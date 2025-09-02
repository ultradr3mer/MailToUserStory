// GraphToTfsUserStories — .NET 8 CLI using SDKs
// -------------------------------------------------
// Uses Microsoft Graph SDK + Azure.Identity and Azure DevOps/TFS .NET client (REST) for cleaner code.
// - Graph: delta queries, reply (draft + subject token), attachments
// - TFS (on-prem): create/update User Story, add comments, upload attachments
// - SQLite: delta tokens + processed message ids + lease
//
// Quick start
//   dotnet new console -n GraphToTfsUserStories -f net8.0
//   dotnet add package Microsoft.Graph
//   dotnet add package Azure.Identity
//   dotnet add package Microsoft.Data.Sqlite
//   dotnet add package Microsoft.Extensions.Configuration
//   dotnet add package Microsoft.Extensions.Configuration.Json
//   dotnet add package Microsoft.Extensions.Configuration.UserSecrets
//   dotnet add package Microsoft.Extensions.Configuration.Binder
//   dotnet add package HtmlAgilityPack
//   dotnet add package ReverseMarkdown
//   dotnet add package Microsoft.TeamFoundationServer.Client
//   dotnet add package Microsoft.VisualStudio.Services.Client
//
// appsettings.json (example)
// {
//   "Project": "ContosoProject",
//   "Graph": {
//     "TenantId": "00000000-0000-0000-0000-000000000000",
//     "ClientId": "11111111-1111-1111-1111-111111111111",
//     "Mailboxes": ["support@contoso.local", "sales@contoso.local"]
//   },
//   "Tfs": {
//     "BaseUrl": "http://tfs:8080/tfs/DefaultCollection",
//     "ProjectCollection": "DefaultCollection",
//     "Project": "ContosoProject"
//   },
//   "Polling": { "Minutes": 5 },
//   "DatabasePath": "stories.db"
// }
// user-secrets:
//   dotnet user-secrets set "Graph:ClientSecret" "<client-secret>"
//   dotnet user-secrets set "Tfs:Pat" "<tfs-pat>"

using Azure.Identity;
using HtmlAgilityPack;
using MailToUserStory;
using MailToUserStory.Data;
using Microsoft.Extensions.Configuration;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using Microsoft.Graph.Users.Item.MailFolders.Item.Messages.Delta;
using Microsoft.TeamFoundation.WorkItemTracking.WebApi;
using Microsoft.TeamFoundation.WorkItemTracking.WebApi.Models;
using Microsoft.VisualBasic;
using Microsoft.VisualStudio.Services.Common;
using Microsoft.VisualStudio.Services.WebApi;
using Microsoft.VisualStudio.Services.WebApi.Patch.Json;
using System.Net;
using System.Net.Mail;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;

// ----------------------------
// Configuration models
// ----------------------------





// ----------------------------
// Bootstrapping
// ----------------------------
var config = new ConfigurationBuilder()
    .AddJsonFile("appsettings.json", optional: false)
    .AddUserSecrets(typeof(AppConfig).Assembly, optional: true)
    //.AddEnvironmentVariables()
    .Build();

var app = config.Get<AppConfig>() ?? throw new Exception("Invalid configuration");
app = app with
{
  Graph = app.Graph with { ClientSecret = config["Graph:ClientSecret"] ?? app.Graph.ClientSecret },
  Tfs = app.Tfs with { Pat = config["Tfs:Pat"] ?? app.Tfs.Pat }
};
Util.Assert(!string.IsNullOrWhiteSpace(app.Graph.ClientSecret), "Graph:ClientSecret missing (user-secrets)");
Util.Assert(!string.IsNullOrWhiteSpace(app.Tfs.Pat), "Tfs:Pat missing (user-secrets)");

using var db = new Db(app.DatabasePath);
Db.InitializeSchema(db);
using var lease = await Lease.AcquireAsync(db, TimeSpan.FromMinutes(Math.Max(app.Polling.Minutes, 10)));

// Graph SDK client
var credential = new ClientSecretCredential(app.Graph.TenantId, app.Graph.ClientId, app.Graph.ClientSecret);
var scopes = new[] { "https://graph.microsoft.com/.default" };
var graph = new GraphServiceClient(credential, scopes);

// TFS/ADO client (REST)
var tfsUri = new Uri(app.Tfs.BaseUrl);
var connection = new VssConnection(tfsUri, new VssBasicCredential(string.Empty, app.Tfs.Pat));
var wit = await connection.GetClientAsync<WorkItemTrackingHttpClient>();

var mdConverter = new ReverseMarkdown.Converter(new ReverseMarkdown.Config
{
  GithubFlavored = true,
  UnknownTags = ReverseMarkdown.Config.UnknownTagsOption.Bypass
});

foreach (var mailbox in app.Graph.Mailboxes)
{
  Console.WriteLine("Processing " + mailbox);
  string? deltaLink = db.GetDeltaLink(mailbox);

  await foreach (var page in GraphConnector.DeltaPagesAsync(graph, mailbox, deltaLink))
  {
    foreach (var msg in page.Messages)
    {
      if (db.WasProcessed(msg.Id!)) continue;
      if (GraphConnector.IsSelf(mailbox, msg)) { db.MarkProcessed(msg.Id!, mailbox, null, "skipped-self"); continue; }

      int? usId = Util.ParseUserStoryId(msg.Subject);
      try
      {
        if (usId is int existingId)
        {
          // Update existing story
          if (!await TfsConnector.WorkItemExistsAsync(wit, existingId))
          {
            await GraphConnector.SendErrorReplyAsync(graph, mailbox, msg, "User Story #" + existingId + " was not found in TFS.");
            db.MarkProcessed(msg.Id!, mailbox, null, "us-not-found");
            continue;
          }

          var prepared = await Util.PrepareContentAsync(graph, mailbox, msg, mdConverter);
          await TfsConnector.AddCommentAndAttachmentsAsync(wit, app.Tfs.Project, existingId, prepared.markdown, prepared.attachments);
          await GraphConnector.SendInfoReplyAsync(graph, mailbox, msg, "User Story [US#" + existingId + "] was updated.", null);
          db.MarkProcessed(msg.Id!, mailbox, existingId, "updated");
        }
        else
        {
          // Create new user story
          var prepared = await Util.PrepareContentAsync(graph, mailbox, msg, mdConverter);
          int newId = await TfsConnector.CreateUserStoryAsync(wit, app.Tfs.Project, msg.Subject ?? "(no subject)", prepared.markdown);
          if (prepared.attachments.Count > 0)
            await TfsConnector.AddAttachmentsAsync(wit, app.Tfs.Project, newId, prepared.attachments);

          db.LinkStory(mailbox, newId);
          await GraphConnector.SendInfoReplyAsync(graph, mailbox, msg,
              "Created new User Story [US#" + newId + "] from your email.", subjectSuffix: " [US#" + newId + "]");
          db.MarkProcessed(msg.Id!, mailbox, newId, "created");
        }
      }
      catch (Exception ex)
      {
        Console.Error.WriteLine("Error processing " + msg.Id + ": " + ex);
        throw; // per requirement
      }
    }

    if (!string.IsNullOrEmpty(page.DeltaLink))
    {
      deltaLink = page.DeltaLink;
      db.UpsertDeltaLink(mailbox, deltaLink);
    }
  }
}

Console.WriteLine("Done.");
return;

// ----------------------------
// TFS helpers (WorkItemTrackingHttpClient)
// ----------------------------


// ----------------------------
// Glue: content prep + parsing
// ----------------------------




// ----------------------------
// Single-instance lease
// ----------------------------
sealed class Lease : IDisposable
{
  private readonly Db _db;
  private bool _released;
  private Lease(Db db) { _db = db; }

  public static async Task<Lease> AcquireAsync(Db db, TimeSpan duration)
  {
    string owner = Environment.MachineName + ":" + Environment.ProcessId;
    while (true)
    {
      using var tx = db.Connection.BeginTransaction();
      using var cmdSel = db.Connection.CreateCommand();
      cmdSel.Transaction = tx;
      cmdSel.CommandText = "SELECT owner, expires_at FROM Lease WHERE id=1";
      using var r = cmdSel.ExecuteReader();
      string? curOwner = null; DateTimeOffset? expires = null;
      if (r.Read())
      {
        curOwner = r.IsDBNull(0) ? null : r.GetString(0);
        expires = r.IsDBNull(1) ? null : DateTimeOffset.Parse(r.GetString(1));
      }
      r.Close();

      bool canTake = curOwner == null || expires == null || expires < DateTimeOffset.UtcNow;
      using var cmd = db.Connection.CreateCommand();
      cmd.Transaction = tx;
      if (canTake)
      {
        cmd.CommandText = @"
INSERT INTO Lease(id, owner, expires_at) VALUES(1, @o, @e)
ON CONFLICT(id) DO UPDATE SET owner=excluded.owner, expires_at=excluded.expires_at";
        cmd.Parameters.AddWithValue("@o", owner);
        cmd.Parameters.AddWithValue("@e", DateTimeOffset.UtcNow.Add(duration).ToString("O"));
        cmd.ExecuteNonQuery();
        tx.Commit();
        return new Lease(db);
      }
      tx.Rollback();
      await Task.Delay(1000);
    }
  }

  public void Dispose()
  {
    if (_released) return;
    using var tx = _db.Connection.BeginTransaction();
    using var cmd = _db.Connection.CreateCommand();
    cmd.Transaction = tx;
    cmd.CommandText = "UPDATE Lease SET owner=NULL, expires_at=NULL WHERE id=1";
    cmd.ExecuteNonQuery();
    tx.Commit();
    _released = true;
  }
}

// ----------------------------
// Misc models
// ----------------------------

