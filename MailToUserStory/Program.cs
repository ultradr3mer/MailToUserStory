// GraphToTfsUserStories — .NET 8 CLI using SDKs
// -------------------------------------------------
// Uses Microsoft Graph SDK + Azure.Identity and Azure DevOps/TFS .NET client (REST) for cleaner code.
// - Graph: delta queries, reply (draft + subject token), attachments
// - TFS (on-prem): create/update User Story, add comments, upload attachments
// - SQLite: delta tokens + processed message ids + lease
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
//
// user-secrets:
//   dotnet user-secrets set "Graph:ClientSecret" "<client-secret>"
//   dotnet user-secrets set "Tfs:Pat" "<tfs-pat>"

using Azure.Identity;
using MailToUserStory;
using Microsoft.Extensions.Configuration;
using Microsoft.Graph;
using Microsoft.TeamFoundation.WorkItemTracking.WebApi;
using Microsoft.VisualStudio.Services.Common;
using Microsoft.VisualStudio.Services.WebApi;

// ----------------------------
// Programm
// ----------------------------
var config = new ConfigurationBuilder()
    .AddJsonFile("appsettings.json", optional: false)
    .AddJsonFile("appsettings.local.json", optional: true)
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

