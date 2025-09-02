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
using Microsoft.Data.Sqlite;
using Microsoft.Extensions.Configuration;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using Microsoft.Graph.Users.Item.MailFolders.Item.Messages.Delta;
using Microsoft.TeamFoundation.WorkItemTracking.WebApi;
using Microsoft.TeamFoundation.WorkItemTracking.WebApi.Models;
using Microsoft.VisualStudio.Services.Common;
using Microsoft.VisualStudio.Services.WebApi;
using System.Net;
using System.Net.Mail;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;

// ----------------------------
// Configuration models
// ----------------------------
record AppConfig
{
  public required string Project { get; init; }
  public required GraphConfig Graph { get; init; }
  public required TfsConfig Tfs { get; init; }
  public PollingConfig Polling { get; init; } = new();
  public string DatabasePath { get; init; } = "stories.db";
}

record GraphConfig
{
  public required string TenantId { get; init; }
  public required string ClientId { get; init; }
  public string? ClientSecret { get; init; }
  public required string[] Mailboxes { get; init; }
}

record TfsConfig
{
  public required string BaseUrl { get; init; }
  public required string ProjectCollection { get; init; }
  public required string Project { get; init; }
  public string? Pat { get; init; }
}

record PollingConfig { public int Minutes { get; init; } = 5; }

// ----------------------------
// Bootstrapping
// ----------------------------
var config = new ConfigurationBuilder()
    .AddJsonFile("appsettings.json", optional: false)
    .AddUserSecrets(typeof(AppConfig).Assembly, optional: true)
    .AddEnvironmentVariables()
    .Build();

var app = config.Get<AppConfig>() ?? throw new Exception("Invalid configuration");
app = app with
{
  Graph = app.Graph with { ClientSecret = config["Graph:ClientSecret"] ?? app.Graph.ClientSecret },
  Tfs = app.Tfs with { Pat = config["Tfs:Pat"] ?? app.Tfs.Pat }
};
Assert(!string.IsNullOrWhiteSpace(app.Graph.ClientSecret), "Graph:ClientSecret missing (user-secrets)");
Assert(!string.IsNullOrWhiteSpace(app.Tfs.Pat), "Tfs:Pat missing (user-secrets)");

using var db = new Db(app.DatabasePath);
InitializeSchema(db);
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

  await foreach (var page in DeltaPagesAsync(graph, mailbox, deltaLink))
  {
    foreach (var msg in page.Messages)
    {
      if (db.WasProcessed(msg.Id!)) continue;
      if (IsSelf(mailbox, msg)) { db.MarkProcessed(msg.Id!, mailbox, null, "skipped-self"); continue; }

      int? usId = ParseUserStoryId(msg.Subject);
      try
      {
        if (usId is int existingId)
        {
          // Update existing story
          if (!await WorkItemExistsAsync(wit, existingId))
          {
            await SendErrorReplyAsync(graph, mailbox, msg, "User Story #" + existingId + " was not found in TFS.");
            db.MarkProcessed(msg.Id!, mailbox, null, "us-not-found");
            continue;
          }

          var prepared = await PrepareContentAsync(graph, mailbox, msg, mdConverter);
          await AddCommentAndAttachmentsAsync(wit, app.Tfs.Project, existingId, prepared.markdown, prepared.attachments);
          await SendInfoReplyAsync(graph, mailbox, msg, "User Story [US#" + existingId + "] was updated.", null);
          db.MarkProcessed(msg.Id!, mailbox, existingId, "updated");
        }
        else
        {
          // Create new user story
          var prepared = await PrepareContentAsync(graph, mailbox, msg, mdConverter);
          int newId = await CreateUserStoryAsync(wit, app.Tfs.Project, msg.Subject ?? "(no subject)", prepared.markdown);
          if (prepared.attachments.Count > 0)
            await AddAttachmentsAsync(wit, app.Tfs.Project, newId, prepared.attachments);

          db.LinkStory(mailbox, newId);
          await SendInfoReplyAsync(graph, mailbox, msg,
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
// Graph helpers (SDK)
// ----------------------------
static async IAsyncEnumerable<DeltaPage> DeltaPagesAsync(GraphServiceClient graph, string mailbox, string? deltaLink)
{
  DeltaResponse? page;
  if (!string.IsNullOrEmpty(deltaLink))
  {
    // Resume from stored delta link
    page = await graph.RequestAdapter.SendAsync<DeltaResponse>(new Microsoft.Kiota.Abstractions.RequestInformation
    {
      HttpMethod = Microsoft.Kiota.Abstractions.Method.GET,
      UrlTemplate = deltaLink,
    });
  }
  else
  {
    page = await graph.Users[mailbox]
        .MailFolders["Inbox"].Messages.Delta
        .GetAsync(r =>
        {
          r.QueryParameters.Select = new[] { "id", "subject", "from", "receivedDateTime", "hasAttachments", "body" };
        });
  }

  while (page != null)
  {
    yield return new DeltaPage
    {
      Messages = page.Value?.ToList() ?? new List<Message>(),
      NextLink = page.OdataNextLink,
      DeltaLink = page.OdataDeltaLink
    };

    if (!string.IsNullOrEmpty(page.OdataNextLink))
    {
      page = await graph.RequestAdapter.SendAsync<DeltaResponse>(new Microsoft.Kiota.Abstractions.RequestInformation
      {
        HttpMethod = Microsoft.Kiota.Abstractions.Method.GET,
        UrlTemplate = page.OdataNextLink,
      });
    }
    else
    {
      break;
    }
  }
}

static bool IsSelf(string mailbox, Message msg)
    => string.Equals(msg.From?.EmailAddress?.Address, mailbox, StringComparison.OrdinalIgnoreCase);

static async Task<List<FileAttachment>> GetFileAttachmentsAsync(GraphServiceClient graph, string mailbox, string messageId)
{
  var result = new List<FileAttachment>();
  var page = await graph.Users[mailbox].Messages[messageId].Attachments.GetAsync();
  foreach (var att in page?.Value ?? Enumerable.Empty<Attachment>())
  {
    if (att is FileAttachment fa && fa.ContentBytes != null)
      result.Add(fa);
  }
  return result;
}

static async Task SendInfoReplyAsync(GraphServiceClient graph, string mailbox, Message original, string infoBody, string? subjectSuffix)
{
  // Create a draft reply to preserve threading, then patch subject/body, then send
  var draft = await graph.Users[mailbox].Messages[original.Id!].CreateReply.PostAsync(new Microsoft.Graph.Users.Item.Messages.Item.CreateReply.CreateReplyPostRequestBody
  {
    Message = new Message
    {
      Body = new ItemBody { ContentType = BodyType.Text, Content = infoBody }
    }
  });

  if (draft == null) throw new Exception("Failed to create reply draft");

  string subject = original.Subject ?? string.Empty;
  if (!string.IsNullOrEmpty(subjectSuffix)) subject = subject + subjectSuffix;

  await graph.Users[mailbox].Messages[draft.Id!].PatchAsync(new Message
  {
    Subject = subject,
    Body = new ItemBody { ContentType = BodyType.Text, Content = infoBody }
  });

  await graph.Users[mailbox].Messages[draft.Id!].Send.PostAsync();
}

static Task SendErrorReplyAsync(GraphServiceClient graph, string mailbox, Message original, string errorText)
    => SendInfoReplyAsync(graph, mailbox, original, errorText, null);

// ----------------------------
// TFS helpers (WorkItemTrackingHttpClient)
// ----------------------------
static async Task<bool> WorkItemExistsAsync(WorkItemTrackingHttpClient wit, int id)
{
  try { _ = await wit.GetWorkItemAsync(id); return true; }
  catch (Microsoft.VisualStudio.Services.WebApi.VssServiceException ex) when (ex.Message.Contains("404")) { return false; }
  catch (Microsoft.VisualStudio.Services.WebApi.VssServiceResponseException ex) when (ex.HttpStatusCode == HttpStatusCode.NotFound) { return false; }
}

static async Task<int> CreateUserStoryAsync(WorkItemTrackingHttpClient wit, string project, string title, string descriptionMarkdown)
{
  var patch = new JsonPatchDocument
    {
        new JsonPatchOperation { Operation = Operation.Add, Path = "/fields/System.Title", Value = title },
        new JsonPatchOperation { Operation = Operation.Add, Path = "/fields/System.Description", Value = MarkdownAsHtml(descriptionMarkdown) }
    };
  var wi = await wit.CreateWorkItemAsync(patch, project, "User Story");
  return wi.Id ?? throw new Exception("No ID returned from CreateWorkItemAsync");
}

static async Task AddCommentAndAttachmentsAsync(WorkItemTrackingHttpClient wit, string project, int id, string commentMarkdown, List<AttachmentPayload> attachments)
{
  var patch = new JsonPatchDocument
    {
        new JsonPatchOperation { Operation = Operation.Add, Path = "/fields/System.History", Value = MarkdownAsHtml(commentMarkdown) }
    };

  foreach (var a in attachments)
  {
    using var ms = new MemoryStream(a.Bytes);
    var ar = await wit.CreateAttachmentAsync(ms, fileName: a.FileName);
    patch.Add(new JsonPatchOperation
    {
      Operation = Operation.Add,
      Path = "/relations/-",
      Value = new WorkItemRelation { Rel = "AttachedFile", Url = ar.Url }
    });
  }

  _ = await wit.UpdateWorkItemAsync(patch, id);
}

static async Task AddAttachmentsAsync(WorkItemTrackingHttpClient wit, string project, int id, List<AttachmentPayload> attachments)
{
  if (attachments.Count == 0) return;
  var patch = new JsonPatchDocument();
  foreach (var a in attachments)
  {
    using var ms = new MemoryStream(a.Bytes);
    var ar = await wit.CreateAttachmentAsync(ms, fileName: a.FileName);
    patch.Add(new JsonPatchOperation
    {
      Operation = Operation.Add,
      Path = "/relations/-",
      Value = new WorkItemRelation { Rel = "AttachedFile", Url = ar.Url }
    });
  }
  _ = await wit.UpdateWorkItemAsync(patch, id);
}

static string MarkdownAsHtml(string md)
{
  string escaped = System.Net.WebUtility.HtmlEncode(md).Replace("\n", "<br/>");
  return "<div>" + escaped + "</div>";
}

// ----------------------------
// Glue: content prep + parsing
// ----------------------------
static async Task<(string markdown, List<AttachmentPayload> attachments)> PrepareContentAsync(GraphServiceClient graph, string mailbox, Message msg, ReverseMarkdown.Converter converter)
{
  string markdown = HtmlToMarkdown(msg.Body, converter);
  var attachments = new List<AttachmentPayload>();
  if (msg.HasAttachments == true)
  {
    var files = await GetFileAttachmentsAsync(graph, mailbox, msg.Id!);
    foreach (var fa in files)
    {
      attachments.Add(new AttachmentPayload
      {
        FileName = fa.Name!,
        Bytes = fa.ContentBytes!
      });
    }
  }

  var meta = new StringBuilder();
  meta.AppendLine();
  meta.AppendLine("---");
  meta.AppendLine("> From: " + msg.From?.EmailAddress?.Name + " <" + msg.From?.EmailAddress?.Address + ">");
  meta.AppendLine("> Received: " + (msg.ReceivedDateTime.HasValue ? msg.ReceivedDateTime.Value.ToString("O") : ""));

  return (markdown + "\n\n" + meta.ToString(), attachments);
}

static string HtmlToMarkdown(ItemBody? body, ReverseMarkdown.Converter converter)
{
  if (body == null) return "(no content)";
  if (body.ContentType == BodyType.Text) return string.IsNullOrWhiteSpace(body.Content) ? "(no content)" : body.Content!.Trim();

  var html = body.Content ?? string.Empty;
  var doc = new HtmlDocument();
  doc.LoadHtml(html);
  foreach (var n in doc.DocumentNode.SelectNodes("//script|//style") ?? Array.Empty<HtmlNode>()) n.Remove();
  string sanitized = doc.DocumentNode.InnerHtml;
  string md = converter.Convert(sanitized);
  return md.Trim();
}

static int? ParseUserStoryId(string? subject)
{
  if (string.IsNullOrEmpty(subject)) return null;
  var rx = new Regex(@"(?ix)
        (?:\[(?:US|User\s*Story)\s*#\s*(?<id>\d{1,10})\])
        |
        (?:\b(?:US|User\s*Story)\s*[:#\-]\s*(?<id>\d{1,10})\b)
    ");
  var m = rx.Match(subject);
  if (m.Success && int.TryParse(m.Groups["id"].Value, out var id)) return id;
  return null;
}

static void Assert(bool condition, string message)
{
  if (!condition) throw new Exception(message);
}

// ----------------------------
// DB layer (SQLite)
// ----------------------------
sealed class Db : IDisposable
{
  private readonly SqliteConnection _conn;
  public Db(string path)
  {
    _conn = new SqliteConnection("Data Source=" + path);
    _conn.Open();
  }
  public SqliteConnection Connection => _conn;
  public void Dispose() => _conn.Dispose();

  public string? GetDeltaLink(string mailbox)
  {
    using var cmd = _conn.CreateCommand();
    cmd.CommandText = "SELECT delta_link FROM Mailboxes WHERE address=@a LIMIT 1";
    cmd.Parameters.AddWithValue("@a", mailbox);
    return cmd.ExecuteScalar() as string;
  }

  public void UpsertDeltaLink(string mailbox, string? delta)
  {
    using var tx = _conn.BeginTransaction();
    using var cmd = _conn.CreateCommand();
    cmd.Transaction = tx;
    cmd.CommandText = @"
INSERT INTO Mailboxes(address, delta_link) VALUES(@a, @d)
ON CONFLICT(address) DO UPDATE SET delta_link=excluded.delta_link";
    cmd.Parameters.AddWithValue("@a", mailbox);
    cmd.Parameters.AddWithValue("@d", (object?)delta ?? DBNull.Value);
    cmd.ExecuteNonQuery();
    tx.Commit();
  }

  public bool WasProcessed(string messageId)
  {
    using var cmd = _conn.CreateCommand();
    cmd.CommandText = "SELECT 1 FROM ProcessedEmails WHERE graph_message_id=@id LIMIT 1";
    cmd.Parameters.AddWithValue("@id", messageId);
    using var r = cmd.ExecuteReader();
    return r.Read();
  }

  public void MarkProcessed(string messageId, string mailbox, int? workItemId, string outcome)
  {
    using var tx = _conn.BeginTransaction();
    using var cmd = _conn.CreateCommand();
    cmd.Transaction = tx;
    cmd.CommandText = @"
INSERT INTO ProcessedEmails(graph_message_id, mailbox, work_item_id, processed_at, outcome)
VALUES(@id, @mb, @wi, @ts, @out)";
    cmd.Parameters.AddWithValue("@id", messageId);
    cmd.Parameters.AddWithValue("@mb", mailbox);
    cmd.Parameters.AddWithValue("@wi", (object?)workItemId ?? DBNull.Value);
    cmd.Parameters.AddWithValue("@ts", DateTimeOffset.UtcNow.ToString("O"));
    cmd.Parameters.AddWithValue("@out", outcome);
    cmd.ExecuteNonQuery();
    tx.Commit();
  }

  public void LinkStory(string mailbox, int workItemId)
  {
    using var tx = _conn.BeginTransaction();
    using var cmd = _conn.CreateCommand();
    cmd.Transaction = tx;
    cmd.CommandText = @"
INSERT OR IGNORE INTO Stories(work_item_id, mailbox) VALUES(@wi, @mb)";
    cmd.Parameters.AddWithValue("@wi", workItemId);
    cmd.Parameters.AddWithValue("@mb", mailbox);
    cmd.ExecuteNonQuery();
    tx.Commit();
  }
}

static void InitializeSchema(Db db)
{
  using var cmd = db.Connection.CreateCommand();
  cmd.CommandText = @"
CREATE TABLE IF NOT EXISTS Mailboxes(
  address TEXT PRIMARY KEY,
  delta_link TEXT
);
CREATE TABLE IF NOT EXISTS Stories(
  work_item_id INTEGER,
  mailbox TEXT,
  PRIMARY KEY(work_item_id, mailbox)
);
CREATE TABLE IF NOT EXISTS ProcessedEmails(
  graph_message_id TEXT PRIMARY KEY,
  mailbox TEXT NOT NULL,
  work_item_id INTEGER NULL,
  processed_at TEXT NOT NULL,
  outcome TEXT NOT NULL
);
CREATE TABLE IF NOT EXISTS Lease(
  id INTEGER PRIMARY KEY CHECK(id=1),
  owner TEXT,
  expires_at TEXT
);
";
  cmd.ExecuteNonQuery();
}

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
sealed class DeltaPage
{
  public required List<Message> Messages { get; init; }
  public string? NextLink { get; init; }
  public string? DeltaLink { get; init; }
}

sealed class AttachmentPayload
{
  public required string FileName { get; init; }
  public required byte[] Bytes { get; init; }
}
