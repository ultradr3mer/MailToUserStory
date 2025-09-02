// GraphToTfsUserStories — single‑file, top‑level .NET 8 CLI
// ---------------------------------------------------------
// What it does
//  - Polls Microsoft Graph mailboxes via delta queries
//  - Creates/updates **User Story** work items in on‑prem TFS
//  - Stores delta tokens + processed message ids in SQLite
//  - Replies to the sender with a canonical token subject: "[US#12345]"
//  - Converts HTML email bodies to Markdown, uploads attachments
//  - Simple single‑instance lease (SQLite) to avoid overlap
//
// How to use
//  1) Create a new console project and replace Program.cs with this file.
//     dotnet new console -n GraphToTfsUserStories -f net8.0
//  2) Add packages:
//     dotnet add package Microsoft.Data.Sqlite
//     dotnet add package Microsoft.Extensions.Configuration
//     dotnet add package Microsoft.Extensions.Configuration.Json
//     dotnet add package Microsoft.Extensions.Configuration.UserSecrets
//     dotnet add package Microsoft.Extensions.Configuration.Binder
//     dotnet add package HtmlAgilityPack
//     dotnet add package ReverseMarkdown
//  3) Add a UserSecretsId to your .csproj (any GUID) to enable dotnet user‑secrets.
//  4) Configure appsettings.json (example below) and user‑secrets:
//       dotnet user-secrets set "Graph:ClientSecret" "<client-secret>"
//       dotnet user-secrets set "Tfs:Pat" "<tfs-pat>"
//  5) Run it (one shot); schedule externally (Task Scheduler, cron, etc.).
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
//     "BaseUrl": "http://tfs:8080/tfs/DefaultCollection/ContosoProject",
//     "ApiVersion": "3.2"
//   },
//   "Polling": { "Minutes": 5 },
//   "DatabasePath": "stories.db"
// }
//
// Notes
//  - This targets **on‑prem TFS** REST APIs. PAT is used via Basic auth (username blank).
//  - Replies are sent using Graph createReply + send to preserve threading.
//  - Error policy: throw; rely on external supervisor to re-run.
//  - Adjust TFS ApiVersion to match your TFS (e.g., 2.0, 3.2). For TFS 2018, 3.2 is typical.

using HtmlAgilityPack;
using Microsoft.Data.Sqlite;
using Microsoft.Extensions.Configuration;
using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;
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
  public string? ClientSecret { get; init; } // pulled from user-secrets at runtime
  public required string[] Mailboxes { get; init; }
}

record TfsConfig
{
  public required string BaseUrl { get; init; } // e.g., http://tfs:8080/tfs/DefaultCollection/Project
  public string ApiVersion { get; init; } = "3.2";
  public string? Pat { get; init; } // pulled from user-secrets
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

// Acquire single-instance lease
using var lease = await Lease.AcquireAsync(db, TimeSpan.FromMinutes(Math.Max(app.Polling.Minutes, 10)));

using var http = new HttpClient();
http.Timeout = TimeSpan.FromSeconds(100);
var graph = new GraphClient(http, app.Graph.TenantId, app.Graph.ClientId, app.Graph.ClientSecret!);
var tfs = new TfsClient(http, app.Tfs.BaseUrl.TrimEnd('/'), app.Tfs.ApiVersion, app.Tfs.Pat!);

var converter = new ReverseMarkdown.Converter(new ReverseMarkdown.Config
{
  GithubFlavored = true,
  UnknownTags = ReverseMarkdown.Config.UnknownTagsOption.Bypass
});

foreach (var mailbox in app.Graph.Mailboxes)
{
  Console.WriteLine($"Processing mailbox: {mailbox}");
  string? deltaLink = db.GetDeltaLink(mailbox);

  await foreach (var page in graph.DeltaPagesAsync(mailbox, deltaLink))
  {
    foreach (var msg in page.Messages)
    {
      if (db.WasProcessed(msg.Id)) continue; // idempotency
      if (msg.From?.EmailAddress?.Address?.Equals(mailbox, StringComparison.OrdinalIgnoreCase) == true)
      {
        // Skip our own mailbox to avoid loops
        db.MarkProcessed(msg.Id, mailbox, null, "skipped-self");
        continue;
      }

      int? usId = ParseUserStoryId(msg.Subject);
      try
      {
        if (usId is int existingId)
        {
          // update existing US
          bool exists = await tfs.WorkItemExistsAsync(existingId);
          if (!exists)
          {
            await graph.SendErrorReplyAsync(mailbox, msg.Id, $"User Story #{existingId} was not found in TFS.");
            db.MarkProcessed(msg.Id, mailbox, null, "us-not-found");
            continue;
          }

          var (markdown, attachments) = await PrepareContentAsync(graph, mailbox, msg, converter);

          await tfs.AddCommentAndAttachmentsAsync(existingId, markdown, attachments);

          await graph.SendInfoReplyAsync(mailbox, msg.Id, $"User Story [US#{existingId}] was updated.", null);

          db.MarkProcessed(msg.Id, mailbox, existingId, "updated");
        }
        else
        {
          // create new US
          var (markdown, attachments) = await PrepareContentAsync(graph, mailbox, msg, converter);

          int newId = await tfs.CreateUserStoryAsync(app.Project, msg.Subject ?? "(no subject)", markdown);

          if (attachments.Count > 0)
          {
            await tfs.AddAttachmentsAsync(newId, attachments);
          }

          db.LinkStory(mailbox, newId);

          await graph.SendInfoReplyAsync(mailbox, msg.Id,
              $"Created new User Story [US#{newId}] from your email.",
              subjectSuffix: $" [US#{newId}]");

          db.MarkProcessed(msg.Id, mailbox, newId, "created");
        }
      }
      catch (Exception ex)
      {
        Console.Error.WriteLine($"Error processing message {msg.Id}: {ex.Message}");
        throw; // per requirements: just throw for now
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
// Helpers & domain logic
// ----------------------------
static async Task<(string markdown, List<AttachmentPayload> attachments)> PrepareContentAsync(GraphClient graph, string mailbox, GraphMessage msg, ReverseMarkdown.Converter converter)
{
  string markdown = HtmlToMarkdown(msg.Body?.Content, converter);

  var attachments = new List<AttachmentPayload>();
  if (msg.HasAttachments == true)
  {
    var atts = await graph.GetAttachmentsAsync(mailbox, msg.Id);
    foreach (var a in atts)
    {
      if (a.ODataType?.EndsWith("fileAttachment", StringComparison.OrdinalIgnoreCase) == true &&
          a.ContentBytes is not null && a.Name is not null)
      {
        attachments.Add(new AttachmentPayload
        {
          FileName = a.Name,
          Bytes = Convert.FromBase64String(a.ContentBytes)
        });
      }
    }
  }

  // Append metadata
  var meta = new StringBuilder();
  meta.AppendLine();
  meta.AppendLine("---");
  meta.AppendLine($"> From: {msg.From?.EmailAddress?.Name} <{msg.From?.EmailAddress?.Address}>");
  meta.AppendLine($"> Received: {msg.ReceivedDateTime:O}");

  return (markdown + "\n\n" + meta.ToString(), attachments);
}

static string HtmlToMarkdown(string? html, ReverseMarkdown.Converter converter)
{
  if (string.IsNullOrWhiteSpace(html)) return "(no content)";
  var doc = new HtmlDocument();
  doc.LoadHtml(html);
  // Remove script/style
  foreach (var n in doc.DocumentNode.SelectNodes("//script|//style") ?? Array.Empty<HtmlNode>())
    n.Remove();
  string sanitized = doc.DocumentNode.InnerHtml;
  string md = converter.Convert(sanitized);
  return md.Trim();
}

static int? ParseUserStoryId(string? subject)
{
  if (string.IsNullOrEmpty(subject)) return null;
  // Accept variants like [US#123], [UserStory#123], US:123, UserStory: 123, US-123, etc.
  var rx = new Regex(@"(?ix)
        (?:\[(?:US|User\s*Story)\s*#\s*(?<id>\d{1,10})\])
        |
        (?:\b(?:US|User\s*Story)\s*[:#\-]\s*(?<id>\d{1,10})\b)
    ");
  var m = rx.Match(subject);
  if (m.Success && int.TryParse(m.Groups["id"].Value, out var id))
    return id;
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
    _conn = new SqliteConnection($"Data Source={path}");
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
    string owner = $"{Environment.MachineName}:{Environment.ProcessId}";
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
// Graph client (minimal REST)
// ----------------------------
sealed class GraphClient
{
  private readonly HttpClient _http;
  private readonly string _tenantId;
  private readonly string _clientId;
  private readonly string _clientSecret;
  private string? _accessToken;
  private DateTimeOffset _tokenExpires;

  public GraphClient(HttpClient http, string tenantId, string clientId, string clientSecret)
  {
    _http = http; _tenantId = tenantId; _clientId = clientId; _clientSecret = clientSecret;
  }

  private async Task<string> TokenAsync()
  {
    if (_accessToken != null && _tokenExpires > DateTimeOffset.UtcNow.AddMinutes(1))
      return _accessToken;
    var content = new FormUrlEncodedContent(new Dictionary<string, string>
    {
      ["client_id"] = _clientId,
      ["client_secret"] = _clientSecret,
      ["grant_type"] = "client_credentials",
      ["scope"] = "https://graph.microsoft.com/.default"
    });
    using var resp = await _http.PostAsync($"https://login.microsoftonline.com/{_tenantId}/oauth2/v2.0/token", content);
    resp.EnsureSuccessStatusCode();
    using var s = await resp.Content.ReadAsStreamAsync();
    var doc = await JsonDocument.ParseAsync(s);
    _accessToken = doc.RootElement.GetProperty("access_token").GetString();
    int expiresIn = doc.RootElement.GetProperty("expires_in").GetInt32();
    _tokenExpires = DateTimeOffset.UtcNow.AddSeconds(expiresIn);
    return _accessToken!;
  }

  private async Task<HttpRequestMessage> CreateRequestAsync(HttpMethod method, string url)
  {
    var req = new HttpRequestMessage(method, url);
    req.Headers.Authorization = new AuthenticationHeaderValue("Bearer", await TokenAsync());
    return req;
  }

  public async IAsyncEnumerable<DeltaPage> DeltaPagesAsync(string mailbox, string? deltaLink)
  {
    string url = deltaLink ?? $"https://graph.microsoft.com/v1.0/users/{Uri.EscapeDataString(mailbox)}/mailFolders('Inbox')/messages/delta?$select=id,subject,from,receivedDateTime,hasAttachments,body";

    while (true)
    {
      using var req = await CreateRequestAsync(HttpMethod.Get, url);
      using var resp = await _http.SendAsync(req);
      resp.EnsureSuccessStatusCode();
      await using var s = await resp.Content.ReadAsStreamAsync();
      var doc = await JsonDocument.ParseAsync(s);

      var msgs = new List<GraphMessage>();
      if (doc.RootElement.TryGetProperty("value", out var arr))
      {
        foreach (var el in arr.EnumerateArray())
        {
          msgs.Add(GraphMessage.From(el));
        }
      }

      string? nextLink = doc.RootElement.TryGetProperty("@odata.nextLink", out var nl) ? nl.GetString() : null;
      string? delta = doc.RootElement.TryGetProperty("@odata.deltaLink", out var dl) ? dl.GetString() : null;

      yield return new DeltaPage { Messages = msgs, NextLink = nextLink, DeltaLink = delta };

      if (nextLink != null) url = nextLink;
      else break;
    }
  }

  public async Task<List<GraphAttachment>> GetAttachmentsAsync(string mailbox, string messageId)
  {
    using var req = await CreateRequestAsync(HttpMethod.Get,
        $"https://graph.microsoft.com/v1.0/users/{Uri.EscapeDataString(mailbox)}/messages/{messageId}/attachments?$select=id,name,contentBytes,@odata.type");
    using var resp = await _http.SendAsync(req);
    resp.EnsureSuccessStatusCode();
    await using var s = await resp.Content.ReadAsStreamAsync();
    var doc = await JsonDocument.ParseAsync(s);
    var list = new List<GraphAttachment>();
    foreach (var el in doc.RootElement.GetProperty("value").EnumerateArray())
      list.Add(GraphAttachment.From(el));
    return list;
  }

  public async Task SendInfoReplyAsync(string mailbox, string messageId, string infoBody, string? subjectSuffix)
  {
    // create draft reply
    using (var req = await CreateRequestAsync(HttpMethod.Post,
        $"https://graph.microsoft.com/v1.0/users/{Uri.EscapeDataString(mailbox)}/messages/{messageId}/createReply"))
    {
      req.Content = new StringContent("{}", Encoding.UTF8, "application/json");
      using var resp = await _http.SendAsync(req);
      resp.EnsureSuccessStatusCode();
      await using var s = await resp.Content.ReadAsStreamAsync();
      var doc = await JsonDocument.ParseAsync(s);
      var replyId = doc.RootElement.GetProperty("id").GetString()!;

      // set subject/body
      var patch = new
      {
        body = new { contentType = "Text", content = infoBody },
      };
      var patchJson = JsonSerializer.Serialize(patch);

      using var req2 = await CreateRequestAsync(HttpMethod.Patch,
          $"https://graph.microsoft.com/v1.0/users/{Uri.EscapeDataString(mailbox)}/messages/{replyId}");
      req2.Content = new StringContent(patchJson, Encoding.UTF8, "application/json");
      using var resp2 = await _http.SendAsync(req2);
      resp2.EnsureSuccessStatusCode();

      if (!string.IsNullOrEmpty(subjectSuffix))
      {
        // Get original subject to append (Graph doesn't echo it back reliably here). We'll fetch message to read subject.
        string subject = await GetMessageSubjectAsync(mailbox, messageId);
        var patchSub = new { subject = subject + subjectSuffix };
        using var reqSub = await CreateRequestAsync(HttpMethod.Patch,
            $"https://graph.microsoft.com/v1.0/users/{Uri.EscapeDataString(mailbox)}/messages/{replyId}");
        reqSub.Content = new StringContent(JsonSerializer.Serialize(patchSub), Encoding.UTF8, "application/json");
        using var respSub = await _http.SendAsync(reqSub);
        respSub.EnsureSuccessStatusCode();
      }

      // send
      using var req3 = await CreateRequestAsync(HttpMethod.Post,
          $"https://graph.microsoft.com/v1.0/users/{Uri.EscapeDataString(mailbox)}/messages/{replyId}/send");
      using var resp3 = await _http.SendAsync(req3);
      resp3.EnsureSuccessStatusCode();
    }
  }

  public async Task SendErrorReplyAsync(string mailbox, string messageId, string errorText)
      => await SendInfoReplyAsync(mailbox, messageId, errorText, null);

  private async Task<string> GetMessageSubjectAsync(string mailbox, string messageId)
  {
    using var req = await CreateRequestAsync(HttpMethod.Get,
        $"https://graph.microsoft.com/v1.0/users/{Uri.EscapeDataString(mailbox)}/messages/{messageId}?$select=subject");
    using var resp = await _http.SendAsync(req);
    resp.EnsureSuccessStatusCode();
    await using var s = await resp.Content.ReadAsStreamAsync();
    var doc = await JsonDocument.ParseAsync(s);
    return doc.RootElement.GetProperty("subject").GetString() ?? string.Empty;
  }
}

sealed class DeltaPage
{
  public required List<GraphMessage> Messages { get; init; }
  public string? NextLink { get; init; }
  public string? DeltaLink { get; init; }
}

sealed class GraphMessage
{
  public required string Id { get; init; }
  public string? Subject { get; init; }
  public GraphRecipient? From { get; init; }
  public DateTimeOffset? ReceivedDateTime { get; init; }
  public bool? HasAttachments { get; init; }
  public GraphItemBody? Body { get; init; }

  public static GraphMessage From(JsonElement e)
  {
    return new GraphMessage
    {
      Id = e.GetProperty("id").GetString()!,
      Subject = e.TryGetProperty("subject", out var sub) ? sub.GetString() : null,
      From = e.TryGetProperty("from", out var fr) ? GraphRecipient.From(fr) : null,
      ReceivedDateTime = e.TryGetProperty("receivedDateTime", out var dt) ? dt.GetDateTimeOffset() : (DateTimeOffset?)null,
      HasAttachments = e.TryGetProperty("hasAttachments", out var ha) && ha.GetBoolean(),
      Body = e.TryGetProperty("body", out var b) ? GraphItemBody.From(b) : null
    };
  }
}

sealed class GraphItemBody
{
  public string? ContentType { get; init; }
  public string? Content { get; init; }
  public static GraphItemBody From(JsonElement e) => new()
  {
    ContentType = e.TryGetProperty("contentType", out var t) ? t.GetString() : null,
    Content = e.TryGetProperty("content", out var c) ? c.GetString() : null
  };
}

sealed class GraphRecipient
{
  public GraphEmailAddress? EmailAddress { get; init; }
  public static GraphRecipient From(JsonElement e) => new()
  {
    EmailAddress = e.TryGetProperty("emailAddress", out var ea) ? GraphEmailAddress.From(ea) : null
  };
}

sealed class GraphEmailAddress
{
  public string? Name { get; init; }
  public string? Address { get; init; }
  public static GraphEmailAddress From(JsonElement e) => new()
  {
    Name = e.TryGetProperty("name", out var n) ? n.GetString() : null,
    Address = e.TryGetProperty("address", out var a) ? a.GetString() : null
  };
}

sealed class GraphAttachment
{
  public string? Id { get; init; }
  public string? Name { get; init; }
  public string? ContentBytes { get; init; }
  public string? ODataType { get; init; }

  public static GraphAttachment From(JsonElement e) => new()
  {
    Id = e.TryGetProperty("id", out var id) ? id.GetString() : null,
    Name = e.TryGetProperty("name", out var n) ? n.GetString() : null,
    ContentBytes = e.TryGetProperty("contentBytes", out var cb) ? cb.GetString() : null,
    ODataType = e.TryGetProperty("@odata.type", out var od) ? od.GetString() : null
  };
}

// ----------------------------
// TFS client (on‑prem REST)
// ----------------------------
sealed class TfsClient
{
  private readonly HttpClient _http;
  private readonly string _base;
  private readonly string _api;
  private readonly string _pat;
  private readonly string _authHeader;

  public TfsClient(HttpClient http, string baseUrl, string apiVersion, string pat)
  {
    _http = http; _base = baseUrl; _api = apiVersion; _pat = pat;
    _authHeader = Convert.ToBase64String(Encoding.ASCII.GetBytes($":{_pat}"));
  }

  private HttpRequestMessage Req(HttpMethod method, string path)
  {
    var req = new HttpRequestMessage(method, path);
    req.Headers.Authorization = new AuthenticationHeaderValue("Basic", _authHeader);
    return req;
  }

  public async Task<bool> WorkItemExistsAsync(int id)
  {
    using var req = Req(HttpMethod.Get, $"{_base}/_apis/wit/workitems/{id}?api-version={_api}");
    using var resp = await _http.SendAsync(req);
    if (resp.StatusCode == System.Net.HttpStatusCode.NotFound) return false;
    resp.EnsureSuccessStatusCode();
    return true;
  }

  public async Task<int> CreateUserStoryAsync(string project, string title, string descriptionMarkdown)
  {
    var ops = new List<object>
        {
            new { op = "add", path = "/fields/System.Title", value = title },
            new { op = "add", path = "/fields/System.Description", value = MarkdownAsHtml(descriptionMarkdown) }
        };
    var json = JsonSerializer.Serialize(ops);

    using var req = Req(new HttpMethod("PATCH"), $"{_base}/{Uri.EscapeDataString(project)}/_apis/wit/workitems/$User%20Story?api-version={_api}");
    req.Content = new StringContent(json, Encoding.UTF8, "application/json-patch+json");
    using var resp = await _http.SendAsync(req);
    resp.EnsureSuccessStatusCode();
    await using var s = await resp.Content.ReadAsStreamAsync();
    var doc = await JsonDocument.ParseAsync(s);
    return doc.RootElement.GetProperty("id").GetInt32();
  }

  public async Task AddCommentAndAttachmentsAsync(int workItemId, string commentMarkdown, List<AttachmentPayload> attachments)
  {
    var ops = new List<object>
        {
            new { op = "add", path = "/fields/System.History", value = MarkdownAsHtml(commentMarkdown) }
        };

    if (attachments.Count > 0)
    {
      foreach (var a in attachments)
      {
        var url = await UploadAttachmentAsync(a.FileName, a.Bytes);
        ops.Add(new
        {
          op = "add",
          path = "/relations/-",
          value = new { rel = "AttachedFile", url }
        });
      }
    }

    var json = JsonSerializer.Serialize(ops);
    using var req = Req(new HttpMethod("PATCH"), $"{_base}/_apis/wit/workitems/{workItemId}?api-version={_api}");
    req.Content = new StringContent(json, Encoding.UTF8, "application/json-patch+json");
    using var resp = await _http.SendAsync(req);
    resp.EnsureSuccessStatusCode();
  }

  public async Task AddAttachmentsAsync(int workItemId, List<AttachmentPayload> attachments)
  {
    if (attachments.Count == 0) return;
    var ops = new List<object>();
    foreach (var a in attachments)
    {
      var url = await UploadAttachmentAsync(a.FileName, a.Bytes);
      ops.Add(new { op = "add", path = "/relations/-", value = new { rel = "AttachedFile", url } });
    }
    var json = JsonSerializer.Serialize(ops);
    using var req = Req(new HttpMethod("PATCH"), $"{_base}/_apis/wit/workitems/{workItemId}?api-version={_api}");
    req.Content = new StringContent(json, Encoding.UTF8, "application/json-patch+json");
    using var resp = await _http.SendAsync(req);
    resp.EnsureSuccessStatusCode();
  }

  private async Task<string> UploadAttachmentAsync(string fileName, byte[] bytes)
  {
    using var req = Req(HttpMethod.Post, $"{_base}/_apis/wit/attachments?fileName={Uri.EscapeDataString(fileName)}&api-version={_api}");
    req.Content = new ByteArrayContent(bytes);
    req.Content.Headers.ContentType = new MediaTypeHeaderValue("application/octet-stream");
    using var resp = await _http.SendAsync(req);
    resp.EnsureSuccessStatusCode();
    await using var s = await resp.Content.ReadAsStreamAsync();
    var doc = await JsonDocument.ParseAsync(s);
    return doc.RootElement.GetProperty("url").GetString()!;
  }

  private static string MarkdownAsHtml(string md)
  {
    // TFS accepts HTML in System.Description/System.History. Minimal wrapping.
    string escaped = System.Net.WebUtility.HtmlEncode(md).Replace("\n", "<br/>");
    return $"<div>{escaped}</div>";
  }
}

// ----------------------------
// Misc models
// ----------------------------
sealed class AttachmentPayload
{
  public required string FileName { get; init; }
  public required byte[] Bytes { get; init; }
}