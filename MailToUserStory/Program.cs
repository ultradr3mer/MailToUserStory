// GraphToTfsUserStories — .NET 8 CLI using SDKs
// -------------------------------------------------
// Uses Microsoft Graph SDK + Azure.Identity and Azure DevOps/TFS .NET client (REST) for cleaner code.
// Responsibilities:
//   - Poll Microsoft Graph mailboxes via delta queries
//   - Create/update **User Story** work items in on-prem TFS
//   - Store delta tokens + processed message IDs in SQLite
//   - Send replies to the sender with canonical subject token: "[US#12345]"
//   - Convert email bodies into sanitized HTML/Markdown and upload attachments
//
// user-secrets:
//   dotnet user-secrets set "Graph:ClientSecret" "<client-secret>"
//   dotnet user-secrets set "Tfs:Password" "<password>"

using Azure.Identity;
using MailToUserStory;
using Microsoft.Extensions.Configuration;
using Microsoft.Graph;
using Microsoft.TeamFoundation.WorkItemTracking.WebApi;
using Microsoft.VisualStudio.Services.Common;
using Microsoft.VisualStudio.Services.WebApi;

// ----------------------------
// Program startup
// ----------------------------
var config = new ConfigurationBuilder()
    .AddJsonFile("appsettings.json", optional: false)
    .AddJsonFile("appsettings.local.json", optional: true)
    .AddUserSecrets(typeof(AppConfig).Assembly, optional: true)
    .Build();

// Load config into strongly typed model
var app = config.Get<AppConfig>() ?? throw new Exception("Invalid configuration");
app = app with
{
  Graph = app.Graph with { ClientSecret = config["Graph:ClientSecret"] ?? app.Graph.ClientSecret },
  Tfs = app.Tfs with { Password = config["Tfs:Password"] ?? app.Tfs.Password }
};

// Verify secrets exist
Util.Assert(!string.IsNullOrWhiteSpace(app.Graph.ClientSecret), "Graph:ClientSecret missing (user-secrets)");
Util.Assert(!string.IsNullOrWhiteSpace(app.Tfs.Password), "Tfs:Password missing (user-secrets)");

// SQLite DB for state (delta tokens, processed mails, lock)
using var db = new Db(app.DatabasePath);
Db.InitializeSchema(db);

// Create Graph SDK client
var credential = new ClientSecretCredential(app.Graph.TenantId, app.Graph.ClientId, app.Graph.ClientSecret);
var scopes = new[] { "https://graph.microsoft.com/.default" };
var graph = new GraphServiceClient(credential, scopes);

// Create TFS client
var tfsUri = new Uri(app.Tfs.BaseUrl);
var connection = new VssConnection(tfsUri, new VssBasicCredential(app.Tfs.User, app.Tfs.Password));
var wit = await connection.GetClientAsync<WorkItemTrackingHttpClient>();

// Markdown converter (for logs/debugging, not persisted to TFS)
var mdConverter = new ReverseMarkdown.Converter(new ReverseMarkdown.Config
{
  GithubFlavored = true,
  UnknownTags = ReverseMarkdown.Config.UnknownTagsOption.Bypass
});

// ----------------------------
// Main loop: process all mailboxes
// ----------------------------
foreach (var mailbox in app.Graph.Mailboxes)
{
  Console.WriteLine($"=== Processing mailbox: {mailbox} ===");
  string? deltaLink = db.GetDeltaLink(mailbox);

  // Iterate all delta pages (new + changed messages)
  await foreach (var page in GraphConnector.DeltaPagesAsync(graph, mailbox, deltaLink))
  {
    Console.WriteLine($"Received {page.Messages.Count} messages from delta query for {mailbox}");

    int nr = 1;
    foreach (var msg in page.Messages)
    {
      Console.WriteLine($"--> Message {nr++} with Subject: {msg.Subject}");

      bool flowControl = await ProcessMessage(mailbox, msg);

      if (!flowControl)
      {
        Console.WriteLine($"Skipped message");
        continue;
      }
    }

    // Store new delta token for next run
    if (!string.IsNullOrEmpty(page.DeltaLink))
    {
      deltaLink = page.DeltaLink;
      db.UpsertDeltaLink(mailbox, deltaLink);
      Console.WriteLine($"Updated delta link for {mailbox}");
    }
  }
}

Console.WriteLine("All mailboxes processed. Done.");
return;

// ----------------------------
// Process a single email
// ----------------------------
async Task<bool> ProcessMessage(string mailbox, Microsoft.Graph.Models.Message msg)
{
  // Skip if we already processed this message
  if (db.WasProcessed(msg.Id!))
  {
    Console.WriteLine($"Message already processed, skipping.");
    return false;
  }

  // Skip if email was sent by the monitored mailbox itself (avoid loops)
  if (GraphConnector.IsSelf(mailbox, msg))
  {
    Console.WriteLine($"Message is self-sent, skipping.");
    db.MarkProcessed(msg.Id!, mailbox, null, "skipped-self");
    return false;
  }

  // Try to extract User Story ID from subject
  int? usId = Util.ParseUserStoryId(msg.Subject);
  try
  {
    if (usId is int existingId)
    {
      // ----------------------------
      // Update existing user story
      // ----------------------------
      Console.WriteLine($"Updating existing User Story #{existingId} for message {msg.Id}");

      if (!await TfsConnector.WorkItemExistsAsync(wit, existingId))
      {
        Console.WriteLine($"User Story #{existingId} not found in TFS. Sending error reply.");
        await GraphConnector.SendErrorReplyAsync(graph, mailbox, msg, $"User Story #{existingId} was not found in TFS.");
        db.MarkProcessed(msg.Id!, mailbox, null, "us-not-found");
        return false;
      }

      var prepared = await Util.PrepareContentAsync(graph, mailbox, msg, mdConverter);

      await TfsConnector.AddCommentAndAttachmentsAsync(
        wit,
        app.Tfs.Project,
        existingId,
        prepared.markdown,
        prepared.attachments
      );

      await GraphConnector.SendInfoReplyAsync(graph, mailbox, msg,
        $"User Story [US#{existingId}] was updated.", null);

      Console.WriteLine($"Updated User Story #{existingId} with new comment/attachments.");
      db.MarkProcessed(msg.Id!, mailbox, existingId, "updated");
    }
    else
    {
      // ----------------------------
      // Create new user story
      // ----------------------------
      Console.WriteLine($"Creating new User Story from message");

      var prepared = await Util.PrepareContentAsync(graph, mailbox, msg, mdConverter);

      int newId = await TfsConnector.CreateUserStoryAsync(
        wit,
        app.Tfs.Project,
        msg.Subject ?? "(no subject)",
        prepared.markdown
      );

      if (prepared.attachments.Count > 0)
      {
        await TfsConnector.AddAttachmentsAsync(wit, app.Tfs.Project, newId, prepared.attachments);
        Console.WriteLine($"Uploaded {prepared.attachments.Count} attachments to User Story #{newId}");
      }

      db.LinkStory(mailbox, newId);

      await GraphConnector.SendInfoReplyAsync(graph, mailbox, msg,
        $"Created new User Story [US#{newId}] from your email.",
        subjectSuffix: $" [US#{newId}]"
      );

      Console.WriteLine($"Created User Story #{newId} for message");
      db.MarkProcessed(msg.Id!, mailbox, newId, "created");
    }
  }
  catch (Exception ex)
  {
    Console.Error.WriteLine($"Error processing message: {ex}");
    throw; // per requirement: crash and let supervisor retry
  }

  return true;
}
