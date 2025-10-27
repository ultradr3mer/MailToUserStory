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
using System.Diagnostics;

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
  Graph = app.Graph with { ClientSecret = config["Graph:ClientSecret"] 
    ?? Environment.GetEnvironmentVariable("MailToUserStoryGraphClientSecret")
    ?? app.Graph.ClientSecret },
  Tfs = app.Tfs with { Password = config["Tfs:Password"]
    ?? Environment.GetEnvironmentVariable("MailToUserStoryTfsPassword")
    ?? app.Tfs.Password }
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
var graphClient = new GraphServiceClient(credential, scopes);
var graph = new GraphConnector(graphClient);

// Create TFS client
var tfsUri = new Uri(app.Tfs.BaseUrl);
var connection = new VssConnection(tfsUri, new VssBasicCredential(app.Tfs.User, app.Tfs.Password));
var wit = await connection.GetClientAsync<WorkItemTrackingHttpClient>();

// Create Ollama client
var ollama = new OllamaClient(app.Ollama.Host, app.Ollama.Model, app.Ollama.SummarizeInstruction);

// Initilize Summary Generator
var summaryGenerator = new SummaryGenerator(ollama, app.Ollama.Enabled);

// Markdown converter (for logs/debugging, not persisted to TFS)
var mdConverter = new ReverseMarkdown.Converter(new ReverseMarkdown.Config
{
  GithubFlavored = true,
  UnknownTags = ReverseMarkdown.Config.UnknownTagsOption.Bypass
});

Console.OutputEncoding = System.Text.Encoding.UTF8;

string processingPrefix = "       ";

// ----------------------------
// Main loop: process all mailboxes
// ----------------------------
foreach (var user in app.Graph.Users)
{
  Console.WriteLine($"=== Processing user: {user} ===");

  foreach (var link in app.Links)
  {
    var mailbox = user + "/" + link.Mailbox;

    Console.WriteLine($"╰→ Processing mailbox: {mailbox}");

    // Iterate all delta pages (new + changed messages)
    string? deltaLink = db.GetDeltaLink(mailbox);
    await foreach (var page in graph.DeltaPagesAsync(mailbox, deltaLink, app.Graph.BeginDate))
    {
      Console.WriteLine($"  ╰→ Received {page.Messages.Count} incomming messages from delta query for {mailbox}");

      int nr = 1;
      foreach (var msg in page.Messages)
      {
        if (msg.Body == null && msg.Subject == null)
        {
          Console.WriteLine($"    ╰→ Message {nr++} was removed");
          continue;
        }

        Console.WriteLine($"    ╰→ Message {nr++} with Subject: {msg.Subject}");

        bool flowControl = await ProcessIncommingMessage(mailbox, link.Project, msg);

        if (!flowControl)
        {
          Console.WriteLine($"       Skipped message");
          continue;
        }
      }

      // Store new delta token for next run
      if (!string.IsNullOrEmpty(page.DeltaLink))
      {
        deltaLink = page.DeltaLink;
        db.UpsertDeltaLink(mailbox, deltaLink);
        Console.WriteLine($"  ╰→ Updated delta link for {mailbox}");
      }
    }
  }

  // For sent messages
  var sentMailbox = GraphConnector.GetSentMailbox(user);
  Console.WriteLine($"╰→ Processing mailbox: {sentMailbox}");

  string? sentDeltaLink = db.GetDeltaLink(sentMailbox);
  await foreach (var page in graph.DeltaPagesAsync(sentMailbox, sentDeltaLink, app.Graph.BeginDate))
  {

    Console.WriteLine($"  ╰→ Received {page.Messages.Count} outgoing messages from delta query for {sentMailbox}");

    int nr = 1;
    foreach (var msg in page.Messages)
    {
      Console.WriteLine($"    ╰→ Message {nr++} with Subject: {msg.Subject}");

      bool flowControl = await ProcessSentMessage(sentMailbox, msg);

      if (!flowControl)
      {
        Console.WriteLine($"       Skipped message");
        continue;
      }
    }

    // Store new delta token for next run
    if (!string.IsNullOrEmpty(page.DeltaLink))
    {
      sentDeltaLink = page.DeltaLink;
      db.UpsertDeltaLink(sentMailbox, sentDeltaLink);
      Console.WriteLine($"  ╰→ Updated delta link for {sentMailbox}");
    }
  }
}

Console.WriteLine("=== All mailboxes processed. Done. ===");

if (app.Pause)
{
  Console.WriteLine("Press any key to continue...");
  Console.ReadKey(intercept: true);
}
return;

// ----------------------------
// Process a single email
// ----------------------------
async Task<bool> ProcessIncommingMessage(string mailbox, string project, Microsoft.Graph.Models.Message msg)
{
  // Skip if we already processed this message
  if (GraphConnector.WasProcessed(msg))
  {
    Console.WriteLine($"{processingPrefix}Message already processed, skipping.");
    return false;
  }

  // Skip if email was sent by the monitored mailbox itself (avoid loops)
  if (GraphConnector.IsSelf(mailbox, msg))
  {
    Console.WriteLine($"{processingPrefix}Message is self-sent, skipping.");
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
      Console.WriteLine($"{processingPrefix}Updating existing User Story #{existingId}");

      string? description = await TfsConnector.WorkItemExistsingDescriptionAsync(wit, existingId);
      if (description == null)
      {
        Console.WriteLine($"{processingPrefix}User Story #{existingId} not found in TFS. Sending error reply.");
        await graph.SendErrorReplyAsync(mailbox, msg,
          errorBody: string.Format(app.Graph.UsNotFoundTemplate, existingId));

        await graph.CategorizeDoneAsync(msg, mailbox);
        Console.WriteLine($"{processingPrefix}Categorized mail as done.");

        db.MarkProcessed(msg.Id!, mailbox, null, "us-not-found");
        return false;
      }

      var prepared = await Util.PrepareContentAsync(graph, mailbox, msg, mdConverter);

      List<string> history = db.GetStoryContent(existingId);
      history.Add(prepared.html);

      var newDesc = await summaryGenerator.Summarize(description, history);

      await TfsConnector.AddCommentAndAttachmentsAsync(
        wit,
        existingId,
        prepared.html,
        prepared.attachments,
        newDesc
      );

      await graph.SendInfoReplyAsync(mailbox, msg,
        infoBody: string.Format(app.Graph.UsUpdatedTemplate, existingId));
      Console.WriteLine($"{processingPrefix}Updated User Story #{existingId} with new comment/attachments.");

      await graph.CategorizeDoneAsync(msg, mailbox);
      Console.WriteLine($"{processingPrefix}Categorized mail as done.");

      db.MarkProcessed(msg.Id!, mailbox, existingId, "updated", prepared.html);
    }
    else
    {
      // ----------------------------
      // Create new user story
      // ----------------------------
      Console.WriteLine($"{processingPrefix}Creating new User Story from message");

      var prepared = await Util.PrepareContentAsync(graph, mailbox, msg, mdConverter);

      string description = "This UserStory was generated form an E-Mail, see History";

      var history = new List<string> { prepared.html };

      var newDesc = await summaryGenerator.Summarize(description, history);

      int newId = await TfsConnector.CreateUserStoryAsync(
        wit,
        project,
        msg.Subject ?? "(no subject)",
        newDesc
      );

      Console.WriteLine($"{processingPrefix}Created User Story #{newId}");


      await TfsConnector.AddCommentAndAttachmentsAsync(
        wit,
        newId,
        prepared.html,
        prepared.attachments
      );

      Console.WriteLine($"{processingPrefix}Updated User Story #{newId} with new comment/attachments.");

      db.LinkStory(mailbox, newId);

      await graph.SendInfoReplyAsync(mailbox, msg,
        string.Format(app.Graph.UsCreatedTemplate, newId),
        subjectSuffix: $" [US#{newId}]"
      );

      await graph.CategorizeDoneAsync(msg, mailbox);
      Console.WriteLine($"{processingPrefix}Categorized mail as done.");

      db.MarkProcessed(msg.Id!, mailbox, newId, "created", prepared.html);
    }
  }
  catch (Exception ex)
  {
    Console.Error.WriteLine($"Error processing message: {ex}");
    throw; // per requirement: crash and let supervisor retry
  }

  return true;
}

// ----------------------------
// Process a single outgoing email
// ----------------------------
async Task<bool> ProcessSentMessage(string mailbox, Microsoft.Graph.Models.Message msg)
{
  // Skip if we already processed this message
  if (GraphConnector.WasProcessed(msg))
  {
    Console.WriteLine($"{processingPrefix}Message already processed, skipping.");
    return false;
  }

  if (GraphConnector.HasNotificationCategory(msg))
  {
    Console.WriteLine($"{processingPrefix}Message is a self emited notification, skipping.");
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
      Console.WriteLine($"{processingPrefix}Updating existing User Story #{existingId}");

      string? description = await TfsConnector.WorkItemExistsingDescriptionAsync(wit, existingId);
      if (description == null)
      {
        await graph.CategorizeDoneAsync(msg, mailbox);
        Console.WriteLine($"{processingPrefix}Categorized mail as done.");

        db.MarkProcessed(msg.Id!, mailbox, null, "us-not-found");
        return false;
      }

      var prepared = await Util.PrepareContentAsync(graph, mailbox, msg, mdConverter);

      List<string> history = db.GetStoryContent(existingId);
      history.Add(prepared.html);

      var newDesc = await summaryGenerator.Summarize(description, history);

      await TfsConnector.AddCommentAndAttachmentsAsync(
        wit,
        existingId,
        prepared.html,
        prepared.attachments,
        newDesc
      );

      Console.WriteLine($"{processingPrefix}Updated User Story #{existingId} with new comment/attachments.");

      await graph.CategorizeDoneAsync(msg, mailbox);
      Console.WriteLine($"{processingPrefix}Categorized mail as done.");

      db.MarkProcessed(msg.Id!, mailbox, existingId, "updated", prepared.html);
    }
    else
    {
      Console.WriteLine($"{processingPrefix}Sent mail withoutUser Story reference is ignored.");
    }
  }
  catch (Exception ex)
  {
    Console.Error.WriteLine($"{processingPrefix}Error processing message: {ex}");
    throw; // per requirement: crash and let supervisor retry
  }

  return true;
}