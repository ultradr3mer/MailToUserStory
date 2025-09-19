# MailToUserStory

## Project file
```
<Project Sdk="Microsoft.NET.Sdk">

  <PropertyGroup>
    <OutputType>Exe</OutputType>
    <TargetFramework>net8.0-windows</TargetFramework>
    <ImplicitUsings>enable</ImplicitUsings>
    <Nullable>enable</Nullable>
    <UserSecretsId>68b2f0b9-9306-4ce9-b25e-0f98c0c6d1c4</UserSecretsId>
  </PropertyGroup>

  <ItemGroup>
    <PackageReference Include="Azure.Identity" Version="1.15.0" />
    <PackageReference Include="HtmlAgilityPack" Version="1.12.2" />
    <PackageReference Include="Microsoft.Data.Sqlite" Version="9.0.8" />
    <PackageReference Include="Microsoft.Extensions.Configuration" Version="9.0.8" />
    <PackageReference Include="Microsoft.Extensions.Configuration.Binder" Version="9.0.8" />
    <PackageReference Include="Microsoft.Extensions.Configuration.Json" Version="9.0.8" />
    <PackageReference Include="Microsoft.Extensions.Configuration.UserSecrets" Version="9.0.8" />
    <PackageReference Include="Microsoft.Graph" Version="5.91.0" />
    <PackageReference Include="Microsoft.TeamFoundationServer.Client" Version="19.225.1" />
    <PackageReference Include="Microsoft.VisualStudio.Services.Client" Version="19.225.1" />
    <PackageReference Include="ReverseMarkdown" Version="4.7.0" />
  </ItemGroup>

  <ItemGroup>
    <None Update="appsettings.local.json">
      <CopyToOutputDirectory>PreserveNewest</CopyToOutputDirectory>
    </None>
    <None Update="appsettings.json">
      <CopyToOutputDirectory>PreserveNewest</CopyToOutputDirectory>
    </None>
  </ItemGroup>

</Project>

```

## Structure
```
MailToUserStory
├── Data
│   ├── AttachmentContainer.cs
│   ├── AttachmentPayload.cs
│   └── DeltaPage.cs
├── AppConfig.cs
├── appsettings.json
├── appsettings.local.json
├── Db.cs
├── GraphConnector.cs
├── MailToUserStory.csproj
├── OllamaClient.cs
├── Program.cs
├── SummaryGenerator.cs
├── TfsConnector.cs
└── Util.cs
```

## Files

### AppConfig.cs

```
namespace MailToUserStory
{
    public record AppConfig
    {
        public required GraphConfig Graph { get; set; }
        public required TfsConfig Tfs { get; set; }
        public string DatabasePath { get; set; }
        public required OllamaConfig Ollama { get; set; }

        public record GraphConfig
        {
            public required string TenantId { get; set; }
            public required string ClientId { get; set; }
            public string? ClientSecret { get; set; }
            public required string[] Mailboxes { get; set; }
            public required string UsCreatedTemplate { get; set; }
            public required string UsUpdatedTemplate { get; set; }
            public required string UsNotFoundTemplate { get; set; }
        }

        public record TfsConfig
        {
            public required string BaseUrl { get; set; }
            public required string ProjectCollection { get; set; }
            public required string Project { get; set; }
            public string? User { get; set; }
            public string Password { get; set; }
        }

        public record OllamaConfig
        {
            public string Host { get; set; }
            public string Model { get; set; }
            public string SummarizeInstruction { get; set; }
            public bool Enabled { get; set; }
        }
    }
}
```

### Data/AttachmentContainer.cs

```
using Microsoft.Graph.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MailToUserStory.Data
{
    // AttachmentContainer is a lightweight data holder that bundles two distinct collections of FileAttachment objects. The InlineAttachments list is intended for attachments that should be embedded directly within the body of a message—think images or documents that appear inline in an email or chat bubble. The FileAttachments list holds attachments that are meant to be downloaded or opened separately, such as PDFs, spreadsheets, or other binary files. By marking both properties as required, the class guarantees that any consumer will always receive a fully populated container, preventing null‑reference pitfalls.
    // 
    // In practice, AttachmentContainer is typically instantiated by higher‑level services that assemble a message payload. For example, an email‑sending service might create an AttachmentContainer, populate its InlineAttachments with images referenced in the HTML body, and fill FileAttachments with any additional files the user wants to send. The container is then passed to a transport layer or a serialization routine that converts the attachments into the appropriate MIME parts or multipart/form‑data payloads. UI components can also consume the container to render a list of attachments, offering preview thumbnails for inline items and download links for file attachments. Because the class depends only on the FileAttachment type, it remains agnostic of the underlying storage or transmission mechanism, making it a clean, reusable contract across the application.
    public class AttachmentContainer
    {
        public required List<FileAttachment> InlineAttachments { get; set; }
        public required List<FileAttachment> FileAttachments { get; set; }
    }
}
```

### Data/AttachmentPayload.cs

```
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MailToUserStory.Data
{
    // AttachmentPayload is a lightweight, immutable data carrier that bundles the essential information for a file attachment: the original file name and its raw byte content. Because it is sealed and uses init‑only properties, it cannot be subclassed or mutated after construction, ensuring that the payload remains consistent throughout its lifecycle. In practice, this class is typically instantiated by higher‑level components that receive or generate file data—such as an API controller handling an upload, a service that prepares an email with attachments, or a persistence layer that writes files to disk or a database. Once created, the payload is passed along to other parts of the system—perhaps a storage service that persists the Bytes, a messaging component that transmits the attachment, or a UI layer that renders the file for download—without exposing the underlying byte array to accidental modification. The required keyword guarantees that every instance contains both a FileName and Bytes, making it safe to use in contexts where a complete attachment representation is mandatory.
    public sealed class AttachmentPayload
    {
        public required string FileName { get; set; }
        public required byte[] Bytes { get; set; }
    }
}
```

### Data/DeltaPage.cs

```
using Microsoft.Graph.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MailToUserStory.Data
{
    // DeltaPage is a lightweight, immutable container that represents a single page of results in a delta‑query workflow. It is sealed, so it cannot be subclassed, ensuring that its shape remains stable across the codebase. The core of the page is the required Messages list, which holds the actual payload objects – each entry is a Message instance that the caller will process or display. Because the property is marked required, any attempt to construct a DeltaPage without supplying a Messages collection will fail at compile time, guaranteeing that the page always carries meaningful data.
    // 
    // The optional NextLink and DeltaLink strings are part of the pagination and change‑tracking contract. NextLink is a URI that the client can follow to retrieve the next page of results when the current page is incomplete. DeltaLink, on the other hand, is a token that the client can store and later use to resume a delta query from the point where the current page left off, ensuring that only new or modified items are returned in subsequent calls.
    // 
    // In practice, DeltaPage is produced by a data‑access layer or an HTTP client that talks to an external service (for example, a Microsoft Graph endpoint). The service layer parses the raw response, populates the Messages list, and extracts the NextLink and DeltaLink values from the response payload. The calling code then consumes the DeltaPage, iterating over the Messages, and if NextLink is present, issuing another request to fetch the following page. When the caller is ready to perform a new delta sync, it can use the stored DeltaLink to start from the last known state, avoiding reprocessing of unchanged data. Thus, DeltaPage acts as the bridge between the raw API response and the higher‑level business logic that needs to handle incremental updates efficiently.
    public sealed class DeltaPage
    {
        public required List<Message> Messages { get; set; }
        public string? NextLink { get; set; }
        public string? DeltaLink { get; set; }
    }
}
```

### Db.cs

```
using Microsoft.Data.Sqlite;

// The `Db` class is a lightweight data‑access wrapper that talks directly to a SQLite database. It implements `IDisposable` so callers can use it in a `using` block or rely on deterministic cleanup. Internally it keeps a single `SqliteConnection` that is opened in the constructor and closed when `Dispose` is called.
// 
// Its public surface is a set of domain‑specific operations:
// 
// * **Delta link handling** – `GetDeltaLink` looks up the last known delta link for a mailbox, while `UpsertDeltaLink` writes or updates that value. These methods are used by code that synchronises mailboxes with an external service, ensuring the next sync starts from the correct point.
// 
// * **Processed‑email tracking** – `WasProcessed` checks whether a particular message ID has already been handled, preventing duplicate work. `MarkProcessed` records the outcome of processing a message, storing the message ID, mailbox, optional work‑item reference, timestamp, outcome status, and any content that should be persisted. This table is the audit trail for all email processing.
// 
// * **Story linking** – `LinkStory` associates a work‑item ID with a mailbox. The `Stories` table is a many‑to‑many bridge that lets the system later retrieve all mailboxes that contributed to a particular story.
// 
// * **Story content retrieval** – `GetStoryContent` pulls the raw content of all processed emails that belong to a given work‑item. This is useful for assembling the full narrative of a story from its constituent messages.
// 
// * **Schema bootstrap** – `InitializeSchema` creates the three tables (`Mailboxes`, `Stories`, `ProcessedEmails`) if they do not already exist. This method is typically called once at application start‑up or during a migration step.
// 
// The class is intentionally decoupled from higher‑level business logic. Other components—such as an email ingestion service, a delta‑sync orchestrator, or a story‑assembly module—call these methods to persist state and query historical data. By encapsulating all SQL statements and transaction handling, `Db` provides a clean, reusable persistence layer that other parts of the system can rely on without needing to know the details of SQLite or SQL syntax.
sealed class Db : IDisposable
{
    private readonly SqliteConnection _conn;
    public Db(string path)
    {
    }

    public SqliteConnection Connection { get; set; }

    public void Dispose()
    {
    }

    public string? GetDeltaLink(string mailbox)
    {
    }

    public void UpsertDeltaLink(string mailbox, string? delta)
    {
    }

    public bool WasProcessed(string messageId)
    {
    }

    public void MarkProcessed(string messageId, string mailbox, int? workItemId, string outcome, string content = "")
    {
    }

    public void LinkStory(string mailbox, int workItemId)
    {
    }

    public List<string> GetStoryContent(int id)
    {
    }

    public static void InitializeSchema(Db db)
    {
    }
}
```

### GraphConnector.cs

```
using MailToUserStory.Data;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using System;
using DeltaGetResponse = Microsoft.Graph.Users.Item.MailFolders.Item.Messages.Delta.DeltaGetResponse;

namespace MailToUserStory
{
    // GraphConnector is a static helper that orchestrates all Microsoft Graph‑based mail operations for a given mailbox. It relies on a GraphServiceClient instance to issue REST calls, and it works with core Graph types such as Message, Attachment, FileAttachment, ItemBody, and the delta‑response classes that Microsoft Graph exposes. The class parses a mailbox string into a user and a folder via the internal GraphUserFolder helper, then resolves the folder’s ID (handling the well‑known “Inbox” and “SentItems” names or walking the child‑folder hierarchy under msgfolderroot).  
    // 
    // The DeltaPagesAsync method streams incremental changes for a folder’s messages. It first obtains the folder ID, then either resumes from a stored deltaLink or starts a new delta query. It consumes the delta response, yielding DeltaPage objects that contain the current batch of Message objects, a nextLink for paging, and a deltaLink for future resumptions. The method uses Kiota’s RequestInformation to follow nextLinks and deltaLinks directly.  
    // 
    // GetFileAttachmentsAsync pulls all attachments for a specific message, separating inline attachments from regular file attachments, and returns them in an AttachmentContainer.  
    // 
    // SendInfoReplyAsync builds a reply Message addressed to the original sender, optionally appending a subject suffix, and tags it with the “MailToTfs‑Notification” category. It posts the message via the SendMail endpoint, ensuring it is saved to SentItems. SendErrorReplyAsync is a thin wrapper that simply forwards an error body to SendInfoReplyAsync.  
    // 
    // HasNotificationCategory checks whether a message carries the notification category, while GetSentMailbox constructs the path to the SentItems folder for a user. Throughout, the class interacts with GraphServiceClient, the Graph request builders, and the Kiota abstractions to perform paging, delta tracking, and message manipulation, providing a cohesive API for mail‑centric workflows.
    public static class GraphConnector
    {
        private const string InboxWellKnownFolderName = "Inbox";
        private const string SentItemsWellKnownFolderName = "SentItems";
        private const string MailToTfsNotificationCategoryName = "MailToTfs-Notification";
        public static async IAsyncEnumerable<DeltaPage> DeltaPagesAsync(GraphServiceClient graph, string mailbox, string? deltaLink)
        {
        }

        private static async Task<string?> ResolveFolderIdAsync(GraphServiceClient graph, string user, string folderName)
        {
        }

        public static bool IsSelf(string mailbox, Message msg)
        {
        }

        public static async Task<AttachmentContainer> GetFileAttachmentsAsync(GraphServiceClient graph, string mailbox, string messageId)
        {
        }

        public static async Task SendInfoReplyAsync(GraphServiceClient graph, string mailbox, Message original, string infoBody, string? subjectSuffix = null)
        {
        }

        public static Task SendErrorReplyAsync(GraphServiceClient graph, string mailbox, Message original, string errorBody)
        {
        }

        internal static string GetSentMailbox(string mailbox)
        {
        }

        internal static bool HasNotificationCategory(Message msg)
        {
        }

        private class GraphUserFolder
        {
            public GraphUserFolder(string mailbox)
            {
            }

            public string User { get; set; }
            public string Folder { get; set; }
        }
    }
}
```

### OllamaClient.cs

```
using Microsoft.Graph.Models.Security;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace MailToUserStory
{
    // The `OllamaClient` is an internal helper that talks to an Ollama inference server. It holds a `HttpClient` for making HTTP requests, a `Regex` for normalising whitespace, and a few configuration strings: the server host, the model name, and a default instruction for generating summaries.  
    // 
    // When another part of the application wants a summary, it calls `GenerateSummary`, passing the conversation history. That method forwards the history and the pre‑configured instruction to `Complete`. `Complete` builds a prompt by concatenating the history with a minified version of the instruction (whitespace collapsed to single spaces). It then serialises a payload containing the model, prompt, and a flag to disable streaming, and posts it to the Ollama `/api/generate` endpoint.  
    // 
    // The response is read as a stream, parsed into a `JsonDocument`, and the `"response"` field is extracted and returned. If the field is missing, the whole JSON is returned as a fallback.  
    // 
    // Internally, the class relies on standard .NET types such as `StringBuilder`, `Encoding`, `JsonSerializer`, and `JsonDocument`. It is designed to be used by higher‑level components that need to generate text completions or summaries without dealing with HTTP or JSON handling directly.
    internal class OllamaClient
    {
        private static Regex condenseSpaces = new Regex(@"\s+", RegexOptions.Compiled);
        private readonly string host;
        private readonly HttpClient http;
        private readonly string model;
        private readonly string generateSummaryInstruction;
        public OllamaClient(string host, string model, string generateSummaryInstruction)
        {
        }

        internal Task<string> GenerateSummary(string history)
        {
        }

        private async Task<string> Complete(string context, string instruction)
        {
        }

        private static string Minify(string text)
        {
        }
    }
}
```

### Program.cs

```
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
var config = new ConfigurationBuilder().AddJsonFile("appsettings.json", optional: false).AddJsonFile("appsettings.local.json", optional: true).AddUserSecrets(typeof(AppConfig).Assembly, optional: true).Build();
// Load config into strongly typed model
var app = config.Get<AppConfig>() ?? throw new Exception("Invalid configuration");
app = app with
{
    Graph = app.Graph with
    {
        ClientSecret = config["Graph:ClientSecret"] ?? app.Graph.ClientSecret
    },
    Tfs = app.Tfs with
    {
        Password = config["Tfs:Password"] ?? app.Tfs.Password
    }
};
// Verify secrets exist
Util.Assert(!string.IsNullOrWhiteSpace(app.Graph.ClientSecret), "Graph:ClientSecret missing (user-secrets)");
Util.Assert(!string.IsNullOrWhiteSpace(app.Tfs.Password), "Tfs:Password missing (user-secrets)");
// SQLite DB for state (delta tokens, processed mails, lock)
using var db = new Db(app.DatabasePath);
Db.InitializeSchema(db);
// Create Graph SDK client
var credential = new ClientSecretCredential(app.Graph.TenantId, app.Graph.ClientId, app.Graph.ClientSecret);
var scopes = new[]
{
    "https://graph.microsoft.com/.default"
};
var graph = new GraphServiceClient(credential, scopes);
// Create TFS client
var tfsUri = new Uri(app.Tfs.BaseUrl);
var connection = new VssConnection(tfsUri, new VssBasicCredential(app.Tfs.User, app.Tfs.Password));
var wit = await connection.GetClientAsync<WorkItemTrackingHttpClient>();
// Create Ollama client
var ollama = new OllamaClient(app.Ollama.Host, app.Ollama.Model, app.Ollama.SummarizeInstruction);
// Initilize Summary Generator
var summaryGenerator = new SummaryGenerator(ollama, app.Ollama.Enabled);
// Markdown converter (for logs/debugging, not persisted to TFS)
var mdConverter = new ReverseMarkdown.Converter(new ReverseMarkdown.Config { GithubFlavored = true, UnknownTags = ReverseMarkdown.Config.UnknownTagsOption.Bypass });
// ----------------------------
// Main loop: process all mailboxes
// ----------------------------
foreach (var mailbox in app.Graph.Mailboxes)
{
    Console.WriteLine($"=== Processing mailbox: {mailbox} ===");
    // Iterate all delta pages (new + changed messages)
    string? deltaLink = db.GetDeltaLink(mailbox);
    await foreach (var page in GraphConnector.DeltaPagesAsync(graph, mailbox, deltaLink))
    {
        Console.WriteLine($"Received {page.Messages.Count} incomming messages from delta query for {mailbox}");
        int nr = 1;
        foreach (var msg in page.Messages)
        {
            Console.WriteLine($"--> Message {nr++} with Subject: {msg.Subject}");
            bool flowControl = await ProcessIncommingMessage(mailbox, msg);
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

    // For sent messages
    var sentMailbox = GraphConnector.GetSentMailbox(mailbox);
    string? sentDeltaLink = db.GetDeltaLink(sentMailbox);
    await foreach (var page in GraphConnector.DeltaPagesAsync(graph, sentMailbox, sentDeltaLink))
    {
        Console.WriteLine($"Received {page.Messages.Count} outgoing messages from delta query for {sentMailbox}");
        int nr = 1;
        foreach (var msg in page.Messages)
        {
            Console.WriteLine($"--> Message {nr++} with Subject: {msg.Subject}");
            bool flowControl = await ProcessSentMessage(sentMailbox, msg);
            if (!flowControl)
            {
                Console.WriteLine($"Skipped message");
                continue;
            }
        }

        // Store new delta token for next run
        if (!string.IsNullOrEmpty(page.DeltaLink))
        {
            sentDeltaLink = page.DeltaLink;
            db.UpsertDeltaLink(sentMailbox, sentDeltaLink);
            Console.WriteLine($"Updated delta link for {sentMailbox}");
        }
    }
}

Console.WriteLine("All mailboxes processed. Done.");
return;
// ----------------------------
// Process a single email
// ----------------------------
async Task<bool> ProcessIncommingMessage(string mailbox, Microsoft.Graph.Models.Message msg)
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
            Console.WriteLine($"Updating existing User Story #{existingId}");
            string? description = await TfsConnector.WorkItemExistsingDescriptionAsync(wit, existingId);
            if (description == null)
            {
                Console.WriteLine($"User Story #{existingId} not found in TFS. Sending error reply.");
                await GraphConnector.SendErrorReplyAsync(graph, mailbox, msg, errorBody: string.Format(app.Graph.UsNotFoundTemplate, existingId));
                db.MarkProcessed(msg.Id!, mailbox, null, "us-not-found");
                return false;
            }

            var prepared = await Util.PrepareContentAsync(graph, mailbox, msg, mdConverter);
            List<string> history = db.GetStoryContent(existingId);
            history.Add(prepared.html);
            var newDesc = await summaryGenerator.Summarize(description, history);
            await TfsConnector.AddCommentAndAttachmentsAsync(wit, app.Tfs.Project, existingId, prepared.html, prepared.attachments, newDesc);
            await GraphConnector.SendInfoReplyAsync(graph, mailbox, msg, infoBody: string.Format(app.Graph.UsUpdatedTemplate, existingId));
            Console.WriteLine($"Updated User Story #{existingId} with new comment/attachments.");
            db.MarkProcessed(msg.Id!, mailbox, existingId, "updated", prepared.html);
        }
        else
        {
            // ----------------------------
            // Create new user story
            // ----------------------------
            Console.WriteLine($"Creating new User Story from message");
            var prepared = await Util.PrepareContentAsync(graph, mailbox, msg, mdConverter);
            string description = "This UserStory was generated form an E-Mail, see History";
            var history = new List<string>
            {
                prepared.html
            };
            var newDesc = await summaryGenerator.Summarize(description, history);
            int newId = await TfsConnector.CreateUserStoryAsync(wit, app.Tfs.Project, msg.Subject ?? "(no subject)", newDesc);
            Console.WriteLine($"Created User Story #{newId}");
            await TfsConnector.AddCommentAndAttachmentsAsync(wit, app.Tfs.Project, newId, prepared.html, prepared.attachments);
            Console.WriteLine($"Updated User Story #{newId} with new comment/attachments.");
            db.LinkStory(mailbox, newId);
            await GraphConnector.SendInfoReplyAsync(graph, mailbox, msg, string.Format(app.Graph.UsCreatedTemplate, newId), subjectSuffix: $" [US#{newId}]");
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
    if (db.WasProcessed(msg.Id!))
    {
        Console.WriteLine($"Message already processed, skipping.");
        return false;
    }

    if (GraphConnector.HasNotificationCategory(msg))
    {
        Console.WriteLine($"Message is a self emited notification, skipping.");
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
            Console.WriteLine($"Updating existing User Story #{existingId}");
            string? description = await TfsConnector.WorkItemExistsingDescriptionAsync(wit, existingId);
            if (description == null)
            {
                db.MarkProcessed(msg.Id!, mailbox, null, "us-not-found");
                return false;
            }

            var prepared = await Util.PrepareContentAsync(graph, mailbox, msg, mdConverter);
            List<string> history = db.GetStoryContent(existingId);
            history.Add(prepared.html);
            var newDesc = await summaryGenerator.Summarize(description, history);
            await TfsConnector.AddCommentAndAttachmentsAsync(wit, app.Tfs.Project, existingId, prepared.html, prepared.attachments, newDesc);
            Console.WriteLine($"Updated User Story #{existingId} with new comment/attachments.");
            db.MarkProcessed(msg.Id!, mailbox, existingId, "updated", prepared.html);
        }
        else
        {
            Console.WriteLine($"Sent mail withoutUser Story reference is ignored.");
            db.MarkProcessed(msg.Id!, mailbox, null, "ignored");
        }
    }
    catch (Exception ex)
    {
        Console.Error.WriteLine($"Error processing message: {ex}");
        throw; // per requirement: crash and let supervisor retry
    }

    return true;
}
```

### SummaryGenerator.cs

```
using System.Text;
using HtmlAgilityPack;

namespace MailToUserStory
{
    // The `SummaryGenerator` class is a helper that orchestrates the creation of an AI‑generated summary for a given email thread. It receives an `OllamaClient` instance, which is responsible for communicating with the underlying language model, and a boolean flag that toggles the summarization feature on or off.
    // 
    // When the `Summarize` method is called, the class first checks whether summarization is enabled. If it is disabled, the method simply returns the current description, optionally stripping out any previously inserted AI marker. If summarization is enabled, the method prepares the conversation history for the model: each entry is sanitized via `Util.SanitizeHtmlForLlm`, appended with a line break and a custom delimiter, and finally the entire history is joined into a single string. This string is then passed to the `OllamaClient.GenerateSummary` method, which returns the AI‑generated summary text.
    // 
    // After receiving the summary, the class inserts it into the current description. If the description already contains the marker `==== AI Generated Summary ====`, the existing marker and any following text are replaced with the new summary. Otherwise, the marker and the new summary are appended to the end of the description. The final output is a string that combines the original content, the marker, and the AI summary, separated by HTML line breaks.
    // 
    // In short, `SummaryGenerator` acts as a thin wrapper that sanitizes input, delegates the heavy lifting to `OllamaClient`, and formats the result for display, while respecting a feature toggle and handling previously inserted AI markers.
    internal class SummaryGenerator
    {
        public SummaryGenerator(OllamaClient client, bool enabled)
        {
        }

        private const string AI_MARKER = "==== AI Generated Summary ====";
        private readonly OllamaClient client;
        private readonly bool enabled;
        public async Task<string> Summarize(string currentDescription, List<string> history)
        {
        }
    }
}
```

### TfsConnector.cs

```
using MailToUserStory.Data;
using Microsoft.TeamFoundation.WorkItemTracking.WebApi;
using Microsoft.TeamFoundation.WorkItemTracking.WebApi.Models;
using Microsoft.VisualStudio.Services.WebApi.Patch.Json;
using System.Net;
using Microsoft.VisualStudio.Services.WebApi.Patch;

namespace MailToUserStory
{
    // The TfsConnector class is a small helper that talks to Azure DevOps (TFS) through the WorkItemTrackingHttpClient API. It exposes four asynchronous operations that are used by higher‑level code to read, create, and update work items and to attach files.
    // 
    // 1. WorkItemExistsingDescriptionAsync  
    //    – Calls GetWorkItemAsync to fetch a work item by its numeric ID, requesting the full field set.  
    //    – Extracts the value of the System.Description field and returns it as a string, or null if the item is not found (HTTP 404).  
    //    – The method swallows the VssServiceResponseException for a 404 and propagates other errors.
    // 
    // 2. CreateUserStoryAsync  
    //    – Builds a JSON‑patch document that adds a title and a Markdown description to a new work item.  
    //    – Sends the patch to CreateWorkItemAsync with the project name and the work‑item type “User Story”.  
    //    – Returns the newly created work‑item ID, throwing if the API does not supply one.
    // 
    // 3. AddCommentAndAttachmentsAsync  
    //    – Prepares a patch that appends a comment to the System.History field.  
    //    – Optionally replaces the System.Description field if a new description is supplied.  
    //    – For each attachment payload, it streams the byte array to CreateAttachmentAsync, obtains the attachment URL, and adds a relation of type AttachedFile to the patch.  
    //    – Finally, it calls UpdateWorkItemAsync to apply the patch to the target work item.
    // 
    // 4. AddAttachmentsAsync  
    //    – Similar to the attachment part of the previous method but only handles file uploads.  
    //    – Builds a patch that adds all attachment relations and updates the work item.
    // 
    // Throughout, the class relies on the Azure DevOps client library types: WorkItemTrackingHttpClient, JsonPatchDocument, JsonPatchOperation, WorkItemRelation, and a custom AttachmentPayload structure that holds file bytes and a file name. The connector abstracts the low‑level HTTP calls and patch construction, allowing other parts of the application to work with work items and attachments without dealing with the intricacies of the REST API.
    public static class TfsConnector
    {
        public static async Task<string?> WorkItemExistsingDescriptionAsync(WorkItemTrackingHttpClient wit, int id)
        {
        }

        public static async Task<int> CreateUserStoryAsync(WorkItemTrackingHttpClient wit, string project, string title, string descriptionMarkdown)
        {
        }

        public static async Task AddCommentAndAttachmentsAsync(WorkItemTrackingHttpClient wit, string project, int id, string commentMarkdown, List<AttachmentPayload> attachments, string? updatedDescription = null)
        {
        }

        public static async Task AddAttachmentsAsync(WorkItemTrackingHttpClient wit, string project, int id, List<AttachmentPayload> attachments)
        {
        }
    }
}
```

### Util.cs

```
using HtmlAgilityPack;
using MailToUserStory.Data;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using System.Net;
using System.Text.RegularExpressions;

namespace MailToUserStory
{
    // The Util class is a helper hub that stitches together several parts of an email‑processing pipeline. It pulls raw data from Microsoft Graph, massages it into a form that can be stored or displayed, and offers a few small utilities for parsing and validation.
    // 
    // When PrepareContentAsync is called, it first asks GraphConnector to fetch all file attachments for a given message. Those attachments are wrapped into AttachmentPayload objects that hold the file name and raw bytes. The method then sanitises the message body: if the body is plain text it is returned as‑is; if it is HTML it is cleaned of scripts and styles, inline images are re‑encoded to JPEG and embedded as data URLs, and the resulting HTML is returned together with the attachment list. The HTML is built line by line, adding German header lines such as “Von”, “An”, “Betreff” and “Gesendet”, and the body content is appended after a line break.
    // 
    // SanitizeHtmlForHistory handles the heavy lifting of turning an ItemBody into safe HTML. It replaces inline image references (cid:…) with base64 JPEG data, strips out script and style tags, and returns the inner HTML. SanitizeHtmlForLlm turns arbitrary HTML into plain text suitable for a language model: it replaces <br> and <hr> with newlines, decodes HTML entities, collapses multiple blank lines, and trims the result.
    // 
    // ReEncode is a small image helper that forces any inline image to JPEG at 70 % quality, ensuring a consistent format for downstream consumers.
    // 
    // ParseUserStoryId looks for a pattern like “[US#123]” in an email subject and extracts the numeric ID, returning null if the pattern is absent.
    // 
    // Assert is a simple guard that throws an exception with a custom message when a condition fails.
    // 
    // Throughout, the class relies on Microsoft Graph SDK types (GraphServiceClient, Message, ItemBody, FileAttachment), the ReverseMarkdown library for converting Markdown to HTML, HtmlAgilityPack for parsing and cleaning HTML, and System.Drawing for image re‑encoding. It does not maintain state; all methods are static and operate purely on the supplied parameters, making it a stateless utility layer that other parts of the application can call to prepare email content for storage, display, or AI processing.
    public static class Util
    {
        public static async Task<(string html, List<AttachmentPayload> attachments)> PrepareContentAsync(GraphServiceClient graph, string mailbox, Message msg, ReverseMarkdown.Converter converter)
        {
        }

        public static string SanitizeHtmlForHistory(ItemBody? body, ReverseMarkdown.Converter converter, List<FileAttachment> inlineAttachments)
        {
        }

        public static string SanitizeHtmlForLlm(string html)
        {
        }

        private static byte[] ReEncode(byte[] bytes)
        {
        }

        public static int? ParseUserStoryId(string? subject)
        {
        }

        public static void Assert(bool condition, string message)
        {
        }
    }
}
```


