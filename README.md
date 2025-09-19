# Mail‑to‑User‑Story

> **Author:** Clara  
> **License:** MIT (see [LICENSE](LICENSE))  

A lightweight .NET 8 command‑line tool that turns e‑mails into Team Foundation Server (TFS) *User Stories*.  
It watches a set of mailboxes via Microsoft Graph, creates or updates work items in TFS, attaches the e‑mail body and any attachments, and optionally generates an AI‑powered summary of the conversation using an Ollama model.

---

## Table of Contents

- [What it does](#what-it-does)
- [Why you’ll love it](#why-youll-love-it)
- [Architecture & key components](#architecture--key-components)
- [Getting started](#getting-started)
  - [Prerequisites](#prerequisites)
  - [Configuration](#configuration)
  - [Secrets](#secrets)
  - [Building & running](#building--running)
- [How it works](#how-it-works)
  - [Processing incoming mail](#processing-incoming-mail)
  - [Processing outgoing mail](#processing-outgoing-mail)
  - [Summary generation](#summary-generation)
- [Extending / Customising](#extending--customising)
- [Troubleshooting](#troubleshooting)
- [Contributing](#contributing)
- [License](#license)
- [Acknowledgements](#acknowledgements)

---

## What it does

| Feature | Description |
|---------|-------------|
| **E‑mail → TFS** | Detects a `[US#123]` marker in the subject line. If the ID exists, it appends a comment and attachments to the work item; otherwise it creates a new *User Story* and replies with a confirmation. |
| **Outgoing mail → TFS** | When you send an e‑mail that contains a `[US#123]` reference, the tool automatically updates the corresponding work item with the e‑mail body and attachments. |
| **State persistence** | Uses a local SQLite database to store Graph delta tokens, processed message IDs, and a simple lock for concurrent runs. |
| **AI summarisation** | Optional AI‑generated summary of the conversation using an Ollama model. The summary is inserted into the work item description. |
| **Attachment handling** | All file attachments are downloaded, re‑encoded to JPEG (70 % quality) if inline, and stored as `AttachmentPayload` objects. |
| **HTML sanitisation** | The e‑mail body is cleaned of scripts/styles, inline images are embedded as data‑URLs, and the resulting HTML is safe for display or storage. |

---

## Architecture & key components

```
src/
├─ AppConfig.cs          → Configuration POCO
├─ Db.cs                 → SQLite wrapper (delta tokens, processed mails)
├─ GraphConnector.cs     → Microsoft Graph SDK helper
├─ TfsConnector.cs       → Azure DevOps / TFS helper
├─ Util.cs               → E‑mail body & attachment processing
├─ SummaryGenerator.cs   → AI summarisation wrapper
├─ SummaryGenerator.cs   → Ollama client
└─ Program.cs            → Main loop (CLI)
```

* **`GraphConnector`** – Handles delta queries for both inbox and sent items, fetches attachments, and sends notification e‑mails.
* **`TfsConnector`** – Builds JSON‑patch documents to create/update work items and upload attachments.
* **`Util`** – Pulls raw data from Graph, sanitises HTML, re‑encodes images, extracts `[US#123]` IDs, and provides simple assertions.
* **`SummaryGenerator`** – Delegates to `OllamaClient` to generate a summary and injects it into the work‑item description.
* **`Db`** – Keeps track of processed message IDs and the latest delta token for each mailbox.

---

## Getting started

### Prerequisites

| Item | Minimum version | Notes |
|------|-----------------|-------|
| .NET SDK | 8.0 | Download from <https://dotnet.microsoft.com/download> |
| Microsoft Graph app registration | – | Must have `Mail.ReadWrite`, `Mail.ReadWrite.Shared`, `Mail.ReadWrite.All` scopes |
| Azure DevOps / TFS instance | – | Project name, user & password |
| Ollama server (optional) | – | For AI summarisation |
| SQLite (bundled with .NET) | – | No external dependency |

### Configuration

Create an `appsettings.json` (or `appsettings.Development.json`) in the project root:

```json
{
  "DatabasePath": "state.db",
  "Graph": {
    "TenantId": "YOUR_TENANT_ID",
    "ClientId": "YOUR_CLIENT_ID",
    "ClientSecret": "YOUR_CLIENT_SECRET",
    "UsCreatedTemplate": "User Story #{0} created successfully.",
    "UsUpdatedTemplate": "User Story #{0} updated.",
    "UsNotFoundTemplate": "User Story #{0} not found."
  },
  "Tfs": {
    "BaseUrl": "https://dev.azure.com/yourorg",
    "User": "your-username",
    "Password": "your-password",
    "Project": "YourProject"
  },
  "Ollama": {
    "Host": "http://localhost:11434",
    "Model": "llama3",
    "SummarizeInstruction": "Summarise the following conversation:",
    "Enabled": true
  }
}
```

> **NOTE**  
> `ClientSecret` and `Password` are sensitive and should **not** be committed.  
> Use **user‑secrets** or environment variables to override them:

```bash
dotnet user-secrets set Graph:ClientSecret "your-client-secret"
dotnet user-secrets set Tfs:Password "your-tfs-password"
```

### Building

```bash
dotnet build -c Release
```

### Running

```bash
dotnet run --project src
```

The tool will:

1. Initialise the SQLite database (`state.db` by default).
2. Create a GraphServiceClient with the supplied credentials.
3. Create a TFS connection.
4. Process each mailbox listed in `Graph.Mailboxes`:
   * Pull new/changed messages via delta queries.
   * Create or update User Stories in TFS.
   * Attach e‑mail body and any attachments.
   * Send a confirmation e‑mail back to the sender.
5. Repeat the same for the *sent* folder to keep the work items in sync.

The console output gives a step‑by‑step trace of what is happening.

---

## How it works – a deeper look

### 1. Incoming e‑mail

* **Duplicate check** – `Db.WasProcessed(msg.Id)` prevents re‑processing.
* **Self‑sent check** – `GraphConnector.IsSelf` skips messages sent by the monitored mailbox.
* **Story ID extraction** – `Util.ParseUserStoryId(msg.Subject)` looks for `[US#123]`.
* **Existing story** – If the ID exists:
  * Retrieve current description (`TfsConnector.WorkItemExistsingDescriptionAsync`).
  * Prepare the new comment (`Util.PrepareContentAsync`).
  * Summarise the thread (`SummaryGenerator.Summarize`).
  * Add comment & attachments (`TfsConnector.AddCommentAndAttachmentsAsync`).
  * Reply with `GraphConnector.SendInfoReplyAsync`.
* **New story** – If no ID:
  * Create a new work item (`TfsConnector.CreateUserStoryAsync`).
  * Add the first comment & attachments.
  * Link the mailbox to the story (`Db.LinkStory`).
  * Reply with a confirmation.

### 2. Outgoing e‑mail

* Similar logic but only updates an existing story if the subject contains a `[US#123]` reference.
* Self‑sent notifications (category `MailSentNotification`) are ignored.

### 3. Summary generation

* If `app.Ollama.Enabled` is `true`, the `SummaryGenerator` builds a conversation history, sanitises it for the LLM, and calls `OllamaClient.GenerateSummary`.
* The AI summary is inserted into the work‑item description, wrapped by the `==== AI Generated Summary ==== ` marker.

### 4. Attachment handling

* Inline images (`cid:`) are re‑encoded to JPEG (70 % quality) and embedded as data URLs.
* All attachments are uploaded to TFS via `CreateAttachmentAsync` and linked to the work item.

---

## Contributing

1. Fork the repository.  
2. Create a feature branch (`git checkout -b feature/xyz`).  
3. Run `dotnet test` (if tests are added).  
4. Submit a pull request.  

All contributions must be licensed under MIT. Please keep the code style consistent with the existing files (use `dotnet format`).

---

## License

This project is licensed under the MIT License – see the [LICENSE](LICENSE) file for details.

---

## Acknowledgements

* [.NET 8](https://dotnet.microsoft.com/) – runtime & SDK  
* [Microsoft Graph SDK](https://github.com/microsoftgraph/msgraph-sdk-dotnet) – e‑mail access  
* [Azure DevOps / TFS Client](https://github.com/microsoft/azure-devops-dotnet) – work‑item API  
* [ReverseMarkdown](https://github.com/ronjones/ReverseMarkdown) – Markdown ↔ HTML conversion  
* [HtmlAgilityPack](https://github.com/zzzprojects/html-agility-pack) – HTML sanitisation  
* [Ollama](https://ollama.ai/) – local LLM inference

Happy coding!