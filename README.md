# MailToUserStory

> **Author:** Clara  
> **Version:** 1.0.0  
> **Target Framework:** .NET 8.0  
> **License:** MIT (see `LICENSE`)

---

## Table of Contents

| Section |
|---------|
| [What is it?](#what-is-it) |
| [Features](#features) |
| [Architecture](#architecture) |
| [Prerequisites](#prerequisites) |
| [Installation](#installation) |
| [Configuration](#configuration) |
| [Running the App](#running-the-app) |
| [How it Works](#how-it-works) |
| [Database Schema](#database-schema) |
| [Extending / Contributing](#extending--contributing) |
| [Troubleshooting](#troubleshooting) |
| [License](#license) |

---

## What is it?

`MailToUserStory` is a lightweight .NET 8 console application that bridges Microsoft Graph mailboxes and an on‑premises Azure DevOps / TFS instance.  
It watches one or more mailboxes for new e‑mails, extracts a **User Story** ID from the subject line (or creates a new one), and updates the corresponding work item in TFS.  
All processed e‑mails are recorded in a local SQLite database to guarantee idempotency and to keep track of delta tokens.

> **Why?**  
> In many teams, stakeholders send e‑mails that describe new features or bug reports.  
> Manually creating or updating work items is tedious and error‑prone.  
> This tool automates the process, ensuring that every e‑mail becomes a traceable work item with attachments and comments.

---

## Features

| Feature | Description |
|---------|-------------|
| **Delta polling** | Uses Microsoft Graph delta queries to fetch only new or changed messages. |
| **Idempotent processing** | Stores processed message IDs in SQLite; skips duplicates. |
| **Self‑sent e‑mail guard** | Detects and ignores e‑mails sent by the monitored mailbox itself. |
| **User Story ID extraction** | Parses `[US#12345]` tokens from the subject line. |
| **Create / Update** | Creates a new User Story if no ID is found; otherwise updates the existing one. |
| **Attachments** | Downloads all file attachments, uploads them to TFS, and attaches them to the work item. |
| **Reply notifications** | Sends a reply back to the sender with a canonical subject token and a short status message. |
| **Markdown conversion** | Converts the e‑mail body to Markdown (via `ReverseMarkdown`) before posting it as a comment. |
| **SQLite state store** | Persists delta tokens, processed message IDs, and mailbox‑story links. |
| **Configuration via JSON + User Secrets** | Keeps secrets out of source control. |
| **Extensible** | All core logic is split into small, testable classes (`GraphConnector`, `TfsConnector`, `Util`, `Db`). |

---

## Architecture

* **GraphConnector** – thin wrapper around the Microsoft Graph SDK.  
* **TfsConnector** – thin wrapper around the Azure DevOps/TFS REST client.  
* **Util** – helper functions for HTML sanitisation, Markdown conversion, and ID parsing.  
* **Db** – simple SQLite wrapper that keeps the application state.

---

## Prerequisites

| Item | Minimum Version | Notes |
|------|-----------------|-------|
| .NET SDK | 8.0 | Install from <https://dotnet.microsoft.com/download> |
| Azure AD App | – | Must have `Mail.ReadWrite` and `Mail.Send` permissions. |
| TFS / Azure DevOps Server | 2022+ | On‑premises instance with Work Item Tracking enabled. |
| SQLite | – | Included via `Microsoft.Data.Sqlite` NuGet package. |
| Git | – | For cloning the repo. |

---

## Installation

```bash
# Clone the repo
git clone https://github.com/yourorg/mail-to-userstory.git
cd mail-to-userstory

# Restore NuGet packages
dotnet restore
```

---

## Configuration

The application uses a combination of JSON files and **User Secrets** to keep sensitive data out of source control.

| File | Purpose | Example |
|------|---------|---------|
| `appsettings.json` | Default configuration (non‑secret values). | `{"Graph":{"TenantId":"<tenant-id>","ClientId":"<client-id>","Mailboxes":["user@contoso.com"]},"Tfs":{"BaseUrl":"https://tfs.contoso.com","ProjectCollection":"DefaultCollection","Project":"MyProject","User":"tfsuser","Password":"<placeholder>"}}` |
| `appsettings.local.json` | Optional overrides for local dev. | `{"Polling":{"Minutes":2}}` |
| User Secrets | Secrets that must not be committed. | `dotnet user-secrets set "Graph:ClientSecret" "<client-secret>"`<br>`dotnet user-secrets set "Tfs:Password" "<tfs-password>"` |

> **Tip:** The `UserSecretsId` is already defined in the `.csproj`.  
> Run the following commands to set the secrets:

```bash
dotnet user-secrets set "Graph:ClientSecret" "<client-secret>"
dotnet user-secrets set "Tfs:Password" "<tfs-password>"
```

---

## Running the App

```bash
dotnet run
```

The console will output the progress for each mailbox, including:

* Delta page counts
* Message subjects
* Actions taken (created, updated, skipped, error)

All processed messages are stored in `stories.db` (or the path specified in `appsettings.json`).

---

## How it Works (Step‑by‑step)

1. **Startup**  
   * Load configuration.  
   * Initialise SQLite database (`Db.InitializeSchema`).  
   * Create Graph and TFS clients.

2. **Mailbox Loop**  
   For each mailbox in `Graph.Mailboxes`:
   * Retrieve the stored delta link (`db.GetDeltaLink`).  
   * Call `GraphConnector.DeltaPagesAsync` to stream all new/changed messages.  
   * For each message:
     * Skip if already processed (`db.WasProcessed`).  
     * Skip if sent by the mailbox itself (`GraphConnector.IsSelf`).  
     * Parse `[US#12345]` from the subject (`Util.ParseUserStoryId`).  
     * **If ID exists**  
       * Verify the work item exists (`TfsConnector.WorkItemExistsAsync`).  
       * Prepare content (`Util.PrepareContentAsync`).  
       * Add comment + attachments (`TfsConnector.AddCommentAndAttachmentsAsync`).  
       * Send reply (`GraphConnector.SendInfoReplyAsync`).  
       * Mark as processed (`db.MarkProcessed`).  
     * **If no ID**  
       * Prepare content.  
       * Create new User Story (`TfsConnector.CreateUserStoryAsync`).  
       * Upload attachments (`TfsConnector.AddAttachmentsAsync`).  
       * Link mailbox to story (`db.LinkStory`).  
       * Send reply.  
       * Mark as processed.  
   * After finishing a delta page, store the new delta link (`db.UpsertDeltaLink`).

3. **Shutdown**  
   * Dispose of the SQLite connection.  
   * Exit.

---

## Database Schema

The SQLite database contains three tables:

| Table | Columns | Purpose |
|-------|---------|---------|
| `Mailboxes` | `mailbox TEXT PRIMARY KEY`, `delta_link TEXT` | Stores the last delta link per mailbox. |
| `Stories` | `mailbox TEXT`, `work_item_id INTEGER`, `PRIMARY KEY(mailbox, work_item_id)` | Keeps a record of which work items belong to which mailbox. |
| `ProcessedEmails` | `graph_message_id TEXT PRIMARY KEY`, `mailbox TEXT`, `work_item_id INTEGER`, `processed_at DATETIME`, `outcome TEXT` | Records every processed message to avoid duplicates. |

The schema is created automatically on first run by `Db.InitializeSchema`.

---

## Extending / Contributing

Feel free to fork, branch, and submit pull requests.  
When adding new features, keep the following in mind:

1. **Keep the core logic testable** – add unit tests for new methods.  
2. **Avoid hard‑coding** – use configuration or dependency injection.  
3. **Document changes** – update the README and add comments where necessary.  

---

## Troubleshooting

| Symptom | Likely Cause | Fix |
|---------|--------------|-----|
| `ClientSecret` or `Password` missing | User secrets not set | Run the `dotnet user-secrets set` commands again. |
| Graph API returns 401 | Wrong tenant/client ID or insufficient permissions | Verify Azure AD app registration and grant `Mail.ReadWrite` & `Mail.Send`. |
| TFS returns 404 on work item | Wrong project name or missing permissions | Check `appsettings.json` for correct `Project` and that the user has *Edit work items* rights. |
| SQLite file not created | Wrong path or permissions | Ensure the directory exists and the process has write access. |
| Attachments not uploaded | Attachment size > 4 MB (Graph limit) | Split large attachments or use Graph's upload session. |

---

## License

MIT – see the `LICENSE` file for details.

---

**Happy coding!**  
If you run into any issues or have feature requests, open an issue or reach out to Clara.