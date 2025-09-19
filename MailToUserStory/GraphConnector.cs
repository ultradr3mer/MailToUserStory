
using MailToUserStory.Data;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using DeltaGetResponse = Microsoft.Graph.Users.Item.MailFolders.Item.Messages.Delta.DeltaGetResponse;

namespace MailToUserStory
{
  public class GraphConnector
  {
    private const string InboxWellKnownFolderName = "Inbox";
    private const string SentItemsWellKnownFolderName = "SentItems";
    private const string MailToTfsNotificationCategoryName = "MailToTfs-Notification";
    private const string MailToTfsDoneCategoryName = "MailToTfs-Done";
    private readonly GraphServiceClient client;

    public GraphConnector(GraphServiceClient client)
    {
      this.client = client;
    }

    public async IAsyncEnumerable<DeltaPage> DeltaPagesAsync(string mailbox, string? deltaLink, DateTime beginDate)
    {
      var userFolder = new GraphUserFolder(mailbox);

      var folderId = await ResolveFolderIdAsync(userFolder.User, userFolder.Folder);

      DeltaGetResponse? page;
      if (!string.IsNullOrEmpty(deltaLink))
      {
        // Resume from stored delta link
        page = await this.client.RequestAdapter.SendAsync<DeltaGetResponse>(new Microsoft.Kiota.Abstractions.RequestInformation
        {
          HttpMethod = Microsoft.Kiota.Abstractions.Method.GET,
          UrlTemplate = deltaLink,
        },
          DeltaGetResponse.CreateFromDiscriminatorValue);
      }
      else
      {
        page = await this.client.Users[userFolder.User]
          .MailFolders[folderId].Messages.Delta
          .GetAsDeltaGetResponseAsync(r =>
          {
            r.QueryParameters.Select = new[]
            {
              "id", "subject", "from", "toRecipients",
              "receivedDateTime", "hasAttachments", "body", "categories"
            };

            r.QueryParameters.Filter = "receivedDateTime ge " + beginDate.ToUniversalTime().ToString("yyyy-MM-ddTHH:mm:ssZ");
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
          page = await this.client.RequestAdapter.SendAsync<DeltaGetResponse>(new Microsoft.Kiota.Abstractions.RequestInformation
          {
            HttpMethod = Microsoft.Kiota.Abstractions.Method.GET,
            UrlTemplate = page.OdataNextLink,
          },
          DeltaGetResponse.CreateFromDiscriminatorValue);
        }
        else
        {
          break;
        }
      }
    }

    private async Task<string?> ResolveFolderIdAsync(string user, string folderName)
    {
      if (folderName == InboxWellKnownFolderName || folderName == SentItemsWellKnownFolderName)
        return folderName;

      // Initial request from msgfolderroot
      var response = await this.client.Users[user]
                                .MailFolders["msgfolderroot"]
                                .ChildFolders
                                .GetAsync();

      while (response != null)
      {
        foreach (var f in response.Value)
        {
          if (string.Equals(f.DisplayName, folderName, StringComparison.OrdinalIgnoreCase))
            return f.Id;
        }

        if (response.OdataNextLink != null)
        {
          // Create new request using the NextLink
          var nextPage = new Microsoft.Graph.Users.Item.MailFolders.Item.ChildFolders.ChildFoldersRequestBuilder(
              response.OdataNextLink,
              this.client.RequestAdapter);

          response = await nextPage.GetAsync();
        }
        else
        {
          response = null;
        }
      }

      return null;
    }

    public static bool IsSelf(string mailbox, Message msg)
      => string.Equals(msg.From?.EmailAddress?.Address, new GraphUserFolder(mailbox).User, StringComparison.OrdinalIgnoreCase);

    public async Task<AttachmentContainer> GetFileAttachmentsAsync(string mailbox, string messageId)
    {
      var fileAttachments = new List<FileAttachment>();
      var inlineAttachments = new List<FileAttachment>();
      var page = await client.Users[new GraphUserFolder(mailbox).User].Messages[messageId].Attachments.GetAsync();
      foreach (var att in page?.Value ?? Enumerable.Empty<Attachment>())
      {
        if (!(att is FileAttachment fa) || fa.ContentBytes == null)
          continue;

        if (att.IsInline == true)
          inlineAttachments.Add(fa);
        else
          fileAttachments.Add(fa);
      }

      return new AttachmentContainer() { FileAttachments = fileAttachments, InlineAttachments = inlineAttachments };
    }

    public async Task SendInfoReplyAsync(string mailbox, Message original, string infoBody, string? subjectSuffix = null)
    {
      string subject = original.Subject ?? string.Empty;
      if (!string.IsNullOrEmpty(subjectSuffix)) subject = subject + subjectSuffix;

      var msg = new Message
      {
        ToRecipients = [original.From],
        Subject = subject,
        Body = new ItemBody { ContentType = BodyType.Html, Content = infoBody },
        Categories = new List<string> { MailToTfsNotificationCategoryName }
      };

      var sendRequest = new Microsoft.Graph.Users.Item.SendMail.SendMailPostRequestBody()
      { Message = msg, SaveToSentItems = true };

      await this.client.Users[new GraphUserFolder(mailbox).User].SendMail.PostAsync(sendRequest);
    }

    public Task SendErrorReplyAsync(string mailbox, Message original, string errorBody)
          => this.SendInfoReplyAsync(mailbox, original, errorBody, null);

    internal static string GetSentMailbox(string mailbox)
    {
      return new GraphUserFolder(mailbox).User + "/" + SentItemsWellKnownFolderName;
    }

    internal static bool HasNotificationCategory(Message msg)
    {
      return msg.Categories?.Any(c => string.Equals(c, MailToTfsNotificationCategoryName, StringComparison.OrdinalIgnoreCase)) == true;
    }

    internal async Task CategorizeDoneAsync(Message msg, string mailbox)
    {
      if (msg == null || string.IsNullOrEmpty(msg.Id))
        throw new ArgumentException("Message must not be null and must have an Id", nameof(msg));

      // Build updated categories list
      var categories = (msg.Categories ?? new List<string>()).ToList();
      if (categories.Contains(MailToTfsDoneCategoryName))
        return;

      categories.Add(MailToTfsDoneCategoryName);

      var update = new Message
      {
        Categories = categories
      };

      await client.Users[new GraphUserFolder(mailbox).User]
          .Messages[msg.Id]
          .PatchAsync(update);
    }

    internal static bool WasProcessed(Message msg)
    {
      return msg.Categories?.Any(c => string.Equals(c, MailToTfsDoneCategoryName, StringComparison.OrdinalIgnoreCase)) == true;
    }

    private class GraphUserFolder
    {
      public GraphUserFolder(string mailbox)
      {
        if (mailbox.Contains("/"))
        {
          var split = mailbox.Split('/');
          this.User = split[0];
          this.Folder = split[1];
        }
        else
        {
          this.User = mailbox;
          this.Folder = InboxWellKnownFolderName;
        }
      }

      public string User { get; }
      public string Folder { get; }
    }
  }
}
