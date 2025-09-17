
using MailToUserStory.Data;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using System;
using DeltaGetResponse = Microsoft.Graph.Users.Item.MailFolders.Item.Messages.Delta.DeltaGetResponse;

namespace MailToUserStory
{
  public static class GraphConnector
  {
    private const string InboxWellKnownFolderName = "Inbox";
    private const string SentItemsWellKnownFolderName = "SentItems";
    private const string MailToTfsNotificationCategoryName = "MailToTfs-Notification";

    public static async IAsyncEnumerable<DeltaPage> DeltaPagesAsync(GraphServiceClient graph, string mailbox, string? deltaLink)
    {
      var userFolder = new GraphUserFolder(mailbox);

      var folderId = await ResolveFolderIdAsync(graph, userFolder.User, userFolder.Folder);

      DeltaGetResponse? page;
      if (!string.IsNullOrEmpty(deltaLink))
      {
        // Resume from stored delta link
        page = await graph.RequestAdapter.SendAsync<DeltaGetResponse>(new Microsoft.Kiota.Abstractions.RequestInformation
        {
          HttpMethod = Microsoft.Kiota.Abstractions.Method.GET,
          UrlTemplate = deltaLink,
        },
          DeltaGetResponse.CreateFromDiscriminatorValue);
      }
      else
      {
        page = await graph.Users[userFolder.User]
          .MailFolders[folderId].Messages.Delta
          .GetAsDeltaGetResponseAsync(r =>
          {
            r.QueryParameters.Select = new[]
            {
              "id", "subject", "from", "toRecipients",
              "receivedDateTime", "hasAttachments", "body", "categories"
            };
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
          page = await graph.RequestAdapter.SendAsync<DeltaGetResponse>(new Microsoft.Kiota.Abstractions.RequestInformation
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

    private static async Task<string?> ResolveFolderIdAsync(GraphServiceClient graph, string user, string folderName)
    {
      if (folderName == InboxWellKnownFolderName || folderName == SentItemsWellKnownFolderName)
        return folderName;

      // Initial request from msgfolderroot
      var response = await graph.Users[user]
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
              graph.RequestAdapter);

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

    public static async Task<AttachmentContainer> GetFileAttachmentsAsync(GraphServiceClient graph, string mailbox, string messageId)
    {
      var fileAttachments = new List<FileAttachment>();
      var inlineAttachments = new List<FileAttachment>();
      var page = await graph.Users[new GraphUserFolder(mailbox).User].Messages[messageId].Attachments.GetAsync();
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

    public static async Task SendInfoReplyAsync(GraphServiceClient graph, string mailbox, Message original, string infoBody, string? subjectSuffix)
    {
      string subject = original.Subject ?? string.Empty;
      if (!string.IsNullOrEmpty(subjectSuffix)) subject = subject + subjectSuffix;

      var msg = new Message
      {
        ToRecipients = [original.From],
        Subject = subject,
        Body = new ItemBody { ContentType = BodyType.Text, Content = infoBody },
        Categories = new List<string> { MailToTfsNotificationCategoryName }
      };

      var sendRequest = new Microsoft.Graph.Users.Item.SendMail.SendMailPostRequestBody()
      { Message = msg, SaveToSentItems = true };

      await graph.Users[new GraphUserFolder(mailbox).User].SendMail.PostAsync(sendRequest);
    }

    public static Task SendErrorReplyAsync(GraphServiceClient graph, string mailbox, Message original, string errorText)
          => SendInfoReplyAsync(graph, mailbox, original, errorText, null);

    internal static string GetSentMailbox(string mailbox)
    {
      return new GraphUserFolder(mailbox).User + "/" + SentItemsWellKnownFolderName;
    }

    internal static bool HasNotificationCategory(Message msg)
    {
      return msg.Categories?.Any(c => string.Equals(c, MailToTfsNotificationCategoryName, StringComparison.OrdinalIgnoreCase)) == true;
    }

    private class GraphUserFolder
    {
      public GraphUserFolder(string mailbox)
      {
        if (mailbox.Contains("/"))
        {
          var split = mailbox.Split('/');
          User = split[0];
          Folder = split[1];
        }
        else
        {
          User = mailbox;
          Folder = InboxWellKnownFolderName;
        }
      }

      public string User { get; }
      public string Folder { get; }
    }
  }
}
