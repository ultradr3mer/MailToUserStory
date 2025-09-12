
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
      if (folderName == InboxWellKnownFolderName)
        return InboxWellKnownFolderName;

      var folders = await graph.Users[user].MailFolders.GetAsync();
      var folder = folders?.Value?.FirstOrDefault(f =>
          string.Equals(f.DisplayName, folderName, StringComparison.OrdinalIgnoreCase));

      return folder?.Id;
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
      // Create a draft reply to preserve threading, then patch subject/body, then send
      var draft = await graph.Users[new GraphUserFolder(mailbox).User].Messages[original.Id!].CreateReply.PostAsync(new Microsoft.Graph.Users.Item.Messages.Item.CreateReply.CreateReplyPostRequestBody
      {
        Message = new Message
        {
          Body = new ItemBody { ContentType = BodyType.Text, Content = infoBody }
        }
      });

      if (draft == null) throw new Exception("Failed to create reply draft");

      string subject = original.Subject ?? string.Empty;
      if (!string.IsNullOrEmpty(subjectSuffix)) subject = subject + subjectSuffix;

      await graph.Users[new GraphUserFolder(mailbox).User].Messages[draft.Id!].PatchAsync(new Message
      {
        Subject = subject,
        Body = new ItemBody { ContentType = BodyType.Text, Content = infoBody }
      });

      await graph.Users[new GraphUserFolder(mailbox).User].Messages[draft.Id!].Send.PostAsync();
    }

    public static Task SendErrorReplyAsync(GraphServiceClient graph, string mailbox, Message original, string errorText)
          => SendInfoReplyAsync(graph, mailbox, original, errorText, null);

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
