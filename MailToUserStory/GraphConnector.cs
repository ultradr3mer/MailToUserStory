
using MailToUserStory.Data;
using Microsoft.Graph;
using Microsoft.Graph.Chats.Item.Messages.Delta;
using Microsoft.Graph.Models;
using Microsoft.Kiota.Abstractions;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DeltaGetResponse = Microsoft.Graph.Users.Item.MailFolders.Item.Messages.Delta.DeltaGetResponse;

namespace MailToUserStory
{
  public static class GraphConnector
  {
    public static async IAsyncEnumerable<DeltaPage> DeltaPagesAsync(GraphServiceClient graph, string mailbox, string? deltaLink)
    {
      Microsoft.Graph.Users.Item.MailFolders.Item.Messages.Delta.DeltaGetResponse? page;
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
        page = await graph.Users[mailbox]
          .MailFolders["Inbox"].Messages.Delta
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
          page = await graph.RequestAdapter.SendAsync<Microsoft.Graph.Users.Item.MailFolders.Item.Messages.Delta.DeltaGetResponse>(new Microsoft.Kiota.Abstractions.RequestInformation
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

    public static bool IsSelf(string mailbox, Message msg)
        => string.Equals(msg.From?.EmailAddress?.Address, mailbox, StringComparison.OrdinalIgnoreCase);

    public static async Task<AttachmentContainer> GetFileAttachmentsAsync(GraphServiceClient graph, string mailbox, string messageId)
    {
      var fileAttachments = new List<FileAttachment>();
      var inlineAttachments = new List<FileAttachment>();
      var page = await graph.Users[mailbox].Messages[messageId].Attachments.GetAsync();
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

    public static Task SendErrorReplyAsync(GraphServiceClient graph, string mailbox, Message original, string errorText)
          => SendInfoReplyAsync(graph, mailbox, original, errorText, null);
  }
}
