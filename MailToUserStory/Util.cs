using HtmlAgilityPack;
using MailToUserStory.Data;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using System.Text;
using System.Text.RegularExpressions;

namespace MailToUserStory
{
  public static class Util
  {
    public static async Task<(string markdown, List<AttachmentPayload> attachments)> PrepareContentAsync(GraphServiceClient graph, string mailbox, Message msg, ReverseMarkdown.Converter converter)
    {
      var attachments = new List<AttachmentPayload>();

      var container = await GraphConnector.GetFileAttachmentsAsync(graph, mailbox, msg.Id!);
      foreach (var fa in container.FileAttachments)
      {
        attachments.Add(new AttachmentPayload
        {
          FileName = fa.Name!,
          Bytes = fa.ContentBytes!
        });
      }

      string html = SanitizeHtml(msg.Body, converter, container.InlineAttachments);

      var meta = new StringBuilder();
      meta.AppendLine();
      meta.AppendLine("---");
      meta.AppendLine("> From: " + msg.From?.EmailAddress?.Name + " <" + msg.From?.EmailAddress?.Address + ">");
      meta.AppendLine("> Received: " + (msg.ReceivedDateTime.HasValue ? msg.ReceivedDateTime.Value.ToString("O") : ""));

      return (html + "\n\n" + meta.ToString(), attachments);
    }

    public static string SanitizeHtml(ItemBody? body, ReverseMarkdown.Converter converter, List<FileAttachment> inlineAttachments)
    {
      if (body == null) return "(no content)";
      if (body.ContentType == BodyType.Text) return string.IsNullOrWhiteSpace(body.Content) ? "(no content)" : body.Content!.Trim();

      var html = body.Content ?? string.Empty;

      foreach (var att in inlineAttachments)
      {
        if (att.IsInline == true && !string.IsNullOrEmpty(att.ContentId) && att.ContentBytes != null)
        {
          var base64 = Convert.ToBase64String(att.ContentBytes);
          var mime = att.ContentType ?? "image/png"; // guess if null
          html = html.Replace(
              $"cid:{att.ContentId}",
              $"data:{mime};base64,{base64}"
          );
        }
      }

      var doc = new HtmlDocument();
      doc.LoadHtml(html);
      foreach (var n in doc.DocumentNode.SelectNodes("//script|//style") ?? new HtmlNodeCollection(doc.DocumentNode)) n.Remove();
      string sanitized = doc.DocumentNode.InnerHtml;
      return "<div>" + sanitized + "</div>";
    }

    public static int? ParseUserStoryId(string? subject)
    {
      if (string.IsNullOrEmpty(subject)) return null;
      var rx = new Regex(@"\[US#(?<id>\d+)\]", RegexOptions.IgnoreCase);

      var m = rx.Match(subject);
      if (m.Success && int.TryParse(m.Groups["id"].Value, out var id)) return id;
      return null;
    }

    public static void Assert(bool condition, string message)
    {
      if (!condition) throw new Exception(message);
    }
  }
}
