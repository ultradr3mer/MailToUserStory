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
      string markdown = HtmlToMarkdown(msg.Body, converter);
      var attachments = new List<AttachmentPayload>();
      if (msg.HasAttachments == true)
      {
        var files = await GraphConnector.GetFileAttachmentsAsync(graph, mailbox, msg.Id!);
        foreach (var fa in files)
        {
          attachments.Add(new AttachmentPayload
          {
            FileName = fa.Name!,
            Bytes = fa.ContentBytes!
          });
        }
      }

      var meta = new StringBuilder();
      meta.AppendLine();
      meta.AppendLine("---");
      meta.AppendLine("> From: " + msg.From?.EmailAddress?.Name + " <" + msg.From?.EmailAddress?.Address + ">");
      meta.AppendLine("> Received: " + (msg.ReceivedDateTime.HasValue ? msg.ReceivedDateTime.Value.ToString("O") : ""));

      return (markdown + "\n\n" + meta.ToString(), attachments);
    }

    public static string HtmlToMarkdown(ItemBody? body, ReverseMarkdown.Converter converter)
    {
      if (body == null) return "(no content)";
      if (body.ContentType == BodyType.Text) return string.IsNullOrWhiteSpace(body.Content) ? "(no content)" : body.Content!.Trim();

      var html = body.Content ?? string.Empty;
      var doc = new HtmlDocument();
      doc.LoadHtml(html);
      foreach (var n in doc.DocumentNode.SelectNodes("//script|//style") ?? new HtmlNodeCollection(doc.DocumentNode)) n.Remove();
      string sanitized = doc.DocumentNode.InnerHtml;
      string md = converter.Convert(sanitized);
      return md.Trim();
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
