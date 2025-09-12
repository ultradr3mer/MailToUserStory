using HtmlAgilityPack;
using MailToUserStory.Data;
using Microsoft.Graph;
using Microsoft.Graph.Models;
using System.Net;
using System.Text.RegularExpressions;

namespace MailToUserStory
{
  public static class Util
  {
    public static async Task<(string html, List<AttachmentPayload> attachments)> PrepareContentAsync(GraphServiceClient graph, string mailbox, Message msg, ReverseMarkdown.Converter converter)
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

      string message = SanitizeHtmlForHistory(msg.Body, converter, container.InlineAttachments);

      var html = new List<string>();
      html.Add("Von: " + msg.From?.EmailAddress?.Name + " <" + msg.From?.EmailAddress?.Address + ">");

      if(msg.ToRecipients != null)
        html.Add("An: " + string.Join(";", msg.ToRecipients.Select(r => r.EmailAddress?.Name + " <" + r.EmailAddress?.Address + ">")));

      if (msg.Subject != null)
        html.Add("Betreff: " + msg.Subject);

      html.Add("Gesendet: " + (msg.ReceivedDateTime.HasValue ? msg.ReceivedDateTime.Value.ToString("O") : "")
            +message);


      return (string.Join("<br>\r\n", html), attachments);
    }

    public static string SanitizeHtmlForHistory(ItemBody? body, ReverseMarkdown.Converter converter, List<FileAttachment> inlineAttachments)
    {
      if (body == null) return "(no content)";
      if (body.ContentType == BodyType.Text) return string.IsNullOrWhiteSpace(body.Content) ? "(no content)" : body.Content!.Trim();

      var html = body.Content ?? string.Empty;

      foreach (var att in inlineAttachments)
      {
        if (att.IsInline == true && !string.IsNullOrEmpty(att.ContentId) && att.ContentBytes != null)
        {
          byte[] compressed;
          try
          {
            compressed = ReEncode(att.ContentBytes!);
          }
          catch
          {
            continue;
          }

          // Convert compressed JPEG to base64
          var base64 = Convert.ToBase64String(compressed);

          // Always set MIME type to jpeg since we forced JPEG re-encoding
          html = html.Replace(
              $"cid:{att.ContentId}",
              $"data:image/jpeg;base64,{base64}"
          );
        }

      }

      var doc = new HtmlDocument();
      doc.LoadHtml(html);
      foreach (var n in doc.DocumentNode.SelectNodes("//script|//style") ?? new HtmlNodeCollection(doc.DocumentNode)) n.Remove();
      string sanitized = doc.DocumentNode.InnerHtml;
      return sanitized;
    }

    public static string SanitizeHtmlForLlm(string html)
    {
      HtmlDocument doc = new HtmlDocument();
      doc.LoadHtml(html);

      // Ersetze <br>, <hr> explizit durch Zeilenumbrüche
      foreach (var br in doc.DocumentNode.SelectNodes("//br|//hr") ?? Enumerable.Empty<HtmlNode>())
        br.ParentNode.ReplaceChild(doc.CreateTextNode("\n"), br);

      // Plaintext extrahieren
      string plain = WebUtility.HtmlDecode(doc.DocumentNode.InnerText);

      // Aufräumen: mehrfach \r\n oder \n zu einfachem Zeilenumbruch, trimmen
      plain = Regex.Replace(plain, @"(\r?\n)\s*(\r?\n)", "\r\n").Trim();

      return plain;
    }

    private static byte[] ReEncode(byte[] bytes)
    {
      // Default to JPEG re-encoding for inline images
      byte[] compressed;
      using (var input = new MemoryStream(bytes))
      using (var img = System.Drawing.Image.FromStream(input))
      using (var ms = new MemoryStream())
      {
        // Get JPEG encoder
        var codec = System.Drawing.Imaging.ImageCodecInfo.GetImageDecoders()
            .First(c => c.FormatID == System.Drawing.Imaging.ImageFormat.Jpeg.Guid);

        // Set compression quality (0–100, lower = more compression)
        var encParams = new System.Drawing.Imaging.EncoderParameters(1);
        encParams.Param[0] = new System.Drawing.Imaging.EncoderParameter(
            System.Drawing.Imaging.Encoder.Quality, 70L); // e.g. 70% quality

        img.Save(ms, codec, encParams);
        compressed = ms.ToArray();
      }

      return compressed;
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
