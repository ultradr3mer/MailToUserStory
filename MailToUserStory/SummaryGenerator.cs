using System.Text;
using HtmlAgilityPack;

namespace MailToUserStory
{
  internal class SummaryGenerator
  {
    public SummaryGenerator(OllamaClient client, bool enabled)
    {
      this.client = client;
      this.enabled = enabled;
    }

    private const string AI_MARKER = "==== AI Generated Summary ====";
    private readonly OllamaClient client;
    private readonly bool enabled;

    public async Task<string> Summarize(string currentDescription, List<string> history)
    {
      if(!enabled)
      {
        if (currentDescription.Contains(AI_MARKER))
        {
          return currentDescription.Split(AI_MARKER)[0];
        }
        
        return currentDescription;
      }

      var sanitizedHistoryList = history.Select(h => Util.SanitizeHtmlForLlm(h) 
      + Environment.NewLine
      + "===== ENDE EINTRAG =====")
        .Append("===== ENDE DES EMAIL VERLAUFES =====")
        .ToList();
      var historyString = string.Join(Environment.NewLine, sanitizedHistoryList);
      var summary = await client.GenerateSummary(historyString);
      if (currentDescription.Contains(AI_MARKER))
      {
        var split = currentDescription.Split(AI_MARKER);
        var desc = new List<string>
        {
          split[0],
          AI_MARKER,
          summary
        };
        return string.Join("<br>\r\n", desc);
      }
      else
      {
        var desc = new List<string>
        {
          currentDescription,
          AI_MARKER,
          summary
        };
        return string.Join("<br>\r\n", desc);
      }
    }
  }

}