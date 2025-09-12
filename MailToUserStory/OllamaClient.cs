using Microsoft.Graph.Models.Security;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace MailToUserStory
{
  internal class OllamaClient
  {
    private static Regex condenseSpaces = new Regex(@"\s+", RegexOptions.Compiled);

    private readonly string host;
    private readonly HttpClient http;
    private readonly string model;
    private readonly string generateSummaryInstruction;

    public OllamaClient(string host, string model, string generateSummaryInstruction)
    {
      this.http = new HttpClient
      {
        Timeout = TimeSpan.FromSeconds(90)
      };
      this.host = host;
      this.model = model;
      this.generateSummaryInstruction = generateSummaryInstruction;
    }

    internal Task<string> GenerateSummary(string history)
    {
      return this.Complete(history,
                            generateSummaryInstruction);
    }

    private async Task<string> Complete(string context, string instruction)
    {
      // Build prompt
      var sb = new StringBuilder();
      sb.AppendLine(context);
      sb.AppendLine(Minify(instruction));

      var prompt = sb.ToString();

      // Prepare request
      var payload = new
      {
        this.model,
        prompt,
        stream = false
      };

      var json = JsonSerializer.Serialize(payload);
      using var content = new StringContent(json, Encoding.UTF8, "application/json");

      var response = await http.PostAsync($"{host.TrimEnd('/')}/api/generate", content);
      response.EnsureSuccessStatusCode();

      using var stream = await response.Content.ReadAsStreamAsync();
      using var doc = await JsonDocument.ParseAsync(stream);

      // Ollama’s /api/generate usually returns { "response": "...", ... }
      if (doc.RootElement.TryGetProperty("response", out var resp))
      {
        return resp.GetString() ?? string.Empty;
      }

      // fallback: return entire JSON if schema changes
      return doc.RootElement.ToString();
    }

    private static string Minify(string text)
    {
      var result = text.Replace("\r", string.Empty)
                       .Replace("\n", string.Empty)
                       .Replace("\t", string.Empty);

      result = condenseSpaces.Replace(result, " ");

      return result;
    }

  }

}