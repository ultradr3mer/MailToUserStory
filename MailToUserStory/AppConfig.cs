namespace MailToUserStory
{
  public record AppConfig
  {
    public required GraphConfig Graph { get; init; }
    public required TfsConfig Tfs { get; init; }
    public string DatabasePath { get; init; } = "stories.db";
    public required OllamaConfig Ollama { get; init; }
    public required bool Pause { get; init; }
    public required List<GraphTfsLink> Links { get; init; }

    public record GraphConfig
    {
      public required string TenantId { get; init; }
      public required string ClientId { get; init; }
      public string? ClientSecret { get; init; }
      public required string UsCreatedTemplate { get; init; }
      public required string UsUpdatedTemplate { get; init; }
      public required string UsNotFoundTemplate { get; init; }
      public required DateTime BeginDate { get; init; }
      public required List<string> Users { get; init; }
    }

    public record GraphTfsLink
    {
      public required string Mailbox { get; init; }
      public required string Project { get; init; }
    }

    public record TfsConfig
    {
      public required string BaseUrl { get; init; }
      public required string ProjectCollection { get; init; }
      public string? User { get; init; }
      public string Password { get; init; }
    }

    public record OllamaConfig
    {
      public string Host { get; init; } = "http://localhost:11434";
      public string Model { get; init; } = "gpt-oss:20b";
      public string SummarizeInstruction { get; init; } = string.Empty;
      public bool Enabled { get; init; } = false;
    }
  }
}
