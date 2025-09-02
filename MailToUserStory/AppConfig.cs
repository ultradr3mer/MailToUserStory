using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MailToUserStory
{
  public record AppConfig
  {
    public required string Project { get; init; }
    public required GraphConfig Graph { get; init; }
    public required TfsConfig Tfs { get; init; }
    public PollingConfig Polling { get; init; } = new();
    public string DatabasePath { get; init; } = "stories.db";

    public record GraphConfig
    {
      public required string TenantId { get; init; }
      public required string ClientId { get; init; }
      public string? ClientSecret { get; init; }
      public required string[] Mailboxes { get; init; }
    }

    public record TfsConfig
    {
      public required string BaseUrl { get; init; }
      public required string ProjectCollection { get; init; }
      public required string Project { get; init; }
      public string? Pat { get; init; }
    }

    public record PollingConfig { public int Minutes { get; init; } = 5; }
  }
}
