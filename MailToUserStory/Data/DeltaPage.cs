using Microsoft.Graph.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MailToUserStory.Data
{
  public sealed class DeltaPage
  {
    public required List<Message> Messages { get; init; }
    public string? NextLink { get; init; }
    public string? DeltaLink { get; init; }
  }
}
