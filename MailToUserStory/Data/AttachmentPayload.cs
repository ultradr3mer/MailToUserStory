using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MailToUserStory.Data
{
  public sealed class AttachmentPayload
  {
    public required string FileName { get; init; }
    public required byte[] Bytes { get; init; }
  }
}
