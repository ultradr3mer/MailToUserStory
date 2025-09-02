using Microsoft.Graph.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MailToUserStory.Data
{
  public class AttachmentContainer
  {
    public required List<FileAttachment> InlineAttachments { get; set; }
    public required List<FileAttachment> FileAttachments { get; set; }
  }
}
