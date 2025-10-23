using MailToUserStory.Data;
using Microsoft.TeamFoundation.WorkItemTracking.WebApi;
using Microsoft.TeamFoundation.WorkItemTracking.WebApi.Models;
using Microsoft.VisualStudio.Services.Common;
using Microsoft.VisualStudio.Services.WebApi;
using Microsoft.VisualStudio.Services.WebApi.Patch;
using Microsoft.VisualStudio.Services.WebApi.Patch.Json;
using System.Net;

namespace MailToUserStory
{
  public static class TfsConnector
  {
    private static string TfsUsNotFoundErrorCode = "TF401232";

    public static async Task<string?> WorkItemExistsingDescriptionAsync(
        WorkItemTrackingHttpClient wit,
        int id)
    {
      try
      {
        var wi = await wit.GetWorkItemAsync(id, expand: WorkItemExpand.Fields);
        wi.Fields.TryGetValue("System.Description", out var desc);
        return desc as string ?? string.Empty;
      }
      catch (VssServiceException ex) 
        when (ex.Message.StartsWith(TfsUsNotFoundErrorCode))
      {
        return null;
      }
    }

    public static async Task<int> CreateUserStoryAsync(WorkItemTrackingHttpClient wit, string project, string title, string descriptionMarkdown)
    {
      var patch = new JsonPatchDocument
      {
        new JsonPatchOperation { Operation = Operation.Add, Path = "/fields/System.Title", Value = title },
        new JsonPatchOperation { Operation = Operation.Add, Path = "/fields/System.Description", Value = descriptionMarkdown }
      };
      var wi = await wit.CreateWorkItemAsync(patch, project, "User Story");
      return wi.Id ?? throw new Exception("No ID returned from CreateWorkItemAsync");
    }

    public static async Task AddCommentAndAttachmentsAsync(
        WorkItemTrackingHttpClient wit,
        int id,
        string commentMarkdown,
        List<AttachmentPayload> attachments,
        string? updatedDescription = null)
    {
      var patch = new JsonPatchDocument
    {
        new JsonPatchOperation
        {
            Operation = Operation.Add,
            Path = "/fields/System.History",
            Value = commentMarkdown
        }
    };

      if (!string.IsNullOrEmpty(updatedDescription))
      {
        patch.Add(new JsonPatchOperation
        {
          Operation = Operation.Replace, 
          Path = "/fields/System.Description",
          Value = updatedDescription
        });
      }

      foreach (var a in attachments)
      {
        using var ms = new MemoryStream(a.Bytes);
        var ar = await wit.CreateAttachmentAsync(ms, fileName: a.FileName);
        patch.Add(new JsonPatchOperation
        {
          Operation = Operation.Add,
          Path = "/relations/-",
          Value = new WorkItemRelation { Rel = "AttachedFile", Url = ar.Url }
        });
      }

      _ = await wit.UpdateWorkItemAsync(patch, id);
    }

    public static async Task AddAttachmentsAsync(WorkItemTrackingHttpClient wit, string project, int id, List<AttachmentPayload> attachments)
    {
      if (attachments.Count == 0) return;
      var patch = new JsonPatchDocument();
      foreach (var a in attachments)
      {
        using var ms = new MemoryStream(a.Bytes);
        var ar = await wit.CreateAttachmentAsync(ms, fileName: a.FileName);
        patch.Add(new JsonPatchOperation
        {
          Operation = Operation.Add,
          Path = "/relations/-",
          Value = new WorkItemRelation { Rel = "AttachedFile", Url = ar.Url }
        });
      }
      _ = await wit.UpdateWorkItemAsync(patch, id);
    }
  }
}
