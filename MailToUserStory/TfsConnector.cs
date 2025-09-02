using MailToUserStory.Data;
using Microsoft.TeamFoundation.WorkItemTracking.WebApi;
using Microsoft.TeamFoundation.WorkItemTracking.WebApi.Models;
using Microsoft.VisualStudio.Services.WebApi.Patch.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using Microsoft.VisualStudio.Services.WebApi.Patch;

namespace MailToUserStory
{
  public static class TfsConnector
  {
    public static async Task<bool> WorkItemExistsAsync(WorkItemTrackingHttpClient wit, int id)
    {
      try { _ = await wit.GetWorkItemAsync(id); return true; }
      //catch (Microsoft.VisualStudio.Services.WebApi.VssServiceException ex) when (ex.Message.Contains("404")) { return false; }
      catch (Microsoft.VisualStudio.Services.WebApi.VssServiceResponseException ex) when (ex.HttpStatusCode == HttpStatusCode.NotFound) { return false; }
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

    public static async Task AddCommentAndAttachmentsAsync(WorkItemTrackingHttpClient wit, string project, int id, string commentMarkdown, List<AttachmentPayload> attachments)
    {
      var patch = new JsonPatchDocument
    {
        new JsonPatchOperation { Operation = Operation.Add, Path = "/fields/System.History", Value = commentMarkdown }
    };

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
