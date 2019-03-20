using HtmlAgilityPack;
using Microsoft.TeamFoundation.WorkItemTracking.Client;
using Microsoft.TeamFoundation.WorkItemTracking.Proxy;
using Serilog;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using WordExporter.Core.WordManipulation.Support;

namespace WordExporter.Core.WorkItems
{
    public static class Extensions
    {
        /// <summary>
        /// Simply create a dictionary of substitution values from all the fields
        /// of the work item.
        /// </summary>
        /// <param name="workItem"></param>
        /// <returns></returns>
        public static Dictionary<String, Object> CreateDictionaryFromWorkItem(this WorkItem workItem)
        {
            var retValue = new Dictionary<String, Object>(StringComparer.OrdinalIgnoreCase);
            retValue["title"] = workItem.Title;
            retValue["id"] = workItem.Id;
            //Some help to avoid forcing the user to use System.AssignedTo etc for most commonly used fields.
            retValue["description"] = new HtmlSubstitution(workItem.EmbedHtmlContent(workItem.Description));
            retValue["assignedto"] = workItem.Fields["System.AssignedTo"].Value?.ToString() ?? String.Empty;
            retValue["createdby"] = workItem.Fields["System.CreatedBy"].Value?.ToString() ?? String.Empty;

            HashSet<String> specialFields = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
            {
                "description",
            };

            //All the fields will be set in raw format.
            foreach (Field field in workItem.Fields)
            {
                if (specialFields.Contains(field.Name))
                    continue; //This is a special field, ignore.

                retValue[field.Name] = retValue[field.ReferenceName] = GetValue(field);
            }

            //ok some of the historical value could be of interests, as an example the last user timestamp for each state change
            //is an information that can be interesting
            if (workItem.Revisions.Count > 0)
            {
                foreach (Revision revision in workItem.Revisions)
                {
                    var fieldsChanged = revision
                        .Fields
                        .OfType<Field>()
                        .Where(f => f.IsChangedInRevision)
                        .ToList();
                    var changedBy = revision.Fields["Changed By"].Value;
                    var changedDate = revision.Fields["Changed Date"].Value;
                    foreach (var field in fieldsChanged)
                    {
                        if (field.ReferenceName.Equals("system.state", StringComparison.OrdinalIgnoreCase))
                        {
                            retValue[$"statechange.{field.Value.ToString().ToLower()}.author"] = changedBy;
                            retValue[$"statechange.{field.Value.ToString().ToLower()}.date"] = ((DateTime)changedDate).ToShortDateString();
                        }
                        else if (field.ReferenceName.Equals("system.areapath", StringComparison.OrdinalIgnoreCase))
                        {
                            retValue[$"lastareapathchange.author"] = changedBy;
                            retValue[$"lastareapathchange.date"] = ((DateTime)changedDate).ToShortDateString();
                        }
                    }
                }
            }
            return retValue;
        }


        public static String GetValue(this Field field)
        {
            if (field.Value == null)
                return String.Empty;

            if (field.Value is DateTime dateTime)
                return dateTime.ToShortDateString();

            return field.Value.ToString();
        }

        public static String EmbedHtmlContent(this WorkItem workItem, String htmlContent)
        {
            HtmlDocument doc = new HtmlDocument();
            doc.LoadHtml(htmlContent);

            var images = doc.DocumentNode.SelectNodes("//img");
            if (images != null)
            {
                foreach (var image in images)
                {
                    //need to understand if it is in base 64 or no, if the answer is no, we need to embed image
                    var src = image.GetAttributeValue("src", "");
                    if (!String.IsNullOrEmpty(src))
                    {
                        if (src.Contains("base64")) // data:image/jpeg;base64,
                        {
                            //image already embedded
                            Log.Debug("found image in html content that was already in base64");
                        }
                        else
                        {
                            Log.Debug("found image in html content that point to external image {src}", src);
                            //is it a internal attached images?
                            var match = Regex.Match(src, @"FileID=(?<id>\d*)");
                            if (match.Success)
                            {
                                var attachment = workItem.Attachments
                                    .OfType<Attachment>()
                                    .FirstOrDefault(_ => _.Id.ToString() == match.Groups["id"].Value);
                                if (attachment != null)
                                {
                                    //ok we can embed in the image as base64
                                    WorkItemServer wise = workItem.Store.TeamProjectCollection.GetService<WorkItemServer>();
                                    var downloadedAttachment = wise.DownloadFile(attachment.Id);
                                    byte[] byteContent = File.ReadAllBytes(downloadedAttachment);
                                    String base64Encoded = Convert.ToBase64String(byteContent);
                                    var newSrcValue = $"data:image/{attachment.Extension.Trim('.')};base64,{base64Encoded}";
                                    image.SetAttributeValue("src", newSrcValue);
                                }
                            }
                        }
                    }
                }
            }

            return doc.DocumentNode.OuterHtml;
        }
    }
}
