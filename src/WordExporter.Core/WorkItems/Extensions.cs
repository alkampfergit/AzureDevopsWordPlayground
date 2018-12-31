using HtmlAgilityPack;
using Microsoft.TeamFoundation.WorkItemTracking.Client;
using Microsoft.TeamFoundation.WorkItemTracking.Proxy;
using Serilog;
using System;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

namespace WordExporter.Core.WorkItems
{
    public static class Extensions
    {
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
