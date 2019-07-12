using HtmlAgilityPack;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WordExporter.Core.Support
{
    public static class HtmlAgilityToolkitExtension
    {
        /// <summary>
        /// Scan the document for all tag in <paramref name="tagNames"/>
        /// and remove all the tags, leaving the content. Optionally after
        /// each substituted tag you can append a <paramref name="nodeToAppendAtClosingTag"/>.
        /// 
        /// </summary>
        /// <param name="document"></param>
        /// <param name="nodeToAppendAtClosingTag">Null if you do not want to append nothing
        /// after the tag is closed, different from null if you want to append a node
        /// after each manipulated node.</param>
        /// <param name="tagNames"></param>
        public static void RemoveTags(
            this HtmlDocument document,
            HtmlNode nodeToAppendAtClosingTag,
            params String[] tagNames)
        {
            foreach (var tag in tagNames)
            {
                var nodesToRemove = document.DocumentNode.SelectNodes($"//{tag}");
                if (nodesToRemove != null)
                {
                    var nodes = nodesToRemove.ToList();
                    foreach (var nodeToRemove in nodes)
                    {
                        var parent = nodeToRemove.ParentNode;
                        var children = nodeToRemove.ChildNodes.ToList();
                        foreach (var child in children)
                        {
                            child.Remove();
                            parent.InsertBefore(child, nodeToRemove);
                        }
                        if (nodeToAppendAtClosingTag != null)
                        {
                            parent.InsertBefore(nodeToAppendAtClosingTag.Clone(), nodeToRemove);
                        }
                        nodeToRemove.Remove();
                    }
                }
            }
        }
    }
}
