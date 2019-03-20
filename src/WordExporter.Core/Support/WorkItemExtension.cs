using Microsoft.TeamFoundation.WorkItemTracking.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WordExporter.Core.Support
{
    public static class WorkItemExtension
    {
        public static String GetFieldValueAsString(this WorkItem workItem, String fieldName)
        {
            return workItem.Fields[fieldName]?.Value?.ToString() ?? String.Empty;
        }
    }
}
