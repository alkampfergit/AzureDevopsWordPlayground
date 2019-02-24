using Sprache;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WordExporter.Core.Support;
using WordExporter.Core.WordManipulation;
using WordExporter.Core.WorkItems;

namespace WordExporter.Core.Templates.Parser
{
    public sealed class QuerySection : Section
    {
        private QuerySection(IEnumerable<KeyValue> keyValuePairList)
        {
            var dicList = keyValuePairList.ToDictionary(
                k => k.Key,
                k => k.Value,
                StringComparer.OrdinalIgnoreCase);
            Query = dicList.TryGetValue("query");
            foreach (var templateKeys in keyValuePairList.Where(k => k.Key
                .StartsWith("template/"))
                .Select(k => new
                {
                    realKey = k.Key.Substring("template/".Length),
                    value = k.Value
                }))
            {
                SpecificTemplates[templateKeys.realKey] = templateKeys.value;
            }
            TableTemplate = keyValuePairList.GetStringValue("tableTemplate");
           Limit = keyValuePairList.GetIntValue("limit", Int32.MaxValue);
        }

        public String Query { get; private set; }

        /// <summary>
        /// If this property is set we have a special export where we create a table of 
        /// work items.
        /// </summary>
        public String TableTemplate { get; set; }
        public Int32 Limit { get; set; }

        private Dictionary<String, String> SpecificTemplates = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        public String GetTemplateForWorkItem(string workItemTypeName)
        {
            return SpecificTemplates.TryGetValue(workItemTypeName);
        }

        #region syntax

        public static Parser<QuerySection> Parser =
          from keyValueList in ConfigurationParser.KeyValueList
          select new QuerySection(keyValueList);

        #endregion

        public override void Assemble(
            WordManipulator manipulator,
            Dictionary<string, object> parameters,
            ConnectionManager connectionManager,
            WordTemplateFolderManager wordTemplateFolderManager,
            string teamProjectName)
        {
            WorkItemManger workItemManger = new WorkItemManger(connectionManager);
            workItemManger.SetTeamProject(teamProjectName);
            var workItems = workItemManger.ExecuteQuery(Query)
                .Take(Limit);

            if (String.IsNullOrEmpty(TableTemplate))
            {
                foreach (var workItem in workItems)
                {
                    if (!SpecificTemplates.TryGetValue(workItem.Type.Name, out var templateName))
                    {
                        templateName = wordTemplateFolderManager.GetTemplateFor(workItem.Type.Name);
                    }
                    else
                    {
                        templateName = wordTemplateFolderManager.GenerateFullFileName(templateName);
                    }

                    manipulator.InsertWorkItem(workItem, templateName, true);
                }
            }
            else
            {
                //We have a table template, we want to export work item as a list
                var tableFile = wordTemplateFolderManager.GenerateFullFileName(TableTemplate);
                var tempFile = wordTemplateFolderManager.CopyFileInTempDirectory(tableFile);
                using (var tableManipulator = new WordManipulator(tempFile, false))
                {
                    tableManipulator.FillTableWithCompositeWorkItems(true, workItems);
                }
                manipulator.AppendOtherWordFile(tempFile);
            }
            base.Assemble(manipulator, parameters, connectionManager, wordTemplateFolderManager, teamProjectName);
        }
    }
}
