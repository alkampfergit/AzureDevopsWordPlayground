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
            Query = keyValuePairList.GetStringValue("query");
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
            QueryParameters = new List<Dictionary<string, string>>();
            foreach (var parameter in keyValuePairList.Where(kvl => kvl.Key == "parameterSet"))
            {
                var set = ConfigurationParser.ParameterSetList.Parse(parameter.Value);

                var dictionary = set.ToDictionary(k => k.Key, v => v.Value);
                QueryParameters.Add(dictionary);
            }
        }

        public String Query { get; private set; }

        /// <summary>
        /// If this property is set we have a special export where we create a table of 
        /// work items.
        /// </summary>
        public String TableTemplate { get; set; }
        public Int32 Limit { get; set; }

        /// <summary>
        /// <para>
        /// Query can be parametric, this means that we should execute query for each 
        /// series of parameters.
        /// </para>
        /// <para>
        /// If there are no parameters, the query will be executed just one time with 
        /// standard set of parameter.
        /// </para>
        /// </summary>
        public List<Dictionary<String, String>> QueryParameters { get; private set; }

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
            parameters = parameters.ToDictionary(k => k.Key, v => v.Value); //clone
            Dictionary<String, Dictionary<String, Object>> queries = new Dictionary<string, Dictionary<string, Object>>();
            if (QueryParameters.Count == 0)
            {
                queries.Add(Query, parameters);
            }
            else
            {
                foreach (var parameterSet in QueryParameters)
                {
                    var query = Query;
                    foreach (var parameter in parameterSet)
                    {
                        query = query.Replace('{' + parameter.Key + '}', parameter.Value);
                        parameters[parameter.Key] = parameter.Value;
                    }
                    queries.Add(query, parameterSet.ToDictionary(k => k.Key, v => (Object)v.Value));
                }
            }
            WorkItemManger workItemManger = new WorkItemManger(connectionManager);
            workItemManger.SetTeamProject(teamProjectName);

            foreach (var query in queries)
            {
                var workItems = workItemManger.ExecuteQuery(query.Key).Take(Limit);

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

                        manipulator.InsertWorkItem(workItem, templateName, true, query.Value);
                    }
                }
                else
                {
                    //We have a table template, we want to export work item as a list
                    var tableFile = wordTemplateFolderManager.GenerateFullFileName(TableTemplate);
                    var tempFile = wordTemplateFolderManager.CopyFileInTempDirectory(tableFile);
                    using (var tableManipulator = new WordManipulator(tempFile, false))
                    {
                        tableManipulator.SubstituteTokens(query.Value);
                        tableManipulator.FillTableWithCompositeWorkItems(true, workItems);
                    }
                    manipulator.AppendOtherWordFile(tempFile);
                }
            }
        
            base.Assemble(manipulator, parameters, connectionManager, wordTemplateFolderManager, teamProjectName);
        }
    }
}
