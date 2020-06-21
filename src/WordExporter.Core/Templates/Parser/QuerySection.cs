using Microsoft.TeamFoundation.WorkItemTracking.Client;
using Serilog;
using Sprache;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using WordExporter.Core.ExcelManipulation;
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
            Name = keyValuePairList.GetStringValue("name");
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
            RepeatForEachIteration = keyValuePairList.GetBooleanValue("repeatForEachIteration");
            var workItemTypes = keyValuePairList.GetStringValue("workItemTypes");
            if (!String.IsNullOrEmpty(workItemTypes))
            {
                this.WorkItemTypes = workItemTypes.Split(',');
            }
            foreach (var parameter in keyValuePairList.Where(kvl => kvl.Key == "parameterSet"))
            {
                var set = ConfigurationParser.ParameterSetList.Parse(parameter.Value);

                var dictionary = set.ToDictionary(k => k.Key, v => v.Value);
                QueryParameters.Add(dictionary);
            }

            var hierarchyModeString = keyValuePairList.GetStringValue("hierarchyMode");
            if (!String.IsNullOrEmpty(hierarchyModeString))
            {
                HierarchyMode = hierarchyModeString.Split(',', ';').Select(s => s.Trim(' ')).ToArray();
            }
            PageBreak = keyValuePairList.GetBooleanValue("pageBreak");
        }

        public String Name { get; private set; }

        public String Query { get; private set; }

        /// <summary>
        /// If this property is set we have a special export where we create a table of 
        /// work items.
        /// </summary>
        public String TableTemplate { get; set; }
        public Int32 Limit { get; set; }

        /// <summary>
        /// If this value contains at least one element, it will used to export
        /// only the work item in this lists. This allows you to do hierarchical
        /// queries, but export only childs or fathers.
        /// </summary>
        public String[] WorkItemTypes { get; set; }

        /// <summary>
        /// If true it will insert a page break after each work item.
        /// </summary>
        public Boolean PageBreak { get; set; }

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

        private readonly Dictionary<String, String> SpecificTemplates = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        public bool RepeatForEachIteration { get; private set; }

        /// <summary>
        /// <para>
        /// This is a special property, if contains series of types that will trigger a real specific query 
        /// where:
        /// </para>
        /// <para>
        /// 1) the component will execute the query.
        /// 2) the component will start from the first type, and for each type it will automatically
        /// traverse all the parent to fulfill hierarchy.
        /// </para>
        /// </summary>
        public String[] HierarchyMode { get; private set; }

        public String GetTemplateForWorkItem(string workItemTypeName)
        {
            return SpecificTemplates.TryGetValue(workItemTypeName);
        }

        #region Syntax

        public readonly static Parser<QuerySection> Parser =
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
            WorkItemManger workItemManger = PrepareWorkItemManager(connectionManager, teamProjectName);

            parameters = parameters.ToDictionary(k => k.Key, v => v.Value); //clone
            //If we do not have query parameters we have a single query or we can have parametrized query with iterationPath
            var queries = PrepareQueries(parameters);

            foreach (var query in queries)
            {
#if DEBUG
                //Limit = 20;
#endif
                List<WorkItem> queryRawReturnValue = workItemManger.ExecuteQuery(query)
                    .Take(Limit)
                    .ToList();
                var workItems = queryRawReturnValue
                    .Where(ShouldExport)
                    .ToList();

                //Add the table only if whe really have work item selected.
                if (String.IsNullOrEmpty(TableTemplate))
                {
                    if (workItems.Count > 0)
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

                            manipulator.InsertWorkItem(workItem, templateName, PageBreak, parameters);
                        }
                    }
                }
                else
                {
                    //We have a table template, we want to export work item as a list
                    var tableFile = wordTemplateFolderManager.GenerateFullFileName(TableTemplate);
                    var tempFile = wordTemplateFolderManager.CopyFileInTempDirectory(tableFile);
                    using (var tableManipulator = new WordManipulator(tempFile, false))
                    {
                        tableManipulator.SubstituteTokens(parameters);
                        tableManipulator.FillTableWithCompositeWorkItems(true, workItems, workItemManger);
                    }
                    manipulator.AppendOtherWordFile(tempFile);
                }
            }

            base.Assemble(manipulator, parameters, connectionManager, wordTemplateFolderManager, teamProjectName);
        }

        public override void AssembleExcel(
            ExcelManipulator manipulator,
            Dictionary<string, object> parameters,
            ConnectionManager connectionManager,
            WordTemplateFolderManager wordTemplateFolderManager,
            string teamProjectName)
        {
            WorkItemManger workItemManger = PrepareWorkItemManager(connectionManager, teamProjectName);
            //If we do not have query parameters we have a single query or we can have parametrized query with iterationPath
            var queries = PrepareQueries(parameters);

            foreach (var query in queries)
            {
                if (HierarchyMode?.Length > 0)
                {
                    var hr = workItemManger.ExecuteHierarchicQuery(query, HierarchyMode);
                    manipulator.FillWorkItems(hr);
                }
                else
                {
                    throw new NotSupportedException("This version of the program only support hierarchy mode for excel");
                }
            }
        }

        public override void Dump(
            StringBuilder stringBuilder,
            Dictionary<string, object> parameters,
            ConnectionManager connectionManager,
            WordTemplateFolderManager wordTemplateFolderManager,
            string teamProjectName)
        {
            WorkItemManger workItemManger = PrepareWorkItemManager(connectionManager, teamProjectName);

            parameters = parameters.ToDictionary(k => k.Key, v => v.Value); //clone
            //If we do not have query parameters we have a single query or we can have parametrized query with iterationPath
            var queries = PrepareQueries(parameters);

            foreach (var query in queries)
            {
                var workItems = workItemManger.ExecuteQuery(query).Take(Limit);
                foreach (var workItem in workItems.Where(ShouldExport))
                {
                    var values = workItem.CreateDictionaryFromWorkItem();
                    foreach (var value in values)
                    {
                        stringBuilder.AppendLine($"{value.Key.PadRight(50, ' ')}={value.Value}");
                    }
                }
            }
        }

        private bool ShouldExport(WorkItem workItem)
        {
            return WorkItemTypes == null || WorkItemTypes.Contains(workItem.Type.Name);
        }

        private static WorkItemManger PrepareWorkItemManager(ConnectionManager connectionManager, string teamProjectName)
        {
            WorkItemManger workItemManger = new WorkItemManger(connectionManager);
            workItemManger.SetTeamProject(teamProjectName);
            return workItemManger;
        }

        private List<String> PrepareQueries(Dictionary<string, object> parameters)
        {
            var retValue = new List<String>();
            if (QueryParameters.Count == 0)
            {
                if (!RepeatForEachIteration)
                {
                    retValue.Add(SubstituteParametersInQuery(Query, parameters));
                }
                else
                {
                    //TODO: THis is a logic that should propably be completely changed with the concept of array parameters.
                    if (parameters.TryGetValue("iterations", out object iterations)
                        && iterations is List<String> iterationList)
                    {
                        foreach (var iteration in iterationList)
                        {
                            var query = Query.Replace("{iterationPath}", iteration);
                            retValue.Add(SubstituteParametersInQuery(query, parameters));
                        }
                    }
                    else
                    {
                        //By convention we should have a list of valid iterations names inside parameters dictionary.
                        Log.Error("Error handling iteration for query {name}, we have RepeatForEachIteration to true but no iteration defined", Name);
                    }
                }
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
                    retValue.Add(SubstituteParametersInQuery(query, parameters));
                }
            }

            return retValue;
        }

        private string SubstituteParametersInQuery(string query, Dictionary<string, object> parameters)
        {
            foreach (var parameter in parameters)
            {
                var value = parameter.Value?.ToString() ?? "";
                query = query.Replace('{' + parameter.Key + '}', value);
            }
            return query;
        }
    }
}
