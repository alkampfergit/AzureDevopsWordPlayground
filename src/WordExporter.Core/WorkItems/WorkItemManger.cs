using Microsoft.TeamFoundation.WorkItemTracking.Client;
using Serilog;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using WordExporter.Core.Support;

namespace WordExporter.Core.WorkItems
{
    public class WorkItemManger
    {
        private readonly ConnectionManager _connection;
        private string _teamProjectName;

        public WorkItemManger(ConnectionManager connection)
        {
            _connection = connection;
        }

        public void SetTeamProject(String teamProjectName)
        {
            _teamProjectName = teamProjectName;
        }

        /// <summary>
        /// Load a series of Work Items for a given area and a given iteration.
        /// </summary>
        /// <param name="areaPath"></param>
        /// <param name="iterationPath"></param>
        /// <returns></returns>
        public List<WorkItem> LoadAllWorkItemForAreaAndIteration(string areaPath, string iterationPath)
        {
            return ExecuteQuery($"SELECT * FROM WorkItems Where [System.AreaPath] UNDER '{areaPath}' AND [System.IterationPath] UNDER '{iterationPath}'");
        }

        public List<WorkItem> ExecuteQuery(string wiqlQuery)
        {
            StringBuilder query = new StringBuilder();
            query.AppendLine(wiqlQuery);
            if (!String.IsNullOrEmpty(_teamProjectName))
            {
                query = query.Replace("{teamProjectName}", _teamProjectName);
            }

            try
            {
                var realQuery = query.ToString();
                Log.Information("About to execute query {query}", realQuery);

                if (!realQuery.Contains("workitemLinks"))
                {
                    return _connection.WorkItemStore.Query(realQuery)
                        .OfType<WorkItem>()
                        .ToList();
                }
                return ExecuteLinkedQuery(realQuery);
            }
            catch (Exception ex)
            {
                Log.Error(ex, "Error executing Query [{message}]\n{query}", ex.Message, query.ToString());
                throw;
            }
        }

        /// <summary>
        /// <para>
        /// Executes the query, then it start considering only the very first type of 
        /// <paramref name="types"/> array as valid. From each work item of that type
        /// it start crawling the hierarchy up to all the remaining types in <paramref name="types"/>.
        /// </para>
        /// <para>
        /// Ex. if types is Task,Requirements,Feature it will execute the query, return all tasks
        /// and all their parents up to Feature.
        /// </para>
        /// </summary>
        /// <param name="wiqlQuery"></param>
        /// <param name="types"></param>
        /// <returns>A list of linked work items to store a work item and all the parents
        /// up to the hierarchy.</returns>
        public List<LinkedWorkItem> ExecuteHierarchicQuery(string wiqlQuery, String[] types)
        {
            if (types.Length == 0)
                return new List<LinkedWorkItem>();

            var queryResult = ExecuteQuery(wiqlQuery);
            var resultList = queryResult
                .Where(wi => wi.Type.Name.Equals(types[0], StringComparison.OrdinalIgnoreCase))
                .Select(wi => new LinkedWorkItem(wi))
                .ToList();

            //ok we need to start crawling up the hierarcy up to the final type.
            IEnumerable<LinkedWorkItem> actualList = resultList;
            foreach (var type in types.Skip(1))
            {
                var allParentId = actualList
                    .Where(wi => wi.ParentId.HasValue)
                    .Select(wi => wi.ParentId.Value);

                //now we need to query all the parent, TODO: execute in block and not with a single query.
                var listOfParentWorkItems = LoadListOfWorkItems(allParentId)
                    .Select(wi => new LinkedWorkItem(wi))
                    .ToList();

                //now create a dictionary to simplify the lookup
                var actualDictionaryList = listOfParentWorkItems.ToDictionary(wi => wi.WorkItem.Id);
                foreach (var item in actualList.Where(wi => wi.ParentId.HasValue))
                {
                    if (actualDictionaryList.TryGetValue(item.ParentId.Value, out var parentWorkItem))
                    {
                        item.Parent = parentWorkItem;
                    }
                }

                actualList = listOfParentWorkItems;
            }

            return resultList;
        }

        /// <summary>
        /// Accepts a list of id of work items and returns all the work items with that id
        /// </summary>
        /// <param name="idList"></param>
        /// <returns></returns>
        public List<WorkItem> LoadListOfWorkItems(IEnumerable<int> idList)
        {
            //optimize, load all parents in a dictionary with a single query
            if (!idList.Any())
            {
                return new List<WorkItem>();
            }

            //ok now we need to grab all parent link, just to grab 
            var query = $@"SELECT
                    [System.Id],
                    [System.Title]
                FROM workitems
                WHERE [System.Id] IN ({String.Join(",", idList)})
                ORDER BY [System.Id]";

            //ok, query all the parents
            return _connection.WorkItemStore.Query(query)
                .OfType<WorkItem>()
                .ToList();
        }

        private List<WorkItem> ExecuteLinkedQuery(string realQuery)
        {
            var linkQuery = new Query(_connection.WorkItemStore, realQuery);
            WorkItemLinkInfo[] witLinkInfos = linkQuery.RunLinkQuery();
            Dictionary<Int32, WorkItem> result = new Dictionary<int, WorkItem>();
            foreach (WorkItemLinkInfo witinfo in witLinkInfos)
            {
                LoadWorkItem(result, witinfo.TargetId);
                LoadWorkItem(result, witinfo.SourceId);
            }
            return result.Values.ToList();
        }

        private void LoadWorkItem(Dictionary<int, WorkItem> result, int workItemId)
        {
            if (workItemId > 0 && !result.ContainsKey(workItemId))
            {
                result[workItemId] = _connection.WorkItemStore.GetWorkItem(workItemId);
            }
        }
    }

    /// <summary>
    /// Simple in-memory structure to store work item and parents.
    /// </summary>
    public class LinkedWorkItem
    {
        public LinkedWorkItem(WorkItem workItem)
        {
            WorkItem = workItem;
            ParentId = workItem.GetParentLink()?.RelatedWorkItemId;
        }

        public WorkItem WorkItem { get; set; }

        public Int32? ParentId { get; set; }

        public LinkedWorkItem Parent { get; set; }
    }
}
