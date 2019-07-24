using Microsoft.TeamFoundation.WorkItemTracking.Client;
using Serilog;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

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
}
