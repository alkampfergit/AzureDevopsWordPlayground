using Microsoft.TeamFoundation.WorkItemTracking.Client;
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
                if (wiqlQuery.IndexOf(" where ", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    query.AppendLine($" AND [System.TeamProject] = '{_teamProjectName}'");
                }
                else
                {
                    query.AppendLine($" WHERE [System.TeamProject] = '{_teamProjectName}'");
                }
            }

            return _connection.WorkItemStore.Query(query.ToString())
                .OfType<WorkItem>()
                .ToList();
        }
    }
}
