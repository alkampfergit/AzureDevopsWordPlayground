using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WordExporter.Core.WorkItems
{
    public class WorkItemHelper
    {
        private readonly Connection _connection;

        public WorkItemHelper(Connection connection)
        {
            _connection = connection;
        }
    }
}
