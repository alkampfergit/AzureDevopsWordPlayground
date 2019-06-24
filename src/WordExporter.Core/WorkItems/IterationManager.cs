using Microsoft.TeamFoundation.Server;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml;
using System.Xml.Linq;

namespace WordExporter.Core.WorkItems
{
    public class IterationManager
    {
        private readonly ConnectionManager _connectionManager;

        public IterationManager(ConnectionManager connectionManager)
        {
            _connectionManager = connectionManager;
        }

        public IEnumerable<IterationInfo> GetAllIterationsForTeamProject(string teamProjectName)
        {
            var project = _connectionManager.GetTeamProject(teamProjectName);
            NodeInfo[] nodes = _connectionManager
                .CommonStructureService.ListStructures(
                project.Uri.AbsoluteUri);
            var iterationRootNode = nodes.Single(n => n.Name.Equals("iteration", StringComparison.OrdinalIgnoreCase));
            List<IterationInfo> retValue = new List<IterationInfo>();
            var itRoot = _connectionManager
               .CommonStructureService
               .GetNodesXml(new string[] { iterationRootNode.Uri }, true);
            var rootNode = itRoot.FirstChild as XmlElement;
            var xml = rootNode.OuterXml;
            var element = XElement.Parse(xml);
            var xmlNodes = element.Descendants("Node");
            foreach (var node in xmlNodes)
            {
                var path = node.Attribute("Path").Value.Trim('\\', '/');
                var splitted = path.Split('\\');
                var normalizedPath = splitted[0] + "\\" + String.Join("\\", splitted.Skip(2));
                retValue.Add(new IterationInfo()
                {
                    Name = node.Attribute("Name").Value.Trim('\\', '/'),
                    Path = normalizedPath,
                    StartDate = node.Attribute("StartDate")?.Value,
                    EndDate = node.Attribute("FinishDate")?.Value,
                }) ;
            }
            return retValue;
        }

        //private void PopulateIterations(
        //    XmlElement node,
        //    List<IterationInfo> retValue)
        //{
        //    retValue.Add(new IterationInfo()
        //    {
        //        Name = node.Attributes["Name"].Value,
        //        Path = node.Attributes["Path"].Value,
        //    });
        //    var children = node.GetElementsByTagName("Children");
        //    foreach (XmlElement child in children)
        //    {
        //        var elementNode = child.FirstChild as XmlElement;
        //        PopulateIterations(elementNode, retValue);
        //    }
        //}

        public class IterationInfo
        {
            public String Name { get; set; }
            public String Path { get; set; }
            public String StartDate { get; set; }

            public String EndDate { get; set; }
        }
    }
}
