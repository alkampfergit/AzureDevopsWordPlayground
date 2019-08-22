using CommandLine;
using Microsoft.TeamFoundation.VersionControl.Client;
using Microsoft.TeamFoundation.WorkItemTracking.Client;
using Serilog;
using Serilog.Exceptions;
using System;
using System.Collections.Generic;
using System.IO;
using WordExporter.Core;
using WordExporter.Core.ExcelManipulation;
using WordExporter.Core.Support;
using WordExporter.Core.Templates;
using WordExporter.Core.WordManipulation;
using WordExporter.Core.WorkItems;
using WordExporter.Support;

namespace WordExporter
{
    public static class Program
    {
        [STAThread]
        private static void Main(string[] args)
        {
            ConfigureSerilog();

            AppDomain.CurrentDomain.UnhandledException += CurrentDomain_UnhandledException;

            var result = CommandLine.Parser.Default.ParseArguments<Options>(args)
               .WithParsed<Options>(opts => options = opts)
               .WithNotParsed<Options>((errs) => HandleParseError(errs));

            if (result.Tag != ParserResultType.Parsed)
            {
                Log.Logger.Error("Command line parameters error, press a key to continue!");
                Console.ReadKey();
                return;
            }

            ConnectionManager connection = new ConnectionManager(options.ServiceAddress, options.GetAccessToken());
            
            //DumpAllIterations(connection);
            //DumpAllTeamProjects(connection);
            //TestExcelExtraction(connection);
            if (String.IsNullOrEmpty(options.TemplateFolder))
            {
                PerformStandardIterationExport(connection);
            }
            else
            {
                PerformTemplateExport(connection);
            }

            if (Environment.UserInteractive)
            {
                Console.WriteLine("Execution completed, press a key to continue");
                Console.ReadKey();
            }
        }

        private static void TestExcelExtraction(ConnectionManager connection)
        {
            WorkItemManger workItemManger = new WorkItemManger(connection);
            workItemManger.SetTeamProject("CMMI Playground");
            var hr = workItemManger.ExecuteHierarchicQuery(@"SELECT
    *
FROM workitemLinks
WHERE
    (
        [Source].[System.TeamProject] = '{teamProjectName}'
        AND[Source].[System.WorkItemType] = 'Feature'
        AND[Source].[Microsoft.VSTS.Scheduling.TargetDate] < '2062-01-01T00:00:00.0000000'
        AND[Source].[Microsoft.VSTS.Scheduling.TargetDate] > '1990-02-02T00:00:00.0000000'
    )
    AND(
        [System.Links.LinkType] = 'System.LinkTypes.Hierarchy-Forward'
    )
    AND(
        [Target].[System.TeamProject] = '{teamProjectName}'
        AND[Target].[System.WorkItemType] <> ''
    )
MODE(Recursive)
",
new[] { "task", "requirement", "feature", "epic" });

            var tempFile = @"c:\temp\test2.xlsx";
            if (File.Exists(tempFile))
                File.Delete(tempFile);

            File.Copy(@"c:\temp\test.xlsx", tempFile);
            using (ExcelManipulator m = new ExcelManipulator(tempFile))
            {
                m.FillWorkItems(hr);
            }
            System.Diagnostics.Process.Start(tempFile);
        }

        private static void DumpAllIterations(ConnectionManager connection)
        {
            var im = new IterationManager(connection);
            var iterations = im.GetAllIterationsForTeamProject("cmmi playground");
            foreach (var iteration in iterations)
            {
                Console.WriteLine(iteration.Path);
            }
        }

        private static void PerformTemplateExport(ConnectionManager connection)
        {
            var wordFolderManager = new WordTemplateFolderManager(options.TemplateFolder);
            var executor = new TemplateExecutor(wordFolderManager);

            //now we need to ask user parameter value
            Dictionary<string, Object> parameters = new Dictionary<string, object>();
            foreach (var parameterName in wordFolderManager.TemplateDefinition.ParameterSection.Parameters.Keys)
            {
                Console.Write($"Parameter {parameterName}:");
                parameters[parameterName] = Console.ReadLine();
            }

            var tempFileName = Path.GetTempPath() + Guid.NewGuid().ToString();
            var generatedName = executor.GenerateFile(tempFileName, connection, options.TeamProject, parameters);
            System.Diagnostics.Process.Start(generatedName);
        }

        private static void PerformStandardIterationExport(ConnectionManager connection)
        {
            WorkItemManger workItemManger = new WorkItemManger(connection);
            workItemManger.SetTeamProject(options.TeamProject);
            var workItems = workItemManger.LoadAllWorkItemForAreaAndIteration(
                options.AreaPath,
                options.IterationPath);

            var fileName = Path.GetTempFileName() + ".docx";
            var templateManager = new TemplateManager("Templates");
            var template = templateManager.GetWordDefinitionTemplate(options.TemplateName);
            using (WordManipulator manipulator = new WordManipulator(fileName, true))
            {
                AddTableContent(manipulator, workItems, template);
                foreach (var workItem in workItems)
                {
                    manipulator.InsertWorkItem(workItem, template.GetTemplateFor(workItem.Type.Name), true);
                }
            }

            System.Diagnostics.Process.Start(fileName);
        }

        private static void AddTableContent(
            WordManipulator manipulator,
            List<WorkItem> workItems,
            WordTemplateFolderManager template)
        {
            string tableFileName = template.GetTable("A", true);
            using (var tableManipulator = new WordManipulator(tableFileName, false))
            {
                List<List<String>> table = new List<List<string>>();
                foreach (WorkItem workItem in workItems)
                {
                    List<String> row = new List<string>();
                    row.Add(workItem.Id.ToString());
                    row.Add(workItem.GetFieldValueAsString("System.AssignedTo"));
                    row.Add(workItem.AttachedFileCount.ToString());
                    table.Add(row);
                }
                tableManipulator.FillTable(true, table);
            }
            manipulator.AppendOtherWordFile(tableFileName);
            File.Delete(tableFileName);
        }

        private static void DumpAllTeamProjects(ConnectionManager connection)
        {
            foreach (var tpname in connection.GetTeamProjectsNames())
            {
                Log.Debug("Team Project {tpname}", tpname);
            }
        }

        private static void CurrentDomain_UnhandledException(object sender, UnhandledExceptionEventArgs e)
        {
            Log.Error(e.ExceptionObject as Exception, "Unhandled exception in the program: {message}", e.ExceptionObject.ToString());
        }

        private static Options options;

        private static void HandleParseError(IEnumerable<Error> errs)
        {
        }

        private static void ConfigureSerilog()
        {
            Log.Logger = new LoggerConfiguration()
                .Enrich.WithExceptionDetails()
                .MinimumLevel.Debug()
                .WriteTo.Console()
                .WriteTo.File(
                    "logs\\logs.txt",
                     rollingInterval: RollingInterval.Day
                )
                .WriteTo.File(
                    "logs\\errors.txt",
                     rollingInterval: RollingInterval.Day,
                     restrictedToMinimumLevel: Serilog.Events.LogEventLevel.Error
                )
                .CreateLogger();
        }
    }
}
