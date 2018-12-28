using CommandLine;
using Microsoft.TeamFoundation.VersionControl.Client;
using Serilog;
using Serilog.Exceptions;
using System;
using System.Collections.Generic;
using System.IO;
using WordExporter.Core;
using WordExporter.Core.WordManipulation;
using WordExporter.Core.WorkItems;
using WordExporter.Support;

namespace WordExporter
{
    class Program
    {
        static void Main(string[] args)
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

            Connection connection = new Connection(options.ServiceAddress, options.GetAccessToken());

            DumpAllTeamProjects(connection);


            WorkItemManger workItemManger = new WorkItemManger(connection);
            workItemManger.SetTeamProject("zoalord insurance");
            var workItems = workItemManger.LoadAllWorkItemForAreaAndIteration(
                "zoalord insurance",
                "zoalord insurance\\Release 1\\Sprint 6");

            var fileName = Path.GetTempFileName() + ".docx";
            using (WordManipulator manipulator = new WordManipulator(fileName, true))
            {
                foreach (var workItem in workItems)
                {
                    manipulator.InsertWorkItem(workItem);
                }
            }

            System.Diagnostics.Process.Start(fileName);

            if (Environment.UserInteractive)
            {
                Console.WriteLine("Execution completed, press a key to continue");
                Console.ReadKey();
            }
        }

        private static void DumpAllTeamProjects(Connection connection)
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
                    "logs.txt",
                     rollingInterval: RollingInterval.Day
                )
                .WriteTo.File(
                    "errors.txt",
                     rollingInterval: RollingInterval.Day,
                     restrictedToMinimumLevel: Serilog.Events.LogEventLevel.Error
                )
                .CreateLogger();
        }
    }
}
