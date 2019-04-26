using GalaSoft.MvvmLight;
using GalaSoft.MvvmLight.Command;
using Microsoft.TeamFoundation.Core.WebApi;
using Microsoft.TeamFoundation.Core.WebApi.Types;
using Microsoft.TeamFoundation.Work.WebApi;
using Microsoft.TeamFoundation.WorkItemTracking.WebApi;
using Microsoft.TeamFoundation.WorkItemTracking.WebApi.Models;
using Microsoft.VisualStudio.Services.WebApi;
using Serilog;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Input;
using WordExporter.Core;
using WordExporter.Core.Templates;
using WordExporter.Core.WordManipulation;
using WordExporter.Core.WorkItems;
using WordExporter.UI.Support;
using WordExporter.UI.ViewModel.SubModels;

namespace WordExporter.UI.ViewModel
{
    /// <summary>
    /// This class contains properties that the main View can data bind to.
    /// <para>
    /// Use the <strong>mvvminpc</strong> snippet to add bindable properties to this ViewModel.
    /// </para>
    /// <para>
    /// You can also use Blend to data bind with the tool's support.
    /// </para>
    /// <para>
    /// See http://www.galasoft.ch/mvvm
    /// </para>
    /// </summary>
    public class MainViewModel : ViewModelBase
    {
        /// <summary>
        /// Initializes a new instance of the MainViewModel class.
        /// </summary>
        public MainViewModel()
        {
            //if (IsInDesignMode)
            //{
            //    // Code runs in Blend --> create design time data.
            //}
            //else
            //{
            //    // Code runs "for real"
            //}

            TemplateFolder = StatePersister.Instance.Load<String>("main.TemplateFolder") ?? @"C:\develop\GitHub\AzureDevopsWordPlayground\src\WordExporter\Templates";
            Connect = new RelayCommand(ConnectMethod);
            GetQueries = new RelayCommand(GetQueriesMethod);
            Export = new RelayCommand(ExportMethod);
            Dump = new RelayCommand(DumpMethod);
            GetIterations = new RelayCommand(GetIterationsMethod);
            Address = StatePersister.Instance.Load<String>("main.Address") ?? String.Empty;
        }

        private Boolean _connected;

        public Boolean Connected
        {
            get
            {
                return _connected;
            }
            set
            {
                Set<Boolean>(() => this.Connected, ref _connected, value);
            }
        }

        private String _status;

        public String Status
        {
            get
            {
                return _status;
            }
            set
            {
                Set<String>(() => this.Status, ref _status, value);
            }
        }

        private String _address;

        public String Address
        {
            get
            {
                return _address;
            }
            set
            {
                Set<String>(() => this.Address, ref _address, value);
                StatePersister.Instance.Save("main.Address", value);
            }
        }

        private ObservableCollection<TeamProject> _teamProjects = new ObservableCollection<TeamProject>();

        public ObservableCollection<TeamProject> TeamProjects
        {
            get
            {
                return _teamProjects;
            }
            set
            {
                _teamProjects = value;
                RaisePropertyChanged(nameof(TeamProjects));
            }
        }

        private TeamProject _selectedTeamProject;

        public TeamProject SelectedTeamProject
        {
            get
            {
                return _selectedTeamProject;
            }
            set
            {
                Set<TeamProject>(() => this.SelectedTeamProject, ref _selectedTeamProject, value);
                GetIterationsMethod();
            }
        }

        private ObservableCollection<QueryViewModel> _queries = new ObservableCollection<QueryViewModel>();

        public ObservableCollection<QueryViewModel> Queries
        {
            get
            {
                return _queries;
            }
            set
            {
                _queries = value;
                RaisePropertyChanged(nameof(Queries));
            }
        }

        private QueryViewModel _selectedQuery;

        public QueryViewModel SelectedQuery
        {
            get
            {
                return _selectedQuery;
            }
            set
            {
                Set<QueryViewModel>(() => this.SelectedQuery, ref _selectedQuery, value);
            }
        }

        private String _templateFolder;

        public String TemplateFolder
        {
            get
            {
                return _templateFolder;
            }
            set
            {
                Set<String>(() => this.TemplateFolder, ref _templateFolder, value);
                Templates.Clear();
                if (Directory.Exists(value))
                {
                    TemplateManager = new TemplateManager(TemplateFolder);
                    foreach (var template in TemplateManager.GetTemplateNames())
                    {
                        var wordTemplate = TemplateManager.GetWordDefinitionTemplate(template);
                        var info = new TemplateInfo(template, wordTemplate);
                        Templates.Add(info);
                    }
                    StatePersister.Instance.Save("main.TemplateFolder", value);
                }
                else
                {
                    TemplateManager = null;
                }
            }
        }

        private TemplateManager _templateManager;

        public TemplateManager TemplateManager
        {
            get
            {
                return _templateManager;
            }
            set
            {
                Set<TemplateManager>(() => this.TemplateManager, ref _templateManager, value);
            }
        }

        private ObservableCollection<TemplateInfo> _templates = new ObservableCollection<TemplateInfo>();

        public ObservableCollection<TemplateInfo> Templates
        {
            get
            {
                return _templates;
            }
            set
            {
                _templates = value;
                RaisePropertyChanged(nameof(Templates));
            }
        }

        private ObservableCollection<ParameterViewModel> _parameters = new ObservableCollection<ParameterViewModel>();

        public ObservableCollection<ParameterViewModel> Parameters
        {
            get
            {
                return _parameters;
            }
            set
            {
                _parameters = value;
                RaisePropertyChanged(nameof(Parameters));
            }
        }

        private ObservableCollection<ParameterViewModel> _arrayParameters = new ObservableCollection<ParameterViewModel>();

        public ObservableCollection<ParameterViewModel> ArrayParameters
        {
            get
            {
                return _arrayParameters;
            }
            set
            {
                _arrayParameters = value;
                RaisePropertyChanged(nameof(ArrayParameters));
            }
        }

        private ObservableCollection<IterationsViewModel> _iterations = new ObservableCollection<IterationsViewModel>();

        public ObservableCollection<IterationsViewModel> Iterations
        {
            get
            {
                return _iterations;
            }
            set
            {
                _iterations = value;
                RaisePropertyChanged(nameof(Iterations));
            }
        }

        private TemplateInfo _selectedTemplate;

        public TemplateInfo SelectedTemplate
        {
            get
            {
                return _selectedTemplate;
            }
            set
            {
                Set<TemplateInfo>(() => this.SelectedTemplate, ref _selectedTemplate, value);
                UpdateSelectionOfTemplate();
            }
        }

        private Boolean _generatePdf;

        public Boolean GeneratePdf
        {
            get
            {
                return _generatePdf;
            }
            set
            {
                Set<Boolean>(() => this.GeneratePdf, ref _generatePdf, value);
            }
        }

        public ICommand Connect { get; private set; }

        public ICommand GetQueries { get; private set; }

        public ICommand Export { get; private set; }

        public ICommand Dump { get; private set; }

        public ICommand GetIterations { get; private set; }

        private async void ConnectMethod()
        {
            try
            {
                Status = "Connecting";
                var connectionManager = new ConnectionManager();
                await connectionManager.ConnectAsync(Address);

                Status = "Connected, Retrieving Team Projects";
                var projectHttpClient = connectionManager.GetClient<ProjectHttpClient>();

                await GetTeamProjectAsync(projectHttpClient);
                Status = "Connected, List of team Project retrieved";
                Connected = true;
            }
            catch (Exception ex)
            {
                Status = $"Error during connection: {ex.Message}";
                Log.Error(ex, "Error during connection");
            }

            //// Connect to VSTS
            //TfsTeamProjectCollection tpc = new TfsTeamProjectCollection(_uri, creds);
            //tpc.EnsureAuthenticated();

            ////// Create instance of VssConnection using Visual Studio sign-in prompt
            //VssConnection connection = new VssConnection(new Uri(Address), new VssClientCredentials());
            //WorkItemTrackingHttpClient witClient = connection.GetClient<WorkItemTrackingHttpClient>();
            //var items = witClient.GetQueriesAsync("Jarvis", depth: 3).Result;

            //NetworkCredential netCred = new NetworkCredential("", "");
            //BasicAuthCredential basicCred = new BasicAuthCredential(netCred);
            //TfsClientCredentials tfsCred = new TfsClientCredentials(basicCred);
            //tfsCred.AllowInteractive = true;

            //TfsTeamProjectCollection collection = new TfsTeamProjectCollection(new Uri(Address), tfsCred);
            //collection.Authenticate();
        }

        private Task GetCollectionAsync(ProjectCollectionHttpClient projectCollectionHttpClient)
        {
            return Task.Factory.StartNew(() =>
            {
                foreach (var projectCollectionReference in projectCollectionHttpClient.GetProjectCollections(10, 0).Result)
                {
                    // retrieve a reference to the actual project collection based on its (reference) .Id
                    var projectCollection = projectCollectionHttpClient.GetProjectCollection(projectCollectionReference.Id.ToString()).Result;

                    // the 'web' Url is the one for the PC itself, the API endpoint one is different, see below
                    var webUrlForProjectCollection = projectCollection.Links.Links["web"] as ReferenceLink;
                }
            });
        }

        private async Task GetTeamProjectAsync(ProjectHttpClient projectHttpClient)
        {
            // then - same as above.. iterate over the project references (with a hard-coded pagination of the first 10 entries only)
            var tpList = await Task<List<TeamProject>>.Run(() =>
            {
                List<TeamProject> tempUnorderedListOfTeamProjects = new List<TeamProject>();
                foreach (var projectReference in projectHttpClient.GetProjects(top: 100, skip: 0).Result)
                {
                    // and then get ahold of the actual project
                    var teamProject = projectHttpClient.GetProject(projectReference.Id.ToString()).Result;
                    var urlForTeamProject = ((ReferenceLink)teamProject.Links.Links["web"]).Href;

                    tempUnorderedListOfTeamProjects.Add(teamProject);
                }
                return tempUnorderedListOfTeamProjects;
            });

            foreach (var teamProject in tpList.OrderBy(tp => tp.Name))
            {
                _teamProjects.Add(teamProject);
            }
        }

        public async void GetQueriesMethod()
        {
            WorkItemTrackingHttpClient witClient = ConnectionManager.Instance.GetClient<WorkItemTrackingHttpClient>();
            var queries = await witClient.GetQueriesAsync(SelectedTeamProject.Name, depth: 2, expand: QueryExpand.Wiql);
            Queries.Clear();
            await PopulateQueries(String.Empty, witClient, queries);
        }

        public async void GetIterationsMethod()
        {
            Iterations.Clear();
            if (SelectedTeamProject == null)
            {
                return;
            }

            Status = "Getting iterations for team project " + SelectedTeamProject.Name;

            WorkHttpClient workClient = ConnectionManager.Instance.GetClient<WorkHttpClient>();
            var allIterations = await workClient.GetTeamIterationsAsync(new TeamContext(SelectedTeamProject.Id));

            foreach (var iteration in allIterations)
            {
                Iterations.Add(new IterationsViewModel(iteration));
            }
            Status = "All iteration loaded";
        }

        public void ExportMethod()
        {
            if (TemplateManager == null)
            {
                return;
            }

            if (SelectedTemplate == null)
            {
                return;
            }

            if (ConnectionManager.Instance == null)
            {
                return;
            }

            try
            {
                InnerExecuteExport();
            }
            catch (Exception ex)
            {
                Log.Error(ex, "Error exporting template: {message}", ex.Message);
            }
        }

        public void DumpMethod()
        {
            if (TemplateManager == null)
            {
                return;
            }

            if (SelectedTemplate == null)
            {
                return;
            }

            if (ConnectionManager.Instance == null)
            {
                return;
            }

            try
            {
                InnerExecuteDump();
            }
            catch (Exception ex)
            {
                Log.Error(ex, "Error exporting template: {message}", ex.Message);
            }
        }

        private void InnerExecuteDump()
        {
            var fileName = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString()) + ".txt";
            if (SelectedTemplate.IsScriptTemplate)
            {
                var executor = new TemplateExecutor(SelectedTemplate.WordTemplateFolderManager);

                //now we need to ask user parameter value
                Dictionary<string, object> parameters = PrepareUserParameters();
                executor.DumpWorkItem(fileName, ConnectionManager.Instance, SelectedTeamProject.Name, parameters);
            }
            else
            {
                var selected = SelectedQuery?.Results?.Where(q => q.Selected).ToList();
                if (selected == null || selected.Count == 0)
                {
                    return;
                }

                var sb = new StringBuilder();
                foreach (var workItemResult in selected)
                {
                    var workItem = workItemResult.WorkItem;
                    var values = workItem.CreateDictionaryFromWorkItem();
                    foreach (var value in values)
                    {
                        sb.AppendLine($"{value.Key.PadRight(50, ' ')}={value.Value}");
                    }
                    File.WriteAllText(fileName, sb.ToString());
                }

            }
            System.Diagnostics.Process.Start(fileName);
        }

        private void InnerExecuteExport()
        {
            var baseFolder = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            if (SelectedTemplate.IsScriptTemplate)
            {
                if (ArrayParameters.Any())
                {
                    var arrayParameters = ArrayParameters.Select(p => new
                    {
                        Name = p.Name,
                        Values = p.Value?.Split(',', ';').ToList() ?? new List<string>()
                    })
                    .ToList();

                    Int32 maxParameterCount = arrayParameters.Max(p => p.Values.Count);
                    StringBuilder fileSuffix = new StringBuilder();

                    for (int i = 0; i < maxParameterCount; i++)
                    {
                        Dictionary<string, object> parameters = PrepareUserParameters();
                        foreach (var arrayParameter in arrayParameters)
                        {
                            var value = arrayParameter.Values.Count > i ? arrayParameter.Values[i] : String.Empty;
                            parameters[arrayParameter.Name] = value;
                            fileSuffix.Append(arrayParameter.Name);
                            fileSuffix.Append("_");
                            fileSuffix.Append(value);
                        }
                        var fileName = Path.Combine(baseFolder, SelectedTemplate.TemplateName + "_" + DateTime.Now.ToString("dd_MM_yyyy hh mm")) + "_" + fileSuffix.ToString() + ".docx";
                        GenerateFileFromScriptTemplate(fileName, parameters);
                    }
                }
                else
                {
                    var fileName = Path.Combine(baseFolder, SelectedTemplate.TemplateName + "_" + DateTime.Now.ToString("dd_MM_yyyy hh mm")) + ".docx";
                    Dictionary<string, object> parameters = PrepareUserParameters();
                    GenerateFileFromScriptTemplate(fileName, parameters);
                }
            }
            else
            {
                var fileName = Path.Combine(baseFolder, SelectedTemplate.TemplateName + "_" + DateTime.Now.ToString("dd_MM_yyyy hh mm")) + ".docx";
                var selected = SelectedQuery?.Results?.Where(q => q.Selected).ToList();
                if (selected == null || selected.Count == 0)
                {
                    return;
                }

                var template = SelectedTemplate.WordTemplateFolderManager;
                using (WordManipulator manipulator = new WordManipulator(fileName, true))
                {
                    foreach (var workItemResult in selected)
                    {
                        var workItem = workItemResult.WorkItem;
                        manipulator.InsertWorkItem(workItem, template.GetTemplateFor(workItem.Type.Name), true);
                    }
                }
                ManageGeneratedWordFile(fileName);
            }
        }

        private void GenerateFileFromScriptTemplate(string fileName, Dictionary<string, object> parameters)
        {
            var executor = new TemplateExecutor(SelectedTemplate.WordTemplateFolderManager);
            executor.GenerateWordFile(fileName, ConnectionManager.Instance, SelectedTeamProject.Name, parameters);
            ManageGeneratedWordFile(fileName);
        }

        private void ManageGeneratedWordFile(string fileName)
        {
            if (GeneratePdf)
            {
                using (WordAutomationHelper helper = new WordAutomationHelper(fileName, false))
                {
                    var pdfFile = helper.ConvertToPdf();
                    if (!String.IsNullOrEmpty(pdfFile))
                    {
                        System.Diagnostics.Process.Start(pdfFile);
                    }
                }
            }
            else
            {
                System.Diagnostics.Process.Start(fileName);
            }
        }

        private Dictionary<string, object> PrepareUserParameters()
        {
            Dictionary<string, Object> parameters = new Dictionary<string, object>();
            foreach (var parameter in Parameters)
            {
                parameters[parameter.Name] = parameter.Value;
            }
            List<String> iterations = Iterations.Where(i => i.Selected).Select(i => i.Path).ToList();
            parameters["iterations"] = iterations;
            return parameters;
        }

        private async Task PopulateQueries(String actualPath, WorkItemTrackingHttpClient witClient, IEnumerable<QueryHierarchyItem> queries)
        {
            foreach (var query in queries)
            {
                if (query.IsFolder != true)
                {
                    Queries.Add(new QueryViewModel(this, actualPath, query));
                }
                if (query.HasChildren == true)
                {
                    var newPath = actualPath + '/' + query.Name;
                    if (query.Children == null)
                    {
                        //need to requery the store to grab reference to the query.
                        var queryReloaded = await witClient.GetQueryAsync(SelectedTeamProject.Id, query.Path, depth: 2, expand: QueryExpand.Wiql);
                        await PopulateQueries(newPath, witClient, queryReloaded.Children);
                    }
                    else
                    {
                        await PopulateQueries(newPath, witClient, query.Children);
                    }
                }
            }
        }

        private void UpdateSelectionOfTemplate()
        {
            Parameters.Clear();
            ArrayParameters.Clear();
            if (SelectedTemplate == null)
            {
                return;
            }

            if (SelectedTemplate.IsScriptTemplate)
            {
                foreach (var parameter in SelectedTemplate.Parameters)
                {
                    Parameters.Add(new ParameterViewModel(parameter.Key, parameter.Value));
                }

                foreach (var parameter in SelectedTemplate.ArrayParameters)
                {
                    ArrayParameters.Add(new ParameterViewModel(parameter.Key, parameter.Value));
                }
            }
        }
    }
}