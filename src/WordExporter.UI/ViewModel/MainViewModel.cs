using GalaSoft.MvvmLight;
using GalaSoft.MvvmLight.Command;
using Microsoft.TeamFoundation.Client;
using Microsoft.TeamFoundation.Core.WebApi;
using Microsoft.TeamFoundation.WorkItemTracking.Client;
using Microsoft.TeamFoundation.WorkItemTracking.WebApi;
using Microsoft.TeamFoundation.WorkItemTracking.WebApi.Models;
using Microsoft.VisualStudio.Services.Client;
using Microsoft.VisualStudio.Services.Common;
using Microsoft.VisualStudio.Services.WebApi;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Input;
using WordExporter.Core;
using WordExporter.Core.Templates;
using WordExporter.Core.WordManipulation;

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
            if (IsInDesignMode)
            {
                // Code runs in Blend --> create design time data.
            }
            else
            {
                // Code runs "for real"
            }

            //TemplateFolder = Path.Combine(
            //    Path.GetDirectoryName(AppDomain.CurrentDomain.BaseDirectory),
            //    "Templates");

            TemplateFolder = @"C:\develop\GitHub\AzureDevopsWordPlayground\src\WordExporter\Templates";
            Connect = new RelayCommand(ConnectMethod);
            GetQueries = new RelayCommand(GetQueriesMethod);
            Export = new RelayCommand(ExportMethod);
            _address = "https://dev.azure.com/gianmariaricci";
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
                    Templates.AddRange(TemplateManager.GetTemplateNames());
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

        private ObservableCollection<String> _templates = new ObservableCollection<String>();

        public ObservableCollection<String> Templates
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

        public String _selectedTemplate;

        public String SelectedTemplate
        {
            get
            {
                return _selectedTemplate;
            }
            set
            {
                Set<String>(() => this.SelectedTemplate, ref _selectedTemplate, value);
            }
        }

        public ICommand Connect { get; private set; }

        public ICommand GetQueries { get; private set; }

        public ICommand Export { get; private set; }

        private async void ConnectMethod()
        {
            Status = "Connecting";
            var connectionManager = new ConnectionManager();
            await connectionManager.ConnectAsync(Address);

            ProjectCollectionHttpClient projectCollectionHttpClient = connectionManager.GetClient<ProjectCollectionHttpClient>();
            Status = "Connected, Retrieving project collection";

            await GetCollecitonAsync(projectCollectionHttpClient);

            Status = "Connected, Retrieving Team Projects";
            var projectHttpClient = connectionManager.GetClient<ProjectHttpClient>();

            await GetTeamProjectAsync(projectHttpClient);
            Status = "Connected, List of team Project retrieved";
            Connected = true;

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

        private Task GetCollecitonAsync(ProjectCollectionHttpClient projectCollectionHttpClient)
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
            var queries =await witClient.GetQueriesAsync(SelectedTeamProject.Name, depth: 2, expand: QueryExpand.Wiql);
            Queries.Clear();
            await PopulateQueries(String.Empty, witClient, queries);
        }

        public void ExportMethod()
        {
            if (TemplateManager == null)
                return;

            if (String.IsNullOrEmpty(SelectedTemplate))
                return;

            var selected = SelectedQuery?.Results?.Where(q => q.Selected).ToList();
            if (selected == null || selected.Count == 0)
                return;

            var template = TemplateManager.GetWordDefinitionTemplate(SelectedTemplate);
            var fileName = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString()) + ".docx";
            using (WordManipulator manipulator = new WordManipulator(fileName, true))
            {
                foreach (var workItemResult in selected)
                {
                    var workItem = workItemResult.WorkItem;
                    manipulator.InsertWorkItem(workItem, template.GetTemplateFor(workItem.Type.Name), true);
                }
            }

            System.Diagnostics.Process.Start(fileName);
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
    }
}