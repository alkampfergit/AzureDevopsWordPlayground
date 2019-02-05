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
using System.Threading.Tasks;
using System.Windows.Input;
using WordExporter.Core;

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

            Connect = new RelayCommand(ConnectMethod);
            GetQueries = new RelayCommand(GetQueriesMethod);

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

        private ObservableCollection<QueryHierarchyItem> _queries = new ObservableCollection<QueryHierarchyItem>();
        public ObservableCollection<QueryHierarchyItem> Queries
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

        public ICommand Connect { get; private set; }
        public ICommand GetQueries { get; private set; }

        private async void ConnectMethod()
        {
            var connectionManager = new ConnectionManager();
            await connectionManager.ConnectAsync(Address);

            ProjectCollectionHttpClient projectCollectionHttpClient = connectionManager.GetClient<ProjectCollectionHttpClient>();

            foreach (var projectCollectionReference in projectCollectionHttpClient.GetProjectCollections(10, 0).Result)
            {
                // retrieve a reference to the actual project collection based on its (reference) .Id
                var projectCollection = projectCollectionHttpClient.GetProjectCollection(projectCollectionReference.Id.ToString()).Result;

                // the 'web' Url is the one for the PC itself, the API endpoint one is different, see below
                var webUrlForProjectCollection = projectCollection.Links.Links["web"] as ReferenceLink;
            }

            var projectHttpClient = connectionManager.GetClient<ProjectHttpClient>();

            // then - same as above.. iterate over the project references (with a hard-coded pagination of the first 10 entries only)
            foreach (var projectReference in projectHttpClient.GetProjects(top: 100, skip: 0).Result)
            {
                // and then get ahold of the actual project
                var teamProject = projectHttpClient.GetProject(projectReference.Id.ToString()).Result;
                var urlForTeamProject = ((ReferenceLink)teamProject.Links.Links["web"]).Href;

                _teamProjects.Add(teamProject);
            }

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

        public async void GetQueriesMethod()
        {
            WorkItemTrackingHttpClient witClient = ConnectionManager.Instance.GetClient<WorkItemTrackingHttpClient>();
            var queries = witClient.GetQueriesAsync(SelectedTeamProject.Name, depth: 2).Result;
            Queries.Clear();
            await PopulateQueries(witClient, queries);
        }

        private async Task PopulateQueries(WorkItemTrackingHttpClient witClient, IEnumerable<QueryHierarchyItem> queries)
        {
            foreach (var query in queries)
            {
                if (query.IsFolder == true)
                {
                    Queries.Add(query);
                }
                if (query.HasChildren == true)
                {
                    if (query.Children == null)
                    {
                        //need to requery the store to grab reference to the query.
                        var queryReloaded = await witClient.GetQueryAsync(SelectedTeamProject.Id, query.Path, depth: 2);
                        await PopulateQueries(witClient, queryReloaded.Children);
                    }
                    else
                    {
                        await PopulateQueries(witClient, query.Children);
                    }
                }
            }
        }
    }
}