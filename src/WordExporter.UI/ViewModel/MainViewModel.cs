using GalaSoft.MvvmLight;
using GalaSoft.MvvmLight.Command;
using Microsoft.TeamFoundation.Client;
using Microsoft.TeamFoundation.WorkItemTracking.WebApi;
using Microsoft.VisualStudio.Services.Client;
using Microsoft.VisualStudio.Services.Common;
using Microsoft.VisualStudio.Services.WebApi;
using System;
using System.Net;
using System.Windows.Input;

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

        public ICommand Connect { get; private set; }

        private void ConnectMethod()
        {
// Create instance of VssConnection using Visual Studio sign-in prompt
VssConnection connection = new VssConnection(new Uri(Address), new VssClientCredentials());
WorkItemTrackingHttpClient witClient = connection.GetClient<WorkItemTrackingHttpClient>();
var items = witClient.GetQueriesAsync("TestMigration").Result;

            NetworkCredential netCred = new NetworkCredential("", "");
            BasicAuthCredential basicCred = new BasicAuthCredential(netCred);
            TfsClientCredentials tfsCred = new TfsClientCredentials(basicCred);
            tfsCred.AllowInteractive = true;

            TfsTeamProjectCollection collection = new TfsTeamProjectCollection(new Uri(Address), tfsCred);
            collection.Authenticate();
            Connected = true;
        }
    }
}