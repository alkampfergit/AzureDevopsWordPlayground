using Microsoft.TeamFoundation.Client;
using Microsoft.TeamFoundation.WorkItemTracking.Client;
using Microsoft.VisualStudio.Services.Client;
using Microsoft.VisualStudio.Services.Common;
using System;
using System.Collections.Generic;
using System.Linq;

namespace WordExporter.Core
{
    public class Connection
    {
        /// <summary>
        /// Perform a connection with an access token, simplest way to give permission to a program
        /// to access your account.
        /// </summary>
        /// <param name="accessToken"></param>
        public Connection(String accountUri, String accessToken)
        {
            ConnectToTfs(accountUri, accessToken);
            _workItemStore = _tfsCollection.GetService<WorkItemStore>();
        }

        private TfsTeamProjectCollection _tfsCollection;
        private WorkItemStore _workItemStore;

        public WorkItemStore WorkItemStore => _workItemStore;

        private bool ConnectToTfs(String accountUri, String accessToken)
        {
            //login for VSTS
            VssCredentials creds = new VssBasicCredential(
                String.Empty,
                accessToken);
            creds.Storage = new VssClientCredentialStorage();

            // Connect to VSTS
            _tfsCollection = new TfsTeamProjectCollection(new Uri(accountUri), creds);
            _tfsCollection.Authenticate();
            return true;
        }

        /// <summary>
        /// Returns a list of all team projects names.
        /// </summary>
        /// <returns></returns>
        public IEnumerable<String> GetTeamProjectsNames()
        {
            return _workItemStore.Projects.OfType<Project>().Select(_ => _.Name);
        }
    }
}
