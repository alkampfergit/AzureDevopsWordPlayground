using GalaSoft.MvvmLight;
using GalaSoft.MvvmLight.Command;
using Microsoft.TeamFoundation.WorkItemTracking.WebApi.Models;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Windows.Input;
using WordExporter.Core;

namespace WordExporter.UI.ViewModel
{
    public class QueryViewModel : ViewModelBase
    {
        public QueryViewModel(
            MainViewModel mainViewModel,
            String parentPath,
            QueryHierarchyItem queryHierarchyItem)
        {
            Query = queryHierarchyItem;
            FullPath = (parentPath + '/' + queryHierarchyItem.Name).Trim('/');
            Execute = new RelayCommand(ExecuteMethod);
            SelectAll = new RelayCommand(ExecuteSelectAll);
            _mainViewModel = mainViewModel;
        }

        private QueryHierarchyItem _query;

        public QueryHierarchyItem Query
        {
            get
            {
                return _query;
            }
            set
            {
                Set<QueryHierarchyItem>(() => this.Query, ref _query, value);
            }
        }

        private String _fullPath;

        public String FullPath
        {
            get
            {
                return _fullPath;
            }
            set
            {
                Set<String>(() => this.FullPath, ref _fullPath, value);
            }
        }

        private Boolean _selectAllSelected;

        public Boolean SelectAllSelected
        {
            get
            {
                return _selectAllSelected;
            }
            set
            {
                Set<Boolean>(() => this.SelectAllSelected, ref _selectAllSelected, value);
            }
        }

        private ObservableCollection<QueryResultViewModel> _results = new ObservableCollection<QueryResultViewModel>();
        private readonly MainViewModel _mainViewModel;

        public ObservableCollection<QueryResultViewModel> Results
        {
            get
            {
                return _results;
            }
            set
            {
                _results = value;
                RaisePropertyChanged(nameof(Results));
            }
        }

        public ICommand Execute { get; private set; }

        private async void ExecuteMethod()
        {
            if (_query.QueryType == QueryType.Flat)
            {
                ////This is the REST API
                //WorkItemTrackingHttpClient witClient = ConnectionManager.Instance.GetClient<WorkItemTrackingHttpClient>();
                //var result = await witClient.QueryByIdAsync(_query.Id);
                //foreach (var wiReference in result.WorkItems)
                //{
                //    var qrvm = new QueryResultViewModel(wiReference);
                //    Results.Add(qrvm);
                //}
                Results.Clear();
                Dictionary<String, String> parameters = new Dictionary<String, String>();
                parameters.Add("project", _mainViewModel.SelectedTeamProject.Name);
                var queryResult = ConnectionManager.Instance.WorkItemStore.Query(_query.Wiql, parameters);
                foreach (Microsoft.TeamFoundation.WorkItemTracking.Client.WorkItem workItem in queryResult)
                {
                    Results.Add(new QueryResultViewModel(workItem));
                }
            }
            else
            {
                //TODO: Show some meaningful error to the caller.
            }
        }

        public ICommand SelectAll { get; private set; }

        private void ExecuteSelectAll()
        {
            foreach (var qrvi in Results)
            {
                qrvi.Selected = SelectAllSelected;
            }
        }
    }
}
