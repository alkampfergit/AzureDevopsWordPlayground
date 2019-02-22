using GalaSoft.MvvmLight;
using Microsoft.TeamFoundation.WorkItemTracking.WebApi.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WordExporter.UI.ViewModel
{
    public class QueryResultViewModel : ViewModelBase
    {
        private readonly Microsoft.TeamFoundation.WorkItemTracking.Client.WorkItem _workItem;

        public QueryResultViewModel(Microsoft.TeamFoundation.WorkItemTracking.Client.WorkItem workItem)
        {
            this._workItem = workItem;
            Id = _workItem.Id;
            Title = _workItem.Title;
        }

        public Microsoft.TeamFoundation.WorkItemTracking.Client.WorkItem WorkItem => _workItem;

        private Boolean _selected;

        public Boolean Selected
        {
            get
            {
                return _selected;
            }
            set
            {
                Set<Boolean>(() => this.Selected, ref _selected, value);
            }
        }

        private Int32 _id;

        public Int32 Id
        {
            get
            {
                return _id;
            }
            set
            {
                Set<Int32>(() => this.Id, ref _id, value);
            }
        }

        private String _title;

        public String Title
        {
            get
            {
                return _title;
            }
            set
            {
                Set<String>(() => this.Title, ref _title, value);
            }
        }
    }
}
