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
        private readonly WorkItemReference wiReference;

        public QueryResultViewModel(WorkItemReference wiReference)
        {
            this.wiReference = wiReference;
            Id = wiReference.Id;
            Url = wiReference.Url;
        }

        private Int64 _id;

        public Int64 Id
        {
            get
            {
                return _id;
            }
            set
            {
                Set<Int64>(() => this.Id, ref _id, value);
            }
        }

        private String _url;

        public String Url
        {
            get
            {
                return _url;
            }
            set
            {
                Set<String>(() => this.Url, ref _url, value);
            }
        }
    }
}
