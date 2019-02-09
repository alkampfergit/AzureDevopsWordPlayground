using GalaSoft.MvvmLight;
using Microsoft.TeamFoundation.WorkItemTracking.WebApi.Models;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WordExporter.UI.ViewModel
{
    public class QueryViewModel : ViewModelBase
    {
        public QueryViewModel(String parentPath, QueryHierarchyItem queryHierarchyItem)
        {
            Query = queryHierarchyItem;
            FullPath = (parentPath + '/' + queryHierarchyItem.Name).Trim('/');
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
    }
}
