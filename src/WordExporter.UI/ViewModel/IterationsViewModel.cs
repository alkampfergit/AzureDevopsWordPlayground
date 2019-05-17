using GalaSoft.MvvmLight;
using Microsoft.TeamFoundation.Work.WebApi;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WordExporter.UI.ViewModel
{
    public class IterationsViewModel : ViewModelBase
    {
        private readonly TeamSettingsIteration _iteration;

        public IterationsViewModel(TeamSettingsIteration iteration)
        {
            _iteration = iteration;
            Path = iteration.Path;
            StartDate = iteration.Attributes?.StartDate;
            EndDate = iteration.Attributes?.FinishDate;
        }

        private String _path;

        public String Path
        {
            get
            {
                return _path;
            }
            set
            {
                Set<String>(() => this.Path, ref _path, value);
            }
        }

        private DateTime? _startDate;

        public DateTime? StartDate
        {
            get
            {
                return _startDate;
            }
            set
            {
                Set<DateTime?>(() => this.StartDate, ref _startDate, value);
            }
        }

        private DateTime? _endDate;

        public DateTime? EndDate
        {
            get
            {
                return _endDate;
            }
            set
            {
                Set<DateTime?>(() => this.EndDate, ref _endDate, value);
            }
        }

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
    }
}
