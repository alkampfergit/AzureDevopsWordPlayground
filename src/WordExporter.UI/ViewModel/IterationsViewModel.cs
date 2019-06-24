using GalaSoft.MvvmLight;
using System;
using static WordExporter.Core.WorkItems.IterationManager;

namespace WordExporter.UI.ViewModel
{
    public class IterationsViewModel : ViewModelBase
    {
        private readonly IterationInfo _iteration;

        public IterationsViewModel(IterationInfo iteration)
        {
            _iteration = iteration;
            Path = iteration.Path;
            if (DateTime.TryParse(iteration.StartDate, out DateTime startDate))
            {
                StartDate = startDate;
            }
            if (DateTime.TryParse(iteration.EndDate, out DateTime endDate))
            {
                EndDate = endDate;
            } 
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
