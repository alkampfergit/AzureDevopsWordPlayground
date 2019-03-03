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
