using GalaSoft.MvvmLight;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WordExporter.UI.ViewModel
{
    public class ParameterViewModel : ViewModelBase
    {
        public ParameterViewModel(String name)
        {
            Name = name;
        }

        private String _name;

        public String Name
        {
            get
            {
                return _name;
            }
            set
            {
                Set<String>(() => this.Name, ref _name, value);
            }
        }

        private String _value;

        public String Value
        {
            get
            {
                return _value;
            }
            set
            {
                Set<String>(() => this.Value, ref _value, value);
            }
        }
    }
}
