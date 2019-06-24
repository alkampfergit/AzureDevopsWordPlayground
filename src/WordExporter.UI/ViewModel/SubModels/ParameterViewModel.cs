using GalaSoft.MvvmLight;
using System;

namespace WordExporter.UI.ViewModel.SubModels
{
    public class ParameterViewModel : ViewModelBase
    {
        /// <summary>
        /// Create a parameter view model
        /// </summary>
        /// <param name="name"></param>
        /// <param name="type">CAn be null or empty string, it will considered to be 
        /// string.</param>
        /// <param name="initialValue"></param>
        /// <param name="allowedValues"></param>
        public ParameterViewModel(
            String name,
            String type,
            String initialValue,
            String[] allowedValues)
        {
            Name = name;
            Type = String.IsNullOrEmpty( type) ? "string" : type;
            Value = initialValue;
            AllowedValues = allowedValues;
            HasAllowedValues = allowedValues?.Length > 0;
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

        private String _type;

        public String Type
        {
            get
            {
                return _type;
            }
            set
            {
                Set<String>(() => this.Type, ref _type, value);
            }
        }

        private Boolean _hasAllowedValues;

        public Boolean HasAllowedValues
        {
            get
            {
                return _hasAllowedValues;
            }
            set
            {
                Set<Boolean>(() => this.HasAllowedValues, ref _hasAllowedValues, value);
            }
        }

        private String[] _allowedValues;

        public String[] AllowedValues
        {
            get
            {
                return _allowedValues;
            }
            set
            {
                Set<String[]>(() => this.AllowedValues, ref _allowedValues, value);
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
