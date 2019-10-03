using GalaSoft.MvvmLight;
using System;
using System.Collections.Generic;
using System.Linq;
using WordExporter.Core.Templates;

namespace WordExporter.UI.ViewModel.SubModels
{
    public class TemplateInfo : ViewModelBase
    {
        public WordTemplateFolderManager WordTemplateFolderManager { get; private set; }

        public TemplateInfo(String templateName, WordTemplateFolderManager wordTemplateFolderManager)
        {
            IsScriptTemplate = wordTemplateFolderManager.HasTemplateDefinition;
            TemplateName = templateName;
            WordTemplateFolderManager = wordTemplateFolderManager;
            if (IsScriptTemplate)
            {
                //copy list of parameters
                var tdef = wordTemplateFolderManager.TemplateDefinition;
                Parameters =
                    tdef?.ParameterSection.Parameters
                        .Select(e => CreateParameterViewModel(tdef, e))
                        .ToList()
                    ?? new List<ParameterViewModel>();
                ArrayParameters = wordTemplateFolderManager.TemplateDefinition.ArrayParameterSection?.ArrayParameters ?? new Dictionary<String, String>();
            }
        }

        private ParameterViewModel CreateParameterViewModel(TemplateDefinition tdef, KeyValuePair<string, string> e)
        {
            String[] allowedValues = null;
            String type = "";
            if (tdef?.ParameterDefinition != null
                && tdef.ParameterDefinition.TryGetValue(e.Key, out var def))
            {
                allowedValues = def.AllowedValues;
                type = def.Type;
            }
            return new ParameterViewModel(e.Key, type, e.Value, allowedValues);
        }

        private Boolean _isScriptTemplate;

        public Boolean IsScriptTemplate
        {
            get
            {
                return _isScriptTemplate;
            }
            set
            {
                Set<Boolean>(() => this.IsScriptTemplate, ref _isScriptTemplate, value);
            }
        }

        private String _templateName;

        public String TemplateName
        {
            get
            {
                return _templateName;
            }
            set
            {
                Set<String>(() => this.TemplateName, ref _templateName, value);
            }
        }

        private List<ParameterViewModel> _parameters;

        public List<ParameterViewModel> Parameters
        {
            get
            {
                return _parameters;
            }
            set
            {
                Set<List<ParameterViewModel>>(() => this.Parameters, ref _parameters, value);
            }
        }

        private Dictionary<String, String> _arrayParameters;

        public Dictionary<String, String> ArrayParameters
        {
            get
            {
                return _arrayParameters;
            }
            set
            {
                Set<Dictionary<String, String>>(() => this.ArrayParameters, ref _arrayParameters, value);
            }
        }

        private Boolean _isSelected;
        public Boolean IsSelected
        {
            get
            {
                return _isSelected;
            }
            set
            {
                Set<Boolean>(() => this.IsSelected, ref _isSelected, value);
            }
        }
    }
}
