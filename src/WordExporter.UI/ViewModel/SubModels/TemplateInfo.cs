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
                Parameters = wordTemplateFolderManager.TemplateDefinition?.ParameterSection.Parameters ?? new Dictionary<string, string>();
                ArrayParameters = wordTemplateFolderManager.TemplateDefinition.ArrayParameterSection?.ArrayParameters ?? new Dictionary<String, String>();
            }
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

        private Dictionary<String, String> _parameters;

        public Dictionary<String, String> Parameters
        {
            get
            {
                return _parameters;
            }
            set
            {
                Set<Dictionary<String, String>>(() => this.Parameters, ref _parameters, value);
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
    }
}
