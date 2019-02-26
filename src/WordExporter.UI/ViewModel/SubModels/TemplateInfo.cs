using GalaSoft.MvvmLight;
using System;
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
    }
}
