using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WordExporter.Core.Templates
{
    public class TemplateManager
    {
        public TemplateManager(String baseTemplateFolder)
        {
            _baseTemplateFolder = baseTemplateFolder ?? throw new ArgumentNullException(nameof(baseTemplateFolder));
            _templates = new Dictionary<string, WordTemplate>(StringComparer.OrdinalIgnoreCase);

            if (!Directory.Exists(baseTemplateFolder))
                throw new ArgumentException($"Template folder {baseTemplateFolder} does not exists.", nameof(baseTemplateFolder));

            ScanFortemplate();
        }

        private readonly String _baseTemplateFolder;
        private readonly Dictionary<String, WordTemplate> _templates;

        public Int32 TemplateCount => _templates.Count;

        public WordTemplate GetWordTemplate(String templateName)
        {
            if (!_templates.TryGetValue(templateName, out var template))
            {
                throw new ArgumentException($"There is no template with name {templateName} in folder {_baseTemplateFolder}");
            }

            return template;
        }

        public IEnumerable<String> GetTemplateNames()
        {
            return _templates.Keys;
        }

        /// <summary>
        /// Scan a base folder for all subfolder for template. Templates are simply
        /// docx files inside a folder.
        /// </summary>
        private void ScanFortemplate()
        {
            var dinfo = new DirectoryInfo(_baseTemplateFolder);
            foreach (var directory in dinfo.EnumerateDirectories())
            {
                _templates.Add(directory.Name, new WordTemplate(directory.FullName));
            }
        }
    }
}
