using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WordExporter.Core.Templates
{
    /// <summary>
    /// This class manage a single folder for word template it scans all
    /// the docx files. Files should have name of the type of work item
    /// they are planned to export, and a WorkItem.docx that is the fallback
    /// in case specific file for work item is not present.
    /// </summary>
    public class WordTemplate
    {
        public WordTemplate(String templateFolder)
        {
            _templateFolder = templateFolder ?? throw new ArgumentNullException(nameof(templateFolder));
            _templateFileNames = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            var dinfo = new DirectoryInfo(templateFolder);

            Name = dinfo.Name;
            ScanFolder();
        }

        private readonly String _templateFolder;
        private readonly Dictionary<String, String> _templateFileNames;

        public String Name { get; private set; }

        private void ScanFolder()
        {
            var files = Directory.EnumerateFiles(_templateFolder, "*.docx");
            foreach (var file in files)
            {
                var finfo = new FileInfo(file);
                _templateFileNames.Add(Path.GetFileNameWithoutExtension(file), finfo.FullName);
            }
        }

        public String GetTemplateFor(string workItemType)
        {
            if (workItemType == null)
                throw new ArgumentNullException(nameof(workItemType));

            if (!_templateFileNames.TryGetValue(workItemType, out var templateFile))
            {
                return _templateFileNames["WorkItem"];
            }
            return templateFile;
        }
    }
}
