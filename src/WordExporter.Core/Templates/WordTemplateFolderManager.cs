using Sprache;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WordExporter.Core.Templates.Parser;

namespace WordExporter.Core.Templates
{
    /// <summary>
    /// This class manage a single folder for word template it scans all
    /// the docx files. Files should have name of the type of work item
    /// they are planned to export, and a WorkItem.docx that is the fallback
    /// in case specific file for work item is not present.
    /// </summary>
    public class WordTemplateFolderManager
    {
        public WordTemplateFolderManager(String templateFolder)
        {
            if (templateFolder == null)
                throw new ArgumentNullException(nameof(templateFolder));

            _templateFileNames = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            var dinfo = new DirectoryInfo(templateFolder);
            _templateFolder = dinfo.FullName;
            Name = dinfo.Name;
            ScanFolder();
        }

        private readonly String _templateFolder;
        private readonly Dictionary<String, String> _templateFileNames;

        public String Name { get; private set; }

        public TemplateDefinition TemplateDefinition { get; private set; }

        public Boolean HasTemplateDefinition => TemplateDefinition != null;

        private void ScanFolder()
        {
            var files = Directory.EnumerateFiles(_templateFolder, "*.docx");
            foreach (var file in files)
            {
                var finfo = new FileInfo(file);
                _templateFileNames.Add(Path.GetFileNameWithoutExtension(file), finfo.FullName);
            }
            var structureFile = Path.Combine(_templateFolder, "structure.txt");
            if (File.Exists(structureFile))
            {
                TemplateDefinition = ConfigurationParser.TemplateDefinition.Parse(File.ReadAllText(structureFile));
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

        internal string GenerateFullFileName(string templateName)
        {
            if (!Path.IsPathRooted(templateName))
            {
                return Path.Combine(_templateFolder, templateName);
            }
            return templateName;
        }

        /// <summary>
        /// Retrieve a table file name, if <paramref name="createTempVersion"/> is true it 
        /// will create a temp file name and then copy the original file to avoid messing
        /// up with the original template file.
        /// </summary>
        /// <param name="tableName"></param>
        /// <param name="createTempVersion">If false it will return the original path of 
        /// docx file in the template folder, pay attention because if you are going to 
        /// modifiy it, it will be modified forever.</param>
        /// <returns></returns>
        public string GetTable(string tableName, Boolean createTempVersion)
        {
            String baseFile = Path.Combine(_templateFolder, "Table" + tableName + ".docx");
            if (!File.Exists(baseFile))
                throw new ArgumentException($"There is no table file for table name {tableName}");

            if (!createTempVersion)
                return baseFile;

            String tempFile = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString() + ".docx");
            File.Copy(baseFile, tempFile);
            return tempFile;
        }

        public String CopyFileInTempDirectory(String localFileName)
        {
            String fileName = Path.Combine(_templateFolder, localFileName);
            if (!File.Exists(fileName))
                throw new ArgumentException($"File {nameof(fileName)} does not exists.", nameof(fileName));

            var tempFile = Path.GetTempPath() + Guid.NewGuid().ToString() + Path.GetExtension(fileName);
            File.Copy(fileName, tempFile);
            return tempFile;
        }
    }
}
