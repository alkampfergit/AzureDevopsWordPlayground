using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using WordExporter.Core.ExcelManipulation;
using WordExporter.Core.WordManipulation;

namespace WordExporter.Core.Templates
{
    /// <summary>
    /// This is a class that can execute a <see cref="TemplateDefinition"/>
    /// to create a document.
    /// </summary>
    public class TemplateExecutor
    {
        private readonly WordTemplateFolderManager _wordTemplateFolderManager;

        public TemplateExecutor(WordTemplateFolderManager wordTemplateFolderManager)
        {
            _wordTemplateFolderManager = wordTemplateFolderManager;
        }

        /// <summary>
        /// <para>
        /// This will generate word file directly from the folder, we just 
        /// need to know the connection and the team project we want to 
        /// use.
        /// </para>
        /// <para>
        /// This can be used only if the folder with the template have a 
        /// text file to parse a <see cref="TemplateDefinition"/>
        /// </para>
        /// </summary>
        /// <param name="fileNameWithoutExtension"></param>
        /// <param name="connectionManager"></param>
        /// <param name="teamProjectName"></param>
        /// <param name="parameters">Parameters should be given from external code
        /// because this class does not know how to ask parameter to the user.</param>
        /// <returns>File name genreated</returns>
        public String GenerateFile(
            String fileNameWithoutExtension,
            ConnectionManager connectionManager,
            String teamProjectName,
            Dictionary<string, Object> parameters)
        {
            var fileName = fileNameWithoutExtension;

            if (_wordTemplateFolderManager.TemplateDefinition == null)
            {
                throw new ApplicationException("Cannot generate work file, template folder name does not contain valid structure.txt file");
            }

            TemplateDefinition templateDefinition = _wordTemplateFolderManager.TemplateDefinition;
            if (templateDefinition.Type == TemplateType.Word)
            {
                fileName += ".docx";
                Boolean createNew = true;
                if (templateDefinition.BaseTemplate != null)
                {
                    _wordTemplateFolderManager.CopyFileToDestination(templateDefinition.BaseTemplate, fileName);
                    File.Copy(templateDefinition.BaseTemplate, fileName);
                    createNew = false;
                }
                //Ok, we start creating the empty file, then proceed to create everything.
                using (WordManipulator manipulator = new WordManipulator(fileName, createNew))
                {
                    //now we need to scan the sections of the definition so we can use
                    //each section to build the file
                    foreach (var section in templateDefinition.AllSections)
                    {
                        //now each section can do something with my standard word manipulator
                        section.Assemble(manipulator, parameters, connectionManager, _wordTemplateFolderManager, teamProjectName);
                    }
                }
            }
            else if (templateDefinition.Type == TemplateType.Excel)
            {
                fileName += ".xlsx";
                _wordTemplateFolderManager.CopyFileToDestination(templateDefinition.BaseTemplate, fileName);
                //Ok, we start creating the empty file, then proceed to create everything.
                using (ExcelManipulator manipulator = new ExcelManipulator(fileName))
                {
                    //now we need to scan the sections of the definition so we can use
                    //each section to build the file
                    foreach (var section in templateDefinition.AllSections)
                    {
                        //now each section can do something with my standard word manipulator
                        section.AssembleExcel(manipulator, parameters, connectionManager, _wordTemplateFolderManager, teamProjectName);
                    }
                }
            }
            else
            {
                throw new NotSupportedException($"Unsupported type {templateDefinition.Type}");
            }

            return fileName;
        }

        public void DumpWorkItem(
            String fileName,
            ConnectionManager connectionManager,
            String teamProjectName,
            Dictionary<string, Object> parameters)
        {
            if (_wordTemplateFolderManager.TemplateDefinition == null)
            {
                throw new ApplicationException("Cannot generate work file, template folder name does not contain valid structure.txt file");
            }

            StringBuilder sb = new StringBuilder();
            //Scan each section to get all work item if there are any.
            foreach (var section in _wordTemplateFolderManager.TemplateDefinition.AllSections)
            {
                //now each section can do something with my standard word manipulator
                section.Dump(sb, parameters, connectionManager, _wordTemplateFolderManager, teamProjectName);
            }
            File.WriteAllText(fileName, sb.ToString());
        }
    }
}
