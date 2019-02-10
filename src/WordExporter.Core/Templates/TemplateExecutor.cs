using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
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

        public TemplateExecutor(WordTemplateFolderManager  wordTemplateFolderManager)
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
        /// <param name="fileName"></param>
        /// <param name="connectionManager"></param>
        /// <param name="parameters">Parameters should be given from external code
        /// because this class does not know how to ask parameter to the user.</param>
        /// <param name="teamProjectName"></param>
        public void GenerateWordFile(
            String fileName,
            ConnectionManager connectionManager,
            String teamProjectName,
            Dictionary<string, Object> parameters)
        {
            if (_wordTemplateFolderManager.TemplateDefinition == null)
                throw new ApplicationException("Cannot generate work file, template folder name does not contain valid structure.txt file");

            //Ok, we start creating the empty file, then proceed to create everything.
            using (WordManipulator manipulator = new WordManipulator(fileName, true))
            {
                //now we need to scan the sections of the definition so we can use
                //each section to build the file
                foreach (var section in _wordTemplateFolderManager.TemplateDefinition.AllSections)
                {
                    //now each section can do something with my standard word manipulator
                    section.Assemble(manipulator, parameters, connectionManager, _wordTemplateFolderManager, teamProjectName);
                }
            }
        }
    }
}
