using Sprache;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WordExporter.Core.Support;
using WordExporter.Core.WordManipulation;

namespace WordExporter.Core.Templates.Parser
{
    public class StaticWordSection : Section
    {
        private StaticWordSection(IEnumerable<KeyValue> keyValuePairList)
        {
            FileName = keyValuePairList.GetStringValue("filename");
            if (FileName == null)
                throw new ArgumentException("Static word section needs a parameter called filename");
            PageBreak = keyValuePairList.GetBooleanValue("PageBreak"); 
        }

        public String FileName { get; private set; }

        public Boolean PageBreak { get; private set; }

        #region Syntax

        public readonly static Parser<StaticWordSection> Parser =
            from keyValueList in ConfigurationParser.KeyValueList
            select new StaticWordSection(keyValueList);

        #endregion

        public override void Assemble(
            WordManipulator manipulator,
            Dictionary<string, Object> parameters,
            ConnectionManager connectionManager,
            WordTemplateFolderManager wordTemplateFolderManager,
            string teamProjectName)
        {
            //ok we simply add a file to the manipulator, but remember we need to perform
            //substitution.
            var fileName = wordTemplateFolderManager.CopyFileInTempDirectory(FileName);
            using (WordManipulator m = new WordManipulator(fileName, false))
            {
                m.SubstituteTokens(parameters);
            }
            //Simply append other file honoring the page break.
            manipulator.AppendOtherWordFile(fileName, PageBreak);
            File.Delete(FileName);
            base.Assemble(manipulator, parameters, connectionManager, wordTemplateFolderManager, teamProjectName);
        }
    }
}
