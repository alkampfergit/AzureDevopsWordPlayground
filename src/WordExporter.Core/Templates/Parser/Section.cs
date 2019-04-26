using Sprache;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WordExporter.Core.WordManipulation;

namespace WordExporter.Core.Templates.Parser
{
    /// <summary>
    /// This is the abstract class represented by a single section.
    /// Any section can have own parser, this is done to simply syntax.
    /// </summary>
    public abstract class Section
    {
        public static Section Create(String sectionName, String sectionContent)
        {
            switch (sectionName.Trim())
            {
                case "parameters":
                    return ParameterSection.Parser.Parse(sectionContent);
                case "arrayParameters":
                    return ArrayParameterSection.Parser.Parse(sectionContent);
                case "static":
                    return StaticWordSection.Parser.Parse(sectionContent);
                case "query":
                    return QuerySection.Parser.Parse(sectionContent);
                default:
                    throw new NotSupportedException($"section name {sectionName} not supported");
            }
        }

        /// <summary>
        /// <para>Each section should be able to generate a part of the word file.</para>
        /// <para>
        /// Not all sections are used to manipulate word file, for those sections
        /// this function could be not overriden
        /// </para>
        /// </summary>
        /// <param name="manipulator"></param>
        /// <param name="parameters"></param>
        /// <param name="connectionManager"></param>
        /// <param name="teamProjectName"></param>
        public virtual void Assemble(
            WordManipulator manipulator,
            Dictionary<string, Object> parameters,
            ConnectionManager connectionManager,
            WordTemplateFolderManager wordTemplateFolderManager,
            string teamProjectName)
        {
            //Do nothing.
        }

        /// <summary>
        /// <para>Can simply dump information for work items, field etc. implementing
        /// this method is completely optional</para>
        /// </summary>
        /// <param name="stringBuilder">A string builder to accumulate all the
        /// data for dumping information</param>
        /// <param name="parameters"></param>
        /// <param name="connectionManager"></param>
        /// <param name="teamProjectName"></param>
        public virtual void Dump(
            StringBuilder stringBuilder,
            Dictionary<string, Object> parameters,
            ConnectionManager connectionManager,
            WordTemplateFolderManager wordTemplateFolderManager,
            string teamProjectName)
        {
            //Do nothing.
        }
    }
}
