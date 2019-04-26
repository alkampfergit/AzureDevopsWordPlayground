using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WordExporter.Core.Templates.Parser;

namespace WordExporter.Core.Templates
{
    /// <summary>
    /// <para>
    /// Definition of a complex template, composed by
    /// multiple sections, parameters, etc.
    /// </para>
    /// <para>
    /// This class is usually created parsing a text file
    /// using syntax defined in <see cref="ConfigurationParser"/>
    /// </para>
    /// </summary>
    public class TemplateDefinition
    {
        public TemplateDefinition(IEnumerable<Section> sections)
        {
            ParameterSection = sections.OfType<ParameterSection>().SingleOrDefault();
            ArrayParameterSection = sections.OfType<ArrayParameterSection>().SingleOrDefault();
            AllSections = sections.ToArray();
        }

        /// <summary>
        /// List of all the sections that composes the template definition
        /// </summary>
        public Section[] AllSections { get; private set; }

        /// <summary>
        /// These are the parameters that the user can specify
        /// and that can be referred inside word template
        /// with standard sytax {{parameter}}
        /// </summary>
        public ParameterSection ParameterSection { get; internal set; }

        /// <summary>
        /// Array Parameters are special parameters, the user can specify multiple values
        /// comma or semicolon separated, and the template will be executed multiple times
        /// once for each array parameter instance.
        /// </summary>
        public ArrayParameterSection ArrayParameterSection { get; internal set; }
    }
}
