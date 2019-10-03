using System;
using System.Collections.Generic;
using System.Linq;
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
            var pdef = sections.OfType<ParameterDefinitionSection>().SingleOrDefault();
            if (pdef != null)
            {
                ParameterDefinition = pdef.Parameters;
            }
            else
            {
                ParameterDefinition = new Dictionary<string, ParameterDefinition>();
            }

            //TODO: Using a visitor pattern could be a better solution.
            var definitionSection = sections.OfType<DefinitionSection>().SingleOrDefault();
            if (definitionSection != null)
            {
                if (definitionSection.Parameters.TryGetValue("type", out var paramType))
                {
                    if (Enum.TryParse<TemplateType>(paramType, true, out var templateType))
                    {
                        Type = templateType;
                    }
                }
                if (definitionSection.Parameters.TryGetValue("baseTemplate", out var baseTemplate))
                {
                    BaseTemplate = baseTemplate;
                }
            }
            ArrayParameterSection = sections.OfType<ArrayParameterSection>().SingleOrDefault();
            AllSections = sections.ToArray();
        }

        public TemplateType Type { get; private set; } = TemplateType.Word;

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
        /// Optional definition of parameters
        /// </summary>
        public Dictionary<String, ParameterDefinition> ParameterDefinition { get; internal set; }

        /// <summary>
        /// Array Parameters are special parameters, the user can specify multiple values
        /// comma or semicolon separated, and the template will be executed multiple times
        /// once for each array parameter instance.
        /// </summary>
        public ArrayParameterSection ArrayParameterSection { get; internal set; }

        /// <summary>
        /// Optionally this is the base template used to generate the report.
        /// </summary>
        public String BaseTemplate { get; set; }
    }

    /// <summary>
    /// This tool was born to export word templates, but actually we can use
    /// to export also into excel, if needed.
    /// </summary>
    public enum TemplateType
    {
        Word = 0,
        Excel = 1
    }
}
