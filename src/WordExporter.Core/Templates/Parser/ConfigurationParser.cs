using Sprache;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WordExporter.Core.Templates.Parser
{
    public class ConfigurationParser
    {
        /// <summary>
        /// Keyvalue is such a typical construct that is defined in the configuration parser
        /// TODO: refactor to a common parser utility
        /// </summary>
        public static Parser<KeyValue> KeyValue =
        (
            from key in Parse.CharExcept(':').Many().Text()
            from separator in Parse.Char(':')
            from value in Parse.CharExcept('\n').Many().Text()
            from eol in Parse.LineTerminator
            select new KeyValue(key.Trim(), value.Trim(' ', '\"', '\n', '\t', '\r'))
        ).Named("keyvalue");

        /// <summary>
        /// Keyvalue is such a typical construct that is defined in the configuration parser
        /// TODO: refactor to a common parser utility
        /// </summary>
        public static Parser<KeyValue> MultiLineKeyValue =
        (
            from key in Parse.CharExcept(':').Many().Text()
            from separator in Parse.Char(':')
            from trailingWs in Parse.WhiteSpace.Optional().Many()
            from openValue in Parse.Char('\"')
            from value in Parse.CharExcept('\"').Many().Text()
            from endValue in Parse.Char('\"')
            from eol in Parse.LineTerminator
            select new KeyValue(key.Trim(), value.Trim(' ', '\"', '\n', '\t', '\r'))
        ).Named("keyvalue");

        public static Parser<IEnumerable<KeyValue>> KeyValueList =
             from keyValue in KeyValue.Or(MultiLineKeyValue).Many()
             select keyValue;

        public static readonly Parser<Section> SectionParser =
        (
            from open in Parse.String("[[")
            from section in Parse.AnyChar.Except(Parse.Char(']')).Many().Text()
            from close in Parse.String("]]")
            from trailingWhitespace in Parse.WhiteSpace.Many()
            from restOfSection in Parse.AnyChar.Except(Parse.String("[[")).Many().Text()
            select Section.Create(section, restOfSection)
        ).Named("section");

        internal static Parser<TemplateDefinition> TemplateDefinition =
             from sections in SectionParser.AtLeastOnce()
             select new TemplateDefinition(sections);

        /// <summary>
        /// This is the entry point that will be called to parse an entire definition.
        /// </summary>
        /// <param name="fullTemplateContent"></param>
        /// <returns></returns>
        public TemplateDefinition ParseTemplateDefinition(string fullTemplateContent)
        {
            return TemplateDefinition.Parse(fullTemplateContent);
        }
    }
}
