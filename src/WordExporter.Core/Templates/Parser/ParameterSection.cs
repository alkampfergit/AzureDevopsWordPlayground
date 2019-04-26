using Sprache;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WordExporter.Core.WordManipulation;

namespace WordExporter.Core.Templates.Parser
{
    public sealed class ParameterSection : Section
    {
        private ParameterSection()
        {
        }

        /// <summary>
        /// Dictionary of parameters and default values
        /// </summary>
        public Dictionary<String, String> Parameters { get; private set; }

        #region syntax

        /// <summary>
        /// Keyvalue is such a typical construct that is defined in the configuration parser
        /// TODO: refactor to a common parser utility
        /// </summary>
        public readonly static Parser<String> ParameterName =
        (
            from parameterName in Parse.CharExcept(new[] { '=', '\r', '\n' }).Many().Text()
            from trailingWs in Parse.WhiteSpace.Optional().Many()
            select parameterName
        ).Named("paramName");

        public readonly static Parser<String> DefaultValue =
        (
            from equal in Parse.Char('=')
            from trailingWs in Parse.WhiteSpace.Optional().Many()
            from parameterDefaultValue in Parse.AnyChar.Except(Parse.LineEnd).Many().Text()
            select parameterDefaultValue
        ).Named("paramValue");

        public readonly static Parser<KeyValuePair<String, String>> Parameter =
        (
            from parameterName in ParameterName
            from defaultValue in DefaultValue.Optional()
            from lineEnd in Parse.LineEnd.Optional()
            select new KeyValuePair<String, String>(parameterName.Trim(), defaultValue.GetOrElse("").Trim(' ', '\"', '\n', '\t', '\r'))
        ).Named("param");

        public readonly static Parser<ParameterSection> Parser =
             from parameters in Parameter.Many()
             select new ParameterSection()
             {
                 Parameters = parameters.ToDictionary(p => p.Key, p => p.Value)
             };

        #endregion
    }
}
