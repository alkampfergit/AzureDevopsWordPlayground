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
        public readonly static Parser<KeyValuePair<String, String>> Parameter =
        (
            from parameterName in Parse.CharExcept('=').Many().Text()
            from separator in Parse.Char('=')
            from trailingWs in Parse.WhiteSpace.Optional().Many()
            from parameterDefaultValue in Parse.AnyChar.Except(Parse.LineEnd).Many().Text()
            select new KeyValuePair<String, String>(parameterName.Trim(), parameterDefaultValue.Trim(' ', '\"', '\n', '\t', '\r'))
        ).Named("paramValue");

        public readonly static Parser<ParameterSection> Parser =
             from parameters in Parameter.Many()
             select new ParameterSection()
             {
                 Parameters = parameters.ToDictionary(p => p.Key, p => p.Value)
             };

        #endregion
    }
}
