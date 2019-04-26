using Sprache;
using System;
using System.Collections.Generic;
using System.Linq;

namespace WordExporter.Core.Templates.Parser
{
    /// <summary>
    /// Formally identical to <see cref="ParameterSection"/> but it is keep
    /// separated to allow for explicit syntax differentiation.
    /// </summary>
    public sealed class ArrayParameterSection : Section
    {
        private ArrayParameterSection()
        {
        }

        /// <summary>
        /// Dictionary of parameters and default values
        /// </summary>
        public Dictionary<String, String> ArrayParameters { get; private set; }

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

        public readonly static Parser<ArrayParameterSection> Parser =
             from parameters in Parameter.Many()
             select new ArrayParameterSection()
             {
                 ArrayParameters = parameters.ToDictionary(p => p.Key, p => p.Value)
             };

        #endregion
    }
}
