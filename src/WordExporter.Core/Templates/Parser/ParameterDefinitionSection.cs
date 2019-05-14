using Sprache;
using System;
using System.Collections.Generic;
using System.Linq;

namespace WordExporter.Core.Templates.Parser
{
    public class ParameterDefinitionSection : Section
    {
        /// <summary>
        /// Dictionary of parameters and default values
        /// </summary>
        public Dictionary<String, ParameterDefinition> Parameters { get; private set; }

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

        public readonly static Parser<ParameterDefinition> ParameterDefinition =
        (
            from equal in Parse.Char('=')
            from trailingWs in Parse.WhiteSpace.Optional().Many()
            from parameterType in Parse.LetterOrDigit.Many().Text()
            from separator in Parse.Char('/').Optional()
            from allowedValues in Parse.AnyChar.Except(Parse.LineEnd).Many().Text()
            select new ParameterDefinition()
            {
                Type = parameterType,
                AllowedValues = allowedValues?.Split('|').ToArray() ?? Array.Empty<String>()
            }
        ).Named("paramDef");

        public readonly static Parser<KeyValuePair<String, ParameterDefinition>> Parameter =
        (
            from parameterName in ParameterName
            from parameterDefinition in ParameterDefinition.Once()
            from lineEnd in Parse.LineEnd.Optional()
            select new KeyValuePair<String, ParameterDefinition>(parameterName.Trim(), parameterDefinition.Single())
        ).Named("param");

        public readonly static Parser<ParameterDefinitionSection> Parser =
             from parameters in Parameter.Many()
             select new ParameterDefinitionSection()
             {
                 Parameters = parameters.ToDictionary(p => p.Key, p => p.Value)
             };

        #endregion
    }

    public class ParameterDefinition
    {
        public String Type { get; set; }

        public String[] AllowedValues { get; set; }
    }
}
