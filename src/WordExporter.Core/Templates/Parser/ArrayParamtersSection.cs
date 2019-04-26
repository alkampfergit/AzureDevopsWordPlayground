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
    /// Formally identical to <see cref="ParameterSection"/> but it is keep
    /// separated to allow for explicit syntax differentiation.
    /// </summary>
    public sealed class ArrayParameterSection : Section
    {
        private ArrayParameterSection()
        {
        }

        public List<String> ParameterNames { get; private set; }

        #region syntax

        private static readonly Parser<string> Line = Parse
            .AnyChar
            .Except(Parse.LineEnd)
            .AtLeastOnce()
            .Token()
            .Text();

        public readonly static Parser<ArrayParameterSection> Parser =
            from sections in Line.Many()
            select new ArrayParameterSection()
            {
                ParameterNames = sections.ToList()
            };

        #endregion
    }
}
