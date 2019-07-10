using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WordExporter.Core
{
    public class WordExporterOptions
    {
        /// <summary>
        /// If true the routine will try to normalize the font in HTML description
        /// removing all font formatting that are embedded in work item html field.
        /// </summary>
        public Boolean NormalizeFontInDescription { get; set; }
    }  
}
