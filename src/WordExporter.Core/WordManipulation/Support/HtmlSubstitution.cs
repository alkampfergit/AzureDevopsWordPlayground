using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WordExporter.Core.WordManipulation.Support
{
    public class HtmlSubstitution
    {
        public HtmlSubstitution(string htmlValue)
        {
            HtmlValue = htmlValue;
        }

        public String HtmlValue { get; private set; }
    }
}
