using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WordExporter.Core
{
    public static class Registry
    {
        static Registry()
        {
            Options = new WordExporterOptions();
        }

        public static WordExporterOptions Options { get; private set; }
    }
}
