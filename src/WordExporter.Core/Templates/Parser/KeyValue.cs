using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WordExporter.Core.Templates.Parser
{
    public class KeyValue
    {
        public KeyValue(string key, string value)
        {
            Key = key;
            Value = value;
        }

        public String Key { get; private set; }

        public String Value { get; private set; }
    }
}
