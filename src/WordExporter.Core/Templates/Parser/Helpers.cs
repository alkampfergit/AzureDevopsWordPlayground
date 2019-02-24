using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WordExporter.Core.Templates.Parser
{
    public static class Helpers
    {
        public static Boolean GetBooleanValue(this IEnumerable<KeyValue> keyValues, String key)
        {
            var entry = keyValues.FirstOrDefault(k => k.Key.Equals(key, StringComparison.OrdinalIgnoreCase));
            if (entry == null)
                return false;

            return "true".Equals(entry.Value, StringComparison.OrdinalIgnoreCase);
        }

        public static String GetStringValue(this IEnumerable<KeyValue> keyValues, String key)
        {
            var entry = keyValues.FirstOrDefault(k => k.Key.Equals(key, StringComparison.OrdinalIgnoreCase));
            return entry?.Value;
        }
    }
}
