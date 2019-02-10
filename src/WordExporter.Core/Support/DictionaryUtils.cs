using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WordExporter.Core.Support
{
    public static class DictionaryUtils
    {
        public static TValue TryGetValue<TKey, TValue>(
            this IDictionary<TKey, TValue> dictionary,
            TKey key) where TKey : class
        {
            if (!dictionary.TryGetValue(key, out TValue retValue))
            {
                retValue = default(TValue);
            }
            return retValue;
        }
    }
}
