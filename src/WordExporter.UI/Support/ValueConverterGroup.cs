using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Data;

namespace WordExporter.UI.Support
{
    /// <summary>
    /// Converter to chain together multiple converters.
    /// </summary>
    public class ValueConverterGroup : List<IValueConverter>, IValueConverter
    {
        public object Convert(object value,
                              Type targetType,
                              object parameter,
                              CultureInfo culture)
        {
            object curValue = value;
            for (int i = 0; i < this.Count; i++)
            {
                curValue = this[i].Convert(curValue, targetType, parameter, culture);
            }

            return (curValue);
        }

        public object ConvertBack(object value,
                                  Type targetType,
                                  object parameter,
                                  CultureInfo culture)
        {
            object curValue = value;
            for (int i = (this.Count - 1); i >= 0; i--)
            {
                curValue = this[i].ConvertBack(curValue, targetType, parameter, culture);
            }

            return (curValue);
        }
    }
}
