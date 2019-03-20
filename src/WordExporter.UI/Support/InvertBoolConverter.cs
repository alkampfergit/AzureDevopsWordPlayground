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
    /// Inverts a boolean value.
    /// </summary>
    public class InvertBoolConverter : IValueConverter
    {
        public object Convert(
            object value,
            Type targetType,
            object parameter,
            CultureInfo culture)
        {
            if (value is bool)
            {
                return (!(bool)value);
            }
            else
            {
                throw new ArgumentException("Value must be of the type bool");
            }
        }

        public object ConvertBack(
            object value,
            Type targetType,
            object parameter,
            CultureInfo culture)
        {
            if (value is bool)
            {
                return (!(bool)value);
            }
            else
            {
                throw new ArgumentException();
            }
        }
    }
}
