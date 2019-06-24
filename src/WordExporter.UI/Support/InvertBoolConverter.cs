using System;
using System.Globalization;
using System.Windows.Data;
using System.Windows.Markup;

namespace WordExporter.UI.Support
{
    /// <summary>
    /// Inverts a boolean value.
    /// </summary>
    public class InvertBoolConverter : MarkupExtension, IValueConverter
    {
        public object Convert(
            object value,
            Type targetType,
            object parameter,
            CultureInfo culture)
        {
            if (value is bool x)
            {
                return !x;
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
            if (value is bool x)
            {
                return !x;
            }
            else
            {
                throw new ArgumentException();
            }
        }

        public override object ProvideValue(IServiceProvider serviceProvider)
        {
            return this;
        }
    }
}
