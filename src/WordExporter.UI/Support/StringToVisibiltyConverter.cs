using System;
using System.Globalization;
using System.Windows;
using System.Windows.Data;
using System.Windows.Markup;

namespace WordExporter.UI.Support
{
    public class StringToVisibiltyConverter : MarkupExtension, IValueConverter
    {
        public Boolean Invert { get; set; }

        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is String str && str.Equals(Value, StringComparison.OrdinalIgnoreCase))
            {
                return Invert ? Visibility.Collapsed : Visibility.Visible;
            }

            return Invert ? Visibility.Visible : Visibility.Collapsed;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }

        public String Value { get; set; }

        public override object ProvideValue(IServiceProvider serviceProvider)
        {
            return this;
        }
    }
}
