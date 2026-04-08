using System;
using System.Globalization;
using System.Windows;
using System.Windows.Data;

namespace QPK_Keynote_Manager
{
    public class ScopeKindToVisibilityConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value == null || parameter == null) return Visibility.Collapsed;
            string kind = value.ToString();
            string desired = parameter.ToString();
            return string.Equals(kind, desired, StringComparison.OrdinalIgnoreCase)
                ? Visibility.Visible
                : Visibility.Collapsed;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
            => throw new NotImplementedException();
    }
}
