using Dalton_Trapper.Model;
using System.Globalization;
using System.Windows;
using System.Windows.Data;

namespace Dalton_Trapper.Utilities
{
    public class BooleanToVisibilityConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is Entries)
            {
                return Visibility.Visible;
            }

            else if (value is CurvedLine)
            {
                return Visibility.Visible;
            }

            else
            { 
                return Visibility.Collapsed; 
            }
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}
