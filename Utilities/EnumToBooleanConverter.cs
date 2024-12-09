using System.Globalization;
using System.Windows;
using System.Windows.Data;

namespace Dalton_Trapper.Utilities
{
    public class EnumToBooleanConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is Enum enumValue && parameter is string enumString)
            {
                var enumParameter = Enum.Parse(enumValue.GetType(), enumString);
                return enumValue.Equals(enumParameter);
            }
            return false;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            if (value is bool isChecked && isChecked && parameter is string enumString)
            {
                return Enum.Parse(targetType, enumString);
            }
            return Binding.DoNothing;
        }
    }
}
