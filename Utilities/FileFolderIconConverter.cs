using System.Diagnostics;
using System.Globalization;
using System.Windows.Data;
using System.Windows.Media.Imaging;

namespace Dalton_Trapper.Utilities
{
    public class FileFolderIconConverter : IMultiValueConverter
    {
        public object Convert(object[] values, Type targetType, object parameter, CultureInfo culture)
        {
            bool isFolder = (bool)values[0];

            string extension = (string)values[2];

            if (isFolder)
            {
                return new BitmapImage(new Uri("pack://application:,,,/View/Icons/folder_icon.png"));
            }
            else if (extension == ".pdf")
            {
                return new BitmapImage(new Uri("pack://application:,,,/View/Icons/pdf_icon.png"));
            }
            else if (extension == ".xlsm")
            {
                return new BitmapImage(new Uri("pack://application:,,,/View/Icons/excel_icon.png"));
            }
            else 
            {
                return new BitmapImage(new Uri("pack://application:,,,/View/Icons/question_mark_icon.png"));
            }
        }

        public object[] ConvertBack(object value, Type[] targetTypes, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}


