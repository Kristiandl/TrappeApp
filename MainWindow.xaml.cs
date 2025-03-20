using Dalton_Trapper.Utilities;
using System.Windows;
using System.Windows.Input;

namespace Dalton_Trapper
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            string version = PublishVersion.GetVersion();
            AppVersion.Text = version.Remove(version.Length - 2, 2);
        }

        private void CloseApp_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

        private void MinimizeApp_Click(object sender, RoutedEventArgs e)
        {
            this.WindowState = WindowState.Minimized;
        }

        private void Window_MouseDown(object sender, MouseButtonEventArgs e)
        {
            if (e.LeftButton == MouseButtonState.Pressed) 
            { 
                if (e.GetPosition(this).Y < 65) 
                    { 
                    this.DragMove(); 
                } 
            }
        }
    }
}