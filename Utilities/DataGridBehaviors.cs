namespace Dalton_Trapper.Utilities
{
    using System.Windows;
    using System.Windows.Controls;
    using System.Windows.Input;
    using System.Windows.Media;

    public static class DataGridBehaviors
    {
        public static readonly DependencyProperty ClearSelectionOnLostFocusProperty =
            DependencyProperty.RegisterAttached(
                "ClearSelectionOnLostFocus",
                typeof(bool),
                typeof(DataGridBehaviors),
                new PropertyMetadata(false, OnClearSelectionOnLostFocusChanged));

        public static bool GetClearSelectionOnLostFocus(DependencyObject obj)
        {
            return (bool)obj.GetValue(ClearSelectionOnLostFocusProperty);
        }

        public static void SetClearSelectionOnLostFocus(DependencyObject obj, bool value)
        {
            obj.SetValue(ClearSelectionOnLostFocusProperty, value);
        }

        private static void OnClearSelectionOnLostFocusChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            if (d is DataGrid dataGrid)
            {
                if ((bool)e.NewValue)
                {
                    // Attach event handlers
                    dataGrid.LostFocus += DataGrid_LostFocus;
                    dataGrid.PreviewMouseLeftButtonDown += DataGrid_MouseDownOutside;
                }
                else
                {
                    // Detach event handlers
                    dataGrid.LostFocus -= DataGrid_LostFocus;
                    dataGrid.PreviewMouseLeftButtonDown -= DataGrid_MouseDownOutside;
                }
            }
        }

        private static void DataGrid_LostFocus(object sender, RoutedEventArgs e)
        {
            if (sender is DataGrid dataGrid)
            {
                dataGrid.SelectedItem = null;
            }
        }

        private static void DataGrid_MouseDownOutside(object sender, MouseButtonEventArgs e)
        {
            if (sender is DataGrid dataGrid)
            {
                var hitTestResult = VisualTreeHelper.HitTest(dataGrid, Mouse.GetPosition(dataGrid));
                if (hitTestResult == null)
                {
                    dataGrid.SelectedItem = null;
                }
            }
        }
    }

}
