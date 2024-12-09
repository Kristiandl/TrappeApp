using System.Windows;
using System.Windows.Input;

namespace Dalton_Trapper.Utilities
{
    public static class ListViewBehaviors
    {
        public static readonly DependencyProperty DoubleClickCommandProperty =
            DependencyProperty.RegisterAttached(
                "DoubleClickCommand",
                typeof(ICommand),
                typeof(ListViewBehaviors),
                new PropertyMetadata(null, OnDoubleClickCommandChanged));

        public static ICommand GetDoubleClickCommand(DependencyObject obj) =>
            (ICommand)obj.GetValue(DoubleClickCommandProperty);

        public static void SetDoubleClickCommand(DependencyObject obj, ICommand value) =>
            obj.SetValue(DoubleClickCommandProperty, value);

        private static void OnDoubleClickCommandChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            if (d is UIElement element)
            {
                element.MouseLeftButtonDown += (sender, args) =>
                {
                    if (args.ClickCount == 2)
                    {
                        var command = GetDoubleClickCommand(d);
                        if (command?.CanExecute(null) == true)
                        {
                            var listView = sender as System.Windows.Controls.ListView;
                            var selectedItem = listView?.SelectedItem;

                            command.Execute(selectedItem);
                        }
                    }
                };
            }
        }
    }
}

