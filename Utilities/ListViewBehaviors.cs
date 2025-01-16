using System.Windows;
using System.Windows.Input;

namespace Dalton_Trapper.Utilities
{
    public static class ListViewBehaviors
    {
        // Double-click command property
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


        // Enter key command property
        public static readonly DependencyProperty EnterKeyCommandProperty =
            DependencyProperty.RegisterAttached(
                "EnterKeyCommand",
                typeof(ICommand),
                typeof(ListViewBehaviors),
                new PropertyMetadata(null, OnEnterKeyCommandChanged));

        public static ICommand GetEnterKeyCommand(DependencyObject obj) =>
            (ICommand)obj.GetValue(EnterKeyCommandProperty);

        public static void SetEnterKeyCommand(DependencyObject obj, ICommand value) =>
            obj.SetValue(EnterKeyCommandProperty, value);

        private static void OnEnterKeyCommandChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            if (d is UIElement element)
            {
                element.KeyDown += (sender, args) =>
                {
                    if (args.Key == Key.Enter)
                    {
                        var command = GetEnterKeyCommand(d);
                        if (command?.CanExecute(null) == true)
                        {
                            var listView = sender as System.Windows.Controls.ListView;
                            var selectedItems = listView?.SelectedItems;

                            if (selectedItems != null && selectedItems.Count > 0)
                            {
                                command.Execute(selectedItems);
                            }
                        }
                    }
                };
            }
        }
    }
}
