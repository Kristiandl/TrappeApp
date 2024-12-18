using System.Collections;
using System.Windows;
using System.Windows.Controls;


namespace Dalton_Trapper.Utilities
{
    public static class ListViewMultiSelectHelper
    {
        public static readonly DependencyProperty SelectedItemsProperty =
            DependencyProperty.RegisterAttached(
                "SelectedItems",
                typeof(IList),
                typeof(ListViewMultiSelectHelper),
                new PropertyMetadata(null, OnSelectedItemsChanged));

        public static IList GetSelectedItems(DependencyObject obj) =>
            (IList)obj.GetValue(SelectedItemsProperty);

        public static void SetSelectedItems(DependencyObject obj, IList value) =>
            obj.SetValue(SelectedItemsProperty, value);

        private static void OnSelectedItemsChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            if (d is ListView listView)
            {
                listView.SelectionChanged -= ListView_SelectionChanged;
                if (e.NewValue is IList newList)
                {
                    listView.SelectionChanged += ListView_SelectionChanged;
                    UpdateSelectedItems(listView, newList);
                }
            }
        }

        private static void ListView_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (sender is ListView listView)
            {
                var selectedItems = GetSelectedItems(listView);
                if (selectedItems == null) return;

                foreach (var item in e.RemovedItems) selectedItems.Remove(item);
                foreach (var item in e.AddedItems) selectedItems.Add(item);
            }
        }

        private static void UpdateSelectedItems(ListView listView, IList selectedItems)
        {
            listView.SelectedItems.Clear();
            foreach (var item in selectedItems)
            {
                listView.SelectedItems.Add(item);
            }
        }
    }
}
