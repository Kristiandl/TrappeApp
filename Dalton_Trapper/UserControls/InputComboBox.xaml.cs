using System.Collections;
using System.Windows;
using System.Windows.Controls;

namespace Dalton_Trapper.UserControls
{
    public partial class InputComboBox : UserControl
    {
        public InputComboBox()
        {
            InitializeComponent();
        }

        // Dependency property for Label
        public static readonly DependencyProperty LabelProperty =
            DependencyProperty.Register("Label", typeof(string), typeof(InputComboBox), new PropertyMetadata(string.Empty));

        public string Label
        {
            get => (string)GetValue(LabelProperty);
            set => SetValue(LabelProperty, value);
        }

        // Dependency property for Items (items in ComboBox)
        public static readonly DependencyProperty ItemsProperty =
            DependencyProperty.Register("Items", typeof(IEnumerable), typeof(InputComboBox), new PropertyMetadata(null));

        public IEnumerable Items
        {
            get => (IEnumerable)GetValue(ItemsProperty);
            set => SetValue(ItemsProperty, value);
        }

        // Dependency property for SelectedItem (selected item in ComboBox)
        public static readonly DependencyProperty SelectedItemProperty =
            DependencyProperty.Register("SelectedItem", typeof(object), typeof(InputComboBox), new PropertyMetadata(null));

        public object SelectedItem
        {
            get => GetValue(SelectedItemProperty);
            set => SetValue(SelectedItemProperty, value);
        }

        // Dependency property for LabelWidth
        public static readonly DependencyProperty LabelWidthProperty =
            DependencyProperty.Register("LabelWidth", typeof(string), typeof(InputComboBox), new PropertyMetadata(string.Empty));

        public string LabelWidth
        {
            get => (string)GetValue(LabelWidthProperty);
            set => SetValue(LabelWidthProperty, value);
        }

        // Dependency property for InputWidth
        public static readonly DependencyProperty InputWidthProperty =
            DependencyProperty.Register("InputWidth", typeof(string), typeof(InputComboBox), new PropertyMetadata(string.Empty));

        public string InputWidth
        {
            get => (string)GetValue(InputWidthProperty);
            set => SetValue(InputWidthProperty, value);
        }
        
        // Dependency property for ContentWidth
        public static readonly DependencyProperty ContentWidthProperty =
            DependencyProperty.Register("ContentWidth", typeof(string), typeof(InputComboBox), new PropertyMetadata(string.Empty));

        public string ContentWidth
        {
            get => (string)GetValue(ContentWidthProperty);
            set => SetValue(ContentWidthProperty, value);
        }
    }
}
