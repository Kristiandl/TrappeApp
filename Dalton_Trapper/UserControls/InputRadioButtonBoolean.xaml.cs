using System.Windows;
using System.Windows.Controls;

namespace Dalton_Trapper.UserControls
{
    /// <summary>
    /// Interaction logic for InputRadioButtonBoolean.xaml
    /// </summary>
    public partial class InputRadioButtonBoolean : UserControl
    {
        public InputRadioButtonBoolean()
        {
            InitializeComponent();

            // Generate a unique group name for each instance
            string uniqueGroupName = $"Group_{Guid.NewGuid()}";

            // Assign the unique group name to both radio buttons
            YesRadioButton.GroupName = uniqueGroupName;
            NoRadioButton.GroupName = uniqueGroupName;

        }

        public string Label
        {
            get { return (string)GetValue(LabelProperty); }
            set { SetValue(LabelProperty, value); }
        }

        public static readonly DependencyProperty LabelProperty =
            DependencyProperty.Register("Label", typeof(string), typeof(InputRadioButtonBoolean), new PropertyMetadata(string.Empty));

        public bool IsYesChecked
        {
            get { return (bool)GetValue(IsYesCheckedProperty); }
            set { SetValue(IsYesCheckedProperty, value); }
        }

        public static readonly DependencyProperty IsYesCheckedProperty =
            DependencyProperty.Register("IsYesChecked", typeof(bool), typeof(InputRadioButtonBoolean), new PropertyMetadata(false));

        public bool IsNoChecked
        {
            get => !IsYesChecked;
            set => IsYesChecked = !value;
        }
    }
}
