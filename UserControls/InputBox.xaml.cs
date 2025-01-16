using Dalton_Trapper.Utilities;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace Dalton_Trapper.UserControls
{
    /// <summary>
    /// Interaction logic for InputBox.xaml
    /// </summary>
    public partial class InputBox : UserControl
    {
        public InputBox()
        {
            InitializeComponent();
        }

        // Dependency property for Label
        public static readonly DependencyProperty LabelProperty =
            DependencyProperty.Register("Label", typeof(string), typeof(InputBox), new PropertyMetadata(string.Empty));

        public string Label
        {
            get => (string)GetValue(LabelProperty);
            set => SetValue(LabelProperty, value);
        }

        // Dependency property for Text
        public static readonly DependencyProperty TextProperty =
            DependencyProperty.Register("Text", typeof(string), typeof(InputBox), new PropertyMetadata(string.Empty));

        public string Text
        {
            get => (string)GetValue(TextProperty);
            set => SetValue(TextProperty, value);
        }

        // Dependency property for LabelWidth
        public static readonly DependencyProperty LabelWidthProperty =
            DependencyProperty.Register("LabelWidth", typeof(string), typeof(InputBox), new PropertyMetadata(string.Empty));

        public string LabelWidth
        {
            get => (string)GetValue(LabelWidthProperty);
            set => SetValue(LabelWidthProperty, value);
        }

        // Dependency property for InputWidth
        public static readonly DependencyProperty InputWidthProperty =
            DependencyProperty.Register("InputWidth", typeof(string), typeof(InputBox), new PropertyMetadata(string.Empty));

        public string InputWidth
        {
            get => (string)GetValue(InputWidthProperty);
            set => SetValue(InputWidthProperty, value);
        }
    }
}
