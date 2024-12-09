using Dalton_Trapper.ViewModel;
using System.Windows.Controls;
using System.Windows.Input;

namespace Dalton_Trapper.View
{
    /// <summary>
    /// Interaction logic for Customers.xaml
    /// </summary>
    public partial class Projektering : UserControl
    {
        public Projektering()
        {
            InitializeComponent();
        }


        private void ListViewItem_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            ((ProjekteringVM)DataContext)
                .OpenDrawingCommand.Execute(((ListViewItem)sender).Content);
        }
    }
}
