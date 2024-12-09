using Dalton_Trapper.Utilities;
using System.Windows;
using System.Windows.Input;

namespace Dalton_Trapper.ViewModel
{
    class NavigationVM : ViewModelBase
    {
        private object _currentView;
        public object CurrentView
        {
            get { return _currentView; }
            set { _currentView = value; OnPropertyChanged(); }
        }

        public ICommand HomeCommand { get; set; }
        public ICommand ProjekteringCommand { get; set; }
        public ICommand FemdesignCommand { get; set; }


        private void Home(object obj)
        { 
            ConfirmAndNavigate(new HomeVM()); 
        }
        private void Projektering(object obj)
        {
            ConfirmAndNavigate(new ProjekteringVM());
        }
        private void Femdesign(object obj)
        {
            ConfirmAndNavigate(new FemdesignVM());
        }

        private void ConfirmAndNavigate(object newView)
        {
            // Check if current view is FemdesignVM and we're navigating away
            if (_currentView is FemdesignVM && !(newView is FemdesignVM))
            {
                var result = MessageBox.Show(
                    "Du er på vej væk fra FEM-Design fanen. Al indtastet data vil blive slettet. Er du sikker på du vil fortsætte væk fra fanen?",
                    "Data vil blive slettet!",
                    MessageBoxButton.YesNo,
                    MessageBoxImage.Warning
                );

                if (result != MessageBoxResult.Yes)
                {
                    return; // Stay on the current view
                }
            }

            CurrentView = newView;
        }

        public NavigationVM()
        {
            HomeCommand = new RelayCommand(Home);
            ProjekteringCommand = new RelayCommand(Projektering);
            FemdesignCommand = new RelayCommand(Femdesign);

            // Startup Page
            CurrentView = new HomeVM();
        }
    }
}
