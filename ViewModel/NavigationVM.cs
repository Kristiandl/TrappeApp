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
        private Visibility _testPageVisibility;
        public Visibility TestPageVisibility
        {
            get { return _testPageVisibility; }
            set { _testPageVisibility = value; OnPropertyChanged(); }
        }

        public ICommand HomeCommand { get; set; }
        public ICommand ProjekteringCommand { get; set; }
        public ICommand FemdesignCommand { get; set; }
        public ICommand TestPageCommand { get; set; }


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
        private void TestPage(object obj)
        {
            ConfirmAndNavigate(new TestPageVM());
        }

        private void ConfirmAndNavigate(object newView)
        {
            // Check if current view is FemdesignVM or TestPage and we're navigating away
            if (_currentView is FemdesignVM && !(newView is FemdesignVM) || _currentView is TestPageVM && !(newView is TestPageVM))
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
            TestPageCommand = new RelayCommand(TestPage);

            // Startup Page
            CurrentView = new HomeVM();

            // Only show testpage for KDL
            TestPageVisibility = Environment.UserName.ToLower() == "kdl" ? Visibility.Visible : Visibility.Hidden;
        }
    }
}
