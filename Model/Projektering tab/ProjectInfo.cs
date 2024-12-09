using System.Collections.ObjectModel;

namespace Dalton_Trapper.Model.Projektering_tab
{
    public class ProjectInfo : Utilities.ViewModelBase
    {
        private string _projectNumber;
        private string _projectName;
        private string _konsekvensklasse;
        private string _miljøklasse;
        private string _liveload;

        public string ProjectNumber
        {
            get => _projectNumber;
            set { _projectNumber = value; OnPropertyChanged(nameof(ProjectNumber)); }
        }

        public string ProjectName
        {
            get => _projectName;
            set { _projectName = value; OnPropertyChanged(nameof(ProjectName)); }
        }

        public string Konsekvensklasse
        {
            get => _konsekvensklasse;
            set { _konsekvensklasse = value; OnPropertyChanged(nameof(Konsekvensklasse)); }
        }

        public string Miljøklasse
        {
            get => _miljøklasse;
            set { _miljøklasse = value; OnPropertyChanged(nameof(Miljøklasse)); }
        }

        public string Liveload
        {
            get => _liveload;
            set { _liveload = value; OnPropertyChanged(nameof(Liveload)); }
        }
    }
}
