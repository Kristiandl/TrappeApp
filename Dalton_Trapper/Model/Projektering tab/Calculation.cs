namespace Dalton_Trapper.Model.Projektering_tab
{
    public class Calculation : Utilities.ViewModelBase
    {
        private string _fileName;
        private string _filePath;

        public string FileName
        {
            get
            {
                return _fileName;
            }
            set
            {
                _fileName = value;
                OnPropertyChanged();
            }
        }

        public string FilePath
        {
            get
            {
                return _filePath;
            }
            set
            {
                _filePath = value;
                OnPropertyChanged();
            }
        }
    }
}
