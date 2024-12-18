namespace Dalton_Trapper.Model.Projektering_tab
{
    public class Calculation : Utilities.ViewModelBase
    {
        private string _fileName;
        private string _filePath;
        private string _fileExtension;
        private bool _isFolder;

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

        public string FileExtension
        {
            get
            {
                return _fileExtension;
            }
            set
            {
                _fileExtension = value;
                OnPropertyChanged();
            }
        }
        public bool IsFolder
        {
            get
            {
                return _isFolder;
            }
            set
            {
                _isFolder = value;
                OnPropertyChanged();
            }
        }
    }
}
