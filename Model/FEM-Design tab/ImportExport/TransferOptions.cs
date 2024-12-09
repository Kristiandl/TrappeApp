namespace Dalton_Trapper.Model.ImportExport
{
    public class TransferOptions : Utilities.ViewModelBase
    {
        private bool _isYesChecked;
        public bool IsYesChecked
        {
            get => _isYesChecked;
            set
            {
                if (_isYesChecked != value)
                {
                    _isYesChecked = value;
                    OnPropertyChanged(nameof(IsYesChecked));
                }
            }
        }

        private bool _isNoChecked;
        public bool IsNoChecked
        {
            get => _isNoChecked;
            set
            {
                if (_isNoChecked != value)
                {
                    _isNoChecked = value;
                    OnPropertyChanged(nameof(IsNoChecked));
                }
            }
        }

        public string Label { get; set; }

    }
}
