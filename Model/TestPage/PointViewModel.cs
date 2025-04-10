namespace Dalton_Trapper.Model
{
    public class PointViewModel : Utilities.ViewModelBase
    {
        private double _x;
        private double _y;
        private string _id;
        public double X
        {
            get => _x;
            set { _x = value; OnPropertyChanged(); }

        }
        public double Y
        {
            get => _y;
            set { _y = value; OnPropertyChanged(); }
        }

        public string ID
        {
            get
            {
                return _id;
            }
            set
            {
                _id = value;
                OnPropertyChanged();
            }
        }
    }
}
