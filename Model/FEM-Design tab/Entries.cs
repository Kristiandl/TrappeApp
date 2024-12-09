namespace Dalton_Trapper.Model
{
    public class Entries : Utilities.ViewModelBase
    {
        private int _x;
        private int _y;
        private int _x2;
        private int _y2;
        private string _id;

        public Entries() { }

        public int X
        {
            get { return _x; }
            set
            {
                _x = value;
                OnPropertyChanged();
            }
        }

        public int Y
        {
            get { return _y; }
            set
            {
                _y = value;
                OnPropertyChanged();
            }
        }

        public int X2
        {
            get { return _x2; }
            set
            {
                _x2 = value;
                OnPropertyChanged();
            }
        }

        public int Y2
        {
            get { return _y2; }
            set
            {
                _y2 = value;
                OnPropertyChanged();
            }
        }

        public float G { get; set; }
        public float Q { get; set; }

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
