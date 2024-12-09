namespace Dalton_Trapper.Model
{
    public class CurvedLine : Utilities.ViewModelBase
    {
        private Entries _P1;
        private Entries _P2;
        private int _radius;
        private string _form;

        public CurvedLine() { }

        public Entries P1
        {
            get { return _P1; }
            set
            {
                _P1 = value;
                OnPropertyChanged();
            }
        }
        public Entries P2
        {
            get { return _P2; }
            set
            {
                _P2 = value;
                OnPropertyChanged();
            }
        }

        public int Radius
        {
            get { return _radius; }
            set
            {
                _radius = value;
                OnPropertyChanged();
            }
        }
        public string Form
        {
            get
            {
                return _form;
            }
            set
            {
                _form = value;
                OnPropertyChanged();
            }
        }
    }
}
