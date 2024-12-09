using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Media;

namespace Dalton_Trapper.Model
{
    public class Labels : Utilities.ViewModelBase
    {
        public Labels() { }

        private double _x;
        private double _y;
        private string _id;
        private Brush _colour;
        private double _angle;


        public double X
        {
            get { return _x; }
            set
            {
                _x = value;
                OnPropertyChanged();
            }
        }

        public double Y
        {
            get { return _y; }
            set
            {
                _y = value;
                OnPropertyChanged();
            }
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

        public Brush Colour
        {
            get
            {
                return _colour;
            }
            set
            {
                _colour = value;
                OnPropertyChanged();
            }
        }

        public double Angle
        {
            get { return _angle; }
            set
            {
                _angle = value;
                OnPropertyChanged();
            }
        }
    }
}
