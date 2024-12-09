using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;

namespace Dalton_Trapper.Model
{
    public class Reinforcement : Utilities.ViewModelBase
    {
        public string Navn { get; set; }

        private double diameter;
        public double Diameter
        {
            get { return diameter; }
            set
            {
                if (diameter != value)
                {
                    diameter = value;
                    OnPropertyChanged(nameof(Diameter));
                    DiameterChanged?.Invoke(this, EventArgs.Empty);
                }
            }
        }

        private int afstand;
        public int Afstand
        {
            get { return afstand; }
            set
            {
                if (afstand != value)
                {
                    afstand = value;
                    OnPropertyChanged(nameof(afstand));
                }
            }
        }
        private double _dæklag;
        public double Dæklag
        {
            get { return _dæklag; }
            set
            {
                if (_dæklag != value)
                {
                    _dæklag = value;
                    OnPropertyChanged(nameof(Dæklag));
                }
            }
        }

        private string _kvalitet;
        public string Kvalitet
        {
            get { return _kvalitet; }
            set
            {
                if (_kvalitet != value)
                {
                    _kvalitet = value;
                    OnPropertyChanged(nameof(Kvalitet));
                }
            }
        }

        public static event EventHandler DiameterChanged;
    }
}
