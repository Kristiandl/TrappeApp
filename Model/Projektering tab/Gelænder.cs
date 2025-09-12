using System;

namespace Dalton_Trapper.Model.Projektering_tab
{
    class Gelænder : Utilities.ViewModelBase
    {
        private string _type;
        private string _name;

        public string Type
        {
            get => _type;
            set { _type = value; OnPropertyChanged(nameof(Type)); }
        }

        public string Name
        {
            get => _name;
            set { _name = value; OnPropertyChanged(nameof(Name)); }
        }
    }
}
