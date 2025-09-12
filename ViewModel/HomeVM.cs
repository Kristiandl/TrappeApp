using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Media;

namespace Dalton_Trapper.ViewModel
{
    class HomeVM : Utilities.ViewModelBase
    {
        public HomeVM()
        {
            string user = Environment.UserName.ToLower();

            if (user == "tho" || user == "nto")
            {
                ImgSource = "/Images/FCM_logo_og_agf.png";
            }

            else
            {
                ImgSource = "/Images/FCM_logo.png";
            }
        }

        private string imgSource;
        public string ImgSource
        {
            get => imgSource;
            set
            {
                imgSource = value;
                OnPropertyChanged(nameof(ImgSource));
            }
        }
    }
}
