using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Xamarin.Forms;

namespace IslamSourceQuranViewer.Xam
{
	public partial class Settings : ContentPage
	{
		public Settings ()
		{
            //this.DataContext = this;
            MyAppSettings = new AppSettings();
            InitializeComponent();
		}
        public AppSettings MyAppSettings { get; set; }
    }
}