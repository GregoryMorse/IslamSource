using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices.WindowsRuntime;
using Windows.Foundation;
using Windows.Foundation.Collections;
using Xamarin.Forms;

// The Blank Page item template is documented at http://go.microsoft.com/fwlink/?LinkId=234238

namespace IslamSourceQuranViewer.Xam
{
    /// <summary>
    /// An empty page that can be used on its own or navigated to within a Frame.
    /// </summary>
    public sealed partial class About : ContentPage
    {
        public About()
        {
            this.InitializeComponent();
            DisplayName.Text = (string)Application.Current.Resources["DisplayName"];
            Description.Text = (string)Application.Current.Resources["Description"];
        }
    }
}
