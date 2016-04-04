using System;
using System.ComponentModel;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices.WindowsRuntime;
using Windows.Foundation;
using Windows.Foundation.Collections;
using Windows.UI.Xaml;
using Windows.UI.Xaml.Controls;
using Windows.UI.Xaml.Controls.Primitives;
using Windows.UI.Xaml.Data;
using Windows.UI.Xaml.Input;
using Windows.UI.Xaml.Media;
using Windows.UI.Xaml.Navigation;

// The Blank Page item template is documented at http://go.microsoft.com/fwlink/?LinkId=234238

namespace IslamSourceQuranViewer
{
    /// <summary>
    /// An empty page that can be used on its own or navigated to within a Frame.
    /// </summary>
    public sealed partial class Settings : Page
    {
        public Settings()
        {
            this.DataContext = this;
            MyAppSettings = new AppSettings();
            this.InitializeComponent();
#if WINDOWS_APP
            AppBarButton BackButton = new AppBarButton() { Icon = new SymbolIcon(Symbol.Back), Label = new Windows.ApplicationModel.Resources.ResourceLoader().GetString("Back/Label") };
            BackButton.Click += Back_Click;
            (this.BottomAppBar as CommandBar).PrimaryCommands.Add(BackButton);
#endif
        }

        private void Back_Click(object sender, RoutedEventArgs e)
        {
            this.Frame.GoBack();
        }
        public AppSettings MyAppSettings { get; set; }
    }
}
