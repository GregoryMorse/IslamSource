using System;
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
            this.InitializeComponent();
        }

        private void Back_Click(object sender, RoutedEventArgs e)
        {
            this.Frame.GoBack();
        }
        public static void InitDefaultSettings()
        {
            if (!Windows.Storage.ApplicationData.Current.LocalSettings.Values.ContainsKey("CurrentFont"))
            {
                Windows.Storage.ApplicationData.Current.LocalSettings.Values["CurrentFont"] = "Times New Roman";
            }
            if (!Windows.Storage.ApplicationData.Current.LocalSettings.Values.ContainsKey("UseColoring"))
            {
                Windows.Storage.ApplicationData.Current.LocalSettings.Values["UseColoring"] = true;
            }
            if (!Windows.Storage.ApplicationData.Current.LocalSettings.Values.ContainsKey("ShowTranslation"))
            {
                Windows.Storage.ApplicationData.Current.LocalSettings.Values["ShowTranslation"] = true;
            }
            if (!Windows.Storage.ApplicationData.Current.LocalSettings.Values.ContainsKey("ShowTransliteration"))
            {
                Windows.Storage.ApplicationData.Current.LocalSettings.Values["ShowTransliteration"] = true;
            }
            if (!Windows.Storage.ApplicationData.Current.LocalSettings.Values.ContainsKey("ShowW4W"))
            {
                Windows.Storage.ApplicationData.Current.LocalSettings.Values["ShowW4W"] = true;
            }
        }
        public string SelectedFont
        {
            get
            {
                return (string)Windows.Storage.ApplicationData.Current.LocalSettings.Values["CurrentFont"];
            }
            set
            {
                Windows.Storage.ApplicationData.Current.LocalSettings.Values["CurrentFont"] = value;
            }
        }
        public List<string> Fonts
        {
            get
            {
                return new List<string> { "Times New Roman", "Traditional Arabic", "Arabic Typesetting", "Sakkal Majalla", "Microsoft Uighur", "Arial", "Global User Interface" };
            }
        }
        public static bool UseColoring {
            get
            {
                return (bool)Windows.Storage.ApplicationData.Current.LocalSettings.Values["UseColoring"];
            }
            set
            {
                Windows.Storage.ApplicationData.Current.LocalSettings.Values["UseColoring"] = value;
            }
        }
        public bool ShowTranslation
        {
            get
            {
                return (bool)Windows.Storage.ApplicationData.Current.LocalSettings.Values["ShowTranslation"];
            }
            set
            {
                Windows.Storage.ApplicationData.Current.LocalSettings.Values["ShowTranslation"] = value;
            }
        }

        public bool ShowTransliteration
        {
            get
            {
                return (bool)Windows.Storage.ApplicationData.Current.LocalSettings.Values["ShowTransliteration"];
            }
            set
            {
                Windows.Storage.ApplicationData.Current.LocalSettings.Values["ShowTransliteration"] = value;
            }
        }

        public bool ShowW4W
        {
            get
            {
                return (bool)Windows.Storage.ApplicationData.Current.LocalSettings.Values["ShowW4W"];
            }
            set
            {
                Windows.Storage.ApplicationData.Current.LocalSettings.Values["ShowW4W"] = value;
            }
        }
    }
}
