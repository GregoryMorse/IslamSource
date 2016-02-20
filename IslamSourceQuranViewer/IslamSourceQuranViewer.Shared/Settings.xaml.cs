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
    public class AppSettings //: INotifyPropertyChanged
    {
        public AppSettings() { }
        public static void InitDefaultSettings()
        {
            if (!Windows.Storage.ApplicationData.Current.LocalSettings.Values.ContainsKey("CurrentFont"))
            {
                strSelectedFont = "Times New Roman";
            }
            if (!Windows.Storage.ApplicationData.Current.LocalSettings.Values.ContainsKey("OtherCurrentFont"))
            {
                strOtherSelectedFont = "Times New Roman";
            }
            if (!Windows.Storage.ApplicationData.Current.LocalSettings.Values.ContainsKey("FontSize"))
            {
                dFontSize = 30.0;
            }
            if (!Windows.Storage.ApplicationData.Current.LocalSettings.Values.ContainsKey("OtherFontSize"))
            {
                dOtherFontSize = 20.0;
            }
            if (!Windows.Storage.ApplicationData.Current.LocalSettings.Values.ContainsKey("UseColoring"))
            {
                bUseColoring = true;
            }
            if (!Windows.Storage.ApplicationData.Current.LocalSettings.Values.ContainsKey("ShowTranslation"))
            {
                bShowTranslation = true;
            }
            if (!Windows.Storage.ApplicationData.Current.LocalSettings.Values.ContainsKey("ShowTransliteration"))
            {
                bShowTransliteration = true;
            }
            if (!Windows.Storage.ApplicationData.Current.LocalSettings.Values.ContainsKey("ShowW4W"))
            {
                bShowW4W = true;
            }
        }
        public static string strSelectedFont { get { return (string)Windows.Storage.ApplicationData.Current.LocalSettings.Values["CurrentFont"]; } set { Windows.Storage.ApplicationData.Current.LocalSettings.Values["CurrentFont"] = value; } }
        public static string strOtherSelectedFont { get { return (string)Windows.Storage.ApplicationData.Current.LocalSettings.Values["OtherCurrentFont"]; } set { Windows.Storage.ApplicationData.Current.LocalSettings.Values["OtherCurrentFont"] = value; } }
        public string SelectedFont { get { return strSelectedFont; } set { strSelectedFont = value; } }
        public string OtherSelectedFont { get { return strOtherSelectedFont; } set { strOtherSelectedFont = value; } }

        public static double dFontSize { get { return (double)Windows.Storage.ApplicationData.Current.LocalSettings.Values["FontSize"]; } set { Windows.Storage.ApplicationData.Current.LocalSettings.Values["FontSize"] = value; } }
        public static double dOtherFontSize { get { return (double)Windows.Storage.ApplicationData.Current.LocalSettings.Values["OtherFontSize"]; } set { Windows.Storage.ApplicationData.Current.LocalSettings.Values["OtherFontSize"] = value; } }
        public string FontSize { get { return dFontSize.ToString(); } set { double fontSize;  if (double.TryParse(value, out fontSize)) { dFontSize = fontSize; } } }
        public string OtherFontSize { get { return dOtherFontSize.ToString(); } set { double fontSize; if (double.TryParse(value, out fontSize)) { dOtherFontSize = fontSize; } } }
        public List<string> Fonts
        {
            get
            {
                return new List<string> { "Times New Roman", "Traditional Arabic", "Arabic Typesetting", "Sakkal Majalla", "Microsoft Uighur", "Arial", "Global User Interface" };
            }
        }
        public static bool bUseColoring { get { return (bool)Windows.Storage.ApplicationData.Current.LocalSettings.Values["UseColoring"]; } set { Windows.Storage.ApplicationData.Current.LocalSettings.Values["UseColoring"] = value; } }
        public static bool bShowTranslation { get { return (bool)Windows.Storage.ApplicationData.Current.LocalSettings.Values["ShowTranslation"]; } set { Windows.Storage.ApplicationData.Current.LocalSettings.Values["ShowTranslation"] = value; } }
        public static bool bShowTransliteration { get { return (bool)Windows.Storage.ApplicationData.Current.LocalSettings.Values["ShowTransliteration"]; } set { Windows.Storage.ApplicationData.Current.LocalSettings.Values["ShowTransliteration"] = value; } }
        public static bool bShowW4W { get { return (bool)Windows.Storage.ApplicationData.Current.LocalSettings.Values["ShowW4W"]; } set { Windows.Storage.ApplicationData.Current.LocalSettings.Values["ShowW4W"] = value; } }
        public bool UseColoring { get { return bUseColoring; } set { bUseColoring = value; } }
        public bool ShowTranslation { get { return bShowTranslation; } set { bShowTranslation = value; } }
        public bool ShowTransliteration { get { return bShowTransliteration; } set { bShowTransliteration = value; } }
        public bool ShowW4W { get { return bShowW4W; } set { bShowW4W = value; } }
        //#region Implementation of INotifyPropertyChanged

        //public event PropertyChangedEventHandler PropertyChanged;

        //#endregion
    }
}
