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
    public class AppSettings : INotifyPropertyChanged
    {
        public AppSettings() { }
        public static void InitDefaultSettings()
        {
#if WINDOWS_PHONE
            XMLRender.PortableMethods.FileIO = new WindowsRTXamFileIO();
            XMLRender.PortableMethods.Settings = new WindowsRTXamSettings();
#else
            XMLRender.PortableMethods.FileIO = new AndroidiOSFileIO();
            XMLRender.PortableMethods.Settings = new AndroidiOSSettings();
#endif
            //if (!Windows.Storage.ApplicationData.Current.LocalSettings.Values.ContainsKey("CurrentFont"))
            //{
            //    strSelectedFont = "Times New Roman";
            //}
            //if (!Windows.Storage.ApplicationData.Current.LocalSettings.Values.ContainsKey("OtherCurrentFont"))
            //{
            //    strOtherSelectedFont = "Arial";
            //}
            //if (!Windows.Storage.ApplicationData.Current.LocalSettings.Values.ContainsKey("FontSize"))
            //{
            //    dFontSize = 30.0;
            //}
            //if (!Windows.Storage.ApplicationData.Current.LocalSettings.Values.ContainsKey("OtherFontSize"))
            //{
            //    dOtherFontSize = 20.0;
            //}
            //if (!Windows.Storage.ApplicationData.Current.LocalSettings.Values.ContainsKey("UseColoring"))
            //{
            //    bUseColoring = true;
            //}
            //if (!Windows.Storage.ApplicationData.Current.LocalSettings.Values.ContainsKey("ShowTranslation"))
            //{
            //    bShowTranslation = true;
            //}
            //if (!Windows.Storage.ApplicationData.Current.LocalSettings.Values.ContainsKey("ShowTransliteration"))
            //{
            //    bShowTransliteration = true;
            //}
            //if (!Windows.Storage.ApplicationData.Current.LocalSettings.Values.ContainsKey("ShowW4W"))
            //{
            //    bShowW4W = true;
            //}
            //if (!Windows.Storage.ApplicationData.Current.LocalSettings.Values.ContainsKey("CurrentTranslation"))
            //{
            //    iSelectedTranslation = IslamMetadata.TanzilReader.GetTranslationIndex(String.Empty);
            //}

        }
        //public static string strSelectedFont { get { return (string)Windows.Storage.ApplicationData.Current.LocalSettings.Values["CurrentFont"]; } set { Windows.Storage.ApplicationData.Current.LocalSettings.Values["CurrentFont"] = value; } }
        //public static string strOtherSelectedFont { get { return (string)Windows.Storage.ApplicationData.Current.LocalSettings.Values["OtherCurrentFont"]; } set { Windows.Storage.ApplicationData.Current.LocalSettings.Values["OtherCurrentFont"] = value; } }
        //public string SelectedFont { get { return strSelectedFont; } set { strSelectedFont = value; } }
        //public string OtherSelectedFont { get { return strOtherSelectedFont; } set { strOtherSelectedFont = value; } }

        //public static double dFontSize { get { return (double)Windows.Storage.ApplicationData.Current.LocalSettings.Values["FontSize"]; } set { Windows.Storage.ApplicationData.Current.LocalSettings.Values["FontSize"] = value; } }
        //public static double dOtherFontSize { get { return (double)Windows.Storage.ApplicationData.Current.LocalSettings.Values["OtherFontSize"]; } set { Windows.Storage.ApplicationData.Current.LocalSettings.Values["OtherFontSize"] = value; } }
        //public string FontSize { get { return dFontSize.ToString(); } set { double fontSize; if (double.TryParse(value, out fontSize)) { dFontSize = fontSize; } } }
        //public string OtherFontSize { get { return dOtherFontSize.ToString(); } set { double fontSize; if (double.TryParse(value, out fontSize)) { dOtherFontSize = fontSize; } } }
        public List<string> GetFontList()
        {
            List<string> fontList = new List<string>();
            //SharpDX.DirectWrite.FontCollection fontCollection = MyUIChanger.DWFactory.GetSystemFontCollection(false);
            //for (int i = 0; i < fontCollection.FontFamilyCount; i++)
            //{
            //    int index = 0;
            //    if (!fontCollection.GetFontFamily(i).FamilyNames.FindLocaleName(System.Globalization.CultureInfo.CurrentCulture.Name, out index))
            //    {
            //        for (int j = 0; j < Windows.Globalization.ApplicationLanguages.Languages.Count; j++)
            //        {
            //            if (fontCollection.GetFontFamily(i).FamilyNames.FindLocaleName(Windows.Globalization.ApplicationLanguages.Languages[j], out index))
            //            {
            //                fontList.Add(fontCollection.GetFontFamily(i).FamilyNames.GetString(index));
            //                break;
            //            }
            //        }

            //    }
            //    else { fontList.Add(fontCollection.GetFontFamily(i).FamilyNames.GetString(index)); }

            //}
            return fontList;// new List<string> { "Times New Roman", "Traditional Arabic", "Arabic Typesetting", "Sakkal Majalla", "Microsoft Uighur", "Arial", "Global User Interface" };
        }
        public List<string> Fonts
        {
            get
            {
                return GetFontList();
            }
        }
        public List<string> OtherFonts
        {
            get
            {
                return GetFontList();
            }
        }
        //public static int iSelectedTranslation { get { return (int)Windows.Storage.ApplicationData.Current.LocalSettings.Values["CurrentTranslation"]; } set { Windows.Storage.ApplicationData.Current.LocalSettings.Values["CurrentTranslation"] = value; } }
        //public string SelectedTranslation { get { return IslamMetadata.CachedData.IslamData.Translations.TranslationList[iSelectedTranslation].Name; } set { if (value != null) { iSelectedTranslation = Array.FindIndex(IslamMetadata.CachedData.IslamData.Translations.TranslationList, (Translation) => Translation.Name == value); } } }
        public List<string> TranslationList
        {
            get
            {
                return new List<string>(IslamMetadata.CachedData.IslamData.Translations.TranslationList.Where((Translation) => { return Translation.FileName.Substring(0, Translation.FileName.IndexOf('.')) == System.Globalization.CultureInfo.CurrentCulture.Name || Translation.FileName.Substring(0, Translation.FileName.IndexOf('.')) == System.Globalization.CultureInfo.CurrentCulture.Parent.Name; }).Select((Translation) => Translation.Name));
            }
        }
        //public static bool bUseColoring { get { return (bool)Windows.Storage.ApplicationData.Current.LocalSettings.Values["UseColoring"]; } set { Windows.Storage.ApplicationData.Current.LocalSettings.Values["UseColoring"] = value; } }
        //public static bool bShowTranslation { get { return (bool)Windows.Storage.ApplicationData.Current.LocalSettings.Values["ShowTranslation"]; } set { Windows.Storage.ApplicationData.Current.LocalSettings.Values["ShowTranslation"] = value; } }
        //public static bool bShowTransliteration { get { return (bool)Windows.Storage.ApplicationData.Current.LocalSettings.Values["ShowTransliteration"]; } set { Windows.Storage.ApplicationData.Current.LocalSettings.Values["ShowTransliteration"] = value; } }
        //public static bool bShowW4W { get { return (bool)Windows.Storage.ApplicationData.Current.LocalSettings.Values["ShowW4W"]; } set { Windows.Storage.ApplicationData.Current.LocalSettings.Values["ShowW4W"] = value; } }
        //public bool UseColoring { get { return bUseColoring; } set { bUseColoring = value; } }
        //public bool ShowTranslation { get { return bShowTranslation; } set { bShowTranslation = value; } }
        //public bool ShowTransliteration { get { return bShowTransliteration; } set { bShowTransliteration = value; } }
        //public bool ShowW4W { get { return bShowW4W; } set { bShowW4W = value; } }
        //public static string strAppLanguage
        //{
        //    get { return (string)Windows.Globalization.ApplicationLanguages.PrimaryLanguageOverride; }
        //    set
        //    {
        //        Windows.Globalization.ApplicationLanguages.PrimaryLanguageOverride = (value == Windows.Globalization.ApplicationLanguages.Languages.First() ? string.Empty : value);
        //        App._resourceContext = null;
        //        if (Windows.UI.Xaml.Window.Current != null && Windows.UI.Xaml.Window.Current.CoreWindow != null) Windows.ApplicationModel.Resources.Core.ResourceContext.GetForCurrentView().Reset();
        //        Windows.ApplicationModel.Resources.Core.ResourceContext.GetForViewIndependentUse().Reset();
        //        Windows.ApplicationModel.Resources.Core.ResourceContext.ResetGlobalQualifierValues();
        //        System.Globalization.CultureInfo.DefaultThreadCurrentCulture = new System.Globalization.CultureInfo(value);
        //        System.Globalization.CultureInfo.DefaultThreadCurrentUICulture = new System.Globalization.CultureInfo(value);
        //        while (System.Globalization.CultureInfo.CurrentUICulture.Name != value)
        //        {
        //            System.Threading.Tasks.Task.Delay(100);
        //        }
        //    }
        //}
        //public string AppLanguage { get { return new Windows.Globalization.Language(strAppLanguage == string.Empty ? Windows.Globalization.ApplicationLanguages.Languages.First() : strAppLanguage).DisplayName + " (" + (strAppLanguage == string.Empty ? Windows.Globalization.ApplicationLanguages.Languages.First() : strAppLanguage) + ")"; } set { strAppLanguage = value.Substring(value.LastIndexOf("(")).Trim('(', ')'); PropertyChanged(this, new PropertyChangedEventArgs("TranslationList")); iSelectedTranslation = IslamMetadata.TanzilReader.GetTranslationIndex(String.Empty); PropertyChanged(this, new PropertyChangedEventArgs("SelectedTranslation")); } }

        //public List<string> AppLanguageList
        //{
        //    get
        //    {
        //        return Windows.Globalization.ApplicationLanguages.ManifestLanguages.Select((Item) => new Windows.Globalization.Language(Item).DisplayName + " (" + Item + ")").ToList();
        //    }
        //}
#region Implementation of INotifyPropertyChanged

        public event PropertyChangedEventHandler PropertyChanged;

#endregion
    }
}