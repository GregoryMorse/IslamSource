using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Xamarin.Forms;

public class AndroidiOSFileIO : XMLRender.PortableFileIO
{
    public /*async*/ string[] GetDirectoryFiles(string Path)
    {
        return System.IO.Directory.GetFiles(Path);
    }
    public /*async*/ Stream LoadStream(string FilePath)
    {
#if __ANDROID__
        return ((IslamSourceQuranViewer.Xam.Droid.MainActivity)global::Xamarin.Forms.Forms.Context).Assets.Open(FilePath.Trim('/'));
#endif
#if __IOS__
        return null;
#endif
        //return File.Open(FilePath, FileMode.Open, FileAccess.Read);
    }
    public /*async*/ void SaveStream(string FilePath, Stream Stream)
    {
        FileStream File = System.IO.File.Open(FilePath, FileMode.Create, FileAccess.Write);
        File.Seek(0, SeekOrigin.Begin);
        byte[] Bytes = new byte[4096];
        int Read;
        Stream.Seek(0, SeekOrigin.Begin);
        Read = Stream.Read(Bytes, 0, Bytes.Length);
        while (Read != 0)
        {
            File.Write(Bytes, 0, Read);
            Read = Stream.Read(Bytes, 0, Bytes.Length);
        }
        File.Dispose();
    }
    public string CombinePath(params string[] Paths)
    {
        return System.IO.Path.Combine(Paths);
    }
    public /*async*/ void DeleteFile(string FilePath)
    {
        File.Delete(FilePath);
    }
    public /*async*/ bool PathExists(string Path)
    {
        return Directory.Exists(Path);
    }
    public /*async*/ void CreateDirectory(string Path)
    {        
        Directory.CreateDirectory(Path);
    }
    public /*async*/ DateTime PathGetLastWriteTimeUtc(string Path)
    {
        return System.IO.File.GetLastWriteTimeUtc(Path);
    }
    public /*async*/ void PathSetLastWriteTimeUtc(string Path, DateTime Time)
    {
        System.IO.File.SetLastWriteTimeUtc(Path, Time);
    }
}
public class AndroidiOSSettings : XMLRender.PortableSettings
{
    public string CacheDirectory
    {
        get
        {
#if __ANDROID__
            return ((IslamSourceQuranViewer.Xam.Droid.MainActivity)global::Xamarin.Forms.Forms.Context).CacheDir.Path;
#endif
#if __IOS__
            return string.Empty;
#endif
        }
    }
    public KeyValuePair<string, string[]>[] Resources
    {
        get
        {
            return (new List<KeyValuePair<string, string[]>>(System.Linq.Enumerable.Select("HostPageUtility=Acct,lang,unicode;IslamResources=Hadith,IslamInfo,IslamSource".Split(';'), Str => new KeyValuePair<string, string[]>(Str.Split('=')[0], Str.Split('=')[1].Split(','))))).ToArray();
        }
    }
    public string[] FuncLibs
    {
        get
        {
            return new string[] { "IslamMetadata" };
        }
    }
    public string GetTemplatePath()
    {
        return GetFilePath("metadata/IslamSource.xml");
    }
    public string GetFilePath(string Path)
    {
        return System.IO.Path.Combine(Directory.GetCurrentDirectory(), Path);
    }
    public string GetUName(char Character)
    {
        return "";
    }
    public string GetResourceString(string baseName, string resourceKey)
    {
        return new System.Resources.ResourceManager(baseName + ".Resources", System.Reflection.Assembly.Load(baseName)).GetString(resourceKey, System.Threading.Thread.CurrentThread.CurrentUICulture);
    }
}

namespace IslamSourceQuranViewer.Xam
{
	public partial class MainPage : ContentPage
	{
		public MainPage ()
		{
            this.BindingContext = this;
            this.ViewModel = new MyTabViewModel();
            UIChanger = new MyUIChanger();
            InitializeComponent();
            ViewModel.Items = System.Linq.Enumerable.Select(IslamMetadata.TanzilReader.GetDivisionTypes(), (Arr, idx) => new MyTabItem { Title = Arr, Index = idx });
        }
        public MyUIChanger UIChanger { get; set; }
        public MyTabViewModel ViewModel { get; set; }
        private void sectionListBox_DoubleTapped(object sender, ItemTappedEventArgs e)
        {
            if (ViewModel.ListSelectedItem == null) return;
            //this.Frame.Navigate(typeof(WordForWordUC), new { Division = ViewModel.SelectedItem.Index, Selection = ViewModel.ListSelectedItem.Index });
        }
    }
    //public static class ListFormattedTextBehavior
    //{
    //    #region FormattedText Attached dependency property

    //    public static List<string> GetFormattedText(DependencyObject obj)
    //    {
    //        return (List<string>)obj.GetValue(FormattedTextProperty);
    //    }

    //    public static void SetFormattedText(DependencyObject obj, List<string> value)
    //    {
    //        obj.SetValue(FormattedTextProperty, value);
    //    }

    //    public static readonly DependencyProperty FormattedTextProperty =
    //        DependencyProperty.RegisterAttached("FormattedText",
    //        typeof(List<string>),
    //        typeof(ListFormattedTextBehavior),
    //        new PropertyMetadata(null, FormattedTextChanged));

    //    private static void FormattedTextChanged(DependencyObject sender, DependencyPropertyChangedEventArgs e)
    //    {
    //        List<string> value = e.NewValue as List<string>;

    //        TextBlock textBlock = sender as TextBlock;

    //        if (textBlock != null)
    //        {
    //            textBlock.Inlines.Clear();
    //            textBlock.Inlines.Add(new Windows.UI.Xaml.Documents.Run() { Text = value[0], FontFamily = new FontFamily(AppSettings.strOtherSelectedFont), FontSize = AppSettings.dOtherFontSize });
    //            if (value.Count > 1)
    //            {
    //                textBlock.Inlines.Add(new Windows.UI.Xaml.Documents.Run() { Text = "(" + value[1] + ")", FontWeight = Windows.UI.Text.FontWeights.Bold, FlowDirection = FlowDirection.RightToLeft, FontFamily = new FontFamily(AppSettings.strSelectedFont), FontSize = AppSettings.dFontSize });
    //                textBlock.Inlines.Add(new Windows.UI.Xaml.Documents.Run() { Text = value[2], FontFamily = new FontFamily(AppSettings.strOtherSelectedFont), FontSize = AppSettings.dOtherFontSize });
    //            }
    //        }
    //    }

    //    #endregion
    //}
    public class MyTabViewModel : INotifyPropertyChanged
    {
        public MyTabViewModel()
        {
        }
        private IEnumerable<MyTabItem> _Items;
        public IEnumerable<MyTabItem> Items
        {
            get
            {
                return _Items;
            }
            set
            {
                _Items = value;
                PropertyChanged(this, new PropertyChangedEventArgs("Items"));
            }
        }

        private IEnumerable<MyListItem> _ListItems;
        public IEnumerable<MyListItem> ListItems
        {
            get
            {
                return _ListItems;
            }
            private set
            {
                _ListItems = value;
                PropertyChanged(this, new PropertyChangedEventArgs("ListItems"));
            }
        }

        public MyListItem ListSelectedItem
        {
            get { return _selectedItem == null ? null : _selectedItem.SelectedItem; }
            set
            {
                _selectedItem.SelectedItem = value;
                PropertyChanged(this, new PropertyChangedEventArgs("ListSelectedItem"));
            }
        }

        private MyTabItem _selectedItem;

        public MyTabItem SelectedItem
        {
            get { return _selectedItem; }
            set
            {
                _selectedItem = value;
                PropertyChanged(this, new PropertyChangedEventArgs("SelectedItem"));
                ListItems = _selectedItem.Items;
                ListSelectedItem = ListItems.First();
            }
        }

#region Implementation of INotifyPropertyChanged

        public event PropertyChangedEventHandler PropertyChanged;

#endregion
    }

    public class MyTabItem
    {
        public string Title { get; set; }
        public int Index { get; set; }
        private IEnumerable<MyListItem> _Items;
        public IEnumerable<MyListItem> Items
        {
            get
            {
                if (_Items == null) { _Items = System.Linq.Enumerable.Select(IslamMetadata.TanzilReader.GetSelectionNames(Index.ToString(), XMLRender.ArabicData.TranslitScheme.RuleBased, String.Empty), (Arr, Idx) => new MyListItem { TextItems = new List<string>(((string)(Arr.Cast<object>()).First()).Split('(', ')')), Index = (int)(Arr.Cast<object>()).Last() }); }
                return _Items;
            }
        }

        private MyListItem _selectedItem;

        public MyListItem SelectedItem
        {
            get { return _selectedItem; }
            set
            {
                _selectedItem = value;
            }
        }
    }

    public class MyListItem
    {
        public List<string> TextItems { get; set; }
        public int Index { get; set; }
    }
}
