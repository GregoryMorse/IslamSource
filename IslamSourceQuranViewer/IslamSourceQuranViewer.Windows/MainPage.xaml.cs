using System;
using System.Collections.Generic;
using System.ComponentModel;
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

public class WindowsRTFileIO : XMLRender.PortableFileIO
{
    public /*async*/ string[] GetDirectoryFiles(string Path)
    {
        System.Threading.Tasks.Task<Windows.Storage.StorageFolder> t = Windows.Storage.StorageFolder.GetFolderFromPathAsync(Path).AsTask();
        t.Wait();
        Windows.Storage.StorageFolder folder = t.Result; //await Windows.Storage.StorageFolder.GetFolderFromPathAsync(Path);
        System.Threading.Tasks.Task<IReadOnlyList<Windows.Storage.StorageFile>> tn = folder.GetFilesAsync().AsTask();
        t.Wait();
        IReadOnlyList<Windows.Storage.StorageFile> files = tn.Result; //await folder.GetFilesAsync();
        return new List<string>(files.Select(file => file.Name)).ToArray();
    }
    public /*async*/ Stream LoadStream(string FilePath)
    {
        System.Threading.Tasks.Task<Windows.Storage.StorageFile> t = Windows.Storage.StorageFile.GetFileFromPathAsync(FilePath).AsTask();
        t.Wait();
        Windows.Storage.StorageFile file = t.Result; //await Windows.Storage.StorageFile.GetFileFromPathAsync(FilePath);
        System.Threading.Tasks.Task<Stream> tn = file.OpenStreamForReadAsync();
        tn.Wait();
        Stream Stream = tn.Result; //await file.OpenStreamForReadAsync();
        return Stream;
    }
    public /*async*/ void SaveStream(string FilePath, Stream Stream)
    {
        System.Threading.Tasks.Task<Windows.Storage.StorageFolder> td = Windows.Storage.StorageFolder.GetFolderFromPathAsync(System.IO.Path.GetDirectoryName(FilePath)).AsTask();
        td.Wait();
        System.Threading.Tasks.Task<Windows.Storage.StorageFile> t = td.Result.CreateFileAsync(System.IO.Path.GetFileName(FilePath)).AsTask();
        t.Wait();
        Windows.Storage.StorageFile file = t.Result; //await (await Windows.Storage.StorageFolder.GetFolderFromPathAsync(System.IO.Path.GetDirectoryName(FilePath))).Result.CreateFileAsync(System.IO.Path.GetFileName(FilePath));
        System.Threading.Tasks.Task<Stream> tn = file.OpenStreamForWriteAsync();
        tn.Wait();
        Stream File = tn.Result; //await file.OpenStreamForWriteAsync();
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
        System.Threading.Tasks.Task<Windows.Storage.StorageFile> t = Windows.Storage.StorageFile.GetFileFromPathAsync(FilePath).AsTask();
        t.Wait();
        Windows.Storage.StorageFile file = t.Result; //await Windows.Storage.StorageFile.GetFileFromPathAsync(FilePath);
        file.DeleteAsync().AsTask().Wait();
        //await file.DeleteAsync();
    }
    public /*async*/ bool PathExists(string Path)
    {
        System.Threading.Tasks.Task<Windows.Storage.StorageFolder> t = Windows.Storage.StorageFolder.GetFolderFromPathAsync(System.IO.Path.GetDirectoryName(Path)).AsTask();
        t.Wait();
        System.Threading.Tasks.Task<Windows.Storage.IStorageItem> tn = t.Result.TryGetItemAsync(System.IO.Path.GetFileName(Path)).AsTask();
        tn.Wait();
        return tn.Result != null;
        //return (await (await Windows.Storage.StorageFolder.GetFolderFromPathAsync(System.IO.Path.GetDirectoryName(Path))).TryGetItemAsync(System.IO.Path.GetFileName(Path))) != null;
    }
    public /*async*/ void CreateDirectory(string Path)
    {
        System.Threading.Tasks.Task<Windows.Storage.StorageFolder> t = Windows.Storage.StorageFolder.GetFolderFromPathAsync(System.IO.Path.GetDirectoryName(Path)).AsTask();
        t.Wait();
        t.Result.CreateFolderAsync(System.IO.Path.GetFileName(Path), Windows.Storage.CreationCollisionOption.FailIfExists).AsTask().Wait();
        //await (await Windows.Storage.StorageFolder.GetFolderFromPathAsync(System.IO.Path.GetDirectoryName(Path))).CreateFolderAsync(System.IO.Path.GetFileName(Path), Windows.Storage.CreationCollisionOption.FailIfExists);
    }
    public /*async*/ DateTime PathGetLastWriteTimeUtc(string Path)
    {
        System.Threading.Tasks.Task<Windows.Storage.StorageFile> t = Windows.Storage.StorageFile.GetFileFromPathAsync(Path).AsTask();
        t.Wait();
        Windows.Storage.StorageFile file = t.Result; //await Windows.Storage.StorageFile.GetFileFromPathAsync(Path);
        System.Threading.Tasks.Task<Windows.Storage.FileProperties.BasicProperties> tn = file.GetBasicPropertiesAsync().AsTask();
        tn.Wait();
        return tn.Result.DateModified.UtcDateTime;
        //return (await file.GetBasicPropertiesAsync()).DateModified;
    }
    public /*async*/ void PathSetLastWriteTimeUtc(string Path, DateTime Time)
    {
        System.Threading.Tasks.Task<Windows.Storage.StorageFile> t = Windows.Storage.StorageFile.GetFileFromPathAsync(Path).AsTask();
        t.Wait();
        Windows.Storage.StorageFile file = t.Result; //await Windows.Storage.StorageFile.GetFileFromPathAsync(Path);
        System.Threading.Tasks.Task<Windows.Storage.StorageStreamTransaction> tn = file.OpenTransactedWriteAsync().AsTask();
        tn.Wait();
        tn.Result.CommitAsync().AsTask().Wait();
        //await(await file.OpenTransactedWriteAsync()).CommitAsync();
    }
}
public class WindowsRTSettings : XMLRender.PortableSettings
{
    public string CacheDirectory
    {
        get {
            //Windows.Storage.ApplicationData.Current.LocalFolder.InstalledLocation;
            return Windows.ApplicationModel.Package.Current.InstalledLocation.Path;
        }
    }
    public KeyValuePair<string, string[]>[] Resources
    {
        get {
            return (new List<KeyValuePair<string, string[]>>(System.Linq.Enumerable.Select("HostPageUtility=Acct,lang,unicode;IslamResources=Hadith,IslamInfo,IslamSource".Split(';'), Str => new KeyValuePair<string, string[]>(Str.Split('=')[0], Str.Split('=')[1].Split(','))))).ToArray();
        }
    }
    public string GetFilePath(string Path)
    {
        return System.IO.Path.Combine(Windows.ApplicationModel.Package.Current.InstalledLocation.Path, Path);
    }
    public string GetUName(char Character)
    {
        return "";
    }
}
// The Blank Page item template is documented at http://go.microsoft.com/fwlink/?LinkId=234238

namespace IslamSourceQuranViewer
{
    /// <summary>
    /// An empty page that can be used on its own or navigated to within a Frame.
    /// </summary>
    public sealed partial class MainPage : Page
    {
        public MainPage()
        {
            XMLRender.PortableMethods.FileIO = new WindowsRTFileIO();
            XMLRender.PortableMethods.Settings = new WindowsRTSettings();
            this.DataContext = this;
            this.ViewModel = new MyTabViewModel();
            this.InitializeComponent();
        }
        public MyTabViewModel ViewModel { get; set; }

        private void sectionListBox_DoubleTapped(object sender, DoubleTappedRoutedEventArgs e)
        {
            this.Frame.Navigate(typeof(WordForWordUC), new {Division = ViewModel.SelectedItem.Index, Selection = ViewModel.ListSelectedItem.Index + 1});
        }

        private void RenderPngs_Click(object sender, RoutedEventArgs e)
        {
            this.Frame.Navigate(typeof(SplashScreen));
        }
    }
    public class MyTabViewModel : INotifyPropertyChanged
    {
        public MyTabViewModel()
        {
            Items = System.Linq.Enumerable.Select(IslamMetadata.TanzilReader.GetDivisionTypes(), (Arr, idx) => new MyTabItem { Title = Arr, Index = idx });
        }

        public IEnumerable<MyTabItem> Items { get; set; }

        public IEnumerable<MyListItem> _ListItems;
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
            get { return _selectedItem.SelectedItem; }
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
                if (_Items == null) { _Items = System.Linq.Enumerable.Select(IslamMetadata.TanzilReader.GetSelectionNames(Index.ToString(), XMLRender.ArabicData.TranslitScheme.RuleBased, "PlainRoman"), (Arr, Idx) => new MyListItem { Name = (string)(Arr.Cast<object>()).First(), Index = Idx }); }
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
        public string Name { get; set; }
        public int Index { get; set; }
    }
}
