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
    public async string[] GetDirectoryFiles(string Path)
    {
        Windows.Storage.StorageFolder folder = await Windows.Storage.StorageFolder.GetFolderFromPathAsync(Path);
        IReadOnlyList<Windows.Storage.StorageFile> files = await folder.GetFilesAsync();
        return files.Select(file => file.Name);
    }
    public async Stream LoadStream(string FilePath)
    {
        Windows.Storage.StorageFile file = await Windows.Storage.StorageFile.GetFileFromPathAsync(FilePath);
        System.IO.Stream Stream = await file.OpenStreamForReadAsync();
        return Stream;
    }
    public async void SaveStream(string FilePath, System.IO.Stream Stream)
    {
        Windows.Storage.StorageFile file = await Windows.Storage.StorageFile.GetFileFromPathAsync(FilePath);
        Stream File = await file.OpenStreamForWriteAsync();
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
    public async void DeleteFile(string FilePath)
    {
        Windows.Storage.StorageFile file = await Windows.Storage.StorageFile.GetFileFromPathAsync(FilePath);
        await file.DeleteAsync();
    }
    public async bool PathExists(string Path)
    {
        return (await (await Windows.Storage.StorageFolder.GetFolderFromPathAsync(System.IO.Path.GetDirectoryName(Path))).TryGetItemAsync(System.IO.Path.GetFileName(Path))) != null;
    }
    public async void CreateDirectory(string Path)
    {
        await (await Windows.Storage.StorageFolder.GetFolderFromPathAsync(System.IO.Path.GetDirectoryName(Path))).CreateFolderAsync(System.IO.Path.GetFileName(Path), Windows.Storage.CreationCollisionOption.FailIfExists);
    }
    public async DateTime PathGetLastWriteTimeUtc(string Path)
    {
        Windows.Storage.StorageFile file = await Windows.Storage.StorageFile.GetFileFromPathAsync(Path);
        return (await file.GetBasicPropertiesAsync()).DateModified;
    }
    public async void PathSetLastWriteTimeUtc(string Path, DateTime Time)
    {
        Windows.Storage.StorageFile file = await Windows.Storage.StorageFile.GetFileFromPathAsync(Path);
        await (await file.OpenTransactedWriteAsync()).CommitAsync();
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
            return (new List<KeyValuePair<string, string[]>>(System.Linq.Enumerable.Select("HostPageUtility=Acct,lang,unicode;IslamResources=Hadith,IslamInfo,IslamSource".Split(';'), Str => new KeyValuePair<String, String[]>(Str.Split('=')[0], Str.Split('=')[1].Split(','))))).ToArray();
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
    }
    public class MyTabViewModel : INotifyPropertyChanged
    {
        public MyTabViewModel()
        {
            Items = System.Linq.Enumerable.Select(IslamMetadata.TanzilReader.GetChapterNames(XMLRender.ArabicData.TranslitScheme.Literal, "PlainRoman"), Arr => new MyTabItem { Title = (string)(Arr.Cast<Object>()).First(), Content = null });
        }

        public IEnumerable<MyTabItem> Items { get; private set; }

        private MyTabItem _selectedItem;

        public MyTabItem SelectedItem
        {
            get { return _selectedItem; }
            set
            {
                _selectedItem = value;
                PropertyChanged(this, new PropertyChangedEventArgs("SelectedItem"));
            }
        }

        #region Implementation of INotifyPropertyChanged

        public event PropertyChangedEventHandler PropertyChanged;

        #endregion
    }

    public class MyTabItem
    {
        public string Title { get; set; }
        public UserControl Content { get; set; }
    }
}
