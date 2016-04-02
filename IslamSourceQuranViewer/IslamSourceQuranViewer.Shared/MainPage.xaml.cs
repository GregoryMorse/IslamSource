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
        Windows.ApplicationModel.Resources.Core.ResourceCandidate rc = null;
        if (IslamSourceQuranViewer.App._resourceContext == null && Windows.UI.Xaml.Window.Current != null && Windows.UI.Xaml.Window.Current.CoreWindow != null) { IslamSourceQuranViewer.App._resourceContext = Windows.ApplicationModel.Resources.Core.ResourceContext.GetForCurrentView(); } else { IslamSourceQuranViewer.App._resourceContext = Windows.ApplicationModel.Resources.Core.ResourceContext.GetForViewIndependentUse(); }
        if (IslamSourceQuranViewer.App._resourceContext != null) { rc = Windows.ApplicationModel.Resources.Core.ResourceManager.Current.MainResourceMap.GetValue(FilePath.Replace(Windows.ApplicationModel.Package.Current.InstalledLocation.Path, "ms-resource:///Files").Replace("\\", "/"), IslamSourceQuranViewer.App._resourceContext); }
        System.Threading.Tasks.Task<Windows.Storage.StorageFile> t;
        if (rc != null && rc.IsMatch) {
            t = rc.GetValueAsFileAsync().AsTask();
        } else {
            t = Windows.Storage.StorageFile.GetFileFromPathAsync(FilePath).AsTask();
        }
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
        System.Threading.Tasks.Task<Windows.Storage.StorageFile> t = td.Result.CreateFileAsync(System.IO.Path.GetFileName(FilePath), Windows.Storage.CreationCollisionOption.ReplaceExisting).AsTask();
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
#if WINDOWS_PHONE_APP
        System.Threading.Tasks.Task<IReadOnlyList<Windows.Storage.IStorageItem>> files = t.Result.GetItemsAsync().AsTask();
        files.Wait();
        return files.Result.FirstOrDefault(p => p.Name == Path) != null;
#else
        System.Threading.Tasks.Task<Windows.Storage.IStorageItem> tn = t.Result.TryGetItemAsync(System.IO.Path.GetFileName(Path)).AsTask();
        tn.Wait();
        return tn.Result != null;
#endif
        //return (await (await Windows.Storage.StorageFolder.GetFolderFromPathAsync(System.IO.Path.GetDirectoryName(Path))).TryGetItemAsync(System.IO.Path.GetFileName(Path))) != null;
    }
    public /*async*/ void CreateDirectory(string Path)
    {
        System.Threading.Tasks.Task<Windows.Storage.StorageFolder> t = Windows.Storage.StorageFolder.GetFolderFromPathAsync(System.IO.Path.GetDirectoryName(Path)).AsTask();
        t.Wait();
        t.Result.CreateFolderAsync(System.IO.Path.GetFileName(Path), Windows.Storage.CreationCollisionOption.OpenIfExists).AsTask().Wait();
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
        get
        {
            //Windows.Storage.ApplicationData.Current.LocalFolder.InstalledLocation;
            //Windows.ApplicationModel.Package.Current.InstalledLocation.Path;
            return Windows.Storage.ApplicationData.Current.TemporaryFolder.Path;
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
        return GetFilePath("metadata\\IslamSource.xml");
    }
    public string GetFilePath(string Path)
    {
        return System.IO.Path.Combine(Windows.ApplicationModel.Package.Current.InstalledLocation.Path, Path);
    }
    public string GetUName(char Character)
    {
        return "";
    }
    public string GetResourceString(string baseName, string resourceKey)
    {
        return Windows.ApplicationModel.Resources.ResourceLoader.GetForViewIndependentUse(baseName + ".Resources").GetString(resourceKey);
    }
    public static async System.Threading.Tasks.Task SavePathImageAsFile(int Width, int Height, string fileName, FrameworkElement element, bool UseRenderTarget = true)
    {
        double oldWidth = element.Width;
        double oldHeight = element.Height;
        double actOldWidth = element.ActualWidth;
        double actOldHeight = element.ActualHeight;
        //if (!UseRenderTarget)
        {
            //engine takes the Ceiling so make sure its below or sometimes off by 1 rounding up from ActualWidth/Height
            element.Width = !UseRenderTarget ? Math.Floor((float)Math.Min(Window.Current.Bounds.Width, Width)) : (float)Math.Min(Window.Current.Bounds.Width, Width);
            element.Height = !UseRenderTarget ? Math.Floor((float)Math.Min(Window.Current.Bounds.Height, Height)) : (float)Math.Min(Window.Current.Bounds.Height, Height);
            //bool bHasCalledUpdateLayout = false;
            //should wrap into another event handler and check a bHasCalledUpdateLayout to ignore early calls and race condition
            //object lockVar = new object();
            //EventHandler<object> eventHandler = null;
            //System.Threading.Tasks.TaskCompletionSource<object> t = new System.Threading.Tasks.TaskCompletionSource<object>();
            //eventHandler = (sender, e) => { lock (lockVar) { if (bHasCalledUpdateLayout && Math.Abs(element.ActualWidth - element.Width) <= 1 && Math.Abs(element.ActualHeight - element.Height) <= 1) { lock (lockVar) { if (bHasCalledUpdateLayout) { bHasCalledUpdateLayout = false; t.SetResult(e); } } } } };
            //element.LayoutUpdated += eventHandler;
            //lock (lockVar) {
            //    element.UpdateLayout();
            //    bHasCalledUpdateLayout = true;
            //}
            ////await element.Dispatcher.RunAsync(Windows.UI.Core.CoreDispatcherPriority.Normal, new Windows.UI.Core.DispatchedHandler(() => element.Dispatcher.ProcessEvents(Windows.UI.Core.CoreProcessEventsOption.ProcessAllIfPresent)));
            //await t.Task;
            //element.LayoutUpdated -= eventHandler;
            await element.Dispatcher.RunAsync(Windows.UI.Core.CoreDispatcherPriority.Low, () => { });
            if (!UseRenderTarget && (element.ActualWidth > element.Width || element.ActualHeight > element.Height))
            {
                if (element.ActualWidth > element.Width) element.Width -= 1;
                if (element.ActualHeight > element.Height) element.Height -= 1;
                //bHasCalledUpdateLayout = false;
                //t = new System.Threading.Tasks.TaskCompletionSource<object>();
                //eventHandler = (sender, e) => { lock (lockVar) { if (bHasCalledUpdateLayout && Math.Abs(element.ActualWidth - element.Width) <= 1 && Math.Abs(element.ActualHeight - element.Height) <= 1) { lock (lockVar) { if (bHasCalledUpdateLayout) { bHasCalledUpdateLayout = false; t.SetResult(e); } } } } };
                //element.LayoutUpdated += eventHandler;
                //lock (lockVar)
                //{
                //    element.UpdateLayout();
                //    bHasCalledUpdateLayout = true;
                //}
                //await t.Task;
                //element.LayoutUpdated -= eventHandler;
                await element.Dispatcher.RunAsync(Windows.UI.Core.CoreDispatcherPriority.Low, () => { });
            }
        }
#if WINDOWS_APP && STORETOOLKIT
        if (!UseRenderTarget) {
            System.IO.MemoryStream memstream = await WinRTXamlToolkit.Composition.WriteableBitmapRenderExtensions.RenderToPngStream(element);
            Windows.Storage.StorageFile file = await Windows.Storage.ApplicationData.Current.LocalFolder.CreateFileAsync(fileName + ".png", Windows.Storage.CreationCollisionOption.ReplaceExisting);
            Windows.Storage.Streams.IRandomAccessStream stream = await file.OpenAsync(Windows.Storage.FileAccessMode.ReadWrite);
            await stream.WriteAsync(memstream.GetWindowsRuntimeBuffer());
            stream.Dispose();
        } else {
#endif
        //Canvas cvs = new Canvas();
        //cvs.Width = Width;
        //cvs.Height = Height;
        //Windows.UI.Xaml.Shapes.Path path = new Windows.UI.Xaml.Shapes.Path();
        //object val;
        //Resources.TryGetValue((object)"PathString", out val);
        //Binding b = new Binding
        //{
        //    Source = (string)val
        //};
        //BindingOperations.SetBinding(path, Windows.UI.Xaml.Shapes.Path.DataProperty, b);
        //cvs.Children.Add(path);
        float dpi = Windows.Graphics.Display.DisplayInformation.GetForCurrentView().LogicalDpi;
        Windows.UI.Xaml.Media.Imaging.RenderTargetBitmap wb = new Windows.UI.Xaml.Media.Imaging.RenderTargetBitmap();
        await wb.RenderAsync(element, (int)((float)Width * 96 / dpi), (int)((float)Height * 96 / dpi));
        Windows.Storage.Streams.IBuffer buf = await wb.GetPixelsAsync();
        Windows.Storage.StorageFile file = await Windows.Storage.ApplicationData.Current.LocalFolder.CreateFileAsync(fileName + ".png", Windows.Storage.CreationCollisionOption.ReplaceExisting);
        Windows.Storage.Streams.IRandomAccessStream stream = await file.OpenAsync(Windows.Storage.FileAccessMode.ReadWrite);
        //Windows.Graphics.Imaging.BitmapPropertySet propertySet = new Windows.Graphics.Imaging.BitmapPropertySet();
        //propertySet.Add("ImageQuality", new Windows.Graphics.Imaging.BitmapTypedValue(1.0, Windows.Foundation.PropertyType.Single)); // Maximum quality
        Windows.Graphics.Imaging.BitmapEncoder be = await Windows.Graphics.Imaging.BitmapEncoder.CreateAsync(Windows.Graphics.Imaging.BitmapEncoder.PngEncoderId, stream);//, propertySet);
        be.SetPixelData(Windows.Graphics.Imaging.BitmapPixelFormat.Bgra8, Windows.Graphics.Imaging.BitmapAlphaMode.Premultiplied, (uint)wb.PixelWidth, (uint)wb.PixelHeight, dpi, dpi, buf.ToArray());
        await be.FlushAsync();
        await stream.GetOutputStreamAt(0).FlushAsync();
        stream.Dispose();
#if WINDOWS_APP && STORETOOLKIT
        }
#endif
        //if (!UseRenderTarget)
        {

            element.Width = oldWidth;
            element.Height = oldHeight;
            //bHasCalledUpdateLayout = false;
            //t = new System.Threading.Tasks.TaskCompletionSource<object>();
            //eventHandler = (sender, e) => { lock (lockVar) { if (bHasCalledUpdateLayout && Math.Abs(element.ActualWidth - actOldWidth) <= 1 && Math.Abs(element.ActualHeight - actOldHeight) <= 1) { lock (lockVar) { if (bHasCalledUpdateLayout) { bHasCalledUpdateLayout = false; t.SetResult(e); } } } } };
            //element.LayoutUpdated += eventHandler;
            //lock (lockVar)
            //{
            //    element.UpdateLayout();
            //    bHasCalledUpdateLayout = true;
            //}
            //await t.Task;
            //element.LayoutUpdated -= eventHandler;
            await element.Dispatcher.RunAsync(Windows.UI.Core.CoreDispatcherPriority.Low, () => { });
        }
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
            this.DataContext = this;
            this.ViewModel = new MyTabViewModel();
            UIChanger = new MyUIChanger();
            this.InitializeComponent();
#if STORETOOLKIT
            AppBarButton RenderButton = new AppBarButton() { Icon = new SymbolIcon(Symbol.Camera), Label = new Windows.ApplicationModel.Resources.ResourceLoader().GetString("Render/Label") };
            RenderButton.Click += RenderPngs_Click;
            (this.BottomAppBar as CommandBar).PrimaryCommands.Add(RenderButton);
#endif
#if WINDOWS_PHONE_APP
            this.NavigationCacheMode = NavigationCacheMode.Required;
#endif
            gestRec = new Windows.UI.Input.GestureRecognizer();
            gestRec.GestureSettings = Windows.UI.Input.GestureSettings.HoldWithMouse | Windows.UI.Input.GestureSettings.Hold | Windows.UI.Input.GestureSettings.Tap | Windows.UI.Input.GestureSettings.DoubleTap | Windows.UI.Input.GestureSettings.RightTap;
            gestRec.Holding += OnHolding;
            gestRec.RightTapped += OnRightTapped;
            gestRec.Tapped += OnTapped;
        }
        private Windows.UI.Input.GestureRecognizer gestRec;
        private object _holdObj;
        void OnPointerReleased(object sender, PointerRoutedEventArgs e)
        {
            var ps = e.GetIntermediatePoints(null);
            if (ps != null && ps.Count > 0)
            {
                gestRec.ProcessUpEvent(ps[0]);
                e.Handled = true;
                gestRec.CompleteGesture();
                _holdObj = null;
            }
        }
        void OnPointerPressed(object sender, PointerRoutedEventArgs e)
        {
            var ps = e.GetIntermediatePoints(null);
            if (ps != null && ps.Count > 0)
            {
                _holdObj = sender;
                gestRec.ProcessDownEvent(ps[0]);
                e.Handled = true;
            }
        }
        void OnPointerMoved(object sender, PointerRoutedEventArgs e)
        {
            _holdObj = sender;
            gestRec.ProcessMoveEvents(e.GetIntermediatePoints(null));
            e.Handled = true;
        }
        void OnPointerCanceled(object sender, PointerRoutedEventArgs e)
        {
            gestRec.CompleteGesture();
            e.Handled = true;
        }
        void OnPointerCaptureLost(object sender, PointerRoutedEventArgs e)
        {
            gestRec.CompleteGesture();
            e.Handled = true;
        }
        void OnPointerExited(object sender, PointerRoutedEventArgs e)
        {
            gestRec.CompleteGesture();
            e.Handled = true;
        }
        void OnHolding(Windows.UI.Input.GestureRecognizer sender, Windows.UI.Input.HoldingEventArgs args)
        {
            if (args.HoldingState == Windows.UI.Input.HoldingState.Started)
            {
                if (_holdObj != null) DoHolding(_holdObj);
                gestRec.CompleteGesture();
            }
        }
        void OnRightTapped(Windows.UI.Input.GestureRecognizer sender, Windows.UI.Input.RightTappedEventArgs args)
        {
            if (_holdObj != null) DoHolding(_holdObj);
        }
        void OnTapped(Windows.UI.Input.GestureRecognizer sender, Windows.UI.Input.TappedEventArgs args)
        {
            if (_holdObj != null) {
                sectionListBox.SelectedItem = (_holdObj as TextBlock).DataContext;
                if (args.TapCount >= 2) {
                    sectionListBox_DoubleTapped(null, null);
                }
            }
        }
        protected override async void OnNavigatedTo(NavigationEventArgs e)
        {
            List<MyTabItem> items = null;
            System.Threading.Tasks.Task t = new System.Threading.Tasks.Task(() =>
            {
                items = System.Linq.Enumerable.Select(IslamMetadata.TanzilReader.GetDivisionTypes(), (Arr, idx) => new MyTabItem { Title = Arr, Index = idx + 1 }).ToList();
                items.Insert(0, new MyTabItem { IsBookmarks = true, Title = new Windows.ApplicationModel.Resources.ResourceLoader().GetString("Bookmarks/Text"), Index = 0 });
            });
            t.Start();
            await t;
            ViewModel.Items = items;
            LoadingRing.IsActive = false;
        }
        public MyUIChanger UIChanger { get; set; }
        public MyTabViewModel ViewModel { get; set; }

        private void sectionListBox_DoubleTapped(object sender, DoubleTappedRoutedEventArgs e)
        {
            if (ViewModel.ListSelectedItem == null) return;
            this.Frame.Navigate(typeof(WordForWordUC), ViewModel.SelectedItem.IsBookmarks ? new { Division = AppSettings.Bookmarks[ViewModel.ListSelectedItem.Index][0], Selection = AppSettings.Bookmarks[ViewModel.ListSelectedItem.Index][1], JumpToChapter = AppSettings.Bookmarks[ViewModel.ListSelectedItem.Index][2], JumpToVerse = AppSettings.Bookmarks[ViewModel.ListSelectedItem.Index][3], StartPlaying = false } : new {Division = ViewModel.SelectedItem.Index - 1, Selection = ViewModel.ListSelectedItem.Index, JumpToChapter = -1, JumpToVerse = -1, StartPlaying = false });
        }

        private void RenderPngs_Click(object sender, RoutedEventArgs e)
        {
            this.Frame.Navigate(typeof(ExtSplashScreen));
        }
        private void Settings_Click(object sender, RoutedEventArgs e)
        {
            this.Frame.Navigate(typeof(Settings));
        }
        private void About_Click(object sender, RoutedEventArgs e)
        {
            this.Frame.Navigate(typeof(About));
        }
        private void RemoveBookmark_Click(object sender, RoutedEventArgs e)
        {
            List<int[]> marks = AppSettings.Bookmarks.ToList();
            marks.RemoveAt(((sender as MenuFlyoutItem).DataContext as MyListItem).Index);
            AppSettings.Bookmarks = marks.ToArray();
            ViewModel.SelectedItem.RefreshItems();
            ViewModel.ListItems = ViewModel.SelectedItem.Items;
        }

        private void TextBlock_Holding(object sender, HoldingRoutedEventArgs e)
        {
            DoHolding(sender);
            e.Handled = true;
        }
        private void DoHolding(object sender)
        {
            if (ViewModel.SelectedItem.IsBookmarks)
            {
                FlyoutBase.ShowAttachedFlyout(sender as TextBlock);
            }
        }
    }
    public static class ListFormattedTextBehavior
    {
        #region FormattedText Attached dependency property

        public static List<string> GetFormattedText(DependencyObject obj)
        {
            return (List<string>)obj.GetValue(FormattedTextProperty);
        }

        public static void SetFormattedText(DependencyObject obj, List<string> value)
        {
            obj.SetValue(FormattedTextProperty, value);
        }

        public static readonly DependencyProperty FormattedTextProperty =
            DependencyProperty.RegisterAttached("FormattedText",
            typeof(List<string>),
            typeof(ListFormattedTextBehavior),
            new PropertyMetadata(null, FormattedTextChanged));

        private static void FormattedTextChanged(DependencyObject sender, DependencyPropertyChangedEventArgs e)
        {
            List<string> value = e.NewValue as List<string>;

            TextBlock textBlock = sender as TextBlock;

            if (textBlock != null)
            {
                textBlock.Inlines.Clear();
                textBlock.Inlines.Add(new Windows.UI.Xaml.Documents.Run() { Text = value[0], FontFamily = new FontFamily(AppSettings.strOtherSelectedFont), FontSize = AppSettings.dOtherFontSize });
                if (value.Count > 1)
                {
                    textBlock.Inlines.Add(new Windows.UI.Xaml.Documents.Run() { Text = "(" + value[1] + ")", FontWeight = Windows.UI.Text.FontWeights.Bold, FlowDirection = FlowDirection.RightToLeft, FontFamily = new FontFamily(AppSettings.strSelectedFont), FontSize = AppSettings.dFontSize });
                    textBlock.Inlines.Add(new Windows.UI.Xaml.Documents.Run() { Text = value[2], FontFamily = new FontFamily(AppSettings.strOtherSelectedFont), FontSize = AppSettings.dOtherFontSize });
                }
            }
        }

        #endregion
    }
    public class MyTabViewModel : INotifyPropertyChanged
    {
        public MyTabViewModel()
        {
        }
        private List<MyTabItem> _Items;
        public IEnumerable<MyTabItem> Items { get
            {
                return _Items;
            }
            set
            {
                _Items = value.ToList();
                int iDefault = AppSettings.iDefaultStartTab;
                PropertyChanged(this, new PropertyChangedEventArgs("Items"));
                SelectedItem = _Items.ElementAt(iDefault);
            }
        }

        private List<MyListItem> _ListItems;
        public IEnumerable<MyListItem> ListItems
        {
            get
            {
                return _ListItems;
            }
            set
            {
                _ListItems = value.ToList();
                if (PropertyChanged != null) PropertyChanged(this, new PropertyChangedEventArgs("ListItems"));
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
                AppSettings.iDefaultStartTab = _selectedItem.Index;
                PropertyChanged(this, new PropertyChangedEventArgs("SelectedItem"));
                ListItems = _selectedItem.Items;
                ListSelectedItem = ListItems.Count() == 0 ? null : ListItems.First();
            }
        }

        #region Implementation of INotifyPropertyChanged

        public event PropertyChangedEventHandler PropertyChanged;

        #endregion
    }

    public class MyTabItem
    {
        public bool IsBookmarks { get; set; }
        public string Title { get; set; }
        public int Index { get; set; }
        private IEnumerable<MyListItem> _Items;
        public IEnumerable<MyListItem> Items
        {
            get
            {
                if (_Items == null) {
                    if (IsBookmarks) {
                        _Items = System.Linq.Enumerable.Select(AppSettings.Bookmarks, (Bookmark, Idx) => new MyListItem { TextItems = new List<string>() { IslamMetadata.TanzilReader.GetSelectionName(Bookmark[0], Bookmark[1], XMLRender.ArabicData.TranslitScheme.RuleBased, String.Empty) + " " + Bookmark[2].ToString() + ":" + Bookmark[3].ToString() }, Index = Idx });
                    } else {
                        _Items = System.Linq.Enumerable.Select(IslamMetadata.TanzilReader.GetSelectionNames((Index - 1).ToString(), XMLRender.ArabicData.TranslitScheme.RuleBased, String.Empty), (Arr, Idx) => new MyListItem { TextItems = new List<string>(((string)(Arr.Cast<object>()).First()).Split('(', ')')), Index = (int)(Arr.Cast<object>()).Last() });
                    }
                }
                return _Items;
            }
        }
        public void RefreshItems() { _Items = null; }

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
