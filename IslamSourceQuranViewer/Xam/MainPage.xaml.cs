using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.Collections;
using System.Diagnostics;
using System.Windows.Input;
using Xamarin.Forms;

#if WINDOWS_PHONE
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
        Stream rc = null;
        rc = IslamSourceQuranViewer.Xam.WinPhone.Resources.AppResources.ResourceManager.GetStream(FilePath.Replace(Windows.ApplicationModel.Package.Current.InstalledLocation.Path, "ms-resource:///Files").Replace("\\", "/"));
        System.Threading.Tasks.Task<Windows.Storage.StorageFile> t;
        if (rc != null)
        {
            return rc;
        }
        else {
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
#if WINDOWS_PHONE
        System.Threading.Tasks.Task<IReadOnlyList<Windows.Storage.IStorageItem>> files = t.Result.GetItemsAsync().AsTask();
        files.Wait();
        return files.Result.FirstOrDefault(p => p.Name == Path) != null;
#else
        System.Threading.Tasks.Task<Windows.Storage.IStorageItem> tn = t.Result.GetItemsAsync();
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
        return new System.Resources.ResourceManager(baseName + ".Resources", System.Reflection.Assembly.Load(baseName)).GetString(resourceKey, System.Threading.Thread.CurrentThread.CurrentUICulture);
    }
    /*public static async System.Threading.Tasks.Task SavePathImageAsFile(int Width, int Height, string fileName, VisualElement element, bool UseRenderTarget = true)
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
    }*/
}
#else
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
#endif

namespace IslamSourceQuranViewer.Xam
{
    public class ItemsView : Grid
    {
        protected StackLayout PagingStackLayout;
        protected ScrollView ScrollView;
        protected readonly ICommand SelectedCommand;
        protected readonly StackLayout ItemsStackLayout;
        protected Boolean Wait = false;


        private Boolean _isScrollAutomaticInitialized;

        public ItemsView()
        {
            ScrollView = new ScrollView
            {
                Orientation = ScrollOrientation.Horizontal
            };

            ScrollView.Scrolled += ScrollView_Scrolled;

            ItemsStackLayout = new StackLayout
            {
                Orientation = StackOrientation.Horizontal,
                Padding = new Thickness(0),
                Spacing = 0,
                HorizontalOptions = LayoutOptions.FillAndExpand
            };

            ScrollView.Content = ItemsStackLayout;
            Children.Add(ScrollView);

            PagingStackLayout = new StackLayout()
            {
                HorizontalOptions = LayoutOptions.Center,
                Orientation = StackOrientation.Horizontal,
                VerticalOptions = LayoutOptions.End,
                Padding = Device.OnPlatform<Thickness>(
                new Thickness(0, 0, 0, 36),
                new Thickness(0, 0, 0, 36),
                new Thickness(0, 0, 0, 60)),
                Opacity = 0.5
            };
            Children.Add(PagingStackLayout);

            var leftArrow = new Image()
            {
                // Replace with your own arrow image
                Source = ImageSource.FromResource("IslamSourceQuranViewer.Xam.Images.ItemsView.LeftArrow.png"),
                HorizontalOptions = LayoutOptions.Start,
                VerticalOptions = LayoutOptions.Center,
                WidthRequest = 50,
            };
            leftArrow.GestureRecognizers.Add(new TapGestureRecognizer
            {
                Command = new Command(async () =>
                {
                    this.ActualElementIndex--;
                    Wait = true;
                    await ScrollToActualAsync();
                })
            });
            Children.Add(leftArrow);

            var rightArrow = new Image()
            {
                // Replace with your own arrow image
                Source = ImageSource.FromResource("IslamSourceQuranViewer.Xam.Images.ItemsView.RightArrow.png"),
                HorizontalOptions = LayoutOptions.End,
                VerticalOptions = LayoutOptions.Center,
                WidthRequest = 50
            };
            rightArrow.GestureRecognizers.Add(new TapGestureRecognizer
            {
                Command = new Command(async () =>
                {
                    this.ActualElementIndex++;
                    Wait = true;
                    await ScrollToActualAsync();
                })
            });
            Children.Add(rightArrow);

            SelectedCommand = new Command<object>(item =>
            {
                var selectable = item as ISelectable;
                if (selectable == null) return;

                SetSelected(selectable);
                SelectedItem = selectable.IsSelected ? selectable : null;
            });

            PropertyChanged += (sender, e) =>
            {
                if (e.PropertyName == "Orientation")
                {
                    ItemsStackLayout.Orientation = ScrollView.Orientation == ScrollOrientation.Horizontal ? StackOrientation.Horizontal : StackOrientation.Vertical;
                }

            };
        }


        public int ItemsCount
        {
            get { return this.ItemsStackLayout.Children.Count; }
        }

        protected virtual void SetSelected(ISelectable selectable)
        {
            selectable.IsSelected = true;
        }

        public View ActualElement
        {
            get
            {
                return ItemsStackLayout.Children[ActualElementIndex];
            }
        }
        public int ActualElementIndex { get; set; }

        public bool ScrollToStartOnSelected { get; set; }

        public event EventHandler SelectedItemChanged;

        public static readonly BindableProperty ItemsSourceProperty =
            BindableProperty.Create<ItemsView, IEnumerable>(p => p.ItemsSource, default(IEnumerable<object>), BindingMode.TwoWay, null, ItemsSourceChanged);

        public IEnumerable ItemsSource
        {
            get { return (IEnumerable)GetValue(ItemsSourceProperty); }
            set { SetValue(ItemsSourceProperty, value); }
        }

        public static readonly BindableProperty SelectedItemProperty =
            BindableProperty.Create<ItemsView, object>(p => p.SelectedItem, default(object), BindingMode.TwoWay, null, OnSelectedItemChanged);

        public object SelectedItem
        {
            get { return (object)GetValue(SelectedItemProperty); }
            set { SetValue(SelectedItemProperty, value); }
        }

        public static readonly BindableProperty ItemTemplateProperty =
            BindableProperty.Create<ItemsView, DataTemplate>(p => p.ItemTemplate, default(DataTemplate));

        public DataTemplate ItemTemplate
        {
            get { return (DataTemplate)GetValue(ItemTemplateProperty); }
            set { SetValue(ItemTemplateProperty, value); }
        }

        private static void ItemsSourceChanged(BindableObject bindable, IEnumerable oldValue, IEnumerable newValue)
        {
            var itemsLayout = (ItemsView)bindable;
            itemsLayout.SetItems();
            itemsLayout.SetPagination();
            itemsLayout.ScrollAutomaticAsync();

        }

        protected virtual void SetItems()
        {
            ItemsStackLayout.Children.Clear();

            if (ItemsSource == null)
                return;

            foreach (var item in ItemsSource)
                ItemsStackLayout.Children.Add(GetItemView(item));

            SelectedItem = ItemsSource.OfType<ISelectable>().FirstOrDefault(x => x.IsSelected);
        }

        protected virtual View GetItemView(object item)
        {
            var content = ItemTemplate.CreateContent();
            var view = content as View;
            if (view == null) return null;

            view.BindingContext = item;

            var gesture = new TapGestureRecognizer
            {
                Command = SelectedCommand,
                CommandParameter = item
            };

            AddGesture(view, gesture);

            return view;
        }

        protected void AddGesture(View view, TapGestureRecognizer gesture)
        {
            view.GestureRecognizers.Add(gesture);

            var layout = view as Layout<View>;

            if (layout == null)
                return;

            foreach (var child in layout.Children)
                AddGesture(child, gesture);
        }

        private static void OnSelectedItemChanged(BindableObject bindable, object oldValue, object newValue)
        {
            var itemsView = (ItemsView)bindable;
            if (newValue == oldValue)
                return;

            var selectable = newValue as ISelectable;
            itemsView.SetSelectedItem(selectable ?? oldValue as ISelectable);
        }

        protected virtual void SetSelectedItem(ISelectable selectedItem)
        {
            var items = ItemsSource;

            foreach (var item in items.OfType<ISelectable>())
                item.IsSelected = selectedItem != null && item == selectedItem && selectedItem.IsSelected;

            var handler = SelectedItemChanged;
            if (handler != null)
                handler(this, EventArgs.Empty);
        }

        protected virtual async void ScrollAutomaticAsync()
        {
            while (!_isScrollAutomaticInitialized)
            {
                _isScrollAutomaticInitialized = true;
                if (Wait)
                {
                    Wait = false;
                    await Task.Delay(5000);
                }
                SetActivePage();
                await Task.Delay(5000);
                this.ActualElementIndex++;
                await ScrollToActualAsync();
                _isScrollAutomaticInitialized = false;

            }
        }

        private async Task ScrollToActualAsync()
        {
            if (this.ActualElementIndex == this.ItemsCount)
                this.ActualElementIndex = 0;

            if (this.ActualElementIndex < 0)
                this.ActualElementIndex = 0;

            try
            {
                await this.ScrollView.ScrollToAsync(this.ActualElement.X, 0, false);
            }
            catch
            {
                //invalid scroll: sometimes happen
            }
        }

        protected virtual void SetPagination()
        {
            this.ActualElementIndex = 0;
            PagingStackLayout.Children.Clear();
            for (int i = 0; i < this.ItemsCount; i++)
            {
                var view = new BoxView() { BackgroundColor = Color.White, WidthRequest = 10, HeightRequest = 10 };
                PagingStackLayout.Children.Add(view);
            }
        }

        protected virtual void SetActivePage()
        {
            try
            {
                for (int i = 0; i < this.ItemsCount; i++)
                {
                    (PagingStackLayout.Children[i] as BoxView).BackgroundColor = Color.White;
                }
                (PagingStackLayout.Children[this.ActualElementIndex] as BoxView).BackgroundColor = Color.Red;

            }
            catch { }
        }

        void ScrollView_Scrolled(object sender, Xamarin.Forms.ScrolledEventArgs e)
        {
            if (e.ScrollX % ItemsStackLayout.Children.First().Width != 0 && !Wait)
                Wait = true;

            for (int i = 1; i < this.ItemsCount; i++)
            {
                var previousItemX = ItemsStackLayout.Children[i - 1].X;
                var actualItemX = i == this.ItemsCount ? this.ItemsCount : ItemsStackLayout.Children[i].X;

                if (e.ScrollX >= previousItemX && e.ScrollX <= actualItemX)
                {
                    this.ActualElementIndex = e.ScrollX == actualItemX ? i : i - 1;
                    SetActivePage();
                }
            }
        }
    }

    public interface ISelectable
    {
        bool IsSelected { get; set; }

        ICommand SelectCommand { get; set; }
    }

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
