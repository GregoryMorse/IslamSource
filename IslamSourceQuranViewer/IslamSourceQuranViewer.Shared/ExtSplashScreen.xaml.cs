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
    public sealed partial class ExtSplashScreen : Page
    {
        internal Rect splashImageRect; // Rect to store splash screen image coordinates.
        private Windows.ApplicationModel.Activation.SplashScreen splash; // Variable to hold the splash screen object.
        internal bool dismissed = false; // Variable to track splash screen dismissal status.
        internal Frame rootFrame;

        public ExtSplashScreen()
        {
            this.InitializeComponent();
            this.BottomAppBar.Visibility = Windows.UI.Xaml.Visibility.Visible;
            this.ProgressGrid.Visibility = Windows.UI.Xaml.Visibility.Collapsed;
            Position();
            LayoutUpdated += SplashScreen_LayoutUpdated;
#if WINDOWS_APP
            AppBarButton BackButton = new AppBarButton() { Icon = new SymbolIcon(Symbol.Back), Label = new Windows.ApplicationModel.Resources.ResourceLoader().GetString("Back/Label") };
            BackButton.Click += Back_Click;
            (this.BottomAppBar as CommandBar).PrimaryCommands.Add(BackButton);
#endif
#if STORETOOLKIT
            AppBarButton RenderButton = new AppBarButton() { Icon = new SymbolIcon(Symbol.Camera), Label = new Windows.ApplicationModel.Resources.ResourceLoader().GetString("Render/Label") };
            RenderButton.Click += RenderPngs_Click;
            (this.BottomAppBar as CommandBar).PrimaryCommands.Add(RenderButton);
#endif
        }
        public ExtSplashScreen(Windows.ApplicationModel.Activation.SplashScreen splashScreen, bool loadState)
        {
            this.InitializeComponent();
            this.BottomAppBar.Visibility = Windows.UI.Xaml.Visibility.Collapsed;
            this.ProgressGrid.Visibility = Windows.UI.Xaml.Visibility.Visible;
            // Listen for window resize events to reposition the extended splash screen image accordingly.
            // This ensures that the extended splash screen formats properly in response to window resizing.
            Window.Current.SizeChanged += new WindowSizeChangedEventHandler(SplashScreen_OnResize);
            LayoutUpdated += SplashScreen_LayoutUpdated;
            splash = splashScreen;
            if (splash != null)
            {
                // Register an event handler to be executed when the splash screen has been dismissed.
                splash.Dismissed += new TypedEventHandler<Windows.ApplicationModel.Activation.SplashScreen, Object>(DismissedEventHandler);

                // Retrieve the window coordinates of the splash screen image.
                splashImageRect = splash.ImageLocation;
                Position();
            }

            // Create a Frame to act as the navigation context 
            rootFrame = new Frame();

            // Restore the saved session state if necessary
            //RestoreStateAsync(loadState);
        }
        void Position()
        {
            //MainGrid.SetValue(Canvas.LeftProperty, splashImageRect.X);
            //MainGrid.SetValue(Canvas.TopProperty, splashImageRect.Y);
            //MainGrid.Height = splashImageRect.Height;
            //MainGrid.Width = splashImageRect.Width;
            //MainGrid.SetValue(Canvas.LeftProperty, Window.Current.Bounds.Left);
            //MainGrid.SetValue(Canvas.TopProperty, Window.Current.Bounds.Top);
            //MainGrid.Height = Window.Current.Bounds.Height;
            //MainGrid.Width = Window.Current.Bounds.Width;
            if (Window.Current.Bounds.Width < Window.Current.Bounds.Height * .94 / 746 * 502.655)
            {
                //MainGrid.ColumnDefinitions[1].Width = new GridLength(Window.Current.Bounds.Width - 2);
                SplashImg.Width = Window.Current.Bounds.Width - 2;
                ProgressGrid.Height = SplashImg.Height = SplashImg.Width / 502.655 * 746;
            }
            else
            {
                SplashImg.Width = Window.Current.Bounds.Height * .94 / 746 * 502.655;
                ProgressGrid.Height = SplashImg.Height = Window.Current.Bounds.Height * .94;
            }
        }
        void SplashScreen_LayoutUpdated(object sender, object e)
        {
            if (splashProgressRing.Height != ProgressGrid.RowDefinitions[1].ActualHeight)
            {
                splashProgressRing.Height = ProgressGrid.RowDefinitions[1].ActualHeight;
                splashProgressRing.Width = ProgressGrid.RowDefinitions[1].ActualHeight;
            }
            if (double.IsNaN(MainGrid.Width) || double.IsNaN(MainGrid.Height)) { Position(); return; }
            if (Math.Min(Window.Current.Bounds.Width, MainGrid.Width) < Math.Min(Window.Current.Bounds.Height, MainGrid.Height) * .94 / 746 * 502.655)
            {
                SplashImg.Width = Math.Min(Window.Current.Bounds.Width, MainGrid.Width) - 2;
                ProgressGrid.Height = SplashImg.Height = SplashImg.Width / 502.655 * 746;
            }
            else
            {
                SplashImg.Width = Math.Min(Window.Current.Bounds.Height, MainGrid.Height) * .94 / 746 * 502.655;
                ProgressGrid.Height = SplashImg.Height = Math.Min(Window.Current.Bounds.Height, MainGrid.Height) * .94;
            }
        }
        void SplashScreen_OnResize(Object sender, Windows.UI.Core.WindowSizeChangedEventArgs e)
        {
            // Safely update the extended splash screen image coordinates. This function will be executed when a user resizes the window.
            if (splash != null)
            {
                // Update the coordinates of the splash screen image.
                splashImageRect = splash.ImageLocation;
                Position();

                // If applicable, include a method for positioning a progress control.
                // PositionRing();
            }
        }

        // Include code to be executed when the system has transitioned from the splash screen to the extended splash screen (application's first view).
        async void DismissedEventHandler(Windows.ApplicationModel.Activation.SplashScreen sender, object e)
        {
            dismissed = true;

            // Complete app setup operations here...
            await DismissExtendedSplash();
        }
        async System.Threading.Tasks.Task DismissExtendedSplash()
        {
            await Windows.ApplicationModel.Core.CoreApplication.MainView.CoreWindow.Dispatcher.RunAsync(Windows.UI.Core.CoreDispatcherPriority.Normal,
                () =>
                {
                    rootFrame.Navigate(typeof(MainPage));
                    Window.Current.Content = rootFrame;
                });
            //Navigate to mainpage
            // Place the frame in the current Window

        }
        async System.Threading.Tasks.Task RestoreStateAsync(bool loadState)
        {
            if (loadState)
            {
                await System.Threading.Tasks.Task.Run(()=>loadState);
                // code to load your app's state here 
            }
        }
        private async void RenderPngs_Click(object sender, RoutedEventArgs e)
        {
            object[,] Win81PhoneLogos = new object[,] {{1152, 1920, "SplashScreen.scale-240"}, {672, 1120, "SplashScreen.scale-140"}, {480, 800, "SplashScreen.scale-100"},
                        {58, 58, "BadgeLogo.scale-240"}, {33, 33, "BadgeLogo.scale-140"}, {24, 24, "BadgeLogo.scale-100"},
                        {120, 120, "StoreLogo.scale-240"}, {70, 70, "StoreLogo.scale-140"}, {50, 50, "StoreLogo.scale-100"},
                        {106, 106, "Square44x44Logo.scale-240"}, {62, 62, "Square44x44Logo.scale-140"}, {44, 44, "Square44x44Logo.scale-100"},
                        {744, 360, "Wide310x150Logo.scale-240"}, {434, 210, "Wide310x150Logo.scale-140"}, {310, 150, "Wide310x150Logo.scale-100"},
                        {360, 360, "Square150x150Logo.scale-240"}, {210, 210, "Square150x150Logo.scale-140"}, {150, 150, "Square150x150Logo.scale-100"},
                        {170, 170, "Square71x71Logo.scale-240"}, {99, 99, "Square71x71Logo.scale-140"}, {71, 71, "Square71x71Logo.scale-100"}};
            object[,] Win8Logos = new object[,] {{1116, 540, "splashscreen.scale-180"}, {868, 420, "splashscreen.scale-140"}, {620, 300, "splashscreen.scale-100"},
                        {43, 43, "badgelogo.scale-180"}, {33, 33, "badgelogo.scale-140"}, {24, 24, "badgelogo.scale-100"},
                        {90, 90, "storelogo.scale-180"}, {70, 70, "storelogo.scale-140"}, {50, 50, "storelogo.scale-100"},
                        {54, 54, "smalllogo.scale-180"}, {42, 42, "smalllogo.scale-140"}, {30, 30, "smalllogo.scale-100"}, {24, 24, "smalllogo.scale-80"},
                        {256, 256, "smalllogo.targetsize-256"}, {48, 48, "smalllogo.targetsize-48"}, {32, 32, "smalllogo.targetsize-32"}, {16, 16, "smalllogo.targetsize-16"},
                        {558, 270, "widelogo.scale-180"}, {434, 210, "widelogo.scale-140"}, {310, 150, "widelogo.scale-100"}, {248, 120, "widelogo.scale-80"},
                        {270, 270, "logo.scale-180"}, {210, 210, "logo.scale-140"}, {150, 150, "logo.scale-100"}, {120, 120, "logo.scale-80"}};
            object[,] Win81Logos = new object[,] {{1116, 540, "SplashScreen.scale-180"}, {868, 420, "SplashScreen.scale-140"}, {620, 300, "SplashScreen.scale-100"},
                        {43, 43, "BadgeLogo.scale-180"}, {33, 33, "BadgeLogo.scale-140"}, {24, 24, "BadgeLogo.scale-100"},
                        {90, 90, "StoreLogo.scale-180"}, {70, 70, "StoreLogo.scale-140"}, {50, 50, "StoreLogo.scale-100"},
                        {54, 54, "SmallLogo.scale-180"}, {42, 42, "SmallLogo.scale-140"}, {30, 30, "SmallLogo.scale-100"}, {24, 24, "SmallLogo.scale-80"},
                        {256, 256, "SmallLogo.targetsize-256"}, {48, 48, "SmallLogo.targetsize-48"}, {32, 32, "SmallLogo.targetsize-32"}, {16, 16, "SmallLogo.targetsize-16"},
                        {558, 558, "Square310x310Logo.scale-180"}, {434, 434, "Square310x310Logo.scale-140"}, {310, 310, "Square310x310Logo.scale-100"}, {248, 248, "Square310x310Logo.scale-80"},
                        {558, 270, "Wide310x150Logo.scale-180"}, {434, 210, "Wide310x150Logo.scale-140"}, {310, 150, "Wide310x150Logo.scale-100"}, {248, 120, "Wide310x150Logo.scale-80"},
                        {270, 270, "Square150x150Logo.scale-180"}, {210, 210, "Square150x150Logo.scale-140"}, {150, 150, "Square150x150Logo.scale-100"}, {120, 120, "Square150x150Logo.scale-80"},
                        {126, 126, "Square70x70Logo.scale-180"}, {98, 98, "Square70x70Logo.scale-140"}, {70, 70, "Square70x70Logo.scale-100"}, {56, 56, "Square70x70Logo.scale-80"}};
            object[,] WinUniversalLogos = new object[,] {{2480, 1200, "SplashScreen.scale-400"}, {1240, 600, "SplashScreen.scale-200"}, {930, 450, "SplashScreen.scale-150"}, {775, 375, "SplashScreen.scale-125"}, {620, 300, "SplashScreen.scale-100"},
                        {96, 96, "LockScreenLogo.scale-400"}, {48, 48, "LockScreenLogo.scale-200"}, {36, 36, "LockScreenLogo.scale-150"}, {30, 30, "LockScreenLogo.scale-125"}, {24, 24, "LockScreenLogo.scale-100"},
                        {200, 200, "StoreLogo.scale-400"}, {100, 100, "StoreLogo.scale-200"}, {75, 75, "StoreLogo.scale-150"}, {63, 63, "StoreLogo.scale-125"}, {50, 50, "StoreLogo.scale-100"},
                        {176, 176, "Square44x44Logo.scale-400"}, {88, 88, "Square44x44Logo.scale-200"}, {44, 44, "Square44x44Logo.scale-100"}, {66, 66, "Square44x44Logo.scale-150"}, {55, 55, "Square44x44Logo.scale-125"},
                        {256, 256, "Square44x44Logo.targetsize-256"}, {48, 48, "Square44x44Logo.targetsize-48"}, {24, 24, "Square44x44Logo.targetsize-24"}, {16, 16, "Square44x44Logo.targetsize-16"},
                        {256, 256, "Square44x44Logo.targetsize-256_altform-unplated"}, {48, 48, "Square44x44Logo.targetsize-48_altform-unplated"}, {24, 24, "Square44x44Logo.targetsize-24_altform-unplated"}, {16, 16, "Square44x44Logo.targetsize-16_altform-unplated"},
                        {1240, 1240, "Square310x310Logo.scale-400"}, {620, 620, "Square310x310Logo.scale-200"}, {310, 310, "Square310x310Logo.scale-100"}, {465, 465, "Square310x310Logo.scale-150"}, {388, 388, "Square310x310Logo.scale-125"},
                        {1240, 600, "Wide310x150Logo.scale-400"}, {620, 300, "Wide310x150Logo.scale-200"}, {310, 150, "Wide310x150Logo.scale-100"}, {465, 225, "Wide310x150Logo.scale-150"}, {388, 188, "Wide310x150Logo.scale-125"},
                        {600, 600, "Square150x150Logo.scale-400"}, {300, 300, "Square150x150Logo.scale-200"}, {150, 150, "Square150x150Logo.scale-100"}, {225, 225, "Square150x150Logo.scale-150"}, {188, 188, "Square150x150Logo.scale-125"},
                        {284, 284, "Square71x71Logo.scale-400"}, {142, 142, "Square71x71Logo.scale-200"}, {71, 71, "Square71x71Logo.scale-100"}, {107, 107, "Square71x71Logo.scale-150"}, {89, 89, "Square71x71Logo.scale-125"}};
            object[,] AppStoreLogos = new object[,] {{846, 468, "appstorepromotional-846x468"}, {558, 756, "appstorepromotional-558x756"}, {414, 468, "appstorepromotional-414x468"},
                        {414, 180, "appstorepromotional-414x180"}, {558, 558, "appstorepromotional-558x558"}, {2400, 1200, "appstorepromotional-2400x1200"},
                        {300, 300, "appstorephonetitleicon-300x300"}, {1000, 800, "appstorephonepromotional-1000x800"}, {358, 358, "appstorephonepromotional-358x358"}, {358, 173, "appstorephonepromotional-358x173"}};
            object[,] OldWin8 = new object[,] { { 360, 360, "logo.scale-240" }, { 106, 106, "smalllogo.scale-240" }, { 744, 360, "widelogo.scale-240" } };
            await Windows.Storage.ApplicationData.Current.LocalFolder.CreateFolderAsync("win8", Windows.Storage.CreationCollisionOption.OpenIfExists);
            await Windows.Storage.ApplicationData.Current.LocalFolder.CreateFolderAsync("win81phone", Windows.Storage.CreationCollisionOption.OpenIfExists);
            await Windows.Storage.ApplicationData.Current.LocalFolder.CreateFolderAsync("win81", Windows.Storage.CreationCollisionOption.OpenIfExists);
            await Windows.Storage.ApplicationData.Current.LocalFolder.CreateFolderAsync("winuniversal", Windows.Storage.CreationCollisionOption.OpenIfExists);
            await Windows.Storage.ApplicationData.Current.LocalFolder.CreateFolderAsync("appstore", Windows.Storage.CreationCollisionOption.OpenIfExists);
            for (int count = 0; count <= Win8Logos.GetUpperBound(0); count++)
            {
                await SavePathImageAsFile((int)Win8Logos[count, 0], (int)Win8Logos[count, 1], "win8\\" + (string)Win8Logos[count, 2], MainGrid);
            }
            for (int count = 0; count <= Win81PhoneLogos.GetUpperBound(0); count++) {
                await SavePathImageAsFile((int)Win81PhoneLogos[count, 0], (int)Win81PhoneLogos[count, 1], "win81phone\\" + (string)Win81PhoneLogos[count, 2], MainGrid);
            }
            for (int count = 0; count <= Win81Logos.GetUpperBound(0); count++)
            {
                await SavePathImageAsFile((int)Win81Logos[count, 0], (int)Win81Logos[count, 1], "win81\\" + (string)Win81Logos[count, 2], MainGrid);
            }
            for (int count = 0; count <= WinUniversalLogos.GetUpperBound(0); count++)
            {
                await SavePathImageAsFile((int)WinUniversalLogos[count, 0], (int)WinUniversalLogos[count, 1], "winuniversal\\" + (string)WinUniversalLogos[count, 2], MainGrid);
            }
            for (int count = 0; count <= AppStoreLogos.GetUpperBound(0); count++)
            {
                await SavePathImageAsFile((int)AppStoreLogos[count, 0], (int)AppStoreLogos[count, 1], "appstore\\" + (string)AppStoreLogos[count, 2], MainGrid);
            }
            GC.Collect(); //causes streams and files to properly close 
        }
        private void Settings_Click(object sender, RoutedEventArgs e)
        {
            //this.Frame.Navigate(typeof(Settings));
        }
        private void Back_Click(object sender, RoutedEventArgs e)
        {
            this.Frame.GoBack();
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
            if (!UseRenderTarget)
            {
                System.IO.MemoryStream memstream = await WinRTXamlToolkit.Composition.WriteableBitmapRenderExtensions.RenderToPngStream(element);
                Windows.Storage.StorageFile file = await Windows.Storage.ApplicationData.Current.LocalFolder.CreateFileAsync(fileName + ".png", Windows.Storage.CreationCollisionOption.ReplaceExisting);
                Windows.Storage.Streams.IRandomAccessStream stream = await file.OpenAsync(Windows.Storage.FileAccessMode.ReadWrite);
                await stream.WriteAsync(memstream.GetWindowsRuntimeBuffer());
                stream.Dispose();
            }
            else
            {
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
}