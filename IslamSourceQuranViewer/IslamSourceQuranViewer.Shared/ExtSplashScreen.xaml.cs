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
        internal System.Threading.SynchronizationContext ctx;

        public ExtSplashScreen()
        {
            this.InitializeComponent();
            this.BottomAppBar.Visibility = Windows.UI.Xaml.Visibility.Visible;
            this.ProgressGrid.Visibility = Windows.UI.Xaml.Visibility.Collapsed;
            Position();
        }
        public ExtSplashScreen(Windows.ApplicationModel.Activation.SplashScreen splashScreen, bool loadState)
        {
            ctx = System.Threading.SynchronizationContext.Current;
            this.InitializeComponent();
            this.BottomAppBar.Visibility = Windows.UI.Xaml.Visibility.Collapsed;
            this.ProgressGrid.Visibility = Windows.UI.Xaml.Visibility.Visible;
            // Listen for window resize events to reposition the extended splash screen image accordingly.
            // This ensures that the extended splash screen formats properly in response to window resizing.
            Window.Current.SizeChanged += new WindowSizeChangedEventHandler(SplashScreen_OnResize);
            
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
            if (Window.Current.Bounds.Width < Window.Current.Bounds.Height * .94 / 722 * 502.655)
            {
                MainGrid.ColumnDefinitions[1].Width = new GridLength(Window.Current.Bounds.Width - 2);
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
        /*async*/ void DismissedEventHandler(Windows.ApplicationModel.Activation.SplashScreen sender, object e)
        {
            dismissed = true;

            // Complete app setup operations here...
            /*await*/ DismissExtendedSplash();
        }
        void /*async System.Threading.Tasks.Task*/ DismissExtendedSplash()
        {
            ctx.Post(delegate {
                    rootFrame.Navigate(typeof(MainPage));
                    Window.Current.Content = rootFrame;
                }, null);
            //Windows Phone has problem getting TPL event sources properly initialized or a crash will occur on first task call
            //await Windows.ApplicationModel.Core.CoreApplication.MainView.CoreWindow.Dispatcher.RunAsync(Windows.UI.Core.CoreDispatcherPriority.Normal,
            //    () =>
            //    {
            //        rootFrame.Navigate(typeof(MainPage));
            //        Window.Current.Content = rootFrame;
            //    });
                // Navigate to mainpage
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
            await WindowsRTSettings.SavePathImageAsFile(50, 50, "storelogo.scale-100", MainGrid);
            await WindowsRTSettings.SavePathImageAsFile(70, 70, "storelogo.scale-140", MainGrid);
            await WindowsRTSettings.SavePathImageAsFile(90, 90, "storelogo.scale-180", MainGrid);
            await WindowsRTSettings.SavePathImageAsFile(120, 120, "logo.scale-80", MainGrid);
            await WindowsRTSettings.SavePathImageAsFile(150, 150, "logo.scale-100", MainGrid);
            await WindowsRTSettings.SavePathImageAsFile(210, 210, "logo.scale-140", MainGrid);
            await WindowsRTSettings.SavePathImageAsFile(270, 270, "logo.scale-180", MainGrid);
            await WindowsRTSettings.SavePathImageAsFile(360, 360, "logo.scale-240", MainGrid);
            await WindowsRTSettings.SavePathImageAsFile(24, 24, "smalllogo.scale-80", MainGrid);
            await WindowsRTSettings.SavePathImageAsFile(30, 30, "smalllogo.scale-100", MainGrid);
            await WindowsRTSettings.SavePathImageAsFile(42, 42, "smalllogo.scale-140", MainGrid);
            await WindowsRTSettings.SavePathImageAsFile(54, 54, "smalllogo.scale-180", MainGrid);
            await WindowsRTSettings.SavePathImageAsFile(106, 106, "smalllogo.scale-240", MainGrid);
            await WindowsRTSettings.SavePathImageAsFile(620, 300, "splashscreen.scale-100", MainGrid);
            await WindowsRTSettings.SavePathImageAsFile(868, 420, "splashscreen.scale-140", MainGrid);
            await WindowsRTSettings.SavePathImageAsFile(1116, 540, "splashscreen.scale-180", MainGrid);
            await WindowsRTSettings.SavePathImageAsFile(248, 120, "widelogo.scale-80", MainGrid);
            await WindowsRTSettings.SavePathImageAsFile(310, 150, "widelogo.scale-100", MainGrid);
            await WindowsRTSettings.SavePathImageAsFile(434, 210, "widelogo.scale-140", MainGrid);
            await WindowsRTSettings.SavePathImageAsFile(558, 270, "widelogo.scale-180", MainGrid);
            await WindowsRTSettings.SavePathImageAsFile(744, 360, "widelogo.scale-240", MainGrid);
            await WindowsRTSettings.SavePathImageAsFile(24, 24, "badgelogo.scale-100", MainGrid);
            await WindowsRTSettings.SavePathImageAsFile(34, 34, "badgelogo.scale-140", MainGrid);
            await WindowsRTSettings.SavePathImageAsFile(34, 34, "badgelogo.scale-180", MainGrid);
            await WindowsRTSettings.SavePathImageAsFile(1152, 1920, "SplashScreen.scale-240", MainGrid);
            await WindowsRTSettings.SavePathImageAsFile(360, 360, "Square150x150Logo.scale-240", MainGrid);
            await WindowsRTSettings.SavePathImageAsFile(106, 106, "Square44x44Logo.scale-240", MainGrid);
            await WindowsRTSettings.SavePathImageAsFile(170, 170, "Square71x71Logo.scale-240", MainGrid);
            await WindowsRTSettings.SavePathImageAsFile(120, 120, "StoreLogo.scale-240", MainGrid);
            await WindowsRTSettings.SavePathImageAsFile(744, 360, "Wide310x150Logo.scale-240", MainGrid);
            await WindowsRTSettings.SavePathImageAsFile(48, 48, "LockScreenLogo.scale-200", MainGrid);
            await WindowsRTSettings.SavePathImageAsFile(1240, 600, "SplashScreen.scale-200", MainGrid);
            await WindowsRTSettings.SavePathImageAsFile(300, 300, "Square150x150Logo.scale-200", MainGrid);
            await WindowsRTSettings.SavePathImageAsFile(88, 88, "Square44x44Logo.scale-200", MainGrid);
            await WindowsRTSettings.SavePathImageAsFile(24, 24, "Square44x44Logo.targetsize-24_altform-unplated", MainGrid);
            await WindowsRTSettings.SavePathImageAsFile(50, 50, "StoreLogo", MainGrid);
            await WindowsRTSettings.SavePathImageAsFile(620, 300, "Wide310x150Logo.scale-200", MainGrid);
            await WindowsRTSettings.SavePathImageAsFile(846, 468, "appstorepromotional-846x468", MainGrid);
            await WindowsRTSettings.SavePathImageAsFile(558, 756, "appstorepromotional-558x756", MainGrid);
            await WindowsRTSettings.SavePathImageAsFile(414, 468, "appstorepromotional-414x468", MainGrid);
            await WindowsRTSettings.SavePathImageAsFile(414, 180, "appstorepromotional-414x180", MainGrid);
            await WindowsRTSettings.SavePathImageAsFile(558, 558, "appstorepromotional-558x558", MainGrid);
            await WindowsRTSettings.SavePathImageAsFile(2400, 1200, "appstorepromotional-2400x1200", MainGrid);
            await WindowsRTSettings.SavePathImageAsFile(300, 300, "appstorephonetitleicon-300x300", MainGrid);
            await WindowsRTSettings.SavePathImageAsFile(1000, 800, "appstorephonepromotional-1000x800", MainGrid);
            await WindowsRTSettings.SavePathImageAsFile(358, 358, "appstorephonepromotional-358x358", MainGrid);
            await WindowsRTSettings.SavePathImageAsFile(358, 173, "appstorephonepromotional-358x173", MainGrid);
            GC.Collect();
        }
        private void Settings_Click(object sender, RoutedEventArgs e)
        {
            //this.Frame.Navigate(typeof(Settings));
        }
        private void Back_Click(object sender, RoutedEventArgs e)
        {
            this.Frame.GoBack();
        }
    }
}