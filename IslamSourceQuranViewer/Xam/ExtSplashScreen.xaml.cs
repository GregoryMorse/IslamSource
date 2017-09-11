using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices.WindowsRuntime;
using Windows.Foundation;
using Windows.Foundation.Collections;
using Xamarin.Forms;

// The Blank Page item template is documented at http://go.microsoft.com/fwlink/?LinkId=234238

namespace ISQV.Xam
{
    /// <summary>
    /// An empty page that can be used on its own or navigated to within a Frame.
    /// </summary>
    public sealed partial class ExtSplashScreen : ContentPage
    {
        internal Rect splashImageRect; // Rect to store splash screen image coordinates.
        private Windows.ApplicationModel.Activation.SplashScreen splash; // Variable to hold the splash screen object.
        internal bool dismissed = false; // Variable to track splash screen dismissal status.
        internal Frame rootFrame;

        public ExtSplashScreen()
        {
            this.InitializeComponent();
            //this.BottomAppBar.Visibility = Windows.UI.Xaml.Visibility.Visible;
            //this.ProgressGrid.Visibility = Windows.UI.Xaml.Visibility.Collapsed;
            Position();
            //LayoutUpdated += SplashScreen_LayoutUpdated;
#if WINDOWS_APP
            AppBarButton BackButton = new AppBarButton() { Icon = new SymbolIcon(Symbol.Back), Label = new Windows.ApplicationModel.Resources.ResourceLoader().GetString("Back/Label") };
            BackButton.Click += Back_Click;
            (this.BottomAppBar as CommandBar).PrimaryCommands.Add(BackButton);
#endif
            System.Threading.Tasks.Task t = DismissExtendedSplash();
        }
        void OnCanvasViewPaintSurface(object sender, SkiaSharp.Views.Forms.SKPaintSurfaceEventArgs args)
        {
            SkiaSharp.SKImageInfo info = args.Info;
            SkiaSharp.SKSurface surface = args.Surface;
            SkiaSharp.SKCanvas canvas = surface.Canvas;
            canvas.Clear();
            SkiaSharp.SKRect bounds;
            SkiaSharp.SKPath path = SkiaSharp.SKPath.ParseSvgPathData((string)this.Resources["PathString"]);
            path.GetTightBounds(out bounds);
            SkiaSharp.SKPaint paint = new SkiaSharp.SKPaint
            {
                Style = SkiaSharp.SKPaintStyle.Stroke,
                Color = SkiaSharp.SKColors.Black,
                StrokeWidth = 10,
                StrokeCap = SkiaSharp.SKStrokeCap.Round,
                StrokeJoin = SkiaSharp.SKStrokeJoin.Round
            };
            canvas.Translate(info.Width / 2, info.Height / 2);

            canvas.Scale(info.Width / (bounds.Width + paint.StrokeWidth),
                         info.Height / (bounds.Height + paint.StrokeWidth));

            canvas.Translate(-bounds.MidX, -bounds.MidY);

            canvas.DrawPath(path, paint);
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
            /*if (Window.Current.Bounds.Width < Window.Current.Bounds.Height * .94 / 746 * 502.655)
            {
                //MainGrid.ColumnDefinitions[1].Width = new GridLength(Window.Current.Bounds.Width - 2);
                SplashImg.Width = Window.Current.Bounds.Width - 2;
                ProgressGrid.Height = SplashImg.Height = SplashImg.Width / 502.655 * 746;
            }
            else
            {
                SplashImg.Width = Window.Current.Bounds.Height * .94 / 746 * 502.655;
                ProgressGrid.Height = SplashImg.Height = Window.Current.Bounds.Height * .94;
            }*/
        }
        void SplashScreen_LayoutUpdated(object sender, object e)
        {
            /*if (splashProgressRing.Height != ProgressGrid.RowDefinitions[1].ActualHeight)
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
            }*/
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
            await IslamSourceQuranViewer.AppSettings.InitDefaultSettings();
            Device.BeginInvokeOnMainThread(async () =>
            {
                await Navigation.PushAsync(new IslamSourceQuranViewer.Xam.MainPage());
                Navigation.RemovePage(this);
            });
            /*await Windows.ApplicationModel.Core.CoreApplication.MainView.CoreWindow.Dispatcher.RunAsync(Windows.UI.Core.CoreDispatcherPriority.Normal,
                () =>
                {
                    //rootFrame.Navigate(typeof(MainPage));
                    //Window.Current.Content = rootFrame;
                });*/
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

        private void Settings_Click(object sender, EventArgs e)
        {
            //this.Frame.Navigate(typeof(Settings));
        }
        private void Back_Click(object sender, EventArgs e)
        {
            //this.Frame.GoBack();
        }
 
    }
}