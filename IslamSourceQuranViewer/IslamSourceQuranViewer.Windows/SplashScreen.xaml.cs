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
    public sealed partial class SplashScreen : Page
    {
        async System.Threading.Tasks.Task SavePathImageAsFile(int Width, int Height, string fileName)
        {
            float dpi = Windows.Graphics.Display.DisplayInformation.GetForCurrentView().LogicalDpi;
            Windows.UI.Xaml.Media.Imaging.RenderTargetBitmap wb = new Windows.UI.Xaml.Media.Imaging.RenderTargetBitmap();
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
            Windows.Storage.StorageFile file = await Windows.Storage.ApplicationData.Current.LocalFolder.CreateFileAsync(fileName + ".png", Windows.Storage.CreationCollisionOption.ReplaceExisting);
            Windows.Storage.Streams.IRandomAccessStream stream = await file.OpenAsync(Windows.Storage.FileAccessMode.ReadWrite);
            await wb.RenderAsync(MainGrid, (int)((float)Width * 96 / dpi), (int)((float)Height * 96 / dpi));
            //Windows.Graphics.Imaging.BitmapPropertySet propertySet = new Windows.Graphics.Imaging.BitmapPropertySet();
            //propertySet.Add("ImageQuality", new Windows.Graphics.Imaging.BitmapTypedValue(1.0, Windows.Foundation.PropertyType.Single)); // Maximum quality
            Windows.Graphics.Imaging.BitmapEncoder be = await Windows.Graphics.Imaging.BitmapEncoder.CreateAsync(Windows.Graphics.Imaging.BitmapEncoder.PngEncoderId, stream);//, propertySet);
            Windows.Storage.Streams.IBuffer buf = await wb.GetPixelsAsync();
            be.SetPixelData(Windows.Graphics.Imaging.BitmapPixelFormat.Bgra8, Windows.Graphics.Imaging.BitmapAlphaMode.Premultiplied, (uint)wb.PixelWidth, (uint)wb.PixelHeight, dpi, dpi, buf.ToArray());
            await be.FlushAsync();
            await stream.GetOutputStreamAt(0).FlushAsync();
            stream.Dispose();
        }
        public SplashScreen()
        {
            this.InitializeComponent();
        }

        private async void Grid_DoubleTapped(object sender, DoubleTappedRoutedEventArgs e)
        {
            await SavePathImageAsFile(50, 50, "storelogo.scale-100");
            await SavePathImageAsFile(70, 70, "storelogo.scale-140");
            await SavePathImageAsFile(90, 90, "storelogo.scale-180");
            await SavePathImageAsFile(120, 120, "logo.scale-80");
            await SavePathImageAsFile(150, 150, "logo.scale-100");
            await SavePathImageAsFile(210, 210, "logo.scale-140");
            await SavePathImageAsFile(270, 270, "logo.scale-180");
            await SavePathImageAsFile(24, 24, "smalllogo.scale-80");
            await SavePathImageAsFile(30, 30, "smalllogo.scale-100");
            await SavePathImageAsFile(42, 42, "smalllogo.scale-140");
            await SavePathImageAsFile(54, 54, "smalllogo.scale-180");
            await SavePathImageAsFile(620, 300, "splashscreen.scale-100");
            await SavePathImageAsFile(868, 420, "splashscreen.scale-140");
            await SavePathImageAsFile(1116, 540, "splashscreen.scale-180");
            await SavePathImageAsFile(248, 120, "widelogo.scale-80");
            await SavePathImageAsFile(310, 150, "widelogo.scale-100");
            await SavePathImageAsFile(434, 210, "widelogo.scale-140");
            await SavePathImageAsFile(558, 270, "widelogo.scale-180");
            await SavePathImageAsFile(24, 24, "badgelogo.scale-100");
            await SavePathImageAsFile(34, 34, "badgelogo.scale-140");
            await SavePathImageAsFile(34, 34, "badgelogo.scale-180");
            await SavePathImageAsFile(1366, 768, "appstorescreenshot-wide");
            await SavePathImageAsFile(768, 1366, "appstorescreenshot-tall");
            await SavePathImageAsFile(846, 468, "appstorepromotional-846x468");
            await SavePathImageAsFile(558, 756, "appstorepromotional-558x756");
            await SavePathImageAsFile(414, 468, "appstorepromotional-414x468");
            await SavePathImageAsFile(414, 180, "appstorepromotional-414x180");
            await SavePathImageAsFile(558, 558, "appstorepromotional-558x558");
            await SavePathImageAsFile(2400, 1200, "appstorepromotional-2400x1200");
            await SavePathImageAsFile(1152, 1920, "SplashScreen.scale-240");
            await SavePathImageAsFile(360, 360, "Square150x150Logo.scale-240");
            await SavePathImageAsFile(106, 106, "Square44x44Logo.scale-240");
            await SavePathImageAsFile(170, 170, "Square71x71Logo.scale-240");
            await SavePathImageAsFile(120, 120, "StoreLogo.scale-240");
            await SavePathImageAsFile(744, 360, "Wide310x150Logo.scale-240");
            await SavePathImageAsFile(300, 300, "appstorephonetitleicon-300x300");
            await SavePathImageAsFile(1000, 800, "appstorephonepromotional-1000x800");
            await SavePathImageAsFile(358, 358, "appstorephonepromotional-358x358");
            await SavePathImageAsFile(358, 173, "appstorephonepromotional-358x173");
            await SavePathImageAsFile(1280, 768, "appstorephonescreenshot-wide");
            await SavePathImageAsFile(768, 1280, "appstorephonescreenshot-tall");
            await SavePathImageAsFile(1280, 720, "appstorephonescreenshot1280x720-wide");
            await SavePathImageAsFile(720, 1280, "appstorephonescreenshot720x1280-tall");
            await SavePathImageAsFile(800, 480, "appstorephonescreenshot800x480-wide");
            await SavePathImageAsFile(480, 800, "appstorephonescreenshot480x800-tall");
            await SavePathImageAsFile(48, 48, "LockScreenLogo.scale-200");
            await SavePathImageAsFile(1240, 600, "SplashScreen.scale-200");
            await SavePathImageAsFile(300, 300, "Square150x150Logo.scale-200");
            await SavePathImageAsFile(88, 88, "Square44x44Logo.scale-200");
            await SavePathImageAsFile(24, 24, "Square44x44Logo.targetsize-24_altform-unplated");
            await SavePathImageAsFile(50, 50, "StoreLogo");
            await SavePathImageAsFile(620, 300, "Wide310x150Logo.scale-200");
            GC.Collect();
        }
    }
}