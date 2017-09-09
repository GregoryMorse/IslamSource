using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using Xamarin.Forms;

#if WINDOWS_UWP
namespace IslamSourceQuranViewer.Xam
{
	public partial class App : Application
	{
		public App ()
		{
			InitializeComponent();

            //AppSettings.InitDefaultSettings().Wait();
			MainPage = new NavigationPage(new ISQV.Xam.ExtSplashScreen());
			// The root page of your application
			//MainPage = new ContentPage {
			//	Content = new StackLayout {
			//		VerticalOptions = LayoutOptions.Center,
			//		Children = {
			//			new Label {
			//				XAlign = TextAlignment.Center,
			//				Text = "Welcome to Xamarin Forms!"
			//			}
			//		}
			//	}
			//};
		}

		protected override void OnStart ()
		{
			// Handle when your app starts
		}

		protected override void OnSleep ()
		{
			// Handle when your app sleeps
		}

		protected override void OnResume ()
		{
			// Handle when your app resumes
		}
	}
}
#endif