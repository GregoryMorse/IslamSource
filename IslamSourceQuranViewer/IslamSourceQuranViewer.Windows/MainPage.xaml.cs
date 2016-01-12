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
            this.InitializeComponent();
        }
        public MyTabViewModel ViewModel { get; set; }
    }
    public class MyTabViewModel : INotifyPropertyChanged
    {
        public MyTabViewModel()
        {
            Items =
                new List<MyTabItem>
                {
                    new MyTabItem
                        {
                            Title = "Overview",
                            Content = null
                        },
                    new MyTabItem
                        {
                            Title = "Detail",
                            Content = null
                        },
                    new MyTabItem
                        {
                            Title = "Reviews",
                            Content = null
                        },
                };
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
