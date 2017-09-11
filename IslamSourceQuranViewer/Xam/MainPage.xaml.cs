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
using System.Globalization;
using System.Collections.ObjectModel;

namespace IslamSourceQuranViewer.Xam
{
    public class ItemsView : Grid
    {
        protected StackLayout PagingStackLayout;
        protected ScrollView ScrollView;
        protected readonly ICommand SelectedCommand;
        protected readonly StackLayout ItemsStackLayout;
        protected Boolean Wait = false;

        public class TapCommand : ICommand
        {
            public event EventHandler CanExecuteChanged;

            public bool CanExecute(object parameter)
            {
                return true;
            }

            public void Execute(object parameter)
            {
                dynamic par = parameter;
                par.ItemsView.SelectedItem = par.Item;
            }
        }

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
                Padding = new Thickness(2, 0, 2, 0),
                Spacing = 1,
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
                new Thickness(0, 0, 0, 5),
                new Thickness(0, 0, 0, 5),
                new Thickness(0, 0, 0, 5)),
                Opacity = 0.5
            };
            Children.Add(PagingStackLayout);

            SelectedCommand = new TapCommand();

            /*var leftArrow = new Image()
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
            Children.Add(rightArrow);*/

            PropertyChanged += (sender, e) =>
            {
                if (e.PropertyName == "Orientation")
                {
                    ItemsStackLayout.Orientation = ScrollView.Orientation == ScrollOrientation.Horizontal ? StackOrientation.Horizontal : StackOrientation.Vertical;
                }

            };
        }

#if !__ANDROID__
        public ItemsView(ListViewCachingStrategy cachingStrategy) : this()
        {
            if ((Device.OS == TargetPlatform.Android) || (Device.OS == TargetPlatform.iOS))
            {
                this.CachingStrategy = cachingStrategy;
            }
        }
        internal ListViewCachingStrategy CachingStrategy { get; private set; }
#endif
        public int ItemsCount
        {
            get { return this.ItemsStackLayout.Children.Count; }
        }

        public View ActualElement
        {
            get
            {
                return ItemsStackLayout.Children[ActualElementIndex];
            }
        }
        public int ActualElementIndex { get; set; }

        public event EventHandler<SelectedItemChangedEventArgs> ItemSelected;
        public event EventHandler<ItemTappedEventArgs> ItemTapped;

        public static readonly BindableProperty HasUnevenColumnsProperty = BindableProperty.Create("HasUnevenColumns", typeof(bool), typeof(ListView), (bool)false, BindingMode.OneWay, null, null, null, null, null);
        public static readonly BindableProperty ColumnWidthProperty = BindableProperty.Create("ColumnWidth", typeof(int), typeof(ListView), (int)(-1), BindingMode.OneWay, null, null, null, null, null);
        public static readonly BindableProperty SeparatorColorProperty = BindableProperty.Create("SeparatorColor", typeof(Color), typeof(ListView), Color.Default, BindingMode.OneWay, null, null, null, null, null);
        public static readonly BindableProperty SeparatorVisibilityProperty = BindableProperty.Create("SeparatorVisibility", typeof(Xamarin.Forms.SeparatorVisibility), typeof(ListView), Xamarin.Forms.SeparatorVisibility.Default, BindingMode.OneWay, null, null, null, null, null);

        public static readonly BindableProperty ItemsSourceProperty =
            BindableProperty.Create("ItemsSource", typeof(IEnumerable), typeof(ItemsView), null, BindingMode.OneWay, null, new BindableProperty.BindingPropertyChangedDelegate(ItemsView.OnItemsSourceChanged), null, null, null);

        public IEnumerable ItemsSource
        {
            get
            {
                return (IEnumerable)base.GetValue(ItemsView.ItemsSourceProperty);
            }
            set
            {
                base.SetValue(ItemsView.ItemsSourceProperty, value);
            }
        }

        public static readonly BindableProperty SelectedItemProperty =
            BindableProperty.Create("SelectedItem", typeof(object), typeof(ItemsView), null, BindingMode.OneWayToSource, null, new BindableProperty.BindingPropertyChangedDelegate(ItemsView.OnSelectedItemChanged), null, null, null);

        public object SelectedItem
        {
            get
            {
                return base.GetValue(SelectedItemProperty);
            }
            set
            {
                base.SetValue(SelectedItemProperty, value);
            }
        }

        public static readonly BindableProperty ItemTemplateProperty =
            BindableProperty.Create("ItemTemplate", typeof(DataTemplate), typeof(ItemsView), null, BindingMode.OneWay, new BindableProperty.ValidateValueDelegate(ItemsView.ValidateItemTemplate), null, null, null, null);

        public DataTemplate ItemTemplate
        {
            get
            {
                return (DataTemplate)base.GetValue(ItemsView.ItemTemplateProperty);
            }
            set
            {
                base.SetValue(ItemsView.ItemTemplateProperty, value);
            }
        }

        private static bool ValidateItemTemplate(BindableObject b, object v)
        {
            ItemsView view = b as ItemsView;
#if !__ANDROID__
            if ((view != null) && (view.CachingStrategy == ListViewCachingStrategy.RetainElement))
            {
                return !(view.ItemTemplate is DataTemplateSelector);
            }
#endif
            return true;
        }

        private static void OnItemsSourceChanged(BindableObject bindable, object oldValue, object newValue)
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

            SelectedItem = ItemsSource.OfType<object>().FirstOrDefault(x => SelectedItem == x);
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
                CommandParameter = new { ItemsView = this, Item = item }
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
            ItemsView view = (ItemsView)bindable;
            for (int count = 0; count < view.ItemsStackLayout.Children.Count - 1; count++)
            {
                if (view.ItemsStackLayout.Children[count].BindingContext == newValue)
                {
                    view.ItemsStackLayout.Children[count].BackgroundColor = Xamarin.Forms.Color.Blue;
                } else
                {
                    view.ItemsStackLayout.Children[count].BackgroundColor = Xamarin.Forms.Color.White;
                }
            }
            if (view.ItemSelected != null)
            {
                view.ItemSelected(view, new SelectedItemChangedEventArgs(newValue));
            }
        }

        //protected override SizeRequest OnSizeRequest(double widthConstraint, double heightConstraint)
        //{
        //    Size size2;
        //    Size minimum = new Size(40.0, 40.0);
            
        //    double width = Math.Min(Device.Info.ScaledScreenSize.Width, Device.Info.ScaledScreenSize.Height);
        //    IList itemsSource = ItemsSource as IList;
        //    if (((itemsSource != null) && !this.HasUnevenColumns) && ((this.ColumnWidth > 0) && !this.IsGroupingEnabled))
        //    {
        //        size2 = new Size(width, (double)(itemsSource.Count * this.ColumnWidth));
        //    }
        //    else
        //    {
        //        size2 = new Size(width, Math.Max(Device.Info.ScaledScreenSize.Width, Device.Info.ScaledScreenSize.Height));
        //    }
        //    return new SizeRequest(size2, minimum);
        //}

        public bool HasUnevenColumns
        {
            get
            {
                return (bool)((bool)base.GetValue(HasUnevenColumnsProperty));
            }
            set
            {
                base.SetValue(HasUnevenColumnsProperty, (bool)value);
            }
        }
        public int ColumnWidth
        {
            get
            {
                return (int)((int)base.GetValue(ColumnWidthProperty));
            }
            set
            {
                base.SetValue(ColumnWidthProperty, (int)value);
            }
        }
        public Color SeparatorColor
        {
            get
            {
                return (Color)base.GetValue(SeparatorColorProperty);
            }
            set
            {
                base.SetValue(SeparatorColorProperty, value);
            }
        }

        public Xamarin.Forms.SeparatorVisibility SeparatorVisibility
        {
            get
            {
                return (Xamarin.Forms.SeparatorVisibility)base.GetValue(SeparatorVisibilityProperty);
            }
            set
            {
                base.SetValue(SeparatorVisibilityProperty, value);
            }
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

    public partial class MainPage : ContentPage
	{
		public MainPage ()
		{
            this.BindingContext = this;
            this.ViewModel = new MyTabViewModel();
            UIChanger = new MyUIChanger();
            List<MyTabItem> items = null;
            items = System.Linq.Enumerable.Select(AppSettings.TR.GetDivisionTypes(), (Arr, idx) => new MyTabItem { Title = Arr, Index = idx + 1 }).ToList();
            items.Insert(0, new MyTabItem { IsBookmarks = true, Title = ISQV.Xam.UWP.Resources.AppResources.Bookmarks_Text, Index = 0 });
            ViewModel.Items = items;
            InitializeComponent();
        }
        public MyUIChanger UIChanger { get; set; }
        public MyTabViewModel ViewModel { get; set; }
        private void sectionListBox_DoubleTapped(object sender, ItemTappedEventArgs e)
        {
            if (ViewModel.ListSelectedItem == null) return;
            this.Navigation.PushAsync(new WordForWordUC(new { Division = ViewModel.SelectedItem.Index, Selection = ViewModel.ListSelectedItem.Index }));
        }
        private void Settings_Click(object sender, EventArgs e)
        {
            Navigation.PushAsync(new Settings());
        }
        private void About_Click(object sender, EventArgs e)
        {
            Navigation.PushAsync(new About());
        }
        private void Tab_ItemSelected(object sender, SelectedItemChangedEventArgs e)
        {
            ViewModel.SelectedItem = e.SelectedItem as MyTabItem;
        }
    }

    public class ListFormattedText : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
        {
            List<string> val = value as List<string>;
            FormattedString fs = new FormattedString();
            fs.Spans.Add(new Span { Text = val[0] });//, FontFamily = new FontFamily(AppSettings.strOtherSelectedFont), FontSize = AppSettings.dOtherFontSize });
            if (val.Count > 1)
            {
                fs.Spans.Add(new Span { Text = "(" + val[1] + ")", FontAttributes = FontAttributes.Bold });//, FlowDirection = FlowDirection.RightToLeft, FontFamily = new FontFamily(AppSettings.strSelectedFont), FontSize = AppSettings.dFontSize });
                fs.Spans.Add(new Span { Text = val[2] });//, FontFamily = new FontFamily(AppSettings.strOtherSelectedFont), FontSize = AppSettings.dOtherFontSize });
            }
            return fs;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}
