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
                items = System.Linq.Enumerable.Select(AppSettings.TR.GetDivisionTypes(), (Arr, idx) => new MyTabItem { Title = Arr, Index = idx + 1 }).ToList();
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
            ViewModel.ListSelectedItem = null;
        }
        private void OnClick(object sender, RoutedEventArgs e)
        {
            sectionListBox_DoubleTapped(sender, null);
        }
        private void OnTapped(object sender, TappedRoutedEventArgs e)
        {
            sectionListBox_DoubleTapped(sender, null);
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
    public class VisibilitySelectConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, string language)
        {
            return (bool)value ? Visibility.Visible : Visibility.Collapsed;
        }

        public object ConvertBack(object value, Type targetType, object parameter, string language)
        {
            throw new NotImplementedException();
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
}
