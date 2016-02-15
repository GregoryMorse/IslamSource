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
using WinRTXamlToolkit.Controls;

namespace IslamSourceQuranViewer
{
    public class MyWrapPanel : WrapPanel
    {
        protected override Size MeasureOverride(Size constraint)
        {
            Size size = base.MeasureOverride(constraint);
            size.Width += 1;
            return size;
        }
    }
    public sealed partial class WordForWordUC : Page
    {
        public WordForWordUC()
        {
            UIChanger = new MyUIChanger();
            this.InitializeComponent();
        }

        protected override void OnNavigatedTo(NavigationEventArgs e)
        {
            dynamic t = e.Parameter;
            Division = t.Division;
            Selection = t.Selection;
            this.DataContext = this;
            this.ViewModel = new MyRenderModel(IslamMetadata.TanzilReader.GetRenderedQuranText((bool)Windows.Storage.ApplicationData.Current.LocalSettings.Values["ShowTransliteration"] ? XMLRender.ArabicData.TranslitScheme.RuleBased : XMLRender.ArabicData.TranslitScheme.None, String.Empty, String.Empty, Division.ToString(), Selection.ToString(), String.Empty, (bool)Windows.Storage.ApplicationData.Current.LocalSettings.Values["ShowW4W"] ? "0" : "4", (bool)Windows.Storage.ApplicationData.Current.LocalSettings.Values["UseColoring"] ? "0" : "1").Items);
            MainControl.ItemsSource = this.ViewModel.RenderItems;
        }
        public MyUIChanger UIChanger { get; set; }
        public MyRenderModel ViewModel { get; set; }
        private async void RenderPngs_Click(object sender, RoutedEventArgs e)
        {
            await WindowsRTSettings.SavePathImageAsFile(1366, 768, "appstorescreenshot-wide", this, false);
            await WindowsRTSettings.SavePathImageAsFile(768, 1366, "appstorescreenshot-tall", this, false);
            await WindowsRTSettings.SavePathImageAsFile(1280, 768, "appstorephonescreenshot-wide", this, false);
            await WindowsRTSettings.SavePathImageAsFile(768, 1280, "appstorephonescreenshot-tall", this, false);
            await WindowsRTSettings.SavePathImageAsFile(1280, 720, "appstorephonescreenshot1280x720-wide", this, false);
            await WindowsRTSettings.SavePathImageAsFile(720, 1280, "appstorephonescreenshot720x1280-tall", this, false);
            await WindowsRTSettings.SavePathImageAsFile(800, 480, "appstorephonescreenshot800x480-wide", this, false);
            await WindowsRTSettings.SavePathImageAsFile(480, 800, "appstorephonescreenshot480x800-tall", this, false);
            GC.Collect();
        }
        private void Settings_Click(object sender, RoutedEventArgs e)
        {
            this.Frame.Navigate(typeof(Settings));
        }
        private void Back_Click(object sender, RoutedEventArgs e)
        {
            this.Frame.GoBack();
        }

        private int Division;
        private int Selection;

        private void ZoomIn_Click(object sender, RoutedEventArgs e)
        {
            UIChanger.AllFontSize += 1.0;
        }

        private void ZoomOut_Click(object sender, RoutedEventArgs e)
        {
            UIChanger.AllFontSize -= 1.0;
        }

        private void Default_Click(object sender, RoutedEventArgs e)
        {
            UIChanger.AllFontSize = 20.0;
        }
    }
    public class MyUIChanger : INotifyPropertyChanged
    {
        public MyUIChanger()
        {
            _AllFontSize = 20.0;
        }
        private double _AllFontSize;
        public double AllFontSize
        {
            get
            {
                return _AllFontSize;
            }
            set
            {
                _AllFontSize = value;
                PropertyChanged(this, new PropertyChangedEventArgs("AllFontSize"));
            }
        }
        public string FontFamily
        {
            get
            {
                return (string)Windows.Storage.ApplicationData.Current.LocalSettings.Values["CurrentFont"];
            }
        }
        
        #region Implementation of INotifyPropertyChanged

        public event PropertyChangedEventHandler PropertyChanged;

        #endregion
    }
    public class MyRenderModel
    {
        public MyRenderModel(List<XMLRender.RenderArray.RenderItem> RendItems)
        {
            RenderItems = System.Linq.Enumerable.Select(RendItems, (Arr) => new MyRenderItem((XMLRender.RenderArray.RenderItem)Arr));
        }

        public IEnumerable<MyRenderItem> RenderItems { get; set; }
    }
    public class MyRenderItem
    {
        public MyRenderItem(XMLRender.RenderArray.RenderItem RendItem)
        {
            Items = System.Linq.Enumerable.Select(RendItem.TextItems.GroupBy((MainItems) => (MainItems.DisplayClass == XMLRender.RenderArray.RenderDisplayClass.eArabic || MainItems.DisplayClass == XMLRender.RenderArray.RenderDisplayClass.eLTR || MainItems.DisplayClass == XMLRender.RenderArray.RenderDisplayClass.eRTL || MainItems.DisplayClass == XMLRender.RenderArray.RenderDisplayClass.eTransliteration) ? (object)MainItems.DisplayClass : (object)MainItems), (Arr) => (Arr.First().Text.GetType() == typeof(List<XMLRender.RenderArray.RenderItem>)) ? (object)new MyRenderModel((List<XMLRender.RenderArray.RenderItem>)Arr.First().Text) : (Arr.First().Text.GetType() == typeof(string) ? (object)new MyChildRenderItem { ItemRuns = System.Linq.Enumerable.Select(Arr, (ArrItem) => new MyChildRenderBlockItem() { ItemText = (string)ArrItem.Text, Color = new SolidColorBrush(Windows.UI.Color.FromArgb(0xFF, XMLRender.Utility.ColorR(ArrItem.Clr), XMLRender.Utility.ColorG(ArrItem.Clr), XMLRender.Utility.ColorB(ArrItem.Clr))) }).ToList(), Direction = Arr.First().DisplayClass == XMLRender.RenderArray.RenderDisplayClass.eArabic || Arr.First().DisplayClass == XMLRender.RenderArray.RenderDisplayClass.eRTL ? FlowDirection.RightToLeft : FlowDirection.LeftToRight } : null)).Where(Arr => Arr != null);
        }
        public IEnumerable<object> Items { get; set; }
    }
    public static class FormattedTextBehavior
    {
        #region FormattedText Attached dependency property

        public static List<MyChildRenderBlockItem> GetFormattedText(DependencyObject obj)
        {
            return (List<MyChildRenderBlockItem>)obj.GetValue(FormattedTextProperty);
        }

        public static void SetFormattedText(DependencyObject obj, List<MyChildRenderBlockItem> value)
        {
            obj.SetValue(FormattedTextProperty, value);
        }

        public static readonly DependencyProperty FormattedTextProperty =
            DependencyProperty.RegisterAttached("FormattedText",
            typeof(List<MyChildRenderBlockItem>),
            typeof(FormattedTextBehavior),
            new PropertyMetadata(null, FormattedTextChanged));

        private static void FormattedTextChanged(DependencyObject sender, DependencyPropertyChangedEventArgs e)
        {
            List<MyChildRenderBlockItem> value = e.NewValue as List<MyChildRenderBlockItem>;

            TextBlock textBlock = sender as TextBlock;

            if (textBlock != null)
            {
                textBlock.Inlines.Clear();
                value.FirstOrDefault((it) => { textBlock.Inlines.Add(new Windows.UI.Xaml.Documents.Run() { Text = it.ItemText, Foreground = it.Color }); return false; });
            }
        }

        #endregion
    }
    public class MyChildRenderItem
    {
        public List<MyChildRenderBlockItem> ItemRuns { get; set; }
        public FlowDirection Direction { get; set; }
    }
    public class MyChildRenderBlockItem
    {
        public SolidColorBrush Color { get; set; }
        public string ItemText { get; set; }
    }
    public class MyDataTemplateSelector : DataTemplateSelector
    {
        public DataTemplate MyTemplate { get; set; }

        protected override DataTemplate SelectTemplateCore(object item, DependencyObject container)
        {
            return MyTemplate;
        }
    }
    public class NormalWordTemplateSelector : DataTemplateSelector
    {
        public DataTemplate WordTemplate { get; set; }

        public DataTemplate NormalTemplate { get; set; }

        protected override DataTemplate SelectTemplateCore(object item, DependencyObject container)
        {
            if (item.GetType() == typeof(MyRenderModel)) { return WordTemplate; }
            return NormalTemplate;
        }

    }
}
