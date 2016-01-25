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
    public sealed partial class WordForWordUC : Page
    {
        public WordForWordUC()
        {
            this.InitializeComponent();
        }

        protected override void OnNavigatedTo(NavigationEventArgs e)
        {
            dynamic t = e.Parameter;
            Division = t.Division;
            Selection = t.Selection;
            this.DataContext = this;
            this.ViewModel = new MyRenderModel(IslamMetadata.TanzilReader.GetRenderedQuranText(XMLRender.ArabicData.TranslitScheme.RuleBased, "PlainRoman", String.Empty, Division.ToString(), Selection.ToString(), String.Empty, "0", "1").Items);
            MainControl.ItemsSource = this.ViewModel.RenderItems;
        }
        public MyRenderModel ViewModel { get; set; }
        private async void RenderPngs_Click(object sender, RoutedEventArgs e)
        {
            await WindowsRTSettings.SavePathImageAsFile(1366, 768, "appstorescreenshot-wide", this);
            await WindowsRTSettings.SavePathImageAsFile(768, 1366, "appstorescreenshot-tall", this);
            await WindowsRTSettings.SavePathImageAsFile(1280, 768, "appstorephonescreenshot-wide", this);
            await WindowsRTSettings.SavePathImageAsFile(768, 1280, "appstorephonescreenshot-tall", this);
            await WindowsRTSettings.SavePathImageAsFile(1280, 720, "appstorephonescreenshot1280x720-wide", this);
            await WindowsRTSettings.SavePathImageAsFile(720, 1280, "appstorephonescreenshot720x1280-tall", this);
            await WindowsRTSettings.SavePathImageAsFile(800, 480, "appstorephonescreenshot800x480-wide", this);
            await WindowsRTSettings.SavePathImageAsFile(480, 800, "appstorephonescreenshot480x800-tall", this);
            GC.Collect();
        }
        private void Settings_Click(object sender, RoutedEventArgs e)
        {
            //this.Frame.Navigate(typeof(Settings));
        }

        private int Division;
        private int Selection;
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
            Items = System.Linq.Enumerable.Select(RendItem.TextItems, (Arr) => (Arr.Text.GetType() == typeof(List<XMLRender.RenderArray.RenderItem>)) ? (object)new MyRenderModel((List<XMLRender.RenderArray.RenderItem>)Arr.Text) : (Arr.Text.GetType() == typeof(string) ? (object)new MyChildRenderItem { ItemText = (string)Arr.Text } : null)).Where(Arr => Arr != null);
        }
        public IEnumerable<object> Items { get; set; }
    }
    public class MyChildRenderItem
    {
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
