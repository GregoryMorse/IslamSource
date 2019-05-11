using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Xamarin.Forms;

using System.Collections;
using System.Collections.Specialized;
using System.ComponentModel;

namespace IslamSourceQuranViewer.Xam
{
	public partial class WordForWordUC : ContentPage
	{
		public WordForWordUC(object Parameter)
		{
            dynamic c = Parameter;
            Division = c.Division;
            Selection = c.Selection;
            this.BindingContext = this;
            this.ViewModel = new VirtualizingWrapPanelAdapter();
            UIChanger = new MyUIChanger();
            InitializeComponent();
            WidthLimiterConverter.BindingContext = this;
            MainGrid.SizeChanged += OnSizeChanged;

            WidthLimiterConverter.TopLevelMaxValue = UIChanger.MaxWidth;
            double curWidth = UIChanger.MaxWidth;
            AppSettings.TR.GetRenderedQuranText(AppSettings.bShowTransliteration ? XMLRender.ArabicData.TranslitScheme.RuleBased : XMLRender.ArabicData.TranslitScheme.None, String.Empty, String.Empty, Division.ToString(), Selection.ToString(), !AppSettings.bShowTranslation ? "None" : AppSettings.ChData.IslamData.Translations.TranslationList[AppSettings.iSelectedTranslation].FileName, AppSettings.bShowW4W ? "0" : "4", AppSettings.bUseColoring ? "0" : "1").ContinueWith((x) => {
                this.ViewModel.RenderModels = VirtualizingWrapPanelAdapter.GroupRenderModels(System.Linq.Enumerable.Select(x.Result.Items, (Arr) => new MyRenderItem((XMLRender.RenderArray.RenderItem)Arr)).ToList(), curWidth);
            });
        }
        public MyUIChanger UIChanger { get; set; }
        public VirtualizingWrapPanelAdapter ViewModel { get; set; }

        private int Division;
        private int Selection;
        private void OnSizeChanged(object sender, EventArgs e)
        {
            UIChanger.MaxWidth = MainGrid.Width;
            WidthLimiterConverter.TopLevelMaxValue = UIChanger.MaxWidth;
            UIChanger.MaxHeight = MainGrid.Height;
            this.ViewModel.RegroupRenderModels(UIChanger.MaxWidth);
        }
        private void Settings_Click(object sender, EventArgs e)
        {
            Navigation.PushAsync(new Settings());
        }
        private void About_Click(object sender, EventArgs e)
        {
            Navigation.PushAsync(new About());
        }
    }
    public class WidthLimiterConverter : BindableObject, IValueConverter
    {
        public double TopLevelMaxValue {
            get { return (double)GetValue(TopLevelMaxValueProperty); }
            set { SetValue(TopLevelMaxValueProperty, value); OnPropertyChanged(); }
        }

        public static readonly BindableProperty TopLevelMaxValueProperty =
            BindableProperty.Create("TopLevelMaxValue",
                                        typeof(double),
                                        typeof(WidthLimiterConverter),
                                        0.0);

        public object Convert(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            return Math.Min((double)value, TopLevelMaxValue);
        }

        public object ConvertBack(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
    public class RunFormattedText : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            List<MyChildRenderBlockItem> val = value as List<MyChildRenderBlockItem>;
            FormattedString fs = new FormattedString();
            val.FirstOrDefault((it) => { fs.Spans.Add(new Span() { Text = it.ItemText, ForegroundColor = new Color(it.Clr) }); return false; });
            return fs;
        }

        public object ConvertBack(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
    public class BackgroundSelectedConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
#if __ANDROID__
            return ((bool)value ? System.Drawing.Color.Beige : System.Drawing.Color.White);
#else
            return ((bool)value ? new Color(245, 245, 220) : Color.White);
#endif
        }

        public object ConvertBack(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
    public class ArabicFlowConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            return (bool)value ? TextAlignment.End : TextAlignment.Start;
        }

        public object ConvertBack(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
#if !__ANDROID__
    public class MyDataTemplateSelector : DataTemplateSelector
    {
        public DataTemplate MyTemplate { get; set; }

        protected override DataTemplate OnSelectTemplate(object item, BindableObject container)
        {
            return MyTemplate;
        }
    }
    public class NormalWordTemplateSelector : DataTemplateSelector
    {
        public DataTemplate WordTemplate { get; set; }
        public DataTemplate ArabicTemplate { get; set; }
        public DataTemplate NormalTemplate { get; set; }

        protected override DataTemplate OnSelectTemplate(object item, BindableObject container)
        {
            if (item.GetType() == typeof(MyChildRenderItem)) { return ((MyChildRenderItem)item).IsArabic ? ArabicTemplate : NormalTemplate; }
            return WordTemplate;
        }
    }
#endif
}
