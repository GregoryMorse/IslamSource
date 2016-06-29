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
		public WordForWordUC (object Parameter)
		{
            dynamic c = Parameter;
            Division = c.Division;
            Selection = c.Selection;
            this.BindingContext = this;
            this.ViewModel = new VirtualizingWrapPanelAdapter();
            UIChanger = new MyUIChanger();
		}
        public MyUIChanger UIChanger { get; set; }
        public VirtualizingWrapPanelAdapter ViewModel { get; set; }

        private int Division;
        private int Selection;
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
