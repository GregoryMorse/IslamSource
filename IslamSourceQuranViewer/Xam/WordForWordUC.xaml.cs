using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Xamarin.Forms;

using System.Collections;
using System.Collections.Specialized;
using System.ComponentModel;


#if __IOS__

using System.Drawing;
using Foundation;
using UIKit;

public static class TextMeterImplementation
{
    public static Xamarin.Forms.Size MeasureTextSize(string text, double width,
        double fontSize, string fontName = null)
    {
        var nsText = new NSString(text);
        var boundSize = new SizeF((float)width, float.MaxValue);
        var options = NSStringDrawingOptions.UsesFontLeading |
            NSStringDrawingOptions.UsesLineFragmentOrigin;

        if (fontName == null)
        {
            fontName = "HelveticaNeue";
        }

        var attributes = new UIStringAttributes
        {
            Font = UIFont.FromName(fontName, (float)fontSize)
        };

        var sizeF = nsText.GetBoundingRect(boundSize, options, attributes, null).Size;

        return new Xamarin.Forms.Size((double)sizeF.Width, (double)sizeF.Height);
    }
}

#endif

#if __ANDROID__

using Android.Widget;
using Android.Util;
using Android.Views;
using Android.Graphics;

	public static class TextMeterImplementation
	{
		private static Typeface textTypeface;

		public static Xamarin.Forms.Size MeasureTextSize(string text, double width,
			double fontSize, string fontName = null)
		{
			var textView = new TextView(global::Android.App.Application.Context);
			textView.Typeface = GetTypeface(fontName);
			textView.SetText(text, TextView.BufferType.Normal);
			textView.SetTextSize(ComplexUnitType.Px, (float)fontSize);

			int widthMeasureSpec = Android.Views.View.MeasureSpec.MakeMeasureSpec(
				(int)width, MeasureSpecMode.AtMost);
			int heightMeasureSpec = Android.Views.View.MeasureSpec.MakeMeasureSpec(
				0, MeasureSpecMode.Unspecified);

			textView.Measure(widthMeasureSpec, heightMeasureSpec);

			return new Xamarin.Forms.Size((double)textView.MeasuredWidth,
				(double)textView.MeasuredHeight);
		}

		private static Typeface GetTypeface(string fontName)
		{
			if (fontName == null)
			{
				return Typeface.Default;
			}

			if (textTypeface == null)
			{
				textTypeface = Typeface.Create(fontName, TypefaceStyle.Normal);
			}

			return textTypeface;
		}
	}

#endif

namespace IslamSourceQuranViewer.Xam
{
	public partial class WordForWordUC : ContentPage
	{
		public WordForWordUC (object Parameter)
		{
            Division = Parameter.Division;
            Selection = Parameter.Selection;
            this.DataContext = this;
            this.ViewModel = new VirtualizingWrapPanelAdapter();
            UIChanger = new MyUIChanger();
            InitializeComponent ();
		}
        private int Division;
        private int Selection;
	}

    public class XamarinWrapPanel : Layout<View>
    {
        private static event EventHandler<NotifyCollectionChangedEventArgs> _collectionChanged;
        /// <summary>
        /// Backing Storage for the Orientation property
        /// </summary>
        public static readonly BindableProperty OrientationProperty =
            BindableProperty.Create<XamarinWrapPanel, StackOrientation>(w => w.Orientation, StackOrientation.Vertical,
                propertyChanged: (bindable, oldvalue, newvalue) => ((XamarinWrapPanel)bindable).OnSizeChanged());

        /// <summary>
        /// Orientation (Horizontal or Vertical)
        /// </summary>
        public StackOrientation Orientation
        {
            get { return (StackOrientation)GetValue(OrientationProperty); }
            set { SetValue(OrientationProperty, value); }
        }

        /// <summary>
        /// Backing Storage for the Spacing property
        /// </summary>
        public static readonly BindableProperty SpacingProperty =
            BindableProperty.Create<XamarinWrapPanel, double>(w => w.Spacing, 6,
                propertyChanged: (bindable, oldvalue, newvalue) => ((XamarinWrapPanel)bindable).OnSizeChanged());

        /// <summary>
        /// Spacing added between elements (both directions)
        /// </summary>
        /// <value>The spacing.</value>
        public double Spacing
        {
            get { return (double)GetValue(SpacingProperty); }
            set { SetValue(SpacingProperty, value); }
        }

        /// <summary>
        /// Backing Storage for the Spacing property
        /// </summary>
        public static readonly BindableProperty ItemTemplateProperty =
            BindableProperty.Create<XamarinWrapPanel, DataTemplate>(w => w.ItemTemplate, null,
                propertyChanged: (bindable, oldvalue, newvalue) => ((XamarinWrapPanel)bindable).OnSizeChanged());

        /// <summary>
        /// Spacing added between elements (both directions)
        /// </summary>
        /// <value>The spacing.</value>
        public DataTemplate ItemTemplate
        {
            get { return (DataTemplate)GetValue(ItemTemplateProperty); }
            set { SetValue(ItemTemplateProperty, value); }
        }

        /// <summary>
        /// Backing Storage for the Spacing property
        /// </summary>
        public static readonly BindableProperty ItemsSourceProperty =
            BindableProperty.Create<XamarinWrapPanel, IEnumerable>(w => w.ItemsSource, null,
                propertyChanged: ItemsSource_OnPropertyChanged);

        /// <summary>
        /// Spacing added between elements (both directions)
        /// </summary>
        /// <value>The spacing.</value>
        public IEnumerable ItemsSource
        {
            get { return (IEnumerable)GetValue(ItemsSourceProperty); }
            set { SetValue(ItemsSourceProperty, value); }
        }

        private static void ItemsSource_OnPropertyChanged(BindableObject bindable, IEnumerable oldvalue, IEnumerable newvalue)
        {
            if (oldvalue != null)
            {
                var coll = (INotifyCollectionChanged)oldvalue;
                // Unsubscribe from CollectionChanged on the old collection
                coll.CollectionChanged -= ItemsSource_OnItemChanged;
            }

            if (newvalue != null)
            {
                var coll = (INotifyCollectionChanged)newvalue;
                // Subscribe to CollectionChanged on the new collection
                coll.CollectionChanged += ItemsSource_OnItemChanged;
            }
        }

        public XamarinWrapPanel()
        {
            _collectionChanged += OnCollectionChanged;
        }

        private void OnCollectionChanged(object sender, NotifyCollectionChangedEventArgs args)
        {
            foreach (object item in args.NewItems)
            {
                var child = ItemTemplate.CreateContent() as View;
                if (child == null)
                    return;

                child.BindingContext = item;
                Children.Add(child);
            }
        }

        private static void ItemsSource_OnItemChanged(object sender, NotifyCollectionChangedEventArgs e)
        {
            if (_collectionChanged != null)
                _collectionChanged(null, e);
        }

        /// <summary>
        /// This is called when the spacing or orientation properties are changed - it forces
        /// the control to go back through a layout pass.
        /// </summary>
        private void OnSizeChanged()
        {
            ForceLayout();
        }

        /// <summary>
        /// This method is called during the measure pass of a layout cycle to get the desired size of an element.
        /// </summary>
        /// <param name="widthConstraint">The available width for the element to use.</param>
        /// <param name="heightConstraint">The available height for the element to use.</param>
        protected override SizeRequest OnSizeRequest(double widthConstraint, double heightConstraint)
        {
            if (WidthRequest > 0)
                widthConstraint = Math.Min(widthConstraint, WidthRequest);
            if (HeightRequest > 0)
                heightConstraint = Math.Min(heightConstraint, HeightRequest);

            double internalWidth = double.IsPositiveInfinity(widthConstraint) ? double.PositiveInfinity : Math.Max(0, widthConstraint);
            double internalHeight = double.IsPositiveInfinity(heightConstraint) ? double.PositiveInfinity : Math.Max(0, heightConstraint);

            return Orientation == StackOrientation.Vertical
                ? DoVerticalMeasure(internalWidth, internalHeight)
                    : DoHorizontalMeasure(internalWidth, internalHeight);

        }

        /// <summary>
        /// Does the vertical measure.
        /// </summary>
        /// <returns>The vertical measure.</returns>
        /// <param name="widthConstraint">Width constraint.</param>
        /// <param name="heightConstraint">Height constraint.</param>
        private SizeRequest DoVerticalMeasure(double widthConstraint, double heightConstraint)
        {
            int columnCount = 1;

            double width = 0;
            double height = 0;
            double minWidth = 0;
            double minHeight = 0;
            double heightUsed = 0;

            foreach (var item in Children)
            {
                var size = item.GetSizeRequest(widthConstraint, heightConstraint);
                width = Math.Max(width, size.Request.Width);

                var newHeight = height + size.Request.Height + Spacing;
                if (newHeight > heightConstraint)
                {
                    columnCount++;
                    heightUsed = Math.Max(height, heightUsed);
                    height = size.Request.Height;
                }
                else
                    height = newHeight;

                minHeight = Math.Max(minHeight, size.Minimum.Height);
                minWidth = Math.Max(minWidth, size.Minimum.Width);
            }

            if (columnCount > 1)
            {
                height = Math.Max(height, heightUsed);
                width *= columnCount;  // take max width
            }

            return new SizeRequest(new Size(width, height), new Size(minWidth, minHeight));
        }

        /// <summary>
        /// Does the horizontal measure.
        /// </summary>
        /// <returns>The horizontal measure.</returns>
        /// <param name="widthConstraint">Width constraint.</param>
        /// <param name="heightConstraint">Height constraint.</param>
        private SizeRequest DoHorizontalMeasure(double widthConstraint, double heightConstraint)
        {
            int rowCount = 1;

            double width = 0;
            double height = 0;
            double minWidth = 0;
            double minHeight = 0;
            double widthUsed = 0;

            foreach (var item in Children)
            {
                var size = item.GetSizeRequest(widthConstraint, heightConstraint);
                height = Math.Max(height, size.Request.Height);

                var newWidth = width + size.Request.Width + Spacing;
                if (newWidth > widthConstraint)
                {
                    rowCount++;
                    widthUsed = Math.Max(width, widthUsed);
                    width = size.Request.Width;
                }
                else
                    width = newWidth;

                minHeight = Math.Max(minHeight, size.Minimum.Height);
                minWidth = Math.Max(minWidth, size.Minimum.Width);
            }

            if (rowCount > 1)
            {
                width = Math.Max(width, widthUsed);
                height = (height + Spacing) * rowCount - Spacing; // via MitchMilam 
            }

            return new SizeRequest(new Size(width, height), new Size(minWidth, minHeight));
        }

        /// <summary>
        /// Positions and sizes the children of a Layout.
        /// </summary>
        /// <param name="x">A value representing the x coordinate of the child region bounding box.</param>
        /// <param name="y">A value representing the y coordinate of the child region bounding box.</param>
        /// <param name="width">A value representing the width of the child region bounding box.</param>
        /// <param name="height">A value representing the height of the child region bounding box.</param>
        protected override void LayoutChildren(double x, double y, double width, double height)
        {
            if (Orientation == StackOrientation.Vertical)
            {
                double colWidth = 0;
                double yPos = y, xPos = x;

                foreach (var child in Children.Where(c => c.IsVisible))
                {
                    var request = child.GetSizeRequest(width, height);

                    double childWidth = request.Request.Width;
                    double childHeight = request.Request.Height;
                    colWidth = Math.Max(colWidth, childWidth);

                    if (yPos + childHeight > height)
                    {
                        yPos = y;
                        xPos += colWidth + Spacing;
                        colWidth = 0;
                    }

                    var region = new Rectangle(xPos, yPos, childWidth, childHeight);
                    LayoutChildIntoBoundingRegion(child, region);
                    yPos += region.Height + Spacing;
                }
            }
            else
            {
                double rowHeight = 0;
                double yPos = y, xPos = x;

                foreach (var child in Children.Where(c => c.IsVisible))
                {
                    var request = child.GetSizeRequest(width, height);

                    double childWidth = request.Request.Width;
                    double childHeight = request.Request.Height;
                    rowHeight = Math.Max(rowHeight, childHeight);

                    if (xPos + childWidth > width)
                    {
                        xPos = x;
                        yPos += rowHeight + Spacing;
                        rowHeight = 0;
                    }

                    var region = new Rectangle(xPos, yPos, childWidth, childHeight);
                    LayoutChildIntoBoundingRegion(child, region);
                    xPos += region.Width + Spacing;
                }

            }
        }
    }

    public class MyUIChanger : INotifyPropertyChanged
    {
        public MyUIChanger()
        {
        }
        //public double FontSize
        //{
        //    get
        //    {
        //        return AppSettings.dFontSize;
        //    }
        //    set
        //    {
        //        AppSettings.dFontSize = value;
        //        //if (_DWArabicFormat != null) _DWArabicFormat.Dispose();
        //        //_DWArabicFormat = null;
        //        PropertyChanged(this, new PropertyChangedEventArgs("FontSize"));
        //    }
        //}
        //public string FontFamily
        //{
        //    get
        //    {
        //        return AppSettings.strSelectedFont;
        //    }
        //}
        //public double OtherFontSize
        //{
        //    get
        //    {
        //        return AppSettings.dOtherFontSize;
        //    }
        //    set
        //    {
        //        AppSettings.dOtherFontSize = value;
        //        //if (_DWNormalFormat != null) _DWNormalFormat.Dispose();
        //        //_DWNormalFormat = null;
        //        PropertyChanged(this, new PropertyChangedEventArgs("OtherFontSize"));
        //    }
        //}
        //public string OtherFontFamily
        //{
        //    get
        //    {
        //        return AppSettings.strOtherSelectedFont;
        //    }
        //}
        public static void CleanupDW()
        {
            //if (_DWNormalFormat != null) _DWNormalFormat.Dispose();
            //_DWNormalFormat = null;
            //if (_DWArabicFormat != null) _DWArabicFormat.Dispose();
            //_DWArabicFormat = null;
            //if (_DWFactory != null) _DWFactory.Dispose();
            //_DWFactory = null;
        }
        //private static SharpDX.DirectWrite.Factory _DWFactory;
        //public static SharpDX.DirectWrite.Factory DWFactory { get { if (_DWFactory == null) _DWFactory = new SharpDX.DirectWrite.Factory(); return _DWFactory; } }
        //private static SharpDX.DirectWrite.TextFormat _DWArabicFormat;
        //public static SharpDX.DirectWrite.TextFormat DWArabicFormat { get { if (_DWArabicFormat == null) _DWArabicFormat = new SharpDX.DirectWrite.TextFormat(MyUIChanger.DWFactory, AppSettings.strSelectedFont, (float)AppSettings.dFontSize); return _DWArabicFormat; } }
        //private static SharpDX.DirectWrite.TextFormat _DWNormalFormat;
        //public static SharpDX.DirectWrite.TextFormat DWNormalFormat { get { if (_DWNormalFormat == null) _DWNormalFormat = new SharpDX.DirectWrite.TextFormat(MyUIChanger.DWFactory, AppSettings.strOtherSelectedFont, (float)AppSettings.dOtherFontSize); return _DWNormalFormat; } }
        private double _MaxWidth;
        public double MaxWidth
        {
            get
            {
                return _MaxWidth;
            }
            set
            {
                _MaxWidth = value;
                if (PropertyChanged != null) PropertyChanged(this, new PropertyChangedEventArgs("MaxWidth"));
            }
        }

        #region Implementation of INotifyPropertyChanged

        public event PropertyChangedEventHandler PropertyChanged;

        #endregion
    }
    public class MyRenderModel : INotifyPropertyChanged
    {
        public MyRenderModel(List<MyRenderItem> NewRenderItems)
        {
            RenderItems = NewRenderItems;
            MaxWidth = CalculateWidth();
        }
        private double _MaxWidth;
        public double MaxWidth { get { return _MaxWidth; } set { _MaxWidth = value; if (PropertyChanged != null) PropertyChanged(this, new PropertyChangedEventArgs("MaxWidth")); } }
        private double CalculateWidth()
        {
            return RenderItems.Select((Item) => Item.MaxWidth).Sum();
        }
        private List<MyRenderItem> _RenderItems;
        public List<MyRenderItem> RenderItems { get { return _RenderItems; } set { _RenderItems = value.ToList(); if (PropertyChanged != null) PropertyChanged(this, new PropertyChangedEventArgs("RenderItems")); } }
        #region Implementation of INotifyPropertyChanged

        public event PropertyChangedEventHandler PropertyChanged;

        #endregion
    }
    public class MyRenderItem : INotifyPropertyChanged
    {
        public MyRenderItem(XMLRender.RenderArray.RenderItem RendItem)
        {
            if (RendItem.TextItems.First().DisplayClass == XMLRender.RenderArray.RenderDisplayClass.eReference) { Chapter = ((int[])RendItem.TextItems.First().Text)[0]; Verse = ((int[])RendItem.TextItems.First().Text)[1]; Word = ((int[])RendItem.TextItems.First().Text).Count() == 2 ? -1 : ((int[])RendItem.TextItems.First().Text)[2]; } else { Chapter = -1; Verse = -1; Word = -1; }
            Items = System.Linq.Enumerable.Select(RendItem.TextItems.GroupBy((MainItems) => (MainItems.DisplayClass == XMLRender.RenderArray.RenderDisplayClass.eArabic || MainItems.DisplayClass == XMLRender.RenderArray.RenderDisplayClass.eLTR || MainItems.DisplayClass == XMLRender.RenderArray.RenderDisplayClass.eRTL || MainItems.DisplayClass == XMLRender.RenderArray.RenderDisplayClass.eTransliteration) ? (object)MainItems.DisplayClass : (object)MainItems), (Arr) => (Arr.First().Text.GetType() == typeof(List<XMLRender.RenderArray.RenderItem>)) ? (object)new VirtualizingWrapPanelAdapter() { RenderModels = new List<MyRenderModel>() { new MyRenderModel(System.Linq.Enumerable.Select((List<XMLRender.RenderArray.RenderItem>)Arr.First().Text, (ArrRend) => new MyRenderItem((XMLRender.RenderArray.RenderItem)ArrRend)).ToList()) } } : (Arr.First().Text.GetType() == typeof(string) ? (object)new MyChildRenderItem(System.Linq.Enumerable.Select(Arr, (ArrItem) => new MyChildRenderBlockItem() { ItemText = (string)ArrItem.Text, Clr = ArrItem.Clr }).ToList(), Arr.First().DisplayClass == XMLRender.RenderArray.RenderDisplayClass.eArabic, Arr.First().DisplayClass == XMLRender.RenderArray.RenderDisplayClass.eArabic || Arr.First().DisplayClass == XMLRender.RenderArray.RenderDisplayClass.eRTL) : null)).Where(Arr => Arr != null).ToList();
            MaxWidth = CalculateWidth();
        }
        public double MaxWidth { get; set; }
        private double CalculateWidth()
        {
            if (Items.Count() == 0) return 0.0;
            return _Items.Select((Item) => Item.GetType() == typeof(MyChildRenderItem) ? ((MyChildRenderItem)Item).MaxWidth : ((VirtualizingWrapPanelAdapter)Item).RenderModels.Select((It) => It.MaxWidth).Max()).Max();
        }
        public void RegroupRenderModels(double maxWidth)
        {
            _Items.FirstOrDefault((Item) => { if (Item.GetType() != typeof(MyChildRenderItem)) { ((VirtualizingWrapPanelAdapter)Item).RegroupRenderModels(maxWidth); } return false; });
        }
        private List<object> _Items;
        public List<object> Items { get { return _Items; } set { _Items = value.ToList(); if (PropertyChanged != null) PropertyChanged(this, new PropertyChangedEventArgs("Items")); } }
        private bool _IsSelected;
        public bool IsSelected { get { return _IsSelected; } set { _IsSelected = value; if (PropertyChanged != null) PropertyChanged(this, new PropertyChangedEventArgs("IsSelected")); } }
        public int Chapter;
        public int Verse;
        public int Word;
        public string VerseText { get { return Chapter.ToString() + ":" + Verse.ToString(); } }
        public string VerseWordText { get { return Chapter.ToString() + ":" + Verse.ToString() + (Word == -1 ? string.Empty : (":" + Word.ToString())); } }
        public string GetText { get { return String.Join(" ", _Items.Select((Item) => Item.GetType() == typeof(MyChildRenderItem) ? ((MyChildRenderItem)Item).GetText : String.Join(string.Empty, ((VirtualizingWrapPanelAdapter)Item).RenderModels.Select((It) => String.Join(string.Empty, It.RenderItems.Select((RecIt) => RecIt.GetText)))))) + " " + VerseWordText; } }
        #region Implementation of INotifyPropertyChanged

        public event PropertyChangedEventHandler PropertyChanged;

        #endregion
    }
    public class VirtualizingWrapPanelAdapter : INotifyPropertyChanged
    {
        private CollectionViewSource _RenderSource;
        public ICollectionView RenderSource
        {
            get
            {
                if (_RenderSource == null) { _RenderSource = new CollectionViewSource(); BindingOperations.SetBinding(_RenderSource, CollectionViewSource.SourceProperty, new Binding() { Source = this, Path = new PropertyPath("RenderModels"), Mode = BindingMode.OneWay }); }
                return _RenderSource.View;
            }
        }
        private List<MyRenderModel> _RenderModels;
        public List<MyRenderModel> RenderModels
        {
            get
            {
                if (_RenderModels == null) { _RenderModels = new List<MyRenderModel>(); }
                return _RenderModels;
            }
            set
            {
                _RenderModels = value;
                VerseReferences = value.SelectMany((Model) => Model.RenderItems).Where((Item) => Item.Chapter != -1).ToList();
                if (PropertyChanged != null) PropertyChanged(this, new PropertyChangedEventArgs("RenderModels"));
                if (PropertyChanged != null) PropertyChanged(this, new PropertyChangedEventArgs("RenderSource"));
            }
        }
        public void RegroupRenderModels(double maxWidth)
        {
            if (_RenderModels != null) { RenderModels = GroupRenderModels(_RenderModels.SelectMany((Item) => Item.RenderItems).ToList(), maxWidth); }
        }
        public static List<MyRenderModel> GroupRenderModels(List<MyRenderItem> value, double maxWidth)
        {
            double width = 0.0;
            int groupIndex = 0;
            List<int[]> GroupIndexes = new List<int[]>();
            value.FirstOrDefault((Item) => {
                Item.RegroupRenderModels(maxWidth); double itemWidth = Math.Min(Item.MaxWidth, maxWidth); if (width + itemWidth > maxWidth) { width = itemWidth; groupIndex++; } else { width += itemWidth; }
                GroupIndexes.Add(new int[] { groupIndex, GroupIndexes.Count }); return false;
            });
            return GroupIndexes.GroupBy((Item) => Item[0], (Item) => value.ElementAt(Item[1])).Select((Item) => new MyRenderModel(Item.ToList())).ToList();
        }
        private List<MyRenderItem> _VerseReferences;
        public List<MyRenderItem> VerseReferences { get { return _VerseReferences; } set { _VerseReferences = value; if (PropertyChanged != null) PropertyChanged(this, new PropertyChangedEventArgs("VerseReferences")); } }
        private MyRenderItem _CurrentVerse;
        public MyRenderItem CurrentVerse { get { return _CurrentVerse; } set { if (_CurrentVerse != null) _CurrentVerse.IsSelected = false; _CurrentVerse = value; _CurrentVerse.IsSelected = true; if (PropertyChanged != null) PropertyChanged(this, new PropertyChangedEventArgs("CurrentVerse")); } }
        #region Implementation of INotifyPropertyChanged

        public event PropertyChangedEventHandler PropertyChanged;

        #endregion
    }
    public class RunFormattedText : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            List<MyChildRenderBlockItem> val = value as List<MyChildRenderBlockItem>;
            FormattedString fs = new FormattedString();
            val.FirstOrDefault((it) => { fs.Add(new Span() { Text = it.ItemText, ForegroundColor = it.Color }); return false; });
            return fs;
        }

        public object ConvertBack(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
    public class MyChildRenderItem : INotifyPropertyChanged
    {
        private double _MaxWidth;
        public double MaxWidth { get { return _MaxWidth; } set { _MaxWidth = value; if (PropertyChanged != null) PropertyChanged(this, new PropertyChangedEventArgs("MaxWidth")); } }
        public static double CalculateWidth(string text, bool IsArabic, float maxWidth, float maxHeight)
        {
#if WINDOWS_PHONE
            SharpDX.DirectWrite.TextLayout layout = new SharpDX.DirectWrite.TextLayout(MyUIChanger.DWFactory, text, IsArabic ? MyUIChanger.DWArabicFormat : MyUIChanger.DWNormalFormat, maxWidth, maxHeight);
            double width = layout.Metrics.WidthIncludingTrailingWhitespace + layout.Metrics.Left;
            layout.Dispose();
#else
            double width = TextMeterImplementation.MeasureTextSize(text, maxWidth, MyUIChanger.FontSize, MyUIChanger.FontFamily);
#endif
            return width;
        }
        public double CalculateWidth()
        {
            //5 margin on both sides
            return 5 + 5 + 1 + 1 + CalculateWidth(string.Join(String.Empty, ItemRuns.Select((Item) => Item.ItemText)), IsArabic, (float)float.MaxValue, float.MaxValue);
        }
        public string GetText { get { return String.Join(string.Empty, _ItemRuns.Select((Item) => Item.ItemText)); } }
#region Implementation of INotifyPropertyChanged

        public event PropertyChangedEventHandler PropertyChanged;

        #endregion
#if WINDOWS_PHONE
        [System.Runtime.InteropServices.StructLayout(System.Runtime.InteropServices.LayoutKind.Sequential)]
        public struct LayoutInfo
        {
            public Rect Rect;
            public float Baseline;
            public int nChar;
            public List<List<List<LayoutInfo>>> Bounds;
            public LayoutInfo(Rect NewRect, float NewBaseline, int NewNChar, List<List<List<LayoutInfo>>> NewBounds)
            {
                this = new LayoutInfo();
                this.Rect = NewRect;
                this.Baseline = NewBaseline;
                this.nChar = NewNChar;
                this.Bounds = NewBounds;
            }
        }

        [System.Runtime.InteropServices.StructLayout(System.Runtime.InteropServices.LayoutKind.Sequential)]
        public struct CharPosInfo
        {
            public int Index;
            public int Length;
            public float Width;
            public float PriorWidth;
            public float X;
            public float Y;
            public float Height;
        }

        const int ERROR_INSUFFICIENT_BUFFER = 122;

        public class TextSource : SharpDX.DirectWrite.TextAnalysisSource
        {
            // Fields
            public SharpDX.DirectWrite.Factory _Factory;
            private string _Str;
            private bool disposedValue;

            // Methods
            public TextSource(string Str, SharpDX.DirectWrite.Factory Factory)
            {
                this._Str = Str;
                this._Factory = Factory;
            }

            public void Dispose()
            {
                this.Dispose(true);
                GC.SuppressFinalize(this);
            }

            protected virtual void Dispose(bool disposing)
            {
                if (!this.disposedValue && disposing)
                {
                }
                this.disposedValue = true;
            }

            public string GetLocaleName(int textPosition, out int textLength)
            {
                textLength = _Str.Length - textPosition;
                return System.Globalization.CultureInfo.CurrentCulture.Name;
            }

            public SharpDX.DirectWrite.NumberSubstitution GetNumberSubstitution(int textPosition, out int textLength)
            {
                textLength = _Str.Length - textPosition;
                return new SharpDX.DirectWrite.NumberSubstitution(this._Factory, SharpDX.DirectWrite.NumberSubstitutionMethod.None, null, true);
            }

            public string GetTextAtPosition(int textPosition)
            {
                return this._Str.Substring(textPosition);
            }

            public string GetTextBeforePosition(int textPosition)
            {
                return this._Str.Substring(0x0, textPosition - 0x1);
            }

            // Properties
            public SharpDX.DirectWrite.ReadingDirection ReadingDirection
            {
                get
                {
                    return SharpDX.DirectWrite.ReadingDirection.RightToLeft;
                }
            }


            public System.IDisposable Shadow { get; set; }
        }

    public class TextSink : SharpDX.DirectWrite.TextAnalysisSink
    {
        public byte _explicitLevel;
        public SharpDX.DirectWrite.LineBreakpoint[] _lineBreakpoints;
        public SharpDX.DirectWrite.NumberSubstitution _numberSubstitution;
        public byte _resolvedLevel;
        public SharpDX.DirectWrite.ScriptAnalysis _scriptAnalysis;
        private bool disposedValue;

        public void Dispose()
        {
            this.Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (!this.disposedValue && disposing)
            {
            }
            this.disposedValue = true;
        }

        public void SetBidiLevel(int textPosition, int textLength, byte explicitLevel, byte resolvedLevel)
        {
            this._explicitLevel = explicitLevel;
            this._resolvedLevel = resolvedLevel;
        }

        public void SetLineBreakpoints(int textPosition, int textLength, SharpDX.DirectWrite.LineBreakpoint[] lineBreakpoints)
        {
            this._lineBreakpoints = lineBreakpoints;
        }

        public void SetNumberSubstitution(int textPosition, int textLength, SharpDX.DirectWrite.NumberSubstitution numberSubstitution)
        {
            this._numberSubstitution = numberSubstitution;
        }

        public void SetScriptAnalysis(int textPosition, int textLength, SharpDX.DirectWrite.ScriptAnalysis scriptAnalysis)
        {
            this._scriptAnalysis = scriptAnalysis;
        }

        public IDisposable Shadow { get; set; }
    }
        //LOGFONT struct
        [System.Runtime.InteropServices.StructLayout(System.Runtime.InteropServices.LayoutKind.Sequential, CharSet = System.Runtime.InteropServices.CharSet.Unicode)]
        public class LOGFONT
        {
            public const int LF_FACESIZE = 32;
            public int lfHeight;
            public int lfWidth;
            public int lfEscapement;
            public int lfOrientation;
            public int lfWeight;
            public byte lfItalic;
            public byte lfUnderline;
            public byte lfStrikeOut;
            public byte lfCharSet;
            public byte lfOutPrecision;
            public byte lfClipPrecision;
            public byte lfQuality;
            public byte lfPitchAndFamily;
            [System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.ByValTStr, SizeConst = LF_FACESIZE)]
            public string lfFaceName;
        }
        public static short[] GetWordDiacriticClusters(string Str, string useFont, float fontSize, bool IsRTL)
        {
            if (Str == string.Empty)
            {
                return null;
            }
            SharpDX.DirectWrite.ScriptAnalysis scriptAnalysis;
            TextSink analysisSink = new TextSink();
            TextSource analysisSource = new TextSource(Str, MyUIChanger.DWFactory);
            MyUIChanger.DWAnalyzer.AnalyzeScript(analysisSource, 0, Str.Length, analysisSink);
            scriptAnalysis = analysisSink._scriptAnalysis;
            int maxGlyphCount = ((Str.Length * 3) / 2) + 0x10;
            short[] clusterMap = new short[(Str.Length - 1) + 1];
            SharpDX.DirectWrite.ShapingTextProperties[] textProps = new SharpDX.DirectWrite.ShapingTextProperties[(Str.Length - 1) + 1];
            short[] glyphIndices = new short[(maxGlyphCount - 1) + 1];
            SharpDX.DirectWrite.ShapingGlyphProperties[] glyphProps = new SharpDX.DirectWrite.ShapingGlyphProperties[(maxGlyphCount - 1) + 1];
            int actualGlyphCount = 0;
            while (true)
            {
                try
                {
                    MyUIChanger.DWAnalyzer.GetGlyphs(Str, Str.Length, MyUIChanger.DWFontFace, false, IsRTL, scriptAnalysis, null, null, new SharpDX.DirectWrite.FontFeature[][] { MyUIChanger.DWFeatureArray }, new int[] { Str.Length }, maxGlyphCount, clusterMap, textProps, glyphIndices, glyphProps, out actualGlyphCount);
                    break;
                }
                catch (SharpDX.SharpDXException exception)
                {
                    if (exception.ResultCode == SharpDX.Result.GetResultFromWin32Error(0x7a))
                    {
                        maxGlyphCount *= 2;
                        glyphIndices = new short[(maxGlyphCount - 1) + 1];
                        glyphProps = new SharpDX.DirectWrite.ShapingGlyphProperties[(maxGlyphCount - 1) + 1];
                    }
                }
            }
            Array.Resize(ref glyphIndices, (actualGlyphCount - 1) + 1);
            Array.Resize(ref glyphProps, (actualGlyphCount - 1) + 1);
            float[] glyphAdvances = new float[(actualGlyphCount - 1) + 1];
            SharpDX.DirectWrite.GlyphOffset[] glyphOffsets = new SharpDX.DirectWrite.GlyphOffset[(actualGlyphCount - 1) + 1];
            SharpDX.DirectWrite.FontFeature[][] features = new SharpDX.DirectWrite.FontFeature[][] { MyUIChanger.DWFeatureArray };
            int[] featureRangeLengths = new int[] { Str.Length };
            MyUIChanger.DWAnalyzer.GetGlyphPlacements(Str, clusterMap, textProps, Str.Length, glyphIndices, glyphProps, actualGlyphCount, MyUIChanger.DWFontFace, fontSize, false, IsRTL, scriptAnalysis, null, features, featureRangeLengths, glyphAdvances, glyphOffsets);
            analysisSource.Shadow.Dispose();
            analysisSink.Shadow.Dispose();
            analysisSource.Dispose();
            analysisSource._Factory = null;
            analysisSink.Dispose();
            return clusterMap;
        }
        public static Size GetWordDiacriticPositionsDWrite(string Str, string useFont, float fontSize, char[] Forms, bool IsRTL, ref float BaseLine, ref CharPosInfo[] Pos)
        {
            if (Str == string.Empty)
            {
                return new Size(0f, 0f);
            }
            SharpDX.DirectWrite.TextAnalyzer analyzer = new SharpDX.DirectWrite.TextAnalyzer(MyUIChanger.DWFactory);
            LOGFONT lf = new LOGFONT();
            lf.lfFaceName = useFont;
            float pointSize = fontSize * Windows.Graphics.Display.DisplayInformation.GetForCurrentView().RawDpiY / 72.0f;
            lf.lfHeight = (int)fontSize;
            lf.lfQuality = 5; //clear type
            SharpDX.DirectWrite.Font font = MyUIChanger.DWFactory.GdiInterop.FromLogFont(lf);
            SharpDX.DirectWrite.FontFace fontFace = new SharpDX.DirectWrite.FontFace(font);
            SharpDX.DirectWrite.ScriptAnalysis scriptAnalysis = new SharpDX.DirectWrite.ScriptAnalysis();
            TextSink analysisSink = new TextSink();
            TextSource analysisSource = new TextSource(Str, MyUIChanger.DWFactory);
            analyzer.AnalyzeScript(analysisSource, 0, Str.Length, analysisSink);
            scriptAnalysis = analysisSink._scriptAnalysis;
            int maxGlyphCount = ((Str.Length * 3) / 2) + 0x10;
            short[] clusterMap = new short[(Str.Length - 1) + 1];
            SharpDX.DirectWrite.ShapingTextProperties[] textProps = new SharpDX.DirectWrite.ShapingTextProperties[(Str.Length - 1) + 1];
            short[] glyphIndices = new short[(maxGlyphCount - 1) + 1];
            SharpDX.DirectWrite.ShapingGlyphProperties[] glyphProps = new SharpDX.DirectWrite.ShapingGlyphProperties[(maxGlyphCount - 1) + 1];
            int actualGlyphCount = 0;
            SharpDX.DirectWrite.FontFeature[] featureArray = new SharpDX.DirectWrite.FontFeature[] { new SharpDX.DirectWrite.FontFeature(SharpDX.DirectWrite.FontFeatureTag.GlyphCompositionDecomposition, 1), new SharpDX.DirectWrite.FontFeature(SharpDX.DirectWrite.FontFeatureTag.DiscretionaryLigatures, 0), new SharpDX.DirectWrite.FontFeature(SharpDX.DirectWrite.FontFeatureTag.StandardLigatures, 0), new SharpDX.DirectWrite.FontFeature(SharpDX.DirectWrite.FontFeatureTag.ContextualAlternates, 0), new SharpDX.DirectWrite.FontFeature(SharpDX.DirectWrite.FontFeatureTag.StylisticSet1, 0) };
            while (true)
            {
                try
                {
                    analyzer.GetGlyphs(Str, Str.Length, fontFace, false, IsRTL, scriptAnalysis, null, null, new SharpDX.DirectWrite.FontFeature[][] { featureArray }, new int[] { Str.Length }, maxGlyphCount, clusterMap, textProps, glyphIndices, glyphProps, out actualGlyphCount);
                    break;
                }
                catch (SharpDX.SharpDXException exception)
                {
                    if (exception.ResultCode == SharpDX.Result.GetResultFromWin32Error(0x7a))
                    {
                        maxGlyphCount *= 2;
                        glyphIndices = new short[(maxGlyphCount - 1) + 1];
                        glyphProps = new SharpDX.DirectWrite.ShapingGlyphProperties[(maxGlyphCount - 1) + 1];
                    }
                }
            }
            Array.Resize(ref glyphIndices, (actualGlyphCount - 1) + 1);
            Array.Resize(ref glyphProps, (actualGlyphCount - 1) + 1);
            float[] glyphAdvances = new float[(actualGlyphCount - 1) + 1];
            SharpDX.DirectWrite.GlyphOffset[] glyphOffsets = new SharpDX.DirectWrite.GlyphOffset[(actualGlyphCount - 1) + 1];
            SharpDX.DirectWrite.FontFeature[][] features = new SharpDX.DirectWrite.FontFeature[][] { featureArray };
            int[] featureRangeLengths = new int[] { Str.Length };
            analyzer.GetGlyphPlacements(Str, clusterMap, textProps, Str.Length, glyphIndices, glyphProps, actualGlyphCount, fontFace, fontSize, false, IsRTL, scriptAnalysis, null, features, featureRangeLengths, glyphAdvances, glyphOffsets);
            List<CharPosInfo> list = new List<CharPosInfo>();
            float PriorWidth = 0f;
            int RunStart = 0;
            int RunRes = clusterMap[0];
            if (IsRTL)
            {
                XMLRender.ArabicData.LigatureInfo[] array = XMLRender.ArabicData.GetLigatures(Str, false, Forms);
                for (int CharCount = 0; CharCount < clusterMap.Length - 1; CharCount++)
                {
                    int RunCount = 0;
                    for (int ResCount = clusterMap[CharCount]; ResCount <= ((CharCount == (clusterMap.Length - 1)) ? (actualGlyphCount - 1) : (clusterMap[CharCount + 1] - 1)); ResCount++)
                    {
                        if ((glyphAdvances[ResCount] == 0f) & ((clusterMap.Length <= (RunStart + RunCount)) || (clusterMap[RunStart] == clusterMap[RunStart + RunCount])))
                        {
                            int Index = Array.FindIndex<XMLRender.ArabicData.LigatureInfo>(array, (lig) => lig.Indexes[0] == RunStart + RunCount);
                            int LigLen = 1;
                            if (Index != -1)
                            {
                                while ((LigLen != array[Index].Indexes.Length) && ((array[Index].Indexes[LigLen - 1] + 1) == array[Index].Indexes[LigLen]))
                                {
                                    LigLen++;
                                }
                                if (LigLen != 1)
                                {
                                    int CheckGlyphCount = 0;
                                    short[] CheckClusterMap = new short[((RunCount + LigLen) - 1) + 1];
                                    SharpDX.DirectWrite.ShapingTextProperties[] CheckTextProps = new SharpDX.DirectWrite.ShapingTextProperties[((RunCount + LigLen) - 1) + 1];
                                    short[] CheckGlyphIndices = new short[(maxGlyphCount - 1) + 1];
                                    SharpDX.DirectWrite.ShapingGlyphProperties[] CheckGlyphProps = new SharpDX.DirectWrite.ShapingGlyphProperties[(maxGlyphCount - 1) + 1];
                                    analyzer.GetGlyphs(Str.Substring(RunStart, RunCount + LigLen), RunCount + LigLen, fontFace, false, IsRTL, scriptAnalysis, null, null, new SharpDX.DirectWrite.FontFeature[][] { featureArray }, new int[] { RunCount + LigLen }, maxGlyphCount, CheckClusterMap, CheckTextProps, CheckGlyphIndices, CheckGlyphProps, out CheckGlyphCount);
                                    if ((CheckGlyphCount != LigLen) & (CheckGlyphCount != (LigLen - (((glyphProps[RunRes].Justification != SharpDX.DirectWrite.ScriptJustify.Blank) & (glyphProps[RunRes].Justification != SharpDX.DirectWrite.ScriptJustify.ArabicBlank)) ? 0 : 1))))
                                    {
                                        LigLen = 1;
                                    }
                                }
                            }
                            if ((!glyphProps[ResCount].IsDiacritic | !glyphProps[ResCount].IsZeroWidthSpace) | !glyphProps[ResCount].IsClusterStart)
                            {
                                CharPosInfo info;
                                if (((LigLen == 1) && System.Text.RegularExpressions.Regex.Match(Str[RunStart + RunCount].ToString(), @"[\p{IsArabic}\p{IsArabicPresentationForms-A}\p{IsArabicPresentationForms-B}]").Success) & (System.Globalization.CharUnicodeInfo.GetUnicodeCategory(Str[RunStart + RunCount]) == System.Globalization.UnicodeCategory.DecimalDigitNumber))
                                {
                                    SharpDX.DirectWrite.GlyphMetrics[] _Mets = fontFace.GetDesignGlyphMetrics(glyphIndices, false);
                                    info = new CharPosInfo
                                    {
                                        Index = RunStart + RunCount,
                                        Length = (Index == -1) ? 1 : LigLen,
                                        PriorWidth = PriorWidth,
                                        Width = 2f * ((_Mets[ResCount].AdvanceWidth * pointSize) / ((float)fontFace.Metrics.DesignUnitsPerEm)),
                                        X = (glyphOffsets[ResCount].AdvanceOffset - glyphAdvances[RunRes]) - (((_Mets[ResCount].AdvanceWidth * pointSize) / ((float)fontFace.Metrics.DesignUnitsPerEm)) / 4f),
                                        Y = glyphOffsets[ResCount].AscenderOffset,
                                        Height = (((_Mets[ResCount].AdvanceHeight + _Mets[ResCount].BottomSideBearing) - _Mets[ResCount].TopSideBearing) * pointSize) / ((float)fontFace.Metrics.DesignUnitsPerEm)
                                    };
                                    list.Add(info);
                                }
                                else
                                {
                                    SharpDX.DirectWrite.GlyphMetrics[] _Mets = fontFace.GetDesignGlyphMetrics(glyphIndices, false);
                                    info = new CharPosInfo
                                    {
                                        Index = RunStart + RunCount,
                                        Length = (Index == -1) ? 1 : LigLen,
                                        PriorWidth = PriorWidth - ((((glyphProps[RunRes].Justification == SharpDX.DirectWrite.ScriptJustify.ArabicKashida) & (RunCount == 1)) & ((((CharCount == (clusterMap.Length - 1)) ? actualGlyphCount : clusterMap[CharCount + 1]) - clusterMap[CharCount]) == (CharCount - RunStart))) ? glyphAdvances[RunRes] : 0f),
                                        Width = glyphAdvances[RunRes] + ((glyphProps[RunRes].IsClusterStart & glyphProps[RunRes].IsDiacritic) ? ((_Mets[RunRes].AdvanceWidth * pointSize) / ((float)fontFace.Metrics.DesignUnitsPerEm)) : 0f),
                                        X = glyphOffsets[ResCount].AdvanceOffset,
                                        Y = glyphOffsets[ResCount].AscenderOffset + ((glyphProps[RunRes].IsClusterStart & glyphProps[RunRes].IsDiacritic) ? ((((_Mets[RunRes].AdvanceHeight - _Mets[RunRes].TopSideBearing) - _Mets[RunRes].VerticalOriginY) * pointSize) / ((float)fontFace.Metrics.DesignUnitsPerEm)) : 0f)
                                    };
                                    list.Add(info);
                                    if (((glyphProps[RunRes].Justification == SharpDX.DirectWrite.ScriptJustify.ArabicKashida) & (RunCount == 1)) & ((((CharCount == (clusterMap.Length - 1)) ? actualGlyphCount : clusterMap[CharCount + 1]) - clusterMap[CharCount]) == (CharCount - RunStart)))
                                    {
                                        info = new CharPosInfo
                                        {
                                            Index = (RunStart + RunCount) + 1,
                                            Length = (Index == -1) ? 1 : LigLen,
                                            PriorWidth = PriorWidth,
                                            Width = glyphAdvances[RunRes] + ((glyphProps[RunRes].IsClusterStart & glyphProps[RunRes].IsDiacritic) ? ((_Mets[RunRes].AdvanceWidth * pointSize) / ((float)fontFace.Metrics.DesignUnitsPerEm)) : 0f),
                                            X = glyphOffsets[ResCount].AdvanceOffset,
                                            Y = glyphOffsets[RunRes].AscenderOffset + ((glyphProps[RunRes].IsClusterStart & glyphProps[RunRes].IsDiacritic) ? ((((_Mets[RunRes].AdvanceHeight - _Mets[RunRes].TopSideBearing) - _Mets[RunRes].VerticalOriginY) * pointSize) / ((float)fontFace.Metrics.DesignUnitsPerEm)) : 0f)
                                        };
                                        list.Add(info);
                                    }
                                }
                            }
                            else
                            {
                                PriorWidth -= glyphOffsets[ResCount].AdvanceOffset;
                            }
                        }
                        if ((CharCount == (clusterMap.Length - 1)) || (clusterMap[CharCount] != clusterMap[CharCount + 1]))
                        {
                            PriorWidth += glyphAdvances[ResCount];
                            int Index = Array.FindIndex<XMLRender.ArabicData.LigatureInfo>(array, (lig) => lig.Indexes[0] == RunStart);
                            if ((Index == -1) || ((((glyphProps[ResCount].Justification != SharpDX.DirectWrite.ScriptJustify.Blank) & (glyphProps[ResCount].Justification != SharpDX.DirectWrite.ScriptJustify.ArabicBlank)) | (Array.IndexOf<int>(array[Index].Indexes, RunStart) == -1)) & ((RunStart + RunCount) != (Str.Length - 1))))
                            {
                                RunCount++;
                            }
                            if ((Index != -1) && (((glyphProps[ResCount].Justification != SharpDX.DirectWrite.ScriptJustify.Blank) & (glyphProps[ResCount].Justification != SharpDX.DirectWrite.ScriptJustify.ArabicBlank)) | (Array.IndexOf<int>(array[Index].Indexes, RunStart) == -1)))
                            {
                                while ((Array.IndexOf<int>(array[Index].Indexes, RunStart + RunCount) != -1) & ((RunStart + RunCount) != (Str.Length - 1)))
                                {
                                    RunCount++;
                                }
                            }
                            if ((clusterMap[CharCount] != ResCount) & !(glyphAdvances[ResCount] == 0f))
                            {
                                RunStart = CharCount;
                                RunCount = 0;
                                RunRes = ResCount;
                            }
                        }
                    }
                    if ((CharCount != (clusterMap.Length - 1)) && (clusterMap[CharCount] != clusterMap[CharCount + 1]))
                    {
                        RunStart = CharCount + 1;
                        if (!(glyphAdvances[clusterMap[CharCount + 1]] == 0f) | (glyphProps[clusterMap[CharCount + 1]].IsClusterStart & glyphProps[clusterMap[CharCount + 1]].IsDiacritic))
                        {
                            RunRes = clusterMap[CharCount + 1];
                        }
                    }
                }
            }
            float Width = 0f;
            float Top = 0f;
            float Bottom = 0f;
            SharpDX.DirectWrite.GlyphMetrics[] designGlyphMetrics = fontFace.GetDesignGlyphMetrics(glyphIndices, false);
            float Left = IsRTL ? 0f : (glyphOffsets[0].AdvanceOffset - ((designGlyphMetrics[0].LeftSideBearing * pointSize) / ((float)fontFace.Metrics.DesignUnitsPerEm)));
            float Right = IsRTL ? (glyphOffsets[0].AdvanceOffset - ((designGlyphMetrics[0].RightSideBearing * pointSize) / ((float)fontFace.Metrics.DesignUnitsPerEm))) : 0f;
            for (int i = 0; i <= designGlyphMetrics.Length - 1; i++)
            {
                Left = IsRTL ? Math.Max(Left, (glyphOffsets[i].AdvanceOffset + Width) - ((Math.Max(0, designGlyphMetrics[i].LeftSideBearing) * pointSize) / ((float)fontFace.Metrics.DesignUnitsPerEm))) : Math.Min(Left, (glyphOffsets[i].AdvanceOffset + Width) - ((designGlyphMetrics[i].LeftSideBearing * pointSize) / ((float)fontFace.Metrics.DesignUnitsPerEm)));
                if (!(glyphAdvances[i] == 0f))
                {
                    Width += (IsRTL ? ((float)(-1)) : ((float)1)) * ((designGlyphMetrics[i].AdvanceWidth * pointSize) / ((float)fontFace.Metrics.DesignUnitsPerEm));
                }
                Right = IsRTL ? Math.Min(Right, (glyphOffsets[i].AdvanceOffset + Width) - ((designGlyphMetrics[i].RightSideBearing * pointSize) / ((float)fontFace.Metrics.DesignUnitsPerEm))) : Math.Max(Right, (glyphOffsets[i].AdvanceOffset + Width) - ((Math.Min(0, designGlyphMetrics[i].RightSideBearing) * pointSize) / ((float)fontFace.Metrics.DesignUnitsPerEm)));
                Top = Math.Max(Top, glyphOffsets[i].AscenderOffset + (((designGlyphMetrics[i].VerticalOriginY - designGlyphMetrics[i].TopSideBearing) * pointSize) / ((float)fontFace.Metrics.DesignUnitsPerEm)));
                Bottom = Math.Min(Bottom, glyphOffsets[i].AscenderOffset + ((((designGlyphMetrics[i].VerticalOriginY - designGlyphMetrics[i].AdvanceHeight) + designGlyphMetrics[i].BottomSideBearing) * pointSize) / ((float)fontFace.Metrics.DesignUnitsPerEm)));
            }
            Pos = list.ToArray();
            Size Size = new Size(IsRTL ? (Left - Right) : (Right - Left), (Top - Bottom) + ((fontFace.Metrics.LineGap * pointSize) / ((float)fontFace.Metrics.DesignUnitsPerEm)));
            BaseLine = Top;
            analysisSource.Shadow.Dispose();
            analysisSink.Shadow.Dispose();
            analysisSource.Dispose();
            analysisSource._Factory = null;
            analysisSink.Dispose();
            fontFace.Dispose();
            font.Dispose();
            analyzer.Dispose();
            return Size;
        }
#endif
        public MyChildRenderItem(List<MyChildRenderBlockItem> NewItemRuns, bool NewIsArabic, bool NewIsRTL)
        {
            IsRTL = NewIsRTL;
            IsArabic = NewIsArabic; //must be set before ItemRuns
            ItemRuns = NewItemRuns;
            MaxWidth = CalculateWidth();
        }
        List<MyChildRenderBlockItem> _ItemRuns;
        public List<MyChildRenderBlockItem> ItemRuns
        {
            get { return _ItemRuns; }
            set
            {
                if (IsArabic)
                {
                    char[] Forms = XMLRender.ArabicData.GetPresentationForms;
                    //XMLRender.ArabicData.LigatureInfo[] ligs = XMLRender.ArabicData.GetLigatures(String.Join(String.Empty, System.Linq.Enumerable.Select(value, (Run) => Run.ItemText)), false, Forms);
                    //short[] chpos = GetWordDiacriticClusters(String.Join(String.Empty, System.Linq.Enumerable.Select(value, (Run) => Run.ItemText)), AppSettings.strSelectedFont, (float)AppSettings.dFontSize, IsArabic);
                    int pos = value[0].ItemText.Length;
                    int count = 1;
                    while (count < value.Count)
                    {
                        //if (value[count].ItemText.Length != 0 && pos != 0 && chpos[pos] == chpos[pos - 1])
                        //{
                        //    //if (!XMLRender.ArabicData.IsDiacritic(value[count].ItemText.First()) && value[count - 1].Clr != Windows.UI.Colors.Black) { value[count - 1].Clr = value[count].Clr; }
                        //    if (chpos[pos] == chpos[pos + value[count].ItemText.Length - 1])
                        //    {
                        //        pos -= value[count - 1].ItemText.Length;
                        //        value[count - 1].ItemText += value[count].ItemText; value.RemoveAt(count); count--;
                        //    }
                        //    else {
                        //        int subcount = pos;
                        //        while (subcount < pos + value[count].ItemText.Length && chpos[subcount] == chpos[subcount + 1]) { subcount++; }
                        //        value[count - 1].ItemText += value[count].ItemText.Substring(0, subcount - pos + 1); value[count].ItemText = value[count].ItemText.Substring(subcount - pos + 1);
                        //        pos += subcount - pos + 1;
                        //    }
                        //}
                        //for (int subcount = 0; subcount < ligs.Length; subcount++) {
                        //    //for indexes which are before, and after, find the maximum index in this group which is after
                        //    if (pos > ligs[subcount].Indexes[0] && pos <= ligs[subcount].Indexes[ligs[subcount].Indexes.Length - 1]) {
                        //        int ligcount;
                        //        for (ligcount = ligs[subcount].Indexes.Length - 1; ligcount >= 0; ligcount--) {
                        //            if (pos + value[count].ItemText.Length <= ligs[subcount].Indexes[ligcount]) {
                        //                if (ligs[subcount].Indexes[ligcount] - pos == value[count].ItemText.Length) {
                        //                    pos -= value[count - 1].ItemText.Length;
                        //                    if (!value[count].ItemText.All((ch) => XMLRender.ArabicData.IsDiacritic(ch))) { value[count - 1].Clr = value[count].Clr; }
                        //                    value[count - 1].ItemText += value[count].ItemText; value.RemoveAt(count); count--;
                        //                } else {
                        //                    value[count - 1].ItemText += value[count].ItemText.Substring(0, ligs[subcount].Indexes[ligcount] - pos); value[count].ItemText = value[count].ItemText.Substring(ligs[subcount].Indexes[ligcount] - pos);
                        //                    pos += ligs[subcount].Indexes[ligcount] - pos;
                        //                }
                        //                break;
                        //            }
                        //        }
                        //        if (ligcount != ligs[subcount].Indexes.Length) { break; }
                        //    }
                        //}
                        pos += value[count].ItemText.Length;
                        count++;
                        //int charcount = 0;
                        //while (charcount < value[count].ItemText.Length && XMLRender.ArabicData.IsDiacritic(value[count].ItemText[charcount])) { charcount++; }
                        //if (charcount != 0)
                        //{
                        //    if (charcount == value[count].ItemText.Length)
                        //    {
                        //        value[count - 1].ItemText += value[count].ItemText; value.RemoveAt(count);
                        //    }
                        //    else { value[count - 1].ItemText += value[count].ItemText.Substring(0, charcount); value[count].ItemText = value[count].ItemText.Substring(charcount); count++; }
                        //}
                        //else { count++; }
                    }
                }
                _ItemRuns = value;
            }
        }
        //public FlowDirection Direction { get; set; }
        public bool IsArabic { get; set; }
        public bool IsRTL { get; set; }
    }
    public class BackgroundSelectedConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, string language)
        {
#if __ANDROID__
            return ((bool)value ? new System.Drawing.Color.Beige : Color.White);
#else
            return ((bool)value ? new Color(245, 245, 220) : Color.White);
#endif
        }

        public object ConvertBack(object value, Type targetType, object parameter, string language)
        {
            throw new NotImplementedException();
        }
    }
    public class ArabicFlowConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, string language)
        {
            return (bool)value ? FlowDirection.RightToLeft : FlowDirection.LeftToRight;
        }

        public object ConvertBack(object value, Type targetType, object parameter, string language)
        {
            throw new NotImplementedException();
        }
    }
    public class MyChildRenderBlockItem
    {
        public int Clr;
        public string ItemText { get; set; }
    }
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
}
