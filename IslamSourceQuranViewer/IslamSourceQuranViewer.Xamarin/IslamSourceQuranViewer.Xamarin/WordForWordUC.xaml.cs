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
		public WordForWordUC ()
		{
			InitializeComponent ();
		}
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
        public double FontSize
        {
            get
            {
                return AppSettings.dFontSize;
            }
            set
            {
                AppSettings.dFontSize = value;
                //if (_DWArabicFormat != null) _DWArabicFormat.Dispose();
                //_DWArabicFormat = null;
                PropertyChanged(this, new PropertyChangedEventArgs("FontSize"));
            }
        }
        public string FontFamily
        {
            get
            {
                return AppSettings.strSelectedFont;
            }
        }
        public double OtherFontSize
        {
            get
            {
                return AppSettings.dOtherFontSize;
            }
            set
            {
                AppSettings.dOtherFontSize = value;
                //if (_DWNormalFormat != null) _DWNormalFormat.Dispose();
                //_DWNormalFormat = null;
                PropertyChanged(this, new PropertyChangedEventArgs("OtherFontSize"));
            }
        }
        public string OtherFontFamily
        {
            get
            {
                return AppSettings.strOtherSelectedFont;
            }
        }
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
        public MyRenderModel(IEnumerable<MyRenderItem> NewRenderItems)
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
        public IEnumerable<MyRenderItem> RenderItems { get { return _RenderItems; } set { _RenderItems = value.ToList(); if (PropertyChanged != null) PropertyChanged(this, new PropertyChangedEventArgs("RenderItems")); } }
        #region Implementation of INotifyPropertyChanged

        public event PropertyChangedEventHandler PropertyChanged;

        #endregion
    }
    public class MyRenderItem : INotifyPropertyChanged
    {
        public MyRenderItem(XMLRender.RenderArray.RenderItem RendItem)
        {
            Items = System.Linq.Enumerable.Select(RendItem.TextItems.GroupBy((MainItems) => (MainItems.DisplayClass == XMLRender.RenderArray.RenderDisplayClass.eArabic || MainItems.DisplayClass == XMLRender.RenderArray.RenderDisplayClass.eLTR || MainItems.DisplayClass == XMLRender.RenderArray.RenderDisplayClass.eRTL || MainItems.DisplayClass == XMLRender.RenderArray.RenderDisplayClass.eTransliteration) ? (object)MainItems.DisplayClass : (object)MainItems), (Arr) => (Arr.First().Text.GetType() == typeof(List<XMLRender.RenderArray.RenderItem>)) ? (object)new MyRenderModel(System.Linq.Enumerable.Select((List<XMLRender.RenderArray.RenderItem>)Arr.First().Text, (ArrRend) => new MyRenderItem((XMLRender.RenderArray.RenderItem)ArrRend))) : (Arr.First().Text.GetType() == typeof(string) ? (object)new MyChildRenderItem(System.Linq.Enumerable.Select(Arr, (ArrItem) => new MyChildRenderBlockItem() { ItemText = (string)ArrItem.Text, Clr = Windows.UI.Color.FromArgb(0xFF, XMLRender.Utility.ColorR(ArrItem.Clr), XMLRender.Utility.ColorG(ArrItem.Clr), XMLRender.Utility.ColorB(ArrItem.Clr)) }).ToList(), Arr.First().DisplayClass == XMLRender.RenderArray.RenderDisplayClass.eArabic, (Arr.First().DisplayClass == XMLRender.RenderArray.RenderDisplayClass.eArabic || Arr.First().DisplayClass == XMLRender.RenderArray.RenderDisplayClass.eRTL) ? FlowDirection.RightToLeft : FlowDirection.LeftToRight) : null)).Where(Arr => Arr != null);
            MaxWidth = CalculateWidth();
        }
        public double MaxWidth { get; set; }
        private double CalculateWidth()
        {
            if (Items.Count() == 0) return 0.0;
            return Items.Select((Item) => Item.GetType() == typeof(MyChildRenderItem) ? ((MyChildRenderItem)Item).MaxWidth : ((MyRenderModel)Item).MaxWidth).Max();
        }
        private List<object> _Items;
        public IEnumerable<object> Items { get { return _Items; } set { _Items = value.ToList(); if (PropertyChanged != null) PropertyChanged(this, new PropertyChangedEventArgs("Items")); } }
        #region Implementation of INotifyPropertyChanged

        public event PropertyChangedEventHandler PropertyChanged;

        #endregion
    }
    public class VirtualizingWrapPanelAdapter : INotifyPropertyChanged
    {
        private List<MyRenderModel> _RenderModels;
        public List<MyRenderModel> RenderModels
        {
            get
            {
                return _RenderModels;
            }
            set
            {
                _RenderModels = value;
                PropertyChanged(this, new PropertyChangedEventArgs("RenderModels"));
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
                double itemWidth = Math.Min(Item.MaxWidth, maxWidth); if (width + itemWidth > maxWidth) { width = itemWidth; groupIndex++; } else { width += itemWidth; }
                GroupIndexes.Add(new int[] { groupIndex, GroupIndexes.Count }); return false;
            });
            return GroupIndexes.GroupBy((Item) => Item[0], (Item) => value.ElementAt(Item[1])).Select((Item) => new MyRenderModel(Item)).ToList();
        }
        #region Implementation of INotifyPropertyChanged

        public event PropertyChangedEventHandler PropertyChanged;

        #endregion
    }
    //public static class FormattedTextBehavior
    //{
    //    #region FormattedText Attached dependency property

    //    public static List<MyChildRenderBlockItem> GetFormattedText(DependencyObject obj)
    //    {
    //        return (List<MyChildRenderBlockItem>)obj.GetValue(FormattedTextProperty);
    //    }

    //    public static void SetFormattedText(DependencyObject obj, List<MyChildRenderBlockItem> value)
    //    {
    //        obj.SetValue(FormattedTextProperty, value);
    //    }

    //    public static readonly DependencyProperty FormattedTextProperty =
    //        DependencyProperty.RegisterAttached("FormattedText",
    //        typeof(List<MyChildRenderBlockItem>),
    //        typeof(FormattedTextBehavior),
    //        new PropertyMetadata(null, FormattedTextChanged));

    //    private static void FormattedTextChanged(DependencyObject sender, DependencyPropertyChangedEventArgs e)
    //    {
    //        List<MyChildRenderBlockItem> value = e.NewValue as List<MyChildRenderBlockItem>;

    //        TextBlock textBlock = sender as TextBlock;

    //        if (textBlock != null)
    //        {
    //            textBlock.Inlines.Clear();
    //            value.FirstOrDefault((it) => { textBlock.Inlines.Add(new Windows.UI.Xaml.Documents.Run() { Text = it.ItemText, Foreground = it.Color }); return false; });
    //        }
    //    }

    //    #endregion
    //}
    public class MyChildRenderItem : INotifyPropertyChanged
    {
        private double _MaxWidth;
        public double MaxWidth { get { return _MaxWidth; } set { _MaxWidth = value; if (PropertyChanged != null) PropertyChanged(this, new PropertyChangedEventArgs("MaxWidth")); } }
        //public static double CalculateWidth(string text, bool IsArabic, float maxWidth, float maxHeight)
        //{
        //    //SharpDX.DirectWrite.TextLayout layout = new SharpDX.DirectWrite.TextLayout(MyUIChanger.DWFactory, text, IsArabic ? MyUIChanger.DWArabicFormat : MyUIChanger.DWNormalFormat, maxWidth, maxHeight);
        //    //double width = layout.Metrics.WidthIncludingTrailingWhitespace + layout.Metrics.Left;
        //    //layout.Dispose();
        //    //return width;
        //}
        //public double CalculateWidth()
        //{
        //    //5 margin on both sides
        //    return 5 + 5 + 1 + 1 + CalculateWidth(string.Join(String.Empty, ItemRuns.Select((Item) => Item.ItemText)), IsArabic, (float)float.MaxValue, float.MaxValue);
        //}
        #region Implementation of INotifyPropertyChanged

        public event PropertyChangedEventHandler PropertyChanged;

        #endregion
 
        //public static short[] GetWordDiacriticClusters(string Str, string useFont, float fontSize, bool IsRTL)
        //{
        //}
        //public static Size GetWordDiacriticPositionsDWrite(string Str, string useFont, float fontSize, char[] Forms, bool IsRTL, ref float BaseLine, ref CharPosInfo[] Pos)
        //{
        //}

        //public MyChildRenderItem(List<MyChildRenderBlockItem> NewItemRuns, bool NewIsArabic, FlowDirection NewDirection)
        //{
        //    IsArabic = NewIsArabic; //must be set before ItemRuns
        //    Direction = NewDirection;
        //    ItemRuns = NewItemRuns;
        //    MaxWidth = CalculateWidth();
        //}
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
    }
    public class MyChildRenderBlockItem
    {
        public Windows.UI.Color Clr;
        //public SolidColorBrush Color { get { return new SolidColorBrush(Clr); } }
        public string ItemText { get; set; }
    }
    //public class MyDataTemplateSelector : DataTemplateSelector
    //{
    //    public DataTemplate MyTemplate { get; set; }

    //    protected override DataTemplate SelectTemplateCore(object item, DependencyObject container)
    //    {
    //        return MyTemplate;
    //    }
    //}
    //public class NormalWordTemplateSelector : DataTemplateSelector
    //{
    //    public DataTemplate WordTemplate { get; set; }
    //    public DataTemplate ArabicTemplate { get; set; }
    //    public DataTemplate NormalTemplate { get; set; }

    //    protected override DataTemplate SelectTemplateCore(object item, DependencyObject container)
    //    {
    //        if (item.GetType() == typeof(MyRenderModel)) { return WordTemplate; }
    //        return ((MyChildRenderItem)item).IsArabic ? ArabicTemplate : NormalTemplate;
    //    }

    //}
}
