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
            this.DataContext = this;
            UIChanger = new MyUIChanger();
            this.InitializeComponent();
#if WINDOWS_APP
            AppBarButton BackButton = new AppBarButton() { Icon = new SymbolIcon(Symbol.Back), Label = new Windows.ApplicationModel.Resources.ResourceLoader().GetString("Back/Label") };
            BackButton.Click += Back_Click;
            (this.BottomAppBar as CommandBar).PrimaryCommands.Add(BackButton);
#endif
#if STORETOOLKIT
            AppBarButton RenderButton = new AppBarButton() { Icon = new SymbolIcon(Symbol.Camera), Label = new Windows.ApplicationModel.Resources.ResourceLoader().GetString("Render/Label") };
            RenderButton.Click += RenderPngs_Click;
            (this.BottomAppBar as CommandBar).PrimaryCommands.Add(RenderButton);
#endif
        }

        protected override void OnNavigatedTo(NavigationEventArgs e)
        {
            dynamic t = e.Parameter;
            Division = t.Division;
            Selection = t.Selection;
            this.ViewModel = new MyRenderModel(IslamMetadata.TanzilReader.GetRenderedQuranText(AppSettings.bShowTransliteration ? XMLRender.ArabicData.TranslitScheme.RuleBased : XMLRender.ArabicData.TranslitScheme.None, String.Empty, String.Empty, Division.ToString(), Selection.ToString(), String.Empty, AppSettings.bShowW4W ? "0" : "4", AppSettings.bUseColoring ? "0" : "1").Items);
            MainControl.ItemsSource = this.ViewModel.RenderItems;
        }
        public MyUIChanger UIChanger { get; set; }
        public MyRenderModel ViewModel { get; set; }
        private async void RenderPngs_Click(object sender, RoutedEventArgs e)
        {
            string origLang = Windows.Globalization.ApplicationLanguages.PrimaryLanguageOverride;
            string[] LangList = "en,hu,zh-Hans,zh-Hant".Split(',');
            //string[] LangList = "ar,ar-sa,ar-ae,ar-bh,ar-dz,ar-eg,ar-iq,ar-jo,ar-kw,ar-lb,ar-ly,ar-ma,ar-om,ar-qa,ar-sy,ar-tn,ar-ye,af,af-za,sq,sq-al,am,am-et,hy,hy-am,as,as-in,az,az-arab,az-arab-az,az-cyrl,az-cyrl-az,az-latn,az-latn-az,eu,eu-es,be,be-by,bn,bn-bd,bn-in,bs,bs-cyrl,bs-cyrl-ba,bs-latn,bs-latn-ba,bg,bg-bg,ca,ca-es,ca-es-valencia,chr-cher,chr-cher-us,chr-latn,zh,zh-Hans,zh-cn,zh-hans-cn,zh-sg,zh-hans-sg,zh-Hant,zh-hk,zh-mo,zh-tw,zh-hant-hk,zh-hant-mo,zh-hant-tw,hr,hr-hr,hr-ba,cs,cs-cz,da,da-dk,prs,prs-af,prs-arab,nl,nl-nl,nl-be,en,en-au,en-ca,en-gb,en-ie,en-in,en-nz,en-sg,en-us,en-za,en-bz,en-hk,en-id,en-jm,en-kz,en-mt,en-my,en-ph,en-pk,en-tt,en-vn,en-zw,en-053,en-021,en-029,en-011,en-018,en-014,et,et-ee,fil,fil-latn,fil-ph,fi,fi-fi,fr,fr-be,fr-ca,fr-ch,fr-fr,fr-lu,fr-015,fr-cd,fr-ci,fr-cm,fr-ht,fr-ma,fr-mc,fr-ml,fr-re,frc-latn,frp-latn,fr-155,fr-029,fr-021,fr-011,gl,gl-es,ka,ka-ge,de,de-at,de-ch,de-de,de-lu,de-li,el,el-gr,gu,gu-in,ha,ha-latn,ha-latn-ng,he,he-il,hi,hi-in,hu,hu-hu,is,is-is,ig-latn,ig-ng,id,id-id,iu-cans,iu-latn,iu-latn-ca,ga,ga-ie,xh,xh-za,zu,zu-za,it,it-it,it-ch,ja,ja-jp,kn,kn-in,kk,kk-kz,km,km-kh,quc-latn,qut-gt,qut-latn,rw,rw-rw,sw,sw-ke,kok,kok-in,ko,ko-kr,ku-arab,ku-arab-iq,ky-kg,ky-cyrl,lo,lo-la,lv,lv-lv,lt,lt-lt,lb,lb-lu,mk,mk-mk,ms,ms-bn,ms-my,ml,ml-in,mt,mt-mt,mi,mi-latn,mi-nz,mr,mr-in,mn-cyrl,mn-mong,mn-mn,mn-phag,ne,ne-np,nb,nb-no,nn,nn-no,no,no-no,,or,or-in,fa,fa-ir,pl,pl-pl,pt-br,pt,pt-pt,pa,pa-arab,pa-arab-pk,pa-deva,pa-in,quz,quz-bo,quz-ec,quz-pe,ro,ro-ro,ru,ru-ru,gd-gb,gd-latn,sr-Latn,sr-latn-cs,sr,sr-latn-ba,sr-latn-me,sr-latn-rs,sr-cyrl,sr-cyrl-ba,sr-cyrl-cs,sr-cyrl-me,sr-cyrl-rs,nso,nso-za,tn,tn-bw,tn-za,sd-arab,sd-arab-pk,sd-deva,si,si-lk,sk,sk-sk,sl,sl-si,es,es-cl,es-co,es-es,es-mx,es-ar,es-bo,es-cr,es-do,es-ec,es-gt,es-hn,es-ni,es-pa,es-pe,es-pr,es-py,es-sv,es-us,es-uy,es-ve,es-019,es-419,sv,sv-se,sv-fi,tg-arab,tg-cyrl,tg-cyrl-tj,tg-latn,ta,ta-in,tt-arab,tt-cyrl,tt-latn,tt-ru,te,te-in,th,th-th,ti,ti-et,tr,tr-tr,tk-cyrl,tk-latn,tk-tm,tk-latn-tr,tk-cyrl-tr,uk,uk-ua,ur,ur-pk,ug-arab,ug-cn,ug-cyrl,ug-latn,uz,uz-cyrl,uz-latn,uz-latn-uz,vi,vi-vn,cy,cy-gb,wo,wo-sn,yo-latn,yo-ng".Split(',');
            for (int count = 0; count < LangList.Length; count++) {
                Windows.Globalization.ApplicationLanguages.PrimaryLanguageOverride = LangList[count];
                await Windows.Storage.ApplicationData.Current.LocalFolder.CreateFolderAsync(LangList[count], Windows.Storage.CreationCollisionOption.OpenIfExists);
                //re-render the UI
                //this.Frame.Navigate(this.GetType(), new { Division = Division, Selection = Selection});
                //this.Frame.BackStack.Remove(this.Frame.BackStack.Last());
                this.ViewModel = new MyRenderModel(IslamMetadata.TanzilReader.GetRenderedQuranText(AppSettings.bShowTransliteration ? XMLRender.ArabicData.TranslitScheme.RuleBased : XMLRender.ArabicData.TranslitScheme.None, String.Empty, String.Empty, Division.ToString(), Selection.ToString(), String.Empty, AppSettings.bShowW4W ? "0" : "4", AppSettings.bUseColoring ? "0" : "1").Items);
                MainControl.ItemsSource = this.ViewModel.RenderItems;
                UpdateLayout();
                await Dispatcher.RunAsync(Windows.UI.Core.CoreDispatcherPriority.Low, () => { });
                await WindowsRTSettings.SavePathImageAsFile(1366, 768, LangList[count] + "\\appstorescreenshot-wide", this, true);
                await WindowsRTSettings.SavePathImageAsFile(768, 1366, LangList[count] + "\\appstorescreenshot-tall", this, true);
                await WindowsRTSettings.SavePathImageAsFile(1280, 768, LangList[count] + "\\appstorephonescreenshot-wide", this, true);
                await WindowsRTSettings.SavePathImageAsFile(768, 1280, LangList[count] + "\\appstorephonescreenshot-tall", this, true);
                await WindowsRTSettings.SavePathImageAsFile(1280, 720, LangList[count] + "\\appstorephonescreenshot1280x720-wide", this, true);
                await WindowsRTSettings.SavePathImageAsFile(720, 1280, LangList[count] + "\\appstorephonescreenshot720x1280-tall", this, true);
                await WindowsRTSettings.SavePathImageAsFile(800, 480, LangList[count] + "\\appstorephonescreenshot800x480-wide", this, true);
                await WindowsRTSettings.SavePathImageAsFile(480, 800, LangList[count] + "\\appstorephonescreenshot480x800-tall", this, true);
            }
            Windows.Globalization.ApplicationLanguages.PrimaryLanguageOverride = origLang;
            this.ViewModel = new MyRenderModel(IslamMetadata.TanzilReader.GetRenderedQuranText(AppSettings.bShowTransliteration ? XMLRender.ArabicData.TranslitScheme.RuleBased : XMLRender.ArabicData.TranslitScheme.None, String.Empty, String.Empty, Division.ToString(), Selection.ToString(), String.Empty, AppSettings.bShowW4W ? "0" : "4", AppSettings.bUseColoring ? "0" : "1").Items);
            MainControl.ItemsSource = this.ViewModel.RenderItems;
            UpdateLayout();
            await Dispatcher.RunAsync(Windows.UI.Core.CoreDispatcherPriority.Low, () => { });
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
            UIChanger.FontSize += 1.0;
            UIChanger.OtherFontSize += 1.0;
        }

        private void ZoomOut_Click(object sender, RoutedEventArgs e)
        {
            UIChanger.FontSize -= 1.0;
            UIChanger.OtherFontSize -= 1.0;
        }

        private void Default_Click(object sender, RoutedEventArgs e)
        {
            UIChanger.FontSize = 30.0;
            UIChanger.OtherFontSize = 20.0;
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
            Items = System.Linq.Enumerable.Select(RendItem.TextItems.GroupBy((MainItems) => (MainItems.DisplayClass == XMLRender.RenderArray.RenderDisplayClass.eArabic || MainItems.DisplayClass == XMLRender.RenderArray.RenderDisplayClass.eLTR || MainItems.DisplayClass == XMLRender.RenderArray.RenderDisplayClass.eRTL || MainItems.DisplayClass == XMLRender.RenderArray.RenderDisplayClass.eTransliteration) ? (object)MainItems.DisplayClass : (object)MainItems), (Arr) => (Arr.First().Text.GetType() == typeof(List<XMLRender.RenderArray.RenderItem>)) ? (object)new MyRenderModel((List<XMLRender.RenderArray.RenderItem>)Arr.First().Text) : (Arr.First().Text.GetType() == typeof(string) ? (object)new MyChildRenderItem { IsArabic = Arr.First().DisplayClass == XMLRender.RenderArray.RenderDisplayClass.eArabic, Direction = (Arr.First().DisplayClass == XMLRender.RenderArray.RenderDisplayClass.eArabic || Arr.First().DisplayClass == XMLRender.RenderArray.RenderDisplayClass.eRTL) ? FlowDirection.RightToLeft : FlowDirection.LeftToRight, ItemRuns = System.Linq.Enumerable.Select(Arr, (ArrItem) => new MyChildRenderBlockItem() { ItemText = (string)ArrItem.Text, Color = new SolidColorBrush(Windows.UI.Color.FromArgb(0xFF, XMLRender.Utility.ColorR(ArrItem.Clr), XMLRender.Utility.ColorG(ArrItem.Clr), XMLRender.Utility.ColorB(ArrItem.Clr))) }).ToList() } : null)).Where(Arr => Arr != null);
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
            SharpDX.DirectWrite.Factory factory = new SharpDX.DirectWrite.Factory();
            SharpDX.DirectWrite.TextAnalyzer analyzer = new SharpDX.DirectWrite.TextAnalyzer(factory);
            LOGFONT lf = new LOGFONT();
            lf.lfFaceName = useFont;
            float pointSize = fontSize * Windows.Graphics.Display.DisplayInformation.GetForCurrentView().RawDpiY / 72.0f;
            lf.lfHeight = (int)fontSize;
            lf.lfQuality = 5; //clear type
            SharpDX.DirectWrite.Font font = factory.GdiInterop.FromLogFont(lf);
            SharpDX.DirectWrite.FontFace fontFace = new SharpDX.DirectWrite.FontFace(font);
            SharpDX.DirectWrite.ScriptAnalysis scriptAnalysis = new SharpDX.DirectWrite.ScriptAnalysis();
            TextSink analysisSink = new TextSink();
            TextSource analysisSource = new TextSource(Str, factory);
            analyzer.AnalyzeScript(analysisSource, 0, Str.Length, analysisSink);
            scriptAnalysis = analysisSink._scriptAnalysis;
            int maxGlyphCount = ((Str.Length * 3) / 2) + 0x10;
            short[] clusterMap = new short[(Str.Length - 1) + 1];
            SharpDX.DirectWrite.ShapingTextProperties[] textProps = new SharpDX.DirectWrite.ShapingTextProperties[(Str.Length - 1) + 1];
            short[] glyphIndices = new short[(maxGlyphCount - 1) + 1];
            SharpDX.DirectWrite.ShapingGlyphProperties[] glyphProps = new SharpDX.DirectWrite.ShapingGlyphProperties[(maxGlyphCount - 1) + 1];
            int actualGlyphCount = 0;
            SharpDX.DirectWrite.FontFeature[] featureArray = new SharpDX.DirectWrite.FontFeature[] { new SharpDX.DirectWrite.FontFeature(SharpDX.DirectWrite.FontFeatureTag.GlyphCompositionDecomposition, 1), new SharpDX.DirectWrite.FontFeature(SharpDX.DirectWrite.FontFeatureTag.DiscretionaryLigatures, 1), new SharpDX.DirectWrite.FontFeature(SharpDX.DirectWrite.FontFeatureTag.StandardLigatures, 1), new SharpDX.DirectWrite.FontFeature(SharpDX.DirectWrite.FontFeatureTag.ContextualAlternates, 1), new SharpDX.DirectWrite.FontFeature(SharpDX.DirectWrite.FontFeatureTag.StylisticSet1, 1) };
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
            return clusterMap;
        }
        public static Size GetWordDiacriticPositionsDWrite(string Str, string useFont, float fontSize, char[] Forms, bool IsRTL, ref float BaseLine, ref CharPosInfo[] Pos)
        {
            if (Str == string.Empty)
            {
                return new Size(0f, 0f);
            }
            SharpDX.DirectWrite.Factory factory = new SharpDX.DirectWrite.Factory();
            SharpDX.DirectWrite.TextAnalyzer analyzer = new SharpDX.DirectWrite.TextAnalyzer(factory);
            LOGFONT lf = new LOGFONT();
            lf.lfFaceName = useFont;
            float pointSize = fontSize * Windows.Graphics.Display.DisplayInformation.GetForCurrentView().RawDpiY / 72.0f;
            lf.lfHeight = (int)fontSize;
            lf.lfQuality = 5; //clear type
            SharpDX.DirectWrite.Font font = factory.GdiInterop.FromLogFont(lf);
            SharpDX.DirectWrite.FontFace fontFace = new SharpDX.DirectWrite.FontFace(font);
            SharpDX.DirectWrite.ScriptAnalysis scriptAnalysis = new SharpDX.DirectWrite.ScriptAnalysis();
            TextSink analysisSink = new TextSink();
            TextSource analysisSource = new TextSource(Str, factory);
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
            factory.Dispose();
            return Size;
        }


        List<MyChildRenderBlockItem> _ItemRuns;
        public List<MyChildRenderBlockItem> ItemRuns { get { return _ItemRuns; } set {
                if (IsArabic)
                {
                    char[] Forms = XMLRender.ArabicData.GetPresentationForms;
                    //XMLRender.ArabicData.LigatureInfo[] ligs = XMLRender.ArabicData.GetLigatures(String.Join(String.Empty, System.Linq.Enumerable.Select(value, (Run) => Run.ItemText)), false, Forms);
                    short[] chpos = GetWordDiacriticClusters(String.Join(String.Empty, System.Linq.Enumerable.Select(value, (Run) => Run.ItemText)), AppSettings.strSelectedFont, (float)AppSettings.dFontSize, IsArabic);
                    int pos = value[0].ItemText.Length;
                    int count = 1;
                    while (count < value.Count)
                    {
                        if (value[count].ItemText.Length != 0 && pos != 0 && chpos[pos] == chpos[pos - 1])
                        {
                            if (!value[count].ItemText.All((ch) => XMLRender.ArabicData.IsDiacritic(ch)) && value[count - 1].Color.Color != Windows.UI.Colors.Black) { value[count - 1].Color = value[count].Color; }
                            if (chpos[pos] == chpos[pos + value[count].ItemText.Length - 1])
                            {
                                pos -= value[count - 1].ItemText.Length;
                                value[count - 1].ItemText += value[count].ItemText; value.RemoveAt(count); count--;
                            }
                            else {
                                int subcount = pos;
                                while (subcount < pos + value[count].ItemText.Length && chpos[subcount] == chpos[subcount + 1]) { subcount++; }
                                value[count - 1].ItemText += value[count].ItemText.Substring(0, subcount - pos + 1); value[count].ItemText = value[count].ItemText.Substring(subcount - pos + 1);
                                pos += subcount - pos + 1;
                            }
                        }
                        //for (int subcount = 0; subcount < ligs.Length; subcount++) {
                        //    //for indexes which are before, and after, find the maximum index in this group which is after
                        //    if (pos > ligs[subcount].Indexes[0] && pos <= ligs[subcount].Indexes[ligs[subcount].Indexes.Length - 1]) {
                        //        int ligcount;
                        //        for (ligcount = ligs[subcount].Indexes.Length - 1; ligcount >= 0; ligcount--) {
                        //            if (pos + value[count].ItemText.Length <= ligs[subcount].Indexes[ligcount]) {
                        //                if (ligs[subcount].Indexes[ligcount] - pos == value[count].ItemText.Length) {
                        //                    pos -= value[count - 1].ItemText.Length;
                        //                    if (!value[count].ItemText.All((ch) => XMLRender.ArabicData.IsDiacritic(ch))) { value[count - 1].Color = value[count].Color; }
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
        public FlowDirection Direction { get; set; }
        public bool IsArabic { get; set; }
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
        public DataTemplate ArabicTemplate { get; set; }
        public DataTemplate NormalTemplate { get; set; }

        protected override DataTemplate SelectTemplateCore(object item, DependencyObject container)
        {
            if (item.GetType() == typeof(MyRenderModel)) { return WordTemplate; }
            return ((MyChildRenderItem)item).IsArabic ? ArabicTemplate : NormalTemplate;
        }

    }
}
