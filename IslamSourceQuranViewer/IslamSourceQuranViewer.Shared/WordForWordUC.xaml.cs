using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
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
using Windows.UI.Xaml.Media.Imaging;
using Windows.UI.Xaml.Navigation;
using WinRTXamlToolkit.Controls;

namespace IslamSourceQuranViewer
{
    public sealed partial class WordForWordUC : Page
    {
        public WordForWordUC()
        {
            this.DataContext = this;
            this.ViewModel = new VirtualizingWrapPanelAdapter();
            UIChanger = new MyUIChanger();
            this.InitializeComponent();

            MainGrid.SizeChanged += OnSizeChanged;
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
            gestRec = new Windows.UI.Input.GestureRecognizer();
            gestRec.GestureSettings = Windows.UI.Input.GestureSettings.HoldWithMouse | Windows.UI.Input.GestureSettings.Hold | Windows.UI.Input.GestureSettings.RightTap | Windows.UI.Input.GestureSettings.Tap;
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
            gestRec.CompleteGesture();
        }
        void OnTapped(Windows.UI.Input.GestureRecognizer sender, Windows.UI.Input.TappedEventArgs args)
        {
            if (_holdObj != null) {
                return;
                object obj = ((_holdObj as StackPanel).DataContext as MyRenderItem).Items.FirstOrDefault((Item) => Item.GetType() == typeof(MyChildRenderStopContinue));
                if (obj != null) {
                    //ContentPresenter, ItemsControl and ItemsPresenter in middle layer between higher and lower StackPanel
                    (obj as MyChildRenderStopContinue).IsStop = !(obj as MyChildRenderStopContinue).IsStop;
                    VirtualizingWrapPanelAdapter panel = ((VisualTreeHelper.GetParent(VisualTreeHelper.GetParent(VisualTreeHelper.GetParent(VisualTreeHelper.GetParent(VisualTreeHelper.GetParent(VisualTreeHelper.GetParent(VisualTreeHelper.GetParent(VisualTreeHelper.GetParent(_holdObj as StackPanel)))))))) as VirtualizingStackPanel).DataContext as VirtualizingWrapPanelAdapter);
                    MyRenderModel model = ((VisualTreeHelper.GetParent(VisualTreeHelper.GetParent(VisualTreeHelper.GetParent(VisualTreeHelper.GetParent(_holdObj as StackPanel)))) as StackPanel).DataContext as MyRenderModel);
                    int mmodel = panel.RenderModels.IndexOf(model);
                    int index = model.RenderItems.IndexOf((_holdObj as StackPanel).DataContext as MyRenderItem);
                    List<int> stops = new List<int>();
                    if ((mmodel >= 1 || index >= 2) && (panel.RenderModels[mmodel - (index <= 1 ? 1 : 0)].RenderItems[index <= 1 ? panel.RenderModels[mmodel - 1].RenderItems.Count - 2 + index : index - 2].Items.FirstOrDefault((Item) => Item.GetType() == typeof(MyChildRenderStopContinue) && (Item as MyChildRenderStopContinue).IsStop) as MyChildRenderStopContinue) != null) stops.Add(-1);
                    string text = (panel.RenderModels[mmodel - (index == 0 ? 1 : 0)].RenderItems[index == 0 ? panel.RenderModels[mmodel - 1].RenderItems.Count - 1 : index - 1].Items.FirstOrDefault((Item) => Item.GetType() == typeof(MyChildRenderItem) && (Item as MyChildRenderItem).IsArabic) as MyChildRenderItem).GetText;
                    MyChildRenderItem MainItem = (model.RenderItems[index].Items.FirstOrDefault((Item) => Item.GetType() == typeof(MyChildRenderItem) && (Item as MyChildRenderItem).IsArabic) as MyChildRenderItem);
                    if ((obj as MyChildRenderStopContinue).IsStop) stops.Add(text.Length + (MainItem != null ? 1 : 0));
                    text += (MainItem != null ? (" " + MainItem.GetText.Trim()) : string.Empty);
                    if (model.RenderItems[index].Chapter == panel.RenderModels[mmodel + (index + 1 == model.RenderItems.Count ? 1 : 0)].RenderItems[index + 1 == model.RenderItems.Count ? 0 : index + 1].Chapter && model.RenderItems[index].Verse == panel.RenderModels[mmodel + (index + 1 == model.RenderItems.Count ? 1 : 0)].RenderItems[index + 1 == model.RenderItems.Count ? 0 : index + 1].Verse && model.RenderItems[index].Word + 1 == panel.RenderModels[mmodel + (index + 1 == model.RenderItems.Count ? 1 : 0)].RenderItems[index + 1 == model.RenderItems.Count ? 0 : index + 1].Word ||
                        model.RenderItems[index].Chapter == panel.RenderModels[mmodel + (index + 1 == model.RenderItems.Count ? 1 : 0)].RenderItems[index + 1 == model.RenderItems.Count ? 0 : index + 1].Chapter && model.RenderItems[index].Verse + 1 == panel.RenderModels[mmodel + (index + 1 == model.RenderItems.Count ? 1 : 0)].RenderItems[index + 1 == model.RenderItems.Count ? 0 : index + 1].Verse && model.RenderItems[index].Word == AppSettings.TR.GetWordCount(model.RenderItems[index].Chapter, model.RenderItems[index].Verse) && panel.RenderModels[mmodel + (index + 1 == model.RenderItems.Count ? 1 : 0)].RenderItems[index + 1 == model.RenderItems.Count ? 0 : index + 1].Word == 1 ||
                        model.RenderItems[index].Chapter + 1 == panel.RenderModels[mmodel + (index + 1 == model.RenderItems.Count ? 1 : 0)].RenderItems[index + 1 == model.RenderItems.Count ? 0 : index + 1].Chapter && model.RenderItems[index].Verse == AppSettings.TR.GetVerseCount(model.RenderItems[index].Chapter) && panel.RenderModels[mmodel + (index + 1 == model.RenderItems.Count ? 1 : 0)].RenderItems[index + 1 == model.RenderItems.Count ? 0 : index + 1].Verse == 1 && model.RenderItems[index].Word == AppSettings.TR.GetWordCount(model.RenderItems[index].Chapter, model.RenderItems[index].Verse) && panel.RenderModels[mmodel + (index + 1 == model.RenderItems.Count ? 1 : 0)].RenderItems[index + 1 == model.RenderItems.Count ? 0 : index + 1].Word == 1) {
                        text += " " + (panel.RenderModels[mmodel + (index + 1 == model.RenderItems.Count ? 1 : 0)].RenderItems[index + 1 == model.RenderItems.Count ? 0 : index + 1].Items.FirstOrDefault((Item) => Item.GetType() == typeof(MyChildRenderItem) && (Item as MyChildRenderItem).IsArabic) as MyChildRenderItem).GetText;
                    }
                    if ((panel.RenderModels[mmodel + (index + 2 >= model.RenderItems.Count ? 1 : 0)].RenderItems[index + 2 == model.RenderItems.Count ? 0 : (index + 1 == model.RenderItems.Count ? 1 : index + 2)].Items.FirstOrDefault((Item) => Item.GetType() == typeof(MyChildRenderStopContinue) && (Item as MyChildRenderStopContinue).IsStop) as MyChildRenderStopContinue) != null) stops.Add(text.Length);
                    if (AppSettings.bUseColoring)
                    {
                        XMLRender.RenderArray.RenderText[][] strs = AppSettings.Arb.ApplyColorRules(text, true, IslamMetadata.Arabic.FilterMetadataStops(text, (obj as MyChildRenderStopContinue).MetaRules, stops.ToArray()));
                        (panel.RenderModels[mmodel - (index == 0 ? 1 : 0)].RenderItems[index == 0 ? panel.RenderModels[mmodel - 1].RenderItems.Count - 1 : index - 1].Items.FirstOrDefault((Item) => Item.GetType() == typeof(MyChildRenderItem) && (Item as MyChildRenderItem).IsArabic) as MyChildRenderItem).ItemRuns = strs[0].Select((Item) => new MyChildRenderBlockItem() { ItemText = (string)Item.Text, Clr = Item.Clr }).ToList();
                        if (MainItem != null) MainItem.ItemRuns = strs[1].Select((Item) => new MyChildRenderBlockItem() { ItemText = (string)Item.Text, Clr = Item.Clr }).ToList();
                        (panel.RenderModels[mmodel + (index + 1 == model.RenderItems.Count ? 1 : 0)].RenderItems[index + 1 == model.RenderItems.Count ? 0 : index + 1].Items.FirstOrDefault((Item) => Item.GetType() == typeof(MyChildRenderItem) && (Item as MyChildRenderItem).IsArabic) as MyChildRenderItem).ItemRuns = strs[(MainItem != null) ? 2 : 1].Select((Item) => new MyChildRenderBlockItem() { ItemText = (string)Item.Text, Clr = Item.Clr }).ToList();
                        XMLRender.RenderArray.RenderText[][] rt = null;
                        strs = AppSettings.Arb.TransliterateWithRulesColor(text, String.Empty, true, false, false, IslamMetadata.Arabic.FilterMetadataStops(text, (obj as MyChildRenderStopContinue).MetaRules, stops.ToArray()), ref rt);
                        (panel.RenderModels[mmodel - (index == 0 ? 1 : 0)].RenderItems[index == 0 ? panel.RenderModels[mmodel - 1].RenderItems.Count - 1 : index - 1].Items.FirstOrDefault((Item) => Item.GetType() == typeof(MyChildRenderItem) && !(Item as MyChildRenderItem).IsArabic) as MyChildRenderItem).ItemRuns = strs[0].Select((Item) => new MyChildRenderBlockItem() { ItemText = (string)Item.Text, Clr = Item.Clr }).ToList();
                        if (MainItem != null && MainItem.GetText.Trim().Length != 1) (model.RenderItems[index].Items.FirstOrDefault((Item) => Item.GetType() == typeof(MyChildRenderItem) && !(Item as MyChildRenderItem).IsArabic) as MyChildRenderItem).ItemRuns = strs[1].Select((Item) => new MyChildRenderBlockItem() { ItemText = (string)Item.Text, Clr = Item.Clr }).ToList();
                        (panel.RenderModels[mmodel + (index + 1 == model.RenderItems.Count ? 1 : 0)].RenderItems[index + 1 == model.RenderItems.Count ? 0 : index + 1].Items.FirstOrDefault((Item) => Item.GetType() == typeof(MyChildRenderItem) && !(Item as MyChildRenderItem).IsArabic) as MyChildRenderItem).ItemRuns = strs[(MainItem != null && MainItem.GetText.Trim().Length != 1) ? 2 : 1].Select((Item) => new MyChildRenderBlockItem() { ItemText = (string)Item.Text, Clr = Item.Clr }).ToList();
                    }
                    else
                    {
                        string[] strs = AppSettings.Arb.TransliterateWithRules(text, String.Empty, false, IslamMetadata.Arabic.FilterMetadataStops(text, (obj as MyChildRenderStopContinue).MetaRules, stops.ToArray())).Split(' ');
                        (panel.RenderModels[mmodel - (index == 0 ? 1 : 0)].RenderItems[index == 0 ? panel.RenderModels[mmodel - 1].RenderItems.Count - 1 : index - 1].Items.FirstOrDefault((Item) => Item.GetType() == typeof(MyChildRenderItem) && (Item as MyChildRenderItem).IsArabic) as MyChildRenderItem).ItemRuns = new List<MyChildRenderBlockItem>() { new MyChildRenderBlockItem() { ItemText = strs[0] } };
                        if (MainItem != null && MainItem.GetText.Trim().Length != 1) MainItem.ItemRuns = new List<MyChildRenderBlockItem>() { new MyChildRenderBlockItem() { ItemText = strs[1] } };
                        (panel.RenderModels[mmodel + (index + 1 == model.RenderItems.Count ? 1 : 0)].RenderItems[index + 1 == model.RenderItems.Count ? 0 : index + 1].Items.FirstOrDefault((Item) => Item.GetType() == typeof(MyChildRenderItem) && (Item as MyChildRenderItem).IsArabic) as MyChildRenderItem).ItemRuns = new List<MyChildRenderBlockItem>() { new MyChildRenderBlockItem() { ItemText = strs[(MainItem != null && MainItem.GetText.Trim().Length != 1) ? 2 : 1] } };
                    }
                    //(((_holdObj as StackPanel).DataContext as MyRenderItem).Items[idx - 1] as MyChildRenderItem).GetText
                }
            }
            gestRec.CompleteGesture();
        }
        private void OnSizeChanged(object sender, SizeChangedEventArgs e)
        {
            UIChanger.MaxWidth = MainGrid.ActualWidth;
            UIChanger.MaxHeight = MainGrid.ActualHeight;
            this.ViewModel.RegroupRenderModels(UIChanger.MaxWidth);
        }

        protected override async void OnNavigatedTo(NavigationEventArgs e)
        {
            dynamic c = e.Parameter;
            double curWidth = UIChanger.MaxWidth;
            Division = c.Division;
            Selection = c.Selection;
            int JumpToChapter = c.JumpToChapter;
            int JumpToVerse = c.JumpToVerse;
            this.ViewModel.RenderModels = VirtualizingWrapPanelAdapter.GroupRenderModels(System.Linq.Enumerable.Select((await AppSettings.TR.GetRenderedQuranText(AppSettings.bShowTransliteration ? XMLRender.ArabicData.TranslitScheme.RuleBased : XMLRender.ArabicData.TranslitScheme.None, String.Empty, String.Empty, Division.ToString(), Selection.ToString(), !AppSettings.bShowTranslation ? "None" : AppSettings.ChData.IslamData.Translations.TranslationList[AppSettings.iSelectedTranslation].FileName, AppSettings.bShowW4W ? "0" : "4", AppSettings.bUseColoring ? "0" : "1")).Items, (Arr) => new MyRenderItem((XMLRender.RenderArray.RenderItem)Arr)).ToList(), curWidth);
            if (curWidth == 0 && UIChanger.MaxWidth != 0) OnSizeChanged(this, null);
            if (JumpToChapter != -1)
            {
                while (ViewModel.VerseReferences.Count() != CurrentPlayingItem && (ViewModel.VerseReferences[CurrentPlayingItem].Chapter != JumpToChapter || ViewModel.VerseReferences[CurrentPlayingItem].Verse != JumpToVerse)) { CurrentPlayingItem++; if (ViewModel.VerseReferences.Count() == CurrentPlayingItem) CurrentPlayingItem = 0; }
            }
            else {
                while (ViewModel.VerseReferences.Count() != CurrentPlayingItem && ViewModel.VerseReferences[CurrentPlayingItem].Chapter == -1) { CurrentPlayingItem++; if (ViewModel.VerseReferences.Count() == CurrentPlayingItem) CurrentPlayingItem = 0; }
            }
            ViewModel.CurrentVerse = ViewModel.VerseReferences[CurrentPlayingItem];
            if (JumpToChapter != -1) { ScrollToVerse(ViewModel.CurrentVerse); }
            if (c.StartPlaying) { PlayPause_Click(null, null); }
            LoadingRing.IsActive = false;
        }
        public MyUIChanger UIChanger { get; set; }
        public VirtualizingWrapPanelAdapter ViewModel { get; set; }
        private async void RenderPngs_Click(object sender, RoutedEventArgs e)
        {
            string origLang = Windows.Globalization.ApplicationLanguages.PrimaryLanguageOverride;
            string[] LangList = Windows.Globalization.ApplicationLanguages.ManifestLanguages.ToArray();
            //string[] LangList = "ar,ar-sa,ar-ae,ar-bh,ar-dz,ar-eg,ar-iq,ar-jo,ar-kw,ar-lb,ar-ly,ar-ma,ar-om,ar-qa,ar-sy,ar-tn,ar-ye,af,af-za,sq,sq-al,am,am-et,hy,hy-am,as,as-in,az,az-arab,az-arab-az,az-cyrl,az-cyrl-az,az-latn,az-latn-az,eu,eu-es,be,be-by,bn,bn-bd,bn-in,bs,bs-cyrl,bs-cyrl-ba,bs-latn,bs-latn-ba,bg,bg-bg,ca,ca-es,ca-es-valencia,chr-cher,chr-cher-us,chr-latn,zh,zh-Hans,zh-cn,zh-hans-cn,zh-sg,zh-hans-sg,zh-Hant,zh-hk,zh-mo,zh-tw,zh-hant-hk,zh-hant-mo,zh-hant-tw,hr,hr-hr,hr-ba,cs,cs-cz,da,da-dk,prs,prs-af,prs-arab,nl,nl-nl,nl-be,en,en-au,en-ca,en-gb,en-ie,en-in,en-nz,en-sg,en-us,en-za,en-bz,en-hk,en-id,en-jm,en-kz,en-mt,en-my,en-ph,en-pk,en-tt,en-vn,en-zw,en-053,en-021,en-029,en-011,en-018,en-014,et,et-ee,fil,fil-latn,fil-ph,fi,fi-fi,fr,fr-be,fr-ca,fr-ch,fr-fr,fr-lu,fr-015,fr-cd,fr-ci,fr-cm,fr-ht,fr-ma,fr-mc,fr-ml,fr-re,frc-latn,frp-latn,fr-155,fr-029,fr-021,fr-011,gl,gl-es,ka,ka-ge,de,de-at,de-ch,de-de,de-lu,de-li,el,el-gr,gu,gu-in,ha,ha-latn,ha-latn-ng,he,he-il,hi,hi-in,hu,hu-hu,is,is-is,ig-latn,ig-ng,id,id-id,iu-cans,iu-latn,iu-latn-ca,ga,ga-ie,xh,xh-za,zu,zu-za,it,it-it,it-ch,ja,ja-jp,kn,kn-in,kk,kk-kz,km,km-kh,quc-latn,qut-gt,qut-latn,rw,rw-rw,sw,sw-ke,kok,kok-in,ko,ko-kr,ku-arab,ku-arab-iq,ky-kg,ky-cyrl,lo,lo-la,lv,lv-lv,lt,lt-lt,lb,lb-lu,mk,mk-mk,ms,ms-bn,ms-my,ml,ml-in,mt,mt-mt,mi,mi-latn,mi-nz,mr,mr-in,mn-cyrl,mn-mong,mn-mn,mn-phag,ne,ne-np,nb,nb-no,nn,nn-no,no,no-no,,or,or-in,fa,fa-ir,pl,pl-pl,pt-br,pt,pt-pt,pa,pa-arab,pa-arab-pk,pa-deva,pa-in,quz,quz-bo,quz-ec,quz-pe,ro,ro-ro,ru,ru-ru,gd-gb,gd-latn,sr-Latn,sr-latn-cs,sr,sr-latn-ba,sr-latn-me,sr-latn-rs,sr-cyrl,sr-cyrl-ba,sr-cyrl-cs,sr-cyrl-me,sr-cyrl-rs,nso,nso-za,tn,tn-bw,tn-za,sd-arab,sd-arab-pk,sd-deva,si,si-lk,sk,sk-sk,sl,sl-si,es,es-cl,es-co,es-es,es-mx,es-ar,es-bo,es-cr,es-do,es-ec,es-gt,es-hn,es-ni,es-pa,es-pe,es-pr,es-py,es-sv,es-us,es-uy,es-ve,es-019,es-419,sv,sv-se,sv-fi,tg-arab,tg-cyrl,tg-cyrl-tj,tg-latn,ta,ta-in,tt-arab,tt-cyrl,tt-latn,tt-ru,te,te-in,th,th-th,ti,ti-et,tr,tr-tr,tk-cyrl,tk-latn,tk-tm,tk-latn-tr,tk-cyrl-tr,uk,uk-ua,ur,ur-pk,ug-arab,ug-cn,ug-cyrl,ug-latn,uz,uz-cyrl,uz-latn,uz-latn-uz,vi,vi-vn,cy,cy-gb,wo,wo-sn,yo-latn,yo-ng".Split(',');
            double curWidth = UIChanger.MaxWidth;
            for (int count = 0; count < LangList.Length; count++)
            {
                AppSettings.strAppLanguage = LangList[count];
                //Windows.Storage.KnownFolders.PicturesLibrary
                await Windows.Storage.ApplicationData.Current.LocalFolder.CreateFolderAsync(LangList[count], Windows.Storage.CreationCollisionOption.OpenIfExists);
                //re-render the UI
                //this.Frame.Navigate(this.GetType(), new { Division = Division, Selection = Selection});
                //this.Frame.BackStack.Remove(this.Frame.BackStack.Last());
                this.ViewModel.RenderModels = VirtualizingWrapPanelAdapter.GroupRenderModels(System.Linq.Enumerable.Select((await AppSettings.TR.GetRenderedQuranText(AppSettings.bShowTransliteration ? XMLRender.ArabicData.TranslitScheme.RuleBased : XMLRender.ArabicData.TranslitScheme.None, String.Empty, String.Empty, Division.ToString(), Selection.ToString(), String.Empty, AppSettings.bShowW4W ? "0" : "4", AppSettings.bUseColoring ? "0" : "1")).Items, (Arr) => new MyRenderItem((XMLRender.RenderArray.RenderItem)Arr)).ToList(), curWidth);
                UpdateLayout();
                await Dispatcher.RunAsync(Windows.UI.Core.CoreDispatcherPriority.Low, () => { });
                await ExtSplashScreen.SavePathImageAsFile(1366, 768, LangList[count] + "\\appstorescreenshot-wide", this, true);
                await ExtSplashScreen.SavePathImageAsFile(768, 1366, LangList[count] + "\\appstorescreenshot-tall", this, true);
                await ExtSplashScreen.SavePathImageAsFile(1280, 768, LangList[count] + "\\appstorephonescreenshot-wide", this, true);
                await ExtSplashScreen.SavePathImageAsFile(768, 1280, LangList[count] + "\\appstorephonescreenshot-tall", this, true);
                await ExtSplashScreen.SavePathImageAsFile(1280, 720, LangList[count] + "\\appstorephonescreenshot1280x720-wide", this, true);
                await ExtSplashScreen.SavePathImageAsFile(720, 1280, LangList[count] + "\\appstorephonescreenshot720x1280-tall", this, true);
                await ExtSplashScreen.SavePathImageAsFile(800, 480, LangList[count] + "\\appstorephonescreenshot800x480-wide", this, true);
                await ExtSplashScreen.SavePathImageAsFile(480, 800, LangList[count] + "\\appstorephonescreenshot480x800-tall", this, true);
            }
            AppSettings.strAppLanguage = origLang;
            this.ViewModel.RenderModels = VirtualizingWrapPanelAdapter.GroupRenderModels(System.Linq.Enumerable.Select((await AppSettings.TR.GetRenderedQuranText(AppSettings.bShowTransliteration ? XMLRender.ArabicData.TranslitScheme.RuleBased : XMLRender.ArabicData.TranslitScheme.None, String.Empty, String.Empty, Division.ToString(), Selection.ToString(), String.Empty, AppSettings.bShowW4W ? "0" : "4", AppSettings.bUseColoring ? "0" : "1")).Items, (Arr) => new MyRenderItem((XMLRender.RenderArray.RenderItem)Arr)).ToList(), curWidth);
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
            this.ViewModel.RegroupRenderModels(UIChanger.MaxWidth);
        }

        private void ZoomOut_Click(object sender, RoutedEventArgs e)
        {
            if (UIChanger.OtherFontSize > 1.0)
            {
                UIChanger.FontSize -= 1.0;
                UIChanger.OtherFontSize -= 1.0;
                this.ViewModel.RegroupRenderModels(UIChanger.MaxWidth);
            }
        }

        private void Default_Click(object sender, RoutedEventArgs e)
        {
            UIChanger.FontSize = 30.0;
            UIChanger.OtherFontSize = 20.0;
            this.ViewModel.RegroupRenderModels(UIChanger.MaxWidth);
        }
        private void PlayPause_Click(object sender, RoutedEventArgs e)
        {
            if ((PlayPause.Icon as SymbolIcon).Symbol == Symbol.Pause)
            {
                PlayPause.Icon = new SymbolIcon(Symbol.Play);
                PlayPause.Label = new Windows.ApplicationModel.Resources.ResourceLoader().GetString("Play/Label");
                VersePlayer.Pause();
            }
            else {
                PlayPause.Icon = new SymbolIcon(Symbol.Pause);
                PlayPause.Label = new Windows.ApplicationModel.Resources.ResourceLoader().GetString("Pause/Label");
                if (VersePlayer.Source != null) { VersePlayer.Play(); }
                else {
                    VersePlayer.Source = new Uri(IslamMetadata.AudioRecitation.GetURL(AppSettings.ChData.IslamData.ReciterList.Reciters[AppSettings.iSelectedReciter].Source, AppSettings.ChData.IslamData.ReciterList.Reciters[AppSettings.iSelectedReciter].Name, ViewModel.VerseReferences[CurrentPlayingItem].Chapter, ViewModel.VerseReferences[CurrentPlayingItem].Verse));
                }
            }
        }
        private void GoToVerse_Click(object sender, RoutedEventArgs e)
        {
            GoToVersePopup.IsOpen = true;
            ViewModel.PropertyChanged += NewVerseSelection;
            GoToVersePopup.Closed += PopupClosed;
        }

        private void PopupClosed(object sender, object e)
        {
            ViewModel.PropertyChanged -= NewVerseSelection;
            GoToVersePopup.Closed -= PopupClosed;
        }

        void NewVerseSelection(object sender, PropertyChangedEventArgs e)
        {
            if (e.PropertyName == "CurrentVerse")
            {
                GoToVersePopup.IsOpen = false;
                ScrollToVerse(ViewModel.CurrentVerse);
            }
        }
        private void ScrollToVerse(MyRenderItem Item)
        {
            ScrollViewer sv = VisualTreeHelper.GetChild(MainControl, 0) as ScrollViewer;
            //ScrollViewer sv = WinRTXamlToolkit.Controls.Extensions.ItemsControlExtensions.GetScrollViewer(MainControl);
            int count;
            for (count = 0; count < ViewModel.RenderModels.Count - 1; count++)
            {
                if (ViewModel.RenderModels[count].RenderItems.Contains(Item)) { break; }
            }
            FrameworkElement ItemContainer = MainControl.ContainerFromItem(ViewModel.RenderModels[count]) as FrameworkElement;
            //Point pos = ItemContainer.TransformToVisual(sv).TransformPoint(new Point());
            sv.ChangeView(null, count + 2/*sv.VerticalOffset + pos.Y*/, null, true);
        }
        private int CurrentPlayingItem;
        private async void MediaElement_CurrentStateChanged(object sender, RoutedEventArgs e)
        {
            MediaElement mediaElement = sender as MediaElement;
            if (mediaElement != null) {
                //if (mediaElement.CurrentState == Windows.UI.Xaml.Media.MediaElementState.Opening) { mediaElement.Play(); }
                if (mediaElement.CurrentState == Windows.UI.Xaml.Media.MediaElementState.Opening || mediaElement.CurrentState == Windows.UI.Xaml.Media.MediaElementState.Playing || mediaElement.CurrentState == Windows.UI.Xaml.Media.MediaElementState.Buffering) {
                } else if (mediaElement.CurrentState == Windows.UI.Xaml.Media.MediaElementState.Closed || mediaElement.CurrentState == Windows.UI.Xaml.Media.MediaElementState.Stopped)
                {
                    PlayPause.Icon = new SymbolIcon(Symbol.Play);
                    PlayPause.Label = new Windows.ApplicationModel.Resources.ResourceLoader().GetString("Play/Label");
                } else if (((PlayPause.Icon as SymbolIcon).Symbol == Symbol.Pause) && mediaElement.CurrentState == Windows.UI.Xaml.Media.MediaElementState.Paused) {
                    if (!AppSettings.bAutomaticAdvanceVerse) return;
                    if (AppSettings.bDelayVerseLengthBeforeAdvancing || AppSettings.iAdditionalVerseAdvanceDelay != 0) {
                        await System.Threading.Tasks.Task.Delay((AppSettings.bDelayVerseLengthBeforeAdvancing ? (int)VersePlayer.NaturalDuration.TimeSpan.TotalMilliseconds : 0) + AppSettings.iAdditionalVerseAdvanceDelay * 1000);
                        if (PlayPause.Label == new Windows.ApplicationModel.Resources.ResourceLoader().GetString("Play/Label")) { return; }
                    }
                    if (AppSettings.LoopingMode == 1) { VersePlayer.Play(); return; }
                    do {
                        CurrentPlayingItem++;
                        if (ViewModel.VerseReferences.Count() == CurrentPlayingItem) {
                            if (AppSettings.LoopingMode == 2) {
                                CurrentPlayingItem = 0;
                            } else {
                                int max = (int)AppSettings.TR.GetSelectionNames(Division.ToString(), XMLRender.ArabicData.TranslitScheme.RuleBased, String.Empty).Last().GetValue(1);
                                if (AppSettings.LoopingMode == 0 && Selection != max) {
                                    this.Frame.Navigate(typeof(WordForWordUC), new { Division = Division, Selection = Selection + 1, JumpToChapter = -1, JumpToVerse = -1, StartPlaying = true });
                                    this.Frame.BackStack.Remove(this.Frame.BackStack.Last());
                                    return;
                                }
                                else if (AppSettings.LoopingMode == 3) {
                                    this.Frame.Navigate(typeof(WordForWordUC), new { Division = Division, Selection = Selection == max ? 0 : (Selection + 1), JumpToChapter = -1, JumpToVerse = -1, StartPlaying = true });
                                    this.Frame.BackStack.Remove(this.Frame.BackStack.Last());
                                    return;
                                } else {
                                    CurrentPlayingItem--;
                                    return;
                                }
                            }
                        }
                    } while (ViewModel.VerseReferences.Count() != CurrentPlayingItem && ViewModel.VerseReferences[CurrentPlayingItem].Chapter == -1);
                    ViewModel.CurrentVerse = ViewModel.VerseReferences[CurrentPlayingItem];
                    ScrollToVerse(ViewModel.VerseReferences[CurrentPlayingItem]);
                    VersePlayer.Source = new Uri(IslamMetadata.AudioRecitation.GetURL(AppSettings.ChData.IslamData.ReciterList.Reciters[AppSettings.iSelectedReciter].Source, AppSettings.ChData.IslamData.ReciterList.Reciters[AppSettings.iSelectedReciter].Name, ViewModel.VerseReferences[CurrentPlayingItem].Chapter, ViewModel.VerseReferences[CurrentPlayingItem].Verse));
                }
            }
        }
        private void StackPanel_Holding(object sender, HoldingRoutedEventArgs e)
        {
            if (e.HoldingState == Windows.UI.Input.HoldingState.Started)
            {
                e.Handled = true;
                DoHolding(sender);
            }
        }
        void DoHolding(object sender)
        {
            if (((sender as StackPanel).DataContext as MyRenderItem).Word != -1)
            {
                Flyout ContextFlyout = new Flyout();
                StackPanel Panel = new StackPanel();
                //Panel.SetValue(NameProperty, "TopLevel");
                //Panel.Name = "TopLevel";
                Panel.DataContext = this;
                ContextFlyout.Content = Panel;
                Panel.Children.Add(new ItemsControl() { ItemsPanel = Resources["VirtualPanelTemplate"] as ItemsPanelTemplate, ItemTemplate = Resources["WrapTemplate"] as DataTemplate, ItemsSource = VirtualizingWrapPanelAdapter.GroupRenderModels(System.Linq.Enumerable.Select(AppSettings.ChData.GetMorphologicalDataForWord(((sender as StackPanel).DataContext as MyRenderItem).Chapter, ((sender as StackPanel).DataContext as MyRenderItem).Verse, ((sender as StackPanel).DataContext as MyRenderItem).Word).Items, (Arr) => new MyRenderItem((XMLRender.RenderArray.RenderItem)Arr)).ToList(), UIChanger.MaxWidth) });
                VirtualizingStackPanel.SetVirtualizationMode((Panel.Children.Last() as ItemsControl), VirtualizationMode.Recycling);
                Panel.Children.Add(new Button() { Content = new Windows.ApplicationModel.Resources.ResourceLoader().GetString("CopyToClipboard/Text") });
                (Panel.Children.Last() as Button).Click += (object _sender, RoutedEventArgs _e) =>
                {
#if WINDOWS_APP || WINDOWS_UWP
                    Windows.ApplicationModel.DataTransfer.DataPackage package = new Windows.ApplicationModel.DataTransfer.DataPackage(); package.SetText(((sender as StackPanel).DataContext as MyRenderItem).GetText); Windows.ApplicationModel.DataTransfer.Clipboard.SetContent(package);
#else
                    global::Windows.System.LauncherOptions options = new global::Windows.System.LauncherOptions(); 
                    options.PreferredApplicationDisplayName = "Clipboarder"; 
                    options.PreferredApplicationPackageFamilyName = "InTheHandLtd.Clipboarder"; 
                    options.DisplayApplicationPicker = false; 
                    global::Windows.System.Launcher.LaunchUriAsync(new Uri(string.Format("clipboard:Set?Text={0}", Uri.EscapeDataString(string.Empty))), options); 
#endif
                };
                Panel.Children.Add(new Button() { Content = new TextBlock() { Text = new Windows.ApplicationModel.Resources.ResourceLoader().GetString("Share/Text") } });
                (Panel.Children.Last() as Button).Click += (object _sender, RoutedEventArgs _e) =>
                {
//#if WINDOWS_APP
                Windows.ApplicationModel.DataTransfer.DataPackage package = new Windows.ApplicationModel.DataTransfer.DataPackage();
                    package.SetText(((sender as StackPanel).DataContext as MyRenderItem).GetText);
                    Windows.ApplicationModel.DataTransfer.DataTransferManager.GetForCurrentView().DataRequested += (Windows.ApplicationModel.DataTransfer.DataTransferManager __sender, Windows.ApplicationModel.DataTransfer.DataRequestedEventArgs __e) =>
                    {
                        __e.Request.Data.Properties.Title = new Windows.ApplicationModel.Resources.ResourceLoader().GetString("DisplayName");
                        __e.Request.Data.Properties.Description = new Windows.ApplicationModel.Resources.ResourceLoader().GetString("Description");
                        __e.Request.Data.SetText(((sender as StackPanel).DataContext as MyRenderItem).GetText);
                        Windows.ApplicationModel.DataTransfer.DataTransferManager.GetForCurrentView().DataRequested -= null;
                    };
                    Windows.ApplicationModel.DataTransfer.DataTransferManager.ShowShareUI();
                //#else
                //            Microsoft.Phone.Tasks.ShareStatusTask shareStatusTask = new Microsoft.Phone.Tasks.ShareStatusTask(); 
                //            shareStatusTask.Status = string.Empty;
                //            shareStatusTask.Show(); 
                //#endif
                };
                FlyoutBase.SetAttachedFlyout(sender as StackPanel, ContextFlyout);
                FlyoutBase.ShowAttachedFlyout(sender as StackPanel);
            }
            else if (((sender as StackPanel).DataContext as MyRenderItem).Chapter != -1)
            {
                MenuFlyout ContextFlyout = new MenuFlyout();
                int idxMark = Array.FindIndex(AppSettings.Bookmarks, (Item) => Item[0] == Division && Item[1] == Selection && Item[2] == ((sender as StackPanel).DataContext as MyRenderItem).Chapter && Item[3] == ((sender as StackPanel).DataContext as MyRenderItem).Verse);
                ContextFlyout.Items.Add(new MenuFlyoutItem() { Text = new Windows.ApplicationModel.Resources.ResourceLoader().GetString((idxMark == -1 ? "AddBookmark" : "RemoveBookmark") + "/Text") });
                (ContextFlyout.Items.Last() as MenuFlyoutItem).Click += (object _sender, RoutedEventArgs _e) =>
                {
                    List<int[]> marks = AppSettings.Bookmarks.ToList();
                    if (idxMark != -1) {
                        marks.RemoveAt(idxMark);
                    } else {
                        marks.Add(new int[] { Division, Selection, ((sender as StackPanel).DataContext as MyRenderItem).Chapter, ((sender as StackPanel).DataContext as MyRenderItem).Verse });
                    }
                    AppSettings.Bookmarks = marks.ToArray();
                };
                ContextFlyout.Items.Add(new MenuFlyoutItem() { Text = new Windows.ApplicationModel.Resources.ResourceLoader().GetString("CopyToClipboard/Text") });
                (ContextFlyout.Items.Last() as MenuFlyoutItem).Click += (object _sender, RoutedEventArgs _e) =>
                {
#if WINDOWS_APP || WINDOWS_UWP
                    Windows.ApplicationModel.DataTransfer.DataPackage package = new Windows.ApplicationModel.DataTransfer.DataPackage(); package.SetText(((sender as StackPanel).DataContext as MyRenderItem).GetText); Windows.ApplicationModel.DataTransfer.Clipboard.SetContent(package);
#else
                    global::Windows.System.LauncherOptions options = new global::Windows.System.LauncherOptions(); 
                    options.PreferredApplicationDisplayName = "Clipboarder"; 
                    options.PreferredApplicationPackageFamilyName = "InTheHandLtd.Clipboarder"; 
                    options.DisplayApplicationPicker = false; 
                    global::Windows.System.Launcher.LaunchUriAsync(new Uri(string.Format("clipboard:Set?Text={0}", Uri.EscapeDataString(string.Empty))), options); 
#endif
                };
                ContextFlyout.Items.Add(new MenuFlyoutItem() { Text = new Windows.ApplicationModel.Resources.ResourceLoader().GetString("Share/Text") });
                (ContextFlyout.Items.Last() as MenuFlyoutItem).Click += (object _sender, RoutedEventArgs _e) =>
                {
//#if WINDOWS_APP
                    Windows.ApplicationModel.DataTransfer.DataPackage package = new Windows.ApplicationModel.DataTransfer.DataPackage();
                    package.SetText(((sender as StackPanel).DataContext as MyRenderItem).GetText);
                    Windows.ApplicationModel.DataTransfer.DataTransferManager.GetForCurrentView().DataRequested += (Windows.ApplicationModel.DataTransfer.DataTransferManager __sender, Windows.ApplicationModel.DataTransfer.DataRequestedEventArgs __e) =>
                    {
                        __e.Request.Data.Properties.Title = new Windows.ApplicationModel.Resources.ResourceLoader().GetString("DisplayName");
                        __e.Request.Data.Properties.Description = new Windows.ApplicationModel.Resources.ResourceLoader().GetString("Description");
                        __e.Request.Data.SetText(((sender as StackPanel).DataContext as MyRenderItem).GetText);
                        Windows.ApplicationModel.DataTransfer.DataTransferManager.GetForCurrentView().DataRequested -= null;
                    };
                    Windows.ApplicationModel.DataTransfer.DataTransferManager.ShowShareUI();
//#else
//            Microsoft.Phone.Tasks.ShareStatusTask shareStatusTask = new Microsoft.Phone.Tasks.ShareStatusTask(); 
//            shareStatusTask.Status = string.Empty;
//            shareStatusTask.Show(); 
//#endif
                };
                ContextFlyout.Items.Add(new MenuFlyoutItem() { Text = new Windows.ApplicationModel.Resources.ResourceLoader().GetString("SetPlaybackVerse/Text") });
                (ContextFlyout.Items.Last() as MenuFlyoutItem).Click += (object _sender, RoutedEventArgs _e) =>
                {
                    ViewModel.CurrentVerse = (sender as StackPanel).DataContext as MyRenderItem;
                    for (int count = 0; count <= ViewModel.VerseReferences.Count() - 1; count++)
                    {
                        if (ViewModel.VerseReferences[count] == ViewModel.CurrentVerse) { CurrentPlayingItem = count; break; }
                        VersePlayer.Source = null;
                    }
                };
                //ContextFlyout.Items.Add(new MenuFlyoutItem() { Text = new Windows.ApplicationModel.Resources.ResourceLoader().GetString("ShowExegesis/Text") });
                //(ContextFlyout.Items.Last() as MenuFlyoutItem).Click += (object _sender, RoutedEventArgs _e) => { };
                FlyoutBase.SetAttachedFlyout(sender as StackPanel, ContextFlyout);
                FlyoutBase.ShowAttachedFlyout(sender as StackPanel);
            }
        }
    }
    public class WidthLimiterConverter : DependencyObject, IValueConverter
    {
        public double TopLevelMaxValue
        {
            get { return (double)GetValue(TopLevelMaxValueProperty); }
            set { SetValue(TopLevelMaxValueProperty, value); }
        }

        public static readonly DependencyProperty TopLevelMaxValueProperty =
            DependencyProperty.Register("TopLevelMaxValue",
                                        typeof(double),
                                        typeof(WidthLimiterConverter),
                                        new PropertyMetadata(0.0));

        public object Convert(object value, Type targetType, object parameter, string language)
        {
            return Math.Min((double)value, TopLevelMaxValue);
        }

        public object ConvertBack(object value, Type targetType, object parameter, string language)
        {
            throw new NotImplementedException();
        }
    }
    public class SurfaceSourceConverter : DependencyObject, IValueConverter
    {
        public double TopLevelFontSize
        {
            get { return (double)GetValue(TopLevelFontSizeProperty); }
            set { SetValue(TopLevelFontSizeProperty, value); }
        }

        public static readonly DependencyProperty TopLevelFontSizeProperty =
            DependencyProperty.Register("TopLevelFontSize",
                                        typeof(double),
                                        typeof(SurfaceSourceConverter),
                                        new PropertyMetadata(0.0));
        public object Convert(object value, Type targetType, object parameter, string language)
        {
            MyChildRenderItem.RenderDataStruct Item = (MyChildRenderItem.RenderDataStruct)value;
            if (Item.FontSize != AppSettings.dFontSize)
            {
                Item.sizeFunc(); //cannot react this late
                Item = (MyChildRenderItem.RenderDataStruct)value;
            }
            int pixelWidth = (int)Item.Width, pixelHeight = (int)Item.Height;
            SurfaceImageSource newSource = new SurfaceImageSource(pixelWidth, pixelHeight, false);
            //SharpDX.Direct3D11.Device D3DDev = new SharpDX.Direct3D11.Device(SharpDX.Direct3D.DriverType.Hardware, SharpDX.Direct3D11.DeviceCreationFlags.BgraSupport);
            //SharpDX.DXGI.Device DXDev = D3DDev.QueryInterface<SharpDX.DXGI.Device>();
#if WINDOWS_PHONE_APP || !STORETOOLKIT
            SharpDX.DXGI.ISurfaceImageSourceNativeWithD2D surfaceImageSourceNative = SharpDX.ComObject.As<SharpDX.DXGI.ISurfaceImageSourceNativeWithD2D>(newSource);
#else
            SharpDX.DXGI.ISurfaceImageSourceNative surfaceImageSourceNative = SharpDX.ComObject.As<SharpDX.DXGI.ISurfaceImageSourceNative>(newSource);
#endif
#if WINDOWS_UWP
            SharpDX.Mathematics.Interop.RawRectangle rt = new SharpDX.Mathematics.Interop.RawRectangle(0, 0, pixelWidth, pixelHeight);
#else
            SharpDX.Rectangle rt = new SharpDX.Rectangle(0, 0, pixelWidth, pixelHeight);
#endif
#if WINDOWS_PHONE_APP || !STORETOOLKIT
#if WINDOWS_UWP
            SharpDX.Mathematics.Interop.RawPoint pt;
#else
            SharpDX.Point pt;
#endif
            IntPtr obj;
            surfaceImageSourceNative.Device = TextShaping.Dev2D;
            surfaceImageSourceNative.BeginDraw(rt, new Guid("e8f7fe7a-191c-466d-ad95-975678bda998"), out obj, out pt); //d2d1_1.h
            SharpDX.Direct2D1.DeviceContext devcxt = SharpDX.ComObject.As<SharpDX.Direct2D1.DeviceContext>(obj);
#else
            SharpDX.DrawingPoint pt;
            surfaceImageSourceNative.Device = TextShaping.DXDev;
            SharpDX.DXGI.Surface surf = surfaceImageSourceNative.BeginDraw(rt, out pt);
            SharpDX.Direct2D1.DeviceContext devcxt = new SharpDX.Direct2D1.DeviceContext(surf);
            //SharpDX.Direct2D1.Bitmap1 bmp = new SharpDX.Direct2D1.Bitmap1(devcxt, surf, new SharpDX.Direct2D1.BitmapProperties1() { DpiX = Windows.Graphics.Display.DisplayInformation.GetForCurrentView().RawDpiX, DpiY = Windows.Graphics.Display.DisplayInformation.GetForCurrentView().RawDpiY, PixelFormat = new SharpDX.Direct2D1.PixelFormat(SharpDX.DXGI.Format.B8G8R8A8_UNorm, SharpDX.Direct2D1.AlphaMode.Premultiplied), BitmapOptions = SharpDX.Direct2D1.BitmapOptions.CannotDraw | SharpDX.Direct2D1.BitmapOptions.Target });
            //devcxt.Target = bmp;
            devcxt.BeginDraw();
#endif
#if WINDOWS_UWP
            devcxt.Clear(new SharpDX.Mathematics.Interop.RawColor4(Windows.UI.Colors.White.R / 255.0f, Windows.UI.Colors.White.G / 255.0f, Windows.UI.Colors.White.B / 255.0f, Windows.UI.Colors.Transparent.A / 255.0f));
#else
            devcxt.Clear(new SharpDX.Color4(Windows.UI.Colors.White.R / 255.0f, Windows.UI.Colors.White.G / 255.0f, Windows.UI.Colors.White.B / 255.0f, Windows.UI.Colors.Transparent.A / 255.0f));
#endif
            SharpDX.Direct2D1.Layer lyr = new SharpDX.Direct2D1.Layer(devcxt);
            //SharpDX.RectangleF.Infinite
#if WINDOWS_UWP
            devcxt.PushLayer(new SharpDX.Direct2D1.LayerParameters1(new SharpDX.Mathematics.Interop.RawRectangleF(float.NegativeInfinity, float.NegativeInfinity, float.PositiveInfinity, float.PositiveInfinity), null, SharpDX.Direct2D1.AntialiasMode.PerPrimitive, new SharpDX.Mathematics.Interop.RawMatrix3x2(1, 0, 0, 1, 0, 0), 1.0f, null, SharpDX.Direct2D1.LayerOptions1.None), lyr);
            devcxt.PushAxisAlignedClip(new SharpDX.Mathematics.Interop.RawRectangleF(pt.X, pt.Y, pt.X + pixelWidth, pt.Y + pixelHeight), SharpDX.Direct2D1.AntialiasMode.PerPrimitive);
            devcxt.Transform = new SharpDX.Mathematics.Interop.RawMatrix3x2(1, 0, 0, 1, pt.X, pt.Y);
#else
            devcxt.PushLayer(new SharpDX.Direct2D1.LayerParameters1(new SharpDX.RectangleF(float.NegativeInfinity, float.NegativeInfinity, float.PositiveInfinity, float.PositiveInfinity), null, SharpDX.Direct2D1.AntialiasMode.PerPrimitive, SharpDX.Matrix3x2.Identity, 1.0f, null, SharpDX.Direct2D1.LayerOptions1.None), lyr);
            devcxt.PushAxisAlignedClip(new SharpDX.RectangleF(pt.X, pt.Y, pt.X + pixelWidth, pt.Y + pixelHeight), SharpDX.Direct2D1.AntialiasMode.PerPrimitive);
            devcxt.Transform = SharpDX.Matrix3x2.Translation(pt.X, pt.Y);
#endif
            SharpDX.DirectWrite.GlyphRun gr = new SharpDX.DirectWrite.GlyphRun();
            gr.FontFace = TextShaping.DWFontFace;
            gr.FontSize = (float)TopLevelFontSize;
            gr.BidiLevel = -1;
            int curlen = 0;
            for (int ct = 0; ct < Item.ItemRuns.Count(); ct++)
            {
                int newlen = curlen + Item.ItemRuns[ct].ItemText.Length;
                gr.Indices = Item.indices.Skip(Item.clusters[curlen]).TakeWhile((indice, idx) => ct == Item.ItemRuns.Count() - 1 || idx < Item.clusters[newlen]).ToArray();
                gr.Offsets = Item.offsets.Skip(Item.clusters[curlen]).TakeWhile((offset, idx) => ct == Item.ItemRuns.Count() - 1 || idx < Item.clusters[newlen]).ToArray();
                gr.Advances = Item.advances.Skip(Item.clusters[curlen]).TakeWhile((advance, idx) => ct == Item.ItemRuns.Count() - 1 || idx < Item.clusters[newlen]).ToArray();
                if (Item.ItemRuns[ct].ItemText[0] == XMLRender.ArabicData.ArabicEndOfAyah) gr.Advances[0] = 0;
#if WINDOWS_UWP
                SharpDX.Direct2D1.SolidColorBrush brsh = new SharpDX.Direct2D1.SolidColorBrush(devcxt, new SharpDX.Mathematics.Interop.RawColor4(XMLRender.Utility.ColorR(Item.ItemRuns[ct].Clr) / 255.0f, XMLRender.Utility.ColorG(Item.ItemRuns[ct].Clr) / 255.0f, XMLRender.Utility.ColorB(Item.ItemRuns[ct].Clr) / 255.0f, 0xFF / 255.0f));
                devcxt.DrawGlyphRun(new SharpDX.Mathematics.Interop.RawVector2((float)pt.X + (float)pixelWidth + Item.offsets[0].AdvanceOffset - (Item.clusters[curlen] == 0 ? 0 : Item.advances.Take(Item.clusters[curlen]).Sum()), Item.BaseLine), gr, brsh, SharpDX.Direct2D1.MeasuringMode.Natural);
#else
                SharpDX.Direct2D1.SolidColorBrush brsh = new SharpDX.Direct2D1.SolidColorBrush(devcxt, new SharpDX.Color4(XMLRender.Utility.ColorR(Item.ItemRuns[ct].Clr) / 255.0f, XMLRender.Utility.ColorG(Item.ItemRuns[ct].Clr) / 255.0f, XMLRender.Utility.ColorB(Item.ItemRuns[ct].Clr) / 255.0f, 0xFF / 255.0f));
                devcxt.DrawGlyphRun(new SharpDX.Vector2((float) pt.X + (float)pixelWidth + Item.offsets[0].AdvanceOffset - (Item.clusters[curlen] == 0 ? 0 : Item.advances.Take(Item.clusters[curlen]).Sum()), Item.BaseLine), gr, brsh, SharpDX.Direct2D1.MeasuringMode.Natural);
#endif
                brsh.Dispose();
                curlen = newlen;
            }
#if WINDOWS_UWP
            devcxt.Transform = new SharpDX.Mathematics.Interop.RawMatrix3x2(1, 0, 0, 1, 0, 0);
#else
            devcxt.Transform = SharpDX.Matrix3x2.Identity;
#endif
            devcxt.PopAxisAlignedClip();
            devcxt.PopLayer();
#if WINDOWS_PHONE_APP || !STORETOOLKIT
            gr.Dispose();
#else
            gr.FontFace = null;
#endif
#if WINDOWS_PHONE_APP || !STORETOOLKIT
#else
            devcxt.EndDraw();
            devcxt.Target = null;
            //bmp.Dispose();
#endif
            devcxt.Dispose();
#if WINDOWS_PHONE_APP || !STORETOOLKIT
#else
            surf.Dispose();
#endif
            surfaceImageSourceNative.EndDraw();
            surfaceImageSourceNative.Device = null;
            surfaceImageSourceNative.Dispose();
            //DXDev.Dispose();
            //D3DDev.Dispose();
            return newSource;
        }

        public object ConvertBack(object value, Type targetType, object parameter, string language)
        {
            throw new NotImplementedException();
        }
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
                value.FirstOrDefault((it) => { textBlock.Inlines.Add(new Windows.UI.Xaml.Documents.Run() { Text = it.ItemText, Foreground = new SolidColorBrush(Windows.UI.Color.FromArgb(0xFF, XMLRender.Utility.ColorR(it.Clr), XMLRender.Utility.ColorG(it.Clr), XMLRender.Utility.ColorB(it.Clr))) }); return false; });
            }
        }

#endregion
    }
    public class RenderSourceConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, string language)
        {
            CollectionViewSource _RenderSource = new CollectionViewSource();
            BindingOperations.SetBinding(_RenderSource, CollectionViewSource.SourceProperty, new Binding() { Source = value, Path = new PropertyPath("RenderModels"), Mode = BindingMode.OneWay });
            return _RenderSource.View;
        }

        public object ConvertBack(object value, Type targetType, object parameter, string language)
        {
            throw new NotImplementedException();
        }
    }
    public class BackgroundSelectedConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, string language)
        {
            return new SolidColorBrush((bool)value ? Windows.UI.Colors.Beige : Windows.UI.Colors.White);
        }

        public object ConvertBack(object value, Type targetType, object parameter, string language)
        {
            throw new NotImplementedException();
        }
    }
    public class StopContinueTextConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, string language)
        {
            return (bool)value ? "\u2B59" : "\u2B45";
        }

        public object ConvertBack(object value, Type targetType, object parameter, string language)
        {
            throw new NotImplementedException();
        }
    }
    public class StopContinueTextColorConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, string language)
        {
            return new SolidColorBrush((bool)value ? Windows.UI.Colors.Red : Windows.UI.Colors.Green);
        }

        public object ConvertBack(object value, Type targetType, object parameter, string language)
        {
            throw new NotImplementedException();
        }
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
        public DataTemplate ArabicRenderTemplate { get; set; }
        public DataTemplate NormalTemplate { get; set; }
        public DataTemplate StopContinueTemplate { get; set; }

        protected override DataTemplate SelectTemplateCore(object item, DependencyObject container)
        {
            if (item.GetType() == typeof(MyChildRenderStopContinue)) return StopContinueTemplate;
            if (item.GetType() == typeof(MyChildRenderItem)) { return ((MyChildRenderItem)item).IsArabic ? /*((MyChildRenderItem)item).ItemRuns.Count > 1 ? ArabicRenderTemplate :*/ ArabicTemplate : NormalTemplate; }
            return WordTemplate;
        }
    }
}
