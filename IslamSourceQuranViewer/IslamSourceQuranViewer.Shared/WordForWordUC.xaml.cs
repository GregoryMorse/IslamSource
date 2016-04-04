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
using Windows.UI.Xaml.Navigation;
using WinRTXamlToolkit.Controls;

namespace IslamSourceQuranViewer
{
    public class VirtualizingWrapPanelAdapterSource : VirtualizingWrapPanelAdapter
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
    }
    public sealed partial class WordForWordUC : Page
    {
        public WordForWordUC()
        {
            this.DataContext = this;
            this.ViewModel = new VirtualizingWrapPanelAdapterSource();
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
            gestRec.GestureSettings = Windows.UI.Input.GestureSettings.HoldWithMouse | Windows.UI.Input.GestureSettings.Hold | Windows.UI.Input.GestureSettings.RightTap;
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
                object obj = ((_holdObj as StackPanel).DataContext as MyRenderItem).Items.FirstOrDefault((Item) => Item.GetType() == typeof(MyChildRenderStopContinue));
                if (obj != null) { (obj as MyChildRenderStopContinue).IsStop = !(obj as MyChildRenderStopContinue).IsStop; }
            }
            gestRec.CompleteGesture();
        }
        private void OnSizeChanged(object sender, SizeChangedEventArgs e)
        {
            UIChanger.MaxWidth = MainGrid.ActualWidth;
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
            this.ViewModel.RenderModels = await System.Threading.Tasks.Task.Run(() => VirtualizingWrapPanelAdapter.GroupRenderModels(System.Linq.Enumerable.Select(IslamMetadata.TanzilReader.GetRenderedQuranText(AppSettings.bShowTransliteration ? XMLRender.ArabicData.TranslitScheme.RuleBased : XMLRender.ArabicData.TranslitScheme.None, String.Empty, String.Empty, Division.ToString(), Selection.ToString(), !AppSettings.bShowTranslation ? "None" : IslamMetadata.CachedData.IslamData.Translations.TranslationList[AppSettings.iSelectedTranslation].FileName, AppSettings.bShowW4W ? "0" : "4", AppSettings.bUseColoring ? "0" : "1").Items, (Arr) => new MyRenderItem((XMLRender.RenderArray.RenderItem)Arr)).ToList(), curWidth));
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
        public VirtualizingWrapPanelAdapterSource ViewModel { get; set; }
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
                this.ViewModel.RenderModels = await System.Threading.Tasks.Task.Run(() => VirtualizingWrapPanelAdapter.GroupRenderModels(System.Linq.Enumerable.Select(IslamMetadata.TanzilReader.GetRenderedQuranText(AppSettings.bShowTransliteration ? XMLRender.ArabicData.TranslitScheme.RuleBased : XMLRender.ArabicData.TranslitScheme.None, String.Empty, String.Empty, Division.ToString(), Selection.ToString(), String.Empty, AppSettings.bShowW4W ? "0" : "4", AppSettings.bUseColoring ? "0" : "1").Items, (Arr) => new MyRenderItem((XMLRender.RenderArray.RenderItem)Arr)).ToList(), curWidth));
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
            this.ViewModel.RenderModels = await System.Threading.Tasks.Task.Run(() => VirtualizingWrapPanelAdapter.GroupRenderModels(System.Linq.Enumerable.Select(IslamMetadata.TanzilReader.GetRenderedQuranText(AppSettings.bShowTransliteration ? XMLRender.ArabicData.TranslitScheme.RuleBased : XMLRender.ArabicData.TranslitScheme.None, String.Empty, String.Empty, Division.ToString(), Selection.ToString(), String.Empty, AppSettings.bShowW4W ? "0" : "4", AppSettings.bUseColoring ? "0" : "1").Items, (Arr) => new MyRenderItem((XMLRender.RenderArray.RenderItem)Arr)).ToList(), curWidth));
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
                    VersePlayer.Source = new Uri(IslamMetadata.AudioRecitation.GetURL(IslamMetadata.CachedData.IslamData.ReciterList.Reciters[AppSettings.iSelectedReciter].Source, IslamMetadata.CachedData.IslamData.ReciterList.Reciters[AppSettings.iSelectedReciter].Name, ViewModel.VerseReferences[CurrentPlayingItem].Chapter, ViewModel.VerseReferences[CurrentPlayingItem].Verse));
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
                    if (AppSettings.LoopingMode == IslamMetadata.CachedData.IslamData.LoopingModeList.LoopingModes[1].Name) { VersePlayer.Play(); return; }
                    do {
                        CurrentPlayingItem++;
                        if (ViewModel.VerseReferences.Count() == CurrentPlayingItem) {
                            if (AppSettings.LoopingMode == IslamMetadata.CachedData.IslamData.LoopingModeList.LoopingModes[2].Name) {
                                CurrentPlayingItem = 0;
                            } else {
                                int max = (int)IslamMetadata.TanzilReader.GetSelectionNames(Division.ToString(), XMLRender.ArabicData.TranslitScheme.RuleBased, String.Empty).Last().GetValue(1);
                                if (AppSettings.LoopingMode == IslamMetadata.CachedData.IslamData.LoopingModeList.LoopingModes[0].Name && Selection != max) {
                                    this.Frame.Navigate(typeof(WordForWordUC), new { Division = Division, Selection = Selection + 1, JumpToChapter = -1, JumpToVerse = -1, StartPlaying = true });
                                    this.Frame.BackStack.Remove(this.Frame.BackStack.Last());
                                    return;
                                }
                                else if (AppSettings.LoopingMode == IslamMetadata.CachedData.IslamData.LoopingModeList.LoopingModes[3].Name) {
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
                    VersePlayer.Source = new Uri(IslamMetadata.AudioRecitation.GetURL(IslamMetadata.CachedData.IslamData.ReciterList.Reciters[AppSettings.iSelectedReciter].Source, IslamMetadata.CachedData.IslamData.ReciterList.Reciters[AppSettings.iSelectedReciter].Name, ViewModel.VerseReferences[CurrentPlayingItem].Chapter, ViewModel.VerseReferences[CurrentPlayingItem].Verse));
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
                ContextFlyout.Content = Panel;
                Panel.Children.Add(new ItemsControl() { ItemsPanel = Resources["VirtualPanelTemplate"] as ItemsPanelTemplate, ItemTemplate = Resources["WrapTemplate"] as DataTemplate, ItemsSource = VirtualizingWrapPanelAdapter.GroupRenderModels(System.Linq.Enumerable.Select(IslamMetadata.CachedData.GetMorphologicalDataForWord(((sender as StackPanel).DataContext as MyRenderItem).Chapter, ((sender as StackPanel).DataContext as MyRenderItem).Verse, ((sender as StackPanel).DataContext as MyRenderItem).Word).Items, (Arr) => new MyRenderItem((XMLRender.RenderArray.RenderItem)Arr)).ToList(), UIChanger.MaxWidth) });
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
                    if (idxMark != -1)
                    {
                        marks.RemoveAt(idxMark);
                    }
                    else {
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
    public class StopContinueText : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, string language)
        {
            return (bool)value ? "\u2B45" : "\u2B59";
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
        public DataTemplate NormalTemplate { get; set; }
        public DataTemplate StopContinueTemplate { get; set; }

        protected override DataTemplate SelectTemplateCore(object item, DependencyObject container)
        {
            if (item.GetType() == typeof(MyChildRenderStopContinue)) return StopContinueTemplate;
            if (item.GetType() == typeof(MyChildRenderItem)) { return ((MyChildRenderItem)item).IsArabic ? ArabicTemplate : NormalTemplate; }
            return WordTemplate;
        }

    }
}
