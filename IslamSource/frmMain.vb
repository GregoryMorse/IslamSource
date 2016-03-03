Imports HostPageUtility
Imports IslamMetadata
Imports XMLRender
Public Class frmMain
    Public Class WindowsSettings
        Implements PortableSettings
        Public ReadOnly Property CacheDirectory As String Implements PortableSettings.CacheDirectory
            Get
                Return IO.Directory.GetCurrentDirectory()
            End Get
        End Property
        Public ReadOnly Property Resources As KeyValuePair(Of String, String())() Implements PortableSettings.Resources
            Get
                Return (New List(Of KeyValuePair(Of String, String()))(Linq.Enumerable.Select("XMLRender=Acct,lang,unicode;IslamResources=Hadith,IslamInfo,IslamSource".Split(";"c), Function(Str As String) New KeyValuePair(Of String, String())(Str.Split("="c)(0), Str.Split("="c)(1).Split(","c))))).ToArray()
            End Get
        End Property
        Public ReadOnly Property FuncLibs As String() Implements PortableSettings.FuncLibs
            Get
                Return {"IslamMetadata"}
            End Get
        End Property
        Public Function GetTemplatePath() As String Implements PortableSettings.GetTemplatePath
            Return GetFilePath("metadata\IslamSource.xml")
        End Function
        Public Function GetFilePath(ByVal Path As String) As String Implements PortableSettings.GetFilePath
            Return "..\..\..\" + Path
        End Function
        <Runtime.InteropServices.DllImport("getuname.dll", EntryPoint:="GetUName")>
        Friend Shared Function GetUName(ByVal wCharCode As UShort, <Runtime.InteropServices.MarshalAs(Runtime.InteropServices.UnmanagedType.LPWStr)> ByVal lpbuf As System.Text.StringBuilder) As Integer
        End Function
        Public Function GetUName(Character As Char) As String Implements PortableSettings.GetUName
            Dim Str As New System.Text.StringBuilder(512)
            Try
                GetUName(CUShort(AscW(Character)), Str)
            Catch e As System.DllNotFoundException
            End Try
            Return Str.ToString()
        End Function
        Public Function GetResourceString(baseName As String, resourceKey As String) As String Implements PortableSettings.GetResourceString
            Return New System.Resources.ResourceManager(baseName + ".Resources", Reflection.Assembly.Load(baseName)).GetString(resourceKey, Threading.Thread.CurrentThread.CurrentUICulture)
        End Function
    End Class
    Private Sub CompareWordCounts()
        Dim Text As List(Of String()) = TanzilReader.GetQuranText(CachedData.XMLDocMain, 1, 1, 114, 6)
        'Dim Lines As String() = Utility.ReadAllLines("..\..\..\metadata\en.w4w.corpus.txt")
        'Dim Lines As String() = Utility.ReadAllLines("..\..\..\metadata\en.w4w.qurandev.txt")
        'Dim Lines As String() = Utility.ReadAllLines("..\..\..\metadata\en.w4w.shehnazshaikh.txt")
        'Dim Lines As String() = Utility.ReadAllLines("..\..\..\metadata\id.w4w.terjemah.txt")
        Dim Lines As String() = Utility.ReadAllLines("..\..\..\metadata\ur.w4w.hafiznazarmuhammad.txt") '1000+ errors needs prefix fixing
        Dim LineCount As Integer = 0
        For Count As Integer = 0 To 113
            For VerseCount = 0 To Text(Count).Length - 1
                If New List(Of String)(Linq.Enumerable.Where(Text(Count)(VerseCount).Split(" "c), Function(It) It.Length <> 1)).Count <> Lines(LineCount).Split("|"c).Length Then
                    Debug.Print("Fault line: " + CStr(LineCount + 1) + "(" + CStr(Count + 1) + ":" + CStr(VerseCount + 1) + ") Arabic word count: " + CStr(New List(Of String)(Linq.Enumerable.Where(Text(Count)(VerseCount).Split(" "c), Function(It) It.Length <> 1)).Count) + " Translation word count: " + CStr(Lines(LineCount).Split("|"c).Length))
                End If
                LineCount += 1
            Next
        Next
    End Sub
    Private Sub UtilityTestCode()
        CompareWordCounts()
        'Dim AvailableTranslationProviders() As String = "af,af-ZA,am,am-ET,ar,ar-AE,ar-BH,ar-DZ,ar-EG,ar-IQ,ar-JO,ar-KW,ar-LB,ar-LY,ar-MA,ar-OM,ar-QA,ar-SA,ar-SY,ar-TN,ar-YE,as,as-IN,az-Latn,az-Latn-AZ,be,be-BY,bg,bg-BG,bn,bn-BD,bn-IN,bs-Cyrl,bs-Cyrl-BA,bs-Latn,bs-Latn-BA,ca,ca-ES,chr-Cher,chr-Cher-US,cs,cs-CZ,cy,cy-GB,da,da-DK,de,de-AT,de-CH,de-DE,de-LI,de-LU,el,el-GR,en,en-029,en-AU,en-BZ,en-CA,en-GB,en-HK,en-IE,en-IN,en-JM,en-MY,en-NG,en-NZ,en-PH,en-PK,en-SG,en-TT,en-US,en-ZA,en-ZW,es,es-AR,es-BO,es-CL,es-CO,es-CR,es-DO,es-EC,es-ES,es-GT,es-HN,es-MX,es-NI,es-PA,es-PE,es-PR,es-PY,es-SV,es-US,es-UY,es-VE,et,et-EE,eu,eu-ES,fa,fa-IR,fi,fi-FI,fil,fil-PH,fr,fr-BE,fr-CA,fr-CH,fr-DZ,fr-FR,fr-LU,fr-MA,fr-MC,fr-TN,ga,ga-IE,gd,gd-GB,gl,gl-ES,gu,gu-IN,ha-Latn,ha-Latn-NG,he,he-IL,hi,hi-IN,hr,hr-HR,hu,hu-HU,hy,hy-AM,id,id-ID,ig,ig-NG,is,is-IS,it,it-CH,it-IT,iu-Latn,iu-Latn-CA,ja,ja-JP,ka,ka-GE,kk,kk-KZ,km,km-KH,kn,kn-IN,ko,ko-KR,kok,kok-IN,ku-Arab,ku-Arab-IQ,ky,ky-KG,lb,lb-LU,lo,lo-LA,lt,lt-LT,lv,lv-LV,mi,mi-NZ,mk,mk-MK,ml,ml-IN,mn-MN,mr,mr-IN,ms,ms-BN,ms-MY,mt,mt-MT,my-MM,nb,nb-NO,ne,ne-NP,nl,nl-BE,nl-NL,nn,nn-NO,nso,nso-ZA,or,or-IN,pa,pa-Arab,pa-Arab-PK,pa-IN,pl,pl-PL,prs,prs-AF,ps,ps-AF,pt,pt-BR,pt-PT,qps-ploc,quc-Latn-GT,quz,quz-PE,rm,rm-CH,ro,ro-MD,ro-RO,ru,ru-KZ,ru-RU,rw,rw-RW,sd-Arab,sd-Arab-PK,si,si-LK,sk,sk-SK,sl,sl-SI,sq,sq-AL,sr-Cyrl,sr-Cyrl-BA,sr-Cyrl-RS,sr-Latn,sr-Latn-ME,sr-Latn-RS,sv,sv-FI,sv-SE,sw,sw-KE,ta,ta-IN,te,te-IN,tg-Cyrl,tg-Cyrl-TJ,th,th-TH,ti,ti-ET,tk,tk-TM,tn,tn-ZA,tr,tr-TR,tt,tt-RU,ug,ug-CN,uk,uk-UA,ur,ur-PK,uz-Cyrl-UZ,uz-Latn,uz-Latn-UZ,vi,vi-VN,wo,wo-SN,xh,xh-ZA,yo,yo-NG,zh-CN,zh-HK,zh-Hans,zh-Hant,zh-SG,zh-TW,zu,zu-ZA".ToLower().Split(","c)
        'Dim StoreCultures() As String = "ar,ar-sa,ar-ae,ar-bh,ar-dz,ar-eg,ar-iq,ar-jo,ar-kw,ar-lb,ar-ly,ar-ma,ar-om,ar-qa,ar-sy,ar-tn,ar-ye,af,af-za,sq,sq-al,am,am-et,hy,hy-am,as,as-in,az,az-arab,az-arab-az,az-cyrl,az-cyrl-az,az-latn,az-latn-az,eu,eu-es,be,be-by,bn,bn-bd,bn-in,bs,bs-cyrl,bs-cyrl-ba,bs-latn,bs-latn-ba,bg,bg-bg,ca,ca-es,ca-es-valencia,chr-cher,chr-cher-us,chr-latn,zh,zh-Hans,zh-cn,zh-hans-cn,zh-sg,zh-hans-sg,zh-Hant,zh-hk,zh-mo,zh-tw,zh-hant-hk,zh-hant-mo,zh-hant-tw,hr,hr-hr,hr-ba,cs,cs-cz,da,da-dk,prs,prs-af,prs-arab,nl,nl-nl,nl-be,en,en-au,en-ca,en-gb,en-ie,en-in,en-nz,en-sg,en-us,en-za,en-bz,en-hk,en-id,en-jm,en-kz,en-mt,en-my,en-ph,en-pk,en-tt,en-vn,en-zw,en-053,en-021,en-029,en-011,en-018,en-014,et,et-ee,fil,fil-latn,fil-ph,fi,fi-fi,fr,fr-be,fr-ca,fr-ch,fr-fr,fr-lu,fr-015,fr-cd,fr-ci,fr-cm,fr-ht,fr-ma,fr-mc,fr-ml,fr-re,frc-latn,frp-latn,fr-155,fr-029,fr-021,fr-011,gl,gl-es,ka,ka-ge,de,de-at,de-ch,de-de,de-lu,de-li,el,el-gr,gu,gu-in,ha,ha-latn,ha-latn-ng,he,he-il,hi,hi-in,hu,hu-hu,is,is-is,ig-latn,ig-ng,id,id-id,iu-cans,iu-latn,iu-latn-ca,ga,ga-ie,xh,xh-za,zu,zu-za,it,it-it,it-ch,ja,ja-jp,kn,kn-in,kk,kk-kz,km,km-kh,quc-latn,qut-gt,qut-latn,rw,rw-rw,sw,sw-ke,kok,kok-in,ko,ko-kr,ku-arab,ku-arab-iq,ky-kg,ky-cyrl,lo,lo-la,lv,lv-lv,lt,lt-lt,lb,lb-lu,mk,mk-mk,ms,ms-bn,ms-my,ml,ml-in,mt,mt-mt,mi,mi-latn,mi-nz,mr,mr-in,mn-cyrl,mn-mong,mn-mn,mn-phag,ne,ne-np,nb,nb-no,nn,nn-no,no,no-no,or,or-in,fa,fa-ir,pl,pl-pl,pt-br,pt,pt-pt,pa,pa-arab,pa-arab-pk,pa-deva,pa-in,quz,quz-bo,quz-ec,quz-pe,ro,ro-ro,ru,ru-ru,gd-gb,gd-latn,sr-Latn,sr-latn-cs,sr,sr-latn-ba,sr-latn-me,sr-latn-rs,sr-cyrl,sr-cyrl-ba,sr-cyrl-cs,sr-cyrl-me,sr-cyrl-rs,nso,nso-za,tn,tn-bw,tn-za,sd-arab,sd-arab-pk,sd-deva,si,si-lk,sk,sk-sk,sl,sl-si,es,es-cl,es-co,es-es,es-mx,es-ar,es-bo,es-cr,es-do,es-ec,es-gt,es-hn,es-ni,es-pa,es-pe,es-pr,es-py,es-sv,es-us,es-uy,es-ve,es-019,es-419,sv,sv-se,sv-fi,tg-arab,tg-cyrl,tg-cyrl-tj,tg-latn,ta,ta-in,tt-arab,tt-cyrl,tt-latn,tt-ru,te,te-in,th,th-th,ti,ti-et,tr,tr-tr,tk-cyrl,tk-latn,tk-tm,tk-latn-tr,tk-cyrl-tr,uk,uk-ua,ur,ur-pk,ug-arab,ug-cn,ug-cyrl,ug-latn,uz,uz-cyrl,uz-latn,uz-latn-uz,vi,vi-vn,cy,cy-gb,wo,wo-sn,yo-latn,yo-ng".ToLower().Split(","c)
        'StoreCultures = New List(Of String)(Linq.Enumerable.Intersect(AvailableTranslationProviders, StoreCultures)).ToArray()
        'Debug.Print(CStr(StoreCultures.Length) + ":" + String.Join(","c, StoreCultures))
        '"gd", "ig", "ky", "tk", "tt", "ug", "yo" cannot be neutral
        'Dim AllCultures() As Globalization.CultureInfo = Globalization.CultureInfo.GetCultures(Globalization.CultureTypes.InstalledWin32Cultures)
        'For Each File In New IO.DirectoryInfo("..\..\..\XMLRender\").EnumerateFiles("*.*.resx")
        '    Dim Count As Integer
        '    For Count = 0 To StoreCultures.Length - 1
        '        If "Resources." + StoreCultures(Count) + ".resx" = File.Name Or "resources." + StoreCultures(Count) + ".resx" = File.Name.ToLower() Then
        '            Exit For
        '        End If
        '    Next
        '    If Count = StoreCultures.Length Then
        '        Debug.Print("Not Found " + File.Name)
        '    End If
        'Next
        'For Count = 0 To StoreCultures.Length - 1
        '    Dim bFound As Boolean = False
        '    For Each File As IO.FileInfo In New IO.DirectoryInfo("..\..\..\XMLRender\").EnumerateFiles("*.*.resx")
        '        If "Resources." + StoreCultures(Count) + ".resx" = File.Name Or "resources." + StoreCultures(Count) + ".resx" = File.Name.ToLower() Then
        '            bFound = True
        '            Exit For
        '        End If
        '    Next
        '    If Not bFound And StoreCultures(Count) <> String.Empty And StoreCultures(Count) <> "en" Then
        '        Debug.Print("Not Found " + StoreCultures(Count))
        '        StoreCultures(Count) = String.Empty
        '    End If
        'Next
        'Dim StoreCultureNames(StoreCultures.Length - 1) As String
        'For Count = 0 To StoreCultures.Length - 1
        '    'frp and frc not in culture info yet and zh aliases zh-Hant and tk aliases both tk-latn and tk-latn-tr but not writing alias detect system yet
        '    If StoreCultures(Count).StartsWith("frp-") Or StoreCultures(Count).StartsWith("frc-") Or StoreCultures(Count) = "zh" Or StoreCultures(Count) = "tk-latn-tr" Then
        '        StoreCultures(Count) = String.Empty
        '        Continue For
        '    End If
        '    If StoreCultures(Count) = "pa-arab" Or (StoreCultures(Count) <> String.Empty And (Globalization.CultureInfo.GetCultureInfo(StoreCultures(Count)).Parent.Name = String.Empty Or Array.IndexOf(StoreCultures, Globalization.CultureInfo.GetCultureInfo(StoreCultures(Count)).Parent.Name.ToLower()) = -1 And Array.IndexOf(StoreCultures, Globalization.CultureInfo.GetCultureInfo(StoreCultures(Count)).Parent.Name) = -1 And (Globalization.CultureInfo.GetCultureInfo(StoreCultures(Count)).Parent.Parent.Name = String.Empty Or Array.IndexOf(StoreCultures, Globalization.CultureInfo.GetCultureInfo(StoreCultures(Count)).Parent.Parent.Name.ToLower()) = -1 And Array.IndexOf(StoreCultures, Globalization.CultureInfo.GetCultureInfo(StoreCultures(Count)).Parent.Parent.Name) = -1))) Then
        '        StoreCultureNames(Count) = If(Globalization.CultureInfo.GetCultureInfo(StoreCultures(Count)).DisplayName.StartsWith("Unknown"), Globalization.CultureInfo.GetCultureInfo(StoreCultures(Count)).Parent.DisplayName, Globalization.CultureInfo.GetCultureInfo(StoreCultures(Count)).DisplayName) + " [" + StoreCultures(Count) + "]"
        '    Else
        '        StoreCultures(Count) = String.Empty
        '    End If
        'Next
        'Array.Sort(StoreCultureNames)
        'For Count = 0 To StoreCultureNames.Length - 1
        '    If StoreCultureNames(Count) <> String.Empty Then Debug.Print(StoreCultureNames(Count))
        'Next
        'TanzilReader.WordFileToResource("..\..\..\metadata\id.w4w.terjemah.txt", "..\..\..\ResourceToolkitTempForeign\QuranResources.resx")
        'TanzilReader.ResourceToWordFile("..\..\..\ResourceToolkitTempForeign\QuranResources.ms.resx", "..\..\..\metadata\ms.w4w.terjemah.txt")
        'TanzilReader.WordFileToResource("..\..\..\metadata\en.w4w.corpus.txt", "..\..\..\ResourceToolkitTemp\QuranResources.resx")
        'TanzilReader.ResourceToWordFile("..\..\..\ResourceToolkitTemp\QuranResources.he.resx", "..\..\..\metadata\he.w4w.corpus.txt")
        'Utility.WordFileToResource("..\..\..\metadata\en.w4w.txt", "..\..\..\ResourceToolkit\W4WResources.resx")
        'Utility.ResourceToWordFile("..\..\..\ResourceToolkit\W4WResources.he.resx", "..\..\..\metadata\he.w4w.txt")
        'clsWarshQuran.ParseQuran()
        'TanzilReader.ChangeQuranFormat(TanzilReader.QuranTexts.Warsh, TanzilReader.QuranTexts.Warsh, TanzilReader.QuranScripts.Uthmani, TanzilReader.ArabicPresentation.Buckwalter)
        'TanzilReader.ChangeQuranFormat(TanzilReader.QuranTexts.Hafs, TanzilReader.QuranTexts.Hafs, TanzilReader.QuranScripts.Uthmani, TanzilReader.ArabicPresentation.Buckwalter)
        'TanzilReader.ChangeQuranFormat(TanzilReader.QuranTexts.Hafs, TanzilReader.QuranTexts.Hafs, TanzilReader.QuranScripts.UthmaniMin, TanzilReader.ArabicPresentation.None)
        'TanzilReader.ChangeQuranFormat(TanzilReader.QuranTexts.Hafs, TanzilReader.QuranTexts.Hafs, TanzilReader.QuranScripts.SimpleEnhanced, TanzilReader.QuranScripts.Uthmani, TanzilReader.ArabicPresentation.Buckwalter)
        'TanzilReader.ChangeQuranFormat(TanzilReader.QuranTexts.Hafs, TanzilReader.QuranTexts.Hafs, TanzilReader.QuranScripts.SimpleEnhanced, TanzilReader.ArabicPresentation.Buckwalter)
        'TanzilReader.ChangeQuranFormat(TanzilReader.QuranTexts.Hafs, TanzilReader.QuranTexts.Hafs, TanzilReader.QuranScripts.Simple, TanzilReader.ArabicPresentation.Buckwalter)
        'TanzilReader.ChangeQuranFormat(TanzilReader.QuranTexts.Hafs, TanzilReader.QuranTexts.Hafs, TanzilReader.QuranScripts.SimpleClean, TanzilReader.ArabicPresentation.Buckwalter)
        'TanzilReader.ChangeQuranFormat(TanzilReader.QuranTexts.Hafs, TanzilReader.QuranTexts.Hafs, TanzilReader.QuranScripts.SimpleMin, TanzilReader.ArabicPresentation.Buckwalter)
        'TanzilReader.ChangeQuranFormat(TanzilReader.QuranTexts.Hafs, TanzilReader.QuranTexts.Hafs, TanzilReader.QuranScripts.SimpleEnhanced, TanzilReader.ArabicPresentation.None)
        'TanzilReader.ChangeQuranFormat(TanzilReader.QuranTexts.Hafs, TanzilReader.QuranTexts.Hafs, TanzilReader.QuranScripts.Simple, TanzilReader.ArabicPresentation.None)
        'TanzilReader.ChangeQuranFormat(TanzilReader.QuranTexts.Hafs, TanzilReader.QuranTexts.Hafs, TanzilReader.QuranScripts.SimpleMin, TanzilReader.ArabicPresentation.None)
        'TanzilReader.ChangeQuranFormat(TanzilReader.QuranTexts.Hafs, TanzilReader.QuranTexts.Hafs, TanzilReader.QuranScripts.SimpleClean, TanzilReader.ArabicPresentation.None)
        'TanzilReader.CompareQuranFormats(TanzilReader.QuranTexts.Hafs, TanzilReader.QuranTexts.Warsh, TanzilReader.QuranScripts.Uthmani, TanzilReader.ArabicPresentation.Buckwalter)
        'TanzilReader.ChangeQuranFormat(TanzilReader.QuranTexts.Hafs, TanzilReader.QuranTexts.Warsh, TanzilReader.QuranScripts.Uthmani, TanzilReader.ArabicPresentation.Buckwalter)
        'TanzilReader.CheckNotablePatterns()
        'TanzilReader.FindMinimalVersesForCoverage()
        'TanzilReader.CheckSequentialRules()
        'TanzilReader.CheckMutualExclusiveRules(False, 4)
        'TanzilReader.CheckMutualExclusiveRules(True, 4)
        Dim IndexToVerse As Integer()() = Nothing
        'Dim Text As String = TanzilReader.QuranTextCombiner(CachedData.XMLDocMain, IndexToVerse, 18, 13)
        'Dim IndVerses As String() = TanzilReader.QuranTextRangeLookup(2, 5, 0, 2, 0, 0)(0)
        Arabic.TransliterateWithRulesColor(Arabic.TransliterateFromBuckwalter("wa{lomurosala`ti EurofFA"), "PlainRoman", True, False, Arabic.GetMetarules(Arabic.TransliterateFromBuckwalter("wa{lomurosala`ti EurofFA"), Nothing, CachedData.RuleMetas("UthmaniQuran")))
        Text = TanzilReader.GetQuranText(CachedData.XMLDocMain, 19, 4, 4)(0) + " " + ArabicData.ArabicEndOfAyah
        'Debug.Print(Arabic.TransliterateToScheme(IndVerses(0), ArabicData.TranslitScheme.RuleBased, String.Empty, CachedData.RuleMetas("UthmaniQuran"), TanzilReader.GenerateDefaultStops(IndVerses(0))))
        Debug.Print(Arabic.TransliterateToScheme(Text, ArabicData.TranslitScheme.RuleBased, String.Empty, Arabic.GetMetarules(Text, TanzilReader.GenerateDefaultStops(Text), CachedData.RuleMetas("UthmaniQuran"))))
        For Selection = 33 To TanzilReader.GetChapterCount()
            TanzilReader.GetRenderedQuranText(ArabicData.TranslitScheme.RuleBased, String.Empty, String.Empty, "0", Selection.ToString(), String.Empty, "0", "1")
            Debug.Print(CStr(Selection))
        Next
        'Debug.Print(Arabic.TransliterateToScheme(System.Text.RegularExpressions.Regex.Match(Text, CachedData.GetPattern("ContinuousMatch")).Value, ArabicData.TranslitScheme.Literal, String.Empty))
        'Utility.SortResX(Utility.GetFilePath("IslamResources\Resources.en.resx"))
        'Utility.SortResX(Utility.GetFilePath("IslamResources\My Project\Resources.resx"))
        CachedData.DoErrorCheck()
        'Dim RenderArr As RenderArray = Phrases.DoGetRenderedCatText(String.Empty, ArabicData.TranslitScheme.RuleBased, String.Empty, CachedData.IslamData.Phrases, 0)
        'HostPageUtility.RenderArray.OutputPdf("test.pdf", RenderArr.Items)
        For Count As Integer = 39 To TanzilReader.GetChapterCount()
            Dim RA As RenderArray = TanzilReader.GetQuranTextBySelection(String.Empty, 0, Count, String.Empty, ArabicData.TranslitScheme.RuleBased, String.Empty, TanzilReader.GetTranslationIndex(String.Empty), True, False, False, True, False, False, True)
            HostPageUtility.RenderArrayWeb.OutputPdf("test" + CStr(Count) + ".pdf", RA.Items)
        Next
    End Sub
    Private PageSet As PageLoader
    Private Sub frmMain_Load(sender As Object, e As EventArgs) Handles Me.Load
        PortableMethods.FileIO = New WindowsWebFileIO
        PortableMethods.Settings = New WindowsSettings
        PageSet = New PageLoader
        UtilityTestCode()
        For Index = 0 To PageSet.Pages.Count - 1
            Dim newNode As TreeNode = tvwMain.Nodes.Add(PageSet.Pages.Item(Index).PageName, Utility.LoadResourceString(PageSet.Pages.Item(Index).Text))
            For SubIndex = 0 To PageSet.Pages.Item(Index).Page.Count - 1
                If PageLoader.IsListItem(PageSet.Pages.Item(Index).Page(SubIndex)) AndAlso DirectCast(PageSet.Pages.Item(Index).Page(SubIndex), PageLoader.ListItem).IsSection Then
                    newNode.Nodes.Add(PageSet.Pages.Item(Index).PageName + CStr(If(DirectCast(PageSet.Pages.Item(Index).Page.Item(SubIndex), PageLoader.ListItem).Name <> String.Empty, "#" + DirectCast(PageSet.Pages.Item(Index).Page.Item(SubIndex), PageLoader.ListItem).Name, String.Empty)), Utility.LoadResourceString(DirectCast(PageSet.Pages.Item(Index).Page.Item(SubIndex), PageLoader.ListItem).Title))
                End If
            Next
        Next
        Dim Renderer As New MultiLangRender
        Renderer.Anchor = AnchorStyles.Left Or AnchorStyles.Top Or AnchorStyles.Right Or AnchorStyles.Bottom
        Renderer.Size = New Size(gbMain.Width, gbMain.Height)
        Renderer.RenderArray = TanzilReader.GetQuranTextBySelection(String.Empty, 0, 1, String.Empty, ArabicData.TranslitScheme.RuleBased, String.Empty, TanzilReader.GetTranslationIndex(String.Empty), True, False, False, True, False, False, True).Items
        gbMain.Controls.Add(Renderer)
    End Sub

    Private Sub gbMain_Resize(sender As Object, e As EventArgs) Handles gbMain.Resize
        If gbMain.Controls.Count <> 0 Then
            gbMain.Controls(0).Size = New Size(gbMain.Width, gbMain.Height)
        End If
    End Sub
End Class
