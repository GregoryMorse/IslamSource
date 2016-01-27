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
        <Runtime.InteropServices.DllImport("getuname.dll", EntryPoint:="GetUName")> _
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
    End Class
    Private PageSet As PageLoader
    Private Sub frmMain_Load(sender As Object, e As EventArgs) Handles Me.Load
        PortableMethods.FileIO = New WindowsWebFileIO
        PortableMethods.Settings = New WindowsSettings
        PageSet = New PageLoader
        'TanzilReader.WordFileToResource("..\..\..\metadata\en.w4w.corpus.txt", "Resources.resx")
        TanzilReader.ResourceToWordFile("..\..\My Project\Resources.hu.resx", "..\..\..\metadata\hu.w4w.corpus.txt")
        'clsWarshQuran.ParseQuran()
        'TanzilReader.ChangeQuranFormat(TanzilReader.QuranTexts.Warsh, TanzilReader.QuranTexts.Warsh, TanzilReader.QuranScripts.Uthmani, TanzilReader.ArabicPresentation.Buckwalter)
        'TanzilReader.ChangeQuranFormat(TanzilReader.QuranTexts.Hafs, TanzilReader.QuranTexts.Hafs, TanzilReader.QuranScripts.Uthmani, TanzilReader.ArabicPresentation.Buckwalter)
        'TanzilReader.ChangeQuranFormat(TanzilReader.QuranTexts.Hafs, TanzilReader.QuranTexts.Hafs, TanzilReader.QuranScripts.UthmaniMin, TanzilReader.ArabicPresentation.None)
        TanzilReader.ChangeQuranFormat(TanzilReader.QuranTexts.Hafs, TanzilReader.QuranTexts.Hafs, TanzilReader.QuranScripts.SimpleEnhanced, TanzilReader.QuranScripts.Uthmani, TanzilReader.ArabicPresentation.Buckwalter)
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
        TanzilReader.CheckMutualExclusiveRules(False)
        TanzilReader.CheckMutualExclusiveRules(True)
        Dim IndexToVerse As Integer()() = Nothing
        Dim Text As String = TanzilReader.QuranTextCombiner(CachedData.XMLDocMain, IndexToVerse, 2, 2)
        Dim IndVerses As String() = TanzilReader.QuranTextRangeLookup(2, 5, 0, 2, 0, 0)(0)
        Debug.Print(Arabic.TransliterateToScheme(IndVerses(0), ArabicData.TranslitScheme.RuleBased, "PlainRoman", CachedData.RuleMetas("UthmaniQuran"), TanzilReader.GenerateDefaultStops(IndVerses(0))))
        Debug.Print(Arabic.TransliterateToScheme(Text, ArabicData.TranslitScheme.RuleBased, "PlainRoman", CachedData.RuleMetas("UthmaniQuran"), TanzilReader.GenerateDefaultStops(Text)))
        'Debug.Print(Arabic.TransliterateToScheme(System.Text.RegularExpressions.Regex.Match(Text, CachedData.GetPattern("ContinuousMatch")).Value, ArabicData.TranslitScheme.Literal, String.Empty))
        Arabic.TransliterateToScheme(Arabic.TransliterateFromBuckwalter("suwrapu {lofaAtiHapi"), ArabicData.TranslitScheme.LearningMode, "PlainRoman", CachedData.RuleMetas("UthmaniQuran"), TanzilReader.GenerateDefaultStops(Arabic.TransliterateFromBuckwalter("suwrapu {lofaAtiHapi")))
        Arabic.TransliterateToScheme(Arabic.TransliterateFromBuckwalter("*ikorFA"), ArabicData.TranslitScheme.LearningMode, "PlainRoman", CachedData.RuleMetas("UthmaniQuran"), TanzilReader.GenerateDefaultStops(Arabic.TransliterateFromBuckwalter("*ikorFA")))
        Arabic.TransliterateToScheme(Arabic.TransliterateFromBuckwalter("bihi."), ArabicData.TranslitScheme.LearningMode, "PlainRoman", CachedData.RuleMetas("UthmaniQuran"), TanzilReader.GenerateDefaultStops(Arabic.TransliterateFromBuckwalter("bihi.")))
        'Utility.SortResX(Utility.GetFilePath("IslamResources\Resources.en.resx"))
        'Utility.SortResX(Utility.GetFilePath("IslamResources\My Project\Resources.resx"))
        CachedData.DoErrorCheck()
        'Dim RenderArr As RenderArray = Phrases.DoGetRenderedCatText(String.Empty, ArabicData.TranslitScheme.RuleBased, "PlainRoman", CachedData.IslamData.Phrases, 0)
        'HostPageUtility.RenderArray.OutputPdf("test.pdf", RenderArr.Items)
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
        For Count As Integer = 39 To TanzilReader.GetChapterCount()
            Renderer.RenderArray = TanzilReader.GetQuranTextBySelection(String.Empty, 0, Count, "en.sahih", ArabicData.TranslitScheme.RuleBased, "PlainRoman", TanzilReader.GetTranslationIndex("en.sahih"), True, False, False, True, False, False, True).Items
            HostPageUtility.RenderArrayWeb.OutputPdf("test" + CStr(Count) + ".pdf", Renderer.RenderArray)
        Next
        Renderer.RenderArray = TanzilReader.GetQuranTextBySelection(String.Empty, 0, 1, "en.sahih", ArabicData.TranslitScheme.RuleBased, "PlainRoman", TanzilReader.GetTranslationIndex("en.sahih"), True, False, False, True, False, False, True).Items
        gbMain.Controls.Add(Renderer)
    End Sub

    Private Sub gbMain_Resize(sender As Object, e As EventArgs) Handles gbMain.Resize
        If gbMain.Controls.Count <> 0 Then
            gbMain.Controls(0).Size = New Size(gbMain.Width, gbMain.Height)
        End If
    End Sub
End Class
