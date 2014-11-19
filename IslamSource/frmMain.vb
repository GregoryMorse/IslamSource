Imports HostPageUtility
Imports IslamMetadata

Public Class frmMain
    Private PageSet As New PageLoader
    Private Sub frmMain_Load(sender As Object, e As EventArgs) Handles Me.Load
        'clsWarshQuran.ParseQuran()
        'TanzilReader.ChangeQuranFormat(TanzilReader.QuranTexts.Warsh, TanzilReader.QuranTexts.Warsh, TanzilReader.QuranScripts.Uthmani, TanzilReader.ArabicPresentation.Buckwalter)
        'TanzilReader.ChangeQuranFormat(TanzilReader.QuranTexts.Hafs, TanzilReader.QuranTexts.Hafs, TanzilReader.QuranScripts.Uthmani, TanzilReader.ArabicPresentation.Buckwalter)
        'TanzilReader.CompareQuranFormats(TanzilReader.QuranTexts.Hafs, TanzilReader.QuranTexts.Warsh, TanzilReader.QuranScripts.Uthmani, TanzilReader.ArabicPresentation.Buckwalter)
        'TanzilReader.ChangeQuranFormat(TanzilReader.QuranTexts.Hafs, TanzilReader.QuranTexts.Warsh, TanzilReader.QuranScripts.Uthmani, TanzilReader.ArabicPresentation.Buckwalter)
        'TanzilReader.CheckNotablePatterns()
        'CachedData.DoErrorCheck()
        Dim RenderArr As RenderArray = Supplications.DoGetRenderedSuppText(ArabicData.TranslitScheme.RuleBased, "PlainRoman", CachedData.IslamData.VerseCategories(0), 0)
        HostPageUtility.RenderArray.OutputPdf("test.pdf", RenderArr.Items)
        For Index = 0 To PageSet.Pages.Count - 1
            Dim newNode As TreeNode = tvwMain.Nodes.Add(PageSet.Pages.Item(Index).PageName, Utility.LoadResourceString(PageSet.Pages.Item(Index).Text))
            For SubIndex = 0 To PageSet.Pages.Item(Index).Page.Count - 1
                If PageLoader.IsListItem(PageSet.Pages.Item(Index).Page(SubIndex)) AndAlso DirectCast(PageSet.Pages.Item(Index).Page(SubIndex), PageLoader.ListItem).IsSection Then
                    newNode.Nodes.Add(PageSet.Pages.Item(Index).PageName + CStr(IIf(DirectCast(PageSet.Pages.Item(Index).Page.Item(SubIndex), PageLoader.ListItem).Name <> String.Empty, "#" + DirectCast(PageSet.Pages.Item(Index).Page.Item(SubIndex), PageLoader.ListItem).Name, String.Empty)), Utility.LoadResourceString(DirectCast(PageSet.Pages.Item(Index).Page.Item(SubIndex), PageLoader.ListItem).Title))
                End If
            Next
        Next
        Dim Renderer As New MultiLangRender
        Renderer.Anchor = AnchorStyles.Left Or AnchorStyles.Top Or AnchorStyles.Right Or AnchorStyles.Bottom
        Renderer.Size = New Size(gbMain.Width, gbMain.Height)
        For Count As Integer = 114 To TanzilReader.GetChapterCount()
            Renderer.RenderArray = TanzilReader.GetQuranTextBySelection(0, Count, String.Empty, ArabicData.TranslitScheme.RuleBased, "PlainRoman", 0).Items
            HostPageUtility.RenderArray.OutputPdf("test" + CStr(Count) + ".pdf", Renderer.RenderArray)
        Next
        Renderer.RenderArray = TanzilReader.GetQuranTextBySelection(0, 1, String.Empty, ArabicData.TranslitScheme.RuleBased, "PlainRoman", 0).Items
        gbMain.Controls.Add(Renderer)
    End Sub

    Private Sub gbMain_Resize(sender As Object, e As EventArgs) Handles gbMain.Resize
        If gbMain.Controls.Count <> 0 Then
            gbMain.Controls(0).Size = New Size(gbMain.Width, gbMain.Height)
        End If
    End Sub
End Class
