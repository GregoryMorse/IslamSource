Imports IslamMetadata

Public Class frmMain
    Private PageSet As New PageLoader
    Private Sub frmMain_Load(sender As Object, e As EventArgs) Handles Me.Load
        For Index = 0 To PageSet.Pages.Count - 1
            Dim newNode As TreeNode = tvwMain.Nodes.Add(PageSet.Pages.Item(Index).PageName, Utility.LoadResourceString(PageSet.Pages.Item(Index).Text))
            For SubIndex = 0 To PageSet.Pages.Item(Index).Page.Count - 1
                If PageLoader.IsListItem(PageSet.Pages.Item(Index).Page(SubIndex)) AndAlso DirectCast(PageSet.Pages.Item(Index).Page(SubIndex), PageLoader.ListItem).IsSection Then
                    newNode.Nodes.Add(PageSet.Pages.Item(Index).PageName + CStr(IIf(DirectCast(PageSet.Pages.Item(Index).Page.Item(SubIndex), PageLoader.ListItem).Name <> String.Empty, "#" + DirectCast(PageSet.Pages.Item(Index).Page.Item(SubIndex), PageLoader.ListItem).Name, String.Empty)), Utility.LoadResourceString(DirectCast(PageSet.Pages.Item(Index).Page.Item(SubIndex), PageLoader.ListItem).Title))
                End If
            Next
        Next

        Dim Renderer As New MultiLangRender
        Renderer.MaximumSize = New Size(gbMain.Width, gbMain.Height)
        Renderer.RenderArray = TanzilReader.DoGetRenderedQuranText(0, 1, String.Empty, 0, 0).Items
        Renderer.Anchor = AnchorStyles.Left Or AnchorStyles.Top Or AnchorStyles.Right Or AnchorStyles.Bottom
        gbMain.Controls.Add(Renderer)
    End Sub

    Private Sub gbMain_Resize(sender As Object, e As EventArgs) Handles gbMain.Resize
        If gbMain.Controls.Count <> 0 Then
            gbMain.Controls(0).MaximumSize = New Size(gbMain.Width, gbMain.Height)
            gbMain.Controls(0).Size = New Size(0, 0)
        End If
    End Sub
End Class
