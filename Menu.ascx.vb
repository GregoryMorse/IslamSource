Partial Class Menu
    Inherits System.Web.UI.UserControl
    Dim PageSet As PageLoader
    Dim CurIndex As Integer
#Region " Web Form Designer Generated Code "

    'This call is required by the Web Form Designer.
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()

    End Sub

    'NOTE: The following placeholder declaration is required by the Web Form Designer.
    'Do not delete or move it.
    Private designerPlaceholderDeclaration As System.Object

    Private Sub Page_Init(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Init
        'CODEGEN: This method call is required by the Web Form Designer
        'Do not modify it using the code editor.
        InitializeComponent()
    End Sub

#End Region
    Public Sub New(ByVal NewPageSet As PageLoader, ByVal NewCurIndex As Integer)
        PageSet = NewPageSet
        CurIndex = NewCurIndex
    End Sub

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'Controls.Add(New WebControls.HyperLink())
    End Sub

    Protected Overrides Sub Render(ByVal writer As System.Web.UI.HtmlTextWriter)
        Dim Index As Integer
        Dim SubIndex As Integer
        Dim Scale As Double
        Dim SizeF As System.Drawing.SizeF
        'no spacing since inline-block element
        writer.WriteBeginTag("div")
        writer.WriteAttribute("id", "menudiv")
        writer.Write(HtmlTextWriter.TagRightChar)
        writer.WriteBeginTag("form")
        writer.WriteAttribute("action", Request.Url.AbsolutePath)
        writer.WriteAttribute("method", "get")
        writer.Write(HtmlTextWriter.TagRightChar)
        writer.WriteFullBeginTag("div")
        For Each Item As String In Request.QueryString.AllKeys
            writer.WriteBeginTag("input")
            writer.WriteAttribute("type", "hidden")
            writer.WriteAttribute("name", Item)
            writer.WriteAttribute("value", Request.QueryString(Item))
            writer.Write(HtmlTextWriter.TagRightChar)
        Next
        writer.WriteBeginTag("select")
        writer.WriteAttribute("name", host.LangSet)
        writer.WriteAttribute("id", host.LangSet)
        writer.WriteAttribute("onchange", "javascript: this.parentElement.parentElement.submit();")
        writer.Write(HtmlTextWriter.TagRightChar)
        IslamData.Initialize()
        Dim AllCultures() As Globalization.CultureInfo = Globalization.CultureInfo.GetCultures(Globalization.CultureTypes.AllCultures)
        For Count As Integer = 0 To AllCultures.Length - 1
            If IO.File.Exists(Utility.GetFilePath(VirtualPathUtility.ToAbsolute("~/App_LocalResources/host.aspx." + AllCultures(Count).Name + ".resx"))) Then
                writer.WriteBeginTag("option")
                writer.WriteAttribute("value", AllCultures(Count).Name)
                If AllCultures(Count).Name = Globalization.CultureInfo.CurrentCulture.Name Or _
                    AllCultures(Count).Name = Globalization.CultureInfo.CurrentCulture.TwoLetterISOLanguageName Then writer.WriteAttribute("selected", "selected")
                writer.Write(HtmlTextWriter.TagRightChar + Utility.HtmlTextEncode(AllCultures(Count).NativeName))
                writer.WriteEndTag("option")
            End If
        Next
        writer.WriteEndTag("select")
        writer.WriteEndTag("div")
        writer.WriteEndTag("form")
        writer.WriteBeginTag("a")
        writer.WriteAttribute("href", "/")
        'cannot use jquery inside on... functions
        writer.WriteAttribute("onmouseover", Utility.HtmlTextEncode("javascript: document.getElementById('mainimage').src = '" + host.GetPageString("Image.gif&Image=Scale&p=menu.hover") + "';"))
        writer.WriteAttribute("onmouseout", Utility.HtmlTextEncode("javascript: document.getElementById('mainimage').src = '" + host.GetPageString("Image.gif&Image=Scale&p=menu.main") + "';"))
        writer.Write(HtmlTextWriter.TagRightChar)
        writer.WriteBeginTag("img")
        writer.WriteAttribute("id", "mainimage")
        writer.WriteAttribute("alt", Utility.LoadResourceString(PageSet.Title))
        writer.WriteAttribute("src", Utility.HtmlTextEncode(host.GetPageString("Image.gif&Image=Scale&p=menu.main")))
        SizeF = Utility.GetImageDimensions(Request.PhysicalApplicationPath + "images/" + PageSet.MainImage)
        Scale = Utility.ComputeImageScale(SizeF.Width, SizeF.Height, 226 - 10 - 10, 200)
        writer.WriteAttribute("width", Convert.ToInt32(SizeF.Width / Scale).ToString())
        writer.WriteAttribute("height", Convert.ToInt32(SizeF.Height / Scale).ToString())
        writer.Write(HtmlTextWriter.TagRightChar)
        writer.WriteEndTag("a")
        UserAccounts.AccountPanel(writer)
        For Index = 0 To PageSet.Pages.Count - 1
            writer.Write(vbCrLf + vbTab + vbTab + vbTab + vbTab)
            If CurIndex = Index Then
                writer.WriteBeginTag("div")
                writer.WriteAttribute("class", "menusel")
                writer.Write(HtmlTextWriter.TagRightChar)
            Else
                writer.WriteFullBeginTag("div")
            End If
            writer.Write(vbCrLf + vbTab + vbTab + vbTab + vbTab + vbTab)
            writer.WriteBeginTag("a")
            writer.WriteAttribute("href", Utility.HtmlTextEncode(host.GetPageString(PageSet.Pages.Item(Index).PageName)))
            writer.Write(HtmlTextWriter.TagRightChar)
            writer.Write(vbCrLf + vbTab + vbTab + vbTab + vbTab + vbTab + vbTab)
            writer.Write(Utility.HtmlTextEncode(CStr(IIf(CurIndex = Index, "-> ", String.Empty)) + Utility.LoadResourceString(PageSet.Pages.Item(Index).Text)) + Utility.HtmlTextEncode(CStr(IIf(CurIndex = Index, " <-", String.Empty))))
            writer.Write(vbCrLf + vbTab + vbTab + vbTab + vbTab + vbTab)
            writer.WriteEndTag("a")
            writer.Write(vbCrLf + vbTab + vbTab + vbTab + vbTab)
            writer.WriteEndTag("div")
            If (CurIndex = Index) Then
                For SubIndex = 0 To PageSet.Pages.Item(Index).Page.Count - 1
                    If PageLoader.IsListItem(PageSet.Pages.Item(Index).Page(SubIndex)) AndAlso DirectCast(PageSet.Pages.Item(Index).Page(SubIndex), PageLoader.ListItem).IsSection Then
                        writer.Write(vbCrLf + vbTab + vbTab + vbTab + vbTab)
                        writer.WriteFullBeginTag("div")
                        writer.Write(Utility.ConvertSpaces(HtmlTextWriter.SpaceChar + HtmlTextWriter.SpaceChar + HtmlTextWriter.SpaceChar) + vbCrLf + vbTab + vbTab + vbTab + vbTab + vbTab)
                        writer.WriteBeginTag("a")
                        writer.WriteAttribute("href", Utility.HtmlTextEncode(host.GetPageString(PageSet.Pages.Item(Index).PageName) + CStr(IIf(DirectCast(PageSet.Pages.Item(Index).Page.Item(SubIndex), PageLoader.ListItem).Name <> String.Empty, "#" + DirectCast(PageSet.Pages.Item(Index).Page.Item(SubIndex), PageLoader.ListItem).Name, String.Empty))))
                        writer.Write(HtmlTextWriter.TagRightChar)
                        writer.Write(vbCrLf + vbTab + vbTab + vbTab + vbTab + vbTab + vbTab)
                        writer.Write(Utility.HtmlTextEncode(Utility.LoadResourceString(DirectCast(PageSet.Pages.Item(Index).Page.Item(SubIndex), PageLoader.ListItem).Title)))
                        writer.Write(vbCrLf + vbTab + vbTab + vbTab + vbTab + vbTab)
                        writer.WriteEndTag("a")
                        writer.Write(vbCrLf + vbTab + vbTab + vbTab + vbTab)
                        writer.WriteEndTag("div")
                    End If
                Next
            End If
            If (Index <> PageSet.Pages.Count - 1) Then
                writer.Write(vbCrLf + vbTab + vbTab + vbTab + vbTab)
                writer.WriteFullBeginTag("br")
            End If
        Next
        writer.Write(vbCrLf + vbTab + vbTab + vbTab)
        writer.WriteEndTag("div")
    End Sub
End Class
