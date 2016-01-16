Imports HostPageUtility
Partial Class Page
    Inherits System.Web.UI.UserControl
    Dim MyPage As PageLoader.PageItem
    Dim UsePanes As Boolean
    Dim UsePrint As Boolean
    Public JSFunctions As New Collections.Generic.List(Of String)
    Public JSInitFuncs As New Collections.Generic.List(Of String)

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
    Public Sub New(ByVal NewPage As PageLoader.PageItem, Optional ByVal NewUsePanes As Boolean = True, Optional ByVal NewIsPrint As Boolean = False)
        MyPage = NewPage
        UsePanes = NewUsePanes
        UsePrint = NewIsPrint
    End Sub
    Function GetTitle() As String
        Return Utility.LoadResourceString(MyPage.Text)
    End Function
    Function GetDescription() As String
        If MyPage.Page.Count <> 0 Then
            If PageLoader.IsListItem(MyPage.Page(0)) Then
                Return Utility.LoadResourceString(CType(MyPage.Page(0), PageLoader.ListItem).Title)
            ElseIf PageLoader.IsTextItem(MyPage.Page(0)) Then
                Return Utility.LoadResourceString(CType(MyPage.Page(0), PageLoader.TextItem).Text)
            End If
        End If
        Return GetTitle()
    End Function
    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        'Put user code to initialize the page here
    End Sub
    Sub AddToJSFunctions(ByVal NewJSFunctions() As String)
        Dim Count As Integer
        If NewJSFunctions.Length > 2 Then
            If Not JSInitFuncs.Contains(NewJSFunctions(1)) Then
                JSInitFuncs.Add(NewJSFunctions(1))
            End If
        End If
        For Count = 2 To NewJSFunctions.Length - 1
            If Not JSFunctions.Contains(NewJSFunctions(Count)) Then
                JSFunctions.Add(NewJSFunctions(Count))
            End If
        Next
    End Sub
    Sub OutputStrings(ByVal writer As System.Web.UI.HtmlTextWriter, ByVal Strings As String())
        Dim Count As Integer
        For Count = 0 To Strings.Length - 1
            writer.Write(Strings(Count))
            If Count <> Strings.Length - 1 Then writer.WriteFullBeginTag("br")
        Next
    End Sub
    Sub WriteTextItem(ByVal writer As System.Web.UI.HtmlTextWriter, ByVal Item As PageLoader.TextItem, ByVal TabCount As Integer, ByVal IndexString As String)
        Dim BaseTabs As String = Utility.MakeTabString(TabCount)
        Dim Output As Object
        Dim Text As String
        writer.Write(vbCrLf + BaseTabs)
        writer.WriteBeginTag("div")
        If Item.Name = "font-tester" Then writer.WriteAttribute("style", "font-family: 'Sakkal Majalla', 'Times New Roman'; display: none;")
        writer.WriteAttribute("class", CStr(If(Item.Name = "font-tester", "font-tester", If(Item.Name = "Error", "error", "normal"))))
        If Item.Name <> String.Empty Then writer.WriteAttribute("id", Item.Name)
        writer.Write(HtmlTextWriter.TagRightChar)
        If (Item.URL <> String.Empty) Then
            writer.Write(vbCrLf + BaseTabs + vbTab)
            writer.WriteBeginTag("a")
            writer.WriteAttribute("href", HttpUtility.HtmlEncode(Item.URL))
            writer.Write(HtmlTextWriter.TagRightChar)
            writer.Write(vbCrLf + BaseTabs + vbTab + vbTab)
            If (Item.ImageURL <> String.Empty) Then
                Dim SizeF As Drawing.SizeF
                Text = "Image.gif&Image=Thumb&p=" + IndexString + "." + Item.Name
                SizeF = Utility.GetThumbSizeFromURL(Item.ImageURL, HttpContext.Current.Request.Url.Host + "_" + Text, 121)
                If SizeF.IsEmpty Then
                    SizeF.Width = 121
                    SizeF.Height = 121 * 3 / 4 + 1
                    Text = Item.ImageURL
                Else
                    Text = host.GetPageString(Text)
                End If
                writer.WriteBeginTag("img")
                writer.WriteAttribute("src", Utility.HtmlTextEncode(Text))
                writer.WriteAttribute("alt", Utility.HtmlTextEncode(Utility.LoadResourceString(Item.Text)))
                writer.WriteAttribute("width", "121")
                writer.WriteAttribute("height", CStr(121 * SizeF.Height / SizeF.Width))
                writer.Write(HtmlTextWriter.TagRightChar)
            End If
            BaseTabs += vbTab
        End If
        If Item.Text <> String.Empty Then
            writer.Write(vbCrLf + BaseTabs + vbTab)
            OutputStrings(writer, Utility.HtmlTextEncode(Utility.LoadResourceString(Item.Text)).Split(New String() {vbCrLf, vbLf}, StringSplitOptions.None))
            writer.Write("&nbsp;&nbsp;")
        End If
        If Not Item.OnRenderFunction Is Nothing Then
            If Item.Text <> String.Empty Then writer.Write("&nbsp;&nbsp;")
            Output = Item.OnRenderFunction.Invoke(Nothing, New Object() {Item})
            If TypeOf Output Is String Then
                OutputStrings(writer, Utility.HtmlTextEncode(CStr(Output)).Split(New String() {vbCrLf, vbLf}, StringSplitOptions.None))
            ElseIf TypeOf Output Is RenderArray Then
                DirectCast(Output, RenderArray).Render(writer, TabCount + 1)
                AddToJSFunctions(DirectCast(Output, RenderArray).GetRenderJS())
            Else
                RenderArray.WriteTable(writer, DirectCast(Output, Object()), TabCount + 1, Item.Name)
                For Each Strs As String() In DirectCast(Output, Object())
                    AddToJSFunctions(Strs)
                Next
            End If
        End If
        If (Item.URL <> String.Empty) Then
            BaseTabs = BaseTabs.Substring(0, BaseTabs.Length - 1)
            If Not Item.URL.StartsWith("/", StringComparison.Ordinal) Then writer.Write(" [" + New Uri(Item.URL).Host + "]")
            writer.Write(vbCrLf + BaseTabs + vbTab)
            writer.WriteEndTag("a")
        End If
        writer.Write(vbCrLf + BaseTabs)
        writer.WriteEndTag("div")
        'writer.Write(vbCrLf + BaseTabs)
        'writer.WriteFullBeginTag("br")
    End Sub
    Sub RenderSingleItem(ByVal writer As System.Web.UI.HtmlTextWriter, ByVal Item As Object, ByVal IndexString As String, ByVal TabCount As Integer)
        Dim SubIndex As Integer
        Dim Scale As Double
        Dim SizeF As System.Drawing.SizeF
        Dim BaseTabs As String = Utility.MakeTabString(TabCount)
        If (PageLoader.IsListItem(Item)) Then
            If (DirectCast(Item, PageLoader.ListItem).IsSection) Then
                writer.Write(vbCrLf + BaseTabs)
                writer.WriteBeginTag("a")
                writer.WriteAttribute("name", Utility.HtmlTextEncode(DirectCast(Item, PageLoader.ListItem).Name))
                writer.Write(HtmlTextWriter.TagRightChar)
                writer.WriteEndTag("a")
            End If
            'no spacing since inline-block element
            writer.WriteBeginTag("div")
            writer.WriteAttribute("class", "list")
            writer.Write(HtmlTextWriter.TagRightChar)
            writer.Write(vbCrLf + BaseTabs + vbTab)
            writer.WriteFullBeginTag("div")
            writer.Write(vbCrLf + BaseTabs + vbTab + vbTab)
            writer.WriteFullBeginTag("b")
            writer.WriteFullBeginTag("big")
            writer.Write(Utility.HtmlTextEncode(Utility.LoadResourceString(DirectCast(Item, PageLoader.ListItem).Title)))
            writer.WriteEndTag("big")
            writer.WriteEndTag("b")
            writer.Write(vbCrLf + BaseTabs + vbTab)
            writer.WriteEndTag("div")
            If Not UsePrint And DirectCast(Item, PageLoader.ListItem).HasForm Then
                writer.Write(vbCrLf + BaseTabs + vbTab)
                writer.WriteBeginTag("form")
                Dim bHasMultiEdit As Boolean = False
                For Count As Integer = 0 To DirectCast(Item, PageLoader.ListItem).List.Count - 1
                    'will not work if nested Edit
                    If PageLoader.IsEditItem(DirectCast(Item, PageLoader.ListItem).List(Count)) Then
                        If DirectCast(DirectCast(Item, PageLoader.ListItem).List(Count), PageLoader.EditItem).Rows > 1 Then
                            bHasMultiEdit = True
                            Exit For
                        End If
                    End If
                Next
                writer.WriteAttribute("action", CStr(If(DirectCast(Item, PageLoader.ListItem).FormPostURL <> String.Empty, DirectCast(Item, PageLoader.ListItem).FormPostURL, host.MainPage)))
                writer.WriteAttribute("method", CStr(If(DirectCast(Item, PageLoader.ListItem).FormPostURL <> String.Empty Or bHasMultiEdit, "post", "get")))
                writer.Write(HtmlTextWriter.TagRightChar)
                writer.Write(vbCrLf + BaseTabs + vbTab + vbTab)
                writer.WriteBeginTag("div")
                writer.Write(HtmlTextWriter.TagRightChar)
                writer.Write(vbCrLf + BaseTabs + vbTab + vbTab + vbTab)
                writer.WriteBeginTag("input")
                writer.WriteAttribute("type", "hidden")
                writer.WriteAttribute("name", "Page")
                writer.WriteAttribute("value", Utility.HtmlTextEncode(MyPage.PageName))
                writer.Write(HtmlTextWriter.TagRightChar)
                TabCount += 2
            End If
            For SubIndex = 0 To DirectCast(Item, PageLoader.ListItem).List.Count - 1
                RenderSingleItem(writer, DirectCast(Item, PageLoader.ListItem).List.Item(SubIndex), IndexString + "." + DirectCast(Item, PageLoader.ListItem).Name, TabCount + 1)
            Next
            If DirectCast(Item, PageLoader.ListItem).HasForm Then
                TabCount -= 2
                writer.Write(vbCrLf + BaseTabs + vbTab + vbTab)
                writer.WriteEndTag("div")
                writer.Write(vbCrLf + BaseTabs + vbTab)
                writer.WriteEndTag("form")
            End If
            writer.Write(vbCrLf + BaseTabs)
            writer.WriteEndTag("div")
            writer.Write(vbCrLf + BaseTabs)
            writer.WriteFullBeginTag("br")
        ElseIf (PageLoader.IsDownloadItem(Item)) Then
            Dim PathName() As String
            If Not DirectCast(Item, PageLoader.DownloadItem).OnRenderFunction Is Nothing Then
                PathName = CType(DirectCast(Item, PageLoader.DownloadItem).OnRenderFunction.Invoke(Nothing, Nothing), String())
            Else
                PathName = New String() {CStr(If(DirectCast(Item, PageLoader.DownloadItem).RelativePath, If(DirectCast(Item, PageLoader.DownloadItem).Path.EndsWith(".h") Or DirectCast(Item, PageLoader.DownloadItem).Path.EndsWith(".vb") Or DirectCast(Item, PageLoader.DownloadItem).Path.EndsWith(".vba") Or DirectCast(Item, PageLoader.DownloadItem).Path.EndsWith(".bat"), host.GetPageString("Source&File=" + HttpUtility.UrlEncode(DirectCast(Item, PageLoader.DownloadItem).Path)), "files/" + DirectCast(Item, PageLoader.DownloadItem).Path), DirectCast(Item, PageLoader.DownloadItem).Path)), _
                                 CStr(If(DirectCast(Item, PageLoader.DownloadItem).UseLink, Utility.LoadResourceString(DirectCast(Item, PageLoader.DownloadItem).Text), DirectCast(Item, PageLoader.DownloadItem).Path))}
            End If
            If PathName.Length = 2 Then
                If DirectCast(Item, PageLoader.DownloadItem).ShowInline Then
                    writer.WriteFullBeginTag("pre")
                    writer.Write(HttpUtility.HtmlEncode(Utility.SourceTextEncode(IO.File.ReadAllText(Utility.GetFilePath("files\" + DirectCast(Item, PageLoader.DownloadItem).Path)))))
                    writer.WriteEndTag("pre")
                End If
                writer.Write(vbCrLf + BaseTabs)
                writer.WriteBeginTag("a")
                writer.WriteAttribute("href", HttpUtility.HtmlEncode(PathName(0)))
                writer.Write(HtmlTextWriter.TagRightChar)
                writer.Write(vbCrLf + BaseTabs + vbTab + String.Format(Utility.LoadResourceString("Acct_Download"), PathName(1)))
                writer.Write(vbCrLf + BaseTabs)
                writer.WriteEndTag("a")
                writer.Write(vbCrLf + BaseTabs)
                writer.WriteFullBeginTag("br")
            End If
        ElseIf (PageLoader.IsEmailItem(Item)) Then
            writer.Write(vbCrLf + BaseTabs)
            If DirectCast(Item, PageLoader.EmailItem).UseImage Then
                writer.WriteBeginTag("img")
                writer.WriteAttribute("src", HttpUtility.HtmlEncode("host.aspx?Page=Image.gif&Image=EMailAddress"))
                writer.WriteAttribute("alt", Utility.HtmlTextEncode(Utility.LoadResourceString("Acct_EmailAddress") + ": " + Utility.ConnectionData.EMailAddress))
                writer.Write(HtmlTextWriter.TagRightChar)
                writer.Write(vbCrLf + BaseTabs)
                writer.WriteFullBeginTag("br")
            Else
                writer.WriteBeginTag("a")
                writer.WriteAttribute("href", HttpUtility.HtmlEncode("mailto: " + Utility.ConnectionData.EMailAddress))
                writer.Write(HtmlTextWriter.TagRightChar)
                writer.Write(vbCrLf + BaseTabs + vbTab + Utility.LoadResourceString("Acct_EmailAddress"))
                writer.Write(vbCrLf + BaseTabs)
                writer.WriteEndTag("a")
            End If
        ElseIf (PageLoader.IsImageItem(Item)) Then
            SizeF = Utility.GetImageDimensions(Utility.GetFilePath("images\" + DirectCast(Item, PageLoader.ImageItem).Path))
            Scale = Utility.ComputeImageScale(SizeF.Width, SizeF.Height, DirectCast(Item, PageLoader.ImageItem).MaxX, DirectCast(Item, PageLoader.ImageItem).MaxY)
            If DirectCast(Item, PageLoader.ImageItem).Link And Scale <> 1 Then
                writer.Write(vbCrLf + BaseTabs)
                writer.WriteBeginTag("a")
                writer.WriteAttribute("href", HttpUtility.HtmlEncode("images/" + DirectCast(Item, PageLoader.ImageItem).Path))
                writer.Write(HtmlTextWriter.TagRightChar)
                writer.Write(vbCrLf + BaseTabs + vbTab)
            Else
                writer.Write(vbCrLf + BaseTabs)
            End If
            writer.WriteBeginTag("img")
            writer.WriteAttribute("src", CStr(If(Scale = 1, "images/" + DirectCast(Item, PageLoader.ImageItem).Path, host.GetPageString("Image.gif&Image=Scale&p=" + IndexString + "." + DirectCast(Item, PageLoader.ImageItem).Name))))
            writer.WriteAttribute("alt", Utility.HtmlTextEncode(Utility.LoadResourceString(DirectCast(Item, PageLoader.ImageItem).Text)))
            writer.WriteAttribute("width", Convert.ToInt32(SizeF.Width / Scale).ToString())
            writer.WriteAttribute("height", Convert.ToInt32(SizeF.Height / Scale).ToString())
            If DirectCast(Item, PageLoader.ImageItem).Name = "fontloading" Then
                writer.WriteAttribute("id", DirectCast(Item, PageLoader.ImageItem).Name)
                writer.WriteAttribute("style", "display: none;")
            End If
            writer.Write(HtmlTextWriter.TagRightChar)
            If DirectCast(Item, PageLoader.ImageItem).Link And Scale <> 1 Then
                writer.Write(vbCrLf + BaseTabs)
                writer.WriteEndTag("a")
            End If
            writer.Write(vbCrLf + BaseTabs)
            writer.WriteFullBeginTag("br")
        ElseIf Not UsePrint And (PageLoader.IsEditItem(Item)) Then
            writer.Write(vbCrLf + BaseTabs)
            If (DirectCast(Item, PageLoader.EditItem).Rows > 1) Then
                writer.WriteBeginTag("textarea")
                writer.WriteAttribute("name", DirectCast(Item, PageLoader.EditItem).Name)
                writer.WriteAttribute("id", DirectCast(Item, PageLoader.EditItem).Name)
                writer.WriteAttribute("cols", "80")
                writer.WriteAttribute("rows", DirectCast(Item, PageLoader.EditItem).Rows.ToString())
                writer.Write(HtmlTextWriter.TagRightChar)
                If Web.HttpContext.Current.Request.Params.Get(DirectCast(Item, PageLoader.EditItem).Name) <> String.Empty Then
                    writer.Write(Web.HttpContext.Current.Request.Params.Get(DirectCast(Item, PageLoader.EditItem).Name))
                ElseIf DirectCast(Item, PageLoader.EditItem).DefaultValue <> String.Empty Then
                    writer.Write(Utility.LoadResourceString(DirectCast(Item, PageLoader.EditItem).DefaultValue))
                End If
                writer.WriteEndTag("textarea")
            Else
                writer.WriteBeginTag("input")
                writer.WriteAttribute("type", CStr(If(DirectCast(Item, PageLoader.EditItem).Rows = 0, "hidden", If(DirectCast(Item, PageLoader.EditItem).Password, "password", "text"))))
                writer.WriteAttribute("name", DirectCast(Item, PageLoader.EditItem).Name)
                writer.WriteAttribute("id", DirectCast(Item, PageLoader.EditItem).Name)
                If DirectCast(Item, PageLoader.EditItem).Name = "fontcustom" And Web.HttpContext.Current.Request.Params.Get("fontselection") <> "custom" Then writer.WriteAttribute("style", "display: none;")
                If Web.HttpContext.Current.Request.Params.Get(DirectCast(Item, PageLoader.EditItem).Name) <> String.Empty Then
                    writer.WriteAttribute("value", Web.HttpContext.Current.Request.Params.Get(DirectCast(Item, PageLoader.EditItem).Name))
                ElseIf DirectCast(Item, PageLoader.EditItem).DefaultValue <> String.Empty Then
                    writer.WriteAttribute("value", Utility.LoadResourceString(DirectCast(Item, PageLoader.EditItem).DefaultValue))
                End If

                If DirectCast(Item, PageLoader.EditItem).Rows <> 0 Then writer.WriteAttribute("size", "20")
                writer.Write(HtmlTextWriter.TagRightChar)
            End If
        ElseIf Not UsePrint And (PageLoader.IsDateItem(Item)) Then
            writer.Write(vbCrLf + BaseTabs)
            writer.WriteBeginTag("input")
            writer.WriteAttribute("type", "text")
            writer.WriteAttribute("name", DirectCast(Item, PageLoader.DateItem).Name)
            writer.WriteAttribute("id", DirectCast(Item, PageLoader.DateItem).Name)
            writer.Write(HtmlTextWriter.TagRightChar)
            writer.Write(vbCrLf + BaseTabs)
            writer.WriteBeginTag("div")
            writer.WriteAttribute("id", DirectCast(Item, PageLoader.DateItem).Name + "_div")
            writer.Write(HtmlTextWriter.TagRightChar)
            writer.Write(vbCrLf + BaseTabs)
            writer.WriteEndTag("div")
            AddToJSFunctions(New String() {String.Empty, "$('#" + DirectCast(Item, PageLoader.DateItem).Name + "').calendarsPicker($.extend({yearRange: 'any', alignment:'bottomLeft', popupContainer: '#" + DirectCast(Item, PageLoader.DateItem).Name + "_div'}, '" + Globalization.CultureInfo.CurrentCulture.Name + "' in $.calendarsPicker.regionalOptions ? $.calendarsPicker.regionalOptions['" + Globalization.CultureInfo.CurrentCulture.Name + "'] : ('" + Globalization.CultureInfo.CurrentCulture.TwoLetterISOLanguageName + "' in $.calendarsPicker.regionalOptions ? $.calendarsPicker.regionalOptions['" + Globalization.CultureInfo.CurrentCulture.TwoLetterISOLanguageName + "'] : $.calendarsPicker.regionalOptions[''])));", String.Empty})
        ElseIf Not UsePrint And (PageLoader.IsRadioItem(Item)) Then
            writer.Write(vbCrLf + BaseTabs)
            writer.WriteBeginTag("span")
            writer.WriteAttribute("id", DirectCast(Item, PageLoader.RadioItem).Name + "_")
            If DirectCast(Item, PageLoader.RadioItem).Name = "fromscript" Or DirectCast(Item, PageLoader.RadioItem).Name = "toscript" Or DirectCast(Item, PageLoader.RadioItem).Name = "operation" Then writer.WriteAttribute("style", "display: none;")
            writer.Write(HtmlTextWriter.TagRightChar)
            writer.Write(Utility.HtmlTextEncode(Utility.LoadResourceString(DirectCast(Item, PageLoader.RadioItem).Description)))
            Dim LoadArray As Object() = Nothing
            Dim Length As Integer
            Dim Text As String
            Dim DefaultValue As String
            Dim OnChangeJS() As String = Nothing
            If Not DirectCast(Item, PageLoader.RadioItem).OnChangeFunction Is Nothing Then
                OnChangeJS = CType(DirectCast(Item, PageLoader.RadioItem).OnChangeFunction.Invoke(Nothing, Nothing), String())
                AddToJSFunctions(OnChangeJS)
            End If
            If Not DirectCast(Item, PageLoader.RadioItem).OnPopulateFunction Is Nothing Then
                LoadArray = CType(DirectCast(Item, PageLoader.RadioItem).OnPopulateFunction.Invoke(Nothing, Nothing), Object())
                Length = LoadArray.Length
            Else
                Length = DirectCast(Item, PageLoader.RadioItem).OptionArray.Length
            End If
            If HttpContext.Current.Request.Params.Get(DirectCast(Item, PageLoader.RadioItem).Name) Is Nothing Then
                DefaultValue = DirectCast(Item, PageLoader.RadioItem).DefaultValue
            Else
                DefaultValue = HttpContext.Current.Request.Params.Get(DirectCast(Item, PageLoader.RadioItem).Name)
            End If
            If (DirectCast(Item, PageLoader.RadioItem).UseList) Then
                writer.Write(vbCrLf + BaseTabs + vbTab)
                writer.WriteBeginTag("select")
                writer.WriteAttribute("name", DirectCast(Item, PageLoader.RadioItem).Name)
                writer.WriteAttribute("id", DirectCast(Item, PageLoader.RadioItem).Name)
                If Not DirectCast(Item, PageLoader.RadioItem).OnChangeFunction Is Nothing Then
                    writer.WriteAttribute("onchange", OnChangeJS(0))
                End If
                writer.Write(HtmlTextWriter.TagRightChar)
                For SubIndex = 0 To Length - 1
                    writer.Write(vbCrLf + BaseTabs + vbTab + vbTab)
                    writer.WriteBeginTag("option")
                    If DirectCast(Item, PageLoader.RadioItem).OnPopulateFunction Is Nothing OrElse Not TypeOf LoadArray Is Array() Then
                        writer.WriteAttribute("value", CStr(SubIndex))
                        If SubIndex = CInt(DefaultValue) Then writer.WriteAttribute("selected", "selected")
                    Else
                        writer.WriteAttribute("value", CStr(DirectCast(LoadArray(SubIndex), Object()).GetValue(1)))
                        If CStr(DirectCast(LoadArray(SubIndex), Object()).GetValue(1)) = DefaultValue Then writer.WriteAttribute("selected", "selected")
                    End If
                    writer.Write(HtmlTextWriter.TagRightChar)
                    If DirectCast(Item, PageLoader.RadioItem).OnPopulateFunction Is Nothing Then
                        Text = Utility.LoadResourceString(DirectCast(Item, PageLoader.RadioItem).OptionArray(SubIndex).Name)
                    Else
                        If TypeOf LoadArray Is Array() Then
                            Text = CStr(DirectCast(LoadArray(SubIndex), Object()).GetValue(0))
                        Else
                            Text = CStr(LoadArray(SubIndex))
                        End If
                    End If
                    writer.Write(vbCrLf + BaseTabs + vbTab + vbTab + vbTab)
                    writer.Write(Utility.HtmlTextEncode(Text))
                    writer.Write(vbCrLf + BaseTabs + vbTab + vbTab)
                    writer.WriteEndTag("option")
                Next
                writer.Write(vbCrLf + BaseTabs)
                writer.WriteEndTag("select")
            ElseIf (DirectCast(Item, PageLoader.RadioItem).UseCheck) Then
                writer.Write(vbCrLf + BaseTabs + vbTab)
                writer.WriteBeginTag("input")
                writer.WriteAttribute("type", "checkbox")
                writer.WriteAttribute("name", DirectCast(Item, PageLoader.RadioItem).Name)
                SubIndex = 0
                writer.WriteAttribute("id", DirectCast(Item, PageLoader.RadioItem).Name)
                writer.WriteAttribute("value", CStr(SubIndex))
                If SubIndex = CInt(DefaultValue) Then writer.WriteAttribute("checked", "checked")
                If Not DirectCast(Item, PageLoader.RadioItem).OnChangeFunction Is Nothing Then
                    writer.WriteAttribute("onchange", OnChangeJS(0))
                End If
                writer.Write(HtmlTextWriter.TagRightChar)
            Else
                For SubIndex = 0 To Length - 1
                    writer.Write(vbCrLf + BaseTabs + vbTab)
                    writer.WriteBeginTag("input")
                    writer.WriteAttribute("type", "radio")
                    writer.WriteAttribute("name", DirectCast(Item, PageLoader.RadioItem).Name)
                    writer.WriteAttribute("id", DirectCast(Item, PageLoader.RadioItem).Name + CStr(SubIndex))
                    If DirectCast(Item, PageLoader.RadioItem).OnPopulateFunction Is Nothing OrElse Not TypeOf LoadArray Is Array() Then
                        writer.WriteAttribute("value", CStr(SubIndex))
                        If SubIndex = CInt(DefaultValue) Then writer.WriteAttribute("checked", "checked")
                    Else
                        writer.WriteAttribute("value", CStr(DirectCast(LoadArray(SubIndex), Object()).GetValue(1)))
                        If CStr(DirectCast(LoadArray(SubIndex), Object()).GetValue(1)) = DefaultValue Then writer.WriteAttribute("checked", "checked")
                    End If
                    If Not DirectCast(Item, PageLoader.RadioItem).OnChangeFunction Is Nothing Then
                        writer.WriteAttribute("onchange", OnChangeJS(0))
                    End If
                    writer.Write(HtmlTextWriter.TagRightChar)
                    If DirectCast(Item, PageLoader.RadioItem).OnPopulateFunction Is Nothing Then
                        Text = Utility.LoadResourceString(DirectCast(Item, PageLoader.RadioItem).OptionArray(SubIndex).Name)
                    Else
                        If TypeOf LoadArray Is Array() Then
                            Text = CStr(DirectCast(LoadArray(SubIndex), Object()).GetValue(0))
                        Else
                            Text = CStr(LoadArray(SubIndex))
                        End If
                    End If
                    writer.Write(Utility.HtmlTextEncode(Text))
                Next
            End If
            writer.WriteEndTag("span")
        ElseIf Not UsePrint And (PageLoader.IsButtonItem(Item)) Then
            writer.Write(vbCrLf + BaseTabs)
            writer.WriteBeginTag("input")
            writer.WriteAttribute("name", DirectCast(Item, PageLoader.ButtonItem).Name)
            writer.WriteAttribute("id", DirectCast(Item, PageLoader.ButtonItem).Name)
            If Not DirectCast(Item, PageLoader.ButtonItem).OnRenderFunction Is Nothing Then
                writer.WriteAttribute("value", CStr(DirectCast(Item, PageLoader.ButtonItem).OnRenderFunction.Invoke(Nothing, New Object() {Item})))
            Else
                writer.WriteAttribute("value", Utility.LoadResourceString(DirectCast(Item, PageLoader.ButtonItem).Description))
            End If
            If DirectCast(Item, PageLoader.ButtonItem).Name = "fontcustomapply" And Web.HttpContext.Current.Request.Params.Get("fontselection") <> "custom" Then writer.WriteAttribute("style", "display: none;")
            If Not DirectCast(Item, PageLoader.ButtonItem).OnClickFunction Is Nothing Then
                Dim OnClickJS As String() = CType(DirectCast(Item, PageLoader.ButtonItem).OnClickFunction.Invoke(Nothing, Nothing), String())
                AddToJSFunctions(OnClickJS)
                writer.WriteAttribute("onclick", OnClickJS(0))
                writer.WriteAttribute("type", "button")
            Else
                writer.WriteAttribute("type", "submit")
            End If
            writer.Write(HtmlTextWriter.TagRightChar)
        ElseIf (PageLoader.IsTextItem(Item)) Then
            WriteTextItem(writer, DirectCast(Item, PageLoader.TextItem), TabCount, IndexString)
        End If
    End Sub
    Protected Overrides Sub Render(ByVal writer As System.Web.UI.HtmlTextWriter)
        Dim Index As Integer
        If UsePanes Then
            'no spacing since inline-block element
            writer.WriteBeginTag("div")
            writer.WriteAttribute("id", "panediv")
            writer.Write(HtmlTextWriter.TagRightChar)
            writer.Write(vbCrLf + vbTab + vbTab + vbTab)
            writer.WriteBeginTag("div")
            writer.WriteAttribute("id", "panewrapper")
            writer.Write(HtmlTextWriter.TagRightChar)
        End If
        For Index = 0 To MyPage.Page.Count - 1
            writer.Write(vbCrLf + vbTab + vbTab + vbTab + vbTab)
            writer.WriteBeginTag("div")
            If Index <> MyPage.Page.Count - 1 AndAlso PageLoader.IsDownloadItem(MyPage.Page.Item(Index + 1)) Then
                writer.Write(HtmlTextWriter.TagRightChar)
                writer.Write(Utility.ConvertSpaces(HtmlTextWriter.SpaceChar + HtmlTextWriter.SpaceChar + HtmlTextWriter.SpaceChar + HtmlTextWriter.SpaceChar + HtmlTextWriter.SpaceChar))
            Else
                writer.Write(HtmlTextWriter.TagRightChar)
            End If
            RenderSingleItem(writer, MyPage.Page.Item(Index), MyPage.PageName, 5)
            writer.Write(vbCrLf + vbTab + vbTab + vbTab + vbTab)
            writer.WriteEndTag("div")
        Next
        If UsePanes Then
            writer.Write(vbCrLf + vbTab + vbTab + vbTab)
            writer.WriteEndTag("div")
            writer.Write(vbCrLf + vbTab + vbTab + vbTab)
            writer.WriteEndTag("div")
        End If
    End Sub
End Class