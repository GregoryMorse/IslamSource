Imports HostPageUtility
Imports XMLRender
Partial Class host
    Inherits System.Web.UI.Page
    Dim _IsHtml As Boolean = False
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
    Public Const RootPath As String = "/"
    Public Const MainPage As String = "host.aspx"
    Public Const PageQuery As String = "Page"
    Public Const PagePrintQuery As String = "PagePrint"
    Public Const LangSet As String = "LangSet"
    Public Shared Function GetPageString(Page As String) As String
        Return RootPath + MainPage + "?" + PageQuery + "=" + Page
    End Function
    'Web.Config requires configuration -> system.webServer -> defaultDocument -> files -> <clear /><add value="host.aspx" />
    'TODO: Upload script which also handles ImageItem and DownloadItem files
    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Dim _dlgt As AsyncTaskDelegate = Nothing
        'Context.Response.BufferOutput = True
        AsyncTimeout = New TimeSpan(0, 0, 240)
        RegisterAsyncTask(New PageAsyncTask(Function(_sender As Object, _e As EventArgs, cb As AsyncCallback, extraData As Object) As IAsyncResult
                                                'adapter for .NET 4
                                                'Dim Context As HttpContext = HttpContext.Current
                                                _dlgt = New AsyncTaskDelegate(Sub() DoAsyncLoad().Wait())
                                                Return _dlgt.BeginInvoke(cb, extraData)
                                            End Function,
                                            Sub(ar As IAsyncResult)
                                                _dlgt.EndInvoke(ar)
                                            End Sub,
                                            Sub(ar As IAsyncResult)

                                            End Sub, "Async"))
        'ExecuteRegisteredAsyncTasks()
    End Sub
    Protected Delegate Sub AsyncTaskDelegate()
    Private _PortableMethods As PortableMethods
    Private UWeb As UtilityWeb
    Private SD As SiteDatabase
    Private UA As UserAccounts
    Private ArbData As ArabicData
    Dim PageSet As PageLoader
    Public Function IsLoggedIn(Context As HttpContext) As Boolean
        Return UA.IsLoggedIn(Context)
    End Function
    Public Async Function DoAsyncLoad() As Threading.Tasks.Task
        Dim Index As Integer
        Dim IsPrint As Boolean = False
        Dim bmp As Bitmap = Nothing
        Dim ResultBmp As Bitmap = Nothing
        If Application("PortableMethods") Is Nothing Then
            Application.Add("PortableMethods", New XMLRender.PortableMethods(New HostPageUtility.WindowsWebFileIO(), New HostPageUtility.WindowsWebSettings(Context.Request.PhysicalApplicationPath)))
            Await CType(Application("PortableMethods"), PortableMethods).Init()
        End If
        _PortableMethods = CType(Application("PortableMethods"), PortableMethods)
        If Application("ArbData") Is Nothing Then
            Application.Add("ArbData", New ArabicData(_PortableMethods))
            Await CType(Application("ArbData"), ArabicData).Init()
        End If
        ArbData = CType(Application("ArbData"), ArabicData)
        If Application("UWeb") Is Nothing Then
            Application.Add("UWeb", New UtilityWeb(_PortableMethods, ArbData, AddressOf host.GetPageString, AddressOf UserAccounts.GetUserID, AddressOf IsLoggedIn))
            CType(_PortableMethods.Settings, HostPageUtility.WindowsWebSettings).Init(CType(Application("UWeb"), UtilityWeb))
        End If
        UWeb = CType(Application("UWeb"), UtilityWeb)
        If Application("SD") Is Nothing Then
            Application.Add("SD", New SiteDatabase(UWeb))
            Await UWeb.Init(CType(Application("SD"), SiteDatabase))
        End If
        SD = CType(Application("SD"), SiteDatabase)
        If Application("PageSet") Is Nothing Then
            Application.Add("PageSet", New PageLoader(_PortableMethods, UWeb))
            Await CType(Application("PageSet"), PageLoader).Init(Context.Request.Url.Host)
        End If
        PageSet = CType(Application("PageSet"), PageLoader)
        If Application("UA") Is Nothing Then Application.Add("UA", New UserAccounts(_PortableMethods, UWeb, ArbData, PageSet, SD))
        UA = CType(Application("UA"), UserAccounts)
        Controls.Clear() 'clear viewstate and default HTML rendering control container
        If Context.Request.QueryString.Get(LangSet) <> String.Empty Then
            UserAccounts.SetLanguage(Context.Request.QueryString.Get(LangSet), Context)
        End If
        Dim LangID As String = UserAccounts.GetLanguage(Context)
        If LangID <> String.Empty Then
            Threading.Thread.CurrentThread.CurrentUICulture = Globalization.CultureInfo.CreateSpecificCulture(LangID)
            Threading.Thread.CurrentThread.CurrentCulture = Globalization.CultureInfo.CreateSpecificCulture(LangID)
        End If
        If Context.Request.QueryString.Get(PageQuery) = UserAccounts.ID_ClearCache AndAlso UA.IsAdmin(Context) Then
            Dim enumerator As IDictionaryEnumerator = Cache.GetEnumerator()
            Dim keys As New Collections.Generic.List(Of String)
            While enumerator.MoveNext
                keys.Add(CStr(enumerator.Key))
            End While
            For Index = 0 To keys.Count() - 1
                Cache.Remove(keys(Index))
            Next
            GC.Collect()
        ElseIf Context.Request.QueryString.Get(PageQuery) = UserAccounts.ID_ViewHeaders AndAlso UA.IsAdmin(Context) Then
            Dim encoding As System.Text.Encoding = System.Text.Encoding.ASCII
            Context.Response.ContentType = "text/plain;charset=" + encoding.WebName
            Context.Response.Write("Headers for: " + Context.Request.Url.Host + vbCrLf)
            For Index = 0 To Context.Request.Headers.AllKeys().Length - 1
                Context.Response.Write(Context.Request.Headers.AllKeys()(Index) + ": ")
                Dim Count As Integer
                For Count = 0 To Context.Request.Headers.GetValues(Context.Request.Headers.AllKeys()(Index)).Length - 1
                    Context.Response.Write(Context.Request.Headers.GetValues(Context.Request.Headers.AllKeys()(Index))(Count) + CStr(If(Count <> Context.Request.Headers.GetValues(Context.Request.Headers.AllKeys()(Index)).Length - 1, ", ", String.Empty)))
                Next
                Context.Response.Write(vbCrLf)
            Next
        ElseIf Context.Request.QueryString.Get(PageQuery) = UserAccounts.ID_DotNetVersion AndAlso UA.IsAdmin(Context) Then
            Dim encoding As System.Text.Encoding = System.Text.Encoding.ASCII
            Context.Response.ContentType = "text/plain;charset=" + encoding.WebName
            Context.Response.Write(Environment.Version)
        ElseIf Context.Request.QueryString.Get(PageQuery) = UserAccounts.ID_ViewCache AndAlso UA.IsAdmin(Context) Then
            Dim enumerator As IDictionaryEnumerator = Cache.GetEnumerator()
            Dim keys As New Generic.List(Of Object)
            Dim encoding As System.Text.Encoding = System.Text.Encoding.ASCII
            Context.Response.ContentType = "text/plain;charset=" + encoding.WebName
            While enumerator.MoveNext
                Context.Response.Write(CStr(enumerator.Key) + vbCrLf)
            End While
        ElseIf Context.Request.QueryString.Get(PageQuery) = UserAccounts.ID_ClearDiskCache AndAlso UA.IsAdmin(Context) Then
            Await _PortableMethods.DiskCache.DeleteUnusedCacheItems(New String() {})
            GC.Collect()
        ElseIf Context.Request.QueryString.Get(PageQuery) = UserAccounts.ID_ViewDiskCache AndAlso UA.IsAdmin(Context) Then
            Dim CacheItems() As String = Await _PortableMethods.DiskCache.GetCacheItems()
            Dim encoding As System.Text.Encoding = System.Text.Encoding.ASCII
            Context.Response.ContentType = "text/plain;charset=" + encoding.WebName
            For Index = 0 To CacheItems.Length - 1
                Dim FileInfo As New IO.FileInfo(CacheItems(Index))
                Context.Response.Write("Name: " + CacheItems(Index) + " Size: " + CStr(FileInfo.Length) + " Last Modified: " + FileInfo.LastWriteTimeUtc.ToString((New Globalization.DateTimeFormatInfo).FullDateTimePattern) + vbCrLf)
            Next
        ElseIf Context.Request.QueryString.Get(PageQuery) = UserAccounts.ID_CreateDatabase AndAlso UA.IsAdmin(Context) Then
            SD.CreateDatabase()
            UWeb.LookupClassMember("IslamSiteDatabase::CreateDatabase").Invoke(Nothing)
        ElseIf Context.Request.QueryString.Get(PageQuery) = UserAccounts.ID_RemoveDatabase AndAlso UA.IsAdmin(Context) Then
            SD.RemoveDatabase()
            UWeb.LookupClassMember("IslamSiteDatabase::RemoveDatabase").Invoke(Nothing)
        ElseIf Context.Request.QueryString.Get(PageQuery) = UserAccounts.ID_CleanupState AndAlso UA.IsAdmin(Context) Then
            SD.CleanupStaleLoginSessions()
            SD.CleanupStaleActivations()
        ElseIf Context.Request.QueryString.Get(PageQuery) = UserAccounts.ID_ViewCertRequests AndAlso UA.IsAdmin(Context) Then
            Dim Store As New System.Security.Cryptography.X509Certificates.X509Store("REQUEST", System.Security.Cryptography.X509Certificates.StoreLocation.LocalMachine)
            Store.Open(System.Security.Cryptography.X509Certificates.OpenFlags.MaxAllowed)
            Dim encoding As System.Text.Encoding = System.Text.Encoding.ASCII
            Context.Response.ContentType = "text/plain;charset=" + encoding.WebName
            Context.Response.Write("Machine Certificate Enrollment Requests:" + vbCrLf)
            For Each Cert As System.Security.Cryptography.X509Certificates.X509Certificate2 In Store.Certificates
                Context.Response.Write(Cert.Subject + vbCrLf)
            Next
            Store.Close()
            Store = New System.Security.Cryptography.X509Certificates.X509Store(System.Security.Cryptography.X509Certificates.StoreName.CertificateAuthority, System.Security.Cryptography.X509Certificates.StoreLocation.LocalMachine)
            Store.Open(System.Security.Cryptography.X509Certificates.OpenFlags.MaxAllowed)
            Context.Response.Write("Machine Intermediate Certificate Authorities:" + vbCrLf)
            For Each Cert As System.Security.Cryptography.X509Certificates.X509Certificate2 In Store.Certificates
                Context.Response.Write(Cert.Subject + vbCrLf)
            Next
            Store.Close()
            Store = New System.Security.Cryptography.X509Certificates.X509Store(System.Security.Cryptography.X509Certificates.StoreName.My, System.Security.Cryptography.X509Certificates.StoreLocation.LocalMachine)
            Store.Open(System.Security.Cryptography.X509Certificates.OpenFlags.MaxAllowed)
            Context.Response.Write("Machine Personal:" + vbCrLf)
            For Each Cert As System.Security.Cryptography.X509Certificates.X509Certificate2 In Store.Certificates
                Context.Response.Write(Cert.Subject + vbCrLf)
            Next
            Store.Close()
            Store = New System.Security.Cryptography.X509Certificates.X509Store("REQUEST", System.Security.Cryptography.X509Certificates.StoreLocation.CurrentUser)
            Store.Open(System.Security.Cryptography.X509Certificates.OpenFlags.MaxAllowed)
            Context.Response.Write("Certificate Enrollment Requests:" + vbCrLf)
            For Each Cert As System.Security.Cryptography.X509Certificates.X509Certificate2 In Store.Certificates
                Context.Response.Write(Cert.Subject + vbCrLf)
            Next
            Store.Close()
            Store = New System.Security.Cryptography.X509Certificates.X509Store(System.Security.Cryptography.X509Certificates.StoreName.CertificateAuthority, System.Security.Cryptography.X509Certificates.StoreLocation.CurrentUser)
            Store.Open(System.Security.Cryptography.X509Certificates.OpenFlags.MaxAllowed)
            Context.Response.Write("Intermediate Certificate Authorities:" + vbCrLf)
            For Each Cert As System.Security.Cryptography.X509Certificates.X509Certificate2 In Store.Certificates
                Context.Response.Write(Cert.Subject + vbCrLf)
            Next
            Store.Close()
            Store = New System.Security.Cryptography.X509Certificates.X509Store(System.Security.Cryptography.X509Certificates.StoreName.My, System.Security.Cryptography.X509Certificates.StoreLocation.CurrentUser)
            Store.Open(System.Security.Cryptography.X509Certificates.OpenFlags.MaxAllowed)
            Context.Response.Write("Personal:" + vbCrLf)
            For Each Cert As System.Security.Cryptography.X509Certificates.X509Certificate2 In Store.Certificates
                Context.Response.Write(Cert.Subject + vbCrLf)
            Next
            Store.Close()
        ElseIf Context.Request.QueryString.Get(PageQuery) = UserAccounts.ID_ViewSites AndAlso UA.IsAdmin(Context) Then
            Dim ServerManager As New Microsoft.Web.Administration.ServerManager
            Dim encoding As System.Text.Encoding = System.Text.Encoding.ASCII
            Context.Response.ContentType = "text/plain;charset=" + encoding.WebName
            For Each Site As Microsoft.Web.Administration.Site In ServerManager.Sites
                Context.Response.Write("Site Name: " + Site.Name + vbCrLf)
                For Each Binding As Microsoft.Web.Administration.Binding In Site.Bindings
                    Context.Response.Write("Binding: " + Binding.BindingInformation + vbCrLf)
                    If Not Binding.CertificateStoreName Is Nothing Then Context.Response.Write("Store Name: " + Binding.CertificateStoreName + vbCrLf)
                    If Not Binding.CertificateHash Is Nothing Then Context.Response.Write("Certificate Hash: " + String.Concat(Linq.Enumerable.Select(Of Byte, String)(Binding.CertificateHash, Function(bt As Byte) bt.ToString("X2"))) + vbCrLf)
                Next
                Context.Response.Write(vbCrLf)
            Next
        ElseIf Context.Request.QueryString.Get(PageQuery) = UserAccounts.ID_EnableSSL AndAlso UA.IsAdmin(Context) Then
            Dim ServerManager As New Microsoft.Web.Administration.ServerManager
            For Each Domain As String In UtilityWeb.ConnectionData.SiteDomains
                ServerManager.Sites(Domain).Bindings.Add("*:443:", "https")
            Next
            ServerManager.CommitChanges()
        ElseIf Context.Request.QueryString.Get(PageQuery) = UserAccounts.ID_CurrentUser AndAlso UA.IsAdmin(Context) Then
            Dim encoding As System.Text.Encoding = System.Text.Encoding.ASCII
            Context.Response.ContentType = "text/plain;charset=" + encoding.WebName
            Context.Response.Write("Execution User: " + System.Security.Principal.WindowsIdentity.GetCurrent().Name() + vbCrLf)
            Context.Response.Write("Http Context User: " + Context.User.Identity.Name() + vbCrLf)
        ElseIf Context.Request.QueryString.Get(PageQuery) = UserAccounts.ID_HadithRanking Then
            If UA.IsLoggedIn(Context) Then
                UWeb.LookupClassMember("IslamSiteDatabase::UpdateRankingData").Invoke(New Object() {UserAccounts.GetUserID(Context), Context})
            End If
            UWeb.LookupClassMember("IslamSiteDatabase::WriteRankingData").Invoke(New Object() {Context})
        ElseIf (Context.Request.QueryString.Get(PageQuery) = "Image.gif") Then
            Dim DateModified As Date = Now
            If Context.Request.QueryString.Get("Image") = "EMailAddress" Or
                Context.Request.QueryString.Get("Image") = "GradientBackground" Or
                Context.Request.QueryString.Get("Image") = "GradientBackgroundBottom" Then
                DateModified = IO.File.GetLastWriteTimeUtc(_PortableMethods.Settings.GetFilePath("bin\HostPage.dll"))
            ElseIf Context.Request.QueryString.Get("Image") = "Scale" Then
                Dim FetchImageItem As PageLoader.ImageItem
                If Context.Request.QueryString.Get("p") = "menu.main" Then
                    FetchImageItem = New PageLoader.ImageItem(String.Empty, String.Empty, PageSet.MainImage, Nothing, 206, 200)
                ElseIf Context.Request.QueryString.Get("p") = "menu.hover" Then
                    FetchImageItem = New PageLoader.ImageItem(String.Empty, String.Empty, PageSet.HoverImage, Nothing, 206, 200)
                Else
                    FetchImageItem = DirectCast(PageSet.GetPageItem(Context.Request.QueryString.Get("p")), PageLoader.ImageItem)
                End If
                'check XML timestamp also
                DateModified = IO.File.GetLastWriteTimeUtc(_PortableMethods.Settings.GetFilePath("images\" + FetchImageItem.Path))
            ElseIf Context.Request.QueryString.Get("Image") = "Thumb" Then
                Dim FetchImageItem As PageLoader.TextItem = DirectCast(PageSet.GetPageItem(Context.Request.QueryString.Get("p")), PageLoader.TextItem)
                'check XML timestamp also
                If CInt((Await _PortableMethods.DiskCache.GetCacheItems()).Length * New Random().NextDouble()) = 0 Then
                    DateModified = UtilityWeb.GetURLLastModified(FetchImageItem.ImageURL)
                Else
                    DateModified = Date.MinValue
                End If
            End If
            Dim Bytes() As Byte = Await _PortableMethods.DiskCache.GetCacheItem(Context.Request.Url.Host + "_" + Context.Request.QueryString().ToString(), DateModified)
            If Not Bytes Is Nothing Then
                ResultBmp = DirectCast(Bitmap.FromStream(New IO.MemoryStream(Bytes)), Bitmap)
            End If
            If ResultBmp Is Nothing Then
                If Context.Request.QueryString.Get("Image") = "EMailAddress" Then
                    Dim oFont As New Font("Arial", 13)
                    Dim TextExtent As SizeF = UtilityWeb.GetTextExtent(UtilityWeb.ConnectionData.EMailAddress, oFont)
                    bmp = New Bitmap(CInt(Math.Ceiling(Math.Ceiling(TextExtent.Width + 1) * 96.0F / 72.0F)), CInt(Math.Ceiling(Math.Ceiling(TextExtent.Height + 1) * 96.0F / 72.0F)))
                    Dim g As Graphics = Graphics.FromImage(bmp)
                    g.PageUnit = GraphicsUnit.Point
                    g.TextRenderingHint = Drawing.Text.TextRenderingHint.AntiAliasGridFit
                    g.TextContrast = 0
                    g.FillRectangle(Brushes.White, New RectangleF(0, 0, CInt(Math.Ceiling(Math.Ceiling(TextExtent.Width + 1) * 96.0F / 72.0F)), CInt(Math.Ceiling(Math.Ceiling(TextExtent.Height + 1) * 96.0F / 72.0F))))
                    g.DrawString(UtilityWeb.ConnectionData.EMailAddress, oFont, Brushes.Black, New RectangleF(0, 0, CSng(Math.Ceiling(TextExtent.Width + 1)), CSng(Math.Ceiling(TextExtent.Height + 1))), Drawing.StringFormat.GenericTypographic)
                    bmp.MakeTransparent(Color.White)
                    oFont.Dispose()
                ElseIf Context.Request.QueryString.Get("Image") = "MandelbrotFractal" Or Context.Request.QueryString.Get("Image") = "JuliaFractal" Then
                    Dim FractalBmp As New Bitmap(50, 50)
                    If Context.Request.QueryString.Get("Image") = "MandelbrotFractal" Then
                        ImageQuantization.Fractal.StandardMandelbrotSet(FractalBmp)
                    Else
                        ImageQuantization.Fractal.StandardJuliaSet(FractalBmp)
                    End If
                    bmp = New Bitmap(32 * 3, 32 * 3)
                    Dim g As Graphics = Graphics.FromImage(bmp)
                    Dim adj As Single = CSng(32 * Math.Sqrt(2))
                    g.ResetTransform()
                    g.RotateTransform(45.0F)
                    g.DrawImage(FractalBmp, -50 + adj, -25) 'top left
                    g.ResetTransform()
                    g.RotateTransform(45.0F)
                    g.RotateTransform(90.0F)
                    g.DrawImage(FractalBmp, -50 + adj - adj * 2 + adj / 2, -25 - adj * 2 + adj / 2) 'top right
                    g.ResetTransform()
                    g.RotateTransform(45.0F)
                    g.ScaleTransform(-1.0F, 1.0F)
                    g.DrawImage(FractalBmp, -50 + adj - adj * 3, -25) 'bottom left
                    g.ResetTransform()
                    g.RotateTransform(45.0F)
                    g.RotateTransform(90.0F)
                    g.ScaleTransform(-1.0F, 1.0F)
                    g.DrawImage(FractalBmp, -50 + adj - adj * 2 + adj / 2, -25 - adj * 2 + adj / 2) 'bottom right
                    g.ResetTransform()
                    g.DrawImage(FractalBmp, 0, 32, 32, 32) 'left
                    g.ResetTransform()
                    g.RotateTransform(90.0F)
                    g.DrawImage(FractalBmp, 0, -32 - 32, 32, 32) 'top
                    g.ResetTransform()
                    g.ScaleTransform(-1.0F, 1.0F)
                    g.DrawImage(FractalBmp, -32 - 32 - 32, 32, 32, 32) 'right
                    g.ResetTransform()
                    g.RotateTransform(90.0F)
                    g.ScaleTransform(-1.0F, 1.0F)
                    g.DrawImage(FractalBmp, -32 - 32 - 32, -32 - 32, 32, 32) 'bottom
                ElseIf Context.Request.QueryString.Get("Image") = "GradientBackground" Then
                    bmp = New Bitmap(1, 320)
                    Dim g As Graphics = Graphics.FromImage(bmp)
                    Dim oBrush As New Drawing2D.LinearGradientBrush(New Rectangle(0, 0, 1, 320), Color.LightBlue, Color.LightGray, 90)
                    g.FillRectangle(oBrush, New Rectangle(0, 0, 1, 320))
                ElseIf Context.Request.QueryString.Get("Image") = "GradientBackgroundBottom" Then
                    bmp = New Bitmap(1, 320)
                    Dim g As Graphics = Graphics.FromImage(bmp)
                    Dim oBrush As New Drawing2D.LinearGradientBrush(New Rectangle(0, 0, 1, 320), Color.LightGray, Color.LightBlue, 90)
                    g.FillRectangle(oBrush, New Rectangle(0, 0, 1, 320))
                ElseIf Context.Request.QueryString.Get("Image") = "Scale" Then
                    Dim SizeF As System.Drawing.SizeF
                    Dim Scale As Double
                    Dim FetchImageItem As PageLoader.ImageItem
                    Dim OriginalBmp As Bitmap
                    If Context.Request.QueryString.Get("p") = "menu.main" Then
                        FetchImageItem = New PageLoader.ImageItem(String.Empty, String.Empty, PageSet.MainImage, Nothing, 206, 200)
                        OriginalBmp = DirectCast(Bitmap.FromFile(_PortableMethods.Settings.GetFilePath("images\" + FetchImageItem.Path)), Bitmap)
                        UtilityWeb.AddTextLogo(OriginalBmp, _PortableMethods.LoadResourceString(PageSet.Title))
                    ElseIf Context.Request.QueryString.Get("p") = "menu.hover" Then
                        FetchImageItem = New PageLoader.ImageItem(String.Empty, String.Empty, PageSet.HoverImage, Nothing, 206, 200)
                        OriginalBmp = DirectCast(Bitmap.FromFile(_PortableMethods.Settings.GetFilePath("images\" + FetchImageItem.Path)), Bitmap)
                        UtilityWeb.AddTextLogo(OriginalBmp, _PortableMethods.LoadResourceString(PageSet.Title))
                    Else
                        'check boundaries
                        FetchImageItem = DirectCast(PageSet.GetPageItem(Context.Request.QueryString.Get("p")), PageLoader.ImageItem)
                        OriginalBmp = DirectCast(Bitmap.FromFile(_PortableMethods.Settings.GetFilePath("images\" + FetchImageItem.Path)), Bitmap)
                    End If
                    SizeF = OriginalBmp.GetBounds(Drawing.GraphicsUnit.Pixel).Size
                    Scale = UtilityWeb.ComputeImageScale(SizeF.Width, SizeF.Height, FetchImageItem.MaxX, FetchImageItem.MaxY)
                    bmp = UtilityWeb.MakeThumbnail(OriginalBmp, Convert.ToInt32(SizeF.Width / Scale), Convert.ToInt32(SizeF.Height / Scale))
                    If Context.Request.QueryString.Get("p") = "menu.main" Then
                        UtilityWeb.ApplyTransparencyFilter(bmp, Color.FromArgb(&HFFF0F0F0), Color.White)
                    End If
                ElseIf Context.Request.QueryString.Get("Image") = "UnicodeChar" Then
                    bmp = UWeb.GetUnicodeChar(If(Context.Request.QueryString.Get("Size") = Nothing, 13, CInt(Context.Request.QueryString.Get("Size"))), Context.Request.QueryString.Get("Font"), Char.ConvertFromUtf32(Integer.Parse(Context.Request.QueryString.Get("Char"), Globalization.NumberStyles.HexNumber)))
                ElseIf Context.Request.QueryString.Get("Image") = "Thumb" Then
                    Dim FetchImageItem As PageLoader.TextItem = DirectCast(PageSet.GetPageItem(Context.Request.QueryString.Get("p")), PageLoader.TextItem)
                    bmp = UtilityWeb.MakeThumbFromURL(FetchImageItem.ImageURL, 121)
                End If
                If Not bmp Is Nothing Then
                    Dim quantizer As ImageQuantization.OctreeQuantizer = New ImageQuantization.OctreeQuantizer(255, 8, Not Bitmap.IsAlphaPixelFormat(bmp.PixelFormat), Color.White)
                    ResultBmp = quantizer.QuantizeBitmap(bmp)
                    bmp.Dispose()
                End If
            End If
            If Not ResultBmp Is Nothing Then
                Context.Response.ContentType = "image/gif"
                'Save crashes because it calls get_Position on the stream
                'ResultBmp.Save(Response.OutputStream, System.Drawing.Imaging.ImageFormat.Gif)
                Dim MemStream As New IO.MemoryStream()
                ResultBmp.Save(MemStream, CType(If(Object.Equals(ResultBmp.RawFormat, Drawing.Imaging.ImageFormat.MemoryBmp), Drawing.Imaging.ImageFormat.Gif, ResultBmp.RawFormat), Drawing.Imaging.ImageFormat))
                If Bytes Is Nothing Then Await _PortableMethods.DiskCache.CacheItem(Context.Request.Url.Host + "_" + Context.Request.QueryString().ToString(), DateModified, MemStream.GetBuffer())
                Context.Response.Cache.SetCacheability(HttpCacheability.Public)
                Context.Response.OutputStream.Write(MemStream.ToArray(), 0, CInt(MemStream.Length))
                Context.Response.Flush()
            End If
            ResultBmp.Dispose()
            GC.Collect()
        ElseIf Context.Request.QueryString.Get(PageQuery) = "Source" Then
            Dim f As IO.FileStream = New IO.FileStream(CStr(If(IO.File.Exists(_PortableMethods.Settings.GetFilePath("files\" + Context.Request.QueryString.Get("File"))), _PortableMethods.Settings.GetFilePath("files\" + Context.Request.QueryString.Get("File")), _PortableMethods.Settings.GetFilePath("metadata\" + Context.Request.QueryString.Get("File")))), IO.FileMode.Open, IO.FileAccess.Read, IO.FileShare.Read)
            Dim buffer As Byte() = New Byte(CInt(f.Length) - 1) {}
            Dim count As Integer = f.Read(buffer, 0, CInt(f.Length))
            Dim encoding As System.Text.Encoding = UtilityWeb.DetectEncoding(buffer)
            If encoding Is Nothing Then
                encoding = System.Text.Encoding.ASCII
            End If
            'Convert tabs to spaces
            Context.Response.ContentType = "text/plain;charset=" + encoding.WebName
            Context.Response.Write(UtilityWeb.SourceTextEncode(encoding.GetChars(buffer, encoding.GetPreamble().Length, count - encoding.GetPreamble().Length)))
            Context.Response.Flush()
        Else
            If Context.Request.Params(PageQuery) = "Print" Then
                IsPrint = True
            ElseIf Context.Request.Params(PageQuery) = "PrintFlashcardPdf" Or Context.Request.Params(PageQuery) = "PrintPdf" Or Context.Request.Params(PageQuery) = "PrintDocX" Then
                Dim MemStream As New IO.MemoryStream()
                Dim RenderItems As New Generic.List(Of RenderArray.RenderItem)
                Index = PageSet.GetPageIndex(Context.Request.Params(PagePrintQuery))
                For Count As Integer = 0 To PageSet.Pages(Index).Page.Count - 1
                    If PageLoader.IsListItem(PageSet.Pages(Index).Page(Count)) Then
                        For SubIndex As Integer = 0 To DirectCast(PageSet.Pages(Index).Page(Count), PageLoader.ListItem).List.Count - 1
                            If PageLoader.IsTextItem(DirectCast(PageSet.Pages(Index).Page(Count), PageLoader.ListItem).List.Item(SubIndex)) Then
                                Dim Item As PageLoader.TextItem = CType(CType(PageSet.Pages(Index).Page(Count), PageLoader.ListItem).List.Item(SubIndex), PageLoader.TextItem)
                                If Not Item.OnRenderFunction Is Nothing Then
                                    Dim Output As Object = Item.OnRenderFunction.Invoke(New Object() {Item, Context})
                                    If UtilityWeb.IsTask(Output) Then
                                        If TypeOf Output Is Threading.Tasks.Task(Of RenderArray) Then
                                            Output = Await CType(Output, Threading.Tasks.Task(Of RenderArray))
                                        Else
                                            Output = Await CType(Output, Threading.Tasks.Task(Of Array()))
                                        End If
                                    End If
                                    If TypeOf Output Is RenderArray Then
                                        RenderItems.AddRange(CType(Output, RenderArray).Items)
                                    ElseIf TypeOf Output Is Array() Then
                                        RenderItems.Add(New RenderArray.RenderItem(RenderArray.RenderTypes.eText, New RenderArray.RenderText() {New RenderArray.RenderText(RenderArray.RenderDisplayClass.eList, Output)}))
                                    End If
                                End If
                            End If
                        Next
                    ElseIf PageLoader.IsTextItem(PageSet.Pages(Index).Page(Count)) Then
                        Dim Item As PageLoader.TextItem = CType(PageSet.Pages(Index).Page(Count), PageLoader.TextItem)
                        If Not Item.OnRenderFunction Is Nothing Then
                            Dim Output As Object = Item.OnRenderFunction.Invoke(New Object() {Item, Context})
                            If UtilityWeb.IsTask(Output) Then Output = Await CType(Output, Threading.Tasks.Task(Of RenderArray))
                            If TypeOf Output Is RenderArray Then
                                RenderItems.AddRange(CType(Output, RenderArray).Items)
                            End If
                        End If
                    End If
                Next
                If RenderItems.Count <> 0 Then
                    Dim RAWeb As New RenderArrayWeb(Nothing, _PortableMethods, ArbData, UWeb)
                    If Context.Request.Params(PageQuery) = "PrintFlashcardPdf" Then
                        RAWeb.OutputFlashcardPdf(MemStream, RenderItems)
                        Context.Response.ContentType = "application/pdf"
                    ElseIf Context.Request.Params(PageQuery) = "PrintPdf" Then
                        RAWeb.OutputPdf(MemStream, RenderItems)
                        Context.Response.ContentType = "application/pdf"
                    Else
                        RAWeb.OutputDocX(MemStream, RenderItems)
                        Context.Response.ContentType = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    End If
                    Context.Response.OutputStream.Write(MemStream.ToArray(), 0, CInt(MemStream.Length))
                    Context.Response.Flush()
                    'Dim Bytes(4096) As Byte
                    'Dim Read As Integer
                    'MemStream.Seek(0, IO.SeekOrigin.Begin)
                    'Read = MemStream.Read(Bytes, 0, Bytes.Length)
                    'While Read <> 0
                    '    Context.Response.OutputStream.Write(Bytes, 0, Read)
                    '    Read = MemStream.Read(Bytes, 0, Bytes.Length)
                    'End While
                    Return
                End If
            ElseIf Context.Request.QueryString.Get(PageQuery) = UserAccounts.ID_Register Then
                UA.Register(PageSet, Context, Context.Request.Form.Get(UserAccounts.ID_Username), Context.Request.Form.Get(UserAccounts.ID_Password), Context.Request.Form.Get(UserAccounts.ID_ConfirmPassword), Context.Request.Form.Get(UserAccounts.ID_EmailAddress), Context.Request.Form.Get(UserAccounts.ID_ConfirmEmailAddress), Context.Request.Form.Get(UserAccounts.ID_Register))
            ElseIf Context.Request.QueryString.Get(PageQuery) = UserAccounts.ID_Login Then
                UA.Login(PageSet, Context, Context.Request.Form.Get(UserAccounts.ID_Username), Context.Request.Form.Get(UserAccounts.ID_Password), Context.Request.Form.Get(UserAccounts.ID_Remember), Context.Request.Form.Get(UserAccounts.ID_Login))
            ElseIf Context.Request.QueryString.Get(PageQuery) = UserAccounts.ID_Logoff Then
                UA.Logoff(PageSet, Context)
            ElseIf Context.Request.QueryString.Get(PageQuery) = UserAccounts.ID_ForgotUsername Then
                UA.ForgotUserName(PageSet, Context, Context.Request.Form.Get(UserAccounts.ID_EmailAddress), Context.Request.Form.Get(UserAccounts.ID_RetrieveUsername))
            ElseIf Context.Request.QueryString.Get(PageQuery) = UserAccounts.ID_ForgotPassword Then
                UA.ForgotPassword(PageSet, Context, Context.Request.Form.Get(UserAccounts.ID_Username), Context.Request.Form.Get(UserAccounts.ID_RetrievePassword))
            ElseIf Context.Request.QueryString.Get(PageQuery) = UserAccounts.ID_ResetPassword Then
                If Context.Request.HttpMethod = "POST" Then
                    UA.ResetPassword(PageSet, Context, String.Empty, Context.Request.Form.Get(UserAccounts.ID_Username), Context.Request.Form.Get(UserAccounts.ID_PasswordResetCode), Context.Request.Form.Get(UserAccounts.ID_Password), Context.Request.Form.Get(UserAccounts.ID_ConfirmPassword), Context.Request.Form.Get(UserAccounts.ID_ResetPassword))
                Else
                    UA.ResetPassword(PageSet, Context, Context.Request.QueryString.Get(UserAccounts.ID_UserID), String.Empty, Context.Request.QueryString.Get(UserAccounts.ID_PasswordResetCode), String.Empty, String.Empty, UserAccounts.ID_ResetPassword)
                End If
            ElseIf Context.Request.QueryString.Get(PageQuery) = UserAccounts.ID_ActivateAccount Then
                If Context.Request.HttpMethod = "POST" Then
                    UA.ActivateAccount(PageSet, Context, String.Empty, Context.Request.Form.Get(UserAccounts.ID_Username), Context.Request.Form.Get(UserAccounts.ID_ActivationCode), Context.Request.Form.Get(UserAccounts.ID_ActivateAccount))
                Else
                    UA.ActivateAccount(PageSet, Context, Context.Request.QueryString.Get(UserAccounts.ID_UserID), String.Empty, Context.Request.QueryString.Get(UserAccounts.ID_ActivationCode), UserAccounts.ID_ActivateAccount)
                End If
            ElseIf Context.Request.QueryString.Get(PageQuery) = UserAccounts.ID_SendActivationCode Then
                UA.SendActivation(PageSet, Context, Context.Request.Form.Get(UserAccounts.ID_Username), Context.Request.Form.Get(UserAccounts.ID_SendActivationCode))
            ElseIf Context.Request.QueryString.Get(PageQuery) = UserAccounts.ID_ControlPanel Then
                UA.ControlPanel(PageSet, Context)
            ElseIf Context.Request.QueryString.Get(PageQuery) = UserAccounts.ID_DeleteAccount Then
                UA.DeleteAccount(PageSet, Context, Context.Request.Form.Get(UserAccounts.ID_Certain), Context.Request.Form.Get(UserAccounts.ID_DeleteAccount))
            ElseIf Context.Request.QueryString.Get(PageQuery) = UserAccounts.ID_ChangeEMailAddress Then
                UA.ChangeEMailAddress(PageSet, Context, Context.Request.Form.Get(UserAccounts.ID_EmailAddress), Context.Request.Form.Get(UserAccounts.ID_ConfirmEmailAddress), Context.Request.Form.Get(UserAccounts.ID_ChangeEMailAddress))
            ElseIf Context.Request.QueryString.Get(PageQuery) = UserAccounts.ID_ChangePassword Then
                UA.ChangePassword(PageSet, Context, Context.Request.Form.Get(UserAccounts.ID_Password), Context.Request.Form.Get(UserAccounts.ID_ConfirmPassword), Context.Request.Form.Get(UserAccounts.ID_ChangePassword))
            ElseIf Context.Request.QueryString.Get(PageQuery) = UserAccounts.ID_ChangeUsername Then
                UA.ChangeUserName(PageSet, Context, Context.Request.Form.Get(UserAccounts.ID_Username), Context.Request.Form.Get(UserAccounts.ID_ChangeUsername))
            ElseIf Context.Request.QueryString.Get(PageQuery) = UserAccounts.ID_CertInstall AndAlso UA.IsAdmin(Context) Then
                UA.UploadCertificate(PageSet, Context, Context.Request.Form.Get(UserAccounts.ID_CertInstall), Context.Request.Form.Get(UserAccounts.ID_Certificate), Context.Request.Form.Get(UserAccounts.ID_CertRequest))
            ElseIf Context.Request.QueryString.Get(PageQuery) = UserAccounts.ID_CertInstallIntermed AndAlso UA.IsAdmin(Context) Then
                UA.InstallIntermediateCert(PageSet, Context, Context.Request.Form.Get(UserAccounts.ID_CertInstallIntermed), Context.Request.Form.Get(UserAccounts.ID_Certificate))
            ElseIf Context.Request.QueryString.Get(PageQuery) = UserAccounts.ID_CertRequest AndAlso UA.IsAdmin(Context) Then
                UA.CreateCertificateRequest(PageSet, Context, Context.Request.Form.Get(UserAccounts.ID_CertRequest), Context.Request.Form.Get(UserAccounts.ID_PrivateKey))
            ElseIf Context.Request.QueryString.Get(PageQuery) = UserAccounts.ID_DeleteCertRequest AndAlso UA.IsAdmin(Context) Then
                UA.DeleteCertificateRequest(PageSet, Context, Context.Request.Form.Get(UserAccounts.ID_DeleteCertRequest), Context.Request.Form.Get(UserAccounts.ID_Certificate))
            End If
            _IsHtml = True
            Index = PageSet.GetPageIndex(Context.Request.Params(If(IsPrint, PagePrintQuery, PageQuery)))
            If Not IsPrint Then
                Controls.Add(New Menu(_PortableMethods, UA, PageSet, Index))
            End If
            Controls.Add(New Page(_PortableMethods, UWeb, ArbData, PageSet.Pages.Item(Index), True, IsPrint))
            Await CType(Controls(Controls.Count - 1), Page).DoGetAsyncData()
            Context.Response.ContentType = "text/html;charset=" + System.Text.Encoding.UTF8.WebName
        End If
    End Function
    Protected Overrides Sub Render(ByVal writer As System.Web.UI.HtmlTextWriter)
        Dim Count As Integer
        If _IsHtml Then
            Dim StringWriter As New IO.StringWriter
            Dim NewWriter As New System.Web.UI.HtmlTextWriter(StringWriter)
            RenderChildren(NewWriter)
            writer.Write("<!DOCTYPE HTML PUBLIC " + HtmlTextWriter.DoubleQuoteChar + "-//W3C//DTD HTML 4.01//EN" + HtmlTextWriter.DoubleQuoteChar + " " + HtmlTextWriter.DoubleQuoteChar + "http://www.w3.org/TR/html4/strict.dtd" + HtmlTextWriter.DoubleQuoteChar + ">" + vbCrLf)
            writer.WriteBeginTag("html")
            writer.WriteAttribute("xmlns", "http://www.w3.org/1999/xhtml")
            writer.WriteAttribute("prefix", "og: http://ogp.me/ns# fb: http://ogp.me/ns/fb# article: http://ogp.me/ns/article#")
            writer.Write(HtmlTextWriter.TagRightChar)
            writer.Write(vbCrLf + vbTab)
            writer.WriteFullBeginTag("head")
            writer.Write("<meta http-equiv='Content-type' content='text/html;charset=" + System.Text.Encoding.UTF8.WebName + "'>")
            writer.Write("<meta property=""og:title"" content=""" + CType(Controls(Controls.Count - 1), Page).GetTitle() + """>")
            writer.Write("<meta property=""og:site_name"" content=""" + _PortableMethods.LoadResourceString(PageSet.Title) + """>")
            writer.Write("<meta property=""og:url"" content=""" + UtilityWeb.HtmlTextEncode(HttpContext.Current.Request.Url.AbsoluteUri) + """>")
            writer.Write("<meta property=""og:description"" content=""" + CType(Controls(Controls.Count - 1), Page).GetDescription() + """>")
            writer.Write("<meta property=""og:image"" content=""" + Web.HttpContext.Current.Request.Url.GetLeftPart(UriPartial.Authority) + UtilityWeb.HtmlTextEncode(GetPageString("Image.gif&Image=Scale&p=menu.main")) + """>")
            writer.Write("<meta property=""fb:app_id"" content=""" + UtilityWeb.ConnectionData.FBAppID + """>")
            writer.Write("<meta property=""og:type"" content=""article"">")
            writer.Write("<meta property=""og:locale"" content=""en_US"">")
            writer.Write("<meta property=""article:author"" content=""" + UtilityWeb.ConnectionData.AuthorName + """>")
            writer.Write("<meta property=""article:publisher"" content=""" + UtilityWeb.ConnectionData.AuthorName + """>")
            writer.WriteFullBeginTag("title")
            writer.Write(_PortableMethods.LoadResourceString(PageSet.Title))
            writer.WriteEndTag("title")
            If CType(Controls(Controls.Count - 1), Page).JSFunctions.Count <> 0 Then
                writer.Write(vbCrLf + vbTab + vbTab)
                writer.WriteBeginTag("script")
                writer.WriteAttribute("type", "text/javascript")
                writer.WriteAttribute("src", "Scripts/jquery-1.11.3.min.js")
                writer.Write(HtmlTextWriter.TagRightChar + vbCrLf + vbTab + vbTab)
                writer.WriteEndTag("script")
                writer.Write(vbCrLf + vbTab + vbTab)

                writer.WriteBeginTag("script")
                writer.WriteAttribute("type", "text/javascript")
                writer.WriteAttribute("src", "Scripts/jquery-ui-1.10.4.min.js")
                writer.Write(HtmlTextWriter.TagRightChar + vbCrLf + vbTab + vbTab)
                writer.WriteEndTag("script")
                writer.Write(vbCrLf + vbTab + vbTab)
                writer.WriteBeginTag("script")
                writer.WriteAttribute("type", "text/javascript")
                writer.WriteAttribute("src", "Scripts/jquery.plugin.min.js")
                writer.Write(HtmlTextWriter.TagRightChar + vbCrLf + vbTab + vbTab)
                writer.WriteEndTag("script")
                writer.Write(vbCrLf + vbTab + vbTab)
                writer.WriteBeginTag("script")
                writer.WriteAttribute("type", "text/javascript")
                writer.WriteAttribute("src", "Scripts/jquery.calendars.all.min.js")
                writer.Write(HtmlTextWriter.TagRightChar + vbCrLf + vbTab + vbTab)
                writer.WriteEndTag("script")
                writer.Write(vbCrLf + vbTab + vbTab)
                writer.WriteBeginTag("script")
                writer.WriteAttribute("type", "text/javascript")
                writer.WriteAttribute("src", "Scripts/jquery.calendars.picker.lang.min.js")
                writer.Write(HtmlTextWriter.TagRightChar + vbCrLf + vbTab + vbTab)
                writer.WriteEndTag("script")
                writer.Write(vbCrLf + vbTab + vbTab)
                writer.WriteBeginTag("script")
                writer.WriteAttribute("type", "text/javascript")
                writer.WriteAttribute("src", "Scripts/jquery.calendars.islamic.min.js")
                writer.Write(HtmlTextWriter.TagRightChar + vbCrLf + vbTab + vbTab)
                writer.WriteEndTag("script")
                writer.Write(vbCrLf + vbTab + vbTab)
                writer.WriteBeginTag("script")
                writer.WriteAttribute("type", "text/javascript")
                writer.WriteAttribute("src", "Scripts/jquery.calendars.ummalqura.min.js")
                writer.Write(HtmlTextWriter.TagRightChar + vbCrLf + vbTab + vbTab)
                writer.WriteEndTag("script")
                writer.Write(vbCrLf + vbTab + vbTab)
                writer.WriteBeginTag("link")
                writer.WriteAttribute("rel", "stylesheet")
                writer.WriteAttribute("type", "text/css")
                writer.WriteAttribute("href", "Content/jquery.calendars.picker.css")
                writer.Write(HtmlTextWriter.TagRightChar + vbCrLf + vbTab + vbTab)

                writer.WriteBeginTag("script")
                writer.WriteAttribute("type", "text/javascript")
                writer.Write(HtmlTextWriter.TagRightChar)
                For Count = 0 To CType(Controls(Controls.Count - 1), Page).JSFunctions.Count - 1
                    writer.Write(vbCrLf + vbTab + vbTab + vbTab)
                    writer.Write(CType(Controls(Controls.Count - 1), Page).JSFunctions(Count))
                Next
                writer.Write(vbCrLf + vbTab + vbTab)
                writer.WriteEndTag("script")
            End If
            writer.Write(vbCrLf + vbTab + vbTab)
            writer.WriteBeginTag("link")
            writer.WriteAttribute("href", "Content/Styles.css")
            writer.WriteAttribute("rel", "stylesheet")
            writer.WriteAttribute("type", "text/css")
            writer.Write(HtmlTextWriter.TagRightChar + vbCrLf + vbTab)
            writer.WriteEndTag("head")
            writer.Write(vbCrLf + vbTab)
            writer.WriteBeginTag("body")
            If CType(Controls(Controls.Count - 1), Page).JSInitFuncs.Count <> 0 Then
                Dim LoadJS As String = String.Empty
                For Count = 0 To CType(Controls(Controls.Count - 1), Page).JSInitFuncs.Count - 1
                    LoadJS += CType(Controls(Controls.Count - 1), Page).JSInitFuncs(Count)
                Next
                writer.WriteAttribute("onload", LoadJS)
            End If
            If Controls.Count = 2 Then
                writer.Write(HtmlTextWriter.TagRightChar + vbCrLf + vbTab + vbTab)
                writer.WriteBeginTag("div")
                writer.WriteAttribute("id", "fb-root")
                writer.Write(HtmlTextWriter.TagRightChar + vbCrLf + vbTab + vbTab)
                writer.WriteEndTag("div")
                writer.Write(vbCrLf + vbTab + vbTab)
                writer.WriteBeginTag("script")
                writer.WriteAttribute("type", "text/javascript")
                writer.Write(HtmlTextWriter.TagRightChar)
                writer.Write("(function(d, s, id) {" +
    "  var js, fjs = d.getElementsByTagName(s)[0];" +
    "  if (d.getElementById(id)) return;" +
    "  js = d.createElement(s); js.id = id;" +
    "  js.src = ""//connect.facebook.net/en_US/all.js#xfbml=1&appId=" + UtilityWeb.ConnectionData.FBAppID + """;" +
    "  fjs.parentNode.insertBefore(js, fjs);" +
    "}(document, 'script', 'facebook-jssdk'));")
                writer.Write(vbCrLf + vbTab + vbTab)
                writer.WriteEndTag("script")

                writer.Write(vbCrLf + vbTab + vbTab)
                writer.WriteBeginTag("div")
                writer.WriteAttribute("id", "bgtopdiv")
                writer.Write(HtmlTextWriter.TagRightChar + vbCrLf + vbTab + vbTab + vbTab)
                writer.WriteBeginTag("img")
                writer.WriteAttribute("alt", String.Empty)
                writer.WriteAttribute("width", "100%")
                writer.WriteAttribute("height", "100%")
                writer.WriteAttribute("src", UtilityWeb.HtmlTextEncode(GetPageString("Image.gif&Image=GradientBackground")))
                writer.Write(HtmlTextWriter.TagRightChar + vbCrLf + vbTab + vbTab)
                writer.WriteEndTag("div")
                writer.Write(vbCrLf + vbTab + vbTab)
                writer.WriteBeginTag("div")
                writer.WriteAttribute("id", "maindiv")
                writer.Write(HtmlTextWriter.TagRightChar)
            End If
            writer.Write(StringWriter.ToString())
            If Controls.Count = 2 Then
                writer.WriteEndTag("div")

                writer.Write(vbCrLf + vbTab + vbTab)
                writer.WriteBeginTag("div")
                writer.WriteAttribute("id", "bgbottomdiv")
                writer.Write(HtmlTextWriter.TagRightChar + vbCrLf + vbTab + vbTab + vbTab)
                writer.WriteBeginTag("img")
                writer.WriteAttribute("alt", String.Empty)
                writer.WriteAttribute("width", "100%")
                writer.WriteAttribute("height", "100%")
                writer.WriteAttribute("src", UtilityWeb.HtmlTextEncode(GetPageString("Image.gif&Image=GradientBackgroundBottom")))
                writer.Write(HtmlTextWriter.TagRightChar + vbCrLf + vbTab + vbTab)
                writer.WriteEndTag("div")

                writer.Write(vbCrLf + vbTab + vbTab)
                writer.WriteBeginTag("div")
                writer.WriteAttribute("class", "fb-like")
                writer.WriteAttribute("data-href", UtilityWeb.HtmlTextEncode(HttpContext.Current.Request.Url.AbsoluteUri))
                writer.WriteAttribute("data-layout", "standard")
                writer.WriteAttribute("data-action", "like")
                writer.WriteAttribute("data-show-faces", "true")
                writer.WriteAttribute("data-share", "true")
                writer.Write(HtmlTextWriter.TagRightChar + vbCrLf + vbTab + vbTab + vbTab)
                writer.WriteEndTag("div")
            End If

            writer.Write(vbCrLf + vbTab)
            writer.WriteEndTag("body")
            writer.Write(vbCrLf)
            writer.WriteEndTag("html")
        End If
    End Sub
End Class
Public Class UserAccounts
    Public Const ID_MinimumUserName As String = "MinimumUsername"
    Public Const ID_MinimumPassword As String = "MinimumPassword"
    Public Const ID_BadEmailFormat As String = "BadEmailFormat"
    Public Const ID_Login As String = "Login"
    Public Const ID_Logoff As String = "Logoff"
    Public Const ID_UserID As String = "UserID"
    Public Const ID_Secret As String = "Secret"
    Public Const ID_Username As String = "Username"
    Public Const ID_Password As String = "Password"
    Public Const ID_Remember As String = "Remember"
    Public Const ID_ConfirmPassword As String = "ConfirmPassword"
    Public Const ID_EmailAddress As String = "EmailAddress"
    Public Const ID_ConfirmEmailAddress As String = "ConfirmEmailAddress"
    Public Const ID_Home As String = "Home"
    Public Const ID_Register As String = "Register"
    Public Const ID_Error As String = "Error"
    Public Const ID_PasswordsNotMatch As String = "PasswordsNotMatch"
    Public Const ID_EmailsNotMatch As String = "EmailsNotMatch"
    Public Const ID_DuplicateUser As String = "DuplicateUser"
    Public Const ID_EmailInUse As String = "EmailInUse"
    Public Const ID_InvalidUsernamePassword As String = "InvalidUsernamePassword"
    Public Const ID_AccountNotActivated As String = "AccountNotActivated"
    Public Const ID_ForgotUsername As String = "ForgotUsername"
    Public Const ID_ForgotPassword As String = "ForgotPassword"
    Public Const ID_ResetPassword As String = "ResetPassword"
    Public Const ID_ActivateAccount As String = "ActivateAccount"
    Public Const ID_ControlPanel As String = "ControlPanel"
    Public Const ID_LoginSuccess As String = "LoginSuccess"
    Public Const ID_LogoffSuccess As String = "LogoffSuccess"
    Public Const ID_ChangeUsername As String = "ChangeUsername"
    Public Const ID_ChangePassword As String = "ChangePassword"
    Public Const ID_DeleteAccount As String = "DeleteAccount"
    Public Const ID_ChangeEMailAddress As String = "ChangeEMailAddress"
    Public Const ID_RetrieveUsername As String = "RetrieveUsername"
    Public Const ID_RetrievePassword As String = "RetrievePassword"
    Public Const ID_NoUserExists As String = "NoUserExists"
    Public Const ID_PasswordAlreadyReset As String = "PasswordAlreadyReset"
    Public Const ID_UserNotFound As String = "UserNotFound"
    Public Const ID_InvalidInput As String = "InvalidInput"
    Public Const ID_IncorrectPasswordResetCode As String = "IncorrectPasswordResetCode"
    Public Const ID_PasswordResetSuccess As String = "PasswordResetSuccess"
    Public Const ID_SendActivationCode As String = "SendActivationCode"
    Public Const ID_AccountAlreadyActivated As String = "AccountAlreadyActivated"
    Public Const ID_IncorrectActivationCode As String = "IncorrectActivationCode"
    Public Const ID_EmailChangeReactivate As String = "EmailChangeReactivate"
    Public Const ID_AccountActivateSuccess As String = "AccountActivateSuccess"
    Public Const ID_AccountDeleteSuccess As String = "AccountDeleteSuccess"
    Public Const ID_CheckEmailNotice As String = "CheckEmailNotice"
    Public Const ID_UsernameChangeSuccess As String = "UsernameChangeSuccess"
    Public Const ID_PasswordChangeSuccess As String = "PasswordChangeSuccess"
    Public Const ID_NotConfirmDelete As String = "NotConfirmDelete"
    Public Const ID_ClearCache As String = "ClearCache"
    Public Const ID_ViewHeaders As String = "ViewHeaders"
    Public Const ID_DotNetVersion As String = "DotNetVersion"
    Public Const ID_ViewCache As String = "ViewCache"
    Public Const ID_ClearDiskCache As String = "ClearDiskCache"
    Public Const ID_ViewDiskCache As String = "ViewDiskCache"
    Public Const ID_CreateDatabase As String = "CreateDatabase"
    Public Const ID_RemoveDatabase As String = "RemoveDatabase"
    Public Const ID_CleanupState As String = "CleanupStale"
    Public Const ID_CertRequest As String = "CertRequest"
    Public Const ID_PrivateKey As String = "PrivateKey"
    Public Const ID_CertInstall As String = "CertInstall"
    Public Const ID_CertInstallIntermed As String = "CertInstallIntermed"
    Public Const ID_Certificate As String = "Certificate"
    Public Const ID_DeleteCertRequest As String = "DeleteCertRequest"
    Public Const ID_ViewCertRequests As String = "ViewCertRequests"
    Public Const ID_ViewSites As String = "ViewSites"
    Public Const ID_EnableSSL As String = "EnableSSL"
    Public Const ID_CurrentUser As String = "CurrentUser"
    Public Const ID_HadithRanking As String = "HadithRanking"
    Public Const ID_PasswordResetCode As String = "PasswordResetCode"
    Public Const ID_ActivationCode As String = "ActivationCode"
    Public Const ID_ResendActivationEmail As String = "ResendActivationEmail"
    Public Const ID_PasswordEmailSent As String = "PasswordEmailSent"
    Public Const ID_UsernameEmailSent As String = "UsernameEmailSent"
    Public Const ID_Yes As String = "Yes"
    Public Const ID_No As String = "No"
    Public Const ID_Certain As String = "Certain"
    Public Const ID_Lang As String = "Lang"
    Public Const ID_LangID As String = "LangID"
    Private _PortableMethods As PortableMethods
    Private UWeb As UtilityWeb
    Private ArbData As ArabicData
    Private SD As SiteDatabase
    Private PL As PageLoader
    Private MD As MailDispatcher
    Public Sub New(NewPortableMethods As PortableMethods, NewUWeb As UtilityWeb, NewArbData As ArabicData, NewPL As PageLoader, NewSD As SiteDatabase)
        _PortableMethods = NewPortableMethods
        UWeb = NewUWeb
        ArbData = NewArbData
        SD = NewSD
        PL = NewPL
        MD = New MailDispatcher(_PortableMethods, UWeb)
    End Sub
    Private Shared Function ValidateUserName(ByVal UserName As String) As String
        If UserName.Length < 4 Then
            Return "Acct_" + ID_MinimumUserName
        End If
        'allowable character set?
        Return String.Empty
    End Function
    Private Shared Function ValidatePassword(ByVal Password As String) As String
        If Password.Length < 8 Then
            Return "Acct_" + ID_MinimumPassword
        End If
        'allowable character set?
        Return String.Empty
    End Function
    Private Shared Function ValidateEMail(ByVal EMail As String) As String
        If Not New UtilityWeb.EMailValidator().IsValidEMail(EMail) Then
            Return "Acct_" + ID_BadEmailFormat
        End If
        Return String.Empty
    End Function
    Public Function IsAdmin(Context As HttpContext) As Boolean
        Return IsLoggedIn(Context) And SD.CheckAccess(GetUserID(Context)) = 1
    End Function
    Public Shared Function GetLanguage(Context As HttpContext) As String
        If Array.IndexOf(Context.Response.Cookies.AllKeys, ID_Lang) <> -1 Then
            Return Context.Response.Cookies.Get(ID_Lang).Item(ID_LangID)
        ElseIf Array.IndexOf(Context.Request.Cookies.AllKeys, ID_Lang) <> -1 Then
            Return Context.Request.Cookies.Get(ID_Lang).Item(ID_LangID)
        Else
            Return String.Empty
        End If
    End Function
    Public Shared Sub SetLanguage(LangID As String, Context As HttpContext)
        Dim Cookie As New HttpCookie(ID_Lang)
        Cookie.Item(ID_LangID) = LangID
        Context.Response.Cookies.Set(Cookie)
    End Sub
    Public Function IsLoggedIn(Context As HttpContext) As Boolean
        If Array.IndexOf(Context.Response.Cookies.AllKeys, ID_Login) <> -1 AndAlso
                Context.Response.Cookies.Get(ID_Login).Item(ID_UserID) <> String.Empty Then Return True
        Dim UserID As Integer = GetUserID(Context)
        If UserID = -1 Then Return False
        Return SD.CheckLogin(UserID,
            CInt(Context.Request.Cookies.Get(ID_Login).Item(ID_Secret)))
    End Function
    Public Function GetUserName(Context As HttpContext) As String
        Return SD.GetUserName(GetUserID(Context))
    End Function
    Public Shared Function GetUserID(Context As HttpContext) As Integer
        If Context.Request.Cookies.Get(ID_Login) Is Nothing Then Return -1
        Return CInt(Context.Request.Cookies.Get(ID_Login).Item(ID_UserID))
    End Function
    Public Sub SetLoginCookie(ByVal UserID As Integer, ByVal Persist As Boolean, Context As HttpContext)
        Dim Secret As Integer = SD.SetLogin(UserID, Persist)
        Dim Cookie As New HttpCookie(ID_Login)
        If Not Persist Then Cookie.Expires = Now.AddHours(1)
        Cookie.Item(ID_UserID) = CStr(UserID)
        Cookie.Item(ID_Secret) = CStr(Secret)
        Context.Response.Cookies.Set(Cookie)
    End Sub
    Public Sub ClearLoginCookie(Context As HttpContext)
        SD.ClearLogin(GetUserID(Context))
        Dim Cookie As New HttpCookie(ID_Login)
        Cookie.Expires = DateTime.Now.AddDays(-1D)
        Cookie.Item(ID_UserID) = String.Empty
        Cookie.Item(ID_Secret) = String.Empty
        Context.Response.Cookies.Set(Cookie)
    End Sub
    Public Sub Register(PageSet As PageLoader, Context As HttpContext, ByVal UserName As String, ByVal Password As String, ByVal PasswordConfirm As String, ByVal EMail As String, ByVal EMailConfirm As String, ByVal Register As String)
        Dim Errors As New Generic.List(Of String)
        Dim UserID As Integer
        If IsLoggedIn(Context) Then Return
        If Register = _PortableMethods.LoadResourceString("Acct_" + ID_Register) Then
            Dim ErrorString As String = ValidateUserName(UserName)
            If ErrorString <> String.Empty Then
                Errors.Add(ErrorString)
            End If
            ErrorString = ValidatePassword(Password)
            If ErrorString <> String.Empty Then
                Errors.Add(ErrorString)
            End If
            ErrorString = ValidateEMail(EMail)
            If ErrorString <> String.Empty Then
                Errors.Add(ErrorString)
            End If
            If Password <> PasswordConfirm Then
                Errors.Add("Acct_" + ID_PasswordsNotMatch)
            End If
            If EMail <> EMailConfirm Then
                Errors.Add("Acct_" + ID_EmailsNotMatch)
            End If
            If Errors.Count = 0 Then
                If SD.GetUserID(UserName) <> -1 Then
                    Errors.Add("Acct_" + ID_DuplicateUser)
                ElseIf SD.GetUserIDByEMail(EMail) <> -1 Then
                    Errors.Add("Acct_" + ID_EmailInUse)
                Else
                    SD.AddUser(UserName, Password, EMail)
                    UserID = SD.GetUserID(UserName)
                    MD.SendActivationEMail(UserName, EMail, UserID, SD.GetUserActivated(UserID), Context)
                End If
            End If
        End If
        Dim Controls As New Generic.List(Of Object)
        Dim Form As New Generic.List(Of Object)
        If Register <> _PortableMethods.LoadResourceString("Acct_" + ID_Register) Or Errors.Count <> 0 Then
            'check form first otherwise send new form
            Controls.AddRange(Linq.Enumerable.Select(Errors.ToArray(), Function(ErrorID As String) CType(New PageLoader.TextItem(ID_Error, ErrorID), Object)))
            Form.Add(New PageLoader.TextItem(String.Empty, "Acct_" + ID_Username))
            Form.Add(New PageLoader.EditItem(ID_Username, UserName, 1))
            Form.Add(New PageLoader.TextItem(String.Empty, "Acct_" + ID_Password))
            Form.Add(New PageLoader.EditItem(ID_Password, String.Empty, 1, True))
            Form.Add(New PageLoader.TextItem(String.Empty, "Acct_" + ID_ConfirmPassword))
            Form.Add(New PageLoader.EditItem(ID_ConfirmPassword, String.Empty, 1, True))
            Form.Add(New PageLoader.TextItem(String.Empty, "Acct_" + ID_EmailAddress))
            Form.Add(New PageLoader.EditItem(ID_EmailAddress, EMail, 1))
            Form.Add(New PageLoader.TextItem(String.Empty, "Acct_" + ID_ConfirmEmailAddress))
            Form.Add(New PageLoader.EditItem(ID_ConfirmEmailAddress, EMailConfirm, 1))
            Form.Add(New PageLoader.ButtonItem(ID_Register, "Acct_" + ID_Register))
            Controls.Add(New PageLoader.ListItem("Acct_" + ID_Register, ID_Register, Form, False, True, host.GetPageString(ID_Register)))
        Else
            Controls.Add(New PageLoader.TextItem(String.Empty, "Acct_" + ID_CheckEmailNotice))
            Controls.Add(New PageLoader.TextItem("Acct_" + ID_Home, ID_Home, host.RootPath))
        End If
        PageSet.Pages.Add(New PageLoader.PageItem(Controls, ID_Register, "Acct_" + ID_Register))
    End Sub
    Public Sub Login(PageSet As PageLoader, Context As HttpContext, ByVal UserName As String, ByVal Password As String, ByVal Persist As String, ByVal Login As String)
        Dim Errors As New Generic.List(Of String)
        Dim UserID As Integer
        Dim Forgot As Boolean = False
        Dim NoActivate As Boolean = False
        If IsLoggedIn(Context) Then Return
        If Login = _PortableMethods.LoadResourceString("Acct_" + ID_Login) Then
            Dim ErrorString As String = ValidateUserName(UserName)
            If ErrorString <> String.Empty Then
                Errors.Add(ErrorString)
            End If
            ErrorString = ValidatePassword(Password)
            If ErrorString <> String.Empty Then
                Errors.Add(ErrorString)
            End If
            If Errors.Count = 0 Then
                UserID = SD.GetUserID(UserName, Password)
                If UserID = -1 Then
                    Errors.Add("Acct_" + ID_InvalidUsernamePassword)
                    Forgot = True
                Else
                    If SD.GetUserActivated(UserID) <> -1 Then
                        Errors.Add("Acct_" + ID_AccountNotActivated)
                        NoActivate = True
                    Else
                        SetLoginCookie(UserID, Persist = "0", Context)
                    End If
                End If
            End If
        End If
        Dim Controls As New Generic.List(Of Object)
        Dim Form As New Generic.List(Of Object)
        If Login <> _PortableMethods.LoadResourceString("Acct_" + ID_Login) Or Errors.Count <> 0 Then
            If Errors.Count <> 0 Then
                Controls.AddRange(Linq.Enumerable.Select(Errors.ToArray(), Function(ErrorID As String) CType(New PageLoader.TextItem(ID_Error, ErrorID), Object)))
                If Forgot Then
                    Controls.Add(New PageLoader.TextItem(ID_ForgotUsername, "Acct_" + ID_ForgotUsername, host.GetPageString(ID_ForgotUsername)))
                    Controls.Add(New PageLoader.TextItem(ID_ForgotPassword, "Acct_" + ID_ForgotPassword, host.GetPageString(ID_ForgotPassword)))
                ElseIf NoActivate Then
                    Controls.Add(New PageLoader.TextItem(ID_ResendActivationEmail, "Acct_" + ID_ResendActivationEmail, host.GetPageString(ID_SendActivationCode)))
                End If
            End If
            Form.Add(New PageLoader.TextItem(String.Empty, "Acct_" + ID_Username))
            Form.Add(New PageLoader.EditItem(ID_Username, UserName, 1))
            Form.Add(New PageLoader.TextItem(String.Empty, "Acct_" + ID_Password))
            Form.Add(New PageLoader.EditItem(ID_Password, String.Empty, 1, True))
            Form.Add(New PageLoader.RadioItem(ID_Remember, "Acct_" + ID_Remember, CStr(If(Persist = String.Empty, "1", Persist)), New PageLoader.OptionItem() {New PageLoader.OptionItem("Acct_" + ID_Yes), New PageLoader.OptionItem("Acct_" + ID_No)}))
            Form.Add(New PageLoader.ButtonItem(ID_Login, "Acct_" + ID_Login))
            Controls.Add(New PageLoader.ListItem("Acct_" + ID_Login, ID_Login, Form, False, True, host.GetPageString(ID_Login)))
            Controls.Add(New PageLoader.TextItem(ID_Register, "Acct_" + ID_Register, host.GetPageString(ID_Register)))
        Else
            Controls.Add(New PageLoader.TextItem(String.Empty, "Acct_" + ID_LoginSuccess))
            Controls.Add(New PageLoader.TextItem(ID_Home, "Acct_" + ID_Home, host.RootPath))
        End If
        PageSet.Pages.Add(New PageLoader.PageItem(Controls, ID_Login, "Acct_" + ID_Login))
    End Sub
    Public Sub Logoff(PageSet As PageLoader, Context As HttpContext)
        If Not IsLoggedIn(Context) Then Return
        ClearLoginCookie(Context)
        Dim Controls As New Generic.List(Of Object)
        Controls.Add(New PageLoader.TextItem(String.Empty, "Acct_" + ID_LogoffSuccess))
        Controls.Add(New PageLoader.TextItem(ID_Home, "Acct_" + ID_Home, host.RootPath))
        PageSet.Pages.Add(New PageLoader.PageItem(Controls, ID_Logoff, "Acct_" + ID_Logoff))
    End Sub
    Public Sub ForgotUserName(PageSet As PageLoader, Context As HttpContext, ByVal EMail As String, ByVal RetrieveUserName As String)
        Dim Errors As New Generic.List(Of String)
        If IsLoggedIn(Context) Then Return
        If RetrieveUserName = _PortableMethods.LoadResourceString("Acct_" + ID_RetrieveUsername) Then
            Dim UserID As Integer
            Dim ErrorString As String = ValidateEMail(EMail)
            If ErrorString <> String.Empty Then Errors.Add(ErrorString)
            If Errors.Count = 0 Then
                UserID = SD.GetUserIDByEMail(EMail)
                If UserID = -1 Then
                    Errors.Add("Acct_" + ID_NoUserExists)
                Else
                    MD.SendUserNameReminderEMail(SD.GetUserName(UserID), EMail, Context)
                End If
            End If
        End If
        Dim Controls As New Generic.List(Of Object)
        Dim Form As New Generic.List(Of Object)
        If RetrieveUserName <> _PortableMethods.LoadResourceString("Acct_" + ID_RetrieveUsername) Or Errors.Count <> 0 Then
            If Errors.Count <> 0 Then
                Controls.AddRange(Linq.Enumerable.Select(Errors.ToArray(), Function(ErrorID As String) CType(New PageLoader.TextItem(ID_Error, ErrorID), Object)))
            End If
            Form.Add(New PageLoader.TextItem(String.Empty, "Acct_" + ID_EmailAddress))
            Form.Add(New PageLoader.EditItem(ID_EmailAddress, EMail, 1))
            Form.Add(New PageLoader.ButtonItem(ID_RetrieveUsername, "Acct_" + ID_RetrieveUsername))
            Controls.Add(New PageLoader.ListItem("Acct_" + ID_RetrieveUsername, ID_RetrieveUsername, Form, False, True, host.GetPageString(ID_RetrieveUsername)))
        Else
            Controls.Add(New PageLoader.TextItem(String.Empty, "Acct_" + ID_UsernameEmailSent))
            Controls.Add(New PageLoader.TextItem(ID_Home, "Acct_" + ID_Home, host.RootPath))
        End If
        PageSet.Pages.Add(New PageLoader.PageItem(Controls, ID_ForgotUsername, "Acct_" + ID_ForgotUsername))
    End Sub
    Public Sub ForgotPassword(PageSet As PageLoader, Context As HttpContext, ByVal UserName As String, ByVal RetrievePassword As String)
        Dim Errors As New Generic.List(Of String)
        Dim UserID As Integer
        If IsLoggedIn(Context) Then Return
        If RetrievePassword = _PortableMethods.LoadResourceString("Acct_" + ID_RetrievePassword) Then
            Dim ErrorString As String = ValidateUserName(UserName)
            If ErrorString <> String.Empty Then Errors.Add(ErrorString)
            If Errors.Count = 0 Then
                UserID = SD.GetUserID(UserName)
                If UserID = -1 Then
                    Errors.Add("Acct_" + ID_UserNotFound)
                Else
                    MD.SendPasswordResetEMail(UserName, SD.GetUserEMail(UserID), UserID, SD.GetUserResetCode(UserID), Context) 'secret code is crc32 of password hash
                End If
            End If
        End If
        Dim Controls As New Generic.List(Of Object)
        Dim Form As New Generic.List(Of Object)
        If RetrievePassword <> _PortableMethods.LoadResourceString("Acct_" + ID_RetrievePassword) Or Errors.Count <> 0 Then
            If Errors.Count <> 0 Then
                Controls.AddRange(Linq.Enumerable.Select(Errors.ToArray(), Function(ErrorID As String) CType(New PageLoader.TextItem(ID_Error, ErrorID), Object)))
            End If
            Form.Add(New PageLoader.TextItem(String.Empty, "Acct_" + ID_Username))
            Form.Add(New PageLoader.EditItem(ID_Username, UserName, 1))
            Form.Add(New PageLoader.ButtonItem(ID_RetrievePassword, "Acct_" + ID_RetrievePassword))
            Controls.Add(New PageLoader.ListItem("Acct_" + ID_RetrievePassword, ID_RetrievePassword, Form, False, True, host.GetPageString(ID_ForgotPassword)))
        Else
            Controls.Add(New PageLoader.TextItem(String.Empty, "Acct_" + ID_PasswordEmailSent))
            Controls.Add(New PageLoader.TextItem(ID_Home, "Acct_" + ID_Home, host.RootPath))
        End If
        PageSet.Pages.Add(New PageLoader.PageItem(Controls, ID_ForgotPassword, "Acct_" + ID_ForgotPassword))
    End Sub
    Public Sub ResetPassword(PageSet As PageLoader, Context As HttpContext, ByVal UserID As String, ByVal UserName As String, ByVal ResetCode As String, ByVal Password As String, ByVal PasswordConfirm As String, ByVal ResetPassword As String)
        Dim Errors As New Generic.List(Of String)
        Dim UID As Integer
        Dim ResCode As UInteger
        If IsLoggedIn(Context) Then Return
        If UserID <> String.Empty Or ResetPassword = _PortableMethods.LoadResourceString("Acct_" + ID_ResetPassword) Then
            If UserID <> String.Empty And UserName = String.Empty Then
                UID = CInt(UserID)
            ElseIf UserID = String.Empty And UserName <> String.Empty Then
                Dim ErrorString As String = ValidateUserName(UserName)
                If ErrorString <> String.Empty Then Errors.Add(ErrorString)
                If Errors.Count = 0 Then
                    UID = SD.GetUserID(UserName)
                    If UID = -1 Then Errors.Add("Acct_" + ID_UserNotFound)
                End If
            Else
                Errors.Add("Acct_" + ID_InvalidInput)
            End If
            If Errors.Count = 0 Then
                ResCode = SD.GetUserResetCode(UID)
                If ResCode = -1 Then
                    Errors.Add("Acct_" + ID_PasswordAlreadyReset)
                ElseIf ResCode = CUInt(ResetCode) Then
                    If UserID = String.Empty Then
                        Dim ErrorString As String = ValidatePassword(Password)
                        If ErrorString <> String.Empty Then Errors.Add(ErrorString)
                        If Password <> PasswordConfirm Then
                            Errors.Add("Acct_" + ID_PasswordsNotMatch)
                        End If
                        If Errors.Count = 0 Then
                            SD.ChangeUserPassword(UID, Password)
                        End If
                    Else
                        UserName = SD.GetUserName(UID)
                    End If
                Else
                    Errors.Add("Acct_" + ID_IncorrectPasswordResetCode)
                End If
            End If
        End If
        Dim Controls As New Generic.List(Of Object)
        Dim Form As New Generic.List(Of Object)
        If ResetPassword <> _PortableMethods.LoadResourceString("Acct_" + ID_ResetPassword) Or UserID <> String.Empty Or Errors.Count <> 0 Then
            If Errors.Count <> 0 Then
                Controls.AddRange(Linq.Enumerable.Select(Errors.ToArray(), Function(ErrorID As String) CType(New PageLoader.TextItem(ID_Error, ErrorID), Object)))
            End If
            Form.Add(New PageLoader.TextItem(String.Empty, "Acct_" + ID_Username))
            Form.Add(New PageLoader.EditItem(ID_Username, UserName, 1))
            Form.Add(New PageLoader.TextItem(String.Empty, "Acct_" + ID_PasswordResetCode))
            Form.Add(New PageLoader.EditItem(ID_PasswordResetCode, ResetCode, 1))
            Form.Add(New PageLoader.TextItem(String.Empty, "Acct_" + ID_Password))
            Form.Add(New PageLoader.EditItem(ID_Password, String.Empty, 1, True))
            Form.Add(New PageLoader.TextItem(String.Empty, "Acct_" + ID_ConfirmPassword))
            Form.Add(New PageLoader.EditItem(ID_ConfirmPassword, String.Empty, 1, True))
            Form.Add(New PageLoader.ButtonItem(ID_ResetPassword, "Acct_" + ID_ResetPassword))
            Controls.Add(New PageLoader.ListItem("Acct_" + ID_ResetPassword, ID_ResetPassword, Form, False, True, host.GetPageString(ID_ResetPassword)))
        Else
            Controls.Add(New PageLoader.TextItem(String.Empty, "Acct_" + ID_PasswordResetSuccess))
            Controls.Add(New PageLoader.TextItem(ID_Home, "Acct_" + ID_Home, host.RootPath))
        End If
        PageSet.Pages.Add(New PageLoader.PageItem(Controls, ID_ResetPassword, "Acct_" + ID_ResetPassword))
    End Sub
    Public Sub ActivateAccount(PageSet As PageLoader, Context As HttpContext, ByVal UserID As String, ByVal UserName As String, ByVal ActivationCode As String, ByVal ActivateAccount As String)
        Dim Errors As New Generic.List(Of String)
        Dim UID As Integer
        Dim ActivateCode As Integer
        If IsLoggedIn(Context) Then Return
        If UserID <> String.Empty Or ActivateAccount = _PortableMethods.LoadResourceString("Acct_" + ID_ActivateAccount) Then
            If UserID <> String.Empty And UserName = String.Empty Then
                UID = CInt(UserID)
            ElseIf UserID = String.Empty And UserName <> String.Empty Then
                Dim ErrorString As String = ValidateUserName(UserName)
                If ErrorString <> String.Empty Then Errors.Add(ErrorString)
                If Errors.Count = 0 Then
                    UID = SD.GetUserID(UserName)
                    If UID = -1 Then Errors.Add("Acct_" + ID_UserNotFound)
                End If
            Else
                Errors.Add("Acct_" + ID_InvalidInput)
            End If
            If Errors.Count = 0 Then
                ActivateCode = SD.GetUserActivated(UID)
                If ActivateCode = -1 Then
                    Errors.Add("Acct_" + ID_AccountAlreadyActivated)
                ElseIf ActivateCode = CInt(ActivationCode) Then
                    SD.SetUserActivated(UID)
                Else
                    Errors.Add("Acct_" + ID_IncorrectActivationCode)
                End If
            End If
        End If
        Dim Controls As New Generic.List(Of Object)
        Dim Form As New Generic.List(Of Object)
        If (UserID = String.Empty And ActivateAccount <> _PortableMethods.LoadResourceString("Acct_" + ID_ActivateAccount)) Or Errors.Count <> 0 Then
            If Errors.Count <> 0 Then
                Controls.AddRange(Linq.Enumerable.Select(Errors.ToArray(), Function(ErrorID As String) CType(New PageLoader.TextItem(ID_Error, ErrorID), Object)))
            End If
            Form.Add(New PageLoader.TextItem(String.Empty, "Acct_" + ID_Username))
            Form.Add(New PageLoader.EditItem(ID_Username, UserName, 1))
            Form.Add(New PageLoader.TextItem(String.Empty, "Acct_" + ID_ActivationCode))
            Form.Add(New PageLoader.EditItem(ID_ActivationCode, ActivationCode, 1))
            Form.Add(New PageLoader.ButtonItem(ID_ActivateAccount, "Acct_" + ID_ActivateAccount))
            Controls.Add(New PageLoader.ListItem("Acct_" + ID_ActivateAccount, ID_ActivateAccount, Form, False, True, host.GetPageString(ID_ActivateAccount)))
        Else
            Controls.Add(New PageLoader.TextItem(String.Empty, "Acct_" + ID_AccountActivateSuccess))
            Controls.Add(New PageLoader.TextItem(ID_Home, "Acct_" + ID_Home, host.RootPath))
        End If
        PageSet.Pages.Add(New PageLoader.PageItem(Controls, ID_ActivateAccount, "Acct_" + ID_ActivateAccount))
    End Sub
    Public Sub SendActivation(PageSet As PageLoader, Context As HttpContext, ByVal UserName As String, ByVal SendActivation As String)
        Dim Errors As New Generic.List(Of String)
        Dim UserID As Integer
        If IsLoggedIn(Context) Then Return
        If SendActivation = _PortableMethods.LoadResourceString("Acct_" + ID_SendActivationCode) Then
            Dim ErrorString As String = ValidateUserName(UserName)
            If ErrorString <> String.Empty Then Errors.Add(ErrorString)
            If Errors.Count = 0 Then
                UserID = SD.GetUserID(UserName)
                If UserID = -1 Then
                    Errors.Add("Acct_" + ID_UserNotFound)
                Else
                    MD.SendActivationEMail(UserName, SD.GetUserEMail(UserID), UserID, SD.GetUserActivated(UserID), Context)
                End If
            End If
        End If
        Dim Controls As New Generic.List(Of Object)
        Dim Form As New Generic.List(Of Object)
        If SendActivation <> _PortableMethods.LoadResourceString("Acct_" + ID_SendActivationCode) Or Errors.Count <> 0 Then
            If Errors.Count <> 0 Then
                Controls.AddRange(Linq.Enumerable.Select(Errors.ToArray(), Function(ErrorID As String) CType(New PageLoader.TextItem(ID_Error, ErrorID), Object)))
            End If
            Form.Add(New PageLoader.TextItem(String.Empty, "Acct_" + ID_Username))
            Form.Add(New PageLoader.EditItem(ID_Username, UserName, 1))
            Form.Add(New PageLoader.ButtonItem(ID_SendActivationCode, "Acct_" + ID_SendActivationCode))
            Controls.Add(New PageLoader.ListItem("Acct_" + ID_SendActivationCode, ID_SendActivationCode, Form, False, True, host.GetPageString(ID_SendActivationCode)))
        Else
            Controls.Add(New PageLoader.TextItem(String.Empty, "Acct_" + ID_CheckEmailNotice))
            Controls.Add(New PageLoader.TextItem(ID_Home, "Acct_" + ID_Home, host.RootPath))
        End If
        PageSet.Pages.Add(New PageLoader.PageItem(Controls, ID_SendActivationCode, "Acct_" + ID_SendActivationCode))
    End Sub
    Public Sub ControlPanel(PageSet As PageLoader, Context As HttpContext)
        Dim Controls As New Generic.List(Of Object)
        Controls.Add(New PageLoader.TextItem(ID_ChangeUsername, "Acct_" + ID_ChangeUsername, host.GetPageString(ID_ChangeUsername), String.Empty))
        Controls.Add(New PageLoader.TextItem(ID_ChangePassword, "Acct_" + ID_ChangePassword, host.GetPageString(ID_ChangePassword), String.Empty))
        Controls.Add(New PageLoader.TextItem(ID_ChangeEMailAddress, "Acct_" + ID_ChangeEMailAddress, host.GetPageString(ID_ChangeEMailAddress), String.Empty))
        Controls.Add(New PageLoader.TextItem(ID_DeleteAccount, "Acct_" + ID_DeleteAccount, host.GetPageString(ID_DeleteAccount), String.Empty))
        If (IsAdmin(Context)) Then
            Controls.Add(New PageLoader.TextItem(ID_ClearCache, "Acct_" + ID_ClearCache, host.GetPageString(ID_ClearCache), String.Empty))
            Controls.Add(New PageLoader.TextItem(ID_ViewHeaders, "Acct_" + ID_ViewHeaders, host.GetPageString(ID_ViewHeaders), String.Empty))
            Controls.Add(New PageLoader.TextItem(ID_DotNetVersion, "Acct_" + ID_DotNetVersion, host.GetPageString(ID_DotNetVersion), String.Empty))
            Controls.Add(New PageLoader.TextItem(ID_ViewCache, "Acct_" + ID_ViewCache, host.GetPageString(ID_ViewCache), String.Empty))
            Controls.Add(New PageLoader.TextItem(ID_ClearDiskCache, "Acct_" + ID_ClearDiskCache, host.GetPageString(ID_ClearDiskCache), String.Empty))
            Controls.Add(New PageLoader.TextItem(ID_ViewDiskCache, "Acct_" + ID_ViewDiskCache, host.GetPageString(ID_ViewDiskCache), String.Empty))
            Controls.Add(New PageLoader.TextItem(ID_CreateDatabase, "Acct_" + ID_CreateDatabase, host.GetPageString(ID_CreateDatabase), String.Empty))
            Controls.Add(New PageLoader.TextItem(ID_RemoveDatabase, "Acct_" + ID_RemoveDatabase, host.GetPageString(ID_RemoveDatabase), String.Empty))
            Controls.Add(New PageLoader.TextItem(ID_CleanupState, "Acct_" + ID_CleanupState, host.GetPageString(ID_CleanupState), String.Empty))
            Controls.Add(New PageLoader.TextItem(ID_CertRequest, "Acct_" + ID_CertRequest, host.GetPageString(ID_CertRequest), String.Empty))
            Controls.Add(New PageLoader.TextItem(ID_CertInstall, "Acct_" + ID_CertInstall, host.GetPageString(ID_CertInstall), String.Empty))
            Controls.Add(New PageLoader.TextItem(ID_CertInstallIntermed, "Acct_" + ID_CertInstallIntermed, host.GetPageString(ID_CertInstallIntermed), String.Empty))
            Controls.Add(New PageLoader.TextItem(ID_DeleteCertRequest, "Acct_" + ID_DeleteCertRequest, host.GetPageString(ID_DeleteCertRequest), String.Empty))
            Controls.Add(New PageLoader.TextItem(ID_ViewCertRequests, "Acct_" + ID_ViewCertRequests, host.GetPageString(ID_ViewCertRequests), String.Empty))
            Controls.Add(New PageLoader.TextItem(ID_ViewSites, "Acct_" + ID_ViewSites, host.GetPageString(ID_ViewSites), String.Empty))
            Controls.Add(New PageLoader.TextItem(ID_EnableSSL, "Acct_" + ID_EnableSSL, host.GetPageString(ID_EnableSSL), String.Empty))
            Controls.Add(New PageLoader.TextItem(ID_CurrentUser, "Acct_" + ID_CurrentUser, host.GetPageString(ID_CurrentUser), String.Empty))
        End If
        Controls.Add(New PageLoader.TextItem(ID_Home, "Acct_" + ID_Home, host.RootPath))
        PageSet.Pages.Add(New PageLoader.PageItem(Controls, ID_ControlPanel, "Acct_" + ID_ControlPanel))
    End Sub
    Public Sub DeleteAccount(PageSet As PageLoader, Context As HttpContext, ByVal Certain As String, ByVal DeleteAccount As String)
        Dim Errors As New Generic.List(Of String)
        If Not IsLoggedIn(Context) Then Return
        If DeleteAccount = _PortableMethods.LoadResourceString("Acct_" + ID_DeleteAccount) Then
            If Certain <> "0" Then
                Errors.Add("Acct_" + ID_NotConfirmDelete)
            End If
            If Errors.Count = 0 Then
                SD.RemoveUser(GetUserID(Context))
                ClearLoginCookie(Context)
            End If
        End If
        Dim Controls As New Generic.List(Of Object)
        Dim Form As New Generic.List(Of Object)
        If DeleteAccount <> _PortableMethods.LoadResourceString("Acct_" + ID_DeleteAccount) Or Errors.Count <> 0 Then
            If Errors.Count <> 0 Then
                Controls.AddRange(Linq.Enumerable.Select(Errors.ToArray(), Function(ErrorID As String) CType(New PageLoader.TextItem(ID_Error, ErrorID), Object)))
            End If
            Form.Add(New PageLoader.RadioItem(ID_Certain, "Acct_" + ID_Certain, "1", New PageLoader.OptionItem() {New PageLoader.OptionItem("Acct_" + ID_Yes), New PageLoader.OptionItem("Acct_" + ID_No)}))
            Form.Add(New PageLoader.ButtonItem(ID_DeleteAccount, "Acct_" + ID_DeleteAccount))
            Controls.Add(New PageLoader.ListItem("Acct_" + ID_DeleteAccount, ID_DeleteAccount, Form, False, True, host.GetPageString(ID_DeleteAccount)))
        Else
            Controls.Add(New PageLoader.TextItem(String.Empty, "Acct_" + ID_AccountDeleteSuccess))
            Controls.Add(New PageLoader.TextItem(ID_Home, "Acct_" + ID_Home, host.RootPath))
        End If
        PageSet.Pages.Add(New PageLoader.PageItem(Controls, ID_DeleteAccount, "Acct_" + ID_DeleteAccount))
    End Sub
    Public Sub ChangeEMailAddress(PageSet As PageLoader, Context As HttpContext, ByVal EMail As String, ByVal EMailConfirm As String, ByVal ChangeEMailAddress As String)
        Dim Errors As New Generic.List(Of String)
        Dim UserID As Integer
        If Not IsLoggedIn(Context) Then Return
        If ChangeEMailAddress = _PortableMethods.LoadResourceString("Acct_" + ID_ChangeEMailAddress) Then
            Dim ErrorString As String = ValidateEMail(EMail)
            If ErrorString <> String.Empty Then Errors.Add(ErrorString)
            If EMail <> EMailConfirm Then
                Errors.Add("Acct_" + ID_EmailsNotMatch)
            End If
            If Errors.Count = 0 Then
                UserID = GetUserID(Context)
                SD.ChangeUserEMail(UserID, EMail)
                MD.SendActivationEMail(SD.GetUserName(UserID), EMail, UserID, SD.GetUserActivated(UserID), Context)
                ClearLoginCookie(Context)
            End If
        End If
        Dim Controls As New Generic.List(Of Object)
        Dim Form As New Generic.List(Of Object)
        If ChangeEMailAddress <> _PortableMethods.LoadResourceString("Acct_" + ID_ChangeEMailAddress) Or Errors.Count <> 0 Then
            If Errors.Count <> 0 Then
                Controls.AddRange(Linq.Enumerable.Select(Errors.ToArray(), Function(ErrorID As String) CType(New PageLoader.TextItem(ID_Error, ErrorID), Object)))
            End If
            Form.Add(New PageLoader.TextItem(String.Empty, "Acct_" + ID_EmailAddress))
            Form.Add(New PageLoader.EditItem(ID_EmailAddress, String.Empty, 1))
            Form.Add(New PageLoader.TextItem(String.Empty, "Acct_" + ID_ConfirmEmailAddress))
            Form.Add(New PageLoader.EditItem(ID_ConfirmEmailAddress, String.Empty, 1))
            Form.Add(New PageLoader.ButtonItem(ID_ChangeEMailAddress, "Acct_" + ID_ChangeEMailAddress))
            Controls.Add(New PageLoader.ListItem("Acct_" + ID_ChangeEMailAddress, ID_ChangeEMailAddress, Form, False, True, host.GetPageString(ID_ChangeEMailAddress)))
        Else
            Controls.Add(New PageLoader.TextItem(String.Empty, "Acct_" + ID_EmailChangeReactivate))
            Controls.Add(New PageLoader.TextItem(ID_Home, "Acct_" + ID_Home, host.RootPath))
        End If
        PageSet.Pages.Add(New PageLoader.PageItem(Controls, ID_ChangeEMailAddress, "Acct_" + ID_ChangeEMailAddress))
    End Sub
    Public Sub ChangePassword(PageSet As PageLoader, Context As HttpContext, ByVal Password As String, ByVal PasswordConfirm As String, ByVal ChangePassword As String)
        Dim Errors As New Generic.List(Of String)
        Dim UserID As Integer
        If Not IsLoggedIn(Context) Then Return
        If ChangePassword = _PortableMethods.LoadResourceString("Acct_" + ID_ChangePassword) Then
            Dim ErrorString As String = ValidatePassword(Password)
            If ErrorString <> String.Empty Then Errors.Add(ErrorString)
            If Password <> PasswordConfirm Then
                Errors.Add("Acct_" + ID_PasswordsNotMatch)
            End If
            If Errors.Count = 0 Then
                UserID = GetUserID(Context)
                SD.ChangeUserPassword(UserID, Password)
                MD.SendPasswordChangedEMail(SD.GetUserName(UserID), SD.GetUserEMail(UserID), Context)
            End If
        End If
        Dim Controls As New Generic.List(Of Object)
        Dim Form As New Generic.List(Of Object)
        If ChangePassword <> _PortableMethods.LoadResourceString("Acct_" + ID_ChangePassword) Or Errors.Count <> 0 Then
            If Errors.Count <> 0 Then
                Controls.AddRange(Linq.Enumerable.Select(Errors.ToArray(), Function(ErrorID As String) CType(New PageLoader.TextItem(ID_Error, ErrorID), Object)))
            End If
            Form.Add(New PageLoader.TextItem(String.Empty, "Acct_" + ID_Password))
            Form.Add(New PageLoader.EditItem(ID_Password, String.Empty, 1, True))
            Form.Add(New PageLoader.TextItem(String.Empty, "Acct_" + ID_ConfirmPassword))
            Form.Add(New PageLoader.EditItem(ID_ConfirmPassword, String.Empty, 1, True))
            Form.Add(New PageLoader.ButtonItem(ID_ChangePassword, "Acct_" + ID_ChangePassword))
            Controls.Add(New PageLoader.ListItem("Acct_" + ID_ChangePassword, ID_ChangePassword, Form, False, True, host.GetPageString(ID_ChangePassword)))
        Else
            Controls.Add(New PageLoader.TextItem(String.Empty, "Acct_" + ID_PasswordChangeSuccess))
            Controls.Add(New PageLoader.TextItem(ID_Home, "Acct_" + ID_Home, host.RootPath))
        End If
        PageSet.Pages.Add(New PageLoader.PageItem(Controls, ID_ChangePassword, "Acct_" + ID_ChangePassword))
    End Sub
    Public Sub ChangeUserName(PageSet As PageLoader, Context As HttpContext, ByVal UserName As String, ByVal ChangeUserName As String)
        Dim Errors As New Generic.List(Of String)
        Dim UserID As Integer
        If Not IsLoggedIn(Context) Then Return
        If ChangeUserName = _PortableMethods.LoadResourceString("Acct_" + ID_ChangeUsername) Then
            Dim ErrorString As String = ValidateUserName(UserName)
            If ErrorString <> String.Empty Then Errors.Add(ErrorString)
            If Errors.Count = 0 Then
                UserID = GetUserID(Context)
                SD.ChangeUserName(UserID, UserName)
                MD.SendUserNameChangedEMail(SD.GetUserName(UserID), SD.GetUserEMail(UserID), Context)
            End If
        End If
        Dim Controls As New Generic.List(Of Object)
        Dim Form As New Generic.List(Of Object)
        If ChangeUserName <> _PortableMethods.LoadResourceString("Acct_" + ID_ChangeUsername) Or Errors.Count <> 0 Then
            If Errors.Count <> 0 Then
                Controls.AddRange(Linq.Enumerable.Select(Errors.ToArray(), Function(ErrorID As String) CType(New PageLoader.TextItem(ID_Error, ErrorID), Object)))
            End If
            Form.Add(New PageLoader.TextItem(String.Empty, "Acct_" + ID_Username))
            Form.Add(New PageLoader.EditItem(ID_Username, String.Empty, 1))
            Form.Add(New PageLoader.ButtonItem(ID_ChangeUsername, "Acct_" + ID_ChangeUsername))
            Controls.Add(New PageLoader.ListItem("Acct_" + ID_ChangeUsername, ID_ChangeUsername, Form, False, True, host.GetPageString(ID_ChangeUsername)))
        Else
            Controls.Add(New PageLoader.TextItem(String.Empty, "Acct_" + ID_UsernameChangeSuccess))
            Controls.Add(New PageLoader.TextItem(ID_Home, "Acct_" + ID_Home, host.RootPath))
        End If
        PageSet.Pages.Add(New PageLoader.PageItem(Controls, ID_ChangeUsername, "Acct_" + ID_ChangeUsername))
    End Sub
    'requires applicationHost.config configuration->system.applicationHost->sites->site->bindings->binding for https on *:443:, also command netsh http add sslcert hostnameport=localhost:<por> certhash=<certhash> certstorename=MY appid=<appid>
    'add 127.0.0.1  <domainname> to %windir%\system32\drivers\etc\hosts for testing
    'netsh http add sslcert hostnameport=<domainname>:<port> certhash=<certhash> certstorename=MY appid=<appid>
    'appid=Assembly.GetExecutingAssembly().GetType().GUID.ToString()
    Public Sub UploadCertificate(PageSet As PageLoader, Context As HttpContext, CertInstall As String, Certificate As String, CertRequest As String)
        If Not IsLoggedIn(Context) And IsAdmin(Context) Then Return
        Dim Output As String = String.Empty
        If CertInstall = _PortableMethods.LoadResourceString("Acct_" + ID_CertInstall) Then
            Dim objEnroll As New CERTENROLLLib.CX509Enrollment()
            Try
                'Install the certificate
                If CertRequest.Length <> 0 Then
                    Dim objCertReq As New CERTENROLLLib.CX509CertificateRequestPkcs10
                    objCertReq.InitializeFromCertificate(CERTENROLLLib.X509CertificateEnrollmentContext.ContextUser, CertRequest)
                    'objCertReq.InitializeDecode(CertRequest)
                    objEnroll.InitializeFromRequest(CType(objCertReq, CERTENROLLLib.IX509CertificateRequest))
                Else
                    objEnroll.Initialize(CERTENROLLLib.X509CertificateEnrollmentContext.ContextUser)
                End If
                objEnroll.InstallResponse(
                    CERTENROLLLib.InstallResponseRestrictionFlags.AllowUntrustedRoot,
                    Certificate, CERTENROLLLib.EncodingType.XCN_CRYPT_STRING_BASE64, Nothing)
                Output = "Certificate Installed!"
            Catch ex As Exception
                Output = "Error: " + ex.Message
            End Try
        End If
        Dim Controls As New Generic.List(Of Object)
        Dim Form As New Generic.List(Of Object)
        Controls.Add(New PageLoader.ListItem("Acct_" + ID_CertInstall, ID_CertInstall, Form, False, True, host.GetPageString(ID_CertInstall)))
        Form.Add(New PageLoader.EditItem(ID_Certificate, String.Empty, 10))
        Form.Add(New PageLoader.EditItem(ID_CertRequest, String.Empty, 10))
        Form.Add(New PageLoader.TextItem("", Output, , , PL.DoLookup("Utility::TextRender")))
        Form.Add(New PageLoader.ButtonItem(ID_CertInstall, "Acct_" + ID_CertInstall))
        PageSet.Pages.Add(New PageLoader.PageItem(Controls, ID_CertInstall, "Acct_" + ID_CertInstall))
    End Sub
    Public Sub DeleteCertificateRequest(PageSet As PageLoader, Context As HttpContext, DeleteCertRequest As String, Certificate As String)
        If Not IsLoggedIn(Context) And IsAdmin(Context) Then Return
        Dim Output As String = String.Empty
        If DeleteCertRequest = _PortableMethods.LoadResourceString("Acct_" + ID_DeleteCertRequest) Then
            Dim Store As New System.Security.Cryptography.X509Certificates.X509Store("REQUEST", System.Security.Cryptography.X509Certificates.StoreLocation.LocalMachine)
            Store.Open(System.Security.Cryptography.X509Certificates.OpenFlags.MaxAllowed)
            For Each Cert As System.Security.Cryptography.X509Certificates.X509Certificate2 In Store.Certificates
                If Cert.Subject = Certificate Then
                    Store.Certificates.Remove(Cert)
                    Output = "Certificate Deleted!"
                End If
            Next
            Store.Close()
        End If
        Dim Controls As New Generic.List(Of Object)
        Dim Form As New Generic.List(Of Object)
        Controls.Add(New PageLoader.ListItem("Acct_" + ID_DeleteCertRequest, ID_DeleteCertRequest, Form, False, True, host.GetPageString(ID_DeleteCertRequest)))
        Form.Add(New PageLoader.EditItem(ID_Certificate, String.Empty, 10))
        Form.Add(New PageLoader.TextItem("", Output, , , PL.DoLookup("Utility::TextRender")))
        Form.Add(New PageLoader.ButtonItem(ID_DeleteCertRequest, "Acct_" + ID_DeleteCertRequest))
        PageSet.Pages.Add(New PageLoader.PageItem(Controls, ID_DeleteCertRequest, "Acct_" + ID_DeleteCertRequest))
    End Sub
    Public Sub InstallIntermediateCert(PageSet As PageLoader, Context As HttpContext, CertInstallIntermed As String, Certificate As String)
        If Not IsLoggedIn(Context) And IsAdmin(Context) Then Return
        Dim Output As String = String.Empty
        If CertInstallIntermed = _PortableMethods.LoadResourceString("Acct_" + ID_CertInstallIntermed) Then
            Dim Store As New System.Security.Cryptography.X509Certificates.X509Store(System.Security.Cryptography.X509Certificates.StoreName.CertificateAuthority, System.Security.Cryptography.X509Certificates.StoreLocation.CurrentUser)
            Store.Open(System.Security.Cryptography.X509Certificates.OpenFlags.MaxAllowed)
            Dim Cms As New System.Security.Cryptography.Pkcs.SignedCms
            Cms.Decode(Convert.FromBase64String(Certificate.Replace(vbLf, String.Empty).Replace(vbCr, String.Empty)))
            Store.AddRange(Cms.Certificates)
            Store.Close()
            Output = "Installed Certificate!"
        End If
        Dim Controls As New Generic.List(Of Object)
        Dim Form As New Generic.List(Of Object)
        Controls.Add(New PageLoader.ListItem("Acct_" + ID_CertInstallIntermed, ID_CertInstallIntermed, Form, False, True, host.GetPageString(ID_CertInstallIntermed)))
        Form.Add(New PageLoader.EditItem(ID_Certificate, String.Empty, 10))
        Form.Add(New PageLoader.TextItem("", Output, , , PL.DoLookup("Utility::TextRender")))
        Form.Add(New PageLoader.ButtonItem(ID_CertInstallIntermed, "Acct_" + ID_CertInstallIntermed))
        PageSet.Pages.Add(New PageLoader.PageItem(Controls, ID_CertInstallIntermed, "Acct_" + ID_CertInstallIntermed))
    End Sub
    Public Sub CreateCertificateRequest(PageSet As PageLoader, Context As HttpContext, CertRequest As String, PrivateKey As String)
        If Not IsLoggedIn(Context) And IsAdmin(Context) Then Return
        Dim Output As String = String.Empty
        If CertRequest = _PortableMethods.LoadResourceString("Acct_" + ID_CertRequest) Then
            'Create all the objects that will be required
            Dim objPkcs10 As New CERTENROLLLib.CX509CertificateRequestPkcs10()
            Dim objPrivateKey As New CERTENROLLLib.CX509PrivateKey()
            Dim objCSP As New CERTENROLLLib.CCspInformation()
            Dim objCSPs As New CERTENROLLLib.CCspInformations()
            Dim objDN As New CERTENROLLLib.CX500DistinguishedName()
            Dim objEnroll As New CERTENROLLLib.CX509Enrollment()
            Dim objObjectIds As New CERTENROLLLib.CObjectIds()
            Dim objObjectId As New CERTENROLLLib.CObjectId()
            Dim objExtensionKeyUsage As New CERTENROLLLib.CX509ExtensionKeyUsage()
            Dim objX509ExtensionEnhancedKeyUsage As New CERTENROLLLib.CX509ExtensionEnhancedKeyUsage()
            Dim objExtension As New CERTENROLLLib.CX509ExtensionAlternativeNames
            Dim objAlternativeNames As New CERTENROLLLib.CAlternativeNames
            Dim objAlternativeName As CERTENROLLLib.CAlternativeName
            Dim strRequest As String

            Try
                'Initialize the csp object using the desired Cryptograhic Service Provider (CSP)
                objCSP.InitializeFromName("Microsoft RSA Schannel Cryptographic Provider")
                'Add this CSP object to the CSP collection object
                objCSPs.Add(objCSP)

                'Provide key container name, key length and key spec to the private key object
                'objPrivateKey.ContainerName = Utility.ConnectionData.KeyContainerName
                objPrivateKey.Length = 2048
                objPrivateKey.KeySpec = CERTENROLLLib.X509KeySpec.XCN_AT_KEYEXCHANGE
                objPrivateKey.KeyUsage = CERTENROLLLib.X509PrivateKeyUsageFlags.XCN_NCRYPT_ALLOW_ALL_USAGES
                objPrivateKey.MachineContext = True
                objPrivateKey.ExportPolicy = CERTENROLLLib.X509PrivateKeyExportFlags.XCN_NCRYPT_ALLOW_EXPORT_FLAG

                'Provide the CSP collection object (in this case containing only 1 CSP object)
                'to the private key object
                objPrivateKey.CspInformations = objCSPs

                Dim strCert As String = PrivateKey
                If strCert.Length <> 0 Then
                    objPrivateKey.Import("PRIVATEBLOB", strCert)
                Else
                    'Create the actual key pair
                    objPrivateKey.Create()
                End If

                'Initialize the PKCS#10 certificate request object based on the private key.
                'Using the context, indicate that this is a user certificate request and don't
                'provide a template name
                objPkcs10.InitializeFromPrivateKey(
                    CERTENROLLLib.X509CertificateEnrollmentContext.ContextMachine,
                    objPrivateKey, "")

                'Key Usage Extension 
                objExtensionKeyUsage.InitializeEncode(
                    CERTENROLLLib.X509KeyUsageFlags.XCN_CERT_DIGITAL_SIGNATURE_KEY_USAGE Or
                    CERTENROLLLib.X509KeyUsageFlags.XCN_CERT_KEY_ENCIPHERMENT_KEY_USAGE)
                objPkcs10.X509Extensions.Add(CType(objExtensionKeyUsage, CERTENROLLLib.CX509Extension))

                'Enhanced Key Usage Extension
                objObjectId.InitializeFromValue("1.3.6.1.5.5.7.3.1") ' OID for Server Authentication usage
                objObjectIds.Add(objObjectId)
                objObjectId = New CERTENROLLLib.CObjectId
                objObjectId.InitializeFromValue("1.3.6.1.5.5.7.3.2") ' OID for Client Authentication usage
                objObjectIds.Add(objObjectId)
                For Each Domain As String In UtilityWeb.ConnectionData.CertExtraDomains
                    objAlternativeName = New CERTENROLLLib.CAlternativeName
                    objAlternativeName.InitializeFromString(CERTENROLLLib.AlternativeNameType.XCN_CERT_ALT_NAME_DNS_NAME, Domain)
                    objAlternativeNames.Add(objAlternativeName)
                Next
                objExtension.InitializeEncode(objAlternativeNames)
                objPkcs10.X509Extensions.Add(CType(objExtension, CERTENROLLLib.CX509Extension))
                objX509ExtensionEnhancedKeyUsage.InitializeEncode(objObjectIds)
                objPkcs10.X509Extensions.Add(CType(objX509ExtensionEnhancedKeyUsage, CERTENROLLLib.CX509Extension))

                'Encode the name in using the Distinguished Name object
                objDN.Encode(UtilityWeb.ConnectionData.DistinguishedName,
                    CERTENROLLLib.X500NameFlags.XCN_CERT_X500_NAME_STR)

                'Assing the subject name by using the Distinguished Name object initialized above
                objPkcs10.Subject = objDN
                'Create enrollment request
                objEnroll.InitializeFromRequest(objPkcs10)

                strRequest = objEnroll.CreateRequest(
                    CERTENROLLLib.EncodingType.XCN_CRYPT_STRING_BASE64REQUESTHEADER)
                Output = "Private key: " + vbCrLf + objPrivateKey.Export("PRIVATEBLOB", CERTENROLLLib.EncodingType.XCN_CRYPT_STRING_BASE64 Or CERTENROLLLib.EncodingType.XCN_CRYPT_STRING_NOCR Or CERTENROLLLib.EncodingType.XCN_CRYPT_STRING_NOCRLF) + vbCrLf + strRequest
                Dim Store As New System.Security.Cryptography.X509Certificates.X509Store("REQUEST", System.Security.Cryptography.X509Certificates.StoreLocation.LocalMachine)
                Store.Open(System.Security.Cryptography.X509Certificates.OpenFlags.MaxAllowed)
                For Each Cert As System.Security.Cryptography.X509Certificates.X509Certificate2 In Store.Certificates
                    If Cert.Subject = objDN.Name Then
                        'Cert.Export(System.Security.Cryptography.X509Certificates.X509ContentType.Pfx, String.Empty)
                        Output += "-----BEGIN CERTIFICATE-----" + vbCrLf + Convert.ToBase64String(Cert.Export(System.Security.Cryptography.X509Certificates.X509ContentType.Cert), Base64FormattingOptions.InsertLineBreaks) + vbCrLf + "-----END CERTIFICATE-----" + vbCrLf
                    End If
                Next
            Catch ex As Exception
                Output = "Error: "
            End Try
        End If
        Dim Controls As New Generic.List(Of Object)
        Dim Form As New Generic.List(Of Object)
        Controls.Add(New PageLoader.ListItem("Acct_" + ID_CertRequest, ID_CertRequest, Form, False, True, host.GetPageString(ID_CertRequest)))
        Form.Add(New PageLoader.EditItem(ID_PrivateKey, String.Empty, 10))
        Form.Add(New PageLoader.TextItem("", Output, , , PL.DoLookup("Utility::TextRender")))
        Form.Add(New PageLoader.ButtonItem(ID_CertRequest, "Acct_" + ID_CertRequest))
        PageSet.Pages.Add(New PageLoader.PageItem(Controls, ID_CertRequest, "Acct_" + ID_CertRequest))
    End Sub
    Public Async Function AccountPanel(ByVal writer As HtmlTextWriter, Context As HttpContext) As Threading.Tasks.Task
        If IsLoggedIn(Context) Then
            writer.Write(vbCrLf + vbTab + vbTab + vbTab + vbTab)
            writer.WriteFullBeginTag("div")
            writer.Write(vbCrLf + vbTab + vbTab + vbTab + vbTab + vbTab)
            writer.Write(UtilityWeb.HtmlTextEncode(GetUserName(Context)))
            writer.Write(vbCrLf + vbTab + vbTab + vbTab + vbTab + vbTab)
            writer.WriteBeginTag("a")
            writer.WriteAttribute("href", UtilityWeb.HtmlTextEncode(host.GetPageString(ID_Logoff)))
            writer.Write(HtmlTextWriter.TagRightChar)
            writer.Write(vbCrLf + vbTab + vbTab + vbTab + vbTab + vbTab + vbTab)
            writer.Write(UtilityWeb.HtmlTextEncode(ID_Logoff))
            writer.Write(vbCrLf + vbTab + vbTab + vbTab + vbTab + vbTab)
            writer.WriteEndTag("a")
            writer.Write(vbCrLf + vbTab + vbTab + vbTab + vbTab + vbTab)
            writer.WriteBeginTag("a")
            writer.WriteAttribute("href", UtilityWeb.HtmlTextEncode(host.GetPageString(ID_ControlPanel)))
            writer.Write(HtmlTextWriter.TagRightChar)
            writer.Write(vbCrLf + vbTab + vbTab + vbTab + vbTab + vbTab + vbTab)
            writer.Write(UtilityWeb.HtmlTextEncode(_PortableMethods.LoadResourceString("Acct_" + ID_ControlPanel)))
            writer.Write(vbCrLf + vbTab + vbTab + vbTab + vbTab + vbTab)
            writer.WriteEndTag("a")
            writer.Write(vbCrLf + vbTab + vbTab + vbTab + vbTab)
            writer.WriteEndTag("div")
        Else
            Dim Form As New Generic.List(Of Object)
            Dim Controls As New Generic.List(Of Object)
            Form.Add(New PageLoader.TextItem(String.Empty, "Acct_" + ID_Username))
            Form.Add(New PageLoader.EditItem(ID_Username, String.Empty, 1))
            Form.Add(New PageLoader.TextItem(String.Empty, "Acct_" + ID_Password))
            Form.Add(New PageLoader.EditItem(ID_Password, String.Empty, 1, True))
            Form.Add(New PageLoader.RadioItem(ID_Remember, "Acct_" + ID_Remember, "1", New PageLoader.OptionItem() {New PageLoader.OptionItem("Acct_" + ID_Yes), New PageLoader.OptionItem("Acct_" + ID_No)}))
            Form.Add(New PageLoader.ButtonItem(ID_Login, "Acct_" + ID_Login))
            Form.Add(New PageLoader.TextItem(ID_Register, "Acct_" + ID_Register, host.GetPageString(ID_Register)))
            Controls.Add(New PageLoader.ListItem("Acct_" + ID_Login, ID_Login, Form, False, True, host.GetPageString(ID_Login)))
            Dim RenderPage As New Page(_PortableMethods, UWeb, ArbData, New PageLoader.PageItem(Controls, ID_Login, "Acct_" + ID_Login), False)
            Await RenderPage.DoGetAsyncData()
            RenderPage.RenderControl(writer)
        End If
    End Function
End Class