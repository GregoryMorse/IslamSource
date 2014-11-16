Option Explicit On
Option Strict On
Imports System.Drawing
Imports System.Web
Imports System.Web.UI
Public Class Utility
    Delegate Function _GetUserID() As Integer
    Public Shared GetUserID As _GetUserID
    Delegate Function _IsLoggedIn() As Boolean
    Public Shared IsLoggedIn As _IsLoggedIn
    Delegate Function _GetPageString(Page As String) As String
    Public Shared GetPageString As _GetPageString
    'HttpContext.Current.Trace.Write(Text)
    Public Shared Sub Initialize(NewGetPageString As _GetPageString, NewGetUserID As _GetUserID, NewIsLoggedIn As _IsLoggedIn)
        GetPageString = NewGetPageString
        GetUserID = NewGetUserID
        IsLoggedIn = NewIsLoggedIn
    End Sub
    Public Const LocalConfig As String = "~/web.config"
    Public Class ConnectionData
        Public Shared ReadOnly Property IslamSourceAdminEMail As String
            Get
                Return GetConfigSetting("islamsourceadminemail")
            End Get
        End Property
        Public Shared ReadOnly Property IslamSourceAdminEMailPass As String
            Get
                Return DoDecrypt(GetConfigSetting("islamsourceadminemailpass"))
            End Get
        End Property
        Public Shared ReadOnly Property IslamSourceAdminName As String
            Get
                Return GetConfigSetting("islamsourceadminname")
            End Get
        End Property
        Public Shared ReadOnly Property IslamSourceMailServer As String
            Get
                Return GetConfigSetting("islamsourcemailserver")
            End Get
        End Property
        Public Shared ReadOnly Property EMailAddress As String
            Get
                Return GetConfigSetting("emailaddress")
            End Get
        End Property
        Public Shared ReadOnly Property AuthorName As String
            Get
                Return GetConfigSetting("authorname")
            End Get
        End Property
        Public Shared ReadOnly Property FBAppID As String
            Get
                Return GetConfigSetting("fbappid")
            End Get
        End Property
        Public Shared ReadOnly Property SiteDomains As String()
            Get
                Return GetConfigSetting("sitedomains").Split(";"c)
            End Get
        End Property
        Public Shared ReadOnly Property SiteXMLs As String()
            Get
                Return GetConfigSetting("sitexmls").Split(";"c)
            End Get
        End Property
        Public Shared ReadOnly Property DocXML As String
            Get
                Return GetConfigSetting("docxml")
            End Get
        End Property
        Public Shared ReadOnly Property DefaultXML As String
            Get
                Return SiteXMLs(0)
            End Get
        End Property
        Public Shared ReadOnly Property AlternatePath As String
            Get
                Return GetConfigSetting("alternatepath")
            End Get
        End Property
        Public Shared ReadOnly Property CertExtraDomains As String()
            Get
                Return GetConfigSetting("certextradomains").Split(";"c)
            End Get
        End Property
        Public Shared ReadOnly Property DistinguishedName As String
            Get
                Return GetConfigSetting("distinguishedname")
            End Get
        End Property
        Public Shared ReadOnly Property IPInfoDBAPIKey As String
            Get
                Return GetConfigSetting("ipinfodbapikey")
            End Get
        End Property
        Public Shared ReadOnly Property Resources As KeyValuePair(Of String, String())()
            Get
                Return Array.ConvertAll(GetConfigSetting("resources").Split(";"c), Function(Str As String) New KeyValuePair(Of String, String())(Str.Split("="c)(0), Str.Split("="c)(1).Split(","c)))
            End Get
        End Property
        Public Shared ReadOnly Property FuncLibs As String()
            Get
                Return GetConfigSetting("funclibs").Split(","c)
            End Get
        End Property
        Public Const KeyFileName As String = "prv.key"
        Public Const KeyContainerName As String = "HOSTPAGE_CRYPT"

        Public Shared ReadOnly Property DbConnServer As String
            Get
                Return GetConfigSetting("mysqldbserver", "localhost")
            End Get
        End Property
        Public Shared ReadOnly Property DbConnUid As String
            Get
                Return GetConfigSetting("mysqldbuid")
            End Get
        End Property
        Public Shared ReadOnly Property DbConnPwd As String
            Get
                Return DoDecrypt(GetConfigSetting("mysqldbpwd"))
            End Get
        End Property
        Public Shared ReadOnly Property DbConnDatabase As String
            Get
                Return GetConfigSetting("mysqldbname")
            End Get
        End Property
    End Class
    Public Shared Function IsDesktopApp() As Boolean
        Return Not Reflection.Assembly.GetEntryAssembly() Is Nothing AndAlso New Reflection.AssemblyName(Reflection.Assembly.GetEntryAssembly().FullName).Name = "IslamSource"
    End Function
    Public Shared Function GetTemplatePath() As String
        If IsDesktopApp() Then
            Return GetFilePath("metadata\IslamSource.xml")
        Else
            Dim Index As Integer = Array.FindIndex(ConnectionData.SiteDomains(), Function(Domain As String) HttpContext.Current.Request.Url.Host.EndsWith(Domain))
            If Index = -1 Then
                Return GetFilePath("metadata\" + ConnectionData.DefaultXML + ".xml")
            Else
                Return GetFilePath("metadata\" + ConnectionData.SiteXMLs()(Index) + ".xml")
            End If
        End If
    End Function
    Public Shared Function GetFilePath(ByVal Path As String) As String
        If IsDesktopApp() Then
            Return "..\..\..\" + Path
        Else
            Return CStr(IIf(IO.File.Exists(HttpContext.Current.Request.PhysicalApplicationPath + Path), HttpContext.Current.Request.PhysicalApplicationPath + Path, HttpContext.Current.Request.PhysicalApplicationPath + ConnectionData.AlternatePath + Path))
        End If
    End Function
    Friend Shared Function GetStringHashCode(ByVal s As String) As Integer
        Dim spin As System.Runtime.InteropServices.GCHandle = System.Runtime.InteropServices.GCHandle.Alloc(s, Runtime.InteropServices.GCHandleType.Pinned)
        Dim str As IntPtr = spin.AddrOfPinnedObject()
        Dim chPtr As IntPtr = str
        Dim num As Long = &H15051505
        Dim num2 As Long = num
        Dim numPtr As IntPtr = chPtr
        Dim i As Integer = s.Length
        Do While (i > 0)
            num = ((((num << 5) + num) + (num >> &H1B)) Xor System.Runtime.InteropServices.Marshal.ReadInt32(numPtr))
            If (i <= 2) Then
                Exit Do
            End If
            num2 = ((((num2 << 5) + num2) + (num2 >> &H1B)) Xor System.Runtime.InteropServices.Marshal.ReadInt32(New IntPtr(numPtr.ToInt64() + 4)))
            numPtr = New IntPtr(numPtr.ToInt64() + 8)
            i = (i - 4)
        Loop
        spin.Free()
        Return CInt((num + (num2 * &H5D588B65)) And &H800000007FFFFFFFL)
    End Function
    Public Shared Function LoadResourceString(resourceKey As String) As String
        LoadResourceString = Nothing
        For Each Pair In ConnectionData.Resources
            If Array.FindIndex(Pair.Value, Function(Str As String) Str = resourceKey Or resourceKey.StartsWith(Str + "_")) <> -1 Then
                LoadResourceString = New System.Resources.ResourceManager(Pair.Key + ".Resources", Reflection.Assembly.Load(Pair.Key)).GetString(resourceKey, Threading.Thread.CurrentThread.CurrentUICulture)
            End If
        Next
        If LoadResourceString = Nothing Then
            LoadResourceString = String.Empty
            'System.Diagnostics.Debug.WriteLine("  <data name=""" + resourceKey + """ xml:space=""preserve"">" + vbCrLf + "    <value>" + System.Text.RegularExpressions.Regex.Replace(System.Text.RegularExpressions.Regex.Replace(resourceKey, ".*_", String.Empty), "(.+?)([A-Z])", "$1 $2") + "</value>" + vbCrLf + "  </data>")
        End If
    End Function
    Public Shared Function DefaultValue(Value As String, DefValue As String) As String
        If Value Is Nothing Then Return DefValue
        Return Value
    End Function
    Public Shared Function GetConfigSetting(Key As String, Optional DefaultValue As String = "") As String
        Dim rootWebConfig As System.Configuration.Configuration = Web.Configuration.WebConfigurationManager.OpenWebConfiguration(LocalConfig)
        If rootWebConfig.AppSettings.Settings.Count > 0 Then
            Dim customSetting As System.Configuration.KeyValueConfigurationElement
            customSetting = rootWebConfig.AppSettings.Settings(Key)
            If Not customSetting.Value = Nothing Then Return customSetting.Value
        End If
        Return DefaultValue
    End Function
    Public Shared Function DoEncrypt(EncodeStr As String) As String
        Dim cspParams As New System.Security.Cryptography.CspParameters(1, "Microsoft Base Cryptographic Provider v1.0", Utility.ConnectionData.KeyContainerName)
        cspParams.KeyNumber = System.Security.Cryptography.KeyNumber.Exchange
        cspParams.Flags = System.Security.Cryptography.CspProviderFlags.NoFlags
        Dim Transform As New System.Security.Cryptography.RSACryptoServiceProvider(512, cspParams)
        Dim EncodeBytes As Byte() = Transform.Encrypt(System.Text.Encoding.UTF8.GetBytes(EncodeStr), False)
        Transform.Clear()
        Array.Reverse(EncodeBytes) '.NET uses reverse from order of CryptEncrypt
        IO.File.WriteAllBytes(Utility.GetFilePath("bin\" + Utility.ConnectionData.KeyFileName), Transform.ExportCspBlob(True))
        Return String.Join(String.Empty, Array.ConvertAll(EncodeBytes, Function(Convert As Byte) Convert.ToString("X2")))
    End Function
    Public Shared Function DoDecrypt(DecryptStr As String) As String
        Dim cspParams As New System.Security.Cryptography.CspParameters(1, "Microsoft Base Cryptographic Provider v1.0", Utility.ConnectionData.KeyContainerName)
        cspParams.KeyNumber = System.Security.Cryptography.KeyNumber.Exchange
        cspParams.Flags = System.Security.Cryptography.CspProviderFlags.UseMachineKeyStore 'user may change to must use machine store
        Dim Transform As New System.Security.Cryptography.RSACryptoServiceProvider(512, cspParams)
        Dim CspBlob As Byte() = IO.File.ReadAllBytes(Utility.GetFilePath("bin\" + Utility.ConnectionData.KeyFileName))
        Transform.PersistKeyInCsp = False
        Transform.ImportCspBlob(CspBlob)
        Dim Bytes(DecryptStr.Length \ 2 - 1) As Byte '.NET uses reverse from order of CryptDecrypt
        For Count As Integer = 0 To DecryptStr.Length - 1 Step 2
            Bytes(DecryptStr.Length \ 2 - 1 - Count \ 2) = Byte.Parse(DecryptStr.Substring(Count, 2), Globalization.NumberStyles.HexNumber)
        Next
        Dim Str As String = System.Text.Encoding.UTF8.GetString(Transform.Decrypt(Bytes, False)).TrimEnd(Chr(0)) 'not using OAEP when calling CryptDe/Encrypt
        Transform.Clear()
        Return Str
    End Function
    Public Shared Function ConvertSpaces(ByVal Text As String) As String
        Dim Location As Integer = 0
        Do
            Location = Text.IndexOf("  ", Location)
            If Location = -1 Then
                Exit Do
            End If
            Text = Text.Remove(Location, 1).Insert(Location, "&nbsp;")
        Loop
        ConvertSpaces = Text
    End Function
    Public Shared Function GetDigitLength(ByVal Number As Integer) As Integer
        If Number = 0 Then Return 1
        Return CInt(Math.Floor(Math.Log10(Number))) + 1
    End Function
    Public Shared Function ZeroPad(ByVal PadString As String, ByVal ZeroCount As Integer) As String
        Dim RetString As String = StrDup(ZeroCount, "0") + PadString
        Return RetString.Substring(RetString.Length - ZeroCount)
    End Function
    Public Class EMailValidator
        Dim invalid As Boolean
        Public Function IsValidEMail(ByVal strIn As String) As Boolean
            invalid = False
            If String.IsNullOrEmpty(strIn) Then Return False
            'Use IdnMapping class to convert Unicode domain names.
            strIn = System.Text.RegularExpressions.Regex.Replace(strIn, "(@)(.+)$", New System.Text.RegularExpressions.MatchEvaluator(AddressOf DomainMapper))
            If invalid Then Return False
            'Return true if strIn is in valid e-mail format.
            'not javascript compatible due to lookbehind
            Return System.Text.RegularExpressions.Regex.IsMatch(strIn, _
              "^(?("")(""[^""]+?""@)|(([0-9a-z]((\.(?!\.))|[-!#\$%&'\*\+/=\?\^`\{\}\|~\w])*)(?<=[0-9a-z])@))" + _
              "(?(\[)(\[(\d{1,3}\.){3}\d{1,3}\])|(([0-9a-z][-\w]*[0-9a-z]*\.)+[a-z0-9]{2,17}))$", _
              System.Text.RegularExpressions.RegexOptions.IgnoreCase)
        End Function
        Private Function DomainMapper(ByVal match As System.Text.RegularExpressions.Match) As String
            'IdnMapping class with default property values.
            Dim idn As Globalization.IdnMapping = New Globalization.IdnMapping()
            Dim domainName As String = match.Groups(2).Value
            Try
                domainName = idn.GetAscii(domainName)
            Catch e As ArgumentException
                invalid = True
            End Try
            Return match.Groups(1).Value + domainName
        End Function
    End Class
    Class CompareNameValueArray
        Implements Collections.IComparer
        'Compares an array of structures with a String and Integer element
        Public Function Compare(ByVal x As Object, ByVal y As Object) As Integer _
            Implements Collections.IComparer.Compare
            If CInt(CType(x, Array).GetValue(1)) = CInt(CType(y, Array).GetValue(1)) Then
                Compare = 0
            Else
                Compare = CInt(IIf(CInt(CType(x, Array).GetValue(1)) > CInt(CType(y, Array).GetValue(1)), 1, -1))
            End If
        End Function
    End Class
    Public Shared Function EscapeJS(Str As String) As String
        Return Str.Replace("\", "\\")
    End Function
    Public Shared Function EncodeJS(Str As String) As String
        Return Str.Replace("'", "\'")
    End Function
    Public Shared Function MakeJSString(Str As String) As String
        Return "'" + EncodeJS(Str) + "'"
    End Function
    Public Shared Function MakeJSArray(ByVal StringArray As String(), Optional ByVal bObject As Boolean = False) As String
        Dim JSArray As String = "["
        Dim Count As Integer
        For Count = 0 To StringArray.Length() - 1
            If StringArray(Count) Is Nothing Then
                JSArray += "null"
            ElseIf bObject Then
                JSArray += StringArray(Count)
            Else
                JSArray += MakeJSString(StringArray(Count))
            End If
            If (Count <> StringArray.Length() - 1) Then JSArray += ", "
        Next
        JSArray += "]"
        Return JSArray
    End Function
    Public Shared Function MakeJSIndexedObject(ByVal IndexNamesArray As String(), ByVal StringsArray As Array(), ByVal bObject As Boolean) As String
        Dim JSArray As String = String.Empty
        Dim Count As Integer
        Dim SubCount As Integer
        For Count = 0 To StringsArray.Length - 1
            JSArray += "{"
            For SubCount = 0 To IndexNamesArray.Length - 1
                JSArray += "'" + EncodeJS(IndexNamesArray(SubCount)) + "':"
                If CType(StringsArray(Count), Object())(SubCount) Is Nothing Then
                    JSArray += "null"
                ElseIf bObject Then
                    JSArray += CStr(CType(StringsArray(Count), Object())(SubCount))
                Else
                    JSArray += MakeJSString(CStr(CType(StringsArray(Count), String())(SubCount)))
                End If
                If (SubCount <> IndexNamesArray.Length() - 1) Then JSArray += ", "
            Next
            JSArray += "}"
            If (Count <> StringsArray.Length - 1) Then JSArray += ", "
        Next
        Return JSArray
    End Function
    Public Shared Function MakeTabString(ByVal Index As Integer) As String
        MakeTabString = StrDup(Index, vbTab)
    End Function
    Public Shared Function GetTextExtent(ByVal Text As String, ByVal MeasureFont As Font) As SizeF
        Dim bmp As New Bitmap(1, 1)
        Dim g As Graphics = Graphics.FromImage(bmp)
        g.PageUnit = GraphicsUnit.Pixel
        g.TextRenderingHint = Drawing.Text.TextRenderingHint.SystemDefault
        GetTextExtent = g.MeasureString(Text, MeasureFont, New PointF(0, 0), Drawing.StringFormat.GenericTypographic)
        g.Dispose()
        bmp.Dispose()
    End Function
    Public Shared Function HtmlTextEncode(ByVal Text As String) As String
        HtmlTextEncode = ConvertSpaces(HttpUtility.HtmlEncode(New System.Text.UTF8Encoding().GetString(System.Text.Encoding.UTF8.GetBytes(Text))))
    End Function
    Public Shared Function SourceTextEncode(ByVal Text As String) As String
        SourceTextEncode = Text.Replace(vbTab, New String(" "c, 4))
    End Function
    Public Shared Function DetectEncoding(ByVal Bytes As Byte()) As System.Text.Encoding
        'must check longest encodings first
        Dim encodingInfo As System.Text.EncodingInfo
        Dim encoding As System.Text.Encoding
        Dim preamble As Byte()
        Dim Count As Integer
        DetectEncoding = Nothing
        Dim preambles As New Collections.Generic.SortedList(Of Integer, System.Text.Encoding)
        For Each encodingInfo In System.Text.Encoding.GetEncodings()
            preamble = encodingInfo.GetEncoding().GetPreamble()
            If (preamble.Length > 0) Then
                preambles.Add(-(preamble.Length * 65536 + encodingInfo.CodePage), encodingInfo.GetEncoding())
            End If
        Next
        For Each encoding In preambles.Values
            preamble = encoding.GetPreamble()
            If Bytes.Length >= preamble.Length Then
                For Count = 0 To preamble.Length - 1
                    If preamble(Count) <> Bytes(Count) Then Exit For
                Next
                If (Count = preamble.Length) Then
                    DetectEncoding = encoding
                    Exit For
                End If
            End If
        Next
    End Function
    Public Class PrefixComparer
        Implements Collections.IComparer
        Public Function Compare(ByVal x As Object, ByVal y As Object) As Integer Implements System.Collections.IComparer.Compare
            Dim StrLeft As String() = CStr(x).Substring(0, CInt(IIf(CStr(x).IndexOf(":") <> -1, CStr(x).IndexOf(":"), CStr(x).Length))).Split(New Char() {"."c})
            Dim StrRight As String() = CStr(y).Substring(0, CInt(IIf(CStr(y).IndexOf(":") <> -1, CStr(y).IndexOf(":"), CStr(y).Length))).Split(New Char() {"."c})
            If StrLeft.Length = 0 And StrRight.Length = 0 Then Return 0
            If StrLeft.Length = 0 Then Return -1
            If StrRight.Length = 0 Then Return 1
            Dim Check As Integer = String.Compare(StrLeft(0), StrRight(0))
            If Check <> 0 Then Return Check
            If StrLeft.Length = 1 And StrRight.Length = 1 Then Return 0
            If StrLeft.Length = 1 Then Return -1
            If StrRight.Length = 1 Then Return 1
            If StrLeft.Length = 2 And StrRight.Length = 2 Then Return String.Compare(StrLeft(1), StrRight(1))
            If StrLeft.Length = 2 And StrRight.Length = 3 Then Return String.Compare(StrLeft(1), StrRight(2))
            If StrLeft.Length = 3 And StrRight.Length = 2 Then Return String.Compare(StrLeft(2), StrRight(1))
            Check = String.Compare(StrLeft(1), StrRight(1))
            If Check <> 0 Then Return Check
            Return String.Compare(StrLeft(2), StrRight(2))
        End Function
    End Class
    Public Shared Function GetFileLinesByNumberPrefix(Strings() As String, ByVal Prefix As String) As String()
        Dim Index As Integer = Array.BinarySearch(Strings, Prefix, New PrefixComparer)
        If Index < 0 OrElse Index >= Strings.Length OrElse (New PrefixComparer).Compare(Prefix, Strings(Index)) <> 0 Then Return New String() {}
        Dim StartIndex As Integer = Index - 1
        While StartIndex >= 0 _
            AndAlso (New PrefixComparer).Compare(Prefix, Strings(StartIndex)) = 0
            StartIndex -= 1
        End While
        StartIndex += 1
        Index += 1
        While Index < Strings.Length AndAlso _
            (New PrefixComparer).Compare(Prefix, Strings(Index)) = 0
            Index += 1
        End While
        Index -= 1
        Dim ReturnStrings(Index - StartIndex) As String
        Array.ConstrainedCopy(Strings, StartIndex, ReturnStrings, 0, Index - StartIndex + 1)
        For Index = 0 To ReturnStrings.Length - 1
            ReturnStrings(Index) = ReturnStrings(Index).Substring(CInt(IIf(ReturnStrings(Index).IndexOf(":") <> -1, ReturnStrings(Index).IndexOf(":") + 2, 0)))
        Next
        Return ReturnStrings
    End Function
    Public Shared Function GetImageDimensions(ByVal Path As String) As Drawing.SizeF
        Dim bmp As New Bitmap(Path)
        GetImageDimensions = bmp.GetBounds(Drawing.GraphicsUnit.Pixel).Size
        bmp.Dispose()
    End Function
    Public Shared Function ComputeImageScale(ByVal Width As Double, ByVal Height As Double, ByVal MaxWidth As Double, ByVal MaxHeight As Double) As Double
        ComputeImageScale = 1
        If (MaxWidth <> 0 AndAlso Width > MaxWidth) Then
            ComputeImageScale = Math.Max(Width / MaxWidth, ComputeImageScale)
        End If
        If (MaxHeight <> 0 AndAlso Height > MaxHeight) Then
            ComputeImageScale = Math.Max(Height / MaxHeight, ComputeImageScale)
        End If
    End Function
    Public Shared Function MakeThumbnail(ByVal inputImage As Bitmap, ByVal width As Integer, ByVal height As Integer) As Bitmap
        Dim outputImage As New Bitmap(width, height, Imaging.PixelFormat.Format32bppArgb)
        Dim g As Graphics = Graphics.FromImage(outputImage)
        g.CompositingMode = Drawing2D.CompositingMode.SourceCopy
        g.InterpolationMode = Drawing2D.InterpolationMode.HighQualityBicubic
        Dim destRect As Rectangle = New Rectangle(0, 0, width, height)
        g.DrawImage(inputImage, destRect, 0, 0, inputImage.Width, inputImage.Height, GraphicsUnit.Pixel)
        g.Dispose()
        If Not Bitmap.IsAlphaPixelFormat(inputImage.PixelFormat) Then outputImage.MakeTransparent(Color.White)
        MakeThumbnail = outputImage
    End Function
    Public Shared Sub ApplyTransparencyFilter(ByRef inputImage As Bitmap, ByVal initialColor As Color, ByVal finalColor As Color)
        Dim xCount As Integer
        Dim yCount As Integer
        For xCount = 0 To CInt(inputImage.GetBounds(Drawing.GraphicsUnit.Pixel).Size.Width) - 1
            For yCount = 0 To CInt(inputImage.GetBounds(Drawing.GraphicsUnit.Pixel).Size.Height) - 1
                If (inputImage.GetPixel(xCount, yCount).R >= initialColor.R) And _
                   (inputImage.GetPixel(xCount, yCount).G >= initialColor.G) And _
                   (inputImage.GetPixel(xCount, yCount).B >= initialColor.B) And _
                   (inputImage.GetPixel(xCount, yCount).R <= finalColor.R) And _
                   (inputImage.GetPixel(xCount, yCount).G <= finalColor.G) And _
                   (inputImage.GetPixel(xCount, yCount).B <= finalColor.B) Then
                    inputImage.SetPixel(xCount, yCount, Color.Transparent)
                End If
            Next
        Next
    End Sub
    Public Shared Function GetURLLastModified(ByVal URL As String) As Date
        Dim MyWebRequest As Net.HttpWebRequest = DirectCast(Net.WebRequest.Create(URL), Net.HttpWebRequest)
        MyWebRequest.Method = "HEAD"
        MyWebRequest.Accept = "*/*"
        MyWebRequest.Referer = New Uri(URL).GetLeftPart(UriPartial.Authority)
        Try
            Dim Response As Net.HttpWebResponse = DirectCast(MyWebRequest.GetResponse(), Net.HttpWebResponse)
            Dim DataStream As IO.Stream = Response.GetResponseStream()
            GetURLLastModified = Response.LastModified.ToUniversalTime()
            Response.Close()
        Catch ' e As System.Net.WebException
            GetURLLastModified = Date.MinValue
        End Try
    End Function
    Public Shared Function MakeThumbFromURL(ByVal URL As String, ByVal MaxWidth As Double, Optional ByRef ModifiedDate As Date = Nothing) As Bitmap
        Dim MyWebRequest As Net.HttpWebRequest = DirectCast(Net.WebRequest.Create(URL), Net.HttpWebRequest)
        Dim bmp As Bitmap = Nothing
        MyWebRequest.Accept = "image/*"
        MyWebRequest.Referer = New Uri(URL).GetLeftPart(UriPartial.Authority)
        Try
            Dim Response As Net.HttpWebResponse = DirectCast(MyWebRequest.GetResponse(), Net.HttpWebResponse)
            Dim DataStream As IO.Stream = Response.GetResponseStream()
            Try
                Dim SizeF As Drawing.SizeF
                Dim Scale As Double
                Dim Bitmap As New Bitmap(DataStream)
                SizeF = Bitmap.GetBounds(Drawing.GraphicsUnit.Pixel).Size
                Scale = Utility.ComputeImageScale(SizeF.Width, SizeF.Height, MaxWidth, MaxWidth * SizeF.Height / SizeF.Width)
                bmp = Utility.MakeThumbnail(Bitmap, Convert.ToInt32(SizeF.Width / Scale), Convert.ToInt32(SizeF.Height / Scale))
                ModifiedDate = Response.LastModified.ToUniversalTime()
            Catch
            End Try
            Response.Close()
        Catch
        End Try
        Return bmp
    End Function
    Public Shared Function GetThumbSizeFromURL(ByVal URL As String, ByVal CacheURL As String, ByVal MaxWidth As Double) As SizeF
        Dim ResultBmp As Bitmap = Nothing
        Dim Bytes() As Byte
        Dim DateModified As Date
        If CInt(DiskCache.GetCacheItems().Length * New Random().NextDouble()) = 0 Then
            DateModified = GetURLLastModified(URL)
        Else
            DateModified = Date.MinValue
        End If
        Bytes = DiskCache.GetCacheItem(CacheURL, DateModified)
        If Not Bytes Is Nothing Then
            ResultBmp = DirectCast(Bitmap.FromStream(New IO.MemoryStream(Bytes)), Bitmap)
        End If
        If ResultBmp Is Nothing Then
            'Must have a way to initialize this potentially long operation
            Dim bmp As Bitmap = Utility.MakeThumbFromURL(URL, MaxWidth, DateModified)
            If Not bmp Is Nothing Then
                Dim quantizer As ImageQuantization.OctreeQuantizer = New ImageQuantization.OctreeQuantizer(255, 8, Not Bitmap.IsAlphaPixelFormat(bmp.PixelFormat), Color.White)
                ResultBmp = quantizer.QuantizeBitmap(bmp)
                bmp.Dispose()
                Dim MemStream As New IO.MemoryStream()
                ResultBmp.Save(MemStream, DirectCast(IIf(Object.Equals(ResultBmp.RawFormat, Drawing.Imaging.ImageFormat.MemoryBmp), Drawing.Imaging.ImageFormat.Gif, ResultBmp.RawFormat), Drawing.Imaging.ImageFormat))
                DiskCache.CacheItem(CacheURL, DateModified, MemStream.GetBuffer())
            End If
            'save thumb to cache
        End If
        If ResultBmp Is Nothing Then
            GetThumbSizeFromURL = SizeF.Empty
        Else
            GetThumbSizeFromURL = ResultBmp.GetBounds(Drawing.GraphicsUnit.Pixel).Size
            ResultBmp.Dispose()
        End If
    End Function
    Public Shared Sub AddTextLogo(ByVal bmp As Bitmap, ByVal Text As String)
        Dim g As Graphics = Graphics.FromImage(bmp)
        Dim FontSize As Integer = 0
        Dim oFont As Font
        Dim TextExtent As SizeF
        g.PageUnit = GraphicsUnit.Pixel
        g.TextRenderingHint = Drawing.Text.TextRenderingHint.SystemDefault
        g.TextContrast = 0
        Do
            FontSize += 1
            oFont = New Font("Arial", FontSize, FontStyle.Regular, GraphicsUnit.Pixel)
            TextExtent = g.MeasureString(Text, oFont, New PointF(0, 0), Drawing.StringFormat.GenericTypographic)
            If TextExtent.Width > bmp.GetBounds(Drawing.GraphicsUnit.Pixel).Size.Width Then
                oFont.Dispose()
                oFont = New Font("Arial", FontSize - 1, FontStyle.Regular, GraphicsUnit.Pixel)
                TextExtent = g.MeasureString(Text, oFont, New PointF(0, 0), Drawing.StringFormat.GenericTypographic)
                Exit Do
            End If
        Loop While TextExtent.Width < bmp.GetBounds(Drawing.GraphicsUnit.Pixel).Size.Width * 4 / 5
        Dim Format As StringFormat = Drawing.StringFormat.GenericTypographic
        Format.LineAlignment = StringAlignment.Center
        Format.Alignment = StringAlignment.Center
        g.DrawString(Text, oFont, New SolidBrush(Color.FromArgb(128, Color.MintCream)), New RectangleF(0, CSng(bmp.GetBounds(Drawing.GraphicsUnit.Pixel).Size.Height - Math.Ceiling(TextExtent.Height) - 2), bmp.GetBounds(Drawing.GraphicsUnit.Pixel).Size.Width, CSng(Math.Ceiling(TextExtent.Height))), Format)
        g.Dispose()
        oFont.Dispose()
    End Sub
    Public Shared Function LookupClassMember(ByVal Text As String) As Reflection.MethodInfo
        Dim ClassMember As String() = Text.Split(":"c)
        LookupClassMember = Nothing
        If (ClassMember.Length = 3 AndAlso ClassMember(1) = String.Empty) Then
            Dim CheckType As Type = Reflection.Assembly.GetExecutingAssembly().GetType("HostPageUtility." + ClassMember(0))
            If Not CheckType Is Nothing Then
                LookupClassMember = CheckType.GetMethod(ClassMember(2))
            End If
            If LookupClassMember Is Nothing Then
                For Each Key As String In ConnectionData.FuncLibs
                    Dim Asm As Reflection.Assembly = Reflection.Assembly.Load(Key)
                    If Not Asm Is Nothing Then
                        CheckType = Asm.GetType(Key + "." + ClassMember(0))
                        If Not CheckType Is Nothing Then
                            LookupClassMember = CheckType.GetMethod(ClassMember(2))
                            If Not LookupClassMember Is Nothing Then Exit For
                        End If
                    End If
                Next
            End If
        End If
    End Function
    Public Shared Function TextRender(ByVal Item As PageLoader.TextItem) As String
        Return Item.Text
    End Function
    Public Shared Function GetOnPrintJS() As String()
        Return New String() {"javascript: openPrintable(this);", String.Empty, "function openPrintable(btn) { var input = document.createElement('input'); input.type = 'hidden'; input.name = 'PagePrint'; input.value = btn.form.elements['Page'].value; btn.form.appendChild(input); btn.form.target = '_blank'; btn.form.elements['Page'].value = 'PrintPdf'; btn.form.submit(); btn.form.target = ''; btn.form.elements['Page'].value = btn.form.elements['PagePrint'].value; btn.form.removeChild(input); }"}
    End Function
    Public Shared Function GetClearOptionListJS() As String
        Return "function clearOptionList(selectObject) { while (selectObject.options.length) { selectObject.options.remove(selectObject.options.length - 1); } }"
    End Function
    Public Shared Function GetLookupStyleSheetJS() As String
        Return "function findStyleSheetRule(ruleName) { var iCount, iIndex; for (iCount = 0; iCount < document.styleSheets.length; iCount++) { var rules = document.styleSheets.item(iCount); rules = rules.cssRules || rules.rules; for (var iIndex = 0; iIndex < rules.length; iIndex++) { if (rules.item(iIndex).selectorText == ruleName) { return rules.item(iIndex); } } } return null; }"
    End Function
    Public Shared Function GetBrowserTestJS() As String
        Return "var isChrome = /chrome/.test(navigator.userAgent.toLowerCase()); var isMac = /mac/i.test(navigator.platform); var isSafari = /Safari/i.test(navigator.userAgent);"
    End Function
    Public Shared Function IsInArrayJS() As String
        Return "function isInArray(array, item) { var length = array.length; for (var count = 0; count < length; count++) { if (array[count] === item) return true; } return false; }"
    End Function
    Public Shared Function GetAddStyleSheetJS() As String
        Return "function newStyleSheet() { if (document.createStyleSheet) return document.createStyleSheet(); else { var newSE = document.createElement('style'); newSE.type = 'text/css'; $('head').get(0).appendChild(newSE); return newSE.styleSheet ? newSE.styleSheet : (newSE.sheet ? newSE.sheet : newSE); } }"
    End Function
    Public Shared Function GetAddStyleSheetRuleJS() As String
        Return "function addStyleSheetRule(sheet, selectorText, ruleText) { if (sheet.tagName) { sheet.innerHTML = sheet.innerHTML + selectorText + ' {' + ruleText + '}'; } else if (sheet.addRule) { if (selectorText == '@font-face' && sheet.cssText != null) sheet.cssText = selectorText + ' {' + ruleText + '}'; else sheet.addRule(selectorText, ruleText); } else if (sheet.insertRule) sheet.insertRule(selectorText + ' {' + ruleText + '}', sheet.cssRules.length); }"
    End Function
    Public Shared Function ParseValue(ByVal XMLItemNode As System.Xml.XmlNode, ByVal DefaultValue As String) As String
        If XMLItemNode Is Nothing Then
            ParseValue = DefaultValue
        Else
            ParseValue = XMLItemNode.Value
        End If
    End Function
    Public Shared Function GetChildNode(ByVal NodeName As String, ByVal ChildNodes As System.Xml.XmlNodeList) As System.Xml.XmlNode
        Dim XMLNode As System.Xml.XmlNode
        For Each XMLNode In ChildNodes
            If XMLNode.Name = NodeName Then
                Return XMLNode
            End If
        Next
        Return Nothing
    End Function
    Public Shared Function GetChildNodes(ByVal NodeName As String, ByVal ChildNodes As System.Xml.XmlNodeList) As System.Xml.XmlNode()
        Dim XMLNode As System.Xml.XmlNode
        Dim XMLNodeList As New ArrayList
        For Each XMLNode In ChildNodes
            If XMLNode.Name = NodeName Then
                XMLNodeList.Add(XMLNode)
            End If
        Next
        Return DirectCast(XMLNodeList.ToArray(GetType(System.Xml.XmlNode)), System.Xml.XmlNode())
    End Function
    Public Shared Function GetChildNodeByIndex(ByVal NodeName As String, ByVal IndexName As String, ByVal Index As Integer, ByVal ChildNodes As System.Xml.XmlNodeList) As System.Xml.XmlNode
        Dim XMLNode As System.Xml.XmlNode = ChildNodes.Item(Index)
        Dim AttributeNode As System.Xml.XmlNode
        If Index - 1 < ChildNodes.Count Then
            XMLNode = ChildNodes.Item(Index - 1)
            If XMLNode.Name = NodeName Then
                AttributeNode = XMLNode.Attributes.GetNamedItem(IndexName)
                If Not AttributeNode Is Nothing AndAlso CInt(AttributeNode.Value) = Index Then
                    Return XMLNode
                End If
            End If
        End If
        For Each XMLNode In ChildNodes
            If XMLNode.Name = NodeName Then
                AttributeNode = XMLNode.Attributes.GetNamedItem(IndexName)
                If Not AttributeNode Is Nothing AndAlso CInt(AttributeNode.Value) = Index Then
                    Return XMLNode
                End If
            End If
        Next
        Return Nothing
    End Function
    Public Shared Function GetChildNodeCount(ByVal NodeName As String, ByVal Node As System.Xml.XmlNode) As Integer
        Dim Index As Integer
        Dim Count As Integer = 0
        For Index = 0 To Node.ChildNodes.Count - 1
            If Node.ChildNodes.Item(Index).Name = NodeName Then Count += 1
        Next
        Return Count
    End Function
End Class
Public Class DiskCache
    Shared Function GetCacheDirectory() As String
        Dim Path As String
        Path = IO.Path.Combine(HttpRuntime.CodegenDir, "DiskCache")
        If Not IO.Directory.Exists(Path) Then IO.Directory.CreateDirectory(Path)
        Return Path
    End Function
    Public Shared Function GetCacheItem(ByVal Name As String, ByVal ModifiedUtc As Date) As Byte()
        If Not IO.File.Exists(IO.Path.Combine(GetCacheDirectory, Name)) OrElse _
            ModifiedUtc > IO.File.GetLastWriteTimeUtc(IO.Path.Combine(GetCacheDirectory(), Name)) Then Return Nothing
        Dim File As IO.FileStream = IO.File.Open(IO.Path.Combine(GetCacheDirectory(), Name), IO.FileMode.Open, IO.FileAccess.Read)
        Dim Bytes(CInt(File.Length) - 1) As Byte
        File.Read(Bytes, 0, CInt(File.Length))
        File.Close()
        Return Bytes
    End Function
    Public Shared Function TransmitCacheItem(ByVal Name As String, ByVal ModifiedUtc As Date) As Boolean
        If Not IO.File.Exists(IO.Path.Combine(GetCacheDirectory, Name)) OrElse _
            ModifiedUtc > IO.File.GetLastWriteTimeUtc(IO.Path.Combine(GetCacheDirectory(), Name)) Then Return False
        HttpContext.Current.Response.TransmitFile(IO.Path.Combine(GetCacheDirectory(), Name))
        Return True
    End Function
    Public Shared Sub CacheItem(ByVal Name As String, ByVal LastModifiedUtc As Date, ByVal Data() As Byte)
        Dim File As IO.FileStream = IO.File.Open(IO.Path.Combine(GetCacheDirectory(), Name), IO.FileMode.Create, IO.FileAccess.Write)
        File.Write(Data, 0, Data.Length)
        File.Close()
        If LastModifiedUtc = DateTime.MinValue Then LastModifiedUtc = DateTime.Now
        IO.File.SetLastWriteTimeUtc(IO.Path.Combine(GetCacheDirectory(), Name), LastModifiedUtc)
    End Sub
    Public Shared Function GetCacheItems() As String()
        Return IO.Directory.GetFiles(GetCacheDirectory())
    End Function
    Public Shared Sub DeleteUnusedCacheItems(ByVal ActiveNames() As String)
        Dim Files() As String = IO.Directory.GetFiles(GetCacheDirectory())
        Dim Count As Integer
        For Count = 0 To Files.Length - 1
            If Array.IndexOf(ActiveNames, Files(Count)) = -1 Then DeleteCacheItem(Files(Count))
        Next
    End Sub
    Public Shared Sub DeleteCacheItem(ByVal Name As String)
        IO.File.Delete(Name)
    End Sub
End Class
Public Class UnitConversions
    Public Shared Function GetPIJS() As String
        Return "function getPI() { return 3.14159265358979323846; }" + _
        "function degToDegMinSec(deg) { return Math.floor(deg).toString() + '\u00B0 ' + Math.floor((deg % 1) * 60).toString() + '\' ' + ((((deg % 1) * 60) % 1) * 60).toString() + '""'; }" + _
        "function degMinSecToDeg(deg, min, sec) { return Number(deg) + min / 60 + sec / 3600; }" + _
        "function degToRad(deg) { return deg * getPI() /  180; }" + _
        "function radToDeg(rad) { return rad * 180 / getPI(); }" + _
        "function toBearing(rad) { return (radToDeg(rad) + 360) % 360; }" + _
        "function getSphericalDistance(lat1, lon1, lat2, lon2) { var R = 6378.137, dLon = degToRad(lon2 - lon1), dLat = degToRad(lat2 - lat1); a = Math.sin(dLat / 2) * Math.sin(dLat / 2) + Math.sin(dLon / 2) * Math.sin(dLon / 2) * Math.cos(degToRad(lat1)) * Math.cos(degToRad(lat2)); return R * 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a)); }" + _
        "function getDegreeBearing(lat1, lon1, lat2, lon2) { var dLon = degToRad(lon1 - lon2); return toBearing(-Math.atan2(Math.sin(dLon), Math.cos(degToRad(lat1)) * Math.tan(degToRad(lat2)) - Math.sin(degToRad(lat1)) * Math.cos(dLon))); }"
    End Function
    Public Shared Function GetTimeDistToSpeedJS() As String()
        Return New String() {"javascript: doTimeDistToSpeed();", String.Empty, _
        "function doTimeDistToSpeed() { $('#speedresult').text($('#mode0').prop('checked') ? ($('#distance').val() / (($('#minutes').val() * 60 + parseFloat($('#seconds').val())) / 3600 + parseFloat($('#hours').val()))) : ((parseFloat($('#minutes').val()) + (parseFloat($('#seconds').val()) + $('#hours').val() * 3600) / 60) / $('#distance').val())); }"}
    End Function
    Public Shared Function GetUnitConversionJS() As String()
        Return New String() {"javascript: doUnitConversion();", String.Empty, _
        "function doUnitConversion() { $('#distanceresult').text($('#convunits0').prop('checked') ? ($('#convdistance').val() / 1.609344) : ($('#convdistance').val() * 1.609344)); }"}
    End Function
    Public Shared Function GetDegreeConversionUnitChangeJS() As String()
        Return New String() {"javascript: doDegreeConversionUnitChange();", String.Empty, _
        "function doDegreeConversionUnitChange() { $('#minutes').css('display', $('#convunits0').prop('checked') ? 'block' : 'none'); $('#seconds').css('display', $('#convunits0').prop('checked') ? 'block' : 'none'); }"}
    End Function
    Public Shared Function GetDegreeConversionJS() As String()
        Return New String() {"javascript: doDegreeConversion();", String.Empty, GetPIJS(), _
        "function doDegreeConversion() { $('#result').text($('#convunits0').prop('checked') ? degMinSecToDeg($('#degrees').val(), $('#minutes').val(), $('#seconds').val()).toString() + '\u00B0\r\n' + degToRad(degMinSecToDeg($('#degrees').val(), $('#minutes').val(), $('#seconds').val())).toString() + 'rad' : ($('#convunits1').prop('checked') ? degToDegMinSec($('#degrees').val()).toString() + '\r\n' + degToRad($('#degrees').val()) + 'rad' : radToDeg($('#degrees').val()) + '\u00B0\r\n' + degToDegMinSec(radToDeg($('#degrees').val())))); }"}
    End Function
    Public Shared Function GetDegreeOffsetJS() As String()
        Return New String() {"javascript: doDegreeOffset();", String.Empty, _
        "function doDegreeOffset() { $('#resultdist').text(getSphericalDistance(degMinSecToDeg($('#latdegrees').val(), $('#latminutes').val(), $('#latseconds').val()), degMinSecToDeg($('#londegrees').val(), $('#lonminutes').val(), $('#lonseconds').val()), degMinSecToDeg($('#destlatdegrees').val(), $('#destlatminutes').val(), $('#destlatseconds').val()), degMinSecToDeg($('#destlondegrees').val(), $('#destlonminutes').val(), $('#destlonseconds').val())) + 'km\r\n' + getDegreeBearing(degMinSecToDeg($('#latdegrees').val(), $('#latminutes').val(), $('#latseconds').val()), degMinSecToDeg($('#londegrees').val(), $('#lonminutes').val(), $('#lonseconds').val()), degMinSecToDeg($('#destlatdegrees').val(), $('#destlatminutes').val(), $('#destlatseconds').val()), degMinSecToDeg($('#destlondegrees').val(), $('#destlonminutes').val(), $('#destlonseconds').val())) + '\u00B0'); }"}
    End Function
    Public Shared Function GetDateOffsetJS() As String()
        Return New String() {"javascript: doDateOffset();", String.Empty, _
        "function doDateOffset() { var d = new Date($('#date').val()); d.setDate($('#convdates0').prop('checked') ? (d.getDate() - (-$('#offset').val())) : (d.getDate() - $('#offset').val())); $('#resultdate').text(d.toDateString()); }"}
    End Function
    Public Shared Function GetDataConversionJS() As String()
        Return New String() {"javascript: doDateConversion();", String.Empty, _
        "function doDateConversion() { $('#resultcal').text($.calendars.instance($('#convcalendars0').prop('checked') ? 'gregorian' : ($('#convcalendars1').prop('checked') ? 'islamic' : 'ummalqura')).fromJD($.calendars.instance('gregorian').fromJSDate(new Date($('#dateconv').val())).toJD()).formatDate($.calendars.instance().FULL)); }"}
    End Function
End Class
Public Class XMLCoding
    Public Shared Function PerformCoding() As String()
        Return New String() {"javascript: doCoding();", String.Empty, _
        "function doCoding() { $('#xmlresult').text($('#convdir0').prop('checked') ? ($('#convattr0').prop('checked') ? $('#convxml').val().replace(/&/g, '&amp;').replace(/\r/g, '&#13;').replace(/\n/g, '&#10;') : $('#convxml').val().replace(/&/g, '&amp;')).replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/\'/g, '&apos;').replace(/\""/g, '&quot;') : ($('#convattr0').prop('checked') ? $('#convxml').val().replace(/&#13;/g, '\r').replace(/&#10;/g, '\n') : $('#convxml').val()).replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&apos;/g, '\'').replace(/&quot;/g, '\""').replace(/&amp;/g, '&')); }"}
    End Function
End Class
Public Class HTTPCoding
    Public Shared Function XmlEncode(Str As String) As String
        Return Str.Replace("&", "&amp;").Replace("<", "&lt;").Replace(">", "&gt;").Replace("'", "&apos;").Replace("""", "&quot;")
    End Function
    Public Shared Function XmlDecode(Str As String) As String
        Return Str.Replace("&lt;", "<").Replace("&gt;", ">").Replace("&apos;", "'").Replace("&quot;", """").Replace("&amp;", "&")
    End Function
    Public Shared Function PerformCoding() As String()
        Return New String() {"javascript: doHttpCoding();", String.Empty, _
        "function doHttpCoding() { $('#httpresult').text($('#convhttpdir0').prop('checked') ? ($('#convhttpattr0').prop('checked') ? encodeURI($('#convhttp').val()) : encodeURIComponent($('#convhttp').val())) : ($('#convhttpattr0').prop('checked') ? unescape($('#convhttp').val()) : decodeURIComponent($('#convhttp').val()))); }"}
    End Function
End Class
Public Class HTMLCoding
    Public Shared Function PerformCoding() As String()
        Return New String() {"javascript: doHtmlCoding();", String.Empty, _
        "function doHtmlCoding() { $('#htmlresult').text($('#convhtmldir0').prop('checked') ? escape($('#convhtml').val()) : unescape($('#convhtml').val())); }"}
    End Function
End Class
Public Class JSCoding
    Public Shared Function PerformCoding() As String()
        'parsing javascript for beautification must ignore quotes and count the number of open braces
        'simple naive technique used for now
        Return New String() {"javascript: doJSCoding();", String.Empty, _
        "function doJSCoding() { $('#jsresult').text($('#convjsdir0').prop('checked') ? $('#convjs').val().replace(/\r?\n|\r|\t/gm, ' ') : $('#convjs').val().replace(/(?:{|;|})\s+/g, '$1\r\n')); }"}
    End Function
End Class
Public Class Document
    Public Shared Function GetDocument(ByVal Item As PageLoader.TextItem) As String
        Return String.Empty
    End Function
    Public Shared Function GetXML(ByVal Item As PageLoader.TextItem) As Array
        Dim XMLDoc As New System.Xml.XmlDocument
        XMLDoc.Load(Utility.GetFilePath("metadata\" + Utility.ConnectionData.DocXML))
        Dim RetArray(2 + XMLDoc.DocumentElement.ChildNodes.Count) As Array
        RetArray(0) = New String() {"javascript: doOnCheck(this);", "doSort();", "function doSort() { var child = $('#render').children('tr'); child.shift(); child.sort(function(a, b) { if (window.localstorage.getItem(a.children('td')(3).text()) == window.localstorage.getItem(b.children('td')(3).text())) return new Date(a.children('td')(1).text()) > new Data(b.children('td')(1).text()); return (window.localstorage.getItem(a.children('td')(3).text())) ? 1 : -1; }); child.detach().appendTo($('#render')); } function doOnCheck(element) { element.checked = !element.checked; if (element.checked) { window.localstorage.setItem($(element).parent().children('td')(3), true); } else { window.localstorage.removeItem($(element).parent().children('td')(3)); } doSort(); }"}
        RetArray(1) = New String() {"check", String.Empty, String.Empty, "hidden"}
        RetArray(2) = New String() {String.Empty, String.Empty, String.Empty, String.Empty}
        For Count As Integer = 0 To XMLDoc.DocumentElement.ChildNodes.Count - 1
            RetArray(2 + Count) = New String() {"Separate?", XMLDoc.DocumentElement.ChildNodes.Item(Count).Attributes("date").Value, XMLDoc.DocumentElement.ChildNodes.Item(Count).Attributes("message").Value, XMLDoc.DocumentElement.ChildNodes.Item(Count).Attributes("id").Value}
        Next
        Return RetArray
    End Function
End Class
Public Class Geolocation
    Public Shared Function GetIP(ByVal Item As PageLoader.TextItem) As String
        Return HttpContext.Current.Request.UserHostAddress
    End Function
    Public Shared Function GetGeoData() As String()
        Dim URL As String = "http://api.ipinfodb.com/v3/ip-city/?key=" + Utility.ConnectionData.IPInfoDBAPIKey + "&ip=" + HttpContext.Current.Request.UserHostAddress
        Dim MyWebRequest As Net.HttpWebRequest = DirectCast(Net.WebRequest.Create(URL), Net.HttpWebRequest)
        Dim Data As String = String.Empty
        Try
            Dim Response As Net.HttpWebResponse = DirectCast(MyWebRequest.GetResponse(), Net.HttpWebResponse)
            Dim DataStream As IO.StreamReader = New IO.StreamReader(Response.GetResponseStream())
            Try
                Data = DataStream.ReadToEnd()
            Catch
            End Try
            Response.Close()
        Catch
        End Try
        Return Data.Split(";"c)
    End Function
    Public Shared Function GetGeoInfo(ByVal Item As PageLoader.TextItem) As Array()
        Dim Strings As String() = GetGeoData()
        If Strings.Length <> 11 Then Return New Array() {}
        Return New Array() {New String() {"Status Code", "Status Message", "IP Address", "Country Code", "Country Name", "Region Name", "City Name", "Zip Code", "Latitude", "Longitude", "Time Zone"}, New String() {Strings(0), Strings(1), Strings(2), Strings(3), Strings(4), Strings(5), Strings(6), Strings(7), Strings(8), Strings(9), Strings(10)}}
    End Function
    Public Shared Function GetElevationData(ByVal lat As String, ByVal lng As String) As String
        Dim URL As String = "http://maps.googleapis.com/maps/api/elevation/xml?locations=" + lat + "," + lng + "&sensor=false"
        Dim MyWebRequest As Net.HttpWebRequest = DirectCast(Net.WebRequest.Create(URL), Net.HttpWebRequest)
        Dim Data As String = String.Empty
        Try
            Dim Response As Net.HttpWebResponse = DirectCast(MyWebRequest.GetResponse(), Net.HttpWebResponse)
            Dim DataStream As IO.StreamReader = New IO.StreamReader(Response.GetResponseStream())
            Try
                Data = DataStream.ReadToEnd()
            Catch
            End Try
            Response.Close()
        Catch
        End Try
        Dim XMLDoc As New System.Xml.XmlDocument
        Dim XMLChildNode As System.Xml.XmlNode
        Dim XMLNode As System.Xml.XmlNode
        XMLDoc.LoadXml(Data)
        For Each XMLNode In XMLDoc.DocumentElement.ChildNodes
            If XMLNode.Name = "result" Then
                For Each XMLChildNode In XMLNode
                    If XMLChildNode.Name = "elevation" Then
                        Return XMLChildNode.InnerText
                    End If
                    If XMLChildNode.Name = "resolution" Then
                        'Return XMLChildNode.InnerText
                    End If
                Next
            End If
        Next
        Return String.Empty
    End Function
    Public Shared Function GetElevation(ByVal Item As PageLoader.TextItem) As String
        Dim Strings As String() = GetGeoData()
        If Strings.Length <> 11 Then Return String.Empty
        Return GetElevationData(Strings(8), Strings(9))
    End Function
End Class
<CLSCompliant(True)> _
Public Class PageLoader
    Structure PageItem
        Dim Page As ArrayList
        Dim PageName As String
        Dim Text As String
        Public Sub New(ByVal NewPage As ArrayList, ByVal NewPageName As String, ByVal NewText As String)
            Page = NewPage
            PageName = NewPageName
            Text = NewText
        End Sub
    End Structure
    Structure ListItem
        Dim List As ArrayList
        Dim Title As String
        Dim Name As String
        Dim IsSection As Boolean
        Dim HasForm As Boolean
        Dim FormPostURL As String
        Public Sub New(ByVal NewTitle As String, ByVal NewName As String, ByVal NewList As ArrayList, ByVal NewIsSection As Boolean, ByVal NewHasForm As Boolean, Optional ByVal NewFormPostURL As String = "")
            List = NewList
            Name = NewName
            Title = NewTitle
            IsSection = NewIsSection
            HasForm = NewHasForm
            FormPostURL = NewFormPostURL
        End Sub
    End Structure
    Enum ContentType
        eImage
        eText
        eDownload
    End Enum
    Structure EmailItem
        Dim UseImage As Boolean
        Public Sub New(ByVal NewUseImage As Boolean)
            UseImage = NewUseImage
        End Sub
    End Structure
    Structure TextItem
        Dim Name As String
        Dim Text As String
        Dim URL As String
        Dim ImageURL As String
        Dim OnRenderFunction As Reflection.MethodInfo
        Public Sub New(ByVal NewName As String, ByVal NewText As String, Optional ByVal NewURL As String = "", Optional ByVal NewImageURL As String = "", Optional ByVal NewOnRender As String = "")
            Name = NewName
            Text = NewText
            URL = NewURL
            ImageURL = NewImageURL
            If NewOnRender <> String.Empty Then OnRenderFunction = Utility.LookupClassMember(NewOnRender)
        End Sub
    End Structure
    Structure EditItem
        Dim Name As String
        Dim DefaultValue As String
        Dim Rows As Integer
        Dim Password As Boolean
        Public Sub New(ByVal NewName As String, ByVal NewDefaultValue As String, ByVal NewRows As Integer, Optional ByVal NewPassword As Boolean = False)
            Name = NewName
            DefaultValue = NewDefaultValue
            Rows = NewRows
            Password = NewPassword
        End Sub
    End Structure
    Structure DateItem
        Dim Name As String
        Dim Description As String
        Public Sub New(ByVal NewName As String, ByVal NewDescription As String, Optional ByVal NewOnClick As String = "")
            Name = NewName
            Description = NewDescription
        End Sub
    End Structure
    Structure ButtonItem
        Dim Name As String
        Dim Description As String
        Dim OnClickFunction As Reflection.MethodInfo
        Dim OnRenderFunction As Reflection.MethodInfo
        Public Sub New(ByVal NewName As String, ByVal NewDescription As String, Optional ByVal NewOnClick As String = "", Optional ByVal NewOnRender As String = "")
            Name = NewName
            Description = NewDescription
            If NewOnClick <> String.Empty Then OnClickFunction = Utility.LookupClassMember(NewOnClick)
            If NewOnRender <> String.Empty Then OnRenderFunction = Utility.LookupClassMember(NewOnRender)
        End Sub
    End Structure
    Structure RadioItem
        Dim Name As String
        Dim Description As String
        Dim UseList As Boolean
        Dim UseCheck As Boolean
        Dim DefaultValue As String
        Dim OptionArray() As OptionItem
        Dim OnPopulateFunction As Reflection.MethodInfo
        Dim OnChangeFunction As Reflection.MethodInfo
        Public Sub New(ByVal NewName As String, ByVal NewDescription As String, ByVal NewDefaultValue As String, ByVal NewOptionArray() As OptionItem, Optional ByVal NewUseList As Boolean = False, Optional ByVal NewUseCheck As Boolean = False, Optional ByVal NewOnPopulate As String = "", Optional ByVal NewOnChange As String = "")
            Name = NewName
            Description = NewDescription
            DefaultValue = NewDefaultValue
            OptionArray = NewOptionArray
            UseList = NewUseList
            UseCheck = NewUseCheck
            If NewOnPopulate <> String.Empty Then OnPopulateFunction = Utility.LookupClassMember(NewOnPopulate)
            If NewOnChange <> String.Empty Then OnChangeFunction = Utility.LookupClassMember(NewOnChange)
        End Sub
    End Structure
    Structure OptionItem
        Dim Name As String
        Public Sub New(ByVal NewName As String)
            Name = NewName
        End Sub
    End Structure
    Structure ImageItem
        Dim Name As String
        Dim Text As String
        Dim Path As String
        Dim Link As Boolean
        Dim MaxX As Integer
        Dim MaxY As Integer
        Public Sub New(ByVal NewName As String, ByVal NewText As String, ByVal NewPath As String, Optional ByVal NewLink As Boolean = False, Optional ByVal NewMaxX As Integer = 0, Optional ByVal NewMaxY As Integer = 0)
            Name = NewName
            Text = NewText
            Path = NewPath
            Link = NewLink
            MaxX = NewMaxX
            MaxY = NewMaxY
        End Sub
    End Structure
    Structure DownloadItem
        Dim Text As String
        Dim Path As String
        Dim OnRenderFunction As Reflection.MethodInfo
        Dim RelativePath As Boolean
        Dim UseLink As Boolean
        Dim ShowInline As Boolean
        Public Sub New(ByVal NewText As String, ByVal NewPath As String, ByVal NewOnRender As String, Optional ByVal NewRelativePath As Boolean = True, Optional ByVal NewUseLink As Boolean = False, Optional ByVal NewShowInline As Boolean = False)
            Text = NewText
            Path = NewPath
            RelativePath = NewRelativePath
            UseLink = NewUseLink
            ShowInline = NewShowInline
            If NewOnRender <> String.Empty Then OnRenderFunction = Utility.LookupClassMember(NewOnRender)
        End Sub
    End Structure
    Public Shared Function IsDownloadItem(ByVal Item As Object) As Boolean
        IsDownloadItem = TypeOf Item Is DownloadItem
    End Function
    Public Shared Function IsEmailItem(ByVal Item As Object) As Boolean
        IsEmailItem = TypeOf Item Is EmailItem
    End Function
    Public Shared Function IsImageItem(ByVal Item As Object) As Boolean
        IsImageItem = TypeOf Item Is ImageItem
    End Function
    Public Shared Function IsEditItem(ByVal Item As Object) As Boolean
        IsEditItem = TypeOf Item Is EditItem
    End Function
    Public Shared Function IsDateItem(ByVal Item As Object) As Boolean
        IsDateItem = TypeOf Item Is DateItem
    End Function
    Public Shared Function IsRadioItem(ByVal Item As Object) As Boolean
        IsRadioItem = TypeOf Item Is RadioItem
    End Function
    Public Shared Function IsButtonItem(ByVal Item As Object) As Boolean
        IsButtonItem = TypeOf Item Is ButtonItem
    End Function
    Public Shared Function IsTextItem(ByVal Item As Object) As Boolean
        IsTextItem = TypeOf Item Is TextItem
    End Function
    Public Shared Function IsListItem(ByVal Item As Object) As Boolean
        IsListItem = TypeOf Item Is ListItem
    End Function
    Public Pages As New Collections.Generic.List(Of PageItem)
    Public Title As String
    Public MainImage As String
    Public HoverImage As String
    Public Function GetPage(ByVal Name As String) As PageItem
        Dim Count As Integer
        For Count = 0 To Pages.Count - 1
            If Pages(Count).PageName = Name Then Return Pages(Count)
        Next
        Return Pages(0) 'default page is 0
    End Function
    Public Function GetPageIndex(ByVal Name As String) As Integer
        Dim Index As Integer
        For Index = 0 To Pages.Count - 1
            If (Name Is Nothing Or Name = String.Empty) And Index = 0 Or _
                Name <> String.Empty And Name = (Pages.Item(Index).PageName) Then
                Return Index
            End If
        Next
        Return 0
    End Function
    Public Shared Function GetItem(ByVal Name As String, ByVal Item As ArrayList) As Object
        Dim Count As Integer
        For Count = 0 To Item.Count - 1
            If IsListItem(Item(Count)) Then
                If DirectCast(Item(Count), ListItem).Name = Name Then Return Item(Count)
            ElseIf IsImageItem(Item(Count)) Then
                If DirectCast(Item(Count), ImageItem).Name = Name Then Return Item(Count)
            ElseIf IsTextItem(Item(Count)) Then
                If DirectCast(Item(Count), TextItem).Name = Name Then Return Item(Count)
            End If
        Next
        Return Item(0) 'default item should be an image or text item
    End Function
    Public Function GetPageItem(ByVal Path As String) As Object
        Dim Index As Integer
        Dim StrArray As String() = Path.Split("."c)
        Dim Item As ArrayList = GetPage(StrArray(0)).Page
        Dim ObjItem As Object = Item
        For Index = 1 To StrArray.Length - 1
            ObjItem = PageLoader.GetItem(StrArray(Index), Item)
            If PageLoader.IsListItem(ObjItem) Then
                Item = DirectCast(ObjItem, ListItem).List
            End If
        Next
        Return ObjItem
    End Function
    Sub ParseSingleElement(ByRef XMLChildNode As System.Xml.XmlNode, ByRef List As ArrayList, ByVal IsTopLevel As Boolean)
        If XMLChildNode.Name = "frame" Then
            Dim XMLListNode As System.Xml.XmlNode
            Dim ListArray As New ArrayList
            For Each XMLListNode In XMLChildNode.ChildNodes
                ParseSingleElement(XMLListNode, ListArray, False)
            Next
            List.Add(New ListItem( _
                XMLChildNode.Attributes.GetNamedItem("description").Value, _
                XMLChildNode.Attributes.GetNamedItem("name").Value, ListArray, IsTopLevel, _
                Utility.ParseValue(XMLChildNode.Attributes.GetNamedItem("hasform"), "false") = "true"))
        ElseIf XMLChildNode.Name = "button" Then
            List.Add(New ButtonItem(XMLChildNode.Attributes.GetNamedItem("name").Value, _
                                    XMLChildNode.Attributes.GetNamedItem("description").Value, _
                                    Utility.ParseValue(XMLChildNode.Attributes.GetNamedItem("onclick"), String.Empty), _
                                    Utility.ParseValue(XMLChildNode.Attributes.GetNamedItem("onrender"), String.Empty)))
        ElseIf XMLChildNode.Name = "edit" Then
            List.Add(New EditItem(XMLChildNode.Attributes.GetNamedItem("name").Value, _
                                  Utility.ParseValue(XMLChildNode.Attributes.GetNamedItem("defaultvalue"), String.Empty), _
                                  CInt(Utility.ParseValue(XMLChildNode.Attributes.GetNamedItem("rows"), "1"))))
        ElseIf XMLChildNode.Name = "date" Then
            List.Add(New DateItem(XMLChildNode.Attributes.GetNamedItem("name").Value, _
                                  XMLChildNode.Attributes.GetNamedItem("description").Value))
        ElseIf XMLChildNode.Name = "radio" Then
            Dim XMLOptionNode As System.Xml.XmlNode
            Dim OptionArray As New ArrayList
            Dim DefaultValue As String = Utility.ParseValue(XMLChildNode.Attributes.GetNamedItem("defaultvalue"), "-1")
            For Each XMLOptionNode In XMLChildNode.ChildNodes
                If XMLOptionNode.Name = "option" Then
                    If Utility.ParseValue(XMLOptionNode.Attributes.GetNamedItem("defaultvalue"), "false") = "true" Then
                        DefaultValue = CStr(OptionArray.Count)
                    End If
                    OptionArray.Add(New OptionItem(XMLOptionNode.Attributes.GetNamedItem("name").Value))
                End If
            Next
            List.Add(New RadioItem(XMLChildNode.Attributes.GetNamedItem("name").Value, _
                                   XMLChildNode.Attributes.GetNamedItem("description").Value, _
                                   DefaultValue, _
                                   DirectCast(OptionArray.ToArray(GetType(OptionItem)), OptionItem()), _
                                   Utility.ParseValue(XMLChildNode.Attributes.GetNamedItem("uselist"), "false") = "true", _
                                   Utility.ParseValue(XMLChildNode.Attributes.GetNamedItem("usecheck"), "false") = "true", _
                                   Utility.ParseValue(XMLChildNode.Attributes.GetNamedItem("onpopulate"), String.Empty), _
                                   Utility.ParseValue(XMLChildNode.Attributes.GetNamedItem("onchange"), String.Empty)))
        ElseIf XMLChildNode.Name = "ipaddr" Then
        ElseIf XMLChildNode.Name = "static" Then
            List.Add(New TextItem( _
                            XMLChildNode.Attributes.GetNamedItem("name").Value, _
                            XMLChildNode.Attributes.GetNamedItem("description").Value, _
                            Utility.ParseValue(XMLChildNode.Attributes.GetNamedItem("url"), String.Empty), _
                            Utility.ParseValue(XMLChildNode.Attributes.GetNamedItem("imageurl"), String.Empty), _
                            Utility.ParseValue(XMLChildNode.Attributes.GetNamedItem("onrender"), String.Empty)))
        ElseIf XMLChildNode.Name = "image" Then
            List.Add(New ImageItem( _
                Utility.ParseValue(XMLChildNode.Attributes.GetNamedItem("name"), String.Empty), _
                XMLChildNode.Attributes.GetNamedItem("text").Value, _
                XMLChildNode.Attributes.GetNamedItem("source").Value, _
                Utility.ParseValue(XMLChildNode.Attributes.GetNamedItem("usethumbonmax"), "false") = "true", _
                CInt(Utility.ParseValue(XMLChildNode.Attributes.GetNamedItem("maxwidth"), "0")), _
                CInt(Utility.ParseValue(XMLChildNode.Attributes.GetNamedItem("maxheight"), "0"))))
        ElseIf XMLChildNode.Name = "email" Then
            List.Add(New EmailItem( _
               Utility.ParseValue(XMLChildNode.Attributes.GetNamedItem("useimage"), "true") = "true"))
        ElseIf XMLChildNode.Name = "download" Then
            List.Add(New DownloadItem( _
                XMLChildNode.Attributes.GetNamedItem("text").Value, _
                XMLChildNode.Attributes.GetNamedItem("path").Value, _
                Utility.ParseValue(XMLChildNode.Attributes.GetNamedItem("onrender"), String.Empty), _
                Utility.ParseValue(XMLChildNode.Attributes.GetNamedItem("userelativepath"), "true") = "true", _
                Utility.ParseValue(XMLChildNode.Attributes.GetNamedItem("userelativepath"), "true") <> "true", _
                Utility.ParseValue(XMLChildNode.Attributes.GetNamedItem("showinline"), "false") = "true"))
        End If
    End Sub
    Public Sub New()
        Dim XMLDoc As New System.Xml.XmlDocument
        Dim XMLNode As System.Xml.XmlNode
        Dim XMLChildNode As System.Xml.XmlNode
        XMLDoc.Load(Utility.GetTemplatePath())
        Title = Utility.ParseValue(XMLDoc.DocumentElement.Attributes.GetNamedItem("title"), String.Empty)
        MainImage = Utility.ParseValue(XMLDoc.DocumentElement.Attributes.GetNamedItem("mainimage"), String.Empty)
        HoverImage = Utility.ParseValue(XMLDoc.DocumentElement.Attributes.GetNamedItem("hoverimage"), String.Empty)
        Dim PageList As ArrayList
        For Each XMLNode In XMLDoc.DocumentElement.ChildNodes
            If XMLNode.Name = "page" Then
                PageList = New ArrayList
                For Each XMLChildNode In XMLNode.ChildNodes
                    If XMLChildNode.Name = "child" Then
                    ElseIf XMLChildNode.Name = "addlist" Then
                    Else
                        ParseSingleElement(XMLChildNode, PageList, True)
                    End If
                Next
                Pages.Add(New PageItem(PageList, XMLNode.Attributes.GetNamedItem("name").Value, _
                                       XMLNode.Attributes.GetNamedItem("description").Value))
            End If
        Next
    End Sub
End Class
Public Class ArabicData
    <System.Xml.Serialization.XmlRoot("arabicdata")> _
    Class ArabicXMLData
        Public Structure ArabicCombo
            <System.Xml.Serialization.XmlAttribute("symbolname")> _
            Public SymbolName As String
            <System.Xml.Serialization.XmlAttribute("uname")> _
            Public UnicodeName As String
            Public Shaping() As Char
            <System.Xml.Serialization.XmlAttribute("shaping")> _
            Property ShapingParse As String
                Get
                    If Shaping.Length = 0 Then Return String.Empty
                    Return String.Join(","c, Array.ConvertAll(Shaping, Function(Ch As Char) Asc(Ch).ToString("X2")))
                End Get
                Set(value As String)
                    If Not value Is Nothing Then
                        Shaping = Array.ConvertAll(value.Split(","c), Function(Str As String) If(Str = String.Empty, ChrW(0), ChrW(Integer.Parse(Str, System.Globalization.NumberStyles.HexNumber))))
                    End If
                End Set
            End Property
        End Structure
        <System.Xml.Serialization.XmlArray("arabiccombos")> _
        <System.Xml.Serialization.XmlArrayItem("combo")> _
        Public ArabicCombos() As ArabicCombo
        Public Structure ArabicSymbol
            <System.Xml.Serialization.XmlAttribute("symbolname")> _
            Public SymbolName As String
            <System.Xml.Serialization.XmlAttribute("uname")> _
            Public UnicodeName As String
            Public Symbol As Char
            <System.Xml.Serialization.XmlAttribute("symbol")> _
            Property SymbolParse As String
                Get
                    Return Asc(Symbol).ToString("X2")
                End Get
                Set(value As String)
                    Symbol = ChrW(Integer.Parse(value, System.Globalization.NumberStyles.HexNumber))
                End Set
            End Property
            Public Shaping() As Char
            <System.Xml.Serialization.XmlAttribute("shaping")> _
            Property ShapingParse As String
                Get
                    If Shaping.Length = 0 Then Return String.Empty
                    Return String.Join(","c, Array.ConvertAll(Shaping, Function(Ch As Char) Asc(Ch).ToString("X2")))
                End Get
                Set(value As String)
                    If Not value Is Nothing Then
                        Shaping = Array.ConvertAll(value.Split(","c), Function(Str As String) If(Str = String.Empty, ChrW(0), ChrW(Integer.Parse(Str, System.Globalization.NumberStyles.HexNumber))))
                    End If
                End Set
            End Property
            <System.Xml.Serialization.XmlAttribute("ipavalue")> _
            Public IPAValue As String
            <System.Xml.Serialization.XmlAttribute("extendedbuckwalter")> _
            Public ExtendedBuckwalterLetter As Char
            <System.Xml.Serialization.XmlAttribute("terminating")> _
            Public Terminating As Boolean
            <System.Xml.Serialization.XmlAttribute("connecting")> _
            Public Connecting As Boolean
            <System.Xml.Serialization.XmlAttribute("assimilate")> _
            Public Assimilate As Boolean
        End Structure
        <System.Xml.Serialization.XmlArray("arabicletters")> _
        <System.Xml.Serialization.XmlArrayItem("arabicsymbol")> _
        Public ArabicLetters() As ArabicSymbol
    End Class
    Shared _ArabicXMLData As ArabicXMLData
    Public Shared ReadOnly Property Data As ArabicXMLData
        Get
            If _ArabicXMLData Is Nothing Then
                Dim fs As IO.FileStream = New IO.FileStream(Utility.GetFilePath("metadata\arabicdata.xml"), IO.FileMode.Open, IO.FileAccess.Read)
                Dim xs As System.Xml.Serialization.XmlSerializer = New System.Xml.Serialization.XmlSerializer(GetType(ArabicXMLData))
                _ArabicXMLData = CType(xs.Deserialize(fs), ArabicXMLData)
                fs.Close()
            End If
            Return _ArabicXMLData
        End Get
    End Property
    <CLSCompliant(False)> _
    Declare Function GetFontUnicodeRanges Lib "gdi32.dll" (ByVal hds As IntPtr, ByVal lpgs As IntPtr) As UInteger
    <CLSCompliant(False)> _
    Declare Function SelectObject Lib "gdi32.dll" (ByVal hDc As IntPtr, ByVal hObject As IntPtr) As IntPtr
    <CLSCompliant(False)> _
    Public Structure FontRange
        Public Low As UShort
        Public High As UShort
    End Structure
    <CLSCompliant(False)> _
    Shared Function Unsign(ByVal Input As Int16) As UShort
        If Input > -1 Then
            Return CUShort(Input)
        Else
            Return CUShort(UShort.MaxValue - (Not Input))
        End If
    End Function
    <CLSCompliant(False)> _
    Shared Function GetUnicodeRangesForFont(ByVal font As Font) As Generic.List(Of FontRange)
        Dim g As Graphics
        Dim hdc, hFont, old, glyphSet As IntPtr
        Dim size As UInteger
        Dim fontRanges As Generic.List(Of FontRange)
        Dim count As Integer
        g = Graphics.FromHwnd(IntPtr.Zero)
        hdc = g.GetHdc()
        hFont = font.ToHfont()
        old = SelectObject(hdc, hFont)
        size = GetFontUnicodeRanges(hdc, IntPtr.Zero)
        glyphSet = Runtime.InteropServices.Marshal.AllocHGlobal(CInt(size))
        GetFontUnicodeRanges(hdc, glyphSet)
        fontRanges = New Generic.List(Of FontRange)
        count = Runtime.InteropServices.Marshal.ReadInt32(glyphSet, 12)
        For i As Integer = 0 To count - 1
            Dim range As FontRange = New FontRange
            range.Low = Unsign(Runtime.InteropServices.Marshal.ReadInt16(glyphSet, 16 + (i * 4)))
            range.High = range.Low + Unsign(Runtime.InteropServices.Marshal.ReadInt16(glyphSet, 18 + (i * 4))) - 1US
            fontRanges.Add(range)
        Next
        SelectObject(hdc, old)
        Runtime.InteropServices.Marshal.FreeHGlobal(glyphSet)
        g.ReleaseHdc(hdc)
        g.Dispose()
        Return fontRanges
    End Function
    <CLSCompliant(False)> _
    Shared Function CheckIfCharInFont(ByVal character As Char, ByVal font As Font) As Boolean
        Dim intval As UInt16 = Convert.ToUInt16(character)
        Dim ranges As Generic.List(Of FontRange) = GetUnicodeRangesForFont(font)
        Dim isCharacterPresent As Boolean = False
        For Each range As FontRange In ranges
            If intval >= range.Low And intval <= range.High Then
                isCharacterPresent = True
                Exit For
            End If
        Next range
        Return isCharacterPresent
    End Function
    <CLSCompliant(False)> _
    Declare Function GetUName Lib "getuname.dll" (ByVal wCharCode As UShort, <Runtime.InteropServices.MarshalAs(Runtime.InteropServices.UnmanagedType.LPWStr)> ByVal lpbuf As System.Text.StringBuilder) As Integer
    Public Shared Function GetUnicodeName(Character As Char) As String
        Dim Str As New System.Text.StringBuilder(512)
        Try
            GetUName(CUShort(AscW(Character)), Str)
        Catch e As System.DllNotFoundException
            Return Utility.LoadResourceString("unicode_" + Data.ArabicLetters(FindLetterBySymbol(Character)).UnicodeName)
        End Try
        Return Str.ToString()
    End Function
    Public Shared Function IsTerminating(Str As String, Index As Integer) As Boolean
        Dim bIsEnd = True 'default to non-connecting end
        'should probably check for any non-arabic letters also
        For CharCount As Integer = Index + 1 To Str.Length - 1
            If Array.IndexOf(RecitationCombiningSymbols, Str(CharCount)) = -1 Then
                bIsEnd = False
                Exit For
            End If
        Next
        Return bIsEnd
    End Function
    Public Shared Function IsLastConnecting(Str As String, Index As Integer) As Boolean
        Dim bLastConnects = False 'default to non-connecting beginning 
        For CharCount As Integer = Index - 1 To 0 Step -1
            If Array.IndexOf(RecitationCombiningSymbols, Str(CharCount)) = -1 Then
                bLastConnects = Not Data.ArabicLetters(FindLetterBySymbol(Str(CharCount))).Terminating
                Exit For
            End If
        Next
        Return bLastConnects
    End Function
    Public Shared Function GetShapeIndex(bConnects As Boolean, bLastConnects As Boolean, bIsEnd As Boolean) As Integer
        If Not bLastConnects And (Not bConnects Or bConnects And bIsEnd) Then
            Return 0
        ElseIf bLastConnects And (Not bConnects Or bConnects And bIsEnd) Then
            Return 1
        ElseIf Not bLastConnects And bConnects And Not bIsEnd Then
            Return 2
        ElseIf bLastConnects And bConnects And Not bIsEnd Then
            Return 3
        End If
        Return -1
    End Function
    Public Shared Function GetShapeIndexFromString(Str As String, Index As Integer, Length As Integer) As Integer
        'ignore all transparent characters
        'isolated - non-connecting + (non-connecting letter | connecting letter + end)
        'final - connecting + (non-connecting letter | connecting letter + end)
        'initial - non-connecting + connecting letter + not end
        'medial - connecting + connecting letter + not end
        Dim bIsEnd = IsTerminating(Str, Index + Length - 1)
        Dim bConnects As Boolean = Not Data.ArabicLetters(FindLetterBySymbol(Str.Chars(Index + Length - 1))).Terminating
        Dim bLastConnects As Boolean = Data.ArabicLetters(FindLetterBySymbol(Str.Chars(Index + Length - 1))).Connecting And IsLastConnecting(Str, Index)
        Return GetShapeIndex(bConnects, bLastConnects, bIsEnd)
    End Function
    Public Shared Function ConvertLigatures(Str As String, Dir As Boolean) As String
        Dim Count As Integer
        Dim SubCount As Integer
        ConvertLigatures = String.Empty
        'Division seleciton between Presentation A and B forms can be done here though wasl and gunnah need consideration
        Dim Combos(ArabicData.Data.ArabicLetters.Length + ArabicData.Data.ArabicCombos.Length - 1) As ArabicXMLData.ArabicCombo
        Data.ArabicLetters.CopyTo(Data.ArabicCombos, 0)
        For Count = 0 To Data.ArabicLetters.Length - 1
            'do not need to transfer UnicodeName as it is not used here
            Combos(Data.ArabicCombos.Length + Count).SymbolName = Data.ArabicLetters(Count).Symbol
            Combos(Data.ArabicCombos.Length + Count).Shaping = Data.ArabicLetters(Count).Shaping
        Next
        Array.Sort(Combos, Function(Com1 As ArabicXMLData.ArabicCombo, Com2 As ArabicXMLData.ArabicCombo) If(Com1.SymbolName.Length = Com2.SymbolName.Length, Com1.SymbolName.CompareTo(Com2.SymbolName), If(Com1.SymbolName.Length > Com2.SymbolName.Length, -1, 1)))
        For Count = 0 To Str.Length - 1
            For SubCount = 0 To Data.ArabicCombos.Length - 1
                If Dir Then
                    If Array.IndexOf(ArabicData.Data.ArabicCombos(SubCount).Shaping, Str.Chars(Count)) <> -1 Then
                        ConvertLigatures += TransliterateFromBuckwalter(Data.ArabicCombos(SubCount).SymbolName)
                        Exit For
                    End If
                Else
                    If Str.Length - Count >= Combos(SubCount).SymbolName.Length _
                        AndAlso TransliterateFromBuckwalter(Combos(SubCount).SymbolName) = Str.Substring(Count, Combos(SubCount).SymbolName.Length) Then
                        Dim Shape As Integer = GetShapeIndexFromString(Str, Count, Combos(SubCount).SymbolName.Length)
                        If Combos(SubCount).Shaping(Shape) <> ChrW(0) Then
                            Str += Combos(SubCount).Shaping(Shape)
                        End If
                    End If
                End If
            Next
            If SubCount = Data.ArabicCombos.Length Then
                ConvertLigatures += Str.Chars(Count)
            End If
        Next
    End Function
    Public Shared Function TransliterateFromBuckwalter(ByVal Buckwalter As String) As String
        Dim ArabicString As String = String.Empty
        Dim Count As Integer
        Dim Index As Integer
        For Count = 0 To Buckwalter.Length - 1
            If Buckwalter(Count) = "\" Then
                Count += 1
                If Buckwalter(Count) = "," Then
                    ArabicString += ArabicData.ArabicComma
                Else
                    ArabicString += Buckwalter(Count)
                End If
            Else
                For Index = 0 To Data.ArabicLetters.Length - 1
                    If Buckwalter(Count) = Data.ArabicLetters(Index).ExtendedBuckwalterLetter Then
                        ArabicString += Data.ArabicLetters(Index).Symbol
                        Exit For
                    End If
                Next
                If Index = Data.ArabicLetters.Length Then
                    ArabicString += Buckwalter(Count)
                End If
            End If
        Next
        Return ArabicString
    End Function
    Public Shared Function FindLetterBySymbol(Symbol As Char) As Integer
        Return Array.FindIndex(Data.ArabicLetters, Function(Letter As ArabicXMLData.ArabicSymbol) Letter.Symbol = Symbol)
    End Function
    Public Const Space As Char = ChrW(32)
    Public Const ExclamationMark As Char = ChrW(33)
    Public Const QuotationMark As Char = ChrW(34)
    Public Const Comma As Char = ChrW(44)
    Public Const FullStop As Char = ChrW(46)
    Public Const HyphenMinus As Char = ChrW(&H2D)
    Public Const Colon As Char = ChrW(&H3A)
    Public Const LeftParenthesis As Char = ChrW(&H5B)
    Public Const RightParenthesis As Char = ChrW(&H5D)
    Public Const LeftSquareBracket As Char = ChrW(&H5B)
    Public Const RightSquareBracket As Char = ChrW(&H5D)
    Public Const LeftCurlyBracket As Char = ChrW(&H7B)
    Public Const RightCurlyBracket As Char = ChrW(&H7D)
    Public Const NoBreakSpace As Char = ChrW(&HA0)
    Public Const LeftPointingDoubleAngleQuotationMark As Char = ChrW(&HAB)
    Public Const RightPointingDoubleAngleQuotationMark As Char = ChrW(&HBB)
    Public Const ArabicComma As Char = ChrW(1548)
    Public Const ArabicLetterHamza As Char = ChrW(1569)
    Public Const ArabicLetterAlefWithMaddaAbove As Char = ChrW(1570)
    Public Const ArabicLetterAlefWithHamzaAbove As Char = ChrW(1571)
    Public Const ArabicLetterWawWithHamzaAbove As Char = ChrW(1572)
    Public Const ArabicLetterAlefWithHamzaBelow As Char = ChrW(1573)
    Public Const ArabicLetterYehWithHamzaAbove As Char = ChrW(1574)
    Public Const ArabicLetterAlef As Char = ChrW(1575)
    Public Const ArabicLetterBeh As Char = ChrW(1576)
    Public Const ArabicLetterTehMarbuta As Char = ChrW(1577)
    Public Const ArabicLetterTeh As Char = ChrW(1578)
    Public Const ArabicLetterTheh As Char = ChrW(1579)
    Public Const ArabicLetterJeem As Char = ChrW(1580)
    Public Const ArabicLetterHah As Char = ChrW(1581)
    Public Const ArabicLetterKhah As Char = ChrW(1582)
    Public Const ArabicLetterDal As Char = ChrW(1583)
    Public Const ArabicLetterThal As Char = ChrW(1584)
    Public Const ArabicLetterReh As Char = ChrW(1585)
    Public Const ArabicLetterZain As Char = ChrW(1586)
    Public Const ArabicLetterSeen As Char = ChrW(1587)
    Public Const ArabicLetterSheen As Char = ChrW(1588)
    Public Const ArabicLetterSad As Char = ChrW(1589)
    Public Const ArabicLetterDad As Char = ChrW(1590)
    Public Const ArabicLetterTah As Char = ChrW(1591)
    Public Const ArabicLetterZah As Char = ChrW(1592)
    Public Const ArabicLetterAin As Char = ChrW(1593)
    Public Const ArabicLetterGhain As Char = ChrW(1594)
    Public Const ArabicTatweel As Char = ChrW(1600)
    Public Const ArabicLetterFeh As Char = ChrW(1601)
    Public Const ArabicLetterQaf As Char = ChrW(1602)
    Public Const ArabicLetterKaf As Char = ChrW(1603)
    Public Const ArabicLetterLam As Char = ChrW(1604)
    Public Const ArabicLetterMeem As Char = ChrW(1605)
    Public Const ArabicLetterNoon As Char = ChrW(1606)
    Public Const ArabicLetterHeh As Char = ChrW(1607)
    Public Const ArabicLetterWaw As Char = ChrW(1608)
    Public Const ArabicLetterAlefMaksura As Char = ChrW(1609)
    Public Const ArabicLetterYeh As Char = ChrW(1610)

    Public Const ArabicFathatan As Char = ChrW(1611)
    Public Const ArabicDammatan As Char = ChrW(1612)
    Public Const ArabicKasratan As Char = ChrW(1613)
    Public Const ArabicFatha As Char = ChrW(1614)
    Public Const ArabicDamma As Char = ChrW(1615)
    Public Const ArabicKasra As Char = ChrW(1616)
    Public Const ArabicShadda As Char = ChrW(1617)
    Public Const ArabicSukun As Char = ChrW(1618)
    Public Const ArabicMaddahAbove As Char = ChrW(1619)
    Public Const ArabicHamzaAbove As Char = ChrW(1620)
    Public Const ArabicHamzaBelow As Char = ChrW(1621)
    Public Const ArabicVowelSignDotBelow As Char = ChrW(&H65C)
    Public Const Bullet As Char = ChrW(&H2022)
    Public Const ArabicLetterSuperscriptAlef As Char = ChrW(1648)
    Public Const ArabicLetterAlefWasla As Char = ChrW(1649)
    Public Const ArabicSmallHighLigatureSadWithLamWithAlefMaksura As Char = ChrW(1750)
    Public Const ArabicSmallHighLigatureQafWithLamWithAlefMaksura As Char = ChrW(1751)
    Public Const ArabicSmallHighMeemInitialForm As Char = ChrW(1752)
    Public Const ArabicSmallHighLamAlef As Char = ChrW(1753)
    Public Const ArabicSmallHighJeem As Char = ChrW(1754)
    Public Const ArabicSmallHighThreeDots As Char = ChrW(1755)
    Public Const ArabicSmallHighSeen As Char = ChrW(1756)
    Public Const ArabicEndOfAyah As Char = ChrW(1757)
    Public Const ArabicStartOfRubElHizb As Char = ChrW(1758)
    Public Const ArabicSmallHighRoundedZero As Char = ChrW(1759)
    Public Const ArabicSmallHighUprightRectangularZero As Char = ChrW(1760)
    Public Const ArabicSmallHighMeemIsolatedForm As Char = ChrW(1762)
    Public Const ArabicSmallLowSeen As Char = ChrW(1763)
    Public Const ArabicSmallWaw As Char = ChrW(1765)
    Public Const ArabicSmallYeh As Char = ChrW(1766)
    Public Const ArabicSmallHighNoon As Char = ChrW(1768)
    Public Const ArabicPlaceOfSajdah As Char = ChrW(1769)
    Public Const ArabicEmptyCentreLowStop As Char = ChrW(1770)
    Public Const ArabicEmptyCentreHighStop As Char = ChrW(1771)
    Public Const ArabicRoundedHighStopWithFilledCentre As Char = ChrW(1772)
    Public Const ArabicSmallLowMeem As Char = ChrW(1773)
    Public Const ArabicSemicolon As Char = ChrW(&H61B)
    Public Const ArabicLetterMark As Char = ChrW(&H61C)
    Public Const ArabicQuestionMark As Char = ChrW(&H61F)
    Public Const ArabicLetterPeh As Char = ChrW(&H67E)
    Public Const ArabicLetterTcheh As Char = ChrW(&H686)
    Public Const ArabicLetterVeh As Char = ChrW(&H6A4)
    Public Const ArabicLetterGaf As Char = ChrW(&H6AF)
    Public Const ArabicLetterNoonGhunna As Char = ChrW(&H6BA)
    Public Const ZeroWidthSpace As Char = ChrW(&H200B)
    Public Const ZeroWidthNonJoiner As Char = ChrW(&H200C)
    Public Const ZeroWidthJoiner As Char = ChrW(&H200D)
    Public Const LeftToRightMark As Char = ChrW(&H200E)
    Public Const RightToLeftMark As Char = ChrW(&H200F)
    Public Const PopDirectionalFormatting As Char = ChrW(&H202C)
    Public Const LeftToRightOverride As Char = ChrW(&H202D)
    Public Const RightToLeftOverride As Char = ChrW(&H202E)
    Public Const NarrowNoBreakSpace As Char = ChrW(&H202F)
    Public Const DottedCircle As Char = ChrW(&H25CC)
    Public Const OrnateLeftParenthesis As Char = ChrW(&HFD3E)
    Public Const OrnateRightParenthesis As Char = ChrW(&HFD3F)
    'http://www.unicode.org/Public/7.0.0/ucd/UnicodeData.txt
    Public Shared LTRCategories As String() = New String() {"L"}
    Public Shared RTLCategories As String() = New String() {"R", "AL"}
    Public Shared ALCategories As String() = New String() {"AL"}
    Public Shared NeutralCategories As String() = New String() {"B", "S", "WS", "ON"}
    Public Shared WeakCategories As String() = New String() {"EN", "ES", "ET", "AN", "CS", "NSM", "BN"}
    Public Shared ExplicitCategories As String() = New String() {"LRE", "LRO", "RLE", "RLO", "PDF", "LRI", "RLI", "FSI", "PDI"}
    Public Shared Function GetUniCats() As String
        Return "function IsLTR(c) { " + MakeUniCategory(LTRCategories) + " }" + vbCrLf + _
        "function IsRTL(c) { " + MakeUniCategory(RTLCategories) + " }" + vbCrLf + _
        "function IsAL(c) { " + MakeUniCategory(ALCategories) + " }" + vbCrLf + _
        "function IsNeutral(c) { " + MakeUniCategory(NeutralCategories) + " }" + vbCrLf + _
        "function IsWeak(c) { " + MakeUniCategory(WeakCategories) + " }" + vbCrLf + _
        "function IsExplicit(c) { " + MakeUniCategory(ExplicitCategories) + " }"
    End Function
    Public Shared Function MakeUniCategory(Cats As String()) As String
        Dim Strs As String() = IO.File.ReadAllLines("..\..\..\IslamMetadata\UnicodeData.txt")
        Dim Ranges As New ArrayList
        For Count = 0 To Strs.Length - 1
            If Array.IndexOf(Cats, Strs(Count).Split(";"c)(4)) <> -1 Then
                Dim NewRangeMatch As Integer = Integer.Parse(Strs(Count).Split(";"c)(0), Globalization.NumberStyles.AllowHexSpecifier)
                If Ranges.Count <> 0 AndAlso CInt(CType(Ranges(Ranges.Count - 1), ArrayList)(CType(Ranges(Ranges.Count - 1), ArrayList).Count - 1)) + 1 = NewRangeMatch Then
                    CType(Ranges(Ranges.Count - 1), ArrayList).Add(NewRangeMatch)
                Else
                    Ranges.Add(New ArrayList From {NewRangeMatch})
                End If
            End If
        Next
        Return "return " + String.Join("||", Array.ConvertAll(Of ArrayList, String)(CType(Ranges.ToArray(GetType(ArrayList)), ArrayList()), Function(Arr As ArrayList) If(Arr.Count = 1, "c===0x" + Hex(Arr(0)), "(c>=0x" + Hex(Arr(0)) + "&&c<=0x" + Hex(Arr(Arr.Count - 1)) + ")"))) + ";"
    End Function
    'ArabicLetterAlefWithMaddahAbove is used in simple script but not uthmani
    'ArabicEndOfAyah is added later
    'ArabicHamzaBelow never used along with ArabicData.ArabicLetterPeh, ArabicData.ArabicLetterTcheh, ArabicData.ArabicLetterVeh, ArabicData.ArabicLetterGaf
    Public Shared RecitationSymbols() As Char = {Space, _
        ArabicLetterHamza, ArabicLetterAlefWithHamzaAbove, ArabicLetterWawWithHamzaAbove, _
        ArabicLetterAlefWithHamzaBelow, ArabicLetterYehWithHamzaAbove, _
        ArabicLetterAlef, ArabicLetterBeh, ArabicLetterTehMarbuta, ArabicLetterTeh, _
        ArabicLetterTheh, ArabicLetterJeem, ArabicLetterHah, ArabicLetterKhah, ArabicLetterDal,
        ArabicLetterThal, ArabicLetterReh, ArabicLetterZain, ArabicLetterSeen, ArabicLetterSheen, _
        ArabicLetterSad, ArabicLetterDad, ArabicLetterTah, ArabicLetterZah, ArabicLetterAin, _
        ArabicLetterGhain, ArabicTatweel, ArabicLetterFeh, ArabicLetterQaf, ArabicLetterKaf, _
        ArabicLetterLam, ArabicLetterMeem, ArabicLetterNoon, ArabicLetterHeh, ArabicLetterWaw, _
        ArabicLetterAlefMaksura, ArabicLetterYeh, ArabicFathatan, ArabicDammatan, ArabicKasratan, _
        ArabicFatha, ArabicDamma, ArabicKasra, ArabicShadda, ArabicSukun, ArabicMaddahAbove, _
        ArabicHamzaAbove, ArabicLetterSuperscriptAlef, ArabicLetterAlefWasla, _
        ArabicSmallHighLigatureSadWithLamWithAlefMaksura, _
        ArabicSmallHighLigatureQafWithLamWithAlefMaksura, ArabicSmallHighMeemInitialForm, _
        ArabicSmallHighLamAlef, ArabicSmallHighJeem, ArabicSmallHighThreeDots, _
        ArabicSmallHighSeen, ArabicStartOfRubElHizb, ArabicSmallHighRoundedZero, _
        ArabicSmallHighUprightRectangularZero, ArabicSmallHighMeemIsolatedForm, _
        ArabicSmallLowSeen, ArabicSmallWaw, ArabicSmallYeh, ArabicSmallHighNoon, _
        ArabicPlaceOfSajdah, ArabicEmptyCentreLowStop, ArabicEmptyCentreHighStop, _
        ArabicRoundedHighStopWithFilledCentre, ArabicSmallLowMeem}
    Public Shared Function GetRecitationSymbols() As Array()
        Return Array.ConvertAll(RecitationSymbols, Function(Ch As Char) New Object() {Data.ArabicLetters(FindLetterBySymbol(Ch)).UnicodeName + " (" + FixStartingCombiningSymbol(Ch) + LeftToRightOverride + ")" + PopDirectionalFormatting, FindLetterBySymbol(Ch)})
    End Function
    Public Shared RecitationLetters As Char() = {ArabicLetterHamza, ArabicLetterAlefWithHamzaAbove, ArabicLetterWawWithHamzaAbove, _
        ArabicLetterAlefWithHamzaBelow, ArabicLetterYehWithHamzaAbove, _
        ArabicLetterAlef, ArabicLetterBeh, ArabicLetterTehMarbuta, ArabicLetterTeh, _
        ArabicLetterTheh, ArabicLetterJeem, ArabicLetterHah, ArabicLetterKhah, ArabicLetterDal,
        ArabicLetterThal, ArabicLetterReh, ArabicLetterZain, ArabicLetterSeen, ArabicLetterSheen, _
        ArabicLetterSad, ArabicLetterDad, ArabicLetterTah, ArabicLetterZah, ArabicLetterAin, _
        ArabicLetterGhain, ArabicTatweel, ArabicLetterFeh, ArabicLetterQaf, ArabicLetterKaf, _
        ArabicLetterLam, ArabicLetterMeem, ArabicLetterNoon, ArabicLetterHeh, ArabicLetterWaw, _
        ArabicLetterAlefMaksura, ArabicLetterYeh, ArabicLetterAlefWasla}
    Public Shared RecitationDiacritics As Char() = {ArabicFathatan, ArabicDammatan, ArabicKasratan, _
        ArabicFatha, ArabicDamma, ArabicKasra, ArabicShadda, ArabicSukun, ArabicMaddahAbove, _
        ArabicHamzaAbove, ArabicLetterSuperscriptAlef}
    Public Shared RecitationAlefs As Char() = {ArabicLetterAlefWithHamzaAbove, ArabicLetterAlefWithHamzaBelow, ArabicLetterAlef, _
                        ArabicLetterAlefWithMaddaAbove, ArabicLetterAlefMaksura, ArabicLetterAlefWasla, ArabicLetterSuperscriptAlef}
    Public Shared RecitationHamzas As Char() = {ArabicLetterAlefWithHamzaAbove, ArabicLetterAlefWithHamzaBelow, ArabicLetterWawWithHamzaAbove, _
                        ArabicLetterYehWithHamzaAbove, ArabicLetterHamza, ArabicHamzaAbove}
    Public Shared RecitationSpecialSymbols As Char() = {ArabicSmallHighLigatureSadWithLamWithAlefMaksura, _
        ArabicSmallHighLigatureQafWithLamWithAlefMaksura, ArabicSmallHighMeemInitialForm, _
        ArabicSmallHighLamAlef, ArabicSmallHighJeem, ArabicSmallHighThreeDots, _
        ArabicSmallHighSeen, ArabicStartOfRubElHizb, ArabicSmallHighRoundedZero, _
        ArabicSmallHighUprightRectangularZero, ArabicSmallHighMeemIsolatedForm, _
        ArabicSmallLowSeen, ArabicSmallWaw, ArabicSmallYeh, ArabicSmallHighNoon, _
        ArabicPlaceOfSajdah, ArabicEmptyCentreLowStop, ArabicEmptyCentreHighStop, _
        ArabicRoundedHighStopWithFilledCentre, ArabicSmallLowMeem}
    Public Shared RecitationSpecialSymbolsNotStop As Char() = {
        ArabicSmallHighLamAlef, _
        ArabicSmallHighSeen, ArabicStartOfRubElHizb, ArabicSmallHighRoundedZero, _
        ArabicSmallHighUprightRectangularZero, ArabicSmallHighMeemIsolatedForm, _
        ArabicSmallLowSeen, ArabicSmallWaw, ArabicSmallYeh, ArabicSmallHighNoon, _
        ArabicPlaceOfSajdah, ArabicEmptyCentreLowStop, ArabicEmptyCentreHighStop, _
        ArabicRoundedHighStopWithFilledCentre, ArabicSmallLowMeem}
    Public Shared RecitationLettersDiacritics As Char() = {ArabicLetterHamza, ArabicLetterAlefWithHamzaAbove, ArabicLetterWawWithHamzaAbove, _
        ArabicLetterAlefWithHamzaBelow, ArabicLetterYehWithHamzaAbove, _
        ArabicLetterAlef, ArabicLetterBeh, ArabicLetterTehMarbuta, ArabicLetterTeh, _
        ArabicLetterTheh, ArabicLetterJeem, ArabicLetterHah, ArabicLetterKhah, ArabicLetterDal,
        ArabicLetterThal, ArabicLetterReh, ArabicLetterZain, ArabicLetterSeen, ArabicLetterSheen, _
        ArabicLetterSad, ArabicLetterDad, ArabicLetterTah, ArabicLetterZah, ArabicLetterAin, _
        ArabicLetterGhain, ArabicTatweel, ArabicLetterFeh, ArabicLetterQaf, ArabicLetterKaf, _
        ArabicLetterLam, ArabicLetterMeem, ArabicLetterNoon, ArabicLetterHeh, ArabicLetterWaw, _
        ArabicLetterAlefMaksura, ArabicLetterYeh, ArabicFathatan, ArabicDammatan, ArabicKasratan, _
        ArabicFatha, ArabicDamma, ArabicKasra, ArabicShadda, ArabicSukun, ArabicMaddahAbove, _
        ArabicHamzaAbove, ArabicLetterSuperscriptAlef, ArabicLetterAlefWasla}
    Public Shared Function FixStartingCombiningSymbol(Str As String) As String
        Return If(Array.IndexOf(RecitationCombiningSymbols, Str.Chars(0)) <> -1 Or Str.Length = 1, LeftToRightOverride + Str + PopDirectionalFormatting, Str)
    End Function
    'Kasra when following shadda may actually be below the shadda but above the character
    Public Shared RecitationCombiningSymbolsBelow As Char() = {ArabicKasratan, ArabicKasra, ArabicSmallLowSeen, ArabicEmptyCentreLowStop, ArabicSmallLowMeem, ArabicHamzaBelow}
    Public Shared RecitationCombiningSymbols As Char() = {ArabicFathatan, ArabicDammatan, ArabicKasratan, _
        ArabicFatha, ArabicDamma, ArabicKasra, ArabicShadda, ArabicSukun, ArabicMaddahAbove, _
        ArabicHamzaAbove, ArabicLetterSuperscriptAlef, _
        ArabicSmallHighLigatureSadWithLamWithAlefMaksura, _
        ArabicSmallHighLigatureQafWithLamWithAlefMaksura, ArabicSmallHighMeemInitialForm, _
        ArabicSmallHighLamAlef, ArabicSmallHighJeem, ArabicSmallHighThreeDots, _
        ArabicSmallHighSeen, ArabicSmallHighRoundedZero, _
        ArabicSmallHighUprightRectangularZero, ArabicSmallHighMeemIsolatedForm, _
        ArabicSmallLowSeen, ArabicSmallHighNoon, ArabicEmptyCentreLowStop, ArabicEmptyCentreHighStop, _
        ArabicRoundedHighStopWithFilledCentre, ArabicSmallLowMeem}
    Public Shared RecitationConnectingFollowerSymbols As Char() = {ArabicLetterAlefWithHamzaAbove, ArabicLetterWawWithHamzaAbove, _
        ArabicLetterAlefWithHamzaBelow, ArabicLetterYehWithHamzaAbove, _
        ArabicLetterAlef, ArabicLetterBeh, ArabicLetterTehMarbuta, ArabicLetterTeh, _
        ArabicLetterTheh, ArabicLetterJeem, ArabicLetterHah, ArabicLetterKhah, ArabicLetterDal,
        ArabicLetterThal, ArabicLetterReh, ArabicLetterZain, ArabicLetterSeen, ArabicLetterSheen, _
        ArabicLetterSad, ArabicLetterDad, ArabicLetterTah, ArabicLetterZah, ArabicLetterAin, _
        ArabicLetterGhain, ArabicTatweel, ArabicLetterFeh, ArabicLetterQaf, ArabicLetterKaf, _
        ArabicLetterLam, ArabicLetterMeem, ArabicLetterNoon, ArabicLetterHeh, ArabicLetterWaw, _
        ArabicLetterAlefMaksura, ArabicLetterYeh}
    Public Shared Function IsLetter(Index As Integer) As Boolean
        Return Array.FindIndex(ArabicLetters, Function(Str As String) Str = Data.ArabicLetters(Index).Symbol) <> -1
    End Function
    Public Shared Function IsPunctuation(Index As Integer) As Boolean
        Return Array.FindIndex(PunctuationSymbols, Function(Str As String) Str = Data.ArabicLetters(Index).Symbol) <> -1
    End Function
    Public Shared Function IsStop(Index As Integer) As Boolean
        Return Array.FindIndex(ArabicStopLetters, Function(Str As String) Str = Data.ArabicLetters(Index).Symbol) <> -1
    End Function
    Public Shared Function IsWhitespace(Index As Integer) As Boolean
        Return Array.FindIndex(WhitespaceSymbols, Function(Str As String) Str = Data.ArabicLetters(Index).Symbol) <> -1
    End Function
    Public Shared ArabicUniqueLetters As String() = {"Al^m^", "Al^m^S^", "Al^r", "Al^m^r", "k^hyE^S^", "Th", "Ts^m^", "Ts^", "ys^", "S^", "Hm^", "E^s^q^", "q^", "n^"}
    Public Shared ArabicNumbers As String() = {"0", "1", "2", "3", "4", "5", "6", "7", "8", "9"}
    Public Shared ArabicWaslKasraExceptions As String() = {"{mo$uwA", "{}otuwA", "{qoDuwA", "{bonuwA", "{moDuwA", "{mora>ata", "{somu", "{vonatayni", "{vonayni", "{bonatu", "{bonu", "{moru&NA"}
    Public Shared ArabicFathaDammaKasra As String() = {ArabicFatha, ArabicDamma, ArabicKasra}
    Public Shared ArabicTanweens As String() = {ArabicFathatan, ArabicDammatan, ArabicKasratan}
    Public Shared ArabicLongVowels As String() = {ArabicFatha + ArabicLetterAlef, ArabicDamma + ArabicLetterWaw, ArabicKasra + ArabicLetterYeh}
    Public Shared ArabicSunLetters As String() = {ArabicLetterTeh, ArabicLetterTheh, ArabicLetterDal, ArabicLetterThal, ArabicLetterReh, ArabicLetterZain, ArabicLetterSeen, ArabicLetterSheen, ArabicLetterSad, ArabicLetterDad, ArabicLetterTah, ArabicLetterZah, ArabicLetterLam, ArabicLetterNoon}
    Public Shared ArabicSunLettersNoLam As String() = {ArabicLetterTeh, ArabicLetterTheh, ArabicLetterDal, ArabicLetterThal, ArabicLetterReh, ArabicLetterZain, ArabicLetterSeen, ArabicLetterSheen, ArabicLetterSad, ArabicLetterDad, ArabicLetterTah, ArabicLetterZah, ArabicLetterNoon}
    Public Shared ArabicMoonLetters As String() = {ArabicLetterAlef, ArabicLetterBeh, ArabicLetterJeem, ArabicLetterHah, ArabicLetterKhah, ArabicLetterAin, ArabicLetterGhain, ArabicLetterFeh, ArabicLetterQaf, ArabicLetterKaf, ArabicLetterMeem, ArabicLetterHeh, ArabicLetterWaw, ArabicLetterYeh}
    Public Shared ArabicMoonLettersNoVowels As String() = {ArabicLetterBeh, ArabicLetterJeem, ArabicLetterHah, ArabicLetterKhah, ArabicLetterAin, ArabicLetterGhain, ArabicLetterFeh, ArabicLetterQaf, ArabicLetterKaf, ArabicLetterMeem, ArabicLetterHeh}
    Public Shared ArabicSpecialLeadingGutteral As String() = {ArabicLetterHamza, ArabicLetterHah, ArabicLetterAin}
    Public Shared ArabicSpecialGutteral As String() = {ArabicLetterHah, ArabicLetterAin, ArabicLetterSad, ArabicLetterDad, ArabicLetterTah, ArabicLetterZah}
    Public Shared ArabicLetters As String() = {ArabicLetterTeh, ArabicLetterTheh, ArabicLetterDal, ArabicLetterThal, ArabicLetterReh, ArabicLetterZain, ArabicLetterSeen, ArabicLetterSheen, ArabicLetterSad, ArabicLetterDad, ArabicLetterTah, ArabicLetterZah, ArabicLetterLam, ArabicLetterNoon, ArabicLetterAlef, ArabicLetterBeh, ArabicLetterJeem, ArabicLetterHah, ArabicLetterKhah, ArabicLetterAin, ArabicLetterGhain, ArabicLetterFeh, ArabicLetterQaf, ArabicLetterKaf, ArabicLetterMeem, ArabicLetterHeh, ArabicLetterWaw, ArabicLetterYeh}
    Public Shared ArabicLettersInOrder As Char() = {ArabicLetterAlef, ArabicLetterBeh, ArabicLetterTeh, ArabicLetterTheh, ArabicLetterJeem, ArabicLetterHah, ArabicLetterKhah, ArabicLetterDal, ArabicLetterThal, ArabicLetterReh, ArabicLetterZain, ArabicLetterSeen, ArabicLetterSheen, ArabicLetterSad, ArabicLetterDad, ArabicLetterTah, ArabicLetterZah, ArabicLetterAin, ArabicLetterGhain, ArabicLetterFeh, ArabicLetterQaf, ArabicLetterKaf, ArabicLetterLam, ArabicLetterMeem, ArabicLetterNoon, ArabicLetterHeh, ArabicLetterWaw, ArabicLetterYeh}
    Public Shared ArabicHamzas As Char() = {ArabicLetterHamza, ArabicLetterAlefWithMaddaAbove, ArabicLetterAlefWithHamzaAbove, ArabicLetterWawWithHamzaAbove, ArabicLetterAlefWithHamzaBelow, ArabicLetterYehWithHamzaAbove, ArabicHamzaAbove, ArabicLetterAlefWasla, ArabicHamzaBelow}
    Public Shared ArabicSpecialLetters As Char() = {ArabicLetterTehMarbuta, ArabicLetterTehMarbuta, ArabicLetterTehMarbuta, ArabicLetterAlefMaksura, ArabicLetterSuperscriptAlef, ArabicLetterNoonGhunna}
    Public Shared ArabicVowels As String() = {ArabicFatha, ArabicDamma, ArabicKasra, ArabicFathatan, ArabicDammatan, ArabicKasratan, ArabicFatha + ArabicLetterAlef, ArabicDamma + ArabicLetterWaw, ArabicKasra + ArabicLetterYeh, ArabicFatha + ArabicLetterWaw, ArabicFatha + ArabicLetterYeh, ArabicShadda, ArabicSukun}
    Public Shared ArabicLeadingGutterals As String() = {ArabicFatha + ArabicLetterAlef, ArabicFatha, ArabicDamma + ArabicLetterWaw, ArabicDamma, ArabicKasra + ArabicLetterYeh, ArabicKasra, ArabicFathatan, ArabicDammatan, ArabicKasratan, ArabicSukun}
    Public Shared ArabicTrailingGutterals As String() = {ArabicFatha + ArabicLetterAlef, ArabicFatha, ArabicDamma + ArabicLetterWaw, ArabicDamma, ArabicKasra + ArabicLetterYeh, ArabicKasra}
    Public Shared ArabicTajweed As Char() = {ArabicSmallHighSeen, ArabicSmallHighMeemIsolatedForm, ArabicSmallLowSeen, ArabicSmallWaw, ArabicSmallYeh, ArabicSmallHighNoon, ArabicSmallLowMeem}
    Public Shared ArabicPunctuation As Char() = {Space, ExclamationMark, QuotationMark, Comma, HyphenMinus, FullStop, Colon, LeftSquareBracket, RightSquareBracket, LeftCurlyBracket, RightCurlyBracket, NoBreakSpace, LeftPointingDoubleAngleQuotationMark, RightPointingDoubleAngleQuotationMark, ArabicComma, ArabicSemicolon, ArabicQuestionMark, ZeroWidthNonJoiner, NarrowNoBreakSpace, OrnateLeftParenthesis, OrnateRightParenthesis}
    Public Shared NonArabicLetters As Char() = {ArabicLetterPeh, ArabicLetterTcheh, ArabicLetterVeh, ArabicLetterGaf}
    Public Shared PunctuationSymbols As Char() = {ExclamationMark, QuotationMark, FullStop, Comma, ArabicComma, OrnateLeftParenthesis, OrnateRightParenthesis}
    Public Shared ArabicPunctuationSymbols As String() = {ArabicComma, OrnateLeftParenthesis, OrnateRightParenthesis}
    Public Shared Function ArabicTranslitLetters() As String()
        Dim Lets As New List(Of String)
        Lets.AddRange(Array.ConvertAll(ArabicLettersInOrder, Function(Ch As Char) CStr(Ch)))
        Lets.AddRange(Array.ConvertAll(ArabicHamzas, Function(Ch As Char) CStr(Ch)))
        Lets.AddRange(Array.ConvertAll(ArabicSpecialLetters, Function(Ch As Char) CStr(Ch)))
        Lets.AddRange(ArabicVowels)
        Lets.AddRange(ArabicLeadingGutterals)
        Lets.AddRange(ArabicTrailingGutterals)
        Lets.AddRange(Array.ConvertAll(ArabicTajweed, Function(Ch As Char) CStr(Ch)))
        Lets.AddRange(Array.ConvertAll(ArabicPunctuation, Function(Ch As Char) CStr(Ch)))
        Lets.AddRange(Array.ConvertAll(NonArabicLetters, Function(Ch As Char) CStr(Ch)))
        Return Lets.ToArray()
    End Function
    Public Shared WhitespaceSymbols As Char() = {Space}
    Public Shared ArabicStopLetters As Char() = {ArabicSmallHighLigatureSadWithLamWithAlefMaksura, _
                                                   ArabicSmallHighLigatureQafWithLamWithAlefMaksura, _
                                ArabicSmallHighMeemInitialForm, ArabicSmallHighLamAlef, _
                                ArabicSmallHighJeem, ArabicSmallHighThreeDots, ArabicSmallHighSeen}
    Public Shared Function MakeUniRegEx(Input As String) As String
        Return String.Join(String.Empty, Array.ConvertAll(Of Char, String)(Input.ToCharArray(), Function(Ch As Char) "\u" + AscW(Ch).ToString("X4")))
    End Function
    Public Shared Function MakeRegMultiEx(Input As String()) As String
        Return String.Join("|", Input)
    End Function
    Public Enum TranslitScheme
        None = 0
        Literal = 1
        RuleBased = 2
    End Enum
End Class
Public Class RenderArray
    Enum RenderTypes
        eHeaderLeft
        eHeaderCenter
        eHeaderRight
        eText
        eInteractive
    End Enum
    Enum RenderDisplayClass
        eNested
        eArabic
        eTransliteration
        eLTR
        eRTL
        eContinueStop
        eRanking
        eList
    End Enum
    Structure RenderText
        Public DisplayClass As RenderDisplayClass
        Public Clr As Color
        Public Text As Object
        Sub New(ByVal NewDisplayClass As RenderDisplayClass, ByVal NewText As Object)
            DisplayClass = NewDisplayClass
            Text = NewText
            Clr = Color.Black 'default
        End Sub
    End Structure
    Structure RenderItem
        Public Type As RenderTypes
        Public TextItems() As RenderText
        Sub New(ByVal NewType As RenderTypes, ByVal NewTextItems() As RenderText)
            Type = NewType
            TextItems = NewTextItems
        End Sub
    End Structure
    Public Items As New Collections.Generic.List(Of RenderItem)
    Structure LayoutInfo
        Public Sub New(NewRect As RectangleF, NewBaseline As Single, NewNChar As Integer, NewBounds As Generic.List(Of Generic.List(Of Generic.List(Of LayoutInfo))))
            Rect = NewRect
            Baseline = NewBaseline
            nChar = NewNChar
            Bounds = NewBounds
        End Sub
        Dim Rect As RectangleF
        Dim Baseline As Single
        Dim nChar As Integer
        Dim Bounds As Generic.List(Of Generic.List(Of Generic.List(Of LayoutInfo)))
    End Structure
    Public Shared Sub DoRenderPdf(Doc As iTextSharp.text.Document, Writer As iTextSharp.text.pdf.PdfWriter, Font As iTextSharp.text.Font, DrawFont As Drawing.Font, CurRenderArray As List(Of HostPageUtility.RenderArray.RenderItem), _Bounds As Generic.List(Of Generic.List(Of Generic.List(Of LayoutInfo))), ByRef PageOffset As PointF, BaseOffset As PointF)
        For Count As Integer = 0 To CurRenderArray.Count - 1
            Dim MaxRect As New RectangleF(Doc.PageSize.Width, Doc.PageSize.Height, 0, 0)
            For SubCount As Integer = 0 To CurRenderArray(Count).TextItems.Length - 1
                Dim NextCount As Integer
                For NextCount = 0 To _Bounds(Count)(SubCount).Count - 1
                    If _Bounds(Count)(SubCount)(NextCount).Rect.Top + PageOffset.Y + BaseOffset.Y > Doc.PageSize.Height - Doc.BottomMargin - Doc.TopMargin Then
                        Writer.DirectContent.SetLineWidth(1)
                        Writer.DirectContent.Rectangle(MaxRect.Left + Doc.LeftMargin - 2, Doc.PageSize.Height - Doc.TopMargin - MaxRect.Bottom + Font.BaseFont.GetFontDescriptor(iTextSharp.text.pdf.BaseFont.AWT_LEADING, Font.Size) * 2, MaxRect.Width - 2, MaxRect.Height - Font.BaseFont.GetFontDescriptor(iTextSharp.text.pdf.BaseFont.AWT_LEADING, Font.Size) * 2)
                        Writer.DirectContent.Stroke()
                        MaxRect = New RectangleF(Doc.PageSize.Width, Doc.PageSize.Height, 0, 0)
                        Doc.NewPage()
                        PageOffset.Y = -_Bounds(Count)(0)(0).Rect.Top - BaseOffset.Y
                        Exit For
                    End If
                Next
                If NextCount <> _Bounds(Count)(SubCount).Count Then Exit For
            Next
            For SubCount As Integer = 0 To CurRenderArray(Count).TextItems.Length - 1
                If CurRenderArray(Count).TextItems(SubCount).DisplayClass = RenderArray.RenderDisplayClass.eNested Then
                    DoRenderPdf(Doc, Writer, Font, DrawFont, CType(CurRenderArray(Count).TextItems(SubCount).Text, List(Of RenderArray.RenderItem)), _Bounds(Count)(SubCount)(0).Bounds, PageOffset, New PointF(_Bounds(Count)(SubCount)(0).Rect.Location.X, _Bounds(Count)(SubCount)(0).Rect.Location.Y))
                ElseIf CurRenderArray(Count).TextItems(SubCount).DisplayClass = RenderArray.RenderDisplayClass.eArabic Or CurRenderArray(Count).TextItems(SubCount).DisplayClass = HostPageUtility.RenderArray.RenderDisplayClass.eLTR Or CurRenderArray(Count).TextItems(SubCount).DisplayClass = RenderArray.RenderDisplayClass.eRTL Or CurRenderArray(Count).TextItems(SubCount).DisplayClass = RenderArray.RenderDisplayClass.eTransliteration Then
                    Dim theText As String = CStr(CurRenderArray(Count).TextItems(SubCount).Text)
                    For NextCount As Integer = 0 To _Bounds(Count)(SubCount).Count - 1
                        If _Bounds(Count)(SubCount)(NextCount).Rect.Top + PageOffset.Y + BaseOffset.Y > Doc.PageSize.Height - Doc.BottomMargin - Doc.TopMargin Then
                            Writer.DirectContent.SetLineWidth(1)
                            Writer.DirectContent.Rectangle(MaxRect.Left + Doc.LeftMargin - 2, Doc.PageSize.Height - Doc.TopMargin - MaxRect.Bottom + Font.BaseFont.GetFontDescriptor(iTextSharp.text.pdf.BaseFont.AWT_LEADING, Font.Size) * 2, MaxRect.Width - 2, MaxRect.Height - Font.BaseFont.GetFontDescriptor(iTextSharp.text.pdf.BaseFont.AWT_LEADING, Font.Size) * 2)
                            Writer.DirectContent.Stroke()
                            MaxRect = New RectangleF(Doc.PageSize.Width, Doc.PageSize.Height, 0, 0)
                            Doc.NewPage()
                            PageOffset.Y = -_Bounds(Count)(SubCount)(0).Rect.Top - BaseOffset.Y
                            Exit For
                        End If
                    Next
                    For NextCount As Integer = 0 To _Bounds(Count)(SubCount).Count - 1
                        Dim Rect As RectangleF = _Bounds(Count)(SubCount)(NextCount).Rect
                        Dim Text As String = theText.Substring(0, _Bounds(Count)(SubCount)(NextCount).nChar)
                        Rect.Offset(BaseOffset)
                        Rect.Offset(PageOffset)
                        MaxRect.X = Math.Min(MaxRect.Left, Rect.Left)
                        MaxRect.Y = Math.Min(MaxRect.Top, Rect.Top)
                        MaxRect.Width = Math.Max(MaxRect.Right, Rect.Right) - MaxRect.Left + 1
                        MaxRect.Height = Math.Max(MaxRect.Bottom, Rect.Bottom) - MaxRect.Top + 1
                        Dim ct As iTextSharp.text.pdf.ColumnText
                        Dim Index As Integer = 0
                        Do
                            Index = Text.IndexOf(ArabicData.ArabicEndOfAyah, Index)
                            If Index <> -1 Then
                                Dim NumCount As Integer = 0
                                Do While NumCount <> 3 And (Index + NumCount + 1) <= Text.Length - 1 AndAlso Char.IsDigit(Text(Index + NumCount + 1))
                                    NumCount += 1
                                Loop
                                Dim s As New SizeF
                                Dim Baseline As Single
                                GetTextWidthPdf(Font, Text.Substring(0, Index + 1 + NumCount), Doc.PageSize.Width, True, s, Baseline)
                                Dim ChBounds As Integer() = Font.BaseFont.GetCharBBox(AscW(Text(Index)))
                                Dim n As New SizeF
                                GetTextWidthPdf(Font, Text.Substring(Index + 1, NumCount), Doc.PageSize.Width, True, n, Baseline)
                                ct = New iTextSharp.text.pdf.ColumnText(Writer.DirectContent)
                                ct.RunDirection = iTextSharp.text.pdf.PdfWriter.RUN_DIRECTION_RTL
                                ct.ArabicOptions = iTextSharp.text.pdf.ColumnText.AR_COMPOSEDTASHKEEL
                                ct.UseAscender = False
                                Dim FitFont As New iTextSharp.text.Font(Font)
                                FitFont.Size = (ChBounds(2) + ChBounds(0)) * 0.001F * Font.Size / n.Width * Font.Size
                                GetTextWidthPdf(FitFont, Text.Substring(Index + 1, NumCount), Doc.PageSize.Width, True, n, Baseline)
                                ct.SetSimpleColumn(Rect.Left + Doc.LeftMargin + Rect.Width + 1 - (ChBounds(2) - ChBounds(0)) * 0.001F * Font.Size / 4, Doc.PageSize.Height - Doc.TopMargin - Rect.Bottom - _Bounds(Count)(SubCount)(NextCount).Baseline, Rect.Right + 1 + Doc.LeftMargin - s.Width - (ChBounds(2) - ChBounds(0)) * 0.001F * Font.Size / 4, Doc.PageSize.Height - Doc.TopMargin - Rect.Top + 1 - _Bounds(Count)(SubCount)(NextCount).Baseline, Font.BaseFont.GetFontDescriptor(iTextSharp.text.pdf.BaseFont.AWT_LEADING, Font.Size), iTextSharp.text.Element.ALIGN_CENTER Or iTextSharp.text.Element.ALIGN_BASELINE)
                                ct.AddText(New iTextSharp.text.Chunk(Text.Substring(Index + 1, NumCount), FitFont))
                                ct.Go()
                                Text = Text.Remove(Index + 1, NumCount)
                                Index = Index + 1
                            End If
                        Loop While Index <> -1
                        Index = 1
                        Do
                            Index = Text.IndexOfAny(ArabicData.RecitationCombiningSymbols, Index)
                            If Index <> -1 Then
                                Dim s As New SizeF
                                Dim Baseline As Single
                                Dim ChBounds As Integer()
                                If (Text(Index - 1) = " "c) Then
                                    'stopping symbols handled by normal rendering engine
                                    Index += 1
                                    Continue Do
                                ElseIf Text(Index - 1) = ArabicData.ArabicTatweel Then
                                    ChBounds = Font.BaseFont.GetCharBBox(AscW(Text(Index - 1)))
                                Else
                                    Dim ShapeCh As Char = ArabicData.Data.ArabicLetters(ArabicData.FindLetterBySymbol(Text(Index - 1))).Shaping(ArabicData.GetShapeIndexFromString(Text, Index - 1, 1))
                                    ChBounds = Font.BaseFont.GetCharBBox(AscW(ShapeCh))
                                End If
                                'partial shaping will never work
                                'must either convert all to shaped characters or subtract last character
                                GetTextWidthPdf(Font, Text.Substring(0, Index), Doc.PageSize.Width, True, s, Baseline)
                                If Text(Index - 1) <> " "c And Text(Index - 1) <> ArabicData.ArabicTatweel And ArabicData.GetShapeIndexFromString(Text, Index - 1, 1) = 2 Then
                                    s.Width -= Font.BaseFont.GetCharBBox(AscW(ArabicData.Data.ArabicLetters(ArabicData.FindLetterBySymbol(Text(Index - 1))).Shaping(0)))(2) * 0.001F * Font.Size
                                ElseIf Text(Index - 1) <> " "c And Text(Index - 1) <> ArabicData.ArabicTatweel And ArabicData.GetShapeIndexFromString(Text, Index - 1, 1) = 3 Then
                                    s.Width -= Font.BaseFont.GetCharBBox(AscW(ArabicData.Data.ArabicLetters(ArabicData.FindLetterBySymbol(Text(Index - 1))).Shaping(1)))(2) * 0.001F * Font.Size
                                Else
                                    s.Width -= ChBounds(2) * 0.001F * Font.Size
                                End If
                                Dim LastCenter As Integer = 0
                                Dim Offset As Integer = 0
                                Do
                                    ct = New iTextSharp.text.pdf.ColumnText(Writer.DirectContent)
                                    ct.RunDirection = iTextSharp.text.pdf.PdfWriter.RUN_DIRECTION_RTL
                                    ct.ArabicOptions = iTextSharp.text.pdf.ColumnText.AR_COMPOSEDTASHKEEL
                                    ct.UseAscender = False
                                    Dim DiaBounds As Integer() = Font.BaseFont.GetCharBBox(AscW(Text(Index)))
                                    Offset = If(DiaBounds(1) < 0 And DiaBounds(3) > If(Offset < 0, Offset, ChBounds(1)), -(DiaBounds(3) - If(Offset < 0, Offset, ChBounds(1))), If(DiaBounds(1) >= 0 And DiaBounds(1) < If(Offset > 0, Offset, ChBounds(3)), -(DiaBounds(1) - If(Offset > 0, Offset, ChBounds(3))), 0))
                                    ct.SetSimpleColumn(Rect.Left + Doc.LeftMargin + (Rect.Width - 3 - s.Width - (ChBounds(2) - If(LastCenter <> 0, LastCenter - (DiaBounds(2) - DiaBounds(0)) \ 2, 0)) * 0.001F * Font.Size), Doc.PageSize.Height - Doc.TopMargin - Rect.Bottom - _Bounds(Count)(SubCount)(NextCount).Baseline + Offset * 0.001F * Font.Size, Rect.Right + 1 + Doc.LeftMargin - s.Width + If(LastCenter <> 0, LastCenter + (DiaBounds(2) - DiaBounds(0)) \ 2, 0) * 0.001F * Font.Size, Doc.PageSize.Height - Doc.TopMargin - Rect.Top + 1 - _Bounds(Count)(SubCount)(NextCount).Baseline + Offset * 0.001F * Font.Size, Font.BaseFont.GetFontDescriptor(iTextSharp.text.pdf.BaseFont.AWT_LEADING, Font.Size), iTextSharp.text.Element.ALIGN_RIGHT Or iTextSharp.text.Element.ALIGN_BASELINE)
                                    ct.AddText(New iTextSharp.text.Chunk(Text(Index), Font))
                                    ct.Go()
                                    If DiaBounds(1) < 0 Then Offset = DiaBounds(1) - Offset
                                    If DiaBounds(1) >= 0 Then Offset = DiaBounds(3) + Offset
                                    LastCenter += (DiaBounds(2) - DiaBounds(0)) \ 2
                                    Index += 1
                                Loop While Index <> Text.Length AndAlso Array.IndexOf(ArabicData.RecitationCombiningSymbols, Text(Index)) <> -1
                            End If
                        Loop While Index <> -1
                        Text = System.Text.RegularExpressions.Regex.Replace(Text, "(?<!\s+)(?:" + ArabicData.MakeRegMultiEx(Array.ConvertAll(ArabicData.RecitationCombiningSymbols, Function(Ch As Char) CStr(Ch))) + ")", String.Empty)
                        ct = New iTextSharp.text.pdf.ColumnText(Writer.DirectContent)
                        If CurRenderArray(Count).TextItems(SubCount).DisplayClass = RenderArray.RenderDisplayClass.eArabic Or CurRenderArray(Count).TextItems(SubCount).DisplayClass = RenderArray.RenderDisplayClass.eRTL Then
                            ct.RunDirection = iTextSharp.text.pdf.PdfWriter.RUN_DIRECTION_RTL
                            ct.ArabicOptions = iTextSharp.text.pdf.ColumnText.AR_COMPOSEDTASHKEEL
                            ct.UseAscender = False
                        Else
                            ct.RunDirection = iTextSharp.text.pdf.PdfWriter.RUN_DIRECTION_LTR
                        End If
                        ct.SetSimpleColumn(Rect.Left + Doc.LeftMargin, Doc.PageSize.Height - Doc.TopMargin - Rect.Bottom - _Bounds(Count)(SubCount)(NextCount).Baseline, Rect.Right + 1 + Doc.LeftMargin, Doc.PageSize.Height - Doc.TopMargin - Rect.Top + 1 - _Bounds(Count)(SubCount)(NextCount).Baseline, Font.BaseFont.GetFontDescriptor(iTextSharp.text.pdf.BaseFont.AWT_LEADING, Font.Size), If(ct.RunDirection = iTextSharp.text.pdf.PdfWriter.RUN_DIRECTION_LTR, iTextSharp.text.Element.ALIGN_RIGHT, iTextSharp.text.Element.ALIGN_RIGHT) Or iTextSharp.text.Element.ALIGN_BASELINE)
                        ct.AddText(New iTextSharp.text.Chunk(Text, Font))
                        ct.Go()
                        theText = theText.Substring(_Bounds(Count)(SubCount)(NextCount).nChar)
                    Next
                End If
            Next
            Writer.DirectContent.SetLineWidth(1)
            Writer.DirectContent.Rectangle(MaxRect.Left + Doc.LeftMargin - 2, Doc.PageSize.Height - Doc.TopMargin - MaxRect.Bottom + Font.BaseFont.GetFontDescriptor(iTextSharp.text.pdf.BaseFont.AWT_LEADING, Font.Size) * 2, MaxRect.Width - 2, MaxRect.Height - Font.BaseFont.GetFontDescriptor(iTextSharp.text.pdf.BaseFont.AWT_LEADING, Font.Size) * 2)
            Writer.DirectContent.Stroke()
        Next
    End Sub
    Public Shared Function GetFontPath(Index As Integer) As String
        'Return Utility.GetFilePath("files\" + "Scheherazade-R.ttf")
        Dim Fonts As String() = {"times.ttf", "me_quran.ttf", "Scheherazade.ttf", "PDMS_Saleem.ttf", "KFC_naskh.otf", "trado.ttf", "arabtype.ttf", "majalla.ttf", "msuighur.ttf", "ARIALUNI.ttf"}
        Return If(Index < 1 Or Index > 4, IO.Path.Combine(Environment.GetEnvironmentVariable("windir"), "Fonts\" + Fonts(Index)), Utility.GetFilePath("files\" + Fonts(Index)))
    End Function
    Public Shared Sub OutputPdf(Path As String, CurRenderArray As List(Of RenderArray.RenderItem))
        OutputPdf(New IO.FileStream(Path, IO.FileMode.Create, IO.FileAccess.Write, IO.FileShare.None), CurRenderArray)
    End Sub
    Public Shared Sub OutputPdf(Stream As IO.Stream, CurRenderArray As List(Of RenderArray.RenderItem))
        Dim Doc As New iTextSharp.text.Document
        Dim Writer As iTextSharp.text.pdf.PdfWriter = iTextSharp.text.pdf.PdfWriter.GetInstance(Doc, Stream)
        Doc.Open()
        Doc.NewPage()
        Dim BaseFont As iTextSharp.text.pdf.BaseFont = iTextSharp.text.pdf.BaseFont.CreateFont(GetFontPath(0), iTextSharp.text.pdf.BaseFont.IDENTITY_H, iTextSharp.text.pdf.BaseFont.NOT_EMBEDDED)
        Dim Font As New iTextSharp.text.Font(BaseFont, 20, iTextSharp.text.Font.BOLD)
        Dim _Bounds As New Generic.List(Of Generic.List(Of Generic.List(Of LayoutInfo)))
        Dim PrivateFontColl As New Drawing.Text.PrivateFontCollection
        PrivateFontColl.AddFontFile(GetFontPath(0))
        Dim DrawFont As New Drawing.Font(PrivateFontColl.Families(0), 20, FontStyle.Bold, GraphicsUnit.Point)
        'divide into pages by heights
        GetLayout(CurRenderArray, Doc.PageSize.Width - Doc.LeftMargin - Doc.RightMargin, _Bounds, GetTextWidthFromPdf(Font))
        Dim PageOffset As New PointF(0, 0)
        DoRenderPdf(Doc, Writer, Font, DrawFont, CurRenderArray, _Bounds, PageOffset, New PointF(0, 0))
        Writer.CloseStream = False
        Doc.Close()
    End Sub
    Delegate Function GetTextWidth(Str As String, MaxWidth As Single, IsRTL As Boolean, ByRef s As SizeF, ByRef Baseline As Single) As Integer
    Private Shared Function GetTextWidthPdf(Font As iTextSharp.text.Font, Str As String, MaxWidth As Single, IsRTL As Boolean, ByRef s As SizeF, ByRef Baseline As Single) As Integer
        Font.BaseFont.CorrectArabicAdvance()
        Dim Index As Integer = 0
        Do
            Index = Str.IndexOf(ArabicData.ArabicEndOfAyah, Index)
            If Index <> -1 Then
                Dim NumCount As Integer = 0
                Do While NumCount <> 3 And (Index + NumCount + 1) <= Str.Length - 1 AndAlso Char.IsDigit(Str(Index + NumCount + 1))
                    Str = Str.Remove(Index + NumCount + 1, 1).Insert(Index + NumCount + 1, ArabicData.ZeroWidthSpace)
                    NumCount += 1
                Loop
                Index = Index + 1
            End If
        Loop While Index <> -1
        'must count removed characters or convert to shaping characters
        'although should not significantly effect calculation only the height in a generally absorbed way
        'use leading plus one extra leading up and down for extra spacing
        Dim Text As String = Str
        Str = System.Text.RegularExpressions.Regex.Replace(Str, "(?<!\s+)(?:" + ArabicData.MakeRegMultiEx(Array.ConvertAll(ArabicData.RecitationCombiningSymbols, Function(Ch As Char) CStr(Ch))) + ")", String.Empty)
        s.Width = iTextSharp.text.pdf.ColumnText.GetWidth(New iTextSharp.text.Phrase(New iTextSharp.text.Chunk(Str, Font)), If(IsRTL, iTextSharp.text.pdf.PdfWriter.RUN_DIRECTION_RTL, iTextSharp.text.pdf.PdfWriter.RUN_DIRECTION_LTR), iTextSharp.text.pdf.ColumnText.AR_COMPOSEDTASHKEEL)
        Str = Text
        Dim Len As Integer = Str.Length
        Dim Search As Integer = Len
        'binary search the maximum characters
        If s.Width > MaxWidth Then
            While Search <> 1
                Search = Search \ 2
                If s.Width > MaxWidth Then
                    Len -= Search
                Else
                    Len += Search
                End If
                'cannot split arabic words due to shaping issues
                Str = Str.Substring(0, If(Str.IndexOf(" "c, Len - 1) = -1, Str.Length, Str.IndexOf(" "c, Len - 1) + 1))
                Str = System.Text.RegularExpressions.Regex.Replace(Str, "(?<!\s+)(?:" + ArabicData.MakeRegMultiEx(Array.ConvertAll(ArabicData.RecitationCombiningSymbols, Function(Ch As Char) CStr(Ch))) + ")", String.Empty)
                s.Width = iTextSharp.text.pdf.ColumnText.GetWidth(New iTextSharp.text.Phrase(Str, Font), If(IsRTL, iTextSharp.text.pdf.PdfWriter.RUN_DIRECTION_RTL, iTextSharp.text.pdf.PdfWriter.RUN_DIRECTION_LTR), iTextSharp.text.pdf.ColumnText.AR_COMPOSEDTASHKEEL)
                Dim SpecLen As Integer = System.Text.RegularExpressions.Regex.Match(Str, "\s+(?:" + ArabicData.MakeRegMultiEx(Array.ConvertAll(ArabicData.RecitationCombiningSymbols, Function(Ch As Char) CStr(Ch))) + ")\s+$").Length
                If SpecLen <> 0 Then
                    s.Width += iTextSharp.text.pdf.ColumnText.GetWidth(New iTextSharp.text.Phrase(Str.Substring(Str.Length - SpecLen), Font), If(IsRTL, iTextSharp.text.pdf.PdfWriter.RUN_DIRECTION_RTL, iTextSharp.text.pdf.PdfWriter.RUN_DIRECTION_LTR), iTextSharp.text.pdf.ColumnText.AR_COMPOSEDTASHKEEL)
                End If
                Str = Text
            End While
            Len = If(Str.IndexOf(" "c, Len - 1) = -1, Str.Length, Str.IndexOf(" "c, Len - 1) + 1)
            If s.Width > MaxWidth Then
                Len = Str.LastIndexOf(" "c, Len - 1 - 1) + 1 'factor towards fitting not overflowing
                Str = Str.Substring(0, Len)
                Str = System.Text.RegularExpressions.Regex.Replace(Str, "(?<!\s+)(?:" + ArabicData.MakeRegMultiEx(Array.ConvertAll(ArabicData.RecitationCombiningSymbols, Function(Ch As Char) CStr(Ch))) + ")", String.Empty)
                s.Width = iTextSharp.text.pdf.ColumnText.GetWidth(New iTextSharp.text.Phrase(Str, Font), If(IsRTL, iTextSharp.text.pdf.PdfWriter.RUN_DIRECTION_RTL, iTextSharp.text.pdf.PdfWriter.RUN_DIRECTION_LTR), iTextSharp.text.pdf.ColumnText.AR_COMPOSEDTASHKEEL)
                Dim SpecLen As Integer = System.Text.RegularExpressions.Regex.Match(Str, "\s+(?:" + ArabicData.MakeRegMultiEx(Array.ConvertAll(ArabicData.RecitationCombiningSymbols, Function(Ch As Char) CStr(Ch))) + ")\s+$").Length
                If SpecLen <> 0 Then
                    s.Width += iTextSharp.text.pdf.ColumnText.GetWidth(New iTextSharp.text.Phrase(Str.Substring(Str.Length - SpecLen), Font), If(IsRTL, iTextSharp.text.pdf.PdfWriter.RUN_DIRECTION_RTL, iTextSharp.text.pdf.PdfWriter.RUN_DIRECTION_LTR), iTextSharp.text.pdf.ColumnText.AR_COMPOSEDTASHKEEL)
                End If
                Str = Text
            End If
        End If
        If Len <> Text.Length AndAlso Array.IndexOf(ArabicData.RecitationCombiningSymbols, Text(Len)) <> -1 Then Len -= 1
        Text = Text.Substring(0, Len)
        Index = 1
        Dim MaxAscent As Integer = 0
        Dim MinAscent As Integer = 0
        Do
            Index = Text.IndexOfAny(ArabicData.RecitationCombiningSymbols, Index)
            If Index <> -1 Then
                Dim ChBounds As Integer()
                If (Text(Index - 1) = " "c) Then
                    'stopping symbols handled by normal rendering engine
                    Index += 1
                    Continue Do
                ElseIf Text(Index - 1) = ArabicData.ArabicTatweel Then
                    ChBounds = Font.BaseFont.GetCharBBox(AscW(Text(Index - 1)))
                Else
                    Dim ShapeCh As Char = ArabicData.Data.ArabicLetters(ArabicData.FindLetterBySymbol(Text(Index - 1))).Shaping(ArabicData.GetShapeIndexFromString(Text, Index - 1, 1))
                    ChBounds = Font.BaseFont.GetCharBBox(AscW(ShapeCh))
                End If
                Dim Offset As Integer = 0
                Do
                    Dim DiaBounds As Integer() = Font.BaseFont.GetCharBBox(AscW(Text(Index)))
                    Offset = If(DiaBounds(1) < 0 And DiaBounds(3) > If(Offset < 0, Offset, ChBounds(1)), -(DiaBounds(3) - If(Offset < 0, Offset, ChBounds(1))), If(DiaBounds(1) >= 0 And DiaBounds(1) < If(Offset > 0, Offset, ChBounds(3)), -(DiaBounds(1) - If(Offset > 0, Offset, ChBounds(3))), 0))
                    If DiaBounds(1) < 0 Then Offset = DiaBounds(1) - Offset
                    If DiaBounds(1) >= 0 Then Offset = DiaBounds(3) + Offset
                    If Offset > 0 Then MaxAscent = Math.Max(Offset, MaxAscent)
                    If Offset < 0 Then MinAscent = Math.Min(Offset, MinAscent)
                    Index += 1
                Loop While Index <> Text.Length AndAlso Array.IndexOf(ArabicData.RecitationCombiningSymbols, Text(Index)) <> -1
            End If
        Loop While Index <> -1
        Baseline = Math.Max(MaxAscent * 0.001F * Font.Size, Font.BaseFont.GetAscentPoint(Text, Font.Size)) + Font.BaseFont.GetFontDescriptor(iTextSharp.text.pdf.BaseFont.AWT_LEADING, Font.Size) * 2
        s.Height = Font.BaseFont.GetFontDescriptor(iTextSharp.text.pdf.BaseFont.AWT_LEADING, Font.Size) * 4 + Baseline - Math.Min(MinAscent * 0.001F * Font.Size, Font.BaseFont.GetDescentPoint(Text, Font.Size))
        Return Len
    End Function
    Private Shared Function GetTextWidthFromPdf(Font As iTextSharp.text.Font) As GetTextWidth
        Return Function(Str As String, MaxWidth As Single, IsRTL As Boolean, ByRef s As SizeF, ByRef Baseline As Single)
                   Dim Ret As Integer = GetTextWidthPdf(Font, Str, MaxWidth - 4, IsRTL, s, Baseline)
                   s.Width += 4 '1 unit for line and 1 for spacing on each side
                   Return Ret
               End Function
    End Function
    Structure OverInfo
        Sub New(NewIndex As Integer, NewSubIndex As Integer, NewMaxRight As Single)
            Index = NewIndex
            SubIndex = NewSubIndex
            MaxRight = NewMaxRight
        End Sub
        Dim Index As Integer
        Dim SubIndex As Integer
        Dim MaxRight As Single
    End Structure
    Public Shared Function GetLayout(CurRenderArray As List(Of RenderArray.RenderItem), _Width As Single, ByRef Bounds As Generic.List(Of Generic.List(Of Generic.List(Of LayoutInfo))), WidthFunc As GetTextWidth) As SizeF
        Dim MaxRight As Single = _Width
        Dim Top As Single = 0
        Dim NextRight As Single = _Width
        Dim LastCurTop As Single = 0
        Dim LastRight As Single = _Width
        Dim OverIndexes As New List(Of OverInfo)
        For Count As Integer = 0 To CurRenderArray.Count - 1
            Dim IsOverflow As Boolean = False
            Dim MaxWidth As Single = 0
            Dim Right As Single = NextRight
            Dim CurTop As Single = 0
            Dim MaxTop As Single = 0
            Bounds.Add(New Generic.List(Of Generic.List(Of LayoutInfo)))
            For SubCount As Integer = 0 To CurRenderArray(Count).TextItems.Length - 1
                Bounds(Count).Add(New Generic.List(Of LayoutInfo))
                Dim s As Drawing.SizeF
                If CurRenderArray(Count).TextItems(SubCount).DisplayClass = RenderArray.RenderDisplayClass.eNested Then
                    Dim SubBounds As New Generic.List(Of Generic.List(Of Generic.List(Of LayoutInfo)))
                    s = GetLayout(CType(CurRenderArray(Count).TextItems(SubCount).Text, List(Of RenderArray.RenderItem)), _Width, SubBounds, WidthFunc)
                    If s.Width > NextRight Then
                        OverIndexes.Add(New OverInfo(Count, SubCount, NextRight))
                        NextRight = _Width
                        IsOverflow = True
                    End If
                    Right = NextRight
                    Bounds(Count)(SubCount).Add(New LayoutInfo(New RectangleF(Right, Top + CurTop, s.Width, s.Height), 0, 0, SubBounds))
                    MaxWidth = Math.Max(MaxWidth, s.Width)
                ElseIf CurRenderArray(Count).TextItems(SubCount).DisplayClass = RenderArray.RenderDisplayClass.eArabic Or CurRenderArray(Count).TextItems(SubCount).DisplayClass = RenderArray.RenderDisplayClass.eLTR Or CurRenderArray(Count).TextItems(SubCount).DisplayClass = RenderArray.RenderDisplayClass.eRTL Or CurRenderArray(Count).TextItems(SubCount).DisplayClass = RenderArray.RenderDisplayClass.eTransliteration Then
                    Dim theText As String = CStr(CurRenderArray(Count).TextItems(SubCount).Text)
                    While theText <> String.Empty
                        Dim nChar As Integer
                        Dim Baseline As Single
                        nChar = WidthFunc(theText, _Width, CurRenderArray(Count).TextItems(SubCount).DisplayClass = RenderArray.RenderDisplayClass.eArabic Or CurRenderArray(Count).TextItems(SubCount).DisplayClass = RenderArray.RenderDisplayClass.eRTL, s, Baseline)
                        'break up string on previous word boundary unless beginning of string
                        'arabic strings cannot be broken up in the middle due to letters joining which would throw off calculations
                        If nChar = 0 Then
                            nChar = theText.Length 'If no room for even a letter then just use placeholder
                        ElseIf nChar <> theText.Length Then
                            Dim idx As Integer = Array.FindLastIndex(theText.ToCharArray(), nChar - 1, nChar, Function(ch As Char) Char.IsWhiteSpace(ch))
                            If idx <> -1 Then nChar = idx + 1
                        End If
                        If theText.Substring(nChar) <> String.Empty Then
                            WidthFunc(theText.Substring(0, nChar), _Width, CurRenderArray(Count).TextItems(SubCount).DisplayClass = RenderArray.RenderDisplayClass.eArabic Or CurRenderArray(Count).TextItems(SubCount).DisplayClass = RenderArray.RenderDisplayClass.eRTL, s, Baseline)
                        End If
                        theText = theText.Substring(nChar)
                        If theText <> String.Empty Or s.Width > NextRight Then
                            If s.Width > NextRight Then OverIndexes.Add(New OverInfo(Count, SubCount, NextRight))
                            NextRight = _Width
                            IsOverflow = True
                        End If
                        If s.Width > NextRight Then Right = NextRight
                        Bounds(Count)(SubCount).Add(New LayoutInfo(New RectangleF(Right, Top + CurTop, s.Width, s.Height), Baseline, nChar, Nothing))
                        MaxTop = Math.Max(CurTop + s.Height, MaxTop)
                        If theText <> String.Empty Then
                            CurTop += s.Height
                        End If
                        MaxWidth = Math.Max(MaxWidth, s.Width)
                    End While
                End If
                MaxTop = Math.Max(CurTop + s.Height, MaxTop)
                If Bounds(Count)(SubCount).Count <> 0 Then
                    CurTop += s.Height
                End If
            Next
            'centering must come after maximum width is calculated
            For SubCount = 0 To Bounds(Count).Count - 1
                For NextCount = 0 To Bounds(Count)(SubCount).Count - 1
                    MaxRight = Math.Min(If(IsOverflow Or NextCount <> Bounds(Count)(SubCount).Count - 1, _Width, Bounds(Count)(SubCount)(NextCount).Rect.Left) - ((MaxWidth + Bounds(Count)(SubCount)(NextCount).Rect.Width) / 2), MaxRight)
                    Bounds(Count)(SubCount)(NextCount) = New LayoutInfo(New RectangleF(If(IsOverflow Or NextCount <> Bounds(Count)(SubCount).Count - 1, _Width, Bounds(Count)(SubCount)(NextCount).Rect.Left) - ((MaxWidth + Bounds(Count)(SubCount)(NextCount).Rect.Width) / 2), Bounds(Count)(SubCount)(NextCount).Rect.Top + If(IsOverflow, LastCurTop, 0), Bounds(Count)(SubCount)(NextCount).Rect.Width, Bounds(Count)(SubCount)(NextCount).Rect.Height), Bounds(Count)(SubCount)(NextCount).Baseline, Bounds(Count)(SubCount)(NextCount).nChar, Bounds(Count)(SubCount)(NextCount).Bounds)
                Next
            Next
            If Count <> 0 AndAlso ((CurRenderArray(Count).Type = RenderTypes.eHeaderLeft Or CurRenderArray(Count - 1).Type = RenderTypes.eHeaderRight) Or (CurRenderArray(Count).Type = RenderTypes.eHeaderCenter And CurRenderArray(Count - 1).Type <> RenderTypes.eHeaderLeft) Or (CurRenderArray(Count).Type <> RenderTypes.eHeaderRight And CurRenderArray(Count - 1).Type = RenderTypes.eHeaderCenter)) Then
                Top += MaxTop + LastCurTop
                CurTop = 0
                MaxTop = 0
                LastCurTop = 0
                OverIndexes.Add(New OverInfo(Count + 1, 0, NextRight - MaxWidth))
                NextRight = _Width
                Right = NextRight
            ElseIf IsOverflow Then
                Top += LastCurTop
                LastCurTop = 0
                NextRight -= MaxWidth
                Right = NextRight
            Else
                NextRight -= MaxWidth
            End If
            LastCurTop = Math.Max(MaxTop, LastCurTop)
            LastRight = NextRight
            If Count = CurRenderArray.Count - 1 Then
                Top += CurTop + Bounds(Count)(CurRenderArray(Count).TextItems.Length - 1)(Bounds(Count)(CurRenderArray(Count).TextItems.Length - 1).Count - 1).Rect.Height
                OverIndexes.Add(New OverInfo(Count + 1, 0, NextRight))
            End If
        Next
        Dim NextOverIndex As Integer = 0
        For Count = 0 To Bounds.Count - 1
            For SubCount = 0 To Bounds(Count).Count - 1
                Dim CenterAdj As Single = 0
                If NextOverIndex <> OverIndexes.Count AndAlso (OverIndexes(NextOverIndex).Index < Count Or _
                        OverIndexes(NextOverIndex).Index = Count) Then
                    NextOverIndex += 1
                End If
                If NextOverIndex <> OverIndexes.Count Then
                    CenterAdj = (MaxRight - OverIndexes(NextOverIndex).MaxRight) / 2
                End If
                For NextCount = 0 To Bounds(Count)(SubCount).Count - 1
                    'overall centering can be done here though must calculate an overall line width
                    Bounds(Count)(SubCount)(NextCount) = New LayoutInfo(New RectangleF(Bounds(Count)(SubCount)(NextCount).Rect.Left - MaxRight + CenterAdj, Bounds(Count)(SubCount)(NextCount).Rect.Top, Bounds(Count)(SubCount)(NextCount).Rect.Width, Bounds(Count)(SubCount)(NextCount).Rect.Height), Bounds(Count)(SubCount)(NextCount).Baseline, Bounds(Count)(SubCount)(NextCount).nChar, Bounds(Count)(SubCount)(NextCount).Bounds)
                Next
            Next
        Next
        Return New SizeF(_Width - MaxRight, Top)
    End Function
    Public Sub Render(ByVal writer As System.Web.UI.HtmlTextWriter, ByVal TabCount As Integer)
        DoRender(writer, TabCount, Items, String.Empty)
    End Sub
    Public Shared Function GetQuoteModeJS() As String()
        Return New String() {"javascript: quoteMode();", String.Empty, Utility.GetLookupStyleSheetJS(), "function quoteMode() { var rule = findStyleSheetRule('span.copy'); rule.style.display = $('#quotemode').prop('checked') === true ? 'block' : 'none'; }"}
    End Function
    Public Shared Function GetStarRatingJS() As String
        Return "function changeStarRating(e, item, val, data) { $(item).parent().find('span').each(function (index, Element) { if (Element.textContent !== '\u26D2') { Element.style.color = (index < val) ? '#00a4e4' : '#cccccc'; Element.innerText = (index < val) ? '\u2605' : '\u2606'; } }); data['Rating'] = val.toString(); $.ajax({url: '" + Utility.GetPageString("HadithRanking") + "', data: data, type: 'POST', success: function(data) { $(item).parent().parent().children('span').text(data); }, dataType: 'text'}); } " + _
            "function restoreStarRating(e, item) { $(item).parent().find('span').each(function (index, Element) { if (Element.textContent !== '\u26D2') Element.style.color = (Element.textContent === '\u2605') ? '#00a4e4' : '#cccccc'; }); } " + _
            "function updateStarRating(e, item, val) { $(item).parent().find('span').each(function (index, Element) { if (Element.textContent !== '\u26D2') Element.style.color = (index < val) ? '#aa1010' : ((Element.textContent === '\u2605') ? '#00a4e4' : '#cccccc'); }); }"
        'Return "function changeStarRating(e, item, data) { $(item).find('div').get(0).style.width = (Math.ceil((e.pageX - $(item).parent().offset().left) / $(item).outerWidth() * 10) * 10).toString() + '%'; data['Rating'] = Math.ceil((e.pageX - $(item).parent().offset().left) / $(item).outerWidth() * 10).toString(); $.ajax({url: '" + host.GetPageString("HadithRanking") + "', data: data, type: 'POST', success: function(data) { $(item).parent().parent().find('span').text(data); }, dataType: 'text'}); } " + _
        '    "function restoreStarRating(e, item) { $(item).find('div').get(1).style.width = '0%'; $(item).find('div').get(1).style.zIndex = 102; } " + _
        '    "function updateStarRating(e, item) { $(item).find('div').get(1).style.width = (Math.ceil((e.pageX - $(item).parent().offset().left) / $(item).outerWidth() * 10) * 10).toString() + '%'; $(item).find('div').get(0).style.zIndex = parseFloat($(item).find('div').get(1).style.width) > parseFloat($(item).find('div').get(0).style.width) ? 103 : 102; }"
    End Function
    Public Shared Function GetContinueStopJS() As String
        Return "function changeContinueStop(e, item, data) {}"
    End Function
    Public Shared Function GetCopyClipboardJS() As String
        Return "function setClipboardText(text) { if (window.clipboardData) { window.clipboardData.setData('Text', text); } }"
    End Function
    Public Shared Function GetSetClipboardJS() As String
        Return "function getText(top, child) { var iCount, item = child === '' ? renderList[top] : renderList[top].children[child], text, chtxt, str; text = item['title'] === '' ? '' : getText(item['title'], '') + '\t'; for (iCount = 0; iCount < item['arabic'].length; iCount++) { if (item['arabic'][iCount] !== '') text += $('#' + item['arabic'][iCount]).text() + '\t'; } for (iCount = 0; iCount < item['translit'].length; iCount++) { if (item['translit'][iCount] !== '') text += $('#' + item['translit'][iCount]).text() + '\t'; } for (iCount = 0; iCount < item['translate'].length; iCount++) { if (item['translate'][iCount] !== '') text += $('#' + item['translate'][iCount]).text() + '\t'; } chtxt = ''; for (k in item['children']) { if (chtxt !== '') chtxt += ' '; for (iCount = 0; iCount < item['children'][k]['arabic'].length; iCount++) { if (item['children'][k]['arabic'][iCount] !== '') chtxt += $('#' + item['children'][k]['arabic'][iCount]).text(); } str = ''; for (iCount = 0; iCount < item['children'][k]['translit'].length; iCount++) { if (item['children'][k]['translit'][iCount] !== '') str += $('#' + item['children'][k]['translit'][iCount]).text(); } chtxt += (str !== '' ? '(' + str + ')' : '') + '='; for (iCount = 0; iCount < item['children'][k]['translate'].length; iCount++) { if (item['children'][k]['translate'][iCount] !== '') chtxt += $('#' + item['children'][k]['translate'][iCount]).text(); }; } if (chtxt !== '') text += '[' + chtxt + ']'; return text; }"
    End Function
    Public Function GetRenderJS() As String()
        Return DoGetRenderJS(Items)
    End Function
    Public Shared Function GetInitJS(Items As Collections.Generic.List(Of RenderItem)) As String
        Dim Objects As Object() = GetInitJSItems(Items, String.Empty, String.Empty)
        Dim ListJS As String()() = CType(Objects(2), List(Of String())).ToArray()
        Dim ListJSInit As New List(Of String)
        Dim ListJSAfter As New List(Of String)
        For Count = 0 To ListJS.Length - 1
            For SubCount = 0 To ListJS(Count).Length - 1
                If SubCount = 1 Then
                    ListJSInit.Add(ListJS(Count)(SubCount))
                ElseIf SubCount >= 2 Then
                    ListJSAfter.Add(ListJS(Count)(SubCount))
                End If
            Next
        Next
        Return "var renderList = " + Utility.MakeJSIndexedObject(CType(CType(Objects(0), ArrayList).ToArray(GetType(String)), String()), New Array() {CType(CType(Objects(1), ArrayList).ToArray(GetType(String)), String())}, True) + ";" + String.Join(String.Empty, ListJSInit.ToArray()) + String.Join(String.Empty, ListJSAfter.ToArray())
    End Function
    Public Shared Function GetInitJSItems(Items As Collections.Generic.List(Of RenderItem), Title As String, NestPrefix As String) As Object()
        Dim Count As Integer
        Dim Index As Integer
        Dim Objects As ArrayList = New ArrayList From {New ArrayList, New ArrayList, New List(Of String())}
        Dim Names As New ArrayList
        Dim LastTitle As String = Title
        For Count = 0 To Items.Count - 1
            Dim Arabic As New ArrayList
            Dim Translit As New ArrayList
            Dim Translate As New ArrayList
            Dim Children As ArrayList = New ArrayList From {New ArrayList, New ArrayList}
            For Index = 0 To Items(Count).TextItems.Length - 1
                If Items(Count).TextItems(Index).DisplayClass = RenderDisplayClass.eNested Then
                    Dim Objs As Object() = GetInitJSItems(CType(Items(Count).TextItems(Index).Text, Collections.Generic.List(Of RenderItem)), LastTitle, CStr(Count))
                    CType(Children(0), ArrayList).AddRange(CType(Objs(0), ArrayList))
                    CType(Children(1), ArrayList).AddRange(CType(Objs(1), ArrayList))
                ElseIf Items(Count).TextItems(Index).DisplayClass = RenderDisplayClass.eList Then
                    CType(Objects(2), List(Of String())).AddRange(GetTableJSFunctions(CType(Items(Count).TextItems(Index).Text, Object())))
                ElseIf Items(Count).TextItems(Index).DisplayClass = RenderDisplayClass.eRanking Or _
                    Items(Count).TextItems(Index).DisplayClass = RenderDisplayClass.eContinueStop Then
                Else
                    If Items(Count).TextItems(Index).DisplayClass = RenderDisplayClass.eArabic Then
                        Arabic.Add("arabic" + CStr(IIf(NestPrefix = String.Empty, String.Empty, NestPrefix + "_")) + CStr(Count) + "_" + CStr(Index))
                    ElseIf Items(Count).TextItems(Index).DisplayClass = RenderDisplayClass.eTransliteration Then
                        Translit.Add("translit" + CStr(IIf(NestPrefix = String.Empty, String.Empty, NestPrefix + "_")) + CStr(Count) + "_" + CStr(Index))
                    Else
                        Translate.Add("translate" + CStr(IIf(NestPrefix = String.Empty, String.Empty, NestPrefix + "_")) + CStr(Count) + "_" + CStr(Index))
                    End If
                End If
            Next
            If Items(Count).Type = RenderTypes.eHeaderCenter Then
                LastTitle = "ri" + CStr(Count)
            End If
            If Arabic.Count <> 0 Or Translit.Count <> 0 Or Translate.Count <> 0 Then
                If Arabic.Count = 0 Then Arabic.Add(String.Empty)
                If Translit.Count = 0 Then Translit.Add(String.Empty)
                If Translate.Count = 0 Then Translate.Add(String.Empty)
                CType(Objects(0), ArrayList).Add("ri" + CStr(IIf(NestPrefix = String.Empty, String.Empty, NestPrefix + "_")) + CStr(Count))
                CType(Objects(1), ArrayList).Add(Utility.MakeJSIndexedObject(New String() {"title", "arabic", "translit", "translate", "children"}, New Array() {New String() {"'" + CStr(IIf(Items(Count).Type <> RenderTypes.eHeaderLeft And Items(Count).Type <> RenderTypes.eHeaderCenter And Items(Count).Type <> RenderTypes.eHeaderRight, Utility.EncodeJS(LastTitle), String.Empty)) + "'", Utility.MakeJSArray(CType(Arabic.ToArray(GetType(String)), String())), Utility.MakeJSArray(CType(Translit.ToArray(GetType(String)), String())), Utility.MakeJSArray(CType(Translate.ToArray(GetType(String)), String())), Utility.MakeJSIndexedObject(CType(CType(Children(0), ArrayList).ToArray(GetType(String)), String()), New Array() {CType(CType(Children(1), ArrayList).ToArray(GetType(String)), String())}, True)}}, True))
            End If
        Next
        Return CType(Objects.ToArray(GetType(Object)), Object())
    End Function
    Public Shared Function DoGetRenderJS(Items As Collections.Generic.List(Of RenderItem)) As String()
        Return New String() {String.Empty, String.Empty, GetInitJS(Items), GetCopyClipboardJS(), GetSetClipboardJS(), GetStarRatingJS(), GetContinueStopJS()}
    End Function
    Public Shared Function GetTableJSFunctions(ByVal Output As Object()) As String()()
        Dim Count As Integer
        Dim JSFuncs As New List(Of String())
        Dim OutArray As Object() = Output
        If Output.Length = 0 Then Return JSFuncs.ToArray()
        JSFuncs.Add(CType(OutArray(0), String()))
        For Count = 2 To OutArray.Length - 1
            If TypeOf OutArray(Count) Is Object() Then
                Dim InnerArray As Object() = DirectCast(OutArray(Count), Object())
                For Index = 0 To InnerArray.Length - 1
                    If TypeOf InnerArray(Index) Is Object() Then
                        JSFuncs.AddRange(GetTableJSFunctions(DirectCast(InnerArray(Index), Object())))
                    End If
                Next
            End If
        Next
        Return JSFuncs.ToArray()
    End Function
    Public Shared Sub WriteTable(ByVal writer As System.Web.UI.HtmlTextWriter, ByVal Output As Object(), ByVal TabCount As Integer, Prefix As String)
        '2 dimensional array for table
        Dim BaseTabs As String = Utility.MakeTabString(TabCount)
        Dim Count As Integer
        Dim Index As Integer
        If Output.Length = 0 Then Return
        Dim OutArray As Object() = Output
        writer.Write(vbCrLf + BaseTabs)
        writer.WriteBeginTag("table")
        writer.WriteAttribute("id", "render" + Prefix)
        writer.Write(HtmlTextWriter.TagRightChar)
        For Count = 2 To OutArray.Length - 1
            writer.Write(vbCrLf + BaseTabs + vbTab)
            writer.WriteFullBeginTag("tr")
            If TypeOf OutArray(Count) Is Object() Then
                Dim InnerArray As Object() = DirectCast(OutArray(Count), Object())
                For Index = 0 To InnerArray.Length - 1
                    writer.Write(vbCrLf + BaseTabs + vbTab + vbTab)
                    writer.WriteFullBeginTag(CStr(IIf(Count = 2, "th", "td")))
                    writer.Write(vbCrLf + BaseTabs + vbTab + vbTab + vbTab)
                    writer.WriteBeginTag("span")
                    If Count <> 2 Then writer.WriteAttribute("id", "render" + CStr(IIf(Prefix <> String.Empty, Prefix + "_", String.Empty)) + CStr(Count - 3) + "_" + CStr(Index))
                    If (CStr(DirectCast(OutArray(1), Object())(Index)) <> String.Empty) Then
                        writer.WriteAttribute("class", CStr(DirectCast(OutArray(1), Object())(Index)))
                        If CStr(DirectCast(OutArray(1), Object())(Index)) = "transliteration" Then
                            writer.WriteAttribute("style", "display: " + CStr(IIf(CInt(HttpContext.Current.Request.QueryString.Get("translitscheme")) <> 0, "block", "none")) + ";")
                        ElseIf CStr(DirectCast(OutArray(1), Object())(Index)) = "check" Then
                            writer.Write(HtmlTextWriter.TagRightChar)
                            writer.WriteBeginTag("input")
                            writer.WriteAttribute("id", "check" + CStr(IIf(Prefix <> String.Empty, Prefix + "_", String.Empty)) + CStr(Count - 3) + "_" + CStr(Index))
                            writer.WriteAttribute("type", "checkbox")
                            writer.WriteAttribute("onchange", CType(OutArray(0), String())(0))
                        ElseIf CStr(DirectCast(OutArray(1), Object())(Index)) = "hidden" Then
                            writer.WriteAttribute("style", "display: none;")
                        End If
                    End If
                    writer.Write(HtmlTextWriter.TagRightChar)
                    If TypeOf InnerArray(Index) Is Object() Then
                        WriteTable(writer, DirectCast(InnerArray(Index), Object()), TabCount + 4, CStr(Count - 3) + "_" + CStr(Index))
                    Else
                        writer.Write(Utility.HtmlTextEncode(CStr(InnerArray(Index))).Replace(vbCrLf, "<br>"))
                    End If
                    writer.WriteEndTag("span")
                    writer.Write(vbCrLf + BaseTabs + vbTab + vbTab)
                    writer.WriteEndTag(CStr(IIf(Count = 2, "th", "td")))
                Next
            End If
            writer.Write(vbCrLf + BaseTabs + vbTab)
            writer.WriteEndTag("tr")
        Next
        writer.Write(vbCrLf + BaseTabs)
        writer.WriteEndTag("table")
    End Sub
    Public Shared Sub DoRender(ByVal writer As System.Web.UI.HtmlTextWriter, ByVal TabCount As Integer, Items As Collections.Generic.List(Of RenderItem), NestPrefix As String)
        Dim BaseTabs As String = Utility.MakeTabString(TabCount)
        Dim Count As Integer
        Dim Index As Integer
        Dim Base As Integer = 0
        For Count = 0 To Items.Count - 1
            If Count <> 0 AndAlso ((Items(Count).Type = RenderTypes.eHeaderLeft Or Items(Count - 1).Type = RenderTypes.eHeaderRight) Or (Items(Count).Type = RenderTypes.eHeaderCenter And Items(Count - 1).Type <> RenderTypes.eHeaderLeft) Or (Items(Count).Type <> RenderTypes.eHeaderRight And Items(Count - 1).Type = RenderTypes.eHeaderCenter)) Then
                writer.Write(vbCrLf + BaseTabs)
                writer.WriteFullBeginTag("br")
            End If
            If Count <> 0 AndAlso (Items(Count).Type = RenderTypes.eText And Items(Count - 1).Type <> RenderTypes.eText) Then Base = Count
            'no spacing since inline-block element
            writer.WriteBeginTag("div")
            writer.WriteAttribute("class", "multidisplay")
            writer.Write(HtmlTextWriter.TagRightChar)
            For Index = 0 To Items(Count).TextItems.Length - 1
                writer.Write(vbCrLf + BaseTabs + vbTab)
                writer.WriteBeginTag(CStr(IIf(Items(Count).TextItems(Index).DisplayClass = RenderDisplayClass.eNested Or Items(Count).TextItems(Index).DisplayClass = RenderDisplayClass.eRanking Or Items(Count).TextItems(Index).DisplayClass = RenderDisplayClass.eList, "div", "span")))
                If Items(Count).TextItems(Index).DisplayClass <> RenderDisplayClass.eList And (Items(Count).Type = RenderTypes.eHeaderCenter Or (Items(Count).Type = RenderTypes.eText And (Count - Base) Mod 2 = 1)) Then writer.WriteAttribute("style", "background-color: 0xD0D0D0;")
                If Items(Count).TextItems(Index).DisplayClass = RenderDisplayClass.eNested Then
                    writer.Write(HtmlTextWriter.TagRightChar)
                    DoRender(writer, TabCount, CType(Items(Count).TextItems(Index).Text, Collections.Generic.List(Of RenderItem)), CStr(Count))
                ElseIf Items(Count).TextItems(Index).DisplayClass = RenderDisplayClass.eList Then
                    writer.WriteAttribute("style", "direction: ltr;")
                    writer.Write(HtmlTextWriter.TagRightChar)
                    WriteTable(writer, CType(Items(Count).TextItems(Index).Text, Object()), TabCount, CStr(Count))
                ElseIf Items(Count).TextItems(Index).DisplayClass = RenderDisplayClass.eContinueStop Then
                    'U+2BC3 is horizontal stop sign make red color, U+2B45/6 is left/rightwards quadruple arrow make green color
                    writer.WriteAttribute("style", "color: " + If(CStr(Items(Count).TextItems(Index).Text) <> String.Empty, "#ff0000", "#00ff00") + ";")
                    writer.WriteAttribute("onclick", "javascript: changeContinueStop(event, this, {Reference:'" + String.Empty + "'});")
                    writer.Write(HtmlTextWriter.TagRightChar + If(CStr(Items(Count).TextItems(Index).Text) <> String.Empty, "&#2BC3;", "&#x2B45;"))
                ElseIf Items(Count).TextItems(Index).DisplayClass = RenderDisplayClass.eRanking Then
                    Dim Data As String() = CStr(Items(Count).TextItems(Index).Text).Split("|"c)
                    writer.Write(HtmlTextWriter.TagRightChar)
                    writer.WriteBeginTag("div")

                    'writer.WriteAttribute("class", "classification")
                    'writer.Write(HtmlTextWriter.TagRightChar)
                    'writer.WriteBeginTag("div")
                    'writer.WriteAttribute("class", "cover")
                    'writer.WriteAttribute("onclick", "javascript: changeStarRating(event, this, {Collection:'" + Data(0) + "', Book:'" + Data(1) + "', Hadith:'" + Data(2) + "'});")
                    'writer.WriteAttribute("onmousemove", "javascript: updateStarRating(event, this);")
                    'writer.WriteAttribute("onmouseover", "javascript: updateStarRating(event, this);")
                    'writer.WriteAttribute("onmouseout", "javascript: restoreStarRating(event, this);")
                    'writer.Write(HtmlTextWriter.TagRightChar)
                    'writer.WriteBeginTag("div")
                    'writer.WriteAttribute("class", "progress")
                    'writer.WriteAttribute("style", "width: " + CStr(IIf(CInt(Data(5)) = -1, 0, CInt(Data(5)) * 10)) + "%;")
                    'writer.Write(HtmlTextWriter.TagRightChar)
                    'writer.WriteEndTag("div")
                    'writer.WriteBeginTag("div")
                    'writer.WriteAttribute("class", "change")
                    'writer.WriteAttribute("style", "width: 0%;")
                    'writer.Write(HtmlTextWriter.TagRightChar)
                    'writer.WriteEndTag("div")
                    'writer.WriteEndTag("div")

                    writer.WriteAttribute("style", "padding: 0; margin: 0;")
                    writer.Write(HtmlTextWriter.TagRightChar)
                    For StarCount As Integer = 1 To 10
                        writer.WriteBeginTag("span")
                        writer.WriteAttribute("class", CStr(IIf(StarCount Mod 2 = 1, "classification", "classificationalt")))
                        writer.WriteAttribute("style", "color: " + CStr(IIf(CInt(IIf(CInt(Data(5)) = -1, 0, Data(5))) < StarCount, "#cccccc", "#00a4e4")) + ";")
                        writer.WriteAttribute("onclick", "javascript: changeStarRating(event, this, " + CStr(StarCount) + ", {Collection:'" + Data(0) + "', Book:'" + Data(1) + "', Hadith:'" + Data(2) + "'});")
                        writer.WriteAttribute("onmousemove", "javascript: updateStarRating(event, this, " + CStr(StarCount) + ");")
                        writer.WriteAttribute("onmouseover", "javascript: updateStarRating(event, this, " + CStr(StarCount) + ");")
                        writer.WriteAttribute("onmouseout", "javascript: restoreStarRating(event, this);")
                        writer.Write(HtmlTextWriter.TagRightChar + CStr(IIf(CInt(IIf(CInt(Data(5)) = -1, 0, Data(5))) < StarCount, "&#x2606;", "&#x2605;")))
                        writer.WriteEndTag("span")
                    Next
                    writer.WriteBeginTag("span")
                    writer.WriteAttribute("style", "color:#ff0000;display:inline-block;height:1em;font-size:1em;text-align:center;line-height:1;overflow:hidden;cursor:pointer;margin-right:0;margin-left:1em;")
                    writer.WriteAttribute("onclick", "javascript: changeStarRating(event, this, 0, {Collection:'" + Data(0) + "', Book:'" + Data(1) + "', Hadith:'" + Data(2) + "'});")
                    writer.WriteAttribute("onmousemove", "javascript: updateStarRating(event, this, 0);")
                    writer.WriteAttribute("onmouseover", "javascript: updateStarRating(event, this, 0);")
                    writer.WriteAttribute("onmouseout", "javascript: restoreStarRating(event, this);")
                    writer.Write(HtmlTextWriter.TagRightChar + "&#x26D2;")
                    writer.WriteEndTag("span")
                    writer.WriteEndTag("div")
                    writer.WriteBeginTag("span")
                    writer.Write(HtmlTextWriter.TagRightChar)
                    If (CInt(Data(4)) <> 0) Then writer.Write("Average of " + CStr(CInt(Data(3)) / CInt(Data(4)) / 2) + " out of " + Data(4) + " rankings")
                    writer.WriteEndTag("span")
                Else
                    If Items(Count).TextItems(Index).DisplayClass = RenderDisplayClass.eArabic Then
                        writer.WriteAttribute("class", "arabic")
                        writer.WriteAttribute("dir", "rtl")
                        writer.WriteAttribute("id", "arabic" + CStr(IIf(NestPrefix = String.Empty, String.Empty, NestPrefix + "_")) + CStr(Count) + "_" + CStr(Index))
                        writer.WriteAttribute("style", "color: " + System.Drawing.ColorTranslator.ToHtml(Items(Count).TextItems(Index).Clr) + ";")
                    ElseIf Items(Count).TextItems(Index).DisplayClass = RenderDisplayClass.eTransliteration Then
                        writer.WriteAttribute("class", "transliteration")
                        writer.WriteAttribute("dir", "ltr")
                        writer.WriteAttribute("style", "color: " + System.Drawing.ColorTranslator.ToHtml(Items(Count).TextItems(Index).Clr) + "; display: " + CStr(IIf(CInt(HttpContext.Current.Request.QueryString.Get("translitscheme")) <> ArabicData.TranslitScheme.None, "block", "none")) + ";")
                        writer.WriteAttribute("id", "translit" + CStr(IIf(NestPrefix = String.Empty, String.Empty, NestPrefix + "_")) + CStr(Count) + "_" + CStr(Index))
                    Else
                        writer.WriteAttribute("class", "translation")
                        writer.WriteAttribute("dir", CStr(IIf(Items(Count).TextItems(Index).DisplayClass = RenderDisplayClass.eRTL, "rtl", "ltr")))
                        writer.WriteAttribute("id", "translate" + CStr(IIf(NestPrefix = String.Empty, String.Empty, NestPrefix + "_")) + CStr(Count) + "_" + CStr(Index))
                        writer.WriteAttribute("style", "color: " + System.Drawing.ColorTranslator.ToHtml(Items(Count).TextItems(Index).Clr) + ";")
                    End If
                    writer.Write(HtmlTextWriter.TagRightChar)
                    writer.Write(Utility.HtmlTextEncode(CStr(Items(Count).TextItems(Index).Text)).Replace(vbCrLf, "<br>"))
                End If
                writer.WriteEndTag(CStr(IIf(Items(Count).TextItems(Index).DisplayClass = RenderDisplayClass.eNested Or Items(Count).TextItems(Index).DisplayClass = RenderDisplayClass.eRanking Or Items(Count).TextItems(Index).DisplayClass = RenderDisplayClass.eList, "div", "span")))
            Next

            If NestPrefix = String.Empty Then
                writer.Write(vbCrLf + BaseTabs + vbTab)
                writer.WriteBeginTag("span")
                writer.WriteAttribute("class", "copy")
                writer.Write(HtmlTextWriter.TagRightChar)

                writer.WriteBeginTag("input")
                writer.WriteAttribute("value", "Copy")
                writer.WriteAttribute("onclick", "javascript: setClipboardText(getText('ri" + CStr(IIf(NestPrefix = String.Empty, Count, NestPrefix)) + "', '" + CStr(IIf(NestPrefix = String.Empty, String.Empty, "ri" + NestPrefix + "_" + CStr(Count))) + "'));")
                writer.WriteAttribute("type", "button")
                writer.Write(HtmlTextWriter.TagRightChar)

                writer.WriteEndTag("span")
            End If

            writer.Write(vbCrLf + BaseTabs)
            writer.WriteEndTag("div")
        Next
    End Sub
End Class
Public Class MailDispatcher
    Public Shared Sub SendEMail(ByVal EMail As String, ByVal Subject As String, ByVal Body As String)
        Dim SmtpClient As New Net.Mail.SmtpClient
        'encrypt and unencrypt password credential
        SmtpClient.Credentials = New Net.NetworkCredential(Utility.ConnectionData.IslamSourceAdminEMail, Utility.ConnectionData.IslamSourceAdminEMailPass)
        SmtpClient.Port = 587
        SmtpClient.Host = Utility.ConnectionData.IslamSourceMailServer
        Dim SmtpMail As New Net.Mail.MailMessage
        SmtpMail.From = New Net.Mail.MailAddress(Utility.ConnectionData.IslamSourceAdminEMail, Utility.ConnectionData.IslamSourceAdminName)
        SmtpMail.To.Add(EMail)
        SmtpMail.Subject = Subject
        SmtpMail.Body = Body
        Try
            SmtpClient.Send(SmtpMail)
        Catch eException As Net.Mail.SmtpException
        End Try
    End Sub
    Public Shared Sub SendActivationEMail(ByVal UserName As String, ByVal EMail As String, ByVal UserID As Integer, ByVal ActivationCode As Integer)
        SendEMail(EMail, String.Format(Utility.LoadResourceString("Acct_ActivationAccountSubject"), HttpContext.Current.Request.Url.Host), _
            String.Format(Utility.LoadResourceString("Acct_ActivationAccountBody"), HttpContext.Current.Request.Url.Host, UserName, "http://" + HttpContext.Current.Request.Url.Host + "/" + Utility.GetPageString("ActivateAccount&UserID=" + CStr(UserID) + "&ActivationCode=" + CStr(ActivationCode)), "http://" + HttpContext.Current.Request.Url.Host + "/" + Utility.GetPageString("ActivateAccount"), CStr(ActivationCode)))
    End Sub
    Public Shared Sub SendUserNameReminderEMail(ByVal UserName As String, ByVal EMail As String)
        SendEMail(EMail, String.Format(Utility.LoadResourceString("Acct_UsernameReminderSubject"), HttpContext.Current.Request.Url.Host), _
            String.Format(Utility.LoadResourceString("Acct_UsernameReminderBody"), HttpContext.Current.Request.Url.Host, UserName))
    End Sub
    Public Shared Sub SendPasswordResetEMail(ByVal UserName As String, ByVal EMail As String, ByVal UserID As Integer, ByVal PasswordResetCode As UInteger)
        SendEMail(EMail, String.Format(Utility.LoadResourceString("Acct_PasswordResetSubject"), HttpContext.Current.Request.Url.Host), _
            String.Format(Utility.LoadResourceString("Acct_PasswordResetBody"), HttpContext.Current.Request.Url.Host, UserName, "http://" + HttpContext.Current.Request.Url.Host + "/" + Utility.GetPageString("ResetPassword&UserID=" + CStr(UserID) + "&PasswordResetCode=" + CStr(PasswordResetCode)), "http://" + HttpContext.Current.Request.Url.Host + "/" + Utility.GetPageString("ResetPassword"), CStr(PasswordResetCode)))
    End Sub
    Public Shared Sub SendUserNameChangedEMail(ByVal UserName As String, ByVal EMail As String)
        SendEMail(EMail, String.Format(Utility.LoadResourceString("Acct_UsernameChangedSubject"), HttpContext.Current.Request.Url.Host), _
            String.Format(Utility.LoadResourceString("Acct_UsernameChangedBody"), HttpContext.Current.Request.Url.Host, UserName))
    End Sub
    Public Shared Sub SendPasswordChangedEMail(ByVal UserName As String, ByVal EMail As String)
        SendEMail(EMail, String.Format(Utility.LoadResourceString("Acct_PasswordChangedSubject"), HttpContext.Current.Request.Url.Host), _
            String.Format(Utility.LoadResourceString("Acct_PasswordChangedBody"), HttpContext.Current.Request.Url.Host, UserName))
    End Sub
End Class
Public Class SiteDatabase
    Public Shared Function GetConnection() As MySql.Data.MySqlClient.MySqlConnection
        Dim Connection As MySql.Data.MySqlClient.MySqlConnection = New MySql.Data.MySqlClient.MySqlConnection("Server=" + Utility.ConnectionData.DbConnServer + ";Uid=" + Utility.ConnectionData.DbConnUid + ";Pwd=" + Utility.ConnectionData.DbConnPwd + ";Database=" + Utility.ConnectionData.DbConnDatabase + ";")
        Try
            Connection.Open()
        Catch e As MySql.Data.MySqlClient.MySqlException
            Return Nothing
        Catch e As TimeoutException
            Return Nothing
        End Try
        Return Connection
    End Function
    Public Shared Sub ExecuteNonQuery(ByVal Connection As MySql.Data.MySqlClient.MySqlConnection, ByVal Query As String, Optional Parameters As Generic.Dictionary(Of String, Object) = Nothing)
        Dim Command As MySql.Data.MySqlClient.MySqlCommand = Connection.CreateCommand()
        Command.CommandText = Query
        If Not Parameters Is Nothing Then
            For Each Key As String In Parameters.Keys
                Command.Parameters.AddWithValue(Key, Parameters(Key))
            Next
        End If
        Command.ExecuteNonQuery()
    End Sub
    Public Shared Sub CreateDatabase()
        Dim Connection As MySql.Data.MySqlClient.MySqlConnection = GetConnection()
        If Connection Is Nothing Then Return
        'SHA1 produces 20 bytes not available in MySQL 5.1
        'should salt the password
        ExecuteNonQuery(Connection, "CREATE TABLE Users (UserID int NOT NULL AUTO_INCREMENT, " + _
        "PRIMARY KEY(UserID), " + _
        "UserName VARCHAR(15) UNIQUE, " + _
        "Password BINARY(20), " + _
        "EMail VARCHAR(254) UNIQUE, " + _
        "Access int NOT NULL DEFAULT 0, " + _
        "ActivationCode int, " + _
        "LoginSecret int DEFAULT NULL, " + _
        "LoginTime TIMESTAMP NULL)")
        Connection.Close()
    End Sub
    Public Shared Sub RemoveDatabase()
        Dim Connection As MySql.Data.MySqlClient.MySqlConnection = GetConnection()
        If Connection Is Nothing Then Return
        Dim Command As MySql.Data.MySqlClient.MySqlCommand = Connection.CreateCommand()
        ExecuteNonQuery(Connection, "DROP TABLE Users")
        Connection.Close()
    End Sub
    Public Shared Sub CleanupStaleActivations()
        Dim Connection As MySql.Data.MySqlClient.MySqlConnection = GetConnection()
        If Connection Is Nothing Then Return
        ExecuteNonQuery(Connection, "DELETE FROM Users WHERE ActivationCode IS NOT NULL AND (LoginTime IS NULL OR UTC_TIMESTAMP > TIMESTAMPADD(DAY, 10, LoginTime))")
        Connection.Close()
    End Sub
    Public Shared Sub CleanupStaleLoginSessions()
        Dim Connection As MySql.Data.MySqlClient.MySqlConnection = GetConnection()
        If Connection Is Nothing Then Return
        ExecuteNonQuery(Connection, "UPDATE Users SET LoginSecret=NULL, LoginTime=NULL WHERE ActivationCode IS NOT NULL AND (LoginTime IS NOT NULL AND UTC_TIMESTAMP > TIMESTAMPADD(HOUR, 1, LoginTime))")
        Connection.Close()
    End Sub
    Public Shared Sub AddUser(ByVal UserName As String, ByVal Password As String, ByVal EMail As String)
        Dim Connection As MySql.Data.MySqlClient.MySqlConnection = GetConnection()
        If Connection Is Nothing Then Return
        Dim Generator As New System.Random()
        ExecuteNonQuery(Connection, "INSERT INTO Users (UserName, Password, EMail, ActivationCode, LoginTime) VALUES (@UserName, UNHEX(SHA1(@Password)), @EMail, @Code, UTC_TIMESTAMP)", _
                        New Generic.Dictionary(Of String, Object) From {{"@UserName", UserName}, {"@Password", Password}, {"@EMail", EMail}, {"@Code", CStr(Generator.Next(0, 99999999))}})
        Connection.Close()
    End Sub
    Public Shared Function GetUserID(ByVal UserName As String, ByVal Password As String) As Integer
        Dim Connection As MySql.Data.MySqlClient.MySqlConnection = GetConnection()
        If Connection Is Nothing Then Return -1
        Dim Command As MySql.Data.MySqlClient.MySqlCommand = Connection.CreateCommand()
        Command.CommandText = "SELECT UserID FROM Users WHERE UserName=@UserName AND Password=UNHEX(SHA1(@Password))"
        Command.Parameters.AddWithValue("@UserName", UserName)
        Command.Parameters.AddWithValue("@Password", Password)
        Dim Reader As MySql.Data.MySqlClient.MySqlDataReader = Command.ExecuteReader()
        If Reader.Read() AndAlso Not Reader.IsDBNull(0) Then
            GetUserID = Reader.GetInt32("UserID")
        Else
            GetUserID = -1
        End If
        Reader.Close()
        Connection.Close()
    End Function
    Public Shared Function GetUserID(ByVal UserName As String) As Integer
        Dim Connection As MySql.Data.MySqlClient.MySqlConnection = GetConnection()
        If Connection Is Nothing Then Return -1
        Dim Command As MySql.Data.MySqlClient.MySqlCommand = Connection.CreateCommand()
        Command.CommandText = "SELECT UserID FROM Users WHERE UserName=@UserName"
        Command.Parameters.AddWithValue("@UserName", UserName)
        Dim Reader As MySql.Data.MySqlClient.MySqlDataReader = Command.ExecuteReader()
        If Reader.Read() AndAlso Not Reader.IsDBNull(0) Then
            GetUserID = Reader.GetInt32("UserID")
        Else
            GetUserID = -1
        End If
        Reader.Close()
        Connection.Close()
    End Function
    Public Shared Function GetUserIDByEMail(ByVal EMail As String) As Integer
        Dim Connection As MySql.Data.MySqlClient.MySqlConnection = GetConnection()
        If Connection Is Nothing Then Return -1
        Dim Command As MySql.Data.MySqlClient.MySqlCommand = Connection.CreateCommand()
        Command.CommandText = "SELECT UserID FROM Users WHERE EMail=@EMail"
        Command.Parameters.AddWithValue("@EMail", EMail)
        Dim Reader As MySql.Data.MySqlClient.MySqlDataReader = Command.ExecuteReader()
        If Reader.Read() AndAlso Not Reader.IsDBNull(0) Then
            GetUserIDByEMail = Reader.GetInt32("UserID")
        Else
            GetUserIDByEMail = -1
        End If
        Reader.Close()
        Connection.Close()
    End Function
    Public Shared Function GetUserResetCode(ByVal UserID As Integer) As UInteger
        Dim Connection As MySql.Data.MySqlClient.MySqlConnection = GetConnection()
        If Connection Is Nothing Then Return &HFFFFFFFFL
        Dim Command As MySql.Data.MySqlClient.MySqlCommand = Connection.CreateCommand()
        Command.CommandText = "SELECT CRC32(Password) FROM Users WHERE UserID=" + CStr(UserID)
        Dim Reader As MySql.Data.MySqlClient.MySqlDataReader = Command.ExecuteReader()
        If Reader.Read() AndAlso Not Reader.IsDBNull(0) Then
            GetUserResetCode = Reader.GetUInt32("CRC32(Password)")
        Else
            GetUserResetCode = &HFFFFFFFFL
        End If
        Reader.Close()
        Connection.Close()
    End Function
    Public Shared Function CheckLogin(ByVal UserID As Integer, ByVal Secret As Integer) As Boolean
        Dim Connection As MySql.Data.MySqlClient.MySqlConnection = GetConnection()
        If Connection Is Nothing Then Return False
        Dim Command As MySql.Data.MySqlClient.MySqlCommand = Connection.CreateCommand()
        Command.CommandText = "SELECT UserID FROM Users WHERE UserID=" + CStr(UserID) + " AND ActivationCode IS NULL AND LoginSecret=" + CStr(Secret) + CStr(IIf(Secret Mod 2 = 0, " AND LoginTime IS NULL", " AND LoginTime IS NOT NULL AND UTC_TIMESTAMP < TIMESTAMPADD(HOUR, 1, LoginTime)"))
        Dim Reader As MySql.Data.MySqlClient.MySqlDataReader = Command.ExecuteReader()
        If Reader.Read() AndAlso Not Reader.IsDBNull(0) Then
            CheckLogin = CInt(Reader.Item("UserID")) = UserID
        Else
            CheckLogin = False
        End If
        Reader.Close()
        Connection.Close()
    End Function
    Public Shared Function CheckAccess(ByVal UserID As Integer) As Integer
        Dim Connection As MySql.Data.MySqlClient.MySqlConnection = GetConnection()
        If Connection Is Nothing Then Return 0
        Dim Command As MySql.Data.MySqlClient.MySqlCommand = Connection.CreateCommand()
        Command.CommandText = "SELECT Access FROM Users WHERE UserID=" + CStr(UserID)
        Dim Reader As MySql.Data.MySqlClient.MySqlDataReader = Command.ExecuteReader()
        If Reader.Read() AndAlso Not Reader.IsDBNull(0) Then
            CheckAccess = CInt(Reader.Item("Access"))
        Else
            CheckAccess = 0
        End If
        Reader.Close()
        Connection.Close()
    End Function
    Public Shared Function SetLogin(ByVal UserID As Integer, ByVal Persist As Boolean) As Integer
        Dim Connection As MySql.Data.MySqlClient.MySqlConnection = GetConnection()
        If Connection Is Nothing Then Return -1
        Dim Generator As New System.Random()
        'Persistant login secret is even, non-persistant is odd
        SetLogin = (Generator.Next(0, 99999999) \ 2) * 2 + CInt(IIf(Persist, 0, 1))
        ExecuteNonQuery(Connection, "UPDATE Users SET LoginSecret=" + CStr(SetLogin) + ", LoginTime=" + CStr(IIf(Persist, "NULL", "UTC_TIMESTAMP")) + " WHERE UserID=" + CStr(UserID))
        Connection.Close()
    End Function
    Public Shared Sub ClearLogin(ByVal UserID As Integer)
        Dim Connection As MySql.Data.MySqlClient.MySqlConnection = GetConnection()
        If Connection Is Nothing Then Return
        ExecuteNonQuery(Connection, "UPDATE Users SET LoginSecret=NULL, LoginTime=NULL WHERE UserID=" + CStr(UserID))
        Connection.Close()
    End Sub
    Public Shared Function GetUserActivated(ByVal UserID As Integer) As Integer
        Dim Connection As MySql.Data.MySqlClient.MySqlConnection = GetConnection()
        If Connection Is Nothing Then Return -1
        Dim Command As MySql.Data.MySqlClient.MySqlCommand = Connection.CreateCommand()
        Command.CommandText = "SELECT ActivationCode FROM Users WHERE UserID=" + CStr(UserID)
        Dim Reader As MySql.Data.MySqlClient.MySqlDataReader = Command.ExecuteReader()
        'If Reader.IsDBNull(0) Then Return -1 'NULL is activated
        If Not Reader.Read() OrElse Reader.IsDBNull(0) Then
            GetUserActivated = -1 'NULL is activated
        Else
            GetUserActivated = Reader.GetInt32("ActivationCode")
        End If
        Reader.Close()
        Connection.Close()
    End Function
    Public Shared Function GetUserName(ByVal UserID As Integer) As String
        Dim Connection As MySql.Data.MySqlClient.MySqlConnection = GetConnection()
        If Connection Is Nothing Then Return String.Empty
        Dim Command As MySql.Data.MySqlClient.MySqlCommand = Connection.CreateCommand()
        Command.CommandText = "SELECT UserName FROM Users WHERE UserID=" + CStr(UserID)
        Dim Reader As MySql.Data.MySqlClient.MySqlDataReader = Command.ExecuteReader()
        If Reader.Read() AndAlso Not Reader.IsDBNull(0) Then
            GetUserName = Reader.GetString("UserName")
        Else
            GetUserName = String.Empty
        End If
        Reader.Close()
        Connection.Close()
    End Function
    Public Shared Function GetUserEMail(ByVal UserID As Integer) As String
        Dim Connection As MySql.Data.MySqlClient.MySqlConnection = GetConnection()
        If Connection Is Nothing Then Return String.Empty
        Dim Command As MySql.Data.MySqlClient.MySqlCommand = Connection.CreateCommand()
        Command.CommandText = "SELECT EMail FROM Users WHERE UserID=" + CStr(UserID)
        Dim Reader As MySql.Data.MySqlClient.MySqlDataReader = Command.ExecuteReader()
        If Reader.Read() AndAlso Not Reader.IsDBNull(0) Then
            GetUserEMail = Reader.GetString("EMail")
        Else
            GetUserEMail = String.Empty
        End If
        Reader.Close()
        Connection.Close()
    End Function
    Public Shared Sub ChangeUserName(ByVal UserID As Integer, ByVal UserName As String)
        Dim Connection As MySql.Data.MySqlClient.MySqlConnection = GetConnection()
        If Connection Is Nothing Then Return
        ExecuteNonQuery(Connection, "UPDATE Users SET UserName=@UserName WHERE UserID=" + CStr(UserID), New Generic.Dictionary(Of String, Object) From {{"@UserName", UserName}})
        Connection.Close()
    End Sub
    Public Shared Sub ChangeUserPassword(ByVal UserID As Integer, ByVal Password As String)
        Dim Connection As MySql.Data.MySqlClient.MySqlConnection = GetConnection()
        If Connection Is Nothing Then Return
        ExecuteNonQuery(Connection, "UPDATE Users SET Password=UNHEX(SHA1(@Password)) WHERE UserID=" + CStr(UserID), New Generic.Dictionary(Of String, Object) From {{"@Password", Password}})
        Connection.Close()
    End Sub
    Public Shared Sub ChangeUserEMail(ByVal UserID As Integer, ByVal EMail As String)
        Dim Connection As MySql.Data.MySqlClient.MySqlConnection = GetConnection()
        If Connection Is Nothing Then Return
        Dim Generator As New System.Random()
        ExecuteNonQuery(Connection, "UPDATE Users SET EMail=@EMail, ActivationCode='" + CStr(Generator.Next(0, 99999999)) + "' WHERE UserID=" + CStr(UserID), New Generic.Dictionary(Of String, Object) From {{"@EMail", EMail}})
        Connection.Close()
    End Sub
    Public Shared Sub SetUserActivated(ByVal UserID As Integer)
        Dim Connection As MySql.Data.MySqlClient.MySqlConnection = GetConnection()
        If Connection Is Nothing Then Return
        Dim Generator As New System.Random()
        ExecuteNonQuery(Connection, "UPDATE Users SET ActivationCode=NULL, LoginTime=NULL WHERE UserID=" + CStr(UserID))
        Connection.Close()
    End Sub
    Public Shared Sub RemoveUser(ByVal UserID As Integer)
        Dim Connection As MySql.Data.MySqlClient.MySqlConnection = GetConnection()
        If Connection Is Nothing Then Return
        ExecuteNonQuery(Connection, "DELETE FROM Users WHERE UserID=" + CStr(UserID))
        Connection.Close()
    End Sub
End Class