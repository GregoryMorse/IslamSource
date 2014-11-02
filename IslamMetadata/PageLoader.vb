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
        Public Shared ReadOnly Property GlobalRes As String
            Get
                Return GetConfigSetting("globalres")
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
        If resourceKey.StartsWith("Acct_") Or _
            resourceKey.StartsWith("Hadith_") Or _
            resourceKey.StartsWith("IslamInfo_") Or _
            resourceKey.StartsWith("IslamSource_") Or _
            resourceKey.StartsWith("lang_") Or _
            resourceKey.StartsWith("unicode_") Or resourceKey = "IslamSource" Then
            'LoadResourceString = CStr(HttpContext.GetLocalResourceObject(LocalFile, resourceKey))
            LoadResourceString = New System.Resources.ResourceManager("IslamResources.Resources", Reflection.Assembly.Load("IslamResources")).GetString(resourceKey, Threading.Thread.CurrentThread.CurrentUICulture)
        Else
            'LoadResourceString = CStr(HttpContext.GetGlobalResourceObject(ConnectionData.GlobalRes, resourceKey))
            LoadResourceString = New System.Resources.ResourceManager("GMorseCodeResources.Resources", Reflection.Assembly.Load("GMorseCodeResources")).GetString(resourceKey, Threading.Thread.CurrentThread.CurrentUICulture)
        End If
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
        If (ClassMember.Length = 3 AndAlso ClassMember(1) = String.Empty) Then
            Return Type.GetType("IslamMetadata." + ClassMember(0)).GetMethod(ClassMember(2))
        End If
        Return Nothing
    End Function
    Public Shared Function TextRender(ByVal Item As PageLoader.TextItem) As String
        Return Item.Text
    End Function
    Public Shared Function GetOnPrintJS() As String()
        Return New String() {"javascript: openPrintable(this);", String.Empty, "function openPrintable(btn) { var input = document.createElement('input'); input.type = 'hidden'; input.name = 'PagePrint'; input.value = btn.form.elements['Page'].value; btn.form.appendChild(input); btn.form.target = '_blank'; btn.form.elements['Page'].value = 'Print'; btn.form.submit(); btn.form.target = ''; btn.form.elements['Page'].value = btn.form.elements['PagePrint'].value; btn.form.removeChild(input); }"}
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
        Return "function isInArray(array, item) { var length = array.length; for (var count = 0; count < length; count++) { if (array[count] == item) return true; } return false; }"
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
Public Class PrayerTime
    Public Shared Function GetMonthName(ByVal Item As PageLoader.TextItem) As String
        Dim CultureInfo As Globalization.CultureInfo
        If Item.Name = "hijrimonthname" Then
            CultureInfo = New Globalization.CultureInfo("ar-SA")
            CultureInfo.DateTimeFormat.Calendar = New Globalization.HijriCalendar
        ElseIf Item.Name = "umalquramonthname" Then
            CultureInfo = New Globalization.CultureInfo("ar-SA")
            CultureInfo.DateTimeFormat.Calendar = New Globalization.UmAlQuraCalendar
        Else
            CultureInfo = Globalization.CultureInfo.CurrentCulture 'Globalization.CultureInfo.CurrentCulture.LCID
        End If
        'If Array.Exists(Globalization.CultureInfo.CurrentCulture.OptionalCalendars, Function(Cal As Globalization.Calendar) Cal.ToString() = Calendar.ToString()) Then
        GetMonthName = CultureInfo.DateTimeFormat.MonthNames(CultureInfo.DateTimeFormat.Calendar.GetMonth(Today) - 1)
    End Function
    Public Shared Function GetCalendar(ByVal Item As PageLoader.TextItem) As Array()
        Dim Count As Integer
        Dim Calendar As Globalization.Calendar
        If Item.Name = "hijricalendar" Then
            Calendar = New Globalization.HijriCalendar
        ElseIf Item.Name = "umalquracalendar" Then
            Calendar = New Globalization.UmAlQuraCalendar
        Else
            Calendar = Globalization.CultureInfo.CurrentCulture.Calendar
        End If
        Dim RetArray(Calendar.GetWeekOfYear(New Date(Calendar.GetYear(Today), Calendar.GetMonth(Today), Calendar.GetDaysInMonth(Calendar.GetYear(Today), Calendar.GetMonth(Today)), Calendar), Globalization.CalendarWeekRule.FirstDay, DayOfWeek.Sunday) - Calendar.GetWeekOfYear(New Date(Calendar.GetYear(Today), Calendar.GetMonth(Today), 1, Calendar), Globalization.CalendarWeekRule.FirstDay, DayOfWeek.Sunday) + 1 + 2) As Array
        RetArray(0) = New String() {}
        RetArray(1) = New String() {String.Empty, String.Empty, String.Empty, String.Empty, String.Empty, String.Empty, String.Empty}
        RetArray(2) = New String() {Globalization.CultureInfo.CurrentCulture.DateTimeFormat.DayNames(0), Globalization.CultureInfo.CurrentCulture.DateTimeFormat.DayNames(1), Globalization.CultureInfo.CurrentCulture.DateTimeFormat.DayNames(2), Globalization.CultureInfo.CurrentCulture.DateTimeFormat.DayNames(3), Globalization.CultureInfo.CurrentCulture.DateTimeFormat.DayNames(4), Globalization.CultureInfo.CurrentCulture.DateTimeFormat.DayNames(5), Globalization.CultureInfo.CurrentCulture.DateTimeFormat.DayNames(6)}
        For Count = 1 To Calendar.GetDaysInMonth(Calendar.GetYear(Today), Calendar.GetMonth(Today))
            RetArray(2 + 1 + Calendar.GetWeekOfYear(New Date(Calendar.GetYear(Today), Calendar.GetMonth(Today), Count, Calendar), Globalization.CalendarWeekRule.FirstDay, DayOfWeek.Sunday) - Calendar.GetWeekOfYear(New Date(Calendar.GetYear(Today), Calendar.GetMonth(Today), 1, Calendar), Globalization.CalendarWeekRule.FirstDay, DayOfWeek.Sunday)) = New String() {String.Empty, String.Empty, String.Empty, String.Empty, String.Empty, String.Empty, String.Empty}
            Do
                CType(RetArray(2 + 1 + Calendar.GetWeekOfYear(New Date(Calendar.GetYear(Today), Calendar.GetMonth(Today), Count, Calendar), Globalization.CalendarWeekRule.FirstDay, DayOfWeek.Sunday) - Calendar.GetWeekOfYear(New Date(Calendar.GetYear(Today), Calendar.GetMonth(Today), 1, Calendar), Globalization.CalendarWeekRule.FirstDay, DayOfWeek.Sunday)), String())(Calendar.GetDayOfWeek(New Date(Calendar.GetYear(Today), Calendar.GetMonth(Today), Count, Calendar))) = CStr(IIf(Calendar.GetDayOfMonth(Today) = Count, ">", String.Empty)) + CStr(Calendar.GetDayOfMonth(New Date(Calendar.GetYear(Today), Calendar.GetMonth(Today), Count, Calendar))) + CStr(IIf(Calendar.GetDayOfMonth(Today) = Count, "<", String.Empty))
                Count += 1
            Loop While Count <= Calendar.GetDaysInMonth(Calendar.GetYear(Today), Calendar.GetMonth(Today)) AndAlso _
                Calendar.GetDayOfWeek(New Date(Calendar.GetYear(Today), Calendar.GetMonth(Today), Count, Calendar)) <> DayOfWeek.Sunday
            Count -= 1
        Next
        Return RetArray
    End Function
    Public Shared Function GetPrayerTimes(ByVal Item As PageLoader.TextItem) As Array()
        Dim Strings As String() = Geolocation.GetGeoData()
        If Strings.Length <> 11 OrElse Strings(0) = "ERROR" Then Return New Array() {}
        Dim GeoData As String = Geolocation.GetElevationData(Strings(8), Strings(9))
        Dim PrayTimes As New PrayTime.PrayTime
        Dim Count As Integer
        'Dim Times As String() = PrayTimes.getDatePrayerTimes(Today.Year, Today.Month, Today.Day, CDbl(Strings(8)), CDbl(Strings(9)), CInt(Strings(10).Split(":")(0)) + IIf(CInt(Strings(10).Split(":")(0)) >= 0, CInt(Strings(10).Split(":")(1)) / 60, -CInt(Strings(10).Split(":")(1)) / 60), 0)
        Dim RetArray(Date.DaysInMonth(Today.Year, Today.Month) + 2) As Array
        RetArray(0) = New String() {}
        RetArray(1) = New String() {String.Empty, String.Empty, String.Empty, String.Empty, String.Empty, String.Empty, String.Empty, String.Empty, String.Empty, String.Empty, String.Empty}
        RetArray(2) = New String() {Utility.LoadResourceString("IslamInfo_Date"), Utility.LoadResourceString("IslamInfo_Day"), Utility.LoadResourceString("IslamInfo_PrayTime6"), Utility.LoadResourceString("IslamInfo_PrayTime1"), Utility.LoadResourceString("IslamInfo_PrayTime7"), Utility.LoadResourceString("IslamInfo_PrayTime2"), Utility.LoadResourceString("IslamInfo_PrayTime3"), Utility.LoadResourceString("IslamInfo_PrayTime8"), Utility.LoadResourceString("IslamInfo_PrayTime4"), Utility.LoadResourceString("IslamInfo_PrayTime5"), Utility.LoadResourceString("IslamInfo_PrayTime9")}
        For Count = 1 To Date.DaysInMonth(Today.Year, Today.Month)
            Dim Times As String() = PrayTimes.getDatePrayerTimes(Today.Year, Today.Month, Count, CDbl(Strings(8)), CDbl(Strings(9)), CInt(Strings(10).Split(":"c)(0)) + CInt(IIf(CInt(Strings(10).Split(":"c)(0)) >= 0, CInt(Strings(10).Split(":"c)(1)) / 60, -CInt(Strings(10).Split(":"c)(1)) / 60)), CInt(GeoData))
            RetArray(Count + 2) = New String() {CStr(Count), New Date(Today.Year, Today.Month, Count).ToString("dddd", Globalization.CultureInfo.CurrentCulture), Times(0), Times(1), Times(2), Times(3), Times(4), Times(5), Times(6), Times(7), Times(8)}
        Next
        Return RetArray
    End Function
    Public Shared Function GetQiblaDirection(ByVal Item As PageLoader.TextItem) As String
        Const QiblaLat As Double = 21.42252
        Const QiblaLon As Double = 39.82621
        Dim Strings As String() = Geolocation.GetGeoData()
        If Strings.Length <> 11 Then Return String.Empty
        Return DegreeBearing(CDbl(Strings(8)), CDbl(Strings(9)), QiblaLat, QiblaLon).ToString() + " " + SphericalDistance(QiblaLat, QiblaLon, CDbl(Strings(8)), CDbl(Strings(9))).ToString()
    End Function

    Public Shared Function SphericalDistance(ByVal lat1 As Double, ByVal lon1 As Double, ByVal lat2 As Double, ByVal lon2 As Double) As Double
        Const R As Double = 6378.137 'earth’s mean radius (volumetric radius = 6,371km) according to the WGS84 system
        Dim dLon As Double = ToRad(lon2 - lon1)
        Dim dLat As Double = ToRad(lat2 - lat1)
        Dim a As Double = Math.Sin(dLat / 2) * Math.Sin(dLat / 2) + Math.Sin(dLon / 2) * Math.Sin(dLon / 2) * Math.Cos(ToRad(lat1)) * Math.Cos(ToRad(lat2))
        Return R * 2 * Math.Atan2(Math.Sqrt(a), Math.Sqrt(1 - a))
    End Function
    Public Shared Function DegreeBearing(ByVal lat1 As Double, ByVal lon1 As Double, ByVal lat2 As Double, ByVal lon2 As Double) As Double
        Dim dLon As Double = ToRad(lon1 - lon2)
        'Great circle
        Return ToBearing(-Math.Atan2(Math.Sin(dLon), Math.Cos(ToRad(lat1)) * Math.Tan(ToRad(lat2)) - Math.Sin(ToRad(lat1)) * Math.Cos(dLon)))
        'Rhumm lines is incorrect
        'Dim dPhi As Double = Math.Log(Math.Tan(ToRad(lat2) / 2 + Math.PI / 4) / Math.Tan(ToRad(lat1) / 2 + Math.PI / 4))
        'If (Math.Abs(dLon) > Math.PI) Then dLon = IIf(dLon > 0, -(2 * Math.PI - dLon), (2 * Math.PI + dLon))
        'ToBearing(Math.Atan2(dLon, dPhi))
    End Function
    Public Shared Function ToRad(ByVal degrees As Double) As Double
        Return degrees * (Math.PI / 180)
    End Function
    Public Shared Function ToDegrees(ByVal radians As Double) As Double
        Return radians * 180 / Math.PI
    End Function
    Public Shared Function ToBearing(ByVal radians As Double) As Double
        'convert radians to degrees (as bearing: 0...360)
        Return (ToDegrees(radians) + 360) Mod 360
    End Function
End Class
Public Class Arabic
    Class StringLengthComparer
        Implements Collections.IComparer
        Public Sub New(Scheme As String)
            _Scheme = Scheme
        End Sub
        Private _Scheme As String
        Public Function Compare(ByVal x As Object, ByVal y As Object) As Integer _
            Implements Collections.IComparer.Compare
            Compare = GetSchemeValueFromSymbol(DirectCast(x, IslamData.ArabicSymbol), _Scheme).Length - _
                GetSchemeValueFromSymbol(DirectCast(y, IslamData.ArabicSymbol), _Scheme).Length
        End Function
    End Class
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
            Return Utility.LoadResourceString("unicode_" + CachedData.IslamData.ArabicLetters(FindLetterBySymbol(Character)).UnicodeName)
        End Try
        Return Str.ToString()
    End Function
    Public Shared Function TransliterateFromBuckwalter(ByVal Buckwalter As String) As String
        Dim ArabicString As String = String.Empty
        Dim Count As Integer
        Dim Index As Integer
        For Count = 0 To Buckwalter.Length - 1
            If Buckwalter(Count) = "\" Then
                Count += 1
                If Buckwalter(Count) = "," Then
                    ArabicString += ArabicComma
                Else
                    ArabicString += Buckwalter(Count)
                End If
            Else
                For Index = 0 To CachedData.IslamData.ArabicLetters.Length - 1
                    If Buckwalter(Count) = CachedData.IslamData.ArabicLetters(Index).ExtendedBuckwalterLetter Then
                        ArabicString += CachedData.IslamData.ArabicLetters(Index).Symbol
                        Exit For
                    End If
                Next
                If Index = CachedData.IslamData.ArabicLetters.Length Then
                    ArabicString += Buckwalter(Count)
                End If
            End If
        Next
        Return ArabicString
    End Function
    Public Enum TranslitScheme
        None = 0
        Literal = 1
        RuleBased = 2
    End Enum
    Public Shared Function TransliterateToScheme(ByVal ArabicString As String, SchemeType As TranslitScheme, Scheme As String) As String
        If SchemeType = TranslitScheme.RuleBased Then
            Return TransliterateWithRules(ArabicString, Scheme)
        ElseIf SchemeType = TranslitScheme.Literal Then
            Return TransliterateToRoman(ArabicString, Scheme)
        Else
            Return New String(System.Array.FindAll(ArabicString.ToCharArray(), Function(Check As Char) Check = " "c))
        End If
    End Function
    Public Shared Function GetSchemeValueFromSymbol(Symbol As IslamData.ArabicSymbol, Scheme As String) As String
        Dim Sch As IslamData.TranslitScheme = Nothing
        Dim Count As Integer
        For Count = 0 To CachedData.IslamData.TranslitSchemes.Length - 1
            If CachedData.IslamData.TranslitSchemes(Count).Name = Scheme Then
                Sch = CachedData.IslamData.TranslitSchemes(Count)
            End If
        Next
        If Count = CachedData.IslamData.TranslitSchemes.Length Then Return String.Empty
        If Array.IndexOf(ArabicLettersInOrder, Symbol.Symbol) <> 0 Then
            Return Sch.Alphabet(Array.IndexOf(ArabicLettersInOrder, Symbol.Symbol))
        ElseIf Array.IndexOf(ArabicHamzas, Symbol.Symbol) <> 0 Then
            Return Sch.Hamza(Array.IndexOf(ArabicHamzas, Symbol.Symbol))
        ElseIf Array.IndexOf(ArabicSpecialLetters, Symbol.Symbol) <> 0 Then
            Return Sch.SpecialLetters(Array.IndexOf(ArabicSpecialLetters, Symbol.Symbol))
        ElseIf Array.IndexOf(ArabicVowels, Symbol.Symbol) <> 0 Then
            Return Sch.Vowels(Array.IndexOf(ArabicVowels, Symbol.Symbol))
        ElseIf Array.IndexOf(ArabicTajweed, Symbol.Symbol) <> 0 Then
            Return Sch.Tajweed(Array.IndexOf(ArabicTajweed, Symbol.Symbol))
        ElseIf Array.IndexOf(ArabicPunctuation, Symbol.Symbol) <> 0 Then
            Return Sch.Punctuation(Array.IndexOf(ArabicPunctuation, Symbol.Symbol))
        ElseIf Array.IndexOf(NonArabicLetters, Symbol.Symbol) <> 0 Then
            Return Sch.NonArabic(Array.IndexOf(NonArabicLetters, Symbol.Symbol))
        End If
        Return String.Empty
    End Function
    Public Shared Function TransliterateToRoman(ByVal ArabicString As String, Scheme As String) As String
        Dim RomanString As String = String.Empty
        Dim Count As Integer
        Dim Index As Integer
        Dim Letters(CachedData.IslamData.ArabicLetters.Length - 1) As IslamData.ArabicSymbol
        CachedData.IslamData.ArabicLetters.CopyTo(Letters, 0)
        Array.Sort(Letters, New StringLengthComparer(Scheme))
        For Count = 0 To ArabicString.Length - 1
            If ArabicString(Count) = "\" Then
                Count += 1
                If ArabicString(Count) = "," Then
                    RomanString += ArabicComma
                Else
                    RomanString += ArabicString(Count)
                End If
            Else
                For Index = 0 To Letters.Length - 1
                    If ArabicString(Count) = Letters(Index).Symbol Then
                        RomanString += CStr(IIf(Scheme = String.Empty, Letters(Index).ExtendedBuckwalterLetter, GetSchemeValueFromSymbol(Letters(Index), Scheme)))
                        Exit For
                    End If
                Next
                If Index = Letters.Length Then
                    RomanString += ArabicString(Count)
                End If
            End If
        Next
        Return RomanString
    End Function
    Public Shared Function FindLetterBySymbol(Symbol As Char) As Integer
        Return Array.FindIndex(CachedData.IslamData.ArabicLetters, Function(Letter As IslamData.ArabicSymbol) Letter.Symbol = Symbol)
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
                If Ranges.Count <> 0 AndAlso Ranges(Ranges.Count - 1)(Ranges(Ranges.Count - 1).Count - 1) + 1 = NewRangeMatch Then
                    Ranges(Ranges.Count - 1).Add(NewRangeMatch)
                Else
                    Ranges.Add(New ArrayList From {NewRangeMatch})
                End If
            End If
        Next
        Return "return " + String.Join("||", Array.ConvertAll(Of ArrayList, String)(Ranges.ToArray(GetType(ArrayList)), Function(Arr As ArrayList) If(Arr.Count = 1, "c===0x" + Hex(Arr(0)), "(c>=0x" + Hex(Arr(0)) + "&&c<=0x" + Hex(Arr(Arr.Count - 1)) + ")"))) + ";"
    End Function
    'ArabicLetterAlefWithMaddahAbove is used in simple script but not uthmani
    'ArabicEndOfAyah is added later
    'ArabicHamzaBelow never used along with ArabicLetterPeh, ArabicLetterTcheh, ArabicLetterVeh, ArabicLetterGaf
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
        Return Array.ConvertAll(RecitationSymbols, Function(Ch As Char) New Object() {CachedData.IslamData.ArabicLetters(FindLetterBySymbol(Ch)).UnicodeName + " (" + Arabic.FixStartingCombiningSymbol(Ch) + Arabic.LeftToRightOverride + ")" + Arabic.PopDirectionalFormatting, FindLetterBySymbol(Ch)})
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
        Return If(Array.IndexOf(RecitationCombiningSymbols, Str.Chars(0)) <> -1 Or Str.Length = 1, Arabic.LeftToRightOverride + Str + Arabic.PopDirectionalFormatting, Str)
    End Function
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
        Return Array.FindIndex(ArabicLetters, Function(Str As String) Str = CachedData.IslamData.ArabicLetters(Index).Symbol) <> -1
    End Function
    Public Shared Function IsPunctuation(Index As Integer) As Boolean
        Return Array.FindIndex(PunctuationSymbols, Function(Str As String) Str = CachedData.IslamData.ArabicLetters(Index).Symbol) <> -1
    End Function
    Public Shared Function IsStop(Index As Integer) As Boolean
        Return Array.FindIndex(ArabicStopLetters, Function(Str As String) Str = CachedData.IslamData.ArabicLetters(Index).Symbol) <> -1
    End Function
    Public Shared Function IsWhitespace(Index As Integer) As Boolean
        Return Array.FindIndex(WhitespaceSymbols, Function(Str As String) Str = CachedData.IslamData.ArabicLetters(Index).Symbol) <> -1
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
    Public Shared ArabicSpecialLeadingGutteral As String() = {ArabicLetterHah, ArabicLetterAin}
    Public Shared ArabicSpecialGutteral As String() = {ArabicLetterHamza, ArabicLetterHah, ArabicLetterAin, ArabicLetterSad, ArabicLetterDad, ArabicLetterTah, ArabicLetterZah}
    Public Shared ArabicLetters As String() = {ArabicLetterTeh, ArabicLetterTheh, ArabicLetterDal, ArabicLetterThal, ArabicLetterReh, ArabicLetterZain, ArabicLetterSeen, ArabicLetterSheen, ArabicLetterSad, ArabicLetterDad, ArabicLetterTah, ArabicLetterZah, ArabicLetterLam, ArabicLetterNoon, ArabicLetterAlef, ArabicLetterBeh, ArabicLetterJeem, ArabicLetterHah, ArabicLetterKhah, ArabicLetterAin, ArabicLetterGhain, ArabicLetterFeh, ArabicLetterQaf, ArabicLetterKaf, ArabicLetterMeem, ArabicLetterHeh, ArabicLetterWaw, ArabicLetterYeh}
    Public Shared ArabicLettersInOrder As String() = {ArabicLetterAlef, ArabicLetterBeh, ArabicLetterTeh, ArabicLetterTheh, ArabicLetterJeem, ArabicLetterHah, ArabicLetterKhah, ArabicLetterDal, ArabicLetterThal, ArabicLetterReh, ArabicLetterZain, ArabicLetterSeen, ArabicLetterSheen, ArabicLetterSad, ArabicLetterDad, ArabicLetterTah, ArabicLetterZah, ArabicLetterAin, ArabicLetterGhain, ArabicLetterFeh, ArabicLetterQaf, ArabicLetterKaf, ArabicLetterLam, ArabicLetterMeem, ArabicLetterNoon, ArabicLetterHeh, ArabicLetterWaw, ArabicLetterYeh}
    Public Shared ArabicHamzas As String() = {ArabicLetterHamza, ArabicLetterAlefWithMaddaAbove, ArabicLetterAlefWithHamzaAbove, ArabicLetterWawWithHamzaAbove, ArabicLetterAlefWithHamzaBelow, ArabicLetterYehWithHamzaAbove, ArabicHamzaAbove, ArabicLetterAlefWasla, ArabicHamzaBelow}
    Public Shared ArabicSpecialLetters As String() = {ArabicLetterTehMarbuta, ArabicLetterTehMarbuta, ArabicLetterTehMarbuta, ArabicLetterAlefMaksura, ArabicLetterSuperscriptAlef, ArabicLetterNoonGhunna}
    Public Shared ArabicVowels As String() = {ArabicFatha, ArabicDamma, ArabicKasra, ArabicFathatan, ArabicDammatan, ArabicKasratan, ArabicFatha + ArabicLetterAlef, ArabicDamma + ArabicLetterWaw, ArabicKasra + ArabicLetterYeh, ArabicFatha + ArabicLetterWaw, ArabicFatha + ArabicLetterYeh, ArabicShadda, ArabicSukun}
    Public Shared ArabicTajweed As String() = {ArabicSmallHighSeen, ArabicSmallHighMeemIsolatedForm, ArabicSmallLowSeen, ArabicSmallWaw, ArabicSmallYeh, ArabicSmallHighNoon, ArabicSmallLowMeem}
    Public Shared ArabicPunctuation As String() = {Space, ExclamationMark, QuotationMark, Comma, HyphenMinus, FullStop, Colon, LeftSquareBracket, RightSquareBracket, LeftCurlyBracket, RightCurlyBracket, NoBreakSpace, LeftPointingDoubleAngleQuotationMark, RightPointingDoubleAngleQuotationMark, ArabicComma, ArabicSemicolon, ArabicQuestionMark, ZeroWidthNonJoiner, NarrowNoBreakSpace, OrnateLeftParenthesis, OrnateRightParenthesis}
    Public Shared NonArabicLetters As String() = {ArabicLetterPeh, ArabicLetterTcheh, ArabicLetterVeh, ArabicLetterGaf}
    Public Shared PunctuationSymbols As String() = {ExclamationMark, QuotationMark, FullStop, Comma, ArabicComma, OrnateLeftParenthesis, OrnateRightParenthesis}
    Public Shared ArabicPunctuationSymbols As String() = {ArabicComma, OrnateLeftParenthesis, OrnateRightParenthesis}
    Public Shared WhitespaceSymbols As String() = {Space}
    Public Shared ArabicStopLetters As String() = {ArabicSmallHighLigatureSadWithLamWithAlefMaksura, _
                                                   ArabicSmallHighLigatureQafWithLamWithAlefMaksura, _
                                ArabicSmallHighMeemInitialForm, ArabicSmallHighLamAlef, _
                                ArabicSmallHighJeem, ArabicSmallHighThreeDots, ArabicSmallHighSeen}
    Structure RuleMetadata
        Sub New(NewIndex As Integer, NewLength As Integer, NewType As String)
            Index = NewIndex
            Length = NewLength
            Type = NewType
        End Sub
        Public Index As Integer
        Public Length As Integer
        Public Type As String
        Public Children As RuleMetadata()
    End Structure
    Structure RuleMetadataTranslation
        Public Rule As String
        Public Match As String
        Public Evaluator As String()
    End Structure
    Structure RuleTranslation
        Public Rule As String
        Public Match As String
        Public Evaluator As String
        Public NegativeMatch As String
        Public RuleFunc As RuleFuncs
    End Structure
    Structure ColorRule
        Public Rule As String
        Public Match As String
        Public Color As Color
    End Structure
    Public Shared AlDuriOrthography As RuleTranslation() = {
        New RuleTranslation With {.Rule = "Feh", .Match = ArabicLetterFeh, .Evaluator = ChrW(&H6A2)},
        New RuleTranslation With {.Rule = "ImalaE", .Match = String.Empty, .Evaluator = ChrW(&H65C)}
    }
    Public Shared WarshOrthography As RuleTranslation() = { _
        New RuleTranslation With {.Rule = "FehBeginMiddle", .Match = MakeUniRegEx(ArabicLetterFeh) + "(" + MakeRegMultiEx(Array.ConvertAll(RecitationCombiningSymbols, Function(C As Char) MakeUniRegEx(C))) + ")*(?!(" + MakeRegMultiEx(Array.ConvertAll(RecitationConnectingFollowerSymbols, Function(C As Char) MakeUniRegEx(C))) + "))", .Evaluator = ChrW(&H6A1)},
        New RuleTranslation With {.Rule = "FehIsolatedEnd", .Match = MakeUniRegEx(ArabicLetterFeh) + "(" + MakeRegMultiEx(Array.ConvertAll(RecitationCombiningSymbols, Function(C As Char) MakeUniRegEx(C))) + ")*(" + MakeRegMultiEx(Array.ConvertAll(RecitationConnectingFollowerSymbols, Function(C As Char) MakeUniRegEx(C))) + ")", .Evaluator = ChrW(&H6A2)},
        New RuleTranslation With {.Rule = "QafBeginMiddle", .Match = MakeUniRegEx(ArabicLetterQaf) + "(" + MakeRegMultiEx(Array.ConvertAll(RecitationCombiningSymbols, Function(C As Char) MakeUniRegEx(C))) + ")*(?!(" + MakeRegMultiEx(Array.ConvertAll(RecitationConnectingFollowerSymbols, Function(C As Char) MakeUniRegEx(C))) + "))", .Evaluator = ChrW(&H66F)},
        New RuleTranslation With {.Rule = "QafIsolatedEnd", .Match = MakeUniRegEx(ArabicLetterQaf) + "(" + MakeRegMultiEx(Array.ConvertAll(RecitationCombiningSymbols, Function(C As Char) MakeUniRegEx(C))) + ")*(" + MakeRegMultiEx(Array.ConvertAll(RecitationConnectingFollowerSymbols, Function(C As Char) MakeUniRegEx(C))) + ")", .Evaluator = ChrW(&H6A7)},
        New RuleTranslation With {.Rule = "Kaf", .Match = MakeUniRegEx(ArabicLetterKaf), .Evaluator = ChrW(&H6A9)},
        New RuleTranslation With {.Rule = "NoonBeginMiddle", .Match = MakeUniRegEx(ArabicLetterNoon) + "(" + MakeRegMultiEx(Array.ConvertAll(RecitationCombiningSymbols, Function(C As Char) MakeUniRegEx(C))) + ")*(?!(" + MakeRegMultiEx(Array.ConvertAll(RecitationConnectingFollowerSymbols, Function(C As Char) MakeUniRegEx(C))) + "))", .Evaluator = ChrW(&H6BA)},
        New RuleTranslation With {.Rule = "NoonIsolatedEnd", .Match = MakeUniRegEx(ArabicLetterNoon) + "(" + MakeRegMultiEx(Array.ConvertAll(RecitationCombiningSymbols, Function(C As Char) MakeUniRegEx(C))) + ")*(" + MakeRegMultiEx(Array.ConvertAll(RecitationConnectingFollowerSymbols, Function(C As Char) MakeUniRegEx(C))) + ")", .Evaluator = ChrW(&H646)},
        New RuleTranslation With {.Rule = "ImalaE", .Match = String.Empty, .Evaluator = ChrW(&H65C)},
        New RuleTranslation With {.Rule = "IIFinal", .Match = "(" + MakeUniRegEx(ArabicKasra) + ")(?=" + MakeUniRegEx(ArabicLetterYeh) + "(^\s*|\s+))", .Evaluator = "$1" + ChrW(&H6D2)}
    }
    Public Shared SimpleTrailingAlef As String = ArabicLetterAlef + ArabicSmallHighRoundedZero
    Public Shared SimpleSuperscriptAlef As String = ArabicLetterSuperscriptAlef
    Public Shared UthmaniShortVowelsBeforeLongVowelsSuperscriptAlef As String = ArabicFatha + ArabicLetterSuperscriptAlef
    Public Shared UthmaniShortVowelsBeforeLongVowelsAlefAlefMaksura As String = ArabicFatha + ArabicLetterAlefMaksura
    Public Shared UthmaniShortVowelsBeforeLongVowelsYehAlefMaksura As String = ArabicKasra + ArabicLetterAlefMaksura
    Public Shared UthmaniShortVowelsBeforeLongVowelsAlef As String = ArabicFatha + ArabicLetterAlef
    Public Shared UthmaniShortVowelsBeforeLongVowelsWaw As String = ArabicDamma + ArabicLetterWaw
    Public Shared UthmaniShortVowelsBeforeLongVowelsSmallWaw As String = ArabicDamma + ArabicSmallWaw
    Public Shared UthmaniShortVowelsBeforeLongVowelsYeh As String = ArabicKasra + ArabicLetterYeh
    Public Shared UthmaniShortVowelsBeforeLongVowelsSmallYeh As String = ArabicKasra + ArabicSmallYeh
    Public Shared UthmaniMinimalScript As RuleTranslation() = { _
        New RuleTranslation With {.Rule = "SmallYehSmallWawAfterPronounHeh", .Match = "(?:(" + MakeUniRegEx(ArabicLetterHeh) + MakeUniRegEx(ArabicKasra) + ")" + MakeUniRegEx(ArabicSmallYeh) + "|(" + MakeUniRegEx(ArabicLetterHeh) + MakeUniRegEx(ArabicDamma) + ")" + MakeUniRegEx(ArabicSmallWaw) + ")" + MakeUniRegEx(ArabicMaddahAbove) + "?(?=\s*$|\s+)", _
            .Evaluator = "$1$2"}, _
        New RuleTranslation With {.Rule = "AlefMaksuraToYeh", .Match = "(" + MakeUniRegEx(ArabicFatha) + ")" + MakeUniRegEx(ArabicLetterAlefMaksura) + "(?=" + MakeUniRegEx(ArabicSukun) + MakeUniRegEx(ArabicLetterHamza) + ")|(" + MakeUniRegEx(ArabicKasra) + ")" + MakeUniRegEx(ArabicLetterAlefMaksura) + "(?=" + MakeUniRegEx(ArabicSmallHighRoundedZero) + "|" + MakeUniRegEx(ArabicShadda) + MakeUniRegEx(ArabicKasra) + MakeUniRegEx(ArabicSmallYeh) + "|" + MakeUniRegEx(ArabicMaddahAbove) + MakeUniRegEx(ArabicLetterHamza) + ")|(" + MakeUniRegEx(ArabicSukun) + ")" + MakeUniRegEx(ArabicLetterAlefMaksura) + "(?=" + MakeUniRegEx(ArabicKasra) + MakeUniRegEx(ArabicSmallYeh) + ")|(" + MakeUniRegEx(ArabicSmallHighRoundedZero) + ")" + MakeUniRegEx(ArabicLetterAlefMaksura) + "(?=" + MakeUniRegEx(ArabicMaddahAbove) + "|" + MakeUniRegEx(ArabicSukun) + MakeUniRegEx(ArabicLetterHamza) + ")", _
            .Evaluator = "$1$2$3$4" + CStr(ArabicLetterYeh)}, _
        New RuleTranslation With {.Rule = "ShortVowelsBeforeLongVowels", .Match = "(?:(?:" + MakeUniRegEx(ArabicFatha) + "(?=" + MakeUniRegEx(ArabicLetterWaw) + MakeUniRegEx(ArabicLetterSuperscriptAlef) + "(?!(?:" + MakeUniRegEx(ArabicShadda) + "|" + MakeRegMultiEx(Array.ConvertAll(ArabicFathaDammaKasra, Function(Str As String) MakeUniRegEx(Str))) + "|" + MakeRegMultiEx(Array.ConvertAll(ArabicTanweens, Function(Str As String) MakeUniRegEx(Str))) + "|(?:" + MakeUniRegEx(ArabicLetterAlef) + MakeUniRegEx(ArabicSmallHighRoundedZero) + ")?(?: (?:" + MakeUniRegEx(ArabicSmallHighLamAlef) + "|" + MakeUniRegEx(ArabicSmallHighJeem) + "|" + MakeUniRegEx(ArabicSmallHighLigatureSadWithLamWithAlefMaksura) + "))? " + MakeUniRegEx(ArabicLetterAlefWasla) + "|" + MakeUniRegEx(ArabicSmallHighRoundedZero) + "|" + MakeUniRegEx(ArabicRoundedHighStopWithFilledCentre) + "|" + MakeUniRegEx(ArabicSukun) + ")))|" + MakeUniRegEx(ArabicFatha) + "(?=" + MakeUniRegEx(ArabicLetterSuperscriptAlef) + "(?!$)(?!(?:" + MakeUniRegEx(ArabicShadda) + "|" + MakeRegMultiEx(Array.ConvertAll(ArabicFathaDammaKasra, Function(Str As String) MakeUniRegEx(Str))) + "|" + MakeRegMultiEx(Array.ConvertAll(ArabicTanweens, Function(Str As String) MakeUniRegEx(Str))) + "|(?:" + MakeUniRegEx(ArabicLetterAlef) + MakeUniRegEx(ArabicSmallHighRoundedZero) + ")?(?: (?:" + MakeUniRegEx(ArabicSmallHighLamAlef) + "|" + MakeUniRegEx(ArabicSmallHighJeem) + "|" + MakeUniRegEx(ArabicSmallHighLigatureSadWithLamWithAlefMaksura) + "))? " + MakeUniRegEx(ArabicLetterAlefWasla) + "|" + MakeUniRegEx(ArabicSmallHighRoundedZero) + "|" + MakeUniRegEx(ArabicRoundedHighStopWithFilledCentre) + "|" + MakeUniRegEx(ArabicSukun) + ")))|" + MakeUniRegEx(ArabicDamma) + "(?=" + MakeUniRegEx(ArabicSmallWaw) + "(?!(?:" + MakeUniRegEx(ArabicShadda) + "|" + MakeRegMultiEx(Array.ConvertAll(ArabicFathaDammaKasra, Function(Str As String) MakeUniRegEx(Str))) + "|" + MakeRegMultiEx(Array.ConvertAll(ArabicTanweens, Function(Str As String) MakeUniRegEx(Str))) + "|(?:" + MakeUniRegEx(ArabicLetterAlef) + MakeUniRegEx(ArabicSmallHighRoundedZero) + ")?(?: (?:" + MakeUniRegEx(ArabicSmallHighLamAlef) + "|" + MakeUniRegEx(ArabicSmallHighJeem) + "|" + MakeUniRegEx(ArabicSmallHighLigatureSadWithLamWithAlefMaksura) + "))? " + MakeUniRegEx(ArabicLetterAlefWasla) + "|" + MakeUniRegEx(ArabicSmallHighRoundedZero) + "|" + MakeUniRegEx(ArabicRoundedHighStopWithFilledCentre) + "|" + MakeUniRegEx(ArabicSukun) + ")))|" + MakeUniRegEx(ArabicFatha) + "(?=" + MakeUniRegEx(ArabicLetterSuperscriptAlef) + "(?!$)(?!(?:" + MakeUniRegEx(ArabicShadda) + "|" + MakeRegMultiEx(Array.ConvertAll(ArabicFathaDammaKasra, Function(Str As String) MakeUniRegEx(Str))) + "|" + MakeRegMultiEx(Array.ConvertAll(ArabicTanweens, Function(Str As String) MakeUniRegEx(Str))) + "|(?:" + MakeUniRegEx(ArabicLetterAlef) + MakeUniRegEx(ArabicSmallHighRoundedZero) + ")?(?: (?:" + MakeUniRegEx(ArabicSmallHighLamAlef) + "|" + MakeUniRegEx(ArabicSmallHighJeem) + "|" + MakeUniRegEx(ArabicSmallHighLigatureSadWithLamWithAlefMaksura) + "))? " + MakeUniRegEx(ArabicLetterAlefWasla) + "|" + MakeUniRegEx(ArabicSmallHighRoundedZero) + "|" + MakeUniRegEx(ArabicRoundedHighStopWithFilledCentre) + "|" + MakeUniRegEx(ArabicSukun) + ")))|" + MakeUniRegEx(ArabicKasra) + "(?=" + MakeUniRegEx(ArabicSmallYeh) + "(?!(?:" + MakeUniRegEx(ArabicShadda) + "|" + MakeRegMultiEx(Array.ConvertAll(ArabicFathaDammaKasra, Function(Str As String) MakeUniRegEx(Str))) + "|" + MakeRegMultiEx(Array.ConvertAll(ArabicTanweens, Function(Str As String) MakeUniRegEx(Str))) + "|(?:" + MakeUniRegEx(ArabicLetterAlef) + MakeUniRegEx(ArabicSmallHighRoundedZero) + ")?(?: (?:" + MakeUniRegEx(ArabicSmallHighLamAlef) + "|" + MakeUniRegEx(ArabicSmallHighJeem) + "|" + MakeUniRegEx(ArabicSmallHighLigatureSadWithLamWithAlefMaksura) + "))? " + MakeUniRegEx(ArabicLetterAlefWasla) + "|" + MakeUniRegEx(ArabicSmallHighRoundedZero) + "|" + MakeUniRegEx(ArabicRoundedHighStopWithFilledCentre) + "|" + MakeUniRegEx(ArabicSukun) + ")))|" + MakeUniRegEx(ArabicFatha) + "(?=" + MakeUniRegEx(ArabicLetterSuperscriptAlef) + "(?!$)(?!(?:" + MakeUniRegEx(ArabicShadda) + "|" + MakeRegMultiEx(Array.ConvertAll(ArabicFathaDammaKasra, Function(Str As String) MakeUniRegEx(Str))) + "|" + MakeRegMultiEx(Array.ConvertAll(ArabicTanweens, Function(Str As String) MakeUniRegEx(Str))) + "|(?:" + MakeUniRegEx(ArabicLetterAlef) + MakeUniRegEx(ArabicSmallHighRoundedZero) + ")?(?: (?:" + MakeUniRegEx(ArabicSmallHighLamAlef) + "|" + MakeUniRegEx(ArabicSmallHighJeem) + "|" + MakeUniRegEx(ArabicSmallHighLigatureSadWithLamWithAlefMaksura) + "))? " + MakeUniRegEx(ArabicLetterAlefWasla) + "|" + MakeUniRegEx(ArabicSmallHighRoundedZero) + "|" + MakeUniRegEx(ArabicRoundedHighStopWithFilledCentre) + "|" + MakeUniRegEx(ArabicSukun) + ")))|" + MakeRegMultiEx(Array.ConvertAll(ArabicLongVowels, Function(StrV As String) MakeUniRegEx(StrV(0)) + "(?=" + MakeUniRegEx(StrV(1)) + "(?!(?:" + MakeUniRegEx(ArabicShadda) + "|" + MakeRegMultiEx(Array.ConvertAll(ArabicFathaDammaKasra, Function(Str As String) MakeUniRegEx(Str))) + "|" + MakeRegMultiEx(Array.ConvertAll(ArabicTanweens, Function(Str As String) MakeUniRegEx(Str))) + "|(?:" + MakeUniRegEx(ArabicLetterAlef) + MakeUniRegEx(ArabicSmallHighRoundedZero) + ")?(?: (?:" + MakeUniRegEx(ArabicSmallHighLamAlef) + "|" + MakeUniRegEx(ArabicSmallHighJeem) + "|" + MakeUniRegEx(ArabicSmallHighLigatureSadWithLamWithAlefMaksura) + "))? " + MakeUniRegEx(ArabicLetterAlefWasla) + "|" + MakeUniRegEx(ArabicSmallHighRoundedZero) + "|" + MakeUniRegEx(ArabicRoundedHighStopWithFilledCentre) + "|" + MakeUniRegEx(ArabicSukun) + ")))|" + MakeUniRegEx(ArabicFatha) + "(?=" + MakeUniRegEx(ArabicLetterSuperscriptAlef) + "(?!$)(?!(?:" + MakeUniRegEx(ArabicShadda) + "|" + MakeRegMultiEx(Array.ConvertAll(ArabicFathaDammaKasra, Function(Str As String) MakeUniRegEx(Str))) + "|" + MakeRegMultiEx(Array.ConvertAll(ArabicTanweens, Function(Str As String) MakeUniRegEx(Str))) + "|(?:" + MakeUniRegEx(ArabicLetterAlef) + MakeUniRegEx(ArabicSmallHighRoundedZero) + ")?(?: (?:" + MakeUniRegEx(ArabicSmallHighLamAlef) + "|" + MakeUniRegEx(ArabicSmallHighJeem) + "|" + MakeUniRegEx(ArabicSmallHighLigatureSadWithLamWithAlefMaksura) + "))? " + MakeUniRegEx(ArabicLetterAlefWasla) + "|" + MakeUniRegEx(ArabicSmallHighRoundedZero) + "|" + MakeUniRegEx(ArabicRoundedHighStopWithFilledCentre) + "|" + MakeUniRegEx(ArabicSukun) + ")))")) + ")|(?:(?:" + MakeUniRegEx(ArabicFatha) + "(?=" + MakeUniRegEx(ArabicLetterAlefMaksura) + "(?!$)(?!(?:" + MakeUniRegEx(ArabicShadda) + "|" + MakeRegMultiEx(Array.ConvertAll(ArabicFathaDammaKasra, Function(Str As String) MakeUniRegEx(Str))) + "|" + MakeRegMultiEx(Array.ConvertAll(ArabicTanweens, Function(Str As String) MakeUniRegEx(Str))) + "|(?:" + MakeUniRegEx(ArabicLetterAlef) + MakeUniRegEx(ArabicSmallHighRoundedZero) + ")?(?: (?:" + MakeUniRegEx(ArabicSmallHighLamAlef) + "|" + MakeUniRegEx(ArabicSmallHighJeem) + "|" + MakeUniRegEx(ArabicSmallHighLigatureSadWithLamWithAlefMaksura) + "))? " + MakeUniRegEx(ArabicLetterAlefWasla) + "|" + MakeUniRegEx(ArabicSmallHighRoundedZero) + "|" + MakeUniRegEx(ArabicRoundedHighStopWithFilledCentre) + "|" + MakeUniRegEx(ArabicSukun) + ")))|" + MakeUniRegEx(ArabicFatha) + "(?=" + MakeUniRegEx(ArabicLetterSuperscriptAlef) + "(?!$)(?!(?:" + MakeUniRegEx(ArabicShadda) + "|" + MakeRegMultiEx(Array.ConvertAll(ArabicFathaDammaKasra, Function(Str As String) MakeUniRegEx(Str))) + "|" + MakeRegMultiEx(Array.ConvertAll(ArabicTanweens, Function(Str As String) MakeUniRegEx(Str))) + "|(?:" + MakeUniRegEx(ArabicLetterAlef) + MakeUniRegEx(ArabicSmallHighRoundedZero) + ")?(?: (?:" + MakeUniRegEx(ArabicSmallHighLamAlef) + "|" + MakeUniRegEx(ArabicSmallHighJeem) + "|" + MakeUniRegEx(ArabicSmallHighLigatureSadWithLamWithAlefMaksura) + "))? " + MakeUniRegEx(ArabicLetterAlefWasla) + "|" + MakeUniRegEx(ArabicSmallHighRoundedZero) + "|" + MakeUniRegEx(ArabicRoundedHighStopWithFilledCentre) + "|" + MakeUniRegEx(ArabicSukun) + ")))))|" + MakeUniRegEx(ArabicKasra) + "(?=" + MakeUniRegEx(ArabicLetterAlefMaksura) + "(?!$)(?!(?:" + MakeUniRegEx(ArabicShadda) + "|" + MakeRegMultiEx(Array.ConvertAll(ArabicFathaDammaKasra, Function(Str As String) MakeUniRegEx(Str))) + "|" + MakeRegMultiEx(Array.ConvertAll(ArabicTanweens, Function(Str As String) MakeUniRegEx(Str))) + "|(?:" + MakeUniRegEx(ArabicLetterAlef) + MakeUniRegEx(ArabicSmallHighRoundedZero) + ")?(?: (?:" + MakeUniRegEx(ArabicSmallHighLamAlef) + "|" + MakeUniRegEx(ArabicSmallHighJeem) + "|" + MakeUniRegEx(ArabicSmallHighLigatureSadWithLamWithAlefMaksura) + "))? " + MakeUniRegEx(ArabicLetterAlefWasla) + "|" + MakeUniRegEx(ArabicSmallHighRoundedZero) + "|" + MakeUniRegEx(ArabicRoundedHighStopWithFilledCentre) + "|" + MakeUniRegEx(ArabicSukun) + "))))", _
            .Evaluator = String.Empty}, _
        New RuleTranslation With {.Rule = "ShortVowelsBeforeLongVowels", .Match = "((?:" + MakeUniRegEx(ArabicLetterKhah) + "|" + MakeUniRegEx(ArabicLetterKaf) + MakeUniRegEx(ArabicSukun) + MakeUniRegEx(ArabicLetterReh) + "|" + MakeUniRegEx(ArabicKasra) + MakeUniRegEx(ArabicLetterNoon) + MakeUniRegEx(ArabicFatha) + MakeUniRegEx(ArabicLetterFeh) + MakeUniRegEx(ArabicSukun) + MakeUniRegEx(ArabicLetterSeen) + ")" + MakeUniRegEx(ArabicKasra) + "(?=" + MakeUniRegEx(ArabicLetterAlefMaksura) + "$))|" + MakeUniRegEx(ArabicKasra) + "(?=" + MakeUniRegEx(ArabicLetterAlefMaksura) + "$)", _
            .Evaluator = "$1"}, _
        New RuleTranslation With {.Rule = "ShortVowelsBeforeLongVowels", .Match = "(" + MakeUniRegEx(ArabicLetterAlefWithHamzaAbove) + ")(?:" + MakeUniRegEx(ArabicFatha) + "(?=(?:" + MakeUniRegEx(ArabicLetterAlef) + "|" + MakeUniRegEx(ArabicLetterSuperscriptAlef) + "|" + MakeUniRegEx(ArabicLetterAlefMaksura) + MakeUniRegEx(ArabicLetterSuperscriptAlef) + ")(?!" + MakeUniRegEx(ArabicLetterSuperscriptAlef) + "|" + MakeUniRegEx(ArabicSmallHighRoundedZero) + "))|" + MakeUniRegEx(ArabicDamma) + "(?=(?:" + MakeUniRegEx(ArabicLetterWaw) + "|" + MakeUniRegEx(ArabicSmallWaw) + ")(?!" + MakeUniRegEx(ArabicLetterSuperscriptAlef) + "|" + MakeUniRegEx(ArabicSmallHighRoundedZero) + ")))|(" + MakeUniRegEx(ArabicLetterAlefWithHamzaBelow) + ")" + MakeUniRegEx(ArabicKasra) + "(?=(" + MakeUniRegEx(ArabicLetterYeh) + "|" + MakeUniRegEx(ArabicSmallYeh) + "|" + MakeUniRegEx(ArabicLetterAlefMaksura) + ")(?!" + MakeUniRegEx(ArabicShadda) + "|" + MakeUniRegEx(ArabicLetterAlef) + "|" + MakeUniRegEx(ArabicSmallHighRoundedZero) + "))", _
            .Evaluator = "$1$2"}, _
        New RuleTranslation With {.Rule = "SmallSadAboveHamzaWasl", .Match = MakeUniRegEx(ArabicLetterAlefWasla), _
            .Evaluator = CStr(ArabicLetterAlef)}, _
        New RuleTranslation With {.Rule = "Sukun", .Match = MakeUniRegEx(ArabicSukun), _
            .Evaluator = String.Empty}, _
        New RuleTranslation With {.Rule = "Madda", .Match = MakeUniRegEx(ArabicLetterAlefWithMaddaAbove), _
            .Evaluator = CStr(ArabicLetterAlef)}, _
        New RuleTranslation With {.Rule = "Madda", .Match = MakeUniRegEx(ArabicMaddahAbove), _
            .Evaluator = String.Empty}, _
        New RuleTranslation With {.Rule = "SmallCircleShowingNonReadLetters", .Match = "(" + MakeUniRegEx(ArabicLetterWaw) + "(?:" + MakeUniRegEx(ArabicDamma) + "?" + MakeUniRegEx(ArabicSmallWaw) + "?)" + MakeUniRegEx(ArabicLetterAlef) + ")" + MakeUniRegEx(ArabicSmallHighRoundedZero) + "(?=\s*$|\s+)|(" + MakeUniRegEx(ArabicLetterAlefWithHamzaAbove) + MakeUniRegEx(ArabicDamma) + MakeUniRegEx(ArabicLetterWaw) + ")" + MakeUniRegEx(ArabicSmallHighRoundedZero) + "(?=" + MakeUniRegEx(ArabicLetterLam) + ")", _
            .Evaluator = "$1$2"}, _
        New RuleTranslation With {.Rule = "ShaddaIdgham", .Match = "((?:\s+|^\s*)(?:" + MakeRegMultiEx(Array.ConvertAll(ArabicSunLetters, Function(Str As String) MakeUniRegEx(Str))) + "|" + MakeUniRegEx(ArabicLetterWaw) + "|" + MakeUniRegEx(ArabicLetterMeem) + "|" + MakeUniRegEx(ArabicLetterYeh) + "|" + MakeUniRegEx(ArabicLetterBeh) + "|" + MakeUniRegEx(ArabicLetterAin) + "|" + MakeUniRegEx(ArabicLetterFeh) + "))" + MakeUniRegEx(ArabicShadda), _
            .Evaluator = "$1"}, _
        New RuleTranslation With {.Rule = "ShaddaIdgham", .Match = "(" + MakeUniRegEx(ArabicLetterYeh) + MakeUniRegEx(ArabicLetterYeh) + "|" + MakeUniRegEx(ArabicLetterDal) + MakeUniRegEx(ArabicLetterTeh) + "|" + MakeUniRegEx(ArabicLetterHeh) + "|(?:" + MakeUniRegEx(ArabicKasra) + "|" + MakeUniRegEx(ArabicLetterKaf) + "|" + MakeUniRegEx(ArabicLetterQaf) + ")" + MakeUniRegEx(ArabicLetterKaf) + ")" + MakeUniRegEx(ArabicShadda) + "(?=(" + MakeUniRegEx(ArabicDamma) + "|" + MakeUniRegEx(ArabicFatha) + "|" + MakeUniRegEx(ArabicKasra) + ")(" + MakeUniRegEx(ArabicLetterHeh) + "|" + MakeUniRegEx(ArabicLetterKaf) + ")?((" + MakeUniRegEx(ArabicDamma) + "|" + MakeUniRegEx(ArabicFatha) + ")?(" + MakeUniRegEx(ArabicSmallWaw) + "|(" + MakeUniRegEx(ArabicLetterMeem) + "(" + MakeUniRegEx(ArabicDamma) + "|" + MakeUniRegEx(ArabicLetterWaw) + MakeUniRegEx(ArabicLetterHeh) + MakeUniRegEx(ArabicDamma) + MakeUniRegEx(ArabicLetterMeem) + ")?)|" + MakeUniRegEx(ArabicLetterNoon) + "(" + MakeUniRegEx(ArabicShadda) + MakeUniRegEx(ArabicFatha) + "|" + MakeUniRegEx(ArabicLetterAlef) + ")|" + MakeUniRegEx(ArabicLetterAlef) + ")?(\s*$|\s+)))", _
            .Evaluator = "$1"}, _
        New RuleTranslation With {.Rule = "SmallMeemIghlab", .Match = MakeUniRegEx(ArabicSmallHighMeemIsolatedForm) + "|" + MakeUniRegEx(ArabicSmallLowMeem), _
            .Evaluator = String.Empty}, _
        New RuleTranslation With {.Rule = "SuperscriptAlefHamzaAbove", .Match = MakeUniRegEx(ArabicLetterSuperscriptAlef) + MakeUniRegEx(ArabicHamzaAbove), _
            .Evaluator = CStr(ArabicFatha) + CStr(ArabicLetterHamza)},
        New RuleTranslation With {.Rule = "AlefRoundedHighStopWithFilledCentre", .Match = MakeUniRegEx(ArabicLetterAlef) + MakeUniRegEx(ArabicRoundedHighStopWithFilledCentre), _
            .Evaluator = CStr(ArabicLetterAlefWithHamzaAbove) + CStr(ArabicFatha)}
    }
    Public Shared SimpleEnhancedScript As RuleTranslation() = { _
        New RuleTranslation With {.Rule = "AlefRoundedHighStopWithFilledCentre", .Match = MakeUniRegEx(ArabicLetterAlef) + MakeUniRegEx(ArabicRoundedHighStopWithFilledCentre), _
            .Evaluator = CStr(ArabicLetterAlefWithHamzaAbove) + CStr(ArabicFatha)}, _
        New RuleTranslation With {.Rule = "EmptyCenterLowStop", .Match = MakeUniRegEx(ArabicEmptyCentreLowStop) + MakeUniRegEx(ArabicLetterAlefMaksura) + MakeUniRegEx(ArabicLetterSuperscriptAlef), _
            .Evaluator = CStr(ArabicFatha) + CStr(ArabicLetterAlef)}, _
        New RuleTranslation With {.Rule = "SmallYehSmallWawAfterPronounHeh", .Match = "(?:(" + MakeUniRegEx(ArabicLetterHeh) + MakeUniRegEx(ArabicKasra) + ")" + MakeUniRegEx(ArabicSmallYeh) + "|(" + MakeUniRegEx(ArabicLetterHeh) + MakeUniRegEx(ArabicDamma) + ")" + MakeUniRegEx(ArabicSmallWaw) + ")" + MakeUniRegEx(ArabicMaddahAbove) + "?(?=\s*$|\s+)", _
            .Evaluator = "$1$2"}, _
        New RuleTranslation With {.Rule = "SmallWaw", .Match = MakeUniRegEx(ArabicSmallWaw), _
            .Evaluator = CStr(ArabicLetterWaw)}, _
        New RuleTranslation With {.Rule = "SmallYeh", .Match = MakeUniRegEx(ArabicSmallYeh), _
            .Evaluator = CStr(ArabicLetterYeh)}, _
        New RuleTranslation With {.Rule = "TrailingAlef", .Match = "((?:(?:(" + MakeUniRegEx(ArabicLetterAlefWithHamzaAbove) + "|" + MakeUniRegEx(ArabicLetterWawWithHamzaAbove) + "|" + MakeUniRegEx(ArabicLetterWaw) + "|(?:(?:" + MakeUniRegEx(ArabicLetterZain) + "| (?:" + MakeUniRegEx(ArabicLetterWaw) + MakeUniRegEx(ArabicFatha) + ")?" + MakeUniRegEx(ArabicLetterTeh) + ")" + MakeUniRegEx(ArabicFatha) + "|" + MakeUniRegEx(ArabicLetterTeh) + MakeUniRegEx(ArabicDamma) + "|" + MakeUniRegEx(ArabicLetterAlefWasla) + ")" + MakeUniRegEx(ArabicLetterDal) + "|" + MakeUniRegEx(ArabicLetterKhah) + "|" + MakeUniRegEx(ArabicLetterHeh) + "|" + MakeUniRegEx(ArabicLetterQaf) + "|" + MakeUniRegEx(ArabicLetterYeh) + "|(?:" + MakeUniRegEx(ArabicLetterAlefWasla) + "|" + MakeUniRegEx(ArabicDamma) + ")" + MakeUniRegEx(ArabicLetterBeh) + "|(?:" + MakeUniRegEx(ArabicLetterAlefWasla) + "|" + MakeUniRegEx(ArabicLetterAlefWithHamzaAbove) + MakeUniRegEx(ArabicDamma) + ")" + MakeUniRegEx(ArabicLetterReh) + "|(?:(?:" + MakeUniRegEx(ArabicLetterTeh) + "|" + MakeUniRegEx(ArabicSukun) + MakeUniRegEx(ArabicLetterYeh) + ")" + MakeUniRegEx(ArabicFatha) + "|" + MakeUniRegEx(ArabicLetterAlefWasla) + ")" + MakeUniRegEx(ArabicLetterAin) + "|" + MakeUniRegEx(ArabicLetterGhain) + "|(?:" + MakeUniRegEx(ArabicLetterTeh) + MakeUniRegEx(ArabicDamma) + "|" + MakeUniRegEx(ArabicSukun) + MakeUniRegEx(ArabicLetterYeh) + MakeUniRegEx(ArabicFatha) + ")" + MakeUniRegEx(ArabicLetterBeh) + "|" + MakeUniRegEx(ArabicLetterLam) + "|" + MakeUniRegEx(ArabicLetterNoon) + "|" + MakeUniRegEx(ArabicLetterYehWithHamzaAbove) + ")?" + MakeUniRegEx(ArabicSukun) + "|" + MakeUniRegEx(ArabicSmallHighRoundedZero) + "|((?:" + MakeUniRegEx(ArabicLetterDal) + "|" + MakeUniRegEx(ArabicLetterAin) + "|" + MakeUniRegEx(ArabicLetterReh) + "|" + MakeUniRegEx(ArabicLetterTeh) + "|" + MakeUniRegEx(ArabicLetterHeh) + "|" + MakeUniRegEx(ArabicLetterThal) + "|" + MakeUniRegEx(ArabicLetterJeem) + "|" + MakeUniRegEx(ArabicLetterKhah) + "|" + MakeUniRegEx(ArabicLetterKaf) + "|" + MakeUniRegEx(ArabicLetterQaf) + "|" + MakeUniRegEx(ArabicLetterHamza) + "|" + MakeUniRegEx(ArabicLetterBeh) + "|" + MakeUniRegEx(ArabicLetterNoon) + "|" + MakeUniRegEx(ArabicLetterTah) + "|" + MakeUniRegEx(ArabicLetterDad) + "|" + MakeUniRegEx(ArabicLetterMeem) + "|" + MakeUniRegEx(ArabicLetterLam) + "|" + MakeUniRegEx(ArabicLetterWaw) + MakeUniRegEx(ArabicShadda) + "|" + MakeUniRegEx(ArabicTatweel) + MakeUniRegEx(ArabicHamzaAbove) + ")" + MakeUniRegEx(ArabicFatha) + ")?" + MakeUniRegEx(ArabicLetterAlef) + "|(" + MakeUniRegEx(ArabicTatweel) + "|(?:" + MakeUniRegEx(ArabicLetterQaf) + "|" + MakeUniRegEx(ArabicLetterYeh) + MakeUniRegEx(ArabicDamma) + MakeUniRegEx(ArabicLetterLam) + "|" + MakeUniRegEx(ArabicLetterTah) + ")" + MakeUniRegEx(ArabicFatha) + ")?" + MakeUniRegEx(ArabicLetterSuperscriptAlef) + ")(?:" + MakeUniRegEx(ArabicLetterAin) + "|" + MakeUniRegEx(ArabicLetterSeen) + "|" + MakeUniRegEx(ArabicLetterHah) + "|" + MakeUniRegEx(ArabicSukun) + MakeUniRegEx(ArabicLetterReh) + MakeUniRegEx(ArabicKasra) + MakeUniRegEx(ArabicLetterKaf) + "|(" + MakeUniRegEx(ArabicFatha) + ")?" + MakeUniRegEx(ArabicLetterFeh) + "|" + MakeUniRegEx(ArabicLetterNoon) + "|" + MakeUniRegEx(ArabicLetterMeem) + "|" + MakeUniRegEx(ArabicLetterQaf) + "|" + MakeUniRegEx(ArabicLetterDal) + "|" + MakeUniRegEx(ArabicLetterHamza) + "|" + MakeUniRegEx(ArabicLetterReh) + MakeUniRegEx(ArabicShadda) + "|((?:" + MakeUniRegEx(ArabicLetterAin) + "|" + MakeUniRegEx(ArabicLetterKhah) + "|" + MakeUniRegEx(ArabicLetterBeh) + "|" + MakeUniRegEx(ArabicLetterMeem) + "|" + MakeUniRegEx(ArabicLetterTeh) + "|" + MakeUniRegEx(ArabicLetterKaf) + MakeUniRegEx(ArabicShadda) + "|" + MakeUniRegEx(ArabicLetterYeh) + MakeUniRegEx(ArabicShadda) + "|" + MakeUniRegEx(ArabicLetterAlefWithHamzaAbove) + "|" + MakeUniRegEx(ArabicTatweel) + MakeUniRegEx(ArabicHamzaAbove) + ")" + MakeUniRegEx(ArabicFatha) + "|(?:" + MakeUniRegEx(ArabicLetterMeem) + "|" + MakeUniRegEx(ArabicLetterTah) + "|" + MakeUniRegEx(ArabicLetterTeh) + "|" + MakeUniRegEx(ArabicLetterKhah) + "|" + MakeUniRegEx(ArabicLetterSad) + "|" + MakeUniRegEx(ArabicLetterDal) + "|" + MakeUniRegEx(ArabicLetterMeem) + MakeUniRegEx(ArabicShadda) + "|" + MakeUniRegEx(ArabicLetterDal) + MakeUniRegEx(ArabicShadda) + "|(?:(?:" + MakeUniRegEx(ArabicLetterAlefWithHamzaAbove) + "(?:" + MakeUniRegEx(ArabicDamma) + "|" + MakeUniRegEx(ArabicFatha) + ")" + MakeUniRegEx(ArabicLetterReh) + "|" + MakeUniRegEx(ArabicLetterBeh) + "|" + MakeUniRegEx(ArabicLetterGhain) + ")" + MakeUniRegEx(ArabicSukun) + ")" + MakeUniRegEx(ArabicLetterSeen) + ")" + MakeUniRegEx(ArabicKasra) + "|(?:" + MakeUniRegEx(ArabicLetterKaf) + "|" + MakeUniRegEx(ArabicLetterKhah) + "|" + MakeUniRegEx(ArabicLetterTeh) + ")" + MakeUniRegEx(ArabicDamma) + "|" + MakeUniRegEx(ArabicKasra) + MakeUniRegEx(ArabicLetterYeh) + ")?" + MakeUniRegEx(ArabicLetterLam) + "|" + MakeUniRegEx(ArabicLetterThal) + "|" + MakeUniRegEx(ArabicLetterKaf) + "|" + MakeUniRegEx(ArabicKasra) + MakeUniRegEx(ArabicLetterSeen) + "|" + MakeUniRegEx(ArabicLetterBeh) + "|" + MakeUniRegEx(ArabicLetterTeh) + "|(" + MakeUniRegEx(ArabicKasra) + "|" + MakeUniRegEx(ArabicDamma) + ")?" + MakeUniRegEx(ArabicLetterJeem) + ")|" + MakeUniRegEx(ArabicLetterBeh) + MakeUniRegEx(ArabicFatha) + MakeUniRegEx(ArabicLetterNoon) + "|" + MakeUniRegEx(ArabicLetterAlef) + MakeUniRegEx(ArabicLetterSeen) + MakeUniRegEx(ArabicKasra) + MakeUniRegEx(ArabicLetterTah) + "|" + MakeUniRegEx(ArabicLetterSheen) + MakeUniRegEx(ArabicKasra) + MakeUniRegEx(ArabicLetterFeh) + "|(?:" + MakeUniRegEx(ArabicLetterAlef) + MakeUniRegEx(ArabicLetterReh) + "|" + MakeUniRegEx(ArabicLetterMeem) + MakeUniRegEx(ArabicDamma) + MakeUniRegEx(ArabicLetterHeh) + MakeUniRegEx(ArabicSukun) + MakeUniRegEx(ArabicLetterLam) + ")" + MakeUniRegEx(ArabicKasra) + MakeUniRegEx(ArabicLetterKaf) + "|" + MakeUniRegEx(ArabicLetterYehWithHamzaAbove) + MakeUniRegEx(ArabicKasra) + MakeUniRegEx(ArabicLetterQaf) + "|" + MakeUniRegEx(ArabicLetterMeem) + MakeUniRegEx(ArabicDamma) + MakeUniRegEx(ArabicLetterReh) + MakeUniRegEx(ArabicSukun) + MakeUniRegEx(ArabicLetterSeen) + MakeUniRegEx(ArabicKasra) + MakeUniRegEx(ArabicLetterLam) + "|" + MakeUniRegEx(ArabicLetterAlef) + MakeUniRegEx(ArabicLetterKaf) + MakeUniRegEx(ArabicKasra) + MakeUniRegEx(ArabicLetterSeen) + ")(" + MakeUniRegEx(ArabicShadda) + ")?" + MakeUniRegEx(ArabicDamma) + MakeUniRegEx(ArabicLetterWaw) + MakeUniRegEx(ArabicMaddahAbove) + "?" + MakeUniRegEx(ArabicFatha) + "?)" + MakeUniRegEx(ArabicLetterAlef) + MakeUniRegEx(ArabicSmallHighRoundedZero) + "|" + MakeUniRegEx(ArabicLetterAlef) + MakeUniRegEx(ArabicSmallHighRoundedZero) + "(?=" + MakeUniRegEx(ArabicLetterAlefMaksura) + ")|((?:" + MakeUniRegEx(ArabicLetterLam) + "|" + MakeUniRegEx(ArabicLetterDal) + "|" + MakeUniRegEx(ArabicLetterReh) + "|" + MakeUniRegEx(ArabicLetterTeh) + "|" + MakeUniRegEx(ArabicLetterYeh) + "|" + MakeUniRegEx(ArabicLetterAlefWithHamzaAbove) + ")" + MakeUniRegEx(ArabicFatha) + "|" + MakeUniRegEx(ArabicLetterWawWithHamzaAbove) + "(?:" + MakeUniRegEx(ArabicDamma) + "|" + MakeUniRegEx(ArabicDammatan) + MakeUniRegEx(ArabicSmallLowMeem) + "?))" + MakeUniRegEx(ArabicLetterAlef) + MakeUniRegEx(ArabicSmallHighRoundedZero) + "(?=\s*$|\s+|" + MakeUniRegEx(ArabicLetterThal) + "|" + MakeUniRegEx(ArabicLetterYeh) + ")", _
            .Evaluator = "$1$9", .NegativeMatch = "$2$3$4$5$6$7$8"}, _
        New RuleTranslation With {.Rule = "TrailingAlef", .Match = "(" + MakeUniRegEx(ArabicLetterAlef) + MakeUniRegEx(ArabicMaddahAbove) + MakeUniRegEx(ArabicLetterHamza) + MakeUniRegEx(ArabicDamma) + MakeUniRegEx(ArabicLetterWaw) + MakeUniRegEx(ArabicMaddahAbove) + "?)(?=\s*$|\s+)", _
            .Evaluator = "$1" + CStr(ArabicLetterAlef)}, _
        New RuleTranslation With {.Rule = "AlefMaksuraToYeh", .Match = "(" + MakeUniRegEx(ArabicDamma) + MakeUniRegEx(ArabicLetterHah) + MakeUniRegEx(ArabicSukun) + ")" + MakeUniRegEx(ArabicLetterAlefMaksura) + "(" + MakeUniRegEx(ArabicKasra) + ")(?=\s*$|\s+)", _
            .Evaluator = "$1" + CStr(ArabicLetterYeh) + "$2" + CStr(ArabicLetterYeh)}, _
        New RuleTranslation With {.Rule = "AlefMaksuraToYeh", .Match = "(" + MakeUniRegEx(ArabicSmallHighRoundedZero) + "|" + MakeUniRegEx(ArabicSukun) + "|" + MakeUniRegEx(ArabicKasra) + ")" + MakeUniRegEx(ArabicLetterAlefMaksura) + "|(" + MakeUniRegEx(ArabicLetterAlef) + "|" + MakeUniRegEx(ArabicLetterSuperscriptAlef) + ")" + MakeUniRegEx(ArabicLetterAlefMaksura) + "(?=" + MakeUniRegEx(ArabicFatha) + ")|(" + MakeUniRegEx(ArabicFatha) + ")" + MakeUniRegEx(ArabicLetterAlefMaksura) + "(?=" + MakeUniRegEx(ArabicSukun) + "|" + MakeUniRegEx(ArabicShadda) + "|" + MakeUniRegEx(ArabicKasra) + ")", _
            .Evaluator = "$1$2$3" + CStr(ArabicLetterYeh)}, _
        New RuleTranslation With {.Rule = "SmallSadAboveHamzaWasl", .Match = MakeUniRegEx(ArabicLetterAlefWasla) + "(?=" + MakeUniRegEx(ArabicLetterLam) + MakeUniRegEx(ArabicShadda) + "(?!" + MakeUniRegEx(ArabicFatha) + "(?:" + MakeUniRegEx(ArabicLetterThal) + MakeUniRegEx(ArabicKasra) + "|" + MakeUniRegEx(ArabicLetterTeh) + ")))", _
            .Evaluator = ArabicLetterAlef + ArabicLetterLam}, _
        New RuleTranslation With {.Rule = "SmallSadAboveHamzaWasl", .Match = MakeUniRegEx(ArabicLetterAlefWasla), _
            .Evaluator = CStr(ArabicLetterAlef)}, _
        New RuleTranslation With {.Rule = "Alef", .Match = "((?:^\s*|\s+)" + MakeUniRegEx(ArabicLetterYeh) + MakeUniRegEx(ArabicFatha) + ")(?=" + MakeUniRegEx(ArabicLetterBeh) + MakeUniRegEx(ArabicSukun) + MakeUniRegEx(ArabicLetterNoon) + ")", _
            .Evaluator = "$1" + CStr(ArabicLetterAlef) + " " + CStr(ArabicLetterAlef)}, _
        New RuleTranslation With {.Rule = "Alef", .Match = "(" + MakeUniRegEx(ArabicLetterBeh) + MakeUniRegEx(ArabicSukun) + MakeUniRegEx(ArabicLetterNoon) + MakeUniRegEx(ArabicFatha) + ")" + MakeUniRegEx(ArabicLetterWawWithHamzaAbove) + "(?=" + MakeUniRegEx(ArabicDamma) + ")", _
            .Evaluator = "$1 " + CStr(ArabicLetterAlefWithHamzaAbove)}, _
        New RuleTranslation With {.Rule = "Alef", .Match = "(^\s*|\s+)(?=" + MakeUniRegEx(ArabicLetterLam) + MakeUniRegEx(ArabicSukun) + MakeUniRegEx(ArabicTatweel) + MakeUniRegEx(ArabicHamzaAbove) + ")|((?:^\s*|\s+)" + MakeUniRegEx(ArabicLetterLam) + MakeUniRegEx(ArabicFatha) + ")(?=" + MakeUniRegEx(ArabicLetterTeh) + MakeUniRegEx(ArabicShadda) + ")|((?:^\s*|\s+)(?:" + MakeUniRegEx(ArabicLetterWaw) + "|" + MakeUniRegEx(ArabicLetterFeh) + ")" + MakeUniRegEx(ArabicFatha) + ")(?=" + MakeUniRegEx(ArabicLetterSeen) + MakeUniRegEx(ArabicSukun) + ")|(" + MakeUniRegEx(ArabicSukun) + "(?:" + MakeUniRegEx(ArabicLetterReh) + "|" + MakeUniRegEx(ArabicLetterLam) + ")" + MakeUniRegEx(ArabicFatha) + MakeUniRegEx(ArabicLetterTeh) + MakeUniRegEx(ArabicFatha) + "|" + MakeUniRegEx(ArabicShadda) + MakeUniRegEx(ArabicKasra) + MakeUniRegEx(ArabicLetterNoon) + MakeUniRegEx(ArabicFatha) + ")" + MakeUniRegEx(ArabicLetterAlefMaksura) + MakeUniRegEx(ArabicLetterSuperscriptAlef) + MakeUniRegEx(ArabicMaddahAbove) + "?(?=(?:\s*$|\s+(?!" + MakeUniRegEx(ArabicLetterLam) + "|" + MakeUniRegEx(ArabicLetterHamza) + ")))|(" + MakeUniRegEx(ArabicLetterAin) + "(?:" + MakeUniRegEx(ArabicFatha) + MakeUniRegEx(ArabicLetterTeh) + ")?" + MakeUniRegEx(ArabicFatha) + MakeUniRegEx(ArabicLetterWaw) + MakeUniRegEx(ArabicSukun) + "|" + MakeUniRegEx(ArabicShadda) + MakeUniRegEx(ArabicFatha) + MakeUniRegEx(ArabicLetterHamza) + MakeUniRegEx(ArabicDamma) + MakeUniRegEx(ArabicLetterWaw) + ")(?=\s*$|\s+)|" + MakeUniRegEx(ArabicLetterWaw) + MakeUniRegEx(ArabicLetterSuperscriptAlef) + MakeUniRegEx(ArabicLetterAlef) + MakeUniRegEx(ArabicSmallHighRoundedZero) + "(?=\s*$|\s+)", _
            .Evaluator = "$1$2$3$4$5" + CStr(ArabicLetterAlef)}, _
        New RuleTranslation With {.Rule = "HamzaAlefEnding", .Match = "(" + MakeUniRegEx(ArabicLetterSuperscriptAlef) + MakeUniRegEx(ArabicMaddahAbove) + MakeUniRegEx(ArabicLetterHamza) + ")" + MakeUniRegEx(ArabicFatha) + MakeUniRegEx(ArabicLetterAlef) + "(?= " + MakeUniRegEx(ArabicLetterAlef) + ")", _
            .Evaluator = "$1" + CStr(ArabicFatha) + CStr(ArabicLetterAlefMaksura)}, _
        New RuleTranslation With {.Rule = "HamzaAlefEnding", .Match = "(?:" + MakeUniRegEx(ArabicLetterHamza) + "|" + MakeUniRegEx(ArabicTatweel) + MakeUniRegEx(ArabicHamzaAbove) + ")" + MakeUniRegEx(ArabicFatha) + MakeUniRegEx(ArabicLetterAlef) + "(?= " + MakeUniRegEx(ArabicLetterAlef) + ")", _
            .Evaluator = CStr(ArabicLetterAlefWithHamzaAbove) + CStr(ArabicFatha) + CStr(ArabicLetterAlefMaksura)}, _
        New RuleTranslation With {.Rule = "HamzaAlefEnding", .Match = "(?:(" + MakeUniRegEx(ArabicLetterReh) + MakeUniRegEx(ArabicFatha) + ")" + MakeUniRegEx(ArabicLetterHamza) + "|" + MakeUniRegEx(ArabicTatweel) + MakeUniRegEx(ArabicHamzaAbove) + ")" + MakeUniRegEx(ArabicFatha) + MakeUniRegEx(ArabicLetterAlef) + "(?=(?:\s*$|\s+))", _
            .Evaluator = "$1" + CStr(ArabicLetterAlefWithHamzaAbove) + CStr(ArabicFatha) + CStr(ArabicLetterAlefMaksura) + CStr(ArabicLetterSuperscriptAlef)}, _
        New RuleTranslation With {.Rule = "AlefMaksura", .Match = "((?:" + MakeUniRegEx(ArabicLetterTah) + MakeUniRegEx(ArabicFatha) + MakeUniRegEx(ArabicLetterGhain) + "|" + MakeUniRegEx(ArabicLetterSad) + "|" + MakeUniRegEx(ArabicLetterLam) + MakeUniRegEx(ArabicFatha) + MakeUniRegEx(ArabicLetterDal) + ")" + MakeUniRegEx(ArabicFatha) + ")" + MakeUniRegEx(ArabicLetterAlef) + "(?=(?:\s*$|\s+))", _
            .Evaluator = "$1" + CStr(ArabicLetterAlefMaksura)}, _
        New RuleTranslation With {.Rule = "RehAlefEnding", .Match = "(" + MakeUniRegEx(ArabicLetterTeh) + MakeUniRegEx(ArabicSukun) + MakeUniRegEx(ArabicLetterReh) + "|" + MakeUniRegEx(ArabicLetterReh) + MakeUniRegEx(ArabicShadda) + "?" + MakeUniRegEx(ArabicFatha) + MakeUniRegEx(ArabicLetterHamza) + ")" + MakeUniRegEx(ArabicFatha) + MakeUniRegEx(ArabicLetterAlef) + MakeUniRegEx(ArabicMaddahAbove) + "?(?=(?:\s*$|\s+))", _
            .Evaluator = "$1" + CStr(ArabicFatha) + CStr(ArabicLetterAlefMaksura) + CStr(ArabicLetterSuperscriptAlef)}, _
        New RuleTranslation With {.Rule = "Madda", .Match = MakeUniRegEx(ArabicLetterWaw) + MakeUniRegEx(ArabicLetterSuperscriptAlef) + MakeUniRegEx(ArabicMaddahAbove) + "(?=" + MakeUniRegEx(ArabicLetterAlef) + MakeUniRegEx(ArabicSmallHighRoundedZero) + ")", _
            .Evaluator = String.Empty}, _
        New RuleTranslation With {.Rule = "Madda", .Match = "(?:" + MakeUniRegEx(ArabicLetterHamza) + "|" + MakeUniRegEx(ArabicLetterAlefWithHamzaAbove) + ")" + MakeUniRegEx(ArabicFatha) + "(?:" + MakeUniRegEx(ArabicLetterAlef) + "|" + MakeUniRegEx(ArabicLetterSuperscriptAlef) + ")|(" + MakeUniRegEx(ArabicFatha) + "|" + MakeUniRegEx(ArabicLetterBeh) + MakeUniRegEx(ArabicKasra) + "|" + MakeUniRegEx(ArabicSukun) + "|" + MakeUniRegEx(ArabicLetterSuperscriptAlef) + MakeUniRegEx(ArabicMaddahAbove) + ")" + MakeUniRegEx(ArabicTatweel) + MakeUniRegEx(ArabicHamzaAbove) + MakeUniRegEx(ArabicFatha) + "(?:" + MakeUniRegEx(ArabicLetterAlef) + "|" + MakeUniRegEx(ArabicTatweel) + "?" + MakeUniRegEx(ArabicLetterSuperscriptAlef) + ")", _
            .Evaluator = "$1" + CStr(ArabicLetterAlefWithMaddaAbove)}, _
        New RuleTranslation With {.Rule = "Madda", .Match = MakeUniRegEx(ArabicLetterSuperscriptAlef) + MakeUniRegEx(ArabicMaddahAbove) + "(?=" + MakeUniRegEx(ArabicLetterAlefWithHamzaAbove) + "|" + MakeUniRegEx(ArabicLetterAlefWithMaddaAbove) + ")|((?:^\s*|\s+)(?:" + MakeUniRegEx(ArabicLetterWaw) + MakeUniRegEx(ArabicFatha) + ")?" + MakeUniRegEx(ArabicLetterYeh) + MakeUniRegEx(ArabicFatha) + ")" + MakeUniRegEx(ArabicLetterSuperscriptAlef), _
            .Evaluator = "$1" + CStr(ArabicLetterAlef) + " "},
        New RuleTranslation With {.Rule = "Madda", .Match = MakeUniRegEx(ArabicMaddahAbove), _
            .Evaluator = String.Empty},
        New RuleTranslation With {.Rule = "SmallCircleShowingNonReadLetters", .Match = MakeUniRegEx(ArabicSmallHighRoundedZero) + "|" + MakeUniRegEx(ArabicSmallHighUprightRectangularZero) + "|" + MakeUniRegEx(ArabicEmptyCentreHighStop), _
            .Evaluator = String.Empty}, _
        New RuleTranslation With {.Rule = "SadSmallHighSeen", .Match = MakeUniRegEx(ArabicLetterSad) + MakeUniRegEx(ArabicSmallHighSeen), _
            .Evaluator = CStr(ArabicLetterSeen)}, _
        New RuleTranslation With {.Rule = "SmallLowSeen", .Match = MakeUniRegEx(ArabicSmallLowSeen), _
            .Evaluator = String.Empty}, _
        New RuleTranslation With {.Rule = "SmallMeemIghlab", .Match = MakeUniRegEx(ArabicSmallHighMeemIsolatedForm) + "|" + MakeUniRegEx(ArabicSmallLowMeem), _
            .Evaluator = String.Empty}, _
        New RuleTranslation With {.Rule = "SmallHighNoon", .Match = MakeUniRegEx(ArabicSmallHighNoon), _
            .Evaluator = CStr(ArabicLetterNoon)}, _
        New RuleTranslation With {.Rule = "Noon", .Match = "(\S" + MakeUniRegEx(ArabicLetterAlefWithHamzaAbove) + MakeUniRegEx(ArabicFatha) + ")(?=" + MakeUniRegEx(ArabicLetterLam) + MakeUniRegEx(ArabicShadda) + MakeUniRegEx(ArabicFatha) + MakeUniRegEx(ArabicLetterWaw) + ")", _
            .Evaluator = "$1" + CStr(ArabicLetterNoon) + " "}, _
        New RuleTranslation With {.Rule = "AlefWithHamzaAbove", .Match = "(" + MakeUniRegEx(ArabicFatha) + ")" + MakeUniRegEx(ArabicTatweel) + MakeUniRegEx(ArabicHamzaAbove) + MakeUniRegEx(ArabicFathatan) + MakeUniRegEx(ArabicLetterAlef), _
            .Evaluator = "$1" + CStr(ArabicLetterAlefWithHamzaAbove) + CStr(ArabicFathatan)}, _
        New RuleTranslation With {.Rule = "AlefWithHamzaAbove", .Match = "((?:^\s*|\s+)" + MakeUniRegEx(ArabicLetterAlefWithHamzaAbove) + MakeUniRegEx(ArabicFatha) + ")" + MakeUniRegEx(ArabicLetterHamza) + "(?=" + MakeUniRegEx(ArabicDamma) + ")|(^\s*|\s+|" + MakeUniRegEx(ArabicFatha) + ")" + MakeUniRegEx(ArabicLetterHamza) + "(?=" + MakeUniRegEx(ArabicFatha) + ")|(?:" + MakeUniRegEx(ArabicLetterAlef) + "|" + MakeUniRegEx(ArabicLetterSuperscriptAlef) + ")" + MakeUniRegEx(ArabicHamzaAbove) + "|(" + MakeUniRegEx(ArabicFatha) + ")" + MakeUniRegEx(ArabicLetterWawWithHamzaAbove) + "(?!" + MakeUniRegEx(ArabicDamma) + "(?!\s*$|\s+))|((?:(?:" + MakeUniRegEx(ArabicLetterTeh) + "|" + MakeUniRegEx(ArabicLetterYeh) + ")" + MakeUniRegEx(ArabicFatha) + MakeUniRegEx(ArabicLetterYeh) + "|" + MakeUniRegEx(ArabicLetterLam) + "|" + MakeUniRegEx(ArabicLetterJeem) + "|" + MakeUniRegEx(ArabicLetterTah) + "|" + MakeUniRegEx(ArabicLetterNoon) + "|" + MakeUniRegEx(ArabicLetterSeen) + "|" + MakeUniRegEx(ArabicLetterSheen) + ")" + MakeUniRegEx(ArabicSukun) + "|" + MakeUniRegEx(ArabicFatha) + ")" + MakeUniRegEx(ArabicTatweel) + MakeUniRegEx(ArabicHamzaAbove) + "(?=" + MakeUniRegEx(ArabicFatha) + "|" + MakeUniRegEx(ArabicFathatan) + "(?!" + MakeUniRegEx(ArabicLetterAlef) + ")|" + MakeUniRegEx(ArabicSukun) + ")", _
            .Evaluator = "$1$2$3$4" + CStr(ArabicLetterAlefWithHamzaAbove)}, _
        New RuleTranslation With {.Rule = "WawWithHamzaAbove", .Match = "(" + MakeUniRegEx(ArabicDamma) + ")(?:" + MakeUniRegEx(ArabicTatweel) + MakeUniRegEx(ArabicHamzaAbove) + "|" + MakeUniRegEx(ArabicLetterHamza) + "(?=" + MakeUniRegEx(ArabicSukun) + "))", _
            .Evaluator = "$1" + CStr(ArabicLetterWawWithHamzaAbove)}, _
        New RuleTranslation With {.Rule = "AlefWithHamzaBelow", .Match = "(" + MakeUniRegEx(ArabicFatha) + ")" + MakeUniRegEx(ArabicLetterHamza) + "(?=" + MakeUniRegEx(ArabicKasra) + ")", _
            .Evaluator = "$1" + CStr(ArabicLetterAlefWithHamzaBelow)}, _
        New RuleTranslation With {.Rule = "YehWithHamzaAbove", .Match = "(?:(" + MakeUniRegEx(ArabicSukun) + "|" + MakeUniRegEx(ArabicKasra) + "|" + MakeUniRegEx(ArabicFatha) + "|" + MakeUniRegEx(ArabicLetterYeh) + ")" + MakeUniRegEx(ArabicTatweel) + MakeUniRegEx(ArabicHamzaAbove) + "(?=" + MakeUniRegEx(ArabicFatha) + "|" + MakeUniRegEx(ArabicFathatan) + "|" + MakeUniRegEx(ArabicKasra) + "|" + MakeUniRegEx(ArabicDamma) + ")|(" + MakeUniRegEx(ArabicLetterSuperscriptAlef) + "|" + MakeUniRegEx(ArabicLetterAlef) + ")" + MakeUniRegEx(ArabicTatweel) + MakeUniRegEx(ArabicHamzaAbove) + "(?!" + MakeUniRegEx(ArabicDamma) + ")|(" + MakeUniRegEx(ArabicLetterSuperscriptAlef) + "|(?:" + MakeUniRegEx(ArabicLetterZain) + "|" + MakeUniRegEx(ArabicLetterReh) + ")" + MakeUniRegEx(ArabicKasra) + ")" + MakeUniRegEx(ArabicLetterHamza) + ")(?=(" + MakeUniRegEx(ArabicKasra) + "|" + MakeUniRegEx(ArabicDamma) + "|" + MakeUniRegEx(ArabicFatha) + "(?!" + MakeUniRegEx(ArabicLetterAlefMaksura) + ")|" + MakeUniRegEx(ArabicFathatan) + "|" + MakeUniRegEx(ArabicSukun) + ")(?!\s*$|\s+))", _
            .Evaluator = "$1$2$3" + CStr(ArabicLetterYehWithHamzaAbove)}, _
        New RuleTranslation With {.Rule = "YehWithHamzaAbove", .Match = "(" + MakeUniRegEx(ArabicLetterAlef) + ")" + MakeUniRegEx(ArabicLetterHamza) + MakeUniRegEx(ArabicKasra) + "(?=" + MakeUniRegEx(ArabicLetterYeh) + ")|(" + MakeUniRegEx(ArabicLetterWaw) + MakeUniRegEx(ArabicFatha) + MakeUniRegEx(ArabicLetterMeem) + MakeUniRegEx(ArabicFatha) + MakeUniRegEx(ArabicLetterLam) + MakeUniRegEx(ArabicFatha) + ")" + MakeUniRegEx(ArabicLetterAlefWithHamzaBelow) + MakeUniRegEx(ArabicKasra) + MakeUniRegEx(ArabicLetterYeh) + "(?!\s*$|\s+|" + MakeUniRegEx(ArabicLetterMeem) + ")", _
            .Evaluator = "$1$2" + CStr(ArabicLetterYehWithHamzaAbove) + CStr(ArabicKasra)}, _
        New RuleTranslation With {.Rule = "NonReadWawAndYeh", .Match = "(" + MakeUniRegEx(ArabicLetterYeh) + MakeUniRegEx(ArabicSukun) + "|" + MakeUniRegEx(ArabicFatha) + "(?:" + MakeUniRegEx(ArabicLetterFeh) + "|" + MakeUniRegEx(ArabicLetterBeh) + ")" + (ArabicFatha) + MakeUniRegEx(ArabicLetterAlefWithHamzaBelow) + MakeUniRegEx(ArabicKasra) + ")" + MakeUniRegEx(ArabicLetterYeh) + "|(" + MakeUniRegEx(ArabicFatha) + MakeUniRegEx(ArabicLetterAlefWithHamzaAbove) + MakeUniRegEx(ArabicDamma) + ")" + MakeUniRegEx(ArabicLetterWaw) + "(?=" + MakeUniRegEx(ArabicLetterReh) + ")", _
            .Evaluator = "$1$2"}, _
        New RuleTranslation With {.Rule = "WawAlefMaksuraSuperscriptAlef", .Match = "(" + MakeUniRegEx(ArabicFatha) + ")(?:" + MakeUniRegEx(ArabicLetterWaw) + "|" + MakeUniRegEx(ArabicLetterAlefMaksura) + ")" + MakeUniRegEx(ArabicLetterSuperscriptAlef) + "(?!\s*$|\s+)", _
            .Evaluator = "$1" + CStr(ArabicLetterAlef)}, _
        New RuleTranslation With {.Rule = "SuperscriptAlef", .Match = "((" + MakeUniRegEx(ArabicLetterHah) + MakeUniRegEx(ArabicSukun) + MakeUniRegEx(ArabicLetterMeem) + "|(?:^\s*|\s+)(?:(?:(?:" + MakeUniRegEx(ArabicLetterWaw) + "|" + MakeUniRegEx(ArabicLetterFeh) + "|" + MakeUniRegEx(ArabicLetterAlefWithHamzaAbove) + ")" + MakeUniRegEx(ArabicFatha) + ")?(?:(?:" + MakeUniRegEx(ArabicLetterKaf) + "|" + MakeUniRegEx(ArabicLetterWaw) + ")" + MakeUniRegEx(ArabicFatha) + "|(?:" + MakeUniRegEx(ArabicLetterBeh) + "|" + MakeUniRegEx(ArabicLetterLam) + ")" + MakeUniRegEx(ArabicKasra) + "))?" + MakeUniRegEx(ArabicLetterThal) + MakeUniRegEx(ArabicShadda) + "?|" + MakeUniRegEx(ArabicLetterAlefWithHamzaBelow) + MakeUniRegEx(ArabicKasra) + MakeUniRegEx(ArabicLetterLam) + "|(?:^\s*|\s+)" + MakeUniRegEx(ArabicLetterWaw) + MakeUniRegEx(ArabicSukun) + MakeUniRegEx(ArabicLetterLam) + ")?" + MakeUniRegEx(ArabicFatha) + MakeUniRegEx(ArabicTatweel) + "?)" + MakeUniRegEx(ArabicLetterSuperscriptAlef) + "(?=" + MakeUniRegEx(ArabicLetterGhain) + "|" + MakeUniRegEx(ArabicLetterHeh) + MakeUniRegEx(ArabicDamma) + MakeUniRegEx(ArabicLetterNoon) + MakeUniRegEx(ArabicFatha) + MakeUniRegEx(ArabicLetterAlef) + ")?(?!\s*$|\s+|" + MakeUniRegEx(ArabicLetterLam) + MakeUniRegEx(ArabicKasra) + MakeUniRegEx(ArabicLetterKaf) + "(?:" + MakeUniRegEx(ArabicFatha) + "|" + MakeUniRegEx(ArabicDamma) + MakeUniRegEx(ArabicLetterMeem) + MakeUniRegEx(ArabicDamma) + "|" + MakeUniRegEx(ArabicDamma) + MakeUniRegEx(ArabicLetterNoon) + MakeUniRegEx(ArabicShadda) + MakeUniRegEx(ArabicFatha) + ")|" + MakeUniRegEx(ArabicLetterWawWithHamzaAbove) + MakeUniRegEx(ArabicDamma) + MakeUniRegEx(ArabicLetterLam) + MakeUniRegEx(ArabicFatha) + MakeUniRegEx(ArabicLetterAlef) + MakeUniRegEx(ArabicLetterHamza) + MakeUniRegEx(ArabicKasra) + "|" + MakeUniRegEx(ArabicLetterKaf) + MakeUniRegEx(ArabicFatha) + MakeUniRegEx(ArabicLetterThal) + MakeUniRegEx(ArabicFatha) + MakeUniRegEx(ArabicLetterAlef) + "|" + MakeUniRegEx(ArabicLetterKaf) + MakeUniRegEx(ArabicKasra) + MakeUniRegEx(ArabicLetterNoon) + "(?:" + MakeUniRegEx(ArabicShadda) + MakeUniRegEx(ArabicKasra) + MakeUniRegEx(ArabicLetterYeh) + "|" + MakeUniRegEx(ArabicShadda) + MakeUniRegEx(ArabicFatha) + "(?:" + MakeUniRegEx(ArabicLetterHeh) + MakeUniRegEx(ArabicDamma) + "(?:" + MakeUniRegEx(ArabicLetterMeem) + MakeUniRegEx(ArabicSukun) + ")?|" + MakeUniRegEx(ArabicLetterKaf) + MakeUniRegEx(ArabicDamma) + MakeUniRegEx(ArabicLetterMeem) + MakeUniRegEx(ArabicSukun) + "|" + MakeUniRegEx(ArabicLetterAlef) + ")?|" + MakeUniRegEx(ArabicKasra) + "|" + MakeUniRegEx(ArabicSukun) + ")?(?:\s*$|\s+)|" + MakeUniRegEx(ArabicLetterThal) + "(?:" + MakeUniRegEx(ArabicFatha) + "(?:" + MakeUniRegEx(ArabicLetterAlef) + "|" + MakeUniRegEx(ArabicLetterSuperscriptAlef) + ")(?:" + MakeUniRegEx(ArabicLetterNoon) + MakeUniRegEx(ArabicKasra) + ")?|" + MakeUniRegEx(ArabicKasra) + MakeUniRegEx(ArabicLetterHeh) + MakeUniRegEx(ArabicKasra) + ")|" + MakeUniRegEx(ArabicLetterHeh) + "(?:" + MakeUniRegEx(ArabicKasra) + "|" + MakeUniRegEx(ArabicDamma) + MakeUniRegEx(ArabicLetterKaf) + MakeUniRegEx(ArabicDamma) + MakeUniRegEx(ArabicLetterMeem) + MakeUniRegEx(ArabicSukun) + "|" + MakeUniRegEx(ArabicFathatan) + MakeUniRegEx(ArabicLetterAlef) + "|" + MakeUniRegEx(ArabicFatha) + "(?:" + MakeUniRegEx(ArabicLetterKaf) + MakeUniRegEx(ArabicFatha) + ")?)(?:\s+|\s*$)|" + MakeUniRegEx(ArabicLetterYehWithHamzaAbove) + MakeUniRegEx(ArabicKasra) + MakeUniRegEx(ArabicLetterKaf) + "(?:" + MakeUniRegEx(ArabicFatha) + "|" + MakeUniRegEx(ArabicDamma) + MakeUniRegEx(ArabicLetterMeem) + MakeUniRegEx(ArabicSukun) + ")(?:\s+|\s*$))|(" + MakeUniRegEx(ArabicLetterThal) + MakeUniRegEx(ArabicShadda) + "?" + MakeUniRegEx(ArabicFatha) + ")?" + MakeUniRegEx(ArabicLetterSuperscriptAlef) + "(?=" + MakeUniRegEx(ArabicLetterLam) + MakeUniRegEx(ArabicKasra) + MakeUniRegEx(ArabicLetterKaf) + MakeUniRegEx(ArabicFatha) + ")|(" + MakeUniRegEx(ArabicLetterWaw) + MakeUniRegEx(ArabicFatha) + MakeUniRegEx(ArabicLetterLam) + MakeUniRegEx(ArabicFatha) + ")" + MakeUniRegEx(ArabicLetterSuperscriptAlef) + "(?=" + MakeUniRegEx(ArabicLetterYeh) + MakeUniRegEx(ArabicFatha) + ")|(" + MakeUniRegEx(ArabicLetterMeem) + MakeUniRegEx(ArabicFatha) + MakeUniRegEx(ArabicLetterSeen) + MakeUniRegEx(ArabicFatha) + ")" + MakeUniRegEx(ArabicLetterSuperscriptAlef), _
            .Evaluator = "$1$4$5" + CStr(ArabicLetterAlef), .NegativeMatch = "$2$3"}, _
        New RuleTranslation With {.Rule = "Hamza", .Match = MakeUniRegEx(ArabicTatweel) + MakeUniRegEx(ArabicHamzaAbove) + "|(" + MakeUniRegEx(ArabicDamma) + MakeUniRegEx(ArabicLetterWaw) + ")" + MakeUniRegEx(ArabicLetterAlefWithHamzaAbove) + "(?=(?:" + MakeUniRegEx(ArabicFatha) + "|" + MakeUniRegEx(ArabicDamma) + ")(?:\s*$|\s+))|(" + MakeUniRegEx(ArabicLetterAlef) + "|" + MakeUniRegEx(ArabicLetterAlefWithMaddaAbove) + ")(?:" + MakeUniRegEx(ArabicLetterWawWithHamzaAbove) + "(?=(?:" + MakeUniRegEx(ArabicDamma) + "|" + MakeUniRegEx(ArabicDammatan) + ")(?:\s*$|\s+))|" + MakeUniRegEx(ArabicLetterYehWithHamzaAbove) + "(?=" + MakeUniRegEx(ArabicKasra) + "(?:\s*$|\s+)))", _
            .Evaluator = "$1$2" + CStr(ArabicLetterHamza)}, _
        New RuleTranslation With {.Rule = "Fathatan", .Match = "(" + MakeUniRegEx(ArabicLetterJeem) + ")" + MakeUniRegEx(ArabicFatha) + "(?=" + MakeUniRegEx(ArabicLetterAlef) + "(?:\s*$|\s+))", _
            .Evaluator = "$1" + CStr(ArabicFathatan)}, _
        New RuleTranslation With {.Rule = "Fatha", .Match = "((?:" + MakeUniRegEx(ArabicLetterNoon) + MakeUniRegEx(ArabicFatha) + "|" + MakeUniRegEx(ArabicLetterReh) + MakeUniRegEx(ArabicKasra) + ")" + MakeUniRegEx(ArabicLetterJeem) + ")" + MakeUniRegEx(ArabicFathatan) + "(?=" + MakeUniRegEx(ArabicLetterAlef) + "(?:\s*$|\s+))", _
            .Evaluator = "$1" + CStr(ArabicFatha)}
    }
    'New RuleTranslation With {.Rule = "SuperscriptAlef", .Match = "(?:((?:" + MakeRegMultiEx(Array.ConvertAll(SimpleSuperscriptAlefBefore, Function(Str As String) MakeUniRegEx(TransliterateFromBuckwalter(Str.Replace(".", "").Replace("""", "").Replace("@", "").Replace("[", "").Replace("]", "").Replace("-", "").Replace("^", ""))))) + "))|((?:" + MakeRegMultiEx(Array.ConvertAll(SimpleSuperscriptAlefNotBefore, Function(Str As String) MakeUniRegEx(TransliterateFromBuckwalter(Str.Replace(".", "").Replace("""", "").Replace("@", "").Replace("[", "").Replace("]", "").Replace("-", "").Replace("^", ""))))) + ")))" + MakeUniRegEx(ArabicLetterSuperscriptAlef) + "(?=(?:" + MakeRegMultiEx(Array.ConvertAll(SimpleSuperscriptAlefAfter, Function(Str As String) MakeUniRegEx(TransliterateFromBuckwalter(Str.Replace(".", "").Replace("""", "").Replace("@", "").Replace("[", "").Replace("]", "").Replace("-", "").Replace("^", ""))))) + "))?(?!" + MakeRegMultiEx(Array.ConvertAll(SimpleSuperscriptAlefNotAfter, Function(Str As String) MakeUniRegEx(TransliterateFromBuckwalter(Str.Replace(".", "").Replace("""", "").Replace("@", "").Replace("[", "").Replace("]", "").Replace("-", "").Replace("^", ""))))) + ")", _
    '    .Evaluator = "$1" + CStr(ArabicLetterAlef), .NegativeMatch = "$2"}, _

    Public Shared SimpleScript As RuleTranslation() = { _
        New RuleTranslation With {.Rule = "RemoveShadda", .Match = "((?:^\s*|\s+)(?:" + MakeRegMultiEx(Array.ConvertAll(ArabicLetters, Function(Str As String) MakeUniRegEx(Str))) + "))" + MakeUniRegEx(ArabicShadda) + "|((?:" + MakeUniRegEx(ArabicLetterQaf) + "|" + MakeUniRegEx(ArabicLetterKaf) + ")" + MakeUniRegEx(ArabicLetterKaf) + "|" + MakeUniRegEx(ArabicLetterDal) + MakeUniRegEx(ArabicLetterTeh) + "|" + MakeUniRegEx(ArabicLetterYeh) + MakeUniRegEx(ArabicLetterYeh) + "|" + MakeUniRegEx(ArabicLetterHeh) + MakeUniRegEx(ArabicLetterHeh) + ")" + MakeUniRegEx(ArabicShadda) + "(?=(?:" + MakeUniRegEx(ArabicKasra) + MakeUniRegEx(ArabicLetterKaf) + "|" + MakeUniRegEx(ArabicFatha) + MakeUniRegEx(ArabicLetterHeh) + ")?(?:" + MakeUniRegEx(ArabicDamma) + "|" + MakeUniRegEx(ArabicFatha) + ")(?:(?:(?:" + MakeUniRegEx(ArabicLetterNoon) + "|" + MakeUniRegEx(ArabicLetterHeh) + MakeUniRegEx(ArabicDamma) + "?)?" + MakeUniRegEx(ArabicFatha) + MakeUniRegEx(ArabicLetterAlef) + ")|" + MakeUniRegEx(ArabicLetterNoon) + MakeUniRegEx(ArabicShadda) + MakeUniRegEx(ArabicFatha) + "|(?:(?:(?:" + MakeUniRegEx(ArabicLetterMeem) + MakeUniRegEx(ArabicDamma) + MakeUniRegEx(ArabicLetterWaw) + ")?" + MakeUniRegEx(ArabicLetterHeh) + "|" + MakeUniRegEx(ArabicLetterKaf) + ")(?:" + MakeUniRegEx(ArabicDamma) + "|" + MakeUniRegEx(ArabicFatha) + "))?(?:" + MakeUniRegEx(ArabicLetterMeem) + "(?:" + MakeUniRegEx(ArabicSukun) + "|" + MakeUniRegEx(ArabicDamma) + ")?)?)(?:\s*$|\s+))", _
            .Evaluator = "$1$2"}, _
        New RuleTranslation With {.Rule = "AddSukun", .Match = "((?:^\s*|\s+)(?:" + MakeRegMultiEx(Array.ConvertAll(ArabicUniqueLetters, Function(Str As String) MakeUniRegEx(TransliterateFromBuckwalter(Str).Replace(CStr(ArabicMaddahAbove), String.Empty)))) + ")(?:\s+|\s*$))|(" + MakeRegMultiEx(Array.ConvertAll(ArabicSunLetters, Function(Str As String) MakeUniRegEx(Str))) + "|" + MakeRegMultiEx(Array.ConvertAll(ArabicMoonLettersNoVowels, Function(Str As String) MakeUniRegEx(Str))) + "|" + MakeUniRegEx(ArabicFatha) + "(" + MakeUniRegEx(ArabicLetterWaw) + "|" + MakeUniRegEx(ArabicLetterYeh) + "))(?=\s*$|\s+|(" + MakeRegMultiEx(Array.ConvertAll(ArabicSunLetters, Function(Str As String) MakeUniRegEx(Str))) + ")(?!" + MakeUniRegEx(ArabicShadda) + ")|" + MakeRegMultiEx(Array.ConvertAll(ArabicMoonLetters, Function(Str As String) MakeUniRegEx(Str))) + ")", _
            .Evaluator = "$2" + CStr(ArabicSukun), .NegativeMatch = "$1"}
    }
    Public Shared SimpleCleanScript As RuleTranslation() = { _
        New RuleTranslation With {.Rule = "RemoveDiacritics", .Match = MakeUniRegEx(ArabicShadda) + "|" + MakeUniRegEx(ArabicSukun) + "|" + MakeUniRegEx(ArabicFatha) + "|" + MakeUniRegEx(ArabicKasra) + "|" + MakeUniRegEx(ArabicDamma) + "|" + MakeUniRegEx(ArabicFathatan) + "|" + MakeUniRegEx(ArabicKasratan) + "|" + MakeUniRegEx(ArabicDammatan) + "|" + MakeUniRegEx(ArabicLetterSuperscriptAlef), _
            .Evaluator = String.Empty}
    }
    Public Shared SimpleMinimalScript As RuleTranslation() = { _
        New RuleTranslation With {.Rule = "RemoveShadda", .Match = "((?:^\s*|\s+)(?:" + MakeRegMultiEx(Array.ConvertAll(ArabicLetters, Function(Str As String) MakeUniRegEx(Str))) + "))" + MakeUniRegEx(ArabicShadda) + "|((?:" + MakeUniRegEx(ArabicLetterQaf) + "|" + MakeUniRegEx(ArabicLetterKaf) + ")" + MakeUniRegEx(ArabicLetterKaf) + "|" + MakeUniRegEx(ArabicLetterDal) + MakeUniRegEx(ArabicLetterTeh) + "|" + MakeUniRegEx(ArabicLetterYeh) + MakeUniRegEx(ArabicLetterYeh) + "|" + MakeUniRegEx(ArabicLetterHeh) + MakeUniRegEx(ArabicLetterHeh) + ")" + MakeUniRegEx(ArabicShadda) + "(?=(?:" + MakeUniRegEx(ArabicKasra) + MakeUniRegEx(ArabicLetterKaf) + "|" + MakeUniRegEx(ArabicFatha) + MakeUniRegEx(ArabicLetterHeh) + ")?(?:" + MakeUniRegEx(ArabicDamma) + "|" + MakeUniRegEx(ArabicFatha) + ")(?:(?:(?:" + MakeUniRegEx(ArabicLetterNoon) + "|" + MakeUniRegEx(ArabicLetterHeh) + MakeUniRegEx(ArabicDamma) + "?)?" + MakeUniRegEx(ArabicFatha) + MakeUniRegEx(ArabicLetterAlef) + ")|" + MakeUniRegEx(ArabicLetterNoon) + MakeUniRegEx(ArabicShadda) + MakeUniRegEx(ArabicFatha) + "|(?:(?:(?:" + MakeUniRegEx(ArabicLetterMeem) + MakeUniRegEx(ArabicDamma) + MakeUniRegEx(ArabicLetterWaw) + ")?" + MakeUniRegEx(ArabicLetterHeh) + "|" + MakeUniRegEx(ArabicLetterKaf) + ")(?:" + MakeUniRegEx(ArabicDamma) + "|" + MakeUniRegEx(ArabicFatha) + "))?(?:" + MakeUniRegEx(ArabicLetterMeem) + "(?:" + MakeUniRegEx(ArabicSukun) + "|" + MakeUniRegEx(ArabicDamma) + ")?)?)(?:\s*$|\s+))", _
            .Evaluator = "$1$2"}, _
        New RuleTranslation With {.Rule = "RemoveSukun", .Match = MakeUniRegEx(ArabicSukun), _
            .Evaluator = String.Empty}, _
        New RuleTranslation With {.Rule = "RemoveLongVowelDiacritics", .Match = "((?:^\s*|\s+)(?:" + MakeUniRegEx(ArabicLetterKaf) + "|" + MakeUniRegEx(ArabicLetterWaw) + "|(?:" + MakeUniRegEx(ArabicLetterAlefWithHamzaAbove) + MakeUniRegEx(ArabicFatha) + ")?" + MakeUniRegEx(ArabicLetterFeh) + "))" + MakeUniRegEx(ArabicFatha) + "(?=" + MakeUniRegEx(ArabicLetterAlef) + "(?!" + MakeUniRegEx(ArabicFatha) + "|" + MakeUniRegEx(ArabicFathatan) + MakeUniRegEx(ArabicLetterAlef) + "?|" + MakeUniRegEx(ArabicDamma) + "|" + MakeUniRegEx(ArabicShadda) + "|" + MakeUniRegEx(ArabicLetterAlef) + "?(?: " + MakeUniRegEx(ArabicSmallHighJeem) + "| " + MakeUniRegEx(ArabicSmallHighLamAlef) + ")? " + MakeUniRegEx(ArabicLetterAlef) + ")(?:(?:" + MakeUniRegEx(ArabicLetterTeh) + "|" + MakeUniRegEx(ArabicLetterTheh) + "|" + MakeUniRegEx(ArabicLetterDad) + "|" + MakeUniRegEx(ArabicLetterQaf) + ")" + MakeUniRegEx(ArabicFatha) + "|" + MakeUniRegEx(ArabicLetterHamza) + "|(?:" + MakeUniRegEx(ArabicLetterJeem) + "|" + MakeUniRegEx(ArabicLetterBeh) + "|" + MakeUniRegEx(ArabicLetterReh) + "|" + MakeUniRegEx(ArabicLetterSeen) + "|" + MakeUniRegEx(ArabicLetterSad) + "|" + MakeUniRegEx(ArabicLetterTah) + "|" + MakeUniRegEx(ArabicLetterHah) + ")" + MakeUniRegEx(ArabicKasra) + "|" + MakeUniRegEx(ArabicLetterKaf) + MakeUniRegEx(ArabicKasra) + MakeUniRegEx(ArabicLetterHeh) + "|" + MakeUniRegEx(ArabicLetterHeh) + "(?:" + MakeUniRegEx(ArabicKasra) + "|" + MakeUniRegEx(ArabicDamma) + ")|(?:" + MakeUniRegEx(ArabicLetterDal) + "|" + MakeUniRegEx(ArabicLetterQaf) + "|" + MakeUniRegEx(ArabicLetterLam) + ")(?:" + MakeUniRegEx(ArabicDamma) + "|" + MakeUniRegEx(ArabicKasra) + "|" + MakeUniRegEx(ArabicKasratan) + ")|(?:" + MakeUniRegEx(ArabicLetterZain) + "|" + MakeUniRegEx(ArabicLetterAin) + ")(?:" + MakeUniRegEx(ArabicKasra) + "|" + MakeUniRegEx(ArabicFatha) + ")|" + MakeUniRegEx(ArabicLetterNoon) + MakeUniRegEx(ArabicKasratan) + "))|((?:^\s*|\s+)" + MakeUniRegEx(ArabicLetterLam) + ")" + MakeUniRegEx(ArabicFatha) + "(?=" + MakeUniRegEx(ArabicLetterAlef) + "(?! " + MakeUniRegEx(ArabicLetterAlef) + "|" + MakeUniRegEx(ArabicLetterNoon) + "|" + MakeUniRegEx(ArabicLetterFeh) + "|" + MakeUniRegEx(ArabicLetterReh) + "|" + MakeUniRegEx(ArabicLetterTeh) + "|" + MakeUniRegEx(ArabicLetterBeh) + "|" + MakeUniRegEx(ArabicLetterKhah) + "|" + MakeUniRegEx(ArabicLetterSeen) + "|" + MakeUniRegEx(ArabicLetterSad) + "))|(^" + MakeUniRegEx(ArabicLetterLam) + ")" + MakeUniRegEx(ArabicFatha) + "(?=" + MakeUniRegEx(ArabicLetterAlef) + MakeUniRegEx(ArabicLetterBeh) + ")(?!" + MakeUniRegEx(ArabicFatha) + "|" + MakeUniRegEx(ArabicFathatan) + MakeUniRegEx(ArabicLetterAlef) + "?|" + MakeUniRegEx(ArabicDamma) + "|" + MakeUniRegEx(ArabicShadda) + "|" + MakeUniRegEx(ArabicLetterAlef) + "?(?: " + MakeUniRegEx(ArabicSmallHighJeem) + "| " + MakeUniRegEx(ArabicSmallHighLamAlef) + ")? " + MakeUniRegEx(ArabicLetterAlef) + ")|" + MakeUniRegEx(ArabicFatha) + "(?=" + MakeUniRegEx(ArabicLetterSuperscriptAlef) + "|" + MakeUniRegEx(ArabicLetterAlefMaksura) + "(?!(?:\s*$|\s+" + MakeUniRegEx(ArabicSmallHighLamAlef) + "| " + MakeUniRegEx(ArabicSmallHighLigatureSadWithLamWithAlefMaksura) + "))(?!" + MakeUniRegEx(ArabicFatha) + "|" + MakeUniRegEx(ArabicFathatan) + MakeUniRegEx(ArabicLetterAlef) + "?|" + MakeUniRegEx(ArabicDamma) + "|" + MakeUniRegEx(ArabicShadda) + "|" + MakeUniRegEx(ArabicLetterAlef) + "?(?: " + MakeUniRegEx(ArabicSmallHighJeem) + "| " + MakeUniRegEx(ArabicSmallHighLamAlef) + ")? " + MakeUniRegEx(ArabicLetterAlef) + "))|((?:^\s*|\s+)(?:(?:" + MakeUniRegEx(ArabicLetterWaw) + MakeUniRegEx(ArabicFatha) + ")?" + MakeUniRegEx(ArabicLetterTeh) + "|" + MakeUniRegEx(ArabicLetterWaw) + "|" + MakeUniRegEx(ArabicLetterFeh) + "|" + MakeUniRegEx(ArabicLetterKaf) + "))?" + MakeUniRegEx(ArabicFatha) + "(?=" + MakeUniRegEx(ArabicLetterAlef) + MakeUniRegEx(ArabicLetterLam) + ")|((?:^\s*|\s+)(?:" + MakeUniRegEx(ArabicLetterWaw) + "|(?:" + MakeUniRegEx(ArabicLetterAlefWithHamzaAbove) + MakeUniRegEx(ArabicFatha) + ")?" + MakeUniRegEx(ArabicLetterFeh) + "|" + MakeUniRegEx(ArabicLetterLam) + ")|" + MakeUniRegEx(ArabicLetterYehWithHamzaAbove) + ")?" + MakeUniRegEx(ArabicFatha) + "(?=" + MakeUniRegEx(ArabicLetterAlef) + "(?!" + MakeUniRegEx(ArabicLetterLam) + "(?!" + MakeUniRegEx(ArabicKasra) + "|" + MakeUniRegEx(ArabicDamma) + "|" + MakeUniRegEx(ArabicKasratan) + ")|" + MakeUniRegEx(ArabicFatha) + "|" + MakeUniRegEx(ArabicFathatan) + MakeUniRegEx(ArabicLetterAlef) + "?|" + MakeUniRegEx(ArabicDamma) + "|" + MakeUniRegEx(ArabicShadda) + "|" + MakeUniRegEx(ArabicLetterAlef) + "?(?: " + MakeUniRegEx(ArabicSmallHighJeem) + "| " + MakeUniRegEx(ArabicSmallHighLamAlef) + ")? " + MakeUniRegEx(ArabicLetterAlef) + "))|(" + MakeUniRegEx(ArabicLetterJeem) + ")?" + MakeUniRegEx(ArabicKasra) + "(?=" + MakeUniRegEx(ArabicLetterYeh) + MakeUniRegEx(ArabicLetterHamza) + ")(?!" + MakeUniRegEx(ArabicFatha) + "|" + MakeUniRegEx(ArabicFathatan) + MakeUniRegEx(ArabicLetterAlef) + "?|" + MakeUniRegEx(ArabicDamma) + "|" + MakeUniRegEx(ArabicShadda) + "|" + MakeUniRegEx(ArabicLetterAlef) + "?(?: " + MakeUniRegEx(ArabicSmallHighJeem) + "| " + MakeUniRegEx(ArabicSmallHighLamAlef) + ")? " + MakeUniRegEx(ArabicLetterAlef) + ")|(" + MakeUniRegEx(ArabicLetterAlefWithHamzaAbove) + MakeUniRegEx(ArabicFatha) + MakeUniRegEx(ArabicLetterKhah) + "|" + MakeUniRegEx(ArabicKasra) + MakeUniRegEx(ArabicLetterNoon) + MakeUniRegEx(ArabicFatha) + MakeUniRegEx(ArabicLetterFeh) + MakeUniRegEx(ArabicLetterSeen) + "| " + MakeUniRegEx(ArabicLetterThal) + MakeUniRegEx(ArabicKasra) + MakeUniRegEx(ArabicLetterKaf) + MakeUniRegEx(ArabicLetterReh) + ")" + MakeUniRegEx(ArabicKasra) + "(?=" + MakeUniRegEx(ArabicLetterYeh) + "(?!$)(?!" + MakeUniRegEx(ArabicFatha) + "|" + MakeUniRegEx(ArabicFathatan) + MakeUniRegEx(ArabicLetterAlef) + "?|" + MakeUniRegEx(ArabicDamma) + "|" + MakeUniRegEx(ArabicShadda) + "|" + MakeUniRegEx(ArabicLetterAlef) + "?(?: " + MakeUniRegEx(ArabicSmallHighJeem) + "| " + MakeUniRegEx(ArabicSmallHighLamAlef) + ")? " + MakeUniRegEx(ArabicLetterAlef) + "))|(" + MakeUniRegEx(ArabicLetterAlefWithHamzaAbove) + MakeUniRegEx(ArabicFatha) + MakeUniRegEx(ArabicLetterKhah) + "|" + MakeUniRegEx(ArabicKasra) + MakeUniRegEx(ArabicLetterNoon) + MakeUniRegEx(ArabicFatha) + MakeUniRegEx(ArabicLetterFeh) + MakeUniRegEx(ArabicLetterSeen) + "| " + MakeUniRegEx(ArabicLetterThal) + MakeUniRegEx(ArabicKasra) + MakeUniRegEx(ArabicLetterKaf) + MakeUniRegEx(ArabicLetterReh) + ")?" + MakeUniRegEx(ArabicKasra) + "(?=" + MakeUniRegEx(ArabicLetterYeh) + "(?!" + MakeUniRegEx(ArabicLetterHamza) + ")(?!" + MakeUniRegEx(ArabicFatha) + "|" + MakeUniRegEx(ArabicFathatan) + MakeUniRegEx(ArabicLetterAlef) + "?|" + MakeUniRegEx(ArabicDamma) + "|" + MakeUniRegEx(ArabicShadda) + "|" + MakeUniRegEx(ArabicLetterAlef) + "?(?: " + MakeUniRegEx(ArabicSmallHighJeem) + "| " + MakeUniRegEx(ArabicSmallHighLamAlef) + ")? " + MakeUniRegEx(ArabicLetterAlef) + "))|((?:^\s*|\s+)(?:" + MakeUniRegEx(ArabicLetterWaw) + MakeUniRegEx(ArabicFatha) + ")?" + MakeUniRegEx(ArabicLetterAlefWithHamzaAbove) + ")" + MakeUniRegEx(ArabicDamma) + "(?=" + MakeUniRegEx(ArabicLetterWaw) + "(?!" + MakeUniRegEx(ArabicLetterLam) + "(?:" + MakeUniRegEx(ArabicFatha) + "(?:" + MakeUniRegEx(ArabicLetterSuperscriptAlef) + "|" + MakeUniRegEx(ArabicLetterAlef) + "(?:" + MakeUniRegEx(ArabicLetterTeh) + "|" + MakeUniRegEx(ArabicLetterHamza) + "))|" + MakeUniRegEx(ArabicKasra) + MakeUniRegEx(ArabicLetterYeh) + "|" + MakeUniRegEx(ArabicDamma) + MakeUniRegEx(ArabicLetterWaw) + "))(?!" + MakeUniRegEx(ArabicFatha) + "|" + MakeUniRegEx(ArabicFathatan) + MakeUniRegEx(ArabicLetterAlef) + "?|" + MakeUniRegEx(ArabicDamma) + "|" + MakeUniRegEx(ArabicShadda) + "|" + MakeUniRegEx(ArabicLetterAlef) + "?(?: " + MakeUniRegEx(ArabicSmallHighJeem) + "| " + MakeUniRegEx(ArabicSmallHighLamAlef) + ")? " + MakeUniRegEx(ArabicLetterAlef) + "))|((?:^\s*|\s+)(?:" + MakeUniRegEx(ArabicLetterWaw) + MakeUniRegEx(ArabicFatha) + ")?" + MakeUniRegEx(ArabicLetterAlefWithHamzaAbove) + ")?" + MakeUniRegEx(ArabicDamma) + "(?=" + MakeUniRegEx(ArabicLetterWaw) + "(?!" + MakeUniRegEx(ArabicLetterLam) + "(?:" + MakeUniRegEx(ArabicFatha) + MakeUniRegEx(ArabicLetterSuperscriptAlef) + MakeUniRegEx(ArabicLetterYehWithHamzaAbove) + "|" + MakeUniRegEx(ArabicKasra) + MakeUniRegEx(ArabicLetterYeh) + " " + MakeUniRegEx(ArabicLetterAlef) + "))(?!" + MakeUniRegEx(ArabicFatha) + "|" + MakeUniRegEx(ArabicFathatan) + MakeUniRegEx(ArabicLetterAlef) + "?|" + MakeUniRegEx(ArabicDamma) + "|" + MakeUniRegEx(ArabicShadda) + "|" + MakeUniRegEx(ArabicLetterAlef) + "?(?: " + MakeUniRegEx(ArabicSmallHighJeem) + "| " + MakeUniRegEx(ArabicSmallHighLamAlef) + ")? " + MakeUniRegEx(ArabicLetterAlef) + "))", _
            .Evaluator = "$1$2$3$7$9", .NegativeMatch = "$4$5$6$8$10"}
    }

    Public Shared ColoringRules As ColorRule() = { _
        New ColorRule With {.Rule = "Normal", .Match = "empty|helperfatha|helperkasra|helperdamma|helpermeem|assimilator|assimilatorincomplete|dipthong|compulsorystop|endofversestop|prostration|canstoporcontinue|prohibitedtostop|stopatfirstnotsecond|stopatsecondnotfirst|bettertostopbutpermissibletocontinue|bettertocontinuebutpermissibletostop|subtlestopwithoutbreath", .Color = Color.Black}, _
        New ColorRule With {.Rule = "NecessaryProlongation", .Match = "necessaryprolong", .Color = Color.DarkRed}, _
        New ColorRule With {.Rule = "ObligatoryProlongation", .Match = "obligatoryprolong", .Color = Color.FromArgb(175, 17, 28)}, _
        New ColorRule With {.Rule = "PermissibleProlongation", .Match = "permissibleprolong", .Color = Color.OrangeRed}, _
        New ColorRule With {.Rule = "NormalProlongation", .Match = "normalprolong", .Color = Color.FromArgb(213, 139, 24)}, _
        New ColorRule With {.Rule = "Nasalization", .Match = "nasalize", .Color = Color.Green}, _
        New ColorRule With {.Rule = "Unannounced", .Match = "assimilate|assimilateincomplete", .Color = Color.Gray}, _
        New ColorRule With {.Rule = "EmphaticPronounciation", .Match = "emphasis", .Color = Color.DarkBlue}, _
        New ColorRule With {.Rule = "UnrestLetters", .Match = "bounce", .Color = Color.Blue} _
    }

    'New RuleTranslation With {.Rule = "TehMarbutaPartiallyFlexible", .Match = "(" + MakeUniRegEx(ArabicLetterAlefWasla) + MakeUniRegEx(ArabicLetterLam) + "\S+)?" + MakeUniRegEx(ArabicLetterTehMarbuta) + "(?:" + MakeRegMultiEx(Array.ConvertAll(ArabicFathaDammaKasra, Function(Str As String) MakeUniRegEx(Str))) + ")",
    '                        .Evaluator = "$&", .NegativeMatch = "$1"}

    Public Shared ErrorCheckRules As RuleTranslation() = {
        New RuleTranslation With {.Rule = "NotValidStartEndCombination", .Match = MakeUniRegEx(ArabicLetterAlefWasla) + MakeUniRegEx(ArabicLetterLam) + "\S+(?:" + MakeRegMultiEx(Array.ConvertAll(ArabicTanweens, Function(Str As String) MakeUniRegEx(Str))) + ")",
                                    .Evaluator = "$&"}, _
        New RuleTranslation With {.Rule = "OnlyAtEndOfWord", .Match = "(" + MakeUniRegEx(ArabicFatha) + MakeUniRegEx(ArabicLetterAlef) + MakeUniRegEx(ArabicLetterAlefMaksura) + "(?=" + MakeUniRegEx(ArabicFatha) + "(?:\s*$|\s+)))|(" + MakeUniRegEx(ArabicSukun) + "|" + MakeUniRegEx(ArabicFatha) + "|" + MakeUniRegEx(ArabicKasra) + "|" + MakeUniRegEx(ArabicLetterSuperscriptAlef) + "|" + MakeUniRegEx(ArabicSmallHighRoundedZero) + "|" + MakeUniRegEx(ArabicEmptyCentreLowStop) + "|" + MakeUniRegEx(ArabicFathatan) + MakeUniRegEx(ArabicSmallLowMeem) + "?)?" + MakeUniRegEx(ArabicLetterAlefMaksura) + "|" + MakeUniRegEx(ArabicLetterAlefMaksura) + "(?!(?:" + MakeRegMultiEx(Array.ConvertAll(ArabicFathaDammaKasra, Function(Str As String) MakeUniRegEx(Str))) + "|" + MakeUniRegEx(ArabicDammatan) + "|" + MakeUniRegEx(ArabicShadda) + "(?:" + MakeRegMultiEx(Array.ConvertAll(ArabicFathaDammaKasra, Function(Str As String) MakeUniRegEx(Str))) + "|" + MakeUniRegEx(ArabicDammatan) + "|" + MakeUniRegEx(ArabicKasratan) + ")|" + MakeUniRegEx(ArabicMaddahAbove) + "(?:" + MakeUniRegEx(ArabicFatha) + "|" + MakeUniRegEx(ArabicDamma) + "|" + MakeUniRegEx(ArabicDammatan) + ")?|" + MakeUniRegEx(ArabicSukun) + "(?:" + MakeUniRegEx(ArabicHamzaAbove) + "(?:" + MakeUniRegEx(ArabicDammatan) + MakeUniRegEx(ArabicKasratan) + "))?|" + MakeUniRegEx(ArabicLetterSuperscriptAlef) + "(?:" + MakeUniRegEx(ArabicMaddahAbove) + "|" + MakeUniRegEx(ArabicLetterTehMarbuta) + "(?:" + MakeRegMultiEx(Array.ConvertAll(ArabicFathaDammaKasra, Function(Str As String) MakeUniRegEx(Str))) + "|" + MakeUniRegEx(ArabicFathatan) + "|" + MakeUniRegEx(ArabicKasratan) + ")|" + MakeUniRegEx(ArabicLetterMeem) + MakeUniRegEx(ArabicKasra) + "|" + MakeUniRegEx(ArabicLetterNoon) + "(?:" + MakeUniRegEx(ArabicFatha) + MakeUniRegEx(ArabicLetterAlef) + MakeUniRegEx(ArabicMaddahAbove) + "?|" + MakeUniRegEx(ArabicKasra) + "(?:" + MakeUniRegEx(ArabicLetterYeh) + MakeUniRegEx(ArabicLetterHeh) + MakeUniRegEx(ArabicDamma) + "|" + MakeUniRegEx(ArabicLetterAlefMaksura) + "(?:" + MakeUniRegEx(ArabicMaddahAbove) + "|" + MakeUniRegEx(ArabicFatha) + ")?)?)|" + MakeUniRegEx(ArabicLetterHeh) + "(?:" + MakeUniRegEx(ArabicKasra) + MakeUniRegEx(ArabicLetterNoon) + MakeUniRegEx(ArabicFatha) + MakeUniRegEx(ArabicLetterAlefMaksura) + MakeUniRegEx(ArabicKasra) + "|" + MakeUniRegEx(ArabicFatha) + MakeUniRegEx(ArabicLetterAlef) + MakeUniRegEx(ArabicMaddahAbove) + "?|" + MakeUniRegEx(ArabicDamma) + "(?:" + MakeUniRegEx(ArabicLetterNoon) + MakeUniRegEx(ArabicShadda) + MakeUniRegEx(ArabicFatha) + "|" + MakeUniRegEx(ArabicLetterMeem) + "(?:" + MakeUniRegEx(ArabicSukun) + "|" + MakeUniRegEx(ArabicDamma) + "|" + MakeUniRegEx(ArabicFatha) + MakeUniRegEx(ArabicLetterAlef) + ")?)?)|" + MakeUniRegEx(ArabicLetterKaf) + "(?:" + MakeUniRegEx(ArabicFatha) + "(?:" + MakeUniRegEx(ArabicLetterHeh) + MakeUniRegEx(ArabicDamma) + MakeUniRegEx(ArabicLetterMeem) + ")?|" + MakeUniRegEx(ArabicKasra) + "|" + MakeUniRegEx(ArabicDamma) + MakeUniRegEx(ArabicLetterMeem) + "(?:" + MakeUniRegEx(ArabicSukun) + "|" + MakeUniRegEx(ArabicDamma) + "|" + MakeUniRegEx(ArabicFatha) + MakeUniRegEx(ArabicLetterAlef) + ")?)|" + MakeUniRegEx(ArabicLetterLam) + MakeUniRegEx(ArabicFatha) + ")?)?(?:\s*$|\s+))", _
                                    .Evaluator = "$&", .NegativeMatch = "$1$2"}, _
        New RuleTranslation With {.Rule = "OnlyAtEndOfWord", .Match = "(" + MakeUniRegEx(ArabicFatha) + "(?:(?:" + MakeUniRegEx(ArabicLetterAlefMaksura) + "|" + MakeUniRegEx(ArabicLetterWaw) + ")" + MakeUniRegEx(ArabicLetterSuperscriptAlef) + ")?)?" + MakeUniRegEx(ArabicLetterTehMarbuta), _
                                    .Evaluator = "$&", .NegativeMatch = "$1"}, _
        New RuleTranslation With {.Rule = "OnlyAtEndOfWord", .Match = MakeUniRegEx(ArabicLetterTehMarbuta) + "(?!(?:" + MakeRegMultiEx(Array.ConvertAll(ArabicFathaDammaKasra, Function(Str As String) MakeUniRegEx(Str))) + "|(?:" + MakeRegMultiEx(Array.ConvertAll(ArabicTanweens, Function(Str As String) MakeUniRegEx(Str))) + ")(?:" + MakeUniRegEx(ArabicSmallLowMeem) + "|" + MakeUniRegEx(ArabicSmallHighMeemIsolatedForm) + ")?)(?:\s*$|\s+))", _
                                    .Evaluator = "$&"}, _
        New RuleTranslation With {.Rule = "OnlyAtEndOfWord", .Match = "(" + MakeUniRegEx(ArabicLetterTehMarbuta) + ")?" + MakeUniRegEx(ArabicFathatan) + "(?:" + MakeUniRegEx(ArabicSmallLowMeem) + "|" + MakeUniRegEx(ArabicSmallHighMeemIsolatedForm) + ")?(?:\s+|\s*$)|(" + MakeUniRegEx(ArabicLetterHamza) + ")" + MakeUniRegEx(ArabicFathatan) + "(?:" + MakeUniRegEx(ArabicSmallLowMeem) + "|" + MakeUniRegEx(ArabicSmallHighMeemIsolatedForm) + ")?" + MakeUniRegEx(ArabicLetterAlef) + "?(?:\s+|\s*$)|(?:" + MakeUniRegEx(ArabicLetterAlef) + "|" + MakeUniRegEx(ArabicLetterAlefMaksura) + "|" + MakeUniRegEx(ArabicLetterTehMarbuta) + "|" + MakeUniRegEx(ArabicLetterHamza) + ")" + MakeUniRegEx(ArabicFathatan) + "|" + MakeUniRegEx(ArabicFathatan) + "(?!" + MakeUniRegEx(ArabicSmallLowMeem) + "?(?:" + MakeUniRegEx(ArabicSmallHighMeemIsolatedForm) + "?" + MakeUniRegEx(ArabicLetterAlef) + "|" + MakeUniRegEx(ArabicLetterAlefMaksura) + ")(?:\s+|\s*$))", _
                                    .Evaluator = "$&", .NegativeMatch = "$1$2"}, _
        New RuleTranslation With {.Rule = "OnlyAtStartOfWord", .Match = "(?:(?:^\s*|\s+)(" + MakeUniRegEx(ArabicLetterWaw) + MakeUniRegEx(ArabicFatha) + "|(?:" + MakeUniRegEx(ArabicLetterWaw) + MakeUniRegEx(ArabicFatha) + ")?" + MakeUniRegEx(ArabicLetterTeh) + MakeUniRegEx(ArabicFatha) + "|(?:" + MakeUniRegEx(ArabicLetterWaw) + MakeUniRegEx(ArabicFatha) + "|" + MakeUniRegEx(ArabicLetterLam) + MakeUniRegEx(ArabicFatha) + "|" + MakeUniRegEx(ArabicLetterAlefWithHamzaAbove) + MakeUniRegEx(ArabicFatha) + "(" + MakeUniRegEx(ArabicLetterFeh) + MakeUniRegEx(ArabicFatha) + ")?)?" + MakeUniRegEx(ArabicLetterBeh) + MakeUniRegEx(ArabicKasra) + "|(?:" + MakeUniRegEx(ArabicLetterAlefWithHamzaAbove) + MakeUniRegEx(ArabicFatha) + ")?" + MakeUniRegEx(ArabicLetterFeh) + MakeUniRegEx(ArabicFatha) + "|" + MakeUniRegEx(ArabicLetterLam) + MakeUniRegEx(ArabicFatha) + "|" + MakeUniRegEx(ArabicLetterWaw) + MakeUniRegEx(ArabicShadda) + MakeUniRegEx(ArabicFatha) + "|" + MakeUniRegEx(ArabicLetterLam) + MakeUniRegEx(ArabicShadda) + MakeUniRegEx(ArabicFatha) + "|" + MakeUniRegEx(ArabicLetterKaf) + MakeUniRegEx(ArabicFatha) + "|" + MakeUniRegEx(ArabicLetterAlefWasla) + "?" + MakeUniRegEx(ArabicLetterLam) + MakeUniRegEx(ArabicKasra) + ")|\S+)" + MakeUniRegEx(ArabicLetterAlefWasla) + "|" + MakeUniRegEx(ArabicLetterAlefWasla) + "(?=" + MakeUniRegEx(ArabicLetterAlefWasla) + "|" + MakeUniRegEx(ArabicLetterAlefWithHamzaAbove) + "|" + MakeUniRegEx(ArabicLetterAlefWithHamzaBelow) + "|" + MakeUniRegEx(ArabicLetterAlefMaksura) + "|" + MakeUniRegEx(ArabicLetterAlef) + "|" + MakeUniRegEx(ArabicLetterSuperscriptAlef) + "|" + MakeUniRegEx(ArabicLetterHamza) + "|" + MakeUniRegEx(ArabicLetterTehMarbuta) + "|" + MakeUniRegEx(ArabicLetterWaw) + "|" + MakeUniRegEx(ArabicLetterYeh) + "|" + MakeUniRegEx(ArabicFatha) + "|" + MakeUniRegEx(ArabicKasra) + "|" + MakeUniRegEx(ArabicDamma) + "|" + MakeUniRegEx(ArabicSukun) + "|" + MakeUniRegEx(ArabicShadda) + "|" + MakeUniRegEx(ArabicMaddahAbove) + "|" + MakeUniRegEx(ArabicFathatan) + "|" + MakeUniRegEx(ArabicKasratan) + "|" + MakeUniRegEx(ArabicDammatan) + "|" + MakeUniRegEx(ArabicHamzaAbove) + ")", _
                                .Evaluator = "$&", .NegativeMatch = "$1"}, _
        New RuleTranslation With {.Rule = "NotValidCombination", .Match = "(" + MakeUniRegEx(ArabicLetterAlef) + "|" + MakeUniRegEx(ArabicLetterSuperscriptAlef) + "|" + MakeUniRegEx(ArabicLetterAlefMaksura) + "|" + MakeUniRegEx(ArabicSmallWaw) + "|" + MakeUniRegEx(ArabicSmallYeh) + "|" + MakeUniRegEx(ArabicLetterSeen) + "|" + MakeUniRegEx(ArabicLetterSad) + "|" + MakeUniRegEx(ArabicLetterAin) + "|" + MakeUniRegEx(ArabicLetterQaf) + "|" + MakeUniRegEx(ArabicLetterKaf) + "|" + MakeUniRegEx(ArabicLetterLam) + "|" + MakeUniRegEx(ArabicLetterMeem) + "|" + MakeUniRegEx(ArabicLetterNoon) + "|" + MakeUniRegEx(ArabicLetterWaw) + "|" + MakeUniRegEx(ArabicLetterYeh) + ")?" + MakeUniRegEx(ArabicMaddahAbove), _
                                .Evaluator = "$&", .NegativeMatch = "$1"}, _
        New RuleTranslation With {.Rule = "NotValidCombination", .Match = MakeUniRegEx(ArabicLetterAlef) + "(?:" + MakeUniRegEx(ArabicFatha) + "|" + MakeUniRegEx(ArabicKasra) + "|" + MakeUniRegEx(ArabicDamma) + "|" + MakeUniRegEx(ArabicSukun) + "|" + MakeUniRegEx(ArabicShadda) + ")|(?:" + MakeUniRegEx(ArabicLetterAlef) + "|" + MakeUniRegEx(ArabicLetterYeh) + ")" + MakeUniRegEx(ArabicDammatan) + "|(?:" + MakeUniRegEx(ArabicLetterAlef) + "|" + MakeUniRegEx(ArabicLetterWaw) + "|" + MakeUniRegEx(ArabicLetterAlefMaksura) + ")" + MakeUniRegEx(ArabicKasratan) + "|" + MakeUniRegEx(ArabicKasratan) + "(?!(?:" + MakeUniRegEx(ArabicSmallLowMeem) + "|" + MakeUniRegEx(ArabicSmallHighMeemIsolatedForm) + ")?(?:\s+|\s*$))|" + MakeUniRegEx(ArabicDammatan) + "(?!(?:" + MakeUniRegEx(ArabicSmallLowMeem) + "|" + MakeUniRegEx(ArabicSmallHighMeemIsolatedForm) + ")?(?:" + MakeUniRegEx(ArabicLetterAlef) + MakeUniRegEx(ArabicSmallHighRoundedZero) + "?)?(?:\s+|\s*$))", _
                                .Evaluator = "$&"}, _
        New RuleTranslation With {.Rule = "NotValidCombination", .Match = "(?:\S+(" + MakeUniRegEx(ArabicFatha) + "|" + MakeUniRegEx(ArabicKasra) + "|" + MakeUniRegEx(ArabicSukun) + "|" + MakeUniRegEx(ArabicMaddahAbove) + ")|\S+)" + MakeUniRegEx(ArabicLetterAlefWithHamzaBelow) + "|(?:\S+(" + MakeUniRegEx(ArabicFatha) + "|" + MakeUniRegEx(ArabicKasra) + "|" + MakeUniRegEx(ArabicSukun) + "|" + MakeUniRegEx(ArabicMaddahAbove) + ")|\S+)" + MakeUniRegEx(ArabicLetterAlefWithHamzaAbove) + "|(?:\S+(" + MakeUniRegEx(ArabicFatha) + "|" + MakeUniRegEx(ArabicKasra) + "|" + MakeUniRegEx(ArabicDamma) + "|" + MakeUniRegEx(ArabicSukun) + "|" + MakeUniRegEx(ArabicMaddahAbove) + ")|\S+)" + MakeUniRegEx(ArabicLetterHamza) + "|(" + MakeUniRegEx(ArabicFatha) + "|" + MakeUniRegEx(ArabicLetterAlefWasla) + "|" + MakeUniRegEx(ArabicDamma) + "|" + MakeUniRegEx(ArabicMaddahAbove) + ")?" + MakeUniRegEx(ArabicLetterWawWithHamzaAbove) + "|(" + MakeUniRegEx(ArabicFatha) + "|" + MakeUniRegEx(ArabicKasra) + "|" + MakeUniRegEx(ArabicDamma) + "|" + MakeUniRegEx(ArabicLetterAlef) + MakeUniRegEx(ArabicSmallHighRoundedZero) + "|" + MakeUniRegEx(ArabicLetterAlefWasla) + "|" + MakeUniRegEx(ArabicSukun) + "|" + MakeUniRegEx(ArabicMaddahAbove) + ")?" + MakeUniRegEx(ArabicLetterYehWithHamzaAbove) + "|" + MakeUniRegEx(ArabicLetterAlefWithHamzaBelow) + "(?!" + MakeUniRegEx(ArabicKasra) + "|" + MakeUniRegEx(ArabicKasratan) + ")|" + MakeUniRegEx(ArabicLetterAlefWithHamzaAbove) + "(?!" + MakeUniRegEx(ArabicSukun) + "|" + MakeUniRegEx(ArabicFatha) + "|" + MakeUniRegEx(ArabicDamma) + "|" + MakeUniRegEx(ArabicDammatan) + ")|" + MakeUniRegEx(ArabicLetterWawWithHamzaAbove) + "(?!" + MakeUniRegEx(ArabicSukun) + "|" + MakeUniRegEx(ArabicFatha) + "|" + MakeUniRegEx(ArabicDamma) + "|" + MakeUniRegEx(ArabicKasra) + "|" + MakeUniRegEx(ArabicFathatan) + "|" + MakeUniRegEx(ArabicDammatan) + ")|" + MakeUniRegEx(ArabicLetterYehWithHamzaAbove) + "(?!" + MakeUniRegEx(ArabicSukun) + "|" + MakeUniRegEx(ArabicFatha) + "|" + MakeUniRegEx(ArabicDamma) + "|" + MakeUniRegEx(ArabicKasra) + "|" + MakeUniRegEx(ArabicFathatan) + "|" + MakeUniRegEx(ArabicKasratan) + ")|" + MakeUniRegEx(ArabicLetterHamza) + "(?!" + MakeUniRegEx(ArabicSukun) + "|" + MakeUniRegEx(ArabicFatha) + "|" + MakeUniRegEx(ArabicDamma) + "|" + MakeUniRegEx(ArabicKasra) + "|" + MakeUniRegEx(ArabicFathatan) + "|" + MakeUniRegEx(ArabicKasratan) + "|" + MakeUniRegEx(ArabicDammatan) + ")", _
                                .Evaluator = "$&", .NegativeMatch = "$1$2$3$4$5"}, _
        New RuleTranslation With {.Rule = "NotValidCombination", .Match = "(" + MakeUniRegEx(ArabicLetterWaw) + "|" + MakeUniRegEx(ArabicLetterAlefMaksura) + "|" + MakeUniRegEx(ArabicFatha) + ")?" + MakeUniRegEx(ArabicLetterSuperscriptAlef), _
                                .Evaluator = "$&", .NegativeMatch = "$1"}, _
        New RuleTranslation With {.Rule = "NotValidCombination", .Match = MakeUniRegEx(ArabicLetterAlef) + MakeUniRegEx(ArabicLetterLam) + MakeUniRegEx(ArabicMaddahAbove) + "(" + MakeUniRegEx(ArabicLetterReh) + "|" + MakeUniRegEx(ArabicLetterMeem) + MakeUniRegEx(ArabicMaddahAbove) + "(?:" + MakeUniRegEx(ArabicLetterSad) + MakeUniRegEx(ArabicMaddahAbove) + "|" + MakeUniRegEx(ArabicLetterReh) + ")?)|(?:(" + MakeUniRegEx(ArabicLetterWaw) + "|" + MakeUniRegEx(ArabicSmallWaw) + "|" + MakeUniRegEx(ArabicDammatan) + MakeUniRegEx(ArabicSmallLowMeem) + "?|" + MakeUniRegEx(ArabicFathatan) + "(" + MakeUniRegEx(ArabicSmallHighMeemIsolatedForm) + "|" + MakeUniRegEx(ArabicSmallLowMeem) + ")?|" + MakeUniRegEx(ArabicMaddahAbove) + "|" + MakeUniRegEx(ArabicFatha) + "|" + MakeUniRegEx(ArabicKasra) + "|" + MakeUniRegEx(ArabicDamma) + "|" + MakeUniRegEx(ArabicLetterSuperscriptAlef) + "|" + MakeUniRegEx(ArabicSukun) + ")?|(?:^\s*|\s+))" + MakeUniRegEx(ArabicLetterAlef), _
                                .Evaluator = "$&", .NegativeMatch = "$1$2"}, _
        New RuleTranslation With {.Rule = "NotValidCombination", .Match = "(?:\S+(" + MakeUniRegEx(ArabicLetterAlef) + MakeUniRegEx(ArabicSmallHighRoundedZero) + "?|" + MakeUniRegEx(ArabicLetterYeh) + "|" + MakeUniRegEx(ArabicLetterHeh) + "|" + MakeUniRegEx(ArabicFatha) + "|" + MakeUniRegEx(ArabicKasra) + "|" + MakeUniRegEx(ArabicDamma) + "|" + MakeUniRegEx(ArabicLetterSuperscriptAlef) + "|" + MakeUniRegEx(ArabicSukun) + ")|\S+)" + MakeUniRegEx(ArabicLetterYeh), _
                                .Evaluator = "$&", .NegativeMatch = "$1"}, _
        New RuleTranslation With {.Rule = "NotValidCombination", .Match = "(?:\S+(" + MakeUniRegEx(ArabicLetterAlef) + "|" + MakeUniRegEx(ArabicFatha) + "|" + MakeUniRegEx(ArabicKasra) + "|" + MakeUniRegEx(ArabicDamma) + "|" + MakeUniRegEx(ArabicLetterSuperscriptAlef) + "|" + MakeUniRegEx(ArabicSukun) + ")|\S+)" + MakeUniRegEx(ArabicLetterWaw), _
                                .Evaluator = "$&", .NegativeMatch = "$1"}, _
        New RuleTranslation With {.Rule = "NotValidCombination", .Match = MakeUniRegEx(ArabicLetterGhain) + MakeUniRegEx(ArabicShadda) + "|" + MakeUniRegEx(ArabicLetterWaw) + "(?:" + MakeUniRegEx(ArabicLetterGhain) + "|" + MakeUniRegEx(ArabicLetterAlefMaksura) + "|" + MakeUniRegEx(ArabicLetterWaw) + "|" + MakeUniRegEx(ArabicLetterYeh) + ")|" + MakeUniRegEx(ArabicLetterYeh) + "(?:" + MakeUniRegEx(ArabicLetterAlef) + "|" + MakeUniRegEx(ArabicLetterAlefMaksura) + "|" + MakeUniRegEx(ArabicLetterWaw) + ")", _
                                .Evaluator = "$&"}, _
        New RuleTranslation With {.Rule = "NotValidCombination", .Match = "(" + MakeUniRegEx(ArabicFatha) + MakeUniRegEx(ArabicLetterSuperscriptAlef) + MakeUniRegEx(ArabicHamzaAbove) + MakeUniRegEx(ArabicSukun) + ")|(?:" + MakeUniRegEx(ArabicFatha) + "|" + MakeUniRegEx(ArabicKasra) + "|" + MakeUniRegEx(ArabicDamma) + "|" + MakeUniRegEx(ArabicFathatan) + "|" + MakeUniRegEx(ArabicKasratan) + "|" + MakeUniRegEx(ArabicDammatan) + "|" + MakeUniRegEx(ArabicLetterSuperscriptAlef) + "|" + MakeUniRegEx(ArabicSukun) + "|" + MakeUniRegEx(ArabicMaddahAbove) + "|" + MakeUniRegEx(ArabicHamzaAbove) + "|" + MakeUniRegEx(ArabicShadda) + "|" + MakeUniRegEx(ArabicTatweel) + ")" + MakeUniRegEx(ArabicShadda) + "|(?:" + MakeUniRegEx(ArabicFatha) + "|" + MakeUniRegEx(ArabicKasra) + "|" + MakeUniRegEx(ArabicDamma) + "|" + MakeUniRegEx(ArabicFathatan) + "|" + MakeUniRegEx(ArabicKasratan) + "|" + MakeUniRegEx(ArabicDammatan) + "|" + MakeUniRegEx(ArabicLetterSuperscriptAlef) + "|" + MakeUniRegEx(ArabicSukun) + "|" + MakeUniRegEx(ArabicMaddahAbove) + "|" + MakeUniRegEx(ArabicHamzaAbove) + "|" + MakeUniRegEx(ArabicTatweel) + ")(?:" + MakeUniRegEx(ArabicKasratan) + "|" + MakeUniRegEx(ArabicDammatan) + ")|(?:" + MakeUniRegEx(ArabicFatha) + "|" + MakeUniRegEx(ArabicKasra) + "|" + MakeUniRegEx(ArabicDamma) + "|" + MakeUniRegEx(ArabicFathatan) + "|" + MakeUniRegEx(ArabicKasratan) + "|" + MakeUniRegEx(ArabicDammatan) + "|" + MakeUniRegEx(ArabicLetterSuperscriptAlef) + "|" + MakeUniRegEx(ArabicSukun) + "|" + MakeUniRegEx(ArabicMaddahAbove) + "|" + MakeUniRegEx(ArabicTatweel) + ")" + MakeUniRegEx(ArabicFathatan) + "|(?:" + MakeUniRegEx(ArabicFatha) + "|" + MakeUniRegEx(ArabicKasra) + "|" + MakeUniRegEx(ArabicDamma) + "|" + MakeUniRegEx(ArabicFathatan) + "|" + MakeUniRegEx(ArabicKasratan) + "|" + MakeUniRegEx(ArabicDammatan) + "|" + MakeUniRegEx(ArabicSukun) + "|" + MakeUniRegEx(ArabicMaddahAbove) + "|" + MakeUniRegEx(ArabicHamzaAbove) + "|" + MakeUniRegEx(ArabicShadda) + "|" + MakeUniRegEx(ArabicTatweel) + ")" + MakeUniRegEx(ArabicMaddahAbove) + "|(?:" + MakeUniRegEx(ArabicFatha) + "|" + MakeUniRegEx(ArabicKasra) + "|" + MakeUniRegEx(ArabicDamma) + "|" + MakeUniRegEx(ArabicFathatan) + "|" + MakeUniRegEx(ArabicKasratan) + "|" + MakeUniRegEx(ArabicDammatan) + "|" + MakeUniRegEx(ArabicLetterSuperscriptAlef) + "|" + MakeUniRegEx(ArabicSukun) + "|" + MakeUniRegEx(ArabicMaddahAbove) + "|" + MakeUniRegEx(ArabicShadda) + "|" + MakeUniRegEx(ArabicTatweel) + ")" + MakeUniRegEx(ArabicSukun) + "|(?:" + MakeUniRegEx(ArabicFatha) + "|" + MakeUniRegEx(ArabicKasra) + "|" + MakeUniRegEx(ArabicDamma) + "|" + MakeUniRegEx(ArabicFathatan) + "|" + MakeUniRegEx(ArabicKasratan) + "|" + MakeUniRegEx(ArabicDammatan) + "|" + MakeUniRegEx(ArabicLetterSuperscriptAlef) + "|" + MakeUniRegEx(ArabicSukun) + "|" + MakeUniRegEx(ArabicMaddahAbove) + "|" + MakeUniRegEx(ArabicTatweel) + ")(?:" + MakeUniRegEx(ArabicKasra) + "|" + MakeUniRegEx(ArabicDamma) + "|" + MakeUniRegEx(ArabicFatha) + ")|(?:" + MakeUniRegEx(ArabicKasra) + "|" + MakeUniRegEx(ArabicDamma) + "|" + MakeUniRegEx(ArabicFathatan) + "|" + MakeUniRegEx(ArabicKasratan) + "|" + MakeUniRegEx(ArabicDammatan) + "|" + MakeUniRegEx(ArabicLetterSuperscriptAlef) + "|" + MakeUniRegEx(ArabicSukun) + "|" + MakeUniRegEx(ArabicMaddahAbove) + "|" + MakeUniRegEx(ArabicHamzaAbove) + "|" + MakeUniRegEx(ArabicShadda) + "|" + MakeUniRegEx(ArabicTatweel) + ")" + MakeUniRegEx(ArabicLetterSuperscriptAlef) + "|(?:" + MakeUniRegEx(ArabicFatha) + "|" + MakeUniRegEx(ArabicKasra) + "|" + MakeUniRegEx(ArabicDamma) + "|" + MakeUniRegEx(ArabicFathatan) + "|" + MakeUniRegEx(ArabicKasratan) + "|" + MakeUniRegEx(ArabicDammatan) + "|" + MakeUniRegEx(ArabicLetterSuperscriptAlef) + "|" + MakeUniRegEx(ArabicSukun) + "|" + MakeUniRegEx(ArabicMaddahAbove) + "|" + MakeUniRegEx(ArabicHamzaAbove) + "|" + MakeUniRegEx(ArabicShadda) + "|" + MakeUniRegEx(ArabicTatweel) + ")" + MakeUniRegEx(ArabicShadda) + "|(?:" + MakeUniRegEx(ArabicFatha) + "|" + MakeUniRegEx(ArabicKasra) + "|" + MakeUniRegEx(ArabicDamma) + "|" + MakeUniRegEx(ArabicFathatan) + "|" + MakeUniRegEx(ArabicKasratan) + "|" + MakeUniRegEx(ArabicDammatan) + "|" + MakeUniRegEx(ArabicLetterSuperscriptAlef) + "|" + MakeUniRegEx(ArabicSukun) + "|" + MakeUniRegEx(ArabicMaddahAbove) + "|" + MakeUniRegEx(ArabicHamzaAbove) + "|" + MakeUniRegEx(ArabicShadda) + ")" + MakeUniRegEx(ArabicHamzaAbove) + "|(?:" + MakeUniRegEx(ArabicFathatan) + "|" + MakeUniRegEx(ArabicKasratan) + "|" + MakeUniRegEx(ArabicDammatan) + "|" + MakeUniRegEx(ArabicLetterSuperscriptAlef) + "|" + MakeUniRegEx(ArabicHamzaAbove) + "|" + MakeUniRegEx(ArabicShadda) + "|" + MakeUniRegEx(ArabicTatweel) + ")" + MakeUniRegEx(ArabicTatweel), _
                                .Evaluator = "$&", .NegativeMatch = "$1"}, _
        New RuleTranslation With {.Rule = "MissingDiacritic", .Match = "(" + MakeUniRegEx(ArabicLetterKaf) + MakeUniRegEx(ArabicMaddahAbove) + MakeUniRegEx(ArabicLetterHeh) + MakeUniRegEx(ArabicLetterYeh) + MakeUniRegEx(ArabicLetterAin) + MakeUniRegEx(ArabicMaddahAbove) + MakeUniRegEx(ArabicLetterSad) + MakeUniRegEx(ArabicMaddahAbove) + "|" + MakeUniRegEx(ArabicLetterHah) + MakeUniRegEx(ArabicLetterMeem) + MakeUniRegEx(ArabicMaddahAbove) + "|" + MakeUniRegEx(ArabicLetterTah) + "(?:" + MakeUniRegEx(ArabicLetterHeh) + "|" + MakeUniRegEx(ArabicLetterSeen) + MakeUniRegEx(ArabicMaddahAbove) + "(?:" + MakeUniRegEx(ArabicLetterMeem) + MakeUniRegEx(ArabicMaddahAbove) + ")?)(?:\s*$|\s+)|(?:" + MakeUniRegEx(ArabicLetterHeh) + MakeUniRegEx(ArabicLetterHeh) + "|(?:" + MakeUniRegEx(ArabicLetterKaf) + "|" + MakeUniRegEx(ArabicLetterQaf) + ")" + MakeUniRegEx(ArabicLetterKaf) + "|" + MakeUniRegEx(ArabicLetterDal) + MakeUniRegEx(ArabicLetterTeh) + "|" + MakeUniRegEx(ArabicFatha) + MakeUniRegEx(ArabicLetterYeh) + MakeUniRegEx(ArabicLetterYeh) + "|" + MakeUniRegEx(ArabicLetterLam) + "(?:" + MakeRegMultiEx(Array.ConvertAll(ArabicSunLetters, Function(Str As String) MakeUniRegEx(Str))) + "))" + MakeUniRegEx(ArabicShadda) + "|" + MakeUniRegEx(ArabicLetterNoon) + "(?:" + MakeUniRegEx(ArabicLetterTeh) + "|" + MakeUniRegEx(ArabicLetterTheh) + "|" + MakeUniRegEx(ArabicLetterJeem) + "|" + MakeUniRegEx(ArabicLetterDal) + "|" + MakeUniRegEx(ArabicLetterThal) + "|" + MakeUniRegEx(ArabicLetterZain) + "|" + MakeUniRegEx(ArabicLetterSeen) + "|" + MakeUniRegEx(ArabicLetterSheen) + "|" + MakeUniRegEx(ArabicLetterSad) + "|" + MakeUniRegEx(ArabicLetterDad) + "|" + MakeUniRegEx(ArabicLetterTah) + "|" + MakeUniRegEx(ArabicLetterZah) + "|" + MakeUniRegEx(ArabicLetterFeh) + "|" + MakeUniRegEx(ArabicLetterQaf) + "|" + MakeUniRegEx(ArabicLetterKaf) + ")|" + MakeUniRegEx(ArabicLetterTah) + MakeUniRegEx(ArabicLetterTeh) + "|" + MakeUniRegEx(ArabicFatha) + MakeUniRegEx(ArabicLetterWaw) + MakeUniRegEx(ArabicLetterAlef) + ")|(?:" + MakeRegMultiEx(Array.ConvertAll(ArabicSunLetters, Function(Str As String) MakeUniRegEx(Str))) + "|" + MakeRegMultiEx(Array.ConvertAll(ArabicMoonLettersNoVowels, Function(Str As String) MakeUniRegEx(Str))) + "|(?:" + MakeUniRegEx(ArabicSukun) + "|" + MakeUniRegEx(ArabicFatha) + "|" + MakeUniRegEx(ArabicDamma) + ")" + MakeUniRegEx(ArabicLetterYeh) + "|(?:" + MakeUniRegEx(ArabicSukun) + "|" + MakeUniRegEx(ArabicFatha) + "|" + MakeUniRegEx(ArabicKasra) + ")" + MakeUniRegEx(ArabicLetterWaw) + ")(?:" + MakeRegMultiEx(Array.ConvertAll(ArabicSunLetters, Function(Str As String) MakeUniRegEx(Str))) + "|" + MakeRegMultiEx(Array.ConvertAll(ArabicMoonLetters, Function(Str As String) MakeUniRegEx(Str))) + ")|(?:" + MakeUniRegEx(ArabicLetterHeh) + MakeUniRegEx(ArabicSukun) + MakeUniRegEx(ArabicLetterHeh) + "|(?:" + MakeUniRegEx(ArabicLetterKaf) + "|" + MakeUniRegEx(ArabicLetterQaf) + ")" + MakeUniRegEx(ArabicSukun) + MakeUniRegEx(ArabicLetterKaf) + "|" + MakeUniRegEx(ArabicLetterDal) + MakeUniRegEx(ArabicSukun) + MakeUniRegEx(ArabicLetterTeh) + "|" + MakeUniRegEx(ArabicFatha) + MakeUniRegEx(ArabicLetterYeh) + MakeUniRegEx(ArabicSukun) + MakeUniRegEx(ArabicLetterYeh) + "|" + MakeUniRegEx(ArabicLetterLam) + MakeUniRegEx(ArabicSukun) + "(?:" + MakeRegMultiEx(Array.ConvertAll(ArabicSunLetters, Function(Str As String) MakeUniRegEx(Str))) + "))" + MakeUniRegEx(ArabicShadda) + "|" + MakeUniRegEx(ArabicLetterNoon) + MakeUniRegEx(ArabicSukun) + "(?:" + MakeUniRegEx(ArabicLetterTeh) + "|" + MakeUniRegEx(ArabicLetterTheh) + "|" + MakeUniRegEx(ArabicLetterJeem) + "|" + MakeUniRegEx(ArabicLetterDal) + "|" + MakeUniRegEx(ArabicLetterThal) + "|" + MakeUniRegEx(ArabicLetterZain) + "|" + MakeUniRegEx(ArabicLetterSeen) + "|" + MakeUniRegEx(ArabicLetterSheen) + "|" + MakeUniRegEx(ArabicLetterSad) + "|" + MakeUniRegEx(ArabicLetterDad) + "|" + MakeUniRegEx(ArabicLetterTah) + "|" + MakeUniRegEx(ArabicLetterZah) + "|" + MakeUniRegEx(ArabicLetterFeh) + "|" + MakeUniRegEx(ArabicLetterQaf) + "|" + MakeUniRegEx(ArabicLetterKaf) + ")|" + MakeUniRegEx(ArabicLetterTah) + MakeUniRegEx(ArabicSukun) + MakeUniRegEx(ArabicLetterTeh), _
                                .Evaluator = "$&", .NegativeMatch = "$1"}, _
        New RuleTranslation With {.Rule = "MissingEndOfWordDiacritic", .Match = "((?:" + MakeUniRegEx(ArabicFatha) + MakeUniRegEx(ArabicLetterWaw) + MakeUniRegEx(ArabicLetterAlef) + MakeUniRegEx(ArabicSmallHighRoundedZero) + "\s+" + MakeUniRegEx(ArabicLetterWaw) + "|" + MakeUniRegEx(ArabicLetterBeh) + "\s+(?:" + MakeUniRegEx(ArabicLetterBeh) + "|" + MakeUniRegEx(ArabicLetterMeem) + ")|" + MakeUniRegEx(ArabicFatha) + MakeUniRegEx(ArabicLetterWaw) + "\s+" + MakeUniRegEx(ArabicLetterWaw) + "|" + MakeUniRegEx(ArabicLetterThal) + "\s+(?:" + MakeUniRegEx(ArabicLetterThal) + "|" + MakeUniRegEx(ArabicLetterZah) + ")|" + MakeUniRegEx(ArabicLetterAin) + "\s+" + MakeUniRegEx(ArabicLetterAin) + "|" + MakeUniRegEx(ArabicLetterFeh) + "\s+" + MakeUniRegEx(ArabicLetterFeh) + "|" + MakeUniRegEx(ArabicLetterReh) + "\s+" + MakeUniRegEx(ArabicLetterReh) + "|" + MakeUniRegEx(ArabicLetterMeem) + "\s+" + MakeUniRegEx(ArabicLetterMeem) + "|" + MakeUniRegEx(ArabicLetterNoon) + "\s+(?:" + MakeUniRegEx(ArabicLetterNoon) + "|" + MakeUniRegEx(ArabicLetterLam) + "|" + MakeUniRegEx(ArabicLetterReh) + "|" + MakeUniRegEx(ArabicLetterMeem) + ")|" + MakeUniRegEx(ArabicLetterTeh) + "\s+(?:" + MakeUniRegEx(ArabicLetterTeh) + "|" + MakeUniRegEx(ArabicLetterDal) + "|" + MakeUniRegEx(ArabicLetterTah) + ")|" + MakeUniRegEx(ArabicLetterDal) + "\s+(?:" + MakeUniRegEx(ArabicLetterDal) + "|" + MakeUniRegEx(ArabicLetterTeh) + ")|" + MakeUniRegEx(ArabicLetterLam) + "\s+(?:" + MakeRegMultiEx(Array.ConvertAll(ArabicSunLetters, Function(Str As String) MakeUniRegEx(Str))) + "))" + MakeUniRegEx(ArabicShadda) + "|" + MakeUniRegEx(ArabicLetterMeem) + "\s+" + MakeUniRegEx(ArabicLetterBeh) + "|" + MakeUniRegEx(ArabicLetterNoon) + "\s+(?:" + MakeUniRegEx(ArabicLetterTeh) + "|" + MakeUniRegEx(ArabicLetterTheh) + "|" + MakeUniRegEx(ArabicLetterJeem) + "|" + MakeUniRegEx(ArabicLetterDal) + "|" + MakeUniRegEx(ArabicLetterThal) + "|" + MakeUniRegEx(ArabicLetterZain) + "|" + MakeUniRegEx(ArabicLetterSeen) + "|" + MakeUniRegEx(ArabicLetterSheen) + "|" + MakeUniRegEx(ArabicLetterSad) + "|" + MakeUniRegEx(ArabicLetterDad) + "|" + MakeUniRegEx(ArabicLetterTah) + "|" + MakeUniRegEx(ArabicLetterZah) + "|" + MakeUniRegEx(ArabicLetterFeh) + "|" + MakeUniRegEx(ArabicLetterQaf) + "|" + MakeUniRegEx(ArabicLetterKaf) + "|" + MakeUniRegEx(ArabicLetterWaw) + "|" + MakeUniRegEx(ArabicLetterYeh) + "))|(?:" + MakeRegMultiEx(Array.ConvertAll(ArabicSunLetters, Function(Str As String) MakeUniRegEx(Str))) + "|" + MakeRegMultiEx(Array.ConvertAll(ArabicMoonLettersNoVowels, Function(Str As String) MakeUniRegEx(Str))) + "|(?:" + MakeUniRegEx(ArabicSukun) + "|" + MakeUniRegEx(ArabicFatha) + "|" + MakeUniRegEx(ArabicDamma) + ")" + MakeUniRegEx(ArabicLetterYeh) + "|(?:" + MakeUniRegEx(ArabicSukun) + "|" + MakeUniRegEx(ArabicFatha) + "|" + MakeUniRegEx(ArabicKasra) + ")" + MakeUniRegEx(ArabicLetterWaw) + ")\s+(?:$|" + MakeRegMultiEx(Array.ConvertAll(ArabicSunLetters, Function(Str As String) MakeUniRegEx(Str))) + "|" + MakeRegMultiEx(Array.ConvertAll(ArabicMoonLetters, Function(Str As String) MakeUniRegEx(Str))) + ")|" + MakeUniRegEx(ArabicSukun) + "\s+" + MakeUniRegEx(ArabicLetterAlefWasla) + "|(?:" + MakeUniRegEx(ArabicFatha) + MakeUniRegEx(ArabicLetterWaw) + MakeUniRegEx(ArabicSukun) + MakeUniRegEx(ArabicLetterAlef) + MakeUniRegEx(ArabicSmallHighRoundedZero) + "\s+" + MakeUniRegEx(ArabicLetterWaw) + "|" + MakeUniRegEx(ArabicLetterBeh) + MakeUniRegEx(ArabicSukun) + "\s+(?:" + MakeUniRegEx(ArabicLetterBeh) + "|" + MakeUniRegEx(ArabicLetterMeem) + ")|" + MakeUniRegEx(ArabicFatha) + MakeUniRegEx(ArabicLetterWaw) + MakeUniRegEx(ArabicSukun) + "\s+" + MakeUniRegEx(ArabicLetterWaw) + "|" + MakeUniRegEx(ArabicLetterThal) + MakeUniRegEx(ArabicSukun) + "\s+(?:" + MakeUniRegEx(ArabicLetterThal) + "|" + MakeUniRegEx(ArabicLetterZah) + ")|" + MakeUniRegEx(ArabicLetterAin) + MakeUniRegEx(ArabicSukun) + "\s+" + MakeUniRegEx(ArabicLetterAin) + "|" + MakeUniRegEx(ArabicLetterFeh) + MakeUniRegEx(ArabicSukun) + "\s+" + MakeUniRegEx(ArabicLetterFeh) + "|" + MakeUniRegEx(ArabicLetterReh) + MakeUniRegEx(ArabicSukun) + "\s+" + MakeUniRegEx(ArabicLetterReh) + "|" + MakeUniRegEx(ArabicLetterMeem) + MakeUniRegEx(ArabicSukun) + "\s+" + MakeUniRegEx(ArabicLetterMeem) + "|" + MakeUniRegEx(ArabicLetterNoon) + MakeUniRegEx(ArabicSukun) + "\s+(?:" + MakeUniRegEx(ArabicLetterNoon) + "|" + MakeUniRegEx(ArabicLetterLam) + "|" + MakeUniRegEx(ArabicLetterReh) + "|" + MakeUniRegEx(ArabicLetterMeem) + ")|" + MakeUniRegEx(ArabicLetterTeh) + MakeUniRegEx(ArabicSukun) + "\s+(?:" + MakeUniRegEx(ArabicLetterTeh) + "|" + MakeUniRegEx(ArabicLetterDal) + "|" + MakeUniRegEx(ArabicLetterTah) + ")|" + MakeUniRegEx(ArabicLetterDal) + MakeUniRegEx(ArabicSukun) + "\s+(?:" + MakeUniRegEx(ArabicLetterDal) + "|" + MakeUniRegEx(ArabicLetterTeh) + ")|" + MakeUniRegEx(ArabicLetterLam) + MakeUniRegEx(ArabicSukun) + "\s+(?:" + MakeRegMultiEx(Array.ConvertAll(ArabicSunLetters, Function(Str As String) MakeUniRegEx(Str))) + "))" + MakeUniRegEx(ArabicShadda) + "|" + MakeUniRegEx(ArabicLetterMeem) + MakeUniRegEx(ArabicSukun) + "\s+" + MakeUniRegEx(ArabicLetterBeh) + "|" + MakeUniRegEx(ArabicLetterNoon) + MakeUniRegEx(ArabicSukun) + "\s+(?:" + MakeUniRegEx(ArabicLetterTeh) + "|" + MakeUniRegEx(ArabicLetterTheh) + "|" + MakeUniRegEx(ArabicLetterJeem) + "|" + MakeUniRegEx(ArabicLetterDal) + "|" + MakeUniRegEx(ArabicLetterThal) + "|" + MakeUniRegEx(ArabicLetterZain) + "|" + MakeUniRegEx(ArabicLetterSeen) + "|" + MakeUniRegEx(ArabicLetterSheen) + "|" + MakeUniRegEx(ArabicLetterSad) + "|" + MakeUniRegEx(ArabicLetterDad) + "|" + MakeUniRegEx(ArabicLetterTah) + "|" + MakeUniRegEx(ArabicLetterZah) + "|" + MakeUniRegEx(ArabicLetterFeh) + "|" + MakeUniRegEx(ArabicLetterQaf) + "|" + MakeUniRegEx(ArabicLetterKaf) + "|" + MakeUniRegEx(ArabicLetterWaw) + "|" + MakeUniRegEx(ArabicLetterYeh) + ")", _
                                .Evaluator = "$&", .NegativeMatch = "$1"}, _
        New RuleTranslation With {.Rule = "NotValidStartEnd", .Match = "(?:^\s*|\s+)(?:" + MakeUniRegEx(ArabicFatha) + "|" + MakeUniRegEx(ArabicKasra) + "|" + MakeUniRegEx(ArabicDamma) + "|" + MakeUniRegEx(ArabicFathatan) + "|" + MakeUniRegEx(ArabicKasratan) + "|" + MakeUniRegEx(ArabicDammatan) + "|" + MakeUniRegEx(ArabicSukun) + "|" + MakeUniRegEx(ArabicMaddahAbove) + "|" + MakeUniRegEx(ArabicHamzaAbove) + "|" + MakeUniRegEx(ArabicShadda) + ")|(?:" + MakeUniRegEx(ArabicLetterJeem) + "|" + MakeUniRegEx(ArabicLetterHah) + "|" + MakeUniRegEx(ArabicLetterKhah) + "|" + MakeUniRegEx(ArabicLetterZain) + "|" + MakeUniRegEx(ArabicLetterSeen) + "|" + MakeUniRegEx(ArabicLetterSheen) + "|" + MakeUniRegEx(ArabicLetterSad) + "|" + MakeUniRegEx(ArabicLetterDad) + "|" + MakeUniRegEx(ArabicLetterTah) + "|" + MakeUniRegEx(ArabicLetterGhain) + ")(?=\s*$|\s+)", _
                                .Evaluator = "$&"}}
    'New RuleTranslation With {.Rule = "NeedsRecomposition", .Match = "(" + MakeUniRegEx(ArabicLetterAlef) + MakeUniRegEx(ArabicMaddahAbove) + "|" + MakeUniRegEx(ArabicLetterHamza) + MakeUniRegEx(ArabicFatha) + MakeUniRegEx(ArabicLetterAlef) + "|" + MakeUniRegEx(ArabicLetterAlef) + MakeUniRegEx(ArabicLetterHamza) + "|" + MakeUniRegEx(ArabicLetterWaw) + MakeUniRegEx(ArabicLetterHamza) + "|" + MakeUniRegEx(ArabicLetterAlefMaksura) + MakeUniRegEx(ArabicLetterHamza) + ")", _
    '                            .Evaluator = "$&"}
    '} '"Missing diacritic", "Can only appear at end of word", "Must not appear at beginning of word", "Must appear at beginning of word", "Not a valid combination", "Needs to be recomposed"

    Public Enum RuleFuncs As Integer
        eNone
        eUpperCase
        eSpellNumber
        eSpellLetter
        eLookupLetter
        eDivideTanween
        eDivideLetterSymbol
        eStopOption
    End Enum
    Public Delegate Function RuleFunction(Str As String) As String()
    Public Shared RuleFunctions As RuleFunction() = {
        Function(Str As String) {UCase(Str)},
        Function(Str As String) {TransliterateWithRules(TransliterateFromBuckwalter(ArabicWordFromNumber(CInt(TransliterateToScheme(Str, TranslitScheme.Literal, String.Empty)), True, False, False)), String.Empty)},
        Function(Str As String) {ArabicLetterSpelling(Str)},
        Function(Str As String) {Arabic.GetSchemeValueFromSymbol(CachedData.IslamData.ArabicLetters(FindLetterBySymbol(Str)), String.Empty)},
        Function(Str As String) {ArabicFathaDammaKasra(Array.IndexOf(ArabicTanweens, Str)), ArabicLetterNoon},
        Function(Str As String) {String.Empty, String.Empty}
    }
        'Javascript does not support negative or positive lookbehind in regular expressions
    Public Shared RomanizationRules As RuleTranslation() = { _
        New RuleTranslation With {.Rule = "Shadda", .Match = "(" + MakeRegMultiEx(Array.ConvertAll(ArabicLetters, Function(Str As String) MakeUniRegEx(Str))) + ")" + MakeUniRegEx(ArabicShadda), _
                                    .Evaluator = "$1-$1"}, _
        New RuleTranslation With {.Rule = "GutteralRules", .Match = "(" + MakeRegMultiEx(Array.ConvertAll(ArabicSpecialGutteral, Function(Str As String) MakeUniRegEx(Str))) + "|" + MakeRegMultiEx(Array.ConvertAll(ArabicSpecialLeadingGutteral, Function(Str As String) MakeUniRegEx(Str))) + ")(?:" + MakeUniRegEx(ArabicFatha) + MakeUniRegEx(ArabicLetterAlef) + ")", _
                                    .Evaluator = "$1waa"}, _
        New RuleTranslation With {.Rule = "GutteralRules", .Match = "(" + MakeRegMultiEx(Array.ConvertAll(ArabicSpecialGutteral, Function(Str As String) MakeUniRegEx(Str))) + "|" + MakeRegMultiEx(Array.ConvertAll(ArabicSpecialLeadingGutteral, Function(Str As String) MakeUniRegEx(Str))) + ")(?:" + MakeUniRegEx(ArabicFatha) + ")", _
                                    .Evaluator = "$1aw"}, _
        New RuleTranslation With {.Rule = "GutteralRules", .Match = "(" + MakeRegMultiEx(Array.ConvertAll(ArabicSpecialGutteral, Function(Str As String) MakeUniRegEx(Str))) + "|" + MakeRegMultiEx(Array.ConvertAll(ArabicSpecialLeadingGutteral, Function(Str As String) MakeUniRegEx(Str))) + ")(?:" + MakeUniRegEx(ArabicDamma) + MakeUniRegEx(ArabicLetterWaw) + ")", _
                                    .Evaluator = "$1eoo"}, _
        New RuleTranslation With {.Rule = "GutteralRules", .Match = "(" + MakeRegMultiEx(Array.ConvertAll(ArabicSpecialGutteral, Function(Str As String) MakeUniRegEx(Str))) + "|" + MakeRegMultiEx(Array.ConvertAll(ArabicSpecialLeadingGutteral, Function(Str As String) MakeUniRegEx(Str))) + ")(?:" + MakeUniRegEx(ArabicDamma) + ")", _
                                    .Evaluator = "$1o"}, _
        New RuleTranslation With {.Rule = "GutteralRules", .Match = "(" + MakeRegMultiEx(Array.ConvertAll(ArabicSpecialGutteral, Function(Str As String) MakeUniRegEx(Str))) + "|" + MakeRegMultiEx(Array.ConvertAll(ArabicSpecialLeadingGutteral, Function(Str As String) MakeUniRegEx(Str))) + ")(?:" + MakeUniRegEx(ArabicKasra) + MakeUniRegEx(ArabicLetterYeh) + ")", _
                                    .Evaluator = "$1aee"}, _
        New RuleTranslation With {.Rule = "GutteralRules", .Match = "(" + MakeRegMultiEx(Array.ConvertAll(ArabicSpecialGutteral, Function(Str As String) MakeUniRegEx(Str))) + "|" + MakeRegMultiEx(Array.ConvertAll(ArabicSpecialLeadingGutteral, Function(Str As String) MakeUniRegEx(Str))) + ")(?:" + MakeUniRegEx(ArabicKasra) + ")", _
                                    .Evaluator = "$1e"}, _
        New RuleTranslation With {.Rule = "GutteralRules", .Match = "(" + MakeRegMultiEx(Array.ConvertAll(ArabicSpecialGutteral, Function(Str As String) MakeUniRegEx(Str))) + "|" + MakeRegMultiEx(Array.ConvertAll(ArabicSpecialLeadingGutteral, Function(Str As String) MakeUniRegEx(Str))) + ")(?:" + MakeUniRegEx(ArabicFathatan) + ")", _
                                    .Evaluator = "$1awn"}, _
        New RuleTranslation With {.Rule = "GutteralRules", .Match = "(" + MakeRegMultiEx(Array.ConvertAll(ArabicSpecialGutteral, Function(Str As String) MakeUniRegEx(Str))) + "|" + MakeRegMultiEx(Array.ConvertAll(ArabicSpecialLeadingGutteral, Function(Str As String) MakeUniRegEx(Str))) + ")(?:" + MakeUniRegEx(ArabicDammatan) + ")", _
                                    .Evaluator = "$1on"}, _
        New RuleTranslation With {.Rule = "GutteralRules", .Match = "(" + MakeRegMultiEx(Array.ConvertAll(ArabicSpecialGutteral, Function(Str As String) MakeUniRegEx(Str))) + "|" + MakeRegMultiEx(Array.ConvertAll(ArabicSpecialLeadingGutteral, Function(Str As String) MakeUniRegEx(Str))) + ")(?:" + MakeUniRegEx(ArabicKasratan) + ")", _
                                    .Evaluator = "$1en"}, _
        New RuleTranslation With {.Rule = "LeadingGutteralRules", .Match = "(?:" + MakeUniRegEx(ArabicFatha) + MakeUniRegEx(ArabicLetterAlef) + ")(" + MakeUniRegEx(ArabicLetterHamza) + "|" + MakeRegMultiEx(Array.ConvertAll(ArabicSpecialLeadingGutteral, Function(Str As String) MakeUniRegEx(Str))) + ")", _
                                    .Evaluator = "aae$1"}, _
        New RuleTranslation With {.Rule = "LeadingGutteralRules", .Match = "(?:" + MakeUniRegEx(ArabicFatha) + ")(" + MakeUniRegEx(ArabicLetterHamza) + "|" + MakeRegMultiEx(Array.ConvertAll(ArabicSpecialLeadingGutteral, Function(Str As String) MakeUniRegEx(Str))) + ")", _
                                    .Evaluator = "ah$1"}, _
        New RuleTranslation With {.Rule = "LeadingGutteralRules", .Match = "(?:" + MakeUniRegEx(ArabicKasra) + MakeUniRegEx(ArabicLetterYeh) + ")(" + MakeUniRegEx(ArabicLetterHamza) + "|" + MakeRegMultiEx(Array.ConvertAll(ArabicSpecialLeadingGutteral, Function(Str As String) MakeUniRegEx(Str))) + ")", _
                                    .Evaluator = "eeh$1"}, _
        New RuleTranslation With {.Rule = "LeadingGutteralRules", .Match = "(?:" + MakeUniRegEx(ArabicKasra) + ")(" + MakeUniRegEx(ArabicLetterHamza) + "|" + MakeRegMultiEx(Array.ConvertAll(ArabicSpecialLeadingGutteral, Function(Str As String) MakeUniRegEx(Str))) + ")", _
                                    .Evaluator = "ik$1"}, _
        New RuleTranslation With {.Rule = "LeadingGutteralRules", .Match = "(?:" + MakeUniRegEx(ArabicDamma) + MakeUniRegEx(ArabicLetterWaw) + ")(" + MakeUniRegEx(ArabicLetterHamza) + "|" + MakeRegMultiEx(Array.ConvertAll(ArabicSpecialLeadingGutteral, Function(Str As String) MakeUniRegEx(Str))) + ")", _
                                    .Evaluator = "ooh$1"}, _
        New RuleTranslation With {.Rule = "LeadingGutteralRules", .Match = "(?:" + MakeUniRegEx(ArabicDamma) + ")(" + MakeUniRegEx(ArabicLetterHamza) + "|" + MakeRegMultiEx(Array.ConvertAll(ArabicSpecialLeadingGutteral, Function(Str As String) MakeUniRegEx(Str))) + ")", _
                                    .Evaluator = "o$1"}, _
        New RuleTranslation With {.Rule = "FathaAlef", .Match = MakeUniRegEx(ArabicFatha) + MakeUniRegEx(ArabicLetterAlef), _
                                    .Evaluator = "aa"}, _
        New RuleTranslation With {.Rule = "DammaWaw", .Match = MakeUniRegEx(ArabicDamma) + MakeUniRegEx(ArabicLetterWaw), _
                                    .Evaluator = "oo"}, _
        New RuleTranslation With {.Rule = "KasraYeh", .Match = MakeUniRegEx(ArabicKasra) + MakeUniRegEx(ArabicLetterYeh), _
                                    .Evaluator = "ee"}, _
        New RuleTranslation With {.Rule = "ResolveAmbiguity", .Match = "(" + MakeUniRegEx(ArabicLetterSeen) + "|" + MakeUniRegEx(ArabicLetterTeh) + "|" + MakeUniRegEx(ArabicLetterTah) + ")(" + MakeUniRegEx(ArabicLetterHeh) + ")", _
                                  .Evaluator = "$1-$2"}, _
        New RuleTranslation With {.Rule = "LettersTanweensVowelsHamza", .Match = "(" + MakeRegMultiEx(Array.ConvertAll(ArabicLetters, Function(Str As String) MakeUniRegEx(Str))) + "|" + MakeRegMultiEx(Array.ConvertAll(ArabicTanweens, Function(Str As String) MakeUniRegEx(Str))) + "|" + MakeRegMultiEx(Array.ConvertAll(ArabicFathaDammaKasra, Function(Str As String) MakeUniRegEx(Str))) + "|" + MakeUniRegEx(ArabicLetterHamza) + ")", _
                                    .Evaluator = "$&", .RuleFunc = RuleFuncs.eLookupLetter}, _
        New RuleTranslation With {.Rule = "Punctuation", .Match = "(" + MakeRegMultiEx(Array.ConvertAll(ArabicPunctuationSymbols, Function(Str As String) MakeUniRegEx(Str))) + ")", _
                                    .Evaluator = "$&", .RuleFunc = RuleFuncs.eLookupLetter}
    }
    Public Shared AllowZeroLength As String() = {"helperlparen", "helperrparen"}
    Public Shared ColoringSpelledOutRules As RuleTranslation() = { _
        New RuleTranslation With {.Rule = "DivideTanween", .Match = "dividetanween",
                                  .Evaluator = "{0}", .RuleFunc = RuleFuncs.eDivideTanween}, _
        New RuleTranslation With {.Rule = "DivideLetterSymbol", .Match = "dividelettersymbol",
                                  .Evaluator = "{0}", .RuleFunc = RuleFuncs.eDivideLetterSymbol}, _
        New RuleTranslation With {.Rule = "LeftParenthesis", .Match = "helperlparen",
                                  .Evaluator = "("}, _
        New RuleTranslation With {.Rule = "RightParenthesis", .Match = "helperrparen",
                                  .Evaluator = ")"}, _
        New RuleTranslation With {.Rule = "HelperHeh", .Match = "helperheh",
                                  .Evaluator = ArabicLetterHeh}, _
        New RuleTranslation With {.Rule = "HelperTeh", .Match = "helperteh",
                                  .Evaluator = ArabicLetterTeh}, _
        New RuleTranslation With {.Rule = "LetterSpelling", .Match = "spellletter", _
                                    .Evaluator = "{0}", .RuleFunc = RuleFuncs.eSpellLetter}, _
        New RuleTranslation With {.Rule = "NumberSpelling", .Match = "spellnumber", _
                                    .Evaluator = "{0}", .RuleFunc = RuleFuncs.eSpellNumber}, _
        New RuleTranslation With {.Rule = "HelperPrefix", .Match = "helperprefix", _
                                    .Evaluator = "{0}-"}, _
        New RuleTranslation With {.Rule = "HelperFatha", .Match = "helperfatha", _
                                    .Evaluator = ArabicFatha}, _
        New RuleTranslation With {.Rule = "HelperKasra", .Match = "helperkasra", _
                                    .Evaluator = ArabicKasra}, _
        New RuleTranslation With {.Rule = "HelperDamma", .Match = "helperdamma", _
                                    .Evaluator = ArabicDamma}, _
        New RuleTranslation With {.Rule = "HelperAlef", .Match = "helperalef", _
                                    .Evaluator = ArabicLetterAlef}, _
        New RuleTranslation With {.Rule = "HelperWaw", .Match = "helperwaw", _
                                    .Evaluator = ArabicLetterWaw}, _
        New RuleTranslation With {.Rule = "HelperYeh", .Match = "helperyeh", _
                                    .Evaluator = ArabicLetterYeh}, _
        New RuleTranslation With {.Rule = "HelperMeem", .Match = "helpermeem", _
                                    .Evaluator = ArabicLetterMeem}, _
        New RuleTranslation With {.Rule = "HelperNoon", .Match = "helpernoon", _
                                    .Evaluator = ArabicLetterNoon}, _
        New RuleTranslation With {.Rule = "HelperSeen", .Match = "helperseen", _
                                    .Evaluator = ArabicLetterSeen}, _
        New RuleTranslation With {.Rule = "HelperMadda", .Match = "helpermadda", _
                                    .Evaluator = ArabicLetterHamza + ArabicFatha + ArabicLetterAlef}, _
        New RuleTranslation With {.Rule = "Empty", .Match = "empty|compulsorystop|startofhizb|endofversestop|prostration|canstoporcontinue|prohibitedtostop|stopatfirstnotsecond|stopatsecondnotfirst|bettertostopbutpermissibletocontinue|bettertocontinuebutpermissibletostop|subtlestopwithoutbreath", _
                                    .Evaluator = String.Empty}, _
        New RuleTranslation With {.Rule = "Normal", .Match = "helpermeem|assimilator|assimilatorincomplete|dipthong", .Evaluator = "{0}"}, _
        New RuleTranslation With {.Rule = "NecessaryProlongation", .Match = "necessaryprolong", .Evaluator = "{0}-{0}-{0}-{0}-{0}-{0}"}, _
        New RuleTranslation With {.Rule = "ObligatoryProlongation", .Match = "obligatoryprolong", .Evaluator = "{0}-{0}-{0}-{0}(-{0})"}, _
        New RuleTranslation With {.Rule = "PermissibleProlongation", .Match = "permissibleprolong", .Evaluator = "{0}-{0}(-{0}-{0})(-{0}-{0})"}, _
        New RuleTranslation With {.Rule = "NormalProlongation", .Match = "normalprolong", .Evaluator = "{0}-{0}"}, _
        New RuleTranslation With {.Rule = "Nasalization", .Match = "nasalize", .Evaluator = "{0}"}, _
        New RuleTranslation With {.Rule = "Unannounced", .Match = "assimilate|assimilateincomplete", .Evaluator = String.Empty}, _
        New RuleTranslation With {.Rule = "EmphaticPronounciation", .Match = "emphasis", .Evaluator = "{0}", .RuleFunc = RuleFuncs.eUpperCase}, _
        New RuleTranslation With {.Rule = "UnrestLetters", .Match = "bounce", .Evaluator = "{0}-{0}"} _
    }
    Public Shared PrefixPattern As String = ""
    Public Shared CertainStopPattern As String = MakeUniRegEx(ArabicSmallHighMeemInitialForm) + "|" + MakeRegMultiEx(Array.ConvertAll(PunctuationSymbols, Function(Ch As Char) MakeUniRegEx(Ch)))
    Public Shared CertainNotStopPattern As String = MakeRegMultiEx(Array.ConvertAll(RecitationSymbols, Function(Str As String) MakeUniRegEx(Str)))
    Public Shared OptionalStopPattern As String = MakeUniRegEx(ArabicEndOfAyah) + "|" + MakeUniRegEx(ArabicSmallHighJeem) + "|" + MakeUniRegEx(ArabicSmallHighThreeDots) + "|" + MakeUniRegEx(ArabicSmallHighLigatureQafWithLamWithAlefMaksura) + "|" + MakeUniRegEx(ArabicSmallHighLigatureSadWithLamWithAlefMaksura)
    Public Shared OptionalNotStopPattern As String = String.Empty
    Public Shared RulesOfRecitationRegEx As RuleMetadataTranslation() = { _
        New RuleMetadataTranslation With {.Rule = "LetterSpelling", .Match = "(^\s*|\s+)(" + MakeRegMultiEx(Array.ConvertAll(ArabicUniqueLetters, Function(Str As String) MakeUniRegEx(TransliterateFromBuckwalter(Str)))) + ")(?=\s*$|\s+)", _
            .Evaluator = New String() {Nothing, "spellletter"}}, _
        New RuleMetadataTranslation With {.Rule = "NumberSpelling", .Match = "(^\s*|\s+)((?:" + MakeRegMultiEx(Array.ConvertAll(ArabicNumbers, Function(Str As String) MakeUniRegEx(TransliterateFromBuckwalter(Str)))) + ")+)(?=\s*$|\s+)", _
            .Evaluator = New String() {Nothing, "spellnumber"}}, _
        New RuleMetadataTranslation With {.Rule = "NumberSpelling", .Match = "(" + MakeUniRegEx(ArabicEndOfAyah) + ")()((?:" + MakeRegMultiEx(Array.ConvertAll(ArabicNumbers, Function(Str As String) MakeUniRegEx(TransliterateFromBuckwalter(Str)))) + ")+)()(?=\s*$|\s+)", _
            .Evaluator = New String() {Nothing, "helperlparen", "spellnumber", "helperrparen"}}, _
        New RuleMetadataTranslation With {.Rule = "Stopping", .Match = "(?:(" + MakeUniRegEx(ArabicSmallHighMeemInitialForm) + ")|(" + MakeUniRegEx(ArabicStartOfRubElHizb) + ")|(" + MakeUniRegEx(ArabicEndOfAyah) + ")|(" + MakeUniRegEx(ArabicPlaceOfSajdah) + "))(?=\s*$|\s+)",
            .Evaluator = New String() {"compulsorystop", "startofhizb", "endofversestop", "prostration"}}, _
        New RuleMetadataTranslation With {.Rule = "Stopping", .Match = "(" + MakeUniRegEx(ArabicSmallHighJeem) + ")|(" + MakeUniRegEx(ArabicSmallHighLamAlef) + ")|(" + MakeUniRegEx(ArabicSmallHighThreeDots) + ")(.*)(" + MakeUniRegEx(ArabicSmallHighThreeDots) + ")|(" + MakeUniRegEx(ArabicSmallHighLigatureQafWithLamWithAlefMaksura) + ")|(" + MakeUniRegEx(ArabicSmallHighLigatureSadWithLamWithAlefMaksura) + ")|(" + MakeUniRegEx(ArabicSmallHighSeen) + ")(?:\s+|\s*$)",
            .Evaluator = New String() {"canstoporcontinue", "prohibitedtostop", "stopatfirstnotsecond", Nothing, "stopatsecondnotfirst", "bettertostopbutpermissibletocontinue", "bettertocontinuebutpermissibletostop", "subtlestopwithoutbreath"}}, _
        New RuleMetadataTranslation With {.Rule = "Stopping", .Match = "(" + MakeUniRegEx(ArabicFathatan) + ")(?=" + MakeUniRegEx(ArabicLetterAlef) + "\s*($|" + CertainStopPattern + "|" + OptionalStopPattern + "))",
            .Evaluator = New String() {"helperfatha"}}, _
        New RuleMetadataTranslation With {.Rule = "Stopping", .Match = "([^" + MakeUniRegEx(ArabicLetterTehMarbuta) + "])(" + MakeRegMultiEx(Array.ConvertAll(ArabicFathaDammaKasra, Function(Str As String) MakeUniRegEx(Str))) + "|" + MakeUniRegEx(ArabicKasratan) + "|" + MakeUniRegEx(ArabicDammatan) + ")(?=\s*($|" + CertainStopPattern + "|" + OptionalStopPattern + "))",
            .Evaluator = New String() {Nothing, "empty"}}, _
        New RuleMetadataTranslation With {.Rule = "Stopping", .Match = "(" + MakeUniRegEx(ArabicLetterTehMarbuta) + ")(" + MakeRegMultiEx(Array.ConvertAll(ArabicFathaDammaKasra, Function(Str As String) MakeUniRegEx(Str))) + "|" + MakeRegMultiEx(Array.ConvertAll(ArabicTanweens, Function(Str As String) MakeUniRegEx(Str))) + ")(?=\s*($|" + CertainStopPattern + "|" + OptionalStopPattern + "))",
            .Evaluator = New String() {"helperheh", "empty"}}, _
        New RuleMetadataTranslation With {.Rule = "Continuing", .Match = "(" + MakeUniRegEx(ArabicLetterTehMarbuta) + ")(?=(" + MakeRegMultiEx(Array.ConvertAll(ArabicFathaDammaKasra, Function(Str As String) MakeUniRegEx(Str))) + "|" + MakeRegMultiEx(Array.ConvertAll(ArabicTanweens, Function(Str As String) MakeUniRegEx(Str))) + ")\s*\S)",
            .Evaluator = New String() {"helperteh"}}, _
        New RuleMetadataTranslation With {.Rule = "SmallBounce", .Match = "(" + MakeRegMultiEx(Array.ConvertAll(ArabicFathaDammaKasra, Function(Str As String) MakeUniRegEx(Str))) + ")(" + MakeUniRegEx(ArabicLetterDal) + "|" + MakeUniRegEx(ArabicLetterBeh) + "|" + MakeUniRegEx(ArabicLetterQaf) + "|" + MakeUniRegEx(ArabicLetterTah) + "|" + MakeUniRegEx(ArabicLetterJeem) + ")(?=" + MakeUniRegEx(ArabicSukun) + "\S)", _
            .Evaluator = New String() {Nothing, "bounce"}}, _
        New RuleMetadataTranslation With {.Rule = "ModerateBounce", .Match = "(" + MakeUniRegEx(ArabicLetterDal) + "|" + MakeUniRegEx(ArabicLetterBeh) + "|" + MakeUniRegEx(ArabicLetterQaf) + "|" + MakeUniRegEx(ArabicLetterTah) + "|" + MakeUniRegEx(ArabicLetterJeem) + ")(?=(" + MakeRegMultiEx(Array.ConvertAll(ArabicTanweens, Function(Str As String) MakeUniRegEx(Str))) + "|" + MakeRegMultiEx(Array.ConvertAll(ArabicFathaDammaKasra, Function(Str As String) MakeUniRegEx(Str))) + ")\s*$)", _
            .Evaluator = New String() {"bounce"}}, _
        New RuleMetadataTranslation With {.Rule = "GreatBounce", .Match = "(" + MakeUniRegEx(ArabicLetterDal) + "|" + MakeUniRegEx(ArabicLetterBeh) + "|" + MakeUniRegEx(ArabicLetterQaf) + "|" + MakeUniRegEx(ArabicLetterTah) + "|" + MakeUniRegEx(ArabicLetterJeem) + ")(?=" + MakeUniRegEx(ArabicShadda) + MakeUniRegEx(ArabicSukun) + "\s*$)", _
            .Evaluator = New String() {"bounce"}}, _
        New RuleMetadataTranslation With {.Rule = "NasalizeCharacterDoubled", .Match = "(" + MakeUniRegEx(ArabicLetterNoon) + ")(" + MakeUniRegEx(ArabicShadda) + ")",
            .Evaluator = New String() {"nasalize|normalprolong", "empty"}}, _
        New RuleMetadataTranslation With {.Rule = "VowellessNoonClear", .Match = "(?:(" + MakeUniRegEx(ArabicLetterNoon) + MakeUniRegEx(ArabicSukun) + ")|(?:(" + MakeUniRegEx(ArabicLetterNoon) + ")|(" + MakeRegMultiEx(Array.ConvertAll(ArabicTanweens, Function(Str As String) MakeUniRegEx(Str))) + "))(?:\s+))(" + MakeUniRegEx(ArabicLetterAin) + "|" + MakeUniRegEx(ArabicLetterGhain) + "|" + MakeUniRegEx(ArabicLetterJeem) + "|" + MakeUniRegEx(ArabicLetterHah) + "|" + MakeUniRegEx(ArabicLetterHeh) + "|" + MakeUniRegEx(ArabicLetterHamza) + ")",
            .Evaluator = Nothing}, _
        New RuleMetadataTranslation With {.Rule = "VowellessNoonCovered", .Match = "(" + MakeUniRegEx(ArabicLetterNoon) + ")(" + MakeUniRegEx(ArabicSukun) + "?)(" + MakeUniRegEx(ArabicSmallHighMeemIsolatedForm) + ")|(?:(?:(" + MakeUniRegEx(ArabicFathatan) + "|" + MakeUniRegEx(ArabicDammatan) + "|" + MakeUniRegEx(ArabicKasratan) + ")(" + MakeUniRegEx(ArabicSmallHighMeemIsolatedForm) + ")|(" + MakeUniRegEx(ArabicKasratan) + ")(" + MakeUniRegEx(ArabicSmallLowMeem) + "))(?:\s+))(?=" + MakeUniRegEx(ArabicLetterBeh) + "(" + MakeUniRegEx(ArabicFatha) + "|" + MakeUniRegEx(ArabicDamma) + "|" + MakeUniRegEx(ArabicKasra) + "))",
            .Evaluator = New String() {"empty", Nothing, "nasalize", "dividetanween(,empty)", "nasalize", "dividetanween(,empty)", "nasalize"}}, _
        New RuleMetadataTranslation With {.Rule = "VowellessNoonAssimilatingNasalization", .Match = "(?:(" + MakeUniRegEx(ArabicLetterNoon) + MakeUniRegEx(ArabicSukun) + ")|(?:(" + MakeUniRegEx(ArabicLetterNoon) + ")|(" + MakeRegMultiEx(Array.ConvertAll(ArabicTanweens, Function(Str As String) MakeUniRegEx(Str))) + "))(?:\s+))(" + MakeUniRegEx(ArabicLetterNoon) + "|" + MakeUniRegEx(ArabicLetterWaw) + "|" + MakeUniRegEx(ArabicLetterMeem) + "|" + MakeUniRegEx(ArabicLetterYeh) + ")",
            .Evaluator = New String() {"assimilate", "assimilate", "dividetanween(,assimilate)", "nasalize|normalprolong|assimilator"}}, _
        New RuleMetadataTranslation With {.Rule = "VowellessNoonAssimilating", .Match = "(?:(" + MakeUniRegEx(ArabicLetterNoon) + MakeUniRegEx(ArabicSukun) + ")|(?:(" + MakeUniRegEx(ArabicLetterNoon) + ")|(" + MakeRegMultiEx(Array.ConvertAll(ArabicTanweens, Function(Str As String) MakeUniRegEx(Str))) + "))(?:\s+))(" + MakeUniRegEx(ArabicLetterLam) + "|" + MakeUniRegEx(ArabicLetterReh) + ")",
            .Evaluator = New String() {"assimilate", "assimilate", "dividetanween(,assimilate)", "assimilator"}}, _
        New RuleMetadataTranslation With {.Rule = "VowellessNoonHideHeaviness", .Match = "(?:(" + MakeUniRegEx(ArabicLetterNoon) + MakeUniRegEx(ArabicSukun) + ")|(?:(" + MakeUniRegEx(ArabicLetterNoon) + ")|(" + MakeRegMultiEx(Array.ConvertAll(ArabicTanweens, Function(Str As String) MakeUniRegEx(Str))) + "))(?:\s+))(?=" + MakeUniRegEx(ArabicLetterSad) + "|" + MakeUniRegEx(ArabicLetterDad) + "|" + MakeUniRegEx(ArabicLetterTah) + "|" + MakeUniRegEx(ArabicLetterZah) + "|" + MakeUniRegEx(ArabicLetterQaf) + ")",
            .Evaluator = New String() {"nasalize|normalprolong", "nasalize|normalprolong", "dividetanween(,normalprolong nasalize)"}}, _
        New RuleMetadataTranslation With {.Rule = "VowellessNoonHideLightness", .Match = "(?:(" + MakeUniRegEx(ArabicLetterNoon) + MakeUniRegEx(ArabicSukun) + ")|(?:(" + MakeUniRegEx(ArabicLetterNoon) + ")|(" + MakeRegMultiEx(Array.ConvertAll(ArabicTanweens, Function(Str As String) MakeUniRegEx(Str))) + "))(?:\s+))(?=" + MakeUniRegEx(ArabicLetterTeh) + "|" + MakeUniRegEx(ArabicLetterTheh) + "|" + MakeUniRegEx(ArabicLetterJeem) + "|" + MakeUniRegEx(ArabicLetterDal) + "|" + MakeUniRegEx(ArabicLetterThal) + "|" + MakeUniRegEx(ArabicLetterZain) + "|" + MakeUniRegEx(ArabicLetterSeen) + "|" + MakeUniRegEx(ArabicLetterSheen) + "|" + MakeUniRegEx(ArabicLetterFeh) + "|" + MakeUniRegEx(ArabicLetterKaf) + ")",
            .Evaluator = New String() {"nasalize|normalprolong", "nasalize|normalprolong", "dividetanween(,normalprolong nasalize)"}}, _
        New RuleMetadataTranslation With {.Rule = "CharacterDoubled", .Match = "(" + MakeUniRegEx(ArabicLetterMeem) + ")(" + MakeUniRegEx(ArabicShadda) + ")",
            .Evaluator = New String() {"nasalize|normalprolong", "empty"}}, _
        New RuleMetadataTranslation With {.Rule = "VowellessMeemHide", .Match = "(?:(" + MakeUniRegEx(ArabicLetterMeem) + MakeUniRegEx(ArabicSukun) + ")|(" + MakeUniRegEx(ArabicLetterMeem) + ")(?:\s+))(" + MakeUniRegEx(ArabicLetterBeh) + "(?=" + MakeUniRegEx(ArabicFatha) + "|" + MakeUniRegEx(ArabicDamma) + "|" + MakeUniRegEx(ArabicKasra) + "))",
            .Evaluator = New String() {"assimilate", "assimilate", "nasalize|normalprolong|assimilator"}}, _
        New RuleMetadataTranslation With {.Rule = "VowellessMeemAssimilating", .Match = "(?:(" + MakeUniRegEx(ArabicLetterMeem) + MakeUniRegEx(ArabicSukun) + ")|(" + MakeUniRegEx(ArabicLetterMeem) + ")(?:\s+))(" + MakeUniRegEx(ArabicLetterMeem) + "(?=" + MakeUniRegEx(ArabicFatha) + "|" + MakeUniRegEx(ArabicDamma) + "|" + MakeUniRegEx(ArabicKasra) + "))",
            .Evaluator = New String() {"assimilate", "assimilate", "nasalize|normalprolong|assimilator"}}, _
        New RuleMetadataTranslation With {.Rule = "VowellessMeemClear", .Match = "(?:(" + MakeUniRegEx(ArabicLetterMeem) + MakeUniRegEx(ArabicSukun) + ")|(" + MakeUniRegEx(ArabicLetterMeem) + ")(?:\s+))(?=(" + MakeUniRegEx(ArabicLetterTeh) + "|" + MakeUniRegEx(ArabicLetterTheh) + "|" + MakeUniRegEx(ArabicLetterJeem) + "|" + MakeUniRegEx(ArabicLetterHah) + "|" + MakeUniRegEx(ArabicLetterKhah) + "|" + MakeUniRegEx(ArabicLetterDal) + "|" + MakeUniRegEx(ArabicLetterThal) + "|" + MakeUniRegEx(ArabicLetterReh) + "|" + MakeUniRegEx(ArabicLetterZain) + "|" + MakeUniRegEx(ArabicLetterSeen) + "|" + MakeUniRegEx(ArabicLetterSheen) + "|" + MakeUniRegEx(ArabicLetterSad) + "|" + MakeUniRegEx(ArabicLetterDad) + "|" + MakeUniRegEx(ArabicLetterTah) + "|" + MakeUniRegEx(ArabicLetterZah) + "|" + MakeUniRegEx(ArabicLetterAin) + "|" + MakeUniRegEx(ArabicLetterGhain) + "|" + MakeUniRegEx(ArabicLetterQaf) + "|" + MakeUniRegEx(ArabicLetterKaf) + "|" + MakeUniRegEx(ArabicLetterLam) + "|" + MakeUniRegEx(ArabicLetterNoon) + "|" + MakeUniRegEx(ArabicLetterHamza) + "|((" + MakeUniRegEx(ArabicSukun) + "|" + MakeUniRegEx(ArabicFatha) + ")(" + MakeUniRegEx(ArabicLetterYeh) + ")))(" + MakeRegMultiEx(Array.ConvertAll(ArabicFathaDammaKasra, Function(Str As String) MakeUniRegEx(Str))) + "))",
            .Evaluator = Nothing}, _
        New RuleMetadataTranslation With {.Rule = "VowellessMeemClearGreater", .Match = "(?:(" + MakeUniRegEx(ArabicLetterMeem) + MakeUniRegEx(ArabicSukun) + ")|(" + MakeUniRegEx(ArabicLetterMeem) + ")(?:\s+))(?=(" + MakeUniRegEx(ArabicLetterFeh) + "|((" + MakeUniRegEx(ArabicSukun) + "|" + MakeUniRegEx(ArabicFatha) + ")(" + MakeUniRegEx(ArabicLetterWaw) + ")))(" + MakeRegMultiEx(Array.ConvertAll(ArabicFathaDammaKasra, Function(Str As String) MakeUniRegEx(Str))) + "))",
            .Evaluator = Nothing}, _
        New RuleMetadataTranslation With {.Rule = "AlefSmallHighRoundedZero", .Match = "(" + MakeUniRegEx(ArabicLetterAlef) + MakeUniRegEx(ArabicSmallHighRoundedZero) + ")(?=\s*$|\s+)",
            .Evaluator = New String() {"empty"}}, _
        New RuleMetadataTranslation With {.Rule = "AlefSmallHighRoundedZero", .Match = "(" + MakeUniRegEx(ArabicLetterWaw) + MakeUniRegEx(ArabicSmallHighRoundedZero) + ")(?=\s*$|\s+)",
            .Evaluator = New String() {"empty"}}, _
        New RuleMetadataTranslation With {.Rule = "SmallHighUprightRectangularZero", .Match = "(" + MakeUniRegEx(ArabicSmallHighUprightRectangularZero) + ")",
            .Evaluator = New String() {"empty"}}, _
        New RuleMetadataTranslation With {.Rule = "EmptyCentreLowStop", .Match = "(" + MakeUniRegEx(ArabicEmptyCentreLowStop) + ")(" + MakeUniRegEx(ArabicLetterAlefMaksura) + ")(" + MakeUniRegEx(ArabicLetterSuperscriptAlef) + ")",
            .Evaluator = New String() {"empty", "helperfatha", "helperalef"}}, _
        New RuleMetadataTranslation With {.Rule = "EmptyCentreHighStop", .Match = "(" + MakeUniRegEx(ArabicEmptyCentreHighStop) + ")",
            .Evaluator = New String() {"empty"}}, _
        New RuleMetadataTranslation With {.Rule = "RoundedHighStopWithFilledCentre", .Match = "(" + MakeUniRegEx(ArabicLetterAlef) + ")()(" + MakeUniRegEx(ArabicRoundedHighStopWithFilledCentre) + ")",
            .Evaluator = New String() {Nothing, "helperhamza", "helperfatha"}}, _
        New RuleMetadataTranslation With {.Rule = "EmptyHamza", .Match = "((?:^|" + CertainStopPattern + "|" + OptionalStopPattern + ")\s*)(" + MakeUniRegEx(ArabicLetterAlefWasla) + ")(?=" + MakeUniRegEx(ArabicLetterLam) + "((" + MakeRegMultiEx(Array.ConvertAll(ArabicSunLetters, Function(Str As String) MakeUniRegEx(Str))) + ")?" + MakeUniRegEx(ArabicShadda) + "|" + MakeUniRegEx(ArabicSukun) + "?(" + MakeRegMultiEx(Array.ConvertAll(ArabicMoonLetters, Function(Str As String) MakeUniRegEx(Str))) + "|" + MakeUniRegEx(ArabicLetterAlefWithHamzaAbove) + "|" + MakeUniRegEx(ArabicLetterAlefWithHamzaBelow) + "|" + MakeUniRegEx(ArabicLetterHamza) + "|" + MakeUniRegEx(ArabicTatweel) + MakeUniRegEx(ArabicHamzaAbove) + ")))",
            .Evaluator = New String() {Nothing, "helperfatha"}}, _
        New RuleMetadataTranslation With {.Rule = "EmptyHamza", .Match = "((?:^|" + CertainStopPattern + "|" + OptionalStopPattern + ")\s*)(" + MakeUniRegEx(ArabicLetterAlefWasla) + ")(?=(" + MakeRegMultiEx(Array.ConvertAll(ArabicMoonLetters, Function(Str As String) MakeUniRegEx(Str))) + "|" + MakeRegMultiEx(Array.ConvertAll(ArabicSunLettersNoLam, Function(Str As String) MakeUniRegEx(Str))) + ")(" + MakeUniRegEx(ArabicSukun) + "|" + MakeRegMultiEx(Array.ConvertAll(ArabicFathaDammaKasra, Function(Str As String) MakeUniRegEx(Str))) + ")?(" + MakeRegMultiEx(Array.ConvertAll(ArabicLetters, Function(Str As String) MakeUniRegEx(Str))) + ")(" + MakeUniRegEx(ArabicFatha) + "|" + MakeUniRegEx(ArabicKasra) + ")|" + MakeRegMultiEx(Array.ConvertAll(ArabicWaslKasraExceptions, Function(Str As String) MakeUniRegEx(TransliterateFromBuckwalter(Str)))) + ")",
            .Evaluator = New String() {Nothing, "helperkasra"}}, _
        New RuleMetadataTranslation With {.Rule = "EmptyHamza", .Match = "((?:^|" + CertainStopPattern + "|" + OptionalStopPattern + ")\s*)(" + MakeUniRegEx(ArabicLetterAlefWasla) + ")(?=(" + MakeRegMultiEx(Array.ConvertAll(ArabicLetters, Function(Str As String) MakeUniRegEx(Str))) + ")(" + MakeUniRegEx(ArabicSukun) + "|" + MakeRegMultiEx(Array.ConvertAll(ArabicFathaDammaKasra, Function(Str As String) MakeUniRegEx(Str))) + ")?(" + MakeRegMultiEx(Array.ConvertAll(ArabicLetters, Function(Str As String) MakeUniRegEx(Str))) + ")" + MakeUniRegEx(ArabicDamma) + ")",
            .Evaluator = New String() {Nothing, "helperdamma"}}, _
        New RuleMetadataTranslation With {.Rule = "EmptyHamza", .Match = "(\S+|" + CertainNotStopPattern + "\s*|" + OptionalNotStopPattern + "\s*)(" + MakeUniRegEx(ArabicLetterAlefWasla) + ")",
            .Evaluator = New String() {Nothing, "empty"}}, _
        New RuleMetadataTranslation With {.Rule = "AssimilateLaamSunLetter", .Match = "((?:^\s*|\s+)(?:" + MakeUniRegEx(ArabicLetterWaw) + MakeUniRegEx(ArabicFatha) + "|" + MakeUniRegEx(ArabicLetterBeh) + MakeUniRegEx(ArabicKasra) + "|" + MakeUniRegEx(ArabicLetterTeh) + MakeUniRegEx(ArabicFatha) + "|" + MakeUniRegEx(ArabicLetterKaf) + MakeUniRegEx(ArabicFatha) + "|" + MakeUniRegEx(ArabicLetterLam) + MakeUniRegEx(ArabicKasra) + ")?" + MakeUniRegEx(ArabicLetterAlefWasla) + ")(" + MakeUniRegEx(ArabicLetterLam) + ")(" + MakeRegMultiEx(Array.ConvertAll(ArabicSunLetters, Function(Str As String) MakeUniRegEx(Str))) + ")(?=" + MakeUniRegEx(ArabicShadda) + ")",
            .Evaluator = New String() {Nothing, "assimilate", "assimilator"}}, _
        New RuleMetadataTranslation With {.Rule = "ClearLaamMoonLetter", .Match = "((?:^\s*|\s+)(?:" + MakeUniRegEx(ArabicLetterWaw) + MakeUniRegEx(ArabicFatha) + "|" + MakeUniRegEx(ArabicLetterBeh) + MakeUniRegEx(ArabicKasra) + "|" + MakeUniRegEx(ArabicLetterTeh) + MakeUniRegEx(ArabicFatha) + "|" + MakeUniRegEx(ArabicLetterKaf) + MakeUniRegEx(ArabicFatha) + "|" + MakeUniRegEx(ArabicLetterLam) + MakeUniRegEx(ArabicKasra) + ")?" + MakeUniRegEx(ArabicLetterAlefWasla) + ")(" + MakeUniRegEx(ArabicLetterLam) + ")(?=" + MakeUniRegEx(ArabicSukun) + "?(?:" + MakeRegMultiEx(Array.ConvertAll(ArabicMoonLetters, Function(Str As String) MakeUniRegEx(Str))) + "))",
            .Evaluator = New String() {Nothing, "helperprefix", Nothing}}, _
        New RuleMetadataTranslation With {.Rule = "Sukun", .Match = "(" + MakeUniRegEx(ArabicSukun) + ")", _
            .Evaluator = New String() {"empty"}}, _
        New RuleMetadataTranslation With {.Rule = "AlefMaksuraDaggerAlef", .Match = "(" + MakeUniRegEx(ArabicLetterAlefMaksura) + ")(?=" + MakeUniRegEx(ArabicLetterSuperscriptAlef) + ")", _
            .Evaluator = New String() {"helperfatha"}}, _
        New RuleMetadataTranslation With {.Rule = "AlefMaddah", .Match = "(" + MakeUniRegEx(ArabicLetterAlefWithMaddaAbove) + ")", _
            .Evaluator = New String() {"dividelettersymbol(helperalef,helpermadda)"}}, _
        New RuleMetadataTranslation With {.Rule = "Maddah", .Match = "(" + MakeUniRegEx(ArabicLetterAlef) + ")(" + MakeUniRegEx(ArabicMaddahAbove) + ")", _
            .Evaluator = New String() {Nothing, "helpermadda"}}, _
        New RuleMetadataTranslation With {.Rule = "Maddah", .Match = "(" + MakeUniRegEx(ArabicLetterYeh) + "|" + MakeUniRegEx(ArabicSmallYeh) + "|" + MakeUniRegEx(ArabicKasra) + MakeUniRegEx(ArabicLetterAlefMaksura) + ")(" + MakeUniRegEx(ArabicMaddahAbove) + ")", _
            .Evaluator = New String() {Nothing, "permissibleprolong"}}, _
        New RuleMetadataTranslation With {.Rule = "AlefWaslaWawHamza", .Match = "((?:^|" + CertainStopPattern + "|" + OptionalStopPattern + ")\s*" + MakeUniRegEx(ArabicLetterAlefWasla) + ")(" + MakeUniRegEx(ArabicLetterWawWithHamzaAbove) + ")", _
            .Evaluator = New String() {Nothing, "dividelettersymbol(helperwaw,empty)"}}, _
        New RuleMetadataTranslation With {.Rule = "AlefWaslaAlefMaksuraHamza", .Match = "((?:^|" + CertainStopPattern + "|" + OptionalStopPattern + ")\s*" + MakeUniRegEx(ArabicLetterAlefWasla) + ")(" + MakeUniRegEx(ArabicLetterYehWithHamzaAbove) + ")", _
            .Evaluator = New String() {Nothing, "dividelettersymbol(helperyeh,empty)"}}, _
        New RuleMetadataTranslation With {.Rule = "AlefHamza", .Match = "((?:^|" + CertainStopPattern + "|" + OptionalStopPattern + ")\s*)(" + MakeUniRegEx(ArabicLetterAlefWithHamzaAbove) + "|" + MakeUniRegEx(ArabicLetterAlefWithHamzaBelow) + "|" + MakeUniRegEx(ArabicLetterHamza) + ")", _
            .Evaluator = New String() {Nothing, "empty"}}, _
        New RuleMetadataTranslation With {.Rule = "AlefHamza", .Match = "(\S+|" + CertainNotStopPattern + "\s*|" + OptionalNotStopPattern + "\s*)(" + MakeUniRegEx(ArabicLetterAlefWithHamzaAbove) + "|" + MakeUniRegEx(ArabicLetterAlefWithHamzaBelow) + ")", _
            .Evaluator = New String() {Nothing, "dividelettersymbol(empty,helperhamza)"}}, _
        New RuleMetadataTranslation With {.Rule = "AlefHamza", .Match = "(\S+|" + CertainNotStopPattern + "\s*|" + OptionalNotStopPattern + "\s*)(" + MakeUniRegEx(ArabicLetterAlefWithHamzaAbove) + "|" + MakeUniRegEx(ArabicLetterAlefWithHamzaBelow) + ")", _
            .Evaluator = New String() {Nothing, "dividelettersymbol(empty,helperhamza)"}}, _
        New RuleMetadataTranslation With {.Rule = "WawHamzah", .Match = "(" + MakeUniRegEx(ArabicLetterWawWithHamzaAbove) + ")", _
            .Evaluator = New String() {"dividelettersymbol(empty,helperhamza)"}}, _
        New RuleMetadataTranslation With {.Rule = "AlefMaksuraHamzah", .Match = "(" + MakeUniRegEx(ArabicKasra) + "|" + MakeUniRegEx(ArabicFatha) + "(?:" + MakeUniRegEx(ArabicLetterAlef) + MakeUniRegEx(ArabicMaddahAbove) + "?)?|" + MakeUniRegEx(ArabicLetterSuperscriptAlef) + MakeUniRegEx(ArabicMaddahAbove) + ")(" + MakeUniRegEx(ArabicLetterYehWithHamzaAbove) + ")", _
            .Evaluator = New String() {Nothing, "dividelettersymbol(empty,helperhamza)"}}, _
        New RuleMetadataTranslation With {.Rule = "SmallWaw", .Match = "(" + MakeUniRegEx(ArabicSmallWaw) + ")", _
            .Evaluator = New String() {"helperwaw"}}, _
        New RuleMetadataTranslation With {.Rule = "SmallYeh", .Match = "(" + MakeUniRegEx(ArabicSmallYeh) + ")", _
            .Evaluator = New String() {"helperyeh"}}, _
        New RuleMetadataTranslation With {.Rule = "SmallMeem", .Match = "(" + MakeUniRegEx(ArabicSmallLowMeem) + "|" + MakeUniRegEx(ArabicSmallHighMeemIsolatedForm) + ")", _
            .Evaluator = New String() {"helpermeem"}}, _
        New RuleMetadataTranslation With {.Rule = "SmallNoon", .Match = "(" + MakeUniRegEx(ArabicSmallHighNoon) + ")", _
            .Evaluator = New String() {"helpernoon"}}, _
        New RuleMetadataTranslation With {.Rule = "SmallHighSeen", .Match = "(" + MakeUniRegEx(ArabicSmallHighSeen) + ")\S+", _
            .Evaluator = New String() {"helperseen"}}, _
        New RuleMetadataTranslation With {.Rule = "SmallLowSeen", .Match = "(" + MakeUniRegEx(ArabicSmallLowSeen) + ")", _
            .Evaluator = New String() {"empty"}}, _
        New RuleMetadataTranslation With {.Rule = "DaggerAlef", .Match = "(" + MakeUniRegEx(ArabicLetterSuperscriptAlef) + ")", _
            .Evaluator = New String() {"helperalef"}}, _
        New RuleMetadataTranslation With {.Rule = "SmallHamza", .Match = "(" + MakeUniRegEx(ArabicTatweel) + MakeUniRegEx(ArabicHamzaAbove) + "|" + MakeUniRegEx(ArabicHamzaBelow) + ")", _
            .Evaluator = New String() {"helperhamza"}}, _
        New RuleMetadataTranslation With {.Rule = "FathaAlefMaksura", .Match = "(" + MakeUniRegEx(ArabicFatha) + ")(" + MakeUniRegEx(ArabicLetterAlefMaksura) + ")", _
            .Evaluator = New String() {Nothing, "helperalef"}}, _
        New RuleMetadataTranslation With {.Rule = "KasraAlefMaksura", .Match = "(" + MakeUniRegEx(ArabicKasra) + ")(" + MakeUniRegEx(ArabicLetterAlefMaksura) + ")", _
            .Evaluator = New String() {Nothing, "helperyeh"}}
    }
        'U+2BC3 is horizontal stop sign make red color, U+2B45/6 is left/rightwards quadruple arrow make green color
    Public Shared Function MakeUniRegEx(Input As String) As String
        Return String.Join(String.Empty, Array.ConvertAll(Of Char, String)(Input.ToCharArray(), Function(Ch As Char) "\u" + AscW(Ch).ToString("X4")))
    End Function
    Public Shared Function MakeRegMultiEx(Input As String()) As String
        Return String.Join("|", Input)
    End Function
    Public Shared Function ArabicLetterSpelling(Input As String) As String
        Dim Output As String = String.Empty
        For Each Ch As Char In Input
            Dim Index As Integer = FindLetterBySymbol(Ch)
            If Index <> -1 AndAlso IsLetter(Index) Then
                If Output <> String.Empty Then Output += " "
                Output += CachedData.IslamData.ArabicLetters(Index).SymbolName
            ElseIf Index <> -1 AndAlso CachedData.IslamData.ArabicLetters(Index).Symbol = ArabicMaddahAbove Then
                Output += Ch
            End If
        Next
        Return TransliterateFromBuckwalter(Output)
    End Function
    Class RuleMetadataComparer
        Implements Collections.Generic.IComparer(Of RuleMetadata)
        Public Function Compare(x As RuleMetadata, y As RuleMetadata) As Integer Implements Generic.IComparer(Of RuleMetadata).Compare
            If x.Index = y.Index Then
                Return y.Length.CompareTo(x.Length)
            Else
                Return y.Index.CompareTo(x.Index)
            End If
        End Function
    End Class
    Public Shared Function ApplyColorRules(ByVal ArabicString As String) As RenderArray.RenderText()
        Dim Count As Integer
        Dim Index As Integer
        Dim MetadataList As New Generic.List(Of RuleMetadata)
        Dim Strings As New Generic.List(Of RenderArray.RenderText)
        For Count = 0 To RulesOfRecitationRegEx.Length - 1
            If Not RulesOfRecitationRegEx(Count).Evaluator Is Nothing Then
                Dim Matches As System.Text.RegularExpressions.MatchCollection = System.Text.RegularExpressions.Regex.Matches(ArabicString, RulesOfRecitationRegEx(Count).Match)
                For MatchIndex As Integer = 0 To Matches.Count - 1
                    For SubCount As Integer = 0 To RulesOfRecitationRegEx(Count).Evaluator.Length - 1
                        If Not RulesOfRecitationRegEx(Count).Evaluator(SubCount) Is Nothing Then
                            MetadataList.Add(New RuleMetadata(Matches(MatchIndex).Groups(SubCount + 1).Index, Matches(MatchIndex).Groups(SubCount + 1).Length, RulesOfRecitationRegEx(Count).Evaluator(SubCount)))
                        End If
                    Next
                Next
            End If
        Next
        MetadataList.Sort(New RuleMetadataComparer)
        For Index = 0 To MetadataList.Count - 1
            If If(Index <> 0, MetadataList(Index - 1).Index + MetadataList(Index - 1).Length, 0) <> MetadataList(Index).Index Then
                Strings.Add(New RenderArray.RenderText(RenderArray.RenderDisplayClass.eArabic, ArabicString.Substring(If(Index <> 0, MetadataList(Index - 1).Index + MetadataList(Index - 1).Length, 0), MetadataList(Index).Index - If(Index <> 0, MetadataList(Index - 1).Index + MetadataList(Index - 1).Length, 0))))
            End If
            Strings.Add(New RenderArray.RenderText(RenderArray.RenderDisplayClass.eArabic, ArabicString.Substring(MetadataList(Index).Index, MetadataList(Index).Length)))
            For Count = 0 To ColoringRules.Length - 1
                Dim Match As Integer = Array.FindIndex(ColoringRules(Count).Match.Split("|"c), Function(Str As String) Array.IndexOf(MetadataList(Index).Type.Split("|"c), Str) <> -1)
                If Match <> -1 Then
                    'ApplyColorRules(Strings(Strings.Count - 1).Text)
                    Dim Text As RenderArray.RenderText = Strings(Strings.Count - 1)
                    Text.Clr = ColoringRules(Count).Color
                    Strings(Strings.Count - 1) = Text
                End If
            Next
        Next
        If MetadataList.Count = 0 OrElse MetadataList(MetadataList.Count - 1).Index <> ArabicString.Length - 1 Then
            Strings.Add(New RenderArray.RenderText(RenderArray.RenderDisplayClass.eArabic, ArabicString.Substring(If(MetadataList.Count = 0, 0, MetadataList(MetadataList.Count - 1).Index))))
        End If
        Return Strings.ToArray()
    End Function
    Public Shared ArabicOrdinalNumbers As String() = {">aw~alN", "vaAniyN", "vaAlivN", "raAbiEN", "xaAmisN", "saAdisN", "saAbiEN", "vaAminN", "taAsiEN", "EaA$irN"}
    Public Shared ArabicOrdinalExtraNumbers As String() = {"HaAdiy"} '">uwalaY" "ap"
    Public Shared ArabicFractionNumbers As String() = {"nisof", "vuluv", "rubuE", "xumus"}
    Public Shared ArabicBaseNumbers As String() = {"waAHidN", "<ivonaAni", "valaAvapN", ">arobaEapN", "xamosapN", "sit~apN", "saboEapN", "vamaAnoyapN", "tisoEapN"}
    Public Shared ArabicBaseExtraNumbers As String() = {">aHada", "<ivonaA"}
    Public Shared ArabicBaseTenNumbers As String() = {"SiforN", "Ea$arapN", "Ei$oruwna", "valaAvuwna", ">arobaEuwna", "xamosuwna", "sit~uwna", "saboEuwna", "vamaAnuwna", "tisoEuwna"}
    Public Shared ArabicBaseHundredNumbers As String() = {"mi}apN", "mi}ataAni", "valaAvumi}apK", ">arobaEumi}apK", "xamosumi}apK", "sit~umi}apK", "saboEumi}apK", "vamaAnumi}apK", "tisoEumi}apK"} '"miA}apN"
    Public Shared ArabicBaseThousandNumbers As String() = {">alofN", ">alofaAni", "|laAfK"}
    Public Shared ArabicBaseMillionNumbers As String() = {"miloyuwnu", "miloyuwnaAni", "malaAyiyna"}
    Public Shared ArabicBaseBillionNumbers As String() = {"biloyuwnu", "biloyuwnaAni", "balaAyiyna"}
    Public Shared ArabicBaseMilliardNumbers As String() = {"miloyaAru", "miloyaAraAni", "miloyaAraAtK"}
    Public Shared ArabicBaseTrillionNumbers As String() = {"toriloyuwnu", "toriloyuwnaAni", "triloyuwnaAtK"}
    Public Shared Function ArabicWordForLessThanThousand(ByVal Number As Integer, UseClassic As Boolean, UseAlefHundred As Boolean) As String
        Dim Str As String = String.Empty
        Dim HundStr As String = String.Empty
        If Number >= 100 Then
            HundStr = If(UseAlefHundred, ArabicBaseHundredNumbers((Number \ 100) - 1).Insert(2, "A"), ArabicBaseHundredNumbers((Number \ 100) - 1))
            If (Number Mod 100) = 0 Then Return HundStr
            Number = Number Mod 100
        End If
        If (Number Mod 10) <> 0 And Number <> 11 And Number <> 12 Then
            Str = ArabicBaseNumbers((Number Mod 10) - 1)
        End If
        If Number >= 11 AndAlso Number < 20 Then
            If Number = 11 Or Number = 12 Then
                Str += ArabicBaseExtraNumbers(Number - 11)
            Else
                Str = Str.Remove(Str.Length - 1) + "a"
            End If
            Str += " Ea$ara"
        ElseIf (Number = 0 And Str = String.Empty) Or Number = 10 Or Number >= 20 Then
            Str = If(Str = String.Empty, String.Empty, Str + " wa") + ArabicBaseTenNumbers(Number \ 10)
        End If
        Return If(UseClassic, If(Str = String.Empty, String.Empty, Str + If(HundStr = String.Empty, String.Empty, " wa")) + HundStr, If(HundStr = String.Empty, String.Empty, HundStr + If(Str = String.Empty, String.Empty, " wa")) + Str)
    End Function
    Public Shared Function ArabicWordFromNumber(ByVal Number As Long, UseClassic As Boolean, UseAlefHundred As Boolean, UseMilliard As Boolean) As String
        Dim Str As String = String.Empty
        Dim NextStr As String = String.Empty
        Dim CurBase As Integer = 3
        Dim BaseNums As Long() = {1000, 1000000, 1000000000, 1000000000000}
        Dim Bases As String()() = {ArabicBaseThousandNumbers, ArabicBaseMillionNumbers, If(UseMilliard, ArabicBaseMilliardNumbers, ArabicBaseBillionNumbers), ArabicBaseTrillionNumbers}
        Do
            If Number >= BaseNums(CurBase) And Number < 2 * BaseNums(CurBase) Then
                NextStr = Bases(CurBase)(0)
            ElseIf Number >= 2 * BaseNums(CurBase) And Number < 3 * BaseNums(CurBase) Then
                NextStr = Bases(CurBase)(1)
            ElseIf Number >= 3 * BaseNums(CurBase) And Number < 10 * BaseNums(CurBase) Then
                NextStr = ArabicBaseNumbers(Number \ BaseNums(CurBase) - 1).Remove(ArabicBaseNumbers(Number \ BaseNums(CurBase) - 1).Length - 1) + "u " + Bases(CurBase)(2)
            ElseIf Number >= 10 * BaseNums(CurBase) And Number < 11 * BaseNums(CurBase) Then
                NextStr = ArabicBaseTenNumbers(1).Remove(ArabicBaseTenNumbers(1).Length - 1) + "u " + Bases(CurBase)(2)
            ElseIf Number >= BaseNums(CurBase) Then
                NextStr = ArabicWordForLessThanThousand((Number \ BaseNums(CurBase)) Mod 100, UseClassic, UseAlefHundred)
                If Number >= 100 * BaseNums(CurBase) And Number < If(UseClassic, 200, 101) * BaseNums(CurBase) Then
                    NextStr = NextStr.Remove(NextStr.Length - 1) + "u " + Bases(CurBase)(0).Remove(Bases(CurBase)(0).Length - 1) + "K"
                ElseIf Number >= 200 * BaseNums(CurBase) And Number < If(UseClassic, 300, 201) * BaseNums(CurBase) Then
                    NextStr = NextStr.Remove(NextStr.Length - 2) + " " + Bases(CurBase)(0).Remove(Bases(CurBase)(0).Length - 1) + "K"
                ElseIf Number >= 300 * BaseNums(CurBase) And (UseClassic Or (Number \ BaseNums(CurBase)) Mod 100 = 0) Then
                    NextStr = NextStr.Remove(NextStr.Length - 1) + "i " + Bases(CurBase)(0).Remove(Bases(CurBase)(0).Length - 1) + "K"
                Else
                    NextStr += " " + Bases(CurBase)(0).Remove(Bases(CurBase)(0).Length - 1) + "FA"
                End If
            End If
            Number = Number Mod BaseNums(CurBase)
            CurBase -= 1
            Str = If(UseClassic, If(NextStr = String.Empty, String.Empty, NextStr + If(Str = String.Empty, String.Empty, " wa")) + Str, If(Str = String.Empty, String.Empty, Str + If(NextStr = String.Empty, String.Empty, " wa")) + NextStr)
            NextStr = String.Empty
        Loop While CurBase >= 0
        If Number <> 0 Or Str = String.Empty Then NextStr = ArabicWordForLessThanThousand(Number, UseClassic, UseAlefHundred)
        Return If(UseClassic, If(NextStr = String.Empty, String.Empty, NextStr + If(Str = String.Empty, String.Empty, " wa")) + Str, If(Str = String.Empty, String.Empty, Str + If(NextStr = String.Empty, String.Empty, " wa")) + NextStr)
    End Function
    Public Shared Function NegativeMatchEliminator(NegativeMatch As String, Evaluator As String) As System.Text.RegularExpressions.MatchEvaluator
        Return Function(Match As System.Text.RegularExpressions.Match)
                   Return If(NegativeMatch <> String.Empty AndAlso Match.Result(NegativeMatch) <> String.Empty, Match.Value, Match.Result(Evaluator))
               End Function
    End Function
    Public Shared Function ChangeScript(ArabicString As String, ScriptType As TanzilReader.QuranScripts) As String
        If ScriptType = TanzilReader.QuranScripts.UthmaniMin Then
            For Count = 0 To UthmaniMinimalScript.Length - 1
                ArabicString = System.Text.RegularExpressions.Regex.Replace(ArabicString, UthmaniMinimalScript(Count).Match, NegativeMatchEliminator(UthmaniMinimalScript(Count).NegativeMatch, UthmaniMinimalScript(Count).Evaluator))
            Next
        ElseIf ScriptType = TanzilReader.QuranScripts.SimpleEnhanced Then
            For Count = 0 To SimpleEnhancedScript.Length - 1
                ArabicString = System.Text.RegularExpressions.Regex.Replace(ArabicString, SimpleEnhancedScript(Count).Match, NegativeMatchEliminator(SimpleEnhancedScript(Count).NegativeMatch, SimpleEnhancedScript(Count).Evaluator))
            Next
        ElseIf ScriptType = TanzilReader.QuranScripts.Simple Then
            For Count = 0 To SimpleEnhancedScript.Length - 1
                ArabicString = System.Text.RegularExpressions.Regex.Replace(ArabicString, SimpleEnhancedScript(Count).Match, NegativeMatchEliminator(SimpleEnhancedScript(Count).NegativeMatch, SimpleEnhancedScript(Count).Evaluator))
            Next
            For Count = 0 To SimpleScript.Length - 1
                ArabicString = System.Text.RegularExpressions.Regex.Replace(ArabicString, SimpleScript(Count).Match, NegativeMatchEliminator(SimpleScript(Count).NegativeMatch, SimpleScript(Count).Evaluator))
            Next
        ElseIf ScriptType = TanzilReader.QuranScripts.SimpleClean Then
            For Count = 0 To SimpleEnhancedScript.Length - 1
                ArabicString = System.Text.RegularExpressions.Regex.Replace(ArabicString, SimpleEnhancedScript(Count).Match, NegativeMatchEliminator(SimpleEnhancedScript(Count).NegativeMatch, SimpleEnhancedScript(Count).Evaluator))
            Next
            For Count = 0 To SimpleCleanScript.Length - 1
                ArabicString = System.Text.RegularExpressions.Regex.Replace(ArabicString, SimpleCleanScript(Count).Match, NegativeMatchEliminator(SimpleCleanScript(Count).NegativeMatch, SimpleCleanScript(Count).Evaluator))
            Next
        ElseIf ScriptType = TanzilReader.QuranScripts.SimpleMin Then
            For Count = 0 To SimpleEnhancedScript.Length - 1
                ArabicString = System.Text.RegularExpressions.Regex.Replace(ArabicString, SimpleEnhancedScript(Count).Match, NegativeMatchEliminator(SimpleEnhancedScript(Count).NegativeMatch, SimpleEnhancedScript(Count).Evaluator))
            Next
            For Count = 0 To SimpleMinimalScript.Length - 1
                ArabicString = System.Text.RegularExpressions.Regex.Replace(ArabicString, SimpleMinimalScript(Count).Match, NegativeMatchEliminator(SimpleMinimalScript(Count).NegativeMatch, SimpleMinimalScript(Count).Evaluator))
            Next
        End If
        Return ArabicString
    End Function
    Public Shared Function ReplaceMetadata(ArabicString As String, MetadataRule As RuleMetadata) As String
        For Count As Integer = 0 To ColoringSpelledOutRules.Length - 1
            Dim Match As String = Array.Find(ColoringSpelledOutRules(Count).Match.Split("|"c), Function(Str As String) Array.IndexOf(Array.ConvertAll(MetadataRule.Type.Split("|"c), Function(S As String) System.Text.RegularExpressions.Regex.Replace(S, "\(.*\)", String.Empty)), Str) <> -1)
            If Match <> Nothing Then
                Dim Str As String = String.Format(ColoringSpelledOutRules(Count).Evaluator, ArabicString.Substring(MetadataRule.Index, MetadataRule.Length))
                If ColoringSpelledOutRules(Count).RuleFunc <> RuleFuncs.eNone Then
                    Dim Args As String() = RuleFunctions(ColoringSpelledOutRules(Count).RuleFunc - 1)(Str)
                    If Args.Length = 1 Then
                        Str = Args(0)
                    Else
                        Dim MetaArgs As String() = System.Text.RegularExpressions.Regex.Match(MetadataRule.Type, Match + "\((.*)\)").Groups(1).Value.Split(",")
                        Str = String.Empty
                        For Index As Integer = 0 To Args.Length - 1
                            Str += ReplaceMetadata(Args(Index), New RuleMetadata(0, Args(Index).Length, MetaArgs(Index).Replace(" "c, "|"c)))
                        Next
                    End If
                End If
                ArabicString = ArabicString.Insert(MetadataRule.Index + MetadataRule.Length, Str).Remove(MetadataRule.Index, MetadataRule.Length)
            End If
        Next
        Return ArabicString
    End Function
    Public Shared Sub DoErrorCheck(ByVal ArabicString As String)
        'need to check for decomposed first
        Dim Count As Integer
        For Count = 0 To ErrorCheckRules.Length - 1
            Dim Matches As System.Text.RegularExpressions.MatchCollection = System.Text.RegularExpressions.Regex.Matches(ArabicString, ErrorCheckRules(Count).Match)
            For MatchIndex As Integer = 0 To Matches.Count - 1
                If ErrorCheckRules(Count).NegativeMatch Is Nothing OrElse Matches(MatchIndex).Result(ErrorCheckRules(Count).NegativeMatch) = String.Empty Then
                    Debug.Print(ErrorCheckRules(Count).Rule + ": " + TransliterateToScheme(ArabicString, TranslitScheme.Literal, String.Empty).Insert(Matches(MatchIndex).Index, "<!-- -->"))
                End If
            Next
        Next
    End Sub
    Public Shared Function TransliterateWithRules(ByVal ArabicString As String, Scheme As String) As String
        Dim Count As Integer
        ArabicString.Replace(Arabic.RightToLeftMark, String.Empty)
        Dim MetadataList As New Generic.List(Of RuleMetadata)
        DoErrorCheck(ArabicString)
        For Count = 0 To RulesOfRecitationRegEx.Length - 1
            If Not RulesOfRecitationRegEx(Count).Evaluator Is Nothing Then
                Dim Matches As System.Text.RegularExpressions.MatchCollection = System.Text.RegularExpressions.Regex.Matches(ArabicString, RulesOfRecitationRegEx(Count).Match)
                For MatchIndex As Integer = 0 To Matches.Count - 1
                    For SubCount As Integer = 0 To RulesOfRecitationRegEx(Count).Evaluator.Length - 1
                        If Not RulesOfRecitationRegEx(Count).Evaluator(SubCount) Is Nothing And (Matches(MatchIndex).Groups(SubCount + 1).Length <> 0 Or Array.IndexOf(AllowZeroLength, RulesOfRecitationRegEx(Count).Evaluator(SubCount)) <> -1) Then
                            MetadataList.Add(New RuleMetadata(Matches(MatchIndex).Groups(SubCount + 1).Index, Matches(MatchIndex).Groups(SubCount + 1).Length, RulesOfRecitationRegEx(Count).Evaluator(SubCount)))
                        End If
                    Next
                Next
            End If
        Next
        MetadataList.Sort(New RuleMetadataComparer)
        Dim Index As Integer
        For Index = 0 To MetadataList.Count - 1
            ArabicString = ReplaceMetadata(ArabicString, MetadataList(Index))
        Next
        'redundant romanization rules should have -'s such as seen/teh/kaf-heh
        For Count = 0 To RomanizationRules.Length - 1
            If RomanizationRules(Count).RuleFunc = RuleFuncs.eNone Then
                ArabicString = System.Text.RegularExpressions.Regex.Replace(ArabicString, RomanizationRules(Count).Match, RomanizationRules(Count).Evaluator)
            Else
                ArabicString = System.Text.RegularExpressions.Regex.Replace(ArabicString, RomanizationRules(Count).Match, Function(Match As System.Text.RegularExpressions.Match) RuleFunctions(RomanizationRules(Count).RuleFunc - 1)(Match.Result(RomanizationRules(Count).Evaluator))(0))
            End If
        Next

        'process wasl loanwords and names
        'process loanwords and names
        Return ArabicString
    End Function
    Shared Function GetArabicSymbolJSArray() As String
        Dim Letters(CachedData.IslamData.ArabicLetters.Length - 1) As IslamData.ArabicSymbol
        CachedData.IslamData.ArabicLetters.CopyTo(Letters, 0)
        Array.Sort(Letters, New StringLengthComparer("RomanTranslit"))
        GetArabicSymbolJSArray = "var arabicLetters = " + _
                                Utility.MakeJSArray(New String() {Utility.MakeJSIndexedObject(New String() {"Symbol", "Shaping", "Assimilate", "TranslitLetter", "RomanTranslit", "PlainRoman"}, _
                                Array.ConvertAll(Of IslamData.ArabicSymbol, String())(Letters, Function(Convert As IslamData.ArabicSymbol) New String() {CStr(AscW(Convert.Symbol)), If(Convert.Shaping = Nothing, String.Empty, Utility.MakeJSArray(Array.ConvertAll(Convert.Shaping, Function(Ch As Char) CStr(AscW(Ch))))), CStr(IIf(Convert.Assimilate, "true", String.Empty)), CStr(IIf(Convert.ExtendedBuckwalterLetter = ChrW(0), String.Empty, Convert.ExtendedBuckwalterLetter)), GetSchemeValueFromSymbol(Convert, "RomanTranslit"), GetSchemeValueFromSymbol(Convert, "PlainRoman")}), False)}, True) + ";"
    End Function
    Public Shared FindLetterBySymbolJS As String = "function findLetterBySymbol(chVal) { var iSubCount; for (iSubCount = 0; iSubCount < arabicLetters.length; iSubCount++) { if (chVal === parseInt(arabicLetters[iSubCount].Symbol, 10)) return iSubCount; for (var iShapeCount = 0; iShapeCount < arabicLetters[iSubCount].Shaping.length; iShapeCount++) { if (chVal === parseInt(arabicLetters[iSubCount].Shaping[iShapeCount], 10)) return iSubCount; } } return -1; }"
    Public Shared TransliterateGenJS As String() = {
        FindLetterBySymbolJS,
        "function isLetterDiacritic(index) { return (" + String.Join("||", Array.ConvertAll(Arabic.RecitationLettersDiacritics, Function(C As Char) "parseInt(arabicLetters[index].Symbol, 10) === 0x" + Hex(AscW(C)))) + "); }", _
        "function isSpecialSymbol(index) { return (" + String.Join("||", Array.ConvertAll(Arabic.RecitationSpecialSymbols, Function(C As Char) "parseInt(arabicLetters[index].Symbol, 10) === 0x" + Hex(AscW(C)))) + "); }", _
        "function isCombiningSymbol(index) { return (" + String.Join("||", Array.ConvertAll(Arabic.RecitationCombiningSymbols, Function(C As Char) "parseInt(arabicLetters[index].Symbol, 10) === 0x" + Hex(AscW(C)))) + "); }", _
        "function doTransliterate(sVal, direction, conversion) { var iCount, iSubCount, sOutVal = ''; for (iCount = 0; iCount < sVal.length; iCount++) { if (sVal.charAt(iCount) === '\\') { iCount++; if (sVal.charAt(iCount) === ',') { sOutVal += String.fromCharCode(1548); } else { sOutVal += sVal.charAt(iCount); } } else { for (iSubCount = 0; iSubCount < arabicLetters.length; iSubCount++) { if (direction ? sVal.charCodeAt(iCount) === parseInt(arabicLetters[iSubCount].Symbol, 10) : sVal.charAt(iCount) === unescape((conversion ? arabicLetters[iSubCount].TranslitLetter : arabicLetters[iSubCount].RomanTranslit))) { sOutVal += (direction ? (conversion ? arabicLetters[iSubCount].TranslitLetter : arabicLetters[iSubCount].RomanTranslit) : ((isCombiningSymbol(iSubCount) && (iSubCount === 0 || findLetterBySymbol(sOutVal.charCodeAt(sOutVal.length - 1)) === -1 || !isLetterDiacritic(findLetterBySymbol(sOutVal.charCodeAt(sOutVal.length - 1))) && !isSpecialSymbol(findLetterBySymbol(sOutVal.charCodeAt(sOutVal.length - 1))))) ? String.fromCharCode(0x202D) + String.fromCharCode(0x25CC) : '') + String.fromCharCode(arabicLetters[iSubCount].Symbol) + ((isCombiningSymbol(iSubCount) && (iSubCount === 0 || findLetterBySymbol(sOutVal.charCodeAt(sOutVal.length - 1)) === -1 || !isLetterDiacritic(findLetterBySymbol(sOutVal.charCodeAt(sOutVal.length - 1))) && !isSpecialSymbol(findLetterBySymbol(sOutVal.charCodeAt(sOutVal.length - 1))))) ? String.fromCharCode(0x202C) : '')); break; } } if (iSubCount === arabicLetters.length) sOutVal += sVal.charAt(iCount); } } return unescape(sOutVal); }"
    }
    Public Shared IsDiacriticJS As String = "function isDiacritic(index) { return (" + String.Join("||", Array.ConvertAll(Arabic.RecitationDiacritics, Function(C As Char) "parseInt(arabicLetters[index].Symbol, 10) === 0x" + Hex(AscW(C)))) + "); }"
    Public Shared DiacriticJS As String() =
        {"function doDiacritics(sVal, direction) { var iCount, sOutVal = ''; for (iCount = 0; iCount < sVal.length; iCount++) { sOutVal += findLetterBySymbol(sVal.charCodeAt(iCount)) === -1 || !isDiacritic(findLetterBySymbol(sVal.charCodeAt(iCount))) ? sVal[iCount] : ''; } return sOutVal; }", _
            IsDiacriticJS, _
            FindLetterBySymbolJS}
    Public Shared PlainTransliterateGenJS As String() = {FindLetterBySymbolJS, IsDiacriticJS, _
            "function isWhitespace(index) { return (" + String.Join("||", Array.ConvertAll(Arabic.WhitespaceSymbols, Function(C As Char) "parseInt(arabicLetters[index].Symbol, 10) === 0x" + Hex(AscW(C)))) + "); }", _
            "function isPunctuation(index) { return (" + String.Join("||", Array.ConvertAll(Arabic.PunctuationSymbols, Function(C As Char) "parseInt(arabicLetters[index].Symbol, 10) === 0x" + Hex(AscW(C)))) + "); }", _
            "function isStop(index) { return (" + String.Join("||", Array.ConvertAll(Arabic.ArabicStopLetters, Function(C As Char) "parseInt(arabicLetters[index].Symbol, 10) === 0x" + Hex(AscW(C)))) + "); }", _
            "function applyColorRules(sVal) {}", _
            "function changeScript(sVal, scriptType) {}", _
            "function arabicLetterSpelling(sVal) { var count, index, output = ''; for (count = 0; count < sVal.length; count++) { index = findLetterBySymbol(sVal.charCodeAt(count)); if (index !== -1 && isLetter(index)) { if (output !== '') output += ' '; output += arabicLetters[index].SymbolName; } else if (index !== -1 && arabicLetters[index].Symbol === 1619) { output += sVal.charCodeAt(count); } } return doTransliterate(output, false, true); }", _
            "String.prototype.format = function() { var formatted = this; for (var i = 0; i < arguments.length; i++) { formatted = formatted.replace(new RegExp('\\{'+i+'\\}', 'gi'), arguments[i]); } return formatted; };", _
            "RegExp.matchResult = function(subexp, match, offset, str) { var args = arguments; return subexp.replace(/\$(\$|&|`|\'|[0-9]+)/g, function(m, p) { if (p === '$') return '$'; if (p === '`') return str.slice(0, offset); if (p === '\'') return str.slice(offset + match.length); if (p === '&' || parseInt(p, 10) <= 0 || parseInt(p, 10) >= args.length - 3) return match; return args[3 + parseInt(p, 10)]; }); };", _
            "var ruleFunctions = [function(str) { return [str.toUpperCase()]; }, function(str) { return [transliterateWithRules(doTransliterate(arabicWordFromNumber(parseInt(doTransliterate(str, true, true), 10), true, false, false), false, true))]; }, function(str) { return [arabicLetterSpelling(str)]; }, function(str) { return [arabicLetters[findLetterBySymbol(str.charCodeAt(0))].PlainRoman]; }, function(str) { return [" + Utility.MakeJSArray(ArabicFathaDammaKasra) + "[" + Utility.MakeJSArray(ArabicTanweens) + ".indexOf(str)], '" + ArabicLetterNoon + "']; }, function (str) { return ['', '']; }];", _
            "function isLetter(index) { return (" + String.Join("||", Array.ConvertAll(Arabic.RecitationLetters, Function(C As Char) "parseInt(arabicLetters[index].Symbol, 10) === 0x" + Hex(AscW(C)))) + "); }", _
            "function isSymbol(index) { return (" + String.Join("||", Array.ConvertAll(Arabic.GetRecitationSymbols(), Function(A As Array) "parseInt(arabicLetters[index].Symbol, 10) === 0x" + Hex(AscW(CachedData.IslamData.ArabicLetters(A(1)).Symbol)))) + "); }", _
            "var uthmaniMinimalScript = " + Utility.MakeJSArray(New String() {Utility.MakeJSIndexedObject(New String() {"rule", "match", "evaluator", "ruleFunc"}, _
                                    Array.ConvertAll(Of RuleTranslation, Object())(UthmaniMinimalScript, Function(Convert As RuleTranslation) New Object() {Utility.MakeJSString(Convert.Rule), Utility.MakeJSString(Utility.EscapeJS(Convert.Match)), Utility.MakeJSString(Convert.Evaluator), Convert.RuleFunc}), True)}, True) + ";", _
            "var simpleEnhancedScript = " + Utility.MakeJSArray(New String() {Utility.MakeJSIndexedObject(New String() {"rule", "match", "evaluator", "ruleFunc"}, _
                                    Array.ConvertAll(Of RuleTranslation, Object())(SimpleEnhancedScript, Function(Convert As RuleTranslation) New Object() {Utility.MakeJSString(Convert.Rule), Utility.MakeJSString(Utility.EscapeJS(Convert.Match)), Utility.MakeJSString(Convert.Evaluator), Convert.RuleFunc}), True)}, True) + ";", _
            "var simpleScript = " + Utility.MakeJSArray(New String() {Utility.MakeJSIndexedObject(New String() {"rule", "match", "evaluator", "ruleFunc"}, _
                                    Array.ConvertAll(Of RuleTranslation, Object())(SimpleScript, Function(Convert As RuleTranslation) New Object() {Utility.MakeJSString(Convert.Rule), Utility.MakeJSString(Utility.EscapeJS(Convert.Match)), Utility.MakeJSString(Convert.Evaluator), Convert.RuleFunc}), True)}, True) + ";", _
            "var simpleMinimalScript = " + Utility.MakeJSArray(New String() {Utility.MakeJSIndexedObject(New String() {"rule", "match", "evaluator", "ruleFunc"}, _
                                    Array.ConvertAll(Of RuleTranslation, Object())(SimpleMinimalScript, Function(Convert As RuleTranslation) New Object() {Utility.MakeJSString(Convert.Rule), Utility.MakeJSString(Utility.EscapeJS(Convert.Match)), Utility.MakeJSString(Convert.Evaluator), Convert.RuleFunc}), True)}, True) + ";", _
            "var simpleCleanScript = " + Utility.MakeJSArray(New String() {Utility.MakeJSIndexedObject(New String() {"rule", "match", "evaluator", "ruleFunc"}, _
                                    Array.ConvertAll(Of RuleTranslation, Object())(SimpleCleanScript, Function(Convert As RuleTranslation) New Object() {Utility.MakeJSString(Convert.Rule), Utility.MakeJSString(Utility.EscapeJS(Convert.Match)), Utility.MakeJSString(Convert.Evaluator), Convert.RuleFunc}), True)}, True) + ";", _
            "var errorCheckRules = " + Utility.MakeJSArray(New String() {Utility.MakeJSIndexedObject(New String() {"rule", "match", "evaluator", "ruleFunc"}, _
                                    Array.ConvertAll(Of RuleTranslation, Object())(ErrorCheckRules, Function(Convert As RuleTranslation) New Object() {Utility.MakeJSString(Convert.Rule), Utility.MakeJSString(Utility.EscapeJS(Convert.Match)), Utility.MakeJSString(Convert.Evaluator), Convert.RuleFunc}), True)}, True) + ";", _
            "var coloringSpelledOutRules = " + Utility.MakeJSArray(New String() {Utility.MakeJSIndexedObject(New String() {"rule", "match", "evaluator", "ruleFunc"}, _
                                    Array.ConvertAll(Of RuleTranslation, Object())(ColoringSpelledOutRules, Function(Convert As RuleTranslation) New Object() {Utility.MakeJSString(Convert.Rule), Utility.MakeJSString(Utility.EscapeJS(Convert.Match)), Utility.MakeJSString(Convert.Evaluator), Convert.RuleFunc}), True)}, True) + ";", _
            "var romanizationRules = " + Utility.MakeJSArray(New String() {Utility.MakeJSIndexedObject(New String() {"rule", "match", "evaluator", "ruleFunc"}, _
                                    Array.ConvertAll(Of RuleTranslation, Object())(RomanizationRules, Function(Convert As RuleTranslation) New Object() {Utility.MakeJSString(Convert.Rule), Utility.MakeJSString(Utility.EscapeJS(Convert.Match)), Utility.MakeJSString(Convert.Evaluator), Convert.RuleFunc}), True)}, True) + ";", _
            "var coloringRules = " + Utility.MakeJSArray(New String() {Utility.MakeJSIndexedObject(New String() {"rule", "match", "color"}, _
                                    Array.ConvertAll(Of ColorRule, String())(ColoringRules, Function(Convert As ColorRule) New String() {Convert.Rule, Utility.EscapeJS(Convert.Match), System.Drawing.ColorTranslator.ToHtml(Convert.Color)}), False)}, True) + ";", _
            "var rulesOfRecitationRegEx = " + Utility.MakeJSArray(New String() {Utility.MakeJSIndexedObject(New String() {"rule", "match", "evaluator"}, _
                                    Array.ConvertAll(Of RuleMetadataTranslation, Object())(RulesOfRecitationRegEx, Function(Convert As RuleMetadataTranslation) New Object() {Utility.MakeJSString(Convert.Rule), Utility.MakeJSString(Utility.EscapeJS(Convert.Match)), If(Convert.Evaluator Is Nothing, Nothing, Utility.MakeJSArray(Convert.Evaluator))}), True)}, True) + ";", _
            "var allowZeroLength = " + Utility.MakeJSArray(AllowZeroLength) + ";", _
            "function ruleMetadataComparer(a, b) { return (a.index === b.index) ? b.length - a.length : b.index - a.index; }", _
            "function replaceMetadata(sVal, metadataRule) { var count, elimParen = function(s) { return s.replace(/\(.*\)/, ''); }; for (count = 0; count < coloringSpelledOutRules.length; count++) { var index, match = null; for (index = 0; index < coloringSpelledOutRules[count].match.split('|').length; index++) { if (metadataRule.type.split('|').map(elimParen).indexOf(coloringSpelledOutRules[count].match.split('|')[index]) !== -1) { match = coloringSpelledOutRules[count].match.split('|')[index]; break; } } if (match !== null) { var str = coloringSpelledOutRules[count].evaluator.format(sVal.substr(metadataRule.index, metadataRule.length)); if (coloringSpelledOutRules[count].ruleFunc !== 0) { var args = ruleFunctions[coloringSpelledOutRules[count].ruleFunc - 1](str); if (args.length === 1) { str = args[0]; } else { var metaArgs = metadataRule.type.match(/\((.*)\)/)[1].split(','); str = ''; for (index = 0; index < args.length; index++) { str += replaceMetadata(args[index], {index: 0, length: args[index].length, type: metaArgs[index].replace(' ', '|')}); } } } sVal = sVal.substr(0, metadataRule.index) + str + sVal.substr(metadataRule.index + metadataRule.length); } } return sVal; }", _
            "function transliterateWithRules(sVal) { var count, index, arr, re, metadataList = [], replaceFunc = function(f, e) { return function(m) { return f(RegExp.matchResult(e, m, arguments[arguments.length - 2], arguments[arguments.length - 1]))[0]; }; }; sVal = sVal.replace(/\u200F/g, ''); for (count = 0; count < errorCheckRules.length; count++) { re = new RegExp(errorCheckRules[count].match, 'g'); while ((arr = re.exec(sVal)) !== null) { console.log(errorCheckRules[count].rule + ': ' + doTransliterate(sVal.substr(0, arr.index), true, true) + '<!-- -->' + doTransliterate(sVal.substr(arr.index), true, true)); } } for (count = 0; count < rulesOfRecitationRegEx.length; count++) { if (rulesOfRecitationRegEx[count].evaluator !== null) { var subcount, lindex; re = new RegExp(rulesOfRecitationRegEx[count].match, 'g'); while ((arr = re.exec(sVal)) !== null) { lindex = arr.index; for (subcount = 0; subcount < rulesOfRecitationRegEx[count].evaluator.length; subcount++) { if (rulesOfRecitationRegEx[count].evaluator[subcount] !== null && (arr[subcount + 1] && arr[subcount + 1].length !== 0 || allowZeroLength.indexOf(rulesOfRecitationRegEx[count].evaluator[subcount]) !== -1)) { metadataList.push({index: lindex, length: arr[subcount + 1] ? arr[subcount + 1].length : 0, type: rulesOfRecitationRegEx[count].evaluator[subcount]}); } lindex += (arr[subcount + 1] ? arr[subcount + 1].length : 0); } } } } metadataList.sort(ruleMetadataComparer); for (index = 0; index < metadataList.length; index++) { sVal = replaceMetadata(sVal, metadataList[index]); } for (count = 0; count < romanizationRules.length; count++) { sVal = sVal.replace(new RegExp(romanizationRules[count].match, 'g'), (romanizationRules[count].ruleFunc === 0) ? romanizationRules[count].evaluator : replaceFunc(ruleFunctions[romanizationRules[count].ruleFunc - 1], romanizationRules[count].evaluator)); } return sVal; }"}
    Public Shared NumberGenJS As String() = {"var arabicOrdinalNumbers = " + Utility.MakeJSArray(ArabicOrdinalNumbers) + ";", _
                "var arabicOrdinalExtraNumbers = " + Utility.MakeJSArray(ArabicOrdinalExtraNumbers) + ";", _
                "var arabicFractionNumbers = " + Utility.MakeJSArray(ArabicFractionNumbers) + ";", _
                "var arabicBaseNumbers = " + Utility.MakeJSArray(ArabicBaseNumbers) + ";", _
                "var arabicBaseExtraNumbers = " + Utility.MakeJSArray(ArabicBaseExtraNumbers) + ";", _
                "var arabicBaseTenNumbers = " + Utility.MakeJSArray(ArabicBaseTenNumbers) + ";", _
                "var arabicBaseHundredNumbers = " + Utility.MakeJSArray(ArabicBaseHundredNumbers) + ";", _
                "var arabicBaseThousandNumbers = " + Utility.MakeJSArray(ArabicBaseThousandNumbers) + ";", _
                "var arabicBaseMillionNumbers = " + Utility.MakeJSArray(ArabicBaseMillionNumbers) + ";", _
                "var arabicBaseBillionNumbers = " + Utility.MakeJSArray(ArabicBaseBillionNumbers) + ";", _
                "var arabicBaseMilliardNumbers = " + Utility.MakeJSArray(ArabicBaseMilliardNumbers) + ";", _
                "var arabicBaseTrillionNumbers = " + Utility.MakeJSArray(ArabicBaseTrillionNumbers) + ";", _
                "function doTransliterateNum() { $('#translitvalue').text(doTransliterate(arabicWordFromNumber($('#translitedit').val(), $('#useclassic0').prop('checked'), $('#usehundredform0').prop('checked'), $('#usemilliard0').prop('checked')), false, true)); }", _
                "function arabicWordForLessThanThousand(number, useclassic, usealefhundred) { var str = '', hundstr = ''; if (number >= 100) { hundstr = usealefhundred ? arabicBaseHundredNumbers[Math.floor(number / 100) - 1].substr(0, 2) + 'A' + arabicBaseHundredNumbers[Math.floor(number / 100) - 1].substr(2) : arabicBaseHundredNumbers[Math.floor(number / 100) - 1]; if ((number % 100) === 0) { return hundstr; } number = number % 100; } if ((number % 10) !== 0 && number !== 11 && number !== 12) { str = arabicBaseNumbers[number % 10 - 1]; } if (number >= 11 && number < 20) { if (number == 11 || number == 12) { str += arabicBaseExtraNumbers[number - 11]; } else { str = str.slice(0, -1) + 'a'; } str += ' Ea$ara'; } else if ((number === 0 && str === '') || number === 10 || number >= 20) { str = ((str === '') ? '' : str + ' wa') + arabicBaseTenNumbers[Math.floor(number / 10)]; } return useclassic ? (((str === '') ? '' : str + ((hundstr === '') ? '' : ' wa')) + hundstr) : (((hundstr === '') ? '' : hundstr + ((str === '') ? '' : ' wa')) + str); }", _
                "function arabicWordFromNumber(number, useclassic, usealefhundred, usemilliard) { var str = '', nextstr = '', curbase = 3, basenums = [1000, 1000000, 1000000000, 1000000000000], bases = [arabicBaseThousandNumbers, arabicBaseMillionNumbers, usemilliard ? arabicBaseMilliardNumbers : arabicBaseBillionNumbers, arabicBaseTrillionNumbers]; do { if (number >= basenums[curbase] && number < 2 * basenums[curbase]) { nextstr = bases[curbase][0]; } else if (number >= 2 * basenums[curbase] && number < 3 * basenums[curbase]) { nextstr = bases[curbase][1]; } else if (number >= 3 * basenums[curbase] && number < 10 * basenums[curbase]) { nextstr = arabicBaseNumbers[Math.floor(Number / basenums[curbase]) - 1].slice(0, -1) + 'u ' + bases[curbase][2]; } else if (number >= 10 * basenums[curbase] && number < 11 * basenums[curbase]) { nextstr = arabicBaseTenNumbers[1].slice(0, -1) + 'u ' + bases[curbase][2]; } else if (number >= basenums[curbase]) { nextstr = arabicWordForLessThanThousand(Math.floor(number / basenums[curbase]) % 100, useclassic, usealefhundred); if (number >= 100 * basenums[curbase] && number < (useclassic ? 200 : 101) * basenums[curbase]) { nextstr = nextstr.slice(0, -1) + 'u ' + bases[curbase][0].slice(0, -1) + 'K'; } else if (number >= 200 * basenums[curbase] && number < (useclassic ? 300 : 201) * basenums[curbase]) { nextstr = nextstr.slice(0, -2) + ' ' + bases[curbase][0].slice(0, -1) + 'K'; } else if (number >= 300 * basenums[curbase] && (useclassic || Math.floor(number / basenums[curbase]) % 100 === 0)) { nextstr = nextstr.slice(0, -1) + 'i ' + bases[curbase][0].slice(0, -1) + 'K'; } else { nextstr += ' ' + bases[curbase][0].slice(0, -1) + 'FA'; } } number = number % basenums[curbase]; curbase--; str = useclassic ? (((nextstr === '') ? '' : nextstr + ((str === '') ? '' : ' wa')) + str) : (((str === '') ? '' : str + ((nextstr === '') ? '' : ' wa')) + nextstr); nextstr = ''; } while (curbase >= 0); if (number !== 0 || str === '') { nextstr = arabicWordForLessThanThousand(number, useclassic, usealefhundred); } return useclassic ? (((nextstr === '') ? '' : nextstr + ((str === '') ? '' : ' wa')) + str) : (((str === '') ? '' : str + ((nextstr === '') ? '' : ' wa')) + nextstr); }"}
    Public Shared Function GetTransliterateNumberJS() As String()
        Dim GetJS As New List(Of String) From {"javascript: doTransliterateNum();", String.Empty, GetArabicSymbolJSArray()}
        GetJS.AddRange(TransliterateGenJS)
        GetJS.AddRange(NumberGenJS)
        Return GetJS.ToArray()
    End Function
    Public Shared Function GetTransliterateJS() As String()
        Dim GetJS As New List(Of String) From {"javascript: doTransliterateDisplay();", String.Empty, GetArabicSymbolJSArray(), _
        "function IsLTR(c) { return (c>=0x41&&c<=0x5A)||(c>=0x61&&c<=0x7A)||c===0xAA||c===0xB5||c===0xBA||(c>=0xC0&&c<=0xD6)||(c>=0xD8&&c<=0xF6)||(c>=0xF8&&c<=0x2B8)||(c>=0x2BB&&c<=0x2C1)||(c>=0x2D0&&c<=0x2D1)||(c>=0x2E0&&c<=0x2E4)||c===0x2EE||(c>=0x370&&c<=0x373)||(c>=0x376&&c<=0x377)||(c>=0x37A&&c<=0x37D)||c===0x37F||c===0x386||(c>=0x388&&c<=0x38A)||c===0x38C||(c>=0x38E&&c<=0x3A1)||(c>=0x3A3&&c<=0x3F5)||(c>=0x3F7&&c<=0x482)||(c>=0x48A&&c<=0x52F)||(c>=0x531&&c<=0x556)||(c>=0x559&&c<=0x55F)||(c>=0x561&&c<=0x587)||c===0x589||(c>=0x903&&c<=0x939)||c===0x93B||(c>=0x93D&&c<=0x940)||(c>=0x949&&c<=0x94C)||(c>=0x94E&&c<=0x950)||(c>=0x958&&c<=0x961)||(c>=0x964&&c<=0x980)||(c>=0x982&&c<=0x983)||(c>=0x985&&c<=0x98C)||(c>=0x98F&&c<=0x990)||(c>=0x993&&c<=0x9A8)||(c>=0x9AA&&c<=0x9B0)||c===0x9B2||(c>=0x9B6&&c<=0x9B9)||(c>=0x9BD&&c<=0x9C0)||(c>=0x9C7&&c<=0x9C8)||(c>=0x9CB&&c<=0x9CC)||c===0x9CE||c===0x9D7||(c>=0x9DC&&c<=0x9DD)||(c>=0x9DF&&c<=0x9E1)||(c>=0x9E6&&c<=0x9F1)||(c>=0x9F4&&c<=0x9FA)||c===0xA03||(c>=0xA05&&c<=0xA0A)||(c>=0xA0F&&c<=0xA10)||(c>=0xA13&&c<=0xA28)||(c>=0xA2A&&c<=0xA30)||(c>=0xA32&&c<=0xA33)||(c>=0xA35&&c<=0xA36)||(c>=0xA38&&c<=0xA39)||(c>=0xA3E&&c<=0xA40)||(c>=0xA59&&c<=0xA5C)||c===0xA5E||(c>=0xA66&&c<=0xA6F)||(c>=0xA72&&c<=0xA74)||c===0xA83||(c>=0xA85&&c<=0xA8D)||(c>=0xA8F&&c<=0xA91)||(c>=0xA93&&c<=0xAA8)||(c>=0xAAA&&c<=0xAB0)||(c>=0xAB2&&c<=0xAB3)||(c>=0xAB5&&c<=0xAB9)||(c>=0xABD&&c<=0xAC0)||c===0xAC9||(c>=0xACB&&c<=0xACC)||c===0xAD0||(c>=0xAE0&&c<=0xAE1)||(c>=0xAE6&&c<=0xAF0)||(c>=0xB02&&c<=0xB03)||(c>=0xB05&&c<=0xB0C)||(c>=0xB0F&&c<=0xB10)||(c>=0xB13&&c<=0xB28)||(c>=0xB2A&&c<=0xB30)||(c>=0xB32&&c<=0xB33)||(c>=0xB35&&c<=0xB39)||(c>=0xB3D&&c<=0xB3E)||c===0xB40||(c>=0xB47&&c<=0xB48)||(c>=0xB4B&&c<=0xB4C)||c===0xB57||(c>=0xB5C&&c<=0xB5D)||(c>=0xB5F&&c<=0xB61)||(c>=0xB66&&c<=0xB77)||c===0xB83||(c>=0xB85&&c<=0xB8A)||(c>=0xB8E&&c<=0xB90)||(c>=0xB92&&c<=0xB95)||(c>=0xB99&&c<=0xB9A)||c===0xB9C||(c>=0xB9E&&c<=0xB9F)||(c>=0xBA3&&c<=0xBA4)||(c>=0xBA8&&c<=0xBAA)||(c>=0xBAE&&c<=0xBB9)||(c>=0xBBE&&c<=0xBBF)||(c>=0xBC1&&c<=0xBC2)||(c>=0xBC6&&c<=0xBC8)||(c>=0xBCA&&c<=0xBCC)||c===0xBD0||c===0xBD7||(c>=0xBE6&&c<=0xBF2)||(c>=0xC01&&c<=0xC03)||(c>=0xC05&&c<=0xC0C)||(c>=0xC0E&&c<=0xC10)||(c>=0xC12&&c<=0xC28)||(c>=0xC2A&&c<=0xC39)||c===0xC3D||(c>=0xC41&&c<=0xC44)||(c>=0xC58&&c<=0xC59)||(c>=0xC60&&c<=0xC61)||(c>=0xC66&&c<=0xC6F)||c===0xC7F||(c>=0xC82&&c<=0xC83)||(c>=0xC85&&c<=0xC8C)||(c>=0xC8E&&c<=0xC90)||(c>=0xC92&&c<=0xCA8)||(c>=0xCAA&&c<=0xCB3)||(c>=0xCB5&&c<=0xCB9)||(c>=0xCBD&&c<=0xCC4)||(c>=0xCC6&&c<=0xCC8)||(c>=0xCCA&&c<=0xCCB)||(c>=0xCD5&&c<=0xCD6)||c===0xCDE||(c>=0xCE0&&c<=0xCE1)||(c>=0xCE6&&c<=0xCEF)||(c>=0xCF1&&c<=0xCF2)||(c>=0xD02&&c<=0xD03)||(c>=0xD05&&c<=0xD0C)||(c>=0xD0E&&c<=0xD10)||(c>=0xD12&&c<=0xD3A)||(c>=0xD3D&&c<=0xD40)||(c>=0xD46&&c<=0xD48)||(c>=0xD4A&&c<=0xD4C)||c===0xD4E||c===0xD57||(c>=0xD60&&c<=0xD61)||(c>=0xD66&&c<=0xD75)||(c>=0xD79&&c<=0xD7F)||(c>=0xD82&&c<=0xD83)||(c>=0xD85&&c<=0xD96)||(c>=0xD9A&&c<=0xDB1)||(c>=0xDB3&&c<=0xDBB)||c===0xDBD||(c>=0xDC0&&c<=0xDC6)||(c>=0xDCF&&c<=0xDD1)||(c>=0xDD8&&c<=0xDDF)||(c>=0xDE6&&c<=0xDEF)||(c>=0xDF2&&c<=0xDF4)||(c>=0xE01&&c<=0xE30)||(c>=0xE32&&c<=0xE33)||(c>=0xE40&&c<=0xE46)||(c>=0xE4F&&c<=0xE5B)||(c>=0xE81&&c<=0xE82)||c===0xE84||(c>=0xE87&&c<=0xE88)||c===0xE8A||c===0xE8D||(c>=0xE94&&c<=0xE97)||(c>=0xE99&&c<=0xE9F)||(c>=0xEA1&&c<=0xEA3)||c===0xEA5||c===0xEA7||(c>=0xEAA&&c<=0xEAB)||(c>=0xEAD&&c<=0xEB0)||(c>=0xEB2&&c<=0xEB3)||c===0xEBD||(c>=0xEC0&&c<=0xEC4)||c===0xEC6||(c>=0xED0&&c<=0xED9)||(c>=0xEDC&&c<=0xEDF)||(c>=0xF00&&c<=0xF17)||(c>=0xF1A&&c<=0xF34)||c===0xF36||c===0xF38||(c>=0xF3E&&c<=0xF47)||(c>=0xF49&&c<=0xF6C)||c===0xF7F||c===0xF85||(c>=0xF88&&c<=0xF8C)||(c>=0xFBE&&c<=0xFC5)||(c>=0xFC7&&c<=0xFCC)||(c>=0xFCE&&c<=0xFDA)||(c>=0x1000&&c<=0x102C)||c===0x1031||c===0x1038||(c>=0x103B&&c<=0x103C)||(c>=0x103F&&c<=0x1057)||(c>=0x105A&&c<=0x105D)||(c>=0x1061&&c<=0x1070)||(c>=0x1075&&c<=0x1081)||(c>=0x1083&&c<=0x1084)||(c>=0x1087&&c<=0x108C)||(c>=0x108E&&c<=0x109C)||(c>=0x109E&&c<=0x10C5)||c===0x10C7||c===0x10CD||(c>=0x10D0&&c<=0x1248)||(c>=0x124A&&c<=0x124D)||(c>=0x1250&&c<=0x1256)||c===0x1258||(c>=0x125A&&c<=0x125D)||(c>=0x1260&&c<=0x1288)||(c>=0x128A&&c<=0x128D)||(c>=0x1290&&c<=0x12B0)||(c>=0x12B2&&c<=0x12B5)||(c>=0x12B8&&c<=0x12BE)||c===0x12C0||(c>=0x12C2&&c<=0x12C5)||(c>=0x12C8&&c<=0x12D6)||(c>=0x12D8&&c<=0x1310)||(c>=0x1312&&c<=0x1315)||(c>=0x1318&&c<=0x135A)||(c>=0x1360&&c<=0x137C)||(c>=0x1380&&c<=0x138F)||(c>=0x13A0&&c<=0x13F4)||(c>=0x1401&&c<=0x167F)||(c>=0x1681&&c<=0x169A)||(c>=0x16A0&&c<=0x16F8)||(c>=0x1700&&c<=0x170C)||(c>=0x170E&&c<=0x1711)||(c>=0x1720&&c<=0x1731)||(c>=0x1735&&c<=0x1736)||(c>=0x1740&&c<=0x1751)||(c>=0x1760&&c<=0x176C)||(c>=0x176E&&c<=0x1770)||(c>=0x1780&&c<=0x17B3)||c===0x17B6||(c>=0x17BE&&c<=0x17C5)||(c>=0x17C7&&c<=0x17C8)||(c>=0x17D4&&c<=0x17DA)||c===0x17DC||(c>=0x17E0&&c<=0x17E9)||(c>=0x1810&&c<=0x1819)||(c>=0x1820&&c<=0x1877)||(c>=0x1880&&c<=0x18A8)||c===0x18AA||(c>=0x18B0&&c<=0x18F5)||(c>=0x1900&&c<=0x191E)||(c>=0x1923&&c<=0x1926)||(c>=0x1929&&c<=0x192B)||(c>=0x1930&&c<=0x1931)||(c>=0x1933&&c<=0x1938)||(c>=0x1946&&c<=0x196D)||(c>=0x1970&&c<=0x1974)||(c>=0x1980&&c<=0x19AB)||(c>=0x19B0&&c<=0x19C9)||(c>=0x19D0&&c<=0x19DA)||(c>=0x1A00&&c<=0x1A16)||(c>=0x1A19&&c<=0x1A1A)||(c>=0x1A1E&&c<=0x1A55)||c===0x1A57||c===0x1A61||(c>=0x1A63&&c<=0x1A64)||(c>=0x1A6D&&c<=0x1A72)||(c>=0x1A80&&c<=0x1A89)||(c>=0x1A90&&c<=0x1A99)||(c>=0x1AA0&&c<=0x1AAD)||(c>=0x1B04&&c<=0x1B33)||c===0x1B35||c===0x1B3B||(c>=0x1B3D&&c<=0x1B41)||(c>=0x1B43&&c<=0x1B4B)||(c>=0x1B50&&c<=0x1B6A)||(c>=0x1B74&&c<=0x1B7C)||(c>=0x1B82&&c<=0x1BA1)||(c>=0x1BA6&&c<=0x1BA7)||c===0x1BAA||(c>=0x1BAE&&c<=0x1BE5)||c===0x1BE7||(c>=0x1BEA&&c<=0x1BEC)||c===0x1BEE||(c>=0x1BF2&&c<=0x1BF3)||(c>=0x1BFC&&c<=0x1C2B)||(c>=0x1C34&&c<=0x1C35)||(c>=0x1C3B&&c<=0x1C49)||(c>=0x1C4D&&c<=0x1C7F)||(c>=0x1CC0&&c<=0x1CC7)||c===0x1CD3||c===0x1CE1||(c>=0x1CE9&&c<=0x1CEC)||(c>=0x1CEE&&c<=0x1CF3)||(c>=0x1CF5&&c<=0x1CF6)||(c>=0x1D00&&c<=0x1DBF)||(c>=0x1E00&&c<=0x1F15)||(c>=0x1F18&&c<=0x1F1D)||(c>=0x1F20&&c<=0x1F45)||(c>=0x1F48&&c<=0x1F4D)||(c>=0x1F50&&c<=0x1F57)||c===0x1F59||c===0x1F5B||c===0x1F5D||(c>=0x1F5F&&c<=0x1F7D)||(c>=0x1F80&&c<=0x1FB4)||(c>=0x1FB6&&c<=0x1FBC)||c===0x1FBE||(c>=0x1FC2&&c<=0x1FC4)||(c>=0x1FC6&&c<=0x1FCC)||(c>=0x1FD0&&c<=0x1FD3)||(c>=0x1FD6&&c<=0x1FDB)||(c>=0x1FE0&&c<=0x1FEC)||(c>=0x1FF2&&c<=0x1FF4)||(c>=0x1FF6&&c<=0x1FFC)||c===0x200E||c===0x2071||c===0x207F||(c>=0x2090&&c<=0x209C)||c===0x2102||c===0x2107||(c>=0x210A&&c<=0x2113)||c===0x2115||(c>=0x2119&&c<=0x211D)||c===0x2124||c===0x2126||c===0x2128||(c>=0x212A&&c<=0x212D)||(c>=0x212F&&c<=0x2139)||(c>=0x213C&&c<=0x213F)||(c>=0x2145&&c<=0x2149)||(c>=0x214E&&c<=0x214F)||(c>=0x2160&&c<=0x2188)||(c>=0x2336&&c<=0x237A)||c===0x2395||(c>=0x249C&&c<=0x24E9)||c===0x26AC||(c>=0x2800&&c<=0x28FF)||(c>=0x2C00&&c<=0x2C2E)||(c>=0x2C30&&c<=0x2C5E)||(c>=0x2C60&&c<=0x2CE4)||(c>=0x2CEB&&c<=0x2CEE)||(c>=0x2CF2&&c<=0x2CF3)||(c>=0x2D00&&c<=0x2D25)||c===0x2D27||c===0x2D2D||(c>=0x2D30&&c<=0x2D67)||(c>=0x2D6F&&c<=0x2D70)||(c>=0x2D80&&c<=0x2D96)||(c>=0x2DA0&&c<=0x2DA6)||(c>=0x2DA8&&c<=0x2DAE)||(c>=0x2DB0&&c<=0x2DB6)||(c>=0x2DB8&&c<=0x2DBE)||(c>=0x2DC0&&c<=0x2DC6)||(c>=0x2DC8&&c<=0x2DCE)||(c>=0x2DD0&&c<=0x2DD6)||(c>=0x2DD8&&c<=0x2DDE)||(c>=0x3005&&c<=0x3007)||(c>=0x3021&&c<=0x3029)||(c>=0x302E&&c<=0x302F)||(c>=0x3031&&c<=0x3035)||(c>=0x3038&&c<=0x303C)||(c>=0x3041&&c<=0x3096)||(c>=0x309D&&c<=0x309F)||(c>=0x30A1&&c<=0x30FA)||(c>=0x30FC&&c<=0x30FF)||(c>=0x3105&&c<=0x312D)||(c>=0x3131&&c<=0x318E)||(c>=0x3190&&c<=0x31BA)||(c>=0x31F0&&c<=0x321C)||(c>=0x3220&&c<=0x324F)||(c>=0x3260&&c<=0x327B)||(c>=0x327F&&c<=0x32B0)||(c>=0x32C0&&c<=0x32CB)||(c>=0x32D0&&c<=0x32FE)||(c>=0x3300&&c<=0x3376)||(c>=0x337B&&c<=0x33DD)||(c>=0x33E0&&c<=0x33FE)||c===0x3400||c===0x4DB5||c===0x4E00||c===0x9FCC||(c>=0xA000&&c<=0xA48C)||(c>=0xA4D0&&c<=0xA60C)||(c>=0xA610&&c<=0xA62B)||(c>=0xA640&&c<=0xA66E)||(c>=0xA680&&c<=0xA69D)||(c>=0xA6A0&&c<=0xA6EF)||(c>=0xA6F2&&c<=0xA6F7)||(c>=0xA722&&c<=0xA787)||(c>=0xA789&&c<=0xA78E)||(c>=0xA790&&c<=0xA7AD)||(c>=0xA7B0&&c<=0xA7B1)||(c>=0xA7F7&&c<=0xA801)||(c>=0xA803&&c<=0xA805)||(c>=0xA807&&c<=0xA80A)||(c>=0xA80C&&c<=0xA824)||c===0xA827||(c>=0xA830&&c<=0xA837)||(c>=0xA840&&c<=0xA873)||(c>=0xA880&&c<=0xA8C3)||(c>=0xA8CE&&c<=0xA8D9)||(c>=0xA8F2&&c<=0xA8FB)||(c>=0xA900&&c<=0xA925)||(c>=0xA92E&&c<=0xA946)||(c>=0xA952&&c<=0xA953)||(c>=0xA95F&&c<=0xA97C)||(c>=0xA983&&c<=0xA9B2)||(c>=0xA9B4&&c<=0xA9B5)||(c>=0xA9BA&&c<=0xA9BB)||(c>=0xA9BD&&c<=0xA9CD)||(c>=0xA9CF&&c<=0xA9D9)||(c>=0xA9DE&&c<=0xA9E4)||(c>=0xA9E6&&c<=0xA9FE)||(c>=0xAA00&&c<=0xAA28)||(c>=0xAA2F&&c<=0xAA30)||(c>=0xAA33&&c<=0xAA34)||(c>=0xAA40&&c<=0xAA42)||(c>=0xAA44&&c<=0xAA4B)||c===0xAA4D||(c>=0xAA50&&c<=0xAA59)||(c>=0xAA5C&&c<=0xAA7B)||(c>=0xAA7D&&c<=0xAAAF)||c===0xAAB1||(c>=0xAAB5&&c<=0xAAB6)||(c>=0xAAB9&&c<=0xAABD)||c===0xAAC0||c===0xAAC2||(c>=0xAADB&&c<=0xAAEB)||(c>=0xAAEE&&c<=0xAAF5)||(c>=0xAB01&&c<=0xAB06)||(c>=0xAB09&&c<=0xAB0E)||(c>=0xAB11&&c<=0xAB16)||(c>=0xAB20&&c<=0xAB26)||(c>=0xAB28&&c<=0xAB2E)||(c>=0xAB30&&c<=0xAB5F)||(c>=0xAB64&&c<=0xAB65)||(c>=0xABC0&&c<=0xABE4)||(c>=0xABE6&&c<=0xABE7)||(c>=0xABE9&&c<=0xABEC)||(c>=0xABF0&&c<=0xABF9)||c===0xAC00||c===0xD7A3||(c>=0xD7B0&&c<=0xD7C6)||(c>=0xD7CB&&c<=0xD7FB)||c===0xD800||(c>=0xDB7F&&c<=0xDB80)||(c>=0xDBFF&&c<=0xDC00)||(c>=0xDFFF&&c<=0xE000)||(c>=0xF8FF&&c<=0xFA6D)||(c>=0xFA70&&c<=0xFAD9)||(c>=0xFB00&&c<=0xFB06)||(c>=0xFB13&&c<=0xFB17)||(c>=0xFF21&&c<=0xFF3A)||(c>=0xFF41&&c<=0xFF5A)||(c>=0xFF66&&c<=0xFFBE)||(c>=0xFFC2&&c<=0xFFC7)||(c>=0xFFCA&&c<=0xFFCF)||(c>=0xFFD2&&c<=0xFFD7)||(c>=0xFFDA&&c<=0xFFDC)||(c>=0x10000&&c<=0x1000B)||(c>=0x1000D&&c<=0x10026)||(c>=0x10028&&c<=0x1003A)||(c>=0x1003C&&c<=0x1003D)||(c>=0x1003F&&c<=0x1004D)||(c>=0x10050&&c<=0x1005D)||(c>=0x10080&&c<=0x100FA)||c===0x10100||c===0x10102||(c>=0x10107&&c<=0x10133)||(c>=0x10137&&c<=0x1013F)||(c>=0x101D0&&c<=0x101FC)||(c>=0x10280&&c<=0x1029C)||(c>=0x102A0&&c<=0x102D0)||(c>=0x10300&&c<=0x10323)||(c>=0x10330&&c<=0x1034A)||(c>=0x10350&&c<=0x10375)||(c>=0x10380&&c<=0x1039D)||(c>=0x1039F&&c<=0x103C3)||(c>=0x103C8&&c<=0x103D5)||(c>=0x10400&&c<=0x1049D)||(c>=0x104A0&&c<=0x104A9)||(c>=0x10500&&c<=0x10527)||(c>=0x10530&&c<=0x10563)||c===0x1056F||(c>=0x10600&&c<=0x10736)||(c>=0x10740&&c<=0x10755)||(c>=0x10760&&c<=0x10767)||c===0x11000||(c>=0x11002&&c<=0x11037)||(c>=0x11047&&c<=0x1104D)||(c>=0x11066&&c<=0x1106F)||(c>=0x11082&&c<=0x110B2)||(c>=0x110B7&&c<=0x110B8)||(c>=0x110BB&&c<=0x110C1)||(c>=0x110D0&&c<=0x110E8)||(c>=0x110F0&&c<=0x110F9)||(c>=0x11103&&c<=0x11126)||c===0x1112C||(c>=0x11136&&c<=0x11143)||(c>=0x11150&&c<=0x11172)||(c>=0x11174&&c<=0x11176)||(c>=0x11182&&c<=0x111B5)||(c>=0x111BF&&c<=0x111C8)||c===0x111CD||(c>=0x111D0&&c<=0x111DA)||(c>=0x111E1&&c<=0x111F4)||(c>=0x11200&&c<=0x11211)||(c>=0x11213&&c<=0x1122E)||(c>=0x11232&&c<=0x11233)||c===0x11235||(c>=0x11238&&c<=0x1123D)||(c>=0x112B0&&c<=0x112DE)||(c>=0x112E0&&c<=0x112E2)||(c>=0x112F0&&c<=0x112F9)||(c>=0x11302&&c<=0x11303)||(c>=0x11305&&c<=0x1130C)||(c>=0x1130F&&c<=0x11310)||(c>=0x11313&&c<=0x11328)||(c>=0x1132A&&c<=0x11330)||(c>=0x11332&&c<=0x11333)||(c>=0x11335&&c<=0x11339)||(c>=0x1133D&&c<=0x1133F)||(c>=0x11341&&c<=0x11344)||(c>=0x11347&&c<=0x11348)||(c>=0x1134B&&c<=0x1134D)||c===0x11357||(c>=0x1135D&&c<=0x11363)||(c>=0x11480&&c<=0x114B2)||c===0x114B9||(c>=0x114BB&&c<=0x114BE)||c===0x114C1||(c>=0x114C4&&c<=0x114C7)||(c>=0x114D0&&c<=0x114D9)||(c>=0x11580&&c<=0x115B1)||(c>=0x115B8&&c<=0x115BB)||c===0x115BE||(c>=0x115C1&&c<=0x115C9)||(c>=0x11600&&c<=0x11632)||(c>=0x1163B&&c<=0x1163C)||c===0x1163E||(c>=0x11641&&c<=0x11644)||(c>=0x11650&&c<=0x11659)||(c>=0x11680&&c<=0x116AA)||c===0x116AC||(c>=0x116AE&&c<=0x116AF)||c===0x116B6||(c>=0x116C0&&c<=0x116C9)||(c>=0x118A0&&c<=0x118F2)||c===0x118FF||(c>=0x11AC0&&c<=0x11AF8)||(c>=0x12000&&c<=0x12398)||(c>=0x12400&&c<=0x1246E)||(c>=0x12470&&c<=0x12474)||(c>=0x13000&&c<=0x1342E)||(c>=0x16800&&c<=0x16A38)||(c>=0x16A40&&c<=0x16A5E)||(c>=0x16A60&&c<=0x16A69)||(c>=0x16A6E&&c<=0x16A6F)||(c>=0x16AD0&&c<=0x16AED)||c===0x16AF5||(c>=0x16B00&&c<=0x16B2F)||(c>=0x16B37&&c<=0x16B45)||(c>=0x16B50&&c<=0x16B59)||(c>=0x16B5B&&c<=0x16B61)||(c>=0x16B63&&c<=0x16B77)||(c>=0x16B7D&&c<=0x16B8F)||(c>=0x16F00&&c<=0x16F44)||(c>=0x16F50&&c<=0x16F7E)||(c>=0x16F93&&c<=0x16F9F)||(c>=0x1B000&&c<=0x1B001)||(c>=0x1BC00&&c<=0x1BC6A)||(c>=0x1BC70&&c<=0x1BC7C)||(c>=0x1BC80&&c<=0x1BC88)||(c>=0x1BC90&&c<=0x1BC99)||c===0x1BC9C||c===0x1BC9F||(c>=0x1D000&&c<=0x1D0F5)||(c>=0x1D100&&c<=0x1D126)||(c>=0x1D129&&c<=0x1D166)||(c>=0x1D16A&&c<=0x1D172)||(c>=0x1D183&&c<=0x1D184)||(c>=0x1D18C&&c<=0x1D1A9)||(c>=0x1D1AE&&c<=0x1D1DD)||(c>=0x1D360&&c<=0x1D371)||(c>=0x1D400&&c<=0x1D454)||(c>=0x1D456&&c<=0x1D49C)||(c>=0x1D49E&&c<=0x1D49F)||c===0x1D4A2||(c>=0x1D4A5&&c<=0x1D4A6)||(c>=0x1D4A9&&c<=0x1D4AC)||(c>=0x1D4AE&&c<=0x1D4B9)||c===0x1D4BB||(c>=0x1D4BD&&c<=0x1D4C3)||(c>=0x1D4C5&&c<=0x1D505)||(c>=0x1D507&&c<=0x1D50A)||(c>=0x1D50D&&c<=0x1D514)||(c>=0x1D516&&c<=0x1D51C)||(c>=0x1D51E&&c<=0x1D539)||(c>=0x1D53B&&c<=0x1D53E)||(c>=0x1D540&&c<=0x1D544)||c===0x1D546||(c>=0x1D54A&&c<=0x1D550)||(c>=0x1D552&&c<=0x1D6A5)||(c>=0x1D6A8&&c<=0x1D6DA)||(c>=0x1D6DC&&c<=0x1D714)||(c>=0x1D716&&c<=0x1D74E)||(c>=0x1D750&&c<=0x1D788)||(c>=0x1D78A&&c<=0x1D7C2)||(c>=0x1D7C4&&c<=0x1D7CB)||(c>=0x1F110&&c<=0x1F12E)||(c>=0x1F130&&c<=0x1F169)||(c>=0x1F170&&c<=0x1F19A)||(c>=0x1F1E6&&c<=0x1F202)||(c>=0x1F210&&c<=0x1F23A)||(c>=0x1F240&&c<=0x1F248)||(c>=0x1F250&&c<=0x1F251)||c===0x20000||c===0x2A6D6||c===0x2A700||c===0x2B734||c===0x2B740||c===0x2B81D||(c>=0x2F800&&c<=0x2FA1D)||c===0xF0000||c===0xFFFFD||c===0x100000||c===0x10FFFD; }", _
        "function IsRTL(c) { return c===0x5BE||c===0x5C0||c===0x5C3||c===0x5C6||(c>=0x5D0&&c<=0x5EA)||(c>=0x5F0&&c<=0x5F4)||c===0x608||c===0x60B||c===0x60D||(c>=0x61B&&c<=0x61C)||(c>=0x61E&&c<=0x64A)||(c>=0x66D&&c<=0x66F)||(c>=0x671&&c<=0x6D5)||(c>=0x6E5&&c<=0x6E6)||(c>=0x6EE&&c<=0x6EF)||(c>=0x6FA&&c<=0x70D)||(c>=0x70F&&c<=0x710)||(c>=0x712&&c<=0x72F)||(c>=0x74D&&c<=0x7A5)||c===0x7B1||(c>=0x7C0&&c<=0x7EA)||(c>=0x7F4&&c<=0x7F5)||c===0x7FA||(c>=0x800&&c<=0x815)||c===0x81A||c===0x824||c===0x828||(c>=0x830&&c<=0x83E)||(c>=0x840&&c<=0x858)||c===0x85E||(c>=0x8A0&&c<=0x8B2)||c===0x200F||c===0xFB1D||(c>=0xFB1F&&c<=0xFB28)||(c>=0xFB2A&&c<=0xFB36)||(c>=0xFB38&&c<=0xFB3C)||c===0xFB3E||(c>=0xFB40&&c<=0xFB41)||(c>=0xFB43&&c<=0xFB44)||(c>=0xFB46&&c<=0xFBC1)||(c>=0xFBD3&&c<=0xFD3D)||(c>=0xFD50&&c<=0xFD8F)||(c>=0xFD92&&c<=0xFDC7)||(c>=0xFDF0&&c<=0xFDFC)||(c>=0xFE70&&c<=0xFE74)||(c>=0xFE76&&c<=0xFEFC)||(c>=0x10800&&c<=0x10805)||c===0x10808||(c>=0x1080A&&c<=0x10835)||(c>=0x10837&&c<=0x10838)||c===0x1083C||(c>=0x1083F&&c<=0x10855)||(c>=0x10857&&c<=0x1089E)||(c>=0x108A7&&c<=0x108AF)||(c>=0x10900&&c<=0x1091B)||(c>=0x10920&&c<=0x10939)||c===0x1093F||(c>=0x10980&&c<=0x109B7)||(c>=0x109BE&&c<=0x109BF)||c===0x10A00||(c>=0x10A10&&c<=0x10A13)||(c>=0x10A15&&c<=0x10A17)||(c>=0x10A19&&c<=0x10A33)||(c>=0x10A40&&c<=0x10A47)||(c>=0x10A50&&c<=0x10A58)||(c>=0x10A60&&c<=0x10A9F)||(c>=0x10AC0&&c<=0x10AE4)||(c>=0x10AEB&&c<=0x10AF6)||(c>=0x10B00&&c<=0x10B35)||(c>=0x10B40&&c<=0x10B55)||(c>=0x10B58&&c<=0x10B72)||(c>=0x10B78&&c<=0x10B91)||(c>=0x10B99&&c<=0x10B9C)||(c>=0x10BA9&&c<=0x10BAF)||(c>=0x10C00&&c<=0x10C48)||(c>=0x1E800&&c<=0x1E8C4)||(c>=0x1E8C7&&c<=0x1E8CF)||(c>=0x1EE00&&c<=0x1EE03)||(c>=0x1EE05&&c<=0x1EE1F)||(c>=0x1EE21&&c<=0x1EE22)||c===0x1EE24||c===0x1EE27||(c>=0x1EE29&&c<=0x1EE32)||(c>=0x1EE34&&c<=0x1EE37)||c===0x1EE39||c===0x1EE3B||c===0x1EE42||c===0x1EE47||c===0x1EE49||c===0x1EE4B||(c>=0x1EE4D&&c<=0x1EE4F)||(c>=0x1EE51&&c<=0x1EE52)||c===0x1EE54||c===0x1EE57||c===0x1EE59||c===0x1EE5B||c===0x1EE5D||c===0x1EE5F||(c>=0x1EE61&&c<=0x1EE62)||c===0x1EE64||(c>=0x1EE67&&c<=0x1EE6A)||(c>=0x1EE6C&&c<=0x1EE72)||(c>=0x1EE74&&c<=0x1EE77)||(c>=0x1EE79&&c<=0x1EE7C)||c===0x1EE7E||(c>=0x1EE80&&c<=0x1EE89)||(c>=0x1EE8B&&c<=0x1EE9B)||(c>=0x1EEA1&&c<=0x1EEA3)||(c>=0x1EEA5&&c<=0x1EEA9)||(c>=0x1EEAB&&c<=0x1EEBB); }", _
        "function IsAL(c) { return c===0x608||c===0x60B||c===0x60D||(c>=0x61B&&c<=0x61C)||(c>=0x61E&&c<=0x64A)||(c>=0x66D&&c<=0x66F)||(c>=0x671&&c<=0x6D5)||(c>=0x6E5&&c<=0x6E6)||(c>=0x6EE&&c<=0x6EF)||(c>=0x6FA&&c<=0x70D)||(c>=0x70F&&c<=0x710)||(c>=0x712&&c<=0x72F)||(c>=0x74D&&c<=0x7A5)||c===0x7B1||(c>=0x8A0&&c<=0x8B2)||(c>=0xFB50&&c<=0xFBC1)||(c>=0xFBD3&&c<=0xFD3D)||(c>=0xFD50&&c<=0xFD8F)||(c>=0xFD92&&c<=0xFDC7)||(c>=0xFDF0&&c<=0xFDFC)||(c>=0xFE70&&c<=0xFE74)||(c>=0xFE76&&c<=0xFEFC)||(c>=0x1EE00&&c<=0x1EE03)||(c>=0x1EE05&&c<=0x1EE1F)||(c>=0x1EE21&&c<=0x1EE22)||c===0x1EE24||c===0x1EE27||(c>=0x1EE29&&c<=0x1EE32)||(c>=0x1EE34&&c<=0x1EE37)||c===0x1EE39||c===0x1EE3B||c===0x1EE42||c===0x1EE47||c===0x1EE49||c===0x1EE4B||(c>=0x1EE4D&&c<=0x1EE4F)||(c>=0x1EE51&&c<=0x1EE52)||c===0x1EE54||c===0x1EE57||c===0x1EE59||c===0x1EE5B||c===0x1EE5D||c===0x1EE5F||(c>=0x1EE61&&c<=0x1EE62)||c===0x1EE64||(c>=0x1EE67&&c<=0x1EE6A)||(c>=0x1EE6C&&c<=0x1EE72)||(c>=0x1EE74&&c<=0x1EE77)||(c>=0x1EE79&&c<=0x1EE7C)||c===0x1EE7E||(c>=0x1EE80&&c<=0x1EE89)||(c>=0x1EE8B&&c<=0x1EE9B)||(c>=0x1EEA1&&c<=0x1EEA3)||(c>=0x1EEA5&&c<=0x1EEA9)||(c>=0x1EEAB&&c<=0x1EEBB); }", _
        "function IsNeutral(c) { return (c>=0x9&&c<=0xD)||(c>=0x1C&&c<=0x22)||(c>=0x26&&c<=0x2A)||(c>=0x3B&&c<=0x40)||(c>=0x5B&&c<=0x60)||(c>=0x7B&&c<=0x7E)||c===0x85||c===0xA1||(c>=0xA6&&c<=0xA9)||(c>=0xAB&&c<=0xAC)||(c>=0xAE&&c<=0xAF)||c===0xB4||(c>=0xB6&&c<=0xB8)||(c>=0xBB&&c<=0xBF)||c===0xD7||c===0xF7||(c>=0x2B9&&c<=0x2BA)||(c>=0x2C2&&c<=0x2CF)||(c>=0x2D2&&c<=0x2DF)||(c>=0x2E5&&c<=0x2ED)||(c>=0x2EF&&c<=0x2FF)||(c>=0x374&&c<=0x375)||c===0x37E||(c>=0x384&&c<=0x385)||c===0x387||c===0x3F6||c===0x58A||(c>=0x58D&&c<=0x58E)||(c>=0x606&&c<=0x607)||(c>=0x60E&&c<=0x60F)||c===0x6DE||c===0x6E9||(c>=0x7F6&&c<=0x7F9)||(c>=0xBF3&&c<=0xBF8)||c===0xBFA||(c>=0xC78&&c<=0xC7E)||(c>=0xF3A&&c<=0xF3D)||(c>=0x1390&&c<=0x1399)||c===0x1400||c===0x1680||(c>=0x169B&&c<=0x169C)||(c>=0x17F0&&c<=0x17F9)||(c>=0x1800&&c<=0x180A)||c===0x1940||(c>=0x1944&&c<=0x1945)||(c>=0x19DE&&c<=0x19FF)||c===0x1FBD||(c>=0x1FBF&&c<=0x1FC1)||(c>=0x1FCD&&c<=0x1FCF)||(c>=0x1FDD&&c<=0x1FDF)||(c>=0x1FED&&c<=0x1FEF)||(c>=0x1FFD&&c<=0x1FFE)||(c>=0x2000&&c<=0x200A)||(c>=0x2010&&c<=0x2029)||(c>=0x2035&&c<=0x2043)||(c>=0x2045&&c<=0x205F)||(c>=0x207C&&c<=0x207E)||(c>=0x208C&&c<=0x208E)||(c>=0x2100&&c<=0x2101)||(c>=0x2103&&c<=0x2106)||(c>=0x2108&&c<=0x2109)||c===0x2114||(c>=0x2116&&c<=0x2118)||(c>=0x211E&&c<=0x2123)||c===0x2125||c===0x2127||c===0x2129||(c>=0x213A&&c<=0x213B)||(c>=0x2140&&c<=0x2144)||(c>=0x214A&&c<=0x214D)||(c>=0x2150&&c<=0x215F)||c===0x2189||(c>=0x2190&&c<=0x2211)||(c>=0x2214&&c<=0x2335)||(c>=0x237B&&c<=0x2394)||(c>=0x2396&&c<=0x23FA)||(c>=0x2400&&c<=0x2426)||(c>=0x2440&&c<=0x244A)||(c>=0x2460&&c<=0x2487)||(c>=0x24EA&&c<=0x26AB)||(c>=0x26AD&&c<=0x27FF)||(c>=0x2900&&c<=0x2B73)||(c>=0x2B76&&c<=0x2B95)||(c>=0x2B98&&c<=0x2BB9)||(c>=0x2BBD&&c<=0x2BC8)||(c>=0x2BCA&&c<=0x2BD1)||(c>=0x2CE5&&c<=0x2CEA)||(c>=0x2CF9&&c<=0x2CFF)||(c>=0x2E00&&c<=0x2E42)||(c>=0x2E80&&c<=0x2E99)||(c>=0x2E9B&&c<=0x2EF3)||(c>=0x2F00&&c<=0x2FD5)||(c>=0x2FF0&&c<=0x2FFB)||(c>=0x3000&&c<=0x3004)||(c>=0x3008&&c<=0x3020)||c===0x3030||(c>=0x3036&&c<=0x3037)||(c>=0x303D&&c<=0x303F)||(c>=0x309B&&c<=0x309C)||c===0x30A0||c===0x30FB||(c>=0x31C0&&c<=0x31E3)||(c>=0x321D&&c<=0x321E)||(c>=0x3250&&c<=0x325F)||(c>=0x327C&&c<=0x327E)||(c>=0x32B1&&c<=0x32BF)||(c>=0x32CC&&c<=0x32CF)||(c>=0x3377&&c<=0x337A)||(c>=0x33DE&&c<=0x33DF)||c===0x33FF||(c>=0x4DC0&&c<=0x4DFF)||(c>=0xA490&&c<=0xA4C6)||(c>=0xA60D&&c<=0xA60F)||c===0xA673||(c>=0xA67E&&c<=0xA67F)||(c>=0xA700&&c<=0xA721)||c===0xA788||(c>=0xA828&&c<=0xA82B)||(c>=0xA874&&c<=0xA877)||(c>=0xFD3E&&c<=0xFD3F)||c===0xFDFD||(c>=0xFE10&&c<=0xFE19)||(c>=0xFE30&&c<=0xFE4F)||c===0xFE51||c===0xFE54||(c>=0xFE56&&c<=0xFE5E)||(c>=0xFE60&&c<=0xFE61)||(c>=0xFE64&&c<=0xFE66)||c===0xFE68||c===0xFE6B||(c>=0xFF01&&c<=0xFF02)||(c>=0xFF06&&c<=0xFF0A)||(c>=0xFF1B&&c<=0xFF20)||(c>=0xFF3B&&c<=0xFF40)||(c>=0xFF5B&&c<=0xFF65)||(c>=0xFFE2&&c<=0xFFE4)||(c>=0xFFE8&&c<=0xFFEE)||(c>=0xFFF9&&c<=0xFFFD)||c===0x10101||(c>=0x10140&&c<=0x1018C)||(c>=0x10190&&c<=0x1019B)||c===0x101A0||c===0x1091F||(c>=0x10B39&&c<=0x10B3F)||(c>=0x11052&&c<=0x11065)||(c>=0x1D200&&c<=0x1D241)||c===0x1D245||(c>=0x1D300&&c<=0x1D356)||c===0x1D6DB||c===0x1D715||c===0x1D74F||c===0x1D789||c===0x1D7C3||(c>=0x1EEF0&&c<=0x1EEF1)||(c>=0x1F000&&c<=0x1F02B)||(c>=0x1F030&&c<=0x1F093)||(c>=0x1F0A0&&c<=0x1F0AE)||(c>=0x1F0B1&&c<=0x1F0BF)||(c>=0x1F0C1&&c<=0x1F0CF)||(c>=0x1F0D1&&c<=0x1F0F5)||(c>=0x1F10B&&c<=0x1F10C)||(c>=0x1F16A&&c<=0x1F16B)||(c>=0x1F300&&c<=0x1F32C)||(c>=0x1F330&&c<=0x1F37D)||(c>=0x1F380&&c<=0x1F3CE)||(c>=0x1F3D4&&c<=0x1F3F7)||(c>=0x1F400&&c<=0x1F4FE)||(c>=0x1F500&&c<=0x1F54A)||(c>=0x1F550&&c<=0x1F579)||(c>=0x1F57B&&c<=0x1F5A3)||(c>=0x1F5A5&&c<=0x1F642)||(c>=0x1F645&&c<=0x1F6CF)||(c>=0x1F6E0&&c<=0x1F6EC)||(c>=0x1F6F0&&c<=0x1F6F3)||(c>=0x1F700&&c<=0x1F773)||(c>=0x1F780&&c<=0x1F7D4)||(c>=0x1F800&&c<=0x1F80B)||(c>=0x1F810&&c<=0x1F847)||(c>=0x1F850&&c<=0x1F859)||(c>=0x1F860&&c<=0x1F887)||(c>=0x1F890&&c<=0x1F8AD); }", _
        "function IsWeak(c) { return (c>=0x0&&c<=0x8)||(c>=0xE&&c<=0x1B)||(c>=0x23&&c<=0x25)||(c>=0x2B&&c<=0x3A)||(c>=0x7F&&c<=0x84)||(c>=0x86&&c<=0xA0)||(c>=0xA2&&c<=0xA5)||c===0xAD||(c>=0xB0&&c<=0xB3)||c===0xB9||(c>=0x300&&c<=0x36F)||(c>=0x483&&c<=0x489)||c===0x58F||(c>=0x591&&c<=0x5BD)||c===0x5BF||(c>=0x5C1&&c<=0x5C2)||(c>=0x5C4&&c<=0x5C5)||c===0x5C7||(c>=0x600&&c<=0x605)||(c>=0x609&&c<=0x60A)||c===0x60C||(c>=0x610&&c<=0x61A)||(c>=0x64B&&c<=0x66C)||c===0x670||(c>=0x6D6&&c<=0x6DD)||(c>=0x6DF&&c<=0x6E4)||(c>=0x6E7&&c<=0x6E8)||(c>=0x6EA&&c<=0x6ED)||(c>=0x6F0&&c<=0x6F9)||c===0x711||(c>=0x730&&c<=0x74A)||(c>=0x7A6&&c<=0x7B0)||(c>=0x7EB&&c<=0x7F3)||(c>=0x816&&c<=0x819)||(c>=0x81B&&c<=0x823)||(c>=0x825&&c<=0x827)||(c>=0x829&&c<=0x82D)||(c>=0x859&&c<=0x85B)||(c>=0x8E4&&c<=0x902)||c===0x93A||c===0x93C||(c>=0x941&&c<=0x948)||c===0x94D||(c>=0x951&&c<=0x957)||(c>=0x962&&c<=0x963)||c===0x981||c===0x9BC||(c>=0x9C1&&c<=0x9C4)||c===0x9CD||(c>=0x9E2&&c<=0x9E3)||(c>=0x9F2&&c<=0x9F3)||c===0x9FB||(c>=0xA01&&c<=0xA02)||c===0xA3C||(c>=0xA41&&c<=0xA42)||(c>=0xA47&&c<=0xA48)||(c>=0xA4B&&c<=0xA4D)||c===0xA51||(c>=0xA70&&c<=0xA71)||c===0xA75||(c>=0xA81&&c<=0xA82)||c===0xABC||(c>=0xAC1&&c<=0xAC5)||(c>=0xAC7&&c<=0xAC8)||c===0xACD||(c>=0xAE2&&c<=0xAE3)||c===0xAF1||c===0xB01||c===0xB3C||c===0xB3F||(c>=0xB41&&c<=0xB44)||c===0xB4D||c===0xB56||(c>=0xB62&&c<=0xB63)||c===0xB82||c===0xBC0||c===0xBCD||c===0xBF9||c===0xC00||(c>=0xC3E&&c<=0xC40)||(c>=0xC46&&c<=0xC48)||(c>=0xC4A&&c<=0xC4D)||(c>=0xC55&&c<=0xC56)||(c>=0xC62&&c<=0xC63)||c===0xC81||c===0xCBC||(c>=0xCCC&&c<=0xCCD)||(c>=0xCE2&&c<=0xCE3)||c===0xD01||(c>=0xD41&&c<=0xD44)||c===0xD4D||(c>=0xD62&&c<=0xD63)||c===0xDCA||(c>=0xDD2&&c<=0xDD4)||c===0xDD6||c===0xE31||(c>=0xE34&&c<=0xE3A)||c===0xE3F||(c>=0xE47&&c<=0xE4E)||c===0xEB1||(c>=0xEB4&&c<=0xEB9)||(c>=0xEBB&&c<=0xEBC)||(c>=0xEC8&&c<=0xECD)||(c>=0xF18&&c<=0xF19)||c===0xF35||c===0xF37||c===0xF39||(c>=0xF71&&c<=0xF7E)||(c>=0xF80&&c<=0xF84)||(c>=0xF86&&c<=0xF87)||(c>=0xF8D&&c<=0xF97)||(c>=0xF99&&c<=0xFBC)||c===0xFC6||(c>=0x102D&&c<=0x1030)||(c>=0x1032&&c<=0x1037)||(c>=0x1039&&c<=0x103A)||(c>=0x103D&&c<=0x103E)||(c>=0x1058&&c<=0x1059)||(c>=0x105E&&c<=0x1060)||(c>=0x1071&&c<=0x1074)||c===0x1082||(c>=0x1085&&c<=0x1086)||c===0x108D||c===0x109D||(c>=0x135D&&c<=0x135F)||(c>=0x1712&&c<=0x1714)||(c>=0x1732&&c<=0x1734)||(c>=0x1752&&c<=0x1753)||(c>=0x1772&&c<=0x1773)||(c>=0x17B4&&c<=0x17B5)||(c>=0x17B7&&c<=0x17BD)||c===0x17C6||(c>=0x17C9&&c<=0x17D3)||c===0x17DB||c===0x17DD||(c>=0x180B&&c<=0x180E)||c===0x18A9||(c>=0x1920&&c<=0x1922)||(c>=0x1927&&c<=0x1928)||c===0x1932||(c>=0x1939&&c<=0x193B)||(c>=0x1A17&&c<=0x1A18)||c===0x1A1B||c===0x1A56||(c>=0x1A58&&c<=0x1A5E)||c===0x1A60||c===0x1A62||(c>=0x1A65&&c<=0x1A6C)||(c>=0x1A73&&c<=0x1A7C)||c===0x1A7F||(c>=0x1AB0&&c<=0x1ABE)||(c>=0x1B00&&c<=0x1B03)||c===0x1B34||(c>=0x1B36&&c<=0x1B3A)||c===0x1B3C||c===0x1B42||(c>=0x1B6B&&c<=0x1B73)||(c>=0x1B80&&c<=0x1B81)||(c>=0x1BA2&&c<=0x1BA5)||(c>=0x1BA8&&c<=0x1BA9)||(c>=0x1BAB&&c<=0x1BAD)||c===0x1BE6||(c>=0x1BE8&&c<=0x1BE9)||c===0x1BED||(c>=0x1BEF&&c<=0x1BF1)||(c>=0x1C2C&&c<=0x1C33)||(c>=0x1C36&&c<=0x1C37)||(c>=0x1CD0&&c<=0x1CD2)||(c>=0x1CD4&&c<=0x1CE0)||(c>=0x1CE2&&c<=0x1CE8)||c===0x1CED||c===0x1CF4||(c>=0x1CF8&&c<=0x1CF9)||(c>=0x1DC0&&c<=0x1DF5)||(c>=0x1DFC&&c<=0x1DFF)||(c>=0x200B&&c<=0x200D)||(c>=0x202F&&c<=0x2034)||c===0x2044||(c>=0x2060&&c<=0x2064)||(c>=0x206A&&c<=0x2070)||(c>=0x2074&&c<=0x207B)||(c>=0x2080&&c<=0x208B)||(c>=0x20A0&&c<=0x20BD)||(c>=0x20D0&&c<=0x20F0)||c===0x212E||(c>=0x2212&&c<=0x2213)||(c>=0x2488&&c<=0x249B)||(c>=0x2CEF&&c<=0x2CF1)||c===0x2D7F||(c>=0x2DE0&&c<=0x2DFF)||(c>=0x302A&&c<=0x302D)||(c>=0x3099&&c<=0x309A)||(c>=0xA66F&&c<=0xA672)||(c>=0xA674&&c<=0xA67D)||c===0xA69F||(c>=0xA6F0&&c<=0xA6F1)||c===0xA802||c===0xA806||c===0xA80B||(c>=0xA825&&c<=0xA826)||(c>=0xA838&&c<=0xA839)||c===0xA8C4||(c>=0xA8E0&&c<=0xA8F1)||(c>=0xA926&&c<=0xA92D)||(c>=0xA947&&c<=0xA951)||(c>=0xA980&&c<=0xA982)||c===0xA9B3||(c>=0xA9B6&&c<=0xA9B9)||c===0xA9BC||c===0xA9E5||(c>=0xAA29&&c<=0xAA2E)||(c>=0xAA31&&c<=0xAA32)||(c>=0xAA35&&c<=0xAA36)||c===0xAA43||c===0xAA4C||c===0xAA7C||c===0xAAB0||(c>=0xAAB2&&c<=0xAAB4)||(c>=0xAAB7&&c<=0xAAB8)||(c>=0xAABE&&c<=0xAABF)||c===0xAAC1||(c>=0xAAEC&&c<=0xAAED)||c===0xAAF6||c===0xABE5||c===0xABE8||c===0xABED||c===0xFB1E||c===0xFB29||(c>=0xFE00&&c<=0xFE0F)||(c>=0xFE20&&c<=0xFE2D)||c===0xFE50||c===0xFE52||c===0xFE55||c===0xFE5F||(c>=0xFE62&&c<=0xFE63)||(c>=0xFE69&&c<=0xFE6A)||c===0xFEFF||(c>=0xFF03&&c<=0xFF05)||(c>=0xFF0B&&c<=0xFF1A)||(c>=0xFFE0&&c<=0xFFE1)||(c>=0xFFE5&&c<=0xFFE6)||c===0x101FD||(c>=0x102E0&&c<=0x102FB)||(c>=0x10376&&c<=0x1037A)||(c>=0x10A01&&c<=0x10A03)||(c>=0x10A05&&c<=0x10A06)||(c>=0x10A0C&&c<=0x10A0F)||(c>=0x10A38&&c<=0x10A3A)||c===0x10A3F||(c>=0x10AE5&&c<=0x10AE6)||(c>=0x10E60&&c<=0x10E7E)||c===0x11001||(c>=0x11038&&c<=0x11046)||(c>=0x1107F&&c<=0x11081)||(c>=0x110B3&&c<=0x110B6)||(c>=0x110B9&&c<=0x110BA)||(c>=0x11100&&c<=0x11102)||(c>=0x11127&&c<=0x1112B)||(c>=0x1112D&&c<=0x11134)||c===0x11173||(c>=0x11180&&c<=0x11181)||(c>=0x111B6&&c<=0x111BE)||(c>=0x1122F&&c<=0x11231)||c===0x11234||(c>=0x11236&&c<=0x11237)||c===0x112DF||(c>=0x112E3&&c<=0x112EA)||c===0x11301||c===0x1133C||c===0x11340||(c>=0x11366&&c<=0x1136C)||(c>=0x11370&&c<=0x11374)||(c>=0x114B3&&c<=0x114B8)||c===0x114BA||(c>=0x114BF&&c<=0x114C0)||(c>=0x114C2&&c<=0x114C3)||(c>=0x115B2&&c<=0x115B5)||(c>=0x115BC&&c<=0x115BD)||(c>=0x115BF&&c<=0x115C0)||(c>=0x11633&&c<=0x1163A)||c===0x1163D||(c>=0x1163F&&c<=0x11640)||c===0x116AB||c===0x116AD||(c>=0x116B0&&c<=0x116B5)||c===0x116B7||(c>=0x16AF0&&c<=0x16AF4)||(c>=0x16B30&&c<=0x16B36)||(c>=0x16F8F&&c<=0x16F92)||(c>=0x1BC9D&&c<=0x1BC9E)||(c>=0x1BCA0&&c<=0x1BCA3)||(c>=0x1D167&&c<=0x1D169)||(c>=0x1D173&&c<=0x1D182)||(c>=0x1D185&&c<=0x1D18B)||(c>=0x1D1AA&&c<=0x1D1AD)||(c>=0x1D242&&c<=0x1D244)||(c>=0x1D7CE&&c<=0x1D7FF)||(c>=0x1E8D0&&c<=0x1E8D6)||(c>=0x1F100&&c<=0x1F10A)||c===0xE0001||(c>=0xE0020&&c<=0xE007F)||(c>=0xE0100&&c<=0xE01EF); }", _
        "function IsExplicit(c) { return (c>=0x202A&&c<=0x202E)||(c>=0x2066&&c<=0x2069); }", _
        "function doFixDirection(sVal, direction) { var iCount, sOutVal = direction ? '\u200E' : '\u200F', bInv = false; for (iCount = 0; iCount < sVal.length; iCount++) { sOutVal += (sVal.charCodeAt(iCount) == 0x200E || sVal.charCodeAt(iCount) == 0x200F || IsExplicit(sVal.charCodeAt(iCount))) ? '' : (!bInv && IsNeutral(sVal.charCodeAt(iCount)) ? '' : (direction ? '\u202A' : '\u202B')) + sVal[iCount]; } return sOutVal; }", _
        "function doTransliterateDisplay() { $('#translitvalue').css('direction', $('#direction0').prop('checked') ? 'ltr' : 'rtl'); $('#translitvalue').text($('#scheme0').prop('checked') ? doTransliterate($('#translitedit').val(), $('#direction0').prop('checked'), $('#translitscheme0').prop('checked')) : ($('#scheme1').prop('checked') ? doDiacritics($('#translitedit').val(), $('#diacriticscheme0').prop('checked')) : doFixDirection($('#translitedit').val(), $('#diacriticscheme0').prop('checked')))); }"}
        GetJS.AddRange(TransliterateGenJS)
        GetJS.AddRange(DiacriticJS)
        Return GetJS.ToArray()
    End Function
    Public Shared Function GetSchemeChangeJS() As String()
        Return New String() {"javascript: doSchemeChange();", String.Empty, "function doSchemeChange() { $('#diacriticscheme_').css('display', $('#scheme0').prop('checked') ? 'none' : 'block'); $('#translitscheme_').css('display', $('#scheme0').prop('checked') ? 'block' : 'none'); $('#direction_').css('display', $('#scheme0').prop('checked') ? 'block' : 'none'); }"}
    End Function
    Public Shared Function GetTransliterationSchemes() As Array()
        Dim Count As Integer
        Dim Strings(CachedData.IslamData.TranslitSchemes.Length - 1 + 2) As Array
        Strings(0) = New String() {Utility.LoadResourceString("IslamSource_Off"), "IslamSource_Off"}
        Strings(1) = New String() {Utility.LoadResourceString("IslamSource_ExtendedBuckwalter"), "IslamSource_ExtendedBuckwalter"}
        For Count = 0 To CachedData.IslamData.TranslitSchemes.Length - 1
            Strings(Count + 2) = New String() {Utility.LoadResourceString("IslamSource_" + CachedData.IslamData.TranslitSchemes(Count).Name), "IslamSource_" + CachedData.IslamData.TranslitSchemes(Count).Name}
        Next
        Return Strings
    End Function
    Public Shared Function GetChangeTransliterationJS() As String()
        Dim GetJS As New List(Of String) From {"javascript: changeTransliteration();", String.Empty, Utility.GetLookupStyleSheetJS(), GetArabicSymbolJSArray(), _
        "function changeTransliteration() { var k, child, iSubCount, text; $('span.transliteration').each(function() { $(this).css('display', $('#translitscheme').val() === '0' ? 'none' : 'block'); }); for (k in renderList) { text = ''; for (child in renderList[k]['children']) { for (iSubCount = 0; iSubCount < renderList[k]['children'][child]['arabic'].length; iSubCount++) { if ($('#translitscheme').val() === '1' && renderList[k]['children'][child]['arabic'][iSubCount] !== '' && (renderList[k]['children'][child]['arabic'][iSubCount].length !== 1 || !isStop(findLetterBySymbol(renderList[k]['children'][child]['arabic'][iSubCount].charCodeAt(0)))) && renderList[k]['children'][child]['translit'][iSubCount] !== '') { if (text !== '') text += ' '; text += $('#' + renderList[k]['children'][child]['arabic'][iSubCount]).text(); } else { if (renderList[k]['children'][child]['translit'][iSubCount] !== '') $('#' + renderList[k]['children'][child]['translit'][iSubCount]).text(($('#translitscheme').val() === '0' || renderList[k]['children'][child]['arabic'][iSubCount] === '') ? '' : doTransliterate($('#' + renderList[k]['children'][child]['arabic'][iSubCount]).text(), true, $('#translitscheme').val() === '3')); } } } if ($('#translitscheme').val() === '1') { text = transliterateWithRules(text).split(' '); for (child in renderList[k]['children']) { for (iSubCount = 0; iSubCount < renderList[k]['children'][child]['translit'].length; iSubCount++) { if (renderList[k]['children'][child]['translit'][iSubCount] !== '') $('#' + renderList[k]['children'][child]['translit'][iSubCount]).text(text.shift()); } } } for (iSubCount = 0; iSubCount < renderList[k]['arabic'].length; iSubCount++) { if (renderList[k]['translit'][iSubCount] !== '') $('#' + renderList[k]['translit'][iSubCount]).text(($('#translitscheme').val() === '0' || renderList[k]['arabic'][iSubCount] === '') ? '' : ($('#translitscheme').val() === '1' ? transliterateWithRules($('#' + renderList[k]['arabic'][iSubCount]).text()) : doTransliterate($('#' + renderList[k]['arabic'][iSubCount]).text(), true, $('#translitscheme').val() === '3'))); } } }"}
        GetJS.AddRange(TransliterateGenJS)
        GetJS.AddRange(NumberGenJS)
        GetJS.AddRange(PlainTransliterateGenJS)
        Return GetJS.ToArray()
    End Function
    Public Shared Function GetCategories() As String()
        Dim RetCat As New ArrayList(Array.ConvertAll(CachedData.IslamData.VocabularyCategories, Function(Convert As IslamData.VocabCategory) Utility.LoadResourceString("IslamInfo_" + Convert.Title)))
        RetCat.Add(Utility.LoadResourceString("IslamInfo_Months"))
        RetCat.Add(Utility.LoadResourceString("IslamInfo_DaysOfWeek"))
        RetCat.Add(Utility.LoadResourceString("IslamInfo_PrayerTimes"))
        RetCat.Add(Utility.LoadResourceString("IslamInfo_Prayers"))
        'RetCat.Add(Utility.LoadResourceString("IslamInfo_Fasting"))
        Return CType(RetCat.ToArray(GetType(String)), String())
    End Function
    Public Shared Function GetRenderJS() As String()
        Dim Count As Integer = CInt(HttpContext.Current.Request.QueryString.Get("selection"))
        If Count = -1 Then Count = 0
        Dim Objects As ArrayList = New ArrayList From {New ArrayList, New ArrayList}
        Dim Total As Integer
        If Count = CachedData.IslamData.VocabularyCategories.Length Then
            Total = CachedData.IslamData.Months.Length
        ElseIf Count = CachedData.IslamData.VocabularyCategories.Length + 1 Then
            Total = CachedData.IslamData.DaysOfWeek.Length
        ElseIf Count = CachedData.IslamData.VocabularyCategories.Length + 2 Then
            Total = CachedData.IslamData.PrayerTimes.Length
        ElseIf Count = CachedData.IslamData.VocabularyCategories.Length + 3 Then
            Total = CachedData.IslamData.Prayers.Length
        Else
            Total = CachedData.IslamData.VocabularyCategories(Count).Words.Length
        End If
        For SubCount As Integer = 0 To Total - 1
            CType(Objects(0), ArrayList).Add("render" + CStr(SubCount))
            CType(Objects(1), ArrayList).Add(Utility.MakeJSIndexedObject(New String() {"title", "arabic", "translit", "translate", "children"}, New Array() {New String() {"''", Utility.MakeJSArray(New String() {"render" + CStr(SubCount) + "_0"}), Utility.MakeJSArray(New String() {"render" + CStr(SubCount) + "_1"}), Utility.MakeJSArray(New String() {"render" + CStr(SubCount) + "_2"}), Utility.MakeJSIndexedObject(New String() {}, New Array() {New String() {}}, True)}}, True))
        Next
        Return New String() {String.Empty, String.Empty, "var renderList = " + Utility.MakeJSIndexedObject(CType(CType(Objects(0), ArrayList).ToArray(GetType(String)), String()), New Array() {CType(CType(Objects(1), ArrayList).ToArray(GetType(String)), String())}, True) + ";"}
    End Function
    Public Shared Function DisplayTranslation(ByVal Item As PageLoader.TextItem) As Array()
        Dim Count As Integer = CInt(HttpContext.Current.Request.QueryString.Get("selection"))
        If Count = -1 Then Count = 0
        Dim SchemeType As Arabic.TranslitScheme = If(CInt(HttpContext.Current.Request.QueryString.Get("translitscheme")) >= 2, 2 + CInt(HttpContext.Current.Request.QueryString.Get("translitscheme")) Mod 2, CInt(HttpContext.Current.Request.QueryString.Get("translitscheme")))
        Dim Scheme As String = If(CInt(HttpContext.Current.Request.QueryString.Get("translitscheme")) >= 2, CachedData.IslamData.TranslitSchemes((CInt(HttpContext.Current.Request.QueryString.Get("translitscheme")) - 2) \ 2).Name, String.Empty)
        If Count = CachedData.IslamData.VocabularyCategories.Length Then
            Return DoDisplayTranslation(CachedData.IslamData.Months, SchemeType, Scheme)
        ElseIf Count = CachedData.IslamData.VocabularyCategories.Length + 1 Then
            Return DoDisplayTranslation(CachedData.IslamData.DaysOfWeek, SchemeType, Scheme)
        ElseIf Count = CachedData.IslamData.VocabularyCategories.Length + 2 Then
            Return DoDisplayTranslation(CachedData.IslamData.PrayerTimes, SchemeType, Scheme)
        ElseIf Count = CachedData.IslamData.VocabularyCategories.Length + 3 Then
            Return DoDisplayTranslation(CachedData.IslamData.Prayers, SchemeType, Scheme)
        Else
            Return DoDisplayTranslation(CachedData.IslamData.VocabularyCategories(Count), SchemeType, Scheme)
        End If
    End Function
    Public Shared Function DoDisplayTranslation(Category As Object, SchemeType As Arabic.TranslitScheme, Scheme As String) As Array()
        Dim Output As New ArrayList
        Output.Add(GetRenderJS())
        If TypeOf Category Is IslamData.PrayerTime Then
            Output.Add(New String() {"arabic", "transliteration", String.Empty, String.Empty, String.Empty, String.Empty})
            Output.Add(New String() {Utility.LoadResourceString("IslamInfo_Arabic"), Utility.LoadResourceString("IslamInfo_Transliteration"), Utility.LoadResourceString("IslamInfo_Translation"), Utility.LoadResourceString("IslamInfo_Before"), Utility.LoadResourceString("IslamInfo_PrescribedTime"), Utility.LoadResourceString("IslamInfo_After")})
        ElseIf TypeOf Category Is IslamData.PrayerType Then
            Output.Add(New String() {"arabic", "transliteration", String.Empty, String.Empty, String.Empty})
            Output.Add(New String() {Utility.LoadResourceString("IslamInfo_Arabic"), Utility.LoadResourceString("IslamInfo_Transliteration"), Utility.LoadResourceString("IslamInfo_Translation"), Utility.LoadResourceString("IslamInfo_Classification"), Utility.LoadResourceString("IslamInfo_PrayerUnits")})
        Else
            Output.Add(New String() {"arabic", "transliteration", String.Empty})
            Output.Add(New String() {Utility.LoadResourceString("IslamInfo_Arabic"), Utility.LoadResourceString("IslamInfo_Transliteration"), Utility.LoadResourceString("IslamInfo_Translation")})
        End If
        If TypeOf Category Is IslamData.Month Then
            For SubCount As Integer = 0 To CType(Category, IslamData.Month()).Length - 1
                Output.Add(New String() {Arabic.RightToLeftMark + TransliterateFromBuckwalter(CType(Category, IslamData.Month())(SubCount).Name), TransliterateToScheme(TransliterateFromBuckwalter(CType(Category, IslamData.Month())(SubCount).Name), SchemeType, Scheme).Trim(), Utility.LoadResourceString("IslamInfo_" + CType(Category, IslamData.Month())(SubCount).TranslationID)})
            Next
        ElseIf TypeOf Category Is IslamData.DayOfWeek Then
            For SubCount As Integer = 0 To CType(Category, IslamData.DayOfWeek()).Length - 1
                Output.Add(New String() {Arabic.RightToLeftMark + TransliterateFromBuckwalter(CType(Category, IslamData.DayOfWeek())(SubCount).Name), TransliterateToScheme(TransliterateFromBuckwalter(CType(Category, IslamData.DayOfWeek())(SubCount).Name), SchemeType, Scheme).Trim(), Utility.LoadResourceString("IslamInfo_" + CType(Category, IslamData.DayOfWeek())(SubCount).TranslationID)})
            Next
        ElseIf TypeOf Category Is IslamData.PrayerTime Then
            Dim Table As New Hashtable
            Array.ForEach(CachedData.IslamData.Prayers, Sub(Convert As IslamData.PrayerType) Array.ForEach(Convert.PrayerUnits.Split(","c), Sub(Part As String) If Part.Contains("="c) Then If Table.ContainsKey(Part.Substring(0, Part.IndexOf("="c))) Then Table.Item(Part.Substring(0, Part.IndexOf("="c))) = CStr(Table.Item(Part.Substring(0, Part.IndexOf("="c)))) + vbCrLf + Part.Substring(Part.IndexOf("="c) + 1).Replace("|"c, " or ") + " " + CStr(IIf(Utility.LoadResourceString("IslamInfo_" + Convert.TranslationID) <> "Prescribed time", Utility.LoadResourceString("IslamInfo_" + Convert.TranslationID) + " - ", String.Empty)) + Convert.Classification Else Table.Add(Part.Substring(0, Part.IndexOf("="c)), Part.Substring(Part.IndexOf("="c) + 1).Replace("|"c, " or ") + " " + Convert.Classification)))
            For SubCount As Integer = 0 To CType(Category, IslamData.PrayerTime()).Length - 1
                Output.Add(New String() {Arabic.RightToLeftMark + TransliterateFromBuckwalter(CType(Category, IslamData.PrayerTime())(SubCount).Name), TransliterateToScheme(TransliterateFromBuckwalter(CType(Category, IslamData.PrayerTime())(SubCount).Name), SchemeType, Scheme).Trim(), Utility.LoadResourceString("IslamInfo_" + CType(Category, IslamData.PrayerTime())(SubCount).TranslationID), CStr(IIf(Table.ContainsKey("-"c + Utility.LoadResourceString("IslamInfo_" + CType(Category, IslamData.PrayerTime())(SubCount).TranslationID)), Table.Item("-"c + Utility.LoadResourceString("IslamInfo_" + CType(Category, IslamData.PrayerTime())(SubCount).TranslationID)), String.Empty)), CStr(IIf(Table.ContainsKey(Utility.LoadResourceString("IslamInfo_" + CType(Category, IslamData.PrayerTime())(SubCount).TranslationID)), Table.Item(Utility.LoadResourceString("IslamInfo_" + CType(Category, IslamData.PrayerTime())(SubCount).TranslationID)), String.Empty)), CStr(IIf(Table.ContainsKey("+"c + Utility.LoadResourceString("IslamInfo_" + CType(Category, IslamData.PrayerTime())(SubCount).TranslationID)), Table.Item("+"c + Utility.LoadResourceString("IslamInfo_" + CType(Category, IslamData.PrayerTime())(SubCount).TranslationID)), String.Empty))})
            Next
        ElseIf TypeOf Category Is IslamData.PrayerType Then
            For SubCount As Integer = 0 To CType(Category, IslamData.PrayerType()).Length - 1
                Output.Add(New String() {Arabic.RightToLeftMark + TransliterateFromBuckwalter(CType(Category, IslamData.PrayerType())(SubCount).Name), TransliterateToScheme(TransliterateFromBuckwalter(CType(Category, IslamData.PrayerType())(SubCount).Name), SchemeType, Scheme).Trim(), Utility.LoadResourceString("IslamInfo_" + CType(Category, IslamData.PrayerType())(SubCount).TranslationID), CType(Category, IslamData.PrayerType())(SubCount).Classification, CType(Category, IslamData.PrayerType())(SubCount).PrayerUnits})
            Next
        Else
            For SubCount As Integer = 0 To CType(Category, IslamData.VocabCategory).Words.Length - 1
                Output.Add(New String() {Arabic.RightToLeftMark + TransliterateFromBuckwalter(CType(Category, IslamData.VocabCategory).Words(SubCount).Text), TransliterateToScheme(TransliterateFromBuckwalter(CType(Category, IslamData.VocabCategory).Words(SubCount).Text), SchemeType, Scheme).Trim(), Utility.LoadResourceString("IslamInfo_" + CType(Category, IslamData.VocabCategory).Words(SubCount).TranslationID)})
            Next
        End If
        Return DirectCast(Output.ToArray(GetType(Array)), Array())
    End Function
    Public Shared Function DisplayDict(ByVal Item As PageLoader.TextItem) As Array()
        Dim Lines As String() = IO.File.ReadAllLines(Utility.GetFilePath("metadata\HansWeir.txt"))
        Dim Count As Integer
        Dim Words As New List(Of String())
        Words.Add(New String() {})
        Words.Add(New String() {"arabic", String.Empty})
        Words.Add(New String() {Utility.LoadResourceString("IslamInfo_Arabic"), Utility.LoadResourceString("IslamInfo_Meaning")})
        For Count = 0 To Lines.Length - 1
            '"ā", "au", "ai", "ē", "ī", "ō", "ū", "t", "d", "ḳ", "š", "ṣ", "ḍ", "ṯ", "ẓ", "‘", "’"
            '".", "│"
            Dim Matches As System.Text.RegularExpressions.MatchCollection = System.Text.RegularExpressions.Regex.Matches(Lines(Count), "(\p{IsArabic}+\s*―\s*\p{IsArabic}+)([1-6]?) ([^\s,]+\s*―\s*[^\s,]+)(?:,?)")
            For MatchCount = 0 To Matches.Count - 1
                Dim Val As String = String.Empty
                Dim Len As Integer = Matches(MatchCount).Groups(3).Value.Length - 1
                For SubCount = Matches(MatchCount).Groups(1).Value.Length - 1 To 0 Step -1
                    '()
                    '" see "
                    '"pl. -āt"
                    If (Matches(MatchCount).Groups(3).Value(Len) = "a") Then
                        Val = ArabicFatha + Val
                    ElseIf (Matches(MatchCount).Groups(3).Value(Len) = "i") Then
                        Val = ArabicKasra + Val
                    ElseIf (Matches(MatchCount).Groups(3).Value(Len) = "u") Then
                        Val = ArabicDamma + Val
                    ElseIf (Matches(MatchCount).Groups(3).Value(Len) = "ā" Or Len <> 0 AndAlso ((Matches(MatchCount).Groups(3).Value(Len) = "i" Or Matches(MatchCount).Groups(3).Value(SubCount) = "u") And Matches(MatchCount).Groups(3).Value(SubCount - 1) = "a")) Then
                        Val = ArabicFatha + Val
                    ElseIf (Matches(MatchCount).Groups(3).Value(Len) = "ī" Or Matches(MatchCount).Groups(3).Value(Len) = "ē") Then
                        Val = ArabicKasra + Val
                    ElseIf (Matches(MatchCount).Groups(3).Value(Len) = "ū" Or Matches(MatchCount).Groups(3).Value(Len) = "ō") Then
                        Val = ArabicDamma + Val
                    ElseIf Len <> 0 AndAlso Matches(MatchCount).Groups(3).Value(Len) = Matches(MatchCount).Groups(3).Value(Len - 1) Then
                        Val = ArabicShadda + Val
                    ElseIf Matches(MatchCount).Groups(3).Value(Len) = "." Then
                    Else
                        Val = Matches(MatchCount).Groups(1).Value(SubCount) + Val
                    End If
                Next
                Words.Add(New String() {Val})
            Next
        Next
        Return Words.ToArray()
    End Function
    Public Shared Function CheckShapingOrder(Index As Integer, UName As String) As Boolean
        Return UName.EndsWith("Isolated Form") And Index = 0 Or UName.EndsWith("Final Form") And Index = 1 Or _
            UName.EndsWith("Initial Form") And Index = 2 Or UName.EndsWith("Medial Form") And Index = 3
    End Function
    Public Shared Function DisplayCombo(ByVal Item As PageLoader.TextItem) As Array()
        Dim Count As Integer
        Dim Output(CachedData.IslamData.ArabicCombos.Length + 2) As Array
        Output(0) = New String() {}
        Output(1) = New String() {"arabic", "arabic", String.Empty, String.Empty}
        Output(2) = New String() {Utility.LoadResourceString("IslamInfo_LetterName"), Utility.LoadResourceString("IslamInfo_Arabic"), Utility.LoadResourceString("IslamSource_ExtendedBuckwalter"), Utility.LoadResourceString("IslamInfo_Shaping")}
        Dim Combos As IslamData.ArabicCombo() = CachedData.IslamData.ArabicCombos
        'Array.Sort(Combos, Function(Key As IslamData.ArabicCombo, NextKey As IslamData.ArabicCombo) Key.SymbolName.CompareTo(NextKey.SymbolName))
        For Count = 0 To CachedData.IslamData.ArabicCombos.Length - 1
            Output(Count + 3) = New String() {String.Join(" ", Array.ConvertAll(TransliterateFromBuckwalter(CachedData.IslamData.ArabicCombos(Count).SymbolName).ToCharArray(), Function(Ch As Char) TransliterateFromBuckwalter(CachedData.IslamData.ArabicLetters(FindLetterBySymbol(Ch)).SymbolName))), _
                                       TransliterateFromBuckwalter(CachedData.IslamData.ArabicCombos(Count).SymbolName), _
                                       CachedData.IslamData.ArabicCombos(Count).SymbolName,
                                       String.Join(", ", Array.ConvertAll(CachedData.IslamData.ArabicCombos(Count).Shaping, Function(Shape As Char) If(Shape = ChrW(0), String.Empty, Shape + " " + CStr(AscW(Shape)) + " " + If(CheckShapingOrder(Array.IndexOf(CachedData.IslamData.ArabicCombos(Count).Shaping, Shape), GetUnicodeName(Shape)), String.Empty, "!!!") + GetUnicodeName(Shape))))}
        Next
        Return Output
    End Function
    Public Shared Function LetterDisplay(Symbols() As IslamData.ArabicSymbol) As Array()
        Dim Count As Integer
        Dim Output(CachedData.IslamData.ArabicLetters.Length + 2) As Array
        'Dim oFont As New Font(DefaultValue(HttpContext.Current.Request.QueryString.Get("fontcustom"), "Arial"), 13)
        'CheckIfCharInFont(CachedData.IslamData.ArabicLetters(Count).Symbol, oFont)
        Output(0) = New String() {}
        Output(1) = New String() {"arabic", String.Empty, "arabic", String.Empty, String.Empty, String.Empty, String.Empty, String.Empty, String.Empty}
        Output(2) = New String() {Utility.LoadResourceString("IslamInfo_LetterName"), Utility.LoadResourceString("IslamInfo_UnicodeName"), Utility.LoadResourceString("IslamInfo_Arabic"), Utility.LoadResourceString("IslamInfo_UnicodeValue"), Utility.LoadResourceString("IslamSource_ExtendedBuckwalter"), Utility.LoadResourceString("IslamInfo_Terminating"), Utility.LoadResourceString("IslamInfo_Connecting"), Utility.LoadResourceString("IslamInfo_Assimilate"), Utility.LoadResourceString("IslamInfo_Shaping")}
        For Count = 0 To Symbols.Length - 1
            Output(Count + 3) = New String() {Arabic.RightToLeftMark + TransliterateFromBuckwalter(Symbols(Count).SymbolName), _
                                              GetUnicodeName(Symbols(Count).Symbol), _
                                       CStr(Symbols(Count).Symbol), _
                                       CStr(AscW(Symbols(Count).Symbol)), _
                                       CStr(IIf(Symbols(Count).ExtendedBuckwalterLetter = ChrW(0), String.Empty, Symbols(Count).ExtendedBuckwalterLetter)), _
                                       CStr(Symbols(Count).Terminating), _
                                       CStr(Symbols(Count).Connecting), _
                                       CStr(Symbols(Count).Assimilate),
                                       If(Symbols(Count).Shaping = Nothing, String.Empty, String.Join(", ", Array.ConvertAll(Symbols(Count).Shaping, Function(Shape As Char) If(Shape = ChrW(0), String.Empty, Shape + " " + CStr(AscW(Shape)) + " " + If(CheckShapingOrder(Array.IndexOf(Symbols(Count).Shaping, Shape), GetUnicodeName(Shape)), String.Empty, "!!!") + GetUnicodeName(Shape)))))}
        Next
        Return Output
    End Function
    Public Shared Function DisplayAll(ByVal Item As PageLoader.TextItem) As Array()
        Return LetterDisplay(CachedData.IslamData.ArabicLetters)
    End Function
    Public Shared Function DisplayTranslitSchemes(ByVal Item As PageLoader.TextItem) As Array()
        Dim Count As Integer
        Dim Output(CachedData.IslamData.ArabicLetters.Length + 2) As Array
        'Dim oFont As New Font(DefaultValue(HttpContext.Current.Request.QueryString.Get("fontcustom"), "Arial"), 13)
        'CheckIfCharInFont(CachedData.IslamData.ArabicLetters(Count).Symbol, oFont)
        Output(0) = New String() {}
        Output(1) = New String() {"arabic", String.Empty, "arabic", String.Empty}
        Array.Resize(Of String)(Output(1), 4 + CachedData.IslamData.TranslitSchemes.Length)
        Output(2) = New String() {Utility.LoadResourceString("IslamInfo_LetterName"), Utility.LoadResourceString("IslamInfo_UnicodeName"), Utility.LoadResourceString("IslamInfo_Arabic"), Utility.LoadResourceString("IslamSource_ExtendedBuckwalter")}
        Array.Resize(Of String)(Output(2), 4 + CachedData.IslamData.TranslitSchemes.Length)
        For SchemeCount = 0 To CachedData.IslamData.TranslitSchemes.Length - 1
            Output(1)(4 + SchemeCount) = String.Empty
            Output(2)(4 + SchemeCount) = Utility.LoadResourceString("IslamSource_" + CachedData.IslamData.TranslitSchemes(SchemeCount).Name)
        Next
        For Count = 0 To CachedData.IslamData.ArabicLetters.Length - 1
            Output(Count + 3) = New String() {Arabic.RightToLeftMark + TransliterateFromBuckwalter(CachedData.IslamData.ArabicLetters(Count).SymbolName), _
                                              GetUnicodeName(CachedData.IslamData.ArabicLetters(Count).Symbol), _
                                       CStr(CachedData.IslamData.ArabicLetters(Count).Symbol), _
                                       CStr(IIf(CachedData.IslamData.ArabicLetters(Count).ExtendedBuckwalterLetter = ChrW(0), String.Empty, CachedData.IslamData.ArabicLetters(Count).ExtendedBuckwalterLetter))}
            Array.Resize(Of String)(Output(Count + 3), 4 + CachedData.IslamData.TranslitSchemes.Length)
            For SchemeCount = 0 To CachedData.IslamData.TranslitSchemes.Length - 1
                Output(Count + 3)(4 + SchemeCount) = GetSchemeValueFromSymbol(CachedData.IslamData.ArabicLetters(Count), CachedData.IslamData.TranslitSchemes(SchemeCount).Name)
            Next
        Next
        Return Output
    End Function
    Public Shared Function DisplayParticle(Category As IslamData.GrammarCategory) As Array()
        Dim Count As Integer
        Dim Output(2 + Category.Words.Length) As Array
        Dim Build As New Generic.Dictionary(Of String, Generic.Dictionary(Of String, String))
        Output(0) = New String() {}
        Output(1) = New String() {"arabic", String.Empty, String.Empty}
        Output(2) = New String() {"Particle", "Translation", "Grammar Feature"}
        For Count = 0 To Category.Words.Length - 1
            Output(3 + Count) = New String() {TransliterateFromBuckwalter(Category.Words(Count).Text), Category.Words(Count).TranslationID, Utility.DefaultValue(Category.Words(Count).Grammar, String.Empty)}
        Next
        Return Output
    End Function
    Public Shared Function DisplayPronoun(Category As IslamData.GrammarCategory, Personal As Boolean) As Array()
        Dim Count As Integer
        Dim Output(2 + If(Personal, 6, 2)) As Array
        Dim Build As New Generic.Dictionary(Of String, Generic.Dictionary(Of String, String))
        Output(0) = New String() {}
        Output(1) = New String() {"arabic", "arabic", "arabic", String.Empty}
        Output(2) = New String() {"Plural", "Dual", "Singular", String.Empty}
        For Count = 0 To Category.Words.Length - 1
            Array.ForEach(Category.Words(Count).Grammar.Split(","c)(0).Split("|"c),
                          Sub(Str As String)
                              Dim Key As String = Str.Chars(0)
                              If Personal Then '"123".Contains(Str.Chars(0))
                                  Key += Str.Chars(1)
                              End If
                              If Not Build.ContainsKey(Key) Then
                                  Build.Add(Key, New Generic.Dictionary(Of String, String))
                              End If
                              If Build.Item(Key).ContainsKey(Str.Chars(If(Personal, 2, 1))) Then
                                  Build.Item(Key).Item(Str.Chars(If(Personal, 2, 1))) += " " + TransliterateFromBuckwalter(Category.Words(Count).Text)
                              Else
                                  Build.Item(Key).Add(Str.Chars(If(Personal, 2, 1)), TransliterateFromBuckwalter(Category.Words(Count).Text))
                              End If
                          End Sub)
        Next
        If Personal Then
            Output(3) = New String() {Build("3m")("p"), Build("3m")("d"), Build("3m")("s"), "Third Person Masculine"}
            Output(4) = New String() {Build("3f")("p"), Build("3f")("d"), Build("3f")("s"), "Third Person Feminine"}
            Output(5) = New String() {Build("2m")("p"), Build("2m")("d"), Build("2m")("s"), "Second Person Masculine"}
            Output(6) = New String() {Build("2f")("p"), Build("2f")("d"), Build("2f")("s"), "Second Person Feminine"}
            Output(7) = New String() {Build("1m")("p"), Build("1m")("d"), Build("1m")("s"), "First Person Masculine"}
            Output(8) = New String() {Build("1f")("p"), Build("1f")("d"), Build("1f")("s"), "First Person Feminine"}
        Else
            Output(3) = New String() {Build("m")("p"), Build("m")("d"), Build("m")("s"), "Masculine"}
            Output(4) = New String() {Build("f")("p"), Build("f")("d"), Build("f")("s"), "Feminine"}
        End If
        Return Output
    End Function
    Public Shared Function DisplayProximals(ByVal Item As PageLoader.TextItem) As Array()
        Return DisplayPronoun(CachedData.IslamData.GrammarCategories(1), False)
    End Function
    Public Shared Function DisplayDistals(ByVal Item As PageLoader.TextItem) As Array()
        Return DisplayPronoun(CachedData.IslamData.GrammarCategories(2), False)
    End Function
    Public Shared Function DisplayRelatives(ByVal Item As PageLoader.TextItem) As Array()
        Return DisplayPronoun(CachedData.IslamData.GrammarCategories(4), False)
    End Function
    Public Shared Function DisplayPersonals(ByVal Item As PageLoader.TextItem) As Array()
        Return DisplayPronoun(CachedData.IslamData.GrammarCategories(5), True)
    End Function
    Public Shared Function DisplayDeterminerPersonals(ByVal Item As PageLoader.TextItem) As Array()
        Return DisplayPronoun(CachedData.IslamData.GrammarCategories(6), True)
    End Function
    Public Shared Function DisplayPastVerbsFamilyI(ByVal Item As PageLoader.TextItem) As Array()
        Return DisplayPronoun(CachedData.IslamData.GrammarCategories(7), True)
    End Function
    Public Shared Function DisplayPresentVerbsFamilyI(ByVal Item As PageLoader.TextItem) As Array()
        Return DisplayPronoun(CachedData.IslamData.GrammarCategories(8), True)
    End Function
    Public Shared Function DisplayCommandVerbsFamilyI(ByVal Item As PageLoader.TextItem) As Array()
        Return DisplayPronoun(CachedData.IslamData.GrammarCategories(9), False)
    End Function
    Public Shared Function DisplayResponseParticles(ByVal Item As PageLoader.TextItem) As Array()
        Return DisplayParticle(CachedData.IslamData.GrammarCategories(10))
    End Function
    Public Shared Function DisplayInterogativeParticles(ByVal Item As PageLoader.TextItem) As Array()
        Return DisplayParticle(CachedData.IslamData.GrammarCategories(11))
    End Function
    Public Shared Function DisplayLocationParticles(ByVal Item As PageLoader.TextItem) As Array()
        Return DisplayParticle(CachedData.IslamData.GrammarCategories(12))
    End Function
    Public Shared Function DisplayTimeParticles(ByVal Item As PageLoader.TextItem) As Array()
        Return DisplayParticle(CachedData.IslamData.GrammarCategories(13))
    End Function
    Public Shared Function DisplayPrepositions(ByVal Item As PageLoader.TextItem) As Array()
        Return DisplayParticle(CachedData.IslamData.GrammarCategories(14))
    End Function
    Public Shared Function DisplayParticles(ByVal Item As PageLoader.TextItem) As Array()
        Return DisplayParticle(CachedData.IslamData.GrammarCategories(15))
    End Function
    Public Shared Function DisplayOtherParticles(ByVal Item As PageLoader.TextItem) As Array()
        Return DisplayParticle(CachedData.IslamData.GrammarCategories(16))
    End Function
    Public Shared Function NounDisplay(Category As IslamData.GrammarCategory) As Array()
        Dim Count As Integer
        Dim Output(2 + Category.Words.Length) As Array
        Dim Build As New Generic.Dictionary(Of String, Generic.Dictionary(Of String, String))
        Output(0) = New String() {}
        Output(1) = New String() {"arabic", "arabic", "arabic", String.Empty}
        Output(2) = New String() {"Noun", "Dual", "Plural", "Singular Translation"}
        For Count = 0 To Category.Words.Length - 1
            Output(3 + Count) = New String() {TransliterateFromBuckwalter(Category.Words(Count).Text), String.Empty, String.Empty, Category.Words(Count).TranslationID}
        Next
        Return Output
    End Function
    Public Shared Function DisplayNouns(ByVal Item As PageLoader.TextItem) As Array()
        Return NounDisplay(CachedData.IslamData.GrammarCategories(17))
    End Function
    Public Shared Function VerbDisplay(Category As IslamData.GrammarCategory) As Array()
        Dim Count As Integer
        Dim Output(2 + Category.Words.Length) As Array
        Dim Build As New Generic.Dictionary(Of String, Generic.Dictionary(Of String, String))
        Output(0) = New String() {}
        Output(1) = New String() {"arabic", "arabic", "arabic", "arabic", "arabic", "arabic", "arabic", "arabic", "arabic"}
        Output(2) = New String() {"Past Root", "Present Root", "Command Root", "Forbidding Root", "Passive Past Root", "Passive Present Root", "Verbal Doer", "Passive Noun", "Particles"}
        For Count = 0 To Category.Words.Length - 1
            Dim Grammar As String
            Dim Text As String
            Dim Present As String
            Dim Command As String
            Dim Forbidding As String
            Dim PassivePast As String
            Dim PassivePresent As String
            Dim VerbalDoer As String
            Dim PassiveNoun As String
            If (Not Category.Words(Count).Grammar Is Nothing AndAlso Category.Words(Count).Grammar.StartsWith("form=")) Then
                Text = TransliterateFromBuckwalter(Category.Words(Count).Grammar.Substring(5).Split(","c)(0).Replace("f", Category.Words(Count).Text.Chars(0)).Replace("E", Category.Words(Count).Text.Chars(1)).Replace("l", Category.Words(Count).Text.Chars(2)))
                Present = TransliterateFromBuckwalter(Category.Words(Count).Grammar.Substring(5).Split(","c)(1).Replace("f", Category.Words(Count).Text.Chars(0)).Replace("E", Category.Words(Count).Text.Chars(1)).Replace("l", Category.Words(Count).Text.Chars(2)))
                Command = TransliterateFromBuckwalter("{foE$1lo".Replace("f", Category.Words(Count).Text.Chars(0)).Replace("E", Category.Words(Count).Text.Chars(1)).Replace("l", Category.Words(Count).Text.Chars(2)).Replace("$1", Category.Words(Count).Grammar.Substring(5).Split(","c)(1).Chars(5)))
                Forbidding = TransliterateFromBuckwalter("laA tafoE$1lo".Replace("f", Category.Words(Count).Text.Chars(0)).Replace("E", Category.Words(Count).Text.Chars(1)).Replace("l", Category.Words(Count).Text.Chars(2)).Replace("$1", Category.Words(Count).Grammar.Substring(5).Split(","c)(1).Chars(5)))
                PassivePast = TransliterateFromBuckwalter("fuEila".Replace("f", Category.Words(Count).Text.Chars(0)).Replace("E", Category.Words(Count).Text.Chars(1)).Replace("l", Category.Words(Count).Text.Chars(2)))
                PassivePresent = TransliterateFromBuckwalter("yufoEalu".Replace("f", Category.Words(Count).Text.Chars(0)).Replace("E", Category.Words(Count).Text.Chars(1)).Replace("l", Category.Words(Count).Text.Chars(2)))
                VerbalDoer = TransliterateFromBuckwalter("faAEilN".Replace("f", Category.Words(Count).Text.Chars(0)).Replace("E", Category.Words(Count).Text.Chars(1)).Replace("l", Category.Words(Count).Text.Chars(2)))
                PassiveNoun = TransliterateFromBuckwalter("mafoEuwlN".Replace("f", Category.Words(Count).Text.Chars(0)).Replace("E", Category.Words(Count).Text.Chars(1)).Replace("l", Category.Words(Count).Text.Chars(2)))
                '"faEiylN" passive noun
                '"mafoEilN" time
                '"mafoEalN" place
                '"faEiylN" derived adjective "faEalN" "faEulaAnu" "&gt;ufoEalN" "faEolN" "fuEaAlN" "faEuwlN"
                '"&gt;afoEalu" "fuEolaY" comparative derived noun
                '"faE~aAlN" "mifoEaAlN" "faEuwlN" "fuE~uwl" "faEilN" "fiE~iylN" "fuEalapN" intensive derived noun
                '"mifoEalapN" "mifoEalN" "mifoEaAlN" instrument of action
                Grammar = String.Empty
            Else
                Text = TransliterateFromBuckwalter(Category.Words(Count).Text)
                Present = String.Empty
                Command = String.Empty
                Forbidding = String.Empty
                PassivePast = String.Empty
                PassivePresent = String.Empty
                VerbalDoer = String.Empty
                PassiveNoun = String.Empty
                Grammar = Utility.DefaultValue(Category.Words(Count).Grammar, String.Empty)
            End If
            Output(3 + Count) = New String() {Text, Present, Command, Forbidding, PassivePast, PassivePresent, VerbalDoer, PassiveNoun, Grammar}
        Next
        Return Output
    End Function
    Public Shared Function DisplayVerbs(ByVal Item As PageLoader.TextItem) As Array()
        Return VerbDisplay(CachedData.IslamData.GrammarCategories(18))
    End Function
End Class
Public Class RenderArray
    Enum RenderTypes
        eHeaderLeft
        eHeaderCenter
        eHeaderRight
        eText
        eRanking
    End Enum
    Enum RenderDisplayClass
        eNested
        eArabic
        eTransliteration
        eLTR
        eRTL
        eRanking
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
        Return "var renderList = " + Utility.MakeJSIndexedObject(CType(CType(Objects(0), ArrayList).ToArray(GetType(String)), String()), New Array() {CType(CType(Objects(1), ArrayList).ToArray(GetType(String)), String())}, True) + ";"
    End Function
    Public Shared Function GetInitJSItems(Items As Collections.Generic.List(Of RenderItem), Title As String, NestPrefix As String) As Object()
        Dim Count As Integer
        Dim Index As Integer
        Dim Objects As ArrayList = New ArrayList From {New ArrayList, New ArrayList}
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
                ElseIf Items(Count).TextItems(Index).DisplayClass = RenderDisplayClass.eRanking Then
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
        Return New String() {String.Empty, String.Empty, GetInitJS(Items), GetCopyClipboardJS(), GetSetClipboardJS(), GetStarRatingJS()}
    End Function
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
                writer.WriteBeginTag(CStr(IIf(Items(Count).TextItems(Index).DisplayClass = RenderDisplayClass.eNested Or Items(Count).TextItems(Index).DisplayClass = RenderDisplayClass.eRanking, "div", "span")))
                If Items(Count).Type = RenderTypes.eHeaderCenter Or (Items(Count).Type = RenderTypes.eText And (Count - Base) Mod 2 = 1) Then writer.WriteAttribute("style", "background-color: 0xD0D0D0;")
                If Items(Count).TextItems(Index).DisplayClass = RenderDisplayClass.eNested Then
                    writer.Write(HtmlTextWriter.TagRightChar)
                    DoRender(writer, TabCount, CType(Items(Count).TextItems(Index).Text, Collections.Generic.List(Of RenderItem)), CStr(Count))
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
                        writer.WriteAttribute("style", "color: " + System.Drawing.ColorTranslator.ToHtml(Items(Count).TextItems(Index).Clr) + "; display: " + CStr(IIf(CInt(HttpContext.Current.Request.QueryString.Get("translitscheme")) <> Arabic.TranslitScheme.None, "block", "none")) + ";")
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
                writer.WriteEndTag(CStr(IIf(Items(Count).TextItems(Index).DisplayClass = RenderDisplayClass.eNested Or Items(Count).TextItems(Index).DisplayClass = RenderDisplayClass.eRanking, "div", "span")))
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
Public Class ArabicFont
    'Web.Config requires: configuration -> system.webServer -> staticContent -> <mimeMap fileExtension=".otf" mimeType="application/octet-stream" />
    'Web.Config requires for cross site scripting: configuration -> system.WebServer -> httpProtocol -> customHeaders -> <add name="Access-Control-Allow-Origin" value="*" />
    Public Shared Function GetFontList() As Array()
        Dim Count As Integer
        Dim Strings(CachedData.IslamData.ArabicFonts.Length - 1) As Array
        For Count = 0 To CachedData.IslamData.ArabicFonts.Length - 1
            Strings(Count) = New String() {Utility.LoadResourceString("IslamInfo_" + CachedData.IslamData.ArabicFonts(Count).Name), CachedData.IslamData.ArabicFonts(Count).ID}
        Next
        Return Strings
    End Function
    Public Shared Function GetArabicFontListJS() As String
        Return "var fontList = " + _
        Utility.MakeJSIndexedObject(Array.ConvertAll(CachedData.IslamData.ArabicFonts, Function(Convert As IslamData.ArabicFontList) Convert.ID), New Array() {Array.ConvertAll(CachedData.IslamData.ArabicFonts, Function(Convert As IslamData.ArabicFontList) Utility.MakeJSIndexedObject(New String() {"family", "embed", "file", "scale"}, New Array() {New String() {Convert.Family, Convert.EmbedName, Convert.FileName, CStr(Convert.Scale)}}, False))}, True) + _
        ";var fontPrefs = " + Utility.MakeJSIndexedObject(Array.ConvertAll(CachedData.IslamData.ScriptFonts, Function(Convert As IslamData.ScriptFont) Convert.Name), _
                                                          New Array() {Array.ConvertAll(Of IslamData.ScriptFont, String)(CachedData.IslamData.ScriptFonts, Function(Convert As IslamData.ScriptFont) Utility.MakeJSArray(Array.ConvertAll(Of IslamData.ScriptFont.Font, String)(Convert.FontList, Function(SubConv As IslamData.ScriptFont.Font) SubConv.ID)))}, True) + ";"
    End Function
    Public Shared Function GetFontEmbedJS() As String
        Return "function embedFontStyle(fontID) { if (isInArray(fontID, embeddedFonts)) return; embeddedFonts.push(fontID); var font=fontList[fontID]; var style = 'font-family: \'' + font.embed + '\';' + 'src: url(\'/files/\' + font.file + \'.eot\');' + 'src: local(\'' + font.family + '\'), url(\'/files/' + font.file + ((font.file == 'KFC_naskh') ? '.otf\') format(\'opentype\');' : '.ttf\') format(\'truetype\');'); addStyleSheetRule(newStyleSheet(), '@font-face', style);  }"
    End Function
    Public Shared Function GetFontInitJS() As String
        Return "var tryFontCounter = 0; var embeddedFonts = " + Utility.MakeJSArray(New String() {"null"}, True) + "; var baseFont = 'Times New Roman';"
    End Function
    Public Shared Function GetFontIDJS() As String
        Return "function getFontID() { var fontID = $('#fontselection').val(); if (fontID == 'def') { fontID = 'me_quran'; if (isMac && isSafari) fontID = 'scheherazade'; if (isChrome) fontID = getPrefInstalledFont('uthmani'); } return fontID; }"
    End Function
    Public Shared Function GetFontFaceJS() As String
        Return "function getFontFace(fontID) { return fontList[fontID].family + (fontList[fontID].embed ? ',' + fontList[fontID].embed : ''); }"
    End Function
    Public Shared Function GetFontWidthJS() As String
        Return "function fontWidth(fontName, text) { text = text || '" + Utility.EncodeJS(TanzilReader.GetTextVerse(TanzilReader.GetTextChapter(CachedData.XMLDocMain, 3), 9).Attributes.GetNamedItem("text").Value) + "' ; if (text == 2) text = '" + Utility.EncodeJS(Utility.LoadResourceString("IslamInfo_InTheNameOfAllah")) + "," + Utility.EncodeJS(TanzilReader.GetTextVerse(TanzilReader.GetTextChapter(CachedData.XMLDocMain, 1), 1).Attributes.GetNamedItem("text").Value) + "'; var tester = $('#font-tester'); tester.css('fontFamily', fontName); if (tester.firstChild) tester.remove(tester.firstChild); tester.append(document.createTextNode(text)); tester.css('display', 'block'); var width = tester.offsetWidth; tester.css('display', 'none'); return width; }"
    End Function
    Public Shared Function GetFontExistsJS() As String
        Return "function fontExists(fontName) { var fontFamily = fontName + ', ' + baseFont; return fontWidth(baseFont) * fontWidth(baseFont, 2) != fontWidth(fontFamily) * fontWidth(fontFamily, 2); }"
    End Function
    Public Shared Function GetApplyFontJS() As String
        Return "function applyFont(fontID) { if (!fontExists(getFontFace(fontID))) fontID = getPrefInstalledFont(); var font = fontList[fontID]; findStyleSheetRule('span.arabic').style.fontFamily = getFontFace(fontID); $('#fontloading').css('display', 'none'); }"
    End Function
    Public Shared Function GetTryFontJS() As String
        Return "function tryFont(fontID) { if (++tryFontCounter < 50 && !fontExists(getFontFace(fontID))) { setTimeout('tryFont(\'' + fontID + '\')', 400); return; } $('#fontloading').css('display', 'none'); applyFont(fontID); }"
    End Function
    Public Shared Function GetApplyEmbedFontJS() As String
        Return "function applyEmbedFont(fontID) { embedFontStyle(fontID); $('#fontloading').css('display', ''); tryFontCounter = 0; tryFont(fontID); }"
    End Function
    Public Shared Function GetFontPrefInstalledJS() As String
        Return "function getPrefInstalledFont(type) { var list = fontPrefs[type]; for(var i in list) { var fontID = list[i]; if (fontList[fontID].installed) return fontID; } return 'arial'; }"
    End Function
    Public Shared Function GetCheckInstalledFontsJS() As String
        Return "function checkInstalledFonts() { for (var i in fontList) { var font = fontList[i]; if (font.family && fontExists(font.family)) font.installed = true; } }"
    End Function
    Public Shared Function GetUpdateCustomFontJS() As String
        Return "function updateCustomFont() { var fontID = getFontID(); $('#fontcustom').css('display', fontID == 'custom' ? '' : 'none'); $('#fontcustomapply').css('display', fontID == 'custom' ? '' : 'none'); }"
    End Function
    Public Shared Function GetChangeCustomFontJS() As String()
        Return New String() {"javascript: changeCustomFont();", String.Empty, "function changeCustomFont() { fontList['custom'].family = $('#fontcustom').val(); fontList['custom'].scale = fontWidth(baseFont) / fontWidth(fontList['custom'].family); changeFont(); }"}
    End Function
    Public Shared Function GetChangeFontJS() As String()
        Return New String() {"javascript: changeFont();", "checkInstalledFonts();", Utility.GetLookupStyleSheetJS(), GetArabicFontListJS(), Utility.GetBrowserTestJS(), Utility.GetAddStyleSheetJS(), Utility.GetAddStyleSheetRuleJS(), Utility.GetLookupStyleSheetJS(), Utility.IsInArrayJS(), GetUpdateCustomFontJS(), GetFontInitJS(), GetFontPrefInstalledJS(), GetCheckInstalledFontsJS(), GetFontIDJS(), GetFontFaceJS(), GetFontWidthJS(), GetFontExistsJS(), GetFontEmbedJS(), GetApplyFontJS(), GetTryFontJS(), GetApplyEmbedFontJS(), _
        "function changeFont() { var fontID = getFontID(); updateCustomFont(); if (fontList[fontID].embed) applyEmbedFont(fontID); else applyFont(fontID); }"}
    End Function
    Public Shared Function GetFontSmallerJS() As String()
        Return New String() {"javascript: decreaseFontSize();", String.Empty, Utility.GetLookupStyleSheetJS(), _
        "function decreaseFontSize() { rule = findStyleSheetRule('span.arabic'); rule.style.fontSize = Math.max(parseInt(rule.style.fontSize.replace('px', '')) - 1, 1) + 'px'; }"}
    End Function
    Public Shared Function GetFontDefaultSizeJS() As String()
        Return New String() {"javascript: defaultFontSize();", String.Empty, Utility.GetLookupStyleSheetJS(), _
        "function defaultFontSize() { findStyleSheetRule('span.arabic').style.fontSize = '32px'; }"}
    End Function
    Public Shared Function GetFontBiggerJS() As String()
        Return New String() {"javascript: increaseFontSize();", String.Empty, Utility.GetLookupStyleSheetJS(), _
        "function increaseFontSize() { rule = findStyleSheetRule('span.arabic'); rule.style.fontSize = (parseInt(rule.style.fontSize.replace('px', '')) + 1) + 'px'; }"}
    End Function
End Class
Class AudioRecitation
    Function GetURL(Source As String, ReciterName As String, Chapter As Integer, Verse As Integer) As String
        Dim Base As String = String.Empty
        If Source = "everyayah" Then Base = "http://www.everyayah.com/data/"
        If Source = "tanzil" Then Base = "http://tanzil.net/res/audio/"
        Return Base + ReciterName + "/" + Chapter.ToString("D3") + Verse.ToString("D3") + ".mp3"
    End Function
End Class
<System.Xml.Serialization.XmlRoot("islamdata")> _
Public Class IslamData
    Public Structure PrayerType
        <System.Xml.Serialization.XmlAttribute("name")> _
        Public Name As String
        <System.Xml.Serialization.XmlAttribute("id")> _
        Public TranslationID As String
        <System.Xml.Serialization.XmlAttribute("classification")> _
        Public Classification As String
        <System.Xml.Serialization.XmlAttribute("prayerunits")> _
        Public PrayerUnits As String
    End Structure
    <System.Xml.Serialization.XmlArray("prayers")> _
    <System.Xml.Serialization.XmlArrayItem("prayertype")> _
    Public Prayers() As PrayerType
    Public Structure PrayerTime
        <System.Xml.Serialization.XmlAttribute("name")> _
        Public Name As String
        <System.Xml.Serialization.XmlAttribute("id")> _
        Public TranslationID As String
    End Structure
    <System.Xml.Serialization.XmlArray("prayertimes")> _
    <System.Xml.Serialization.XmlArrayItem("prayertime")> _
    Public PrayerTimes() As PrayerTime
    Public Structure DayOfWeek
        <System.Xml.Serialization.XmlAttribute("name")> _
        Public Name As String
        <System.Xml.Serialization.XmlAttribute("id")> _
        Public TranslationID As String
    End Structure
    <System.Xml.Serialization.XmlArray("daysofweek")> _
    <System.Xml.Serialization.XmlArrayItem("day")> _
    Public DaysOfWeek() As DayOfWeek
    Public Structure Month
        <System.Xml.Serialization.XmlAttribute("name")> _
        Public Name As String
        <System.Xml.Serialization.XmlAttribute("id")> _
        Public TranslationID As String
    End Structure
    <System.Xml.Serialization.XmlArray("months")> _
    <System.Xml.Serialization.XmlArrayItem("month")> _
    Public Months() As Month
    Public Structure VerseCategory
        Public Structure Verse
            <System.Xml.Serialization.XmlAttribute("title")> _
            Public TranslationID As String
            <System.Xml.Serialization.XmlAttribute("text")> _
            Public Arabic As String
        End Structure
        <System.Xml.Serialization.XmlAttribute("title")> _
        Public Title As String
        <System.Xml.Serialization.XmlElement("verse")> _
        Public Verses() As Verse
    End Structure
    <System.Xml.Serialization.XmlArray("verses")> _
    <System.Xml.Serialization.XmlArrayItem("category")> _
    Public VerseCategories() As VerseCategory

    Public Structure VocabCategory
        Public Structure Word
            <System.Xml.Serialization.XmlAttribute("text")> _
            Public Text As String
            <System.Xml.Serialization.XmlAttribute("id")> _
            Public TranslationID As String
        End Structure
        <System.Xml.Serialization.XmlAttribute("title")> _
        Public Title As String
        <System.Xml.Serialization.XmlElement("word")> _
        Public Words() As Word
    End Structure

    <System.Xml.Serialization.XmlArray("vocabulary")> _
    <System.Xml.Serialization.XmlArrayItem("category")> _
    Public VocabularyCategories() As VocabCategory

    <System.Xml.Serialization.XmlArray("abbreviations")> _
    <System.Xml.Serialization.XmlArrayItem("category")> _
    Public Abbreviations() As VocabCategory

    <System.Xml.Serialization.XmlArray("lists")> _
    <System.Xml.Serialization.XmlArrayItem("category")> _
    Public Lists() As VocabCategory

    Public Structure GrammarCategory
        Public Structure GrammarWord
            <System.Xml.Serialization.XmlAttribute("text")> _
            Public Text As String
            <System.Xml.Serialization.XmlAttribute("id")> _
            Public TranslationID As String
            <System.Xml.Serialization.XmlAttribute("grammar")> _
            Public Grammar As String
        End Structure
        <System.Xml.Serialization.XmlAttribute("title")> _
        Public Title As String
        <System.Xml.Serialization.XmlElement("word")> _
        Public Words() As GrammarWord
    End Structure

    <System.Xml.Serialization.XmlArray("grammar")> _
    <System.Xml.Serialization.XmlArrayItem("category")> _
    Public GrammarCategories() As GrammarCategory
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
    Public Structure TranslitScheme
        <System.Xml.Serialization.XmlAttribute("name")> _
        Public Name As String
        Public Alphabet() As String
        <System.Xml.Serialization.XmlAttribute("alphabet")> _
        Property AlphabetParse As String
            Get
                If Alphabet.Length = 0 Then Return String.Empty
                Return String.Join("|"c, Alphabet)
            End Get
            Set(value As String)
                If Not value Is Nothing Then
                    Alphabet = value.Split("|"c)
                End If
            End Set
        End Property
        Public Hamza() As String
        <System.Xml.Serialization.XmlAttribute("hamza")> _
        Property HamzaParse As String
            Get
                If Hamza.Length = 0 Then Return String.Empty
                Return String.Join("|"c, Hamza)
            End Get
            Set(value As String)
                If Not value Is Nothing Then
                    Hamza = value.Split("|"c)
                End If
            End Set
        End Property
        Public SpecialLetters() As String
        <System.Xml.Serialization.XmlAttribute("tehmarbutaalefmaksuradaggeralefgunnah")> _
        Property SpecialLettersParse As String
            Get
                If SpecialLetters.Length = 0 Then Return String.Empty
                Return String.Join("|"c, SpecialLetters)
            End Get
            Set(value As String)
                If Not value Is Nothing Then
                    SpecialLetters = value.Split("|"c)
                End If
            End Set
        End Property
        Public Vowels() As String
        <System.Xml.Serialization.XmlAttribute("fathadammakasratanweenlongvowelsdipthongsshaddasukun")> _
        Property VowelsParse As String
            Get
                If Vowels.Length = 0 Then Return String.Empty
                Return String.Join("|"c, Vowels)
            End Get
            Set(value As String)
                If Not value Is Nothing Then
                    Vowels = value.Split("|"c)
                End If
            End Set
        End Property
        Public Tajweed() As String
        <System.Xml.Serialization.XmlAttribute("tajweed")> _
        Property TajweedParse As String
            Get
                If Tajweed.Length = 0 Then Return String.Empty
                Return String.Join("|"c, Tajweed)
            End Get
            Set(value As String)
                If Not value Is Nothing Then
                    Tajweed = value.Split("|"c)
                End If
            End Set
        End Property
        Public Punctuation() As String
        <System.Xml.Serialization.XmlAttribute("punctuation")> _
        Property PunctuationParse As String
            Get
                If Punctuation.Length = 0 Then Return String.Empty
                Return String.Join("|"c, Punctuation)
            End Get
            Set(value As String)
                If Not value Is Nothing Then
                    Punctuation = value.Split("|"c)
                End If
            End Set
        End Property
        Public Numbers() As String
        <System.Xml.Serialization.XmlAttribute("number")> _
        Property NumbersParse As String
            Get
                If Numbers.Length = 0 Then Return String.Empty
                Return String.Join("|"c, Numbers)
            End Get
            Set(value As String)
                If Not value Is Nothing Then
                    Numbers = value.Split("|"c)
                End If
            End Set
        End Property
        Public NonArabic() As String
        <System.Xml.Serialization.XmlAttribute("nonarabic")> _
        Property NonArabicParse As String
            Get
                If NonArabic.Length = 0 Then Return String.Empty
                Return String.Join("|"c, NonArabic)
            End Get
            Set(value As String)
                If Not value Is Nothing Then
                    NonArabic = value.Split("|"c)
                End If
            End Set
        End Property
    End Structure
    <System.Xml.Serialization.XmlArray("translitschemes")> _
    <System.Xml.Serialization.XmlArrayItem("scheme")> _
    Public TranslitSchemes() As TranslitScheme
    Structure LanguageInfo
        <System.Xml.Serialization.XmlAttribute("code")> _
        Public Code As String
        <System.Xml.Serialization.XmlAttribute("rtl")> _
        Public IsRTL As Boolean
    End Structure
    <System.Xml.Serialization.XmlArray("languages")> _
    <System.Xml.Serialization.XmlArrayItem("language")> _
    Public LanguageList() As LanguageInfo

    Structure ArabicFontList
        <System.Xml.Serialization.XmlAttribute("id")> _
        Public ID As String
        <System.Xml.Serialization.XmlAttribute("name")> _
        Public Name As String
        <System.Xml.Serialization.XmlAttribute("family")> _
        Public Family As String
        <System.Xml.Serialization.XmlAttribute("embedname")> _
        Public EmbedName As String
        <System.Xml.Serialization.XmlAttribute("file")> _
        Public FileName As String
        <System.Xml.Serialization.XmlAttribute("scale")> _
        <ComponentModel.DefaultValueAttribute(-1.0)> _
        Public Scale As Double
    End Structure
    <System.Xml.Serialization.XmlArray("arabicfonts")> _
    <System.Xml.Serialization.XmlArrayItem("arabicfont")> _
    Public ArabicFonts() As ArabicFontList

    Public Structure ScriptFont
        Public Structure Font
            <System.Xml.Serialization.XmlAttribute("id")> _
            Public ID As String
        End Structure
        <System.Xml.Serialization.XmlAttribute("name")> _
        Public Name As String
        <System.Xml.Serialization.XmlElement("font")> _
        Public FontList() As Font
    End Structure

    <System.Xml.Serialization.XmlArray("scriptfonts")> _
    <System.Xml.Serialization.XmlArrayItem("scriptfont")> _
    Public ScriptFonts() As ScriptFont

    Public Class TranslationsInfo
        Public Structure TranslationInfo
            <System.Xml.Serialization.XmlAttribute("name")> _
            Public Name As String
            <System.Xml.Serialization.XmlAttribute("file")> _
            Public FileName As String
            <System.Xml.Serialization.XmlAttribute("translator")> _
            Public Translator As String
        End Structure
        <System.Xml.Serialization.XmlAttribute("default")> _
        Public DefaultTranslation As String
        <System.Xml.Serialization.XmlElement("translation")> _
        Public TranslationList() As TranslationInfo
    End Class
    <System.Xml.Serialization.XmlElement("translations")> _
    Public Translations As TranslationsInfo
    Structure QuranSelection
        Structure QuranSelectionInfo
            <System.Xml.Serialization.XmlAttribute("chapter")> _
            Public ChapterNumber As Integer
            <System.Xml.Serialization.XmlAttribute("startverse")> _
            Public VerseNumber As Integer
            <ComponentModel.DefaultValueAttribute(1)> _
            <System.Xml.Serialization.XmlAttribute("startword")> _
            Public WordNumber As Integer
            <ComponentModel.DefaultValueAttribute(0)> _
            <System.Xml.Serialization.XmlAttribute("endword")> _
            Public EndWordNumber As Integer
            <ComponentModel.DefaultValueAttribute(0)> _
            <System.Xml.Serialization.XmlAttribute("endverse")> _
            Public ExtraVerseNumber As Integer
        End Structure
        <System.Xml.Serialization.XmlAttribute("description")> _
        Public Description As String
        <System.Xml.Serialization.XmlElement("verse")> _
        Public SelectionInfo As QuranSelectionInfo()
    End Structure
    <System.Xml.Serialization.XmlArray("quranselections")> _
    <System.Xml.Serialization.XmlArrayItem("quranselection")> _
    Public QuranSelections As QuranSelection()
    Structure QuranDivision
        <System.Xml.Serialization.XmlAttribute("description")> _
        Public Description As String
    End Structure
    <System.Xml.Serialization.XmlArray("qurandivisions")> _
    <System.Xml.Serialization.XmlArrayItem("division")> _
    Public QuranDivisions As QuranDivision()

    Structure QuranChapter
        <System.Xml.Serialization.XmlAttribute("index")> _
        Public Index As String
        <System.Xml.Serialization.XmlAttribute("name")> _
        Public Name As String
        <ComponentModel.DefaultValueAttribute(0)> _
        <System.Xml.Serialization.XmlAttribute("uniqueletters")> _
        Public UniqueLetters As Integer
    End Structure
    <System.Xml.Serialization.XmlArray("quranchapters")> _
    <System.Xml.Serialization.XmlArrayItem("chapter")> _
    Public QuranChapters As QuranChapter()

    Structure QuranPart
        <System.Xml.Serialization.XmlAttribute("index")> _
        Public Index As String
        <System.Xml.Serialization.XmlAttribute("name")> _
        Public Name As String
        <System.Xml.Serialization.XmlAttribute("id")> _
        Public ID As String
    End Structure
    <System.Xml.Serialization.XmlArray("quranparts")> _
    <System.Xml.Serialization.XmlArrayItem("part")> _
    Public QuranParts As QuranPart()

    Structure CollectionInfo
        Structure CollTranslationInfo
            <System.Xml.Serialization.XmlAttribute("name")> _
            Public Name As String
            <System.Xml.Serialization.XmlAttribute("file")> _
            Public FileName As String
        End Structure
        <System.Xml.Serialization.XmlAttribute("name")> _
        Public Name As String
        <System.Xml.Serialization.XmlAttribute("file")> _
        Public FileName As String
        <System.Xml.Serialization.XmlAttribute("default")> _
        Public DefaultTranslation As String
        <System.Xml.Serialization.XmlArray("translations")> _
        <System.Xml.Serialization.XmlArrayItem("translation")> _
        Public Translations() As CollTranslationInfo
    End Structure
    <System.Xml.Serialization.XmlArray("hadithcollections")> _
    <System.Xml.Serialization.XmlArrayItem("collection")> _
    Public Collections() As CollectionInfo
    Structure PartOfSpeechInfo
        <System.Xml.Serialization.XmlAttribute("symbol")> _
        Public Symbol As String
        <System.Xml.Serialization.XmlAttribute("id")> _
        Public Id As String
    End Structure
    <System.Xml.Serialization.XmlArray("partsofspeech")> _
    <System.Xml.Serialization.XmlArrayItem("pos")> _
    Public PartsOfSpeech() As PartOfSpeechInfo
End Class
Public Class CachedData
    'need disk and memory cache as time consuming to read or build
    Shared _ObjIslamData As IslamData
    Shared _XMLDocMain As System.Xml.XmlDocument 'Tanzil Quran data
    Shared _XMLDocInfo As System.Xml.XmlDocument 'Tanzil metadata
    Shared _XMLDocInfos As Collections.Generic.List(Of System.Xml.XmlDocument) 'Hadiths
    Shared _RootDictionary As New Generic.Dictionary(Of String, ArrayList)
    Shared _FormDictionary As New Generic.Dictionary(Of String, ArrayList)
    Shared _TagDictionary As New Generic.Dictionary(Of String, Generic.Dictionary(Of String, ArrayList))
    Shared _WordDictionary As New Generic.Dictionary(Of String, ArrayList)
    Shared _LetterDictionary As New Generic.Dictionary(Of Char, Generic.Dictionary(Of String, ArrayList))
    Shared _LetterPreDictionary As New Generic.Dictionary(Of Char, Generic.Dictionary(Of String, ArrayList))
    Shared _LetterSufDictionary As New Generic.Dictionary(Of Char, Generic.Dictionary(Of String, ArrayList))
    Shared _PreDictionary As New Generic.Dictionary(Of String, ArrayList)
    Shared _SufDictionary As New Generic.Dictionary(Of String, ArrayList)
    Shared _IsolatedLetterDictionary As New Generic.Dictionary(Of Char, ArrayList)
    Shared _TotalLetters As Integer = 0
    Shared _TotalIsolatedLetters As Integer = 0
    Shared _PartUniqueArray(TanzilReader.GetPartCount() - 1) As Generic.List(Of String)
    Shared _PartArray(TanzilReader.GetPartCount() - 1) As Generic.List(Of String)
    Shared _StationUniqueArray(TanzilReader.GetPartCount() - 1) As Generic.List(Of String)
    Shared _StationArray(TanzilReader.GetPartCount() - 1) As Generic.List(Of String)
    Shared _TotalUniqueWordsInParts As Integer = 0
    Shared _TotalWordsInParts As Integer = 0
    Shared _TotalUniqueWordsInStations As Integer = 0
    Shared _TotalWordsInStations As Integer = 0
    Public Shared Sub GetMorphologicalData()
        Dim Lines As String() = IO.File.ReadAllLines(Utility.GetFilePath("metadata\quranic-corpus-morphology-0.4.txt"))
        For Count As Integer = 0 To Lines.Length - 1
            If Lines(Count).Length <> 0 AndAlso Lines(Count).Chars(0) <> "#" Then
                'LOCATION	FORM	TAG	FEATURES
                Dim Pieces As String() = Lines(Count).Split(CChar(vbTab))
                'FORM can be found identically in tanzil
                If Pieces(0).Chars(0) = "(" Then
                    If Not _FormDictionary.ContainsKey(Pieces(1)) Then
                        _FormDictionary.Add(Pieces(1), New ArrayList)
                    End If
                    'TAG
                    If Not _TagDictionary.ContainsKey(Pieces(2)) Then
                        _TagDictionary.Add(Pieces(2), New Generic.Dictionary(Of String, ArrayList))
                    End If
                    Dim Location As Integer() = Array.ConvertAll(Pieces(0).TrimStart("("c).TrimEnd(")"c).Split(":"c), Function(Str As String) CInt(Str))
                    _FormDictionary.Item(Pieces(1)).Add(Location)
                    If Not _TagDictionary.Item(Pieces(2)).ContainsKey(Pieces(1)) Then
                        _TagDictionary.Item(Pieces(2)).Add(Pieces(1), New ArrayList)
                    End If
                    _TagDictionary.Item(Pieces(2)).Item(Pieces(1)).Add(Location)
                    Dim Parts As String() = Pieces(3).Split("|"c)
                    If Array.Find(Parts, Function(Str As String) Str = "PREFIX" Or Str = "SUFFIX") = String.Empty Then
                        'LEM: or if not present FORM
                        Dim Lem As String = Array.Find(Parts, Function(Str As String) Str.StartsWith("LEM:"))
                        If Lem <> String.Empty Then
                            Lem = Lem.Replace("LEM:", String.Empty)
                        Else
                            Lem = Pieces(1)
                        End If
                        If Not _WordDictionary.ContainsKey(Lem) Then
                            _WordDictionary.Add(Lem, New ArrayList)
                        End If
                        _WordDictionary.Item(Lem).Add(Location)
                    End If
                    If Array.Find(Parts, Function(Str As String) Str = "PREFIX") <> String.Empty Then
                        If Not _PreDictionary.ContainsKey(Pieces(1)) Then
                            _PreDictionary.Add(Pieces(1), New ArrayList)
                        End If
                        _PreDictionary.Item(Pieces(1)).Add(Location)
                    ElseIf Array.Find(Parts, Function(Str As String) Str = "SUFFIX") <> String.Empty Then
                        If Not _SufDictionary.ContainsKey(Pieces(1)) Then
                            _SufDictionary.Add(Pieces(1), New ArrayList)
                        End If
                        _SufDictionary.Item(Pieces(1)).Add(Location)
                    End If
                    'ROOT:
                    Dim Root As String = Array.Find(Parts, Function(Str As String) Str.StartsWith("ROOT:"))
                    If Root <> String.Empty Then
                        Root = Root.Replace("ROOT:", String.Empty)
                        If Not _RootDictionary.ContainsKey(Root) Then
                            _RootDictionary.Add(Root, New ArrayList)
                        End If
                        _RootDictionary.Item(Root).Add(Location)
                    End If
                End If
            End If
        Next
    End Sub
    Public Shared Sub GetMorphDataByDivision(bStation As Boolean, _
                                             ByRef All As Integer, ByRef AllUnique As Integer, _
                                             ByRef PartArray() As Generic.List(Of String), _
                                             ByRef PartUniqueArray() As Generic.List(Of String))
        Dim FreqArray(CachedData.WordDictionary.Keys.Count - 1) As String
        CachedData.WordDictionary.Keys.CopyTo(FreqArray, 0)
        For Count As Integer = 1 To CInt(IIf(bStation, TanzilReader.GetStationCount(), TanzilReader.GetPartCount()))
            PartUniqueArray(Count - 1) = New Generic.List(Of String)
            PartArray(Count - 1) = New Generic.List(Of String)
            Dim Node As System.Xml.XmlNode = CType(IIf(bStation, TanzilReader.GetStationByIndex(Count), TanzilReader.GetPartByIndex(Count)), System.Xml.XmlNode)
            Dim BaseChapter As Integer = CInt(Node.Attributes.GetNamedItem("sura").Value)
            Dim BaseVerse As Integer = CInt(Node.Attributes.GetNamedItem("aya").Value)
            Dim Chapter As Integer
            Dim Verse As Integer
            Node = CType(IIf(bStation, TanzilReader.GetStationByIndex(Count + 1), TanzilReader.GetPartByIndex(Count + 1)), System.Xml.XmlNode)
            If Node Is Nothing Then
                Chapter = TanzilReader.GetChapterCount()
                Verse = TanzilReader.GetVerseCount(Chapter)
            Else
                Chapter = CInt(Node.Attributes.GetNamedItem("sura").Value)
                Verse = CInt(Node.Attributes.GetNamedItem("aya").Value)
                TanzilReader.GetPreviousChapterVerse(Chapter, Verse)
            End If
            For SubCount As Integer = 0 To FreqArray.Length - 1
                Dim RefCount As Integer
                Dim UniCount As Integer = 0
                For RefCount = 0 To CachedData.WordDictionary(FreqArray(SubCount)).Count - 1
                    If (CType(CachedData.WordDictionary(FreqArray(SubCount))(RefCount), Integer())(0) = BaseChapter AndAlso _
                        CType(CachedData.WordDictionary(FreqArray(SubCount))(RefCount), Integer())(1) >= BaseVerse AndAlso _
                        (BaseChapter <> Chapter OrElse _
                        CType(CachedData.WordDictionary(FreqArray(SubCount))(RefCount), Integer())(1) <= Verse)) OrElse _
                        (CType(CachedData.WordDictionary(FreqArray(SubCount))(RefCount), Integer())(0) > BaseChapter AndAlso _
                        CType(CachedData.WordDictionary(FreqArray(SubCount))(RefCount), Integer())(0) < Chapter) OrElse _
                        (CType(CachedData.WordDictionary(FreqArray(SubCount))(RefCount), Integer())(0) = Chapter AndAlso
                        CType(CachedData.WordDictionary(FreqArray(SubCount))(RefCount), Integer())(1) <= Verse) Then
                        UniCount += 1
                    End If
                Next
                If UniCount = CachedData.WordDictionary(FreqArray(SubCount)).Count Then
                    PartUniqueArray(Count - 1).Add(FreqArray(SubCount))
                    AllUnique += 1
                End If
                If UniCount > 0 Then
                    PartArray(Count - 1).Add(FreqArray(SubCount))
                    All += 1
                End If
            Next
        Next
    End Sub
    Public Shared Sub BuildQuranLetterIndex()
        Dim Verses As Collections.Generic.List(Of String())
        Verses = TanzilReader.GetQuranText(CachedData.XMLDocMain, -1, -1, -1, -1)
        For Count As Integer = 0 To Verses.Count - 1
            For SubCount As Integer = 0 To Verses(Count).Length - 1
                For LetCount As Integer = 0 To Verses(Count)(SubCount).Length - 1
                    _TotalLetters += 1
                    If Not _LetterDictionary.ContainsKey(Verses(Count)(SubCount)(LetCount)) Then
                        _LetterDictionary.Add(Verses(Count)(SubCount)(LetCount), New Dictionary(Of String, ArrayList))
                    End If
                    If Not _LetterPreDictionary.ContainsKey(Verses(Count)(SubCount)(LetCount)) Then
                        _LetterPreDictionary.Add(Verses(Count)(SubCount)(LetCount), New Dictionary(Of String, ArrayList))
                    End If
                    If Not _LetterSufDictionary.ContainsKey(Verses(Count)(SubCount)(LetCount)) Then
                        _LetterSufDictionary.Add(Verses(Count)(SubCount)(LetCount), New Dictionary(Of String, ArrayList))
                    End If
                    Dim PrevIndex As Integer = Verses(Count)(SubCount).LastIndexOf(" "c, LetCount) + If(Verses(Count)(SubCount)(LetCount) = " ", 0, 1)
                    Dim NextIndex As Integer = Verses(Count)(SubCount).IndexOf(" "c, LetCount)
                    If NextIndex = -1 Then NextIndex = Verses(Count)(SubCount).Length
                    If Not _LetterDictionary.Item(Verses(Count)(SubCount)(LetCount)).ContainsKey(Verses(Count)(SubCount).Substring(PrevIndex, NextIndex - PrevIndex)) Then
                        _LetterDictionary.Item(Verses(Count)(SubCount)(LetCount)).Add(Verses(Count)(SubCount).Substring(PrevIndex, NextIndex - PrevIndex), New ArrayList)
                    End If
                    _LetterDictionary.Item(Verses(Count)(SubCount)(LetCount))(Verses(Count)(SubCount).Substring(PrevIndex, NextIndex - PrevIndex)).Add(New Integer() {Count, SubCount, LetCount})
                    If Not _LetterPreDictionary.Item(Verses(Count)(SubCount)(LetCount)).ContainsKey(Verses(Count)(SubCount).Substring(PrevIndex, LetCount - PrevIndex)) Then
                        _LetterPreDictionary.Item(Verses(Count)(SubCount)(LetCount)).Add(Verses(Count)(SubCount).Substring(PrevIndex, LetCount - PrevIndex), New ArrayList)
                    End If
                    _LetterPreDictionary.Item(Verses(Count)(SubCount)(LetCount))(Verses(Count)(SubCount).Substring(PrevIndex, LetCount - PrevIndex)).Add(New Integer() {Count, SubCount, LetCount})
                    If LetCount <> NextIndex Then
                        If Not _LetterSufDictionary.Item(Verses(Count)(SubCount)(LetCount)).ContainsKey(Verses(Count)(SubCount).Substring(LetCount + 1, NextIndex - LetCount - 1)) Then
                            _LetterSufDictionary.Item(Verses(Count)(SubCount)(LetCount)).Add(Verses(Count)(SubCount).Substring(LetCount + 1, NextIndex - LetCount - 1), New ArrayList)
                        End If
                        _LetterSufDictionary.Item(Verses(Count)(SubCount)(LetCount))(Verses(Count)(SubCount).Substring(LetCount + 1, NextIndex - LetCount - 1)).Add(New Integer() {Count, SubCount, LetCount})
                    End If
                    If LetCount <> 0 AndAlso LetCount <> Verses(Count)(SubCount).Length - 1 AndAlso _
                        Char.IsWhiteSpace(Verses(Count)(SubCount)(LetCount - 1)) AndAlso Char.IsWhiteSpace(Verses(Count)(SubCount)(LetCount + 1)) Then
                        _TotalIsolatedLetters += 1
                        If Not _IsolatedLetterDictionary.ContainsKey(Verses(Count)(SubCount)(LetCount)) Then
                            _IsolatedLetterDictionary.Add(Verses(Count)(SubCount)(LetCount), New ArrayList)
                        End If
                        _IsolatedLetterDictionary.Item(Verses(Count)(SubCount)(LetCount)).Add(New Integer() {Count, SubCount, LetCount})
                    End If
                Next
            Next
        Next
    End Sub
    Public Shared Sub DoErrorCheck()
        'missing from loader
        'loanwords word value/id, jurisprudence name/id, category name/id, months name/id
        'daysofweek name/id, fasting fastingtype name/id, islamicbooks book title/author
        Dim Count As Integer
        Dim Verses As Collections.Generic.List(Of String())
        Verses = TanzilReader.GetQuranText(CachedData.XMLDocMain, -1, -1, -1, -1)
        For Count = 0 To Verses.Count - 1
            Dim ChapterNode As System.Xml.XmlNode = TanzilReader.GetTextChapter(CachedData.XMLDocMain, Count + 1)
            For SubCount As Integer = 0 To Verses(Count).Length - 1
                If SubCount = 0 AndAlso Not TanzilReader.GetTextVerse(ChapterNode, SubCount + 1).Attributes.GetNamedItem("bismillah") Is Nothing Then
                    Arabic.DoErrorCheck(TanzilReader.GetTextVerse(ChapterNode, SubCount + 1).Attributes.GetNamedItem("bismillah").Value)
                End If
                Arabic.DoErrorCheck(Verses(Count)(SubCount))
            Next
        Next
        For Count = 0 To IslamData.Months.Length - 1
            Arabic.DoErrorCheck(Arabic.TransliterateFromBuckwalter(IslamData.Months(Count).Name))
            Utility.LoadResourceString("IslamInfo_" + IslamData.Months(Count).TranslationID)
        Next
        For Count = 0 To IslamData.DaysOfWeek.Length - 1
            Arabic.DoErrorCheck(Arabic.TransliterateFromBuckwalter(IslamData.DaysOfWeek(Count).Name))
            Utility.LoadResourceString("IslamInfo_" + IslamData.DaysOfWeek(Count).TranslationID)
        Next
        For Count = 0 To IslamData.Prayers.Length - 1
            Arabic.DoErrorCheck(Arabic.TransliterateFromBuckwalter(IslamData.Prayers(Count).Name))
            Utility.LoadResourceString("IslamInfo_" + IslamData.Prayers(Count).TranslationID)
        Next
        For Count = 0 To IslamData.PrayerTimes.Length - 1
            Arabic.DoErrorCheck(Arabic.TransliterateFromBuckwalter(IslamData.PrayerTimes(Count).Name))
            Utility.LoadResourceString("IslamInfo_" + IslamData.PrayerTimes(Count).TranslationID)
        Next
        'must check Trans and WordForWord
        For Count = 0 To IslamData.VerseCategories.Length - 1
            Utility.LoadResourceString("IslamInfo_" + IslamData.VerseCategories(Count).Title)
            For SubCount As Integer = 0 To IslamData.VerseCategories(Count).Verses.Length - 1
                Arabic.DoErrorCheck(Arabic.TransliterateFromBuckwalter(IslamData.VerseCategories(Count).Verses(SubCount).Arabic))
                Utility.LoadResourceString("IslamInfo_" + IslamData.VerseCategories(Count).Verses(SubCount).TranslationID)
            Next
        Next
        For Count = 0 To IslamData.GrammarCategories.Length - 1
            Utility.LoadResourceString("IslamInfo_" + IslamData.GrammarCategories(Count).Title)
            For SubCount As Integer = 0 To IslamData.GrammarCategories(Count).Words.Length - 1
                Arabic.DoErrorCheck(Arabic.TransliterateFromBuckwalter(IslamData.GrammarCategories(Count).Words(SubCount).Text))
                Utility.LoadResourceString("IslamInfo_" + IslamData.GrammarCategories(Count).Words(SubCount).TranslationID)
            Next
        Next
        For Count = 0 To IslamData.VocabularyCategories.Length - 1
            Utility.LoadResourceString("IslamInfo_" + IslamData.VocabularyCategories(Count).Title)
            For SubCount As Integer = 0 To IslamData.VocabularyCategories(Count).Words.Length - 1
                Arabic.DoErrorCheck(Arabic.TransliterateFromBuckwalter(IslamData.VocabularyCategories(Count).Words(SubCount).Text))
                Utility.LoadResourceString("IslamInfo_" + IslamData.VocabularyCategories(Count).Words(SubCount).TranslationID)
            Next
        Next
        For Count = 0 To IslamData.ArabicLetters.Length - 1
            Arabic.DoErrorCheck(Arabic.TransliterateFromBuckwalter(IslamData.ArabicLetters(Count).SymbolName))
            Utility.LoadResourceString("IslamInfo_" + IslamData.ArabicLetters(Count).UnicodeName)
        Next
        For Count = 0 To IslamData.Collections.Length - 1
            Utility.LoadResourceString("IslamInfo_" + IslamData.Collections(Count).Name)
        Next
        For Count = 0 To IslamData.QuranDivisions.Length - 1
            Utility.LoadResourceString("IslamInfo_" + IslamData.QuranDivisions(Count).Description)
        Next
        For Count = 0 To IslamData.QuranSelections.Length - 1
            Utility.LoadResourceString("IslamInfo_" + IslamData.QuranSelections(Count).Description)
        Next
        For Count = 0 To IslamData.QuranChapters.Length - 1
            Arabic.DoErrorCheck(Arabic.TransliterateFromBuckwalter(IslamData.QuranChapters(Count).Name))
        Next
        For Count = 0 To IslamData.QuranParts.Length - 1
            Arabic.DoErrorCheck(Arabic.TransliterateFromBuckwalter(IslamData.QuranParts(Count).Name))
            Utility.LoadResourceString("IslamInfo_" + IslamData.QuranParts(Count).ID)
        Next
        For Count = 0 To IslamData.PartsOfSpeech.Length - 1
            Utility.LoadResourceString("IslamInfo_" + IslamData.PartsOfSpeech(Count).Id)
        Next
    End Sub
    Public Shared ReadOnly Property IslamData As IslamData
        Get
            If _ObjIslamData Is Nothing Then
                Dim fs As IO.FileStream = New IO.FileStream(Utility.GetFilePath("metadata\islaminfo.xml"), IO.FileMode.Open, IO.FileAccess.Read)
                Dim xs As System.Xml.Serialization.XmlSerializer = New System.Xml.Serialization.XmlSerializer(GetType(IslamData))
                _ObjIslamData = CType(xs.Deserialize(fs), IslamData)
                fs.Close()
            End If
            Return _ObjIslamData
        End Get
    End Property
    Public Shared ReadOnly Property XMLDocMain As System.Xml.XmlDocument
        Get
            If _XMLDocMain Is Nothing Then
                _XMLDocMain = New System.Xml.XmlDocument
                _XMLDocMain.Load(Utility.GetFilePath("metadata\quran-uthmani.xml"))
            End If
            Return _XMLDocMain
        End Get
    End Property
    Public Shared ReadOnly Property XMLDocInfo As System.Xml.XmlDocument
        Get
            If _XMLDocInfo Is Nothing Then
                _XMLDocInfo = New System.Xml.XmlDocument
                _XMLDocInfo.Load(Utility.GetFilePath("metadata\quran-data.xml"))
            End If
            Return _XMLDocInfo
        End Get
    End Property
    Public Shared ReadOnly Property XMLDocInfos As Collections.Generic.List(Of System.Xml.XmlDocument)
        Get
            Dim Count As Integer
            If _XMLDocInfos Is Nothing Then
                _XMLDocInfos = New Collections.Generic.List(Of System.Xml.XmlDocument)
                For Count = 0 To CachedData.IslamData.Collections.Length - 1
                    _XMLDocInfos.Add(New System.Xml.XmlDocument)
                    _XMLDocInfos(_XMLDocInfos.Count - 1).Load(Utility.GetFilePath("metadata\" + CachedData.IslamData.Collections(Count).FileName + "-data.xml"))
                Next
            End If
            Return _XMLDocInfos
        End Get
    End Property
    Public Shared ReadOnly Property RootDictionary As Generic.Dictionary(Of String, ArrayList)
        Get
            If _RootDictionary.Keys.Count = 0 Then GetMorphologicalData()
            Return _RootDictionary
        End Get
    End Property
    Public Shared ReadOnly Property FormDictionary As Generic.Dictionary(Of String, ArrayList)
        Get
            If _FormDictionary.Keys.Count = 0 Then GetMorphologicalData()
            Return _FormDictionary
        End Get
    End Property
    Public Shared ReadOnly Property TagDictionary As Generic.Dictionary(Of String, Generic.Dictionary(Of String, ArrayList))
        Get
            If _TagDictionary.Keys.Count = 0 Then GetMorphologicalData()
            Return _TagDictionary
        End Get
    End Property
    Public Shared ReadOnly Property WordDictionary As Generic.Dictionary(Of String, ArrayList)
        Get
            If _WordDictionary.Keys.Count = 0 Then GetMorphologicalData()
            Return _WordDictionary
        End Get
    End Property
    Public Shared ReadOnly Property LetterDictionary As Generic.Dictionary(Of Char, Generic.Dictionary(Of String, ArrayList))
        Get
            If _LetterDictionary.Keys.Count = 0 Then BuildQuranLetterIndex()
            Return _LetterDictionary
        End Get
    End Property
    Public Shared ReadOnly Property LetterPreDictionary As Generic.Dictionary(Of Char, Generic.Dictionary(Of String, ArrayList))
        Get
            If _LetterPreDictionary.Keys.Count = 0 Then BuildQuranLetterIndex()
            Return _LetterPreDictionary
        End Get
    End Property
    Public Shared ReadOnly Property LetterSufDictionary As Generic.Dictionary(Of Char, Generic.Dictionary(Of String, ArrayList))
        Get
            If _LetterSufDictionary.Keys.Count = 0 Then BuildQuranLetterIndex()
            Return _LetterSufDictionary
        End Get
    End Property
    Public Shared ReadOnly Property PreDictionary As Generic.Dictionary(Of String, ArrayList)
        Get
            If _PreDictionary.Keys.Count = 0 Then GetMorphologicalData()
            Return _PreDictionary
        End Get
    End Property
    Public Shared ReadOnly Property SufDictionary As Generic.Dictionary(Of String, ArrayList)
        Get
            If _SufDictionary.Keys.Count = 0 Then GetMorphologicalData()
            Return _SufDictionary
        End Get
    End Property
    Public Shared ReadOnly Property TotalLetters As Integer
        Get
            If _TotalLetters = 0 Then BuildQuranLetterIndex()
            Return _TotalLetters
        End Get
    End Property
    Public Shared ReadOnly Property IsolatedLetterDictionary As Generic.Dictionary(Of Char, ArrayList)
        Get
            If _IsolatedLetterDictionary.Keys.Count = 0 Then BuildQuranLetterIndex()
            Return _IsolatedLetterDictionary
        End Get
    End Property
    Public Shared ReadOnly Property TotalIsolatedLetters As Integer
        Get
            If _TotalIsolatedLetters = 0 Then BuildQuranLetterIndex()
            Return _TotalIsolatedLetters
        End Get
    End Property
    Public Shared ReadOnly Property PartArray As Generic.List(Of String)()
        Get
            If TotalWordsInParts = 0 Then GetMorphDataByDivision(False, _TotalWordsInParts, _TotalUniqueWordsInParts, _PartArray, _PartUniqueArray)
            Return _PartArray
        End Get
    End Property
    Public Shared ReadOnly Property PartUniqueArray As Generic.List(Of String)()
        Get
            If TotalUniqueWordsInParts = 0 Then GetMorphDataByDivision(False, _TotalWordsInParts, _TotalUniqueWordsInParts, _PartArray, _PartUniqueArray)
            Return _PartUniqueArray
        End Get
    End Property
    Public Shared ReadOnly Property TotalWordsInParts As Integer
        Get
            If TotalWordsInParts = 0 Then GetMorphDataByDivision(False, _TotalWordsInParts, _TotalUniqueWordsInParts, _PartArray, _PartUniqueArray)
            Return _TotalWordsInParts
        End Get
    End Property
    Public Shared ReadOnly Property TotalUniqueWordsInParts As Integer
        Get
            If TotalUniqueWordsInParts = 0 Then GetMorphDataByDivision(False, _TotalWordsInParts, _TotalUniqueWordsInParts, _PartArray, _PartUniqueArray)
            Return _TotalUniqueWordsInParts
        End Get
    End Property
    Public Shared ReadOnly Property StationArray As Generic.List(Of String)()
        Get
            If _TotalWordsInStations = 0 Then GetMorphDataByDivision(True, _TotalWordsInStations, _TotalUniqueWordsInStations, _StationArray, _StationUniqueArray)
            Return _StationArray
        End Get
    End Property
    Public Shared ReadOnly Property StationUniqueArray As Generic.List(Of String)()
        Get
            If TotalUniqueWordsInStations = 0 Then GetMorphDataByDivision(True, _TotalWordsInStations, _TotalUniqueWordsInStations, _StationArray, _StationUniqueArray)
            Return _StationUniqueArray
        End Get
    End Property
    Public Shared ReadOnly Property TotalWordsInStations As Integer
        Get
            If _TotalWordsInStations = 0 Then GetMorphDataByDivision(True, _TotalWordsInStations, _TotalUniqueWordsInStations, _StationArray, _StationUniqueArray)
            Return _TotalWordsInStations
        End Get
    End Property
    Public Shared ReadOnly Property TotalUniqueWordsInStations As Integer
        Get
            If TotalUniqueWordsInStations = 0 Then GetMorphDataByDivision(True, _TotalWordsInStations, _TotalUniqueWordsInStations, _StationArray, _StationUniqueArray)
            Return _TotalUniqueWordsInStations
        End Get
    End Property
End Class
Public Class Languages
    Public Shared Function GetLanguageInfoByCode(ByVal Code As String) As IslamData.LanguageInfo
        Dim Count As Integer
        For Count = 0 To CachedData.IslamData.LanguageList.Length - 1
            If CachedData.IslamData.LanguageList(Count).Code = Code Then Return CachedData.IslamData.LanguageList(Count)
        Next
        Return Nothing
    End Function
End Class
Public Class DocBuilder
    Public Shared Function GetRenderedText(ByVal Item As PageLoader.TextItem) As RenderArray
        Return TextFromReferences(HttpContext.Current.Request.QueryString.Get("docedit"))
    End Function
    Public Shared Function TextFromReferences(Strings As String) As RenderArray
        Dim Renderer As New RenderArray
        If Strings = Nothing Then Return Renderer
        Dim Matches As System.Text.RegularExpressions.MatchCollection = System.Text.RegularExpressions.Regex.Matches(Strings, "(.*?)(?:(\{)(.*?)(\})|$)")
        For Count As Integer = 0 To Matches.Count - 1
            If Matches(Count).Length <> 0 Then
                If Matches(Count).Groups(1).Length <> 0 Then

                End If
                If Matches(Count).Groups(3).Length <> 0 Then
                    'text before and after reference matches needs rendering
                    'hadith reference matching {name,book/hadith}
                    If TanzilReader.IsQuranTextReference(Matches(Count).Groups(3).Value) Then
                        Renderer.Items.AddRange(TanzilReader.QuranTextFromReference(Matches(Count).Groups(3).Value).Items)
                    ElseIf Matches(Count).Groups(3).Value.StartsWith("letter:") Then
                    ElseIf Matches(Count).Groups(3).Value.StartsWith("personalpronoun:") Then
                    ElseIf Matches(Count).Groups(3).Value.StartsWith("possessivedeterminerpersonalpronoun:") Then
                    ElseIf Matches(Count).Groups(3).Value.StartsWith("particle:") Then
                    ElseIf Matches(Count).Groups(3).Value.StartsWith("noun:") Then
                    ElseIf Matches(Count).Groups(3).Value.StartsWith("verb:") Then
                    ElseIf Matches(Count).Groups(3).Value.StartsWith("phrase:") Then
                    ElseIf Matches(Count).Groups(3).Value.StartsWith("reference:") Then
                    ElseIf Array.IndexOf(CachedData.IslamData.Abbreviations, Matches(Count).Groups(3).Value) <> 0 Then
                    End If
                End If
            End If
        Next
        Return Renderer
    End Function
End Class
Public Class Supplications
    Public Shared Function GetSuppCategories() As String()
        Return Array.ConvertAll(CachedData.IslamData.VerseCategories, Function(Convert As IslamData.VerseCategory) Utility.LoadResourceString("IslamInfo_" + Convert.Title))
    End Function
    Public Shared Function GetRenderedSuppText(ByVal Item As PageLoader.TextItem) As RenderArray
        Dim SchemeType As Arabic.TranslitScheme = If(CInt(HttpContext.Current.Request.QueryString.Get("translitscheme")) >= 2, 2 + CInt(HttpContext.Current.Request.QueryString.Get("translitscheme")) Mod 2, CInt(HttpContext.Current.Request.QueryString.Get("translitscheme")))
        Dim Scheme As String = If(CInt(HttpContext.Current.Request.QueryString.Get("translitscheme")) >= 2, CachedData.IslamData.TranslitSchemes((CInt(HttpContext.Current.Request.QueryString.Get("translitscheme")) - 2) \ 2).Name, String.Empty)
        Dim Count As Integer = CInt(HttpContext.Current.Request.QueryString.Get("selection"))
        If Count = -1 Then Count = 0
        Return DoGetRenderedSuppText(SchemeType, Scheme, CachedData.IslamData.VerseCategories(Count))
    End Function
    Public Shared Function DoGetRenderedSuppText(SchemeType As Arabic.TranslitScheme, Scheme As String, Category As IslamData.VerseCategory) As RenderArray
        Dim Renderer As New RenderArray
        For SubCount As Integer = 0 To Category.Verses.Length - 1
            Dim Matches As System.Text.RegularExpressions.MatchCollection = System.Text.RegularExpressions.Regex.Matches(Category.Verses(SubCount).Arabic, "(.*?)(?:(\\\{)(.*?)(\\\})|$)")
            For MatchCount As Integer = 0 To Matches.Count - 1
                If Matches(MatchCount).Length <> 0 Then
                    If Matches(MatchCount).Groups(1).Length <> 0 Then
                        Dim EnglishByWord As String() = Utility.LoadResourceString("IslamInfo_" + Category.Verses(SubCount).TranslationID + "WordByWord").Split("|"c)
                        Dim ArabicText As String() = Matches(MatchCount).Groups(1).Value.Split(" "c)
                        Dim Transliteration As String() = Arabic.TransliterateToScheme(Arabic.TransliterateFromBuckwalter(Matches(MatchCount).Groups(1).Value), SchemeType, Scheme).Split(" "c)
                        Renderer.Items.Add(New RenderArray.RenderItem(RenderArray.RenderTypes.eHeaderCenter, New RenderArray.RenderText() {New RenderArray.RenderText(RenderArray.RenderDisplayClass.eLTR, Utility.LoadResourceString("IslamInfo_" + Category.Verses(SubCount).TranslationID))}))
                        Dim Items As New Collections.Generic.List(Of RenderArray.RenderItem)
                        For WordCount As Integer = 0 To EnglishByWord.Length - 1
                            Items.Add(New RenderArray.RenderItem(RenderArray.RenderTypes.eText, New RenderArray.RenderText() {New RenderArray.RenderText(RenderArray.RenderDisplayClass.eArabic, Arabic.RightToLeftMark + Arabic.TransliterateFromBuckwalter(ArabicText(WordCount))), New RenderArray.RenderText(RenderArray.RenderDisplayClass.eTransliteration, Transliteration(WordCount)), New RenderArray.RenderText(RenderArray.RenderDisplayClass.eLTR, EnglishByWord(WordCount))}))
                        Next
                        Renderer.Items.Add(New RenderArray.RenderItem(RenderArray.RenderTypes.eText, New RenderArray.RenderText() {New RenderArray.RenderText(RenderArray.RenderDisplayClass.eNested, Items), New RenderArray.RenderText(RenderArray.RenderDisplayClass.eArabic, Arabic.RightToLeftMark + Arabic.TransliterateFromBuckwalter(Matches(MatchCount).Groups(1).Value)), New RenderArray.RenderText(RenderArray.RenderDisplayClass.eTransliteration, Arabic.TransliterateToScheme(Arabic.TransliterateFromBuckwalter(Matches(MatchCount).Groups(1).Value), SchemeType, Scheme)), New RenderArray.RenderText(RenderArray.RenderDisplayClass.eLTR, Utility.LoadResourceString("IslamInfo_" + Category.Verses(SubCount).TranslationID + "Trans"))}))
                    End If
                    If Matches(MatchCount).Groups(3).Length <> 0 Then
                        If TanzilReader.IsQuranTextReference(Matches(MatchCount).Groups(3).Value) Then
                            Renderer.Items.AddRange(TanzilReader.QuranTextFromReference(Matches(MatchCount).Groups(3).Value).Items)
                        Else
                            Renderer.Items.Add(New RenderArray.RenderItem(RenderArray.RenderTypes.eHeaderCenter, New RenderArray.RenderText() {New RenderArray.RenderText(RenderArray.RenderDisplayClass.eLTR, "(" + Matches(MatchCount).Groups(3).Value + ")")}))
                        End If
                    End If
                End If
            Next
        Next
        Return Renderer
    End Function
End Class
Public Class Quiz
    Public Shared Function GetQuizList() As Array()
        Dim Strings(2) As Array
        Strings(0) = New String() {"ArabicLetters", "arabicletters"}
        Strings(1) = New String() {"ArabicDiacriticsLetters", "arabicdiacriticsletters"}
        Strings(2) = New String() {"ArabicLettersDiacritics", "arabiclettersdiacritics"}
        Return Strings
    End Function
    Public Shared ArabicSpecialLetters As String() = {Arabic.ArabicLetterAlefWithHamzaAbove, Arabic.ArabicLetterAlefWithHamzaBelow, Arabic.ArabicLetterWawWithHamzaAbove, Arabic.ArabicLetterYehWithHamzaAbove, Arabic.ArabicLetterAlefWasla, Arabic.ArabicLetterAlefWithMaddaAbove, Arabic.ArabicLetterHamza, Arabic.ArabicLetterTehMarbuta, Arabic.ArabicLetterAlefMaksura}
    Public Shared ArabicDiacriticsBefore As String() = {Arabic.ArabicSukun, Arabic.ArabicFatha, Arabic.ArabicKasra, Arabic.ArabicDamma}
    Public Shared ArabicDiacriticsAfter As String() = {Arabic.ArabicSukun, Arabic.ArabicFatha, Arabic.ArabicKasra, Arabic.ArabicDamma, Arabic.ArabicKasratan, Arabic.ArabicDammatan, Arabic.ArabicFathatan + Arabic.ArabicLetterAlef, Arabic.ArabicShadda}
    Public Shared Function GetChangeQuizJS() As String()
        Return New String() {"javascript: changeQuiz();", String.Empty, _
                             "function changeQuiz() { qtype = $('#quizselection').val(); qwrong = 0; qright = 0; nextQuestion(); }"}
    End Function
    Public Shared Function DisplayCount(ByVal Item As PageLoader.TextItem) As String
        Return "Wrong: 0 Right: 0"
    End Function
    Public Shared Function GetQuizSet() As String()
        Dim Quiz As Integer = HttpContext.Current.Request.QueryString.Get("quizselection")
        If Quiz = 0 Then
            Return Arabic.ArabicLetters
        ElseIf Quiz = 1 Then
            Dim CurList As New Generic.List(Of String)
            For Count As Integer = 0 To ArabicDiacriticsBefore.Length - 1
                Dim CurLet As String = ArabicDiacriticsBefore(Count)
                CurList.AddRange(Array.ConvertAll(Arabic.ArabicLetters, Function(Str As String) CurLet + Str))
            Next
            Return CurList.ToArray()
        ElseIf Quiz = 2 Then
            Dim CurList As New Generic.List(Of String)
            For Count As Integer = 0 To ArabicDiacriticsAfter.Length - 1
                Dim CurLet As String = ArabicDiacriticsAfter(Count)
                CurList.AddRange(Array.ConvertAll(Arabic.ArabicLetters, Function(Str As String) Str + CurLet))
            Next
            Return CurList.ToArray()
        Else
            Return Nothing
        End If
    End Function
    Public Shared Function DisplayQuestion(ByVal Item As PageLoader.TextItem) As String
        HttpContext.Current.Items.Add("rnd", DateTime.Now.ToFileTime())
        Rnd(-1)
        Randomize(CDbl(HttpContext.Current.Items("rnd")))
        Dim Count As Integer = CInt(Math.Floor(Rnd() * 4))
        While Count
            Rnd()
            Count -= 1
        End While
        Dim QuizSet As String() = GetQuizSet()
        Return QuizSet(CInt(Math.Floor(Rnd() * QuizSet.Length)))
    End Function
    Public Shared Function DisplayAnswer(ByVal Item As PageLoader.ButtonItem) As String
        Dim Count As Integer
        Rnd(-1)
        Randomize(CDbl(HttpContext.Current.Items("rnd")))
        Rnd()
        Dim Quiz As Integer = HttpContext.Current.Request.QueryString.Get("quizselection")
        For Count = 2 To Integer.Parse(Item.Name.Replace("answer", String.Empty))
            Rnd()
            If Quiz <> 0 Then Rnd()
        Next
        Dim QuizSet As String() = GetQuizSet()
        Return Arabic.TransliterateToScheme(QuizSet(CInt(Math.Floor(Rnd() * QuizSet.Length))), CInt(HttpContext.Current.Request.QueryString.Get("translitscheme")) + 1, String.Empty)
    End Function
    Public Shared Function VerifyAnswer() As String()
        Dim JSList As New List(Of String) From {"javascript: verifyAnswer(this);", String.Empty, _
            Arabic.GetArabicSymbolJSArray(), Arabic.FindLetterBySymbolJS, _
            "var arabicLets = " + Utility.MakeJSArray(Arabic.ArabicLetters, False) + ";", _
            "var arabicDiacriticsBefore = " + Utility.MakeJSArray(ArabicDiacriticsBefore, False) + ";", _
            "var arabicDiacriticsAfter = " + Utility.MakeJSArray(ArabicDiacriticsAfter, False) + ";", _
            "var qtype = 'arabicletters', qwrong = 0, qright = 0;", _
            "function getUniqueRnd(excl, count) { var rnd; if (excl.length === count) return 0; do { rnd = Math.floor(Math.random() * count); } while (excl.indexOf(rnd) !== -1); return rnd; }", _
            "function getQuizSet() { if (qtype === 'arabicletters') return arabicLets; if (qtype === 'arabiclettersdiacritics') { var count = 0, arr = []; for (count = 0; count < arabicDiacriticsAfter.length; count++) { arr = arr.concat(arabicLets.map(function(val) { return val + arabicDiacriticsAfter[count]; })); } return arr; } if (qtype === 'arabicdiacriticsletters') { var count = 0, arr = []; for (count = 0; count < arabicDiacriticsBefore.length; count++) { arr = arr.concat(arabicLets.map(function(val) { return arabicDiacriticsBefore[count] + val; })); } return arr; }; return []; }", _
            "function getQA(quizSet, quest, nidx) { if (quest) return quizSet[nidx]; return $('#translitscheme').val() === '0' ? transliterateWithRules(quizSet[nidx]) : doTransliterate(quizSet[nidx], true, $('#translitscheme').val() === '2'); }", _
            "function nextQuestion() { $('#count').text('Wrong: ' + qwrong + ' Right: ' + qright); var i = Math.floor(Math.random() * 4), quizSet = getQuizSet(), pos = quizSet.length, nidx = getUniqueRnd([], pos), aidx = []; aidx[0] = getUniqueRnd([nidx], pos); aidx[1] = getUniqueRnd([nidx, aidx[0]], pos); aidx[2] = getUniqueRnd([nidx, aidx[0], aidx[1]], pos); $('#quizquestion').text(getQA(quizSet, true, nidx)); $('#answer1').prop('value', getQA(quizSet, false, i === 0 ? nidx : aidx[0])); $('#answer2').prop('value', getQA(quizSet, false, i === 1 ? nidx : aidx[i > 1 ? 1 : 0])); $('#answer3').prop('value', getQA(quizSet, false, i === 2 ? nidx : aidx[i > 2 ? 2 : 1])); $('#answer4').prop('value', getQA(quizSet, false, i === 3 ? nidx : aidx[2])); }", _
            "function verifyAnswer(ctl) { $(ctl).prop('value') === ($('#translitscheme').val() === '0' ? transliterateWithRules($('#quizquestion').text().trim()) : doTransliterate($('#quizquestion').text().trim(), true, $('#translitscheme').val() === '2')) ? qright++ : qwrong++; nextQuestion(); }"}
        JSList.AddRange(Arabic.PlainTransliterateGenJS)
        JSList.AddRange(Arabic.TransliterateGenJS)
        Return JSList.ToArray()
    End Function
End Class
Public Class TanzilReader
    Public Shared Function GetDivisionTypes() As String()
        Return Array.ConvertAll(CachedData.IslamData.QuranDivisions, Function(Convert As IslamData.QuranDivision) Utility.LoadResourceString("IslamInfo_" + Convert.Description))
    End Function
    Public Shared Function GetTranslationList() As Array()
        Return Array.ConvertAll(CachedData.IslamData.Translations.TranslationList, Function(Convert As IslamData.TranslationsInfo.TranslationInfo) New String() {Utility.LoadResourceString("lang_local" + Languages.GetLanguageInfoByCode(Convert.FileName.Substring(0, CInt(IIf(Convert.FileName.IndexOf("-") <> -1, Convert.FileName.IndexOf("-"), Convert.FileName.IndexOf("."))))).Code) + ": " + Convert.Name, Convert.FileName})
    End Function
    Public Shared Function GetTranslationIndex(ByVal Translation As String) As Integer
        If String.IsNullOrEmpty(Translation) Then Translation = CachedData.IslamData.Translations.DefaultTranslation 'Default
        Dim Count As Integer = Array.FindIndex(CachedData.IslamData.Translations.TranslationList, Function(Test As IslamData.TranslationsInfo.TranslationInfo) Test.FileName = Translation)
        If Count = -1 Then
            Translation = CachedData.IslamData.Translations.DefaultTranslation 'Default
            Count = Array.FindIndex(CachedData.IslamData.Translations.TranslationList, Function(Test As IslamData.TranslationsInfo.TranslationInfo) Test.FileName = Translation)
        End If
        Return Count
    End Function
    Public Shared Function GetTranslationFileName(ByVal Translation As String) As String
        Dim Index As Integer = GetTranslationIndex(Translation)
        Return CachedData.IslamData.Translations.TranslationList(Index).FileName + ".txt"
    End Function
    Public Shared Function GetDivisionChangeJS() As String()
        Dim JSArrays As String = Utility.MakeJSArray(New String() {Utility.MakeJSArray(Array.ConvertAll(Of Array, String)(TanzilReader.GetChapterNames(), Function(Convert As Array) Utility.MakeJSArray(New String() {CStr(CType(Convert, Object())(0)), CStr(CType(Convert, Object())(1))})), True), _
            Utility.MakeJSArray(Array.ConvertAll(Of Array, String)(TanzilReader.GetChapterNamesByRevelationOrder(), Function(Convert As Array) Utility.MakeJSArray(New String() {CStr(CType(Convert, Object())(0)), CStr(CType(Convert, Object())(1))})), True), _
            Utility.MakeJSArray(Array.ConvertAll(Of Array, String)(TanzilReader.GetPartNames(), Function(Convert As Array) Utility.MakeJSArray(New String() {CStr(CType(Convert, Object())(0)), CStr(CType(Convert, Object())(1))})), True), _
            Utility.MakeJSArray(Array.ConvertAll(Of Array, String)(TanzilReader.GetGroupNames(), Function(Convert As Array) Utility.MakeJSArray(New String() {CStr(CType(Convert, Object())(0)), CStr(CType(Convert, Object())(1))})), True), _
            Utility.MakeJSArray(Array.ConvertAll(Of Array, String)(TanzilReader.GetStationNames(), Function(Convert As Array) Utility.MakeJSArray(New String() {CStr(CType(Convert, Object())(0)), CStr(CType(Convert, Object())(1))})), True), _
            Utility.MakeJSArray(Array.ConvertAll(Of Array, String)(TanzilReader.GetSectionNames(), Function(Convert As Array) Utility.MakeJSArray(New String() {CStr(CType(Convert, Object())(0)), CStr(CType(Convert, Object())(1))})), True), _
            Utility.MakeJSArray(Array.ConvertAll(Of Array, String)(TanzilReader.GetPageNames(), Function(Convert As Array) Utility.MakeJSArray(New String() {CStr(CType(Convert, Object())(0)), CStr(CType(Convert, Object())(1))})), True), _
            Utility.MakeJSArray(Array.ConvertAll(Of Array, String)(TanzilReader.GetSajdaNames(), Function(Convert As Array) Utility.MakeJSArray(New String() {CStr(CType(Convert, Object())(0)), CStr(CType(Convert, Object())(1))})), True), _
            Utility.MakeJSArray(Array.ConvertAll(Of Array, String)(TanzilReader.GetImportantNames(), Function(Convert As Array) Utility.MakeJSArray(New String() {CStr(CType(Convert, Object())(0)), CStr(CType(Convert, Object())(1))})), True), _
            Utility.MakeJSArray(Array.ConvertAll(Of Array, String)(Arabic.GetRecitationSymbols(), Function(Convert As Array) Utility.MakeJSArray(New String() {CStr(CType(Convert, Object())(0)), CStr(CType(Convert, Object())(1))})), True)}, True)
        Return New String() {"javascript: changeQuranDivision(this.selectedIndex);", String.Empty, Utility.GetClearOptionListJS(), _
        "function changeQuranDivision(index) { var iCount; var qurandata = " + JSArrays + "; var eSelect = $('#quranselection').get(0); clearOptionList(eSelect); for (iCount = 0; iCount < qurandata[index].length; iCount++) { eSelect.options.add(new Option(qurandata[index][iCount][0], qurandata[index][iCount][1])); } }"}
    End Function
    Public Shared Function GetWordPartitions() As String()
        Dim Parts As New Generic.List(Of String) From {Utility.LoadResourceString("IslamInfo_Letters"), Utility.LoadResourceString("IslamInfo_Words"), Utility.LoadResourceString("IslamInfo_UniqueWords"), Utility.LoadResourceString("IslamInfo_UniqueWordsPerPart"), Utility.LoadResourceString("IslamInfo_WordsPerPart"), Utility.LoadResourceString("IslamInfo_UniqueWordsPerStation"), Utility.LoadResourceString("IslamInfo_WordsPerStation"), Utility.LoadResourceString("IslamInfo_IsolatedLetters"), Utility.LoadResourceString("IslamInfo_LetterPatterns"), Utility.LoadResourceString("IslamInfo_Prefix"), Utility.LoadResourceString("IslamInfo_Suffix")}
        Parts.AddRange(Array.ConvertAll(CachedData.IslamData.PartsOfSpeech, Function(POS As IslamData.PartOfSpeechInfo) Utility.LoadResourceString("IslamInfo_" + POS.Id)))
        Parts.AddRange(Array.ConvertAll(Arabic.RecitationSymbols, Function(Sym As String) Arabic.GetUnicodeName(Sym)))
        Parts.AddRange(Array.ConvertAll(Arabic.RecitationSymbols, Function(Sym As String) "Prefix of " + Arabic.GetUnicodeName(Sym)))
        Parts.AddRange(Array.ConvertAll(Arabic.RecitationSymbols, Function(Sym As String) "Suffix of " + Arabic.GetUnicodeName(Sym)))
        Return Parts.ToArray()
    End Function
    Public Shared Function GetQuranWordTotalNumber() As Integer
        Dim Total As Integer
        For Each Key As String In CachedData.WordDictionary.Keys
            Total = Total + CachedData.WordDictionary.Item(Key).Count
        Next
        Return Total
    End Function
    Public Shared Function GetQuranWordTotal(ByVal Item As PageLoader.TextItem) As String
        Dim Strings As String
        Dim Index As Integer
        Strings = HttpContext.Current.Request.QueryString.Get("quranselection")
        If Not Strings Is Nothing Then Index = CInt(Strings)
        If Index = 0 Then
            Return CStr(CachedData.TotalLetters)
        ElseIf Index = 7 Then
            Return CStr(CachedData.TotalIsolatedLetters)
        ElseIf Index = 1 Then
            Return CStr(GetQuranWordTotalNumber())
        ElseIf Index = 2 Then
            Return CStr(CachedData.WordDictionary.Keys.Count)
        ElseIf Index = 3 Then
            Return CStr(CachedData.TotalUniqueWordsInParts)
        ElseIf Index = 4 Then
            Return CStr(CachedData.TotalWordsInParts)
        ElseIf Index = 5 Then
            Return CStr(CachedData.TotalUniqueWordsInStations)
        ElseIf Index = 6 Then
            Return CStr(CachedData.TotalWordsInStations)
        ElseIf Index = 8 Then
            Return String.Empty
        ElseIf Index = 9 Then
            Return CStr(CachedData.PreDictionary.Count)
        ElseIf Index = 10 Then
            Return CStr(CachedData.SufDictionary.Count)
        ElseIf Index >= 11 And Index < 11 + CachedData.IslamData.PartsOfSpeech.Length Then
            Return CStr(CachedData.TagDictionary.Item(CachedData.IslamData.PartsOfSpeech(Index - 11).Symbol).Count)
        ElseIf Index >= 11 + CachedData.IslamData.PartsOfSpeech.Length And Index < 11 + CachedData.IslamData.PartsOfSpeech.Length + Arabic.RecitationSymbols.Length Then
            Return CStr(CachedData.LetterDictionary.Item(Arabic.RecitationSymbols(Index - 11 - CachedData.IslamData.PartsOfSpeech.Length)).Count)
        ElseIf Index >= 11 + CachedData.IslamData.PartsOfSpeech.Length + Arabic.RecitationSymbols.Length And Index < 11 + CachedData.IslamData.PartsOfSpeech.Length + Arabic.RecitationSymbols.Length + Arabic.RecitationSymbols.Length Then
            Return CStr(CachedData.LetterPreDictionary.Item(Arabic.RecitationSymbols(Index - 11 - CachedData.IslamData.PartsOfSpeech.Length - Arabic.RecitationSymbols.Length)).Count)
        ElseIf Index >= 11 + CachedData.IslamData.PartsOfSpeech.Length + Arabic.RecitationSymbols.Length + Arabic.RecitationSymbols.Length And Index < 11 + CachedData.IslamData.PartsOfSpeech.Length + Arabic.RecitationSymbols.Length + Arabic.RecitationSymbols.Length + Arabic.RecitationSymbols.Length Then
            Return CStr(CachedData.LetterSufDictionary.Item(Arabic.RecitationSymbols(Index - 11 - CachedData.IslamData.PartsOfSpeech.Length - Arabic.RecitationSymbols.Length - Arabic.RecitationSymbols.Length)).Count)
        Else
            Return String.Empty
        End If
    End Function
    Public Shared Function GetQuranWordFrequency(ByVal Item As PageLoader.TextItem) As Array()
        Dim Output As New ArrayList
        Dim Total As Integer = 0
        Dim All As Double
        Output.Add(New String() {String.Empty, String.Empty, String.Empty, String.Empty, String.Empty})
        Output.Add(New String() {"arabic", "transliteration", String.Empty, String.Empty, String.Empty})
        Output.Add(New String() {Utility.LoadResourceString("IslamInfo_Arabic"), Utility.LoadResourceString("IslamInfo_Transliteration"), Utility.LoadResourceString("IslamSource_WordTotal"), String.Empty, String.Empty})
        Dim Strings As String
        Dim Index As Integer
        Strings = HttpContext.Current.Request.QueryString.Get("quranselection")
        If Not Strings Is Nothing Then Index = CInt(Strings)
        If Index = 0 Then
            All = CachedData.TotalLetters
            Dim LetterFreqArray(CachedData.LetterDictionary.Keys.Count - 1) As Char
            CachedData.LetterDictionary.Keys.CopyTo(LetterFreqArray, 0)
            Array.Sort(LetterFreqArray, Function(Key As Char, NextKey As Char) CachedData.LetterDictionary.Item(NextKey).Count.CompareTo(CachedData.LetterDictionary.Item(Key).Count))
            For Count As Integer = 0 To LetterFreqArray.Length - 1
                Total += CachedData.LetterDictionary.Item(LetterFreqArray(Count)).Count
                Output.Add(New String() {Arabic.LeftToRightMark + Arabic.GetUnicodeName(LetterFreqArray(Count)) + " ( " + Arabic.FixStartingCombiningSymbol(LetterFreqArray(Count)) + Arabic.LeftToRightOverride + " )" + Arabic.PopDirectionalFormatting, String.Empty, CStr(CachedData.LetterDictionary.Item(LetterFreqArray(Count)).Count), (CDbl(CachedData.LetterDictionary.Item(LetterFreqArray(Count)).Count) * 100 / All).ToString("n2"), (CDbl(Total) * 100 / All).ToString("n2")})
            Next
        ElseIf Index = 7 Then
            All = CachedData.TotalIsolatedLetters
            Dim LetterFreqArray(CachedData.IsolatedLetterDictionary.Keys.Count - 1) As Char
            CachedData.IsolatedLetterDictionary.Keys.CopyTo(LetterFreqArray, 0)
            Array.Sort(LetterFreqArray, Function(Key As Char, NextKey As Char) CachedData.IsolatedLetterDictionary.Item(NextKey).Count.CompareTo(CachedData.IsolatedLetterDictionary.Item(Key).Count))
            For Count As Integer = 0 To LetterFreqArray.Length - 1
                Total += CachedData.IsolatedLetterDictionary.Item(LetterFreqArray(Count)).Count
                Output.Add(New String() {Arabic.LeftToRightMark + Arabic.GetUnicodeName(LetterFreqArray(Count)) + " ( " + Arabic.FixStartingCombiningSymbol(LetterFreqArray(Count)) + Arabic.LeftToRightOverride + " )" + Arabic.PopDirectionalFormatting, String.Empty, CStr(CachedData.IsolatedLetterDictionary.Item(LetterFreqArray(Count)).Count), (CDbl(CachedData.IsolatedLetterDictionary.Item(LetterFreqArray(Count)).Count) * 100 / All).ToString("n2"), (CDbl(Total) * 100 / All).ToString("n2")})
            Next
        ElseIf Index = 1 Or Index = 9 Or Index = 10 Or Index >= 11 And Index < 11 + CachedData.IslamData.PartsOfSpeech.Length + Arabic.RecitationSymbols.Length + Arabic.RecitationSymbols.Length + Arabic.RecitationSymbols.Length Then
            Dim Dict As Generic.Dictionary(Of String, ArrayList)
            If Index = 1 Then
                Dict = CachedData.WordDictionary
            ElseIf Index = 9 Then
                Dict = CachedData.PreDictionary
            ElseIf Index = 10 Then
                Dict = CachedData.SufDictionary
            ElseIf Index >= 11 And Index < 11 + CachedData.IslamData.PartsOfSpeech.Length Then
                Dict = CachedData.TagDictionary(CachedData.IslamData.PartsOfSpeech(Index - 11).Symbol)
            ElseIf Index >= 11 + CachedData.IslamData.PartsOfSpeech.Length And Index < 11 + CachedData.IslamData.PartsOfSpeech.Length + Arabic.RecitationSymbols.Length Then
                Dict = CachedData.LetterDictionary(Arabic.RecitationSymbols(Index - 11 - CachedData.IslamData.PartsOfSpeech.Length))
            ElseIf Index >= 11 + CachedData.IslamData.PartsOfSpeech.Length + Arabic.RecitationSymbols.Length And Index < 11 + CachedData.IslamData.PartsOfSpeech.Length + Arabic.RecitationSymbols.Length + Arabic.RecitationSymbols.Length Then
                Dict = CachedData.LetterPreDictionary(Arabic.RecitationSymbols(Index - 11 - CachedData.IslamData.PartsOfSpeech.Length - Arabic.RecitationSymbols.Length))
            ElseIf Index >= 11 + CachedData.IslamData.PartsOfSpeech.Length + Arabic.RecitationSymbols.Length + Arabic.RecitationSymbols.Length And Index < 11 + CachedData.IslamData.PartsOfSpeech.Length + Arabic.RecitationSymbols.Length + Arabic.RecitationSymbols.Length + Arabic.RecitationSymbols.Length Then
                Dict = CachedData.LetterSufDictionary(Arabic.RecitationSymbols(Index - 11 - CachedData.IslamData.PartsOfSpeech.Length - Arabic.RecitationSymbols.Length - Arabic.RecitationSymbols.Length))
            Else
                Dict = Nothing
            End If
            Dim FreqArray(Dict.Keys.Count - 1) As String
            Dict.Keys.CopyTo(FreqArray, 0)
            Total = 0
            All = GetQuranWordTotalNumber()
            Array.Sort(FreqArray, Function(Key As String, NextKey As String) Dict.Item(NextKey).Count.CompareTo(Dict.Item(Key).Count))
            For Count As Integer = 0 To FreqArray.Length - 1
                Total += Dict.Item(FreqArray(Count)).Count
                Output.Add(New String() {Arabic.TransliterateFromBuckwalter(FreqArray(Count)), String.Empty, CStr(Dict.Item(FreqArray(Count)).Count), (CDbl(Dict.Item(FreqArray(Count)).Count) * 100 / All).ToString("n2"), (CDbl(Total) * 100 / All).ToString("n2")})
            Next
        ElseIf Index = 3 Or Index = 4 Or Index = 5 Or Index = 6 Then
            Total = 0
            Dim DivArray As Collections.Generic.List(Of String)()
            If Index = 3 Or Index = 5 Then
                DivArray = If(Index = 5, CachedData.StationUniqueArray, CachedData.PartUniqueArray)
                All = If(Index = 5, CachedData.TotalUniqueWordsInStations, CachedData.TotalUniqueWordsInParts)
                For Count As Integer = 0 To CInt(IIf(Index = 5, TanzilReader.GetStationCount(), TanzilReader.GetPartCount())) - 1
                    Total += DivArray(Count).Count
                    Output.Add(New String() {Arabic.LeftToRightMark + CStr(Count + 1), String.Empty, CStr(DivArray(Count).Count), (CDbl(DivArray(Count).Count) * 100 / All).ToString("n2"), (CDbl(Total) * 100 / All).ToString("n2")})
                Next
            ElseIf Index = 4 Or Index = 6 Then
                DivArray = If(Index = 6, CachedData.StationUniqueArray, CachedData.PartUniqueArray)
                All = If(Index = 6, CachedData.TotalWordsInStations, CachedData.TotalWordsInParts)
                For Count As Integer = 0 To CInt(IIf(Index = 6, TanzilReader.GetStationCount(), TanzilReader.GetPartCount())) - 1
                    Total += DivArray(Count).Count
                    Output.Add(New String() {Arabic.LeftToRightMark + CStr(Count + 1), String.Empty, CStr(DivArray(Count).Count), (CDbl(DivArray(Count).Count) * 100 / All).ToString("n2"), (CDbl(Total) * 100 / All).ToString("n2")})
                Next
            End If
        ElseIf Index = 8 Then
            Output.AddRange(Array.ConvertAll(GetQuranLetterPatterns(), Function(Str As String) {Str}))
        End If
        Return CType(Output.ToArray(GetType(Array)), Array())
    End Function
    Public Shared Function GetQuranLetterPatterns() As String()
        Dim RecSymbols As String = String.Join(String.Empty, Array.ConvertAll(Arabic.RecitationSpecialSymbols, Function(C As Char) CStr(C)))
        Dim LtrSymbols As String = String.Join(String.Empty, Array.ConvertAll(Arabic.RecitationLetters, Function(C As Char) CStr(C)))
        Dim DiaSymbols As String = String.Join(String.Empty, Array.ConvertAll(Arabic.RecitationDiacritics, Function(C As Char) CStr(C)))
        Dim StartWordMultiOnly As New Generic.Dictionary(Of String, String)
        Dim EndWordMultiOnly As New Generic.Dictionary(Of String, String)
        Dim MiddleWordMultiOnly As New Generic.Dictionary(Of String, String)
        Dim StartWordOnly As String = String.Join(String.Empty, Array.ConvertAll(Arabic.RecitationLettersDiacritics, Function(C As Char) CStr(C)))
        Dim NotStartWord As String = String.Join(String.Empty, Array.ConvertAll(Arabic.RecitationLettersDiacritics, Function(C As Char) CStr(C)))
        Dim EndWordOnly As String = String.Join(String.Empty, Array.ConvertAll(Arabic.RecitationLettersDiacritics, Function(C As Char) CStr(C)))
        Dim NotEndWord As String = String.Join(String.Empty, Array.ConvertAll(Arabic.RecitationLettersDiacritics, Function(C As Char) CStr(C)))
        Dim EndWordOnlyNoDia As String = String.Join(String.Empty, Array.ConvertAll(Arabic.RecitationLetters, Function(C As Char) CStr(C)))
        Dim NotEndWordNoDia As String = String.Join(String.Empty, Array.ConvertAll(Arabic.RecitationLetters, Function(C As Char) CStr(C)))
        Dim MiddleWordOnlyNoDia As String = String.Join(String.Empty, Array.ConvertAll(Arabic.RecitationLetters, Function(C As Char) CStr(C)))
        Dim NotMiddleWordNoDia As String = String.Join(String.Empty, Array.ConvertAll(Arabic.RecitationLetters, Function(C As Char) CStr(C)))
        Dim MiddleWordOnly As String = String.Join(String.Empty, Array.ConvertAll(Arabic.RecitationLettersDiacritics, Function(C As Char) CStr(C)))
        Dim NotMiddleWord As String = String.Join(String.Empty, Array.ConvertAll(Arabic.RecitationLettersDiacritics, Function(C As Char) CStr(C)))
        Dim DiaStartWordOnly As String = String.Join(String.Empty, Array.ConvertAll(Arabic.RecitationDiacritics, Function(C As Char) CStr(C)))
        Dim DiaNotStartWord As String = String.Join(String.Empty, Array.ConvertAll(Arabic.RecitationDiacritics, Function(C As Char) CStr(C)))
        Dim DiaEndWordOnly As String = String.Join(String.Empty, Array.ConvertAll(Arabic.RecitationDiacritics, Function(C As Char) CStr(C)))
        Dim DiaNotEndWord As String = String.Join(String.Empty, Array.ConvertAll(Arabic.RecitationDiacritics, Function(C As Char) CStr(C)))
        Dim DiaMiddleWordOnly As String = String.Join(String.Empty, Array.ConvertAll(Arabic.RecitationDiacritics, Function(C As Char) CStr(C)))
        Dim DiaNotMiddleWord As String = String.Join(String.Empty, Array.ConvertAll(Arabic.RecitationDiacritics, Function(C As Char) CStr(C)))
        Dim Combos As String() = String.Join("|", Array.ConvertAll(Arabic.RecitationLettersDiacritics, Function(C As Char) String.Join("|", Array.ConvertAll(Arabic.RecitationLettersDiacritics, Function(Nxt As Char) C + Nxt)))).Split("|")
        Dim DiaCombos As String() = String.Join("|", Array.ConvertAll(Arabic.RecitationDiacritics, Function(C As Char) String.Join("|", Array.ConvertAll(Arabic.RecitationDiacritics, Function(Nxt As Char) C + Nxt)))).Split("|")
        Dim LetCombos As String() = String.Join("|", Array.ConvertAll(Arabic.RecitationLetters, Function(C As Char) String.Join("|", Array.ConvertAll(Arabic.RecitationLetters, Function(Nxt As Char) C + Nxt)))).Split("|")
        For Each Key As String In CachedData.FormDictionary.Keys
            Dim Str As String = New String(Array.FindAll(Key.ToCharArray(), Function(Ch As Char) Not RecSymbols.Contains(CStr(Ch))))
            For Count = 1 To Str.Length - 2
                If Not EndWordMultiOnly.ContainsKey(Str.Substring(Count)) Then
                    EndWordMultiOnly.Add(Str.Substring(Count), Nothing)
                End If
                If Not StartWordMultiOnly.ContainsKey(Str.Substring(0, Count + 1)) Then
                    StartWordMultiOnly.Add(Str.Substring(0, Count + 1), Nothing)
                End If
                For SubCount As Integer = 2 To Str.Length - 1 - Count
                    If Not MiddleWordMultiOnly.ContainsKey(Str.Substring(Count, SubCount)) Then
                        MiddleWordMultiOnly.Add(Str.Substring(Count, SubCount), Nothing)
                    End If
                Next
            Next
        Next
        For Each Key As String In CachedData.FormDictionary.Keys
            Dim Str As String = New String(Array.FindAll(Key.ToCharArray(), Function(Ch As Char) Not RecSymbols.Contains(CStr(Ch))))
            Str = Arabic.TransliterateFromBuckwalter(Str)
            Dim KeyArray(EndWordMultiOnly.Keys.Count - 1) As String
            EndWordMultiOnly.Keys.CopyTo(KeyArray, 0)
            Array.ForEach(KeyArray, Sub(S As String)
                                        If Str.LastIndexOf(S) <> -1 AndAlso Str.LastIndexOf(S) <> 0 Then EndWordMultiOnly.Remove(S)
                                    End Sub)
            ReDim KeyArray(StartWordMultiOnly.Keys.Count - 1)
            StartWordMultiOnly.Keys.CopyTo(KeyArray, 0)
            Array.ForEach(KeyArray, Sub(S As String)
                                        If Str.LastIndexOf(S) <> -1 AndAlso Str.LastIndexOf(S) <> 0 Then StartWordMultiOnly.Remove(S)
                                    End Sub)
            ReDim KeyArray(MiddleWordMultiOnly.Keys.Count - 1)
            MiddleWordMultiOnly.Keys.CopyTo(KeyArray, 0)
            Array.ForEach(KeyArray, Sub(S As String)
                                        If Str.LastIndexOf(S) <> -1 AndAlso Str.LastIndexOf(S) <> 0 Then MiddleWordMultiOnly.Remove(S)
                                    End Sub)
            For Count = 0 To Str.Length - 1
                Dim Index As Integer
                If Count = 0 Or Count = Str.Length - 1 Then
                    Index = MiddleWordOnly.IndexOf(Str.Chars(Count))
                    If Index <> -1 Then
                        MiddleWordOnly = MiddleWordOnly.Remove(Index, 1)
                    End If
                Else
                    Index = NotMiddleWord.IndexOf(Str.Chars(Count))
                    If Index <> -1 Then
                        NotMiddleWord = NotMiddleWord.Remove(Index, 1)
                    End If
                End If
                If Count <> 0 Then
                    Index = StartWordOnly.IndexOf(Str.Chars(Count))
                    If Index <> -1 Then
                        StartWordOnly = StartWordOnly.Remove(Index, 1)
                    End If
                Else
                    Index = NotStartWord.IndexOf(Str.Chars(Count))
                    If Index <> -1 Then
                        NotStartWord = NotStartWord.Remove(Index, 1)
                    End If
                End If
                If Count <> Str.Length - 1 Then
                    Index = EndWordOnly.IndexOf(Str.Chars(Count))
                    If Index <> -1 Then
                        EndWordOnly = EndWordOnly.Remove(Index, 1)
                    End If
                Else
                    Index = NotEndWord.IndexOf(Str.Chars(Count))
                    If Index <> -1 Then
                        NotEndWord = NotEndWord.Remove(Index, 1)
                    End If
                End If
                If Count <= Str.LastIndexOfAny(LtrSymbols.ToCharArray()) Then
                    If Count = 0 Or Count = Str.LastIndexOfAny(LtrSymbols.ToCharArray()) Then
                        Index = MiddleWordOnlyNoDia.IndexOf(Str.Chars(Count))
                        If Index <> -1 Then
                            MiddleWordOnlyNoDia = MiddleWordOnlyNoDia.Remove(Index, 1)
                        End If
                    Else
                        Index = NotMiddleWordNoDia.IndexOf(Str.Chars(Count))
                        If Index <> -1 Then
                            NotMiddleWordNoDia = NotMiddleWordNoDia.Remove(Index, 1)
                        End If
                    End If
                    If Count <> Str.LastIndexOfAny(LtrSymbols.ToCharArray()) Then
                        Index = EndWordOnlyNoDia.IndexOf(Str.Chars(Count))
                        If Index <> -1 Then
                            EndWordOnlyNoDia = EndWordOnlyNoDia.Remove(Index, 1)
                        End If
                    Else
                        Index = NotEndWordNoDia.IndexOf(Str.Chars(Count))
                        If Index <> -1 Then
                            NotEndWordNoDia = NotEndWordNoDia.Remove(Index, 1)
                        End If
                    End If
                End If
            Next
            Combos = Array.FindAll(Combos, Function(S As String) Not Str.Contains(S))
            DiaCombos = Array.FindAll(DiaCombos, Function(S As String) Not Str.Contains(S))
            LetCombos = Array.FindAll(LetCombos, Function(S As String) Not New String(Array.FindAll(Str.ToCharArray(), Function(C As Char) LtrSymbols.Contains(C))).Contains(S))
        Next
        Dim Dict As New Generic.Dictionary(Of Char, String)
        Array.ForEach(Combos, Sub(Str As String)
                                  If Dict.ContainsKey(Str.Chars(0)) Then
                                      Dict.Item(Str.Chars(0)) = Dict.Item(Str.Chars(0)) + Str.Chars(1)
                                  Else
                                      Dict.Add(Str.Chars(0), Str.Chars(1))
                                  End If
                              End Sub)
        Dim Val As String = Arabic.LeftToRightMark + "Combinations: "
        For Each Key As Char In Dict.Keys
            If Dict.Item(Key).Length > (DiaSymbols.Length + LtrSymbols.Length) / 2 Then
                Val += Arabic.FixStartingCombiningSymbol(Key) + Arabic.LeftToRightOverride + " [" + Arabic.PopDirectionalFormatting + String.Join(" ", Array.ConvertAll(New String(Array.FindAll((DiaSymbols + LtrSymbols).ToCharArray(), Function(C As Char) Not Dict.Item(Key).Contains(C))).ToCharArray(), Function(C As Char) Arabic.FixStartingCombiningSymbol(CStr(C)))) + Arabic.LeftToRightOverride + " ]" + Arabic.PopDirectionalFormatting + vbTab
            Else
                Val += Arabic.FixStartingCombiningSymbol(Key) + Arabic.LeftToRightOverride + " ! [ " + Arabic.PopDirectionalFormatting + String.Join(" ", Array.ConvertAll(Dict.Item(Key).ToCharArray(), Function(C As Char) Arabic.FixStartingCombiningSymbol(CStr(C)))) + Arabic.LeftToRightOverride + " ]" + Arabic.PopDirectionalFormatting + vbTab
            End If
        Next
        Dim RevDict As New Generic.Dictionary(Of Char, String)
        Array.ForEach(Combos, Sub(Str As String)
                                  If RevDict.ContainsKey(Str.Chars(1)) Then
                                      RevDict.Item(Str.Chars(1)) = RevDict.Item(Str.Chars(1)) + Str.Chars(0)
                                  Else
                                      RevDict.Add(Str.Chars(1), Str.Chars(0))
                                  End If
                              End Sub)
        Dim RevVal As String = Arabic.LeftToRightMark + "Reverse Combinations: "
        For Each Key As Char In RevDict.Keys
            If RevDict.Item(Key).Length > (DiaSymbols.Length + LtrSymbols.Length) / 2 Then
                RevVal += Arabic.LeftToRightOverride + "[" + Arabic.PopDirectionalFormatting + String.Join(" ", Array.ConvertAll(New String(Array.FindAll((DiaSymbols + LtrSymbols).ToCharArray(), Function(C As Char) Not RevDict.Item(Key).Contains(C))).ToCharArray(), Function(C As Char) Arabic.FixStartingCombiningSymbol(CStr(C)))) + Arabic.LeftToRightOverride + " ] " + Arabic.PopDirectionalFormatting + Arabic.FixStartingCombiningSymbol(Key) + vbTab
            Else
                RevVal += Arabic.LeftToRightOverride + "! [ " + Arabic.PopDirectionalFormatting + String.Join(" ", Array.ConvertAll(RevDict.Item(Key).ToCharArray(), Function(C As Char) Arabic.FixStartingCombiningSymbol(CStr(C)))) + Arabic.LeftToRightOverride + " ] " + Arabic.PopDirectionalFormatting + Arabic.FixStartingCombiningSymbol(Key) + vbTab
            End If
        Next
        Dim DiaDict As New Generic.Dictionary(Of Char, String)
        Array.ForEach(DiaCombos, Sub(Str As String)
                                     If DiaDict.ContainsKey(Str.Chars(0)) Then
                                         DiaDict.Item(Str.Chars(0)) = DiaDict.Item(Str.Chars(0)) + Str.Chars(1)
                                     Else
                                         DiaDict.Add(Str.Chars(0), Str.Chars(1))
                                     End If
                                 End Sub)
        Dim DiaVal As String = Arabic.LeftToRightMark + "Diacritic Only Combinations: "
        For Each Key As Char In DiaDict.Keys
            If DiaDict.Item(Key).Length > DiaSymbols.Length / 2 Then
                DiaVal += Arabic.FixStartingCombiningSymbol(Key) + Arabic.LeftToRightOverride + " [" + Arabic.PopDirectionalFormatting + String.Join(" ", Array.ConvertAll(New String(Array.FindAll(DiaSymbols.ToCharArray(), Function(C As Char) Not DiaDict.Item(Key).Contains(C))).ToCharArray(), Function(C As Char) Arabic.FixStartingCombiningSymbol(CStr(C)))) + Arabic.LeftToRightOverride + " ]" + Arabic.PopDirectionalFormatting + vbTab
            Else
                DiaVal += Arabic.FixStartingCombiningSymbol(Key) + Arabic.LeftToRightOverride + " ! [ " + Arabic.PopDirectionalFormatting + String.Join(" ", Array.ConvertAll(DiaDict.Item(Key).ToCharArray(), Function(C As Char) Arabic.FixStartingCombiningSymbol(CStr(C)))) + Arabic.LeftToRightOverride + " ]" + Arabic.PopDirectionalFormatting + vbTab
            End If
        Next
        Dim LetDict As New Generic.Dictionary(Of Char, String)
        Array.ForEach(LetCombos, Sub(Str As String)
                                     If LetDict.ContainsKey(Str.Chars(0)) Then
                                         LetDict.Item(Str.Chars(0)) = LetDict.Item(Str.Chars(0)) + Str.Chars(1)
                                     Else
                                         LetDict.Add(Str.Chars(0), Str.Chars(1))
                                     End If
                                 End Sub)
        Dim LetVal As String = Arabic.LeftToRightMark + "Letter Only Combinations: "
        For Each Key As Char In LetDict.Keys
            If LetDict.Item(Key).Length > LtrSymbols.Length / 2 Then
                LetVal += Arabic.FixStartingCombiningSymbol(Key) + Arabic.LeftToRightOverride + " [" + Arabic.PopDirectionalFormatting + String.Join(" ", Array.ConvertAll(New String(Array.FindAll(LtrSymbols.ToCharArray(), Function(C As Char) Not LetDict.Item(Key).Contains(C))).ToCharArray(), Function(C As Char) Arabic.FixStartingCombiningSymbol(CStr(C)))) + Arabic.LeftToRightOverride + " ]" + Arabic.PopDirectionalFormatting + vbTab
            Else
                LetVal += Arabic.FixStartingCombiningSymbol(Key) + Arabic.LeftToRightOverride + " ! [ " + Arabic.PopDirectionalFormatting + String.Join(" ", Array.ConvertAll(LetDict.Item(Key).ToCharArray(), Function(C As Char) Arabic.FixStartingCombiningSymbol(CStr(C)))) + Arabic.LeftToRightOverride + " ]" + Arabic.PopDirectionalFormatting + vbTab
            End If
        Next
        Dim LetRevDict As New Generic.Dictionary(Of Char, String)
        Array.ForEach(LetCombos, Sub(Str As String)
                                     If LetRevDict.ContainsKey(Str.Chars(1)) Then
                                         LetRevDict.Item(Str.Chars(1)) = LetRevDict.Item(Str.Chars(1)) + Str.Chars(0)
                                     Else
                                         LetRevDict.Add(Str.Chars(1), Str.Chars(0))
                                     End If
                                 End Sub)
        Dim LetRevVal As String = Arabic.LeftToRightMark + "Reverse Letter Only Combinations: "
        For Each Key As Char In LetRevDict.Keys
            If LetRevDict.Item(Key).Length > LtrSymbols.Length / 2 Then
                LetRevVal += Arabic.LeftToRightOverride + "[" + Arabic.PopDirectionalFormatting + String.Join(" ", Array.ConvertAll(New String(Array.FindAll(LtrSymbols.ToCharArray(), Function(C As Char) Not LetRevDict.Item(Key).Contains(C))).ToCharArray(), Function(C As Char) Arabic.FixStartingCombiningSymbol(CStr(C)))) + Arabic.LeftToRightOverride + " ] " + Arabic.PopDirectionalFormatting + Arabic.FixStartingCombiningSymbol(Key) + vbTab
            Else
                LetRevVal += Arabic.LeftToRightOverride + "! [ " + Arabic.PopDirectionalFormatting + String.Join(" ", Array.ConvertAll(LetRevDict.Item(Key).ToCharArray(), Function(C As Char) Arabic.FixStartingCombiningSymbol(CStr(C)))) + Arabic.LeftToRightOverride + " ] " + Arabic.PopDirectionalFormatting + Arabic.FixStartingCombiningSymbol(Key) + vbTab
            End If
        Next
        Dim StartMulti As String = " "
        For Each Key As String In StartWordMultiOnly.Keys
            StartMulti += Key + " "
        Next
        Dim EndMulti As String = " "
        For Each Key As String In EndWordMultiOnly.Keys
            EndMulti += Key + " "
        Next
        Dim MiddleMulti As String = " "
        For Each Key As String In MiddleWordMultiOnly.Keys
            MiddleMulti += Key + " "
        Next
        Return {Arabic.LeftToRightMark + "Unique Prefix: [" + StartMulti + Arabic.LeftToRightOverride + "]" + Arabic.PopDirectionalFormatting, _
                Arabic.LeftToRightMark + "Unique Suffix: [" + EndMulti + Arabic.LeftToRightOverride + "]" + Arabic.PopDirectionalFormatting, _
                Arabic.LeftToRightMark + "Unique Middle: [" + MiddleMulti + Arabic.LeftToRightOverride + "]" + Arabic.PopDirectionalFormatting, _
                Arabic.LeftToRightMark + "Start Only: [" + String.Join(" ", Array.ConvertAll(StartWordOnly.ToCharArray(), Function(C As Char) Arabic.FixStartingCombiningSymbol(CStr(C)))) + Arabic.LeftToRightOverride + "]" + Arabic.PopDirectionalFormatting, _
                Arabic.LeftToRightMark + "Not Start: [" + String.Join(" ", Array.ConvertAll(NotStartWord.ToCharArray(), Function(C As Char) Arabic.FixStartingCombiningSymbol(CStr(C)))) + Arabic.LeftToRightOverride + "]" + Arabic.PopDirectionalFormatting, _
                Arabic.LeftToRightMark + "End Only: [" + String.Join(" ", Array.ConvertAll(EndWordOnly.ToCharArray(), Function(C As Char) Arabic.FixStartingCombiningSymbol(CStr(C)))) + Arabic.LeftToRightOverride + "]" + Arabic.PopDirectionalFormatting, _
                Arabic.LeftToRightMark + "Not End: [" + String.Join(" ", Array.ConvertAll(NotEndWord.ToCharArray(), Function(C As Char) Arabic.FixStartingCombiningSymbol(CStr(C)))) + Arabic.LeftToRightOverride + "]" + Arabic.PopDirectionalFormatting, _
                Arabic.LeftToRightMark + "End Only No Diacritics: [" + String.Join(" ", Array.ConvertAll(EndWordOnlyNoDia.ToCharArray(), Function(C As Char) Arabic.FixStartingCombiningSymbol(CStr(C)))) + Arabic.LeftToRightOverride + "]" + Arabic.PopDirectionalFormatting, _
                Arabic.LeftToRightMark + "Not End No Diacritics: [" + String.Join(" ", Array.ConvertAll(NotEndWordNoDia.ToCharArray(), Function(C As Char) Arabic.FixStartingCombiningSymbol(CStr(C)))) + Arabic.LeftToRightOverride + "]" + Arabic.PopDirectionalFormatting, _
                Arabic.LeftToRightMark + "Middle Only: [" + String.Join(" ", Array.ConvertAll(MiddleWordOnly.ToCharArray(), Function(C As Char) Arabic.FixStartingCombiningSymbol(CStr(C)))) + Arabic.LeftToRightOverride + "]" + Arabic.PopDirectionalFormatting, _
                Arabic.LeftToRightMark + "Not Middle: [" + String.Join(" ", Array.ConvertAll(NotMiddleWord.ToCharArray(), Function(C As Char) Arabic.FixStartingCombiningSymbol(CStr(C)))) + Arabic.LeftToRightOverride + "]" + Arabic.PopDirectionalFormatting, _
                Arabic.LeftToRightMark + "Middle Only No Diacritics: [" + String.Join(" ", Array.ConvertAll(MiddleWordOnlyNoDia.ToCharArray(), Function(C As Char) Arabic.FixStartingCombiningSymbol(CStr(C)))) + Arabic.LeftToRightOverride + "]" + Arabic.PopDirectionalFormatting, _
                Arabic.LeftToRightMark + "Not Middle No Diacritics: [" + String.Join(" ", Array.ConvertAll(NotMiddleWordNoDia.ToCharArray(), Function(C As Char) Arabic.FixStartingCombiningSymbol(CStr(C)))) + Arabic.LeftToRightOverride + "]" + Arabic.PopDirectionalFormatting, _
                Val, RevVal, DiaVal, LetVal, LetRevVal}
    End Function
    Public Shared Function GetSelectionNames() As Array()
        Dim Division As Integer = 0
        Dim Strings As String = HttpContext.Current.Request.QueryString.Get("qurandivision")
        If Not Strings Is Nothing Then Division = CInt(Strings)
        If Division = 0 Then
            Return TanzilReader.GetChapterNames()
        ElseIf Division = 1 Then
            Return TanzilReader.GetChapterNamesByRevelationOrder()
        ElseIf Division = 2 Then
            Return TanzilReader.GetPartNames()
        ElseIf Division = 3 Then
            Return TanzilReader.GetGroupNames()
        ElseIf Division = 4 Then
            Return TanzilReader.GetStationNames()
        ElseIf Division = 5 Then
            Return TanzilReader.GetSectionNames()
        ElseIf Division = 6 Then
            Return TanzilReader.GetPageNames()
        ElseIf Division = 7 Then
            Return TanzilReader.GetSajdaNames()
        ElseIf Division = 8 Then
            Return TanzilReader.GetImportantNames()
        ElseIf Division = 9 Then
            Return Arabic.GetRecitationSymbols()
        End If
        Return Nothing
    End Function
    Enum QuranScripts
        Uthmani = 0
        UthmaniMin = 1
        Simple = 2
        SimpleMin = 3
        SimpleEnhanced = 4
        SimpleClean = 5
        BuckwalterUthmani = 6
        BuckwalterUthmaniMin = 7
        BuckwalterSimple = 8
        BuckwalterSimpleMin = 9
        BuckwalterSimpleEnhanced = 10
        BuckwalterSimpleClean = 11
        Warsh = 12
        AlDari = 13
    End Enum
    Shared QuranFileNames As String() = {"quran-uthmani.xml", "quran-uthmani-min.xml", "quran-simple.xml", "quran-simple-min.xml", "quran-simple-enhanced.xml", "quran-simple-clean.xml", "quran-buckwalter-uthmani.xml", "quran-buckwalter-uthmani-min.xml", "quran-buckwlater-simple.xml", "quran-buckwalter-simple-min.xml", "quran-buckwalter-simple-enhanced.xml", "quran-buckwalter-simple-clean.xml", "quran-warsh.xml", "quran-alduri.xml"}
    Shared QuranScriptNames As String() = {"Uthmani", "Uthmani Minimal", "Simple", "Simple Minimal", "Simple Enhanced", "Simple Clean"}
    Public Shared Sub CheckNotablePatterns()
        'ComparePatterns(QuranScripts.Uthmani, QuranScripts.UthmaniMin, Arabic.UthmaniShortVowelsBeforeLongVowelsAlef)
        'ComparePatterns(QuranScripts.Uthmani, QuranScripts.UthmaniMin, Arabic.UthmaniShortVowelsBeforeLongVowelsSuperscriptAlef)
        'ComparePatterns(QuranScripts.Uthmani, QuranScripts.UthmaniMin, Arabic.UthmaniShortVowelsBeforeLongVowelsAlefAlefMaksura)
        'ComparePatterns(QuranScripts.Uthmani, QuranScripts.UthmaniMin, Arabic.UthmaniShortVowelsBeforeLongVowelsYehAlefMaksura)
        'ComparePatterns(QuranScripts.Uthmani, QuranScripts.UthmaniMin, Arabic.UthmaniShortVowelsBeforeLongVowelsYeh)
        'ComparePatterns(QuranScripts.Uthmani, QuranScripts.UthmaniMin, Arabic.UthmaniShortVowelsBeforeLongVowelsSmallYeh)
        'ComparePatterns(QuranScripts.Uthmani, QuranScripts.UthmaniMin, Arabic.UthmaniShortVowelsBeforeLongVowelsWaw)
        'ComparePatterns(QuranScripts.Uthmani, QuranScripts.UthmaniMin, Arabic.UthmaniShortVowelsBeforeLongVowelsSmallWaw)
        ComparePatterns(QuranScripts.Uthmani, QuranScripts.SimpleEnhanced, Arabic.SimpleTrailingAlef)
        'this rule should be analyzed after all other rules in Simple Enhanced are processed as it will great simplify its expression while the earlier it is processed the longer it will be
        'ComparePatterns(QuranScripts.Uthmani, QuranScripts.SimpleEnhanced, Arabic.SimpleSuperscriptAlef)
    End Sub
    Public Shared Sub RemoveSubsetPatterns(ByRef Dict As Dictionary(Of String, String), Prefix As Boolean)
        Dim Keys(Dict.Keys.Count - 1) As String
        Dict.Keys.CopyTo(Keys, 0)
        For Count As Integer = 0 To Keys.Length - 1
            For SubCount As Integer = 1 To Keys(Count).Length - 1
                If Dict.ContainsKey(If(Prefix, Keys(Count).Substring(SubCount), Keys(Count).Substring(0, SubCount))) Then
                    Dict.Remove(Keys(Count))
                    Exit For
                End If
            Next
        Next
    End Sub
    Public Shared Function DumpDictionary(Dict As Dictionary(Of String, String)) As String
        Dim Msg As String = String.Empty
        Dim Keys(Dict.Keys.Count - 1) As String
        Dict.Keys.CopyTo(Keys, 0)
        Array.Sort(Keys, StringComparer.Ordinal)
        For Count As Integer = 0 To Keys.Length - 1
            Msg += """" + Arabic.TransliterateToScheme(Keys(Count), Arabic.TranslitScheme.Literal, String.Empty) + """" + If(Count <> Keys.Length - 1, ", ", String.Empty)
        Next
        Return Msg
    End Function
    Public Shared Sub ComparePatterns(ScriptType As QuranScripts, CompScriptType As QuranScripts, LetterPattern As String)
        Dim WordPattern As String = "(?<=^\s*|\s+)\S*" + Arabic.MakeUniRegEx(LetterPattern) + "\S*(?=\s+|\s*$)"
        Dim FirstList As List(Of String) = PatternMatch(ScriptType, WordPattern)
        FirstList.Sort(StringComparer.Ordinal)
        Dim CompList As List(Of String) = PatternMatch(CompScriptType, "(?<=^\s*|\s+)\S*" + Arabic.MakeUniRegEx(LetterPattern.Substring(0, 1)) + "(?=\s+|\s*$)")
        CompList.Sort(StringComparer.Ordinal)
        Dim Index As Integer = 0
        Do While Index < CompList.Count - 1
            Do While Index < CompList.Count - 1 AndAlso CompList(Index + 1) = CompList(Index)
                CompList.RemoveAt(Index)
            Loop
            Index += 1
        Loop
        Index = 0
        Do While Index <= FirstList.Count - 1
            Do While Index < FirstList.Count - 1 AndAlso FirstList(Index + 1) = FirstList(Index)
                FirstList.RemoveAt(Index)
            Loop
            Dim Find As Integer = CompList.BinarySearch(FirstList(Index))
            If Find >= 0 Then
                Do
                    CompList.RemoveAt(Find)
                Loop While Find <= CompList.Count - 1 AndAlso CompList(Find) = FirstList(Index)
                Find -= 1
                While Find <> -1 AndAlso CompList(Find) = FirstList(Index)
                    CompList.RemoveAt(Find)
                    Find -= 1
                End While
            End If
            Index += 1
        Loop
        Dim FirstDict As New Dictionary(Of String, String)
        Dim CompDict As New Dictionary(Of String, String)
        Dim Msg As String = "First: "
        For Each Str As String In FirstList
            Dim SubKey As String = Str.Substring(0, Str.IndexOf(LetterPattern))
            For Count As Integer = 0 To SubKey.Length - 1
                If Not FirstDict.ContainsKey(SubKey.Substring(Count)) Then
                    FirstDict.Add(SubKey.Substring(Count), Str)
                End If
            Next
            Msg += """" + Arabic.TransliterateToScheme(Str, Arabic.TranslitScheme.Literal, String.Empty) + """, "
        Next
        Msg += vbCrLf + "Second: "
        For Each Str As String In CompList
            Dim SubKey As String = Str.Substring(0, Str.IndexOf(LetterPattern))
            For Count As Integer = 0 To SubKey.Length - 1
                If Not CompDict.ContainsKey(SubKey.Substring(Count)) Then
                    CompDict.Add(SubKey.Substring(Count), Str)
                End If
            Next
            Msg += """" + Arabic.TransliterateToScheme(Str, Arabic.TranslitScheme.Literal, String.Empty) + """, "
        Next
        Dim Keys(FirstDict.Keys.Count - 1) As String
        Dim FirstNotInDict As New Dictionary(Of String, String)
        Dim CompNotInDict As New Dictionary(Of String, String)
        FirstDict.Keys.CopyTo(Keys, 0)
        For Count As Integer = 0 To Keys.Length - 1
            If CompDict.ContainsKey(Keys(Count)) Then
                If Not FirstNotInDict.ContainsKey(FirstDict(Keys(Count))) Then
                    FirstNotInDict.Add(Keys(Count), FirstDict(Keys(Count)))
                End If
                If Not CompNotInDict.ContainsKey(CompDict(Keys(Count))) Then
                    CompNotInDict.Add(Keys(Count), CompDict(Keys(Count)))
                End If
                FirstDict.Remove(Keys(Count))
                CompDict.Remove(Keys(Count))
            End If
        Next
        RemoveSubsetPatterns(FirstNotInDict, True)
        RemoveSubsetPatterns(CompNotInDict, True)
        RemoveSubsetPatterns(FirstDict, True)
        RemoveSubsetPatterns(CompDict, True)
        Msg += vbCrLf + "First: " + DumpDictionary(FirstDict) + vbCrLf + "Not First: " + DumpDictionary(FirstNotInDict) + vbCrLf + "Second: " + DumpDictionary(CompDict) + vbCrLf + "Not Second: " + DumpDictionary(CompNotInDict)
        FirstDict = New Dictionary(Of String, String)
        CompDict = New Dictionary(Of String, String)
        For Each Str As String In FirstList
            Dim SubKey As String = Str.Substring(Str.IndexOf(LetterPattern) + 1)
            For Count As Integer = 1 To SubKey.Length
                If Not FirstDict.ContainsKey(SubKey.Substring(0, Count)) Then
                    FirstDict.Add(SubKey.Substring(0, Count), Str)
                End If
            Next
        Next
        For Each Str As String In CompList
            Dim SubKey As String = Str.Substring(Str.IndexOf(LetterPattern) + 1)
            For Count As Integer = 1 To SubKey.Length
                If Not CompDict.ContainsKey(SubKey.Substring(0, Count)) Then
                    CompDict.Add(SubKey.Substring(0, Count), Str)
                End If
            Next
        Next
        FirstNotInDict = New Dictionary(Of String, String)
        CompNotInDict = New Dictionary(Of String, String)
        ReDim Keys(FirstDict.Keys.Count - 1)
        FirstDict.Keys.CopyTo(Keys, 0)
        For Count As Integer = 0 To Keys.Length - 1
            If CompDict.ContainsKey(Keys(Count)) Then
                If Not FirstNotInDict.ContainsKey(FirstDict(Keys(Count))) Then
                    FirstNotInDict.Add(Keys(Count), FirstDict(Keys(Count)))
                End If
                If Not CompNotInDict.ContainsKey(CompDict(Keys(Count))) Then
                    CompNotInDict.Add(Keys(Count), CompDict(Keys(Count)))
                End If
                FirstDict.Remove(Keys(Count))
                CompDict.Remove(Keys(Count))
            End If
        Next
        RemoveSubsetPatterns(FirstNotInDict, False)
        RemoveSubsetPatterns(CompNotInDict, False)
        RemoveSubsetPatterns(FirstDict, False)
        RemoveSubsetPatterns(CompDict, False)
        Msg += vbCrLf + "First: " + DumpDictionary(FirstDict) + vbCrLf + "Not First: " + DumpDictionary(FirstNotInDict) + vbCrLf + "Second: " + DumpDictionary(CompDict) + vbCrLf + "Not Second: " + DumpDictionary(CompNotInDict)
        Debug.Print(Msg)
    End Sub
    Public Shared Function PatternMatch(ScriptType As QuranScripts, Pattern As String) As List(Of String)
        PatternMatch = New List(Of String)
        Dim Doc As New System.Xml.XmlDocument
        If ScriptType = QuranScripts.Uthmani Then
            Doc.Load(Utility.GetFilePath("metadata\quran-uthmani.xml"))
        Else
            Doc.Load(Utility.GetFilePath("IslamMetadata\" + QuranFileNames(ScriptType)))
        End If
        Dim Verses As Collections.Generic.List(Of String())
        Verses = TanzilReader.GetQuranText(Doc, -1, -1, -1, -1)
        For Count As Integer = 0 To Verses.Count - 1
            Dim ChapterNode As System.Xml.XmlNode = GetTextChapter(Doc, Count + 1)
            For SubCount As Integer = 0 To Verses(Count).Length - 1
                If SubCount = 0 AndAlso Not GetTextVerse(ChapterNode, SubCount + 1).Attributes.GetNamedItem("bismillah") Is Nothing Then
                    For Each Val As System.Text.RegularExpressions.Match In System.Text.RegularExpressions.Regex.Matches(GetTextVerse(ChapterNode, SubCount + 1).Attributes.GetNamedItem("bismillah").Value, Pattern)
                        PatternMatch.Add(Val.Value)
                    Next
                End If
                For Each Val As System.Text.RegularExpressions.Match In System.Text.RegularExpressions.Regex.Matches(Verses(Count)(SubCount), Pattern)
                    PatternMatch.Add(Val.Value)
                Next
            Next
        Next
    End Function
    Public Shared Sub ChangeQuranFormat(ScriptType As QuranScripts)
        Dim Doc As New System.Xml.XmlDocument
        Doc.Load(Utility.GetFilePath("metadata\quran-uthmani.xml"))
        Dim Verses As Collections.Generic.List(Of String())
        Dim UseBuckwalter As Boolean = False
        Dim Path As String = Utility.GetFilePath("metadata\-" + QuranFileNames(ScriptType))
        If ScriptType = QuranScripts.BuckwalterUthmani Then
            UseBuckwalter = True
            ScriptType = QuranScripts.Uthmani
        ElseIf ScriptType = QuranScripts.BuckwalterUthmaniMin Then
            UseBuckwalter = True
            ScriptType = QuranScripts.UthmaniMin
        ElseIf ScriptType = QuranScripts.BuckwalterSimple Then
            UseBuckwalter = True
            ScriptType = QuranScripts.Simple
        ElseIf ScriptType = QuranScripts.BuckwalterSimpleEnhanced Then
            UseBuckwalter = True
            ScriptType = QuranScripts.SimpleEnhanced
        ElseIf ScriptType = QuranScripts.BuckwalterSimpleMin Then
            UseBuckwalter = True
            ScriptType = QuranScripts.SimpleMin
        ElseIf ScriptType = QuranScripts.BuckwalterSimpleClean Then
            UseBuckwalter = True
            ScriptType = QuranScripts.SimpleClean
        End If
        Doc.DocumentElement.PreviousSibling.Value = Doc.DocumentElement.PreviousSibling.Value.Replace("Uthmani", QuranScriptNames(ScriptType))
        Verses = TanzilReader.GetQuranText(CachedData.XMLDocMain, -1, -1, -1, -1)
        For Count As Integer = 0 To Verses.Count - 1
            Dim ChapterNode As System.Xml.XmlNode = GetTextChapter(Doc, Count + 1)
            If UseBuckwalter Then
                ChapterNode.Attributes.GetNamedItem("name").Value = Arabic.TransliterateToScheme(ChapterNode.Attributes.GetNamedItem("name").Value, Arabic.TranslitScheme.Literal, String.Empty)
            End If
            For SubCount As Integer = 0 To Verses(Count).Length - 1
                If SubCount = 0 AndAlso Not GetTextVerse(ChapterNode, SubCount + 1).Attributes.GetNamedItem("bismillah") Is Nothing Then
                    GetTextVerse(ChapterNode, SubCount + 1).Attributes.GetNamedItem("bismillah").Value = Arabic.ChangeScript(GetTextVerse(ChapterNode, SubCount + 1).Attributes.GetNamedItem("bismillah").Value, ScriptType)
                    If UseBuckwalter Then
                        GetTextVerse(ChapterNode, SubCount + 1).Attributes.GetNamedItem("bismillah").Value = Arabic.TransliterateToScheme(GetTextVerse(ChapterNode, SubCount + 1).Attributes.GetNamedItem("bismillah").Value, Arabic.TranslitScheme.Literal, String.Empty)
                    End If
                End If
                GetTextVerse(ChapterNode, SubCount + 1).Attributes.GetNamedItem("text").Value = Arabic.ChangeScript(Verses(Count)(SubCount), ScriptType)
                If UseBuckwalter Then
                    GetTextVerse(ChapterNode, SubCount + 1).Attributes.GetNamedItem("text").Value = Arabic.TransliterateToScheme(GetTextVerse(ChapterNode, SubCount + 1).Attributes.GetNamedItem("text").Value, Arabic.TranslitScheme.Literal, String.Empty)
                End If
            Next
        Next
        Doc.Save(Path)
    End Sub
    Public Shared Function IsQuranTextReference(Str As String) As Boolean
        Return System.Text.RegularExpressions.Regex.Match(Str, "^(?:,?(\d+)(?:\:(\d+))?(?:\:(\d+))?(?:-(\d+)(?:\:(\d+))?(?:\:(\d+))?)?)+$").Success
    End Function
    Public Shared Function QuranTextFromReference(Str As String) As RenderArray
        Dim Renderer As New RenderArray
        Dim Matches As System.Text.RegularExpressions.MatchCollection = System.Text.RegularExpressions.Regex.Matches(Str, "(?:,?(\d+)(?:\:(\d+))?(?:\:(\d+))?(?:-(\d+)(?:\:(\d+))?(?:\:(\d+))?)?)")
        Dim SchemeType As Arabic.TranslitScheme = If(CInt(HttpContext.Current.Request.QueryString.Get("translitscheme")) >= 2, 2 + CInt(HttpContext.Current.Request.QueryString.Get("translitscheme")) Mod 2, CInt(HttpContext.Current.Request.QueryString.Get("translitscheme")))
        Dim Scheme As String = If(CInt(HttpContext.Current.Request.QueryString.Get("translitscheme")) >= 2, CachedData.IslamData.TranslitSchemes((CInt(HttpContext.Current.Request.QueryString.Get("translitscheme")) - 2) \ 2).Name, String.Empty)
        Dim TranslationIndex As Integer = GetTranslationIndex(HttpContext.Current.Request.QueryString.Get("qurantranslation"))
        Dim Reference As String
        For Count = 0 To Matches.Count - 1
            Dim BaseChapter As Integer = Matches(Count).Groups(1).Value
            Dim BaseVerse As Integer = If(Matches(Count).Groups(2).Value = String.Empty, 0, CInt(Matches(Count).Groups(2).Value))
            Dim WordNumber As Integer = If(Matches(Count).Groups(3).Value = String.Empty, 0, CInt(Matches(Count).Groups(3).Value))
            Dim EndChapter As Integer = If(Matches(Count).Groups(4).Value = String.Empty, 0, CInt(Matches(Count).Groups(4).Value))
            Dim ExtraVerseNumber As Integer = If(Matches(Count).Groups(5).Value = String.Empty, 0, CInt(Matches(Count).Groups(5).Value))
            Dim EndWordNumber As Integer = If(Matches(Count).Groups(6).Value = String.Empty, 0, CInt(Matches(Count).Groups(6).Value))
            If BaseVerse <> 0 And WordNumber = 0 And EndChapter <> 0 And ExtraVerseNumber = 0 And EndWordNumber = 0 Then
                ExtraVerseNumber = EndChapter
                EndChapter = 0
            ElseIf BaseVerse <> 0 And WordNumber <> 0 And EndChapter <> 0 And ExtraVerseNumber = 0 And EndWordNumber = 0 Then
                EndWordNumber = EndChapter
                EndChapter = 0
            ElseIf BaseVerse <> 0 And WordNumber <> 0 And EndChapter <> 0 And ExtraVerseNumber <> 0 And EndWordNumber = 0 Then
                EndWordNumber = ExtraVerseNumber
                ExtraVerseNumber = EndChapter
                EndChapter = 0
            End If
            Renderer.Items.AddRange(DoGetRenderedQuranText(QuranTextRangeLookup(BaseChapter, BaseVerse, WordNumber, EndChapter, ExtraVerseNumber, EndWordNumber), BaseChapter, BaseVerse, HttpContext.Current.Request.QueryString.Get("qurantranslation"), SchemeType, Scheme, TranslationIndex).Items)
            Reference = CStr(BaseChapter) + If(BaseVerse <> 0, ":" + CStr(BaseVerse), String.Empty) + If(EndChapter <> 0, "-" + CStr(EndChapter) + If(ExtraVerseNumber <> 0, ":" + CStr(ExtraVerseNumber), String.Empty), If(ExtraVerseNumber <> 0, "-" + CStr(ExtraVerseNumber), String.Empty))
            Renderer.Items.Add(New RenderArray.RenderItem(RenderArray.RenderTypes.eHeaderCenter, New RenderArray.RenderText() {New RenderArray.RenderText(RenderArray.RenderDisplayClass.eLTR, "(Qur'an " + Reference + ")")}))
        Next
        Return Renderer
    End Function
    Public Shared Function QuranTextRangeLookup(BaseChapter As Integer, BaseVerse As Integer, WordNumber As Integer, EndChapter As Integer, ExtraVerseNumber As Integer, EndWordNumber As Integer) As Collections.Generic.List(Of String())
        Dim QuranText As New Collections.Generic.List(Of String())
        If EndChapter = 0 Or EndChapter = BaseChapter Then
            QuranText.Add(TanzilReader.GetQuranText(CachedData.XMLDocMain, BaseChapter, BaseVerse, CInt(IIf(ExtraVerseNumber <> 0, ExtraVerseNumber, BaseVerse))))
        Else
            QuranText.AddRange(TanzilReader.GetQuranText(CachedData.XMLDocMain, BaseChapter, BaseVerse, EndChapter, ExtraVerseNumber))
        End If
        If (WordNumber > 1) Then
            Dim VerseIndex As Integer = 0
            For WordCount As Integer = 1 To WordNumber - 1
                VerseIndex = QuranText(0)(0).IndexOf(" "c, VerseIndex) + 1
            Next
            QuranText(0)(0) = System.Text.RegularExpressions.Regex.Replace(QuranText(0)(0).Substring(0, VerseIndex), "(^\s*|\s+)[^\s" + String.Join(String.Empty, Array.ConvertAll(Arabic.ArabicStopLetters, Function(Str As String) Arabic.MakeUniRegEx(Str))) + "]+(?=\s*$|\s+)", "$1") + QuranText(0)(0).Substring(VerseIndex)
        End If
        If (EndWordNumber <> 0) Then
            Dim VerseIndex As Integer = 0
            'selections are always within the same chapter
            Dim LastChapter As Integer = QuranText.Count - 1
            Dim LastVerse As Integer = CInt(IIf(ExtraVerseNumber <> 0, QuranText(LastChapter).Length - 1, 0))
            For WordCount As Integer = 1 To EndWordNumber - 1
                VerseIndex = QuranText(LastChapter)(LastVerse).IndexOf(" "c, VerseIndex) + 1
            Next
            QuranText(LastChapter)(LastVerse) = QuranText(LastChapter)(LastVerse).Substring(0, VerseIndex) + System.Text.RegularExpressions.Regex.Replace(QuranText(LastChapter)(LastVerse).Substring(VerseIndex), "(^\s*|\s+)[^\s" + String.Join(String.Empty, Array.ConvertAll(Arabic.ArabicStopLetters, Function(Str As String) Arabic.MakeUniRegEx(Str))) + "]+(?=\s*$|\s+)", "$1")
        End If
        Return QuranText
    End Function
    Public Shared Function GetQuranTextBySelection(Division As Integer, Index As Integer, Translation As String, SchemeType As Arabic.TranslitScheme, Scheme As String, TranslationIndex As Integer) As RenderArray
        Dim Chapter As Integer
        Dim Verse As Integer
        Dim BaseChapter As Integer
        Dim BaseVerse As Integer
        Dim Node As System.Xml.XmlNode
        Dim Renderer As New RenderArray
        Dim QuranText As Collections.Generic.List(Of String())
        Dim SeperateSectionCount As Integer = 1
        Dim Keys() As String = Nothing
        If Division = 8 Then SeperateSectionCount = CachedData.IslamData.QuranSelections(Index).SelectionInfo.Length
        If Division = 9 Then
            ReDim Keys(CachedData.LetterDictionary(CachedData.IslamData.ArabicLetters(Index).Symbol).Count - 1)
            CachedData.LetterDictionary(CachedData.IslamData.ArabicLetters(Index).Symbol).Keys.CopyTo(Keys, 0)
            SeperateSectionCount = CachedData.LetterDictionary(CachedData.IslamData.ArabicLetters(Index).Symbol).Count
        End If
        For SectionCount As Integer = 0 To SeperateSectionCount - 1
            If Division = 0 Then
                BaseChapter = CInt(GetChapterByIndex(Index).Attributes.GetNamedItem("index").Value)
                BaseVerse = 1
                QuranText = New Collections.Generic.List(Of String())
                QuranText.Add(GetQuranText(CachedData.XMLDocMain, BaseChapter, BaseVerse, -1))
            ElseIf Division = 1 Then
                BaseChapter = CInt(TanzilReader.GetChapterIndexByRevelationOrder(Index).Attributes.GetNamedItem("index").Value)
                BaseVerse = 1
                QuranText = New Collections.Generic.List(Of String())
                QuranText.Add(GetQuranText(CachedData.XMLDocMain, BaseChapter, BaseVerse, -1))
            ElseIf Division = 2 Then
                Node = TanzilReader.GetPartByIndex(Index)
                BaseChapter = CInt(Node.Attributes.GetNamedItem("sura").Value)
                BaseVerse = CInt(Node.Attributes.GetNamedItem("aya").Value)
                Node = TanzilReader.GetPartByIndex(Index + 1)
                If Node Is Nothing Then
                    Chapter = -1
                    Verse = -1
                Else
                    Chapter = CInt(Node.Attributes.GetNamedItem("sura").Value)
                    Verse = CInt(Node.Attributes.GetNamedItem("aya").Value)
                    TanzilReader.GetPreviousChapterVerse(Chapter, Verse)
                End If
                QuranText = GetQuranText(CachedData.XMLDocMain, BaseChapter, BaseVerse, Chapter, Verse)
            ElseIf Division = 3 Then
                Node = TanzilReader.GetGroupByIndex(Index)
                BaseChapter = CInt(Node.Attributes.GetNamedItem("sura").Value)
                BaseVerse = CInt(Node.Attributes.GetNamedItem("aya").Value)
                Node = TanzilReader.GetGroupByIndex(Index + 1)
                If Node Is Nothing Then
                    Chapter = -1
                    Verse = -1
                Else
                    Chapter = CInt(Node.Attributes.GetNamedItem("sura").Value)
                    Verse = CInt(Node.Attributes.GetNamedItem("aya").Value)
                    TanzilReader.GetPreviousChapterVerse(Chapter, Verse)
                End If
                QuranText = GetQuranText(CachedData.XMLDocMain, BaseChapter, BaseVerse, Chapter, Verse)
            ElseIf Division = 4 Then
                Node = TanzilReader.GetStationByIndex(Index)
                BaseChapter = CInt(Node.Attributes.GetNamedItem("sura").Value)
                BaseVerse = CInt(Node.Attributes.GetNamedItem("aya").Value)
                Node = TanzilReader.GetStationByIndex(Index + 1)
                If Node Is Nothing Then
                    Chapter = -1
                    Verse = -1
                Else
                    Chapter = CInt(Node.Attributes.GetNamedItem("sura").Value)
                    Verse = CInt(Node.Attributes.GetNamedItem("aya").Value)
                    TanzilReader.GetPreviousChapterVerse(Chapter, Verse)
                End If
                QuranText = GetQuranText(CachedData.XMLDocMain, BaseChapter, BaseVerse, Chapter, Verse)
            ElseIf Division = 5 Then
                Node = TanzilReader.GetSectionByIndex(Index)
                BaseChapter = CInt(Node.Attributes.GetNamedItem("sura").Value)
                BaseVerse = CInt(Node.Attributes.GetNamedItem("aya").Value)
                Node = TanzilReader.GetSectionByIndex(Index + 1)
                If Node Is Nothing Then
                    Chapter = -1
                    Verse = -1
                Else
                    Chapter = CInt(Node.Attributes.GetNamedItem("sura").Value)
                    Verse = CInt(Node.Attributes.GetNamedItem("aya").Value)
                    TanzilReader.GetPreviousChapterVerse(Chapter, Verse)
                End If
                QuranText = GetQuranText(CachedData.XMLDocMain, BaseChapter, BaseVerse, Chapter, Verse)
            ElseIf Division = 6 Then
                Node = TanzilReader.GetPageByIndex(Index)
                BaseChapter = CInt(Node.Attributes.GetNamedItem("sura").Value)
                BaseVerse = CInt(Node.Attributes.GetNamedItem("aya").Value)
                Node = TanzilReader.GetPageByIndex(Index + 1)
                If Node Is Nothing Then
                    Chapter = -1
                    Verse = -1
                Else
                    Chapter = CInt(Node.Attributes.GetNamedItem("sura").Value)
                    Verse = CInt(Node.Attributes.GetNamedItem("aya").Value)
                    TanzilReader.GetPreviousChapterVerse(Chapter, Verse)
                End If
                QuranText = GetQuranText(CachedData.XMLDocMain, BaseChapter, BaseVerse, Chapter, Verse)
            ElseIf Division = 7 Then
                Node = TanzilReader.GetSajdaByIndex(Index)
                BaseChapter = CInt(Node.Attributes.GetNamedItem("sura").Value)
                BaseVerse = CInt(Node.Attributes.GetNamedItem("aya").Value)
                QuranText = New Collections.Generic.List(Of String())
                QuranText.Add(GetQuranText(CachedData.XMLDocMain, BaseChapter, BaseVerse, BaseVerse))
            ElseIf Division = 8 Then
                BaseChapter = CachedData.IslamData.QuranSelections(Index).SelectionInfo(SectionCount).ChapterNumber
                BaseVerse = CachedData.IslamData.QuranSelections(Index).SelectionInfo(SectionCount).VerseNumber
                QuranText = QuranTextRangeLookup(BaseChapter, BaseVerse, CachedData.IslamData.QuranSelections(Index).SelectionInfo(SectionCount).WordNumber, 0, CachedData.IslamData.QuranSelections(Index).SelectionInfo(SectionCount).ExtraVerseNumber, CachedData.IslamData.QuranSelections(Index).SelectionInfo(SectionCount).EndWordNumber)
            ElseIf Division = 9 Then
                QuranText = New Collections.Generic.List(Of String())
                For SubCount = 0 To CachedData.LetterDictionary(CachedData.IslamData.ArabicLetters(Index).Symbol)(Keys(SectionCount)).Count - 1
                    QuranText.Add(GetQuranText(CachedData.XMLDocMain, CType(CachedData.LetterDictionary(CachedData.IslamData.ArabicLetters(Index).Symbol)(Keys(SectionCount))(SubCount), Integer())(0), CType(CachedData.LetterDictionary(CachedData.IslamData.ArabicLetters(Index).Symbol)(Keys(SectionCount))(SubCount), Integer())(1), CType(CachedData.LetterDictionary(CachedData.IslamData.ArabicLetters(Index).Symbol)(Keys(SectionCount))(SubCount), Integer())(1)))
                    If SubCount <> CachedData.LetterDictionary(CachedData.IslamData.ArabicLetters(Index).Symbol)(Keys(SectionCount)).Count - 1 Then
                        Renderer.Items.AddRange(DoGetRenderedQuranText(QuranText, BaseChapter, BaseVerse, Translation, SchemeType, Scheme, TranslationIndex).Items)
                    End If
                Next
            Else
                QuranText = Nothing
            End If
            Renderer.Items.AddRange(DoGetRenderedQuranText(QuranText, BaseChapter, BaseVerse, Translation, SchemeType, Scheme, TranslationIndex).Items)
        Next
        Return Renderer
    End Function
    Public Shared Function GetRenderedQuranText(ByVal Item As PageLoader.TextItem) As RenderArray
        Dim Division As Integer = 0
        Dim Index As Integer = 1
        Dim Strings As String = HttpContext.Current.Request.QueryString.Get("qurandivision")
        If Not Strings Is Nothing Then Division = CInt(Strings)
        Strings = HttpContext.Current.Request.QueryString.Get("quranselection")
        If Not Strings Is Nothing Then Index = CInt(Strings)
        Dim SchemeType As Arabic.TranslitScheme = If(CInt(HttpContext.Current.Request.QueryString.Get("translitscheme")) >= 2, 2 + CInt(HttpContext.Current.Request.QueryString.Get("translitscheme")) Mod 2, CInt(HttpContext.Current.Request.QueryString.Get("translitscheme")))
        Dim Scheme As String = If(CInt(HttpContext.Current.Request.QueryString.Get("translitscheme")) >= 2, CachedData.IslamData.TranslitSchemes((CInt(HttpContext.Current.Request.QueryString.Get("translitscheme")) - 2) \ 2).Name, String.Empty)
        Dim TranslationIndex As Integer = GetTranslationIndex(HttpContext.Current.Request.QueryString.Get("qurantranslation"))
        Return GetQuranTextBySelection(Division, Index, HttpContext.Current.Request.QueryString.Get("qurantranslation"), SchemeType, Scheme, TranslationIndex)
    End Function
    Public Shared Function DoGetRenderedQuranText(QuranText As Collections.Generic.List(Of String()), BaseChapter As Integer, BaseVerse As Integer, Translation As String, SchemeType As Arabic.TranslitScheme, Scheme As String, TranslationIndex As Integer) As RenderArray
        Dim Text As String
        Dim Node As System.Xml.XmlNode
        Dim Renderer As New RenderArray
        Dim Lines As String() = IO.File.ReadAllLines(Utility.GetFilePath("metadata\" + GetTranslationFileName(Translation)))
        Dim W4WLines As String() = IO.File.ReadAllLines(Utility.GetFilePath("metadata\en.w4w.shehnazshaikh.txt"))
        If Not QuranText Is Nothing Then
            For Chapter = 0 To QuranText.Count - 1
                Dim ChapterNode As System.Xml.XmlNode = GetChapterByIndex(BaseChapter + Chapter)
                Renderer.Items.Add(New RenderArray.RenderItem(RenderArray.RenderTypes.eHeaderLeft, New RenderArray.RenderText() {New RenderArray.RenderText(RenderArray.RenderDisplayClass.eArabic, Arabic.RightToLeftMark + Arabic.TransliterateFromBuckwalter("'aAya`tuhaA " + ChapterNode.Attributes.GetNamedItem("ayas").Value + " ")), New RenderArray.RenderText(RenderArray.RenderDisplayClass.eTransliteration, Arabic.TransliterateToScheme(Arabic.TransliterateFromBuckwalter("'aAya`tuhaA " + ChapterNode.Attributes.GetNamedItem("ayas").Value + " "), SchemeType, Scheme).Trim()), New RenderArray.RenderText(RenderArray.RenderDisplayClass.eLTR, "Verses " + ChapterNode.Attributes.GetNamedItem("ayas").Value + " ")}))
                Renderer.Items.Add(New RenderArray.RenderItem(RenderArray.RenderTypes.eHeaderCenter, New RenderArray.RenderText() {New RenderArray.RenderText(RenderArray.RenderDisplayClass.eArabic, Arabic.RightToLeftMark + Arabic.TransliterateFromBuckwalter("suwrapu " + CachedData.IslamData.QuranChapters(CInt(ChapterNode.Attributes.GetNamedItem("index").Value) - 1).Name + " ")), New RenderArray.RenderText(RenderArray.RenderDisplayClass.eTransliteration, Arabic.TransliterateToScheme(Arabic.TransliterateFromBuckwalter("suwrapu " + CachedData.IslamData.QuranChapters(CInt(ChapterNode.Attributes.GetNamedItem("index").Value) - 1).Name + " "), SchemeType, Scheme).Trim()), New RenderArray.RenderText(RenderArray.RenderDisplayClass.eLTR, "Chapter " + TanzilReader.GetChapterEName(ChapterNode) + " ")}))
                Renderer.Items.Add(New RenderArray.RenderItem(RenderArray.RenderTypes.eHeaderRight, New RenderArray.RenderText() {New RenderArray.RenderText(RenderArray.RenderDisplayClass.eArabic, Arabic.RightToLeftMark + Arabic.TransliterateFromBuckwalter("rukuwEaAtuhaA " + ChapterNode.Attributes.GetNamedItem("rukus").Value + " ")), New RenderArray.RenderText(RenderArray.RenderDisplayClass.eTransliteration, Arabic.TransliterateToScheme(Arabic.TransliterateFromBuckwalter("rukuwEaAtuhaA " + ChapterNode.Attributes.GetNamedItem("rukus").Value + " "), SchemeType, Scheme).Trim()), New RenderArray.RenderText(RenderArray.RenderDisplayClass.eLTR, "Rukus " + ChapterNode.Attributes.GetNamedItem("rukus").Value + " ")}))
                For Verse = 0 To QuranText(Chapter).Length - 1
                    Dim Items As New Collections.Generic.List(Of RenderArray.RenderItem)
                    Text = String.Empty
                    'hizb symbols not needed as Quranic text already contains them
                    'If BaseChapter + Chapter <> 1 AndAlso TanzilReader.IsQuarterStart(BaseChapter + Chapter, CInt(IIf(Chapter = 0, BaseVerse, 1)) + Verse) Then
                    '    Text += Arabic.TransliterateFromBuckwalter("B")
                    '    Items.Add(New RenderArray.RenderItem(RenderArray.RenderTypes.eText, New RenderArray.RenderText() {New RenderArray.RenderText(RenderArray.RenderDisplayClass.eArabic, Arabic.RightToLeftMark + Arabic.TransliterateFromBuckwalter("B"))}))
                    'End If
                    If CInt(IIf(Chapter = 0, BaseVerse, 1)) + Verse = 1 Then
                        Node = GetTextVerse(GetTextChapter(CachedData.XMLDocMain, BaseChapter + Chapter), 1).Attributes.GetNamedItem("bismillah")
                        If Not Node Is Nothing Then
                            Renderer.Items.Add(New RenderArray.RenderItem(RenderArray.RenderTypes.eText, New RenderArray.RenderText() {New RenderArray.RenderText(RenderArray.RenderDisplayClass.eArabic, Arabic.RightToLeftMark + Node.Value + " "), New RenderArray.RenderText(RenderArray.RenderDisplayClass.eTransliteration, Arabic.TransliterateToScheme(Node.Value, SchemeType, Scheme).Trim()), New RenderArray.RenderText(DirectCast(IIf(IsTranslationTextLTR(TranslationIndex), RenderArray.RenderDisplayClass.eLTR, RenderArray.RenderDisplayClass.eRTL), RenderArray.RenderDisplayClass), TanzilReader.GetTranslationVerse(Lines, 1, 1))}))
                        End If
                    End If
                    Dim Words As String() = QuranText(Chapter)(Verse).Split(" "c)
                    Dim TranslitWords As String() = Arabic.TransliterateToScheme(QuranText(Chapter)(Verse), SchemeType, Scheme).Split(" "c)
                    Dim PauseMarks As Integer = 0
                    For Count As Integer = 0 To Words.Length - 1
                        'handle start/end words here which have space placeholders
                        If Words(Count).Length = 1 AndAlso _
                            Arabic.IsStop(Arabic.FindLetterBySymbol(Words(Count)(0))) Then
                            PauseMarks += 1
                            If Items.Count <> 0 Then
                                Items.Add(New RenderArray.RenderItem(RenderArray.RenderTypes.eText, New RenderArray.RenderText() {New RenderArray.RenderText(RenderArray.RenderDisplayClass.eArabic, Arabic.RightToLeftMark + Words(Count)), New RenderArray.RenderText(RenderArray.RenderDisplayClass.eTransliteration, TranslitWords(Count))}))
                            End If
                        ElseIf Words(Count).Length = 0 Then
                        Else
                            Items.Add(New RenderArray.RenderItem(RenderArray.RenderTypes.eText, New RenderArray.RenderText() {New RenderArray.RenderText(RenderArray.RenderDisplayClass.eArabic, Arabic.RightToLeftMark + Words(Count)), New RenderArray.RenderText(RenderArray.RenderDisplayClass.eTransliteration, TranslitWords(Count)), New RenderArray.RenderText(DirectCast(IIf(IsTranslationTextLTR(TranslationIndex), RenderArray.RenderDisplayClass.eLTR, RenderArray.RenderDisplayClass.eRTL), RenderArray.RenderDisplayClass), TanzilReader.GetW4WTranslationVerse(W4WLines, BaseChapter + Chapter, CInt(IIf(Chapter = 0, BaseVerse, 1)) + Verse, Count - PauseMarks))}))
                        End If
                    Next
                    Text += QuranText(Chapter)(Verse) + " "
                    If TanzilReader.IsSajda(BaseChapter + Chapter, CInt(IIf(Chapter = 0, BaseVerse, 1)) + Verse) Then
                        'Sajda markers are already in the text
                        'Text += Arabic.TransliterateFromBuckwalter("R")
                        'Items.Add(New RenderArray.RenderItem(RenderArray.RenderTypes.eText, New RenderArray.RenderText() {New RenderArray.RenderText(RenderArray.RenderDisplayClass.eArabic, Arabic.RightToLeftMark + Arabic.TransliterateFromBuckwalter("R"))}))
                    End If
                    Text += Arabic.TransliterateFromBuckwalter("=" + CStr(CInt(IIf(Chapter = 0, BaseVerse, 1)) + Verse)) + " "
                    Items.Add(New RenderArray.RenderItem(RenderArray.RenderTypes.eText, New RenderArray.RenderText() {New RenderArray.RenderText(RenderArray.RenderDisplayClass.eArabic, Arabic.RightToLeftMark + Arabic.TransliterateFromBuckwalter("=" + CStr(CInt(IIf(Chapter = 0, BaseVerse, 1)) + Verse))), New RenderArray.RenderText(DirectCast(IIf(IsTranslationTextLTR(TranslationIndex), RenderArray.RenderDisplayClass.eLTR, RenderArray.RenderDisplayClass.eRTL), RenderArray.RenderDisplayClass), "(" + CStr(CInt(IIf(Chapter = 0, BaseVerse, 1)) + Verse) + ")")}))
                    'Text += Arabic.TransliterateFromBuckwalter("(" + CStr(IIf(Chapter = 0, BaseVerse, 1) + Verse) + ") ")
                    Renderer.Items.Add(New RenderArray.RenderItem(RenderArray.RenderTypes.eText, New RenderArray.RenderText() {New RenderArray.RenderText(RenderArray.RenderDisplayClass.eNested, Items), New RenderArray.RenderText(RenderArray.RenderDisplayClass.eArabic, Arabic.RightToLeftMark + Text), New RenderArray.RenderText(RenderArray.RenderDisplayClass.eTransliteration, Arabic.TransliterateToScheme(QuranText(Chapter)(Verse) + " " + Arabic.TransliterateFromBuckwalter("=" + CStr(CInt(IIf(Chapter = 0, BaseVerse, 1)) + Verse)) + " ", SchemeType, Scheme).Trim()), New RenderArray.RenderText(DirectCast(IIf(IsTranslationTextLTR(TranslationIndex), RenderArray.RenderDisplayClass.eLTR, RenderArray.RenderDisplayClass.eRTL), RenderArray.RenderDisplayClass), "(" + CStr(CInt(IIf(Chapter = 0, BaseVerse, 1)) + Verse) + ") " + TanzilReader.GetTranslationVerse(Lines, BaseChapter + Chapter, CInt(IIf(Chapter = 0, BaseVerse, 1)) + Verse))}))
                Next
            Next
        End If
        Return Renderer
    End Function
    Public Shared Function GetQuranText(ByVal XMLDocMain As System.Xml.XmlDocument, ByVal StartChapter As Integer, ByVal StartAyat As Integer, ByVal EndChapter As Integer, ByVal EndAyat As Integer) As Collections.Generic.List(Of String())
        Dim Count As Integer
        If StartChapter = -1 Then StartChapter = 1
        If EndChapter = -1 Then EndChapter = GetChapterCount()
        Dim ChapterVerses As New Collections.Generic.List(Of String())
        For Count = StartChapter To EndChapter
            ChapterVerses.Add(GetQuranText(XMLDocMain, Count, CInt(IIf(StartChapter = Count, StartAyat, -1)), CInt(IIf(EndChapter = Count, EndAyat, -1))))
        Next
        Return ChapterVerses
    End Function
    Public Shared Function GetQuranText(ByVal XMLDocMain As System.Xml.XmlDocument, ByVal Chapter As Integer, ByVal StartVerse As Integer, ByVal EndVerse As Integer) As String()
        Dim Count As Integer
        If StartVerse = -1 Then StartVerse = 1
        If EndVerse = -1 Then EndVerse = GetVerseCount(Chapter)
        Dim Verses(EndVerse - StartVerse) As String
        For Count = StartVerse To EndVerse
            Verses(Count - StartVerse) = GetTextVerse(GetTextChapter(XMLDocMain, Chapter), Count).Attributes.GetNamedItem("text").Value
        Next
        Return Verses
    End Function
    Public Shared Function GetTextChapter(ByVal XMLDocMain As System.Xml.XmlDocument, ByVal Chapter As Integer) As System.Xml.XmlNode
        Return Utility.GetChildNodeByIndex("sura", "index", Chapter, XMLDocMain.DocumentElement.ChildNodes)
    End Function
    Public Shared Function GetTextVerse(ByVal ChapterNode As System.Xml.XmlNode, ByVal Verse As Integer) As System.Xml.XmlNode
        Return Utility.GetChildNodeByIndex("aya", "index", Verse, ChapterNode.ChildNodes)
    End Function
    Public Shared Function GetVerseCount(ByVal Chapter As Integer) As Integer
        Return CInt(GetChapterByIndex(Chapter).Attributes.GetNamedItem("ayas").Value)
    End Function
    Public Shared Sub GetPreviousChapterVerse(ByRef Chapter As Integer, ByRef Verse As Integer)
        If Verse = 1 Then
            If Chapter <> 1 Then
                Chapter -= 1
                Verse = GetVerseCount(Chapter)
            End If
        Else
            Verse -= 1
        End If
    End Sub
    Public Shared Function GetVerseNumber(ByVal Chapter As Integer, ByVal Verse As Integer) As Integer
        Return CInt(GetChapterByIndex(Chapter).Attributes.GetNamedItem("start").Value) + Verse
    End Function
    Public Shared Function IsTranslationTextLTR(Index As Integer) As Boolean
        Return Not Languages.GetLanguageInfoByCode(CachedData.IslamData.Translations.TranslationList(Index).FileName.Substring(0, 2)).IsRTL
    End Function
    Public Shared Function GetTranslationVerse(Lines As String(), ByVal Chapter As Integer, ByVal Verse As Integer) As String
        GetTranslationVerse = Lines(GetVerseNumber(Chapter, Verse) - 1)
    End Function
    Public Shared Function GetW4WTranslationVerse(Lines As String(), ByVal Chapter As Integer, ByVal Verse As Integer, ByVal Word As Integer) As String
        Dim Words As String() = Lines(GetVerseNumber(Chapter, Verse) - 1).Split("|"c)
        If Word >= Words.Length Then
            GetW4WTranslationVerse = String.Empty
        Else
            GetW4WTranslationVerse = Words(Word)
        End If
    End Function
    Public Shared Function IsQuarterStart(ByVal Chapter As Integer, ByVal Verse As Integer) As Boolean
        Dim Count As Integer
        Dim Node As System.Xml.XmlNode
        For Count = 0 To Utility.GetChildNode("hizbs", CachedData.XMLDocInfo.DocumentElement.ChildNodes).ChildNodes.Count - 1
            Node = Utility.GetChildNode("hizbs", CachedData.XMLDocInfo.DocumentElement.ChildNodes).ChildNodes.Item(Count)
            If Node.Name = "quarter" AndAlso _
                CInt(Node.Attributes.GetNamedItem("sura").Value) = Chapter AndAlso _
                CInt(Node.Attributes.GetNamedItem("aya").Value) = Verse Then
                Return True
            End If
        Next
        Return False
    End Function
    Public Shared Function IsSajda(ByVal Chapter As Integer, ByVal Verse As Integer) As Boolean
        Dim Count As Integer
        Dim Node As System.Xml.XmlNode
        For Count = 0 To Utility.GetChildNode("sajdas", CachedData.XMLDocInfo.DocumentElement.ChildNodes).ChildNodes.Count - 1
            Node = Utility.GetChildNode("sajdas", CachedData.XMLDocInfo.DocumentElement.ChildNodes).ChildNodes.Item(Count)
            If Node.Name = "sajda" AndAlso _
                CInt(Node.Attributes.GetNamedItem("sura").Value) = Chapter AndAlso _
                CInt(Node.Attributes.GetNamedItem("aya").Value) = Verse Then
                Return True
            End If
        Next
        Return False
    End Function
    Public Shared Function GetImportantNames() As Array()
        Dim Names() As Array = Array.ConvertAll(CachedData.IslamData.QuranSelections, Function(Convert As IslamData.QuranSelection) New Object() {Utility.LoadResourceString("IslamInfo_" + Convert.Description), CInt(Array.IndexOf(CachedData.IslamData.QuranSelections, Convert))})
        Array.Sort(Names, New Utility.CompareNameValueArray)
        Return Names
    End Function
    Public Shared Function GetChapterCount() As Integer
        Return Utility.GetChildNodeCount("sura", Utility.GetChildNode("suras", CachedData.XMLDocInfo.DocumentElement.ChildNodes))
    End Function
    Public Shared Function GetChapterByIndex(ByVal Index As Integer) As System.Xml.XmlNode
        Return Utility.GetChildNodeByIndex("sura", "index", Index, Utility.GetChildNode("suras", CachedData.XMLDocInfo.DocumentElement.ChildNodes).ChildNodes)
    End Function
    Public Shared Function GetChapterNames() As Array()
        Dim SchemeType As Arabic.TranslitScheme = If(CInt(HttpContext.Current.Request.QueryString.Get("translitscheme")) >= 2, 2 + CInt(HttpContext.Current.Request.QueryString.Get("translitscheme")) Mod 2, CInt(HttpContext.Current.Request.QueryString.Get("translitscheme")))
        Dim Scheme As String = If(CInt(HttpContext.Current.Request.QueryString.Get("translitscheme")) >= 2, CachedData.IslamData.TranslitSchemes((CInt(HttpContext.Current.Request.QueryString.Get("translitscheme")) - 2) \ 2).Name, String.Empty)
        Dim Names() As Array = Array.ConvertAll(Utility.GetChildNodes("sura", Utility.GetChildNode("suras", CachedData.XMLDocInfo.DocumentElement.ChildNodes).ChildNodes), Function(Convert As System.Xml.XmlNode) New Object() {Convert.Attributes.GetNamedItem("index").Value + ". " + GetChapterEName(Convert) + " (" + Arabic.RightToLeftMark + Arabic.TransliterateFromBuckwalter("suwrapu " + CachedData.IslamData.QuranChapters(CInt(Convert.Attributes.GetNamedItem("index").Value) - 1).Name) + Arabic.LeftToRightMark + ")" + If(Scheme = Arabic.TranslitScheme.None, String.Empty, " " + Arabic.TransliterateToScheme(Arabic.TransliterateFromBuckwalter(CachedData.IslamData.QuranChapters(CInt(Convert.Attributes.GetNamedItem("index").Value) - 1).Name), SchemeType, Scheme)), CInt(Convert.Attributes.GetNamedItem("index").Value)})
        Array.Sort(Names, New Utility.CompareNameValueArray)
        Return Names
    End Function
    Public Shared Function GetChapterEName(ByVal ChapterNode As System.Xml.XmlNode) As String
        Return Utility.LoadResourceString("IslamInfo_QuranChapter" + ChapterNode.Attributes.GetNamedItem("index").Value)
    End Function
    Public Shared Function GetChapterNamesByRevelationOrder() As Array()
        Dim SchemeType As Arabic.TranslitScheme = If(CInt(HttpContext.Current.Request.QueryString.Get("translitscheme")) >= 2, 2 + CInt(HttpContext.Current.Request.QueryString.Get("translitscheme")) Mod 2, CInt(HttpContext.Current.Request.QueryString.Get("translitscheme")))
        Dim Scheme As String = If(CInt(HttpContext.Current.Request.QueryString.Get("translitscheme")) >= 2, CachedData.IslamData.TranslitSchemes((CInt(HttpContext.Current.Request.QueryString.Get("translitscheme")) - 2) \ 2).Name, String.Empty)
        Dim Names() As Array = Array.ConvertAll(Utility.GetChildNodes("sura", Utility.GetChildNode("suras", CachedData.XMLDocInfo.DocumentElement.ChildNodes).ChildNodes), Function(Convert As System.Xml.XmlNode) New Object() {Convert.Attributes.GetNamedItem("index").Value + ". " + GetChapterEName(Convert) + " (" + Arabic.RightToLeftMark + Arabic.TransliterateFromBuckwalter("suwrapu " + CachedData.IslamData.QuranChapters(CInt(Convert.Attributes.GetNamedItem("index").Value) - 1).Name) + Arabic.LeftToRightMark + ")" + If(Scheme = Arabic.TranslitScheme.None, String.Empty, " " + Arabic.TransliterateToScheme(Arabic.TransliterateFromBuckwalter(CachedData.IslamData.QuranChapters(CInt(Convert.Attributes.GetNamedItem("index").Value) - 1).Name), SchemeType, Scheme)), CInt(Convert.Attributes.GetNamedItem("order").Value)})
        Array.Sort(Names, New Utility.CompareNameValueArray)
        Return Names
    End Function
    Public Shared Function GetChapterIndexByRevelationOrder(ByVal Index As Integer) As System.Xml.XmlNode
        Dim Count As Integer
        Dim ChapterNode As System.Xml.XmlNode
        For Count = 0 To Utility.GetChildNode("suras", CachedData.XMLDocInfo.DocumentElement.ChildNodes).ChildNodes.Count - 1
            ChapterNode = Utility.GetChildNode("suras", CachedData.XMLDocInfo.DocumentElement.ChildNodes).ChildNodes.Item(Count)
            If ChapterNode.Name = "sura" AndAlso CInt(ChapterNode.Attributes.GetNamedItem("order").Value) = Index Then
                Return ChapterNode
            End If
        Next
        Return Nothing
    End Function
    Public Shared Function GetPartNames() As Array()
        Dim Names() As Array = Array.ConvertAll(Utility.GetChildNodes("juz", Utility.GetChildNode("juzs", CachedData.XMLDocInfo.DocumentElement.ChildNodes).ChildNodes), Function(Convert As System.Xml.XmlNode) New Object() {Convert.Attributes.GetNamedItem("index").Value + " (" + Arabic.TransliterateFromBuckwalter("juz " + CachedData.IslamData.QuranParts(CInt(Convert.Attributes.GetNamedItem("index").Value) - 1).Name + " ") + ")", CInt(Convert.Attributes.GetNamedItem("index").Value)})
        Array.Sort(Names, New Utility.CompareNameValueArray)
        Return Names
    End Function
    Public Shared Function GetPartCount() As Integer
        Return Utility.GetChildNodeCount("juz", Utility.GetChildNode("juzs", CachedData.XMLDocInfo.DocumentElement.ChildNodes))
    End Function
    Public Shared Function GetPartByIndex(ByVal Index As Integer) As System.Xml.XmlNode
        Return Utility.GetChildNodeByIndex("juz", "index", Index, Utility.GetChildNode("juzs", CachedData.XMLDocInfo.DocumentElement.ChildNodes).ChildNodes)
    End Function
    Public Shared Function GetGroupNames() As Array()
        Dim Names() As Array = Array.ConvertAll(Utility.GetChildNodes("quarter", Utility.GetChildNode("hizbs", CachedData.XMLDocInfo.DocumentElement.ChildNodes).ChildNodes), Function(Convert As System.Xml.XmlNode) New Object() {Convert.Attributes.GetNamedItem("index").Value, CInt(Convert.Attributes.GetNamedItem("index").Value)})
        Array.Sort(Names, New Utility.CompareNameValueArray)
        Return Names
    End Function
    Public Shared Function GetGroupCount() As Integer
        Return Utility.GetChildNodeCount("quarter", Utility.GetChildNode("hizbs", CachedData.XMLDocInfo.DocumentElement.ChildNodes))
    End Function
    Public Shared Function GetGroupByIndex(ByVal Index As Integer) As System.Xml.XmlNode
        Return Utility.GetChildNodeByIndex("quarter", "index", Index, Utility.GetChildNode("hizbs", CachedData.XMLDocInfo.DocumentElement.ChildNodes).ChildNodes)
    End Function
    Public Shared Function GetStationNames() As Array()
        Dim Names() As Array = Array.ConvertAll(Utility.GetChildNodes("manzil", Utility.GetChildNode("manzils", CachedData.XMLDocInfo.DocumentElement.ChildNodes).ChildNodes), Function(Convert As System.Xml.XmlNode) New Object() {Convert.Attributes.GetNamedItem("index").Value, CInt(Convert.Attributes.GetNamedItem("index").Value)})
        Array.Sort(Names, New Utility.CompareNameValueArray)
        Return Names
    End Function
    Public Shared Function GetStationCount() As Integer
        Return Utility.GetChildNodeCount("manzil", Utility.GetChildNode("manzils", CachedData.XMLDocInfo.DocumentElement.ChildNodes))
    End Function
    Public Shared Function GetStationByIndex(ByVal Index As Integer) As System.Xml.XmlNode
        Return Utility.GetChildNodeByIndex("manzil", "index", Index, Utility.GetChildNode("manzils", CachedData.XMLDocInfo.DocumentElement.ChildNodes).ChildNodes)
    End Function
    Public Shared Function GetSectionNames() As Array()
        Dim Names() As Array = Array.ConvertAll(Utility.GetChildNodes("ruku", Utility.GetChildNode("rukus", CachedData.XMLDocInfo.DocumentElement.ChildNodes).ChildNodes), Function(Convert As System.Xml.XmlNode) New Object() {Convert.Attributes.GetNamedItem("index").Value, CInt(Convert.Attributes.GetNamedItem("index").Value)})
        Array.Sort(Names, New Utility.CompareNameValueArray)
        Return Names
    End Function
    Public Shared Function GetSectionCount() As Integer
        Return Utility.GetChildNodeCount("ruku", Utility.GetChildNode("rukus", CachedData.XMLDocInfo.DocumentElement.ChildNodes))
    End Function
    Public Shared Function GetSectionByIndex(ByVal Index As Integer) As System.Xml.XmlNode
        Return Utility.GetChildNodeByIndex("ruku", "index", Index, Utility.GetChildNode("rukus", CachedData.XMLDocInfo.DocumentElement.ChildNodes).ChildNodes)
    End Function
    Public Shared Function GetPageNames() As Array()
        Dim Names() As Array = Array.ConvertAll(Utility.GetChildNodes("page", Utility.GetChildNode("pages", CachedData.XMLDocInfo.DocumentElement.ChildNodes).ChildNodes), Function(Convert As System.Xml.XmlNode) New Object() {Convert.Attributes.GetNamedItem("index").Value, CInt(Convert.Attributes.GetNamedItem("index").Value)})
        Array.Sort(Names, New Utility.CompareNameValueArray)
        Return Names
    End Function
    Public Shared Function GetPageCount() As Integer
        Return Utility.GetChildNodeCount("page", Utility.GetChildNode("pages", CachedData.XMLDocInfo.DocumentElement.ChildNodes))
    End Function
    Public Shared Function GetPageByIndex(ByVal Index As Integer) As System.Xml.XmlNode
        Return Utility.GetChildNodeByIndex("page", "index", Index, Utility.GetChildNode("pages", CachedData.XMLDocInfo.DocumentElement.ChildNodes).ChildNodes)
    End Function
    Public Shared Function GetSajdaNames() As Array()
        Dim Names() As Array = Array.ConvertAll(Utility.GetChildNodes("sajda", Utility.GetChildNode("sajdas", CachedData.XMLDocInfo.DocumentElement.ChildNodes).ChildNodes), Function(Convert As System.Xml.XmlNode) New Object() {Convert.Attributes.GetNamedItem("index").Value, CInt(Convert.Attributes.GetNamedItem("index").Value)})
        Array.Sort(Names, New Utility.CompareNameValueArray)
        Return Names
    End Function
    Public Shared Function GetSajdaCount() As Integer
        Return Utility.GetChildNodeCount("sajda", Utility.GetChildNode("sajdas", CachedData.XMLDocInfo.DocumentElement.ChildNodes))
    End Function
    Public Shared Function GetSajdaByIndex(ByVal Index As Integer) As System.Xml.XmlNode
        Return Utility.GetChildNodeByIndex("sajda", "index", Index, Utility.GetChildNode("sajdas", CachedData.XMLDocInfo.DocumentElement.ChildNodes).ChildNodes)
    End Function
End Class
Public Class HadithReader
    Public Shared Function GetCollectionChangeOnlyJS() As String
        Dim JSArrays As String = Utility.MakeJSArray(Array.ConvertAll(CachedData.IslamData.Collections, Function(Convert As IslamData.CollectionInfo) Utility.MakeJSArray(Array.ConvertAll(Of IslamData.CollectionInfo.CollTranslationInfo, String)(Convert.Translations, Function(TranslateBlock As IslamData.CollectionInfo.CollTranslationInfo) Utility.MakeJSArray(New String() {Utility.LoadResourceString("lang_local" + Languages.GetLanguageInfoByCode(TranslateBlock.FileName.Substring(0, 2)).Code) + ": " + Utility.LoadResourceString("IslamInfo_" + TranslateBlock.Name), TranslateBlock.FileName})), True)), True)
        Return "function changeHadithCollection(index) { var iCount; var hadithdata = " + JSArrays + "; var eSelect = $('#hadithtranslation').get(0); clearOptionList(eSelect); for (iCount = 0; iCount < hadithdata[index].length; iCount++) { eSelect.options.add(new Option(hadithdata[index][iCount][0], hadithdata[index][iCount][1])); } }"
    End Function
    Public Shared Function GetCollectionChangeWithBooksJS() As String()
        Dim JSArrays As String = Utility.MakeJSArray(Array.ConvertAll(Of IslamData.CollectionInfo, String)(CachedData.IslamData.Collections, Function(Convert As IslamData.CollectionInfo) Utility.MakeJSArray(Array.ConvertAll(Of Array, String)(HadithReader.GetBookNamesByCollection(GetCollectionIndex(Convert.Name)), Function(BookNames As Array) Utility.MakeJSArray(New String() {CStr(BookNames.GetValue(0)), CStr(BookNames.GetValue(1))})), True)), True)
        Return New String() {"javascript: changeHadithCollectionBooks(this.selectedIndex);", String.Empty, Utility.GetClearOptionListJS(), _
                             GetCollectionChangeOnlyJS(), _
        "function changeHadithCollectionBooks(index) { changeHadithCollection(index); var iCount; var hadithtdata = " + JSArrays + "; var eSelect = $('#hadithbook').get(0); clearOptionList(eSelect); for (iCount = 0; iCount < hadithtdata[index].length; iCount++) { eSelect.options.add(new Option(hadithtdata[index][iCount][0], hadithtdata[index][iCount][1])); } }"}
    End Function
    Public Shared Function GetCollectionChangeJS() As String()
        Return New String() {"javascript: changeHadithCollection(this.selectedIndex);", String.Empty, Utility.GetClearOptionListJS(), _
                             GetCollectionChangeOnlyJS()}
    End Function
    Public Shared Function GetCollectionXMLMetaDataDownload() As String()
        Return New String() {Utility.GetPageString("Source&File=" + CachedData.IslamData.Collections(GetCurrentCollection()).FileName + "-data.xml"), Utility.LoadResourceString("IslamInfo_" + CachedData.IslamData.Collections(GetCurrentCollection()).Name) + " XML metadata"}
    End Function
    Public Shared Function GetCollectionXMLDownload() As String()
        Return New String() {Utility.GetPageString("Source&File=" + CachedData.IslamData.Collections(GetCurrentCollection()).FileName + ".xml"), Utility.LoadResourceString("IslamInfo_" + CachedData.IslamData.Collections(GetCurrentCollection()).Name) + " XML source text"}
    End Function
    Public Shared Function GetTranslationXMLMetaDataDownload() As String()
        Dim TranslationIndex As Integer = GetTranslationIndex(GetCurrentCollection(), HttpContext.Current.Request.QueryString.Get("hadithtranslation"))
        If TranslationIndex = -1 Then Return New String() {}
        Return New String() {Utility.GetPageString("Source&File=" + GetTranslationXMLFileName(GetCurrentCollection(), HttpContext.Current.Request.QueryString.Get("hadithtranslation")) + ".xml"), CachedData.IslamData.Collections(GetCurrentCollection()).Translations(TranslationIndex).Name + " XML metadata"}
    End Function
    Public Shared Function GetTranslationTextDownload() As String()
        Dim TranslationIndex As Integer = GetTranslationIndex(GetCurrentCollection(), HttpContext.Current.Request.QueryString.Get("hadithtranslation"))
        If TranslationIndex = -1 Then Return New String() {}
        Return New String() {Utility.GetPageString("Source&File=" + GetTranslationFileName(GetCurrentCollection(), HttpContext.Current.Request.QueryString.Get("hadithtranslation")) + ".txt"), CachedData.IslamData.Collections(GetCurrentCollection()).Translations(TranslationIndex).Name + " raw source text"}
    End Function
    Public Shared Function GetCollectionIndex(ByVal Name As String) As Integer
        Dim Count As Integer
        For Count = 0 To CachedData.IslamData.Collections.Length - 1
            If Name = CachedData.IslamData.Collections(Count).Name Then Return Count
        Next
        Return -1
    End Function
    Public Shared Function GetCollectionNames() As String()
        Return Array.ConvertAll(CachedData.IslamData.Collections, Function(Convert As IslamData.CollectionInfo) Utility.LoadResourceString("IslamInfo_" + Convert.Name))
    End Function
    Public Shared Function GetChapterByIndex(ByVal BookNode As System.Xml.XmlNode, ByVal ChapterIndex As Integer) As System.Xml.XmlNode
        Return Utility.GetChildNodeByIndex("chapter", "index", ChapterIndex, BookNode.ChildNodes)
    End Function
    Public Shared Function GetSubChapterByIndex(ByVal ChapterNode As System.Xml.XmlNode, ByVal SubChapterIndex As Integer) As System.Xml.XmlNode
        Return Utility.GetChildNodeByIndex("subchapter", "index", SubChapterIndex, ChapterNode.ChildNodes)
    End Function
    Public Shared Function GetCurrentCollection() As Integer
        Dim Strings As String = HttpContext.Current.Request.QueryString.Get("hadithcollection")
        If Not Strings Is Nothing Then Return CInt(Strings) Else Return 0
    End Function
    Public Shared Function GetCurrentBook() As Integer
        Dim Strings As String = HttpContext.Current.Request.QueryString.Get("hadithbook")
        If Not Strings Is Nothing Then Return CInt(Strings) Else Return 1
    End Function
    Public Shared Function GetBookNames() As Array()
        Return HadithReader.GetBookNamesByCollection(GetCurrentCollection())
    End Function
    Public Shared Function GetBookEName(ByVal BookNode As System.Xml.XmlNode, CollectionIndex As Integer) As String
        If BookNode Is Nothing Then
            Return String.Empty
        Else
            GetBookEName = Utility.LoadResourceString("IslamInfo_" + CachedData.IslamData.Collections(CollectionIndex).FileName + "Book" + BookNode.Attributes.GetNamedItem("index").Value)
            If GetBookEName Is Nothing Then GetBookEName = String.Empty
        End If
    End Function
    Public Shared Function GetTranslationList() As Array()
        Return Array.ConvertAll(CachedData.IslamData.Collections(GetCurrentCollection()).Translations, Function(Convert As IslamData.CollectionInfo.CollTranslationInfo) New String() {Utility.LoadResourceString("lang_local" + Languages.GetLanguageInfoByCode(Convert.FileName.Substring(0, 2)).Code) + ": " + Utility.LoadResourceString("IslamInfo_" + Convert.Name), Convert.FileName})
    End Function
    Public Shared Function IsTranslationTextLTR(ByVal Index As Integer, Translation As String) As Boolean
        Dim TranslationIndex As Integer = GetTranslationIndex(Index, Translation)
        Return TranslationIndex = -1 OrElse Not Languages.GetLanguageInfoByCode(CachedData.IslamData.Collections(Index).Translations(TranslationIndex).FileName.Substring(0, 2)).IsRTL
    End Function
    Public Shared Function GetTranslationIndex(ByVal Index As Integer, ByVal Translation As String) As Integer
        If String.IsNullOrEmpty(Translation) Then Translation = CachedData.IslamData.Collections(Index).DefaultTranslation
        Dim Count As Integer = Array.FindIndex(CachedData.IslamData.Collections(Index).Translations, Function(Test As IslamData.CollectionInfo.CollTranslationInfo) Test.FileName = Translation)
        If Count = -1 Then
            Translation = CachedData.IslamData.Collections(Index).DefaultTranslation
            Count = Array.FindIndex(CachedData.IslamData.Collections(Index).Translations, Function(Test As IslamData.CollectionInfo.CollTranslationInfo) Test.FileName = Translation)
        End If
        Return Count
    End Function
    Public Shared Function GetTranslationXMLFileName(ByVal Index As Integer, ByVal Translation As String) As String
        Dim TranslationIndex As Integer = GetTranslationIndex(Index, Translation)
        Return CachedData.IslamData.Collections(Index).FileName + "." + CachedData.IslamData.Collections(Index).Translations(TranslationIndex).FileName + "-data"
    End Function
    Public Shared Function GetTranslationFileName(ByVal Index As Integer, ByVal Translation As String) As String
        Dim TranslationIndex As Integer = GetTranslationIndex(Index, Translation)
        Return CachedData.IslamData.Collections(Index).FileName + "." + CachedData.IslamData.Collections(Index).Translations(TranslationIndex).FileName
    End Function
    Public Shared Function GetHadithMappingText(ByVal Item As PageLoader.TextItem) As Array()
        Dim Index As Integer = GetCurrentCollection()
        Dim XMLDocTranslate As New System.Xml.XmlDocument
        If CachedData.IslamData.Collections(Index).Translations.Length = 0 Then Return New Array() {}
        XMLDocTranslate.Load(Utility.GetFilePath("metadata\" + GetTranslationXMLFileName(Index, HttpContext.Current.Request.QueryString.Get("hadithtranslation")) + ".xml"))
        Dim Output As New ArrayList
        Output.Add(New String() {})
        If HadithReader.HasVolumes(Index) Then
            Output.Add(New String() {String.Empty, String.Empty, String.Empty, String.Empty, String.Empty, String.Empty})
            Output.Add(New String() {Utility.LoadResourceString("Hadith_Volume"), Utility.LoadResourceString("Hadith_Book"), Utility.LoadResourceString("Hadith_Index"), Utility.LoadResourceString("Hadith_Chapters"), Utility.LoadResourceString("Hadith_Hadiths"), Utility.LoadResourceString("Hadith_Translation")})
        Else
            Output.Add(New String() {String.Empty, String.Empty, String.Empty, String.Empty, String.Empty})
            Output.Add(New String() {Utility.LoadResourceString("Hadith_Book"), Utility.LoadResourceString("Hadith_Index"), Utility.LoadResourceString("Hadith_Chapters"), Utility.LoadResourceString("Hadith_Hadiths"), Utility.LoadResourceString("Hadith_Translation")})
        End If
        If HadithReader.HasVolumes(Index) Then
            Output.AddRange(Array.ConvertAll(Utility.GetChildNodes("book", Utility.GetChildNode("books", CachedData.XMLDocInfos(Index).DocumentElement.ChildNodes).ChildNodes), Function(Convert As System.Xml.XmlNode) New Object() {CStr(HadithReader.GetVolumeIndex(Index, CInt(Convert.Attributes.GetNamedItem("index").Value))), GetBookEName(Convert, Index), Convert.Attributes.GetNamedItem("index").Value, CStr(HadithReader.GetChapterCount(Index, CInt(Convert.Attributes.GetNamedItem("index").Value))), CStr(HadithReader.GetHadithCount(Index, CInt(Convert.Attributes.GetNamedItem("index").Value))), HadithReader.GetBookHadithMapping(XMLDocTranslate, Index, CInt(Convert.Attributes.GetNamedItem("index").Value))}))
        Else
            Output.AddRange(Array.ConvertAll(Utility.GetChildNodes("book", Utility.GetChildNode("books", CachedData.XMLDocInfos(Index).DocumentElement.ChildNodes).ChildNodes), Function(Convert As System.Xml.XmlNode) New Object() {GetBookEName(Convert, Index), Convert.Attributes.GetNamedItem("index").Value, CStr(HadithReader.GetChapterCount(Index, CInt(Convert.Attributes.GetNamedItem("index").Value))), CStr(HadithReader.GetHadithCount(Index, CInt(Convert.Attributes.GetNamedItem("index").Value))), HadithReader.GetBookHadithMapping(XMLDocTranslate, Index, CInt(Convert.Attributes.GetNamedItem("index").Value))}))
        End If
        Return DirectCast(Output.ToArray(GetType(Array)), Array())
    End Function
    Public Shared Function GetRenderedText(ByVal Item As PageLoader.TextItem) As RenderArray
        Dim SchemeType As Arabic.TranslitScheme = If(CInt(HttpContext.Current.Request.QueryString.Get("translitscheme")) >= 2, 2 + CInt(HttpContext.Current.Request.QueryString.Get("translitscheme")) Mod 2, CInt(HttpContext.Current.Request.QueryString.Get("translitscheme")))
        Dim Scheme As String = If(CInt(HttpContext.Current.Request.QueryString.Get("translitscheme")) >= 2, CachedData.IslamData.TranslitSchemes((CInt(HttpContext.Current.Request.QueryString.Get("translitscheme")) - 2) \ 2).Name, String.Empty)
        Dim Translation As String = HttpContext.Current.Request.QueryString.Get("hadithtranslation")
        Return DoGetRenderedText(SchemeType, Scheme, Translation)
    End Function
    Public Shared Function DoGetRenderedText(SchemeType As Arabic.TranslitScheme, Scheme As String, Translation As String) As RenderArray
        Dim Renderer As New RenderArray
        Dim Hadith As Integer
        Dim Index As Integer = GetCurrentCollection()
        Dim BookIndex As Integer = GetCurrentBook()
        Dim HadithText As Collections.Generic.List(Of Collections.Generic.List(Of Object)) = HadithReader.GetHadithText(BookIndex)
        Dim ChapterIndex As Integer = -1
        Dim SubChapterIndex As Integer = -1
        Dim BookNode As System.Xml.XmlNode = HadithReader.GetBookByIndex(Index, BookIndex)
        Dim ChapterNode As System.Xml.XmlNode = Nothing
        Dim SubChapterNode As System.Xml.XmlNode
        If Not BookNode Is Nothing Then
            Renderer.Items.Add(New RenderArray.RenderItem(RenderArray.RenderTypes.eHeaderLeft, New RenderArray.RenderText() {New RenderArray.RenderText(RenderArray.RenderDisplayClass.eArabic, Arabic.RightToLeftMark + Arabic.TransliterateFromBuckwalter("Had~iv " + BookNode.Attributes.GetNamedItem("hadiths").Value + " ")), New RenderArray.RenderText(RenderArray.RenderDisplayClass.eTransliteration, Arabic.TransliterateToScheme(Arabic.TransliterateFromBuckwalter("Had~iv " + BookNode.Attributes.GetNamedItem("hadiths").Value + " "), SchemeType, Scheme).Trim()), New RenderArray.RenderText(RenderArray.RenderDisplayClass.eLTR, "Hadiths: " + BookNode.Attributes.GetNamedItem("hadiths").Value + " ")}))
            Renderer.Items.Add(New RenderArray.RenderItem(RenderArray.RenderTypes.eHeaderCenter, New RenderArray.RenderText() {New RenderArray.RenderText(RenderArray.RenderDisplayClass.eArabic, Arabic.RightToLeftMark + Arabic.TransliterateFromBuckwalter("{lokita`bu " + CStr(BookIndex)) + " " + BookNode.Attributes.GetNamedItem("name").Value + " "), New RenderArray.RenderText(RenderArray.RenderDisplayClass.eTransliteration, Arabic.TransliterateToScheme(Arabic.TransliterateFromBuckwalter("{lokita`bu " + CStr(BookIndex)) + " " + BookNode.Attributes.GetNamedItem("name").Value + " ", SchemeType, Scheme).Trim()), New RenderArray.RenderText(RenderArray.RenderDisplayClass.eLTR, "Book " + CStr(BookIndex) + ": " + GetBookEName(BookNode, Index) + " ")}))
            'Renderer.Items.Add(New RenderArray.RenderItem(RenderArray.RenderTypes.eHeaderRight, New RenderArray.RenderText() {New RenderArray.RenderText(RenderArray.RenderDisplayClass.eArabic, Arabic.RightToLeftMark + Arabic.TransliterateFromBuckwalter("mjld " + Utility.GetChildNode("books", XMLDocInfo(Index).DocumentElement.ChildNodes).ChildNodes.Item(BookIndex).Attributes.GetNamedItem("volume").Value + " ")), New RenderArray.RenderText(RenderArray.RenderDisplayClass.eTransliteration, Arabic.TransliterateToScheme(Arabic.TransliterateFromBuckwalter("mjld " + Utility.GetChildNode("books", XMLDocInfo(Index).DocumentElement.ChildNodes).ChildNodes.Item(BookIndex).Attributes.GetNamedItem("volume").Value + " "), SchemeType, Scheme).Trim()), New RenderArray.RenderText(RenderArray.RenderDisplayClass.eLTR, "Volume " + Utility.GetChildNode("books", XMLDocInfo(Index).DocumentElement.ChildNodes).ChildNodes.Item(BookIndex).Attributes.GetNamedItem("volume").Value + " ")}))
            Dim XMLDocTranslate As New System.Xml.XmlDocument
            Dim Strings() As String = Nothing
            If CachedData.IslamData.Collections(Index).Translations.Length <> 0 Then
                XMLDocTranslate.Load(Utility.GetFilePath("metadata\" + GetTranslationXMLFileName(Index, Translation) + ".xml"))
                Strings = IO.File.ReadAllLines(Utility.GetFilePath("metadata\" + GetTranslationFileName(Index, Translation) + ".txt"))
            End If
            For Hadith = 0 To HadithText.Count - 1
                'Handle missing or excess chapter indexes
                If ChapterIndex <> CInt(HadithText(Hadith)(1)) Then
                    ChapterIndex = CInt(HadithText(Hadith)(1))
                    ChapterNode = GetChapterByIndex(BookNode, ChapterIndex)
                    If Not ChapterNode Is Nothing Then
                        Renderer.Items.Add(New RenderArray.RenderItem(RenderArray.RenderTypes.eHeaderLeft, New RenderArray.RenderText() {New RenderArray.RenderText(RenderArray.RenderDisplayClass.eArabic, Arabic.RightToLeftMark + Arabic.TransliterateFromBuckwalter("Had~iv " + ChapterNode.Attributes.GetNamedItem("hadiths").Value + " ")), New RenderArray.RenderText(RenderArray.RenderDisplayClass.eTransliteration, Arabic.TransliterateToScheme(Arabic.TransliterateFromBuckwalter("Had~iv " + ChapterNode.Attributes.GetNamedItem("hadiths").Value + " "), SchemeType, Scheme).Trim()), New RenderArray.RenderText(RenderArray.RenderDisplayClass.eLTR, "Hadiths: " + ChapterNode.Attributes.GetNamedItem("hadiths").Value + " ")}))
                        Renderer.Items.Add(New RenderArray.RenderItem(RenderArray.RenderTypes.eHeaderCenter, New RenderArray.RenderText() {New RenderArray.RenderText(RenderArray.RenderDisplayClass.eArabic, Arabic.RightToLeftMark + Arabic.TransliterateFromBuckwalter("bAb " + CStr(ChapterIndex)) + " " + ChapterNode.Attributes.GetNamedItem("name").Value + " "), New RenderArray.RenderText(RenderArray.RenderDisplayClass.eTransliteration, Arabic.TransliterateToScheme(Arabic.TransliterateFromBuckwalter("bAb " + CStr(ChapterIndex)) + " " + ChapterNode.Attributes.GetNamedItem("name").Value + " ", SchemeType, Scheme).Trim()), New RenderArray.RenderText(RenderArray.RenderDisplayClass.eLTR, "Chapter " + CStr(ChapterIndex) + ": " + Utility.DefaultValue(Utility.LoadResourceString("IslamInfo_" + CachedData.IslamData.Collections(Index).FileName + "Book" + BookNode.Attributes.GetNamedItem("index").Value + "Chapter" + ChapterNode.Attributes.GetNamedItem("index").Value), String.Empty) + " ")}))
                    End If
                    SubChapterIndex = -1
                End If
                'Handle missing or excess subchapter indexes
                If SubChapterIndex <> CInt(HadithText(Hadith)(2)) Then
                    SubChapterIndex = CInt(HadithText(Hadith)(2))
                    If Not ChapterNode Is Nothing Then
                        SubChapterNode = GetSubChapterByIndex(ChapterNode, SubChapterIndex)
                        If Not SubChapterNode Is Nothing Then
                            Renderer.Items.Add(New RenderArray.RenderItem(RenderArray.RenderTypes.eHeaderLeft, New RenderArray.RenderText() {New RenderArray.RenderText(RenderArray.RenderDisplayClass.eArabic, Arabic.RightToLeftMark + Arabic.TransliterateFromBuckwalter("Had~iv " + SubChapterNode.Attributes.GetNamedItem("hadiths").Value + " ")), New RenderArray.RenderText(RenderArray.RenderDisplayClass.eTransliteration, Arabic.TransliterateToScheme(Arabic.TransliterateFromBuckwalter("Had~iv " + SubChapterNode.Attributes.GetNamedItem("hadiths").Value + " "), SchemeType, Scheme).Trim()), New RenderArray.RenderText(RenderArray.RenderDisplayClass.eLTR, "Hadiths: " + SubChapterNode.Attributes.GetNamedItem("hadiths").Value + " ")}))
                            Renderer.Items.Add(New RenderArray.RenderItem(RenderArray.RenderTypes.eHeaderCenter, New RenderArray.RenderText() {New RenderArray.RenderText(RenderArray.RenderDisplayClass.eArabic, Arabic.RightToLeftMark + Arabic.TransliterateFromBuckwalter("bAb " + CStr(SubChapterIndex)) + " " + SubChapterNode.Attributes.GetNamedItem("name").Value + " "), New RenderArray.RenderText(RenderArray.RenderDisplayClass.eTransliteration, Arabic.TransliterateToScheme(Arabic.TransliterateFromBuckwalter("bAb " + CStr(SubChapterIndex)) + " " + SubChapterNode.Attributes.GetNamedItem("name").Value + " ", SchemeType, Scheme).Trim()), New RenderArray.RenderText(RenderArray.RenderDisplayClass.eLTR, "Sub-Chapter " + CStr(SubChapterIndex) + ": " + Utility.DefaultValue(Utility.LoadResourceString("IslamInfo_" + CachedData.IslamData.Collections(Index).FileName + "Book" + BookNode.Attributes.GetNamedItem("index").Value + "Chapter" + ChapterNode.Attributes.GetNamedItem("index").Value + "Subchapter" + SubChapterNode.Attributes.GetNamedItem("index").Value), String.Empty) + " ")}))
                        End If
                    End If
                End If
                Dim HadithTranslation As String = String.Empty
                If CInt(HadithText(Hadith)(0)) <> 0 Then
                    Dim TranslationLines() As String = HadithReader.GetTranslationHadith(XMLDocTranslate, Strings, Index, BookIndex - 1, CInt(HadithText(Hadith)(0)))
                    Dim Count As Integer
                    For Count = 0 To TranslationLines.Length - 1
                        HadithTranslation += vbCrLf + TranslationLines(Count)
                    Next
                End If
                'Arabic.TransliterateFromBuckwalter("(" + HadithText(Hadith)(0) + ") ")
                Renderer.Items.Add(New RenderArray.RenderItem(RenderArray.RenderTypes.eText, New RenderArray.RenderText() {New RenderArray.RenderText(RenderArray.RenderDisplayClass.eArabic, Arabic.RightToLeftMark + CStr(HadithText(Hadith)(3)) + " " + Arabic.TransliterateFromBuckwalter("=" + CStr(HadithText(Hadith)(0))) + " "), New RenderArray.RenderText(RenderArray.RenderDisplayClass.eTransliteration, Arabic.TransliterateToScheme(CStr(HadithText(Hadith)(3)) + " " + Arabic.TransliterateFromBuckwalter("=" + CStr(HadithText(Hadith)(0))) + " ", SchemeType, Scheme).Trim()), New RenderArray.RenderText(DirectCast(IIf(IsTranslationTextLTR(Index, Translation), RenderArray.RenderDisplayClass.eLTR, RenderArray.RenderDisplayClass.eRTL), RenderArray.RenderDisplayClass), "(" + CStr(HadithText(Hadith)(0)) + ") " + HadithTranslation)}))
                Dim Ranking As Integer() = SiteDatabase.GetHadithRankingData(CachedData.IslamData.Collections(Index).FileName, BookIndex, CInt(HadithText(Hadith)(0)))
                Dim UserRanking As Integer
                If Utility.IsLoggedIn() Then
                    UserRanking = SiteDatabase.GetUserHadithRankingData(Utility.GetUserID(), CachedData.IslamData.Collections(Index).FileName, BookIndex, CInt(HadithText(Hadith)(0)))
                Else
                    UserRanking = -1
                End If
                Renderer.Items.Add(New RenderArray.RenderItem(RenderArray.RenderTypes.eRanking, New RenderArray.RenderText() {New RenderArray.RenderText(RenderArray.RenderDisplayClass.eRanking, CachedData.IslamData.Collections(Index).FileName + "|" + CStr(BookIndex) + "|" + CStr(HadithText(Hadith)(0)) + "|" + CStr(Ranking(0)) + "|" + CStr(Ranking(1)) + "|" + CStr(UserRanking))}))
            Next
        End If
        Return Renderer
    End Function
    Public Shared Function GetHadithTextBook(ByVal XMLDocMain As System.Xml.XmlDocument, ByVal BookIndex As Integer) As System.Xml.XmlNode
        Return Utility.GetChildNodeByIndex("book", "index", BookIndex, XMLDocMain.DocumentElement.ChildNodes)
    End Function
    Public Shared Function GetHadithText(ByVal BookIndex As Integer) As Collections.Generic.List(Of Collections.Generic.List(Of Object))
        Dim Count As Integer
        Dim XMLDocMain As New System.Xml.XmlDocument
        XMLDocMain.Load(Utility.GetFilePath("metadata\" + CachedData.IslamData.Collections(GetCurrentCollection()).FileName + ".xml"))
        Dim BookNode As System.Xml.XmlNode = GetHadithTextBook(XMLDocMain, BookIndex)
        Dim HadithNode As System.Xml.XmlNode
        Dim Hadiths As New Collections.Generic.List(Of Collections.Generic.List(Of Object))
        For Count = 0 To BookNode.ChildNodes.Count - 1
            HadithNode = BookNode.ChildNodes.Item(Count)
            If HadithNode.Name = "hadith" Then
                Dim NextEntry As New Collections.Generic.List(Of Object)
                NextEntry.AddRange(New Object() {CInt(HadithNode.Attributes.GetNamedItem("index").Value), _
                                              CInt(Utility.ParseValue(HadithNode.Attributes.GetNamedItem("sectionindex"), "-1")), _
                                              CInt(Utility.ParseValue(HadithNode.Attributes.GetNamedItem("subsectionindex"), "-1")), _
                                              HadithNode.Attributes.GetNamedItem("text").Value})
                Hadiths.Add(NextEntry)
            End If
        Next
        Return Hadiths
    End Function
    Public Shared Function GetBookCount(ByVal Index As Integer) As Integer
        Return CInt(Utility.GetChildNode("books", CachedData.XMLDocInfos(Index).DocumentElement.ChildNodes).Attributes("count").Value)
    End Function
    Public Shared Function GetBookByIndex(ByVal Index As Integer, ByVal BookIndex As Integer) As System.Xml.XmlNode
        Return Utility.GetChildNodeByIndex("book", "index", BookIndex, Utility.GetChildNode("books", CachedData.XMLDocInfos(Index).DocumentElement.ChildNodes).ChildNodes)
    End Function
    Public Shared Function GetBookNamesByCollection(ByVal Index As Integer) As Array()
        Dim Names() As Array = Array.ConvertAll(Utility.GetChildNodes("book", Utility.GetChildNode("books", CachedData.XMLDocInfos(Index).DocumentElement.ChildNodes).ChildNodes), Function(Convert As System.Xml.XmlNode) New Object() {Convert.Attributes.GetNamedItem("index").Value + ". " + GetBookEName(Convert, Index), CInt(Convert.Attributes.GetNamedItem("index").Value)})
        Array.Sort(Names, New Utility.CompareNameValueArray)
        Return Names
    End Function
    Public Shared Function HasVolumes(ByVal Index As Integer) As Boolean
        Return Not Utility.GetChildNode("books", CachedData.XMLDocInfos(Index).DocumentElement.ChildNodes).Attributes("volumes") Is Nothing
    End Function
    Public Shared Function GetVolumeIndex(ByVal Index As Integer, ByVal BookIndex As Integer) As Integer
        Dim Node As System.Xml.XmlNode = GetBookByIndex(Index, BookIndex).Attributes.GetNamedItem("volume")
        If Node Is Nothing Then Return -1
        Return CInt(Node.Value)
    End Function
    Public Shared Function GetChapterCount(ByVal Index As Integer, ByVal BookIndex As Integer) As Integer
        Dim Node As System.Xml.XmlNode = GetBookByIndex(Index, BookIndex).Attributes.GetNamedItem("chapters")
        If Node Is Nothing Then Return -1
        Return CInt(Node.Value)
    End Function
    Public Shared Function GetChapterIndex(ByVal Index As Integer, ByVal BookIndex As Integer, ByVal HadithIndex As Integer) As Integer
        Dim BookNode As System.Xml.XmlNode = GetBookByIndex(Index, BookIndex)
        Dim ChapterNode As System.Xml.XmlNode
        Dim Count As Integer
        For Count = 0 To BookNode.ChildNodes.Count - 1
            ChapterNode = BookNode.ChildNodes.Item(Count)
            If ChapterNode.Name = "chapter" AndAlso _
                (CInt(ChapterNode.Attributes.GetNamedItem("starthadith").Value) <= HadithIndex And _
                CInt(ChapterNode.Attributes.GetNamedItem("starthadith").Value) + CInt(ChapterNode.Attributes.GetNamedItem("hadiths ").Value) > HadithIndex) Then Return Count
        Next
        Return -1
    End Function
    Public Shared Function GetHadithCount(ByVal Index As Integer, ByVal BookIndex As Integer) As Integer
        Dim Node As System.Xml.XmlNode = GetBookByIndex(Index, BookIndex).Attributes.GetNamedItem("hadiths")
        If Node Is Nothing Then Return -1
        Return CInt(Node.Value)
    End Function
    Public Shared Function GetHadithStart(ByVal Index As Integer, ByVal BookIndex As Integer) As Integer
        Dim Node As System.Xml.XmlNode = GetBookByIndex(Index, BookIndex).Attributes.GetNamedItem("starthadith")
        If Node Is Nothing Then Return -1
        Return CInt(Node.Value)
    End Function
    Public Shared Function ParseBookTranslationIndex(ByVal BookString As String) As Integer()
        Return Array.ConvertAll(BookString.Split("|"c), Function(MakeNumeric As String) Integer.Parse(MakeNumeric))
    End Function
    Public Shared Function ExpandIndexes(ByVal ExpandString As String) As Collections.Generic.List(Of Object)
        Dim Count As Integer
        Dim SubCount As Integer
        Dim Groupings As String() = ExpandString.Split(","c)
        Dim Indexes As New Collections.Generic.List(Of Object)
        For Count = 0 To Groupings.Length - 1
            If Groupings(Count) = String.Empty Then
                Indexes.Add(-1)
            Else
                Dim Ranges As String() = Groupings(Count).Split("-"c)
                If Ranges.Length = 1 Then
                    Dim Combined As String() = Ranges(0).Split("+"c)
                    If Combined.Length = 1 Then
                        Indexes.Add(Integer.Parse(Ranges(0)))
                    Else
                        Indexes.Add(Array.ConvertAll(Combined, Function(MakeNumeric As String) Integer.Parse(MakeNumeric)))
                    End If
                ElseIf Ranges.Length = 2 Then
                    Dim Combined As String() = Ranges(1).Split("+"c)
                    For SubCount = Integer.Parse(Ranges(0)) To Integer.Parse(Combined(0))
                        If Combined.Length > 1 AndAlso SubCount = Integer.Parse(Combined(0)) Then
                            Indexes.Add(Array.ConvertAll(Combined, Function(MakeNumeric As String) Integer.Parse(MakeNumeric)))
                        Else
                            Indexes.Add(SubCount)
                        End If
                    Next
                End If
            End If
        Next
        Return Indexes
    End Function
    Public Shared Function ParseHadithTranslationIndex(ByVal HadithString As String) As Collections.Generic.List(Of Collections.Generic.List(Of Object))
        ParseHadithTranslationIndex = New Collections.Generic.List(Of Collections.Generic.List(Of Object))
        ParseHadithTranslationIndex.AddRange(Array.ConvertAll(HadithString.Split("|"c), Function(IndexString As String) ExpandIndexes(IndexString)))
    End Function
    Public Shared Function TranslationHasVolumes(ByVal XMLDocTranslate As System.Xml.XmlDocument) As Boolean
        Return Not Utility.GetChildNode("books", XMLDocTranslate.DocumentElement.ChildNodes).Attributes("volumes") Is Nothing
    End Function
    Public Shared Function GetTranslateMaxHadith(ByVal XMLDocTranslate As System.Xml.XmlDocument) As Integer
        Dim BookNode As System.Xml.XmlNode
        Dim Count As Integer
        Dim MaxHadith As Integer = 0
        For Count = 0 To Utility.GetChildNode("books", XMLDocTranslate.DocumentElement.ChildNodes).ChildNodes.Count - 1
            BookNode = Utility.GetChildNode("books", XMLDocTranslate.DocumentElement.ChildNodes).ChildNodes.Item(Count)
            If BookNode.Name = "book" Then
                If CInt(BookNode.Attributes.GetNamedItem("hadiths").Value) <> 0 Then
                    MaxHadith = Math.Max(MaxHadith, CInt(BookNode.Attributes.GetNamedItem("starthadith").Value) + CInt(BookNode.Attributes.GetNamedItem("hadiths").Value) - 1)
                End If
            End If
        Next
        Return MaxHadith
    End Function
    Public Shared Function GetMaxChapter(ByVal XMLDocTranslate As System.Xml.XmlDocument) As Integer
        Dim BookNode As System.Xml.XmlNode
        Dim ChapterNode As System.Xml.XmlNode
        Dim Count As Integer
        Dim SubCount As Integer
        Dim MaxChapter As Integer = 0
        For Count = 0 To Utility.GetChildNode("books", XMLDocTranslate.DocumentElement.ChildNodes).ChildNodes.Count - 1
            BookNode = Utility.GetChildNode("books", XMLDocTranslate.DocumentElement.ChildNodes).ChildNodes.Item(Count)
            If BookNode.Name = "book" Then
                For SubCount = 0 To BookNode.ChildNodes.Count - 1
                    ChapterNode = BookNode.ChildNodes.Item(SubCount)
                    If ChapterNode.Name = "chapter" Then
                        MaxChapter = Math.Max(MaxChapter, CInt(ChapterNode.Attributes.GetNamedItem("index").Value))
                    End If
                Next
            End If
        Next
        Return MaxChapter
    End Function
    Public Shared Function GetHadithChapter(ByVal BookNode As System.Xml.XmlNode, ByVal Hadith As Integer) As Integer
        Dim ChapterNode As System.Xml.XmlNode
        Dim Node As System.Xml.XmlNode
        Dim Count As Integer
        For Count = 0 To BookNode.ChildNodes.Count - 1
            ChapterNode = BookNode.ChildNodes.Item(Count)
            If ChapterNode.Name = "chapter" Then
                Node = ChapterNode.Attributes.GetNamedItem("starthadith")
                If Not Node Is Nothing AndAlso (Hadith >= CInt(Node.Value) And _
                    Hadith < CInt(Node.Value) + CInt(ChapterNode.Attributes.GetNamedItem("hadiths").Value)) Then
                    Return CInt(ChapterNode.Attributes.GetNamedItem("index").Value)
                End If
            End If
        Next
        Return -1
    End Function
    Public Shared Function BuildTranslationIndex(ByVal XMLDocTranslate As System.Xml.XmlDocument, ByVal Volume As Integer, ByVal Book As Integer, ByVal Chapter As Integer, ByVal Hadith As Integer, ByVal SharedHadith As Integer) As String
        Dim MaxVolume As Integer = 0
        Dim MaxBook As Integer = CInt(Utility.GetChildNode("books", XMLDocTranslate.DocumentElement.ChildNodes).Attributes.GetNamedItem("count").Value)
        Dim MaxChapter As Integer
        Dim MaxHadith As Integer
        Dim bHasSharedHadith As Boolean = Not Utility.GetChildNode("books", XMLDocTranslate.DocumentElement.ChildNodes).Attributes.GetNamedItem("sharedhadiths") Is Nothing
        MaxHadith = GetTranslateMaxHadith(XMLDocTranslate)
        Dim bHasChapters As Boolean = Not Utility.GetChildNode("books", XMLDocTranslate.DocumentElement.ChildNodes).Attributes.GetNamedItem("chapters") Is Nothing
        If Chapter <> -1 Then
            MaxChapter = GetMaxChapter(XMLDocTranslate)
        End If
        If Volume <> -1 Then
            MaxVolume = CInt(Utility.GetChildNode("books", XMLDocTranslate.DocumentElement.ChildNodes).Attributes.GetNamedItem("volumes").Value)
        End If
        Return CStr(IIf(Volume = -1, String.Empty, Utility.ZeroPad(CStr(Volume), Utility.GetDigitLength(MaxVolume)) + ".")) + Utility.ZeroPad(CStr(Book), Utility.GetDigitLength(MaxBook)) + "." + CStr(IIf(Chapter = -1, String.Empty, Utility.ZeroPad(CStr(Chapter), Utility.GetDigitLength(MaxChapter)) + ".")) + Utility.ZeroPad(CStr(Hadith), Utility.GetDigitLength(MaxHadith)) + CStr(IIf(bHasSharedHadith, IIf(SharedHadith = 0, IIf(Chapter = -1 Or Utility.GetChildNode("books", XMLDocTranslate.DocumentElement.ChildNodes).Attributes.GetNamedItem("sourced") Is Nothing, " ", String.Empty), Chr(64 + SharedHadith)), String.Empty)) + ":"
    End Function
    Public Shared Function MapIndexes(ByVal ExpandString As String, ByVal HadithIndex As Integer) As Object()
        Dim Count As Integer
        Dim SubCount As Integer
        Dim Groupings As String() = ExpandString.Split(","c)
        Dim Indexes As New ArrayList
        Indexes.Add(New String() {})
        Indexes.Add(New String() {String.Empty, String.Empty})
        Indexes.Add(New String() {Utility.LoadResourceString("Hadith_SourceHadithIndex"), Utility.LoadResourceString("Hadith_TranslationHadithIndex")})
        Dim HadithCount As Integer = 0
        For Count = 0 To Groupings.Length - 1
            If Groupings(Count) = String.Empty Then
                Indexes.Add(New String() {CStr(HadithIndex), Utility.LoadResourceString("Hadith_NoTranslation")})
                HadithIndex += 1
            Else
                Dim Ranges As String() = Groupings(Count).Split("-"c)
                If Ranges.Length = 1 Then
                    Dim Compile As String = String.Empty
                    Dim Combined As String() = Ranges(0).Split("+"c)
                    For SubCount = 0 To Combined.Length - 1
                        Compile += Combined(SubCount) + CStr(IIf(SubCount <> Combined.Length - 1, "&", String.Empty))
                    Next
                    Indexes.Add(New String() {CStr(HadithIndex), Compile})
                    HadithIndex += 1
                    HadithCount += Combined.Length
                ElseIf Ranges.Length = 2 Then
                    Dim Compile As String
                    Dim Combined As String() = Ranges(1).Split("+"c)
                    Compile = Ranges(0) + "-"
                    For SubCount = 0 To Combined.Length - 1
                        Compile += Combined(SubCount) + CStr(IIf(SubCount <> Combined.Length - 1, "&", String.Empty))
                    Next
                    Indexes.Add(New String() {CStr(HadithIndex) + "-" + CStr(HadithIndex + Integer.Parse(Combined(0)) - Integer.Parse(Ranges(0))), Compile})
                    HadithIndex += Integer.Parse(Combined(0)) - Integer.Parse(Ranges(0)) + 1
                    HadithCount += Integer.Parse(Combined(0)) - Integer.Parse(Ranges(0)) + Combined.Length
                End If
            End If
        Next
        Return New Object() {CStr(HadithCount), Indexes.ToArray(GetType(Array))}
    End Function
    Public Shared Function GetBookHadithMapping(ByVal XMLDocTranslate As System.Xml.XmlDocument, ByVal Index As Integer, ByVal BookIndex As Integer) As Object()
        Dim Count As Integer
        Dim SourceStart As Integer
        Dim Volume As Integer
        Dim Mapping As New ArrayList
        Mapping.Add(New String() {})
        If TranslationHasVolumes(XMLDocTranslate) Then
            Mapping.Add(New String() {String.Empty, String.Empty, String.Empty, String.Empty})
            Mapping.Add(New String() {Utility.LoadResourceString("Hadith_TranslationVolume"), Utility.LoadResourceString("Hadith_TranslationBook"), Utility.LoadResourceString("Hadith_TranslationHadithCount"), Utility.LoadResourceString("Hadith_TranslationMapping")})
        Else
            Mapping.Add(New String() {String.Empty, String.Empty, String.Empty})
            Mapping.Add(New String() {Utility.LoadResourceString("Hadith_TranslationBook"), Utility.LoadResourceString("Hadith_TranslationHadithCount"), Utility.LoadResourceString("Hadith_TranslationMapping")})
        End If
        Dim bHasSharedHadith As Boolean = Not Utility.GetChildNode("books", XMLDocTranslate.DocumentElement.ChildNodes).Attributes.GetNamedItem("sharedhadiths") Is Nothing
        Dim Books() As Integer
        Dim TranslateBookIndex As Integer
        Dim BookNode As System.Xml.XmlNode
        Dim Node As System.Xml.XmlNode
        For Count = 0 To Utility.GetChildNode("books", XMLDocTranslate.DocumentElement.ChildNodes).ChildNodes.Count - 1
            BookNode = Utility.GetChildNode("books", XMLDocTranslate.DocumentElement.ChildNodes).ChildNodes.Item(Count)
            If BookNode.Name = "book" Then
                Node = BookNode.Attributes.GetNamedItem("sourcebook")
                If Node Is Nothing Then
                    Books = New Integer() {CInt(BookNode.Attributes.GetNamedItem("index").Value)}
                Else
                    Books = ParseBookTranslationIndex(Node.Value)
                End If
                Node = BookNode.Attributes.GetNamedItem("volume")
                If Node Is Nothing Then
                    Volume = -1
                Else
                    Volume = CInt(Node.Value)
                End If
                TranslateBookIndex = Array.FindIndex(Books, Function(CheckIndex As Integer) CheckIndex = BookIndex)
                If TranslateBookIndex <> -1 Then
                    'Must handle shared hadiths
                    Node = BookNode.Attributes.GetNamedItem("sourcestart")
                    If Node Is Nothing Then
                        If CInt(BookNode.Attributes.GetNamedItem("hadiths").Value) = 0 Then
                            SourceStart = 0
                        Else
                            SourceStart = CInt(BookNode.Attributes.GetNamedItem("starthadith").Value)
                        End If
                    Else
                        SourceStart = CInt(Node.Value.Split("|"c)(TranslateBookIndex))
                    End If
                    Node = BookNode.Attributes.GetNamedItem("sourceindex")
                    If Node Is Nothing OrElse Node.Value = String.Empty Then
                        If TranslationHasVolumes(XMLDocTranslate) Then
                            Mapping.Add(New String() {CStr(Volume), BookNode.Attributes.GetNamedItem("index").Value, _
                                        CStr(BookNode.Attributes.GetNamedItem("hadiths").Value), Utility.LoadResourceString("Hadith_IdenticalNumbering")})
                        Else
                            Mapping.Add(New String() {BookNode.Attributes.GetNamedItem("index").Value, _
                                        CStr(BookNode.Attributes.GetNamedItem("hadiths").Value), Utility.LoadResourceString("Hadith_IdenticalNumbering")})
                        End If
                    Else
                        Dim RetObject As Object()
                        RetObject = MapIndexes(Node.Value.Split("|"c)(TranslateBookIndex), SourceStart)
                        Dim SharedHadith As Integer
                        If bHasSharedHadith Then
                            SharedHadith = CInt(Utility.ParseValue(BookNode.Attributes.GetNamedItem("sharedhadiths"), "0"))
                        Else
                            SharedHadith = 0
                        End If
                        If TranslationHasVolumes(XMLDocTranslate) Then
                            Mapping.Add(New Object() {CStr(Volume), BookNode.Attributes.GetNamedItem("index").Value, _
                                        String.Format(Utility.LoadResourceString("Hadith_MappedOf"), CStr(RetObject(0)), CStr(CInt(BookNode.Attributes.GetNamedItem("hadiths").Value) + SharedHadith)), RetObject(1)})
                        Else
                            Mapping.Add(New Object() {BookNode.Attributes.GetNamedItem("index").Value, _
                                        String.Format(Utility.LoadResourceString("Hadith_MappedOf"), CStr(RetObject(0)), CStr(CInt(BookNode.Attributes.GetNamedItem("hadiths").Value) + SharedHadith)), RetObject(1)})
                        End If
                    End If
                End If
            End If
        Next
        Return Mapping.ToArray()
    End Function
    Public Shared Function GetSharedHadithIndex(ByVal TranslationIndexes As Collections.Generic.List(Of Collections.Generic.List(Of Object)), ByVal TranslateBookIndex As Integer, ByVal HadithIndex As Integer, ByVal SubCount As Integer) As Integer
        Dim Count As Integer
        Dim HadithCount As Integer
        Dim ChildCount As Integer
        Dim HadithValue As Integer
        Dim SharedHadith As Integer = -1
        If TypeOf TranslationIndexes(TranslateBookIndex).Item(HadithIndex) Is Integer() Then
            HadithValue = DirectCast(TranslationIndexes(TranslateBookIndex)(HadithIndex), Integer())(SubCount)
        Else
            HadithValue = CInt(TranslationIndexes(TranslateBookIndex)(HadithIndex))
        End If
        For Count = 0 To TranslateBookIndex
            For HadithCount = 0 To CInt(IIf(Count = TranslateBookIndex, HadithIndex, TranslationIndexes(Count).Count - 1))
                If TypeOf TranslationIndexes(Count)(HadithCount) Is Integer() Then
                    For ChildCount = 0 To CInt(IIf(Count = TranslateBookIndex And HadithCount = HadithIndex, SubCount - 1, DirectCast(TranslationIndexes(Count)(HadithCount), Integer()).Length - 1))
                        If DirectCast(TranslationIndexes(Count)(HadithCount), Integer())(ChildCount) = HadithValue Then
                            SharedHadith += 1
                        End If
                    Next
                Else
                    If CInt(TranslationIndexes(Count)(HadithCount)) = HadithValue Then
                        SharedHadith += 1
                    End If
                End If
            Next
        Next
        Return SharedHadith
    End Function
    Public Shared Function GetTranslationHadith(XMLDocTranslate As System.Xml.XmlDocument, Strings() As String, ByVal Index As Integer, ByVal BookIndex As Integer, ByVal HadithIndex As Integer) As String()
        Dim BookNode As System.Xml.XmlNode
        Dim Node As System.Xml.XmlNode
        Dim Count As Integer
        Dim SubCount As Integer
        Dim SourceStart As Integer
        Dim Volume As Integer
        Dim Books() As Integer
        Dim TranslateBookIndex As Integer
        Dim TranslationIndexes As Collections.Generic.List(Of Collections.Generic.List(Of Object))
        Dim TranslationHadith As New ArrayList
        If CachedData.IslamData.Collections(Index).Translations.Length = 0 Then Return New String() {}
        For Count = 0 To Utility.GetChildNode("books", XMLDocTranslate.DocumentElement.ChildNodes).ChildNodes.Count - 1
            BookNode = Utility.GetChildNode("books", XMLDocTranslate.DocumentElement.ChildNodes).ChildNodes.Item(Count)
            If BookNode.Name = "book" Then
                Node = BookNode.Attributes.GetNamedItem("sourcebook")
                If Node Is Nothing Then
                    Books = New Integer() {Count + 1}
                Else
                    Books = ParseBookTranslationIndex(Node.Value)
                End If
                Node = BookNode.Attributes.GetNamedItem("volume")
                If Node Is Nothing Then
                    Volume = -1
                Else
                    Volume = CInt(Node.Value)
                End If
                TranslateBookIndex = Array.FindIndex(Books, Function(CheckIndex As Integer) CheckIndex = BookIndex + 1)
                If TranslateBookIndex <> -1 Then
                    Node = BookNode.Attributes.GetNamedItem("sourcestart")
                    If Node Is Nothing Then
                        SourceStart = HadithIndex
                    Else
                        SourceStart = CInt(Node.Value.Split("|"c)(TranslateBookIndex))
                    End If
                    Node = BookNode.Attributes.GetNamedItem("sourceindex")
                    If Node Is Nothing Then
                        TranslationHadith.Add(CStr(IIf(Volume = -1, String.Empty, Utility.LoadResourceString("Hadith_Volume") + ": " + CStr(Volume) + " ")) + Utility.LoadResourceString("Hadith_Books") + ": " + CStr(Books(TranslateBookIndex)) + " " + Utility.LoadResourceString("Hadith_Hadith") + ": " + CStr(HadithIndex))
                        TranslationHadith.AddRange(Utility.GetFileLinesByNumberPrefix(Strings, BuildTranslationIndex(XMLDocTranslate, Volume, CInt(BookNode.Attributes.GetNamedItem("index").Value), GetHadithChapter(BookNode, HadithIndex), HadithIndex, 0)))
                    Else
                        TranslationIndexes = ParseHadithTranslationIndex(Node.Value)
                        If HadithIndex >= SourceStart AndAlso HadithIndex - SourceStart < TranslationIndexes(TranslateBookIndex).Count Then
                            If TypeOf TranslationIndexes(TranslateBookIndex)(HadithIndex - SourceStart) Is Integer() Then
                                For SubCount = 0 To DirectCast(TranslationIndexes(TranslateBookIndex)(HadithIndex - SourceStart), Integer()).Length - 1
                                    Dim SharedHadithIndex As Integer = GetSharedHadithIndex(TranslationIndexes, TranslateBookIndex, HadithIndex - SourceStart, SubCount)
                                    TranslationHadith.Add(CStr(IIf(Volume = -1, String.Empty, Utility.LoadResourceString("Hadith_Volume") + ": " + CStr(Volume) + " ")) + Utility.LoadResourceString("Hadith_Book") + ": " + BookNode.Attributes.GetNamedItem("index").Value + " " + Utility.LoadResourceString("Hadith_Hadith") + ": " + CStr(DirectCast(TranslationIndexes(TranslateBookIndex)(HadithIndex - SourceStart), Integer())(SubCount)))
                                    TranslationHadith.AddRange(Utility.GetFileLinesByNumberPrefix(Strings, BuildTranslationIndex(XMLDocTranslate, Volume, CInt(BookNode.Attributes.GetNamedItem("index").Value), GetHadithChapter(BookNode, DirectCast(TranslationIndexes(TranslateBookIndex)(HadithIndex - SourceStart), Integer())(SubCount)), DirectCast(TranslationIndexes(TranslateBookIndex)(HadithIndex - SourceStart), Integer())(SubCount), SharedHadithIndex)))
                                Next
                            Else
                                If CInt(TranslationIndexes(TranslateBookIndex)(HadithIndex - SourceStart)) = -1 Then Return New String() {}
                                Dim SharedHadithIndex As Integer = GetSharedHadithIndex(TranslationIndexes, TranslateBookIndex, HadithIndex - SourceStart, -1)
                                TranslationHadith.Add(CStr(IIf(Volume = -1, String.Empty, Utility.LoadResourceString("Hadith_Volume") + ": " + CStr(Volume) + " ")) + Utility.LoadResourceString("Hadith_Book") + ": " + BookNode.Attributes.GetNamedItem("index").Value + " " + Utility.LoadResourceString("Hadith_Hadith") + ": " + CStr(TranslationIndexes(TranslateBookIndex)(HadithIndex - SourceStart)))
                                TranslationHadith.AddRange(Utility.GetFileLinesByNumberPrefix(Strings, BuildTranslationIndex(XMLDocTranslate, Volume, CInt(BookNode.Attributes.GetNamedItem("index").Value), GetHadithChapter(BookNode, CInt(TranslationIndexes(TranslateBookIndex)(HadithIndex - SourceStart))), CInt(TranslationIndexes(TranslateBookIndex)(HadithIndex - SourceStart)), SharedHadithIndex)))
                            End If
                        End If
                    End If
                End If
            End If
        Next
        Return DirectCast(TranslationHadith.ToArray(GetType(String)), String())
    End Function
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
        ExecuteNonQuery(Connection, "CREATE TABLE HadithRankings (UserID int NOT NULL, " + _
        "Collection VARCHAR(254) NOT NULL, " + _
        "BookIndex int, " + _
        "HadithIndex int NOT NULL, " + _
        "Ranking int NOT NULL)")
        Connection.Close()
    End Sub
    Public Shared Sub RemoveDatabase()
        Dim Connection As MySql.Data.MySqlClient.MySqlConnection = GetConnection()
        If Connection Is Nothing Then Return
        Dim Command As MySql.Data.MySqlClient.MySqlCommand = Connection.CreateCommand()
        ExecuteNonQuery(Connection, "DROP TABLE Users")
        ExecuteNonQuery(Connection, "DROP TABLE HadithRankings")
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
    Public Shared Function GetHadithCollectionRankingData(ByVal Collection As String) As Double
        Dim Connection As MySql.Data.MySqlClient.MySqlConnection = GetConnection()
        If Connection Is Nothing Then Return -1
        Dim Command As MySql.Data.MySqlClient.MySqlCommand = Connection.CreateCommand()
        Command.CommandText = "SELECT AVG(Ranking) AS AvgRank FROM HadithRankings Collection=@Collection"
        Command.Parameters.AddWithValue("@Collection", Collection)
        Dim Reader As MySql.Data.MySqlClient.MySqlDataReader = Command.ExecuteReader()
        If Reader.Read() AndAlso Not Reader.IsDBNull(0) Then
            GetHadithCollectionRankingData = Reader.GetInt32("AvgRank")
        Else
            GetHadithCollectionRankingData = -1
        End If
        Reader.Close()
        Connection.Close()
    End Function
    Public Shared Function GetHadithBookRankingData(ByVal Collection As String, ByVal Book As Integer) As Double
        Dim Connection As MySql.Data.MySqlClient.MySqlConnection = GetConnection()
        If Connection Is Nothing Then Return -1
        Dim Command As MySql.Data.MySqlClient.MySqlCommand = Connection.CreateCommand()
        Command.CommandText = "SELECT AVG(Ranking) AS AvgRank FROM HadithRankings WHERE Collection=@Collection AND BookIndex=" + CStr(Book)
        Command.Parameters.AddWithValue("@Collection", Collection)
        Dim Reader As MySql.Data.MySqlClient.MySqlDataReader = Command.ExecuteReader()
        If Reader.Read() AndAlso Not Reader.IsDBNull(0) Then
            GetHadithBookRankingData = Reader.GetInt32("AvgRank")
        Else
            GetHadithBookRankingData = -1
        End If
        Reader.Close()
        Connection.Close()
    End Function
    Public Shared Function GetHadithRankingData(ByVal Collection As String, ByVal Book As Integer, ByVal Hadith As Integer) As Integer()
        Dim Connection As MySql.Data.MySqlClient.MySqlConnection = GetConnection()
        If Connection Is Nothing Then Return {0, 0}
        Dim Command As MySql.Data.MySqlClient.MySqlCommand = Connection.CreateCommand()
        Command.CommandText = "SELECT SUM(Ranking) AS SumRank, COUNT(Ranking) AS CountRank FROM HadithRankings WHERE Collection=@Collection AND BookIndex=" + CStr(Book) + " AND HadithIndex=" + CStr(Hadith)
        Command.Parameters.AddWithValue("@Collection", Collection)
        Dim Reader As MySql.Data.MySqlClient.MySqlDataReader = Command.ExecuteReader()
        If Reader.Read() AndAlso Not Reader.IsDBNull(0) Then
            GetHadithRankingData = {Reader.GetInt32("SumRank"), Reader.GetInt32("CountRank")}
        Else
            GetHadithRankingData = {0, 0}
        End If
        Reader.Close()
        Connection.Close()
    End Function
    Public Shared Function GetUserHadithRankingData(ByVal UserID As Integer, ByVal Collection As String, ByVal Book As Integer, ByVal Hadith As Integer) As Integer
        Dim Connection As MySql.Data.MySqlClient.MySqlConnection = GetConnection()
        If Connection Is Nothing Then Return -1
        Dim Command As MySql.Data.MySqlClient.MySqlCommand = Connection.CreateCommand()
        Command.CommandText = "SELECT Ranking FROM HadithRankings WHERE UserID=" + CStr(UserID) + " AND Collection=@Collection AND BookIndex=" + CStr(Book) + " AND HadithIndex=" + CStr(Hadith)
        Command.Parameters.AddWithValue("@Collection", Collection)
        Dim Reader As MySql.Data.MySqlClient.MySqlDataReader = Command.ExecuteReader()
        If Reader.Read() AndAlso Not Reader.IsDBNull(0) Then
            GetUserHadithRankingData = Reader.GetInt32("Ranking")
        Else
            GetUserHadithRankingData = -1
        End If
        Reader.Close()
        Connection.Close()
    End Function
    Public Shared Sub SetUserHadithRankingData(ByVal UserID As Integer, ByVal Collection As String, ByVal Book As Integer, ByVal Hadith As Integer, ByVal Rank As Integer)
        Dim Connection As MySql.Data.MySqlClient.MySqlConnection = GetConnection()
        If Connection Is Nothing Then Return
        ExecuteNonQuery(Connection, "INSERT INTO HadithRankings (UserID, Collection, BookIndex, HadithIndex, Ranking) VALUES (" + CStr(UserID) + ", @Collection, " + CStr(Book) + ", " + CStr(Hadith) + ", " + CStr(Rank) + ")", New Generic.Dictionary(Of String, Object) From {{"@Collection", Collection}})
        Connection.Close()
    End Sub
    Public Shared Sub UpdateUserHadithRankingData(ByVal UserID As Integer, ByVal Collection As String, ByVal Book As Integer, ByVal Hadith As Integer, ByVal Rank As Integer)
        Dim Connection As MySql.Data.MySqlClient.MySqlConnection = GetConnection()
        If Connection Is Nothing Then Return
        ExecuteNonQuery(Connection, "UPDATE HadithRankings SET Ranking=" + CStr(Rank) + " WHERE UserID=" + CStr(UserID) + " AND Collection=@Collection AND BookIndex=" + CStr(Book) + " AND HadithIndex=" + CStr(Hadith), New Generic.Dictionary(Of String, Object) From {{"@Collection", Collection}})
        Connection.Close()
    End Sub
    Public Shared Sub RemoveUserHadithRankingData(ByVal UserID As Integer, ByVal Collection As String, ByVal Book As Integer, ByVal Hadith As Integer)
        Dim Connection As MySql.Data.MySqlClient.MySqlConnection = GetConnection()
        If Connection Is Nothing Then Return
        ExecuteNonQuery(Connection, "DELETE FROM HadithRankings WHERE UserID=" + CStr(UserID) + " AND Collection=@Collection AND BookIndex=" + CStr(Book) + " AND HadithIndex=" + CStr(Hadith), New Generic.Dictionary(Of String, Object) From {{"@Collection", Collection}})
        Connection.Close()
    End Sub
End Class