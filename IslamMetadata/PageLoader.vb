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
    'Web.HttpContext.Current.Trace.Write(Text)
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
        Return New Reflection.AssemblyName(Reflection.Assembly.GetEntryAssembly().FullName).Name = "IslamSource"
    End Function
    Public Shared Function GetTemplatePath() As String
        If IsDesktopApp() Then
            Return "..\..\..\metadata\IslamSource.xml"
        Else
            Dim Index As Integer = Array.FindIndex(ConnectionData.SiteDomains(), Function(Domain As String) Web.HttpContext.Current.Request.Url.Host.EndsWith(Domain))
            If Index = -1 Then
                Return GetFilePath("metadata\" + ConnectionData.DefaultXML + ".xml")
            Else
                Return GetFilePath("metadata\" + ConnectionData.SiteXMLs()(Index) + ".xml")
            End If
        End If
    End Function
    Public Shared Function GetFilePath(ByVal Path As String) As String
        If IsDesktopApp() Then
            Return Path
        Else
            Return CStr(IIf(IO.File.Exists(Web.HttpContext.Current.Request.PhysicalApplicationPath + Path), Web.HttpContext.Current.Request.PhysicalApplicationPath + Path, Web.HttpContext.Current.Request.PhysicalApplicationPath + ConnectionData.AlternatePath + Path))
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
            LoadResourceString = New System.Resources.ResourceManager("GMorseCodeResources.Resources", Reflection.Assembly.Load("GMorseCodeResources")).GetString(resourceKey, Threading.Thread.CurrentThread.CurrentUICulture)
            'LoadResourceString = CStr(HttpContext.GetGlobalResourceObject(ConnectionData.GlobalRes, resourceKey))
        End If
        If LoadResourceString = Nothing Then
            LoadResourceString = String.Empty
            System.Diagnostics.Debug.WriteLine("  <data name=""" + resourceKey + """ xml:space=""preserve"">" + vbCrLf + "    <value>" + System.Text.RegularExpressions.Regex.Replace(System.Text.RegularExpressions.Regex.Replace(resourceKey, ".*_", String.Empty), "(.+?)([A-Z])", "$1 $2") + "</value>" + vbCrLf + "  </data>")
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
        cspParams.Flags = System.Security.Cryptography.CspProviderFlags.NoFlags
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
    Public Shared Function EncodeJS(Str As String) As String
        Return Str.Replace("'", "\'")
    End Function
    Public Shared Function MakeJSArray(ByVal StringArray As String(), Optional ByVal bObject As Boolean = False) As String
        Dim JSArray As String = "["
        Dim Count As Integer
        For Count = 0 To StringArray.Length() - 1
            If bObject Then
                JSArray += StringArray(Count)
            Else
                JSArray += "'" + EncodeJS(StringArray(Count)) + "'"
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
                If bObject Then
                    JSArray += CType(StringsArray(Count), String())(SubCount)
                Else
                    JSArray += "'" + EncodeJS(CStr(CType(StringsArray(Count), String())(SubCount))) + "'"
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
        Dim File As IO.FileStream = IO.File.Open(IO.Path.Combine(GetCacheDirectory(), Name), IO.FileMode.Open)
        Dim Bytes(CInt(File.Length) - 1) As Byte
        File.Read(Bytes, 0, CInt(File.Length))
        File.Close()
        Return Bytes
    End Function
    Public Shared Function TransmitCacheItem(ByVal Name As String, ByVal ModifiedUtc As Date) As Boolean
        If Not IO.File.Exists(IO.Path.Combine(GetCacheDirectory, Name)) OrElse _
            ModifiedUtc > IO.File.GetLastWriteTimeUtc(IO.Path.Combine(GetCacheDirectory(), Name)) Then Return False
        Web.HttpContext.Current.Response.TransmitFile(IO.Path.Combine(GetCacheDirectory(), Name))
        Return True
    End Function
    Public Shared Sub CacheItem(ByVal Name As String, ByVal LastModifiedUtc As Date, ByVal Data() As Byte)
        Dim File As IO.FileStream = IO.File.Open(IO.Path.Combine(GetCacheDirectory(), Name), IO.FileMode.Create)
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
        RetArray(0) = New String() {"javascript: doOnCheck(this);", "doSort();", "function doSort() { var child = $('#render').children('tr'); child.shift(); child.sort(function(a, b) { if (window.localstorage.getItem(a.children('td')(3).text()) == window.localstorage.getItem(b.children('td')(3).text())) return new Date(a.children('td')(1).text()) > new Data(b.children('td')(1).text()); return (window.localstorage.getItem(a.children('td')(3).text())) ? 1 : -1; }); child.detach().appendTo($('#render')); }" + "function doOnCheck(element) { element.checked = !element.checked; if (element.checked) { window.localstorage.setItem($(element).parent().children('td')(3), true); } else { window.localstorage.removeItem($(element).parent().children('td')(3)); } doSort(); }"}
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
        Return Web.HttpContext.Current.Request.UserHostAddress
    End Function
    Public Shared Function GetGeoData() As String()
        Dim URL As String = "http://api.ipinfodb.com/v3/ip-city/?key=" + Utility.ConnectionData.IPInfoDBAPIKey + "&ip=" + Web.HttpContext.Current.Request.UserHostAddress
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
        Public Function Compare(ByVal x As Object, ByVal y As Object) As Integer _
            Implements Collections.IComparer.Compare
            Compare = DirectCast(x, IslamData.ArabicSymbol).RomanTranslit.Length - _
                DirectCast(y, IslamData.ArabicSymbol).RomanTranslit.Length
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
            For Index = 0 To CachedData.IslamData.ArabicLetters.Length - 1
                If Buckwalter(Count) = CachedData.IslamData.ArabicLetters(Index).ExtendedBuckwalterLetter Then
                    ArabicString += CachedData.IslamData.ArabicLetters(Index).Symbol
                    Exit For
                End If
            Next
            If Index = CachedData.IslamData.ArabicLetters.Length Then
                ArabicString += Buckwalter(Count)
            End If
        Next
        Return ArabicString
    End Function
    Public Shared Function TransliterateToScheme(ByVal ArabicString As String) As String
        Dim Scheme As Integer = CInt(HttpContext.Current.Request.QueryString.Get("translitscheme"))
        If Scheme = 1 Then
            Return TransliterateToPlainRoman(ArabicString)
        ElseIf Scheme = 2 Then
            Return TransliterateToRoman(ArabicString, False)
        ElseIf Scheme = 3 Then
            Return TransliterateToRoman(ArabicString, True)
        Else
            Return New String(System.Array.FindAll(ArabicString.ToCharArray(), Function(Check As Char) Check = " "c))
        End If
    End Function
    Public Shared Function TransliterateToRoman(ByVal ArabicString As String, UseBuckwalter As Boolean) As String
        Dim RomanString As String = String.Empty
        Dim Count As Integer
        Dim Index As Integer
        Dim Letters(CachedData.IslamData.ArabicLetters.Length - 1) As IslamData.ArabicSymbol
        CachedData.IslamData.ArabicLetters.CopyTo(Letters, 0)
        Array.Sort(Letters, New StringLengthComparer)
        For Count = 0 To ArabicString.Length - 1
            For Index = 0 To Letters.Length - 1
                If ArabicString(Count) = Letters(Index).Symbol Then
                    RomanString += CStr(IIf(UseBuckwalter, Letters(Index).ExtendedBuckwalterLetter, Letters(Index).RomanTranslit))
                    Exit For
                End If
            Next
            If Index = Letters.Length Then
                RomanString += ArabicString(Count)
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
    Public Const FullStop As Char = Chr(46)
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
    Public Const LeftToRightMark As Char = ChrW(&H200E)
    Public Const RightToLeftMark As Char = ChrW(&H200F)
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
        Return Array.ConvertAll(RecitationSymbols, Function(Ch As Char) New Object() {CachedData.IslamData.ArabicLetters(FindLetterBySymbol(Ch)).UnicodeName + " (" + Ch + ")", FindLetterBySymbol(Ch)})
    End Function
    Public Shared Function IsLetter(Index As Integer) As Boolean
        Return Array.FindIndex(ArabicLetters, Function(Str As String) Str = CachedData.IslamData.ArabicLetters(Index).Symbol) <> -1
    End Function
    Public Shared Function IsPunctuation(Index As Integer) As Boolean
        Return CachedData.IslamData.ArabicLetters(Index).Symbol = ExclamationMark Or CachedData.IslamData.ArabicLetters(Index).Symbol = QuotationMark Or CachedData.IslamData.ArabicLetters(Index).Symbol = FullStop Or CachedData.IslamData.ArabicLetters(Index).Symbol = Comma Or CachedData.IslamData.ArabicLetters(Index).Symbol = ArabicComma
    End Function
    Public Shared Function IsStop(Index As Integer) As Boolean
        Return Array.FindIndex(ArabicStopLetters, Function(Str As String) Str = CachedData.IslamData.ArabicLetters(Index).Symbol) <> -1
    End Function
    Public Shared Function IsIgnored(Index As Integer) As Boolean
        Return (CachedData.IslamData.ArabicLetters(Index).Symbol = ArabicSmallWaw) Or _
            (CachedData.IslamData.ArabicLetters(Index).Symbol = ArabicSmallHighMeemIsolatedForm)
    End Function
    Public Shared Function IsWhitespace(Index As Integer) As Boolean
        Return CachedData.IslamData.ArabicLetters(Index).Symbol = Space
    End Function
    Public Shared Function IsAmbiguousGutteral(Index As Integer, After As Boolean) As Boolean
        Return CachedData.IslamData.ArabicLetters(Index).Symbol = ArabicLetterHamza Or _
            CachedData.IslamData.ArabicLetters(Index).Symbol = ArabicLetterHah Or _
            Not After And (CachedData.IslamData.ArabicLetters(Index).Symbol = ArabicLetterSad Or _
                CachedData.IslamData.ArabicLetters(Index).Symbol = ArabicLetterDad Or _
                CachedData.IslamData.ArabicLetters(Index).Symbol = ArabicLetterTah Or _
                CachedData.IslamData.ArabicLetters(Index).Symbol = ArabicLetterZah) Or _
                CachedData.IslamData.ArabicLetters(Index).Symbol = ArabicLetterAin
    End Function
    Public Shared ArabicUniqueLetters As String() = {"Al^m^", "Al^m^S^", "Al^r", "Al^m^r", "k^hyE^S^", "Th", "Ts^m^", "Ts^", "ys^", "S^", "Hm^", "E^s^q^", "q^", "n^"}
    Public Shared ArabicWaslKasraExceptions As String() = {"{mo$uwA", "{}otuwA", "{qoDuwA", "{bonuwA", "{moDuwA", "{mora>ata", "{somu", "{vonatayni", "{vonayni", "{bonatu", "{bonu", "{moru&NA"}
    Public Shared ArabicFathaDammaKasra As String() = {ArabicFatha, ArabicDamma, ArabicKasra}
    Public Shared ArabicTanweens As String() = {ArabicFathatan, ArabicDammatan, ArabicKasratan}
    Public Shared ArabicLongVowels As String() = {ArabicFatha + ArabicLetterAlef, ArabicDamma + ArabicLetterWaw, ArabicKasra + ArabicLetterYeh}
    Public Shared ArabicSunLetters As String() = {ArabicLetterTeh, ArabicLetterTheh, ArabicLetterDal, ArabicLetterThal, ArabicLetterReh, ArabicLetterZain, ArabicLetterSeen, ArabicLetterSheen, ArabicLetterSad, ArabicLetterDad, ArabicLetterTah, ArabicLetterZah, ArabicLetterLam, ArabicLetterNoon}
    Public Shared ArabicMoonLetters As String() = {ArabicLetterAlef, ArabicLetterBeh, ArabicLetterJeem, ArabicLetterHah, ArabicLetterKhah, ArabicLetterAin, ArabicLetterGhain, ArabicLetterFeh, ArabicLetterQaf, ArabicLetterKaf, ArabicLetterMeem, ArabicLetterHeh, ArabicLetterWaw, ArabicLetterYeh}
    Public Shared ArabicMoonLettersNoVowels As String() = {ArabicLetterBeh, ArabicLetterJeem, ArabicLetterHah, ArabicLetterKhah, ArabicLetterAin, ArabicLetterGhain, ArabicLetterFeh, ArabicLetterQaf, ArabicLetterKaf, ArabicLetterMeem, ArabicLetterHeh}
    Public Shared ArabicSpecialLeadingGutteral As String() = {ArabicLetterHah, ArabicLetterAin}
    Public Shared ArabicSpecialGutteral As String() = {ArabicLetterHamza, ArabicLetterHah, ArabicLetterAin, ArabicLetterSad, ArabicLetterDad, ArabicLetterTah, ArabicLetterZah}
    Public Shared ArabicLetters As String() = {ArabicLetterTeh, ArabicLetterTheh, ArabicLetterDal, ArabicLetterThal, ArabicLetterReh, ArabicLetterZain, ArabicLetterSeen, ArabicLetterSheen, ArabicLetterSad, ArabicLetterDad, ArabicLetterTah, ArabicLetterZah, ArabicLetterLam, ArabicLetterNoon, ArabicLetterAlef, ArabicLetterBeh, ArabicLetterJeem, ArabicLetterHah, ArabicLetterKhah, ArabicLetterAin, ArabicLetterGhain, ArabicLetterFeh, ArabicLetterQaf, ArabicLetterKaf, ArabicLetterMeem, ArabicLetterHeh, ArabicLetterWaw, ArabicLetterYeh}
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
    Delegate Function RuleMetadataEvaluator(Match As System.Text.RegularExpressions.Match) As RuleMetadata()
    Structure RuleMetadataTranslation
        Public Rule As String
        Public RuleName As String
        Public Match As String
        Public Evaluator As RuleMetadataEvaluator
    End Structure
    Structure RuleTranslation
        Public Rule As String
        Public RuleName As String
        Public Match As String
        Public Evaluator As System.Text.RegularExpressions.MatchEvaluator
    End Structure
    Structure ColorRule
        Public Rule As String
        Public RuleName As String
        Public Match As String
        Public Color As Color
    End Structure
    Public Shared ColoringRules As ColorRule() = { _
        New ColorRule With {.RuleName = "Normal", .Match = "empty|helperfatha|helperkasra|helperdamma|helpermeem|assimilator|assimilatorincomplete|dipthong|compulsorystop|endofversestop|prostration|canstoporcontinue|betternottostop|stopatfirstnotsecond|stopatsecondnotfirst|bettertostopbutpermissibletocontinue|bettertocontinuebutpermissibletostop|subtlestopwithoutbreath", .Color = Color.Black}, _
        New ColorRule With {.RuleName = "NecessaryProlongation", .Match = "necessaryprolong", .Color = Color.DarkRed}, _
        New ColorRule With {.RuleName = "ObligatoryProlongation", .Match = "obligatoryprolong", .Color = Color.FromArgb(175, 17, 28)}, _
        New ColorRule With {.RuleName = "PermissibleProlongation", .Match = "permissibleprolong", .Color = Color.OrangeRed}, _
        New ColorRule With {.RuleName = "NormalProlongation", .Match = "normalprolong", .Color = Color.FromArgb(213, 139, 24)}, _
        New ColorRule With {.RuleName = "Nasalization", .Match = "nasalize", .Color = Color.Green}, _
        New ColorRule With {.RuleName = "Unannounced", .Match = "assimilate|assimilateincomplete", .Color = Color.Gray}, _
        New ColorRule With {.RuleName = "EmphaticPronounciation", .Match = "emphasis", .Color = Color.DarkBlue}, _
        New ColorRule With {.RuleName = "UnrestLetters", .Match = "bounce", .Color = Color.Blue} _
    }
    Public Shared BreakdownRules As RuleTranslation() = { _
            New RuleTranslation With {.RuleName = "AlefMaddah", .Match = MakeUniRegEx(ArabicLetterAlefWithMaddaAbove), _
                                      .Evaluator = Function(Match As System.Text.RegularExpressions.Match) ArabicLetterAlef + ArabicMaddahAbove}, _
            New RuleTranslation With {.RuleName = "Maddah", .Match = MakeUniRegEx(ArabicMaddahAbove), _
                .Evaluator = Function(Match As System.Text.RegularExpressions.Match) ArabicLetterHamza + ArabicFatha + ArabicLetterAlef}, _
            New RuleTranslation With {.RuleName = "AlefHamzah", .Match = MakeUniRegEx(ArabicLetterAlefWithHamzaAbove) + "|" + MakeUniRegEx(ArabicLetterAlefWithHamzaBelow), _
                                      .Evaluator = Function(Match As System.Text.RegularExpressions.Match) ArabicLetterAlef + ArabicLetterHamza}, _
            New RuleTranslation With {.RuleName = "WawHamzah", .Match = MakeUniRegEx(ArabicLetterWawWithHamzaAbove), _
                                      .Evaluator = Function(Match As System.Text.RegularExpressions.Match) ArabicLetterWaw + ArabicLetterHamza}, _
            New RuleTranslation With {.RuleName = "YehHamzah", .Match = MakeUniRegEx(ArabicLetterYehWithHamzaAbove), _
                                      .Evaluator = Function(Match As System.Text.RegularExpressions.Match) ArabicLetterAlefMaksura + ArabicLetterHamza}, _
            New RuleTranslation With {.RuleName = "SmallWaw", .Match = MakeUniRegEx(ArabicSmallWaw), _
                                      .Evaluator = Function(Match As System.Text.RegularExpressions.Match) ArabicLetterWaw}, _
            New RuleTranslation With {.RuleName = "SmallYeh", .Match = MakeUniRegEx(ArabicSmallYeh), _
                                      .Evaluator = Function(Match As System.Text.RegularExpressions.Match) ArabicLetterYeh}, _
            New RuleTranslation With {.RuleName = "SmallMeem", .Match = MakeUniRegEx(ArabicSmallLowMeem) + "|" + MakeUniRegEx(ArabicSmallHighMeemIsolatedForm), _
                                      .Evaluator = Function(Match As System.Text.RegularExpressions.Match) ArabicLetterAlef + ArabicLetterHamza}, _
            New RuleTranslation With {.RuleName = "SmallNoon", .Match = MakeUniRegEx(ArabicSmallHighNoon), _
                                      .Evaluator = Function(Match As System.Text.RegularExpressions.Match) ArabicLetterNoon}, _
            New RuleTranslation With {.RuleName = "SmallSeen", .Match = MakeUniRegEx(ArabicSmallHighSeen), _
                                      .Evaluator = Function(Match As System.Text.RegularExpressions.Match) ArabicLetterSeen}, _
            New RuleTranslation With {.RuleName = "Sukun", .Match = MakeUniRegEx(ArabicSukun), _
                                      .Evaluator = Function(Match As System.Text.RegularExpressions.Match) Match.Value.Insert(1, "</empty>").Insert(0, "<empty>")}, _
            New RuleTranslation With {.RuleName = "AlefMaksuraDaggerAlef", .Match = MakeUniRegEx(ArabicLetterAlefMaksura) + MakeUniRegEx(ArabicLetterSuperscriptAlef), _
                                      .Evaluator = Function(Match As System.Text.RegularExpressions.Match) Match.Value.Insert(1, "</helperfatha>").Insert(0, "<helperfatha>")}, _
            New RuleTranslation With {.RuleName = "HelperFatha", .Match = "<helperfatha>.*?</helperfatha>", _
                                      .Evaluator = Function(Match As System.Text.RegularExpressions.Match) ArabicFatha}, _
            New RuleTranslation With {.RuleName = "HelperKasra", .Match = "<helperkasra>.*?</helperkasra>", _
                                      .Evaluator = Function(Match As System.Text.RegularExpressions.Match) ArabicKasra}, _
            New RuleTranslation With {.RuleName = "HelperDamma", .Match = "<helperdamma>.*?</helperdamma>", _
                                      .Evaluator = Function(Match As System.Text.RegularExpressions.Match) ArabicDamma}, _
            New RuleTranslation With {.RuleName = "DaggerAlef", .Match = MakeUniRegEx(ArabicLetterSuperscriptAlef), _
                                      .Evaluator = Function(Match As System.Text.RegularExpressions.Match) ArabicLetterAlef}, _
            New RuleTranslation With {.RuleName = "SmallHamza", .Match = MakeUniRegEx(ArabicHamzaAbove) + "|" + MakeUniRegEx(ArabicHamzaBelow), _
                                      .Evaluator = Function(Match As System.Text.RegularExpressions.Match) ArabicLetterHamza}, _
            New RuleTranslation With {.RuleName = "LetterSpelling", .Match = "\b(" + MakeRegMultiEx(Array.ConvertAll(ArabicUniqueLetters, Function(Str As String) MakeUniRegEx(TransliterateFromBuckwalter(Str)))) + ")\b", _
                                      .Evaluator = Function(Match As System.Text.RegularExpressions.Match) ArabicLetterSpelling(Match.Value)}, _
            New RuleTranslation With {.RuleName = "FathaAlefMaksura", .Match = MakeUniRegEx(ArabicFatha) + MakeUniRegEx(ArabicLetterAlefMaksura), _
                                      .Evaluator = Function(Match As System.Text.RegularExpressions.Match) Match.Value.Remove(1).Insert(1, ArabicLetterAlef)}, _
            New RuleTranslation With {.RuleName = "KasraAlefMaksura", .Match = MakeUniRegEx(ArabicKasra) + MakeUniRegEx(ArabicLetterAlefMaksura), _
                                      .Evaluator = Function(Match As System.Text.RegularExpressions.Match) Match.Value.Remove(1).Insert(1, ArabicLetterYeh)}, _
            New RuleTranslation With {.RuleName = "TehMarbuta", .Match = MakeUniRegEx(ArabicKasra) + MakeUniRegEx(ArabicLetterAlefMaksura), _
                                      .Evaluator = Function(Match As System.Text.RegularExpressions.Match) Match.Value.Remove(1).Insert(1, ArabicLetterYeh)}, _
            New RuleTranslation With {.RuleName = "TehMarbuta", .Match = MakeUniRegEx(ArabicLetterTehMarbuta) + "(" + MakeRegMultiEx(Array.ConvertAll(ArabicFathaDammaKasra, Function(Str As String) MakeUniRegEx(Str))) + "|" + MakeRegMultiEx(Array.ConvertAll(ArabicTanweens, Function(Str As String) MakeUniRegEx(Str))) + ")", _
                                      .Evaluator = Function(Match As System.Text.RegularExpressions.Match) Match.Value.Remove(0).Insert(0, ArabicLetterTeh)}, _
            New RuleTranslation With {.RuleName = "TehMarbuta", .Match = MakeUniRegEx(ArabicLetterTehMarbuta), _
                                      .Evaluator = Function(Match As System.Text.RegularExpressions.Match) ArabicLetterHeh}
        }
    Public Shared RomanizationRules As RuleTranslation() = { _
            New RuleTranslation With {.RuleName = "GutteralRulesBoth", .Match = "(" + MakeRegMultiEx(Array.ConvertAll(ArabicLongVowels, Function(Str As String) MakeUniRegEx(Str))) + "|" + MakeRegMultiEx(Array.ConvertAll(ArabicFathaDammaKasra, Function(Str As String) MakeUniRegEx(Str))) + ")(" + MakeRegMultiEx(Array.ConvertAll(ArabicSpecialLeadingGutteral, Function(Str As String) MakeUniRegEx(Str))) + ")(" + MakeRegMultiEx(Array.ConvertAll(ArabicLongVowels, Function(Str As String) MakeUniRegEx(Str))) + "|" + MakeRegMultiEx(Array.ConvertAll(ArabicFathaDammaKasra, Function(Str As String) MakeUniRegEx(Str))) + "|" + MakeRegMultiEx(Array.ConvertAll(ArabicTanweens, Function(Str As String) MakeUniRegEx(Str))) + ")", _
                                      .Evaluator = Function(Match As System.Text.RegularExpressions.Match) GutteralRules(Match.Value, 0, True, True)}, _
            New RuleTranslation With {.RuleName = "GutteralRules", .Match = "(" + MakeRegMultiEx(Array.ConvertAll(ArabicSpecialGutteral, Function(Str As String) MakeUniRegEx(Str))) + ")(" + MakeRegMultiEx(Array.ConvertAll(ArabicLongVowels, Function(Str As String) MakeUniRegEx(Str))) + "|" + MakeRegMultiEx(Array.ConvertAll(ArabicFathaDammaKasra, Function(Str As String) MakeUniRegEx(Str))) + "|" + MakeRegMultiEx(Array.ConvertAll(ArabicTanweens, Function(Str As String) MakeUniRegEx(Str))) + ")", _
                                      .Evaluator = Function(Match As System.Text.RegularExpressions.Match) GutteralRules(Match.Value, 0, False, True)}, _
            New RuleTranslation With {.RuleName = "LeadingGutteralRules", .Match = "(" + MakeRegMultiEx(Array.ConvertAll(ArabicLongVowels, Function(Str As String) MakeUniRegEx(Str))) + "|" + MakeRegMultiEx(Array.ConvertAll(ArabicFathaDammaKasra, Function(Str As String) MakeUniRegEx(Str))) + ")(" + MakeUniRegEx(ArabicLetterHamza) + "|" + MakeRegMultiEx(Array.ConvertAll(ArabicSpecialLeadingGutteral, Function(Str As String) MakeUniRegEx(Str))) + ")", _
                                      .Evaluator = Function(Match As System.Text.RegularExpressions.Match) GutteralRules(Match.Value, 0, True, False)}, _
            New RuleTranslation With {.RuleName = "Shadda", .Match = "(" + MakeRegMultiEx(Array.ConvertAll(ArabicLetters, Function(Str As String) MakeUniRegEx(Str))) + ")" + MakeUniRegEx(ArabicShadda), _
                                      .Evaluator = Function(Match As System.Text.RegularExpressions.Match) Match.Value.Remove(1, 1).Insert(1, "-" + Match.Value(0))}, _
            New RuleTranslation With {.RuleName = "FathaAlef", .Match = MakeUniRegEx(ArabicFatha) + MakeUniRegEx(ArabicLetterAlef), _
                                      .Evaluator = Function(Match As System.Text.RegularExpressions.Match) "aa"}, _
            New RuleTranslation With {.RuleName = "DammaWaw", .Match = MakeUniRegEx(ArabicDamma) + MakeUniRegEx(ArabicLetterWaw), _
                                      .Evaluator = Function(Match As System.Text.RegularExpressions.Match) "oo"}, _
            New RuleTranslation With {.RuleName = "KasraYeh", .Match = MakeUniRegEx(ArabicKasra) + MakeUniRegEx(ArabicLetterYeh), _
                                      .Evaluator = Function(Match As System.Text.RegularExpressions.Match) "ee"}, _
            New RuleTranslation With {.RuleName = "LettersTanweensVowelsHamza", .Match = "(" + MakeRegMultiEx(Array.ConvertAll(ArabicLetters, Function(Str As String) MakeUniRegEx(Str))) + "|" + MakeRegMultiEx(Array.ConvertAll(ArabicTanweens, Function(Str As String) MakeUniRegEx(Str))) + "|" + MakeRegMultiEx(Array.ConvertAll(ArabicFathaDammaKasra, Function(Str As String) MakeUniRegEx(Str))) + "|" + MakeUniRegEx(ArabicLetterHamza) + ")", _
                                      .Evaluator = Function(Match As System.Text.RegularExpressions.Match) Match.Value.Insert(1, CachedData.IslamData.ArabicLetters(FindLetterBySymbol(Match.Value(0))).PlainRoman).Remove(0, 1)} _
        }
    Public Shared ColoringSpelledOutRules As RuleTranslation() = { _
        New RuleTranslation With {.RuleName = "Normal", .Match = "empty|helperfatha|helperkasra|helperdamma|helpermeem|assimilator|assimilatorincomplete|dipthong|compulsorystop|endofversestop|prostration|canstoporcontinue|betternottostop|stopatfirstnotsecond|stopatsecondnotfirst|bettertostopbutpermissibletocontinue|bettertocontinuebutpermissibletostop|subtlestopwithoutbreath", .Evaluator = Function(Match As System.Text.RegularExpressions.Match) Match.Value}, _
        New RuleTranslation With {.RuleName = "NecessaryProlongation", .Match = "necessaryprolong", .Evaluator = Function(Match As System.Text.RegularExpressions.Match) Match.Value + "-" + Match.Value + "-" + Match.Value + "-" + Match.Value + "-" + Match.Value + "-" + Match.Value}, _
        New RuleTranslation With {.RuleName = "ObligatoryProlongation", .Match = "obligatoryprolong", .Evaluator = Function(Match As System.Text.RegularExpressions.Match) Match.Value + "-" + Match.Value + "-" + Match.Value + "-" + Match.Value + "(-" + Match.Value + ")"}, _
        New RuleTranslation With {.RuleName = "PermissibleProlongation", .Match = "permissibleprolong", .Evaluator = Function(Match As System.Text.RegularExpressions.Match) Match.Value + "-" + Match.Value + "(-" + Match.Value + "-" + Match.Value + ")(-" + Match.Value + "-" + Match.Value + ")"}, _
        New RuleTranslation With {.RuleName = "NormalProlongation", .Match = "normalprolong", .Evaluator = Function(Match As System.Text.RegularExpressions.Match) Match.Value + "-" + Match.Value}, _
        New RuleTranslation With {.RuleName = "Nasalization", .Match = "nasalize", .Evaluator = Function(Match As System.Text.RegularExpressions.Match) Match.Value}, _
        New RuleTranslation With {.RuleName = "Unannounced", .Match = "assimilate|assimilateincomplete", .Evaluator = Function(Match As System.Text.RegularExpressions.Match) Match.Value}, _
        New RuleTranslation With {.RuleName = "EmphaticPronounciation", .Match = "emphasis", .Evaluator = Function(Match As System.Text.RegularExpressions.Match) UCase(Match.Value)}, _
        New RuleTranslation With {.RuleName = "UnrestLetters", .Match = "bounce", .Evaluator = Function(Match As System.Text.RegularExpressions.Match) Match.Value + "-" + Match.Value} _
        }
    Public Shared RulesOfRecitationRegEx As RuleMetadataTranslation() = { _
            New RuleMetadataTranslation With {.RuleName = "Stopping", .Match = "(" + MakeUniRegEx(ArabicSmallHighMeemIsolatedForm) + "|" + MakeUniRegEx(ArabicStartOfRubElHizb) + "|" + MakeUniRegEx(ArabicEndOfAyah) + ")\s*\b",
                .Evaluator = Function(Match As System.Text.RegularExpressions.Match)
                                 If Match.Groups(1).Value = ArabicSmallHighMeemIsolatedForm Then
                                     Return New RuleMetadata() {New RuleMetadata(Match.Groups(1).Index - Match.Index, Match.Groups(1).Length, "compulsorystop")}
                                 ElseIf Match.Groups(1).Value = ArabicStartOfRubElHizb Then
                                     Return New RuleMetadata() {New RuleMetadata(Match.Groups(1).Index - Match.Index, Match.Groups(1).Length, "endofversestop")}
                                 ElseIf Match.Groups(1).Value = ArabicEndOfAyah Then
                                     Return New RuleMetadata() {New RuleMetadata(Match.Groups(1).Index - Match.Index, Match.Groups(1).Length, "prostration")}
                                 Else
                                     Return Nothing
                                 End If
                             End Function}, _
            New RuleMetadataTranslation With {.RuleName = "Stopping", .Match = MakeUniRegEx(ArabicSmallHighJeem) + "|" + MakeUniRegEx(ArabicSmallHighLamAlef) + "|" + MakeUniRegEx(ArabicSmallHighThreeDots) + "\w*" + MakeUniRegEx(ArabicSmallHighThreeDots) + "|" + MakeUniRegEx(ArabicSmallHighLigatureQafWithLamWithAlefMaksura) + "|" + MakeUniRegEx(ArabicSmallHighLigatureSadWithLamWithAlefMaksura) + "|" + MakeUniRegEx(ArabicSmallHighSeen),
                .Evaluator = Function(Match As System.Text.RegularExpressions.Match)
                                 If Match.Groups(1).Value = ArabicSmallHighJeem Then
                                     Return New RuleMetadata() {New RuleMetadata(0, Match.Length, "canstoporcontinue")}
                                 ElseIf Match.Groups(1).Value = ArabicSmallHighLamAlef Then
                                     Return New RuleMetadata() {New RuleMetadata(0, Match.Length, "betternottostop")}
                                 ElseIf Match.Groups(1).Value.StartsWith(ArabicSmallHighThreeDots) Then
                                     Return New RuleMetadata() {New RuleMetadata(Match.Length - 1, Match.Length, "stopatsecondnotfirst"), New RuleMetadata(0, 1, "stopatfirstnotsecond")}
                                 ElseIf Match.Groups(1).Value = ArabicSmallHighLigatureQafWithLamWithAlefMaksura Then
                                     Return New RuleMetadata() {New RuleMetadata(0, Match.Length, "bettertostopbutpermissibletocontinue")}
                                 ElseIf Match.Groups(1).Value = ArabicSmallHighLigatureSadWithLamWithAlefMaksura Then
                                     Return New RuleMetadata() {New RuleMetadata(0, Match.Length, "bettertocontinuebutpermissibletostop")}
                                 ElseIf Match.Groups(1).Value = ArabicSmallHighSeen Then
                                     Return New RuleMetadata() {New RuleMetadata(0, Match.Length, "subtlestopwithoutbreath")}
                                 Else
                                     Return Nothing
                                 End If
                             End Function}, _
            New RuleMetadataTranslation With {.RuleName = "Stopping", .Match = "(" + MakeRegMultiEx(Array.ConvertAll(ArabicFathaDammaKasra, Function(Str As String) MakeUniRegEx(Str))) + "|" + MakeUniRegEx(ArabicKasratan) + "|" + MakeUniRegEx(ArabicDammatan) + ")\s*$",
                .Evaluator = Function(Match As System.Text.RegularExpressions.Match)
                                 Return New RuleMetadata() {New RuleMetadata(Match.Groups(1).Index - Match.Index, Match.Groups(1).Length, "empty")}
                             End Function}, _
            New RuleMetadataTranslation With {.RuleName = "Stopping", .Match = MakeUniRegEx(ArabicLetterTehMarbuta) + "(" + MakeRegMultiEx(Array.ConvertAll(ArabicFathaDammaKasra, Function(Str As String) MakeUniRegEx(Str))) + "|" + MakeRegMultiEx(Array.ConvertAll(ArabicTanweens, Function(Str As String) MakeUniRegEx(Str))) + ")",
                .Evaluator = Function(Match As System.Text.RegularExpressions.Match)
                                 Return New RuleMetadata() {New RuleMetadata(0, 1, "helperheh")}
                             End Function}, _
            New RuleMetadataTranslation With {.RuleName = "SmallBounce", .Match = "(" + MakeRegMultiEx(Array.ConvertAll(ArabicFathaDammaKasra, Function(Str As String) MakeUniRegEx(Str))) + ")(" + MakeUniRegEx(ArabicLetterDal) + "|" + MakeUniRegEx(ArabicLetterBeh) + "|" + MakeUniRegEx(ArabicLetterQaf) + "|" + MakeUniRegEx(ArabicLetterTah) + "|" + MakeUniRegEx(ArabicLetterJeem) + ")" + MakeUniRegEx(ArabicSukun) + "\B", _
                .Evaluator = Function(Match As System.Text.RegularExpressions.Match)
                                 Return New RuleMetadata() {New RuleMetadata(Match.Groups(2).Index - Match.Index, Match.Groups(2).Length, "bounce")}
                             End Function}, _
            New RuleMetadataTranslation With {.RuleName = "ModerateBounce", .Match = "(" + MakeUniRegEx(ArabicLetterDal) + "|" + MakeUniRegEx(ArabicLetterBeh) + "|" + MakeUniRegEx(ArabicLetterQaf) + "|" + MakeUniRegEx(ArabicLetterTah) + "|" + MakeUniRegEx(ArabicLetterJeem) + ")" + "(" + MakeRegMultiEx(Array.ConvertAll(ArabicTanweens, Function(Str As String) MakeUniRegEx(Str))) + "|" + MakeRegMultiEx(Array.ConvertAll(ArabicFathaDammaKasra, Function(Str As String) MakeUniRegEx(Str))) + ")\s*$", _
                .Evaluator = Function(Match As System.Text.RegularExpressions.Match)
                                 Return New RuleMetadata() {New RuleMetadata(Match.Groups(1).Index - Match.Index, Match.Groups(1).Length, "bounce")}
                             End Function}, _
            New RuleMetadataTranslation With {.RuleName = "GreatBounce", .Match = "(" + MakeUniRegEx(ArabicLetterDal) + "|" + MakeUniRegEx(ArabicLetterBeh) + "|" + MakeUniRegEx(ArabicLetterQaf) + "|" + MakeUniRegEx(ArabicLetterTah) + "|" + MakeUniRegEx(ArabicLetterJeem) + ")" + MakeUniRegEx(ArabicShadda) + MakeUniRegEx(ArabicSukun) + "\s*$", _
                .Evaluator = Function(Match As System.Text.RegularExpressions.Match)
                                 Return New RuleMetadata() {New RuleMetadata(Match.Groups(1).Index - Match.Index, Match.Groups(1).Length, "bounce")}
                             End Function}, _
            New RuleMetadataTranslation With {.RuleName = "NasalizeCharacterDoubled", .Match = MakeUniRegEx(ArabicLetterNoon) + MakeUniRegEx(ArabicShadda),
                .Evaluator = Function(Match As System.Text.RegularExpressions.Match)
                                 Return New RuleMetadata() {New RuleMetadata(0, Match.Length, "nasalize normalprolong")}
                             End Function}, _
            New RuleMetadataTranslation With {.RuleName = "VowellessNoonClear", .Match = "(" + MakeUniRegEx(ArabicLetterNoon) + MakeUniRegEx(ArabicSukun) + "|(" + MakeUniRegEx(ArabicLetterNoon) + "|" + MakeRegMultiEx(Array.ConvertAll(ArabicTanweens, Function(Str As String) MakeUniRegEx(Str))) + ")\b\s*)(" + MakeUniRegEx(ArabicLetterAin) + "|" + MakeUniRegEx(ArabicLetterGhain) + "|" + MakeUniRegEx(ArabicLetterJeem) + "|" + MakeUniRegEx(ArabicLetterHah) + "|" + MakeUniRegEx(ArabicLetterHeh) + "|" + MakeUniRegEx(ArabicLetterHamza) + ")",
                .Evaluator = Function(Match As System.Text.RegularExpressions.Match)
                                 Return Nothing
                             End Function}, _
            New RuleMetadataTranslation With {.RuleName = "VowellessNoonCovered", .Match = "(" + MakeUniRegEx(ArabicLetterNoon) + MakeUniRegEx(ArabicSukun) + MakeUniRegEx(ArabicSmallHighMeemIsolatedForm) + "|(" + MakeUniRegEx(ArabicLetterNoon) + MakeUniRegEx(ArabicSmallHighMeemIsolatedForm) + "|" + MakeUniRegEx(ArabicFathatan) + MakeUniRegEx(ArabicSmallHighMeemIsolatedForm) + "|" + MakeUniRegEx(ArabicDammatan) + MakeUniRegEx(ArabicSmallHighMeemIsolatedForm) + "|" + MakeUniRegEx(ArabicKasratan) + MakeUniRegEx(ArabicSmallLowMeem) + "|" + ")\b\s*)(" + MakeUniRegEx(ArabicLetterBeh) + MakeUniRegEx(ArabicFatha) + "|" + MakeUniRegEx(ArabicLetterBeh) + MakeUniRegEx(ArabicDamma) + "|" + MakeUniRegEx(ArabicLetterBeh) + MakeUniRegEx(ArabicKasra) + ")",
                .Evaluator = Function(Match As System.Text.RegularExpressions.Match)
                                 If Match.Groups(1).Value = ArabicLetterNoon + ArabicSukun + ArabicSmallHighMeemIsolatedForm Then
                                     Return New RuleMetadata() {New RuleMetadata(Match.Groups(1).Length - 1 + Match.Groups(1).Index - Match.Index, 1, "nasalize"), New RuleMetadata(0, 2, "empty")}
                                 ElseIf Match.Groups(1).Value.StartsWith(ArabicLetterNoon + ArabicSmallHighMeemIsolatedForm) Then
                                     Return New RuleMetadata() {New RuleMetadata(1 + Match.Groups(1).Index - Match.Index, 1, "nasalize"), New RuleMetadata(0, 1, "empty")}
                                 Else
                                     Return New RuleMetadata() {New RuleMetadata(1 + Match.Groups(1).Index - Match.Index, 1, "nasalize"), New RuleMetadata(0, 1, "dividetanween(,empty)")}
                                 End If
                             End Function}, _
            New RuleMetadataTranslation With {.RuleName = "VowellessNoonAssimilatingNasalization", .Match = "(" + MakeUniRegEx(ArabicLetterNoon) + MakeUniRegEx(ArabicSukun) + "|(" + MakeUniRegEx(ArabicLetterNoon) + "|" + MakeRegMultiEx(Array.ConvertAll(ArabicTanweens, Function(Str As String) MakeUniRegEx(Str))) + ")\b\s*)(" + MakeUniRegEx(ArabicLetterNoon) + "|" + MakeUniRegEx(ArabicLetterWaw) + "|" + MakeUniRegEx(ArabicLetterMeem) + "|" + MakeUniRegEx(ArabicLetterYeh) + ")",
                .Evaluator = Function(Match As System.Text.RegularExpressions.Match)
                                 If Match.Groups(1).Value = ArabicLetterNoon + ArabicSukun Then
                                     Return New RuleMetadata() {New RuleMetadata(Match.Groups(2).Index - Match.Index, Match.Groups(2).Length, "nasalize normalprolong assimilator"), New RuleMetadata(0, 2, "assimilate")}
                                 ElseIf Match.Groups(1).Value.StartsWith(ArabicLetterNoon) Then
                                     Return New RuleMetadata() {New RuleMetadata(Match.Groups(2).Index - Match.Index, Match.Groups(2).Length, "nasalize normalprolong assimilator"), New RuleMetadata(0, 1, "assimilate")}
                                 Else
                                     Return New RuleMetadata() {New RuleMetadata(Match.Groups(2).Index - Match.Index, Match.Groups(2).Length, "nasalize normalprolong assimilator"), New RuleMetadata(0, 1, "dividetanween(,assimilate)")}
                                 End If
                             End Function}, _
            New RuleMetadataTranslation With {.RuleName = "VowellessNoonAssimilating", .Match = "(" + MakeUniRegEx(ArabicLetterNoon) + MakeUniRegEx(ArabicSukun) + "|(" + MakeUniRegEx(ArabicLetterNoon) + "|" + MakeRegMultiEx(Array.ConvertAll(ArabicTanweens, Function(Str As String) MakeUniRegEx(Str))) + ")\b\s*)(" + MakeUniRegEx(ArabicLetterLam) + "|" + MakeUniRegEx(ArabicLetterReh) + ")",
                .Evaluator = Function(Match As System.Text.RegularExpressions.Match)
                                 If Match.Groups(1).Value = ArabicLetterNoon + ArabicSukun Then
                                     Return New RuleMetadata() {New RuleMetadata(Match.Groups(2).Index - Match.Index, Match.Groups(2).Length, "assimilator"), New RuleMetadata(0, 2, "assimilate")}
                                 ElseIf Match.Groups(1).Value.StartsWith(ArabicLetterNoon) Then
                                     Return New RuleMetadata() {New RuleMetadata(Match.Groups(2).Index - Match.Index, Match.Groups(2).Length, "assimilator"), New RuleMetadata(0, 1, "assimilate")}
                                 Else
                                     Return New RuleMetadata() {New RuleMetadata(Match.Groups(2).Index - Match.Index, Match.Groups(2).Length, "assimilator"), New RuleMetadata(0, 1, "dividetanween(,assimilate)")}
                                 End If
                             End Function}, _
            New RuleMetadataTranslation With {.RuleName = "VowellessNoonHideHeaviness", .Match = "(" + MakeUniRegEx(ArabicLetterNoon) + MakeUniRegEx(ArabicSukun) + "|(" + MakeUniRegEx(ArabicLetterNoon) + "|" + MakeRegMultiEx(Array.ConvertAll(ArabicTanweens, Function(Str As String) MakeUniRegEx(Str))) + ")\b\s*)(" + MakeUniRegEx(ArabicLetterSad) + "|" + MakeUniRegEx(ArabicLetterDad) + "|" + MakeUniRegEx(ArabicLetterTah) + "|" + MakeUniRegEx(ArabicLetterZah) + "|" + MakeUniRegEx(ArabicLetterQaf) + ")",
                .Evaluator = Function(Match As System.Text.RegularExpressions.Match)
                                 If Match.Groups(1).Value = ArabicLetterNoon + ArabicSukun Then
                                     Return New RuleMetadata() {New RuleMetadata(0, 2, "nasalize normalprolong")}
                                 ElseIf Match.Groups(1).Value.StartsWith(ArabicLetterNoon) Then
                                     Return New RuleMetadata() {New RuleMetadata(0, 1, "nasalize normalprolong")}
                                 Else
                                     Return New RuleMetadata() {New RuleMetadata(0, 1, "dividetanween(,normalprolong nasalize)")}
                                 End If
                             End Function}, _
            New RuleMetadataTranslation With {.RuleName = "VowellessNoonHideLightness", .Match = "(" + MakeUniRegEx(ArabicLetterNoon) + MakeUniRegEx(ArabicSukun) + "|(" + MakeUniRegEx(ArabicLetterNoon) + "|" + MakeRegMultiEx(Array.ConvertAll(ArabicTanweens, Function(Str As String) MakeUniRegEx(Str))) + ")\b\s*)(" + MakeUniRegEx(ArabicLetterTeh) + "|" + MakeUniRegEx(ArabicLetterTheh) + "|" + MakeUniRegEx(ArabicLetterJeem) + "|" + MakeUniRegEx(ArabicLetterDal) + "|" + MakeUniRegEx(ArabicLetterThal) + "|" + MakeUniRegEx(ArabicLetterZain) + "|" + MakeUniRegEx(ArabicLetterSeen) + "|" + MakeUniRegEx(ArabicLetterSheen) + "|" + MakeUniRegEx(ArabicLetterFeh) + "|" + MakeUniRegEx(ArabicLetterKaf) + ")",
                .Evaluator = Function(Match As System.Text.RegularExpressions.Match)
                                 If Match.Groups(1).Value = ArabicLetterNoon + ArabicSukun Then
                                     Return New RuleMetadata() {New RuleMetadata(0, 2, "nasalize normalprolong")}
                                 ElseIf Match.Groups(1).Value.StartsWith(ArabicLetterNoon) Then
                                     Return New RuleMetadata() {New RuleMetadata(0, 1, "nasalize normalprolong")}
                                 Else
                                     Return New RuleMetadata() {New RuleMetadata(0, 1, "dividetanween(,normalprolong nasalize)")}
                                 End If
                             End Function}, _
            New RuleMetadataTranslation With {.RuleName = "CharacterDoubled", .Match = MakeUniRegEx(ArabicLetterMeem) + MakeUniRegEx(ArabicShadda),
                .Evaluator = Function(Match As System.Text.RegularExpressions.Match)
                                 Return New RuleMetadata() {New RuleMetadata(0, Match.Length, "nasalize normalprolong")}
                             End Function}, _
            New RuleMetadataTranslation With {.RuleName = "VowellessMeemHide", .Match = "(" + MakeUniRegEx(ArabicLetterMeem) + MakeUniRegEx(ArabicSukun) + "|" + MakeUniRegEx(ArabicLetterMeem) + "\b\s*)(" + MakeUniRegEx(ArabicLetterBeh) + MakeUniRegEx(ArabicFatha) + "|" + MakeUniRegEx(ArabicLetterBeh) + MakeUniRegEx(ArabicDamma) + "|" + MakeUniRegEx(ArabicLetterBeh) + MakeUniRegEx(ArabicKasra) + ")",
                .Evaluator = Function(Match As System.Text.RegularExpressions.Match)
                                 If Match.Groups(1).Value = ArabicLetterMeem + ArabicSukun Then
                                     Return New RuleMetadata() {New RuleMetadata(0, 2, "nasalize normalprolong")}
                                 Else
                                     Return New RuleMetadata() {New RuleMetadata(0, 1, "nasalize normalprolong")}
                                 End If
                                 Return New RuleMetadata() {New RuleMetadata(Match.Groups(2).Index - Match.Index, Match.Groups(2).Length, "nasalize assimilator"), New RuleMetadata(Match.Groups(1).Index - Match.Index, Match.Groups(1).Length, "assimilate")}
                             End Function}, _
            New RuleMetadataTranslation With {.RuleName = "VowellessMeemAssimilating", .Match = "(" + MakeUniRegEx(ArabicLetterMeem) + MakeUniRegEx(ArabicSukun) + "|" + MakeUniRegEx(ArabicLetterMeem) + "\b\s*)(" + MakeUniRegEx(ArabicLetterMeem) + MakeUniRegEx(ArabicFatha) + "|" + MakeUniRegEx(ArabicLetterMeem) + MakeUniRegEx(ArabicDamma) + "|" + MakeUniRegEx(ArabicLetterMeem) + MakeUniRegEx(ArabicKasra) + ")",
                .Evaluator = Function(Match As System.Text.RegularExpressions.Match)
                                 If Match.Groups(1).Value = ArabicLetterMeem + ArabicSukun Then
                                     Return New RuleMetadata() {New RuleMetadata(Match.Groups(2).Index - Match.Index, Match.Groups(2).Length, "nasalize normalprolong assimilator"), New RuleMetadata(0, 2, "assimilate")}
                                 Else
                                     Return New RuleMetadata() {New RuleMetadata(Match.Groups(2).Index - Match.Index, Match.Groups(2).Length, "nasalize normalprolong assimilator"), New RuleMetadata(0, 1, "assimilate")}
                                 End If
                             End Function}, _
            New RuleMetadataTranslation With {.RuleName = "VowellessMeemClear", .Match = "(" + MakeUniRegEx(ArabicLetterMeem) + MakeUniRegEx(ArabicSukun) + "|" + MakeUniRegEx(ArabicLetterMeem) + "\b\s*)(" + MakeUniRegEx(ArabicLetterTeh) + "|" + MakeUniRegEx(ArabicLetterTheh) + "|" + MakeUniRegEx(ArabicLetterJeem) + "|" + MakeUniRegEx(ArabicLetterHah) + "|" + MakeUniRegEx(ArabicLetterKhah) + "|" + MakeUniRegEx(ArabicLetterDal) + "|" + MakeUniRegEx(ArabicLetterThal) + "|" + MakeUniRegEx(ArabicLetterReh) + "|" + MakeUniRegEx(ArabicLetterZain) + "|" + MakeUniRegEx(ArabicLetterSeen) + "|" + MakeUniRegEx(ArabicLetterSheen) + "|" + MakeUniRegEx(ArabicLetterSad) + "|" + MakeUniRegEx(ArabicLetterDad) + "|" + MakeUniRegEx(ArabicLetterTah) + "|" + MakeUniRegEx(ArabicLetterZah) + "|" + MakeUniRegEx(ArabicLetterAin) + "|" + MakeUniRegEx(ArabicLetterGhain) + "|" + MakeUniRegEx(ArabicLetterQaf) + "|" + MakeUniRegEx(ArabicLetterKaf) + "|" + MakeUniRegEx(ArabicLetterLam) + "|" + MakeUniRegEx(ArabicLetterNoon) + "|" + MakeUniRegEx(ArabicLetterHamza) + "|((" + MakeUniRegEx(ArabicSukun) + "|" + MakeUniRegEx(ArabicFatha) + ")(" + MakeUniRegEx(ArabicLetterYeh) + "))" + ")(" + MakeRegMultiEx(Array.ConvertAll(ArabicFathaDammaKasra, Function(Str As String) MakeUniRegEx(Str))) + ")",
                .Evaluator = Function(Match As System.Text.RegularExpressions.Match)
                                 Return Nothing
                             End Function}, _
            New RuleMetadataTranslation With {.RuleName = "VowellessMeemClearGreater", .Match = "(" + MakeUniRegEx(ArabicLetterMeem) + MakeUniRegEx(ArabicSukun) + "|" + MakeUniRegEx(ArabicLetterMeem) + "\b\s*)(" + MakeUniRegEx(ArabicLetterFeh) + "|((" + MakeUniRegEx(ArabicSukun) + "|" + MakeUniRegEx(ArabicFatha) + ")(" + MakeUniRegEx(ArabicLetterWaw) + "))" + ")(" + MakeRegMultiEx(Array.ConvertAll(ArabicFathaDammaKasra, Function(Str As String) MakeUniRegEx(Str))) + ")",
                .Evaluator = Function(Match As System.Text.RegularExpressions.Match)
                                 Return Nothing
                             End Function}, _
            New RuleMetadataTranslation With {.RuleName = "EmptyHamza", .Match = "^\s*(" + MakeUniRegEx(ArabicLetterAlefWasla) + ")" + MakeUniRegEx(ArabicLetterLam) + "((" + MakeRegMultiEx(Array.ConvertAll(ArabicSunLetters, Function(Str As String) MakeUniRegEx(Str))) + ")" + MakeUniRegEx(ArabicShadda) + "|(" + MakeRegMultiEx(Array.ConvertAll(ArabicMoonLetters, Function(Str As String) MakeUniRegEx(Str))) + "))",
                .Evaluator = Function(Match As System.Text.RegularExpressions.Match)
                                 Return New RuleMetadata() {New RuleMetadata(Match.Groups(1).Index - Match.Index, Match.Groups(1).Length, "helperfatha")}
                             End Function}, _
            New RuleMetadataTranslation With {.RuleName = "EmptyHamza", .Match = "^\s*(" + MakeUniRegEx(ArabicLetterAlefWasla) + "(" + MakeRegMultiEx(Array.ConvertAll(ArabicLetters, Function(Str As String) MakeUniRegEx(Str))) + ")" + "(" + MakeRegMultiEx(Array.ConvertAll(ArabicFathaDammaKasra, Function(Str As String) MakeUniRegEx(Str))) + ")?" + "(" + MakeRegMultiEx(Array.ConvertAll(ArabicLetters, Function(Str As String) MakeUniRegEx(Str))) + ")(" + MakeUniRegEx(ArabicFatha) + "|" + MakeUniRegEx(ArabicKasra) + ")|" + MakeRegMultiEx(Array.ConvertAll(ArabicWaslKasraExceptions, Function(Str As String) MakeUniRegEx(TransliterateFromBuckwalter(Str)))) + ")",
                .Evaluator = Function(Match As System.Text.RegularExpressions.Match)
                                 Return New RuleMetadata() {New RuleMetadata(Match.Groups(1).Index - Match.Index, 1, "helperkasra")}
                             End Function}, _
            New RuleMetadataTranslation With {.RuleName = "EmptyHamza", .Match = "^\s*(" + MakeUniRegEx(ArabicLetterAlefWasla) + ")(" + MakeRegMultiEx(Array.ConvertAll(ArabicLetters, Function(Str As String) MakeUniRegEx(Str))) + ")" + "(" + MakeRegMultiEx(Array.ConvertAll(ArabicFathaDammaKasra, Function(Str As String) MakeUniRegEx(Str))) + ")?" + "(" + MakeRegMultiEx(Array.ConvertAll(ArabicLetters, Function(Str As String) MakeUniRegEx(Str))) + ")" + MakeUniRegEx(ArabicDamma),
                .Evaluator = Function(Match As System.Text.RegularExpressions.Match)
                                 Return New RuleMetadata() {New RuleMetadata(Match.Groups(1).Index - Match.Index, Match.Groups(1).Length, "helperdamma")}
                             End Function}, _
            New RuleMetadataTranslation With {.RuleName = "EmptyHamza", .Match = "\w+\s*(" + MakeUniRegEx(ArabicLetterAlefWasla) + ")",
                .Evaluator = Function(Match As System.Text.RegularExpressions.Match)
                                 Return New RuleMetadata() {New RuleMetadata(Match.Groups(1).Index - Match.Index, Match.Groups(1).Length, "empty")}
                             End Function}, _
            New RuleMetadataTranslation With {.RuleName = "AssimilateLaamSunLetter", .Match = "(^\s*|\b)(" + MakeUniRegEx(ArabicLetterWaw) + MakeUniRegEx(ArabicFatha) + "|" + MakeUniRegEx(ArabicLetterBeh) + MakeUniRegEx(ArabicKasra) + "|" + MakeUniRegEx(ArabicLetterTeh) + MakeUniRegEx(ArabicFatha) + "|" + MakeUniRegEx(ArabicLetterKaf) + MakeUniRegEx(ArabicFatha) + "|" + MakeUniRegEx(ArabicLetterLam) + MakeUniRegEx(ArabicKasra) + ")?(" + MakeUniRegEx(ArabicLetterAlefWasla) + MakeUniRegEx(ArabicLetterLam) + ")(" + MakeRegMultiEx(Array.ConvertAll(ArabicSunLetters, Function(Str As String) MakeUniRegEx(Str))) + ")" + MakeUniRegEx(ArabicShadda),
                .Evaluator = Function(Match As System.Text.RegularExpressions.Match)
                                 Return New RuleMetadata() {New RuleMetadata(Match.Groups(4).Index - Match.Index, Match.Groups(4).Length, "assimilator"), New RuleMetadata(Match.Groups(3).Length - 1 + Match.Groups(3).Index - Match.Index, 1, "assimilate")}
                             End Function}, _
            New RuleMetadataTranslation With {.RuleName = "ClearLaamMoonLetter", .Match = "(^\s*|\b)(" + MakeUniRegEx(ArabicLetterWaw) + MakeUniRegEx(ArabicFatha) + "|" + MakeUniRegEx(ArabicLetterBeh) + MakeUniRegEx(ArabicKasra) + "|" + MakeUniRegEx(ArabicLetterTeh) + MakeUniRegEx(ArabicFatha) + "|" + MakeUniRegEx(ArabicLetterKaf) + MakeUniRegEx(ArabicFatha) + "|" + MakeUniRegEx(ArabicLetterLam) + MakeUniRegEx(ArabicKasra) + ")?(" + MakeUniRegEx(ArabicLetterAlefWasla) + MakeUniRegEx(ArabicLetterLam) + ")(" + MakeRegMultiEx(Array.ConvertAll(ArabicSunLetters, Function(Str As String) MakeUniRegEx(Str))) + ")" + MakeUniRegEx(ArabicShadda),
                .Evaluator = Function(Match As System.Text.RegularExpressions.Match)
                                 Return New RuleMetadata() {New RuleMetadata(Match.Groups(3).Index - Match.Index, Match.Groups(3).Length, "assimilator"), New RuleMetadata(Match.Groups(2).Length - 1 + Match.Groups(2).Index - Match.Index, 1, "assimilate")}
                             End Function} _
        }
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
            If IsLetter(Index) Then
                Output += CachedData.IslamData.ArabicLetters(Index).SymbolName
            End If
        Next
        Return Output
    End Function
    Public Shared Function GutteralRules(ArabicString As String, Index As Integer, bLeading As Boolean, bTrailing As Boolean) As String
        If bLeading Then
            If ArabicString(Index) = ArabicFatha Then
                If ArabicString(Index + 1) = ArabicLetterAlef Then
                    Dim Letter As String = CStr(IIf(bTrailing, String.Empty, CachedData.IslamData.ArabicLetters(FindLetterBySymbol(ArabicString(Index + 2))).PlainRoman))
                    ArabicString = ArabicString.Remove(Index, CInt(IIf(bTrailing, 2, 3))).Insert(Index, "aae" + Letter)
                    If bTrailing Then Index += 3
                Else
                    Dim Letter As String = CStr(IIf(bTrailing, String.Empty, CachedData.IslamData.ArabicLetters(FindLetterBySymbol(ArabicString(Index + 1))).PlainRoman))
                    ArabicString = ArabicString.Remove(Index, CInt(IIf(bTrailing, 1, 2))).Insert(Index, "ah" + Letter)
                    If bTrailing Then Index += 2
                End If
            ElseIf ArabicString(Index) = ArabicDamma Then
                If ArabicString(Index + 1) = ArabicLetterWaw Then
                    Dim Letter As String = CStr(IIf(bTrailing, String.Empty, CachedData.IslamData.ArabicLetters(FindLetterBySymbol(ArabicString(Index + 2))).PlainRoman))
                    ArabicString = ArabicString.Remove(Index, CInt(IIf(bTrailing, 2, 3))).Insert(Index, "eeh" + Letter)
                    If bTrailing Then Index += 3
                Else
                    Dim Letter As String = CStr(IIf(bTrailing, String.Empty, CachedData.IslamData.ArabicLetters(FindLetterBySymbol(ArabicString(Index + 1))).PlainRoman))
                    ArabicString = ArabicString.Remove(Index, CInt(IIf(bTrailing, 1, 2))).Insert(Index, "k" + Letter)
                    If bTrailing Then Index += 1
                End If
            ElseIf ArabicString(Index) = ArabicKasra Then
                If ArabicString(Index + 1) = ArabicLetterYeh Then
                    Dim Letter As String = CStr(IIf(bTrailing, String.Empty, CachedData.IslamData.ArabicLetters(FindLetterBySymbol(ArabicString(Index + 2))).PlainRoman))
                    ArabicString = ArabicString.Remove(Index, CInt(IIf(bTrailing, 2, 3))).Insert(Index, "ooh" + Letter)
                    If bTrailing Then Index += 3
                Else
                    Dim Letter As String = CStr(IIf(bTrailing, String.Empty, CachedData.IslamData.ArabicLetters(FindLetterBySymbol(ArabicString(Index + 1))).PlainRoman))
                    ArabicString = ArabicString.Remove(Index, CInt(IIf(bTrailing, 1, 2))).Insert(Index, "o" + Letter)
                    If bTrailing Then Index += 1
                End If
            End If
        End If
        If bTrailing Then
            Dim Letter As String = CachedData.IslamData.ArabicLetters(FindLetterBySymbol(ArabicString(Index))).PlainRoman
            If ArabicString(Index + 1) = ArabicFatha Then
                If ArabicString.Length <> Index + 2 AndAlso ArabicString(Index + 2) = ArabicLetterAlef Then
                    ArabicString = ArabicString.Remove(Index, 3).Insert(Index, Letter + "waa")
                Else
                    ArabicString = ArabicString.Remove(Index, 2).Insert(Index, Letter + "aw")
                End If
            ElseIf ArabicString(Index + 1) = ArabicDamma Then
                If ArabicString.Length <> Index + 2 AndAlso ArabicString(Index + 2) = ArabicLetterWaw Then
                    ArabicString = ArabicString.Remove(Index, 3).Insert(Index, Letter + "eoo")
                Else
                    ArabicString = ArabicString.Remove(Index, 2).Insert(Index, Letter + "o")
                End If
            ElseIf ArabicString(Index + 1) = ArabicKasra Then
                If ArabicString.Length <> Index + 2 AndAlso ArabicString(Index + 2) = ArabicLetterYeh Then
                    ArabicString = ArabicString.Remove(Index, 3).Insert(Index, Letter + "aee")
                Else
                    ArabicString = ArabicString.Remove(Index, 2).Insert(Index, Letter + "e")
                End If
            ElseIf ArabicString(Index + 1) = ArabicFathatan Then
                ArabicString = ArabicString.Remove(Index, 2).Insert(Index, Letter + "awn")
            ElseIf ArabicString(Index + 1) = ArabicDammatan Then
                ArabicString = ArabicString.Remove(Index, 2).Insert(Index, Letter + "on")
            ElseIf ArabicString(Index + 1) = ArabicKasratan Then
                ArabicString = ArabicString.Remove(Index, 2).Insert(Index, Letter + "en")
            End If
        End If
        Return ArabicString
    End Function
    Public Shared Function ColorRuleSplit() As String
        Return "(" + MakeRegMultiEx(Array.ConvertAll(ColoringRules, Function(ClrRule As ColorRule) MakeRegMultiEx(Array.ConvertAll(ClrRule.Match.Split("|"c), Function(Rule As String) "<" + Rule + ">(.+?)</" + Rule + ">")))) + ")" ' paranthesis causes delimeter capture
    End Function
    Class RuleMetadataComparer
        Implements Collections.Generic.IComparer(Of RuleMetadata)
        Public Function Compare(x As RuleMetadata, y As RuleMetadata) As Integer Implements Generic.IComparer(Of RuleMetadata).Compare
            If x.Index = y.Index Then
                Return x.Length.CompareTo(y.Length)
            Else
                Return x.Index.CompareTo(y.Index)
            End If
        End Function
    End Class
    Public Shared Function ApplyColorRules(ByVal ArabicString As String) As RenderArray.RenderText()
        Dim Count As Integer
        Dim Index As Integer
        Dim MetadataList As New Generic.List(Of RuleMetadata)
        Dim Strings As New Generic.List(Of RenderArray.RenderText)
        For Count = 0 To RulesOfRecitationRegEx.Length - 1
            MetadataList.AddRange(RulesOfRecitationRegEx(Count).Evaluator(System.Text.RegularExpressions.Regex.Match(ArabicString, RulesOfRecitationRegEx(Count).Match)))
        Next
        MetadataList.Sort(New RuleMetadataComparer)
        For Index = 0 To MetadataList.Count - 1
            If If(Index <> 0, MetadataList(Index - 1).Index + MetadataList(Index - 1).Length, 0) <> MetadataList(Index).Index Then
                Strings.Add(New RenderArray.RenderText(RenderArray.RenderDisplayClass.eArabic, ArabicString.Substring(If(Index <> 0, MetadataList(Index - 1).Index + MetadataList(Index - 1).Length, 0), MetadataList(Index).Index - If(Index <> 0, MetadataList(Index - 1).Index + MetadataList(Index - 1).Length, 0))))
            End If
            Strings.Add(New RenderArray.RenderText(RenderArray.RenderDisplayClass.eArabic, ArabicString.Substring(MetadataList(Index).Index, MetadataList(Index).Length)))
            For Count = 0 To ColoringRules.Length - 1
                Dim Match As Integer = Array.FindIndex(ColoringRules(Count).Match.Split("|"c), Function(Str As String) Str = MetadataList(Index).Type)
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
    Public Shared Function TransliterateToPlainRoman(ByVal ArabicString As String) As String
        Dim RomanString As String = String.Empty
        'need to check for decomposed first
        If System.Text.RegularExpressions.Regex.Matches(ArabicString, MakeUniRegEx(ArabicLetterAlefWithHamzaAbove) + "[^" + MakeUniRegEx(ArabicFatha) + MakeUniRegEx(ArabicDamma) + "]|" + MakeUniRegEx(ArabicLetterAlefWithHamzaBelow) + "[^" + MakeUniRegEx(ArabicKasra) + "]").Count <> 0 Then
            Diagnostics.Debug.Print("Missing diacritic: " + ArabicString)
        End If
        If System.Text.RegularExpressions.Regex.Matches(ArabicString, "(" + MakeRegMultiEx(ArabicSunLetters) + "|" + MakeRegMultiEx(ArabicMoonLettersNoVowels) + "|" + MakeUniRegEx(ArabicFatha) + MakeUniRegEx(ArabicLetterYeh) + "|" + MakeUniRegEx(ArabicFatha) + MakeUniRegEx(ArabicLetterWaw) + ")(" + MakeRegMultiEx(ArabicSunLetters) + "|" + MakeRegMultiEx(ArabicMoonLettersNoVowels) + "|" + MakeUniRegEx(ArabicLetterAlef) + "|" + MakeUniRegEx(ArabicLetterWaw) + "|" + MakeUniRegEx(ArabicLetterYeh) + ")").Count <> 0 Then
            Diagnostics.Debug.Print("Missing diacritic: " + ArabicString)
        End If
        If System.Text.RegularExpressions.Regex.Matches(ArabicString, MakeUniRegEx(ArabicLetterAlefMaksura) + "(" + MakeRegMultiEx(Array.ConvertAll(ArabicFathaDammaKasra, Function(Str As String) MakeUniRegEx(Str))) + "|" + MakeRegMultiEx(Array.ConvertAll(ArabicTanweens, Function(Str As String) MakeUniRegEx(Str))) + "|" + MakeUniRegEx(ArabicLetterSuperscriptAlef) + ")?\B|" + MakeUniRegEx(ArabicLetterTehMarbuta) + "(" + MakeRegMultiEx(Array.ConvertAll(ArabicFathaDammaKasra, Function(Str As String) MakeUniRegEx(Str))) + "|" + MakeRegMultiEx(Array.ConvertAll(ArabicTanweens, Function(Str As String) MakeUniRegEx(Str))) + ")?\B").Count <> 0 Then
            Diagnostics.Debug.Print("Can only appear at end of word: " + ArabicString)
        End If
        If System.Text.RegularExpressions.Regex.Matches(ArabicString, "\b[" + MakeUniRegEx(ArabicLetterAlef) + MakeUniRegEx(ArabicShadda) + MakeUniRegEx(ArabicSukun) + MakeRegMultiEx(Array.ConvertAll(ArabicFathaDammaKasra, Function(Str As String) MakeUniRegEx(Str))) + MakeRegMultiEx(Array.ConvertAll(ArabicTanweens, Function(Str As String) MakeUniRegEx(Str))) + "]").Count <> 0 Then
            Diagnostics.Debug.Print("Must not appear at beginning of word: " + ArabicString)
        End If
        If System.Text.RegularExpressions.Regex.Matches(ArabicString, "\B" + MakeUniRegEx(ArabicLetterAlefWasla)).Count <> 0 Then
            Diagnostics.Debug.Print("Must appear at beginning of word: " + ArabicString)
        End If
        If System.Text.RegularExpressions.Regex.Matches(ArabicString, MakeUniRegEx(ArabicDamma) + MakeUniRegEx(ArabicLetterYeh) + "|" + MakeUniRegEx(ArabicKasra) + MakeUniRegEx(ArabicLetterWaw)).Count <> 0 Then
            Diagnostics.Debug.Print("Not a valid combination: " + ArabicString)
        End If
        If System.Text.RegularExpressions.Regex.Matches(ArabicString, "(" + MakeUniRegEx(ArabicLetterAlef) + MakeUniRegEx(ArabicMaddahAbove) + "|" + MakeUniRegEx(ArabicLetterHamza) + MakeUniRegEx(ArabicFatha) + MakeUniRegEx(ArabicLetterAlef) + "|" + MakeUniRegEx(ArabicLetterAlef) + MakeUniRegEx(ArabicLetterHamza) + "|" + MakeUniRegEx(ArabicLetterWaw) + MakeUniRegEx(ArabicLetterHamza) + "|" + MakeUniRegEx(ArabicLetterAlefMaksura) + MakeUniRegEx(ArabicLetterHamza) + ")").Count <> 0 Then
            Diagnostics.Debug.Print("Needs to be recomposed: " + ArabicString)
        End If
        Dim Count As Integer
        Dim MetadataList As New Generic.List(Of RuleMetadata)
        For Count = 0 To RulesOfRecitationRegEx.Length - 1
            MetadataList.AddRange(RulesOfRecitationRegEx(Count).Evaluator(System.Text.RegularExpressions.Regex.Match(ArabicString, RulesOfRecitationRegEx(Count).Match)))
        Next
        For Count = 0 To BreakdownRules.Length - 1
            ArabicString = System.Text.RegularExpressions.Regex.Replace(ArabicString, BreakdownRules(Count).Match, BreakdownRules(Count).Evaluator)
        Next
        For Count = 0 To ColoringSpelledOutRules.Length - 1
            ArabicString = System.Text.RegularExpressions.Regex.Replace(ArabicString, ColoringSpelledOutRules(Count).Match, ColoringSpelledOutRules(Count).Evaluator)
        Next
        For Count = 0 To RomanizationRules.Length - 1
            ArabicString = System.Text.RegularExpressions.Regex.Replace(ArabicString, RomanizationRules(Count).Match, RomanizationRules(Count).Evaluator)
        Next

        'process madda loanwords and names
        'process loanwords and names
        RomanString = ArabicString
        'Dim Count As Integer
        'Dim Index As Integer
        'Dim WordStart As Integer = 0
        'Dim PreviousIndex As Integer = -1
        'Dim PreviousPreviousIndex As Integer = -1
        'For Count = 0 To ArabicString.Length - 1
        '    Index = FindLetterBySymbol(ArabicString(Count))
        '    If Index = -1 Then
        '        RomanString += ArabicString(Count)
        '    ElseIf IsIgnored(Index) Or IsWhitespace(Index) Then
        '        WordStart = Count + 1
        '        RomanString += CachedData.IslamData.ArabicLetters(Index).PlainRoman
        '        Continue For 'do not record as previous?
        '    ElseIf CachedData.IslamData.ArabicLetters(Index).Symbol = ArabicEndOfAyah Then
        '        Dim NextIndex As Integer
        '        Do
        '            NextIndex = FindLetterBySymbol(ArabicString(Count + 1))
        '            Count += 1
        '        Loop While (ArabicString.Length - 1 <> Count) And NextIndex <> -1 AndAlso Not IsWhitespace(NextIndex)
        '    ElseIf CachedData.IslamData.ArabicLetters(Index).Symbol = ArabicShadda And PreviousIndex <> -1 Then
        '        'double last non-diacritic letter and not a letter assimalated to the noun particle
        '        If WordStart = Count - 3 AndAlso FindLetterBySymbol(ArabicString(Count - 2)) <> -1 AndAlso _
        '            CachedData.IslamData.ArabicLetters(FindLetterBySymbol(ArabicString(Count - 2))).Symbol = ArabicLetterLam AndAlso FindLetterBySymbol(ArabicString(Count - 3)) <> -1 AndAlso _
        '            (CachedData.IslamData.ArabicLetters(FindLetterBySymbol(ArabicString(Count - 3))).Symbol = ArabicLetterAlef Or _
        '             CachedData.IslamData.ArabicLetters(FindLetterBySymbol(ArabicString(Count - 3))).Symbol = ArabicLetterAlefWasla) Then
        '        Else
        '            RomanString += "-" + CachedData.IslamData.ArabicLetters(PreviousIndex).PlainRoman
        '        End If
        '    ElseIf (CachedData.IslamData.ArabicLetters(Index).Symbol = ArabicLetterAlef Or _
        '            CachedData.IslamData.ArabicLetters(Index).Symbol = ArabicLetterAlefWasla Or _
        '            CachedData.IslamData.ArabicLetters(Index).Symbol = ArabicLetterSuperscriptAlef) And _
        '        PreviousIndex <> -1 AndAlso _
        '    (CachedData.IslamData.ArabicLetters(PreviousIndex).Symbol = ArabicFatha Or _
        '     CachedData.IslamData.ArabicLetters(PreviousIndex).Symbol = ArabicDamma Or _
        '     CachedData.IslamData.ArabicLetters(PreviousIndex).Symbol = ArabicKasra Or _
        '     CachedData.IslamData.ArabicLetters(PreviousIndex).Symbol = ArabicLetterAlef AndAlso PreviousPreviousIndex <> -1 AndAlso _
        '     CachedData.IslamData.ArabicLetters(PreviousPreviousIndex).Symbol = ArabicFatha Or _
        '     CachedData.IslamData.ArabicLetters(PreviousIndex).Symbol = ArabicLetterWaw AndAlso PreviousPreviousIndex <> -1 AndAlso _
        '     CachedData.IslamData.ArabicLetters(PreviousPreviousIndex).Symbol = ArabicDamma Or _
        '     CachedData.IslamData.ArabicLetters(PreviousIndex).Symbol = ArabicLetterYeh AndAlso PreviousPreviousIndex <> -1 AndAlso _
        '     CachedData.IslamData.ArabicLetters(PreviousPreviousIndex).Symbol = ArabicKasra) Then
        '        'drop if previous letter is vowel diacritic or long vowel 
        '    ElseIf CachedData.IslamData.ArabicLetters(Index).Symbol = ArabicLetterTehMarbuta Then
        '        'add the t if a feminine genitive construct
        '        RomanString += CachedData.IslamData.ArabicLetters(Index).PlainRoman
        '        If (ArabicString.Length - 1 = Count) Then
        '            RomanString += "h"
        '        Else
        '            Dim NextIndex As Integer = FindLetterBySymbol(ArabicString(Count + 1))
        '            If NextIndex <> -1 AndAlso _
        '                (IsWhitespace(NextIndex) And _
        '                (ArabicString.Length - 2 <> Count)) Then
        '                NextIndex = FindLetterBySymbol(ArabicString(Count + 2))
        '            End If
        '            If NextIndex = -1 OrElse IsStop(NextIndex) OrElse IsPunctuation(NextIndex) Then
        '                RomanString += "h"
        '            Else
        '                RomanString += "t"
        '                Diagnostics.Debug.Print("Invalid Teh Marbuta in middle of word: " + ArabicString.Substring(0, Count) + "<" + ArabicString.Substring(Count + 1))
        '            End If
        '        End If
        '    ElseIf CachedData.IslamData.ArabicLetters(Index).Symbol = ArabicLetterAlef And (ArabicString.Length - 1 <> Count) Then
        '        'drop if next letter is vowel diacritic
        '        Dim NextIndex As Integer = FindLetterBySymbol(ArabicString(Count + 1))
        '        If NextIndex = -1 Then
        '            RomanString += CachedData.IslamData.ArabicLetters(Index).PlainRoman
        '        ElseIf CachedData.IslamData.ArabicLetters(NextIndex).Symbol = ArabicFatha Or _
        '                    CachedData.IslamData.ArabicLetters(NextIndex).Symbol = ArabicDamma Or _
        '                    CachedData.IslamData.ArabicLetters(NextIndex).Symbol = ArabicKasra Or _
        '                    CachedData.IslamData.ArabicLetters(NextIndex).Symbol = ArabicFathatan Or _
        '                    CachedData.IslamData.ArabicLetters(NextIndex).Symbol = ArabicDammatan Or _
        '                    CachedData.IslamData.ArabicLetters(NextIndex).Symbol = ArabicKasratan Then
        '        ElseIf CachedData.IslamData.ArabicLetters(NextIndex).Symbol = ArabicHamzaAbove Then
        '        Else
        '            RomanString += CachedData.IslamData.ArabicLetters(Index).PlainRoman
        '        End If
        '    ElseIf CachedData.IslamData.ArabicLetters(Index).Symbol = ArabicLetterLam And _
        '        (Count - 1 = WordStart AndAlso _
        '         (CachedData.IslamData.ArabicLetters(PreviousIndex).Symbol = ArabicLetterAlef Or _
        '            CachedData.IslamData.ArabicLetters(PreviousIndex).Symbol = ArabicLetterAlefWasla)) Then
        '        'if previous letter is >alif and beginning of word and following letter is assimilating then assimilate
        '        Dim NextIndex As Integer = FindLetterBySymbol(ArabicString(Count + 1))
        '        If NextIndex = -1 Then
        '            RomanString += CachedData.IslamData.ArabicLetters(Index).PlainRoman
        '        ElseIf CachedData.IslamData.ArabicLetters(NextIndex).Symbol = ArabicShadda Then
        '        ElseIf CachedData.IslamData.ArabicLetters(NextIndex).Assimilate Then
        '            If FindLetterBySymbol(ArabicString(Count + 2)) = -1 OrElse CachedData.IslamData.ArabicLetters(FindLetterBySymbol(ArabicString(Count + 2))).Symbol <> ArabicShadda Then
        '                Diagnostics.Debug.Print("Missing Shadda after assimilating letter: " + ArabicString.Substring(0, Count) + "<" + ArabicString.Substring(Count + 1))
        '            End If
        '            RomanString += CachedData.IslamData.ArabicLetters(NextIndex).PlainRoman + "-"
        '            Else
        '                RomanString += CachedData.IslamData.ArabicLetters(Index).PlainRoman + "-"
        '            End If
        '    ElseIf CachedData.IslamData.ArabicLetters(Index).Symbol = ArabicFatha Or _
        '        CachedData.IslamData.ArabicLetters(Index).Symbol = ArabicDamma Or _
        '        CachedData.IslamData.ArabicLetters(Index).Symbol = ArabicKasra Then
        '        'if end of verse then drop
        '        If (ArabicString.Length - 1 = Count) Then
        '        Else
        '            'if next letter makes long sound then change
        '            Dim NextIndex As Integer = FindLetterBySymbol(ArabicString(Count + 1))
        '            If NextIndex <> -1 AndAlso IsIgnored(NextIndex) Then
        '                If (ArabicString.Length - 1 <> Count + 1) Then
        '                    NextIndex = FindLetterBySymbol(ArabicString(Count + 2))
        '                Else
        '                    NextIndex = -1
        '                End If
        '            End If
        '            If NextIndex <> -1 AndAlso CachedData.IslamData.ArabicLetters(NextIndex).Symbol = ArabicShadda Then
        '                If WordStart = Count - 3 AndAlso FindLetterBySymbol(ArabicString(Count - 2)) <> -1 AndAlso _
        '                    CachedData.IslamData.ArabicLetters(FindLetterBySymbol(ArabicString(Count - 2))).Symbol = ArabicLetterLam AndAlso FindLetterBySymbol(ArabicString(Count - 3)) <> -1 AndAlso _
        '                    (CachedData.IslamData.ArabicLetters(FindLetterBySymbol(ArabicString(Count - 3))).Symbol = ArabicLetterAlef Or _
        '                     CachedData.IslamData.ArabicLetters(FindLetterBySymbol(ArabicString(Count - 3))).Symbol = ArabicLetterAlefWasla) Then
        '                Else
        '                    RomanString += "-" + CachedData.IslamData.ArabicLetters(PreviousIndex).PlainRoman
        '                End If
        '                Count = Count + 1
        '                If (ArabicString.Length - 1 <> Count) Then NextIndex = FindLetterBySymbol(ArabicString(Count + 1))
        '            End If
        '            If NextIndex = -1 OrElse _
        '                (IsStop(NextIndex) Or IsPunctuation(NextIndex)) Then
        '            ElseIf (CachedData.IslamData.ArabicLetters(NextIndex).Symbol = ArabicLetterAlef Or _
        '                    CachedData.IslamData.ArabicLetters(NextIndex).Symbol = ArabicLetterAlefWasla Or _
        '                CachedData.IslamData.ArabicLetters(NextIndex).Symbol = ArabicLetterAlefWithHamzaBelow Or _
        '                CachedData.IslamData.ArabicLetters(NextIndex).Symbol = ArabicLetterAlefWithHamzaAbove Or _
        '                    CachedData.IslamData.ArabicLetters(NextIndex).Symbol = ArabicLetterAlefMaksura Or _
        '                    (IsWhitespace(NextIndex) And _
        '                    (ArabicString.Length - 1 <> Count) AndAlso (ArabicString.Length - 1 <> Count + 1) AndAlso FindLetterBySymbol(ArabicString(Count + 2)) <> -1 AndAlso _
        '                    (CachedData.IslamData.ArabicLetters(FindLetterBySymbol(ArabicString(Count + 2))).Symbol <> ArabicLetterAlefWasla) AndAlso Not IsStop(FindLetterBySymbol(ArabicString(Count + 2))) AndAlso Not IsPunctuation(FindLetterBySymbol(ArabicString(Count + 2))))) And _
        '                        CachedData.IslamData.ArabicLetters(Index).Symbol = ArabicFatha Then
        '                RomanString += "a"
        '                If Not IsWhitespace(NextIndex) Then
        '                    If CachedData.IslamData.ArabicLetters(NextIndex).Symbol = ArabicLetterAlefWithHamzaBelow Or _
        '                        CachedData.IslamData.ArabicLetters(NextIndex).Symbol = ArabicLetterAlefWithHamzaAbove Then RomanString += CachedData.IslamData.ArabicLetters(NextIndex).PlainRoman
        '                    Count = Count + 1
        '                    PreviousPreviousIndex = PreviousIndex
        '                    PreviousIndex = Index
        '                    Index = NextIndex
        '                    'Prefixed particle causes start of word to be advanced
        '                    If Count - 2 = WordStart AndAlso CachedData.IslamData.ArabicLetters(PreviousPreviousIndex).Symbol = ArabicLetterWaw Then
        '                        WordStart = Count
        '                    End If
        '                End If
        '            ElseIf (CachedData.IslamData.ArabicLetters(NextIndex).Symbol = ArabicLetterWaw Or _
        '                    CachedData.IslamData.ArabicLetters(NextIndex).Symbol = ArabicLetterWawWithHamzaAbove Or _
        '                    (IsWhitespace(NextIndex) And _
        '                    (ArabicString.Length - 1 <> Count + 1) AndAlso FindLetterBySymbol(ArabicString(Count + 2)) <> -1 AndAlso _
        '                    (CachedData.IslamData.ArabicLetters(FindLetterBySymbol(ArabicString(Count + 2))).Symbol <> ArabicLetterAlefWasla) AndAlso Not IsStop(FindLetterBySymbol(ArabicString(Count + 2))) AndAlso Not IsPunctuation(FindLetterBySymbol(ArabicString(Count + 2))))) And _
        '                CachedData.IslamData.ArabicLetters(Index).Symbol = ArabicDamma Then
        '                RomanString += "oo"
        '                If Not IsWhitespace(NextIndex) Then
        '                    If CachedData.IslamData.ArabicLetters(NextIndex).Symbol = ArabicLetterWawWithHamzaAbove Then RomanString += CachedData.IslamData.ArabicLetters(NextIndex).PlainRoman
        '                    Count = Count + 1
        '                    PreviousPreviousIndex = PreviousIndex
        '                    PreviousIndex = Index
        '                    Index = NextIndex
        '                End If
        '            ElseIf (CachedData.IslamData.ArabicLetters(NextIndex).Symbol = ArabicLetterYeh Or _
        '                    CachedData.IslamData.ArabicLetters(NextIndex).Symbol = ArabicLetterYehWithHamzaAbove Or _
        '                    CachedData.IslamData.ArabicLetters(NextIndex).Symbol = ArabicLetterAlefMaksura Or _
        '                (IsWhitespace(NextIndex) And _
        '                    (ArabicString.Length - 1 <> Count + 1) AndAlso FindLetterBySymbol(ArabicString(Count + 2)) <> -1 AndAlso _
        '                    (CachedData.IslamData.ArabicLetters(FindLetterBySymbol(ArabicString(Count + 2))).Symbol <> ArabicLetterAlefWasla) AndAlso Not IsStop(FindLetterBySymbol(ArabicString(Count + 2))) AndAlso Not IsPunctuation(FindLetterBySymbol(ArabicString(Count + 2))))) And _
        '                CachedData.IslamData.ArabicLetters(Index).Symbol = ArabicKasra Then
        '                RomanString += "ee"
        '                If Not IsWhitespace(NextIndex) Then
        '                    If CachedData.IslamData.ArabicLetters(NextIndex).Symbol = ArabicLetterYehWithHamzaAbove Then RomanString += CachedData.IslamData.ArabicLetters(NextIndex).PlainRoman
        '                    Count = Count + 1
        '                    PreviousPreviousIndex = PreviousIndex
        '                    PreviousIndex = Index
        '                    Index = NextIndex
        '                End If
        '            ElseIf (PreviousIndex <> -1 AndAlso _
        '                        IsAmbiguousGutteral(PreviousIndex, False) Or _
        '                    IsAmbiguousGutteral(NextIndex, True)) Then
        '                If (CachedData.IslamData.ArabicLetters(Index).Symbol = ArabicFatha) Then
        '                    If (PreviousIndex <> -1 AndAlso _
        '                        IsAmbiguousGutteral(PreviousIndex, False)) Then
        '                        RomanString += "aw"
        '                    End If
        '                    If IsAmbiguousGutteral(NextIndex, True) Then
        '                        RomanString += "ah"
        '                    End If
        '                ElseIf (CachedData.IslamData.ArabicLetters(Index).Symbol = ArabicDamma) Then
        '                    If (PreviousIndex <> -1 AndAlso _
        '                        IsAmbiguousGutteral(PreviousIndex, False)) Then
        '                        RomanString += "o"
        '                    End If
        '                    If IsAmbiguousGutteral(NextIndex, True) Then
        '                        RomanString += "o"
        '                    End If
        '                ElseIf (CachedData.IslamData.ArabicLetters(Index).Symbol = ArabicKasra) Then
        '                    If (PreviousIndex <> -1 AndAlso _
        '                        IsAmbiguousGutteral(PreviousIndex, False)) Then
        '                        RomanString += "e"
        '                    End If
        '                    If IsAmbiguousGutteral(NextIndex, True) Then
        '                        RomanString += "k"
        '                    End If
        '                End If
        '                ElseIf IsWhitespace(NextIndex) And _
        '                    (ArabicString.Length - 2 <> Count) Then
        '                    NextIndex = FindLetterBySymbol(ArabicString(Count + 2))
        '                    If NextIndex = -1 OrElse (IsStop(NextIndex) Or IsPunctuation(NextIndex)) Then
        '                    Else
        '                        RomanString += CachedData.IslamData.ArabicLetters(Index).PlainRoman
        '                    End If
        '                Else
        '                    RomanString += CachedData.IslamData.ArabicLetters(Index).PlainRoman
        '                End If
        '            End If
        '            ElseIf CachedData.IslamData.ArabicLetters(Index).Symbol = ArabicLetterNoon Or _
        '                CachedData.IslamData.ArabicLetters(Index).Symbol = ArabicFathatan Or _
        '                    CachedData.IslamData.ArabicLetters(Index).Symbol = ArabicDammatan Or _
        '                    CachedData.IslamData.ArabicLetters(Index).Symbol = ArabicKasratan Then
        '            'if end of verse then drop
        '            If (ArabicString.Length - 1 = Count) Then
        '                If CachedData.IslamData.ArabicLetters(Index).Symbol = ArabicLetterNoon Then
        '                    RomanString += CachedData.IslamData.ArabicLetters(Index).PlainRoman
        '                End If
        '            Else
        '                'if next letter is Idghaam without ghunnah then drop the -n
        '                'if next letter is Iqlaab then change to miym
        '                Dim NextIndex As Integer = FindLetterBySymbol(ArabicString(Count + 1))
        '                If NextIndex <> -1 AndAlso IsIgnored(NextIndex) Then
        '                    If (ArabicString.Length - 1 <> Count + 1) Then
        '                        NextIndex = FindLetterBySymbol(ArabicString(Count + 2))
        '                    Else
        '                        NextIndex = -1
        '                    End If
        '                End If
        '                If NextIndex <> -1 AndAlso IsWhitespace(NextIndex) And _
        '                        (ArabicString.Length - 2 <> Count) Then
        '                    NextIndex = FindLetterBySymbol(ArabicString(Count + 2))
        '                End If
        '                If NextIndex = -1 OrElse _
        '                  (CachedData.IslamData.ArabicLetters(Index).Symbol <> ArabicLetterNoon And _
        '                   (IsStop(NextIndex) Or IsPunctuation(NextIndex))) Then
        '            ElseIf CachedData.IslamData.ArabicLetters(Index).Symbol = ArabicFathatan And _
        '                CachedData.IslamData.ArabicLetters(FindLetterBySymbol(ArabicString(Count + 1))).Symbol = ArabicLetterAlef Then
        '                RomanString += CachedData.IslamData.ArabicLetters(CInt(IIf(ArabicString.Length - 2 = Count, NextIndex, Index))).PlainRoman
        '                Count = Count + 1
        '                PreviousPreviousIndex = PreviousIndex
        '                PreviousIndex = Index
        '                Index = NextIndex
        '                ElseIf CachedData.IslamData.ArabicLetters(NextIndex).Symbol = ArabicLetterLam Or _
        '                            CachedData.IslamData.ArabicLetters(NextIndex).Symbol = ArabicLetterReh Then
        '                    RomanString += CachedData.IslamData.ArabicLetters(Index).PlainRoman.Substring(0, CachedData.IslamData.ArabicLetters(Index).PlainRoman.Length - 1)
        '                ElseIf CachedData.IslamData.ArabicLetters(NextIndex).Symbol = ArabicLetterBeh And _
        '                    (ArabicString.Length - 1 <> Count + 1) AndAlso FindLetterBySymbol(ArabicString(Count + 2)) <> -1 AndAlso _
        '                    (CachedData.IslamData.ArabicLetters(FindLetterBySymbol(ArabicString(Count + 2))).Symbol = ArabicFatha Or _
        '                      CachedData.IslamData.ArabicLetters(FindLetterBySymbol(ArabicString(Count + 2))).Symbol = ArabicDamma Or _
        '                      CachedData.IslamData.ArabicLetters(FindLetterBySymbol(ArabicString(Count + 2))).Symbol = ArabicKasra) Then
        '                    RomanString += "m"
        '                Else
        '                    RomanString += CachedData.IslamData.ArabicLetters(Index).PlainRoman
        '                End If
        '            End If
        '    Else
        '        If CachedData.IslamData.ArabicLetters(Index).Symbol = ArabicLetterAlef Or _
        '        CachedData.IslamData.ArabicLetters(Index).Symbol = ArabicLetterAlefWasla Or _
        '        CachedData.IslamData.ArabicLetters(Index).Symbol = ArabicLetterSuperscriptAlef Or _
        '        CachedData.IslamData.ArabicLetters(Index).Symbol = ArabicLetterAlefWithHamzaAbove Or _
        '        CachedData.IslamData.ArabicLetters(Index).Symbol = ArabicLetterAlefWithHamzaBelow Or _
        '        CachedData.IslamData.ArabicLetters(Index).Symbol = ArabicLetterAlefMaksura Or _
        '        CachedData.IslamData.ArabicLetters(Index).Symbol = ArabicLetterWaw Or _
        '        CachedData.IslamData.ArabicLetters(Index).Symbol = ArabicLetterWawWithHamzaAbove Or _
        '        CachedData.IslamData.ArabicLetters(Index).Symbol = ArabicLetterYeh Or _
        '        CachedData.IslamData.ArabicLetters(Index).Symbol = ArabicLetterYehWithHamzaAbove Or IsLetter(Index) Then
        '            If PreviousIndex <> -1 AndAlso ( _
        '                Not IsWhitespace(PreviousIndex) Or _
        '                CachedData.IslamData.ArabicLetters(PreviousIndex).Symbol <> ArabicFatha Or _
        '                CachedData.IslamData.ArabicLetters(PreviousIndex).Symbol <> ArabicDamma Or _
        '                CachedData.IslamData.ArabicLetters(PreviousIndex).Symbol <> ArabicKasra Or _
        '                CachedData.IslamData.ArabicLetters(PreviousIndex).Symbol <> ArabicSukun) Then
        '                Diagnostics.Debug.Print("Missing diacritic: " + ArabicString.Substring(0, Count) + "<" + ArabicString.Substring(Count + 1))
        '            End If
        '        End If
        '        RomanString += CachedData.IslamData.ArabicLetters(Index).PlainRoman
        '        End If
        '        PreviousPreviousIndex = PreviousIndex
        '        PreviousIndex = Index
        'Next
        Return RomanString
    End Function
    Shared Function GetArabicSymbolJSArray() As String
        Dim Letters(CachedData.IslamData.ArabicLetters.Length - 1) As IslamData.ArabicSymbol
        CachedData.IslamData.ArabicLetters.CopyTo(Letters, 0)
        Array.Sort(Letters, New StringLengthComparer)
        GetArabicSymbolJSArray = "var ArabicLetters = " + _
                                Utility.MakeJSArray(New String() {Utility.MakeJSIndexedObject(New String() {"Symbol", "Shaping", "Assimilate", "TranslitLetter", "RomanTranslit", "PlainRoman"}, _
                                Array.ConvertAll(Of IslamData.ArabicSymbol, String())(Letters, Function(Convert As IslamData.ArabicSymbol) New String() {CStr(AscW(Convert.Symbol)), If(Convert.Shaping = Nothing, String.Empty, Utility.MakeJSArray(Array.ConvertAll(Convert.Shaping, Function(Ch As Char) CStr(AscW(Ch))))), CStr(IIf(Convert.Assimilate, "true", String.Empty)), CStr(IIf(Convert.ExtendedBuckwalterLetter = ChrW(0), String.Empty, Convert.ExtendedBuckwalterLetter)), Convert.RomanTranslit, Convert.PlainRoman}), False)}, True) + ";"
    End Function
    Public Shared Function GetTransliterateGenJS() As String
        Return "function doTransliterate(sVal, direction, conversion) { var iCount, iSubCount, sOutVal = ''; for (iCount = 0; iCount < sVal.length; iCount++) { for (iSubCount = 0; iSubCount < ArabicLetters.length; iSubCount++) { if (direction ? sVal.charCodeAt(iCount) === parseInt(ArabicLetters[iSubCount].Symbol, 10) : sVal.charAt(iCount) === unescape((conversion ? ArabicLetters[iSubCount].TranslitLetter : ArabicLetters[iSubCount].RomanTranslit))) { sOutVal += (direction ? (conversion ? ArabicLetters[iSubCount].TranslitLetter : ArabicLetters[iSubCount].RomanTranslit) : String.fromCharCode(ArabicLetters[iSubCount].Symbol)); break; } } if (iSubCount === ArabicLetters.length) sOutVal += sVal.charAt(iCount); } return unescape(sOutVal); }"
    End Function
    Public Shared Function GetDiacriticJS() As String
        Return "function doDiacritics(sVal, direction) { var iCount, sOutVal = ''; for (iCount = 0; iCount < sVal.length; iCount++) { sOutVal += findLetterBySymbol(sVal.charCodeAt(iCount)) === -1 || !isDiacritic(findLetterBySymbol(sVal.charCodeAt(iCount))) ? sVal[iCount] : ''; } return sOutVal; }" + _
            IsDiacriticJS() + _
            FindLetterBySymbolJS()
    End Function
    Public Shared Function FindLetterBySymbolJS() As String
        Return "function findLetterBySymbol(chVal) { var iSubCount; for (iSubCount = 0; iSubCount < ArabicLetters.length; iSubCount++) { if (chVal === parseInt(ArabicLetters[iSubCount].Symbol, 10)) return iSubCount; for (var iShapeCount = 0; iShapeCount < ArabicLetters[iSubCount].Shaping.length; iShapeCount++) { if (chVal === parseInt(ArabicLetters[iSubCount].Shaping[iShapeCount], 10)) return iSubCount; } } return -1; }"
    End Function
    Public Shared Function IsDiacriticJS() As String
        Return "function isDiacritic(index) { return (parseInt(ArabicLetters[index].Symbol, 10) === 1611 || parseInt(ArabicLetters[index].Symbol, 10) === 1612 || parseInt(ArabicLetters[index].Symbol, 10) === 1613 || parseInt(ArabicLetters[index].Symbol, 10) === 1614 || parseInt(ArabicLetters[index].Symbol, 10) === 1615 || parseInt(ArabicLetters[index].Symbol, 10) === 1616 || parseInt(ArabicLetters[index].Symbol, 10) === 1617 || parseInt(ArabicLetters[index].Symbol, 10) === 1618 || parseInt(ArabicLetters[index].Symbol, 10) === 1619 || parseInt(ArabicLetters[index].Symbol, 10) === 1620 || parseInt(ArabicLetters[index].Symbol, 10) === 1621); }"
    End Function
    Public Shared Function GetPlainTransliterateGenJS() As String
        Return FindLetterBySymbolJS() + _
            IsDiacriticJS() + _
            "function isIgnored(index) { return (parseInt(ArabicLetters[index].Symbol, 10) === 1765 || parseInt(ArabicLetters[index].Symbol, 10) === 1762); }" + _
            "function isWhitespace(index) { return (parseInt(ArabicLetters[index].Symbol, 10) === 32); }" + _
            "function isPunctuation(index) { return (parseInt(ArabicLetters[index].Symbol, 10) === 33 || parseInt(ArabicLetters[index].Symbol, 10) === 34 || parseInt(ArabicLetters[index].Symbol, 10) === 46 || parseInt(ArabicLetters[index].Symbol, 10) === 44 || parseInt(ArabicLetters[index].Symbol, 10) === 1548); }" + _
            "function isStop(index) { return (parseInt(ArabicLetters[index].Symbol, 10) === 1750 || parseInt(ArabicLetters[index].Symbol, 10) === 1751 || parseInt(ArabicLetters[index].Symbol, 10) === 1752 || parseInt(ArabicLetters[index].Symbol, 10) === 1753 || parseInt(ArabicLetters[index].Symbol, 10) === 1754 || parseInt(ArabicLetters[index].Symbol, 10) === 1755 || parseInt(ArabicLetters[index].Symbol, 10) === 1756 || parseInt(ArabicLetters[index].Symbol, 10) === 1757); }" + _
            "function doPlainTransliterate(sVal) { var iCount, index, wordStart = 0, previousIndex = -1, previousPreviousIndex = -1, sOutVal = ''; for (iCount = 0; iCount < sVal.length; iCount++) { index = findLetterBySymbol(sVal.charCodeAt(iCount)); if (index === -1) { sOutVal += sVal.charAt(iCount); " + _
            "} else if (isIgnored(index) || isWhitespace(index)) { wordStart = iCount + 1; sOutVal += ArabicLetters[index].PlainRoman; continue; " + _
            "} else if (parseInt(ArabicLetters[index].Symbol, 10) === 1757) { var nextIndex; do { nextIndex = findLetterBySymbol(sVal.charCodeAt(iCount + 1)); iCount++; } while (sVal.length - 1 !== iCount && nextIndex !== -1 && !isWhitespace(nextIndex)); " + _
            "} else if (parseInt(ArabicLetters[index].Symbol, 10) === 1617 && previousIndex !== -1) { if (wordStart === iCount - 3 && findLetterBySymbol(sVal.charCodeAt(iCount - 2)) !== -1 && parseInt(ArabicLetters[findLetterBySymbol(sVal.charCodeAt(iCount - 2))].Symbol, 10) === 1604 && findLetterBySymbol(sVal.charCodeAt(iCount - 3)) !== -1 && (parseInt(ArabicLetters[findLetterBySymbol(sVal.charCodeAt(iCount - 3))].Symbol, 10) === 1575 || parseInt(ArabicLetters[findLetterBySymbol(sVal.charCodeAt(iCount - 3))].Symbol, 10) === 1649)) {} else { sOutVal += '-' + ArabicLetters[previousIndex].PlainRoman; } " + _
            "} else if ((parseInt(ArabicLetters[index].Symbol, 10) === 1575 || parseInt(ArabicLetters[index].Symbol, 10) === 1649 || parseInt(ArabicLetters[index].Symbol, 10) === 1648) && previousIndex !== -1 && (parseInt(ArabicLetters[previousIndex].Symbol, 10) === 1614 || parseInt(ArabicLetters[previousIndex].Symbol, 10) === 1615 || parseInt(ArabicLetters[previousIndex].Symbol, 10) === 1616 || parseInt(ArabicLetters[previousIndex].Symbol, 10) === 1575 && parseInt(ArabicLetters[previousPreviousIndex].Symbol, 10) === 1614 || parseInt(ArabicLetters[previousIndex].Symbol, 10) === 1608 && parseInt(ArabicLetters[previousPreviousIndex].Symbol, 10) === 1615 || parseInt(ArabicLetters[previousIndex].Symbol, 10) === 1610 && parseInt(ArabicLetters[previousPreviousIndex].Symbol, 10) === 1616)) { " + _
            "} else if (parseInt(ArabicLetters[index].Symbol, 10) === 1577) { sOutVal += ArabicLetters[index].PlainRoman; if (sVal.length - 1 === iCount) { sOutVal += 'h'; } else { var nextIndex = findLetterBySymbol(sVal.charCodeAt(iCount + 1)); if (nextIndex !== -1 && isWhitespace(nextIndex) && sVal.length - 2 !== iCount) nextIndex = findLetterBySymbol(sVal.charCodeAt(iCount + 2)); if (nextIndex === -1 || isStop(nextIndex) || isPunctuation(nextIndex)) { sOutVal += 'h'; } else { sOutVal += 't'; } } " +
            "} else if (parseInt(ArabicLetters[index].Symbol, 10) === 1575 && sVal.length - 1 !== iCount) { var nextIndex = findLetterBySymbol(sVal.charCodeAt(iCount + 1)); if (nextIndex === -1) { sOutVal += ArabicLetters[index].PlainRoman; } else if (parseInt(ArabicLetters[nextIndex].Symbol, 10) === 1614 || parseInt(ArabicLetters[nextIndex].Symbol, 10) === 1615 || parseInt(ArabicLetters[nextIndex].Symbol, 10) === 1616 || parseInt(ArabicLetters[nextIndex].Symbol, 10) === 1611 || parseInt(ArabicLetters[nextIndex].Symbol, 10) === 1612 || parseInt(ArabicLetters[nextIndex].Symbol, 10) === 1613) {} else if (parseInt(ArabicLetters[nextIndex].Symbol, 10) === 1620) {} else { sOutVal += ArabicLetters[index].PlainRoman; } " +
            "} else if (parseInt(ArabicLetters[index].Symbol, 10) === 1604 && (iCount - 1 === wordStart && (parseInt(ArabicLetters[previousIndex].Symbol, 10) === 1575 || parseInt(ArabicLetters[previousIndex].Symbol, 10) === 1649))) { var nextIndex = findLetterBySymbol(sVal.charCodeAt(iCount + 1)); if (nextIndex === -1) { sOutVal += ArabicLetters[index].PlainRoman; } else if (parseInt(ArabicLetters[nextIndex].Symbol, 10) === 1617) {} else if (ArabicLetters[nextIndex].Assimilate) { sOutVal += ArabicLetters[nextIndex].PlainRoman + '-'; } else { sOutVal += ArabicLetters[index].PlainRoman + '-'; } " +
            "} else if (parseInt(ArabicLetters[index].Symbol, 10) === 1614 || parseInt(ArabicLetters[index].Symbol, 10) === 1615 || parseInt(ArabicLetters[index].Symbol, 10) === 1616) { if (sVal.length - 1 === iCount) {} else { var nextIndex = findLetterBySymbol(sVal.charCodeAt(iCount + 1)); if (nextIndex !== -1 && isIgnored(nextIndex)) { if (sVal.length - 1 !== iCount + 1) { nextIndex = findLetterBySymbol(sVal.charCodeAt(iCount + 2)); } else { nextIndex = -1; } } if (nextIndex !== -1 && parseInt(ArabicLetters[nextIndex].Symbol, 10) === 1617) { if (wordStart === iCount - 3 && findLetterBySymbol(sVal.charCodeAt(iCount - 2)) !== -1 && parseInt(ArabicLetters[findLetterBySymbol(sVal.charCodeAt(iCount - 2))].Symbol, 10) === 1604 && findLetterBySymbol(sVal.charCodeAt(iCount - 3)) !== -1 && (parseInt(ArabicLetters[findLetterBySymbol(sVal.charCodeAt(iCount - 3))].Symbol, 10) === 1575 || parseInt(ArabicLetters[findLetterBySymbol(sVal.charCodeAt(iCount - 3))].Symbol, 10) === 1649)) {} else { sOutVal += '-' + ArabicLetters[previousIndex].PlainRoman; } iCount += 1; if (sVal.length - 1 !== iCount) nextIndex = findLetterBySymbol(sVal.charCodeAt(iCount + 1)); } if (nextIndex === -1 || isStop(nextIndex) || isPunctuation(nextIndex)) {} " + _
                "else if ((parseInt(ArabicLetters[nextIndex].Symbol, 10) === 1575 || parseInt(ArabicLetters[nextIndex].Symbol, 10) === 1649 || parseInt(ArabicLetters[nextIndex].Symbol, 10) === 1573 || parseInt(ArabicLetters[nextIndex].Symbol, 10) === 1571 || parseInt(ArabicLetters[nextIndex].Symbol, 10) === 1609 || (isWhitespace(nextIndex) && sVal.length - 1 !== iCount && sVal.length - 1 !== iCount + 1 && findLetterBySymbol(sVal.charCodeAt(iCount + 2)) !== -1 && parseInt(ArabicLetters[findLetterBySymbol(sVal.charCodeAt(iCount + 2))].Symbol, 10) === 1649 && !isStop(findLetterBySymbol(sVal.charCodeAt(iCount + 2))) && !isPunctuation(findLetterBySymbol(sVal.charCodeAt(iCount + 2))))) && parseInt(ArabicLetters[index].Symbol, 10) === 1614) { sOutVal += 'a'; if (!isWhitespace(nextIndex)) { if (parseInt(ArabicLetters[nextIndex].Symbol, 10) === 1573 || parseInt(ArabicLetters[nextIndex].Symbol, 10) === 1571) sOutVal += ArabicLetters[nextIndex].PlainRoman; iCount++; previousPreviousIndex = previousIndex; previousIndex = index; index = nextIndex; if (iCount - 2 === wordStart && parseInt(ArabicLetters[previousPreviousIndex].Symbol, 10) === 1608) wordStart = iCount; } } " +
                "else if ((parseInt(ArabicLetters[nextIndex].Symbol, 10) === 1608 || parseInt(ArabicLetters[nextIndex].Symbol, 10) === 1572 || (isWhitespace(nextIndex) && sVal.length - 1 !== iCount + 1 && findLetterBySymbol(sVal.charCodeAt(iCount + 2)) !== -1 && parseInt(ArabicLetters[findLetterBySymbol(sVal.charCodeAt(iCount + 2))].Symbol, 10) !== 1649 && !isStop(findLetterBySymbol(sVal.charCodeAt(iCount + 2))) && !isPunctuation(findLetterBySymbol(sVal.charCodeAt(iCount + 2))))) && parseInt(ArabicLetters[index].Symbol, 10) === 1615) { sOutVal += 'oo'; if (!isWhitespace(nextIndex)) { if (parseInt(ArabicLetters[nextIndex].Symbol, 10) === 1572) sOutVal += ArabicLetters[nextIndex].PlainRoman; iCount++; previousPreviousIndex = previousIndex; previousIndex = index; index = nextIndex; } } " +
                "else if ((parseInt(ArabicLetters[nextIndex].Symbol, 10) === 1610 || parseInt(ArabicLetters[nextIndex].Symbol, 10) === 1574 || parseInt(ArabicLetters[nextIndex].Symbol, 10) === 1609 || (isWhitespace(nextIndex) && sVal.length - 1 !== iCount + 1 && findLetterBySymbol(sVal.charCodeAt(iCount + 2)) !== -1 && parseInt(ArabicLetters[findLetterBySymbol(sVal.charCodeAt(iCount + 2))].Symbol, 10) !== 1649 && !isStop(findLetterBySymbol(sVal.charCodeAt(iCount + 2))) && !isPunctuation(findLetterBySymbol(sVal.charCodeAt(iCount + 2))))) && parseInt(ArabicLetters[index].Symbol, 10) === 1616) { sOutVal += 'ee'; if (!isWhitespace(nextIndex)) { if (parseInt(ArabicLetters[nextIndex].Symbol, 10) === 1574) sOutVal += ArabicLetters[nextIndex].PlainRoman; iCount++; previousPreviousIndex = previousIndex; previousIndex = index; index = nextIndex; } } else if (isWhitespace(nextIndex) && sVal.length - 2 !== iCount) { nextIndex = findLetterBySymbol(sVal.charCodeAt(iCount + 2)); if (nextIndex === -1 || isStop(nextIndex) || isPunctuation(nextIndex)) {} else { sOutVal += ArabicLetters[index].PlainRoman; } } else { sOutVal += ArabicLetters[index].PlainRoman; } } " +
            "} else if (parseInt(ArabicLetters[index].Symbol, 10) === 1606 || parseInt(ArabicLetters[index].Symbol, 10) === 1611 || parseInt(ArabicLetters[index].Symbol, 10) === 1612 || parseInt(ArabicLetters[index].Symbol, 10) === 1613) { if (sVal.length - 1 === iCount) { if (parseInt(ArabicLetters[index].Symbol, 10) === 1606) sOutVal += ArabicLetters[index].PlainRoman; } else { var nextIndex = findLetterBySymbol(sVal.charCodeAt(iCount + 1)); if (nextIndex !== -1 && isIgnored(nextIndex)) { if (sVal.length - 1 !== iCount + 1) { nextIndex = findLetterBySymbol(sVal.charCodeAt(iCount + 2)); } else { nextIndex = -1; } } if (nextIndex !== -1 && isWhitespace(nextIndex) && sVal.length - 2 !== iCount) nextIndex = findLetterBySymbol(sVal.charCodeAt(iCount + 2)); if (nextIndex === -1 || parseInt(ArabicLetters[index].Symbol, 10) !== 1606 && (isStop(nextIndex) || isPunctuation(nextIndex))) {} else if (parseInt(ArabicLetters[index].Symbol, 10) === 1611 && parseInt(ArabicLetters[findLetterBySymbol(sVal.charCodeAt(iCount + 1))].Symbol, 10) === 1575) { sOutVal += ArabicLetters[sVal.length - 2 == iCount ? nextIndex : index].PlainRoman; iCount++; previousPreviousIndex = previousIndex; previousIndex = index; index = nextIndex; } else if (parseInt(ArabicLetters[nextIndex].Symbol, 10) === 1604 || parseInt(ArabicLetters[nextIndex].Symbol, 10) === 1585) { sOutVal += ArabicLetters[index].PlainRoman.substr(0, ArabicLetters[index].PlainRoman.length - 1); } else if (parseInt(ArabicLetters[nextIndex].Symbol, 10) === 1576 && sVal.length - 1 !== iCount + 1 && findLetterBySymbol(sVal.charCodeAt(iCount + 2)) !== -1 && (parseInt(ArabicLetters[findLetterBySymbol(sVal.charCodeAt(iCount + 2))].Symbol, 10) === 1614 || parseInt(ArabicLetters[findLetterBySymbol(sVal.charCodeAt(iCount + 2))].Symbol, 10) === 1615 || parseInt(ArabicLetters[findLetterBySymbol(sVal.charCodeAt(iCount + 2))].Symbol, 10) === 1616)) { sOutVal += 'm'; } else { sOutVal += ArabicLetters[index].PlainRoman; } } " +
            "} else { sOutVal += ArabicLetters[index].PlainRoman; } previousPreviousIndex = previousIndex; previousIndex = index; } return sOutVal; }"
    End Function
    Public Shared Function GetTransliterateJS() As String()
        Return New String() {"javascript: doTransliterateDisplay();", String.Empty, GetArabicSymbolJSArray(), GetTransliterateGenJS(), GetDiacriticJS(), _
        "function isRTL(c) { return ((c===0x05BE)||(c===0x05C0)||(c===0x05C3)||(c===0x05C6)||((c>=0x05D0)&&(c<=0x05F4))||(c===0x0608)||(c===0x060B)||(c===0x060D)||((c>=0x061B)&&(c<=0x064A))||((c>=0x066D)&&(c<=0x066F))||((c>=0x0671)&&(c<=0x06D5))||((c>=0x06E5)&&(c<=0x06E6))||((c>=0x06EE)&&(c<=0x06EF))||((c>=0x06FA)&&(c<=0x0710))||((c>=0x0712)&&(c<=0x072F))||((c>=0x074D)&&(c<=0x07A5))||((c>=0x07B1)&&(c<=0x07EA))||((c>=0x07F4)&&(c<=0x07F5))||((c>=0x07FA)&&(c<=0x0815))||(c===0x081A)||(c===0x0824)||(c===0x0828)||((c>=0x0830)&&(c<=0x0858))||((c>=0x085E)&&(c<=0x08AC))||(c===0x200F)||(c===0xFB1D)||((c>=0xFB1F)&&(c<=0xFB28))||((c>=0xFB2A)&&(c<=0xFD3D))||((c>=0xFD50)&&(c<=0xFDFC))||((c>=0xFE70)&&(c<=0xFEFC))||((c>=0x10800)&&(c<=0x1091B))||((c>=0x10920)&&(c<=0x10A00))||((c>=0x10A10)&&(c<=0x10A33))||((c>=0x10A40)&&(c<=0x10B35))||((c>=0x10B40)&&(c<=0x10C48))||((c>=0x1EE00)&&(c<=0x1EEBB))); }", _
        "function doFixDirection(sVal, direction) { var iCount, sOutVal = direction ? '\u200E' : '\u200F', bInv = false; for (iCount = 0; iCount < sVal.length; iCount++) { sOutVal += (sVal.charCodeAt(iCount) == 0x200E || sVal.charCodeAt(iCount) == 0x200F || sVal.charCodeAt(iCount) == 0x202A || sVal.charCodeAt(iCount) == 0x202B || sVal.charCodeAt(iCount) == 0x202C || sVal.charCodeAt(iCount) == 0x202D || sVal.charCodeAt(iCount) == 0x202E) ? '' : (!bInv && isRTL(sVal.charCodeAt(iCount)) === direction ? '' : (direction ? '\u202A' : '\u202B')) + sVal[iCount]; } return sOutVal; }", _
        "function doTransliterateDisplay() { $('#translitvalue').text($('#scheme0').prop('checked') ? doTransliterate($('#translitedit').val(), $('#direction0').prop('checked'), $('#translitscheme0').prop('checked')) : ($('#scheme1').prop('checked') ? doDiacritics($('#translitedit').val(), $('#diacriticscheme0').prop('checked')) : doFixDirection($('#translitedit').val(), $('#diacriticscheme0').prop('checked')))); }"}
    End Function
    Public Shared Function GetSchemeChangeJS() As String()
        Return New String() {"javascript: doSchemeChange();", String.Empty, "function doSchemeChange() { $('#diacriticscheme_').css('display', $('#scheme0').prop('checked') ? 'none' : 'block'); $('#translitscheme_').css('display', $('#scheme0').prop('checked') ? 'block' : 'none'); $('#direction_').css('display', $('#scheme0').prop('checked') ? 'block' : 'none'); }"}
    End Function
    Public Shared Function GetChangeTransliterationJS() As String()
        Return New String() {"javascript: changeTransliteration();", String.Empty, Utility.GetLookupStyleSheetJS(), GetArabicSymbolJSArray(), GetTransliterateGenJS(), GetPlainTransliterateGenJS(), _
        "function changeTransliteration() { var k, child, iSubCount, text; $('span.transliteration').each(function() { $(this).css('display', $('#translitscheme').val() === '0' ? 'none' : 'block'); }); for (k in renderList) { text = ''; for (child in renderList[k]['children']) { for (iSubCount = 0; iSubCount < renderList[k]['children'][child]['arabic'].length; iSubCount++) { if ($('#translitscheme').val() === '1' && renderList[k]['children'][child]['arabic'][iSubCount] !== '' && (renderList[k]['children'][child]['arabic'][iSubCount].length !== 1 || !isStop(findLetterBySymbol(renderList[k]['children'][child]['arabic'][iSubCount].charCodeAt(0)))) && renderList[k]['children'][child]['translit'][iSubCount] !== '') { if (text !== '') text += ' '; text += $('#' + renderList[k]['children'][child]['arabic'][iSubCount]).text(); } else { if (renderList[k]['children'][child]['translit'][iSubCount] !== '') $('#' + renderList[k]['children'][child]['translit'][iSubCount]).text(($('#translitscheme').val() === '0' || renderList[k]['children'][child]['arabic'][iSubCount] === '') ? '' : doTransliterate($('#' + renderList[k]['children'][child]['arabic'][iSubCount]).text(), true, $('#translitscheme').val() === '3')); } } } if ($('#translitscheme').val() === '1') { text = doPlainTransliterate(text).split(' '); for (child in renderList[k]['children']) { for (iSubCount = 0; iSubCount < renderList[k]['children'][child]['translit'].length; iSubCount++) { if (renderList[k]['children'][child]['translit'][iSubCount] !== '') $('#' + renderList[k]['children'][child]['translit'][iSubCount]).text(text.shift()); } } } for (iSubCount = 0; iSubCount < renderList[k]['arabic'].length; iSubCount++) { if (renderList[k]['translit'][iSubCount] !== '') $('#' + renderList[k]['translit'][iSubCount]).text(($('#translitscheme').val() === '0' || renderList[k]['arabic'][iSubCount] === '') ? '' : ($('#translitscheme').val() === '1' ? doPlainTransliterate($('#' + renderList[k]['arabic'][iSubCount]).text()) : doTransliterate($('#' + renderList[k]['arabic'][iSubCount]).text(), true, $('#translitscheme').val() === '3'))); } } }"}
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
        Dim Count As Integer = CInt(Web.HttpContext.Current.Request.QueryString.Get("selection"))
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
        Dim Output As New ArrayList
        Output.Add(GetRenderJS())
        Dim Count As Integer = CInt(Web.HttpContext.Current.Request.QueryString.Get("selection"))
        If CachedData.IslamData.VocabularyCategories.Length + 2 = Count Then
            Output.Add(New String() {"arabic", "transliteration", String.Empty, String.Empty, String.Empty, String.Empty})
            Output.Add(New String() {Utility.LoadResourceString("IslamInfo_Arabic"), Utility.LoadResourceString("IslamInfo_Transliteration"), Utility.LoadResourceString("IslamInfo_Translation"), Utility.LoadResourceString("IslamInfo_Before"), Utility.LoadResourceString("IslamInfo_PrescribedTime"), Utility.LoadResourceString("IslamInfo_After")})
        ElseIf CachedData.IslamData.VocabularyCategories.Length + 3 = Count Then
            Output.Add(New String() {"arabic", "transliteration", String.Empty, String.Empty, String.Empty})
            Output.Add(New String() {Utility.LoadResourceString("IslamInfo_Arabic"), Utility.LoadResourceString("IslamInfo_Transliteration"), Utility.LoadResourceString("IslamInfo_Translation"), Utility.LoadResourceString("IslamInfo_Classification"), Utility.LoadResourceString("IslamInfo_PrayerUnits")})
        Else
            Output.Add(New String() {"arabic", "transliteration", String.Empty})
            Output.Add(New String() {Utility.LoadResourceString("IslamInfo_Arabic"), Utility.LoadResourceString("IslamInfo_Transliteration"), Utility.LoadResourceString("IslamInfo_Translation")})
        End If
        If Count = -1 Then Count = 0
        If CachedData.IslamData.VocabularyCategories.Length = Count Then
            For SubCount As Integer = 0 To CachedData.IslamData.Months.Length - 1
                Output.Add(New String() {Arabic.RightToLeftMark + TransliterateFromBuckwalter(CachedData.IslamData.Months(SubCount).Name), TransliterateToScheme(TransliterateFromBuckwalter(CachedData.IslamData.Months(SubCount).Name)).Trim(), Utility.LoadResourceString("IslamInfo_" + CachedData.IslamData.Months(SubCount).TranslationID)})
            Next
        ElseIf CachedData.IslamData.VocabularyCategories.Length + 1 = Count Then
            For SubCount As Integer = 0 To CachedData.IslamData.DaysOfWeek.Length - 1
                Output.Add(New String() {Arabic.RightToLeftMark + TransliterateFromBuckwalter(CachedData.IslamData.DaysOfWeek(SubCount).Name), TransliterateToScheme(TransliterateFromBuckwalter(CachedData.IslamData.DaysOfWeek(SubCount).Name)).Trim(), Utility.LoadResourceString("IslamInfo_" + CachedData.IslamData.DaysOfWeek(SubCount).TranslationID)})
            Next
        ElseIf CachedData.IslamData.VocabularyCategories.Length + 2 = Count Then
            Dim Table As New Hashtable
            Array.ForEach(CachedData.IslamData.Prayers, Sub(Convert As IslamData.PrayerType) Array.ForEach(Convert.PrayerUnits.Split(","c), Sub(Part As String) If Part.Contains("="c) Then If Table.ContainsKey(Part.Substring(0, Part.IndexOf("="c))) Then Table.Item(Part.Substring(0, Part.IndexOf("="c))) = CStr(Table.Item(Part.Substring(0, Part.IndexOf("="c)))) + vbCrLf + Part.Substring(Part.IndexOf("="c) + 1).Replace("|"c, " or ") + " " + CStr(IIf(Utility.LoadResourceString("IslamInfo_" + Convert.TranslationID) <> "Prescribed time", Utility.LoadResourceString("IslamInfo_" + Convert.TranslationID) + " - ", String.Empty)) + Convert.Classification Else Table.Add(Part.Substring(0, Part.IndexOf("="c)), Part.Substring(Part.IndexOf("="c) + 1).Replace("|"c, " or ") + " " + Convert.Classification)))
            For SubCount As Integer = 0 To CachedData.IslamData.PrayerTimes.Length - 1
                Output.Add(New String() {Arabic.RightToLeftMark + TransliterateFromBuckwalter(CachedData.IslamData.PrayerTimes(SubCount).Name), TransliterateToScheme(TransliterateFromBuckwalter(CachedData.IslamData.PrayerTimes(SubCount).Name)).Trim(), Utility.LoadResourceString("IslamInfo_" + CachedData.IslamData.PrayerTimes(SubCount).TranslationID), CStr(IIf(Table.ContainsKey("-"c + Utility.LoadResourceString("IslamInfo_" + CachedData.IslamData.PrayerTimes(SubCount).TranslationID)), Table.Item("-"c + Utility.LoadResourceString("IslamInfo_" + CachedData.IslamData.PrayerTimes(SubCount).TranslationID)), String.Empty)), CStr(IIf(Table.ContainsKey(Utility.LoadResourceString("IslamInfo_" + CachedData.IslamData.PrayerTimes(SubCount).TranslationID)), Table.Item(Utility.LoadResourceString("IslamInfo_" + CachedData.IslamData.PrayerTimes(SubCount).TranslationID)), String.Empty)), CStr(IIf(Table.ContainsKey("+"c + Utility.LoadResourceString("IslamInfo_" + CachedData.IslamData.PrayerTimes(SubCount).TranslationID)), Table.Item("+"c + Utility.LoadResourceString("IslamInfo_" + CachedData.IslamData.PrayerTimes(SubCount).TranslationID)), String.Empty))})
            Next
        ElseIf CachedData.IslamData.VocabularyCategories.Length + 3 = Count Then
            For SubCount As Integer = 0 To CachedData.IslamData.Prayers.Length - 1
                Output.Add(New String() {Arabic.RightToLeftMark + TransliterateFromBuckwalter(CachedData.IslamData.Prayers(SubCount).Name), TransliterateToScheme(TransliterateFromBuckwalter(CachedData.IslamData.Prayers(SubCount).Name)).Trim(), Utility.LoadResourceString("IslamInfo_" + CachedData.IslamData.Prayers(SubCount).TranslationID), CachedData.IslamData.Prayers(SubCount).Classification, CachedData.IslamData.Prayers(SubCount).PrayerUnits})
            Next
        Else
            For SubCount As Integer = 0 To CachedData.IslamData.VocabularyCategories(Count).Words.Length - 1
                Output.Add(New String() {Arabic.RightToLeftMark + TransliterateFromBuckwalter(CachedData.IslamData.VocabularyCategories(Count).Words(SubCount).Text), TransliterateToScheme(TransliterateFromBuckwalter(CachedData.IslamData.VocabularyCategories(Count).Words(SubCount).Text)).Trim(), Utility.LoadResourceString("IslamInfo_" + CachedData.IslamData.VocabularyCategories(Count).Words(SubCount).TranslationID)})
            Next
        End If
        Return DirectCast(Output.ToArray(GetType(Array)), Array())
    End Function
    Public Shared Function DisplayAll(ByVal Item As PageLoader.TextItem) As Array()
        Dim Count As Integer
        Dim Output(CachedData.IslamData.ArabicLetters.Length + 2) As Array
        'Dim oFont As New Font(DefaultValue(Web.HttpContext.Current.Request.QueryString.Get("fontcustom"), "Arial"), 13)
        'CheckIfCharInFont(CachedData.IslamData.ArabicLetters(Count).Symbol, oFont)
        Output(0) = New String() {}
        Output(1) = New String() {"arabic", String.Empty, "arabic", String.Empty, String.Empty, String.Empty, String.Empty, String.Empty, String.Empty, String.Empty}
        Output(2) = New String() {Utility.LoadResourceString("IslamInfo_LetterName"), Utility.LoadResourceString("IslamInfo_UnicodeName"), Utility.LoadResourceString("IslamInfo_Arabic"), Utility.LoadResourceString("IslamInfo_UnicodeValue"), Utility.LoadResourceString("IslamSource_ExtendedBuckwalter"), Utility.LoadResourceString("IslamSource_EnglishRomanized"), Utility.LoadResourceString("IslamSource_PlainRoman"), Utility.LoadResourceString("IslamInfo_Terminating"), Utility.LoadResourceString("IslamInfo_Connecting"), Utility.LoadResourceString("IslamInfo_Assimilate")}
        For Count = 0 To CachedData.IslamData.ArabicLetters.Length - 1
            Output(Count + 3) = New String() {Arabic.RightToLeftMark + TransliterateFromBuckwalter(CachedData.IslamData.ArabicLetters(Count).SymbolName), _
                                              GetUnicodeName(CachedData.IslamData.ArabicLetters(Count).Symbol), _
                                       CStr(CachedData.IslamData.ArabicLetters(Count).Symbol), _
                                       CStr(AscW(CachedData.IslamData.ArabicLetters(Count).Symbol)), _
                                       CStr(IIf(CachedData.IslamData.ArabicLetters(Count).ExtendedBuckwalterLetter = ChrW(0), String.Empty, CachedData.IslamData.ArabicLetters(Count).ExtendedBuckwalterLetter)), _
                                       CachedData.IslamData.ArabicLetters(Count).RomanTranslit, _
                                       CachedData.IslamData.ArabicLetters(Count).PlainRoman, _
                                       CStr(CachedData.IslamData.ArabicLetters(Count).Terminating), _
                                       CStr(CachedData.IslamData.ArabicLetters(Count).Connecting), _
                                       CStr(CachedData.IslamData.ArabicLetters(Count).Assimilate)}
        Next
        Return Output
    End Function
    Public Shared Function DisplayParticle(Index As Integer) As Array()
        Dim Count As Integer
        Dim Output(2 + CachedData.IslamData.GrammarCategories(Index).Words.Length) As Array
        Dim Build As New Generic.Dictionary(Of String, Generic.Dictionary(Of String, String))
        Output(0) = New String() {}
        Output(1) = New String() {"arabic", String.Empty, String.Empty}
        Output(2) = New String() {"Particle", "Translation", "Grammar Feature"}
        For Count = 0 To CachedData.IslamData.GrammarCategories(Index).Words.Length - 1
            Output(3 + Count) = New String() {TransliterateFromBuckwalter(CachedData.IslamData.GrammarCategories(Index).Words(Count).Text), CachedData.IslamData.GrammarCategories(Index).Words(Count).TranslationID, Utility.DefaultValue(CachedData.IslamData.GrammarCategories(Index).Words(Count).Grammar, String.Empty)}
        Next
        Return Output
    End Function
    Public Shared Function DisplayPronoun(Index As Integer, Personal As Boolean) As Array()
        Dim Count As Integer
        Dim Output(2 + If(Personal, 6, 2)) As Array
        Dim Build As New Generic.Dictionary(Of String, Generic.Dictionary(Of String, String))
        Output(0) = New String() {}
        Output(1) = New String() {"arabic", "arabic", "arabic", String.Empty}
        Output(2) = New String() {"Plural", "Dual", "Singular", String.Empty}
        For Count = 0 To CachedData.IslamData.GrammarCategories(Index).Words.Length - 1
            Array.ForEach(CachedData.IslamData.GrammarCategories(Index).Words(Count).Grammar.Split(","c)(0).Split("|"c),
                          Sub(Str As String)
                              Dim Key As String = Str.Chars(0)
                              If Personal Then '"123".Contains(Str.Chars(0))
                                  Key += Str.Chars(1)
                              End If
                              If Not Build.ContainsKey(Key) Then
                                  Build.Add(Key, New Generic.Dictionary(Of String, String))
                              End If
                              If Build.Item(Key).ContainsKey(Str.Chars(If(Personal, 2, 1))) Then
                                  Build.Item(Key).Item(Str.Chars(If(Personal, 2, 1))) += " " + TransliterateFromBuckwalter(CachedData.IslamData.GrammarCategories(Index).Words(Count).Text)
                              Else
                                  Build.Item(Key).Add(Str.Chars(If(Personal, 2, 1)), TransliterateFromBuckwalter(CachedData.IslamData.GrammarCategories(Index).Words(Count).Text))
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
        Return DisplayPronoun(1, False)
    End Function
    Public Shared Function DisplayDistals(ByVal Item As PageLoader.TextItem) As Array()
        Return DisplayPronoun(2, False)
    End Function
    Public Shared Function DisplayRelatives(ByVal Item As PageLoader.TextItem) As Array()
        Return DisplayPronoun(4, False)
    End Function
    Public Shared Function DisplayPersonals(ByVal Item As PageLoader.TextItem) As Array()
        Return DisplayPronoun(5, True)
    End Function
    Public Shared Function DisplayDeterminerPersonals(ByVal Item As PageLoader.TextItem) As Array()
        Return DisplayPronoun(6, True)
    End Function
    Public Shared Function DisplayPastVerbsFamilyI(ByVal Item As PageLoader.TextItem) As Array()
        Return DisplayPronoun(7, True)
    End Function
    Public Shared Function DisplayPresentVerbsFamilyI(ByVal Item As PageLoader.TextItem) As Array()
        Return DisplayPronoun(8, True)
    End Function
    Public Shared Function DisplayCommandVerbsFamilyI(ByVal Item As PageLoader.TextItem) As Array()
        Return DisplayPronoun(9, False)
    End Function
    Public Shared Function DisplayResponseParticles(ByVal Item As PageLoader.TextItem) As Array()
        Return DisplayParticle(10)
    End Function
    Public Shared Function DisplayInterogativeParticles(ByVal Item As PageLoader.TextItem) As Array()
        Return DisplayParticle(11)
    End Function
    Public Shared Function DisplayLocationParticles(ByVal Item As PageLoader.TextItem) As Array()
        Return DisplayParticle(12)
    End Function
    Public Shared Function DisplayTimeParticles(ByVal Item As PageLoader.TextItem) As Array()
        Return DisplayParticle(13)
    End Function
    Public Shared Function DisplayPrepositions(ByVal Item As PageLoader.TextItem) As Array()
        Return DisplayParticle(14)
    End Function
    Public Shared Function DisplayParticles(ByVal Item As PageLoader.TextItem) As Array()
        Return DisplayParticle(15)
    End Function
    Public Shared Function DisplayOtherParticles(ByVal Item As PageLoader.TextItem) As Array()
        Return DisplayParticle(16)
    End Function
    Public Shared Function DisplayNouns(ByVal Item As PageLoader.TextItem) As Array()
        Dim Count As Integer
        Dim Index As Integer = 17
        Dim Output(2 + CachedData.IslamData.GrammarCategories(Index).Words.Length) As Array
        Dim Build As New Generic.Dictionary(Of String, Generic.Dictionary(Of String, String))
        Output(0) = New String() {}
        Output(1) = New String() {"arabic", "arabic", "arabic", String.Empty}
        Output(2) = New String() {"Noun", "Dual", "Plural", "Singular Translation"}
        For Count = 0 To CachedData.IslamData.GrammarCategories(Index).Words.Length - 1
            Output(3 + Count) = New String() {TransliterateFromBuckwalter(CachedData.IslamData.GrammarCategories(Index).Words(Count).Text), String.Empty, String.Empty, CachedData.IslamData.GrammarCategories(Index).Words(Count).TranslationID}
        Next
        Return Output
    End Function
    Public Shared Function DisplayVerbs(ByVal Item As PageLoader.TextItem) As Array()
        Dim Count As Integer
        Dim Index As Integer = 18
        Dim Output(2 + CachedData.IslamData.GrammarCategories(Index).Words.Length) As Array
        Dim Build As New Generic.Dictionary(Of String, Generic.Dictionary(Of String, String))
        Output(0) = New String() {}
        Output(1) = New String() {"arabic", "arabic", "arabic", "arabic", "arabic", "arabic", "arabic", "arabic", "arabic"}
        Output(2) = New String() {"Past Root", "Present Root", "Command Root", "Forbidding Root", "Passive Past Root", "Passive Present Root", "Verbal Doer", "Passive Noun", "Particles"}
        For Count = 0 To CachedData.IslamData.GrammarCategories(Index).Words.Length - 1
            Dim Grammar As String
            Dim Text As String
            Dim Present As String
            Dim Command As String
            Dim Forbidding As String
            Dim PassivePast As String
            Dim PassivePresent As String
            Dim VerbalDoer As String
            Dim PassiveNoun As String
            If (Not CachedData.IslamData.GrammarCategories(Index).Words(Count).Grammar Is Nothing AndAlso CachedData.IslamData.GrammarCategories(Index).Words(Count).Grammar.StartsWith("form=")) Then
                Text = TransliterateFromBuckwalter(CachedData.IslamData.GrammarCategories(Index).Words(Count).Grammar.Substring(5).Split(","c)(0).Replace("f", CachedData.IslamData.GrammarCategories(Index).Words(Count).Text.Chars(0)).Replace("E", CachedData.IslamData.GrammarCategories(Index).Words(Count).Text.Chars(1)).Replace("l", CachedData.IslamData.GrammarCategories(Index).Words(Count).Text.Chars(2)))
                Present = TransliterateFromBuckwalter(CachedData.IslamData.GrammarCategories(Index).Words(Count).Grammar.Substring(5).Split(","c)(1).Replace("f", CachedData.IslamData.GrammarCategories(Index).Words(Count).Text.Chars(0)).Replace("E", CachedData.IslamData.GrammarCategories(Index).Words(Count).Text.Chars(1)).Replace("l", CachedData.IslamData.GrammarCategories(Index).Words(Count).Text.Chars(2)))
                Command = TransliterateFromBuckwalter("{foE$1lo".Replace("f", CachedData.IslamData.GrammarCategories(Index).Words(Count).Text.Chars(0)).Replace("E", CachedData.IslamData.GrammarCategories(Index).Words(Count).Text.Chars(1)).Replace("l", CachedData.IslamData.GrammarCategories(Index).Words(Count).Text.Chars(2)).Replace("$1", CachedData.IslamData.GrammarCategories(Index).Words(Count).Grammar.Substring(5).Split(","c)(1).Chars(5)))
                Forbidding = TransliterateFromBuckwalter("laA " + "tafoE$1lo".Replace("f", CachedData.IslamData.GrammarCategories(Index).Words(Count).Text.Chars(0)).Replace("E", CachedData.IslamData.GrammarCategories(Index).Words(Count).Text.Chars(1)).Replace("l", CachedData.IslamData.GrammarCategories(Index).Words(Count).Text.Chars(2)).Replace("$1", CachedData.IslamData.GrammarCategories(Index).Words(Count).Grammar.Substring(5).Split(","c)(1).Chars(5)))
                PassivePast = TransliterateFromBuckwalter("fuEila".Replace("f", CachedData.IslamData.GrammarCategories(Index).Words(Count).Text.Chars(0)).Replace("E", CachedData.IslamData.GrammarCategories(Index).Words(Count).Text.Chars(1)).Replace("l", CachedData.IslamData.GrammarCategories(Index).Words(Count).Text.Chars(2)))
                PassivePresent = TransliterateFromBuckwalter("yufoEalu".Replace("f", CachedData.IslamData.GrammarCategories(Index).Words(Count).Text.Chars(0)).Replace("E", CachedData.IslamData.GrammarCategories(Index).Words(Count).Text.Chars(1)).Replace("l", CachedData.IslamData.GrammarCategories(Index).Words(Count).Text.Chars(2)))
                VerbalDoer = TransliterateFromBuckwalter("faAEilN".Replace("f", CachedData.IslamData.GrammarCategories(Index).Words(Count).Text.Chars(0)).Replace("E", CachedData.IslamData.GrammarCategories(Index).Words(Count).Text.Chars(1)).Replace("l", CachedData.IslamData.GrammarCategories(Index).Words(Count).Text.Chars(2)))
                PassiveNoun = TransliterateFromBuckwalter("mafoEuwlN".Replace("f", CachedData.IslamData.GrammarCategories(Index).Words(Count).Text.Chars(0)).Replace("E", CachedData.IslamData.GrammarCategories(Index).Words(Count).Text.Chars(1)).Replace("l", CachedData.IslamData.GrammarCategories(Index).Words(Count).Text.Chars(2)))
                '"faEiylN" passive noun
                '"mafoEilN" time
                '"mafoEalN" place
                '"faEiylN" derived adjective "faEalN" "faEulaAnu" "&gt;ufoEalN" "faEolN" "fuEaAlN" "faEuwlN"
                '"&gt;afoEalu" "fuEolaY" comparative derived noun
                '"faE~aAlN" "mifoEaAlN" "faEuwlN" "fuE~uwl" "faEilN" "fiE~iylN" "fuEalapN" intensive derived noun
                '"mifoEalapN" "mifoEalN" "mifoEaAlN" instrument of action
                Grammar = String.Empty
            Else
                Text = TransliterateFromBuckwalter(CachedData.IslamData.GrammarCategories(Index).Words(Count).Text)
                Present = String.Empty
                Command = String.Empty
                Forbidding = String.Empty
                PassivePast = String.Empty
                PassivePresent = String.Empty
                VerbalDoer = String.Empty
                PassiveNoun = String.Empty
                Grammar = Utility.DefaultValue(CachedData.IslamData.GrammarCategories(Index).Words(Count).Grammar, String.Empty)
            End If
            Output(3 + Count) = New String() {Text, Present, Command, Forbidding, PassivePast, PassivePresent, VerbalDoer, PassiveNoun, Grammar}
        Next
        Return Output
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
                    ElseIf Items(Count).TextItems(Index).DisplayClass = RenderDisplayClass.eTransliteration Then
                        writer.WriteAttribute("class", "transliteration")
                        writer.WriteAttribute("dir", "ltr")
                        writer.WriteAttribute("style", "display: " + CStr(IIf(CInt(HttpContext.Current.Request.QueryString.Get("translitscheme")) <> 0, "block", "none")) + ";")
                        writer.WriteAttribute("id", "translit" + CStr(IIf(NestPrefix = String.Empty, String.Empty, NestPrefix + "_")) + CStr(Count) + "_" + CStr(Index))
                    Else
                        writer.WriteAttribute("class", "translation")
                        writer.WriteAttribute("dir", CStr(IIf(Items(Count).TextItems(Index).DisplayClass = RenderDisplayClass.eRTL, "rtl", "ltr")))
                        writer.WriteAttribute("id", "translate" + CStr(IIf(NestPrefix = String.Empty, String.Empty, NestPrefix + "_")) + CStr(Count) + "_" + CStr(Index))
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
        Return "function fontWidth(fontName, text) { text = text || '" + TanzilReader.GetTextVerse(TanzilReader.GetTextChapter(CachedData.XMLDocMain, 3), 9).Attributes.GetNamedItem("text").Value + "' ; if (text == 2) text = '" + Utility.LoadResourceString("IslamInfo_InTheNameOfAllah") + "," + TanzilReader.GetTextVerse(TanzilReader.GetTextChapter(CachedData.XMLDocMain, 1), 1).Attributes.GetNamedItem("text").Value + "'; var tester = $('#font-tester'); tester.css('fontFamily', fontName); if (tester.firstChild) tester.remove(tester.firstChild); tester.append(document.createTextNode(text)); tester.css('display', 'block'); var width = tester.offsetWidth; tester.css('display', 'none'); return width; }"
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
        Property ShapingParse As String
            Get
                If Shaping.Length = 0 Then Return String.Empty
                Return String.Join(","c, Array.ConvertAll(Shaping, Function(Ch As Char) Asc(Ch).ToString("X2")))
            End Get
            Set(value As String)
                If Not value Is Nothing Then
                    Shaping = Array.ConvertAll(value.Split(","c), Function(Str As String) ChrW(Integer.Parse(Str, System.Globalization.NumberStyles.HexNumber)))
                End If
            End Set
        End Property
        <System.Xml.Serialization.XmlAttribute("ipavalue")> _
        Public IPAValue As String
        <System.Xml.Serialization.XmlAttribute("extendedbuckwalter")> _
        Public ExtendedBuckwalterLetter As Char
        <System.Xml.Serialization.XmlAttribute("romantranslit")> _
        Public RomanTranslit As String
        <System.Xml.Serialization.XmlAttribute("plainroman")> _
        Public PlainRoman As String
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
End Class
Public Class CachedData
    'need disk and memory cache as time consuming to read or build
    Shared _ObjIslamData As IslamData
    Shared _XMLDocMain As System.Xml.XmlDocument 'Tanzil Quran data
    Shared _XMLDocInfo As System.Xml.XmlDocument 'Tanzil metadata
    Shared _XMLDocInfos As Collections.Generic.List(Of System.Xml.XmlDocument) 'Hadiths
    Shared _RootDictionary As New Generic.Dictionary(Of String, ArrayList)
    Shared _TagDictionary As New Generic.Dictionary(Of String, ArrayList)
    Shared _WordDictionary As New Generic.Dictionary(Of String, ArrayList)
    Shared _LetterDictionary As New Generic.Dictionary(Of Char, ArrayList)
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
                    'TAG
                    If Not _TagDictionary.ContainsKey(Pieces(2)) Then
                        _TagDictionary.Add(Pieces(2), New ArrayList)
                    End If
                    Dim Location As Integer() = Array.ConvertAll(Pieces(0).TrimStart("("c).TrimEnd(")"c).Split(":"c), Function(Str As String) CInt(Str))
                    _TagDictionary.Item(Pieces(2)).Add(Location)
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
                        _LetterDictionary.Add(Verses(Count)(SubCount)(LetCount), New ArrayList)
                    End If
                    _LetterDictionary.Item(Verses(Count)(SubCount)(LetCount)).Add(New Integer() {Count, SubCount, LetCount})
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
    Public Shared ReadOnly Property IslamData As IslamData
        Get
            If _ObjIslamData Is Nothing Then
                Dim fs As IO.FileStream = New IO.FileStream(Utility.GetFilePath("metadata\islaminfo.xml"), IO.FileMode.Open)
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
    Public Shared ReadOnly Property TagDictionary As Generic.Dictionary(Of String, ArrayList)
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
    Public Shared ReadOnly Property LetterDictionary As Generic.Dictionary(Of Char, ArrayList)
        Get
            If _LetterDictionary.Keys.Count = 0 Then BuildQuranLetterIndex()
            Return _LetterDictionary
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
Public Class Supplications
    Public Shared Function GetSuppCategories() As String()
        Return Array.ConvertAll(CachedData.IslamData.VerseCategories, Function(Convert As IslamData.VerseCategory) Utility.LoadResourceString("IslamInfo_" + Convert.Title))
    End Function
    Public Shared Function GetRenderedSuppText(ByVal Item As PageLoader.TextItem) As RenderArray
        Dim Renderer As New RenderArray
        Dim Count As Integer = CInt(Web.HttpContext.Current.Request.QueryString.Get("selection"))
        If Count = -1 Then Count = 0
        For SubCount As Integer = 0 To CachedData.IslamData.VerseCategories(Count).Verses.Length - 1
            Dim EnglishByWord As String() = Utility.LoadResourceString("IslamInfo_" + CachedData.IslamData.VerseCategories(Count).Verses(SubCount).TranslationID + "WordByWord").Split("|"c)
            Dim ArabicText As String() = CachedData.IslamData.VerseCategories(Count).Verses(SubCount).Arabic.Split(" "c)
            Dim Transliteration As String() = Arabic.TransliterateToScheme(Arabic.TransliterateFromBuckwalter(CachedData.IslamData.VerseCategories(Count).Verses(SubCount).Arabic)).Split(" "c)
            Renderer.Items.Add(New RenderArray.RenderItem(RenderArray.RenderTypes.eHeaderCenter, New RenderArray.RenderText() {New RenderArray.RenderText(RenderArray.RenderDisplayClass.eLTR, Utility.LoadResourceString("IslamInfo_" + CachedData.IslamData.VerseCategories(Count).Verses(SubCount).TranslationID))}))
            Dim Items As New Collections.Generic.List(Of RenderArray.RenderItem)
            For WordCount As Integer = 0 To EnglishByWord.Length - 1
                Items.Add(New RenderArray.RenderItem(RenderArray.RenderTypes.eText, New RenderArray.RenderText() {New RenderArray.RenderText(RenderArray.RenderDisplayClass.eArabic, Arabic.RightToLeftMark + Arabic.TransliterateFromBuckwalter(ArabicText(WordCount))), New RenderArray.RenderText(RenderArray.RenderDisplayClass.eTransliteration, Transliteration(WordCount)), New RenderArray.RenderText(RenderArray.RenderDisplayClass.eLTR, EnglishByWord(WordCount))}))
            Next
            Renderer.Items.Add(New RenderArray.RenderItem(RenderArray.RenderTypes.eText, New RenderArray.RenderText() {New RenderArray.RenderText(RenderArray.RenderDisplayClass.eNested, Items), New RenderArray.RenderText(RenderArray.RenderDisplayClass.eLTR, Utility.LoadResourceString("IslamInfo_" + CachedData.IslamData.VerseCategories(Count).Verses(SubCount).TranslationID + "Trans"))}))
        Next
        Return Renderer
    End Function
End Class
Public Class Quiz
    Public Shared Function GetQuizList() As Array()
        Dim Strings(1) As Array
        Strings(0) = New String() {"ArabicLetters", "arabicletters"}
        Strings(1) = New String() {"ArabicLettersDiacritics", "arabiclettersdiacritics"}
        Return Strings
    End Function
    Public Shared Function GetChangeQuizJS() As String()
        Return New String() {"javascript: changeQuiz();", String.Empty, _
                             "function changeQuiz() { qtype = $('#quizselection').val(); qwrong = 0; qright = 0; nextQuestion(); }"}
    End Function
    Public Shared Function DisplayCount(ByVal Item As PageLoader.TextItem) As String
        Return "Wrong: 0 Right: 0"
    End Function
    Public Shared Function DisplayQuestion(ByVal Item As PageLoader.TextItem) As String
        Web.HttpContext.Current.Items.Add("rnd", DateTime.Now.ToFileTime())
        Rnd(-1)
        Randomize(CDbl(Web.HttpContext.Current.Items("rnd")))
        Dim Count As Integer = CInt(Math.Floor(Rnd() * 4))
        While Count
            Rnd()
            Count -= 1
        End While
        Return CachedData.IslamData.ArabicLetters(Arabic.FindLetterBySymbol(Arabic.ArabicLetters(CInt(Math.Floor(Rnd() * Arabic.ArabicLetters.Length))).Chars(0))).Symbol
    End Function
    Public Shared Function DisplayAnswer(ByVal Item As PageLoader.ButtonItem) As String
        Dim Count As Integer
        Rnd(-1)
        Randomize(CDbl(Web.HttpContext.Current.Items("rnd")))
        Rnd()
        For Count = 2 To Integer.Parse(Item.Name.Replace("answer", String.Empty))
            Rnd()
        Next
        Return CachedData.IslamData.ArabicLetters(Arabic.FindLetterBySymbol(Arabic.ArabicLetters(CInt(Math.Floor(Rnd() * Arabic.ArabicLetters.Length))).Chars(0))).RomanTranslit
    End Function
    Public Shared Function VerifyAnswer() As String()
        Return New String() {"javascript: verifyAnswer(this);", String.Empty, _
                             Arabic.GetArabicSymbolJSArray(), Arabic.FindLetterBySymbolJS(), _
                             "var ArabicLets = " + Utility.MakeJSArray(Array.ConvertAll(Arabic.ArabicLetters, Function(Str As String) CStr(AscW(Str.Chars(0)))), True) + ";", _
                             "var qtype = 'arabicletters', qwrong = 0, qright = 0; " + _
                             "function getUniqueRnd(excl, count) { var rnd; do { rnd = Math.floor(Math.random() * count); } while (excl.indexOf(rnd) !== -1); return rnd; } " + _
                             "function nextQuestion() { $('#count').text('Wrong: ' + qwrong + ' Right: ' + qright); var i = Math.floor(Math.random() * 4), nidx = getUniqueRnd([], ArabicLets.length), aidx = []; aidx[0] = getUniqueRnd([nidx], ArabicLets.length); aidx[1] = getUniqueRnd([nidx, aidx[0]], ArabicLets.length); aidx[2] = getUniqueRnd([nidx, aidx[0], aidx[1]], ArabicLets.length); $('#question').text(String.fromCharCode(ArabicLets[nidx])); $('#answer1').prop('value', ArabicLetters[findLetterBySymbol(ArabicLets[i === 0 ? nidx : aidx[0]])].RomanTranslit); $('#answer2').prop('value', ArabicLetters[findLetterBySymbol(ArabicLets[i === 1 ? nidx : aidx[i > 1 ? 1 : 0]])].RomanTranslit); $('#answer3').prop('value', ArabicLetters[findLetterBySymbol(ArabicLets[i === 2 ? nidx : aidx[i > 2 ? 2 : 1]])].RomanTranslit); $('#answer4').prop('value', ArabicLetters[findLetterBySymbol(ArabicLets[i === 3 ? nidx : aidx[2]])].RomanTranslit); } " + _
                             "function verifyAnswer(ctl) { $(ctl).prop('value') === ArabicLetters[findLetterBySymbol($('#question').text().charCodeAt(0))].RomanTranslit ? qright++ : qwrong++; nextQuestion(); }"}
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
        Return New String() {Utility.LoadResourceString("IslamInfo_Letters"), Utility.LoadResourceString("IslamInfo_Words"), Utility.LoadResourceString("IslamInfo_UniqueWords"), Utility.LoadResourceString("IslamInfo_UniqueWordsPerPart"), Utility.LoadResourceString("IslamInfo_WordsPerPart"), Utility.LoadResourceString("IslamInfo_UniqueWordsPerStation"), Utility.LoadResourceString("IslamInfo_WordsPerStation"), Utility.LoadResourceString("IslamInfo_Letters")}
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
        Strings = Web.HttpContext.Current.Request.QueryString.Get("quranselection")
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
        Strings = Web.HttpContext.Current.Request.QueryString.Get("quranselection")
        If Not Strings Is Nothing Then Index = CInt(Strings)
        If Index = 0 Then
            All = CachedData.TotalLetters
            Dim LetterFreqArray(CachedData.LetterDictionary.Keys.Count - 1) As Char
            CachedData.LetterDictionary.Keys.CopyTo(LetterFreqArray, 0)
            Array.Sort(LetterFreqArray, Function(Key As Char, NextKey As Char) CachedData.LetterDictionary.Item(NextKey).Count.CompareTo(CachedData.LetterDictionary.Item(Key).Count))
            For Count As Integer = 0 To LetterFreqArray.Length - 1
                Total += CachedData.LetterDictionary.Item(LetterFreqArray(Count)).Count
                Output.Add(New String() {CachedData.IslamData.ArabicLetters(Arabic.FindLetterBySymbol(LetterFreqArray(Count))).UnicodeName + " (" + Arabic.RightToLeftMark + " " + LetterFreqArray(Count) + " " + Arabic.LeftToRightMark + ")", String.Empty, CStr(CachedData.LetterDictionary.Item(LetterFreqArray(Count)).Count), (CDbl(CachedData.LetterDictionary.Item(LetterFreqArray(Count)).Count) * 100 / All).ToString("n2"), (CDbl(Total) * 100 / All).ToString("n2")})
            Next
        ElseIf Index = 7 Then
            All = CachedData.TotalIsolatedLetters
            Dim LetterFreqArray(CachedData.IsolatedLetterDictionary.Keys.Count - 1) As Char
            CachedData.IsolatedLetterDictionary.Keys.CopyTo(LetterFreqArray, 0)
            Array.Sort(LetterFreqArray, Function(Key As Char, NextKey As Char) CachedData.IsolatedLetterDictionary.Item(NextKey).Count.CompareTo(CachedData.IsolatedLetterDictionary.Item(Key).Count))
            For Count As Integer = 0 To LetterFreqArray.Length - 1
                Total += CachedData.IsolatedLetterDictionary.Item(LetterFreqArray(Count)).Count
                Output.Add(New String() {CachedData.IslamData.ArabicLetters(Arabic.FindLetterBySymbol(LetterFreqArray(Count))).UnicodeName + " (" + Arabic.RightToLeftMark + " " + LetterFreqArray(Count) + " " + Arabic.LeftToRightMark + ")", String.Empty, CStr(CachedData.IsolatedLetterDictionary.Item(LetterFreqArray(Count)).Count), (CDbl(CachedData.IsolatedLetterDictionary.Item(LetterFreqArray(Count)).Count) * 100 / All).ToString("n2"), (CDbl(Total) * 100 / All).ToString("n2")})
            Next
        ElseIf Index = 1 Then
            Dim FreqArray(CachedData.WordDictionary.Keys.Count - 1) As String
            CachedData.WordDictionary.Keys.CopyTo(FreqArray, 0)
            Total = 0
            All = GetQuranWordTotalNumber()
            Array.Sort(FreqArray, Function(Key As String, NextKey As String) CachedData.WordDictionary.Item(NextKey).Count.CompareTo(CachedData.WordDictionary.Item(Key).Count))
            For Count As Integer = 0 To FreqArray.Length - 1
                Total += CachedData.WordDictionary.Item(FreqArray(Count)).Count
                Output.Add(New String() {Arabic.RightToLeftMark + Arabic.TransliterateFromBuckwalter(FreqArray(Count)), String.Empty, CStr(CachedData.WordDictionary.Item(FreqArray(Count)).Count), (CDbl(CachedData.WordDictionary.Item(FreqArray(Count)).Count) * 100 / All).ToString("n2"), (CDbl(Total) * 100 / All).ToString("n2")})
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
        End If
        Return CType(Output.ToArray(GetType(Array())), Array())
    End Function
    Public Shared Function GetSelectionNames() As Array()
        Dim Division As Integer = 0
        Dim Strings As String = Web.HttpContext.Current.Request.QueryString.Get("qurandivision")
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
    Public Shared Function GetRenderedQuranText(ByVal Item As PageLoader.TextItem) As RenderArray
        Dim Division As Integer = 0
        Dim Index As Integer = 1
        Dim Chapter As Integer
        Dim Verse As Integer
        Dim BaseChapter As Integer
        Dim BaseVerse As Integer
        Dim Text As String
        Dim Node As System.Xml.XmlNode
        Dim Renderer As New RenderArray
        Dim Strings As String = Web.HttpContext.Current.Request.QueryString.Get("qurandivision")
        If Not Strings Is Nothing Then Division = CInt(Strings)
        Strings = Web.HttpContext.Current.Request.QueryString.Get("quranselection")
        If Not Strings Is Nothing Then Index = CInt(Strings)
        Dim QuranText As Collections.Generic.List(Of String())
        Dim SeperateSectionCount As Integer = 1
        If Division = 8 Then SeperateSectionCount = CachedData.IslamData.QuranSelections(Index).SelectionInfo.Length
        If Division = 9 Then SeperateSectionCount = CachedData.LetterDictionary(CachedData.IslamData.ArabicLetters(Index).Symbol).Count
        Dim Lines As String() = IO.File.ReadAllLines(Utility.GetFilePath("metadata\" + GetTranslationFileName(Web.HttpContext.Current.Request.QueryString.Get("qurantranslation"))))
        Dim W4WLines As String() = IO.File.ReadAllLines(Utility.GetFilePath("metadata\en.w4w.shehnazshaikh.txt"))
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
                QuranText = New Collections.Generic.List(Of String())
                QuranText.Add(GetQuranText(CachedData.XMLDocMain, BaseChapter, BaseVerse, CInt(IIf(CachedData.IslamData.QuranSelections(Index).SelectionInfo(SectionCount).ExtraVerseNumber <> 0, CachedData.IslamData.QuranSelections(Index).SelectionInfo(SectionCount).ExtraVerseNumber, BaseVerse))))
                If (CachedData.IslamData.QuranSelections(Index).SelectionInfo(SectionCount).WordNumber > 1) Then
                    Dim VerseIndex As Integer = 0
                    For Count As Integer = 1 To CachedData.IslamData.QuranSelections(Index).SelectionInfo(SectionCount).WordNumber - 1
                        VerseIndex = QuranText(0)(0).IndexOf(" "c, VerseIndex) + 1
                    Next
                    QuranText(0)(0) = QuranText(0)(0).Substring(VerseIndex)
                End If
                If (CachedData.IslamData.QuranSelections(Index).SelectionInfo(SectionCount).EndWordNumber <> 0) Then
                    Dim VerseIndex As Integer = 0
                    'selections are always within the same chapter
                    Dim LastVerse As Integer = CInt(IIf(CachedData.IslamData.QuranSelections(Index).SelectionInfo(SectionCount).ExtraVerseNumber <> 0, 1, 0))
                    For Count As Integer = CachedData.IslamData.QuranSelections(Index).SelectionInfo(SectionCount).WordNumber To CachedData.IslamData.QuranSelections(Index).SelectionInfo(SectionCount).EndWordNumber - 1
                        VerseIndex = QuranText(0)(LastVerse).IndexOf(" "c, VerseIndex) + 1
                    Next
                    QuranText(0)(LastVerse) = QuranText(0)(LastVerse).Substring(0, VerseIndex + 1)
                End If
            ElseIf Division = 9 Then
                QuranText = New Collections.Generic.List(Of String())
                QuranText.Add(GetQuranText(CachedData.XMLDocMain, CType(CachedData.LetterDictionary(CachedData.IslamData.ArabicLetters(Index).Symbol)(SectionCount), Integer())(0), CType(CachedData.LetterDictionary(CachedData.IslamData.ArabicLetters(Index).Symbol)(SectionCount), Integer())(1), CType(CachedData.LetterDictionary(CachedData.IslamData.ArabicLetters(Index).Symbol)(SectionCount), Integer())(1)))
            Else
                QuranText = Nothing
            End If
            If Not QuranText Is Nothing Then
                For Chapter = 0 To QuranText.Count - 1
                    Dim ChapterNode As System.Xml.XmlNode = GetChapterByIndex(BaseChapter + Chapter)
                    Renderer.Items.Add(New RenderArray.RenderItem(RenderArray.RenderTypes.eHeaderLeft, New RenderArray.RenderText() {New RenderArray.RenderText(RenderArray.RenderDisplayClass.eArabic, Arabic.RightToLeftMark + Arabic.TransliterateFromBuckwalter("|yAthA " + ChapterNode.Attributes.GetNamedItem("ayas").Value + " ")), New RenderArray.RenderText(RenderArray.RenderDisplayClass.eTransliteration, Arabic.TransliterateToScheme(Arabic.TransliterateFromBuckwalter("|yAthA " + ChapterNode.Attributes.GetNamedItem("ayas").Value + " ")).Trim()), New RenderArray.RenderText(RenderArray.RenderDisplayClass.eLTR, "Verses " + ChapterNode.Attributes.GetNamedItem("ayas").Value + " ")}))
                    Renderer.Items.Add(New RenderArray.RenderItem(RenderArray.RenderTypes.eHeaderCenter, New RenderArray.RenderText() {New RenderArray.RenderText(RenderArray.RenderDisplayClass.eArabic, Arabic.RightToLeftMark + Arabic.TransliterateFromBuckwalter("suwrapu " + CachedData.IslamData.QuranChapters(CInt(ChapterNode.Attributes.GetNamedItem("index").Value) - 1).Name + " ")), New RenderArray.RenderText(RenderArray.RenderDisplayClass.eTransliteration, Arabic.TransliterateToScheme(Arabic.TransliterateFromBuckwalter("suwrapu " + CachedData.IslamData.QuranChapters(CInt(ChapterNode.Attributes.GetNamedItem("index").Value) - 1).Name + " ")).Trim()), New RenderArray.RenderText(RenderArray.RenderDisplayClass.eLTR, "Chapter " + TanzilReader.GetChapterEName(ChapterNode) + " ")}))
                    Renderer.Items.Add(New RenderArray.RenderItem(RenderArray.RenderTypes.eHeaderRight, New RenderArray.RenderText() {New RenderArray.RenderText(RenderArray.RenderDisplayClass.eArabic, Arabic.RightToLeftMark + Arabic.TransliterateFromBuckwalter("rukwEAthA " + ChapterNode.Attributes.GetNamedItem("rukus").Value + " ")), New RenderArray.RenderText(RenderArray.RenderDisplayClass.eTransliteration, Arabic.TransliterateToScheme(Arabic.TransliterateFromBuckwalter("rukwEAthA " + ChapterNode.Attributes.GetNamedItem("rukus").Value + " ")).Trim()), New RenderArray.RenderText(RenderArray.RenderDisplayClass.eLTR, "Rukus " + ChapterNode.Attributes.GetNamedItem("rukus").Value + " ")}))
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
                                Renderer.Items.Add(New RenderArray.RenderItem(RenderArray.RenderTypes.eText, New RenderArray.RenderText() {New RenderArray.RenderText(RenderArray.RenderDisplayClass.eArabic, Arabic.RightToLeftMark + Node.Value + " "), New RenderArray.RenderText(RenderArray.RenderDisplayClass.eTransliteration, Arabic.TransliterateToScheme(Node.Value).Trim()), New RenderArray.RenderText(DirectCast(IIf(IsTranslationTextLTR(), RenderArray.RenderDisplayClass.eLTR, RenderArray.RenderDisplayClass.eRTL), RenderArray.RenderDisplayClass), TanzilReader.GetTranslationVerse(Lines, 1, 1))}))
                            End If
                        End If
                        Dim Words As String() = QuranText(Chapter)(Verse).Split(" "c)
                        Dim TranslitWords As String() = Arabic.TransliterateToScheme(QuranText(Chapter)(Verse)).Split(" "c)
                        Dim PauseMarks As Integer = 0
                        For Count As Integer = 0 To Words.Length - 1
                            If Words(Count).Length = 1 AndAlso _
                                Arabic.IsStop(Arabic.FindLetterBySymbol(Words(Count)(0))) Then
                                PauseMarks += 1
                                Items.Add(New RenderArray.RenderItem(RenderArray.RenderTypes.eText, New RenderArray.RenderText() {New RenderArray.RenderText(RenderArray.RenderDisplayClass.eArabic, Arabic.RightToLeftMark + Words(Count)), New RenderArray.RenderText(RenderArray.RenderDisplayClass.eTransliteration, TranslitWords(Count))}))
                            Else
                                Items.Add(New RenderArray.RenderItem(RenderArray.RenderTypes.eText, New RenderArray.RenderText() {New RenderArray.RenderText(RenderArray.RenderDisplayClass.eArabic, Arabic.RightToLeftMark + Words(Count)), New RenderArray.RenderText(RenderArray.RenderDisplayClass.eTransliteration, TranslitWords(Count)), New RenderArray.RenderText(DirectCast(IIf(IsTranslationTextLTR(), RenderArray.RenderDisplayClass.eLTR, RenderArray.RenderDisplayClass.eRTL), RenderArray.RenderDisplayClass), TanzilReader.GetW4WTranslationVerse(W4WLines, BaseChapter + Chapter, CInt(IIf(Chapter = 0, BaseVerse, 1)) + Verse, Count - PauseMarks))}))
                            End If
                        Next
                        Text += QuranText(Chapter)(Verse) + " "
                        If TanzilReader.IsSajda(BaseChapter + Chapter, CInt(IIf(Chapter = 0, BaseVerse, 1)) + Verse) Then
                            'Sajda markers are already in the text
                            'Text += Arabic.TransliterateFromBuckwalter("R")
                            'Items.Add(New RenderArray.RenderItem(RenderArray.RenderTypes.eText, New RenderArray.RenderText() {New RenderArray.RenderText(RenderArray.RenderDisplayClass.eArabic, Arabic.RightToLeftMark + Arabic.TransliterateFromBuckwalter("R"))}))
                        End If
                        Text += Arabic.TransliterateFromBuckwalter("=" + CStr(CInt(IIf(Chapter = 0, BaseVerse, 1)) + Verse)) + " "
                        Items.Add(New RenderArray.RenderItem(RenderArray.RenderTypes.eText, New RenderArray.RenderText() {New RenderArray.RenderText(RenderArray.RenderDisplayClass.eArabic, Arabic.RightToLeftMark + Arabic.TransliterateFromBuckwalter("=" + CStr(CInt(IIf(Chapter = 0, BaseVerse, 1)) + Verse))), New RenderArray.RenderText(DirectCast(IIf(IsTranslationTextLTR(), RenderArray.RenderDisplayClass.eLTR, RenderArray.RenderDisplayClass.eRTL), RenderArray.RenderDisplayClass), "(" + CStr(CInt(IIf(Chapter = 0, BaseVerse, 1)) + Verse) + ")")}))
                        'Text += Arabic.TransliterateFromBuckwalter("(" + CStr(IIf(Chapter = 0, BaseVerse, 1) + Verse) + ") ")
                        Renderer.Items.Add(New RenderArray.RenderItem(RenderArray.RenderTypes.eText, New RenderArray.RenderText() {New RenderArray.RenderText(RenderArray.RenderDisplayClass.eNested, Items), New RenderArray.RenderText(RenderArray.RenderDisplayClass.eArabic, Arabic.RightToLeftMark + Text), New RenderArray.RenderText(RenderArray.RenderDisplayClass.eTransliteration, Arabic.TransliterateToScheme(QuranText(Chapter)(Verse)).Trim()), New RenderArray.RenderText(DirectCast(IIf(IsTranslationTextLTR(), RenderArray.RenderDisplayClass.eLTR, RenderArray.RenderDisplayClass.eRTL), RenderArray.RenderDisplayClass), "(" + CStr(CInt(IIf(Chapter = 0, BaseVerse, 1)) + Verse) + ") " + TanzilReader.GetTranslationVerse(Lines, BaseChapter + Chapter, CInt(IIf(Chapter = 0, BaseVerse, 1)) + Verse))}))
                    Next
                Next
            End If
        Next
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
    Public Shared Function IsTranslationTextLTR() As Boolean
        Dim Index As Integer = GetTranslationIndex(Web.HttpContext.Current.Request.QueryString.Get("qurantranslation"))
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
        Dim Names() As Array = Array.ConvertAll(Utility.GetChildNodes("sura", Utility.GetChildNode("suras", CachedData.XMLDocInfo.DocumentElement.ChildNodes).ChildNodes), Function(Convert As System.Xml.XmlNode) New Object() {Convert.Attributes.GetNamedItem("index").Value + ". " + GetChapterEName(Convert) + " (" + Arabic.TransliterateFromBuckwalter("suwrapu " + CachedData.IslamData.QuranChapters(CInt(Convert.Attributes.GetNamedItem("index").Value) - 1).Name + " ") + ")", CInt(Convert.Attributes.GetNamedItem("index").Value)})
        Array.Sort(Names, New Utility.CompareNameValueArray)
        Return Names
    End Function
    Public Shared Function GetChapterEName(ByVal ChapterNode As System.Xml.XmlNode) As String
        Return Utility.LoadResourceString("IslamInfo_QuranChapter" + ChapterNode.Attributes.GetNamedItem("index").Value)
    End Function
    Public Shared Function GetChapterNamesByRevelationOrder() As Array()
        Dim Names() As Array = Array.ConvertAll(Utility.GetChildNodes("sura", Utility.GetChildNode("suras", CachedData.XMLDocInfo.DocumentElement.ChildNodes).ChildNodes), Function(Convert As System.Xml.XmlNode) New Object() {Convert.Attributes.GetNamedItem("index").Value + ". " + GetChapterEName(Convert) + " (" + Arabic.TransliterateFromBuckwalter("suwrapu " + CachedData.IslamData.QuranChapters(CInt(Convert.Attributes.GetNamedItem("index").Value) - 1).Name + " ") + ")", CInt(Convert.Attributes.GetNamedItem("order").Value)})
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
        Dim TranslationIndex As Integer = GetTranslationIndex(GetCurrentCollection(), Web.HttpContext.Current.Request.QueryString.Get("hadithtranslation"))
        If TranslationIndex = -1 Then Return New String() {}
        Return New String() {Utility.GetPageString("Source&File=" + GetTranslationXMLFileName(GetCurrentCollection(), Web.HttpContext.Current.Request.QueryString.Get("hadithtranslation")) + ".xml"), CachedData.IslamData.Collections(GetCurrentCollection()).Translations(TranslationIndex).Name + " XML metadata"}
    End Function
    Public Shared Function GetTranslationTextDownload() As String()
        Dim TranslationIndex As Integer = GetTranslationIndex(GetCurrentCollection(), Web.HttpContext.Current.Request.QueryString.Get("hadithtranslation"))
        If TranslationIndex = -1 Then Return New String() {}
        Return New String() {Utility.GetPageString("Source&File=" + GetTranslationFileName(GetCurrentCollection(), Web.HttpContext.Current.Request.QueryString.Get("hadithtranslation")) + ".txt"), CachedData.IslamData.Collections(GetCurrentCollection()).Translations(TranslationIndex).Name + " raw source text"}
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
        Dim Strings As String = Web.HttpContext.Current.Request.QueryString.Get("hadithcollection")
        If Not Strings Is Nothing Then Return CInt(Strings) Else Return 0
    End Function
    Public Shared Function GetCurrentBook() As Integer
        Dim Strings As String = Web.HttpContext.Current.Request.QueryString.Get("hadithbook")
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
    Public Shared Function IsTranslationTextLTR(ByVal Index As Integer) As Boolean
        Dim TranslationIndex As Integer = GetTranslationIndex(Index, Web.HttpContext.Current.Request.QueryString.Get("hadithtranslation"))
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
        XMLDocTranslate.Load(Utility.GetFilePath("metadata\" + GetTranslationXMLFileName(Index, Web.HttpContext.Current.Request.QueryString.Get("hadithtranslation")) + ".xml"))
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
            Renderer.Items.Add(New RenderArray.RenderItem(RenderArray.RenderTypes.eHeaderLeft, New RenderArray.RenderText() {New RenderArray.RenderText(RenderArray.RenderDisplayClass.eArabic, Arabic.RightToLeftMark + Arabic.TransliterateFromBuckwalter("Had~iv " + BookNode.Attributes.GetNamedItem("hadiths").Value + " ")), New RenderArray.RenderText(RenderArray.RenderDisplayClass.eTransliteration, Arabic.TransliterateToScheme(Arabic.TransliterateFromBuckwalter("Had~iv " + BookNode.Attributes.GetNamedItem("hadiths").Value + " ")).Trim()), New RenderArray.RenderText(RenderArray.RenderDisplayClass.eLTR, "Hadiths: " + BookNode.Attributes.GetNamedItem("hadiths").Value + " ")}))
            Renderer.Items.Add(New RenderArray.RenderItem(RenderArray.RenderTypes.eHeaderCenter, New RenderArray.RenderText() {New RenderArray.RenderText(RenderArray.RenderDisplayClass.eArabic, Arabic.RightToLeftMark + Arabic.TransliterateFromBuckwalter("{lokita`bu " + CStr(BookIndex)) + " " + BookNode.Attributes.GetNamedItem("name").Value + " "), New RenderArray.RenderText(RenderArray.RenderDisplayClass.eTransliteration, Arabic.TransliterateToScheme(Arabic.TransliterateFromBuckwalter("{lokita`bu " + CStr(BookIndex)) + " " + BookNode.Attributes.GetNamedItem("name").Value + " ").Trim()), New RenderArray.RenderText(RenderArray.RenderDisplayClass.eLTR, "Book " + CStr(BookIndex) + ": " + GetBookEName(BookNode, Index) + " ")}))
            'Renderer.Items.Add(New RenderArray.RenderItem(RenderArray.RenderTypes.eHeaderRight, New RenderArray.RenderText() {New RenderArray.RenderText(RenderArray.RenderDisplayClass.eArabic, Arabic.RightToLeftMark + Arabic.TransliterateFromBuckwalter("mjld " + Utility.GetChildNode("books", XMLDocInfo(Index).DocumentElement.ChildNodes).ChildNodes.Item(BookIndex).Attributes.GetNamedItem("volume").Value + " ")), New RenderArray.RenderText(RenderArray.RenderDisplayClass.eTransliteration, Arabic.TransliterateToScheme(Arabic.TransliterateFromBuckwalter("mjld " + Utility.GetChildNode("books", XMLDocInfo(Index).DocumentElement.ChildNodes).ChildNodes.Item(BookIndex).Attributes.GetNamedItem("volume").Value + " ")).Trim()), New RenderArray.RenderText(RenderArray.RenderDisplayClass.eLTR, "Volume " + Utility.GetChildNode("books", XMLDocInfo(Index).DocumentElement.ChildNodes).ChildNodes.Item(BookIndex).Attributes.GetNamedItem("volume").Value + " ")}))
            Dim XMLDocTranslate As New System.Xml.XmlDocument
            Dim Strings() As String = Nothing
            If CachedData.IslamData.Collections(Index).Translations.Length <> 0 Then
                XMLDocTranslate.Load(Utility.GetFilePath("metadata\" + GetTranslationXMLFileName(Index, Web.HttpContext.Current.Request.QueryString.Get("hadithtranslation")) + ".xml"))
                Strings = IO.File.ReadAllLines(Utility.GetFilePath("metadata\" + GetTranslationFileName(Index, Web.HttpContext.Current.Request.QueryString.Get("hadithtranslation")) + ".txt"))
            End If
            For Hadith = 0 To HadithText.Count - 1
                'Handle missing or excess chapter indexes
                If ChapterIndex <> CInt(HadithText(Hadith)(1)) Then
                    ChapterIndex = CInt(HadithText(Hadith)(1))
                    ChapterNode = GetChapterByIndex(BookNode, ChapterIndex)
                    If Not ChapterNode Is Nothing Then
                        Renderer.Items.Add(New RenderArray.RenderItem(RenderArray.RenderTypes.eHeaderLeft, New RenderArray.RenderText() {New RenderArray.RenderText(RenderArray.RenderDisplayClass.eArabic, Arabic.RightToLeftMark + Arabic.TransliterateFromBuckwalter("Had~iv " + ChapterNode.Attributes.GetNamedItem("hadiths").Value + " ")), New RenderArray.RenderText(RenderArray.RenderDisplayClass.eTransliteration, Arabic.TransliterateToScheme(Arabic.TransliterateFromBuckwalter("Had~iv " + ChapterNode.Attributes.GetNamedItem("hadiths").Value + " ")).Trim()), New RenderArray.RenderText(RenderArray.RenderDisplayClass.eLTR, "Hadiths: " + ChapterNode.Attributes.GetNamedItem("hadiths").Value + " ")}))
                        Renderer.Items.Add(New RenderArray.RenderItem(RenderArray.RenderTypes.eHeaderCenter, New RenderArray.RenderText() {New RenderArray.RenderText(RenderArray.RenderDisplayClass.eArabic, Arabic.RightToLeftMark + Arabic.TransliterateFromBuckwalter("bAb " + CStr(ChapterIndex)) + " " + ChapterNode.Attributes.GetNamedItem("name").Value + " "), New RenderArray.RenderText(RenderArray.RenderDisplayClass.eTransliteration, Arabic.TransliterateToScheme(Arabic.TransliterateFromBuckwalter("bAb " + CStr(ChapterIndex)) + " " + ChapterNode.Attributes.GetNamedItem("name").Value + " ").Trim()), New RenderArray.RenderText(RenderArray.RenderDisplayClass.eLTR, "Chapter " + CStr(ChapterIndex) + ": " + Utility.DefaultValue(Utility.LoadResourceString("IslamInfo_" + CachedData.IslamData.Collections(Index).FileName + "Book" + BookNode.Attributes.GetNamedItem("index").Value + "Chapter" + ChapterNode.Attributes.GetNamedItem("index").Value), String.Empty) + " ")}))
                    End If
                    SubChapterIndex = -1
                End If
                'Handle missing or excess subchapter indexes
                If SubChapterIndex <> CInt(HadithText(Hadith)(2)) Then
                    SubChapterIndex = CInt(HadithText(Hadith)(2))
                    If Not ChapterNode Is Nothing Then
                        SubChapterNode = GetSubChapterByIndex(ChapterNode, SubChapterIndex)
                        If Not SubChapterNode Is Nothing Then
                            Renderer.Items.Add(New RenderArray.RenderItem(RenderArray.RenderTypes.eHeaderLeft, New RenderArray.RenderText() {New RenderArray.RenderText(RenderArray.RenderDisplayClass.eArabic, Arabic.RightToLeftMark + Arabic.TransliterateFromBuckwalter("Had~iv " + SubChapterNode.Attributes.GetNamedItem("hadiths").Value + " ")), New RenderArray.RenderText(RenderArray.RenderDisplayClass.eTransliteration, Arabic.TransliterateToScheme(Arabic.TransliterateFromBuckwalter("Had~iv " + SubChapterNode.Attributes.GetNamedItem("hadiths").Value + " ")).Trim()), New RenderArray.RenderText(RenderArray.RenderDisplayClass.eLTR, "Hadiths: " + SubChapterNode.Attributes.GetNamedItem("hadiths").Value + " ")}))
                            Renderer.Items.Add(New RenderArray.RenderItem(RenderArray.RenderTypes.eHeaderCenter, New RenderArray.RenderText() {New RenderArray.RenderText(RenderArray.RenderDisplayClass.eArabic, Arabic.RightToLeftMark + Arabic.TransliterateFromBuckwalter("bAb " + CStr(SubChapterIndex)) + " " + SubChapterNode.Attributes.GetNamedItem("name").Value + " "), New RenderArray.RenderText(RenderArray.RenderDisplayClass.eTransliteration, Arabic.TransliterateToScheme(Arabic.TransliterateFromBuckwalter("bAb " + CStr(SubChapterIndex)) + " " + SubChapterNode.Attributes.GetNamedItem("name").Value + " ").Trim()), New RenderArray.RenderText(RenderArray.RenderDisplayClass.eLTR, "Sub-Chapter " + CStr(SubChapterIndex) + ": " + Utility.DefaultValue(Utility.LoadResourceString("IslamInfo_" + CachedData.IslamData.Collections(Index).FileName + "Book" + BookNode.Attributes.GetNamedItem("index").Value + "Chapter" + ChapterNode.Attributes.GetNamedItem("index").Value + "Subchapter" + SubChapterNode.Attributes.GetNamedItem("index").Value), String.Empty) + " ")}))
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
                Renderer.Items.Add(New RenderArray.RenderItem(RenderArray.RenderTypes.eText, New RenderArray.RenderText() {New RenderArray.RenderText(RenderArray.RenderDisplayClass.eArabic, Arabic.RightToLeftMark + CStr(HadithText(Hadith)(3)) + " " + Arabic.TransliterateFromBuckwalter("=" + CStr(HadithText(Hadith)(0))) + " "), New RenderArray.RenderText(RenderArray.RenderDisplayClass.eTransliteration, Arabic.TransliterateToScheme(CStr(HadithText(Hadith)(3)) + " " + Arabic.TransliterateFromBuckwalter("=" + CStr(HadithText(Hadith)(0))) + " ").Trim()), New RenderArray.RenderText(DirectCast(IIf(IsTranslationTextLTR(Index), RenderArray.RenderDisplayClass.eLTR, RenderArray.RenderDisplayClass.eRTL), RenderArray.RenderDisplayClass), "(" + CStr(HadithText(Hadith)(0)) + ") " + HadithTranslation)}))
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
        SendEMail(EMail, String.Format(Utility.LoadResourceString("Acct_ActivationAccountSubject"), Web.HttpContext.Current.Request.Url.Host), _
            String.Format(Utility.LoadResourceString("Acct_ActivationAccountBody"), Web.HttpContext.Current.Request.Url.Host, UserName, "http://" + Web.HttpContext.Current.Request.Url.Host + "/" + Utility.GetPageString("ActivateAccount&UserID=" + CStr(UserID) + "&ActivationCode=" + CStr(ActivationCode)), "http://" + Web.HttpContext.Current.Request.Url.Host + "/" + Utility.GetPageString("ActivateAccount"), CStr(ActivationCode)))
    End Sub
    Public Shared Sub SendUserNameReminderEMail(ByVal UserName As String, ByVal EMail As String)
        SendEMail(EMail, String.Format(Utility.LoadResourceString("Acct_UsernameReminderSubject"), Web.HttpContext.Current.Request.Url.Host), _
            String.Format(Utility.LoadResourceString("Acct_UsernameReminderBody"), Web.HttpContext.Current.Request.Url.Host, UserName))
    End Sub
    Public Shared Sub SendPasswordResetEMail(ByVal UserName As String, ByVal EMail As String, ByVal UserID As Integer, ByVal PasswordResetCode As UInteger)
        SendEMail(EMail, String.Format(Utility.LoadResourceString("Acct_PasswordResetSubject"), Web.HttpContext.Current.Request.Url.Host), _
            String.Format(Utility.LoadResourceString("Acct_PasswordResetBody"), Web.HttpContext.Current.Request.Url.Host, UserName, "http://" + Web.HttpContext.Current.Request.Url.Host + "/" + Utility.GetPageString("ResetPassword&UserID=" + CStr(UserID) + "&PasswordResetCode=" + CStr(PasswordResetCode)), "http://" + Web.HttpContext.Current.Request.Url.Host + "/" + Utility.GetPageString("ResetPassword"), CStr(PasswordResetCode)))
    End Sub
    Public Shared Sub SendUserNameChangedEMail(ByVal UserName As String, ByVal EMail As String)
        SendEMail(EMail, String.Format(Utility.LoadResourceString("Acct_UsernameChangedSubject"), Web.HttpContext.Current.Request.Url.Host), _
            String.Format(Utility.LoadResourceString("Acct_UsernameChangedBody"), Web.HttpContext.Current.Request.Url.Host, UserName))
    End Sub
    Public Shared Sub SendPasswordChangedEMail(ByVal UserName As String, ByVal EMail As String)
        SendEMail(EMail, String.Format(Utility.LoadResourceString("Acct_PasswordChangedSubject"), Web.HttpContext.Current.Request.Url.Host), _
            String.Format(Utility.LoadResourceString("Acct_PasswordChangedBody"), Web.HttpContext.Current.Request.Url.Host, UserName))
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