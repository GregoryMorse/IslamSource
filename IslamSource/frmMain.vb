Imports HostPageUtility
Imports IslamMetadata
Imports XMLRender
Public Class frmMain
    Private _PortableMethods As PortableMethods
    Private TR As TanzilReader
    Private Arb As Arabic
    Private ArbData As ArabicData
    Private DocBuild As DocBuilder
    Private ChData As CachedData
    Private RAWeb As HostPageUtility.RenderArrayWeb
    Private UWeb As UtilityWeb
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
        Public Function GetTemplatePath(Selector As String) As String Implements PortableSettings.GetTemplatePath
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
    Private Shared Function ReadByteArray(ByVal s As IO.Stream) As Byte()
        Dim buffer As Byte() = New Byte(4 - 1) {}
        If (s.Read(buffer, 0, buffer.Length) <> buffer.Length) Then
            Throw New SystemException("Stream did not contain properly formatted byte array")
        End If
        Dim buffer2 As Byte() = New Byte(BitConverter.ToInt32(buffer, 0) - 1) {}
        If (s.Read(buffer2, 0, buffer2.Length) <> buffer2.Length) Then
            Throw New SystemException("Did not read byte array properly")
        End If
        Return buffer2
    End Function
    Public Shared Function Decrypt(ByVal cipherText As String, ByVal unknown As String) As String
        If String.IsNullOrEmpty(cipherText) Then
            Throw New ArgumentNullException("cipherText")
        End If
        If String.IsNullOrEmpty(unknown) Then
            Throw New ArgumentNullException("unknown")
        End If
        Dim managed As Security.Cryptography.RijndaelManaged = Nothing
        Try
            Dim bytes As New Security.Cryptography.Rfc2898DeriveBytes(unknown, System.Text.Encoding.ASCII.GetBytes("o6B06642kbM7c5"))
            Using stream As IO.MemoryStream = New IO.MemoryStream(Convert.FromBase64String(cipherText))
                managed = New Security.Cryptography.RijndaelManaged
                managed.Key = bytes.GetBytes((managed.KeySize \ 8))
                managed.IV = ReadByteArray(stream)
                Dim transform As Security.Cryptography.ICryptoTransform = managed.CreateDecryptor(managed.Key, managed.IV)
                Using stream2 As Security.Cryptography.CryptoStream = New Security.Cryptography.CryptoStream(stream, transform, Security.Cryptography.CryptoStreamMode.Read)
                    Using reader As IO.StreamReader = New IO.StreamReader(stream2)
                        Return reader.ReadToEnd
                    End Using
                End Using
            End Using
        Finally
            If (Not managed Is Nothing) Then
                managed.Clear()
            End If
        End Try
        Return Nothing
    End Function
    Private Sub CheckMorphology()
        For TestCount = 0 To ChData.IslamData.PartsOfSpeech.Length - 1
            _PortableMethods.LoadResourceString("IslamInfo_" + ChData.IslamData.PartsOfSpeech(TestCount).Id)
        Next
        For TestCount = 0 To ChData.IslamData.FeaturesOfSpeech.Length - 1
            _PortableMethods.LoadResourceString("IslamInfo_" + ChData.IslamData.FeaturesOfSpeech(TestCount).Id)
        Next
        Dim Text As List(Of String()) = TR.GetQuranText(ChData.XMLDocMain, 1, 1, 114, 6)
        For Count As Integer = 0 To 113
            For VerseCount = 0 To Text(Count).Length - 1
                Dim Words As String() = New List(Of String)(Linq.Enumerable.Where(Text(Count)(VerseCount).Split(" "c), Function(It) It.Length <> 1)).ToArray()
                For WordCount = 0 To Words.Length - 1
                    ChData.GetMorphologicalDataForWord(Count + 1, VerseCount + 1, WordCount + 1)
                Next
            Next
        Next
    End Sub
    Private Async Function CompareWordCounts() As Threading.Tasks.Task
        Dim Text As List(Of String()) = TR.GetQuranText(ChData.XMLDocMain, 1, 1, 114, 6)
        'Dim Lines As String() = Await _PortableMethods.ReadAllLines("..\..\..\metadata\en.w4w.corpus.txt")
        'Dim Lines As String() = Await _PortableMethods.ReadAllLines("..\..\..\metadata\en.w4w.qurandev.txt")
        'Dim Lines As String() = Await _PortableMethods.ReadAllLines("..\..\..\metadata\en.w4w.shehnazshaikh.txt")
        'Dim Lines As String() = Await _PortableMethods.ReadAllLines("..\..\..\metadata\id.w4w.terjemah.txt")
        'Dim Lines As String() = Await _PortableMethods.ReadAllLines("..\..\..\metadata\ur.w4w.hafiznazarmuhammad.txt")
        Dim Lines As String() = Await _PortableMethods.ReadAllLines("..\..\..\metadata\tr.w4w.suleymanates.txt")
        Dim LineCount As Integer = 0
        For Count As Integer = 0 To 113
            For VerseCount = 0 To Text(Count).Length - 1
                Dim Words As String() = New List(Of String)(Linq.Enumerable.Where(Text(Count)(VerseCount).Split(" "c), Function(It) It.Length <> 1)).ToArray()
                Dim TranslationWords As String() = Lines(LineCount).Split("|"c)
                If Words.Length <> TranslationWords.Length Then
                    Dim Str As String = String.Empty
                    For WordCount = 0 To Math.Max(Words.Length, TranslationWords.Length) - 1
                        Str += "["
                        If WordCount < Words.Length Then Str += Words(WordCount)
                        Str += " : "
                        If WordCount < TranslationWords.Length Then Str += TranslationWords(WordCount)
                        Str += "] "
                    Next
                    Debug.Print("Fault line: " + CStr(LineCount + 1) + "(" + CStr(Count + 1) + ":" + CStr(VerseCount + 1) + ") Arabic word count: " + CStr(Words.Length) + " Translation word count: " + CStr(TranslationWords.Length) + " " + CStr(Words.Length - TranslationWords.Length) + "(" + Str + ")")
                End If
                LineCount += 1
            Next
        Next
    End Function

    Private Async Function UtilityTestCode() As Threading.Tasks.Task
        'ChData.GetMorphologicalDataByVerbScale()
        Await TR.MakeQuranCacheMetarules()
        Dim Rules As List(Of Arabic.RuleMetadata) = Await TR.GetQuranCacheMetarules()
        Dim IndexToVerse As Integer()() = Nothing
        TR.QuranTextCombiner(ChData.XMLDocMain, IndexToVerse)
        Await _PortableMethods.WriteAllLines(_PortableMethods.FileIO.CombinePath(Await _PortableMethods.DiskCache.GetCacheDirectory(), "_QuranTajweedData.txt"), Arabic.MakeCacheMetarules(Rules, IndexToVerse))
        'Debug.Print(Decrypt("EAAAAKTSWJHpN/u15OHqSqZ3RhDB7UNHKMeY9Lk2sxW7Rcsc", "!234Qwer)987Poiu"))
        'clsWarshQuran.HindiW4W()
        'Dim Strs As New List(Of String)
        'For Chapter = 1 To 114
        '    For SubCount = 1 To TR.GetVerseCount(Chapter)
        '        Dim Text As String() = _PortableMethods.ReadAllLines("..\..\..\Resources\tr.w4w\kelime.asp@sure=" + CStr(Chapter) + "&ayet=" + CStr(SubCount))
        '        Dim Str As String = String.Empty
        '        Dim SplitStr As String = String.Empty
        '        Dim bComb As Integer = 0
        '        For ParseCount As Integer = 0 To Text.Length - 1
        '            If Text(ParseCount).StartsWith("			<font face=""Shaikh Hamdullah Mushaf"" style=""font-size: 20px"">") Then
        '                If Text(ParseCount).Replace("			<font face=""Shaikh Hamdullah Mushaf"" style=""font-size: 20px"">", String.Empty).Replace(" </font><br>", String.Empty) = "&#1610;&#1614;&#1619;&#1575;" Or Text(ParseCount).Replace("			<font face=""Shaikh Hamdullah Mushaf"" style=""font-size: 20px"">", String.Empty).Replace(" </font><br>", String.Empty) = "&#1610;&#1614;&#1575;" Then
        '                    bComb = 1
        '                Else
        '                    SplitStr = String.Join(String.Empty, Linq.Enumerable.Where(Text(ParseCount).Replace("			<font face=""Shaikh Hamdullah Mushaf"" style=""font-size: 20px"">", String.Empty).Replace("&#1610;&#1614;&#1619;&#1575; ", String.Empty).Replace("&#1610;&#1614;&#1575; ", String.Empty).Replace(" &#1769;", String.Empty).Replace(" </font><br>", String.Empty).ToCharArray(), Function(Ch) Ch = " "c)).Replace(" "c, "|"c)
        '                End If
        '            ElseIf Text(ParseCount).StartsWith("			<font face=""Verdana"" size=""2"" color=""#008000"">") Then
        '                If Str <> String.Empty And bComb <> 2 Then Str += "|"
        '                If bComb = 2 Then
        '                    Str += " "
        '                    bComb = 0
        '                End If
        '                If bComb = 1 Then bComb = 2
        '                Str += Text(ParseCount).Replace("			<font face=""Verdana"" size=""2"" color=""#008000"">", String.Empty).Replace("</font></p>", String.Empty) + SplitStr
        '            End If
        '        Next
        '        Strs.Add(Str)
        '    Next
        'Next
        '_PortableMethods.WriteAllLines("..\..\..\metadata\tr.w4w.suleyman.ates.txt", Strs.ToArray())
        'CheckMorphology()        
        'dataroot -> ayat -> TarjumaLafziDrFarhatHashmi, TarjumaLafziFahmulQuran, TarjumaLafziNazarAhmad
        'CompareWordCounts()
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
        'state=".*".*>(?:(?!.*trans-unit.*).*\n)*?.*</target>\r\n.*</note> -> state="new"></target>
        'state="needs-review-translation"(.*)&lt;html&gt;(?:(?!.*trans-unit.*).*\n)+?.*&gt;</ -> state="new"></
        '\s*<it id="1" pos="open">&lt;(.*)&gt;</it> -> $1
        'state="needs-review-translation"(.*)" onmouseover=.*(?=</target>) -> state="new">
        'TR.WordFileToResource("..\..\..\metadata\id.w4w.terjemah.txt", "..\..\..\ResourceToolkitTempForeign\QuranResources.resx")
        'TR.ResourceToWordFile("..\..\..\ResourceToolkitTempForeign\QuranResources.ms.resx", "..\..\..\metadata\ms.w4w.terjemah.txt")
        'TR.WordFileToResource("..\..\..\metadata\en.w4w.corpus.txt", "..\..\..\ResourceToolkitTemp\QuranResources.resx")
        'TR.FileToResource("..\..\..\metadata\en.sahih.txt", "..\..\..\ResourceToolkitQuran\QuranResources.resx")
        For Each File As IO.FileInfo In New IO.DirectoryInfo("..\..\..\ResourceToolkitTemp\").EnumerateFiles("QuranResources.*.resx")
            Await TR.ResourceToWordFile(File.FullName, "..\..\..\metadata\" + File.Name.Replace("QuranResources.", String.Empty).Replace(".resx", String.Empty) + ".w4w.corpus.txt")
        Next
        '_PortableMethods.WordFileToResource("..\..\..\metadata\en.w4w.txt", "..\..\..\ResourceToolkit\W4WResources.resx")
        For Each File As IO.FileInfo In New IO.DirectoryInfo("..\..\..\ResourceToolkit\").EnumerateFiles("W4WResources.*.resx")
            Await _PortableMethods.ResourceToWordFile(File.FullName, "..\..\..\metadata\" + File.Name.Replace("W4WResources.", String.Empty).Replace(".resx", String.Empty) + ".w4w.txt")
        Next
        'clsWarshQuran.ParseQuran()
        'Await TR.ChangeQuranFormat(TanzilReader.QuranTexts.Warsh, TanzilReader.QuranTexts.Warsh, TanzilReader.QuranScripts.Uthmani, TanzilReader.ArabicPresentation.Buckwalter)
        'Await TR.ChangeQuranFormat(TanzilReader.QuranTexts.Hafs, TanzilReader.QuranTexts.Hafs, TanzilReader.QuranScripts.Uthmani, TanzilReader.ArabicPresentation.Buckwalter)
        'Await TR.ChangeQuranFormat(TanzilReader.QuranTexts.Hafs, TanzilReader.QuranTexts.Hafs, TanzilReader.QuranScripts.UthmaniMin, TanzilReader.ArabicPresentation.None)
        Await TR.ChangeQuranFormat(TanzilReader.QuranTexts.Hafs, TanzilReader.QuranTexts.Hafs, TanzilReader.QuranScripts.SimpleEnhanced, TanzilReader.QuranScripts.Uthmani, TanzilReader.ArabicPresentation.Buckwalter)
        'Await TR.ChangeQuranFormat(TanzilReader.QuranTexts.Hafs, TanzilReader.QuranTexts.Hafs, TanzilReader.QuranScripts.SimpleEnhanced, TanzilReader.ArabicPresentation.Buckwalter)
        'Await TR.ChangeQuranFormat(TanzilReader.QuranTexts.Hafs, TanzilReader.QuranTexts.Hafs, TanzilReader.QuranScripts.Simple, TanzilReader.ArabicPresentation.Buckwalter)
        'Await TR.ChangeQuranFormat(TanzilReader.QuranTexts.Hafs, TanzilReader.QuranTexts.Hafs, TanzilReader.QuranScripts.SimpleClean, TanzilReader.ArabicPresentation.Buckwalter)
        'Await TR.ChangeQuranFormat(TanzilReader.QuranTexts.Hafs, TanzilReader.QuranTexts.Hafs, TanzilReader.QuranScripts.SimpleMin, TanzilReader.ArabicPresentation.Buckwalter)
        'Await TR.ChangeQuranFormat(TanzilReader.QuranTexts.Hafs, TanzilReader.QuranTexts.Hafs, TanzilReader.QuranScripts.SimpleEnhanced, TanzilReader.ArabicPresentation.None)
        'Await TR.ChangeQuranFormat(TanzilReader.QuranTexts.Hafs, TanzilReader.QuranTexts.Hafs, TanzilReader.QuranScripts.Simple, TanzilReader.ArabicPresentation.None)
        'Await TR.ChangeQuranFormat(TanzilReader.QuranTexts.Hafs, TanzilReader.QuranTexts.Hafs, TanzilReader.QuranScripts.SimpleMin, TanzilReader.ArabicPresentation.None)
        'Await TR.ChangeQuranFormat(TanzilReader.QuranTexts.Hafs, TanzilReader.QuranTexts.Hafs, TanzilReader.QuranScripts.SimpleClean, TanzilReader.ArabicPresentation.None)
        'Await TR.CompareQuranFormats(TanzilReader.QuranTexts.Hafs, TanzilReader.QuranTexts.Warsh, TanzilReader.QuranScripts.Uthmani, TanzilReader.ArabicPresentation.Buckwalter)
        'Await TR.ChangeQuranFormat(TanzilReader.QuranTexts.Hafs, TanzilReader.QuranTexts.Warsh, TanzilReader.QuranScripts.Uthmani, TanzilReader.ArabicPresentation.Buckwalter)
        'TR.CheckNotablePatterns()
        'TR.FindMinimalVersesForCoverage()
        'TR.CheckSequentialRules()
        'TR.CheckMutualExclusiveRules(False, 4)
        'TR.CheckMutualExclusiveRules(True, 4)

        'Dim Text As String = TanzilReader.QuranTextCombiner(CachedData.XMLDocMain, IndexToVerse)
        'Dim IndVerses As String() = TanzilReader.QuranTextRangeLookup(2, 5, 0, 2, 0, 0)(0)
        Arb.TransliterateWithRulesColor(Arb.TransliterateFromBuckwalter("mina {lojin~api wa{ln~aAsi"), "PlainRoman", True, False, Arabic.FilterMetadataStops(Arb.TransliterateFromBuckwalter("mina {lojin~api wa{ln~aAsi"), Arb.GetMetarules(Arb.TransliterateFromBuckwalter("mina {lojin~api wa{ln~aAsi"), ChData.RuleMetas("UthmaniQuran")), Nothing))
        Text = TanzilReader.GetQuranText(ChData.XMLDocMain, 19, 4, 4)(0) + " " + ArabicData.ArabicEndOfAyah
        'Debug.Print(Arb.TransliterateToScheme(IndVerses(0), ArabicData.TranslitScheme.RuleBased, String.Empty, CachedData.RuleMetas("UthmaniQuran"), TanzilReader.GenerateDefaultStops(IndVerses(0))))
        Debug.Print(Arb.TransliterateToScheme(Text, ArabicData.TranslitScheme.RuleBased, String.Empty, Arabic.FilterMetadataStops(Text, Arb.GetMetarules(Text, ChData.RuleMetas("UthmaniQuran")), TR.GenerateDefaultStops(Text))))
        For Selection = 33 To TR.GetChapterCount()
            Await TR.GetRenderedQuranText(ArabicData.TranslitScheme.RuleBased, String.Empty, String.Empty, "0", Selection.ToString(), String.Empty, "0", "1")
            Debug.Print(CStr(Selection))
        Next
        'Debug.Print(Arabic.TransliterateToScheme(System.Text.RegularExpressions.Regex.Match(Text, CachedData.GetPattern("ContinuousMatch")).Value, ArabicData.TranslitScheme.Literal, String.Empty))
        'Utility.SortResX(Utility.GetFilePath("IslamResources\Resources.en.resx"))
        'Utility.SortResX(Utility.GetFilePath("IslamResources\My Project\Resources.resx"))
        Await ChData.DoErrorCheck(TR, DocBuild)
        'Dim RenderArr As RenderArray = Phrases.DoGetRenderedCatText(String.Empty, ArabicData.TranslitScheme.RuleBased, String.Empty, CachedData.IslamData.Phrases, 0)
        'HostPageUtility.RenderArray.OutputPdf("test.pdf", RenderArr.Items)
        For Count As Integer = 39 To TR.GetChapterCount()
            Dim RA As RenderArray = Await TR.GetQuranTextBySelection(String.Empty, 0, Count, String.Empty, ArabicData.TranslitScheme.RuleBased, String.Empty, TR.GetTranslationIndex(String.Empty), True, False, False, True, False, False, True)
            RAWeb.OutputPdf("test" + CStr(Count) + ".pdf", RA.Items)
        Next
    End Function
    Private Async Function Init() As Threading.Tasks.Task
        _PortableMethods = New PortableMethods(New WindowsWebFileIO, New WindowsSettings)
        Await _PortableMethods.Init()
        ArbData = New XMLRender.ArabicData(_PortableMethods)
        Await ArbData.Init()
        Arb = New IslamMetadata.Arabic(_PortableMethods, ArbData)
        ChData = New IslamMetadata.CachedData(_PortableMethods, ArbData, Arb)
        Await ChData.Init()
        Await Arb.Init(ChData)
        TR = New IslamMetadata.TanzilReader(_PortableMethods, Arb, ArbData, ChData)
        Await TR.Init()
        DocBuild = New IslamMetadata.DocBuilder(_PortableMethods, Arb, ArbData, ChData)
        UWeb = New UtilityWeb(_PortableMethods, ArbData, Nothing, Nothing, Nothing)
        PageSet = New PageLoader(_PortableMethods, UWeb)
        Await UtilityTestCode()
        RenderItems = (Await TR.GetQuranTextBySelection(String.Empty, 0, 1, String.Empty, ArabicData.TranslitScheme.RuleBased, String.Empty, TR.GetTranslationIndex(String.Empty), True, False, False, True, False, False, True)).Items
    End Function
    Private PageSet As PageLoader
    Private RenderItems As List(Of RenderArray.RenderItem)
    Private Async Sub frmMain_Load(sender As Object, e As EventArgs) Handles Me.Load
        Await Init()
        For Index = 0 To PageSet.Pages.Count - 1
            Dim newNode As TreeNode = tvwMain.Nodes.Add(PageSet.Pages.Item(Index).PageName, _PortableMethods.LoadResourceString(PageSet.Pages.Item(Index).Text))
            For SubIndex = 0 To PageSet.Pages.Item(Index).Page.Count - 1
                If PageLoader.IsListItem(PageSet.Pages.Item(Index).Page(SubIndex)) AndAlso DirectCast(PageSet.Pages.Item(Index).Page(SubIndex), PageLoader.ListItem).IsSection Then
                    newNode.Nodes.Add(PageSet.Pages.Item(Index).PageName + CStr(If(DirectCast(PageSet.Pages.Item(Index).Page.Item(SubIndex), PageLoader.ListItem).Name <> String.Empty, "#" + DirectCast(PageSet.Pages.Item(Index).Page.Item(SubIndex), PageLoader.ListItem).Name, String.Empty)), _PortableMethods.LoadResourceString(DirectCast(PageSet.Pages.Item(Index).Page.Item(SubIndex), PageLoader.ListItem).Title))
                End If
            Next
        Next
        Dim Renderer As New MultiLangRender
        Renderer.Anchor = AnchorStyles.Left Or AnchorStyles.Top Or AnchorStyles.Right Or AnchorStyles.Bottom
        Renderer.Size = New Size(gbMain.Width, gbMain.Height)
        Renderer.RenderArray = RenderItems
        gbMain.Controls.Add(Renderer)
    End Sub

    Private Sub gbMain_Resize(sender As Object, e As EventArgs) Handles gbMain.Resize
        If gbMain.Controls.Count <> 0 Then
            gbMain.Controls(0).Size = New Size(gbMain.Width, gbMain.Height)
        End If
    End Sub
End Class
