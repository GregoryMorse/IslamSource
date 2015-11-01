Option Explicit On
Option Strict On
Imports HostPageUtility
Imports System.Drawing
Imports System.Web
Imports System.Web.UI
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
            Compare = GetSchemeValueFromSymbol(DirectCast(x, ArabicData.ArabicSymbol), _Scheme).Length - _
                GetSchemeValueFromSymbol(DirectCast(y, ArabicData.ArabicSymbol), _Scheme).Length
            If Compare = 0 Then Compare = GetSchemeValueFromSymbol(DirectCast(x, ArabicData.ArabicSymbol), _Scheme).CompareTo(GetSchemeValueFromSymbol(DirectCast(y, ArabicData.ArabicSymbol), _Scheme))
        End Function
    End Class
    Public Shared Function GetRecitationSymbols() As Array()
        Return Array.ConvertAll(CachedData.RecitationSymbols, Function(Ch As String) New Object() {ArabicData.ArabicLetters(ArabicData.FindLetterBySymbol(Ch.Chars(0))).UnicodeName + " (" + ArabicData.FixStartingCombiningSymbol(Ch) + ArabicData.LeftToRightEmbedding + ")" + ArabicData.PopDirectionalFormatting, ArabicData.FindLetterBySymbol(Ch.Chars(0))})
    End Function
    Public Shared _BuckwalterMap As Dictionary(Of Char, Integer)
    Public Shared ReadOnly Property BuckwalterMap As Dictionary(Of Char, Integer)
        Get
            If _BuckwalterMap Is Nothing Then
                If Not DiskCache.GetCacheItem("BuckwalterMap", DateTime.MinValue) Is Nothing Then
                    _BuckwalterMap = CType((New System.Runtime.Serialization.Formatters.Binary.BinaryFormatter).Deserialize(New IO.MemoryStream(DiskCache.GetCacheItem("BuckwalterMap", DateTime.MinValue))), Dictionary(Of Char, Integer))
                Else
                    _BuckwalterMap = New Dictionary(Of Char, Integer)
                    For Index = 0 To ArabicData.ArabicLetters.Length - 1
                        If GetSchemeValueFromSymbol(ArabicData.ArabicLetters(Index), "ExtendedBuckwalter").Length <> 0 Then
                            _BuckwalterMap.Add(GetSchemeValueFromSymbol(ArabicData.ArabicLetters(Index), "ExtendedBuckwalter").Chars(0), Index)
                        End If
                    Next
                    Dim MemStream As New IO.MemoryStream
                    Dim Ser As New System.Runtime.Serialization.Formatters.Binary.BinaryFormatter
                    Ser.Serialize(MemStream, _BuckwalterMap)
                    DiskCache.CacheItem("BuckwalterMap", Now, MemStream.ToArray())
                    MemStream.Close()
                End If
            End If
            Return _BuckwalterMap
        End Get
    End Property
    Public Shared Function TransliterateFromBuckwalter(ByVal Buckwalter As String) As String
        Dim ArabicString As New System.Text.StringBuilder
        Dim Count As Integer
        If Buckwalter Is Nothing Then Return ArabicString.ToString()
        For Count = 0 To Buckwalter.Length - 1
            If Buckwalter(Count) = "\" Then
                Count += 1
                If Buckwalter(Count) = "," Then
                    ArabicString.Append(ArabicData.ArabicComma)
                ElseIf Buckwalter(Count) = ";" Then
                    ArabicString.Append(ArabicData.ArabicSemicolon)
                ElseIf Buckwalter(Count) = "?" Then
                    ArabicString.Append(ArabicData.ArabicQuestionMark)
                Else
                    ArabicString.Append(Buckwalter(Count))
                End If
            Else
                If BuckwalterMap.ContainsKey(Buckwalter(Count)) Then
                    ArabicString.Append(ArabicData.ArabicLetters(BuckwalterMap.Item(Buckwalter(Count))).Symbol)
                Else
                    ArabicString.Append(Buckwalter(Count))
                End If
            End If
        Next
        Return ArabicString.ToString()
    End Function
    Public Shared Function TransliterateToScheme(ByVal ArabicString As String, SchemeType As ArabicData.TranslitScheme, Scheme As String, Optional OptionalStops() As Integer = Nothing, Optional PreString As String = "", Optional PostString As String = "", Optional PreStop As Boolean = True, Optional PostStop As Boolean = True) As String
        If SchemeType = ArabicData.TranslitScheme.LearningMode Then
            Return TransliterateWithRules(ArabicString, Scheme, OptionalStops, True)
        ElseIf SchemeType = ArabicData.TranslitScheme.RuleBased And PreString = "" And PostString = "" Then
            Return TransliterateWithRules(ArabicString, Scheme, OptionalStops, False)
        ElseIf SchemeType = ArabicData.TranslitScheme.RuleBased Then
            Return TransliterateContigWithRules(ArabicString, PreString, PostString, PreStop, PostStop, Scheme, OptionalStops)
        ElseIf SchemeType = ArabicData.TranslitScheme.Literal Then
            Return TransliterateToRoman(ArabicString, Scheme)
        Else
            Return New String(System.Array.FindAll(ArabicString.ToCharArray(), Function(Check As Char) Check = " "c))
        End If
    End Function
    Shared _SchemeTable As Dictionary(Of String, IslamData.TranslitScheme)
    Public Shared ReadOnly Property SchemeTable() As Dictionary(Of String, IslamData.TranslitScheme)
        Get
            If _SchemeTable Is Nothing Then
                _SchemeTable = New Dictionary(Of String, IslamData.TranslitScheme)
                Dim Count As Integer
                For Count = 0 To CachedData.IslamData.TranslitSchemes.Length - 1
                    _SchemeTable.Add(CachedData.IslamData.TranslitSchemes(Count).Name, CachedData.IslamData.TranslitSchemes(Count))
                Next
            End If
            Return _SchemeTable
        End Get
    End Property
    Public Shared Function GetSchemeSpecialValue(Str As String, Index As Integer, Scheme As String) As String
        If Not SchemeTable.ContainsKey(Scheme) Then Return String.Empty
        Dim Sch As IslamData.TranslitScheme = SchemeTable(Scheme)
        Return Sch.SpecialLetters(Index).Replace("&first;", If(Str = CachedData.ArabicSpecialLetters(Index), "*", GetSchemeValueFromSymbol(ArabicData.ArabicLetters(ArabicData.FindLetterBySymbol(Str(0))), Scheme)))
    End Function
    Public Shared Function GetSchemeSpecialFromMatch(Str As String, bExp As Boolean) As Integer
        If bExp Then
            For Count = 0 To CachedData.ArabicSpecialLetters.Length - 1
                If System.Text.RegularExpressions.Regex.Match(CachedData.ArabicSpecialLetters(Count), Str).Success Then Return Count
            Next
        Else
            If Array.IndexOf(CachedData.ArabicSpecialLetters, Str) <> -1 Then
                Return Array.IndexOf(CachedData.ArabicSpecialLetters, Str)
            End If
        End If
        Return -1
    End Function
    Public Shared Function GetSchemeLongVowel(Str As String, Scheme As String) As Integer
        If Not SchemeTable.ContainsKey(Scheme) Then Return -1
        Dim Sch As IslamData.TranslitScheme = SchemeTable(Scheme)
        If Array.IndexOf(CachedData.ArabicMultis, Str) <> -1 Then
            Return Array.IndexOf(CachedData.ArabicMultis, Str)
        End If
        Return -1
    End Function
    Public Shared Function GetSchemeLongVowelFromString(Str As String, Scheme As String) As String
        If Not SchemeTable.ContainsKey(Scheme) Then Return String.Empty
        Dim Sch As IslamData.TranslitScheme = SchemeTable(Scheme)
        If Array.IndexOf(CachedData.ArabicMultis, Str) <> -1 Then
            Return Sch.Multis(Array.IndexOf(CachedData.ArabicMultis, Str))
        End If
        Return String.Empty
    End Function
    Public Shared Function GetSchemeGutteralFromString(Str As String, Scheme As String, Leading As Boolean) As String
        If Not SchemeTable.ContainsKey(Scheme) Then Return String.Empty
        Dim Sch As IslamData.TranslitScheme = SchemeTable(Scheme)
        If Array.IndexOf(CachedData.ArabicLeadingGutterals, Str) <> -1 Then
            Return Sch.Gutterals(Array.IndexOf(CachedData.ArabicLeadingGutterals, Str) + If(Leading, CachedData.ArabicLeadingGutterals.Length, 0))
        End If
        Return String.Empty
    End Function
    Public Shared Function GetSchemeValueFromSymbol(Symbol As ArabicData.ArabicSymbol, Scheme As String) As String
        If Not SchemeTable.ContainsKey(Scheme) Then Return String.Empty
        Dim Sch As IslamData.TranslitScheme = SchemeTable(Scheme)
        If Array.IndexOf(CachedData.ArabicLettersInOrder, CStr(Symbol.Symbol)) <> -1 Then
            Return Sch.Alphabet(Array.IndexOf(CachedData.ArabicLettersInOrder, CStr(Symbol.Symbol)))
        ElseIf Array.IndexOf(CachedData.ArabicHamzas, CStr(Symbol.Symbol)) <> -1 Then
            Return Sch.Hamza(Array.IndexOf(CachedData.ArabicHamzas, CStr(Symbol.Symbol)))
        ElseIf Array.IndexOf(CachedData.ArabicVowels, CStr(Symbol.Symbol)) <> -1 Then
            Return Sch.Vowels(Array.IndexOf(CachedData.ArabicVowels, CStr(Symbol.Symbol)))
        ElseIf Array.IndexOf(CachedData.ArabicTajweed, CStr(Symbol.Symbol)) <> -1 Then
            Return Sch.Tajweed(Array.IndexOf(CachedData.ArabicTajweed, CStr(Symbol.Symbol)))
        ElseIf Array.IndexOf(CachedData.ArabicSilent, CStr(Symbol.Symbol)) <> -1 Then
            Return Sch.Silent(Array.IndexOf(CachedData.ArabicSilent, CStr(Symbol.Symbol)))
        ElseIf Array.IndexOf(CachedData.ArabicPunctuation, CStr(Symbol.Symbol)) <> -1 Then
            Return Sch.Punctuation(Array.IndexOf(CachedData.ArabicPunctuation, CStr(Symbol.Symbol)))
        ElseIf Array.IndexOf(CachedData.ArabicNums, CStr(Symbol.Symbol)) <> -1 Then
            Return Sch.Numbers(Array.IndexOf(CachedData.ArabicNums, CStr(Symbol.Symbol)))
        ElseIf Array.IndexOf(CachedData.NonArabicLetters, CStr(Symbol.Symbol)) <> -1 Then
            Return Sch.NonArabic(Array.IndexOf(CachedData.NonArabicLetters, CStr(Symbol.Symbol)))
        End If
        Return String.Empty
    End Function
    Public Shared Function SchemeHasValue(Str As String, Scheme As String) As Boolean
        If Not SchemeTable.ContainsKey(Scheme) Then Return False
        Dim Sch As IslamData.TranslitScheme = SchemeTable(Scheme)
        For Count = 0 To Sch.Alphabet.Length - 1
            If Sch.Alphabet(Count) = Str Then Return True
        Next
        Return False
    End Function
    Shared _Letters As Dictionary(Of Char, Integer)
    Public Shared Function GetSortedLetters(Scheme As String) As Dictionary(Of Char, Integer)
        If _Letters Is Nothing Then
            Dim Letters(ArabicData.ArabicLetters.Length - 1) As ArabicData.ArabicSymbol
            ArabicData.ArabicLetters.CopyTo(Letters, 0)
            Array.Sort(Letters, New StringLengthComparer(Scheme))
            _Letters = New Dictionary(Of Char, Integer)
            For Count = 0 To Letters.Length - 1
                _Letters.Add(Letters(Count).Symbol, Count)
            Next
        End If
        Return _Letters
    End Function
    Public Shared Function TransliterateToRoman(ByVal ArabicString As String, Scheme As String) As String
        Dim RomanString As New System.Text.StringBuilder
        Dim Count As Integer = 0
        While Count <= ArabicString.Length - 1
            If ArabicString(Count) = "\" Then
                Count += 1
                If ArabicString(Count) = "," Then
                    RomanString.Append(ArabicData.ArabicComma)
                ElseIf ArabicString(Count) = ";" Then
                    RomanString.Append(ArabicData.ArabicSemicolon)
                ElseIf ArabicString(Count) = "?" Then
                    RomanString.Append(ArabicData.ArabicQuestionMark)
                Else
                    RomanString.Append(ArabicString(Count))
                End If
            Else
                If GetSchemeSpecialFromMatch(ArabicString.Substring(Count), False) <> -1 Then
                    RomanString.Append(GetSchemeSpecialValue(ArabicString.Substring(Count), GetSchemeSpecialFromMatch(ArabicString.Substring(Count), False), If(Scheme = String.Empty, "ExtendedBuckwalter", Scheme)))
                    Count += System.Text.RegularExpressions.Regex.Match(CachedData.ArabicSpecialLetters(GetSchemeSpecialFromMatch(ArabicString.Substring(Count), False)), ArabicString.Substring(Count)).Value.Length - 1
                ElseIf ArabicString.Length - Count > 1 AndAlso GetSchemeLongVowel(ArabicString.Substring(Count, 2), If(Scheme = String.Empty, "ExtendedBuckwalter", Scheme)) <> -1 Then
                    RomanString.Append(GetSchemeLongVowelFromString(ArabicString.Substring(Count, 2), If(Scheme = String.Empty, "ExtendedBuckwalter", Scheme)))
                    Count += 1
                ElseIf GetSortedLetters(Scheme).ContainsKey(ArabicString(Count)) Then
                    RomanString.Append(GetSchemeValueFromSymbol(ArabicData.ArabicLetters(ArabicData.FindLetterBySymbol(ArabicString(Count))), If(Scheme = String.Empty, "ExtendedBuckwalter", Scheme)))
                Else
                    RomanString.Append(ArabicString(Count))
                End If
            End If
            Count += 1
        End While
        Return RomanString.ToString()
    End Function
    Structure RuleMetadata
        Sub New(NewIndex As Integer, NewLength As Integer, NewType As String, NewOrigOrder As Integer)
            Index = NewIndex
            Length = NewLength
            Type = NewType
            OrigOrder = NewOrigOrder
        End Sub
        Public Index As Integer
        Public Length As Integer
        Public Type As String
        Public OrigOrder As Integer
        Public Children As RuleMetadata()
    End Structure
    Public Shared SimpleTrailingAlef As String = ArabicData.ArabicLetterAlef + ArabicData.ArabicSmallHighRoundedZero
    Public Shared SimpleSuperscriptAlef As String = ArabicData.ArabicLetterSuperscriptAlef
    Public Shared UthmaniShortVowelsBeforeLongVowelsSuperscriptAlef As String = ArabicData.ArabicFatha + ArabicData.ArabicLetterSuperscriptAlef
    Public Shared UthmaniShortVowelsBeforeLongVowelsAlefAlefMaksura As String = ArabicData.ArabicFatha + ArabicData.ArabicLetterAlefMaksura
    Public Shared UthmaniShortVowelsBeforeLongVowelsYehAlefMaksura As String = ArabicData.ArabicKasra + ArabicData.ArabicLetterAlefMaksura
    Public Shared UthmaniShortVowelsBeforeLongVowelsAlef As String = ArabicData.ArabicFatha + ArabicData.ArabicLetterAlef
    Public Shared UthmaniShortVowelsBeforeLongVowelsWaw As String = ArabicData.ArabicDamma + ArabicData.ArabicLetterWaw
    Public Shared UthmaniShortVowelsBeforeLongVowelsSmallWaw As String = ArabicData.ArabicDamma + ArabicData.ArabicSmallWaw
    Public Shared UthmaniShortVowelsBeforeLongVowelsYeh As String = ArabicData.ArabicKasra + ArabicData.ArabicLetterYeh
    Public Shared UthmaniShortVowelsBeforeLongVowelsSmallYeh As String = ArabicData.ArabicKasra + ArabicData.ArabicSmallYeh

    Public Shared PrefixPattern As String = ""

    Public Enum RuleFuncs As Integer
        eNone
        eUpperCase
        eSpellNumber
        eSpellLetter
        eLookupLetter
        eLookupLongVowelDipthong
        eDivideTanween
        eLeadingGutteral
        eTrailingGutteral
        eResolveAmbiguity
        eLearningMode
    End Enum
    Public Delegate Function RuleFunction(Str As String, Scheme As String, LearningMode As Boolean) As String()
    Public Shared RuleFunctions As RuleFunction() = {
        Function(Str As String, Scheme As String, LearningMode As Boolean) {UCase(Str)},
        Function(Str As String, Scheme As String, LearningMode As Boolean) {TransliterateWithRules(Arabic.TransliterateFromBuckwalter(Arabic.ArabicWordFromNumber(CInt(TransliterateToScheme(Str, ArabicData.TranslitScheme.Literal, String.Empty)), True, False, False)), Scheme, Nothing, LearningMode)},
        Function(Str As String, Scheme As String, LearningMode As Boolean) {TransliterateWithRules(ArabicLetterSpelling(Str, True), Scheme, Nothing, LearningMode)},
        Function(Str As String, Scheme As String, LearningMode As Boolean) {GetSchemeValueFromSymbol(ArabicData.ArabicLetters(ArabicData.FindLetterBySymbol(Str.Chars(0))), Scheme)},
        Function(Str As String, Scheme As String, LearningMode As Boolean) {GetSchemeLongVowelFromString(Str, Scheme)},
        Function(Str As String, Scheme As String, LearningMode As Boolean) {CachedData.ArabicFathaDammaKasra(Array.IndexOf(CachedData.ArabicTanweens, Str)), ArabicData.ArabicLetterNoon},
        Function(Str As String, Scheme As String, LearningMode As Boolean) {GetSchemeGutteralFromString(Str.Remove(Str.Length - 1), Scheme, True) + Str.Chars(Str.Length - 1)},
        Function(Str As String, Scheme As String, LearningMode As Boolean) {Str.Chars(0) + GetSchemeGutteralFromString(Str.Remove(0, 1), Scheme, False)},
        Function(Str As String, Scheme As String, LearningMode As Boolean) {If(SchemeHasValue(GetSchemeValueFromSymbol(ArabicData.ArabicLetters(ArabicData.FindLetterBySymbol(Str.Chars(0))), Scheme) + GetSchemeValueFromSymbol(ArabicData.ArabicLetters(ArabicData.FindLetterBySymbol(Str.Chars(1))), Scheme), Scheme), Str.Chars(0) + "-" + Str.Chars(1), Str)},
        Function(Str As String, Scheme As String, LearningMode As Boolean) If(LearningMode, {Str, String.Empty}, {String.Empty, Str})
    }
        'Javascript does not support negative or positive lookbehind in regular expressions
    Public Shared AllowZeroLength As String() = {"helperfatha", "helperdamma", "helperkasra", "helperlparen", "helperrparen", "learningmode(helperslash,)", "learningmode(helperlbracket,)", "learningmode(helperrbracket,)", "learningmode(helperfathatan,)", "learningmode(helperteh,)"}
    Public Shared Function IsLetter(Index As Integer) As Boolean
        Return Array.FindIndex(CachedData.ArabicLetters, Function(Str As String) Str = ArabicData.ArabicLetters(Index).Symbol) <> -1
    End Function
    Public Shared Function IsPunctuation(Index As Integer) As Boolean
        Return Array.FindIndex(CachedData.PunctuationSymbols, Function(Str As String) Str = ArabicData.ArabicLetters(Index).Symbol) <> -1
    End Function
    Public Shared Function IsStop(Index As Integer) As Boolean
        Return Array.FindIndex(CachedData.ArabicStopLetters, Function(Str As String) Str = ArabicData.ArabicLetters(Index).Symbol) <> -1
    End Function
    Public Shared Function IsWhitespace(Index As Integer) As Boolean
        Return Array.FindIndex(CachedData.WhitespaceSymbols, Function(Str As String) Str = ArabicData.ArabicLetters(Index).Symbol) <> -1
    End Function
    Public Shared Function ArabicLetterSpelling(Input As String, Quranic As Boolean) As String
        Dim Output As String = String.Empty
        For Each Ch As Char In Input
            Dim Index As Integer = ArabicData.FindLetterBySymbol(Ch)
            If Index <> -1 AndAlso IsLetter(Index) Then
                If Output <> String.Empty And Not Quranic Then Output += " "
                Dim Idx As Integer = Array.IndexOf(CachedData.ArabicLettersInOrder, CStr(ArabicData.ArabicLetters(Index).Symbol))
                Output += If(Quranic, CachedData.ArabicAlphabet(Idx).Remove(CachedData.ArabicAlphabet(Idx).Length - 1) + If(CachedData.ArabicAlphabet(Idx).EndsWith("n"), String.Empty, "o"), CachedData.ArabicAlphabet(Idx))
            ElseIf Index <> -1 AndAlso ArabicData.ArabicLetters(Index).Symbol = ArabicData.ArabicMaddahAbove Then
                If Not Quranic Then Output += Ch
            End If
        Next
        Return Arabic.TransliterateFromBuckwalter(Output)
    End Function
    Class RuleMetadataComparer
        Implements Collections.Generic.IComparer(Of RuleMetadata)
        Public Function Compare(x As RuleMetadata, y As RuleMetadata) As Integer Implements Generic.IComparer(Of RuleMetadata).Compare
            If x.Index = y.Index Then
                If x.Length = y.Length Then
                    Return y.OrigOrder.CompareTo(x.OrigOrder)
                Else
                    Return y.Length.CompareTo(x.Length)
                End If
            Else
                Return y.Index.CompareTo(x.Index)
            End If
        End Function
    End Class
    Public Shared Function ApplyContigColorRules(ByVal ArabicString As String, PreString As String, PostString As String, PreStop As Boolean, PostStop As Boolean, OptionalStops As Integer(), BreakWords As Boolean) As RenderArray.RenderText()()
        Dim RendererList As New List(Of RenderArray.RenderText())
        RendererList.AddRange(ApplyColorRules(JoinContig(ArabicString, PreString, PostString, PreStop, PostStop, OptionalStops), OptionalStops, True))
        Dim Index As Integer = ArabicString.IndexOf(ArabicData.ArabicEndOfAyah)
        If PreString <> String.Empty AndAlso Index <> -1 Then
            RendererList.RemoveAt(0)
        End If
        Index = ArabicString.LastIndexOf(ArabicData.ArabicEndOfAyah)
        If PostString <> String.Empty AndAlso Index <> -1 Then
            RendererList.RemoveAt(RendererList.Count - 1)
        End If
        If Not BreakWords Then
            Dim RenderTexts As New List(Of RenderArray.RenderText)
            For Count As Integer = 0 To RendererList.Count - 1
                RenderTexts.AddRange(RendererList(Count))
            Next
            RendererList.Clear()
            RendererList.Add(RenderTexts.ToArray())
        End If
        Return RendererList.ToArray()
    End Function
    Public Shared Function ApplyColorRules(ByVal ArabicString As String, OptionalStops As Integer(), BreakWords As Boolean) As RenderArray.RenderText()()
        Dim Count As Integer
        Dim Index As Integer
        Dim MetadataList As New Generic.List(Of RuleMetadata)
        Dim Strings As New Generic.List(Of RenderArray.RenderText)
        For Count = 0 To CachedData.RulesOfRecitationRegEx.Length - 1
            If Not CachedData.RulesOfRecitationRegEx(Count).Evaluator Is Nothing Then
                Dim Matches As System.Text.RegularExpressions.MatchCollection = System.Text.RegularExpressions.Regex.Matches(ArabicString, CachedData.RulesOfRecitationRegEx(Count).Match)
                For MatchIndex As Integer = 0 To Matches.Count - 1
                    Dim SubCount As Integer
                    For SubCount = 0 To CachedData.RulesOfRecitationRegEx(Count).Evaluator.Length - 1
                        If (CachedData.RulesOfRecitationRegEx(Count).Evaluator(SubCount) = "optionalstop" AndAlso (OptionalStops Is Nothing AndAlso Matches(MatchIndex).Groups(SubCount + 1).Value = ArabicData.ArabicSmallHighLigatureSadWithLamWithAlefMaksura OrElse (Not OptionalStops Is Nothing AndAlso Matches(MatchIndex).Groups(SubCount + 1).Value <> String.Empty AndAlso Array.IndexOf(OptionalStops, Matches(MatchIndex).Groups(SubCount + 1).Index) = -1))) OrElse (CachedData.RulesOfRecitationRegEx(Count).Evaluator(SubCount) = "optionalnotstop" AndAlso (OptionalStops Is Nothing AndAlso Matches(MatchIndex).Groups(SubCount + 1).Value <> String.Empty AndAlso Matches(MatchIndex).Groups(SubCount + 1).Value <> ArabicData.ArabicSmallHighLigatureSadWithLamWithAlefMaksura OrElse (Not OptionalStops Is Nothing AndAlso Matches(MatchIndex).Groups(SubCount + 1).Value <> String.Empty AndAlso Array.IndexOf(OptionalStops, Matches(MatchIndex).Groups(SubCount + 1).Index) <> -1))) Then Exit For
                    Next
                    If SubCount <> CachedData.RulesOfRecitationRegEx(Count).Evaluator.Length Then Continue For
                    For SubCount = 0 To CachedData.RulesOfRecitationRegEx(Count).Evaluator.Length - 1
                        If Not CachedData.RulesOfRecitationRegEx(Count).Evaluator(SubCount) Is Nothing Then
                            MetadataList.Add(New RuleMetadata(Matches(MatchIndex).Groups(SubCount + 1).Index, Matches(MatchIndex).Groups(SubCount + 1).Length, CachedData.RulesOfRecitationRegEx(Count).Evaluator(SubCount), SubCount))
                        End If
                    Next
                Next
            End If
        Next
        MetadataList.Sort(New RuleMetadataComparer)
        Dim RuleIndexes As New List(Of Integer)
        For Count = 0 To ArabicString.Length - 1
            RuleIndexes.Add(0)
        Next
        For Index = 0 To MetadataList.Count - 1
            For Count = 0 To CachedData.IslamData.ColorRules.Length - 1
                Dim Match As Integer = Array.FindIndex(CachedData.IslamData.ColorRules(Count).Match.Split("|"c), Function(Str As String) Array.IndexOf(Array.ConvertAll(MetadataList(Index).Type.Split("|"c), Function(S As String) System.Text.RegularExpressions.Regex.Replace(S, "\(.*\)", String.Empty)), Str) <> -1)
                If Match <> -1 Then
                    For MetaCount As Integer = MetadataList(Index).Index To MetadataList(Index).Index + MetadataList(Index).Length - 1
                        RuleIndexes(MetaCount) = Count
                    Next
                    'ApplyColorRules(Strings(Strings.Count - 1).Text)
                End If
            Next
        Next
        Dim Base As Integer = 0
        Dim WordRenderers As New List(Of RenderArray.RenderText())
        Dim Renderers As New List(Of RenderArray.RenderText)
        For Count = 0 To RuleIndexes.Count - 1
            If Count = RuleIndexes.Count - 1 Or (ArabicString(Count) = " "c And BreakWords) Then
                Renderers.Add(New RenderArray.RenderText(RenderArray.RenderDisplayClass.eArabic, If((ArabicString(Count) = " "c And BreakWords), ArabicString.Substring(Base, Count - Base), ArabicString.Substring(Base))) With {.Clr = CachedData.IslamData.ColorRules(RuleIndexes(Count) Mod CachedData.IslamData.ColorRules.Length).Color})
                WordRenderers.Add(Renderers.ToArray())
                Renderers = New List(Of RenderArray.RenderText)
                Base = Count + 1
            ElseIf RuleIndexes(Count) <> RuleIndexes(Count + 1) Then
                Renderers.Add(New RenderArray.RenderText(RenderArray.RenderDisplayClass.eArabic, ArabicString.Substring(Base, Count - Base + 1)) With {.Clr = CachedData.IslamData.ColorRules(RuleIndexes(Count) Mod CachedData.IslamData.ColorRules.Length).Color})
                Base = Count + 1
            End If
        Next
        Return WordRenderers.ToArray()
    End Function
    Public Shared Function ArabicWordForLessThanThousand(ByVal Number As Integer, UseClassic As Boolean, UseAlefHundred As Boolean) As String
        Dim Str As String = String.Empty
        Dim HundStr As String = String.Empty
        If Number >= 100 Then
            HundStr = If(UseAlefHundred, CachedData.ArabicBaseHundredNumbers((Number \ 100) - 1).Insert(2, "A"), CachedData.ArabicBaseHundredNumbers((Number \ 100) - 1))
            If (Number Mod 100) = 0 Then Return HundStr
            Number = Number Mod 100
        End If
        If (Number Mod 10) <> 0 And Number <> 11 And Number <> 12 Then
            Str = CachedData.ArabicBaseNumbers((Number Mod 10) - 1)
        End If
        If Number >= 11 AndAlso Number < 20 Then
            If Number = 11 Or Number = 12 Then
                Str += CachedData.ArabicBaseExtraNumbers(Number - 11)
            Else
                Str = Str.Remove(Str.Length - 1) + "a"
            End If
            Str += " " + CachedData.ArabicBaseTenNumbers(1).Remove(CachedData.ArabicBaseTenNumbers(1).Length - 2)
        ElseIf (Number = 0 And Str = String.Empty) Or Number = 10 Or Number >= 20 Then
            Str = If(Str = String.Empty, String.Empty, Str + " " + CachedData.ArabicCombiners(0)) + CachedData.ArabicBaseTenNumbers(Number \ 10)
        End If
        Return If(UseClassic, If(Str = String.Empty, String.Empty, Str + If(HundStr = String.Empty, String.Empty, " " + CachedData.ArabicCombiners(0))) + HundStr, If(HundStr = String.Empty, String.Empty, HundStr + If(Str = String.Empty, String.Empty, " " + CachedData.ArabicCombiners(0))) + Str)
    End Function
    Public Shared Function ArabicWordFromNumber(ByVal Number As Long, UseClassic As Boolean, UseAlefHundred As Boolean, UseMilliard As Boolean) As String
        Dim Str As String = String.Empty
        Dim NextStr As String = String.Empty
        Dim CurBase As Integer = 3
        Dim BaseNums As Long() = {1000, 1000000, 1000000000, 1000000000000}
        Dim Bases As String()() = {CachedData.ArabicBaseThousandNumbers, CachedData.ArabicBaseMillionNumbers, If(UseMilliard, CachedData.ArabicBaseMilliardNumbers, CachedData.ArabicBaseBillionNumbers), CachedData.ArabicBaseTrillionNumbers}
        Do
            If Number >= BaseNums(CurBase) And Number < 2 * BaseNums(CurBase) Then
                NextStr = Bases(CurBase)(0)
            ElseIf Number >= 2 * BaseNums(CurBase) And Number < 3 * BaseNums(CurBase) Then
                NextStr = Bases(CurBase)(1)
            ElseIf Number >= 3 * BaseNums(CurBase) And Number < 10 * BaseNums(CurBase) Then
                NextStr = CachedData.ArabicBaseNumbers(CInt(Number \ BaseNums(CurBase) - 1)).Remove(CachedData.ArabicBaseNumbers(CInt(Number \ BaseNums(CurBase) - 1)).Length - 1) + "u " + Bases(CurBase)(2)
            ElseIf Number >= 10 * BaseNums(CurBase) And Number < 11 * BaseNums(CurBase) Then
                NextStr = CachedData.ArabicBaseTenNumbers(1).Remove(CachedData.ArabicBaseTenNumbers(1).Length - 1) + "u " + Bases(CurBase)(2)
            ElseIf Number >= BaseNums(CurBase) Then
                NextStr = Arabic.ArabicWordForLessThanThousand(CInt((Number \ BaseNums(CurBase)) Mod 100), UseClassic, UseAlefHundred)
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
            Str = If(UseClassic, If(NextStr = String.Empty, String.Empty, NextStr + If(Str = String.Empty, String.Empty, " " + CachedData.ArabicCombiners(0))) + Str, If(Str = String.Empty, String.Empty, Str + If(NextStr = String.Empty, String.Empty, " " + CachedData.ArabicCombiners(0))) + NextStr)
            NextStr = String.Empty
        Loop While CurBase >= 0
        If Number <> 0 Or Str = String.Empty Then NextStr = Arabic.ArabicWordForLessThanThousand(CInt(Number), UseClassic, UseAlefHundred)
        Return If(UseClassic, If(NextStr = String.Empty, String.Empty, NextStr + If(Str = String.Empty, String.Empty, " " + CachedData.ArabicCombiners(0))) + Str, If(Str = String.Empty, String.Empty, Str + If(NextStr = String.Empty, String.Empty, " " + CachedData.ArabicCombiners(0))) + NextStr)
    End Function
    Public Shared Function NegativeMatchEliminator(NegativeMatch As String, Evaluator As String) As System.Text.RegularExpressions.MatchEvaluator
        Return Function(Match As System.Text.RegularExpressions.Match)
                   Return If(NegativeMatch <> String.Empty AndAlso Match.Result(NegativeMatch) <> String.Empty, Match.Value, Match.Result(Evaluator))
               End Function
    End Function
    Public Shared Function ProcessTransform(ArabicString As String, Rules As IslamData.RuleTranslationCategory.RuleTranslation(), bPriority As Boolean) As String
        'mutual exclusivity required and makes the rules far more accurate and self-documenting and explanatory
        Dim Replacements As New List(Of RuleMetadata)
        For Count = 0 To Rules.Length - 1
            If bPriority Then
                ArabicString = System.Text.RegularExpressions.Regex.Replace(ArabicString, Rules(Count).Match, NegativeMatchEliminator(Rules(Count).NegativeMatch, Rules(Count).Evaluator))
            Else
                Dim Matches As System.Text.RegularExpressions.MatchCollection = System.Text.RegularExpressions.Regex.Matches(ArabicString, Rules(Count).Match)
                For MatchCount = 0 To Matches.Count - 1
                    If Rules(Count).NegativeMatch = String.Empty OrElse Matches(MatchCount).Result(Rules(Count).NegativeMatch) = String.Empty Then
                        Dim DupCount As Integer
                        For DupCount = 0 To Matches(MatchCount).Result(Rules(Count).Evaluator).Length - 1
                            If Matches(MatchCount).Index + DupCount >= ArabicString.Length OrElse ArabicString(Matches(MatchCount).Index + DupCount) <> Matches(MatchCount).Result(Rules(Count).Evaluator)(DupCount) Then
                                Exit For
                            End If
                        Next
                        Replacements.Add(New RuleMetadata(Matches(MatchCount).Index + DupCount, Matches(MatchCount).Length - DupCount, Matches(MatchCount).Result(Rules(Count).Evaluator).Substring(DupCount), Count))
                    End If
                Next
            End If
        Next
        Replacements.Sort(New RuleMetadataComparer)
        For Count = 0 To Replacements.Count - 1
            If Count <> 0 AndAlso (Replacements(Count).Index + Replacements(Count).Length > Replacements(Count - 1).Index) Then
                Debug.Print(Rules(Replacements(Count - 1).OrigOrder).Name + ":" + Arabic.TransliterateToScheme(Replacements(Count - 1).Type, ArabicData.TranslitScheme.Literal, String.Empty) + "-" + Rules(Replacements(Count).OrigOrder).Name + ":" + Arabic.TransliterateToScheme(Replacements(Count).Type, ArabicData.TranslitScheme.Literal, String.Empty) + "-" + Arabic.TransliterateToScheme(ArabicString.Substring(Math.Max(Replacements(Count).Index - 15, 0), Math.Min(Replacements(Count - 1).Index + Replacements(Count - 1).Length + 15, ArabicString.Length) - Math.Max(Replacements(Count).Index - 15, 0)), ArabicData.TranslitScheme.Literal, String.Empty))
            End If
            ArabicString = ArabicString.Substring(0, Replacements(Count).Index) + Replacements(Count).Type + ArabicString.Substring(Replacements(Count).Index + Replacements(Count).Length)
        Next
        Return ArabicString
    End Function
    Public Shared Function ChangeBaseScript(ArabicString As String, BaseText As TanzilReader.QuranTexts, ByVal PreString As String, ByVal PostString As String) As String
        If BaseText = TanzilReader.QuranTexts.Warsh Then
            ArabicString = UnjoinContig(ProcessTransform(JoinContig(ArabicString, PreString, PostString, False, False, Nothing), CachedData.WarshScript, True), PreString, PostString)
        End If
        Return ArabicString
    End Function
    Public Shared Function ChangeScript(ArabicString As String, ScriptType As TanzilReader.QuranScripts, ByVal PreString As String, ByVal PostString As String) As String
        If ScriptType = TanzilReader.QuranScripts.UthmaniMin Then
            ArabicString = UnjoinContig(ProcessTransform(JoinContig(ArabicString, PreString, PostString, False, False, Nothing), CachedData.UthmaniMinimalScript, False), PreString, PostString)
        ElseIf ScriptType = TanzilReader.QuranScripts.SimpleEnhanced Then
            Dim ScriptCombine As New List(Of IslamMetadata.IslamData.RuleTranslationCategory.RuleTranslation)
            ScriptCombine.AddRange(CachedData.SimpleScriptBase)
            ScriptCombine.AddRange(CachedData.SimpleEnhancedScript)
            ArabicString = UnjoinContig(ProcessTransform(JoinContig(ArabicString, PreString, PostString, False, False, Nothing), ScriptCombine.ToArray(), False), PreString, PostString)
        ElseIf ScriptType = TanzilReader.QuranScripts.Simple Then
            Dim ScriptCombine As New List(Of IslamMetadata.IslamData.RuleTranslationCategory.RuleTranslation)
            ScriptCombine.AddRange(CachedData.SimpleScriptBase)
            ScriptCombine.AddRange(CachedData.SimpleScript)
            ArabicString = UnjoinContig(ProcessTransform(JoinContig(ArabicString, PreString, PostString, False, False, Nothing), ScriptCombine.ToArray(), False), PreString, PostString)
        ElseIf ScriptType = TanzilReader.QuranScripts.SimpleClean Then
            Dim ScriptCombine As New List(Of IslamMetadata.IslamData.RuleTranslationCategory.RuleTranslation)
            ScriptCombine.AddRange(CachedData.SimpleScriptBase)
            ScriptCombine.AddRange(CachedData.SimpleCleanScript)
            ArabicString = UnjoinContig(ProcessTransform(JoinContig(ArabicString, PreString, PostString, False, False, Nothing), ScriptCombine.ToArray(), False), PreString, PostString)
        ElseIf ScriptType = TanzilReader.QuranScripts.SimpleMin Then
            Dim ScriptCombine As New List(Of IslamMetadata.IslamData.RuleTranslationCategory.RuleTranslation)
            ScriptCombine.AddRange(CachedData.SimpleScriptBase)
            ScriptCombine.AddRange(CachedData.SimpleMinimalScript)
            ArabicString = UnjoinContig(ProcessTransform(JoinContig(ArabicString, PreString, PostString, False, False, Nothing), ScriptCombine.ToArray(), False), PreString, PostString)
        End If
        Return ArabicString
    End Function
    Public Shared Function ReplaceMetadata(ArabicString As String, MetadataRule As RuleMetadata, Scheme As String, LearningMode As Boolean) As String
        For Count As Integer = 0 To CachedData.ColoringSpelledOutRules.Length - 1
            Dim Match As String = Array.Find(CachedData.ColoringSpelledOutRules(Count).Match.Split("|"c), Function(Str As String) Array.IndexOf(Array.ConvertAll(MetadataRule.Type.Split("|"c), Function(S As String) System.Text.RegularExpressions.Regex.Replace(S, "\(.*\)", String.Empty)), Str) <> -1)
            If Match <> Nothing Then
                Dim Str As String = String.Format(CachedData.ColoringSpelledOutRules(Count).Evaluator, ArabicString.Substring(MetadataRule.Index, MetadataRule.Length))
                If CachedData.ColoringSpelledOutRules(Count).RuleFunc <> RuleFuncs.eNone Then
                    Dim Args As String() = RuleFunctions(CachedData.ColoringSpelledOutRules(Count).RuleFunc - 1)(Str, Scheme, LearningMode)
                    If Args.Length = 1 Then
                        Str = Args(0)
                    Else
                        Dim MetaArgs As String() = System.Text.RegularExpressions.Regex.Match(MetadataRule.Type, Match + "\((.*)\)").Groups(1).Value.Split(","c)
                        Str = String.Empty
                        For Index As Integer = 0 To Args.Length - 1
                            If Not Args(Index) Is Nothing And (LearningMode Or CachedData.ColoringSpelledOutRules(Count).RuleFunc <> RuleFuncs.eLearningMode Or Index <> 0) Then
                                Str += ReplaceMetadata(Args(Index), New RuleMetadata(0, Args(Index).Length, MetaArgs(Index).Replace(" "c, "|"c), Index), Scheme, LearningMode)
                            End If
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
        For Count = 0 To CachedData.ErrorCheckRules.Length - 1
            Dim Matches As System.Text.RegularExpressions.MatchCollection = System.Text.RegularExpressions.Regex.Matches(ArabicString, CachedData.ErrorCheckRules(Count).Match)
            For MatchIndex As Integer = 0 To Matches.Count - 1
                If CachedData.ErrorCheckRules(Count).NegativeMatch Is Nothing OrElse Matches(MatchIndex).Result(CachedData.ErrorCheckRules(Count).NegativeMatch) = String.Empty Then
                    'Debug.Print(ErrorCheckRules(Count).Rule + ": " + TransliterateToScheme(ArabicString, ArabicData.TranslitScheme.Literal, String.Empty).Insert(Matches(MatchIndex).Index, "<!-- -->"))
                End If
            Next
        Next
    End Sub
    Public Shared Function JoinContig(ByVal ArabicString As String, ByVal PreString As String, ByVal PostString As String, PreStop As Boolean, PostStop As Boolean, ByRef OptionalStops As Integer()) As String
        Dim Index As Integer = PreString.LastIndexOf(" "c)
        'take last word of pre string and first word of post string or another if it is a pause marker
        'end of ayah sign without number is used as a proper place holder
        If Index <> -1 And PreString.Length - 2 = Index Then Index = PreString.LastIndexOf(" "c, Index - 1)
        If Index <> -1 Then PreString = PreString.Substring(Index + 1)
        If PreString <> String.Empty Then PreString += " " + ArabicData.ArabicEndOfAyah + " "
        Index = PostString.IndexOf(" "c)
        If Index = 2 Then Index = PreString.IndexOf(" "c, Index + 1)
        If Index <> -1 Then PostString = PostString.Substring(0, Index)
        If PostString <> String.Empty Then PostString = " " + ArabicData.ArabicEndOfAyah + " " + PostString
        If Not OptionalStops Is Nothing Then
            Dim Stops As New List(Of Integer)
            If PreString <> String.Empty Then Stops.Add(PreString.Length - 2)
            For Count As Integer = 0 To OptionalStops.Length - 1
                Stops.Add(OptionalStops(Count) + PreString.Length)
            Next
            If PostString <> String.Empty Then Stops.Add(PreString.Length + ArabicString.Length + 1)
        End If
        Return PreString + ArabicString + PostString
    End Function
    Public Shared Function UnjoinContig(ByVal ArabicString As String, ByVal PreString As String, ByVal PostString As String) As String
        Dim Index As Integer = ArabicString.IndexOf(ArabicData.ArabicEndOfAyah)
        If PreString <> String.Empty AndAlso Index <> -1 Then
            ArabicString = ArabicString.Substring(Index + 1 + 1)
        End If
        Index = ArabicString.LastIndexOf(ArabicData.ArabicEndOfAyah)
        If PostString <> String.Empty AndAlso Index <> -1 Then
            ArabicString = ArabicString.Substring(0, Index - 1)
        End If
        Return ArabicString
    End Function
    Public Shared Function TransliterateContigWithRules(ByVal ArabicString As String, ByVal PreString As String, ByVal PostString As String, PreStop As Boolean, PostStop As Boolean, Scheme As String, OptionalStops As Integer()) As String
        Return UnjoinContig(TransliterateWithRules(JoinContig(ArabicString, PreString, PostString, PreStop, PostStop, OptionalStops), Scheme, OptionalStops, False), PreString, PostString)
    End Function
    Public Shared Function TransliterateWithRules(ByVal ArabicString As String, Scheme As String, OptionalStops As Integer(), LearningMode As Boolean) As String
        Dim Count As Integer
        Dim MetadataList As New Generic.List(Of RuleMetadata)
        DoErrorCheck(ArabicString)
        For Count = 0 To CachedData.RulesOfRecitationRegEx.Length - 1
            If Not CachedData.RulesOfRecitationRegEx(Count).Evaluator Is Nothing Then
                Dim Matches As System.Text.RegularExpressions.MatchCollection = System.Text.RegularExpressions.Regex.Matches(ArabicString, CachedData.RulesOfRecitationRegEx(Count).Match)
                For MatchIndex As Integer = 0 To Matches.Count - 1
                    Dim SubCount As Integer
                    For SubCount = 0 To CachedData.RulesOfRecitationRegEx(Count).Evaluator.Length - 1
                        If (CachedData.RulesOfRecitationRegEx(Count).Evaluator(SubCount) = "optionalstop" AndAlso (OptionalStops Is Nothing AndAlso Matches(MatchIndex).Groups(SubCount + 1).Value = ArabicData.ArabicSmallHighLigatureSadWithLamWithAlefMaksura OrElse (Not OptionalStops Is Nothing AndAlso Matches(MatchIndex).Groups(SubCount + 1).Value <> String.Empty AndAlso Array.IndexOf(OptionalStops, Matches(MatchIndex).Groups(SubCount + 1).Index) = -1))) OrElse (CachedData.RulesOfRecitationRegEx(Count).Evaluator(SubCount) = "optionalnotstop" AndAlso (OptionalStops Is Nothing AndAlso Matches(MatchIndex).Groups(SubCount + 1).Value <> String.Empty AndAlso Matches(MatchIndex).Groups(SubCount + 1).Value <> ArabicData.ArabicSmallHighLigatureSadWithLamWithAlefMaksura OrElse (Matches(MatchIndex).Groups(SubCount + 1).Length = 0 AndAlso (Matches(MatchIndex).Groups(SubCount + 1).Index = 0 Or Matches(MatchIndex).Groups(SubCount + 1).Index = ArabicString.Length)) OrElse (Not OptionalStops Is Nothing AndAlso Matches(MatchIndex).Groups(SubCount + 1).Value <> String.Empty AndAlso Array.IndexOf(OptionalStops, Matches(MatchIndex).Groups(SubCount + 1).Index) <> -1))) Then Exit For
                    Next
                    If SubCount <> CachedData.RulesOfRecitationRegEx(Count).Evaluator.Length Then Continue For
                    For SubCount = 0 To CachedData.RulesOfRecitationRegEx(Count).Evaluator.Length - 1
                        If Not CachedData.RulesOfRecitationRegEx(Count).Evaluator(SubCount) Is Nothing And (Matches(MatchIndex).Groups(SubCount + 1).Length <> 0 Or Array.IndexOf(AllowZeroLength, CachedData.RulesOfRecitationRegEx(Count).Evaluator(SubCount)) <> -1) Then
                            MetadataList.Add(New RuleMetadata(Matches(MatchIndex).Groups(SubCount + 1).Index, Matches(MatchIndex).Groups(SubCount + 1).Length, CachedData.RulesOfRecitationRegEx(Count).Evaluator(SubCount), SubCount))
                            'Debug.Print(CachedData.RulesOfRecitationRegEx(Count).Name + " Index: " + CStr(Matches(MatchIndex).Groups(SubCount + 1).Index) + " Length: " + CStr(Matches(MatchIndex).Groups(SubCount + 1).Length) + " Ruling: " + CachedData.RulesOfRecitationRegEx(Count).Evaluator(SubCount))
                        End If
                    Next
                Next
            End If
        Next
        MetadataList.Sort(New RuleMetadataComparer)
        Dim Index As Integer
        For Index = 0 To MetadataList.Count - 1
            ArabicString = ReplaceMetadata(ArabicString, MetadataList(Index), Scheme, LearningMode)
        Next
        'redundant romanization rules should have -'s such as seen/teh/kaf-heh
        For Count = 0 To CachedData.RomanizationRules.Length - 1
            If CachedData.RomanizationRules(Count).RuleFunc = RuleFuncs.eNone Then
                ArabicString = System.Text.RegularExpressions.Regex.Replace(ArabicString, CachedData.RomanizationRules(Count).Match, CachedData.RomanizationRules(Count).Evaluator)
            Else
                ArabicString = System.Text.RegularExpressions.Regex.Replace(ArabicString, CachedData.RomanizationRules(Count).Match, Function(Match As System.Text.RegularExpressions.Match) RuleFunctions(CachedData.RomanizationRules(Count).RuleFunc - 1)(Match.Result(CachedData.RomanizationRules(Count).Evaluator), Scheme, LearningMode)(0))
            End If
        Next

        'process wasl loanwords and names
        'process loanwords and names
        If System.Text.RegularExpressions.Regex.Match(ArabicString, "[\p{IsArabic}|\p{IsArabicPresentationForms-A}|\p{IsArabicPresentationForms-B}]").Success Then Debug.Print(ArabicString)
        Return ArabicString
    End Function
    Public Shared Function GetTransliterationSchemeTable(ByVal Item As PageLoader.TextItem) As RenderArray
        Dim Scheme As String = If(CInt(HttpContext.Current.Request.Params("translitscheme")) >= 2, CachedData.IslamData.TranslitSchemes(CInt(HttpContext.Current.Request.Params("translitscheme")) \ 2).Name, String.Empty)
        Return New RenderArray("translitscheme") With {.Items = GetTransliterationTable(Scheme)}
    End Function
    Shared Function SubOutPatterns(Str As String) As String
        Return Str.Replace(CachedData.TehMarbutaStopRule, String.Empty).Replace(CachedData.TehMarbutaContinueRule, "...").Replace(CachedData.TranslateRegEx("(?=(?:{ArabicMoonLetters}|{ArabicSunLettersNoLam})(?:{ArabicSukun}|{ArabicShadda}?(?:{ArabicFathaDammaKasra}))?(?:{ArabicLetters})(?:{ArabicFatha}|{ArabicKasra})|{ArabicWaslKasraExceptions})", True), ArabicData.ArabicKasra).Replace(CachedData.TranslateRegEx("(?=(?:{ArabicMoonLetters}|{ArabicSunLettersNoLam})(?:{ArabicSukun}|{ArabicShadda}?(?:{ArabicFathaDammaKasra}))?(?:{ArabicLetters}){ArabicDamma})", True), ArabicData.ArabicDamma).Replace(CachedData.TranslateRegEx("({ArabicSunLetters}|{ArabicMoonLettersNoVowels}|{ArabicLetterWaw}|{ArabicLetterYeh}|{ArabicLetterAlefMaksura}|{ArabicSmallYeh})", True), ChrW(&H66D))
    End Function
    Shared Function GetTransliterationTable(Scheme As String) As List(Of RenderArray.RenderItem)
        Dim Items As New List(Of RenderArray.RenderItem)
        Items.AddRange(Array.ConvertAll(CachedData.ArabicLettersInOrder, Function(Letter As String) New RenderArray.RenderItem(RenderArray.RenderTypes.eText, New RenderArray.RenderText() {New RenderArray.RenderText(RenderArray.RenderDisplayClass.eArabic, Letter), New RenderArray.RenderText(RenderArray.RenderDisplayClass.eTransliteration, GetSchemeValueFromSymbol(ArabicData.ArabicLetters(ArabicData.FindLetterBySymbol(Letter.Chars(0))), Scheme))})))
        Items.AddRange(Array.ConvertAll(CachedData.ArabicSpecialLetters, Function(Combo As String) New RenderArray.RenderItem(RenderArray.RenderTypes.eText, New RenderArray.RenderText() {New RenderArray.RenderText(RenderArray.RenderDisplayClass.eArabic, System.Text.RegularExpressions.Regex.Replace(SubOutPatterns(Combo), "\(?\\u([0-9a-fA-F]{4})\)?", Function(Match As System.Text.RegularExpressions.Match) ChrW(Integer.Parse(Match.Groups(1).Value, Globalization.NumberStyles.HexNumber)))), New RenderArray.RenderText(RenderArray.RenderDisplayClass.eTransliteration, GetSchemeSpecialValue(Combo, GetSchemeSpecialFromMatch(Combo, False), Scheme))})))
        Items.AddRange(Array.ConvertAll(CachedData.ArabicHamzas, Function(Letter As String) New RenderArray.RenderItem(RenderArray.RenderTypes.eText, New RenderArray.RenderText() {New RenderArray.RenderText(RenderArray.RenderDisplayClass.eArabic, Letter), New RenderArray.RenderText(RenderArray.RenderDisplayClass.eTransliteration, GetSchemeValueFromSymbol(ArabicData.ArabicLetters(ArabicData.FindLetterBySymbol(Letter.Chars(0))), Scheme))})))
        Items.AddRange(Array.ConvertAll(CachedData.ArabicVowels, Function(Letter As String) New RenderArray.RenderItem(RenderArray.RenderTypes.eText, New RenderArray.RenderText() {New RenderArray.RenderText(RenderArray.RenderDisplayClass.eArabic, Letter), New RenderArray.RenderText(RenderArray.RenderDisplayClass.eTransliteration, GetSchemeValueFromSymbol(ArabicData.ArabicLetters(ArabicData.FindLetterBySymbol(Letter.Chars(0))), Scheme))})))
        Items.AddRange(Array.ConvertAll(CachedData.ArabicMultis, Function(Combo As String) New RenderArray.RenderItem(RenderArray.RenderTypes.eText, New RenderArray.RenderText() {New RenderArray.RenderText(RenderArray.RenderDisplayClass.eArabic, Combo), New RenderArray.RenderText(RenderArray.RenderDisplayClass.eTransliteration, GetSchemeLongVowelFromString(Combo, Scheme))})))
        Items.AddRange(Array.ConvertAll(CachedData.ArabicTajweed, Function(Letter As String) New RenderArray.RenderItem(RenderArray.RenderTypes.eText, New RenderArray.RenderText() {New RenderArray.RenderText(RenderArray.RenderDisplayClass.eArabic, Letter), New RenderArray.RenderText(RenderArray.RenderDisplayClass.eTransliteration, GetSchemeValueFromSymbol(ArabicData.ArabicLetters(ArabicData.FindLetterBySymbol(Letter.Chars(0))), Scheme))})))
        Return Items
    End Function
    Public Shared Function ArabicTranslitLetters() As String()
        Dim Lets As New List(Of String)
        Lets.AddRange(Array.ConvertAll(CachedData.ArabicLettersInOrder, Function(Ch As String) Ch))
        Lets.AddRange(Array.ConvertAll(CachedData.ArabicHamzas, Function(Ch As String) Ch))
        Lets.AddRange(CachedData.ArabicVowels)
        Lets.AddRange(Array.ConvertAll(CachedData.ArabicTajweed, Function(Ch As String) Ch))
        Lets.AddRange(Array.ConvertAll(CachedData.ArabicSilent, Function(Ch As String) Ch))
        Lets.AddRange(Array.ConvertAll(CachedData.ArabicPunctuation, Function(Ch As String) Ch))
        Lets.AddRange(Array.ConvertAll(CachedData.ArabicNums, Function(Ch As String) Ch))
        Lets.AddRange(Array.ConvertAll(CachedData.NonArabicLetters, Function(Ch As String) Ch))
        Return Lets.ToArray()
    End Function
    Public Shared Function GetTranslitSchemeMetadata(ID As String) As Array()
        Dim Output(CachedData.IslamData.TranslitSchemes.Length + 2) As Array
        Output(0) = New String() {}
        Output(1) = New String() {String.Empty, String.Empty}
        Output(2) = New String() {Utility.LoadResourceString("IslamInfo_Name"), Utility.LoadResourceString("IslamInfo_Translation")}
        For Count = 0 To CachedData.IslamData.TranslitSchemes.Length - 1
            Output(3 + Count) = {CachedData.IslamData.TranslitSchemes(Count).Name, Utility.LoadResourceString("IslamSource_" + CachedData.IslamData.TranslitSchemes(Count).Name)}
        Next
        Return RenderArray.MakeTableJSFunctions(Output, ID)
    End Function
    Shared Function GetTranslitSchemeJSArray() As String
        'Dim Letters(ArabicData.ArabicLetters.Length - 1) As IslamData.ArabicSymbol
        'ArabicData.ArabicLetters.CopyTo(Letters, 0)
        'Array.Sort(Letters, New StringLengthComparer("RomanTranslit"))
        Return "var translitSchemes = " + Utility.MakeJSIndexedObject(Array.ConvertAll(CachedData.IslamData.TranslitSchemes, Function(TranslitScheme As IslamData.TranslitScheme) CStr(Array.IndexOf(CachedData.IslamData.TranslitSchemes, TranslitScheme) + 1)), _
                                                                          New Array() {Array.ConvertAll(Of IslamData.TranslitScheme, String)(CachedData.IslamData.TranslitSchemes, Function(TranslitScheme As IslamData.TranslitScheme) Utility.MakeJSIndexedObject({"standard", "multi", "special", "gutteral"}, New Array() {New String() {Utility.MakeJSIndexedObject(Array.ConvertAll(ArabicTranslitLetters(), Function(Str As String) System.Text.RegularExpressions.Regex.Replace(Str, "\(?\\u([0-9a-fA-F]{4})\)?", Function(Match As System.Text.RegularExpressions.Match) ChrW(Integer.Parse(Match.Groups(1).Value, Globalization.NumberStyles.HexNumber)))), New Array() {Array.ConvertAll(Of String, String)(ArabicTranslitLetters(),
                                                                            Function(Str As String) GetSchemeValueFromSymbol(ArabicData.ArabicLetters(ArabicData.FindLetterBySymbol(Str.Chars(0))), TranslitScheme.Name))}, False), Utility.MakeJSArray(TranslitScheme.Multis), Utility.MakeJSArray(TranslitScheme.SpecialLetters), Utility.MakeJSArray(TranslitScheme.Gutterals)}}, True))}, True)

    End Function
    Shared Function GetArabicSymbolJSArray() As String
        GetArabicSymbolJSArray = "var arabicLetters = " + _
                                Utility.MakeJSArray(New String() {Utility.MakeJSIndexedObject(New String() {"Symbol", "Shaping"}, _
                                Array.ConvertAll(Of ArabicData.ArabicSymbol, String())(Array.FindAll(ArabicData.ArabicLetters, Function(Letter As ArabicData.ArabicSymbol) GetSchemeValueFromSymbol(Letter, "ExtendedBuckwalter") <> String.Empty), Function(Convert As ArabicData.ArabicSymbol) New String() {CStr(AscW(Convert.Symbol)), If(Convert.Shaping = Nothing, String.Empty, Utility.MakeJSArray(Array.ConvertAll(Convert.Shaping, Function(Ch As Char) CStr(AscW(Ch)))))}), False)}, True) + ";"
    End Function
    Public Shared FindLetterBySymbolJS As String = "function findLetterBySymbol(chVal) { var iSubCount; for (iSubCount = 0; iSubCount < arabicLetters.length; iSubCount++) { if (chVal === parseInt(arabicLetters[iSubCount].Symbol, 10)) return iSubCount; for (var iShapeCount = 0; iShapeCount < arabicLetters[iSubCount].Shaping.length; iShapeCount++) { if (chVal === parseInt(arabicLetters[iSubCount].Shaping[iShapeCount], 10)) return iSubCount; } } return -1; }"
    Public Shared TransliterateGenJS As String() = {
        FindLetterBySymbolJS,
        "function isLetterDiacritic(index) { return (" + String.Join("||", Array.ConvertAll(CachedData.RecitationLettersDiacritics, Function(C As String) "parseInt(arabicLetters[index].Symbol, 10) === 0x" + Hex(AscW(C.Chars(0))))) + "); }", _
        "function isSpecialSymbol(index) { return (" + String.Join("||", Array.ConvertAll(CachedData.RecitationSpecialSymbols, Function(C As String) "parseInt(arabicLetters[index].Symbol, 10) === 0x" + Hex(AscW(C.Chars(0))))) + "); }", _
        "function isCombiningSymbol(index) { return (" + String.Join("||", Array.ConvertAll(CachedData.RecitationCombiningSymbols, Function(C As String) "parseInt(arabicLetters[index].Symbol, 10) === 0x" + Hex(AscW(C.Chars(0))))) + "); }", _
        "var arabicMultis = " + Utility.MakeJSArray(CachedData.ArabicMultis) + ";", _
        "var arabicSpecials = " + Utility.MakeJSArray(CachedData.ArabicSpecialLetters) + ";", _
        "function getSchemeSpecialFromMatch(str, bExp) { if (bExp) { for (var count = 0; count < arabicSpecials.length; count++) { re = new RegExp(arabicSpecials[count], ''); if (re.exec(str) !== null) return count; } } else { return arabicSpecials.indexOf(str); } return -1; }", _
        "function generateDefaultStops(str) { var re = new RegExp('(^\s*|\s+)(" + ArabicData.MakeUniRegEx(ArabicData.ArabicEndOfAyah) + "[\p{Nd}]{1,3}|" + CachedData.OptionalPattern + ")(?=\s*$|\s+)', 'g'); var arr, defstops = [], dottoggle = false; while ((arr = re.exec(str)) !== null) { if (arr[2] === String.fromCharCode(0x6D6) && (arr[2] !== String.fromCharCode(0x6DB) || dottoggle)) defstops.push(arr.index + arr[1].length); if (arr[2] === String.fromCharCode(0x6DB)) dottoggle = !dottoggle; } }", _
        "function doTransliterate(sVal, direction, conversion) { var iCount, iSubCount, sOutVal = ''; if (conversion === 0) return sVal; if (direction && (conversion % 2) === 0) return transliterateWithRules(sVal, Math.floor((conversion - 2) / 2) + 2, generateDefaultStops(sVal), false); for (iCount = 0; iCount < sVal.length; iCount++) { if (sVal.charAt(iCount) === '\\' || direction && (sVal.charAt(iCount) === String.fromCharCode(0x60C) || sVal.charAt(iCount) === String.fromCharCode(0x61B) || sVal.charAt(iCount) === String.fromCharCode(0x61F))) { if (!direction) iCount++; if (sVal.charAt(iCount) === (direction ? String.fromCharCode(0x60C) : ',')) { sOutVal += (direction ? '\\,' : String.fromCharCode(0x60C)); } else if (sVal.charAt(iCount) === (direction ? String.fromCharCode(0x61B) : ';')) { sOutVal += (direction ? '\\;' : String.fromCharCode(0x61B)); } else if (sVal.charAt(iCount) === (direction ? String.fromCharCode(0x61F) : '?')) { sOutVal += (direction ? '\\?' : String.fromCharCode(0x61F)); } else { sOutVal += String.fromCharCode(0x202A) + sVal.charAt(iCount) + String.fromCharCode(0x202C); } } else { if (getSchemeSpecialFromMatch(sVal.slice(iCount), false) !== -1) { sOutVal += translitSchemes[(Math.floor((conversion - 2) / 2) + 2).toString()].special[getSchemeSpecialFromMatch(sVal.slice(iCount), false)]; re = new RegExp(arabicSpecials[getSchemeSpecialFromMatch(sVal.slice(iCount), false)], ''); iCount += re.exec(sVal.slice(iCount))[0].length - 1; } else if (sVal.length - iCount > 1 && arabicMultis.indexOf(sVal.slice(iCount, 2)) !== -1) { sOutVal += translitSchemes[(Math.floor((conversion - 2) / 2) + 2).toString()].multi[arabicMultis.indexOf(sVal.slice(iCount, 2))]; iCount++; } else { for (iSubCount = 0; iSubCount < arabicLetters.length; iSubCount++) { if (direction ? sVal.charCodeAt(iCount) === parseInt(arabicLetters[iSubCount].Symbol, 10) : sVal.charAt(iCount) === unescape(translitSchemes[(Math.floor((conversion - 2) / 2) + 2).toString()].standard[String.fromCharCode(arabicLetters[iSubCount].Symbol)])) { sOutVal += (direction ? (translitSchemes[(Math.floor((conversion - 2) / 2) + 2).toString()].standard[String.fromCharCode(arabicLetters[iSubCount].Symbol)] ? translitSchemes[(Math.floor((conversion - 2) / 2) + 2).toString()].standard[String.fromCharCode(arabicLetters[iSubCount].Symbol)] : '') : (((isCombiningSymbol(iSubCount) && (iSubCount === 0 || findLetterBySymbol(sOutVal.charCodeAt(sOutVal.length - 1)) === -1 || !isLetterDiacritic(findLetterBySymbol(sOutVal.charCodeAt(sOutVal.length - 1))) && !isSpecialSymbol(findLetterBySymbol(sOutVal.charCodeAt(sOutVal.length - 1))))) ? String.fromCharCode(0x202B) : '') + String.fromCharCode(arabicLetters[iSubCount].Symbol) + ((isCombiningSymbol(iSubCount) && (iSubCount === 0 || findLetterBySymbol(sOutVal.charCodeAt(sOutVal.length - 1)) === -1 || !isLetterDiacritic(findLetterBySymbol(sOutVal.charCodeAt(sOutVal.length - 1))) && !isSpecialSymbol(findLetterBySymbol(sOutVal.charCodeAt(sOutVal.length - 1))))) ? String.fromCharCode(0x202C) : ''))); break; } } if (iSubCount === arabicLetters.length && sVal.charCodeAt(iCount) !== 0x200E && sVal.charCodeAt(iCount) !== 0x200F && !IsExplicit(sVal.charCodeAt(iCount))) sOutVal += ((direction && conversion === 1 && sVal.charAt(iCount) !== '\n') ? '\\' : '') + sVal.charAt(iCount); } } } return unescape(sOutVal); }"
    }
    Public Shared IsDiacriticJS As String = "function isDiacritic(index) { return (" + String.Join("||", Array.ConvertAll(CachedData.RecitationDiacritics, Function(C As String) "parseInt(arabicLetters[index].Symbol, 10) === 0x" + Hex(AscW(C.Chars(0))))) + "); }"
    Public Shared DiacriticJS As String() =
        {"function doDiacritics(sVal, direction) { var iCount, sOutVal = ''; for (iCount = 0; iCount < sVal.length; iCount++) { sOutVal += (sVal.charCodeAt(iCount) === 0x671 ? String.fromCharCode(0x627) : ((findLetterBySymbol(sVal.charCodeAt(iCount)) === -1 || !isDiacritic(findLetterBySymbol(sVal.charCodeAt(iCount)))) ? sVal[iCount] : '')); } return sOutVal; }", _
            IsDiacriticJS, _
            FindLetterBySymbolJS}
    Public Shared PlainTransliterateGenJS As String() = {FindLetterBySymbolJS, IsDiacriticJS, _
            "function isWhitespace(index) { return (" + String.Join("||", Array.ConvertAll(CachedData.WhitespaceSymbols, Function(C As String) "parseInt(arabicLetters[index].Symbol, 10) === 0x" + Hex(AscW(C.Chars(0))))) + "); }", _
            "function isPunctuation(index) { return (" + String.Join("||", Array.ConvertAll(CachedData.PunctuationSymbols, Function(C As String) "parseInt(arabicLetters[index].Symbol, 10) === 0x" + Hex(AscW(C.Chars(0))))) + "); }", _
            "function isStop(index) { return (" + String.Join("||", Array.ConvertAll(CachedData.ArabicStopLetters, Function(C As String) "parseInt(arabicLetters[index].Symbol, 10) === 0x" + Hex(AscW(C.Chars(0))))) + "); }", _
            "function applyColorRules(sVal) {}", _
            "function changeScript(sVal, scriptType) {}", _
            "var arabicAlphabet = " + Utility.MakeJSArray(CachedData.ArabicAlphabet) + ";", _
            "var arabicLettersInOrder = " + Utility.MakeJSArray(CachedData.ArabicLettersInOrder) + ";", _
            "var arabicLeadingGutterals = " + Utility.MakeJSArray(CachedData.ArabicLeadingGutterals) + ";", _
            "function getSchemeGutteralFromString(str, scheme, leading) { if (arabicLeadingGutterals.indexOf(str) !== -1) { return translitSchemes[scheme].gutteral[arabicLeadingGutterals.indexOf(str) + (leading ? arabicLeadingGutterals.length : 0)]; } return ''; }", _
            "function arabicLetterSpelling(sVal, bQuranic) { var count, index, output = ''; for (count = 0; count < sVal.length; count++) { index = findLetterBySymbol(sVal.charCodeAt(count)); if (index !== -1 && isLetter(index)) { if (output !== '' && !bQuranic) output += ' '; var idx = arabicLettersInOrder.indexOf(String.fromCharCode(parseInt(arabicLetters[index].Symbol, 10))); output += arabicAlphabet[idx].slice(0, -1) + ((arabicAlphabet[idx][arabicAlphabet[idx].length - 1] == 'n') ? '' : 'o'); } else if (index !== -1 && arabicLetters[index].Symbol === 0x653) { output += sVal.charCodeAt(count); } } return doTransliterate(output, false, 1); }", _
            "function schemeHasValue(str, scheme) { for (var k in translitSchemes[scheme]) { if (translitSchemes[scheme].hasOwnProperty(k) && str === translitSchemes[scheme][k]) return true; } return false; }", _
            "String.prototype.format = function() { var formatted = this; for (var i = 0; i < arguments.length; i++) { formatted = formatted.replace(new RegExp('\\{'+i+'\\}', 'gi'), arguments[i]); } return formatted; };", _
            "RegExp.matchResult = function(subexp, offset, str, matches) { return subexp.replace(/\$(\$|&|`|\'|[0-9]+)/g, function(m, p) { if (p === '$') return '$'; if (p === '`') return str.slice(0, offset); if (p === '\'') return str.slice(offset + matches[0].length); if (p === '&' || parseInt(p, 10) <= 0 || parseInt(p, 10) >= matches.length) return matches[0]; return matches[parseInt(p, 10)]; }); };", _
            "var ruleFunctions = [function(str, scheme, learningMode) { return [str.toUpperCase()]; }, function(str, scheme, learningMode) { return [transliterateWithRules(doTransliterate(arabicWordFromNumber(parseInt(doTransliterate(str, true, 1), 10), true, false, false), false, 1), scheme, null, learningMode)]; }, function(str, scheme, learningMode) { return [transliterateWithRules(arabicLetterSpelling(str, true), scheme, null, learningMode)]; }, function(str, scheme, learningMode) { return [translitSchemes[scheme.toString()].standard[str]]; }, function(str, scheme, learningMode) { return [translitSchemes[scheme.toString()].multi[arabicMultis.indexOf(str)]]; }, function(str, scheme, learningMode) { return [" + Utility.MakeJSArray(CachedData.ArabicFathaDammaKasra) + "[" + Utility.MakeJSArray(CachedData.ArabicTanweens) + ".indexOf(str)], '" + ArabicData.ArabicLetterNoon + "']; }, function (str, scheme, learningMode) { return [getSchemeGutteralFromString(str.slice(0, -1), scheme, true) + str[str.length - 1]]; }, function(str, scheme, learningMode) { return [str[0] + getSchemeGutteralFromString(str.slice(1), scheme, false)]; }, function(str, scheme, learningMode) { return [schemeHasValue(translitSchemes[scheme.toString()].standard[str[0]] + translitSchemes[scheme.toString()].standard[str[1]]) ? str[0] + '-' + str[1] : str]; }, function(str, scheme, learningMode) { return learningMode ? [str, ''] : ['', str]; }];", _
            "function isLetter(index) { return (" + String.Join("||", Array.ConvertAll(CachedData.RecitationLetters, Function(C As String) "parseInt(arabicLetters[index].Symbol, 10) === 0x" + Hex(AscW(C.Chars(0))))) + "); }", _
            "function isSymbol(index) { return (" + String.Join("||", Array.ConvertAll(GetRecitationSymbols(), Function(A As Array) "parseInt(arabicLetters[index].Symbol, 10) === 0x" + Hex(AscW(ArabicData.ArabicLetters(CInt(A.GetValue(1))).Symbol)))) + "); }", _
            "var uthmaniMinimalScript = " + Utility.MakeJSArray(New String() {Utility.MakeJSIndexedObject(New String() {"rule", "match", "evaluator", "negativematch", "ruleFunc"}, _
                                    Array.ConvertAll(Of IslamData.RuleTranslationCategory.RuleTranslation, Object())(CachedData.UthmaniMinimalScript, Function(Convert As IslamData.RuleTranslationCategory.RuleTranslation) New Object() {Utility.MakeJSString(Convert.Name), Utility.MakeJSString(Utility.EscapeJS(Convert.Match)), Utility.MakeJSString(Convert.Evaluator), Utility.MakeJSString(Convert.NegativeMatch), Convert.RuleFunc}), True)}, True) + ";", _
            "var simpleEnhancedScript = " + Utility.MakeJSArray(New String() {Utility.MakeJSIndexedObject(New String() {"rule", "match", "evaluator", "negativematch", "ruleFunc"}, _
                                    Array.ConvertAll(Of IslamData.RuleTranslationCategory.RuleTranslation, Object())(CachedData.SimpleEnhancedScript, Function(Convert As IslamData.RuleTranslationCategory.RuleTranslation) New Object() {Utility.MakeJSString(Convert.Name), Utility.MakeJSString(Utility.EscapeJS(Convert.Match)), Utility.MakeJSString(Convert.Evaluator), Utility.MakeJSString(Convert.NegativeMatch), Convert.RuleFunc}), True)}, True) + ";", _
            "var simpleScript = " + Utility.MakeJSArray(New String() {Utility.MakeJSIndexedObject(New String() {"rule", "match", "evaluator", "negativematch", "ruleFunc"}, _
                                    Array.ConvertAll(Of IslamData.RuleTranslationCategory.RuleTranslation, Object())(CachedData.SimpleScript, Function(Convert As IslamData.RuleTranslationCategory.RuleTranslation) New Object() {Utility.MakeJSString(Convert.Name), Utility.MakeJSString(Utility.EscapeJS(Convert.Match)), Utility.MakeJSString(Convert.Evaluator), Utility.MakeJSString(Convert.NegativeMatch), Convert.RuleFunc}), True)}, True) + ";", _
            "var simpleMinimalScript = " + Utility.MakeJSArray(New String() {Utility.MakeJSIndexedObject(New String() {"rule", "match", "evaluator", "negativematch", "ruleFunc"}, _
                                    Array.ConvertAll(Of IslamData.RuleTranslationCategory.RuleTranslation, Object())(CachedData.SimpleMinimalScript, Function(Convert As IslamData.RuleTranslationCategory.RuleTranslation) New Object() {Utility.MakeJSString(Convert.Name), Utility.MakeJSString(Utility.EscapeJS(Convert.Match)), Utility.MakeJSString(Convert.Evaluator), Utility.MakeJSString(Convert.NegativeMatch), Convert.RuleFunc}), True)}, True) + ";", _
            "var simpleCleanScript = " + Utility.MakeJSArray(New String() {Utility.MakeJSIndexedObject(New String() {"rule", "match", "evaluator", "negativematch", "ruleFunc"}, _
                                    Array.ConvertAll(Of IslamData.RuleTranslationCategory.RuleTranslation, Object())(CachedData.SimpleCleanScript, Function(Convert As IslamData.RuleTranslationCategory.RuleTranslation) New Object() {Utility.MakeJSString(Convert.Name), Utility.MakeJSString(Utility.EscapeJS(Convert.Match)), Utility.MakeJSString(Convert.Evaluator), Utility.MakeJSString(Convert.NegativeMatch), Convert.RuleFunc}), True)}, True) + ";", _
            "var errorCheckRules = " + Utility.MakeJSArray(New String() {Utility.MakeJSIndexedObject(New String() {"rule", "match", "evaluator", "negativematch", "ruleFunc"}, _
                                    Array.ConvertAll(Of IslamData.RuleTranslationCategory.RuleTranslation, Object())(CachedData.ErrorCheckRules, Function(Convert As IslamData.RuleTranslationCategory.RuleTranslation) New Object() {Utility.MakeJSString(Convert.Name), Utility.MakeJSString(Utility.EscapeJS(Convert.Match)), Utility.MakeJSString(Convert.Evaluator), Utility.MakeJSString(Convert.NegativeMatch), Convert.RuleFunc}), True)}, True) + ";", _
            "var coloringSpelledOutRules = " + Utility.MakeJSArray(New String() {Utility.MakeJSIndexedObject(New String() {"rule", "match", "evaluator", "ruleFunc"}, _
                                    Array.ConvertAll(Of IslamData.RuleTranslationCategory.RuleTranslation, Object())(CachedData.ColoringSpelledOutRules, Function(Convert As IslamData.RuleTranslationCategory.RuleTranslation) New Object() {Utility.MakeJSString(Convert.Name), Utility.MakeJSString(Utility.EscapeJS(Convert.Match)), Utility.MakeJSString(Convert.Evaluator), Convert.RuleFunc}), True)}, True) + ";", _
            "var romanizationRules = " + Utility.MakeJSArray(New String() {Utility.MakeJSIndexedObject(New String() {"rule", "match", "evaluator", "ruleFunc"}, _
                                    Array.ConvertAll(Of IslamData.RuleTranslationCategory.RuleTranslation, Object())(CachedData.RomanizationRules, Function(Convert As IslamData.RuleTranslationCategory.RuleTranslation) New Object() {Utility.MakeJSString(Convert.Name), Utility.MakeJSString(Utility.EscapeJS(Convert.Match)), Utility.MakeJSString(Convert.Evaluator), Convert.RuleFunc}), True)}, True) + ";", _
            "var coloringRules = " + Utility.MakeJSArray(New String() {Utility.MakeJSIndexedObject(New String() {"rule", "match", "color"}, _
                                    Array.ConvertAll(Of IslamData.ColorRule, String())(CachedData.IslamData.ColorRules, Function(Convert As IslamData.ColorRule) New String() {Convert.Name, Utility.EscapeJS(Convert.Match), System.Drawing.ColorTranslator.ToHtml(Convert.Color)}), False)}, True) + ";", _
            "var rulesOfRecitationRegEx = " + Utility.MakeJSArray(New String() {Utility.MakeJSIndexedObject(New String() {"rule", "match", "evaluator"}, _
                                    Array.ConvertAll(Of IslamData.RuleMetaSet.RuleMetadataTranslation, Object())(CachedData.RulesOfRecitationRegEx, Function(Convert As IslamData.RuleMetaSet.RuleMetadataTranslation) New Object() {Utility.MakeJSString(Convert.Name), Utility.MakeJSString(Utility.EscapeJS(Convert.Match)), If(Convert.Evaluator Is Nothing, Nothing, Utility.MakeJSArray(Convert.Evaluator))}), True)}, True) + ";", _
            "var allowZeroLength = " + Utility.MakeJSArray(AllowZeroLength) + ";", _
            "function ruleMetadataComparer(a, b) { return (a.index === b.index) ? (b.length === a.length ? b.origOrder - a.origOrder : b.length - a.length) : b.index - a.index; }", _
            "function replaceMetadata(sVal, metadataRule, scheme, learningMode) { var count, elimParen = function(s) { return s.replace(/\(.*\)/, ''); }; for (count = 0; count < coloringSpelledOutRules.length; count++) { var index, match = null; for (index = 0; index < coloringSpelledOutRules[count].match.split('|').length; index++) { if (metadataRule.type.split('|').map(elimParen).indexOf(coloringSpelledOutRules[count].match.split('|')[index]) !== -1) { match = coloringSpelledOutRules[count].match.split('|')[index]; break; } } if (match !== null) { var str = coloringSpelledOutRules[count].evaluator.format(sVal.substr(metadataRule.index, metadataRule.length)); if (coloringSpelledOutRules[count].ruleFunc !== 0) { var args = ruleFunctions[coloringSpelledOutRules[count].ruleFunc - 1](str, scheme, learningMode); if (args.length === 1) { str = args[0]; } else { var metaArgs = metadataRule.type.match(/\((.*)\)/)[1].split(','); str = ''; for (index = 0; index < args.length; index++) { if (args[index] !== undefined && args[index] !== null && (learningMode || coloringSpelledOutRules[count].ruleFunc !== " + CStr(RuleFuncs.eLearningMode) + " || index !== 0)) str += replaceMetadata(args[index], {index: 0, length: args[index].length, type: metaArgs[index].replace(' ', '|'), origOrder: index}, scheme, learningMode); } } } sVal = sVal.substr(0, metadataRule.index) + str + sVal.substr(metadataRule.index + metadataRule.length); } } return sVal; }", _
            "function joinContig(sVal, preString, postString) { var index = preString.lastIndexOf(' '); if (index !== -1 && preString.length - 2 === index) index = preString.lastIndexOf(' ', index - 1); if (index !== -1) preString = preString.substring(index + 1); if (preString !== '') preString += ' ' + String.fromCharCode(0x6DD) + ' '; index = postString.indexOf(' '); if (index === 2) index = preString.indexOf(' ', index + 1); if (index !== -1) postString = postString.substring(0, index); if (postString !== '') postString = ' ' + String.fromCharCode(0x6DD) + ' ' + postString; return preString + sVal + postString; }", _
            "function unjoinContig(sVal, preString, postString) { var index = sVal.indexOf(String.fromCharCode(0x6DD)); if (preString !== '' && index !== -1) sVal = sVal.substring(index + 1 +  1); index = sVal.lastIndexOf(String.fromCharCode(0x6DD)); if (postString !== '' && index !== -1) sVal = sVal.substring(0, index - 1); return sVal; }", _
            "function transliterateContigWithRules(sVal, preString, postString, scheme, optionalStops) { return unjoinContig(transliterateWithRules(JoinContig(sVal, preString, postString), scheme, optionalStops, false), preString, postString); }", _
            "function transliterateWithRules(sVal, scheme, optionalStops, learningMode) { var count, index, arr, re, metadataList = [], replaceFunc = function(f, e) { return function() { return f(RegExp.matchResult(e, arguments[arguments.length - 2], arguments[arguments.length - 1], Array.prototype.slice.call(arguments).slice(0, -2)), scheme)[0]; }; }; for (count = 0; count < errorCheckRules.length; count++) { re = new RegExp(errorCheckRules[count].match, 'g'); while ((arr = re.exec(sVal)) !== null) { if (!errorCheckRules[count].negativematch || RegExp.matchResult(errorCheckRules[count].negativematch, arr.index, sVal, arr) === '') { console.log(errorCheckRules[count].rule + ': ' + doTransliterate(sVal.substr(0, arr.index), true, 1) + '<!-- -->' + doTransliterate(sVal.substr(arr.index), true, 1)); } } } for (count = 0; count < rulesOfRecitationRegEx.length; count++) { if (rulesOfRecitationRegEx[count].evaluator !== null) { var subcount, lindex; re = new RegExp(rulesOfRecitationRegEx[count].match, 'g'); while ((arr = re.exec(sVal)) !== null) { lindex = arr.index; for (subcount = 0; subcount < rulesOfRecitationRegEx[count].evaluator.length; subcount++) { if (rulesOfRecitationRegEx[count].evaluator[subcount] === 'optionalstop' && (optionalStops === null && arr[subcount + 1] === String.fromCharCode(0x6D6) || (optionalStops !== null && arr[subcount + 1] !== '' && optionalStops.indexOf(lindex) === -1)) || rulesOfRecitationRegEx[count].evaluator[subcount] === 'optionalnotstop' && (optionalStops === null && arr[subcount + 1] !== '' && arr[subcount + 1] !== String.fromCharCode(0x6D6) || (arr[subcount + 1].length === 0 && lindex === 0) || optionalStops !== null && arr[subcount + 1] !== '' && optionalStops.indexOf(lindex) !== -1)) break; } if (subcount !== rulesOfRecitationRegEx[count].evaluator.length) continue; for (subcount = 0; subcount < rulesOfRecitationRegEx[count].evaluator.length; subcount++) { if (rulesOfRecitationRegEx[count].evaluator[subcount] !== null && (arr[subcount + 1] && arr[subcount + 1].length !== 0 || allowZeroLength.indexOf(rulesOfRecitationRegEx[count].evaluator[subcount]) !== -1)) { metadataList.push({index: lindex, length: arr[subcount + 1] ? arr[subcount + 1].length : 0, type: rulesOfRecitationRegEx[count].evaluator[subcount], origOrder: subcount}); } lindex += (arr[subcount + 1] ? arr[subcount + 1].length : 0); } } } } metadataList.sort(ruleMetadataComparer); for (index = 0; index < metadataList.length; index++) { sVal = replaceMetadata(sVal, metadataList[index], scheme, learningMode); } for (count = 0; count < romanizationRules.length; count++) { sVal = sVal.replace(new RegExp(romanizationRules[count].match, 'g'), (romanizationRules[count].ruleFunc === 0) ? romanizationRules[count].evaluator : replaceFunc(ruleFunctions[romanizationRules[count].ruleFunc - 1], romanizationRules[count].evaluator)); } return sVal; }"}
    Public Shared NumberGenJS As String() = {"var arabicOrdinalNumbers = " + Utility.MakeJSArray(CachedData.ArabicOrdinalNumbers) + ";", _
                "var arabicOrdinalExtraNumbers = " + Utility.MakeJSArray(CachedData.ArabicOrdinalExtraNumbers) + ";", _
                "var arabicFractionNumbers = " + Utility.MakeJSArray(CachedData.ArabicFractionNumbers) + ";", _
                "var arabicBaseNumbers = " + Utility.MakeJSArray(CachedData.ArabicBaseNumbers) + ";", _
                "var arabicBaseExtraNumbers = " + Utility.MakeJSArray(CachedData.ArabicBaseExtraNumbers) + ";", _
                "var arabicBaseTenNumbers = " + Utility.MakeJSArray(CachedData.ArabicBaseTenNumbers) + ";", _
                "var arabicBaseHundredNumbers = " + Utility.MakeJSArray(CachedData.ArabicBaseHundredNumbers) + ";", _
                "var arabicBaseThousandNumbers = " + Utility.MakeJSArray(CachedData.ArabicBaseThousandNumbers) + ";", _
                "var arabicBaseMillionNumbers = " + Utility.MakeJSArray(CachedData.ArabicBaseMillionNumbers) + ";", _
                "var arabicBaseBillionNumbers = " + Utility.MakeJSArray(CachedData.ArabicBaseBillionNumbers) + ";", _
                "var arabicBaseMilliardNumbers = " + Utility.MakeJSArray(CachedData.ArabicBaseMilliardNumbers) + ";", _
                "var arabicBaseTrillionNumbers = " + Utility.MakeJSArray(CachedData.ArabicBaseTrillionNumbers) + ";", _
                "var arabicCombiners = " + Utility.MakeJSArray(CachedData.ArabicCombiners) + ";", _
                "function doTransliterateNum() { $('#translitvalue').text(doTransliterate(arabicWordFromNumber($('#translitedit').val(), $('#useclassic0').prop('checked'), $('#usehundredform0').prop('checked'), $('#usemilliard0').prop('checked')), false, 1)); }", _
                "function arabicWordForLessThanThousand(number, useclassic, usealefhundred) { var str = '', hundstr = ''; if (number >= 100) { hundstr = usealefhundred ? arabicBaseHundredNumbers[Math.floor(number / 100) - 1].substr(0, 2) + 'A' + arabicBaseHundredNumbers[Math.floor(number / 100) - 1].substr(2) : arabicBaseHundredNumbers[Math.floor(number / 100) - 1]; if ((number % 100) === 0) { return hundstr; } number = number % 100; } if ((number % 10) !== 0 && number !== 11 && number !== 12) { str = arabicBaseNumbers[number % 10 - 1]; } if (number >= 11 && number < 20) { if (number == 11 || number == 12) { str += arabicBaseExtraNumbers[number - 11]; } else { str = str.slice(0, -1) + 'a'; } str += ' ' + arabicBaseTenNumbers[1].slice(0, -2); } else if ((number === 0 && str === '') || number === 10 || number >= 20) { str = ((str === '') ? '' : str + ' ' + arabicCombiners[0]) + arabicBaseTenNumbers[Math.floor(number / 10)]; } return useclassic ? (((str === '') ? '' : str + ((hundstr === '') ? '' : ' ' + arabicCombiners[0])) + hundstr) : (((hundstr === '') ? '' : hundstr + ((str === '') ? '' : ' ' + arabicCombiners[0])) + str); }", _
                "function arabicWordFromNumber(number, useclassic, usealefhundred, usemilliard) { var str = '', nextstr = '', curbase = 3, basenums = [1000, 1000000, 1000000000, 1000000000000], bases = [arabicBaseThousandNumbers, arabicBaseMillionNumbers, usemilliard ? arabicBaseMilliardNumbers : arabicBaseBillionNumbers, arabicBaseTrillionNumbers]; do { if (number >= basenums[curbase] && number < 2 * basenums[curbase]) { nextstr = bases[curbase][0]; } else if (number >= 2 * basenums[curbase] && number < 3 * basenums[curbase]) { nextstr = bases[curbase][1]; } else if (number >= 3 * basenums[curbase] && number < 10 * basenums[curbase]) { nextstr = arabicBaseNumbers[Math.floor(Number / basenums[curbase]) - 1].slice(0, -1) + 'u ' + bases[curbase][2]; } else if (number >= 10 * basenums[curbase] && number < 11 * basenums[curbase]) { nextstr = arabicBaseTenNumbers[1].slice(0, -1) + 'u ' + bases[curbase][2]; } else if (number >= basenums[curbase]) { nextstr = arabicWordForLessThanThousand(Math.floor(number / basenums[curbase]) % 100, useclassic, usealefhundred); if (number >= 100 * basenums[curbase] && number < (useclassic ? 200 : 101) * basenums[curbase]) { nextstr = nextstr.slice(0, -1) + 'u ' + bases[curbase][0].slice(0, -1) + 'K'; } else if (number >= 200 * basenums[curbase] && number < (useclassic ? 300 : 201) * basenums[curbase]) { nextstr = nextstr.slice(0, -2) + ' ' + bases[curbase][0].slice(0, -1) + 'K'; } else if (number >= 300 * basenums[curbase] && (useclassic || Math.floor(number / basenums[curbase]) % 100 === 0)) { nextstr = nextstr.slice(0, -1) + 'i ' + bases[curbase][0].slice(0, -1) + 'K'; } else { nextstr += ' ' + bases[curbase][0].slice(0, -1) + 'FA'; } } number = number % basenums[curbase]; curbase--; str = useclassic ? (((nextstr === '') ? '' : nextstr + ((str === '') ? '' : ' ' + arabicCombiners[0])) + str) : (((str === '') ? '' : str + ((nextstr === '') ? '' : ' ' + arabicCombiners[0])) + nextstr); nextstr = ''; } while (curbase >= 0); if (number !== 0 || str === '') { nextstr = arabicWordForLessThanThousand(number, useclassic, usealefhundred); } return useclassic ? (((nextstr === '') ? '' : nextstr + ((str === '') ? '' : ' ' + arabicCombiners[0])) + str) : (((str === '') ? '' : str + ((nextstr === '') ? '' : ' ' + arabicCombiners[0])) + nextstr); }"}
    Public Shared Function GetTransliterateNumberJS() As String()
        Dim GetJS As New List(Of String) From {"javascript: doTransliterateNum();", String.Empty, GetArabicSymbolJSArray(), GetTranslitSchemeJSArray()}
        GetJS.AddRange(ArabicData.GetUniCats())
        GetJS.AddRange(PlainTransliterateGenJS)
        GetJS.AddRange(TransliterateGenJS)
        GetJS.AddRange(NumberGenJS)
        Return GetJS.ToArray()
    End Function
    Public Shared Function GetTransliterateJS() As String()
        Dim GetJS As New List(Of String) From {"javascript: doTransliterateDisplay();", String.Empty, GetArabicSymbolJSArray(), GetTranslitSchemeJSArray(), _
        "function doDirectionDom(elem, sVal, direction) { elem.css('direction', direction ? 'ltr' : 'rtl'); var stack = [], lastStrong = -1, lastCount = 0, iCount; for (iCount = 0; iCount < sVal.length; iCount++) { if (sVal.charCodeAt(iCount) === 0x200E || sVal.charCodeAt(iCount) === 0x200F || sVal.charCodeAt(iCount) === 0x61C) { if (lastStrong !== iCount - 1) {  } } else if (IsExplicit(sVal.charCodeAt(iCount))) { if (sVal.charCodeAt(iCount) === 0x202C || sVal.charCodeAt(iCount) === 0x2069) { stack.pop()[1].add(document.createTextNode(sVal.substring(lastCount, iCount - 1))); lastCount = iCount + 1; lastStrong = -1; } else { (stack.length === 0 ? elem : stack[stack.length - 1][1]).add(document.createTextNode(sVal.substring(lastCount, iCount - 1))); lastCount = iCount + 1; lastStrong = -1; stack.push([sVal[iCount], (stack.length === 0 ? elem : stack[stack.length - 1][1]).add('span')]); stack[stack.length - 1][1].css('direction', (sVal[iCount] === 0x202D || sVal[iCount] === 0x202A || sVal[iCount] === 0x2066) ? 'ltr' : 'rtl'); } } else if (!IsNeutral(sVal.charCodeAt(iCount))) { lastStrong = iCount; } } (stack.length === 0 ? elem : stack[stack.length - 1][1]).add(document.createTextNode(sVal.substring(lastCount, iCount - 1))); }", _
        "function doTransliterateDisplay() { $('#translitvalue').css('direction', !$('#scheme1').prop('checked') && $('#direction0').prop('checked') ? 'ltr' : 'rtl'); $('#translitvalue').empty(); $('#translitvalue').text($('#scheme0').prop('checked') ? doTransliterate($('#translitedit').val(), $('#direction0').prop('checked'), parseInt($('#translitscheme').val(), 10)) : ($('#scheme1').prop('checked') ? doDiacritics($('#translitedit').val(), $('#diacriticscheme0').prop('checked')) : $('#translitedit').val())); $('#translitvalue').html($('#translitvalue').html().replace(/\n/g, '<br>')); }"}
        GetJS.AddRange(ArabicData.GetUniCats())
        GetJS.AddRange(PlainTransliterateGenJS)
        GetJS.AddRange(TransliterateGenJS)
        GetJS.AddRange(DiacriticJS)
        Return GetJS.ToArray()
    End Function
    Public Shared Function GetSchemeChangeJS() As String()
        Return New String() {"javascript: doSchemeChange();", String.Empty, "function doSchemeChange() { $('#diacriticscheme_').css('display', $('#scheme0').prop('checked') ? 'none' : 'block'); $('#translitscheme').css('display', $('#scheme0').prop('checked') ? 'block' : 'none'); $('#direction_').css('display', $('#scheme0').prop('checked') ? 'block' : 'none'); }"}
    End Function
    Public Shared Function GetTransliterationSchemes() As Array()
        Dim Count As Integer
        Dim Strings(CachedData.IslamData.TranslitSchemes.Length * 2 + 2 - 2 - 1) As Array
        Strings(0) = New String() {Utility.LoadResourceString("IslamSource_Off"), "0"}
        Strings(1) = New String() {Utility.LoadResourceString("IslamSource_ExtendedBuckwalter"), "1"}
        For Count = 0 To CachedData.IslamData.TranslitSchemes.Length - 2
            Strings(Count * 2 + 2) = New String() {Utility.LoadResourceString("IslamSource_" + CachedData.IslamData.TranslitSchemes(Count + 1).Name), CStr(Count * 2 + 2)}
            Strings(Count * 2 + 1 + 2) = New String() {Utility.LoadResourceString("IslamSource_" + CachedData.IslamData.TranslitSchemes(Count + 1).Name) + " Literal", CStr(Count * 2 + 1 + 2)}
        Next
        Return Strings
    End Function
    Public Shared Function GetChangeTransliterationJS() As String()
        Dim GetJS As New List(Of String) From {"javascript: changeTransliteration();", String.Empty, Utility.GetLookupStyleSheetJS(), GetArabicSymbolJSArray(), GetTranslitSchemeJSArray(), _
        "function processTransliteration(list) { var k, child, iSubCount, text; $('span.transliteration').each(function() { $(this).css('display', $('#translitscheme').val() === '0' ? 'none' : 'block'); }); for (k in list) { text = ''; if (list.hasOwnProperty(k) && list[k]['linkchild']) { for (child in list[k]['children']) { if (list[k]['children'].hasOwnProperty(child)) { processTransliteration(list[k]['children'][child]['children']); for (iSubCount = 0; iSubCount < list[k]['children'][child]['arabic'].length; iSubCount++) { if ($('#translitscheme').val() !== '0' && $('#translitscheme').val() !== '1'  && parseInt($('#translitscheme').val(), 10) % 2 !== 1 && list[k]['children'][child]['arabic'][iSubCount] !== '' && list[k]['children'][child]['translit'][iSubCount] !== '') { if (text !== '') text += ' '; text += $('#' + list[k]['children'][child]['arabic'][iSubCount]).text(); } else { if (list[k]['children'][child]['translit'][iSubCount] !== '') $('#' + list[k]['children'][child]['translit'][iSubCount]).text(($('#translitscheme').val() === '0' || list[k]['children'][child]['arabic'][iSubCount] === '') ? '' : doTransliterate($('#' + list[k]['children'][child]['arabic'][iSubCount]).text(), true, parseInt($('#translitscheme').val(), 10))); } } } } if ($('#translitscheme').val() !== '0' && $('#translitscheme').val() !== '1' && parseInt($('#translitscheme').val(), 10) % 2 !== 1) { text = transliterateWithRules(text, Math.floor((parseInt($('#translitscheme').val(), 10) - 2) / 2) + 2, null, false).split(' '); for (child in list[k]['children']) { if (list[k]['children'].hasOwnProperty(child)) { for (iSubCount = 0; iSubCount < list[k]['children'][child]['translit'].length; iSubCount++) { if (list[k]['children'][child]['arabic'][iSubCount] !== '' && list[k]['children'][child]['translit'][iSubCount] !== '') $('#' + list[k]['children'][child]['translit'][iSubCount]).text(text.shift()); } } } } } else { processTransliteration(list[k]['children']); } for (iSubCount = 0; iSubCount < list[k]['arabic'].length; iSubCount++) { if (list[k]['translit'][iSubCount] !== '') $('#' + list[k]['translit'][iSubCount]).text(($('#translitscheme').val() === '0' || list[k]['arabic'][iSubCount] === '') ? '' : (($('#translitscheme').val() !== '0' && $('#translitscheme').val() !== '1' && parseInt($('#translitscheme').val(), 10) % 2 !== 1) ? transliterateWithRules($('#' + list[k]['arabic'][iSubCount]).text(), parseInt($('#translitscheme').val(), 10) >= 2 ? Math.floor((parseInt($('#translitscheme').val(), 10) - 2) / 2) + 2 : parseInt($('#translitscheme').val(), 10), null, false) : doTransliterate($('#' + list[k]['arabic'][iSubCount]).text(), true, parseInt($('#translitscheme').val(), 10)))); } } }", _
        "function changeTransliteration() { changeChapterTranslit(); var i; for (i = 0; i < renderList.length; i++) { processTransliteration(renderList[i]); } }", _
        "function changeChapterTranslit() { var i; for (i = 0; i < $('#quranselection').get(0).options.length; i++) { $('#quranselection').get(0).options[i].text = $('#quranselection').get(0).options[i].text.replace(/(\(.*? )(.*?)(\))( )?(.*)/g, function (m, open, a, close) { return open + a + close + ' ' + (($('#translitscheme').val() === '0' || a === '') ? '' : (($('#translitscheme').val() !== '0' && $('#translitscheme').val() !== '1' && parseInt($('#translitscheme').val(), 10) % 2 !== 1) ? transliterateWithRules(a, parseInt($('#translitscheme').val(), 10) >= 2 ? Math.floor((parseInt($('#translitscheme').val(), 10) - 2) / 2) + 2 : parseInt($('#translitscheme').val(), 10), null, false) : doTransliterate(a, true, parseInt($('#translitscheme').val(), 10)))); }); } }"}
        GetJS.AddRange(ArabicData.GetUniCats())
        GetJS.AddRange(PlainTransliterateGenJS)
        GetJS.AddRange(TransliterateGenJS)
        GetJS.AddRange(NumberGenJS)
        Return GetJS.ToArray()
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
                        Val = ArabicData.ArabicFatha + Val
                    ElseIf (Matches(MatchCount).Groups(3).Value(Len) = "i") Then
                        Val = ArabicData.ArabicKasra + Val
                    ElseIf (Matches(MatchCount).Groups(3).Value(Len) = "u") Then
                        Val = ArabicData.ArabicDamma + Val
                    ElseIf (Matches(MatchCount).Groups(3).Value(Len) = "ā" Or Len <> 0 AndAlso ((Matches(MatchCount).Groups(3).Value(Len) = "i" Or Matches(MatchCount).Groups(3).Value(SubCount) = "u") And Matches(MatchCount).Groups(3).Value(SubCount - 1) = "a")) Then
                        Val = ArabicData.ArabicFatha + Val
                    ElseIf (Matches(MatchCount).Groups(3).Value(Len) = "ī" Or Matches(MatchCount).Groups(3).Value(Len) = "ē") Then
                        Val = ArabicData.ArabicKasra + Val
                    ElseIf (Matches(MatchCount).Groups(3).Value(Len) = "ū" Or Matches(MatchCount).Groups(3).Value(Len) = "ō") Then
                        Val = ArabicData.ArabicDamma + Val
                    ElseIf Len <> 0 AndAlso Matches(MatchCount).Groups(3).Value(Len) = Matches(MatchCount).Groups(3).Value(Len - 1) Then
                        Val = ArabicData.ArabicShadda + Val
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
        Dim SchemeType As ArabicData.TranslitScheme = CType(If(CInt(HttpContext.Current.Request.Params("translitscheme")) >= 2, 2 - CInt(HttpContext.Current.Request.Params("translitscheme")) Mod 2, CInt(HttpContext.Current.Request.Params("translitscheme"))), ArabicData.TranslitScheme)
        Dim Scheme As String = If(CInt(HttpContext.Current.Request.Params("translitscheme")) >= 2, CachedData.IslamData.TranslitSchemes(CInt(HttpContext.Current.Request.Params("translitscheme")) \ 2).Name, String.Empty)
        Dim Output As New List(Of String())
        Output.Add(New String() {})
        Output.Add(New String() {"arabic", "transliteration", "arabic", String.Empty, String.Empty, String.Empty, String.Empty})
        Output.Add(New String() {Utility.LoadResourceString("IslamInfo_LetterName"), Utility.LoadResourceString("IslamInfo_Transliteration"), Utility.LoadResourceString("IslamInfo_Arabic"), Utility.LoadResourceString("IslamSource_ExtendedBuckwalter"), Utility.LoadResourceString("IslamInfo_Terminating"), Utility.LoadResourceString("IslamInfo_Connecting"), Utility.LoadResourceString("IslamInfo_Shaping")})
        'Dim Combos(ArabicData.Data.ArabicCombos.Length - 1) As IslamData.ArabicCombo
        'ArabicData.ArabicLetters.CopyTo(ArabicData.Data.ArabicCombos, 0)
        'Array.Sort(Combos, Function(Key As IslamData.ArabicCombo, NextKey As IslamData.ArabicCombo) Key.SymbolName.CompareTo(NextKey.SymbolName))
        For Count = 0 To ArabicData.ArabicCombos.Length - 1
            If Array.TrueForAll(ArabicData.ArabicCombos(Count).Symbol, Function(Ch As Char) GetSchemeValueFromSymbol(ArabicData.ArabicLetters(ArabicData.FindLetterBySymbol(Ch)), "ExtendedBuckwalter") <> String.Empty) Then
                Output.Add(New String() {ArabicLetterSpelling(String.Join(String.Empty, Array.ConvertAll(ArabicData.ArabicCombos(Count).Symbol, Function(Sym As Char) CStr(Sym))), False), _
                                         TransliterateToScheme(ArabicLetterSpelling(String.Join(String.Empty, Array.ConvertAll(ArabicData.ArabicCombos(Count).Symbol, Function(Sym As Char) CStr(Sym))), False), SchemeType, Scheme), _
                                                    String.Join(String.Empty, Array.ConvertAll(ArabicData.ArabicCombos(Count).Symbol, Function(Sym As Char) CStr(Sym))), _
                                                    TransliterateToScheme(String.Join(String.Empty, Array.ConvertAll(ArabicData.ArabicCombos(Count).Symbol, Function(Sym As Char) CStr(Sym))), ArabicData.TranslitScheme.Literal, String.Empty), _
                                                    CStr(ArabicData.ArabicCombos(Count).Terminating), _
                                                  CStr(ArabicData.ArabicCombos(Count).Connecting), _
                                                    String.Join(vbCrLf, Array.ConvertAll(ArabicData.ArabicCombos(Count).Shaping, Function(Shape As Char) If(Shape = ChrW(0), String.Empty, Shape + " " + CStr(Hex(AscW(Shape))) + " " + If(CheckShapingOrder(Array.IndexOf(ArabicData.ArabicCombos(Count).Shaping, Shape), ArabicData.GetUnicodeName(Shape)), String.Empty, "!!!") + ArabicData.GetUnicodeName(Shape))))})
            End If
        Next
        Return Output.ToArray()
    End Function
    Public Shared Function SymbolDisplay(Symbols() As ArabicData.ArabicSymbol, SchemeType As ArabicData.TranslitScheme, Scheme As String, Cols As String()) As Array()
        Dim Count As Integer
        Dim Output(Symbols.Length + 2) As Array
        Dim ColSet As String() = {"LetterName", "Transliteration", "UnicodeName", "Arabic", "UnicodeValue", "ExtendedBuckwalter", "Terminating", "Connecting", "Shaping"}
        'validation of Cols needed...
        If Cols Is Nothing Then Cols = ColSet
        Dim ColStyles As String() = {"arabic", "transliteration", String.Empty, "arabic", String.Empty, String.Empty, String.Empty, String.Empty, String.Empty}
        'Dim oFont As New Font(DefaultValue(HttpContext.Current.Request.QueryString.Get("fontcustom"), "Arial"), 13)
        'CheckIfCharInFont(ArabicData.ArabicLetters(Count).Symbol, oFont)
        Output(0) = New String() {}
        Output(1) = Array.ConvertAll(Cols, Function(Str As String) ColStyles(Array.IndexOf(ColSet, Str)))
        Output(2) = Array.ConvertAll(Cols, Function(Str As String) Utility.LoadResourceString("IslamInfo_" + Str))
        For Count = 0 To Symbols.Length - 1
            '"arabic", String.Empty
            'Utility.LoadResourceString("IslamInfo_LetterName"), Utility.LoadResourceString("IslamInfo_Assimilate")
            'Arabic.TransliterateFromBuckwalter(Symbols(Count).SymbolName), _
            'CStr(Symbols(Count).Assimilate),
            Output(Count + 3) = Array.ConvertAll(Cols,
                Function(Str As String)
                    Select Case Array.IndexOf(ColSet, Str)
                        Case 0
                            Return ArabicLetterSpelling(CStr(Symbols(Count).Symbol), False)
                        Case 1
                            Return TransliterateToScheme(ArabicLetterSpelling(CStr(Symbols(Count).Symbol), False), SchemeType, Scheme)
                        Case 2
                            Return ArabicData.GetUnicodeName(Symbols(Count).Symbol)
                        Case 3
                            Return CStr(Symbols(Count).Symbol)
                        Case 4
                            Return CStr(Hex(AscW(Symbols(Count).Symbol)))
                        Case 5
                            Return CStr(If(GetSchemeValueFromSymbol(Symbols(Count), "ExtendedBuckwalter").Length = 0, String.Empty, GetSchemeValueFromSymbol(Symbols(Count), "ExtendedBuckwalter").Chars(0)))
                        Case 6
                            Return CStr(Symbols(Count).Terminating)
                        Case 7
                            Return CStr(Symbols(Count).Connecting)
                        Case 8
                            Return If(Symbols(Count).Shaping = Nothing, String.Empty, String.Join(vbCrLf, Array.ConvertAll(Symbols(Count).Shaping, Function(Shape As Char) If(Shape = ChrW(0), String.Empty, Shape + " " + CStr(Hex(AscW(Shape))) + " " + If(CheckShapingOrder(Array.IndexOf(Symbols(Count).Shaping, Shape), ArabicData.GetUnicodeName(Shape)), String.Empty, "!!!") + ArabicData.GetUnicodeName(Shape)))))
                        Case Else
                            Return Nothing
                    End Select
                End Function)
        Next
        Return Output
    End Function
    Public Shared Function DisplayAll(ByVal Item As PageLoader.TextItem) As Array()
        Dim SchemeType As ArabicData.TranslitScheme = CType(If(CInt(HttpContext.Current.Request.Params("translitscheme")) >= 2, 2 - CInt(HttpContext.Current.Request.Params("translitscheme")) Mod 2, CInt(HttpContext.Current.Request.Params("translitscheme"))), ArabicData.TranslitScheme)
        Dim Scheme As String = If(CInt(HttpContext.Current.Request.Params("translitscheme")) >= 2, CachedData.IslamData.TranslitSchemes(CInt(HttpContext.Current.Request.Params("translitscheme")) \ 2).Name, String.Empty)
        Return SymbolDisplay(Array.FindAll(ArabicData.ArabicLetters, Function(Letter As ArabicData.ArabicSymbol) GetSchemeValueFromSymbol(Letter, "ExtendedBuckwalter") <> String.Empty), SchemeType, Scheme, Nothing)
    End Function
    Public Shared Function DisplayTranslitSchemes(ByVal Item As PageLoader.TextItem) As Array()
        Dim Count As Integer
        Dim Output As New List(Of String())
        'Dim oFont As New Font(DefaultValue(HttpContext.Current.Request.QueryString.Get("fontcustom"), "Arial"), 13)
        'CheckIfCharInFont(ArabicData.ArabicLetters(Count).Symbol, oFont)
        Output.Add(New String() {})
        Dim Strs As String() = New String() {"arabic", String.Empty, "arabic"}
        Array.Resize(Of String)(Strs, 3 + CachedData.IslamData.TranslitSchemes.Length)
        Output.Add(Strs)
        Strs = New String() {Utility.LoadResourceString("IslamInfo_LetterName"), Utility.LoadResourceString("IslamInfo_UnicodeName"), Utility.LoadResourceString("IslamInfo_Arabic")}
        Array.Resize(Of String)(Strs, 3 + CachedData.IslamData.TranslitSchemes.Length)
        For SchemeCount = 0 To CachedData.IslamData.TranslitSchemes.Length - 1
            CType(Output(1), String())(3 + SchemeCount) = String.Empty
            Strs(3 + SchemeCount) = Utility.LoadResourceString("IslamSource_" + CachedData.IslamData.TranslitSchemes(SchemeCount).Name)
        Next
        Output.Add(Strs)
        For Count = 0 To ArabicData.ArabicLetters.Length - 1
            If GetSchemeValueFromSymbol(ArabicData.ArabicLetters(Count), "ExtendedBuckwalter") <> String.Empty Then
                Strs = New String() {ArabicLetterSpelling(ArabicData.ArabicLetters(Count).Symbol, False),
                                           ArabicData.GetUnicodeName(ArabicData.ArabicLetters(Count).Symbol), _
                                           CStr(ArabicData.ArabicLetters(Count).Symbol)}
                Array.Resize(Of String)(Strs, 3 + CachedData.IslamData.TranslitSchemes.Length)
                For SchemeCount = 0 To CachedData.IslamData.TranslitSchemes.Length - 1
                    Strs(3 + SchemeCount) = GetSchemeValueFromSymbol(ArabicData.ArabicLetters(Count), CachedData.IslamData.TranslitSchemes(SchemeCount).Name)
                Next
                Output.Add(Strs)
            End If
        Next
        For Count = 0 To CachedData.ArabicSpecialLetters.Length - 1
            Dim Str As String = System.Text.RegularExpressions.Regex.Replace(SubOutPatterns(CachedData.ArabicSpecialLetters(Count)), "\(?\\u([0-9a-fA-F]{4})\)?", Function(Match As System.Text.RegularExpressions.Match) ChrW(Integer.Parse(Match.Groups(1).Value, Globalization.NumberStyles.HexNumber)))
            Strs = New String() {ArabicLetterSpelling(Str, False), String.Empty, Str, _
                                       TransliterateToScheme(Str, ArabicData.TranslitScheme.Literal, String.Empty)}
            Array.Resize(Of String)(Strs, 3 + CachedData.IslamData.TranslitSchemes.Length)
            For SchemeCount = 0 To CachedData.IslamData.TranslitSchemes.Length - 1
                Strs(3 + SchemeCount) = GetSchemeSpecialValue(CachedData.ArabicSpecialLetters(Count), GetSchemeSpecialFromMatch(CachedData.ArabicSpecialLetters(Count), False), CachedData.IslamData.TranslitSchemes(SchemeCount).Name)
            Next
            Output.Add(Strs)
        Next
        For Count = 0 To CachedData.ArabicLongVowels.Length - 1
            Strs = New String() {ArabicLetterSpelling(CachedData.ArabicLongVowels(Count), False), _
                                       String.Empty, CachedData.ArabicLongVowels(Count), _
                                       TransliterateToScheme(CachedData.ArabicLongVowels(Count), ArabicData.TranslitScheme.Literal, String.Empty)}
            Array.Resize(Of String)(Strs, 3 + CachedData.IslamData.TranslitSchemes.Length)
            For SchemeCount = 0 To CachedData.IslamData.TranslitSchemes.Length - 1
                Strs(3 + SchemeCount) = GetSchemeLongVowelFromString(CachedData.ArabicLongVowels(Count), CachedData.IslamData.TranslitSchemes(SchemeCount).Name)
            Next
            Output.Add(Strs)
        Next
        Return Output.ToArray()
    End Function
    Public Shared Function DisplayParticle(Category As IslamData.GrammarSet.GrammarParticle(), ID As String, SchemeType As ArabicData.TranslitScheme, Scheme As String, ColSels() As String) As Array()
        Dim Count As Integer
        Dim Output(2 + Category.Length) As Array
        If ColSels Is Nothing Then ColSels = {"posspron"}
        Dim Build As New Generic.Dictionary(Of String, Generic.Dictionary(Of String, String))
        Output(0) = New String() {}
        Dim Strings As New List(Of String)
        Strings.AddRange({"arabic", "transliteration", "translation"})
        If Array.IndexOf(ColSels, "posspron") <> -1 Then
            Strings.Add(String.Empty)
        End If
        Strings.Add(String.Empty)
        Output(1) = Strings.ToArray()
        Strings.Clear()
        Strings.AddRange({"Particle", Utility.LoadResourceString("IslamInfo_Transliteration"), Utility.LoadResourceString("IslamInfo_Translation")})
        If Array.IndexOf(ColSels, "posspron") <> -1 Then
            Strings.Add("Prepositional Attached Pronoun")
        End If
        Strings.Add("Grammar Feature")
        Output(2) = Strings.ToArray()
        For Count = 0 To Category.Length - 1
            Dim Objs As New List(Of Object)
            Objs.AddRange({Arabic.TransliterateFromBuckwalter(Category(Count).Text), Arabic.TransliterateToScheme(Arabic.TransliterateFromBuckwalter(Category(Count).Text), If(SchemeType = ArabicData.TranslitScheme.RuleBased, ArabicData.TranslitScheme.LearningMode, SchemeType), Scheme), Utility.LoadResourceString("IslamInfo_" + Category(Count).TranslationID)})
            If Array.IndexOf(ColSels, "posspron") <> -1 Then
                If Array.FindIndex(Utility.DefaultValue(Category(Count).Grammar, String.Empty).Split(","c), Function(S As String) Array.IndexOf(S.Split("|"c), "prep") <> -1) <> -1 Then
                    Objs.Add(DisplayTransform(Arabic.TransliterateFromBuckwalter(Category(Count).Text), GetTransform("posspron"), ID, True, True, SchemeType, Scheme, Nothing))
                Else
                    Objs.Add(String.Empty)
                End If
            End If
            Objs.Add(Utility.DefaultValue(Category(Count).Grammar, String.Empty))
            Output(3 + Count) = Objs.ToArray()
        Next
        Return RenderArray.MakeTableJSFunctions(CType(Output, Array()), ID)
    End Function
    Public Shared Function DisplayPronoun(Category As IslamData.GrammarSet.GrammarNoun(), ID As String, Personal As Boolean, SchemeType As ArabicData.TranslitScheme, Scheme As String, ColSels() As String) As Array()
        Dim Count As Integer
        Dim Cols As String()
        Dim ColVals As String()
        If ColSels Is Nothing Then ColSels = {"p", "d", "s"}
        If Personal Then
            Cols = {"3m", "3f", "2m", "2f", "1m", "1f"}
            ColVals = {"Third Person Masculine", "Third Person Feminine", "Second Person Masculine", "Second Person Feminine", "First Person Masculine", "First Person Feminine"}
        Else
            Cols = {"m", "f"}
            ColVals = {"Masculine", "Feminine"}
        End If
        Dim Output(2 + Cols.Length) As Array
        Dim Build As New Generic.Dictionary(Of String, Generic.Dictionary(Of String, String()))
        Output(0) = New String() {}
        Dim Strings As New List(Of String)
        If Array.IndexOf(ColSels, "p") <> -1 Then
            Strings.AddRange({"arabic", "transliteration", "translation"})
        End If
        If Array.IndexOf(ColSels, "d") <> -1 Then
            Strings.AddRange({"arabic", "transliteration", "translation"})
        End If
        If Array.IndexOf(ColSels, "s") <> -1 Then
            Strings.AddRange({"arabic", "transliteration", "translation"})
        End If
        Strings.Add(String.Empty)
        Output(1) = Strings.ToArray()
        Strings.Clear()
        If Array.IndexOf(ColSels, "p") <> -1 Then
            Strings.AddRange({"Plural " + Utility.LoadResourceString("IslamInfo_Arabic"), "Plural " + Utility.LoadResourceString("IslamInfo_Transliteration"), "Plural " + Utility.LoadResourceString("IslamInfo_Translation")})
        End If
        If Array.IndexOf(ColSels, "d") <> -1 Then
            Strings.AddRange({"Dual " + Utility.LoadResourceString("IslamInfo_Arabic"), "Dual " + Utility.LoadResourceString("IslamInfo_Transliteration"), "Dual " + Utility.LoadResourceString("IslamInfo_Translation")})
        End If
        If Array.IndexOf(ColSels, "s") <> -1 Then
            Strings.AddRange({"Singular " + Utility.LoadResourceString("IslamInfo_Arabic"), "Singular " + Utility.LoadResourceString("IslamInfo_Transliteration"), "Singular " + Utility.LoadResourceString("IslamInfo_Translation")})
        End If
        Strings.Add("Person and Gender")
        Output(2) = Strings.ToArray()
        Strings.Clear()
        For Count = 0 To Category.Length - 1
            Dim Translat As String = Utility.LoadResourceString("IslamInfo_" + Category(Count).TranslationID)
            Array.ForEach(Category(Count).Grammar.Split(","c), Sub(Sets As String) Array.ForEach(Sets.Split("|"c),
                Sub(Str As String)
                    If System.Text.RegularExpressions.Regex.Match(Str, "^(?:" + ArabicData.MakeRegMultiEx(Cols) + ")[pds]$").Success Then
                        Dim Key As String = Str.Chars(0)
                        If Personal Then '"123".Contains(Str.Chars(0))
                            Key += Str.Chars(1)
                        End If
                        If Not Build.ContainsKey(Key) Then
                            Build.Add(Key, New Generic.Dictionary(Of String, String()))
                        End If
                        If Build.Item(Key).ContainsKey(Str.Chars(If(Personal, 2, 1))) Then
                            Build.Item(Key).Item(Str.Chars(If(Personal, 2, 1)))(0) += " " + Arabic.TransliterateFromBuckwalter(Category(Count).Text)
                        Else
                            Build.Item(Key).Add(Str.Chars(If(Personal, 2, 1)), {Arabic.TransliterateFromBuckwalter(Category(Count).Text), Translat})
                        End If
                    End If
                End Sub))
        Next
        For Index = 0 To Cols.Length - 1
            If Build.ContainsKey(Cols(Index)) Then
                Dim Strs(3 * ColSels.Length) As String
                For SubIndex = 0 To ColSels.Length - 1
                    If Not Build(Cols(Index)).ContainsKey(ColSels(SubIndex)) Then Build(Cols(Index)).Add(ColSels(SubIndex), {String.Empty, String.Empty})
                    Strs(3 * SubIndex) = Build(Cols(Index))(ColSels(SubIndex))(0)
                    Strs(3 * SubIndex + 1) = TransliterateToScheme(Build(Cols(Index))(ColSels(SubIndex))(0), If(SchemeType = ArabicData.TranslitScheme.RuleBased, ArabicData.TranslitScheme.LearningMode, SchemeType), Scheme)
                    Strs(3 * SubIndex + 2) = Build(Cols(Index))(ColSels(SubIndex))(1)
                Next
                Strs(3 * ColSels.Length) = ColVals(Index)
                Output(3 + Index) = Strs
            End If
        Next
        Return RenderArray.MakeTableJSFunctions(CType(Output, Array()), ID)
    End Function
    Public Shared Function DisplayTransform(Text As String, Category As IslamData.GrammarSet.GrammarTransform(), ID As String, Personal As Boolean, Noun As Boolean, SchemeType As ArabicData.TranslitScheme, Scheme As String, ColSels As String()) As Array()
        Dim Count As Integer
        Dim Cols As String()
        Dim ColVals As String()
        If ColSels Is Nothing Then ColSels = {"p", "d", "s"}
        If Personal Then
            Cols = {"3m", "3f", "2m", "2f", "1m", "1f"}
            ColVals = {"Third Person Masculine", "Third Person Feminine", "Second Person Masculine", "Second Person Feminine", "First Person Masculine", "First Person Feminine"}
        Else
            Cols = {"m", "f"}
            ColVals = {"Masculine", "Feminine"}
        End If
        Dim Output(2 + Cols.Length) As Array
        Dim Build As New Generic.Dictionary(Of String, Generic.Dictionary(Of String, String()))
        Output(0) = New String() {}
        Dim Strings As New List(Of String)
        If Array.IndexOf(ColSels, "p") <> -1 Then
            Strings.AddRange({"arabic", "transliteration", "translation"})
        End If
        If Array.IndexOf(ColSels, "d") <> -1 Then
            Strings.AddRange({"arabic", "transliteration", "translation"})
        End If
        If Array.IndexOf(ColSels, "s") <> -1 Then
            Strings.AddRange({"arabic", "transliteration", "translation"})
        End If
        Strings.Add(String.Empty)
        Output(1) = Strings.ToArray()
        Strings.Clear()
        If Array.IndexOf(ColSels, "p") <> -1 Then
            Strings.AddRange({"Plural " + Utility.LoadResourceString("IslamInfo_Arabic"), "Plural " + Utility.LoadResourceString("IslamInfo_Transliteration"), "Plural " + Utility.LoadResourceString("IslamInfo_Translation")})
        End If
        If Array.IndexOf(ColSels, "d") <> -1 Then
            Strings.AddRange({"Dual " + Utility.LoadResourceString("IslamInfo_Arabic"), "Dual " + Utility.LoadResourceString("IslamInfo_Transliteration"), "Dual " + Utility.LoadResourceString("IslamInfo_Translation")})
        End If
        If Array.IndexOf(ColSels, "s") <> -1 Then
            Strings.AddRange({"Singular " + Utility.LoadResourceString("IslamInfo_Arabic"), "Singular " + Utility.LoadResourceString("IslamInfo_Transliteration"), "Singular " + Utility.LoadResourceString("IslamInfo_Translation")})
        End If
        Strings.Add("Person and Gender")
        Output(2) = Strings.ToArray()
        Strings.Clear()
        For Count = 0 To Category.Length - 1
            Dim Translat As String = Utility.LoadResourceString("IslamInfo_" + Category(Count).TranslationID)
            Array.ForEach(Category(Count).Grammar.Split(","c), Sub(Sets As String) Array.ForEach(Sets.Split("|"c),
                Sub(Str As String)
                    If Array.FindIndex(Category(Count).Grammar.Split(","c), Function(S As String) Array.IndexOf(S.Split("|"c), If(Noun, "noun", "verb")) <> -1) <> -1 And System.Text.RegularExpressions.Regex.Match(Str, "^(?:" + ArabicData.MakeRegMultiEx(Cols) + ")[pds]$").Success Then
                        Dim Key As String = Str.Chars(0)
                        If Personal Then '"123".Contains(Str.Chars(0))
                            Key += Str.Chars(1)
                        End If
                        If Not Build.ContainsKey(Key) Then
                            Build.Add(Key, New Generic.Dictionary(Of String, String()))
                        End If
                        If Build.Item(Key).ContainsKey(Str.Chars(If(Personal, 2, 1))) Then
                            Build.Item(Key).Item(Str.Chars(If(Personal, 2, 1)))(0) = If(Text = String.Empty, Build.Item(Key).Item(Str.Chars(If(Personal, 2, 1)))(0) + " " + CachedData.TranslateRegEx(Category(Count).Text, False), " " + ApplyTransform({Category(Count)}, Build.Item(Key).Item(Str.Chars(If(Personal, 2, 1)))(0)))
                            Build.Item(Key).Item(Str.Chars(If(Personal, 2, 1)))(1) = Translat
                        Else
                            Build.Item(Key).Add(Str.Chars(If(Personal, 2, 1)), {If(Text = String.Empty, CachedData.TranslateRegEx(Category(Count).Text, False), ApplyTransform({Category(Count)}, Text)), Translat})
                        End If
                    End If
                End Sub))
        Next
        For Index = 0 To Cols.Length - 1
            If Build.ContainsKey(Cols(Index)) Then
                Dim Strs(3 * ColSels.Length) As String
                For SubIndex = 0 To ColSels.Length - 1
                    If Not Build(Cols(Index)).ContainsKey(ColSels(SubIndex)) Then Build(Cols(Index)).Add(ColSels(SubIndex), {String.Empty, String.Empty})
                    Strs(3 * SubIndex) = Build(Cols(Index))(ColSels(SubIndex))(0)
                    Strs(3 * SubIndex + 1) = TransliterateToScheme(Build(Cols(Index))(ColSels(SubIndex))(0), If(SchemeType = ArabicData.TranslitScheme.RuleBased, ArabicData.TranslitScheme.LearningMode, SchemeType), Scheme)
                    Strs(3 * SubIndex + 2) = Build(Cols(Index))(ColSels(SubIndex))(1)
                Next
                Strs(3 * ColSels.Length) = ColVals(Index)
                Output(3 + Index) = Strs
            End If
        Next
        Return RenderArray.MakeTableJSFunctions(CType(Output, Array()), ID)
    End Function
    Public Shared Function GetCatWords(SelArr As String()) As IslamData.GrammarSet.GrammarWord()
        Dim Words As New List(Of IslamData.GrammarSet.GrammarWord)
        For Count = 0 To SelArr.Length - 1
            Dim Word As Nullable(Of IslamData.GrammarSet.GrammarWord)
            Word = GetCatWord(SelArr(Count))
            If Word.HasValue Then Words.Add(Word.Value)
        Next
        Return Words.ToArray()
    End Function
    Public Shared Function GetCatWord(ID As String) As Nullable(Of IslamData.GrammarSet.GrammarWord)
        GetCatWord = New Nullable(Of IslamData.GrammarSet.GrammarWord)
        Dim Particles As IslamData.GrammarSet.GrammarParticle() = GetParticles(ID)
        If Not Particles Is Nothing AndAlso Particles.Length <> 0 Then GetCatWord = New IslamData.GrammarSet.GrammarWord(Particles(0))
        Dim Nouns As IslamData.GrammarSet.GrammarNoun() = GetCatNoun(ID)
        If Not Nouns Is Nothing AndAlso Nouns.Length <> 0 Then GetCatWord = New IslamData.GrammarSet.GrammarWord(Nouns(0))
        Dim Verbs As IslamData.GrammarSet.GrammarVerb() = GetVerb(ID)
        If Not Verbs Is Nothing AndAlso Verbs.Length <> 0 Then GetCatWord = New IslamData.GrammarSet.GrammarWord(Verbs(0))
        Dim Transforms As IslamData.GrammarSet.GrammarTransform() = GetTransform(ID)
        If Not Transforms Is Nothing AndAlso Transforms.Length <> 0 Then GetCatWord = New IslamData.GrammarSet.GrammarWord(Transforms(0))
    End Function
    Public Shared Function DisplayWord(Category As IslamData.GrammarSet.GrammarWord(), ID As String, SchemeType As ArabicData.TranslitScheme, Scheme As String) As Array()
        Dim Count As Integer
        Dim Output(2 + Category.Length) As Array
        Dim Build As New Generic.Dictionary(Of String, Generic.Dictionary(Of String, String))
        Output(0) = New String() {}
        Output(1) = New String() {"arabic", "transliteration", "translation", String.Empty}
        Output(2) = New String() {Utility.LoadResourceString("IslamInfo_Arabic"), Utility.LoadResourceString("IslamInfo_Transliteration"), Utility.LoadResourceString("IslamInfo_Translation"), Utility.LoadResourceString("IslamInfo_Grammar")}
        For Count = 0 To Category.Length - 1
            Output(3 + Count) = New String() {Arabic.TransliterateFromBuckwalter(Category(Count).Text), Arabic.TransliterateToScheme(Arabic.TransliterateFromBuckwalter(Category(Count).Text), If(SchemeType = ArabicData.TranslitScheme.RuleBased, ArabicData.TranslitScheme.LearningMode, SchemeType), Scheme), Utility.LoadResourceString("IslamInfo_" + Category(Count).TranslationID), If(Category(Count).Grammar Is Nothing, String.Empty, Category(Count).Grammar)}
        Next
        Return RenderArray.MakeTableJSFunctions(CType(Output, Array()), ID)
    End Function
    Shared _NounIDs As Dictionary(Of String, List(Of IslamData.GrammarSet.GrammarNoun))
    Public Shared ReadOnly Property NounIDs As Dictionary(Of String, List(Of IslamData.GrammarSet.GrammarNoun))
        Get
            If _NounIDs Is Nothing Then
                _NounIDs = New Dictionary(Of String, List(Of IslamData.GrammarSet.GrammarNoun))
                For Count As Integer = 0 To CachedData.IslamData.Grammar.Nouns.Length - 1
                    If Not _NounIDs.ContainsKey(CachedData.IslamData.Grammar.Nouns(Count).TranslationID) Then
                        _NounIDs.Add(CachedData.IslamData.Grammar.Nouns(Count).TranslationID, New List(Of IslamData.GrammarSet.GrammarNoun) From {CachedData.IslamData.Grammar.Nouns(Count)})
                    Else
                        Debug.Print("Duplicate Noun ID: " + CachedData.IslamData.Grammar.Nouns(Count).TranslationID)
                    End If
                    Dim Noun As IslamData.GrammarSet.GrammarNoun = CachedData.IslamData.Grammar.Nouns(Count)
                    If Not CachedData.IslamData.Grammar.Nouns(Count).Grammar Is Nothing AndAlso CachedData.IslamData.Grammar.Nouns(Count).Grammar.Length <> 0 Then
                        Array.ForEach(CachedData.IslamData.Grammar.Nouns(Count).Grammar.Split(","c),
                            Sub(Str As String)
                                If Not _NounIDs.ContainsKey(Str) Then
                                    _NounIDs.Add(Str, New List(Of IslamData.GrammarSet.GrammarNoun))
                                End If
                                _NounIDs(Str).Add(Noun)
                            End Sub)
                    End If
                Next
            End If
            Return _NounIDs
        End Get
    End Property
    Public Shared Function GetCatNoun(ID As String) As IslamData.GrammarSet.GrammarNoun()
        Return If(NounIDs.ContainsKey(ID), NounIDs(ID).ToArray(), Nothing)
    End Function
    Shared _TransformIDs As Dictionary(Of String, List(Of IslamData.GrammarSet.GrammarTransform))
    Public Shared ReadOnly Property TransformIDs As Dictionary(Of String, List(Of IslamData.GrammarSet.GrammarTransform))
        Get
            If _TransformIDs Is Nothing Then
                _TransformIDs = New Dictionary(Of String, List(Of IslamData.GrammarSet.GrammarTransform))
                For Count As Integer = 0 To CachedData.IslamData.Grammar.Transforms.Length - 1
                    If Not _TransformIDs.ContainsKey(CachedData.IslamData.Grammar.Transforms(Count).TranslationID) Then
                        _TransformIDs.Add(CachedData.IslamData.Grammar.Transforms(Count).TranslationID, New List(Of IslamData.GrammarSet.GrammarTransform) From {CachedData.IslamData.Grammar.Transforms(Count)})
                    Else
                        Debug.Print("Duplicate Transform ID: " + CachedData.IslamData.Grammar.Transforms(Count).TranslationID)
                    End If
                    Dim Transform As IslamData.GrammarSet.GrammarTransform = CachedData.IslamData.Grammar.Transforms(Count)
                    If Not CachedData.IslamData.Grammar.Transforms(Count).Grammar Is Nothing AndAlso CachedData.IslamData.Grammar.Transforms(Count).Grammar.Length <> 0 Then
                        Array.ForEach(CachedData.IslamData.Grammar.Transforms(Count).Grammar.Split(","c),
                            Sub(GroupStr As String)
                                Array.ForEach(GroupStr.Split("|"c),
                                    Sub(Str As String)
                                        If Not _TransformIDs.ContainsKey(Str) Then
                                            _TransformIDs.Add(Str, New List(Of IslamData.GrammarSet.GrammarTransform))
                                        End If
                                        _TransformIDs(Str).Add(Transform)
                                    End Sub)
                            End Sub)
                    End If
                Next
            End If
            Return _TransformIDs
        End Get
    End Property
    Public Shared Function GetTransform(ID As String) As IslamData.GrammarSet.GrammarTransform()
        Return If(TransformIDs.ContainsKey(ID), TransformIDs(ID).ToArray(), Nothing)
    End Function
    Public Shared Function GetTransformMatch(IDs As String()) As IslamData.GrammarSet.GrammarTransform()
        Dim Transforms As IslamData.GrammarSet.GrammarTransform()() = Array.ConvertAll(IDs, Function(ID As String) GetTransform(ID))
        'union all the results together
        Return Array.FindAll(Transforms(0), Function(Tr As IslamData.GrammarSet.GrammarTransform) Array.TrueForAll(Transforms, Function(Trs As IslamData.GrammarSet.GrammarTransform()) Array.IndexOf(Trs, Tr) <> -1))
    End Function
    Shared _ParticleIDs As Dictionary(Of String, List(Of IslamData.GrammarSet.GrammarParticle))
    Public Shared ReadOnly Property ParticleIDs As Dictionary(Of String, List(Of IslamData.GrammarSet.GrammarParticle))
        Get
            If _ParticleIDs Is Nothing Then
                _ParticleIDs = New Dictionary(Of String, List(Of IslamData.GrammarSet.GrammarParticle))
                For Count As Integer = 0 To CachedData.IslamData.Grammar.Particles.Length - 1
                    If Not _ParticleIDs.ContainsKey(CachedData.IslamData.Grammar.Particles(Count).TranslationID) Then
                        _ParticleIDs.Add(CachedData.IslamData.Grammar.Particles(Count).TranslationID, New List(Of IslamData.GrammarSet.GrammarParticle) From {CachedData.IslamData.Grammar.Particles(Count)})
                    Else
                        Debug.Print("Duplicate Particle ID: " + CachedData.IslamData.Grammar.Particles(Count).TranslationID)
                    End If
                    Dim Particle As IslamData.GrammarSet.GrammarParticle = CachedData.IslamData.Grammar.Particles(Count)
                    If Not CachedData.IslamData.Grammar.Particles(Count).Grammar Is Nothing AndAlso CachedData.IslamData.Grammar.Particles(Count).Grammar.Length <> 0 Then
                        Array.ForEach(CachedData.IslamData.Grammar.Particles(Count).Grammar.Split(","c),
                            Sub(Str As String)
                                If Not _ParticleIDs.ContainsKey(Str) Then
                                    _ParticleIDs.Add(Str, New List(Of IslamData.GrammarSet.GrammarParticle))
                                End If
                                _ParticleIDs(Str).Add(Particle)
                            End Sub)
                    End If
                Next
            End If
            Return _ParticleIDs
        End Get
    End Property
    Public Shared Function GetParticles(ID As String) As IslamData.GrammarSet.GrammarParticle()
        Return If(ParticleIDs.ContainsKey(ID), ParticleIDs(ID).ToArray(), Nothing)
    End Function
    Shared _VerbIDs As Dictionary(Of String, List(Of IslamData.GrammarSet.GrammarVerb))
    Public Shared ReadOnly Property VerbIDs As Dictionary(Of String, List(Of IslamData.GrammarSet.GrammarVerb))
        Get
            If _VerbIDs Is Nothing Then
                _VerbIDs = New Dictionary(Of String, List(Of IslamData.GrammarSet.GrammarVerb))
                For Count As Integer = 0 To CachedData.IslamData.Grammar.Verbs.Length - 1
                    If Not _VerbIDs.ContainsKey(CachedData.IslamData.Grammar.Verbs(Count).TranslationID) Then
                        _VerbIDs.Add(CachedData.IslamData.Grammar.Verbs(Count).TranslationID, New List(Of IslamData.GrammarSet.GrammarVerb) From {CachedData.IslamData.Grammar.Verbs(Count)})
                    Else
                        Debug.Print("Duplicate Verb ID: " + CachedData.IslamData.Grammar.Verbs(Count).TranslationID)
                    End If
                    Dim Verb As IslamData.GrammarSet.GrammarVerb = CachedData.IslamData.Grammar.Verbs(Count)
                    If Not CachedData.IslamData.Grammar.Verbs(Count).Grammar Is Nothing AndAlso CachedData.IslamData.Grammar.Verbs(Count).Grammar.Length <> 0 Then
                        Array.ForEach(CachedData.IslamData.Grammar.Verbs(Count).Grammar.Split(","c),
                            Sub(Str As String)
                                If Not _VerbIDs.ContainsKey(Str) Then
                                    _VerbIDs.Add(Str, New List(Of IslamData.GrammarSet.GrammarVerb))
                                End If
                                _VerbIDs(Str).Add(Verb)
                            End Sub)
                    End If
                Next
            End If
            Return _VerbIDs
        End Get
    End Property
    Public Shared Function GetVerb(ID As String) As IslamData.GrammarSet.GrammarVerb()
        Return If(VerbIDs.ContainsKey(ID), VerbIDs(ID).ToArray(), Nothing)
    End Function
    Public Shared Function DisplayProximals(ByVal Item As PageLoader.TextItem) As Array()
        Dim SchemeType As ArabicData.TranslitScheme = CType(If(CInt(HttpContext.Current.Request.QueryString.Get("translitscheme")) >= 2, 2 - CInt(HttpContext.Current.Request.QueryString.Get("translitscheme")) Mod 2, CInt(HttpContext.Current.Request.QueryString.Get("translitscheme"))), ArabicData.TranslitScheme)
        Dim Scheme As String = If(CInt(HttpContext.Current.Request.QueryString.Get("translitscheme")) >= 2, CachedData.IslamData.TranslitSchemes(CInt(HttpContext.Current.Request.QueryString.Get("translitscheme")) \ 2).Name, String.Empty)
        Return DisplayPronoun(GetCatNoun("proxdemo"), Item.Name, False, SchemeType, Scheme, Nothing)
    End Function
    Public Shared Function DisplayDistals(ByVal Item As PageLoader.TextItem) As Array()
        Dim SchemeType As ArabicData.TranslitScheme = CType(If(CInt(HttpContext.Current.Request.QueryString.Get("translitscheme")) >= 2, 2 - CInt(HttpContext.Current.Request.QueryString.Get("translitscheme")) Mod 2, CInt(HttpContext.Current.Request.QueryString.Get("translitscheme"))), ArabicData.TranslitScheme)
        Dim Scheme As String = If(CInt(HttpContext.Current.Request.QueryString.Get("translitscheme")) >= 2, CachedData.IslamData.TranslitSchemes(CInt(HttpContext.Current.Request.QueryString.Get("translitscheme")) \ 2).Name, String.Empty)
        Return DisplayPronoun(GetCatNoun("distdemo"), Item.Name, False, SchemeType, Scheme, Nothing)
    End Function
    Public Shared Function DisplayRelatives(ByVal Item As PageLoader.TextItem) As Array()
        Dim SchemeType As ArabicData.TranslitScheme = CType(If(CInt(HttpContext.Current.Request.QueryString.Get("translitscheme")) >= 2, 2 - CInt(HttpContext.Current.Request.QueryString.Get("translitscheme")) Mod 2, CInt(HttpContext.Current.Request.QueryString.Get("translitscheme"))), ArabicData.TranslitScheme)
        Dim Scheme As String = If(CInt(HttpContext.Current.Request.QueryString.Get("translitscheme")) >= 2, CachedData.IslamData.TranslitSchemes(CInt(HttpContext.Current.Request.QueryString.Get("translitscheme")) \ 2).Name, String.Empty)
        Return DisplayPronoun(GetCatNoun("relpro"), Item.Name, False, SchemeType, Scheme, Nothing)
    End Function
    Public Shared Function DisplayPersonals(ByVal Item As PageLoader.TextItem) As Array()
        Dim SchemeType As ArabicData.TranslitScheme = CType(If(CInt(HttpContext.Current.Request.QueryString.Get("translitscheme")) >= 2, 2 - CInt(HttpContext.Current.Request.QueryString.Get("translitscheme")) Mod 2, CInt(HttpContext.Current.Request.QueryString.Get("translitscheme"))), ArabicData.TranslitScheme)
        Dim Scheme As String = If(CInt(HttpContext.Current.Request.QueryString.Get("translitscheme")) >= 2, CachedData.IslamData.TranslitSchemes(CInt(HttpContext.Current.Request.QueryString.Get("translitscheme")) \ 2).Name, String.Empty)
        Return DisplayPronoun(GetCatNoun("perspro"), Item.Name, True, SchemeType, Scheme, Nothing)
    End Function
    Public Shared Function DisplayDeterminerPersonals(ByVal Item As PageLoader.TextItem) As Array()
        Dim SchemeType As ArabicData.TranslitScheme = CType(If(CInt(HttpContext.Current.Request.QueryString.Get("translitscheme")) >= 2, 2 - CInt(HttpContext.Current.Request.QueryString.Get("translitscheme")) Mod 2, CInt(HttpContext.Current.Request.QueryString.Get("translitscheme"))), ArabicData.TranslitScheme)
        Dim Scheme As String = If(CInt(HttpContext.Current.Request.QueryString.Get("translitscheme")) >= 2, CachedData.IslamData.TranslitSchemes(CInt(HttpContext.Current.Request.QueryString.Get("translitscheme")) \ 2).Name, String.Empty)
        Return DisplayTransform(String.Empty, GetTransform("posspron"), Item.Name, True, True, SchemeType, Scheme, Nothing)
    End Function
    Public Shared Function DisplayPastVerbsFamilyI(ByVal Item As PageLoader.TextItem) As Array()
        Dim SchemeType As ArabicData.TranslitScheme = CType(If(CInt(HttpContext.Current.Request.QueryString.Get("translitscheme")) >= 2, 2 - CInt(HttpContext.Current.Request.QueryString.Get("translitscheme")) Mod 2, CInt(HttpContext.Current.Request.QueryString.Get("translitscheme"))), ArabicData.TranslitScheme)
        Dim Scheme As String = If(CInt(HttpContext.Current.Request.QueryString.Get("translitscheme")) >= 2, CachedData.IslamData.TranslitSchemes(CInt(HttpContext.Current.Request.QueryString.Get("translitscheme")) \ 2).Name, String.Empty)
        Return DisplayTransform(String.Empty, GetTransform("pastverbi"), Item.Name, True, False, SchemeType, Scheme, Nothing)
    End Function
    Public Shared Function DisplayPresentVerbsFamilyI(ByVal Item As PageLoader.TextItem) As Array()
        Dim SchemeType As ArabicData.TranslitScheme = CType(If(CInt(HttpContext.Current.Request.QueryString.Get("translitscheme")) >= 2, 2 - CInt(HttpContext.Current.Request.QueryString.Get("translitscheme")) Mod 2, CInt(HttpContext.Current.Request.QueryString.Get("translitscheme"))), ArabicData.TranslitScheme)
        Dim Scheme As String = If(CInt(HttpContext.Current.Request.QueryString.Get("translitscheme")) >= 2, CachedData.IslamData.TranslitSchemes(CInt(HttpContext.Current.Request.QueryString.Get("translitscheme")) \ 2).Name, String.Empty)
        Return DisplayTransform(String.Empty, GetTransform("presverbi"), Item.Name, True, False, SchemeType, Scheme, Nothing)
    End Function
    Public Shared Function DisplayCommandVerbsFamilyI(ByVal Item As PageLoader.TextItem) As Array()
        Dim SchemeType As ArabicData.TranslitScheme = CType(If(CInt(HttpContext.Current.Request.QueryString.Get("translitscheme")) >= 2, 2 - CInt(HttpContext.Current.Request.QueryString.Get("translitscheme")) Mod 2, CInt(HttpContext.Current.Request.QueryString.Get("translitscheme"))), ArabicData.TranslitScheme)
        Dim Scheme As String = If(CInt(HttpContext.Current.Request.QueryString.Get("translitscheme")) >= 2, CachedData.IslamData.TranslitSchemes(CInt(HttpContext.Current.Request.QueryString.Get("translitscheme")) \ 2).Name, String.Empty)
        Return DisplayTransform(String.Empty, GetTransform("commverbi"), Item.Name, False, False, SchemeType, Scheme, Nothing)
    End Function
    Public Shared Function DisplayResponseParticles(ByVal Item As PageLoader.TextItem) As Array()
        Dim SchemeType As ArabicData.TranslitScheme = CType(If(CInt(HttpContext.Current.Request.QueryString.Get("translitscheme")) >= 2, 2 - CInt(HttpContext.Current.Request.QueryString.Get("translitscheme")) Mod 2, CInt(HttpContext.Current.Request.QueryString.Get("translitscheme"))), ArabicData.TranslitScheme)
        Dim Scheme As String = If(CInt(HttpContext.Current.Request.QueryString.Get("translitscheme")) >= 2, CachedData.IslamData.TranslitSchemes(CInt(HttpContext.Current.Request.QueryString.Get("translitscheme")) \ 2).Name, String.Empty)
        Return DisplayParticle(GetParticles("resp"), Item.Name, SchemeType, Scheme, Nothing)
    End Function
    Public Shared Function DisplayInterogativeParticles(ByVal Item As PageLoader.TextItem) As Array()
        Dim SchemeType As ArabicData.TranslitScheme = CType(If(CInt(HttpContext.Current.Request.QueryString.Get("translitscheme")) >= 2, 2 - CInt(HttpContext.Current.Request.QueryString.Get("translitscheme")) Mod 2, CInt(HttpContext.Current.Request.QueryString.Get("translitscheme"))), ArabicData.TranslitScheme)
        Dim Scheme As String = If(CInt(HttpContext.Current.Request.QueryString.Get("translitscheme")) >= 2, CachedData.IslamData.TranslitSchemes(CInt(HttpContext.Current.Request.QueryString.Get("translitscheme")) \ 2).Name, String.Empty)
        Return DisplayParticle(GetParticles("intg"), Item.Name, SchemeType, Scheme, Nothing)
    End Function
    Public Shared Function DisplayLocationParticles(ByVal Item As PageLoader.TextItem) As Array()
        Dim SchemeType As ArabicData.TranslitScheme = CType(If(CInt(HttpContext.Current.Request.QueryString.Get("translitscheme")) >= 2, 2 - CInt(HttpContext.Current.Request.QueryString.Get("translitscheme")) Mod 2, CInt(HttpContext.Current.Request.QueryString.Get("translitscheme"))), ArabicData.TranslitScheme)
        Dim Scheme As String = If(CInt(HttpContext.Current.Request.QueryString.Get("translitscheme")) >= 2, CachedData.IslamData.TranslitSchemes(CInt(HttpContext.Current.Request.QueryString.Get("translitscheme")) \ 2).Name, String.Empty)
        Return DisplayParticle(GetParticles("loc"), Item.Name, SchemeType, Scheme, Nothing)
    End Function
    Public Shared Function DisplayTimeParticles(ByVal Item As PageLoader.TextItem) As Array()
        Dim SchemeType As ArabicData.TranslitScheme = CType(If(CInt(HttpContext.Current.Request.QueryString.Get("translitscheme")) >= 2, 2 - CInt(HttpContext.Current.Request.QueryString.Get("translitscheme")) Mod 2, CInt(HttpContext.Current.Request.QueryString.Get("translitscheme"))), ArabicData.TranslitScheme)
        Dim Scheme As String = If(CInt(HttpContext.Current.Request.QueryString.Get("translitscheme")) >= 2, CachedData.IslamData.TranslitSchemes(CInt(HttpContext.Current.Request.QueryString.Get("translitscheme")) \ 2).Name, String.Empty)
        Return DisplayParticle(GetParticles("time"), Item.Name, SchemeType, Scheme, Nothing)
    End Function
    Public Shared Function DisplayPrepositions(ByVal Item As PageLoader.TextItem) As Array()
        Dim SchemeType As ArabicData.TranslitScheme = CType(If(CInt(HttpContext.Current.Request.QueryString.Get("translitscheme")) >= 2, 2 - CInt(HttpContext.Current.Request.QueryString.Get("translitscheme")) Mod 2, CInt(HttpContext.Current.Request.QueryString.Get("translitscheme"))), ArabicData.TranslitScheme)
        Dim Scheme As String = If(CInt(HttpContext.Current.Request.QueryString.Get("translitscheme")) >= 2, CachedData.IslamData.TranslitSchemes(CInt(HttpContext.Current.Request.QueryString.Get("translitscheme")) \ 2).Name, String.Empty)
        Return DisplayParticle(GetParticles("prep"), Item.Name, SchemeType, Scheme, Nothing)
    End Function
    Public Shared Function DisplayParticles(ByVal Item As PageLoader.TextItem) As Array()
        Dim SchemeType As ArabicData.TranslitScheme = CType(If(CInt(HttpContext.Current.Request.QueryString.Get("translitscheme")) >= 2, 2 - CInt(HttpContext.Current.Request.QueryString.Get("translitscheme")) Mod 2, CInt(HttpContext.Current.Request.QueryString.Get("translitscheme"))), ArabicData.TranslitScheme)
        Dim Scheme As String = If(CInt(HttpContext.Current.Request.QueryString.Get("translitscheme")) >= 2, CachedData.IslamData.TranslitSchemes(CInt(HttpContext.Current.Request.QueryString.Get("translitscheme")) \ 2).Name, String.Empty)
        Return DisplayParticle(GetParticles("particle"), Item.Name, SchemeType, Scheme, Nothing)
    End Function
    Public Shared Function ApplyTransform(Transforms As IslamData.GrammarSet.GrammarTransform(), Str As String) As String
        Dim Text As String = Str
        For Count = 0 To Transforms.Length - 1
            Text = New System.Text.RegularExpressions.Regex(If(Transforms(Count).From Is Nothing, "$", CachedData.TranslateRegEx(Transforms(Count).From, True))).Replace(Text, CachedData.TranslateRegEx(Transforms(Count).Text, False), 1)
        Next
        Return Text
    End Function
    Public Shared Function NounDisplay(Category As IslamData.GrammarSet.GrammarNoun(), ID As String, SchemeType As ArabicData.TranslitScheme, Scheme As String, ColSels As String()) As Array()
        Dim Count As Integer
        Dim Output(2 + Category.Length) As Array
        Output(0) = New String() {}
        If ColSels Is Nothing Then ColSels = {"deft", "fem", "reladj", "grammar"}
        Dim Strings As New List(Of String)
        Strings.Add(String.Empty)
        If Array.IndexOf(ColSels, "deft") <> -1 Then
            Strings.Add(String.Empty)
        End If
        If Array.IndexOf(ColSels, "fem") <> -1 Then
            Strings.Add(String.Empty)
        End If
        If Array.IndexOf(ColSels, "reladj") <> -1 Then
            Strings.Add(String.Empty)
        End If
        If Array.IndexOf(ColSels, "grammar") <> -1 Then
            Strings.Add(String.Empty)
        End If
        Output(1) = Strings.ToArray()
        Strings.Clear()
        Strings.Add("Base Noun")
        If Array.IndexOf(ColSels, "deft") <> -1 Then
            Strings.Add("Definite")
        End If
        If Array.IndexOf(ColSels, "fem") <> -1 Then
            Strings.Add("Feminine")
        End If
        If Array.IndexOf(ColSels, "reladj") <> -1 Then
            Strings.Add("Relational Adjective")
        End If
        If Array.IndexOf(ColSels, "grammar") <> -1 Then
            Strings.Add("Grammar")
        End If
        Output(2) = Strings.ToArray()
        Strings.Clear()
        For Count = 0 To Category.Length - 1
            Dim Tables As New List(Of Object)
            Tables.Add(DeclineNoun(Category(Count), ID, SchemeType, Scheme, ColSels))
            If Array.IndexOf(ColSels, "deft") <> -1 Then
                Dim Text As String = ApplyTransform(GetTransform("deft"), ApplyTransform(GetTransform("strip"), Arabic.TransliterateFromBuckwalter(Category(Count).Text)))
                Tables.Add(DeclineNoun(New IslamData.GrammarSet.GrammarNoun With {.Text = Text, .Grammar = "flex,def," + If(Array.FindIndex(Utility.DefaultValue(Category(Count).Grammar, String.Empty).Split(","c), Function(S As String) Array.IndexOf(S.Split("|"c), "fs") <> -1) <> -1, "fs", "ms") + If(Array.FindIndex(Utility.DefaultValue(Category(Count).Grammar, String.Empty).Split(","c), Function(S As String) Array.IndexOf(S.Split("|"c), "reladj") <> -1) <> -1, ",reladj", String.Empty), .TranslationID = Category(Count).TranslationID}, ID, SchemeType, Scheme, ColSels))
            End If
            If Array.IndexOf(ColSels, "fem") <> -1 Then
                If Array.FindIndex(Utility.DefaultValue(Category(Count).Grammar, String.Empty).Split(","c), Function(S As String) Array.IndexOf(S.Split("|"c), "fs") <> -1) = -1 Then
                    Dim Text As String = ApplyTransform(GetTransform("fem"), ApplyTransform(GetTransform("strip"), Arabic.TransliterateFromBuckwalter(Category(Count).Text)))
                    Tables.Add(NounDisplay({New IslamData.GrammarSet.GrammarNoun With {.Text = Text, .Grammar = "flex,indef,fs" + If(Array.FindIndex(Utility.DefaultValue(Category(Count).Grammar, String.Empty).Split(","c), Function(S As String) Array.IndexOf(S.Split("|"c), "reladj") <> -1) <> -1, ",reladj", String.Empty), .TranslationID = Category(Count).TranslationID}}, ID, SchemeType, Scheme, ColSels))
                Else
                    Tables.Add(String.Empty)
                End If
            End If
            If Array.IndexOf(ColSels, "reladj") <> -1 Then
                If Array.FindIndex(Utility.DefaultValue(Category(Count).Grammar, String.Empty).Split(","c), Function(S As String) Array.IndexOf(S.Split("|"c), "adj") <> -1 Or Array.IndexOf(S.Split("|"c), "reladj") <> -1 Or Array.IndexOf(S.Split("|"c), "fs") <> -1) = -1 Then
                    Dim Text As String = ApplyTransform(GetTransform("reladj"), ApplyTransform(GetTransform("strip"), Arabic.TransliterateFromBuckwalter(Category(Count).Text)))
                    Tables.Add(NounDisplay({New IslamData.GrammarSet.GrammarNoun With {.Text = Text, .Grammar = "flex,reladj,indef,ms", .TranslationID = Category(Count).TranslationID}}, ID, SchemeType, Scheme, ColSels))
                Else
                    Tables.Add(String.Empty)
                End If
            End If
            If Array.IndexOf(ColSels, "grammar") <> -1 Then
                Tables.Add(Category(Count).Grammar)
            End If
            Output(3 + Count) = Tables.ToArray()
            Tables.Clear()
        Next
        Return RenderArray.MakeTableJSFunctions(CType(Output, Array()), ID)
    End Function
    Public Shared Function DisplayNouns(ByVal Item As PageLoader.TextItem) As Array()
        Dim SchemeType As ArabicData.TranslitScheme = CType(If(CInt(HttpContext.Current.Request.QueryString.Get("translitscheme")) >= 2, 2 - CInt(HttpContext.Current.Request.QueryString.Get("translitscheme")) Mod 2, CInt(HttpContext.Current.Request.QueryString.Get("translitscheme"))), ArabicData.TranslitScheme)
        Dim Scheme As String = If(CInt(HttpContext.Current.Request.QueryString.Get("translitscheme")) >= 2, CachedData.IslamData.TranslitSchemes(CInt(HttpContext.Current.Request.QueryString.Get("translitscheme")) \ 2).Name, String.Empty)
        Return NounDisplay(CachedData.IslamData.Grammar.Nouns, Item.Name, SchemeType, Scheme, Nothing)
    End Function
    Public Shared Function DeclineNoun(Category As IslamData.GrammarSet.GrammarNoun, ID As String, SchemeType As ArabicData.TranslitScheme, Scheme As String, ColSels As String()) As Array()
        Dim Sels As String() = {"nom", "acc", "pos"}
        Dim SelTexts As String() = {"Nominative", "Accusative", "Possessive"}
        Dim Output(2 + Sels.Length) As Array
        If ColSels Is Nothing Then ColSels = {"s", "d", "p", "posspron"}
        Dim HasPoss As Boolean = Not Array.FindIndex(Utility.DefaultValue(Category.Grammar, String.Empty).Split(","c), Function(S As String) Array.IndexOf(S.Split("|"c), "def") <> -1) <> -1 And Array.IndexOf(ColSels, "posspron") <> -1
        Output(0) = New String() {}
        Dim Strings As New List(Of String)
        If Array.IndexOf(ColSels, "s") <> -1 Then
            Strings.AddRange({"arabic", "transliteration", "translation"})
            If HasPoss Then Strings.Add(String.Empty)
        End If
        If Array.IndexOf(ColSels, "d") <> -1 Then
            Strings.AddRange({"arabic", "transliteration", "translation"})
            If HasPoss Then Strings.Add(String.Empty)
        End If
        If Array.IndexOf(ColSels, "p") <> -1 Then
            Strings.AddRange({"arabic", "transliteration", "translation"})
            If HasPoss Then Strings.Add(String.Empty)
        End If
        Strings.Add(String.Empty)
        Output(1) = Strings.ToArray()
        Strings.Clear()
        If Array.IndexOf(ColSels, "s") <> -1 Then
            Strings.AddRange({"Singular " + Utility.LoadResourceString("IslamInfo_Arabic"), "Singular " + Utility.LoadResourceString("IslamInfo_Transliteration"), "Singular " + Utility.LoadResourceString("IslamInfo_Translation")})
            If HasPoss Then Strings.Add("Attached Possessive Pronoun")
        End If
        If Array.IndexOf(ColSels, "d") <> -1 Then
            Strings.AddRange({"Dual " + Utility.LoadResourceString("IslamInfo_Arabic"), "Dual " + Utility.LoadResourceString("IslamInfo_Transliteration"), "Dual " + Utility.LoadResourceString("IslamInfo_Translation")})
            If HasPoss Then Strings.Add("Attached Possessive Pronoun")
        End If
        If Array.IndexOf(ColSels, "p") <> -1 Then
            Strings.AddRange({"Plural " + Utility.LoadResourceString("IslamInfo_Arabic"), "Plural " + Utility.LoadResourceString("IslamInfo_Transliteration"), "Plural " + Utility.LoadResourceString("IslamInfo_Translation")})
            If HasPoss Then Strings.Add("Attached Possessive Pronoun")
        End If
        Strings.Add("Case")
        Output(2) = Strings.ToArray()
        Strings.Clear()
        For Count = 0 To Sels.Length - 1
            Dim Objs As New List(Of Object)
            'Nisbah has a whole slow of suffix possibilities from like -ese or -ism or -ist -ar
            If Array.IndexOf(ColSels, "s") <> -1 Then
                Dim Text As String = ApplyTransform(GetTransformMatch({"flex", Sels(Count), If(Array.FindIndex(Utility.DefaultValue(Category.Grammar, String.Empty).Split(","c), Function(S As String) Array.IndexOf(S.Split("|"c), "def") <> -1) <> -1, "def", "indef"), If(Array.FindIndex(Utility.DefaultValue(Category.Grammar, String.Empty).Split(","c), Function(S As String) Array.IndexOf(S.Split("|"c), "fs") <> -1) <> -1, "fs", "ms")}), ApplyTransform(GetTransform("strip"), Arabic.TransliterateFromBuckwalter(Category.Text)))
                Objs.AddRange({Text, Arabic.TransliterateToScheme(Text, If(SchemeType = ArabicData.TranslitScheme.RuleBased, ArabicData.TranslitScheme.LearningMode, SchemeType), Scheme), If(Array.FindIndex(Utility.DefaultValue(Category.Grammar, String.Empty).Split(","c), Function(S As String) Array.IndexOf(S.Split("|"c), "reladj") <> -1) <> -1, "Relating to ", String.Empty) + If(Array.FindIndex(Utility.DefaultValue(Category.Grammar, String.Empty).Split(","c), Function(S As String) Array.IndexOf(S.Split("|"c), "def") <> -1) <> -1, "The ", String.Empty) + Utility.LoadResourceString("IslamInfo_" + Category.TranslationID) + If(Array.FindIndex(Utility.DefaultValue(Category.Grammar, String.Empty).Split(","c), Function(S As String) Array.IndexOf(S.Split("|"c), "fs") <> -1) <> -1, " Feminine", " Masculine")})
                If HasPoss Then
                    Objs.Add(DisplayTransform(ApplyTransform(GetTransform("constpos"), Text), GetTransform("posspron"), ID, True, True, SchemeType, Scheme, Array.FindAll(ColSels, Function(S As String) Array.IndexOf({"p", "d", "s"}, S) <> -1)))
                End If
            End If
            If Array.IndexOf(ColSels, "d") <> -1 Then
                Dim Text As String = ApplyTransform(GetTransformMatch({"flex", Sels(Count), If(Array.FindIndex(Utility.DefaultValue(Category.Grammar, String.Empty).Split(","c), Function(S As String) Array.IndexOf(S.Split("|"c), "def") <> -1) <> -1, "def", "indef"), If(Array.FindIndex(Utility.DefaultValue(Category.Grammar, String.Empty).Split(","c), Function(S As String) Array.IndexOf(S.Split("|"c), "fs") <> -1) <> -1, "fd", "md")}), ApplyTransform(GetTransform("strip"), Arabic.TransliterateFromBuckwalter(Category.Text)))
                Objs.AddRange({Text, Arabic.TransliterateToScheme(Text, If(SchemeType = ArabicData.TranslitScheme.RuleBased, ArabicData.TranslitScheme.LearningMode, SchemeType), Scheme), If(Array.FindIndex(Utility.DefaultValue(Category.Grammar, String.Empty).Split(","c), Function(S As String) Array.IndexOf(S.Split("|"c), "reladj") <> -1) <> -1, "Relating to ", String.Empty) + If(Array.FindIndex(Utility.DefaultValue(Category.Grammar, String.Empty).Split(","c), Function(S As String) Array.IndexOf(S.Split("|"c), "def") <> -1) <> -1, "The ", String.Empty) + "Two " + Utility.LoadResourceString("IslamInfo_" + Category.TranslationID) + "s" + If(Array.FindIndex(Utility.DefaultValue(Category.Grammar, String.Empty).Split(","c), Function(S As String) Array.IndexOf(S.Split("|"c), "fs") <> -1) <> -1, " Feminine", " Masculine")})
                If HasPoss Then
                    Objs.Add(DisplayTransform(ApplyTransform(GetTransform("constpos"), Text), GetTransform("posspron"), ID, True, True, SchemeType, Scheme, Array.FindAll(ColSels, Function(S As String) Array.IndexOf({"p", "d", "s"}, S) <> -1)))
                End If
            End If
            If Array.IndexOf(ColSels, "p") <> -1 Then
                Dim Text As String = ApplyTransform(GetTransformMatch({"flex", Sels(Count), If(Array.FindIndex(Utility.DefaultValue(Category.Grammar, String.Empty).Split(","c), Function(S As String) Array.IndexOf(S.Split("|"c), "def") <> -1) <> -1, "def", "indef"), If(Array.FindIndex(Utility.DefaultValue(Category.Grammar, String.Empty).Split(","c), Function(S As String) Array.IndexOf(S.Split("|"c), "fs") <> -1) <> -1, "fp", "mp")}), ApplyTransform(GetTransform("strip"), Arabic.TransliterateFromBuckwalter(Category.Text)))
                Objs.AddRange({Text, Arabic.TransliterateToScheme(Text, If(SchemeType = ArabicData.TranslitScheme.RuleBased, ArabicData.TranslitScheme.LearningMode, SchemeType), Scheme), If(Array.FindIndex(Utility.DefaultValue(Category.Grammar, String.Empty).Split(","c), Function(S As String) Array.IndexOf(S.Split("|"c), "reladj") <> -1) <> -1, "Relating to ", String.Empty) + If(Array.FindIndex(Utility.DefaultValue(Category.Grammar, String.Empty).Split(","c), Function(S As String) Array.IndexOf(S.Split("|"c), "def") <> -1) <> -1, "The ", String.Empty) + Utility.LoadResourceString("IslamInfo_" + Category.TranslationID) + "s" + If(Array.FindIndex(Utility.DefaultValue(Category.Grammar, String.Empty).Split(","c), Function(S As String) Array.IndexOf(S.Split("|"c), "fs") <> -1) <> -1, " Feminine", " Masculine")})
                If HasPoss Then
                    Objs.Add(DisplayTransform(ApplyTransform(GetTransform("constpos"), Text), GetTransform("posspron"), ID, True, True, SchemeType, Scheme, Array.FindAll(ColSels, Function(S As String) Array.IndexOf({"p", "d", "s"}, S) <> -1)))
                End If
            End If
            Objs.Add(SelTexts(Count))
            Output(3 + Count) = Objs.ToArray()
            Objs.Clear()
        Next
        Return RenderArray.MakeTableJSFunctions(CType(Output, Array()), ID)
    End Function
    Public Shared Function VerbDisplay(Category As IslamData.GrammarSet.GrammarVerb(), ID As String, SchemeType As ArabicData.TranslitScheme, Scheme As String, ColSels As String()) As Array()
        Dim Count As Integer
        Dim Output(2 + Category.Length) As Array
        If ColSels Is Nothing Then ColSels = {"past", "pres", "comm", "forbid", "pasvpast", "pasvpres", "doernoun", "pasvnoun", "part"}
        Dim Build As New Generic.Dictionary(Of String, Generic.Dictionary(Of String, String))
        Output(0) = New String() {}
        Dim Strings As New List(Of String)
        If Array.IndexOf(ColSels, "past") <> -1 Then
            Strings.AddRange({"arabic", "transliteration", "translation"})
        End If
        If Array.IndexOf(ColSels, "pres") <> -1 Then
            Strings.AddRange({"arabic", "transliteration", "translation"})
        End If
        If Array.IndexOf(ColSels, "comm") <> -1 Then
            Strings.AddRange({"arabic", "transliteration", "translation"})
        End If
        If Array.IndexOf(ColSels, "forbid") <> -1 Then
            Strings.AddRange({"arabic", "transliteration", "translation"})
        End If
        If Array.IndexOf(ColSels, "pasvpast") <> -1 Then
            Strings.AddRange({"arabic", "transliteration", "translation"})
        End If
        If Array.IndexOf(ColSels, "pasvpres") <> -1 Then
            Strings.AddRange({"arabic", "transliteration", "translation"})
        End If
        If Array.IndexOf(ColSels, "doernoun") <> -1 Then
            Strings.AddRange({"arabic", "transliteration", "translation"})
        End If
        If Array.IndexOf(ColSels, "pasvnoun") <> -1 Then
            Strings.AddRange({"arabic", "transliteration", "translation"})
        End If
        If Array.IndexOf(ColSels, "part") <> -1 Then
            Strings.AddRange({"arabic", "transliteration", "translation"})
        End If
        Output(1) = Strings.ToArray()
        Strings.Clear()
        If Array.IndexOf(ColSels, "past") <> -1 Then
            Strings.AddRange({"Past Root", Utility.LoadResourceString("IslamInfo_Transliteration"), Utility.LoadResourceString("IslamInfo_Translation")})
        End If
        If Array.IndexOf(ColSels, "pres") <> -1 Then
            Strings.AddRange({"Present Root", Utility.LoadResourceString("IslamInfo_Transliteration"), Utility.LoadResourceString("IslamInfo_Translation")})
        End If
        If Array.IndexOf(ColSels, "comm") <> -1 Then
            Strings.AddRange({"Command Root", Utility.LoadResourceString("IslamInfo_Transliteration"), Utility.LoadResourceString("IslamInfo_Translation")})
        End If
        If Array.IndexOf(ColSels, "forbid") <> -1 Then
            Strings.AddRange({"Forbidding Root", Utility.LoadResourceString("IslamInfo_Transliteration"), Utility.LoadResourceString("IslamInfo_Translation")})
        End If
        If Array.IndexOf(ColSels, "pasvpast") <> -1 Then
            Strings.AddRange({"Passive Past Root", Utility.LoadResourceString("IslamInfo_Transliteration"), Utility.LoadResourceString("IslamInfo_Translation")})
        End If
        If Array.IndexOf(ColSels, "pasvpres") <> -1 Then
            Strings.AddRange({"Passive Present Root", Utility.LoadResourceString("IslamInfo_Transliteration"), Utility.LoadResourceString("IslamInfo_Translation")})
        End If
        If Array.IndexOf(ColSels, "doernoun") <> -1 Then
            Strings.AddRange({"Verbal Doer", Utility.LoadResourceString("IslamInfo_Transliteration"), Utility.LoadResourceString("IslamInfo_Translation")})
        End If
        If Array.IndexOf(ColSels, "pasvnoun") <> -1 Then
            Strings.AddRange({"Passive Noun", Utility.LoadResourceString("IslamInfo_Transliteration"), Utility.LoadResourceString("IslamInfo_Translation")})
        End If
        If Array.IndexOf(ColSels, "part") <> -1 Then
            Strings.AddRange({"Particles", Utility.LoadResourceString("IslamInfo_Transliteration"), Utility.LoadResourceString("IslamInfo_Translation")})
        End If
        Output(2) = Strings.ToArray()
        Strings.Clear()
        For Count = 0 To Category.Length - 1
            Dim Grammar As String
            Dim Text As String
            Dim Present As String
            Dim Command As String
            Dim Forbidding As String
            Dim PassivePast As String
            Dim PassivePresent As String
            Dim VerbalDoer As String
            Dim PassiveNoun As String
            If (Not Category(Count).Grammar Is Nothing AndAlso Category(Count).Grammar.StartsWith("form=")) Then
                Text = Arabic.TransliterateFromBuckwalter(Category(Count).Grammar.Substring(5).Split(","c)(0).Replace("f", Category(Count).Text.Chars(0)).Replace("E", Category(Count).Text.Chars(1)).Replace("l", Category(Count).Text.Chars(2)))
                Present = Arabic.TransliterateFromBuckwalter(Category(Count).Grammar.Substring(5).Split(","c)(1).Replace("f", Category(Count).Text.Chars(0)).Replace("E", Category(Count).Text.Chars(1)).Replace("l", Category(Count).Text.Chars(2)))
                Command = Arabic.TransliterateFromBuckwalter(GetTransform("VerbTypeICommandYouMasculinePattern")(0).Text.Replace("f", Category(Count).Text.Chars(0)).Replace("E", Category(Count).Text.Chars(1)).Replace("l", Category(Count).Text.Chars(2)).Replace("\1", Category(Count).Grammar.Substring(5).Split(","c)(1).Chars(5)))
                Dim Multi As String() = GetTransform("VerbTypeIForbiddingYouMasculinePattern")(0).Text.Split(" "c)
                Forbidding = Arabic.TransliterateFromBuckwalter(Multi(0) + " " + Multi(1).Replace("f", Category(Count).Text.Chars(0)).Replace("E", Category(Count).Text.Chars(1)).Replace("l", Category(Count).Text.Chars(2)).Replace("\1", Category(Count).Grammar.Substring(5).Split(","c)(1).Chars(5)))
                PassivePast = Arabic.TransliterateFromBuckwalter(GetTransform("VerbTypeIPassivePastHePattern")(0).Text.Replace("f", Category(Count).Text.Chars(0)).Replace("E", Category(Count).Text.Chars(1)).Replace("l", Category(Count).Text.Chars(2)))
                PassivePresent = Arabic.TransliterateFromBuckwalter(GetTransform("VerbTypeIPassivePresentHeMasculinePattern")(0).Text.Replace("f", Category(Count).Text.Chars(0)).Replace("E", Category(Count).Text.Chars(1)).Replace("l", Category(Count).Text.Chars(2)))
                VerbalDoer = Arabic.TransliterateFromBuckwalter(GetTransform("VerbTypeIVerbalDoerPattern")(0).Text.Replace("f", Category(Count).Text.Chars(0)).Replace("E", Category(Count).Text.Chars(1)).Replace("l", Category(Count).Text.Chars(2)))
                PassiveNoun = Arabic.TransliterateFromBuckwalter(GetTransform("VerbTypeIPassiveNounPattern")(0).Text.Replace("f", Category(Count).Text.Chars(0)).Replace("E", Category(Count).Text.Chars(1)).Replace("l", Category(Count).Text.Chars(2)))
                Grammar = String.Empty
            Else
                Text = Arabic.TransliterateFromBuckwalter(Category(Count).Text)
                Present = String.Empty
                Command = String.Empty
                Forbidding = String.Empty
                PassivePast = String.Empty
                PassivePresent = String.Empty
                VerbalDoer = String.Empty
                PassiveNoun = String.Empty
                Grammar = Utility.DefaultValue(Category(Count).Grammar, String.Empty)
            End If
            If Array.IndexOf(ColSels, "past") <> -1 Then
                Strings.AddRange({Text, TransliterateToScheme(Text, If(SchemeType = ArabicData.TranslitScheme.RuleBased, ArabicData.TranslitScheme.LearningMode, SchemeType), Scheme), Utility.LoadResourceString("IslamInfo_" + Category(Count).TranslationID)})
            End If
            If Array.IndexOf(ColSels, "pres") <> -1 Then
                Strings.AddRange({Present, TransliterateToScheme(Present, If(SchemeType = ArabicData.TranslitScheme.RuleBased, ArabicData.TranslitScheme.LearningMode, SchemeType), Scheme), String.Empty})
            End If
            If Array.IndexOf(ColSels, "comm") <> -1 Then
                Strings.AddRange({Command, TransliterateToScheme(Command, If(SchemeType = ArabicData.TranslitScheme.RuleBased, ArabicData.TranslitScheme.LearningMode, SchemeType), Scheme), String.Empty})
            End If
            If Array.IndexOf(ColSels, "forbid") <> -1 Then
                Strings.AddRange({Forbidding, TransliterateToScheme(Forbidding, If(SchemeType = ArabicData.TranslitScheme.RuleBased, ArabicData.TranslitScheme.LearningMode, SchemeType), Scheme), String.Empty})
            End If
            If Array.IndexOf(ColSels, "pasvpast") <> -1 Then
                Strings.AddRange({PassivePast, TransliterateToScheme(PassivePast, If(SchemeType = ArabicData.TranslitScheme.RuleBased, ArabicData.TranslitScheme.LearningMode, SchemeType), Scheme), String.Empty})
            End If
            If Array.IndexOf(ColSels, "pasvpres") <> -1 Then
                Strings.AddRange({PassivePresent, TransliterateToScheme(PassivePresent, If(SchemeType = ArabicData.TranslitScheme.RuleBased, ArabicData.TranslitScheme.LearningMode, SchemeType), Scheme), String.Empty})
            End If
            If Array.IndexOf(ColSels, "doernoun") <> -1 Then
                Strings.AddRange({VerbalDoer, TransliterateToScheme(VerbalDoer, If(SchemeType = ArabicData.TranslitScheme.RuleBased, ArabicData.TranslitScheme.LearningMode, SchemeType), Scheme), String.Empty})
            End If
            If Array.IndexOf(ColSels, "pasvnoun") <> -1 Then
                Strings.AddRange({PassiveNoun, TransliterateToScheme(PassiveNoun, If(SchemeType = ArabicData.TranslitScheme.RuleBased, ArabicData.TranslitScheme.LearningMode, SchemeType), Scheme), String.Empty})
            End If
            If Array.IndexOf(ColSels, "part") <> -1 Then
                Strings.AddRange({Grammar, TransliterateToScheme(Grammar, If(SchemeType = ArabicData.TranslitScheme.RuleBased, ArabicData.TranslitScheme.LearningMode, SchemeType), Scheme), String.Empty})
            End If
            Output(3 + Count) = Strings.ToArray()
            Strings.Clear()
        Next
        Return RenderArray.MakeTableJSFunctions(CType(Output, Array()), ID)
    End Function
    Public Shared Function DisplayVerbs(ByVal Item As PageLoader.TextItem) As Array()
        Dim SchemeType As ArabicData.TranslitScheme = CType(If(CInt(HttpContext.Current.Request.QueryString.Get("translitscheme")) >= 2, 2 - CInt(HttpContext.Current.Request.QueryString.Get("translitscheme")) Mod 2, CInt(HttpContext.Current.Request.QueryString.Get("translitscheme"))), ArabicData.TranslitScheme)
        Dim Scheme As String = If(CInt(HttpContext.Current.Request.QueryString.Get("translitscheme")) >= 2, CachedData.IslamData.TranslitSchemes(CInt(HttpContext.Current.Request.QueryString.Get("translitscheme")) \ 2).Name, String.Empty)
        Return VerbDisplay(CachedData.IslamData.Grammar.Verbs, Item.Name, SchemeType, Scheme, Nothing)
    End Function
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
        Return "function embedFontStyle(fontID) { if (isInArray(embeddedFonts, fontID)) return; embeddedFonts.push(fontID); var font=fontList[fontID]; var style = 'font-family: \'' + font.embed + '\';' + 'src: url(\'/files/' + font.file + '.eot\');' + 'src: local(\'' + font.family + '\'), url(\'/files/' + font.file + ((font.file == 'KFC_naskh') ? '.otf\') format(\'opentype\');' : '.ttf\') format(\'truetype\');'); addStyleSheetRule(newStyleSheet(), '@font-face', style);  }"
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
        Return "function getPrefInstalledFont(type) { var list = fontPrefs[type]; for(var i in list) { if (list.hasOwnProperty(i)) { var fontID = list[i]; if (fontList[fontID].installed) return fontID; } } return 'arial'; }"
    End Function
    Public Shared Function GetCheckInstalledFontsJS() As String
        Return "function checkInstalledFonts() { for (var i in fontList) { if (fontList.hasOwnProperty(i)) { var font = fontList[i]; if (font.family && fontExists(font.family)) font.installed = true; } } }"
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
        "function decreaseFontSize() { rule = findStyleSheetRule('span.arabic'); rule.style.fontSize = Math.max(parseInt(rule.style.fontSize.replace('px', ''), 10) - 1, 1) + 'px'; $('.arabic > img').each(function (i) { this.src = this.src.replace(/Size=(\d+)/g, function (mat, p) { return 'Size=' + (parseInt(p) - 1).toString(); }); }); }"}
    End Function
    Public Shared Function GetFontDefaultSizeJS() As String()
        Return New String() {"javascript: defaultFontSize();", String.Empty, Utility.GetLookupStyleSheetJS(), _
        "function defaultFontSize() { findStyleSheetRule('span.arabic').style.fontSize = '32px'; $('.arabic > img').each(function (i) { this.src = this.src.replace(/Size=(\d+)/g, 'Size=32'); }); }"}
    End Function
    Public Shared Function GetFontBiggerJS() As String()
        Return New String() {"javascript: increaseFontSize();", String.Empty, Utility.GetLookupStyleSheetJS(), _
        "function increaseFontSize() { rule = findStyleSheetRule('span.arabic'); rule.style.fontSize = (parseInt(rule.style.fontSize.replace('px', ''), 10) + 1) + 'px'; $('.arabic > img').each(function (i) { this.src = this.src.replace(/Size=(\d+)/g, function (mat, p) { return 'Size=' + (parseInt(p) + 1).toString(); }); }); }"}
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
    Public Structure ListCategory
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
    <System.Xml.Serialization.XmlArray("lists")> _
    <System.Xml.Serialization.XmlArrayItem("category")> _
    Public Lists() As ListCategory

    Public Structure Phrase
        <System.Xml.Serialization.XmlAttribute("text")> _
        Public Text As String
        <System.Xml.Serialization.XmlAttribute("id")> _
        Public TranslationID As String
    End Structure
    <System.Xml.Serialization.XmlArray("phrases")> _
    <System.Xml.Serialization.XmlArrayItem("word")> _
    Public Phrases() As Phrase

    Public Structure AbbrevWord
        <System.Xml.Serialization.XmlAttribute("text")> _
        Public Text As String
        <System.Xml.Serialization.XmlAttribute("font")> _
        Public Font As String
        <System.Xml.Serialization.XmlAttribute("id")> _
        Public TranslationID As String
    End Structure
    <System.Xml.Serialization.XmlArray("abbreviations")> _
    <System.Xml.Serialization.XmlArrayItem("word")> _
    Public Abbreviations() As AbbrevWord

    Public Structure GrammarSet
        Public Structure GrammarWord
            Public Sub New(Particle As GrammarParticle)
                Text = Particle.Text
                TranslationID = Particle.TranslationID
                Grammar = Particle.Grammar
            End Sub
            Public Sub New(Noun As GrammarNoun)
                Text = Noun.Text
                TranslationID = Noun.TranslationID
                Grammar = Noun.Grammar
            End Sub
            Public Sub New(Verb As GrammarVerb)
                Text = Verb.Text
                TranslationID = Verb.TranslationID
                Grammar = Verb.Grammar
            End Sub
            Public Sub New(Transform As GrammarTransform)
                Text = Transform.Text
                TranslationID = Transform.TranslationID
                Grammar = Transform.Grammar
            End Sub
            Public Text As String
            Public TranslationID As String
            Public Grammar As String
        End Structure
        Public Structure GrammarTransform
            <System.Xml.Serialization.XmlAttribute("text")> _
            Public Text As String
            <System.Xml.Serialization.XmlAttribute("id")> _
            Public TranslationID As String
            <System.Xml.Serialization.XmlAttribute("grammar")> _
            Public Grammar As String
            <System.Xml.Serialization.XmlAttribute("match")> _
            Public Match As String
            <System.Xml.Serialization.XmlAttribute("from")> _
            Public From As String
        End Structure
        <System.Xml.Serialization.XmlArray("transforms")> _
        <System.Xml.Serialization.XmlArrayItem("transform")> _
        Public Transforms() As GrammarTransform
        Public Structure GrammarParticle
            <System.Xml.Serialization.XmlAttribute("text")> _
            Public Text As String
            <System.Xml.Serialization.XmlAttribute("id")> _
            Public TranslationID As String
            <System.Xml.Serialization.XmlAttribute("grammar")> _
            Public Grammar As String
        End Structure
        <System.Xml.Serialization.XmlArray("particles")> _
        <System.Xml.Serialization.XmlArrayItem("particle")> _
        Public Particles() As GrammarParticle
        Public Structure GrammarNoun
            <System.Xml.Serialization.XmlAttribute("text")> _
            Public Text As String
            <System.Xml.Serialization.XmlAttribute("id")> _
            Public TranslationID As String
            <System.Xml.Serialization.XmlAttribute("plural")> _
            Public Plural As String
            <System.Xml.Serialization.XmlAttribute("grammar")> _
            Public Grammar As String
        End Structure
        <System.Xml.Serialization.XmlArray("nouns")> _
        <System.Xml.Serialization.XmlArrayItem("noun")> _
        Public Nouns() As GrammarNoun
        Public Structure GrammarVerb
            <System.Xml.Serialization.XmlAttribute("text")> _
            Public Text As String
            <System.Xml.Serialization.XmlAttribute("id")> _
            Public TranslationID As String
            <System.Xml.Serialization.XmlAttribute("poss")> _
            Public Possessives As String
            <System.Xml.Serialization.XmlAttribute("grammar")> _
            Public Grammar As String
        End Structure
        <System.Xml.Serialization.XmlArray("verbs")> _
        <System.Xml.Serialization.XmlArrayItem("verb")> _
        Public Verbs() As GrammarVerb
    End Structure
    <System.Xml.Serialization.XmlElement("grammar")> _
    Public Grammar As GrammarSet

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
                Return String.Join("|"c, Array.ConvertAll(Hamza, Function(Str As String) Str.Replace("|", "&pipe;")))
            End Get
            Set(value As String)
                If Not value Is Nothing Then
                    Hamza = Array.ConvertAll(value.Split("|"c), Function(Str As String) Str.Replace("&pipe;", "|"))
                End If
            End Set
        End Property
        Public SpecialLetters() As String
        <System.Xml.Serialization.XmlAttribute("literals")> _
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
        Public Multis() As String
        <System.Xml.Serialization.XmlAttribute("multis")> _
        Property MultisParse As String
            Get
                If Multis.Length = 0 Then Return String.Empty
                Return String.Join("|"c, Multis)
            End Get
            Set(value As String)
                If Not value Is Nothing Then
                    Multis = value.Split("|"c)
                End If
            End Set
        End Property
        Public Gutterals() As String
        <System.Xml.Serialization.XmlAttribute("gutterals")> _
        Property GutteralsParse As String
            Get
                If Gutterals.Length = 0 Then Return String.Empty
                Return String.Join("|"c, Gutterals)
            End Get
            Set(value As String)
                If Not value Is Nothing Then
                    Gutterals = value.Split("|"c)
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
        Public Silent() As String
        <System.Xml.Serialization.XmlAttribute("silent")> _
        Property SilentParse As String
            Get
                If Silent.Length = 0 Then Return String.Empty
                Return String.Join("|"c, Silent)
            End Get
            Set(value As String)
                If Not value Is Nothing Then
                    Silent = value.Split("|"c)
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
    Structure ArabicCapInfo
        <System.Xml.Serialization.XmlAttribute("name")> _
        Public Name As String
        <System.Xml.Serialization.XmlAttribute("text")> _
        Public _Text As String
        ReadOnly Property Text As String()
            Get
                Return _Text.Split({"  "}, StringSplitOptions.None)
            End Get
        End Property
    End Structure
    <System.Xml.Serialization.XmlArray("arabiccaptures")> _
    <System.Xml.Serialization.XmlArrayItem("caps")> _
    Public ArabicCaptures() As ArabicCapInfo
    Structure ArabicNumInfo
        <System.Xml.Serialization.XmlAttribute("name")> _
        Public Name As String
        <System.Xml.Serialization.XmlAttribute("text")> _
        Public _Text As String
        ReadOnly Property Text As String()
            Get
                Return _Text.Split(" "c)
            End Get
        End Property
    End Structure
    <System.Xml.Serialization.XmlArray("arabicnumbers")> _
    <System.Xml.Serialization.XmlArrayItem("nums")> _
    Public ArabicNumbers() As ArabicNumInfo
    Structure ArabicPattern
        <System.Xml.Serialization.XmlAttribute("name")> _
        Public Name As String
        <System.Xml.Serialization.XmlAttribute("match")> _
        Public Match As String
    End Structure
    <System.Xml.Serialization.XmlArray("arabicpatterns")> _
    <System.Xml.Serialization.XmlArrayItem("pattern")> _
    Public ArabicPatterns() As ArabicPattern
    Structure ArabicGroup
        <System.Xml.Serialization.XmlAttribute("name")> _
        Public Name As String
        <System.Xml.Serialization.XmlAttribute("text")> _
        Public _Text As String
        ReadOnly Property Text As String()
            Get
                Return _Text.Split(" "c)
            End Get
        End Property
    End Structure
    <System.Xml.Serialization.XmlArray("arabicgroups")> _
    <System.Xml.Serialization.XmlArrayItem("group")> _
    Public ArabicGroups() As ArabicGroup
    Structure ColorRule
        <System.Xml.Serialization.XmlAttribute("name")> _
        Public Name As String
        <System.Xml.Serialization.XmlAttribute("match")> _
        Public Match As String
        <System.Xml.Serialization.XmlAttribute("color")> _
        Public _Color As String
        ReadOnly Property Color As Drawing.Color
            Get
                If _Color.Contains(",") Then
                    Dim RGB As Integer() = Array.ConvertAll(_Color.Split(","c), Function(Str As String) CInt(Str))
                    Return Drawing.Color.FromArgb(RGB(0), RGB(1), RGB(2))
                Else
                    Return Drawing.Color.FromName(_Color)
                End If
            End Get
        End Property
    End Structure
    <System.Xml.Serialization.XmlArray("colorrules")> _
    <System.Xml.Serialization.XmlArrayItem("colorrule")> _
    Public ColorRules() As ColorRule

    Structure RuleTranslationCategory
        Structure RuleTranslation
            <System.Xml.Serialization.XmlAttribute("name")> _
            Public Name As String
            <System.Xml.Serialization.XmlAttribute("match")> _
            Public Match As String
            <System.Xml.Serialization.XmlAttribute("evaluator")> _
            Public Evaluator As String
            <System.Xml.Serialization.XmlAttribute("negativematch")> _
            Public NegativeMatch As String
            <System.Xml.Serialization.XmlAttribute("rulefunc")> _
            Public _RuleFunc As String
            ReadOnly Property RuleFunc As Arabic.RuleFuncs
                Get
                    If _RuleFunc = "eLearningMode" Then Return Arabic.RuleFuncs.eLearningMode
                    If _RuleFunc = "eDivideTanween" Then Return Arabic.RuleFuncs.eDivideTanween
                    If _RuleFunc = "eLeadingGutteral" Then Return Arabic.RuleFuncs.eLeadingGutteral
                    If _RuleFunc = "eLookupLetter" Then Return Arabic.RuleFuncs.eLookupLetter
                    If _RuleFunc = "eLookupLongVowelDipthong" Then Return Arabic.RuleFuncs.eLookupLongVowelDipthong
                    If _RuleFunc = "eSpellLetter" Then Return Arabic.RuleFuncs.eSpellLetter
                    If _RuleFunc = "eSpellNumber" Then Return Arabic.RuleFuncs.eSpellNumber
                    If _RuleFunc = "eTrailingGutteral" Then Return Arabic.RuleFuncs.eTrailingGutteral
                    If _RuleFunc = "eUpperCase" Then Return Arabic.RuleFuncs.eUpperCase
                    If _RuleFunc = "eResolveAmbiguity" Then Return Arabic.RuleFuncs.eResolveAmbiguity
                    'If _RuleFunc = "eNone" Then
                    Return Arabic.RuleFuncs.eNone
                End Get
            End Property
        End Structure
        <System.Xml.Serialization.XmlAttribute("name")> _
        Public Name As String
        <System.Xml.Serialization.XmlElement("rule")> _
        Public Rules() As RuleTranslation
    End Structure
    <System.Xml.Serialization.XmlArray("translitrules")> _
    <System.Xml.Serialization.XmlArrayItem("ruleset")> _
    Public RuleSets() As RuleTranslationCategory
    Structure VerificationData
        <System.Xml.Serialization.XmlAttribute("name")> _
        Public Name As String
        <System.Xml.Serialization.XmlAttribute("match")> _
        Public Match As String
        <System.Xml.Serialization.XmlAttribute("evaluator")> _
        Public _Evaluator As String
        ReadOnly Property Evaluator As String()
            Get
                Return _Evaluator.Split("|"c)
            End Get
        End Property
        <System.Xml.Serialization.XmlAttribute("metarules")> _
        Public _MetaRules As String
        ReadOnly Property MetaRules As String()
            Get
                Return _MetaRules.Split("|"c)
            End Get
        End Property
    End Structure
    <System.Xml.Serialization.XmlArray("verificationset")> _
    <System.Xml.Serialization.XmlArrayItem("verification")> _
    Public VerificationSet() As VerificationData
    Structure RuleMetaSet
        Structure RuleMetadataTranslation
            <System.Xml.Serialization.XmlAttribute("name")> _
            Public Name As String
            <System.Xml.Serialization.XmlAttribute("match")> _
            Public Match As String
            <System.Xml.Serialization.XmlAttribute("evaluator")> _
            Public _Evaluator As String
            ReadOnly Property Evaluator As String()
                Get
                    Return _Evaluator.Split(";"c)
                End Get
            End Property
        End Structure
        <System.Xml.Serialization.XmlAttribute("name")> _
        Public Name As String
        <System.Xml.Serialization.XmlElement("metarule")> _
        Public Rules() As RuleMetadataTranslation
    End Structure
    <System.Xml.Serialization.XmlArray("metarules")> _
    <System.Xml.Serialization.XmlArrayItem("metaruleset")> _
    Public MetaRules() As RuleMetaSet
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
    Structure VerseNumberScheme
        <System.Xml.Serialization.XmlAttribute("name")> _
        Public Name As String
        Public ExtraVerses As Integer()()
        <System.Xml.Serialization.XmlAttribute("extraverses")> _
        Property ExtraVersesStr As String
            Get
                Return String.Join(",", Array.ConvertAll(ExtraVerses, Function(Ints As Integer()) String.Join(":", Array.ConvertAll(Ints, Function(Int As Integer) CStr(Int)))))
            End Get
            Set(value As String)
                If value Is Nothing OrElse value.Length = 0 Then
                    ExtraVerses = Nothing
                Else
                    ExtraVerses = Array.ConvertAll(value.Split(","c), Function(Str As String) Array.ConvertAll(Str.Split(":"c), Function(Verse As String) CInt(Verse)))
                End If
            End Set
        End Property
        Public CombinedVerses As Integer()()
        <System.Xml.Serialization.XmlAttribute("combinedverses")> _
        Property CombinedVersesStr As String
            Get
                Return String.Join(",", Array.ConvertAll(CombinedVerses, Function(Ints As Integer()) String.Join(":", Array.ConvertAll(Ints, Function(Int As Integer) CStr(Int)))))
            End Get
            Set(value As String)
                If value Is Nothing OrElse value.Length = 0 Then
                    CombinedVerses = Nothing
                Else
                    CombinedVerses = Array.ConvertAll(value.Split(","c), Function(Str As String) Array.ConvertAll(Str.Split(":"c), Function(Verse As String) CInt(Verse)))
                End If
            End Set
        End Property
    End Structure
    <System.Xml.Serialization.XmlArray("versenumberschemes")> _
    <System.Xml.Serialization.XmlArrayItem("versenumberscheme")> _
    Public VerseNumberSchemes As VerseNumberScheme()

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
        <System.Xml.Serialization.XmlAttribute("color")> _
        Public _Color As String
        ReadOnly Property Color As Drawing.Color
            Get
                If _Color.Contains(",") Then
                    Dim RGB As Integer() = Array.ConvertAll(_Color.Split(","c), Function(Str As String) CInt(Str))
                    Return Drawing.Color.FromArgb(RGB(0), RGB(1), RGB(2))
                Else
                    Return Drawing.Color.FromName(_Color)
                End If
            End Get
        End Property
    End Structure
    <System.Xml.Serialization.XmlArray("partsofspeech")> _
    <System.Xml.Serialization.XmlArrayItem("pos")> _
    Public PartsOfSpeech() As PartOfSpeechInfo
End Class
Public Class CachedData
    'need disk and memory cache as time consuming to read or build
    Shared _ObjIslamData As IslamData
    Shared _AlDuri As IslamData.RuleTranslationCategory.RuleTranslation()
    Shared _Warsh As IslamData.RuleTranslationCategory.RuleTranslation()
    Shared _WarshScript As IslamData.RuleTranslationCategory.RuleTranslation()
    Shared _UthmaniMinimalScript As IslamData.RuleTranslationCategory.RuleTranslation()
    Shared _SimpleScriptBase As IslamData.RuleTranslationCategory.RuleTranslation()
    Shared _SimpleEnhancedScript As IslamData.RuleTranslationCategory.RuleTranslation()
    Shared _SimpleScript As IslamData.RuleTranslationCategory.RuleTranslation()
    Shared _SimpleCleanScript As IslamData.RuleTranslationCategory.RuleTranslation()
    Shared _SimpleMinimalScript As IslamData.RuleTranslationCategory.RuleTranslation()
    Shared _RomanizationRules As IslamData.RuleTranslationCategory.RuleTranslation()
    Shared _ColoringSpelledOutRules As IslamData.RuleTranslationCategory.RuleTranslation()
    Shared _ErrorCheck As IslamData.RuleTranslationCategory.RuleTranslation()
    Shared _RulesOfRecitationRegEx As IslamData.RuleMetaSet.RuleMetadataTranslation()
    Shared _SavedPatterns As New Dictionary(Of String, String)
    Shared _SavedGroups As New Dictionary(Of String, String())
    Public Shared Function GetNum(Name As String) As String()
        Dim Count As Integer
        For Count = 0 To CachedData.IslamData.ArabicNumbers.Length - 1
            If CachedData.IslamData.ArabicNumbers(Count).Name = Name Then
                Return CachedData.IslamData.ArabicNumbers(Count).Text
            End If
        Next
        Return {}
    End Function
    Public Shared Function GetCap(Name As String) As String()
        Dim Count As Integer
        For Count = 0 To CachedData.IslamData.ArabicCaptures.Length - 1
            If CachedData.IslamData.ArabicCaptures(Count).Name = Name Then
                Return CachedData.IslamData.ArabicCaptures(Count).Text
            End If
        Next
        Return {}
    End Function
    Public Shared Function GetPattern(Name As String) As String
        Dim Count As Integer
        If _SavedPatterns.ContainsKey(Name) Then Return _SavedPatterns(Name)
        For Count = 0 To CachedData.IslamData.ArabicPatterns.Length - 1
            If CachedData.IslamData.ArabicPatterns(Count).Name = Name Then
                'Recursive and may already add pattern
                Dim TRegEx As String = TranslateRegEx(CachedData.IslamData.ArabicPatterns(Count).Match, True)
                If Not _SavedPatterns.ContainsKey(Name) Then _SavedPatterns.Add(Name, TRegEx)
                Return _SavedPatterns(Name)
            End If
        Next
        Return String.Empty
    End Function
    Public Shared Function GetGroup(Name As String) As String()
        Dim Characteristics As String() = {"Audibility", "Whispering", "Weakness", "Moderation", "Strength", "Lowness", "Elevation", "Opening", "Closing", "Restraint", "Fluency", "Vibration", "Inclination", "Repetition", "Whistling", "Diffusion", "Elongation", "Nasal", "Ease"}
        Dim Count As Integer
        If _SavedGroups.ContainsKey(Name) Then Return _SavedGroups(Name)
        For Count = 0 To CachedData.IslamData.ArabicGroups.Length - 1
            If CachedData.IslamData.ArabicGroups(Count).Name = Name Then
                _SavedGroups.Add(Name, Array.ConvertAll(CachedData.IslamData.ArabicGroups(Count).Text, Function(Str As String) TranslateRegEx(Str, Name = "ArabicSpecialLetters" Or Name = "ArabicAssimilateSameWord" Or Name = "ArabicAssimilateAcrossWord" Or Name = "ArabicAssimilateLeenAcrossWord" Or Array.IndexOf(Characteristics, Name) <> -1)))
                Return _SavedGroups(Name)
            End If
        Next
        Return {}
    End Function
    Public Shared Function GetRuleSet(Name As String) As IslamData.RuleTranslationCategory.RuleTranslation()
        Dim Count As Integer
        For Count = 0 To CachedData.IslamData.RuleSets.Length - 1
            If CachedData.IslamData.RuleSets(Count).Name = Name Then
                Dim RuleSet As IslamData.RuleTranslationCategory.RuleTranslation() = CachedData.IslamData.RuleSets(Count).Rules
                For SubCount As Integer = 0 To RuleSet.Length - 1
                    RuleSet(SubCount).Match = TranslateRegEx(RuleSet(SubCount).Match, True)
                    RuleSet(SubCount).Evaluator = TranslateRegEx(RuleSet(SubCount).Evaluator, False)
                Next
                Return RuleSet
            End If
        Next
        Return Nothing
    End Function
    Shared _ArabicUniqueLetters As String()
    Shared _ArabicAlphabet As String()
    Shared _ArabicNumbers As String()
    Shared _ArabicWaslKasraExceptions As String()
    Shared _ArabicBaseNumbers As String()
    Shared _ArabicBaseExtraNumbers As String()
    Shared _ArabicBaseTenNumbers As String()
    Shared _ArabicBaseHundredNumbers As String()
    Shared _ArabicBaseThousandNumbers As String()
    Shared _ArabicBaseMillionNumbers As String()
    Shared _ArabicBaseMilliardNumbers As String()
    Shared _ArabicBaseBillionNumbers As String()
    Shared _ArabicBaseTrillionNumbers As String()
    Shared _ArabicFractionNumbers As String()
    Shared _ArabicOrdinalNumbers As String()
    Shared _ArabicOrdinalExtraNumbers As String()
    Shared _ArabicCombiners As String()
    Shared _QuranHeaders As String()
    Public Shared ReadOnly Property ArabicUniqueLetters As String()
        Get
            If _ArabicUniqueLetters Is Nothing Then
                _ArabicUniqueLetters = GetNum("ArabicUniqueLetters")
            End If
            Return _ArabicUniqueLetters
        End Get
    End Property
    Public Shared ReadOnly Property ArabicAlphabet As String()
        Get
            If _ArabicAlphabet Is Nothing Then
                _ArabicAlphabet = GetNum("ArabicAlphabet")
            End If
            Return _ArabicAlphabet
        End Get
    End Property
    Public Shared ReadOnly Property ArabicNumbers As String()
        Get
            If _ArabicNumbers Is Nothing Then
                _ArabicNumbers = GetNum("ArabicNumbers")
            End If
            Return _ArabicNumbers
        End Get
    End Property
    Public Shared ReadOnly Property ArabicWaslKasraExceptions As String()
        Get
            If _ArabicWaslKasraExceptions Is Nothing Then
                _ArabicWaslKasraExceptions = GetNum("ArabicWaslKasraExceptions")
            End If
            Return _ArabicWaslKasraExceptions
        End Get
    End Property
    Public Shared ReadOnly Property ArabicBaseNumbers As String()
        Get
            If _ArabicBaseNumbers Is Nothing Then
                _ArabicBaseNumbers = GetNum("base")
            End If
            Return _ArabicBaseNumbers
        End Get
    End Property
    Public Shared ReadOnly Property ArabicBaseExtraNumbers As String()
        Get
            If _ArabicBaseExtraNumbers Is Nothing Then
                _ArabicBaseExtraNumbers = GetNum("baseextras")
            End If
            Return _ArabicBaseExtraNumbers
        End Get
    End Property
    Public Shared ReadOnly Property ArabicBaseTenNumbers As String()
        Get
            If _ArabicBaseTenNumbers Is Nothing Then
                _ArabicBaseTenNumbers = GetNum("baseten")
            End If
            Return _ArabicBaseTenNumbers
        End Get
    End Property
    Public Shared ReadOnly Property ArabicBaseHundredNumbers As String()
        Get
            If _ArabicBaseHundredNumbers Is Nothing Then
                _ArabicBaseHundredNumbers = GetNum("basehundred")
            End If
            Return _ArabicBaseHundredNumbers
        End Get
    End Property
    Public Shared ReadOnly Property ArabicBaseThousandNumbers As String()
        Get
            If _ArabicBaseThousandNumbers Is Nothing Then
                _ArabicBaseThousandNumbers = GetNum("thousands")
            End If
            Return _ArabicBaseThousandNumbers
        End Get
    End Property
    Public Shared ReadOnly Property ArabicBaseMillionNumbers As String()
        Get
            If _ArabicBaseMillionNumbers Is Nothing Then
                _ArabicBaseMillionNumbers = GetNum("millions")
            End If
            Return _ArabicBaseMillionNumbers
        End Get
    End Property
    Public Shared ReadOnly Property ArabicBaseMilliardNumbers As String()
        Get
            If _ArabicBaseMilliardNumbers Is Nothing Then
                _ArabicBaseMilliardNumbers = GetNum("milliard")
            End If
            Return _ArabicBaseMilliardNumbers
        End Get
    End Property
    Public Shared ReadOnly Property ArabicBaseBillionNumbers As String()
        Get
            If _ArabicBaseBillionNumbers Is Nothing Then
                _ArabicBaseBillionNumbers = GetNum("billions")
            End If
            Return _ArabicBaseBillionNumbers
        End Get
    End Property
    Public Shared ReadOnly Property ArabicBaseTrillionNumbers As String()
        Get
            If _ArabicBaseTrillionNumbers Is Nothing Then
                _ArabicBaseTrillionNumbers = GetNum("trillions")
            End If
            Return _ArabicBaseTrillionNumbers
        End Get
    End Property
    Public Shared ReadOnly Property ArabicFractionNumbers As String()
        Get
            If _ArabicFractionNumbers Is Nothing Then
                _ArabicFractionNumbers = GetNum("fractions")
            End If
            Return _ArabicFractionNumbers
        End Get
    End Property
    Public Shared ReadOnly Property ArabicOrdinalNumbers As String()
        Get
            If _ArabicOrdinalNumbers Is Nothing Then
                _ArabicOrdinalNumbers = GetNum("ordinals")
            End If
            Return _ArabicOrdinalNumbers
        End Get
    End Property
    Public Shared ReadOnly Property ArabicOrdinalExtraNumbers As String()
        Get
            If _ArabicOrdinalExtraNumbers Is Nothing Then
                _ArabicOrdinalExtraNumbers = GetNum("ordinalextras")
            End If
            Return _ArabicOrdinalExtraNumbers
        End Get
    End Property
    Public Shared ReadOnly Property ArabicCombiners As String()
        Get
            If _ArabicCombiners Is Nothing Then
                _ArabicCombiners = GetNum("combiners")
            End If
            Return _ArabicCombiners
        End Get
    End Property
    Public Shared ReadOnly Property QuranHeaders As String()
        Get
            If _QuranHeaders Is Nothing Then
                _QuranHeaders = GetNum("quranheaders")
            End If
            Return _QuranHeaders
        End Get
    End Property
    Shared _CertainStopPattern As String
    Shared _OptionalPattern As String
    Shared _OptionalPatternNotEndOfAyah As String
    Shared _CertainNotStopPattern As String
    Shared _TehMarbutaStopRule As String
    Shared _TehMarbutaContinueRule As String
    Public Shared ReadOnly Property CertainStopPattern As String
        Get
            If _CertainStopPattern Is Nothing Then
                _CertainStopPattern = GetPattern("CertainStopPattern")
            End If
            Return _CertainStopPattern
        End Get
    End Property
    Public Shared ReadOnly Property OptionalPattern As String
        Get
            If _OptionalPattern Is Nothing Then
                _OptionalPattern = GetPattern("OptionalPattern")
            End If
            Return _OptionalPattern
        End Get
    End Property
    Public Shared ReadOnly Property OptionalPatternNotEndOfAyah As String
        Get
            If _OptionalPatternNotEndOfAyah Is Nothing Then
                _OptionalPatternNotEndOfAyah = GetPattern("OptionalPatternNotEndOfAyah")
            End If
            Return _OptionalPatternNotEndOfAyah
        End Get
    End Property
    Public Shared ReadOnly Property CertainNotStopPattern As String
        Get
            If _CertainNotStopPattern Is Nothing Then
                _CertainNotStopPattern = GetPattern("CertainNotStopPattern")
            End If
            Return _CertainNotStopPattern
        End Get
    End Property
    Public Shared ReadOnly Property TehMarbutaStopRule As String
        Get
            If _TehMarbutaStopRule Is Nothing Then
                _TehMarbutaStopRule = GetPattern("TehMarbutaStopRule")
            End If
            Return _TehMarbutaStopRule
        End Get
    End Property
    Public Shared ReadOnly Property TehMarbutaContinueRule As String
        Get
            If _TehMarbutaContinueRule Is Nothing Then
                _TehMarbutaContinueRule = GetPattern("TehMarbutaContinueRule")
            End If
            Return _TehMarbutaContinueRule
        End Get
    End Property
    Shared _ArabicLongVowels As String()
    Shared _ArabicTanweens As String()
    Shared _ArabicFathaDammaKasra As String()
    Shared _ArabicStopLetters As String()
    Shared _ArabicSpecialGutteral As String()
    Shared _ArabicSpecialLeadingGutteral As String()
    Shared _ArabicPunctuationSymbols As String()
    Shared _ArabicLetters As String()
    Shared _ArabicSunLettersNoLam As String()
    Shared _ArabicSunLetters As String()
    Shared _ArabicMoonLettersNoVowels As String()
    Shared _ArabicMoonLetters As String()
    Shared _RecitationCombiningSymbols As String()
    Shared _RecitationConnectingFollowerSymbols As String()
    Shared _RecitationSymbols As String()
    Shared _ArabicLettersInOrder As String()
    Shared _ArabicSpecialLetters As String()
    Shared _ArabicHamzas As String()
    Shared _ArabicVowels As String()
    Shared _ArabicMultis As String()
    Shared _ArabicTajweed As String()
    Shared _ArabicSilent As String()
    Shared _ArabicPunctuation As String()
    Shared _ArabicNums As String()
    Shared _NonArabicLetters As String()
    Shared _WhitespaceSymbols As String()
    Shared _PunctuationSymbols As String()
    Shared _RecitationDiacritics As String()
    Shared _RecitationLettersDiacritics As String()
    Shared _RecitationSpecialSymbols As String()
    Shared _ArabicLeadingGutterals As String()
    Shared _RecitationLetters As String()
    Shared _ArabicTrailingGutterals As String()
    Shared _RecitationSpecialSymbolsNotStop As String()
    Public Shared ReadOnly Property ArabicLongVowels As String()
        Get
            If _ArabicLongVowels Is Nothing Then
                _ArabicLongVowels = GetGroup("ArabicLongVowels")
            End If
            Return _ArabicLongVowels
        End Get
    End Property
    Public Shared ReadOnly Property ArabicTanweens As String()
        Get
            If _ArabicTanweens Is Nothing Then
                _ArabicTanweens = GetGroup("ArabicTanweens")
            End If
            Return _ArabicTanweens
        End Get
    End Property
    Public Shared ReadOnly Property ArabicFathaDammaKasra As String()
        Get
            If _ArabicFathaDammaKasra Is Nothing Then
                _ArabicFathaDammaKasra = GetGroup("ArabicFathaDammaKasra")
            End If
            Return _ArabicFathaDammaKasra
        End Get
    End Property
    Public Shared ReadOnly Property ArabicStopLetters As String()
        Get
            If _ArabicStopLetters Is Nothing Then
                _ArabicStopLetters = GetGroup("ArabicStopLetters")
            End If
            Return _ArabicStopLetters
        End Get
    End Property
    Public Shared ReadOnly Property ArabicSpecialGutteral As String()
        Get
            If _ArabicSpecialGutteral Is Nothing Then
                _ArabicSpecialGutteral = GetGroup("ArabicSpecialGutteral")
            End If
            Return _ArabicSpecialGutteral
        End Get
    End Property
    Public Shared ReadOnly Property ArabicSpecialLeadingGutteral As String()
        Get
            If _ArabicSpecialLeadingGutteral Is Nothing Then
                _ArabicSpecialLeadingGutteral = GetGroup("ArabicSpecialLeadingGutteral")
            End If
            Return _ArabicSpecialLeadingGutteral
        End Get
    End Property
    Public Shared ReadOnly Property ArabicPunctuationSymbols As String()
        Get
            If _ArabicPunctuationSymbols Is Nothing Then
                _ArabicPunctuationSymbols = GetGroup("ArabicPunctuationSymbols")
            End If
            Return _ArabicPunctuationSymbols
        End Get
    End Property
    Public Shared Function GetLetterCharacteristics(Ch As String) As String()
        Dim ArticulationPoints As String() = {"NasalPassage", "TwoLips", "TwoLipsFromInside", "DeepestPartOfTongue", "UnderTheDeepestPartOfTongue", "EdgeOfTongue", "EdgeOfTongueLowestPart", "UnderTipOfTongue", "CloseToUnderTipOfTongueWithTop", "MiddleOfTongue", "TipOfTongueLetters", "ElevatedGumLetters", "GumLetters", "DeepestPartOfThroat", "MiddleOfThroat", "ClosestPartOfThroat", "ChestInterior"} '"TipLetters"
        Dim Characteristics As String() = {"Audibility", "Whispering", "Weakness", "Moderation", "Strength", "Lowness", "Elevation", "Opening", "Closing", "Restraint", "Fluency", "Vibration", "Inclination", "Repetition", "Whistling", "Diffusion", "Elongation", "Nasal", "Ease"}
        Dim Matches As New List(Of String)
        For Count = 0 To Characteristics.Length - 1
            If Array.IndexOf(GetGroup(Characteristics(Count)), Ch) <> -1 Then Matches.Add(Characteristics(Count))
        Next
        Return Matches.ToArray()
    End Function
    Public Shared Function ArabicLetterCharacteristics(ByVal Item As PageLoader.TextItem) As String
        Dim MutualExclusiveChars As String() = {"Audibility", "Whispering", "Weakness", "Moderation", "Strength", "Lowness", "Elevation", "Opening", "Closing", "Restraint", "Fluency"}
        Dim Lets As New List(Of String)
        Lets.AddRange(Array.ConvertAll(ArabicSunLetters, Function(Str As String) ArabicData.MakeUniRegEx(Str)))
        Lets.AddRange(Array.ConvertAll(ArabicMoonLettersNoVowels, Function(Str As String) ArabicData.MakeUniRegEx(Str)))
        Lets.Add(GetPattern("WawExceptLengthening"))
        Lets.Add(GetPattern("YehExceptLengthening"))
        Lets.Add(ArabicData.MakeUniRegEx(ArabicData.ArabicLetterHamza))
        Dim LetCombs As New List(Of String)
        For Count As Integer = 0 To Lets.Count - 1
            For DupCount As Integer = Count + 1 To Lets.Count - 1
                Dim Chars As String() = GetLetterCharacteristics(Lets(Count))
                Dim DupChars As String() = GetLetterCharacteristics(Lets(DupCount))
                'Intersect to get matches
                Dim Union As String() = Array.FindAll(Chars, Function(Ch As String) Array.IndexOf(DupChars, Ch) <> -1)
                'Subtract the intersection to get differences
                Chars = Array.FindAll(Chars, Function(Ch As String) Array.IndexOf(Union, Ch) = -1)
                DupChars = Array.FindAll(DupChars, Function(Ch As String) Array.IndexOf(Union, Ch) = -1)
                'Mutually exclusive differences counted only once
                Dim MutExc As Integer = Array.FindAll(Chars, Function(Ch As String) Array.IndexOf(MutualExclusiveChars, Ch) <> -1).Length
                LetCombs.Add(ArabicData.LeftToRightEmbedding + (Union.Length - Chars.Length - DupChars.Length + MutExc + 8).ToString("00") + "    " + ArabicData.PopDirectionalFormatting + DocBuilder.GetRegExText(Lets(Count)) + "+" + DocBuilder.GetRegExText(Lets(DupCount)) + ArabicData.LeftToRightEmbedding + "    " + CStr(Union.Length) + "    " + String.Join(", ", Chars) + " <> " + String.Join(", ", DupChars) + ArabicData.PopDirectionalFormatting)
            Next
        Next
        LetCombs.Sort(StringComparer.Ordinal)
        Return String.Join(vbCrLf, LetCombs.ToArray())
    End Function
    Public Shared ReadOnly Property ArabicLetters As String()
        Get
            If _ArabicLetters Is Nothing Then
                _ArabicLetters = GetGroup("ArabicLetters")
            End If
            Return _ArabicLetters
        End Get
    End Property
    Public Shared ReadOnly Property ArabicSunLettersNoLam As String()
        Get
            If _ArabicSunLettersNoLam Is Nothing Then
                _ArabicSunLettersNoLam = GetGroup("ArabicSunLettersNoLam")
            End If
            Return _ArabicSunLettersNoLam
        End Get
    End Property
    Public Shared ReadOnly Property ArabicSunLetters As String()
        Get
            If _ArabicSunLetters Is Nothing Then
                _ArabicSunLetters = GetGroup("ArabicSunLetters")
            End If
            Return _ArabicSunLetters
        End Get
    End Property
    Public Shared ReadOnly Property ArabicMoonLettersNoVowels As String()
        Get
            If _ArabicMoonLettersNoVowels Is Nothing Then
                _ArabicMoonLettersNoVowels = GetGroup("ArabicMoonLettersNoVowels")
            End If
            Return _ArabicMoonLettersNoVowels
        End Get
    End Property
    Public Shared ReadOnly Property ArabicMoonLetters As String()
        Get
            If _ArabicMoonLetters Is Nothing Then
                _ArabicMoonLetters = GetGroup("ArabicMoonLetters")
            End If
            Return _ArabicMoonLetters
        End Get
    End Property
    Public Shared ReadOnly Property RecitationCombiningSymbols As String()
        Get
            If _RecitationCombiningSymbols Is Nothing Then
                _RecitationCombiningSymbols = GetGroup("RecitationCombiningSymbols")
            End If
            Return _RecitationCombiningSymbols
        End Get
    End Property
    Public Shared ReadOnly Property RecitationConnectingFollowerSymbols As String()
        Get
            If _RecitationConnectingFollowerSymbols Is Nothing Then
                _RecitationConnectingFollowerSymbols = GetGroup("RecitationConnectingFollowerSymbols")
            End If
            Return _RecitationConnectingFollowerSymbols
        End Get
    End Property
    Public Shared ReadOnly Property RecitationSymbols As String()
        Get
            If _RecitationSymbols Is Nothing Then
                _RecitationSymbols = GetGroup("RecitationSymbols")
            End If
            Return _RecitationSymbols
        End Get
    End Property
    Public Shared ReadOnly Property ArabicLettersInOrder As String()
        Get
            If _ArabicLettersInOrder Is Nothing Then
                _ArabicLettersInOrder = GetGroup("ArabicLettersInOrder")
            End If
            Return _ArabicLettersInOrder
        End Get
    End Property
    Public Shared ReadOnly Property ArabicSpecialLetters As String()
        Get
            If _ArabicSpecialLetters Is Nothing Then
                _ArabicSpecialLetters = GetGroup("ArabicSpecialLetters")
            End If
            Return _ArabicSpecialLetters
        End Get
    End Property
    Public Shared ReadOnly Property ArabicHamzas As String()
        Get
            If _ArabicHamzas Is Nothing Then
                _ArabicHamzas = GetGroup("ArabicHamzas")
            End If
            Return _ArabicHamzas
        End Get
    End Property
    Public Shared ReadOnly Property ArabicVowels As String()
        Get
            If _ArabicVowels Is Nothing Then
                _ArabicVowels = GetGroup("ArabicVowels")
            End If
            Return _ArabicVowels
        End Get
    End Property
    Public Shared ReadOnly Property ArabicMultis As String()
        Get
            If _ArabicMultis Is Nothing Then
                _ArabicMultis = GetGroup("ArabicMultis")
            End If
            Return _ArabicMultis
        End Get
    End Property
    Public Shared ReadOnly Property ArabicTajweed As String()
        Get
            If _ArabicTajweed Is Nothing Then
                _ArabicTajweed = GetGroup("ArabicTajweed")
            End If
            Return _ArabicTajweed
        End Get
    End Property
    Public Shared ReadOnly Property ArabicSilent As String()
        Get
            If _ArabicSilent Is Nothing Then
                _ArabicSilent = GetGroup("ArabicSilent")
            End If
            Return _ArabicSilent
        End Get
    End Property
    Public Shared ReadOnly Property ArabicPunctuation As String()
        Get
            If _ArabicPunctuation Is Nothing Then
                _ArabicPunctuation = GetGroup("ArabicPunctuation")
            End If
            Return _ArabicPunctuation
        End Get
    End Property
    Public Shared ReadOnly Property ArabicNums As String()
        Get
            If _ArabicNums Is Nothing Then
                _ArabicNums = GetGroup("ArabicNums")
            End If
            Return _ArabicNums
        End Get
    End Property
    Public Shared ReadOnly Property NonArabicLetters As String()
        Get
            If _NonArabicLetters Is Nothing Then
                _NonArabicLetters = GetGroup("NonArabicLetters")
            End If
            Return _NonArabicLetters
        End Get
    End Property
    Public Shared ReadOnly Property WhitespaceSymbols As String()
        Get
            If _WhitespaceSymbols Is Nothing Then
                _WhitespaceSymbols = GetGroup("WhitespaceSymbols")
            End If
            Return _WhitespaceSymbols
        End Get
    End Property
    Public Shared ReadOnly Property PunctuationSymbols As String()
        Get
            If _PunctuationSymbols Is Nothing Then
                _PunctuationSymbols = GetGroup("PunctuationSymbols")
            End If
            Return _PunctuationSymbols
        End Get
    End Property
    Public Shared ReadOnly Property RecitationDiacritics As String()
        Get
            If _RecitationDiacritics Is Nothing Then
                _RecitationDiacritics = GetGroup("RecitationDiacritics")
            End If
            Return _RecitationDiacritics
        End Get
    End Property
    Public Shared ReadOnly Property RecitationLettersDiacritics As String()
        Get
            If _RecitationLettersDiacritics Is Nothing Then
                _RecitationLettersDiacritics = GetGroup("RecitationLettersDiacritics")
            End If
            Return _RecitationLettersDiacritics
        End Get
    End Property
    Public Shared ReadOnly Property RecitationSpecialSymbols As String()
        Get
            If _RecitationSpecialSymbols Is Nothing Then
                _RecitationSpecialSymbols = GetGroup("RecitationSpecialSymbols")
            End If
            Return _RecitationSpecialSymbols
        End Get
    End Property
    Public Shared ReadOnly Property ArabicLeadingGutterals As String()
        Get
            If _ArabicLeadingGutterals Is Nothing Then
                _ArabicLeadingGutterals = GetGroup("ArabicLeadingGutterals")
            End If
            Return _ArabicLeadingGutterals
        End Get
    End Property
    Public Shared ReadOnly Property RecitationLetters As String()
        Get
            If _RecitationLetters Is Nothing Then
                _RecitationLetters = GetGroup("RecitationLetters")
            End If
            Return _RecitationLetters
        End Get
    End Property
    Public Shared ReadOnly Property ArabicTrailingGutterals As String()
        Get
            If _ArabicTrailingGutterals Is Nothing Then
                _ArabicTrailingGutterals = GetGroup("ArabicTrailingGutterals")
            End If
            Return _ArabicTrailingGutterals
        End Get
    End Property
    Public Shared ReadOnly Property RecitationSpecialSymbolsNotStop As String()
        Get
            If _RecitationSpecialSymbolsNotStop Is Nothing Then
                _RecitationSpecialSymbolsNotStop = GetGroup("RecitationSpecialSymbolsNotStop")
            End If
            Return _RecitationSpecialSymbolsNotStop
        End Get
    End Property
    Shared _ArabicCamelCaseDict As Dictionary(Of String, Integer)
    Shared _ArabicComboCamelCaseDict As Dictionary(Of String, Integer)
    Public Shared ReadOnly Property ArabicCamelCaseDict As Dictionary(Of String, Integer)
        Get
            If _ArabicCamelCaseDict Is Nothing Then
                _ArabicCamelCaseDict = New Dictionary(Of String, Integer)
                For Count = 0 To ArabicData.ArabicLetters.Length - 1
                    If Not ArabicData.ArabicLetters(Count).UnicodeName.StartsWith("<") Then _ArabicCamelCaseDict.Add(ArabicData.ToCamelCase(ArabicData.ArabicLetters(Count).UnicodeName), Count)
                Next
            End If
            Return _ArabicCamelCaseDict
        End Get
    End Property
    Public Shared ReadOnly Property ArabicComboCamelCaseDict As Dictionary(Of String, Integer)
        Get
            If _ArabicComboCamelCaseDict Is Nothing Then
                _ArabicComboCamelCaseDict = New Dictionary(Of String, Integer)
                For Count = 0 To ArabicData.ArabicCombos.Length - 1
                    For SubCount = 0 To ArabicData.ArabicCombos(Count).UnicodeName.Length - 1
                        If Not ArabicData.ArabicCombos(Count).UnicodeName(SubCount) Is Nothing AndAlso ArabicData.ArabicCombos(Count).UnicodeName(SubCount).Length <> 0 Then _ArabicComboCamelCaseDict.Add(ArabicData.ToCamelCase(ArabicData.ArabicCombos(Count).UnicodeName(SubCount)), Count)
                    Next
                Next
            End If
            Return _ArabicComboCamelCaseDict
        End Get
    End Property
    Public Shared Function TranslateRegEx(Value As String, bAll As Boolean) As String
        Return System.Text.RegularExpressions.Regex.Replace(Value, "\{(.*?)\}",
            Function(Match As System.Text.RegularExpressions.Match)
                If bAll Then
                    If GetPattern(Match.Groups(1).Value) <> String.Empty Then Return GetPattern(Match.Groups(1).Value)
                    If GetCap(Match.Groups(1).Value.Split(";"c)(0)).Length <> 0 Then
                        Return ArabicData.MakeRegMultiEx(Array.ConvertAll(GetCap(Match.Groups(1).Value.Split(";"c)(0)), Function(Str As String)
                                                                                                                            Dim Strs As String() = Str.Split(" "c)
                                                                                                                            Str = String.Empty
                                                                                                                            Dim CapLimit As Integer = If(Match.Groups(1).Value.Contains(";"), Integer.Parse(Match.Groups(1).Value.Split(";"c)(1)), 0)
                                                                                                                            For StrCount As Integer = 0 To Strs.Length - 1
                                                                                                                                If CapLimit = 0 Then
                                                                                                                                    Str += ArabicData.MakeUniRegEx(Arabic.TransliterateFromBuckwalter(Strs(StrCount)))
                                                                                                                                ElseIf StrCount + 1 <= CapLimit Then
                                                                                                                                    Str += "(" + ArabicData.MakeUniRegEx(Arabic.TransliterateFromBuckwalter(Strs(StrCount))) + ")"
                                                                                                                                ElseIf StrCount + 1 > CapLimit Then
                                                                                                                                    Str += "(?=" + ArabicData.MakeUniRegEx(Arabic.TransliterateFromBuckwalter(Strs(StrCount))) + ")"
                                                                                                                                End If
                                                                                                                            Next
                                                                                                                            Return Str
                                                                                                                        End Function))
                    End If
                    If GetNum(Match.Groups(1).Value).Length <> 0 Then
                        Return ArabicData.MakeRegMultiEx(Array.ConvertAll(GetNum(Match.Groups(1).Value), Function(Str As String) ArabicData.MakeUniRegEx(Arabic.TransliterateFromBuckwalter(Str))))
                    End If
                    If GetGroup(Match.Groups(1).Value).Length <> 0 Then
                        Return ArabicData.MakeRegMultiEx(Array.ConvertAll(GetGroup(Match.Groups(1).Value), Function(Str As String) ArabicData.MakeUniRegEx(Str)))
                    End If
                End If
                If System.Text.RegularExpressions.Regex.Match(Match.Groups(1).Value, "0x([0-9a-fA-F]{4})").Success Then
                    Return If(bAll, ArabicData.MakeUniRegEx(ChrW(Integer.Parse(Match.Groups(1).Value.Substring(2), System.Globalization.NumberStyles.HexNumber))), ChrW(Integer.Parse(Match.Groups(1).Value.Substring(2), System.Globalization.NumberStyles.HexNumber)))
                End If
                If ArabicCamelCaseDict.ContainsKey(Match.Groups(1).Value) Then
                    Return If(bAll, ArabicData.MakeUniRegEx(ArabicData.ArabicLetters(ArabicCamelCaseDict(Match.Groups(1).Value)).Symbol), ArabicData.ArabicLetters(ArabicCamelCaseDict(Match.Groups(1).Value)).Symbol)
                End If
                If ArabicComboCamelCaseDict.ContainsKey(Match.Groups(1).Value) Then
                    Return If(bAll, ArabicData.MakeUniRegEx(If(ArabicData.ArabicCombos(ArabicComboCamelCaseDict(Match.Groups(1).Value)).Shaping.Length = 1, ArabicData.ArabicCombos(ArabicComboCamelCaseDict(Match.Groups(1).Value)).Shaping(0), String.Join(String.Empty, Array.ConvertAll(ArabicData.ArabicCombos(ArabicComboCamelCaseDict(Match.Groups(1).Value)).Symbol, Function(Sym As Char) CStr(Sym))))), If(ArabicData.ArabicCombos(ArabicComboCamelCaseDict(Match.Groups(1).Value)).Shaping.Length = 1, ArabicData.ArabicCombos(ArabicComboCamelCaseDict(Match.Groups(1).Value)).Shaping(0), String.Join(String.Empty, Array.ConvertAll(ArabicData.ArabicCombos(ArabicComboCamelCaseDict(Match.Groups(1).Value)).Symbol, Function(Sym As Char) CStr(Sym)))))
                End If
                '{0} ignore
                If Not IsNumeric(Match.Groups(1).Value) Then Debug.Print("Unknown Group: " + Match.Groups(1).Value)
                Return Match.Value
            End Function)
    End Function
    Public Shared ReadOnly Property RulesOfRecitationRegEx As IslamData.RuleMetaSet.RuleMetadataTranslation()
        Get
            If _RulesOfRecitationRegEx Is Nothing Then
                _RulesOfRecitationRegEx = TanzilReader.GetMetaRuleSet("UthmaniQuran").Rules
                For SubCount As Integer = 0 To _RulesOfRecitationRegEx.Length - 1
                    _RulesOfRecitationRegEx(SubCount).Match = TranslateRegEx(_RulesOfRecitationRegEx(SubCount).Match, True)
                Next
            End If
            Return _RulesOfRecitationRegEx
        End Get
    End Property
    Public Shared ReadOnly Property WarshScript As IslamData.RuleTranslationCategory.RuleTranslation()
        Get
            If _WarshScript Is Nothing Then
                _WarshScript = GetRuleSet("WarshScript")
            End If
            Return _WarshScript
        End Get
    End Property
    Public Shared ReadOnly Property UthmaniMinimalScript As IslamData.RuleTranslationCategory.RuleTranslation()
        Get
            If _UthmaniMinimalScript Is Nothing Then
                _UthmaniMinimalScript = GetRuleSet("UthmaniMinimalScript")
            End If
            Return _UthmaniMinimalScript
        End Get
    End Property
    Public Shared ReadOnly Property SimpleScriptBase As IslamData.RuleTranslationCategory.RuleTranslation()
        Get
            If _SimpleScriptBase Is Nothing Then
                _SimpleScriptBase = GetRuleSet("SimpleScriptBase")
            End If
            Return _SimpleScriptBase
        End Get
    End Property
    Public Shared ReadOnly Property SimpleEnhancedScript As IslamData.RuleTranslationCategory.RuleTranslation()
        Get
            If _SimpleEnhancedScript Is Nothing Then
                _SimpleEnhancedScript = GetRuleSet("SimpleEnhancedScript")
            End If
            Return _SimpleEnhancedScript
        End Get
    End Property
    Public Shared ReadOnly Property SimpleScript As IslamData.RuleTranslationCategory.RuleTranslation()
        Get
            If _SimpleScript Is Nothing Then
                _SimpleScript = GetRuleSet("SimpleScript")
            End If
            Return _SimpleScript
        End Get
    End Property
    Public Shared ReadOnly Property SimpleCleanScript As IslamData.RuleTranslationCategory.RuleTranslation()
        Get
            If _SimpleCleanScript Is Nothing Then
                _SimpleCleanScript = GetRuleSet("SimpleCleanScript")
            End If
            Return _SimpleCleanScript
        End Get
    End Property
    Public Shared ReadOnly Property SimpleMinimalScript As IslamData.RuleTranslationCategory.RuleTranslation()
        Get
            If _SimpleMinimalScript Is Nothing Then
                _SimpleMinimalScript = GetRuleSet("SimpleMinimalScript")
            End If
            Return _SimpleMinimalScript
        End Get
    End Property
    Public Shared ReadOnly Property RomanizationRules As IslamData.RuleTranslationCategory.RuleTranslation()
        Get
            If _RomanizationRules Is Nothing Then
                _RomanizationRules = GetRuleSet("RomanizationRules")
            End If
            Return _RomanizationRules
        End Get
    End Property
    Public Shared ReadOnly Property ColoringSpelledOutRules As IslamData.RuleTranslationCategory.RuleTranslation()
        Get
            If _ColoringSpelledOutRules Is Nothing Then
                _ColoringSpelledOutRules = GetRuleSet("ColoringSpelledOutRules")
            End If
            Return _ColoringSpelledOutRules
        End Get
    End Property
    Public Shared ReadOnly Property ErrorCheckRules As IslamData.RuleTranslationCategory.RuleTranslation()
        Get
            If _ErrorCheck Is Nothing Then
                _ErrorCheck = GetRuleSet("ErrorCheck")
            End If
            Return _ErrorCheck
        End Get
    End Property
    Shared _XMLDocMain As System.Xml.XmlDocument 'Tanzil Quran data
    Shared _XMLDocInfo As System.Xml.XmlDocument 'Tanzil metadata
    Shared _XMLDocInfos As Collections.Generic.List(Of System.Xml.XmlDocument) 'Hadiths
    Shared _RootDictionary As New Generic.Dictionary(Of String, ArrayList)
    Shared _FormDictionary As New Generic.Dictionary(Of String, ArrayList)
    Shared _TagDictionary As New Generic.Dictionary(Of String, Generic.Dictionary(Of String, ArrayList))
    Shared _WordDictionary As New Generic.Dictionary(Of String, ArrayList)
    Shared _RealWordDictionary As New Generic.Dictionary(Of String, ArrayList)
    Shared _LetterDictionary As New Generic.Dictionary(Of Char, Generic.Dictionary(Of String, ArrayList))
    Shared _LetterPreDictionary As New Generic.Dictionary(Of Char, Generic.Dictionary(Of String, ArrayList))
    Shared _LetterSufDictionary As New Generic.Dictionary(Of Char, Generic.Dictionary(Of String, ArrayList))
    Shared _PreDictionary As New Generic.Dictionary(Of String, ArrayList)
    Shared _SufDictionary As New Generic.Dictionary(Of String, ArrayList)
    Shared _LocDictionary As New Generic.Dictionary(Of String, Object())
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
                    _LocDictionary.Add(Pieces(0).TrimStart("("c).TrimEnd(")"c), {Pieces(1), Array.Find(Parts, Function(Str As String) Str = "PREFIX") <> String.Empty, Array.Find(Parts, Function(Str As String) Str = "SUFFIX") <> String.Empty, Pieces(2)})
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
                        If _WordDictionary.Item(Lem).IndexOf(Pieces(1)) = -1 Then _WordDictionary.Item(Lem).Add(Pieces(1))
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
                    For FormCount As Integer = 0 To CachedData.FormDictionary(CStr(CachedData.WordDictionary(FreqArray(SubCount))(RefCount))).Count - 1
                        If (CType(CachedData.FormDictionary(CStr(CachedData.WordDictionary(FreqArray(SubCount))(RefCount)))(FormCount), Integer())(0) = BaseChapter AndAlso _
                            CType(CachedData.FormDictionary(CStr(CachedData.WordDictionary(FreqArray(SubCount))(RefCount)))(FormCount), Integer())(1) >= BaseVerse AndAlso _
                            (BaseChapter <> Chapter OrElse _
                            CType(CachedData.FormDictionary(CStr(CachedData.WordDictionary(FreqArray(SubCount))(RefCount)))(FormCount), Integer())(1) <= Verse)) OrElse _
                            (CType(CachedData.FormDictionary(CStr(CachedData.WordDictionary(FreqArray(SubCount))(RefCount)))(FormCount), Integer())(0) > BaseChapter AndAlso _
                            CType(CachedData.FormDictionary(CStr(CachedData.WordDictionary(FreqArray(SubCount))(RefCount)))(FormCount), Integer())(0) < Chapter) OrElse _
                            (CType(CachedData.FormDictionary(CStr(CachedData.WordDictionary(FreqArray(SubCount))(RefCount)))(FormCount), Integer())(0) = Chapter AndAlso
                            CType(CachedData.FormDictionary(CStr(CachedData.WordDictionary(FreqArray(SubCount))(RefCount)))(FormCount), Integer())(1) <= Verse) Then
                            UniCount += 1
                        End If
                    Next
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
        'Bismillah are not counted in here
        Dim Verses As Collections.Generic.List(Of String())
        Verses = TanzilReader.GetQuranText(CachedData.XMLDocMain, -1, -1, -1, -1)
        For Count As Integer = 0 To Verses.Count - 1
            For SubCount As Integer = 0 To Verses(Count).Length - 1
                Dim Words As String() = Verses(Count)(SubCount).Split(" "c)
                For LetCount As Integer = 0 To Words.Length - 1
                    If Not _RealWordDictionary.ContainsKey(Words(LetCount)) Then _RealWordDictionary.Add(Words(LetCount), Nothing)
                Next
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
        For Count = 0 To IslamData.Lists.Length - 1
            Utility.LoadResourceString("IslamInfo_" + IslamData.Lists(Count).Title)
            If Not IslamData.Lists(Count).Words Is Nothing Then
                For SubCount As Integer = 0 To IslamData.Lists(Count).Words.Length - 1
                    DocBuilder.DoErrorCheckBuckwalterText(IslamData.Lists(Count).Words(SubCount).Text, String.Empty)
                    Utility.LoadResourceString("IslamInfo_" + IslamData.Lists(Count).Words(SubCount).TranslationID)
                Next
            End If
        Next
        For Count = 0 To IslamData.Grammar.Transforms.Length - 1
            Arabic.DoErrorCheck(Arabic.TransliterateFromBuckwalter(IslamData.Grammar.Transforms(Count).Text))
            Utility.LoadResourceString("IslamInfo_" + IslamData.Grammar.Transforms(Count).TranslationID)
        Next
        For Count = 0 To IslamData.Grammar.Nouns.Length - 1
            Arabic.DoErrorCheck(Arabic.TransliterateFromBuckwalter(IslamData.Grammar.Nouns(Count).Text))
            Utility.LoadResourceString("IslamInfo_" + IslamData.Grammar.Nouns(Count).TranslationID)
        Next
        For Count = 0 To IslamData.Grammar.Particles.Length - 1
            Arabic.DoErrorCheck(Arabic.TransliterateFromBuckwalter(IslamData.Grammar.Particles(Count).Text))
            Utility.LoadResourceString("IslamInfo_" + IslamData.Grammar.Particles(Count).TranslationID)
        Next
        For Count = 0 To IslamData.Grammar.Verbs.Length - 1
            Arabic.DoErrorCheck(Arabic.TransliterateFromBuckwalter(IslamData.Grammar.Verbs(Count).Text))
            Utility.LoadResourceString("IslamInfo_" + IslamData.Grammar.Verbs(Count).TranslationID)
        Next
        For Count = 0 To IslamData.Phrases.Length - 1
            DocBuilder.DoErrorCheckBuckwalterText(Arabic.TransliterateFromBuckwalter(IslamData.Phrases(Count).Text), IslamData.Phrases(Count).TranslationID)
        Next
        For Count = 0 To IslamData.Abbreviations.Length - 1
            If Not Phrases.GetPhraseCat(IslamData.Abbreviations(Count).TranslationID).HasValue Then
                Debug.Print("Missing Phrase ID: " + IslamData.Abbreviations(Count).TranslationID)
            End If
        Next
        For Count = 0 To ArabicData.ArabicLetters.Length - 1
            Arabic.DoErrorCheck(Arabic.ArabicLetterSpelling(ArabicData.ArabicLetters(Count).Symbol, False))
            If Not ArabicData.ArabicLetters(Count).UnicodeName.StartsWith("<") Then
                Utility.LoadResourceString("IslamInfo_" + ArabicData.ToCamelCase(ArabicData.ArabicLetters(Count).UnicodeName))
            End If
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
                _XMLDocMain.Load(Utility.GetFilePath("metadata\" + TanzilReader.QuranTextNames(0) + ".xml"))
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
    Public Shared ReadOnly Property RealWordDictionary As Generic.Dictionary(Of String, ArrayList)
        Get
            If _RealWordDictionary.Keys.Count = 0 Then BuildQuranLetterIndex()
            Return _RealWordDictionary
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
    Public Shared ReadOnly Property LocDictionary As Generic.Dictionary(Of String, Object())
        Get
            If _LocDictionary.Keys.Count = 0 Then GetMorphologicalData()
            Return _LocDictionary
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
    Public Shared Function GetListRenderedText(ByVal Item As PageLoader.TextItem) As RenderArray
        Dim SchemeType As ArabicData.TranslitScheme = CType(If(CInt(HttpContext.Current.Request.QueryString.Get("translitscheme")) >= 2, 2 - CInt(HttpContext.Current.Request.QueryString.Get("translitscheme")) Mod 2, CInt(HttpContext.Current.Request.QueryString.Get("translitscheme"))), ArabicData.TranslitScheme)
        Dim Scheme As String = If(CInt(HttpContext.Current.Request.QueryString.Get("translitscheme")) >= 2, CachedData.IslamData.TranslitSchemes(CInt(HttpContext.Current.Request.QueryString.Get("translitscheme")) \ 2).Name, String.Empty)
        Dim Count As Integer = CInt(HttpContext.Current.Request.QueryString.Get("selection"))
        If Count = -1 Then Count = 0
        Dim Renderer As New RenderArray(Item.Name)
        If Not CachedData.IslamData.Lists(Count).Words Is Nothing Then
            For SubCount = 0 To CachedData.IslamData.Lists(Count).Words.Length - 1
                Renderer.Items.AddRange(BuckwalterTextFromReferences(Item.Name, SchemeType, Scheme, CachedData.IslamData.Lists(Count).Words(SubCount).Text, String.Empty, TanzilReader.GetTranslationIndex(HttpContext.Current.Request.QueryString.Get("qurantranslation"))).Items)
            Next
        End If
        Return Renderer
    End Function
    Public Shared Function GetRenderedText(ByVal Item As PageLoader.TextItem) As RenderArray
        Dim SchemeType As ArabicData.TranslitScheme = CType(If(CInt(HttpContext.Current.Request.Params("translitscheme")) >= 2, 2 - CInt(HttpContext.Current.Request.Params("translitscheme")) Mod 2, CInt(HttpContext.Current.Request.Params("translitscheme"))), ArabicData.TranslitScheme)
        Dim Scheme As String = If(CInt(HttpContext.Current.Request.Params("translitscheme")) >= 2, CachedData.IslamData.TranslitSchemes(CInt(HttpContext.Current.Request.Params("translitscheme")) \ 2).Name, String.Empty)
        Return NormalTextFromReferences(Item.Name, HttpContext.Current.Request.Params("docedit"), SchemeType, Scheme, TanzilReader.GetTranslationIndex(HttpContext.Current.Request.Params("qurantranslation")))
    End Function
    Public Structure XYZColor
        Public X As Double
        Public Y As Double
        Public Z As Double
    End Structure
    Public Shared WhiteReference As New XYZColor With {.X = 95.047, .Y = 100.0, .Z = 108.883}
    Public Const Epsilon As Double = 0.008856 'Intent is 216/24389
    Public Const Kappa As Double = 903.3 'Intent is 24389/27
    Public Shared Function PivotRGB(N As Double) As Double
        Return If(N > 0.04045, Math.Pow((N + 0.055) / 1.055, 2.4), N / 12.92) * 100.0
    End Function
    Public Shared Function ToRGB(N As Double) As Double
        Dim Result As Double = N * 255.0
        If Result < 0 Then Return 0
        If Result > 255 Then Return 255
        Return Result
    End Function
    Public Shared Function RGBToXYZ(clr As Color) As XYZColor
        Dim r As Double = PivotRGB(clr.R / 255.0)
        Dim g As Double = PivotRGB(clr.G / 255.0)
        Dim b As Double = PivotRGB(clr.B / 255.0)
        Return New XYZColor With {.X = r * 0.4124 + g * 0.3576 + b * 0.1805, _
                                  .Y = r * 0.2126 + g * 0.7152 + b * 0.0722, _
                                  .Z = r * 0.0193 + g * 0.1192 + b * 0.9505}
    End Function
    Public Shared Function XYZToRGB(clr As XYZColor) As Color
        Dim x As Double = clr.X / 100.0
        Dim y As Double = clr.Y / 100.0
        Dim z As Double = clr.Z / 100.0
        Dim r As Double = x * 3.2406 + y * -1.5372 + z * -0.4986
        Dim g As Double = x * -0.9689 + y * 1.8758 + z * 0.0415
        Dim b As Double = x * 0.0557 + y * -0.204 + z * 1.057
        r = If(r > 0.0031308, 1.055 * Math.Pow(r, 1 / 2.4) - 0.055, 12.92 * r)
        g = If(g > 0.0031308, 1.055 * Math.Pow(g, 1 / 2.4) - 0.055, 12.92 * g)
        b = If(b > 0.0031308, 1.055 * Math.Pow(b, 1 / 2.4) - 0.055, 12.92 * b)
        Return Color.FromArgb(CInt(ToRGB(r)), CInt(ToRGB(g)), CInt(ToRGB(b)))
    End Function
    Public Structure LABColor
        Public L As Double
        Public A As Double
        Public B As Double
    End Structure
    Public Shared Function PivotXYZ(N As Double) As Double
        Return If(N > Epsilon, Math.Pow(N, 1.0 / 3.0), (Kappa * N + 16) / 116)
    End Function
    Public Shared Function RGBToLAB(clr As Color) As LABColor
        Dim XYZCol As XYZColor = RGBToXYZ(clr)
        Dim x As Double = PivotXYZ(XYZCol.X / WhiteReference.X)
        Dim y As Double = PivotXYZ(XYZCol.Y / WhiteReference.Y)
        Dim z As Double = PivotXYZ(XYZCol.Z / WhiteReference.Z)
        Return New LABColor With {.L = Math.Max(0, 116 * y - 16), .A = 500 * (x - y), .B = 200 * (y - z)}
    End Function
    Public Shared Function LABToRGB(clr As LABColor) As Color
        Dim y As Double = (clr.L + 16.0) / 116.0
        Dim x As Double = clr.A / 500.0 + y
        Dim z As Double = y - clr.B / 200.0
        Dim X3 As Double = x * x * x
        Dim Z3 As Double = z * z * z
        Return XYZToRGB(New XYZColor With {.X = WhiteReference.X * If(X3 > Epsilon, X3, (x - 16.0 / 116.0) / 7.787), _
                                         .Y = WhiteReference.Y * If(clr.L > (Kappa * Epsilon), Math.Pow((clr.L + 16.0) / 116.0, 3), clr.L / Kappa), _
                                         .Z = WhiteReference.Z * If(Z3 > Epsilon, Z3, (z - 16.0 / 116.0) / 7.787)})
    End Function
    Public Shared Function CMCCompareColors(ColorA As LABColor, ColorB As LABColor, Lightness As Double, Chroma As Double) As Double
        Dim deltaL As Double = ColorA.L - ColorB.L
        Dim h As Double = Math.Atan2(ColorB.B, ColorA.A)
        Dim C1 As Double = Math.Sqrt(ColorA.A * ColorA.A + ColorA.B * ColorA.B)
        Dim C2 As Double = Math.Sqrt(ColorB.A * ColorB.A + ColorB.B * ColorB.B)
        Dim deltaC As Double = C1 - C2
        Dim deltaH As Double = Math.Sqrt((ColorA.A - ColorB.A) * (ColorA.A - ColorB.A) + (ColorA.B - ColorB.B) * (ColorA.B - ColorB.B) - deltaC * deltaC)
        Dim C1_4 As Double = C1 * C1
        C1_4 *= C1_4
        Dim t As Double = If(164 <= h Or h >= 345, 0.56 + Math.Abs(0.2 * Math.Cos(h + 168.0)), 0.36 + Math.Abs(0.4 * Math.Cos(h + 35.0)))
        Dim f As Double = Math.Sqrt(C1_4 / (C1_4 + 1900.0))
        Dim sL As Double = If(ColorA.L < 16, 0.511, (0.040975 * ColorA.L) / (1.0 + 0.01765 * ColorA.L))
        Dim sC As Double = (0.0638 * C1) / (1 + 0.0131 * C1) + 0.638
        Dim sH As Double = sC * (f * t + 1 - f)
        Return Math.Sqrt(deltaL * deltaL / (Lightness * Lightness * sL * sL) + deltaC * deltaC / (Chroma * Chroma * sC * sC) + deltaH * deltaH / (sH * sH))
    End Function
    Public Shared Function GCD(A As Integer, B As Integer) As Integer 'Euclid's algorithm
        If B = 0 Then Return A
        Return GCD(B, A Mod B)
    End Function
    Public Shared Function GenerateNDistinctColors(N As Integer, Threshold As Integer, Interleave As Integer) As Color()
        'To best support individuals with colorblindness (deuteranopia or protanopia) keep a set to 0; vary only L and b.
        Dim LABColors As New List(Of LABColor)
        Dim LowThresholds As New List(Of Double)
        LABColors.Add(RGBToLAB(Color.Black)) 'Start with pivot forecolor
        LowThresholds.Add(100)
        LABColors.Add(RGBToLAB(Color.White)) 'Start with background color
        LowThresholds.Add(100)
        For A = 0 To 200 'Pivot around 0 and move towards 100/-100
            For L = 0 To 100 / 2 'dark to light yet for readability do not exceed half of the spectrum
                For B = 0 To 200 'Pivot around 0 and move towards 100/-100
                    Dim CurColCount As Integer
                    Dim LowThreshold As Double = 100
                    For CurColCount = 0 To LABColors.Count - 1
                        LowThreshold = Math.Min(LowThreshold, CMCCompareColors(LABColors(CurColCount), New LABColor With {.L = L, .A = ((A \ 2) + If((A Mod 2) = 1, 1, 0)) * If((A Mod 2) = 1, 1, -1), .B = ((B \ 2) + If((B Mod 2) = 1, 1, 0)) * If((B Mod 2) = 1, 1, -1)}, 1.0, 1.0))
                        If LowThreshold < Threshold Then Exit For
                    Next
                    If CurColCount = LABColors.Count Then
                        Dim Idx As Integer = LowThresholds.BinarySearch(LowThreshold)
                        If Idx < 0 Then Idx = Idx Xor -1
                        If Idx <> 0 Or LowThresholds.Count <> N + 1 Then
                            LABColors.Insert(Idx, New LABColor With {.L = L, .A = ((A \ 2) + If((A Mod 2) = 1, 1, 0)) * If((A Mod 2) = 1, 1, -1), .B = ((B \ 2) + If((B Mod 2) = 1, 1, 0)) * If((B Mod 2) = 1, 1, -1)})
                            LowThresholds.Insert(Idx, LowThreshold)
                            If LowThresholds.Count > N + 1 Then
                                LABColors.RemoveAt(0)
                                LowThresholds.RemoveAt(0)
                            End If
                        End If
                    End If
                Next
            Next
            If LABColors.Count >= N + 1 Then Exit For
        Next
        LABColors.RemoveAt(N) 'Remove background color
        'if less than N colors found then try with lower threshold
        If LABColors.Count < N Then Return GenerateNDistinctColors(N, Threshold - 1, Interleave)
        Dim Cols(N - 1) As Color
        For Count As Integer = 0 To N - 1
            'Least Common Multiple LCM = a * b \ GCD(A, B)
            Cols((Count * Interleave) \ (N * Interleave \ GCD(N, Interleave)) + (Count * Interleave) Mod N) = LABToRGB(LABColors(Count))
        Next
        Return Cols
    End Function
    Public Shared Function ColorizeList(Strs As String(), bArabic As Boolean) As RenderArray.RenderText()
        Dim Cols As Color() = GenerateNDistinctColors(Strs.Length + 1, 15, 5)
        Dim Renderers As New List(Of RenderArray.RenderText)
        For Count As Integer = 0 To Strs.Length - 1
            Renderers.Add(New RenderArray.RenderText(If(bArabic, RenderArray.RenderDisplayClass.eArabic, RenderArray.RenderDisplayClass.eLTR), Strs(Count) + If(Not bArabic And (Strs(Count) = String.Empty Or Strs(Count).StartsWith(ArabicData.LeftToRightEmbedding)), "NULL" + CStr(Count), "_" + CStr(Count))) With {.Clr = Cols(Count + 1)})
            'Renderers.Add(New RenderArray.RenderText(RenderArray.RenderDisplayClass.eLTR, CStr(Count)) With {.Clr = Color.DarkRed})
            If Not bArabic And Count <> Strs.Length - 1 Then Renderers.Add(New RenderArray.RenderText(RenderArray.RenderDisplayClass.eLTR, ";"))
        Next
        Return Renderers.ToArray()
    End Function
    Public Shared Function ColorizeRegExGroups(Str As String, bReplaceGroup As Boolean) As RenderArray.RenderText()
        'Define an numeric partial ordering of groups that is non-contiguous based on their nearest parent parenthesis
        Dim ParenPos As New List(Of Integer)
        For Count As Integer = 0 To Str.Length - 1
            ParenPos.Add(0)
        Next
        Dim CurNum As Integer = 0
        Dim NumStack As New Stack(Of Integer())
        Dim Matches As System.Text.RegularExpressions.MatchCollection = System.Text.RegularExpressions.Regex.Matches(Str, If(bReplaceGroup, "(\\\$)?(\$\d+)", "(\\\(|\\\))?(\(\??|\))"))
        For MatchCount As Integer = 0 To Matches.Count - 1
            If Matches(MatchCount).Groups(2).Value.Chars(0) = "$" Then
                Dim Num As Integer = Integer.Parse(Matches(MatchCount).Groups(2).Value.Substring(1))
                CurNum = Math.Max(CurNum, Num)
                For Count As Integer = Matches(MatchCount).Groups(2).Index To Matches(MatchCount).Groups(2).Index + Matches(MatchCount).Groups(2).Length - 1
                    ParenPos(Count) = Num
                Next
            ElseIf Matches(MatchCount).Groups(2).Value.Chars(0) = "("c Then
                If Matches(MatchCount).Groups(2).Value.Length = 1 Then CurNum += 1
                NumStack.Push(New Integer() {CurNum, If(Matches(MatchCount).Groups(2).Value.Length = 1, Matches(MatchCount).Groups(2).Index, Str.Length)})
            Else
                Debug.Assert(NumStack.Count <> 0) 'Misbalance parenthesis is exception
                Dim Nums As Integer() = NumStack.Pop()
                For Count As Integer = Nums(1) To Matches(MatchCount).Groups(2).Index
                    If Nums(0) > ParenPos(Count) Then ParenPos(Count) = Nums(0)
                Next
            End If
        Next
        Debug.Assert(NumStack.Count = 0) 'Misbalance parenthesis is exception
        'Proper coloring requires that parent-child and neighboring siblings have different colors
        'yet the current partial ordering does not define either of those relationships
        'must maintain neighbor and color list to properly color
        Dim Base As Integer = 0
        Dim Cols As Color() = GenerateNDistinctColors(CurNum + 1, 15, 5)
        Dim Renderers As New List(Of RenderArray.RenderText)
        For Count As Integer = 0 To ParenPos.Count - 1
            If Count = ParenPos.Count - 1 Then
                Renderers.Add(New RenderArray.RenderText(RenderArray.RenderDisplayClass.eLTR, Str.Substring(Base)) With {.Clr = Cols(ParenPos(Count))})
            ElseIf ParenPos(Count) <> ParenPos(Count + 1) Then
                Renderers.Add(New RenderArray.RenderText(RenderArray.RenderDisplayClass.eLTR, Str.Substring(Base, Count - Base + 1)) With {.Clr = Cols(ParenPos(Count))})
                Base = Count + 1
                Renderers.Add(New RenderArray.RenderText(RenderArray.RenderDisplayClass.eLTR, CStr(ParenPos(Count))) With {.Clr = Color.DarkRed, .Font = "Courier New"})
            End If
        Next
        Return Renderers.ToArray()
    End Function
    Public Shared Function GetRegExText(Str As String) As String
        Return System.Text.RegularExpressions.Regex.Replace(System.Text.RegularExpressions.Regex.Replace(Str, "\\u([0-9a-fA-F]{4})", Function(Match As System.Text.RegularExpressions.Match) ChrW(Integer.Parse(Match.Groups(1).Value, Globalization.NumberStyles.HexNumber))), "[\p{IsArabic}\p{IsArabicPresentationForms-A}\p{IsArabicPresentationForms-B}]+", ArabicData.LeftToRightEmbedding + "$&" + ArabicData.PopDirectionalFormatting)
    End Function
    Public Shared Function GetMetadataRules(ID As String) As Array()
        Dim Output(CachedData.IslamData.MetaRules.Length + 2) As Array
        Output(0) = New String() {}
        Output(1) = New String() {String.Empty, String.Empty, String.Empty}
        Output(2) = New String() {Utility.LoadResourceString("IslamInfo_Name"), Utility.LoadResourceString("IslamInfo_Translation"), Utility.LoadResourceString("IslamInfo_Translation")}
        For Count = 0 To TanzilReader.GetMetaRuleSet("UthmaniQuran").Rules.Length - 1
            Output(3 + Count) = New Object() {TanzilReader.GetMetaRuleSet("UthmaniQuran").Rules(Count).Name, New RenderArray.RenderItem() {New RenderArray.RenderItem(RenderArray.RenderTypes.eText, ColorizeRegExGroups(GetRegExText(TanzilReader.GetMetaRuleSet("UthmaniQuran").Rules(Count).Match), False))}, New RenderArray.RenderItem() {New RenderArray.RenderItem(RenderArray.RenderTypes.eText, ColorizeList(Array.ConvertAll(TanzilReader.GetMetaRuleSet("UthmaniQuran").Rules(Count).Evaluator, Function(Str As String) GetRegExText(Str)), False))}}
        Next
        Return RenderArray.MakeTableJSFunctions(Output, ID)
    End Function
    Public Shared Function GetRuleSetRules(ID As String, Data As IslamData.RuleTranslationCategory.RuleTranslation()) As Array()
        Dim Output(Data.Length + 2) As Array
        Output(0) = New String() {}
        Output(1) = New String() {String.Empty, String.Empty, String.Empty}
        Output(2) = New String() {Utility.LoadResourceString("IslamInfo_Name"), Utility.LoadResourceString("IslamInfo_Translation"), Utility.LoadResourceString("IslamInfo_Translation")}
        For Count = 0 To Data.Length - 1
            Output(3 + Count) = New Object() {Data(Count).Name, New RenderArray.RenderItem() {New RenderArray.RenderItem(RenderArray.RenderTypes.eText, ColorizeRegExGroups(GetRegExText(Data(Count).Match), False))}, New RenderArray.RenderItem() {New RenderArray.RenderItem(RenderArray.RenderTypes.eText, ColorizeRegExGroups(GetRegExText(Data(Count).Evaluator), True))}}
        Next
        Return RenderArray.MakeTableJSFunctions(Output, ID)
    End Function
    Public Shared Function GetRenderedHelpText(ByVal Item As PageLoader.TextItem) As RenderArray
        Dim SchemeType As ArabicData.TranslitScheme = CType(If(CInt(HttpContext.Current.Request.Params("translitscheme")) >= 2, 2 - CInt(HttpContext.Current.Request.Params("translitscheme")) Mod 2, CInt(HttpContext.Current.Request.Params("translitscheme"))), ArabicData.TranslitScheme)
        Dim Scheme As String = If(CInt(HttpContext.Current.Request.Params("translitscheme")) >= 2, CachedData.IslamData.TranslitSchemes(CInt(HttpContext.Current.Request.Params("translitscheme")) \ 2).Name, String.Empty)
        Dim Renderer As New RenderArray(Item.Name)
        Renderer.Items.Add(New RenderArray.RenderItem(RenderArray.RenderTypes.eText, New RenderArray.RenderText() {New RenderArray.RenderText(RenderArray.RenderDisplayClass.eList, Arabic.GetTranslitSchemeMetadata("0"))}))
        Renderer.Items.Add(New RenderArray.RenderItem(RenderArray.RenderTypes.eText, New RenderArray.RenderText() {New RenderArray.RenderText(RenderArray.RenderDisplayClass.eList, TanzilReader.GetTranslationMetadata("1"))}))
        Renderer.Items.Add(New RenderArray.RenderItem(RenderArray.RenderTypes.eText, New RenderArray.RenderText() {New RenderArray.RenderText(RenderArray.RenderDisplayClass.eList, GetMetadataRules("2"))}))
        Renderer.Items.Add(New RenderArray.RenderItem(RenderArray.RenderTypes.eText, New RenderArray.RenderText() {New RenderArray.RenderText(RenderArray.RenderDisplayClass.eList, GetRuleSetRules("3", CachedData.RomanizationRules))}))
        Renderer.Items.Add(New RenderArray.RenderItem(RenderArray.RenderTypes.eText, New RenderArray.RenderText() {New RenderArray.RenderText(RenderArray.RenderDisplayClass.eList, GetRuleSetRules("4", CachedData.ColoringSpelledOutRules))}))
        Renderer.Items.Add(New RenderArray.RenderItem(RenderArray.RenderTypes.eText, New RenderArray.RenderText() {New RenderArray.RenderText(RenderArray.RenderDisplayClass.eList, GetRuleSetRules("5", CachedData.ErrorCheckRules))}))
        Renderer.Items.Add(New RenderArray.RenderItem(RenderArray.RenderTypes.eText, New RenderArray.RenderText() {New RenderArray.RenderText(RenderArray.RenderDisplayClass.eList, GetRuleSetRules("6", CachedData.WarshScript))}))
        Renderer.Items.Add(New RenderArray.RenderItem(RenderArray.RenderTypes.eText, New RenderArray.RenderText() {New RenderArray.RenderText(RenderArray.RenderDisplayClass.eList, GetRuleSetRules("7", CachedData.UthmaniMinimalScript))}))
        Renderer.Items.Add(New RenderArray.RenderItem(RenderArray.RenderTypes.eText, New RenderArray.RenderText() {New RenderArray.RenderText(RenderArray.RenderDisplayClass.eList, GetRuleSetRules("8", CachedData.SimpleEnhancedScript))}))
        Renderer.Items.Add(New RenderArray.RenderItem(RenderArray.RenderTypes.eText, New RenderArray.RenderText() {New RenderArray.RenderText(RenderArray.RenderDisplayClass.eList, GetRuleSetRules("9", CachedData.SimpleScript))}))
        Renderer.Items.Add(New RenderArray.RenderItem(RenderArray.RenderTypes.eText, New RenderArray.RenderText() {New RenderArray.RenderText(RenderArray.RenderDisplayClass.eList, GetRuleSetRules("10", CachedData.SimpleCleanScript))}))
        Renderer.Items.Add(New RenderArray.RenderItem(RenderArray.RenderTypes.eText, New RenderArray.RenderText() {New RenderArray.RenderText(RenderArray.RenderDisplayClass.eList, GetRuleSetRules("11", CachedData.SimpleMinimalScript))}))
        Return Renderer
    End Function
    Public Shared Sub DoErrorCheckBuckwalterText(Strings As String, TranslationID As String)
        If Strings = Nothing Then Return
        Dim Matches As System.Text.RegularExpressions.MatchCollection = System.Text.RegularExpressions.Regex.Matches(Strings, "(.*?)(?:(\\\{)(.*?)(\\\})|$)", System.Text.RegularExpressions.RegexOptions.Singleline)
        Dim EnglishByWord As String() = If(TranslationID = Nothing, {}, Utility.LoadResourceString("IslamInfo_" + TranslationID + "WordByWord").Split("|"c))
        For MatchCount As Integer = 0 To Matches.Count - 1
            If Matches(MatchCount).Length <> 0 Then
                If Matches(MatchCount).Groups(1).Length <> 0 Then
                    Dim ArabicText As String() = Matches(MatchCount).Groups(1).Value.Split(" "c)
                    If ArabicText.Length > 1 And EnglishByWord.Length = ArabicText.Length Then
                    End If
                    Arabic.DoErrorCheck(Arabic.TransliterateFromBuckwalter(Matches(MatchCount).Groups(1).Value))
                    Utility.LoadResourceString("IslamInfo_" + TranslationID)
                End If
                If Matches(MatchCount).Groups(3).Length <> 0 Then
                    ErrorCheckTextFromReferences(Matches(MatchCount).Groups(3).Value)
                End If
            End If
        Next
    End Sub
    Public Shared Function BuckwalterTextFromReferences(ID As String, SchemeType As ArabicData.TranslitScheme, Scheme As String, Strings As String, TranslationID As String, TranslationIndex As Integer) As RenderArray
        Dim Renderer As New RenderArray(ID)
        If Strings = Nothing Then Return Renderer
        Dim Matches As System.Text.RegularExpressions.MatchCollection = System.Text.RegularExpressions.Regex.Matches(Strings, "(.*?)(?:(\\\{)(.*?)(\\\})|$)", System.Text.RegularExpressions.RegexOptions.Singleline)
        Dim EnglishByWord As String() = If(TranslationID = Nothing, {}, Utility.LoadResourceString("IslamInfo_" + TranslationID + "WordByWord").Split("|"c))
        For MatchCount As Integer = 0 To Matches.Count - 1
            If Matches(MatchCount).Length <> 0 Then
                If Matches(MatchCount).Groups(1).Length <> 0 Then
                    Dim ArabicText As String() = Matches(MatchCount).Groups(1).Value.Split(" "c)
                    If ArabicText.Length > 1 Then 'And EnglishByWord.Length = ArabicText.Length Then
                        Dim Transliteration As String() = Arabic.TransliterateToScheme(Arabic.TransliterateFromBuckwalter(Matches(MatchCount).Groups(1).Value), SchemeType, Scheme).Split(" "c)
                        'Renderer.Items.Add(New RenderArray.RenderItem(RenderArray.RenderTypes.eHeaderCenter, New RenderArray.RenderText() {New RenderArray.RenderText(RenderArray.RenderDisplayClass.eLTR, Utility.LoadResourceString("IslamInfo_" + TranslationID))}))
                        Dim Items As New Collections.Generic.List(Of RenderArray.RenderItem)
                        For WordCount As Integer = 0 To EnglishByWord.Length - 1
                            Items.Add(New RenderArray.RenderItem(RenderArray.RenderTypes.eText, New RenderArray.RenderText() {New RenderArray.RenderText(RenderArray.RenderDisplayClass.eArabic, Arabic.TransliterateFromBuckwalter(If(WordCount >= ArabicText.Length, String.Empty, ArabicText(WordCount)))), New RenderArray.RenderText(RenderArray.RenderDisplayClass.eTransliteration, If(WordCount >= Transliteration.Length, String.Empty, Transliteration(WordCount))), New RenderArray.RenderText(RenderArray.RenderDisplayClass.eLTR, EnglishByWord(WordCount))}))
                        Next
                        Renderer.Items.Add(New RenderArray.RenderItem(RenderArray.RenderTypes.eText, New RenderArray.RenderText() {New RenderArray.RenderText(RenderArray.RenderDisplayClass.eNested, Items), New RenderArray.RenderText(RenderArray.RenderDisplayClass.eArabic, Arabic.TransliterateFromBuckwalter(Matches(MatchCount).Groups(1).Value)), New RenderArray.RenderText(RenderArray.RenderDisplayClass.eTransliteration, Arabic.TransliterateToScheme(Arabic.TransliterateFromBuckwalter(Matches(MatchCount).Groups(1).Value), SchemeType, Scheme)), New RenderArray.RenderText(RenderArray.RenderDisplayClass.eLTR, Utility.LoadResourceString("IslamInfo_" + TranslationID))}))
                    Else
                        Renderer.Items.Add(New RenderArray.RenderItem(RenderArray.RenderTypes.eText, New RenderArray.RenderText() {New RenderArray.RenderText(RenderArray.RenderDisplayClass.eArabic, Arabic.TransliterateFromBuckwalter(Matches(MatchCount).Groups(1).Value)), New RenderArray.RenderText(RenderArray.RenderDisplayClass.eTransliteration, Arabic.TransliterateToScheme(Arabic.TransliterateFromBuckwalter(Matches(MatchCount).Groups(1).Value), SchemeType, Scheme)), New RenderArray.RenderText(RenderArray.RenderDisplayClass.eLTR, Utility.LoadResourceString("IslamInfo_" + TranslationID))}))
                    End If
                End If
                If Matches(MatchCount).Groups(3).Length <> 0 Then
                    Renderer.Items.AddRange(TextFromReferences(ID, Matches(MatchCount).Groups(3).Value, SchemeType, Scheme, TranslationIndex).Items)
                End If
            End If
        Next
        Return Renderer
    End Function
    Public Shared Function NormalTextFromReferences(ID As String, Strings As String, SchemeType As ArabicData.TranslitScheme, Scheme As String, TranslationIndex As Integer) As RenderArray
        Dim Renderer As New RenderArray(ID)
        If Strings = Nothing Then Return Renderer
        Dim Matches As System.Text.RegularExpressions.MatchCollection = System.Text.RegularExpressions.Regex.Matches(Strings, "(.*?)(?:(\{)(.*?)(\})|$)", System.Text.RegularExpressions.RegexOptions.Singleline)
        For Count As Integer = 0 To Matches.Count - 1
            If Matches(Count).Length <> 0 Then
                If Matches(Count).Groups(1).Length <> 0 Then
                    Renderer.Items.Add(New RenderArray.RenderItem(RenderArray.RenderTypes.eText, New RenderArray.RenderText() {New RenderArray.RenderText(RenderArray.RenderDisplayClass.ePassThru, Matches(Count).Groups(1).Value)}))
                End If
                If Matches(Count).Groups(3).Length <> 0 Then
                    Renderer.Items.AddRange(TextFromReferences(ID + CStr(Renderer.Items.Count), Matches(Count).Groups(3).Value, SchemeType, Scheme, TranslationIndex).Items)
                End If
            End If
        Next
        Return Renderer
    End Function
    Shared _Abbrevs As Dictionary(Of String, IslamData.AbbrevWord)
    Public Shared ReadOnly Property Abbrevs As Dictionary(Of String, IslamData.AbbrevWord)
        Get
            If _Abbrevs Is Nothing Then
                _Abbrevs = New Dictionary(Of String, IslamData.AbbrevWord)
                For Count = 0 To CachedData.IslamData.Abbreviations.Length - 1
                    Dim AbbrevWord As IslamData.AbbrevWord = CachedData.IslamData.Abbreviations(Count)
                    If CachedData.IslamData.Abbreviations(Count).Text <> String.Empty Then Array.ForEach(CachedData.IslamData.Abbreviations(Count).Text.Split("|"c), Sub(Str As String) _Abbrevs.Add(Str, AbbrevWord))
                    If CachedData.IslamData.Abbreviations(Count).TranslationID <> String.Empty AndAlso Array.IndexOf(CachedData.IslamData.Abbreviations(Count).Text.Split("|"c), CachedData.IslamData.Abbreviations(Count).TranslationID) = -1 Then _Abbrevs.Add(CachedData.IslamData.Abbreviations(Count).TranslationID, CachedData.IslamData.Abbreviations(Count))
                Next
            End If
            Return _Abbrevs
        End Get
    End Property
    Public Shared Sub ErrorCheckTextFromReferences(Strings As String)
        Dim _Options As String() = Strings.Split(";"c)
        Strings = _Options(0)
        Dim Count As Integer = 0
        While _Options(Count).EndsWith("&leftbrace") Or _Options(Count).EndsWith("&rightbrace") Or _Options(Count).EndsWith("&comma") Or _Options(Count).EndsWith("&semicolon")
            _Options(0) = _Options(0) + ";" + _Options(Count + 1)
            Count += 1
        End While
        If TanzilReader.IsQuranTextReference(Strings) Then
        ElseIf Strings.StartsWith("symbol:") Then
            Dim SelArr As String() = Strings.Replace("symbol:", String.Empty).Split(","c)
            For SubCount = 0 To ArabicData.ArabicLetters.Length - 1
                If Array.IndexOf(SelArr, ArabicData.ToCamelCase(ArabicData.ArabicLetters(SubCount).UnicodeName).Replace("ArabicLetter", String.Empty).Replace("Arabic", String.Empty)) <> -1 Then
                End If
            Next
        ElseIf Strings.StartsWith("personalpronoun:") Or Strings.StartsWith("proximaldemonstratives:") Or Strings.StartsWith("distaldemonstratives:") Then
            Dim Words As IslamData.GrammarSet.GrammarNoun()
            Dim SelArr As String()
            If Strings.StartsWith("proximaldemonstratives") Then
                Words = Arabic.GetCatNoun("proxdemo")
                SelArr = Strings.Replace("proximaldemonstratives:", String.Empty).Split(","c)
            ElseIf Strings.StartsWith("distaldemonstratives") Then
                Words = Arabic.GetCatNoun("distdemo")
                SelArr = Strings.Replace("distaldemonstratives:", String.Empty).Split(","c)
            ElseIf Strings.StartsWith("personalpronoun:") Then
                Words = Arabic.GetCatNoun("perspro")
                SelArr = Strings.Replace("personalpronoun:", String.Empty).Split(","c)
            Else
                Words = Nothing
                SelArr = Nothing
            End If
            For Count = 0 To SelArr.Length - 1
                Dim S As String = SelArr(Count)
                If Words Is Nothing OrElse Array.FindIndex(Words, Function(Word As IslamData.GrammarSet.GrammarNoun) S = Word.TranslationID) = -1 Then Debug.Print("Noun Subject ID Not Found: " + SelArr(Count))
            Next
        ElseIf Strings.StartsWith("plurals:") Or Strings.StartsWith("possessivedeterminerpersonalpronoun:") Then
            Dim Words As IslamData.GrammarSet.GrammarTransform()
            Dim SelArr As String()
            If Strings.StartsWith("plurals") Then
                Words = Arabic.GetTransform("mp|fp|bp")
                SelArr = Strings.Replace("plurals:", String.Empty).Split(","c)
            ElseIf Strings.StartsWith("possessivedeterminerpersonalpronoun") Then
                Words = Arabic.GetTransform("posspron")
                SelArr = Strings.Replace("possessivedeterminerpersonalpronoun:", String.Empty).Split(","c)
            Else
                Words = Nothing
                SelArr = Nothing
            End If
            For Count = 0 To SelArr.Length - 1
                Dim S As String = SelArr(Count)
                If Words Is Nothing OrElse Array.FindIndex(Words, Function(Word As IslamData.GrammarSet.GrammarTransform) S = Word.TranslationID) = -1 Then Debug.Print("Transform Subject ID Not Found: " + SelArr(Count))
            Next
        ElseIf Strings.StartsWith("particle:") Then
            Dim SelArr As String() = Strings.Replace("particle:", String.Empty).Split(","c)
            For Count = 0 To SelArr.Length - 1
                If Arabic.GetParticles(SelArr(Count)) Is Nothing Then Debug.Print("Particle ID Not Found: " + SelArr(Count))
            Next
        ElseIf Strings.StartsWith("noun:") Then
            Dim SelArr As String() = Strings.Replace("noun:", String.Empty).Split(","c)
            For Count = 0 To SelArr.Length - 1
                If Arabic.GetCatNoun(SelArr(Count)) Is Nothing Then Debug.Print("Noun ID Not Found: " + SelArr(Count))
            Next
        ElseIf Strings.StartsWith("verb:") Then
            Dim SelArr As String() = Strings.Replace("verb:", String.Empty).Split(","c)
            For Count = 0 To SelArr.Length - 1
                If Arabic.GetVerb(SelArr(Count)) Is Nothing Then Debug.Print("Verb ID Not Found: " + SelArr(Count))
            Next
        ElseIf Strings.StartsWith("transform:") Then
            Dim SelArr As String() = Strings.Replace("transform:", String.Empty).Split(","c)
            For Count = 0 To SelArr.Length - 1
                If Arabic.GetTransform(SelArr(Count)) Is Nothing Then Debug.Print("Transform ID Not Found: " + SelArr(Count))
            Next
        ElseIf Strings.StartsWith("word:") Then
            Dim SelArr As String() = Strings.Replace("word:", String.Empty).Split(","c)
            For Count = 0 To SelArr.Length - 1
                If Not Arabic.GetCatWord(SelArr(Count)).HasValue Then Debug.Print("Word ID Not Found: " + SelArr(Count))
            Next
        ElseIf Strings.StartsWith("phrase:") Then
            Dim SelArr As String() = Strings.Replace("phrase:", String.Empty).Split(","c)
            For Count = 0 To SelArr.Length - 1
                If Not Phrases.GetPhraseCat(SelArr(Count)).HasValue Then Debug.Print("Phrase ID Not Found: " + SelArr(Count))
            Next
        ElseIf Strings.StartsWith("list:") Then
            Dim SelArr As String() = Strings.Replace("list:", String.Empty).Split(","c)
            For Count = 0 To SelArr.Length - 1
                If GetListCats({SelArr(Count)}).Length = 0 Then Debug.Print("List ID Not Found: " + SelArr(Count))
            Next
        ElseIf Abbrevs.ContainsKey(Strings) Then
        ElseIf Strings.StartsWith("reference:") Then
        ElseIf Strings.StartsWith("arabic:") Then
        ElseIf Strings.StartsWith("text:") Then
        Else
            Debug.Print("Unknown tag:" + Strings)
        End If
    End Sub
    Public Shared Function TextFromReferences(ID As String, Strings As String, SchemeType As ArabicData.TranslitScheme, Scheme As String, TranslationIndex As Integer) As RenderArray
        Dim _Options As String() = Strings.Split(";"c)
        Dim Count As Integer = 0
        While _Options(Count).EndsWith("&leftbrace") Or _Options(Count).EndsWith("&rightbrace") Or _Options(Count).EndsWith("&comma") Or _Options(Count).EndsWith("&semicolon")
            _Options(0) = _Options(0) + ";" + _Options(Count + 1)
            Count += 1
        End While
        Strings = _Options(0)
        Dim Options As New Dictionary(Of String, String())
        For Count = 1 To _Options.Length - 1
            Dim Vals As String() = _Options(Count).Split("="c)
            Options(Vals(0)) = If(Vals.Length = 1, Nothing, Vals(1).Split(","c))
        Next
        If Options.ContainsKey("Translation") Then TranslationIndex = TanzilReader.GetTranslationIndex(Options("Translation")(0))
        If Options.ContainsKey("TranslitStyle") Then
            If (Options("TranslitStyle")(0) = "Literal") Then
                SchemeType = ArabicData.TranslitScheme.Literal
            ElseIf (Options("TranslitStyle")(0) = "RuleBased") Then
                SchemeType = ArabicData.TranslitScheme.RuleBased
            ElseIf (Options("TranslitStyle")(0) = "LearningMode") Then
                SchemeType = ArabicData.TranslitScheme.LearningMode
            Else
                SchemeType = ArabicData.TranslitScheme.None
            End If
        End If
        If Options.ContainsKey("TranslitScheme") Then Scheme = Options("TranslitScheme")(0)
        Dim Renderer As New RenderArray(String.Empty)
        If Strings = Nothing Then Return Renderer
        'text before and after reference matches needs rendering
        'hadith reference matching {name,book/hadith}
        If TanzilReader.IsQuranTextReference(Strings) Then
            Renderer.Items.AddRange(TanzilReader.QuranTextFromReference(Strings, SchemeType, Scheme, TranslationIndex, Options.ContainsKey("W4W") Or Options.ContainsKey("W4WNum"), Options.ContainsKey("W4WNum"), Options.ContainsKey("NoArabic"), Options.ContainsKey("Header"), Options.ContainsKey("NoRef"), Options.ContainsKey("Colorize"), Options.ContainsKey("Verses")).Items)
        ElseIf Strings.StartsWith("search:") Then
            Dim SelArr As String() = Strings.Replace("search:", String.Empty).Split(","c)
            Renderer.Items.AddRange(TanzilReader.QuranTextFromSearch(SelArr(0).Replace("&leftbrace;", "{").Replace("&rightbrace;", "}").Replace("&comma;", ",").Replace("&semicolon;", ";"), SchemeType, Scheme, TranslationIndex, Options.ContainsKey("W4W") Or Options.ContainsKey("W4WNum"), Options.ContainsKey("W4WNum"), Options.ContainsKey("NoArabic"), Options.ContainsKey("Header"), Options.ContainsKey("NoRef"), Options.ContainsKey("Colorize"), Options.ContainsKey("Verses")).Items)
        ElseIf Strings.StartsWith("symbol:") Then
            Dim Symbols As New List(Of ArabicData.ArabicSymbol)
            Dim SelArr As String() = Strings.Replace("symbol:", String.Empty).Split(","c)
            For SubCount = 0 To ArabicData.ArabicLetters.Length - 1
                If Array.IndexOf(SelArr, ArabicData.ToCamelCase(ArabicData.ArabicLetters(SubCount).UnicodeName).Replace("ArabicLetter", String.Empty).Replace("Arabic", String.Empty)) <> -1 Then
                    Symbols.Add(ArabicData.ArabicLetters(SubCount))
                End If
            Next
            Renderer.Items.Add(New RenderArray.RenderItem(RenderArray.RenderTypes.eText, New RenderArray.RenderText() {New RenderArray.RenderText(RenderArray.RenderDisplayClass.eList, Arabic.SymbolDisplay(Symbols.ToArray(), SchemeType, Scheme, If(Options.ContainsKey("Cols"), Options("Cols"), Nothing)))}))
        ElseIf Strings.StartsWith("personalpronoun:") Or Strings.StartsWith("proximaldemonstratives:") Or Strings.StartsWith("distaldemonstratives:") Then
            Dim Words As IslamData.GrammarSet.GrammarNoun()
            Dim SelArr As String()
            If Strings.StartsWith("proximaldemonstratives") Then
                Words = Arabic.GetCatNoun("proxdemo")
                SelArr = Strings.Replace("proximaldemonstratives:", String.Empty).Split(","c)
            ElseIf Strings.StartsWith("distaldemonstratives") Then
                Words = Arabic.GetCatNoun("distdemo")
                SelArr = Strings.Replace("distaldemonstratives:", String.Empty).Split(","c)
            ElseIf Strings.StartsWith("personalpronoun:") Then
                Words = Arabic.GetCatNoun("perspro")
                SelArr = Strings.Replace("personalpronoun:", String.Empty).Split(","c)
            Else
                Words = Nothing
                SelArr = Nothing
            End If
            Renderer.Items.Add(New RenderArray.RenderItem(RenderArray.RenderTypes.eText, New RenderArray.RenderText() {New RenderArray.RenderText(RenderArray.RenderDisplayClass.eList, Arabic.DisplayPronoun(Array.FindAll(Words, Function(Word As IslamData.GrammarSet.GrammarNoun) Array.IndexOf(SelArr, Word.TranslationID) <> -1), ID, True, SchemeType, Scheme, If(Options.ContainsKey("Cols"), Options("Cols"), Nothing)))}))
        ElseIf Strings.StartsWith("plurals:") Or Strings.StartsWith("possessivedeterminerpersonalpronoun:") Then
            Dim Words As IslamData.GrammarSet.GrammarTransform()
            Dim SelArr As String()
            If Strings.StartsWith("plurals") Then
                Words = Arabic.GetTransform("mp|fp|bp")
                SelArr = Strings.Replace("plurals:", String.Empty).Split(","c)
            ElseIf Strings.StartsWith("possessivedeterminerpersonalpronoun") Then
                Words = Arabic.GetTransform("posspron")
                SelArr = Strings.Replace("possessivedeterminerpersonalpronoun:", String.Empty).Split(","c)
            Else
                Words = Nothing
                SelArr = Nothing
            End If
            If Not Words Is Nothing Then Renderer.Items.Add(New RenderArray.RenderItem(RenderArray.RenderTypes.eText, New RenderArray.RenderText() {New RenderArray.RenderText(RenderArray.RenderDisplayClass.eList, Arabic.DisplayTransform(String.Empty, Array.FindAll(Words, Function(Word As IslamData.GrammarSet.GrammarTransform) Array.IndexOf(SelArr, Word.TranslationID) <> -1), ID, True, True, SchemeType, Scheme, If(Options.ContainsKey("Cols"), Options("Cols"), Nothing)))}))
        ElseIf Strings.StartsWith("particle:") Then
            Dim SelArr As String() = Strings.Replace("particle:", String.Empty).Split(","c)
            Dim Words As New List(Of IslamData.GrammarSet.GrammarParticle)
            For Count = 0 To SelArr.Length - 1
                Words.AddRange(Arabic.GetParticles(SelArr(Count)))
            Next
            Renderer.Items.Add(New RenderArray.RenderItem(RenderArray.RenderTypes.eText, New RenderArray.RenderText() {New RenderArray.RenderText(RenderArray.RenderDisplayClass.eList, Arabic.DisplayParticle(Words.ToArray(), ID, SchemeType, Scheme, If(Options.ContainsKey("Cols"), Options("Cols"), Nothing)))}))
        ElseIf Strings.StartsWith("noun:") Then
            Dim SelArr As String() = Strings.Replace("noun:", String.Empty).Split(","c)
            Dim Words As New List(Of IslamData.GrammarSet.GrammarNoun)
            For Count = 0 To SelArr.Length - 1
                Words.AddRange(Arabic.GetCatNoun(SelArr(Count)))
            Next
            Renderer.Items.Add(New RenderArray.RenderItem(RenderArray.RenderTypes.eText, New RenderArray.RenderText() {New RenderArray.RenderText(RenderArray.RenderDisplayClass.eList, Arabic.NounDisplay(Words.ToArray(), ID, SchemeType, Scheme, If(Options.ContainsKey("Cols"), Options("Cols"), Nothing)))}))
        ElseIf Strings.StartsWith("verb:") Then
            Dim SelArr As String() = Strings.Replace("verb:", String.Empty).Split(","c)
            Dim Words As New List(Of IslamData.GrammarSet.GrammarVerb)
            For Count = 0 To SelArr.Length - 1
                Words.AddRange(Arabic.GetVerb(SelArr(Count)))
            Next
            Renderer.Items.Add(New RenderArray.RenderItem(RenderArray.RenderTypes.eText, New RenderArray.RenderText() {New RenderArray.RenderText(RenderArray.RenderDisplayClass.eList, Arabic.VerbDisplay(Words.ToArray(), ID, SchemeType, Scheme, If(Options.ContainsKey("Cols"), Options("Cols"), Nothing)))}))
        ElseIf Strings.StartsWith("transform:") Then
            Dim SelArr As String() = Strings.Replace("transform:", String.Empty).Split(","c)
            Dim Words As New List(Of IslamData.GrammarSet.GrammarTransform)
            For Count = 0 To SelArr.Length - 1
                Words.AddRange(Arabic.GetTransform(SelArr(Count)))
            Next
            Renderer.Items.Add(New RenderArray.RenderItem(RenderArray.RenderTypes.eText, New RenderArray.RenderText() {New RenderArray.RenderText(RenderArray.RenderDisplayClass.eList, Arabic.DisplayTransform(String.Empty, Words.ToArray(), ID, False, False, SchemeType, Scheme, If(Options.ContainsKey("Cols"), Options("Cols"), Nothing)))}))
        ElseIf Strings.StartsWith("word:") Then
            Dim Words As IslamData.GrammarSet.GrammarWord() = Arabic.GetCatWords(Strings.Replace("word:", String.Empty).Split(","c))
            Renderer.Items.Add(New RenderArray.RenderItem(RenderArray.RenderTypes.eText, New RenderArray.RenderText() {New RenderArray.RenderText(RenderArray.RenderDisplayClass.eList, Arabic.DisplayWord(Words, ID, SchemeType, Scheme))}))
        ElseIf Strings.StartsWith("phrase:") Then
            Dim SelArr As String() = Strings.Replace("phrase:", String.Empty).Split(","c)
            Dim PhraseCats As IslamData.Phrase() = Phrases.GetPhraseCats(SelArr)
            For Count = 0 To PhraseCats.Length - 1
                Renderer.Items.AddRange(Phrases.DoGetRenderedPhraseText(SchemeType, Scheme, PhraseCats(Count), TranslationIndex))
            Next
        ElseIf Strings.StartsWith("list:") Then
            Dim SelArr As String() = Strings.Replace("list:", String.Empty).Split(","c)
            Dim ListCats As IslamData.ListCategory.Word() = GetListCats(SelArr)
            For Count = 0 To ListCats.Length - 1
                Renderer.Items.Add(New RenderArray.RenderItem(RenderArray.RenderTypes.eHeaderCenter, New RenderArray.RenderText() {New RenderArray.RenderText(RenderArray.RenderDisplayClass.eLTR, Utility.LoadResourceString("IslamInfo_" + ListCats(Count).TranslationID))}))
                Renderer.Items.AddRange(BuckwalterTextFromReferences(ID, SchemeType, Scheme, ListCats(Count).Text, String.Empty, TranslationIndex).Items)
            Next
        ElseIf Abbrevs.ContainsKey(Strings) Then
            Dim AbbrevWord As IslamData.AbbrevWord = Abbrevs(Strings)
            Dim PhraseCat As Nullable(Of IslamData.Phrase) = Phrases.GetPhraseCat(AbbrevWord.TranslationID)
            Dim GrammarWord As Nullable(Of IslamData.GrammarSet.GrammarWord) = Arabic.GetCatWord(AbbrevWord.TranslationID)
            Dim Items As New List(Of RenderArray.RenderItem)
            If AbbrevWord.Font <> String.Empty Then
                If Options.ContainsKey("Char") Then Options("Char") = Array.ConvertAll(Options("Char"), Function(Str As String) Char.ConvertFromUtf32(Integer.Parse(Str, Globalization.NumberStyles.HexNumber)))
                If Options.Count = 0 Then Items.Add(New RenderArray.RenderItem(RenderArray.RenderTypes.eText, New RenderArray.RenderText() {New RenderArray.RenderText(RenderArray.RenderDisplayClass.eTag, AbbrevWord.TranslationID + "|" + AbbrevWord.Text)}))
                Array.ForEach(AbbrevWord.Font.Split("|"c),
                    Sub(Str As String)
                        Dim Font As String = String.Empty
                        If Str.Contains(";") Then
                            Font = Str.Split(";"c)(0)
                            Str = Str.Split(";"c)(1)
                        End If
                        If Not Options.ContainsKey("Font") OrElse Array.IndexOf(Options("Font"), Font) <> -1 Then
                            Array.ForEach(Str.Split(","c),
                                Sub(SubStr As String)
                                    If Not Options.ContainsKey("Char") OrElse Array.IndexOf(Options("Char"), String.Join(String.Empty, Array.ConvertAll(SubStr.Split("+"c), Function(S As String) Char.ConvertFromUtf32(Integer.Parse(S, Globalization.NumberStyles.HexNumber))))) <> -1 Then
                                        Dim RendText As New RenderArray.RenderText(RenderArray.RenderDisplayClass.eArabic, String.Join(String.Empty, Array.ConvertAll(SubStr.Split("+"c), Function(Split As String) Char.ConvertFromUtf32(Integer.Parse(Split, System.Globalization.NumberStyles.HexNumber)))))
                                        RendText.Font = Font
                                        Items.Add(New RenderArray.RenderItem(RenderArray.RenderTypes.eText, New RenderArray.RenderText() {RendText}))
                                    End If
                                End Sub)
                        End If
                    End Sub)
            End If
            If PhraseCat.HasValue Then
                Renderer.Items.AddRange(Phrases.DoGetRenderedPhraseText(SchemeType, Scheme, PhraseCat.Value, TranslationIndex))
                Renderer.Items.AddRange(Items)
            End If
            If GrammarWord.HasValue Then
                If AbbrevWord.Font <> String.Empty Then
                    Renderer.Items.Add(New RenderArray.RenderItem(RenderArray.RenderTypes.eText, New RenderArray.RenderText() {New RenderArray.RenderText(RenderArray.RenderDisplayClass.eArabic, Arabic.TransliterateFromBuckwalter(GrammarWord.Value.Text)), New RenderArray.RenderText(RenderArray.RenderDisplayClass.eTransliteration, Arabic.TransliterateToScheme(Arabic.TransliterateFromBuckwalter(GrammarWord.Value.Text), SchemeType, Scheme).Trim()), New RenderArray.RenderText(RenderArray.RenderDisplayClass.eLTR, Utility.LoadResourceString("IslamInfo_" + GrammarWord.Value.TranslationID)), New RenderArray.RenderText(RenderArray.RenderDisplayClass.eNested, Items)}))
                Else
                    Renderer.Items.Add(New RenderArray.RenderItem(RenderArray.RenderTypes.eText, New RenderArray.RenderText() {New RenderArray.RenderText(RenderArray.RenderDisplayClass.eArabic, Arabic.TransliterateFromBuckwalter(GrammarWord.Value.Text)), New RenderArray.RenderText(RenderArray.RenderDisplayClass.eTransliteration, Arabic.TransliterateToScheme(Arabic.TransliterateFromBuckwalter(GrammarWord.Value.Text), SchemeType, Scheme).Trim()), New RenderArray.RenderText(RenderArray.RenderDisplayClass.eLTR, Utility.LoadResourceString("IslamInfo_" + GrammarWord.Value.TranslationID))}))
                End If
            End If
        ElseIf Strings.StartsWith("reference:") Then
            Renderer.Items.Add(New RenderArray.RenderItem(RenderArray.RenderTypes.eHeaderCenter, New RenderArray.RenderText() {New RenderArray.RenderText(RenderArray.RenderDisplayClass.eLTR, "(" + Strings.Replace("reference:", String.Empty) + ")")}))
        ElseIf Strings.StartsWith("arabic:") Then
            Renderer.Items.Add(New RenderArray.RenderItem(RenderArray.RenderTypes.eText, New RenderArray.RenderText() {New RenderArray.RenderText(RenderArray.RenderDisplayClass.eArabic, Arabic.TransliterateFromBuckwalter(Strings.Replace("arabic:", String.Empty).Replace("&leftbrace;", "{").Replace("&rightbrace;", "}").Replace("&comma;", ",").Replace("&semicolon;", ";")))}))
        ElseIf Strings.StartsWith("text:") Then
            Renderer.Items.Add(New RenderArray.RenderItem(RenderArray.RenderTypes.eText, New RenderArray.RenderText() {New RenderArray.RenderText(RenderArray.RenderDisplayClass.eLTR, Strings.Replace("text:", String.Empty))}))
        Else
        End If
        Return Renderer
    End Function
    Public Shared Function GetListCategories() As String()
        Return Array.ConvertAll(CachedData.IslamData.Lists, Function(Convert As IslamData.ListCategory) Utility.LoadResourceString("IslamInfo_" + Convert.Title))
    End Function
    Public Shared Function GetListCat(ID As String) As Nullable(Of IslamData.ListCategory.Word)
        GetListCat = New Nullable(Of IslamData.ListCategory.Word)
        If ListIDs.ContainsKey(ID) Then GetListCat = ListIDs(ID)
    End Function
    Public Shared _ListIDs As Dictionary(Of String, IslamData.ListCategory.Word)
    Public Shared _ListTitles As Dictionary(Of String, IslamData.ListCategory)
    Public Shared ReadOnly Property ListIDs As Dictionary(Of String, IslamData.ListCategory.Word)
        Get
            If _ListIDs Is Nothing Then
                _ListIDs = New Dictionary(Of String, IslamData.ListCategory.Word)
                For Count As Integer = 0 To CachedData.IslamData.Lists.Length - 1
                    If Not CachedData.IslamData.Lists(Count).Words Is Nothing Then
                        For SubCount As Integer = 0 To CachedData.IslamData.Lists(Count).Words.Length - 1
                            _ListIDs.Add(CachedData.IslamData.Lists(Count).Words(SubCount).TranslationID, CachedData.IslamData.Lists(Count).Words(SubCount))
                        Next
                    End If
                Next
            End If
            Return _ListIDs
        End Get
    End Property
    Public Shared ReadOnly Property ListTitles As Dictionary(Of String, IslamData.ListCategory)
        Get
            If _ListTitles Is Nothing Then
                _ListTitles = New Dictionary(Of String, IslamData.ListCategory)
                For Count As Integer = 0 To CachedData.IslamData.Lists.Length - 1
                    _ListTitles.Add(CachedData.IslamData.Lists(Count).Title, CachedData.IslamData.Lists(Count))
                Next
            End If
            Return _ListTitles
        End Get
    End Property
    Public Shared Function GetListCats(SelArr As String()) As IslamData.ListCategory.Word()
        Dim ListCats As New List(Of IslamData.ListCategory.Word)
        For SelCount As Integer = 0 To SelArr.Length - 1
            If ListTitles.ContainsKey(SelArr(SelCount)) Then ListCats.AddRange(ListTitles(SelArr(SelCount)).Words)
            Dim Word As Nullable(Of IslamData.ListCategory.Word) = GetListCat(SelArr(SelCount))
            If Word.HasValue Then ListCats.Add(Word.Value)
        Next
        Return ListCats.ToArray()
    End Function
End Class
Public Class Phrases
    Public Shared Function GetPhraseCat(ID As String) As Nullable(Of IslamData.Phrase)
        GetPhraseCat = New Nullable(Of IslamData.Phrase)
        If PhraseIDs.ContainsKey(ID) Then GetPhraseCat = PhraseIDs(ID)
    End Function
    Public Shared _PhraseIDs As Dictionary(Of String, IslamData.Phrase)
    Public Shared ReadOnly Property PhraseIDs As Dictionary(Of String, IslamData.Phrase)
        Get
            If _PhraseIDs Is Nothing Then
                _PhraseIDs = New Dictionary(Of String, IslamData.Phrase)
                For Count As Integer = 0 To CachedData.IslamData.Phrases.Length - 1
                    _PhraseIDs.Add(CachedData.IslamData.Phrases(Count).TranslationID, CachedData.IslamData.Phrases(Count))
                Next
            End If
            Return _PhraseIDs
        End Get
    End Property
    Public Shared Function GetPhraseCats(SelArr As String()) As IslamData.Phrase()
        Dim PhraseCats As New List(Of IslamData.Phrase)
        For SelCount As Integer = 0 To SelArr.Length - 1
            Dim Word As Nullable(Of IslamData.Phrase) = GetPhraseCat(SelArr(SelCount))
            If Word.HasValue Then PhraseCats.Add(Word.Value)
        Next
        Return PhraseCats.ToArray()
    End Function
    Public Shared Function GetRenderedPhraseText(ByVal Item As PageLoader.TextItem) As RenderArray
        Dim SchemeType As ArabicData.TranslitScheme = CType(If(CInt(HttpContext.Current.Request.QueryString.Get("translitscheme")) >= 2, 2 - CInt(HttpContext.Current.Request.QueryString.Get("translitscheme")) Mod 2, CInt(HttpContext.Current.Request.QueryString.Get("translitscheme"))), ArabicData.TranslitScheme)
        Dim Scheme As String = If(CInt(HttpContext.Current.Request.QueryString.Get("translitscheme")) >= 2, CachedData.IslamData.TranslitSchemes(CInt(HttpContext.Current.Request.QueryString.Get("translitscheme")) \ 2).Name, String.Empty)
        Return DoGetRenderedCatText(Item.Name, SchemeType, Scheme, CachedData.IslamData.Phrases, TanzilReader.GetTranslationIndex(HttpContext.Current.Request.QueryString.Get("qurantranslation")))
    End Function
    Public Shared Function DoGetRenderedPhraseText(SchemeType As ArabicData.TranslitScheme, Scheme As String, Verse As IslamData.Phrase, TranslationIndex As Integer) As List(Of RenderArray.RenderItem)
        Return DocBuilder.BuckwalterTextFromReferences(String.Empty, SchemeType, Scheme, Verse.Text, Verse.TranslationID, TranslationIndex).Items
    End Function
    Public Shared Function DoGetRenderedCatText(ID As String, SchemeType As ArabicData.TranslitScheme, Scheme As String, Category As IslamData.Phrase(), TranslationIndex As Integer) As RenderArray
        Dim Renderer As New RenderArray(ID)
        For SubCount As Integer = 0 To Category.Length - 1
            Renderer.Items.AddRange(DoGetRenderedPhraseText(SchemeType, Scheme, Category(SubCount), TranslationIndex))
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
    Public Shared ArabicSpecialLetters As String() = {ArabicData.ArabicLetterAlefWithHamzaAbove, ArabicData.ArabicLetterAlefWithHamzaBelow, ArabicData.ArabicLetterWawWithHamzaAbove, ArabicData.ArabicLetterYehWithHamzaAbove, ArabicData.ArabicLetterAlefWasla, ArabicData.ArabicLetterAlefWithMaddaAbove, ArabicData.ArabicLetterHamza, ArabicData.ArabicLetterTehMarbuta, ArabicData.ArabicLetterAlefMaksura}
    Public Shared ArabicDiacriticsBefore As String() = {ArabicData.ArabicSukun, ArabicData.ArabicFatha, ArabicData.ArabicKasra, ArabicData.ArabicDamma}
    Public Shared ArabicDiacriticsAfter As String() = {ArabicData.ArabicSukun, ArabicData.ArabicFatha, ArabicData.ArabicKasra, ArabicData.ArabicDamma, ArabicData.ArabicKasratan, ArabicData.ArabicDammatan, ArabicData.ArabicFathatan + ArabicData.ArabicLetterAlef, ArabicData.ArabicShadda}
    Public Shared Function GetChangeQuizJS() As String()
        Return New String() {"javascript: changeQuiz();", String.Empty, _
                             "function changeQuiz() { qtype = $('#quizselection').val(); qwrong = 0; qright = 0; nextQuestion(); }"}
    End Function
    Public Shared Function DisplayCount(ByVal Item As PageLoader.TextItem) As String
        Return "Wrong: 0 Right: 0"
    End Function
    Public Shared Function GetQuizSet() As String()
        Dim Quiz As Integer = CInt(HttpContext.Current.Request.QueryString.Get("quizselection"))
        If Quiz = 0 Then
            Return CachedData.ArabicLetters
        ElseIf Quiz = 1 Then
            Dim CurList As New Generic.List(Of String)
            For Count As Integer = 0 To ArabicDiacriticsBefore.Length - 1
                Dim CurLet As String = ArabicDiacriticsBefore(Count)
                CurList.AddRange(Array.ConvertAll(CachedData.ArabicLetters, Function(Str As String) CurLet + Str))
            Next
            Return CurList.ToArray()
        ElseIf Quiz = 2 Then
            Dim CurList As New Generic.List(Of String)
            For Count As Integer = 0 To ArabicDiacriticsAfter.Length - 1
                Dim CurLet As String = ArabicDiacriticsAfter(Count)
                CurList.AddRange(Array.ConvertAll(CachedData.ArabicLetters, Function(Str As String) Str + CurLet))
            Next
            Return CurList.ToArray()
        Else
            Return Nothing
        End If
    End Function
    Public Shared Function DisplayQuestion(ByVal Item As PageLoader.TextItem) As String
        HttpContext.Current.Items.Add("rnd", Timer())
        Rnd(-1)
        Randomize(CDbl(HttpContext.Current.Items("rnd")))
        Dim Count As Integer = CInt(Math.Floor(Rnd() * 4))
        Dim QuizSet As New List(Of String)
        QuizSet.AddRange(GetQuizSet())
        While Count <> 0
            QuizSet.RemoveAt(CInt(Math.Floor(Rnd() * QuizSet.Count)))
            Count -= 1
        End While
        Return QuizSet(CInt(Math.Floor(Rnd() * QuizSet.Count)))
    End Function
    Public Shared Function DisplayAnswer(ByVal Item As PageLoader.ButtonItem) As String
        Dim Count As Integer
        Rnd(-1)
        Randomize(CDbl(HttpContext.Current.Items("rnd")))
        Rnd()
        Dim Quiz As Integer = CInt(HttpContext.Current.Request.QueryString.Get("quizselection"))
        Dim QuizSet As New List(Of String)
        QuizSet.AddRange(GetQuizSet())
        For Count = 2 To Integer.Parse(Item.Name.Replace("answer", String.Empty))
            QuizSet.RemoveAt(CInt(Math.Floor(Rnd() * QuizSet.Count)))
        Next
        Dim SchemeType As ArabicData.TranslitScheme = CType(If(CInt(HttpContext.Current.Request.QueryString.Get("translitscheme")) >= 2, 2 - CInt(HttpContext.Current.Request.QueryString.Get("translitscheme")) Mod 2, CInt(HttpContext.Current.Request.QueryString.Get("translitscheme"))), ArabicData.TranslitScheme)
        Dim Scheme As String = If(CInt(HttpContext.Current.Request.QueryString.Get("translitscheme")) >= 2, CachedData.IslamData.TranslitSchemes(CInt(HttpContext.Current.Request.QueryString.Get("translitscheme")) \ 2).Name, String.Empty)
        If SchemeType = ArabicData.TranslitScheme.None Then
            SchemeType = ArabicData.TranslitScheme.LearningMode
            Scheme = CachedData.IslamData.TranslitSchemes(3).Name
        End If
        Return Arabic.TransliterateToScheme(QuizSet(CInt(Math.Floor(Rnd() * QuizSet.Count))), SchemeType, Scheme)
    End Function
    Public Shared Function VerifyAnswer() As String()
        Dim JSList As New List(Of String) From {"javascript: verifyAnswer(this);", String.Empty, _
            Arabic.GetArabicSymbolJSArray(), Arabic.GetTranslitSchemeJSArray(), Arabic.FindLetterBySymbolJS, _
            "var arabicLets = " + Utility.MakeJSArray(CachedData.ArabicLetters, False) + ";", _
            "var arabicDiacriticsBefore = " + Utility.MakeJSArray(ArabicDiacriticsBefore, False) + ";", _
            "var arabicDiacriticsAfter = " + Utility.MakeJSArray(ArabicDiacriticsAfter, False) + ";", _
            "var qtype = 'arabicletters', qwrong = 0, qright = 0; if (typeof renderList == 'undefined') { renderList = []; }", _
            "function getUniqueRnd(excl, count) { var rnd; if (excl.length === count) return 0; do { rnd = Math.floor(Math.random() * count); } while (excl.indexOf(rnd) !== -1); return rnd; }", _
            "function getQuizSet() { if (qtype === 'arabicletters') return arabicLets; if (qtype === 'arabiclettersdiacritics') { var count = 0, arr = []; for (count = 0; count < arabicDiacriticsAfter.length; count++) { arr = arr.concat(arabicLets.map(function(val) { return val + arabicDiacriticsAfter[count]; })); } return arr; } if (qtype === 'arabicdiacriticsletters') { var count = 0, arr = []; for (count = 0; count < arabicDiacriticsBefore.length; count++) { arr = arr.concat(arabicLets.map(function(val) { return arabicDiacriticsBefore[count] + val; })); } return arr; }; return []; }", _
            "function getQA(quizSet, quest, nidx) { if (quest) return quizSet[nidx]; return (parseInt($('#translitscheme').val(), 10) % 2) === 0 ? transliterateWithRules(quizSet[nidx], parseInt($('#translitscheme').val(), 10) >= 2 ? Math.floor((parseInt($('#translitscheme').val(), 10) - 2) / 2) + 2 : 5, null, false) : doTransliterate(quizSet[nidx], true, parseInt($('#translitscheme').val(), 10)); }", _
            "function nextQuestion() { $('#count').text('Wrong: ' + qwrong + ' Right: ' + qright); var i = Math.floor(Math.random() * 4), quizSet = getQuizSet(), pos = quizSet.length, nidx = getUniqueRnd([], pos), aidx = []; aidx[0] = getUniqueRnd([nidx], pos); aidx[1] = getUniqueRnd([nidx, aidx[0]], pos); aidx[2] = getUniqueRnd([nidx, aidx[0], aidx[1]], pos); $('#quizquestion').text(getQA(quizSet, true, nidx)); $('#answer1').prop('value', getQA(quizSet, false, i === 0 ? nidx : aidx[0])); $('#answer2').prop('value', getQA(quizSet, false, i === 1 ? nidx : aidx[i > 1 ? 1 : 0])); $('#answer3').prop('value', getQA(quizSet, false, i === 2 ? nidx : aidx[i > 2 ? 2 : 1])); $('#answer4').prop('value', getQA(quizSet, false, i === 3 ? nidx : aidx[2])); }", _
            "function verifyAnswer(ctl) { $(ctl).prop('value') === ((parseInt($('#translitscheme').val(), 10) % 2) === 0 ? transliterateWithRules($('#quizquestion').text().trim(), parseInt($('#translitscheme').val(), 10) >= 2 ? Math.floor((parseInt($('#translitscheme').val(), 10) - 2) / 2) + 2 : 5, null, false) : doTransliterate($('#quizquestion').text().trim(), true, parseInt($('#translitscheme').val(), 10))) ? qright++ : qwrong++; nextQuestion(); }"}
        JSList.AddRange(ArabicData.GetUniCats())
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
    Public Shared Function GetTranslationMetadata(ID As String) As Array()
        Dim Output(CachedData.IslamData.Translations.TranslationList.Length + 2) As Array
        Output(0) = New String() {}
        Output(1) = New String() {String.Empty, String.Empty}
        Output(2) = New String() {Utility.LoadResourceString("IslamInfo_Name"), Utility.LoadResourceString("IslamInfo_Translation")}
        For Count = 0 To CachedData.IslamData.Translations.TranslationList.Length - 1
            Output(3 + Count) = {CachedData.IslamData.Translations.TranslationList(Count).FileName, CachedData.IslamData.Translations.TranslationList(Count).Translator + " (" + CachedData.IslamData.Translations.TranslationList(Count).Name + ")"}
        Next
        Return RenderArray.MakeTableJSFunctions(Output, ID)
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
            Utility.MakeJSArray(Array.ConvertAll(Of Array, String)(Arabic.GetRecitationSymbols(), Function(Convert As Array) Utility.MakeJSArray(New String() {CStr(CType(Convert, Object())(0)), CStr(CType(Convert, Object())(1))})), True), _
            Utility.MakeJSArray(Array.ConvertAll(Of Array, String)(GetRecitationRules(), Function(Convert As Array) Utility.MakeJSArray(New String() {CStr(CType(Convert, Object())(0)), CStr(CType(Convert, Object())(1))})), True)}, True)
        Return New String() {"javascript: changeQuranDivision(this.selectedIndex);", String.Empty, Utility.GetClearOptionListJS(), _
        "function changeQuranDivision(index) { var iCount; var qurandata = " + JSArrays + "; var eSelect = $('#quranselection').get(0); clearOptionList(eSelect); for (iCount = 0; iCount < qurandata[index].length; iCount++) { eSelect.options.add(new Option(qurandata[index][iCount][0], qurandata[index][iCount][1])); } }"}
    End Function
    Public Shared Function GetWordVerseModeJS() As String()
        Return New String() {"javascript: changeWordVerseMode(this.selectedIndex);", String.Empty, _
                             "function changeWordVerseMode(index) {}"}
    End Function
    Public Shared Function GetColorCueModeJS() As String()
        Return New String() {"javascript: changeColorCueMode(this.selectedIndex);", String.Empty, _
                             "function changeColorCueMode(index) {}"}
    End Function
    Public Shared Function GetWordPartitions() As String()
        Dim Parts As New Generic.List(Of String) From {Utility.LoadResourceString("IslamInfo_Letters"), Utility.LoadResourceString("IslamInfo_Words"), Utility.LoadResourceString("IslamInfo_UniqueWords"), Utility.LoadResourceString("IslamInfo_UniqueWordsPerPart"), Utility.LoadResourceString("IslamInfo_WordsPerPart"), Utility.LoadResourceString("IslamInfo_UniqueWordsPerStation"), Utility.LoadResourceString("IslamInfo_WordsPerStation"), Utility.LoadResourceString("IslamInfo_IsolatedLetters"), Utility.LoadResourceString("IslamInfo_LetterPatterns"), Utility.LoadResourceString("IslamInfo_LetterPatterns"), Utility.LoadResourceString("IslamInfo_LetterPatterns"), Utility.LoadResourceString("IslamInfo_Prefix"), Utility.LoadResourceString("IslamInfo_Suffix")}
        Parts.AddRange(Array.ConvertAll(CachedData.IslamData.PartsOfSpeech, Function(POS As IslamData.PartOfSpeechInfo) Utility.LoadResourceString("IslamInfo_" + POS.Id)))
        Parts.AddRange(Array.ConvertAll(CachedData.RecitationSymbols, Function(Sym As String) ArabicData.GetUnicodeName(Sym.Chars(0))))
        Parts.AddRange(Array.ConvertAll(CachedData.RecitationSymbols, Function(Sym As String) "Prefix of " + ArabicData.GetUnicodeName(Sym.Chars(0))))
        Parts.AddRange(Array.ConvertAll(CachedData.RecitationSymbols, Function(Sym As String) "Suffix of " + ArabicData.GetUnicodeName(Sym.Chars(0))))
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
            Return String.Empty
        ElseIf Index = 10 Then
            Return String.Empty
        ElseIf Index = 11 Then
            Return CStr(CachedData.PreDictionary.Count)
        ElseIf Index = 12 Then
            Return CStr(CachedData.SufDictionary.Count)
        ElseIf Index >= 13 And Index < 13 + CachedData.IslamData.PartsOfSpeech.Length Then
            Return CStr(CachedData.TagDictionary.Item(CachedData.IslamData.PartsOfSpeech(Index - 13).Symbol).Count)
        ElseIf Index >= 13 + CachedData.IslamData.PartsOfSpeech.Length And Index < 13 + CachedData.IslamData.PartsOfSpeech.Length + CachedData.RecitationSymbols.Length Then
            Return CStr(CachedData.LetterDictionary.Item(CachedData.RecitationSymbols(Index - 13 - CachedData.IslamData.PartsOfSpeech.Length).Chars(0)).Count)
        ElseIf Index >= 13 + CachedData.IslamData.PartsOfSpeech.Length + CachedData.RecitationSymbols.Length And Index < 13 + CachedData.IslamData.PartsOfSpeech.Length + CachedData.RecitationSymbols.Length + CachedData.RecitationSymbols.Length Then
            Return CStr(CachedData.LetterPreDictionary.Item(CachedData.RecitationSymbols(Index - 13 - CachedData.IslamData.PartsOfSpeech.Length - CachedData.RecitationSymbols.Length).Chars(0)).Count)
        ElseIf Index >= 13 + CachedData.IslamData.PartsOfSpeech.Length + CachedData.RecitationSymbols.Length + CachedData.RecitationSymbols.Length And Index < 13 + CachedData.IslamData.PartsOfSpeech.Length + CachedData.RecitationSymbols.Length + CachedData.RecitationSymbols.Length + CachedData.RecitationSymbols.Length Then
            Return CStr(CachedData.LetterSufDictionary.Item(CachedData.RecitationSymbols(Index - 13 - CachedData.IslamData.PartsOfSpeech.Length - CachedData.RecitationSymbols.Length - CachedData.RecitationSymbols.Length).Chars(0)).Count)
        Else
            Return String.Empty
        End If
    End Function
    Public Shared Function GetQuranWordFrequency(ByVal Item As PageLoader.TextItem) As Array()
        Dim Output As New ArrayList
        Dim Total As Integer = 0
        Dim All As Double
        Dim SchemeType As ArabicData.TranslitScheme = CType(If(CInt(HttpContext.Current.Request.QueryString.Get("translitscheme")) >= 2, 2 - CInt(HttpContext.Current.Request.QueryString.Get("translitscheme")) Mod 2, CInt(HttpContext.Current.Request.QueryString.Get("translitscheme"))), ArabicData.TranslitScheme)
        Dim Scheme As String = If(CInt(HttpContext.Current.Request.QueryString.Get("translitscheme")) >= 2, CachedData.IslamData.TranslitSchemes(CInt(HttpContext.Current.Request.QueryString.Get("translitscheme")) \ 2).Name, String.Empty)
        Dim Strings As String
        Dim Index As Integer
        Strings = HttpContext.Current.Request.QueryString.Get("quranselection")
        If Not Strings Is Nothing Then Index = CInt(Strings)
        Output.Add(New String() {String.Empty, String.Empty, String.Empty, String.Empty, String.Empty, String.Empty})
        Output.Add(New String() {If(Index = 8 Or Index = 9, "transliteration", "arabic"), "transliteration", "translation", String.Empty, String.Empty, String.Empty})
        Output.Add(New String() {Utility.LoadResourceString(If(Index = 8 Or Index = 9, "IslamInfo_Transliteration", "IslamInfo_Arabic")), Utility.LoadResourceString("IslamInfo_Transliteration"), Utility.LoadResourceString("IslamInfo_Translation"), Utility.LoadResourceString("IslamSource_WordTotal"), String.Empty, String.Empty})
        If Index = 0 Then
            All = CachedData.TotalLetters
            Dim LetterFreqArray(CachedData.LetterDictionary.Keys.Count - 1) As Char
            CachedData.LetterDictionary.Keys.CopyTo(LetterFreqArray, 0)
            Array.Sort(LetterFreqArray, Function(Key As Char, NextKey As Char) CachedData.LetterDictionary.Item(NextKey).Count.CompareTo(CachedData.LetterDictionary.Item(Key).Count))
            For Count As Integer = 0 To LetterFreqArray.Length - 1
                Total += CachedData.LetterDictionary.Item(LetterFreqArray(Count)).Count
                Output.Add(New String() {ArabicData.LeftToRightEmbedding + ArabicData.GetUnicodeName(LetterFreqArray(Count)) + " ( " + ArabicData.PopDirectionalFormatting + ArabicData.FixStartingCombiningSymbol(LetterFreqArray(Count)) + ArabicData.LeftToRightEmbedding + " )" + ArabicData.PopDirectionalFormatting, String.Empty, String.Empty, CStr(CachedData.LetterDictionary.Item(LetterFreqArray(Count)).Count), (CDbl(CachedData.LetterDictionary.Item(LetterFreqArray(Count)).Count) * 100 / All).ToString("n2"), (CDbl(Total) * 100 / All).ToString("n2")})
            Next
        ElseIf Index = 7 Then
            All = CachedData.TotalIsolatedLetters
            Dim LetterFreqArray(CachedData.IsolatedLetterDictionary.Keys.Count - 1) As Char
            CachedData.IsolatedLetterDictionary.Keys.CopyTo(LetterFreqArray, 0)
            Array.Sort(LetterFreqArray, Function(Key As Char, NextKey As Char) CachedData.IsolatedLetterDictionary.Item(NextKey).Count.CompareTo(CachedData.IsolatedLetterDictionary.Item(Key).Count))
            For Count As Integer = 0 To LetterFreqArray.Length - 1
                Total += CachedData.IsolatedLetterDictionary.Item(LetterFreqArray(Count)).Count
                Output.Add(New String() {ArabicData.LeftToRightEmbedding + ArabicData.GetUnicodeName(LetterFreqArray(Count)) + " ( " + ArabicData.PopDirectionalFormatting + ArabicData.FixStartingCombiningSymbol(LetterFreqArray(Count)) + ArabicData.LeftToRightEmbedding + " )" + ArabicData.PopDirectionalFormatting, String.Empty, String.Empty, CStr(CachedData.IsolatedLetterDictionary.Item(LetterFreqArray(Count)).Count), (CDbl(CachedData.IsolatedLetterDictionary.Item(LetterFreqArray(Count)).Count) * 100 / All).ToString("n2"), (CDbl(Total) * 100 / All).ToString("n2")})
            Next
        ElseIf Index = 1 Or Index = 11 Or Index = 12 Or Index >= 13 And Index < 13 + CachedData.IslamData.PartsOfSpeech.Length + CachedData.RecitationSymbols.Length + CachedData.RecitationSymbols.Length + CachedData.RecitationSymbols.Length Then
            Dim Dict As Generic.Dictionary(Of String, ArrayList)
            If Index = 1 Then
                Dict = New Dictionary(Of String, ArrayList)
                For Each KV As KeyValuePair(Of String, ArrayList) In CachedData.WordDictionary
                    Dim Str As String = KV.Key + vbCrLf + String.Join(vbCrLf, CType(KV.Value.ToArray(GetType(String)), String()))
                    Dict.Add(Str, New ArrayList)
                    For Count As Integer = 0 To KV.Value.Count - 1
                        Dict(Str).AddRange(CachedData.FormDictionary(CStr(KV.Value(Count))))
                    Next
                Next
            ElseIf Index = 11 Then
                Dict = CachedData.PreDictionary
            ElseIf Index = 12 Then
                Dict = CachedData.SufDictionary
            ElseIf Index >= 13 And Index < 13 + CachedData.IslamData.PartsOfSpeech.Length Then
                Dict = CachedData.TagDictionary(CachedData.IslamData.PartsOfSpeech(Index - 13).Symbol)
            ElseIf Index >= 13 + CachedData.IslamData.PartsOfSpeech.Length And Index < 13 + CachedData.IslamData.PartsOfSpeech.Length + CachedData.RecitationSymbols.Length Then
                Dict = CachedData.LetterDictionary(CachedData.RecitationSymbols(Index - 13 - CachedData.IslamData.PartsOfSpeech.Length).Chars(0))
            ElseIf Index >= 13 + CachedData.IslamData.PartsOfSpeech.Length + CachedData.RecitationSymbols.Length And Index < 13 + CachedData.IslamData.PartsOfSpeech.Length + CachedData.RecitationSymbols.Length + CachedData.RecitationSymbols.Length Then
                Dict = CachedData.LetterPreDictionary(CachedData.RecitationSymbols(Index - 13 - CachedData.IslamData.PartsOfSpeech.Length - CachedData.RecitationSymbols.Length).Chars(0))
            ElseIf Index >= 13 + CachedData.IslamData.PartsOfSpeech.Length + CachedData.RecitationSymbols.Length + CachedData.RecitationSymbols.Length And Index < 13 + CachedData.IslamData.PartsOfSpeech.Length + CachedData.RecitationSymbols.Length + CachedData.RecitationSymbols.Length + CachedData.RecitationSymbols.Length Then
                Dict = CachedData.LetterSufDictionary(CachedData.RecitationSymbols(Index - 13 - CachedData.IslamData.PartsOfSpeech.Length - CachedData.RecitationSymbols.Length - CachedData.RecitationSymbols.Length).Chars(0))
            Else
                Dict = Nothing
            End If
            Dim FreqArray(Dict.Keys.Count - 1) As String
            Dict.Keys.CopyTo(FreqArray, 0)
            Total = 0
            All = GetQuranWordTotalNumber()
            Array.Sort(FreqArray, Function(Key As String, NextKey As String) Dict.Item(NextKey).Count.CompareTo(Dict.Item(Key).Count))
            Dim W4WLines As String() = IO.File.ReadAllLines(Utility.GetFilePath("metadata\en.w4w.shehnazshaikh.txt"))
            For Count As Integer = 0 To FreqArray.Length - 1
                Dim TranslationDict As New Dictionary(Of String, ArrayList)
                For WordCount As Integer = 0 To Dict.Item(FreqArray(Count)).Count - 1
                    Dim CheckStr As String = TanzilReader.GetW4WTranslationVerse(W4WLines, CType(Dict.Item(FreqArray(Count))(WordCount), Integer())(0), CType(Dict.Item(FreqArray(Count))(WordCount), Integer())(1), CType(Dict.Item(FreqArray(Count))(WordCount), Integer())(2) - 1)
                    If Not TranslationDict.ContainsKey(CheckStr) Then
                        TranslationDict.Add(CheckStr, New ArrayList)
                    End If
                    TranslationDict(CheckStr).Add(CType(Dict.Item(FreqArray(Count))(WordCount), Integer()))
                Next
                Dim TranslationArray(TranslationDict.Keys.Count - 1) As String
                TranslationDict.Keys.CopyTo(TranslationArray, 0)
                For WordCount As Integer = 0 To TranslationArray.Length - 1
                    TranslationArray(WordCount) += vbCrLf + "(" + String.Join(",", Array.ConvertAll(CType(TranslationDict(TranslationArray(WordCount)).ToArray(GetType(Integer())), Integer()()), Function(Indexes As Integer()) String.Join(":", Array.ConvertAll(Indexes, Function(Idx As Integer) CStr(Idx))))) + ")"
                Next
                Total += Dict.Item(FreqArray(Count)).Count
                Output.Add(New String() {Arabic.TransliterateFromBuckwalter(FreqArray(Count)), Arabic.TransliterateToScheme(Arabic.TransliterateFromBuckwalter(FreqArray(Count)), If(SchemeType = ArabicData.TranslitScheme.RuleBased, ArabicData.TranslitScheme.LearningMode, SchemeType), Scheme), String.Join(vbCrLf, TranslationArray), CStr(Dict.Item(FreqArray(Count)).Count), (CDbl(Dict.Item(FreqArray(Count)).Count) * 100 / All).ToString("n2"), (CDbl(Total) * 100 / All).ToString("n2")})
            Next
        ElseIf Index = 3 Or Index = 4 Or Index = 5 Or Index = 6 Then
            Total = 0
            Dim DivArray As Collections.Generic.List(Of String)()
            If Index = 3 Or Index = 5 Then
                DivArray = If(Index = 5, CachedData.StationUniqueArray, CachedData.PartUniqueArray)
                All = If(Index = 5, CachedData.TotalUniqueWordsInStations, CachedData.TotalUniqueWordsInParts)
                For Count As Integer = 0 To CInt(IIf(Index = 5, TanzilReader.GetStationCount(), TanzilReader.GetPartCount())) - 1
                    Total += DivArray(Count).Count
                    Output.Add(New String() {ArabicData.LeftToRightEmbedding + CStr(Count + 1) + ArabicData.PopDirectionalFormatting, String.Empty, String.Empty, CStr(DivArray(Count).Count), (CDbl(DivArray(Count).Count) * 100 / All).ToString("n2"), (CDbl(Total) * 100 / All).ToString("n2")})
                Next
            ElseIf Index = 4 Or Index = 6 Then
                DivArray = If(Index = 6, CachedData.StationUniqueArray, CachedData.PartUniqueArray)
                All = If(Index = 6, CachedData.TotalWordsInStations, CachedData.TotalWordsInParts)
                For Count As Integer = 0 To CInt(IIf(Index = 6, TanzilReader.GetStationCount(), TanzilReader.GetPartCount())) - 1
                    Total += DivArray(Count).Count
                    Output.Add(New String() {ArabicData.LeftToRightEmbedding + CStr(Count + 1) + ArabicData.PopDirectionalFormatting, String.Empty, String.Empty, CStr(DivArray(Count).Count), (CDbl(DivArray(Count).Count) * 100 / All).ToString("n2"), (CDbl(Total) * 100 / All).ToString("n2")})
                Next
            End If
        ElseIf Index = 8 Then
            Output.AddRange(Array.ConvertAll(GetQuranLetterPatterns(), Function(Str As String) {Str, String.Empty, String.Empty, String.Empty, String.Empty, String.Empty}))
        ElseIf Index = 9 Then
            Output.AddRange(Array.ConvertAll(PatternAnalysis(), Function(Str As String) {Str, String.Empty, String.Empty, String.Empty, String.Empty, String.Empty}))
        ElseIf Index = 10 Then
            Output.AddRange(Array.ConvertAll(GetQuranHamzaMaddDoubleLetterPatterns(), Function(Str As String) {Str, String.Empty, String.Empty, String.Empty, String.Empty, String.Empty}))
        End If
        Return CType(Output.ToArray(GetType(Array)), Array())
    End Function
    Public Shared Function IsSingletonDictionary(Dict As Dictionary(Of String, Object())) As Boolean
        Dim Keys(Dict.Keys.Count - 1) As String
        Dict.Keys.CopyTo(Keys, 0)
        Return Dict.Keys.Count = 1 AndAlso (CType(Dict(Keys(0))(0), Dictionary(Of String, Object())).Keys.Count = 0 OrElse IsSingletonDictionary(CType(Dict(Keys(0))(0), Dictionary(Of String, Object()))))
    End Function
    Public Shared Function DumpRecDictionary(Dict As Dictionary(Of String, Object()), Post As Boolean, Depth As Integer, Dual As Boolean) As String
        Dim Str As String = String.Empty
        For Each KV As KeyValuePair(Of String, Object()) In Dict
            If Str <> String.Empty Then Str = If(Post, Str + "/", "/" + Str)
            'if all are 1 recursively
            If IsSingletonDictionary(CType(KV.Value(0), Dictionary(Of String, Object()))) Then
                If Post Then
                    Str += Arabic.TransliterateToScheme(KV.Key, ArabicData.TranslitScheme.Literal, String.Empty) + DumpRecDictionary(CType(KV.Value(0), Dictionary(Of String, Object())), Post, Depth + 1, Dual) + If(Dual AndAlso CType(KV.Value(1), Dictionary(Of String, Object())).Keys.Count <> 0, "!(" + DumpRecDictionary(CType(KV.Value(1), Dictionary(Of String, Object())), Not Post, Depth + 1, False) + ")!", String.Empty)
                Else
                    Str = DumpRecDictionary(CType(KV.Value(0), Dictionary(Of String, Object())), Post, Depth + 1, Dual) + If(Dual AndAlso CType(KV.Value(1), Dictionary(Of String, Object())).Keys.Count <> 0, "!(" + DumpRecDictionary(CType(KV.Value(1), Dictionary(Of String, Object())), Not Post, Depth + 1, False) + ")!", String.Empty) + Arabic.TransliterateToScheme(KV.Key, ArabicData.TranslitScheme.Literal, String.Empty) + Str
                End If
            ElseIf Post Then
                Str += Arabic.TransliterateToScheme(KV.Key, ArabicData.TranslitScheme.Literal, String.Empty) + If(CType(KV.Value(0), Dictionary(Of String, Object())).Keys.Count = 0, String.Empty, vbCrLf + New String(" "c, Depth * 4) + "(" + DumpRecDictionary(CType(KV.Value(0), Dictionary(Of String, Object())), Post, Depth + 1, Dual) + ")" + vbCrLf + New String(" "c, Depth * 4)) + If(Dual AndAlso CType(KV.Value(1), Dictionary(Of String, Object())).Keys.Count <> 0, vbCrLf + New String(" "c, Depth * 4) + "!(" + DumpRecDictionary(CType(KV.Value(1), Dictionary(Of String, Object())), Not Post, Depth + 1, False) + ")!" + vbCrLf + New String(" "c, Depth * 4), String.Empty)
            Else
                Str = If(CType(KV.Value(0), Dictionary(Of String, Object())).Keys.Count = 0, String.Empty, vbCrLf + New String(" "c, Depth * 4) + "(" + DumpRecDictionary(CType(KV.Value(0), Dictionary(Of String, Object())), Post, Depth + 1, Dual) + ")" + vbCrLf + New String(" "c, Depth * 4)) + If(Dual AndAlso CType(KV.Value(1), Dictionary(Of String, Object())).Keys.Count <> 0, vbCrLf + New String(" "c, Depth * 4) + "!(" + DumpRecDictionary(CType(KV.Value(1), Dictionary(Of String, Object())), Not Post, Depth + 1, False) + ")!" + vbCrLf + New String(" "c, Depth * 4), String.Empty) + Arabic.TransliterateToScheme(KV.Key, ArabicData.TranslitScheme.Literal, String.Empty) + Str
            End If
        Next
        Return Str
    End Function
    Public Shared Function CloneDict(Dict As Dictionary(Of String, Object())) As Dictionary(Of String, Object())
        If Dict Is Nothing Then Return Nothing
        Dim Clone As New Dictionary(Of String, Object())
        For Each KV As KeyValuePair(Of String, Object()) In Dict
            Clone.Add(KV.Key, {CloneDict(CType(KV.Value(0), Dictionary(Of String, Object()))), CloneDict(CType(KV.Value(1), Dictionary(Of String, Object())))})
        Next
        Return Clone
    End Function
    Public Shared Sub IntersectDict(ByRef Dict As Dictionary(Of String, Object()), OtherDict As Dictionary(Of String, Object()))
        For Each KV As KeyValuePair(Of String, Object()) In Dict
            If Not OtherDict.ContainsKey(KV.Key) Then
                Dict.Remove(KV.Key)
            Else
                IntersectDict(CType(KV.Value(0), Dictionary(Of String, Object())), CType(OtherDict.Item(KV.Key)(0), Dictionary(Of String, Object())))
                IntersectDict(CType(KV.Value(1), Dictionary(Of String, Object())), CType(OtherDict.Item(KV.Key)(1), Dictionary(Of String, Object())))
            End If
        Next
    End Sub
    Public Shared Sub TrimDict(ByRef Dict As Dictionary(Of String, Object()), OtherDict As Dictionary(Of String, Object()))
        For Each KV As KeyValuePair(Of String, Object()) In Dict
            If OtherDict.ContainsKey(KV.Key) Then

            End If
        Next
    End Sub
    Public Shared Sub PruneDict(ByRef Dict As Dictionary(Of String, Object()))
        For Each KV As KeyValuePair(Of String, Object()) In Dict
            PruneDict(CType(KV.Value(0), Dictionary(Of String, Object())))
            Dim Clone As New Dictionary(Of String, Object())
            Clone = CloneDict(CType(KV.Value(1), Dictionary(Of String, Object())))
            'Intersect all children
            For Each CheckChildren As KeyValuePair(Of String, Object()) In CType(KV.Value(0), Dictionary(Of String, Object()))
                IntersectDict(Clone, CType(CheckChildren.Value(0), Dictionary(Of String, Object())))
            Next
            'Compare and prune intersections at child level
            'cannot trim non-intersections at parent level unless termination markers are included
            For Each CheckChildren As KeyValuePair(Of String, Object()) In CType(KV.Value(0), Dictionary(Of String, Object()))
                TrimDict(CType(CheckChildren.Value(0), Dictionary(Of String, Object())), Clone)
            Next
        Next
    End Sub
    Public Shared Sub AddRecDictionary(ByRef Dict As Dictionary(Of String, Object()), Anchor As String, Str As String, OthStr As String, Post As Boolean)
        If Not Dict.ContainsKey(Anchor) Then
            Dict.Add(Anchor, {New Dictionary(Of String, Object()), New Dictionary(Of String, Object())})
        End If
        Dim CurDict As Dictionary(Of String, Object()) = CType(Dict(Anchor)(0), Dictionary(Of String, Object()))
        Dim CurOthDict As Dictionary(Of String, Object()) = CType(Dict(Anchor)(1), Dictionary(Of String, Object()))
        For StrCount = If(Post, 0, Str.Length - 1) To If(Post, Str.Length - 1, 0) Step If(Post, 1, -1)
            If Not CurDict.ContainsKey(Str(StrCount)) Then
                CurDict.Add(Str(StrCount), {New Dictionary(Of String, Object()), New Dictionary(Of String, Object())})
            End If
            For StrOthCount = If(Post, OthStr.Length - 1, 0) To If(Post, 0, OthStr.Length - 1) Step If(Post, -1, 1)
                If Not CurOthDict.ContainsKey(OthStr(StrOthCount)) Then
                    CurOthDict.Add(OthStr(StrOthCount), {New Dictionary(Of String, Object()), Nothing})
                End If
                CurOthDict = CType(CurOthDict(OthStr(StrOthCount))(0), Dictionary(Of String, Object()))
            Next
            CurOthDict = CType(CurDict(Str(StrCount))(1), Dictionary(Of String, Object()))
            CurDict = CType(CurDict(Str(StrCount))(0), Dictionary(Of String, Object()))
        Next
    End Sub
    Public Shared Function PatternAnalysis() As String()
        Dim Verses As List(Of String()) = GetQuranText(CachedData.XMLDocMain, -1, -1, -1, -1)
        Dim PreDict As Dictionary(Of String, Object())
        Dim PostDict As Dictionary(Of String, Object())
        Dim Strings(CachedData.RecitationSymbols.Length - 1) As String
        For LetCount = 0 To CachedData.RecitationSymbols.Length - 1
            PreDict = New Dictionary(Of String, Object())
            PostDict = New Dictionary(Of String, Object())
            For Count As Integer = 0 To Verses.Count - 1
                For SubCount As Integer = 0 To Verses(Count).Length - 1
                    Dim Matches As System.Text.RegularExpressions.MatchCollection = System.Text.RegularExpressions.Regex.Matches(Verses(Count)(SubCount), If(CachedData.RecitationSymbols(LetCount) = " ", "(\S*)(^|\s+|$)(\S*)", "(\S*)(" + ArabicData.MakeUniRegEx(CachedData.RecitationSymbols(LetCount)) + ")(\S*)"))
                    For MatchCount As Integer = 0 To Matches.Count - 1
                        AddRecDictionary(PostDict, Matches(MatchCount).Groups(3).Value(0), Matches(MatchCount).Groups(3).Value.Substring(1), Matches(MatchCount).Groups(1).Value, True)
                        AddRecDictionary(PreDict, Matches(MatchCount).Groups(1).Value(0), Matches(MatchCount).Groups(1).Value.Substring(0, Matches(MatchCount).Groups(1).Value.Length - 1), Matches(MatchCount).Groups(3).Value, False)
                    Next
                Next
            Next
            Strings(LetCount) = ArabicData.LeftToRightEmbedding + DumpRecDictionary(PreDict, False, 0, True) + "\" + Arabic.TransliterateToScheme(CachedData.RecitationSymbols(LetCount), ArabicData.TranslitScheme.Literal, String.Empty) + "\" + DumpRecDictionary(PostDict, True, 0, True) + ArabicData.PopDirectionalFormatting
        Next
        Return Strings
    End Function
    Public Shared Function GetQuranHamzaMaddDoubleLetterPatterns() As String()
        Dim CurPat As String = ArabicData.MakeRegMultiEx(CachedData.RecitationSymbols)
        Dim Prefixes As New Dictionary(Of String, ArrayList)
        Dim Suffixes As New Dictionary(Of String, ArrayList)
        Dim PreMidSuf As New Dictionary(Of String, Object()) 'Prefix indexed
        Dim SufMidPre As New Dictionary(Of String, Object()) 'Suffix indexed
        For Each Key As String In CachedData.FormDictionary.Keys
            Dim Matches As System.Text.RegularExpressions.MatchCollection = System.Text.RegularExpressions.Regex.Matches(Arabic.TransliterateFromBuckwalter(Key), CurPat)
            For Count As Integer = 0 To Matches.Count - 1
                If Matches(Count).Index = 0 Then
                    For SubCount As Integer = 0 To CachedData.FormDictionary(Key).Count - 1
                        Dim Loc(3) As Integer
                        CType(CachedData.FormDictionary(Key)(SubCount), Integer()).CopyTo(Loc, 0)
                        If (Not CBool(CachedData.LocDictionary(String.Join(":", Loc))(1)) And Not CBool(CachedData.LocDictionary(String.Join(":", Loc))(2))) Or CStr(CachedData.LocDictionary(String.Join(":", Loc))(3)) = "DET" Or (Matches(Count).Value = ArabicData.ArabicLetterAlefWasla And CStr(CachedData.LocDictionary(String.Join(":", Loc))(3)) = "PN") Then
                            Dim Pre As String = String.Empty
                            For LocCount = 1 To CType(CachedData.FormDictionary(Key)(SubCount), Integer())(3) - 1
                                Loc(3) = LocCount
                                Pre += CStr(CachedData.LocDictionary(String.Join(":", Loc))(0))
                            Next
                            If Pre <> String.Empty Then
                                Loc(3) = CType(CachedData.FormDictionary(Key)(SubCount), Integer())(3)
                                If CStr(CachedData.LocDictionary(String.Join(":", Loc))(3)) = "DET" Or (Matches(Count).Value = ArabicData.ArabicLetterAlefWasla And CStr(CachedData.LocDictionary(String.Join(":", Loc))(3)) = "PN") Then
                                    If Not Prefixes.ContainsKey("Al+") Then Prefixes.Add("Al+", New ArrayList)
                                    Prefixes("Al+").Add(StrReverse(Pre))
                                Else
                                    If Not Prefixes.ContainsKey(Matches(Count).Value) Then Prefixes.Add(Matches(Count).Value, New ArrayList)
                                    Prefixes(Matches(Count).Value).Add(StrReverse(Pre))
                                End If
                                If Matches(Count).Value <> ArabicData.ArabicLetterAlefWasla And CStr(CachedData.LocDictionary(String.Join(":", Loc))(3)) <> "DET" Then
                                    If Not Prefixes.ContainsKey("!" + ArabicData.ArabicLetterAlefWasla) Then Prefixes.Add("!" + ArabicData.ArabicLetterAlefWasla, New ArrayList)
                                    Prefixes("!" + ArabicData.ArabicLetterAlefWasla).Add(StrReverse(Pre))
                                End If
                            End If
                        End If
                    Next
                ElseIf Matches(Count).Index = Key.Length - 1 Then
                    For SubCount As Integer = 0 To CachedData.FormDictionary(Key).Count - 1
                        Dim Loc(3) As Integer
                        CType(CachedData.FormDictionary(Key)(SubCount), Integer()).CopyTo(Loc, 0)
                        If (Not CBool(CachedData.LocDictionary(String.Join(":", Loc))(1)) And Not CBool(CachedData.LocDictionary(String.Join(":", Loc))(2))) Then
                            Dim Sup As String = String.Empty
                            Dim LocCount As Integer = CType(CachedData.FormDictionary(Key)(SubCount), Integer())(3) + 1
                            Do
                                Loc(3) = LocCount
                                If Not CachedData.LocDictionary.ContainsKey(String.Join(":", Loc)) Then Exit Do
                                Sup += CStr(CachedData.LocDictionary(String.Join(":", Loc))(0))
                                LocCount += 1
                            Loop While True
                            If Sup <> String.Empty Then
                                If Not Suffixes.ContainsKey(Matches(Count).Value) Then Suffixes.Add(Matches(Count).Value, New ArrayList)
                                Suffixes(Matches(Count).Value).Add(Sup)
                            End If
                        End If
                    Next
                End If
                If Matches(Count).Value = ArabicData.ArabicLetterHamza Or Matches(Count).Value = ArabicData.ArabicLetterAlefWithHamzaAbove Or Matches(Count).Value = ArabicData.ArabicLetterAlefWithHamzaBelow Or Matches(Count).Value = ArabicData.ArabicLetterWawWithHamzaAbove Or Matches(Count).Value = ArabicData.ArabicLetterYehWithHamzaAbove Or Matches(Count).Value = ArabicData.ArabicHamzaAbove Then
                    'At the end of the word, or in the middle of the word when followed by sukun, look before
                    'At the beginning of the word regardless of prefix, or in the middle of the word when not followed by sukun look after
                    For SubCount As Integer = 0 To CachedData.FormDictionary(Key).Count - 1
                        Dim PreCheck As String = String.Empty
                        Dim Loc(3) As Integer
                        CType(CachedData.FormDictionary(Key)(SubCount), Integer()).CopyTo(Loc, 0)
                        'Hamza prefix then must look before
                        'If PreCheck.IndexOfAny({ArabicData.ArabicLetterHamza, ArabicData.ArabicLetterAlefWithHamzaAbove, ArabicData.ArabicLetterAlefWithHamzaBelow, ArabicData.ArabicLetterWawWithHamzaAbove, ArabicData.ArabicLetterYehWithHamzaAbove, ArabicData.ArabicHamzaAbove}) = -1 Then PreCheck = String.Empty
                        Dim Pre As String = String.Empty
                        Dim LocCount As Integer
                        PreCheck += Arabic.TransliterateFromBuckwalter(Key.Substring(0, Matches(Count).Index))
                        For SupCount As Integer = PreCheck.Length - 1 To 0 Step -1
                            Pre = PreCheck(SupCount) + Pre
                            If Array.IndexOf(CachedData.ArabicSunLetters, CStr(PreCheck(SupCount))) <> -1 Or Array.IndexOf(CachedData.ArabicMoonLettersNoVowels, CStr(PreCheck(SupCount))) <> -1 Then Exit For
                        Next
                        If Pre.Length = PreCheck.Length Then
                            PreCheck = String.Empty
                            For LocCount = 1 To CType(CachedData.FormDictionary(Key)(SubCount), Integer())(3) - 1
                                Loc(3) = LocCount
                                PreCheck += Arabic.TransliterateFromBuckwalter(CStr(CachedData.LocDictionary(String.Join(":", Loc))(0)))
                            Next
                            If PreCheck <> String.Empty Then Pre = ";" + PreCheck + ";" + Pre
                        End If
                        Dim Sup As String = Arabic.TransliterateFromBuckwalter(Key.Substring(Matches(Count).Index + 1))
                        LocCount = CType(CachedData.FormDictionary(Key)(SubCount), Integer())(3) + 1
                        Do
                            Loc(3) = LocCount
                            If Not CachedData.LocDictionary.ContainsKey(String.Join(":", Loc)) Then Exit Do
                            Sup += Arabic.TransliterateFromBuckwalter(CStr(CachedData.LocDictionary(String.Join(":", Loc))(0)))
                            LocCount += 1
                        Loop While True
                        Dim Suf As String = String.Empty
                        For SufCount As Integer = 0 To Sup.Length - 1
                            Suf += Sup(SufCount)
                            If Array.IndexOf(CachedData.ArabicSunLetters, CStr(Sup(SufCount))) <> -1 Or Array.IndexOf(CachedData.ArabicMoonLettersNoVowels, CStr(Sup(SufCount))) <> -1 Or ArabicData.ArabicLetterTehMarbuta = Sup(SufCount) Then Exit For
                        Next
                        If (Suf(0) <> ArabicData.ArabicSukun Or Matches(Count).Index = 0) And Suf.Length <> Sup.Length Then
                            AddRecDictionary(SufMidPre, If(Pre.Length <> 0 AndAlso Pre(Pre.Length - 1) = ArabicData.ArabicTatweel, ArabicData.ArabicTatweel, String.Empty) + Matches(Count).Value, Suf, If(Pre.Length <> 0 AndAlso Pre(Pre.Length - 1) = ArabicData.ArabicTatweel, Pre.Substring(0, Pre.Length - 1), Pre), True)
                        Else
                            Debug.Assert(Suf(0) = ArabicData.ArabicSukun Or Suf.Length = Sup.Length)
                            AddRecDictionary(PreMidSuf, If(Pre.Length <> 0 AndAlso Pre(Pre.Length - 1) = ArabicData.ArabicTatweel, ArabicData.ArabicTatweel, String.Empty) + Matches(Count).Value, If(Pre.Length <> 0 AndAlso Pre(Pre.Length - 1) = ArabicData.ArabicTatweel, Pre.Substring(0, Pre.Length - 1), Pre), Suf, False)
                        End If
                    Next
                End If
            Next
        Next
        'PreMidSuf.Clear()
        For Each Key As String In CachedData.RealWordDictionary.Keys
            Dim Matches As System.Text.RegularExpressions.MatchCollection = System.Text.RegularExpressions.Regex.Matches(Key, CurPat)
            For Count As Integer = 0 To Matches.Count - 1
                'If Matches(Count).Value = ArabicData.ArabicLetterHamza Or Matches(Count).Value = ArabicData.ArabicLetterAlefWithHamzaAbove Or Matches(Count).Value = ArabicData.ArabicLetterAlefWithHamzaBelow Or Matches(Count).Value = ArabicData.ArabicLetterWawWithHamzaAbove Or Matches(Count).Value = ArabicData.ArabicLetterYehWithHamzaAbove Or Matches(Count).Value = ArabicData.ArabicHamzaAbove Then
                'If Matches(Count).Value = ArabicData.ArabicLetterAlef Or Matches(Count).Value = ArabicData.ArabicLetterWaw Or Matches(Count).Value = ArabicData.ArabicLetterYeh Or Matches(Count).Value = ArabicData.ArabicLetterAlefMaksura Or Matches(Count).Value = ArabicData.ArabicSmallWaw Or Matches(Count).Value = ArabicData.ArabicSmallYeh Or Matches(Count).Value = ArabicData.ArabicLetterSuperscriptAlef Then
                If Matches(Count).Index = Key.Length - 1 And (Array.IndexOf(CachedData.ArabicSunLetters, Matches(Count).Value) = -1 And Array.IndexOf(CachedData.ArabicMoonLettersNoVowels, Matches(Count).Value) = -1 And ArabicData.ArabicLetterTehMarbuta <> Matches(Count).Value) Then
                    Dim Pre As String = String.Empty
                    For SubCount As Integer = Matches(Count).Index - 1 To 0 Step -1
                        If Array.IndexOf(CachedData.ArabicSunLetters, CStr(Key(SubCount))) = -1 And Array.IndexOf(CachedData.ArabicMoonLettersNoVowels, CStr(Key(SubCount))) = -1 Then
                            Pre = Key(SubCount) + Pre
                        Else
                            Exit For
                        End If
                    Next
                    Dim Suf As String = String.Empty
                    For SubCount As Integer = Matches(Count).Index + 1 To Key.Length - 1
                        Suf += Key(SubCount)
                        If Array.IndexOf(CachedData.ArabicSunLetters, CStr(Key(SubCount))) <> -1 Or Array.IndexOf(CachedData.ArabicMoonLettersNoVowels, CStr(Key(SubCount))) <> -1 Or ArabicData.ArabicLetterTehMarbuta = Key(SubCount) Then Exit For
                    Next
                    AddRecDictionary(PreMidSuf, Matches(Count).Value, Pre, Suf, False)
                End If
            Next
        Next
        Dim Strings(Prefixes.Count + Suffixes.Count + PreMidSuf.Count + SufMidPre.Count - 1) As String
        Dim CurNum As Integer = 0
        For Each Key As String In Prefixes.Keys
            Prefixes(Key).Sort()
            Dim Pres As String = String.Empty
            For Count = 0 To Prefixes(Key).Count - 1
                If Count = Prefixes(Key).Count - 1 OrElse CStr(Prefixes(Key)(Count)) <> CStr(Prefixes(Key)(Count + 1)) Then
                    If Pres <> String.Empty Then Pres += " / "
                    Pres += StrReverse(CStr(Prefixes(Key)(Count)))
                End If
            Next
            Strings(CurNum) = ArabicData.LeftToRightEmbedding + "Pre: " + Pres + " Key: " + Key + ArabicData.PopDirectionalFormatting
            CurNum += 1
        Next
        For Each Key As String In Suffixes.Keys
            Suffixes(Key).Sort()
            Dim Pres As String = String.Empty
            For Count = 0 To Suffixes(Key).Count - 1
                If Count = Suffixes(Key).Count - 1 OrElse CStr(Suffixes(Key)(Count)) <> CStr(Suffixes(Key)(Count + 1)) Then
                    If Pres <> String.Empty Then Pres += " / "
                    Pres += CStr(Suffixes(Key)(Count))
                End If
            Next
            Strings(CurNum) = ArabicData.LeftToRightEmbedding + "Key: " + Key + " Suf: " + Pres + ArabicData.PopDirectionalFormatting
            CurNum += 1
        Next
        For Each Key As String In PreMidSuf.Keys
            Strings(CurNum) = ArabicData.LeftToRightEmbedding + "Key:" + Key + "\" + Arabic.TransliterateToScheme(Key, ArabicData.TranslitScheme.Literal, String.Empty) + "\" + DumpRecDictionary(CType(PreMidSuf(Key)(0), Dictionary(Of String, Object())), False, 0, True) + "\" + ArabicData.PopDirectionalFormatting
            CurNum += 1
        Next
        For Each Key As String In SufMidPre.Keys
            Strings(CurNum) = ArabicData.LeftToRightEmbedding + "Key:" + Key + "\" + Arabic.TransliterateToScheme(Key, ArabicData.TranslitScheme.Literal, String.Empty) + "\" + DumpRecDictionary(CType(SufMidPre(Key)(0), Dictionary(Of String, Object())), False, 0, True) + "\" + ArabicData.PopDirectionalFormatting
            CurNum += 1
        Next
        Return Strings
    End Function
    Public Shared Function GetQuranLetterPatterns() As String()
        Dim RecSymbols As String = String.Join(String.Empty, Array.ConvertAll(CachedData.RecitationSpecialSymbols, Function(C As String) C))
        Dim LtrSymbols As String = String.Join(String.Empty, Array.ConvertAll(CachedData.RecitationLetters, Function(C As String) C))
        Dim DiaSymbols As String = String.Join(String.Empty, Array.ConvertAll(CachedData.RecitationDiacritics, Function(C As String) C))
        Dim StartWordMultiOnly As New Generic.Dictionary(Of String, String)
        Dim EndWordMultiOnly As New Generic.Dictionary(Of String, String)
        Dim MiddleWordMultiOnly As New Generic.Dictionary(Of String, String)
        Dim StartWordOnly As String = String.Join(String.Empty, Array.ConvertAll(CachedData.RecitationLettersDiacritics, Function(C As String) C))
        Dim NotStartWord As String = String.Join(String.Empty, Array.ConvertAll(CachedData.RecitationLettersDiacritics, Function(C As String) C))
        Dim EndWordOnly As String = String.Join(String.Empty, Array.ConvertAll(CachedData.RecitationLettersDiacritics, Function(C As String) C))
        Dim NotEndWord As String = String.Join(String.Empty, Array.ConvertAll(CachedData.RecitationLettersDiacritics, Function(C As String) C))
        Dim EndWordOnlyNoDia As String = String.Join(String.Empty, Array.ConvertAll(CachedData.RecitationLetters, Function(C As String) C))
        Dim NotEndWordNoDia As String = String.Join(String.Empty, Array.ConvertAll(CachedData.RecitationLetters, Function(C As String) C))
        Dim MiddleWordOnlyNoDia As String = String.Join(String.Empty, Array.ConvertAll(CachedData.RecitationLetters, Function(C As String) C))
        Dim NotMiddleWordNoDia As String = String.Join(String.Empty, Array.ConvertAll(CachedData.RecitationLetters, Function(C As String) C))
        Dim MiddleWordOnly As String = String.Join(String.Empty, Array.ConvertAll(CachedData.RecitationLettersDiacritics, Function(C As String) C))
        Dim NotMiddleWord As String = String.Join(String.Empty, Array.ConvertAll(CachedData.RecitationLettersDiacritics, Function(C As String) C))
        Dim DiaStartWordOnly As String = String.Join(String.Empty, Array.ConvertAll(CachedData.RecitationDiacritics, Function(C As String) C))
        Dim DiaNotStartWord As String = String.Join(String.Empty, Array.ConvertAll(CachedData.RecitationDiacritics, Function(C As String) C))
        Dim DiaEndWordOnly As String = String.Join(String.Empty, Array.ConvertAll(CachedData.RecitationDiacritics, Function(C As String) C))
        Dim DiaNotEndWord As String = String.Join(String.Empty, Array.ConvertAll(CachedData.RecitationDiacritics, Function(C As String) C))
        Dim DiaMiddleWordOnly As String = String.Join(String.Empty, Array.ConvertAll(CachedData.RecitationDiacritics, Function(C As String) C))
        Dim DiaNotMiddleWord As String = String.Join(String.Empty, Array.ConvertAll(CachedData.RecitationDiacritics, Function(C As String) C))
        Dim Combos As String() = String.Join("|", Array.ConvertAll(CachedData.RecitationLettersDiacritics, Function(C As String) String.Join("|", Array.ConvertAll(CachedData.RecitationLettersDiacritics, Function(Nxt As String) C + Nxt)))).Split("|"c)
        Dim DiaCombos As String() = String.Join("|", Array.ConvertAll(CachedData.RecitationDiacritics, Function(C As String) String.Join("|", Array.ConvertAll(CachedData.RecitationDiacritics, Function(Nxt As String) C + Nxt)))).Split("|"c)
        Dim LetCombos As String() = String.Join("|", Array.ConvertAll(CachedData.RecitationLetters, Function(C As String) String.Join("|", Array.ConvertAll(CachedData.RecitationLetters, Function(Nxt As String) C + Nxt)))).Split("|"c)
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
        Dim Val As String = ArabicData.LeftToRightEmbedding + "Combinations: " + ArabicData.PopDirectionalFormatting
        For Each Key As Char In Dict.Keys
            If Dict.Item(Key).Length > (DiaSymbols.Length + LtrSymbols.Length) / 2 Then
                Val += ArabicData.FixStartingCombiningSymbol(Key) + ArabicData.LeftToRightEmbedding + " [" + ArabicData.PopDirectionalFormatting + String.Join(" ", Array.ConvertAll(New String(Array.FindAll((DiaSymbols + LtrSymbols).ToCharArray(), Function(C As Char) Not Dict.Item(Key).Contains(C))).ToCharArray(), Function(C As Char) ArabicData.FixStartingCombiningSymbol(CStr(C)))) + ArabicData.LeftToRightEmbedding + " ]" + ArabicData.PopDirectionalFormatting + vbTab
            Else
                Val += ArabicData.FixStartingCombiningSymbol(Key) + ArabicData.LeftToRightEmbedding + " ! [ " + ArabicData.PopDirectionalFormatting + String.Join(" ", Array.ConvertAll(Dict.Item(Key).ToCharArray(), Function(C As Char) ArabicData.FixStartingCombiningSymbol(CStr(C)))) + ArabicData.LeftToRightEmbedding + " ]" + ArabicData.PopDirectionalFormatting + vbTab
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
        Dim RevVal As String = ArabicData.LeftToRightEmbedding + "Reverse Combinations: " + ArabicData.PopDirectionalFormatting
        For Each Key As Char In RevDict.Keys
            If RevDict.Item(Key).Length > (DiaSymbols.Length + LtrSymbols.Length) / 2 Then
                RevVal += ArabicData.LeftToRightEmbedding + "[" + ArabicData.PopDirectionalFormatting + String.Join(" ", Array.ConvertAll(New String(Array.FindAll((DiaSymbols + LtrSymbols).ToCharArray(), Function(C As Char) Not RevDict.Item(Key).Contains(C))).ToCharArray(), Function(C As Char) ArabicData.FixStartingCombiningSymbol(CStr(C)))) + ArabicData.LeftToRightEmbedding + " ] " + ArabicData.PopDirectionalFormatting + ArabicData.FixStartingCombiningSymbol(Key) + vbTab
            Else
                RevVal += ArabicData.LeftToRightEmbedding + "! [ " + ArabicData.PopDirectionalFormatting + String.Join(" ", Array.ConvertAll(RevDict.Item(Key).ToCharArray(), Function(C As Char) ArabicData.FixStartingCombiningSymbol(CStr(C)))) + ArabicData.LeftToRightEmbedding + " ] " + ArabicData.PopDirectionalFormatting + ArabicData.FixStartingCombiningSymbol(Key) + vbTab
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
        Dim DiaVal As String = ArabicData.LeftToRightEmbedding + "Diacritic Only Combinations: " + ArabicData.PopDirectionalFormatting
        For Each Key As Char In DiaDict.Keys
            If DiaDict.Item(Key).Length > DiaSymbols.Length / 2 Then
                DiaVal += ArabicData.FixStartingCombiningSymbol(Key) + ArabicData.LeftToRightEmbedding + " [" + ArabicData.PopDirectionalFormatting + String.Join(" ", Array.ConvertAll(New String(Array.FindAll(DiaSymbols.ToCharArray(), Function(C As Char) Not DiaDict.Item(Key).Contains(C))).ToCharArray(), Function(C As Char) ArabicData.FixStartingCombiningSymbol(CStr(C)))) + ArabicData.LeftToRightEmbedding + " ]" + ArabicData.PopDirectionalFormatting + vbTab
            Else
                DiaVal += ArabicData.FixStartingCombiningSymbol(Key) + ArabicData.LeftToRightEmbedding + " ! [ " + ArabicData.PopDirectionalFormatting + String.Join(" ", Array.ConvertAll(DiaDict.Item(Key).ToCharArray(), Function(C As Char) ArabicData.FixStartingCombiningSymbol(CStr(C)))) + ArabicData.LeftToRightEmbedding + " ]" + ArabicData.PopDirectionalFormatting + vbTab
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
        Dim LetVal As String = ArabicData.LeftToRightEmbedding + "Letter Only Combinations: " + ArabicData.PopDirectionalFormatting
        For Each Key As Char In LetDict.Keys
            If LetDict.Item(Key).Length > LtrSymbols.Length / 2 Then
                LetVal += ArabicData.FixStartingCombiningSymbol(Key) + ArabicData.LeftToRightEmbedding + " [" + ArabicData.PopDirectionalFormatting + String.Join(" ", Array.ConvertAll(New String(Array.FindAll(LtrSymbols.ToCharArray(), Function(C As Char) Not LetDict.Item(Key).Contains(C))).ToCharArray(), Function(C As Char) ArabicData.FixStartingCombiningSymbol(CStr(C)))) + ArabicData.LeftToRightEmbedding + " ]" + ArabicData.PopDirectionalFormatting + vbTab
            Else
                LetVal += ArabicData.FixStartingCombiningSymbol(Key) + ArabicData.LeftToRightEmbedding + " ! [ " + ArabicData.PopDirectionalFormatting + String.Join(" ", Array.ConvertAll(LetDict.Item(Key).ToCharArray(), Function(C As Char) ArabicData.FixStartingCombiningSymbol(CStr(C)))) + ArabicData.LeftToRightEmbedding + " ]" + ArabicData.PopDirectionalFormatting + vbTab
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
        Dim LetRevVal As String = ArabicData.LeftToRightEmbedding + "Reverse Letter Only Combinations: " + ArabicData.PopDirectionalFormatting
        For Each Key As Char In LetRevDict.Keys
            If LetRevDict.Item(Key).Length > LtrSymbols.Length / 2 Then
                LetRevVal += ArabicData.LeftToRightEmbedding + "[" + ArabicData.PopDirectionalFormatting + String.Join(" ", Array.ConvertAll(New String(Array.FindAll(LtrSymbols.ToCharArray(), Function(C As Char) Not LetRevDict.Item(Key).Contains(C))).ToCharArray(), Function(C As Char) ArabicData.FixStartingCombiningSymbol(CStr(C)))) + ArabicData.LeftToRightEmbedding + " ] " + ArabicData.PopDirectionalFormatting + ArabicData.FixStartingCombiningSymbol(Key) + vbTab
            Else
                LetRevVal += ArabicData.LeftToRightEmbedding + "! [ " + ArabicData.PopDirectionalFormatting + String.Join(" ", Array.ConvertAll(LetRevDict.Item(Key).ToCharArray(), Function(C As Char) ArabicData.FixStartingCombiningSymbol(CStr(C)))) + ArabicData.LeftToRightEmbedding + " ] " + ArabicData.PopDirectionalFormatting + ArabicData.FixStartingCombiningSymbol(Key) + vbTab
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
        Return {ArabicData.LeftToRightEmbedding + "Unique Prefix: [" + ArabicData.PopDirectionalFormatting + StartMulti + ArabicData.LeftToRightEmbedding + "]" + ArabicData.PopDirectionalFormatting, _
                ArabicData.LeftToRightEmbedding + "Unique Suffix: [" + ArabicData.PopDirectionalFormatting + EndMulti + ArabicData.LeftToRightEmbedding + "]" + ArabicData.PopDirectionalFormatting, _
                ArabicData.LeftToRightEmbedding + "Unique Middle: [" + ArabicData.PopDirectionalFormatting + MiddleMulti + ArabicData.LeftToRightEmbedding + "]" + ArabicData.PopDirectionalFormatting, _
                ArabicData.LeftToRightEmbedding + "Start Only: [" + ArabicData.PopDirectionalFormatting + String.Join(" ", Array.ConvertAll(StartWordOnly.ToCharArray(), Function(C As Char) ArabicData.FixStartingCombiningSymbol(CStr(C)))) + ArabicData.LeftToRightEmbedding + "]" + ArabicData.PopDirectionalFormatting, _
                ArabicData.LeftToRightEmbedding + "Not Start: [" + ArabicData.PopDirectionalFormatting + String.Join(" ", Array.ConvertAll(NotStartWord.ToCharArray(), Function(C As Char) ArabicData.FixStartingCombiningSymbol(CStr(C)))) + ArabicData.LeftToRightEmbedding + "]" + ArabicData.PopDirectionalFormatting, _
                ArabicData.LeftToRightEmbedding + "End Only: [" + ArabicData.PopDirectionalFormatting + String.Join(" ", Array.ConvertAll(EndWordOnly.ToCharArray(), Function(C As Char) ArabicData.FixStartingCombiningSymbol(CStr(C)))) + ArabicData.LeftToRightEmbedding + "]" + ArabicData.PopDirectionalFormatting, _
                ArabicData.LeftToRightEmbedding + "Not End: [" + ArabicData.PopDirectionalFormatting + String.Join(" ", Array.ConvertAll(NotEndWord.ToCharArray(), Function(C As Char) ArabicData.FixStartingCombiningSymbol(CStr(C)))) + ArabicData.LeftToRightEmbedding + "]" + ArabicData.PopDirectionalFormatting, _
                ArabicData.LeftToRightEmbedding + "End Only No Diacritics: [" + ArabicData.PopDirectionalFormatting + String.Join(" ", Array.ConvertAll(EndWordOnlyNoDia.ToCharArray(), Function(C As Char) ArabicData.FixStartingCombiningSymbol(CStr(C)))) + ArabicData.LeftToRightEmbedding + "]" + ArabicData.PopDirectionalFormatting, _
                ArabicData.LeftToRightEmbedding + "Not End No Diacritics: [" + ArabicData.PopDirectionalFormatting + String.Join(" ", Array.ConvertAll(NotEndWordNoDia.ToCharArray(), Function(C As Char) ArabicData.FixStartingCombiningSymbol(CStr(C)))) + ArabicData.LeftToRightEmbedding + "]" + ArabicData.PopDirectionalFormatting, _
                ArabicData.LeftToRightEmbedding + "Middle Only: [" + ArabicData.PopDirectionalFormatting + String.Join(" ", Array.ConvertAll(MiddleWordOnly.ToCharArray(), Function(C As Char) ArabicData.FixStartingCombiningSymbol(CStr(C)))) + ArabicData.LeftToRightEmbedding + "]" + ArabicData.PopDirectionalFormatting, _
                ArabicData.LeftToRightEmbedding + "Not Middle: [" + ArabicData.PopDirectionalFormatting + String.Join(" ", Array.ConvertAll(NotMiddleWord.ToCharArray(), Function(C As Char) ArabicData.FixStartingCombiningSymbol(CStr(C)))) + ArabicData.LeftToRightEmbedding + "]" + ArabicData.PopDirectionalFormatting, _
                ArabicData.LeftToRightEmbedding + "Middle Only No Diacritics: [" + ArabicData.PopDirectionalFormatting + String.Join(" ", Array.ConvertAll(MiddleWordOnlyNoDia.ToCharArray(), Function(C As Char) ArabicData.FixStartingCombiningSymbol(CStr(C)))) + ArabicData.LeftToRightEmbedding + "]" + ArabicData.PopDirectionalFormatting, _
                ArabicData.LeftToRightEmbedding + "Not Middle No Diacritics: [" + ArabicData.PopDirectionalFormatting + String.Join(" ", Array.ConvertAll(NotMiddleWordNoDia.ToCharArray(), Function(C As Char) ArabicData.FixStartingCombiningSymbol(CStr(C)))) + ArabicData.LeftToRightEmbedding + "]" + ArabicData.PopDirectionalFormatting, _
                Val, RevVal, DiaVal, LetVal, LetRevVal}
    End Function
    Public Shared Function GetMetaRuleSet(Name As String) As IslamData.RuleMetaSet
        For Count = 0 To CachedData.IslamData.MetaRules.Length - 1
            If CachedData.IslamData.MetaRules(Count).Name = Name Then Return CachedData.IslamData.MetaRules(Count)
        Next
        Return Nothing
    End Function
    Public Shared Function GetRecitationRules() As Array()
        Dim Names() As Array = Array.ConvertAll(GetMetaRuleSet("UthmaniQuran").Rules, Function(Convert As IslamData.RuleMetaSet.RuleMetadataTranslation) New Object() {Utility.LoadResourceString("IslamInfo_" + Convert.Name), CInt(Array.IndexOf(CachedData.IslamData.MetaRules, Convert))})
        Array.Sort(Names, New Utility.CompareNameValueArray)
        Return Names
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
        ElseIf Division = 10 Then
            Return GetRecitationRules()
        End If
        Return Nothing
    End Function
    Public Enum QuranScripts
        Uthmani = 0
        UthmaniMin = 1
        Simple = 2
        SimpleMin = 3
        SimpleEnhanced = 4
        SimpleClean = 5
    End Enum
    Public Enum QuranTexts
        Hafs = 0
        Warsh = 1
        AlDari = 2
    End Enum
    Public Enum ArabicPresentation
        None = 0
        PresentationLigatures = 1
        Buckwalter = 2
    End Enum
    Public Shared QuranTextNames As String() = {"quran-hafs", "quran-warsh", "quran-alduri"}
    Public Shared QuranFileNames As String() = {"uthmani", "uthmani-min", "simple", "simple-min", "simple-enhanced", "simple-clean"}
    Public Shared QuranScriptNames As String() = {"Uthmani", "Uthmani Minimal", "Simple", "Simple Minimal", "Simple Enhanced", "Simple Clean"}
    Public Shared PresentationCacheNames As String() = {String.Empty, "pres", "buckwalter"}
    Public Shared Sub FindMinimalVersesForCoverage()
        Dim Letters As String() = CachedData.RecitationSymbols 'input is letter coverage required 
        'CachedData.ArabicLetters '3:154 and 48:29
        '{ArabicData.ArabicHamzaAbove, ArabicData.ArabicLetterHamza, ArabicData.ArabicLetterAlefWithHamzaAbove, ArabicData.ArabicLetterAlefWithHamzaBelow, ArabicData.ArabicLetterWawWithHamzaAbove, ArabicData.ArabicLetterYehWithHamzaAbove} '2:31, 3:49, 5:41, 8:19, 24:62 
        Dim SubAyah As Boolean = False 'input whether or not to break ayahs into "sub ayah" with 5 tajweed pause markers not including the 6th of forced continue
        Dim VerseDict As New Dictionary(Of String, List(Of Integer(,)))
        Dim Verses As List(Of String()) = GetQuranText(CachedData.XMLDocMain, -1, -1, -1, -1)
        'initialize the letter profile dictionary
        Dim TotalVerses As Integer = 0
        For Count As Integer = 0 To Verses.Count - 1
            For SubCount As Integer = 0 To Verses(Count).Length - 1
                Dim Hits(Letters.Length - 1) As Boolean
                For LetCount As Integer = 0 To Letters.Length - 1
                    If Verses(Count)(SubCount).Contains(Letters(LetCount)) Then
                        Hits(LetCount) = True
                    End If
                Next
                Dim Str As String = String.Empty
                For LetCount As Integer = 0 To Hits.Length - 1
                    If Hits(LetCount) Then Str += Letters(LetCount)
                Next
                If Str.Length = Letters.Length Then Debug.Print(CStr(Count) + ":" + CStr(SubCount))
                If Not VerseDict.ContainsKey(Str) Then VerseDict.Add(Str, New List(Of Integer(,)))
                VerseDict(Str).Add(New Integer(0, 1) {{Count, SubCount}})
            Next
            TotalVerses += GetVerseCount(Count + 1)
        Next
        Dim CombineCount As Integer = 1 'order of n^2*n! which needs to be heuristically reduced towards n^3
        Dim CombineVerseDict As New Dictionary(Of String, List(Of Integer(,)))(VerseDict) 'shallow copy
        Do
            Dim NextCombineVerseDict As New Dictionary(Of String, List(Of Integer(,)))
            For Each LetterProfile As KeyValuePair(Of String, List(Of Integer(,))) In CombineVerseDict
                For Each VerseLetProfile As KeyValuePair(Of String, List(Of Integer(,))) In VerseDict
                    For Count As Integer = 0 To LetterProfile.Value.Count - 1
                        For SubCount As Integer = 0 To VerseLetProfile.Value.Count - 1
                            Dim SearchCount As Integer
                            For SearchCount = 0 To LetterProfile.Value(Count).GetLength(0) - 1
                                If LetterProfile.Value(Count)(SearchCount, 0) = VerseLetProfile.Value(SubCount)(0, 0) And LetterProfile.Value(Count)(SearchCount, 1) = VerseLetProfile.Value(SubCount)(0, 1) Then Exit For
                            Next
                            If SearchCount = LetterProfile.Value(Count).GetLength(0) Then
                                Dim Hits(Letters.Length - 1) As Boolean
                                For LetCount As Integer = 0 To Letters.Length - 1
                                    If LetterProfile.Key.Contains(Letters(LetCount)) Or VerseLetProfile.Key.Contains(Letters(LetCount)) Then
                                        Hits(LetCount) = True
                                    End If
                                Next
                                Dim Str As String = String.Empty
                                For LetCount As Integer = 0 To Hits.Length - 1
                                    If Hits(LetCount) Then Str += Letters(LetCount)
                                Next
                                If Str.Length = Letters.Length Then
                                    Debug.Print(CStr(VerseLetProfile.Value(SubCount)(0, 0)) + ":" + CStr(VerseLetProfile.Value(SubCount)(0, 1)) + ",")
                                    For SearchCount = 0 To LetterProfile.Value(Count).GetLength(0) - 1
                                        Debug.Print(String.Join(",", CStr(LetterProfile.Value(Count)(SearchCount, 0)) + ":" + CStr(LetterProfile.Value(Count)(SearchCount, 1))))
                                    Next
                                End If
                                If Not NextCombineVerseDict.ContainsKey(Str) Then NextCombineVerseDict.Add(Str, New List(Of Integer(,)))
                                Dim NewList As Integer(,) = New Integer(LetterProfile.Value(Count).GetLength(0) + VerseLetProfile.Value(SubCount).GetLength(0) - 1, 1) {}
                                Array.Copy(LetterProfile.Value(Count), NewList, 2 * LetterProfile.Value(Count).GetLength(0))
                                NewList(LetterProfile.Value(Count).GetLength(0), 0) = VerseLetProfile.Value(SubCount)(0, 0)
                                NewList(LetterProfile.Value(Count).GetLength(0), 1) = VerseLetProfile.Value(SubCount)(0, 1)
                                NextCombineVerseDict(Str).Add(NewList)
                            End If
                        Next
                    Next
                Next
            Next
            CombineVerseDict = NextCombineVerseDict
            CombineCount += 1
        Loop While CombineCount < TotalVerses
    End Sub
    Public Shared Sub CheckNotablePatterns()
        'ComparePatterns(QuranScripts.Uthmani, QuranScripts.UthmaniMin, Arabic.UthmaniShortVowelsBeforeLongVowelsAlef)
        'ComparePatterns(QuranScripts.Uthmani, QuranScripts.UthmaniMin, Arabic.UthmaniShortVowelsBeforeLongVowelsSuperscriptAlef)
        'ComparePatterns(QuranScripts.Uthmani, QuranScripts.UthmaniMin, Arabic.UthmaniShortVowelsBeforeLongVowelsAlefAlefMaksura)
        'ComparePatterns(QuranScripts.Uthmani, QuranScripts.UthmaniMin, Arabic.UthmaniShortVowelsBeforeLongVowelsYehAlefMaksura)
        'ComparePatterns(QuranScripts.Uthmani, QuranScripts.UthmaniMin, Arabic.UthmaniShortVowelsBeforeLongVowelsYeh)
        'ComparePatterns(QuranScripts.Uthmani, QuranScripts.UthmaniMin, Arabic.UthmaniShortVowelsBeforeLongVowelsSmallYeh)
        'ComparePatterns(QuranScripts.Uthmani, QuranScripts.UthmaniMin, Arabic.UthmaniShortVowelsBeforeLongVowelsWaw)
        'ComparePatterns(QuranScripts.Uthmani, QuranScripts.UthmaniMin, Arabic.UthmaniShortVowelsBeforeLongVowelsSmallWaw)
        ComparePatterns(QuranTexts.Hafs, QuranScripts.Uthmani, QuranScripts.SimpleEnhanced, Arabic.SimpleTrailingAlef)
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
            Msg += """" + Arabic.TransliterateToScheme(Keys(Count), ArabicData.TranslitScheme.Literal, String.Empty) + """" + If(Count <> Keys.Length - 1, ", ", String.Empty)
        Next
        Return Msg
    End Function
    Public Shared Sub ComparePatterns(BaseText As QuranTexts, ScriptType As QuranScripts, CompScriptType As QuranScripts, LetterPattern As String)
        Dim WordPattern As String = "(?<=^\s*|\s+)\S*" + ArabicData.MakeUniRegEx(LetterPattern) + "\S*(?=\s+|\s*$)"
        Dim FirstList As List(Of String) = PatternMatch(BaseText, ScriptType, ArabicPresentation.None, WordPattern)
        FirstList.Sort(StringComparer.Ordinal)
        Dim CompList As List(Of String) = PatternMatch(BaseText, CompScriptType, ArabicPresentation.None, "(?<=^\s*|\s+)\S*" + ArabicData.MakeUniRegEx(LetterPattern.Substring(0, 1)) + "(?=\s+|\s*$)")
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
            Msg += """" + Arabic.TransliterateToScheme(Str, ArabicData.TranslitScheme.Literal, String.Empty) + """, "
        Next
        Msg += vbCrLf + "Second: "
        For Each Str As String In CompList
            Dim SubKey As String = Str.Substring(0, Str.IndexOf(LetterPattern))
            For Count As Integer = 0 To SubKey.Length - 1
                If Not CompDict.ContainsKey(SubKey.Substring(Count)) Then
                    CompDict.Add(SubKey.Substring(Count), Str)
                End If
            Next
            Msg += """" + Arabic.TransliterateToScheme(Str, ArabicData.TranslitScheme.Literal, String.Empty) + """, "
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
    Public Shared Function PatternMatch(BaseText As QuranTexts, ScriptType As QuranScripts, Presentation As ArabicPresentation, Pattern As String) As List(Of String)
        PatternMatch = New List(Of String)
        Dim Doc As New System.Xml.XmlDocument
        If ScriptType = QuranScripts.Uthmani Then
            Doc.Load(Utility.GetFilePath("metadata\" + QuranTextNames(BaseText) + ".xml"))
        Else
            Doc.Load(Utility.GetFilePath("metadata\" + QuranTextNames(BaseText) + "-" + QuranFileNames(ScriptType) + If(Presentation <> ArabicPresentation.None, "-" + PresentationCacheNames(Presentation), String.Empty) + ".xml"))
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
    Public Shared Sub CompareQuranFormats(BaseText As QuranTexts, TargetBaseText As QuranTexts, ScriptType As QuranScripts, Presentation As ArabicPresentation)
        Dim Doc As New System.Xml.XmlDocument
        Doc.Load(Utility.GetFilePath("metadata\" + QuranTextNames(BaseText) + ".xml"))
        Dim TargetDoc As New System.Xml.XmlDocument
        If BaseText = TargetBaseText Then
            TargetDoc.Load(Utility.GetFilePath("metadata\" + QuranTextNames(TargetBaseText) + "-" + QuranFileNames(ScriptType) + If(Presentation <> ArabicPresentation.None, "-" + PresentationCacheNames(Presentation), String.Empty) + ".xml"))
        Else
            TargetDoc.Load(Utility.GetFilePath("metadata\" + QuranTextNames(TargetBaseText) + "-" + QuranFileNames(ScriptType) + If(Presentation <> ArabicPresentation.None, "-" + PresentationCacheNames(Presentation), String.Empty) + ".xml"))
        End If
        Dim Verses As Collections.Generic.List(Of String())
        Dim TargetVerses As Collections.Generic.List(Of String())
        Verses = TanzilReader.GetQuranText(Doc, -1, -1, -1, -1)
        TargetVerses = TanzilReader.GetQuranText(TargetDoc, -1, -1, -1, -1)
        For Count As Integer = 0 To Verses.Count - 1
            Dim SubCount As Integer = 0
            Dim TargetSubCount As Integer = 0
            Dim Total As Integer = 0
            Dim TargetTotal As Integer = 0
            Do
                Dim Words As String() = {}
                Dim TargetWords As String() = {}
                If Total <= TargetTotal Then
                    If Total <= TargetTotal And SubCount <= Verses(Count).Length - 1 Then
                        Words = Verses(Count)(SubCount).Split(" "c)
                        Total += Words.Length - Array.FindAll(Words, Function(Str As String) Str.Length = 1).Length
                        SubCount += 1
                    End If
                    While (TargetTotal < Total Or TargetTotal = Total And SubCount = Verses(Count).Length) And TargetSubCount <= TargetVerses(Count).Length - 1
                        TargetWords = TargetVerses(Count)(TargetSubCount).Split(" "c)
                        TargetTotal += TargetWords.Length - Array.FindAll(TargetWords, Function(Str As String) Str.Length = 1).Length
                        TargetSubCount += 1
                        If Total <> TargetTotal Then
                            Debug.Print("Chapter: " + CStr(Count + 1) + " Verse: " + CStr(SubCount) + " Words: " + CStr(Words.Length) + " Verse: " + CStr(TargetSubCount) + " Words: " + CStr(TargetWords.Length) + If(Words.Length > TargetWords.Length, "  +", "  -") + CStr(Count + 1) + ":" + CStr(Math.Max(SubCount, TargetSubCount)) + ":" + CStr(Math.Min(Words.Length, TargetWords.Length) + 1))
                        End If
                    End While
                Else
                    If TargetTotal <= Total And TargetSubCount <= TargetVerses(Count).Length - 1 Then
                        TargetWords = TargetVerses(Count)(TargetSubCount).Split(" "c)
                        TargetTotal += TargetWords.Length - Array.FindAll(TargetWords, Function(Str As String) Str.Length = 1).Length
                        TargetSubCount += 1
                    End If
                    While Total < TargetTotal And SubCount <= Verses(Count).Length - 1
                        Words = Verses(Count)(SubCount).Split(" "c)
                        Total += Words.Length - Array.FindAll(Words, Function(Str As String) Str.Length = 1).Length
                        SubCount += 1
                        If Total <> TargetTotal Then
                            Debug.Print("Chapter: " + CStr(Count + 1) + " Verse: " + CStr(SubCount) + " Words: " + CStr(Words.Length) + " Verse: " + CStr(TargetSubCount) + " Words: " + CStr(TargetWords.Length) + If(Words.Length > TargetWords.Length, "  +", "  -") + CStr(Count + 1) + ":" + CStr(Math.Max(SubCount, TargetSubCount)) + ":" + CStr(Math.Min(Words.Length, TargetWords.Length) + 1))
                        End If
                    End While
                End If
            Loop While Total <= TargetTotal And SubCount <= Verses(Count).Length - 1 Or TargetTotal <= Total And TargetSubCount <= TargetVerses(Count).Length - 1
        Next
    End Sub
    Public Shared Sub ChangeQuranFormat(BaseText As QuranTexts, TargetBaseText As QuranTexts, ScriptType As QuranScripts, Presentation As ArabicPresentation)
        Dim Doc As New System.Xml.XmlDocument
        Doc.Load(Utility.GetFilePath("metadata\" + QuranTextNames(BaseText) + ".xml"))
        Dim Verses As Collections.Generic.List(Of String())
        Dim UseBuckwalter As Boolean = False
        Dim Path As String
        If BaseText = TargetBaseText Then
            Path = QuranTextNames(TargetBaseText) + "-" + QuranFileNames(ScriptType) + If(Presentation <> ArabicPresentation.None, "-" + PresentationCacheNames(Presentation), String.Empty) + ".xml"
        Else
            Path = QuranTextNames(TargetBaseText) + "-" + QuranFileNames(ScriptType) + If(Presentation <> ArabicPresentation.None, "-" + PresentationCacheNames(Presentation), String.Empty) + ".xml"
        End If
        If Presentation = ArabicPresentation.Buckwalter Then
            UseBuckwalter = True
        End If
        Dim IndexToVerse As Integer()() = Nothing
        Doc.DocumentElement.PreviousSibling.Value = Doc.DocumentElement.PreviousSibling.Value.Replace("Uthmani", QuranScriptNames(ScriptType))
        Verses = TanzilReader.GetQuranText(Doc, -1, -1, -1, -1)
        For Count As Integer = 0 To Verses.Count - 1
            Dim VerseAdjust As Integer = 0
            Dim ChapterNode As System.Xml.XmlNode = GetTextChapter(Doc, Count + 1)
            If UseBuckwalter Then
                ChapterNode.Attributes.GetNamedItem("name").Value = Arabic.TransliterateToScheme(ChapterNode.Attributes.GetNamedItem("name").Value, ArabicData.TranslitScheme.Literal, String.Empty)
            End If
            Dim SubCount As Integer = 0
            While SubCount <= Verses(Count).Length - 1 - VerseAdjust
                Dim CurVerse As Xml.XmlNode = GetTextVerse(ChapterNode, SubCount + 1)
                Dim PreVerse As String = String.Empty
                Dim NextVerse As String = String.Empty
                If SubCount = 0 AndAlso Not CurVerse.Attributes.GetNamedItem("bismillah") Is Nothing Then
                    If Count <> 0 Then
                        PreVerse = GetTextVerse(GetTextChapter(Doc, Count), GetTextChapter(Doc, Count).ChildNodes.Count - 1).Attributes.GetNamedItem("text").Value
                        If UseBuckwalter Then PreVerse = Arabic.TransliterateFromBuckwalter(PreVerse)
                    End If
                    CurVerse.Attributes.GetNamedItem("bismillah").Value = If(BaseText = TargetBaseText, Arabic.ChangeScript(CurVerse.Attributes.GetNamedItem("bismillah").Value, ScriptType, PreVerse, CurVerse.Attributes.GetNamedItem("text").Value), Arabic.ChangeBaseScript(CurVerse.Attributes.GetNamedItem("bismillah").Value, TargetBaseText, PreVerse, CurVerse.Attributes.GetNamedItem("text").Value))
                    If UseBuckwalter Then
                        CurVerse.Attributes.GetNamedItem("bismillah").Value = Arabic.TransliterateToScheme(CurVerse.Attributes.GetNamedItem("bismillah").Value, ArabicData.TranslitScheme.Literal, String.Empty)
                    End If
                    PreVerse = CurVerse.Attributes.GetNamedItem("bismillah").Value
                End If
                If SubCount + 1 <= Verses(Count).Length - 1 - VerseAdjust Then
                    NextVerse = GetTextVerse(ChapterNode, SubCount + 1 + 1).Attributes.GetNamedItem("text").Value
                ElseIf Count <> Verses.Count - 1 Then
                    NextVerse = GetTextVerse(GetTextChapter(Doc, Count + 1 + 1), 1).Attributes.GetNamedItem(If(GetTextVerse(GetTextChapter(Doc, Count + 1 + 1), 1).Attributes.GetNamedItem("bismillah") Is Nothing, "text", "bismillah")).Value
                End If
                If Count <> 0 AndAlso SubCount = 0 AndAlso CurVerse.Attributes.GetNamedItem("bismillah") Is Nothing Then
                    PreVerse = GetTextVerse(GetTextChapter(Doc, Count), GetTextChapter(Doc, Count).ChildNodes.Count - 1).Attributes.GetNamedItem("text").Value
                ElseIf SubCount <> 0 Then
                    PreVerse = GetTextVerse(ChapterNode, SubCount).Attributes.GetNamedItem("text").Value
                End If
                If UseBuckwalter Then PreVerse = Arabic.TransliterateFromBuckwalter(PreVerse)
                CurVerse.Attributes.GetNamedItem("text").Value = If(BaseText = TargetBaseText, Arabic.ChangeScript(CurVerse.Attributes.GetNamedItem("text").Value, ScriptType, PreVerse, NextVerse), Arabic.ChangeBaseScript(CurVerse.Attributes.GetNamedItem("text").Value, TargetBaseText, PreVerse, NextVerse))
                If UseBuckwalter Then
                    CurVerse.Attributes.GetNamedItem("text").Value = Arabic.TransliterateToScheme(CurVerse.Attributes.GetNamedItem("text").Value, ArabicData.TranslitScheme.Literal, String.Empty)
                End If
                If BaseText = QuranTexts.Hafs And TargetBaseText = QuranTexts.Warsh Then
                    Dim TCount As Integer = Count
                    Dim Index As Integer = Array.FindIndex(CachedData.IslamData.VerseNumberSchemes(0).CombinedVerses, Function(Ints As Integer()) TCount + 1 = Ints(0) And SubCount + 1 + VerseAdjust - 1 = Ints(1))
                    If Index <> -1 Then
                        If Count = 0 And SubCount = 1 Then
                            Dim NewAttr As Xml.XmlAttribute = Doc.CreateAttribute("bismillah")
                            NewAttr.Value = GetTextVerse(ChapterNode, SubCount).Attributes.GetNamedItem("text").Value
                            GetTextVerse(ChapterNode, SubCount).Attributes.InsertAfter(NewAttr, CType(GetTextVerse(ChapterNode, SubCount).Attributes.GetNamedItem("text"), Xml.XmlAttribute))
                            GetTextVerse(ChapterNode, SubCount).Attributes.GetNamedItem("text").Value = CurVerse.Attributes.GetNamedItem("text").Value
                        Else
                            GetTextVerse(ChapterNode, SubCount).Attributes.GetNamedItem("text").Value = GetTextVerse(ChapterNode, SubCount).Attributes.GetNamedItem("text").Value + " " + CurVerse.Attributes.GetNamedItem("text").Value
                        End If
                        CurVerse.ParentNode.RemoveChild(CurVerse)
                        CurVerse = GetTextVerse(ChapterNode, SubCount)
                        VerseAdjust += 1
                        SubCount -= 1
                        For Index = SubCount + 2 To Verses(Count).Length - 1 - VerseAdjust + 1
                            GetTextVerse(ChapterNode, Index + 1).Attributes.GetNamedItem("index").Value = CStr(CInt(GetTextVerse(ChapterNode, Index + 1).Attributes.GetNamedItem("index").Value) - 1)
                        Next
                    End If
                    Index = Array.FindIndex(CachedData.IslamData.VerseNumberSchemes(0).ExtraVerses, Function(Ints As Integer()) TCount + 1 = Ints(0) And SubCount + 1 + VerseAdjust = Ints(1))
                    If Index <> -1 Then
                        Dim NewNode As Xml.XmlNode = CurVerse.Clone()
                        If Not NewNode.Attributes.GetNamedItem("bismillah") Is Nothing Then
                            NewNode.Attributes.RemoveNamedItem("bismillah")
                        End If
                        Index = CachedData.IslamData.VerseNumberSchemes(0).ExtraVerses(Index)(2)
                        While Index <> 1
                            NewNode.Attributes.GetNamedItem("text").Value = NewNode.Attributes.GetNamedItem("text").Value.Substring(NewNode.Attributes.GetNamedItem("text").Value.IndexOf(" "c) + 1)
                            Index -= 1
                        End While
                        CurVerse.Attributes.GetNamedItem("text").Value = CurVerse.Attributes.GetNamedItem("text").Value.Substring(0, CurVerse.Attributes.GetNamedItem("text").Value.Length - NewNode.Attributes.GetNamedItem("text").Value.Length - 1)
                        CurVerse.ParentNode.InsertAfter(NewNode, CurVerse)
                        VerseAdjust -= 1
                        SubCount += 1
                        For Index = Verses(Count).Length - 1 - VerseAdjust - 1 To SubCount Step -1
                            GetTextVerse(ChapterNode, Index + 1).Attributes.GetNamedItem("index").Value = CStr(CInt(GetTextVerse(ChapterNode, Index + 1).Attributes.GetNamedItem("index").Value) + 1)
                        Next
                        NewNode.Attributes.GetNamedItem("index").Value = CStr(CInt(NewNode.Attributes.GetNamedItem("index").Value) + 1)
                    End If
                End If
                SubCount += 1
            End While
        Next
        Doc.Save(Path)
    End Sub
    Public Shared Function IsQuranTextReference(Str As String) As Boolean
        If Not System.Text.RegularExpressions.Regex.Match(Str, "^(?:,?(\d+)(?:\:(\d+))?(?:\:(\d+))?(?:-(\d+)(?:\:(\d+))?(?:\:(\d+))?)?)+$").Success Then Return False
        Dim Matches As System.Text.RegularExpressions.MatchCollection = System.Text.RegularExpressions.Regex.Matches(Str, "(?:,?(\d+)(?:\:(\d+))?(?:\:(\d+))?(?:-(\d+)(?:\:(\d+))?(?:\:(\d+))?)?)")
        For Count = 0 To Matches.Count - 1
            Dim BaseChapter As Integer = CInt(Matches(Count).Groups(1).Value)
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
            If BaseVerse = 0 Then
                BaseVerse += 1
                ExtraVerseNumber = GetTextChapter(CachedData.XMLDocMain, BaseChapter).ChildNodes.Count
            End If
            If WordNumber = 0 Then WordNumber += 1
            If BaseChapter < 1 Or BaseChapter > GetChapterCount() Then Return False
            If EndChapter <> 0 AndAlso (EndChapter < BaseChapter Or EndChapter < 1 Or EndChapter > GetChapterCount()) Then Return False
            If BaseVerse <> 0 AndAlso (BaseVerse < 1 Or BaseVerse > GetVerseCount(BaseChapter)) Then Return False
            If ExtraVerseNumber <> 0 AndAlso ((BaseChapter = If(EndChapter = 0, BaseChapter, EndChapter) And BaseVerse <> 0 And ExtraVerseNumber < BaseVerse) Or ExtraVerseNumber < 1 Or ExtraVerseNumber > GetVerseCount(If(EndChapter = 0, BaseChapter, EndChapter))) Then Return False
            Dim Check As Collections.Generic.List(Of String()) = QuranTextRangeLookup(BaseChapter, BaseVerse, 0, EndChapter, ExtraVerseNumber, 0)
            If WordNumber < 1 Or WordNumber > Array.FindAll(Check(0)(0).ToCharArray(), Function(Ch As Char) Ch = " "c).Length + 1 Then Return False
            If EndWordNumber <> 0 AndAlso (BaseChapter = If(EndChapter = 0, BaseChapter, EndChapter) And BaseVerse = If(ExtraVerseNumber = 0, BaseVerse, ExtraVerseNumber) And WordNumber <> 0 And EndWordNumber < WordNumber Or EndWordNumber < 1 Or EndWordNumber > Array.FindAll(Check(Check.Count - 1)(Check(Check.Count - 1).Length - 1).ToCharArray(), Function(Ch As Char) Ch = " "c).Length + 1) Then Return False
        Next
        Return True
    End Function
    Public Shared Function QuranTextFromReference(Str As String, SchemeType As ArabicData.TranslitScheme, Scheme As String, TranslationIndex As Integer, W4W As Boolean, W4WNum As Boolean, NoArabic As Boolean, Header As Boolean, NoRef As Boolean, Colorize As Boolean, Verses As Boolean) As RenderArray
        Dim Renderer As New RenderArray(String.Empty)
        Dim Matches As System.Text.RegularExpressions.MatchCollection = System.Text.RegularExpressions.Regex.Matches(Str, "(?:,?(\d+)(?:\:(\d+))?(?:\:(\d+))?(?:-(\d+)(?:\:(\d+))?(?:\:(\d+))?)?)")
        Dim Reference As String
        For Count = 0 To Matches.Count - 1
            Dim BaseChapter As Integer = CInt(Matches(Count).Groups(1).Value)
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
            If BaseVerse = 0 Then
                BaseVerse += 1
                ExtraVerseNumber = GetTextChapter(CachedData.XMLDocMain, BaseChapter).ChildNodes.Count
            End If
            If WordNumber = 0 Then WordNumber += 1
            Renderer.Items.AddRange(DoGetRenderedQuranText(QuranTextRangeLookup(BaseChapter, BaseVerse, WordNumber, EndChapter, ExtraVerseNumber, EndWordNumber), BaseChapter, BaseVerse, CachedData.IslamData.Translations.TranslationList(TranslationIndex).Name, SchemeType, Scheme, TranslationIndex, W4W, W4WNum, NoArabic, Header, NoRef, Colorize, Verses).Items)
            Reference = CStr(BaseChapter) + If(BaseVerse <> 0, ":" + CStr(BaseVerse), String.Empty) + If(EndChapter <> 0, "-" + CStr(EndChapter) + If(ExtraVerseNumber <> 0, ":" + CStr(ExtraVerseNumber), String.Empty), If(ExtraVerseNumber <> 0, "-" + CStr(ExtraVerseNumber), String.Empty))
            Renderer.Items.Add(New RenderArray.RenderItem(RenderArray.RenderTypes.eHeaderCenter, New RenderArray.RenderText() {New RenderArray.RenderText(RenderArray.RenderDisplayClass.eLTR, "(Qur'an " + Reference + ")")}))
        Next
        Return Renderer
    End Function
    'much faster to make a word index...
    Public Shared Function QuranTextFromSearch(Str As String, SchemeType As ArabicData.TranslitScheme, Scheme As String, TranslationIndex As Integer, W4W As Boolean, W4WNum As Boolean, NoArabic As Boolean, Header As Boolean, NoRef As Boolean, Colorize As Boolean, UseVerses As Boolean) As RenderArray
        Dim Renderer As New RenderArray(String.Empty)
        Dim Verses As List(Of String()) = GetQuranText(CachedData.XMLDocMain, -1, -1, -1, -1)
        Dim RefList As String = String.Empty
        Dim RefCount As Integer = 0
        For Count As Integer = 0 To Verses.Count - 1
            For SubCount As Integer = 0 To Verses(Count).Length - 1
                Dim Matches As System.Text.RegularExpressions.MatchCollection = System.Text.RegularExpressions.Regex.Matches(Arabic.TransliterateToScheme(Verses(Count)(SubCount), ArabicData.TranslitScheme.Literal, String.Empty), Str)
                For MatchCount As Integer = 0 To Matches.Count - 1
                    Renderer.Items.AddRange(DoGetRenderedQuranText(QuranTextRangeLookup(Count + 1, SubCount + 1, Array.FindAll(Verses(Count)(SubCount).Substring(0, Matches(MatchCount).Index).ToCharArray(), Function(Ch As Char) Ch = " "c).Length + 1, Count + 1, SubCount + 1, Array.FindAll(Verses(Count)(SubCount).Substring(0, Matches(MatchCount).Index + Matches(MatchCount).Length).ToCharArray(), Function(Ch As Char) Ch = " "c).Length + 1), Count + 1, SubCount + 1, CachedData.IslamData.Translations.TranslationList(TranslationIndex).Name, SchemeType, Scheme, TranslationIndex, W4W, W4WNum, NoArabic, Header, NoRef, Colorize, UseVerses).Items)
                    Dim Reference As String = CStr(Count + 1) + ":" + CStr(SubCount + 1)
                    RefList += If(RefList <> String.Empty, ",", String.Empty) + Reference
                    RefCount += 1
                    If Array.FindAll(Verses(Count)(SubCount).Substring(0, Matches(MatchCount).Index).ToCharArray(), Function(Ch As Char) Ch = " "c).Length <> 0 OrElse Array.FindAll(Verses(Count)(SubCount).Substring(Matches(MatchCount).Index + Matches(MatchCount).Length).ToCharArray(), Function(Ch As Char) Ch = " "c).Length <> 0 Then
                        RefList += ":" + CStr(Array.FindAll(Verses(Count)(SubCount).Substring(0, Matches(MatchCount).Index).ToCharArray(), Function(Ch As Char) Ch = " "c).Length + 1)
                        If Array.FindAll(Verses(Count)(SubCount).Substring(Matches(MatchCount).Index + Matches(MatchCount).Length).ToCharArray(), Function(Ch As Char) Ch = " "c).Length <> 0 Then RefList += "-" + CStr(Array.FindAll(Verses(Count)(SubCount).Substring(0, Matches(MatchCount).Index + Matches(MatchCount).Length).ToCharArray(), Function(Ch As Char) Ch = " "c).Length + 1)
                    End If
                    Renderer.Items.Add(New RenderArray.RenderItem(RenderArray.RenderTypes.eHeaderCenter, New RenderArray.RenderText() {New RenderArray.RenderText(RenderArray.RenderDisplayClass.eLTR, "(" + Arabic.TransliterateToScheme(Arabic.GetCatNoun("QuranReadingRecitation")(0).Text, SchemeType, Scheme) + " " + Reference + ")")}))
                Next
            Next
        Next
        Renderer.Items.Add(New RenderArray.RenderItem(RenderArray.RenderTypes.eHeaderCenter, New RenderArray.RenderText() {New RenderArray.RenderText(RenderArray.RenderDisplayClass.eLTR, "(" + Arabic.TransliterateToScheme(Arabic.GetCatNoun("QuranReadingRecitation")(0).Text, SchemeType, Scheme) + " " + RefList + ") " + CStr(RefCount) + " Total")}))
        Return Renderer
    End Function
    Public Shared Function TextPositionToMorphology(Text As String, WordPos As Integer) As String
        Dim Chapter As Integer = System.Text.RegularExpressions.Regex.Matches(Text.Substring(0, WordPos), ArabicData.ArabicEndOfAyah + Arabic.TransliterateFromBuckwalter("1") + "\s").Count
        Dim Verse As Integer = Integer.Parse(Arabic.TransliterateToScheme(System.Text.RegularExpressions.Regex.Match(Text.Substring(WordPos), ArabicData.ArabicEndOfAyah + "(\d{1,3})").Groups(1).Value, ArabicData.TranslitScheme.Literal, String.Empty))
        If Verse = 1 Then Chapter += 1
        Dim Word As Integer = System.Text.RegularExpressions.Regex.Matches(Text.Substring(0, WordPos).Substring(Text.Substring(0, WordPos).LastIndexOf(ArabicData.ArabicEndOfAyah) + 1), "(\s.)?\s").Count
        Dim Lines As String() = IO.File.ReadAllLines(Utility.GetFilePath("metadata\quranic-corpus-morphology-0.4.txt"))
        TextPositionToMorphology = String.Empty
        For Count As Integer = 0 To Lines.Length - 1
            If Lines(Count).Length <> 0 AndAlso Lines(Count).Chars(0) <> "#" Then
                Dim Pieces As String() = Lines(Count).Split(CChar(vbTab))
                If Pieces(0).Chars(0) = "(" Then
                    Dim Location As Integer() = Array.ConvertAll(Pieces(0).TrimStart("("c).TrimEnd(")"c).Split(":"c), Function(Str As String) CInt(Str))
                    If Location(0) = Chapter And Location(1) = Verse And Location(2) = Word Then TextPositionToMorphology += If(TextPositionToMorphology = String.Empty, String.Empty, vbCrLf) + Lines(Count)
                End If
            End If
        Next
    End Function
    Public Shared Sub CheckSequentialRules()
        Dim Rules As IslamData.RuleTranslationCategory.RuleTranslation() = CachedData.GetRuleSet("SimpleScriptHamzaWriting") '"HamzaWriting")
        Dim IndexToVerse As Integer()() = Nothing
        Dim XMLDocAlt As New System.Xml.XmlDocument
        XMLDocAlt.Load(QuranTextNames(QuranTexts.Hafs) + "-" + QuranFileNames(QuranScripts.SimpleEnhanced) + ".xml")
        Dim Text As String = QuranTextCombiner(XMLDocAlt, IndexToVerse) 'CachedData.XMLDocMain
        Dim Matches As System.Text.RegularExpressions.MatchCollection = System.Text.RegularExpressions.Regex.Matches(Text, CachedData.GetPattern("Hamzas"))
        Dim CheckMatches As New Dictionary(Of Integer, String)
        Debug.Print(CStr(Matches.Count))
        For Count = 0 To Matches.Count - 1
            If Matches(Count).Length = 0 Then Continue For 'Avoid zero width matches for end of string anchors
            If Not CheckMatches.ContainsKey(Matches(Count).Index) Then CheckMatches.Add(Matches(Count).Index, String.Empty)
            CheckMatches(Matches(Count).Index) += "0-"
        Next
        For MainCount = 0 To Rules.Length - 1
            Matches = System.Text.RegularExpressions.Regex.Matches(Text, Rules(MainCount).Match)
            Dim NegativeCount As Integer = 0
            For Count = 0 To Matches.Count - 1
                If Rules(MainCount).NegativeMatch <> String.Empty AndAlso Matches(Count).Result(Rules(MainCount).NegativeMatch) <> String.Empty Then
                    NegativeCount += 1
                ElseIf Matches(Count).Result(Rules(MainCount).Evaluator) <> Matches(Count).Value Then
                    Debug.Print(CStr(Matches(Count).Groups(2 + If(Rules(MainCount).NegativeMatch <> String.Empty, 1, 0)).Index) + ":" + If(CheckMatches(Matches(Count).Groups(2 + If(Rules(MainCount).NegativeMatch <> String.Empty, 1, 0)).Index).EndsWith("-"), String.Empty, "-") + Rules(MainCount).Name + If(CheckMatches(Matches(Count).Groups(2 + If(Rules(MainCount).NegativeMatch <> String.Empty, 1, 0)).Index).Length = 2, String.Empty, "-") + ":" + Arabic.TransliterateToScheme(Text(Matches(Count).Groups(2 + If(Rules(MainCount).NegativeMatch <> String.Empty, 1, 0)).Index), ArabicData.TranslitScheme.Literal, String.Empty) + ":" + Arabic.TransliterateToScheme(Text.Substring(Math.Max(0, Matches(Count).Groups(2 + If(Rules(MainCount).NegativeMatch <> String.Empty, 1, 0)).Index - 15), 30), ArabicData.TranslitScheme.Literal, String.Empty))
                Else
                    If Not CheckMatches.ContainsKey(Matches(Count).Groups(2 + If(Rules(MainCount).NegativeMatch <> String.Empty, 1, 0)).Index) Then CheckMatches.Add(Matches(Count).Groups(2 + If(Rules(MainCount).NegativeMatch <> String.Empty, 1, 0)).Index, String.Empty)
                    CheckMatches(Matches(Count).Groups(2 + If(Rules(MainCount).NegativeMatch <> String.Empty, 1, 0)).Index) += If(CheckMatches(Matches(Count).Groups(2 + If(Rules(MainCount).NegativeMatch <> String.Empty, 1, 0)).Index).EndsWith("-"), String.Empty, "-") + Rules(MainCount).Name + If(CheckMatches(Matches(Count).Groups(2 + If(Rules(MainCount).NegativeMatch <> String.Empty, 1, 0)).Index).Length = 2, String.Empty, "-")
                End If
            Next
            Debug.Print(Rules(MainCount).Name + ": " + CStr(Matches.Count - NegativeCount))
        Next
        Dim Keys(CheckMatches.Keys.Count - 1) As Integer
        CheckMatches.Keys.CopyTo(Keys, 0)
        Array.Sort(Keys)
        For Count = 0 To Keys.Length - 1
            If CheckMatches(Keys(Count)).EndsWith("-") And CheckMatches(Keys(Count)) <> "0-LetterHamzaEnd-YehHamzaKasra-" And CheckMatches(Keys(Count)) <> "0-FathaAlefHamzaAboveSukun-TatweelHamzaSukun-" And CheckMatches(Keys(Count)) <> "0-LetterHamzaEnd-FathaAlefHamzaAboveEnd-" Then
                Debug.Print(CStr(Keys(Count)) + ":" + CheckMatches(Keys(Count)) + ":" + Arabic.TransliterateToScheme(Text(Keys(Count)), ArabicData.TranslitScheme.Literal, String.Empty) + ":" + Arabic.TransliterateToScheme(Text.Substring(Math.Max(0, Keys(Count) - 15), 30), ArabicData.TranslitScheme.Literal, String.Empty))
            End If
        Next
    End Sub
    Public Shared Sub CheckMutualExclusiveRules(bAssumeContinue As Boolean)
        'Dim Verify As String() = {CStr(ArabicData.ArabicLetterHamza), ArabicData.ArabicTatweel + "?" + ArabicData.ArabicHamzaAbove, ArabicData.ArabicLetterAlefWithHamzaAbove, ArabicData.ArabicLetterAlefWithHamzaBelow, ArabicData.ArabicLetterWawWithHamzaAbove, ArabicData.ArabicLetterYehWithHamzaAbove}
        Dim VerIndex As Integer = 16
        Dim IndexToVerse As Integer()() = Nothing
        Dim Text As String = QuranTextCombiner(CachedData.XMLDocMain, IndexToVerse)
        Dim Matches As System.Text.RegularExpressions.MatchCollection = System.Text.RegularExpressions.Regex.Matches(Text, CachedData.TranslateRegEx(CachedData.IslamData.VerificationSet(VerIndex).Match, True))
        Dim CheckMatches As New Dictionary(Of Integer, String)
        Debug.Print(CStr(Matches.Count))
        For Count = 0 To Matches.Count - 1
            If Matches(Count).Length = 0 Then Continue For 'Avoid zero width matches for end of string anchors
            For LenCount = 0 To Matches(Count).Length - 1
                If Not CheckMatches.ContainsKey(Matches(Count).Index + LenCount) Then CheckMatches.Add(Matches(Count).Index + LenCount, String.Empty)
                CheckMatches(Matches(Count).Index + LenCount) += "0"
            Next
        Next
        Dim MatchMetadata As String() = CachedData.IslamData.VerificationSet(VerIndex).Evaluator
        For MainCount = 0 To CachedData.IslamData.VerificationSet(VerIndex).MetaRules.Length - 1
            Dim MetaCount As Integer
            For MetaCount = 0 To CachedData.RulesOfRecitationRegEx.Length - 1
                If CachedData.IslamData.VerificationSet(VerIndex).MetaRules(MainCount) = CachedData.RulesOfRecitationRegEx(MetaCount).Name Then
                    Exit For
                End If
            Next
            Matches = System.Text.RegularExpressions.Regex.Matches(Text, CachedData.RulesOfRecitationRegEx(MetaCount).Match)
            Dim SieveCount As Integer = 0
            For Count = 0 To Matches.Count - 1
                Dim MetaRules As String() = CachedData.RulesOfRecitationRegEx(MetaCount).Evaluator
                If Count = 0 AndAlso Matches(Count).Groups.Count <> MetaRules.Length + 1 Then Debug.Print("Discrepency in metadata:" + CStr(MainCount + 1) + ":" + CStr(MetaRules.Length) + ":Got:" + CStr(Matches(Count).Groups.Count - 1))
                Dim bSieve As Boolean = False
                For SubCount = 0 To MetaRules.Length - 1
                    If Array.IndexOf(MetaRules(SubCount).Split("|"c), If(bAssumeContinue, "optionalstop", "optionalnotstop")) <> -1 And Matches(Count).Groups(SubCount + 1).Success Then
                        'check continuity in prior patterns
                        Dim CheckLen As Integer = 0
                        For CheckCount = 1 To SubCount
                            CheckLen += Matches(Count).Groups(CheckCount).Length
                        Next
                        If Matches(Count).Index + CheckLen <> Matches(Count).Groups(SubCount + 1).Index Then Debug.Print("Non-Sequential Capture:" + CStr(MainCount + 1) + ":" + Arabic.TransliterateToScheme(Text.Substring(Math.Max(0, Matches(Count).Groups(SubCount + 1).Index - 15), 30), ArabicData.TranslitScheme.Literal, String.Empty))
                        bSieve = True
                        Exit For
                    End If
                Next
                If Not bSieve Then
                    For SubCount = 0 To MetaRules.Length - 1
                        If Array.IndexOf(MatchMetadata, MetaRules(SubCount).Split("|"c)(0)) <> -1 And Matches(Count).Groups(SubCount + 1).Success Then
                            'If Verify.Length <> 0 And Not System.Text.RegularExpressions.Regex.Match(Matches(Count).Groups(SubCount + 1).Value, Verify(MainCount - 1)).Success Then
                            '    Debug.Print("Erroneous Match: " + Check(MainCount, 0) + " " + Arabic.TransliterateToScheme(Text.Substring(Math.Max(0, Matches(Count).Groups(SubCount + 1).Index - 15), 30), ArabicData.TranslitScheme.Literal, String.Empty))
                            '    'Debug.Print(TextPositionToMorphology(Text, Matches(Count).Groups(SubCount + 1).Index))
                            'End If
                            'check continuity in previous patterns
                            Dim CheckLen As Integer = 0
                            For CheckCount = 1 To SubCount
                                CheckLen += Matches(Count).Groups(CheckCount).Length
                            Next
                            If Matches(Count).Index + CheckLen <> Matches(Count).Groups(SubCount + 1).Index Then Debug.Print("Non-Sequential Capture:" + CStr(MainCount + 1) + ":" + Arabic.TransliterateToScheme(Text.Substring(Math.Max(0, Matches(Count).Groups(SubCount + 1).Index - 15), 30), ArabicData.TranslitScheme.Literal, String.Empty))
                            For LenCount = 0 To Matches(Count).Groups(SubCount + 1).Length - 1
                                If Not CheckMatches.ContainsKey(Matches(Count).Groups(SubCount + 1).Index + LenCount) Then CheckMatches.Add(Matches(Count).Groups(SubCount + 1).Index + LenCount, String.Empty)
                                CheckMatches(Matches(Count).Groups(SubCount + 1).Index + LenCount) += Hex(MainCount + 1)
                            Next
                            SieveCount += 1
                        End If
                    Next
                End If
            Next
            Debug.Print(CStr(SieveCount))
        Next
        Dim Keys(CheckMatches.Keys.Count - 1) As Integer
        CheckMatches.Keys.CopyTo(Keys, 0)
        Array.Sort(Keys)
        For Count = 0 To Keys.Length - 1
            If CheckMatches(Keys(Count)).Length <> 2 Then
                Debug.Print(CStr(Keys(Count)) + ":" + CheckMatches(Keys(Count)) + ":" + Arabic.TransliterateToScheme(Text(Keys(Count)), ArabicData.TranslitScheme.Literal, String.Empty) + ":" + Arabic.TransliterateToScheme(Text.Substring(Math.Max(0, Keys(Count) - 15), 30), ArabicData.TranslitScheme.Literal, String.Empty))
            End If
        Next
    End Sub
    Public Shared Function QuranTextCombiner(XMLDoc As System.Xml.XmlDocument, ByRef IndexToVerse As Integer()()) As String
        Dim Verses As List(Of String()) = GetQuranText(XMLDoc, -1, -1, -1, -1)
        Dim IndexToVerseList As New List(Of Integer())
        Dim Str As New System.Text.StringBuilder
        For Count As Integer = 0 To Verses.Count - 1
            For SubCount As Integer = 0 To Verses(Count).Length - 1
                Dim Words As String()
                Dim Index As Integer
                If SubCount = 0 Then
                    Dim Node As System.Xml.XmlNode
                    Node = GetTextVerse(GetTextChapter(XMLDoc, Count + 1), 1).Attributes.GetNamedItem("bismillah")
                    If Not Node Is Nothing Then
                        Words = Node.Value.Split(" "c)
                        Index = Str.Length
                        For WordCount = 0 To Words.Length - 1
                            IndexToVerseList.Add(New Integer() {Count + 1, SubCount, WordCount + 1, Index, Words(WordCount).Length})
                            Index += Words(WordCount).Length + 1
                        Next
                        Str.Append(Node.Value + " ")
                    End If
                End If
                Words = Verses(Count)(SubCount).Split(" "c)
                Index = Str.Length
                For WordCount = 0 To Words.Length - 1
                    IndexToVerseList.Add(New Integer() {Count + 1, SubCount + 1, WordCount + 1, Index, Words(WordCount).Length})
                    Index += Words(WordCount).Length + 1
                Next
                Str.Append(Verses(Count)(SubCount) + Arabic.TransliterateFromBuckwalter(" =" + CStr(SubCount + 1)) + " ")
            Next
        Next
        IndexToVerse = IndexToVerseList.ToArray()
        Return Str.ToString()
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
            QuranText(0)(0) = System.Text.RegularExpressions.Regex.Replace(System.Text.RegularExpressions.Regex.Replace(QuranText(0)(0).Substring(0, VerseIndex), "(^\s*|\s+)[^\s" + String.Join(String.Empty, Array.ConvertAll(CachedData.ArabicStopLetters, Function(Str As String) ArabicData.MakeUniRegEx(Str))) + ArabicData.ArabicStartOfRubElHizb + ArabicData.ArabicPlaceOfSajdah + "]+(?=\s*$|\s+)", "$1"), String.Join("|", Array.ConvertAll(CachedData.ArabicStopLetters, Function(Str As String) ArabicData.MakeUniRegEx(Str))) + "|" + ArabicData.ArabicStartOfRubElHizb + "|" + ArabicData.ArabicPlaceOfSajdah, ChrW(0)) + QuranText(0)(0).Substring(VerseIndex)
        End If
        If (EndWordNumber <> 0) Then
            Dim VerseIndex As Integer = 0
            'selections are always within the same chapter
            Dim LastChapter As Integer = QuranText.Count - 1
            Dim LastVerse As Integer = CInt(IIf(ExtraVerseNumber <> 0, QuranText(LastChapter).Length - 1, 0))
            While QuranText(LastChapter)(LastVerse)(VerseIndex) = ChrW(0) Or QuranText(LastChapter)(LastVerse)(VerseIndex) = " "
                VerseIndex += 1
            End While
            For WordCount As Integer = WordNumber - 1 To EndWordNumber - 1
                VerseIndex = QuranText(LastChapter)(LastVerse).IndexOf(" "c, VerseIndex) + 1
            Next
            If VerseIndex = 0 Then VerseIndex = QuranText(LastChapter)(LastVerse).Length
            QuranText(LastChapter)(LastVerse) = QuranText(LastChapter)(LastVerse).Substring(0, VerseIndex) + System.Text.RegularExpressions.Regex.Replace(System.Text.RegularExpressions.Regex.Replace(QuranText(LastChapter)(LastVerse).Substring(VerseIndex), "(^\s*|\s+)[^\s" + String.Join(String.Empty, Array.ConvertAll(CachedData.ArabicStopLetters, Function(Str As String) ArabicData.MakeUniRegEx(Str))) + ArabicData.ArabicStartOfRubElHizb + ArabicData.ArabicPlaceOfSajdah + "]+(?=\s*$|\s+)", "$1"), String.Join("|", Array.ConvertAll(CachedData.ArabicStopLetters, Function(Str As String) ArabicData.MakeUniRegEx(Str))) + "|" + ArabicData.ArabicStartOfRubElHizb + "|" + ArabicData.ArabicPlaceOfSajdah, ChrW(0))
        End If
        Return QuranText
    End Function
    Public Shared Function GetQuranTextBySelection(ID As String, Division As Integer, Index As Integer, Translation As String, SchemeType As ArabicData.TranslitScheme, Scheme As String, TranslationIndex As Integer, W4W As Boolean, W4WNum As Boolean, NoArabic As Boolean, Header As Boolean, NoRef As Boolean, Colorize As Boolean, Verses As Boolean) As RenderArray
        Dim Chapter As Integer
        Dim Verse As Integer
        Dim BaseChapter As Integer
        Dim BaseVerse As Integer
        Dim Node As System.Xml.XmlNode
        Dim Renderer As New RenderArray(ID)
        Dim QuranText As Collections.Generic.List(Of String())
        Dim SeperateSectionCount As Integer = 1
        Dim Keys() As String = Nothing
        If Division = 8 Then SeperateSectionCount = CachedData.IslamData.QuranSelections(Index).SelectionInfo.Length
        If Division = 9 Then
            ReDim Keys(CachedData.LetterDictionary(ArabicData.ArabicLetters(Index).Symbol).Count - 1)
            CachedData.LetterDictionary(ArabicData.ArabicLetters(Index).Symbol).Keys.CopyTo(Keys, 0)
            SeperateSectionCount = CachedData.LetterDictionary(ArabicData.ArabicLetters(Index).Symbol).Count
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
                For SubCount = 0 To CachedData.LetterDictionary(ArabicData.ArabicLetters(Index).Symbol)(Keys(SectionCount)).Count - 1
                    BaseChapter = CType(CachedData.LetterDictionary(ArabicData.ArabicLetters(Index).Symbol)(Keys(SectionCount))(SubCount), Integer())(0)
                    BaseVerse = CType(CachedData.LetterDictionary(ArabicData.ArabicLetters(Index).Symbol)(Keys(SectionCount))(SubCount), Integer())(1)
                    QuranText.Add(GetQuranText(CachedData.XMLDocMain, BaseChapter, BaseVerse, BaseVerse))
                    If SubCount <> CachedData.LetterDictionary(ArabicData.ArabicLetters(Index).Symbol)(Keys(SectionCount)).Count - 1 Then
                        Renderer.Items.AddRange(DoGetRenderedQuranText(QuranText, BaseChapter, BaseVerse, Translation, SchemeType, Scheme, TranslationIndex, W4W, W4WNum, NoArabic, Header, NoRef, Colorize, Verses).Items)
                    End If
                Next
            ElseIf Division = 10 Then
                QuranText = New Collections.Generic.List(Of String())
                Dim IndexToVerse As Integer()() = Nothing
                Dim Text As String = QuranTextCombiner(CachedData.XMLDocMain, IndexToVerse)
                Dim Matches As System.Text.RegularExpressions.MatchCollection = System.Text.RegularExpressions.Regex.Matches(Text, GetMetaRuleSet("UthmaniQuran").Rules(Index).Match)
                Renderer.Items.AddRange(New RenderArray.RenderItem() {New RenderArray.RenderItem(RenderArray.RenderTypes.eText, New RenderArray.RenderText() {New RenderArray.RenderText(RenderArray.RenderDisplayClass.eLTR, GetMetaRuleSet("UthmaniQuran").Rules(Index).Name)}), New RenderArray.RenderItem(RenderArray.RenderTypes.eText, DocBuilder.ColorizeRegExGroups(DocBuilder.GetRegExText(GetMetaRuleSet("UthmaniQuran").Rules(Index).Match), False)), New RenderArray.RenderItem(RenderArray.RenderTypes.eText, DocBuilder.ColorizeList(Array.ConvertAll(GetMetaRuleSet("UthmaniQuran").Rules(Index).Evaluator, Function(Str As String) DocBuilder.GetRegExText(Str)), False))})
                For SubCount = 0 To Matches.Count - 1
                    Dim StartWordIndex As Integer = Array.BinarySearch(IndexToVerse, Matches(SubCount).Index, New QuranWordIndexComparer)
                    If StartWordIndex < 0 Then StartWordIndex = (StartWordIndex Xor -1) - 1
                    Dim EndWordIndex As Integer = Array.BinarySearch(IndexToVerse, Matches(SubCount).Index + Matches(SubCount).Length - 1, New QuranWordIndexComparer)
                    If EndWordIndex < 0 Then EndWordIndex = (EndWordIndex Xor -1) + 1
                    Dim Renderers As New List(Of RenderArray.RenderText)
                    If IndexToVerse(StartWordIndex)(3) <> Matches(SubCount).Index Then
                        Renderers.Add(New RenderArray.RenderText(RenderArray.RenderDisplayClass.eArabic, Text.Substring(IndexToVerse(StartWordIndex)(3), Matches(SubCount).Index - IndexToVerse(StartWordIndex)(3))))
                    End If
                    Dim Groups(Matches(SubCount).Groups.Count - 1 - 1) As String
                    For GroupCount As Integer = 1 To Matches(SubCount).Groups.Count - 1
                        Groups(GroupCount - 1) = Matches(SubCount).Groups(GroupCount).Value
                    Next
                    Renderers.AddRange(DocBuilder.ColorizeList(Groups, True))
                    If IndexToVerse(EndWordIndex)(3) + IndexToVerse(EndWordIndex)(4) <> Matches(SubCount).Index + Matches(SubCount).Length Then
                        Renderers.Add(New RenderArray.RenderText(RenderArray.RenderDisplayClass.eArabic, Text.Substring(Matches(SubCount).Index + Matches(SubCount).Length, IndexToVerse(EndWordIndex)(3) + IndexToVerse(EndWordIndex)(4) - (Matches(SubCount).Index + Matches(SubCount).Length))))
                    End If
                    Renderer.Items.Add(New RenderArray.RenderItem(RenderArray.RenderTypes.eText, Renderers.ToArray()))
                    Renderer.Items.Add(New RenderArray.RenderItem(RenderArray.RenderTypes.eHeaderCenter, New RenderArray.RenderText() {New RenderArray.RenderText(RenderArray.RenderDisplayClass.eLTR, "(Qur'an " + CStr(IndexToVerse(StartWordIndex)(0)) + ":" + CStr(IndexToVerse(StartWordIndex)(1)) + ":" + CStr(IndexToVerse(StartWordIndex)(2)) + If(StartWordIndex <> EndWordIndex, "-" + CStr(IndexToVerse(EndWordIndex)(0)) + ":" + CStr(IndexToVerse(EndWordIndex)(1)) + ":" + CStr(IndexToVerse(EndWordIndex)(2)), String.Empty) + ")")}))
                Next
            Else
                QuranText = Nothing
            End If
            Renderer.Items.AddRange(DoGetRenderedQuranText(QuranText, BaseChapter, BaseVerse, Translation, SchemeType, Scheme, TranslationIndex, W4W, W4WNum, NoArabic, Header, NoRef, Colorize, Verses).Items)
        Next
        Return Renderer
    End Function
    Public Class QuranWordIndexComparer
        Implements IComparer(Of Object)
        Public Function Compare(x As Object, y As Object) As Integer Implements IComparer(Of Object).Compare
            If TypeOf x Is Integer() Then
                If CInt(y) >= CType(x, Integer())(3) And CInt(y) < (CType(x, Integer())(3) + CType(x, Integer())(4)) Then Return 0
                Return If(CInt(y) < CType(x, Integer())(3), 1, -1)
            Else
                If CInt(x) >= CType(y, Integer())(3) And CInt(x) < (CType(y, Integer())(3) + CType(y, Integer())(4)) Then Return 0
                Return If(CInt(x) < CType(y, Integer())(3), 1, -1)
            End If
        End Function
    End Class
    Public Shared Function GetRenderedQuranText(ByVal Item As PageLoader.TextItem) As RenderArray
        Dim Division As Integer = 0
        Dim Index As Integer = 1
        Dim Strings As String = HttpContext.Current.Request.QueryString.Get("qurandivision")
        If Not Strings Is Nothing Then Division = CInt(Strings)
        Strings = HttpContext.Current.Request.QueryString.Get("quranselection")
        If Not Strings Is Nothing Then Index = CInt(Strings)
        Dim SchemeType As ArabicData.TranslitScheme = CType(If(CInt(HttpContext.Current.Request.QueryString.Get("translitscheme")) >= 2, 2 - CInt(HttpContext.Current.Request.QueryString.Get("translitscheme")) Mod 2, CInt(HttpContext.Current.Request.QueryString.Get("translitscheme"))), ArabicData.TranslitScheme)
        Dim Scheme As String = If(CInt(HttpContext.Current.Request.QueryString.Get("translitscheme")) >= 2, CachedData.IslamData.TranslitSchemes(CInt(HttpContext.Current.Request.QueryString.Get("translitscheme")) \ 2).Name, String.Empty)
        Dim TranslationIndex As Integer = GetTranslationIndex(HttpContext.Current.Request.QueryString.Get("qurantranslation"))
        Return GetQuranTextBySelection(Item.Name, Division, Index, CachedData.IslamData.Translations.TranslationList(TranslationIndex).FileName, SchemeType, Scheme, TranslationIndex, CInt(HttpContext.Current.Request.QueryString.Get("wordversemode")) <> 4, CInt(HttpContext.Current.Request.QueryString.Get("wordversemode")) Mod 2 = 1, False, True, False, CInt(HttpContext.Current.Request.QueryString.Get("colorcuemode")) = 0, CInt(HttpContext.Current.Request.QueryString.Get("wordversemode")) / 2 <> 1)
    End Function
    Public Shared Function GenerateDefaultStops(Str As String) As Integer()
        Dim Matches As System.Text.RegularExpressions.MatchCollection = System.Text.RegularExpressions.Regex.Matches(Str, "(^\s*|\s+)(" + ArabicData.MakeUniRegEx(ArabicData.ArabicEndOfAyah) + "[\p{Nd}]{1,3}|" + CachedData.OptionalPattern + ")(?=\s*$|\s+)")
        Dim DefStops As New List(Of Integer)
        Dim DotToggle As Boolean = False
        For Count As Integer = 0 To Matches.Count - 1
            If Matches(Count).Groups(2).Value <> ArabicData.ArabicSmallHighLigatureSadWithLamWithAlefMaksura AndAlso (Matches(Count).Groups(2).Value <> ArabicData.ArabicSmallHighThreeDots Or DotToggle) Then DefStops.Add(Matches(Count).Groups(2).Index)
            If Matches(Count).Groups(2).Value = ArabicData.ArabicSmallHighThreeDots Then DotToggle = Not DotToggle
        Next
        Return DefStops.ToArray()
    End Function
    Public Shared Function DoGetRenderedQuranText(QuranText As Collections.Generic.List(Of String()), BaseChapter As Integer, BaseVerse As Integer, Translation As String, SchemeType As ArabicData.TranslitScheme, Scheme As String, TranslationIndex As Integer, W4W As Boolean, W4WNum As Boolean, NoArabic As Boolean, Header As Boolean, NoRef As Boolean, Colorize As Boolean, Verses As Boolean) As RenderArray
        Dim Text As String
        Dim Node As System.Xml.XmlNode
        Dim Renderer As New RenderArray(String.Empty)
        Dim Lines As String() = IO.File.ReadAllLines(Utility.GetFilePath("metadata\" + GetTranslationFileName(Translation)))
        Dim W4WLines As String() = If(W4W, IO.File.ReadAllLines(Utility.GetFilePath("metadata\en.w4w.shehnazshaikh.txt")), Nothing)
        If Not QuranText Is Nothing Then
            For Chapter = 0 To QuranText.Count - 1
                Dim ChapterNode As System.Xml.XmlNode = GetChapterByIndex(BaseChapter + Chapter)
                Dim Texts As New List(Of RenderArray.RenderText)
                If Header Then
                    If Not NoArabic Then
                        Texts.Add(New RenderArray.RenderText(RenderArray.RenderDisplayClass.eArabic, Arabic.TransliterateFromBuckwalter(CachedData.QuranHeaders(0) + " " + ChapterNode.Attributes.GetNamedItem("ayas").Value + " ")))
                        Texts.Add(New RenderArray.RenderText(RenderArray.RenderDisplayClass.eTransliteration, If(SchemeType <> ArabicData.TranslitScheme.None, Arabic.TransliterateToScheme(Arabic.TransliterateFromBuckwalter(CachedData.QuranHeaders(0) + " " + ChapterNode.Attributes.GetNamedItem("ayas").Value + " "), SchemeType, Scheme).Trim(), String.Empty)))
                    End If
                    If Translation <> String.Empty Then Texts.Add(New RenderArray.RenderText(RenderArray.RenderDisplayClass.eLTR, "Verses " + ChapterNode.Attributes.GetNamedItem("ayas").Value + " "))
                    Renderer.Items.Add(New RenderArray.RenderItem(RenderArray.RenderTypes.eHeaderLeft, Texts.ToArray()))
                    Texts.Clear()
                    If Not NoArabic Then
                        Texts.Add(New RenderArray.RenderText(RenderArray.RenderDisplayClass.eArabic, Arabic.TransliterateFromBuckwalter(CachedData.QuranHeaders(1) + " " + CachedData.IslamData.QuranChapters(CInt(ChapterNode.Attributes.GetNamedItem("index").Value) - 1).Name + " ")))
                        Texts.Add(New RenderArray.RenderText(RenderArray.RenderDisplayClass.eTransliteration, If(SchemeType <> ArabicData.TranslitScheme.None, Arabic.TransliterateToScheme(Arabic.TransliterateFromBuckwalter(CachedData.QuranHeaders(1) + " " + CachedData.IslamData.QuranChapters(CInt(ChapterNode.Attributes.GetNamedItem("index").Value) - 1).Name + " "), SchemeType, Scheme).Trim(), String.Empty)))
                    End If
                    If Translation <> String.Empty Then Texts.Add(New RenderArray.RenderText(RenderArray.RenderDisplayClass.eLTR, "Chapter " + TanzilReader.GetChapterEName(ChapterNode) + " "))
                    Renderer.Items.Add(New RenderArray.RenderItem(RenderArray.RenderTypes.eHeaderCenter, Texts.ToArray()))
                    Texts.Clear()
                    If Not NoArabic Then
                        Texts.Add(New RenderArray.RenderText(RenderArray.RenderDisplayClass.eArabic, Arabic.TransliterateFromBuckwalter(CachedData.QuranHeaders(2) + " " + ChapterNode.Attributes.GetNamedItem("rukus").Value + " ")))
                        Texts.Add(New RenderArray.RenderText(RenderArray.RenderDisplayClass.eTransliteration, If(SchemeType <> ArabicData.TranslitScheme.None, Arabic.TransliterateToScheme(Arabic.TransliterateFromBuckwalter(CachedData.QuranHeaders(2) + " " + ChapterNode.Attributes.GetNamedItem("rukus").Value + " "), SchemeType, Scheme).Trim(), String.Empty)))
                    End If
                    If Translation <> String.Empty Then Texts.Add(New RenderArray.RenderText(RenderArray.RenderDisplayClass.eLTR, "Rukus " + ChapterNode.Attributes.GetNamedItem("rukus").Value + " "))
                    Renderer.Items.Add(New RenderArray.RenderItem(RenderArray.RenderTypes.eHeaderRight, Texts.ToArray()))
                    Texts.Clear()
                End If
                For Verse = 0 To QuranText(Chapter).Length - 1
                    Dim Items As New Collections.Generic.List(Of RenderArray.RenderItem)
                    Text = String.Empty
                    'hizb symbols not needed as Quranic text already contains them
                    'If BaseChapter + Chapter <> 1 AndAlso TanzilReader.IsQuarterStart(BaseChapter + Chapter, CInt(IIf(Chapter = 0, BaseVerse, 1)) + Verse) Then
                    '    Text += Arabic.TransliterateFromBuckwalter("B")
                    '    Items.Add(New RenderArray.RenderItem(RenderArray.RenderTypes.eText, New RenderArray.RenderText() {New RenderArray.RenderText(RenderArray.RenderDisplayClass.eArabic, Arabic.TransliterateFromBuckwalter("B"))}))
                    'End If
                    If CInt(IIf(Chapter = 0, BaseVerse, 1)) + Verse = 1 Then
                        Node = GetTextVerse(GetTextChapter(CachedData.XMLDocMain, BaseChapter + Chapter), 1).Attributes.GetNamedItem("bismillah")
                        If Not Node Is Nothing Then
                            If Not NoArabic Then
                                Texts.Add(New RenderArray.RenderText(RenderArray.RenderDisplayClass.eArabic, Node.Value + " "))
                                Texts.Add(New RenderArray.RenderText(RenderArray.RenderDisplayClass.eTransliteration, If(SchemeType <> ArabicData.TranslitScheme.None, Arabic.TransliterateToScheme(Node.Value, SchemeType, Scheme).Trim(), String.Empty)))
                            End If
                            If Translation <> String.Empty Then Texts.Add(New RenderArray.RenderText(DirectCast(IIf(IsTranslationTextLTR(TranslationIndex), RenderArray.RenderDisplayClass.eLTR, RenderArray.RenderDisplayClass.eRTL), RenderArray.RenderDisplayClass), TanzilReader.GetTranslationVerse(Lines, 1, 1)))
                            Renderer.Items.Add(New RenderArray.RenderItem(RenderArray.RenderTypes.eText, Texts.ToArray()))
                            Texts.Clear()
                        End If
                    End If
                    Text += QuranText(Chapter)(Verse).Trim(" "c, ChrW(0)) + If(NoRef, String.Empty, " ")
                    If TanzilReader.IsSajda(BaseChapter + Chapter, CInt(IIf(Chapter = 0, BaseVerse, 1)) + Verse) Then
                        'Sajda markers are already in the text
                        'Text += Arabic.TransliterateFromBuckwalter("R")
                        'Items.Add(New RenderArray.RenderItem(RenderArray.RenderTypes.eText, New RenderArray.RenderText() {New RenderArray.RenderText(RenderArray.RenderDisplayClass.eArabic, Arabic.TransliterateFromBuckwalter("R"))}))
                    End If
                    If NoRef Then Text = Arabic.TransliterateFromBuckwalter("(") + Text
                    Text += Arabic.TransliterateFromBuckwalter(If(NoRef, ")", "=" + CStr(CInt(IIf(Chapter = 0, BaseVerse, 1)) + Verse))) + " "
                    If W4W And Translation <> String.Empty Then
                        Dim DefStops As Integer() = GenerateDefaultStops(QuranText(Chapter)(Verse))
                        Dim Words As String() = If(QuranText(Chapter)(Verse) Is Nothing, {}, QuranText(Chapter)(Verse).Split(" "c))
                        Dim TranslitWords As String() = Arabic.TransliterateToScheme(QuranText(Chapter)(Verse), SchemeType, Scheme, DefStops).Split(" "c)
                        Dim PauseMarks As Integer = 0
                        If NoRef Then
                            If Not NoArabic Then
                                Texts.Add(New RenderArray.RenderText(RenderArray.RenderDisplayClass.eArabic, Arabic.TransliterateFromBuckwalter("(")))
                                Texts.Add(New RenderArray.RenderText(RenderArray.RenderDisplayClass.eTransliteration, If(SchemeType <> ArabicData.TranslitScheme.None, Arabic.TransliterateToScheme(Arabic.TransliterateFromBuckwalter(")"), SchemeType, Scheme), String.Empty)))
                            End If
                            If Translation <> String.Empty Then Texts.Add(New RenderArray.RenderText(DirectCast(IIf(IsTranslationTextLTR(TranslationIndex), RenderArray.RenderDisplayClass.eLTR, RenderArray.RenderDisplayClass.eRTL), RenderArray.RenderDisplayClass), ")"))
                            Items.Add(New RenderArray.RenderItem(RenderArray.RenderTypes.eText, Texts.ToArray()))
                            Texts.Clear()
                        End If
                        Dim WordColors As RenderArray.RenderText()() = Nothing
                        If Colorize Then WordColors = Arabic.ApplyColorRules(Text, GenerateDefaultStops(Text), True)
                        Dim Pos As Integer = 0
                        For Count As Integer = 0 To Words.Length - 1
                            'handle start/end words here which have space placeholders
                            If Words(Count).Length = 1 AndAlso _
                                Words(Count)(0) = ChrW(0) Then
                                PauseMarks += 1
                            ElseIf Words(Count).Length = 1 AndAlso _
                                (Arabic.IsStop(ArabicData.FindLetterBySymbol(Words(Count)(0))) Or Words(Count)(0) = ArabicData.ArabicStartOfRubElHizb Or Words(Count)(0) = ArabicData.ArabicPlaceOfSajdah) Then
                                PauseMarks += 1
                                If Not NoArabic Then
                                    Texts.Add(New RenderArray.RenderText(RenderArray.RenderDisplayClass.eArabic, " " + Words(Count)))
                                    Texts.Add(New RenderArray.RenderText(RenderArray.RenderDisplayClass.eTransliteration, If(SchemeType <> ArabicData.TranslitScheme.None, TranslitWords(Count), String.Empty)))
                                    If Count <> Words.Length - 1 AndAlso Words(Count + 1).Length <> 1 Then Texts.Add(New RenderArray.RenderText(RenderArray.RenderDisplayClass.eContinueStop, Array.IndexOf(DefStops, Pos) = -1))
                                End If
                                If W4WNum Then Texts.Add(New RenderArray.RenderText(RenderArray.RenderDisplayClass.eLTR, CStr(Count + 1)))
                                Items.Add(New RenderArray.RenderItem(RenderArray.RenderTypes.eText, Texts.ToArray()))
                                Texts.Clear()
                            ElseIf Words(Count).Length <> 0 Then
                                If Not NoArabic Then
                                    If Colorize Then
                                        Texts.AddRange(WordColors(Count))
                                    Else
                                        Texts.Add(New RenderArray.RenderText(RenderArray.RenderDisplayClass.eArabic, Words(Count)))
                                    End If
                                    Texts.Add(New RenderArray.RenderText(RenderArray.RenderDisplayClass.eTransliteration, If(SchemeType <> ArabicData.TranslitScheme.None, TranslitWords(Count), String.Empty)))
                                End If
                                If Translation <> String.Empty Then Texts.Add(New RenderArray.RenderText(DirectCast(IIf(IsTranslationTextLTR(TranslationIndex), RenderArray.RenderDisplayClass.eLTR, RenderArray.RenderDisplayClass.eRTL), RenderArray.RenderDisplayClass), TanzilReader.GetW4WTranslationVerse(W4WLines, BaseChapter + Chapter, CInt(IIf(Chapter = 0, BaseVerse, 1)) + Verse, Count - PauseMarks)))
                                If W4WNum Then Texts.Add(New RenderArray.RenderText(RenderArray.RenderDisplayClass.eLTR, CStr(Count + 1)))
                                Items.Add(New RenderArray.RenderItem(RenderArray.RenderTypes.eText, Texts.ToArray()))
                                If Not NoArabic AndAlso Count <> Words.Length - 1 AndAlso Words(Count + 1).Length <> 1 Then Items.Add(New RenderArray.RenderItem(RenderArray.RenderTypes.eText, {New RenderArray.RenderText(RenderArray.RenderDisplayClass.eContinueStop, True)}))
                                Texts.Clear()
                            End If
                            Pos += Words(Count).Length + 1
                        Next
                        If Not NoArabic Then
                            Texts.Add(New RenderArray.RenderText(RenderArray.RenderDisplayClass.eArabic, Arabic.TransliterateFromBuckwalter(If(NoRef, ")", "=" + CStr(CInt(IIf(Chapter = 0, BaseVerse, 1)) + Verse)))))
                            Texts.Add(New RenderArray.RenderText(RenderArray.RenderDisplayClass.eTransliteration, If(SchemeType <> ArabicData.TranslitScheme.None, Arabic.TransliterateToScheme(Arabic.TransliterateFromBuckwalter(If(NoRef, "(", "=" + CStr(CInt(IIf(Chapter = 0, BaseVerse, 1)) + Verse))), SchemeType, Scheme), String.Empty)))
                            Texts.Add(New RenderArray.RenderText(RenderArray.RenderDisplayClass.eContinueStop, False))
                        End If
                        If Translation <> String.Empty Then Texts.Add(New RenderArray.RenderText(DirectCast(IIf(IsTranslationTextLTR(TranslationIndex), RenderArray.RenderDisplayClass.eLTR, RenderArray.RenderDisplayClass.eRTL), RenderArray.RenderDisplayClass), "(" + If(NoRef, String.Empty, CStr(CInt(IIf(Chapter = 0, BaseVerse, 1)) + Verse) + ")")))
                        Items.Add(New RenderArray.RenderItem(RenderArray.RenderTypes.eText, Texts.ToArray()))
                        Texts.Clear()
                        'Text += Arabic.TransliterateFromBuckwalter("(" + CStr(IIf(Chapter = 0, BaseVerse, 1) + Verse) + ") ")
                        Texts.Add(New RenderArray.RenderText(RenderArray.RenderDisplayClass.eNested, Items))
                    End If
                    If Verses Then
                        If Not NoArabic Then
                            If Colorize Then
                                Texts.AddRange(Arabic.ApplyColorRules(Text, GenerateDefaultStops(Text), False)(0))
                            Else
                                Texts.Add(New RenderArray.RenderText(RenderArray.RenderDisplayClass.eArabic, Text))
                            End If
                            Texts.Add(New RenderArray.RenderText(RenderArray.RenderDisplayClass.eTransliteration, If(SchemeType <> ArabicData.TranslitScheme.None, Arabic.TransliterateToScheme(If(NoRef, Arabic.TransliterateFromBuckwalter("("), String.Empty) + QuranText(Chapter)(Verse).Trim(" "c, ChrW(0)) + Arabic.TransliterateFromBuckwalter(If(NoRef, ")", " " + "=" + CStr(CInt(IIf(Chapter = 0, BaseVerse, 1)) + Verse) + " ")), SchemeType, Scheme, GenerateDefaultStops(If(NoRef, Arabic.TransliterateFromBuckwalter("("), String.Empty) + QuranText(Chapter)(Verse).Trim(" "c, ChrW(0)) + Arabic.TransliterateFromBuckwalter(If(NoRef, ")", " " + "=" + CStr(CInt(IIf(Chapter = 0, BaseVerse, 1)) + Verse)) + " "))).Trim(), String.Empty)))
                        End If
                        If Translation <> String.Empty Then Texts.Add(New RenderArray.RenderText(DirectCast(IIf(IsTranslationTextLTR(TranslationIndex), RenderArray.RenderDisplayClass.eLTR, RenderArray.RenderDisplayClass.eRTL), RenderArray.RenderDisplayClass), If(NoRef, String.Empty, "(" + CStr(CInt(IIf(Chapter = 0, BaseVerse, 1)) + Verse) + ") ") + TanzilReader.GetTranslationVerse(Lines, BaseChapter + Chapter, CInt(IIf(Chapter = 0, BaseVerse, 1)) + Verse)))
                    End If
                    Renderer.Items.Add(New RenderArray.RenderItem(RenderArray.RenderTypes.eText, Texts.ToArray()))
                    If Not NoArabic Then Renderer.Items.Add(New RenderArray.RenderItem(RenderArray.RenderTypes.eText, {New RenderArray.RenderText(RenderArray.RenderDisplayClass.eContinueStop, False)}))
                    Texts.Clear()
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
        If EndVerse = -1 Then EndVerse = GetTextChapter(XMLDocMain, Chapter).ChildNodes.Count 'GetVerseCount(Chapter)
        Dim Verses(EndVerse - StartVerse) As String
        For Count = StartVerse To EndVerse
            Dim VerseNode As Xml.XmlNode = GetTextVerse(GetTextChapter(XMLDocMain, Chapter), Count)
            If Not VerseNode Is Nothing Then
                Dim AttrNode As Xml.XmlNode = VerseNode.Attributes.GetNamedItem("text")
                If Not AttrNode Is Nothing Then
                    Verses(Count - StartVerse) = AttrNode.Value
                End If
            End If
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
        Dim SchemeType As ArabicData.TranslitScheme = CType(If(CInt(HttpContext.Current.Request.QueryString.Get("translitscheme")) >= 2, 2 - CInt(HttpContext.Current.Request.QueryString.Get("translitscheme")) Mod 2, CInt(HttpContext.Current.Request.QueryString.Get("translitscheme"))), ArabicData.TranslitScheme)
        Dim Scheme As String = If(CInt(HttpContext.Current.Request.QueryString.Get("translitscheme")) >= 2, CachedData.IslamData.TranslitSchemes(CInt(HttpContext.Current.Request.QueryString.Get("translitscheme")) \ 2).Name, String.Empty)
        Dim Names() As Array = Array.ConvertAll(Utility.GetChildNodes("sura", Utility.GetChildNode("suras", CachedData.XMLDocInfo.DocumentElement.ChildNodes).ChildNodes), Function(Convert As System.Xml.XmlNode) New Object() {Convert.Attributes.GetNamedItem("index").Value + ". " + GetChapterEName(Convert) + " (" + ArabicData.RightToLeftEmbedding + Arabic.TransliterateFromBuckwalter(CachedData.QuranHeaders(1) + " " + CachedData.IslamData.QuranChapters(CInt(Convert.Attributes.GetNamedItem("index").Value) - 1).Name) + ArabicData.PopDirectionalFormatting + ")" + If(SchemeType = ArabicData.TranslitScheme.None, String.Empty, " " + Arabic.TransliterateToScheme(Arabic.TransliterateFromBuckwalter(CachedData.IslamData.QuranChapters(CInt(Convert.Attributes.GetNamedItem("index").Value) - 1).Name), SchemeType, Scheme)), CInt(Convert.Attributes.GetNamedItem("index").Value)})
        Array.Sort(Names, New Utility.CompareNameValueArray)
        Return Names
    End Function
    Public Shared Function GetChapterEName(ByVal ChapterNode As System.Xml.XmlNode) As String
        Return Utility.LoadResourceString("IslamInfo_QuranChapter" + ChapterNode.Attributes.GetNamedItem("index").Value)
    End Function
    Public Shared Function GetChapterNamesByRevelationOrder() As Array()
        Dim SchemeType As ArabicData.TranslitScheme = CType(If(CInt(HttpContext.Current.Request.QueryString.Get("translitscheme")) >= 2, 2 - CInt(HttpContext.Current.Request.QueryString.Get("translitscheme")) Mod 2, CInt(HttpContext.Current.Request.QueryString.Get("translitscheme"))), ArabicData.TranslitScheme)
        Dim Scheme As String = If(CInt(HttpContext.Current.Request.QueryString.Get("translitscheme")) >= 2, CachedData.IslamData.TranslitSchemes(CInt(HttpContext.Current.Request.QueryString.Get("translitscheme")) \ 2).Name, String.Empty)
        Dim Names() As Array = Array.ConvertAll(Utility.GetChildNodes("sura", Utility.GetChildNode("suras", CachedData.XMLDocInfo.DocumentElement.ChildNodes).ChildNodes), Function(Convert As System.Xml.XmlNode) New Object() {Convert.Attributes.GetNamedItem("index").Value + ". " + GetChapterEName(Convert) + " (" + ArabicData.RightToLeftEmbedding + Arabic.TransliterateFromBuckwalter(CachedData.QuranHeaders(1) + " " + CachedData.IslamData.QuranChapters(CInt(Convert.Attributes.GetNamedItem("index").Value) - 1).Name) + ArabicData.PopDirectionalFormatting + ")" + If(SchemeType = ArabicData.TranslitScheme.None, String.Empty, " " + Arabic.TransliterateToScheme(Arabic.TransliterateFromBuckwalter(CachedData.IslamData.QuranChapters(CInt(Convert.Attributes.GetNamedItem("index").Value) - 1).Name), SchemeType, Scheme)), CInt(Convert.Attributes.GetNamedItem("order").Value)})
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
        Dim SchemeType As ArabicData.TranslitScheme = CType(If(CInt(HttpContext.Current.Request.QueryString.Get("translitscheme")) >= 2, 2 - CInt(HttpContext.Current.Request.QueryString.Get("translitscheme")) Mod 2, CInt(HttpContext.Current.Request.QueryString.Get("translitscheme"))), ArabicData.TranslitScheme)
        Dim Scheme As String = If(CInt(HttpContext.Current.Request.QueryString.Get("translitscheme")) >= 2, CachedData.IslamData.TranslitSchemes(CInt(HttpContext.Current.Request.QueryString.Get("translitscheme")) \ 2).Name, String.Empty)
        Dim Translation As String = HttpContext.Current.Request.QueryString.Get("hadithtranslation")
        Return DoGetRenderedText(Item.Name, SchemeType, Scheme, Translation)
    End Function
    Public Shared Function DoGetRenderedText(ID As String, SchemeType As ArabicData.TranslitScheme, Scheme As String, Translation As String) As RenderArray
        Dim Renderer As New RenderArray(ID)
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
            Renderer.Items.Add(New RenderArray.RenderItem(RenderArray.RenderTypes.eHeaderLeft, New RenderArray.RenderText() {New RenderArray.RenderText(RenderArray.RenderDisplayClass.eArabic, Arabic.TransliterateFromBuckwalter("Had~iv " + BookNode.Attributes.GetNamedItem("hadiths").Value + " ")), New RenderArray.RenderText(RenderArray.RenderDisplayClass.eTransliteration, Arabic.TransliterateToScheme(Arabic.TransliterateFromBuckwalter("Had~iv " + BookNode.Attributes.GetNamedItem("hadiths").Value + " "), SchemeType, Scheme).Trim()), New RenderArray.RenderText(RenderArray.RenderDisplayClass.eLTR, "Hadiths: " + BookNode.Attributes.GetNamedItem("hadiths").Value + " ")}))
            Renderer.Items.Add(New RenderArray.RenderItem(RenderArray.RenderTypes.eHeaderCenter, New RenderArray.RenderText() {New RenderArray.RenderText(RenderArray.RenderDisplayClass.eArabic, Arabic.TransliterateFromBuckwalter("{lokita`bu " + CStr(BookIndex)) + " " + BookNode.Attributes.GetNamedItem("name").Value + " "), New RenderArray.RenderText(RenderArray.RenderDisplayClass.eTransliteration, Arabic.TransliterateToScheme(Arabic.TransliterateFromBuckwalter("{lokita`bu " + CStr(BookIndex)) + " " + BookNode.Attributes.GetNamedItem("name").Value + " ", SchemeType, Scheme).Trim()), New RenderArray.RenderText(RenderArray.RenderDisplayClass.eLTR, "Book " + CStr(BookIndex) + ": " + GetBookEName(BookNode, Index) + " ")}))
            'Renderer.Items.Add(New RenderArray.RenderItem(RenderArray.RenderTypes.eHeaderRight, New RenderArray.RenderText() {New RenderArray.RenderText(RenderArray.RenderDisplayClass.eArabic, Arabic.TransliterateFromBuckwalter("mjld " + Utility.GetChildNode("books", XMLDocInfo(Index).DocumentElement.ChildNodes).ChildNodes.Item(BookIndex).Attributes.GetNamedItem("volume").Value + " ")), New RenderArray.RenderText(RenderArray.RenderDisplayClass.eTransliteration, Arabic.TransliterateToScheme(Arabic.TransliterateFromBuckwalter("mjld " + Utility.GetChildNode("books", XMLDocInfo(Index).DocumentElement.ChildNodes).ChildNodes.Item(BookIndex).Attributes.GetNamedItem("volume").Value + " "), SchemeType, Scheme).Trim()), New RenderArray.RenderText(RenderArray.RenderDisplayClass.eLTR, "Volume " + Utility.GetChildNode("books", XMLDocInfo(Index).DocumentElement.ChildNodes).ChildNodes.Item(BookIndex).Attributes.GetNamedItem("volume").Value + " ")}))
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
                        Renderer.Items.Add(New RenderArray.RenderItem(RenderArray.RenderTypes.eHeaderLeft, New RenderArray.RenderText() {New RenderArray.RenderText(RenderArray.RenderDisplayClass.eArabic, Arabic.TransliterateFromBuckwalter("Had~iv " + ChapterNode.Attributes.GetNamedItem("hadiths").Value + " ")), New RenderArray.RenderText(RenderArray.RenderDisplayClass.eTransliteration, Arabic.TransliterateToScheme(Arabic.TransliterateFromBuckwalter("Had~iv " + ChapterNode.Attributes.GetNamedItem("hadiths").Value + " "), SchemeType, Scheme).Trim()), New RenderArray.RenderText(RenderArray.RenderDisplayClass.eLTR, "Hadiths: " + ChapterNode.Attributes.GetNamedItem("hadiths").Value + " ")}))
                        Dim Heads As New List(Of RenderArray.RenderText)
                        Heads.AddRange(Array.ConvertAll(Of String, RenderArray.RenderText)(System.Text.RegularExpressions.Regex.Split(Arabic.TransliterateFromBuckwalter("bAb " + CStr(ChapterIndex)) + " " + ChapterNode.Attributes.GetNamedItem("name").Value + " ", "(\d+\.\d+(?:-\d+)?)"), Function(Str As String) If(System.Text.RegularExpressions.Regex.Match(Str, "(\d+)\.(\d+(?:-\d+)?)").Success, New RenderArray.RenderText(RenderArray.RenderDisplayClass.eLink, {"/host.aspx?Page=docbuild&docedit=%7B" + Str.Replace(".", "%3A") + "%7D&selectiondisplay=Display&translitscheme=0&fontselection=def&fontcustom=Lotus", Str}), New RenderArray.RenderText(RenderArray.RenderDisplayClass.eArabic, Str))))
                        Heads.Add(New RenderArray.RenderText(RenderArray.RenderDisplayClass.eTransliteration, Arabic.TransliterateToScheme(Arabic.TransliterateFromBuckwalter("bAb " + CStr(ChapterIndex)) + " " + System.Text.RegularExpressions.Regex.Replace(ChapterNode.Attributes.GetNamedItem("name").Value, "(\d+).(\d+(?:-\d+)?)", String.Empty) + " ", SchemeType, Scheme).Trim()))
                        Heads.AddRange(Array.ConvertAll(Of String, RenderArray.RenderText)(System.Text.RegularExpressions.Regex.Split("Chapter " + CStr(ChapterIndex) + ": " + Utility.DefaultValue(Utility.LoadResourceString("IslamInfo_" + CachedData.IslamData.Collections(Index).FileName + "Book" + BookNode.Attributes.GetNamedItem("index").Value + "Chapter" + ChapterNode.Attributes.GetNamedItem("index").Value), String.Empty) + " ", "(\d+\.\d+(?:-\d+)?)"), Function(Str As String) If(System.Text.RegularExpressions.Regex.Match(Str, "(\d+)\.(\d+(?:-\d+)?)").Success, New RenderArray.RenderText(RenderArray.RenderDisplayClass.eLink, {"/host.aspx?Page=docbuild&docedit=%7B" + Str.Replace(".", "%3A") + "%7D&selectiondisplay=Display&translitscheme=0&fontselection=def&fontcustom=Lotus", Str}), New RenderArray.RenderText(RenderArray.RenderDisplayClass.eLTR, Str))))
                        Renderer.Items.Add(New RenderArray.RenderItem(RenderArray.RenderTypes.eHeaderCenter, Heads.ToArray()))
                    End If
                    SubChapterIndex = -1
                End If
                'Handle missing or excess subchapter indexes
                If SubChapterIndex <> CInt(HadithText(Hadith)(2)) Then
                    SubChapterIndex = CInt(HadithText(Hadith)(2))
                    If Not ChapterNode Is Nothing Then
                        SubChapterNode = GetSubChapterByIndex(ChapterNode, SubChapterIndex)
                        If Not SubChapterNode Is Nothing Then
                            Renderer.Items.Add(New RenderArray.RenderItem(RenderArray.RenderTypes.eHeaderLeft, New RenderArray.RenderText() {New RenderArray.RenderText(RenderArray.RenderDisplayClass.eArabic, Arabic.TransliterateFromBuckwalter("Had~iv " + SubChapterNode.Attributes.GetNamedItem("hadiths").Value + " ")), New RenderArray.RenderText(RenderArray.RenderDisplayClass.eTransliteration, Arabic.TransliterateToScheme(Arabic.TransliterateFromBuckwalter("Had~iv " + SubChapterNode.Attributes.GetNamedItem("hadiths").Value + " "), SchemeType, Scheme).Trim()), New RenderArray.RenderText(RenderArray.RenderDisplayClass.eLTR, "Hadiths: " + SubChapterNode.Attributes.GetNamedItem("hadiths").Value + " ")}))
                            Dim Heads As New List(Of RenderArray.RenderText)
                            Heads.AddRange(Array.ConvertAll(Of String, RenderArray.RenderText)(System.Text.RegularExpressions.Regex.Split(Arabic.TransliterateFromBuckwalter("bAb " + CStr(SubChapterIndex)) + " " + SubChapterNode.Attributes.GetNamedItem("name").Value + " ", "(\d+\.\d+(?:-\d+)?)"), Function(Str As String) If(System.Text.RegularExpressions.Regex.Match(Str, "(\d+)\.(\d+(?:-\d+)?)").Success, New RenderArray.RenderText(RenderArray.RenderDisplayClass.eLink, {"/host.aspx?Page=docbuild&docedit=%7B" + Str.Replace(".", "%3A") + "%7D&selectiondisplay=Display&translitscheme=0&fontselection=def&fontcustom=Lotus", Str}), New RenderArray.RenderText(RenderArray.RenderDisplayClass.eArabic, Str))))
                            Heads.Add(New RenderArray.RenderText(RenderArray.RenderDisplayClass.eTransliteration, Arabic.TransliterateToScheme(Arabic.TransliterateFromBuckwalter("bAb " + CStr(SubChapterIndex)) + " " + System.Text.RegularExpressions.Regex.Replace(SubChapterNode.Attributes.GetNamedItem("name").Value, "(\d+).(\d+(?:-\d+)?)", String.Empty) + " ", SchemeType, Scheme).Trim()))
                            Heads.AddRange(Array.ConvertAll(Of String, RenderArray.RenderText)(System.Text.RegularExpressions.Regex.Split("Sub-Chapter " + CStr(SubChapterIndex) + ": " + Utility.DefaultValue(Utility.LoadResourceString("IslamInfo_" + CachedData.IslamData.Collections(Index).FileName + "Book" + BookNode.Attributes.GetNamedItem("index").Value + "Chapter" + ChapterNode.Attributes.GetNamedItem("index").Value + "Subchapter" + SubChapterNode.Attributes.GetNamedItem("index").Value), String.Empty) + " ", "(\d+\.\d+(?:-\d+)?)"), Function(Str As String) If(System.Text.RegularExpressions.Regex.Match(Str, "(\d+)\.(\d+(?:-\d+)?)").Success, New RenderArray.RenderText(RenderArray.RenderDisplayClass.eLink, {"/host.aspx?Page=docbuild&docedit=%7B" + Str.Replace(".", "%3A") + "%7D&selectiondisplay=Display&translitscheme=0&fontselection=def&fontcustom=Lotus", Str}), New RenderArray.RenderText(RenderArray.RenderDisplayClass.eLTR, Str))))
                            Renderer.Items.Add(New RenderArray.RenderItem(RenderArray.RenderTypes.eHeaderCenter, Heads.ToArray()))
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
                Dim Texts As New List(Of RenderArray.RenderText)
                Texts.AddRange(Array.ConvertAll(Of String, RenderArray.RenderText)(System.Text.RegularExpressions.Regex.Split(CStr(HadithText(Hadith)(3)) + " " + Arabic.TransliterateFromBuckwalter("=" + CStr(HadithText(Hadith)(0)) + " "), "(\d+\.\d+(?:-\d+)?)"), Function(Str As String) If(System.Text.RegularExpressions.Regex.Match(Str, "(\d+)\.(\d+(?:-\d+)?)").Success, New RenderArray.RenderText(RenderArray.RenderDisplayClass.eLink, {"/host.aspx?Page=docbuild&docedit=%7B" + Str.Replace(".", "%3A") + "%7D&selectiondisplay=Display&translitscheme=0&fontselection=def&fontcustom=Lotus", Str}), New RenderArray.RenderText(RenderArray.RenderDisplayClass.eArabic, Str))))
                Texts.Add(New RenderArray.RenderText(RenderArray.RenderDisplayClass.eTransliteration, Arabic.TransliterateToScheme(System.Text.RegularExpressions.Regex.Replace(CStr(HadithText(Hadith)(3)), "(\d+).(\d+(?:-\d+)?)", String.Empty) + " " + Arabic.TransliterateFromBuckwalter("=" + CStr(HadithText(Hadith)(0))) + " ", SchemeType, Scheme).Trim()))
                Texts.AddRange(Array.ConvertAll(Of String, RenderArray.RenderText)(System.Text.RegularExpressions.Regex.Split("(" + CStr(HadithText(Hadith)(0)) + ") " + HadithTranslation, "(\d+\.\d+(?:-\d+)?)"), Function(Str As String) If(System.Text.RegularExpressions.Regex.Match(Str, "(\d+)\.(\d+(?:-\d+)?)").Success, New RenderArray.RenderText(RenderArray.RenderDisplayClass.eLink, {"/host.aspx?Page=docbuild&docedit=%7B" + Str.Replace(".", "%3A") + "%7D&selectiondisplay=Display&translitscheme=0&fontselection=def&fontcustom=Lotus", Str}), New RenderArray.RenderText(DirectCast(IIf(IsTranslationTextLTR(Index, Translation), RenderArray.RenderDisplayClass.eLTR, RenderArray.RenderDisplayClass.eRTL), RenderArray.RenderDisplayClass), Str))))
                Renderer.Items.Add(New RenderArray.RenderItem(RenderArray.RenderTypes.eText, Texts.ToArray()))
                Dim Ranking As Integer() = IslamSiteDatabase.GetHadithRankingData(CachedData.IslamData.Collections(Index).FileName, BookIndex, CInt(HadithText(Hadith)(0)))
                Dim UserRanking As Integer
                If Utility.IsLoggedIn() Then
                    UserRanking = IslamSiteDatabase.GetUserHadithRankingData(Utility.GetUserID(), CachedData.IslamData.Collections(Index).FileName, BookIndex, CInt(HadithText(Hadith)(0)))
                Else
                    UserRanking = -1
                End If
                Renderer.Items.Add(New RenderArray.RenderItem(RenderArray.RenderTypes.eInteractive, New RenderArray.RenderText() {New RenderArray.RenderText(RenderArray.RenderDisplayClass.eRanking, CachedData.IslamData.Collections(Index).FileName + "|" + CStr(BookIndex) + "|" + CStr(HadithText(Hadith)(0)) + "|" + CStr(Ranking(0)) + "|" + CStr(Ranking(1)) + "|" + CStr(UserRanking))}))
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
Class IslamSiteDatabase
    Public Shared Sub CreateDatabase()
        Dim Connection As MySql.Data.MySqlClient.MySqlConnection = SiteDatabase.GetConnection()
        If Connection Is Nothing Then Return
        SiteDatabase.ExecuteNonQuery(Connection, "CREATE TABLE HadithRankings (UserID int NOT NULL, " + _
        "Collection VARCHAR(254) NOT NULL, " + _
        "BookIndex int, " + _
        "HadithIndex int NOT NULL, " + _
        "Ranking int NOT NULL)")
        Connection.Close()
    End Sub
    Public Shared Sub RemoveDatabase()
        Dim Connection As MySql.Data.MySqlClient.MySqlConnection = SiteDatabase.GetConnection()
        If Connection Is Nothing Then Return
        Dim Command As MySql.Data.MySqlClient.MySqlCommand = Connection.CreateCommand()
        SiteDatabase.ExecuteNonQuery(Connection, "DROP TABLE HadithRankings")
        Connection.Close()
    End Sub
    Public Shared Sub UpdateRankingData(ByVal UserID As Integer)
        If GetUserHadithRankingData(UserID, HttpContext.Current.Request.Form.Get("Collection"), CInt(HttpContext.Current.Request.Form.Get("Book")), CInt(HttpContext.Current.Request.Form.Get("Hadith"))) = -1 Then
            If CInt(HttpContext.Current.Request.Form.Get("Rating")) <> 0 Then
                SetUserHadithRankingData(UserID, HttpContext.Current.Request.Form.Get("Collection"), CInt(HttpContext.Current.Request.Form.Get("Book")), CInt(HttpContext.Current.Request.Form.Get("Hadith")), CInt(HttpContext.Current.Request.Form.Get("Rating")))
            End If
        Else
            If CInt(HttpContext.Current.Request.Form.Get("Rating")) <> 0 Then
                UpdateUserHadithRankingData(UserID, HttpContext.Current.Request.Form.Get("Collection"), CInt(HttpContext.Current.Request.Form.Get("Book")), CInt(HttpContext.Current.Request.Form.Get("Hadith")), CInt(HttpContext.Current.Request.Form.Get("Rating")))
            Else
                RemoveUserHadithRankingData(UserID, HttpContext.Current.Request.Form.Get("Collection"), CInt(HttpContext.Current.Request.Form.Get("Book")), CInt(HttpContext.Current.Request.Form.Get("Hadith")))
            End If
        End If
    End Sub
    Public Shared Sub WriteRankingData()
        Dim Data As Integer() = GetHadithRankingData(HttpContext.Current.Request.Form.Get("Collection"), CInt(HttpContext.Current.Request.Form.Get("Book")), CInt(HttpContext.Current.Request.Form.Get("Hadith")))
        If Data(1) <> 0 Then HttpContext.Current.Response.Write("Average of " + CStr(Data(0) / Data(1) / 2) + " out of " + CStr(Data(1)) + " rankings")
    End Sub
    Public Shared Function GetHadithCollectionRankingData(ByVal Collection As String) As Double
        Dim Connection As MySql.Data.MySqlClient.MySqlConnection = SiteDatabase.GetConnection()
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
        Dim Connection As MySql.Data.MySqlClient.MySqlConnection = SiteDatabase.GetConnection()
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
        Dim Connection As MySql.Data.MySqlClient.MySqlConnection = SiteDatabase.GetConnection()
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
        Dim Connection As MySql.Data.MySqlClient.MySqlConnection = SiteDatabase.GetConnection()
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
        Dim Connection As MySql.Data.MySqlClient.MySqlConnection = SiteDatabase.GetConnection()
        If Connection Is Nothing Then Return
        SiteDatabase.ExecuteNonQuery(Connection, "INSERT INTO HadithRankings (UserID, Collection, BookIndex, HadithIndex, Ranking) VALUES (" + CStr(UserID) + ", @Collection, " + CStr(Book) + ", " + CStr(Hadith) + ", " + CStr(Rank) + ")", New Generic.Dictionary(Of String, Object) From {{"@Collection", Collection}})
        Connection.Close()
    End Sub
    Public Shared Sub UpdateUserHadithRankingData(ByVal UserID As Integer, ByVal Collection As String, ByVal Book As Integer, ByVal Hadith As Integer, ByVal Rank As Integer)
        Dim Connection As MySql.Data.MySqlClient.MySqlConnection = SiteDatabase.GetConnection()
        If Connection Is Nothing Then Return
        SiteDatabase.ExecuteNonQuery(Connection, "UPDATE HadithRankings SET Ranking=" + CStr(Rank) + " WHERE UserID=" + CStr(UserID) + " AND Collection=@Collection AND BookIndex=" + CStr(Book) + " AND HadithIndex=" + CStr(Hadith), New Generic.Dictionary(Of String, Object) From {{"@Collection", Collection}})
        Connection.Close()
    End Sub
    Public Shared Sub RemoveUserHadithRankingData(ByVal UserID As Integer, ByVal Collection As String, ByVal Book As Integer, ByVal Hadith As Integer)
        Dim Connection As MySql.Data.MySqlClient.MySqlConnection = SiteDatabase.GetConnection()
        If Connection Is Nothing Then Return
        SiteDatabase.ExecuteNonQuery(Connection, "DELETE FROM HadithRankings WHERE UserID=" + CStr(UserID) + " AND Collection=@Collection AND BookIndex=" + CStr(Book) + " AND HadithIndex=" + CStr(Hadith), New Generic.Dictionary(Of String, Object) From {{"@Collection", Collection}})
        Connection.Close()
    End Sub
End Class