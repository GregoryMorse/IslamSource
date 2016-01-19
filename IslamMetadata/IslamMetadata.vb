Option Explicit On
Option Strict On
Imports XMLRender

Public Class PrayerTime
    Public Shared Function GetMonthName(Name As String) As String
        Dim CultureInfo As Globalization.CultureInfo
        If Name = "hijrimonthname" Then
            CultureInfo = New Globalization.CultureInfo("ar-SA")
            CultureInfo.DateTimeFormat.Calendar = New HijriCalendar
        ElseIf Name = "umalquramonthname" Then
            CultureInfo = New Globalization.CultureInfo("ar-SA")
            CultureInfo.DateTimeFormat.Calendar = New UmAlQuraCalendar
        Else
            CultureInfo = Globalization.CultureInfo.CurrentCulture 'Globalization.CultureInfo.CurrentCulture.LCID
        End If
        'If Array.Exists(Globalization.CultureInfo.CurrentCulture.OptionalCalendars, Function(Cal As Globalization.Calendar) Cal.ToString() = Calendar.ToString()) Then
        GetMonthName = CultureInfo.DateTimeFormat.MonthNames(CultureInfo.DateTimeFormat.Calendar.GetMonth(DateTime.Today) - 1)
    End Function
    Public Shared Function GetCalendar(Name As String) As Array()
        Dim Count As Integer
        Dim Calendar As Globalization.Calendar
        If Name = "hijricalendar" Then
            Calendar = New HijriCalendar
        ElseIf Name = "umalquracalendar" Then
            Calendar = New UmAlQuraCalendar
        Else
            Calendar = Globalization.CultureInfo.CurrentCulture.Calendar
        End If
        Dim TimeNow As DateTime = DateTime.Today
        Dim RetArray(Calendar.GetWeekOfYear(Calendar.ToDateTime(Calendar.GetYear(TimeNow), Calendar.GetMonth(TimeNow), Calendar.GetDaysInMonth(Calendar.GetYear(TimeNow), Calendar.GetMonth(TimeNow)), TimeNow.Hour, TimeNow.Minute, TimeNow.Second, TimeNow.Millisecond), Globalization.CalendarWeekRule.FirstDay, DayOfWeek.Sunday) - Calendar.GetWeekOfYear(Calendar.ToDateTime(Calendar.GetYear(TimeNow), Calendar.GetMonth(TimeNow), 1, TimeNow.Hour, TimeNow.Minute, TimeNow.Second, TimeNow.Millisecond), Globalization.CalendarWeekRule.FirstDay, DayOfWeek.Sunday) + 1 + 2) As Array
        RetArray(0) = New String() {}
        RetArray(1) = New String() {String.Empty, String.Empty, String.Empty, String.Empty, String.Empty, String.Empty, String.Empty}
        RetArray(2) = New String() {Globalization.CultureInfo.CurrentCulture.DateTimeFormat.DayNames(0), Globalization.CultureInfo.CurrentCulture.DateTimeFormat.DayNames(1), Globalization.CultureInfo.CurrentCulture.DateTimeFormat.DayNames(2), Globalization.CultureInfo.CurrentCulture.DateTimeFormat.DayNames(3), Globalization.CultureInfo.CurrentCulture.DateTimeFormat.DayNames(4), Globalization.CultureInfo.CurrentCulture.DateTimeFormat.DayNames(5), Globalization.CultureInfo.CurrentCulture.DateTimeFormat.DayNames(6)}
        For Count = 1 To Calendar.GetDaysInMonth(Calendar.GetYear(TimeNow), Calendar.GetMonth(TimeNow))
            RetArray(2 + 1 + Calendar.GetWeekOfYear(Calendar.ToDateTime(Calendar.GetYear(TimeNow), Calendar.GetMonth(TimeNow), Count, TimeNow.Hour, TimeNow.Minute, TimeNow.Second, TimeNow.Millisecond), Globalization.CalendarWeekRule.FirstDay, DayOfWeek.Sunday) - Calendar.GetWeekOfYear(Calendar.ToDateTime(Calendar.GetYear(TimeNow), Calendar.GetMonth(TimeNow), 1, TimeNow.Hour, TimeNow.Minute, TimeNow.Second, TimeNow.Millisecond), Globalization.CalendarWeekRule.FirstDay, DayOfWeek.Sunday)) = New String() {String.Empty, String.Empty, String.Empty, String.Empty, String.Empty, String.Empty, String.Empty}
            Do
                CType(RetArray(2 + 1 + Calendar.GetWeekOfYear(Calendar.ToDateTime(Calendar.GetYear(TimeNow), Calendar.GetMonth(TimeNow), Count, TimeNow.Hour, TimeNow.Minute, TimeNow.Second, TimeNow.Millisecond), Globalization.CalendarWeekRule.FirstDay, DayOfWeek.Sunday) - Calendar.GetWeekOfYear(Calendar.ToDateTime(Calendar.GetYear(TimeNow), Calendar.GetMonth(TimeNow), 1, TimeNow.Hour, TimeNow.Minute, TimeNow.Second, TimeNow.Millisecond), Globalization.CalendarWeekRule.FirstDay, DayOfWeek.Sunday)), String())(Calendar.GetDayOfWeek(Calendar.ToDateTime(Calendar.GetYear(TimeNow), Calendar.GetMonth(TimeNow), Count, TimeNow.Hour, TimeNow.Minute, TimeNow.Second, TimeNow.Millisecond))) = CStr(If(Calendar.GetDayOfMonth(TimeNow) = Count, ">", String.Empty)) + CStr(Calendar.GetDayOfMonth(Calendar.ToDateTime(Calendar.GetYear(TimeNow), Calendar.GetMonth(TimeNow), Count, TimeNow.Hour, TimeNow.Minute, TimeNow.Second, TimeNow.Millisecond))) + CStr(If(Calendar.GetDayOfMonth(TimeNow) = Count, "<", String.Empty))
                Count += 1
            Loop While Count <= Calendar.GetDaysInMonth(Calendar.GetYear(TimeNow), Calendar.GetMonth(TimeNow)) AndAlso _
                Calendar.GetDayOfWeek(Calendar.ToDateTime(Calendar.GetYear(TimeNow), Calendar.GetMonth(TimeNow), Count, TimeNow.Hour, TimeNow.Minute, TimeNow.Second, TimeNow.Millisecond)) <> DayOfWeek.Sunday
            Count -= 1
        Next
        Return RetArray
    End Function
    Public Shared Function GetPrayerTimes(Strings As String(), GeoData As String) As Array()
        Dim PrayTimes As New PrayTime.PrayTime
        Dim Count As Integer
        Dim TimeNow As DateTime = DateTime.Today
        If Strings.Length <> 11 OrElse Strings(0) = "ERROR" Then Return New Array() {}
        'Dim Times As String() = PrayTimes.getDatePrayerTimes(Today.Year, Today.Month, Today.Day, CDbl(Strings(8)), CDbl(Strings(9)), CInt(Strings(10).Split(":")(0)) + If(CInt(Strings(10).Split(":")(0)) >= 0, CInt(Strings(10).Split(":")(1)) / 60, -CInt(Strings(10).Split(":")(1)) / 60), 0)
        Dim RetArray(Date.DaysInMonth(TimeNow.Year, TimeNow.Month) + 2) As Array
        RetArray(0) = New String() {}
        RetArray(1) = New String() {String.Empty, String.Empty, String.Empty, String.Empty, String.Empty, String.Empty, String.Empty, String.Empty, String.Empty, String.Empty, String.Empty}
        RetArray(2) = New String() {Utility.LoadResourceString("IslamInfo_Date"), Utility.LoadResourceString("IslamInfo_Day"), Utility.LoadResourceString("IslamInfo_PrayTime6"), Utility.LoadResourceString("IslamInfo_PrayTime1"), Utility.LoadResourceString("IslamInfo_PrayTime7"), Utility.LoadResourceString("IslamInfo_PrayTime2"), Utility.LoadResourceString("IslamInfo_PrayTime3"), Utility.LoadResourceString("IslamInfo_PrayTime8"), Utility.LoadResourceString("IslamInfo_PrayTime4"), Utility.LoadResourceString("IslamInfo_PrayTime5"), Utility.LoadResourceString("IslamInfo_PrayTime9")}
        For Count = 1 To Date.DaysInMonth(TimeNow.Year, TimeNow.Month)
            Dim Times As String() = PrayTimes.getDatePrayerTimes(TimeNow.Year, TimeNow.Month, Count, CDbl(Strings(8)), CDbl(Strings(9)), CInt(Strings(10).Split(":"c)(0)) + CInt(If(CInt(Strings(10).Split(":"c)(0)) >= 0, CInt(Strings(10).Split(":"c)(1)) / 60, -CInt(Strings(10).Split(":"c)(1)) / 60)), CInt(GeoData))
            RetArray(Count + 2) = New String() {CStr(Count), New Date(TimeNow.Year, TimeNow.Month, Count).ToString("dddd", Globalization.CultureInfo.CurrentCulture), Times(0), Times(1), Times(2), Times(3), Times(4), Times(5), Times(6), Times(7), Times(8)}
        Next
        Return RetArray
    End Function
    Public Shared Function GetQiblaDirection(Strings As String()) As String
        Const QiblaLat As Double = 21.42252
        Const QiblaLon As Double = 39.82621
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
        'If (Math.Abs(dLon) > Math.PI) Then dLon = If(dLon > 0, -(2 * Math.PI - dLon), (2 * Math.PI + dLon))
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
        Return New List(Of Object())(Linq.Enumerable.Select(CachedData.RecitationSymbols, Function(Ch As String) New Object() {ArabicData.ArabicLetters(ArabicData.FindLetterBySymbol(Ch.Chars(0))).UnicodeName + " (" + ArabicData.FixStartingCombiningSymbol(Ch) + ArabicData.LeftToRightEmbedding + ")" + ArabicData.PopDirectionalFormatting, ArabicData.FindLetterBySymbol(Ch.Chars(0))})).ToArray()
    End Function
    Public Shared _BuckwalterMap As Dictionary(Of Char, Integer)
    Public Shared ReadOnly Property BuckwalterMap As Dictionary(Of Char, Integer)
        Get
            If _BuckwalterMap Is Nothing Then
                If Not DiskCache.GetCacheItem("BuckwalterMap", DateTime.MinValue) Is Nothing Then
                    _BuckwalterMap = CType((New Runtime.Serialization.DataContractSerializer(GetType(Dictionary(Of Char, Integer)))).ReadObject(New IO.MemoryStream(DiskCache.GetCacheItem("BuckwalterMap", DateTime.MinValue))), Dictionary(Of Char, Integer))
                Else
                    _BuckwalterMap = New Dictionary(Of Char, Integer)
                    For Index = 0 To ArabicData.ArabicLetters.Length - 1
                        If GetSchemeValueFromSymbol(ArabicData.ArabicLetters(Index), "ExtendedBuckwalter").Length <> 0 Then
                            _BuckwalterMap.Add(GetSchemeValueFromSymbol(ArabicData.ArabicLetters(Index), "ExtendedBuckwalter").Chars(0), Index)
                        End If
                    Next
                    Dim MemStream As New IO.MemoryStream
                    Dim Ser As New Runtime.Serialization.DataContractSerializer(GetType(Dictionary(Of Char, Integer)))
                    Ser.WriteObject(MemStream, _BuckwalterMap)
                    DiskCache.CacheItem("BuckwalterMap", DateTime.Now, MemStream.ToArray())
                    'MemStream.Close()
                    MemStream = Nothing
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
    Public Shared Function TransliterateToScheme(ByVal ArabicString As String, SchemeType As ArabicData.TranslitScheme, Scheme As String, RuleSet As IslamMetadata.IslamData.RuleMetaSet.RuleMetadataTranslation(), Optional OptionalStops() As Integer = Nothing, Optional PreString As String = "", Optional PostString As String = "", Optional PreStop As Boolean = True, Optional PostStop As Boolean = True) As String
        If SchemeType = ArabicData.TranslitScheme.LearningMode Then
            Return TransliterateWithRules(ArabicString, Scheme, OptionalStops, True, RuleSet)
        ElseIf SchemeType = ArabicData.TranslitScheme.RuleBased And PreString = "" And PostString = "" Then
            Return TransliterateWithRules(ArabicString, Scheme, OptionalStops, False, RuleSet)
        ElseIf SchemeType = ArabicData.TranslitScheme.RuleBased Then
            Return TransliterateContigWithRules(ArabicString, PreString, PostString, PreStop, PostStop, Scheme, OptionalStops, RuleSet)
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
        eSpellLongLetter
        eSpellLongMergedLetter
        eLookupLetter
        eLookupLongVowelDipthong
        eDivideTanween
        eLeadingGutteral
        eTrailingGutteral
        eResolveAmbiguity
        eLearningMode
        eObligatory
    End Enum
    Public Delegate Function RuleFunction(Str As String, Scheme As String, LearningMode As Boolean) As String()
    Public Shared RuleFunctions As RuleFunction() = {
        Function(Str As String, Scheme As String, LearningMode As Boolean) {Str.ToUpper()},
        Function(Str As String, Scheme As String, LearningMode As Boolean) {TransliterateWithRules(Arabic.TransliterateFromBuckwalter(Arabic.ArabicWordFromNumber(CInt(TransliterateToScheme(Str, ArabicData.TranslitScheme.Literal, String.Empty, Nothing)), True, False, False)), Scheme, Nothing, LearningMode, CachedData.RuleMetas("Normal"))},
        Function(Str As String, Scheme As String, LearningMode As Boolean) {TransliterateWithRules(ArabicLetterSpelling(Str, True, False, False), Scheme, Nothing, LearningMode, CachedData.RuleMetas("UthmaniQuran"))},
        Function(Str As String, Scheme As String, LearningMode As Boolean) {TransliterateWithRules(ArabicLetterSpelling(Str, True, True, False), Scheme, Nothing, LearningMode, CachedData.RuleMetas("UthmaniQuran"))},
        Function(Str As String, Scheme As String, LearningMode As Boolean) {TransliterateWithRules(ArabicLetterSpelling(Str, True, True, True), Scheme, Nothing, LearningMode, CachedData.RuleMetas("UthmaniQuran"))},
        Function(Str As String, Scheme As String, LearningMode As Boolean) {GetSchemeValueFromSymbol(ArabicData.ArabicLetters(ArabicData.FindLetterBySymbol(Str.Chars(0))), Scheme)},
        Function(Str As String, Scheme As String, LearningMode As Boolean) {GetSchemeLongVowelFromString(Str, Scheme)},
        Function(Str As String, Scheme As String, LearningMode As Boolean) {CachedData.ArabicFathaDammaKasra(Array.IndexOf(CachedData.ArabicTanweens, Str)), ArabicData.ArabicLetterNoon},
        Function(Str As String, Scheme As String, LearningMode As Boolean) {GetSchemeGutteralFromString(Str.Remove(Str.Length - 1), Scheme, True) + Str.Chars(Str.Length - 1)},
        Function(Str As String, Scheme As String, LearningMode As Boolean) {Str.Chars(0) + GetSchemeGutteralFromString(Str.Remove(0, 1), Scheme, False)},
        Function(Str As String, Scheme As String, LearningMode As Boolean) {If(SchemeHasValue(GetSchemeValueFromSymbol(ArabicData.ArabicLetters(ArabicData.FindLetterBySymbol(Str.Chars(0))), Scheme) + GetSchemeValueFromSymbol(ArabicData.ArabicLetters(ArabicData.FindLetterBySymbol(Str.Chars(1))), Scheme), Scheme), Str.Chars(0) + "-" + Str.Chars(1), Str)},
        Function(Str As String, Scheme As String, LearningMode As Boolean) If(LearningMode, {Str, String.Empty}, {String.Empty, Str}),
        Function(Str As String, Scheme As String, LearningMode As Boolean) {Str + "-" + Str + "(-" + Str(0) + ")"}
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
    Public Shared Function ArabicLetterSpelling(Input As String, Quranic As Boolean, IsLong As Boolean, Merged As Boolean) As String
        Dim Output As New System.Text.StringBuilder
        For Each Ch As Char In Input
            Dim Index As Integer = ArabicData.FindLetterBySymbol(Ch)
            If Index <> -1 AndAlso IsLetter(Index) Then
                If Output.Length <> 0 And Not Quranic Then Output.Append(" ")
                Dim Idx As Integer = Array.IndexOf(CachedData.ArabicLettersInOrder, CStr(ArabicData.ArabicLetters(Index).Symbol))
                Output.Append(If(Quranic, If(IsLong, If(Merged, CachedData.ArabicAlphabet(Idx).Remove(CachedData.ArabicAlphabet(Idx).Length - 1).Insert(CachedData.ArabicAlphabet(Idx).Length - 2, ArabicData.ArabicMaddahAbove).Insert(1, ArabicData.ArabicShadda), CachedData.ArabicAlphabet(Idx).Remove(CachedData.ArabicAlphabet(Idx).Length - 1).Insert(CachedData.ArabicAlphabet(Idx).Length - 2, ArabicData.ArabicMaddahAbove)), CachedData.ArabicAlphabet(Idx).Remove(CachedData.ArabicAlphabet(Idx).Length - 1)) + If(CachedData.ArabicAlphabet(Idx).EndsWith("n"), String.Empty, "o"), CachedData.ArabicAlphabet(Idx)))
            ElseIf Index <> -1 AndAlso ArabicData.ArabicLetters(Index).Symbol = ArabicData.ArabicMaddahAbove Then
                If Not Quranic Then Output.Append(Ch)
            End If
        Next
        Return Arabic.TransliterateFromBuckwalter(Output.ToString())
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
    Public Shared Function ApplyContigColorRules(ByVal ArabicString As String, PreString As String, PostString As String, PreStop As Boolean, PostStop As Boolean, OptionalStops As Integer(), BreakWords As Boolean, RuleSet As IslamMetadata.IslamData.RuleMetaSet.RuleMetadataTranslation()) As RenderArray.RenderText()()
        Dim RendererList As New List(Of RenderArray.RenderText())
        RendererList.AddRange(ApplyColorRules(JoinContig(ArabicString, PreString, PostString, PreStop, PostStop, OptionalStops), OptionalStops, True, RuleSet))
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
    Public Shared Function ApplyColorRules(ByVal ArabicString As String, OptionalStops As Integer(), BreakWords As Boolean, RuleSet As IslamMetadata.IslamData.RuleMetaSet.RuleMetadataTranslation()) As RenderArray.RenderText()()
        Dim Count As Integer
        Dim Index As Integer
        Dim MetadataList As New Generic.List(Of RuleMetadata)
        Dim Strings As New Generic.List(Of RenderArray.RenderText)
        For Count = 0 To RuleSet.Length - 1
            If Not RuleSet(Count).Evaluator Is Nothing Then
                Dim Matches As System.Text.RegularExpressions.MatchCollection = System.Text.RegularExpressions.Regex.Matches(ArabicString, RuleSet(Count).Match)
                For MatchIndex As Integer = 0 To Matches.Count - 1
                    Dim SubCount As Integer
                    For SubCount = 0 To RuleSet(Count).Evaluator.Length - 1
                        If (Array.IndexOf(RuleSet(Count).Evaluator(SubCount).Split("|"c), "optionalstop") <> -1 AndAlso (OptionalStops Is Nothing AndAlso Matches(MatchIndex).Groups(SubCount + 1).Value = ArabicData.ArabicSmallHighLigatureSadWithLamWithAlefMaksura OrElse (Not OptionalStops Is Nothing AndAlso Matches(MatchIndex).Groups(SubCount + 1).Value <> String.Empty AndAlso Array.IndexOf(OptionalStops, Matches(MatchIndex).Groups(SubCount + 1).Index) = -1))) OrElse (Array.IndexOf(RuleSet(Count).Evaluator(SubCount).Split("|"c), "optionalnotstop") <> -1 AndAlso (OptionalStops Is Nothing AndAlso Matches(MatchIndex).Groups(SubCount + 1).Value <> String.Empty AndAlso Matches(MatchIndex).Groups(SubCount + 1).Value <> ArabicData.ArabicSmallHighLigatureSadWithLamWithAlefMaksura OrElse (Matches(MatchIndex).Groups(SubCount + 1).Success AndAlso Matches(MatchIndex).Groups(SubCount + 1).Length = 0 AndAlso (Matches(MatchIndex).Groups(SubCount + 1).Index = 0 Or Matches(MatchIndex).Groups(SubCount + 1).Index = ArabicString.Length)) OrElse (Not OptionalStops Is Nothing AndAlso Matches(MatchIndex).Groups(SubCount + 1).Value <> String.Empty AndAlso Array.IndexOf(OptionalStops, Matches(MatchIndex).Groups(SubCount + 1).Index) <> -1))) Then Exit For
                    Next
                    If SubCount <> RuleSet(Count).Evaluator.Length Then Continue For
                    For SubCount = 0 To RuleSet(Count).Evaluator.Length - 1
                        If Not RuleSet(Count).Evaluator(SubCount) Is Nothing AndAlso RuleSet(Count).Evaluator(SubCount) <> String.Empty And (Matches(MatchIndex).Groups(SubCount + 1).Length <> 0 Or Array.IndexOf(AllowZeroLength, RuleSet(Count).Evaluator(SubCount)) <> -1) Then
                            MetadataList.Add(New RuleMetadata(Matches(MatchIndex).Groups(SubCount + 1).Index, Matches(MatchIndex).Groups(SubCount + 1).Length, RuleSet(Count).Evaluator(SubCount), SubCount))
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
                Dim Match As Integer = Array.FindIndex(CachedData.IslamData.ColorRules(Count).Match.Split("|"c), Function(Str As String) Array.IndexOf(New List(Of String)(Linq.Enumerable.Select(MetadataList(Index).Type.Split("|"c), Function(S As String) System.Text.RegularExpressions.Regex.Replace(S, "\(.*\)", String.Empty))).ToArray(), Str) <> -1)
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
                'Debug.Print(Rules(Replacements(Count - 1).OrigOrder).Name + ":" + Arabic.TransliterateToScheme(Replacements(Count - 1).Type, ArabicData.TranslitScheme.Literal, String.Empty, Nothing) + "-" + Rules(Replacements(Count).OrigOrder).Name + ":" + Arabic.TransliterateToScheme(Replacements(Count).Type, ArabicData.TranslitScheme.Literal, String.Empty, Nothing) + "-" + Arabic.TransliterateToScheme(ArabicString.Substring(Math.Max(Replacements(Count).Index - 15, 0), Math.Min(Replacements(Count - 1).Index + Replacements(Count - 1).Length + 15, ArabicString.Length) - Math.Max(Replacements(Count).Index - 15, 0)), ArabicData.TranslitScheme.Literal, String.Empty, Nothing))
            End If
            ArabicString = ArabicString.Substring(0, Replacements(Count).Index) + Replacements(Count).Type + ArabicString.Substring(Replacements(Count).Index + Replacements(Count).Length)
        Next
        Return ArabicString
    End Function
    Public Shared Function ChangeBaseScript(ArabicString As String, BaseText As TanzilReader.QuranTexts, ByVal PreString As String, ByVal PostString As String) As String
        If BaseText = TanzilReader.QuranTexts.Warsh Then
            ArabicString = UnjoinContig(ProcessTransform(JoinContig(ArabicString, PreString, PostString, False, False, Nothing), CachedData.RuleTranslations("WarshScript"), True), PreString, PostString)
        End If
        Return ArabicString
    End Function
    Public Shared Function ChangeScript(ArabicString As String, SrcScriptType As TanzilReader.QuranScripts, ScriptType As TanzilReader.QuranScripts, ByVal PreString As String, ByVal PostString As String) As String
        If SrcScriptType = TanzilReader.QuranScripts.Uthmani Then
            If ScriptType = TanzilReader.QuranScripts.UthmaniMin Then
                ArabicString = UnjoinContig(ProcessTransform(JoinContig(ArabicString, PreString, PostString, False, False, Nothing), CachedData.RuleTranslations("UthmaniMinimalScript"), False), PreString, PostString)
            ElseIf ScriptType = TanzilReader.QuranScripts.SimpleEnhanced Then
                Dim ScriptCombine As New List(Of IslamMetadata.IslamData.RuleTranslationCategory.RuleTranslation)
                ScriptCombine.AddRange(CachedData.RuleTranslations("SimpleScriptBase"))
                ScriptCombine.AddRange(CachedData.RuleTranslations("SimpleEnhancedScript"))
                ArabicString = UnjoinContig(ProcessTransform(JoinContig(ArabicString, PreString, PostString, False, False, Nothing), ScriptCombine.ToArray(), False), PreString, PostString)
            ElseIf ScriptType = TanzilReader.QuranScripts.Simple Then
                Dim ScriptCombine As New List(Of IslamMetadata.IslamData.RuleTranslationCategory.RuleTranslation)
                ScriptCombine.AddRange(CachedData.RuleTranslations("SimpleScriptBase"))
                ScriptCombine.AddRange(CachedData.RuleTranslations("SimpleScript"))
                ArabicString = UnjoinContig(ProcessTransform(JoinContig(ArabicString, PreString, PostString, False, False, Nothing), ScriptCombine.ToArray(), False), PreString, PostString)
            ElseIf ScriptType = TanzilReader.QuranScripts.SimpleClean Then
                Dim ScriptCombine As New List(Of IslamMetadata.IslamData.RuleTranslationCategory.RuleTranslation)
                ScriptCombine.AddRange(CachedData.RuleTranslations("SimpleScriptBase"))
                ScriptCombine.AddRange(CachedData.RuleTranslations("SimpleCleanScript"))
                ArabicString = UnjoinContig(ProcessTransform(JoinContig(ArabicString, PreString, PostString, False, False, Nothing), ScriptCombine.ToArray(), False), PreString, PostString)
            ElseIf ScriptType = TanzilReader.QuranScripts.SimpleMin Then
                Dim ScriptCombine As New List(Of IslamMetadata.IslamData.RuleTranslationCategory.RuleTranslation)
                ScriptCombine.AddRange(CachedData.RuleTranslations("SimpleScriptBase"))
                ScriptCombine.AddRange(CachedData.RuleTranslations("SimpleMinimalScript"))
                ArabicString = UnjoinContig(ProcessTransform(JoinContig(ArabicString, PreString, PostString, False, False, Nothing), ScriptCombine.ToArray(), False), PreString, PostString)
            End If
        ElseIf SrcScriptType = TanzilReader.QuranScripts.UthmaniMin Then
            If ScriptType = TanzilReader.QuranScripts.Uthmani Then
                ArabicString = UnjoinContig(ProcessTransform(JoinContig(ArabicString, PreString, PostString, False, False, Nothing), CachedData.RuleTranslations("ReverseUthmaniMinimalScript"), False), PreString, PostString)
            End If
        ElseIf SrcScriptType = TanzilReader.QuranScripts.SimpleEnhanced Then
            If ScriptType = TanzilReader.QuranScripts.Uthmani Then
                Dim ScriptCombine As New List(Of IslamMetadata.IslamData.RuleTranslationCategory.RuleTranslation)
                ScriptCombine.AddRange(CachedData.RuleTranslations("ReverseSimpleScriptBase"))
                ScriptCombine.AddRange(CachedData.RuleTranslations("ReverseSimpleEnhancedScript"))
                ArabicString = UnjoinContig(ProcessTransform(JoinContig(ArabicString, PreString, PostString, False, False, Nothing), ScriptCombine.ToArray(), False), PreString, PostString)
            End If
        ElseIf SrcScriptType = TanzilReader.QuranScripts.Simple Then
        ElseIf SrcScriptType = TanzilReader.QuranScripts.SimpleClean Then
        ElseIf SrcScriptType = TanzilReader.QuranScripts.SimpleMin Then
        End If
        Return ArabicString
    End Function
    Public Shared Function ReplaceMetadata(ArabicString As String, MetadataRule As RuleMetadata, Scheme As String, LearningMode As Boolean) As String
        For Count As Integer = 0 To CachedData.RuleTranslations("ColoringSpelledOutRules").Length - 1
            Dim Match As String = Array.Find(CachedData.RuleTranslations("ColoringSpelledOutRules")(Count).Match.Split("|"c), Function(Str As String) Array.IndexOf(New List(Of String)(Linq.Enumerable.Select(MetadataRule.Type.Split("|"c), Function(S As String) System.Text.RegularExpressions.Regex.Replace(S, "\(.*\)", String.Empty))).ToArray(), Str) <> -1)
            If Match <> Nothing Then
                Dim Str As New System.Text.StringBuilder
                Str.Append(String.Format(CachedData.RuleTranslations("ColoringSpelledOutRules")(Count).Evaluator, ArabicString.Substring(MetadataRule.Index, MetadataRule.Length)))
                If CachedData.RuleTranslations("ColoringSpelledOutRules")(Count).RuleFunc <> RuleFuncs.eNone Then
                    Dim Args As String() = RuleFunctions(CachedData.RuleTranslations("ColoringSpelledOutRules")(Count).RuleFunc - 1)(Str.ToString(), Scheme, LearningMode)
                    If Args.Length = 1 Then
                        Str.Clear()
                        Str.Append(Args(0))
                    Else
                        Dim MetaArgs As String() = System.Text.RegularExpressions.Regex.Match(MetadataRule.Type, Match + "\((.*)\)").Groups(1).Value.Split(","c)
                        Str.Clear()
                        For Index As Integer = 0 To Args.Length - 1
                            If Not Args(Index) Is Nothing And (LearningMode Or CachedData.RuleTranslations("ColoringSpelledOutRules")(Count).RuleFunc <> RuleFuncs.eLearningMode Or Index <> 0) Then
                                Str.Append(ReplaceMetadata(Args(Index), New RuleMetadata(0, Args(Index).Length, MetaArgs(Index).Replace(" "c, "|"c), Index), Scheme, LearningMode))
                            End If
                        Next
                    End If
                End If
                ArabicString = ArabicString.Insert(MetadataRule.Index + MetadataRule.Length, Str.ToString()).Remove(MetadataRule.Index, MetadataRule.Length)
            End If
        Next
        Return ArabicString
    End Function
    Public Shared Sub DoErrorCheck(ByVal ArabicString As String)
        'need to check for decomposed first
        Dim Count As Integer
        For Count = 0 To CachedData.RuleTranslations("ErrorCheck").Length - 1
            Dim Matches As System.Text.RegularExpressions.MatchCollection = System.Text.RegularExpressions.Regex.Matches(ArabicString, CachedData.RuleTranslations("ErrorCheck")(Count).Match)
            For MatchIndex As Integer = 0 To Matches.Count - 1
                If CachedData.RuleTranslations("ErrorCheck")(Count).NegativeMatch Is Nothing OrElse Matches(MatchIndex).Result(CachedData.RuleTranslations("ErrorCheck")(Count).NegativeMatch) = String.Empty Then
                    'Debug.Print(CachedData.RuleTranslations("ErrorCheck")(Count).Rule + ": " + TransliterateToScheme(ArabicString, ArabicData.TranslitScheme.Literal, String.Empty).Insert(Matches(MatchIndex).Index, "<!-- -->"))
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
        If Index = 1 Then Index = PostString.IndexOf(" "c, Index + 1)
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
    Public Shared Function TransliterateContigWithRules(ByVal ArabicString As String, ByVal PreString As String, ByVal PostString As String, PreStop As Boolean, PostStop As Boolean, Scheme As String, OptionalStops As Integer(), RuleSet As IslamMetadata.IslamData.RuleMetaSet.RuleMetadataTranslation()) As String
        Return UnjoinContig(TransliterateWithRules(JoinContig(ArabicString, PreString, PostString, PreStop, PostStop, OptionalStops), Scheme, OptionalStops, False, RuleSet), PreString, PostString)
    End Function
    Public Shared Function TransliterateWithRules(ByVal ArabicString As String, Scheme As String, OptionalStops As Integer(), LearningMode As Boolean, RuleSet As IslamMetadata.IslamData.RuleMetaSet.RuleMetadataTranslation()) As String
        Dim Count As Integer
        Dim MetadataList As New Generic.List(Of RuleMetadata)
        DoErrorCheck(ArabicString)
        For Count = 0 To RuleSet.Length - 1
            If Not RuleSet(Count).Evaluator Is Nothing Then
                Dim Matches As System.Text.RegularExpressions.MatchCollection = System.Text.RegularExpressions.Regex.Matches(ArabicString, RuleSet(Count).Match)
                For MatchIndex As Integer = 0 To Matches.Count - 1
                    Dim SubCount As Integer
                    For SubCount = 0 To RuleSet(Count).Evaluator.Length - 1
                        If (Array.IndexOf(RuleSet(Count).Evaluator(SubCount).Split("|"c), "optionalstop") <> -1 AndAlso (OptionalStops Is Nothing AndAlso Matches(MatchIndex).Groups(SubCount + 1).Value = ArabicData.ArabicSmallHighLigatureSadWithLamWithAlefMaksura OrElse (Matches(MatchIndex).Groups(SubCount + 1).Success AndAlso (Matches(MatchIndex).Groups(SubCount + 1).Index <> 0 And Matches(MatchIndex).Groups(SubCount + 1).Index <> ArabicString.Length)) OrElse (Not OptionalStops Is Nothing AndAlso Matches(MatchIndex).Groups(SubCount + 1).Value <> String.Empty AndAlso Array.IndexOf(OptionalStops, Matches(MatchIndex).Groups(SubCount + 1).Index) = -1))) OrElse (Array.IndexOf(RuleSet(Count).Evaluator(SubCount).Split("|"c), "optionalnotstop") <> -1 AndAlso (OptionalStops Is Nothing AndAlso Matches(MatchIndex).Groups(SubCount + 1).Value <> ArabicData.ArabicSmallHighLigatureSadWithLamWithAlefMaksura AndAlso ((Matches(MatchIndex).Groups(SubCount + 1).Success AndAlso Matches(MatchIndex).Groups(SubCount + 1).Length = 0 AndAlso (Matches(MatchIndex).Groups(SubCount + 1).Index = 0 Or Matches(MatchIndex).Groups(SubCount + 1).Index = ArabicString.Length)) OrElse (Not OptionalStops Is Nothing AndAlso Matches(MatchIndex).Groups(SubCount + 1).Value <> String.Empty AndAlso Array.IndexOf(OptionalStops, Matches(MatchIndex).Groups(SubCount + 1).Index) <> -1)))) Then Exit For
                    Next
                    If SubCount <> RuleSet(Count).Evaluator.Length Then Continue For
                    For SubCount = 0 To RuleSet(Count).Evaluator.Length - 1
                        If Not RuleSet(Count).Evaluator(SubCount) Is Nothing AndAlso RuleSet(Count).Evaluator(SubCount) <> String.Empty And (Matches(MatchIndex).Groups(SubCount + 1).Length <> 0 Or Array.IndexOf(AllowZeroLength, RuleSet(Count).Evaluator(SubCount)) <> -1) Then
                            MetadataList.Add(New RuleMetadata(Matches(MatchIndex).Groups(SubCount + 1).Index, Matches(MatchIndex).Groups(SubCount + 1).Length, RuleSet(Count).Evaluator(SubCount), SubCount))
                            'Debug.Print(RuleSet(Count).Name + " Index: " + CStr(Matches(MatchIndex).Groups(SubCount + 1).Index) + " Length: " + CStr(Matches(MatchIndex).Groups(SubCount + 1).Length) + " Ruling: " + RuleSet(Count).Evaluator(SubCount))
                        End If
                    Next
                Next
            End If
        Next
        MetadataList.Sort(New RuleMetadataComparer)
        Dim Index As Integer = 0
        While Index <= MetadataList.Count - 1
            Do While Index <> MetadataList.Count - 1 AndAlso MetadataList(Index).Index = MetadataList(Index + 1).Index AndAlso MetadataList(Index).Length = MetadataList(Index + 1).Length
                Dim FirstRule As String() = MetadataList(Index).Type.Split("|"c)
                Dim SecondRule As New List(Of String)
                SecondRule.AddRange(MetadataList(Index + 1).Type.Split("|"c))
                Dim FirstUpdate As New System.Text.StringBuilder
                For FirstIndex As Integer = 0 To FirstRule.Length - 1
                    Dim SecondIndex As Integer
                    For SecondIndex = 0 To SecondRule.Count - 1
                        If System.Text.RegularExpressions.Regex.Replace(FirstRule(FirstIndex), "\(.*\)", String.Empty) = System.Text.RegularExpressions.Regex.Replace(SecondRule(SecondIndex), "\(.*\)", String.Empty) Then
                            Dim Matches As String() = System.Text.RegularExpressions.Regex.Replace(FirstRule(FirstIndex), ".*\((.*)\)", "$1").Split(","c)
                            If Matches.Length <> 1 Then
                                Dim AddMatches As String() = System.Text.RegularExpressions.Regex.Replace(SecondRule(SecondIndex), ".*\((.*)\)", "$1").Split(","c)
                                If FirstUpdate.Length <> 0 Then FirstUpdate.Append("|")
                                FirstUpdate.Append(System.Text.RegularExpressions.Regex.Replace(FirstRule(FirstIndex), "\(.*\)", String.Empty) + "(")
                                For Count = 0 To Matches.Length - 1
                                    FirstUpdate.Append(Matches(Count) + If(Matches(Count) <> String.Empty And AddMatches(Count) <> String.Empty, " ", String.Empty) + AddMatches(Count))
                                    If Count <> Matches.Length - 1 Then FirstUpdate.Append(",")
                                Next
                                FirstUpdate.Append(")")
                            Else
                                FirstUpdate.Append(If(FirstUpdate.Length <> 0, "|", String.Empty) + FirstRule(FirstIndex))
                            End If
                            SecondRule.RemoveAt(SecondIndex)
                            SecondIndex -= 1
                            Exit For
                        End If
                    Next
                    If SecondIndex = SecondRule.Count Then FirstUpdate.Append(If(FirstUpdate.Length <> 0, "|", String.Empty) + FirstRule(FirstIndex))
                Next
                SecondRule.Insert(0, FirstUpdate.ToString())
                'Debug.Print("First: " + MetadataList(Index).Type + " Second: " + MetadataList(Index + 1).Type + " After: " + String.Join("|"c, SecondRule.ToArray()))
                MetadataList(Index) = New RuleMetadata(MetadataList(Index).Index, MetadataList(Index).Length, String.Join("|"c, SecondRule.ToArray()), MetadataList(Index).OrigOrder)
                MetadataList.RemoveAt(Index + 1)
            Loop
            ArabicString = ReplaceMetadata(ArabicString, MetadataList(Index), Scheme, LearningMode)
            Index += 1
        End While
        'redundant romanization rules should have -'s such as seen/teh/kaf-heh
        For Count = 0 To CachedData.RuleTranslations("RomanizationRules").Length - 1
            If CachedData.RuleTranslations("RomanizationRules")(Count).RuleFunc = RuleFuncs.eNone Then
                ArabicString = System.Text.RegularExpressions.Regex.Replace(ArabicString, CachedData.RuleTranslations("RomanizationRules")(Count).Match, CachedData.RuleTranslations("RomanizationRules")(Count).Evaluator)
            Else
                ArabicString = System.Text.RegularExpressions.Regex.Replace(ArabicString, CachedData.RuleTranslations("RomanizationRules")(Count).Match, Function(Match As System.Text.RegularExpressions.Match) RuleFunctions(CachedData.RuleTranslations("RomanizationRules")(Count).RuleFunc - 1)(Match.Result(CachedData.RuleTranslations("RomanizationRules")(Count).Evaluator), Scheme, LearningMode)(0))
            End If
        Next

        'process wasl loanwords and names
        'process loanwords and names
        'If System.Text.RegularExpressions.Regex.Match(ArabicString, "[\p{IsArabic}|\p{IsArabicPresentationForms-A}|\p{IsArabicPresentationForms-B}]").Success Then Debug.Print(ArabicString.Substring(System.Text.RegularExpressions.Regex.Match(ArabicString, "[\p{IsArabic}|\p{IsArabicPresentationForms-A}|\p{IsArabicPresentationForms-B}]").Index) + " --- " + ArabicString)
        Return ArabicString
    End Function
    Public Shared Function DecodeTranslitScheme(Str As String) As String
        'QueryString instead of Params?
        Return If(CInt(Str) >= 2, CachedData.IslamData.TranslitSchemes(CInt(Str) \ 2).Name, String.Empty)
    End Function
    Public Shared Function DecodeTranslitSchemeType(Str As String) As ArabicData.TranslitScheme
        'QueryString instead of Params?
        Return CType(If(CInt(Str) >= 2, 2 - CInt(Str) Mod 2, CInt(Str)), ArabicData.TranslitScheme)
    End Function
    Shared Function SubOutPatterns(Str As String) As String
        Return Str.Replace(CachedData.TehMarbutaStopRule, String.Empty).Replace(CachedData.TehMarbutaContinueRule, "...").Replace(CachedData.TranslateRegEx("(?=(?:{ArabicMoonLetters}|{ArabicSunLettersNoLam})(?:{ArabicSukun}|{ArabicShadda}?(?:{ArabicFathaDammaKasra}))?(?:{ArabicLetters})(?:{ArabicFatha}|{ArabicKasra})|{ArabicWaslKasraExceptions})", True), ArabicData.ArabicKasra).Replace(CachedData.TranslateRegEx("(?=(?:{ArabicMoonLetters}|{ArabicSunLettersNoLam})(?:{ArabicSukun}|{ArabicShadda}?(?:{ArabicFathaDammaKasra}))?(?:{ArabicLetters}){ArabicDamma})", True), ArabicData.ArabicDamma).Replace(CachedData.TranslateRegEx("({ArabicSunLetters}|{ArabicMoonLettersNoVowels}|{ArabicLetterWaw}|{ArabicLetterYeh}|{ArabicLetterAlefMaksura}|{ArabicSmallYeh})", True), ChrW(&H66D))
    End Function
    Shared Function GetTransliterationTable(Scheme As String) As List(Of RenderArray.RenderItem)
        Dim Items As New List(Of RenderArray.RenderItem)
        Items.AddRange(Linq.Enumerable.Select(CachedData.ArabicLettersInOrder, Function(Letter As String) New RenderArray.RenderItem(RenderArray.RenderTypes.eText, New RenderArray.RenderText() {New RenderArray.RenderText(RenderArray.RenderDisplayClass.eArabic, Letter), New RenderArray.RenderText(RenderArray.RenderDisplayClass.eTransliteration, GetSchemeValueFromSymbol(ArabicData.ArabicLetters(ArabicData.FindLetterBySymbol(Letter.Chars(0))), Scheme))})))
        Items.AddRange(Linq.Enumerable.Select(CachedData.ArabicSpecialLetters, Function(Combo As String) New RenderArray.RenderItem(RenderArray.RenderTypes.eText, New RenderArray.RenderText() {New RenderArray.RenderText(RenderArray.RenderDisplayClass.eArabic, System.Text.RegularExpressions.Regex.Replace(SubOutPatterns(Combo), "\(?\\u([0-9a-fA-F]{4})\)?", Function(Match As System.Text.RegularExpressions.Match) ChrW(Integer.Parse(Match.Groups(1).Value, Globalization.NumberStyles.HexNumber)))), New RenderArray.RenderText(RenderArray.RenderDisplayClass.eTransliteration, GetSchemeSpecialValue(Combo, GetSchemeSpecialFromMatch(Combo, False), Scheme))})))
        Items.AddRange(Linq.Enumerable.Select(CachedData.ArabicHamzas, Function(Letter As String) New RenderArray.RenderItem(RenderArray.RenderTypes.eText, New RenderArray.RenderText() {New RenderArray.RenderText(RenderArray.RenderDisplayClass.eArabic, Letter), New RenderArray.RenderText(RenderArray.RenderDisplayClass.eTransliteration, GetSchemeValueFromSymbol(ArabicData.ArabicLetters(ArabicData.FindLetterBySymbol(Letter.Chars(0))), Scheme))})))
        Items.AddRange(Linq.Enumerable.Select(CachedData.ArabicVowels, Function(Letter As String) New RenderArray.RenderItem(RenderArray.RenderTypes.eText, New RenderArray.RenderText() {New RenderArray.RenderText(RenderArray.RenderDisplayClass.eArabic, Letter), New RenderArray.RenderText(RenderArray.RenderDisplayClass.eTransliteration, GetSchemeValueFromSymbol(ArabicData.ArabicLetters(ArabicData.FindLetterBySymbol(Letter.Chars(0))), Scheme))})))
        Items.AddRange(Linq.Enumerable.Select(CachedData.ArabicMultis, Function(Combo As String) New RenderArray.RenderItem(RenderArray.RenderTypes.eText, New RenderArray.RenderText() {New RenderArray.RenderText(RenderArray.RenderDisplayClass.eArabic, Combo), New RenderArray.RenderText(RenderArray.RenderDisplayClass.eTransliteration, GetSchemeLongVowelFromString(Combo, Scheme))})))
        Items.AddRange(Linq.Enumerable.Select(CachedData.ArabicTajweed, Function(Letter As String) New RenderArray.RenderItem(RenderArray.RenderTypes.eText, New RenderArray.RenderText() {New RenderArray.RenderText(RenderArray.RenderDisplayClass.eArabic, Letter), New RenderArray.RenderText(RenderArray.RenderDisplayClass.eTransliteration, GetSchemeValueFromSymbol(ArabicData.ArabicLetters(ArabicData.FindLetterBySymbol(Letter.Chars(0))), Scheme))})))
        Return Items
    End Function
    Public Shared Function ArabicTranslitLetters() As String()
        Dim Lets As New List(Of String)
        Lets.AddRange(Linq.Enumerable.Select(CachedData.ArabicLettersInOrder, Function(Ch As String) Ch))
        Lets.AddRange(Linq.Enumerable.Select(CachedData.ArabicHamzas, Function(Ch As String) Ch))
        Lets.AddRange(CachedData.ArabicVowels)
        Lets.AddRange(Linq.Enumerable.Select(CachedData.ArabicTajweed, Function(Ch As String) Ch))
        Lets.AddRange(Linq.Enumerable.Select(CachedData.ArabicSilent, Function(Ch As String) Ch))
        Lets.AddRange(Linq.Enumerable.Select(CachedData.ArabicPunctuation, Function(Ch As String) Ch))
        Lets.AddRange(Linq.Enumerable.Select(CachedData.ArabicNums, Function(Ch As String) Ch))
        Lets.AddRange(Linq.Enumerable.Select(CachedData.NonArabicLetters, Function(Ch As String) Ch))
        Return Lets.ToArray()
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
    Public Shared Function DisplayDict() As Array()
        Dim Lines As String() = Utility.ReadAllLines(PortableMethods.Settings.GetFilePath("metadata\HansWeir.txt"))
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
    Public Shared Function DisplayCombo(SchemeType As ArabicData.TranslitScheme, Scheme As String) As Array()
        Dim Count As Integer
        Dim Output As New List(Of String())
        Output.Add(New String() {})
        Output.Add(New String() {"arabic", "transliteration", "arabic", String.Empty, String.Empty, String.Empty, String.Empty})
        Output.Add(New String() {Utility.LoadResourceString("IslamInfo_LetterName"), Utility.LoadResourceString("IslamInfo_Transliteration"), Utility.LoadResourceString("IslamInfo_Arabic"), Utility.LoadResourceString("IslamSource_ExtendedBuckwalter"), Utility.LoadResourceString("IslamInfo_Terminating"), Utility.LoadResourceString("IslamInfo_Connecting"), Utility.LoadResourceString("IslamInfo_Shaping")})
        'Dim Combos(ArabicData.Data.ArabicCombos.Length - 1) As IslamData.ArabicCombo
        'ArabicData.ArabicLetters.CopyTo(ArabicData.Data.ArabicCombos, 0)
        'Array.Sort(Combos, Function(Key As IslamData.ArabicCombo, NextKey As IslamData.ArabicCombo) Key.SymbolName.CompareTo(NextKey.SymbolName))
        For Count = 0 To ArabicData.ArabicCombos.Length - 1
            If Array.TrueForAll(ArabicData.ArabicCombos(Count).Symbol, Function(Ch As Char) GetSchemeValueFromSymbol(ArabicData.ArabicLetters(ArabicData.FindLetterBySymbol(Ch)), "ExtendedBuckwalter") <> String.Empty) Then
                Output.Add(New String() {ArabicLetterSpelling(String.Join(String.Empty, Linq.Enumerable.Select(ArabicData.ArabicCombos(Count).Symbol, Function(Sym As Char) CStr(Sym))), False, False, False), _
                                         TransliterateToScheme(ArabicLetterSpelling(String.Join(String.Empty, Linq.Enumerable.Select(ArabicData.ArabicCombos(Count).Symbol, Function(Sym As Char) CStr(Sym))), False, False, False), SchemeType, Scheme, CachedData.RuleMetas("Normal")), _
                                                    String.Join(String.Empty, Linq.Enumerable.Select(ArabicData.ArabicCombos(Count).Symbol, Function(Sym As Char) CStr(Sym))), _
                                                    TransliterateToScheme(String.Join(String.Empty, Linq.Enumerable.Select(ArabicData.ArabicCombos(Count).Symbol, Function(Sym As Char) CStr(Sym))), ArabicData.TranslitScheme.Literal, String.Empty, Nothing), _
                                                    CStr(ArabicData.ArabicCombos(Count).Terminating), _
                                                  CStr(ArabicData.ArabicCombos(Count).Connecting), _
                                                    String.Join(vbCrLf, Linq.Enumerable.Select(ArabicData.ArabicCombos(Count).Shaping, Function(Shape As Char) If(Shape = ChrW(0), String.Empty, Shape + " " + CStr(Convert.ToString(AscW(Shape), 16)) + " " + If(CheckShapingOrder(Array.IndexOf(ArabicData.ArabicCombos(Count).Shaping, Shape), ArabicData.GetUnicodeName(Shape)), String.Empty, "!!!") + ArabicData.GetUnicodeName(Shape))))})
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
        Output(1) = New List(Of String)(Linq.Enumerable.Select(Cols, Function(Str As String) ColStyles(Array.IndexOf(ColSet, Str)))).ToArray()
        Output(2) = New List(Of String)(Linq.Enumerable.Select(Cols, Function(Str As String) Utility.LoadResourceString("IslamInfo_" + Str))).ToArray()
        For Count = 0 To Symbols.Length - 1
            '"arabic", String.Empty
            'Utility.LoadResourceString("IslamInfo_LetterName"), Utility.LoadResourceString("IslamInfo_Assimilate")
            'Arabic.TransliterateFromBuckwalter(Symbols(Count).SymbolName), _
            'CStr(Symbols(Count).Assimilate),
            Output(Count + 3) = New List(Of String)(Linq.Enumerable.Select(Cols,
                Function(Str As String)
                    Select Case Array.IndexOf(ColSet, Str)
                        Case 0
                            Return ArabicLetterSpelling(CStr(Symbols(Count).Symbol), False, False, False)
                        Case 1
                            Return TransliterateToScheme(ArabicLetterSpelling(CStr(Symbols(Count).Symbol), False, False, False), SchemeType, Scheme, CachedData.RuleMetas("Normal"))
                        Case 2
                            Return ArabicData.GetUnicodeName(Symbols(Count).Symbol)
                        Case 3
                            Return CStr(Symbols(Count).Symbol)
                        Case 4
                            Return Convert.ToString(AscW(Symbols(Count).Symbol), 16)
                        Case 5
                            Return CStr(If(GetSchemeValueFromSymbol(Symbols(Count), "ExtendedBuckwalter").Length = 0, String.Empty, GetSchemeValueFromSymbol(Symbols(Count), "ExtendedBuckwalter").Chars(0)))
                        Case 6
                            Return CStr(Symbols(Count).Terminating)
                        Case 7
                            Return CStr(Symbols(Count).Connecting)
                        Case 8
                            Return If(Symbols(Count).Shaping = Nothing, String.Empty, String.Join(vbCrLf, Linq.Enumerable.Select(Symbols(Count).Shaping, Function(Shape As Char) If(Shape = ChrW(0), String.Empty, Shape + " " + CStr(Convert.ToString(AscW(Shape), 16)) + " " + If(CheckShapingOrder(Array.IndexOf(Symbols(Count).Shaping, Shape), ArabicData.GetUnicodeName(Shape)), String.Empty, "!!!") + ArabicData.GetUnicodeName(Shape)))))
                        Case Else
                            Return Nothing
                    End Select
                End Function)).ToArray()
        Next
        Return Output
    End Function
    Public Shared Function DisplayAll(SchemeType As ArabicData.TranslitScheme, Scheme As String) As Array()
        Return SymbolDisplay(Array.FindAll(ArabicData.ArabicLetters, Function(Letter As ArabicData.ArabicSymbol) GetSchemeValueFromSymbol(Letter, "ExtendedBuckwalter") <> String.Empty), SchemeType, Scheme, Nothing)
    End Function
    Public Shared Function DisplayTranslitSchemes() As Array()
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
                Strs = New String() {ArabicLetterSpelling(ArabicData.ArabicLetters(Count).Symbol, False, False, False),
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
            Strs = New String() {ArabicLetterSpelling(Str, False, False, False), String.Empty, Str, _
                                       TransliterateToScheme(Str, ArabicData.TranslitScheme.Literal, String.Empty, Nothing)}
            Array.Resize(Of String)(Strs, 3 + CachedData.IslamData.TranslitSchemes.Length)
            For SchemeCount = 0 To CachedData.IslamData.TranslitSchemes.Length - 1
                Strs(3 + SchemeCount) = GetSchemeSpecialValue(CachedData.ArabicSpecialLetters(Count), GetSchemeSpecialFromMatch(CachedData.ArabicSpecialLetters(Count), False), CachedData.IslamData.TranslitSchemes(SchemeCount).Name)
            Next
            Output.Add(Strs)
        Next
        For Count = 0 To CachedData.ArabicLongVowels.Length - 1
            Strs = New String() {ArabicLetterSpelling(CachedData.ArabicLongVowels(Count), False, False, False), _
                                       String.Empty, CachedData.ArabicLongVowels(Count), _
                                       TransliterateToScheme(CachedData.ArabicLongVowels(Count), ArabicData.TranslitScheme.Literal, String.Empty, Nothing)}
            Array.Resize(Of String)(Strs, 3 + CachedData.IslamData.TranslitSchemes.Length)
            For SchemeCount = 0 To CachedData.IslamData.TranslitSchemes.Length - 1
                Strs(3 + SchemeCount) = GetSchemeLongVowelFromString(CachedData.ArabicLongVowels(Count), CachedData.IslamData.TranslitSchemes(SchemeCount).Name)
            Next
            Output.Add(Strs)
        Next
        Return Output.ToArray()
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
    Shared _NounIDs As Dictionary(Of String, List(Of IslamData.GrammarSet.GrammarNoun))
    Public Shared ReadOnly Property NounIDs As Dictionary(Of String, List(Of IslamData.GrammarSet.GrammarNoun))
        Get
            If _NounIDs Is Nothing Then
                _NounIDs = New Dictionary(Of String, List(Of IslamData.GrammarSet.GrammarNoun))
                For Count As Integer = 0 To CachedData.IslamData.Grammar.Nouns.Length - 1
                    If Not _NounIDs.ContainsKey(CachedData.IslamData.Grammar.Nouns(Count).TranslationID) Then
                        _NounIDs.Add(CachedData.IslamData.Grammar.Nouns(Count).TranslationID, New List(Of IslamData.GrammarSet.GrammarNoun) From {CachedData.IslamData.Grammar.Nouns(Count)})
                    Else
                        'Debug.Print("Duplicate Noun ID: " + CachedData.IslamData.Grammar.Nouns(Count).TranslationID)
                    End If
                    Dim Noun As IslamData.GrammarSet.GrammarNoun = CachedData.IslamData.Grammar.Nouns(Count)
                    If Not CachedData.IslamData.Grammar.Nouns(Count).Grammar Is Nothing AndAlso CachedData.IslamData.Grammar.Nouns(Count).Grammar.Length <> 0 Then
                        For Each Str As String In CachedData.IslamData.Grammar.Nouns(Count).Grammar.Split(","c)
                            If Not _NounIDs.ContainsKey(Str) Then
                                _NounIDs.Add(Str, New List(Of IslamData.GrammarSet.GrammarNoun))
                            End If
                            _NounIDs(Str).Add(Noun)
                        Next
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
                        'Debug.Print("Duplicate Transform ID: " + CachedData.IslamData.Grammar.Transforms(Count).TranslationID)
                    End If
                    Dim Transform As IslamData.GrammarSet.GrammarTransform = CachedData.IslamData.Grammar.Transforms(Count)
                    If Not CachedData.IslamData.Grammar.Transforms(Count).Grammar Is Nothing AndAlso CachedData.IslamData.Grammar.Transforms(Count).Grammar.Length <> 0 Then
                        For Each GroupStr As String In CachedData.IslamData.Grammar.Transforms(Count).Grammar.Split(","c)
                            For Each Str As String In GroupStr.Split("|"c)
                                If Not _TransformIDs.ContainsKey(Str) Then
                                    _TransformIDs.Add(Str, New List(Of IslamData.GrammarSet.GrammarTransform))
                                End If
                                _TransformIDs(Str).Add(Transform)
                            Next
                        Next
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
        Dim Transforms As IslamData.GrammarSet.GrammarTransform()() = New List(Of IslamData.GrammarSet.GrammarTransform())(Linq.Enumerable.Select(IDs, Function(ID As String) GetTransform(ID))).ToArray()
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
                        'Debug.Print("Duplicate Particle ID: " + CachedData.IslamData.Grammar.Particles(Count).TranslationID)
                    End If
                    Dim Particle As IslamData.GrammarSet.GrammarParticle = CachedData.IslamData.Grammar.Particles(Count)
                    If Not CachedData.IslamData.Grammar.Particles(Count).Grammar Is Nothing AndAlso CachedData.IslamData.Grammar.Particles(Count).Grammar.Length <> 0 Then
                        For Each Str As String In CachedData.IslamData.Grammar.Particles(Count).Grammar.Split(","c)
                            If Not _ParticleIDs.ContainsKey(Str) Then
                                _ParticleIDs.Add(Str, New List(Of IslamData.GrammarSet.GrammarParticle))
                            End If
                            _ParticleIDs(Str).Add(Particle)
                        Next
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
                        'Debug.Print("Duplicate Verb ID: " + CachedData.IslamData.Grammar.Verbs(Count).TranslationID)
                    End If
                    Dim Verb As IslamData.GrammarSet.GrammarVerb = CachedData.IslamData.Grammar.Verbs(Count)
                    If Not CachedData.IslamData.Grammar.Verbs(Count).Grammar Is Nothing AndAlso CachedData.IslamData.Grammar.Verbs(Count).Grammar.Length <> 0 Then
                        For Each Str As String In CachedData.IslamData.Grammar.Verbs(Count).Grammar.Split(","c)
                            If Not _VerbIDs.ContainsKey(Str) Then
                                _VerbIDs.Add(Str, New List(Of IslamData.GrammarSet.GrammarVerb))
                            End If
                            _VerbIDs(Str).Add(Verb)
                        Next
                    End If
                Next
            End If
            Return _VerbIDs
        End Get
    End Property
    Public Shared Function GetVerb(ID As String) As IslamData.GrammarSet.GrammarVerb()
        Return If(VerbIDs.ContainsKey(ID), VerbIDs(ID).ToArray(), Nothing)
    End Function
    Public Shared Function ApplyTransform(Transforms As IslamData.GrammarSet.GrammarTransform(), Str As String) As String
        Dim Text As String = Str
        For Count = 0 To Transforms.Length - 1
            Text = New System.Text.RegularExpressions.Regex(If(Transforms(Count).From Is Nothing, "$", CachedData.TranslateRegEx(Transforms(Count).From, True))).Replace(Text, CachedData.TranslateRegEx(Transforms(Count).Text, False), 1)
        Next
        Return Text
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
<Xml.Serialization.XmlRoot("islamdata")> _
Public Class IslamData
    Public Structure ListCategory
        Public Structure Word
            <Xml.Serialization.XmlAttribute("text")> _
            Public Text As String
            <Xml.Serialization.XmlAttribute("id")> _
            Public TranslationID As String
        End Structure
        <Xml.Serialization.XmlAttribute("title")> _
        Public Title As String
        <Xml.Serialization.XmlElement("word")> _
        Public Words() As Word
    End Structure
    <Xml.Serialization.XmlArray("lists")> _
    <Xml.Serialization.XmlArrayItem("category")> _
    Public Lists() As ListCategory

    Public Structure Phrase
        <Xml.Serialization.XmlAttribute("text")> _
        Public Text As String
        <Xml.Serialization.XmlAttribute("id")> _
        Public TranslationID As String
    End Structure
    <Xml.Serialization.XmlArray("phrases")> _
    <Xml.Serialization.XmlArrayItem("word")> _
    Public Phrases() As Phrase

    Public Structure AbbrevWord
        <Xml.Serialization.XmlAttribute("text")> _
        Public Text As String
        <Xml.Serialization.XmlAttribute("font")> _
        Public Font As String
        <Xml.Serialization.XmlAttribute("id")> _
        Public TranslationID As String
    End Structure
    <Xml.Serialization.XmlArray("abbreviations")> _
    <Xml.Serialization.XmlArrayItem("word")> _
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
            <Xml.Serialization.XmlAttribute("text")> _
            Public Text As String
            <Xml.Serialization.XmlAttribute("id")> _
            Public TranslationID As String
            <Xml.Serialization.XmlAttribute("grammar")> _
            Public Grammar As String
            <Xml.Serialization.XmlAttribute("match")> _
            Public Match As String
            <Xml.Serialization.XmlAttribute("from")> _
            Public From As String
        End Structure
        <Xml.Serialization.XmlArray("transforms")> _
        <Xml.Serialization.XmlArrayItem("transform")> _
        Public Transforms() As GrammarTransform
        Public Structure GrammarParticle
            <Xml.Serialization.XmlAttribute("text")> _
            Public Text As String
            <Xml.Serialization.XmlAttribute("id")> _
            Public TranslationID As String
            <Xml.Serialization.XmlAttribute("grammar")> _
            Public Grammar As String
        End Structure
        <Xml.Serialization.XmlArray("particles")> _
        <Xml.Serialization.XmlArrayItem("particle")> _
        Public Particles() As GrammarParticle
        Public Structure GrammarNoun
            <Xml.Serialization.XmlAttribute("text")> _
            Public Text As String
            <Xml.Serialization.XmlAttribute("id")> _
            Public TranslationID As String
            <Xml.Serialization.XmlAttribute("plural")> _
            Public Plural As String
            <Xml.Serialization.XmlAttribute("grammar")> _
            Public Grammar As String
        End Structure
        <Xml.Serialization.XmlArray("nouns")> _
        <Xml.Serialization.XmlArrayItem("noun")> _
        Public Nouns() As GrammarNoun
        Public Structure GrammarVerb
            <Xml.Serialization.XmlAttribute("text")> _
            Public Text As String
            <Xml.Serialization.XmlAttribute("id")> _
            Public TranslationID As String
            <Xml.Serialization.XmlAttribute("poss")> _
            Public Possessives As String
            <Xml.Serialization.XmlAttribute("grammar")> _
            Public Grammar As String
        End Structure
        <Xml.Serialization.XmlArray("verbs")> _
        <Xml.Serialization.XmlArrayItem("verb")> _
        Public Verbs() As GrammarVerb
    End Structure
    <Xml.Serialization.XmlElement("grammar")> _
    Public Grammar As GrammarSet

    Public Structure TranslitScheme
        <Xml.Serialization.XmlAttribute("name")> _
        Public Name As String
        Public Alphabet() As String
        <Xml.Serialization.XmlAttribute("alphabet")> _
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
        <Xml.Serialization.XmlAttribute("hamza")> _
        Property HamzaParse As String
            Get
                If Hamza.Length = 0 Then Return String.Empty
                Return String.Join("|"c, Linq.Enumerable.Select(Hamza, Function(Str As String) Str.Replace("|", "&pipe;")))
            End Get
            Set(value As String)
                If Not value Is Nothing Then
                    Hamza = New List(Of String)(Linq.Enumerable.Select(value.Split("|"c), Function(Str As String) Str.Replace("&pipe;", "|"))).ToArray()
                End If
            End Set
        End Property
        Public SpecialLetters() As String
        <Xml.Serialization.XmlAttribute("literals")> _
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
        <Xml.Serialization.XmlAttribute("fathadammakasratanweenlongvowelsdipthongsshaddasukun")> _
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
        <Xml.Serialization.XmlAttribute("multis")> _
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
        <Xml.Serialization.XmlAttribute("gutterals")> _
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
        <Xml.Serialization.XmlAttribute("tajweed")> _
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
        <Xml.Serialization.XmlAttribute("silent")> _
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
        <Xml.Serialization.XmlAttribute("punctuation")> _
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
        <Xml.Serialization.XmlAttribute("number")> _
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
        <Xml.Serialization.XmlAttribute("nonarabic")> _
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
    <Xml.Serialization.XmlArray("translitschemes")> _
    <Xml.Serialization.XmlArrayItem("scheme")> _
    Public TranslitSchemes() As TranslitScheme
    Structure ArabicCapInfo
        <Xml.Serialization.XmlAttribute("name")> _
        Public Name As String
        <Xml.Serialization.XmlAttribute("text")> _
        Public _Text As String
        ReadOnly Property Text As String()
            Get
                Return _Text.Split({"  "}, StringSplitOptions.None)
            End Get
        End Property
    End Structure
    <Xml.Serialization.XmlArray("arabiccaptures")> _
    <Xml.Serialization.XmlArrayItem("caps")> _
    Public ArabicCaptures() As ArabicCapInfo
    Structure ArabicNumInfo
        <Xml.Serialization.XmlAttribute("name")> _
        Public Name As String
        <Xml.Serialization.XmlAttribute("text")> _
        Public _Text As String
        ReadOnly Property Text As String()
            Get
                Return _Text.Split(" "c)
            End Get
        End Property
    End Structure
    <Xml.Serialization.XmlArray("arabicnumbers")> _
    <Xml.Serialization.XmlArrayItem("nums")> _
    Public ArabicNumbers() As ArabicNumInfo
    Structure ArabicPattern
        <Xml.Serialization.XmlAttribute("name")> _
        Public Name As String
        <Xml.Serialization.XmlAttribute("match")> _
        Public Match As String
    End Structure
    <Xml.Serialization.XmlArray("arabicpatterns")> _
    <Xml.Serialization.XmlArrayItem("pattern")> _
    Public ArabicPatterns() As ArabicPattern
    Structure ArabicGroup
        <Xml.Serialization.XmlAttribute("name")> _
        Public Name As String
        <Xml.Serialization.XmlAttribute("text")> _
        Public _Text As String
        ReadOnly Property Text As String()
            Get
                Return _Text.Split(" "c)
            End Get
        End Property
    End Structure
    <Xml.Serialization.XmlArray("arabicgroups")> _
    <Xml.Serialization.XmlArrayItem("group")> _
    Public ArabicGroups() As ArabicGroup
    Structure ColorRule
        <Xml.Serialization.XmlAttribute("name")> _
        Public Name As String
        <Xml.Serialization.XmlAttribute("match")> _
        Public Match As String
        <Xml.Serialization.XmlAttribute("color")> _
        Public _Color As String
        ReadOnly Property Color As Integer
            Get
                If _Color.Contains(",") Then
                    Dim RGB As Byte() = New List(Of Byte)(Linq.Enumerable.Select(_Color.Split(","c), Function(Str As String) CByte(Str))).ToArray()
                    Return Utility.MakeArgb(&HFF, RGB(0), RGB(1), RGB(2))
                Else
                    Return Utility.ColorFromName(_Color)
                End If
            End Get
        End Property
    End Structure
    <Xml.Serialization.XmlArray("colorrules")> _
    <Xml.Serialization.XmlArrayItem("colorrule")> _
    Public ColorRules() As ColorRule

    Structure RuleTranslationCategory
        Structure RuleTranslation
            <Xml.Serialization.XmlAttribute("name")> _
            Public Name As String
            <Xml.Serialization.XmlAttribute("match")> _
            Public Match As String
            <Xml.Serialization.XmlAttribute("evaluator")> _
            Public Evaluator As String
            <Xml.Serialization.XmlAttribute("negativematch")> _
            Public NegativeMatch As String
            <Xml.Serialization.XmlAttribute("rulefunc")> _
            Public _RuleFunc As String
            ReadOnly Property RuleFunc As Arabic.RuleFuncs
                Get
                    If _RuleFunc = "eLearningMode" Then Return Arabic.RuleFuncs.eLearningMode
                    If _RuleFunc = "eDivideTanween" Then Return Arabic.RuleFuncs.eDivideTanween
                    If _RuleFunc = "eLeadingGutteral" Then Return Arabic.RuleFuncs.eLeadingGutteral
                    If _RuleFunc = "eLookupLetter" Then Return Arabic.RuleFuncs.eLookupLetter
                    If _RuleFunc = "eLookupLongVowelDipthong" Then Return Arabic.RuleFuncs.eLookupLongVowelDipthong
                    If _RuleFunc = "eSpellLetter" Then Return Arabic.RuleFuncs.eSpellLetter
                    If _RuleFunc = "eSpellLongLetter" Then Return Arabic.RuleFuncs.eSpellLongLetter
                    If _RuleFunc = "eSpellLongMergedLetter" Then Return Arabic.RuleFuncs.eSpellLongMergedLetter
                    If _RuleFunc = "eSpellNumber" Then Return Arabic.RuleFuncs.eSpellNumber
                    If _RuleFunc = "eTrailingGutteral" Then Return Arabic.RuleFuncs.eTrailingGutteral
                    If _RuleFunc = "eUpperCase" Then Return Arabic.RuleFuncs.eUpperCase
                    If _RuleFunc = "eResolveAmbiguity" Then Return Arabic.RuleFuncs.eResolveAmbiguity
                    If _RuleFunc = "eObligatory" Then Return Arabic.RuleFuncs.eObligatory
                    'If _RuleFunc = "eNone" Then
                    Return Arabic.RuleFuncs.eNone
                End Get
            End Property
        End Structure
        <Xml.Serialization.XmlAttribute("name")> _
        Public Name As String
        <Xml.Serialization.XmlElement("rule")> _
        Public Rules() As RuleTranslation
    End Structure
    <Xml.Serialization.XmlArray("translitrules")> _
    <Xml.Serialization.XmlArrayItem("ruleset")> _
    Public RuleSets() As RuleTranslationCategory
    Structure VerificationData
        <Xml.Serialization.XmlAttribute("name")> _
        Public Name As String
        <Xml.Serialization.XmlAttribute("match")> _
        Public Match As String
        <Xml.Serialization.XmlAttribute("evaluator")> _
        Public _Evaluator As String
        ReadOnly Property Evaluator As String()
            Get
                Return _Evaluator.Split("|"c)
            End Get
        End Property
        <Xml.Serialization.XmlAttribute("metarules")> _
        Public _MetaRules As String
        ReadOnly Property MetaRules As String()
            Get
                Return _MetaRules.Split("|"c)
            End Get
        End Property
    End Structure
    <Xml.Serialization.XmlArray("verificationset")> _
    <Xml.Serialization.XmlArrayItem("verification")> _
    Public VerificationSet() As VerificationData
    Structure RuleMetaSet
        Structure RuleMetadataTranslation
            <Xml.Serialization.XmlAttribute("name")> _
            Public Name As String
            <Xml.Serialization.XmlAttribute("match")> _
            Public Match As String
            <Xml.Serialization.XmlAttribute("evaluator")> _
            Public _Evaluator As String
            ReadOnly Property Evaluator As String()
                Get
                    Return _Evaluator.Split(";"c)
                End Get
            End Property
        End Structure
        <Xml.Serialization.XmlAttribute("from")> _
        Public From As String
        <Xml.Serialization.XmlAttribute("name")> _
        Public Name As String
        <Xml.Serialization.XmlElement("metarule")> _
        Public _Rules() As RuleMetadataTranslation
        ReadOnly Property Rules As RuleMetadataTranslation()
            Get
                Dim BuildRules As New List(Of RuleMetadataTranslation)
                If From <> String.Empty Then BuildRules.AddRange(TanzilReader.GetMetaRuleSet(From).Rules)
                BuildRules.AddRange(_Rules)
                Return BuildRules.ToArray()
            End Get
        End Property
    End Structure
    <Xml.Serialization.XmlArray("metarules")> _
    <Xml.Serialization.XmlArrayItem("metaruleset")> _
    Public MetaRules() As RuleMetaSet
    Structure LanguageInfo
        <Xml.Serialization.XmlAttribute("code")> _
        Public Code As String
        <Xml.Serialization.XmlAttribute("rtl")> _
        Public IsRTL As Boolean
    End Structure
    <Xml.Serialization.XmlArray("languages")> _
    <Xml.Serialization.XmlArrayItem("language")> _
    Public LanguageList() As LanguageInfo

    Structure ArabicFontList
        <Xml.Serialization.XmlAttribute("id")> _
        Public ID As String
        <Xml.Serialization.XmlAttribute("name")> _
        Public Name As String
        <Xml.Serialization.XmlAttribute("family")> _
        Public Family As String
        <Xml.Serialization.XmlAttribute("embedname")> _
        Public EmbedName As String
        <Xml.Serialization.XmlAttribute("file")> _
        Public FileName As String
        <Xml.Serialization.XmlAttribute("scale")> _
        <ComponentModel.DefaultValueAttribute(-1.0)> _
        Public Scale As Double
    End Structure
    <Xml.Serialization.XmlArray("arabicfonts")> _
    <Xml.Serialization.XmlArrayItem("arabicfont")> _
    Public ArabicFonts() As ArabicFontList

    Public Structure ScriptFont
        Public Structure Font
            <Xml.Serialization.XmlAttribute("id")> _
            Public ID As String
        End Structure
        <Xml.Serialization.XmlAttribute("name")> _
        Public Name As String
        <Xml.Serialization.XmlElement("font")> _
        Public FontList() As Font
    End Structure

    <Xml.Serialization.XmlArray("scriptfonts")> _
    <Xml.Serialization.XmlArrayItem("scriptfont")> _
    Public ScriptFonts() As ScriptFont

    Public Class TranslationsInfo
        Public Structure TranslationInfo
            <Xml.Serialization.XmlAttribute("name")> _
            Public Name As String
            <Xml.Serialization.XmlAttribute("file")> _
            Public FileName As String
            <Xml.Serialization.XmlAttribute("translator")> _
            Public Translator As String
        End Structure
        <Xml.Serialization.XmlAttribute("default")> _
        Public DefaultTranslation As String
        <Xml.Serialization.XmlElement("translation")> _
        Public TranslationList() As TranslationInfo
    End Class
    Structure VerseNumberScheme
        <Xml.Serialization.XmlAttribute("name")> _
        Public Name As String
        Public ExtraVerses As Integer()()
        <Xml.Serialization.XmlAttribute("extraverses")> _
        Property ExtraVersesStr As String
            Get
                Return String.Join(",", Linq.Enumerable.Select(ExtraVerses, Function(Ints As Integer()) String.Join(":", Linq.Enumerable.Select(Ints, Function(Int As Integer) CStr(Int)))))
            End Get
            Set(value As String)
                If value Is Nothing OrElse value.Length = 0 Then
                    ExtraVerses = Nothing
                Else
                    ExtraVerses = New List(Of Integer())(Linq.Enumerable.Select(value.Split(","c), Function(Str As String) New List(Of Integer)(Linq.Enumerable.Select(Str.Split(":"c), Function(Verse As String) CInt(Verse))).ToArray())).ToArray()
                End If
            End Set
        End Property
        Public CombinedVerses As Integer()()
        <Xml.Serialization.XmlAttribute("combinedverses")> _
        Property CombinedVersesStr As String
            Get
                Return String.Join(",", Linq.Enumerable.Select(CombinedVerses, Function(Ints As Integer()) String.Join(":", Linq.Enumerable.Select(Ints, Function(Int As Integer) CStr(Int)))))
            End Get
            Set(value As String)
                If value Is Nothing OrElse value.Length = 0 Then
                    CombinedVerses = Nothing
                Else
                    CombinedVerses = New List(Of Integer())(Linq.Enumerable.Select(value.Split(","c), Function(Str As String) New List(Of Integer)(Linq.Enumerable.Select(Str.Split(":"c), Function(Verse As String) CInt(Verse))).ToArray())).ToArray()
                End If
            End Set
        End Property
    End Structure
    <Xml.Serialization.XmlArray("versenumberschemes")> _
    <Xml.Serialization.XmlArrayItem("versenumberscheme")> _
    Public VerseNumberSchemes As VerseNumberScheme()

    <Xml.Serialization.XmlElement("translations")> _
    Public Translations As TranslationsInfo
    Structure QuranSelection
        Structure QuranSelectionInfo
            <Xml.Serialization.XmlAttribute("chapter")> _
            Public ChapterNumber As Integer
            <Xml.Serialization.XmlAttribute("startverse")> _
            Public VerseNumber As Integer
            <ComponentModel.DefaultValueAttribute(1)> _
            <Xml.Serialization.XmlAttribute("startword")> _
            Public WordNumber As Integer
            <ComponentModel.DefaultValueAttribute(0)> _
            <Xml.Serialization.XmlAttribute("endword")> _
            Public EndWordNumber As Integer
            <ComponentModel.DefaultValueAttribute(0)> _
            <Xml.Serialization.XmlAttribute("endverse")> _
            Public ExtraVerseNumber As Integer
        End Structure
        <Xml.Serialization.XmlAttribute("description")> _
        Public Description As String
        <Xml.Serialization.XmlElement("verse")> _
        Public SelectionInfo As QuranSelectionInfo()
    End Structure
    <Xml.Serialization.XmlArray("quranselections")> _
    <Xml.Serialization.XmlArrayItem("quranselection")> _
    Public QuranSelections As QuranSelection()
    Structure QuranDivision
        <Xml.Serialization.XmlAttribute("description")> _
        Public Description As String
    End Structure
    <Xml.Serialization.XmlArray("qurandivisions")> _
    <Xml.Serialization.XmlArrayItem("division")> _
    Public QuranDivisions As QuranDivision()

    Structure QuranChapter
        <Xml.Serialization.XmlAttribute("index")> _
        Public Index As String
        <Xml.Serialization.XmlAttribute("name")> _
        Public Name As String
        <ComponentModel.DefaultValueAttribute(0)> _
        <Xml.Serialization.XmlAttribute("uniqueletters")> _
        Public UniqueLetters As Integer
    End Structure
    <Xml.Serialization.XmlArray("quranchapters")> _
    <Xml.Serialization.XmlArrayItem("chapter")> _
    Public QuranChapters As QuranChapter()

    Structure QuranPart
        <Xml.Serialization.XmlAttribute("index")> _
        Public Index As String
        <Xml.Serialization.XmlAttribute("name")> _
        Public Name As String
        <Xml.Serialization.XmlAttribute("id")> _
        Public ID As String
    End Structure
    <Xml.Serialization.XmlArray("quranparts")> _
    <Xml.Serialization.XmlArrayItem("part")> _
    Public QuranParts As QuranPart()

    Structure CollectionInfo
        Structure CollTranslationInfo
            <Xml.Serialization.XmlAttribute("name")> _
            Public Name As String
            <Xml.Serialization.XmlAttribute("file")> _
            Public FileName As String
        End Structure
        <Xml.Serialization.XmlAttribute("name")> _
        Public Name As String
        <Xml.Serialization.XmlAttribute("file")> _
        Public FileName As String
        <Xml.Serialization.XmlAttribute("default")> _
        Public DefaultTranslation As String
        <Xml.Serialization.XmlArray("translations")> _
        <Xml.Serialization.XmlArrayItem("translation")> _
        Public Translations() As CollTranslationInfo
    End Structure
    <Xml.Serialization.XmlArray("hadithcollections")> _
    <Xml.Serialization.XmlArrayItem("collection")> _
    Public Collections() As CollectionInfo
    Structure PartOfSpeechInfo
        <Xml.Serialization.XmlAttribute("symbol")> _
        Public Symbol As String
        <Xml.Serialization.XmlAttribute("id")> _
        Public Id As String
        <Xml.Serialization.XmlAttribute("color")> _
        Public _Color As String
        ReadOnly Property Color As Integer
            Get
                If _Color.Contains(",") Then
                    Dim RGB As Byte() = New List(Of Byte)(Linq.Enumerable.Select(_Color.Split(","c), Function(Str As String) CByte(Str))).ToArray()
                    Return Utility.MakeArgb(&HFF, RGB(0), RGB(1), RGB(2))
                Else
                    Return Utility.ColorFromName(_Color)
                End If
            End Get
        End Property
    End Structure
    <Xml.Serialization.XmlArray("partsofspeech")> _
    <Xml.Serialization.XmlArrayItem("pos")> _
    Public PartsOfSpeech() As PartOfSpeechInfo
End Class
Public Class CachedData
    'need disk and memory cache as time consuming to read or build
    Shared _ObjIslamData As IslamData
    Shared _RuleTranslations As New Dictionary(Of String, IslamData.RuleTranslationCategory.RuleTranslation())
    Shared _RuleMetas As New Dictionary(Of String, IslamData.RuleMetaSet.RuleMetadataTranslation())
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
    Private Shared Characteristics As String() = {"Audibility", "Whispering", "Weakness", "Moderation", "Strength", "Lowness", "Elevation", "Opening", "Closing", "Restraint", "Fluency", "Vibration", "Inclination", "Repetition", "Whistling", "Diffusion", "Elongation", "Nasal", "Ease"}
    Private Shared PatternAllowed As String() = {"ArabicSpecialLetters", "ArabicAssimilateSameWord", "ArabicAssimilateAcrossWord", "ArabicAssimilateLeenAcrossWord", "ArabicBehSakinCombos", "ArabicTehSakinCombos", "ArabicThehSakinCombos", "ArabicJeemSakinCombos", "ArabicHahSakinCombos", "ArabicKhahSakinCombos", "ArabicDalSakinCombos", "ArabicThalSakinCombos", "ArabicRehSakinCombos", "ArabicZainSakinCombos", "ArabicSeenSakinCombos", "ArabicSheenSakinCombos", "ArabicSadSakinCombos", "ArabicDadSakinCombos", "ArabicTahSakinCombos", "ArabicZahSakinCombos", "ArabicAinSakinCombos", "ArabicGhainSakinCombos", "ArabicFehSakinCombos", "ArabicQafSakinCombos", "ArabicKafSakinCombos", "ArabicLamSakinCombos", "ArabicMeemSakinCombos", "ArabicNoonSakinCombos", "ArabicHehSakinCombos", "ArabicWawSakinCombos", "ArabicYehSakinCombos", "ArabicHamzaSakinCombos"}
    Public Shared Function GetGroup(Name As String) As String()
        Dim Count As Integer
        If _SavedGroups.ContainsKey(Name) Then Return _SavedGroups(Name)
        For Count = 0 To CachedData.IslamData.ArabicGroups.Length - 1
            If CachedData.IslamData.ArabicGroups(Count).Name = Name Then
                _SavedGroups.Add(Name, New List(Of String)(Linq.Enumerable.Select(CachedData.IslamData.ArabicGroups(Count).Text, Function(Str As String) TranslateRegEx(Str, Array.IndexOf(PatternAllowed, Name) <> -1 Or Array.IndexOf(Characteristics, Name) <> -1))).ToArray())
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
    Public Shared Function ArabicLetterCharacteristics() As String
        Dim MutualExclusiveChars As String() = {"Audibility", "Whispering", "Weakness", "Moderation", "Strength", "Lowness", "Elevation", "Opening", "Closing", "Restraint", "Fluency"}
        Dim Lets As New List(Of String)
        Lets.AddRange(Linq.Enumerable.Select(ArabicSunLetters, Function(Str As String) ArabicData.MakeUniRegEx(Str)))
        Lets.AddRange(Linq.Enumerable.Select(ArabicMoonLettersNoVowels, Function(Str As String) ArabicData.MakeUniRegEx(Str)))
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
                        Return ArabicData.MakeRegMultiEx(New List(Of String)(Linq.Enumerable.Select(GetCap(Match.Groups(1).Value.Split(";"c)(0)),
                            Function(Str As String)
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
                            End Function)).ToArray())
                    End If
                    If GetNum(Match.Groups(1).Value).Length <> 0 Then
                        Return ArabicData.MakeRegMultiEx(New List(Of String)(Linq.Enumerable.Select(GetNum(Match.Groups(1).Value), Function(Str As String) ArabicData.MakeUniRegEx(Arabic.TransliterateFromBuckwalter(Str)))).ToArray())
                    End If
                    If GetGroup(Match.Groups(1).Value).Length <> 0 Then
                        Return ArabicData.MakeRegMultiEx(New List(Of String)(Linq.Enumerable.Select(GetGroup(Match.Groups(1).Value), Function(Str As String) ArabicData.MakeUniRegEx(Str))).ToArray())
                    End If
                End If
                If System.Text.RegularExpressions.Regex.Match(Match.Groups(1).Value, "0x([0-9a-fA-F]{4})").Success Then
                    Return If(bAll, ArabicData.MakeUniRegEx(ChrW(Integer.Parse(Match.Groups(1).Value.Substring(2), System.Globalization.NumberStyles.HexNumber))), ChrW(Integer.Parse(Match.Groups(1).Value.Substring(2), System.Globalization.NumberStyles.HexNumber)))
                End If
                If ArabicCamelCaseDict.ContainsKey(Match.Groups(1).Value) Then
                    Return If(bAll, ArabicData.MakeUniRegEx(ArabicData.ArabicLetters(ArabicCamelCaseDict(Match.Groups(1).Value)).Symbol), ArabicData.ArabicLetters(ArabicCamelCaseDict(Match.Groups(1).Value)).Symbol)
                End If
                If ArabicComboCamelCaseDict.ContainsKey(Match.Groups(1).Value) Then
                    Return If(bAll, ArabicData.MakeUniRegEx(If(ArabicData.ArabicCombos(ArabicComboCamelCaseDict(Match.Groups(1).Value)).Shaping.Length = 1, ArabicData.ArabicCombos(ArabicComboCamelCaseDict(Match.Groups(1).Value)).Shaping(0), String.Join(String.Empty, Linq.Enumerable.Select(ArabicData.ArabicCombos(ArabicComboCamelCaseDict(Match.Groups(1).Value)).Symbol, Function(Sym As Char) CStr(Sym))))), If(ArabicData.ArabicCombos(ArabicComboCamelCaseDict(Match.Groups(1).Value)).Shaping.Length = 1, ArabicData.ArabicCombos(ArabicComboCamelCaseDict(Match.Groups(1).Value)).Shaping(0), String.Join(String.Empty, Linq.Enumerable.Select(ArabicData.ArabicCombos(ArabicComboCamelCaseDict(Match.Groups(1).Value)).Symbol, Function(Sym As Char) CStr(Sym)))))
                End If
                '{0} ignore
                'If Not IsNumeric(Match.Groups(1).Value) Then Debug.Print("Unknown Group: " + Match.Groups(1).Value)
                Return Match.Value
            End Function)
    End Function
    Public Shared ReadOnly Property RuleMetas(Name As String) As IslamData.RuleMetaSet.RuleMetadataTranslation()
        Get
            If Not _RuleMetas.ContainsKey(Name) Then
                _RuleMetas.Add(Name, TanzilReader.GetMetaRuleSet(Name).Rules)
                For SubCount As Integer = 0 To _RuleMetas(Name).Length - 1
                    _RuleMetas(Name)(SubCount).Match = TranslateRegEx(_RuleMetas(Name)(SubCount).Match, True)
                Next
            End If
            Return _RuleMetas(Name)
        End Get
    End Property
    Public Shared ReadOnly Property RuleTranslations(Name As String) As IslamData.RuleTranslationCategory.RuleTranslation()
        Get
            If Not _RuleTranslations.ContainsKey(Name) Then
                _RuleTranslations.Add(Name, GetRuleSet(Name))
            End If
            Return _RuleTranslations(Name)
        End Get
    End Property
    Shared _XMLDocMain As Xml.Linq.XDocument 'Tanzil Quran data
    Shared _XMLDocInfo As Xml.Linq.XDocument 'Tanzil metadata
    Shared _XMLDocInfos As Collections.Generic.List(Of Xml.Linq.XDocument) 'Hadiths
    Shared _RootDictionary As New Generic.Dictionary(Of String, List(Of Integer()))
    Shared _FormDictionary As New Generic.Dictionary(Of String, List(Of Integer()))
    Shared _TagDictionary As New Generic.Dictionary(Of String, Generic.Dictionary(Of String, List(Of Integer())))
    Shared _WordDictionary As New Generic.Dictionary(Of String, List(Of String))
    Shared _RealWordDictionary As New Generic.Dictionary(Of String, List(Of Integer()))
    Shared _LetterDictionary As New Generic.Dictionary(Of Char, Generic.Dictionary(Of String, List(Of Integer())))
    Shared _LetterPreDictionary As New Generic.Dictionary(Of Char, Generic.Dictionary(Of String, List(Of Integer())))
    Shared _LetterSufDictionary As New Generic.Dictionary(Of Char, Generic.Dictionary(Of String, List(Of Integer())))
    Shared _PreDictionary As New Generic.Dictionary(Of String, List(Of Integer()))
    Shared _SufDictionary As New Generic.Dictionary(Of String, List(Of Integer()))
    Shared _LocDictionary As New Generic.Dictionary(Of String, Object())
    Shared _IsolatedLetterDictionary As New Generic.Dictionary(Of Char, List(Of Integer()))
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
        Dim Lines As String() = Utility.ReadAllLines(PortableMethods.Settings.GetFilePath("metadata\quranic-corpus-morphology-0.4.txt"))
        For Count As Integer = 0 To Lines.Length - 1
            If Lines(Count).Length <> 0 AndAlso Lines(Count).Chars(0) <> "#" Then
                'LOCATION	FORM	TAG	FEATURES
                Dim Pieces As String() = Lines(Count).Split(CChar(vbTab))
                'FORM can be found identically in tanzil
                If Pieces(0).Chars(0) = "(" Then
                    If Not _FormDictionary.ContainsKey(Pieces(1)) Then
                        _FormDictionary.Add(Pieces(1), New List(Of Integer()))
                    End If
                    'TAG
                    If Not _TagDictionary.ContainsKey(Pieces(2)) Then
                        _TagDictionary.Add(Pieces(2), New Generic.Dictionary(Of String, List(Of Integer())))
                    End If
                    Dim Location As Integer() = New List(Of Integer)(Linq.Enumerable.Select(Pieces(0).TrimStart("("c).TrimEnd(")"c).Split(":"c), Function(Str As String) CInt(Str))).ToArray()
                    _FormDictionary.Item(Pieces(1)).Add(Location)
                    If Not _TagDictionary.Item(Pieces(2)).ContainsKey(Pieces(1)) Then
                        _TagDictionary.Item(Pieces(2)).Add(Pieces(1), New List(Of Integer()))
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
                            _WordDictionary.Add(Lem, New List(Of String))
                        End If
                        If _WordDictionary.Item(Lem).IndexOf(Pieces(1)) = -1 Then _WordDictionary.Item(Lem).Add(Pieces(1))
                    End If
                    If Array.Find(Parts, Function(Str As String) Str = "PREFIX") <> String.Empty Then
                        If Not _PreDictionary.ContainsKey(Pieces(1)) Then
                            _PreDictionary.Add(Pieces(1), New List(Of Integer()))
                        End If
                        _PreDictionary.Item(Pieces(1)).Add(Location)
                    ElseIf Array.Find(Parts, Function(Str As String) Str = "SUFFIX") <> String.Empty Then
                        If Not _SufDictionary.ContainsKey(Pieces(1)) Then
                            _SufDictionary.Add(Pieces(1), New List(Of Integer()))
                        End If
                        _SufDictionary.Item(Pieces(1)).Add(Location)
                    End If
                    'ROOT:
                    Dim Root As String = Array.Find(Parts, Function(Str As String) Str.StartsWith("ROOT:"))
                    If Root <> String.Empty Then
                        Root = Root.Replace("ROOT:", String.Empty)
                        If Not _RootDictionary.ContainsKey(Root) Then
                            _RootDictionary.Add(Root, New List(Of Integer()))
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
        For Count As Integer = 1 To CInt(If(bStation, TanzilReader.GetStationCount(), TanzilReader.GetPartCount()))
            PartUniqueArray(Count - 1) = New Generic.List(Of String)
            PartArray(Count - 1) = New Generic.List(Of String)
            Dim Node As Xml.Linq.XElement = CType(If(bStation, TanzilReader.GetStationByIndex(Count), TanzilReader.GetPartByIndex(Count)), Xml.Linq.XElement)
            Dim BaseChapter As Integer = CInt(Node.Attribute("sura").Value)
            Dim BaseVerse As Integer = CInt(Node.Attribute("aya").Value)
            Dim Chapter As Integer
            Dim Verse As Integer
            Node = CType(If(bStation, TanzilReader.GetStationByIndex(Count + 1), TanzilReader.GetPartByIndex(Count + 1)), Xml.Linq.XElement)
            If Node Is Nothing Then
                Chapter = TanzilReader.GetChapterCount()
                Verse = TanzilReader.GetVerseCount(Chapter)
            Else
                Chapter = CInt(Node.Attribute("sura").Value)
                Verse = CInt(Node.Attribute("aya").Value)
                TanzilReader.GetPreviousChapterVerse(Chapter, Verse)
            End If
            For SubCount As Integer = 0 To FreqArray.Length - 1
                Dim RefCount As Integer
                Dim UniCount As Integer = 0
                For RefCount = 0 To CachedData.WordDictionary(FreqArray(SubCount)).Count - 1
                    For FormCount As Integer = 0 To CachedData.FormDictionary(CStr(CachedData.WordDictionary(FreqArray(SubCount))(RefCount))).Count - 1
                        If (CachedData.FormDictionary(CStr(CachedData.WordDictionary(FreqArray(SubCount))(RefCount)))(FormCount)(0) = BaseChapter AndAlso _
                            CachedData.FormDictionary(CStr(CachedData.WordDictionary(FreqArray(SubCount))(RefCount)))(FormCount)(1) >= BaseVerse AndAlso _
                            (BaseChapter <> Chapter OrElse _
                            CachedData.FormDictionary(CStr(CachedData.WordDictionary(FreqArray(SubCount))(RefCount)))(FormCount)(1) <= Verse)) OrElse _
                            (CachedData.FormDictionary(CStr(CachedData.WordDictionary(FreqArray(SubCount))(RefCount)))(FormCount)(0) > BaseChapter AndAlso _
                            CachedData.FormDictionary(CStr(CachedData.WordDictionary(FreqArray(SubCount))(RefCount)))(FormCount)(0) < Chapter) OrElse _
                            (CachedData.FormDictionary(CStr(CachedData.WordDictionary(FreqArray(SubCount))(RefCount)))(FormCount)(0) = Chapter AndAlso
                            CachedData.FormDictionary(CStr(CachedData.WordDictionary(FreqArray(SubCount))(RefCount)))(FormCount)(1) <= Verse) Then
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
                        _LetterDictionary.Add(Verses(Count)(SubCount)(LetCount), New Dictionary(Of String, List(Of Integer())))
                    End If
                    If Not _LetterPreDictionary.ContainsKey(Verses(Count)(SubCount)(LetCount)) Then
                        _LetterPreDictionary.Add(Verses(Count)(SubCount)(LetCount), New Dictionary(Of String, List(Of Integer())))
                    End If
                    If Not _LetterSufDictionary.ContainsKey(Verses(Count)(SubCount)(LetCount)) Then
                        _LetterSufDictionary.Add(Verses(Count)(SubCount)(LetCount), New Dictionary(Of String, List(Of Integer())))
                    End If
                    Dim PrevIndex As Integer = Verses(Count)(SubCount).LastIndexOf(" "c, LetCount) + If(Verses(Count)(SubCount)(LetCount) = " ", 0, 1)
                    Dim NextIndex As Integer = Verses(Count)(SubCount).IndexOf(" "c, LetCount)
                    If NextIndex = -1 Then NextIndex = Verses(Count)(SubCount).Length
                    If Not _LetterDictionary.Item(Verses(Count)(SubCount)(LetCount)).ContainsKey(Verses(Count)(SubCount).Substring(PrevIndex, NextIndex - PrevIndex)) Then
                        _LetterDictionary.Item(Verses(Count)(SubCount)(LetCount)).Add(Verses(Count)(SubCount).Substring(PrevIndex, NextIndex - PrevIndex), New List(Of Integer()))
                    End If
                    _LetterDictionary.Item(Verses(Count)(SubCount)(LetCount))(Verses(Count)(SubCount).Substring(PrevIndex, NextIndex - PrevIndex)).Add(New Integer() {Count, SubCount, LetCount})
                    If Not _LetterPreDictionary.Item(Verses(Count)(SubCount)(LetCount)).ContainsKey(Verses(Count)(SubCount).Substring(PrevIndex, LetCount - PrevIndex)) Then
                        _LetterPreDictionary.Item(Verses(Count)(SubCount)(LetCount)).Add(Verses(Count)(SubCount).Substring(PrevIndex, LetCount - PrevIndex), New List(Of Integer()))
                    End If
                    _LetterPreDictionary.Item(Verses(Count)(SubCount)(LetCount))(Verses(Count)(SubCount).Substring(PrevIndex, LetCount - PrevIndex)).Add(New Integer() {Count, SubCount, LetCount})
                    If LetCount <> NextIndex Then
                        If Not _LetterSufDictionary.Item(Verses(Count)(SubCount)(LetCount)).ContainsKey(Verses(Count)(SubCount).Substring(LetCount + 1, NextIndex - LetCount - 1)) Then
                            _LetterSufDictionary.Item(Verses(Count)(SubCount)(LetCount)).Add(Verses(Count)(SubCount).Substring(LetCount + 1, NextIndex - LetCount - 1), New List(Of Integer()))
                        End If
                        _LetterSufDictionary.Item(Verses(Count)(SubCount)(LetCount))(Verses(Count)(SubCount).Substring(LetCount + 1, NextIndex - LetCount - 1)).Add(New Integer() {Count, SubCount, LetCount})
                    End If
                    If LetCount <> 0 AndAlso LetCount <> Verses(Count)(SubCount).Length - 1 AndAlso _
                        Char.IsWhiteSpace(Verses(Count)(SubCount)(LetCount - 1)) AndAlso Char.IsWhiteSpace(Verses(Count)(SubCount)(LetCount + 1)) Then
                        _TotalIsolatedLetters += 1
                        If Not _IsolatedLetterDictionary.ContainsKey(Verses(Count)(SubCount)(LetCount)) Then
                            _IsolatedLetterDictionary.Add(Verses(Count)(SubCount)(LetCount), New List(Of Integer()))
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
            Dim ChapterNode As Xml.Linq.XElement = TanzilReader.GetTextChapter(CachedData.XMLDocMain, Count + 1)
            For SubCount As Integer = 0 To Verses(Count).Length - 1
                If SubCount = 0 AndAlso Not TanzilReader.GetTextVerse(ChapterNode, SubCount + 1).Attribute("bismillah") Is Nothing Then
                    Arabic.DoErrorCheck(TanzilReader.GetTextVerse(ChapterNode, SubCount + 1).Attribute("bismillah").Value)
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
                'Debug.Print("Missing Phrase ID: " + IslamData.Abbreviations(Count).TranslationID)
            End If
        Next
        For Count = 0 To ArabicData.ArabicLetters.Length - 1
            Arabic.DoErrorCheck(Arabic.ArabicLetterSpelling(ArabicData.ArabicLetters(Count).Symbol, False, False, False))
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
                Dim fs As IO.Stream = PortableMethods.FileIO.LoadStream(PortableMethods.Settings.GetFilePath("metadata\islaminfo.xml"))
                Dim xs As Xml.Serialization.XmlSerializer = New Xml.Serialization.XmlSerializer(GetType(IslamData))
                _ObjIslamData = CType(xs.Deserialize(fs), IslamData)
                fs.Dispose()
            End If
            Return _ObjIslamData
        End Get
    End Property
    Public Shared ReadOnly Property XMLDocMain As Xml.Linq.XDocument
        Get
            If _XMLDocMain Is Nothing Then
                Dim Stream As IO.Stream = PortableMethods.FileIO.LoadStream(PortableMethods.Settings.GetFilePath("metadata\" + TanzilReader.QuranTextNames(0) + ".xml"))
                _XMLDocMain = Xml.Linq.XDocument.Load(Stream)
                Stream.Dispose()
            End If
            Return _XMLDocMain
        End Get
    End Property
    Public Shared ReadOnly Property XMLDocInfo As Xml.Linq.XDocument
        Get
            If _XMLDocInfo Is Nothing Then
                Dim Stream As IO.Stream = PortableMethods.FileIO.LoadStream(PortableMethods.Settings.GetFilePath("metadata\quran-data.xml"))
                _XMLDocInfo = Xml.Linq.XDocument.Load(Stream)
                Stream.Dispose()
            End If
            Return _XMLDocInfo
        End Get
    End Property
    Public Shared ReadOnly Property XMLDocInfos As Collections.Generic.List(Of Xml.Linq.XDocument)
        Get
            Dim Count As Integer
            If _XMLDocInfos Is Nothing Then
                _XMLDocInfos = New Collections.Generic.List(Of Xml.Linq.XDocument)
                For Count = 0 To CachedData.IslamData.Collections.Length - 1
                    Dim Stream As IO.Stream = PortableMethods.FileIO.LoadStream(PortableMethods.Settings.GetFilePath("metadata\" + CachedData.IslamData.Collections(Count).FileName + "-data.xml"))
                    _XMLDocInfos.Add(Xml.Linq.XDocument.Load(Stream))
                    Stream.Dispose()
                Next
            End If
            Return _XMLDocInfos
        End Get
    End Property
    Public Shared ReadOnly Property RootDictionary As Generic.Dictionary(Of String, List(Of Integer()))
        Get
            If _RootDictionary.Keys.Count = 0 Then GetMorphologicalData()
            Return _RootDictionary
        End Get
    End Property
    Public Shared ReadOnly Property FormDictionary As Generic.Dictionary(Of String, List(Of Integer()))
        Get
            If _FormDictionary.Keys.Count = 0 Then GetMorphologicalData()
            Return _FormDictionary
        End Get
    End Property
    Public Shared ReadOnly Property TagDictionary As Generic.Dictionary(Of String, Generic.Dictionary(Of String, List(Of Integer())))
        Get
            If _TagDictionary.Keys.Count = 0 Then GetMorphologicalData()
            Return _TagDictionary
        End Get
    End Property
    Public Shared ReadOnly Property WordDictionary As Generic.Dictionary(Of String, List(Of String))
        Get
            If _WordDictionary.Keys.Count = 0 Then GetMorphologicalData()
            Return _WordDictionary
        End Get
    End Property
    Public Shared ReadOnly Property RealWordDictionary As Generic.Dictionary(Of String, List(Of Integer()))
        Get
            If _RealWordDictionary.Keys.Count = 0 Then BuildQuranLetterIndex()
            Return _RealWordDictionary
        End Get
    End Property
    Public Shared ReadOnly Property LetterDictionary As Generic.Dictionary(Of Char, Generic.Dictionary(Of String, List(Of Integer())))
        Get
            If _LetterDictionary.Keys.Count = 0 Then BuildQuranLetterIndex()
            Return _LetterDictionary
        End Get
    End Property
    Public Shared ReadOnly Property LetterPreDictionary As Generic.Dictionary(Of Char, Generic.Dictionary(Of String, List(Of Integer())))
        Get
            If _LetterPreDictionary.Keys.Count = 0 Then BuildQuranLetterIndex()
            Return _LetterPreDictionary
        End Get
    End Property
    Public Shared ReadOnly Property LetterSufDictionary As Generic.Dictionary(Of Char, Generic.Dictionary(Of String, List(Of Integer())))
        Get
            If _LetterSufDictionary.Keys.Count = 0 Then BuildQuranLetterIndex()
            Return _LetterSufDictionary
        End Get
    End Property
    Public Shared ReadOnly Property PreDictionary As Generic.Dictionary(Of String, List(Of Integer()))
        Get
            If _PreDictionary.Keys.Count = 0 Then GetMorphologicalData()
            Return _PreDictionary
        End Get
    End Property
    Public Shared ReadOnly Property SufDictionary As Generic.Dictionary(Of String, List(Of Integer()))
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
    Public Shared ReadOnly Property IsolatedLetterDictionary As Generic.Dictionary(Of Char, List(Of Integer()))
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
Public Class ColorGenerator
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
    Public Shared Function ToRGB(N As Double) As Byte
        Dim Result As Double = N * 255.0
        If Result < 0 Then Return 0
        If Result > 255 Then Return 255
        Return CByte(Result)
    End Function
    Public Shared Function RGBToXYZ(clr As Integer) As XYZColor
        Dim r As Double = PivotRGB(Utility.ColorR(clr) / 255.0)
        Dim g As Double = PivotRGB(Utility.ColorG(clr) / 255.0)
        Dim b As Double = PivotRGB(Utility.ColorB(clr) / 255.0)
        Return New XYZColor With {.X = r * 0.4124 + g * 0.3576 + b * 0.1805, _
                                  .Y = r * 0.2126 + g * 0.7152 + b * 0.0722, _
                                  .Z = r * 0.0193 + g * 0.1192 + b * 0.9505}
    End Function
    Public Shared Function XYZToRGB(clr As XYZColor) As Integer
        Dim x As Double = clr.X / 100.0
        Dim y As Double = clr.Y / 100.0
        Dim z As Double = clr.Z / 100.0
        Dim r As Double = x * 3.2406 + y * -1.5372 + z * -0.4986
        Dim g As Double = x * -0.9689 + y * 1.8758 + z * 0.0415
        Dim b As Double = x * 0.0557 + y * -0.204 + z * 1.057
        r = If(r > 0.0031308, 1.055 * Math.Pow(r, 1 / 2.4) - 0.055, 12.92 * r)
        g = If(g > 0.0031308, 1.055 * Math.Pow(g, 1 / 2.4) - 0.055, 12.92 * g)
        b = If(b > 0.0031308, 1.055 * Math.Pow(b, 1 / 2.4) - 0.055, 12.92 * b)
        Return Utility.MakeArgb(&HFF, ToRGB(r), ToRGB(g), ToRGB(b))
    End Function
    Public Structure LABColor
        Public L As Double
        Public A As Double
        Public B As Double
    End Structure
    Public Shared Function PivotXYZ(N As Double) As Double
        Return If(N > Epsilon, Math.Pow(N, 1.0 / 3.0), (Kappa * N + 16) / 116)
    End Function
    Public Shared Function RGBToLAB(clr As Integer) As LABColor
        Dim XYZCol As XYZColor = RGBToXYZ(clr)
        Dim x As Double = PivotXYZ(XYZCol.X / WhiteReference.X)
        Dim y As Double = PivotXYZ(XYZCol.Y / WhiteReference.Y)
        Dim z As Double = PivotXYZ(XYZCol.Z / WhiteReference.Z)
        Return New LABColor With {.L = Math.Max(0, 116 * y - 16), .A = 500 * (x - y), .B = 200 * (y - z)}
    End Function
    Public Shared Function LABToRGB(clr As LABColor) As Integer
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
    Public Shared Function GenerateNDistinctColors(N As Integer, Threshold As Integer, Interleave As Integer) As Integer()
        'To best support individuals with colorblindness (deuteranopia or protanopia) keep a set to 0; vary only L and b.
        Dim LABColors As New List(Of LABColor)
        Dim LowThresholds As New List(Of Double)
        LABColors.Add(RGBToLAB(Utility.ColorFromName("Black"))) 'Start with pivot forecolor
        LowThresholds.Add(100)
        LABColors.Add(RGBToLAB(Utility.ColorFromName("White"))) 'Start with background color
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
        Dim Cols(N - 1) As Integer
        For Count As Integer = 0 To N - 1
            'Least Common Multiple LCM = a * b \ GCD(A, B)
            Cols((Count * Interleave) \ (N * Interleave \ GCD(N, Interleave)) + (Count * Interleave) Mod N) = LABToRGB(LABColors(Count))
        Next
        Return Cols
    End Function
End Class
Public Class DocBuilder
    Public Shared Function ColorizeList(Strs As String(), bArabic As Boolean) As RenderArray.RenderText()
        Dim Cols As Integer() = ColorGenerator.GenerateNDistinctColors(Strs.Length + 1, 15, 5)
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
        Dim Cols As Integer() = ColorGenerator.GenerateNDistinctColors(CurNum + 1, 15, 5)
        Dim Renderers As New List(Of RenderArray.RenderText)
        For Count As Integer = 0 To ParenPos.Count - 1
            If Count = ParenPos.Count - 1 Then
                Renderers.Add(New RenderArray.RenderText(RenderArray.RenderDisplayClass.eLTR, Str.Substring(Base)) With {.Clr = Cols(ParenPos(Count))})
            ElseIf ParenPos(Count) <> ParenPos(Count + 1) Then
                Renderers.Add(New RenderArray.RenderText(RenderArray.RenderDisplayClass.eLTR, Str.Substring(Base, Count - Base + 1)) With {.Clr = Cols(ParenPos(Count))})
                Base = Count + 1
                Renderers.Add(New RenderArray.RenderText(RenderArray.RenderDisplayClass.eLTR, CStr(ParenPos(Count))) With {.Clr = Utility.ColorFromName("DarkRed"), .Font = "Courier New"})
            End If
        Next
        Return Renderers.ToArray()
    End Function
    Public Shared Function GetRegExText(Str As String) As String
        Return System.Text.RegularExpressions.Regex.Replace(System.Text.RegularExpressions.Regex.Replace(Str, "\\u([0-9a-fA-F]{4})", Function(Match As System.Text.RegularExpressions.Match) ChrW(Integer.Parse(Match.Groups(1).Value, Globalization.NumberStyles.HexNumber))), "[\p{IsArabic}\p{IsArabicPresentationForms-A}\p{IsArabicPresentationForms-B}]+", ArabicData.LeftToRightEmbedding + "$&" + ArabicData.PopDirectionalFormatting)
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
    Shared _Abbrevs As Dictionary(Of String, IslamData.AbbrevWord)
    Public Shared ReadOnly Property Abbrevs As Dictionary(Of String, IslamData.AbbrevWord)
        Get
            If _Abbrevs Is Nothing Then
                _Abbrevs = New Dictionary(Of String, IslamData.AbbrevWord)
                For Count = 0 To CachedData.IslamData.Abbreviations.Length - 1
                    Dim AbbrevWord As IslamData.AbbrevWord = CachedData.IslamData.Abbreviations(Count)
                    If CachedData.IslamData.Abbreviations(Count).Text <> String.Empty Then
                        For Each Str As String In CachedData.IslamData.Abbreviations(Count).Text.Split("|"c)
                            _Abbrevs.Add(Str, AbbrevWord)
                        Next
                    End If
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
                'If Words Is Nothing OrElse Array.FindIndex(Words, Function(Word As IslamData.GrammarSet.GrammarNoun) S = Word.TranslationID) = -1 Then Debug.Print("Noun Subject ID Not Found: " + SelArr(Count))
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
                'If Words Is Nothing OrElse Array.FindIndex(Words, Function(Word As IslamData.GrammarSet.GrammarTransform) S = Word.TranslationID) = -1 Then Debug.Print("Transform Subject ID Not Found: " + SelArr(Count))
            Next
        ElseIf Strings.StartsWith("particle:") Then
            Dim SelArr As String() = Strings.Replace("particle:", String.Empty).Split(","c)
            For Count = 0 To SelArr.Length - 1
                'If Arabic.GetParticles(SelArr(Count)) Is Nothing Then Debug.Print("Particle ID Not Found: " + SelArr(Count))
            Next
        ElseIf Strings.StartsWith("noun:") Then
            Dim SelArr As String() = Strings.Replace("noun:", String.Empty).Split(","c)
            For Count = 0 To SelArr.Length - 1
                'If Arabic.GetCatNoun(SelArr(Count)) Is Nothing Then Debug.Print("Noun ID Not Found: " + SelArr(Count))
            Next
        ElseIf Strings.StartsWith("verb:") Then
            Dim SelArr As String() = Strings.Replace("verb:", String.Empty).Split(","c)
            For Count = 0 To SelArr.Length - 1
                'If Arabic.GetVerb(SelArr(Count)) Is Nothing Then Debug.Print("Verb ID Not Found: " + SelArr(Count))
            Next
        ElseIf Strings.StartsWith("transform:") Then
            Dim SelArr As String() = Strings.Replace("transform:", String.Empty).Split(","c)
            For Count = 0 To SelArr.Length - 1
                'If Arabic.GetTransform(SelArr(Count)) Is Nothing Then Debug.Print("Transform ID Not Found: " + SelArr(Count))
            Next
        ElseIf Strings.StartsWith("word:") Then
            Dim SelArr As String() = Strings.Replace("word:", String.Empty).Split(","c)
            For Count = 0 To SelArr.Length - 1
                'If Not Arabic.GetCatWord(SelArr(Count)).HasValue Then Debug.Print("Word ID Not Found: " + SelArr(Count))
            Next
        ElseIf Strings.StartsWith("phrase:") Then
            Dim SelArr As String() = Strings.Replace("phrase:", String.Empty).Split(","c)
            For Count = 0 To SelArr.Length - 1
                'If Not Phrases.GetPhraseCat(SelArr(Count)).HasValue Then Debug.Print("Phrase ID Not Found: " + SelArr(Count))
            Next
        ElseIf Strings.StartsWith("list:") Then
            Dim SelArr As String() = Strings.Replace("list:", String.Empty).Split(","c)
            For Count = 0 To SelArr.Length - 1
                'If GetListCats({SelArr(Count)}).Length = 0 Then Debug.Print("List ID Not Found: " + SelArr(Count))
            Next
        ElseIf Abbrevs.ContainsKey(Strings) Then
        ElseIf Strings.StartsWith("reference:") Then
        ElseIf Strings.StartsWith("arabic:") Then
        ElseIf Strings.StartsWith("text:") Then
        Else
            'Debug.Print("Unknown tag:" + Strings)
        End If
    End Sub
    Public Shared Function GetListCategories() As String()
        Return New List(Of String)(Linq.Enumerable.Select(CachedData.IslamData.Lists, Function(Convert As IslamData.ListCategory) Utility.LoadResourceString("IslamInfo_" + Convert.Title))).ToArray()
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
End Class
Public Class TanzilReader
    Public Shared Function GetDivisionTypes() As String()
        Return New List(Of String)(Linq.Enumerable.Select(CachedData.IslamData.QuranDivisions, Function(Convert As IslamData.QuranDivision) Utility.LoadResourceString("IslamInfo_" + Convert.Description))).ToArray()
    End Function
    Public Shared Function GetTranslationList() As Array()
        Return New List(Of String())(Linq.Enumerable.Select(CachedData.IslamData.Translations.TranslationList, Function(Convert As IslamData.TranslationsInfo.TranslationInfo) New String() {Utility.LoadResourceString("lang_local" + Languages.GetLanguageInfoByCode(Convert.FileName.Substring(0, CInt(If(Convert.FileName.IndexOf("-") <> -1, Convert.FileName.IndexOf("-"), Convert.FileName.IndexOf("."))))).Code) + ": " + Convert.Name, Convert.FileName})).ToArray()
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

    Public Shared Function GetWordPartitions() As String()
        Dim Parts As New Generic.List(Of String) From {Utility.LoadResourceString("IslamInfo_Letters"), Utility.LoadResourceString("IslamInfo_Words"), Utility.LoadResourceString("IslamInfo_UniqueWords"), Utility.LoadResourceString("IslamInfo_UniqueWordsPerPart"), Utility.LoadResourceString("IslamInfo_WordsPerPart"), Utility.LoadResourceString("IslamInfo_UniqueWordsPerStation"), Utility.LoadResourceString("IslamInfo_WordsPerStation"), Utility.LoadResourceString("IslamInfo_IsolatedLetters"), Utility.LoadResourceString("IslamInfo_LetterPatterns"), Utility.LoadResourceString("IslamInfo_LetterPatterns"), Utility.LoadResourceString("IslamInfo_LetterPatterns"), Utility.LoadResourceString("IslamInfo_Prefix"), Utility.LoadResourceString("IslamInfo_Suffix")}
        Parts.AddRange(Linq.Enumerable.Select(CachedData.IslamData.PartsOfSpeech, Function(POS As IslamData.PartOfSpeechInfo) Utility.LoadResourceString("IslamInfo_" + POS.Id)))
        Parts.AddRange(Linq.Enumerable.Select(CachedData.RecitationSymbols, Function(Sym As String) ArabicData.GetUnicodeName(Sym.Chars(0))))
        Parts.AddRange(Linq.Enumerable.Select(CachedData.RecitationSymbols, Function(Sym As String) "Prefix of " + ArabicData.GetUnicodeName(Sym.Chars(0))))
        Parts.AddRange(Linq.Enumerable.Select(CachedData.RecitationSymbols, Function(Sym As String) "Suffix of " + ArabicData.GetUnicodeName(Sym.Chars(0))))
        Return Parts.ToArray()
    End Function
    Public Shared Function GetQuranWordTotalNumber() As Integer
        Dim Total As Integer
        For Each Key As String In CachedData.WordDictionary.Keys
            Total = Total + CachedData.WordDictionary.Item(Key).Count
        Next
        Return Total
    End Function
    Public Shared Function GetQuranWordTotal(Strings As String) As String
        Dim Index As Integer
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
    Public Shared Function GetQuranWordFrequency(SchemeType As ArabicData.TranslitScheme, Scheme As String, Strings As String) As Array()
        Dim Output As New List(Of Object)
        Dim Total As Integer = 0
        Dim All As Double
        Dim Index As Integer
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
            Dim Dict As Generic.Dictionary(Of String, List(Of Integer()))
            If Index = 1 Then
                Dict = New Dictionary(Of String, List(Of Integer()))
                For Each KV As KeyValuePair(Of String, List(Of String)) In CachedData.WordDictionary
                    Dim Str As String = KV.Key + vbCrLf + String.Join(vbCrLf, CType(KV.Value.ToArray(), String()))
                    Dict.Add(Str, New List(Of Integer()))
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
            Dim W4WLines As String() = Utility.ReadAllLines(PortableMethods.Settings.GetFilePath("metadata\en.w4w.corpus.txt"))
            For Count As Integer = 0 To FreqArray.Length - 1
                Dim TranslationDict As New Dictionary(Of String, List(Of Integer()))
                For WordCount As Integer = 0 To Dict.Item(FreqArray(Count)).Count - 1
                    Dim CheckStr As String = TanzilReader.GetW4WTranslationVerse(W4WLines, Dict.Item(FreqArray(Count))(WordCount)(0), Dict.Item(FreqArray(Count))(WordCount)(1), Dict.Item(FreqArray(Count))(WordCount)(2) - 1)
                    If Not TranslationDict.ContainsKey(CheckStr) Then
                        TranslationDict.Add(CheckStr, New List(Of Integer()))
                    End If
                    TranslationDict(CheckStr).Add(Dict.Item(FreqArray(Count))(WordCount))
                Next
                Dim TranslationArray(TranslationDict.Keys.Count - 1) As String
                TranslationDict.Keys.CopyTo(TranslationArray, 0)
                For WordCount As Integer = 0 To TranslationArray.Length - 1
                    TranslationArray(WordCount) += vbCrLf + "(" + String.Join(",", Linq.Enumerable.Select(TranslationDict(TranslationArray(WordCount)).ToArray(), Function(Indexes As Integer()) String.Join(":", Linq.Enumerable.Select(Indexes, Function(Idx As Integer) CStr(Idx))))) + ")"
                Next
                Total += Dict.Item(FreqArray(Count)).Count
                Output.Add(New String() {Arabic.TransliterateFromBuckwalter(FreqArray(Count)), Arabic.TransliterateToScheme(Arabic.TransliterateFromBuckwalter(FreqArray(Count)), If(SchemeType = ArabicData.TranslitScheme.RuleBased, ArabicData.TranslitScheme.LearningMode, SchemeType), Scheme, CachedData.RuleMetas("Normal")), String.Join(vbCrLf, TranslationArray), CStr(Dict.Item(FreqArray(Count)).Count), (CDbl(Dict.Item(FreqArray(Count)).Count) * 100 / All).ToString("n2"), (CDbl(Total) * 100 / All).ToString("n2")})
            Next
        ElseIf Index = 3 Or Index = 4 Or Index = 5 Or Index = 6 Then
            Total = 0
            Dim DivArray As Collections.Generic.List(Of String)()
            If Index = 3 Or Index = 5 Then
                DivArray = If(Index = 5, CachedData.StationUniqueArray, CachedData.PartUniqueArray)
                All = If(Index = 5, CachedData.TotalUniqueWordsInStations, CachedData.TotalUniqueWordsInParts)
                For Count As Integer = 0 To CInt(If(Index = 5, TanzilReader.GetStationCount(), TanzilReader.GetPartCount())) - 1
                    Total += DivArray(Count).Count
                    Output.Add(New String() {ArabicData.LeftToRightEmbedding + CStr(Count + 1) + ArabicData.PopDirectionalFormatting, String.Empty, String.Empty, CStr(DivArray(Count).Count), (CDbl(DivArray(Count).Count) * 100 / All).ToString("n2"), (CDbl(Total) * 100 / All).ToString("n2")})
                Next
            ElseIf Index = 4 Or Index = 6 Then
                DivArray = If(Index = 6, CachedData.StationUniqueArray, CachedData.PartUniqueArray)
                All = If(Index = 6, CachedData.TotalWordsInStations, CachedData.TotalWordsInParts)
                For Count As Integer = 0 To CInt(If(Index = 6, TanzilReader.GetStationCount(), TanzilReader.GetPartCount())) - 1
                    Total += DivArray(Count).Count
                    Output.Add(New String() {ArabicData.LeftToRightEmbedding + CStr(Count + 1) + ArabicData.PopDirectionalFormatting, String.Empty, String.Empty, CStr(DivArray(Count).Count), (CDbl(DivArray(Count).Count) * 100 / All).ToString("n2"), (CDbl(Total) * 100 / All).ToString("n2")})
                Next
            End If
        ElseIf Index = 8 Then
            Output.AddRange(Linq.Enumerable.Select(GetQuranLetterPatterns(), Function(Str As String) {Str, String.Empty, String.Empty, String.Empty, String.Empty, String.Empty}))
        ElseIf Index = 9 Then
            Output.AddRange(Linq.Enumerable.Select(PatternAnalysis(), Function(Str As String) {Str, String.Empty, String.Empty, String.Empty, String.Empty, String.Empty}))
        ElseIf Index = 10 Then
            Output.AddRange(Linq.Enumerable.Select(GetQuranHamzaMaddDoubleLetterPatterns(), Function(Str As String) {Str, String.Empty, String.Empty, String.Empty, String.Empty, String.Empty}))
        End If
        Return CType(Output.ToArray(), Array())
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
                    Str += Arabic.TransliterateToScheme(KV.Key, ArabicData.TranslitScheme.Literal, String.Empty, Nothing) + DumpRecDictionary(CType(KV.Value(0), Dictionary(Of String, Object())), Post, Depth + 1, Dual) + If(Dual AndAlso CType(KV.Value(1), Dictionary(Of String, Object())).Keys.Count <> 0, "!(" + DumpRecDictionary(CType(KV.Value(1), Dictionary(Of String, Object())), Not Post, Depth + 1, False) + ")!", String.Empty)
                Else
                    Str = DumpRecDictionary(CType(KV.Value(0), Dictionary(Of String, Object())), Post, Depth + 1, Dual) + If(Dual AndAlso CType(KV.Value(1), Dictionary(Of String, Object())).Keys.Count <> 0, "!(" + DumpRecDictionary(CType(KV.Value(1), Dictionary(Of String, Object())), Not Post, Depth + 1, False) + ")!", String.Empty) + Arabic.TransliterateToScheme(KV.Key, ArabicData.TranslitScheme.Literal, String.Empty, Nothing) + Str
                End If
            ElseIf Post Then
                Str += Arabic.TransliterateToScheme(KV.Key, ArabicData.TranslitScheme.Literal, String.Empty, Nothing) + If(CType(KV.Value(0), Dictionary(Of String, Object())).Keys.Count = 0, String.Empty, vbCrLf + New String(" "c, Depth * 4) + "(" + DumpRecDictionary(CType(KV.Value(0), Dictionary(Of String, Object())), Post, Depth + 1, Dual) + ")" + vbCrLf + New String(" "c, Depth * 4)) + If(Dual AndAlso CType(KV.Value(1), Dictionary(Of String, Object())).Keys.Count <> 0, vbCrLf + New String(" "c, Depth * 4) + "!(" + DumpRecDictionary(CType(KV.Value(1), Dictionary(Of String, Object())), Not Post, Depth + 1, False) + ")!" + vbCrLf + New String(" "c, Depth * 4), String.Empty)
            Else
                Str = If(CType(KV.Value(0), Dictionary(Of String, Object())).Keys.Count = 0, String.Empty, vbCrLf + New String(" "c, Depth * 4) + "(" + DumpRecDictionary(CType(KV.Value(0), Dictionary(Of String, Object())), Post, Depth + 1, Dual) + ")" + vbCrLf + New String(" "c, Depth * 4)) + If(Dual AndAlso CType(KV.Value(1), Dictionary(Of String, Object())).Keys.Count <> 0, vbCrLf + New String(" "c, Depth * 4) + "!(" + DumpRecDictionary(CType(KV.Value(1), Dictionary(Of String, Object())), Not Post, Depth + 1, False) + ")!" + vbCrLf + New String(" "c, Depth * 4), String.Empty) + Arabic.TransliterateToScheme(KV.Key, ArabicData.TranslitScheme.Literal, String.Empty, Nothing) + Str
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
            Strings(LetCount) = ArabicData.LeftToRightEmbedding + DumpRecDictionary(PreDict, False, 0, True) + "\" + Arabic.TransliterateToScheme(CachedData.RecitationSymbols(LetCount), ArabicData.TranslitScheme.Literal, String.Empty, Nothing) + "\" + DumpRecDictionary(PostDict, True, 0, True) + ArabicData.PopDirectionalFormatting
        Next
        Return Strings
    End Function
    Public Shared Function GetQuranHamzaMaddDoubleLetterPatterns() As String()
        Dim CurPat As String = ArabicData.MakeRegMultiEx(CachedData.RecitationSymbols)
        Dim Prefixes As New Dictionary(Of String, List(Of String))
        Dim Suffixes As New Dictionary(Of String, List(Of String))
        Dim PreMidSuf As New Dictionary(Of String, Object()) 'Prefix indexed
        Dim SufMidPre As New Dictionary(Of String, Object()) 'Suffix indexed
        For Each Key As String In CachedData.FormDictionary.Keys
            Dim Matches As System.Text.RegularExpressions.MatchCollection = System.Text.RegularExpressions.Regex.Matches(Arabic.TransliterateFromBuckwalter(Key), CurPat)
            For Count As Integer = 0 To Matches.Count - 1
                If Matches(Count).Index = 0 Then
                    For SubCount As Integer = 0 To CachedData.FormDictionary(Key).Count - 1
                        Dim Loc(3) As Integer
                        CachedData.FormDictionary(Key)(SubCount).CopyTo(Loc, 0)
                        If (Not CBool(CachedData.LocDictionary(String.Join(":", Loc))(1)) And Not CBool(CachedData.LocDictionary(String.Join(":", Loc))(2))) Or CStr(CachedData.LocDictionary(String.Join(":", Loc))(3)) = "DET" Or (Matches(Count).Value = ArabicData.ArabicLetterAlefWasla And CStr(CachedData.LocDictionary(String.Join(":", Loc))(3)) = "PN") Then
                            Dim Pre As String = String.Empty
                            For LocCount = 1 To CachedData.FormDictionary(Key)(SubCount)(3) - 1
                                Loc(3) = LocCount
                                Pre += CStr(CachedData.LocDictionary(String.Join(":", Loc))(0))
                            Next
                            If Pre <> String.Empty Then
                                Loc(3) = CachedData.FormDictionary(Key)(SubCount)(3)
                                If CStr(CachedData.LocDictionary(String.Join(":", Loc))(3)) = "DET" Or (Matches(Count).Value = ArabicData.ArabicLetterAlefWasla And CStr(CachedData.LocDictionary(String.Join(":", Loc))(3)) = "PN") Then
                                    If Not Prefixes.ContainsKey("Al+") Then Prefixes.Add("Al+", New List(Of String))
                                    Prefixes("Al+").Add(New String(New List(Of Char)(Linq.Enumerable.Reverse(Pre.ToCharArray())).ToArray()))
                                Else
                                    If Not Prefixes.ContainsKey(Matches(Count).Value) Then Prefixes.Add(Matches(Count).Value, New List(Of String))
                                    Prefixes(Matches(Count).Value).Add(New String(New List(Of Char)(Linq.Enumerable.Reverse(Pre.ToCharArray())).ToArray()))
                                End If
                                If Matches(Count).Value <> ArabicData.ArabicLetterAlefWasla And CStr(CachedData.LocDictionary(String.Join(":", Loc))(3)) <> "DET" Then
                                    If Not Prefixes.ContainsKey("!" + ArabicData.ArabicLetterAlefWasla) Then Prefixes.Add("!" + ArabicData.ArabicLetterAlefWasla, New List(Of String))
                                    Prefixes("!" + ArabicData.ArabicLetterAlefWasla).Add(New String(New List(Of Char)(Linq.Enumerable.Reverse(Pre.ToCharArray())).ToArray()))
                                End If
                            End If
                        End If
                    Next
                ElseIf Matches(Count).Index = Key.Length - 1 Then
                    For SubCount As Integer = 0 To CachedData.FormDictionary(Key).Count - 1
                        Dim Loc(3) As Integer
                        CachedData.FormDictionary(Key)(SubCount).CopyTo(Loc, 0)
                        If (Not CBool(CachedData.LocDictionary(String.Join(":", Loc))(1)) And Not CBool(CachedData.LocDictionary(String.Join(":", Loc))(2))) Then
                            Dim Sup As String = String.Empty
                            Dim LocCount As Integer = CachedData.FormDictionary(Key)(SubCount)(3) + 1
                            Do
                                Loc(3) = LocCount
                                If Not CachedData.LocDictionary.ContainsKey(String.Join(":", Loc)) Then Exit Do
                                Sup += CStr(CachedData.LocDictionary(String.Join(":", Loc))(0))
                                LocCount += 1
                            Loop While True
                            If Sup <> String.Empty Then
                                If Not Suffixes.ContainsKey(Matches(Count).Value) Then Suffixes.Add(Matches(Count).Value, New List(Of String))
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
                        CachedData.FormDictionary(Key)(SubCount).CopyTo(Loc, 0)
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
                            For LocCount = 1 To CachedData.FormDictionary(Key)(SubCount)(3) - 1
                                Loc(3) = LocCount
                                PreCheck += Arabic.TransliterateFromBuckwalter(CStr(CachedData.LocDictionary(String.Join(":", Loc))(0)))
                            Next
                            If PreCheck <> String.Empty Then Pre = ";" + PreCheck + ";" + Pre
                        End If
                        Dim Sup As String = Arabic.TransliterateFromBuckwalter(Key.Substring(Matches(Count).Index + 1))
                        LocCount = CachedData.FormDictionary(Key)(SubCount)(3) + 1
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
                    Pres += New String(New List(Of Char)(Linq.Enumerable.Reverse(CStr(Prefixes(Key)(Count)).ToCharArray())).ToArray())
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
            Strings(CurNum) = ArabicData.LeftToRightEmbedding + "Key:" + Key + "\" + Arabic.TransliterateToScheme(Key, ArabicData.TranslitScheme.Literal, String.Empty, Nothing) + "\" + DumpRecDictionary(CType(PreMidSuf(Key)(0), Dictionary(Of String, Object())), False, 0, True) + "\" + ArabicData.PopDirectionalFormatting
            CurNum += 1
        Next
        For Each Key As String In SufMidPre.Keys
            Strings(CurNum) = ArabicData.LeftToRightEmbedding + "Key:" + Key + "\" + Arabic.TransliterateToScheme(Key, ArabicData.TranslitScheme.Literal, String.Empty, Nothing) + "\" + DumpRecDictionary(CType(SufMidPre(Key)(0), Dictionary(Of String, Object())), False, 0, True) + "\" + ArabicData.PopDirectionalFormatting
            CurNum += 1
        Next
        Return Strings
    End Function
    Public Shared Function GetQuranLetterPatterns() As String()
        Dim RecSymbols As String = String.Join(String.Empty, Linq.Enumerable.Select(CachedData.RecitationSpecialSymbols, Function(C As String) C))
        Dim LtrSymbols As String = String.Join(String.Empty, Linq.Enumerable.Select(CachedData.RecitationLetters, Function(C As String) C))
        Dim DiaSymbols As String = String.Join(String.Empty, Linq.Enumerable.Select(CachedData.RecitationDiacritics, Function(C As String) C))
        Dim StartWordMultiOnly As New Generic.Dictionary(Of String, String)
        Dim EndWordMultiOnly As New Generic.Dictionary(Of String, String)
        Dim MiddleWordMultiOnly As New Generic.Dictionary(Of String, String)
        Dim StartWordOnly As String = String.Join(String.Empty, Linq.Enumerable.Select(CachedData.RecitationLettersDiacritics, Function(C As String) C))
        Dim NotStartWord As String = String.Join(String.Empty, Linq.Enumerable.Select(CachedData.RecitationLettersDiacritics, Function(C As String) C))
        Dim EndWordOnly As String = String.Join(String.Empty, Linq.Enumerable.Select(CachedData.RecitationLettersDiacritics, Function(C As String) C))
        Dim NotEndWord As String = String.Join(String.Empty, Linq.Enumerable.Select(CachedData.RecitationLettersDiacritics, Function(C As String) C))
        Dim EndWordOnlyNoDia As String = String.Join(String.Empty, Linq.Enumerable.Select(CachedData.RecitationLetters, Function(C As String) C))
        Dim NotEndWordNoDia As String = String.Join(String.Empty, Linq.Enumerable.Select(CachedData.RecitationLetters, Function(C As String) C))
        Dim MiddleWordOnlyNoDia As String = String.Join(String.Empty, Linq.Enumerable.Select(CachedData.RecitationLetters, Function(C As String) C))
        Dim NotMiddleWordNoDia As String = String.Join(String.Empty, Linq.Enumerable.Select(CachedData.RecitationLetters, Function(C As String) C))
        Dim MiddleWordOnly As String = String.Join(String.Empty, Linq.Enumerable.Select(CachedData.RecitationLettersDiacritics, Function(C As String) C))
        Dim NotMiddleWord As String = String.Join(String.Empty, Linq.Enumerable.Select(CachedData.RecitationLettersDiacritics, Function(C As String) C))
        Dim DiaStartWordOnly As String = String.Join(String.Empty, Linq.Enumerable.Select(CachedData.RecitationDiacritics, Function(C As String) C))
        Dim DiaNotStartWord As String = String.Join(String.Empty, Linq.Enumerable.Select(CachedData.RecitationDiacritics, Function(C As String) C))
        Dim DiaEndWordOnly As String = String.Join(String.Empty, Linq.Enumerable.Select(CachedData.RecitationDiacritics, Function(C As String) C))
        Dim DiaNotEndWord As String = String.Join(String.Empty, Linq.Enumerable.Select(CachedData.RecitationDiacritics, Function(C As String) C))
        Dim DiaMiddleWordOnly As String = String.Join(String.Empty, Linq.Enumerable.Select(CachedData.RecitationDiacritics, Function(C As String) C))
        Dim DiaNotMiddleWord As String = String.Join(String.Empty, Linq.Enumerable.Select(CachedData.RecitationDiacritics, Function(C As String) C))
        Dim Combos As String() = String.Join("|", Linq.Enumerable.Select(CachedData.RecitationLettersDiacritics, Function(C As String) String.Join("|", Linq.Enumerable.Select(CachedData.RecitationLettersDiacritics, Function(Nxt As String) C + Nxt)))).Split("|"c)
        Dim DiaCombos As String() = String.Join("|", Linq.Enumerable.Select(CachedData.RecitationDiacritics, Function(C As String) String.Join("|", Linq.Enumerable.Select(CachedData.RecitationDiacritics, Function(Nxt As String) C + Nxt)))).Split("|"c)
        Dim LetCombos As String() = String.Join("|", Linq.Enumerable.Select(CachedData.RecitationLetters, Function(C As String) String.Join("|", Linq.Enumerable.Select(CachedData.RecitationLetters, Function(Nxt As String) C + Nxt)))).Split("|"c)
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
            For Each S As String In KeyArray
                If Str.LastIndexOf(S) <> -1 AndAlso Str.LastIndexOf(S) <> 0 Then EndWordMultiOnly.Remove(S)
            Next
            ReDim KeyArray(StartWordMultiOnly.Keys.Count - 1)
            StartWordMultiOnly.Keys.CopyTo(KeyArray, 0)
            For Each S As String In KeyArray
                If Str.LastIndexOf(S) <> -1 AndAlso Str.LastIndexOf(S) <> 0 Then StartWordMultiOnly.Remove(S)
            Next
            ReDim KeyArray(MiddleWordMultiOnly.Keys.Count - 1)
            MiddleWordMultiOnly.Keys.CopyTo(KeyArray, 0)
            For Each S As String In KeyArray
                If Str.LastIndexOf(S) <> -1 AndAlso Str.LastIndexOf(S) <> 0 Then MiddleWordMultiOnly.Remove(S)
            Next
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
        For Each Str As String In Combos
            If Dict.ContainsKey(Str.Chars(0)) Then
                Dict.Item(Str.Chars(0)) = Dict.Item(Str.Chars(0)) + Str.Chars(1)
            Else
                Dict.Add(Str.Chars(0), Str.Chars(1))
            End If
        Next
        Dim Val As String = ArabicData.LeftToRightEmbedding + "Combinations: " + ArabicData.PopDirectionalFormatting
        For Each Key As Char In Dict.Keys
            If Dict.Item(Key).Length > (DiaSymbols.Length + LtrSymbols.Length) / 2 Then
                Val += ArabicData.FixStartingCombiningSymbol(Key) + ArabicData.LeftToRightEmbedding + " [" + ArabicData.PopDirectionalFormatting + String.Join(" ", Linq.Enumerable.Select(New String(Array.FindAll((DiaSymbols + LtrSymbols).ToCharArray(), Function(C As Char) Not Dict.Item(Key).Contains(C))).ToCharArray(), Function(C As Char) ArabicData.FixStartingCombiningSymbol(CStr(C)))) + ArabicData.LeftToRightEmbedding + " ]" + ArabicData.PopDirectionalFormatting + vbTab
            Else
                Val += ArabicData.FixStartingCombiningSymbol(Key) + ArabicData.LeftToRightEmbedding + " ! [ " + ArabicData.PopDirectionalFormatting + String.Join(" ", Linq.Enumerable.Select(Dict.Item(Key).ToCharArray(), Function(C As Char) ArabicData.FixStartingCombiningSymbol(CStr(C)))) + ArabicData.LeftToRightEmbedding + " ]" + ArabicData.PopDirectionalFormatting + vbTab
            End If
        Next
        Dim RevDict As New Generic.Dictionary(Of Char, String)
        For Each Str As String In Combos
            If RevDict.ContainsKey(Str.Chars(1)) Then
                RevDict.Item(Str.Chars(1)) = RevDict.Item(Str.Chars(1)) + Str.Chars(0)
            Else
                RevDict.Add(Str.Chars(1), Str.Chars(0))
            End If
        Next
        Dim RevVal As String = ArabicData.LeftToRightEmbedding + "Reverse Combinations: " + ArabicData.PopDirectionalFormatting
        For Each Key As Char In RevDict.Keys
            If RevDict.Item(Key).Length > (DiaSymbols.Length + LtrSymbols.Length) / 2 Then
                RevVal += ArabicData.LeftToRightEmbedding + "[" + ArabicData.PopDirectionalFormatting + String.Join(" ", Linq.Enumerable.Select(New String(Array.FindAll((DiaSymbols + LtrSymbols).ToCharArray(), Function(C As Char) Not RevDict.Item(Key).Contains(C))).ToCharArray(), Function(C As Char) ArabicData.FixStartingCombiningSymbol(CStr(C)))) + ArabicData.LeftToRightEmbedding + " ] " + ArabicData.PopDirectionalFormatting + ArabicData.FixStartingCombiningSymbol(Key) + vbTab
            Else
                RevVal += ArabicData.LeftToRightEmbedding + "! [ " + ArabicData.PopDirectionalFormatting + String.Join(" ", Linq.Enumerable.Select(RevDict.Item(Key).ToCharArray(), Function(C As Char) ArabicData.FixStartingCombiningSymbol(CStr(C)))) + ArabicData.LeftToRightEmbedding + " ] " + ArabicData.PopDirectionalFormatting + ArabicData.FixStartingCombiningSymbol(Key) + vbTab
            End If
        Next
        Dim DiaDict As New Generic.Dictionary(Of Char, String)
        For Each Str As String In DiaCombos
            If DiaDict.ContainsKey(Str.Chars(0)) Then
                DiaDict.Item(Str.Chars(0)) = DiaDict.Item(Str.Chars(0)) + Str.Chars(1)
            Else
                DiaDict.Add(Str.Chars(0), Str.Chars(1))
            End If
        Next
        Dim DiaVal As String = ArabicData.LeftToRightEmbedding + "Diacritic Only Combinations: " + ArabicData.PopDirectionalFormatting
        For Each Key As Char In DiaDict.Keys
            If DiaDict.Item(Key).Length > DiaSymbols.Length / 2 Then
                DiaVal += ArabicData.FixStartingCombiningSymbol(Key) + ArabicData.LeftToRightEmbedding + " [" + ArabicData.PopDirectionalFormatting + String.Join(" ", Linq.Enumerable.Select(New String(Array.FindAll(DiaSymbols.ToCharArray(), Function(C As Char) Not DiaDict.Item(Key).Contains(C))).ToCharArray(), Function(C As Char) ArabicData.FixStartingCombiningSymbol(CStr(C)))) + ArabicData.LeftToRightEmbedding + " ]" + ArabicData.PopDirectionalFormatting + vbTab
            Else
                DiaVal += ArabicData.FixStartingCombiningSymbol(Key) + ArabicData.LeftToRightEmbedding + " ! [ " + ArabicData.PopDirectionalFormatting + String.Join(" ", Linq.Enumerable.Select(DiaDict.Item(Key).ToCharArray(), Function(C As Char) ArabicData.FixStartingCombiningSymbol(CStr(C)))) + ArabicData.LeftToRightEmbedding + " ]" + ArabicData.PopDirectionalFormatting + vbTab
            End If
        Next
        Dim LetDict As New Generic.Dictionary(Of Char, String)
        For Each Str As String In LetCombos
            If LetDict.ContainsKey(Str.Chars(0)) Then
                LetDict.Item(Str.Chars(0)) = LetDict.Item(Str.Chars(0)) + Str.Chars(1)
            Else
                LetDict.Add(Str.Chars(0), Str.Chars(1))
            End If
        Next
        Dim LetVal As String = ArabicData.LeftToRightEmbedding + "Letter Only Combinations: " + ArabicData.PopDirectionalFormatting
        For Each Key As Char In LetDict.Keys
            If LetDict.Item(Key).Length > LtrSymbols.Length / 2 Then
                LetVal += ArabicData.FixStartingCombiningSymbol(Key) + ArabicData.LeftToRightEmbedding + " [" + ArabicData.PopDirectionalFormatting + String.Join(" ", Linq.Enumerable.Select(New String(Array.FindAll(LtrSymbols.ToCharArray(), Function(C As Char) Not LetDict.Item(Key).Contains(C))).ToCharArray(), Function(C As Char) ArabicData.FixStartingCombiningSymbol(CStr(C)))) + ArabicData.LeftToRightEmbedding + " ]" + ArabicData.PopDirectionalFormatting + vbTab
            Else
                LetVal += ArabicData.FixStartingCombiningSymbol(Key) + ArabicData.LeftToRightEmbedding + " ! [ " + ArabicData.PopDirectionalFormatting + String.Join(" ", Linq.Enumerable.Select(LetDict.Item(Key).ToCharArray(), Function(C As Char) ArabicData.FixStartingCombiningSymbol(CStr(C)))) + ArabicData.LeftToRightEmbedding + " ]" + ArabicData.PopDirectionalFormatting + vbTab
            End If
        Next
        Dim LetRevDict As New Generic.Dictionary(Of Char, String)
        For Each Str As String In LetCombos
            If LetRevDict.ContainsKey(Str.Chars(1)) Then
                LetRevDict.Item(Str.Chars(1)) = LetRevDict.Item(Str.Chars(1)) + Str.Chars(0)
            Else
                LetRevDict.Add(Str.Chars(1), Str.Chars(0))
            End If
        Next
        Dim LetRevVal As String = ArabicData.LeftToRightEmbedding + "Reverse Letter Only Combinations: " + ArabicData.PopDirectionalFormatting
        For Each Key As Char In LetRevDict.Keys
            If LetRevDict.Item(Key).Length > LtrSymbols.Length / 2 Then
                LetRevVal += ArabicData.LeftToRightEmbedding + "[" + ArabicData.PopDirectionalFormatting + String.Join(" ", Linq.Enumerable.Select(New String(Array.FindAll(LtrSymbols.ToCharArray(), Function(C As Char) Not LetRevDict.Item(Key).Contains(C))).ToCharArray(), Function(C As Char) ArabicData.FixStartingCombiningSymbol(CStr(C)))) + ArabicData.LeftToRightEmbedding + " ] " + ArabicData.PopDirectionalFormatting + ArabicData.FixStartingCombiningSymbol(Key) + vbTab
            Else
                LetRevVal += ArabicData.LeftToRightEmbedding + "! [ " + ArabicData.PopDirectionalFormatting + String.Join(" ", Linq.Enumerable.Select(LetRevDict.Item(Key).ToCharArray(), Function(C As Char) ArabicData.FixStartingCombiningSymbol(CStr(C)))) + ArabicData.LeftToRightEmbedding + " ] " + ArabicData.PopDirectionalFormatting + ArabicData.FixStartingCombiningSymbol(Key) + vbTab
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
                ArabicData.LeftToRightEmbedding + "Start Only: [" + ArabicData.PopDirectionalFormatting + String.Join(" ", Linq.Enumerable.Select(StartWordOnly.ToCharArray(), Function(C As Char) ArabicData.FixStartingCombiningSymbol(CStr(C)))) + ArabicData.LeftToRightEmbedding + "]" + ArabicData.PopDirectionalFormatting, _
                ArabicData.LeftToRightEmbedding + "Not Start: [" + ArabicData.PopDirectionalFormatting + String.Join(" ", Linq.Enumerable.Select(NotStartWord.ToCharArray(), Function(C As Char) ArabicData.FixStartingCombiningSymbol(CStr(C)))) + ArabicData.LeftToRightEmbedding + "]" + ArabicData.PopDirectionalFormatting, _
                ArabicData.LeftToRightEmbedding + "End Only: [" + ArabicData.PopDirectionalFormatting + String.Join(" ", Linq.Enumerable.Select(EndWordOnly.ToCharArray(), Function(C As Char) ArabicData.FixStartingCombiningSymbol(CStr(C)))) + ArabicData.LeftToRightEmbedding + "]" + ArabicData.PopDirectionalFormatting, _
                ArabicData.LeftToRightEmbedding + "Not End: [" + ArabicData.PopDirectionalFormatting + String.Join(" ", Linq.Enumerable.Select(NotEndWord.ToCharArray(), Function(C As Char) ArabicData.FixStartingCombiningSymbol(CStr(C)))) + ArabicData.LeftToRightEmbedding + "]" + ArabicData.PopDirectionalFormatting, _
                ArabicData.LeftToRightEmbedding + "End Only No Diacritics: [" + ArabicData.PopDirectionalFormatting + String.Join(" ", Linq.Enumerable.Select(EndWordOnlyNoDia.ToCharArray(), Function(C As Char) ArabicData.FixStartingCombiningSymbol(CStr(C)))) + ArabicData.LeftToRightEmbedding + "]" + ArabicData.PopDirectionalFormatting, _
                ArabicData.LeftToRightEmbedding + "Not End No Diacritics: [" + ArabicData.PopDirectionalFormatting + String.Join(" ", Linq.Enumerable.Select(NotEndWordNoDia.ToCharArray(), Function(C As Char) ArabicData.FixStartingCombiningSymbol(CStr(C)))) + ArabicData.LeftToRightEmbedding + "]" + ArabicData.PopDirectionalFormatting, _
                ArabicData.LeftToRightEmbedding + "Middle Only: [" + ArabicData.PopDirectionalFormatting + String.Join(" ", Linq.Enumerable.Select(MiddleWordOnly.ToCharArray(), Function(C As Char) ArabicData.FixStartingCombiningSymbol(CStr(C)))) + ArabicData.LeftToRightEmbedding + "]" + ArabicData.PopDirectionalFormatting, _
                ArabicData.LeftToRightEmbedding + "Not Middle: [" + ArabicData.PopDirectionalFormatting + String.Join(" ", Linq.Enumerable.Select(NotMiddleWord.ToCharArray(), Function(C As Char) ArabicData.FixStartingCombiningSymbol(CStr(C)))) + ArabicData.LeftToRightEmbedding + "]" + ArabicData.PopDirectionalFormatting, _
                ArabicData.LeftToRightEmbedding + "Middle Only No Diacritics: [" + ArabicData.PopDirectionalFormatting + String.Join(" ", Linq.Enumerable.Select(MiddleWordOnlyNoDia.ToCharArray(), Function(C As Char) ArabicData.FixStartingCombiningSymbol(CStr(C)))) + ArabicData.LeftToRightEmbedding + "]" + ArabicData.PopDirectionalFormatting, _
                ArabicData.LeftToRightEmbedding + "Not Middle No Diacritics: [" + ArabicData.PopDirectionalFormatting + String.Join(" ", Linq.Enumerable.Select(NotMiddleWordNoDia.ToCharArray(), Function(C As Char) ArabicData.FixStartingCombiningSymbol(CStr(C)))) + ArabicData.LeftToRightEmbedding + "]" + ArabicData.PopDirectionalFormatting, _
                Val, RevVal, DiaVal, LetVal, LetRevVal}
    End Function
    Public Shared Function GetMetaRuleSet(Name As String) As IslamData.RuleMetaSet
        For Count = 0 To CachedData.IslamData.MetaRules.Length - 1
            If CachedData.IslamData.MetaRules(Count).Name = Name Then Return CachedData.IslamData.MetaRules(Count)
        Next
        Return Nothing
    End Function
    Public Shared Function GetRecitationRules() As Array()
        Dim Names() As Array = New List(Of Object())(Linq.Enumerable.Select(GetMetaRuleSet("UthmaniQuran").Rules, Function(Convert As IslamData.RuleMetaSet.RuleMetadataTranslation) New Object() {Utility.LoadResourceString("IslamInfo_" + Convert.Name), CInt(Array.IndexOf(CachedData.IslamData.MetaRules, Convert))})).ToArray()
        Array.Sort(Names, New Utility.CompareNameValueArray)
        Return Names
    End Function
    Public Shared Function GetSelectionNames(Strings As String, SchemeType As ArabicData.TranslitScheme, Scheme As String) As Array()
        Dim Division As Integer = 0
        If Not Strings Is Nothing Then Division = CInt(Strings)
        If Division = 0 Then
            Return TanzilReader.GetChapterNames(SchemeType, Scheme)
        ElseIf Division = 1 Then
            Return TanzilReader.GetChapterNamesByRevelationOrder(SchemeType, Scheme)
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
    Public Shared Function GetScriptFormatsSrc() As Array()
        Return {New Object() {"Uthmani", QuranScripts.Uthmani},
                New Object() {"Uthmani Minimal", QuranScripts.UthmaniMin},
                New Object() {"Simple Minimal", QuranScripts.SimpleMin},
                New Object() {"Simple Enhanced", QuranScripts.SimpleEnhanced},
                New Object() {"Simple Clean", QuranScripts.SimpleClean}}
    End Function
    Public Shared Function GetScriptFormats() As Array()
        Return {New Object() {"Uthmani Minimal", QuranScripts.UthmaniMin},
                New Object() {"Simple", QuranScripts.Simple},
                New Object() {"Simple Minimal", QuranScripts.SimpleMin},
                New Object() {"Simple Enhanced", QuranScripts.SimpleEnhanced},
                New Object() {"Simple Clean", QuranScripts.SimpleClean}}
    End Function
    Public Shared Function GetScriptFormatChangeJS() As String()
        Return New String() {"javascript: doScriptFormatOptChange(this);", String.Empty, "function doScriptFormatOptChange(obj) { var ct, oth = $(obj.id === 'toscript' ? '#fromscript' : '#toscript')[0]; for (ct = 0; ct < oth.options.length - 1 - 1; ct++) { if (obj.options[ct].value !== oth.options[(ct >= oth.selectedIndex) ? ct + 1 : ct].value) break; } oth.options.add(new Option(obj.options[ct].text, obj.options[ct].value), (ct >= oth.selectedIndex) ? ct + 1 : ct); for (ct = 0; ct < oth.options.length - 1; ct++) { if (oth.options[ct].value === obj.options[obj.selectedIndex].value) { oth.options.remove(ct); } } }"}
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
                'If Str.Length = Letters.Length Then Debug.Print(CStr(Count) + ":" + CStr(SubCount))
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
                                    'Debug.Print(CStr(VerseLetProfile.Value(SubCount)(0, 0)) + ":" + CStr(VerseLetProfile.Value(SubCount)(0, 1)) + ",")
                                    For SearchCount = 0 To LetterProfile.Value(Count).GetLength(0) - 1
                                        'Debug.Print(String.Join(",", CStr(LetterProfile.Value(Count)(SearchCount, 0)) + ":" + CStr(LetterProfile.Value(Count)(SearchCount, 1))))
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
            Msg += """" + Arabic.TransliterateToScheme(Keys(Count), ArabicData.TranslitScheme.Literal, String.Empty, Nothing) + """" + If(Count <> Keys.Length - 1, ", ", String.Empty)
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
            Msg += """" + Arabic.TransliterateToScheme(Str, ArabicData.TranslitScheme.Literal, String.Empty, Nothing) + """, "
        Next
        Msg += vbCrLf + "Second: "
        For Each Str As String In CompList
            Dim SubKey As String = Str.Substring(0, Str.IndexOf(LetterPattern))
            For Count As Integer = 0 To SubKey.Length - 1
                If Not CompDict.ContainsKey(SubKey.Substring(Count)) Then
                    CompDict.Add(SubKey.Substring(Count), Str)
                End If
            Next
            Msg += """" + Arabic.TransliterateToScheme(Str, ArabicData.TranslitScheme.Literal, String.Empty, Nothing) + """, "
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
        'Debug.Print(Msg)
    End Sub
    Public Shared Function PatternMatch(BaseText As QuranTexts, ScriptType As QuranScripts, Presentation As ArabicPresentation, Pattern As String) As List(Of String)
        PatternMatch = New List(Of String)
        Dim Doc As Xml.Linq.XDocument
        Dim Stream As IO.Stream
        If ScriptType = QuranScripts.Uthmani Then
            Stream = PortableMethods.FileIO.LoadStream(PortableMethods.Settings.GetFilePath("metadata\" + QuranTextNames(BaseText) + ".xml"))
        Else
            Stream = PortableMethods.FileIO.LoadStream(PortableMethods.Settings.GetFilePath("metadata\" + QuranTextNames(BaseText) + "-" + QuranFileNames(ScriptType) + If(Presentation <> ArabicPresentation.None, "-" + PresentationCacheNames(Presentation), String.Empty) + ".xml"))
        End If
        Doc = Xml.Linq.XDocument.Load(Stream)
        Stream.Dispose()
        Dim Verses As Collections.Generic.List(Of String())
        Verses = TanzilReader.GetQuranText(Doc, -1, -1, -1, -1)
        For Count As Integer = 0 To Verses.Count - 1
            Dim ChapterNode As Xml.Linq.XElement = GetTextChapter(Doc, Count + 1)
            For SubCount As Integer = 0 To Verses(Count).Length - 1
                If SubCount = 0 AndAlso Not GetTextVerse(ChapterNode, SubCount + 1).Attribute("bismillah") Is Nothing Then
                    For Each Val As System.Text.RegularExpressions.Match In System.Text.RegularExpressions.Regex.Matches(GetTextVerse(ChapterNode, SubCount + 1).Attribute("bismillah").Value, Pattern)
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
        Dim Doc As Xml.Linq.XDocument
        Dim Stream As IO.Stream = PortableMethods.FileIO.LoadStream(PortableMethods.Settings.GetFilePath("metadata\" + QuranTextNames(BaseText) + ".xml"))
        Doc = Xml.Linq.XDocument.Load(Stream)
        Stream.Dispose()
        Dim TargetDoc As Xml.Linq.XDocument
        If BaseText = TargetBaseText Then
            Stream = PortableMethods.FileIO.LoadStream(PortableMethods.Settings.GetFilePath("metadata\" + QuranTextNames(TargetBaseText) + "-" + QuranFileNames(ScriptType) + If(Presentation <> ArabicPresentation.None, "-" + PresentationCacheNames(Presentation), String.Empty) + ".xml"))
        Else
            Stream = PortableMethods.FileIO.LoadStream(PortableMethods.Settings.GetFilePath("metadata\" + QuranTextNames(TargetBaseText) + "-" + QuranFileNames(ScriptType) + If(Presentation <> ArabicPresentation.None, "-" + PresentationCacheNames(Presentation), String.Empty) + ".xml"))
        End If
        TargetDoc = Xml.Linq.XDocument.Load(Stream)
        Stream.Dispose()
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
                            'Debug.Print("Chapter: " + CStr(Count + 1) + " Verse: " + CStr(SubCount) + " Words: " + CStr(Words.Length) + " Verse: " + CStr(TargetSubCount) + " Words: " + CStr(TargetWords.Length) + If(Words.Length > TargetWords.Length, "  +", "  -") + CStr(Count + 1) + ":" + CStr(Math.Max(SubCount, TargetSubCount)) + ":" + CStr(Math.Min(Words.Length, TargetWords.Length) + 1))
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
                            'Debug.Print("Chapter: " + CStr(Count + 1) + " Verse: " + CStr(SubCount) + " Words: " + CStr(Words.Length) + " Verse: " + CStr(TargetSubCount) + " Words: " + CStr(TargetWords.Length) + If(Words.Length > TargetWords.Length, "  +", "  -") + CStr(Count + 1) + ":" + CStr(Math.Max(SubCount, TargetSubCount)) + ":" + CStr(Math.Min(Words.Length, TargetWords.Length) + 1))
                        End If
                    End While
                End If
            Loop While Total <= TargetTotal And SubCount <= Verses(Count).Length - 1 Or TargetTotal <= Total And TargetSubCount <= TargetVerses(Count).Length - 1
        Next
    End Sub
    Public Shared Sub ChangeQuranFormat(BaseText As QuranTexts, TargetBaseText As QuranTexts, SrcScriptType As QuranScripts, ScriptType As QuranScripts, Presentation As ArabicPresentation)
        Dim Doc As Xml.Linq.XDocument
        Dim Stream As IO.Stream
        If SrcScriptType = QuranScripts.Uthmani Then
            Stream = PortableMethods.FileIO.LoadStream(PortableMethods.Settings.GetFilePath("metadata\" + QuranTextNames(BaseText) + ".xml"))
        Else
            Stream = PortableMethods.FileIO.LoadStream(PortableMethods.Settings.GetFilePath("IslamMetadata\quran-" + QuranFileNames(SrcScriptType) + ".xml"))
        End If
        Doc = Xml.Linq.XDocument.Load(Stream)
        Stream.Dispose()
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
        CType(Doc.Root.PreviousNode, Xml.Linq.XElement).Value = CType(Doc.Root.PreviousNode, Xml.Linq.XElement).Value.Replace(QuranScriptNames(SrcScriptType), QuranScriptNames(ScriptType))
        Verses = TanzilReader.GetQuranText(Doc, -1, -1, -1, -1)
        For Count As Integer = 0 To Verses.Count - 1
            Dim VerseAdjust As Integer = 0
            Dim ChapterNode As Xml.Linq.XElement = GetTextChapter(Doc, Count + 1)
            If UseBuckwalter Then
                ChapterNode.Attribute("name").Value = Arabic.TransliterateToScheme(ChapterNode.Attribute("name").Value, ArabicData.TranslitScheme.Literal, String.Empty, Nothing)
            End If
            Dim SubCount As Integer = 0
            While SubCount <= Verses(Count).Length - 1 - VerseAdjust
                Dim CurVerse As Xml.Linq.XElement = GetTextVerse(ChapterNode, SubCount + 1)
                Dim PreVerse As String = String.Empty
                Dim NextVerse As String = String.Empty
                If SubCount = 0 AndAlso Not CurVerse.Attribute("bismillah") Is Nothing Then
                    If Count <> 0 Then
                        PreVerse = GetTextVerse(GetTextChapter(Doc, Count), New List(Of Xml.Linq.XNode)(GetTextChapter(Doc, Count).Nodes).Count).Attribute("text").Value
                        If UseBuckwalter Then PreVerse = Arabic.TransliterateFromBuckwalter(PreVerse)
                    End If
                    CurVerse.Attribute("bismillah").Value = If(BaseText = TargetBaseText, Arabic.ChangeScript(CurVerse.Attribute("bismillah").Value, SrcScriptType, ScriptType, PreVerse, CurVerse.Attribute("text").Value), Arabic.ChangeBaseScript(CurVerse.Attribute("bismillah").Value, TargetBaseText, PreVerse, CurVerse.Attribute("text").Value))
                    If UseBuckwalter Then
                        CurVerse.Attribute("bismillah").Value = Arabic.TransliterateToScheme(CurVerse.Attribute("bismillah").Value, ArabicData.TranslitScheme.Literal, String.Empty, Nothing)
                    End If
                    PreVerse = CurVerse.Attribute("bismillah").Value
                End If
                If SubCount + 1 <= Verses(Count).Length - 1 - VerseAdjust Then
                    NextVerse = GetTextVerse(ChapterNode, SubCount + 1 + 1).Attribute("text").Value
                ElseIf Count <> Verses.Count - 1 Then
                    NextVerse = GetTextVerse(GetTextChapter(Doc, Count + 1 + 1), 1).Attribute(If(GetTextVerse(GetTextChapter(Doc, Count + 1 + 1), 1).Attribute("bismillah") Is Nothing, "text", "bismillah")).Value
                End If
                If Count <> 0 AndAlso SubCount = 0 AndAlso CurVerse.Attribute("bismillah") Is Nothing Then
                    PreVerse = GetTextVerse(GetTextChapter(Doc, Count), New List(Of Xml.Linq.XNode)(GetTextChapter(Doc, Count).Nodes).Count).Attribute("text").Value
                ElseIf SubCount <> 0 Then
                    PreVerse = GetTextVerse(ChapterNode, SubCount).Attribute("text").Value
                End If
                If UseBuckwalter Then PreVerse = Arabic.TransliterateFromBuckwalter(PreVerse)
                CurVerse.Attribute("text").Value = If(BaseText = TargetBaseText, Arabic.ChangeScript(CurVerse.Attribute("text").Value, SrcScriptType, ScriptType, PreVerse, NextVerse), Arabic.ChangeBaseScript(CurVerse.Attribute("text").Value, TargetBaseText, PreVerse, NextVerse))
                If UseBuckwalter Then
                    CurVerse.Attribute("text").Value = Arabic.TransliterateToScheme(CurVerse.Attribute("text").Value, ArabicData.TranslitScheme.Literal, String.Empty, Nothing)
                End If
                If BaseText = QuranTexts.Hafs And TargetBaseText = QuranTexts.Warsh Then
                    Dim TCount As Integer = Count
                    Dim Index As Integer = Array.FindIndex(CachedData.IslamData.VerseNumberSchemes(0).CombinedVerses, Function(Ints As Integer()) TCount + 1 = Ints(0) And SubCount + 1 + VerseAdjust - 1 = Ints(1))
                    If Index <> -1 Then
                        If Count = 0 And SubCount = 1 Then
                            Dim NewAttr As Xml.Linq.XAttribute = New Xml.Linq.XAttribute("bismillah", GetTextVerse(ChapterNode, SubCount).Attribute("text").Value)
                            GetTextVerse(ChapterNode, SubCount).Add(NewAttr)
                            GetTextVerse(ChapterNode, SubCount).Attribute("text").Value = CurVerse.Attribute("text").Value
                        Else
                            GetTextVerse(ChapterNode, SubCount).Attribute("text").Value = GetTextVerse(ChapterNode, SubCount).Attribute("text").Value + " " + CurVerse.Attribute("text").Value
                        End If
                        CurVerse.Remove()
                        CurVerse = GetTextVerse(ChapterNode, SubCount)
                        VerseAdjust += 1
                        SubCount -= 1
                        For Index = SubCount + 2 To Verses(Count).Length - 1 - VerseAdjust + 1
                            GetTextVerse(ChapterNode, Index + 1).Attribute("index").Value = CStr(CInt(GetTextVerse(ChapterNode, Index + 1).Attribute("index").Value) - 1)
                        Next
                    End If
                    Index = Array.FindIndex(CachedData.IslamData.VerseNumberSchemes(0).ExtraVerses, Function(Ints As Integer()) TCount + 1 = Ints(0) And SubCount + 1 + VerseAdjust = Ints(1))
                    If Index <> -1 Then
                        Dim NewNode As New Xml.Linq.XElement(CurVerse)
                        If Not NewNode.Attribute("bismillah") Is Nothing Then
                            NewNode.Attribute("bismillah").Remove()
                        End If
                        Index = CachedData.IslamData.VerseNumberSchemes(0).ExtraVerses(Index)(2)
                        While Index <> 1
                            NewNode.Attribute("text").Value = NewNode.Attribute("text").Value.Substring(NewNode.Attribute("text").Value.IndexOf(" "c) + 1)
                            Index -= 1
                        End While
                        CurVerse.Attribute("text").Value = CurVerse.Attribute("text").Value.Substring(0, CurVerse.Attribute("text").Value.Length - NewNode.Attribute("text").Value.Length - 1)
                        CurVerse.AddAfterSelf(NewNode)
                        VerseAdjust -= 1
                        SubCount += 1
                        For Index = Verses(Count).Length - 1 - VerseAdjust - 1 To SubCount Step -1
                            GetTextVerse(ChapterNode, Index + 1).Attribute("index").Value = CStr(CInt(GetTextVerse(ChapterNode, Index + 1).Attribute("index").Value) + 1)
                        Next
                        NewNode.Attribute("index").Value = CStr(CInt(NewNode.Attribute("index").Value) + 1)
                    End If
                End If
                SubCount += 1
            End While
        Next
        Dim MemStream As New IO.MemoryStream
        Doc.Save(MemStream)
        PortableMethods.FileIO.SaveStream(Path, MemStream)
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
                ExtraVerseNumber = New List(Of Xml.Linq.XNode)(GetTextChapter(CachedData.XMLDocMain, BaseChapter).Nodes).Count
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
                ExtraVerseNumber = New List(Of Xml.Linq.XNode)(GetTextChapter(CachedData.XMLDocMain, BaseChapter).Nodes).Count
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
                Dim Matches As System.Text.RegularExpressions.MatchCollection = System.Text.RegularExpressions.Regex.Matches(Arabic.TransliterateToScheme(Verses(Count)(SubCount), ArabicData.TranslitScheme.Literal, String.Empty, Nothing), Str)
                For MatchCount As Integer = 0 To Matches.Count - 1
                    Renderer.Items.AddRange(DoGetRenderedQuranText(QuranTextRangeLookup(Count + 1, SubCount + 1, Array.FindAll(Verses(Count)(SubCount).Substring(0, Matches(MatchCount).Index).ToCharArray(), Function(Ch As Char) Ch = " "c).Length + 1, Count + 1, SubCount + 1, Array.FindAll(Verses(Count)(SubCount).Substring(0, Matches(MatchCount).Index + Matches(MatchCount).Length).ToCharArray(), Function(Ch As Char) Ch = " "c).Length + 1), Count + 1, SubCount + 1, CachedData.IslamData.Translations.TranslationList(TranslationIndex).Name, SchemeType, Scheme, TranslationIndex, W4W, W4WNum, NoArabic, Header, NoRef, Colorize, UseVerses).Items)
                    Dim Reference As String = CStr(Count + 1) + ":" + CStr(SubCount + 1)
                    RefList += If(RefList <> String.Empty, ",", String.Empty) + Reference
                    RefCount += 1
                    If Array.FindAll(Verses(Count)(SubCount).Substring(0, Matches(MatchCount).Index).ToCharArray(), Function(Ch As Char) Ch = " "c).Length <> 0 OrElse Array.FindAll(Verses(Count)(SubCount).Substring(Matches(MatchCount).Index + Matches(MatchCount).Length).ToCharArray(), Function(Ch As Char) Ch = " "c).Length <> 0 Then
                        RefList += ":" + CStr(Array.FindAll(Verses(Count)(SubCount).Substring(0, Matches(MatchCount).Index).ToCharArray(), Function(Ch As Char) Ch = " "c).Length + 1)
                        If Array.FindAll(Verses(Count)(SubCount).Substring(Matches(MatchCount).Index + Matches(MatchCount).Length).ToCharArray(), Function(Ch As Char) Ch = " "c).Length <> 0 Then RefList += "-" + CStr(Array.FindAll(Verses(Count)(SubCount).Substring(0, Matches(MatchCount).Index + Matches(MatchCount).Length).ToCharArray(), Function(Ch As Char) Ch = " "c).Length + 1)
                    End If
                    Renderer.Items.Add(New RenderArray.RenderItem(RenderArray.RenderTypes.eHeaderCenter, New RenderArray.RenderText() {New RenderArray.RenderText(RenderArray.RenderDisplayClass.eLTR, "(" + Arabic.TransliterateToScheme(Arabic.GetCatNoun("QuranReadingRecitation")(0).Text, SchemeType, Scheme, CachedData.RuleMetas("Normal")) + " " + Reference + ")")}))
                Next
            Next
        Next
        Renderer.Items.Add(New RenderArray.RenderItem(RenderArray.RenderTypes.eHeaderCenter, New RenderArray.RenderText() {New RenderArray.RenderText(RenderArray.RenderDisplayClass.eLTR, "(" + Arabic.TransliterateToScheme(Arabic.GetCatNoun("QuranReadingRecitation")(0).Text, SchemeType, Scheme, CachedData.RuleMetas("Normal")) + " " + RefList + ") " + CStr(RefCount) + " Total")}))
        Return Renderer
    End Function
    Public Shared Function TextPositionToMorphology(Text As String, WordPos As Integer) As String
        Dim Chapter As Integer = System.Text.RegularExpressions.Regex.Matches(Text.Substring(0, WordPos), ArabicData.ArabicEndOfAyah + Arabic.TransliterateFromBuckwalter("1") + "\s").Count
        Dim Verse As Integer = Integer.Parse(Arabic.TransliterateToScheme(System.Text.RegularExpressions.Regex.Match(Text.Substring(WordPos), ArabicData.ArabicEndOfAyah + "(\d{1,3})").Groups(1).Value, ArabicData.TranslitScheme.Literal, String.Empty, Nothing))
        If Verse = 1 Then Chapter += 1
        Dim Word As Integer = System.Text.RegularExpressions.Regex.Matches(Text.Substring(0, WordPos).Substring(Text.Substring(0, WordPos).LastIndexOf(ArabicData.ArabicEndOfAyah) + 1), "(\s.)?\s").Count
        Dim Lines As String() = Utility.ReadAllLines(PortableMethods.Settings.GetFilePath("metadata\quranic-corpus-morphology-0.4.txt"))
        TextPositionToMorphology = String.Empty
        For Count As Integer = 0 To Lines.Length - 1
            If Lines(Count).Length <> 0 AndAlso Lines(Count).Chars(0) <> "#" Then
                Dim Pieces As String() = Lines(Count).Split(CChar(vbTab))
                If Pieces(0).Chars(0) = "(" Then
                    Dim Location As Integer() = New List(Of Integer)(Linq.Enumerable.Select(Pieces(0).TrimStart("("c).TrimEnd(")"c).Split(":"c), Function(Str As String) CInt(Str))).ToArray()
                    If Location(0) = Chapter And Location(1) = Verse And Location(2) = Word Then TextPositionToMorphology += If(TextPositionToMorphology = String.Empty, String.Empty, vbCrLf) + Lines(Count)
                End If
            End If
        Next
    End Function
    Public Shared Sub CheckSequentialRules()
        Dim Rules As IslamData.RuleTranslationCategory.RuleTranslation() = CachedData.GetRuleSet("SimpleScriptHamzaWriting") '"HamzaWriting")
        Dim IndexToVerse As Integer()() = Nothing
        Dim XMLDocAlt As Xml.Linq.XDocument
        Dim Stream As IO.Stream = PortableMethods.FileIO.LoadStream(QuranTextNames(QuranTexts.Hafs) + "-" + QuranFileNames(QuranScripts.SimpleEnhanced) + ".xml")
        XMLDocAlt = Xml.Linq.XDocument.Load(Stream)
        Stream.Dispose()
        Dim Text As String = QuranTextCombiner(XMLDocAlt, IndexToVerse) 'CachedData.XMLDocMain
        Dim Matches As System.Text.RegularExpressions.MatchCollection = System.Text.RegularExpressions.Regex.Matches(Text, CachedData.GetPattern("Hamzas"))
        Dim CheckMatches As New Dictionary(Of Integer, String)
        'Debug.Print(CStr(Matches.Count))
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
                    'Debug.Print(CStr(Matches(Count).Groups(2 + If(Rules(MainCount).NegativeMatch <> String.Empty, 1, 0)).Index) + ":" + If(CheckMatches(Matches(Count).Groups(2 + If(Rules(MainCount).NegativeMatch <> String.Empty, 1, 0)).Index).EndsWith("-"), String.Empty, "-") + Rules(MainCount).Name + If(CheckMatches(Matches(Count).Groups(2 + If(Rules(MainCount).NegativeMatch <> String.Empty, 1, 0)).Index).Length = 2, String.Empty, "-") + ":" + Arabic.TransliterateToScheme(Text(Matches(Count).Groups(2 + If(Rules(MainCount).NegativeMatch <> String.Empty, 1, 0)).Index), ArabicData.TranslitScheme.Literal, String.Empty, Nothing) + ":" + Arabic.TransliterateToScheme(Text.Substring(Math.Max(0, Matches(Count).Groups(2 + If(Rules(MainCount).NegativeMatch <> String.Empty, 1, 0)).Index - 15), 30), ArabicData.TranslitScheme.Literal, String.Empty, Nothing))
                Else
                    If Not CheckMatches.ContainsKey(Matches(Count).Groups(2 + If(Rules(MainCount).NegativeMatch <> String.Empty, 1, 0)).Index) Then CheckMatches.Add(Matches(Count).Groups(2 + If(Rules(MainCount).NegativeMatch <> String.Empty, 1, 0)).Index, String.Empty)
                    CheckMatches(Matches(Count).Groups(2 + If(Rules(MainCount).NegativeMatch <> String.Empty, 1, 0)).Index) += If(CheckMatches(Matches(Count).Groups(2 + If(Rules(MainCount).NegativeMatch <> String.Empty, 1, 0)).Index).EndsWith("-"), String.Empty, "-") + Rules(MainCount).Name + If(CheckMatches(Matches(Count).Groups(2 + If(Rules(MainCount).NegativeMatch <> String.Empty, 1, 0)).Index).Length = 2, String.Empty, "-")
                End If
            Next
            'Debug.Print(Rules(MainCount).Name + ": " + CStr(Matches.Count - NegativeCount))
        Next
        Dim Keys(CheckMatches.Keys.Count - 1) As Integer
        CheckMatches.Keys.CopyTo(Keys, 0)
        Array.Sort(Keys)
        For Count = 0 To Keys.Length - 1
            If CheckMatches(Keys(Count)).EndsWith("-") And CheckMatches(Keys(Count)) <> "0-LetterHamzaEnd-YehHamzaKasra-" And CheckMatches(Keys(Count)) <> "0-FathaAlefHamzaAboveSukun-TatweelHamzaSukun-" And CheckMatches(Keys(Count)) <> "0-LetterHamzaEnd-FathaAlefHamzaAboveEnd-" Then
                'Debug.Print(CStr(Keys(Count)) + ":" + CheckMatches(Keys(Count)) + ":" + Arabic.TransliterateToScheme(Text(Keys(Count)), ArabicData.TranslitScheme.Literal, String.Empty, Nothing) + ":" + Arabic.TransliterateToScheme(Text.Substring(Math.Max(0, Keys(Count) - 15), 30), ArabicData.TranslitScheme.Literal, String.Empty, Nothing))
            End If
        Next
    End Sub
    Public Shared Sub CheckMutualExclusiveRules(bAssumeContinue As Boolean)
        'Dim Verify As String() = {CStr(ArabicData.ArabicLetterHamza), ArabicData.ArabicTatweel + "?" + ArabicData.ArabicHamzaAbove, ArabicData.ArabicLetterAlefWithHamzaAbove, ArabicData.ArabicLetterAlefWithHamzaBelow, ArabicData.ArabicLetterWawWithHamzaAbove, ArabicData.ArabicLetterYehWithHamzaAbove}
        Dim VerIndex As Integer = 4
        Dim IndexToVerse As Integer()() = Nothing
        Dim Text As String = QuranTextCombiner(CachedData.XMLDocMain, IndexToVerse)
        Dim Matches As System.Text.RegularExpressions.MatchCollection = System.Text.RegularExpressions.Regex.Matches(Text, CachedData.TranslateRegEx(CachedData.IslamData.VerificationSet(VerIndex).Match, True))
        Dim CheckMatches As New Dictionary(Of Integer, String)
        'Debug.Print(CStr(Matches.Count))
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
            Dim bVerSet As Boolean = False
            For MetaCount = 0 To CachedData.RuleMetas("UthmaniQuran").Length - 1
                If CachedData.IslamData.VerificationSet(VerIndex).MetaRules(MainCount) = CachedData.RuleMetas("UthmaniQuran")(MetaCount).Name Then
                    Exit For
                End If
            Next
            If MetaCount = CachedData.RuleMetas("UthmaniQuran").Length Then
                bVerSet = True
                For MetaCount = 0 To CachedData.RuleMetas("UthmaniQuran").Length - 1
                    If CachedData.IslamData.VerificationSet(VerIndex).MetaRules(MainCount) = CachedData.RuleMetas("UthmaniQuran")(MetaCount).Name Then
                        Exit For
                    End If
                Next
            End If
            Matches = System.Text.RegularExpressions.Regex.Matches(Text, If(bVerSet, CachedData.RuleMetas("UthmaniQuran")(MetaCount).Match, CachedData.RuleMetas("UthmaniQuran")(MetaCount).Match))
            Dim SieveCount As Integer = 0
            For Count = 0 To Matches.Count - 1
                Dim MetaRules As String() = If(bVerSet, CachedData.RuleMetas("UthmaniQuran")(MetaCount).Evaluator, CachedData.RuleMetas("UthmaniQuran")(MetaCount).Evaluator)
                'If Count = 0 AndAlso Matches(Count).Groups.Count <> MetaRules.Length + 1 Then Debug.Print("Discrepency in metadata:" + CStr(MainCount + 1) + ":" + CStr(MetaRules.Length) + ":Got:" + CStr(Matches(Count).Groups.Count - 1))
                Dim bSieve As Boolean = False
                For SubCount = 0 To MetaRules.Length - 1
                    If Array.IndexOf(MetaRules(SubCount).Split("|"c), If(bAssumeContinue, "optionalstop", "optionalnotstop")) <> -1 And Matches(Count).Groups(SubCount + 1).Success Then
                        'check continuity in prior patterns
                        Dim CheckLen As Integer = 0
                        For CheckCount = 1 To SubCount
                            CheckLen += Matches(Count).Groups(CheckCount).Length
                        Next
                        'If Matches(Count).Index + CheckLen <> Matches(Count).Groups(SubCount + 1).Index Then Debug.Print("Non-Sequential Capture:" + CStr(MainCount + 1) + ":" + Arabic.TransliterateToScheme(Text.Substring(Math.Max(0, Matches(Count).Groups(SubCount + 1).Index - 15), 30), ArabicData.TranslitScheme.Literal, String.Empty, Nothing))
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
                            'If Matches(Count).Index + CheckLen <> Matches(Count).Groups(SubCount + 1).Index Then Debug.Print("Non-Sequential Capture:" + CStr(MainCount + 1) + ":" + Arabic.TransliterateToScheme(Text.Substring(Math.Max(0, Matches(Count).Groups(SubCount + 1).Index - 15), 30), ArabicData.TranslitScheme.Literal, String.Empty, Nothing))
                            For LenCount = 0 To Matches(Count).Groups(SubCount + 1).Length - 1
                                If Not CheckMatches.ContainsKey(Matches(Count).Groups(SubCount + 1).Index + LenCount) Then CheckMatches.Add(Matches(Count).Groups(SubCount + 1).Index + LenCount, String.Empty)
                                CheckMatches(Matches(Count).Groups(SubCount + 1).Index + LenCount) += Convert.ToString(MainCount + 1, 16)
                            Next
                            SieveCount += 1
                        End If
                    Next
                End If
            Next
            'Debug.Print(CStr(SieveCount))
        Next
        Dim Keys(CheckMatches.Keys.Count - 1) As Integer
        CheckMatches.Keys.CopyTo(Keys, 0)
        Array.Sort(Keys)
        For Count = 0 To Keys.Length - 1
            If CheckMatches(Keys(Count)).Length <> 2 Then
                'Debug.Print(CStr(Keys(Count)) + ":" + CheckMatches(Keys(Count)) + ":" + Arabic.TransliterateToScheme(Text(Keys(Count)), ArabicData.TranslitScheme.Literal, String.Empty, Nothing) + ":" + Arabic.TransliterateToScheme(Text.Substring(Math.Max(0, Keys(Count) - 15), 30), ArabicData.TranslitScheme.Literal, String.Empty, Nothing))
            End If
        Next
    End Sub
    Public Shared Function QuranTextCombiner(XMLDoc As Xml.Linq.XDocument, ByRef IndexToVerse As Integer()(), Optional StartChapter As Integer = 1, Optional EndChapter As Integer = 0) As String
        Dim Verses As List(Of String()) = GetQuranText(XMLDoc, -1, -1, -1, -1)
        Dim IndexToVerseList As New List(Of Integer())
        Dim Str As New System.Text.StringBuilder
        For Count As Integer = StartChapter - 1 To If(EndChapter <> 0, EndChapter - 1, Verses.Count - 1)
            For SubCount As Integer = 0 To Verses(Count).Length - 1
                Dim Words As String()
                Dim Index As Integer
                If SubCount = 0 Then
                    Dim Node As Xml.Linq.XAttribute
                    Node = GetTextVerse(GetTextChapter(XMLDoc, Count + 1), 1).Attribute("bismillah")
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
            QuranText.Add(TanzilReader.GetQuranText(CachedData.XMLDocMain, BaseChapter, BaseVerse, CInt(If(ExtraVerseNumber <> 0, ExtraVerseNumber, BaseVerse))))
        Else
            QuranText.AddRange(TanzilReader.GetQuranText(CachedData.XMLDocMain, BaseChapter, BaseVerse, EndChapter, ExtraVerseNumber))
        End If
        If (WordNumber > 1) Then
            Dim VerseIndex As Integer = 0
            For WordCount As Integer = 1 To WordNumber - 1
                VerseIndex = QuranText(0)(0).IndexOf(" "c, VerseIndex) + 1
            Next
            QuranText(0)(0) = System.Text.RegularExpressions.Regex.Replace(System.Text.RegularExpressions.Regex.Replace(QuranText(0)(0).Substring(0, VerseIndex), "(^\s*|\s+)[^\s" + String.Join(String.Empty, Linq.Enumerable.Select(CachedData.ArabicStopLetters, Function(Str As String) ArabicData.MakeUniRegEx(Str))) + ArabicData.ArabicStartOfRubElHizb + ArabicData.ArabicPlaceOfSajdah + "]+(?=\s*$|\s+)", "$1"), String.Join("|", Linq.Enumerable.Select(CachedData.ArabicStopLetters, Function(Str As String) ArabicData.MakeUniRegEx(Str))) + "|" + ArabicData.ArabicStartOfRubElHizb + "|" + ArabicData.ArabicPlaceOfSajdah, ChrW(0)) + QuranText(0)(0).Substring(VerseIndex)
        End If
        If (EndWordNumber <> 0) Then
            Dim VerseIndex As Integer = 0
            'selections are always within the same chapter
            Dim LastChapter As Integer = QuranText.Count - 1
            Dim LastVerse As Integer = CInt(If(ExtraVerseNumber <> 0, QuranText(LastChapter).Length - 1, 0))
            While QuranText(LastChapter)(LastVerse)(VerseIndex) = ChrW(0) Or QuranText(LastChapter)(LastVerse)(VerseIndex) = " "
                VerseIndex += 1
            End While
            For WordCount As Integer = WordNumber - 1 To EndWordNumber - 1
                VerseIndex = QuranText(LastChapter)(LastVerse).IndexOf(" "c, VerseIndex) + 1
            Next
            If VerseIndex = 0 Then VerseIndex = QuranText(LastChapter)(LastVerse).Length
            QuranText(LastChapter)(LastVerse) = QuranText(LastChapter)(LastVerse).Substring(0, VerseIndex) + System.Text.RegularExpressions.Regex.Replace(System.Text.RegularExpressions.Regex.Replace(QuranText(LastChapter)(LastVerse).Substring(VerseIndex), "(^\s*|\s+)[^\s" + String.Join(String.Empty, Linq.Enumerable.Select(CachedData.ArabicStopLetters, Function(Str As String) ArabicData.MakeUniRegEx(Str))) + ArabicData.ArabicStartOfRubElHizb + ArabicData.ArabicPlaceOfSajdah + "]+(?=\s*$|\s+)", "$1"), String.Join("|", Linq.Enumerable.Select(CachedData.ArabicStopLetters, Function(Str As String) ArabicData.MakeUniRegEx(Str))) + "|" + ArabicData.ArabicStartOfRubElHizb + "|" + ArabicData.ArabicPlaceOfSajdah, ChrW(0))
        End If
        Return QuranText
    End Function
    Public Shared Function GetQuranTextBySelection(ID As String, Division As Integer, Index As Integer, Translation As String, SchemeType As ArabicData.TranslitScheme, Scheme As String, TranslationIndex As Integer, W4W As Boolean, W4WNum As Boolean, NoArabic As Boolean, Header As Boolean, NoRef As Boolean, Colorize As Boolean, Verses As Boolean) As RenderArray
        Dim Chapter As Integer
        Dim Verse As Integer
        Dim BaseChapter As Integer
        Dim BaseVerse As Integer
        Dim Node As Xml.Linq.XElement
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
                BaseChapter = CInt(GetChapterByIndex(Index).Attribute("index").Value)
                BaseVerse = 1
                QuranText = New Collections.Generic.List(Of String())
                QuranText.Add(GetQuranText(CachedData.XMLDocMain, BaseChapter, BaseVerse, -1))
            ElseIf Division = 1 Then
                BaseChapter = CInt(TanzilReader.GetChapterIndexByRevelationOrder(Index).Attribute("index").Value)
                BaseVerse = 1
                QuranText = New Collections.Generic.List(Of String())
                QuranText.Add(GetQuranText(CachedData.XMLDocMain, BaseChapter, BaseVerse, -1))
            ElseIf Division = 2 Then
                Node = TanzilReader.GetPartByIndex(Index)
                BaseChapter = CInt(Node.Attribute("sura").Value)
                BaseVerse = CInt(Node.Attribute("aya").Value)
                Node = TanzilReader.GetPartByIndex(Index + 1)
                If Node Is Nothing Then
                    Chapter = -1
                    Verse = -1
                Else
                    Chapter = CInt(Node.Attribute("sura").Value)
                    Verse = CInt(Node.Attribute("aya").Value)
                    TanzilReader.GetPreviousChapterVerse(Chapter, Verse)
                End If
                QuranText = GetQuranText(CachedData.XMLDocMain, BaseChapter, BaseVerse, Chapter, Verse)
            ElseIf Division = 3 Then
                Node = TanzilReader.GetGroupByIndex(Index)
                BaseChapter = CInt(Node.Attribute("sura").Value)
                BaseVerse = CInt(Node.Attribute("aya").Value)
                Node = TanzilReader.GetGroupByIndex(Index + 1)
                If Node Is Nothing Then
                    Chapter = -1
                    Verse = -1
                Else
                    Chapter = CInt(Node.Attribute("sura").Value)
                    Verse = CInt(Node.Attribute("aya").Value)
                    TanzilReader.GetPreviousChapterVerse(Chapter, Verse)
                End If
                QuranText = GetQuranText(CachedData.XMLDocMain, BaseChapter, BaseVerse, Chapter, Verse)
            ElseIf Division = 4 Then
                Node = TanzilReader.GetStationByIndex(Index)
                BaseChapter = CInt(Node.Attribute("sura").Value)
                BaseVerse = CInt(Node.Attribute("aya").Value)
                Node = TanzilReader.GetStationByIndex(Index + 1)
                If Node Is Nothing Then
                    Chapter = -1
                    Verse = -1
                Else
                    Chapter = CInt(Node.Attribute("sura").Value)
                    Verse = CInt(Node.Attribute("aya").Value)
                    TanzilReader.GetPreviousChapterVerse(Chapter, Verse)
                End If
                QuranText = GetQuranText(CachedData.XMLDocMain, BaseChapter, BaseVerse, Chapter, Verse)
            ElseIf Division = 5 Then
                Node = TanzilReader.GetSectionByIndex(Index)
                BaseChapter = CInt(Node.Attribute("sura").Value)
                BaseVerse = CInt(Node.Attribute("aya").Value)
                Node = TanzilReader.GetSectionByIndex(Index + 1)
                If Node Is Nothing Then
                    Chapter = -1
                    Verse = -1
                Else
                    Chapter = CInt(Node.Attribute("sura").Value)
                    Verse = CInt(Node.Attribute("aya").Value)
                    TanzilReader.GetPreviousChapterVerse(Chapter, Verse)
                End If
                QuranText = GetQuranText(CachedData.XMLDocMain, BaseChapter, BaseVerse, Chapter, Verse)
            ElseIf Division = 6 Then
                Node = TanzilReader.GetPageByIndex(Index)
                BaseChapter = CInt(Node.Attribute("sura").Value)
                BaseVerse = CInt(Node.Attribute("aya").Value)
                Node = TanzilReader.GetPageByIndex(Index + 1)
                If Node Is Nothing Then
                    Chapter = -1
                    Verse = -1
                Else
                    Chapter = CInt(Node.Attribute("sura").Value)
                    Verse = CInt(Node.Attribute("aya").Value)
                    TanzilReader.GetPreviousChapterVerse(Chapter, Verse)
                End If
                QuranText = GetQuranText(CachedData.XMLDocMain, BaseChapter, BaseVerse, Chapter, Verse)
            ElseIf Division = 7 Then
                Node = TanzilReader.GetSajdaByIndex(Index)
                BaseChapter = CInt(Node.Attribute("sura").Value)
                BaseVerse = CInt(Node.Attribute("aya").Value)
                QuranText = New Collections.Generic.List(Of String())
                QuranText.Add(GetQuranText(CachedData.XMLDocMain, BaseChapter, BaseVerse, BaseVerse))
            ElseIf Division = 8 Then
                BaseChapter = CachedData.IslamData.QuranSelections(Index).SelectionInfo(SectionCount).ChapterNumber
                BaseVerse = CachedData.IslamData.QuranSelections(Index).SelectionInfo(SectionCount).VerseNumber
                QuranText = QuranTextRangeLookup(BaseChapter, BaseVerse, CachedData.IslamData.QuranSelections(Index).SelectionInfo(SectionCount).WordNumber, 0, CachedData.IslamData.QuranSelections(Index).SelectionInfo(SectionCount).ExtraVerseNumber, CachedData.IslamData.QuranSelections(Index).SelectionInfo(SectionCount).EndWordNumber)
            ElseIf Division = 9 Then
                QuranText = New Collections.Generic.List(Of String())
                For SubCount = 0 To CachedData.LetterDictionary(ArabicData.ArabicLetters(Index).Symbol)(Keys(SectionCount)).Count - 1
                    BaseChapter = CachedData.LetterDictionary(ArabicData.ArabicLetters(Index).Symbol)(Keys(SectionCount))(SubCount)(0)
                    BaseVerse = CachedData.LetterDictionary(ArabicData.ArabicLetters(Index).Symbol)(Keys(SectionCount))(SubCount)(1)
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
                Renderer.Items.AddRange(New RenderArray.RenderItem() {New RenderArray.RenderItem(RenderArray.RenderTypes.eText, New RenderArray.RenderText() {New RenderArray.RenderText(RenderArray.RenderDisplayClass.eLTR, GetMetaRuleSet("UthmaniQuran").Rules(Index).Name)}), New RenderArray.RenderItem(RenderArray.RenderTypes.eText, DocBuilder.ColorizeRegExGroups(DocBuilder.GetRegExText(GetMetaRuleSet("UthmaniQuran").Rules(Index).Match), False)), New RenderArray.RenderItem(RenderArray.RenderTypes.eText, DocBuilder.ColorizeList(New List(Of String)(Linq.Enumerable.Select(GetMetaRuleSet("UthmaniQuran").Rules(Index).Evaluator, Function(Str As String) DocBuilder.GetRegExText(Str))).ToArray(), False))})
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
    Public Shared Function GetRenderedQuranText(SchemeType As ArabicData.TranslitScheme, Scheme As String, Name As String, Strings As String, QuranSelection As String, QuranTranslation As String, WordVerseMode As String, ColorCueMode As String) As RenderArray
        Dim Division As Integer = 0
        Dim Index As Integer = 1
        If Not Strings Is Nothing Then Division = CInt(Strings)
        If Not QuranSelection Is Nothing Then Index = CInt(QuranSelection)
        Dim TranslationIndex As Integer = GetTranslationIndex(QuranTranslation)
        Return GetQuranTextBySelection(Name, Division, Index, CachedData.IslamData.Translations.TranslationList(TranslationIndex).FileName, SchemeType, Scheme, TranslationIndex, CInt(WordVerseMode) <> 4, CInt(WordVerseMode) Mod 2 = 1, False, True, False, CInt(ColorCueMode) = 0, CInt(WordVerseMode) / 2 <> 1)
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
    Public Shared Sub WordFileToResource(WordFilePath As String, ResFilePath As String)
        '"metadata\en.w4w.corpus.txt"
        Dim W4WLines As String() = Utility.ReadAllLines(WordFilePath)
        Dim Stream As IO.Stream = PortableMethods.FileIO.LoadStream(ResFilePath)
        Dim XMLDoc As Xml.Linq.XDocument = Xml.Linq.XDocument.Load(Stream)
        Stream.Dispose()
        For Count = 0 To W4WLines.Length - 1
            Dim Words As String() = W4WLines(Count).Split("|"c)
            For SubCount = 0 To Words.Length - 1
                Dim NewData As New Xml.Linq.XElement("data")
                NewData.Add(New Xml.Linq.XAttribute("name", "Quran" + CStr(Count + 1) + "." + CStr(SubCount + 1)))
                NewData.Add(New Xml.Linq.XAttribute("xml:space", "preserve"))
                Dim Inner As New Xml.Linq.XElement("value")
                Inner.Value = Words(SubCount)
                NewData.Add(Inner)
                XMLDoc.Root.Add(NewData)
            Next
        Next
        Dim MemStream As New IO.MemoryStream
        XMLDoc.Save(MemStream)
        PortableMethods.FileIO.SaveStream(ResFilePath, MemStream)
        MemStream.Dispose()
    End Sub
    Public Shared Sub ResourceToWordFile(ResFilePath As String, WordFilePath As String)
        Dim W4WLines As New List(Of List(Of String))
        Dim Stream As IO.Stream = PortableMethods.FileIO.LoadStream(ResFilePath)
        Dim XMLDoc As Xml.Linq.XDocument = XML.Linq.XDocument.Load(Stream)
        Stream.Dispose()
        Dim AllNodes As Xml.Linq.XElement() = Nothing '(New List(Of Xml.Linq.XElement)(System.Xml.XPath.Extensions.XPathSelectElements(XMLDoc.Root, "data/@name"))).ToArray()
        For Each Item As Xml.Linq.XElement In AllNodes
            If System.Text.RegularExpressions.Regex.Match(Item.Attribute("name").Value, "^Quran%d+\.%d+$").Success Then
                Dim Line As Integer = Integer.Parse(Item.Attribute("name").Value.Substring(5, Item.Attribute("name").Value.IndexOf("."c) - 5))
                Dim Word As Integer = Integer.Parse(Item.Attribute("name").Value.Substring(Item.Attribute("name").Value.IndexOf("."c) + 1))
                If W4WLines(Line - 1) Is Nothing Then
                    W4WLines.Insert(Line, New List(Of String))
                End If
                W4WLines(Line - 1).Insert(Word - 1, CType(Item.FirstNode, Xml.Linq.XElement).Value)
            End If
        Next
        Utility.WriteAllLines(WordFilePath, New List(Of String)(Linq.Enumerable.Select(W4WLines, Function(Input As List(Of String)) String.Join("|"c, Input.ToArray()))).ToArray())
    End Sub
    Public Shared Function DoGetRenderedQuranText(QuranText As Collections.Generic.List(Of String()), BaseChapter As Integer, BaseVerse As Integer, Translation As String, SchemeType As ArabicData.TranslitScheme, Scheme As String, TranslationIndex As Integer, W4W As Boolean, W4WNum As Boolean, NoArabic As Boolean, Header As Boolean, NoRef As Boolean, Colorize As Boolean, Verses As Boolean) As RenderArray
        Dim Text As String
        Dim Node As Xml.Linq.XAttribute
        Dim Renderer As New RenderArray(String.Empty)
        Dim Lines As String() = Utility.ReadAllLines(PortableMethods.Settings.GetFilePath("metadata\" + GetTranslationFileName(Translation)))
        Dim W4WLines As String() = If(W4W, Utility.ReadAllLines(PortableMethods.Settings.GetFilePath("metadata\en.w4w.corpus.txt")), Nothing)
        If Not QuranText Is Nothing Then
            For Chapter = 0 To QuranText.Count - 1
                Dim ChapterNode As Xml.Linq.XElement = GetChapterByIndex(BaseChapter + Chapter)
                Dim Texts As New List(Of RenderArray.RenderText)
                If Header Then
                    If Not NoArabic Then
                        Texts.Add(New RenderArray.RenderText(RenderArray.RenderDisplayClass.eArabic, Arabic.TransliterateFromBuckwalter(CachedData.QuranHeaders(0) + " " + ChapterNode.Attribute("ayas").Value)))
                        Texts.Add(New RenderArray.RenderText(RenderArray.RenderDisplayClass.eTransliteration, If(SchemeType <> ArabicData.TranslitScheme.None, Arabic.TransliterateToScheme(Arabic.TransliterateFromBuckwalter(CachedData.QuranHeaders(0) + " " + ChapterNode.Attribute("ayas").Value), SchemeType, Scheme, CachedData.RuleMetas("Normal")).Trim(), String.Empty)))
                    End If
                    If Translation <> String.Empty Then Texts.Add(New RenderArray.RenderText(RenderArray.RenderDisplayClass.eLTR, "Verses " + ChapterNode.Attribute("ayas").Value + " "))
                    Renderer.Items.Add(New RenderArray.RenderItem(RenderArray.RenderTypes.eHeaderLeft, Texts.ToArray()))
                    Texts.Clear()
                    If Not NoArabic Then
                        Texts.Add(New RenderArray.RenderText(RenderArray.RenderDisplayClass.eArabic, Arabic.TransliterateFromBuckwalter(CachedData.QuranHeaders(1) + " " + CachedData.IslamData.QuranChapters(CInt(ChapterNode.Attribute("index").Value) - 1).Name)))
                        Texts.Add(New RenderArray.RenderText(RenderArray.RenderDisplayClass.eTransliteration, If(SchemeType <> ArabicData.TranslitScheme.None, Arabic.TransliterateToScheme(Arabic.TransliterateFromBuckwalter(CachedData.QuranHeaders(1) + " " + CachedData.IslamData.QuranChapters(CInt(ChapterNode.Attribute("index").Value) - 1).Name), SchemeType, Scheme, CachedData.RuleMetas("Normal")).Trim(), String.Empty)))
                    End If
                    If Translation <> String.Empty Then Texts.Add(New RenderArray.RenderText(RenderArray.RenderDisplayClass.eLTR, "Chapter " + TanzilReader.GetChapterEName(ChapterNode) + " "))
                    Renderer.Items.Add(New RenderArray.RenderItem(RenderArray.RenderTypes.eHeaderCenter, Texts.ToArray()))
                    Texts.Clear()
                    If Not NoArabic Then
                        Texts.Add(New RenderArray.RenderText(RenderArray.RenderDisplayClass.eArabic, Arabic.TransliterateFromBuckwalter(CachedData.QuranHeaders(2) + " " + ChapterNode.Attribute("rukus").Value)))
                        Texts.Add(New RenderArray.RenderText(RenderArray.RenderDisplayClass.eTransliteration, If(SchemeType <> ArabicData.TranslitScheme.None, Arabic.TransliterateToScheme(Arabic.TransliterateFromBuckwalter(CachedData.QuranHeaders(2) + " " + ChapterNode.Attribute("rukus").Value), SchemeType, Scheme, CachedData.RuleMetas("Normal")).Trim(), String.Empty)))
                    End If
                    If Translation <> String.Empty Then Texts.Add(New RenderArray.RenderText(RenderArray.RenderDisplayClass.eLTR, "Rukus " + ChapterNode.Attribute("rukus").Value + " "))
                    Renderer.Items.Add(New RenderArray.RenderItem(RenderArray.RenderTypes.eHeaderRight, Texts.ToArray()))
                    Texts.Clear()
                End If
                For Verse = 0 To QuranText(Chapter).Length - 1
                    Dim Items As New Collections.Generic.List(Of RenderArray.RenderItem)
                    Text = String.Empty
                    'hizb symbols not needed as Quranic text already contains them
                    'If BaseChapter + Chapter <> 1 AndAlso TanzilReader.IsQuarterStart(BaseChapter + Chapter, CInt(If(Chapter = 0, BaseVerse, 1)) + Verse) Then
                    '    Text += Arabic.TransliterateFromBuckwalter("B")
                    '    Items.Add(New RenderArray.RenderItem(RenderArray.RenderTypes.eText, New RenderArray.RenderText() {New RenderArray.RenderText(RenderArray.RenderDisplayClass.eArabic, Arabic.TransliterateFromBuckwalter("B"))}))
                    'End If
                    If CInt(If(Chapter = 0, BaseVerse, 1)) + Verse = 1 Then
                        Node = GetTextVerse(GetTextChapter(CachedData.XMLDocMain, BaseChapter + Chapter), 1).Attribute("bismillah")
                        If Not Node Is Nothing Then
                            If Not NoArabic Then
                                Texts.Add(New RenderArray.RenderText(RenderArray.RenderDisplayClass.eArabic, Node.Value + " "))
                                Texts.Add(New RenderArray.RenderText(RenderArray.RenderDisplayClass.eTransliteration, If(SchemeType <> ArabicData.TranslitScheme.None, Arabic.TransliterateToScheme(Node.Value, SchemeType, Scheme, CachedData.RuleMetas("UthmaniQuran")).Trim(), String.Empty)))
                            End If
                            If Translation <> String.Empty Then Texts.Add(New RenderArray.RenderText(If(IsTranslationTextLTR(TranslationIndex), RenderArray.RenderDisplayClass.eLTR, RenderArray.RenderDisplayClass.eRTL), TanzilReader.GetTranslationVerse(Lines, 1, 1)))
                            Renderer.Items.Add(New RenderArray.RenderItem(RenderArray.RenderTypes.eText, Texts.ToArray()))
                            Texts.Clear()
                        End If
                    End If
                    Text += QuranText(Chapter)(Verse).Trim(" "c, ChrW(0)) + If(NoRef, String.Empty, " ")
                    If TanzilReader.IsSajda(BaseChapter + Chapter, CInt(If(Chapter = 0, BaseVerse, 1)) + Verse) Then
                        'Sajda markers are already in the text
                        'Text += Arabic.TransliterateFromBuckwalter("R")
                        'Items.Add(New RenderArray.RenderItem(RenderArray.RenderTypes.eText, New RenderArray.RenderText() {New RenderArray.RenderText(RenderArray.RenderDisplayClass.eArabic, Arabic.TransliterateFromBuckwalter("R"))}))
                    End If
                    If NoRef Then Text = Arabic.TransliterateFromBuckwalter("(") + Text
                    Text += Arabic.TransliterateFromBuckwalter(If(NoRef, ")", "=" + CStr(CInt(If(Chapter = 0, BaseVerse, 1)) + Verse))) + " "
                    If W4W And Translation <> String.Empty Then
                        Dim DefStops As Integer() = GenerateDefaultStops(QuranText(Chapter)(Verse))
                        Dim Words As String() = If(QuranText(Chapter)(Verse) Is Nothing, {}, QuranText(Chapter)(Verse).Split(" "c))
                        Dim TranslitWords As String() = Arabic.TransliterateToScheme(QuranText(Chapter)(Verse), SchemeType, Scheme, CachedData.RuleMetas("UthmaniQuran"), DefStops).Split(" "c)
                        Dim PauseMarks As Integer = 0
                        If NoRef Then
                            If Not NoArabic Then
                                Texts.Add(New RenderArray.RenderText(RenderArray.RenderDisplayClass.eArabic, Arabic.TransliterateFromBuckwalter("(")))
                                Texts.Add(New RenderArray.RenderText(RenderArray.RenderDisplayClass.eTransliteration, If(SchemeType <> ArabicData.TranslitScheme.None, Arabic.TransliterateToScheme(Arabic.TransliterateFromBuckwalter(")"), SchemeType, Scheme, CachedData.RuleMetas("Normal")), String.Empty)))
                            End If
                            If Translation <> String.Empty Then Texts.Add(New RenderArray.RenderText(If(IsTranslationTextLTR(TranslationIndex), RenderArray.RenderDisplayClass.eLTR, RenderArray.RenderDisplayClass.eRTL), ")"))
                            Items.Add(New RenderArray.RenderItem(RenderArray.RenderTypes.eText, Texts.ToArray()))
                            Texts.Clear()
                        End If
                        Dim WordColors As RenderArray.RenderText()() = Nothing
                        If Colorize Then WordColors = Arabic.ApplyColorRules(Text, GenerateDefaultStops(Text), True, CachedData.RuleMetas("UthmaniQuran"))
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
                                    Texts.Add(New RenderArray.RenderText(RenderArray.RenderDisplayClass.eTransliteration, If(SchemeType <> ArabicData.TranslitScheme.None, TranslitWords(Count - PauseMarks), String.Empty)))
                                End If
                                If Translation <> String.Empty Then Texts.Add(New RenderArray.RenderText(If(IsTranslationTextLTR(TranslationIndex), RenderArray.RenderDisplayClass.eLTR, RenderArray.RenderDisplayClass.eRTL), TanzilReader.GetW4WTranslationVerse(W4WLines, BaseChapter + Chapter, CInt(If(Chapter = 0, BaseVerse, 1)) + Verse, Count - PauseMarks)))
                                If W4WNum Then Texts.Add(New RenderArray.RenderText(RenderArray.RenderDisplayClass.eLTR, CStr(Count + 1)))
                                Items.Add(New RenderArray.RenderItem(RenderArray.RenderTypes.eText, Texts.ToArray()))
                                If Not NoArabic AndAlso Count <> Words.Length - 1 AndAlso Words(Count + 1).Length <> 1 Then Items.Add(New RenderArray.RenderItem(RenderArray.RenderTypes.eText, {New RenderArray.RenderText(RenderArray.RenderDisplayClass.eContinueStop, True)}))
                                Texts.Clear()
                            End If
                            Pos += Words(Count).Length + 1
                        Next
                        If Not NoArabic Then
                            Texts.Add(New RenderArray.RenderText(RenderArray.RenderDisplayClass.eArabic, Arabic.TransliterateFromBuckwalter(If(NoRef, ")", "=" + CStr(CInt(If(Chapter = 0, BaseVerse, 1)) + Verse)))))
                            Texts.Add(New RenderArray.RenderText(RenderArray.RenderDisplayClass.eTransliteration, If(SchemeType <> ArabicData.TranslitScheme.None, Arabic.TransliterateToScheme(Arabic.TransliterateFromBuckwalter(If(NoRef, "(", "=" + CStr(CInt(If(Chapter = 0, BaseVerse, 1)) + Verse))), SchemeType, Scheme, CachedData.RuleMetas("Normal")), String.Empty)))
                            Texts.Add(New RenderArray.RenderText(RenderArray.RenderDisplayClass.eContinueStop, False))
                        End If
                        If Translation <> String.Empty Then Texts.Add(New RenderArray.RenderText(If(IsTranslationTextLTR(TranslationIndex), RenderArray.RenderDisplayClass.eLTR, RenderArray.RenderDisplayClass.eRTL), "(" + If(NoRef, String.Empty, CStr(CInt(If(Chapter = 0, BaseVerse, 1)) + Verse) + ")")))
                        Items.Add(New RenderArray.RenderItem(RenderArray.RenderTypes.eText, Texts.ToArray()))
                        Texts.Clear()
                        'Text += Arabic.TransliterateFromBuckwalter("(" + CStr(If(Chapter = 0, BaseVerse, 1) + Verse) + ") ")
                        Texts.Add(New RenderArray.RenderText(RenderArray.RenderDisplayClass.eNested, Items))
                    End If
                    If Verses Then
                        If Not NoArabic Then
                            If Colorize Then
                                Texts.AddRange(Arabic.ApplyColorRules(Text, GenerateDefaultStops(Text), False, CachedData.RuleMetas("UthmaniQuran"))(0))
                            Else
                                Texts.Add(New RenderArray.RenderText(RenderArray.RenderDisplayClass.eArabic, Text))
                            End If
                            Texts.Add(New RenderArray.RenderText(RenderArray.RenderDisplayClass.eTransliteration, If(SchemeType <> ArabicData.TranslitScheme.None, Arabic.TransliterateToScheme(If(NoRef, Arabic.TransliterateFromBuckwalter("("), String.Empty) + QuranText(Chapter)(Verse).Trim(" "c, ChrW(0)) + Arabic.TransliterateFromBuckwalter(If(NoRef, ")", " " + "=" + CStr(CInt(If(Chapter = 0, BaseVerse, 1)) + Verse) + " ")), SchemeType, Scheme, CachedData.RuleMetas("UthmaniQuran"), GenerateDefaultStops(If(NoRef, Arabic.TransliterateFromBuckwalter("("), String.Empty) + QuranText(Chapter)(Verse).Trim(" "c, ChrW(0)) + Arabic.TransliterateFromBuckwalter(If(NoRef, ")", " " + "=" + CStr(CInt(If(Chapter = 0, BaseVerse, 1)) + Verse)) + " "))).Trim(), String.Empty)))
                        End If
                        If Translation <> String.Empty Then Texts.Add(New RenderArray.RenderText(If(IsTranslationTextLTR(TranslationIndex), RenderArray.RenderDisplayClass.eLTR, RenderArray.RenderDisplayClass.eRTL), If(NoRef, String.Empty, "(" + CStr(CInt(If(Chapter = 0, BaseVerse, 1)) + Verse) + ") ") + TanzilReader.GetTranslationVerse(Lines, BaseChapter + Chapter, CInt(If(Chapter = 0, BaseVerse, 1)) + Verse)))
                    End If
                    Renderer.Items.Add(New RenderArray.RenderItem(RenderArray.RenderTypes.eText, Texts.ToArray()))
                    If Not NoArabic Then Renderer.Items.Add(New RenderArray.RenderItem(RenderArray.RenderTypes.eText, {New RenderArray.RenderText(RenderArray.RenderDisplayClass.eContinueStop, False)}))
                    Texts.Clear()
                Next
            Next
        End If
        Return Renderer
    End Function
    Public Shared Function GetQuranText(ByVal XMLDocMain As Xml.Linq.XDocument, ByVal StartChapter As Integer, ByVal StartAyat As Integer, ByVal EndChapter As Integer, ByVal EndAyat As Integer) As Collections.Generic.List(Of String())
        Dim Count As Integer
        If StartChapter = -1 Then StartChapter = 1
        If EndChapter = -1 Then EndChapter = GetChapterCount()
        Dim ChapterVerses As New Collections.Generic.List(Of String())
        For Count = StartChapter To EndChapter
            ChapterVerses.Add(GetQuranText(XMLDocMain, Count, CInt(If(StartChapter = Count, StartAyat, -1)), CInt(If(EndChapter = Count, EndAyat, -1))))
        Next
        Return ChapterVerses
    End Function
    Public Shared Function GetQuranText(ByVal XMLDocMain As Xml.Linq.XDocument, ByVal Chapter As Integer, ByVal StartVerse As Integer, ByVal EndVerse As Integer) As String()
        Dim Count As Integer
        If StartVerse = -1 Then StartVerse = 1
        If EndVerse = -1 Then EndVerse = New List(Of Xml.Linq.XNode)(GetTextChapter(XMLDocMain, Chapter).Nodes).Count 'GetVerseCount(Chapter)
        Dim Verses(EndVerse - StartVerse) As String
        For Count = StartVerse To EndVerse
            Dim VerseNode As Xml.Linq.XElement = GetTextVerse(GetTextChapter(XMLDocMain, Chapter), Count)
            If Not VerseNode Is Nothing Then
                Dim AttrNode As Xml.Linq.XAttribute = VerseNode.Attribute("text")
                If Not AttrNode Is Nothing Then
                    Verses(Count - StartVerse) = AttrNode.Value
                End If
            End If
        Next
        Return Verses
    End Function
    Public Shared Function GetTextChapter(ByVal XMLDocMain As Xml.Linq.XDocument, ByVal Chapter As Integer) As Xml.Linq.XElement
        Return Utility.GetChildNodeByIndex("sura", "index", Chapter, New List(Of Xml.Linq.XElement)(XMLDocMain.Root.Elements).ToArray())
    End Function
    Public Shared Function GetTextVerse(ByVal ChapterNode As Xml.Linq.XElement, ByVal Verse As Integer) As Xml.Linq.XElement
        Return Utility.GetChildNodeByIndex("aya", "index", Verse, New List(Of Xml.Linq.XElement)(ChapterNode.Elements).ToArray())
    End Function
    Public Shared Function GetVerseCount(ByVal Chapter As Integer) As Integer
        Return CInt(GetChapterByIndex(Chapter).Attribute("ayas").Value)
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
        Return CInt(GetChapterByIndex(Chapter).Attribute("start").Value) + Verse
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
        For Each Node As Xml.Linq.XElement In Utility.GetChildNode("hizbs", New List(Of Xml.Linq.XElement)(CachedData.XMLDocInfo.Root.Elements).ToArray()).Elements
            If Node.Name = "quarter" AndAlso _
                CInt(Node.Attribute("sura").Value) = Chapter AndAlso _
                CInt(Node.Attribute("aya").Value) = Verse Then
                Return True
            End If
        Next
        Return False
    End Function
    Public Shared Function IsSajda(ByVal Chapter As Integer, ByVal Verse As Integer) As Boolean
        Dim Node As Xml.Linq.XElement
        For Each Node In Utility.GetChildNode("sajdas", New List(Of Xml.Linq.XElement)(CachedData.XMLDocInfo.Root.Elements).ToArray()).Elements
            If Node.Name = "sajda" AndAlso _
                CInt(Node.Attribute("sura").Value) = Chapter AndAlso _
                CInt(Node.Attribute("aya").Value) = Verse Then
                Return True
            End If
        Next
        Return False
    End Function
    Public Shared Function GetImportantNames() As Array()
        Dim Names() As Array = New List(Of Object())(Linq.Enumerable.Select(CachedData.IslamData.QuranSelections, Function(Convert As IslamData.QuranSelection) New Object() {Utility.LoadResourceString("IslamInfo_" + Convert.Description), CInt(Array.IndexOf(CachedData.IslamData.QuranSelections, Convert))})).ToArray()
        Array.Sort(Names, New Utility.CompareNameValueArray)
        Return Names
    End Function
    Public Shared Function GetChapterCount() As Integer
        Return Utility.GetChildNodeCount("sura", Utility.GetChildNode("suras", New List(Of Xml.Linq.XElement)(CachedData.XMLDocInfo.Root.Elements).ToArray()))
    End Function
    Public Shared Function GetChapterByIndex(ByVal Index As Integer) As Xml.Linq.XElement
        Return Utility.GetChildNodeByIndex("sura", "index", Index, New List(Of Xml.Linq.XElement)(Utility.GetChildNode("suras", New List(Of Xml.Linq.XElement)(CachedData.XMLDocInfo.Root.Elements).ToArray()).Elements).ToArray())
    End Function
    Public Shared Function GetChapterNames(SchemeType As ArabicData.TranslitScheme, Scheme As String) As Array()
        Dim Names() As Array = New List(Of Object())(Linq.Enumerable.Select(Utility.GetChildNodes("sura", New List(Of Xml.Linq.XElement)(Utility.GetChildNode("suras", New List(Of Xml.Linq.XElement)(CachedData.XMLDocInfo.Root.Elements).ToArray()).Elements).ToArray()), Function(Convert As Xml.Linq.XElement) New Object() {Convert.Attribute("index").Value + ". " + GetChapterEName(Convert) + " (" + ArabicData.RightToLeftEmbedding + Arabic.TransliterateFromBuckwalter(CachedData.QuranHeaders(1) + " " + CachedData.IslamData.QuranChapters(CInt(Convert.Attribute("index").Value) - 1).Name) + ArabicData.PopDirectionalFormatting + ")" + If(SchemeType = ArabicData.TranslitScheme.None, String.Empty, " " + Arabic.TransliterateToScheme(Arabic.TransliterateFromBuckwalter(CachedData.IslamData.QuranChapters(CInt(Convert.Attribute("index").Value) - 1).Name), SchemeType, Scheme, CachedData.RuleMetas("Normal"))), CInt(Convert.Attribute("index").Value)})).ToArray()
        Array.Sort(Names, New Utility.CompareNameValueArray)
        Return Names
    End Function
    Public Shared Function GetChapterEName(ByVal ChapterNode As Xml.Linq.XElement) As String
        Return Utility.LoadResourceString("IslamInfo_QuranChapter" + ChapterNode.Attribute("index").Value)
    End Function
    Public Shared Function GetChapterNamesByRevelationOrder(SchemeType As ArabicData.TranslitScheme, Scheme As String) As Array()
        Dim Names() As Array = New List(Of Object())(Linq.Enumerable.Select(Utility.GetChildNodes("sura", New List(Of Xml.Linq.XElement)(Utility.GetChildNode("suras", New List(Of Xml.Linq.XElement)(CachedData.XMLDocInfo.Root.Elements).ToArray()).Elements).ToArray()), Function(Convert As Xml.Linq.XElement) New Object() {Convert.Attribute("index").Value + ". " + GetChapterEName(Convert) + " (" + ArabicData.RightToLeftEmbedding + Arabic.TransliterateFromBuckwalter(CachedData.QuranHeaders(1) + " " + CachedData.IslamData.QuranChapters(CInt(Convert.Attribute("index").Value) - 1).Name) + ArabicData.PopDirectionalFormatting + ")" + If(SchemeType = ArabicData.TranslitScheme.None, String.Empty, " " + Arabic.TransliterateToScheme(Arabic.TransliterateFromBuckwalter(CachedData.IslamData.QuranChapters(CInt(Convert.Attribute("index").Value) - 1).Name), SchemeType, Scheme, CachedData.RuleMetas("Normal"))), CInt(Convert.Attribute("order").Value)})).ToArray()
        Array.Sort(Names, New Utility.CompareNameValueArray)
        Return Names
    End Function
    Public Shared Function GetChapterIndexByRevelationOrder(ByVal Index As Integer) As Xml.Linq.XElement
        For Each ChapterNode As Xml.Linq.XElement In Utility.GetChildNode("suras", New List(Of Xml.Linq.XElement)(CachedData.XMLDocInfo.Root.Elements).ToArray()).Elements
            If ChapterNode.Name = "sura" AndAlso CInt(ChapterNode.Attribute("order").Value) = Index Then
                Return ChapterNode
            End If
        Next
        Return Nothing
    End Function
    Public Shared Function GetPartNames() As Array()
        Dim Names() As Array = New List(Of Object())(Linq.Enumerable.Select(Utility.GetChildNodes("juz", New List(Of Xml.Linq.XElement)(Utility.GetChildNode("juzs", New List(Of Xml.Linq.XElement)(CachedData.XMLDocInfo.Root.Elements).ToArray()).Elements).ToArray()), Function(Convert As Xml.Linq.XElement) New Object() {Convert.Attribute("index").Value + " (" + Arabic.TransliterateFromBuckwalter("juz " + CachedData.IslamData.QuranParts(CInt(Convert.Attribute("index").Value) - 1).Name + " ") + ")", CInt(Convert.Attribute("index").Value)})).ToArray()
        Array.Sort(Names, New Utility.CompareNameValueArray)
        Return Names
    End Function
    Public Shared Function GetPartCount() As Integer
        Return Utility.GetChildNodeCount("juz", Utility.GetChildNode("juzs", New List(Of Xml.Linq.XElement)(CachedData.XMLDocInfo.Root.Elements).ToArray()))
    End Function
    Public Shared Function GetPartByIndex(ByVal Index As Integer) As Xml.Linq.XElement
        Return Utility.GetChildNodeByIndex("juz", "index", Index, New List(Of Xml.Linq.XElement)(Utility.GetChildNode("juzs", New List(Of Xml.Linq.XElement)(CachedData.XMLDocInfo.Root.Elements).ToArray()).Elements).ToArray())
    End Function
    Public Shared Function GetGroupNames() As Array()
        Dim Names() As Array = New List(Of Object())(Linq.Enumerable.Select(Utility.GetChildNodes("quarter", New List(Of Xml.Linq.XElement)(Utility.GetChildNode("hizbs", New List(Of Xml.Linq.XElement)(CachedData.XMLDocInfo.Root.Elements).ToArray()).Elements).ToArray()), Function(Convert As Xml.Linq.XElement) New Object() {Convert.Attribute("index").Value, CInt(Convert.Attribute("index").Value)})).ToArray()
        Array.Sort(Names, New Utility.CompareNameValueArray)
        Return Names
    End Function
    Public Shared Function GetGroupCount() As Integer
        Return Utility.GetChildNodeCount("quarter", Utility.GetChildNode("hizbs", New List(Of Xml.Linq.XElement)(CachedData.XMLDocInfo.Root.Elements).ToArray()))
    End Function
    Public Shared Function GetGroupByIndex(ByVal Index As Integer) As Xml.Linq.XElement
        Return Utility.GetChildNodeByIndex("quarter", "index", Index, New List(Of Xml.Linq.XElement)(Utility.GetChildNode("hizbs", New List(Of Xml.Linq.XElement)(CachedData.XMLDocInfo.Root.Elements).ToArray()).Elements).ToArray())
    End Function
    Public Shared Function GetStationNames() As Array()
        Dim Names() As Array = New List(Of Object())(Linq.Enumerable.Select(Utility.GetChildNodes("manzil", New List(Of Xml.Linq.XElement)(Utility.GetChildNode("manzils", New List(Of Xml.Linq.XElement)(CachedData.XMLDocInfo.Root.Elements).ToArray()).Elements).ToArray()), Function(Convert As Xml.Linq.XElement) New Object() {Convert.Attribute("index").Value, CInt(Convert.Attribute("index").Value)})).ToArray()
        Array.Sort(Names, New Utility.CompareNameValueArray)
        Return Names
    End Function
    Public Shared Function GetStationCount() As Integer
        Return Utility.GetChildNodeCount("manzil", Utility.GetChildNode("manzils", New List(Of Xml.Linq.XElement)(CachedData.XMLDocInfo.Root.Elements).ToArray()))
    End Function
    Public Shared Function GetStationByIndex(ByVal Index As Integer) As Xml.Linq.XElement
        Return Utility.GetChildNodeByIndex("manzil", "index", Index, New List(Of Xml.Linq.XElement)(Utility.GetChildNode("manzils", New List(Of Xml.Linq.XElement)(CachedData.XMLDocInfo.Root.Elements).ToArray()).Elements).ToArray())
    End Function
    Public Shared Function GetSectionNames() As Array()
        Dim Names() As Array = New List(Of Object())(Linq.Enumerable.Select(Utility.GetChildNodes("ruku", New List(Of Xml.Linq.XElement)(Utility.GetChildNode("rukus", New List(Of Xml.Linq.XElement)(CachedData.XMLDocInfo.Root.Elements).ToArray()).Elements).ToArray()), Function(Convert As Xml.Linq.XElement) New Object() {Convert.Attribute("index").Value, CInt(Convert.Attribute("index").Value)})).ToArray()
        Array.Sort(Names, New Utility.CompareNameValueArray)
        Return Names
    End Function
    Public Shared Function GetSectionCount() As Integer
        Return Utility.GetChildNodeCount("ruku", Utility.GetChildNode("rukus", New List(Of Xml.Linq.XElement)(CachedData.XMLDocInfo.Root.Elements).ToArray()))
    End Function
    Public Shared Function GetSectionByIndex(ByVal Index As Integer) As Xml.Linq.XElement
        Return Utility.GetChildNodeByIndex("ruku", "index", Index, New List(Of Xml.Linq.XElement)(Utility.GetChildNode("rukus", New List(Of Xml.Linq.XElement)(CachedData.XMLDocInfo.Root.Elements).ToArray()).Elements).ToArray())
    End Function
    Public Shared Function GetPageNames() As Array()
        Dim Names() As Array = New List(Of Object())(Linq.Enumerable.Select(Utility.GetChildNodes("page", New List(Of Xml.Linq.XElement)(Utility.GetChildNode("pages", New List(Of Xml.Linq.XElement)(CachedData.XMLDocInfo.Root.Elements).ToArray()).Elements).ToArray()), Function(Convert As Xml.Linq.XElement) New Object() {Convert.Attribute("index").Value, CInt(Convert.Attribute("index").Value)})).ToArray()
        Array.Sort(Names, New Utility.CompareNameValueArray)
        Return Names
    End Function
    Public Shared Function GetPageCount() As Integer
        Return Utility.GetChildNodeCount("page", Utility.GetChildNode("pages", New List(Of Xml.Linq.XElement)(CachedData.XMLDocInfo.Root.Elements).ToArray()))
    End Function
    Public Shared Function GetPageByIndex(ByVal Index As Integer) As Xml.Linq.XElement
        Return Utility.GetChildNodeByIndex("page", "index", Index, New List(Of Xml.Linq.XElement)(Utility.GetChildNode("pages", New List(Of Xml.Linq.XElement)(CachedData.XMLDocInfo.Root.Elements).ToArray()).Elements).ToArray())
    End Function
    Public Shared Function GetSajdaNames() As Array()
        Dim Names() As Array = New List(Of Object())(Linq.Enumerable.Select(Utility.GetChildNodes("sajda", New List(Of Xml.Linq.XElement)(Utility.GetChildNode("sajdas", New List(Of Xml.Linq.XElement)(CachedData.XMLDocInfo.Root.Elements).ToArray()).Elements).ToArray()), Function(Convert As Xml.Linq.XElement) New Object() {Convert.Attribute("index").Value, CInt(Convert.Attribute("index").Value)})).ToArray()
        Array.Sort(Names, New Utility.CompareNameValueArray)
        Return Names
    End Function
    Public Shared Function GetSajdaCount() As Integer
        Return Utility.GetChildNodeCount("sajda", Utility.GetChildNode("sajdas", New List(Of Xml.Linq.XElement)(CachedData.XMLDocInfo.Root.Elements).ToArray()))
    End Function
    Public Shared Function GetSajdaByIndex(ByVal Index As Integer) As Xml.Linq.XElement
        Return Utility.GetChildNodeByIndex("sajda", "index", Index, New List(Of Xml.Linq.XElement)(Utility.GetChildNode("sajdas", New List(Of Xml.Linq.XElement)(CachedData.XMLDocInfo.Root.Elements).ToArray()).Elements).ToArray())
    End Function
End Class
Public Interface IslamSiteRankingData
    Function GetHadithRankingData(ByVal Collection As String, ByVal Book As Integer, ByVal Hadith As Integer) As Integer()
    Function GetUserHadithRankingData(ByVal Collection As String, ByVal Book As Integer, ByVal Hadith As Integer) As Integer
End Interface
Public Class HadithReader
    Public Shared Function GetCollectionIndex(ByVal Name As String) As Integer
        Dim Count As Integer
        For Count = 0 To CachedData.IslamData.Collections.Length - 1
            If Name = CachedData.IslamData.Collections(Count).Name Then Return Count
        Next
        Return -1
    End Function
    Public Shared Function GetCollectionNames() As String()
        Return New List(Of String)(Linq.Enumerable.Select(CachedData.IslamData.Collections, Function(Convert As IslamData.CollectionInfo) Utility.LoadResourceString("IslamInfo_" + Convert.Name))).ToArray()
    End Function
    Public Shared Function GetChapterByIndex(ByVal BookNode As Xml.Linq.XElement, ByVal ChapterIndex As Integer) As Xml.Linq.XElement
        Return Utility.GetChildNodeByIndex("chapter", "index", ChapterIndex, New List(Of Xml.Linq.XElement)(BookNode.Elements).ToArray())
    End Function
    Public Shared Function GetSubChapterByIndex(ByVal ChapterNode As Xml.Linq.XElement, ByVal SubChapterIndex As Integer) As Xml.Linq.XElement
        Return Utility.GetChildNodeByIndex("subchapter", "index", SubChapterIndex, New List(Of Xml.Linq.XElement)(ChapterNode.Elements).ToArray())
    End Function
    Public Shared Function GetBookEName(ByVal BookNode As Xml.Linq.XElement, CollectionIndex As Integer) As String
        If BookNode Is Nothing Then
            Return String.Empty
        Else
            GetBookEName = Utility.LoadResourceString("IslamInfo_" + CachedData.IslamData.Collections(CollectionIndex).FileName + "Book" + BookNode.Attribute("index").Value)
            If GetBookEName Is Nothing Then GetBookEName = String.Empty
        End If
    End Function
    Public Shared Function GetTranslationList(Collection As Integer) As Array()
        Return New List(Of String())(Linq.Enumerable.Select(CachedData.IslamData.Collections(Collection).Translations, Function(Convert As IslamData.CollectionInfo.CollTranslationInfo) New String() {Utility.LoadResourceString("lang_local" + Languages.GetLanguageInfoByCode(Convert.FileName.Substring(0, 2)).Code) + ": " + Utility.LoadResourceString("IslamInfo_" + Convert.Name), Convert.FileName})).ToArray()
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
    Public Shared Function GetHadithMappingText(Index As Integer, HadithTranslation As String) As Array()
        Dim XMLDocTranslate As Xml.Linq.XDocument
        If CachedData.IslamData.Collections(Index).Translations.Length = 0 Then Return New Array() {}
        Dim Stream As IO.Stream = PortableMethods.FileIO.LoadStream(PortableMethods.Settings.GetFilePath("metadata\" + GetTranslationXMLFileName(Index, HadithTranslation) + ".xml"))
        XMLDocTranslate = Xml.Linq.XDocument.Load(Stream)
        Stream.Dispose()
        Dim Output As New List(Of Object)
        Output.Add(New String() {})
        If HadithReader.HasVolumes(Index) Then
            Output.Add(New String() {String.Empty, String.Empty, String.Empty, String.Empty, String.Empty, String.Empty})
            Output.Add(New String() {Utility.LoadResourceString("Hadith_Volume"), Utility.LoadResourceString("Hadith_Book"), Utility.LoadResourceString("Hadith_Index"), Utility.LoadResourceString("Hadith_Chapters"), Utility.LoadResourceString("Hadith_Hadiths"), Utility.LoadResourceString("Hadith_Translation")})
        Else
            Output.Add(New String() {String.Empty, String.Empty, String.Empty, String.Empty, String.Empty})
            Output.Add(New String() {Utility.LoadResourceString("Hadith_Book"), Utility.LoadResourceString("Hadith_Index"), Utility.LoadResourceString("Hadith_Chapters"), Utility.LoadResourceString("Hadith_Hadiths"), Utility.LoadResourceString("Hadith_Translation")})
        End If
        If HadithReader.HasVolumes(Index) Then
            Output.AddRange(Linq.Enumerable.Select(Utility.GetChildNodes("book", New List(Of Xml.Linq.XElement)(Utility.GetChildNode("books", New List(Of Xml.Linq.XElement)(CachedData.XMLDocInfos(Index).Root.Elements).ToArray()).Elements).ToArray()), Function(Convert As Xml.Linq.XElement) New Object() {CStr(HadithReader.GetVolumeIndex(Index, CInt(Convert.Attribute("index").Value))), GetBookEName(Convert, Index), Convert.Attribute("index").Value, CStr(HadithReader.GetChapterCount(Index, CInt(Convert.Attribute("index").Value))), CStr(HadithReader.GetHadithCount(Index, CInt(Convert.Attribute("index").Value))), HadithReader.GetBookHadithMapping(XMLDocTranslate, Index, CInt(Convert.Attribute("index").Value))}))
        Else
            Output.AddRange(Linq.Enumerable.Select(Utility.GetChildNodes("book", New List(Of Xml.Linq.XElement)(Utility.GetChildNode("books", New List(Of Xml.Linq.XElement)(CachedData.XMLDocInfos(Index).Root.Elements).ToArray()).Elements).ToArray()), Function(Convert As Xml.Linq.XElement) New Object() {GetBookEName(Convert, Index), Convert.Attribute("index").Value, CStr(HadithReader.GetChapterCount(Index, CInt(Convert.Attribute("index").Value))), CStr(HadithReader.GetHadithCount(Index, CInt(Convert.Attribute("index").Value))), HadithReader.GetBookHadithMapping(XMLDocTranslate, Index, CInt(Convert.Attribute("index").Value))}))
        End If
        Return DirectCast(Output.ToArray(), Array())
    End Function
    Public Shared Function DoGetRenderedText(ID As String, SchemeType As ArabicData.TranslitScheme, Scheme As String, Translation As String, Index As Integer, BookIndex As Integer, RankInterface As IslamSiteRankingData) As RenderArray
        Dim Renderer As New RenderArray(ID)
        Dim Hadith As Integer
        Dim HadithText As Collections.Generic.List(Of Collections.Generic.List(Of Object)) = HadithReader.GetHadithText(BookIndex, Index)
        Dim ChapterIndex As Integer = -1
        Dim SubChapterIndex As Integer = -1
        Dim BookNode As Xml.Linq.XElement = HadithReader.GetBookByIndex(Index, BookIndex)
        Dim ChapterNode As Xml.Linq.XElement = Nothing
        Dim SubChapterNode As Xml.Linq.XElement
        If Not BookNode Is Nothing Then
            Renderer.Items.Add(New RenderArray.RenderItem(RenderArray.RenderTypes.eHeaderLeft, New RenderArray.RenderText() {New RenderArray.RenderText(RenderArray.RenderDisplayClass.eArabic, Arabic.TransliterateFromBuckwalter("Had~iv " + BookNode.Attribute("hadiths").Value + " ")), New RenderArray.RenderText(RenderArray.RenderDisplayClass.eTransliteration, Arabic.TransliterateToScheme(Arabic.TransliterateFromBuckwalter("Had~iv " + BookNode.Attribute("hadiths").Value + " "), SchemeType, Scheme, CachedData.RuleMetas("Normal")).Trim()), New RenderArray.RenderText(RenderArray.RenderDisplayClass.eLTR, "Hadiths: " + BookNode.Attribute("hadiths").Value + " ")}))
            Renderer.Items.Add(New RenderArray.RenderItem(RenderArray.RenderTypes.eHeaderCenter, New RenderArray.RenderText() {New RenderArray.RenderText(RenderArray.RenderDisplayClass.eArabic, Arabic.TransliterateFromBuckwalter("{lokita`bu " + CStr(BookIndex)) + " " + BookNode.Attribute("name").Value + " "), New RenderArray.RenderText(RenderArray.RenderDisplayClass.eTransliteration, Arabic.TransliterateToScheme(Arabic.TransliterateFromBuckwalter("{lokita`bu " + CStr(BookIndex)) + " " + BookNode.Attribute("name").Value + " ", SchemeType, Scheme, CachedData.RuleMetas("Normal")).Trim()), New RenderArray.RenderText(RenderArray.RenderDisplayClass.eLTR, "Book " + CStr(BookIndex) + ": " + GetBookEName(BookNode, Index) + " ")}))
            'Renderer.Items.Add(New RenderArray.RenderItem(RenderArray.RenderTypes.eHeaderRight, New RenderArray.RenderText() {New RenderArray.RenderText(RenderArray.RenderDisplayClass.eArabic, Arabic.TransliterateFromBuckwalter("mjld " + New List(Of Xml.Linq.XElement)(Utility.GetChildNode("books", New List(Of Xml.Linq.XElement)(XMLDocInfo(Index).Root.Elements).ToArray()).Elements).ToArray()(BookIndex).Attribute("volume").Value + " ")), New RenderArray.RenderText(RenderArray.RenderDisplayClass.eTransliteration, Arabic.TransliterateToScheme(Arabic.TransliterateFromBuckwalter("mjld " + New List(Of Xml.Linq.XElement)(Utility.GetChildNode("books", New List(Of Xml.Linq.XElement)(XMLDocInfo(Index).Root.Elements).ToArray()).Elements).ToArray()(BookIndex).Attribute("volume").Value + " "), SchemeType, Scheme).Trim()), New RenderArray.RenderText(RenderArray.RenderDisplayClass.eLTR, "Volume " + Utility.GetChildNode("books", XMLDocInfo(Index).Root.Nodes).Nodes.Item(BookIndex).Attribute("volume").Value + " ")}))
            Dim XMLDocTranslate As Xml.Linq.XDocument = Nothing
            Dim Strings() As String = Nothing
            If CachedData.IslamData.Collections(Index).Translations.Length <> 0 Then
                Dim Stream As IO.Stream = PortableMethods.FileIO.LoadStream(PortableMethods.Settings.GetFilePath("metadata\" + GetTranslationXMLFileName(Index, Translation) + ".xml"))
                XMLDocTranslate = Xml.Linq.XDocument.Load(Stream)
                Stream.Dispose()
                Strings = Utility.ReadAllLines(PortableMethods.Settings.GetFilePath("metadata\" + GetTranslationFileName(Index, Translation) + ".txt"))
            End If
            For Hadith = 0 To HadithText.Count - 1
                'Handle missing or excess chapter indexes
                If ChapterIndex <> CInt(HadithText(Hadith)(1)) Then
                    ChapterIndex = CInt(HadithText(Hadith)(1))
                    ChapterNode = GetChapterByIndex(BookNode, ChapterIndex)
                    If Not ChapterNode Is Nothing Then
                        Renderer.Items.Add(New RenderArray.RenderItem(RenderArray.RenderTypes.eHeaderLeft, New RenderArray.RenderText() {New RenderArray.RenderText(RenderArray.RenderDisplayClass.eArabic, Arabic.TransliterateFromBuckwalter("Had~iv " + ChapterNode.Attribute("hadiths").Value + " ")), New RenderArray.RenderText(RenderArray.RenderDisplayClass.eTransliteration, Arabic.TransliterateToScheme(Arabic.TransliterateFromBuckwalter("Had~iv " + ChapterNode.Attribute("hadiths").Value + " "), SchemeType, Scheme, CachedData.RuleMetas("Normal")).Trim()), New RenderArray.RenderText(RenderArray.RenderDisplayClass.eLTR, "Hadiths: " + ChapterNode.Attribute("hadiths").Value + " ")}))
                        Dim Heads As New List(Of RenderArray.RenderText)
                        Heads.AddRange(Linq.Enumerable.Select(Of String, RenderArray.RenderText)(System.Text.RegularExpressions.Regex.Split(Arabic.TransliterateFromBuckwalter("bAb " + CStr(ChapterIndex)) + " " + ChapterNode.Attribute("name").Value + " ", "(\d+\.\d+(?:-\d+)?)"), Function(Str As String) If(System.Text.RegularExpressions.Regex.Match(Str, "(\d+)\.(\d+(?:-\d+)?)").Success, New RenderArray.RenderText(RenderArray.RenderDisplayClass.eLink, {"/host.aspx?Page=docbuild&docedit=%7B" + Str.Replace(".", "%3A") + "%7D&selectiondisplay=Display&translitscheme=0&fontselection=def&fontcustom=Lotus", Str}), New RenderArray.RenderText(RenderArray.RenderDisplayClass.eArabic, Str))))
                        Heads.Add(New RenderArray.RenderText(RenderArray.RenderDisplayClass.eTransliteration, Arabic.TransliterateToScheme(Arabic.TransliterateFromBuckwalter("bAb " + CStr(ChapterIndex)) + " " + System.Text.RegularExpressions.Regex.Replace(ChapterNode.Attribute("name").Value, "(\d+).(\d+(?:-\d+)?)", String.Empty) + " ", SchemeType, Scheme, CachedData.RuleMetas("Normal")).Trim()))
                        Heads.AddRange(Linq.Enumerable.Select(Of String, RenderArray.RenderText)(System.Text.RegularExpressions.Regex.Split("Chapter " + CStr(ChapterIndex) + ": " + Utility.DefaultValue(Utility.LoadResourceString("IslamInfo_" + CachedData.IslamData.Collections(Index).FileName + "Book" + BookNode.Attribute("index").Value + "Chapter" + ChapterNode.Attribute("index").Value), String.Empty) + " ", "(\d+\.\d+(?:-\d+)?)"), Function(Str As String) If(System.Text.RegularExpressions.Regex.Match(Str, "(\d+)\.(\d+(?:-\d+)?)").Success, New RenderArray.RenderText(RenderArray.RenderDisplayClass.eLink, {"/host.aspx?Page=docbuild&docedit=%7B" + Str.Replace(".", "%3A") + "%7D&selectiondisplay=Display&translitscheme=0&fontselection=def&fontcustom=Lotus", Str}), New RenderArray.RenderText(RenderArray.RenderDisplayClass.eLTR, Str))))
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
                            Renderer.Items.Add(New RenderArray.RenderItem(RenderArray.RenderTypes.eHeaderLeft, New RenderArray.RenderText() {New RenderArray.RenderText(RenderArray.RenderDisplayClass.eArabic, Arabic.TransliterateFromBuckwalter("Had~iv " + SubChapterNode.Attribute("hadiths").Value + " ")), New RenderArray.RenderText(RenderArray.RenderDisplayClass.eTransliteration, Arabic.TransliterateToScheme(Arabic.TransliterateFromBuckwalter("Had~iv " + SubChapterNode.Attribute("hadiths").Value + " "), SchemeType, Scheme, CachedData.RuleMetas("Normal")).Trim()), New RenderArray.RenderText(RenderArray.RenderDisplayClass.eLTR, "Hadiths: " + SubChapterNode.Attribute("hadiths").Value + " ")}))
                            Dim Heads As New List(Of RenderArray.RenderText)
                            Heads.AddRange(Linq.Enumerable.Select(Of String, RenderArray.RenderText)(System.Text.RegularExpressions.Regex.Split(Arabic.TransliterateFromBuckwalter("bAb " + CStr(SubChapterIndex)) + " " + SubChapterNode.Attribute("name").Value + " ", "(\d+\.\d+(?:-\d+)?)"), Function(Str As String) If(System.Text.RegularExpressions.Regex.Match(Str, "(\d+)\.(\d+(?:-\d+)?)").Success, New RenderArray.RenderText(RenderArray.RenderDisplayClass.eLink, {"/host.aspx?Page=docbuild&docedit=%7B" + Str.Replace(".", "%3A") + "%7D&selectiondisplay=Display&translitscheme=0&fontselection=def&fontcustom=Lotus", Str}), New RenderArray.RenderText(RenderArray.RenderDisplayClass.eArabic, Str))))
                            Heads.Add(New RenderArray.RenderText(RenderArray.RenderDisplayClass.eTransliteration, Arabic.TransliterateToScheme(Arabic.TransliterateFromBuckwalter("bAb " + CStr(SubChapterIndex)) + " " + System.Text.RegularExpressions.Regex.Replace(SubChapterNode.Attribute("name").Value, "(\d+).(\d+(?:-\d+)?)", String.Empty) + " ", SchemeType, Scheme, CachedData.RuleMetas("Normal")).Trim()))
                            Heads.AddRange(Linq.Enumerable.Select(Of String, RenderArray.RenderText)(System.Text.RegularExpressions.Regex.Split("Sub-Chapter " + CStr(SubChapterIndex) + ": " + Utility.DefaultValue(Utility.LoadResourceString("IslamInfo_" + CachedData.IslamData.Collections(Index).FileName + "Book" + BookNode.Attribute("index").Value + "Chapter" + ChapterNode.Attribute("index").Value + "Subchapter" + SubChapterNode.Attribute("index").Value), String.Empty) + " ", "(\d+\.\d+(?:-\d+)?)"), Function(Str As String) If(System.Text.RegularExpressions.Regex.Match(Str, "(\d+)\.(\d+(?:-\d+)?)").Success, New RenderArray.RenderText(RenderArray.RenderDisplayClass.eLink, {"/host.aspx?Page=docbuild&docedit=%7B" + Str.Replace(".", "%3A") + "%7D&selectiondisplay=Display&translitscheme=0&fontselection=def&fontcustom=Lotus", Str}), New RenderArray.RenderText(RenderArray.RenderDisplayClass.eLTR, Str))))
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
                Texts.AddRange(Linq.Enumerable.Select(Of String, RenderArray.RenderText)(System.Text.RegularExpressions.Regex.Split(CStr(HadithText(Hadith)(3)) + " " + Arabic.TransliterateFromBuckwalter("=" + CStr(HadithText(Hadith)(0)) + " "), "(\d+\.\d+(?:-\d+)?)"), Function(Str As String) If(System.Text.RegularExpressions.Regex.Match(Str, "(\d+)\.(\d+(?:-\d+)?)").Success, New RenderArray.RenderText(RenderArray.RenderDisplayClass.eLink, {"/host.aspx?Page=docbuild&docedit=%7B" + Str.Replace(".", "%3A") + "%7D&selectiondisplay=Display&translitscheme=0&fontselection=def&fontcustom=Lotus", Str}), New RenderArray.RenderText(RenderArray.RenderDisplayClass.eArabic, Str))))
                Texts.Add(New RenderArray.RenderText(RenderArray.RenderDisplayClass.eTransliteration, Arabic.TransliterateToScheme(System.Text.RegularExpressions.Regex.Replace(CStr(HadithText(Hadith)(3)), "(\d+).(\d+(?:-\d+)?)", String.Empty) + " " + Arabic.TransliterateFromBuckwalter("=" + CStr(HadithText(Hadith)(0))) + " ", SchemeType, Scheme, CachedData.RuleMetas("Normal")).Trim()))
                Texts.AddRange(Linq.Enumerable.Select(Of String, RenderArray.RenderText)(System.Text.RegularExpressions.Regex.Split("(" + CStr(HadithText(Hadith)(0)) + ") " + HadithTranslation, "(\d+\.\d+(?:-\d+)?)"), Function(Str As String) If(System.Text.RegularExpressions.Regex.Match(Str, "(\d+)\.(\d+(?:-\d+)?)").Success, New RenderArray.RenderText(RenderArray.RenderDisplayClass.eLink, {"/host.aspx?Page=docbuild&docedit=%7B" + Str.Replace(".", "%3A") + "%7D&selectiondisplay=Display&translitscheme=0&fontselection=def&fontcustom=Lotus", Str}), New RenderArray.RenderText(If(IsTranslationTextLTR(Index, Translation), RenderArray.RenderDisplayClass.eLTR, RenderArray.RenderDisplayClass.eRTL), Str))))
                Renderer.Items.Add(New RenderArray.RenderItem(RenderArray.RenderTypes.eText, Texts.ToArray()))
                Dim Ranking As Integer() = RankInterface.GetHadithRankingData(CachedData.IslamData.Collections(Index).FileName, BookIndex, CInt(HadithText(Hadith)(0)))
                Dim UserRanking As Integer = RankInterface.GetUserHadithRankingData(CachedData.IslamData.Collections(Index).FileName, BookIndex, CInt(HadithText(Hadith)(0)))
                Renderer.Items.Add(New RenderArray.RenderItem(RenderArray.RenderTypes.eInteractive, New RenderArray.RenderText() {New RenderArray.RenderText(RenderArray.RenderDisplayClass.eRanking, CachedData.IslamData.Collections(Index).FileName + "|" + CStr(BookIndex) + "|" + CStr(HadithText(Hadith)(0)) + "|" + CStr(Ranking(0)) + "|" + CStr(Ranking(1)) + "|" + CStr(UserRanking))}))
            Next
        End If
        Return Renderer
    End Function
    Public Shared Function GetHadithTextBook(ByVal XMLDocMain As Xml.Linq.XDocument, ByVal BookIndex As Integer) As Xml.Linq.XElement
        Return Utility.GetChildNodeByIndex("book", "index", BookIndex, New List(Of Xml.Linq.XElement)(XMLDocMain.Root.Elements).ToArray())
    End Function
    Public Shared Function GetHadithText(ByVal BookIndex As Integer, Collection As Integer) As Collections.Generic.List(Of Collections.Generic.List(Of Object))
        Dim XMLDocMain As Xml.Linq.XDocument
        Dim Stream As IO.Stream = PortableMethods.FileIO.LoadStream(PortableMethods.Settings.GetFilePath("metadata\" + CachedData.IslamData.Collections(Collection).FileName + ".xml"))
        XMLDocMain = Xml.Linq.XDocument.Load(Stream)
        Stream.Dispose()
        Dim BookNode As Xml.Linq.XElement = GetHadithTextBook(XMLDocMain, BookIndex)
        Dim Hadiths As New Collections.Generic.List(Of Collections.Generic.List(Of Object))
        For Each HadithNode As Xml.Linq.XElement In BookNode.Elements
            If HadithNode.Name = "hadith" Then
                Dim NextEntry As New Collections.Generic.List(Of Object)
                NextEntry.AddRange(New Object() {CInt(HadithNode.Attribute("index").Value), _
                                              CInt(Utility.ParseValue(HadithNode.Attribute("sectionindex"), "-1")), _
                                              CInt(Utility.ParseValue(HadithNode.Attribute("subsectionindex"), "-1")), _
                                              HadithNode.Attribute("text").Value})
                Hadiths.Add(NextEntry)
            End If
        Next
        Return Hadiths
    End Function
    Public Shared Function GetBookCount(ByVal Index As Integer) As Integer
        Return CInt(Utility.GetChildNode("books", New List(Of Xml.Linq.XElement)(CachedData.XMLDocInfos(Index).Root.Elements).ToArray()).Attribute("count").Value)
    End Function
    Public Shared Function GetBookByIndex(ByVal Index As Integer, ByVal BookIndex As Integer) As Xml.Linq.XElement
        Return Utility.GetChildNodeByIndex("book", "index", BookIndex, New List(Of Xml.Linq.XElement)(Utility.GetChildNode("books", New List(Of Xml.Linq.XElement)(CachedData.XMLDocInfos(Index).Root.Elements).ToArray()).Elements).ToArray())
    End Function
    Public Shared Function GetBookNamesByCollection(ByVal Index As Integer) As Array()
        Dim Names() As Array = New List(Of Object())(Linq.Enumerable.Select(Utility.GetChildNodes("book", New List(Of Xml.Linq.XElement)(Utility.GetChildNode("books", New List(Of Xml.Linq.XElement)(CachedData.XMLDocInfos(Index).Root.Elements).ToArray()).Elements).ToArray()), Function(Convert As Xml.Linq.XElement) New Object() {Convert.Attribute("index").Value + ". " + GetBookEName(Convert, Index), CInt(Convert.Attribute("index").Value)})).ToArray()
        Array.Sort(Names, New Utility.CompareNameValueArray)
        Return Names
    End Function
    Public Shared Function HasVolumes(ByVal Index As Integer) As Boolean
        Return Not Utility.GetChildNode("books", New List(Of Xml.Linq.XElement)(CachedData.XMLDocInfos(Index).Root.Elements).ToArray()).Attributes("volumes") Is Nothing
    End Function
    Public Shared Function GetVolumeIndex(ByVal Index As Integer, ByVal BookIndex As Integer) As Integer
        Dim Node As Xml.Linq.XAttribute = GetBookByIndex(Index, BookIndex).Attribute("volume")
        If Node Is Nothing Then Return -1
        Return CInt(Node.Value)
    End Function
    Public Shared Function GetChapterCount(ByVal Index As Integer, ByVal BookIndex As Integer) As Integer
        Dim Node As Xml.Linq.XAttribute = GetBookByIndex(Index, BookIndex).Attribute("chapters")
        If Node Is Nothing Then Return -1
        Return CInt(Node.Value)
    End Function
    Public Shared Function GetChapterIndex(ByVal Index As Integer, ByVal BookIndex As Integer, ByVal HadithIndex As Integer) As Integer
        Dim BookNode As Xml.Linq.XElement = GetBookByIndex(Index, BookIndex)
        For Each ChapterNode As Xml.Linq.XElement In BookNode.Elements
            If ChapterNode.Name = "chapter" AndAlso _
                (CInt(ChapterNode.Attribute("starthadith").Value) <= HadithIndex And _
                CInt(ChapterNode.Attribute("starthadith").Value) + CInt(ChapterNode.Attribute("hadiths ").Value) > HadithIndex) Then Return New List(Of Xml.Linq.XElement)(ChapterNode.ElementsBeforeSelf).ToArray().Length
        Next
        Return -1
    End Function
    Public Shared Function GetHadithCount(ByVal Index As Integer, ByVal BookIndex As Integer) As Integer
        Dim Node As Xml.Linq.XAttribute = GetBookByIndex(Index, BookIndex).Attribute("hadiths")
        If Node Is Nothing Then Return -1
        Return CInt(Node.Value)
    End Function
    Public Shared Function GetHadithStart(ByVal Index As Integer, ByVal BookIndex As Integer) As Integer
        Dim Node As Xml.Linq.XAttribute = GetBookByIndex(Index, BookIndex).Attribute("starthadith")
        If Node Is Nothing Then Return -1
        Return CInt(Node.Value)
    End Function
    Public Shared Function ParseBookTranslationIndex(ByVal BookString As String) As Integer()
        Return New List(Of Integer)(Linq.Enumerable.Select(BookString.Split("|"c), Function(MakeNumeric As String) Integer.Parse(MakeNumeric))).ToArray()
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
                        Indexes.Add(Linq.Enumerable.Select(Combined, Function(MakeNumeric As String) Integer.Parse(MakeNumeric)))
                    End If
                ElseIf Ranges.Length = 2 Then
                    Dim Combined As String() = Ranges(1).Split("+"c)
                    For SubCount = Integer.Parse(Ranges(0)) To Integer.Parse(Combined(0))
                        If Combined.Length > 1 AndAlso SubCount = Integer.Parse(Combined(0)) Then
                            Indexes.Add(Linq.Enumerable.Select(Combined, Function(MakeNumeric As String) Integer.Parse(MakeNumeric)))
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
        ParseHadithTranslationIndex.AddRange(Linq.Enumerable.Select(HadithString.Split("|"c), Function(IndexString As String) ExpandIndexes(IndexString)))
    End Function
    Public Shared Function TranslationHasVolumes(ByVal XMLDocTranslate As Xml.Linq.XDocument) As Boolean
        Return Not Utility.GetChildNode("books", New List(Of Xml.Linq.XElement)(XMLDocTranslate.Root.Elements).ToArray()).Attributes("volumes") Is Nothing
    End Function
    Public Shared Function GetTranslateMaxHadith(ByVal XMLDocTranslate As Xml.Linq.XDocument) As Integer
        Dim MaxHadith As Integer = 0
        For Each BookNode As Xml.Linq.XElement In Utility.GetChildNode("books", New List(Of Xml.Linq.XElement)(XMLDocTranslate.Root.Elements).ToArray()).Elements
            If BookNode.Name = "book" Then
                If CInt(BookNode.Attribute("hadiths").Value) <> 0 Then
                    MaxHadith = Math.Max(MaxHadith, CInt(BookNode.Attribute("starthadith").Value) + CInt(BookNode.Attribute("hadiths").Value) - 1)
                End If
            End If
        Next
        Return MaxHadith
    End Function
    Public Shared Function GetMaxChapter(ByVal XMLDocTranslate As Xml.Linq.XDocument) As Integer
        Dim MaxChapter As Integer = 0
        For Each BookNode In Utility.GetChildNode("books", New List(Of Xml.Linq.XElement)(XMLDocTranslate.Root.Elements).ToArray()).Elements
            If BookNode.Name = "book" Then
                For Each ChapterNode As Xml.Linq.XElement In BookNode.Elements
                    If ChapterNode.Name = "chapter" Then
                        MaxChapter = Math.Max(MaxChapter, CInt(ChapterNode.Attribute("index").Value))
                    End If
                Next
            End If
        Next
        Return MaxChapter
    End Function
    Public Shared Function GetHadithChapter(ByVal BookNode As Xml.Linq.XElement, ByVal Hadith As Integer) As Integer
        Dim Node As Xml.Linq.XAttribute
        For Each ChapterNode As Xml.Linq.XElement In BookNode.Elements
            If ChapterNode.Name = "chapter" Then
                Node = ChapterNode.Attribute("starthadith")
                If Not Node Is Nothing AndAlso (Hadith >= CInt(Node.Value) And _
                    Hadith < CInt(Node.Value) + CInt(ChapterNode.Attribute("hadiths").Value)) Then
                    Return CInt(ChapterNode.Attribute("index").Value)
                End If
            End If
        Next
        Return -1
    End Function
    Public Shared Function BuildTranslationIndex(ByVal XMLDocTranslate As Xml.Linq.XDocument, ByVal Volume As Integer, ByVal Book As Integer, ByVal Chapter As Integer, ByVal Hadith As Integer, ByVal SharedHadith As Integer) As String
        Dim MaxVolume As Integer = 0
        Dim MaxBook As Integer = CInt(Utility.GetChildNode("books", New List(Of Xml.Linq.XElement)(XMLDocTranslate.Root.Elements).ToArray()).Attribute("count").Value)
        Dim MaxChapter As Integer
        Dim MaxHadith As Integer
        Dim bHasSharedHadith As Boolean = Not Utility.GetChildNode("books", New List(Of Xml.Linq.XElement)(XMLDocTranslate.Root.Elements).ToArray()).Attribute("sharedhadiths") Is Nothing
        MaxHadith = GetTranslateMaxHadith(XMLDocTranslate)
        Dim bHasChapters As Boolean = Not Utility.GetChildNode("books", New List(Of Xml.Linq.XElement)(XMLDocTranslate.Root.Elements).ToArray()).Attribute("chapters") Is Nothing
        If Chapter <> -1 Then
            MaxChapter = GetMaxChapter(XMLDocTranslate)
        End If
        If Volume <> -1 Then
            MaxVolume = CInt(Utility.GetChildNode("books", New List(Of Xml.Linq.XElement)(XMLDocTranslate.Root.Elements).ToArray()).Attribute("volumes").Value)
        End If
        Return CStr(If(Volume = -1, String.Empty, Utility.ZeroPad(CStr(Volume), Utility.GetDigitLength(MaxVolume)) + ".")) + Utility.ZeroPad(CStr(Book), Utility.GetDigitLength(MaxBook)) + "." + CStr(If(Chapter = -1, String.Empty, Utility.ZeroPad(CStr(Chapter), Utility.GetDigitLength(MaxChapter)) + ".")) + Utility.ZeroPad(CStr(Hadith), Utility.GetDigitLength(MaxHadith)) + CStr(If(bHasSharedHadith, If(SharedHadith = 0, If(Chapter = -1 Or Utility.GetChildNode("books", New List(Of Xml.Linq.XElement)(XMLDocTranslate.Root.Elements).ToArray()).Attribute("sourced") Is Nothing, " ", String.Empty), Strings.ChrW(64 + SharedHadith)), String.Empty)) + ":"
    End Function
    Public Shared Function MapIndexes(ByVal ExpandString As String, ByVal HadithIndex As Integer) As Object()
        Dim Count As Integer
        Dim SubCount As Integer
        Dim Groupings As String() = ExpandString.Split(","c)
        Dim Indexes As New List(Of String())
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
                        Compile += Combined(SubCount) + CStr(If(SubCount <> Combined.Length - 1, "&", String.Empty))
                    Next
                    Indexes.Add(New String() {CStr(HadithIndex), Compile})
                    HadithIndex += 1
                    HadithCount += Combined.Length
                ElseIf Ranges.Length = 2 Then
                    Dim Compile As String
                    Dim Combined As String() = Ranges(1).Split("+"c)
                    Compile = Ranges(0) + "-"
                    For SubCount = 0 To Combined.Length - 1
                        Compile += Combined(SubCount) + CStr(If(SubCount <> Combined.Length - 1, "&", String.Empty))
                    Next
                    Indexes.Add(New String() {CStr(HadithIndex) + "-" + CStr(HadithIndex + Integer.Parse(Combined(0)) - Integer.Parse(Ranges(0))), Compile})
                    HadithIndex += Integer.Parse(Combined(0)) - Integer.Parse(Ranges(0)) + 1
                    HadithCount += Integer.Parse(Combined(0)) - Integer.Parse(Ranges(0)) + Combined.Length
                End If
            End If
        Next
        Return New Object() {CStr(HadithCount), Indexes.ToArray()}
    End Function
    Public Shared Function GetBookHadithMapping(ByVal XMLDocTranslate As Xml.Linq.XDocument, ByVal Index As Integer, ByVal BookIndex As Integer) As Object()
        Dim SourceStart As Integer
        Dim Volume As Integer
        Dim Mapping As New List(Of Object())
        Mapping.Add(New String() {})
        If TranslationHasVolumes(XMLDocTranslate) Then
            Mapping.Add(New String() {String.Empty, String.Empty, String.Empty, String.Empty})
            Mapping.Add(New String() {Utility.LoadResourceString("Hadith_TranslationVolume"), Utility.LoadResourceString("Hadith_TranslationBook"), Utility.LoadResourceString("Hadith_TranslationHadithCount"), Utility.LoadResourceString("Hadith_TranslationMapping")})
        Else
            Mapping.Add(New String() {String.Empty, String.Empty, String.Empty})
            Mapping.Add(New String() {Utility.LoadResourceString("Hadith_TranslationBook"), Utility.LoadResourceString("Hadith_TranslationHadithCount"), Utility.LoadResourceString("Hadith_TranslationMapping")})
        End If
        Dim bHasSharedHadith As Boolean = Not Utility.GetChildNode("books", New List(Of Xml.Linq.XElement)(XMLDocTranslate.Root.Elements).ToArray()).Attribute("sharedhadiths") Is Nothing
        Dim Books() As Integer
        Dim TranslateBookIndex As Integer
        Dim Node As Xml.Linq.XAttribute
        For Each BookNode As Xml.Linq.XElement In Utility.GetChildNode("books", New List(Of Xml.Linq.XElement)(XMLDocTranslate.Root.Elements).ToArray()).Elements
            If BookNode.Name = "book" Then
                Node = BookNode.Attribute("sourcebook")
                If Node Is Nothing Then
                    Books = New Integer() {CInt(BookNode.Attribute("index").Value)}
                Else
                    Books = ParseBookTranslationIndex(Node.Value)
                End If
                Node = BookNode.Attribute("volume")
                If Node Is Nothing Then
                    Volume = -1
                Else
                    Volume = CInt(Node.Value)
                End If
                TranslateBookIndex = Array.FindIndex(Books, Function(CheckIndex As Integer) CheckIndex = BookIndex)
                If TranslateBookIndex <> -1 Then
                    'Must handle shared hadiths
                    Node = BookNode.Attribute("sourcestart")
                    If Node Is Nothing Then
                        If CInt(BookNode.Attribute("hadiths").Value) = 0 Then
                            SourceStart = 0
                        Else
                            SourceStart = CInt(BookNode.Attribute("starthadith").Value)
                        End If
                    Else
                        SourceStart = CInt(Node.Value.Split("|"c)(TranslateBookIndex))
                    End If
                    Node = BookNode.Attribute("sourceindex")
                    If Node Is Nothing OrElse Node.Value = String.Empty Then
                        If TranslationHasVolumes(XMLDocTranslate) Then
                            Mapping.Add(New String() {CStr(Volume), BookNode.Attribute("index").Value, _
                                        CStr(BookNode.Attribute("hadiths").Value), Utility.LoadResourceString("Hadith_IdenticalNumbering")})
                        Else
                            Mapping.Add(New String() {BookNode.Attribute("index").Value, _
                                        CStr(BookNode.Attribute("hadiths").Value), Utility.LoadResourceString("Hadith_IdenticalNumbering")})
                        End If
                    Else
                        Dim RetObject As Object()
                        RetObject = MapIndexes(Node.Value.Split("|"c)(TranslateBookIndex), SourceStart)
                        Dim SharedHadith As Integer
                        If bHasSharedHadith Then
                            SharedHadith = CInt(Utility.ParseValue(BookNode.Attribute("sharedhadiths"), "0"))
                        Else
                            SharedHadith = 0
                        End If
                        If TranslationHasVolumes(XMLDocTranslate) Then
                            Mapping.Add(New Object() {CStr(Volume), BookNode.Attribute("index").Value, _
                                        String.Format(Utility.LoadResourceString("Hadith_MappedOf"), CStr(RetObject(0)), CStr(CInt(BookNode.Attribute("hadiths").Value) + SharedHadith)), RetObject(1)})
                        Else
                            Mapping.Add(New Object() {BookNode.Attribute("index").Value, _
                                        String.Format(Utility.LoadResourceString("Hadith_MappedOf"), CStr(RetObject(0)), CStr(CInt(BookNode.Attribute("hadiths").Value) + SharedHadith)), RetObject(1)})
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
            For HadithCount = 0 To CInt(If(Count = TranslateBookIndex, HadithIndex, TranslationIndexes(Count).Count - 1))
                If TypeOf TranslationIndexes(Count)(HadithCount) Is Integer() Then
                    For ChildCount = 0 To CInt(If(Count = TranslateBookIndex And HadithCount = HadithIndex, SubCount - 1, DirectCast(TranslationIndexes(Count)(HadithCount), Integer()).Length - 1))
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
    Public Shared Function GetTranslationHadith(XMLDocTranslate As Xml.Linq.XDocument, Strings() As String, ByVal Index As Integer, ByVal BookIndex As Integer, ByVal HadithIndex As Integer) As String()
        Dim Node As Xml.Linq.XAttribute
        Dim Count As Integer
        Dim SubCount As Integer
        Dim SourceStart As Integer
        Dim Volume As Integer
        Dim Books() As Integer
        Dim TranslateBookIndex As Integer
        Dim TranslationIndexes As Collections.Generic.List(Of Collections.Generic.List(Of Object))
        Dim TranslationHadith As New List(Of String)
        If CachedData.IslamData.Collections(Index).Translations.Length = 0 Then Return New String() {}
        For Each BookNode As Xml.Linq.XElement In Utility.GetChildNode("books", New List(Of Xml.Linq.XElement)(XMLDocTranslate.Root.Elements).ToArray()).Elements
            If BookNode.Name = "book" Then
                Node = BookNode.Attribute("sourcebook")
                If Node Is Nothing Then
                    Books = New Integer() {Count + 1}
                Else
                    Books = ParseBookTranslationIndex(Node.Value)
                End If
                Node = BookNode.Attribute("volume")
                If Node Is Nothing Then
                    Volume = -1
                Else
                    Volume = CInt(Node.Value)
                End If
                TranslateBookIndex = Array.FindIndex(Books, Function(CheckIndex As Integer) CheckIndex = BookIndex + 1)
                If TranslateBookIndex <> -1 Then
                    Node = BookNode.Attribute("sourcestart")
                    If Node Is Nothing Then
                        SourceStart = HadithIndex
                    Else
                        SourceStart = CInt(Node.Value.Split("|"c)(TranslateBookIndex))
                    End If
                    Node = BookNode.Attribute("sourceindex")
                    If Node Is Nothing Then
                        TranslationHadith.Add(CStr(If(Volume = -1, String.Empty, Utility.LoadResourceString("Hadith_Volume") + ": " + CStr(Volume) + " ")) + Utility.LoadResourceString("Hadith_Books") + ": " + CStr(Books(TranslateBookIndex)) + " " + Utility.LoadResourceString("Hadith_Hadith") + ": " + CStr(HadithIndex))
                        TranslationHadith.AddRange(Utility.GetFileLinesByNumberPrefix(Strings, BuildTranslationIndex(XMLDocTranslate, Volume, CInt(BookNode.Attribute("index").Value), GetHadithChapter(BookNode, HadithIndex), HadithIndex, 0)))
                    Else
                        TranslationIndexes = ParseHadithTranslationIndex(Node.Value)
                        If HadithIndex >= SourceStart AndAlso HadithIndex - SourceStart < TranslationIndexes(TranslateBookIndex).Count Then
                            If TypeOf TranslationIndexes(TranslateBookIndex)(HadithIndex - SourceStart) Is Integer() Then
                                For SubCount = 0 To DirectCast(TranslationIndexes(TranslateBookIndex)(HadithIndex - SourceStart), Integer()).Length - 1
                                    Dim SharedHadithIndex As Integer = GetSharedHadithIndex(TranslationIndexes, TranslateBookIndex, HadithIndex - SourceStart, SubCount)
                                    TranslationHadith.Add(CStr(If(Volume = -1, String.Empty, Utility.LoadResourceString("Hadith_Volume") + ": " + CStr(Volume) + " ")) + Utility.LoadResourceString("Hadith_Book") + ": " + BookNode.Attribute("index").Value + " " + Utility.LoadResourceString("Hadith_Hadith") + ": " + CStr(DirectCast(TranslationIndexes(TranslateBookIndex)(HadithIndex - SourceStart), Integer())(SubCount)))
                                    TranslationHadith.AddRange(Utility.GetFileLinesByNumberPrefix(Strings, BuildTranslationIndex(XMLDocTranslate, Volume, CInt(BookNode.Attribute("index").Value), GetHadithChapter(BookNode, DirectCast(TranslationIndexes(TranslateBookIndex)(HadithIndex - SourceStart), Integer())(SubCount)), DirectCast(TranslationIndexes(TranslateBookIndex)(HadithIndex - SourceStart), Integer())(SubCount), SharedHadithIndex)))
                                Next
                            Else
                                If CInt(TranslationIndexes(TranslateBookIndex)(HadithIndex - SourceStart)) = -1 Then Return New String() {}
                                Dim SharedHadithIndex As Integer = GetSharedHadithIndex(TranslationIndexes, TranslateBookIndex, HadithIndex - SourceStart, -1)
                                TranslationHadith.Add(CStr(If(Volume = -1, String.Empty, Utility.LoadResourceString("Hadith_Volume") + ": " + CStr(Volume) + " ")) + Utility.LoadResourceString("Hadith_Book") + ": " + BookNode.Attribute("index").Value + " " + Utility.LoadResourceString("Hadith_Hadith") + ": " + CStr(TranslationIndexes(TranslateBookIndex)(HadithIndex - SourceStart)))
                                TranslationHadith.AddRange(Utility.GetFileLinesByNumberPrefix(Strings, BuildTranslationIndex(XMLDocTranslate, Volume, CInt(BookNode.Attribute("index").Value), GetHadithChapter(BookNode, CInt(TranslationIndexes(TranslateBookIndex)(HadithIndex - SourceStart))), CInt(TranslationIndexes(TranslateBookIndex)(HadithIndex - SourceStart)), SharedHadithIndex)))
                            End If
                        End If
                    End If
                End If
            End If
        Next
        Return DirectCast(TranslationHadith.ToArray(), String())
    End Function
End Class