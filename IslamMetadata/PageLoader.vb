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
            Compare = GetSchemeValueFromSymbol(DirectCast(x, ArabicData.ArabicXMLData.ArabicSymbol), _Scheme).Length - _
                GetSchemeValueFromSymbol(DirectCast(y, ArabicData.ArabicXMLData.ArabicSymbol), _Scheme).Length
        End Function
    End Class
    Public Shared Function TransliterateToScheme(ByVal ArabicString As String, SchemeType As ArabicData.TranslitScheme, Scheme As String) As String
        If SchemeType = ArabicData.TranslitScheme.RuleBased Then
            Return TransliterateWithRules(ArabicString, Scheme, Nothing)
        ElseIf SchemeType = ArabicData.TranslitScheme.Literal Then
            Return TransliterateToRoman(ArabicString, Scheme)
        Else
            Return New String(System.Array.FindAll(ArabicString.ToCharArray(), Function(Check As Char) Check = " "c))
        End If
    End Function
    Public Shared Function GetSchemeLongVowelFromString(Str As String, Scheme As String) As String
        Dim Sch As IslamData.TranslitScheme = Nothing
        Dim Count As Integer
        For Count = 0 To CachedData.IslamData.TranslitSchemes.Length - 1
            If CachedData.IslamData.TranslitSchemes(Count).Name = Scheme Then
                Sch = CachedData.IslamData.TranslitSchemes(Count)
                Exit For
            End If
        Next
        If Count = CachedData.IslamData.TranslitSchemes.Length Then Return String.Empty
        If Array.IndexOf(ArabicData.ArabicVowels, Str) <> -1 Then
            Return Sch.Vowels(Array.IndexOf(ArabicData.ArabicVowels, Str))
        End If
        Return String.Empty
    End Function
    Public Shared Function GetSchemeGutteralFromString(Str As String, Scheme As String, Leading As Boolean) As String
        Dim Sch As IslamData.TranslitScheme = Nothing
        Dim Count As Integer
        For Count = 0 To CachedData.IslamData.TranslitSchemes.Length - 1
            If CachedData.IslamData.TranslitSchemes(Count).Name = Scheme Then
                Sch = CachedData.IslamData.TranslitSchemes(Count)
                Exit For
            End If
        Next
        If Count = CachedData.IslamData.TranslitSchemes.Length Then Return String.Empty
        If Array.IndexOf(ArabicData.ArabicLeadingGutterals, Str) <> -1 Then
            Return Sch.Vowels(Array.IndexOf(ArabicData.ArabicLeadingGutterals, Str) + ArabicData.ArabicVowels.Length + If(Leading, ArabicData.ArabicLeadingGutterals.Length, 0))
        End If
        Return String.Empty
    End Function
    Public Shared Function GetSchemeValueFromSymbol(Symbol As ArabicData.ArabicXMLData.ArabicSymbol, Scheme As String) As String
        Dim Sch As IslamData.TranslitScheme = Nothing
        Dim Count As Integer
        For Count = 0 To CachedData.IslamData.TranslitSchemes.Length - 1
            If CachedData.IslamData.TranslitSchemes(Count).Name = Scheme Then
                Sch = CachedData.IslamData.TranslitSchemes(Count)
                Exit For
            End If
        Next
        If Count = CachedData.IslamData.TranslitSchemes.Length Then Return String.Empty
        If Array.IndexOf(ArabicData.ArabicLettersInOrder, Symbol.Symbol) <> -1 Then
            Return Sch.Alphabet(Array.IndexOf(ArabicData.ArabicLettersInOrder, Symbol.Symbol))
        ElseIf Array.IndexOf(ArabicData.ArabicHamzas, Symbol.Symbol) <> -1 Then
            Return Sch.Hamza(Array.IndexOf(ArabicData.ArabicHamzas, Symbol.Symbol))
        ElseIf Array.IndexOf(ArabicData.ArabicSpecialLetters, Symbol.Symbol) <> -1 Then
            Return Sch.SpecialLetters(Array.IndexOf(ArabicData.ArabicSpecialLetters, Symbol.Symbol))
        ElseIf Array.IndexOf(ArabicData.ArabicVowels, CStr(Symbol.Symbol)) <> -1 Then
            Return Sch.Vowels(Array.IndexOf(ArabicData.ArabicVowels, CStr(Symbol.Symbol)))
        ElseIf Array.IndexOf(ArabicData.ArabicTajweed, Symbol.Symbol) <> -1 Then
            Return Sch.Tajweed(Array.IndexOf(ArabicData.ArabicTajweed, Symbol.Symbol))
        ElseIf Array.IndexOf(ArabicData.ArabicPunctuation, Symbol.Symbol) <> -1 Then
            Return Sch.Punctuation(Array.IndexOf(ArabicData.ArabicPunctuation, Symbol.Symbol))
        ElseIf Array.IndexOf(ArabicData.NonArabicLetters, Symbol.Symbol) <> -1 Then
            Return Sch.NonArabic(Array.IndexOf(ArabicData.NonArabicLetters, Symbol.Symbol))
        End If
        Return String.Empty
    End Function
    Public Shared Function TransliterateToRoman(ByVal ArabicString As String, Scheme As String) As String
        Dim RomanString As String = String.Empty
        Dim Count As Integer
        Dim Index As Integer
        Dim Letters(ArabicData.Data.ArabicLetters.Length - 1) As ArabicData.ArabicXMLData.ArabicSymbol
        ArabicData.Data.ArabicLetters.CopyTo(Letters, 0)
        Array.Sort(Letters, New StringLengthComparer(Scheme))
        For Count = 0 To ArabicString.Length - 1
            If ArabicString(Count) = "\" Then
                Count += 1
                If ArabicString(Count) = "," Then
                    RomanString += ArabicData.ArabicComma
                ElseIf ArabicString(Count) = ";" Then
                    RomanString += ArabicData.ArabicSemicolon
                ElseIf ArabicString(Count) = "?" Then
                    RomanString += ArabicData.ArabicQuestionMark
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
        New RuleTranslation With {.Rule = "Feh", .Match = ArabicData.ArabicLetterFeh, .Evaluator = ChrW(&H6A2)},
        New RuleTranslation With {.Rule = "ImalaE", .Match = String.Empty, .Evaluator = ChrW(&H65C)}
    }
    Public Shared WarshOrthography As RuleTranslation() = { _
        New RuleTranslation With {.Rule = "FehBeginMiddle", .Match = ArabicData.MakeUniRegEx(ArabicData.ArabicLetterFeh) + "(" + ArabicData.MakeRegMultiEx(Array.ConvertAll(ArabicData.RecitationCombiningSymbols, Function(C As Char) ArabicData.MakeUniRegEx(C))) + ")*(?!(" + ArabicData.MakeRegMultiEx(Array.ConvertAll(ArabicData.RecitationConnectingFollowerSymbols, Function(C As Char) ArabicData.MakeUniRegEx(C))) + "))", .Evaluator = ChrW(&H6A1)},
        New RuleTranslation With {.Rule = "FehIsolatedEnd", .Match = ArabicData.MakeUniRegEx(ArabicData.ArabicLetterFeh) + "(" + ArabicData.MakeRegMultiEx(Array.ConvertAll(ArabicData.RecitationCombiningSymbols, Function(C As Char) ArabicData.MakeUniRegEx(C))) + ")*(" + ArabicData.MakeRegMultiEx(Array.ConvertAll(ArabicData.RecitationConnectingFollowerSymbols, Function(C As Char) ArabicData.MakeUniRegEx(C))) + ")", .Evaluator = ChrW(&H6A2)},
        New RuleTranslation With {.Rule = "QafBeginMiddle", .Match = ArabicData.MakeUniRegEx(ArabicData.ArabicLetterQaf) + "(" + ArabicData.MakeRegMultiEx(Array.ConvertAll(ArabicData.RecitationCombiningSymbols, Function(C As Char) ArabicData.MakeUniRegEx(C))) + ")*(?!(" + ArabicData.MakeRegMultiEx(Array.ConvertAll(ArabicData.RecitationConnectingFollowerSymbols, Function(C As Char) ArabicData.MakeUniRegEx(C))) + "))", .Evaluator = ChrW(&H66F)},
        New RuleTranslation With {.Rule = "QafIsolatedEnd", .Match = ArabicData.MakeUniRegEx(ArabicData.ArabicLetterQaf) + "(" + ArabicData.MakeRegMultiEx(Array.ConvertAll(ArabicData.RecitationCombiningSymbols, Function(C As Char) ArabicData.MakeUniRegEx(C))) + ")*(" + ArabicData.MakeRegMultiEx(Array.ConvertAll(ArabicData.RecitationConnectingFollowerSymbols, Function(C As Char) ArabicData.MakeUniRegEx(C))) + ")", .Evaluator = ChrW(&H6A7)},
        New RuleTranslation With {.Rule = "Kaf", .Match = ArabicData.MakeUniRegEx(ArabicData.ArabicLetterKaf), .Evaluator = ChrW(&H6A9)},
        New RuleTranslation With {.Rule = "NoonBeginMiddle", .Match = ArabicData.MakeUniRegEx(ArabicData.ArabicLetterNoon) + "(" + ArabicData.MakeRegMultiEx(Array.ConvertAll(ArabicData.RecitationCombiningSymbols, Function(C As Char) ArabicData.MakeUniRegEx(C))) + ")*(?!(" + ArabicData.MakeRegMultiEx(Array.ConvertAll(ArabicData.RecitationConnectingFollowerSymbols, Function(C As Char) ArabicData.MakeUniRegEx(C))) + "))", .Evaluator = ChrW(&H6BA)},
        New RuleTranslation With {.Rule = "NoonIsolatedEnd", .Match = ArabicData.MakeUniRegEx(ArabicData.ArabicLetterNoon) + "(" + ArabicData.MakeRegMultiEx(Array.ConvertAll(ArabicData.RecitationCombiningSymbols, Function(C As Char) ArabicData.MakeUniRegEx(C))) + ")*(" + ArabicData.MakeRegMultiEx(Array.ConvertAll(ArabicData.RecitationConnectingFollowerSymbols, Function(C As Char) ArabicData.MakeUniRegEx(C))) + ")", .Evaluator = ChrW(&H646)},
        New RuleTranslation With {.Rule = "ImalaE", .Match = String.Empty, .Evaluator = ChrW(&H65C)},
        New RuleTranslation With {.Rule = "IIFinal", .Match = "(" + ArabicData.MakeUniRegEx(ArabicData.ArabicKasra) + ")(?=" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterYeh) + "(^\s*|\s+))", .Evaluator = "$1" + ChrW(&H6D2)}
    }
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

    Public Shared WarshScript As RuleTranslation() = { _
        New RuleTranslation With {.Rule = "Imaala", .Match = "(" + ArabicData.MakeUniRegEx(ArabicData.ArabicDamma) + ")" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterWawWithHamzaAbove) + "(" + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + ")", _
            .Evaluator = "$1" + ArabicData.ArabicLetterWaw + ArabicData.ArabicRoundedHighStopWithFilledCentre + "$2"}, _
        New RuleTranslation With {.Rule = "SuperscriptAlef", .Match = ArabicData.MakeUniRegEx(ArabicData.ArabicTatweel) + ArabicData.MakeUniRegEx(ArabicData.ArabicHamzaAbove) + ArabicData.MakeUniRegEx(ArabicData.ArabicSukun), _
            .Evaluator = ArabicData.ArabicLetterSuperscriptAlef}, _
        New RuleTranslation With {.Rule = "ShaddaImaala", .Match = "(" + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + ")" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterHamza) + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + "(?=" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterYeh) + ")", _
            .Evaluator = "$1" + ArabicData.ArabicLetterSuperscriptAlef + ArabicData.ArabicMaddahAbove}, _
        New RuleTranslation With {.Rule = "ShaddaImaala", .Match = "(" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterReh) + ArabicData.MakeUniRegEx(ArabicData.ArabicShadda) + "?)" + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + "(" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterHamza) + ")" + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha), _
            .Evaluator = "$1" + ArabicData.ArabicVowelSignDotBelow + "$2" + ArabicData.ArabicVowelSignDotBelow}, _
        New RuleTranslation With {.Rule = "ShaddaImaala", .Match = "(" + ArabicData.MakeUniRegEx(ArabicData.ArabicShadda) + ")" + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + "(?=" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlef) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterReh) + ArabicData.MakeUniRegEx(ArabicData.ArabicKasra) + ")", _
            .Evaluator = "$1" + ArabicData.ArabicVowelSignDotBelow}, _
        New RuleTranslation With {.Rule = "LongVowelAddition", .Match = "((?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicKasra) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicDamma) + ")" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterMeem) + ")(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicSukun) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicDamma) + ")(?=(?:(?:\s+" + ArabicData.MakeUniRegEx(ArabicData.ArabicEndOfAyah) + ")?(?:\s+(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicStartOfRubElHizb) + "|" + ArabicData.MakeRegMultiEx(Array.ConvertAll(ArabicData.ArabicStopLetters, Function(Str As String) ArabicData.MakeUniRegEx(Str))) + "))?)?\s+(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterHamza) + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlefWithHamzaBelow) + "(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicKasra) + ")|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlefWithHamzaAbove) + "(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicDamma) + ")))", _
            .Evaluator = "$1" + ArabicData.ArabicDamma + ArabicData.ArabicLetterWaw + ArabicData.ArabicMaddahAbove}, _
        New RuleTranslation With {.Rule = "Madda", .Match = "(" + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + ")" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterHamza) + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + "(?!" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlef) + ")", _
            .Evaluator = "$1" + ArabicData.ArabicLetterSuperscriptAlef + ArabicData.ArabicMaddahAbove}, _
        New RuleTranslation With {.Rule = "InitialAlefMadda", .Match = "((?:^\s*|" + ArabicData.MakeUniRegEx(ArabicData.ArabicEndOfAyah) + "\s+)|(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicFathatan) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlef) + "?|" + ArabicData.MakeUniRegEx(ArabicData.ArabicKasratan) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicDammatan) + ")\s+)" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterHamza) + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + "(?=" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlef) + ")", _
            .Evaluator = "$1" + ArabicData.ArabicTatweel}, _
        New RuleTranslation With {.Rule = "InitialAlefMadda", .Match = "(" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterLam) + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicKasra) + ")" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterHamza) + "(" + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + ")" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlef), _
            .Evaluator = "$1" + ArabicData.ArabicLetterAlefWithHamzaAbove + "$2"}, _
        New RuleTranslation With {.Rule = "InitialAlefMadda", .Match = "(" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlefWithHamzaAbove) + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + ")" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterHamza) + ArabicData.MakeUniRegEx(ArabicData.ArabicKasra), _
            .Evaluator = "$1" + ArabicData.Bullet}, _
        New RuleTranslation With {.Rule = "InitialAlefMadda", .Match = "(?:(" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterHamza) + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + ")|(" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterHeh) + ")" + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterSuperscriptAlef) + ArabicData.MakeUniRegEx(ArabicData.ArabicMaddahAbove) + ")" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlefWithHamzaAbove) + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + "|(" + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + ")" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlefWithHamzaAbove) + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + "(?=" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterHamza) + ")", _
            .Evaluator = "$1$2$2" + ArabicData.ArabicLetterAlefWithMaddaAbove}, _
        New RuleTranslation With {.Rule = "InitialAlefMadda", .Match = ArabicData.MakeUniRegEx(ArabicData.ArabicSukun) + "(\s+)" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterHamza) + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + "(?=" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlef) + ")", _
            .Evaluator = ArabicData.ArabicFatha + "$1" + ArabicData.ArabicTatweel}, _
        New RuleTranslation With {.Rule = "InitialAlefMadda", .Match = ArabicData.MakeUniRegEx(ArabicData.ArabicSukun) + "(\s+)" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterHamza) + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlefWithHamzaAbove) + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha), _
            .Evaluator = ArabicData.ArabicFatha + "$1" + ArabicData.ArabicTatweel + ArabicData.ArabicLetterAlefWithMaddaAbove}, _
        New RuleTranslation With {.Rule = "InitialAlefMadda", .Match = "(" + ArabicData.MakeUniRegEx(ArabicData.ArabicSukun) + ")?(\s+)" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterHamza) + "(?=" + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlef) + ")", _
            .Evaluator = "$2" + ArabicData.ArabicLetterAlefMaksura + ArabicData.ArabicHamzaBelow, .NegativeMatch = "$1"}, _
        New RuleTranslation With {.Rule = "InitialAlefMadda", .Match = "(" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterLam) + ")" + ArabicData.MakeUniRegEx(ArabicData.ArabicSukun) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterHamza) + "(?=" + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlef) + ")", _
            .Evaluator = "$1"}, _
        New RuleTranslation With {.Rule = "InitialAlefMadda", .Match = "(" + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + ")" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterSuperscriptAlef) + ArabicData.MakeUniRegEx(ArabicData.ArabicMaddahAbove) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterHamza), _
            .Evaluator = "$1" + ArabicData.ArabicLetterAlefWithMaddaAbove + ArabicData.ArabicLetterAlefMaksura + ArabicData.ArabicHamzaBelow}, _
        New RuleTranslation With {.Rule = "DroppedHamza", .Match = ArabicData.MakeUniRegEx(ArabicData.ArabicSukun) + "((?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlef) + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallHighRoundedZero) + ")?(?:(?:\s+" + ArabicData.MakeUniRegEx(ArabicData.ArabicEndOfAyah) + ")?(?:\s+(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicStartOfRubElHizb) + "|" + ArabicData.MakeRegMultiEx(Array.ConvertAll(ArabicData.ArabicStopLetters, Function(Str As String) ArabicData.MakeUniRegEx(Str))) + "))?)?\s+)" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlefWithHamzaAbove) + "(?=" + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + ")", _
            .Evaluator = ArabicData.ArabicFatha + "$1" + ArabicData.ArabicLetterAlef}, _
        New RuleTranslation With {.Rule = "DroppedHamza", .Match = ArabicData.MakeUniRegEx(ArabicData.ArabicSukun) + "((?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlef) + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallHighRoundedZero) + ")?(?:(?:\s+" + ArabicData.MakeUniRegEx(ArabicData.ArabicEndOfAyah) + ")?(?:\s+(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicStartOfRubElHizb) + "|" + ArabicData.MakeRegMultiEx(Array.ConvertAll(ArabicData.ArabicStopLetters, Function(Str As String) ArabicData.MakeUniRegEx(Str))) + "))?)?\s+)" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlefWithHamzaAbove) + "(?=" + ArabicData.MakeUniRegEx(ArabicData.ArabicDamma) + ")", _
            .Evaluator = ArabicData.ArabicDamma + "$1" + ArabicData.ArabicLetterAlef}, _
        New RuleTranslation With {.Rule = "DroppedHamza", .Match = "(" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterHamza) + "(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicKasra) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicDamma) + ")(?:(?:\s+" + ArabicData.MakeUniRegEx(ArabicData.ArabicEndOfAyah) + ")?(?:\s+(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicStartOfRubElHizb) + "|" + ArabicData.MakeRegMultiEx(Array.ConvertAll(ArabicData.ArabicStopLetters, Function(Str As String) ArabicData.MakeUniRegEx(Str))) + "))?)?\s+)" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlefWithHamzaAbove), _
            .Evaluator = "$1" + ArabicData.ArabicLetterAlef + ArabicData.ArabicRoundedHighStopWithFilledCentre}, _
        New RuleTranslation With {.Rule = "DroppedHamza", .Match = "((?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicFathatan) + "(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlef) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlefMaksura) + ")?|" + ArabicData.MakeUniRegEx(ArabicData.ArabicKasratan) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicDammatan) + ")(?:(?:\s+" + ArabicData.MakeUniRegEx(ArabicData.ArabicEndOfAyah) + ")?(?:\s+(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicStartOfRubElHizb) + "|" + ArabicData.MakeRegMultiEx(Array.ConvertAll(ArabicData.ArabicStopLetters, Function(Str As String) ArabicData.MakeUniRegEx(Str))) + "))?)?\s+)(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlefWithHamzaAbove) + "(?=" + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicDamma) + ")|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlefWithHamzaBelow) + "(?=" + ArabicData.MakeUniRegEx(ArabicData.ArabicKasra) + "))", _
            .Evaluator = "$1" + ArabicData.ArabicLetterAlef}, _
        New RuleTranslation With {.Rule = "DroppedHamza", .Match = "((?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterHamza) + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterHamza) + ArabicData.MakeUniRegEx(ArabicData.ArabicKasra) + ")(?:(?:\s+" + ArabicData.MakeUniRegEx(ArabicData.ArabicEndOfAyah) + ")?(?:\s+(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicStartOfRubElHizb) + "|" + ArabicData.MakeRegMultiEx(Array.ConvertAll(ArabicData.ArabicStopLetters, Function(Str As String) ArabicData.MakeUniRegEx(Str))) + "))?)?\s+)(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlefWithHamzaAbove) + "(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicDamma) + ")|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlefWithHamzaBelow) + ArabicData.MakeUniRegEx(ArabicData.ArabicKasra) + ")", _
            .Evaluator = "$1" + ArabicData.ArabicLetterAlef}, _
        New RuleTranslation With {.Rule = "DroppedHamza", .Match = ArabicData.MakeUniRegEx(ArabicData.ArabicSukun) + "((?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlef) + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallHighRoundedZero) + ")?(?:(?:\s+" + ArabicData.MakeUniRegEx(ArabicData.ArabicEndOfAyah) + ")?(?:\s+(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicStartOfRubElHizb) + "|" + ArabicData.MakeRegMultiEx(Array.ConvertAll(ArabicData.ArabicStopLetters, Function(Str As String) ArabicData.MakeUniRegEx(Str))) + "))?)?\s+)" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlefWithHamzaBelow) + "(?=" + ArabicData.MakeUniRegEx(ArabicData.ArabicKasra) + ")", _
            .Evaluator = ArabicData.ArabicKasra + "$1" + ArabicData.ArabicLetterAlef}, _
        New RuleTranslation With {.Rule = "DroppedHamza", .Match = ArabicData.MakeUniRegEx(ArabicData.ArabicSukun) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlefWithHamzaAbove) + "(" + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + ")", _
            .Evaluator = "$1" + ArabicData.ArabicLetterAlef}, _
        New RuleTranslation With {.Rule = "DroppedHamza", .Match = ArabicData.MakeUniRegEx(ArabicData.ArabicSukun) + "(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlefWithHamzaAbove) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlefWithHamzaBelow) + ")", _
            .Evaluator = ArabicData.ArabicLetterAlef}, _
        New RuleTranslation With {.Rule = "DroppedHamza", .Match = "((?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterTeh) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterYeh) + ")" + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + ")" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlefWithHamzaAbove) + ArabicData.MakeUniRegEx(ArabicData.ArabicSukun) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlefWithHamzaAbove) + ArabicData.MakeUniRegEx(ArabicData.ArabicSukun) + "(?=" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterTeh) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterKaf) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterMeem) + ")", _
            .Evaluator = "$1" + ArabicData.ArabicLetterAlef}, _
        New RuleTranslation With {.Rule = "DroppedHamza", .Match = ArabicData.MakeUniRegEx(ArabicData.ArabicLetterWawWithHamzaAbove) + ArabicData.MakeUniRegEx(ArabicData.ArabicSukun), _
            .Evaluator = ArabicData.ArabicLetterWaw}, _
        New RuleTranslation With {.Rule = "DroppedHamza", .Match = "(" + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterBeh) + ArabicData.MakeUniRegEx(ArabicData.ArabicKasra) + ")" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterYehWithHamzaAbove) + ArabicData.MakeUniRegEx(ArabicData.ArabicSukun), _
            .Evaluator = "$1" + ArabicData.ArabicLetterYeh}, _
        New RuleTranslation With {.Rule = "DroppedHamza", .Match = "(" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterLam) + ")" + ArabicData.MakeUniRegEx(ArabicData.ArabicSukun) + ArabicData.MakeUniRegEx(ArabicData.ArabicTatweel) + ArabicData.MakeUniRegEx(ArabicData.ArabicHamzaAbove), _
            .Evaluator = "$1"}, _
        New RuleTranslation With {.Rule = "AddedHamza", .Match = "(" + ArabicData.MakeUniRegEx(ArabicData.ArabicDamma) + ")" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterWaw) + "(?=" + ArabicData.MakeUniRegEx(ArabicData.ArabicFathatan) + ")", _
            .Evaluator = "$1" + ArabicData.ArabicLetterWawWithHamzaAbove}, _
        New RuleTranslation With {.Rule = "AddedHamza", .Match = "((?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterJeem) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterSheen) + ")" + ArabicData.MakeUniRegEx(ArabicData.ArabicKasra) + ")" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterYeh) + "(?=" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterTeh) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterNoon) + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlef) + ")", _
            .Evaluator = "$1" + ArabicData.ArabicLetterYehWithHamzaAbove + ArabicData.ArabicSukun}, _
        New RuleTranslation With {.Rule = "HamzaBelow", .Match = ArabicData.MakeUniRegEx(ArabicData.ArabicLetterHamza) + "(?!" + ArabicData.MakeUniRegEx(ArabicData.ArabicVowelSignDotBelow) + ")|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterYehWithHamzaAbove) + "(?=" + ArabicData.MakeUniRegEx(ArabicData.ArabicKasra) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicKasratan) + ")", _
            .Evaluator = ArabicData.ArabicLetterAlefMaksura + ArabicData.ArabicHamzaBelow}, _
        New RuleTranslation With {.Rule = "HamzaAbove", .Match = ArabicData.MakeUniRegEx(ArabicData.ArabicLetterYehWithHamzaAbove) + "(?=" + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + "(?=" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterNoon) + ArabicData.MakeUniRegEx(ArabicData.ArabicShadda) + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + "(?!" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterHeh) + ")))|(" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterSheen) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicShadda) + ")" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterYehWithHamzaAbove) + "(?=" + ArabicData.MakeUniRegEx(ArabicData.ArabicDamma) + ")", _
            .Evaluator = "$1" + ArabicData.ArabicLetterAlefMaksura + ArabicData.ArabicHamzaAbove}, _
        New RuleTranslation With {.Rule = "ShaddaImaala", .Match = "((?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicShadda) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterNoon) + ")" + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterBeh) + ArabicData.MakeUniRegEx(ArabicData.ArabicKasra) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlefMaksura) + ")" + ArabicData.MakeUniRegEx(ArabicData.ArabicShadda) + "(?=" + ArabicData.MakeUniRegEx(ArabicData.ArabicDamma) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicDammatan) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicKasra) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicKasratan) + ")", _
            .Evaluator = "$1" + ArabicData.ArabicMaddahAbove + ArabicData.ArabicLetterAlefMaksura + ArabicData.ArabicHamzaBelow}, _
        New RuleTranslation With {.Rule = "HelperMeems", .Match = "(?:(" + ArabicData.MakeUniRegEx(ArabicData.ArabicFathatan) + ")(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallHighMeemIsolatedForm) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallLowMeem) + ")(" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlef) + ")|(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallHighMeemIsolatedForm) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallLowMeem) + ")(?!" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlef) + "))(?!(?:(?:(?:\s+" + ArabicData.MakeUniRegEx(ArabicData.ArabicEndOfAyah) + ")?(?:\s+(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicStartOfRubElHizb) + "|" + ArabicData.MakeRegMultiEx(Array.ConvertAll(ArabicData.ArabicStopLetters, Function(Str As String) ArabicData.MakeUniRegEx(Str))) + "))?)?\s+)?" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterBeh) + ")", _
            .Evaluator = "$1$2"}, _
        New RuleTranslation With {.Rule = "Fatha", .Match = "((?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicKasra) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterNoon) + ArabicData.MakeUniRegEx(ArabicData.ArabicShadda) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicHamzaBelow) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterHah) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterLam) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterReh) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicShadda) + "|(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAin) + ArabicData.MakeUniRegEx(ArabicData.ArabicSukun) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlef) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterNoon) + ")" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterDal) + ")" + ArabicData.MakeUniRegEx(ArabicData.ArabicKasra) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlefMaksura) + ")" + ArabicData.MakeUniRegEx(ArabicData.ArabicMaddahAbove), _
            .Evaluator = "$1" + ArabicData.ArabicFatha}, _
        New RuleTranslation With {.Rule = "Madda", .Match = ArabicData.MakeUniRegEx(ArabicData.ArabicSmallHighUprightRectangularZero) + "(?=\s+" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlefWithHamzaAbove) + ")|(" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlefMaksura) + ")" + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + "(?=" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlefMaksura) + ")", _
            .Evaluator = "$1" + ArabicData.ArabicMaddahAbove}, _
        New RuleTranslation With {.Rule = "SmallYeh", .Match = ArabicData.MakeUniRegEx(ArabicData.ArabicSmallYeh) + "(?=" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterMeem) + ")", _
            .Evaluator = ArabicData.ArabicLetterYeh}, _
        New RuleTranslation With {.Rule = "StopSymbols", .Match = "(^\s*|" + ArabicData.MakeUniRegEx(ArabicData.ArabicEndOfAyah) + "\s+)(?:" + ArabicData.MakeRegMultiEx(Array.ConvertAll(ArabicData.ArabicStopLetters, Function(Str As String) ArabicData.MakeUniRegEx(Str))) + ")\s+|\s+" + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallHighSeen), _
            .Evaluator = "$1"}, _
        New RuleTranslation With {.Rule = "Imaala", .Match = "((?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAin) + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicKasra) + ")" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterLam) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterHah) + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterTeh) + ArabicData.MakeUniRegEx(ArabicData.ArabicShadda) + ")?" + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + "(?=" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlefMaksura) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterSuperscriptAlef) + ")|(" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterNoon) + ArabicData.MakeUniRegEx(ArabicData.ArabicSukun) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterYeh) + ")" + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + "|(" + ArabicData.MakeUniRegEx(ArabicData.ArabicSukun) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterKaf) + ")" + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + "(?=" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterSuperscriptAlef) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterFeh) + ")|" + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + "(?=" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlef) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterReh) + "(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicKasra) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicKasratan) + ")(?:\s+|\s*$))",
           .Evaluator = "$2$3" + ArabicData.ArabicVowelSignDotBelow, .NegativeMatch = "$1"}, _
        New RuleTranslation With {.Rule = "Tatweel", .Match = "(" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlefMaksura) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterWaw) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlef) + "|(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterReh) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterZain) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterDal) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterThal) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterWaw) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlef) + ")" + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + ")?(" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterSuperscriptAlef) + ")", _
            .Evaluator = ArabicData.ArabicTatweel + "$2", .NegativeMatch = "$1"}
        }

    Public Shared UthmaniMinimalScript As RuleTranslation() = { _
        New RuleTranslation With {.Rule = "SmallYehSmallWawAfterPronounHeh", .Match = "(?:(" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterHeh) + ArabicData.MakeUniRegEx(ArabicData.ArabicKasra) + ")" + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallYeh) + "|(" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterHeh) + ArabicData.MakeUniRegEx(ArabicData.ArabicDamma) + ")" + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallWaw) + ")" + ArabicData.MakeUniRegEx(ArabicData.ArabicMaddahAbove) + "?(?=\s*$|\s+)", _
            .Evaluator = "$1$2"}, _
        New RuleTranslation With {.Rule = "AlefMaksuraToYeh", .Match = "(" + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + ")" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlefMaksura) + "(?=" + ArabicData.MakeUniRegEx(ArabicData.ArabicSukun) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterHamza) + ")|(" + ArabicData.MakeUniRegEx(ArabicData.ArabicKasra) + ")" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlefMaksura) + "(?=" + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallHighRoundedZero) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicShadda) + ArabicData.MakeUniRegEx(ArabicData.ArabicKasra) + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallYeh) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicMaddahAbove) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterHamza) + ")|(" + ArabicData.MakeUniRegEx(ArabicData.ArabicSukun) + ")" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlefMaksura) + "(?=" + ArabicData.MakeUniRegEx(ArabicData.ArabicKasra) + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallYeh) + ")|(" + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallHighRoundedZero) + ")" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlefMaksura) + "(?=" + ArabicData.MakeUniRegEx(ArabicData.ArabicMaddahAbove) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicSukun) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterHamza) + ")", _
            .Evaluator = "$1$2$3$4" + CStr(ArabicData.ArabicLetterYeh)}, _
        New RuleTranslation With {.Rule = "ShortVowelsBeforeLongVowels", .Match = "(?:(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + "(?=" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterWaw) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterSuperscriptAlef) + "(?!(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicShadda) + "|" + ArabicData.MakeRegMultiEx(Array.ConvertAll(ArabicData.ArabicFathaDammaKasra, Function(Str As String) ArabicData.MakeUniRegEx(Str))) + "|" + ArabicData.MakeRegMultiEx(Array.ConvertAll(ArabicData.ArabicTanweens, Function(Str As String) ArabicData.MakeUniRegEx(Str))) + "|(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlef) + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallHighRoundedZero) + ")?(?: (?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallHighLamAlef) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallHighJeem) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallHighLigatureSadWithLamWithAlefMaksura) + "))? " + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlefWasla) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallHighRoundedZero) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicRoundedHighStopWithFilledCentre) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicSukun) + ")))|" + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + "(?=" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterSuperscriptAlef) + "(?!$)(?!(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicShadda) + "|" + ArabicData.MakeRegMultiEx(Array.ConvertAll(ArabicData.ArabicFathaDammaKasra, Function(Str As String) ArabicData.MakeUniRegEx(Str))) + "|" + ArabicData.MakeRegMultiEx(Array.ConvertAll(ArabicData.ArabicTanweens, Function(Str As String) ArabicData.MakeUniRegEx(Str))) + "|(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlef) + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallHighRoundedZero) + ")?(?: (?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallHighLamAlef) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallHighJeem) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallHighLigatureSadWithLamWithAlefMaksura) + "))? " + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlefWasla) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallHighRoundedZero) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicRoundedHighStopWithFilledCentre) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicSukun) + ")))|" + ArabicData.MakeUniRegEx(ArabicData.ArabicDamma) + "(?=" + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallWaw) + "(?!(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicShadda) + "|" + ArabicData.MakeRegMultiEx(Array.ConvertAll(ArabicData.ArabicFathaDammaKasra, Function(Str As String) ArabicData.MakeUniRegEx(Str))) + "|" + ArabicData.MakeRegMultiEx(Array.ConvertAll(ArabicData.ArabicTanweens, Function(Str As String) ArabicData.MakeUniRegEx(Str))) + "|(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlef) + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallHighRoundedZero) + ")?(?: (?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallHighLamAlef) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallHighJeem) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallHighLigatureSadWithLamWithAlefMaksura) + "))? " + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlefWasla) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallHighRoundedZero) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicRoundedHighStopWithFilledCentre) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicSukun) + ")))|" + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + "(?=" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterSuperscriptAlef) + "(?!$)(?!(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicShadda) + "|" + ArabicData.MakeRegMultiEx(Array.ConvertAll(ArabicData.ArabicFathaDammaKasra, Function(Str As String) ArabicData.MakeUniRegEx(Str))) + "|" + ArabicData.MakeRegMultiEx(Array.ConvertAll(ArabicData.ArabicTanweens, Function(Str As String) ArabicData.MakeUniRegEx(Str))) + "|(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlef) + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallHighRoundedZero) + ")?(?: (?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallHighLamAlef) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallHighJeem) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallHighLigatureSadWithLamWithAlefMaksura) + "))? " + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlefWasla) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallHighRoundedZero) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicRoundedHighStopWithFilledCentre) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicSukun) + ")))|" + ArabicData.MakeUniRegEx(ArabicData.ArabicKasra) + "(?=" + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallYeh) + "(?!(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicShadda) + "|" + ArabicData.MakeRegMultiEx(Array.ConvertAll(ArabicData.ArabicFathaDammaKasra, Function(Str As String) ArabicData.MakeUniRegEx(Str))) + "|" + ArabicData.MakeRegMultiEx(Array.ConvertAll(ArabicData.ArabicTanweens, Function(Str As String) ArabicData.MakeUniRegEx(Str))) + "|(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlef) + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallHighRoundedZero) + ")?(?: (?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallHighLamAlef) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallHighJeem) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallHighLigatureSadWithLamWithAlefMaksura) + "))? " + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlefWasla) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallHighRoundedZero) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicRoundedHighStopWithFilledCentre) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicSukun) + ")))|" + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + "(?=" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterSuperscriptAlef) + "(?!$)(?!(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicShadda) + "|" + ArabicData.MakeRegMultiEx(Array.ConvertAll(ArabicData.ArabicFathaDammaKasra, Function(Str As String) ArabicData.MakeUniRegEx(Str))) + "|" + ArabicData.MakeRegMultiEx(Array.ConvertAll(ArabicData.ArabicTanweens, Function(Str As String) ArabicData.MakeUniRegEx(Str))) + "|(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlef) + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallHighRoundedZero) + ")?(?: (?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallHighLamAlef) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallHighJeem) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallHighLigatureSadWithLamWithAlefMaksura) + "))? " + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlefWasla) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallHighRoundedZero) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicRoundedHighStopWithFilledCentre) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicSukun) + ")))|" + ArabicData.MakeRegMultiEx(Array.ConvertAll(ArabicData.ArabicLongVowels, Function(StrV As String) ArabicData.MakeUniRegEx(StrV(0)) + "(?=" + ArabicData.MakeUniRegEx(StrV(1)) + "(?!(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicShadda) + "|" + ArabicData.MakeRegMultiEx(Array.ConvertAll(ArabicData.ArabicFathaDammaKasra, Function(Str As String) ArabicData.MakeUniRegEx(Str))) + "|" + ArabicData.MakeRegMultiEx(Array.ConvertAll(ArabicData.ArabicTanweens, Function(Str As String) ArabicData.MakeUniRegEx(Str))) + "|(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlef) + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallHighRoundedZero) + ")?(?: (?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallHighLamAlef) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallHighJeem) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallHighLigatureSadWithLamWithAlefMaksura) + "))? " + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlefWasla) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallHighRoundedZero) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicRoundedHighStopWithFilledCentre) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicSukun) + ")))|" + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + "(?=" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterSuperscriptAlef) + "(?!$)(?!(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicShadda) + "|" + ArabicData.MakeRegMultiEx(Array.ConvertAll(ArabicData.ArabicFathaDammaKasra, Function(Str As String) ArabicData.MakeUniRegEx(Str))) + "|" + ArabicData.MakeRegMultiEx(Array.ConvertAll(ArabicData.ArabicTanweens, Function(Str As String) ArabicData.MakeUniRegEx(Str))) + "|(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlef) + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallHighRoundedZero) + ")?(?: (?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallHighLamAlef) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallHighJeem) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallHighLigatureSadWithLamWithAlefMaksura) + "))? " + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlefWasla) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallHighRoundedZero) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicRoundedHighStopWithFilledCentre) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicSukun) + ")))")) + ")|(?:(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + "(?=" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlefMaksura) + "(?!$)(?!(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicShadda) + "|" + ArabicData.MakeRegMultiEx(Array.ConvertAll(ArabicData.ArabicFathaDammaKasra, Function(Str As String) ArabicData.MakeUniRegEx(Str))) + "|" + ArabicData.MakeRegMultiEx(Array.ConvertAll(ArabicData.ArabicTanweens, Function(Str As String) ArabicData.MakeUniRegEx(Str))) + "|(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlef) + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallHighRoundedZero) + ")?(?: (?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallHighLamAlef) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallHighJeem) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallHighLigatureSadWithLamWithAlefMaksura) + "))? " + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlefWasla) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallHighRoundedZero) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicRoundedHighStopWithFilledCentre) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicSukun) + ")))|" + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + "(?=" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterSuperscriptAlef) + "(?!$)(?!(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicShadda) + "|" + ArabicData.MakeRegMultiEx(Array.ConvertAll(ArabicData.ArabicFathaDammaKasra, Function(Str As String) ArabicData.MakeUniRegEx(Str))) + "|" + ArabicData.MakeRegMultiEx(Array.ConvertAll(ArabicData.ArabicTanweens, Function(Str As String) ArabicData.MakeUniRegEx(Str))) + "|(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlef) + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallHighRoundedZero) + ")?(?: (?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallHighLamAlef) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallHighJeem) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallHighLigatureSadWithLamWithAlefMaksura) + "))? " + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlefWasla) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallHighRoundedZero) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicRoundedHighStopWithFilledCentre) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicSukun) + ")))))|" + ArabicData.MakeUniRegEx(ArabicData.ArabicKasra) + "(?=" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlefMaksura) + "(?!$)(?!(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicShadda) + "|" + ArabicData.MakeRegMultiEx(Array.ConvertAll(ArabicData.ArabicFathaDammaKasra, Function(Str As String) ArabicData.MakeUniRegEx(Str))) + "|" + ArabicData.MakeRegMultiEx(Array.ConvertAll(ArabicData.ArabicTanweens, Function(Str As String) ArabicData.MakeUniRegEx(Str))) + "|(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlef) + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallHighRoundedZero) + ")?(?: (?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallHighLamAlef) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallHighJeem) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallHighLigatureSadWithLamWithAlefMaksura) + "))? " + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlefWasla) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallHighRoundedZero) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicRoundedHighStopWithFilledCentre) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicSukun) + "))))", _
            .Evaluator = String.Empty}, _
        New RuleTranslation With {.Rule = "ShortVowelsBeforeLongVowels", .Match = "((?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterKhah) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterKaf) + ArabicData.MakeUniRegEx(ArabicData.ArabicSukun) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterReh) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicKasra) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterNoon) + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterFeh) + ArabicData.MakeUniRegEx(ArabicData.ArabicSukun) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterSeen) + ")" + ArabicData.MakeUniRegEx(ArabicData.ArabicKasra) + "(?=" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlefMaksura) + "$))|" + ArabicData.MakeUniRegEx(ArabicData.ArabicKasra) + "(?=" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlefMaksura) + "$)", _
            .Evaluator = "$1"}, _
        New RuleTranslation With {.Rule = "ShortVowelsBeforeLongVowels", .Match = "(" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlefWithHamzaAbove) + ")(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + "(?=(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlef) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterSuperscriptAlef) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlefMaksura) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterSuperscriptAlef) + ")(?!" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterSuperscriptAlef) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallHighRoundedZero) + "))|" + ArabicData.MakeUniRegEx(ArabicData.ArabicDamma) + "(?=(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterWaw) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallWaw) + ")(?!" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterSuperscriptAlef) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallHighRoundedZero) + ")))|(" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlefWithHamzaBelow) + ")" + ArabicData.MakeUniRegEx(ArabicData.ArabicKasra) + "(?=(" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterYeh) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallYeh) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlefMaksura) + ")(?!" + ArabicData.MakeUniRegEx(ArabicData.ArabicShadda) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlef) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallHighRoundedZero) + "))", _
            .Evaluator = "$1$2"}, _
        New RuleTranslation With {.Rule = "SmallSadAboveHamzaWasl", .Match = ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlefWasla), _
            .Evaluator = CStr(ArabicData.ArabicLetterAlef)}, _
        New RuleTranslation With {.Rule = "Sukun", .Match = ArabicData.MakeUniRegEx(ArabicData.ArabicSukun), _
            .Evaluator = String.Empty}, _
        New RuleTranslation With {.Rule = "Madda", .Match = ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlefWithMaddaAbove), _
            .Evaluator = CStr(ArabicData.ArabicLetterAlef)}, _
        New RuleTranslation With {.Rule = "Madda", .Match = ArabicData.MakeUniRegEx(ArabicData.ArabicMaddahAbove), _
            .Evaluator = String.Empty}, _
        New RuleTranslation With {.Rule = "SmallCircleShowingNonReadLetters", .Match = "(" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterWaw) + "(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicDamma) + "?" + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallWaw) + "?)" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlef) + ")" + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallHighRoundedZero) + "(?=\s*$|\s+)|(" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlefWithHamzaAbove) + ArabicData.MakeUniRegEx(ArabicData.ArabicDamma) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterWaw) + ")" + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallHighRoundedZero) + "(?=" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterLam) + ")", _
            .Evaluator = "$1$2"}, _
        New RuleTranslation With {.Rule = "ShaddaIdgham", .Match = "((?:\s+|^\s*)(?:" + ArabicData.MakeRegMultiEx(Array.ConvertAll(ArabicData.ArabicSunLetters, Function(Str As String) ArabicData.MakeUniRegEx(Str))) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterWaw) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterMeem) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterYeh) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterBeh) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAin) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterFeh) + "))" + ArabicData.MakeUniRegEx(ArabicData.ArabicShadda), _
            .Evaluator = "$1"}, _
        New RuleTranslation With {.Rule = "ShaddaIdgham", .Match = "(" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterYeh) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterYeh) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterDal) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterTeh) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterHeh) + "|(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicKasra) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterKaf) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterQaf) + ")" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterKaf) + ")" + ArabicData.MakeUniRegEx(ArabicData.ArabicShadda) + "(?=(" + ArabicData.MakeUniRegEx(ArabicData.ArabicDamma) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicKasra) + ")(" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterHeh) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterKaf) + ")?((" + ArabicData.MakeUniRegEx(ArabicData.ArabicDamma) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + ")?(" + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallWaw) + "|(" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterMeem) + "(" + ArabicData.MakeUniRegEx(ArabicData.ArabicDamma) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterWaw) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterHeh) + ArabicData.MakeUniRegEx(ArabicData.ArabicDamma) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterMeem) + ")?)|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterNoon) + "(" + ArabicData.MakeUniRegEx(ArabicData.ArabicShadda) + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlef) + ")|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlef) + ")?(\s*$|\s+)))", _
            .Evaluator = "$1"}, _
        New RuleTranslation With {.Rule = "SmallMeemIghlab", .Match = ArabicData.MakeUniRegEx(ArabicData.ArabicSmallHighMeemIsolatedForm) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallLowMeem), _
            .Evaluator = String.Empty}, _
        New RuleTranslation With {.Rule = "SuperscriptAlefHamzaAbove", .Match = ArabicData.MakeUniRegEx(ArabicData.ArabicLetterSuperscriptAlef) + ArabicData.MakeUniRegEx(ArabicData.ArabicHamzaAbove), _
            .Evaluator = CStr(ArabicData.ArabicFatha) + CStr(ArabicData.ArabicLetterHamza)},
        New RuleTranslation With {.Rule = "AlefRoundedHighStopWithFilledCentre", .Match = ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlef) + ArabicData.MakeUniRegEx(ArabicData.ArabicRoundedHighStopWithFilledCentre), _
            .Evaluator = CStr(ArabicData.ArabicLetterAlefWithHamzaAbove) + CStr(ArabicData.ArabicFatha)}
    }
    Public Shared SimpleEnhancedScript As RuleTranslation() = { _
        New RuleTranslation With {.Rule = "AlefRoundedHighStopWithFilledCentre", .Match = ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlef) + ArabicData.MakeUniRegEx(ArabicData.ArabicRoundedHighStopWithFilledCentre), _
            .Evaluator = CStr(ArabicData.ArabicLetterAlefWithHamzaAbove) + CStr(ArabicData.ArabicFatha)}, _
        New RuleTranslation With {.Rule = "EmptyCenterLowStop", .Match = ArabicData.MakeUniRegEx(ArabicData.ArabicEmptyCentreLowStop) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlefMaksura) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterSuperscriptAlef), _
            .Evaluator = CStr(ArabicData.ArabicFatha) + CStr(ArabicData.ArabicLetterAlef)}, _
        New RuleTranslation With {.Rule = "SmallYehSmallWawAfterPronounHeh", .Match = "(?:(" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterHeh) + ArabicData.MakeUniRegEx(ArabicData.ArabicKasra) + ")" + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallYeh) + "|(" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterHeh) + ArabicData.MakeUniRegEx(ArabicData.ArabicDamma) + ")" + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallWaw) + ")" + ArabicData.MakeUniRegEx(ArabicData.ArabicMaddahAbove) + "?(?=\s*$|\s+)", _
            .Evaluator = "$1$2"}, _
        New RuleTranslation With {.Rule = "SmallWaw", .Match = ArabicData.MakeUniRegEx(ArabicData.ArabicSmallWaw), _
            .Evaluator = CStr(ArabicData.ArabicLetterWaw)}, _
        New RuleTranslation With {.Rule = "SmallYeh", .Match = ArabicData.MakeUniRegEx(ArabicData.ArabicSmallYeh), _
            .Evaluator = CStr(ArabicData.ArabicLetterYeh)}, _
        New RuleTranslation With {.Rule = "TrailingAlef", .Match = "((?:(?:(" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlefWithHamzaAbove) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterWawWithHamzaAbove) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterWaw) + "|(?:(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterZain) + "| (?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterWaw) + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + ")?" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterTeh) + ")" + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterTeh) + ArabicData.MakeUniRegEx(ArabicData.ArabicDamma) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlefWasla) + ")" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterDal) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterKhah) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterHeh) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterQaf) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterYeh) + "|(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlefWasla) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicDamma) + ")" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterBeh) + "|(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlefWasla) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlefWithHamzaAbove) + ArabicData.MakeUniRegEx(ArabicData.ArabicDamma) + ")" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterReh) + "|(?:(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterTeh) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicSukun) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterYeh) + ")" + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlefWasla) + ")" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAin) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterGhain) + "|(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterTeh) + ArabicData.MakeUniRegEx(ArabicData.ArabicDamma) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicSukun) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterYeh) + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + ")" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterBeh) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterLam) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterNoon) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterYehWithHamzaAbove) + ")?" + ArabicData.MakeUniRegEx(ArabicData.ArabicSukun) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallHighRoundedZero) + "|((?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterDal) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAin) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterReh) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterTeh) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterHeh) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterThal) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterJeem) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterKhah) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterKaf) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterQaf) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterHamza) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterBeh) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterNoon) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterTah) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterDad) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterMeem) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterLam) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterWaw) + ArabicData.MakeUniRegEx(ArabicData.ArabicShadda) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicTatweel) + ArabicData.MakeUniRegEx(ArabicData.ArabicHamzaAbove) + ")" + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + ")?" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlef) + "|(" + ArabicData.MakeUniRegEx(ArabicData.ArabicTatweel) + "|(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterQaf) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterYeh) + ArabicData.MakeUniRegEx(ArabicData.ArabicDamma) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterLam) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterTah) + ")" + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + ")?" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterSuperscriptAlef) + ")(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAin) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterSeen) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterHah) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicSukun) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterReh) + ArabicData.MakeUniRegEx(ArabicData.ArabicKasra) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterKaf) + "|(" + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + ")?" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterFeh) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterNoon) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterMeem) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterQaf) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterDal) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterHamza) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterReh) + ArabicData.MakeUniRegEx(ArabicData.ArabicShadda) + "|((?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAin) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterKhah) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterBeh) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterMeem) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterTeh) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterKaf) + ArabicData.MakeUniRegEx(ArabicData.ArabicShadda) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterYeh) + ArabicData.MakeUniRegEx(ArabicData.ArabicShadda) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlefWithHamzaAbove) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicTatweel) + ArabicData.MakeUniRegEx(ArabicData.ArabicHamzaAbove) + ")" + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + "|(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterMeem) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterTah) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterTeh) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterKhah) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterSad) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterDal) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterMeem) + ArabicData.MakeUniRegEx(ArabicData.ArabicShadda) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterDal) + ArabicData.MakeUniRegEx(ArabicData.ArabicShadda) + "|(?:(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlefWithHamzaAbove) + "(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicDamma) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + ")" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterReh) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterBeh) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterGhain) + ")" + ArabicData.MakeUniRegEx(ArabicData.ArabicSukun) + ")" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterSeen) + ")" + ArabicData.MakeUniRegEx(ArabicData.ArabicKasra) + "|(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterKaf) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterKhah) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterTeh) + ")" + ArabicData.MakeUniRegEx(ArabicData.ArabicDamma) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicKasra) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterYeh) + ")?" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterLam) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterThal) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterKaf) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicKasra) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterSeen) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterBeh) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterTeh) + "|(" + ArabicData.MakeUniRegEx(ArabicData.ArabicKasra) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicDamma) + ")?" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterJeem) + ")|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterBeh) + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterNoon) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlef) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterSeen) + ArabicData.MakeUniRegEx(ArabicData.ArabicKasra) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterTah) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterSheen) + ArabicData.MakeUniRegEx(ArabicData.ArabicKasra) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterFeh) + "|(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlef) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterReh) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterMeem) + ArabicData.MakeUniRegEx(ArabicData.ArabicDamma) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterHeh) + ArabicData.MakeUniRegEx(ArabicData.ArabicSukun) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterLam) + ")" + ArabicData.MakeUniRegEx(ArabicData.ArabicKasra) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterKaf) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterYehWithHamzaAbove) + ArabicData.MakeUniRegEx(ArabicData.ArabicKasra) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterQaf) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterMeem) + ArabicData.MakeUniRegEx(ArabicData.ArabicDamma) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterReh) + ArabicData.MakeUniRegEx(ArabicData.ArabicSukun) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterSeen) + ArabicData.MakeUniRegEx(ArabicData.ArabicKasra) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterLam) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlef) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterKaf) + ArabicData.MakeUniRegEx(ArabicData.ArabicKasra) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterSeen) + ")(" + ArabicData.MakeUniRegEx(ArabicData.ArabicShadda) + ")?" + ArabicData.MakeUniRegEx(ArabicData.ArabicDamma) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterWaw) + ArabicData.MakeUniRegEx(ArabicData.ArabicMaddahAbove) + "?" + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + "?)" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlef) + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallHighRoundedZero) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlef) + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallHighRoundedZero) + "(?=" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlefMaksura) + ")|((?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterLam) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterDal) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterReh) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterTeh) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterYeh) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlefWithHamzaAbove) + ")" + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterWawWithHamzaAbove) + "(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicDamma) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicDammatan) + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallLowMeem) + "?))" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlef) + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallHighRoundedZero) + "(?=\s*$|\s+|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterThal) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterYeh) + ")", _
            .Evaluator = "$1$9", .NegativeMatch = "$2$3$4$5$6$7$8"}, _
        New RuleTranslation With {.Rule = "TrailingAlef", .Match = "(" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlef) + ArabicData.MakeUniRegEx(ArabicData.ArabicMaddahAbove) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterHamza) + ArabicData.MakeUniRegEx(ArabicData.ArabicDamma) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterWaw) + ArabicData.MakeUniRegEx(ArabicData.ArabicMaddahAbove) + "?)(?=\s*$|\s+)", _
            .Evaluator = "$1" + CStr(ArabicData.ArabicLetterAlef)}, _
        New RuleTranslation With {.Rule = "AlefMaksuraToYeh", .Match = "(" + ArabicData.MakeUniRegEx(ArabicData.ArabicDamma) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterHah) + ArabicData.MakeUniRegEx(ArabicData.ArabicSukun) + ")" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlefMaksura) + "(" + ArabicData.MakeUniRegEx(ArabicData.ArabicKasra) + ")(?=\s*$|\s+)", _
            .Evaluator = "$1" + CStr(ArabicData.ArabicLetterYeh) + "$2" + CStr(ArabicData.ArabicLetterYeh)}, _
        New RuleTranslation With {.Rule = "AlefMaksuraToYeh", .Match = "(" + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallHighRoundedZero) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicSukun) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicKasra) + ")" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlefMaksura) + "|(" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlef) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterSuperscriptAlef) + ")" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlefMaksura) + "(?=" + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + ")|(" + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + ")" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlefMaksura) + "(?=" + ArabicData.MakeUniRegEx(ArabicData.ArabicSukun) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicShadda) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicKasra) + ")", _
            .Evaluator = "$1$2$3" + CStr(ArabicData.ArabicLetterYeh)}, _
        New RuleTranslation With {.Rule = "SmallSadAboveHamzaWasl", .Match = ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlefWasla) + "(?=" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterLam) + ArabicData.MakeUniRegEx(ArabicData.ArabicShadda) + "(?!" + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + "(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterThal) + ArabicData.MakeUniRegEx(ArabicData.ArabicKasra) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterTeh) + ")))", _
            .Evaluator = ArabicData.ArabicLetterAlef + ArabicData.ArabicLetterLam}, _
        New RuleTranslation With {.Rule = "SmallSadAboveHamzaWasl", .Match = ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlefWasla), _
            .Evaluator = CStr(ArabicData.ArabicLetterAlef)}, _
        New RuleTranslation With {.Rule = "Alef", .Match = "((?:^\s*|\s+)" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterYeh) + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + ")(?=" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterBeh) + ArabicData.MakeUniRegEx(ArabicData.ArabicSukun) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterNoon) + ")", _
            .Evaluator = "$1" + CStr(ArabicData.ArabicLetterAlef) + " " + CStr(ArabicData.ArabicLetterAlef)}, _
        New RuleTranslation With {.Rule = "Alef", .Match = "(" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterBeh) + ArabicData.MakeUniRegEx(ArabicData.ArabicSukun) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterNoon) + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + ")" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterWawWithHamzaAbove) + "(?=" + ArabicData.MakeUniRegEx(ArabicData.ArabicDamma) + ")", _
            .Evaluator = "$1 " + CStr(ArabicData.ArabicLetterAlefWithHamzaAbove)}, _
        New RuleTranslation With {.Rule = "Alef", .Match = "(^\s*|\s+)(?=" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterLam) + ArabicData.MakeUniRegEx(ArabicData.ArabicSukun) + ArabicData.MakeUniRegEx(ArabicData.ArabicTatweel) + ArabicData.MakeUniRegEx(ArabicData.ArabicHamzaAbove) + ")|((?:^\s*|\s+)" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterLam) + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + ")(?=" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterTeh) + ArabicData.MakeUniRegEx(ArabicData.ArabicShadda) + ")|((?:^\s*|\s+)(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterWaw) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterFeh) + ")" + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + ")(?=" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterSeen) + ArabicData.MakeUniRegEx(ArabicData.ArabicSukun) + ")|(" + ArabicData.MakeUniRegEx(ArabicData.ArabicSukun) + "(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterReh) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterLam) + ")" + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterTeh) + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicShadda) + ArabicData.MakeUniRegEx(ArabicData.ArabicKasra) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterNoon) + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + ")" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlefMaksura) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterSuperscriptAlef) + ArabicData.MakeUniRegEx(ArabicData.ArabicMaddahAbove) + "?(?=(?:\s*$|\s+(?!" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterLam) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterHamza) + ")))|(" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAin) + "(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterTeh) + ")?" + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterWaw) + ArabicData.MakeUniRegEx(ArabicData.ArabicSukun) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicShadda) + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterHamza) + ArabicData.MakeUniRegEx(ArabicData.ArabicDamma) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterWaw) + ")(?=\s*$|\s+)|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterWaw) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterSuperscriptAlef) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlef) + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallHighRoundedZero) + "(?=\s*$|\s+)", _
            .Evaluator = "$1$2$3$4$5" + CStr(ArabicData.ArabicLetterAlef)}, _
        New RuleTranslation With {.Rule = "HamzaAlefEnding", .Match = "(" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterSuperscriptAlef) + ArabicData.MakeUniRegEx(ArabicData.ArabicMaddahAbove) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterHamza) + ")" + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlef) + "(?= " + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlef) + ")", _
            .Evaluator = "$1" + CStr(ArabicData.ArabicFatha) + CStr(ArabicData.ArabicLetterAlefMaksura)}, _
        New RuleTranslation With {.Rule = "HamzaAlefEnding", .Match = "(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterHamza) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicTatweel) + ArabicData.MakeUniRegEx(ArabicData.ArabicHamzaAbove) + ")" + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlef) + "(?= " + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlef) + ")", _
            .Evaluator = CStr(ArabicData.ArabicLetterAlefWithHamzaAbove) + CStr(ArabicData.ArabicFatha) + CStr(ArabicData.ArabicLetterAlefMaksura)}, _
        New RuleTranslation With {.Rule = "HamzaAlefEnding", .Match = "(?:(" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterReh) + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + ")" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterHamza) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicTatweel) + ArabicData.MakeUniRegEx(ArabicData.ArabicHamzaAbove) + ")" + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlef) + "(?=(?:\s*$|\s+))", _
            .Evaluator = "$1" + CStr(ArabicData.ArabicLetterAlefWithHamzaAbove) + CStr(ArabicData.ArabicFatha) + CStr(ArabicData.ArabicLetterAlefMaksura) + CStr(ArabicData.ArabicLetterSuperscriptAlef)}, _
        New RuleTranslation With {.Rule = "AlefMaksura", .Match = "((?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterTah) + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterGhain) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterSad) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterLam) + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterDal) + ")" + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + ")" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlef) + "(?=(?:\s*$|\s+))", _
            .Evaluator = "$1" + CStr(ArabicData.ArabicLetterAlefMaksura)}, _
        New RuleTranslation With {.Rule = "RehAlefEnding", .Match = "(" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterTeh) + ArabicData.MakeUniRegEx(ArabicData.ArabicSukun) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterReh) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterReh) + ArabicData.MakeUniRegEx(ArabicData.ArabicShadda) + "?" + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterHamza) + ")" + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlef) + ArabicData.MakeUniRegEx(ArabicData.ArabicMaddahAbove) + "?(?=(?:\s*$|\s+))", _
            .Evaluator = "$1" + CStr(ArabicData.ArabicFatha) + CStr(ArabicData.ArabicLetterAlefMaksura) + CStr(ArabicData.ArabicLetterSuperscriptAlef)}, _
        New RuleTranslation With {.Rule = "Madda", .Match = ArabicData.MakeUniRegEx(ArabicData.ArabicLetterWaw) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterSuperscriptAlef) + ArabicData.MakeUniRegEx(ArabicData.ArabicMaddahAbove) + "(?=" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlef) + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallHighRoundedZero) + ")", _
            .Evaluator = String.Empty}, _
        New RuleTranslation With {.Rule = "Madda", .Match = "(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterHamza) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlefWithHamzaAbove) + ")" + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + "(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlef) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterSuperscriptAlef) + ")|(" + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterBeh) + ArabicData.MakeUniRegEx(ArabicData.ArabicKasra) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicSukun) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterSuperscriptAlef) + ArabicData.MakeUniRegEx(ArabicData.ArabicMaddahAbove) + ")" + ArabicData.MakeUniRegEx(ArabicData.ArabicTatweel) + ArabicData.MakeUniRegEx(ArabicData.ArabicHamzaAbove) + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + "(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlef) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicTatweel) + "?" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterSuperscriptAlef) + ")", _
            .Evaluator = "$1" + CStr(ArabicData.ArabicLetterAlefWithMaddaAbove)}, _
        New RuleTranslation With {.Rule = "Madda", .Match = ArabicData.MakeUniRegEx(ArabicData.ArabicLetterSuperscriptAlef) + ArabicData.MakeUniRegEx(ArabicData.ArabicMaddahAbove) + "(?=" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlefWithHamzaAbove) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlefWithMaddaAbove) + ")|((?:^\s*|\s+)(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterWaw) + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + ")?" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterYeh) + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + ")" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterSuperscriptAlef), _
            .Evaluator = "$1" + CStr(ArabicData.ArabicLetterAlef) + " "},
        New RuleTranslation With {.Rule = "Madda", .Match = ArabicData.MakeUniRegEx(ArabicData.ArabicMaddahAbove), _
            .Evaluator = String.Empty},
        New RuleTranslation With {.Rule = "SmallCircleShowingNonReadLetters", .Match = ArabicData.MakeUniRegEx(ArabicData.ArabicSmallHighRoundedZero) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallHighUprightRectangularZero) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicEmptyCentreHighStop), _
            .Evaluator = String.Empty}, _
        New RuleTranslation With {.Rule = "SadSmallHighSeen", .Match = ArabicData.MakeUniRegEx(ArabicData.ArabicLetterSad) + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallHighSeen), _
            .Evaluator = CStr(ArabicData.ArabicLetterSeen)}, _
        New RuleTranslation With {.Rule = "SmallLowSeen", .Match = ArabicData.MakeUniRegEx(ArabicData.ArabicSmallLowSeen), _
            .Evaluator = String.Empty}, _
        New RuleTranslation With {.Rule = "SmallMeemIghlab", .Match = ArabicData.MakeUniRegEx(ArabicData.ArabicSmallHighMeemIsolatedForm) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallLowMeem), _
            .Evaluator = String.Empty}, _
        New RuleTranslation With {.Rule = "SmallHighNoon", .Match = ArabicData.MakeUniRegEx(ArabicData.ArabicSmallHighNoon), _
            .Evaluator = CStr(ArabicData.ArabicLetterNoon)}, _
        New RuleTranslation With {.Rule = "Noon", .Match = "(\S" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlefWithHamzaAbove) + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + ")(?=" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterLam) + ArabicData.MakeUniRegEx(ArabicData.ArabicShadda) + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterWaw) + ")", _
            .Evaluator = "$1" + CStr(ArabicData.ArabicLetterNoon) + " "}, _
        New RuleTranslation With {.Rule = "AlefWithHamzaAbove", .Match = "(" + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + ")" + ArabicData.MakeUniRegEx(ArabicData.ArabicTatweel) + ArabicData.MakeUniRegEx(ArabicData.ArabicHamzaAbove) + ArabicData.MakeUniRegEx(ArabicData.ArabicFathatan) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlef), _
            .Evaluator = "$1" + CStr(ArabicData.ArabicLetterAlefWithHamzaAbove) + CStr(ArabicData.ArabicFathatan)}, _
        New RuleTranslation With {.Rule = "AlefWithHamzaAbove", .Match = "((?:^\s*|\s+)" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlefWithHamzaAbove) + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + ")" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterHamza) + "(?=" + ArabicData.MakeUniRegEx(ArabicData.ArabicDamma) + ")|(^\s*|\s+|" + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + ")" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterHamza) + "(?=" + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + ")|(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlef) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterSuperscriptAlef) + ")" + ArabicData.MakeUniRegEx(ArabicData.ArabicHamzaAbove) + "|(" + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + ")" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterWawWithHamzaAbove) + "(?!" + ArabicData.MakeUniRegEx(ArabicData.ArabicDamma) + "(?!\s*$|\s+))|((?:(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterTeh) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterYeh) + ")" + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterYeh) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterLam) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterJeem) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterTah) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterNoon) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterSeen) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterSheen) + ")" + ArabicData.MakeUniRegEx(ArabicData.ArabicSukun) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + ")" + ArabicData.MakeUniRegEx(ArabicData.ArabicTatweel) + ArabicData.MakeUniRegEx(ArabicData.ArabicHamzaAbove) + "(?=" + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicFathatan) + "(?!" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlef) + ")|" + ArabicData.MakeUniRegEx(ArabicData.ArabicSukun) + ")", _
            .Evaluator = "$1$2$3$4" + CStr(ArabicData.ArabicLetterAlefWithHamzaAbove)}, _
        New RuleTranslation With {.Rule = "WawWithHamzaAbove", .Match = "(" + ArabicData.MakeUniRegEx(ArabicData.ArabicDamma) + ")(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicTatweel) + ArabicData.MakeUniRegEx(ArabicData.ArabicHamzaAbove) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterHamza) + "(?=" + ArabicData.MakeUniRegEx(ArabicData.ArabicSukun) + "))", _
            .Evaluator = "$1" + CStr(ArabicData.ArabicLetterWawWithHamzaAbove)}, _
        New RuleTranslation With {.Rule = "AlefWithHamzaBelow", .Match = "(" + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + ")" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterHamza) + "(?=" + ArabicData.MakeUniRegEx(ArabicData.ArabicKasra) + ")", _
            .Evaluator = "$1" + CStr(ArabicData.ArabicLetterAlefWithHamzaBelow)}, _
        New RuleTranslation With {.Rule = "YehWithHamzaAbove", .Match = "(?:(" + ArabicData.MakeUniRegEx(ArabicData.ArabicSukun) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicKasra) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterYeh) + ")" + ArabicData.MakeUniRegEx(ArabicData.ArabicTatweel) + ArabicData.MakeUniRegEx(ArabicData.ArabicHamzaAbove) + "(?=" + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicFathatan) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicKasra) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicDamma) + ")|(" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterSuperscriptAlef) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlef) + ")" + ArabicData.MakeUniRegEx(ArabicData.ArabicTatweel) + ArabicData.MakeUniRegEx(ArabicData.ArabicHamzaAbove) + "(?!" + ArabicData.MakeUniRegEx(ArabicData.ArabicDamma) + ")|(" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterSuperscriptAlef) + "|(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterZain) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterReh) + ")" + ArabicData.MakeUniRegEx(ArabicData.ArabicKasra) + ")" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterHamza) + ")(?=(" + ArabicData.MakeUniRegEx(ArabicData.ArabicKasra) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicDamma) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + "(?!" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlefMaksura) + ")|" + ArabicData.MakeUniRegEx(ArabicData.ArabicFathatan) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicSukun) + ")(?!\s*$|\s+))", _
            .Evaluator = "$1$2$3" + CStr(ArabicData.ArabicLetterYehWithHamzaAbove)}, _
        New RuleTranslation With {.Rule = "YehWithHamzaAbove", .Match = "(" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlef) + ")" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterHamza) + ArabicData.MakeUniRegEx(ArabicData.ArabicKasra) + "(?=" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterYeh) + ")|(" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterWaw) + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterMeem) + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterLam) + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + ")" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlefWithHamzaBelow) + ArabicData.MakeUniRegEx(ArabicData.ArabicKasra) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterYeh) + "(?!\s*$|\s+|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterMeem) + ")", _
            .Evaluator = "$1$2" + CStr(ArabicData.ArabicLetterYehWithHamzaAbove) + CStr(ArabicData.ArabicKasra)}, _
        New RuleTranslation With {.Rule = "NonReadWawAndYeh", .Match = "(" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterYeh) + ArabicData.MakeUniRegEx(ArabicData.ArabicSukun) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + "(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterFeh) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterBeh) + ")" + (ArabicData.ArabicFatha) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlefWithHamzaBelow) + ArabicData.MakeUniRegEx(ArabicData.ArabicKasra) + ")" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterYeh) + "|(" + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlefWithHamzaAbove) + ArabicData.MakeUniRegEx(ArabicData.ArabicDamma) + ")" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterWaw) + "(?=" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterReh) + ")", _
            .Evaluator = "$1$2"}, _
        New RuleTranslation With {.Rule = "WawAlefMaksuraSuperscriptAlef", .Match = "(" + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + ")(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterWaw) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlefMaksura) + ")" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterSuperscriptAlef) + "(?!\s*$|\s+)", _
            .Evaluator = "$1" + CStr(ArabicData.ArabicLetterAlef)}, _
        New RuleTranslation With {.Rule = "SuperscriptAlef", .Match = "((" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterHah) + ArabicData.MakeUniRegEx(ArabicData.ArabicSukun) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterMeem) + "|(?:^\s*|\s+)(?:(?:(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterWaw) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterFeh) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlefWithHamzaAbove) + ")" + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + ")?(?:(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterKaf) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterWaw) + ")" + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + "|(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterBeh) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterLam) + ")" + ArabicData.MakeUniRegEx(ArabicData.ArabicKasra) + "))?" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterThal) + ArabicData.MakeUniRegEx(ArabicData.ArabicShadda) + "?|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlefWithHamzaBelow) + ArabicData.MakeUniRegEx(ArabicData.ArabicKasra) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterLam) + "|(?:^\s*|\s+)" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterWaw) + ArabicData.MakeUniRegEx(ArabicData.ArabicSukun) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterLam) + ")?" + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + ArabicData.MakeUniRegEx(ArabicData.ArabicTatweel) + "?)" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterSuperscriptAlef) + "(?=" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterGhain) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterHeh) + ArabicData.MakeUniRegEx(ArabicData.ArabicDamma) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterNoon) + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlef) + ")?(?!\s*$|\s+|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterLam) + ArabicData.MakeUniRegEx(ArabicData.ArabicKasra) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterKaf) + "(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicDamma) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterMeem) + ArabicData.MakeUniRegEx(ArabicData.ArabicDamma) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicDamma) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterNoon) + ArabicData.MakeUniRegEx(ArabicData.ArabicShadda) + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + ")|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterWawWithHamzaAbove) + ArabicData.MakeUniRegEx(ArabicData.ArabicDamma) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterLam) + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlef) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterHamza) + ArabicData.MakeUniRegEx(ArabicData.ArabicKasra) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterKaf) + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterThal) + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlef) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterKaf) + ArabicData.MakeUniRegEx(ArabicData.ArabicKasra) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterNoon) + "(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicShadda) + ArabicData.MakeUniRegEx(ArabicData.ArabicKasra) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterYeh) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicShadda) + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + "(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterHeh) + ArabicData.MakeUniRegEx(ArabicData.ArabicDamma) + "(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterMeem) + ArabicData.MakeUniRegEx(ArabicData.ArabicSukun) + ")?|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterKaf) + ArabicData.MakeUniRegEx(ArabicData.ArabicDamma) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterMeem) + ArabicData.MakeUniRegEx(ArabicData.ArabicSukun) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlef) + ")?|" + ArabicData.MakeUniRegEx(ArabicData.ArabicKasra) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicSukun) + ")?(?:\s*$|\s+)|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterThal) + "(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + "(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlef) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterSuperscriptAlef) + ")(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterNoon) + ArabicData.MakeUniRegEx(ArabicData.ArabicKasra) + ")?|" + ArabicData.MakeUniRegEx(ArabicData.ArabicKasra) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterHeh) + ArabicData.MakeUniRegEx(ArabicData.ArabicKasra) + ")|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterHeh) + "(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicKasra) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicDamma) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterKaf) + ArabicData.MakeUniRegEx(ArabicData.ArabicDamma) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterMeem) + ArabicData.MakeUniRegEx(ArabicData.ArabicSukun) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicFathatan) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlef) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + "(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterKaf) + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + ")?)(?:\s+|\s*$)|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterYehWithHamzaAbove) + ArabicData.MakeUniRegEx(ArabicData.ArabicKasra) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterKaf) + "(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicDamma) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterMeem) + ArabicData.MakeUniRegEx(ArabicData.ArabicSukun) + ")(?:\s+|\s*$))|(" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterThal) + ArabicData.MakeUniRegEx(ArabicData.ArabicShadda) + "?" + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + ")?" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterSuperscriptAlef) + "(?=" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterLam) + ArabicData.MakeUniRegEx(ArabicData.ArabicKasra) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterKaf) + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + ")|(" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterWaw) + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterLam) + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + ")" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterSuperscriptAlef) + "(?=" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterYeh) + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + ")|(" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterMeem) + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterSeen) + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + ")" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterSuperscriptAlef), _
            .Evaluator = "$1$4$5" + CStr(ArabicData.ArabicLetterAlef), .NegativeMatch = "$2$3"}, _
        New RuleTranslation With {.Rule = "Hamza", .Match = ArabicData.MakeUniRegEx(ArabicData.ArabicTatweel) + ArabicData.MakeUniRegEx(ArabicData.ArabicHamzaAbove) + "|(" + ArabicData.MakeUniRegEx(ArabicData.ArabicDamma) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterWaw) + ")" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlefWithHamzaAbove) + "(?=(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicDamma) + ")(?:\s*$|\s+))|(" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlef) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlefWithMaddaAbove) + ")(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterWawWithHamzaAbove) + "(?=(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicDamma) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicDammatan) + ")(?:\s*$|\s+))|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterYehWithHamzaAbove) + "(?=" + ArabicData.MakeUniRegEx(ArabicData.ArabicKasra) + "(?:\s*$|\s+)))", _
            .Evaluator = "$1$2" + CStr(ArabicData.ArabicLetterHamza)}, _
        New RuleTranslation With {.Rule = "Fathatan", .Match = "(" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterJeem) + ")" + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + "(?=" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlef) + "(?:\s*$|\s+))", _
            .Evaluator = "$1" + CStr(ArabicData.ArabicFathatan)}, _
        New RuleTranslation With {.Rule = "Fatha", .Match = "((?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterNoon) + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterReh) + ArabicData.MakeUniRegEx(ArabicData.ArabicKasra) + ")" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterJeem) + ")" + ArabicData.MakeUniRegEx(ArabicData.ArabicFathatan) + "(?=" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlef) + "(?:\s*$|\s+))", _
            .Evaluator = "$1" + CStr(ArabicData.ArabicFatha)}
    }
    'New RuleTranslation With {.Rule = "SuperscriptAlef", .Match = "(?:((?:" + ArabicData.MakeRegMultiEx(Array.ConvertAll(SimpleSuperscriptAlefBefore, Function(Str As String) ArabicData.MakeUniRegEx(ArabicData.TransliterateFromBuckwalter(Str.Replace(".", "").Replace("""", "").Replace("@", "").Replace("[", "").Replace("]", "").Replace("-", "").Replace("^", ""))))) + "))|((?:" + ArabicData.MakeRegMultiEx(Array.ConvertAll(SimpleSuperscriptAlefNotBefore, Function(Str As String) ArabicData.MakeUniRegEx(ArabicData.TransliterateFromBuckwalter(Str.Replace(".", "").Replace("""", "").Replace("@", "").Replace("[", "").Replace("]", "").Replace("-", "").Replace("^", ""))))) + ")))" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterSuperscriptAlef) + "(?=(?:" + ArabicData.MakeRegMultiEx(Array.ConvertAll(SimpleSuperscriptAlefAfter, Function(Str As String) ArabicData.MakeUniRegEx(ArabicData.TransliterateFromBuckwalter(Str.Replace(".", "").Replace("""", "").Replace("@", "").Replace("[", "").Replace("]", "").Replace("-", "").Replace("^", ""))))) + "))?(?!" + ArabicData.MakeRegMultiEx(Array.ConvertAll(SimpleSuperscriptAlefNotAfter, Function(Str As String) ArabicData.MakeUniRegEx(ArabicData.TransliterateFromBuckwalter(Str.Replace(".", "").Replace("""", "").Replace("@", "").Replace("[", "").Replace("]", "").Replace("-", "").Replace("^", ""))))) + ")", _
    '    .Evaluator = "$1" + CStr(ArabicData.ArabicLetterAlef), .NegativeMatch = "$2"}, _

    Public Shared SimpleScript As RuleTranslation() = { _
        New RuleTranslation With {.Rule = "RemoveShadda", .Match = "((?:^\s*|\s+)(?:" + ArabicData.MakeRegMultiEx(Array.ConvertAll(ArabicData.ArabicLetters, Function(Str As String) ArabicData.MakeUniRegEx(Str))) + "))" + ArabicData.MakeUniRegEx(ArabicData.ArabicShadda) + "|((?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterQaf) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterKaf) + ")" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterKaf) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterDal) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterTeh) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterYeh) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterYeh) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterHeh) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterHeh) + ")" + ArabicData.MakeUniRegEx(ArabicData.ArabicShadda) + "(?=(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicKasra) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterKaf) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterHeh) + ")?(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicDamma) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + ")(?:(?:(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterNoon) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterHeh) + ArabicData.MakeUniRegEx(ArabicData.ArabicDamma) + "?)?" + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlef) + ")|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterNoon) + ArabicData.MakeUniRegEx(ArabicData.ArabicShadda) + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + "|(?:(?:(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterMeem) + ArabicData.MakeUniRegEx(ArabicData.ArabicDamma) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterWaw) + ")?" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterHeh) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterKaf) + ")(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicDamma) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + "))?(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterMeem) + "(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicSukun) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicDamma) + ")?)?)(?:\s*$|\s+))", _
            .Evaluator = "$1$2"}, _
        New RuleTranslation With {.Rule = "AddSukun", .Match = "((?:^\s*|\s+)(?:" + ArabicData.MakeRegMultiEx(Array.ConvertAll(ArabicData.ArabicUniqueLetters, Function(Str As String) ArabicData.MakeUniRegEx(ArabicData.TransliterateFromBuckwalter(Str).Replace(CStr(ArabicData.ArabicMaddahAbove), String.Empty)))) + ")(?:\s+|\s*$))|(" + ArabicData.MakeRegMultiEx(Array.ConvertAll(ArabicData.ArabicSunLetters, Function(Str As String) ArabicData.MakeUniRegEx(Str))) + "|" + ArabicData.MakeRegMultiEx(Array.ConvertAll(ArabicData.ArabicMoonLettersNoVowels, Function(Str As String) ArabicData.MakeUniRegEx(Str))) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + "(" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterWaw) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterYeh) + "))(?=\s*$|\s+|(" + ArabicData.MakeRegMultiEx(Array.ConvertAll(ArabicData.ArabicSunLetters, Function(Str As String) ArabicData.MakeUniRegEx(Str))) + ")(?!" + ArabicData.MakeUniRegEx(ArabicData.ArabicShadda) + ")|" + ArabicData.MakeRegMultiEx(Array.ConvertAll(ArabicData.ArabicMoonLetters, Function(Str As String) ArabicData.MakeUniRegEx(Str))) + ")", _
            .Evaluator = "$2" + CStr(ArabicData.ArabicSukun), .NegativeMatch = "$1"}
    }
    Public Shared SimpleCleanScript As RuleTranslation() = { _
        New RuleTranslation With {.Rule = "RemoveDiacritics", .Match = ArabicData.MakeUniRegEx(ArabicData.ArabicShadda) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicSukun) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicKasra) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicDamma) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicFathatan) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicKasratan) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicDammatan) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterSuperscriptAlef), _
            .Evaluator = String.Empty}
    }
    Public Shared SimpleMinimalScript As RuleTranslation() = { _
        New RuleTranslation With {.Rule = "RemoveShadda", .Match = "((?:^\s*|\s+)(?:" + ArabicData.MakeRegMultiEx(Array.ConvertAll(ArabicData.ArabicLetters, Function(Str As String) ArabicData.MakeUniRegEx(Str))) + "))" + ArabicData.MakeUniRegEx(ArabicData.ArabicShadda) + "|((?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterQaf) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterKaf) + ")" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterKaf) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterDal) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterTeh) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterYeh) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterYeh) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterHeh) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterHeh) + ")" + ArabicData.MakeUniRegEx(ArabicData.ArabicShadda) + "(?=(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicKasra) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterKaf) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterHeh) + ")?(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicDamma) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + ")(?:(?:(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterNoon) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterHeh) + ArabicData.MakeUniRegEx(ArabicData.ArabicDamma) + "?)?" + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlef) + ")|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterNoon) + ArabicData.MakeUniRegEx(ArabicData.ArabicShadda) + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + "|(?:(?:(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterMeem) + ArabicData.MakeUniRegEx(ArabicData.ArabicDamma) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterWaw) + ")?" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterHeh) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterKaf) + ")(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicDamma) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + "))?(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterMeem) + "(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicSukun) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicDamma) + ")?)?)(?:\s*$|\s+))", _
            .Evaluator = "$1$2"}, _
        New RuleTranslation With {.Rule = "RemoveSukun", .Match = ArabicData.MakeUniRegEx(ArabicData.ArabicSukun), _
            .Evaluator = String.Empty}, _
        New RuleTranslation With {.Rule = "RemoveLongVowelDiacritics", .Match = "((?:^\s*|\s+)(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterKaf) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterWaw) + "|(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlefWithHamzaAbove) + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + ")?" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterFeh) + "))" + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + "(?=" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlef) + "(?!" + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicFathatan) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlef) + "?|" + ArabicData.MakeUniRegEx(ArabicData.ArabicDamma) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicShadda) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlef) + "?(?: " + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallHighJeem) + "| " + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallHighLamAlef) + ")? " + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlef) + ")(?:(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterTeh) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterTheh) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterDad) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterQaf) + ")" + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterHamza) + "|(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterJeem) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterBeh) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterReh) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterSeen) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterSad) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterTah) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterHah) + ")" + ArabicData.MakeUniRegEx(ArabicData.ArabicKasra) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterKaf) + ArabicData.MakeUniRegEx(ArabicData.ArabicKasra) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterHeh) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterHeh) + "(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicKasra) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicDamma) + ")|(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterDal) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterQaf) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterLam) + ")(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicDamma) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicKasra) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicKasratan) + ")|(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterZain) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAin) + ")(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicKasra) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + ")|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterNoon) + ArabicData.MakeUniRegEx(ArabicData.ArabicKasratan) + "))|((?:^\s*|\s+)" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterLam) + ")" + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + "(?=" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlef) + "(?! " + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlef) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterNoon) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterFeh) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterReh) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterTeh) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterBeh) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterKhah) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterSeen) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterSad) + "))|(^" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterLam) + ")" + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + "(?=" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlef) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterBeh) + ")(?!" + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicFathatan) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlef) + "?|" + ArabicData.MakeUniRegEx(ArabicData.ArabicDamma) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicShadda) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlef) + "?(?: " + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallHighJeem) + "| " + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallHighLamAlef) + ")? " + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlef) + ")|" + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + "(?=" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterSuperscriptAlef) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlefMaksura) + "(?!(?:\s*$|\s+" + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallHighLamAlef) + "| " + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallHighLigatureSadWithLamWithAlefMaksura) + "))(?!" + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicFathatan) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlef) + "?|" + ArabicData.MakeUniRegEx(ArabicData.ArabicDamma) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicShadda) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlef) + "?(?: " + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallHighJeem) + "| " + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallHighLamAlef) + ")? " + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlef) + "))|((?:^\s*|\s+)(?:(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterWaw) + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + ")?" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterTeh) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterWaw) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterFeh) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterKaf) + "))?" + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + "(?=" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlef) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterLam) + ")|((?:^\s*|\s+)(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterWaw) + "|(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlefWithHamzaAbove) + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + ")?" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterFeh) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterLam) + ")|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterYehWithHamzaAbove) + ")?" + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + "(?=" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlef) + "(?!" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterLam) + "(?!" + ArabicData.MakeUniRegEx(ArabicData.ArabicKasra) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicDamma) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicKasratan) + ")|" + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicFathatan) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlef) + "?|" + ArabicData.MakeUniRegEx(ArabicData.ArabicDamma) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicShadda) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlef) + "?(?: " + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallHighJeem) + "| " + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallHighLamAlef) + ")? " + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlef) + "))|(" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterJeem) + ")?" + ArabicData.MakeUniRegEx(ArabicData.ArabicKasra) + "(?=" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterYeh) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterHamza) + ")(?!" + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicFathatan) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlef) + "?|" + ArabicData.MakeUniRegEx(ArabicData.ArabicDamma) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicShadda) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlef) + "?(?: " + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallHighJeem) + "| " + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallHighLamAlef) + ")? " + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlef) + ")|(" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlefWithHamzaAbove) + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterKhah) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicKasra) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterNoon) + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterFeh) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterSeen) + "| " + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterThal) + ArabicData.MakeUniRegEx(ArabicData.ArabicKasra) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterKaf) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterReh) + ")" + ArabicData.MakeUniRegEx(ArabicData.ArabicKasra) + "(?=" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterYeh) + "(?!$)(?!" + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicFathatan) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlef) + "?|" + ArabicData.MakeUniRegEx(ArabicData.ArabicDamma) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicShadda) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlef) + "?(?: " + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallHighJeem) + "| " + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallHighLamAlef) + ")? " + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlef) + "))|(" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlefWithHamzaAbove) + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterKhah) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicKasra) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterNoon) + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterFeh) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterSeen) + "| " + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterThal) + ArabicData.MakeUniRegEx(ArabicData.ArabicKasra) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterKaf) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterReh) + ")?" + ArabicData.MakeUniRegEx(ArabicData.ArabicKasra) + "(?=" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterYeh) + "(?!" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterHamza) + ")(?!" + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicFathatan) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlef) + "?|" + ArabicData.MakeUniRegEx(ArabicData.ArabicDamma) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicShadda) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlef) + "?(?: " + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallHighJeem) + "| " + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallHighLamAlef) + ")? " + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlef) + "))|((?:^\s*|\s+)(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterWaw) + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + ")?" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlefWithHamzaAbove) + ")" + ArabicData.MakeUniRegEx(ArabicData.ArabicDamma) + "(?=" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterWaw) + "(?!" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterLam) + "(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + "(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterSuperscriptAlef) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlef) + "(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterTeh) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterHamza) + "))|" + ArabicData.MakeUniRegEx(ArabicData.ArabicKasra) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterYeh) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicDamma) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterWaw) + "))(?!" + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicFathatan) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlef) + "?|" + ArabicData.MakeUniRegEx(ArabicData.ArabicDamma) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicShadda) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlef) + "?(?: " + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallHighJeem) + "| " + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallHighLamAlef) + ")? " + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlef) + "))|((?:^\s*|\s+)(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterWaw) + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + ")?" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlefWithHamzaAbove) + ")?" + ArabicData.MakeUniRegEx(ArabicData.ArabicDamma) + "(?=" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterWaw) + "(?!" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterLam) + "(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterSuperscriptAlef) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterYehWithHamzaAbove) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicKasra) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterYeh) + " " + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlef) + "))(?!" + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicFathatan) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlef) + "?|" + ArabicData.MakeUniRegEx(ArabicData.ArabicDamma) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicShadda) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlef) + "?(?: " + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallHighJeem) + "| " + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallHighLamAlef) + ")? " + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlef) + "))", _
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

    'New RuleTranslation With {.Rule = "TehMarbutaPartiallyFlexible", .Match = "(" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlefWasla) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterLam) + "\S+)?" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterTehMarbuta) + "(?:" + ArabicData.MakeRegMultiEx(Array.ConvertAll(ArabicData.ArabicFathaDammaKasra, Function(Str As String) ArabicData.MakeUniRegEx(Str))) + ")",
    '                        .Evaluator = "$&", .NegativeMatch = "$1"}
    Public Shared PrefixPattern As String = ""
    'these cannot be empty strings
    Public Shared CertainStopPattern As String = ArabicData.MakeUniRegEx(ArabicData.ArabicSmallHighMeemInitialForm) + "|" + ArabicData.MakeRegMultiEx(Array.ConvertAll(ArabicData.PunctuationSymbols, Function(Ch As Char) ArabicData.MakeUniRegEx(Ch)))
    Public Shared CertainNotStopPattern As String = ArabicData.MakeRegMultiEx(Array.ConvertAll(ArabicData.RecitationLettersDiacritics, Function(Str As String) ArabicData.MakeUniRegEx(Str))) + "|" + ArabicData.MakeRegMultiEx(Array.ConvertAll(ArabicData.RecitationSpecialSymbolsNotStop, Function(Str As String) ArabicData.MakeUniRegEx(Str)))
    Public Shared OptionalStopPattern As String = ArabicData.MakeUniRegEx(ArabicData.ArabicEndOfAyah) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallHighJeem) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallHighThreeDots) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallHighLigatureQafWithLamWithAlefMaksura)
    Public Shared OptionalStopPatternNotEndOfAyah As String = ArabicData.MakeUniRegEx(ArabicData.ArabicSmallHighJeem) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallHighThreeDots) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallHighLigatureQafWithLamWithAlefMaksura) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallHighLigatureSadWithLamWithAlefMaksura)
    Public Shared OptionalNotStopPattern As String = ArabicData.MakeUniRegEx(ArabicData.ArabicSmallHighLigatureSadWithLamWithAlefMaksura)

    Public Shared ErrorCheckRules As RuleTranslation() = {
        New RuleTranslation With {.Rule = "NotValidStartEndCombination", .Match = ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlefWasla) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterLam) + "\S+(?:" + ArabicData.MakeRegMultiEx(Array.ConvertAll(ArabicData.ArabicTanweens, Function(Str As String) ArabicData.MakeUniRegEx(Str))) + ")",
                                    .Evaluator = "$&"}, _
        New RuleTranslation With {.Rule = "NotValidEnd", .Match = "(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlef) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicKasra) + "(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlefMaksura) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallYeh) + ")|" + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlefMaksura) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterSuperscriptAlef) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicDamma) + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallWaw) + ")(?:(?:\s+" + ArabicData.MakeUniRegEx(ArabicData.ArabicEndOfAyah) + ")?(?:\s+(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicStartOfRubElHizb) + "|" + ArabicData.MakeRegMultiEx(Array.ConvertAll(ArabicData.ArabicStopLetters, Function(Str As String) ArabicData.MakeUniRegEx(Str))) + "))?)?\s+(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlefWithHamzaAbove) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlefWithHamzaBelow) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterHamza) + ")",
                                 .Evaluator = "$&"}, _
        New RuleTranslation With {.Rule = "NotValidEnd", .Match = "(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicMaddahAbove) + ")(?:(?:\s+" + ArabicData.MakeUniRegEx(ArabicData.ArabicEndOfAyah) + ")?(?:\s+(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicStartOfRubElHizb) + "|" + ArabicData.MakeRegMultiEx(Array.ConvertAll(ArabicData.ArabicStopLetters, Function(Str As String) ArabicData.MakeUniRegEx(Str))) + "))?)?\s+(" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlefWithHamzaAbove) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlefWithHamzaBelow) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterHamza) + ")",
                                 .Evaluator = "$&", .NegativeMatch = "$1"}, _
        New RuleTranslation With {.Rule = "OnlyAtEndOfWord", .Match = "(" + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlef) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlefMaksura) + "(?=" + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + "(?:\s*(?:$|" + CertainStopPattern + "|" + OptionalStopPattern + ")|\s+)))|(" + ArabicData.MakeUniRegEx(ArabicData.ArabicSukun) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicKasra) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterSuperscriptAlef) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallHighRoundedZero) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicEmptyCentreLowStop) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicFathatan) + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallLowMeem) + "?)?" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlefMaksura) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlefMaksura) + "(?!(?:" + ArabicData.MakeRegMultiEx(Array.ConvertAll(ArabicData.ArabicFathaDammaKasra, Function(Str As String) ArabicData.MakeUniRegEx(Str))) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicDammatan) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicShadda) + "(?:" + ArabicData.MakeRegMultiEx(Array.ConvertAll(ArabicData.ArabicFathaDammaKasra, Function(Str As String) ArabicData.MakeUniRegEx(Str))) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicDammatan) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicKasratan) + ")|" + ArabicData.MakeUniRegEx(ArabicData.ArabicMaddahAbove) + "(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicDamma) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicDammatan) + ")?|" + ArabicData.MakeUniRegEx(ArabicData.ArabicSukun) + "(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicHamzaAbove) + "(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicDammatan) + ArabicData.MakeUniRegEx(ArabicData.ArabicKasratan) + "))?|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterSuperscriptAlef) + "(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicMaddahAbove) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterTehMarbuta) + "(?:" + ArabicData.MakeRegMultiEx(Array.ConvertAll(ArabicData.ArabicFathaDammaKasra, Function(Str As String) ArabicData.MakeUniRegEx(Str))) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicFathatan) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicKasratan) + ")|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterMeem) + ArabicData.MakeUniRegEx(ArabicData.ArabicKasra) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterNoon) + "(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlef) + ArabicData.MakeUniRegEx(ArabicData.ArabicMaddahAbove) + "?|" + ArabicData.MakeUniRegEx(ArabicData.ArabicKasra) + "(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterYeh) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterHeh) + ArabicData.MakeUniRegEx(ArabicData.ArabicDamma) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlefMaksura) + "(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicMaddahAbove) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + ")?)?)|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterHeh) + "(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicKasra) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterNoon) + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlefMaksura) + ArabicData.MakeUniRegEx(ArabicData.ArabicKasra) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlef) + ArabicData.MakeUniRegEx(ArabicData.ArabicMaddahAbove) + "?|" + ArabicData.MakeUniRegEx(ArabicData.ArabicDamma) + "(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterNoon) + ArabicData.MakeUniRegEx(ArabicData.ArabicShadda) + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterMeem) + "(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicSukun) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicDamma) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlef) + ")?)?)|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterKaf) + "(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + "(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterHeh) + ArabicData.MakeUniRegEx(ArabicData.ArabicDamma) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterMeem) + ")?|" + ArabicData.MakeUniRegEx(ArabicData.ArabicKasra) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicDamma) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterMeem) + "(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicSukun) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicDamma) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlef) + ")?)|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterLam) + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + ")?)?(?:\s*(?:$|" + CertainStopPattern + "|" + OptionalStopPattern + ")|\s+))", _
                                    .Evaluator = "$&", .NegativeMatch = "$1$2"}, _
        New RuleTranslation With {.Rule = "OnlyAtEndOfWord", .Match = "(" + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + "(?:(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlefMaksura) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterWaw) + ")" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterSuperscriptAlef) + ")?)?" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterTehMarbuta), _
                                    .Evaluator = "$&", .NegativeMatch = "$1"}, _
        New RuleTranslation With {.Rule = "OnlyAtEndOfWord", .Match = ArabicData.MakeUniRegEx(ArabicData.ArabicLetterTehMarbuta) + "(?!(?:" + ArabicData.MakeRegMultiEx(Array.ConvertAll(ArabicData.ArabicFathaDammaKasra, Function(Str As String) ArabicData.MakeUniRegEx(Str))) + "|(?:" + ArabicData.MakeRegMultiEx(Array.ConvertAll(ArabicData.ArabicTanweens, Function(Str As String) ArabicData.MakeUniRegEx(Str))) + ")(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallLowMeem) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallHighMeemIsolatedForm) + ")?)(?:\s*(?:$|" + CertainStopPattern + "|" + OptionalStopPattern + ")|\s+))", _
                                    .Evaluator = "$&"}, _
        New RuleTranslation With {.Rule = "OnlyAtEndOfWord", .Match = "(" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterTehMarbuta) + ")?" + ArabicData.MakeUniRegEx(ArabicData.ArabicFathatan) + "(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallLowMeem) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallHighMeemIsolatedForm) + ")?(?:\s+|\s*(?:$|" + CertainStopPattern + "|" + OptionalStopPattern + "))|(" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterHamza) + ")" + ArabicData.MakeUniRegEx(ArabicData.ArabicFathatan) + "(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallLowMeem) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallHighMeemIsolatedForm) + ")?" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlef) + "?(?:\s+|\s*(?:$|" + CertainStopPattern + "|" + OptionalStopPattern + "))|(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlef) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlefMaksura) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterTehMarbuta) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterHamza) + ")" + ArabicData.MakeUniRegEx(ArabicData.ArabicFathatan) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicFathatan) + "(?!(?:(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallLowMeem) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallHighMeemIsolatedForm) + ")?" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlef) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallLowMeem) + "?" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlefMaksura) + ")(?:\s+|\s*(?:$|" + CertainStopPattern + "|" + OptionalStopPattern + ")))", _
                                    .Evaluator = "$&", .NegativeMatch = "$1$2"}, _
        New RuleTranslation With {.Rule = "OnlyAtStartOfWord", .Match = "(?:(?:^\s*|\s+)(" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterWaw) + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + "|(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterWaw) + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + ")?" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterTeh) + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + "|(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterWaw) + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterLam) + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlefWithHamzaAbove) + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + "(" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterFeh) + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + ")?)?" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterBeh) + ArabicData.MakeUniRegEx(ArabicData.ArabicKasra) + "|(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlefWithHamzaAbove) + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + ")?" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterFeh) + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterLam) + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterWaw) + ArabicData.MakeUniRegEx(ArabicData.ArabicShadda) + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterLam) + ArabicData.MakeUniRegEx(ArabicData.ArabicShadda) + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterKaf) + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlefWasla) + "?" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterLam) + ArabicData.MakeUniRegEx(ArabicData.ArabicKasra) + ")|\S+)" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlefWasla) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlefWasla) + "(?=" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlefWasla) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlefWithHamzaAbove) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlefWithHamzaBelow) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlefMaksura) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlef) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterSuperscriptAlef) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterHamza) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterTehMarbuta) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterWaw) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterYeh) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicKasra) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicDamma) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicSukun) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicShadda) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicMaddahAbove) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicFathatan) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicKasratan) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicDammatan) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicHamzaAbove) + ")", _
                                .Evaluator = "$&", .NegativeMatch = "$1"}, _
        New RuleTranslation With {.Rule = "NotValidCombination", .Match = "(" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlef) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterSuperscriptAlef) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlefMaksura) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallWaw) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallYeh) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterSeen) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterSad) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAin) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterQaf) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterKaf) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterLam) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterMeem) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterNoon) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterWaw) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterYeh) + ")?" + ArabicData.MakeUniRegEx(ArabicData.ArabicMaddahAbove), _
                                .Evaluator = "$&", .NegativeMatch = "$1"}, _
        New RuleTranslation With {.Rule = "NotValidCombination", .Match = ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlef) + "(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicKasra) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicDamma) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicSukun) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicShadda) + ")|(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlef) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterYeh) + ")" + ArabicData.MakeUniRegEx(ArabicData.ArabicDammatan) + "|(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlef) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterWaw) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlefMaksura) + ")" + ArabicData.MakeUniRegEx(ArabicData.ArabicKasratan) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicKasratan) + "(?!(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallLowMeem) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallHighMeemIsolatedForm) + ")?(?:\s+|\s*(?:$|" + CertainStopPattern + "|" + OptionalStopPattern + ")))|" + ArabicData.MakeUniRegEx(ArabicData.ArabicDammatan) + "(?!(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallLowMeem) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallHighMeemIsolatedForm) + ")?(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlef) + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallHighRoundedZero) + "?)?(?:\s+|\s*(?:$|" + CertainStopPattern + "|" + OptionalStopPattern + ")))", _
                                .Evaluator = "$&"}, _
        New RuleTranslation With {.Rule = "NotValidCombination", .Match = "(?:\S+(" + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicKasra) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicSukun) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicMaddahAbove) + ")|\S+)" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlefWithHamzaBelow) + "|(?:\S+(" + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicKasra) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicSukun) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicMaddahAbove) + ")|\S+)" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlefWithHamzaAbove) + "|(?:\S+(" + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicKasra) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicDamma) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicSukun) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicMaddahAbove) + ")|\S+)" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterHamza) + "|(" + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlefWasla) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicDamma) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicMaddahAbove) + ")?" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterWawWithHamzaAbove) + "|(" + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicKasra) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicDamma) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlef) + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallHighRoundedZero) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlefWasla) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicSukun) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicMaddahAbove) + ")?" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterYehWithHamzaAbove) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlefWithHamzaBelow) + "(?!" + ArabicData.MakeUniRegEx(ArabicData.ArabicKasra) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicKasratan) + ")|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlefWithHamzaAbove) + "(?!" + ArabicData.MakeUniRegEx(ArabicData.ArabicSukun) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicDamma) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicDammatan) + ")|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterWawWithHamzaAbove) + "(?!" + ArabicData.MakeUniRegEx(ArabicData.ArabicSukun) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicDamma) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicKasra) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicFathatan) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicDammatan) + ")|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterYehWithHamzaAbove) + "(?!" + ArabicData.MakeUniRegEx(ArabicData.ArabicSukun) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicDamma) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicKasra) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicFathatan) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicKasratan) + ")|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterHamza) + "(?!" + ArabicData.MakeUniRegEx(ArabicData.ArabicSukun) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicDamma) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicKasra) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicFathatan) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicKasratan) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicDammatan) + ")", _
                                .Evaluator = "$&", .NegativeMatch = "$1$2$3$4$5"}, _
        New RuleTranslation With {.Rule = "NotValidCombination", .Match = "(" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterWaw) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlefMaksura) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + ")?" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterSuperscriptAlef), _
                                .Evaluator = "$&", .NegativeMatch = "$1"}, _
        New RuleTranslation With {.Rule = "NotValidCombination", .Match = ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlef) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterLam) + ArabicData.MakeUniRegEx(ArabicData.ArabicMaddahAbove) + "(" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterReh) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterMeem) + ArabicData.MakeUniRegEx(ArabicData.ArabicMaddahAbove) + "(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterSad) + ArabicData.MakeUniRegEx(ArabicData.ArabicMaddahAbove) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterReh) + ")?)|(?:(" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterWaw) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallWaw) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicDammatan) + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallLowMeem) + "?|" + ArabicData.MakeUniRegEx(ArabicData.ArabicFathatan) + "(" + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallHighMeemIsolatedForm) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallLowMeem) + ")?|" + ArabicData.MakeUniRegEx(ArabicData.ArabicMaddahAbove) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicKasra) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicDamma) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterSuperscriptAlef) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicSukun) + ")?|(?:^\s*|\s+))" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlef), _
                                .Evaluator = "$&", .NegativeMatch = "$1$2"}, _
        New RuleTranslation With {.Rule = "NotValidCombination", .Match = "(?:\S+(" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlef) + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallHighRoundedZero) + "?|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterYeh) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterHeh) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicKasra) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicDamma) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterSuperscriptAlef) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicSukun) + ")|\S+)" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterYeh), _
                                .Evaluator = "$&", .NegativeMatch = "$1"}, _
        New RuleTranslation With {.Rule = "NotValidCombination", .Match = "(?:\S+(" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlef) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicKasra) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicDamma) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterSuperscriptAlef) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicSukun) + ")|\S+)" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterWaw), _
                                .Evaluator = "$&", .NegativeMatch = "$1"}, _
        New RuleTranslation With {.Rule = "NotValidCombination", .Match = ArabicData.MakeUniRegEx(ArabicData.ArabicLetterGhain) + ArabicData.MakeUniRegEx(ArabicData.ArabicShadda) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterWaw) + "(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterGhain) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlefMaksura) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterWaw) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterYeh) + ")|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterYeh) + "(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlef) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlefMaksura) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterWaw) + ")", _
                                .Evaluator = "$&"}, _
        New RuleTranslation With {.Rule = "NotValidCombination", .Match = "(" + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterSuperscriptAlef) + ArabicData.MakeUniRegEx(ArabicData.ArabicHamzaAbove) + ArabicData.MakeUniRegEx(ArabicData.ArabicSukun) + ")|(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicKasra) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicDamma) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicFathatan) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicKasratan) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicDammatan) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterSuperscriptAlef) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicSukun) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicMaddahAbove) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicHamzaAbove) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicShadda) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicTatweel) + ")" + ArabicData.MakeUniRegEx(ArabicData.ArabicShadda) + "|(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicKasra) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicDamma) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicFathatan) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicKasratan) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicDammatan) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterSuperscriptAlef) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicSukun) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicMaddahAbove) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicHamzaAbove) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicTatweel) + ")(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicKasratan) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicDammatan) + ")|(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicKasra) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicDamma) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicFathatan) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicKasratan) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicDammatan) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterSuperscriptAlef) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicSukun) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicMaddahAbove) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicTatweel) + ")" + ArabicData.MakeUniRegEx(ArabicData.ArabicFathatan) + "|(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicKasra) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicDamma) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicFathatan) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicKasratan) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicDammatan) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicSukun) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicMaddahAbove) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicHamzaAbove) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicShadda) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicTatweel) + ")" + ArabicData.MakeUniRegEx(ArabicData.ArabicMaddahAbove) + "|(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicKasra) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicDamma) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicFathatan) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicKasratan) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicDammatan) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterSuperscriptAlef) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicSukun) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicMaddahAbove) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicShadda) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicTatweel) + ")" + ArabicData.MakeUniRegEx(ArabicData.ArabicSukun) + "|(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicKasra) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicDamma) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicFathatan) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicKasratan) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicDammatan) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterSuperscriptAlef) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicSukun) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicMaddahAbove) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicTatweel) + ")(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicKasra) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicDamma) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + ")|(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicKasra) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicDamma) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicFathatan) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicKasratan) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicDammatan) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterSuperscriptAlef) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicSukun) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicMaddahAbove) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicHamzaAbove) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicShadda) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicTatweel) + ")" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterSuperscriptAlef) + "|(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicKasra) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicDamma) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicFathatan) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicKasratan) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicDammatan) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterSuperscriptAlef) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicSukun) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicMaddahAbove) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicHamzaAbove) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicShadda) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicTatweel) + ")" + ArabicData.MakeUniRegEx(ArabicData.ArabicShadda) + "|(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicKasra) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicDamma) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicFathatan) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicKasratan) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicDammatan) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterSuperscriptAlef) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicSukun) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicMaddahAbove) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicHamzaAbove) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicShadda) + ")" + ArabicData.MakeUniRegEx(ArabicData.ArabicHamzaAbove) + "|(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicFathatan) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicKasratan) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicDammatan) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterSuperscriptAlef) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicHamzaAbove) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicShadda) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicTatweel) + ")" + ArabicData.MakeUniRegEx(ArabicData.ArabicTatweel), _
                                .Evaluator = "$&", .NegativeMatch = "$1"}, _
        New RuleTranslation With {.Rule = "MissingDiacritic", .Match = "(" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterKaf) + ArabicData.MakeUniRegEx(ArabicData.ArabicMaddahAbove) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterHeh) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterYeh) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAin) + ArabicData.MakeUniRegEx(ArabicData.ArabicMaddahAbove) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterSad) + ArabicData.MakeUniRegEx(ArabicData.ArabicMaddahAbove) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterHah) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterMeem) + ArabicData.MakeUniRegEx(ArabicData.ArabicMaddahAbove) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterTah) + "(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterHeh) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterSeen) + ArabicData.MakeUniRegEx(ArabicData.ArabicMaddahAbove) + "(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterMeem) + ArabicData.MakeUniRegEx(ArabicData.ArabicMaddahAbove) + ")?)(?:\s*$|\s+)|(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterHeh) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterHeh) + "|(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterKaf) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterQaf) + ")" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterKaf) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterDal) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterTeh) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterYeh) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterYeh) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterLam) + "(?:" + ArabicData.MakeRegMultiEx(Array.ConvertAll(ArabicData.ArabicSunLetters, Function(Str As String) ArabicData.MakeUniRegEx(Str))) + "))" + ArabicData.MakeUniRegEx(ArabicData.ArabicShadda) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterNoon) + "(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterTeh) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterTheh) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterJeem) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterDal) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterThal) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterZain) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterSeen) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterSheen) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterSad) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterDad) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterTah) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterZah) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterFeh) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterQaf) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterKaf) + ")|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterTah) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterTeh) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterWaw) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlef) + ")|(?:" + ArabicData.MakeRegMultiEx(Array.ConvertAll(ArabicData.ArabicSunLetters, Function(Str As String) ArabicData.MakeUniRegEx(Str))) + "|" + ArabicData.MakeRegMultiEx(Array.ConvertAll(ArabicData.ArabicMoonLettersNoVowels, Function(Str As String) ArabicData.MakeUniRegEx(Str))) + "|(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicSukun) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicDamma) + ")" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterYeh) + "|(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicSukun) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicKasra) + ")" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterWaw) + ")(?:" + ArabicData.MakeRegMultiEx(Array.ConvertAll(ArabicData.ArabicSunLetters, Function(Str As String) ArabicData.MakeUniRegEx(Str))) + "|" + ArabicData.MakeRegMultiEx(Array.ConvertAll(ArabicData.ArabicMoonLetters, Function(Str As String) ArabicData.MakeUniRegEx(Str))) + ")|(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterHeh) + ArabicData.MakeUniRegEx(ArabicData.ArabicSukun) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterHeh) + "|(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterKaf) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterQaf) + ")" + ArabicData.MakeUniRegEx(ArabicData.ArabicSukun) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterKaf) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterDal) + ArabicData.MakeUniRegEx(ArabicData.ArabicSukun) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterTeh) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterYeh) + ArabicData.MakeUniRegEx(ArabicData.ArabicSukun) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterYeh) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterLam) + ArabicData.MakeUniRegEx(ArabicData.ArabicSukun) + "(?:" + ArabicData.MakeRegMultiEx(Array.ConvertAll(ArabicData.ArabicSunLetters, Function(Str As String) ArabicData.MakeUniRegEx(Str))) + "))" + ArabicData.MakeUniRegEx(ArabicData.ArabicShadda) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterNoon) + ArabicData.MakeUniRegEx(ArabicData.ArabicSukun) + "(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterTeh) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterTheh) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterJeem) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterDal) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterThal) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterZain) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterSeen) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterSheen) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterSad) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterDad) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterTah) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterZah) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterFeh) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterQaf) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterKaf) + ")|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterTah) + ArabicData.MakeUniRegEx(ArabicData.ArabicSukun) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterTeh), _
                                .Evaluator = "$&", .NegativeMatch = "$1"}, _
        New RuleTranslation With {.Rule = "MissingEndOfWordDiacritic", .Match = "((?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterWaw) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlef) + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallHighRoundedZero) + "\s+" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterWaw) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterBeh) + "\s+(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterBeh) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterMeem) + ")|" + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterWaw) + "\s+" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterWaw) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterThal) + "\s+(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterThal) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterZah) + ")|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAin) + "\s+" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAin) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterFeh) + "\s+" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterFeh) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterReh) + "\s+" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterReh) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterMeem) + "\s+" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterMeem) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterNoon) + "\s+(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterNoon) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterLam) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterReh) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterMeem) + ")|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterTeh) + "\s+(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterTeh) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterDal) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterTah) + ")|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterDal) + "\s+(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterDal) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterTeh) + ")|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterLam) + "\s+(?:" + ArabicData.MakeRegMultiEx(Array.ConvertAll(ArabicData.ArabicSunLetters, Function(Str As String) ArabicData.MakeUniRegEx(Str))) + "))" + ArabicData.MakeUniRegEx(ArabicData.ArabicShadda) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterMeem) + "\s+" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterBeh) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterNoon) + "\s+(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterTeh) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterTheh) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterJeem) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterDal) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterThal) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterZain) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterSeen) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterSheen) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterSad) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterDad) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterTah) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterZah) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterFeh) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterQaf) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterKaf) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterWaw) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterYeh) + "))|(?:" + ArabicData.MakeRegMultiEx(Array.ConvertAll(ArabicData.ArabicSunLetters, Function(Str As String) ArabicData.MakeUniRegEx(Str))) + "|" + ArabicData.MakeRegMultiEx(Array.ConvertAll(ArabicData.ArabicMoonLettersNoVowels, Function(Str As String) ArabicData.MakeUniRegEx(Str))) + "|(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicSukun) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicDamma) + ")" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterYeh) + "|(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicSukun) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicKasra) + ")" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterWaw) + ")\s+(?:$|" + ArabicData.MakeRegMultiEx(Array.ConvertAll(ArabicData.ArabicSunLetters, Function(Str As String) ArabicData.MakeUniRegEx(Str))) + "|" + ArabicData.MakeRegMultiEx(Array.ConvertAll(ArabicData.ArabicMoonLetters, Function(Str As String) ArabicData.MakeUniRegEx(Str))) + ")|" + ArabicData.MakeUniRegEx(ArabicData.ArabicSukun) + "\s+" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlefWasla) + "|(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterWaw) + ArabicData.MakeUniRegEx(ArabicData.ArabicSukun) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlef) + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallHighRoundedZero) + "\s+" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterWaw) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterBeh) + ArabicData.MakeUniRegEx(ArabicData.ArabicSukun) + "\s+(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterBeh) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterMeem) + ")|" + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterWaw) + ArabicData.MakeUniRegEx(ArabicData.ArabicSukun) + "\s+" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterWaw) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterThal) + ArabicData.MakeUniRegEx(ArabicData.ArabicSukun) + "\s+(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterThal) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterZah) + ")|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAin) + ArabicData.MakeUniRegEx(ArabicData.ArabicSukun) + "\s+" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAin) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterFeh) + ArabicData.MakeUniRegEx(ArabicData.ArabicSukun) + "\s+" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterFeh) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterReh) + ArabicData.MakeUniRegEx(ArabicData.ArabicSukun) + "\s+" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterReh) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterMeem) + ArabicData.MakeUniRegEx(ArabicData.ArabicSukun) + "\s+" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterMeem) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterNoon) + ArabicData.MakeUniRegEx(ArabicData.ArabicSukun) + "\s+(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterNoon) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterLam) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterReh) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterMeem) + ")|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterTeh) + ArabicData.MakeUniRegEx(ArabicData.ArabicSukun) + "\s+(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterTeh) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterDal) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterTah) + ")|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterDal) + ArabicData.MakeUniRegEx(ArabicData.ArabicSukun) + "\s+(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterDal) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterTeh) + ")|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterLam) + ArabicData.MakeUniRegEx(ArabicData.ArabicSukun) + "\s+(?:" + ArabicData.MakeRegMultiEx(Array.ConvertAll(ArabicData.ArabicSunLetters, Function(Str As String) ArabicData.MakeUniRegEx(Str))) + "))" + ArabicData.MakeUniRegEx(ArabicData.ArabicShadda) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterMeem) + ArabicData.MakeUniRegEx(ArabicData.ArabicSukun) + "\s+" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterBeh) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterNoon) + ArabicData.MakeUniRegEx(ArabicData.ArabicSukun) + "\s+(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterTeh) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterTheh) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterJeem) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterDal) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterThal) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterZain) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterSeen) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterSheen) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterSad) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterDad) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterTah) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterZah) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterFeh) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterQaf) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterKaf) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterWaw) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterYeh) + ")", _
                                .Evaluator = "$&", .NegativeMatch = "$1"}, _
        New RuleTranslation With {.Rule = "NotValidStartEnd", .Match = "(?:^\s*|\s+)(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicKasra) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicDamma) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicFathatan) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicKasratan) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicDammatan) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicSukun) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicMaddahAbove) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicHamzaAbove) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicShadda) + ")|(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterJeem) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterHah) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterKhah) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterZain) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterSeen) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterSheen) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterSad) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterDad) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterTah) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterGhain) + ")(?=\s*$|\s+)", _
                                .Evaluator = "$&"}, _
        New RuleTranslation With {.Rule = "SmallMeem", .Match = "(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicKasratan) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicDammatan) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicFathatan) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlef) + "|(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterHamza) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterTehMarbuta) + ")" + ArabicData.MakeUniRegEx(ArabicData.ArabicFathatan) + ")\s+" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterBeh), _
                                .Evaluator = "$&"}
        }
    'New RuleTranslation With {.Rule = "NeedsRecomposition", .Match = "(" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlef) + ArabicData.MakeUniRegEx(ArabicData.ArabicMaddahAbove) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterHamza) + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlef) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlef) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterHamza) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterWaw) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterHamza) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlefMaksura) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterHamza) + ")", _
    '                            .Evaluator = "$&"}
    '} '"Missing diacritic", "Can only appear at end of word", "Must not appear at beginning of word", "Must appear at beginning of word", "Not a valid combination", "Needs to be recomposed"

    Public Enum RuleFuncs As Integer
        eNone
        eUpperCase
        eSpellNumber
        eSpellLetter
        eLookupLetter
        eLookupLongVowel
        eDivideTanween
        eDivideLetterSymbol
        eStopOption
        eLeadingGutteral
        eTrailingGutteral
    End Enum
    Public Delegate Function RuleFunction(Str As String, Scheme As String) As String()
    Public Shared RuleFunctions As RuleFunction() = {
        Function(Str As String, Scheme As String) {UCase(Str)},
        Function(Str As String, Scheme As String) {TransliterateWithRules(ArabicData.TransliterateFromBuckwalter(Arabic.ArabicWordFromNumber(CInt(TransliterateToScheme(Str, ArabicData.TranslitScheme.Literal, String.Empty)), True, False, False)), Scheme, Nothing)},
        Function(Str As String, Scheme As String) {TransliterateWithRules(ArabicLetterSpelling(Str, True), Scheme, Nothing)},
        Function(Str As String, Scheme As String) {GetSchemeValueFromSymbol(ArabicData.Data.ArabicLetters(ArabicData.FindLetterBySymbol(Str.Chars(0))), Scheme)},
        Function(Str As String, Scheme As String) {GetSchemeLongVowelFromString(Str, Scheme)},
        Function(Str As String, Scheme As String) {ArabicData.ArabicFathaDammaKasra(Array.IndexOf(ArabicData.ArabicTanweens, Str)), ArabicData.ArabicLetterNoon},
        Function(Str As String, Scheme As String) {String.Empty, String.Empty},
        Function(Str As String, Scheme As String) {String.Empty},
        Function(Str As String, Scheme As String) {GetSchemeGutteralFromString(Str.Remove(Str.Length - 1), Scheme, True) + Str.Chars(Str.Length - 1)},
        Function(Str As String, Scheme As String) {Str.Chars(0) + GetSchemeGutteralFromString(Str.Remove(0, 1), Scheme, False)}
    }
    'Javascript does not support negative or positive lookbehind in regular expressions
    Public Shared RomanizationRules As RuleTranslation() = { _
        New RuleTranslation With {.Rule = "Shadda", .Match = "(" + ArabicData.MakeRegMultiEx(Array.ConvertAll(ArabicData.ArabicLetters, Function(Str As String) ArabicData.MakeUniRegEx(Str))) + ")" + ArabicData.MakeUniRegEx(ArabicData.ArabicShadda), _
                                    .Evaluator = "$1-$1"}, _
        New RuleTranslation With {.Rule = "GutteralRules", .Match = "(?:" + ArabicData.MakeRegMultiEx(Array.ConvertAll(ArabicData.ArabicSpecialGutteral, Function(Str As String) ArabicData.MakeUniRegEx(Str))) + ")(" + ArabicData.MakeUniRegEx(ArabicData.ArabicSukun) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlef) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicDamma) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterWaw) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicDamma) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicKasra) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterYeh) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicKasra) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicFathatan) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicDammatan) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicKasratan) + ")", _
                                    .Evaluator = "$&", .RuleFunc = RuleFuncs.eTrailingGutteral}, _
        New RuleTranslation With {.Rule = "LeadingGutteralRules", .Match = "(" + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlef) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicKasra) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterYeh) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicKasra) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicDamma) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterWaw) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicDamma) + ")(?:" + ArabicData.MakeRegMultiEx(Array.ConvertAll(ArabicData.ArabicSpecialLeadingGutteral, Function(Str As String) ArabicData.MakeUniRegEx(Str))) + ")", _
                                    .Evaluator = "$&", .RuleFunc = RuleFuncs.eLeadingGutteral}, _
        New RuleTranslation With {.Rule = "LongVowels", .Match = ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlef) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicDamma) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterWaw) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicKasra) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterYeh), _
                                    .Evaluator = "$&", .RuleFunc = RuleFuncs.eLookupLongVowel}, _
        New RuleTranslation With {.Rule = "ResolveAmbiguity", .Match = "(" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterSeen) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterTeh) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterTah) + ")(" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterHeh) + ")", _
                                  .Evaluator = "$1-$2"}, _
        New RuleTranslation With {.Rule = "LettersTanweensVowelsHamza", .Match = "(" + ArabicData.MakeRegMultiEx(Array.ConvertAll(ArabicData.ArabicLetters, Function(Str As String) ArabicData.MakeUniRegEx(Str))) + "|" + ArabicData.MakeRegMultiEx(Array.ConvertAll(ArabicData.ArabicTanweens, Function(Str As String) ArabicData.MakeUniRegEx(Str))) + "|" + ArabicData.MakeRegMultiEx(Array.ConvertAll(ArabicData.ArabicFathaDammaKasra, Function(Str As String) ArabicData.MakeUniRegEx(Str))) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterHamza) + ")", _
                                    .Evaluator = "$&", .RuleFunc = RuleFuncs.eLookupLetter}, _
        New RuleTranslation With {.Rule = "Punctuation", .Match = "(" + ArabicData.MakeRegMultiEx(Array.ConvertAll(ArabicData.ArabicPunctuationSymbols, Function(Str As String) ArabicData.MakeUniRegEx(Str))) + ")", _
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
                                  .Evaluator = ArabicData.ArabicLetterHeh}, _
        New RuleTranslation With {.Rule = "HelperTeh", .Match = "helperteh",
                                  .Evaluator = ArabicData.ArabicLetterTeh}, _
        New RuleTranslation With {.Rule = "LetterSpelling", .Match = "spellletter", _
                                    .Evaluator = "{0}", .RuleFunc = RuleFuncs.eSpellLetter}, _
        New RuleTranslation With {.Rule = "NumberSpelling", .Match = "spellnumber", _
                                    .Evaluator = "{0}", .RuleFunc = RuleFuncs.eSpellNumber}, _
        New RuleTranslation With {.Rule = "HelperPrefix", .Match = "helperprefix", _
                                    .Evaluator = "{0}-"}, _
        New RuleTranslation With {.Rule = "HelperFatha", .Match = "helperfatha", _
                                    .Evaluator = ArabicData.ArabicFatha}, _
        New RuleTranslation With {.Rule = "HelperKasra", .Match = "helperkasra", _
                                    .Evaluator = ArabicData.ArabicKasra}, _
        New RuleTranslation With {.Rule = "HelperDamma", .Match = "helperdamma", _
                                    .Evaluator = ArabicData.ArabicDamma}, _
        New RuleTranslation With {.Rule = "HelperAlef", .Match = "helperalef", _
                                    .Evaluator = ArabicData.ArabicLetterAlef}, _
        New RuleTranslation With {.Rule = "HelperWaw", .Match = "helperwaw", _
                                    .Evaluator = ArabicData.ArabicLetterWaw}, _
        New RuleTranslation With {.Rule = "HelperYeh", .Match = "helperyeh", _
                                    .Evaluator = ArabicData.ArabicLetterYeh}, _
        New RuleTranslation With {.Rule = "HelperMeem", .Match = "helpermeem", _
                                    .Evaluator = ArabicData.ArabicLetterMeem}, _
        New RuleTranslation With {.Rule = "HelperNoon", .Match = "helpernoon", _
                                    .Evaluator = ArabicData.ArabicLetterNoon}, _
        New RuleTranslation With {.Rule = "HelperSeen", .Match = "helperseen", _
                                    .Evaluator = ArabicData.ArabicLetterSeen}, _
        New RuleTranslation With {.Rule = "HelperHamza", .Match = "helperhamza", _
                                    .Evaluator = ArabicData.ArabicLetterHamza}, _
        New RuleTranslation With {.Rule = "HelperMadda", .Match = "helpermadda", _
                                    .Evaluator = ArabicData.ArabicLetterHamza + ArabicData.ArabicFatha + ArabicData.ArabicLetterAlef}, _
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
    Public Shared RulesOfRecitationRegEx As RuleMetadataTranslation() = { _
        New RuleMetadataTranslation With {.Rule = "LetterSpelling", .Match = "((?:^|" + CertainStopPattern + "|" + OptionalStopPattern + ")\s*|\s+)(" + ArabicData.MakeRegMultiEx(Array.ConvertAll(ArabicData.ArabicUniqueLetters, Function(Str As String) ArabicData.MakeUniRegEx(ArabicData.TransliterateFromBuckwalter(Str)))) + ")(?=\s*$|\s+)", _
            .Evaluator = New String() {Nothing, "spellletter"}}, _
        New RuleMetadataTranslation With {.Rule = "NumberSpelling", .Match = "((?:^|" + CertainStopPattern + "|" + OptionalStopPatternNotEndOfAyah + ")\s*|\s+)((?:" + ArabicData.MakeRegMultiEx(Array.ConvertAll(ArabicData.ArabicNumbers, Function(Str As String) ArabicData.MakeUniRegEx(ArabicData.TransliterateFromBuckwalter(Str)))) + ")+)(?=\s*$|\s+)", _
            .Evaluator = New String() {Nothing, "spellnumber"}}, _
        New RuleMetadataTranslation With {.Rule = "NumberSpelling", .Match = "(" + ArabicData.MakeUniRegEx(ArabicData.ArabicEndOfAyah) + ")()((?:" + ArabicData.MakeRegMultiEx(Array.ConvertAll(ArabicData.ArabicNumbers, Function(Str As String) ArabicData.MakeUniRegEx(ArabicData.TransliterateFromBuckwalter(Str)))) + ")+)()(?=\s*$|\s+)", _
            .Evaluator = New String() {"empty", "helperlparen", "spellnumber", "helperrparen"}}, _
        New RuleMetadataTranslation With {.Rule = "Stopping", .Match = "(?:(" + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallHighMeemInitialForm) + ")|(" + ArabicData.MakeUniRegEx(ArabicData.ArabicStartOfRubElHizb) + ")|(" + ArabicData.MakeUniRegEx(ArabicData.ArabicEndOfAyah) + ")|(" + ArabicData.MakeUniRegEx(ArabicData.ArabicPlaceOfSajdah) + "))(?=\s*$|\s+)",
            .Evaluator = New String() {"compulsorystop", "startofhizb", "endofversestop", "prostration"}}, _
        New RuleMetadataTranslation With {.Rule = "Stopping", .Match = "(" + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallHighJeem) + ")|(" + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallHighLamAlef) + ")|(" + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallHighThreeDots) + ")(.*)(" + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallHighThreeDots) + ")|(" + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallHighLigatureQafWithLamWithAlefMaksura) + ")|(" + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallHighLigatureSadWithLamWithAlefMaksura) + ")|(" + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallHighSeen) + ")(?:\s+|\s*$)",
            .Evaluator = New String() {"canstoporcontinue", "prohibitedtostop", "stopatfirstnotsecond", Nothing, "stopatsecondnotfirst", "bettertostopbutpermissibletocontinue", "bettertocontinuebutpermissibletostop", "subtlestopwithoutbreath"}}, _
        New RuleMetadataTranslation With {.Rule = "Stopping", .Match = "(" + ArabicData.MakeUniRegEx(ArabicData.ArabicFathatan) + ")(" + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallLowMeem) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallHighMeemIsolatedForm) + ")?(?=(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlef) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlefMaksura) + ")\s*(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicPlaceOfSajdah) + "?\s*(?:$|" + ArabicData.MakeUniRegEx(ArabicData.ArabicEndOfAyah) + ")|" + CertainStopPattern + "|" + OptionalStopPattern + "))",
            .Evaluator = New String() {"helperfatha", "empty"}}, _
        New RuleMetadataTranslation With {.Rule = "Stopping", .Match = "(?:([^" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterTehMarbuta) + "])(?:(" + ArabicData.MakeRegMultiEx(Array.ConvertAll(ArabicData.ArabicFathaDammaKasra, Function(Str As String) ArabicData.MakeUniRegEx(Str))) + ")|(" + ArabicData.MakeUniRegEx(ArabicData.ArabicKasratan) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicDammatan) + ")(" + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallLowMeem) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallHighMeemIsolatedForm) + ")?)|(" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterHamza) + ")(" + ArabicData.MakeUniRegEx(ArabicData.ArabicFathatan) + ")(" + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallLowMeem) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallHighMeemIsolatedForm) + ")?|(" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterNoon) + ")(" + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallLowMeem) + "))(?=\s*(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicPlaceOfSajdah) + "?\s*(?:$|" + ArabicData.MakeUniRegEx(ArabicData.ArabicEndOfAyah) + ")|" + CertainStopPattern + "|" + OptionalStopPattern + "))",
            .Evaluator = New String() {Nothing, "empty", "empty", "empty", Nothing, "empty", "empty", Nothing, "empty"}}, _
        New RuleMetadataTranslation With {.Rule = "Stopping", .Match = "(" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterTehMarbuta) + ")(?:(" + ArabicData.MakeRegMultiEx(Array.ConvertAll(ArabicData.ArabicFathaDammaKasra, Function(Str As String) ArabicData.MakeUniRegEx(Str))) + ")|(" + ArabicData.MakeRegMultiEx(Array.ConvertAll(ArabicData.ArabicTanweens, Function(Str As String) ArabicData.MakeUniRegEx(Str))) + ")(" + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallLowMeem) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallHighMeemIsolatedForm) + ")?)(?=\s*(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicPlaceOfSajdah) + "?\s*(?:$|" + ArabicData.MakeUniRegEx(ArabicData.ArabicEndOfAyah) + ")|" + CertainStopPattern + "|" + OptionalStopPattern + "))",
            .Evaluator = New String() {"helperheh", "empty", "empty", "empty"}}, _
        New RuleMetadataTranslation With {.Rule = "Continuing", .Match = "(" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterTehMarbuta) + ")(?:(?:" + ArabicData.MakeRegMultiEx(Array.ConvertAll(ArabicData.ArabicFathaDammaKasra, Function(Str As String) ArabicData.MakeUniRegEx(Str))) + ")|(?:" + ArabicData.MakeRegMultiEx(Array.ConvertAll(ArabicData.ArabicTanweens, Function(Str As String) ArabicData.MakeUniRegEx(Str))) + ")(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallLowMeem) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallHighMeemIsolatedForm) + ")?)(?!\s*(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicPlaceOfSajdah) + "?\s*(?:$|" + ArabicData.MakeUniRegEx(ArabicData.ArabicEndOfAyah) + ")|" + CertainStopPattern + "|" + OptionalStopPattern + "))",
            .Evaluator = New String() {"helperteh"}}, _
        New RuleMetadataTranslation With {.Rule = "Continuing", .Match = "(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicFathatan) + ")(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallLowMeem) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallHighMeemIsolatedForm) + ")?(" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlef) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlefMaksura) + ")(?!\s*(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicPlaceOfSajdah) + "?\s*(?:$|" + ArabicData.MakeUniRegEx(ArabicData.ArabicEndOfAyah) + ")|" + CertainStopPattern + "|" + OptionalStopPattern + "))",
            .Evaluator = New String() {"empty"}}, _
        New RuleMetadataTranslation With {.Rule = "SmallBounce", .Match = "(" + ArabicData.MakeRegMultiEx(Array.ConvertAll(ArabicData.ArabicFathaDammaKasra, Function(Str As String) ArabicData.MakeUniRegEx(Str))) + ")(" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterDal) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterBeh) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterQaf) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterTah) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterJeem) + ")(?=" + ArabicData.MakeUniRegEx(ArabicData.ArabicSukun) + "\S)", _
            .Evaluator = New String() {Nothing, "bounce"}}, _
        New RuleMetadataTranslation With {.Rule = "ModerateBounce", .Match = "(" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterDal) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterBeh) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterQaf) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterTah) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterJeem) + ")(?=(?:(?:" + ArabicData.MakeRegMultiEx(Array.ConvertAll(ArabicData.ArabicTanweens, Function(Str As String) ArabicData.MakeUniRegEx(Str))) + ")(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallLowMeem) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallHighMeemIsolatedForm) + ")?|" + ArabicData.MakeRegMultiEx(Array.ConvertAll(ArabicData.ArabicFathaDammaKasra, Function(Str As String) ArabicData.MakeUniRegEx(Str))) + ")\s*(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicPlaceOfSajdah) + "?\s*(?:$|" + ArabicData.MakeUniRegEx(ArabicData.ArabicEndOfAyah) + ")|" + CertainStopPattern + "|" + OptionalStopPattern + "))", _
            .Evaluator = New String() {"bounce"}}, _
        New RuleMetadataTranslation With {.Rule = "GreatBounce", .Match = "(" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterDal) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterBeh) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterQaf) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterTah) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterJeem) + ")(?=" + ArabicData.MakeUniRegEx(ArabicData.ArabicSukun) + "\s*(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicPlaceOfSajdah) + "?\s*(?:$|" + ArabicData.MakeUniRegEx(ArabicData.ArabicEndOfAyah) + ")|" + CertainStopPattern + "|" + OptionalStopPattern + "))", _
            .Evaluator = New String() {"bounce"}}, _
        New RuleMetadataTranslation With {.Rule = "NasalizeCharacterDoubled", .Match = "(" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterNoon) + ")(" + ArabicData.MakeUniRegEx(ArabicData.ArabicShadda) + ")",
            .Evaluator = New String() {"nasalize|normalprolong", "empty"}}, _
        New RuleMetadataTranslation With {.Rule = "VowellessNoonClear", .Match = "(?:(" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterNoon) + ArabicData.MakeUniRegEx(ArabicData.ArabicSukun) + ")|(?:(" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterNoon) + ")|(" + ArabicData.MakeRegMultiEx(Array.ConvertAll(ArabicData.ArabicTanweens, Function(Str As String) ArabicData.MakeUniRegEx(Str))) + "))(?:\s+))(" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAin) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterGhain) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterHah) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterHeh) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterHamza) + ")",
            .Evaluator = Nothing}, _
        New RuleMetadataTranslation With {.Rule = "VowellessNoonCovered", .Match = "(?:(" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterNoon) + ")(" + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallHighMeemIsolatedForm) + ")?(?:\s+)?|(?:(?:(" + ArabicData.MakeUniRegEx(ArabicData.ArabicFathatan) + ")(" + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallHighMeemIsolatedForm) + ")?" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlef) + "|(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterHamza) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterTehMarbuta) + ")(" + ArabicData.MakeUniRegEx(ArabicData.ArabicFathatan) + ")(" + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallHighMeemIsolatedForm) + ")?|(" + ArabicData.MakeUniRegEx(ArabicData.ArabicDammatan) + ")(" + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallHighMeemIsolatedForm) + ")?|(" + ArabicData.MakeUniRegEx(ArabicData.ArabicKasratan) + ")(" + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallLowMeem) + ")?)(?:\s+(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicPlaceOfSajdah) + "\s*$|(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallHighLamAlef) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallHighLigatureSadWithLamWithAlefMaksura) + ")\s+)?)))(?=" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterBeh) + "(" + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicDamma) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicKasra) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicDammatan) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicKasratan) + "))",
            .Evaluator = New String() {"empty", "nasalize|helpermeem", "dividetanween(,empty)", "nasalize|helpermeem", "dividetanween(,empty)", "nasalize|helpermeem", "dividetanween(,empty)", "nasalize|helpermeem", "dividetanween(,empty)", "nasalize|helpermeem"}}, _
        New RuleMetadataTranslation With {.Rule = "VowellessNoonAssimilatingNasalization", .Match = "(?:(?:(" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterNoon) + ")|(" + ArabicData.MakeUniRegEx(ArabicData.ArabicFathatan) + ")(" + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallLowMeem) + ")?(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlef) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlefMaksura) + ")|(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterHamza) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterTehMarbuta) + ")(" + ArabicData.MakeUniRegEx(ArabicData.ArabicFathatan) + ")(" + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallLowMeem) + ")?|(" + ArabicData.MakeUniRegEx(ArabicData.ArabicDammatan) + ")(?:(" + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallLowMeem) + ")(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlef) + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallHighRoundedZero) + ")?)?|(" + ArabicData.MakeUniRegEx(ArabicData.ArabicKasratan) + ")(" + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallHighMeemIsolatedForm) + ")?)(?:\s+(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicPlaceOfSajdah) + "\s*$|(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallHighLamAlef) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallHighLigatureSadWithLamWithAlefMaksura) + ")\s+)?))(" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterNoon) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterWaw) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterMeem) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterYeh) + ")(?=" + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicKasra) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicDamma) + ")",
            .Evaluator = New String() {"assimilate", "dividetanween(,assimilate)", "empty", "dividetanween(,assimilate)", "empty", "dividetanween(,assimilate)", "empty", "dividetanween(,assimilate)", "empty", "nasalize|normalprolong|assimilator"}}, _
        New RuleMetadataTranslation With {.Rule = "VowellessNoonAssimilatingNasalization", .Match = "(?:(?:(" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterNoon) + ")|(" + ArabicData.MakeUniRegEx(ArabicData.ArabicFathatan) + ")(" + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallLowMeem) + ")?(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlef) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlefMaksura) + ")|(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterHamza) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterTehMarbuta) + ")(" + ArabicData.MakeUniRegEx(ArabicData.ArabicFathatan) + ")(" + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallLowMeem) + ")?|(" + ArabicData.MakeUniRegEx(ArabicData.ArabicDammatan) + ")(?:(" + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallLowMeem) + ")(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlef) + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallHighRoundedZero) + ")?)?|(" + ArabicData.MakeUniRegEx(ArabicData.ArabicKasratan) + ")(" + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallHighMeemIsolatedForm) + ")?)(?:\s+(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicPlaceOfSajdah) + "\s*$|(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallHighLamAlef) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallHighLigatureSadWithLamWithAlefMaksura) + ")\s+)?))(" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterNoon) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterWaw) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterMeem) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterYeh) + ")(?=" + ArabicData.MakeUniRegEx(ArabicData.ArabicShadda) + ")",
            .Evaluator = New String() {"assimilate", "dividetanween(,assimilate)", "empty", "dividetanween(,assimilate)", "empty", "dividetanween(,assimilate)", "empty", "dividetanween(,assimilate)", "empty", "nasalize|assimilator"}}, _
        New RuleMetadataTranslation With {.Rule = "VowellessNoonAssimilating", .Match = "(?:(?:(" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterNoon) + ")|(" + ArabicData.MakeUniRegEx(ArabicData.ArabicFathatan) + ")(" + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallLowMeem) + ")?(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlef) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlefMaksura) + ")|(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterHamza) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterTehMarbuta) + ")(" + ArabicData.MakeUniRegEx(ArabicData.ArabicFathatan) + ")(" + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallLowMeem) + ")?|(" + ArabicData.MakeUniRegEx(ArabicData.ArabicDammatan) + ")(?:(" + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallLowMeem) + ")(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlef) + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallHighRoundedZero) + ")?)?|(" + ArabicData.MakeUniRegEx(ArabicData.ArabicKasratan) + ")(" + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallHighMeemIsolatedForm) + ")?)(?:\s+(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicPlaceOfSajdah) + "\s*$|(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallHighLamAlef) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallHighLigatureSadWithLamWithAlefMaksura) + ")\s+)?))(" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterLam) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterReh) + ")(?=" + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicKasra) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicDamma) + ")",
            .Evaluator = New String() {"assimilate", "dividetanween(,assimilate)", "empty", "dividetanween(,assimilate)", "empty", "dividetanween(,assimilate)", "empty", "dividetanween(,assimilate)", "empty", "normalprolong|assimilator"}}, _
        New RuleMetadataTranslation With {.Rule = "VowellessNoonAssimilating", .Match = "(?:(?:(" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterNoon) + ")|(" + ArabicData.MakeUniRegEx(ArabicData.ArabicFathatan) + ")(" + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallLowMeem) + ")?(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlef) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlefMaksura) + ")|(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterHamza) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterTehMarbuta) + ")(" + ArabicData.MakeUniRegEx(ArabicData.ArabicFathatan) + ")(" + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallLowMeem) + ")?|(" + ArabicData.MakeUniRegEx(ArabicData.ArabicDammatan) + ")(?:(" + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallLowMeem) + ")(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlef) + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallHighRoundedZero) + ")?)?|(" + ArabicData.MakeUniRegEx(ArabicData.ArabicKasratan) + ")(" + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallHighMeemIsolatedForm) + ")?)(?:\s+(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicPlaceOfSajdah) + "\s*$|(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallHighLamAlef) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallHighLigatureSadWithLamWithAlefMaksura) + ")\s+)?))(" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterLam) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterReh) + ")(?=" + ArabicData.MakeUniRegEx(ArabicData.ArabicShadda) + ")",
            .Evaluator = New String() {"assimilate", "dividetanween(,assimilate)", "empty", "dividetanween(,assimilate)", "empty", "dividetanween(,assimilate)", "empty", "dividetanween(,assimilate)", "empty", "assimilator"}}, _
        New RuleMetadataTranslation With {.Rule = "VowellessNoonHideHeaviness", .Match = "(?:(?:(" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterNoon) + ")|(" + ArabicData.MakeUniRegEx(ArabicData.ArabicFathatan) + ")(" + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallLowMeem) + ")?(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlef) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlefMaksura) + ")|(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterHamza) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterTehMarbuta) + ")(" + ArabicData.MakeUniRegEx(ArabicData.ArabicFathatan) + ")(" + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallLowMeem) + ")?|(" + ArabicData.MakeUniRegEx(ArabicData.ArabicDammatan) + ")(?:(" + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallLowMeem) + ")(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlef) + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallHighRoundedZero) + ")?)?|(" + ArabicData.MakeUniRegEx(ArabicData.ArabicKasratan) + ")(" + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallHighMeemIsolatedForm) + ")?)(?:\s+(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicPlaceOfSajdah) + "\s*$|(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallHighLamAlef) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallHighLigatureSadWithLamWithAlefMaksura) + ")\s+)?))(?=" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterSad) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterDad) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterTah) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterZah) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterQaf) + ")",
            .Evaluator = New String() {"nasalize|normalprolong", "dividetanween(,normalprolong nasalize)", "empty", "dividetanween(,normalprolong nasalize)", "empty", "dividetanween(,normalprolong nasalize)", "empty", "dividetanween(,normalprolong nasalize)", "empty"}}, _
        New RuleMetadataTranslation With {.Rule = "VowellessNoonHideLightness", .Match = "(?:(?:(" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterNoon) + ")|(" + ArabicData.MakeUniRegEx(ArabicData.ArabicFathatan) + ")(" + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallLowMeem) + ")?(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlef) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlefMaksura) + ")|(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterHamza) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterTehMarbuta) + ")(" + ArabicData.MakeUniRegEx(ArabicData.ArabicFathatan) + ")(" + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallLowMeem) + ")?|(" + ArabicData.MakeUniRegEx(ArabicData.ArabicDammatan) + ")(?:(" + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallLowMeem) + ")(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlef) + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallHighRoundedZero) + ")?)?|(" + ArabicData.MakeUniRegEx(ArabicData.ArabicKasratan) + ")(" + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallHighMeemIsolatedForm) + ")?)(?:\s+(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicPlaceOfSajdah) + "\s*$|(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallHighLamAlef) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallHighLigatureSadWithLamWithAlefMaksura) + ")\s+)?))(?=" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterTeh) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterTheh) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterJeem) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterDal) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterThal) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterZain) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterSeen) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterSheen) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterFeh) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterKaf) + ")",
            .Evaluator = New String() {"nasalize|normalprolong", "dividetanween(,normalprolong nasalize)", "empty", "dividetanween(,normalprolong nasalize)", "empty", "dividetanween(,normalprolong nasalize)", "empty", "dividetanween(,normalprolong nasalize)", "empty"}}, _
        New RuleMetadataTranslation With {.Rule = "CharacterDoubled", .Match = "(" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterMeem) + ")(" + ArabicData.MakeUniRegEx(ArabicData.ArabicShadda) + ")",
            .Evaluator = New String() {"nasalize|normalprolong", "empty"}}, _
        New RuleMetadataTranslation With {.Rule = "VowellessMeemHide", .Match = "(?:(" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterMeem) + ArabicData.MakeUniRegEx(ArabicData.ArabicSukun) + ")|(" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterMeem) + ")(?:\s+))(" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterBeh) + "(?=" + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicDamma) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicKasra) + "))",
            .Evaluator = New String() {"assimilate", "assimilate", "nasalize|normalprolong|assimilator"}}, _
        New RuleMetadataTranslation With {.Rule = "VowellessMeemAssimilating", .Match = "(?:(" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterMeem) + ArabicData.MakeUniRegEx(ArabicData.ArabicSukun) + ")|(" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterMeem) + ")(?:\s+))(" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterMeem) + "(?=" + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicDamma) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicKasra) + "))",
            .Evaluator = New String() {"assimilate", "assimilate", "nasalize|normalprolong|assimilator"}}, _
        New RuleMetadataTranslation With {.Rule = "VowellessMeemAssimilating", .Match = "(?:(" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterMeem) + ArabicData.MakeUniRegEx(ArabicData.ArabicSukun) + ")|(" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterMeem) + ")(?:\s+))(" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterMeem) + "(?=" + ArabicData.MakeUniRegEx(ArabicData.ArabicShadda) + "))",
            .Evaluator = New String() {"assimilate", "assimilate", "nasalize|assimilator"}}, _
        New RuleMetadataTranslation With {.Rule = "VowellessMeemClear", .Match = "(?:(" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterMeem) + ArabicData.MakeUniRegEx(ArabicData.ArabicSukun) + ")|(" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterMeem) + ")(?:\s+))(?=(" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterTeh) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterTheh) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterJeem) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterHah) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterKhah) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterDal) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterThal) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterReh) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterZain) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterSeen) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterSheen) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterSad) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterDad) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterTah) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterZah) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAin) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterGhain) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterQaf) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterKaf) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterLam) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterNoon) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterHamza) + "|((" + ArabicData.MakeUniRegEx(ArabicData.ArabicSukun) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + ")(" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterYeh) + ")))(" + ArabicData.MakeRegMultiEx(Array.ConvertAll(ArabicData.ArabicFathaDammaKasra, Function(Str As String) ArabicData.MakeUniRegEx(Str))) + "))",
            .Evaluator = Nothing}, _
        New RuleMetadataTranslation With {.Rule = "VowellessMeemClearGreater", .Match = "(?:(" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterMeem) + ArabicData.MakeUniRegEx(ArabicData.ArabicSukun) + ")|(" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterMeem) + ")(?:\s+))(?=(" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterFeh) + "|((" + ArabicData.MakeUniRegEx(ArabicData.ArabicSukun) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + ")(" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterWaw) + ")))(" + ArabicData.MakeRegMultiEx(Array.ConvertAll(ArabicData.ArabicFathaDammaKasra, Function(Str As String) ArabicData.MakeUniRegEx(Str))) + "))",
            .Evaluator = Nothing}, _
        New RuleMetadataTranslation With {.Rule = "AlefSmallHighRoundedZero", .Match = "(" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlef) + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallHighRoundedZero) + ")(?=\s*$|\s+|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterYehWithHamzaAbove) + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlefMaksura) + "(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicSukun) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicMaddahAbove) + ")|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterYeh) + ArabicData.MakeUniRegEx(ArabicData.ArabicSukun) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterThal) + ArabicData.MakeUniRegEx(ArabicData.ArabicSukun) + ")",
            .Evaluator = New String() {"empty"}}, _
        New RuleMetadataTranslation With {.Rule = "AlefSmallHighRoundedZero", .Match = ArabicData.MakeUniRegEx(ArabicData.ArabicDamma) + "(" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterWaw) + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallHighRoundedZero) + ")(?=" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterLam) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterReh) + ")",
            .Evaluator = New String() {"empty"}}, _
        New RuleMetadataTranslation With {.Rule = "AlefSmallHighRoundedZero", .Match = ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterYeh) + ArabicData.MakeUniRegEx(ArabicData.ArabicSukun) + "(" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterYeh) + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallHighRoundedZero) + ")|(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlefWithHamzaBelow) + ArabicData.MakeUniRegEx(ArabicData.ArabicKasra) + "((?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterYeh) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlefMaksura) + ")" + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallHighRoundedZero) + "))",
            .Evaluator = New String() {"empty", "empty"}}, _
        New RuleMetadataTranslation With {.Rule = "SmallHighUprightRectangularZero", .Match = "(" + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallHighUprightRectangularZero) + ")",
            .Evaluator = New String() {"empty"}}, _
        New RuleMetadataTranslation With {.Rule = "EmptyCentreLowStop", .Match = "(" + ArabicData.MakeUniRegEx(ArabicData.ArabicEmptyCentreLowStop) + ")(" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlefMaksura) + ")(" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterSuperscriptAlef) + ")",
            .Evaluator = New String() {"empty", "helperfatha", "helperalef"}}, _
        New RuleMetadataTranslation With {.Rule = "EmptyCentreHighStop", .Match = "(" + ArabicData.MakeUniRegEx(ArabicData.ArabicEmptyCentreHighStop) + ")",
            .Evaluator = New String() {"empty"}}, _
        New RuleMetadataTranslation With {.Rule = "RoundedHighStopWithFilledCentre", .Match = "(" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlef) + ")()(" + ArabicData.MakeUniRegEx(ArabicData.ArabicRoundedHighStopWithFilledCentre) + ")",
            .Evaluator = New String() {Nothing, "helperhamza", "helperfatha"}}, _
        New RuleMetadataTranslation With {.Rule = "EmptyShadda", .Match = "((?:^|" + CertainStopPattern + "|" + OptionalStopPattern + ")\s*)(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterBeh) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterTeh) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterDal) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterThal) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterReh) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterTah) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterZah) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAin) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterFeh) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterLam) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterMeem) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterNoon) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterWaw) + ")(" + ArabicData.MakeUniRegEx(ArabicData.ArabicShadda) + ")",
            .Evaluator = New String() {Nothing, "empty"}}, _
        New RuleMetadataTranslation With {.Rule = "EmptyHamza", .Match = "((?:^|" + CertainStopPattern + "|" + OptionalStopPattern + ")\s*)(" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlefWasla) + ")(?=" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterLam) + "(?:(?:" + ArabicData.MakeRegMultiEx(Array.ConvertAll(ArabicData.ArabicSunLetters, Function(Str As String) ArabicData.MakeUniRegEx(Str))) + ")?" + ArabicData.MakeUniRegEx(ArabicData.ArabicShadda) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicSukun) + "?(?:" + ArabicData.MakeRegMultiEx(Array.ConvertAll(ArabicData.ArabicMoonLetters, Function(Str As String) ArabicData.MakeUniRegEx(Str))) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlefWithHamzaAbove) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlefWithHamzaBelow) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterHamza) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicTatweel) + ArabicData.MakeUniRegEx(ArabicData.ArabicHamzaAbove) + ")))",
            .Evaluator = New String() {Nothing, "helperfatha"}}, _
        New RuleMetadataTranslation With {.Rule = "EmptyHamza", .Match = "((?:^|" + CertainStopPattern + "|" + OptionalStopPattern + ")\s*)(" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlefWasla) + ")(?=(?:" + ArabicData.MakeRegMultiEx(Array.ConvertAll(ArabicData.ArabicMoonLetters, Function(Str As String) ArabicData.MakeUniRegEx(Str))) + "|" + ArabicData.MakeRegMultiEx(Array.ConvertAll(ArabicData.ArabicSunLettersNoLam, Function(Str As String) ArabicData.MakeUniRegEx(Str))) + ")(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicSukun) + "|" + ArabicData.MakeRegMultiEx(Array.ConvertAll(ArabicData.ArabicFathaDammaKasra, Function(Str As String) ArabicData.MakeUniRegEx(Str))) + ")?(?:" + ArabicData.MakeRegMultiEx(Array.ConvertAll(ArabicData.ArabicLetters, Function(Str As String) ArabicData.MakeUniRegEx(Str))) + ")(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicKasra) + ")|" + ArabicData.MakeRegMultiEx(Array.ConvertAll(ArabicData.ArabicWaslKasraExceptions, Function(Str As String) ArabicData.MakeUniRegEx(ArabicData.TransliterateFromBuckwalter(Str)))) + ")",
            .Evaluator = New String() {Nothing, "helperkasra"}}, _
        New RuleMetadataTranslation With {.Rule = "EmptyHamza", .Match = "((?:^|" + CertainStopPattern + "|" + OptionalStopPattern + ")\s*)(" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlefWasla) + ")(?=(?:" + ArabicData.MakeRegMultiEx(Array.ConvertAll(ArabicData.ArabicLetters, Function(Str As String) ArabicData.MakeUniRegEx(Str))) + ")(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicSukun) + "|" + ArabicData.MakeRegMultiEx(Array.ConvertAll(ArabicData.ArabicFathaDammaKasra, Function(Str As String) ArabicData.MakeUniRegEx(Str))) + ")?(?:" + ArabicData.MakeRegMultiEx(Array.ConvertAll(ArabicData.ArabicLetters, Function(Str As String) ArabicData.MakeUniRegEx(Str))) + ")" + ArabicData.MakeUniRegEx(ArabicData.ArabicDamma) + ")",
            .Evaluator = New String() {Nothing, "helperdamma"}}, _
        New RuleMetadataTranslation With {.Rule = "EmptyHamza", .Match = "((?:" + CertainNotStopPattern + "|" + OptionalNotStopPattern + ")\s*)(" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlefWasla) + ")",
            .Evaluator = New String() {Nothing, "empty"}}, _
        New RuleMetadataTranslation With {.Rule = "AssimilateLaamSunLetter", .Match = "((?:(?:^|" + CertainStopPattern + "|" + OptionalStopPattern + ")\s*|\s+)(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterWaw) + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterBeh) + ArabicData.MakeUniRegEx(ArabicData.ArabicKasra) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterTeh) + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterKaf) + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterLam) + ArabicData.MakeUniRegEx(ArabicData.ArabicKasra) + ")?" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlefWasla) + ")(" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterLam) + ")(" + ArabicData.MakeRegMultiEx(Array.ConvertAll(ArabicData.ArabicSunLetters, Function(Str As String) ArabicData.MakeUniRegEx(Str))) + ")(?=" + ArabicData.MakeUniRegEx(ArabicData.ArabicShadda) + ")",
            .Evaluator = New String() {Nothing, "assimilate", "assimilator"}}, _
        New RuleMetadataTranslation With {.Rule = "ClearLaamMoonLetter", .Match = "((?:(?:^|" + CertainStopPattern + "|" + OptionalStopPattern + ")\s*|\s+)(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterWaw) + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterBeh) + ArabicData.MakeUniRegEx(ArabicData.ArabicKasra) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterTeh) + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterKaf) + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterLam) + ArabicData.MakeUniRegEx(ArabicData.ArabicKasra) + ")?" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlefWasla) + ")(" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterLam) + ")(?=" + ArabicData.MakeUniRegEx(ArabicData.ArabicSukun) + "?(?:" + ArabicData.MakeRegMultiEx(Array.ConvertAll(ArabicData.ArabicMoonLetters, Function(Str As String) ArabicData.MakeUniRegEx(Str))) + "))",
            .Evaluator = New String() {Nothing, "helperprefix", Nothing}}, _
        New RuleMetadataTranslation With {.Rule = "Sukun", .Match = "(" + ArabicData.MakeUniRegEx(ArabicData.ArabicSukun) + ")", _
            .Evaluator = New String() {"empty"}}, _
        New RuleMetadataTranslation With {.Rule = "AlefMaksuraDaggerAlef", .Match = "(" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlefMaksura) + ")(?=" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterSuperscriptAlef) + ")", _
            .Evaluator = New String() {"helperfatha"}}, _
        New RuleMetadataTranslation With {.Rule = "AlefMaddah", .Match = "(" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlefWithMaddaAbove) + ")", _
            .Evaluator = New String() {"dividelettersymbol(helperalef permissibleprolong,empty)"}}, _
        New RuleMetadataTranslation With {.Rule = "Maddah", .Match = "(" + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlef) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + "(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterWaw) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlefMaksura) + ")?" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterSuperscriptAlef) + ")(" + ArabicData.MakeUniRegEx(ArabicData.ArabicMaddahAbove) + ")", _
            .Evaluator = New String() {"permissibleprolong", "empty"}}, _
        New RuleMetadataTranslation With {.Rule = "Maddah", .Match = "(" + ArabicData.MakeUniRegEx(ArabicData.ArabicDamma) + "(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallWaw) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterWaw) + ")|" + ArabicData.MakeUniRegEx(ArabicData.ArabicKasra) + "(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallYeh) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterYeh) + ")|" + ArabicData.MakeUniRegEx(ArabicData.ArabicKasra) + "(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlef) + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallHighRoundedZero) + ")?" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlefMaksura) + ")(" + ArabicData.MakeUniRegEx(ArabicData.ArabicMaddahAbove) + ")", _
            .Evaluator = New String() {"permissibleprolong", "empty"}}, _
        New RuleMetadataTranslation With {.Rule = "AlefWaslaWawHamza", .Match = "((?:^|" + CertainStopPattern + "|" + OptionalStopPattern + ")\s*" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlefWasla) + ")(" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterWawWithHamzaAbove) + ")", _
            .Evaluator = New String() {Nothing, "dividelettersymbol(helperwaw,empty)"}}, _
        New RuleMetadataTranslation With {.Rule = "AlefWaslaAlefMaksuraHamza", .Match = "((?:^|" + CertainStopPattern + "|" + OptionalStopPattern + ")\s*" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlefWasla) + ")(" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterYehWithHamzaAbove) + ")", _
            .Evaluator = New String() {Nothing, "dividelettersymbol(helperyeh,empty)"}}, _
        New RuleMetadataTranslation With {.Rule = "AlefHamza", .Match = "((?:^|" + CertainStopPattern + "|" + OptionalStopPattern + ")\s*)(" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlefWithHamzaAbove) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlefWithHamzaBelow) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterHamza) + ")", _
            .Evaluator = New String() {Nothing, "empty"}}, _
        New RuleMetadataTranslation With {.Rule = "AlefHamza", .Match = "((?:" + CertainNotStopPattern + "|" + OptionalNotStopPattern + ")\s*)(" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlefWithHamzaAbove) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlefWithHamzaBelow) + ")", _
            .Evaluator = New String() {Nothing, "dividelettersymbol(empty,helperhamza)"}}, _
        New RuleMetadataTranslation With {.Rule = "AlefHamza", .Match = "((?:" + CertainNotStopPattern + "|" + OptionalNotStopPattern + ")\s*)(" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlefWithHamzaAbove) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlefWithHamzaBelow) + ")", _
            .Evaluator = New String() {Nothing, "dividelettersymbol(empty,helperhamza)"}}, _
        New RuleMetadataTranslation With {.Rule = "WawHamzah", .Match = "(" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterWawWithHamzaAbove) + ")", _
            .Evaluator = New String() {"dividelettersymbol(empty,helperhamza)"}}, _
        New RuleMetadataTranslation With {.Rule = "AlefMaksuraHamzah", .Match = "(" + ArabicData.MakeUniRegEx(ArabicData.ArabicKasra) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + "(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlef) + ArabicData.MakeUniRegEx(ArabicData.ArabicMaddahAbove) + "?)?|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterSuperscriptAlef) + ArabicData.MakeUniRegEx(ArabicData.ArabicMaddahAbove) + ")(" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterYehWithHamzaAbove) + ")", _
            .Evaluator = New String() {Nothing, "dividelettersymbol(empty,helperhamza)"}}, _
        New RuleMetadataTranslation With {.Rule = "SmallWaw", .Match = "(" + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallWaw) + ")", _
            .Evaluator = New String() {"helperwaw"}}, _
        New RuleMetadataTranslation With {.Rule = "SmallYeh", .Match = "(" + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallYeh) + ")", _
            .Evaluator = New String() {"helperyeh"}}, _
        New RuleMetadataTranslation With {.Rule = "SmallNoon", .Match = "(" + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallHighNoon) + ")", _
            .Evaluator = New String() {"helpernoon"}}, _
        New RuleMetadataTranslation With {.Rule = "SmallHighSeen", .Match = "(" + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallHighSeen) + ")\S+", _
            .Evaluator = New String() {"helperseen"}}, _
        New RuleMetadataTranslation With {.Rule = "SmallLowSeen", .Match = "(" + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallLowSeen) + ")", _
            .Evaluator = New String() {"empty"}}, _
        New RuleMetadataTranslation With {.Rule = "DaggerAlef", .Match = "(" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterSuperscriptAlef) + ")", _
            .Evaluator = New String() {"helperalef"}}, _
        New RuleMetadataTranslation With {.Rule = "SmallHamza", .Match = "((?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicTatweel) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterSuperscriptAlef) + ")" + ArabicData.MakeUniRegEx(ArabicData.ArabicHamzaAbove) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicHamzaBelow) + ")", _
            .Evaluator = New String() {"helperhamza"}}, _
        New RuleMetadataTranslation With {.Rule = "FathaAlefMaksura", .Match = "(" + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + ")(" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlefMaksura) + ")(?=\s+|\s*$|" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterSuperscriptAlef) + ")", _
            .Evaluator = New String() {Nothing, "helperalef"}}, _
        New RuleMetadataTranslation With {.Rule = "FathaAlefMaksura", .Match = "(" + ArabicData.MakeUniRegEx(ArabicData.ArabicFathatan) + "(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallLowMeem) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallHighMeemIsolatedForm) + ")?)(" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlefMaksura) + ")(?=\s*(?:" + ArabicData.MakeUniRegEx(ArabicData.ArabicPlaceOfSajdah) + "?\s*(?:$|" + ArabicData.MakeUniRegEx(ArabicData.ArabicEndOfAyah) + ")|" + CertainStopPattern + "|" + OptionalStopPattern + "))", _
            .Evaluator = New String() {Nothing, "helperalef"}}, _
        New RuleMetadataTranslation With {.Rule = "FathaAlefMaksura", .Match = "(" + ArabicData.MakeUniRegEx(ArabicData.ArabicFatha) + ")(" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlefMaksura) + ")(?=" + ArabicData.MakeUniRegEx(ArabicData.ArabicSukun) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicKasra) + "|" + ArabicData.MakeUniRegEx(ArabicData.ArabicShadda) + ")", _
            .Evaluator = New String() {Nothing, "helperyeh"}}, _
        New RuleMetadataTranslation With {.Rule = "KasraAlefMaksura", .Match = "(" + ArabicData.MakeUniRegEx(ArabicData.ArabicKasra) + ")(" + ArabicData.MakeUniRegEx(ArabicData.ArabicLetterAlefMaksura) + ")(?!" + ArabicData.MakeUniRegEx(ArabicData.ArabicSmallHighRoundedZero) + ")", _
            .Evaluator = New String() {Nothing, "helperyeh"}}
    }
    Public Shared Function ArabicLetterSpelling(Input As String, Quranic As Boolean) As String
        Dim Output As String = String.Empty
        For Each Ch As Char In Input
            Dim Index As Integer = ArabicData.FindLetterBySymbol(Ch)
            If Index <> -1 AndAlso ArabicData.IsLetter(Index) Then
                If Output <> String.Empty And Not Quranic Then Output += " "
                Output += If(Quranic, ArabicData.Data.ArabicLetters(Index).SymbolName.Remove(ArabicData.Data.ArabicLetters(Index).SymbolName.Length - 1) + If(ArabicData.Data.ArabicLetters(Index).SymbolName.EndsWith("n"), String.Empty, "o"), ArabicData.Data.ArabicLetters(Index).SymbolName)
            ElseIf Index <> -1 AndAlso ArabicData.Data.ArabicLetters(Index).Symbol = ArabicData.ArabicMaddahAbove Then
                If Not Quranic Then Output += Ch
            End If
        Next
        Return ArabicData.TransliterateFromBuckwalter(Output)
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
            HundStr = If(UseAlefHundred, Arabic.ArabicBaseHundredNumbers((Number \ 100) - 1).Insert(2, "A"), Arabic.ArabicBaseHundredNumbers((Number \ 100) - 1))
            If (Number Mod 100) = 0 Then Return HundStr
            Number = Number Mod 100
        End If
        If (Number Mod 10) <> 0 And Number <> 11 And Number <> 12 Then
            Str = Arabic.ArabicBaseNumbers((Number Mod 10) - 1)
        End If
        If Number >= 11 AndAlso Number < 20 Then
            If Number = 11 Or Number = 12 Then
                Str += Arabic.ArabicBaseExtraNumbers(Number - 11)
            Else
                Str = Str.Remove(Str.Length - 1) + "a"
            End If
            Str += " Ea$ara"
        ElseIf (Number = 0 And Str = String.Empty) Or Number = 10 Or Number >= 20 Then
            Str = If(Str = String.Empty, String.Empty, Str + " wa") + Arabic.ArabicBaseTenNumbers(Number \ 10)
        End If
        Return If(UseClassic, If(Str = String.Empty, String.Empty, Str + If(HundStr = String.Empty, String.Empty, " wa")) + HundStr, If(HundStr = String.Empty, String.Empty, HundStr + If(Str = String.Empty, String.Empty, " wa")) + Str)
    End Function
    Public Shared Function ArabicWordFromNumber(ByVal Number As Long, UseClassic As Boolean, UseAlefHundred As Boolean, UseMilliard As Boolean) As String
        Dim Str As String = String.Empty
        Dim NextStr As String = String.Empty
        Dim CurBase As Integer = 3
        Dim BaseNums As Long() = {1000, 1000000, 1000000000, 1000000000000}
        Dim Bases As String()() = {ArabicBaseThousandNumbers, Arabic.ArabicBaseMillionNumbers, If(UseMilliard, Arabic.ArabicBaseMilliardNumbers, Arabic.ArabicBaseBillionNumbers), Arabic.ArabicBaseTrillionNumbers}
        Do
            If Number >= BaseNums(CurBase) And Number < 2 * BaseNums(CurBase) Then
                NextStr = Bases(CurBase)(0)
            ElseIf Number >= 2 * BaseNums(CurBase) And Number < 3 * BaseNums(CurBase) Then
                NextStr = Bases(CurBase)(1)
            ElseIf Number >= 3 * BaseNums(CurBase) And Number < 10 * BaseNums(CurBase) Then
                NextStr = Arabic.ArabicBaseNumbers(CInt(Number \ BaseNums(CurBase) - 1)).Remove(Arabic.ArabicBaseNumbers(CInt(Number \ BaseNums(CurBase) - 1)).Length - 1) + "u " + Bases(CurBase)(2)
            ElseIf Number >= 10 * BaseNums(CurBase) And Number < 11 * BaseNums(CurBase) Then
                NextStr = Arabic.ArabicBaseTenNumbers(1).Remove(Arabic.ArabicBaseTenNumbers(1).Length - 1) + "u " + Bases(CurBase)(2)
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
            Str = If(UseClassic, If(NextStr = String.Empty, String.Empty, NextStr + If(Str = String.Empty, String.Empty, " wa")) + Str, If(Str = String.Empty, String.Empty, Str + If(NextStr = String.Empty, String.Empty, " wa")) + NextStr)
            NextStr = String.Empty
        Loop While CurBase >= 0
        If Number <> 0 Or Str = String.Empty Then NextStr = Arabic.ArabicWordForLessThanThousand(CInt(Number), UseClassic, UseAlefHundred)
        Return If(UseClassic, If(NextStr = String.Empty, String.Empty, NextStr + If(Str = String.Empty, String.Empty, " wa")) + Str, If(Str = String.Empty, String.Empty, Str + If(NextStr = String.Empty, String.Empty, " wa")) + NextStr)
    End Function
    Public Shared Function NegativeMatchEliminator(NegativeMatch As String, Evaluator As String) As System.Text.RegularExpressions.MatchEvaluator
        Return Function(Match As System.Text.RegularExpressions.Match)
                   Return If(NegativeMatch <> String.Empty AndAlso Match.Result(NegativeMatch) <> String.Empty, Match.Value, Match.Result(Evaluator))
               End Function
    End Function
    Public Shared Function ProcessTransform(ArabicString As String, Rules As RuleTranslation()) As String
        For Count = 0 To Rules.Length - 1
            ArabicString = System.Text.RegularExpressions.Regex.Replace(ArabicString, Rules(Count).Match, NegativeMatchEliminator(Rules(Count).NegativeMatch, Rules(Count).Evaluator))
        Next
        Return ArabicString
    End Function
    Public Shared Function ChangeBaseScript(ArabicString As String, BaseText As TanzilReader.QuranTexts, ByVal PreString As String, ByVal PostString As String) As String
        If BaseText = TanzilReader.QuranTexts.Warsh Then
            ArabicString = UnjoinContig(ProcessTransform(JoinContig(ArabicString, PreString, PostString), WarshScript), PreString, PostString)
        End If
        Return ArabicString
    End Function
    Public Shared Function ChangeScript(ArabicString As String, ScriptType As TanzilReader.QuranScripts, ByVal PreString As String, ByVal PostString As String) As String
        If ScriptType = TanzilReader.QuranScripts.UthmaniMin Then
            ArabicString = ProcessTransform(ArabicString, UthmaniMinimalScript)
        ElseIf ScriptType = TanzilReader.QuranScripts.SimpleEnhanced Then
            ArabicString = ProcessTransform(ArabicString, SimpleEnhancedScript)
        ElseIf ScriptType = TanzilReader.QuranScripts.Simple Then
            ArabicString = ProcessTransform(ArabicString, SimpleEnhancedScript)
            ArabicString = ProcessTransform(ArabicString, SimpleScript)
        ElseIf ScriptType = TanzilReader.QuranScripts.SimpleClean Then
            ArabicString = ProcessTransform(ArabicString, SimpleEnhancedScript)
            ArabicString = ProcessTransform(ArabicString, SimpleCleanScript)
        ElseIf ScriptType = TanzilReader.QuranScripts.SimpleMin Then
            ArabicString = ProcessTransform(ArabicString, SimpleEnhancedScript)
            ArabicString = ProcessTransform(ArabicString, SimpleMinimalScript)
        End If
        Return ArabicString
    End Function
    Public Shared Function ReplaceMetadata(ArabicString As String, MetadataRule As RuleMetadata, Scheme As String, OptionalStops As Boolean()) As String
        For Count As Integer = 0 To ColoringSpelledOutRules.Length - 1
            Dim Match As String = Array.Find(ColoringSpelledOutRules(Count).Match.Split("|"c), Function(Str As String) Array.IndexOf(Array.ConvertAll(MetadataRule.Type.Split("|"c), Function(S As String) System.Text.RegularExpressions.Regex.Replace(S, "\(.*\)", String.Empty)), Str) <> -1)
            If Match <> Nothing Then
                Dim Str As String = String.Format(ColoringSpelledOutRules(Count).Evaluator, ArabicString.Substring(MetadataRule.Index, MetadataRule.Length))
                If ColoringSpelledOutRules(Count).RuleFunc <> RuleFuncs.eNone Then
                    Dim Args As String() = RuleFunctions(ColoringSpelledOutRules(Count).RuleFunc - 1)(Str, Scheme)
                    If Args.Length = 1 Then
                        Str = Args(0)
                    Else
                        Dim MetaArgs As String() = System.Text.RegularExpressions.Regex.Match(MetadataRule.Type, Match + "\((.*)\)").Groups(1).Value.Split(","c)
                        Str = String.Empty
                        For Index As Integer = 0 To Args.Length - 1
                            If Not Args(Index) Is Nothing Then
                                Str += ReplaceMetadata(Args(Index), New RuleMetadata(0, Args(Index).Length, MetaArgs(Index).Replace(" "c, "|"c)), Scheme, OptionalStops)
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
        For Count = 0 To ErrorCheckRules.Length - 1
            Dim Matches As System.Text.RegularExpressions.MatchCollection = System.Text.RegularExpressions.Regex.Matches(ArabicString, ErrorCheckRules(Count).Match)
            For MatchIndex As Integer = 0 To Matches.Count - 1
                If ErrorCheckRules(Count).NegativeMatch Is Nothing OrElse Matches(MatchIndex).Result(ErrorCheckRules(Count).NegativeMatch) = String.Empty Then
                    Debug.Print(ErrorCheckRules(Count).Rule + ": " + TransliterateToScheme(ArabicString, ArabicData.TranslitScheme.Literal, String.Empty).Insert(Matches(MatchIndex).Index, "<!-- -->"))
                End If
            Next
        Next
    End Sub
    Public Shared Function JoinContig(ByVal ArabicString As String, ByVal PreString As String, ByVal PostString As String) As String
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
    Public Shared Function TransliterateContigWithRules(ByVal ArabicString As String, ByVal PreString As String, ByVal PostString As String, Scheme As String, OptionalStops As Boolean(), PreOptionalStops As Boolean(), PostOptionalStops As Boolean()) As String
        Return UnjoinContig(TransliterateWithRules(JoinContig(ArabicString, PreString, PostString), Scheme, Nothing), PreString, PostString)
    End Function
    Public Shared Function TransliterateWithRules(ByVal ArabicString As String, Scheme As String, OptionalStops As Boolean()) As String
        Dim Count As Integer
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
            ArabicString = ReplaceMetadata(ArabicString, MetadataList(Index), Scheme, OptionalStops)
        Next
        'redundant romanization rules should have -'s such as seen/teh/kaf-heh
        For Count = 0 To RomanizationRules.Length - 1
            If RomanizationRules(Count).RuleFunc = RuleFuncs.eNone Then
                ArabicString = System.Text.RegularExpressions.Regex.Replace(ArabicString, RomanizationRules(Count).Match, RomanizationRules(Count).Evaluator)
            Else
                ArabicString = System.Text.RegularExpressions.Regex.Replace(ArabicString, RomanizationRules(Count).Match, Function(Match As System.Text.RegularExpressions.Match) RuleFunctions(RomanizationRules(Count).RuleFunc - 1)(Match.Result(RomanizationRules(Count).Evaluator), Scheme)(0))
            End If
        Next

        'process wasl loanwords and names
        'process loanwords and names
        Return ArabicString
    End Function
    Shared Function GetTranslitSchemeJSArray() As String
        'Dim Letters(ArabicData.Data.ArabicLetters.Length - 1) As IslamData.ArabicSymbol
        'ArabicData.Data.ArabicLetters.CopyTo(Letters, 0)
        'Array.Sort(Letters, New StringLengthComparer("RomanTranslit"))
        Return "var translitSchemes = " + Utility.MakeJSIndexedObject(Array.ConvertAll(CachedData.IslamData.TranslitSchemes, Function(TranslitScheme As IslamData.TranslitScheme) CStr(Array.IndexOf(CachedData.IslamData.TranslitSchemes, TranslitScheme) + 2)), _
                                                                          New Array() {Array.ConvertAll(Of IslamData.TranslitScheme, String)(CachedData.IslamData.TranslitSchemes, Function(TranslitScheme As IslamData.TranslitScheme) Utility.MakeJSIndexedObject({"standard", "gutteral"}, New Array() {New String() {Utility.MakeJSIndexedObject(ArabicData.ArabicTranslitLetters(), New Array() {Array.ConvertAll(Of String, String)(ArabicData.ArabicTranslitLetters(), Function(Str As String) If(Str.Length = 2, GetSchemeLongVowelFromString(Str, TranslitScheme.Name), GetSchemeValueFromSymbol(ArabicData.Data.ArabicLetters(ArabicData.FindLetterBySymbol(Str.Chars(0))), TranslitScheme.Name)))}, False), Utility.MakeJSArray(TranslitScheme.Vowels)}}, True))}, True)

    End Function
    Shared Function GetArabicSymbolJSArray() As String
        GetArabicSymbolJSArray = "var arabicLetters = " + _
                                Utility.MakeJSArray(New String() {Utility.MakeJSIndexedObject(New String() {"Symbol", "Shaping", "Assimilate", "TranslitLetter"}, _
                                Array.ConvertAll(Of ArabicData.ArabicXMLData.ArabicSymbol, String())(ArabicData.Data.ArabicLetters, Function(Convert As ArabicData.ArabicXMLData.ArabicSymbol) New String() {CStr(AscW(Convert.Symbol)), If(Convert.Shaping = Nothing, String.Empty, Utility.MakeJSArray(Array.ConvertAll(Convert.Shaping, Function(Ch As Char) CStr(AscW(Ch))))), CStr(IIf(Convert.Assimilate, "true", String.Empty)), CStr(IIf(Convert.ExtendedBuckwalterLetter = ChrW(0), String.Empty, Convert.ExtendedBuckwalterLetter))}), False)}, True) + ";"
    End Function
    Public Shared FindLetterBySymbolJS As String = "function findLetterBySymbol(chVal) { var iSubCount; for (iSubCount = 0; iSubCount < arabicLetters.length; iSubCount++) { if (chVal === parseInt(arabicLetters[iSubCount].Symbol, 10)) return iSubCount; for (var iShapeCount = 0; iShapeCount < arabicLetters[iSubCount].Shaping.length; iShapeCount++) { if (chVal === parseInt(arabicLetters[iSubCount].Shaping[iShapeCount], 10)) return iSubCount; } } return -1; }"
    Public Shared TransliterateGenJS As String() = {
        FindLetterBySymbolJS,
        "function isLetterDiacritic(index) { return (" + String.Join("||", Array.ConvertAll(ArabicData.RecitationLettersDiacritics, Function(C As Char) "parseInt(arabicLetters[index].Symbol, 10) === 0x" + Hex(AscW(C)))) + "); }", _
        "function isSpecialSymbol(index) { return (" + String.Join("||", Array.ConvertAll(ArabicData.RecitationSpecialSymbols, Function(C As Char) "parseInt(arabicLetters[index].Symbol, 10) === 0x" + Hex(AscW(C)))) + "); }", _
        "function isCombiningSymbol(index) { return (" + String.Join("||", Array.ConvertAll(ArabicData.RecitationCombiningSymbols, Function(C As Char) "parseInt(arabicLetters[index].Symbol, 10) === 0x" + Hex(AscW(C)))) + "); }", _
        "function doTransliterate(sVal, direction, conversion) { var iCount, iSubCount, sOutVal = ''; if (conversion === 0) return sVal; if (direction && (conversion % 2) === 0) return transliterateWithRules(sVal, Math.floor((conversion - 2) / 2) + 2); for (iCount = 0; iCount < sVal.length; iCount++) { if (sVal.charAt(iCount) === '\\') { iCount++; if (sVal.charAt(iCount) === ',') { sOutVal += String.fromCharCode(1548); } else { sOutVal += sVal.charAt(iCount); } } else { for (iSubCount = 0; iSubCount < arabicLetters.length; iSubCount++) { if (direction ? sVal.charCodeAt(iCount) === parseInt(arabicLetters[iSubCount].Symbol, 10) : sVal.charAt(iCount) === unescape((conversion === 1 ? arabicLetters[iSubCount].TranslitLetter : translitSchemes[(Math.floor((conversion - 2) / 2) + 2).toString()].standard[String.fromCharCode(arabicLetters[iSubCount].Symbol)]))) { sOutVal += (direction ? (conversion === 1 ? arabicLetters[iSubCount].TranslitLetter : translitSchemes[(Math.floor((conversion - 2) / 2) + 2).toString()].standard[String.fromCharCode(arabicLetters[iSubCount].Symbol)]) : ((isCombiningSymbol(iSubCount) && (iSubCount === 0 || findLetterBySymbol(sOutVal.charCodeAt(sOutVal.length - 1)) === -1 || !isLetterDiacritic(findLetterBySymbol(sOutVal.charCodeAt(sOutVal.length - 1))) && !isSpecialSymbol(findLetterBySymbol(sOutVal.charCodeAt(sOutVal.length - 1))))) ? String.fromCharCode(0x202D) + String.fromCharCode(0x25CC) : '') + String.fromCharCode(arabicLetters[iSubCount].Symbol) + ((isCombiningSymbol(iSubCount) && (iSubCount === 0 || findLetterBySymbol(sOutVal.charCodeAt(sOutVal.length - 1)) === -1 || !isLetterDiacritic(findLetterBySymbol(sOutVal.charCodeAt(sOutVal.length - 1))) && !isSpecialSymbol(findLetterBySymbol(sOutVal.charCodeAt(sOutVal.length - 1))))) ? String.fromCharCode(0x202C) : '')); break; } } if (iSubCount === arabicLetters.length) sOutVal += sVal.charAt(iCount); } } return unescape(sOutVal); }"
    }
    Public Shared IsDiacriticJS As String = "function isDiacritic(index) { return (" + String.Join("||", Array.ConvertAll(ArabicData.RecitationDiacritics, Function(C As Char) "parseInt(arabicLetters[index].Symbol, 10) === 0x" + Hex(AscW(C)))) + "); }"
    Public Shared DiacriticJS As String() =
        {"function doDiacritics(sVal, direction) { var iCount, sOutVal = ''; for (iCount = 0; iCount < sVal.length; iCount++) { sOutVal += findLetterBySymbol(sVal.charCodeAt(iCount)) === -1 || !isDiacritic(findLetterBySymbol(sVal.charCodeAt(iCount))) ? sVal[iCount] : ''; } return sOutVal; }", _
            IsDiacriticJS, _
            FindLetterBySymbolJS}
    Public Shared PlainTransliterateGenJS As String() = {FindLetterBySymbolJS, IsDiacriticJS, _
            "function isWhitespace(index) { return (" + String.Join("||", Array.ConvertAll(ArabicData.WhitespaceSymbols, Function(C As Char) "parseInt(arabicLetters[index].Symbol, 10) === 0x" + Hex(AscW(C)))) + "); }", _
            "function isPunctuation(index) { return (" + String.Join("||", Array.ConvertAll(ArabicData.PunctuationSymbols, Function(C As Char) "parseInt(arabicLetters[index].Symbol, 10) === 0x" + Hex(AscW(C)))) + "); }", _
            "function isStop(index) { return (" + String.Join("||", Array.ConvertAll(ArabicData.ArabicStopLetters, Function(C As Char) "parseInt(arabicLetters[index].Symbol, 10) === 0x" + Hex(AscW(C)))) + "); }", _
            "function applyColorRules(sVal) {}", _
            "function changeScript(sVal, scriptType) {}", _
            "var arabicLeadingGutterals = " + Utility.MakeJSArray(ArabicData.ArabicLeadingGutterals) + ";", _
            "function getSchemeGutteralFromString(str, scheme, leading) { if (arabicLeadingGutterals.indexOf(str) !== -1) { return translitSchemes[scheme].gutteral[arabicLeadingGutterals.indexOf(str) + " + CStr(ArabicData.ArabicVowels.Length) + " + (leading ? arabicLeadingGutterals.length : 0)]; } return ''; }", _
            "function arabicLetterSpelling(sVal) { var count, index, output = ''; for (count = 0; count < sVal.length; count++) { index = findLetterBySymbol(sVal.charCodeAt(count)); if (index !== -1 && isLetter(index)) { if (output !== '') output += ' '; output += arabicLetters[index].SymbolName; } else if (index !== -1 && arabicLetters[index].Symbol === 1619) { output += sVal.charCodeAt(count); } } return doTransliterate(output, false, 1); }", _
            "String.prototype.format = function() { var formatted = this; for (var i = 0; i < arguments.length; i++) { formatted = formatted.replace(new RegExp('\\{'+i+'\\}', 'gi'), arguments[i]); } return formatted; };", _
            "RegExp.matchResult = function(subexp, match, offset, str) { var args = arguments; return subexp.replace(/\$(\$|&|`|\'|[0-9]+)/g, function(m, p) { if (p === '$') return '$'; if (p === '`') return str.slice(0, offset); if (p === '\'') return str.slice(offset + match.length); if (p === '&' || parseInt(p, 10) <= 0 || parseInt(p, 10) >= args.length - 3) return match; return args[3 + parseInt(p, 10)]; }); };", _
            "var ruleFunctions = [function(str, scheme) { return [str.toUpperCase()]; }, function(str, scheme) { return [transliterateWithRules(doTransliterate(arabicWordFromNumber(parseInt(doTransliterate(str, true, 1), 10), true, false, false), false, 1), scheme)]; }, function(str, scheme) { return [arabicLetterSpelling(str)]; }, function(str, scheme) { return [translitSchemes[scheme.toString()].standard[str]]; }, function(str, scheme) { return [translitSchemes[scheme.toString()].standard[str]]; }, function(str, scheme) { return [" + Utility.MakeJSArray(ArabicData.ArabicFathaDammaKasra) + "[" + Utility.MakeJSArray(ArabicData.ArabicTanweens) + ".indexOf(str)], '" + ArabicData.ArabicLetterNoon + "']; }, function (str, scheme) { return ['', '']; }, function (str, scheme) { return ['']; }, function (str, scheme) { return [getSchemeGutteralFromString(str.slice(0, -1), scheme, true) + str[str.length - 1]]; }, function(str, scheme) { return [str[0] + getSchemeGutteralFromString(str.slice(1), scheme, false)]; }];", _
            "function isLetter(index) { return (" + String.Join("||", Array.ConvertAll(ArabicData.RecitationLetters, Function(C As Char) "parseInt(arabicLetters[index].Symbol, 10) === 0x" + Hex(AscW(C)))) + "); }", _
            "function isSymbol(index) { return (" + String.Join("||", Array.ConvertAll(ArabicData.GetRecitationSymbols(), Function(A As Array) "parseInt(arabicLetters[index].Symbol, 10) === 0x" + Hex(AscW(ArabicData.Data.ArabicLetters(CInt(A.GetValue(1))).Symbol)))) + "); }", _
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
            "function replaceMetadata(sVal, metadataRule, scheme) { var count, elimParen = function(s) { return s.replace(/\(.*\)/, ''); }; for (count = 0; count < coloringSpelledOutRules.length; count++) { var index, match = null; for (index = 0; index < coloringSpelledOutRules[count].match.split('|').length; index++) { if (metadataRule.type.split('|').map(elimParen).indexOf(coloringSpelledOutRules[count].match.split('|')[index]) !== -1) { match = coloringSpelledOutRules[count].match.split('|')[index]; break; } } if (match !== null) { var str = coloringSpelledOutRules[count].evaluator.format(sVal.substr(metadataRule.index, metadataRule.length)); if (coloringSpelledOutRules[count].ruleFunc !== 0) { var args = ruleFunctions[coloringSpelledOutRules[count].ruleFunc - 1](str, scheme); if (args.length === 1) { str = args[0]; } else { var metaArgs = metadataRule.type.match(/\((.*)\)/)[1].split(','); str = ''; for (index = 0; index < args.length; index++) { if (args[index]) str += replaceMetadata(args[index], {index: 0, length: args[index].length, type: metaArgs[index].replace(' ', '|')}, scheme); } } } sVal = sVal.substr(0, metadataRule.index) + str + sVal.substr(metadataRule.index + metadataRule.length); } } return sVal; }", _
            "function transliterateWithRules(sVal, scheme) { var count, index, arr, re, metadataList = [], replaceFunc = function(f, e) { return function(m) { return f(RegExp.matchResult(e, m, arguments[arguments.length - 2], arguments[arguments.length - 1]), scheme)[0]; }; }; sVal = sVal.replace(/\u200F/g, ''); for (count = 0; count < errorCheckRules.length; count++) { re = new RegExp(errorCheckRules[count].match, 'g'); while ((arr = re.exec(sVal)) !== null) { console.log(errorCheckRules[count].rule + ': ' + doTransliterate(sVal.substr(0, arr.index), true, 1) + '<!-- -->' + doTransliterate(sVal.substr(arr.index), true, 1)); } } for (count = 0; count < rulesOfRecitationRegEx.length; count++) { if (rulesOfRecitationRegEx[count].evaluator !== null) { var subcount, lindex; re = new RegExp(rulesOfRecitationRegEx[count].match, 'g'); while ((arr = re.exec(sVal)) !== null) { lindex = arr.index; for (subcount = 0; subcount < rulesOfRecitationRegEx[count].evaluator.length; subcount++) { if (rulesOfRecitationRegEx[count].evaluator[subcount] !== null && (arr[subcount + 1] && arr[subcount + 1].length !== 0 || allowZeroLength.indexOf(rulesOfRecitationRegEx[count].evaluator[subcount]) !== -1)) { metadataList.push({index: lindex, length: arr[subcount + 1] ? arr[subcount + 1].length : 0, type: rulesOfRecitationRegEx[count].evaluator[subcount]}); } lindex += (arr[subcount + 1] ? arr[subcount + 1].length : 0); } } } } metadataList.sort(ruleMetadataComparer); for (index = 0; index < metadataList.length; index++) { sVal = replaceMetadata(sVal, metadataList[index], scheme); } for (count = 0; count < romanizationRules.length; count++) { sVal = sVal.replace(new RegExp(romanizationRules[count].match, 'g'), (romanizationRules[count].ruleFunc === 0) ? romanizationRules[count].evaluator : replaceFunc(ruleFunctions[romanizationRules[count].ruleFunc - 1], romanizationRules[count].evaluator)); } return sVal; }"}
    Public Shared NumberGenJS As String() = {"var arabicOrdinalNumbers = " + Utility.MakeJSArray(Arabic.ArabicOrdinalNumbers) + ";", _
                "var arabicOrdinalExtraNumbers = " + Utility.MakeJSArray(Arabic.ArabicOrdinalExtraNumbers) + ";", _
                "var arabicFractionNumbers = " + Utility.MakeJSArray(Arabic.ArabicFractionNumbers) + ";", _
                "var arabicBaseNumbers = " + Utility.MakeJSArray(Arabic.ArabicBaseNumbers) + ";", _
                "var arabicBaseExtraNumbers = " + Utility.MakeJSArray(Arabic.ArabicBaseExtraNumbers) + ";", _
                "var arabicBaseTenNumbers = " + Utility.MakeJSArray(Arabic.ArabicBaseTenNumbers) + ";", _
                "var arabicBaseHundredNumbers = " + Utility.MakeJSArray(Arabic.ArabicBaseHundredNumbers) + ";", _
                "var arabicBaseThousandNumbers = " + Utility.MakeJSArray(Arabic.ArabicBaseThousandNumbers) + ";", _
                "var arabicBaseMillionNumbers = " + Utility.MakeJSArray(Arabic.ArabicBaseMillionNumbers) + ";", _
                "var arabicBaseBillionNumbers = " + Utility.MakeJSArray(Arabic.ArabicBaseBillionNumbers) + ";", _
                "var arabicBaseMilliardNumbers = " + Utility.MakeJSArray(Arabic.ArabicBaseMilliardNumbers) + ";", _
                "var arabicBaseTrillionNumbers = " + Utility.MakeJSArray(Arabic.ArabicBaseTrillionNumbers) + ";", _
                "function doTransliterateNum() { $('#translitvalue').text(doTransliterate(arabicWordFromNumber($('#translitedit').val(), $('#useclassic0').prop('checked'), $('#usehundredform0').prop('checked'), $('#usemilliard0').prop('checked')), false, 1)); }", _
                "function arabicWordForLessThanThousand(number, useclassic, usealefhundred) { var str = '', hundstr = ''; if (number >= 100) { hundstr = usealefhundred ? arabicBaseHundredNumbers[Math.floor(number / 100) - 1].substr(0, 2) + 'A' + arabicBaseHundredNumbers[Math.floor(number / 100) - 1].substr(2) : arabicBaseHundredNumbers[Math.floor(number / 100) - 1]; if ((number % 100) === 0) { return hundstr; } number = number % 100; } if ((number % 10) !== 0 && number !== 11 && number !== 12) { str = arabicBaseNumbers[number % 10 - 1]; } if (number >= 11 && number < 20) { if (number == 11 || number == 12) { str += arabicBaseExtraNumbers[number - 11]; } else { str = str.slice(0, -1) + 'a'; } str += ' Ea$ara'; } else if ((number === 0 && str === '') || number === 10 || number >= 20) { str = ((str === '') ? '' : str + ' wa') + arabicBaseTenNumbers[Math.floor(number / 10)]; } return useclassic ? (((str === '') ? '' : str + ((hundstr === '') ? '' : ' wa')) + hundstr) : (((hundstr === '') ? '' : hundstr + ((str === '') ? '' : ' wa')) + str); }", _
                "function arabicWordFromNumber(number, useclassic, usealefhundred, usemilliard) { var str = '', nextstr = '', curbase = 3, basenums = [1000, 1000000, 1000000000, 1000000000000], bases = [arabicBaseThousandNumbers, arabicBaseMillionNumbers, usemilliard ? arabicBaseMilliardNumbers : arabicBaseBillionNumbers, arabicBaseTrillionNumbers]; do { if (number >= basenums[curbase] && number < 2 * basenums[curbase]) { nextstr = bases[curbase][0]; } else if (number >= 2 * basenums[curbase] && number < 3 * basenums[curbase]) { nextstr = bases[curbase][1]; } else if (number >= 3 * basenums[curbase] && number < 10 * basenums[curbase]) { nextstr = arabicBaseNumbers[Math.floor(Number / basenums[curbase]) - 1].slice(0, -1) + 'u ' + bases[curbase][2]; } else if (number >= 10 * basenums[curbase] && number < 11 * basenums[curbase]) { nextstr = arabicBaseTenNumbers[1].slice(0, -1) + 'u ' + bases[curbase][2]; } else if (number >= basenums[curbase]) { nextstr = arabicWordForLessThanThousand(Math.floor(number / basenums[curbase]) % 100, useclassic, usealefhundred); if (number >= 100 * basenums[curbase] && number < (useclassic ? 200 : 101) * basenums[curbase]) { nextstr = nextstr.slice(0, -1) + 'u ' + bases[curbase][0].slice(0, -1) + 'K'; } else if (number >= 200 * basenums[curbase] && number < (useclassic ? 300 : 201) * basenums[curbase]) { nextstr = nextstr.slice(0, -2) + ' ' + bases[curbase][0].slice(0, -1) + 'K'; } else if (number >= 300 * basenums[curbase] && (useclassic || Math.floor(number / basenums[curbase]) % 100 === 0)) { nextstr = nextstr.slice(0, -1) + 'i ' + bases[curbase][0].slice(0, -1) + 'K'; } else { nextstr += ' ' + bases[curbase][0].slice(0, -1) + 'FA'; } } number = number % basenums[curbase]; curbase--; str = useclassic ? (((nextstr === '') ? '' : nextstr + ((str === '') ? '' : ' wa')) + str) : (((str === '') ? '' : str + ((nextstr === '') ? '' : ' wa')) + nextstr); nextstr = ''; } while (curbase >= 0); if (number !== 0 || str === '') { nextstr = arabicWordForLessThanThousand(number, useclassic, usealefhundred); } return useclassic ? (((nextstr === '') ? '' : nextstr + ((str === '') ? '' : ' wa')) + str) : (((str === '') ? '' : str + ((nextstr === '') ? '' : ' wa')) + nextstr); }"}
    Public Shared Function GetTransliterateNumberJS() As String()
        Dim GetJS As New List(Of String) From {"javascript: doTransliterateNum();", String.Empty, GetArabicSymbolJSArray(), GetTranslitSchemeJSArray()}
        GetJS.AddRange(PlainTransliterateGenJS)
        GetJS.AddRange(TransliterateGenJS)
        GetJS.AddRange(NumberGenJS)
        Return GetJS.ToArray()
    End Function
    Public Shared Function GetTransliterateJS() As String()
        Dim GetJS As New List(Of String) From {"javascript: doTransliterateDisplay();", String.Empty, GetArabicSymbolJSArray(), GetTranslitSchemeJSArray(), _
        "function IsLTR(c) { return (c>=0x41&&c<=0x5A)||(c>=0x61&&c<=0x7A)||c===0xAA||c===0xB5||c===0xBA||(c>=0xC0&&c<=0xD6)||(c>=0xD8&&c<=0xF6)||(c>=0xF8&&c<=0x2B8)||(c>=0x2BB&&c<=0x2C1)||(c>=0x2D0&&c<=0x2D1)||(c>=0x2E0&&c<=0x2E4)||c===0x2EE||(c>=0x370&&c<=0x373)||(c>=0x376&&c<=0x377)||(c>=0x37A&&c<=0x37D)||c===0x37F||c===0x386||(c>=0x388&&c<=0x38A)||c===0x38C||(c>=0x38E&&c<=0x3A1)||(c>=0x3A3&&c<=0x3F5)||(c>=0x3F7&&c<=0x482)||(c>=0x48A&&c<=0x52F)||(c>=0x531&&c<=0x556)||(c>=0x559&&c<=0x55F)||(c>=0x561&&c<=0x587)||c===0x589||(c>=0x903&&c<=0x939)||c===0x93B||(c>=0x93D&&c<=0x940)||(c>=0x949&&c<=0x94C)||(c>=0x94E&&c<=0x950)||(c>=0x958&&c<=0x961)||(c>=0x964&&c<=0x980)||(c>=0x982&&c<=0x983)||(c>=0x985&&c<=0x98C)||(c>=0x98F&&c<=0x990)||(c>=0x993&&c<=0x9A8)||(c>=0x9AA&&c<=0x9B0)||c===0x9B2||(c>=0x9B6&&c<=0x9B9)||(c>=0x9BD&&c<=0x9C0)||(c>=0x9C7&&c<=0x9C8)||(c>=0x9CB&&c<=0x9CC)||c===0x9CE||c===0x9D7||(c>=0x9DC&&c<=0x9DD)||(c>=0x9DF&&c<=0x9E1)||(c>=0x9E6&&c<=0x9F1)||(c>=0x9F4&&c<=0x9FA)||c===0xA03||(c>=0xA05&&c<=0xA0A)||(c>=0xA0F&&c<=0xA10)||(c>=0xA13&&c<=0xA28)||(c>=0xA2A&&c<=0xA30)||(c>=0xA32&&c<=0xA33)||(c>=0xA35&&c<=0xA36)||(c>=0xA38&&c<=0xA39)||(c>=0xA3E&&c<=0xA40)||(c>=0xA59&&c<=0xA5C)||c===0xA5E||(c>=0xA66&&c<=0xA6F)||(c>=0xA72&&c<=0xA74)||c===0xA83||(c>=0xA85&&c<=0xA8D)||(c>=0xA8F&&c<=0xA91)||(c>=0xA93&&c<=0xAA8)||(c>=0xAAA&&c<=0xAB0)||(c>=0xAB2&&c<=0xAB3)||(c>=0xAB5&&c<=0xAB9)||(c>=0xABD&&c<=0xAC0)||c===0xAC9||(c>=0xACB&&c<=0xACC)||c===0xAD0||(c>=0xAE0&&c<=0xAE1)||(c>=0xAE6&&c<=0xAF0)||(c>=0xB02&&c<=0xB03)||(c>=0xB05&&c<=0xB0C)||(c>=0xB0F&&c<=0xB10)||(c>=0xB13&&c<=0xB28)||(c>=0xB2A&&c<=0xB30)||(c>=0xB32&&c<=0xB33)||(c>=0xB35&&c<=0xB39)||(c>=0xB3D&&c<=0xB3E)||c===0xB40||(c>=0xB47&&c<=0xB48)||(c>=0xB4B&&c<=0xB4C)||c===0xB57||(c>=0xB5C&&c<=0xB5D)||(c>=0xB5F&&c<=0xB61)||(c>=0xB66&&c<=0xB77)||c===0xB83||(c>=0xB85&&c<=0xB8A)||(c>=0xB8E&&c<=0xB90)||(c>=0xB92&&c<=0xB95)||(c>=0xB99&&c<=0xB9A)||c===0xB9C||(c>=0xB9E&&c<=0xB9F)||(c>=0xBA3&&c<=0xBA4)||(c>=0xBA8&&c<=0xBAA)||(c>=0xBAE&&c<=0xBB9)||(c>=0xBBE&&c<=0xBBF)||(c>=0xBC1&&c<=0xBC2)||(c>=0xBC6&&c<=0xBC8)||(c>=0xBCA&&c<=0xBCC)||c===0xBD0||c===0xBD7||(c>=0xBE6&&c<=0xBF2)||(c>=0xC01&&c<=0xC03)||(c>=0xC05&&c<=0xC0C)||(c>=0xC0E&&c<=0xC10)||(c>=0xC12&&c<=0xC28)||(c>=0xC2A&&c<=0xC39)||c===0xC3D||(c>=0xC41&&c<=0xC44)||(c>=0xC58&&c<=0xC59)||(c>=0xC60&&c<=0xC61)||(c>=0xC66&&c<=0xC6F)||c===0xC7F||(c>=0xC82&&c<=0xC83)||(c>=0xC85&&c<=0xC8C)||(c>=0xC8E&&c<=0xC90)||(c>=0xC92&&c<=0xCA8)||(c>=0xCAA&&c<=0xCB3)||(c>=0xCB5&&c<=0xCB9)||(c>=0xCBD&&c<=0xCC4)||(c>=0xCC6&&c<=0xCC8)||(c>=0xCCA&&c<=0xCCB)||(c>=0xCD5&&c<=0xCD6)||c===0xCDE||(c>=0xCE0&&c<=0xCE1)||(c>=0xCE6&&c<=0xCEF)||(c>=0xCF1&&c<=0xCF2)||(c>=0xD02&&c<=0xD03)||(c>=0xD05&&c<=0xD0C)||(c>=0xD0E&&c<=0xD10)||(c>=0xD12&&c<=0xD3A)||(c>=0xD3D&&c<=0xD40)||(c>=0xD46&&c<=0xD48)||(c>=0xD4A&&c<=0xD4C)||c===0xD4E||c===0xD57||(c>=0xD60&&c<=0xD61)||(c>=0xD66&&c<=0xD75)||(c>=0xD79&&c<=0xD7F)||(c>=0xD82&&c<=0xD83)||(c>=0xD85&&c<=0xD96)||(c>=0xD9A&&c<=0xDB1)||(c>=0xDB3&&c<=0xDBB)||c===0xDBD||(c>=0xDC0&&c<=0xDC6)||(c>=0xDCF&&c<=0xDD1)||(c>=0xDD8&&c<=0xDDF)||(c>=0xDE6&&c<=0xDEF)||(c>=0xDF2&&c<=0xDF4)||(c>=0xE01&&c<=0xE30)||(c>=0xE32&&c<=0xE33)||(c>=0xE40&&c<=0xE46)||(c>=0xE4F&&c<=0xE5B)||(c>=0xE81&&c<=0xE82)||c===0xE84||(c>=0xE87&&c<=0xE88)||c===0xE8A||c===0xE8D||(c>=0xE94&&c<=0xE97)||(c>=0xE99&&c<=0xE9F)||(c>=0xEA1&&c<=0xEA3)||c===0xEA5||c===0xEA7||(c>=0xEAA&&c<=0xEAB)||(c>=0xEAD&&c<=0xEB0)||(c>=0xEB2&&c<=0xEB3)||c===0xEBD||(c>=0xEC0&&c<=0xEC4)||c===0xEC6||(c>=0xED0&&c<=0xED9)||(c>=0xEDC&&c<=0xEDF)||(c>=0xF00&&c<=0xF17)||(c>=0xF1A&&c<=0xF34)||c===0xF36||c===0xF38||(c>=0xF3E&&c<=0xF47)||(c>=0xF49&&c<=0xF6C)||c===0xF7F||c===0xF85||(c>=0xF88&&c<=0xF8C)||(c>=0xFBE&&c<=0xFC5)||(c>=0xFC7&&c<=0xFCC)||(c>=0xFCE&&c<=0xFDA)||(c>=0x1000&&c<=0x102C)||c===0x1031||c===0x1038||(c>=0x103B&&c<=0x103C)||(c>=0x103F&&c<=0x1057)||(c>=0x105A&&c<=0x105D)||(c>=0x1061&&c<=0x1070)||(c>=0x1075&&c<=0x1081)||(c>=0x1083&&c<=0x1084)||(c>=0x1087&&c<=0x108C)||(c>=0x108E&&c<=0x109C)||(c>=0x109E&&c<=0x10C5)||c===0x10C7||c===0x10CD||(c>=0x10D0&&c<=0x1248)||(c>=0x124A&&c<=0x124D)||(c>=0x1250&&c<=0x1256)||c===0x1258||(c>=0x125A&&c<=0x125D)||(c>=0x1260&&c<=0x1288)||(c>=0x128A&&c<=0x128D)||(c>=0x1290&&c<=0x12B0)||(c>=0x12B2&&c<=0x12B5)||(c>=0x12B8&&c<=0x12BE)||c===0x12C0||(c>=0x12C2&&c<=0x12C5)||(c>=0x12C8&&c<=0x12D6)||(c>=0x12D8&&c<=0x1310)||(c>=0x1312&&c<=0x1315)||(c>=0x1318&&c<=0x135A)||(c>=0x1360&&c<=0x137C)||(c>=0x1380&&c<=0x138F)||(c>=0x13A0&&c<=0x13F4)||(c>=0x1401&&c<=0x167F)||(c>=0x1681&&c<=0x169A)||(c>=0x16A0&&c<=0x16F8)||(c>=0x1700&&c<=0x170C)||(c>=0x170E&&c<=0x1711)||(c>=0x1720&&c<=0x1731)||(c>=0x1735&&c<=0x1736)||(c>=0x1740&&c<=0x1751)||(c>=0x1760&&c<=0x176C)||(c>=0x176E&&c<=0x1770)||(c>=0x1780&&c<=0x17B3)||c===0x17B6||(c>=0x17BE&&c<=0x17C5)||(c>=0x17C7&&c<=0x17C8)||(c>=0x17D4&&c<=0x17DA)||c===0x17DC||(c>=0x17E0&&c<=0x17E9)||(c>=0x1810&&c<=0x1819)||(c>=0x1820&&c<=0x1877)||(c>=0x1880&&c<=0x18A8)||c===0x18AA||(c>=0x18B0&&c<=0x18F5)||(c>=0x1900&&c<=0x191E)||(c>=0x1923&&c<=0x1926)||(c>=0x1929&&c<=0x192B)||(c>=0x1930&&c<=0x1931)||(c>=0x1933&&c<=0x1938)||(c>=0x1946&&c<=0x196D)||(c>=0x1970&&c<=0x1974)||(c>=0x1980&&c<=0x19AB)||(c>=0x19B0&&c<=0x19C9)||(c>=0x19D0&&c<=0x19DA)||(c>=0x1A00&&c<=0x1A16)||(c>=0x1A19&&c<=0x1A1A)||(c>=0x1A1E&&c<=0x1A55)||c===0x1A57||c===0x1A61||(c>=0x1A63&&c<=0x1A64)||(c>=0x1A6D&&c<=0x1A72)||(c>=0x1A80&&c<=0x1A89)||(c>=0x1A90&&c<=0x1A99)||(c>=0x1AA0&&c<=0x1AAD)||(c>=0x1B04&&c<=0x1B33)||c===0x1B35||c===0x1B3B||(c>=0x1B3D&&c<=0x1B41)||(c>=0x1B43&&c<=0x1B4B)||(c>=0x1B50&&c<=0x1B6A)||(c>=0x1B74&&c<=0x1B7C)||(c>=0x1B82&&c<=0x1BA1)||(c>=0x1BA6&&c<=0x1BA7)||c===0x1BAA||(c>=0x1BAE&&c<=0x1BE5)||c===0x1BE7||(c>=0x1BEA&&c<=0x1BEC)||c===0x1BEE||(c>=0x1BF2&&c<=0x1BF3)||(c>=0x1BFC&&c<=0x1C2B)||(c>=0x1C34&&c<=0x1C35)||(c>=0x1C3B&&c<=0x1C49)||(c>=0x1C4D&&c<=0x1C7F)||(c>=0x1CC0&&c<=0x1CC7)||c===0x1CD3||c===0x1CE1||(c>=0x1CE9&&c<=0x1CEC)||(c>=0x1CEE&&c<=0x1CF3)||(c>=0x1CF5&&c<=0x1CF6)||(c>=0x1D00&&c<=0x1DBF)||(c>=0x1E00&&c<=0x1F15)||(c>=0x1F18&&c<=0x1F1D)||(c>=0x1F20&&c<=0x1F45)||(c>=0x1F48&&c<=0x1F4D)||(c>=0x1F50&&c<=0x1F57)||c===0x1F59||c===0x1F5B||c===0x1F5D||(c>=0x1F5F&&c<=0x1F7D)||(c>=0x1F80&&c<=0x1FB4)||(c>=0x1FB6&&c<=0x1FBC)||c===0x1FBE||(c>=0x1FC2&&c<=0x1FC4)||(c>=0x1FC6&&c<=0x1FCC)||(c>=0x1FD0&&c<=0x1FD3)||(c>=0x1FD6&&c<=0x1FDB)||(c>=0x1FE0&&c<=0x1FEC)||(c>=0x1FF2&&c<=0x1FF4)||(c>=0x1FF6&&c<=0x1FFC)||c===0x200E||c===0x2071||c===0x207F||(c>=0x2090&&c<=0x209C)||c===0x2102||c===0x2107||(c>=0x210A&&c<=0x2113)||c===0x2115||(c>=0x2119&&c<=0x211D)||c===0x2124||c===0x2126||c===0x2128||(c>=0x212A&&c<=0x212D)||(c>=0x212F&&c<=0x2139)||(c>=0x213C&&c<=0x213F)||(c>=0x2145&&c<=0x2149)||(c>=0x214E&&c<=0x214F)||(c>=0x2160&&c<=0x2188)||(c>=0x2336&&c<=0x237A)||c===0x2395||(c>=0x249C&&c<=0x24E9)||c===0x26AC||(c>=0x2800&&c<=0x28FF)||(c>=0x2C00&&c<=0x2C2E)||(c>=0x2C30&&c<=0x2C5E)||(c>=0x2C60&&c<=0x2CE4)||(c>=0x2CEB&&c<=0x2CEE)||(c>=0x2CF2&&c<=0x2CF3)||(c>=0x2D00&&c<=0x2D25)||c===0x2D27||c===0x2D2D||(c>=0x2D30&&c<=0x2D67)||(c>=0x2D6F&&c<=0x2D70)||(c>=0x2D80&&c<=0x2D96)||(c>=0x2DA0&&c<=0x2DA6)||(c>=0x2DA8&&c<=0x2DAE)||(c>=0x2DB0&&c<=0x2DB6)||(c>=0x2DB8&&c<=0x2DBE)||(c>=0x2DC0&&c<=0x2DC6)||(c>=0x2DC8&&c<=0x2DCE)||(c>=0x2DD0&&c<=0x2DD6)||(c>=0x2DD8&&c<=0x2DDE)||(c>=0x3005&&c<=0x3007)||(c>=0x3021&&c<=0x3029)||(c>=0x302E&&c<=0x302F)||(c>=0x3031&&c<=0x3035)||(c>=0x3038&&c<=0x303C)||(c>=0x3041&&c<=0x3096)||(c>=0x309D&&c<=0x309F)||(c>=0x30A1&&c<=0x30FA)||(c>=0x30FC&&c<=0x30FF)||(c>=0x3105&&c<=0x312D)||(c>=0x3131&&c<=0x318E)||(c>=0x3190&&c<=0x31BA)||(c>=0x31F0&&c<=0x321C)||(c>=0x3220&&c<=0x324F)||(c>=0x3260&&c<=0x327B)||(c>=0x327F&&c<=0x32B0)||(c>=0x32C0&&c<=0x32CB)||(c>=0x32D0&&c<=0x32FE)||(c>=0x3300&&c<=0x3376)||(c>=0x337B&&c<=0x33DD)||(c>=0x33E0&&c<=0x33FE)||c===0x3400||c===0x4DB5||c===0x4E00||c===0x9FCC||(c>=0xA000&&c<=0xA48C)||(c>=0xA4D0&&c<=0xA60C)||(c>=0xA610&&c<=0xA62B)||(c>=0xA640&&c<=0xA66E)||(c>=0xA680&&c<=0xA69D)||(c>=0xA6A0&&c<=0xA6EF)||(c>=0xA6F2&&c<=0xA6F7)||(c>=0xA722&&c<=0xA787)||(c>=0xA789&&c<=0xA78E)||(c>=0xA790&&c<=0xA7AD)||(c>=0xA7B0&&c<=0xA7B1)||(c>=0xA7F7&&c<=0xA801)||(c>=0xA803&&c<=0xA805)||(c>=0xA807&&c<=0xA80A)||(c>=0xA80C&&c<=0xA824)||c===0xA827||(c>=0xA830&&c<=0xA837)||(c>=0xA840&&c<=0xA873)||(c>=0xA880&&c<=0xA8C3)||(c>=0xA8CE&&c<=0xA8D9)||(c>=0xA8F2&&c<=0xA8FB)||(c>=0xA900&&c<=0xA925)||(c>=0xA92E&&c<=0xA946)||(c>=0xA952&&c<=0xA953)||(c>=0xA95F&&c<=0xA97C)||(c>=0xA983&&c<=0xA9B2)||(c>=0xA9B4&&c<=0xA9B5)||(c>=0xA9BA&&c<=0xA9BB)||(c>=0xA9BD&&c<=0xA9CD)||(c>=0xA9CF&&c<=0xA9D9)||(c>=0xA9DE&&c<=0xA9E4)||(c>=0xA9E6&&c<=0xA9FE)||(c>=0xAA00&&c<=0xAA28)||(c>=0xAA2F&&c<=0xAA30)||(c>=0xAA33&&c<=0xAA34)||(c>=0xAA40&&c<=0xAA42)||(c>=0xAA44&&c<=0xAA4B)||c===0xAA4D||(c>=0xAA50&&c<=0xAA59)||(c>=0xAA5C&&c<=0xAA7B)||(c>=0xAA7D&&c<=0xAAAF)||c===0xAAB1||(c>=0xAAB5&&c<=0xAAB6)||(c>=0xAAB9&&c<=0xAABD)||c===0xAAC0||c===0xAAC2||(c>=0xAADB&&c<=0xAAEB)||(c>=0xAAEE&&c<=0xAAF5)||(c>=0xAB01&&c<=0xAB06)||(c>=0xAB09&&c<=0xAB0E)||(c>=0xAB11&&c<=0xAB16)||(c>=0xAB20&&c<=0xAB26)||(c>=0xAB28&&c<=0xAB2E)||(c>=0xAB30&&c<=0xAB5F)||(c>=0xAB64&&c<=0xAB65)||(c>=0xABC0&&c<=0xABE4)||(c>=0xABE6&&c<=0xABE7)||(c>=0xABE9&&c<=0xABEC)||(c>=0xABF0&&c<=0xABF9)||c===0xAC00||c===0xD7A3||(c>=0xD7B0&&c<=0xD7C6)||(c>=0xD7CB&&c<=0xD7FB)||c===0xD800||(c>=0xDB7F&&c<=0xDB80)||(c>=0xDBFF&&c<=0xDC00)||(c>=0xDFFF&&c<=0xE000)||(c>=0xF8FF&&c<=0xFA6D)||(c>=0xFA70&&c<=0xFAD9)||(c>=0xFB00&&c<=0xFB06)||(c>=0xFB13&&c<=0xFB17)||(c>=0xFF21&&c<=0xFF3A)||(c>=0xFF41&&c<=0xFF5A)||(c>=0xFF66&&c<=0xFFBE)||(c>=0xFFC2&&c<=0xFFC7)||(c>=0xFFCA&&c<=0xFFCF)||(c>=0xFFD2&&c<=0xFFD7)||(c>=0xFFDA&&c<=0xFFDC)||(c>=0x10000&&c<=0x1000B)||(c>=0x1000D&&c<=0x10026)||(c>=0x10028&&c<=0x1003A)||(c>=0x1003C&&c<=0x1003D)||(c>=0x1003F&&c<=0x1004D)||(c>=0x10050&&c<=0x1005D)||(c>=0x10080&&c<=0x100FA)||c===0x10100||c===0x10102||(c>=0x10107&&c<=0x10133)||(c>=0x10137&&c<=0x1013F)||(c>=0x101D0&&c<=0x101FC)||(c>=0x10280&&c<=0x1029C)||(c>=0x102A0&&c<=0x102D0)||(c>=0x10300&&c<=0x10323)||(c>=0x10330&&c<=0x1034A)||(c>=0x10350&&c<=0x10375)||(c>=0x10380&&c<=0x1039D)||(c>=0x1039F&&c<=0x103C3)||(c>=0x103C8&&c<=0x103D5)||(c>=0x10400&&c<=0x1049D)||(c>=0x104A0&&c<=0x104A9)||(c>=0x10500&&c<=0x10527)||(c>=0x10530&&c<=0x10563)||c===0x1056F||(c>=0x10600&&c<=0x10736)||(c>=0x10740&&c<=0x10755)||(c>=0x10760&&c<=0x10767)||c===0x11000||(c>=0x11002&&c<=0x11037)||(c>=0x11047&&c<=0x1104D)||(c>=0x11066&&c<=0x1106F)||(c>=0x11082&&c<=0x110B2)||(c>=0x110B7&&c<=0x110B8)||(c>=0x110BB&&c<=0x110C1)||(c>=0x110D0&&c<=0x110E8)||(c>=0x110F0&&c<=0x110F9)||(c>=0x11103&&c<=0x11126)||c===0x1112C||(c>=0x11136&&c<=0x11143)||(c>=0x11150&&c<=0x11172)||(c>=0x11174&&c<=0x11176)||(c>=0x11182&&c<=0x111B5)||(c>=0x111BF&&c<=0x111C8)||c===0x111CD||(c>=0x111D0&&c<=0x111DA)||(c>=0x111E1&&c<=0x111F4)||(c>=0x11200&&c<=0x11211)||(c>=0x11213&&c<=0x1122E)||(c>=0x11232&&c<=0x11233)||c===0x11235||(c>=0x11238&&c<=0x1123D)||(c>=0x112B0&&c<=0x112DE)||(c>=0x112E0&&c<=0x112E2)||(c>=0x112F0&&c<=0x112F9)||(c>=0x11302&&c<=0x11303)||(c>=0x11305&&c<=0x1130C)||(c>=0x1130F&&c<=0x11310)||(c>=0x11313&&c<=0x11328)||(c>=0x1132A&&c<=0x11330)||(c>=0x11332&&c<=0x11333)||(c>=0x11335&&c<=0x11339)||(c>=0x1133D&&c<=0x1133F)||(c>=0x11341&&c<=0x11344)||(c>=0x11347&&c<=0x11348)||(c>=0x1134B&&c<=0x1134D)||c===0x11357||(c>=0x1135D&&c<=0x11363)||(c>=0x11480&&c<=0x114B2)||c===0x114B9||(c>=0x114BB&&c<=0x114BE)||c===0x114C1||(c>=0x114C4&&c<=0x114C7)||(c>=0x114D0&&c<=0x114D9)||(c>=0x11580&&c<=0x115B1)||(c>=0x115B8&&c<=0x115BB)||c===0x115BE||(c>=0x115C1&&c<=0x115C9)||(c>=0x11600&&c<=0x11632)||(c>=0x1163B&&c<=0x1163C)||c===0x1163E||(c>=0x11641&&c<=0x11644)||(c>=0x11650&&c<=0x11659)||(c>=0x11680&&c<=0x116AA)||c===0x116AC||(c>=0x116AE&&c<=0x116AF)||c===0x116B6||(c>=0x116C0&&c<=0x116C9)||(c>=0x118A0&&c<=0x118F2)||c===0x118FF||(c>=0x11AC0&&c<=0x11AF8)||(c>=0x12000&&c<=0x12398)||(c>=0x12400&&c<=0x1246E)||(c>=0x12470&&c<=0x12474)||(c>=0x13000&&c<=0x1342E)||(c>=0x16800&&c<=0x16A38)||(c>=0x16A40&&c<=0x16A5E)||(c>=0x16A60&&c<=0x16A69)||(c>=0x16A6E&&c<=0x16A6F)||(c>=0x16AD0&&c<=0x16AED)||c===0x16AF5||(c>=0x16B00&&c<=0x16B2F)||(c>=0x16B37&&c<=0x16B45)||(c>=0x16B50&&c<=0x16B59)||(c>=0x16B5B&&c<=0x16B61)||(c>=0x16B63&&c<=0x16B77)||(c>=0x16B7D&&c<=0x16B8F)||(c>=0x16F00&&c<=0x16F44)||(c>=0x16F50&&c<=0x16F7E)||(c>=0x16F93&&c<=0x16F9F)||(c>=0x1B000&&c<=0x1B001)||(c>=0x1BC00&&c<=0x1BC6A)||(c>=0x1BC70&&c<=0x1BC7C)||(c>=0x1BC80&&c<=0x1BC88)||(c>=0x1BC90&&c<=0x1BC99)||c===0x1BC9C||c===0x1BC9F||(c>=0x1D000&&c<=0x1D0F5)||(c>=0x1D100&&c<=0x1D126)||(c>=0x1D129&&c<=0x1D166)||(c>=0x1D16A&&c<=0x1D172)||(c>=0x1D183&&c<=0x1D184)||(c>=0x1D18C&&c<=0x1D1A9)||(c>=0x1D1AE&&c<=0x1D1DD)||(c>=0x1D360&&c<=0x1D371)||(c>=0x1D400&&c<=0x1D454)||(c>=0x1D456&&c<=0x1D49C)||(c>=0x1D49E&&c<=0x1D49F)||c===0x1D4A2||(c>=0x1D4A5&&c<=0x1D4A6)||(c>=0x1D4A9&&c<=0x1D4AC)||(c>=0x1D4AE&&c<=0x1D4B9)||c===0x1D4BB||(c>=0x1D4BD&&c<=0x1D4C3)||(c>=0x1D4C5&&c<=0x1D505)||(c>=0x1D507&&c<=0x1D50A)||(c>=0x1D50D&&c<=0x1D514)||(c>=0x1D516&&c<=0x1D51C)||(c>=0x1D51E&&c<=0x1D539)||(c>=0x1D53B&&c<=0x1D53E)||(c>=0x1D540&&c<=0x1D544)||c===0x1D546||(c>=0x1D54A&&c<=0x1D550)||(c>=0x1D552&&c<=0x1D6A5)||(c>=0x1D6A8&&c<=0x1D6DA)||(c>=0x1D6DC&&c<=0x1D714)||(c>=0x1D716&&c<=0x1D74E)||(c>=0x1D750&&c<=0x1D788)||(c>=0x1D78A&&c<=0x1D7C2)||(c>=0x1D7C4&&c<=0x1D7CB)||(c>=0x1F110&&c<=0x1F12E)||(c>=0x1F130&&c<=0x1F169)||(c>=0x1F170&&c<=0x1F19A)||(c>=0x1F1E6&&c<=0x1F202)||(c>=0x1F210&&c<=0x1F23A)||(c>=0x1F240&&c<=0x1F248)||(c>=0x1F250&&c<=0x1F251)||c===0x20000||c===0x2A6D6||c===0x2A700||c===0x2B734||c===0x2B740||c===0x2B81D||(c>=0x2F800&&c<=0x2FA1D)||c===0xF0000||c===0xFFFFD||c===0x100000||c===0x10FFFD; }", _
        "function IsRTL(c) { return c===0x5BE||c===0x5C0||c===0x5C3||c===0x5C6||(c>=0x5D0&&c<=0x5EA)||(c>=0x5F0&&c<=0x5F4)||c===0x608||c===0x60B||c===0x60D||(c>=0x61B&&c<=0x61C)||(c>=0x61E&&c<=0x64A)||(c>=0x66D&&c<=0x66F)||(c>=0x671&&c<=0x6D5)||(c>=0x6E5&&c<=0x6E6)||(c>=0x6EE&&c<=0x6EF)||(c>=0x6FA&&c<=0x70D)||(c>=0x70F&&c<=0x710)||(c>=0x712&&c<=0x72F)||(c>=0x74D&&c<=0x7A5)||c===0x7B1||(c>=0x7C0&&c<=0x7EA)||(c>=0x7F4&&c<=0x7F5)||c===0x7FA||(c>=0x800&&c<=0x815)||c===0x81A||c===0x824||c===0x828||(c>=0x830&&c<=0x83E)||(c>=0x840&&c<=0x858)||c===0x85E||(c>=0x8A0&&c<=0x8B2)||c===0x200F||c===0xFB1D||(c>=0xFB1F&&c<=0xFB28)||(c>=0xFB2A&&c<=0xFB36)||(c>=0xFB38&&c<=0xFB3C)||c===0xFB3E||(c>=0xFB40&&c<=0xFB41)||(c>=0xFB43&&c<=0xFB44)||(c>=0xFB46&&c<=0xFBC1)||(c>=0xFBD3&&c<=0xFD3D)||(c>=0xFD50&&c<=0xFD8F)||(c>=0xFD92&&c<=0xFDC7)||(c>=0xFDF0&&c<=0xFDFC)||(c>=0xFE70&&c<=0xFE74)||(c>=0xFE76&&c<=0xFEFC)||(c>=0x10800&&c<=0x10805)||c===0x10808||(c>=0x1080A&&c<=0x10835)||(c>=0x10837&&c<=0x10838)||c===0x1083C||(c>=0x1083F&&c<=0x10855)||(c>=0x10857&&c<=0x1089E)||(c>=0x108A7&&c<=0x108AF)||(c>=0x10900&&c<=0x1091B)||(c>=0x10920&&c<=0x10939)||c===0x1093F||(c>=0x10980&&c<=0x109B7)||(c>=0x109BE&&c<=0x109BF)||c===0x10A00||(c>=0x10A10&&c<=0x10A13)||(c>=0x10A15&&c<=0x10A17)||(c>=0x10A19&&c<=0x10A33)||(c>=0x10A40&&c<=0x10A47)||(c>=0x10A50&&c<=0x10A58)||(c>=0x10A60&&c<=0x10A9F)||(c>=0x10AC0&&c<=0x10AE4)||(c>=0x10AEB&&c<=0x10AF6)||(c>=0x10B00&&c<=0x10B35)||(c>=0x10B40&&c<=0x10B55)||(c>=0x10B58&&c<=0x10B72)||(c>=0x10B78&&c<=0x10B91)||(c>=0x10B99&&c<=0x10B9C)||(c>=0x10BA9&&c<=0x10BAF)||(c>=0x10C00&&c<=0x10C48)||(c>=0x1E800&&c<=0x1E8C4)||(c>=0x1E8C7&&c<=0x1E8CF)||(c>=0x1EE00&&c<=0x1EE03)||(c>=0x1EE05&&c<=0x1EE1F)||(c>=0x1EE21&&c<=0x1EE22)||c===0x1EE24||c===0x1EE27||(c>=0x1EE29&&c<=0x1EE32)||(c>=0x1EE34&&c<=0x1EE37)||c===0x1EE39||c===0x1EE3B||c===0x1EE42||c===0x1EE47||c===0x1EE49||c===0x1EE4B||(c>=0x1EE4D&&c<=0x1EE4F)||(c>=0x1EE51&&c<=0x1EE52)||c===0x1EE54||c===0x1EE57||c===0x1EE59||c===0x1EE5B||c===0x1EE5D||c===0x1EE5F||(c>=0x1EE61&&c<=0x1EE62)||c===0x1EE64||(c>=0x1EE67&&c<=0x1EE6A)||(c>=0x1EE6C&&c<=0x1EE72)||(c>=0x1EE74&&c<=0x1EE77)||(c>=0x1EE79&&c<=0x1EE7C)||c===0x1EE7E||(c>=0x1EE80&&c<=0x1EE89)||(c>=0x1EE8B&&c<=0x1EE9B)||(c>=0x1EEA1&&c<=0x1EEA3)||(c>=0x1EEA5&&c<=0x1EEA9)||(c>=0x1EEAB&&c<=0x1EEBB); }", _
        "function IsAL(c) { return c===0x608||c===0x60B||c===0x60D||(c>=0x61B&&c<=0x61C)||(c>=0x61E&&c<=0x64A)||(c>=0x66D&&c<=0x66F)||(c>=0x671&&c<=0x6D5)||(c>=0x6E5&&c<=0x6E6)||(c>=0x6EE&&c<=0x6EF)||(c>=0x6FA&&c<=0x70D)||(c>=0x70F&&c<=0x710)||(c>=0x712&&c<=0x72F)||(c>=0x74D&&c<=0x7A5)||c===0x7B1||(c>=0x8A0&&c<=0x8B2)||(c>=0xFB50&&c<=0xFBC1)||(c>=0xFBD3&&c<=0xFD3D)||(c>=0xFD50&&c<=0xFD8F)||(c>=0xFD92&&c<=0xFDC7)||(c>=0xFDF0&&c<=0xFDFC)||(c>=0xFE70&&c<=0xFE74)||(c>=0xFE76&&c<=0xFEFC)||(c>=0x1EE00&&c<=0x1EE03)||(c>=0x1EE05&&c<=0x1EE1F)||(c>=0x1EE21&&c<=0x1EE22)||c===0x1EE24||c===0x1EE27||(c>=0x1EE29&&c<=0x1EE32)||(c>=0x1EE34&&c<=0x1EE37)||c===0x1EE39||c===0x1EE3B||c===0x1EE42||c===0x1EE47||c===0x1EE49||c===0x1EE4B||(c>=0x1EE4D&&c<=0x1EE4F)||(c>=0x1EE51&&c<=0x1EE52)||c===0x1EE54||c===0x1EE57||c===0x1EE59||c===0x1EE5B||c===0x1EE5D||c===0x1EE5F||(c>=0x1EE61&&c<=0x1EE62)||c===0x1EE64||(c>=0x1EE67&&c<=0x1EE6A)||(c>=0x1EE6C&&c<=0x1EE72)||(c>=0x1EE74&&c<=0x1EE77)||(c>=0x1EE79&&c<=0x1EE7C)||c===0x1EE7E||(c>=0x1EE80&&c<=0x1EE89)||(c>=0x1EE8B&&c<=0x1EE9B)||(c>=0x1EEA1&&c<=0x1EEA3)||(c>=0x1EEA5&&c<=0x1EEA9)||(c>=0x1EEAB&&c<=0x1EEBB); }", _
        "function IsNeutral(c) { return (c>=0x9&&c<=0xD)||(c>=0x1C&&c<=0x22)||(c>=0x26&&c<=0x2A)||(c>=0x3B&&c<=0x40)||(c>=0x5B&&c<=0x60)||(c>=0x7B&&c<=0x7E)||c===0x85||c===0xA1||(c>=0xA6&&c<=0xA9)||(c>=0xAB&&c<=0xAC)||(c>=0xAE&&c<=0xAF)||c===0xB4||(c>=0xB6&&c<=0xB8)||(c>=0xBB&&c<=0xBF)||c===0xD7||c===0xF7||(c>=0x2B9&&c<=0x2BA)||(c>=0x2C2&&c<=0x2CF)||(c>=0x2D2&&c<=0x2DF)||(c>=0x2E5&&c<=0x2ED)||(c>=0x2EF&&c<=0x2FF)||(c>=0x374&&c<=0x375)||c===0x37E||(c>=0x384&&c<=0x385)||c===0x387||c===0x3F6||c===0x58A||(c>=0x58D&&c<=0x58E)||(c>=0x606&&c<=0x607)||(c>=0x60E&&c<=0x60F)||c===0x6DE||c===0x6E9||(c>=0x7F6&&c<=0x7F9)||(c>=0xBF3&&c<=0xBF8)||c===0xBFA||(c>=0xC78&&c<=0xC7E)||(c>=0xF3A&&c<=0xF3D)||(c>=0x1390&&c<=0x1399)||c===0x1400||c===0x1680||(c>=0x169B&&c<=0x169C)||(c>=0x17F0&&c<=0x17F9)||(c>=0x1800&&c<=0x180A)||c===0x1940||(c>=0x1944&&c<=0x1945)||(c>=0x19DE&&c<=0x19FF)||c===0x1FBD||(c>=0x1FBF&&c<=0x1FC1)||(c>=0x1FCD&&c<=0x1FCF)||(c>=0x1FDD&&c<=0x1FDF)||(c>=0x1FED&&c<=0x1FEF)||(c>=0x1FFD&&c<=0x1FFE)||(c>=0x2000&&c<=0x200A)||(c>=0x2010&&c<=0x2029)||(c>=0x2035&&c<=0x2043)||(c>=0x2045&&c<=0x205F)||(c>=0x207C&&c<=0x207E)||(c>=0x208C&&c<=0x208E)||(c>=0x2100&&c<=0x2101)||(c>=0x2103&&c<=0x2106)||(c>=0x2108&&c<=0x2109)||c===0x2114||(c>=0x2116&&c<=0x2118)||(c>=0x211E&&c<=0x2123)||c===0x2125||c===0x2127||c===0x2129||(c>=0x213A&&c<=0x213B)||(c>=0x2140&&c<=0x2144)||(c>=0x214A&&c<=0x214D)||(c>=0x2150&&c<=0x215F)||c===0x2189||(c>=0x2190&&c<=0x2211)||(c>=0x2214&&c<=0x2335)||(c>=0x237B&&c<=0x2394)||(c>=0x2396&&c<=0x23FA)||(c>=0x2400&&c<=0x2426)||(c>=0x2440&&c<=0x244A)||(c>=0x2460&&c<=0x2487)||(c>=0x24EA&&c<=0x26AB)||(c>=0x26AD&&c<=0x27FF)||(c>=0x2900&&c<=0x2B73)||(c>=0x2B76&&c<=0x2B95)||(c>=0x2B98&&c<=0x2BB9)||(c>=0x2BBD&&c<=0x2BC8)||(c>=0x2BCA&&c<=0x2BD1)||(c>=0x2CE5&&c<=0x2CEA)||(c>=0x2CF9&&c<=0x2CFF)||(c>=0x2E00&&c<=0x2E42)||(c>=0x2E80&&c<=0x2E99)||(c>=0x2E9B&&c<=0x2EF3)||(c>=0x2F00&&c<=0x2FD5)||(c>=0x2FF0&&c<=0x2FFB)||(c>=0x3000&&c<=0x3004)||(c>=0x3008&&c<=0x3020)||c===0x3030||(c>=0x3036&&c<=0x3037)||(c>=0x303D&&c<=0x303F)||(c>=0x309B&&c<=0x309C)||c===0x30A0||c===0x30FB||(c>=0x31C0&&c<=0x31E3)||(c>=0x321D&&c<=0x321E)||(c>=0x3250&&c<=0x325F)||(c>=0x327C&&c<=0x327E)||(c>=0x32B1&&c<=0x32BF)||(c>=0x32CC&&c<=0x32CF)||(c>=0x3377&&c<=0x337A)||(c>=0x33DE&&c<=0x33DF)||c===0x33FF||(c>=0x4DC0&&c<=0x4DFF)||(c>=0xA490&&c<=0xA4C6)||(c>=0xA60D&&c<=0xA60F)||c===0xA673||(c>=0xA67E&&c<=0xA67F)||(c>=0xA700&&c<=0xA721)||c===0xA788||(c>=0xA828&&c<=0xA82B)||(c>=0xA874&&c<=0xA877)||(c>=0xFD3E&&c<=0xFD3F)||c===0xFDFD||(c>=0xFE10&&c<=0xFE19)||(c>=0xFE30&&c<=0xFE4F)||c===0xFE51||c===0xFE54||(c>=0xFE56&&c<=0xFE5E)||(c>=0xFE60&&c<=0xFE61)||(c>=0xFE64&&c<=0xFE66)||c===0xFE68||c===0xFE6B||(c>=0xFF01&&c<=0xFF02)||(c>=0xFF06&&c<=0xFF0A)||(c>=0xFF1B&&c<=0xFF20)||(c>=0xFF3B&&c<=0xFF40)||(c>=0xFF5B&&c<=0xFF65)||(c>=0xFFE2&&c<=0xFFE4)||(c>=0xFFE8&&c<=0xFFEE)||(c>=0xFFF9&&c<=0xFFFD)||c===0x10101||(c>=0x10140&&c<=0x1018C)||(c>=0x10190&&c<=0x1019B)||c===0x101A0||c===0x1091F||(c>=0x10B39&&c<=0x10B3F)||(c>=0x11052&&c<=0x11065)||(c>=0x1D200&&c<=0x1D241)||c===0x1D245||(c>=0x1D300&&c<=0x1D356)||c===0x1D6DB||c===0x1D715||c===0x1D74F||c===0x1D789||c===0x1D7C3||(c>=0x1EEF0&&c<=0x1EEF1)||(c>=0x1F000&&c<=0x1F02B)||(c>=0x1F030&&c<=0x1F093)||(c>=0x1F0A0&&c<=0x1F0AE)||(c>=0x1F0B1&&c<=0x1F0BF)||(c>=0x1F0C1&&c<=0x1F0CF)||(c>=0x1F0D1&&c<=0x1F0F5)||(c>=0x1F10B&&c<=0x1F10C)||(c>=0x1F16A&&c<=0x1F16B)||(c>=0x1F300&&c<=0x1F32C)||(c>=0x1F330&&c<=0x1F37D)||(c>=0x1F380&&c<=0x1F3CE)||(c>=0x1F3D4&&c<=0x1F3F7)||(c>=0x1F400&&c<=0x1F4FE)||(c>=0x1F500&&c<=0x1F54A)||(c>=0x1F550&&c<=0x1F579)||(c>=0x1F57B&&c<=0x1F5A3)||(c>=0x1F5A5&&c<=0x1F642)||(c>=0x1F645&&c<=0x1F6CF)||(c>=0x1F6E0&&c<=0x1F6EC)||(c>=0x1F6F0&&c<=0x1F6F3)||(c>=0x1F700&&c<=0x1F773)||(c>=0x1F780&&c<=0x1F7D4)||(c>=0x1F800&&c<=0x1F80B)||(c>=0x1F810&&c<=0x1F847)||(c>=0x1F850&&c<=0x1F859)||(c>=0x1F860&&c<=0x1F887)||(c>=0x1F890&&c<=0x1F8AD); }", _
        "function IsWeak(c) { return (c>=0x0&&c<=0x8)||(c>=0xE&&c<=0x1B)||(c>=0x23&&c<=0x25)||(c>=0x2B&&c<=0x3A)||(c>=0x7F&&c<=0x84)||(c>=0x86&&c<=0xA0)||(c>=0xA2&&c<=0xA5)||c===0xAD||(c>=0xB0&&c<=0xB3)||c===0xB9||(c>=0x300&&c<=0x36F)||(c>=0x483&&c<=0x489)||c===0x58F||(c>=0x591&&c<=0x5BD)||c===0x5BF||(c>=0x5C1&&c<=0x5C2)||(c>=0x5C4&&c<=0x5C5)||c===0x5C7||(c>=0x600&&c<=0x605)||(c>=0x609&&c<=0x60A)||c===0x60C||(c>=0x610&&c<=0x61A)||(c>=0x64B&&c<=0x66C)||c===0x670||(c>=0x6D6&&c<=0x6DD)||(c>=0x6DF&&c<=0x6E4)||(c>=0x6E7&&c<=0x6E8)||(c>=0x6EA&&c<=0x6ED)||(c>=0x6F0&&c<=0x6F9)||c===0x711||(c>=0x730&&c<=0x74A)||(c>=0x7A6&&c<=0x7B0)||(c>=0x7EB&&c<=0x7F3)||(c>=0x816&&c<=0x819)||(c>=0x81B&&c<=0x823)||(c>=0x825&&c<=0x827)||(c>=0x829&&c<=0x82D)||(c>=0x859&&c<=0x85B)||(c>=0x8E4&&c<=0x902)||c===0x93A||c===0x93C||(c>=0x941&&c<=0x948)||c===0x94D||(c>=0x951&&c<=0x957)||(c>=0x962&&c<=0x963)||c===0x981||c===0x9BC||(c>=0x9C1&&c<=0x9C4)||c===0x9CD||(c>=0x9E2&&c<=0x9E3)||(c>=0x9F2&&c<=0x9F3)||c===0x9FB||(c>=0xA01&&c<=0xA02)||c===0xA3C||(c>=0xA41&&c<=0xA42)||(c>=0xA47&&c<=0xA48)||(c>=0xA4B&&c<=0xA4D)||c===0xA51||(c>=0xA70&&c<=0xA71)||c===0xA75||(c>=0xA81&&c<=0xA82)||c===0xABC||(c>=0xAC1&&c<=0xAC5)||(c>=0xAC7&&c<=0xAC8)||c===0xACD||(c>=0xAE2&&c<=0xAE3)||c===0xAF1||c===0xB01||c===0xB3C||c===0xB3F||(c>=0xB41&&c<=0xB44)||c===0xB4D||c===0xB56||(c>=0xB62&&c<=0xB63)||c===0xB82||c===0xBC0||c===0xBCD||c===0xBF9||c===0xC00||(c>=0xC3E&&c<=0xC40)||(c>=0xC46&&c<=0xC48)||(c>=0xC4A&&c<=0xC4D)||(c>=0xC55&&c<=0xC56)||(c>=0xC62&&c<=0xC63)||c===0xC81||c===0xCBC||(c>=0xCCC&&c<=0xCCD)||(c>=0xCE2&&c<=0xCE3)||c===0xD01||(c>=0xD41&&c<=0xD44)||c===0xD4D||(c>=0xD62&&c<=0xD63)||c===0xDCA||(c>=0xDD2&&c<=0xDD4)||c===0xDD6||c===0xE31||(c>=0xE34&&c<=0xE3A)||c===0xE3F||(c>=0xE47&&c<=0xE4E)||c===0xEB1||(c>=0xEB4&&c<=0xEB9)||(c>=0xEBB&&c<=0xEBC)||(c>=0xEC8&&c<=0xECD)||(c>=0xF18&&c<=0xF19)||c===0xF35||c===0xF37||c===0xF39||(c>=0xF71&&c<=0xF7E)||(c>=0xF80&&c<=0xF84)||(c>=0xF86&&c<=0xF87)||(c>=0xF8D&&c<=0xF97)||(c>=0xF99&&c<=0xFBC)||c===0xFC6||(c>=0x102D&&c<=0x1030)||(c>=0x1032&&c<=0x1037)||(c>=0x1039&&c<=0x103A)||(c>=0x103D&&c<=0x103E)||(c>=0x1058&&c<=0x1059)||(c>=0x105E&&c<=0x1060)||(c>=0x1071&&c<=0x1074)||c===0x1082||(c>=0x1085&&c<=0x1086)||c===0x108D||c===0x109D||(c>=0x135D&&c<=0x135F)||(c>=0x1712&&c<=0x1714)||(c>=0x1732&&c<=0x1734)||(c>=0x1752&&c<=0x1753)||(c>=0x1772&&c<=0x1773)||(c>=0x17B4&&c<=0x17B5)||(c>=0x17B7&&c<=0x17BD)||c===0x17C6||(c>=0x17C9&&c<=0x17D3)||c===0x17DB||c===0x17DD||(c>=0x180B&&c<=0x180E)||c===0x18A9||(c>=0x1920&&c<=0x1922)||(c>=0x1927&&c<=0x1928)||c===0x1932||(c>=0x1939&&c<=0x193B)||(c>=0x1A17&&c<=0x1A18)||c===0x1A1B||c===0x1A56||(c>=0x1A58&&c<=0x1A5E)||c===0x1A60||c===0x1A62||(c>=0x1A65&&c<=0x1A6C)||(c>=0x1A73&&c<=0x1A7C)||c===0x1A7F||(c>=0x1AB0&&c<=0x1ABE)||(c>=0x1B00&&c<=0x1B03)||c===0x1B34||(c>=0x1B36&&c<=0x1B3A)||c===0x1B3C||c===0x1B42||(c>=0x1B6B&&c<=0x1B73)||(c>=0x1B80&&c<=0x1B81)||(c>=0x1BA2&&c<=0x1BA5)||(c>=0x1BA8&&c<=0x1BA9)||(c>=0x1BAB&&c<=0x1BAD)||c===0x1BE6||(c>=0x1BE8&&c<=0x1BE9)||c===0x1BED||(c>=0x1BEF&&c<=0x1BF1)||(c>=0x1C2C&&c<=0x1C33)||(c>=0x1C36&&c<=0x1C37)||(c>=0x1CD0&&c<=0x1CD2)||(c>=0x1CD4&&c<=0x1CE0)||(c>=0x1CE2&&c<=0x1CE8)||c===0x1CED||c===0x1CF4||(c>=0x1CF8&&c<=0x1CF9)||(c>=0x1DC0&&c<=0x1DF5)||(c>=0x1DFC&&c<=0x1DFF)||(c>=0x200B&&c<=0x200D)||(c>=0x202F&&c<=0x2034)||c===0x2044||(c>=0x2060&&c<=0x2064)||(c>=0x206A&&c<=0x2070)||(c>=0x2074&&c<=0x207B)||(c>=0x2080&&c<=0x208B)||(c>=0x20A0&&c<=0x20BD)||(c>=0x20D0&&c<=0x20F0)||c===0x212E||(c>=0x2212&&c<=0x2213)||(c>=0x2488&&c<=0x249B)||(c>=0x2CEF&&c<=0x2CF1)||c===0x2D7F||(c>=0x2DE0&&c<=0x2DFF)||(c>=0x302A&&c<=0x302D)||(c>=0x3099&&c<=0x309A)||(c>=0xA66F&&c<=0xA672)||(c>=0xA674&&c<=0xA67D)||c===0xA69F||(c>=0xA6F0&&c<=0xA6F1)||c===0xA802||c===0xA806||c===0xA80B||(c>=0xA825&&c<=0xA826)||(c>=0xA838&&c<=0xA839)||c===0xA8C4||(c>=0xA8E0&&c<=0xA8F1)||(c>=0xA926&&c<=0xA92D)||(c>=0xA947&&c<=0xA951)||(c>=0xA980&&c<=0xA982)||c===0xA9B3||(c>=0xA9B6&&c<=0xA9B9)||c===0xA9BC||c===0xA9E5||(c>=0xAA29&&c<=0xAA2E)||(c>=0xAA31&&c<=0xAA32)||(c>=0xAA35&&c<=0xAA36)||c===0xAA43||c===0xAA4C||c===0xAA7C||c===0xAAB0||(c>=0xAAB2&&c<=0xAAB4)||(c>=0xAAB7&&c<=0xAAB8)||(c>=0xAABE&&c<=0xAABF)||c===0xAAC1||(c>=0xAAEC&&c<=0xAAED)||c===0xAAF6||c===0xABE5||c===0xABE8||c===0xABED||c===0xFB1E||c===0xFB29||(c>=0xFE00&&c<=0xFE0F)||(c>=0xFE20&&c<=0xFE2D)||c===0xFE50||c===0xFE52||c===0xFE55||c===0xFE5F||(c>=0xFE62&&c<=0xFE63)||(c>=0xFE69&&c<=0xFE6A)||c===0xFEFF||(c>=0xFF03&&c<=0xFF05)||(c>=0xFF0B&&c<=0xFF1A)||(c>=0xFFE0&&c<=0xFFE1)||(c>=0xFFE5&&c<=0xFFE6)||c===0x101FD||(c>=0x102E0&&c<=0x102FB)||(c>=0x10376&&c<=0x1037A)||(c>=0x10A01&&c<=0x10A03)||(c>=0x10A05&&c<=0x10A06)||(c>=0x10A0C&&c<=0x10A0F)||(c>=0x10A38&&c<=0x10A3A)||c===0x10A3F||(c>=0x10AE5&&c<=0x10AE6)||(c>=0x10E60&&c<=0x10E7E)||c===0x11001||(c>=0x11038&&c<=0x11046)||(c>=0x1107F&&c<=0x11081)||(c>=0x110B3&&c<=0x110B6)||(c>=0x110B9&&c<=0x110BA)||(c>=0x11100&&c<=0x11102)||(c>=0x11127&&c<=0x1112B)||(c>=0x1112D&&c<=0x11134)||c===0x11173||(c>=0x11180&&c<=0x11181)||(c>=0x111B6&&c<=0x111BE)||(c>=0x1122F&&c<=0x11231)||c===0x11234||(c>=0x11236&&c<=0x11237)||c===0x112DF||(c>=0x112E3&&c<=0x112EA)||c===0x11301||c===0x1133C||c===0x11340||(c>=0x11366&&c<=0x1136C)||(c>=0x11370&&c<=0x11374)||(c>=0x114B3&&c<=0x114B8)||c===0x114BA||(c>=0x114BF&&c<=0x114C0)||(c>=0x114C2&&c<=0x114C3)||(c>=0x115B2&&c<=0x115B5)||(c>=0x115BC&&c<=0x115BD)||(c>=0x115BF&&c<=0x115C0)||(c>=0x11633&&c<=0x1163A)||c===0x1163D||(c>=0x1163F&&c<=0x11640)||c===0x116AB||c===0x116AD||(c>=0x116B0&&c<=0x116B5)||c===0x116B7||(c>=0x16AF0&&c<=0x16AF4)||(c>=0x16B30&&c<=0x16B36)||(c>=0x16F8F&&c<=0x16F92)||(c>=0x1BC9D&&c<=0x1BC9E)||(c>=0x1BCA0&&c<=0x1BCA3)||(c>=0x1D167&&c<=0x1D169)||(c>=0x1D173&&c<=0x1D182)||(c>=0x1D185&&c<=0x1D18B)||(c>=0x1D1AA&&c<=0x1D1AD)||(c>=0x1D242&&c<=0x1D244)||(c>=0x1D7CE&&c<=0x1D7FF)||(c>=0x1E8D0&&c<=0x1E8D6)||(c>=0x1F100&&c<=0x1F10A)||c===0xE0001||(c>=0xE0020&&c<=0xE007F)||(c>=0xE0100&&c<=0xE01EF); }", _
        "function IsExplicit(c) { return (c>=0x202A&&c<=0x202E)||(c>=0x2066&&c<=0x2069); }", _
        "function doFixDirection(sVal, direction) { var iCount, sOutVal = direction ? '\u200E' : '\u200F', bInv = false; for (iCount = 0; iCount < sVal.length; iCount++) { sOutVal += (sVal.charCodeAt(iCount) == 0x200E || sVal.charCodeAt(iCount) == 0x200F || IsExplicit(sVal.charCodeAt(iCount))) ? '' : (!bInv && IsNeutral(sVal.charCodeAt(iCount)) ? '' : (direction ? '\u202A' : '\u202B')) + sVal[iCount]; } return sOutVal; }", _
        "function doTransliterateDisplay() { $('#translitvalue').css('direction', $('#direction0').prop('checked') ? 'ltr' : 'rtl'); $('#translitvalue').text($('#scheme0').prop('checked') ? doTransliterate($('#translitedit').val(), $('#direction0').prop('checked'), parseInt($('#translitscheme').val(), 10)) : ($('#scheme1').prop('checked') ? doDiacritics($('#translitedit').val(), $('#diacriticscheme0').prop('checked')) : doFixDirection($('#translitedit').val(), $('#diacriticscheme0').prop('checked')))); }"}
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
        Dim Strings(CachedData.IslamData.TranslitSchemes.Length * 2 + 2 - 1) As Array
        Strings(0) = New String() {Utility.LoadResourceString("IslamSource_Off"), "0"}
        Strings(1) = New String() {Utility.LoadResourceString("IslamSource_ExtendedBuckwalter"), "1"}
        For Count = 0 To CachedData.IslamData.TranslitSchemes.Length - 1
            Strings(Count * 2 + 2) = New String() {Utility.LoadResourceString("IslamSource_" + CachedData.IslamData.TranslitSchemes(Count).Name), CStr(Count * 2 + 2)}
            Strings(Count * 2 + 1 + 2) = New String() {Utility.LoadResourceString("IslamSource_" + CachedData.IslamData.TranslitSchemes(Count).Name) + " Literal", CStr(Count * 2 + 1 + 2)}
        Next
        Return Strings
    End Function
    Public Shared Function GetChangeTransliterationJS() As String()
        Dim GetJS As New List(Of String) From {"javascript: changeTransliteration();", String.Empty, Utility.GetLookupStyleSheetJS(), GetArabicSymbolJSArray(), GetTranslitSchemeJSArray(), _
        "function changeTransliteration() { var k, child, iSubCount, text; $('span.transliteration').each(function() { $(this).css('display', $('#translitscheme').val() === '0' ? 'none' : 'block'); }); for (k in renderList) { text = ''; for (child in renderList[k]['children']) { for (iSubCount = 0; iSubCount < renderList[k]['children'][child]['arabic'].length; iSubCount++) { if ($('#translitscheme').val() !== '0' && $('#translitscheme').val() !== '1' && renderList[k]['children'][child]['arabic'][iSubCount] !== '' && renderList[k]['children'][child]['translit'][iSubCount] !== '') { if (text !== '') text += ' '; text += $('#' + renderList[k]['children'][child]['arabic'][iSubCount]).text(); } else { if (renderList[k]['children'][child]['translit'][iSubCount] !== '') $('#' + renderList[k]['children'][child]['translit'][iSubCount]).text(($('#translitscheme').val() === '0' || renderList[k]['children'][child]['arabic'][iSubCount] === '') ? '' : doTransliterate($('#' + renderList[k]['children'][child]['arabic'][iSubCount]).text(), true, parseInt($('#translitscheme').val(), 10))); } } } if ($('#translitscheme').val() !== '0' && $('#translitscheme').val() !== '1') { text = transliterateWithRules(text, Math.floor((parseInt($('#translitscheme').val(), 10) - 2) / 2) + 2).split(' '); for (child in renderList[k]['children']) { for (iSubCount = 0; iSubCount < renderList[k]['children'][child]['translit'].length; iSubCount++) { if (renderList[k]['children'][child]['translit'][iSubCount] !== '') $('#' + renderList[k]['children'][child]['translit'][iSubCount]).text(text.shift()); } } } for (iSubCount = 0; iSubCount < renderList[k]['arabic'].length; iSubCount++) { if (renderList[k]['translit'][iSubCount] !== '') $('#' + renderList[k]['translit'][iSubCount]).text(($('#translitscheme').val() === '0' || renderList[k]['arabic'][iSubCount] === '') ? '' : (($('#translitscheme').val() !== '0' && $('#translitscheme').val() !== '1') ? transliterateWithRules($('#' + renderList[k]['arabic'][iSubCount]).text(), parseInt($('#translitscheme').val(), 10) >= 2 ? Math.floor((parseInt($('#translitscheme').val(), 10) - 2) / 2) + 2 : parseInt($('#translitscheme').val(), 10)) : doTransliterate($('#' + renderList[k]['arabic'][iSubCount]).text(), true, parseInt($('#translitscheme').val(), 10)))); } } }"}
        GetJS.AddRange(PlainTransliterateGenJS)
        GetJS.AddRange(TransliterateGenJS)
        GetJS.AddRange(NumberGenJS)
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
        Dim SchemeType As ArabicData.TranslitScheme = CType(If(CInt(HttpContext.Current.Request.QueryString.Get("translitscheme")) >= 2, 2 - CInt(HttpContext.Current.Request.QueryString.Get("translitscheme")) Mod 2, CInt(HttpContext.Current.Request.QueryString.Get("translitscheme"))), ArabicData.TranslitScheme)
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
    Public Shared Function DoDisplayTranslation(Category As Object, SchemeType As ArabicData.TranslitScheme, Scheme As String) As Array()
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
                Output.Add(New String() {ArabicData.TransliterateFromBuckwalter(CType(Category, IslamData.Month())(SubCount).Name), TransliterateToScheme(ArabicData.TransliterateFromBuckwalter(CType(Category, IslamData.Month())(SubCount).Name), SchemeType, Scheme).Trim(), Utility.LoadResourceString("IslamInfo_" + CType(Category, IslamData.Month())(SubCount).TranslationID)})
            Next
        ElseIf TypeOf Category Is IslamData.DayOfWeek Then
            For SubCount As Integer = 0 To CType(Category, IslamData.DayOfWeek()).Length - 1
                Output.Add(New String() {ArabicData.TransliterateFromBuckwalter(CType(Category, IslamData.DayOfWeek())(SubCount).Name), TransliterateToScheme(ArabicData.TransliterateFromBuckwalter(CType(Category, IslamData.DayOfWeek())(SubCount).Name), SchemeType, Scheme).Trim(), Utility.LoadResourceString("IslamInfo_" + CType(Category, IslamData.DayOfWeek())(SubCount).TranslationID)})
            Next
        ElseIf TypeOf Category Is IslamData.PrayerTime Then
            Dim Table As New Hashtable
            Array.ForEach(CachedData.IslamData.Prayers, Sub(Convert As IslamData.PrayerType) Array.ForEach(Convert.PrayerUnits.Split(","c), Sub(Part As String) If Part.Contains("="c) Then If Table.ContainsKey(Part.Substring(0, Part.IndexOf("="c))) Then Table.Item(Part.Substring(0, Part.IndexOf("="c))) = CStr(Table.Item(Part.Substring(0, Part.IndexOf("="c)))) + vbCrLf + Part.Substring(Part.IndexOf("="c) + 1).Replace("|"c, " or ") + " " + CStr(IIf(Utility.LoadResourceString("IslamInfo_" + Convert.TranslationID) <> "Prescribed time", Utility.LoadResourceString("IslamInfo_" + Convert.TranslationID) + " - ", String.Empty)) + Convert.Classification Else Table.Add(Part.Substring(0, Part.IndexOf("="c)), Part.Substring(Part.IndexOf("="c) + 1).Replace("|"c, " or ") + " " + Convert.Classification)))
            For SubCount As Integer = 0 To CType(Category, IslamData.PrayerTime()).Length - 1
                Output.Add(New String() {ArabicData.TransliterateFromBuckwalter(CType(Category, IslamData.PrayerTime())(SubCount).Name), TransliterateToScheme(ArabicData.TransliterateFromBuckwalter(CType(Category, IslamData.PrayerTime())(SubCount).Name), SchemeType, Scheme).Trim(), Utility.LoadResourceString("IslamInfo_" + CType(Category, IslamData.PrayerTime())(SubCount).TranslationID), CStr(IIf(Table.ContainsKey("-"c + Utility.LoadResourceString("IslamInfo_" + CType(Category, IslamData.PrayerTime())(SubCount).TranslationID)), Table.Item("-"c + Utility.LoadResourceString("IslamInfo_" + CType(Category, IslamData.PrayerTime())(SubCount).TranslationID)), String.Empty)), CStr(IIf(Table.ContainsKey(Utility.LoadResourceString("IslamInfo_" + CType(Category, IslamData.PrayerTime())(SubCount).TranslationID)), Table.Item(Utility.LoadResourceString("IslamInfo_" + CType(Category, IslamData.PrayerTime())(SubCount).TranslationID)), String.Empty)), CStr(IIf(Table.ContainsKey("+"c + Utility.LoadResourceString("IslamInfo_" + CType(Category, IslamData.PrayerTime())(SubCount).TranslationID)), Table.Item("+"c + Utility.LoadResourceString("IslamInfo_" + CType(Category, IslamData.PrayerTime())(SubCount).TranslationID)), String.Empty))})
            Next
        ElseIf TypeOf Category Is IslamData.PrayerType Then
            For SubCount As Integer = 0 To CType(Category, IslamData.PrayerType()).Length - 1
                Output.Add(New String() {ArabicData.TransliterateFromBuckwalter(CType(Category, IslamData.PrayerType())(SubCount).Name), TransliterateToScheme(ArabicData.TransliterateFromBuckwalter(CType(Category, IslamData.PrayerType())(SubCount).Name), SchemeType, Scheme).Trim(), Utility.LoadResourceString("IslamInfo_" + CType(Category, IslamData.PrayerType())(SubCount).TranslationID), CType(Category, IslamData.PrayerType())(SubCount).Classification, CType(Category, IslamData.PrayerType())(SubCount).PrayerUnits})
            Next
        Else
            For SubCount As Integer = 0 To CType(Category, IslamData.VocabCategory).Words.Length - 1
                Output.Add(New String() {ArabicData.TransliterateFromBuckwalter(CType(Category, IslamData.VocabCategory).Words(SubCount).Text), TransliterateToScheme(ArabicData.TransliterateFromBuckwalter(CType(Category, IslamData.VocabCategory).Words(SubCount).Text), SchemeType, Scheme).Trim(), Utility.LoadResourceString("IslamInfo_" + CType(Category, IslamData.VocabCategory).Words(SubCount).TranslationID)})
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
        Dim Output(ArabicData.Data.ArabicCombos.Length + 2) As Array
        Output(0) = New String() {}
        Output(1) = New String() {"arabic", "arabic", String.Empty, String.Empty}
        Output(2) = New String() {Utility.LoadResourceString("IslamInfo_LetterName"), Utility.LoadResourceString("IslamInfo_Arabic"), Utility.LoadResourceString("IslamSource_ExtendedBuckwalter"), Utility.LoadResourceString("IslamInfo_Shaping")}
        'Dim Combos(ArabicData.Data.ArabicCombos.Length - 1) As IslamData.ArabicCombo
        'ArabicData.Data.ArabicLetters.CopyTo(ArabicData.Data.ArabicCombos, 0)
        'Array.Sort(Combos, Function(Key As IslamData.ArabicCombo, NextKey As IslamData.ArabicCombo) Key.SymbolName.CompareTo(NextKey.SymbolName))
        For Count = 0 To ArabicData.Data.ArabicCombos.Length - 1
            Output(Count + 3) = New String() {String.Join(" ", Array.ConvertAll(ArabicData.TransliterateFromBuckwalter(ArabicData.Data.ArabicCombos(Count).SymbolName).ToCharArray(), Function(Ch As Char) ArabicData.TransliterateFromBuckwalter(ArabicData.Data.ArabicLetters(ArabicData.FindLetterBySymbol(Ch)).SymbolName))), _
                                       ArabicData.TransliterateFromBuckwalter(ArabicData.Data.ArabicCombos(Count).SymbolName), _
                                       ArabicData.Data.ArabicCombos(Count).SymbolName,
                                       String.Join(vbCrLf, Array.ConvertAll(ArabicData.Data.ArabicCombos(Count).Shaping, Function(Shape As Char) If(Shape = ChrW(0), String.Empty, Shape + " " + CStr(Hex(AscW(Shape))) + " " + If(CheckShapingOrder(Array.IndexOf(ArabicData.Data.ArabicCombos(Count).Shaping, Shape), ArabicData.GetUnicodeName(Shape)), String.Empty, "!!!") + ArabicData.GetUnicodeName(Shape))))}
        Next
        Return Output
    End Function
    Public Shared Function SymbolDisplay(Symbols() As ArabicData.ArabicXMLData.ArabicSymbol) As Array()
        Dim Count As Integer
        Dim Output(Symbols.Length + 2) As Array
        'Dim oFont As New Font(DefaultValue(HttpContext.Current.Request.QueryString.Get("fontcustom"), "Arial"), 13)
        'CheckIfCharInFont(ArabicData.Data.ArabicLetters(Count).Symbol, oFont)
        Output(0) = New String() {}
        Output(1) = New String() {"arabic", String.Empty, "arabic", String.Empty, String.Empty, String.Empty, String.Empty, String.Empty, String.Empty}
        Output(2) = New String() {Utility.LoadResourceString("IslamInfo_LetterName"), Utility.LoadResourceString("IslamInfo_UnicodeName"), Utility.LoadResourceString("IslamInfo_Arabic"), Utility.LoadResourceString("IslamInfo_UnicodeValue"), Utility.LoadResourceString("IslamSource_ExtendedBuckwalter"), Utility.LoadResourceString("IslamInfo_Terminating"), Utility.LoadResourceString("IslamInfo_Connecting"), Utility.LoadResourceString("IslamInfo_Assimilate"), Utility.LoadResourceString("IslamInfo_Shaping")}
        For Count = 0 To Symbols.Length - 1
            Output(Count + 3) = New String() {ArabicData.TransliterateFromBuckwalter(Symbols(Count).SymbolName), _
                                              ArabicData.GetUnicodeName(Symbols(Count).Symbol), _
                                       CStr(Symbols(Count).Symbol), _
                                       CStr(Hex(AscW(Symbols(Count).Symbol))), _
                                       CStr(IIf(Symbols(Count).ExtendedBuckwalterLetter = ChrW(0), String.Empty, Symbols(Count).ExtendedBuckwalterLetter)), _
                                       CStr(Symbols(Count).Terminating), _
                                       CStr(Symbols(Count).Connecting), _
                                       CStr(Symbols(Count).Assimilate),
                                       If(Symbols(Count).Shaping = Nothing, String.Empty, String.Join(vbCrLf, Array.ConvertAll(Symbols(Count).Shaping, Function(Shape As Char) If(Shape = ChrW(0), String.Empty, Shape + " " + CStr(Hex(AscW(Shape))) + " " + If(CheckShapingOrder(Array.IndexOf(Symbols(Count).Shaping, Shape), ArabicData.GetUnicodeName(Shape)), String.Empty, "!!!") + ArabicData.GetUnicodeName(Shape)))))}
        Next
        Return Output
    End Function
    Public Shared Function DisplayAll(ByVal Item As PageLoader.TextItem) As Array()
        Return SymbolDisplay(ArabicData.Data.ArabicLetters)
    End Function
    Public Shared Function DisplayTranslitSchemes(ByVal Item As PageLoader.TextItem) As Array()
        Dim Count As Integer
        Dim Output(ArabicData.Data.ArabicLetters.Length + ArabicData.ArabicLongVowels.Length + 2) As Array
        'Dim oFont As New Font(DefaultValue(HttpContext.Current.Request.QueryString.Get("fontcustom"), "Arial"), 13)
        'CheckIfCharInFont(ArabicData.Data.ArabicLetters(Count).Symbol, oFont)
        Output(0) = New String() {}
        Dim Strs As String() = New String() {"arabic", String.Empty, "arabic", String.Empty}
        Array.Resize(Of String)(Strs, 4 + CachedData.IslamData.TranslitSchemes.Length)
        Output(1) = Strs
        Strs = New String() {Utility.LoadResourceString("IslamInfo_LetterName"), Utility.LoadResourceString("IslamInfo_UnicodeName"), Utility.LoadResourceString("IslamInfo_Arabic"), Utility.LoadResourceString("IslamSource_ExtendedBuckwalter")}
        Array.Resize(Of String)(Strs, 4 + CachedData.IslamData.TranslitSchemes.Length)
        For SchemeCount = 0 To CachedData.IslamData.TranslitSchemes.Length - 1
            CType(Output(1), String())(4 + SchemeCount) = String.Empty
            Strs(4 + SchemeCount) = Utility.LoadResourceString("IslamSource_" + CachedData.IslamData.TranslitSchemes(SchemeCount).Name)
        Next
        Output(2) = Strs
        For Count = 0 To ArabicData.Data.ArabicLetters.Length - 1
            Strs = New String() {ArabicData.TransliterateFromBuckwalter(ArabicData.Data.ArabicLetters(Count).SymbolName), _
                                              ArabicData.GetUnicodeName(ArabicData.Data.ArabicLetters(Count).Symbol), _
                                       CStr(ArabicData.Data.ArabicLetters(Count).Symbol), _
                                       CStr(IIf(ArabicData.Data.ArabicLetters(Count).ExtendedBuckwalterLetter = ChrW(0), String.Empty, ArabicData.Data.ArabicLetters(Count).ExtendedBuckwalterLetter))}
            Array.Resize(Of String)(Strs, 4 + CachedData.IslamData.TranslitSchemes.Length)
            For SchemeCount = 0 To CachedData.IslamData.TranslitSchemes.Length - 1
                Strs(4 + SchemeCount) = GetSchemeValueFromSymbol(ArabicData.Data.ArabicLetters(Count), CachedData.IslamData.TranslitSchemes(SchemeCount).Name)
            Next
            Output(Count + 3) = Strs
        Next
        For Count = 0 To ArabicData.ArabicLongVowels.Length - 1
            Strs = New String() {String.Join(" ", Array.ConvertAll(ArabicData.ArabicLongVowels(Count).ToCharArray(), Function(Ch As Char) ArabicData.TransliterateFromBuckwalter(ArabicData.Data.ArabicLetters(ArabicData.FindLetterBySymbol(Ch)).SymbolName))), _
                                              String.Empty, _
                                       ArabicData.ArabicLongVowels(Count), _
                                       TransliterateToScheme(ArabicData.ArabicLongVowels(Count), ArabicData.TranslitScheme.Literal, String.Empty)}
            Array.Resize(Of String)(Strs, 4 + CachedData.IslamData.TranslitSchemes.Length)
            For SchemeCount = 0 To CachedData.IslamData.TranslitSchemes.Length - 1
                Strs(4 + SchemeCount) = GetSchemeLongVowelFromString(ArabicData.ArabicLongVowels(Count), CachedData.IslamData.TranslitSchemes(SchemeCount).Name)
            Next
            Output(ArabicData.Data.ArabicLetters.Length + Count + 3) = Strs
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
            Output(3 + Count) = New String() {ArabicData.TransliterateFromBuckwalter(Category.Words(Count).Text), Category.Words(Count).TranslationID, Utility.DefaultValue(Category.Words(Count).Grammar, String.Empty)}
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
                                  Build.Item(Key).Item(Str.Chars(If(Personal, 2, 1))) += " " + ArabicData.TransliterateFromBuckwalter(Category.Words(Count).Text)
                              Else
                                  Build.Item(Key).Add(Str.Chars(If(Personal, 2, 1)), ArabicData.TransliterateFromBuckwalter(Category.Words(Count).Text))
                              End If
                          End Sub)
        Next
        If Personal Then
            If Build.ContainsKey("3m") Then
                If Not Build("3m").ContainsKey("p") Then Build("3m").Add("p", String.Empty)
                If Not Build("3m").ContainsKey("d") Then Build("3m").Add("d", String.Empty)
                If Not Build("3m").ContainsKey("s") Then Build("3m").Add("s", String.Empty)
                Output(3) = New String() {Build("3m")("p"), Build("3m")("d"), Build("3m")("s"), "Third Person Masculine"}
            End If
            If Build.ContainsKey("3f") Then
                If Not Build("3f").ContainsKey("p") Then Build("3f").Add("p", String.Empty)
                If Not Build("3f").ContainsKey("d") Then Build("3f").Add("d", String.Empty)
                If Not Build("3f").ContainsKey("s") Then Build("3f").Add("s", String.Empty)
                Output(4) = New String() {Build("3f")("p"), Build("3f")("d"), Build("3f")("s"), "Third Person Feminine"}
            End If
            If Build.ContainsKey("2m") Then
                If Not Build("2m").ContainsKey("p") Then Build("2m").Add("p", String.Empty)
                If Not Build("2m").ContainsKey("d") Then Build("2m").Add("d", String.Empty)
                If Not Build("2m").ContainsKey("s") Then Build("2m").Add("s", String.Empty)
                Output(5) = New String() {Build("2m")("p"), Build("2m")("d"), Build("2m")("s"), "Second Person Masculine"}
            End If
            If Build.ContainsKey("2f") Then
                If Not Build("2f").ContainsKey("p") Then Build("2f").Add("p", String.Empty)
                If Not Build("2f").ContainsKey("d") Then Build("2f").Add("d", String.Empty)
                If Not Build("2f").ContainsKey("s") Then Build("2f").Add("s", String.Empty)
                Output(6) = New String() {Build("2f")("p"), Build("2f")("d"), Build("2f")("s"), "Second Person Feminine"}
            End If
            If Build.ContainsKey("1m") Then
                If Not Build("1m").ContainsKey("p") Then Build("1m").Add("p", String.Empty)
                If Not Build("1m").ContainsKey("d") Then Build("1m").Add("d", String.Empty)
                If Not Build("1m").ContainsKey("s") Then Build("1m").Add("s", String.Empty)
                Output(7) = New String() {Build("1m")("p"), Build("1m")("d"), Build("1m")("s"), "First Person Masculine"}
            End If
            If Build.ContainsKey("1f") Then
                If Not Build("1f").ContainsKey("p") Then Build("1f").Add("p", String.Empty)
                If Not Build("1f").ContainsKey("d") Then Build("1f").Add("d", String.Empty)
                If Not Build("1f").ContainsKey("s") Then Build("1f").Add("s", String.Empty)
                Output(8) = New String() {Build("1f")("p"), Build("1f")("d"), Build("1f")("s"), "First Person Feminine"}
            End If
        Else
            If Build.ContainsKey("m") Then
                If Not Build("m").ContainsKey("p") Then Build("m").Add("p", String.Empty)
                If Not Build("m").ContainsKey("d") Then Build("m").Add("d", String.Empty)
                If Not Build("m").ContainsKey("s") Then Build("m").Add("s", String.Empty)
                Output(3) = New String() {Build("m")("p"), Build("m")("d"), Build("m")("s"), "Masculine"}
            End If
            If Build.ContainsKey("f") Then
                If Not Build("f").ContainsKey("p") Then Build("f").Add("p", String.Empty)
                If Not Build("f").ContainsKey("d") Then Build("f").Add("d", String.Empty)
                If Not Build("f").ContainsKey("s") Then Build("f").Add("s", String.Empty)
                Output(4) = New String() {Build("f")("p"), Build("f")("d"), Build("f")("s"), "Feminine"}
            End If
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
            Output(3 + Count) = New String() {ArabicData.TransliterateFromBuckwalter(Category.Words(Count).Text), String.Empty, String.Empty, Category.Words(Count).TranslationID}
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
                Text = ArabicData.TransliterateFromBuckwalter(Category.Words(Count).Grammar.Substring(5).Split(","c)(0).Replace("f", Category.Words(Count).Text.Chars(0)).Replace("E", Category.Words(Count).Text.Chars(1)).Replace("l", Category.Words(Count).Text.Chars(2)))
                Present = ArabicData.TransliterateFromBuckwalter(Category.Words(Count).Grammar.Substring(5).Split(","c)(1).Replace("f", Category.Words(Count).Text.Chars(0)).Replace("E", Category.Words(Count).Text.Chars(1)).Replace("l", Category.Words(Count).Text.Chars(2)))
                Command = ArabicData.TransliterateFromBuckwalter("{foE$1lo".Replace("f", Category.Words(Count).Text.Chars(0)).Replace("E", Category.Words(Count).Text.Chars(1)).Replace("l", Category.Words(Count).Text.Chars(2)).Replace("$1", Category.Words(Count).Grammar.Substring(5).Split(","c)(1).Chars(5)))
                Forbidding = ArabicData.TransliterateFromBuckwalter("laA tafoE$1lo".Replace("f", Category.Words(Count).Text.Chars(0)).Replace("E", Category.Words(Count).Text.Chars(1)).Replace("l", Category.Words(Count).Text.Chars(2)).Replace("$1", Category.Words(Count).Grammar.Substring(5).Split(","c)(1).Chars(5)))
                PassivePast = ArabicData.TransliterateFromBuckwalter("fuEila".Replace("f", Category.Words(Count).Text.Chars(0)).Replace("E", Category.Words(Count).Text.Chars(1)).Replace("l", Category.Words(Count).Text.Chars(2)))
                PassivePresent = ArabicData.TransliterateFromBuckwalter("yufoEalu".Replace("f", Category.Words(Count).Text.Chars(0)).Replace("E", Category.Words(Count).Text.Chars(1)).Replace("l", Category.Words(Count).Text.Chars(2)))
                VerbalDoer = ArabicData.TransliterateFromBuckwalter("faAEilN".Replace("f", Category.Words(Count).Text.Chars(0)).Replace("E", Category.Words(Count).Text.Chars(1)).Replace("l", Category.Words(Count).Text.Chars(2)))
                PassiveNoun = ArabicData.TransliterateFromBuckwalter("mafoEuwlN".Replace("f", Category.Words(Count).Text.Chars(0)).Replace("E", Category.Words(Count).Text.Chars(1)).Replace("l", Category.Words(Count).Text.Chars(2)))
                '"faEiylN" passive noun
                '"mafoEilN" time
                '"mafoEalN" place
                '"faEiylN" derived adjective "faEalN" "faEulaAnu" "&gt;ufoEalN" "faEolN" "fuEaAlN" "faEuwlN"
                '"&gt;afoEalu" "fuEolaY" comparative derived noun
                '"faE~aAlN" "mifoEaAlN" "faEuwlN" "fuE~uwl" "faEilN" "fiE~iylN" "fuEalapN" intensive derived noun
                '"mifoEalapN" "mifoEalN" "mifoEaAlN" instrument of action
                Grammar = String.Empty
            Else
                Text = ArabicData.TransliterateFromBuckwalter(Category.Words(Count).Text)
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
        Return "function embedFontStyle(fontID) { if (isInArray(embeddedFonts, fontID)) return; embeddedFonts.push(fontID); var font=fontList[fontID]; var style = 'font-family: \'' + font.embed + '\';' + 'src: url(\'/files/\' + font.file + \'.eot\');' + 'src: local(\'' + font.family + '\'), url(\'/files/' + font.file + ((font.file == 'KFC_naskh') ? '.otf\') format(\'opentype\');' : '.ttf\') format(\'truetype\');'); addStyleSheetRule(newStyleSheet(), '@font-face', style);  }"
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
        "function decreaseFontSize() { rule = findStyleSheetRule('span.arabic'); rule.style.fontSize = Math.max(parseInt(rule.style.fontSize.replace('px', ''), 10) - 1, 1) + 'px'; }"}
    End Function
    Public Shared Function GetFontDefaultSizeJS() As String()
        Return New String() {"javascript: defaultFontSize();", String.Empty, Utility.GetLookupStyleSheetJS(), _
        "function defaultFontSize() { findStyleSheetRule('span.arabic').style.fontSize = '32px'; }"}
    End Function
    Public Shared Function GetFontBiggerJS() As String()
        Return New String() {"javascript: increaseFontSize();", String.Empty, Utility.GetLookupStyleSheetJS(), _
        "function increaseFontSize() { rule = findStyleSheetRule('span.arabic'); rule.style.fontSize = (parseInt(rule.style.fontSize.replace('px', ''), 10) + 1) + 'px'; }"}
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
            Arabic.DoErrorCheck(ArabicData.TransliterateFromBuckwalter(IslamData.Months(Count).Name))
            Utility.LoadResourceString("IslamInfo_" + IslamData.Months(Count).TranslationID)
        Next
        For Count = 0 To IslamData.DaysOfWeek.Length - 1
            Arabic.DoErrorCheck(ArabicData.TransliterateFromBuckwalter(IslamData.DaysOfWeek(Count).Name))
            Utility.LoadResourceString("IslamInfo_" + IslamData.DaysOfWeek(Count).TranslationID)
        Next
        For Count = 0 To IslamData.Prayers.Length - 1
            Arabic.DoErrorCheck(ArabicData.TransliterateFromBuckwalter(IslamData.Prayers(Count).Name))
            Utility.LoadResourceString("IslamInfo_" + IslamData.Prayers(Count).TranslationID)
        Next
        For Count = 0 To IslamData.PrayerTimes.Length - 1
            Arabic.DoErrorCheck(ArabicData.TransliterateFromBuckwalter(IslamData.PrayerTimes(Count).Name))
            Utility.LoadResourceString("IslamInfo_" + IslamData.PrayerTimes(Count).TranslationID)
        Next
        'must check Trans and WordForWord
        For Count = 0 To IslamData.VerseCategories.Length - 1
            Utility.LoadResourceString("IslamInfo_" + IslamData.VerseCategories(Count).Title)
            For SubCount As Integer = 0 To IslamData.VerseCategories(Count).Verses.Length - 1
                DocBuilder.DoErrorCheckBuckwalterText(IslamData.VerseCategories(Count).Verses(SubCount).Arabic)
                Utility.LoadResourceString("IslamInfo_" + IslamData.VerseCategories(Count).Verses(SubCount).TranslationID)
            Next
        Next
        For Count = 0 To IslamData.GrammarCategories.Length - 1
            Utility.LoadResourceString("IslamInfo_" + IslamData.GrammarCategories(Count).Title)
            For SubCount As Integer = 0 To IslamData.GrammarCategories(Count).Words.Length - 1
                Arabic.DoErrorCheck(ArabicData.TransliterateFromBuckwalter(IslamData.GrammarCategories(Count).Words(SubCount).Text))
                Utility.LoadResourceString("IslamInfo_" + IslamData.GrammarCategories(Count).Words(SubCount).TranslationID)
            Next
        Next
        For Count = 0 To IslamData.VocabularyCategories.Length - 1
            Utility.LoadResourceString("IslamInfo_" + IslamData.VocabularyCategories(Count).Title)
            For SubCount As Integer = 0 To IslamData.VocabularyCategories(Count).Words.Length - 1
                Arabic.DoErrorCheck(ArabicData.TransliterateFromBuckwalter(IslamData.VocabularyCategories(Count).Words(SubCount).Text))
                Utility.LoadResourceString("IslamInfo_" + IslamData.VocabularyCategories(Count).Words(SubCount).TranslationID)
            Next
        Next
        For Count = 0 To ArabicData.Data.ArabicLetters.Length - 1
            Arabic.DoErrorCheck(ArabicData.TransliterateFromBuckwalter(ArabicData.Data.ArabicLetters(Count).SymbolName))
            Utility.LoadResourceString("IslamInfo_" + ArabicData.Data.ArabicLetters(Count).UnicodeName)
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
            Arabic.DoErrorCheck(ArabicData.TransliterateFromBuckwalter(IslamData.QuranChapters(Count).Name))
        Next
        For Count = 0 To IslamData.QuranParts.Length - 1
            Arabic.DoErrorCheck(ArabicData.TransliterateFromBuckwalter(IslamData.QuranParts(Count).Name))
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
    Public Shared Function GetListRenderedText(ByVal Item As PageLoader.TextItem) As RenderArray
        Dim SchemeType As ArabicData.TranslitScheme = CType(If(CInt(HttpContext.Current.Request.QueryString.Get("translitscheme")) >= 2, 2 - CInt(HttpContext.Current.Request.QueryString.Get("translitscheme")) Mod 2, CInt(HttpContext.Current.Request.QueryString.Get("translitscheme"))), ArabicData.TranslitScheme)
        Dim Scheme As String = If(CInt(HttpContext.Current.Request.QueryString.Get("translitscheme")) >= 2, CachedData.IslamData.TranslitSchemes((CInt(HttpContext.Current.Request.QueryString.Get("translitscheme")) - 2) \ 2).Name, String.Empty)
        Dim Count As Integer = CInt(HttpContext.Current.Request.QueryString.Get("selection"))
        If Count = -1 Then Count = 0
        Return BuckwalterTextFromReferences(SchemeType, Scheme, CachedData.IslamData.Lists(0).Words(Count).Text, String.Empty, TanzilReader.GetTranslationIndex(HttpContext.Current.Request.QueryString.Get("qurantranslation")))
    End Function
    Public Shared Function GetListCategories() As String()
        Return Array.ConvertAll(CachedData.IslamData.Lists(0).Words, Function(Convert As IslamData.VocabCategory.Word) Utility.LoadResourceString("IslamInfo_" + Convert.TranslationID))
    End Function
    Public Shared Function GetRenderedText(ByVal Item As PageLoader.TextItem) As RenderArray
        Dim SchemeType As ArabicData.TranslitScheme = CType(If(CInt(HttpContext.Current.Request.QueryString.Get("translitscheme")) >= 2, 2 - CInt(HttpContext.Current.Request.QueryString.Get("translitscheme")) Mod 2, CInt(HttpContext.Current.Request.QueryString.Get("translitscheme"))), ArabicData.TranslitScheme)
        Dim Scheme As String = If(CInt(HttpContext.Current.Request.QueryString.Get("translitscheme")) >= 2, CachedData.IslamData.TranslitSchemes((CInt(HttpContext.Current.Request.QueryString.Get("translitscheme")) - 2) \ 2).Name, String.Empty)
        Return NormalTextFromReferences(HttpContext.Current.Request.QueryString.Get("docedit"), SchemeType, Scheme, TanzilReader.GetTranslationIndex(HttpContext.Current.Request.QueryString.Get("qurantranslation")))
    End Function
    Public Shared Sub DoErrorCheckBuckwalterText(Strings As String)
        If Strings = Nothing Then Return
        Dim Matches As System.Text.RegularExpressions.MatchCollection = System.Text.RegularExpressions.Regex.Matches(Strings, "(.*?)(?:(\\\{)(.*?)(\\\})|$)")
        For MatchCount As Integer = 0 To Matches.Count - 1
            If Matches(MatchCount).Length <> 0 Then
                If Matches(MatchCount).Groups(1).Length <> 0 Then
                    Arabic.DoErrorCheck(ArabicData.TransliterateFromBuckwalter(Matches(MatchCount).Groups(1).Value))
                End If
                If Matches(MatchCount).Groups(3).Length <> 0 Then
                End If
            End If
        Next
    End Sub
    Public Shared Function BuckwalterTextFromReferences(SchemeType As ArabicData.TranslitScheme, Scheme As String, Strings As String, TranslationID As String, TranslationIndex As Integer) As RenderArray
        Dim Renderer As New RenderArray
        If Strings = Nothing Then Return Renderer
        Dim Matches As System.Text.RegularExpressions.MatchCollection = System.Text.RegularExpressions.Regex.Matches(Strings, "(.*?)(?:(\\\{)(.*?)(\\\})|$)")
        For MatchCount As Integer = 0 To Matches.Count - 1
            If Matches(MatchCount).Length <> 0 Then
                If Matches(MatchCount).Groups(1).Length <> 0 Then
                    Dim EnglishByWord As String() = Utility.LoadResourceString("IslamInfo_" + TranslationID + "WordByWord").Split("|"c)
                    Dim ArabicText As String() = Matches(MatchCount).Groups(1).Value.Split(" "c)
                    Dim Transliteration As String() = Arabic.TransliterateToScheme(ArabicData.TransliterateFromBuckwalter(Matches(MatchCount).Groups(1).Value), SchemeType, Scheme).Split(" "c)
                    Renderer.Items.Add(New RenderArray.RenderItem(RenderArray.RenderTypes.eHeaderCenter, New RenderArray.RenderText() {New RenderArray.RenderText(RenderArray.RenderDisplayClass.eLTR, Utility.LoadResourceString("IslamInfo_" + TranslationID))}))
                    Dim Items As New Collections.Generic.List(Of RenderArray.RenderItem)
                    For WordCount As Integer = 0 To EnglishByWord.Length - 1
                        Items.Add(New RenderArray.RenderItem(RenderArray.RenderTypes.eText, New RenderArray.RenderText() {New RenderArray.RenderText(RenderArray.RenderDisplayClass.eArabic, ArabicData.TransliterateFromBuckwalter(ArabicText(WordCount))), New RenderArray.RenderText(RenderArray.RenderDisplayClass.eTransliteration, Transliteration(WordCount)), New RenderArray.RenderText(RenderArray.RenderDisplayClass.eLTR, EnglishByWord(WordCount))}))
                    Next
                    Renderer.Items.Add(New RenderArray.RenderItem(RenderArray.RenderTypes.eText, New RenderArray.RenderText() {New RenderArray.RenderText(RenderArray.RenderDisplayClass.eNested, Items), New RenderArray.RenderText(RenderArray.RenderDisplayClass.eArabic, ArabicData.TransliterateFromBuckwalter(Matches(MatchCount).Groups(1).Value)), New RenderArray.RenderText(RenderArray.RenderDisplayClass.eTransliteration, Arabic.TransliterateToScheme(ArabicData.TransliterateFromBuckwalter(Matches(MatchCount).Groups(1).Value), SchemeType, Scheme)), New RenderArray.RenderText(RenderArray.RenderDisplayClass.eLTR, Utility.LoadResourceString("IslamInfo_" + TranslationID + "Trans"))}))
                End If
                If Matches(MatchCount).Groups(3).Length <> 0 Then
                    Renderer.Items.AddRange(TextFromReferences(Matches(MatchCount).Groups(3).Value, SchemeType, Scheme, TranslationIndex).Items)
                End If
            End If
        Next
        Return Renderer
    End Function
    Public Shared Function NormalTextFromReferences(Strings As String, SchemeType As ArabicData.TranslitScheme, Scheme As String, TranslationIndex As Integer) As RenderArray
        Dim Renderer As New RenderArray
        If Strings = Nothing Then Return Renderer
        Dim Matches As System.Text.RegularExpressions.MatchCollection = System.Text.RegularExpressions.Regex.Matches(Strings, "(.*?)(?:(\{)(.*?)(\})|$)")
        For Count As Integer = 0 To Matches.Count - 1
            If Matches(Count).Length <> 0 Then
                If Matches(Count).Groups(1).Length <> 0 Then

                End If
                If Matches(Count).Groups(3).Length <> 0 Then
                    Renderer.Items.AddRange(TextFromReferences(Matches(Count).Groups(3).Value, SchemeType, Scheme, TranslationIndex).Items)
                End If
            End If
        Next
        Return Renderer
    End Function
    Public Shared Function TextFromReferences(Strings As String, SchemeType As ArabicData.TranslitScheme, Scheme As String, TranslationIndex As Integer) As RenderArray
        Dim Renderer As New RenderArray
        If Strings = Nothing Then Return Renderer
        'text before and after reference matches needs rendering
        'hadith reference matching {name,book/hadith}
        If TanzilReader.IsQuranTextReference(Strings) Then
            Renderer.Items.AddRange(TanzilReader.QuranTextFromReference(Strings, SchemeType, Scheme, TranslationIndex).Items)
        ElseIf Strings.StartsWith("symbol:") Then
            Dim Symbols As New List(Of ArabicData.ArabicXMLData.ArabicSymbol)
            Dim SelArr As String() = Strings.Replace("symbol:", String.Empty).Split(","c)
            For SubCount = 0 To ArabicData.Data.ArabicLetters.Length - 1
                If Array.IndexOf(SelArr, ArabicData.Data.ArabicLetters(SubCount).UnicodeName.Replace("ArabicLetter", String.Empty).Replace("Arabic", String.Empty)) <> -1 Then
                    Symbols.Add(ArabicData.Data.ArabicLetters(SubCount))
                End If
            Next
            Renderer.Items.Add(New RenderArray.RenderItem(RenderArray.RenderTypes.eText, New RenderArray.RenderText() {New RenderArray.RenderText(RenderArray.RenderDisplayClass.eList, Arabic.SymbolDisplay(Symbols.ToArray()))}))
        ElseIf Strings.StartsWith("personalpronoun:") Or Strings.StartsWith("possessivedeterminerpersonalpronoun:") Then
            Dim GrammarCat As New IslamData.GrammarCategory
            Dim Cat As IslamData.GrammarCategory
            Dim Words As New List(Of IslamData.GrammarCategory.GrammarWord)
            Dim SelArr As String()
            If Strings.StartsWith("personalpronoun:") Then
                Cat = CachedData.IslamData.GrammarCategories(5)
                SelArr = Strings.Replace("personalpronoun:", String.Empty).Split(","c)
            Else
                Cat = CachedData.IslamData.GrammarCategories(6)
                SelArr = Strings.Replace("possessivedeterminerpersonalpronoun:", String.Empty).Split(","c)
            End If
            GrammarCat.Title = Cat.Title
            For SubCount = 0 To Cat.Words.Length - 1
                If Array.IndexOf(SelArr, Cat.Words(SubCount).TranslationID) <> -1 Then
                    Words.Add(Cat.Words(SubCount))
                End If
            Next
            GrammarCat.Words = Words.ToArray()
            Renderer.Items.Add(New RenderArray.RenderItem(RenderArray.RenderTypes.eText, New RenderArray.RenderText() {New RenderArray.RenderText(RenderArray.RenderDisplayClass.eList, Arabic.DisplayPronoun(GrammarCat, True))}))
        ElseIf Strings.StartsWith("particle:") Then
            Dim SelArr As String() = Strings.Replace("particle:", String.Empty).Split(","c)
            'Arabic.DisplayParticle()
        ElseIf Strings.StartsWith("noun:") Then
            Dim GrammarCat As New IslamData.GrammarCategory
            Dim Words As New List(Of IslamData.GrammarCategory.GrammarWord)
            Dim SelArr As String() = Strings.Replace("noun:", String.Empty).Split(","c)
            GrammarCat.Title = CachedData.IslamData.GrammarCategories(17).Title
            Strings.Replace("noun:", String.Empty).Split(","c)
            For SubCount = 0 To CachedData.IslamData.GrammarCategories(17).Words.Length - 1
                If Array.IndexOf(SelArr, CachedData.IslamData.GrammarCategories(17).Words(SubCount).TranslationID) <> -1 Then
                    Words.Add(CachedData.IslamData.GrammarCategories(17).Words(SubCount))
                End If
            Next
            GrammarCat.Words = Words.ToArray()
            Renderer.Items.Add(New RenderArray.RenderItem(RenderArray.RenderTypes.eText, New RenderArray.RenderText() {New RenderArray.RenderText(RenderArray.RenderDisplayClass.eList, Arabic.NounDisplay(GrammarCat))}))
        ElseIf Strings.StartsWith("verb:") Then
            Dim SelArr As String() = Strings.Replace("verb:", String.Empty).Split(","c)
            'Arabic.VerbDisplay()
        ElseIf Strings.StartsWith("phrase:") Then
            Dim SelArr As String() = Strings.Replace("phrase:", String.Empty).Split(","c)
            'Arabic.DoDisplayTranslation()
        ElseIf Strings.StartsWith("supp:") Then
            Dim SelArr As String() = Strings.Replace("supp:", String.Empty).Split(","c)
            For SelCount As Integer = 0 To SelArr.Length - 1
                For Count As Integer = 0 To CachedData.IslamData.VerseCategories.Length - 1
                    For SubCount As Integer = 0 To CachedData.IslamData.VerseCategories(Count).Verses.Length - 1
                        If SelArr(SelCount) = CachedData.IslamData.VerseCategories(Count).Verses(SubCount).TranslationID Then
                            Renderer.Items.AddRange(Supplications.DoGetRenderedVerseText(SchemeType, Scheme, CachedData.IslamData.VerseCategories(Count).Verses(SubCount), TranslationIndex))
                        End If
                    Next
                Next
            Next
        ElseIf Array.IndexOf(CachedData.IslamData.Abbreviations, Strings) <> 0 Then
        ElseIf Strings.StartsWith("reference:") Then
            Renderer.Items.Add(New RenderArray.RenderItem(RenderArray.RenderTypes.eHeaderCenter, New RenderArray.RenderText() {New RenderArray.RenderText(RenderArray.RenderDisplayClass.eLTR, "(" + Strings.Replace("reference:", String.Empty) + ")")}))
            'ElseIf Strings.StartsWith("text:") Then
        Else
                Renderer.Items.Add(New RenderArray.RenderItem(RenderArray.RenderTypes.eText, New RenderArray.RenderText() {New RenderArray.RenderText(RenderArray.RenderDisplayClass.eLTR, Strings)}))
        End If
        Return Renderer
    End Function
End Class
Public Class Supplications
    Public Shared Function GetSuppCategories() As String()
        Return Array.ConvertAll(CachedData.IslamData.VerseCategories, Function(Convert As IslamData.VerseCategory) Utility.LoadResourceString("IslamInfo_" + Convert.Title))
    End Function
    Public Shared Function GetRenderedSuppText(ByVal Item As PageLoader.TextItem) As RenderArray
        Dim SchemeType As ArabicData.TranslitScheme = CType(If(CInt(HttpContext.Current.Request.QueryString.Get("translitscheme")) >= 2, 2 - CInt(HttpContext.Current.Request.QueryString.Get("translitscheme")) Mod 2, CInt(HttpContext.Current.Request.QueryString.Get("translitscheme"))), ArabicData.TranslitScheme)
        Dim Scheme As String = If(CInt(HttpContext.Current.Request.QueryString.Get("translitscheme")) >= 2, CachedData.IslamData.TranslitSchemes((CInt(HttpContext.Current.Request.QueryString.Get("translitscheme")) - 2) \ 2).Name, String.Empty)
        Dim Count As Integer = CInt(HttpContext.Current.Request.QueryString.Get("selection"))
        If Count = -1 Then Count = 0
        Return DoGetRenderedSuppText(SchemeType, Scheme, CachedData.IslamData.VerseCategories(Count), TanzilReader.GetTranslationIndex(HttpContext.Current.Request.QueryString.Get("qurantranslation")))
    End Function
    Public Shared Function DoGetRenderedVerseText(SchemeType As ArabicData.TranslitScheme, Scheme As String, Verse As IslamData.VerseCategory.Verse, TranslationIndex As Integer) As List(Of RenderArray.RenderItem)
        Return DocBuilder.BuckwalterTextFromReferences(SchemeType, Scheme, Verse.Arabic, Verse.TranslationID, TranslationIndex).Items
    End Function
    Public Shared Function DoGetRenderedSuppText(SchemeType As ArabicData.TranslitScheme, Scheme As String, Category As IslamData.VerseCategory, TranslationIndex As Integer) As RenderArray
        Dim Renderer As New RenderArray
        For SubCount As Integer = 0 To Category.Verses.Length - 1
            Renderer.Items.AddRange(DoGetRenderedVerseText(SchemeType, Scheme, Category.Verses(SubCount), TranslationIndex))
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
            Return ArabicData.ArabicLetters
        ElseIf Quiz = 1 Then
            Dim CurList As New Generic.List(Of String)
            For Count As Integer = 0 To ArabicDiacriticsBefore.Length - 1
                Dim CurLet As String = ArabicDiacriticsBefore(Count)
                CurList.AddRange(Array.ConvertAll(ArabicData.ArabicLetters, Function(Str As String) CurLet + Str))
            Next
            Return CurList.ToArray()
        ElseIf Quiz = 2 Then
            Dim CurList As New Generic.List(Of String)
            For Count As Integer = 0 To ArabicDiacriticsAfter.Length - 1
                Dim CurLet As String = ArabicDiacriticsAfter(Count)
                CurList.AddRange(Array.ConvertAll(ArabicData.ArabicLetters, Function(Str As String) Str + CurLet))
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
        Dim Scheme As String = If(CInt(HttpContext.Current.Request.QueryString.Get("translitscheme")) >= 2, CachedData.IslamData.TranslitSchemes((CInt(HttpContext.Current.Request.QueryString.Get("translitscheme")) - 2) \ 2).Name, String.Empty)
        If SchemeType = ArabicData.TranslitScheme.None Then
            SchemeType = ArabicData.TranslitScheme.RuleBased
            Scheme = CachedData.IslamData.TranslitSchemes(3).Name
        End If
        Return Arabic.TransliterateToScheme(QuizSet(CInt(Math.Floor(Rnd() * QuizSet.Count))), SchemeType, Scheme)
    End Function
    Public Shared Function VerifyAnswer() As String()
        Dim JSList As New List(Of String) From {"javascript: verifyAnswer(this);", String.Empty, _
            Arabic.GetArabicSymbolJSArray(), Arabic.GetTranslitSchemeJSArray(), Arabic.FindLetterBySymbolJS, _
            "var arabicLets = " + Utility.MakeJSArray(ArabicData.ArabicLetters, False) + ";", _
            "var arabicDiacriticsBefore = " + Utility.MakeJSArray(ArabicDiacriticsBefore, False) + ";", _
            "var arabicDiacriticsAfter = " + Utility.MakeJSArray(ArabicDiacriticsAfter, False) + ";", _
            "var renderList = [], qtype = 'arabicletters', qwrong = 0, qright = 0;", _
            "function getUniqueRnd(excl, count) { var rnd; if (excl.length === count) return 0; do { rnd = Math.floor(Math.random() * count); } while (excl.indexOf(rnd) !== -1); return rnd; }", _
            "function getQuizSet() { if (qtype === 'arabicletters') return arabicLets; if (qtype === 'arabiclettersdiacritics') { var count = 0, arr = []; for (count = 0; count < arabicDiacriticsAfter.length; count++) { arr = arr.concat(arabicLets.map(function(val) { return val + arabicDiacriticsAfter[count]; })); } return arr; } if (qtype === 'arabicdiacriticsletters') { var count = 0, arr = []; for (count = 0; count < arabicDiacriticsBefore.length; count++) { arr = arr.concat(arabicLets.map(function(val) { return arabicDiacriticsBefore[count] + val; })); } return arr; }; return []; }", _
            "function getQA(quizSet, quest, nidx) { if (quest) return quizSet[nidx]; return (parseInt($('#translitscheme').val(), 10) % 2) === 0 ? transliterateWithRules(quizSet[nidx], parseInt($('#translitscheme').val(), 10) >= 2 ? Math.floor((parseInt($('#translitscheme').val(), 10) - 2) / 2) + 2 : 5) : doTransliterate(quizSet[nidx], true, parseInt($('#translitscheme').val(), 10)); }", _
            "function nextQuestion() { $('#count').text('Wrong: ' + qwrong + ' Right: ' + qright); var i = Math.floor(Math.random() * 4), quizSet = getQuizSet(), pos = quizSet.length, nidx = getUniqueRnd([], pos), aidx = []; aidx[0] = getUniqueRnd([nidx], pos); aidx[1] = getUniqueRnd([nidx, aidx[0]], pos); aidx[2] = getUniqueRnd([nidx, aidx[0], aidx[1]], pos); $('#quizquestion').text(getQA(quizSet, true, nidx)); $('#answer1').prop('value', getQA(quizSet, false, i === 0 ? nidx : aidx[0])); $('#answer2').prop('value', getQA(quizSet, false, i === 1 ? nidx : aidx[i > 1 ? 1 : 0])); $('#answer3').prop('value', getQA(quizSet, false, i === 2 ? nidx : aidx[i > 2 ? 2 : 1])); $('#answer4').prop('value', getQA(quizSet, false, i === 3 ? nidx : aidx[2])); }", _
            "function verifyAnswer(ctl) { $(ctl).prop('value') === ((parseInt($('#translitscheme').val(), 10) % 2) === 0 ? transliterateWithRules($('#quizquestion').text().trim(), parseInt($('#translitscheme').val(), 10) >= 2 ? Math.floor((parseInt($('#translitscheme').val(), 10) - 2) / 2) + 2 : 5) : doTransliterate($('#quizquestion').text().trim(), true, parseInt($('#translitscheme').val(), 10))) ? qright++ : qwrong++; nextQuestion(); }"}
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
            Utility.MakeJSArray(Array.ConvertAll(Of Array, String)(ArabicData.GetRecitationSymbols(), Function(Convert As Array) Utility.MakeJSArray(New String() {CStr(CType(Convert, Object())(0)), CStr(CType(Convert, Object())(1))})), True)}, True)
        Return New String() {"javascript: changeQuranDivision(this.selectedIndex);", String.Empty, Utility.GetClearOptionListJS(), _
        "function changeQuranDivision(index) { var iCount; var qurandata = " + JSArrays + "; var eSelect = $('#quranselection').get(0); clearOptionList(eSelect); for (iCount = 0; iCount < qurandata[index].length; iCount++) { eSelect.options.add(new Option(qurandata[index][iCount][0], qurandata[index][iCount][1])); } }"}
    End Function
    Public Shared Function GetWordPartitions() As String()
        Dim Parts As New Generic.List(Of String) From {Utility.LoadResourceString("IslamInfo_Letters"), Utility.LoadResourceString("IslamInfo_Words"), Utility.LoadResourceString("IslamInfo_UniqueWords"), Utility.LoadResourceString("IslamInfo_UniqueWordsPerPart"), Utility.LoadResourceString("IslamInfo_WordsPerPart"), Utility.LoadResourceString("IslamInfo_UniqueWordsPerStation"), Utility.LoadResourceString("IslamInfo_WordsPerStation"), Utility.LoadResourceString("IslamInfo_IsolatedLetters"), Utility.LoadResourceString("IslamInfo_LetterPatterns"), Utility.LoadResourceString("IslamInfo_Prefix"), Utility.LoadResourceString("IslamInfo_Suffix")}
        Parts.AddRange(Array.ConvertAll(CachedData.IslamData.PartsOfSpeech, Function(POS As IslamData.PartOfSpeechInfo) Utility.LoadResourceString("IslamInfo_" + POS.Id)))
        Parts.AddRange(Array.ConvertAll(ArabicData.RecitationSymbols, Function(Sym As Char) ArabicData.GetUnicodeName(Sym)))
        Parts.AddRange(Array.ConvertAll(ArabicData.RecitationSymbols, Function(Sym As Char) "Prefix of " + ArabicData.GetUnicodeName(Sym)))
        Parts.AddRange(Array.ConvertAll(ArabicData.RecitationSymbols, Function(Sym As Char) "Suffix of " + ArabicData.GetUnicodeName(Sym)))
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
        ElseIf Index >= 11 + CachedData.IslamData.PartsOfSpeech.Length And Index < 11 + CachedData.IslamData.PartsOfSpeech.Length + ArabicData.RecitationSymbols.Length Then
            Return CStr(CachedData.LetterDictionary.Item(ArabicData.RecitationSymbols(Index - 11 - CachedData.IslamData.PartsOfSpeech.Length)).Count)
        ElseIf Index >= 11 + CachedData.IslamData.PartsOfSpeech.Length + ArabicData.RecitationSymbols.Length And Index < 11 + CachedData.IslamData.PartsOfSpeech.Length + ArabicData.RecitationSymbols.Length + ArabicData.RecitationSymbols.Length Then
            Return CStr(CachedData.LetterPreDictionary.Item(ArabicData.RecitationSymbols(Index - 11 - CachedData.IslamData.PartsOfSpeech.Length - ArabicData.RecitationSymbols.Length)).Count)
        ElseIf Index >= 11 + CachedData.IslamData.PartsOfSpeech.Length + ArabicData.RecitationSymbols.Length + ArabicData.RecitationSymbols.Length And Index < 11 + CachedData.IslamData.PartsOfSpeech.Length + ArabicData.RecitationSymbols.Length + ArabicData.RecitationSymbols.Length + ArabicData.RecitationSymbols.Length Then
            Return CStr(CachedData.LetterSufDictionary.Item(ArabicData.RecitationSymbols(Index - 11 - CachedData.IslamData.PartsOfSpeech.Length - ArabicData.RecitationSymbols.Length - ArabicData.RecitationSymbols.Length)).Count)
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
                Output.Add(New String() {ArabicData.LeftToRightOverride + ArabicData.GetUnicodeName(LetterFreqArray(Count)) + " ( " + ArabicData.PopDirectionalFormatting + ArabicData.FixStartingCombiningSymbol(LetterFreqArray(Count)) + ArabicData.LeftToRightOverride + " )" + ArabicData.PopDirectionalFormatting, String.Empty, CStr(CachedData.LetterDictionary.Item(LetterFreqArray(Count)).Count), (CDbl(CachedData.LetterDictionary.Item(LetterFreqArray(Count)).Count) * 100 / All).ToString("n2"), (CDbl(Total) * 100 / All).ToString("n2")})
            Next
        ElseIf Index = 7 Then
            All = CachedData.TotalIsolatedLetters
            Dim LetterFreqArray(CachedData.IsolatedLetterDictionary.Keys.Count - 1) As Char
            CachedData.IsolatedLetterDictionary.Keys.CopyTo(LetterFreqArray, 0)
            Array.Sort(LetterFreqArray, Function(Key As Char, NextKey As Char) CachedData.IsolatedLetterDictionary.Item(NextKey).Count.CompareTo(CachedData.IsolatedLetterDictionary.Item(Key).Count))
            For Count As Integer = 0 To LetterFreqArray.Length - 1
                Total += CachedData.IsolatedLetterDictionary.Item(LetterFreqArray(Count)).Count
                Output.Add(New String() {ArabicData.LeftToRightOverride + ArabicData.GetUnicodeName(LetterFreqArray(Count)) + " ( " + ArabicData.PopDirectionalFormatting + ArabicData.FixStartingCombiningSymbol(LetterFreqArray(Count)) + ArabicData.LeftToRightOverride + " )" + ArabicData.PopDirectionalFormatting, String.Empty, CStr(CachedData.IsolatedLetterDictionary.Item(LetterFreqArray(Count)).Count), (CDbl(CachedData.IsolatedLetterDictionary.Item(LetterFreqArray(Count)).Count) * 100 / All).ToString("n2"), (CDbl(Total) * 100 / All).ToString("n2")})
            Next
        ElseIf Index = 1 Or Index = 9 Or Index = 10 Or Index >= 11 And Index < 11 + CachedData.IslamData.PartsOfSpeech.Length + ArabicData.RecitationSymbols.Length + ArabicData.RecitationSymbols.Length + ArabicData.RecitationSymbols.Length Then
            Dim Dict As Generic.Dictionary(Of String, ArrayList)
            If Index = 1 Then
                Dict = CachedData.WordDictionary
            ElseIf Index = 9 Then
                Dict = CachedData.PreDictionary
            ElseIf Index = 10 Then
                Dict = CachedData.SufDictionary
            ElseIf Index >= 11 And Index < 11 + CachedData.IslamData.PartsOfSpeech.Length Then
                Dict = CachedData.TagDictionary(CachedData.IslamData.PartsOfSpeech(Index - 11).Symbol)
            ElseIf Index >= 11 + CachedData.IslamData.PartsOfSpeech.Length And Index < 11 + CachedData.IslamData.PartsOfSpeech.Length + ArabicData.RecitationSymbols.Length Then
                Dict = CachedData.LetterDictionary(ArabicData.RecitationSymbols(Index - 11 - CachedData.IslamData.PartsOfSpeech.Length))
            ElseIf Index >= 11 + CachedData.IslamData.PartsOfSpeech.Length + ArabicData.RecitationSymbols.Length And Index < 11 + CachedData.IslamData.PartsOfSpeech.Length + ArabicData.RecitationSymbols.Length + ArabicData.RecitationSymbols.Length Then
                Dict = CachedData.LetterPreDictionary(ArabicData.RecitationSymbols(Index - 11 - CachedData.IslamData.PartsOfSpeech.Length - ArabicData.RecitationSymbols.Length))
            ElseIf Index >= 11 + CachedData.IslamData.PartsOfSpeech.Length + ArabicData.RecitationSymbols.Length + ArabicData.RecitationSymbols.Length And Index < 11 + CachedData.IslamData.PartsOfSpeech.Length + ArabicData.RecitationSymbols.Length + ArabicData.RecitationSymbols.Length + ArabicData.RecitationSymbols.Length Then
                Dict = CachedData.LetterSufDictionary(ArabicData.RecitationSymbols(Index - 11 - CachedData.IslamData.PartsOfSpeech.Length - ArabicData.RecitationSymbols.Length - ArabicData.RecitationSymbols.Length))
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
                Output.Add(New String() {ArabicData.TransliterateFromBuckwalter(FreqArray(Count)), String.Empty, CStr(Dict.Item(FreqArray(Count)).Count), (CDbl(Dict.Item(FreqArray(Count)).Count) * 100 / All).ToString("n2"), (CDbl(Total) * 100 / All).ToString("n2")})
            Next
        ElseIf Index = 3 Or Index = 4 Or Index = 5 Or Index = 6 Then
            Total = 0
            Dim DivArray As Collections.Generic.List(Of String)()
            If Index = 3 Or Index = 5 Then
                DivArray = If(Index = 5, CachedData.StationUniqueArray, CachedData.PartUniqueArray)
                All = If(Index = 5, CachedData.TotalUniqueWordsInStations, CachedData.TotalUniqueWordsInParts)
                For Count As Integer = 0 To CInt(IIf(Index = 5, TanzilReader.GetStationCount(), TanzilReader.GetPartCount())) - 1
                    Total += DivArray(Count).Count
                    Output.Add(New String() {ArabicData.LeftToRightOverride + CStr(Count + 1) + ArabicData.PopDirectionalFormatting, String.Empty, CStr(DivArray(Count).Count), (CDbl(DivArray(Count).Count) * 100 / All).ToString("n2"), (CDbl(Total) * 100 / All).ToString("n2")})
                Next
            ElseIf Index = 4 Or Index = 6 Then
                DivArray = If(Index = 6, CachedData.StationUniqueArray, CachedData.PartUniqueArray)
                All = If(Index = 6, CachedData.TotalWordsInStations, CachedData.TotalWordsInParts)
                For Count As Integer = 0 To CInt(IIf(Index = 6, TanzilReader.GetStationCount(), TanzilReader.GetPartCount())) - 1
                    Total += DivArray(Count).Count
                    Output.Add(New String() {ArabicData.LeftToRightOverride + CStr(Count + 1) + ArabicData.PopDirectionalFormatting, String.Empty, CStr(DivArray(Count).Count), (CDbl(DivArray(Count).Count) * 100 / All).ToString("n2"), (CDbl(Total) * 100 / All).ToString("n2")})
                Next
            End If
        ElseIf Index = 8 Then
            Output.AddRange(Array.ConvertAll(GetQuranLetterPatterns(), Function(Str As String) {Str}))
        End If
        Return CType(Output.ToArray(GetType(Array)), Array())
    End Function
    Public Shared Function GetQuranLetterPatterns() As String()
        Dim RecSymbols As String = String.Join(String.Empty, Array.ConvertAll(ArabicData.RecitationSpecialSymbols, Function(C As Char) CStr(C)))
        Dim LtrSymbols As String = String.Join(String.Empty, Array.ConvertAll(ArabicData.RecitationLetters, Function(C As Char) CStr(C)))
        Dim DiaSymbols As String = String.Join(String.Empty, Array.ConvertAll(ArabicData.RecitationDiacritics, Function(C As Char) CStr(C)))
        Dim StartWordMultiOnly As New Generic.Dictionary(Of String, String)
        Dim EndWordMultiOnly As New Generic.Dictionary(Of String, String)
        Dim MiddleWordMultiOnly As New Generic.Dictionary(Of String, String)
        Dim StartWordOnly As String = String.Join(String.Empty, Array.ConvertAll(ArabicData.RecitationLettersDiacritics, Function(C As Char) CStr(C)))
        Dim NotStartWord As String = String.Join(String.Empty, Array.ConvertAll(ArabicData.RecitationLettersDiacritics, Function(C As Char) CStr(C)))
        Dim EndWordOnly As String = String.Join(String.Empty, Array.ConvertAll(ArabicData.RecitationLettersDiacritics, Function(C As Char) CStr(C)))
        Dim NotEndWord As String = String.Join(String.Empty, Array.ConvertAll(ArabicData.RecitationLettersDiacritics, Function(C As Char) CStr(C)))
        Dim EndWordOnlyNoDia As String = String.Join(String.Empty, Array.ConvertAll(ArabicData.RecitationLetters, Function(C As Char) CStr(C)))
        Dim NotEndWordNoDia As String = String.Join(String.Empty, Array.ConvertAll(ArabicData.RecitationLetters, Function(C As Char) CStr(C)))
        Dim MiddleWordOnlyNoDia As String = String.Join(String.Empty, Array.ConvertAll(ArabicData.RecitationLetters, Function(C As Char) CStr(C)))
        Dim NotMiddleWordNoDia As String = String.Join(String.Empty, Array.ConvertAll(ArabicData.RecitationLetters, Function(C As Char) CStr(C)))
        Dim MiddleWordOnly As String = String.Join(String.Empty, Array.ConvertAll(ArabicData.RecitationLettersDiacritics, Function(C As Char) CStr(C)))
        Dim NotMiddleWord As String = String.Join(String.Empty, Array.ConvertAll(ArabicData.RecitationLettersDiacritics, Function(C As Char) CStr(C)))
        Dim DiaStartWordOnly As String = String.Join(String.Empty, Array.ConvertAll(ArabicData.RecitationDiacritics, Function(C As Char) CStr(C)))
        Dim DiaNotStartWord As String = String.Join(String.Empty, Array.ConvertAll(ArabicData.RecitationDiacritics, Function(C As Char) CStr(C)))
        Dim DiaEndWordOnly As String = String.Join(String.Empty, Array.ConvertAll(ArabicData.RecitationDiacritics, Function(C As Char) CStr(C)))
        Dim DiaNotEndWord As String = String.Join(String.Empty, Array.ConvertAll(ArabicData.RecitationDiacritics, Function(C As Char) CStr(C)))
        Dim DiaMiddleWordOnly As String = String.Join(String.Empty, Array.ConvertAll(ArabicData.RecitationDiacritics, Function(C As Char) CStr(C)))
        Dim DiaNotMiddleWord As String = String.Join(String.Empty, Array.ConvertAll(ArabicData.RecitationDiacritics, Function(C As Char) CStr(C)))
        Dim Combos As String() = String.Join("|", Array.ConvertAll(ArabicData.RecitationLettersDiacritics, Function(C As Char) String.Join("|", Array.ConvertAll(ArabicData.RecitationLettersDiacritics, Function(Nxt As Char) C + Nxt)))).Split("|"c)
        Dim DiaCombos As String() = String.Join("|", Array.ConvertAll(ArabicData.RecitationDiacritics, Function(C As Char) String.Join("|", Array.ConvertAll(ArabicData.RecitationDiacritics, Function(Nxt As Char) C + Nxt)))).Split("|"c)
        Dim LetCombos As String() = String.Join("|", Array.ConvertAll(ArabicData.RecitationLetters, Function(C As Char) String.Join("|", Array.ConvertAll(ArabicData.RecitationLetters, Function(Nxt As Char) C + Nxt)))).Split("|"c)
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
            Str = ArabicData.TransliterateFromBuckwalter(Str)
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
        Dim Val As String = ArabicData.LeftToRightOverride + "Combinations: " + ArabicData.PopDirectionalFormatting
        For Each Key As Char In Dict.Keys
            If Dict.Item(Key).Length > (DiaSymbols.Length + LtrSymbols.Length) / 2 Then
                Val += ArabicData.FixStartingCombiningSymbol(Key) + ArabicData.LeftToRightOverride + " [" + ArabicData.PopDirectionalFormatting + String.Join(" ", Array.ConvertAll(New String(Array.FindAll((DiaSymbols + LtrSymbols).ToCharArray(), Function(C As Char) Not Dict.Item(Key).Contains(C))).ToCharArray(), Function(C As Char) ArabicData.FixStartingCombiningSymbol(CStr(C)))) + ArabicData.LeftToRightOverride + " ]" + ArabicData.PopDirectionalFormatting + vbTab
            Else
                Val += ArabicData.FixStartingCombiningSymbol(Key) + ArabicData.LeftToRightOverride + " ! [ " + ArabicData.PopDirectionalFormatting + String.Join(" ", Array.ConvertAll(Dict.Item(Key).ToCharArray(), Function(C As Char) ArabicData.FixStartingCombiningSymbol(CStr(C)))) + ArabicData.LeftToRightOverride + " ]" + ArabicData.PopDirectionalFormatting + vbTab
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
        Dim RevVal As String = ArabicData.LeftToRightOverride + "Reverse Combinations: " + ArabicData.PopDirectionalFormatting
        For Each Key As Char In RevDict.Keys
            If RevDict.Item(Key).Length > (DiaSymbols.Length + LtrSymbols.Length) / 2 Then
                RevVal += ArabicData.LeftToRightOverride + "[" + ArabicData.PopDirectionalFormatting + String.Join(" ", Array.ConvertAll(New String(Array.FindAll((DiaSymbols + LtrSymbols).ToCharArray(), Function(C As Char) Not RevDict.Item(Key).Contains(C))).ToCharArray(), Function(C As Char) ArabicData.FixStartingCombiningSymbol(CStr(C)))) + ArabicData.LeftToRightOverride + " ] " + ArabicData.PopDirectionalFormatting + ArabicData.FixStartingCombiningSymbol(Key) + vbTab
            Else
                RevVal += ArabicData.LeftToRightOverride + "! [ " + ArabicData.PopDirectionalFormatting + String.Join(" ", Array.ConvertAll(RevDict.Item(Key).ToCharArray(), Function(C As Char) ArabicData.FixStartingCombiningSymbol(CStr(C)))) + ArabicData.LeftToRightOverride + " ] " + ArabicData.PopDirectionalFormatting + ArabicData.FixStartingCombiningSymbol(Key) + vbTab
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
        Dim DiaVal As String = ArabicData.LeftToRightOverride + "Diacritic Only Combinations: " + ArabicData.PopDirectionalFormatting
        For Each Key As Char In DiaDict.Keys
            If DiaDict.Item(Key).Length > DiaSymbols.Length / 2 Then
                DiaVal += ArabicData.FixStartingCombiningSymbol(Key) + ArabicData.LeftToRightOverride + " [" + ArabicData.PopDirectionalFormatting + String.Join(" ", Array.ConvertAll(New String(Array.FindAll(DiaSymbols.ToCharArray(), Function(C As Char) Not DiaDict.Item(Key).Contains(C))).ToCharArray(), Function(C As Char) ArabicData.FixStartingCombiningSymbol(CStr(C)))) + ArabicData.LeftToRightOverride + " ]" + ArabicData.PopDirectionalFormatting + vbTab
            Else
                DiaVal += ArabicData.FixStartingCombiningSymbol(Key) + ArabicData.LeftToRightOverride + " ! [ " + ArabicData.PopDirectionalFormatting + String.Join(" ", Array.ConvertAll(DiaDict.Item(Key).ToCharArray(), Function(C As Char) ArabicData.FixStartingCombiningSymbol(CStr(C)))) + ArabicData.LeftToRightOverride + " ]" + ArabicData.PopDirectionalFormatting + vbTab
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
        Dim LetVal As String = ArabicData.LeftToRightOverride + "Letter Only Combinations: " + ArabicData.PopDirectionalFormatting
        For Each Key As Char In LetDict.Keys
            If LetDict.Item(Key).Length > LtrSymbols.Length / 2 Then
                LetVal += ArabicData.FixStartingCombiningSymbol(Key) + ArabicData.LeftToRightOverride + " [" + ArabicData.PopDirectionalFormatting + String.Join(" ", Array.ConvertAll(New String(Array.FindAll(LtrSymbols.ToCharArray(), Function(C As Char) Not LetDict.Item(Key).Contains(C))).ToCharArray(), Function(C As Char) ArabicData.FixStartingCombiningSymbol(CStr(C)))) + ArabicData.LeftToRightOverride + " ]" + ArabicData.PopDirectionalFormatting + vbTab
            Else
                LetVal += ArabicData.FixStartingCombiningSymbol(Key) + ArabicData.LeftToRightOverride + " ! [ " + ArabicData.PopDirectionalFormatting + String.Join(" ", Array.ConvertAll(LetDict.Item(Key).ToCharArray(), Function(C As Char) ArabicData.FixStartingCombiningSymbol(CStr(C)))) + ArabicData.LeftToRightOverride + " ]" + ArabicData.PopDirectionalFormatting + vbTab
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
        Dim LetRevVal As String = ArabicData.LeftToRightOverride + "Reverse Letter Only Combinations: " + ArabicData.PopDirectionalFormatting
        For Each Key As Char In LetRevDict.Keys
            If LetRevDict.Item(Key).Length > LtrSymbols.Length / 2 Then
                LetRevVal += ArabicData.LeftToRightOverride + "[" + ArabicData.PopDirectionalFormatting + String.Join(" ", Array.ConvertAll(New String(Array.FindAll(LtrSymbols.ToCharArray(), Function(C As Char) Not LetRevDict.Item(Key).Contains(C))).ToCharArray(), Function(C As Char) ArabicData.FixStartingCombiningSymbol(CStr(C)))) + ArabicData.LeftToRightOverride + " ] " + ArabicData.PopDirectionalFormatting + ArabicData.FixStartingCombiningSymbol(Key) + vbTab
            Else
                LetRevVal += ArabicData.LeftToRightOverride + "! [ " + ArabicData.PopDirectionalFormatting + String.Join(" ", Array.ConvertAll(LetRevDict.Item(Key).ToCharArray(), Function(C As Char) ArabicData.FixStartingCombiningSymbol(CStr(C)))) + ArabicData.LeftToRightOverride + " ] " + ArabicData.PopDirectionalFormatting + ArabicData.FixStartingCombiningSymbol(Key) + vbTab
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
        Return {ArabicData.LeftToRightOverride + "Unique Prefix: [" + ArabicData.PopDirectionalFormatting + StartMulti + ArabicData.LeftToRightOverride + "]" + ArabicData.PopDirectionalFormatting, _
                ArabicData.LeftToRightOverride + "Unique Suffix: [" + ArabicData.PopDirectionalFormatting + EndMulti + ArabicData.LeftToRightOverride + "]" + ArabicData.PopDirectionalFormatting, _
                ArabicData.LeftToRightOverride + "Unique Middle: [" + ArabicData.PopDirectionalFormatting + MiddleMulti + ArabicData.LeftToRightOverride + "]" + ArabicData.PopDirectionalFormatting, _
                ArabicData.LeftToRightOverride + "Start Only: [" + ArabicData.PopDirectionalFormatting + String.Join(" ", Array.ConvertAll(StartWordOnly.ToCharArray(), Function(C As Char) ArabicData.FixStartingCombiningSymbol(CStr(C)))) + ArabicData.LeftToRightOverride + "]" + ArabicData.PopDirectionalFormatting, _
                ArabicData.LeftToRightOverride + "Not Start: [" + ArabicData.PopDirectionalFormatting + String.Join(" ", Array.ConvertAll(NotStartWord.ToCharArray(), Function(C As Char) ArabicData.FixStartingCombiningSymbol(CStr(C)))) + ArabicData.LeftToRightOverride + "]" + ArabicData.PopDirectionalFormatting, _
                ArabicData.LeftToRightOverride + "End Only: [" + ArabicData.PopDirectionalFormatting + String.Join(" ", Array.ConvertAll(EndWordOnly.ToCharArray(), Function(C As Char) ArabicData.FixStartingCombiningSymbol(CStr(C)))) + ArabicData.LeftToRightOverride + "]" + ArabicData.PopDirectionalFormatting, _
                ArabicData.LeftToRightOverride + "Not End: [" + ArabicData.PopDirectionalFormatting + String.Join(" ", Array.ConvertAll(NotEndWord.ToCharArray(), Function(C As Char) ArabicData.FixStartingCombiningSymbol(CStr(C)))) + ArabicData.LeftToRightOverride + "]" + ArabicData.PopDirectionalFormatting, _
                ArabicData.LeftToRightOverride + "End Only No Diacritics: [" + ArabicData.PopDirectionalFormatting + String.Join(" ", Array.ConvertAll(EndWordOnlyNoDia.ToCharArray(), Function(C As Char) ArabicData.FixStartingCombiningSymbol(CStr(C)))) + ArabicData.LeftToRightOverride + "]" + ArabicData.PopDirectionalFormatting, _
                ArabicData.LeftToRightOverride + "Not End No Diacritics: [" + ArabicData.PopDirectionalFormatting + String.Join(" ", Array.ConvertAll(NotEndWordNoDia.ToCharArray(), Function(C As Char) ArabicData.FixStartingCombiningSymbol(CStr(C)))) + ArabicData.LeftToRightOverride + "]" + ArabicData.PopDirectionalFormatting, _
                ArabicData.LeftToRightOverride + "Middle Only: [" + ArabicData.PopDirectionalFormatting + String.Join(" ", Array.ConvertAll(MiddleWordOnly.ToCharArray(), Function(C As Char) ArabicData.FixStartingCombiningSymbol(CStr(C)))) + ArabicData.LeftToRightOverride + "]" + ArabicData.PopDirectionalFormatting, _
                ArabicData.LeftToRightOverride + "Not Middle: [" + ArabicData.PopDirectionalFormatting + String.Join(" ", Array.ConvertAll(NotMiddleWord.ToCharArray(), Function(C As Char) ArabicData.FixStartingCombiningSymbol(CStr(C)))) + ArabicData.LeftToRightOverride + "]" + ArabicData.PopDirectionalFormatting, _
                ArabicData.LeftToRightOverride + "Middle Only No Diacritics: [" + ArabicData.PopDirectionalFormatting + String.Join(" ", Array.ConvertAll(MiddleWordOnlyNoDia.ToCharArray(), Function(C As Char) ArabicData.FixStartingCombiningSymbol(CStr(C)))) + ArabicData.LeftToRightOverride + "]" + ArabicData.PopDirectionalFormatting, _
                ArabicData.LeftToRightOverride + "Not Middle No Diacritics: [" + ArabicData.PopDirectionalFormatting + String.Join(" ", Array.ConvertAll(NotMiddleWordNoDia.ToCharArray(), Function(C As Char) ArabicData.FixStartingCombiningSymbol(CStr(C)))) + ArabicData.LeftToRightOverride + "]" + ArabicData.PopDirectionalFormatting, _
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
            Return ArabicData.GetRecitationSymbols()
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
            Path = Utility.GetFilePath("metadata\" + QuranTextNames(TargetBaseText) + "-" + QuranFileNames(ScriptType) + If(Presentation <> ArabicPresentation.None, "-" + PresentationCacheNames(Presentation), String.Empty) + ".xml")
        Else
            Path = Utility.GetFilePath("metadata\" + QuranTextNames(TargetBaseText) + "-" + QuranFileNames(ScriptType) + If(Presentation <> ArabicPresentation.None, "-" + PresentationCacheNames(Presentation), String.Empty) + ".xml")
        End If
        If Presentation = ArabicPresentation.Buckwalter Then
            UseBuckwalter = True
        End If
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
                        If UseBuckwalter Then PreVerse = ArabicData.TransliterateFromBuckwalter(PreVerse)
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
                ElseIf Count <> 0 And SubCount <> 0 Then
                    PreVerse = GetTextVerse(ChapterNode, SubCount).Attributes.GetNamedItem("text").Value
                End If
                If UseBuckwalter Then PreVerse = ArabicData.TransliterateFromBuckwalter(PreVerse)
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
        Return System.Text.RegularExpressions.Regex.Match(Str, "^(?:,?(\d+)(?:\:(\d+))?(?:\:(\d+))?(?:-(\d+)(?:\:(\d+))?(?:\:(\d+))?)?)+$").Success
    End Function
    Public Shared Function QuranTextFromReference(Str As String, SchemeType As ArabicData.TranslitScheme, Scheme As String, TranslationIndex As Integer) As RenderArray
        Dim Renderer As New RenderArray
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
            Renderer.Items.AddRange(DoGetRenderedQuranText(QuranTextRangeLookup(BaseChapter, BaseVerse, WordNumber, EndChapter, ExtraVerseNumber, EndWordNumber), BaseChapter, BaseVerse, CachedData.IslamData.Translations.TranslationList(TranslationIndex).Name, SchemeType, Scheme, TranslationIndex).Items)
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
            QuranText(0)(0) = System.Text.RegularExpressions.Regex.Replace(System.Text.RegularExpressions.Regex.Replace(QuranText(0)(0).Substring(0, VerseIndex), "(^\s*|\s+)[^\s" + String.Join(String.Empty, Array.ConvertAll(ArabicData.ArabicStopLetters, Function(Str As String) ArabicData.MakeUniRegEx(Str))) + ArabicData.ArabicStartOfRubElHizb + ArabicData.ArabicPlaceOfSajdah + "]+(?=\s*$|\s+)", "$1"), String.Join("|", Array.ConvertAll(ArabicData.ArabicStopLetters, Function(Str As String) ArabicData.MakeUniRegEx(Str))) + "|" + ArabicData.ArabicStartOfRubElHizb + "|" + ArabicData.ArabicPlaceOfSajdah, ChrW(0)) + QuranText(0)(0).Substring(VerseIndex)
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
            QuranText(LastChapter)(LastVerse) = QuranText(LastChapter)(LastVerse).Substring(0, VerseIndex) + System.Text.RegularExpressions.Regex.Replace(System.Text.RegularExpressions.Regex.Replace(QuranText(LastChapter)(LastVerse).Substring(VerseIndex), "(^\s*|\s+)[^\s" + String.Join(String.Empty, Array.ConvertAll(ArabicData.ArabicStopLetters, Function(Str As String) ArabicData.MakeUniRegEx(Str))) + ArabicData.ArabicStartOfRubElHizb + ArabicData.ArabicPlaceOfSajdah + "]+(?=\s*$|\s+)", "$1"), String.Join("|", Array.ConvertAll(ArabicData.ArabicStopLetters, Function(Str As String) ArabicData.MakeUniRegEx(Str))) + "|" + ArabicData.ArabicStartOfRubElHizb + "|" + ArabicData.ArabicPlaceOfSajdah, ChrW(0))
        End If
        Return QuranText
    End Function
    Public Shared Function GetQuranTextBySelection(Division As Integer, Index As Integer, Translation As String, SchemeType As ArabicData.TranslitScheme, Scheme As String, TranslationIndex As Integer) As RenderArray
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
            ReDim Keys(CachedData.LetterDictionary(ArabicData.Data.ArabicLetters(Index).Symbol).Count - 1)
            CachedData.LetterDictionary(ArabicData.Data.ArabicLetters(Index).Symbol).Keys.CopyTo(Keys, 0)
            SeperateSectionCount = CachedData.LetterDictionary(ArabicData.Data.ArabicLetters(Index).Symbol).Count
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
                For SubCount = 0 To CachedData.LetterDictionary(ArabicData.Data.ArabicLetters(Index).Symbol)(Keys(SectionCount)).Count - 1
                    BaseChapter = CType(CachedData.LetterDictionary(ArabicData.Data.ArabicLetters(Index).Symbol)(Keys(SectionCount))(SubCount), Integer())(0)
                    BaseVerse = CType(CachedData.LetterDictionary(ArabicData.Data.ArabicLetters(Index).Symbol)(Keys(SectionCount))(SubCount), Integer())(1)
                    QuranText.Add(GetQuranText(CachedData.XMLDocMain, BaseChapter, BaseVerse, BaseVerse))
                    If SubCount <> CachedData.LetterDictionary(ArabicData.Data.ArabicLetters(Index).Symbol)(Keys(SectionCount)).Count - 1 Then
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
        Dim SchemeType As ArabicData.TranslitScheme = CType(If(CInt(HttpContext.Current.Request.QueryString.Get("translitscheme")) >= 2, 2 - CInt(HttpContext.Current.Request.QueryString.Get("translitscheme")) Mod 2, CInt(HttpContext.Current.Request.QueryString.Get("translitscheme"))), ArabicData.TranslitScheme)
        Dim Scheme As String = If(CInt(HttpContext.Current.Request.QueryString.Get("translitscheme")) >= 2, CachedData.IslamData.TranslitSchemes((CInt(HttpContext.Current.Request.QueryString.Get("translitscheme")) - 2) \ 2).Name, String.Empty)
        Dim TranslationIndex As Integer = GetTranslationIndex(HttpContext.Current.Request.QueryString.Get("qurantranslation"))
        Return GetQuranTextBySelection(Division, Index, HttpContext.Current.Request.QueryString.Get("qurantranslation"), SchemeType, Scheme, TranslationIndex)
    End Function
    Public Shared Function DoGetRenderedQuranText(QuranText As Collections.Generic.List(Of String()), BaseChapter As Integer, BaseVerse As Integer, Translation As String, SchemeType As ArabicData.TranslitScheme, Scheme As String, TranslationIndex As Integer) As RenderArray
        Dim Text As String
        Dim Node As System.Xml.XmlNode
        Dim Renderer As New RenderArray
        Dim Lines As String() = IO.File.ReadAllLines(Utility.GetFilePath("metadata\" + GetTranslationFileName(Translation)))
        Dim W4WLines As String() = IO.File.ReadAllLines(Utility.GetFilePath("metadata\en.w4w.shehnazshaikh.txt"))
        If Not QuranText Is Nothing Then
            For Chapter = 0 To QuranText.Count - 1
                Dim ChapterNode As System.Xml.XmlNode = GetChapterByIndex(BaseChapter + Chapter)
                Renderer.Items.Add(New RenderArray.RenderItem(RenderArray.RenderTypes.eHeaderLeft, New RenderArray.RenderText() {New RenderArray.RenderText(RenderArray.RenderDisplayClass.eArabic, ArabicData.TransliterateFromBuckwalter("'aAya`tuhaA " + ChapterNode.Attributes.GetNamedItem("ayas").Value + " ")), New RenderArray.RenderText(RenderArray.RenderDisplayClass.eTransliteration, Arabic.TransliterateToScheme(ArabicData.TransliterateFromBuckwalter("'aAya`tuhaA " + ChapterNode.Attributes.GetNamedItem("ayas").Value + " "), SchemeType, Scheme).Trim()), New RenderArray.RenderText(RenderArray.RenderDisplayClass.eLTR, "Verses " + ChapterNode.Attributes.GetNamedItem("ayas").Value + " ")}))
                Renderer.Items.Add(New RenderArray.RenderItem(RenderArray.RenderTypes.eHeaderCenter, New RenderArray.RenderText() {New RenderArray.RenderText(RenderArray.RenderDisplayClass.eArabic, ArabicData.TransliterateFromBuckwalter("suwrapu " + CachedData.IslamData.QuranChapters(CInt(ChapterNode.Attributes.GetNamedItem("index").Value) - 1).Name + " ")), New RenderArray.RenderText(RenderArray.RenderDisplayClass.eTransliteration, Arabic.TransliterateToScheme(ArabicData.TransliterateFromBuckwalter("suwrapu " + CachedData.IslamData.QuranChapters(CInt(ChapterNode.Attributes.GetNamedItem("index").Value) - 1).Name + " "), SchemeType, Scheme).Trim()), New RenderArray.RenderText(RenderArray.RenderDisplayClass.eLTR, "Chapter " + TanzilReader.GetChapterEName(ChapterNode) + " ")}))
                Renderer.Items.Add(New RenderArray.RenderItem(RenderArray.RenderTypes.eHeaderRight, New RenderArray.RenderText() {New RenderArray.RenderText(RenderArray.RenderDisplayClass.eArabic, ArabicData.TransliterateFromBuckwalter("rukuwEaAtuhaA " + ChapterNode.Attributes.GetNamedItem("rukus").Value + " ")), New RenderArray.RenderText(RenderArray.RenderDisplayClass.eTransliteration, Arabic.TransliterateToScheme(ArabicData.TransliterateFromBuckwalter("rukuwEaAtuhaA " + ChapterNode.Attributes.GetNamedItem("rukus").Value + " "), SchemeType, Scheme).Trim()), New RenderArray.RenderText(RenderArray.RenderDisplayClass.eLTR, "Rukus " + ChapterNode.Attributes.GetNamedItem("rukus").Value + " ")}))
                For Verse = 0 To QuranText(Chapter).Length - 1
                    Dim Items As New Collections.Generic.List(Of RenderArray.RenderItem)
                    Text = String.Empty
                    'hizb symbols not needed as Quranic text already contains them
                    'If BaseChapter + Chapter <> 1 AndAlso TanzilReader.IsQuarterStart(BaseChapter + Chapter, CInt(IIf(Chapter = 0, BaseVerse, 1)) + Verse) Then
                    '    Text += ArabicData.TransliterateFromBuckwalter("B")
                    '    Items.Add(New RenderArray.RenderItem(RenderArray.RenderTypes.eText, New RenderArray.RenderText() {New RenderArray.RenderText(RenderArray.RenderDisplayClass.eArabic, ArabicData.TransliterateFromBuckwalter("B"))}))
                    'End If
                    If CInt(IIf(Chapter = 0, BaseVerse, 1)) + Verse = 1 Then
                        Node = GetTextVerse(GetTextChapter(CachedData.XMLDocMain, BaseChapter + Chapter), 1).Attributes.GetNamedItem("bismillah")
                        If Not Node Is Nothing Then
                            Renderer.Items.Add(New RenderArray.RenderItem(RenderArray.RenderTypes.eText, New RenderArray.RenderText() {New RenderArray.RenderText(RenderArray.RenderDisplayClass.eArabic, Node.Value + " "), New RenderArray.RenderText(RenderArray.RenderDisplayClass.eTransliteration, Arabic.TransliterateToScheme(Node.Value, SchemeType, Scheme).Trim()), New RenderArray.RenderText(DirectCast(IIf(IsTranslationTextLTR(TranslationIndex), RenderArray.RenderDisplayClass.eLTR, RenderArray.RenderDisplayClass.eRTL), RenderArray.RenderDisplayClass), TanzilReader.GetTranslationVerse(Lines, 1, 1))}))
                        End If
                    End If
                    Dim Words As String() = If(QuranText(Chapter)(Verse) Is Nothing, {}, QuranText(Chapter)(Verse).Split(" "c))
                    Dim TranslitWords As String() = Arabic.TransliterateToScheme(QuranText(Chapter)(Verse), SchemeType, Scheme).Split(" "c)
                    Dim PauseMarks As Integer = 0
                    For Count As Integer = 0 To Words.Length - 1
                        'handle start/end words here which have space placeholders
                        If Words(Count).Length = 1 AndAlso _
                            Words(Count)(0) = ChrW(0) Then
                            PauseMarks += 1
                        ElseIf Words(Count).Length = 1 AndAlso _
                            (ArabicData.IsStop(ArabicData.FindLetterBySymbol(Words(Count)(0))) Or Words(Count)(0) = ArabicData.ArabicStartOfRubElHizb Or Words(Count)(0) = ArabicData.ArabicPlaceOfSajdah) Then
                            PauseMarks += 1
                            Items.Add(New RenderArray.RenderItem(RenderArray.RenderTypes.eText, New RenderArray.RenderText() {New RenderArray.RenderText(RenderArray.RenderDisplayClass.eArabic, " " + Words(Count)), New RenderArray.RenderText(RenderArray.RenderDisplayClass.eTransliteration, TranslitWords(Count))}))
                        ElseIf Words(Count).Length <> 0 Then
                            Items.Add(New RenderArray.RenderItem(RenderArray.RenderTypes.eText, New RenderArray.RenderText() {New RenderArray.RenderText(RenderArray.RenderDisplayClass.eArabic, Words(Count)), New RenderArray.RenderText(RenderArray.RenderDisplayClass.eTransliteration, TranslitWords(Count)), New RenderArray.RenderText(DirectCast(IIf(IsTranslationTextLTR(TranslationIndex), RenderArray.RenderDisplayClass.eLTR, RenderArray.RenderDisplayClass.eRTL), RenderArray.RenderDisplayClass), TanzilReader.GetW4WTranslationVerse(W4WLines, BaseChapter + Chapter, CInt(IIf(Chapter = 0, BaseVerse, 1)) + Verse, Count - PauseMarks))}))
                        End If
                    Next
                    Text += QuranText(Chapter)(Verse) + " "
                    If TanzilReader.IsSajda(BaseChapter + Chapter, CInt(IIf(Chapter = 0, BaseVerse, 1)) + Verse) Then
                        'Sajda markers are already in the text
                        'Text += ArabicData.TransliterateFromBuckwalter("R")
                        'Items.Add(New RenderArray.RenderItem(RenderArray.RenderTypes.eText, New RenderArray.RenderText() {New RenderArray.RenderText(RenderArray.RenderDisplayClass.eArabic, ArabicData.TransliterateFromBuckwalter("R"))}))
                    End If
                    Text += ArabicData.TransliterateFromBuckwalter("=" + CStr(CInt(IIf(Chapter = 0, BaseVerse, 1)) + Verse)) + " "
                    Items.Add(New RenderArray.RenderItem(RenderArray.RenderTypes.eText, New RenderArray.RenderText() {New RenderArray.RenderText(RenderArray.RenderDisplayClass.eArabic, ArabicData.TransliterateFromBuckwalter("=" + CStr(CInt(IIf(Chapter = 0, BaseVerse, 1)) + Verse))), New RenderArray.RenderText(DirectCast(IIf(IsTranslationTextLTR(TranslationIndex), RenderArray.RenderDisplayClass.eLTR, RenderArray.RenderDisplayClass.eRTL), RenderArray.RenderDisplayClass), "(" + CStr(CInt(IIf(Chapter = 0, BaseVerse, 1)) + Verse) + ")")}))
                    'Text += ArabicData.TransliterateFromBuckwalter("(" + CStr(IIf(Chapter = 0, BaseVerse, 1) + Verse) + ") ")
                    Renderer.Items.Add(New RenderArray.RenderItem(RenderArray.RenderTypes.eText, New RenderArray.RenderText() {New RenderArray.RenderText(RenderArray.RenderDisplayClass.eNested, Items), New RenderArray.RenderText(RenderArray.RenderDisplayClass.eArabic, Text), New RenderArray.RenderText(RenderArray.RenderDisplayClass.eTransliteration, Arabic.TransliterateToScheme(QuranText(Chapter)(Verse) + " " + ArabicData.TransliterateFromBuckwalter("=" + CStr(CInt(IIf(Chapter = 0, BaseVerse, 1)) + Verse)) + " ", SchemeType, Scheme).Trim()), New RenderArray.RenderText(DirectCast(IIf(IsTranslationTextLTR(TranslationIndex), RenderArray.RenderDisplayClass.eLTR, RenderArray.RenderDisplayClass.eRTL), RenderArray.RenderDisplayClass), "(" + CStr(CInt(IIf(Chapter = 0, BaseVerse, 1)) + Verse) + ") " + TanzilReader.GetTranslationVerse(Lines, BaseChapter + Chapter, CInt(IIf(Chapter = 0, BaseVerse, 1)) + Verse))}))
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
        Dim Scheme As String = If(CInt(HttpContext.Current.Request.QueryString.Get("translitscheme")) >= 2, CachedData.IslamData.TranslitSchemes((CInt(HttpContext.Current.Request.QueryString.Get("translitscheme")) - 2) \ 2).Name, String.Empty)
        Dim Names() As Array = Array.ConvertAll(Utility.GetChildNodes("sura", Utility.GetChildNode("suras", CachedData.XMLDocInfo.DocumentElement.ChildNodes).ChildNodes), Function(Convert As System.Xml.XmlNode) New Object() {Convert.Attributes.GetNamedItem("index").Value + ". " + GetChapterEName(Convert) + " (" + ArabicData.RightToLeftOverride + ArabicData.TransliterateFromBuckwalter("suwrapu " + CachedData.IslamData.QuranChapters(CInt(Convert.Attributes.GetNamedItem("index").Value) - 1).Name) + ArabicData.PopDirectionalFormatting + ")" + If(SchemeType = ArabicData.TranslitScheme.None, String.Empty, " " + Arabic.TransliterateToScheme(ArabicData.TransliterateFromBuckwalter(CachedData.IslamData.QuranChapters(CInt(Convert.Attributes.GetNamedItem("index").Value) - 1).Name), SchemeType, Scheme)), CInt(Convert.Attributes.GetNamedItem("index").Value)})
        Array.Sort(Names, New Utility.CompareNameValueArray)
        Return Names
    End Function
    Public Shared Function GetChapterEName(ByVal ChapterNode As System.Xml.XmlNode) As String
        Return Utility.LoadResourceString("IslamInfo_QuranChapter" + ChapterNode.Attributes.GetNamedItem("index").Value)
    End Function
    Public Shared Function GetChapterNamesByRevelationOrder() As Array()
        Dim SchemeType As ArabicData.TranslitScheme = CType(If(CInt(HttpContext.Current.Request.QueryString.Get("translitscheme")) >= 2, 2 - CInt(HttpContext.Current.Request.QueryString.Get("translitscheme")) Mod 2, CInt(HttpContext.Current.Request.QueryString.Get("translitscheme"))), ArabicData.TranslitScheme)
        Dim Scheme As String = If(CInt(HttpContext.Current.Request.QueryString.Get("translitscheme")) >= 2, CachedData.IslamData.TranslitSchemes((CInt(HttpContext.Current.Request.QueryString.Get("translitscheme")) - 2) \ 2).Name, String.Empty)
        Dim Names() As Array = Array.ConvertAll(Utility.GetChildNodes("sura", Utility.GetChildNode("suras", CachedData.XMLDocInfo.DocumentElement.ChildNodes).ChildNodes), Function(Convert As System.Xml.XmlNode) New Object() {Convert.Attributes.GetNamedItem("index").Value + ". " + GetChapterEName(Convert) + " (" + ArabicData.RightToLeftOverride + ArabicData.TransliterateFromBuckwalter("suwrapu " + CachedData.IslamData.QuranChapters(CInt(Convert.Attributes.GetNamedItem("index").Value) - 1).Name) + ArabicData.PopDirectionalFormatting + ")" + If(SchemeType = ArabicData.TranslitScheme.None, String.Empty, " " + Arabic.TransliterateToScheme(ArabicData.TransliterateFromBuckwalter(CachedData.IslamData.QuranChapters(CInt(Convert.Attributes.GetNamedItem("index").Value) - 1).Name), SchemeType, Scheme)), CInt(Convert.Attributes.GetNamedItem("order").Value)})
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
        Dim Names() As Array = Array.ConvertAll(Utility.GetChildNodes("juz", Utility.GetChildNode("juzs", CachedData.XMLDocInfo.DocumentElement.ChildNodes).ChildNodes), Function(Convert As System.Xml.XmlNode) New Object() {Convert.Attributes.GetNamedItem("index").Value + " (" + ArabicData.TransliterateFromBuckwalter("juz " + CachedData.IslamData.QuranParts(CInt(Convert.Attributes.GetNamedItem("index").Value) - 1).Name + " ") + ")", CInt(Convert.Attributes.GetNamedItem("index").Value)})
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
        Dim Scheme As String = If(CInt(HttpContext.Current.Request.QueryString.Get("translitscheme")) >= 2, CachedData.IslamData.TranslitSchemes((CInt(HttpContext.Current.Request.QueryString.Get("translitscheme")) - 2) \ 2).Name, String.Empty)
        Dim Translation As String = HttpContext.Current.Request.QueryString.Get("hadithtranslation")
        Return DoGetRenderedText(SchemeType, Scheme, Translation)
    End Function
    Public Shared Function DoGetRenderedText(SchemeType As ArabicData.TranslitScheme, Scheme As String, Translation As String) As RenderArray
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
            Renderer.Items.Add(New RenderArray.RenderItem(RenderArray.RenderTypes.eHeaderLeft, New RenderArray.RenderText() {New RenderArray.RenderText(RenderArray.RenderDisplayClass.eArabic, ArabicData.TransliterateFromBuckwalter("Had~iv " + BookNode.Attributes.GetNamedItem("hadiths").Value + " ")), New RenderArray.RenderText(RenderArray.RenderDisplayClass.eTransliteration, Arabic.TransliterateToScheme(ArabicData.TransliterateFromBuckwalter("Had~iv " + BookNode.Attributes.GetNamedItem("hadiths").Value + " "), SchemeType, Scheme).Trim()), New RenderArray.RenderText(RenderArray.RenderDisplayClass.eLTR, "Hadiths: " + BookNode.Attributes.GetNamedItem("hadiths").Value + " ")}))
            Renderer.Items.Add(New RenderArray.RenderItem(RenderArray.RenderTypes.eHeaderCenter, New RenderArray.RenderText() {New RenderArray.RenderText(RenderArray.RenderDisplayClass.eArabic, ArabicData.TransliterateFromBuckwalter("{lokita`bu " + CStr(BookIndex)) + " " + BookNode.Attributes.GetNamedItem("name").Value + " "), New RenderArray.RenderText(RenderArray.RenderDisplayClass.eTransliteration, Arabic.TransliterateToScheme(ArabicData.TransliterateFromBuckwalter("{lokita`bu " + CStr(BookIndex)) + " " + BookNode.Attributes.GetNamedItem("name").Value + " ", SchemeType, Scheme).Trim()), New RenderArray.RenderText(RenderArray.RenderDisplayClass.eLTR, "Book " + CStr(BookIndex) + ": " + GetBookEName(BookNode, Index) + " ")}))
            'Renderer.Items.Add(New RenderArray.RenderItem(RenderArray.RenderTypes.eHeaderRight, New RenderArray.RenderText() {New RenderArray.RenderText(RenderArray.RenderDisplayClass.eArabic, ArabicData.TransliterateFromBuckwalter("mjld " + Utility.GetChildNode("books", XMLDocInfo(Index).DocumentElement.ChildNodes).ChildNodes.Item(BookIndex).Attributes.GetNamedItem("volume").Value + " ")), New RenderArray.RenderText(RenderArray.RenderDisplayClass.eTransliteration, Arabic.TransliterateToScheme(ArabicData.TransliterateFromBuckwalter("mjld " + Utility.GetChildNode("books", XMLDocInfo(Index).DocumentElement.ChildNodes).ChildNodes.Item(BookIndex).Attributes.GetNamedItem("volume").Value + " "), SchemeType, Scheme).Trim()), New RenderArray.RenderText(RenderArray.RenderDisplayClass.eLTR, "Volume " + Utility.GetChildNode("books", XMLDocInfo(Index).DocumentElement.ChildNodes).ChildNodes.Item(BookIndex).Attributes.GetNamedItem("volume").Value + " ")}))
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
                        Renderer.Items.Add(New RenderArray.RenderItem(RenderArray.RenderTypes.eHeaderLeft, New RenderArray.RenderText() {New RenderArray.RenderText(RenderArray.RenderDisplayClass.eArabic, ArabicData.TransliterateFromBuckwalter("Had~iv " + ChapterNode.Attributes.GetNamedItem("hadiths").Value + " ")), New RenderArray.RenderText(RenderArray.RenderDisplayClass.eTransliteration, Arabic.TransliterateToScheme(ArabicData.TransliterateFromBuckwalter("Had~iv " + ChapterNode.Attributes.GetNamedItem("hadiths").Value + " "), SchemeType, Scheme).Trim()), New RenderArray.RenderText(RenderArray.RenderDisplayClass.eLTR, "Hadiths: " + ChapterNode.Attributes.GetNamedItem("hadiths").Value + " ")}))
                        Renderer.Items.Add(New RenderArray.RenderItem(RenderArray.RenderTypes.eHeaderCenter, New RenderArray.RenderText() {New RenderArray.RenderText(RenderArray.RenderDisplayClass.eArabic, ArabicData.TransliterateFromBuckwalter("bAb " + CStr(ChapterIndex)) + " " + ChapterNode.Attributes.GetNamedItem("name").Value + " "), New RenderArray.RenderText(RenderArray.RenderDisplayClass.eTransliteration, Arabic.TransliterateToScheme(ArabicData.TransliterateFromBuckwalter("bAb " + CStr(ChapterIndex)) + " " + ChapterNode.Attributes.GetNamedItem("name").Value + " ", SchemeType, Scheme).Trim()), New RenderArray.RenderText(RenderArray.RenderDisplayClass.eLTR, "Chapter " + CStr(ChapterIndex) + ": " + Utility.DefaultValue(Utility.LoadResourceString("IslamInfo_" + CachedData.IslamData.Collections(Index).FileName + "Book" + BookNode.Attributes.GetNamedItem("index").Value + "Chapter" + ChapterNode.Attributes.GetNamedItem("index").Value), String.Empty) + " ")}))
                    End If
                    SubChapterIndex = -1
                End If
                'Handle missing or excess subchapter indexes
                If SubChapterIndex <> CInt(HadithText(Hadith)(2)) Then
                    SubChapterIndex = CInt(HadithText(Hadith)(2))
                    If Not ChapterNode Is Nothing Then
                        SubChapterNode = GetSubChapterByIndex(ChapterNode, SubChapterIndex)
                        If Not SubChapterNode Is Nothing Then
                            Renderer.Items.Add(New RenderArray.RenderItem(RenderArray.RenderTypes.eHeaderLeft, New RenderArray.RenderText() {New RenderArray.RenderText(RenderArray.RenderDisplayClass.eArabic, ArabicData.TransliterateFromBuckwalter("Had~iv " + SubChapterNode.Attributes.GetNamedItem("hadiths").Value + " ")), New RenderArray.RenderText(RenderArray.RenderDisplayClass.eTransliteration, Arabic.TransliterateToScheme(ArabicData.TransliterateFromBuckwalter("Had~iv " + SubChapterNode.Attributes.GetNamedItem("hadiths").Value + " "), SchemeType, Scheme).Trim()), New RenderArray.RenderText(RenderArray.RenderDisplayClass.eLTR, "Hadiths: " + SubChapterNode.Attributes.GetNamedItem("hadiths").Value + " ")}))
                            Renderer.Items.Add(New RenderArray.RenderItem(RenderArray.RenderTypes.eHeaderCenter, New RenderArray.RenderText() {New RenderArray.RenderText(RenderArray.RenderDisplayClass.eArabic, ArabicData.TransliterateFromBuckwalter("bAb " + CStr(SubChapterIndex)) + " " + SubChapterNode.Attributes.GetNamedItem("name").Value + " "), New RenderArray.RenderText(RenderArray.RenderDisplayClass.eTransliteration, Arabic.TransliterateToScheme(ArabicData.TransliterateFromBuckwalter("bAb " + CStr(SubChapterIndex)) + " " + SubChapterNode.Attributes.GetNamedItem("name").Value + " ", SchemeType, Scheme).Trim()), New RenderArray.RenderText(RenderArray.RenderDisplayClass.eLTR, "Sub-Chapter " + CStr(SubChapterIndex) + ": " + Utility.DefaultValue(Utility.LoadResourceString("IslamInfo_" + CachedData.IslamData.Collections(Index).FileName + "Book" + BookNode.Attributes.GetNamedItem("index").Value + "Chapter" + ChapterNode.Attributes.GetNamedItem("index").Value + "Subchapter" + SubChapterNode.Attributes.GetNamedItem("index").Value), String.Empty) + " ")}))
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
                'ArabicData.TransliterateFromBuckwalter("(" + HadithText(Hadith)(0) + ") ")
                Renderer.Items.Add(New RenderArray.RenderItem(RenderArray.RenderTypes.eText, New RenderArray.RenderText() {New RenderArray.RenderText(RenderArray.RenderDisplayClass.eArabic, ArabicData.TransliterateFromBuckwalter(CStr(HadithText(Hadith)(3)) + " " + "=" + CStr(HadithText(Hadith)(0))) + " "), New RenderArray.RenderText(RenderArray.RenderDisplayClass.eTransliteration, Arabic.TransliterateToScheme(CStr(HadithText(Hadith)(3)) + " " + ArabicData.TransliterateFromBuckwalter("=" + CStr(HadithText(Hadith)(0))) + " ", SchemeType, Scheme).Trim()), New RenderArray.RenderText(DirectCast(IIf(IsTranslationTextLTR(Index, Translation), RenderArray.RenderDisplayClass.eLTR, RenderArray.RenderDisplayClass.eRTL), RenderArray.RenderDisplayClass), "(" + CStr(HadithText(Hadith)(0)) + ") " + HadithTranslation)}))
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