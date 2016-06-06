Option Explicit On
Option Strict On
Imports System.Threading.Tasks
Imports IslamMetadata
Imports XMLRender
Public Class InitClass
    Implements Utility.IInitClass
    Private _PortableMethods As PortableMethods
    Private ArbData As ArabicData
    Public Arb As Arabic
    Public ChData As CachedData
    Public TR As TanzilReader
    Public DocBuild As DocBuilder
    Public Sub New(NewPortableMethods As PortableMethods, NewArbData As ArabicData)
        _PortableMethods = NewPortableMethods
        ArbData = NewArbData
    End Sub
    Public Async Function Init() As Task Implements Utility.IInitClass.Init
        Arb = New IslamMetadata.Arabic(_PortableMethods, ArbData)
        ChData = New IslamMetadata.CachedData(_PortableMethods, ArbData, Arb)
        Await ChData.Init(False, True)
        Await Arb.Init(ChData)
        TR = New IslamMetadata.TanzilReader(_PortableMethods, Arb, ArbData, ChData)
        DocBuild = New IslamMetadata.DocBuilder(_PortableMethods, Arb, ArbData, ChData)
    End Function
    Public Function LookupObject(ClassName As String) As Object Implements Utility.IInitClass.LookupObject
        If ClassName = "Arabic" Then
            Return Arb
        ElseIf ClassName = "TanzilReader" Then
            Return TR
        ElseIf ClassName = "DocBuilder" Then
            Return DocBuild
        Else
            Return Nothing
        End If
    End Function
    Public Function GetDependency() As Nullable(Of KeyValuePair(Of String, Utility.IInitClass)) Implements Utility.IInitClass.GetDependency
        Return Nothing
    End Function
End Class
Public Class PrayerTime
    Private Shared _PortableMethods As PortableMethods
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
            Loop While Count <= Calendar.GetDaysInMonth(Calendar.GetYear(TimeNow), Calendar.GetMonth(TimeNow)) AndAlso
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
        RetArray(2) = New String() {_PortableMethods.LoadResourceString("IslamInfo_Date"), _PortableMethods.LoadResourceString("IslamInfo_Day"), _PortableMethods.LoadResourceString("IslamInfo_PrayTime6"), _PortableMethods.LoadResourceString("IslamInfo_PrayTime1"), _PortableMethods.LoadResourceString("IslamInfo_PrayTime7"), _PortableMethods.LoadResourceString("IslamInfo_PrayTime2"), _PortableMethods.LoadResourceString("IslamInfo_PrayTime3"), _PortableMethods.LoadResourceString("IslamInfo_PrayTime8"), _PortableMethods.LoadResourceString("IslamInfo_PrayTime4"), _PortableMethods.LoadResourceString("IslamInfo_PrayTime5"), _PortableMethods.LoadResourceString("IslamInfo_PrayTime9")}
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
        Public Sub New(Scheme As String, Arb As Arabic)
            _Scheme = Scheme
            _Arb = Arb
        End Sub
        Private _Scheme As String
        Private _Arb As Arabic
        Public Function Compare(ByVal x As Object, ByVal y As Object) As Integer _
            Implements Collections.IComparer.Compare
            Compare = _Arb.GetSchemeValueFromSymbol(DirectCast(x, ArabicData.ArabicSymbol), _Scheme).Length -
                _Arb.GetSchemeValueFromSymbol(DirectCast(y, ArabicData.ArabicSymbol), _Scheme).Length
            If Compare = 0 Then Compare = _Arb.GetSchemeValueFromSymbol(DirectCast(x, ArabicData.ArabicSymbol), _Scheme).CompareTo(_Arb.GetSchemeValueFromSymbol(DirectCast(y, ArabicData.ArabicSymbol), _Scheme))
        End Function
    End Class
    Private _PortableMethods As PortableMethods
    Private ArbData As ArabicData
    Private ChData As CachedData
    Public Sub New(NewPortableMethods As PortableMethods, NewArbData As ArabicData)
        _PortableMethods = NewPortableMethods
        ArbData = NewArbData
    End Sub
    Public Async Function Init(NewChData As CachedData) As Threading.Tasks.Task
        ChData = NewChData
        If _SchemeTable Is Nothing Then
            _SchemeTable = New Dictionary(Of String, IslamData.TranslitScheme)
            'Await ChData.Init(False)
            Dim Count As Integer
            For Count = 0 To ChData.IslamData.TranslitSchemes.Length - 1
                _SchemeTable.Add(ChData.IslamData.TranslitSchemes(Count).Name, ChData.IslamData.TranslitSchemes(Count))
            Next
        End If
        If _BuckwalterMap Is Nothing Then
            Dim BuckwalterBytes As Byte() = Await _PortableMethods.DiskCache.GetCacheItem("BuckwalterMap", DateTime.MinValue)
            If Not BuckwalterBytes Is Nothing Then
                _BuckwalterMap = CType((New Runtime.Serialization.DataContractSerializer(GetType(Dictionary(Of Char, Integer)))).ReadObject(New IO.MemoryStream(BuckwalterBytes)), Dictionary(Of Char, Integer))
            Else
                _BuckwalterMap = New Dictionary(Of Char, Integer)
                For Index = 0 To ArbData.ArabicLetters.Length - 1
                    If (GetSchemeValueFromSymbol(ArbData.ArabicLetters(Index), "ExtendedBuckwalter")).Length <> 0 Then
                        _BuckwalterMap.Add((GetSchemeValueFromSymbol(ArbData.ArabicLetters(Index), "ExtendedBuckwalter")).Chars(0), Index)
                    End If
                Next
                Dim MemStream As New IO.MemoryStream
                Dim Ser As New Runtime.Serialization.DataContractSerializer(GetType(Dictionary(Of Char, Integer)))
                Ser.WriteObject(MemStream, _BuckwalterMap)
                Await _PortableMethods.DiskCache.CacheItem("BuckwalterMap", DateTime.Now, MemStream.ToArray())
                'MemStream.Close()
                MemStream = Nothing
            End If
        End If
        If _Letters Is Nothing Then
            Dim Letters(ArbData.ArabicLetters.Length - 1) As ArabicData.ArabicSymbol
            ArbData.ArabicLetters.CopyTo(Letters, 0)
            _Letters = New Dictionary(Of String, Dictionary(Of Char, Integer))
            For Index = 0 To ChData.IslamData.TranslitSchemes.Length - 1
                _Letters.Add(ChData.IslamData.TranslitSchemes(Index).Name, New Dictionary(Of Char, Integer))
                Array.Sort(Letters, New StringLengthComparer(ChData.IslamData.TranslitSchemes(Index).Name, Me))
                For Count = 0 To Letters.Length - 1
                    _Letters(ChData.IslamData.TranslitSchemes(Index).Name).Add(Letters(Count).Symbol, Count)
                Next
            Next
        End If
        If ErrorRegExDict.Count = 0 Or RegExDict.Count = 0 Then Await ChData.Init(True)
        If ErrorRegExDict.Count = 0 Then
            For Count = 0 To ChData.RuleTranslations("ErrorCheck").Length - 1
                ErrorRegExDict.Add(ChData.RuleTranslations("ErrorCheck")(Count).Name + CStr(Count), New System.Text.RegularExpressions.Regex(ChData.RuleTranslations("ErrorCheck")(Count).Match))
            Next
        End If
        If RegExDict.Count = 0 Then
            For RuleSetCount = 0 To ChData.IslamData.MetaRules.Length - 1
                For Count = 0 To ChData.RuleMetas(ChData.IslamData.MetaRules(RuleSetCount).Name).Rules.Length - 1
                    If Not ChData.RuleMetas(ChData.IslamData.MetaRules(RuleSetCount).Name).Rules(Count).Evaluator Is Nothing Then
                        RegExDict.Add(String.Join(String.Empty, ChData.IslamData.MetaRules(RuleSetCount).Name, "_", ChData.RuleMetas(ChData.IslamData.MetaRules(RuleSetCount).Name).Rules(Count).Name, CStr(Count)), New System.Text.RegularExpressions.Regex(ChData.RuleMetas(ChData.IslamData.MetaRules(RuleSetCount).Name).Rules(Count).Match))
                    End If
                Next
            Next
        End If
        If _NounIDs Is Nothing Then
            _NounIDs = New Dictionary(Of String, List(Of IslamData.GrammarSet.GrammarNoun))
            For Count = 0 To ChData.IslamData.Grammar.Nouns.Length - 1
                If Not _NounIDs.ContainsKey(ChData.IslamData.Grammar.Nouns(Count).TranslationID) Then
                    _NounIDs.Add(ChData.IslamData.Grammar.Nouns(Count).TranslationID, New List(Of IslamData.GrammarSet.GrammarNoun) From {ChData.IslamData.Grammar.Nouns(Count)})
                Else
                    'Debug.Print("Duplicate Noun ID: " + CachedData.IslamData.Grammar.Nouns(Count).TranslationID)
                End If
                Dim Noun As IslamData.GrammarSet.GrammarNoun = ChData.IslamData.Grammar.Nouns(Count)
                If Not ChData.IslamData.Grammar.Nouns(Count).Grammar Is Nothing AndAlso ChData.IslamData.Grammar.Nouns(Count).Grammar.Length <> 0 Then
                    For Each Str As String In ChData.IslamData.Grammar.Nouns(Count).Grammar.Split(","c)
                        If Not _NounIDs.ContainsKey(Str) Then
                            _NounIDs.Add(Str, New List(Of IslamData.GrammarSet.GrammarNoun))
                        End If
                        _NounIDs(Str).Add(Noun)
                    Next
                End If
            Next
        End If
        If _TransformIDs Is Nothing Then
            _TransformIDs = New Dictionary(Of String, List(Of IslamData.GrammarSet.GrammarTransform))
            For Count = 0 To ChData.IslamData.Grammar.Transforms.Length - 1
                If Not _TransformIDs.ContainsKey(ChData.IslamData.Grammar.Transforms(Count).TranslationID) Then
                    _TransformIDs.Add(ChData.IslamData.Grammar.Transforms(Count).TranslationID, New List(Of IslamData.GrammarSet.GrammarTransform) From {ChData.IslamData.Grammar.Transforms(Count)})
                Else
                    'Debug.Print("Duplicate Transform ID: " + CachedData.IslamData.Grammar.Transforms(Count).TranslationID)
                End If
                Dim Transform As IslamData.GrammarSet.GrammarTransform = ChData.IslamData.Grammar.Transforms(Count)
                If Not ChData.IslamData.Grammar.Transforms(Count).Grammar Is Nothing AndAlso ChData.IslamData.Grammar.Transforms(Count).Grammar.Length <> 0 Then
                    For Each GroupStr As String In ChData.IslamData.Grammar.Transforms(Count).Grammar.Split(","c)
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
        If _ParticleIDs Is Nothing Then
            _ParticleIDs = New Dictionary(Of String, List(Of IslamData.GrammarSet.GrammarParticle))
            For Count = 0 To ChData.IslamData.Grammar.Particles.Length - 1
                If Not _ParticleIDs.ContainsKey(ChData.IslamData.Grammar.Particles(Count).TranslationID) Then
                    _ParticleIDs.Add(ChData.IslamData.Grammar.Particles(Count).TranslationID, New List(Of IslamData.GrammarSet.GrammarParticle) From {ChData.IslamData.Grammar.Particles(Count)})
                Else
                    'Debug.Print("Duplicate Particle ID: " + CachedData.IslamData.Grammar.Particles(Count).TranslationID)
                End If
                Dim Particle As IslamData.GrammarSet.GrammarParticle = ChData.IslamData.Grammar.Particles(Count)
                If Not ChData.IslamData.Grammar.Particles(Count).Grammar Is Nothing AndAlso ChData.IslamData.Grammar.Particles(Count).Grammar.Length <> 0 Then
                    For Each Str As String In ChData.IslamData.Grammar.Particles(Count).Grammar.Split(","c)
                        If Not _ParticleIDs.ContainsKey(Str) Then
                            _ParticleIDs.Add(Str, New List(Of IslamData.GrammarSet.GrammarParticle))
                        End If
                        _ParticleIDs(Str).Add(Particle)
                    Next
                End If
            Next
        End If
        If _VerbIDs Is Nothing Then
            _VerbIDs = New Dictionary(Of String, List(Of IslamData.GrammarSet.GrammarVerb))
            For Count = 0 To ChData.IslamData.Grammar.Verbs.Length - 1
                If Not _VerbIDs.ContainsKey(ChData.IslamData.Grammar.Verbs(Count).TranslationID) Then
                    _VerbIDs.Add(ChData.IslamData.Grammar.Verbs(Count).TranslationID, New List(Of IslamData.GrammarSet.GrammarVerb) From {ChData.IslamData.Grammar.Verbs(Count)})
                Else
                    'Debug.Print("Duplicate Verb ID: " + CachedData.IslamData.Grammar.Verbs(Count).TranslationID)
                End If
                Dim Verb As IslamData.GrammarSet.GrammarVerb = ChData.IslamData.Grammar.Verbs(Count)
                If Not ChData.IslamData.Grammar.Verbs(Count).Grammar Is Nothing AndAlso ChData.IslamData.Grammar.Verbs(Count).Grammar.Length <> 0 Then
                    For Each Str As String In ChData.IslamData.Grammar.Verbs(Count).Grammar.Split(","c)
                        If Not _VerbIDs.ContainsKey(Str) Then
                            _VerbIDs.Add(Str, New List(Of IslamData.GrammarSet.GrammarVerb))
                        End If
                        _VerbIDs(Str).Add(Verb)
                    Next
                End If
            Next
        End If
    End Function
    Public Function GetRecitationSymbol(Index As Integer, SchemeType As ArabicData.TranslitScheme, Scheme As String) As String
        Return ArbData.ArabicLetters(Index).UnicodeName + " (" + ArbData.FixStartingCombiningSymbol(ArbData.ArabicLetters(Index).Symbol) + ")" + If(SchemeType = ArabicData.TranslitScheme.None, String.Empty, " " + TransliterateToScheme(ArbData.ArabicLetters(Index).Symbol, SchemeType, Scheme, Arabic.FilterMetadataStops(ArbData.ArabicLetters(Index).Symbol, GetMetarules(ArbData.ArabicLetters(Index).Symbol, ChData.RuleMetas("Normal")), Nothing)))
    End Function
    Public Function GetRecitationSymbols(SchemeType As ArabicData.TranslitScheme, Scheme As String) As Array()
        Return New List(Of Object())(Linq.Enumerable.Select(ChData.RecitationSymbols, Function(Ch As String) New Object() {ArbData.ArabicLetters(ArbData.FindLetterBySymbol(Ch.Chars(0))).UnicodeName + " (" + ArbData.FixStartingCombiningSymbol(Ch) + ")" + If(SchemeType = ArabicData.TranslitScheme.None, String.Empty, " " + TransliterateToScheme(Ch, SchemeType, Scheme, Arabic.FilterMetadataStops(Ch, GetMetarules(Ch, ChData.RuleMetas("Normal")), Nothing))), ArbData.FindLetterBySymbol(Ch.Chars(0))})).ToArray()
    End Function
    Private _BuckwalterMap As Dictionary(Of Char, Integer)
    Public Function BuckwalterMap() As Dictionary(Of Char, Integer)
        Return _BuckwalterMap
    End Function
    Public Function TransliterateFromBuckwalter(ByVal Buckwalter As String) As String
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
                    ArabicString.Append(ArbData.ArabicLetters(BuckwalterMap.Item(Buckwalter(Count))).Symbol)
                Else
                    ArabicString.Append(Buckwalter(Count))
                End If
            End If
        Next
        Return ArabicString.ToString()
    End Function
    Public Function TransliterateToScheme(ByVal ArabicString As String, SchemeType As ArabicData.TranslitScheme, Scheme As String, MetadataList As Generic.List(Of RuleMetadata)) As String
        If SchemeType = ArabicData.TranslitScheme.LearningMode Then
            Return TransliterateWithRules(ArabicString, Scheme, True, MetadataList)
        ElseIf SchemeType = ArabicData.TranslitScheme.RuleBased Then
            Return TransliterateWithRules(ArabicString, Scheme, False, MetadataList)
        ElseIf SchemeType = ArabicData.TranslitScheme.Literal Then
            Return TransliterateToLiteral(ArabicString, Scheme)
        Else
            Return New String(New List(Of Char)(Linq.Enumerable.Where(ArabicString.ToCharArray(), Function(Check As Char) Check = " "c)).ToArray())
        End If
    End Function
    Shared _SchemeTable As Dictionary(Of String, IslamData.TranslitScheme)
    Public ReadOnly Property SchemeTable() As Dictionary(Of String, IslamData.TranslitScheme)
        Get
            Return _SchemeTable
        End Get
    End Property
    Public Function GetSchemeSpecialValue(Str As String, Index As Integer, Scheme As String) As String
        If Not SchemeTable.ContainsKey(Scheme) Then Return String.Empty
        Dim Sch As IslamData.TranslitScheme = SchemeTable(Scheme)
        Return Sch.SpecialLetters(Index).Replace("&first;", If(Str = ChData.ArabicSpecialLetters(Index), "*", GetSchemeValueFromSymbol(ArbData.ArabicLetters(ArbData.FindLetterBySymbol(Str(0))), Scheme)))
    End Function
    Public Function GetSchemeSpecialFromMatch(Str As String, bExp As Boolean) As Integer
        If bExp Then
            For Count = 0 To ChData.ArabicSpecialLetters.Length - 1
                If System.Text.RegularExpressions.Regex.Match(ChData.ArabicSpecialLetters(Count), Str).Success Then Return Count
            Next
        Else
            If Array.IndexOf(ChData.ArabicSpecialLetters, Str) <> -1 Then
                Return Array.IndexOf(ChData.ArabicSpecialLetters, Str)
            End If
        End If
        Return -1
    End Function
    Public Function GetSchemeLongVowel(Str As String, Scheme As String) As Integer
        If Not SchemeTable.ContainsKey(Scheme) Then Return -1
        Dim Sch As IslamData.TranslitScheme = SchemeTable(Scheme)
        If Array.IndexOf(ChData.ArabicMultis, Str) <> -1 Then
            Return Array.IndexOf(ChData.ArabicMultis, Str)
        End If
        Return -1
    End Function
    Public Function GetSchemeLongVowelFromString(Str As String, Scheme As String) As String
        If Not SchemeTable.ContainsKey(Scheme) Then Return String.Empty
        Dim Sch As IslamData.TranslitScheme = SchemeTable(Scheme)
        If Array.IndexOf(ChData.ArabicMultis, Str) <> -1 Then
            Return Sch.Multis(Array.IndexOf(ChData.ArabicMultis, Str))
        End If
        Return String.Empty
    End Function
    Public Function GetSchemeGutteralFromString(Str As String, Scheme As String, Leading As Boolean) As String
        If Not SchemeTable.ContainsKey(Scheme) Then Return String.Empty
        Dim Sch As IslamData.TranslitScheme = SchemeTable(Scheme)
        If Array.IndexOf(ChData.ArabicLeadingGutterals, Str) <> -1 Then
            Return Sch.Gutterals(Array.IndexOf(ChData.ArabicLeadingGutterals, Str) + If(Leading, ChData.ArabicLeadingGutterals.Length, 0))
        End If
        Return String.Empty
    End Function
    Public Function GetSchemeValueFromSymbol(Symbol As ArabicData.ArabicSymbol, Scheme As String) As String
        If Not SchemeTable.ContainsKey(Scheme) Then Return String.Empty
        Dim Sch As IslamData.TranslitScheme = SchemeTable(Scheme)
        If Array.IndexOf(ChData.ArabicLettersInOrder, CStr(Symbol.Symbol)) <> -1 Then
            Return Sch.Alphabet(Array.IndexOf(ChData.ArabicLettersInOrder, CStr(Symbol.Symbol)))
        ElseIf Array.IndexOf(ChData.ArabicHamzas, CStr(Symbol.Symbol)) <> -1 Then
            Return Sch.Hamza(Array.IndexOf(ChData.ArabicHamzas, CStr(Symbol.Symbol)))
        ElseIf Array.IndexOf(ChData.ArabicVowels, CStr(Symbol.Symbol)) <> -1 Then
            Return Sch.Vowels(Array.IndexOf(ChData.ArabicVowels, CStr(Symbol.Symbol)))
        ElseIf Array.IndexOf(ChData.ArabicTajweed, CStr(Symbol.Symbol)) <> -1 Then
            Return Sch.Tajweed(Array.IndexOf(ChData.ArabicTajweed, CStr(Symbol.Symbol)))
        ElseIf Array.IndexOf(ChData.ArabicSilent, CStr(Symbol.Symbol)) <> -1 Then
            Return Sch.Silent(Array.IndexOf(ChData.ArabicSilent, CStr(Symbol.Symbol)))
        ElseIf Array.IndexOf(ChData.ArabicPunctuation, CStr(Symbol.Symbol)) <> -1 Then
            Return Sch.Punctuation(Array.IndexOf(ChData.ArabicPunctuation, CStr(Symbol.Symbol)))
        ElseIf Array.IndexOf(ChData.ArabicNums, CStr(Symbol.Symbol)) <> -1 Then
            Return Sch.Numbers(Array.IndexOf(ChData.ArabicNums, CStr(Symbol.Symbol)))
        ElseIf Array.IndexOf(ChData.NonArabicLetters, CStr(Symbol.Symbol)) <> -1 Then
            Return Sch.NonArabic(Array.IndexOf(ChData.NonArabicLetters, CStr(Symbol.Symbol)))
        End If
        Return String.Empty
    End Function
    Public Function SchemeHasValue(Str As String, Scheme As String) As Boolean
        If Not SchemeTable.ContainsKey(Scheme) Then Return False
        Dim Sch As IslamData.TranslitScheme = SchemeTable(Scheme)
        For Count = 0 To Sch.Alphabet.Length - 1
            If Sch.Alphabet(Count) = Str Then Return True
        Next
        Return False
    End Function
    Shared _Letters As Dictionary(Of String, Dictionary(Of Char, Integer))
    Public Function GetSortedLetters(Scheme As String) As Dictionary(Of Char, Integer)
        Return _Letters(Scheme)
    End Function
    Public Function TransliterateToLiteral(ByVal ArabicString As String, Scheme As String) As String
        Dim LiteralString As New System.Text.StringBuilder
        Dim Count As Integer = 0
        If Scheme = String.Empty Then Scheme = "ExtendedBuckwalter"
        While Count <= ArabicString.Length - 1
            If ArabicString(Count) = "\" Then
                Count += 1
                If ArabicString(Count) = "," Then
                    LiteralString.Append(ArabicData.ArabicComma)
                ElseIf ArabicString(Count) = ";" Then
                    LiteralString.Append(ArabicData.ArabicSemicolon)
                ElseIf ArabicString(Count) = "?" Then
                    LiteralString.Append(ArabicData.ArabicQuestionMark)
                Else
                    LiteralString.Append(ArabicString(Count))
                End If
            Else
                If GetSchemeSpecialFromMatch(ArabicString.Substring(Count), False) <> -1 Then
                    LiteralString.Append(GetSchemeSpecialValue(ArabicString.Substring(Count), GetSchemeSpecialFromMatch(ArabicString.Substring(Count), False), Scheme))
                    Count += System.Text.RegularExpressions.Regex.Match(ChData.ArabicSpecialLetters(GetSchemeSpecialFromMatch(ArabicString.Substring(Count), False)), ArabicString.Substring(Count)).Value.Length - 1
                ElseIf ArabicString.Length - Count > 1 AndAlso (GetSchemeLongVowel(ArabicString.Substring(Count, 2), Scheme)) <> -1 Then
                    LiteralString.Append(GetSchemeLongVowelFromString(ArabicString.Substring(Count, 2), Scheme))
                    Count += 1
                ElseIf GetSortedLetters(Scheme).ContainsKey(ArabicString(Count)) Then
                    LiteralString.Append(GetSchemeValueFromSymbol(ArbData.ArabicLetters(ArbData.FindLetterBySymbol(ArabicString(Count))), Scheme))
                Else
                    LiteralString.Append(ArabicString(Count))
                End If
            End If
            Count += 1
        End While
        Return LiteralString.ToString()
    End Function
    Structure RuleMetadata
        Sub New(NewIndex As Integer, NewLength As Integer, ByVal NewType As IslamData.RuleMetaSet.RuleMetadataTranslation.RuleWithArgs(), NewOrigOrder As Integer)
            Index = NewIndex
            Length = NewLength
            Type = NewType
            OrigOrder = NewOrigOrder
        End Sub
        Public Index As Integer
        Public Length As Integer
        Public Type As IslamData.RuleMetaSet.RuleMetadataTranslation.RuleWithArgs()
        Public OrigOrder As Integer
        Public Children As RuleMetadata()
        Public Dependencies As RuleMetadata()
    End Structure
    Public Const SimpleTrailingAlef As String = ArabicData.ArabicLetterAlef + ArabicData.ArabicSmallHighRoundedZero
    Public Const SimpleSuperscriptAlef As String = ArabicData.ArabicLetterSuperscriptAlef
    Public Const UthmaniShortVowelsBeforeLongVowelsSuperscriptAlef As String = ArabicData.ArabicFatha + ArabicData.ArabicLetterSuperscriptAlef
    Public Const UthmaniShortVowelsBeforeLongVowelsAlefAlefMaksura As String = ArabicData.ArabicFatha + ArabicData.ArabicLetterAlefMaksura
    Public Const UthmaniShortVowelsBeforeLongVowelsYehAlefMaksura As String = ArabicData.ArabicKasra + ArabicData.ArabicLetterAlefMaksura
    Public Const UthmaniShortVowelsBeforeLongVowelsAlef As String = ArabicData.ArabicFatha + ArabicData.ArabicLetterAlef
    Public Const UthmaniShortVowelsBeforeLongVowelsWaw As String = ArabicData.ArabicDamma + ArabicData.ArabicLetterWaw
    Public Const UthmaniShortVowelsBeforeLongVowelsSmallWaw As String = ArabicData.ArabicDamma + ArabicData.ArabicSmallWaw
    Public Const UthmaniShortVowelsBeforeLongVowelsYeh As String = ArabicData.ArabicKasra + ArabicData.ArabicLetterYeh
    Public Const UthmaniShortVowelsBeforeLongVowelsSmallYeh As String = ArabicData.ArabicKasra + ArabicData.ArabicSmallYeh

    Public Const PrefixPattern As String = ""
    Public Enum MetaRuleFuncs As Integer
        eNone
        eUpperCase
        eSpellNumber
        eSpellLetter
        eSpellLongLetter
        eSpellLongMergedLetter
        eDivideTanween
        eLearningMode
        eObligatory
    End Enum
    Public Enum RuleFuncs As Integer
        eNone
        eLookupLetter
        eLookupLongVowelDipthong
        eLeadingGutteral
        eTrailingGutteral
        eResolveAmbiguity
    End Enum

    Public Delegate Function MetaRuleFunction(Str As String, LearningMode As Boolean) As String()
    Public MetaRuleFunctions As MetaRuleFunction() = {
        Function(Str As String, LearningMode As Boolean) {Str.ToUpper()},
        Function(Str As String, LearningMode As Boolean) {TransliterateFromBuckwalter(ArabicWordFromNumber(CInt(TransliterateToScheme(Str, ArabicData.TranslitScheme.Literal, String.Empty, Nothing)), True, False, False))},
        Function(Str As String, LearningMode As Boolean) {ArabicLetterSpelling(Str, True, False, False)},
        Function(Str As String, LearningMode As Boolean) {ArabicLetterSpelling(Str, True, True, False)},
        Function(Str As String, LearningMode As Boolean) {ArabicLetterSpelling(Str, True, True, True)},
        Function(Str As String, LearningMode As Boolean) {ChData.ArabicFathaDammaKasra(Array.IndexOf(ChData.ArabicTanweens, Str)), ArabicData.ArabicLetterNoon},
        Function(Str As String, LearningMode As Boolean) If(LearningMode, {Str, String.Empty}, {String.Empty, Str}),
        Function(Str As String, LearningMode As Boolean) {Str + "-" + Str + "(-" + Str(0) + ")"}
    }

    Public Delegate Function RuleFunction(Str As String, Scheme As String) As String()
    Public RuleFunctions As RuleFunction() = {
        Function(Str As String, Scheme As String) {GetSchemeValueFromSymbol(ArbData.ArabicLetters(ArbData.FindLetterBySymbol(Str.Chars(0))), Scheme)},
        Function(Str As String, Scheme As String) {GetSchemeLongVowelFromString(Str, Scheme)},
        Function(Str As String, Scheme As String) {GetSchemeGutteralFromString(Str.Remove(Str.Length - 1), Scheme, True) + Str.Chars(Str.Length - 1)},
        Function(Str As String, Scheme As String) {Str.Chars(0) + GetSchemeGutteralFromString(Str.Remove(0, 1), Scheme, False)},
        Function(Str As String, Scheme As String) {If(SchemeHasValue(GetSchemeValueFromSymbol(ArbData.ArabicLetters(ArbData.FindLetterBySymbol(Str.Chars(0))), Scheme) + GetSchemeValueFromSymbol(ArbData.ArabicLetters(ArbData.FindLetterBySymbol(Str.Chars(1))), Scheme), Scheme), Str.Chars(0) + "-" + Str.Chars(1), Str)}
    }
    'Javascript does not support negative or positive lookbehind in regular expressions
    Public Shared RecursiveMetadata As String() = {"spellnumber", "spellletter", "spelllongletter", "spelllongmergedletter"}
    Public Shared AllowZeroLength As String() = {"helperfatha", "helperdamma", "helperkasra", "helperlparen", "helperrparen", "learningmode(helperslash,)", "learningmode(helperlbracket,)", "learningmode(helperrbracket,)", "learningmode(helperfathatan,)", "learningmode(helperteh,)"}
    Public Function IsLetter(Index As Integer) As Boolean
        Return New List(Of String)(Linq.Enumerable.TakeWhile(ChData.ArabicLetters, Function(Str As String) Str <> ArbData.ArabicLetters(Index).Symbol)).Count <> ChData.ArabicLetters.Length
    End Function
    Public Function IsPunctuation(Index As Integer) As Boolean
        Return New List(Of String)(Linq.Enumerable.TakeWhile(ChData.PunctuationSymbols, Function(Str As String) Str <> ArbData.ArabicLetters(Index).Symbol)).Count <> ChData.PunctuationSymbols.Length
    End Function
    Public Function IsStop(Index As Integer) As Boolean
        Return New List(Of String)(Linq.Enumerable.TakeWhile(ChData.ArabicStopLetters, Function(Str As String) Str <> ArbData.ArabicLetters(Index).Symbol)).Count <> ChData.ArabicStopLetters.Length
    End Function
    Public Function IsWhitespace(Index As Integer) As Boolean
        Return New List(Of String)(Linq.Enumerable.TakeWhile(ChData.WhitespaceSymbols, Function(Str As String) Str <> ArbData.ArabicLetters(Index).Symbol)).Count <> ChData.WhitespaceSymbols.Length
    End Function
    Public Function ArabicLetterSpelling(Input As String, Quranic As Boolean, IsLong As Boolean, Merged As Boolean) As String
        Dim Output As New System.Text.StringBuilder
        For Each Ch As Char In Input
            Dim Index As Integer = ArbData.FindLetterBySymbol(Ch)
            If Index <> -1 AndAlso IsLetter(Index) Then
                If Output.Length <> 0 And Not Quranic Then Output.Append(" ")
                Dim Idx As Integer = Array.IndexOf(ChData.ArabicLettersInOrder, CStr(ArbData.ArabicLetters(Index).Symbol))
                Output.Append(If(Quranic, If(IsLong, If(Merged, ChData.ArabicAlphabet(Idx).Remove(ChData.ArabicAlphabet(Idx).Length - 1).Insert(ChData.ArabicAlphabet(Idx).Length - 2, ArabicData.ArabicMaddahAbove).Insert(1, ArabicData.ArabicShadda), ChData.ArabicAlphabet(Idx).Remove(ChData.ArabicAlphabet(Idx).Length - 1).Insert(ChData.ArabicAlphabet(Idx).Length - 2, ArabicData.ArabicMaddahAbove)), ChData.ArabicAlphabet(Idx).Remove(ChData.ArabicAlphabet(Idx).Length - 1)) + If(ChData.ArabicAlphabet(Idx).EndsWith("n"), String.Empty, "o"), ChData.ArabicAlphabet(Idx)))
            ElseIf Index <> -1 AndAlso ArbData.ArabicLetters(Index).Symbol = ArabicData.ArabicMaddahAbove Then
                If Not Quranic Then Output.Append(Ch)
            End If
        Next
        Return TransliterateFromBuckwalter(Output.ToString())
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
    Public Function ApplyColorRules(ByVal ArabicString As String, BreakWords As Boolean, MetadataList As Generic.List(Of RuleMetadata)) As RenderArray.RenderText()()
        Dim Count As Integer
        Dim Index As Integer
        Dim Strings As New Generic.List(Of RenderArray.RenderText)
        Dim RuleIndexes As New List(Of Integer)
        For Count = 0 To ArabicString.Length - 1
            RuleIndexes.Add(0)
        Next
        For Index = 0 To MetadataList.Count - 1
            For Count = 0 To ChData.IslamData.ColorRuleSets(1).ColorRules.Length - 1
                Dim Match As Integer = New List(Of String)(Linq.Enumerable.TakeWhile(ChData.IslamData.ColorRuleSets(1).ColorRules(Count).Match, Function(Str As String) Not Linq.Enumerable.Any(MetadataList(Index).Type, Function(RuleWithArg) RuleWithArg.RuleName = Str))).Count
                If Match <> ChData.IslamData.ColorRuleSets(1).ColorRules(Count).Match.Length Then
                    For MetaCount As Integer = MetadataList(Index).Index To MetadataList(Index).Index + MetadataList(Index).Length - 1
                        If ChData.IslamData.ColorRuleSets(1).ColorRules(RuleIndexes(MetaCount)).Color = &HFF000000 Then RuleIndexes(MetaCount) = Count
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
                If (If((ArabicString(Count) = " "c And BreakWords), Count - Base <> 0, Base <> ArabicString.Length)) Then Renderers.Add(New RenderArray.RenderText(RenderArray.RenderDisplayClass.eArabic, If((ArabicString(Count) = " "c And BreakWords), ArabicString.Substring(Base, Count - Base), ArabicString.Substring(Base))) With {.Clr = ChData.IslamData.ColorRuleSets(1).ColorRules(RuleIndexes(Count) Mod ChData.IslamData.ColorRuleSets(1).ColorRules.Length).Color})
                WordRenderers.Add(Renderers.ToArray())
                Renderers = New List(Of RenderArray.RenderText)
                Base = Count + 1
            ElseIf RuleIndexes(Count) <> RuleIndexes(Count + 1) Then
                Renderers.Add(New RenderArray.RenderText(RenderArray.RenderDisplayClass.eArabic, ArabicString.Substring(Base, Count - Base + 1)) With {.Clr = ChData.IslamData.ColorRuleSets(1).ColorRules(RuleIndexes(Count) Mod ChData.IslamData.ColorRuleSets(1).ColorRules.Length).Color})
                Base = Count + 1
            End If
        Next
        Return WordRenderers.ToArray()
    End Function
    Public Function ArabicWordForLessThanThousand(ByVal Number As Integer, UseClassic As Boolean, UseAlefHundred As Boolean) As String
        Dim Str As String = String.Empty
        Dim HundStr As String = String.Empty
        If Number >= 100 Then
            HundStr = If(UseAlefHundred, ChData.ArabicBaseHundredNumbers((Number \ 100) - 1).Insert(2, "A"), ChData.ArabicBaseHundredNumbers((Number \ 100) - 1))
            If (Number Mod 100) = 0 Then Return HundStr
            Number = Number Mod 100
        End If
        If (Number Mod 10) <> 0 And Number <> 11 And Number <> 12 Then
            Str = ChData.ArabicBaseNumbers((Number Mod 10) - 1)
        End If
        If Number >= 11 AndAlso Number < 20 Then
            If Number = 11 Or Number = 12 Then
                Str += ChData.ArabicBaseExtraNumbers(Number - 11)
            Else
                Str = Str.Remove(Str.Length - 1) + "a"
            End If
            Str += " " + ChData.ArabicBaseTenNumbers(1).Remove(ChData.ArabicBaseTenNumbers(1).Length - 2)
        ElseIf (Number = 0 And Str = String.Empty) Or Number = 10 Or Number >= 20 Then
            Str = If(Str = String.Empty, String.Empty, Str + " " + ChData.ArabicCombiners(0)) + ChData.ArabicBaseTenNumbers(Number \ 10)
        End If
        Return If(UseClassic, If(Str = String.Empty, String.Empty, Str + If(HundStr = String.Empty, String.Empty, " " + ChData.ArabicCombiners(0))) + HundStr, If(HundStr = String.Empty, String.Empty, HundStr + If(Str = String.Empty, String.Empty, " " + ChData.ArabicCombiners(0))) + Str)
    End Function
    Public Function ArabicWordFromNumber(ByVal Number As Long, UseClassic As Boolean, UseAlefHundred As Boolean, UseMilliard As Boolean) As String
        Dim Str As String = String.Empty
        Dim NextStr As String = String.Empty
        Dim CurBase As Integer = 3
        Dim BaseNums As Long() = {1000, 1000000, 1000000000, 1000000000000}
        Dim Bases As String()() = {ChData.ArabicBaseThousandNumbers, ChData.ArabicBaseMillionNumbers, If(UseMilliard, ChData.ArabicBaseMilliardNumbers, ChData.ArabicBaseBillionNumbers), ChData.ArabicBaseTrillionNumbers}
        Do
            If Number >= BaseNums(CurBase) And Number < 2 * BaseNums(CurBase) Then
                NextStr = Bases(CurBase)(0)
            ElseIf Number >= 2 * BaseNums(CurBase) And Number < 3 * BaseNums(CurBase) Then
                NextStr = Bases(CurBase)(1)
            ElseIf Number >= 3 * BaseNums(CurBase) And Number < 10 * BaseNums(CurBase) Then
                NextStr = ChData.ArabicBaseNumbers(CInt(Number \ BaseNums(CurBase) - 1)).Remove(ChData.ArabicBaseNumbers(CInt(Number \ BaseNums(CurBase) - 1)).Length - 1) + "u " + Bases(CurBase)(2)
            ElseIf Number >= 10 * BaseNums(CurBase) And Number < 11 * BaseNums(CurBase) Then
                NextStr = ChData.ArabicBaseTenNumbers(1).Remove(ChData.ArabicBaseTenNumbers(1).Length - 1) + "u " + Bases(CurBase)(2)
            ElseIf Number >= BaseNums(CurBase) Then
                NextStr = ArabicWordForLessThanThousand(CInt((Number \ BaseNums(CurBase)) Mod 100), UseClassic, UseAlefHundred)
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
            Str = If(UseClassic, If(NextStr = String.Empty, String.Empty, NextStr + If(Str = String.Empty, String.Empty, " " + ChData.ArabicCombiners(0))) + Str, If(Str = String.Empty, String.Empty, Str + If(NextStr = String.Empty, String.Empty, " " + ChData.ArabicCombiners(0))) + NextStr)
            NextStr = String.Empty
        Loop While CurBase >= 0
        If Number <> 0 Or Str = String.Empty Then NextStr = ArabicWordForLessThanThousand(CInt(Number), UseClassic, UseAlefHundred)
        Return If(UseClassic, If(NextStr = String.Empty, String.Empty, NextStr + If(Str = String.Empty, String.Empty, " " + ChData.ArabicCombiners(0))) + Str, If(Str = String.Empty, String.Empty, Str + If(NextStr = String.Empty, String.Empty, " " + ChData.ArabicCombiners(0))) + NextStr)
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
                        Replacements.Add(New RuleMetadata(Matches(MatchCount).Index + DupCount, Matches(MatchCount).Length - DupCount, New IslamData.RuleMetaSet.RuleMetadataTranslation.RuleWithArgs() {New IslamData.RuleMetaSet.RuleMetadataTranslation.RuleWithArgs With {.RuleName = Matches(MatchCount).Result(Rules(Count).Evaluator).Substring(DupCount)}}, Count))
                    End If
                Next
            End If
        Next
        Replacements.Sort(New RuleMetadataComparer)
        For Count = 0 To Replacements.Count - 1
            If Count <> 0 AndAlso (Replacements(Count).Index + Replacements(Count).Length > Replacements(Count - 1).Index) Then
                'Debug.Print(Rules(Replacements(Count - 1).OrigOrder).Name + ":" + Arabic.TransliterateToScheme(Replacements(Count - 1).Type, ArabicData.TranslitScheme.Literal, String.Empty, Nothing) + "-" + Rules(Replacements(Count).OrigOrder).Name + ":" + Arabic.TransliterateToScheme(Replacements(Count).Type, ArabicData.TranslitScheme.Literal, String.Empty, Nothing) + "-" + Arabic.TransliterateToScheme(ArabicString.Substring(Math.Max(Replacements(Count).Index - 15, 0), Math.Min(Replacements(Count - 1).Index + Replacements(Count - 1).Length + 15, ArabicString.Length) - Math.Max(Replacements(Count).Index - 15, 0)), ArabicData.TranslitScheme.Literal, String.Empty, Nothing))
            End If
            ArabicString = ArabicString.Substring(0, Replacements(Count).Index) + Replacements(Count).Type(0).RuleName + ArabicString.Substring(Replacements(Count).Index + Replacements(Count).Length)
        Next
        Return ArabicString
    End Function
    Public Function ChangeBaseScript(ArabicString As String, BaseText As TanzilReader.QuranTexts, ByVal PreString As String, ByVal PostString As String) As String
        If BaseText = TanzilReader.QuranTexts.Warsh Then
            ArabicString = UnjoinContig(ProcessTransform(JoinContig(ArabicString, PreString, PostString, False, False), ChData.RuleTranslations("WarshScript"), True), PreString, PostString)
        End If
        Return ArabicString
    End Function
    Public Function ChangeScript(ArabicString As String, SrcScriptType As TanzilReader.QuranScripts, ScriptType As TanzilReader.QuranScripts, ByVal PreString As String, ByVal PostString As String) As String
        If SrcScriptType = TanzilReader.QuranScripts.Uthmani Then
            If ScriptType = TanzilReader.QuranScripts.UthmaniMin Then
                ArabicString = UnjoinContig(ProcessTransform(JoinContig(ArabicString, PreString, PostString, False, False), ChData.RuleTranslations("UthmaniMinimalScript"), False), PreString, PostString)
            ElseIf ScriptType = TanzilReader.QuranScripts.SimpleEnhanced Then
                Dim ScriptCombine As New List(Of IslamMetadata.IslamData.RuleTranslationCategory.RuleTranslation)
                ScriptCombine.AddRange(ChData.RuleTranslations("SimpleScriptBase"))
                ScriptCombine.AddRange(ChData.RuleTranslations("SimpleEnhancedScript"))
                ArabicString = UnjoinContig(ProcessTransform(JoinContig(ArabicString, PreString, PostString, False, False), ScriptCombine.ToArray(), False), PreString, PostString)
            ElseIf ScriptType = TanzilReader.QuranScripts.Simple Then
                Dim ScriptCombine As New List(Of IslamMetadata.IslamData.RuleTranslationCategory.RuleTranslation)
                ScriptCombine.AddRange(ChData.RuleTranslations("SimpleScriptBase"))
                ScriptCombine.AddRange(ChData.RuleTranslations("SimpleScript"))
                ArabicString = UnjoinContig(ProcessTransform(JoinContig(ArabicString, PreString, PostString, False, False), ScriptCombine.ToArray(), False), PreString, PostString)
            ElseIf ScriptType = TanzilReader.QuranScripts.SimpleClean Then
                Dim ScriptCombine As New List(Of IslamMetadata.IslamData.RuleTranslationCategory.RuleTranslation)
                ScriptCombine.AddRange(ChData.RuleTranslations("SimpleScriptBase"))
                ScriptCombine.AddRange(ChData.RuleTranslations("SimpleCleanScript"))
                ArabicString = UnjoinContig(ProcessTransform(JoinContig(ArabicString, PreString, PostString, False, False), ScriptCombine.ToArray(), False), PreString, PostString)
            ElseIf ScriptType = TanzilReader.QuranScripts.SimpleMin Then
                Dim ScriptCombine As New List(Of IslamMetadata.IslamData.RuleTranslationCategory.RuleTranslation)
                ScriptCombine.AddRange(ChData.RuleTranslations("SimpleScriptBase"))
                ScriptCombine.AddRange(ChData.RuleTranslations("SimpleMinimalScript"))
                ArabicString = UnjoinContig(ProcessTransform(JoinContig(ArabicString, PreString, PostString, False, False), ScriptCombine.ToArray(), False), PreString, PostString)
            End If
        ElseIf SrcScriptType = TanzilReader.QuranScripts.UthmaniMin Then
            If ScriptType = TanzilReader.QuranScripts.Uthmani Then
                ArabicString = UnjoinContig(ProcessTransform(JoinContig(ArabicString, PreString, PostString, False, False), ChData.RuleTranslations("ReverseUthmaniMinimalScript"), False), PreString, PostString)
            End If
        ElseIf SrcScriptType = TanzilReader.QuranScripts.SimpleEnhanced Then
            If ScriptType = TanzilReader.QuranScripts.Uthmani Then
                Dim ScriptCombine As New List(Of IslamMetadata.IslamData.RuleTranslationCategory.RuleTranslation)
                ScriptCombine.AddRange(ChData.RuleTranslations("ReverseSimpleScriptBase"))
                ScriptCombine.AddRange(ChData.RuleTranslations("ReverseSimpleEnhancedScript"))
                ArabicString = UnjoinContig(ProcessTransform(JoinContig(ArabicString, PreString, PostString, False, False), ScriptCombine.ToArray(), False), PreString, PostString)
            End If
        ElseIf SrcScriptType = TanzilReader.QuranScripts.Simple Then
        ElseIf SrcScriptType = TanzilReader.QuranScripts.SimpleClean Then
        ElseIf SrcScriptType = TanzilReader.QuranScripts.SimpleMin Then
        End If
        Return ArabicString
    End Function
    Public Function ReplaceMetadata(ArabicString As String, MetadataRule As RuleMetadata, LearningMode As Boolean) As String
        If MetadataRule.Type.Length = 0 OrElse MetadataRule.Type.Length = 1 And MetadataRule.Type(0).RuleName = String.Empty Then Return ArabicString
        'use a dictionary/map here...
        For Count As Integer = 0 To ChData.IslamData.ColorRuleSets(0).ColorRules.Length - 1
            Dim TypeIndex As Integer = 0
            Dim Match As String = Linq.Enumerable.FirstOrDefault(ChData.IslamData.ColorRuleSets(0).ColorRules(Count).Match, Function(Str As String)
                                                                                                                                TypeIndex = New List(Of IslamData.RuleMetaSet.RuleMetadataTranslation.RuleWithArgs)(Linq.Enumerable.TakeWhile(MetadataRule.Type, Function(RuleWithArg As IslamData.RuleMetaSet.RuleMetadataTranslation.RuleWithArgs) RuleWithArg.RuleName <> Str)).Count
                                                                                                                                Return TypeIndex <> MetadataRule.Type.Length
                                                                                                                            End Function)
            If Match <> Nothing Then
                Dim Str As New System.Text.StringBuilder
                Str.Append(String.Format(ChData.IslamData.ColorRuleSets(0).ColorRules(Count).Evaluator, ArabicString.Substring(MetadataRule.Index, MetadataRule.Length)))
                If ChData.IslamData.ColorRuleSets(0).ColorRules(Count).MetaRuleFunc <> MetaRuleFuncs.eNone Then
                    Dim Args As String() = MetaRuleFunctions(ChData.IslamData.ColorRuleSets(0).ColorRules(Count).MetaRuleFunc - 1)(Str.ToString(), LearningMode)
                    If Args.Length = 1 Then
                        Str.Clear()
                        Str.Append(Args(0))
                    Else
                        Str.Clear()
                        For Index As Integer = 0 To Args.Length - 1
                            If Not Args(Index) Is Nothing And (LearningMode Or ChData.IslamData.ColorRuleSets(0).ColorRules(Count).MetaRuleFunc <> MetaRuleFuncs.eLearningMode Or Index <> 0) Then
                                Str.Append(ReplaceMetadata(Args(Index), New RuleMetadata(0, Args(Index).Length, New List(Of IslamData.RuleMetaSet.RuleMetadataTranslation.RuleWithArgs)(Linq.Enumerable.Select(MetadataRule.Type(TypeIndex).Args(Index), Function(S) New IslamData.RuleMetaSet.RuleMetadataTranslation.RuleWithArgs With {.RuleName = S})).ToArray(), Index), LearningMode))
                            End If
                        Next
                    End If
                End If
                ArabicString = ArabicString.Insert(MetadataRule.Index + MetadataRule.Length, Str.ToString()).Remove(MetadataRule.Index, MetadataRule.Length)
            End If
        Next
        Return ArabicString
    End Function
    Private Shared ErrorRegExDict As New Dictionary(Of String, System.Text.RegularExpressions.Regex)
    Public Sub DoErrorCheck(ByVal ArabicString As String)
        'need to check for decomposed first
        Dim Count As Integer
        For Count = 0 To ChData.RuleTranslations("ErrorCheck").Length - 1
            Dim Matches As System.Text.RegularExpressions.MatchCollection = ErrorRegExDict(ChData.RuleTranslations("ErrorCheck")(Count).Name + CStr(Count)).Matches(ArabicString)
            For MatchIndex As Integer = 0 To Matches.Count - 1
                If ChData.RuleTranslations("ErrorCheck")(Count).NegativeMatch Is Nothing OrElse Matches(MatchIndex).Result(ChData.RuleTranslations("ErrorCheck")(Count).NegativeMatch) = String.Empty Then
                    'Debug.Print(CachedData.RuleTranslations("ErrorCheck")(Count).Rule + ": " + TransliterateToScheme(ArabicString, ArabicData.TranslitScheme.Literal, String.Empty).Insert(Matches(MatchIndex).Index, "<!-- -->"))
                End If
            Next
        Next
    End Sub
    Public Shared Function JoinContig(ByVal ArabicString As String, ByVal PreString As String, ByVal PostString As String, PreStop As Boolean, PostStop As Boolean) As String
        Dim Index As Integer = PreString.LastIndexOf(" "c)
        'take last word of pre string and first word of post string or another if it is a pause marker
        'end of ayah sign without number is used as a proper place holder
        If Index <> -1 And PreString.Length - 2 = Index Then Index = PreString.LastIndexOf(" "c, Index - 1)
        If Index <> -1 Then PreString = PreString.Substring(Index + 1)
        If PreString <> String.Empty Then PreString += String.Join(Of Char)(String.Empty, New Char() {" "c, ArabicData.ArabicEndOfAyah, " "c})
        Index = PostString.IndexOf(" "c)
        If Index = 1 Then Index = PostString.IndexOf(" "c, Index + 1)
        If Index <> -1 Then PostString = PostString.Substring(0, Index)
        If PostString <> String.Empty Then PostString = String.Join(String.Empty, {" "c, ArabicData.ArabicEndOfAyah, " "c, PostString})
        'If Not OptionalStops Is Nothing Then
        '    Dim Stops As New List(Of Integer)
        '    If PreString <> String.Empty Then Stops.Add(PreString.Length - 2)
        '    For Count As Integer = 0 To OptionalStops.Length - 1
        '        Stops.Add(OptionalStops(Count) + PreString.Length)
        '    Next
        '    If PostString <> String.Empty Then Stops.Add(PreString.Length + ArabicString.Length + 1)
        'End If
        Return String.Join(String.Empty, {PreString, ArabicString, PostString})
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
    Private Shared RegExDict As New Dictionary(Of String, System.Text.RegularExpressions.Regex)
    Public Shared Function FilterMetadataStopsContig(ByVal ArabicStringLength As Integer, MetadataList As Generic.List(Of RuleMetadata), OptionalStops() As Integer, PreStringLength As Integer, Optional PreStop As Boolean = True, Optional PostStop As Boolean = True) As Generic.List(Of RuleMetadata)
        MetadataList = New List(Of RuleMetadata)(Linq.Enumerable.Select(Linq.Enumerable.Where(MetadataList, Function(Item) Item.Index >= PreStringLength And Item.Index + Item.Length <= ArabicStringLength), Function(Item) New RuleMetadata With {.Index = Item.Index - PreStringLength, .Length = Item.Length, .Type = Item.Type, .OrigOrder = Item.OrigOrder, .Dependencies = If(Item.Dependencies Is Nothing, Nothing, New List(Of RuleMetadata)(Linq.Enumerable.Select(Item.Dependencies, Function(It) New RuleMetadata With {.Index = It.Index - PreStringLength, .Length = It.Length, .Type = It.Type, .OrigOrder = It.OrigOrder, .Children = It.Children})).ToArray()), .Children = Item.Children}))
        'If PreStop Then OptionalStops.Add(0)
        'If PostStop Then OptionalStops.Add(ArabicString.Length - 1)
        Return MetadataList 'FilterMetadataStops(ArabicString, MetadataList, OptionalStops)
    End Function
    Public Shared Function FilterMetadataStops(ByVal ArabicString As String, MetadataList As Generic.List(Of RuleMetadata), OptionalStops() As Integer) As Generic.List(Of RuleMetadata)
        MetadataList = New List(Of RuleMetadata)(Linq.Enumerable.Where(MetadataList, Function(Item)
                                                                                         Return Item.Dependencies Is Nothing OrElse
                (Not (Item.Dependencies(0).Type(0).RuleName = "optionalstop" AndAlso (OptionalStops Is Nothing AndAlso (Item.Dependencies(0).Index <> -1 AndAlso Item.Dependencies(0).Index <> ArabicString.Length AndAlso ArabicString.Substring(Item.Dependencies(0).Index, Item.Dependencies(0).Length) = ArabicData.ArabicSmallHighLigatureSadWithLamWithAlefMaksura OrElse (Item.Dependencies(0).Index <> 0 And Item.Dependencies(0).Index <> ArabicString.Length)) OrElse (Not OptionalStops Is Nothing AndAlso Item.Dependencies(0).Length <> 0 AndAlso Linq.Enumerable.All(OptionalStops, Function(Idx) Not (Idx >= Item.Dependencies(0).Index And Idx < Item.Dependencies(0).Index + Item.Dependencies(0).Length))))) AndAlso
                Not (Item.Dependencies(0).Type(0).RuleName = "optionalnotstop" AndAlso (OptionalStops Is Nothing AndAlso (Item.Dependencies(0).Index = -1 OrElse Item.Dependencies(0).Index = ArabicString.Length OrElse ArabicString.Substring(Item.Dependencies(0).Index, Item.Dependencies(0).Length) <> ArabicData.ArabicSmallHighLigatureSadWithLamWithAlefMaksura) AndAlso (Item.Dependencies(0).Length = 0 AndAlso (Item.Dependencies(0).Index = 0 Or Item.Dependencies(0).Index = ArabicString.Length)) OrElse (Not OptionalStops Is Nothing AndAlso Item.Dependencies(0).Length <> 0 AndAlso Linq.Enumerable.Any(OptionalStops, Function(Idx) Idx >= Item.Dependencies(0).Index And Idx < Item.Dependencies(0).Index + Item.Dependencies(0).Length)))))
                                                                                     End Function))
        Dim Index As Integer = 0
        Dim RemoveIndexes As List(Of Integer) = New List(Of Integer)(Linq.Enumerable.Range(0, MetadataList.Count))
        While Index <= MetadataList.Count - 1
            Do While Index <> MetadataList.Count - 1 AndAlso MetadataList(Index).Index = MetadataList(Index + 1).Index AndAlso MetadataList(Index).Length = MetadataList(Index + 1).Length AndAlso CompareRuleMetadata(MetadataList(Index).Children, MetadataList(Index + 1).Children)
                Dim FirstRule As IslamData.RuleMetaSet.RuleMetadataTranslation.RuleWithArgs() = MetadataList(Index).Type
                Dim SecondRule As New List(Of IslamData.RuleMetaSet.RuleMetadataTranslation.RuleWithArgs)
                SecondRule.AddRange(MetadataList(Index + 1).Type)
                Dim FirstUpdate As New List(Of IslamData.RuleMetaSet.RuleMetadataTranslation.RuleWithArgs)
                For FirstIndex As Integer = 0 To FirstRule.Length - 1
                    Dim SecondIndex As Integer
                    For SecondIndex = 0 To SecondRule.Count - 1
                        If FirstRule(FirstIndex).RuleName = SecondRule(SecondIndex).RuleName And Not FirstRule(FirstIndex).Args Is Nothing Then
                            Dim Matches As String()() = FirstRule(FirstIndex).Args
                            If Matches.Length <> 1 Then
                                Dim ArgList As New List(Of String())
                                For Count = 0 To Matches.Length - 1
                                    ArgList.Add(New List(Of String)(Linq.Enumerable.Concat(Matches(Count), SecondRule(SecondIndex).Args(Count))).ToArray())
                                Next
                                FirstUpdate.Add(New IslamData.RuleMetaSet.RuleMetadataTranslation.RuleWithArgs With {.RuleName = FirstRule(FirstIndex).RuleName, .Args = ArgList.ToArray()})
                            Else
                                FirstUpdate.Add(FirstRule(FirstIndex))
                            End If
                            SecondRule.RemoveAt(SecondIndex)
                            SecondIndex -= 1
                            Exit For
                        End If
                    Next
                    If SecondIndex = SecondRule.Count Then FirstUpdate.Add(FirstRule(FirstIndex))
                Next
                SecondRule.InsertRange(0, FirstUpdate)
                'Debug.WriteLine("First: " + MetadataList(Index).Type + " Second: " + MetadataList(Index + 1).Type + " After: " + String.Join("|"c, SecondRule.ToArray()))
                MetadataList(Index + 1) = New RuleMetadata(MetadataList(Index).Index, MetadataList(Index).Length, SecondRule.ToArray(), MetadataList(Index).OrigOrder) With {.Children = MetadataList(Index).Children, .Dependencies = MetadataList(Index).Dependencies}
                RemoveIndexes(Index) = -1
                Index += 1
            Loop
            Index += 1
        End While
        Return New List(Of RuleMetadata)(Linq.Enumerable.Where(MetadataList, Function(Item, Idx) RemoveIndexes(Idx) <> -1))
    End Function
    Public Function TransliterateWithRules(ByVal ArabicString As String, Scheme As String, LearningMode As Boolean, MetadataList As Generic.List(Of RuleMetadata)) As String
        If Scheme = String.Empty Then Scheme = ChData.IslamData.LanguageDefaultInfo.GetLanguageByID(String.Empty).TranslitScheme
        'DoErrorCheck(ArabicString)
        Dim Index As Integer = 0
        While Index <= MetadataList.Count - 1
            ArabicString = ReplaceMetadata(ArabicString, MetadataList(Index), LearningMode)
            If Not MetadataList(Index).Children Is Nothing Then
                For SubCount As Integer = 0 To MetadataList(Index).Children.Length - 1
                    ArabicString = ArabicString.Remove(MetadataList(Index).Index).Insert(MetadataList(Index).Index, ReplaceMetadata(ArabicString.Substring(MetadataList(Index).Index), MetadataList(Index).Children(SubCount), LearningMode))
                Next
            End If
            Index += 1
        End While
        ArabicString = ReplaceTranslitRule(ArabicString, Scheme, LearningMode, Nothing)
        'process wasl loanwords and names
        'process loanwords and names
        If System.Text.RegularExpressions.Regex.Match(ArabicString, "[\p{IsArabic}|\p{IsArabicPresentationForms-A}|\p{IsArabicPresentationForms-B}]").Success Then Debug.WriteLine(ArabicString.Substring(System.Text.RegularExpressions.Regex.Match(ArabicString, "[\p{IsArabic}|\p{IsArabicPresentationForms-A}|\p{IsArabicPresentationForms-B}]").Index) + " --- " + ArabicString)
        Return ArabicString
    End Function
    Public Function GetCacheMetarules(Lines As String(), IndexToVerse As Integer()()) As Generic.List(Of RuleMetadata)
        Dim CacheIn As String() = Lines
        Dim Rules As New List(Of RuleMetadata)
        For Count As Integer = 0 To CacheIn.Length - 1
            Dim KeyVal As String() = CacheIn(Count).Split("="c)
            Dim Key As IslamData.RuleMetaSet.RuleMetadataTranslation.RuleWithArgs() = New List(Of IslamData.RuleMetaSet.RuleMetadataTranslation.RuleWithArgs)(Linq.Enumerable.Select(KeyVal(0).Substring(0, Math.Min(If(KeyVal(0).IndexOf("["c) = -1, KeyVal(0).Length, KeyVal(0).IndexOf("["c)), If(KeyVal(0).IndexOf("{"c) = -1, KeyVal(0).Length, KeyVal(0).IndexOf("{"c)))).Split("|"c), Function(Item) New IslamData.RuleMetaSet.RuleMetadataTranslation.RuleWithArgs() With {.RuleName = Item})).ToArray()
            Dim Vals As Integer()() = New List(Of Integer())(Linq.Enumerable.Select(KeyVal(1).Split(","c), Function(Item) New List(Of Integer)(If(Item Is Nothing Or Item = String.Empty, {}, Linq.Enumerable.Select(Item.Split(":"c), Function(Frag) Integer.Parse(Frag)))).ToArray())).ToArray()
            Dim Length As Integer = 1
            Dim VerseCount As Integer
            For VerseCount = 0 To Vals.Length - 1
                Dim Ch As RuleMetadata() = If(KeyVal(0).IndexOf("["c) = -1, Nothing, New List(Of RuleMetadata)(Linq.Enumerable.Select(KeyVal(0).Substring(KeyVal(0).IndexOf("["c) + 1, KeyVal(0).IndexOf("]"c) - KeyVal(0).IndexOf("["c) - 1).Split(","c), Function(Item)
                                                                                                                                                                                                                                                               Dim Pieces As String() = Item.Split(";"c)
                                                                                                                                                                                                                                                               Return New RuleMetadata(CInt(Pieces(1)), CInt(Pieces(2)), New List(Of IslamData.RuleMetaSet.RuleMetadataTranslation.RuleWithArgs)(Linq.Enumerable.Select(Pieces(0).Split(" "c), Function(It) New IslamData.RuleMetaSet.RuleMetadataTranslation.RuleWithArgs() With {.RuleName = It})).ToArray(), 0)
                                                                                                                                                                                                                                                           End Function)).ToArray())
                Dim Idx As Integer = Array.BinarySearch(IndexToVerse, New Integer() {Vals(VerseCount)(0), Vals(VerseCount)(1), Vals(VerseCount)(2)}, New TanzilReader.QuranWordChapterVerseWordComparer(False))
                If Idx < 0 Then Idx = (Idx Xor -1) - 1 'gaps based off previous word
                Dim Dep As RuleMetadata() = If(KeyVal(0).IndexOf("{"c) = -1, Nothing, New List(Of RuleMetadata)(Linq.Enumerable.Select(KeyVal(0).Substring(KeyVal(0).IndexOf("{"c) + 1, KeyVal(0).IndexOf("}"c) - KeyVal(0).IndexOf("{"c) - 1).Split(","c), Function(Item)
                                                                                                                                                                                                                                                                Dim Pieces As String() = Item.Split(";"c)
                                                                                                                                                                                                                                                                Return New RuleMetadata(CInt(Pieces(1)) + IndexToVerse(Idx)(3) + Vals(VerseCount)(3), CInt(Pieces(2)), New List(Of IslamData.RuleMetaSet.RuleMetadataTranslation.RuleWithArgs)(Linq.Enumerable.Select(Pieces(0).Split(" "c), Function(It) New IslamData.RuleMetaSet.RuleMetadataTranslation.RuleWithArgs() With {.RuleName = It})).ToArray(), 0)
                                                                                                                                                                                                                                                            End Function)).ToArray())
                Rules.Add(New RuleMetadata(IndexToVerse(Idx)(3) + Vals(VerseCount)(3), Vals(VerseCount)(4), Key, -1) With {.Children = Ch, .Dependencies = Dep})
                Length = 1
            Next
        Next
        Return Rules
    End Function
    Public Shared Function MakeCacheMetarules(RuleMetadata As Generic.List(Of RuleMetadata), IndexToVerse As Integer()()) As String()
        Dim RuleDictionary As New Dictionary(Of String, List(Of RuleMetadata))
        For Count As Integer = 0 To RuleMetadata.Count - 1
            Dim Idx As Integer = Count
            Dim Key As String = String.Join("|"c, Linq.Enumerable.Select(RuleMetadata(Count).Type, Function(It) It.RuleName)) + If(Not RuleMetadata(Count).Children Is Nothing AndAlso RuleMetadata(Count).Children.Length <> 0, "[" + String.Join(","c, Linq.Enumerable.Select(RuleMetadata(Count).Children, Function(Item) String.Join(" "c, Linq.Enumerable.Select(Item.Type, Function(It) It.RuleName)) + ";" + CStr(Item.Index) + ";" + CStr(Item.Length))) + "]", String.Empty) + If(Not RuleMetadata(Count).Dependencies Is Nothing AndAlso RuleMetadata(Count).Dependencies.Length <> 0, "{" + String.Join(","c, Linq.Enumerable.Select(RuleMetadata(Count).Dependencies, Function(Item) String.Join(" "c, Linq.Enumerable.Select(Item.Type, Function(It) It.RuleName)) + ";" + CStr(Item.Index - RuleMetadata(Idx).Index) + ";" + CStr(Item.Length))) + "}", String.Empty)
            If Not RuleDictionary.ContainsKey(Key) Then RuleDictionary.Add(Key, New List(Of RuleMetadata))
            RuleDictionary(Key).Add(RuleMetadata(Count))
        Next
        Dim CacheOut As New List(Of String)
        For Each KeyVal In RuleDictionary
            Dim Str As New Text.StringBuilder
            Str.Append(KeyVal.Key)
            Str.Append("="c)
            For DictCount As Integer = 0 To KeyVal.Value.Count - 1
                Dim Index As Integer = Array.BinarySearch(IndexToVerse, KeyVal.Value(DictCount).Index, New TanzilReader.QuranWordIndexComparer)
                If Index < 0 Then Index = (Index Xor -1) - 1 'gaps based off previous word
                If DictCount <> 0 Then Str.Append(","c)
                Str.Append(CStr(IndexToVerse(Index)(0)))
                Str.Append(":"c)
                Str.Append(CStr(IndexToVerse(Index)(1)))
                Str.Append(":"c)
                Str.Append(CStr(IndexToVerse(Index)(2)))
                Str.Append(":"c)
                Str.Append(CStr(KeyVal.Value(DictCount).Index - IndexToVerse(Index)(3)))
                Str.Append(":"c)
                Str.Append(CStr(KeyVal.Value(DictCount).Length))
            Next
            CacheOut.Add(Str.ToString())
        Next
        Return CacheOut.ToArray()
    End Function
    Public Class MetaruleComparer
        Implements IComparer(Of Object)
        Public Function Compare(x As Object, y As Object) As Integer Implements IComparer(Of Object).Compare
            Return If(CInt(x) >= CType(y, RuleMetadata).Index, 1, -1)
        End Function
    End Class
    Public Shared Function GetMetarulesFromCache(ArabicString As String, BasePosition As Integer, Rules As Generic.List(Of RuleMetadata)) As Generic.List(Of RuleMetadata)
        Dim Index As Integer = Array.BinarySearch(Of Object)(CType(CType(Rules.ToArray(), Array), Object()), BasePosition, New MetaruleComparer) Xor -1
        Return New List(Of RuleMetadata)(Linq.Enumerable.TakeWhile(Rules, Function(Item) Item.Index < BasePosition + ArabicString.Length))
        'FilterMetadataStopsContig(ArabicString.Length, , , BasePosition)
    End Function
    Public Shared Function CompareRuleMetadata(RuleMetadata1() As RuleMetadata, RuleMetadata2() As RuleMetadata) As Boolean
        If RuleMetadata1 Is RuleMetadata2 Then Return True
        If RuleMetadata1 Is Nothing Or RuleMetadata2 Is Nothing Then Return False
        If RuleMetadata1.Length <> RuleMetadata2.Length Then Return False
        For Count As Integer = 0 To RuleMetadata1.Length - 1
            If RuleMetadata1(Count).Index <> RuleMetadata2(Count).Index Or RuleMetadata1(Count).Length <> RuleMetadata2(Count).Length Or RuleMetadata1(Count).Type(0).RuleName <> RuleMetadata2(Count).Type(0).RuleName Then Return False
        Next
        Return True
    End Function
    Public Function GetMetarules(ArabicString As String, RuleSet As IslamMetadata.IslamData.RuleMetaSet) As Generic.List(Of RuleMetadata)
        Dim Count As Integer
        Dim MetadataList As New Generic.List(Of RuleMetadata)
        For Count = 0 To RuleSet.Rules.Length - 1
            If Not RuleSet.Rules(Count).Evaluator Is Nothing Then
                Dim Matches As System.Text.RegularExpressions.MatchCollection = RegExDict(String.Join(String.Empty, RuleSet.Name, "_", RuleSet.Rules(Count).Name, CStr(Count))).Matches(ArabicString)
                Dim MatchIndex As Integer
                For MatchIndex = 0 To Matches.Count - 1
                    Dim SubCount As Integer
                    Dim ParentDependency As New List(Of RuleMetadata)
                    Dim StopCount As Integer
                    For StopCount = 0 To RuleSet.Rules(Count).OptionalStopIndexes.Length - 1
                        SubCount = RuleSet.Rules(Count).OptionalStopIndexes(StopCount)
                        If (Matches(MatchIndex).Groups(SubCount + 1).Success AndAlso Array.IndexOf(RuleSet.Rules(Count).OptionalNotStopIndexes, SubCount) <> -1) Then Exit For
                    Next
                    If StopCount <> RuleSet.Rules(Count).OptionalStopIndexes.Length Then Continue For
                    For StopCount = 0 To RuleSet.Rules(Count).OptionalStopIndexes.Length - 1
                        SubCount = RuleSet.Rules(Count).OptionalStopIndexes(StopCount)
                        If Matches(MatchIndex).Groups(SubCount + 1).Success Then ParentDependency.Add(New RuleMetadata(Matches(MatchIndex).Groups(SubCount + 1).Index, Matches(MatchIndex).Groups(SubCount + 1).Length, RuleSet.Rules(Count).Evaluator(SubCount), SubCount) With {.Children = If(Linq.Enumerable.Any(RecursiveMetadata, Function(It) Linq.Enumerable.Any(RuleSet.Rules(Count).Evaluator(SubCount), Function(RuleWithArgs) RuleWithArgs.RuleName = It)), FilterMetadataStops(MetaRuleFunctions((Linq.Enumerable.First(ChData.IslamData.ColorRuleSets(0).ColorRules, Function(Item) Linq.Enumerable.Any(Item.Match, Function(Mat) Linq.Enumerable.Any(RuleSet.Rules(Count).Evaluator(SubCount), Function(RuleWithArgs) RuleWithArgs.RuleName = Mat)))).MetaRuleFunc - 1)(Matches(MatchIndex).Groups(SubCount + 1).Value, False)(0), GetMetarules(MetaRuleFunctions((Linq.Enumerable.First(ChData.IslamData.ColorRuleSets(0).ColorRules, Function(Item) Linq.Enumerable.Any(Item.Match, Function(Mat) Linq.Enumerable.Any(RuleSet.Rules(Count).Evaluator(SubCount), Function(RuleWithArgs) RuleWithArgs.RuleName = Mat)))).MetaRuleFunc - 1)(Matches(MatchIndex).Groups(SubCount + 1).Value, False)(0), If(Linq.Enumerable.Any(RuleSet.Rules(Count).Evaluator(SubCount), Function(RuleWithArgs) RuleWithArgs.RuleName = "spellnumber"), ChData.RuleMetas("Normal"), RuleSet)), Nothing).ToArray(), Nothing)})
                    Next
                    For StopCount = 0 To RuleSet.Rules(Count).OptionalNotStopIndexes.Length - 1
                        SubCount = RuleSet.Rules(Count).OptionalNotStopIndexes(StopCount)
                        If Matches(MatchIndex).Groups(SubCount + 1).Success Then ParentDependency.Add(New RuleMetadata(Matches(MatchIndex).Groups(SubCount + 1).Index, Matches(MatchIndex).Groups(SubCount + 1).Length, RuleSet.Rules(Count).Evaluator(SubCount), SubCount) With {.Children = If(Linq.Enumerable.Any(RecursiveMetadata, Function(It) Linq.Enumerable.Any(RuleSet.Rules(Count).Evaluator(SubCount), Function(RuleWithArgs) RuleWithArgs.RuleName = It)), FilterMetadataStops(MetaRuleFunctions((Linq.Enumerable.First(ChData.IslamData.ColorRuleSets(0).ColorRules, Function(Item) Linq.Enumerable.Any(Item.Match, Function(Mat) Linq.Enumerable.Any(RuleSet.Rules(Count).Evaluator(SubCount), Function(RuleWithArgs) RuleWithArgs.RuleName = Mat)))).MetaRuleFunc - 1)(Matches(MatchIndex).Groups(SubCount + 1).Value, False)(0), GetMetarules(MetaRuleFunctions((Linq.Enumerable.First(ChData.IslamData.ColorRuleSets(0).ColorRules, Function(Item) Linq.Enumerable.Any(Item.Match, Function(Mat) Linq.Enumerable.Any(RuleSet.Rules(Count).Evaluator(SubCount), Function(RuleWithArgs) RuleWithArgs.RuleName = Mat)))).MetaRuleFunc - 1)(Matches(MatchIndex).Groups(SubCount + 1).Value, False)(0), If(Linq.Enumerable.Any(RuleSet.Rules(Count).Evaluator(SubCount), Function(RuleWithArgs) RuleWithArgs.RuleName = "spellnumber"), ChData.RuleMetas("Normal"), RuleSet)), Nothing).ToArray(), Nothing)})
                    Next
                    Dim ParentDeps As RuleMetadata() = If(ParentDependency.Count = 0, Nothing, ParentDependency.ToArray())
                    For SubCount = 0 To RuleSet.Rules(Count).Evaluator.Length - 1
                        If Not RuleSet.Rules(Count).Evaluator(SubCount) Is Nothing AndAlso RuleSet.Rules(Count).Evaluator(SubCount).Length <> 0 AndAlso (RuleSet.Rules(Count).Evaluator(SubCount).Length <> 1 Or (RuleSet.Rules(Count).Evaluator(SubCount)(0).RuleName <> String.Empty And RuleSet.Rules(Count).Evaluator(SubCount)(0).RuleName <> "optionalstop" And RuleSet.Rules(Count).Evaluator(SubCount)(0).RuleName <> "optionalnotstop")) And (Matches(MatchIndex).Groups(SubCount + 1).Length <> 0 Or Array.IndexOf(AllowZeroLength, RuleSet.Rules(Count).Evaluator(SubCount)(0).RuleName) <> -1) Then
                            MetadataList.Add(New RuleMetadata(Matches(MatchIndex).Groups(SubCount + 1).Index, Matches(MatchIndex).Groups(SubCount + 1).Length, RuleSet.Rules(Count).Evaluator(SubCount), SubCount) With {.Children = If(Linq.Enumerable.Any(RecursiveMetadata, Function(It) Linq.Enumerable.Any(RuleSet.Rules(Count).Evaluator(SubCount), Function(RuleWithArgs) RuleWithArgs.RuleName = It)), FilterMetadataStops(MetaRuleFunctions((Linq.Enumerable.First(ChData.IslamData.ColorRuleSets(0).ColorRules, Function(Item) Linq.Enumerable.Any(Item.Match, Function(Mat) Linq.Enumerable.Any(RuleSet.Rules(Count).Evaluator(SubCount), Function(RuleWithArgs) RuleWithArgs.RuleName = Mat)))).MetaRuleFunc - 1)(Matches(MatchIndex).Groups(SubCount + 1).Value, False)(0), GetMetarules(MetaRuleFunctions((Linq.Enumerable.First(ChData.IslamData.ColorRuleSets(0).ColorRules, Function(Item) Linq.Enumerable.Any(Item.Match, Function(Mat) Linq.Enumerable.Any(RuleSet.Rules(Count).Evaluator(SubCount), Function(RuleWithArgs) RuleWithArgs.RuleName = Mat)))).MetaRuleFunc - 1)(Matches(MatchIndex).Groups(SubCount + 1).Value, False)(0), If(Linq.Enumerable.Any(RuleSet.Rules(Count).Evaluator(SubCount), Function(RuleWithArgs) RuleWithArgs.RuleName = "spellnumber"), ChData.RuleMetas("Normal"), RuleSet)), Nothing).ToArray(), Nothing), .Dependencies = ParentDeps})
                            If (Not MetadataList(MetadataList.Count - 1).Children Is Nothing) Then Array.Sort(MetadataList(MetadataList.Count - 1).Children, New RuleMetadataComparer)
                            'Debug.WriteLine(RuleSet(Count).Name + " Index: " + CStr(Matches(MatchIndex).Groups(SubCount + 1).Index) + " Length: " + CStr(Matches(MatchIndex).Groups(SubCount + 1).Length) + " Ruling: " + RuleSet(Count).Evaluator(SubCount))
                        End If
                    Next
                Next
            End If
        Next
        MetadataList.Sort(New RuleMetadataComparer)
        Dim Index As Integer = 0
        Dim RemoveIndexes As List(Of Integer) = New List(Of Integer)(Linq.Enumerable.Range(0, MetadataList.Count))
        While Index <= MetadataList.Count - 1
            Do While Index <> MetadataList.Count - 1 AndAlso MetadataList(Index).Index = MetadataList(Index + 1).Index AndAlso MetadataList(Index).Length = MetadataList(Index + 1).Length AndAlso CompareRuleMetadata(MetadataList(Index).Dependencies, MetadataList(Index + 1).Dependencies) AndAlso CompareRuleMetadata(MetadataList(Index).Children, MetadataList(Index + 1).Children)
                Dim FirstRule As IslamData.RuleMetaSet.RuleMetadataTranslation.RuleWithArgs() = MetadataList(Index).Type
                Dim SecondRule As New List(Of IslamData.RuleMetaSet.RuleMetadataTranslation.RuleWithArgs)
                SecondRule.AddRange(MetadataList(Index + 1).Type)
                Dim FirstUpdate As New List(Of IslamData.RuleMetaSet.RuleMetadataTranslation.RuleWithArgs)
                For FirstIndex As Integer = 0 To FirstRule.Length - 1
                    Dim SecondIndex As Integer
                    For SecondIndex = 0 To SecondRule.Count - 1
                        If FirstRule(FirstIndex).RuleName = SecondRule(SecondIndex).RuleName And Not FirstRule(FirstIndex).Args Is Nothing Then
                            Dim Matches As String()() = FirstRule(FirstIndex).Args
                            If Matches.Length <> 1 Then
                                Dim ArgList As New List(Of String())
                                For Count = 0 To Matches.Length - 1
                                    ArgList.Add(New List(Of String)(Linq.Enumerable.Concat(Matches(Count), SecondRule(SecondIndex).Args(Count))).ToArray())
                                Next
                                FirstUpdate.Add(New IslamData.RuleMetaSet.RuleMetadataTranslation.RuleWithArgs With {.RuleName = FirstRule(FirstIndex).RuleName, .Args = ArgList.ToArray()})
                            Else
                                FirstUpdate.Add(FirstRule(FirstIndex))
                            End If
                            SecondRule.RemoveAt(SecondIndex)
                            SecondIndex -= 1
                            Exit For
                        End If
                    Next
                    If SecondIndex = SecondRule.Count Then FirstUpdate.Add(FirstRule(FirstIndex))
                Next
                SecondRule.InsertRange(0, FirstUpdate)
                'Debug.WriteLine("First: " + MetadataList(Index).Type + " Second: " + MetadataList(Index + 1).Type + " After: " + String.Join("|"c, SecondRule.ToArray()))
                MetadataList(Index + 1) = New RuleMetadata(MetadataList(Index).Index, MetadataList(Index).Length, SecondRule.ToArray(), MetadataList(Index).OrigOrder) With {.Children = MetadataList(Index).Children, .Dependencies = MetadataList(Index).Dependencies}
                RemoveIndexes(Index) = -1
                Index += 1
            Loop
            Index += 1
        End While
        Return New List(Of RuleMetadata)(Linq.Enumerable.Where(MetadataList, Function(Item, Idx) RemoveIndexes(Idx) <> -1))
    End Function
    Public Function ReplaceTranslitRule(ArabicString As String, Scheme As String, LearningMode As Boolean, ByRef DiffMap As List(Of Integer)) As String
        'redundant romanization rules should have -'s such as seen/teh/kaf-heh
        Dim Count As Integer
        Dim _DiffMap As List(Of Integer) = If(DiffMap Is Nothing, Nothing, DiffMap)
        For Count = 0 To ChData.RuleTranslations("RomanizationRules").Length - 1
            Dim AdjRepCount As Integer = 0
            If ChData.RuleTranslations("RomanizationRules")(Count).RuleFunc = RuleFuncs.eNone Then
                ArabicString = System.Text.RegularExpressions.Regex.Replace(ArabicString, ChData.RuleTranslations("RomanizationRules")(Count).Match, Function(Match As System.Text.RegularExpressions.Match)
                                                                                                                                                         Dim Str As String = Match.Result(ChData.RuleTranslations("RomanizationRules")(Count).Evaluator)
                                                                                                                                                         If Not _DiffMap Is Nothing And Str.Length > Match.Length Then _DiffMap.InsertRange(Match.Index + AdjRepCount + Match.Length, Linq.Enumerable.Select(Str.Substring(0, Str.Length - Match.Length).ToCharArray(), Function(It) _DiffMap(Match.Index + AdjRepCount + Match.Length - 1)))
                                                                                                                                                         If Not _DiffMap Is Nothing And Str.Length < Match.Length Then _DiffMap.RemoveRange(Match.Index + AdjRepCount + Str.Length, Match.Length - Str.Length)
                                                                                                                                                         AdjRepCount += Str.Length - Match.Length
                                                                                                                                                         Return Str
                                                                                                                                                     End Function)
            Else
                ArabicString = System.Text.RegularExpressions.Regex.Replace(ArabicString, ChData.RuleTranslations("RomanizationRules")(Count).Match, Function(Match As System.Text.RegularExpressions.Match)
                                                                                                                                                         Dim Str As String = RuleFunctions(ChData.RuleTranslations("RomanizationRules")(Count).RuleFunc - 1)(Match.Result(ChData.RuleTranslations("RomanizationRules")(Count).Evaluator), Scheme)(0)
                                                                                                                                                         If Not _DiffMap Is Nothing And Str.Length > Match.Length Then _DiffMap.InsertRange(Match.Index + AdjRepCount + Match.Length, Linq.Enumerable.Select(Str.Substring(0, Str.Length - Match.Length).ToCharArray(), Function(It) _DiffMap(Match.Index + AdjRepCount + Match.Length - 1)))
                                                                                                                                                         If Not _DiffMap Is Nothing And Str.Length < Match.Length Then _DiffMap.RemoveRange(Match.Index + AdjRepCount + Str.Length, Match.Length - Str.Length)
                                                                                                                                                         AdjRepCount += Str.Length - Match.Length
                                                                                                                                                         Return Str
                                                                                                                                                     End Function)
            End If
        Next
        If Not DiffMap Is Nothing Then DiffMap = _DiffMap
        Return ArabicString
    End Function
    Public Function TransliterateWithRulesColor(ByVal ArabicString As String, Scheme As String, BreakWords As Boolean, LearningMode As Boolean, MetadataList As Generic.List(Of RuleMetadata)) As RenderArray.RenderText()()
        Dim Count As Integer
        If Scheme = String.Empty Then Scheme = ChData.IslamData.LanguageDefaultInfo.GetLanguageByID(String.Empty).TranslitScheme
        DoErrorCheck(ArabicString)
        Dim Index As Integer = 0
        Dim OffsetList As New List(Of Integer)
        Dim RecOffsetList As New List(Of List(Of Integer))
        While Index <= MetadataList.Count - 1
            Dim OldLength As Integer = ArabicString.Length
            OffsetList.Add(ArabicString.Length)
            ArabicString = ReplaceMetadata(ArabicString, MetadataList(Index), LearningMode)
            If Not MetadataList(Index).Children Is Nothing Then
                RecOffsetList.Add(New List(Of Integer))
                For SubCount As Integer = 0 To MetadataList(Index).Children.Length - 1
                    RecOffsetList(Index).Add(ArabicString.Length)
                    ArabicString = ArabicString.Remove(MetadataList(Index).Index).Insert(MetadataList(Index).Index, ReplaceMetadata(ArabicString.Substring(MetadataList(Index).Index), MetadataList(Index).Children(SubCount), LearningMode))
                    RecOffsetList(Index)(RecOffsetList(Index).Count - 1) = ArabicString.Length - RecOffsetList(Index)(RecOffsetList(Index).Count - 1)
                Next
            Else
                RecOffsetList.Add(Nothing)
            End If
            OffsetList(OffsetList.Count - 1) = ArabicString.Length - OffsetList(OffsetList.Count - 1)
            Index += 1
        End While
        Dim RuleIndexes As List(Of Integer) = New List(Of Integer)(Linq.Enumerable.Select(ArabicString.ToCharArray(), Function(It) 0))
        Dim CumOffset As Integer = 0
        For Index = MetadataList.Count - 1 To 0 Step -1
            For Count = 0 To ChData.IslamData.ColorRuleSets(1).ColorRules.Length - 1
                Dim Match As Integer = New List(Of String)(Linq.Enumerable.TakeWhile(ChData.IslamData.ColorRuleSets(1).ColorRules(Count).Match, Function(Str As String) Linq.Enumerable.Any(MetadataList(Index).Type, Function(RuleWithArg) RuleWithArg.RuleName <> Str))).Count
                If Match <> ChData.IslamData.ColorRuleSets(1).ColorRules(Count).Match.Length Then
                    For MetaCount As Integer = MetadataList(Index).Index + CumOffset To MetadataList(Index).Index + CumOffset + MetadataList(Index).Length + OffsetList(Index) - 1
                        If ChData.IslamData.ColorRuleSets(1).ColorRules(RuleIndexes(MetaCount)).Color = &HFF000000 Then RuleIndexes(MetaCount) = Count
                    Next
                    'ApplyColorRules(Strings(Strings.Count - 1).Text)
                End If
            Next
            If Not MetadataList(Index).Children Is Nothing Then
                Dim SubCount As Integer
                For SubCount = MetadataList(Index).Children.Length - 1 To 0 Step -1
                    Dim RecCumOffset As Integer = 0
                    For Count = 0 To ChData.IslamData.ColorRuleSets(1).ColorRules.Length - 1
                        Dim Match As Integer = New List(Of String)(Linq.Enumerable.TakeWhile(ChData.IslamData.ColorRuleSets(1).ColorRules(Count).Match, Function(Str As String) Linq.Enumerable.Any(MetadataList(Index).Children(SubCount).Type, Function(RuleWithArg) RuleWithArg.RuleName <> Str))).Count
                        If Match <> ChData.IslamData.ColorRuleSets(1).ColorRules(Count).Match.Length Then
                            For MetaCount As Integer = MetadataList(Index).Index + CumOffset + MetadataList(Index).Children(SubCount).Index + RecCumOffset To MetadataList(Index).Index + CumOffset + MetadataList(Index).Children(SubCount).Index + RecCumOffset + MetadataList(Index).Children(SubCount).Length + RecOffsetList(Index)(SubCount) - 1
                                If ChData.IslamData.ColorRuleSets(1).ColorRules(RuleIndexes(MetaCount)).Color = &HFF000000 Then RuleIndexes(MetaCount) = Count
                            Next
                            'ApplyColorRules(Strings(Strings.Count - 1).Text)
                        End If
                    Next
                    RecCumOffset += RecOffsetList(Index)(SubCount)
                Next
            End If
            For MetaCount As Integer = MetadataList(Index).Index + CumOffset To MetadataList(Index).Index + CumOffset + MetadataList(Index).Length + OffsetList(Index) - 1
                If ArabicString(MetaCount) = " "c And BreakWords And Linq.Enumerable.Any(RecursiveMetadata, Function(It) Linq.Enumerable.Any(MetadataList(Index).Type, Function(RuleWithArgs) RuleWithArgs.RuleName = It)) Then RuleIndexes(MetaCount) = -1
            Next
            CumOffset += OffsetList(Index)
        Next
        Dim Base As Integer = 0
        Dim WordRenderers As New List(Of RenderArray.RenderText())
        Dim Renderers As New List(Of RenderArray.RenderText)
        ArabicString = ReplaceTranslitRule(ArabicString, Scheme, LearningMode, RuleIndexes)
        For Count = 0 To RuleIndexes.Count - 1
            If RuleIndexes(Count) = -1 Or Count <> RuleIndexes.Count - 1 AndAlso RuleIndexes(Count + 1) = -1 Then Continue For
            If Count = RuleIndexes.Count - 1 Or (ArabicString(Count) = " "c And BreakWords) Then
                If Count <> 0 Then
                    Renderers.Add(New RenderArray.RenderText(RenderArray.RenderDisplayClass.eTransliteration, If((ArabicString(Count) = " "c And BreakWords), ArabicString.Substring(Base, Count - Base), ArabicString.Substring(Base))) With {.Clr = ChData.IslamData.ColorRuleSets(1).ColorRules(RuleIndexes(Count) Mod ChData.IslamData.ColorRuleSets(1).ColorRules.Length).Color})
                    WordRenderers.Add(Renderers.ToArray())
                    Renderers = New List(Of RenderArray.RenderText)
                End If
                Base = Count + 1
            ElseIf RuleIndexes(Count) <> RuleIndexes(Count + 1) Then
                Renderers.Add(New RenderArray.RenderText(RenderArray.RenderDisplayClass.eTransliteration, ArabicString.Substring(Base, Count - Base + 1)) With {.Clr = ChData.IslamData.ColorRuleSets(1).ColorRules(RuleIndexes(Count) Mod ChData.IslamData.ColorRuleSets(1).ColorRules.Length).Color})
                Base = Count + 1
            End If
        Next
        Return WordRenderers.ToArray()
    End Function
    Public Function DecodeTranslitScheme(Str As String) As String
        'QueryString instead of Params?
        Return If(CInt(Str) >= 2, ChData.IslamData.TranslitSchemes(CInt(Str) \ 2).Name, String.Empty)
    End Function
    Public Shared Function DecodeTranslitSchemeType(Str As String) As ArabicData.TranslitScheme
        'QueryString instead of Params?
        Return CType(If(CInt(Str) >= 2, 2 - CInt(Str) Mod 2, CInt(Str)), ArabicData.TranslitScheme)
    End Function
    Function SubOutPatterns(Str As String) As String
        Return Str.Replace(ChData.TehMarbutaStopRule, String.Empty).Replace(ChData.TehMarbutaContinueRule, "...").Replace(ChData.TranslateRegEx("(?=(?:{ArabicMoonLetters}|{ArabicSunLettersNoLam})(?:{ArabicSukun}|{ArabicShadda}?(?:{ArabicFathaDammaKasra}))?(?:{ArabicLetters})(?:{ArabicFatha}|{ArabicKasra})|{ArabicWaslKasraExceptions})", True), ArabicData.ArabicKasra).Replace(ChData.TranslateRegEx("(?=(?:{ArabicMoonLetters}|{ArabicSunLettersNoLam})(?:{ArabicSukun}|{ArabicShadda}?(?:{ArabicFathaDammaKasra}))?(?:{ArabicLetters}){ArabicDamma})", True), ArabicData.ArabicDamma).Replace(ChData.TranslateRegEx("({ArabicSunLetters}|{ArabicMoonLettersNoVowels}|{ArabicLetterWaw}|{ArabicLetterYeh}|{ArabicLetterAlefMaksura}|{ArabicSmallYeh})", True), ChrW(&H66D))
    End Function
    Function GetTransliterationTable(Scheme As String) As List(Of RenderArray.RenderItem)
        Dim Items As New List(Of RenderArray.RenderItem)
        Items.AddRange(Linq.Enumerable.Select(ChData.ArabicLettersInOrder, Function(Letter As String) New RenderArray.RenderItem(RenderArray.RenderTypes.eText, New RenderArray.RenderText() {New RenderArray.RenderText(RenderArray.RenderDisplayClass.eArabic, Letter), New RenderArray.RenderText(RenderArray.RenderDisplayClass.eTransliteration, GetSchemeValueFromSymbol(ArbData.ArabicLetters(ArbData.FindLetterBySymbol(Letter.Chars(0))), Scheme))})))
        Items.AddRange(Linq.Enumerable.Select(ChData.ArabicSpecialLetters, Function(Combo As String) New RenderArray.RenderItem(RenderArray.RenderTypes.eText, New RenderArray.RenderText() {New RenderArray.RenderText(RenderArray.RenderDisplayClass.eArabic, System.Text.RegularExpressions.Regex.Replace(SubOutPatterns(Combo), "\(?\\u([0-9a-fA-F]{4})\)?", Function(Match As System.Text.RegularExpressions.Match) ChrW(Integer.Parse(Match.Groups(1).Value, Globalization.NumberStyles.HexNumber)))), New RenderArray.RenderText(RenderArray.RenderDisplayClass.eTransliteration, GetSchemeSpecialValue(Combo, GetSchemeSpecialFromMatch(Combo, False), Scheme))})))
        Items.AddRange(Linq.Enumerable.Select(ChData.ArabicHamzas, Function(Letter As String) New RenderArray.RenderItem(RenderArray.RenderTypes.eText, New RenderArray.RenderText() {New RenderArray.RenderText(RenderArray.RenderDisplayClass.eArabic, Letter), New RenderArray.RenderText(RenderArray.RenderDisplayClass.eTransliteration, GetSchemeValueFromSymbol(ArbData.ArabicLetters(ArbData.FindLetterBySymbol(Letter.Chars(0))), Scheme))})))
        Items.AddRange(Linq.Enumerable.Select(ChData.ArabicVowels, Function(Letter As String) New RenderArray.RenderItem(RenderArray.RenderTypes.eText, New RenderArray.RenderText() {New RenderArray.RenderText(RenderArray.RenderDisplayClass.eArabic, Letter), New RenderArray.RenderText(RenderArray.RenderDisplayClass.eTransliteration, GetSchemeValueFromSymbol(ArbData.ArabicLetters(ArbData.FindLetterBySymbol(Letter.Chars(0))), Scheme))})))
        Items.AddRange(Linq.Enumerable.Select(ChData.ArabicMultis, Function(Combo As String) New RenderArray.RenderItem(RenderArray.RenderTypes.eText, New RenderArray.RenderText() {New RenderArray.RenderText(RenderArray.RenderDisplayClass.eArabic, Combo), New RenderArray.RenderText(RenderArray.RenderDisplayClass.eTransliteration, GetSchemeLongVowelFromString(Combo, Scheme))})))
        Items.AddRange(Linq.Enumerable.Select(ChData.ArabicTajweed, Function(Letter As String) New RenderArray.RenderItem(RenderArray.RenderTypes.eText, New RenderArray.RenderText() {New RenderArray.RenderText(RenderArray.RenderDisplayClass.eArabic, Letter), New RenderArray.RenderText(RenderArray.RenderDisplayClass.eTransliteration, GetSchemeValueFromSymbol(ArbData.ArabicLetters(ArbData.FindLetterBySymbol(Letter.Chars(0))), Scheme))})))
        Return Items
    End Function
    Public Function ArabicTranslitLetters() As String()
        Dim Lets As New List(Of String)
        Lets.AddRange(Linq.Enumerable.Select(ChData.ArabicLettersInOrder, Function(Ch As String) Ch))
        Lets.AddRange(Linq.Enumerable.Select(ChData.ArabicHamzas, Function(Ch As String) Ch))
        Lets.AddRange(ChData.ArabicVowels)
        Lets.AddRange(Linq.Enumerable.Select(ChData.ArabicTajweed, Function(Ch As String) Ch))
        Lets.AddRange(Linq.Enumerable.Select(ChData.ArabicSilent, Function(Ch As String) Ch))
        Lets.AddRange(Linq.Enumerable.Select(ChData.ArabicPunctuation, Function(Ch As String) Ch))
        Lets.AddRange(Linq.Enumerable.Select(ChData.ArabicNums, Function(Ch As String) Ch))
        Lets.AddRange(Linq.Enumerable.Select(ChData.NonArabicLetters, Function(Ch As String) Ch))
        Return Lets.ToArray()
    End Function
    Public Function GetTransliterationSchemes() As Array()
        Dim Count As Integer
        Dim Strings(ChData.IslamData.TranslitSchemes.Length * 2 + 2 - 2 - 1) As Array
        Strings(0) = New String() {_PortableMethods.LoadResourceString("IslamSource_Off"), "0"}
        Strings(1) = New String() {_PortableMethods.LoadResourceString("IslamSource_ExtendedBuckwalter"), "1"}
        For Count = 0 To ChData.IslamData.TranslitSchemes.Length - 2
            Strings(Count * 2 + 2) = New String() {_PortableMethods.LoadResourceString("IslamSource_" + ChData.IslamData.TranslitSchemes(Count + 1).Name), CStr(Count * 2 + 2)}
            Strings(Count * 2 + 1 + 2) = New String() {_PortableMethods.LoadResourceString("IslamSource_" + ChData.IslamData.TranslitSchemes(Count + 1).Name) + " Literal", CStr(Count * 2 + 1 + 2)}
        Next
        Return Strings
    End Function
    Public Async Function DisplayDict() As Threading.Tasks.Task(Of Array())
        Dim Lines As String() = Await _PortableMethods.ReadAllLines(_PortableMethods.Settings.GetFilePath(_PortableMethods.FileIO.CombinePath("metadata", "HansWeir.txt")))
        Dim Count As Integer
        Dim Words As New List(Of String())
        Words.Add(New String() {})
        Words.Add(New String() {"arabic", String.Empty})
        Words.Add(New String() {_PortableMethods.LoadResourceString("IslamInfo_Arabic"), _PortableMethods.LoadResourceString("IslamInfo_Meaning")})
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
        Return UName.EndsWith("Isolated Form") And Index = 0 Or UName.EndsWith("Final Form") And Index = 1 Or
            UName.EndsWith("Initial Form") And Index = 2 Or UName.EndsWith("Medial Form") And Index = 3
    End Function
    Public Function DisplayCombo(SchemeType As ArabicData.TranslitScheme, Scheme As String) As Array()
        Dim Count As Integer
        Dim Output As New List(Of String())
        Output.Add(New String() {})
        Output.Add(New String() {"arabic", "transliteration", "arabic", String.Empty, String.Empty, String.Empty, String.Empty})
        Output.Add(New String() {_PortableMethods.LoadResourceString("IslamInfo_LetterName"), _PortableMethods.LoadResourceString("IslamInfo_Transliteration"), _PortableMethods.LoadResourceString("IslamInfo_Arabic"), _PortableMethods.LoadResourceString("IslamSource_ExtendedBuckwalter"), _PortableMethods.LoadResourceString("IslamInfo_Terminating"), _PortableMethods.LoadResourceString("IslamInfo_Connecting"), _PortableMethods.LoadResourceString("IslamInfo_Shaping")})
        'Dim Combos(ArabicData.Data.ArabicCombos.Length - 1) As IslamData.ArabicCombo
        'ArabicData.ArabicLetters.CopyTo(ArabicData.Data.ArabicCombos, 0)
        'Array.Sort(Combos, Function(Key As IslamData.ArabicCombo, NextKey As IslamData.ArabicCombo) Key.SymbolName.CompareTo(NextKey.SymbolName))
        For Count = 0 To ArbData.ArabicCombos.Length - 1
            If Linq.Enumerable.All(ArbData.ArabicCombos(Count).Symbol, Function(Ch As Char) GetSchemeValueFromSymbol(ArbData.ArabicLetters(ArbData.FindLetterBySymbol(Ch)), "ExtendedBuckwalter") <> String.Empty) Then
                Dim Str As String = ArabicLetterSpelling(String.Join(String.Empty, Linq.Enumerable.Select(ArbData.ArabicCombos(Count).Symbol, Function(Sym As Char) CStr(Sym))), False, False, False)
                Output.Add(New String() {Str,
                                         TransliterateToScheme(Str, SchemeType, Scheme, GetMetarules(Str, ChData.RuleMetas("Normal"))),
                                                    String.Join(String.Empty, Linq.Enumerable.Select(ArbData.ArabicCombos(Count).Symbol, Function(Sym As Char) CStr(Sym))),
                                                    TransliterateToScheme(String.Join(String.Empty, Linq.Enumerable.Select(ArbData.ArabicCombos(Count).Symbol, Function(Sym As Char) CStr(Sym))), ArabicData.TranslitScheme.Literal, String.Empty, Nothing),
                                                    CStr(ArbData.ArabicCombos(Count).Terminating(ArbData)),
                                                  CStr(ArbData.ArabicCombos(Count).Connecting(ArbData)),
                                                    String.Join(vbCrLf, Linq.Enumerable.Select(ArbData.ArabicCombos(Count).Shaping, Function(Shape As Char) If(Shape = ChrW(0), String.Empty, Shape + " " + CStr(Convert.ToString(AscW(Shape), 16)) + " " + If(CheckShapingOrder(Array.IndexOf(ArbData.ArabicCombos(Count).Shaping, Shape), ArbData.GetUnicodeName(Shape)), String.Empty, "!!!") + ArbData.GetUnicodeName(Shape))))})
            End If
        Next
        Return Output.ToArray()
    End Function
    Public Function SymbolDisplay(Symbols() As ArabicData.ArabicSymbol, SchemeType As ArabicData.TranslitScheme, Scheme As String, Cols As String()) As Array()
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
        Output(2) = New List(Of String)(Linq.Enumerable.Select(Cols, Function(Str As String) _PortableMethods.LoadResourceString("IslamInfo_" + Str))).ToArray()
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
                            Dim SpellStr As String = ArabicLetterSpelling(CStr(Symbols(Count).Symbol), False, False, False)
                            Return TransliterateToScheme(SpellStr, SchemeType, Scheme, Arabic.FilterMetadataStops(SpellStr, GetMetarules(SpellStr, ChData.RuleMetas("Normal")), Nothing))
                        Case 2
                            Return ArbData.GetUnicodeName(Symbols(Count).Symbol)
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
                            Return If(Symbols(Count).Shaping = Nothing, String.Empty, String.Join(vbCrLf, Linq.Enumerable.Select(Symbols(Count).Shaping, Function(Shape As Char) If(Shape = ChrW(0), String.Empty, Shape + " " + CStr(Convert.ToString(AscW(Shape), 16)) + " " + If(CheckShapingOrder(Array.IndexOf(Symbols(Count).Shaping, Shape), ArbData.GetUnicodeName(Shape)), String.Empty, "!!!") + ArbData.GetUnicodeName(Shape)))))
                        Case Else
                            Return Nothing
                    End Select
                End Function)).ToArray()
        Next
        Return Output
    End Function
    Public Function DisplayAll(SchemeType As ArabicData.TranslitScheme, Scheme As String) As Array()
        Return SymbolDisplay(New List(Of ArabicData.ArabicSymbol)(Linq.Enumerable.Where(ArbData.ArabicLetters, Function(Letter As ArabicData.ArabicSymbol) GetSchemeValueFromSymbol(Letter, "ExtendedBuckwalter") <> String.Empty)).ToArray(), SchemeType, Scheme, Nothing)
    End Function
    Public Function DisplayTranslitSchemes() As Array()
        Dim Count As Integer
        Dim Output As New List(Of String())
        'Dim oFont As New Font(DefaultValue(HttpContext.Current.Request.QueryString.Get("fontcustom"), "Arial"), 13)
        'CheckIfCharInFont(ArabicData.ArabicLetters(Count).Symbol, oFont)
        Output.Add(New String() {})
        Dim Strs As String() = New String() {"arabic", String.Empty, "arabic"}
        Array.Resize(Of String)(Strs, 3 + ChData.IslamData.TranslitSchemes.Length)
        Output.Add(Strs)
        Strs = New String() {_PortableMethods.LoadResourceString("IslamInfo_LetterName"), _PortableMethods.LoadResourceString("IslamInfo_UnicodeName"), _PortableMethods.LoadResourceString("IslamInfo_Arabic")}
        Array.Resize(Of String)(Strs, 3 + ChData.IslamData.TranslitSchemes.Length)
        For SchemeCount = 0 To ChData.IslamData.TranslitSchemes.Length - 1
            CType(Output(1), String())(3 + SchemeCount) = String.Empty
            Strs(3 + SchemeCount) = _PortableMethods.LoadResourceString("IslamSource_" + ChData.IslamData.TranslitSchemes(SchemeCount).Name)
        Next
        Output.Add(Strs)
        For Count = 0 To ArbData.ArabicLetters.Length - 1
            If GetSchemeValueFromSymbol(ArbData.ArabicLetters(Count), "ExtendedBuckwalter") <> String.Empty Then
                Strs = New String() {ArabicLetterSpelling(ArbData.ArabicLetters(Count).Symbol, False, False, False),
                                           ArbData.GetUnicodeName(ArbData.ArabicLetters(Count).Symbol),
                                           CStr(ArbData.ArabicLetters(Count).Symbol)}
                Array.Resize(Of String)(Strs, 3 + ChData.IslamData.TranslitSchemes.Length)
                For SchemeCount = 0 To ChData.IslamData.TranslitSchemes.Length - 1
                    Strs(3 + SchemeCount) = GetSchemeValueFromSymbol(ArbData.ArabicLetters(Count), ChData.IslamData.TranslitSchemes(SchemeCount).Name)
                Next
                Output.Add(Strs)
            End If
        Next
        For Count = 0 To ChData.ArabicSpecialLetters.Length - 1
            Dim Str As String = System.Text.RegularExpressions.Regex.Replace(SubOutPatterns(ChData.ArabicSpecialLetters(Count)), "\(?\\u([0-9a-fA-F]{4})\)?", Function(Match As System.Text.RegularExpressions.Match) ChrW(Integer.Parse(Match.Groups(1).Value, Globalization.NumberStyles.HexNumber)))
            Strs = New String() {ArabicLetterSpelling(Str, False, False, False), String.Empty, Str,
                                       TransliterateToScheme(Str, ArabicData.TranslitScheme.Literal, String.Empty, Nothing)}
            Array.Resize(Of String)(Strs, 3 + ChData.IslamData.TranslitSchemes.Length)
            For SchemeCount = 0 To ChData.IslamData.TranslitSchemes.Length - 1
                Strs(3 + SchemeCount) = GetSchemeSpecialValue(ChData.ArabicSpecialLetters(Count), GetSchemeSpecialFromMatch(ChData.ArabicSpecialLetters(Count), False), ChData.IslamData.TranslitSchemes(SchemeCount).Name)
            Next
            Output.Add(Strs)
        Next
        For Count = 0 To ChData.ArabicLongVowels.Length - 1
            Strs = New String() {ArabicLetterSpelling(ChData.ArabicLongVowels(Count), False, False, False),
                                       String.Empty, ChData.ArabicLongVowels(Count),
                                       TransliterateToScheme(ChData.ArabicLongVowels(Count), ArabicData.TranslitScheme.Literal, String.Empty, Nothing)}
            Array.Resize(Of String)(Strs, 3 + ChData.IslamData.TranslitSchemes.Length)
            For SchemeCount = 0 To ChData.IslamData.TranslitSchemes.Length - 1
                Strs(3 + SchemeCount) = GetSchemeLongVowelFromString(ChData.ArabicLongVowels(Count), ChData.IslamData.TranslitSchemes(SchemeCount).Name)
            Next
            Output.Add(Strs)
        Next
        Return Output.ToArray()
    End Function
    Public Function GetCatWords(SelArr As String()) As IslamData.GrammarSet.GrammarWord()
        Dim Words As New List(Of IslamData.GrammarSet.GrammarWord)
        For Count = 0 To SelArr.Length - 1
            Dim Word As Nullable(Of IslamData.GrammarSet.GrammarWord)
            Word = GetCatWord(SelArr(Count))
            If Word.HasValue Then Words.Add(Word.Value)
        Next
        Return Words.ToArray()
    End Function
    Public Function GetCatWord(ID As String) As Nullable(Of IslamData.GrammarSet.GrammarWord)
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
    Public ReadOnly Property NounIDs As Dictionary(Of String, List(Of IslamData.GrammarSet.GrammarNoun))
        Get
            Return _NounIDs
        End Get
    End Property
    Public Function GetCatNoun(ID As String) As IslamData.GrammarSet.GrammarNoun()
        Return If(NounIDs.ContainsKey(ID), NounIDs(ID).ToArray(), Nothing)
    End Function
    Shared _TransformIDs As Dictionary(Of String, List(Of IslamData.GrammarSet.GrammarTransform))
    Public ReadOnly Property TransformIDs As Dictionary(Of String, List(Of IslamData.GrammarSet.GrammarTransform))
        Get
            Return _TransformIDs
        End Get
    End Property
    Public Function GetTransform(ID As String) As IslamData.GrammarSet.GrammarTransform()
        Return If(TransformIDs.ContainsKey(ID), TransformIDs(ID).ToArray(), Nothing)
    End Function
    Public Function GetTransformMatch(IDs As String()) As IslamData.GrammarSet.GrammarTransform()
        Dim Transforms As IslamData.GrammarSet.GrammarTransform()() = New List(Of IslamData.GrammarSet.GrammarTransform())(Linq.Enumerable.Select(IDs, Function(ID As String) GetTransform(ID))).ToArray()
        'union all the results together
        Return New List(Of IslamData.GrammarSet.GrammarTransform)(Linq.Enumerable.Where(Transforms(0), Function(Tr As IslamData.GrammarSet.GrammarTransform) Linq.Enumerable.All(Transforms, Function(Trs As IslamData.GrammarSet.GrammarTransform()) Array.IndexOf(Trs, Tr) <> -1))).ToArray()
    End Function
    Shared _ParticleIDs As Dictionary(Of String, List(Of IslamData.GrammarSet.GrammarParticle))
    Public ReadOnly Property ParticleIDs As Dictionary(Of String, List(Of IslamData.GrammarSet.GrammarParticle))
        Get
            Return _ParticleIDs
        End Get
    End Property
    Public Function GetParticles(ID As String) As IslamData.GrammarSet.GrammarParticle()
        Return If(ParticleIDs.ContainsKey(ID), ParticleIDs(ID).ToArray(), Nothing)
    End Function
    Shared _VerbIDs As Dictionary(Of String, List(Of IslamData.GrammarSet.GrammarVerb))
    Public ReadOnly Property VerbIDs As Dictionary(Of String, List(Of IslamData.GrammarSet.GrammarVerb))
        Get
            Return _VerbIDs
        End Get
    End Property
    Public Function GetVerb(ID As String) As IslamData.GrammarSet.GrammarVerb()
        Return If(VerbIDs.ContainsKey(ID), VerbIDs(ID).ToArray(), Nothing)
    End Function
    Public Function ApplyTransform(Transforms As IslamData.GrammarSet.GrammarTransform(), Str As String) As String
        Dim Text As String = Str
        For Count = 0 To Transforms.Length - 1
            Text = New System.Text.RegularExpressions.Regex(If(Transforms(Count).From Is Nothing, "$", ChData.TranslateRegEx(Transforms(Count).From, True))).Replace(Text, ChData.TranslateRegEx(Transforms(Count).Text, False), 1)
        Next
        Return Text
    End Function
End Class
Public Class AudioRecitation
    Public Shared Function GetURL(Source As String, ReciterName As String, Chapter As Integer, Verse As Integer) As String
        Dim Base As String = String.Empty
        If Source = "everyayah" Then Base = "http://www.everyayah.com/data/"
        If Source = "tanzil" Then Base = "http://tanzil.net/res/audio/"
        Return Base + ReciterName + "/" + Chapter.ToString("D3") + Verse.ToString("D3") + ".mp3"
    End Function
    Public Shared Function GetReciterIndex(ReciterName As String, ReciterList As IslamData.Reciters) As Integer
        If ReciterName = String.Empty Then ReciterName = ReciterList.DefaultReciter
        For Count As Integer = 0 To ReciterList.Reciters.Length - 1
            If ReciterList.Reciters(Count).Name = ReciterName Then
                Return Count
            End If
        Next
        Return -1
    End Function
End Class
<Xml.Serialization.XmlRoot("islamdata")>
Public Class IslamData
    Public Structure Reciters
        Public Structure Reciter
            <Xml.Serialization.XmlAttribute("name")>
            Public Name As String
            <Xml.Serialization.XmlAttribute("reciter")>
            Public Reciter As String
            <Xml.Serialization.XmlAttribute("source")>
            Public Source As String
            <Xml.Serialization.XmlAttribute("bitrate")>
            <ComponentModel.DefaultValue(0)>
            Public BitRate As Integer
        End Structure
        <Xml.Serialization.XmlElement("reciter")>
        Public Reciters() As Reciter
        <Xml.Serialization.XmlAttribute("default")>
        Public DefaultReciter As String
    End Structure
    <Xml.Serialization.XmlElement("reciters")>
    Public ReciterList As Reciters
    Public Structure LoopingModes
        Public Structure LoopingMode
            <Xml.Serialization.XmlAttribute("name")>
            Public Name As String
        End Structure
        <Xml.Serialization.XmlElement("loopingmode")>
        Public LoopingModes() As LoopingMode
        <Xml.Serialization.XmlAttribute("default")>
        Public DefaultLoopingMode As String
    End Structure
    <Xml.Serialization.XmlElement("loopingmodes")>
    Public LoopingModeList As LoopingModes
    Public Structure ListCategory
        Public Structure Word
            <Xml.Serialization.XmlAttribute("text")>
            Public Text As String
            <Xml.Serialization.XmlAttribute("id")>
            Public TranslationID As String
        End Structure
        <Xml.Serialization.XmlAttribute("title")>
        Public Title As String
        <Xml.Serialization.XmlElement("word")>
        Public Words() As Word
    End Structure
    <Xml.Serialization.XmlArray("lists")>
    <Xml.Serialization.XmlArrayItem("category")>
    Public Lists() As ListCategory

    Public Structure Phrase
        <Xml.Serialization.XmlAttribute("text")>
        Public Text As String
        <Xml.Serialization.XmlAttribute("id")>
        Public TranslationID As String
    End Structure
    <Xml.Serialization.XmlArray("phrases")>
    <Xml.Serialization.XmlArrayItem("word")>
    Public Phrases() As Phrase

    Public Structure AbbrevWord
        <Xml.Serialization.XmlAttribute("text")>
        Public Text As String
        <Xml.Serialization.XmlAttribute("font")>
        Public Font As String
        <Xml.Serialization.XmlAttribute("id")>
        Public TranslationID As String
    End Structure
    <Xml.Serialization.XmlArray("abbreviations")>
    <Xml.Serialization.XmlArrayItem("word")>
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
            <Xml.Serialization.XmlAttribute("text")>
            Public Text As String
            <Xml.Serialization.XmlAttribute("id")>
            Public TranslationID As String
            <Xml.Serialization.XmlAttribute("grammar")>
            Public Grammar As String
            <Xml.Serialization.XmlAttribute("match")>
            Public Match As String
            <Xml.Serialization.XmlAttribute("from")>
            Public From As String
        End Structure
        <Xml.Serialization.XmlArray("transforms")>
        <Xml.Serialization.XmlArrayItem("transform")>
        Public Transforms() As GrammarTransform
        Public Structure GrammarParticle
            <Xml.Serialization.XmlAttribute("text")>
            Public Text As String
            <Xml.Serialization.XmlAttribute("id")>
            Public TranslationID As String
            <Xml.Serialization.XmlAttribute("grammar")>
            Public Grammar As String
        End Structure
        <Xml.Serialization.XmlArray("particles")>
        <Xml.Serialization.XmlArrayItem("particle")>
        Public Particles() As GrammarParticle
        Public Structure GrammarNoun
            <Xml.Serialization.XmlAttribute("text")>
            Public Text As String
            <Xml.Serialization.XmlAttribute("id")>
            Public TranslationID As String
            <Xml.Serialization.XmlAttribute("plural")>
            Public Plural As String
            <Xml.Serialization.XmlAttribute("grammar")>
            Public Grammar As String
        End Structure
        <Xml.Serialization.XmlArray("nouns")>
        <Xml.Serialization.XmlArrayItem("noun")>
        Public Nouns() As GrammarNoun
        Public Structure GrammarVerb
            <Xml.Serialization.XmlAttribute("text")>
            Public Text As String
            <Xml.Serialization.XmlAttribute("id")>
            Public TranslationID As String
            <Xml.Serialization.XmlAttribute("poss")>
            Public Possessives As String
            <Xml.Serialization.XmlAttribute("grammar")>
            Public Grammar As String
        End Structure
        <Xml.Serialization.XmlArray("verbs")>
        <Xml.Serialization.XmlArrayItem("verb")>
        Public Verbs() As GrammarVerb
    End Structure
    <Xml.Serialization.XmlElement("grammar")>
    Public Grammar As GrammarSet

    Public Structure TranslitScheme
        <Xml.Serialization.XmlAttribute("name")>
        Public Name As String
        Public Alphabet() As String
        <Xml.Serialization.XmlAttribute("alphabet")>
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
        <Xml.Serialization.XmlAttribute("hamza")>
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
        <Xml.Serialization.XmlAttribute("literals")>
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
        <Xml.Serialization.XmlAttribute("fathadammakasratanweenlongvowelsdipthongsshaddasukun")>
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
        <Xml.Serialization.XmlAttribute("multis")>
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
        <Xml.Serialization.XmlAttribute("gutterals")>
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
        <Xml.Serialization.XmlAttribute("tajweed")>
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
        <Xml.Serialization.XmlAttribute("silent")>
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
        <Xml.Serialization.XmlAttribute("punctuation")>
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
        <Xml.Serialization.XmlAttribute("number")>
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
        <Xml.Serialization.XmlAttribute("nonarabic")>
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
    <Xml.Serialization.XmlArray("translitschemes")>
    <Xml.Serialization.XmlArrayItem("scheme")>
    Public TranslitSchemes() As TranslitScheme
    Structure ArabicCapInfo
        <Xml.Serialization.XmlAttribute("name")>
        Public Name As String
        Public Text As String()
        <Xml.Serialization.XmlAttribute("text")>
        Public Property _Text As String
            Get
                Return String.Join("  ", Text)
            End Get
            Set(value As String)
                Text = value.Split({"  "}, StringSplitOptions.None)
            End Set
        End Property
    End Structure
    <Xml.Serialization.XmlArray("arabiccaptures")>
    <Xml.Serialization.XmlArrayItem("caps")>
    Public ArabicCaptures() As ArabicCapInfo
    Structure ArabicNumInfo
        <Xml.Serialization.XmlAttribute("name")>
        Public Name As String
        Public Text As String()
        <Xml.Serialization.XmlAttribute("text")>
        Public Property _Text As String
            Get
                Return String.Join(" "c, Text)
            End Get
            Set(value As String)
                Text = value.Split(" "c)
            End Set
        End Property
    End Structure
    <Xml.Serialization.XmlArray("arabicnumbers")>
    <Xml.Serialization.XmlArrayItem("nums")>
    Public ArabicNumbers() As ArabicNumInfo
    Structure ArabicPattern
        <Xml.Serialization.XmlAttribute("name")>
        Public Name As String
        <Xml.Serialization.XmlAttribute("match")>
        Public Match As String
    End Structure
    <Xml.Serialization.XmlArray("arabicpatterns")>
    <Xml.Serialization.XmlArrayItem("pattern")>
    Public ArabicPatterns() As ArabicPattern
    Structure ArabicGroup
        <Xml.Serialization.XmlAttribute("name")>
        Public Name As String
        Public Text As String()
        <Xml.Serialization.XmlAttribute("text")>
        Public Property _Text As String
            Get
                Return String.Join(" "c, Text)
            End Get
            Set(value As String)
                Text = value.Split(" "c)
            End Set
        End Property
    End Structure
    <Xml.Serialization.XmlArray("arabicgroups")>
    <Xml.Serialization.XmlArrayItem("group")>
    Public ArabicGroups() As ArabicGroup
    Structure ColorRuleCategory
        Structure ColorRule
            <Xml.Serialization.XmlAttribute("name")>
            Public Name As String
            Public Match As String()
            <Xml.Serialization.XmlAttribute("match")>
            Public Property _Match As String
                Get
                    Return String.Join("|"c, Match)
                End Get
                Set(value As String)
                    Match = value.Split("|"c)
                End Set
            End Property
            Public Color As Integer
            Private __Color As String
            <Xml.Serialization.XmlAttribute("color")>
            Public Property _Color As String
                Get
                    Return __Color
                End Get
                Set(value As String)
                    __Color = value
                    If value.Contains(",") Then
                        Dim RGB As Byte() = New List(Of Byte)(Linq.Enumerable.Select(value.Split(","c), Function(Str As String) CByte(Str))).ToArray()
                        Color = Utility.MakeArgb(&HFF, RGB(0), RGB(1), RGB(2))
                    Else
                        Color = Utility.ColorFromName(value)
                    End If
                End Set
            End Property
            <Xml.Serialization.XmlAttribute("evaluator")>
            Public Property Evaluator As String
            Public MetaRuleFunc As Arabic.MetaRuleFuncs
            Private __MetaRuleFunc As String
            <Xml.Serialization.XmlAttribute("rulefunc")>
            Public Property _MetaRuleFunc As String
                Get
                    Return __MetaRuleFunc
                End Get
                Set(value As String)
                    __MetaRuleFunc = value
                    Select Case value
                        Case "eLearningMode"
                            MetaRuleFunc = Arabic.MetaRuleFuncs.eLearningMode
                            Exit Select
                        Case "eDivideTanween"
                            MetaRuleFunc = Arabic.MetaRuleFuncs.eDivideTanween
                            Exit Select
                        Case "eSpellLetter"
                            MetaRuleFunc = Arabic.MetaRuleFuncs.eSpellLetter
                            Exit Select
                        Case "eSpellLongLetter"
                            MetaRuleFunc = Arabic.MetaRuleFuncs.eSpellLongLetter
                            Exit Select
                        Case "eSpellLongMergedLetter"
                            MetaRuleFunc = Arabic.MetaRuleFuncs.eSpellLongMergedLetter
                            Exit Select
                        Case "eSpellNumber"
                            MetaRuleFunc = Arabic.MetaRuleFuncs.eSpellNumber
                            Exit Select
                        Case "eUpperCase"
                            MetaRuleFunc = Arabic.MetaRuleFuncs.eUpperCase
                            Exit Select
                        Case "eObligatory"
                            MetaRuleFunc = Arabic.MetaRuleFuncs.eObligatory
                            Exit Select
                            'If _RuleFunc = "eNone" Then
                        Case Else
                            MetaRuleFunc = Arabic.MetaRuleFuncs.eNone
                    End Select
                End Set
            End Property
        End Structure
        <Xml.Serialization.XmlElement("colorrule")>
        Public ColorRules() As ColorRule
    End Structure
    <Xml.Serialization.XmlArray("colorrulesets")>
    <Xml.Serialization.XmlArrayItem("colorruleset")>
    Public ColorRuleSets() As ColorRuleCategory

    Structure RuleTranslationCategory
        Structure RuleTranslation
            <Xml.Serialization.XmlAttribute("name")>
            Public Name As String
            <Xml.Serialization.XmlAttribute("match")>
            Public Match As String

            <Xml.Serialization.XmlAttribute("evaluator")>
            Public Evaluator As String
            <Xml.Serialization.XmlAttribute("negativematch")>
            Public NegativeMatch As String
            Public RuleFunc As Arabic.RuleFuncs
            Public __RuleFunc As String
            <Xml.Serialization.XmlAttribute("rulefunc")>
            Public Property _RuleFunc As String
                Get
                    Return __RuleFunc
                End Get
                Set(value As String)
                    __RuleFunc = value
                    Select Case value
                        Case "eLeadingGutteral"
                            RuleFunc = Arabic.RuleFuncs.eLeadingGutteral
                            Exit Select
                        Case "eLookupLetter"
                            RuleFunc = Arabic.RuleFuncs.eLookupLetter
                            Exit Select
                        Case "eLookupLongVowelDipthong"
                            RuleFunc = Arabic.RuleFuncs.eLookupLongVowelDipthong
                            Exit Select
                        Case "eTrailingGutteral"
                            RuleFunc = Arabic.RuleFuncs.eTrailingGutteral
                            Exit Select
                        Case "eResolveAmbiguity"
                            RuleFunc = Arabic.RuleFuncs.eResolveAmbiguity
                            Exit Select
                        Case Else
                            'If _RuleFunc = "eNone" Then
                            RuleFunc = Arabic.RuleFuncs.eNone
                    End Select
                End Set
            End Property
        End Structure
        <Xml.Serialization.XmlAttribute("name")>
        Public Name As String
        <Xml.Serialization.XmlElement("rule")>
        Public Rules() As RuleTranslation
    End Structure
    <Xml.Serialization.XmlArray("translitrules")>
    <Xml.Serialization.XmlArrayItem("ruleset")>
    Public RuleSets() As RuleTranslationCategory
    Structure VerificationData
        <Xml.Serialization.XmlAttribute("name")>
        Public Name As String
        <Xml.Serialization.XmlAttribute("match")>
        Public Match As String
        <Xml.Serialization.XmlAttribute("evaluator")>
        Public Evaluator As String()
        Public Property _Evaluator As String
            Get
                Return String.Join("|"c, Evaluator)
            End Get
            Set(value As String)
                Evaluator = value.Split("|"c)
            End Set
        End Property
        <Xml.Serialization.XmlAttribute("metarules")>
        Public MetaRules As String()
        Public Property _MetaRules As String
            Get
                Return String.Join("|"c, MetaRules)
            End Get
            Set(value As String)
                MetaRules = value.Split("|"c)
            End Set
        End Property
    End Structure
    <Xml.Serialization.XmlArray("verificationset")>
    <Xml.Serialization.XmlArrayItem("verification")>
    Public VerificationSet() As VerificationData
    Structure RuleMetaSet
        Structure RuleMetadataTranslation
            <Xml.Serialization.XmlAttribute("name")>
            Public Name As String
            <Xml.Serialization.XmlAttribute("match")>
            Public Match As String
            <Xml.Serialization.XmlAttribute("evaluator")>
            Public _Evaluator As String
            Structure RuleWithArgs
                Public RuleName As String
                Public Args As String()()
            End Structure
            Public _SplitEvaluator As RuleWithArgs()()
            ReadOnly Property Evaluator As RuleWithArgs()()
                Get
                    If _SplitEvaluator Is Nothing And Not _Evaluator Is Nothing Then _SplitEvaluator = New List(Of RuleWithArgs())(Linq.Enumerable.Select(_Evaluator.Split(";"c), Function(Str) New List(Of RuleWithArgs)(Linq.Enumerable.Select(Str.Split("|"c), Function(S) New RuleWithArgs With {.RuleName = System.Text.RegularExpressions.Regex.Replace(S, "\(.*\)|^null$", String.Empty), .Args = If(System.Text.RegularExpressions.Regex.Match(S, "\((.*)\)").Success, New List(Of String())(Linq.Enumerable.Select(System.Text.RegularExpressions.Regex.Match(S, "\((.*)\)").Groups(1).Value.Split(","c), Function(InnerStr) InnerStr.Split(" "c))).ToArray(), Nothing)})).ToArray())).ToArray()
                    Return _SplitEvaluator
                End Get
            End Property
            Public _OptionalStopIndexes As Integer()
            Public _OptionalNotStopIndexes As Integer()
            Public Sub InitIndexes()
                Dim StopIndexes As New List(Of Integer)
                Dim NotStopIndexes As New List(Of Integer)
                For Count = 0 To Evaluator.Length - 1
                    Dim EvaluatorParts As RuleWithArgs() = Evaluator(Count)
                    If Linq.Enumerable.Any(EvaluatorParts, Function(RuleWithArgs) RuleWithArgs.RuleName = "optionalstop") Then StopIndexes.Add(Count)
                    If Linq.Enumerable.Any(EvaluatorParts, Function(RuleWithArgs) RuleWithArgs.RuleName = "optionalnotstop") Then NotStopIndexes.Add(Count)
                Next
                _OptionalStopIndexes = StopIndexes.ToArray()
                _OptionalNotStopIndexes = NotStopIndexes.ToArray()
            End Sub
            ReadOnly Property OptionalStopIndexes As Integer()
                Get
                    If _OptionalStopIndexes Is Nothing Then InitIndexes()
                    Return _OptionalStopIndexes
                End Get
            End Property
            ReadOnly Property OptionalNotStopIndexes As Integer()
                Get
                    If _OptionalNotStopIndexes Is Nothing Then InitIndexes()
                    Return _OptionalNotStopIndexes
                End Get
            End Property
        End Structure
        <Xml.Serialization.XmlAttribute("from")>
        Public From As String
        <Xml.Serialization.XmlAttribute("name")>
        Public Name As String
        <Xml.Serialization.XmlElement("metarule")>
        Public Rules() As RuleMetadataTranslation
    End Structure
    <Xml.Serialization.XmlArray("metarules")>
    <Xml.Serialization.XmlArrayItem("metaruleset")>
    Public MetaRules() As RuleMetaSet
    Structure LanguageInfo
        <Xml.Serialization.XmlAttribute("code")>
        Public Code As String
        <Xml.Serialization.XmlAttribute("rtl")>
        Public IsRTL As Boolean
    End Structure
    <Xml.Serialization.XmlArray("languages")>
    <Xml.Serialization.XmlArrayItem("language")>
    Public LanguageList() As LanguageInfo

    Structure ArabicFontList
        <Xml.Serialization.XmlAttribute("id")>
        Public ID As String
        <Xml.Serialization.XmlAttribute("name")>
        Public Name As String
        <Xml.Serialization.XmlAttribute("family")>
        Public Family As String
        <Xml.Serialization.XmlAttribute("embedname")>
        Public EmbedName As String
        <Xml.Serialization.XmlAttribute("file")>
        Public FileName As String
        <Xml.Serialization.XmlAttribute("scale")>
        <ComponentModel.DefaultValueAttribute(-1.0)>
        Public Scale As Double
    End Structure
    <Xml.Serialization.XmlArray("arabicfonts")>
    <Xml.Serialization.XmlArrayItem("arabicfont")>
    Public ArabicFonts() As ArabicFontList

    Public Structure ScriptFont
        Public Structure Font
            <Xml.Serialization.XmlAttribute("id")>
            Public ID As String
        End Structure
        <Xml.Serialization.XmlAttribute("name")>
        Public Name As String
        <Xml.Serialization.XmlElement("font")>
        Public FontList() As Font
    End Structure

    <Xml.Serialization.XmlArray("scriptfonts")>
    <Xml.Serialization.XmlArrayItem("scriptfont")>
    Public ScriptFonts() As ScriptFont
    Public Class LanguageDefaults
        Public Class LanguageDefault
            <Xml.Serialization.XmlAttribute("language")>
            Public ID As String
            <Xml.Serialization.XmlAttribute("quran")>
            Public QuranFile As String
            <Xml.Serialization.XmlAttribute("quranw4w")>
            Public QuranW4WFile As String
            <Xml.Serialization.XmlAttribute("transliteration")>
            Public TranslitScheme As String
        End Class
        <Xml.Serialization.XmlAttribute("default")>
        Public DefaultLanguage As String
        <Xml.Serialization.XmlElement("languagedefault")>
        Public LanguageDefaultList() As LanguageDefault
        Public Function GetLanguageByID(LangID As String) As LanguageDefault
            For Count As Integer = 0 To LanguageDefaultList.Length - 1
                If LanguageDefaultList(Count).ID = If(LangID = String.Empty, System.Threading.Thread.CurrentThread.CurrentUICulture.Name, LangID) Then Return LanguageDefaultList(Count)
            Next
            If LangID = String.Empty Then
                For Count As Integer = 0 To LanguageDefaultList.Length - 1
                    If LanguageDefaultList(Count).ID = System.Threading.Thread.CurrentThread.CurrentUICulture.TwoLetterISOLanguageName Then Return LanguageDefaultList(Count)
                Next
                For Count As Integer = 0 To LanguageDefaultList.Length - 1
                    If LanguageDefaultList(Count).ID = System.Threading.Thread.CurrentThread.CurrentUICulture.Parent.Name Then Return LanguageDefaultList(Count)
                Next
                For Count As Integer = 0 To LanguageDefaultList.Length - 1
                    If LanguageDefaultList(Count).ID = System.Threading.Thread.CurrentThread.CurrentUICulture.Parent.TwoLetterISOLanguageName Then Return LanguageDefaultList(Count)
                Next
            End If
            Return If(LangID = String.Empty, LanguageDefaultList(0), GetLanguageByID(String.Empty))
        End Function
    End Class
    <Xml.Serialization.XmlElement("languagedefaults")>
    Public LanguageDefaultInfo As LanguageDefaults
    Public Class TranslationsInfo
        Public Structure TranslationInfo
            <Xml.Serialization.XmlAttribute("name")>
            Public Name As String
            <Xml.Serialization.XmlAttribute("file")>
            Public FileName As String
            <Xml.Serialization.XmlAttribute("translator")>
            Public Translator As String
        End Structure
        <Xml.Serialization.XmlElement("translation")>
        Public TranslationList() As TranslationInfo
    End Class
    Structure VerseNumberScheme
        <Xml.Serialization.XmlAttribute("name")>
        Public Name As String
        Public ExtraVerses As Integer()()
        <Xml.Serialization.XmlAttribute("extraverses")>
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
        <Xml.Serialization.XmlAttribute("combinedverses")>
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
    <Xml.Serialization.XmlArray("versenumberschemes")>
    <Xml.Serialization.XmlArrayItem("versenumberscheme")>
    Public VerseNumberSchemes As VerseNumberScheme()

    <Xml.Serialization.XmlElement("translations")>
    Public Translations As TranslationsInfo
    Structure QuranSelection
        Structure QuranSelectionInfo
            <Xml.Serialization.XmlAttribute("chapter")>
            Public ChapterNumber As Integer
            <Xml.Serialization.XmlAttribute("startverse")>
            Public VerseNumber As Integer
            <ComponentModel.DefaultValueAttribute(1)>
            <Xml.Serialization.XmlAttribute("startword")>
            Public WordNumber As Integer
            <ComponentModel.DefaultValueAttribute(0)>
            <Xml.Serialization.XmlAttribute("endword")>
            Public EndWordNumber As Integer
            <ComponentModel.DefaultValueAttribute(0)>
            <Xml.Serialization.XmlAttribute("endverse")>
            Public ExtraVerseNumber As Integer
        End Structure
        <Xml.Serialization.XmlAttribute("description")>
        Public Description As String
        <Xml.Serialization.XmlElement("verse")>
        Public SelectionInfo As QuranSelectionInfo()
    End Structure
    <Xml.Serialization.XmlArray("quranselections")>
    <Xml.Serialization.XmlArrayItem("quranselection")>
    Public QuranSelections As QuranSelection()
    Structure QuranDivision
        <Xml.Serialization.XmlAttribute("description")>
        Public Description As String
        <Xml.Serialization.XmlAttribute("arabic")>
        Public Arabic As String
    End Structure
    <Xml.Serialization.XmlArray("qurandivisions")>
    <Xml.Serialization.XmlArrayItem("division")>
    Public QuranDivisions As QuranDivision()

    Structure QuranChapter
        <Xml.Serialization.XmlAttribute("index")>
        Public Index As String
        <Xml.Serialization.XmlAttribute("name")>
        Public Name As String
        <ComponentModel.DefaultValueAttribute(0)>
        <Xml.Serialization.XmlAttribute("uniqueletters")>
        Public UniqueLetters As Integer
    End Structure
    <Xml.Serialization.XmlArray("quranchapters")>
    <Xml.Serialization.XmlArrayItem("chapter")>
    Public QuranChapters As QuranChapter()

    Structure QuranPart
        <Xml.Serialization.XmlAttribute("index")>
        Public Index As String
        <Xml.Serialization.XmlAttribute("name")>
        Public Name As String
        <Xml.Serialization.XmlAttribute("id")>
        Public ID As String
    End Structure
    <Xml.Serialization.XmlArray("quranparts")>
    <Xml.Serialization.XmlArrayItem("part")>
    Public QuranParts As QuranPart()

    Structure CollectionInfo
        Structure CollTranslationInfo
            <Xml.Serialization.XmlAttribute("name")>
            Public Name As String
            <Xml.Serialization.XmlAttribute("file")>
            Public FileName As String
        End Structure
        <Xml.Serialization.XmlAttribute("name")>
        Public Name As String
        <Xml.Serialization.XmlAttribute("file")>
        Public FileName As String
        <Xml.Serialization.XmlAttribute("default")>
        Public DefaultTranslation As String
        <Xml.Serialization.XmlArray("translations")>
        <Xml.Serialization.XmlArrayItem("translation")>
        Public Translations() As CollTranslationInfo
    End Structure
    <Xml.Serialization.XmlArray("hadithcollections")>
    <Xml.Serialization.XmlArrayItem("collection")>
    Public Collections() As CollectionInfo
    Structure PartOfSpeechInfo
        <Xml.Serialization.XmlAttribute("symbol")>
        Public Symbol As String
        <Xml.Serialization.XmlAttribute("id")>
        Public Id As String
        <Xml.Serialization.XmlAttribute("color")>
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
    <Xml.Serialization.XmlArray("partsofspeech")>
    <Xml.Serialization.XmlArrayItem("pos")>
    Public PartsOfSpeech() As PartOfSpeechInfo
    Structure FeatureOfSpeechInfo
        <Xml.Serialization.XmlAttribute("symbol")>
        Public Symbol As String
        <Xml.Serialization.XmlAttribute("id")>
        Public Id As String
    End Structure
    <Xml.Serialization.XmlArray("featuresofspeech")>
    <Xml.Serialization.XmlArrayItem("feature")>
    Public FeaturesOfSpeech() As FeatureOfSpeechInfo
End Class
Public Class CachedData
    'need disk and memory cache as time consuming to read or build
    Private _PortableMethods As PortableMethods
    Private ArbData As ArabicData
    Private Arb As Arabic
    Public Sub New(NewPortableMethods As PortableMethods, NewArbData As ArabicData, NewArb As Arabic)
        _PortableMethods = NewPortableMethods
        ArbData = NewArbData
        Arb = NewArb
    End Sub
    Public Async Function Init(Optional bRec As Boolean = False, Optional bHadith As Boolean = False) As Threading.Tasks.Task
        If _ObjIslamData Is Nothing Then
            Dim fs As IO.Stream = Await _PortableMethods.FileIO.LoadStream(_PortableMethods.Settings.GetFilePath(_PortableMethods.FileIO.CombinePath("metadata", "islaminfo.xml")))
            Dim xs As Xml.Serialization.XmlSerializer = New Xml.Serialization.XmlSerializer(GetType(IslamData))
            _ObjIslamData = CType(xs.Deserialize(fs), IslamData)
            fs.Dispose()
        End If
        If _XMLDocMain Is Nothing Then
            Dim Stream As IO.Stream = Await _PortableMethods.FileIO.LoadStream(_PortableMethods.Settings.GetFilePath(_PortableMethods.FileIO.CombinePath("metadata", TanzilReader.QuranTextNames(0) + ".xml")))
            _XMLDocMain = Xml.Linq.XDocument.Load(Stream)
            Stream.Dispose()
        End If
        If _XMLDocInfo Is Nothing Then
            Dim Stream As IO.Stream = Await _PortableMethods.FileIO.LoadStream(_PortableMethods.Settings.GetFilePath(_PortableMethods.FileIO.CombinePath("metadata", "quran-data.xml")))
            _XMLDocInfo = Xml.Linq.XDocument.Load(Stream)
            Stream.Dispose()
        End If
        Dim Count As Integer
        If _XMLDocInfos Is Nothing And bHadith Then
            _XMLDocInfos = New Collections.Generic.List(Of Xml.Linq.XDocument)
            For Count = 0 To IslamData.Collections.Length - 1
                Dim Stream As IO.Stream = Await _PortableMethods.FileIO.LoadStream(_PortableMethods.Settings.GetFilePath(_PortableMethods.FileIO.CombinePath("metadata", IslamData.Collections(Count).FileName + "-data.xml")))
                _XMLDocInfos.Add(Xml.Linq.XDocument.Load(Stream))
                Stream.Dispose()
            Next
        End If
        If MorphLines Is Nothing Then MorphLines = Await _PortableMethods.ReadAllLines(_PortableMethods.Settings.GetFilePath(_PortableMethods.FileIO.CombinePath("metadata", "quranic-corpus-morphology-0.4.txt")))
        If _ArabicCamelCaseDict Is Nothing Then
            _ArabicCamelCaseDict = New Dictionary(Of String, Integer)
            For Count = 0 To ArbData.ArabicLetters.Length - 1
                Dim camelCase As String = ArbData.ToCamelCase(ArbData.ArabicLetters(Count).UnicodeName)
                If Not ArbData.ArabicLetters(Count).UnicodeName.StartsWith("<") And Not _ArabicCamelCaseDict.ContainsKey(camelCase) Then _ArabicCamelCaseDict.Add(camelCase, Count)
            Next
        End If
        If _ArabicComboCamelCaseDict Is Nothing Then
            _ArabicComboCamelCaseDict = New Dictionary(Of String, Integer)
            For Count = 0 To ArbData.ArabicCombos.Length - 1
                For SubCount = 0 To ArbData.ArabicCombos(Count).UnicodeName.Length - 1
                    If Not ArbData.ArabicCombos(Count).UnicodeName(SubCount) Is Nothing AndAlso ArbData.ArabicCombos(Count).UnicodeName(SubCount).Length <> 0 Then _ArabicComboCamelCaseDict.Add(ArbData.ToCamelCase(ArbData.ArabicCombos(Count).UnicodeName(SubCount)), Count)
                Next
            Next
        End If
        If _SavedPatterns.Count = 0 Then
            For Count = 0 To IslamData.ArabicPatterns.Length - 1
                _SavedPatterns.Add(IslamData.ArabicPatterns(Count).Name, Nothing)
            Next
        End If
        If _SavedGroups.Count = 0 Then
            For Count = 0 To IslamData.ArabicGroups.Length - 1
                _SavedGroups.Add(IslamData.ArabicGroups(Count).Name, Nothing)
            Next
        End If
        'only a few necessary for transliteration groupings many could be delayed
        If _ArabicUniqueLetters Is Nothing Then _ArabicUniqueLetters = GetNum("ArabicUniqueLetters")
        If _ArabicAlphabet Is Nothing Then _ArabicAlphabet = GetNum("ArabicAlphabet")
        If _ArabicNumbers Is Nothing Then _ArabicNumbers = GetNum("ArabicNumbers")
        If _ArabicWaslKasraExceptions Is Nothing Then _ArabicWaslKasraExceptions = GetNum("ArabicWaslKasraExceptions")
        If _ArabicBaseNumbers Is Nothing Then _ArabicBaseNumbers = GetNum("base")
        If _ArabicBaseExtraNumbers Is Nothing Then _ArabicBaseExtraNumbers = GetNum("baseextras")
        If _ArabicBaseTenNumbers Is Nothing Then _ArabicBaseTenNumbers = GetNum("baseten")
        If _ArabicBaseHundredNumbers Is Nothing Then _ArabicBaseHundredNumbers = GetNum("basehundred")
        If _ArabicBaseThousandNumbers Is Nothing Then _ArabicBaseThousandNumbers = GetNum("thousands")
        If _ArabicBaseMillionNumbers Is Nothing Then _ArabicBaseMillionNumbers = GetNum("millions")
        If _ArabicBaseMilliardNumbers Is Nothing Then _ArabicBaseMilliardNumbers = GetNum("milliard")
        If _ArabicBaseBillionNumbers Is Nothing Then _ArabicBaseBillionNumbers = GetNum("billions")
        If _ArabicBaseTrillionNumbers Is Nothing Then _ArabicBaseTrillionNumbers = GetNum("trillions")
        If _ArabicFractionNumbers Is Nothing Then _ArabicFractionNumbers = GetNum("fractions")
        If _ArabicOrdinalNumbers Is Nothing Then _ArabicOrdinalNumbers = GetNum("ordinals")
        If _ArabicOrdinalExtraNumbers Is Nothing Then _ArabicOrdinalExtraNumbers = GetNum("ordinalextras")
        If _ArabicCombiners Is Nothing Then _ArabicCombiners = GetNum("combiners")
        If _QuranHeaders Is Nothing Then _QuranHeaders = GetNum("quranheaders")
        If _ArabicLongVowels Is Nothing Then _ArabicLongVowels = GetGroup("ArabicLongVowels")
        If _ArabicTanweens Is Nothing Then _ArabicTanweens = GetGroup("ArabicTanweens")
        If _ArabicFathaDammaKasra Is Nothing Then _ArabicFathaDammaKasra = GetGroup("ArabicFathaDammaKasra")
        If _ArabicStopLetters Is Nothing Then _ArabicStopLetters = GetGroup("ArabicStopLetters")
        If _ArabicSpecialGutteral Is Nothing Then _ArabicSpecialGutteral = GetGroup("ArabicSpecialGutteral")
        If _ArabicSpecialLeadingGutteral Is Nothing Then _ArabicSpecialLeadingGutteral = GetGroup("ArabicSpecialLeadingGutteral")
        If _ArabicPunctuationSymbols Is Nothing Then _ArabicPunctuationSymbols = GetGroup("ArabicPunctuationSymbols")
        If _ArabicSunLetters Is Nothing Then _ArabicSunLetters = GetGroup("ArabicSunLetters")
        If _ArabicLetters Is Nothing Then _ArabicLetters = GetGroup("ArabicLetters")
        If _ArabicSunLettersNoLam Is Nothing Then _ArabicSunLettersNoLam = GetGroup("ArabicSunLettersNoLam")
        If _ArabicMoonLettersNoVowels Is Nothing Then _ArabicMoonLettersNoVowels = GetGroup("ArabicMoonLettersNoVowels")
        If _ArabicMoonLetters Is Nothing Then _ArabicMoonLetters = GetGroup("ArabicMoonLetters")
        If _RecitationCombiningSymbols Is Nothing Then _RecitationCombiningSymbols = GetGroup("RecitationCombiningSymbols")
        If _RecitationConnectingFollowerSymbols Is Nothing Then _RecitationConnectingFollowerSymbols = GetGroup("RecitationConnectingFollowerSymbols")
        If _RecitationSymbols Is Nothing Then _RecitationSymbols = GetGroup("RecitationSymbols")
        If _ArabicLettersInOrder Is Nothing Then _ArabicLettersInOrder = GetGroup("ArabicLettersInOrder")
        If _ArabicHamzas Is Nothing Then _ArabicHamzas = GetGroup("ArabicHamzas")
        If _ArabicVowels Is Nothing Then _ArabicVowels = GetGroup("ArabicVowels")
        If _ArabicMultis Is Nothing Then _ArabicMultis = GetGroup("ArabicMultis")
        If _ArabicTajweed Is Nothing Then _ArabicTajweed = GetGroup("ArabicTajweed")
        If _ArabicSilent Is Nothing Then _ArabicSilent = GetGroup("ArabicSilent")
        If _ArabicPunctuation Is Nothing Then _ArabicPunctuation = GetGroup("ArabicPunctuation")
        If _ArabicNums Is Nothing Then _ArabicNums = GetGroup("ArabicNums")
        If _NonArabicLetters Is Nothing Then _NonArabicLetters = GetGroup("NonArabicLetters")
        If _WhitespaceSymbols Is Nothing Then _WhitespaceSymbols = GetGroup("WhitespaceSymbols")
        If _PunctuationSymbols Is Nothing Then _PunctuationSymbols = GetGroup("PunctuationSymbols")
        If _RecitationDiacritics Is Nothing Then _RecitationDiacritics = GetGroup("RecitationDiacritics")
        If _RecitationLettersDiacritics Is Nothing Then _RecitationLettersDiacritics = GetGroup("RecitationLettersDiacritics")
        If _RecitationSpecialSymbols Is Nothing Then _RecitationSpecialSymbols = GetGroup("RecitationSpecialSymbols")
        If _ArabicLeadingGutterals Is Nothing Then _ArabicLeadingGutterals = GetGroup("ArabicLeadingGutterals")
        If _RecitationLetters Is Nothing Then _RecitationLetters = GetGroup("RecitationLetters")
        If _ArabicTrailingGutterals Is Nothing Then _ArabicTrailingGutterals = GetGroup("ArabicTrailingGutterals")
        If _RecitationSpecialSymbolsNotStop Is Nothing Then _RecitationSpecialSymbolsNotStop = GetGroup("RecitationSpecialSymbolsNotStop")
        If Not bRec And (_CertainStopPattern Is Nothing Or _RuleMetas.Count = 0 Or _RuleTranslations.Count = 0) Then Await Arb.Init(Me)
        For Count = 0 To IslamData.ArabicPatterns.Length - 1
            'Recursive and may already add pattern
            If _SavedPatterns(IslamData.ArabicPatterns(Count).Name) Is Nothing Then _SavedPatterns(IslamData.ArabicPatterns(Count).Name) = TranslateRegEx(IslamData.ArabicPatterns(Count).Match, True)
        Next
        For Count = 0 To IslamData.ArabicGroups.Length - 1
            If _SavedGroups(IslamData.ArabicGroups(Count).Name) Is Nothing Then _SavedGroups(IslamData.ArabicGroups(Count).Name) = New List(Of String)(Linq.Enumerable.Select(IslamData.ArabicGroups(Count).Text, Function(Str As String) TranslateRegEx(Str, Array.IndexOf(PatternAllowed, IslamData.ArabicGroups(Count).Name) <> -1 Or Array.IndexOf(Characteristics, IslamData.ArabicGroups(Count).Name) <> -1))).ToArray()
        Next
        'problematic to initialize before Arabic class
        If _CertainStopPattern Is Nothing Then _CertainStopPattern = GetPattern("CertainStopPattern")
        If _OptionalPattern Is Nothing Then _OptionalPattern = GetPattern("OptionalPattern")
        If _OptionalPatternNotEndOfAyah Is Nothing Then _OptionalPatternNotEndOfAyah = GetPattern("OptionalPatternNotEndOfAyah")
        If _CertainNotStopPattern Is Nothing Then _CertainNotStopPattern = GetPattern("CertainNotStopPattern")
        If _TehMarbutaStopRule Is Nothing Then _TehMarbutaStopRule = GetPattern("TehMarbutaStopRule")
        If _TehMarbutaContinueRule Is Nothing Then _TehMarbutaContinueRule = GetPattern("TehMarbutaContinueRule")
        If _ArabicSpecialLetters Is Nothing Then _ArabicSpecialLetters = GetGroup("ArabicSpecialLetters")
        If _RuleMetas.Count = 0 Then
            For Count = 0 To IslamData.MetaRules.Length - 1
                Dim BuildRules As New List(Of IslamData.RuleMetaSet.RuleMetadataTranslation)
                BuildRules.AddRange(IslamData.MetaRules(Count).Rules)
                For SubCount As Integer = 0 To BuildRules.Count - 1
                    BuildRules(SubCount) = New IslamData.RuleMetaSet.RuleMetadataTranslation With {.Match = TranslateRegEx(BuildRules(SubCount).Match, True), .Name = BuildRules(SubCount).Name, ._Evaluator = BuildRules(SubCount)._Evaluator, ._OptionalNotStopIndexes = BuildRules(SubCount)._OptionalNotStopIndexes, ._OptionalStopIndexes = BuildRules(SubCount)._OptionalStopIndexes, ._SplitEvaluator = BuildRules(SubCount)._SplitEvaluator}
                Next
                If IslamData.MetaRules(Count).From <> String.Empty Then BuildRules.AddRange(_RuleMetas(IslamData.MetaRules(Count).From).Rules)
                _RuleMetas.Add(IslamData.MetaRules(Count).Name, New IslamData.RuleMetaSet With {.Rules = BuildRules.ToArray(), .From = IslamData.MetaRules(Count).From, .Name = IslamData.MetaRules(Count).Name})
            Next
            For Count = 0 To IslamData.ColorRuleSets.Length - 1
                For SubCount As Integer = 0 To IslamData.ColorRuleSets(Count).ColorRules.Length - 1
                    If Not IslamData.ColorRuleSets(Count).ColorRules(SubCount).Evaluator Is Nothing Then IslamData.ColorRuleSets(Count).ColorRules(SubCount).Evaluator = TranslateRegEx(IslamData.ColorRuleSets(Count).ColorRules(SubCount).Evaluator, False)
                Next
            Next
        End If
        If _RuleTranslations.Count = 0 Then
            For Count = 0 To IslamData.RuleSets.Length - 1
                _RuleTranslations.Add(IslamData.RuleSets(Count).Name, GetRuleSet(IslamData.RuleSets(Count).Name))
            Next
        End If
    End Function
    Private _ObjIslamData As IslamData
    Private _RuleTranslations As New Dictionary(Of String, IslamData.RuleTranslationCategory.RuleTranslation())
    Private _RuleMetas As New Dictionary(Of String, IslamData.RuleMetaSet)
    Private _SavedPatterns As New Dictionary(Of String, String)
    Private _SavedGroups As New Dictionary(Of String, String())
    Public Function GetNum(Name As String) As String()
        Dim Count As Integer
        For Count = 0 To IslamData.ArabicNumbers.Length - 1
            If IslamData.ArabicNumbers(Count).Name = Name Then
                Return IslamData.ArabicNumbers(Count).Text
            End If
        Next
        Return {}
    End Function
    Public Function GetCap(Name As String) As String()
        Dim Count As Integer
        For Count = 0 To IslamData.ArabicCaptures.Length - 1
            If IslamData.ArabicCaptures(Count).Name = Name Then
                Return IslamData.ArabicCaptures(Count).Text
            End If
        Next
        Return {}
    End Function
    Dim PatLock As New Object
    Public Function GetPattern(Name As String) As String
        SyncLock PatLock
            If _SavedPatterns.ContainsKey(Name) Then
                If _SavedPatterns(Name) Is Nothing Then
                    For Count = 0 To IslamData.ArabicPatterns.Length - 1
                        'Recursive and may already add pattern
                        If Name = IslamData.ArabicPatterns(Count).Name Then
                            _SavedPatterns(Name) = TranslateRegEx(IslamData.ArabicPatterns(Count).Match, True)
                            Exit For
                        End If
                    Next
                End If
                Return _SavedPatterns(Name)
            End If
            Return String.Empty
        End SyncLock
    End Function
    Private Shared Characteristics As String() = {"Audibility", "Whispering", "Weakness", "Moderation", "Strength", "Lowness", "Elevation", "Opening", "Closing", "Restraint", "Fluency", "Vibration", "Inclination", "Repetition", "Whistling", "Diffusion", "Elongation", "Nasal", "Ease"}
    Private Shared PatternAllowed As String() = {"ArabicSpecialLetters", "ArabicAssimilateSameWord", "ArabicAssimilateAcrossWord", "ArabicAssimilateLeenAcrossWord", "ArabicBehSakinCombos", "ArabicTehSakinCombos", "ArabicThehSakinCombos", "ArabicJeemSakinCombos", "ArabicHahSakinCombos", "ArabicKhahSakinCombos", "ArabicDalSakinCombos", "ArabicThalSakinCombos", "ArabicRehSakinCombos", "ArabicZainSakinCombos", "ArabicSeenSakinCombos", "ArabicSheenSakinCombos", "ArabicSadSakinCombos", "ArabicDadSakinCombos", "ArabicTahSakinCombos", "ArabicZahSakinCombos", "ArabicAinSakinCombos", "ArabicGhainSakinCombos", "ArabicFehSakinCombos", "ArabicQafSakinCombos", "ArabicKafSakinCombos", "ArabicLamSakinCombos", "ArabicMeemSakinCombos", "ArabicNoonSakinCombos", "ArabicHehSakinCombos", "ArabicWawSakinCombos", "ArabicYehSakinCombos", "ArabicHamzaSakinCombos"}
    Dim GrpLock As New Object
    Public Function GetGroup(Name As String) As String()
        SyncLock GrpLock
            If _SavedGroups.ContainsKey(Name) Then
                If _SavedGroups(Name) Is Nothing Then
                    Dim Count As Integer
                    For Count = 0 To IslamData.ArabicGroups.Length - 1
                        If Name = IslamData.ArabicGroups(Count).Name Then
                            _SavedGroups(Name) = New List(Of String)(Linq.Enumerable.Select(IslamData.ArabicGroups(Count).Text, Function(Str As String) TranslateRegEx(Str, Array.IndexOf(PatternAllowed, IslamData.ArabicGroups(Count).Name) <> -1 Or Array.IndexOf(Characteristics, IslamData.ArabicGroups(Count).Name) <> -1))).ToArray()
                            Exit For
                        End If
                    Next
                End If
                Return _SavedGroups(Name)
            End If
            Return {}
        End SyncLock
    End Function
    Public Function GetRuleSet(Name As String) As IslamData.RuleTranslationCategory.RuleTranslation()
        Dim Count As Integer
        For Count = 0 To IslamData.RuleSets.Length - 1
            If IslamData.RuleSets(Count).Name = Name Then
                Dim RuleSet As IslamData.RuleTranslationCategory.RuleTranslation() = IslamData.RuleSets(Count).Rules
                For SubCount As Integer = 0 To RuleSet.Length - 1
                    RuleSet(SubCount).Match = TranslateRegEx(RuleSet(SubCount).Match, True)
                    RuleSet(SubCount).Evaluator = TranslateRegEx(RuleSet(SubCount).Evaluator, False)
                Next
                Return RuleSet
            End If
        Next
        Return Nothing
    End Function
    Private _ArabicUniqueLetters As String()
    Private _ArabicAlphabet As String()
    Private _ArabicNumbers As String()
    Private _ArabicWaslKasraExceptions As String()
    Private _ArabicBaseNumbers As String()
    Private _ArabicBaseExtraNumbers As String()
    Private _ArabicBaseTenNumbers As String()
    Private _ArabicBaseHundredNumbers As String()
    Private _ArabicBaseThousandNumbers As String()
    Private _ArabicBaseMillionNumbers As String()
    Private _ArabicBaseMilliardNumbers As String()
    Private _ArabicBaseBillionNumbers As String()
    Private _ArabicBaseTrillionNumbers As String()
    Private _ArabicFractionNumbers As String()
    Private _ArabicOrdinalNumbers As String()
    Private _ArabicOrdinalExtraNumbers As String()
    Private _ArabicCombiners As String()
    Private _QuranHeaders As String()
    Public ReadOnly Property ArabicUniqueLetters As String()
        Get
            Return _ArabicUniqueLetters
        End Get
    End Property
    Public ReadOnly Property ArabicAlphabet As String()
        Get
            Return _ArabicAlphabet
        End Get
    End Property
    Public ReadOnly Property ArabicNumbers As String()
        Get
            Return _ArabicNumbers
        End Get
    End Property
    Public ReadOnly Property ArabicWaslKasraExceptions As String()
        Get
            Return _ArabicWaslKasraExceptions
        End Get
    End Property
    Public ReadOnly Property ArabicBaseNumbers As String()
        Get
            Return _ArabicBaseNumbers
        End Get
    End Property
    Public ReadOnly Property ArabicBaseExtraNumbers As String()
        Get
            Return _ArabicBaseExtraNumbers
        End Get
    End Property
    Public ReadOnly Property ArabicBaseTenNumbers As String()
        Get
            Return _ArabicBaseTenNumbers
        End Get
    End Property
    Public ReadOnly Property ArabicBaseHundredNumbers As String()
        Get
            Return _ArabicBaseHundredNumbers
        End Get
    End Property
    Public ReadOnly Property ArabicBaseThousandNumbers As String()
        Get
            Return _ArabicBaseThousandNumbers
        End Get
    End Property
    Public ReadOnly Property ArabicBaseMillionNumbers As String()
        Get
            Return _ArabicBaseMillionNumbers
        End Get
    End Property
    Public ReadOnly Property ArabicBaseMilliardNumbers As String()
        Get
            Return _ArabicBaseMilliardNumbers
        End Get
    End Property
    Public ReadOnly Property ArabicBaseBillionNumbers As String()
        Get
            Return _ArabicBaseBillionNumbers
        End Get
    End Property
    Public ReadOnly Property ArabicBaseTrillionNumbers As String()
        Get
            Return _ArabicBaseTrillionNumbers
        End Get
    End Property
    Public ReadOnly Property ArabicFractionNumbers As String()
        Get
            Return _ArabicFractionNumbers
        End Get
    End Property
    Public ReadOnly Property ArabicOrdinalNumbers As String()
        Get
            Return _ArabicOrdinalNumbers
        End Get
    End Property
    Public ReadOnly Property ArabicOrdinalExtraNumbers As String()
        Get
            Return _ArabicOrdinalExtraNumbers
        End Get
    End Property
    Public ReadOnly Property ArabicCombiners As String()
        Get
            Return _ArabicCombiners
        End Get
    End Property
    Public ReadOnly Property QuranHeaders As String()
        Get
            Return _QuranHeaders
        End Get
    End Property
    Private _CertainStopPattern As String
    Private _OptionalPattern As String
    Private _OptionalPatternNotEndOfAyah As String
    Private _CertainNotStopPattern As String
    Private _TehMarbutaStopRule As String
    Private _TehMarbutaContinueRule As String
    Public ReadOnly Property CertainStopPattern As String
        Get
            Return _CertainStopPattern
        End Get
    End Property
    Public ReadOnly Property OptionalPattern As String
        Get
            Return _OptionalPattern
        End Get
    End Property
    Public ReadOnly Property OptionalPatternNotEndOfAyah As String
        Get
            Return _OptionalPatternNotEndOfAyah
        End Get
    End Property
    Public ReadOnly Property CertainNotStopPattern As String
        Get
            Return _CertainNotStopPattern
        End Get
    End Property
    Public ReadOnly Property TehMarbutaStopRule As String
        Get
            Return _TehMarbutaStopRule
        End Get
    End Property
    Public ReadOnly Property TehMarbutaContinueRule As String
        Get
            Return _TehMarbutaContinueRule
        End Get
    End Property
    Private _ArabicLongVowels As String()
    Private _ArabicTanweens As String()
    Private _ArabicFathaDammaKasra As String()
    Private _ArabicStopLetters As String()
    Private _ArabicSpecialGutteral As String()
    Private _ArabicSpecialLeadingGutteral As String()
    Private _ArabicPunctuationSymbols As String()
    Private _ArabicLetters As String()
    Private _ArabicSunLettersNoLam As String()
    Private _ArabicSunLetters As String()
    Private _ArabicMoonLettersNoVowels As String()
    Private _ArabicMoonLetters As String()
    Private _RecitationCombiningSymbols As String()
    Private _RecitationConnectingFollowerSymbols As String()
    Private _RecitationSymbols As String()
    Private _ArabicLettersInOrder As String()
    Private _ArabicSpecialLetters As String()
    Private _ArabicHamzas As String()
    Private _ArabicVowels As String()
    Private _ArabicMultis As String()
    Private _ArabicTajweed As String()
    Private _ArabicSilent As String()
    Private _ArabicPunctuation As String()
    Private _ArabicNums As String()
    Private _NonArabicLetters As String()
    Private _WhitespaceSymbols As String()
    Private _PunctuationSymbols As String()
    Private _RecitationDiacritics As String()
    Private _RecitationLettersDiacritics As String()
    Private _RecitationSpecialSymbols As String()
    Private _ArabicLeadingGutterals As String()
    Private _RecitationLetters As String()
    Private _ArabicTrailingGutterals As String()
    Private _RecitationSpecialSymbolsNotStop As String()
    Public ReadOnly Property ArabicLongVowels As String()
        Get
            Return _ArabicLongVowels
        End Get
    End Property
    Public ReadOnly Property ArabicTanweens As String()
        Get
            Return _ArabicTanweens
        End Get
    End Property
    Public ReadOnly Property ArabicFathaDammaKasra As String()
        Get
            Return _ArabicFathaDammaKasra
        End Get
    End Property
    Public ReadOnly Property ArabicStopLetters As String()
        Get
            Return _ArabicStopLetters
        End Get
    End Property
    Public ReadOnly Property ArabicSpecialGutteral As String()
        Get
            Return _ArabicSpecialGutteral
        End Get
    End Property
    Public ReadOnly Property ArabicSpecialLeadingGutteral As String()
        Get
            Return _ArabicSpecialLeadingGutteral
        End Get
    End Property
    Public ReadOnly Property ArabicPunctuationSymbols As String()
        Get
            Return _ArabicPunctuationSymbols
        End Get
    End Property
    Public Function GetLetterCharacteristics(Ch As String) As String()
        Dim ArticulationPoints As String() = {"NasalPassage", "TwoLips", "TwoLipsFromInside", "DeepestPartOfTongue", "UnderTheDeepestPartOfTongue", "EdgeOfTongue", "EdgeOfTongueLowestPart", "UnderTipOfTongue", "CloseToUnderTipOfTongueWithTop", "MiddleOfTongue", "TipOfTongueLetters", "ElevatedGumLetters", "GumLetters", "DeepestPartOfThroat", "MiddleOfThroat", "ClosestPartOfThroat", "ChestInterior"} '"TipLetters"
        Dim Characteristics As String() = {"Audibility", "Whispering", "Weakness", "Moderation", "Strength", "Lowness", "Elevation", "Opening", "Closing", "Restraint", "Fluency", "Vibration", "Inclination", "Repetition", "Whistling", "Diffusion", "Elongation", "Nasal", "Ease"}
        Dim Matches As New List(Of String)
        For Count = 0 To Characteristics.Length - 1
            If Array.IndexOf(GetGroup(Characteristics(Count)), Ch) <> -1 Then Matches.Add(Characteristics(Count))
        Next
        Return Matches.ToArray()
    End Function
    Public Function ArabicLetterCharacteristics() As String
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
                Dim Union As String() = New List(Of String)(Linq.Enumerable.Where(Chars, Function(Ch As String) Array.IndexOf(DupChars, Ch) <> -1)).ToArray()
                'Subtract the intersection to get differences
                Chars = New List(Of String)(Linq.Enumerable.Where(Chars, Function(Ch As String) Array.IndexOf(Union, Ch) = -1)).ToArray()
                DupChars = New List(Of String)(Linq.Enumerable.Where(DupChars, Function(Ch As String) Array.IndexOf(Union, Ch) = -1)).ToArray()
                'Mutually exclusive differences counted only once
                Dim MutExc As Integer = New List(Of String)(Linq.Enumerable.Where(Chars, Function(Ch As String) Array.IndexOf(MutualExclusiveChars, Ch) <> -1)).Count()
                LetCombs.Add(ArabicData.LeftToRightEmbedding + (Union.Length - Chars.Length - DupChars.Length + MutExc + 8).ToString("00") + "    " + ArabicData.PopDirectionalFormatting + DocBuilder.GetRegExText(Lets(Count)) + "+" + DocBuilder.GetRegExText(Lets(DupCount)) + ArabicData.LeftToRightEmbedding + "    " + CStr(Union.Length) + "    " + String.Join(", ", Chars) + " <> " + String.Join(", ", DupChars) + ArabicData.PopDirectionalFormatting)
            Next
        Next
        LetCombs.Sort(StringComparer.Ordinal)
        Return String.Join(vbCrLf, LetCombs.ToArray())
    End Function
    Public ReadOnly Property ArabicLetters As String()
        Get
            Return _ArabicLetters
        End Get
    End Property
    Public ReadOnly Property ArabicSunLettersNoLam As String()
        Get
            Return _ArabicSunLettersNoLam
        End Get
    End Property
    Public ReadOnly Property ArabicSunLetters As String()
        Get
            Return _ArabicSunLetters
        End Get
    End Property
    Public ReadOnly Property ArabicMoonLettersNoVowels As String()
        Get
            Return _ArabicMoonLettersNoVowels
        End Get
    End Property
    Public ReadOnly Property ArabicMoonLetters As String()
        Get
            Return _ArabicMoonLetters
        End Get
    End Property
    Public ReadOnly Property RecitationCombiningSymbols As String()
        Get
            Return _RecitationCombiningSymbols
        End Get
    End Property
    Public ReadOnly Property RecitationConnectingFollowerSymbols As String()
        Get
            Return _RecitationConnectingFollowerSymbols
        End Get
    End Property
    Public ReadOnly Property RecitationSymbols As String()
        Get
            Return _RecitationSymbols
        End Get
    End Property
    Public ReadOnly Property ArabicLettersInOrder As String()
        Get
            Return _ArabicLettersInOrder
        End Get
    End Property
    Public ReadOnly Property ArabicSpecialLetters As String()
        Get
            Return _ArabicSpecialLetters
        End Get
    End Property
    Public ReadOnly Property ArabicHamzas As String()
        Get
            Return _ArabicHamzas
        End Get
    End Property
    Public ReadOnly Property ArabicVowels As String()
        Get
            Return _ArabicVowels
        End Get
    End Property
    Public ReadOnly Property ArabicMultis As String()
        Get
            Return _ArabicMultis
        End Get
    End Property
    Public ReadOnly Property ArabicTajweed As String()
        Get
            Return _ArabicTajweed
        End Get
    End Property
    Public ReadOnly Property ArabicSilent As String()
        Get
            Return _ArabicSilent
        End Get
    End Property
    Public ReadOnly Property ArabicPunctuation As String()
        Get
            Return _ArabicPunctuation
        End Get
    End Property
    Public ReadOnly Property ArabicNums As String()
        Get
            Return _ArabicNums
        End Get
    End Property
    Public ReadOnly Property NonArabicLetters As String()
        Get
            Return _NonArabicLetters
        End Get
    End Property
    Public ReadOnly Property WhitespaceSymbols As String()
        Get
            Return _WhitespaceSymbols
        End Get
    End Property
    Public ReadOnly Property PunctuationSymbols As String()
        Get
            Return _PunctuationSymbols
        End Get
    End Property
    Public ReadOnly Property RecitationDiacritics As String()
        Get
            Return _RecitationDiacritics
        End Get
    End Property
    Public ReadOnly Property RecitationLettersDiacritics As String()
        Get
            Return _RecitationLettersDiacritics
        End Get
    End Property
    Public ReadOnly Property RecitationSpecialSymbols As String()
        Get
            Return _RecitationSpecialSymbols
        End Get
    End Property
    Public ReadOnly Property ArabicLeadingGutterals As String()
        Get
            Return _ArabicLeadingGutterals
        End Get
    End Property
    Public ReadOnly Property RecitationLetters As String()
        Get
            Return _RecitationLetters
        End Get
    End Property
    Public ReadOnly Property ArabicTrailingGutterals As String()
        Get
            Return _ArabicTrailingGutterals
        End Get
    End Property
    Public ReadOnly Property RecitationSpecialSymbolsNotStop As String()
        Get
            Return _RecitationSpecialSymbolsNotStop
        End Get
    End Property
    Private _ArabicCamelCaseDict As Dictionary(Of String, Integer)
    Private _ArabicComboCamelCaseDict As Dictionary(Of String, Integer)
    Public ReadOnly Property ArabicCamelCaseDict As Dictionary(Of String, Integer)
        Get
            Return _ArabicCamelCaseDict
        End Get
    End Property
    Public ReadOnly Property ArabicComboCamelCaseDict As Dictionary(Of String, Integer)
        Get
            Return _ArabicComboCamelCaseDict
        End Get
    End Property
    Public Function TranslateRegEx(Value As String, bAll As Boolean) As String
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
                                        Str += ArabicData.MakeUniRegEx(Arb.TransliterateFromBuckwalter(Strs(StrCount)))
                                    ElseIf StrCount + 1 <= CapLimit Then
                                        Str += "(" + ArabicData.MakeUniRegEx(Arb.TransliterateFromBuckwalter(Strs(StrCount))) + ")"
                                    ElseIf StrCount + 1 > CapLimit Then
                                        Str += "(?=" + ArabicData.MakeUniRegEx(Arb.TransliterateFromBuckwalter(Strs(StrCount))) + ")"
                                    End If
                                Next
                                Return Str
                            End Function)).ToArray())
                    End If
                    If GetNum(Match.Groups(1).Value).Length <> 0 Then
                        Return ArabicData.MakeRegMultiEx(New List(Of String)(Linq.Enumerable.Select(GetNum(Match.Groups(1).Value), Function(Str As String) ArabicData.MakeUniRegEx(Arb.TransliterateFromBuckwalter(Str)))).ToArray())
                    End If
                    If GetGroup(Match.Groups(1).Value).Length <> 0 Then
                        Return ArabicData.MakeRegMultiEx(New List(Of String)(Linq.Enumerable.Select(GetGroup(Match.Groups(1).Value), Function(Str As String) ArabicData.MakeUniRegEx(Str))).ToArray())
                    End If
                End If
                If System.Text.RegularExpressions.Regex.Match(Match.Groups(1).Value, "0x([0-9a-fA-F]{4})").Success Then
                    Return If(bAll, ArabicData.MakeUniRegEx(ChrW(Integer.Parse(Match.Groups(1).Value.Substring(2), System.Globalization.NumberStyles.HexNumber))), ChrW(Integer.Parse(Match.Groups(1).Value.Substring(2), System.Globalization.NumberStyles.HexNumber)))
                End If
                If ArabicCamelCaseDict.ContainsKey(Match.Groups(1).Value) Then
                    Return If(bAll, ArabicData.MakeUniRegEx(ArbData.ArabicLetters(ArabicCamelCaseDict(Match.Groups(1).Value)).Symbol), ArbData.ArabicLetters(ArabicCamelCaseDict(Match.Groups(1).Value)).Symbol)
                End If
                If ArabicComboCamelCaseDict.ContainsKey(Match.Groups(1).Value) Then
                    Return If(bAll, ArabicData.MakeUniRegEx(If(ArbData.ArabicCombos(ArabicComboCamelCaseDict(Match.Groups(1).Value)).Shaping.Length = 1, ArbData.ArabicCombos(ArabicComboCamelCaseDict(Match.Groups(1).Value)).Shaping(0), String.Join(String.Empty, Linq.Enumerable.Select(ArbData.ArabicCombos(ArabicComboCamelCaseDict(Match.Groups(1).Value)).Symbol, Function(Sym As Char) CStr(Sym))))), If(ArbData.ArabicCombos(ArabicComboCamelCaseDict(Match.Groups(1).Value)).Shaping.Length = 1, ArbData.ArabicCombos(ArabicComboCamelCaseDict(Match.Groups(1).Value)).Shaping(0), String.Join(String.Empty, Linq.Enumerable.Select(ArbData.ArabicCombos(ArabicComboCamelCaseDict(Match.Groups(1).Value)).Symbol, Function(Sym As Char) CStr(Sym)))))
                End If
                '{0} ignore
                'If Not IsNumeric(Match.Groups(1).Value) Then Debug.Print("Unknown Group: " + Match.Groups(1).Value)
                Return Match.Value
            End Function)
    End Function
    Public ReadOnly Property RuleMetas(Name As String) As IslamData.RuleMetaSet
        Get
            Return _RuleMetas(Name)
        End Get
    End Property
    Public ReadOnly Property RuleTranslations(Name As String) As IslamData.RuleTranslationCategory.RuleTranslation()
        Get
            Return _RuleTranslations(Name)
        End Get
    End Property

    Private _XMLDocMain As Xml.Linq.XDocument 'Tanzil Quran data
    Private _XMLDocInfo As Xml.Linq.XDocument 'Tanzil metadata
    Private _XMLDocInfos As Collections.Generic.List(Of Xml.Linq.XDocument) 'Hadiths
    Private _RootDictionary As New Generic.Dictionary(Of String, List(Of Integer()))
    Private _FormDictionary As New Generic.Dictionary(Of String, List(Of Integer()))
    Private _TagDictionary As New Generic.Dictionary(Of String, Generic.Dictionary(Of String, List(Of Integer())))
    Private _WordDictionary As New Generic.Dictionary(Of String, List(Of String))
    Private _RealWordDictionary As New Generic.Dictionary(Of String, List(Of Integer()))
    Private _LetterDictionary As New Generic.Dictionary(Of Char, Generic.Dictionary(Of String, List(Of Integer())))
    Private _LetterPreDictionary As New Generic.Dictionary(Of Char, Generic.Dictionary(Of String, List(Of Integer())))
    Private _LetterSufDictionary As New Generic.Dictionary(Of Char, Generic.Dictionary(Of String, List(Of Integer())))
    Private _PreDictionary As New Generic.Dictionary(Of String, List(Of Integer()))
    Private _SufDictionary As New Generic.Dictionary(Of String, List(Of Integer()))
    Private _LocDictionary As New Generic.Dictionary(Of String, Object())
    Private _IsolatedLetterDictionary As New Generic.Dictionary(Of Char, List(Of Integer()))
    Private _TotalLetters As Integer = 0
    Private _TotalIsolatedLetters As Integer = 0
    Private _PartUniqueArray() As Generic.List(Of String)
    Private _PartArray() As Generic.List(Of String)
    Private _StationUniqueArray() As Generic.List(Of String)
    Private _StationArray() As Generic.List(Of String)
    Private _TotalUniqueWordsInParts As Integer = 0
    Private _TotalWordsInParts As Integer = 0
    Private _TotalUniqueWordsInStations As Integer = 0
    Private _TotalWordsInStations As Integer = 0
    Private _MorphDataToLineNumber As Dictionary(Of Integer(), Integer)
    Public Function GetPartOfSpeech(POS As String) As IslamData.PartOfSpeechInfo
        For Count As Integer = 0 To IslamData.PartsOfSpeech.Length - 1
            If IslamData.PartsOfSpeech(Count).Symbol = POS Then Return IslamData.PartsOfSpeech(Count)
        Next
        Return Nothing
    End Function
    Public Function GetFeatureOfSpeech(FOS As String) As Nullable(Of IslamData.FeatureOfSpeechInfo)
        For Count As Integer = 0 To IslamData.FeaturesOfSpeech.Length - 1
            If IslamData.FeaturesOfSpeech(Count).Symbol = FOS Then Return IslamData.FeaturesOfSpeech(Count)
        Next
        Return Nothing
    End Function
    Public Class ByteArrayComparer
        Inherits EqualityComparer(Of Integer())
        Public Overrides Function Equals(x() As Integer, y() As Integer) As Boolean
            If x Is Nothing Or y Is Nothing Then Return (x Is Nothing) = (y Is Nothing)
            If ReferenceEquals(x, y) Then Return True
            If x.Length <> y.Length Then Return False
            Return Linq.Enumerable.SequenceEqual(x, y)
        End Function
        Public Overrides Function GetHashCode(obj() As Integer) As Integer
            'Maximum chapter Is 7 bits, verse Is 9 bits, word Is 8 bits and sub-word is 2 bits
            Return obj(0) + (obj(1) << 7) + (obj(2) << 16) + (obj(3) << 24)
        End Function
    End Class
    Public MorphLines As String()
    Public Function GetWritingHamza(StrBefore As String, StrAfter As String, bUthmani As Boolean) As String
        Dim Rules As IslamData.RuleTranslationCategory.RuleTranslation() = GetRuleSet(If(bUthmani, "HamzaWriting", "SimpleScriptHamzaWriting"))
        Dim IndexToVerse As Integer()() = Nothing
        Dim Text As String = StrBefore + ArabicData.ArabicLetterHamza + StrAfter
        Dim Matches As System.Text.RegularExpressions.MatchCollection = System.Text.RegularExpressions.Regex.Matches(Text, GetPattern("Hamzas"))
        For MainCount = 0 To Rules.Length - 1
            Matches = System.Text.RegularExpressions.Regex.Matches(Text, Rules(MainCount).Match)
            Dim NegativeCount As Integer = 0
            For Count = 0 To Matches.Count - 1
                If Rules(MainCount).NegativeMatch <> String.Empty AndAlso Matches(Count).Result(Rules(MainCount).NegativeMatch) <> String.Empty Then
                ElseIf Matches(Count).Groups(2 + If(Rules(MainCount).NegativeMatch <> String.Empty, 1, 0)).Index = StrBefore.Length Then
                    Return Matches(Count).Result(Rules(MainCount).Evaluator).Substring(Matches(Count).Groups(0).Length)
                End If
            Next
        Next
        Return String.Empty
    End Function
    Public Function CombineFixHamzas(Str As String) As String
        Return System.Text.RegularExpressions.Regex.Replace(Str, GetPattern("Hamzas"), Function(Match As System.Text.RegularExpressions.Match) GetWritingHamza(Str.Substring(0, Match.Index), Str.Substring(Match.Index + Match.Length), True))
    End Function
    Structure ScaleRef
        Dim Scale As String
        Dim LongShortBoth As Char
        Dim Refs As List(Of Integer())
    End Structure
    Public Function GetMorphologicalDataByVerbScale(TR As TanzilReader) As Object()
        Dim Lines As String() = MorphLines
        Dim RootDictionary As New Dictionary(Of String, ScaleRef())
        For Count = 0 To Lines.Length - 1
            If Lines(Count).Length <> 0 AndAlso Lines(Count).Chars(0) <> "#" Then
                Dim Pieces As String() = Lines(Count).Split(CChar(vbTab))
                If Pieces(0).Chars(0) = "(" AndAlso Pieces(2) = "V" Then
                    Dim Location As Integer() = New List(Of Integer)(Linq.Enumerable.Select(Pieces(0).TrimStart("("c).TrimEnd(")"c).Split(":"c), Function(Str As String) CInt(Str))).ToArray()
                    Dim Parts As String() = Pieces(3).Split("|"c)
                    Dim bNotDefaultPresent As Boolean = False
                    Dim bPassive As Boolean = False
                    Dim Type As Integer = 1
                    Dim Tense As Integer = 0
                    Dim Root As String = String.Empty
                    Dim Lemma As String = String.Empty
                    For SubCount As Integer = 0 To Parts.Length - 1
                        Dim Types As String() = Parts(SubCount).Split(":"c)
                        If Types(0) = "ROOT" Then
                            Root = Types(1)
                        ElseIf Types(0) = "LEM" Then
                            Lemma = Types(1)
                        ElseIf Types(0) = "MOOD" Then
                            bNotDefaultPresent = True
                        ElseIf Types(0) = "(II)" Then
                            Type = 2
                        ElseIf Types(0) = "(III)" Then
                            Type = 3
                        ElseIf Types(0) = "(IV)" Then
                            Type = 4
                        ElseIf Types(0) = "(V)" Then
                            Type = 5
                        ElseIf Types(0) = "(VI)" Then
                            Type = 6
                        ElseIf Types(0) = "(VII)" Then
                            Type = 7
                        ElseIf Types(0) = "(VIII)" Then
                            Type = 8
                        ElseIf Types(0) = "(IX)" Then
                            Type = 9
                        ElseIf Types(0) = "(X)" Then
                            Type = 10
                        ElseIf Types(0) = "(XI)" Then
                            Type = 11
                        ElseIf Types(0) = "(XII)" Then
                            Type = 12
                        ElseIf Types(0) = "PASS" Then
                            bPassive = True
                        ElseIf Types(0) = "PERF" Then
                            Tense = 1
                        ElseIf Types(0) = "IMPF" Then
                            Tense = 2
                        ElseIf Types(0) = "IMPV" Then
                            Tense = 3
                        End If
                    Next
                    If Type = 1 Then
                        If Root.Length = 3 And Not RootDictionary.ContainsKey(Root) Then RootDictionary.Add(Root, New ScaleRef() {New ScaleRef With {.Refs = New List(Of Integer())}, New ScaleRef With {.Refs = New List(Of Integer())}, New ScaleRef With {.Refs = New List(Of Integer())}})
                        If Root.Length = 3 And (Tense = 1 And Not bPassive Or Tense = 2 And Root(0) = Lemma(0)) Then
                            Dim Orig As String = Pieces(1)
                            If Tense = 2 Then Pieces(1) = Lemma
                            If Root(2) = "w" Then
                                Dim Match As Text.RegularExpressions.Match = System.Text.RegularExpressions.Regex.Match(Pieces(1), (If(Root(0) = "A", ">", Root(0) + "~?") + "a" + If(Root(1) = "A", "(?:>|'|_#|&)", Root(1)) + "(.)(?:A|w|Y`?|y|t|$$)").Replace("$", "\$").Replace("*", "\*").Replace("\$\$", "$"))
                                If Match.Success And Match.Captures.Count = 1 And RootDictionary(Root)(0).Scale = String.Empty Then
                                    RootDictionary(Root)(0).Scale = If(Match.Groups(1).Value = "u", "i", Match.Groups(1).Value)
                                    'RootDictionary(Root)(0) = If(Root(0) = "A", ">", Root(0)) + "a" + If(Root(1) = "A", "'", Root(1)) + "aA"
                                ElseIf Not Match.Success Or Not Match.Captures.Count = 1 Or RootDictionary(Root)(0).Scale <> If(Match.Groups(1).Value = "u", "i", Match.Groups(1).Value) Then
                                    Debug.WriteLine(Root + " - " + Pieces(1) + " - " + RootDictionary(Root)(0).Scale)
                                End If
                                If Match.Success And Match.Captures.Count = 1 And Tense <> 2 Then RootDictionary(Root)(0).Refs.Add(Location)
                            ElseIf Root(2) = "y" And Root(1) = "A" And Root(0) = "r" Then
                                'rAy is a special unique exception
                                Dim Match As Text.RegularExpressions.Match = System.Text.RegularExpressions.Regex.Match(Pieces(1), (If(Root(0) = "A", ">", Root(0) + "~?") + "(a)" + If(Root(1) = "A", "(?:>|'|_#|&)", Root(1)) + "(?:" + Root(2) + "|A|t)?").Replace("$", "\$").Replace("*", "\*"))
                                If Match.Success And Match.Captures.Count = 1 And RootDictionary(Root)(0).Scale = String.Empty Then
                                    RootDictionary(Root)(0).Scale = Match.Groups(1).Value
                                    'RootDictionary(Root)(0) = "ya" + Root(0) + Match.Groups(1).Value + Root(2)
                                ElseIf Not Match.Success Or Not Match.Captures.Count = 1 Or RootDictionary(Root)(0).Scale <> Match.Groups(1).Value Then
                                    Debug.WriteLine(Root + " - " + Pieces(1) + " - " + RootDictionary(Root)(0).Scale)
                                End If
                                If Match.Success And Match.Captures.Count = 1 And Tense <> 2 Then RootDictionary(Root)(0).Refs.Add(Location)
                            ElseIf Root(2) = "y" And Root(1) <> "y" Then
                                Dim Match As Text.RegularExpressions.Match = System.Text.RegularExpressions.Regex.Match(Pieces(1), (If(Root(0) = "A", ">", Root(0) + "~?") + "a" + If(Root(1) = "A", "(?:>|'|_#|&)", Root(1)) + "(.)(?:y|Y|t|A|$$)").Replace("$", "\$").Replace("*", "\*").Replace("\$\$", "$"))
                                If Match.Success And Match.Captures.Count = 1 And RootDictionary(Root)(0).Scale = String.Empty Then
                                    RootDictionary(Root)(0).Scale = If(Match.Groups(1).Value = "u", "i", Match.Groups(1).Value)
                                    'RootDictionary(Root)(0) = If(Root(0) = "A", ">", Root(0)) + "a" + If(Root(1) = "A", "'", Root(1)) + "aY"
                                ElseIf Not Match.Success Or Not Match.Captures.Count = 1 Or RootDictionary(Root)(0).Scale <> If(Match.Groups(1).Value = "u", "i", Match.Groups(1).Value) Then
                                    Debug.WriteLine(Root + " - " + Pieces(1) + " - " + RootDictionary(Root)(0).Scale)
                                End If
                                If Match.Success And Match.Captures.Count = 1 And Tense <> 2 Then RootDictionary(Root)(0).Refs.Add(Location)
                            ElseIf Root(1) = Root(2) And System.Text.RegularExpressions.Regex.Match(Pieces(1), (If(Root(0) = "A", ">", Root(0) + "~?") + "a" + If(Root(1) = "A", "(?:>|'|_#|&)", If(Root(1) = "y", "(?:y|Y)", Root(1))) + "(?:~|o$$)").Replace("$", "\$").Replace("*", "\*").Replace("\$\$", "$")).Success Then
                                'RootDictionary(Root)(0) = If(Root(0) = "A", ">", Root(0)) + "a" + If(Root(1) = "A", "'", Root(1)) + "~a"
                                If Tense <> 2 Then RootDictionary(Root)(0).Refs.Add(Location)
                            ElseIf Root(1) = "w" Or Root(1) = "y" And System.Text.RegularExpressions.Regex.Match(Pieces(1), (If(Root(0) = "A", ">", Root(0) + "~?") + If(Root(1) = "w", "u", "i") + If(Root(2) = "A", "(?:>|'|_#|&|\})", Root(2))).Replace("$", "\$").Replace("*", "\*")).Success Then
                                If RootDictionary(Root)(0).Scale = String.Empty Then
                                    RootDictionary(Root)(0).Scale = "a"
                                    'RootDictionary(Root)(0) = If(Root(0) = "A", ">", Root(0)) + "aA" + If(Root(2) = "A", "'", Root(2)) + "a"
                                ElseIf RootDictionary(Root)(0).Scale <> "a" Then
                                    Debug.WriteLine(Root + " - " + Pieces(1) + " - " + RootDictionary(Root)(0).Scale)
                                End If
                                If Tense <> 2 Then RootDictionary(Root)(0).Refs.Add(Location)
                            ElseIf Root(1) = "w" Or Root(1) = "y" Then
                                Dim Match As Text.RegularExpressions.Match = System.Text.RegularExpressions.Regex.Match(Pieces(1), (If(Root(0) = "A", ">", Root(0) + "~?") + "(a)(?:A|(?:" + Root(1) + "o)?)" + If(Root(2) = "A", "\^(?:>|'|_#|&|\})", Root(2))).Replace("$", "\$").Replace("*", "\*"))
                                If Match.Success And Match.Captures.Count = 1 And RootDictionary(Root)(0).Scale = String.Empty Then
                                    RootDictionary(Root)(0).Scale = Match.Groups(1).Value
                                    'RootDictionary(Root)(0) = If(Root(0) = "A", ">", Root(0)) + "aA" + If(Root(2) = "A", "'", Root(2)) + "a"
                                ElseIf Not Match.Success Or Not Match.Captures.Count = 1 And Tense <> 2 Or RootDictionary(Root)(0).Scale <> Match.Groups(1).Value Then
                                    Debug.WriteLine(Root + " - " + Pieces(1) + " - " + RootDictionary(Root)(0).Scale)
                                End If
                                If Match.Success And Match.Captures.Count = 1 Then RootDictionary(Root)(0).Refs.Add(Location)
                            ElseIf System.Text.RegularExpressions.Regex.Match(Pieces(1), (If(Root(0) = "A", ">", Root(0) + "~?") + "i" + If(Root(1) = "A", "(?:>|_#|\})", Root(1)) + "o" + If(Root(2) = "A", "(?:>|\^?'|_#|&)", Root(2) + If(Root(2) = "n", "?", String.Empty))).Replace("$", "\$").Replace("*", "\*")).Success Then
                                Debug.WriteLine("Special past only: " + Root + " - " + Pieces(1))
                            Else
                                Dim Match As Text.RegularExpressions.Match = System.Text.RegularExpressions.Regex.Match(Pieces(1), (If(Root(0) = "A", ">", Root(0) + "~?") + "a`?" + If(Root(1) = "A", "(?:>|_#|\})", Root(1)) + "(.)" + If(Root(2) = "A", "(?:>|\^?'|_#|&)", Root(2) + If(Root(2) = "n" Or Root(2) = "t", "?", String.Empty))).Replace("$", "\$").Replace("*", "\*"))
                                If Match.Success And Match.Captures.Count = 1 And RootDictionary(Root)(0).Scale = String.Empty Then
                                    RootDictionary(Root)(0).Scale = Match.Groups(1).Value
                                    'RootDictionary(Root)(0) = If(Root(0) = "A", ">", Root(0)) + "a" + Root(1) + Match.Groups(1).Value + If(Root(2) = "A", If(Match.Groups(1).Value = "u", "&", If(Match.Groups(1).Value = "a", ">", "{")), Root(2)) + "a"
                                ElseIf Not Match.Success Or Not Match.Captures.Count = 1 Or RootDictionary(Root)(0).Scale <> Match.Groups(1).Value Then
                                    Debug.WriteLine(Root + " - " + Pieces(1) + " - " + RootDictionary(Root)(0).Scale)
                                End If
                                If Match.Success And Match.Captures.Count = 1 And Tense <> 2 Then RootDictionary(Root)(0).Refs.Add(Location)
                            End If
                            If Tense = 2 Then Pieces(1) = Orig
                        End If
                        If Tense = 2 Then
                            If Root.Length = 3 And Not bPassive Then
                                If Root(2) = "w" Then
                                    '-oona merges into it
                                    Dim Match As Text.RegularExpressions.Match = System.Text.RegularExpressions.Regex.Match(Pieces(1), ("a" + If(Root(0) = "A", "(?:A|>)", Root(0)) + "o?" + Root(1) + "(.)(" + Root(2) + "|Y`)?").Replace("$", "\$").Replace("*", "\*"))
                                    If Match.Success And Match.Captures.Count = 1 And (RootDictionary(Root)(1).Scale = String.Empty Or RootDictionary(Root)(1).Scale <> Match.Groups(1).Value And RootDictionary(Root)(2).Scale = String.Empty) Then
                                        RootDictionary(Root)(If(RootDictionary(Root)(1).Scale = String.Empty, 1, 2)).Scale = Match.Groups(1).Value
                                        'RootDictionary(Root)(1) = "ya" + If(Root(0) = "A", ">", Root(0)) + "o" + Root(1) + Match.Groups(1).Value + Root(2)
                                    ElseIf Not Match.Success Or Not Match.Captures.Count = 1 Or (RootDictionary(Root)(1).Scale <> Match.Groups(1).Value And RootDictionary(Root)(2).Scale <> Match.Groups(1).Value) Then
                                        Debug.WriteLine(Root + " - " + Pieces(1) + " - " + RootDictionary(Root)(1).Scale + RootDictionary(Root)(2).Scale)
                                    End If
                                    If Match.Success And Match.Captures.Count = 1 Then RootDictionary(Root)(If(RootDictionary(Root)(1).Scale = Match.Groups(1).Value, 1, 2)).Refs.Add(Location)
                                ElseIf Root(2) = "y" And Root(1) = "A" And Root(0) = "r" Then
                                    'rAy is a special unique exception
                                    Dim Match As Text.RegularExpressions.Match = System.Text.RegularExpressions.Regex.Match(Pieces(1), ("a" + If(Root(0) = "A", ">", Root(0)) + "(.)" + "(?:" + Root(2) + "|Y)?").Replace("$", "\$").Replace("*", "\*"))
                                    If Match.Success And Match.Captures.Count = 1 And (RootDictionary(Root)(1).Scale = String.Empty Or RootDictionary(Root)(1).Scale <> Match.Groups(1).Value And RootDictionary(Root)(2).Scale = String.Empty) Then
                                        RootDictionary(Root)(If(RootDictionary(Root)(1).Scale = String.Empty, 1, 2)).Scale = Match.Groups(1).Value
                                        'RootDictionary(Root)(1) = "ya" + Root(0) + Match.Groups(1).Value + Root(2)
                                    ElseIf Not Match.Success Or Not Match.Captures.Count = 1 Or (RootDictionary(Root)(1).Scale <> Match.Groups(1).Value And RootDictionary(Root)(2).Scale <> Match.Groups(1).Value) Then
                                        Debug.WriteLine(Root + " - " + Pieces(1) + " - " + RootDictionary(Root)(1).Scale + RootDictionary(Root)(2).Scale)
                                    End If
                                    If Match.Success And Match.Captures.Count = 1 Then RootDictionary(Root)(If(RootDictionary(Root)(1).Scale = Match.Groups(1).Value, 1, 2)).Refs.Add(Location)
                                ElseIf Root(0) = "w" And Root(2) = "y" Then
                                    Dim Match As Text.RegularExpressions.Match = System.Text.RegularExpressions.Regex.Match(Pieces(1), ("a" + Root(1) + "(.)" + "(?:" + Root(2) + "|Y)?").Replace("$", "\$").Replace("*", "\*"))
                                    If Match.Success And Match.Captures.Count = 1 And (RootDictionary(Root)(1).Scale = String.Empty Or RootDictionary(Root)(1).Scale <> Match.Groups(1).Value And RootDictionary(Root)(2).Scale = String.Empty) Then
                                        RootDictionary(Root)(If(RootDictionary(Root)(1).Scale = String.Empty, 1, 2)).Scale = Match.Groups(1).Value
                                        'RootDictionary(Root)(1) = "ya" + Root(1) + Match.Groups(1).Value + Root(2)
                                    ElseIf Not Match.Success Or Not Match.Captures.Count = 1 Or (RootDictionary(Root)(1).Scale <> Match.Groups(1).Value And RootDictionary(Root)(2).Scale <> Match.Groups(1).Value) Then
                                        Debug.WriteLine(Root + " - " + Pieces(1) + " - " + RootDictionary(Root)(1).Scale + RootDictionary(Root)(2).Scale)
                                    End If
                                    If Match.Success And Match.Captures.Count = 1 Then RootDictionary(Root)(If(RootDictionary(Root)(1).Scale = Match.Groups(1).Value, 1, 2)).Refs.Add(Location)
                                ElseIf Root(2) = "y" Then
                                    '-iyoona changes to -oona and -aaoona changes to -awna fatha
                                    '"AhHxgE" of ayn makes ayn on a, others on i
                                    Dim Match As Text.RegularExpressions.Match = System.Text.RegularExpressions.Regex.Match(Pieces(1), ("a" + If(Root(0) = "A", "(?:A|>)", Root(0)) + "o?" + If(Root(1) = "y", "(?:y|Y)", If(Root(1) = "A", "_#", Root(1) + If(Root(1) = "w", "?", String.Empty))) + "(.)" + If(Root(1) = "w", "?", String.Empty) + "(?:" + Root(2) + "|Y)?").Replace("$", "\$").Replace("*", "\*"))
                                    If Match.Success And Match.Captures.Count = 1 And (RootDictionary(Root)(1).Scale = String.Empty Or RootDictionary(Root)(1).Scale <> If(Match.Groups(1).Value = "u", "i", Match.Groups(1).Value) And RootDictionary(Root)(2).Scale = String.Empty) Then
                                        RootDictionary(Root)(If(RootDictionary(Root)(1).Scale = String.Empty, 1, 2)).Scale = If(Match.Groups(1).Value = "u", "i", Match.Groups(1).Value)
                                        'RootDictionary(Root)(1) = "ya" + If(Root(0) = "A", ">", Root(0)) + "o" + If(Root(1) = "A", "_#", Root(1)) + Match.Groups(1).Value + Root(2)
                                    ElseIf Not Match.Success Or Not Match.Captures.Count = 1 Or (RootDictionary(Root)(1).Scale <> If(Match.Groups(1).Value = "u", "i", Match.Groups(1).Value) And RootDictionary(Root)(2).Scale <> If(Match.Groups(1).Value = "u", "i", Match.Groups(1).Value)) Then
                                        Debug.WriteLine(Root + " - " + Pieces(1) + " - " + RootDictionary(Root)(1).Scale + RootDictionary(Root)(2).Scale)
                                    End If
                                    If Match.Success And Match.Captures.Count = 1 Then RootDictionary(Root)(If(RootDictionary(Root)(1).Scale = Match.Groups(1).Value, 1, 2)).Refs.Add(Location)
                                ElseIf Root(1) = Root(2) And System.Text.RegularExpressions.Regex.Match(Pieces(1), ("a" + If(Root(0) = "A", "&", Root(0)) + "(.)" + Root(1) + "~").Replace("$", "\$").Replace("*", "\*")).Success Then
                                    Dim Match As Text.RegularExpressions.Match = System.Text.RegularExpressions.Regex.Match(Pieces(1), ("a" + If(Root(0) = "A", "&", Root(0)) + "(.)" + Root(1) + "~").Replace("$", "\$").Replace("*", "\*"))
                                    If Match.Success And Match.Captures.Count = 1 And (RootDictionary(Root)(1).Scale = String.Empty Or RootDictionary(Root)(1).Scale <> Match.Groups(1).Value And RootDictionary(Root)(2).Scale = String.Empty) Then
                                        RootDictionary(Root)(If(RootDictionary(Root)(1).Scale = String.Empty, 1, 2)).Scale = Match.Groups(1).Value
                                        'RootDictionary(Root)(1) = "ya" + If(Root(0) = "A", "&", Root(0)) + Match.Groups(1).Value + Root(1) + "~u"
                                    ElseIf Not Match.Success Or Not Match.Captures.Count = 1 Or (RootDictionary(Root)(1).Scale <> Match.Groups(1).Value And RootDictionary(Root)(2).Scale <> Match.Groups(1).Value) Then
                                        Debug.WriteLine(Root + " - " + Pieces(1) + " - " + RootDictionary(Root)(1).Scale + RootDictionary(Root)(2).Scale)
                                    End If
                                    If Match.Success And Match.Captures.Count = 1 Then RootDictionary(Root)(If(RootDictionary(Root)(1).Scale = Match.Groups(1).Value, 1, 2)).Refs.Add(Location)
                                ElseIf Root(0) = "w" And System.Text.RegularExpressions.Regex.Match(Pieces(1), ("a" + Root(1) + "(.)" + If(Root(2) = "A", "_#", Root(2))).Replace("$", "\$").Replace("*", "\*")).Success Then
                                    '"AhHxgE" of lam makes ayn on a, others on i
                                    Dim Match As Text.RegularExpressions.Match = System.Text.RegularExpressions.Regex.Match(Pieces(1), ("a" + Root(1) + "(.)" + If(Root(2) = "A", "_#", Root(2))).Replace("$", "\$").Replace("*", "\*"))
                                    If Match.Success And Match.Captures.Count = 1 And (RootDictionary(Root)(1).Scale = String.Empty Or RootDictionary(Root)(1).Scale <> Match.Groups(1).Value And RootDictionary(Root)(2).Scale = String.Empty) Then
                                        RootDictionary(Root)(If(RootDictionary(Root)(1).Scale = String.Empty, 1, 2)).LongShortBoth = "S"c
                                        RootDictionary(Root)(If(RootDictionary(Root)(1).Scale = String.Empty, 1, 2)).Scale = Match.Groups(1).Value
                                        'RootDictionary(Root)(1) = "ya" + Root(1) + Match.Groups(1).Value + If(Root(2) = "A", ">", Root(2)) + "u"
                                    ElseIf Not Match.Success Or Not Match.Captures.Count = 1 Or (RootDictionary(Root)(1).Scale <> Match.Groups(1).Value And RootDictionary(Root)(2).Scale <> Match.Groups(1).Value) Then
                                        Debug.WriteLine(Root + " - " + Pieces(1) + " - " + RootDictionary(Root)(1).Scale + RootDictionary(Root)(2).Scale)
                                    End If
                                    If Match.Success And Match.Captures.Count = 1 Then RootDictionary(Root)(If(RootDictionary(Root)(1).Scale = Match.Groups(1).Value, 1, 2)).Refs.Add(Location)
                                ElseIf Root(1) = "w" Then
                                    'kwd, xwf exceptions with ayn on a
                                    Dim Match As Text.RegularExpressions.Match = System.Text.RegularExpressions.Regex.Match(Pieces(1), ("a" + If(Root(0) = "A", "_#", Root(0)) + "(.)" + "(?:" + Root(1) + "|,|A|`)?" + If(Root(2) = "A", "\^>|\^_#|&", Root(2) + If(Root(2) = "n", "?", String.Empty))).Replace("$", "\$").Replace("*", "\*"))
                                    If Match.Success And Match.Captures.Count = 1 And (RootDictionary(Root)(1).Scale = String.Empty Or RootDictionary(Root)(1).Scale <> Match.Groups(1).Value And RootDictionary(Root)(2).Scale = String.Empty) Then
                                        RootDictionary(Root)(If(RootDictionary(Root)(1).Scale = String.Empty, 1, 2)).Scale = Match.Groups(1).Value
                                        'RootDictionary(Root)(1) = "ya" + If(Root(0) = "A", "_#", Root(0)) + Match.Groups(1).Value + If(Match.Groups(1).Value = "a", "A", Root(1)) + If(Root(2) = "A", "^>", Root(2)) + "u"
                                    ElseIf Not Match.Success Or Not Match.Captures.Count = 1 Or (RootDictionary(Root)(1).Scale <> Match.Groups(1).Value And RootDictionary(Root)(2).Scale <> Match.Groups(1).Value) Then
                                        Debug.WriteLine(Root + " - " + Pieces(1) + " - " + RootDictionary(Root)(1).Scale + RootDictionary(Root)(2).Scale)
                                    End If
                                    If Match.Success And Match.Captures.Count = 1 Then RootDictionary(Root)(If(RootDictionary(Root)(1).Scale = Match.Groups(1).Value, 1, 2)).Refs.Add(Location)
                                ElseIf Root(1) = "y" Then
                                    'hamza of lam can make ayn on a, others on i
                                    'nyl and zyl exceptions with ayn on a
                                    Dim Match As Text.RegularExpressions.Match = System.Text.RegularExpressions.Regex.Match(Pieces(1), ("a" + Root(0) + "(.)" + If(Root(2) = "A", "(?:A|`|Y)?", "(?:" + Root(1) + "|Y|A)?") + If(Root(2) = "A", "(?:\^'|<|>|\^&)", Root(2))).Replace("$", "\$").Replace("*", "\*"))
                                    If Match.Success And Match.Captures.Count = 1 And (RootDictionary(Root)(1).Scale = String.Empty Or RootDictionary(Root)(1).Scale <> Match.Groups(1).Value And RootDictionary(Root)(2).Scale = String.Empty) Then
                                        RootDictionary(Root)(If(RootDictionary(Root)(1).Scale = String.Empty, 1, 2)).Scale = Match.Groups(1).Value
                                        'RootDictionary(Root)(1) = "ya" + Root(0) + Match.Groups(1).Value + If(Root(2) = "A", "A", If(Match.Groups(1).Value = "a", "A", Root(1))) + If(Root(2) = "A", "\^'", Root(2)) + "u"
                                    ElseIf Not Match.Success Or Not Match.Captures.Count = 1 Or (RootDictionary(Root)(1).Scale <> Match.Groups(1).Value And RootDictionary(Root)(2).Scale <> Match.Groups(1).Value) Then
                                        Debug.WriteLine(Root + " - " + Pieces(1) + " - " + RootDictionary(Root)(1).Scale + RootDictionary(Root)(2).Scale)
                                    End If
                                    If Match.Success And Match.Captures.Count = 1 Then RootDictionary(Root)(If(RootDictionary(Root)(1).Scale = Match.Groups(1).Value, 1, 2)).Refs.Add(Location)
                                Else
                                    Dim Match As Text.RegularExpressions.Match = System.Text.RegularExpressions.Regex.Match(Pieces(1), ("a(?:A@)?" + If(Root(0) = "A", "(?:A|>)", Root(0)) + "(?:o|\[)?" + If(Root(1) = "A", "_#", If(Root(1) = "s", "(?:S:|s)", Root(1))) + "(.)" + If(Root(2) = "A", "(?:>|'|_#|&)", Root(2) + If(Root(2) = "n", "?", String.Empty))).Replace("$", "\$").Replace("*", "\*"))
                                    If Match.Success And Match.Captures.Count = 1 And (RootDictionary(Root)(1).Scale = String.Empty Or RootDictionary(Root)(1).Scale <> Match.Groups(1).Value And RootDictionary(Root)(2).Scale = String.Empty) Then
                                        If Root(0) = "w" Then RootDictionary(Root)(If(RootDictionary(Root)(1).Scale = String.Empty, 1, 2)).LongShortBoth = "L"c
                                        RootDictionary(Root)(If(RootDictionary(Root)(1).Scale = String.Empty, 1, 2)).Scale = Match.Groups(1).Value
                                        'RootDictionary(Root)(1) = "ya" + If(Root(0) = "A", ">", Root(0)) + "o" + Root(1) + Match.Groups(1).Value + If(Root(2) = "A", If(Match.Groups(1).Value = "u", "&", If(Match.Groups(1).Value = "a", ">", "{")), Root(2)) + "u"
                                    ElseIf Not Match.Success Or Not Match.Captures.Count = 1 Or (RootDictionary(Root)(1).Scale <> Match.Groups(1).Value And RootDictionary(Root)(2).Scale <> Match.Groups(1).Value) Then
                                        Debug.WriteLine(Root + " - " + Pieces(1) + " - " + RootDictionary(Root)(1).Scale + RootDictionary(Root)(2).Scale)
                                    End If
                                    If Match.Success And Match.Captures.Count = 1 Then RootDictionary(Root)(If(RootDictionary(Root)(1).Scale = Match.Groups(1).Value, 1, 2)).Refs.Add(Location)
                                End If
                            End If
                        ElseIf Tense = 3 And Root.Length = 3 And Not bPassive Then
                            If Root(2) = "w" Then
                                '-oona merges into it
                                Dim Match As Text.RegularExpressions.Match = System.Text.RegularExpressions.Regex.Match(Pieces(1), ("\{" + If(Root(0) = "A", ">", Root(0)) + "o" + Root(1) + "(.)" + Root(2) + "?").Replace("$", "\$").Replace("*", "\*"))
                                If Match.Success And Match.Captures.Count = 1 And (RootDictionary(Root)(1).Scale = String.Empty Or RootDictionary(Root)(1).Scale <> Match.Groups(1).Value And RootDictionary(Root)(2).Scale = String.Empty) Then
                                    RootDictionary(Root)(If(RootDictionary(Root)(1).Scale = String.Empty, 1, 2)).Scale = Match.Groups(1).Value
                                    'RootDictionary(Root)(1) = "ya" + If(Root(0) = "A", ">", Root(0)) + "o" + Root(1) + Match.Groups(1).Value + Root(2)
                                ElseIf Not Match.Success Or Not Match.Captures.Count = 1 Or (RootDictionary(Root)(1).Scale <> Match.Groups(1).Value And RootDictionary(Root)(2).Scale <> Match.Groups(1).Value) Then
                                    Debug.WriteLine(Root + " - " + Pieces(1) + " - " + RootDictionary(Root)(1).Scale + RootDictionary(Root)(2).Scale)
                                End If
                                'If Match.Success And Match.Captures.Count = 1 Then RootDictionary(Root)(If(RootDictionary(Root)(1).Scale = Match.Groups(1).Value, 1, 2)).Refs.Add(Location)
                            ElseIf Root(0) = "w" And Root(2) = "y" Then
                                Dim Match As Text.RegularExpressions.Match = System.Text.RegularExpressions.Regex.Match(Pieces(1), (Root(1) + "(.)" + "(?:" + Root(2) + "|Y)?").Replace("$", "\$").Replace("*", "\*"))
                                If Match.Success And Match.Captures.Count = 1 And (RootDictionary(Root)(1).Scale = String.Empty Or RootDictionary(Root)(1).Scale <> If(Match.Groups(1).Value = "u", "i", Match.Groups(1).Value) And RootDictionary(Root)(2).Scale = String.Empty) Then
                                    RootDictionary(Root)(If(RootDictionary(Root)(1).Scale = String.Empty, 1, 2)).Scale = If(Match.Groups(1).Value = "u", "i", Match.Groups(1).Value)
                                    'RootDictionary(Root)(1) = "ya" + Root(1) + Match.Groups(1).Value + Root(2)
                                ElseIf Not Match.Success Or Not Match.Captures.Count = 1 Or (RootDictionary(Root)(1).Scale <> If(Match.Groups(1).Value = "u", "i", Match.Groups(1).Value) And RootDictionary(Root)(2).Scale <> If(Match.Groups(1).Value = "u", "i", Match.Groups(1).Value)) Then
                                    Debug.WriteLine(Root + " - " + Pieces(1) + " - " + RootDictionary(Root)(1).Scale + RootDictionary(Root)(2).Scale)
                                End If
                                'If Match.Success And Match.Captures.Count = 1 Then RootDictionary(Root)(If(RootDictionary(Root)(1).Scale = Match.Groups(1).Value, 1, 2)).Refs.Add(Location)
                            ElseIf Root(2) = "y" Then
                                '-oona changes to -awna fatha
                                '"AhHxgE" of ayn makes ayn on a, others on i
                                Dim Match As Text.RegularExpressions.Match = System.Text.RegularExpressions.Regex.Match(Pieces(1), ("\{?" + If(Root(0) = "A", "(?:A|>|\})", Root(0)) + "o?" + If(Root(1) = "y", "(?:y|Y)", If(Root(1) = "A", "_#", Root(1))) + "(.)" + "(?:" + Root(2) + "|Y)?").Replace("$", "\$").Replace("*", "\*"))
                                If Match.Success And Match.Captures.Count = 1 And (RootDictionary(Root)(1).Scale = String.Empty Or RootDictionary(Root)(1).Scale <> If(Match.Groups(1).Value = "u", "i", Match.Groups(1).Value) And RootDictionary(Root)(2).Scale = String.Empty) Then
                                    RootDictionary(Root)(If(RootDictionary(Root)(1).Scale = String.Empty, 1, 2)).Scale = If(Match.Groups(1).Value = "u", "i", Match.Groups(1).Value)
                                    'RootDictionary(Root)(1) = "ya" + If(Root(0) = "A", ">", Root(0)) + "o" + If(Root(1) = "A", "_#", Root(1)) + Match.Groups(1).Value + Root(2)
                                ElseIf Not Match.Success Or Not Match.Captures.Count = 1 Or (RootDictionary(Root)(1).Scale <> If(Match.Groups(1).Value = "u", "i", Match.Groups(1).Value) And RootDictionary(Root)(2).Scale <> If(Match.Groups(1).Value = "u", "i", Match.Groups(1).Value)) Then
                                    Debug.WriteLine(Root + " - " + Pieces(1) + " - " + RootDictionary(Root)(1).Scale + RootDictionary(Root)(2).Scale)
                                End If
                                'If Match.Success And Match.Captures.Count = 1 Then RootDictionary(Root)(If(RootDictionary(Root)(1).Scale = Match.Groups(1).Value, 1, 2)).Refs.Add(Location)
                            ElseIf Root(1) = Root(2) And System.Text.RegularExpressions.Regex.Match(Pieces(1), (If(Root(0) = "A", ">", Root(0)) + "(.)(?:" + If(Root(1) = "A", "(?:>|'|_#|&)", Root(1)) + "~|(.)" + Root(2) + ")").Replace("$", "\$").Replace("*", "\*")).Success Then
                                Dim Match As Text.RegularExpressions.Match = System.Text.RegularExpressions.Regex.Match(Pieces(1), (If(Root(0) = "A", ">", Root(0)) + "(.)" + If(Root(1) = "A", "(?:>|'|_#|&)", Root(1)) + "~").Replace("$", "\$").Replace("*", "\*"))
                                If Match.Success And Match.Captures.Count = 1 And (RootDictionary(Root)(1).Scale = String.Empty Or RootDictionary(Root)(1).Scale <> Match.Groups(1).Value And RootDictionary(Root)(2).Scale = String.Empty) Then
                                    RootDictionary(Root)(If(RootDictionary(Root)(1).Scale = String.Empty, 1, 2)).Scale = If(Match.Groups(2).Success, Match.Groups(2).Value, Match.Groups(1).Value)
                                    'RootDictionary(Root)(1) = "ya" + If(Root(0) = "A", "&", Root(0)) + Match.Groups(1).Value + Root(1) + "~u"
                                ElseIf Not Match.Success Or Not Match.Captures.Count = 1 Or (RootDictionary(Root)(1).Scale <> Match.Groups(1).Value And RootDictionary(Root)(2).Scale <> Match.Groups(1).Value) Then
                                    Debug.WriteLine(Root + " - " + Pieces(1) + " - " + RootDictionary(Root)(1).Scale + RootDictionary(Root)(2).Scale)
                                End If
                                'If Match.Success And Match.Captures.Count = 1 Then RootDictionary(Root)(If(RootDictionary(Root)(1).Scale = Match.Groups(1).Value, 1, 2)).Refs.Add(Location)
                                If Match.Success And Match.Captures.Count = 1 Then RootDictionary(Root)(If(RootDictionary(Root)(1).Scale = Match.Groups(1).Value, 1, 2)).LongShortBoth = If(If(Match.Groups(2).Success, "S"c, "L"c) = RootDictionary(Root)(If(RootDictionary(Root)(1).Scale = Match.Groups(1).Value, 1, 2)).LongShortBoth, "B"c, If(Match.Groups(2).Success, "L"c, "S"c))
                            ElseIf Root(0) = "w" And System.Text.RegularExpressions.Regex.Match(Pieces(1), ("^" + Root(1) + "(.)" + If(Root(2) = "A", "_#", Root(2))).Replace("$", "\$").Replace("*", "\*")).Success Then
                                '"AhHxgE" of lam makes ayn on a, others on i
                                Dim Match As Text.RegularExpressions.Match = System.Text.RegularExpressions.Regex.Match(Pieces(1), ("^" + Root(1) + "(.)" + If(Root(2) = "A", "_#", Root(2))).Replace("$", "\$").Replace("*", "\*"))
                                If Match.Success And Match.Captures.Count = 1 And (RootDictionary(Root)(1).Scale = String.Empty Or RootDictionary(Root)(1).Scale <> Match.Groups(1).Value And RootDictionary(Root)(2).Scale = String.Empty) Then
                                    RootDictionary(Root)(If(RootDictionary(Root)(1).Scale = String.Empty, 1, 2)).LongShortBoth = "S"c
                                    RootDictionary(Root)(If(RootDictionary(Root)(1).Scale = String.Empty, 1, 2)).Scale = Match.Groups(1).Value
                                    'RootDictionary(Root)(1) = "ya" + Root(1) + Match.Groups(1).Value + If(Root(2) = "A", ">", Root(2)) + "u"
                                ElseIf Not Match.Success Or Not Match.Captures.Count = 1 Or (RootDictionary(Root)(1).Scale <> Match.Groups(1).Value And RootDictionary(Root)(2).Scale <> Match.Groups(1).Value) Then
                                    Debug.WriteLine(Root + " - " + Pieces(1) + " - " + RootDictionary(Root)(1).Scale + RootDictionary(Root)(2).Scale)
                                End If
                                'If Match.Success And Match.Captures.Count = 1 Then RootDictionary(Root)(If(RootDictionary(Root)(1).Scale = Match.Groups(1).Value, 1, 2)).Refs.Add(Location)
                            ElseIf Root(1) = "w" Or (Root(1) = "A" And System.Text.RegularExpressions.Regex.Match(Pieces(1), (If(Root(0) = "A", "_#", Root(0)) + "(.)" + "(?:" + Root(1) + "|A)?" + If(Root(2) = "A", "\^>|&", Root(2))).Replace("$", "\$").Replace("*", "\*")).Success) Then
                                'kwd, xwf exceptions with ayn on a
                                Dim Match As Text.RegularExpressions.Match = System.Text.RegularExpressions.Regex.Match(Pieces(1), (If(Root(0) = "A", "_#", Root(0)) + "(.)" + "(" + Root(1) + "|A)?" + If(Root(2) = "A", "\^>|&", Root(2))).Replace("$", "\$").Replace("*", "\*"))
                                If Match.Success And Match.Captures.Count = 1 And (RootDictionary(Root)(1).Scale = String.Empty Or RootDictionary(Root)(1).Scale <> Match.Groups(1).Value And RootDictionary(Root)(2).Scale = String.Empty) Then
                                    RootDictionary(Root)(If(RootDictionary(Root)(1).Scale = String.Empty, 1, 2)).Scale = Match.Groups(1).Value
                                    'RootDictionary(Root)(1) = "ya" + If(Root(0) = "A", "_#", Root(0)) + Match.Groups(1).Value + If(Match.Groups(1).Value = "a", "A", Root(1)) + If(Root(2) = "A", "^>", Root(2)) + "u"
                                ElseIf Not Match.Success Or Not Match.Captures.Count = 1 Or (RootDictionary(Root)(1).Scale <> Match.Groups(1).Value And RootDictionary(Root)(2).Scale <> Match.Groups(1).Value) Then
                                    Debug.WriteLine(Root + " - " + Pieces(1) + " - " + RootDictionary(Root)(1).Scale + RootDictionary(Root)(2).Scale)
                                End If
                                'If Match.Success And Match.Captures.Count = 1 Then RootDictionary(Root)(If(RootDictionary(Root)(1).Scale = Match.Groups(1).Value, 1, 2)).Refs.Add(Location)
                                If Match.Success And Match.Captures.Count = 1 And Root(1) = "A" Then RootDictionary(Root)(If(RootDictionary(Root)(1).Scale = Match.Groups(1).Value, 1, 2)).LongShortBoth = If(If(Match.Groups(2).Success, "S"c, "L"c) = RootDictionary(Root)(If(RootDictionary(Root)(1).Scale = Match.Groups(1).Value, 1, 2)).LongShortBoth, "B"c, If(Match.Groups(2).Success, "L"c, "S"c))
                            ElseIf Root(1) = "y" Then
                                'hamza of lam can make ayn on a, others on i
                                'nyl and zyl exceptions with ayn on a
                                Dim Match As Text.RegularExpressions.Match = System.Text.RegularExpressions.Regex.Match(Pieces(1), (Root(0) + "(.)" + If(Root(2) = "A", "A?", "(?:" + Root(1) + "o?|Y|A)?") + If(Root(2) = "A", "(?:\^'|\})", Root(2))).Replace("$", "\$").Replace("*", "\*"))
                                If Match.Success And Match.Captures.Count = 1 And (RootDictionary(Root)(1).Scale = String.Empty Or RootDictionary(Root)(1).Scale <> Match.Groups(1).Value And RootDictionary(Root)(2).Scale = String.Empty) Then
                                    RootDictionary(Root)(If(RootDictionary(Root)(1).Scale = String.Empty, 1, 2)).Scale = Match.Groups(1).Value
                                    'RootDictionary(Root)(1) = "ya" + Root(0) + Match.Groups(1).Value + If(Root(2) = "A", "A", If(Match.Groups(1).Value = "a", "A", Root(1))) + If(Root(2) = "A", "\^'", Root(2)) + "u"
                                ElseIf Not Match.Success Or Not Match.Captures.Count = 1 Or (RootDictionary(Root)(1).Scale <> Match.Groups(1).Value And RootDictionary(Root)(2).Scale <> Match.Groups(1).Value) Then
                                    Debug.WriteLine(Root + " - " + Pieces(1) + " - " + RootDictionary(Root)(1).Scale + RootDictionary(Root)(2).Scale)
                                End If
                                'If Match.Success And Match.Captures.Count = 1 Then RootDictionary(Root)(If(RootDictionary(Root)(1).Scale = Match.Groups(1).Value, 1, 2)).Refs.Add(Location)
                            ElseIf Root(0) = "A" Then
                                Dim Match As Text.RegularExpressions.Match = System.Text.RegularExpressions.Regex.Match(Pieces(1), ("(&?)" + If(Root(1) = "A", "_#", Root(1)) + "(.)" + If(Root(2) = "A", "(?:>|'|_#|&)", Root(2))).Replace("$", "\$").Replace("*", "\*"))
                                If Match.Success And Match.Captures.Count = 1 And (RootDictionary(Root)(1).Scale = String.Empty Or RootDictionary(Root)(1).Scale <> Match.Groups(2).Value And RootDictionary(Root)(2).Scale = String.Empty) Then
                                    RootDictionary(Root)(If(RootDictionary(Root)(1).Scale = String.Empty, 1, 2)).LongShortBoth = If(Match.Groups(1).Success, "L"c, "S"c)
                                    RootDictionary(Root)(If(RootDictionary(Root)(1).Scale = String.Empty, 1, 2)).Scale = Match.Groups(2).Value
                                    'RootDictionary(Root)(1) = "{" + If(Root(0) = "A", ">", Root(0)) + "o" + Root(1) + Match.Groups(1).Value + If(Root(2) = "A", If(Match.Groups(1).Value = "u", "&", If(Match.Groups(1).Value = "a", ">", "{")), Root(2)) + "u"
                                ElseIf Not Match.Success Or Not Match.Captures.Count = 1 Or (RootDictionary(Root)(1).Scale <> Match.Groups(2).Value And RootDictionary(Root)(2).Scale <> Match.Groups(2).Value) Then
                                    Debug.WriteLine(Root + " - " + Pieces(1) + " - " + RootDictionary(Root)(1).Scale + RootDictionary(Root)(2).Scale)
                                End If
                                'If Match.Success And Match.Captures.Count = 1 Then RootDictionary(Root)(If(RootDictionary(Root)(1).Scale = Match.Groups(1).Value, 1, 2)).Refs.Add(Location)
                            Else
                                Dim Match As Text.RegularExpressions.Match = System.Text.RegularExpressions.Regex.Match(Pieces(1), ("\{?" + If(Root(0) = "A", "(?:A|>)", Root(0)) + "(?:\[|o)?" + If(Root(1) = "A", "_#", Root(1)) + "(.)" + If(Root(2) = "A", "(?:>|'|_#|&)", Root(2))).Replace("$", "\$").Replace("*", "\*"))
                                If Match.Success And Match.Captures.Count = 1 And (RootDictionary(Root)(1).Scale = String.Empty Or RootDictionary(Root)(1).Scale <> Match.Groups(1).Value And RootDictionary(Root)(2).Scale = String.Empty) Then
                                    If Root(0) = "w" Then RootDictionary(Root)(If(RootDictionary(Root)(1).Scale = String.Empty, 1, 2)).LongShortBoth = "L"c
                                    RootDictionary(Root)(If(RootDictionary(Root)(1).Scale = String.Empty, 1, 2)).Scale = Match.Groups(1).Value
                                    'RootDictionary(Root)(1) = "{" + If(Root(0) = "A", ">", Root(0)) + "o" + Root(1) + Match.Groups(1).Value + If(Root(2) = "A", If(Match.Groups(1).Value = "u", "&", If(Match.Groups(1).Value = "a", ">", "{")), Root(2)) + "u"
                                ElseIf Not Match.Success Or Not Match.Captures.Count = 1 Or (RootDictionary(Root)(1).Scale <> Match.Groups(1).Value And RootDictionary(Root)(2).Scale <> Match.Groups(1).Value) Then
                                    Debug.WriteLine(Root + " - " + Pieces(1) + " - " + RootDictionary(Root)(1).Scale + RootDictionary(Root)(2).Scale)
                                End If
                                'If Match.Success And Match.Captures.Count = 1 Then RootDictionary(Root)(If(RootDictionary(Root)(1).Scale = Match.Groups(1).Value, 1, 2)).Refs.Add(Location)
                            End If
                        End If
                    End If
                End If
            End If
        Next
        Dim Arr(RootDictionary.Count - 1) As String
        RootDictionary.Keys.CopyTo(Arr, 0)
        Dim VerbExceptionsTable As String()() = {
            New String() {"hyA", "uu"}, New String() {"nhw", "uu"}, New String() {"lbb", "ua"}, New String() {"lbb", "ia"},
            New String() {"wrv", "ii"}, New String() {"wvq", "ii"}, New String() {"wmq", "ii"}, New String() {"wfq", "ii"}, New String() {"wrm", "ii"}, New String() {"wry", "ii"}, New String() {"wjd", "ii"}, New String() {"wEq", "ii"}, New String() {"wly", "ii"}, New String() {"wbq", "ii"}, New String() {"wrE", "ii"}, New String() {"wrk", "ii"}, New String() {"wkm", "ii"}, New String() {"wqh", "ii"},
            New String() {"Hsb", "iai"}, New String() {"nEm", "iai"}, New String() {"yAs", "iai"}, New String() {"bAs", "iai"}, New String() {"ybs", "iai"}, New String() {"wHr", "iai"}, New String() {"wgr", "iai"}, New String() {"whl", "iai"}, New String() {"wlh", "iai"},
            New String() {"mrr", "au"}, New String() {"Emm", "au"}, New String() {"hmm", "au"}, New String() {"$kk", "au"},
            New String() {"Sdd", "aui"}, New String() {"vrr", "aui"}, New String() {"$bb", "aui"}, New String() {"drr", "aui"}, New String() {"jdd", "aui"}, New String() {"xrr", "aui"}, New String() {"Enn", "aui"},
            New String() {"dxl", "au"}, New String() {"TlE", "au"}, New String() {"frg", "au"}, New String() {"qEd", "au"}, New String() {"nfx", "au"}, New String() {"zEm", "au"}, New String() {"blg", "au"}, New String() {"mHw", "au"}, New String() {"Ax*", "au"},
            New String() {"rjE", "ai"}, New String() {"wEZ", "ai"}, New String() {"whn", "ai"}, New String() {"wEd", "ai"}, New String() {"zyg", "ai"}, New String() {"nzE", "ai"}, New String() {"nkH", "ai"},
            New String() {"frq", "aua"}, New String() {"blw", "aua"},
            New String() {"lbs", "aia"}, New String() {"hwy", "aai"},
            New String() {"nyl", "aa"}, New String() {"zyl", "aa"}, New String() {"kwd", "aa"}, New String() {"xwf", "aa"}, New String() {"qnT", "aa"}
            }
        Array.Sort(Arr, New VerbRootComparer(RootDictionary, VerbExceptionsTable))
        Dim Output As New List(Of Object())
        Dim IndexToChunk As Integer()() = Nothing
        TR.QuranTextCombiner(XMLDocMain, IndexToChunk, True)
        Dim VerbMinimumTable As String()() = {
                New String() {"  w    ", "au"}, New String() {"Aw     ", "au"}, New String() {"y      ", "ia"}, New String() {"w -    ", "aa"}, New String() {" y-    ", "ia"}, New String() {" Ay    ", "aa"},
                New String() {"w  ia?L", "ia"}, New String() {"w  uu? ", "uuL"}, New String() {" A aa?S", "aa"}, New String() {" A aa?B", "aa"}, New String() {" wyai? ", "ai"}, New String() {" wyia? ", "ia"}, New String() {"  Auu? ", "uu"}, New String() {" A uu? ", "uu"}, New String() {"A  ia? ", "ia"},
                New String() {"  yi   ", "ia"}, New String() {"w yi   ", "ii"}, New String() {"w ya   ", "ai"}, New String() {"  Aa   ", "aa"}, New String() {"  Ai   ", "ia"}, New String() {" A i   ", "ia"},
                New String() {" yA i  ", "ai"}, New String() {" yA a  ", "aa"}, New String() {" wA a  ", "aa"}, New String() {" wA u  ", "au"}, New String() {"  - u  ", "au"}, New String() {"  - a  ", "ia"}, New String() {"  - i  ", "ai"}, New String() {"  y i  ", "ai"}, New String() {" w  u  ", "au"}, New String() {" w  a  ", "aa"}, New String() {" y  a  ", "aa"}, New String() {" y  i  ", "ai"}, New String() {"A   i  ", "ai"}, New String() {"A   u  S", "au"}, New String() {"A   u  L", "au"}, New String() {"A y i  ", "ai"}, New String() {"A y a  ", "aa"},
                New String() {"w  aa  ", "aa"}, New String() {"w  ia  ", "ia"}, New String() {"w  ai  ", "ai"}, New String() {"w  ii  ", "ii"}, New String() {"  yaa  ", "aa"}, New String() {"   au  ", "au"}, New String() {"   ai  ", "ai"}, New String() {"   aa  ", "aa"}, New String() {"   ia  ", "ia"}, New String() {"   uu  ", "uu"}, New String() {"   ii  ", "ii"}, New String() {"   aui ", "aui"}, New String() {"   aua ", "aua"},
                New String() {"   iua ", "iua"}, New String() {"   iu  ", "iua"}
               }
        For KeyCount As Integer = 0 To Arr.Length - 1
            Dim Count As Integer
            For Count = 0 To VerbExceptionsTable.Length - 1
                If Arr(KeyCount)(0) = VerbExceptionsTable(Count)(0)(0) And
                    Arr(KeyCount)(1) = VerbExceptionsTable(Count)(0)(1) And
                    Arr(KeyCount)(2) = VerbExceptionsTable(Count)(0)(2) Then
                    Debug.WriteLine("Exceptional: " + Arr(KeyCount) + ": " + RootDictionary(Arr(KeyCount))(0).Scale + "-" + RootDictionary(Arr(KeyCount))(1).Scale + RootDictionary(Arr(KeyCount))(2).Scale)
                    Exit For
                End If
            Next
            If RootDictionary(Arr(KeyCount))(0).Scale = "a" And If(RootDictionary(Arr(KeyCount))(1).Scale = "a" Or RootDictionary(Arr(KeyCount))(2).Scale = "a", Not "AhHxgE".Contains(Arr(KeyCount)(1)) And Not "AhHxgE".Contains(Arr(KeyCount)(2)), (RootDictionary(Arr(KeyCount))(1).Scale = "u" Or RootDictionary(Arr(KeyCount))(1).Scale = "i") And ("AhHxgE".Contains(Arr(KeyCount)(1)) Or "AhHxgE".Contains(Arr(KeyCount)(2)))) Then Debug.WriteLine("Throat Exception: " + Arr(KeyCount) + ": " + RootDictionary(Arr(KeyCount))(0).Scale + "-" + RootDictionary(Arr(KeyCount))(1).Scale + RootDictionary(Arr(KeyCount))(2).Scale)
            For Count = 0 To VerbMinimumTable.Length - 1
                If (VerbMinimumTable(Count)(0)(0) = " " Or Arr(KeyCount)(0) = VerbMinimumTable(Count)(0)(0)) And
                    (VerbMinimumTable(Count)(0)(1) = " " Or Arr(KeyCount)(1) = VerbMinimumTable(Count)(0)(1)) And
                    (VerbMinimumTable(Count)(0)(2) = " " Or Arr(KeyCount)(2) = If(VerbMinimumTable(Count)(0)(2) = "-", Arr(KeyCount)(1), VerbMinimumTable(Count)(0)(2))) And
                    If(VerbMinimumTable(Count)(0)(5) = "?", (VerbMinimumTable(Count)(0)(3) = " " Or VerbMinimumTable(Count)(0)(3) = RootDictionary(Arr(KeyCount))(0).Scale) Or
                    (VerbMinimumTable(Count)(0)(4) = " " Or VerbMinimumTable(Count)(0)(4) = RootDictionary(Arr(KeyCount))(1).Scale),
                    (VerbMinimumTable(Count)(0)(3) = " " Or VerbMinimumTable(Count)(0)(3) = RootDictionary(Arr(KeyCount))(0).Scale) And
                    (VerbMinimumTable(Count)(0)(4) = " " Or VerbMinimumTable(Count)(0)(4) = RootDictionary(Arr(KeyCount))(1).Scale) And
                    (VerbMinimumTable(Count)(0)(5) = " " Or VerbMinimumTable(Count)(0)(5) = RootDictionary(Arr(KeyCount))(2).Scale) Or
                    (VerbMinimumTable(Count)(0)(4) = " " Or VerbMinimumTable(Count)(0)(4) = RootDictionary(Arr(KeyCount))(2).Scale) And
                    (VerbMinimumTable(Count)(0)(5) = " " Or VerbMinimumTable(Count)(0)(5) = RootDictionary(Arr(KeyCount))(1).Scale)) And
                    (VerbMinimumTable(Count)(0)(6) = " " Or VerbMinimumTable(Count)(0)(6) = "B" Or RootDictionary(Arr(KeyCount))(1).LongShortBoth = " "c Or VerbMinimumTable(Count)(0)(6) = RootDictionary(Arr(KeyCount))(1).LongShortBoth) Then
                    If Not (VerbMinimumTable(Count)(0)(6) = " " Or VerbMinimumTable(Count)(0)(6) = "B" Or VerbMinimumTable(Count)(0)(6) = RootDictionary(Arr(KeyCount))(1).LongShortBoth) Then Debug.WriteLine("Command Filter Ignore: " + Arr(KeyCount) + ": " + RootDictionary(Arr(KeyCount))(0).Scale + "-" + RootDictionary(Arr(KeyCount))(1).Scale + RootDictionary(Arr(KeyCount))(2).Scale + RootDictionary(Arr(KeyCount))(1).LongShortBoth)
                    Exit For
                End If
            Next
            If Count = VerbMinimumTable.Length Then Debug.WriteLine("Lack Info: " + Arr(KeyCount) + ": " + RootDictionary(Arr(KeyCount))(0).Scale + "-" + RootDictionary(Arr(KeyCount))(1).Scale + RootDictionary(Arr(KeyCount))(2).Scale)
            'filter besides past, those who root only or present tense only make full determination
            If Count <> VerbMinimumTable.Length AndAlso (RootDictionary(Arr(KeyCount))(0).Refs.Count <> 0 Or Count <> VerbMinimumTable.Length) AndAlso RootDictionary(Arr(KeyCount))(1).Refs.Count <> 0 Then
                RootDictionary(Arr(KeyCount))(0).Scale = VerbMinimumTable(Count)(1)(0)
                RootDictionary(Arr(KeyCount))(1).Scale = VerbMinimumTable(Count)(1)(1)
                RootDictionary(Arr(KeyCount))(1).LongShortBoth = If(VerbMinimumTable(Count)(1).Length > 2 AndAlso VerbMinimumTable(Count)(1)(2) = "L", VerbMinimumTable(Count)(1)(2), VerbMinimumTable(Count)(0)(6))
                RootDictionary(Arr(KeyCount))(2).Scale = If(VerbMinimumTable(Count)(1).Length > 2 AndAlso VerbMinimumTable(Count)(1)(2) <> "L", VerbMinimumTable(Count)(1)(2), String.Empty)
                Dim PastV As String = String.Empty
                Dim PresV As String = String.Empty
                Dim PresVOth As String = String.Empty
                If Arr(KeyCount)(2) = "y" Then
                    If RootDictionary(Arr(KeyCount))(0).Scale <> String.Empty Then PastV = If(Arr(KeyCount)(0) = "A", "'", Arr(KeyCount)(0)) + "a" + If(Arr(KeyCount)(1) = "A", "'", Arr(KeyCount)(1)) + RootDictionary(Arr(KeyCount))(0).Scale + "Y"
                    If RootDictionary(Arr(KeyCount))(1).Scale <> String.Empty Then PresV = "ya" + If(Arr(KeyCount)(0) = "A", "'", Arr(KeyCount)(0)) + "o" + If(Arr(KeyCount)(1) = "A", "'", Arr(KeyCount)(1)) + RootDictionary(Arr(KeyCount))(1).Scale + "Y"
                ElseIf Arr(KeyCount)(1) = "w" Or Arr(KeyCount)(1) = "y" Then
                    If RootDictionary(Arr(KeyCount))(0).Scale <> String.Empty Then PastV = If(Arr(KeyCount)(0) = "A", "'", Arr(KeyCount)(0)) + "aA" + If(Arr(KeyCount)(2) = "A", "'", Arr(KeyCount)(2)) + "a"
                    If RootDictionary(Arr(KeyCount))(1).Scale = "a" Then
                        If RootDictionary(Arr(KeyCount))(1).Scale <> String.Empty Then PresV = "ya" + If(Arr(KeyCount)(0) = "A", "'", Arr(KeyCount)(0)) + RootDictionary(Arr(KeyCount))(1).Scale + "A" + If(Arr(KeyCount)(2) = "A", "'", Arr(KeyCount)(2)) + "u"
                    ElseIf RootDictionary(Arr(KeyCount))(1).Scale = "i" Then
                        If RootDictionary(Arr(KeyCount))(1).Scale <> String.Empty Then PresV = "ya" + If(Arr(KeyCount)(0) = "A", "'", Arr(KeyCount)(0)) + RootDictionary(Arr(KeyCount))(1).Scale + "y" + If(Arr(KeyCount)(2) = "A", "'", Arr(KeyCount)(2)) + "u"
                    Else
                        If RootDictionary(Arr(KeyCount))(1).Scale <> String.Empty Then PresV = "ya" + If(Arr(KeyCount)(0) = "A", "'", Arr(KeyCount)(0)) + RootDictionary(Arr(KeyCount))(1).Scale + "w" + If(Arr(KeyCount)(2) = "A", "'", Arr(KeyCount)(2)) + "u"
                    End If
                ElseIf Arr(KeyCount)(1) = Arr(KeyCount)(2) Then
                    If RootDictionary(Arr(KeyCount))(0).Scale <> String.Empty Then PastV = If(Arr(KeyCount)(0) = "A", "'", Arr(KeyCount)(0)) + "a" + Arr(KeyCount)(1) + "~a"
                    If RootDictionary(Arr(KeyCount))(1).Scale <> String.Empty Then PresV = "ya" + If(Arr(KeyCount)(0) = "A", "'", Arr(KeyCount)(0)) + RootDictionary(Arr(KeyCount))(1).Scale + Arr(KeyCount)(1) + "~u"
                Else
                    If Arr(KeyCount)(2) = "w" Then
                        If RootDictionary(Arr(KeyCount))(0).Scale <> String.Empty Then PastV = If(Arr(KeyCount)(0) = "A", "'", Arr(KeyCount)(0)) + "a" + If(Arr(KeyCount)(1) = "A", "'", Arr(KeyCount)(1)) + RootDictionary(Arr(KeyCount))(0).Scale + "A"
                    Else
                        If RootDictionary(Arr(KeyCount))(0).Scale <> String.Empty Then PastV = If(Arr(KeyCount)(0) = "A", "'", Arr(KeyCount)(0)) + "a" + If(Arr(KeyCount)(1) = "A", "'", Arr(KeyCount)(1)) + RootDictionary(Arr(KeyCount))(0).Scale + If(Arr(KeyCount)(2) = "A", "'", Arr(KeyCount)(2)) + "a"
                    End If
                    If Arr(KeyCount)(0) = "w" And (RootDictionary(Arr(KeyCount))(0).Scale = "a" Or RootDictionary(Arr(KeyCount))(0).Scale = "i" And RootDictionary(Arr(KeyCount))(1).Scale = "i") Then
                        If RootDictionary(Arr(KeyCount))(1).Scale <> String.Empty Then PresV = "ya" + If(Arr(KeyCount)(1) = "A", "'", Arr(KeyCount)(1)) + RootDictionary(Arr(KeyCount))(1).Scale + If(Arr(KeyCount)(2) = "A", "'", Arr(KeyCount)(2)) + "u"
                    Else
                        If RootDictionary(Arr(KeyCount))(1).Scale <> String.Empty Then PresV = "ya" + If(Arr(KeyCount)(0) = "A", "'", Arr(KeyCount)(0)) + "o" + If(Arr(KeyCount)(1) = "A", "'", Arr(KeyCount)(1)) + RootDictionary(Arr(KeyCount))(1).Scale + If(Arr(KeyCount)(2) = "A", "'", Arr(KeyCount)(2)) + "u"
                        If RootDictionary(Arr(KeyCount))(2).Scale <> String.Empty Then PresVOth = "ya" + If(Arr(KeyCount)(0) = "A", "'", Arr(KeyCount)(0)) + "o" + If(Arr(KeyCount)(1) = "A", "'", Arr(KeyCount)(1)) + RootDictionary(Arr(KeyCount))(2).Scale + If(Arr(KeyCount)(2) = "A", "'", Arr(KeyCount)(2)) + "u"
                    End If
                End If
                If PastV <> String.Empty Then PastV = CombineFixHamzas(Arb.TransliterateFromBuckwalter(PastV))
                If PresV <> String.Empty Then PresV = CombineFixHamzas(Arb.TransliterateFromBuckwalter(PresV))
                If PresVOth <> String.Empty Then PresVOth = CombineFixHamzas(Arb.TransliterateFromBuckwalter(PresVOth))
                'Debug.WriteLine(PastV + " - " + PresV + " - " + PresVOth + " - " + Arr(KeyCount) + ": " + RootDictionary(Arr(KeyCount))(0).Scale + "-" + RootDictionary(Arr(KeyCount))(1).Scale + RootDictionary(Arr(KeyCount))(2).Scale)
                Dim Renderer() As RenderArray = {New RenderArray(String.Empty), New RenderArray(String.Empty), New RenderArray(String.Empty)}
                For RendCount As Integer = 0 To Renderer.Length - 1
                    For Count = 0 To RootDictionary(Arr(KeyCount))(RendCount).Refs.Count - 1
                        Dim IndexToVerse As Integer()() = Nothing
                        Dim Chunk As Integer = Array.BinarySearch(IndexToChunk, New Integer() {RootDictionary(Arr(KeyCount))(RendCount).Refs(Count)(0), RootDictionary(Arr(KeyCount))(RendCount).Refs(Count)(1), RootDictionary(Arr(KeyCount))(RendCount).Refs(Count)(2), RootDictionary(Arr(KeyCount))(RendCount).Refs(Count)(0), RootDictionary(Arr(KeyCount))(RendCount).Refs(Count)(1), RootDictionary(Arr(KeyCount))(RendCount).Refs(Count)(2)}, New TanzilReader.QuranWordChapterVerseWordComparer(False))
                        Dim Idx As Integer = Linq.Enumerable.Last(Linq.Enumerable.TakeWhile(Linq.Enumerable.Reverse(New List(Of Integer())(Linq.Enumerable.Take(IndexToChunk, Chunk + 1))), Function(It) It(0) = IndexToChunk(Chunk)(0) And It(1) = IndexToChunk(Chunk)(1) And It(5) = IndexToChunk(Chunk)(5)))(2)
                        Dim LastIdx As Integer = Linq.Enumerable.Last(Linq.Enumerable.TakeWhile(New List(Of Integer())(Linq.Enumerable.Skip(IndexToChunk, Chunk)), Function(It) It(0) = IndexToChunk(Chunk)(0) And It(1) = IndexToChunk(Chunk)(1) And It(5) = IndexToChunk(Chunk)(5)))(2)
                        Dim QuranText As String = TR.QuranTextCombiner(XMLDocMain, IndexToVerse, False, RootDictionary(Arr(KeyCount))(RendCount).Refs(Count)(0), RootDictionary(Arr(KeyCount))(RendCount).Refs(Count)(1), Idx, RootDictionary(Arr(KeyCount))(RendCount).Refs(Count)(0), RootDictionary(Arr(KeyCount))(RendCount).Refs(Count)(1), LastIdx)
                        Renderer(RendCount).Items.AddRange(TR.DoGetRenderedQuranText(QuranText, IndexToVerse, {}, {}, ArabicData.TranslitScheme.None, String.Empty, {}, {}, False, False, False, False, False, True, True).Items)
                    Next
                Next
                Output.Add(New Object() {Arb.TransliterateFromBuckwalter(Arr(KeyCount)), Arb.TransliterateFromBuckwalter(RootDictionary(Arr(KeyCount))(0).Scale), Arb.TransliterateFromBuckwalter(RootDictionary(Arr(KeyCount))(1).Scale), Arb.TransliterateFromBuckwalter(RootDictionary(Arr(KeyCount))(2).Scale), PastV, PresV, PresVOth, Renderer(0).Items.ToArray(), Renderer(1).Items.ToArray(), Renderer(2).Items.ToArray()})
            End If
        Next
        Return Output.ToArray()
    End Function
    Public Class VerbRootComparer
        Implements IComparer(Of String)
        Dim Dict As Dictionary(Of String, ScaleRef())
        Dim VerbExceptionsTable As String()()
        Public Sub New(NewDict As Dictionary(Of String, ScaleRef()), VET As String()())
            Dict = NewDict
            VerbExceptionsTable = VET
        End Sub
        Public Function ScaleComp(x As ScaleRef(), y As ScaleRef()) As Integer
            If x(0).Scale = y(0).Scale Then Return If(x(1).Scale = Nothing, String.Empty, x(1).Scale).CompareTo(If(y(1).Scale = Nothing, String.Empty, y(1).Scale))
            Return If(x(0).Scale = Nothing, String.Empty, x(0).Scale).CompareTo(If(y(0).Scale = Nothing, String.Empty, y(0).Scale))
        End Function
        Public Function Compare(x As String, y As String) As Integer Implements IComparer(Of String).Compare
            Dim Count As Integer
            Dim OthCount As Integer
            For Count = 0 To VerbExceptionsTable.Length - 1
                If x(0) = VerbExceptionsTable(Count)(0)(0) And
                    x(1) = VerbExceptionsTable(Count)(0)(1) And
                    x(2) = VerbExceptionsTable(Count)(0)(2) Then
                    Exit For
                End If
            Next
            For OthCount = 0 To VerbExceptionsTable.Length - 1
                If y(0) = VerbExceptionsTable(OthCount)(0)(0) And
                    y(1) = VerbExceptionsTable(OthCount)(0)(1) And
                    y(2) = VerbExceptionsTable(OthCount)(0)(2) Then
                    Exit For
                End If
            Next
            If OthCount <> VerbExceptionsTable.Length Or Count <> VerbExceptionsTable.Length Then Return If(Count = VerbExceptionsTable.Length, 1, If(OthCount = VerbExceptionsTable.Length, -1, Count - OthCount))
            If x(0) = "w" And x(2) = "y" And y(0) = "w" And y(2) = "y" Then Return 0
            If x(0) = "w" And x(2) = "y" Then Return -1
            If y(0) = "w" And y(2) = "y" Then Return 1
            If x(2) = "y" And y(2) = "y" Then Return ScaleComp(Dict(x), Dict(y))
            If x(2) = "y" Then Return -1
            If y(2) = "y" Then Return 1
            If x(2) = "w" And y(2) = "w" Then Return ScaleComp(Dict(x), Dict(y))
            If x(2) = "w" Then Return -1
            If y(2) = "w" Then Return 1
            If x(1) = x(2) And y(1) = y(2) Then Return ScaleComp(Dict(x), Dict(y))
            If x(1) = x(2) Then Return -1
            If y(1) = y(2) Then Return 1
            If x(0) = "w" And y(0) = "w" Then Return ScaleComp(Dict(x), Dict(y))
            If x(0) = "w" Then Return -1
            If y(0) = "w" Then Return 1
            If x(1) = "y" And y(1) = "y" Then Return ScaleComp(Dict(x), Dict(y))
            If x(1) = "y" Then Return -1
            If y(1) = "y" Then Return 1
            If x(1) = "w" And y(1) = "w" Then Return ScaleComp(Dict(x), Dict(y))
            If x(1) = "w" Then Return -1
            If y(1) = "w" Then Return 1
            If x(0) = "A" And y(0) = "A" Then Return ScaleComp(Dict(x), Dict(y))
            If x(0) = "A" Then Return -1
            If y(0) = "A" Then Return 1
            If x(1) = "A" And y(1) = "A" Then Return ScaleComp(Dict(x), Dict(y))
            If x(1) = "A" Then Return -1
            If y(1) = "A" Then Return 1
            If x(2) = "A" And y(2) = "A" Then Return ScaleComp(Dict(x), Dict(y))
            If x(2) = "A" Then Return -1
            If y(2) = "A" Then Return 1
            Return ScaleComp(Dict(x), Dict(y))
        End Function
    End Class
    Public Function GetMorphologicalDataForWord(Chapter As Integer, Verse As Integer, Word As Integer) As RenderArray
        Dim Lines As String() = MorphLines
        If _MorphDataToLineNumber Is Nothing Then
            _MorphDataToLineNumber = New Dictionary(Of Integer(), Integer)(New ByteArrayComparer())
            For Count As Integer = 0 To Lines.Length - 1
                If Lines(Count).Length <> 0 AndAlso Lines(Count).Chars(0) <> "#" Then
                    Dim Pieces As String() = Lines(Count).Split(CChar(vbTab))
                    If Pieces(0).Chars(0) = "(" Then
                        Dim Location As Integer() = New List(Of Integer)(Linq.Enumerable.Select(Pieces(0).TrimStart("("c).TrimEnd(")"c).Split(":"c), Function(Str As String) CInt(Str))).ToArray()
                        _MorphDataToLineNumber.Add(Location, Count)
                    End If
                End If
            Next
        End If
        Dim Renderer As New RenderArray(String.Empty)
        Dim LineTest As Integer = 1
        Do
            Dim Pieces As String() = Lines(_MorphDataToLineNumber(New Integer() {Chapter, Verse, Word, LineTest})).Split(CChar(vbTab))
            Dim Renderers As New List(Of RenderArray.RenderText)
            Renderers.Add(New RenderArray.RenderText(RenderArray.RenderDisplayClass.eArabic, Arb.TransliterateFromBuckwalter(Pieces(1))))
            Renderers.Add(New RenderArray.RenderText(RenderArray.RenderDisplayClass.eLTR, _PortableMethods.LoadResourceString("IslamInfo_" + GetPartOfSpeech(Pieces(2)).Id) + " ") With {.Clr = GetPartOfSpeech(Pieces(2)).Color})
            Dim Parts As String() = Pieces(3).Split("|"c)
            Dim bNotDefaultPresent As Boolean = False
            For Count As Integer = 0 To Parts.Length - 1
                Dim Types As String() = Parts(Count).Split(":"c)
                If Types(0) = "ROOT" Or Types(0) = "LEM" Or Types(0) = "SP" Then
                    Renderers.Add(New RenderArray.RenderText(RenderArray.RenderDisplayClass.eNested, New List(Of RenderArray.RenderItem) From {New RenderArray.RenderItem(RenderArray.RenderTypes.eText, New RenderArray.RenderText() {New RenderArray.RenderText(RenderArray.RenderDisplayClass.eLTR, _PortableMethods.LoadResourceString("IslamInfo_" + GetFeatureOfSpeech(Types(0)).Value.Id) + " "), New RenderArray.RenderText(RenderArray.RenderDisplayClass.eArabic, Arb.TransliterateFromBuckwalter(Types(1)))})}))
                ElseIf Types(0) = "POS" Then
                ElseIf Types(0) = "PRON" Then
                    Renderers.Add(New RenderArray.RenderText(RenderArray.RenderDisplayClass.eNested, New List(Of RenderArray.RenderItem) From {New RenderArray.RenderItem(RenderArray.RenderTypes.eText, New RenderArray.RenderText() {New RenderArray.RenderText(RenderArray.RenderDisplayClass.eLTR, _PortableMethods.LoadResourceString("IslamInfo_" + GetFeatureOfSpeech(Types(0)).Value.Id) + " "), New RenderArray.RenderText(RenderArray.RenderDisplayClass.eLTR, String.Join(" "c, Linq.Enumerable.Select(Types(1).ToCharArray(), Function(Ch) _PortableMethods.LoadResourceString("IslamInfo_" + GetFeatureOfSpeech(Ch).Value.Id))) + " ")})}))
                ElseIf Text.RegularExpressions.Regex.IsMatch(Parts(Count), "^[123]?(?:[MF]|[MF]?[SDP])$") Then
                    Renderers.Add(New RenderArray.RenderText(RenderArray.RenderDisplayClass.eLTR, String.Join(" "c, Linq.Enumerable.Select(Parts(Count).ToCharArray(), Function(Ch) _PortableMethods.LoadResourceString("IslamInfo_" + GetFeatureOfSpeech(Ch).Value.Id))) + " "))
                Else
                    'If Not GetFeatureOfSpeech(Parts(Count)).HasValue Then Debug.WriteLine("Not found: " + Parts(Count))
                    Renderers.Add(New RenderArray.RenderText(RenderArray.RenderDisplayClass.eLTR, _PortableMethods.LoadResourceString("IslamInfo_" + GetFeatureOfSpeech(Parts(Count)).Value.Id) + " "))
                End If
                If Parts(Count).StartsWith("MOOD:") Or Parts(Count) = "INDEF" Then bNotDefaultPresent = True
            Next
            If Pieces(2) = "V" And Not bNotDefaultPresent Then Renderers.Add(New RenderArray.RenderText(RenderArray.RenderDisplayClass.eLTR, _PortableMethods.LoadResourceString("IslamInfo_" + GetFeatureOfSpeech("MOOD:IND").Value.Id) + " "))
            If Pieces(2) = "N" And Not bNotDefaultPresent Then Renderers.Add(New RenderArray.RenderText(RenderArray.RenderDisplayClass.eLTR, _PortableMethods.LoadResourceString("IslamInfo_" + GetFeatureOfSpeech("DEF").Value.Id) + " "))
            Renderer.Items.Add(New RenderArray.RenderItem(RenderArray.RenderTypes.eText, Renderers.ToArray()))
            LineTest += 1
        Loop While _MorphDataToLineNumber.ContainsKey(New Integer() {Chapter, Verse, Word, LineTest})
        Return Renderer
        'Return Lines(_MorphDataToLineNumber(New Integer() {Chapter, Verse, Word, 1}))
    End Function
    Public Sub GetMorphologicalData()
        Dim Lines As String() = MorphLines
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
                    _LocDictionary.Add(Pieces(0).TrimStart("("c).TrimEnd(")"c), {Pieces(1), Linq.Enumerable.FirstOrDefault(Parts, Function(Str As String) Str = "PREFIX") <> String.Empty, Linq.Enumerable.FirstOrDefault(Parts, Function(Str As String) Str = "SUFFIX") <> String.Empty, Pieces(2)})
                    If Linq.Enumerable.FirstOrDefault(Parts, Function(Str As String) Str = "PREFIX" Or Str = "SUFFIX") = String.Empty Then
                        'LEM: or if not present FORM
                        Dim Lem As String = Linq.Enumerable.FirstOrDefault(Parts, Function(Str As String) Str.StartsWith("LEM:"))
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
                    If Linq.Enumerable.FirstOrDefault(Parts, Function(Str As String) Str = "PREFIX") <> String.Empty Then
                        If Not _PreDictionary.ContainsKey(Pieces(1)) Then
                            _PreDictionary.Add(Pieces(1), New List(Of Integer()))
                        End If
                        _PreDictionary.Item(Pieces(1)).Add(Location)
                    ElseIf Linq.Enumerable.FirstOrDefault(Parts, Function(Str As String) Str = "SUFFIX") <> String.Empty Then
                        If Not _SufDictionary.ContainsKey(Pieces(1)) Then
                            _SufDictionary.Add(Pieces(1), New List(Of Integer()))
                        End If
                        _SufDictionary.Item(Pieces(1)).Add(Location)
                    End If
                    'ROOT:
                    Dim Root As String = Linq.Enumerable.FirstOrDefault(Parts, Function(Str As String) Str.StartsWith("ROOT:"))
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
    Public Sub GetMorphDataByDivision(bStation As Boolean,
                                             ByRef All As Integer, ByRef AllUnique As Integer,
                                             ByRef PartArray() As Generic.List(Of String),
                                             ByRef PartUniqueArray() As Generic.List(Of String), TR As TanzilReader)
        If Not bStation Then
            ReDim PartUniqueArray(TR.GetPartCount() - 1)
            ReDim PartArray(TR.GetPartCount() - 1)
        Else
            ReDim PartUniqueArray(TR.GetPartCount() - 1)
            ReDim PartArray(TR.GetPartCount() - 1)
        End If
        Dim FreqArray(WordDictionary.Keys.Count - 1) As String
        WordDictionary.Keys.CopyTo(FreqArray, 0)
        For Count As Integer = 1 To CInt(If(bStation, TR.GetStationCount(), TR.GetPartCount()))
            PartUniqueArray(Count - 1) = New Generic.List(Of String)
            PartArray(Count - 1) = New Generic.List(Of String)
            Dim Node As Xml.Linq.XElement = CType(If(bStation, TR.GetStationByIndex(Count), TR.GetPartByIndex(Count)), Xml.Linq.XElement)
            Dim BaseChapter As Integer = CInt(Node.Attribute("sura").Value)
            Dim BaseVerse As Integer = CInt(Node.Attribute("aya").Value)
            Dim Chapter As Integer
            Dim Verse As Integer
            Node = CType(If(bStation, TR.GetStationByIndex(Count + 1), TR.GetPartByIndex(Count + 1)), Xml.Linq.XElement)
            If Node Is Nothing Then
                Chapter = TR.GetChapterCount()
                Verse = TR.GetVerseCount(Chapter)
            Else
                Chapter = CInt(Node.Attribute("sura").Value)
                Verse = CInt(Node.Attribute("aya").Value)
                TR.GetPreviousChapterVerse(Chapter, Verse)
            End If
            For SubCount As Integer = 0 To FreqArray.Length - 1
                Dim RefCount As Integer
                Dim UniCount As Integer = 0
                For RefCount = 0 To WordDictionary(FreqArray(SubCount)).Count - 1
                    For FormCount As Integer = 0 To FormDictionary(CStr(WordDictionary(FreqArray(SubCount))(RefCount))).Count - 1
                        If (FormDictionary(CStr(WordDictionary(FreqArray(SubCount))(RefCount)))(FormCount)(0) = BaseChapter AndAlso
                            FormDictionary(CStr(WordDictionary(FreqArray(SubCount))(RefCount)))(FormCount)(1) >= BaseVerse AndAlso
                            (BaseChapter <> Chapter OrElse
                            FormDictionary(CStr(WordDictionary(FreqArray(SubCount))(RefCount)))(FormCount)(1) <= Verse)) OrElse
                            (FormDictionary(CStr(WordDictionary(FreqArray(SubCount))(RefCount)))(FormCount)(0) > BaseChapter AndAlso
                            FormDictionary(CStr(WordDictionary(FreqArray(SubCount))(RefCount)))(FormCount)(0) <Chapter) OrElse
                            (FormDictionary(CStr(WordDictionary(FreqArray(SubCount))(RefCount)))(FormCount)(0) = Chapter AndAlso
                            FormDictionary(CStr(WordDictionary(FreqArray(SubCount))(RefCount)))(FormCount)(1) <= Verse) Then
                            UniCount += 1
                        End If
                    Next
                Next
                If UniCount = WordDictionary(FreqArray(SubCount)).Count Then
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
    Public Sub BuildQuranLetterIndex(TR As TanzilReader)
        'Bismillah are not counted in here
        Dim Verses As Collections.Generic.List(Of String())
        Verses = TR.GetQuranText(XMLDocMain, -1, -1, -1, -1)
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
                    If LetCount <> 0 AndAlso LetCount <> Verses(Count)(SubCount).Length - 1 AndAlso
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
    Public Async Function DoErrorCheck(TR As TanzilReader, DB As DocBuilder) As Threading.Tasks.Task
        'missing from loader
        'loanwords word value/id, jurisprudence name/id, category name/id, months name/id
        'daysofweek name/id, fasting fastingtype name/id, islamicbooks book title/author
        Dim Count As Integer
        Dim Verses As Collections.Generic.List(Of String())
        Verses = TR.GetQuranText(XMLDocMain, -1, -1, -1, -1)
        For Count = 0 To Verses.Count - 1
            Dim ChapterNode As Xml.Linq.XElement = TanzilReader.GetTextChapter(XMLDocMain, Count + 1)
            For SubCount As Integer = 0 To Verses(Count).Length - 1
                If SubCount = 0 AndAlso Not TanzilReader.GetTextVerse(ChapterNode, SubCount + 1).Attribute("bismillah") Is Nothing Then
                    Arb.DoErrorCheck(TanzilReader.GetTextVerse(ChapterNode, SubCount + 1).Attribute("bismillah").Value)
                End If
                Arb.DoErrorCheck(Verses(Count)(SubCount))
            Next
        Next
        For Count = 0 To IslamData.Lists.Length - 1
            _PortableMethods.LoadResourceString("IslamInfo_" + IslamData.Lists(Count).Title)
            If Not IslamData.Lists(Count).Words Is Nothing Then
                For SubCount As Integer = 0 To IslamData.Lists(Count).Words.Length - 1
                    Await DB.DoErrorCheckBuckwalterText(IslamData.Lists(Count).Words(SubCount).Text, String.Empty, TR)
                    _PortableMethods.LoadResourceString("IslamInfo_" + IslamData.Lists(Count).Words(SubCount).TranslationID)
                Next
            End If
        Next
        For Count = 0 To IslamData.Grammar.Transforms.Length - 1
            Arb.DoErrorCheck(Arb.TransliterateFromBuckwalter(IslamData.Grammar.Transforms(Count).Text))
            _PortableMethods.LoadResourceString("IslamInfo_" + IslamData.Grammar.Transforms(Count).TranslationID)
        Next
        For Count = 0 To IslamData.Grammar.Nouns.Length - 1
            Arb.DoErrorCheck(Arb.TransliterateFromBuckwalter(IslamData.Grammar.Nouns(Count).Text))
            _PortableMethods.LoadResourceString("IslamInfo_" + IslamData.Grammar.Nouns(Count).TranslationID)
        Next
        For Count = 0 To IslamData.Grammar.Particles.Length - 1
            Arb.DoErrorCheck(Arb.TransliterateFromBuckwalter(IslamData.Grammar.Particles(Count).Text))
            _PortableMethods.LoadResourceString("IslamInfo_" + IslamData.Grammar.Particles(Count).TranslationID)
        Next
        For Count = 0 To IslamData.Grammar.Verbs.Length - 1
            Arb.DoErrorCheck(Arb.TransliterateFromBuckwalter(IslamData.Grammar.Verbs(Count).Text))
            _PortableMethods.LoadResourceString("IslamInfo_" + IslamData.Grammar.Verbs(Count).TranslationID)
        Next
        For Count = 0 To IslamData.Phrases.Length - 1
            Await DB.DoErrorCheckBuckwalterText(Arb.TransliterateFromBuckwalter(IslamData.Phrases(Count).Text), IslamData.Phrases(Count).TranslationID, TR)
        Next
        For Count = 0 To IslamData.Abbreviations.Length - 1
            If Not Phrases.GetPhraseCat(IslamData.Abbreviations(Count).TranslationID, IslamData.Phrases).HasValue Then
                'Debug.Print("Missing Phrase ID: " + IslamData.Abbreviations(Count).TranslationID)
            End If
        Next
        For Count = 0 To ArbData.ArabicLetters.Length - 1
            Arb.DoErrorCheck(Arb.ArabicLetterSpelling(ArbData.ArabicLetters(Count).Symbol, False, False, False))
            If Not ArbData.ArabicLetters(Count).UnicodeName.StartsWith("<") Then
                _PortableMethods.LoadResourceString("IslamInfo_" + ArbData.ToCamelCase(ArbData.ArabicLetters(Count).UnicodeName))
            End If
        Next
        For Count = 0 To IslamData.Collections.Length - 1
            _PortableMethods.LoadResourceString("IslamInfo_" + IslamData.Collections(Count).Name)
        Next
        For Count = 0 To IslamData.QuranDivisions.Length - 1
            _PortableMethods.LoadResourceString("IslamInfo_" + IslamData.QuranDivisions(Count).Description)
        Next
        For Count = 0 To IslamData.QuranSelections.Length - 1
            _PortableMethods.LoadResourceString("IslamInfo_" + IslamData.QuranSelections(Count).Description)
        Next
        For Count = 0 To IslamData.QuranChapters.Length - 1
            Arb.DoErrorCheck(Arb.TransliterateFromBuckwalter(IslamData.QuranChapters(Count).Name))
        Next
        For Count = 0 To IslamData.QuranParts.Length - 1
            Arb.DoErrorCheck(Arb.TransliterateFromBuckwalter(IslamData.QuranParts(Count).Name))
            _PortableMethods.LoadResourceString("IslamInfo_" + IslamData.QuranParts(Count).ID)
        Next
        For Count = 0 To IslamData.PartsOfSpeech.Length - 1
            _PortableMethods.LoadResourceString("IslamInfo_" + IslamData.PartsOfSpeech(Count).Id)
        Next
    End Function
    Public ReadOnly Property IslamData() As IslamData
        Get
            Return _ObjIslamData
        End Get
    End Property
    Public ReadOnly Property XMLDocMain() As Xml.Linq.XDocument
        Get
            Return _XMLDocMain
        End Get
    End Property
    Public ReadOnly Property XMLDocInfo() As Xml.Linq.XDocument
        Get
            Return _XMLDocInfo
        End Get
    End Property
    Public ReadOnly Property XMLDocInfos() As Collections.Generic.List(Of Xml.Linq.XDocument)
        Get
            Return _XMLDocInfos
        End Get
    End Property
    Public ReadOnly Property RootDictionary As Generic.Dictionary(Of String, List(Of Integer()))
        Get
            'If _RootDictionary.Keys.Count = 0 Then GetMorphologicalData()
            Return _RootDictionary
        End Get
    End Property
    Public ReadOnly Property FormDictionary As Generic.Dictionary(Of String, List(Of Integer()))
        Get
            'If _FormDictionary.Keys.Count = 0 Then GetMorphologicalData()
            Return _FormDictionary
        End Get
    End Property
    Public ReadOnly Property TagDictionary As Generic.Dictionary(Of String, Generic.Dictionary(Of String, List(Of Integer())))
        Get
            'If _TagDictionary.Keys.Count = 0 Then GetMorphologicalData()
            Return _TagDictionary
        End Get
    End Property
    Public ReadOnly Property WordDictionary As Generic.Dictionary(Of String, List(Of String))
        Get
            'If _WordDictionary.Keys.Count = 0 Then GetMorphologicalData()
            Return _WordDictionary
        End Get
    End Property
    Public ReadOnly Property RealWordDictionary As Generic.Dictionary(Of String, List(Of Integer()))
        Get
            'If _RealWordDictionary.Keys.Count = 0 Then BuildQuranLetterIndex()
            Return _RealWordDictionary
        End Get
    End Property
    Public ReadOnly Property LetterDictionary As Generic.Dictionary(Of Char, Generic.Dictionary(Of String, List(Of Integer())))
        Get
            'If _LetterDictionary.Keys.Count = 0 Then BuildQuranLetterIndex()
            Return _LetterDictionary
        End Get
    End Property
    Public ReadOnly Property LetterPreDictionary As Generic.Dictionary(Of Char, Generic.Dictionary(Of String, List(Of Integer())))
        Get
            'If _LetterPreDictionary.Keys.Count = 0 Then BuildQuranLetterIndex()
            Return _LetterPreDictionary
        End Get
    End Property
    Public ReadOnly Property LetterSufDictionary As Generic.Dictionary(Of Char, Generic.Dictionary(Of String, List(Of Integer())))
        Get
            'If _LetterSufDictionary.Keys.Count = 0 Then BuildQuranLetterIndex()
            Return _LetterSufDictionary
        End Get
    End Property
    Public ReadOnly Property PreDictionary As Generic.Dictionary(Of String, List(Of Integer()))
        Get
            'If _PreDictionary.Keys.Count = 0 Then GetMorphologicalData()
            Return _PreDictionary
        End Get
    End Property
    Public ReadOnly Property SufDictionary As Generic.Dictionary(Of String, List(Of Integer()))
        Get
            'If _SufDictionary.Keys.Count = 0 Then GetMorphologicalData()
            Return _SufDictionary
        End Get
    End Property
    Public ReadOnly Property LocDictionary As Generic.Dictionary(Of String, Object())
        Get
            'If _LocDictionary.Keys.Count = 0 Then GetMorphologicalData()
            Return _LocDictionary
        End Get
    End Property
    Public ReadOnly Property TotalLetters As Integer
        Get
            'If _TotalLetters = 0 Then BuildQuranLetterIndex()
            Return _TotalLetters
        End Get
    End Property
    Public ReadOnly Property IsolatedLetterDictionary As Generic.Dictionary(Of Char, List(Of Integer()))
        Get
            'If _IsolatedLetterDictionary.Keys.Count = 0 Then BuildQuranLetterIndex()
            Return _IsolatedLetterDictionary
        End Get
    End Property
    Public ReadOnly Property TotalIsolatedLetters As Integer
        Get
            'If _TotalIsolatedLetters = 0 Then BuildQuranLetterIndex()
            Return _TotalIsolatedLetters
        End Get
    End Property
    Public ReadOnly Property PartArray As Generic.List(Of String)()
        Get
            'If TotalWordsInParts = 0 Then GetMorphDataByDivision(False, _TotalWordsInParts, _TotalUniqueWordsInParts, _PartArray, _PartUniqueArray)
            Return _PartArray
        End Get
    End Property
    Public ReadOnly Property PartUniqueArray As Generic.List(Of String)()
        Get
            'If TotalUniqueWordsInParts = 0 Then GetMorphDataByDivision(False, _TotalWordsInParts, _TotalUniqueWordsInParts, _PartArray, _PartUniqueArray)
            Return _PartUniqueArray
        End Get
    End Property
    Public ReadOnly Property TotalWordsInParts As Integer
        Get
            'If TotalWordsInParts = 0 Then GetMorphDataByDivision(False, _TotalWordsInParts, _TotalUniqueWordsInParts, _PartArray, _PartUniqueArray)
            Return _TotalWordsInParts
        End Get
    End Property
    Public ReadOnly Property TotalUniqueWordsInParts As Integer
        Get
            'If TotalUniqueWordsInParts = 0 Then GetMorphDataByDivision(False, _TotalWordsInParts, _TotalUniqueWordsInParts, _PartArray, _PartUniqueArray)
            Return _TotalUniqueWordsInParts
        End Get
    End Property
    Public ReadOnly Property StationArray As Generic.List(Of String)()
        Get
            'If _TotalWordsInStations = 0 Then GetMorphDataByDivision(True, _TotalWordsInStations, _TotalUniqueWordsInStations, _StationArray, _StationUniqueArray)
            Return _StationArray
        End Get
    End Property
    Public ReadOnly Property StationUniqueArray As Generic.List(Of String)()
        Get
            'If TotalUniqueWordsInStations = 0 Then GetMorphDataByDivision(True, _TotalWordsInStations, _TotalUniqueWordsInStations, _StationArray, _StationUniqueArray)
            Return _StationUniqueArray
        End Get
    End Property
    Public ReadOnly Property TotalWordsInStations As Integer
        Get
            'If _TotalWordsInStations = 0 Then GetMorphDataByDivision(True, _TotalWordsInStations, _TotalUniqueWordsInStations, _StationArray, _StationUniqueArray)
            Return _TotalWordsInStations
        End Get
    End Property
    Public ReadOnly Property TotalUniqueWordsInStations As Integer
        Get
            'If TotalUniqueWordsInStations = 0 Then GetMorphDataByDivision(True, _TotalWordsInStations, _TotalUniqueWordsInStations, _StationArray, _StationUniqueArray)
            Return _TotalUniqueWordsInStations
        End Get
    End Property
End Class
Public Class Languages
    Public Shared Function GetLanguageInfoByCode(ByVal Code As String, Languages As IslamData.LanguageInfo()) As IslamData.LanguageInfo
        Dim Count As Integer
        For Count = 0 To Languages.Length - 1
            If Languages(Count).Code = Code Then Return Languages(Count)
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
        Return New XYZColor With {.X = r * 0.4124 + g * 0.3576 + b * 0.1805,
                                  .Y = r * 0.2126 + g * 0.7152 + b * 0.0722,
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
        Return XYZToRGB(New XYZColor With {.X = WhiteReference.X * If(X3 > Epsilon, X3, (x - 16.0 / 116.0) / 7.787),
                                         .Y = WhiteReference.Y * If(clr.L > (Kappa * Epsilon), Math.Pow((clr.L + 16.0) / 116.0, 3), clr.L / Kappa),
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
    Private _PortableMethods As PortableMethods
    Private Arb As Arabic
    Private ArbData As ArabicData
    Private ChData As CachedData
    Public Sub New(NewPortableMethods As PortableMethods, NewArb As Arabic, NewArbData As ArabicData, NewChData As CachedData)
        _PortableMethods = NewPortableMethods
        Arb = NewArb
        ArbData = NewArbData
        ChData = NewChData
    End Sub
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
                Debug.Assert(NumStack.Count <> 0) 'Imbalance parenthesis is exception
                Dim Nums As Integer() = NumStack.Pop()
                For Count As Integer = Nums(1) To Matches(MatchCount).Groups(2).Index
                    If Nums(0) > ParenPos(Count) Then ParenPos(Count) = Nums(0)
                Next
            End If
        Next
        Debug.Assert(NumStack.Count = 0) 'Imbalance parenthesis is exception
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
    Public Shared W4WItems As Dictionary(Of String, Dictionary(Of String, String))
    Public Async Function GetW4WItem(ID As String) As Threading.Tasks.Task(Of String)
        Dim LangID As String = Threading.Thread.CurrentThread.CurrentUICulture.TwoLetterISOLanguageName
        If W4WItems Is Nothing Then
            W4WItems = New Dictionary(Of String, Dictionary(Of String, String))
        End If
        If Not W4WItems.ContainsKey(LangID) Then
            If Not Await _PortableMethods.FileIO.PathExists(_PortableMethods.Settings.GetFilePath(_PortableMethods.FileIO.CombinePath("metadata", LangID + ".w4w.txt"))) And W4WItems.ContainsKey(ChData.IslamData.LanguageDefaultInfo.DefaultLanguage) Then
                W4WItems.Add(LangID, W4WItems(ChData.IslamData.LanguageDefaultInfo.DefaultLanguage))
            Else
                If Not Await _PortableMethods.FileIO.PathExists(_PortableMethods.Settings.GetFilePath(_PortableMethods.FileIO.CombinePath("metadata", LangID + ".w4w.txt"))) Then LangID = ChData.IslamData.LanguageDefaultInfo.DefaultLanguage
                W4WItems.Add(LangID, New Dictionary(Of String, String))
                For Each Line As String In Await _PortableMethods.ReadAllLines(_PortableMethods.Settings.GetFilePath(_PortableMethods.FileIO.CombinePath("metadata", LangID + ".w4w.txt")))
                    W4WItems(LangID).Add(Line.Substring(0, Line.IndexOf("="c)), Line.Substring(Line.IndexOf("="c) + 1))
                Next
            End If
        End If
        Return W4WItems(LangID)(ID)
    End Function
    Public Async Function DoErrorCheckBuckwalterText(Strings As String, TranslationID As String, TR As TanzilReader) As Threading.Tasks.Task
        If Strings = Nothing Then Return
        Dim Matches As System.Text.RegularExpressions.MatchCollection = System.Text.RegularExpressions.Regex.Matches(Strings, "(.*?)(?:(\\\{)(.*?)(\\\})|$)", System.Text.RegularExpressions.RegexOptions.Singleline)
        Dim EnglishByWord As String() = If(TranslationID = Nothing, {}, (Await GetW4WItem(TranslationID)).Split("|"c))
        For MatchCount As Integer = 0 To Matches.Count - 1
            If Matches(MatchCount).Length <> 0 Then
                If Matches(MatchCount).Groups(1).Length <> 0 Then
                    Dim ArabicText As String() = Matches(MatchCount).Groups(1).Value.Split(" "c)
                    If ArabicText.Length > 1 And EnglishByWord.Length = ArabicText.Length Then
                    End If
                    Arb.DoErrorCheck(Arb.TransliterateFromBuckwalter(Matches(MatchCount).Groups(1).Value))
                    _PortableMethods.LoadResourceString("IslamInfo_" + TranslationID)
                End If
                If Matches(MatchCount).Groups(3).Length <> 0 Then
                    ErrorCheckTextFromReferences(Matches(MatchCount).Groups(3).Value, TR)
                End If
            End If
        Next
    End Function
    Shared _Abbrevs As Dictionary(Of String, IslamData.AbbrevWord)
    Public ReadOnly Property Abbrevs As Dictionary(Of String, IslamData.AbbrevWord)
        Get
            If _Abbrevs Is Nothing Then
                _Abbrevs = New Dictionary(Of String, IslamData.AbbrevWord)
                For Count = 0 To ChData.IslamData.Abbreviations.Length - 1
                    Dim AbbrevWord As IslamData.AbbrevWord = ChData.IslamData.Abbreviations(Count)
                    If ChData.IslamData.Abbreviations(Count).Text <> String.Empty Then
                        For Each Str As String In ChData.IslamData.Abbreviations(Count).Text.Split("|"c)
                            _Abbrevs.Add(Str, AbbrevWord)
                        Next
                    End If
                    If ChData.IslamData.Abbreviations(Count).TranslationID <> String.Empty AndAlso Array.IndexOf(ChData.IslamData.Abbreviations(Count).Text.Split("|"c), ChData.IslamData.Abbreviations(Count).TranslationID) = -1 Then _Abbrevs.Add(ChData.IslamData.Abbreviations(Count).TranslationID, ChData.IslamData.Abbreviations(Count))
                Next
            End If
            Return _Abbrevs
        End Get
    End Property
    Public Sub ErrorCheckTextFromReferences(Strings As String, TR As TanzilReader)
        Dim _Options As String() = Strings.Split(";"c)
        Strings = _Options(0)
        Dim Count As Integer = 0
        While _Options(Count).EndsWith("&leftbrace") Or _Options(Count).EndsWith("&rightbrace") Or _Options(Count).EndsWith("&comma") Or _Options(Count).EndsWith("&semicolon")
            _Options(0) = _Options(0) + ";" + _Options(Count + 1)
            Count += 1
        End While
        If TR.IsQuranTextReference(Strings) Then
        ElseIf Strings.StartsWith("symbol:") Then
            Dim SelArr As String() = Strings.Replace("symbol:", String.Empty).Split(","c)
            For SubCount = 0 To ArbData.ArabicLetters.Length - 1
                If Array.IndexOf(SelArr, ArbData.ToCamelCase(ArbData.ArabicLetters(SubCount).UnicodeName).Replace("ArabicLetter", String.Empty).Replace("Arabic", String.Empty)) <> -1 Then
                End If
            Next
        ElseIf Strings.StartsWith("personalpronoun:") Or Strings.StartsWith("proximaldemonstratives:") Or Strings.StartsWith("distaldemonstratives:") Then
            Dim Words As IslamData.GrammarSet.GrammarNoun()
            Dim SelArr As String()
            If Strings.StartsWith("proximaldemonstratives") Then
                Words = Arb.GetCatNoun("proxdemo")
                SelArr = Strings.Replace("proximaldemonstratives:", String.Empty).Split(","c)
            ElseIf Strings.StartsWith("distaldemonstratives") Then
                Words = Arb.GetCatNoun("distdemo")
                SelArr = Strings.Replace("distaldemonstratives:", String.Empty).Split(","c)
            ElseIf Strings.StartsWith("personalpronoun:") Then
                Words = Arb.GetCatNoun("perspro")
                SelArr = Strings.Replace("personalpronoun:", String.Empty).Split(","c)
            Else
                Words = Nothing
                SelArr = Nothing
            End If
            For Count = 0 To SelArr.Length - 1
                Dim S As String = SelArr(Count)
                'If Words Is Nothing OrElse New List(Of IslamData.GrammarSet.GrammarNoun)(Linq.Enumerable.TakeWhile(Words, Function(Word As IslamData.GrammarSet.GrammarNoun) S <> Word.TranslationID)).Count() = Words.Length Then Debug.Print("Noun Subject ID Not Found: " + SelArr(Count))
            Next
        ElseIf Strings.StartsWith("plurals:") Or Strings.StartsWith("possessivedeterminerpersonalpronoun:") Then
            Dim Words As IslamData.GrammarSet.GrammarTransform()
            Dim SelArr As String()
            If Strings.StartsWith("plurals") Then
                Words = Arb.GetTransform("mp|fp|bp")
                SelArr = Strings.Replace("plurals:", String.Empty).Split(","c)
            ElseIf Strings.StartsWith("possessivedeterminerpersonalpronoun") Then
                Words = Arb.GetTransform("posspron")
                SelArr = Strings.Replace("possessivedeterminerpersonalpronoun:", String.Empty).Split(","c)
            Else
                Words = Nothing
                SelArr = Nothing
            End If
            For Count = 0 To SelArr.Length - 1
                Dim S As String = SelArr(Count)
                'If Words Is Nothing OrElse New List(Of IslamData.GrammarSet.GrammarTransform)(Linq.Enumerable.TakeWhile(Words, Function(Word As IslamData.GrammarSet.GrammarTransform) S <> Word.TranslationID)).Count() = Words.Length Then Debug.Print("Transform Subject ID Not Found: " + SelArr(Count))
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
    Public Function GetListCategories() As String()
        Return New List(Of String)(Linq.Enumerable.Select(ChData.IslamData.Lists, Function(Convert As IslamData.ListCategory) _PortableMethods.LoadResourceString("IslamInfo_" + Convert.Title))).ToArray()
    End Function
    Public Function GetListCat(ID As String) As Nullable(Of IslamData.ListCategory.Word)
        GetListCat = New Nullable(Of IslamData.ListCategory.Word)
        If ListIDs.ContainsKey(ID) Then GetListCat = ListIDs(ID)
    End Function
    Public Shared _ListIDs As Dictionary(Of String, IslamData.ListCategory.Word)
    Public Shared _ListTitles As Dictionary(Of String, IslamData.ListCategory)
    Public Function ListIDs() As Dictionary(Of String, IslamData.ListCategory.Word)
        If _ListIDs Is Nothing Then
            _ListIDs = New Dictionary(Of String, IslamData.ListCategory.Word)
            For Count As Integer = 0 To ChData.IslamData.Lists.Length - 1
                If Not ChData.IslamData.Lists(Count).Words Is Nothing Then
                    For SubCount As Integer = 0 To ChData.IslamData.Lists(Count).Words.Length - 1
                        _ListIDs.Add(ChData.IslamData.Lists(Count).Words(SubCount).TranslationID, ChData.IslamData.Lists(Count).Words(SubCount))
                    Next
                End If
            Next
        End If
        Return _ListIDs
    End Function
    Public Function ListTitles() As Dictionary(Of String, IslamData.ListCategory)
        If _ListTitles Is Nothing Then
            _ListTitles = New Dictionary(Of String, IslamData.ListCategory)
            For Count As Integer = 0 To ChData.IslamData.Lists.Length - 1
                _ListTitles.Add(ChData.IslamData.Lists(Count).Title, ChData.IslamData.Lists(Count))
            Next
        End If
        Return _ListTitles
    End Function
    Public Function GetListCats(SelArr As String()) As IslamData.ListCategory.Word()
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
    Public Shared Function GetPhraseCat(ID As String, Phrases As IslamData.Phrase()) As Nullable(Of IslamData.Phrase)
        If PhraseIDs(Phrases).ContainsKey(ID) Then Return PhraseIDs(Phrases)(ID)
        Return New Nullable(Of IslamData.Phrase)
    End Function
    Public Shared _PhraseIDs As Dictionary(Of String, IslamData.Phrase)
    Public Shared Function PhraseIDs(Phrases As IslamData.Phrase()) As Dictionary(Of String, IslamData.Phrase)
        If _PhraseIDs Is Nothing Then
            _PhraseIDs = New Dictionary(Of String, IslamData.Phrase)
            For Count As Integer = 0 To Phrases.Length - 1
                _PhraseIDs.Add(Phrases(Count).TranslationID, Phrases(Count))
            Next
        End If
        Return _PhraseIDs
    End Function
    Public Shared Function GetPhraseCats(SelArr As String(), Phrases As IslamData.Phrase()) As IslamData.Phrase()
        Dim PhraseCats As New List(Of IslamData.Phrase)
        For SelCount As Integer = 0 To SelArr.Length - 1
            Dim Word As Nullable(Of IslamData.Phrase) = GetPhraseCat(SelArr(SelCount), Phrases)
            If Word.HasValue Then PhraseCats.Add(Word.Value)
        Next
        Return PhraseCats.ToArray()
    End Function
End Class
Public Class TanzilReader
    Private _PortableMethods As PortableMethods
    Private Arb As Arabic
    Private ArbData As ArabicData
    Private ChData As CachedData
    Dim _IndexToVerse As Integer()()
    Dim _QuranText As String
    Dim _CacheMetarules As List(Of Arabic.RuleMetadata)
    Public Sub New(NewPortableMethods As PortableMethods, NewArb As Arabic, NewArbData As ArabicData, NewChData As CachedData)
        _PortableMethods = NewPortableMethods
        Arb = NewArb
        ArbData = NewArbData
        ChData = NewChData
    End Sub
    Public Async Function Init() As Threading.Tasks.Task
        _QuranText = QuranTextCombiner(ChData.XMLDocMain, _IndexToVerse)
        _CacheMetarules = Await GetQuranCacheMetarules()
    End Function
    Public Async Function MakeQuranCacheMetarules() As Threading.Tasks.Task(Of List(Of Arabic.RuleMetadata))
        Dim Rules As List(Of Arabic.RuleMetadata) = Arb.GetMetarules(_QuranText, ChData.RuleMetas("UthmaniQuran"))
        Await _PortableMethods.WriteAllLines(_PortableMethods.FileIO.CombinePath(Await _PortableMethods.DiskCache.GetCacheDirectory(), "QuranTajweedData.txt"), Arabic.MakeCacheMetarules(Rules, _IndexToVerse))
        Return Rules
    End Function
    Public Async Function GetQuranCacheMetarules() As Threading.Tasks.Task(Of List(Of Arabic.RuleMetadata))
        Dim Path As String = _PortableMethods.FileIO.CombinePath("metadata", "QuranTajweedData.txt")
        If Not Await _PortableMethods.FileIO.PathExists(Path) Then
            Path = _PortableMethods.FileIO.CombinePath(Await _PortableMethods.DiskCache.GetCacheDirectory(), "QuranTajweedData.txt")
            If Not Await _PortableMethods.FileIO.PathExists(Path) Then Return Await MakeQuranCacheMetarules()
        End If
        Return Arb.GetCacheMetarules(Await _PortableMethods.ReadAllLines(Path), _IndexToVerse)
    End Function
    Public Function GetDivisionTypes() As String()
        Return New List(Of String)(Linq.Enumerable.Select(ChData.IslamData.QuranDivisions, Function(Convert As IslamData.QuranDivision) _PortableMethods.LoadResourceString("IslamInfo_" + Convert.Description))).ToArray()
    End Function
    Public Function GetTranslationList() As Array()
        Return New List(Of String())(Linq.Enumerable.Select(ChData.IslamData.Translations.TranslationList, Function(Convert As IslamData.TranslationsInfo.TranslationInfo) New String() {_PortableMethods.LoadResourceString("lang_local" + Languages.GetLanguageInfoByCode(Convert.FileName.Substring(0, CInt(If(Convert.FileName.IndexOf("-") <> -1, Convert.FileName.IndexOf("-"), Convert.FileName.IndexOf(".")))), ChData.IslamData.LanguageList).Code) + ": " + Convert.Name, Convert.FileName})).ToArray()
    End Function
    Public Function GetTranslationIndex(ByVal Translation As String) As Integer
        If Translation = "None" Then Return -1
        If String.IsNullOrEmpty(Translation) Then Translation = ChData.IslamData.LanguageDefaultInfo.GetLanguageByID(String.Empty).QuranFile 'Default
        Dim Count As Integer = New List(Of IslamData.TranslationsInfo.TranslationInfo)(Linq.Enumerable.TakeWhile(ChData.IslamData.Translations.TranslationList, Function(Test As IslamData.TranslationsInfo.TranslationInfo) Test.FileName <> Translation)).Count
        If Count = ChData.IslamData.Translations.TranslationList.Length Then
            Translation = ChData.IslamData.LanguageDefaultInfo.GetLanguageByID(String.Empty).QuranFile 'Default
            Count = New List(Of IslamData.TranslationsInfo.TranslationInfo)(Linq.Enumerable.TakeWhile(ChData.IslamData.Translations.TranslationList, Function(Test As IslamData.TranslationsInfo.TranslationInfo) Test.FileName <> Translation)).Count
        End If
        Return Count
    End Function
    Public Function GetTranslationFileName(ByVal Translation As String) As String
        Dim Index As Integer = GetTranslationIndex(Translation)
        Return ChData.IslamData.Translations.TranslationList(Index).FileName + ".txt"
    End Function

    Public Function GetWordPartitions() As String()
        Dim Parts As New Generic.List(Of String) From {_PortableMethods.LoadResourceString("IslamInfo_Letters"), _PortableMethods.LoadResourceString("IslamInfo_Words"), _PortableMethods.LoadResourceString("IslamInfo_UniqueWords"), _PortableMethods.LoadResourceString("IslamInfo_UniqueWordsPerPart"), _PortableMethods.LoadResourceString("IslamInfo_WordsPerPart"), _PortableMethods.LoadResourceString("IslamInfo_UniqueWordsPerStation"), _PortableMethods.LoadResourceString("IslamInfo_WordsPerStation"), _PortableMethods.LoadResourceString("IslamInfo_IsolatedLetters"), _PortableMethods.LoadResourceString("IslamInfo_LetterPatterns"), _PortableMethods.LoadResourceString("IslamInfo_LetterPatterns"), _PortableMethods.LoadResourceString("IslamInfo_LetterPatterns"), _PortableMethods.LoadResourceString("IslamInfo_Prefix"), _PortableMethods.LoadResourceString("IslamInfo_Suffix")}
        Parts.AddRange(Linq.Enumerable.Select(ChData.IslamData.PartsOfSpeech, Function(POS As IslamData.PartOfSpeechInfo) _PortableMethods.LoadResourceString("IslamInfo_" + POS.Id)))
        Parts.AddRange(Linq.Enumerable.Select(ChData.RecitationSymbols, Function(Sym As String) ArbData.GetUnicodeName(Sym.Chars(0))))
        Parts.AddRange(Linq.Enumerable.Select(ChData.RecitationSymbols, Function(Sym As String) "Prefix of " + ArbData.GetUnicodeName(Sym.Chars(0))))
        Parts.AddRange(Linq.Enumerable.Select(ChData.RecitationSymbols, Function(Sym As String) "Suffix of " + ArbData.GetUnicodeName(Sym.Chars(0))))
        Return Parts.ToArray()
    End Function
    Public Function GetQuranWordTotalNumber() As Integer
        Dim Total As Integer
        For Each Key As String In ChData.WordDictionary.Keys
            Total = Total + ChData.WordDictionary.Item(Key).Count
        Next
        Return Total
    End Function
    Public Function GetQuranWordTotal(Strings As String) As String
        Dim Index As Integer
        If Not Strings Is Nothing Then Index = CInt(Strings)
        If Index = 0 Then
            Return CStr(ChData.TotalLetters)
        ElseIf Index = 7 Then
            Return CStr(ChData.TotalIsolatedLetters)
        ElseIf Index = 1 Then
            Return CStr(GetQuranWordTotalNumber())
        ElseIf Index = 2 Then
            Return CStr(ChData.WordDictionary.Keys.Count)
        ElseIf Index = 3 Then
            Return CStr(ChData.TotalUniqueWordsInParts)
        ElseIf Index = 4 Then
            Return CStr(ChData.TotalWordsInParts)
        ElseIf Index = 5 Then
            Return CStr(ChData.TotalUniqueWordsInStations)
        ElseIf Index = 6 Then
            Return CStr(ChData.TotalWordsInStations)
        ElseIf Index = 8 Then
            Return String.Empty
        ElseIf Index = 9 Then
            Return String.Empty
        ElseIf Index = 10 Then
            Return String.Empty
        ElseIf Index = 11 Then
            Return CStr(ChData.PreDictionary.Count)
        ElseIf Index = 12 Then
            Return CStr(ChData.SufDictionary.Count)
        ElseIf Index >= 13 And Index < 13 + ChData.IslamData.PartsOfSpeech.Length Then
            Return CStr(ChData.TagDictionary.Item(ChData.IslamData.PartsOfSpeech(Index - 13).Symbol).Count)
        ElseIf Index >= 13 + ChData.IslamData.PartsOfSpeech.Length And Index < 13 + ChData.IslamData.PartsOfSpeech.Length + ChData.RecitationSymbols.Length Then
            Return CStr(ChData.LetterDictionary.Item(ChData.RecitationSymbols(Index - 13 - ChData.IslamData.PartsOfSpeech.Length).Chars(0)).Count)
        ElseIf Index >= 13 + ChData.IslamData.PartsOfSpeech.Length + ChData.RecitationSymbols.Length And Index < 13 + ChData.IslamData.PartsOfSpeech.Length + ChData.RecitationSymbols.Length + ChData.RecitationSymbols.Length Then
            Return CStr(ChData.LetterPreDictionary.Item(ChData.RecitationSymbols(Index - 13 - ChData.IslamData.PartsOfSpeech.Length - ChData.RecitationSymbols.Length).Chars(0)).Count)
        ElseIf Index >= 13 + ChData.IslamData.PartsOfSpeech.Length + ChData.RecitationSymbols.Length + ChData.RecitationSymbols.Length And Index < 13 + ChData.IslamData.PartsOfSpeech.Length + ChData.RecitationSymbols.Length + ChData.RecitationSymbols.Length + ChData.RecitationSymbols.Length Then
            Return CStr(ChData.LetterSufDictionary.Item(ChData.RecitationSymbols(Index - 13 - ChData.IslamData.PartsOfSpeech.Length - ChData.RecitationSymbols.Length - ChData.RecitationSymbols.Length).Chars(0)).Count)
        Else
            Return String.Empty
        End If
    End Function
    Public Async Function GetQuranWordFrequency(SchemeType As ArabicData.TranslitScheme, Scheme As String, Strings As String) As Threading.Tasks.Task(Of Object())
        Dim Output As New List(Of Object())
        Dim Total As Integer = 0
        Dim All As Double
        Dim Index As Integer
        If Not Strings Is Nothing Then Index = CInt(Strings)
        Output.Add(New String() {String.Empty, String.Empty, String.Empty, String.Empty, String.Empty, String.Empty})
        Output.Add(New String() {If(Index = 8 Or Index = 9, "transliteration", "arabic"), "transliteration", "translation", String.Empty, String.Empty, String.Empty})
        Output.Add(New String() {_PortableMethods.LoadResourceString(If(Index = 8 Or Index = 9, "IslamInfo_Transliteration", "IslamInfo_Arabic")), _PortableMethods.LoadResourceString("IslamInfo_Transliteration"), _PortableMethods.LoadResourceString("IslamInfo_Translation"), _PortableMethods.LoadResourceString("IslamSource_WordTotal"), String.Empty, String.Empty})
        If Index = 0 Then
            All = ChData.TotalLetters
            Dim LetterFreqArray(ChData.LetterDictionary.Keys.Count - 1) As Char
            ChData.LetterDictionary.Keys.CopyTo(LetterFreqArray, 0)
            Array.Sort(LetterFreqArray, Function(Key As Char, NextKey As Char) ChData.LetterDictionary.Item(NextKey).Count.CompareTo(ChData.LetterDictionary.Item(Key).Count))
            For Count As Integer = 0 To LetterFreqArray.Length - 1
                Total += ChData.LetterDictionary.Item(LetterFreqArray(Count)).Count
                Output.Add(New String() {ArabicData.LeftToRightEmbedding + ArbData.GetUnicodeName(LetterFreqArray(Count)) + " ( " + ArabicData.PopDirectionalFormatting + ArbData.FixStartingCombiningSymbol(LetterFreqArray(Count)) + ArabicData.LeftToRightEmbedding + " )" + ArabicData.PopDirectionalFormatting, String.Empty, String.Empty, CStr(ChData.LetterDictionary.Item(LetterFreqArray(Count)).Count), (CDbl(ChData.LetterDictionary.Item(LetterFreqArray(Count)).Count) * 100 / All).ToString("n2"), (CDbl(Total) * 100 / All).ToString("n2")})
            Next
        ElseIf Index = 7 Then
            All = ChData.TotalIsolatedLetters
            Dim LetterFreqArray(ChData.IsolatedLetterDictionary.Keys.Count - 1) As Char
            ChData.IsolatedLetterDictionary.Keys.CopyTo(LetterFreqArray, 0)
            Array.Sort(LetterFreqArray, Function(Key As Char, NextKey As Char) ChData.IsolatedLetterDictionary.Item(NextKey).Count.CompareTo(ChData.IsolatedLetterDictionary.Item(Key).Count))
            For Count As Integer = 0 To LetterFreqArray.Length - 1
                Total += ChData.IsolatedLetterDictionary.Item(LetterFreqArray(Count)).Count
                Output.Add(New String() {ArabicData.LeftToRightEmbedding + ArbData.GetUnicodeName(LetterFreqArray(Count)) + " ( " + ArabicData.PopDirectionalFormatting + ArbData.FixStartingCombiningSymbol(LetterFreqArray(Count)) + ArabicData.LeftToRightEmbedding + " )" + ArabicData.PopDirectionalFormatting, String.Empty, String.Empty, CStr(ChData.IsolatedLetterDictionary.Item(LetterFreqArray(Count)).Count), (CDbl(ChData.IsolatedLetterDictionary.Item(LetterFreqArray(Count)).Count) * 100 / All).ToString("n2"), (CDbl(Total) * 100 / All).ToString("n2")})
            Next
        ElseIf Index = 1 Or Index = 11 Or Index = 12 Or Index >= 13 And Index < 13 + ChData.IslamData.PartsOfSpeech.Length + ChData.RecitationSymbols.Length + ChData.RecitationSymbols.Length + ChData.RecitationSymbols.Length Then
            Dim Dict As Generic.Dictionary(Of String, List(Of Integer()))
            If Index = 1 Then
                Dict = New Dictionary(Of String, List(Of Integer()))
                For Each KV As KeyValuePair(Of String, List(Of String)) In ChData.WordDictionary
                    Dim Str As String = KV.Key + vbCrLf + String.Join(vbCrLf, CType(KV.Value.ToArray(), String()))
                    Dict.Add(Str, New List(Of Integer()))
                    For Count As Integer = 0 To KV.Value.Count - 1
                        Dict(Str).AddRange(ChData.FormDictionary(CStr(KV.Value(Count))))
                    Next
                Next
            ElseIf Index = 11 Then
                Dict = ChData.PreDictionary
            ElseIf Index = 12 Then
                Dict = ChData.SufDictionary
            ElseIf Index >= 13 And Index < 13 + ChData.IslamData.PartsOfSpeech.Length Then
                Dict = ChData.TagDictionary(ChData.IslamData.PartsOfSpeech(Index - 13).Symbol)
            ElseIf Index >= 13 + ChData.IslamData.PartsOfSpeech.Length And Index < 13 + ChData.IslamData.PartsOfSpeech.Length + ChData.RecitationSymbols.Length Then
                Dict = ChData.LetterDictionary(ChData.RecitationSymbols(Index - 13 - ChData.IslamData.PartsOfSpeech.Length).Chars(0))
            ElseIf Index >= 13 + ChData.IslamData.PartsOfSpeech.Length + ChData.RecitationSymbols.Length And Index < 13 + ChData.IslamData.PartsOfSpeech.Length + ChData.RecitationSymbols.Length + ChData.RecitationSymbols.Length Then
                Dict = ChData.LetterPreDictionary(ChData.RecitationSymbols(Index - 13 - ChData.IslamData.PartsOfSpeech.Length - ChData.RecitationSymbols.Length).Chars(0))
            ElseIf Index >= 13 + ChData.IslamData.PartsOfSpeech.Length + ChData.RecitationSymbols.Length + ChData.RecitationSymbols.Length And Index < 13 + ChData.IslamData.PartsOfSpeech.Length + ChData.RecitationSymbols.Length + ChData.RecitationSymbols.Length + ChData.RecitationSymbols.Length Then
                Dict = ChData.LetterSufDictionary(ChData.RecitationSymbols(Index - 13 - ChData.IslamData.PartsOfSpeech.Length - ChData.RecitationSymbols.Length - ChData.RecitationSymbols.Length).Chars(0))
            Else
                Dict = Nothing
            End If
            Dim FreqArray(Dict.Keys.Count - 1) As String
            Dict.Keys.CopyTo(FreqArray, 0)
            Total = 0
            All = GetQuranWordTotalNumber()
            Array.Sort(FreqArray, Function(Key As String, NextKey As String) Dict.Item(NextKey).Count.CompareTo(Dict.Item(Key).Count))
            Dim W4WLines As String() = Await _PortableMethods.ReadAllLines(_PortableMethods.Settings.GetFilePath(_PortableMethods.FileIO.CombinePath("metadata", ChData.IslamData.LanguageDefaultInfo.GetLanguageByID(Threading.Thread.CurrentThread.CurrentUICulture.TwoLetterISOLanguageName).QuranW4WFile + ".txt")))
            For Count As Integer = 0 To FreqArray.Length - 1
                Dim TranslationDict As New Dictionary(Of String, List(Of Integer()))
                For WordCount As Integer = 0 To Dict.Item(FreqArray(Count)).Count - 1
                    Dim CheckStr As String = GetW4WTranslationVerse(W4WLines, Dict.Item(FreqArray(Count))(WordCount)(0), Dict.Item(FreqArray(Count))(WordCount)(1), Dict.Item(FreqArray(Count))(WordCount)(2))
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
                Dim Str As String = Arb.TransliterateFromBuckwalter(FreqArray(Count))
                Output.Add(New String() {Str, Arb.TransliterateToScheme(Str, If(SchemeType = ArabicData.TranslitScheme.RuleBased, ArabicData.TranslitScheme.LearningMode, SchemeType), Scheme, Arabic.FilterMetadataStops(Str, Arb.GetMetarules(Str, ChData.RuleMetas("Normal")), Nothing)), String.Join(vbCrLf, TranslationArray), CStr(Dict.Item(FreqArray(Count)).Count), (CDbl(Dict.Item(FreqArray(Count)).Count) * 100 / All).ToString("n2"), (CDbl(Total) * 100 / All).ToString("n2")})
            Next
        ElseIf Index = 3 Or Index = 4 Or Index = 5 Or Index = 6 Then
            Total = 0
            Dim DivArray As Collections.Generic.List(Of String)()
            If Index = 3 Or Index = 5 Then
                DivArray = If(Index = 5, ChData.StationUniqueArray, ChData.PartUniqueArray)
                All = If(Index = 5, ChData.TotalUniqueWordsInStations, ChData.TotalUniqueWordsInParts)
                For Count As Integer = 0 To CInt(If(Index = 5, GetStationCount(), GetPartCount())) - 1
                    Total += DivArray(Count).Count
                    Output.Add(New String() {ArabicData.LeftToRightEmbedding + CStr(Count + 1) + ArabicData.PopDirectionalFormatting, String.Empty, String.Empty, CStr(DivArray(Count).Count), (CDbl(DivArray(Count).Count) * 100 / All).ToString("n2"), (CDbl(Total) * 100 / All).ToString("n2")})
                Next
            ElseIf Index = 4 Or Index = 6 Then
                DivArray = If(Index = 6, ChData.StationUniqueArray, ChData.PartUniqueArray)
                All = If(Index = 6, ChData.TotalWordsInStations, ChData.TotalWordsInParts)
                For Count As Integer = 0 To CInt(If(Index = 6, GetStationCount(), GetPartCount())) - 1
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
        Return Output.ToArray()
    End Function
    Public Shared Function IsSingletonDictionary(Dict As Dictionary(Of String, Object())) As Boolean
        Dim Keys(Dict.Keys.Count - 1) As String
        Dict.Keys.CopyTo(Keys, 0)
        Return Dict.Keys.Count = 1 AndAlso (CType(Dict(Keys(0))(0), Dictionary(Of String, Object())).Keys.Count = 0 OrElse IsSingletonDictionary(CType(Dict(Keys(0))(0), Dictionary(Of String, Object()))))
    End Function
    Public Function DumpRecDictionary(Dict As Dictionary(Of String, Object()), Post As Boolean, Depth As Integer, Dual As Boolean) As String
        Dim Str As String = String.Empty
        For Each KV As KeyValuePair(Of String, Object()) In Dict
            If Str <> String.Empty Then Str = If(Post, Str + "/", "/" + Str)
            'if all are 1 recursively
            If IsSingletonDictionary(CType(KV.Value(0), Dictionary(Of String, Object()))) Then
                If Post Then
                    Str += Arb.TransliterateToScheme(KV.Key, ArabicData.TranslitScheme.Literal, String.Empty, Nothing) + DumpRecDictionary(CType(KV.Value(0), Dictionary(Of String, Object())), Post, Depth + 1, Dual) + If(Dual AndAlso CType(KV.Value(1), Dictionary(Of String, Object())).Keys.Count <> 0, "!(" + DumpRecDictionary(CType(KV.Value(1), Dictionary(Of String, Object())), Not Post, Depth + 1, False) + ")!", String.Empty)
                Else
                    Str = DumpRecDictionary(CType(KV.Value(0), Dictionary(Of String, Object())), Post, Depth + 1, Dual) + If(Dual AndAlso CType(KV.Value(1), Dictionary(Of String, Object())).Keys.Count <> 0, "!(" + DumpRecDictionary(CType(KV.Value(1), Dictionary(Of String, Object())), Not Post, Depth + 1, False) + ")!", String.Empty) + Arb.TransliterateToScheme(KV.Key, ArabicData.TranslitScheme.Literal, String.Empty, Nothing) + Str
                End If
            ElseIf Post Then
                Str += Arb.TransliterateToScheme(KV.Key, ArabicData.TranslitScheme.Literal, String.Empty, Nothing) + If(CType(KV.Value(0), Dictionary(Of String, Object())).Keys.Count = 0, String.Empty, vbCrLf + New String(" "c, Depth * 4) + "(" + DumpRecDictionary(CType(KV.Value(0), Dictionary(Of String, Object())), Post, Depth + 1, Dual) + ")" + vbCrLf + New String(" "c, Depth * 4)) + If(Dual AndAlso CType(KV.Value(1), Dictionary(Of String, Object())).Keys.Count <> 0, vbCrLf + New String(" "c, Depth * 4) + "!(" + DumpRecDictionary(CType(KV.Value(1), Dictionary(Of String, Object())), Not Post, Depth + 1, False) + ")!" + vbCrLf + New String(" "c, Depth * 4), String.Empty)
            Else
                Str = If(CType(KV.Value(0), Dictionary(Of String, Object())).Keys.Count = 0, String.Empty, vbCrLf + New String(" "c, Depth * 4) + "(" + DumpRecDictionary(CType(KV.Value(0), Dictionary(Of String, Object())), Post, Depth + 1, Dual) + ")" + vbCrLf + New String(" "c, Depth * 4)) + If(Dual AndAlso CType(KV.Value(1), Dictionary(Of String, Object())).Keys.Count <> 0, vbCrLf + New String(" "c, Depth * 4) + "!(" + DumpRecDictionary(CType(KV.Value(1), Dictionary(Of String, Object())), Not Post, Depth + 1, False) + ")!" + vbCrLf + New String(" "c, Depth * 4), String.Empty) + Arb.TransliterateToScheme(KV.Key, ArabicData.TranslitScheme.Literal, String.Empty, Nothing) + Str
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
    Public Function PatternAnalysis() As String()
        Dim Verses As List(Of String()) = GetQuranText(ChData.XMLDocMain, -1, -1, -1, -1)
        Dim PreDict As Dictionary(Of String, Object())
        Dim PostDict As Dictionary(Of String, Object())
        Dim Strings(ChData.RecitationSymbols.Length - 1) As String
        For LetCount = 0 To ChData.RecitationSymbols.Length - 1
            PreDict = New Dictionary(Of String, Object())
            PostDict = New Dictionary(Of String, Object())
            For Count As Integer = 0 To Verses.Count - 1
                For SubCount As Integer = 0 To Verses(Count).Length - 1
                    Dim Matches As System.Text.RegularExpressions.MatchCollection = System.Text.RegularExpressions.Regex.Matches(Verses(Count)(SubCount), If(ChData.RecitationSymbols(LetCount) = " ", "(\S*)(^|\s+|$)(\S*)", "(\S*)(" + ArabicData.MakeUniRegEx(ChData.RecitationSymbols(LetCount)) + ")(\S*)"))
                    For MatchCount As Integer = 0 To Matches.Count - 1
                        AddRecDictionary(PostDict, Matches(MatchCount).Groups(3).Value(0), Matches(MatchCount).Groups(3).Value.Substring(1), Matches(MatchCount).Groups(1).Value, True)
                        AddRecDictionary(PreDict, Matches(MatchCount).Groups(1).Value(0), Matches(MatchCount).Groups(1).Value.Substring(0, Matches(MatchCount).Groups(1).Value.Length - 1), Matches(MatchCount).Groups(3).Value, False)
                    Next
                Next
            Next
            Strings(LetCount) = ArabicData.LeftToRightEmbedding + DumpRecDictionary(PreDict, False, 0, True) + "\" + Arb.TransliterateToScheme(ChData.RecitationSymbols(LetCount), ArabicData.TranslitScheme.Literal, String.Empty, Nothing) + "\" + DumpRecDictionary(PostDict, True, 0, True) + ArabicData.PopDirectionalFormatting
        Next
        Return Strings
    End Function
    Public Function GetQuranHamzaMaddDoubleLetterPatterns() As String()
        Dim CurPat As String = ArabicData.MakeRegMultiEx(ChData.RecitationSymbols)
        Dim Prefixes As New Dictionary(Of String, List(Of String))
        Dim Suffixes As New Dictionary(Of String, List(Of String))
        Dim PreMidSuf As New Dictionary(Of String, Object()) 'Prefix indexed
        Dim SufMidPre As New Dictionary(Of String, Object()) 'Suffix indexed
        For Each Key As String In ChData.FormDictionary.Keys
            Dim Matches As System.Text.RegularExpressions.MatchCollection = System.Text.RegularExpressions.Regex.Matches(Arb.TransliterateFromBuckwalter(Key), CurPat)
            For Count As Integer = 0 To Matches.Count - 1
                If Matches(Count).Index = 0 Then
                    For SubCount As Integer = 0 To ChData.FormDictionary(Key).Count - 1
                        Dim Loc(3) As Integer
                        ChData.FormDictionary(Key)(SubCount).CopyTo(Loc, 0)
                        If (Not CBool(ChData.LocDictionary(String.Join(":", Loc))(1)) And Not CBool(ChData.LocDictionary(String.Join(":", Loc))(2))) Or CStr(ChData.LocDictionary(String.Join(":", Loc))(3)) = "DET" Or (Matches(Count).Value = ArabicData.ArabicLetterAlefWasla And CStr(ChData.LocDictionary(String.Join(":", Loc))(3)) = "PN") Then
                            Dim Pre As String = String.Empty
                            For LocCount = 1 To ChData.FormDictionary(Key)(SubCount)(3) - 1
                                Loc(3) = LocCount
                                Pre += CStr(ChData.LocDictionary(String.Join(":", Loc))(0))
                            Next
                            If Pre <> String.Empty Then
                                Loc(3) = ChData.FormDictionary(Key)(SubCount)(3)
                                If CStr(ChData.LocDictionary(String.Join(":", Loc))(3)) = "DET" Or (Matches(Count).Value = ArabicData.ArabicLetterAlefWasla And CStr(ChData.LocDictionary(String.Join(":", Loc))(3)) = "PN") Then
                                    If Not Prefixes.ContainsKey("Al+") Then Prefixes.Add("Al+", New List(Of String))
                                    Prefixes("Al+").Add(New String(New List(Of Char)(Linq.Enumerable.Reverse(Pre.ToCharArray())).ToArray()))
                                Else
                                    If Not Prefixes.ContainsKey(Matches(Count).Value) Then Prefixes.Add(Matches(Count).Value, New List(Of String))
                                    Prefixes(Matches(Count).Value).Add(New String(New List(Of Char)(Linq.Enumerable.Reverse(Pre.ToCharArray())).ToArray()))
                                End If
                                If Matches(Count).Value <> ArabicData.ArabicLetterAlefWasla And CStr(ChData.LocDictionary(String.Join(":", Loc))(3)) <> "DET" Then
                                    If Not Prefixes.ContainsKey("!" + ArabicData.ArabicLetterAlefWasla) Then Prefixes.Add("!" + ArabicData.ArabicLetterAlefWasla, New List(Of String))
                                    Prefixes("!" + ArabicData.ArabicLetterAlefWasla).Add(New String(New List(Of Char)(Linq.Enumerable.Reverse(Pre.ToCharArray())).ToArray()))
                                End If
                            End If
                        End If
                    Next
                ElseIf Matches(Count).Index = Key.Length - 1 Then
                    For SubCount As Integer = 0 To ChData.FormDictionary(Key).Count - 1
                        Dim Loc(3) As Integer
                        ChData.FormDictionary(Key)(SubCount).CopyTo(Loc, 0)
                        If (Not CBool(ChData.LocDictionary(String.Join(":", Loc))(1)) And Not CBool(ChData.LocDictionary(String.Join(":", Loc))(2))) Then
                            Dim Sup As String = String.Empty
                            Dim LocCount As Integer = ChData.FormDictionary(Key)(SubCount)(3) + 1
                            Do
                                Loc(3) = LocCount
                                If Not ChData.LocDictionary.ContainsKey(String.Join(":", Loc)) Then Exit Do
                                Sup += CStr(ChData.LocDictionary(String.Join(":", Loc))(0))
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
                    For SubCount As Integer = 0 To ChData.FormDictionary(Key).Count - 1
                        Dim PreCheck As String = String.Empty
                        Dim Loc(3) As Integer
                        ChData.FormDictionary(Key)(SubCount).CopyTo(Loc, 0)
                        'Hamza prefix then must look before
                        'If PreCheck.IndexOfAny({ArabicData.ArabicLetterHamza, ArabicData.ArabicLetterAlefWithHamzaAbove, ArabicData.ArabicLetterAlefWithHamzaBelow, ArabicData.ArabicLetterWawWithHamzaAbove, ArabicData.ArabicLetterYehWithHamzaAbove, ArabicData.ArabicHamzaAbove}) = -1 Then PreCheck = String.Empty
                        Dim Pre As String = String.Empty
                        Dim LocCount As Integer
                        PreCheck += Arb.TransliterateFromBuckwalter(Key.Substring(0, Matches(Count).Index))
                        For SupCount As Integer = PreCheck.Length - 1 To 0 Step -1
                            Pre = PreCheck(SupCount) + Pre
                            If Array.IndexOf(ChData.ArabicSunLetters, CStr(PreCheck(SupCount))) <> -1 Or Array.IndexOf(ChData.ArabicMoonLettersNoVowels, CStr(PreCheck(SupCount))) <> -1 Then Exit For
                        Next
                        If Pre.Length = PreCheck.Length Then
                            PreCheck = String.Empty
                            For LocCount = 1 To ChData.FormDictionary(Key)(SubCount)(3) - 1
                                Loc(3) = LocCount
                                PreCheck += Arb.TransliterateFromBuckwalter(CStr(ChData.LocDictionary(String.Join(":", Loc))(0)))
                            Next
                            If PreCheck <> String.Empty Then Pre = ";" + PreCheck + ";" + Pre
                        End If
                        Dim Sup As String = Arb.TransliterateFromBuckwalter(Key.Substring(Matches(Count).Index + 1))
                        LocCount = ChData.FormDictionary(Key)(SubCount)(3) + 1
                        Do
                            Loc(3) = LocCount
                            If Not ChData.LocDictionary.ContainsKey(String.Join(":", Loc)) Then Exit Do
                            Sup += Arb.TransliterateFromBuckwalter(CStr(ChData.LocDictionary(String.Join(":", Loc))(0)))
                            LocCount += 1
                        Loop While True
                        Dim Suf As String = String.Empty
                        For SufCount As Integer = 0 To Sup.Length - 1
                            Suf += Sup(SufCount)
                            If Array.IndexOf(ChData.ArabicSunLetters, CStr(Sup(SufCount))) <> -1 Or Array.IndexOf(ChData.ArabicMoonLettersNoVowels, CStr(Sup(SufCount))) <> -1 Or ArabicData.ArabicLetterTehMarbuta = Sup(SufCount) Then Exit For
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
        For Each Key As String In ChData.RealWordDictionary.Keys
            Dim Matches As System.Text.RegularExpressions.MatchCollection = System.Text.RegularExpressions.Regex.Matches(Key, CurPat)
            For Count As Integer = 0 To Matches.Count - 1
                'If Matches(Count).Value = ArabicData.ArabicLetterHamza Or Matches(Count).Value = ArabicData.ArabicLetterAlefWithHamzaAbove Or Matches(Count).Value = ArabicData.ArabicLetterAlefWithHamzaBelow Or Matches(Count).Value = ArabicData.ArabicLetterWawWithHamzaAbove Or Matches(Count).Value = ArabicData.ArabicLetterYehWithHamzaAbove Or Matches(Count).Value = ArabicData.ArabicHamzaAbove Then
                'If Matches(Count).Value = ArabicData.ArabicLetterAlef Or Matches(Count).Value = ArabicData.ArabicLetterWaw Or Matches(Count).Value = ArabicData.ArabicLetterYeh Or Matches(Count).Value = ArabicData.ArabicLetterAlefMaksura Or Matches(Count).Value = ArabicData.ArabicSmallWaw Or Matches(Count).Value = ArabicData.ArabicSmallYeh Or Matches(Count).Value = ArabicData.ArabicLetterSuperscriptAlef Then
                If Matches(Count).Index = Key.Length - 1 And (Array.IndexOf(ChData.ArabicSunLetters, Matches(Count).Value) = -1 And Array.IndexOf(ChData.ArabicMoonLettersNoVowels, Matches(Count).Value) = -1 And ArabicData.ArabicLetterTehMarbuta <> Matches(Count).Value) Then
                    Dim Pre As String = String.Empty
                    For SubCount As Integer = Matches(Count).Index - 1 To 0 Step -1
                        If Array.IndexOf(ChData.ArabicSunLetters, CStr(Key(SubCount))) = -1 And Array.IndexOf(ChData.ArabicMoonLettersNoVowels, CStr(Key(SubCount))) = -1 Then
                            Pre = Key(SubCount) + Pre
                        Else
                            Exit For
                        End If
                    Next
                    Dim Suf As String = String.Empty
                    For SubCount As Integer = Matches(Count).Index + 1 To Key.Length - 1
                        Suf += Key(SubCount)
                        If Array.IndexOf(ChData.ArabicSunLetters, CStr(Key(SubCount))) <> -1 Or Array.IndexOf(ChData.ArabicMoonLettersNoVowels, CStr(Key(SubCount))) <> -1 Or ArabicData.ArabicLetterTehMarbuta = Key(SubCount) Then Exit For
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
            Strings(CurNum) = ArabicData.LeftToRightEmbedding + "Key:" + Key + "\" + Arb.TransliterateToScheme(Key, ArabicData.TranslitScheme.Literal, String.Empty, Nothing) + "\" + DumpRecDictionary(CType(PreMidSuf(Key)(0), Dictionary(Of String, Object())), False, 0, True) + "\" + ArabicData.PopDirectionalFormatting
            CurNum += 1
        Next
        For Each Key As String In SufMidPre.Keys
            Strings(CurNum) = ArabicData.LeftToRightEmbedding + "Key:" + Key + "\" + Arb.TransliterateToScheme(Key, ArabicData.TranslitScheme.Literal, String.Empty, Nothing) + "\" + DumpRecDictionary(CType(SufMidPre(Key)(0), Dictionary(Of String, Object())), False, 0, True) + "\" + ArabicData.PopDirectionalFormatting
            CurNum += 1
        Next
        Return Strings
    End Function
    Public Function GetQuranLetterPatterns() As String()
        Dim RecSymbols As String = String.Join(String.Empty, Linq.Enumerable.Select(ChData.RecitationSpecialSymbols, Function(C As String) C))
        Dim LtrSymbols As String = String.Join(String.Empty, Linq.Enumerable.Select(ChData.RecitationLetters, Function(C As String) C))
        Dim DiaSymbols As String = String.Join(String.Empty, Linq.Enumerable.Select(ChData.RecitationDiacritics, Function(C As String) C))
        Dim StartWordMultiOnly As New Generic.Dictionary(Of String, String)
        Dim EndWordMultiOnly As New Generic.Dictionary(Of String, String)
        Dim MiddleWordMultiOnly As New Generic.Dictionary(Of String, String)
        Dim StartWordOnly As String = String.Join(String.Empty, Linq.Enumerable.Select(ChData.RecitationLettersDiacritics, Function(C As String) C))
        Dim NotStartWord As String = String.Join(String.Empty, Linq.Enumerable.Select(ChData.RecitationLettersDiacritics, Function(C As String) C))
        Dim EndWordOnly As String = String.Join(String.Empty, Linq.Enumerable.Select(ChData.RecitationLettersDiacritics, Function(C As String) C))
        Dim NotEndWord As String = String.Join(String.Empty, Linq.Enumerable.Select(ChData.RecitationLettersDiacritics, Function(C As String) C))
        Dim EndWordOnlyNoDia As String = String.Join(String.Empty, Linq.Enumerable.Select(ChData.RecitationLetters, Function(C As String) C))
        Dim NotEndWordNoDia As String = String.Join(String.Empty, Linq.Enumerable.Select(ChData.RecitationLetters, Function(C As String) C))
        Dim MiddleWordOnlyNoDia As String = String.Join(String.Empty, Linq.Enumerable.Select(ChData.RecitationLetters, Function(C As String) C))
        Dim NotMiddleWordNoDia As String = String.Join(String.Empty, Linq.Enumerable.Select(ChData.RecitationLetters, Function(C As String) C))
        Dim MiddleWordOnly As String = String.Join(String.Empty, Linq.Enumerable.Select(ChData.RecitationLettersDiacritics, Function(C As String) C))
        Dim NotMiddleWord As String = String.Join(String.Empty, Linq.Enumerable.Select(ChData.RecitationLettersDiacritics, Function(C As String) C))
        Dim DiaStartWordOnly As String = String.Join(String.Empty, Linq.Enumerable.Select(ChData.RecitationDiacritics, Function(C As String) C))
        Dim DiaNotStartWord As String = String.Join(String.Empty, Linq.Enumerable.Select(ChData.RecitationDiacritics, Function(C As String) C))
        Dim DiaEndWordOnly As String = String.Join(String.Empty, Linq.Enumerable.Select(ChData.RecitationDiacritics, Function(C As String) C))
        Dim DiaNotEndWord As String = String.Join(String.Empty, Linq.Enumerable.Select(ChData.RecitationDiacritics, Function(C As String) C))
        Dim DiaMiddleWordOnly As String = String.Join(String.Empty, Linq.Enumerable.Select(ChData.RecitationDiacritics, Function(C As String) C))
        Dim DiaNotMiddleWord As String = String.Join(String.Empty, Linq.Enumerable.Select(ChData.RecitationDiacritics, Function(C As String) C))
        Dim Combos As String() = String.Join("|", Linq.Enumerable.Select(ChData.RecitationLettersDiacritics, Function(C As String) String.Join("|", Linq.Enumerable.Select(ChData.RecitationLettersDiacritics, Function(Nxt As String) C + Nxt)))).Split("|"c)
        Dim DiaCombos As String() = String.Join("|", Linq.Enumerable.Select(ChData.RecitationDiacritics, Function(C As String) String.Join("|", Linq.Enumerable.Select(ChData.RecitationDiacritics, Function(Nxt As String) C + Nxt)))).Split("|"c)
        Dim LetCombos As String() = String.Join("|", Linq.Enumerable.Select(ChData.RecitationLetters, Function(C As String) String.Join("|", Linq.Enumerable.Select(ChData.RecitationLetters, Function(Nxt As String) C + Nxt)))).Split("|"c)
        For Each Key As String In ChData.FormDictionary.Keys
            Dim Str As String = New String(New List(Of Char)(Linq.Enumerable.Where(Key.ToCharArray(), Function(Ch As Char) Not RecSymbols.Contains(CStr(Ch)))).ToArray())
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
        For Each Key As String In ChData.FormDictionary.Keys
            Dim Str As String = New String(New List(Of Char)(Linq.Enumerable.Where(Key.ToCharArray(), Function(Ch As Char) Not RecSymbols.Contains(CStr(Ch)))).ToArray())
            Str = Arb.TransliterateFromBuckwalter(Str)
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
            Combos = New List(Of String)(Linq.Enumerable.Where(Combos, Function(S As String) Not Str.Contains(S))).ToArray()
            DiaCombos = New List(Of String)(Linq.Enumerable.Where(DiaCombos, Function(S As String) Not Str.Contains(S))).ToArray()
            LetCombos = New List(Of String)(Linq.Enumerable.Where(LetCombos, Function(S As String) Not New String(New List(Of Char)(Linq.Enumerable.Where(Str.ToCharArray(), Function(C As Char) LtrSymbols.Contains(C))).ToArray()).Contains(S))).ToArray()
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
                Val += ArbData.FixStartingCombiningSymbol(Key) + ArabicData.LeftToRightEmbedding + " [" + ArabicData.PopDirectionalFormatting + String.Join(" ", Linq.Enumerable.Select(Linq.Enumerable.Where((DiaSymbols + LtrSymbols).ToCharArray(), Function(C As Char) Not Dict.Item(Key).Contains(C)), Function(C As Char) ArbData.FixStartingCombiningSymbol(CStr(C)))) + ArabicData.LeftToRightEmbedding + " ]" + ArabicData.PopDirectionalFormatting + vbTab
            Else
                Val += ArbData.FixStartingCombiningSymbol(Key) + ArabicData.LeftToRightEmbedding + " ! [ " + ArabicData.PopDirectionalFormatting + String.Join(" ", Linq.Enumerable.Select(Dict.Item(Key).ToCharArray(), Function(C As Char) ArbData.FixStartingCombiningSymbol(CStr(C)))) + ArabicData.LeftToRightEmbedding + " ]" + ArabicData.PopDirectionalFormatting + vbTab
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
                RevVal += ArabicData.LeftToRightEmbedding + "[" + ArabicData.PopDirectionalFormatting + String.Join(" ", Linq.Enumerable.Select(Linq.Enumerable.Where((DiaSymbols + LtrSymbols).ToCharArray(), Function(C As Char) Not RevDict.Item(Key).Contains(C)), Function(C As Char) ArbData.FixStartingCombiningSymbol(CStr(C)))) + ArabicData.LeftToRightEmbedding + " ] " + ArabicData.PopDirectionalFormatting + ArbData.FixStartingCombiningSymbol(Key) + vbTab
            Else
                RevVal += ArabicData.LeftToRightEmbedding + "! [ " + ArabicData.PopDirectionalFormatting + String.Join(" ", Linq.Enumerable.Select(RevDict.Item(Key).ToCharArray(), Function(C As Char) ArbData.FixStartingCombiningSymbol(CStr(C)))) + ArabicData.LeftToRightEmbedding + " ] " + ArabicData.PopDirectionalFormatting + ArbData.FixStartingCombiningSymbol(Key) + vbTab
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
                DiaVal += ArbData.FixStartingCombiningSymbol(Key) + ArabicData.LeftToRightEmbedding + " [" + ArabicData.PopDirectionalFormatting + String.Join(" ", Linq.Enumerable.Select(Linq.Enumerable.Where(DiaSymbols.ToCharArray(), Function(C As Char) Not DiaDict.Item(Key).Contains(C)), Function(C As Char) ArbData.FixStartingCombiningSymbol(CStr(C)))) + ArabicData.LeftToRightEmbedding + " ]" + ArabicData.PopDirectionalFormatting + vbTab
            Else
                DiaVal += ArbData.FixStartingCombiningSymbol(Key) + ArabicData.LeftToRightEmbedding + " ! [ " + ArabicData.PopDirectionalFormatting + String.Join(" ", Linq.Enumerable.Select(DiaDict.Item(Key).ToCharArray(), Function(C As Char) ArbData.FixStartingCombiningSymbol(CStr(C)))) + ArabicData.LeftToRightEmbedding + " ]" + ArabicData.PopDirectionalFormatting + vbTab
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
                LetVal += ArbData.FixStartingCombiningSymbol(Key) + ArabicData.LeftToRightEmbedding + " [" + ArabicData.PopDirectionalFormatting + String.Join(" ", Linq.Enumerable.Select(Linq.Enumerable.Where(LtrSymbols.ToCharArray(), Function(C As Char) Not LetDict.Item(Key).Contains(C)), Function(C As Char) ArbData.FixStartingCombiningSymbol(CStr(C)))) + ArabicData.LeftToRightEmbedding + " ]" + ArabicData.PopDirectionalFormatting + vbTab
            Else
                LetVal += ArbData.FixStartingCombiningSymbol(Key) + ArabicData.LeftToRightEmbedding + " ! [ " + ArabicData.PopDirectionalFormatting + String.Join(" ", Linq.Enumerable.Select(LetDict.Item(Key).ToCharArray(), Function(C As Char) ArbData.FixStartingCombiningSymbol(CStr(C)))) + ArabicData.LeftToRightEmbedding + " ]" + ArabicData.PopDirectionalFormatting + vbTab
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
                LetRevVal += ArabicData.LeftToRightEmbedding + "[" + ArabicData.PopDirectionalFormatting + String.Join(" ", Linq.Enumerable.Select(Linq.Enumerable.Where(LtrSymbols.ToCharArray(), Function(C As Char) Not LetRevDict.Item(Key).Contains(C)), Function(C As Char) ArbData.FixStartingCombiningSymbol(CStr(C)))) + ArabicData.LeftToRightEmbedding + " ] " + ArabicData.PopDirectionalFormatting + ArbData.FixStartingCombiningSymbol(Key) + vbTab
            Else
                LetRevVal += ArabicData.LeftToRightEmbedding + "! [ " + ArabicData.PopDirectionalFormatting + String.Join(" ", Linq.Enumerable.Select(LetRevDict.Item(Key).ToCharArray(), Function(C As Char) ArbData.FixStartingCombiningSymbol(CStr(C)))) + ArabicData.LeftToRightEmbedding + " ] " + ArabicData.PopDirectionalFormatting + ArbData.FixStartingCombiningSymbol(Key) + vbTab
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
        Return {ArabicData.LeftToRightEmbedding + "Unique Prefix: [" + ArabicData.PopDirectionalFormatting + StartMulti + ArabicData.LeftToRightEmbedding + "]" + ArabicData.PopDirectionalFormatting,
                ArabicData.LeftToRightEmbedding + "Unique Suffix: [" + ArabicData.PopDirectionalFormatting + EndMulti + ArabicData.LeftToRightEmbedding + "]" + ArabicData.PopDirectionalFormatting,
                ArabicData.LeftToRightEmbedding + "Unique Middle: [" + ArabicData.PopDirectionalFormatting + MiddleMulti + ArabicData.LeftToRightEmbedding + "]" + ArabicData.PopDirectionalFormatting,
                ArabicData.LeftToRightEmbedding + "Start Only: [" + ArabicData.PopDirectionalFormatting + String.Join(" ", Linq.Enumerable.Select(StartWordOnly.ToCharArray(), Function(C As Char) ArbData.FixStartingCombiningSymbol(CStr(C)))) + ArabicData.LeftToRightEmbedding + "]" + ArabicData.PopDirectionalFormatting,
                ArabicData.LeftToRightEmbedding + "Not Start: [" + ArabicData.PopDirectionalFormatting + String.Join(" ", Linq.Enumerable.Select(NotStartWord.ToCharArray(), Function(C As Char) ArbData.FixStartingCombiningSymbol(CStr(C)))) + ArabicData.LeftToRightEmbedding + "]" + ArabicData.PopDirectionalFormatting,
                ArabicData.LeftToRightEmbedding + "End Only: [" + ArabicData.PopDirectionalFormatting + String.Join(" ", Linq.Enumerable.Select(EndWordOnly.ToCharArray(), Function(C As Char) ArbData.FixStartingCombiningSymbol(CStr(C)))) + ArabicData.LeftToRightEmbedding + "]" + ArabicData.PopDirectionalFormatting,
                ArabicData.LeftToRightEmbedding + "Not End: [" + ArabicData.PopDirectionalFormatting + String.Join(" ", Linq.Enumerable.Select(NotEndWord.ToCharArray(), Function(C As Char) ArbData.FixStartingCombiningSymbol(CStr(C)))) + ArabicData.LeftToRightEmbedding + "]" + ArabicData.PopDirectionalFormatting,
                ArabicData.LeftToRightEmbedding + "End Only No Diacritics: [" + ArabicData.PopDirectionalFormatting + String.Join(" ", Linq.Enumerable.Select(EndWordOnlyNoDia.ToCharArray(), Function(C As Char) ArbData.FixStartingCombiningSymbol(CStr(C)))) + ArabicData.LeftToRightEmbedding + "]" + ArabicData.PopDirectionalFormatting,
                ArabicData.LeftToRightEmbedding + "Not End No Diacritics: [" + ArabicData.PopDirectionalFormatting + String.Join(" ", Linq.Enumerable.Select(NotEndWordNoDia.ToCharArray(), Function(C As Char) ArbData.FixStartingCombiningSymbol(CStr(C)))) + ArabicData.LeftToRightEmbedding + "]" + ArabicData.PopDirectionalFormatting,
                ArabicData.LeftToRightEmbedding + "Middle Only: [" + ArabicData.PopDirectionalFormatting + String.Join(" ", Linq.Enumerable.Select(MiddleWordOnly.ToCharArray(), Function(C As Char) ArbData.FixStartingCombiningSymbol(CStr(C)))) + ArabicData.LeftToRightEmbedding + "]" + ArabicData.PopDirectionalFormatting,
                ArabicData.LeftToRightEmbedding + "Not Middle: [" + ArabicData.PopDirectionalFormatting + String.Join(" ", Linq.Enumerable.Select(NotMiddleWord.ToCharArray(), Function(C As Char) ArbData.FixStartingCombiningSymbol(CStr(C)))) + ArabicData.LeftToRightEmbedding + "]" + ArabicData.PopDirectionalFormatting,
                ArabicData.LeftToRightEmbedding + "Middle Only No Diacritics: [" + ArabicData.PopDirectionalFormatting + String.Join(" ", Linq.Enumerable.Select(MiddleWordOnlyNoDia.ToCharArray(), Function(C As Char) ArbData.FixStartingCombiningSymbol(CStr(C)))) + ArabicData.LeftToRightEmbedding + "]" + ArabicData.PopDirectionalFormatting,
                ArabicData.LeftToRightEmbedding + "Not Middle No Diacritics: [" + ArabicData.PopDirectionalFormatting + String.Join(" ", Linq.Enumerable.Select(NotMiddleWordNoDia.ToCharArray(), Function(C As Char) ArbData.FixStartingCombiningSymbol(CStr(C)))) + ArabicData.LeftToRightEmbedding + "]" + ArabicData.PopDirectionalFormatting,
                Val, RevVal, DiaVal, LetVal, LetRevVal}
    End Function
    Public Function GetRecitationRule(Index As Integer) As String
        Return _PortableMethods.LoadResourceString("IslamInfo_" + ChData.RuleMetas("UthmaniQuran").Rules(Index).Name)
    End Function
    Public Function GetRecitationRules() As Array()
        Dim Names() As Array = New List(Of Object())(Linq.Enumerable.Select(ChData.RuleMetas("UthmaniQuran").Rules, Function(Convert As IslamData.RuleMetaSet.RuleMetadataTranslation) New Object() {_PortableMethods.LoadResourceString("IslamInfo_" + Convert.Name), CInt(Array.IndexOf(ChData.RuleMetas("UthmaniQuran").Rules, Convert))})).ToArray()
        Array.Sort(Names, New Utility.CompareNameValueArray)
        Return Names
    End Function
    Public Function GetGroupSelectionName(DivisionsParts()() As Integer, SchemeType As ArabicData.TranslitScheme, Scheme As String) As Array()
        Return New List(Of Object())(Linq.Enumerable.Select(DivisionsParts, Function(DivisionPart, Index) New Object() {GetSelectionName(DivisionPart(0), DivisionPart(1), SchemeType, Scheme), Index})).ToArray()
    End Function
    Public Function GetSelectionName(Division As Integer, Part As Integer, SchemeType As ArabicData.TranslitScheme, Scheme As String) As String
        If Division = 0 Then
            Return GetChapterName(GetChapterByIndex(Part), SchemeType, Scheme)
        ElseIf Division = 1 Then
            Return GetChapterName(GetChapterIndexByRevelationOrder(Part), SchemeType, Scheme)
        ElseIf Division = 2 Then
            Return GetPartName(Part, SchemeType, Scheme)
        ElseIf Division = 3 Then
            Return GetGroupName(Part)
        ElseIf Division = 4 Then
            Return GetStationName(Part)
        ElseIf Division = 5 Then
            Return GetSectionName(Part)
        ElseIf Division = 6 Then
            Return GetPageName(Part)
        ElseIf Division = 7 Then
            Return GetSajdaName(Part)
        ElseIf Division = 8 Then
            Return GetImportantName(Part)
        ElseIf Division = 9 Then
            Return Arb.GetRecitationSymbol(Part, SchemeType, Scheme)
        ElseIf Division = 10 Then
            Return GetRecitationRule(Part)
        End If
        Return Nothing
    End Function
    Public Function GetSelectionNames(Strings As String, SchemeType As ArabicData.TranslitScheme, Scheme As String) As Array()
        Dim Division As Integer = 0
        If Not Strings Is Nothing Then Division = CInt(Strings)
        If Division = 0 Then
            Return GetChapterNames(SchemeType, Scheme)
        ElseIf Division = 1 Then
            Return GetChapterNamesByRevelationOrder(SchemeType, Scheme)
        ElseIf Division = 2 Then
            Return GetPartNames(SchemeType, Scheme)
        ElseIf Division = 3 Then
            Return GetGroupNames()
        ElseIf Division = 4 Then
            Return GetStationNames()
        ElseIf Division = 5 Then
            Return GetSectionNames()
        ElseIf Division = 6 Then
            Return GetPageNames()
        ElseIf Division = 7 Then
            Return GetSajdaNames()
        ElseIf Division = 8 Then
            Return GetImportantNames()
        ElseIf Division = 9 Then
            Return Arb.GetRecitationSymbols(SchemeType, Scheme)
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
    Public Sub FindMinimalVersesForCoverage()
        Dim Letters As String() = ChData.RecitationSymbols 'input is letter coverage required 
        'CachedData.ArabicLetters '3:154 and 48:29
        '{ArabicData.ArabicHamzaAbove, ArabicData.ArabicLetterHamza, ArabicData.ArabicLetterAlefWithHamzaAbove, ArabicData.ArabicLetterAlefWithHamzaBelow, ArabicData.ArabicLetterWawWithHamzaAbove, ArabicData.ArabicLetterYehWithHamzaAbove} '2:31, 3:49, 5:41, 8:19, 24:62 
        Dim SubAyah As Boolean = False 'input whether or not to break ayahs into "sub ayah" with 5 tajweed pause markers not including the 6th of forced continue
        Dim VerseDict As New Dictionary(Of String, List(Of Integer(,)))
        Dim Verses As List(Of String()) = GetQuranText(ChData.XMLDocMain, -1, -1, -1, -1)
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
    Public Async Function CheckNotablePatterns() As Threading.Tasks.Task
        'Await ComparePatterns(QuranScripts.Uthmani, QuranScripts.UthmaniMin, Arabic.UthmaniShortVowelsBeforeLongVowelsAlef)
        'Await ComparePatterns(QuranScripts.Uthmani, QuranScripts.UthmaniMin, Arabic.UthmaniShortVowelsBeforeLongVowelsSuperscriptAlef)
        'Await ComparePatterns(QuranScripts.Uthmani, QuranScripts.UthmaniMin, Arabic.UthmaniShortVowelsBeforeLongVowelsAlefAlefMaksura)
        'Await ComparePatterns(QuranScripts.Uthmani, QuranScripts.UthmaniMin, Arabic.UthmaniShortVowelsBeforeLongVowelsYehAlefMaksura)
        'Await ComparePatterns(QuranScripts.Uthmani, QuranScripts.UthmaniMin, Arabic.UthmaniShortVowelsBeforeLongVowelsYeh)
        'Await ComparePatterns(QuranScripts.Uthmani, QuranScripts.UthmaniMin, Arabic.UthmaniShortVowelsBeforeLongVowelsSmallYeh)
        'Await ComparePatterns(QuranScripts.Uthmani, QuranScripts.UthmaniMin, Arabic.UthmaniShortVowelsBeforeLongVowelsWaw)
        'Await ComparePatterns(QuranScripts.Uthmani, QuranScripts.UthmaniMin, Arabic.UthmaniShortVowelsBeforeLongVowelsSmallWaw)
        Await ComparePatterns(QuranTexts.Hafs, QuranScripts.Uthmani, QuranScripts.SimpleEnhanced, Arabic.SimpleTrailingAlef)
        'this rule should be analyzed after all other rules in Simple Enhanced are processed as it will great simplify its expression while the earlier it is processed the longer it will be
        'Await ComparePatterns(QuranScripts.Uthmani, QuranScripts.SimpleEnhanced, Arabic.SimpleSuperscriptAlef)
    End Function
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
    Public Function DumpDictionary(Dict As Dictionary(Of String, String)) As String
        Dim Msg As String = String.Empty
        Dim Keys(Dict.Keys.Count - 1) As String
        Dict.Keys.CopyTo(Keys, 0)
        Array.Sort(Keys, StringComparer.Ordinal)
        For Count As Integer = 0 To Keys.Length - 1
            Msg += """" + Arb.TransliterateToScheme(Keys(Count), ArabicData.TranslitScheme.Literal, String.Empty, Nothing) + """" + If(Count <> Keys.Length - 1, ", ", String.Empty)
        Next
        Return Msg
    End Function
    Public Async Function ComparePatterns(BaseText As QuranTexts, ScriptType As QuranScripts, CompScriptType As QuranScripts, LetterPattern As String) As Threading.Tasks.Task
        Dim WordPattern As String = "(?<=^\s*|\s+)\S*" + ArabicData.MakeUniRegEx(LetterPattern) + "\S*(?=\s+|\s*$)"
        Dim FirstList As List(Of String) = Await PatternMatch(BaseText, ScriptType, ArabicPresentation.None, WordPattern)
        FirstList.Sort(StringComparer.Ordinal)
        Dim CompList As List(Of String) = Await PatternMatch(BaseText, CompScriptType, ArabicPresentation.None, "(?<=^\s*|\s+)\S*" + ArabicData.MakeUniRegEx(LetterPattern.Substring(0, 1)) + "(?=\s+|\s*$)")
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
            Msg += """" + Arb.TransliterateToScheme(Str, ArabicData.TranslitScheme.Literal, String.Empty, Nothing) + """, "
        Next
        Msg += vbCrLf + "Second: "
        For Each Str As String In CompList
            Dim SubKey As String = Str.Substring(0, Str.IndexOf(LetterPattern))
            For Count As Integer = 0 To SubKey.Length - 1
                If Not CompDict.ContainsKey(SubKey.Substring(Count)) Then
                    CompDict.Add(SubKey.Substring(Count), Str)
                End If
            Next
            Msg += """" + Arb.TransliterateToScheme(Str, ArabicData.TranslitScheme.Literal, String.Empty, Nothing) + """, "
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
    End Function
    Public Async Function PatternMatch(BaseText As QuranTexts, ScriptType As QuranScripts, Presentation As ArabicPresentation, Pattern As String) As Threading.Tasks.Task(Of List(Of String))
        Dim PatMatch As New List(Of String)
        Dim Doc As Xml.Linq.XDocument
        Dim Stream As IO.Stream
        If ScriptType = QuranScripts.Uthmani Then
            Stream = Await _PortableMethods.FileIO.LoadStream(_PortableMethods.Settings.GetFilePath(_PortableMethods.FileIO.CombinePath("metadata", QuranTextNames(BaseText) + ".xml")))
        Else
            Stream = Await _PortableMethods.FileIO.LoadStream(_PortableMethods.Settings.GetFilePath(_PortableMethods.FileIO.CombinePath("metadata", QuranTextNames(BaseText) + "-" + QuranFileNames(ScriptType) + If(Presentation <> ArabicPresentation.None, "-" + PresentationCacheNames(Presentation), String.Empty) + ".xml")))
        End If
        Doc = Xml.Linq.XDocument.Load(Stream)
        Stream.Dispose()
        Dim Verses As Collections.Generic.List(Of String())
        Verses = GetQuranText(Doc, -1, -1, -1, -1)
        For Count As Integer = 0 To Verses.Count - 1
            Dim ChapterNode As Xml.Linq.XElement = GetTextChapter(Doc, Count + 1)
            For SubCount As Integer = 0 To Verses(Count).Length - 1
                If SubCount = 0 AndAlso Not GetTextVerse(ChapterNode, SubCount + 1).Attribute("bismillah") Is Nothing Then
                    For Each Val As System.Text.RegularExpressions.Match In System.Text.RegularExpressions.Regex.Matches(GetTextVerse(ChapterNode, SubCount + 1).Attribute("bismillah").Value, Pattern)
                        PatMatch.Add(Val.Value)
                    Next
                End If
                For Each Val As System.Text.RegularExpressions.Match In System.Text.RegularExpressions.Regex.Matches(Verses(Count)(SubCount), Pattern)
                    PatMatch.Add(Val.Value)
                Next
            Next
        Next
        Return PatMatch
    End Function
    Public Async Function CompareQuranFormats(BaseText As QuranTexts, TargetBaseText As QuranTexts, ScriptType As QuranScripts, Presentation As ArabicPresentation) As Threading.Tasks.Task
        Dim Doc As Xml.Linq.XDocument
        Dim Stream As IO.Stream = Await _PortableMethods.FileIO.LoadStream(_PortableMethods.Settings.GetFilePath(_PortableMethods.FileIO.CombinePath("metadata", QuranTextNames(BaseText) + ".xml")))
        Doc = Xml.Linq.XDocument.Load(Stream)
        Stream.Dispose()
        Dim TargetDoc As Xml.Linq.XDocument
        If BaseText = TargetBaseText Then
            Stream = Await _PortableMethods.FileIO.LoadStream(_PortableMethods.Settings.GetFilePath(_PortableMethods.FileIO.CombinePath("metadata", QuranTextNames(TargetBaseText) + "-" + QuranFileNames(ScriptType) + If(Presentation <> ArabicPresentation.None, "-" + PresentationCacheNames(Presentation), String.Empty) + ".xml")))
        Else
            Stream = Await _PortableMethods.FileIO.LoadStream(_PortableMethods.Settings.GetFilePath(_PortableMethods.FileIO.CombinePath("metadata", QuranTextNames(TargetBaseText) + "-" + QuranFileNames(ScriptType) + If(Presentation <> ArabicPresentation.None, "-" + PresentationCacheNames(Presentation), String.Empty) + ".xml")))
        End If
        TargetDoc = Xml.Linq.XDocument.Load(Stream)
        Stream.Dispose()
        Dim Verses As Collections.Generic.List(Of String())
        Dim TargetVerses As Collections.Generic.List(Of String())
        Verses = GetQuranText(Doc, -1, -1, -1, -1)
        TargetVerses = GetQuranText(TargetDoc, -1, -1, -1, -1)
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
                        Total += Words.Length - New List(Of String)(Linq.Enumerable.Where(Words, Function(Str As String) Str.Length = 1)).Count
                        SubCount += 1
                    End If
                    While (TargetTotal < Total Or TargetTotal = Total And SubCount = Verses(Count).Length) And TargetSubCount <= TargetVerses(Count).Length - 1
                        TargetWords = TargetVerses(Count)(TargetSubCount).Split(" "c)
                        TargetTotal += TargetWords.Length - New List(Of String)(Linq.Enumerable.Where(TargetWords, Function(Str As String) Str.Length = 1)).Count
                        TargetSubCount += 1
                        If Total <> TargetTotal Then
                            'Debug.Print("Chapter: " + CStr(Count + 1) + " Verse: " + CStr(SubCount) + " Words: " + CStr(Words.Length) + " Verse: " + CStr(TargetSubCount) + " Words: " + CStr(TargetWords.Length) + If(Words.Length > TargetWords.Length, "  +", "  -") + CStr(Count + 1) + ":" + CStr(Math.Max(SubCount, TargetSubCount)) + ":" + CStr(Math.Min(Words.Length, TargetWords.Length) + 1))
                        End If
                    End While
                Else
                    If TargetTotal <= Total And TargetSubCount <= TargetVerses(Count).Length - 1 Then
                        TargetWords = TargetVerses(Count)(TargetSubCount).Split(" "c)
                        TargetTotal += TargetWords.Length - New List(Of String)(Linq.Enumerable.Where(TargetWords, Function(Str As String) Str.Length = 1)).Count
                        TargetSubCount += 1
                    End If
                    While Total < TargetTotal And SubCount <= Verses(Count).Length - 1
                        Words = Verses(Count)(SubCount).Split(" "c)
                        Total += Words.Length - New List(Of String)(Linq.Enumerable.Where(Words, Function(Str As String) Str.Length = 1)).Count
                        SubCount += 1
                        If Total <> TargetTotal Then
                            'Debug.Print("Chapter: " + CStr(Count + 1) + " Verse: " + CStr(SubCount) + " Words: " + CStr(Words.Length) + " Verse: " + CStr(TargetSubCount) + " Words: " + CStr(TargetWords.Length) + If(Words.Length > TargetWords.Length, "  +", "  -") + CStr(Count + 1) + ":" + CStr(Math.Max(SubCount, TargetSubCount)) + ":" + CStr(Math.Min(Words.Length, TargetWords.Length) + 1))
                        End If
                    End While
                End If
            Loop While Total <= TargetTotal And SubCount <= Verses(Count).Length - 1 Or TargetTotal <= Total And TargetSubCount <= TargetVerses(Count).Length - 1
        Next
    End Function
    Public Async Function ChangeQuranFormat(BaseText As QuranTexts, TargetBaseText As QuranTexts, SrcScriptType As QuranScripts, ScriptType As QuranScripts, Presentation As ArabicPresentation) As Threading.Tasks.Task
        Dim Doc As Xml.Linq.XDocument
        Dim Stream As IO.Stream
        If SrcScriptType = QuranScripts.Uthmani Then
            Stream = Await _PortableMethods.FileIO.LoadStream(_PortableMethods.Settings.GetFilePath(_PortableMethods.FileIO.CombinePath("metadata", QuranTextNames(BaseText) + ".xml")))
        Else
            Stream = Await _PortableMethods.FileIO.LoadStream(_PortableMethods.Settings.GetFilePath(_PortableMethods.FileIO.CombinePath("Resources", "quran-" + QuranFileNames(SrcScriptType) + ".xml")))
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
        CType(Doc.Root.PreviousNode, Xml.Linq.XComment).Value = CType(Doc.Root.PreviousNode, Xml.Linq.XComment).Value.Replace(QuranScriptNames(SrcScriptType), QuranScriptNames(ScriptType))
        Verses = GetQuranText(Doc, -1, -1, -1, -1)
        For Count As Integer = 0 To Verses.Count - 1
            Dim VerseAdjust As Integer = 0
            Dim ChapterNode As Xml.Linq.XElement = GetTextChapter(Doc, Count + 1)
            If UseBuckwalter Then
                ChapterNode.Attribute("name").Value = Arb.TransliterateToScheme(ChapterNode.Attribute("name").Value, ArabicData.TranslitScheme.Literal, String.Empty, Nothing)
            End If
            Dim SubCount As Integer = 0
            While SubCount <= Verses(Count).Length - 1 - VerseAdjust
                Dim CurVerse As Xml.Linq.XElement = GetTextVerse(ChapterNode, SubCount + 1)
                Dim PreVerse As String = String.Empty
                Dim NextVerse As String = String.Empty
                If SubCount = 0 AndAlso Not CurVerse.Attribute("bismillah") Is Nothing Then
                    If Count <> 0 Then
                        PreVerse = GetTextVerse(GetTextChapter(Doc, Count), New List(Of Xml.Linq.XNode)(GetTextChapter(Doc, Count).Nodes).Count).Attribute("text").Value
                        If UseBuckwalter Then PreVerse = Arb.TransliterateFromBuckwalter(PreVerse)
                    End If
                    CurVerse.Attribute("bismillah").Value = If(BaseText = TargetBaseText, Arb.ChangeScript(CurVerse.Attribute("bismillah").Value, SrcScriptType, ScriptType, PreVerse, CurVerse.Attribute("text").Value), Arb.ChangeBaseScript(CurVerse.Attribute("bismillah").Value, TargetBaseText, PreVerse, CurVerse.Attribute("text").Value))
                    If UseBuckwalter Then
                        CurVerse.Attribute("bismillah").Value = Arb.TransliterateToScheme(CurVerse.Attribute("bismillah").Value, ArabicData.TranslitScheme.Literal, String.Empty, Nothing)
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
                If UseBuckwalter Then PreVerse = Arb.TransliterateFromBuckwalter(PreVerse)
                CurVerse.Attribute("text").Value = If(BaseText = TargetBaseText, Arb.ChangeScript(CurVerse.Attribute("text").Value, SrcScriptType, ScriptType, PreVerse, NextVerse), Arb.ChangeBaseScript(CurVerse.Attribute("text").Value, TargetBaseText, PreVerse, NextVerse))
                If UseBuckwalter Then
                    CurVerse.Attribute("text").Value = Arb.TransliterateToScheme(CurVerse.Attribute("text").Value, ArabicData.TranslitScheme.Literal, String.Empty, Nothing)
                End If
                If BaseText = QuranTexts.Hafs And TargetBaseText = QuranTexts.Warsh Then
                    Dim TCount As Integer = Count
                    Dim Index As Integer = New List(Of Integer())(Linq.Enumerable.TakeWhile(ChData.IslamData.VerseNumberSchemes(0).CombinedVerses, Function(Ints As Integer()) Not (TCount + 1 = Ints(0) And SubCount + 1 + VerseAdjust - 1 = Ints(1)))).Count
                    If Index <> ChData.IslamData.VerseNumberSchemes(0).CombinedVerses.Length Then
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
                    Index = New List(Of Integer())(Linq.Enumerable.TakeWhile(ChData.IslamData.VerseNumberSchemes(0).ExtraVerses, Function(Ints As Integer()) Not (TCount + 1 = Ints(0) And SubCount + 1 + VerseAdjust = Ints(1)))).Count
                    If Index <> ChData.IslamData.VerseNumberSchemes(0).ExtraVerses.Length Then
                        Dim NewNode As New Xml.Linq.XElement(CurVerse)
                        If Not NewNode.Attribute("bismillah") Is Nothing Then
                            NewNode.Attribute("bismillah").Remove()
                        End If
                        Index = ChData.IslamData.VerseNumberSchemes(0).ExtraVerses(Index)(2)
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
        Await _PortableMethods.FileIO.SaveStream(Path, MemStream)
    End Function
    Public Function IsQuranTextReference(Str As String) As Boolean
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
                ExtraVerseNumber = New List(Of Xml.Linq.XNode)(GetTextChapter(ChData.XMLDocMain, BaseChapter).Nodes).Count
            End If
            If WordNumber = 0 Then WordNumber += 1
            If BaseChapter < 1 Or BaseChapter > GetChapterCount() Then Return False
            If EndChapter <> 0 AndAlso (EndChapter < BaseChapter Or EndChapter < 1 Or EndChapter > GetChapterCount()) Then Return False
            If BaseVerse <> 0 AndAlso (BaseVerse < 1 Or BaseVerse > GetVerseCount(BaseChapter)) Then Return False
            If ExtraVerseNumber <> 0 AndAlso ((BaseChapter = If(EndChapter = 0, BaseChapter, EndChapter) And BaseVerse <> 0 And ExtraVerseNumber < BaseVerse) Or ExtraVerseNumber < 1 Or ExtraVerseNumber > GetVerseCount(If(EndChapter = 0, BaseChapter, EndChapter))) Then Return False
            Dim IndexToVerse As Integer()() = Nothing
            Dim Check As String = QuranTextCombiner(ChData.XMLDocMain, IndexToVerse, True, BaseChapter, BaseVerse, 0, EndChapter, ExtraVerseNumber, 0)
            If WordNumber < 1 Or WordNumber > Linq.Enumerable.Last(Linq.Enumerable.TakeWhile(IndexToVerse, Function(Test As Integer()) Test(0) = BaseChapter And Test(1) = BaseVerse))(3) Then Return False
            If EndWordNumber <> 0 AndAlso (BaseChapter = If(EndChapter = 0, BaseChapter, EndChapter) And BaseVerse = If(ExtraVerseNumber = 0, BaseVerse, ExtraVerseNumber) And WordNumber <> 0 And EndWordNumber < WordNumber Or EndWordNumber < 1 Or EndWordNumber > IndexToVerse(IndexToVerse.Length - 1)(3)) Then Return False
        Next
        Return True
    End Function
    Public Async Function GetW4WLines(Translation As String) As Task(Of String())
        Return Await _PortableMethods.ReadAllLines(_PortableMethods.Settings.GetFilePath(_PortableMethods.FileIO.CombinePath("metadata", Translation + ".txt")))
    End Function
    Public Async Function GetLines(Translation As String) As Task(Of String())
        Return Await _PortableMethods.ReadAllLines(_PortableMethods.Settings.GetFilePath(_PortableMethods.FileIO.CombinePath("metadata", GetTranslationFileName(Translation))))
    End Function
    Public Async Function QuranTextFromReference(Str As String, SchemeType As ArabicData.TranslitScheme, Scheme As String, TranslationIndex As Integer, W4W As Boolean, W4WNum As Boolean, NoArabic As Boolean, Header As Boolean, NoRef As Boolean, Colorize As Boolean, Verses As Boolean) As Threading.Tasks.Task(Of RenderArray)
        Dim Renderer As New RenderArray(String.Empty)
        Dim Matches As System.Text.RegularExpressions.MatchCollection = System.Text.RegularExpressions.Regex.Matches(Str, "(?:,?(\d+)(?:\:(\d+))?(?:\:(\d+))?(?:-(\d+)(?:\:(\d+))?(?:\:(\d+))?)?)")
        Dim Refs As New List(Of Integer())
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
                ExtraVerseNumber = New List(Of Xml.Linq.XNode)(GetTextChapter(ChData.XMLDocMain, BaseChapter).Nodes).Count
            End If
            If WordNumber = 0 Then WordNumber += 1
            Refs.Add(New Integer() {BaseChapter, BaseVerse, WordNumber, EndChapter, ExtraVerseNumber, EndWordNumber})
        Next
        For Count = 0 To Refs.Count - 1
            Dim IndexToVerse As Integer()() = Nothing
            Dim QuranText As String = QuranTextCombiner(ChData.XMLDocMain, IndexToVerse, True, Refs(Count)(0), Refs(Count)(1), Refs(Count)(2), Refs(Count)(3), Refs(Count)(4), Refs(Count)(5))
            Renderer.Items.AddRange(DoGetRenderedQuranText(QuranText, IndexToVerse, {Await GetLines(ChData.IslamData.Translations.TranslationList(TranslationIndex).Name)}, {Await GetW4WLines(ChData.IslamData.LanguageDefaultInfo.GetLanguageByID(Threading.Thread.CurrentThread.CurrentUICulture.TwoLetterISOLanguageName).QuranW4WFile)}, SchemeType, Scheme, {ChData.IslamData.Translations.TranslationList(TranslationIndex).FileName.Substring(0, 2)}, {Languages.GetLanguageInfoByCode(Threading.Thread.CurrentThread.CurrentUICulture.TwoLetterISOLanguageName, ChData.IslamData.LanguageList).IsRTL}, W4W, W4WNum, NoArabic, Header, Colorize, Verses, True).Items)
        Next
        Return Renderer
    End Function
    Public Function SortMergeReferences(Refs As List(Of Integer()), IndexToVerse As Integer()(), HasLetterRefs As Boolean, ByChunk As Boolean, Optional bMergeContiguous As Boolean = True) As List(Of Integer())
        If HasLetterRefs Then
            For Count As Integer = 0 To Refs.Count - 1
                Refs(Count)(2) = IndexToVerse(Array.BinarySearch(IndexToVerse, Refs(Count)(2), New TanzilReader.QuranWordIndexComparer))(3)
            Next
        End If
        If ByChunk Then
            For Count As Integer = 0 To Refs.Count - 1
                Refs(Count)(2) = IndexToVerse(Array.BinarySearch(IndexToVerse, Refs(Count), New TanzilReader.QuranWordChapterVerseWordComparer(False)))(5)
            Next
        End If
        Dim Arr As Integer()() = Refs.ToArray()
        Array.Sort(Arr, New TanzilReader.QuranWordChapterVerseWordComparer(False))
        Refs = New List(Of Integer())(Arr)
        Dim Idx As Integer = 0
        While Idx < Refs.Count - 1
            'identical, overlap, contiguous
            If Refs(Idx)(0) = Refs(Idx + 1)(0) And Refs(Idx)(1) = Refs(Idx + 1)(1) And Refs(Idx)(2) = Refs(Idx + 1)(2) Or
               Refs(Idx)(3) > Refs(Idx + 1)(0) Or
               Refs(Idx)(3) = Refs(Idx + 1)(0) And Refs(Idx)(4) > Refs(Idx + 1)(1) Or
               Refs(Idx)(3) = Refs(Idx + 1)(0) And Refs(Idx)(4) = Refs(Idx + 1)(1) And Refs(Idx)(5) >= Refs(Idx + 1)(2) Or
               bMergeContiguous And If(Refs(Idx)(5) = If(ByChunk, GetChunkCount(Refs(Idx)(3), Refs(Idx)(4)), GetWordCount(Refs(Idx)(3), Refs(Idx)(4))),
               If(Refs(Idx)(4) = GetVerseCount(Refs(Idx)(3)),
               Refs(Idx)(3) + 1 = Refs(Idx + 1)(0) And 1 = Refs(Idx + 1)(1) And 1 = Refs(Idx + 1)(2),
                Refs(Idx)(3) = Refs(Idx + 1)(0) And Refs(Idx)(4) + 1 = Refs(Idx + 1)(1) And 1 = Refs(Idx + 1)(2)),
                Refs(Idx)(3) = Refs(Idx + 1)(0) And Refs(Idx)(4) = Refs(Idx + 1)(1) And Refs(Idx)(5) + 1 = Refs(Idx + 1)(2)) Then
                If Refs(Idx)(3) < Refs(Idx + 1)(3) Or
                    Refs(Idx)(3) = Refs(Idx + 1)(3) And Refs(Idx)(4) < Refs(Idx + 1)(4) Or
                    Refs(Idx)(3) = Refs(Idx + 1)(3) And Refs(Idx)(4) = Refs(Idx + 1)(4) And Refs(Idx)(5) < Refs(Idx + 1)(5) Then
                    Refs(Idx)(3) = Refs(Idx + 1)(3)
                    Refs(Idx)(4) = Refs(Idx + 1)(4)
                    Refs(Idx)(5) = Refs(Idx + 1)(5)
                End If
                Refs.RemoveAt(Idx + 1)
            End If
            Idx += 1
        End While
        If ByChunk Then
            For Count As Integer = 0 To Refs.Count - 1
                Idx = Array.BinarySearch(IndexToVerse, Refs(Count), New TanzilReader.QuranWordChapterVerseWordComparer(True))
                Refs(Count)(2) = Linq.Enumerable.Last(Linq.Enumerable.TakeWhile(Linq.Enumerable.Reverse(Linq.Enumerable.Take(IndexToVerse, Idx + 1)), Function(It) It(0) = IndexToVerse(Idx)(0) And It(1) = IndexToVerse(Idx)(1) And It(5) = IndexToVerse(Idx)(5)))(2)
                Refs(Count)(5) = Linq.Enumerable.Last(Linq.Enumerable.TakeWhile(Linq.Enumerable.Skip(IndexToVerse, Idx), Function(It) It(0) = IndexToVerse(Idx)(0) And It(1) = IndexToVerse(Idx)(1) And It(5) = IndexToVerse(Idx)(5)))(2)
            Next
        End If
        Return Refs
    End Function
    'much faster to make a word index...
    Public Async Function QuranTextFromSearch(Str As String, SchemeType As ArabicData.TranslitScheme, Scheme As String, TranslationIndex As Integer, W4W As Boolean, W4WNum As Boolean, NoArabic As Boolean, Header As Boolean, Colorize As Boolean, UseVerses As Boolean) As Threading.Tasks.Task(Of RenderArray)
        Dim Renderer As New RenderArray(String.Empty)
        Dim Verses As List(Of String()) = GetQuranText(ChData.XMLDocMain, -1, -1, -1, -1)
        Dim RefList As String = String.Empty
        Dim RefCount As Integer = 0
        Dim Refs As New List(Of Integer())
        For Count As Integer = 0 To Verses.Count - 1
            For SubCount As Integer = 0 To Verses(Count).Length - 1
                Dim Matches As System.Text.RegularExpressions.MatchCollection = System.Text.RegularExpressions.Regex.Matches(Arb.TransliterateToScheme(Verses(Count)(SubCount), ArabicData.TranslitScheme.Literal, String.Empty, Nothing), Str)
                For MatchCount As Integer = 0 To Matches.Count - 1
                    Refs.Add(New Integer() {Count + 1, SubCount + 1, New List(Of Char)(Linq.Enumerable.Where(Verses(Count)(SubCount).Substring(0, Matches(MatchCount).Index).ToCharArray(), Function(Ch As Char) Ch = " "c)).Count + 1, Count + 1, SubCount + 1, New List(Of Char)(Linq.Enumerable.Where(Verses(Count)(SubCount).Substring(0, Matches(MatchCount).Index + Matches(MatchCount).Length).ToCharArray(), Function(Ch As Char) Ch = " "c)).Count + 1})
                    Dim Reference As String = CStr(Count + 1) + ":" + CStr(SubCount + 1)
                    RefList += If(RefList <> String.Empty, ",", String.Empty) + Reference
                    RefCount += 1
                    If New List(Of Char)(Linq.Enumerable.Where(Verses(Count)(SubCount).Substring(0, Matches(MatchCount).Index).ToCharArray(), Function(Ch As Char) Ch = " "c)).Count <> 0 OrElse New List(Of Char)(Linq.Enumerable.Where(Verses(Count)(SubCount).Substring(Matches(MatchCount).Index + Matches(MatchCount).Length).ToCharArray(), Function(Ch As Char) Ch = " "c)).Count <> 0 Then
                        RefList += ":" + CStr(New List(Of Char)(Linq.Enumerable.Where(Verses(Count)(SubCount).Substring(0, Matches(MatchCount).Index).ToCharArray(), Function(Ch As Char) Ch = " "c)).Count + 1)
                        If New List(Of Char)(Linq.Enumerable.Where(Verses(Count)(SubCount).Substring(Matches(MatchCount).Index + Matches(MatchCount).Length).ToCharArray(), Function(Ch As Char) Ch = " "c)).Count <> 0 Then RefList += "-" + CStr(New List(Of Char)(Linq.Enumerable.Where(Verses(Count)(SubCount).Substring(0, Matches(MatchCount).Index + Matches(MatchCount).Length).ToCharArray(), Function(Ch As Char) Ch = " "c)).Count + 1)
                    End If
                Next
            Next
        Next
        For Count = 0 To Refs.Count - 1
            Dim IndexToVerse As Integer()() = Nothing
            Dim QuranText As String = QuranTextCombiner(ChData.XMLDocMain, IndexToVerse, False, Refs(Count)(0), Refs(Count)(1), Refs(Count)(2), Refs(Count)(3), Refs(Count)(4), Refs(Count)(5))
            Renderer.Items.AddRange(DoGetRenderedQuranText(QuranText, IndexToVerse, {Await GetLines(ChData.IslamData.Translations.TranslationList(TranslationIndex).Name)}, {Await GetW4WLines(ChData.IslamData.LanguageDefaultInfo.GetLanguageByID(Threading.Thread.CurrentThread.CurrentUICulture.TwoLetterISOLanguageName).QuranW4WFile)}, SchemeType, Scheme, {ChData.IslamData.Translations.TranslationList(TranslationIndex).FileName.Substring(0, 2)}, {Languages.GetLanguageInfoByCode(Threading.Thread.CurrentThread.CurrentUICulture.TwoLetterISOLanguageName, ChData.IslamData.LanguageList).IsRTL}, W4W, W4WNum, NoArabic, Header, Colorize, UseVerses, True).Items)
        Next
        Renderer.Items.Add(New RenderArray.RenderItem(RenderArray.RenderTypes.eHeaderCenter, New RenderArray.RenderText() {New RenderArray.RenderText(RenderArray.RenderDisplayClass.eLTR, "(" + Arb.TransliterateToScheme(Arb.GetCatNoun("QuranReadingRecitation")(0).Text, SchemeType, Scheme, Arabic.FilterMetadataStops(Arb.GetCatNoun("QuranReadingRecitation")(0).Text, Arb.GetMetarules(Arb.GetCatNoun("QuranReadingRecitation")(0).Text, ChData.RuleMetas("Normal")), Nothing)) + " " + RefList + ") " + CStr(RefCount) + " Total")}))
        Return Renderer
    End Function
    Public Function TextPositionToMorphology(Text As String, WordPos As Integer) As String
        Dim Chapter As Integer = System.Text.RegularExpressions.Regex.Matches(Text.Substring(0, WordPos), ArabicData.ArabicEndOfAyah + Arb.TransliterateFromBuckwalter("1") + "\s").Count
        Dim Verse As Integer = Integer.Parse(Arb.TransliterateToScheme(System.Text.RegularExpressions.Regex.Match(Text.Substring(WordPos), ArabicData.ArabicEndOfAyah + "(\d{1,3})").Groups(1).Value, ArabicData.TranslitScheme.Literal, String.Empty, Nothing))
        If Verse = 1 Then Chapter += 1
        Dim Word As Integer = System.Text.RegularExpressions.Regex.Matches(Text.Substring(0, WordPos).Substring(Text.Substring(0, WordPos).LastIndexOf(ArabicData.ArabicEndOfAyah) + 1), "(\s.)?\s").Count
        Dim Lines As String() = ChData.MorphLines
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
    Public Async Function CheckSequentialRules() As Threading.Tasks.Task
        Dim Rules As IslamData.RuleTranslationCategory.RuleTranslation() = ChData.GetRuleSet("SimpleScriptHamzaWriting") '"HamzaWriting")
        Dim IndexToVerse As Integer()() = Nothing
        Dim XMLDocAlt As Xml.Linq.XDocument
        Dim Stream As IO.Stream = Await _PortableMethods.FileIO.LoadStream(QuranTextNames(QuranTexts.Hafs) + "-" + QuranFileNames(QuranScripts.SimpleEnhanced) + ".xml")
        XMLDocAlt = Xml.Linq.XDocument.Load(Stream)
        Stream.Dispose()
        Dim Text As String = QuranTextCombiner(XMLDocAlt, IndexToVerse) 'CachedData.XMLDocMain
        Dim Matches As System.Text.RegularExpressions.MatchCollection = System.Text.RegularExpressions.Regex.Matches(Text, ChData.GetPattern("Hamzas"))
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
    End Function
    Public Sub CheckMutualExclusiveRules(bAssumeContinue As Boolean, VerIndex As Integer)
        'Dim Verify As String() = {CStr(ArabicData.ArabicLetterHamza), ArabicData.ArabicTatweel + "?" + ArabicData.ArabicHamzaAbove, ArabicData.ArabicLetterAlefWithHamzaAbove, ArabicData.ArabicLetterAlefWithHamzaBelow, ArabicData.ArabicLetterWawWithHamzaAbove, ArabicData.ArabicLetterYehWithHamzaAbove}
        Dim IndexToVerse As Integer()() = Nothing
        Dim Text As String = QuranTextCombiner(ChData.XMLDocMain, IndexToVerse)
        Dim Matches As System.Text.RegularExpressions.MatchCollection = System.Text.RegularExpressions.Regex.Matches(Text, ChData.TranslateRegEx(ChData.IslamData.VerificationSet(VerIndex).Match, True))
        Dim CheckMatches As New Dictionary(Of Integer, String)
        Debug.WriteLine(CStr(Matches.Count))
        For Count = 0 To Matches.Count - 1
            If Matches(Count).Length = 0 Then Continue For 'Avoid zero width matches for end of string anchors
            For LenCount = 0 To Matches(Count).Length - 1
                If Not CheckMatches.ContainsKey(Matches(Count).Index + LenCount) Then CheckMatches.Add(Matches(Count).Index + LenCount, String.Empty)
                CheckMatches(Matches(Count).Index + LenCount) += "0"
            Next
        Next
        Dim MatchMetadata As String() = ChData.IslamData.VerificationSet(VerIndex).Evaluator
        For MainCount = 0 To ChData.IslamData.VerificationSet(VerIndex).MetaRules.Length - 1
            Dim MetaCount As Integer
            Dim bVerSet As Boolean = False
            For MetaCount = 0 To ChData.RuleMetas("UthmaniQuran").Rules.Length - 1
                If ChData.IslamData.VerificationSet(VerIndex).MetaRules(MainCount) = ChData.RuleMetas("UthmaniQuran").Rules(MetaCount).Name Then
                    Exit For
                End If
            Next
            If MetaCount = ChData.RuleMetas("UthmaniQuran").Rules.Length Then
                bVerSet = True
                For MetaCount = 0 To ChData.RuleMetas("Verification").Rules.Length - 1
                    If ChData.IslamData.VerificationSet(VerIndex).MetaRules(MainCount) = ChData.RuleMetas("Verification").Rules(MetaCount).Name Then
                        Exit For
                    End If
                Next
            End If
            Matches = System.Text.RegularExpressions.Regex.Matches(Text, If(bVerSet, ChData.RuleMetas("Verification").Rules(MetaCount).Match, ChData.RuleMetas("UthmaniQuran").Rules(MetaCount).Match))
            Dim SieveCount As Integer = 0
            Dim MetaRules As IslamData.RuleMetaSet.RuleMetadataTranslation.RuleWithArgs()() = If(bVerSet, ChData.RuleMetas("Verification").Rules(MetaCount).Evaluator, ChData.RuleMetas("UthmaniQuran").Rules(MetaCount).Evaluator)
            Dim CoverageMap(MetaRules.Length - 1) As Boolean
            For Count = 0 To Matches.Count - 1
                If Count = 0 AndAlso Matches(Count).Groups.Count <> MetaRules.Length + 1 Then Debug.WriteLine("Discrepency in metadata:" + CStr(MainCount + 1) + ":" + CStr(MetaRules.Length) + ":Got:" + CStr(Matches(Count).Groups.Count - 1))
                Dim bSieve As Boolean = False
                For SubCount = 0 To MetaRules.Length - 1
                    If Matches(Count).Groups(SubCount + 1).Success Then CoverageMap(SubCount) = True
                    If New List(Of IslamData.RuleMetaSet.RuleMetadataTranslation.RuleWithArgs)(Linq.Enumerable.TakeWhile(MetaRules(SubCount), Function(Rule As IslamData.RuleMetaSet.RuleMetadataTranslation.RuleWithArgs) Rule.RuleName <> If(bAssumeContinue, "optionalstop", "optionalnotstop"))).Count <> 0 And Matches(Count).Groups(SubCount + 1).Success Then
                        'check continuity in prior patterns
                        Dim CheckLen As Integer = 0
                        For CheckCount = 1 To SubCount
                            CheckLen += Matches(Count).Groups(CheckCount).Length
                        Next
                        If Matches(Count).Index + CheckLen <> Matches(Count).Groups(SubCount + 1).Index Then Debug.WriteLine("Non-Sequential Capture:" + CStr(MainCount + 1) + ":" + Arb.TransliterateToScheme(Text.Substring(Math.Max(0, Matches(Count).Groups(SubCount + 1).Index - 15), 30), ArabicData.TranslitScheme.Literal, String.Empty, Nothing))
                        bSieve = True
                        Exit For
                    End If
                Next
                If Not bSieve Then
                    For SubCount = 0 To MetaRules.Length - 1
                        If Array.IndexOf(MatchMetadata, MetaRules(SubCount)(0).RuleName) <> -1 And Matches(Count).Groups(SubCount + 1).Success Then
                            'If Verify.Length <> 0 And Not System.Text.RegularExpressions.Regex.Match(Matches(Count).Groups(SubCount + 1).Value, Verify(MainCount - 1)).Success Then
                            '    Debug.WriteLine("Erroneous Match: " + Check(MainCount, 0) + " " + Arabic.TransliterateToScheme(Text.Substring(Math.Max(0, Matches(Count).Groups(SubCount + 1).Index - 15), 30), ArabicData.TranslitScheme.Literal, String.Empty, Nothing))
                            '    'Debug.WriteLine(TextPositionToMorphology(Text, Matches(Count).Groups(SubCount + 1).Index))
                            'End If
                            'check continuity in previous patterns
                            Dim CheckLen As Integer = 0
                            For CheckCount = 1 To SubCount
                                CheckLen += Matches(Count).Groups(CheckCount).Length
                            Next
                            If Matches(Count).Index + CheckLen <> Matches(Count).Groups(SubCount + 1).Index Then Debug.WriteLine("Non-Sequential Capture:" + CStr(MainCount + 1) + ":" + Arb.TransliterateToScheme(Text.Substring(Math.Max(0, Matches(Count).Groups(SubCount + 1).Index - 15), 30), ArabicData.TranslitScheme.Literal, String.Empty, Nothing))
                            For LenCount = 0 To Matches(Count).Groups(SubCount + 1).Length - 1
                                If Not CheckMatches.ContainsKey(Matches(Count).Groups(SubCount + 1).Index + LenCount) Then CheckMatches.Add(Matches(Count).Groups(SubCount + 1).Index + LenCount, String.Empty)
                                CheckMatches(Matches(Count).Groups(SubCount + 1).Index + LenCount) += Convert.ToString(MainCount + 1, 16)
                            Next
                            SieveCount += 1
                        End If
                    Next
                End If
            Next
            Debug.WriteLine(CStr(SieveCount))
            For Count = 0 To CoverageMap.Length - 1
                If Not CoverageMap(Count) Then Debug.WriteLine("Unused Metadata: " + CStr(Count) + " " + String.Join("|"c, MetaRules(Count)))
            Next
        Next
        Dim Keys(CheckMatches.Keys.Count - 1) As Integer
        CheckMatches.Keys.CopyTo(Keys, 0)
        Array.Sort(Keys)
        For Count = 0 To Keys.Length - 1
            If CheckMatches(Keys(Count)).Length <> 2 Then
                Debug.WriteLine(CStr(Keys(Count)) + ":" + CheckMatches(Keys(Count)) + ":" + Arb.TransliterateToScheme(Text(Keys(Count)), ArabicData.TranslitScheme.Literal, String.Empty, Nothing) + ":" + Arb.TransliterateToScheme(Text.Substring(Math.Max(0, Keys(Count) - 15), 30), ArabicData.TranslitScheme.Literal, String.Empty, Nothing))
            End If
        Next
    End Sub
    'maintaining begining based indexing and no previous word
    Public Function QuranTextFilter(QuranText As String, ByRef IndexToVerse As Integer()(), Optional StartChapter As Integer = -1, Optional BaseVerse As Integer = -1, Optional WordNumber As Integer = -1, Optional EndChapter As Integer = -1, Optional ExtraVerseNumber As Integer = -1, Optional EndWordNumber As Integer = -1) As String
        Dim Idx As Integer = Array.BinarySearch(IndexToVerse, New Integer() {If(StartChapter < 1, 1, StartChapter), If(BaseVerse = -1, 1, BaseVerse), If(WordNumber < 1, 1, WordNumber)}, New QuranWordChapterVerseWordComparer(False))
        Dim EndIdx As Integer = Idx + 1
        Do While EndIdx < IndexToVerse.Length AndAlso (If(EndChapter = -1, True, IndexToVerse(EndIdx)(0) <= EndChapter) And If(ExtraVerseNumber = -1, True, IndexToVerse(EndIdx)(1) <= ExtraVerseNumber) And If(EndWordNumber = -1, True, IndexToVerse(EndIdx)(2) <= EndWordNumber))
            EndIdx += 1
        Loop
        Dim NewIndexToVerse(EndIdx - Idx - 1)() As Integer
        Array.Copy(IndexToVerse, Idx, NewIndexToVerse, 0, EndIdx - Idx)
        IndexToVerse = NewIndexToVerse
        Return QuranText.Substring(IndexToVerse(0)(3), IndexToVerse(IndexToVerse.Length - 1)(3) - IndexToVerse(0)(3) + IndexToVerse(IndexToVerse.Length - 1)(4))
    End Function
    'index from base word, which is previous word
    '-1 for minimal start of maximal end
    '0 for verse indicates bismillah
    Public Function QuranTextCombiner(XMLDoc As Xml.Linq.XDocument, ByRef IndexToVerse As Integer()(), Optional UseBismillah As Boolean = True, Optional StartChapter As Integer = -1, Optional BaseVerse As Integer = -1, Optional WordNumber As Integer = -1, Optional EndChapter As Integer = -1, Optional ExtraVerseNumber As Integer = -1, Optional EndWordNumber As Integer = -1) As String
        Dim Verses As New List(Of String())
        Dim bBismillahPrecedes As Boolean = UseBismillah And WordNumber <= 1 And (BaseVerse = -1 Or BaseVerse = 1) AndAlso Not GetTextVerse(GetTextChapter(XMLDoc, If(StartChapter = -1, 1, StartChapter)), 1).Attribute("bismillah") Is Nothing
        Dim bChapterRollback As Boolean = StartChapter <> -1 And StartChapter <> 1 And WordNumber <= 1 And (BaseVerse = 0 Or (BaseVerse = -1 Or BaseVerse = 1) And Not bBismillahPrecedes)
        Dim bBismillahTrails As Boolean = UseBismillah And EndChapter <> -1 AndAlso EndChapter <> GetChapterCount() AndAlso (EndWordNumber = -1 Or EndWordNumber = GetWordCount(EndChapter, If(ExtraVerseNumber = -1, GetVerseCount(EndChapter), ExtraVerseNumber))) AndAlso (ExtraVerseNumber = -1 Or ExtraVerseNumber = GetVerseCount(EndChapter)) AndAlso Not GetTextVerse(GetTextChapter(XMLDoc, EndChapter + 1), 1).Attribute("bismillah") Is Nothing
        Dim bChapterRollforward As Boolean = EndChapter <> -1 AndAlso EndChapter <> GetChapterCount() AndAlso (EndWordNumber = -1 Or EndWordNumber = GetWordCount(EndChapter, If(ExtraVerseNumber = -1, GetVerseCount(EndChapter), ExtraVerseNumber))) AndAlso (ExtraVerseNumber = 0 Or (ExtraVerseNumber = -1 Or ExtraVerseNumber = GetVerseCount(EndChapter)) And Not bBismillahTrails)
        If EndChapter = 0 Or EndChapter <> -1 And EndChapter = StartChapter And Not bChapterRollback And Not bChapterRollforward Then
            Verses.Add(GetQuranText(ChData.XMLDocMain, StartChapter, If(WordNumber <= 1 And Not bBismillahPrecedes And Not (StartChapter <= 1 And BaseVerse <= 1), BaseVerse - 1, BaseVerse), If((EndWordNumber = -1 Or EndWordNumber = GetWordCount(If(EndChapter = -1, GetChapterCount(), EndChapter), If(ExtraVerseNumber = -1, GetVerseCount(If(EndChapter = -1, GetChapterCount(), EndChapter)), ExtraVerseNumber))) And Not bBismillahTrails And Not ((EndChapter = -1 Or EndChapter = GetChapterCount()) And (ExtraVerseNumber = -1 Or ExtraVerseNumber = GetVerseCount(If(EndChapter = -1, GetChapterCount(), EndChapter)))), ExtraVerseNumber + 1, ExtraVerseNumber)))
        Else
            Verses.AddRange(GetQuranText(ChData.XMLDocMain, If(bChapterRollback, StartChapter - 1, StartChapter), If(bChapterRollback, GetVerseCount(StartChapter - 1), If(WordNumber <= 1 And Not bBismillahPrecedes And Not (StartChapter <= 1 And BaseVerse <= 1), BaseVerse - 1, BaseVerse)), If(bChapterRollforward, EndChapter + 1, EndChapter), If(bChapterRollforward, If(bBismillahTrails, ExtraVerseNumber, 1), If((EndWordNumber = -1 Or EndWordNumber = GetWordCount(If(EndChapter = -1, GetChapterCount(), EndChapter), If(ExtraVerseNumber = -1, GetVerseCount(If(EndChapter = -1, GetChapterCount(), EndChapter)), ExtraVerseNumber))) And Not bBismillahTrails And Not ((EndChapter = -1 Or EndChapter = GetChapterCount()) And (ExtraVerseNumber = -1 Or ExtraVerseNumber = GetVerseCount(If(EndChapter = -1, GetChapterCount(), EndChapter)))), ExtraVerseNumber + 1, ExtraVerseNumber))))
        End If
        Dim IndexToVerseList As New List(Of Integer())
        If StartChapter <= 1 And BaseVerse <= 1 And WordNumber <= 1 Then IndexToVerseList.Add(New Integer() {0, 0, 0, 0, 0, 0, 0})
        Dim Str As New System.Text.StringBuilder
        For Count As Integer = 0 To Verses.Count
            Dim ChunkCount As Integer
            Dim MarkerCount As Integer
            Dim FilterIndex As Integer = 0
            Dim bNotFilter As Boolean
            Dim WordCount As Integer
            Dim Matches As Text.RegularExpressions.MatchCollection
            If (Verses.Count = 0 And BaseVerse = 0 And ExtraVerseNumber = 0 Or
                Count = 0 And Not bChapterRollback And bBismillahPrecedes Or
                Count = Verses.Count And Not bChapterRollforward And bBismillahTrails Or
                (Count <> 0 And Count < Verses.Count - 1 And If(BaseVerse = -1, 1, BaseVerse) = 1)) And UseBismillah Then
                Dim Node As Xml.Linq.XAttribute
                Node = GetTextVerse(GetTextChapter(XMLDoc, If(StartChapter = -1, 1, StartChapter - If(bChapterRollback, 1, 0)) + Count), 1).Attribute("bismillah")
                If Not Node Is Nothing Then
                    Matches = System.Text.RegularExpressions.Regex.Matches(Node.Value, "(?:^\s*|\s+)(?:([^\s" + String.Join(String.Empty, Linq.Enumerable.Select(ChData.ArabicStopLetters, Function(S As String) ArabicData.MakeUniRegEx(S))) + ArabicData.ArabicStartOfRubElHizb + ArabicData.ArabicPlaceOfSajdah + "]+)|([\s" + String.Join(String.Empty, Linq.Enumerable.Select(ChData.ArabicStopLetters, Function(S As String) ArabicData.MakeUniRegEx(S))) + ArabicData.ArabicStartOfRubElHizb + ArabicData.ArabicPlaceOfSajdah + "]))(?=\s*$|\s+)")
                    ChunkCount = 1
                    MarkerCount = 0
                    For WordCount = 0 To Matches.Count - 1
                        bNotFilter = StartChapter <= 1 And BaseVerse <= 1 And WordNumber <= 1 Or Count <> 0 Or Count = 0 And (bChapterRollback Or Not bBismillahPrecedes) Or WordCount - MarkerCount + 1 >= If(WordNumber <= 1, Linq.Enumerable.Count(Linq.Enumerable.Cast(Of Text.RegularExpressions.Match)(Matches), Function(It) Not It.Groups(2).Success), WordNumber - 1)
                        If (Count <> Verses.Count - 1 Or Count = Verses.Count And Not bChapterRollforward And bBismillahTrails) And WordCount - MarkerCount + 1 > If(EndWordNumber < 1 OrElse EndWordNumber = GetWordCount(EndChapter, If(ExtraVerseNumber = -1, GetVerseCount(EndChapter), ExtraVerseNumber)), 1, EndWordNumber + 1) Then Exit For
                        If Matches(WordCount).Groups(2).Success Then
                            MarkerCount += 1
                            If bNotFilter Then IndexToVerseList.Add(New Integer() {If(StartChapter = -1, 1, StartChapter - If(bChapterRollback, 1, 0)) + Count, 0, WordCount - MarkerCount + 1, Str.Length + Matches(WordCount).Groups(2).Index - FilterIndex, Matches(WordCount).Groups(2).Length, ChunkCount, MarkerCount})
                            If WordCount <> 0 AndAlso Not Matches(WordCount - 1).Groups(2).Success Then ChunkCount += 1
                        Else
                            If bNotFilter Then IndexToVerseList.Add(New Integer() {If(StartChapter = -1, 1, StartChapter - If(bChapterRollback, 1, 0)) + Count, 0, WordCount - MarkerCount + 1, Str.Length + Matches(WordCount).Groups(1).Index - FilterIndex, Matches(WordCount).Groups(1).Length, ChunkCount, 0})
                        End If
                        If Not bNotFilter Then FilterIndex = Matches(WordCount).Index + Matches(WordCount).Length + 1
                    Next
                    If WordCount <> Matches.Count Then
                        Str.Append(Node.Value.Substring(FilterIndex, Matches(WordCount).Groups(If(Matches(WordCount).Groups(2).Success, 2, 1)).Index - FilterIndex))
                    Else
                        Str.Append(Node.Value.Substring(FilterIndex))
                    End If
                    Str.Append(" "c)
                End If
            End If
            If Count <> Verses.Count Then
                For SubCount As Integer = 0 To Verses(Count).Length - 1
                    Matches = System.Text.RegularExpressions.Regex.Matches(Verses(Count)(SubCount), "(?:^\s*|\s+)(?:([^\s" + String.Join(String.Empty, Linq.Enumerable.Select(ChData.ArabicStopLetters, Function(S As String) ArabicData.MakeUniRegEx(S))) + ArabicData.ArabicStartOfRubElHizb + ArabicData.ArabicPlaceOfSajdah + "]+)|([\s" + String.Join(String.Empty, Linq.Enumerable.Select(ChData.ArabicStopLetters, Function(S As String) ArabicData.MakeUniRegEx(S))) + ArabicData.ArabicStartOfRubElHizb + ArabicData.ArabicPlaceOfSajdah + "]))(?=\s*$|\s+)")
                    ChunkCount = 1
                    MarkerCount = 0
                    FilterIndex = 0
                    For WordCount = 0 To Matches.Count - 1
                        bNotFilter = StartChapter <= 1 And BaseVerse <= 1 And WordNumber <= 1 Or Count <> 0 Or Count = 0 And (Not bChapterRollback And bBismillahPrecedes) Or SubCount <> 0 Or WordCount - MarkerCount + 1 >= If(WordNumber <= 1, Linq.Enumerable.Count(Linq.Enumerable.Cast(Of Text.RegularExpressions.Match)(Matches), Function(It) Not It.Groups(2).Success), WordNumber - 1)
                        If Count = Verses.Count - 1 And (Count <> Verses.Count Or bChapterRollforward Or Not bBismillahTrails) And SubCount = Verses(Count).Length - 1 And WordCount - MarkerCount + 1 > If(EndWordNumber < 1 OrElse EndWordNumber = GetWordCount(EndChapter, If(ExtraVerseNumber = -1, GetVerseCount(EndChapter), ExtraVerseNumber)), 1, EndWordNumber + 1) Then Exit For
                        If Matches(WordCount).Groups(2).Success Then
                            MarkerCount += 1
                            If bNotFilter Then IndexToVerseList.Add(New Integer() {If(StartChapter = -1, 1, StartChapter - If(bChapterRollback, 1, 0)) + Count, If(Count <> 0 Or BaseVerse = -1, 1, BaseVerse - If(WordNumber <= 1 And Not bBismillahPrecedes And Not (StartChapter <= 1 And BaseVerse <= 1), 1, 0)) + SubCount, WordCount - MarkerCount + 1, Str.Length + Matches(WordCount).Groups(2).Index - FilterIndex, Matches(WordCount).Groups(2).Length, ChunkCount, MarkerCount})
                            If WordCount <> 0 AndAlso Not Matches(WordCount - 1).Groups(2).Success Then ChunkCount += 1
                        Else
                            If bNotFilter Then IndexToVerseList.Add(New Integer() {If(StartChapter = -1, 1, StartChapter - If(bChapterRollback, 1, 0)) + Count, If(Count <> 0 Or BaseVerse = -1, 1, BaseVerse - If(WordNumber <= 1 And Not bBismillahPrecedes And Not (StartChapter <= 1 And BaseVerse <= 1), 1, 0)) + SubCount, WordCount - MarkerCount + 1, Str.Length + Matches(WordCount).Groups(1).Index - FilterIndex, Matches(WordCount).Groups(1).Length, ChunkCount, 0})
                        End If
                        If Not bNotFilter Then FilterIndex = Matches(WordCount).Index + Matches(WordCount).Length + 1
                    Next
                    If WordCount <> Matches.Count Then
                        Str.Append(Verses(Count)(SubCount).Substring(FilterIndex, Matches(WordCount).Groups(If(Matches(WordCount).Groups(2).Success, 2, 1)).Index - FilterIndex))
                    Else
                        Str.Append(Verses(Count)(SubCount).Substring(FilterIndex))
                        IndexToVerseList.Add(New Integer() {If(StartChapter = -1, 1, StartChapter - If(bChapterRollback, 1, 0)) + Count, If(Count <> 0 Or BaseVerse = -1, 1, BaseVerse - If(WordNumber <= 1 And Not bBismillahPrecedes And Not (StartChapter <= 1 And BaseVerse <= 1), 1, 0)) + SubCount, Matches.Count - MarkerCount, Str.Length + 1, CStr(If(BaseVerse = -1, 1, BaseVerse) + SubCount).Length + 1, ChunkCount, MarkerCount + 1})
                        Str.Append(Arb.TransliterateFromBuckwalter(" =" + CStr(If(Count <> 0 Or BaseVerse = -1, 1, BaseVerse - If(WordNumber <= 1 And Not bBismillahPrecedes And Not (StartChapter <= 1 And BaseVerse <= 1), 1, 0)) + SubCount)))
                    End If
                    Str.Append(" "c)
                Next
            End If
        Next
        If (EndChapter = -1 Or EndChapter = GetChapterCount()) And (ExtraVerseNumber = -1 Or ExtraVerseNumber = GetVerseCount(If(EndChapter = -1, GetChapterCount(), EndChapter))) And (EndWordNumber = -1 Or EndWordNumber = GetWordCount(If(EndChapter = -1, GetChapterCount(), EndChapter), GetVerseCount(If(EndChapter = -1, GetChapterCount(), EndChapter)))) Then IndexToVerseList.Add(New Integer() {GetChapterCount(), GetVerseCount(If(EndChapter = -1, GetChapterCount(), EndChapter)), GetWordCount(If(EndChapter = -1, GetChapterCount(), EndChapter), GetVerseCount(If(EndChapter = -1, GetChapterCount(), EndChapter))), 0, 0, 0, 0})
        IndexToVerse = IndexToVerseList.ToArray()
        Return Str.ToString()
    End Function
    Public Async Function GetQuranTextBySelection(ID As String, Division As Integer, Index As Integer, Translation As String, SchemeType As ArabicData.TranslitScheme, Scheme As String, TranslationIndex As Integer, W4W As Boolean, W4WNum As Boolean, NoArabic As Boolean, Header As Boolean, NoRef As Boolean, Colorize As Boolean, Verses As Boolean) As Threading.Tasks.Task(Of RenderArray)
        Dim Chapter As Integer
        Dim Verse As Integer
        Dim BaseChapter As Integer
        Dim BaseVerse As Integer
        Dim Node As Xml.Linq.XElement
        Dim Renderer As New RenderArray(ID)
        Dim QuranText As String = Nothing
        Dim IndexToVerse As Integer()() = Nothing
        Dim SeperateSectionCount As Integer = 1
        Dim Keys() As String = Nothing
        Dim Lines As String() = Await GetLines(Translation)
        Dim W4WLines As String() = Await GetW4WLines(ChData.IslamData.LanguageDefaultInfo.GetLanguageByID(Threading.Thread.CurrentThread.CurrentUICulture.TwoLetterISOLanguageName).QuranW4WFile)
        If Division = 8 Then SeperateSectionCount = ChData.IslamData.QuranSelections(Index).SelectionInfo.Length
        If Division = 9 Then
            If ChData.LetterDictionary Is Nothing Then ChData.BuildQuranLetterIndex(Me)
            ReDim Keys(ChData.LetterDictionary(ArbData.ArabicLetters(Index).Symbol).Count - 1)
            ChData.LetterDictionary(ArbData.ArabicLetters(Index).Symbol).Keys.CopyTo(Keys, 0)
            SeperateSectionCount = ChData.LetterDictionary(ArbData.ArabicLetters(Index).Symbol).Count
        End If
        For SectionCount As Integer = 0 To SeperateSectionCount - 1
            If Division = 0 Then
                BaseChapter = CInt(GetChapterByIndex(Index).Attribute("index").Value)
                BaseVerse = 1
                QuranText = QuranTextCombiner(ChData.XMLDocMain, IndexToVerse, True, BaseChapter, BaseVerse, -1, BaseChapter)
            ElseIf Division = 1 Then
                BaseChapter = CInt(GetChapterIndexByRevelationOrder(Index).Attribute("index").Value)
                BaseVerse = 1
                QuranText = QuranTextCombiner(ChData.XMLDocMain, IndexToVerse, True, BaseChapter, BaseVerse, -1, BaseChapter)
            ElseIf Division = 2 Then
                Node = GetPartByIndex(Index)
                BaseChapter = CInt(Node.Attribute("sura").Value)
                BaseVerse = CInt(Node.Attribute("aya").Value)
                Node = GetPartByIndex(Index + 1)
                If Node Is Nothing Then
                    Chapter = -1
                    Verse = -1
                Else
                    Chapter = CInt(Node.Attribute("sura").Value)
                    Verse = CInt(Node.Attribute("aya").Value)
                    GetPreviousChapterVerse(Chapter, Verse)
                End If
                QuranText = QuranTextCombiner(ChData.XMLDocMain, IndexToVerse, True, BaseChapter, BaseVerse, -1, Chapter, Verse)
            ElseIf Division = 3 Then
                Node = GetGroupByIndex(Index)
                BaseChapter = CInt(Node.Attribute("sura").Value)
                BaseVerse = CInt(Node.Attribute("aya").Value)
                Node = GetGroupByIndex(Index + 1)
                If Node Is Nothing Then
                    Chapter = -1
                    Verse = -1
                Else
                    Chapter = CInt(Node.Attribute("sura").Value)
                    Verse = CInt(Node.Attribute("aya").Value)
                    GetPreviousChapterVerse(Chapter, Verse)
                End If
                QuranText = QuranTextCombiner(ChData.XMLDocMain, IndexToVerse, True, BaseChapter, BaseVerse, -1, Chapter, Verse)
            ElseIf Division = 4 Then
                Node = GetStationByIndex(Index)
                BaseChapter = CInt(Node.Attribute("sura").Value)
                BaseVerse = CInt(Node.Attribute("aya").Value)
                Node = GetStationByIndex(Index + 1)
                If Node Is Nothing Then
                    Chapter = -1
                    Verse = -1
                Else
                    Chapter = CInt(Node.Attribute("sura").Value)
                    Verse = CInt(Node.Attribute("aya").Value)
                    GetPreviousChapterVerse(Chapter, Verse)
                End If
                QuranText = QuranTextCombiner(ChData.XMLDocMain, IndexToVerse, True, BaseChapter, BaseVerse, -1, Chapter, Verse)
            ElseIf Division = 5 Then
                Node = GetSectionByIndex(Index)
                BaseChapter = CInt(Node.Attribute("sura").Value)
                BaseVerse = CInt(Node.Attribute("aya").Value)
                Node = GetSectionByIndex(Index + 1)
                If Node Is Nothing Then
                    Chapter = -1
                    Verse = -1
                Else
                    Chapter = CInt(Node.Attribute("sura").Value)
                    Verse = CInt(Node.Attribute("aya").Value)
                    GetPreviousChapterVerse(Chapter, Verse)
                End If
                QuranText = QuranTextCombiner(ChData.XMLDocMain, IndexToVerse, True, BaseChapter, BaseVerse, -1, Chapter, Verse)
            ElseIf Division = 6 Then
                Node = GetPageByIndex(Index)
                BaseChapter = CInt(Node.Attribute("sura").Value)
                BaseVerse = CInt(Node.Attribute("aya").Value)
                Node = GetPageByIndex(Index + 1)
                If Node Is Nothing Then
                    Chapter = -1
                    Verse = -1
                Else
                    Chapter = CInt(Node.Attribute("sura").Value)
                    Verse = CInt(Node.Attribute("aya").Value)
                    GetPreviousChapterVerse(Chapter, Verse)
                End If
                QuranText = QuranTextCombiner(ChData.XMLDocMain, IndexToVerse, True, BaseChapter, BaseVerse, -1, Chapter, Verse)
            ElseIf Division = 7 Then
                Node = GetSajdaByIndex(Index)
                BaseChapter = CInt(Node.Attribute("sura").Value)
                BaseVerse = CInt(Node.Attribute("aya").Value)
                QuranText = QuranTextCombiner(ChData.XMLDocMain, IndexToVerse, True, BaseChapter, BaseVerse, -1, BaseChapter, BaseVerse)
            ElseIf Division = 8 Then
                BaseChapter = ChData.IslamData.QuranSelections(Index).SelectionInfo(SectionCount).ChapterNumber
                BaseVerse = ChData.IslamData.QuranSelections(Index).SelectionInfo(SectionCount).VerseNumber
                QuranText = QuranTextCombiner(ChData.XMLDocMain, IndexToVerse, True, BaseChapter, BaseVerse, ChData.IslamData.QuranSelections(Index).SelectionInfo(SectionCount).WordNumber, 0, ChData.IslamData.QuranSelections(Index).SelectionInfo(SectionCount).ExtraVerseNumber, ChData.IslamData.QuranSelections(Index).SelectionInfo(SectionCount).EndWordNumber)
            ElseIf Division = 9 Then
                For SubCount = 0 To ChData.LetterDictionary(ArbData.ArabicLetters(Index).Symbol)(Keys(SectionCount)).Count - 1
                    BaseChapter = ChData.LetterDictionary(ArbData.ArabicLetters(Index).Symbol)(Keys(SectionCount))(SubCount)(0) + 1
                    BaseVerse = ChData.LetterDictionary(ArbData.ArabicLetters(Index).Symbol)(Keys(SectionCount))(SubCount)(1) + 1
                    QuranText = QuranTextCombiner(ChData.XMLDocMain, IndexToVerse, True, BaseChapter, BaseVerse, -1, BaseChapter, BaseVerse)
                    Renderer.Items.AddRange(DoGetRenderedQuranText(QuranText, IndexToVerse, {Lines}, {W4WLines}, SchemeType, Scheme, {ChData.IslamData.Translations.TranslationList(TranslationIndex).FileName.Substring(0, 2)}, {Languages.GetLanguageInfoByCode(Threading.Thread.CurrentThread.CurrentUICulture.TwoLetterISOLanguageName, ChData.IslamData.LanguageList).IsRTL}, W4W, W4WNum, NoArabic, Header, Colorize, Verses, NoRef).Items)
                    QuranText = Nothing
                    IndexToVerse = Nothing
                Next
            ElseIf Division = 10 Then
                Dim Text As String = QuranTextCombiner(ChData.XMLDocMain, IndexToVerse)
                Dim Matches As System.Text.RegularExpressions.MatchCollection = System.Text.RegularExpressions.Regex.Matches(Text, ChData.RuleMetas("UthmaniQuran").Rules(Index).Match)
                Renderer.Items.AddRange(New RenderArray.RenderItem() {New RenderArray.RenderItem(RenderArray.RenderTypes.eText, New RenderArray.RenderText() {New RenderArray.RenderText(RenderArray.RenderDisplayClass.eLTR, ChData.RuleMetas("UthmaniQuran").Rules(Index).Name)}), New RenderArray.RenderItem(RenderArray.RenderTypes.eText, DocBuilder.ColorizeRegExGroups(DocBuilder.GetRegExText(ChData.RuleMetas("UthmaniQuran").Rules(Index).Match), False)), New RenderArray.RenderItem(RenderArray.RenderTypes.eText, DocBuilder.ColorizeList(New List(Of String)(Linq.Enumerable.Select(ChData.RuleMetas("UthmaniQuran").Rules(Index).Evaluator, Function(Str As IslamData.RuleMetaSet.RuleMetadataTranslation.RuleWithArgs()) DocBuilder.GetRegExText(String.Join("|"c, Linq.Enumerable.Select(Str, Function(S) S.RuleName + If(S.Args Is Nothing OrElse S.Args.Length = 0, "", "(" + String.Join(",", Linq.Enumerable.Select(S.Args, Function(Arg) String.Join(" ", Arg))) + ")")))))).ToArray(), False))})
                For SubCount = 0 To Matches.Count - 1
                    Dim StartWordIndex As Integer = Array.BinarySearch(IndexToVerse, Matches(SubCount).Index, New QuranWordIndexComparer)
                    If StartWordIndex < 0 Then StartWordIndex = (StartWordIndex Xor -1) - 1
                    Dim EndWordIndex As Integer = Array.BinarySearch(IndexToVerse, Matches(SubCount).Index + Matches(SubCount).Length - 1, New QuranWordIndexComparer)
                    If EndWordIndex < 0 Then EndWordIndex = (EndWordIndex Xor -1) + 1
                    Dim Renderers As New List(Of RenderArray.RenderText)
                    Renderers.Add(New RenderArray.RenderText(RenderArray.RenderDisplayClass.eReference, New Integer() {IndexToVerse(StartWordIndex)(0), IndexToVerse(StartWordIndex)(1)}))
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
                IndexToVerse = Nothing
            Else
                QuranText = Nothing
            End If
            If Not QuranText Is Nothing Then Renderer.Items.AddRange(DoGetRenderedQuranText(QuranText, IndexToVerse, {Lines}, {W4WLines}, SchemeType, Scheme, {ChData.IslamData.Translations.TranslationList(TranslationIndex).FileName.Substring(0, 2)}, {Languages.GetLanguageInfoByCode(Threading.Thread.CurrentThread.CurrentUICulture.TwoLetterISOLanguageName, ChData.IslamData.LanguageList).IsRTL}, W4W, W4WNum, NoArabic, Header, Colorize, Verses, NoRef).Items)
        Next
        Return Renderer
    End Function
    Public Class QuranWordChapterVerseWordComparer
        Implements IComparer(Of Integer())
        Dim bUseChunk As Boolean
        Public Sub New(UseChunk As Boolean)
            bUseChunk = UseChunk
        End Sub
        Public Function Compare(x As Integer(), y As Integer()) As Integer Implements IComparer(Of Integer()).Compare
            If x(0) > y(0) Then Return 1
            If x(0) < y(0) Then Return -1
            If x(1) > y(1) Then Return 1
            If x(1) < y(1) Then Return -1
            If bUseChunk And x.Length = 6 Then
                If x(2) > y(5) Then Return 1
                If x(2) < y(5) Then Return -1
            ElseIf bUseChunk And y.Length = 6 Then
                If x(5) > y(2) Then Return 1
                If x(5) < y(2) Then Return -1
            End If
            If x(2) > y(2) Then Return 1
            If x(2) < y(2) Then Return -1
            If x.Length = 6 Then
                If y(6) > 0 Then Return 1
            ElseIf y.Length = 6 Then
                If x(6) > 0 Then Return 1
            ElseIf x(6) > y(6) Then
                Return 1
            ElseIf x(6) < y(6) Then
                Return -1
            End If
            Return 0
        End Function
    End Class
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
    Public Async Function GetRenderedQuranText(SchemeType As ArabicData.TranslitScheme, Scheme As String, Name As String, Strings As String, QuranSelection As String, QuranTranslation As String, WordVerseMode As String, ColorCueMode As String) As Threading.Tasks.Task(Of RenderArray)
        Dim Division As Integer = 0
        Dim Index As Integer = 1
        If Not Strings Is Nothing Then Division = CInt(Strings)
        If Not QuranSelection Is Nothing Then Index = CInt(QuranSelection)
        Dim TranslationIndex As Integer = GetTranslationIndex(QuranTranslation)
        Return Await GetQuranTextBySelection(Name, Division, Index, If(TranslationIndex = -1, String.Empty, ChData.IslamData.Translations.TranslationList(TranslationIndex).FileName), SchemeType, Scheme, TranslationIndex, CInt(WordVerseMode) <> 4, CInt(WordVerseMode) Mod 2 = 1, False, True, False, CInt(ColorCueMode) = 0, CInt(WordVerseMode) / 2 <> 1)
    End Function
    Public Function GenerateDefaultStops(Str As String) As Integer()
        Dim Matches As System.Text.RegularExpressions.MatchCollection = System.Text.RegularExpressions.Regex.Matches(Str, "(^\s*|\s+)(" + ArabicData.MakeUniRegEx(ArabicData.ArabicEndOfAyah) + "[\p{Nd}]{1,3}|" + ChData.OptionalPattern + ")(?=\s*$|\s+)")
        Dim DefStops As New List(Of Integer)
        Dim DotToggle As Boolean = False
        For Count As Integer = 0 To Matches.Count - 1
            If Matches(Count).Groups(2).Value <> ArabicData.ArabicSmallHighLigatureSadWithLamWithAlefMaksura AndAlso (Matches(Count).Groups(2).Value <> ArabicData.ArabicSmallHighThreeDots Or DotToggle) Then DefStops.Add(Matches(Count).Groups(2).Index)
            If Matches(Count).Groups(2).Value = ArabicData.ArabicSmallHighThreeDots Then DotToggle = Not DotToggle
        Next
        Return DefStops.ToArray()
    End Function
    Public Async Function FileToResource(FilePath As String, ResFilePath As String) As Threading.Tasks.Task
        Dim Lines As String() = New List(Of String)(Linq.Enumerable.Where(Await _PortableMethods.ReadAllLines(FilePath), Function(Item As String) Not Item.StartsWith("#") And Item <> String.Empty)).ToArray()
        Dim Stream As IO.Stream = Await _PortableMethods.FileIO.LoadStream(ResFilePath)
        Dim XMLDoc As Xml.Linq.XDocument = Xml.Linq.XDocument.Load(Stream)
        Stream.Dispose()
        For Count = 0 To Lines.Length - 1
            Dim NewData As New Xml.Linq.XElement("data")
            NewData.Add(New Xml.Linq.XAttribute("name", "Quran" + CStr(Count + 1)))
            NewData.Add(New Xml.Linq.XAttribute(Xml.Linq.XNamespace.Xml + "space", "preserve"))
            Dim Inner As New Xml.Linq.XElement("value")
            Inner.Value = Lines(Count)
            NewData.Add(Inner)
            XMLDoc.Root.Add(NewData)
        Next
        Dim MemStream As New IO.MemoryStream
        XMLDoc.Save(MemStream)
        Await _PortableMethods.FileIO.SaveStream(ResFilePath, MemStream)
        MemStream.Dispose()
    End Function
    Public Async Function ResourceToFile(ResFilePath As String, FilePath As String) As Threading.Tasks.Task
        Dim Lines As New List(Of String)
        Dim Stream As IO.Stream = Await _PortableMethods.FileIO.LoadStream(ResFilePath)
        Dim XMLDoc As Xml.Linq.XDocument = Xml.Linq.XDocument.Load(Stream)
        Stream.Dispose()
        Dim AllNodes As Xml.Linq.XElement() = (New List(Of Xml.Linq.XElement)(Linq.Enumerable.Where(XMLDoc.Root.Elements, Function(elem) elem.Name = "data" And Not elem.Attribute("name") Is Nothing))).ToArray()
        For Each Item As Xml.Linq.XElement In AllNodes
            If System.Text.RegularExpressions.Regex.Match(Item.Attribute("name").Value, "^Quran\d+$").Success Then
                Dim Line As Integer = Integer.Parse(Item.Attribute("name").Value.Substring(5))
                Lines.Insert(Line - 1, New List(Of Xml.Linq.XElement)(Item.Elements).Item(0).Value)
            End If
        Next
        Await _PortableMethods.WriteAllLines(FilePath, Lines.ToArray())
    End Function
    Public Async Function WordFileToResource(WordFilePath As String, ResFilePath As String) As Threading.Tasks.Task
        Dim W4WLines As String() = Await _PortableMethods.ReadAllLines(WordFilePath)
        Dim Stream As IO.Stream = Await _PortableMethods.FileIO.LoadStream(ResFilePath)
        Dim XMLDoc As Xml.Linq.XDocument = Xml.Linq.XDocument.Load(Stream)
        Stream.Dispose()
        For Count = 0 To W4WLines.Length - 1
            Dim Words As String() = W4WLines(Count).Split("|"c)
            For SubCount = 0 To Words.Length - 1
                Dim NewData As New Xml.Linq.XElement("data")
                NewData.Add(New Xml.Linq.XAttribute("name", "Quran" + CStr(Count + 1) + "." + CStr(SubCount + 1)))
                NewData.Add(New Xml.Linq.XAttribute(Xml.Linq.XNamespace.Xml + "space", "preserve"))
                Dim Inner As New Xml.Linq.XElement("value")
                Inner.Value = Words(SubCount)
                NewData.Add(Inner)
                XMLDoc.Root.Add(NewData)
            Next
        Next
        Dim MemStream As New IO.MemoryStream
        XMLDoc.Save(MemStream)
        Await _PortableMethods.FileIO.SaveStream(ResFilePath, MemStream)
        MemStream.Dispose()
    End Function
    Public Async Function ResourceToWordFile(ResFilePath As String, WordFilePath As String) As Threading.Tasks.Task
        Dim W4WLines As New List(Of List(Of String))
        Dim Stream As IO.Stream = Await _PortableMethods.FileIO.LoadStream(ResFilePath)
        Dim XMLDoc As Xml.Linq.XDocument = Xml.Linq.XDocument.Load(Stream)
        Stream.Dispose()
        Dim AllNodes As Xml.Linq.XElement() = (New List(Of Xml.Linq.XElement)(Linq.Enumerable.Where(XMLDoc.Root.Elements, Function(elem) elem.Name = "data" And Not elem.Attribute("name") Is Nothing))).ToArray()
        For Each Item As Xml.Linq.XElement In AllNodes
            If System.Text.RegularExpressions.Regex.Match(Item.Attribute("name").Value, "^Quran\d+\.\d+$").Success Then
                Dim Line As Integer = Integer.Parse(Item.Attribute("name").Value.Substring(5, Item.Attribute("name").Value.IndexOf("."c) - 5))
                Dim Word As Integer = Integer.Parse(Item.Attribute("name").Value.Substring(Item.Attribute("name").Value.IndexOf("."c) + 1))
                If W4WLines.Count < Line OrElse W4WLines(Line - 1) Is Nothing Then
                    W4WLines.Insert(Line - 1, New List(Of String))
                End If
                W4WLines(Line - 1).Insert(Word - 1, New List(Of Xml.Linq.XElement)(Item.Elements).Item(0).Value)
            End If
        Next
        Await _PortableMethods.WriteAllLines(WordFilePath, New List(Of String)(Linq.Enumerable.Select(W4WLines, Function(Input As List(Of String)) String.Join("|"c, Input.ToArray()))).ToArray())
    End Function
    Public Function DoGetRenderedQuranTextW4W(Text As String, IndexToVerse As Integer()(), MetaRules As Generic.List(Of Arabic.RuleMetadata), UnfilteredMetaRules As Generic.List(Of Arabic.RuleMetadata), DefStops As Integer(), W4WLines As String()(), SchemeType As ArabicData.TranslitScheme, Scheme As String, IsLTRW4WTranslation As Boolean(), W4WNum As Boolean, NoArabic As Boolean, Colorize As Boolean, NoRef As Boolean) As RenderArray.RenderText
        Dim Texts As New List(Of RenderArray.RenderText)
        Dim Items As New Collections.Generic.List(Of RenderArray.RenderItem)
        If NoRef Then
            If Not NoArabic Then
                Texts.Add(New RenderArray.RenderText(RenderArray.RenderDisplayClass.eArabic, Arb.TransliterateFromBuckwalter("(")))
                Texts.Add(New RenderArray.RenderText(RenderArray.RenderDisplayClass.eTransliteration, If(SchemeType <> ArabicData.TranslitScheme.None, Arb.TransliterateToScheme(Arb.TransliterateFromBuckwalter(")"), SchemeType, Scheme, Arabic.FilterMetadataStops(Arb.TransliterateFromBuckwalter(")"), Arb.GetMetarules(Arb.TransliterateFromBuckwalter(")"), ChData.RuleMetas("Normal")), Nothing)), String.Empty)))
            End If
            For TranslationIndex As Integer = 0 To W4WLines.Length - 1
                Texts.Add(New RenderArray.RenderText(If(IsLTRW4WTranslation(TranslationIndex), RenderArray.RenderDisplayClass.eLTR, RenderArray.RenderDisplayClass.eRTL), ")"))
            Next
            Items.Add(New RenderArray.RenderItem(RenderArray.RenderTypes.eText, Texts.ToArray()))
            Texts.Clear()
        End If
        Dim TranslitWords As String() = Nothing
        Dim WordColors As RenderArray.RenderText()() = Nothing
        Dim TranslitColors As RenderArray.RenderText()() = Nothing
        If Colorize Then
            WordColors = Arb.ApplyColorRules(Text, True, MetaRules)
            TranslitColors = Arb.TransliterateWithRulesColor(Text, Scheme, True, False, MetaRules)
        Else
            TranslitWords = Arb.TransliterateToScheme(Text, SchemeType, Scheme, MetaRules).Split(" "c)
        End If
        Dim Pos As Integer = 0
        For Count As Integer = 0 To IndexToVerse.Length - 1
            If IndexToVerse(Count)(4) = 1 AndAlso
                    (Arb.IsStop(ArbData.FindLetterBySymbol(Text(IndexToVerse(Count)(0)))) Or Text(IndexToVerse(Count)(0)) = ArabicData.ArabicStartOfRubElHizb Or Text(IndexToVerse(Count)(0)) = ArabicData.ArabicPlaceOfSajdah) Then
                If Not NoArabic Then
                    Texts.Add(New RenderArray.RenderText(RenderArray.RenderDisplayClass.eArabic, " " + Text.Substring(IndexToVerse(Count)(3), IndexToVerse(Count)(4))))
                    Texts.Add(New RenderArray.RenderText(RenderArray.RenderDisplayClass.eArabic, " "))
                    'Texts.Add(New RenderArray.RenderText(RenderArray.RenderDisplayClass.eTransliteration, If(SchemeType <> ArabicData.TranslitScheme.None, TranslitWords(Count), String.Empty)))
                    Texts.Add(New RenderArray.RenderText(RenderArray.RenderDisplayClass.eContinueStop, New Object() {Array.IndexOf(DefStops, Pos) <> -1, Arabic.FilterMetadataStopsContig(Pos + IndexToVerse(Count)(4) + If(Count <> IndexToVerse(IndexToVerse.Length - 1)(2) AndAlso IndexToVerse(Count + 1)(4) <> 1, 1 + IndexToVerse(Count + 1)(4), 0), UnfilteredMetaRules, DefStops, Pos - If(Count = 0, 0, IndexToVerse(Count - 1)(4) + 1))}))
                End If
                If W4WNum Then Texts.Add(New RenderArray.RenderText(RenderArray.RenderDisplayClass.eLTR, CStr(Count + 1)))
                Items.Add(New RenderArray.RenderItem(RenderArray.RenderTypes.eText, Texts.ToArray()))
                Texts.Clear()
            ElseIf IndexToVerse(Count)(4) <> 0 Then
                If Text(IndexToVerse(Count)(3)) <> ArabicData.ArabicEndOfAyah Then
                    Texts.Add(New RenderArray.RenderText(RenderArray.RenderDisplayClass.eReference, New Integer() {IndexToVerse(Count)(0), IndexToVerse(Count)(1), IndexToVerse(Count)(2)}))
                End If
                If Not NoArabic Then
                    If Colorize Then
                        Texts.AddRange(WordColors(Count))
                        If SchemeType <> ArabicData.TranslitScheme.None Then
                            Texts.AddRange(TranslitColors(IndexToVerse(Count)(2) - If(Text(IndexToVerse(Count)(3)) = ArabicData.ArabicEndOfAyah, 0, 1)))
                        Else
                            Texts.Add(New RenderArray.RenderText(RenderArray.RenderDisplayClass.eTransliteration, String.Empty))
                        End If
                    Else
                        Texts.Add(New RenderArray.RenderText(RenderArray.RenderDisplayClass.eArabic, Text.Substring(IndexToVerse(Count)(3), IndexToVerse(Count)(4))))
                        Texts.Add(New RenderArray.RenderText(RenderArray.RenderDisplayClass.eTransliteration, If(SchemeType <> ArabicData.TranslitScheme.None, TranslitWords(IndexToVerse(Count)(2) - If(Text(IndexToVerse(Count)(3)) = ArabicData.ArabicEndOfAyah, 0, 1)), String.Empty)))
                    End If
                    If Text(IndexToVerse(Count)(3)) = ArabicData.ArabicEndOfAyah And Count <> 0 Then
                        Texts.Add(New RenderArray.RenderText(RenderArray.RenderDisplayClass.eContinueStop, New Object() {Array.IndexOf(DefStops, Pos) <> -1, Arabic.FilterMetadataStopsContig(Pos + IndexToVerse(Count)(4) + If(Count <> IndexToVerse(IndexToVerse.Length - 1)(2) AndAlso IndexToVerse(Count + 1)(4) <> 1, 1 + IndexToVerse(Count + 1)(4), 0), UnfilteredMetaRules, DefStops, Pos - IndexToVerse(Count - 1)(4) - 1)}))
                    End If
                End If
                'If Translation <> String.Empty Then
                If Text(IndexToVerse(Count)(3)) = ArabicData.ArabicEndOfAyah Then
                    For TranslationIndex As Integer = 0 To W4WLines.Length - 1
                        Texts.Add(New RenderArray.RenderText(If(IsLTRW4WTranslation(TranslationIndex), RenderArray.RenderDisplayClass.eLTR, RenderArray.RenderDisplayClass.eRTL), "(" + CStr(IndexToVerse(Count)(1)) + ")"))
                    Next
                Else
                    For TranslationIndex As Integer = 0 To W4WLines.Length - 1
                        Texts.Add(New RenderArray.RenderText(If(IsLTRW4WTranslation(TranslationIndex), RenderArray.RenderDisplayClass.eLTR, RenderArray.RenderDisplayClass.eRTL), GetW4WTranslationVerse(W4WLines(TranslationIndex), IndexToVerse(Count)(0), IndexToVerse(Count)(1), IndexToVerse(Count)(2))))
                    Next
                End If
                'End If
                If W4WNum Then Texts.Add(New RenderArray.RenderText(RenderArray.RenderDisplayClass.eLTR, CStr(IndexToVerse(Count)(2))))
                Items.Add(New RenderArray.RenderItem(RenderArray.RenderTypes.eText, Texts.ToArray()))
                If Not NoArabic AndAlso Count <> IndexToVerse(IndexToVerse.Length - 1)(2) AndAlso IndexToVerse(Count + 1)(4) <> 1 AndAlso Text(IndexToVerse(Count + 1)(3)) <> ArabicData.ArabicEndOfAyah Then
                    Items.Add(New RenderArray.RenderItem(RenderArray.RenderTypes.eText, {New RenderArray.RenderText(RenderArray.RenderDisplayClass.eContinueStop, New Object() {Array.IndexOf(DefStops, Pos + IndexToVerse(Count)(4)) <> -1, Arabic.FilterMetadataStopsContig(Pos + IndexToVerse(Count)(4) + 1 + IndexToVerse(Count + 1)(4), UnfilteredMetaRules, DefStops, Pos)})}))
                End If
                Texts.Clear()
            End If
            Pos += IndexToVerse(Count)(4) + 1
        Next
        If NoRef Then
            If Not NoArabic Then
                Texts.Add(New RenderArray.RenderText(RenderArray.RenderDisplayClass.eArabic, Arb.TransliterateFromBuckwalter(")")))
                Texts.Add(New RenderArray.RenderText(RenderArray.RenderDisplayClass.eTransliteration, If(SchemeType <> ArabicData.TranslitScheme.None, Arb.TransliterateToScheme(Arb.TransliterateFromBuckwalter(")"), SchemeType, Scheme, Arabic.FilterMetadataStops(Arb.TransliterateFromBuckwalter("("), Arb.GetMetarules(Arb.TransliterateFromBuckwalter("("), ChData.RuleMetas("Normal")), Nothing)), String.Empty)))
            End If
            For TranslationIndex As Integer = 0 To W4WLines.Length - 1
                Texts.Add(New RenderArray.RenderText(If(IsLTRW4WTranslation(TranslationIndex), RenderArray.RenderDisplayClass.eLTR, RenderArray.RenderDisplayClass.eRTL), "("))
            Next
            Items.Add(New RenderArray.RenderItem(RenderArray.RenderTypes.eText, Texts.ToArray()))
            Texts.Clear()
        End If
        Return New RenderArray.RenderText(RenderArray.RenderDisplayClass.eNested, Items)
    End Function
    Public Function DoGetRenderedQuranText(QuranText As String, IndexToVerse As Integer()(), Lines As String()(), W4WLines As String()(), SchemeType As ArabicData.TranslitScheme, Scheme As String, TranslationLang As String(), IsLTRW4WTranslation As Boolean(), W4W As Boolean, W4WNum As Boolean, NoArabic As Boolean, Header As Boolean, Colorize As Boolean, Verses As Boolean, NoRef As Boolean) As RenderArray
        Dim Renderer As New RenderArray(String.Empty)
        Dim Index As Integer = If(_CacheMetarules Is Nothing, 1, 0)
        Dim First As Integer = -1
        Dim Last As Integer = -1
        Dim UnfilteredMetaRules As Generic.List(Of Arabic.RuleMetadata) = Nothing
        Dim DefStops As Integer() = Nothing
        While Index <> IndexToVerse.Length - 1 - If(_CacheMetarules Is Nothing, 1, 0) AndAlso IndexToVerse(Index)(6) <> 0
            Index += 1
        End While
        First = Index
        If Not NoArabic Then
            DefStops = GenerateDefaultStops(QuranText)
            UnfilteredMetaRules = If(_CacheMetarules Is Nothing, Arb.GetMetarules(QuranText, ChData.RuleMetas("UthmaniQuran")), Arabic.GetMetarulesFromCache(QuranText, IndexToVerse(First)(3), _CacheMetarules))
        End If
        While Index <> IndexToVerse.Length - 1 - If(_CacheMetarules Is Nothing, 1, 0)
            Dim ChapterNode As Xml.Linq.XElement = GetChapterByIndex(IndexToVerse(Index)(0))
            Dim Texts As New List(Of RenderArray.RenderText)
            If Header Then
                If Not NoArabic Then
                    Dim Str As String = Arb.TransliterateFromBuckwalter(ChData.QuranHeaders(0) + " " + ChapterNode.Attribute("ayas").Value)
                    Texts.Add(New RenderArray.RenderText(RenderArray.RenderDisplayClass.eArabic, Str))
                    Texts.Add(New RenderArray.RenderText(RenderArray.RenderDisplayClass.eTransliteration, If(SchemeType <> ArabicData.TranslitScheme.None, Arb.TransliterateToScheme(Str, SchemeType, Scheme, Arabic.FilterMetadataStops(Str, Arb.GetMetarules(Str, ChData.RuleMetas("Normal")), Nothing)).Trim(), String.Empty)))
                End If
                Texts.Add(New RenderArray.RenderText(If(Languages.GetLanguageInfoByCode(Threading.Thread.CurrentThread.CurrentUICulture.TwoLetterISOLanguageName, ChData.IslamData.LanguageList).IsRTL, RenderArray.RenderDisplayClass.eRTL, RenderArray.RenderDisplayClass.eLTR), _PortableMethods.LoadResourceString("IslamInfo_Verses") + " " + ChapterNode.Attribute("ayas").Value + " "))
                Renderer.Items.Add(New RenderArray.RenderItem(RenderArray.RenderTypes.eHeaderLeft, Texts.ToArray()))
                Texts.Clear()
                If Not NoArabic Then
                    Dim Str As String = Arb.TransliterateFromBuckwalter(ChData.QuranHeaders(1) + " " + ChData.IslamData.QuranChapters(CInt(ChapterNode.Attribute("index").Value) - 1).Name)
                    Texts.Add(New RenderArray.RenderText(RenderArray.RenderDisplayClass.eArabic, Str))
                    Texts.Add(New RenderArray.RenderText(RenderArray.RenderDisplayClass.eTransliteration, If(SchemeType <> ArabicData.TranslitScheme.None, Arb.TransliterateToScheme(Str, SchemeType, Scheme, Arabic.FilterMetadataStops(Str, Arb.GetMetarules(Str, ChData.RuleMetas("Normal")), Nothing)).Trim(), String.Empty)))
                End If
                Texts.Add(New RenderArray.RenderText(If(Languages.GetLanguageInfoByCode(Threading.Thread.CurrentThread.CurrentUICulture.TwoLetterISOLanguageName, ChData.IslamData.LanguageList).IsRTL, RenderArray.RenderDisplayClass.eRTL, RenderArray.RenderDisplayClass.eLTR), _PortableMethods.LoadResourceString("IslamInfo_Chapter") + " " + GetChapterEName(ChapterNode) + " "))
                Renderer.Items.Add(New RenderArray.RenderItem(RenderArray.RenderTypes.eHeaderCenter, Texts.ToArray()))
                Texts.Clear()
                If Not NoArabic Then
                    Dim Str As String = Arb.TransliterateFromBuckwalter(ChData.QuranHeaders(2) + " " + ChapterNode.Attribute("rukus").Value)
                    Texts.Add(New RenderArray.RenderText(RenderArray.RenderDisplayClass.eArabic, Str))
                    Texts.Add(New RenderArray.RenderText(RenderArray.RenderDisplayClass.eTransliteration, If(SchemeType <> ArabicData.TranslitScheme.None, Arb.TransliterateToScheme(Str, SchemeType, Scheme, Arabic.FilterMetadataStops(Str, Arb.GetMetarules(Str, ChData.RuleMetas("Normal")), Nothing)).Trim(), String.Empty)))
                End If
                Texts.Add(New RenderArray.RenderText(If(Languages.GetLanguageInfoByCode(Threading.Thread.CurrentThread.CurrentUICulture.TwoLetterISOLanguageName, ChData.IslamData.LanguageList).IsRTL, RenderArray.RenderDisplayClass.eRTL, RenderArray.RenderDisplayClass.eLTR), _PortableMethods.LoadResourceString("IslamInfo_Rukus") + " " + ChapterNode.Attribute("rukus").Value + " "))
                Renderer.Items.Add(New RenderArray.RenderItem(RenderArray.RenderTypes.eHeaderRight, Texts.ToArray()))
                Texts.Clear()
            End If
            Do
                'hizb symbols not needed as Quranic text already contains them
                'If BaseChapter + Chapter <> 1 AndAlso TanzilReader.IsQuarterStart(BaseChapter + Chapter, CInt(If(Chapter = 0, BaseVerse, 1)) + Verse) Then
                '    Text += Arabic.TransliterateFromBuckwalter("B")
                '    Items.Add(New RenderArray.RenderItem(RenderArray.RenderTypes.eText, New RenderArray.RenderText() {New RenderArray.RenderText(RenderArray.RenderDisplayClass.eArabic, Arabic.TransliterateFromBuckwalter("B"))}))
                'End If
                Texts.Add(New RenderArray.RenderText(RenderArray.RenderDisplayClass.eReference, New Integer() {IndexToVerse(Index)(0), IndexToVerse(Index)(1)}))
                If IsSajda(IndexToVerse(Index)(0), IndexToVerse(Index)(1)) Then
                    'Sajda markers are already in the text
                    'Text += Arabic.TransliterateFromBuckwalter("R")
                    'Items.Add(New RenderArray.RenderItem(RenderArray.RenderTypes.eText, New RenderArray.RenderText() {New RenderArray.RenderText(RenderArray.RenderDisplayClass.eArabic, Arabic.TransliterateFromBuckwalter("R"))}))
                End If
                Dim MetaRules As Generic.List(Of Arabic.RuleMetadata) = Nothing
                Dim AdjDefStops As Integer() = Nothing
                Dim Text As String = String.Empty
                If Not NoArabic Then
                    Dim Ints As Integer() = Linq.Enumerable.FirstOrDefault(Linq.Enumerable.Skip(If(_CacheMetarules Is Nothing, Linq.Enumerable.Take(IndexToVerse, IndexToVerse.Length - 1), IndexToVerse), Index), Function(Elem As Integer()) Elem(0) <> IndexToVerse(Index)(0) Or Elem(1) <> IndexToVerse(Index)(1))
                    If Not _CacheMetarules Is Nothing And (Ints Is Nothing OrElse Ints.Length = 0) Then
                        Text = QuranText.Substring(IndexToVerse(Index)(3) - IndexToVerse(0)(3)).Trim(" "c)
                    Else
                        If Ints Is Nothing OrElse Ints.Length = 0 Then Ints = IndexToVerse(IndexToVerse.Length - 1)
                        Text = QuranText.Substring(IndexToVerse(Index)(3) - IndexToVerse(0)(3), Ints(3) - IndexToVerse(Index)(3)).Trim(" "c)
                    End If
                    MetaRules = Arabic.FilterMetadataStopsContig(IndexToVerse(Index)(3) + Text.Length, UnfilteredMetaRules, DefStops, IndexToVerse(Index)(3))
                    AdjDefStops = New List(Of Integer)(Linq.Enumerable.Select(Linq.Enumerable.Where(DefStops, Function(It) It >= IndexToVerse(Index)(3) And It < IndexToVerse(Index)(3) + Text.Length), Function(It) It - IndexToVerse(Index)(3))).ToArray()
                    MetaRules = Arabic.FilterMetadataStops(Text, MetaRules, AdjDefStops)
                End If
                If W4W And W4WLines.Length <> 0 Then
                    Texts.Add(DoGetRenderedQuranTextW4W(Text, New List(Of Integer())(Linq.Enumerable.Select(Linq.Enumerable.TakeWhile(Linq.Enumerable.Skip(If(_CacheMetarules Is Nothing, Linq.Enumerable.Take(IndexToVerse, IndexToVerse.Length - 1), IndexToVerse), Index), Function(Elem As Integer()) Elem(0) = IndexToVerse(Index)(0) And Elem(1) = IndexToVerse(Index)(1)), Function(It) New Integer() {It(0), It(1), It(2), It(3) - IndexToVerse(Index)(3), It(4), It(5), It(6)})).ToArray(), MetaRules, UnfilteredMetaRules, AdjDefStops, W4WLines, SchemeType, Scheme, IsLTRW4WTranslation, W4WNum, NoArabic, Colorize, NoRef))
                End If
                Last = Index
                Do
                    Index += 1
                Loop While Index <> IndexToVerse.Length - 1 - If(_CacheMetarules Is Nothing, 1, 0) AndAlso (IndexToVerse(Index)(0) = IndexToVerse(Index - 1)(0) And IndexToVerse(Index)(1) = IndexToVerse(Index - 1)(1))
                If Verses Then
                    If Not NoArabic Then
                        If Colorize Then
                            Texts.AddRange(Arb.ApplyColorRules(Text, False, MetaRules)(0))
                            If SchemeType <> ArabicData.TranslitScheme.None Then
                                Texts.AddRange(Arb.TransliterateWithRulesColor(Text, Scheme, False, False, MetaRules)(0))
                            Else
                                Texts.Add(New RenderArray.RenderText(RenderArray.RenderDisplayClass.eTransliteration, String.Empty))
                            End If
                        Else
                            Texts.Add(New RenderArray.RenderText(RenderArray.RenderDisplayClass.eArabic, Text))
                            Texts.Add(New RenderArray.RenderText(RenderArray.RenderDisplayClass.eTransliteration, If(SchemeType <> ArabicData.TranslitScheme.None, Arb.TransliterateToScheme(Text, SchemeType, Scheme, MetaRules).Trim(), String.Empty)))
                        End If
                    End If
                    For TranslationIndex As Integer = 0 To Lines.Length - 1
                        Texts.Add(New RenderArray.RenderText(If(Languages.GetLanguageInfoByCode(TranslationLang(TranslationIndex), ChData.IslamData.LanguageList).IsRTL, RenderArray.RenderDisplayClass.eRTL, RenderArray.RenderDisplayClass.eLTR), If(NoRef And (Index = First Or Index = IndexToVerse.Length - 1 - If(_CacheMetarules Is Nothing, 1, 0)), String.Empty, "(" + CStr(IndexToVerse(Last)(1)) + ") ") + GetTranslationVerse(Lines(TranslationIndex), IndexToVerse(Last)(0), IndexToVerse(Last)(1))))
                    Next
                End If
                If Not NoArabic And Not (NoRef And (Index = First Or Index = IndexToVerse.Length - 1 - If(_CacheMetarules Is Nothing, 1, 0))) Then Renderer.Items.Add(New RenderArray.RenderItem(RenderArray.RenderTypes.eText, {New RenderArray.RenderText(RenderArray.RenderDisplayClass.eContinueStop, New Object() {Array.IndexOf(DefStops, IndexToVerse(Last)(3) + Text.LastIndexOf(" "c, Text.Length - 2) + 1) <> -1, Arabic.FilterMetadataStopsContig(IndexToVerse(Last)(3) + Text.Length, UnfilteredMetaRules, DefStops, IndexToVerse(Last)(3) + Text.LastIndexOf(" "c, Text.Length - 2) + 1)})}))
                Renderer.Items.Add(New RenderArray.RenderItem(RenderArray.RenderTypes.eText, Texts.ToArray()))
                Texts.Clear()
            Loop While Index <> IndexToVerse.Length - 1 - If(_CacheMetarules Is Nothing, 1, 0) AndAlso IndexToVerse(Index)(0) = IndexToVerse(Index - 1)(0)
        End While
        If NoRef And First <> -1 And Last <> -1 Then
            If Not NoArabic Then
                Dim Ref As String = "(" + Arb.GetCatNoun("QuranReadingRecitation")(0).Text + " " + ChData.IslamData.QuranChapters(CInt(GetChapterByIndex(IndexToVerse(First)(0)).Attribute("index").Value) - 1).Name + " " + CStr(IndexToVerse(First)(0)) + "\:" + CStr(IndexToVerse(First)(1)) + If(IndexToVerse(First)(0) = IndexToVerse(Last)(0) And IndexToVerse(First)(1) = IndexToVerse(Last)(1), String.Empty, "\-" + If(IndexToVerse(First)(0) = IndexToVerse(Last)(0), String.Empty, ChData.IslamData.QuranChapters(CInt(GetChapterByIndex(IndexToVerse(Last)(0)).Attribute("index").Value) - 1).Name + " " + CStr(IndexToVerse(Last)(0)) + "\:") + CStr(IndexToVerse(Last)(1))) + ")"
                Renderer.Items.Add(New RenderArray.RenderItem(RenderArray.RenderTypes.eHeaderCenter, New RenderArray.RenderText() {New RenderArray.RenderText(RenderArray.RenderDisplayClass.eArabic, Arb.TransliterateFromBuckwalter(Ref))}))
                Renderer.Items.Add(New RenderArray.RenderItem(RenderArray.RenderTypes.eHeaderCenter, New RenderArray.RenderText() {New RenderArray.RenderText(RenderArray.RenderDisplayClass.eTransliteration, If(SchemeType <> ArabicData.TranslitScheme.None, Arb.TransliterateToScheme(Ref, SchemeType, Scheme, Arabic.FilterMetadataStops(Ref, Arb.GetMetarules(Ref, ChData.RuleMetas("Normal")), Nothing)), String.Empty))}))
            End If
            For TranslationIndex As Integer = 0 To Lines.Length - 1
                Renderer.Items.Add(New RenderArray.RenderItem(RenderArray.RenderTypes.eHeaderCenter, New RenderArray.RenderText() {New RenderArray.RenderText(RenderArray.RenderDisplayClass.eLTR, "(Quran " + GetChapterEName(GetChapterByIndex(IndexToVerse(First)(0))) + " " + CStr(IndexToVerse(First)(0)) + ":" + CStr(IndexToVerse(First)(1)) + If(IndexToVerse(First)(0) = IndexToVerse(Last)(0) And IndexToVerse(First)(1) = IndexToVerse(Last)(1), String.Empty, "-" + If(IndexToVerse(First)(0) = IndexToVerse(Last)(0), String.Empty, GetChapterEName(GetChapterByIndex(IndexToVerse(Last)(0))) + " " + CStr(IndexToVerse(Last)(0)) + ":") + CStr(IndexToVerse(Last)(1))) + ")")}))
            Next
        End If
        Return Renderer
    End Function
    Public Function GetQuranText(ByVal XMLDocMain As Xml.Linq.XDocument, ByVal StartChapter As Integer, ByVal StartAyat As Integer, ByVal EndChapter As Integer, ByVal EndAyat As Integer) As Collections.Generic.List(Of String())
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
    Public Function GetVerseCount(ByVal Chapter As Integer) As Integer
        Return CInt(GetChapterByIndex(Chapter).Attribute("ayas").Value)
    End Function
    Public Function GetWordCount(ByVal Chapter As Integer, ByVal Verse As Integer) As Integer
        Return System.Text.RegularExpressions.Regex.Matches(GetTextVerse(GetTextChapter(ChData.XMLDocMain, Chapter), Verse).Attribute("text").Value, "(?:^\s*|\s+)[^\s" + String.Join(String.Empty, Linq.Enumerable.Select(ChData.ArabicStopLetters, Function(S As String) ArabicData.MakeUniRegEx(S))) + ArabicData.ArabicStartOfRubElHizb + ArabicData.ArabicPlaceOfSajdah + "]+(?=\s*$|\s+)").Count
    End Function
    Public Function GetChunkCount(ByVal Chapter As Integer, ByVal Verse As Integer) As Integer
        Return System.Text.RegularExpressions.Regex.Matches(GetTextVerse(GetTextChapter(ChData.XMLDocMain, Chapter), Verse).Attribute("text").Value, "\s+[\s" + String.Join(String.Empty, Linq.Enumerable.Select(ChData.ArabicStopLetters, Function(S As String) ArabicData.MakeUniRegEx(S))) + ArabicData.ArabicStartOfRubElHizb + ArabicData.ArabicPlaceOfSajdah + "]+\s+").Count
    End Function
    Public Sub GetPreviousChapterVerse(ByRef Chapter As Integer, ByRef Verse As Integer)
        If Verse = 1 Then
            If Chapter <> 1 Then
                Chapter -= 1
                Verse = GetVerseCount(Chapter)
            End If
        Else
            Verse -= 1
        End If
    End Sub
    Public Function GetVerseNumber(ByVal Chapter As Integer, ByVal Verse As Integer) As Integer
        Return CInt(GetChapterByIndex(Chapter).Attribute("start").Value) + Verse
    End Function
    Public Function IsTranslationTextLTR(Index As Integer) As Boolean
        Return Not Languages.GetLanguageInfoByCode(ChData.IslamData.Translations.TranslationList(Index).FileName.Substring(0, 2), ChData.IslamData.LanguageList).IsRTL
    End Function
    Public Function GetTranslationVerse(Lines As String(), ByVal Chapter As Integer, ByVal Verse As Integer) As String
        GetTranslationVerse = Lines(GetVerseNumber(Chapter, Verse) - 1)
    End Function
    Public Function GetW4WTranslationVerse(Lines As String(), ByVal Chapter As Integer, ByVal Verse As Integer, ByVal Word As Integer) As String
        Dim Words As String() = Lines(GetVerseNumber(Chapter, Verse) - 1).Split("|"c)
        If Word > Words.Length Then
            GetW4WTranslationVerse = String.Empty
        Else
            GetW4WTranslationVerse = Words(Word - 1)
        End If
    End Function
    Public Function IsQuarterStart(ByVal Chapter As Integer, ByVal Verse As Integer) As Boolean
        For Each Node As Xml.Linq.XElement In Utility.GetChildNode("hizbs", New List(Of Xml.Linq.XElement)(ChData.XMLDocInfo.Root.Elements).ToArray()).Elements
            If Node.Name = "quarter" AndAlso
                CInt(Node.Attribute("sura").Value) = Chapter AndAlso
                CInt(Node.Attribute("aya").Value) = Verse Then
                Return True
            End If
        Next
        Return False
    End Function
    Public Function IsSajda(ByVal Chapter As Integer, ByVal Verse As Integer) As Boolean
        Dim Node As Xml.Linq.XElement
        For Each Node In Utility.GetChildNode("sajdas", New List(Of Xml.Linq.XElement)(ChData.XMLDocInfo.Root.Elements).ToArray()).Elements
            If Node.Name = "sajda" AndAlso
                CInt(Node.Attribute("sura").Value) = Chapter AndAlso
                CInt(Node.Attribute("aya").Value) = Verse Then
                Return True
            End If
        Next
        Return False
    End Function
    Public Function GetImportantName(Index As Integer) As String
        Return _PortableMethods.LoadResourceString("IslamInfo_" + ChData.IslamData.QuranSelections(Index).Description)
    End Function
    Public Function GetImportantNames() As Array()
        Dim Names() As Array = New List(Of Object())(Linq.Enumerable.Select(ChData.IslamData.QuranSelections, Function(Convert As IslamData.QuranSelection) New Object() {_PortableMethods.LoadResourceString("IslamInfo_" + Convert.Description), CInt(Array.IndexOf(ChData.IslamData.QuranSelections, Convert))})).ToArray()
        Array.Sort(Names, New Utility.CompareNameValueArray)
        Return Names
    End Function
    Public Function GetChapterCount() As Integer
        Return Utility.GetChildNodeCount("sura", Utility.GetChildNode("suras", New List(Of Xml.Linq.XElement)(ChData.XMLDocInfo.Root.Elements).ToArray()))
    End Function
    Public Function GetChapterByIndex(ByVal Index As Integer) As Xml.Linq.XElement
        Return Utility.GetChildNodeByIndex("sura", "index", Index, New List(Of Xml.Linq.XElement)(Utility.GetChildNode("suras", New List(Of Xml.Linq.XElement)(ChData.XMLDocInfo.Root.Elements).ToArray()).Elements).ToArray())
    End Function
    Public Function GetChapterName(Convert As Xml.Linq.XElement, SchemeType As ArabicData.TranslitScheme, Scheme As String) As String
        Return Convert.Attribute("index").Value + ". " + GetChapterEName(Convert) + " (" + ArabicData.RightToLeftEmbedding + Arb.TransliterateFromBuckwalter(ChData.QuranHeaders(1) + " " + ChData.IslamData.QuranChapters(CInt(Convert.Attribute("index").Value) - 1).Name) + ArabicData.PopDirectionalFormatting + ")" + If(SchemeType = ArabicData.TranslitScheme.None, String.Empty, " " + Arb.TransliterateToScheme(Arb.TransliterateFromBuckwalter(ChData.IslamData.QuranChapters(CInt(Convert.Attribute("index").Value) - 1).Name), SchemeType, Scheme, Arabic.FilterMetadataStops(Arb.TransliterateFromBuckwalter(ChData.IslamData.QuranChapters(CInt(Convert.Attribute("index").Value) - 1).Name), Arb.GetMetarules(Arb.TransliterateFromBuckwalter(ChData.IslamData.QuranChapters(CInt(Convert.Attribute("index").Value) - 1).Name), ChData.RuleMetas("Normal")), Nothing)))
    End Function
    Public Function GetChapterNames(SchemeType As ArabicData.TranslitScheme, Scheme As String) As Array()
        Dim Names() As Array = New List(Of Object())(Linq.Enumerable.Select(Utility.GetChildNodes("sura", New List(Of Xml.Linq.XElement)(Utility.GetChildNode("suras", New List(Of Xml.Linq.XElement)(ChData.XMLDocInfo.Root.Elements).ToArray()).Elements).ToArray()), Function(Convert As Xml.Linq.XElement) New Object() {GetChapterName(Convert, SchemeType, Scheme), CInt(Convert.Attribute("index").Value)})).ToArray()
        Array.Sort(Names, New Utility.CompareNameValueArray)
        Return Names
    End Function
    Public Function GetChapterEName(ByVal ChapterNode As Xml.Linq.XElement) As String
        Return _PortableMethods.LoadResourceString("IslamInfo_QuranChapter" + ChapterNode.Attribute("index").Value)
    End Function
    Public Function GetChapterNamesByRevelationOrder(SchemeType As ArabicData.TranslitScheme, Scheme As String) As Array()
        Dim Names() As Array = New List(Of Object())(Linq.Enumerable.Select(Utility.GetChildNodes("sura", New List(Of Xml.Linq.XElement)(Utility.GetChildNode("suras", New List(Of Xml.Linq.XElement)(ChData.XMLDocInfo.Root.Elements).ToArray()).Elements).ToArray()), Function(Convert As Xml.Linq.XElement) New Object() {GetChapterName(Convert, SchemeType, Scheme), CInt(Convert.Attribute("order").Value)})).ToArray()
        Array.Sort(Names, New Utility.CompareNameValueArray)
        Return Names
    End Function
    Public Function GetChapterIndexByRevelationOrder(ByVal Index As Integer) As Xml.Linq.XElement
        For Each ChapterNode As Xml.Linq.XElement In Utility.GetChildNode("suras", New List(Of Xml.Linq.XElement)(ChData.XMLDocInfo.Root.Elements).ToArray()).Elements
            If ChapterNode.Name = "sura" AndAlso CInt(ChapterNode.Attribute("order").Value) = Index Then
                Return ChapterNode
            End If
        Next
        Return Nothing
    End Function
    Public Function GetPartName(Index As Integer, SchemeType As ArabicData.TranslitScheme, Scheme As String) As String
        Dim Convert As Xml.Linq.XElement = Utility.GetChildNodeByIndex("juz", "index", Index, New List(Of Xml.Linq.XElement)(Utility.GetChildNode("juzs", New List(Of Xml.Linq.XElement)(ChData.XMLDocInfo.Root.Elements).ToArray()).Elements).ToArray())
        Return Convert.Attribute("index").Value + ". " + _PortableMethods.LoadResourceString("IslamInfo_" + ChData.IslamData.QuranParts(CInt(Convert.Attribute("index").Value) - 1).ID) + " (" + Arb.TransliterateFromBuckwalter("juzu " + ChData.IslamData.QuranParts(CInt(Convert.Attribute("index").Value) - 1).Name + " ") + ")" + If(SchemeType = ArabicData.TranslitScheme.None, String.Empty, " " + Arb.TransliterateToScheme(Arb.TransliterateFromBuckwalter("juzu " + ChData.IslamData.QuranParts(CInt(Convert.Attribute("index").Value) - 1).Name + " "), SchemeType, Scheme, Arabic.FilterMetadataStops(Arb.TransliterateFromBuckwalter("juzu " + ChData.IslamData.QuranParts(CInt(Convert.Attribute("index").Value) - 1).Name + " "), Arb.GetMetarules(Arb.TransliterateFromBuckwalter("juzu " + ChData.IslamData.QuranParts(CInt(Convert.Attribute("index").Value) - 1).Name + " "), ChData.RuleMetas("Normal")), Nothing)))
    End Function
    Public Function GetPartNames(SchemeType As ArabicData.TranslitScheme, Scheme As String) As Array()
        Dim Names() As Array = New List(Of Object())(Linq.Enumerable.Select(Utility.GetChildNodes("juz", New List(Of Xml.Linq.XElement)(Utility.GetChildNode("juzs", New List(Of Xml.Linq.XElement)(ChData.XMLDocInfo.Root.Elements).ToArray()).Elements).ToArray()), Function(Convert As Xml.Linq.XElement) New Object() {Convert.Attribute("index").Value + ". " + _PortableMethods.LoadResourceString("IslamInfo_" + ChData.IslamData.QuranParts(CInt(Convert.Attribute("index").Value) - 1).ID) + " (" + Arb.TransliterateFromBuckwalter("juzu " + ChData.IslamData.QuranParts(CInt(Convert.Attribute("index").Value) - 1).Name + " ") + ")" + If(SchemeType = ArabicData.TranslitScheme.None, String.Empty, " " + Arb.TransliterateToScheme(Arb.TransliterateFromBuckwalter("juzu " + ChData.IslamData.QuranParts(CInt(Convert.Attribute("index").Value) - 1).Name + " "), SchemeType, Scheme, Arabic.FilterMetadataStops(Arb.TransliterateFromBuckwalter("juzu " + ChData.IslamData.QuranParts(CInt(Convert.Attribute("index").Value) - 1).Name + " "), Arb.GetMetarules(Arb.TransliterateFromBuckwalter("juzu " + ChData.IslamData.QuranParts(CInt(Convert.Attribute("index").Value) - 1).Name + " "), ChData.RuleMetas("Normal")), Nothing))), CInt(Convert.Attribute("index").Value)})).ToArray()
        Array.Sort(Names, New Utility.CompareNameValueArray)
        Return Names
    End Function
    Public Function GetPartCount() As Integer
        Return Utility.GetChildNodeCount("juz", Utility.GetChildNode("juzs", New List(Of Xml.Linq.XElement)(ChData.XMLDocInfo.Root.Elements).ToArray()))
    End Function
    Public Function GetPartByIndex(ByVal Index As Integer) As Xml.Linq.XElement
        Return Utility.GetChildNodeByIndex("juz", "index", Index, New List(Of Xml.Linq.XElement)(Utility.GetChildNode("juzs", New List(Of Xml.Linq.XElement)(ChData.XMLDocInfo.Root.Elements).ToArray()).Elements).ToArray())
    End Function
    Public Function GetGroupName(Index As Integer) As String
        Dim Convert As Xml.Linq.XElement = Utility.GetChildNodeByIndex("quarter", "index", Index, New List(Of Xml.Linq.XElement)(Utility.GetChildNode("hizbs", New List(Of Xml.Linq.XElement)(ChData.XMLDocInfo.Root.Elements).ToArray()).Elements).ToArray())
        Return Convert.Attribute("index").Value
    End Function
    Public Function GetGroupNames() As Array()
        Dim Names() As Array = New List(Of Object())(Linq.Enumerable.Select(Utility.GetChildNodes("quarter", New List(Of Xml.Linq.XElement)(Utility.GetChildNode("hizbs", New List(Of Xml.Linq.XElement)(ChData.XMLDocInfo.Root.Elements).ToArray()).Elements).ToArray()), Function(Convert As Xml.Linq.XElement) New Object() {Convert.Attribute("index").Value, CInt(Convert.Attribute("index").Value)})).ToArray()
        Array.Sort(Names, New Utility.CompareNameValueArray)
        Return Names
    End Function
    Public Function GetGroupCount() As Integer
        Return Utility.GetChildNodeCount("quarter", Utility.GetChildNode("hizbs", New List(Of Xml.Linq.XElement)(ChData.XMLDocInfo.Root.Elements).ToArray()))
    End Function
    Public Function GetGroupByIndex(ByVal Index As Integer) As Xml.Linq.XElement
        Return Utility.GetChildNodeByIndex("quarter", "index", Index, New List(Of Xml.Linq.XElement)(Utility.GetChildNode("hizbs", New List(Of Xml.Linq.XElement)(ChData.XMLDocInfo.Root.Elements).ToArray()).Elements).ToArray())
    End Function
    Public Function GetStationName(Index As Integer) As String
        Dim Convert As Xml.Linq.XElement = Utility.GetChildNodeByIndex("manzil", "index", Index, New List(Of Xml.Linq.XElement)(Utility.GetChildNode("manzils", New List(Of Xml.Linq.XElement)(ChData.XMLDocInfo.Root.Elements).ToArray()).Elements).ToArray())
        Return Convert.Attribute("index").Value
    End Function
    Public Function GetStationNames() As Array()
        Dim Names() As Array = New List(Of Object())(Linq.Enumerable.Select(Utility.GetChildNodes("manzil", New List(Of Xml.Linq.XElement)(Utility.GetChildNode("manzils", New List(Of Xml.Linq.XElement)(ChData.XMLDocInfo.Root.Elements).ToArray()).Elements).ToArray()), Function(Convert As Xml.Linq.XElement) New Object() {Convert.Attribute("index").Value, CInt(Convert.Attribute("index").Value)})).ToArray()
        Array.Sort(Names, New Utility.CompareNameValueArray)
        Return Names
    End Function
    Public Function GetStationCount() As Integer
        Return Utility.GetChildNodeCount("manzil", Utility.GetChildNode("manzils", New List(Of Xml.Linq.XElement)(ChData.XMLDocInfo.Root.Elements).ToArray()))
    End Function
    Public Function GetStationByIndex(ByVal Index As Integer) As Xml.Linq.XElement
        Return Utility.GetChildNodeByIndex("manzil", "index", Index, New List(Of Xml.Linq.XElement)(Utility.GetChildNode("manzils", New List(Of Xml.Linq.XElement)(ChData.XMLDocInfo.Root.Elements).ToArray()).Elements).ToArray())
    End Function
    Public Function GetSectionName(Index As Integer) As String
        Dim Convert As Xml.Linq.XElement = Utility.GetChildNodeByIndex("ruku", "index", Index, New List(Of Xml.Linq.XElement)(Utility.GetChildNode("rukus", New List(Of Xml.Linq.XElement)(ChData.XMLDocInfo.Root.Elements).ToArray()).Elements).ToArray())
        Return Convert.Attribute("index").Value
    End Function
    Public Function GetSectionNames() As Array()
        Dim Names() As Array = New List(Of Object())(Linq.Enumerable.Select(Utility.GetChildNodes("ruku", New List(Of Xml.Linq.XElement)(Utility.GetChildNode("rukus", New List(Of Xml.Linq.XElement)(ChData.XMLDocInfo.Root.Elements).ToArray()).Elements).ToArray()), Function(Convert As Xml.Linq.XElement) New Object() {Convert.Attribute("index").Value, CInt(Convert.Attribute("index").Value)})).ToArray()
        Array.Sort(Names, New Utility.CompareNameValueArray)
        Return Names
    End Function
    Public Function GetSectionCount() As Integer
        Return Utility.GetChildNodeCount("ruku", Utility.GetChildNode("rukus", New List(Of Xml.Linq.XElement)(ChData.XMLDocInfo.Root.Elements).ToArray()))
    End Function
    Public Function GetSectionByIndex(ByVal Index As Integer) As Xml.Linq.XElement
        Return Utility.GetChildNodeByIndex("ruku", "index", Index, New List(Of Xml.Linq.XElement)(Utility.GetChildNode("rukus", New List(Of Xml.Linq.XElement)(ChData.XMLDocInfo.Root.Elements).ToArray()).Elements).ToArray())
    End Function
    Public Function GetPageName(Index As Integer) As String
        Dim Convert As Xml.Linq.XElement = Utility.GetChildNodeByIndex("page", "index", Index, New List(Of Xml.Linq.XElement)(Utility.GetChildNode("pages", New List(Of Xml.Linq.XElement)(ChData.XMLDocInfo.Root.Elements).ToArray()).Elements).ToArray())
        Return Convert.Attribute("index").Value
    End Function
    Public Function GetPageNames() As Array()
        Dim Names() As Array = New List(Of Object())(Linq.Enumerable.Select(Utility.GetChildNodes("page", New List(Of Xml.Linq.XElement)(Utility.GetChildNode("pages", New List(Of Xml.Linq.XElement)(ChData.XMLDocInfo.Root.Elements).ToArray()).Elements).ToArray()), Function(Convert As Xml.Linq.XElement) New Object() {Convert.Attribute("index").Value, CInt(Convert.Attribute("index").Value)})).ToArray()
        Array.Sort(Names, New Utility.CompareNameValueArray)
        Return Names
    End Function
    Public Function GetPageCount() As Integer
        Return Utility.GetChildNodeCount("page", Utility.GetChildNode("pages", New List(Of Xml.Linq.XElement)(ChData.XMLDocInfo.Root.Elements).ToArray()))
    End Function
    Public Function GetPageByIndex(ByVal Index As Integer) As Xml.Linq.XElement
        Return Utility.GetChildNodeByIndex("page", "index", Index, New List(Of Xml.Linq.XElement)(Utility.GetChildNode("pages", New List(Of Xml.Linq.XElement)(ChData.XMLDocInfo.Root.Elements).ToArray()).Elements).ToArray())
    End Function
    Public Function GetSajdaName(Index As Integer) As String
        Dim Convert As Xml.Linq.XElement = Utility.GetChildNodeByIndex("sajda", "index", Index, New List(Of Xml.Linq.XElement)(Utility.GetChildNode("sajdas", New List(Of Xml.Linq.XElement)(ChData.XMLDocInfo.Root.Elements).ToArray()).Elements).ToArray())
        Return Convert.Attribute("index").Value
    End Function
    Public Function GetSajdaNames() As Array()
        Dim Names() As Array = New List(Of Object())(Linq.Enumerable.Select(Utility.GetChildNodes("sajda", New List(Of Xml.Linq.XElement)(Utility.GetChildNode("sajdas", New List(Of Xml.Linq.XElement)(ChData.XMLDocInfo.Root.Elements).ToArray()).Elements).ToArray()), Function(Convert As Xml.Linq.XElement) New Object() {Convert.Attribute("index").Value, CInt(Convert.Attribute("index").Value)})).ToArray()
        Array.Sort(Names, New Utility.CompareNameValueArray)
        Return Names
    End Function
    Public Function GetSajdaCount() As Integer
        Return Utility.GetChildNodeCount("sajda", Utility.GetChildNode("sajdas", New List(Of Xml.Linq.XElement)(ChData.XMLDocInfo.Root.Elements).ToArray()))
    End Function
    Public Function GetSajdaByIndex(ByVal Index As Integer) As Xml.Linq.XElement
        Return Utility.GetChildNodeByIndex("sajda", "index", Index, New List(Of Xml.Linq.XElement)(Utility.GetChildNode("sajdas", New List(Of Xml.Linq.XElement)(ChData.XMLDocInfo.Root.Elements).ToArray()).Elements).ToArray())
    End Function
End Class
Public Class HadithReader
    Private _PortableMethods As PortableMethods
    Private Arb As Arabic
    Private ChData As CachedData
    Public Sub New(NewPortableMethods As PortableMethods, NewArb As Arabic, NewChData As CachedData)
        _PortableMethods = NewPortableMethods
        Arb = NewArb
        ChData = NewChData
    End Sub
    Public Function GetCollectionIndex(ByVal Name As String) As Integer
        Dim Count As Integer
        For Count = 0 To ChData.IslamData.Collections.Length - 1
            If Name = ChData.IslamData.Collections(Count).Name Then Return Count
        Next
        Return -1
    End Function
    Public Function GetCollectionNames() As String()
        Return New List(Of String)(Linq.Enumerable.Select(ChData.IslamData.Collections, Function(Convert As IslamData.CollectionInfo) _PortableMethods.LoadResourceString("IslamInfo_" + Convert.Name))).ToArray()
    End Function
    Public Shared Function GetChapterByIndex(ByVal BookNode As Xml.Linq.XElement, ByVal ChapterIndex As Integer) As Xml.Linq.XElement
        Return Utility.GetChildNodeByIndex("chapter", "index", ChapterIndex, New List(Of Xml.Linq.XElement)(BookNode.Elements).ToArray())
    End Function
    Public Shared Function GetSubChapterByIndex(ByVal ChapterNode As Xml.Linq.XElement, ByVal SubChapterIndex As Integer) As Xml.Linq.XElement
        Return Utility.GetChildNodeByIndex("subchapter", "index", SubChapterIndex, New List(Of Xml.Linq.XElement)(ChapterNode.Elements).ToArray())
    End Function
    Public Function GetBookEName(ByVal BookNode As Xml.Linq.XElement, CollectionIndex As Integer) As String
        If BookNode Is Nothing Then
            Return String.Empty
        Else
            Dim Str As String = _PortableMethods.LoadResourceString("IslamInfo_" + ChData.IslamData.Collections(CollectionIndex).FileName + "Book" + BookNode.Attribute("index").Value)
            If Str Is Nothing Then Return String.Empty
            Return Str
        End If
    End Function
    Public Function GetTranslationList(Collection As Integer) As Array()
        Return New List(Of String())(Linq.Enumerable.Select(ChData.IslamData.Collections(Collection).Translations, Function(Convert As IslamData.CollectionInfo.CollTranslationInfo) New String() {_PortableMethods.LoadResourceString("lang_local" + Languages.GetLanguageInfoByCode(Convert.FileName.Substring(0, 2), ChData.IslamData.LanguageList).Code) + ": " + _PortableMethods.LoadResourceString("IslamInfo_" + Convert.Name), Convert.FileName})).ToArray()
    End Function
    Public Function IsTranslationTextLTR(ByVal Index As Integer, Translation As String) As Boolean
        Dim TranslationIndex As Integer = GetTranslationIndex(Index, Translation)
        Return TranslationIndex = -1 OrElse Not Languages.GetLanguageInfoByCode(ChData.IslamData.Collections(Index).Translations(TranslationIndex).FileName.Substring(0, 2), ChData.IslamData.LanguageList).IsRTL
    End Function
    Public Function GetTranslationIndex(ByVal Index As Integer, ByVal Translation As String) As Integer
        If String.IsNullOrEmpty(Translation) Then Translation = ChData.IslamData.Collections(Index).DefaultTranslation
        Dim Count As Integer = New List(Of IslamData.CollectionInfo.CollTranslationInfo)(Linq.Enumerable.TakeWhile(ChData.IslamData.Collections(Index).Translations, Function(Test As IslamData.CollectionInfo.CollTranslationInfo) Test.FileName <> Translation)).Count
        If Count = ChData.IslamData.Collections(Index).Translations.Length Then
            Translation = ChData.IslamData.Collections(Index).DefaultTranslation
            Count = New List(Of IslamData.CollectionInfo.CollTranslationInfo)(Linq.Enumerable.TakeWhile(ChData.IslamData.Collections(Index).Translations, Function(Test As IslamData.CollectionInfo.CollTranslationInfo) Test.FileName <> Translation)).Count
        End If
        Return Count
    End Function
    Public Function GetTranslationXMLFileName(ByVal Index As Integer, ByVal Translation As String) As String
        Dim TranslationIndex As Integer = GetTranslationIndex(Index, Translation)
        Return ChData.IslamData.Collections(Index).FileName + "." + ChData.IslamData.Collections(Index).Translations(TranslationIndex).FileName + "-data"
    End Function
    Public Function GetTranslationFileName(ByVal Index As Integer, ByVal Translation As String) As String
        Dim TranslationIndex As Integer = GetTranslationIndex(Index, Translation)
        Return ChData.IslamData.Collections(Index).FileName + "." + ChData.IslamData.Collections(Index).Translations(TranslationIndex).FileName
    End Function
    Public Async Function GetHadithMappingText(Index As Integer, HadithTranslation As String) As Threading.Tasks.Task(Of Array())
        Dim XMLDocTranslate As Xml.Linq.XDocument
        If ChData.IslamData.Collections(Index).Translations.Length = 0 Then Return New Array() {}
        Dim Stream As IO.Stream = Await _PortableMethods.FileIO.LoadStream(_PortableMethods.Settings.GetFilePath(_PortableMethods.FileIO.CombinePath("metadata", GetTranslationXMLFileName(Index, HadithTranslation) + ".xml")))
        XMLDocTranslate = Xml.Linq.XDocument.Load(Stream)
        Stream.Dispose()
        Dim Output As New List(Of Object)
        Output.Add(New String() {})
        If HasVolumes(Index) Then
            Output.Add(New String() {String.Empty, String.Empty, String.Empty, String.Empty, String.Empty, String.Empty})
            Output.Add(New String() {_PortableMethods.LoadResourceString("Hadith_Volume"), _PortableMethods.LoadResourceString("Hadith_Book"), _PortableMethods.LoadResourceString("Hadith_Index"), _PortableMethods.LoadResourceString("Hadith_Chapters"), _PortableMethods.LoadResourceString("Hadith_Hadiths"), _PortableMethods.LoadResourceString("Hadith_Translation")})
        Else
            Output.Add(New String() {String.Empty, String.Empty, String.Empty, String.Empty, String.Empty})
            Output.Add(New String() {_PortableMethods.LoadResourceString("Hadith_Book"), _PortableMethods.LoadResourceString("Hadith_Index"), _PortableMethods.LoadResourceString("Hadith_Chapters"), _PortableMethods.LoadResourceString("Hadith_Hadiths"), _PortableMethods.LoadResourceString("Hadith_Translation")})
        End If
        If HasVolumes(Index) Then
            Output.AddRange(Linq.Enumerable.Select(Utility.GetChildNodes("book", New List(Of Xml.Linq.XElement)(Utility.GetChildNode("books", New List(Of Xml.Linq.XElement)(ChData.XMLDocInfos(Index).Root.Elements).ToArray()).Elements).ToArray()), Function(Convert As Xml.Linq.XElement) New Object() {CStr(GetVolumeIndex(Index, CInt(Convert.Attribute("index").Value))), GetBookEName(Convert, Index), Convert.Attribute("index").Value, CStr(GetChapterCount(Index, CInt(Convert.Attribute("index").Value))), CStr(GetHadithCount(Index, CInt(Convert.Attribute("index").Value))), GetBookHadithMapping(XMLDocTranslate, Index, CInt(Convert.Attribute("index").Value))}))
        Else
            Output.AddRange(Linq.Enumerable.Select(Utility.GetChildNodes("book", New List(Of Xml.Linq.XElement)(Utility.GetChildNode("books", New List(Of Xml.Linq.XElement)(ChData.XMLDocInfos(Index).Root.Elements).ToArray()).Elements).ToArray()), Function(Convert As Xml.Linq.XElement) New Object() {GetBookEName(Convert, Index), Convert.Attribute("index").Value, CStr(GetChapterCount(Index, CInt(Convert.Attribute("index").Value))), CStr(GetHadithCount(Index, CInt(Convert.Attribute("index").Value))), GetBookHadithMapping(XMLDocTranslate, Index, CInt(Convert.Attribute("index").Value))}))
        End If
        Return DirectCast(Output.ToArray(), Array())
    End Function
    Public Shared Function GetHadithTextBook(ByVal XMLDocMain As Xml.Linq.XDocument, ByVal BookIndex As Integer) As Xml.Linq.XElement
        Return Utility.GetChildNodeByIndex("book", "index", BookIndex, New List(Of Xml.Linq.XElement)(XMLDocMain.Root.Elements).ToArray())
    End Function
    Public Async Function GetHadithText(ByVal BookIndex As Integer, Collection As Integer) As Threading.Tasks.Task(Of Collections.Generic.List(Of Collections.Generic.List(Of Object)))
        Dim XMLDocMain As Xml.Linq.XDocument
        Dim Stream As IO.Stream = Await _PortableMethods.FileIO.LoadStream(_PortableMethods.Settings.GetFilePath(_PortableMethods.FileIO.CombinePath("metadata", ChData.IslamData.Collections(Collection).FileName + ".xml")))
        XMLDocMain = Xml.Linq.XDocument.Load(Stream)
        Stream.Dispose()
        Dim BookNode As Xml.Linq.XElement = GetHadithTextBook(XMLDocMain, BookIndex)
        Dim Hadiths As New Collections.Generic.List(Of Collections.Generic.List(Of Object))
        For Each HadithNode As Xml.Linq.XElement In BookNode.Elements
            If HadithNode.Name = "hadith" Then
                Dim NextEntry As New Collections.Generic.List(Of Object)
                NextEntry.AddRange(New Object() {CInt(HadithNode.Attribute("index").Value),
                                              CInt(Utility.ParseValue(HadithNode.Attribute("sectionindex"), "-1")),
                                              CInt(Utility.ParseValue(HadithNode.Attribute("subsectionindex"), "-1")),
                                              HadithNode.Attribute("text").Value})
                Hadiths.Add(NextEntry)
            End If
        Next
        Return Hadiths
    End Function
    Public Function GetBookCount(ByVal Index As Integer) As Integer
        Return CInt(Utility.GetChildNode("books", New List(Of Xml.Linq.XElement)(ChData.XMLDocInfos(Index).Root.Elements).ToArray()).Attribute("count").Value)
    End Function
    Public Function GetBookByIndex(ByVal Index As Integer, ByVal BookIndex As Integer) As Xml.Linq.XElement
        Return Utility.GetChildNodeByIndex("book", "index", BookIndex, New List(Of Xml.Linq.XElement)(Utility.GetChildNode("books", New List(Of Xml.Linq.XElement)(ChData.XMLDocInfos(Index).Root.Elements).ToArray()).Elements).ToArray())
    End Function
    Public Function GetBookNamesByCollection(ByVal Index As Integer) As Array()
        Dim Names() As Array = New List(Of Object())(Linq.Enumerable.Select(Utility.GetChildNodes("book", New List(Of Xml.Linq.XElement)(Utility.GetChildNode("books", New List(Of Xml.Linq.XElement)(ChData.XMLDocInfos(Index).Root.Elements).ToArray()).Elements).ToArray()), Function(Convert As Xml.Linq.XElement) New Object() {Convert.Attribute("index").Value + ". " + GetBookEName(Convert, Index), CInt(Convert.Attribute("index").Value)})).ToArray()
        Array.Sort(Names, New Utility.CompareNameValueArray)
        Return Names
    End Function
    Public Function HasVolumes(ByVal Index As Integer) As Boolean
        Return Not Utility.GetChildNode("books", New List(Of Xml.Linq.XElement)(ChData.XMLDocInfos(Index).Root.Elements).ToArray()).Attributes("volumes") Is Nothing
    End Function
    Public Function GetVolumeIndex(ByVal Index As Integer, ByVal BookIndex As Integer) As Integer
        Dim Node As Xml.Linq.XAttribute = GetBookByIndex(Index, BookIndex).Attribute("volume")
        If Node Is Nothing Then Return -1
        Return CInt(Node.Value)
    End Function
    Public Function GetChapterCount(ByVal Index As Integer, ByVal BookIndex As Integer) As Integer
        Dim Node As Xml.Linq.XAttribute = GetBookByIndex(Index, BookIndex).Attribute("chapters")
        If Node Is Nothing Then Return -1
        Return CInt(Node.Value)
    End Function
    Public Function GetChapterIndex(ByVal Index As Integer, ByVal BookIndex As Integer, ByVal HadithIndex As Integer) As Integer
        Dim BookNode As Xml.Linq.XElement = GetBookByIndex(Index, BookIndex)
        For Each ChapterNode As Xml.Linq.XElement In BookNode.Elements
            If ChapterNode.Name = "chapter" AndAlso
                (CInt(ChapterNode.Attribute("starthadith").Value) <= HadithIndex And
                CInt(ChapterNode.Attribute("starthadith").Value) + CInt(ChapterNode.Attribute("hadiths ").Value) > HadithIndex) Then Return New List(Of Xml.Linq.XElement)(ChapterNode.ElementsBeforeSelf).ToArray().Length
        Next
        Return -1
    End Function
    Public Function GetHadithCount(ByVal Index As Integer, ByVal BookIndex As Integer) As Integer
        Dim Node As Xml.Linq.XAttribute = GetBookByIndex(Index, BookIndex).Attribute("hadiths")
        If Node Is Nothing Then Return -1
        Return CInt(Node.Value)
    End Function
    Public Function GetHadithStart(ByVal Index As Integer, ByVal BookIndex As Integer) As Integer
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
                If Not Node Is Nothing AndAlso (Hadith >= CInt(Node.Value) And
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
    Public Function MapIndexes(ByVal ExpandString As String, ByVal HadithIndex As Integer) As Object()
        Dim Count As Integer
        Dim SubCount As Integer
        Dim Groupings As String() = ExpandString.Split(","c)
        Dim Indexes As New List(Of String())
        Indexes.Add(New String() {})
        Indexes.Add(New String() {String.Empty, String.Empty})
        Indexes.Add(New String() {_PortableMethods.LoadResourceString("Hadith_SourceHadithIndex"), _PortableMethods.LoadResourceString("Hadith_TranslationHadithIndex")})
        Dim HadithCount As Integer = 0
        For Count = 0 To Groupings.Length - 1
            If Groupings(Count) = String.Empty Then
                Indexes.Add(New String() {CStr(HadithIndex), _PortableMethods.LoadResourceString("Hadith_NoTranslation")})
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
    Public Function GetBookHadithMapping(ByVal XMLDocTranslate As Xml.Linq.XDocument, ByVal Index As Integer, ByVal BookIndex As Integer) As Object()
        Dim SourceStart As Integer
        Dim Volume As Integer
        Dim Mapping As New List(Of Object())
        Mapping.Add(New String() {})
        If TranslationHasVolumes(XMLDocTranslate) Then
            Mapping.Add(New String() {String.Empty, String.Empty, String.Empty, String.Empty})
            Mapping.Add(New String() {_PortableMethods.LoadResourceString("Hadith_TranslationVolume"), _PortableMethods.LoadResourceString("Hadith_TranslationBook"), _PortableMethods.LoadResourceString("Hadith_TranslationHadithCount"), _PortableMethods.LoadResourceString("Hadith_TranslationMapping")})
        Else
            Mapping.Add(New String() {String.Empty, String.Empty, String.Empty})
            Mapping.Add(New String() {_PortableMethods.LoadResourceString("Hadith_TranslationBook"), _PortableMethods.LoadResourceString("Hadith_TranslationHadithCount"), _PortableMethods.LoadResourceString("Hadith_TranslationMapping")})
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
                TranslateBookIndex = New List(Of Integer)(Linq.Enumerable.TakeWhile(Books, Function(CheckIndex As Integer) CheckIndex <> BookIndex)).Count
                If TranslateBookIndex <> Books.Length Then
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
                            Mapping.Add(New String() {CStr(Volume), BookNode.Attribute("index").Value,
                                        CStr(BookNode.Attribute("hadiths").Value), _PortableMethods.LoadResourceString("Hadith_IdenticalNumbering")})
                        Else
                            Mapping.Add(New String() {BookNode.Attribute("index").Value,
                                        CStr(BookNode.Attribute("hadiths").Value), _PortableMethods.LoadResourceString("Hadith_IdenticalNumbering")})
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
                            Mapping.Add(New Object() {CStr(Volume), BookNode.Attribute("index").Value,
                                        String.Format(_PortableMethods.LoadResourceString("Hadith_MappedOf"), CStr(RetObject(0)), CStr(CInt(BookNode.Attribute("hadiths").Value) + SharedHadith)), RetObject(1)})
                        Else
                            Mapping.Add(New Object() {BookNode.Attribute("index").Value,
                                        String.Format(_PortableMethods.LoadResourceString("Hadith_MappedOf"), CStr(RetObject(0)), CStr(CInt(BookNode.Attribute("hadiths").Value) + SharedHadith)), RetObject(1)})
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
    Public Function GetTranslationHadith(XMLDocTranslate As Xml.Linq.XDocument, Strings() As String, ByVal Index As Integer, ByVal BookIndex As Integer, ByVal HadithIndex As Integer) As String()
        Dim Node As Xml.Linq.XAttribute
        Dim Count As Integer
        Dim SubCount As Integer
        Dim SourceStart As Integer
        Dim Volume As Integer
        Dim Books() As Integer
        Dim TranslateBookIndex As Integer
        Dim TranslationIndexes As Collections.Generic.List(Of Collections.Generic.List(Of Object))
        Dim TranslationHadith As New List(Of String)
        If ChData.IslamData.Collections(Index).Translations.Length = 0 Then Return New String() {}
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
                TranslateBookIndex = New List(Of Integer)(Linq.Enumerable.TakeWhile(Books, Function(CheckIndex As Integer) CheckIndex <> BookIndex + 1)).Count
                If TranslateBookIndex <> Books.Length Then
                    Node = BookNode.Attribute("sourcestart")
                    If Node Is Nothing Then
                        SourceStart = HadithIndex
                    Else
                        SourceStart = CInt(Node.Value.Split("|"c)(TranslateBookIndex))
                    End If
                    Node = BookNode.Attribute("sourceindex")
                    If Node Is Nothing Then
                        TranslationHadith.Add(CStr(If(Volume = -1, String.Empty, _PortableMethods.LoadResourceString("Hadith_Volume") + ": " + CStr(Volume) + " ")) + _PortableMethods.LoadResourceString("Hadith_Books") + ": " + CStr(Books(TranslateBookIndex)) + " " + _PortableMethods.LoadResourceString("Hadith_Hadith") + ": " + CStr(HadithIndex))
                        TranslationHadith.AddRange(Utility.GetFileLinesByNumberPrefix(Strings, BuildTranslationIndex(XMLDocTranslate, Volume, CInt(BookNode.Attribute("index").Value), GetHadithChapter(BookNode, HadithIndex), HadithIndex, 0)))
                    Else
                        TranslationIndexes = ParseHadithTranslationIndex(Node.Value)
                        If HadithIndex >= SourceStart AndAlso HadithIndex - SourceStart < TranslationIndexes(TranslateBookIndex).Count Then
                            If TypeOf TranslationIndexes(TranslateBookIndex)(HadithIndex - SourceStart) Is Integer() Then
                                For SubCount = 0 To DirectCast(TranslationIndexes(TranslateBookIndex)(HadithIndex - SourceStart), Integer()).Length - 1
                                    Dim SharedHadithIndex As Integer = GetSharedHadithIndex(TranslationIndexes, TranslateBookIndex, HadithIndex - SourceStart, SubCount)
                                    TranslationHadith.Add(CStr(If(Volume = -1, String.Empty, _PortableMethods.LoadResourceString("Hadith_Volume") + ": " + CStr(Volume) + " ")) + _PortableMethods.LoadResourceString("Hadith_Book") + ": " + BookNode.Attribute("index").Value + " " + _PortableMethods.LoadResourceString("Hadith_Hadith") + ": " + CStr(DirectCast(TranslationIndexes(TranslateBookIndex)(HadithIndex - SourceStart), Integer())(SubCount)))
                                    TranslationHadith.AddRange(Utility.GetFileLinesByNumberPrefix(Strings, BuildTranslationIndex(XMLDocTranslate, Volume, CInt(BookNode.Attribute("index").Value), GetHadithChapter(BookNode, DirectCast(TranslationIndexes(TranslateBookIndex)(HadithIndex - SourceStart), Integer())(SubCount)), DirectCast(TranslationIndexes(TranslateBookIndex)(HadithIndex - SourceStart), Integer())(SubCount), SharedHadithIndex)))
                                Next
                            Else
                                If CInt(TranslationIndexes(TranslateBookIndex)(HadithIndex - SourceStart)) = -1 Then Return New String() {}
                                Dim SharedHadithIndex As Integer = GetSharedHadithIndex(TranslationIndexes, TranslateBookIndex, HadithIndex - SourceStart, -1)
                                TranslationHadith.Add(CStr(If(Volume = -1, String.Empty, _PortableMethods.LoadResourceString("Hadith_Volume") + ": " + CStr(Volume) + " ")) + _PortableMethods.LoadResourceString("Hadith_Book") + ": " + BookNode.Attribute("index").Value + " " + _PortableMethods.LoadResourceString("Hadith_Hadith") + ": " + CStr(TranslationIndexes(TranslateBookIndex)(HadithIndex - SourceStart)))
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