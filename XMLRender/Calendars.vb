Imports System
Imports System.Runtime.InteropServices
Imports System.Security
Imports System.Globalization

<Runtime.Serialization.DataContract> _
Public Class UmAlQuraCalendar
    Inherits Calendar
    ' Methods
    Shared Sub New()
        Dim time As New DateTime(&H81D, 11, &H10, &H17, &H3B, &H3B, &H3E7)
        UmAlQuraCalendar.maxDate = New DateTime((time.Ticks + &H270F))
    End Sub
    Friend Shared Sub CalendarCheckAddResult(ByVal ticks As Long, ByVal minValue As DateTime, ByVal maxValue As DateTime)
        If ((ticks < minValue.Ticks) OrElse (ticks > maxValue.Ticks)) Then
            Throw New ArgumentException(String.Empty) 'String.Format(CultureInfo.InvariantCulture, Environment.GetResourceString("Argument_ResultCalendarRange"), minValue, maxValue))
        End If
    End Sub
    Friend Shared Function TimeSpanTimeToTicks(ByVal hour As Integer, ByVal minute As Integer, ByVal second As Integer) As Long
        Dim num As Long = (((hour * &HE10) + (minute * 60)) + second)
        If ((num > &HD6BF94D5E5) OrElse (num < -922337203685)) Then
            Throw New ArgumentOutOfRangeException(Nothing, String.Empty) 'Environment.GetResourceString("Overflow_TimeSpanTooLong"))
        End If
        Return (num * &H989680)
    End Function
    Friend Shared Function CalendarTimeToTicks(ByVal hour As Integer, ByVal minute As Integer, ByVal second As Integer, ByVal millisecond As Integer) As Long
        If ((((hour < 0) OrElse (hour >= &H18)) OrElse ((minute < 0) OrElse (minute >= 60))) OrElse ((second < 0) OrElse (second >= 60))) Then
            Throw New ArgumentOutOfRangeException(Nothing, String.Empty) 'Environment.GetResourceString("ArgumentOutOfRange_BadHourMinuteSecond"))
        End If
        If ((millisecond < 0) OrElse (millisecond >= &H3E8)) Then
            Throw New ArgumentOutOfRangeException("millisecond", String.Empty) ' String.Format(CultureInfo.InvariantCulture, Environment.GetResourceString("ArgumentOutOfRange_Range"), 0, &H3E7))
        End If
        Return (TimeSpanTimeToTicks(hour, minute, second) + (millisecond * &H2710))
    End Function
    Public Overrides Function AddMonths(ByVal time As DateTime, ByVal months As Integer) As DateTime
        If ((months < -120000) OrElse (months > &H1D4C0)) Then
            Throw New ArgumentOutOfRangeException("months", String.Empty) 'String.Format(CultureInfo.CurrentCulture, Environment.GetResourceString("ArgumentOutOfRange_Range"), -120000, &H1D4C0))
        End If
        Dim datePart As Integer = Me.GetDatePart(time, 0)
        Dim month As Integer = Me.GetDatePart(time, 2)
        Dim day As Integer = Me.GetDatePart(time, 3)
        Dim num4 As Integer = ((month - 1) + months)
        If (num4 >= 0) Then
            month = ((num4 Mod 12) + 1)
            datePart = (datePart + (num4 / 12))
        Else
            month = (12 + ((num4 + 1) Mod 12))
            datePart = (datePart + ((num4 - 11) / 12))
        End If
        If (day > &H1D) Then
            Dim daysInMonth As Integer = Me.GetDaysInMonth(datePart, month)
            If (day > daysInMonth) Then
                day = daysInMonth
            End If
        End If
        UmAlQuraCalendar.CheckYearRange(datePart, 1)
        Dim time2 As New DateTime(((UmAlQuraCalendar.GetAbsoluteDateUmAlQura(datePart, month, day) * &HC92A69C000) + (time.Ticks Mod &HC92A69C000)))
        CalendarCheckAddResult(time2.Ticks, Me.MinSupportedDateTime, Me.MaxSupportedDateTime)
        Return time2
    End Function

    Public Overrides Function AddYears(ByVal time As DateTime, ByVal years As Integer) As DateTime
        Return Me.AddMonths(time, (years * 12))
    End Function

    Friend Shared Sub CheckEraRange(ByVal era As Integer)
        If ((era <> 0) AndAlso (era <> 1)) Then
            Throw New ArgumentOutOfRangeException("era", String.Empty) 'Environment.GetResourceString("ArgumentOutOfRange_InvalidEraValue"))
        End If
    End Sub

    Friend Shared Sub CheckTicksRange(ByVal ticks As Long)
        If ((ticks < UmAlQuraCalendar.minDate.Ticks) OrElse (ticks > UmAlQuraCalendar.maxDate.Ticks)) Then
            Throw New ArgumentOutOfRangeException("time", String.Empty) 'String.Format(CultureInfo.InvariantCulture, Environment.GetResourceString("ArgumentOutOfRange_CalendarRange"), UmAlQuraCalendar.minDate, UmAlQuraCalendar.maxDate))
        End If
    End Sub

    Friend Shared Sub CheckYearMonthRange(ByVal year As Integer, ByVal month As Integer, ByVal era As Integer)
        UmAlQuraCalendar.CheckYearRange(year, era)
        If ((month < 1) OrElse (month > 12)) Then
            Throw New ArgumentOutOfRangeException("month", String.Empty) 'Environment.GetResourceString("ArgumentOutOfRange_Month"))
        End If
    End Sub

    Friend Shared Sub CheckYearRange(ByVal year As Integer, ByVal era As Integer)
        UmAlQuraCalendar.CheckEraRange(era)
        If ((year < &H526) OrElse (year > &H5DC)) Then
            Throw New ArgumentOutOfRangeException("year", String.Empty) 'String.Format(CultureInfo.CurrentCulture, Environment.GetResourceString("ArgumentOutOfRange_Range"), &H526, &H5DC))
        End If
    End Sub

    Private Shared Sub ConvertGregorianToHijri(ByVal time As DateTime, ByRef HijriYear As Integer, ByRef HijriMonth As Integer, ByRef HijriDay As Integer)
        Dim num5 As Integer = 0
        Dim num6 As Integer = 0
        Dim num7 As Integer = 0
        Dim index As Integer = (CInt(((time.Ticks - UmAlQuraCalendar.minDate.Ticks) / &HC92A69C000)) / &H163)
        Do While (time.CompareTo(UmAlQuraCalendar.HijriYearInfo(++index).GregorianDate) > 0)
        Loop
        If (time.CompareTo(UmAlQuraCalendar.HijriYearInfo(index).GregorianDate) <> 0) Then
            index -= 1
        End If
        Dim span As TimeSpan = time.Subtract(UmAlQuraCalendar.HijriYearInfo(index).GregorianDate)
        num5 = (index + &H526)
        num6 = 1
        num7 = 1
        Dim totalDays As Double = span.TotalDays
        Dim hijriMonthsLengthFlags As Integer = UmAlQuraCalendar.HijriYearInfo(index).HijriMonthsLengthFlags
        Dim num3 As Integer = (&H1D + (hijriMonthsLengthFlags And 1))
        Do While (totalDays >= num3)
            totalDays = (totalDays - num3)
            hijriMonthsLengthFlags = (hijriMonthsLengthFlags >> 1)
            num3 = (&H1D + (hijriMonthsLengthFlags And 1))
            num6 += 1
        Loop
        num7 = (num7 + CInt(totalDays))
        HijriDay = num7
        HijriMonth = num6
        HijriYear = num5
    End Sub

    Private Shared Sub ConvertHijriToGregorian(ByVal HijriYear As Integer, ByVal HijriMonth As Integer, ByVal HijriDay As Integer, ByRef yg As Integer, ByRef mg As Integer, ByRef dg As Integer)
        Dim num3 As Integer = (HijriDay - 1)
        Dim index As Integer = (HijriYear - &H526)
        Dim gregorianDate As DateTime = UmAlQuraCalendar.HijriYearInfo(index).GregorianDate
        Dim hijriMonthsLengthFlags As Integer = UmAlQuraCalendar.HijriYearInfo(index).HijriMonthsLengthFlags
        Dim i As Integer
        For i = 1 To HijriMonth - 1
            num3 = (num3 + (&H1D + (hijriMonthsLengthFlags And 1)))
            hijriMonthsLengthFlags = (hijriMonthsLengthFlags >> 1)
        Next i
        gregorianDate = gregorianDate.AddDays(CDbl(num3))
        yg = gregorianDate.Year
        mg = gregorianDate.Month
        dg = gregorianDate.Day
    End Sub
    Friend Shared Function GregorianCalendarGetAbsoluteDate(ByVal year As Integer, ByVal month As Integer, ByVal day As Integer) As Long
        Dim GregorianCalendarDaysToMonth365 As Integer() = {0, &H1F, &H3B, 90, 120, &H97, &HB5, &HD4, &HF3, &H111, &H130, &H14E, &H16D}
        Dim GregorianCalendarDaysToMonth366 As Integer() = {0, &H1F, 60, &H5B, &H79, &H98, &HB6, &HD5, &HF4, &H112, &H131, &H14F, &H16E}
        If (((year >= 1) AndAlso (year <= &H270F)) AndAlso ((month >= 1) AndAlso (month <= 12))) Then
            Dim numArray As Integer() = If((((year Mod 4) = 0) AndAlso (((year Mod 100) <> 0) OrElse ((year Mod 400) = 0))), GregorianCalendarDaysToMonth366, GregorianCalendarDaysToMonth365)
            If ((day >= 1) AndAlso (day <= (numArray(month) - numArray((month - 1))))) Then
                Dim num As Integer = (year - 1)
                Return CLng((((((((num * &H16D) + (num / 4)) - (num / 100)) + (num / 400)) + numArray((month - 1))) + day) - 1))
            End If
        End If
        Throw New ArgumentOutOfRangeException(Nothing, String.Empty) 'Environment.GetResourceString("ArgumentOutOfRange_BadYearMonthDay"))
    End Function

    Private Shared Function GetAbsoluteDateUmAlQura(ByVal year As Integer, ByVal month As Integer, ByVal day As Integer) As Long
        Dim yg As Integer = 0
        Dim mg As Integer = 0
        Dim dg As Integer = 0
        UmAlQuraCalendar.ConvertHijriToGregorian(year, month, day, yg, mg, dg)
        Return GregorianCalendarGetAbsoluteDate(yg, mg, dg)
    End Function

    Friend Overridable Function GetDatePart(ByVal time As DateTime, ByVal part As Integer) As Integer
        Dim hijriYear As Integer = 0
        Dim hijriMonth As Integer = 0
        Dim hijriDay As Integer = 0
        UmAlQuraCalendar.CheckTicksRange(time.Ticks)
        UmAlQuraCalendar.ConvertGregorianToHijri(time, hijriYear, hijriMonth, hijriDay)
        If (part = 0) Then
            Return hijriYear
        End If
        If (part = 2) Then
            Return hijriMonth
        End If
        If (part = 3) Then
            Return hijriDay
        End If
        If (part <> 1) Then
            Throw New InvalidOperationException(String.Empty) 'Environment.GetResourceString("InvalidOperation_DateTimeParsing"))
        End If
        Return CInt(((UmAlQuraCalendar.GetAbsoluteDateUmAlQura(hijriYear, hijriMonth, hijriDay) - UmAlQuraCalendar.GetAbsoluteDateUmAlQura(hijriYear, 1, 1)) + 1))
    End Function

    Public Overrides Function GetDayOfMonth(ByVal time As DateTime) As Integer
        Return Me.GetDatePart(time, 3)
    End Function

    Public Overrides Function GetDayOfWeek(ByVal time As DateTime) As DayOfWeek
        Return DirectCast((CInt(((time.Ticks / &HC92A69C000) + 1)) Mod 7), DayOfWeek)
    End Function

    Public Overrides Function GetDayOfYear(ByVal time As DateTime) As Integer
        Return Me.GetDatePart(time, 1)
    End Function

    Public Overrides Function GetDaysInMonth(ByVal year As Integer, ByVal month As Integer, ByVal era As Integer) As Integer
        UmAlQuraCalendar.CheckYearMonthRange(year, month, era)
        If ((UmAlQuraCalendar.HijriYearInfo((year - &H526)).HijriMonthsLengthFlags And (CInt(1) << (month - 1))) = 0) Then
            Return &H1D
        End If
        Return 30
    End Function

    Public Overrides Function GetDaysInYear(ByVal year As Integer, ByVal era As Integer) As Integer
        UmAlQuraCalendar.CheckYearRange(year, era)
        Return UmAlQuraCalendar.RealGetDaysInYear(year)
    End Function

    Public Overrides Function GetEra(ByVal time As DateTime) As Integer
        UmAlQuraCalendar.CheckTicksRange(time.Ticks)
        Return 1
    End Function

    Public Overrides Function GetLeapMonth(ByVal year As Integer, ByVal era As Integer) As Integer
        UmAlQuraCalendar.CheckYearRange(year, era)
        Return 0
    End Function

    Public Overrides Function GetMonth(ByVal time As DateTime) As Integer
        Return Me.GetDatePart(time, 2)
    End Function

    Public Overrides Function GetMonthsInYear(ByVal year As Integer, ByVal era As Integer) As Integer
        UmAlQuraCalendar.CheckYearRange(year, era)
        Return 12
    End Function

    Public Overrides Function GetYear(ByVal time As DateTime) As Integer
        Return Me.GetDatePart(time, 0)
    End Function

    Private Shared Function InitDateMapping() As DateMapping()
        Dim numArray As Short() = New Short() {&H2EA, &H76C, 4, 30, &H6E9, &H76D, 4, &H13, &HED2, &H76E, 4, 9, &HEA4, &H76F, 3, 30, &HD4A, &H770, 3, &H12, &HA96, &H771, 3, 7, &H536, &H772, 2, &H18, &HAB5, &H773, 2, 13, &HDAA, &H774, 2, 3, &HBA4, &H775, 1, &H17, &HB49, &H776, 1, 12, &HA93, &H777, 1, 1, &H52B, &H777, 12, &H15, &HA57, &H778, 12, 9, &H4B6, &H779, 11, &H1D, &HAB5, &H77A, 11, &H12, &H5AA, &H77B, 11, 8, &HD55, &H77C, 10, &H1B, &HD2A, &H77D, 10, &H11, &HA56, &H77E, 10, 6, &H4AE, &H77F, 9, &H19, &H95D, &H780, 9, 13, &H2EC, &H781, 9, 3, &H6D5, &H782, 8, &H17, &H6AA, &H783, 8, 13, &H555, &H784, 8, 1, &H4AB, &H785, 7, &H15, &H95B, &H786, 7, 10, &H2BA, &H787, 6, 30, &H575, &H788, 6, &H12, &HBB2, &H789, 6, 8, &H764, &H78A, 5, &H1D, &H749, &H78B, 5, &H12, &H655, &H78C, 5, 6, &H2AB, &H78D, 4, &H19, &H55B, &H78E, 4, 14, &HADA, &H78F, 4, 4, &H6D4, &H790, 3, &H18, &HEC9, &H791, 3, 13, &HD92, &H792, 3, 3, &HD25, &H793, 2, 20, &HA4D, &H794, 2, 9, &H2AD, &H795, 1, &H1C, &H56D, &H796, 1, &H11, &HB6A, &H797, 1, 7, &HB52, &H797, 12, &H1C, &HAA5, &H798, 12, &H10, &HA4B, &H799, 12, 5, &H497, &H79A, 11, &H18, &H937, &H79B, 11, 13, &H2B6, &H79C, 11, 2, &H575, &H79D, 10, &H16, &HD6A, &H79E, 10, 12, &HD52, &H79F, 10, 2, &HA96, &H7A0, 9, 20, &H92D, &H7A1, 9, 9, &H25D, &H7A2, 8, &H1D, &H4DD, &H7A3, 8, &H12, &HADA, &H7A4, 8, 7, &H5D4, &H7A5, 7, &H1C, &HDA9, &H7A6, 7, &H11, &HD52, &H7A7, 7, 7, &HAAA, &H7A8, 6, &H19, &H4D6, &H7A9, 6, 14, &H9B6, &H7AA, 6, 3, &H374, &H7AB, 5, &H18, &H769, &H7AC, 5, 12, &H752, &H7AD, 5, 2, &H6A5, &H7AE, 4, &H15, &H54B, &H7AF, 4, 10, &HAAB, &H7B0, 3, &H1D, &H55A, &H7B1, 3, &H13, &HAD5, &H7B2, 3, 8, &HDD2, &H7B3, 2, &H1A, &HDA4, &H7B4, 2, &H10, &HD49, &H7B5, 2, 4, &HA95, &H7B6, 1, &H18, &H52D, &H7B7, 1, 13, &HA5D, &H7B8, 1, 2, &H55A, &H7B8, 12, &H16, &HAD5, &H7B9, 12, 11, &H6AA, &H7BA, 12, 1, &H695, &H7BB, 11, 20, &H52B, &H7BC, 11, 8, &HA57, &H7BD, 10, &H1C, &H4AE, &H7BE, 10, &H12, &H976, &H7BF, 10, 7, &H56C, &H7C0, 9, &H1A, &HB55, &H7C1, 9, 15, &HAAA, &H7C2, 9, 5, &HA55, &H7C3, 8, &H19, &H4AD, &H7C4, 8, 13, &H95D, &H7C5, 8, 2, 730, &H7C6, 7, &H17, &H5D9, &H7C7, 7, 12, &HDB2, &H7C8, 7, 1, &HBA4, &H7C9, 6, &H15, &HB4A, &H7CA, 6, 10, &HA55, &H7CB, 5, 30, &H2B5, &H7CC, 5, &H12, &H575, &H7CD, 5, 7, &HB6A, &H7CE, 4, &H1B, &HBD2, &H7CF, 4, &H11, &HBC4, &H7D0, 4, 6, &HB89, &H7D1, 3, &H1A, &HA95, &H7D2, 3, 15, &H52D, &H7D3, 3, 4, &H5AD, &H7D4, 2, &H15, &HB6A, &H7D5, 2, 10, &H6D4, &H7D6, 1, &H1F, &HDC9, &H7D7, 1, 20, &HD92, &H7D8, 1, 10, &HAA6, &H7D8, 12, &H1D, &H956, &H7D9, 12, &H12, &H2AE, &H7DA, 12, 7, &H56D, &H7DB, 11, &H1A, &H36A, &H7DC, 11, 15, &HB55, &H7DD, 11, 4, &HAAA, &H7DE, 10, &H19, &H94D, &H7DF, 10, 14, &H49D, &H7E0, 10, 2, &H95D, &H7E1, 9, &H15, &H2BA, &H7E2, 9, 11, &H5B5, &H7E3, 8, &H1F, &H5AA, &H7E4, 8, 20, &HD55, &H7E5, 8, 9, &HA9A, &H7E6, 7, 30, &H92E, &H7E7, 7, &H13, &H26E, &H7E8, 7, 7, &H55D, &H7E9, 6, &H1A, &HADA, &H7EA, 6, &H10, &H6D4, &H7EB, 6, 6, &H6A5, &H7EC, 5, &H19, &H54B, &H7ED, 5, 14, &HA97, &H7EE, 5, 3, &H54E, &H7EF, 4, &H17, &HAAE, &H7F0, 4, 11, &H5AC, &H7F1, 4, 1, &HBA9, &H7F2, 3, &H15, &HD92, &H7F3, 3, 11, &HB25, &H7F4, 2, &H1C, &H64B, &H7F5, 2, &H10, &HCAB, &H7F6, 2, 5, &H55A, &H7F7, 1, &H1A, &HB55, &H7F8, 1, 15, &H6D2, &H7F9, 1, 4, &HEA5, &H7F9, 12, &H18, &HE4A, &H7FA, 12, 14, &HA95, &H7FB, 12, 3, &H52D, &H7FC, 11, &H15, &HAAD, &H7FD, 11, 10, &H36C, &H7FE, 10, &H1F, &H759, &H7FF, 10, 20, &H6D2, &H800, 10, 9, &H695, &H801, 9, &H1C, &H52D, &H802, 9, &H11, &HA5B, &H803, 9, 6, &H4BA, &H804, 8, &H1A, &H9BA, &H805, 8, 15, &H3B4, &H806, 8, 5, &HB69, &H807, 7, &H19, &HB52, &H808, 7, 14, &HAA6, &H809, 7, 3, &H4B6, &H80A, 6, &H16, &H96D, &H80B, 6, 11, &H2EC, &H80C, 5, &H1F, &H6D9, &H80D, 5, 20, &HEB2, &H80E, 5, 10, &HD54, &H80F, 4, 30, &HD2A, &H810, 4, &H12, &HA56, &H811, 4, 7, &H4AE, &H812, 3, &H1B, &H96D, &H813, 3, &H10, &HD6A, &H814, 3, 5, &HB54, &H815, 2, &H17, &HB29, &H816, 2, 12, &HA93, &H817, 2, 1, &H52B, &H818, 1, &H15, &HA57, &H819, 1, 9, &H536, &H819, 12, 30, &HAB5, &H81A, 12, &H13, &H6AA, &H81B, 12, 9, &HE93, &H81C, 11, &H1B, 0, &H81D, 11, &H11}
        Dim mappingArray As DateMapping() = New DateMapping((numArray.Length / 4) - 1) {}
        Dim i As Integer
        For i = 0 To mappingArray.Length - 1
            mappingArray(i) = New DateMapping(numArray((i * 4)), numArray(((i * 4) + 1)), numArray(((i * 4) + 2)), numArray(((i * 4) + 3)))
        Next i
        Return mappingArray
    End Function

    Public Overrides Function IsLeapDay(ByVal year As Integer, ByVal month As Integer, ByVal day As Integer, ByVal era As Integer) As Boolean
        If ((day >= 1) AndAlso (day <= &H1D)) Then
            UmAlQuraCalendar.CheckYearMonthRange(year, month, era)
            Return False
        End If
        Dim num As Integer = Me.GetDaysInMonth(year, month, era)
        If ((day < 1) OrElse (day > num)) Then
            Throw New ArgumentOutOfRangeException("day", String.Empty) 'String.Format(CultureInfo.CurrentCulture, Environment.GetResourceString("ArgumentOutOfRange_Day"), num, month))
        End If
        Return False
    End Function

    Public Overrides Function IsLeapMonth(ByVal year As Integer, ByVal month As Integer, ByVal era As Integer) As Boolean
        UmAlQuraCalendar.CheckYearMonthRange(year, month, era)
        Return False
    End Function

    Public Overrides Function IsLeapYear(ByVal year As Integer, ByVal era As Integer) As Boolean
        UmAlQuraCalendar.CheckYearRange(year, era)
        Return (UmAlQuraCalendar.RealGetDaysInYear(year) = &H163)
    End Function

    Friend Shared Function RealGetDaysInYear(ByVal year As Integer) As Integer
        Dim num As Integer = 0
        Dim hijriMonthsLengthFlags As Integer = UmAlQuraCalendar.HijriYearInfo((year - &H526)).HijriMonthsLengthFlags
        Dim i As Integer = 1
        Do While (i <= 12)
            num = (num + (&H1D + (hijriMonthsLengthFlags And 1)))
            hijriMonthsLengthFlags = (hijriMonthsLengthFlags >> 1)
            i += 1
        Loop
        Return num
    End Function

    Public Overrides Function ToDateTime(ByVal year As Integer, ByVal month As Integer, ByVal day As Integer, ByVal hour As Integer, ByVal minute As Integer, ByVal second As Integer, ByVal millisecond As Integer, ByVal era As Integer) As DateTime
        If ((day >= 1) AndAlso (day <= &H1D)) Then
            UmAlQuraCalendar.CheckYearMonthRange(year, month, era)
        Else
            Dim num As Integer = Me.GetDaysInMonth(year, month, era)
            If ((day < 1) OrElse (day > num)) Then
                Throw New ArgumentOutOfRangeException("day", String.Empty) 'String.Format(CultureInfo.CurrentCulture, Environment.GetResourceString("ArgumentOutOfRange_Day"), num, month))
            End If
        End If
        Dim num2 As Long = UmAlQuraCalendar.GetAbsoluteDateUmAlQura(year, month, day)
        If (num2 < 0) Then
            Throw New ArgumentOutOfRangeException(Nothing, String.Empty) 'Environment.GetResourceString("ArgumentOutOfRange_BadYearMonthDay"))
        End If
        Return New DateTime(((num2 * &HC92A69C000) + CalendarTimeToTicks(hour, minute, second, millisecond)))
    End Function

    Public Overrides Function ToFourDigitYear(ByVal year As Integer) As Integer
        If (year < 0) Then
            Throw New ArgumentOutOfRangeException("year", String.Empty) 'Environment.GetResourceString("ArgumentOutOfRange_NeedNonNegNum"))
        End If
        If (year < 100) Then
            Return MyBase.ToFourDigitYear(year)
        End If
        If ((year < &H526) OrElse (year > &H5DC)) Then
            Throw New ArgumentOutOfRangeException("year", String.Empty) 'String.Format(CultureInfo.CurrentCulture, Environment.GetResourceString("ArgumentOutOfRange_Range"), &H526, &H5DC))
        End If
        Return year
    End Function


    ' Properties
    'Public Overrides ReadOnly Property AlgorithmType As CalendarAlgorithmType
    '    Get
    '        Return CalendarAlgorithmType.LunarCalendar
    '    End Get
    'End Property

    Friend ReadOnly Property BaseCalendarID As Integer
        Get
            Return 6
        End Get
    End Property

    Protected ReadOnly Property DaysInYearBeforeMinSupportedYear As Integer
        Get
            Return &H163
        End Get
    End Property

    Public Overrides ReadOnly Property Eras As Integer()
        Get
            Return New Integer() {1}
        End Get
    End Property

    Friend ReadOnly Property ID As Integer
        Get
            Return &H17
        End Get
    End Property

    Public Overrides ReadOnly Property MaxSupportedDateTime As DateTime
        Get
            Return UmAlQuraCalendar.maxDate
        End Get
    End Property

    Public Overrides ReadOnly Property MinSupportedDateTime As DateTime
        Get
            Return UmAlQuraCalendar.minDate
        End Get
    End Property

    Public Overrides Property TwoDigitYearMax As Integer
        Get
            If (MyBase.TwoDigitYearMax = -1) Then
                MyBase.TwoDigitYearMax = &H5AB 'Calendar.GetSystemTwoDigitYearSetting(Me.ID, &H5AB)
            End If
            Return MyBase.TwoDigitYearMax
        End Get
        Set(ByVal value As Integer)
            If ((value <> &H63) AndAlso ((value < &H526) OrElse (value > &H5DC))) Then
                Throw New ArgumentOutOfRangeException("value", String.Empty) 'String.Format(CultureInfo.CurrentCulture, Environment.GetResourceString("ArgumentOutOfRange_Range"), &H526, &H5DC))
            End If
            'MyBase.VerifyWritable()
            MyBase.TwoDigitYearMax = value
        End Set
    End Property


    ' Fields
    Friend Const DateCycle As Integer = 30
    Friend Const DatePartDay As Integer = 3
    Friend Const DatePartDayOfYear As Integer = 1
    Friend Const DatePartMonth As Integer = 2
    Friend Const DatePartYear As Integer = 0
    Private Const DEFAULT_TWO_DIGIT_YEAR_MAX As Integer = &H5AB
    Private Shared ReadOnly HijriYearInfo As DateMapping() = UmAlQuraCalendar.InitDateMapping
    Friend Const MaxCalendarYear As Integer = &H5DC
    Friend Shared maxDate As DateTime
    Friend Const MinCalendarYear As Integer = &H526
    Friend Shared minDate As DateTime = New DateTime(&H76C, 4, 30)
    Public Const UmAlQuraEra As Integer = 1

    ' Nested Types
    <StructLayout(LayoutKind.Sequential)> _
    Friend Structure DateMapping
        Friend HijriMonthsLengthFlags As Integer
        Friend GregorianDate As DateTime
        Friend Sub New(ByVal MonthsLengthFlags As Integer, ByVal GYear As Integer, ByVal GMonth As Integer, ByVal GDay As Integer)
            Me.HijriMonthsLengthFlags = MonthsLengthFlags
            Me.GregorianDate = New DateTime(GYear, GMonth, GDay)
        End Sub
    End Structure
End Class
<Runtime.Serialization.DataContract, ComVisible(True)> _
Public Class HijriCalendar
    Inherits Calendar
    ' Methods
    Friend Shared Sub CalendarCheckAddResult(ByVal ticks As Long, ByVal minValue As DateTime, ByVal maxValue As DateTime)
        If ((ticks < minValue.Ticks) OrElse (ticks > maxValue.Ticks)) Then
            Throw New ArgumentException(String.Empty) 'String.Format(CultureInfo.InvariantCulture, Environment.GetResourceString("Argument_ResultCalendarRange"), minValue, maxValue))
        End If
    End Sub
    Friend Shared Function TimeSpanTimeToTicks(ByVal hour As Integer, ByVal minute As Integer, ByVal second As Integer) As Long
        Dim num As Long = (((hour * &HE10) + (minute * 60)) + second)
        If ((num > &HD6BF94D5E5) OrElse (num < -922337203685)) Then
            Throw New ArgumentOutOfRangeException(Nothing, String.Empty) 'Environment.GetResourceString("Overflow_TimeSpanTooLong"))
        End If
        Return (num * &H989680)
    End Function
    Friend Shared Function CalendarTimeToTicks(ByVal hour As Integer, ByVal minute As Integer, ByVal second As Integer, ByVal millisecond As Integer) As Long
        If ((((hour < 0) OrElse (hour >= &H18)) OrElse ((minute < 0) OrElse (minute >= 60))) OrElse ((second < 0) OrElse (second >= 60))) Then
            Throw New ArgumentOutOfRangeException(Nothing, String.Empty) 'Environment.GetResourceString("ArgumentOutOfRange_BadHourMinuteSecond"))
        End If
        If ((millisecond < 0) OrElse (millisecond >= &H3E8)) Then
            Throw New ArgumentOutOfRangeException("millisecond", String.Empty) ' String.Format(CultureInfo.InvariantCulture, Environment.GetResourceString("ArgumentOutOfRange_Range"), 0, &H3E7))
        End If
        Return (TimeSpanTimeToTicks(hour, minute, second) + (millisecond * &H2710))
    End Function
    Public Overrides Function AddMonths(ByVal time As DateTime, ByVal months As Integer) As DateTime
        If ((months < -120000) OrElse (months > &H1D4C0)) Then
            Throw New ArgumentOutOfRangeException("months", String.Empty) 'String.Format(CultureInfo.CurrentCulture, Environment.GetResourceString("ArgumentOutOfRange_Range"), -120000, &H1D4C0))
        End If
        Dim datePart As Integer = Me.GetDatePart(time.Ticks, 0)
        Dim month As Integer = Me.GetDatePart(time.Ticks, 2)
        Dim d As Integer = Me.GetDatePart(time.Ticks, 3)
        Dim num4 As Integer = ((month - 1) + months)
        If (num4 >= 0) Then
            month = ((num4 Mod 12) + 1)
            datePart = (datePart + (num4 / 12))
        Else
            month = (12 + ((num4 + 1) Mod 12))
            datePart = (datePart + ((num4 - 11) / 12))
        End If
        Dim daysInMonth As Integer = Me.GetDaysInMonth(datePart, month)
        If (d > daysInMonth) Then
            d = daysInMonth
        End If
        Dim ticks As Long = ((Me.GetAbsoluteDateHijri(datePart, month, d) * &HC92A69C000) + (time.Ticks Mod &HC92A69C000))
        CalendarCheckAddResult(ticks, Me.MinSupportedDateTime, Me.MaxSupportedDateTime)
        Return New DateTime(ticks)
    End Function

    Public Overrides Function AddYears(ByVal time As DateTime, ByVal years As Integer) As DateTime
        Return Me.AddMonths(time, (years * 12))
    End Function

    Friend Shared Sub CheckEraRange(ByVal era As Integer)
        If ((era <> 0) AndAlso (era <> HijriCalendar.HijriEra)) Then
            Throw New ArgumentOutOfRangeException("era", String.Empty) 'Environment.GetResourceString("ArgumentOutOfRange_InvalidEraValue"))
        End If
    End Sub

    Friend Shared Sub CheckTicksRange(ByVal ticks As Long)
        If ((ticks < HijriCalendar.calendarMinValue.Ticks) OrElse (ticks > HijriCalendar.calendarMaxValue.Ticks)) Then
            Throw New ArgumentOutOfRangeException("time", String.Empty) 'String.Format(CultureInfo.InvariantCulture, Environment.GetResourceString("ArgumentOutOfRange_CalendarRange"), HijriCalendar.calendarMinValue, HijriCalendar.calendarMaxValue))
        End If
    End Sub

    Friend Shared Sub CheckYearMonthRange(ByVal year As Integer, ByVal month As Integer, ByVal era As Integer)
        HijriCalendar.CheckYearRange(year, era)
        If ((year = &H25C2) AndAlso (month > 4)) Then
            Throw New ArgumentOutOfRangeException("month", String.Empty) 'String.Format(CultureInfo.CurrentCulture, Environment.GetResourceString("ArgumentOutOfRange_Range"), 1, 4))
        End If
        If ((month < 1) OrElse (month > 12)) Then
            Throw New ArgumentOutOfRangeException("month", String.Empty) 'Environment.GetResourceString("ArgumentOutOfRange_Month"))
        End If
    End Sub

    Friend Shared Sub CheckYearRange(ByVal year As Integer, ByVal era As Integer)
        HijriCalendar.CheckEraRange(era)
        If ((year < 1) OrElse (year > &H25C2)) Then
            Throw New ArgumentOutOfRangeException("year", String.Empty) 'String.Format(CultureInfo.CurrentCulture, Environment.GetResourceString("ArgumentOutOfRange_Range"), 1, &H25C2))
        End If
    End Sub

    Private Function DaysUpToHijriYear(ByVal HijriYear As Integer) As Long
        Dim num2 As Integer = (((HijriYear - 1) / 30) * 30)
        Dim year As Integer = ((HijriYear - num2) - 1)
        Dim num As Long = (((num2 * &H2987) / 30) + &H376C5)
        Do While (year > 0)
            num = (num + (&H162 + If(Me.IsLeapYear(year, 0), 1, 0)))
            year -= 1
        Loop
        Return num
    End Function

    Private Function GetAbsoluteDateHijri(ByVal y As Integer, ByVal m As Integer, ByVal d As Integer) As Long
        Return ((((Me.DaysUpToHijriYear(y) + HijriCalendar.HijriMonthDays((m - 1))) + d) - 1) - Me.HijriAdjustment)
    End Function

    <SecurityCritical> _
    Private Shared Function GetAdvanceHijriDate() As Integer
        Dim num As Integer = 0
        'Dim key As RegistryKey = Nothing
        'Try
        '    'key = Registry.CurrentUser.InternalOpenSubKey("Control Panel\International", False)
        'Catch exception1 As ObjectDisposedException
        '    Return 0
        'Catch exception2 As ArgumentException
        '    Return 0
        'End Try
        'If (Not key Is Nothing) Then
        '    Try
        '        Dim obj2 As Object = key.InternalGetValue("AddHijriDate", Nothing, False, False)
        '        If (obj2 Is Nothing) Then
        '            Return 0
        '        End If
        '        Dim strA As String = obj2.ToString
        '        If (String.Compare(strA, 0, "AddHijriDate", 0, "AddHijriDate".Length, StringComparison.OrdinalIgnoreCase) <> 0) Then
        '            Return num
        '        End If
        '        If (strA.Length = "AddHijriDate".Length) Then
        '            Return -1
        '        End If
        '        strA = strA.Substring("AddHijriDate".Length)
        '        Try
        '            Dim num3 As Integer = Integer.Parse(strA.ToString, CultureInfo.InvariantCulture)
        '            If ((num3 >= -2) AndAlso (num3 <= 2)) Then
        '                num = num3
        '            End If
        '            Return num
        '        Catch exception3 As ArgumentException
        '            Return num
        '        Catch exception4 As FormatException
        '            Return num
        '        Catch exception5 As OverflowException
        '            Return num
        '        End Try
        '    Finally
        '        key.Close()
        '    End Try
        'End If
        Return num
    End Function

    Friend Overridable Function GetDatePart(ByVal ticks As Long, ByVal part As Integer) As Integer
        HijriCalendar.CheckTicksRange(ticks)
        Dim num4 As Long = ((ticks / &HC92A69C000) + 1)
        num4 = (num4 + Me.HijriAdjustment)
        Dim hijriYear As Integer = (CInt((((num4 - &H376C5) * 30) / &H2987)) + 1)
        Dim num5 As Long = Me.DaysUpToHijriYear(hijriYear)
        Dim daysInYear As Long = Me.GetDaysInYear(hijriYear, 0)
        If (num4 < num5) Then
            num5 = (num5 - daysInYear)
            hijriYear -= 1
        ElseIf (num4 = num5) Then
            hijriYear -= 1
            num5 = (num5 - Me.GetDaysInYear(hijriYear, 0))
        ElseIf (num4 > (num5 + daysInYear)) Then
            num5 = (num5 + daysInYear)
            hijriYear += 1
        End If
        If (part = 0) Then
            Return hijriYear
        End If
        Dim num2 As Integer = 1
        num4 = (num4 - num5)
        If (part = 1) Then
            Return CInt(num4)
        End If
        Do While True
            If ((num2 > 12) OrElse (num4 <= HijriCalendar.HijriMonthDays((num2 - 1)))) Then
                num2 -= 1
                If (part = 2) Then
                    Return num2
                End If
                Dim num3 As Integer = (CInt(num4) - HijriCalendar.HijriMonthDays((num2 - 1)))
                If (part <> 3) Then
                    Throw New InvalidOperationException(String.Empty) 'Environment.GetResourceString("InvalidOperation_DateTimeParsing"))
                End If
                Return num3
            End If
            num2 += 1
        Loop
        Return 0 'unreachable
    End Function

    Public Overrides Function GetDayOfMonth(ByVal time As DateTime) As Integer
        Return Me.GetDatePart(time.Ticks, 3)
    End Function

    Public Overrides Function GetDayOfWeek(ByVal time As DateTime) As DayOfWeek
        Return DirectCast((CInt(((time.Ticks / &HC92A69C000) + 1)) Mod 7), DayOfWeek)
    End Function

    Public Overrides Function GetDayOfYear(ByVal time As DateTime) As Integer
        Return Me.GetDatePart(time.Ticks, 1)
    End Function

    Public Overrides Function GetDaysInMonth(ByVal year As Integer, ByVal month As Integer, ByVal era As Integer) As Integer
        HijriCalendar.CheckYearMonthRange(year, month, era)
        If (month = 12) Then
            If Not Me.IsLeapYear(year, 0) Then
                Return &H1D
            End If
            Return 30
        End If
        If ((month Mod 2) <> 1) Then
            Return &H1D
        End If
        Return 30
    End Function

    Public Overrides Function GetDaysInYear(ByVal year As Integer, ByVal era As Integer) As Integer
        HijriCalendar.CheckYearRange(year, era)
        If Not Me.IsLeapYear(year, 0) Then
            Return &H162
        End If
        Return &H163
    End Function

    Public Overrides Function GetEra(ByVal time As DateTime) As Integer
        HijriCalendar.CheckTicksRange(time.Ticks)
        Return HijriCalendar.HijriEra
    End Function

    <ComVisible(False)> _
    Public Overrides Function GetLeapMonth(ByVal year As Integer, ByVal era As Integer) As Integer
        HijriCalendar.CheckYearRange(year, era)
        Return 0
    End Function

    Public Overrides Function GetMonth(ByVal time As DateTime) As Integer
        Return Me.GetDatePart(time.Ticks, 2)
    End Function

    Public Overrides Function GetMonthsInYear(ByVal year As Integer, ByVal era As Integer) As Integer
        HijriCalendar.CheckYearRange(year, era)
        Return 12
    End Function

    Public Overrides Function GetYear(ByVal time As DateTime) As Integer
        Return Me.GetDatePart(time.Ticks, 0)
    End Function

    Public Overrides Function IsLeapDay(ByVal year As Integer, ByVal month As Integer, ByVal day As Integer, ByVal era As Integer) As Boolean
        Dim num As Integer = Me.GetDaysInMonth(year, month, era)
        If ((day < 1) OrElse (day > num)) Then
            Throw New ArgumentOutOfRangeException("day", String.Empty) 'String.Format(CultureInfo.CurrentCulture, Environment.GetResourceString("ArgumentOutOfRange_Day"), num, month))
        End If
        Return ((Me.IsLeapYear(year, era) AndAlso (month = 12)) AndAlso (day = 30))
    End Function

    Public Overrides Function IsLeapMonth(ByVal year As Integer, ByVal month As Integer, ByVal era As Integer) As Boolean
        HijriCalendar.CheckYearMonthRange(year, month, era)
        Return False
    End Function

    Public Overrides Function IsLeapYear(ByVal year As Integer, ByVal era As Integer) As Boolean
        HijriCalendar.CheckYearRange(year, era)
        Return ((((year * 11) + 14) Mod 30) < 11)
    End Function

    Public Overrides Function ToDateTime(ByVal year As Integer, ByVal month As Integer, ByVal day As Integer, ByVal hour As Integer, ByVal minute As Integer, ByVal second As Integer, ByVal millisecond As Integer, ByVal era As Integer) As DateTime
        Dim num As Integer = Me.GetDaysInMonth(year, month, era)
        If ((day < 1) OrElse (day > num)) Then
            Throw New ArgumentOutOfRangeException("day", String.Empty) 'String.Format(CultureInfo.CurrentCulture, Environment.GetResourceString("ArgumentOutOfRange_Day"), num, month))
        End If
        Dim num2 As Long = Me.GetAbsoluteDateHijri(year, month, day)
        If (num2 < 0) Then
            Throw New ArgumentOutOfRangeException(Nothing, String.Empty) 'Environment.GetResourceString("ArgumentOutOfRange_BadYearMonthDay"))
        End If
        Return New DateTime(((num2 * &HC92A69C000) + CalendarTimeToTicks(hour, minute, second, millisecond)))
    End Function

    Public Overrides Function ToFourDigitYear(ByVal year As Integer) As Integer
        If (year < 0) Then
            Throw New ArgumentOutOfRangeException("year", String.Empty) 'Environment.GetResourceString("ArgumentOutOfRange_NeedNonNegNum"))
        End If
        If (year < 100) Then
            Return MyBase.ToFourDigitYear(year)
        End If
        If (year > &H25C2) Then
            Throw New ArgumentOutOfRangeException("year", String.Empty) 'String.Format(CultureInfo.CurrentCulture, Environment.GetResourceString("ArgumentOutOfRange_Range"), 1, &H25C2))
        End If
        Return year
    End Function


    ' Properties
    '<ComVisible(False)> _
    'Public Overrides ReadOnly Property AlgorithmType As CalendarAlgorithmType
    '    Get
    '        Return CalendarAlgorithmType.LunarCalendar
    '    End Get
    'End Property

    Protected ReadOnly Property DaysInYearBeforeMinSupportedYear As Integer
        Get
            Return &H162
        End Get
    End Property

    Public Overrides ReadOnly Property Eras As Integer()
        Get
            Return New Integer() {HijriCalendar.HijriEra}
        End Get
    End Property

    Public Property HijriAdjustment As Integer
        <SecuritySafeCritical> _
        Get
            If (Me.m_HijriAdvance = -2147483648) Then
                Me.m_HijriAdvance = HijriCalendar.GetAdvanceHijriDate
            End If
            Return Me.m_HijriAdvance
        End Get
        Set(ByVal value As Integer)
            If ((value < -2) OrElse (value > 2)) Then
                Throw New ArgumentOutOfRangeException("HijriAdjustment", String.Empty) 'String.Format(CultureInfo.CurrentCulture, Environment.GetResourceString("ArgumentOutOfRange_Bounds_Lower_Upper"), -2, 2))
            End If
            'MyBase.VerifyWritable()
            Me.m_HijriAdvance = value
        End Set
    End Property

    Friend ReadOnly Property ID As Integer
        Get
            Return 6
        End Get
    End Property

    <ComVisible(False)> _
    Public Overrides ReadOnly Property MaxSupportedDateTime As DateTime
        Get
            Return HijriCalendar.calendarMaxValue
        End Get
    End Property

    <ComVisible(False)> _
    Public Overrides ReadOnly Property MinSupportedDateTime As DateTime
        Get
            Return HijriCalendar.calendarMinValue
        End Get
    End Property

    Public Overrides Property TwoDigitYearMax As Integer
        Get
            If (MyBase.TwoDigitYearMax = -1) Then
                MyBase.TwoDigitYearMax = &H5AB 'Calendar.GetSystemTwoDigitYearSetting(Me.ID, &H5AB)
            End If
            Return MyBase.TwoDigitYearMax
        End Get
        Set(ByVal value As Integer)
            'MyBase.VerifyWritable()
            If ((value < &H63) OrElse (value > &H25C2)) Then
                Throw New ArgumentOutOfRangeException("value", String.Empty) 'String.Format(CultureInfo.CurrentCulture, Environment.GetResourceString("ArgumentOutOfRange_Range"), &H63, &H25C2))
            End If
            MyBase.TwoDigitYearMax = value
        End Set
    End Property


    ' Fields
    Friend Shared ReadOnly calendarMaxValue As DateTime = DateTime.MaxValue
    Friend Shared ReadOnly calendarMinValue As DateTime = New DateTime(&H26E, 7, &H12)
    Friend Const DatePartDay As Integer = 3
    Friend Const DatePartDayOfYear As Integer = 1
    Friend Const DatePartMonth As Integer = 2
    Friend Const DatePartYear As Integer = 0
    Private Const DEFAULT_TWO_DIGIT_YEAR_MAX As Integer = &H5AB
    Private Const HijriAdvanceRegKeyEntry As String = "AddHijriDate"
    Public Shared ReadOnly HijriEra As Integer = 1
    Friend Shared ReadOnly HijriMonthDays As Integer() = New Integer() {0, 30, &H3B, &H59, &H76, &H94, &HB1, &HCF, &HEC, &H10A, &H127, &H145, &H163}
    Private Const InternationalRegKey As String = "Control Panel\International"
    Private m_HijriAdvance As Integer = -2147483648
    Friend Const MaxAdvancedHijri As Integer = 2
    Friend Const MaxCalendarDay As Integer = 3
    Friend Const MaxCalendarMonth As Integer = 4
    Friend Const MaxCalendarYear As Integer = &H25C2
    Friend Const MinAdvancedHijri As Integer = -2
End Class
