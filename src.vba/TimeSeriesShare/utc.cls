Private oDate As Double
Private oFrequency  As String
Public dateString As String
Private Const TZ_KEY = "tz"

'INFO:: UPPER DATE LIMIT is 1/1/10000
'INFO:: LOWER DATE LIMIT is 1/1/100 because 1/1/99 tries to convert to 1/1/1999
'TODO working/business days, 5d

Property Let frequency(f As String)
    If Left(LCase(Trim(f)), Len(TZ_KEY)) = TZ_KEY Then
        oFrequency = TZ_KEY
    Else
        oFrequency = Left(LCase(Trim(f)), 1)
    End If
End Property
Property Get frequency() As String
    frequency = oFrequency
End Property

Property Get exactDate() As Double
    currentDate = oDate
End Property

Property Let exactDate(ddate As Double)
    oDate = CDbl(ddate)
End Property

Property Let exactUTC(ddate As LongLong)
     oDate = ddate / 86400 + 25569 ' no TimeZoneOffset because all utc dates should be provided as GMT
End Property

Public Function periodicity(ffrom As String, tto As String) As Long
num = 1
Select Case ffrom
    Case "m"
        Select Case tto
            Case "d", "h", "i", "s"
                num = MonthDays(getMonth, getYear)
        End Select
    Case "q"
        Select Case tto
            Case "d", "h", "i", "s"
                Dim mMonth As Long
                mMonth = getMonth
                mMonth = WorksheetFunction.RoundDown((mMonth - 1) / 3, 0) * 3 + 1
                num = MonthDays(mMonth, getYear) + MonthDays(mMonth + 1, getYear) + MonthDays(mMonth + 2, getYear)
            Case "m"
                num = 3
        End Select
    Case "a"
        Select Case tto
            Case "d", "h", "i", "s"
                num = YearDays(getYear)
            Case "m"
                num = 12
            Case "q"
                num = 4
        End Select
End Select

Select Case tto
    Case "h", "i", "s"
        num = num * 24
        Select Case tto
            Case "i", "s"
                num = num * 60
                If tto = "s" Then
                    num = num * 60
                End If
        End Select
End Select

periodicity = num
End Function

Public Function YearDays(lngyear As Long) As Long
    YearDays = DateDiff("d", DateSerial(lngyear, 1, 1), DateSerial(lngyear + 1, 1, 1))
End Function

Public Function MonthDays(myMonth As Long, myYear As Long) As Long
    MonthDays = Day(DateSerial(myYear, myMonth + 1, 1) - 1)
End Function

Public Function getMonth() As Long
    getMonth = Month(oDate)
End Function

Public Function getYear() As Long
    getYear = Year(oDate)
End Function

Public Function getDay() As Long
    getDay = Day(oDate)
End Function
Public Function getMinute() As Long
    getMinute = Minute(oDate)
End Function

Public Function getHour() As Long
    getHour = Hour(oDate)
End Function

Public Function getSecond() As Long
    getSecond = Second(oDate)
End Function

Public Function utcDate(Optional tzOffset As Double = 0) As LongLong
    utcDate = (oDate - 25569 - tzOffset) * 86400
End Function

Public Function getUTCByFrequency() As LongLong
    Dim n As utc, tzOffset As Double
    Set n = New utc
    macro = Array(1, 1, 1970)
    micro = Array(0, 0, 0)
    tzOffset = 0
    If oFrequency = TZ_KEY Then
        tzOffset = TimeOffset()
    Else
        If oFrequency <> "a" Then
            macro(0) = getMonth
            If oFrequency = "q" Then
                macro(0) = WorksheetFunction.RoundDown((macro(0) - 1) / 3, 0) * 3 + 1
            End If
        End If
        If oFrequency = "" Or InStr("qma", oFrequency) = 0 Then
            macro(1) = getDay
        End If
        If oFrequency = "" Or InStr("qmad", oFrequency) = 0 Then
            micro(0) = getHour
        End If
        If oFrequency = "" Or InStr("qmadh", oFrequency) = 0 Then
            micro(1) = getMinute
        End If
        If oFrequency = "" Or InStr("qmadhi", oFrequency) = 0 Then
            micro(2) = getSecond
        End If
        macro(2) = getYear
    End If
    Me.dateString = DateSerial(macro(2), macro(0), macro(1)) + TimeSerial(micro(0), micro(1), micro(2))
    n.exactDate = DateSerial(macro(2), macro(0), macro(1)) + TimeSerial(micro(0), micro(1), micro(2))
    getUTCByFrequency = n.utcDate(tzOffset)
End Function