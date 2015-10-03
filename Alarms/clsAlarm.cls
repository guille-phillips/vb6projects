VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAlarm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mdEventTime As Date
Public ReminderFrom As Date
Public RemindEvery As Single
Public RemindEveryType As RecurTypes
Public RecurEvery As Single
Public RecurEveryType As RecurTypes
Public EventDescription As String
Public MostRecentAlarmTime As Date
Public Expired As Boolean
Public OnHold As Boolean
Public Duplicated As Boolean
Public GUID As String
Public RecurIndex As Long
Public LastUpdatedDate As Date
Public OwnerID As String
Public Deleted As Boolean

Private mbExpired As Boolean

Public Property Let EventTime(ByVal dEventTime As Date)
    mdEventTime = dEventTime
End Property

Public Property Get EventTime() As Date
    EventTime = mdEventTime
End Property

Public Property Get TimeLeft(dTime As Date) As String
    Dim lYears As Long
    Dim lMonths As Long
    Dim lWeeks As Long
    Dim lDays As Long
    Dim lHours As Long
    Dim lMinutes As Long
    Dim lSeconds As Long
    
    Dim lSecondsDiff As Long
    Dim lMinutesDiff As Long
    Dim lHoursDiff As Long
    Dim lDaysDiff As Long
    Dim lWeeksDiff As Long
    Dim vMonthWeekDiff As Variant
    Dim lYearsDiff As Long
    
    Dim bShowYears As Boolean
    Dim bShowMonths As Boolean
    Dim bShowWeeks As Boolean
    Dim bShowDays As Boolean
    Dim bShowHours As Boolean
    Dim bShowMinutes As Boolean
    Dim bShowSeconds As Boolean

    Dim lDiff As Long
    Dim bOverdue As Boolean
    
    If Expired Then
        Exit Property
    End If
    
    If dTime < mdEventTime Then
        lSecondsDiff = DateDiff("s", dTime, mdEventTime)
        lMinutesDiff = DateDiff("n", dTime, mdEventTime)
        lHoursDiff = DateDiff("h", dTime, mdEventTime)
        lDaysDiff = DateDiff("d", dTime, mdEventTime)
        vMonthWeekDiff = MonthWeekDiff(dTime, mdEventTime)
        lYearsDiff = DateDiff("yyyy", dTime, mdEventTime)
    Else
        lSecondsDiff = DateDiff("s", mdEventTime, dTime)
        lMinutesDiff = DateDiff("n", mdEventTime, dTime)
        lHoursDiff = DateDiff("h", mdEventTime, dTime)
        lDaysDiff = DateDiff("d", mdEventTime, dTime)
        lWeeksDiff = DateDiff("ww", mdEventTime, dTime)
        vMonthWeekDiff = MonthWeekDiff(mdEventTime, dTime)
        lYearsDiff = DateDiff("yyyy", mdEventTime, dTime)
        
        bOverdue = True
    End If

    If lSecondsDiff < 120 Then
        bShowSeconds = True
        If lSecondsDiff >= 60 Then
            bShowMinutes = True
            lMinutes = 1
        End If
        lSeconds = lSecondsDiff Mod 60
    ElseIf lMinutesDiff < 120 Then
        bShowMinutes = True
        If lMinutesDiff >= 60 Then
            bShowHours = True
            lHours = 1
        End If
        lMinutes = lMinutesDiff Mod 60
    ElseIf lHoursDiff < 48 Then
        bShowHours = True
        If lHoursDiff >= 24 Then
            bShowDays = True
            lDays = 1
        End If
        lHours = lHoursDiff Mod 24
    ElseIf lDaysDiff < 14 Then
        bShowDays = True
        If lDaysDiff >= 7 Then
            bShowWeeks = True
            lWeeks = 1
        End If
        lDays = lDaysDiff Mod 7
    ElseIf vMonthWeekDiff(0) = 0 Then
        bShowWeeks = True
        lWeeks = vMonthWeekDiff(1)
    ElseIf vMonthWeekDiff(0) = 1 Then
        bShowWeeks = True
        bShowMonths = True
        lMonths = 1
        lWeeks = vMonthWeekDiff(1)
    ElseIf vMonthWeekDiff(0) < 24 Then
        bShowMonths = True
        If vMonthWeekDiff(0) >= 12 Then
            bShowYears = True
            lYears = 1
        End If
        lMonths = vMonthWeekDiff(0) Mod 12
    Else
        bShowYears = True
        lYears = lYearsDiff
    End If
       
    If bShowYears Then
        TimeLeft = TimeLeft & Format$(lYears, "0") & Pluralise(lYears, " year") & " "
    End If
    
    If bShowMonths Then
        TimeLeft = TimeLeft & Format$(lMonths, "0") & Pluralise(lMonths, " month") & " "
    End If
    
    If bShowWeeks Then
        TimeLeft = TimeLeft & Format$(lWeeks, "0") & Pluralise(lWeeks, " week") & " "
    End If
    
    If bShowDays Then
        TimeLeft = TimeLeft & Format$(lDays, "0") & Pluralise(lDays, " day") & " "
    End If
    
    If bShowHours Then
        TimeLeft = TimeLeft & Format$(lHours, "0") & Pluralise(lHours, " hour") & " "
    End If
    
    If bShowMinutes Then
        TimeLeft = TimeLeft & Format$(lMinutes, "0") & Pluralise(lMinutes, " min") & " "
    End If
    
    If bShowSeconds Then
        TimeLeft = TimeLeft & Format$(lSeconds, "00") & Pluralise(lSeconds, " sec") & " "
    End If
    If bOverdue Then
        TimeLeft = "overdue " & TimeLeft
    End If
End Property

Private Function Pluralise(ByVal lValue As Long, ByVal sString As String) As String
    If lValue <> 1 Then
        Pluralise = sString & "s"
    Else
        Pluralise = sString
    End If
End Function

Private Function MonthWeekDiff(dBaseDate As Date, dTargetDate As Date) As Variant
    Dim lMonthDiff As Long
    Dim lBaseDay As Long
    Dim lBaseMonth As Long
    Dim lBaseYear As Long
    Dim lTargetDay As Long
    Dim lTargetMonth As Long
    Dim lTargetYear As Long
    
    Dim dNextMonth As Date
    
    lMonthDiff = DateDiff("m", dBaseDate, dTargetDate)
    
    lBaseDay = Day(dBaseDate)
    lBaseMonth = Month(dBaseDate)
    lBaseYear = Year(dBaseDate)
    lTargetDay = Day(dTargetDate)
    lTargetMonth = Month(dTargetDate)
    lTargetYear = Year(dTargetDate)
    
    If lMonthDiff >= 2 Then
        MonthWeekDiff = Array(lMonthDiff, 0)
    Else
        If lTargetMonth = lBaseMonth Then
            MonthWeekDiff = Array(0, DateDiff("ww", dBaseDate, dTargetDate, vbMonday))
        Else
            If lBaseDay > lTargetDay Then
               MonthWeekDiff = Array(0, DateDiff("ww", dBaseDate, dTargetDate, vbMonday))
            Else
                dNextMonth = DateSerial(lBaseYear + IIf(lBaseMonth = 12, 1, 0), lBaseMonth + 1 - IIf(lBaseMonth = 12, 12, 0), lBaseDay)
                MonthWeekDiff = Array(1, DateDiff("ww", dNextMonth, dTargetDate, vbMonday))
            End If
        End If
    End If
End Function

Private Function YearDiff(dBaseDate As Date, dTargetDate As Date) As Long
    YearDiff = Year(dTargetDate) - Year(dBaseDate)
    If Month(dTargetDate) < Month(dBaseDate) Then
        YearDiff = YearDiff - 1
    ElseIf Month(dTargetDate) = Month(dBaseDate) Then
        If Day(dTargetDate) < Day(dBaseDate) Then
            YearDiff = YearDiff - 1
        ElseIf Day(dTargetDate) = Day(dBaseDate) Then
            If Hour(dTargetDate) < Hour(dBaseDate) Then
                YearDiff = YearDiff - 1
            ElseIf Hour(dTargetDate) = Hour(dBaseDate) Then
                If Minute(dTargetDate) < Minute(dBaseDate) Then
                    YearDiff = YearDiff - 1
                ElseIf Minute(dTargetDate) = Minute(dBaseDate) Then
                    If Second(dTargetDate) < Second(dBaseDate) Then
                        YearDiff = YearDiff - 1
                    End If
                End If
            End If
        End If
    End If
End Function

Public Function CheckAlarm(dTime As Date) As String
    Dim lRecurTimeSecs As Long
    Dim dNearestTime As Date
    Dim dFrom As Date
    
    Dim dBaseTime As Date

    If Expired Or OnHold Then
        Exit Function
    End If
    
    If ReminderFrom = 0 Then
        dBaseTime = mdEventTime
    Else
        dBaseTime = ReminderFrom
    End If
    
    Do
        If dTime >= MostRecentAlarmTime Then
            CheckAlarm = EventDescription
    
            If RemindEvery > 0 Then
                MostRecentAlarmTime = NextAlarm(dTime, dBaseTime, RemindEveryType, RemindEvery)
            Else
                MostRecentAlarmTime = DateAdd("y", 100, MostRecentAlarmTime)
            End If
        End If
        
        If dTime < mdEventTime And MostRecentAlarmTime > mdEventTime Then
            MostRecentAlarmTime = mdEventTime
        End If
    Loop Until MostRecentAlarmTime > dTime
End Function

Private Function NextAlarm(ByVal dTime As Date, ByVal dBaseTime As Date, ByVal rtRecurType As RecurTypes, ByVal lValue As Long) As Date
    Dim vIntervals As Variant
    Dim lBaseMonthYear As Long
    Dim lTimeMonthYear As Long
    Dim lThisMonthYear As Long
    Dim dThisDate As Date
    Dim lDayCount As Long
    
    vIntervals = Array("n", "h", "d", "ww", "", "", "", "", "", "", "", "", "", "m", "yyyy")
    
    Select Case rtRecurType
        Case rtMinute, rtHour, rtDay, rtWeek, rtMonth, rtYear
            dThisDate = dBaseTime
            While dThisDate <= dTime
                dThisDate = DateAdd(vIntervals(rtRecurType), lValue, dThisDate)
            Wend
            NextAlarm = dThisDate
        Case Is <= rtSunday
            dThisDate = dBaseTime
            If Weekday(dThisDate, vbMonday) = (rtRecurType - rtMonday + 1) Then
                dThisDate = DateAdd("ww", lValue, dThisDate)
            Else
                While Weekday(dThisDate, vbMonday) <> (rtRecurType - rtMonday + 1)
                    dThisDate = DateAdd("d", 1, dThisDate)
                Wend
                dThisDate = DateAdd("ww", lValue - 1, dThisDate)
                While dThisDate <= dTime
                    dThisDate = DateAdd("ww", 1, dThisDate)
                Wend
            End If
            NextAlarm = dThisDate
        Case rtWeekday
            dThisDate = dBaseTime
            While dThisDate < dTime
                lDayCount = lValue

                While lDayCount > 0
                    dThisDate = DateAdd("d", 1, dThisDate)
                    If Weekday(dThisDate, vbSaturday) >= 3 Then
                        lDayCount = lDayCount - 1
                    End If
                Wend
            Wend
            NextAlarm = dThisDate
        Case rtWeekend
            dThisDate = dBaseTime
            While dThisDate < dTime
                lDayCount = lValue
                While lDayCount > 0
                    dThisDate = DateAdd("d", 1, dThisDate)
                    If Weekday(dThisDate, vbSaturday) < 3 Then
                        lDayCount = lDayCount - 1
                    End If
                Wend
            Wend
            NextAlarm = dThisDate
    End Select
End Function

Public Function Duplicate() As clsAlarm
    Set Duplicate = New clsAlarm
    
    Duplicate.EventDescription = EventDescription
    
    Duplicate.LastUpdatedDate = Now
    Duplicate.RemindEvery = RemindEvery
    Duplicate.RemindEveryType = RemindEveryType
    
    Duplicate.RecurEvery = RecurEvery
    Duplicate.RecurEveryType = RecurEveryType

    Duplicate.EventTime = NextAlarm(Now, EventTime, RecurEveryType, RecurEvery)
    Duplicate.MostRecentAlarmTime = Duplicate.EventTime
    If ReminderFrom <> 0 Then
        Duplicate.ReminderFrom = NextAlarm(ReminderFrom, ReminderFrom, RecurEveryType, RecurEvery)
    End If
    Duplicate.GUID = StringFromGUID(StepGUID(GUIDFromString(GUID)))
    Duplicate.RecurIndex = 1
    Duplicate.OwnerID = OwnerID
End Function

Private Sub Class_Initialize()
    GUID = NewGUID
End Sub
