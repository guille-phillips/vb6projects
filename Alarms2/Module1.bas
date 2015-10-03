Attribute VB_Name = "Module1"
Option Explicit

Public Enum RecurTypes
    rtMinute
    rtHour
    rtDay
    rtWeek
    rtMonday
    rtTuesday
    rtWednesday
    rtThursday
    rtFriday
    rtSaturday
    rtSunday
    rtWeekday
    rtWeekend
    rtMonth
    rtYear
End Enum

Public goDateParser As ISaffronObject

Public Function ParseDateExpression(ByVal sDate As String) As Date
    Dim dToday As Date
    Dim oTree As SaffronTree
    Dim lWeekday As Long
    Dim lYear As Long
    Dim lMonth As Long
    Dim lDay As Long
    Dim lHour As Long
    Dim lMinute As Long
    Dim dTime As Date
    Dim lDayOffset As Long
    
    dToday = Now()
    
    SaffronStream.Text = sDate
    Set oTree = New SaffronTree
    If goDateParser.Parse(oTree) Then
        Select Case oTree(1).Index
            Case 1 ' Date Time
                ParseDateExpression = ParseTime(oTree(1)(1)(3)) + ParseDate(oTree(1)(1)(1), dToday)
            Case 2 ' Time Date
                ParseDateExpression = ParseTime(oTree(1)(1)(1)) + ParseDate(oTree(1)(1)(3), dToday)
            Case 3 ' Day Time
                lDay = oTree(1)(1)(1)(1).Index
                lDayOffset = lDay - Weekday(dToday, vbMonday)
                If lDayOffset < 0 Then
                    lDayOffset = lDayOffset + 7
                End If
                ParseDateExpression = DateSerial(Year(dToday), Month(dToday), Day(dToday)) + lDayOffset + ParseTime(oTree(1)(1)(3))
            Case 4 ' Time Day
                lDay = oTree(1)(1)(3)(1).Index
                lDayOffset = lDay - Weekday(dToday, vbMonday)
                If lDayOffset < 0 Then
                    lDayOffset = lDayOffset + 7
                End If
                ParseDateExpression = DateSerial(Year(dToday), Month(dToday), Day(dToday)) + lDayOffset + ParseTime(oTree(1)(1)(1))
            Case 5 ' Date
                ParseDateExpression = ParseDate(oTree(1)(1), dToday)
            Case 6 ' Day Name
                lDay = oTree(1)(1)(1).Index
                lDayOffset = lDay - Weekday(dToday, vbMonday)
                If lDayOffset < 0 Then
                    lDayOffset = lDayOffset + 7
                End If
                ParseDateExpression = DateSerial(Year(dToday), Month(dToday), Day(dToday)) + lDayOffset
            Case 7 ' Time
                dTime = ParseTime(oTree(1)(1))
                If Hour(dTime) > Hour(dToday) Then
                    ParseDateExpression = DateSerial(Year(dToday), Month(dToday), Day(dToday)) + dTime
                ElseIf Hour(dTime) = Hour(dToday) Then
                    If Minute(dTime) >= Minute(dToday) Then
                        ParseDateExpression = DateSerial(Year(dToday), Month(dToday), Day(dToday)) + dTime
                    Else
                        ParseDateExpression = DateSerial(Year(dToday), Month(dToday), Day(dToday)) + dTime + 1
                    End If
                Else
                   ParseDateExpression = DateSerial(Year(dToday), Month(dToday), Day(dToday)) + dTime + 1
                End If
            Case 8 ' Year
                ParseDateExpression = DateSerial(oTree.Text, 1, 1)
            Case 9 ' Day
                lDay = oTree.Text
                If lDay < Day(dToday) Then
                    If Month(dToday) < 12 Then
                        ParseDateExpression = DateSerial(Year(dToday), Month(dToday) + 1, lDay)
                    Else
                        ParseDateExpression = DateSerial(Year(dToday) + 1, 1, lDay)
                    End If
                Else
                    ParseDateExpression = DateSerial(Year(dToday), Month(dToday), lDay)
                End If
            Case 10 ' Month
                lMonth = ParseMonth(oTree(1)(1))
                If lMonth < Month(dToday) Then
                    ParseDateExpression = DateSerial(Year(dToday) + 1, lMonth, 1)
                Else
                    ParseDateExpression = DateSerial(Year(dToday), lMonth, 1)
                End If
        End Select
    End If
End Function

Public Function ParseDate(oTree As SaffronTree, dNow As Date) As Date
    Dim lDay As Long
    Dim lModifier As Long
    Dim lMonth As Long
    Dim lYear As Long
    
    Select Case oTree.Index
        Case 1 ' Date Short 6
            lDay = oTree(1)(1).Index
            lModifier = oTree(1)(3).Index
            If lDay > Weekday(dNow, vbMonday) Then
                ParseDate = DateSerial(Year(dNow), Month(dNow), Day(dNow)) + (lDay - Weekday(dNow, vbMonday)) + lModifier * 7
            Else
                ParseDate = DateSerial(Year(dNow), Month(dNow), Day(dNow)) + (Weekday(dNow, vbMonday) - lDay) + 7 + lModifier * 7
            End If
        
        Case 2 ' Date Long
            ParseDate = DateSerial(oTree(1)(6).Text, ParseMonth(oTree(1)(4)), oTree(1)(2).Text)
        Case 3 ' Unix Date
            ParseDate = DateSerial(oTree(1)(1).Text, ParseMonth(oTree(1)(3)), oTree(1)(5).Text)
        Case 4 ' Date Short 1
            lDay = oTree(1)(1).Text
            lMonth = ParseMonth(oTree(1)(3))
            
            If lMonth < Month(dNow) Then
                ParseDate = DateSerial(Year(dNow) + 1, lMonth, lDay)
            ElseIf lMonth = Month(dNow) Then
                If lDay < Day(dNow) Then
                    ParseDate = DateSerial(Year(dNow) + 1, lMonth, lDay)
                Else
                    ParseDate = DateSerial(Year(dNow), lMonth, lDay)
                End If
            Else
                ParseDate = DateSerial(Year(dNow), lMonth, lDay)
            End If
            
        Case 5 ' Date Short 3
            lMonth = ParseMonth(oTree(1)(1))
            lYear = oTree(1)(3).Text
            
            ParseDate = DateSerial(lYear, lMonth, 1)
        Case 6 ' Date Short 2
            lMonth = ParseMonth(oTree(1)(1))
            lDay = oTree(1)(3).Text

            If lMonth < Month(dNow) Then
                ParseDate = DateSerial(Year(dNow) + 1, lMonth, lDay)
            ElseIf lMonth = Month(dNow) Then
                If lDay < Day(dNow) Then
                    ParseDate = DateSerial(Year(dNow) + 1, lMonth, lDay)
                Else
                    ParseDate = DateSerial(Year(dNow), lMonth, lDay)
                End If
            Else
                ParseDate = DateSerial(Year(dNow), lMonth, lDay)
            End If
        
        Case 7 ' Relative
        Case 8 ' Date Short 4
        Case 9 ' Date Short 5
        Case 10 ' Other
            Select Case oTree(1).Index
                Case 1 ' today
                    ParseDate = DateSerial(Year(dNow), Month(dNow), Day(dNow))
                Case 2 ' tomorrow
                    ParseDate = DateSerial(Year(dNow), Month(dNow), Day(dNow)) + 1
                Case 3 ' yesterday
                    ParseDate = DateSerial(Year(dNow), Month(dNow), Day(dNow)) - 1
                Case 4 ' fortnight
                    ParseDate = DateSerial(Year(dNow), Month(dNow), Day(dNow)) + 14
                Case 5 ' week
                    ParseDate = DateSerial(Year(dNow), Month(dNow), Day(dNow)) + 7
            End Select
    End Select
End Function


Private Function ParseMonth(oTree As SaffronTree) As Long
    If oTree.Index < 4 Then
        ParseMonth = oTree.Text
    Else
        ParseMonth = oTree(1)(1).Index
    End If
End Function

Private Function ParseTime(oTree As SaffronTree) As Date
    ParseTime = TimeSerial(oTree(1).Text, oTree(3).Text, 0)
End Function


