Attribute VB_Name = "Performance"
Option Explicit


Private mlFrequency As Long
Private cCounters(0 To 5) As Currency

Public Sub InitialiseCounters()
    Dim cFrequency As Currency
    QueryPerformanceFrequency cFrequency
    mlFrequency = cFrequency * 10000#
End Sub

Public Sub StartCounter(Optional lCounterIndex As Long)
    QueryPerformanceCounter cCounters(lCounterIndex)
End Sub

Public Function GetCounter(Optional lCounterIndex As Long) As Double
    Dim cCount As Currency
    
    QueryPerformanceCounter cCount
    GetCounter = 10000# * ((cCount - cCounters(lCounterIndex))) / mlFrequency
End Function

Public Function Scientific(ByVal dValue As Double) As String
    Dim lMultiplier As Long
    Dim vNames As Variant
    
    lMultiplier = 5
    vNames = Array("peta", "tera", "giga", "mega", "kilo", "", "milli", "micro", "nano", "pico", "femto")
    If Abs(dValue) < 1 Then
        While Abs(dValue) < 1
            dValue = dValue * 1000
            lMultiplier = lMultiplier + 1
        Wend
    ElseIf Abs(dValue) >= 1000 Then
        While Abs(dValue) >= 1000
            dValue = dValue / 1000
            lMultiplier = lMultiplier - 1
        Wend
    End If
    
    Scientific = Format$(dValue, "0.000") & " " & vNames(lMultiplier)
End Function


'Public Sub UpdateDisplayOrig()
'
'
''Dim lCount As Long
''Dim lCount2 As Long
''Dim dTime As Double
''Dim dFastestTime As Double
''Const lIterations As Long = 10000
''
''dFastestTime = 1000000#
''For lCount2 = 1 To 10
''    StartCounter
''    For lCount = 1 To lIterations
'
'
'    StretchDIBits mlConsoleHDC, 0&, mlCharacterRowBottomScanline, mlRowWidthStretched, -16&, 0&, 0&, mlRowWidth, 8&, myDisplayMemory(0), bmiCharacterRow, DIB_RGB_COLORS, SRCCOPY
'
''    Next
''    dTime = CDbl(GetCounter) / CDbl(lIterations)
''    If dTime < dFastestTime Then
''        dFastestTime = dTime
''    End If
''Next
''
''Debug.Print Scientific(dFastestTime)
''If Dir(App.Path & "\perf.txt") <> "" Then Kill App.Path & "\perf.txt"
''Open App.Path & "\perf.txt" For Binary As #2
''Put #2, , Scientific(dFastestTime)
''Close #2
''
''End
'
'End Sub


