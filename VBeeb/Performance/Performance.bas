Attribute VB_Name = "Performance"
Option Explicit

Private Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Currency) As Long
Private Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Currency) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Private cFrequency As Currency
Private cCounters(0 To 5) As Currency


Private N As Long
Private Z As Long
Private A As Long
Private V As Long
Private C As Long

Public Sub StartCounter(Optional lCounterIndex As Long)
    QueryPerformanceFrequency cFrequency
    QueryPerformanceCounter cCounters(lCounterIndex)
End Sub

Public Function GetCounter(Optional lCounterIndex As Long) As Double
    Dim cCount As Currency
    
    QueryPerformanceFrequency cFrequency
    QueryPerformanceCounter cCount
    GetCounter = Format$((cCount - cCounters(lCounterIndex) - CCur(0.0008)) / cFrequency, "0.000000000")
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
