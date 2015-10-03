Attribute VB_Name = "Throttle"
Option Explicit

Private mlTotalCycles As Long
Private mlThrottleLoop As Long

Private mdUnit As Double
Private mlDelay As Long
Private mlCounter As Long
Private mcFrequency As Currency
Private mdTicksPerSample As Double

Private mlCycleInterval  As Long
Private mcCounter As Currency
Private mlSampleRate As Long

Private mdSumXY As Double
Private mdSumX As Double
Private mdSumY As Double
Private mdSumXX As Double
Private mlPoints As Long

Private mlStep As Long

Public DefinedSpeedControl As Long
Public SpeedControl As Long
Public SpeedControlStep As Long

Public Sub InitialiseThrottle()
    ' Debugging.WriteString "Throttle.InitialiseThrottle"
    
    InitialiseFrequency
    mlSampleRate = 100 ' times per second
    mlCycleInterval = 2000000 / mlSampleRate  ' cycles per sample rate
    mlTotalCycles = mlCycleInterval
    mdTicksPerSample = mcFrequency / mlSampleRate ' ticks per sample
    QueryPerformanceCounter mcCounter
    DefinedSpeedControl = 10000
    SpeedControl = 10000
    SpeedControlStep = 1024
End Sub

Public Sub InitialiseFrequency()
    Dim cFrequency As Currency
    Dim yByte(7) As Byte
    Dim lIndex As Long
    Dim sShow As String

    ' Debugging.WriteString "Throttle.InitialiseFrequency"
    
'    yByte(0) = &H10
'    yByte(1) = &H7A
'    yByte(2) = &HAA
'    yByte(3) = &HCA
'
'    CopyMemory cFrequency, yByte(0), 8&
    
'    QueryPerformanceFrequency yByte(0)
'
'    For lIndex = 7 To 0 Step -1
'        sShow = sShow & HexNum(yByte(lIndex), 2) & " "
'    Next
'    MsgBox sShow
'    End
    
    QueryPerformanceFrequency cFrequency
    mcFrequency = cFrequency * 10000
End Sub

Public Sub ThrottleTick(ByVal lCycles As Long)
    Static mlExpectedSum As Long
    Static mlActualSum As Long

    Dim cCounter As Currency
    Dim lExpectedElapsedCount As Long
    Dim lActualElapsedCount As Long
    
    Dim lDelayCount As Long
    
    ' Debugging.WriteString "Throttle.ThrottleTick"
    
'    Exit Sub
    
    On Error GoTo Overflow
    
    mlTotalCycles = mlTotalCycles - lCycles
    If mlTotalCycles <= 0 Then
        lExpectedElapsedCount = (mlCycleInterval - mlTotalCycles) * mdTicksPerSample / mlCycleInterval
        Do
            lDelayCount = lDelayCount + 1
            QueryPerformanceCounter cCounter
            lActualElapsedCount = (cCounter - mcCounter) * SpeedControl
        Loop Until lActualElapsedCount >= lExpectedElapsedCount
        'Console.Caption = lDelayCount
        If lDelayCount = 1 Then
            Throttle.SpeedControl = Throttle.SpeedControl - SpeedControlStep
            If Throttle.SpeedControl < 100 Then
                Throttle.SpeedControl = 100
            End If
        Else
            SpeedControlStep = SpeedControlStep \ 2
            If SpeedControlStep = 0 Then
                SpeedControlStep = 1
            End If
            If Throttle.SpeedControl > DefinedSpeedControl Then
                Throttle.SpeedControl = Throttle.DefinedSpeedControl
            Else
                Throttle.SpeedControl = Throttle.SpeedControl + SpeedControlStep
            End If
        End If
        mcCounter = cCounter
        mlTotalCycles = mlCycleInterval
    End If
    Exit Sub
Overflow:
    QueryPerformanceCounter mcCounter
End Sub
