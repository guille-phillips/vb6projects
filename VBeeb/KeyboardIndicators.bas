Attribute VB_Name = "KeyboardIndicators"
Option Explicit

Private mlTotalCycles As Long
Private mlLEDs As Long

Public Sub Tick(ByVal lCycles As Long)
    ' Debugging.WriteString "KeyboardIndicators.Tick"
        
    mlTotalCycles = mlTotalCycles + lCycles
    If mlTotalCycles > 10000# Then
        mlTotalCycles = 0
        Console.staBar.SimpleText = IIf((mlLEDs And 1) = 1, "CAPS LOCK ", "") & IIf((mlLEDs And 2) = 2, "SHIFT LOCK ", "") & IIf((mlLEDs And 4) = 4, "CASSETTE MOTOR ", "")
    End If
End Sub

Public Sub SetLED(ByVal lType As Long, ByVal lValue As Long)
    ' Debugging.WriteString "KeyboardIndicators.SetLED"
            
    mlLEDs = (mlLEDs And -(lType + 1)) Or lType * lValue
End Sub


