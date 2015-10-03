Attribute VB_Name = "IRQLine"
Option Explicit

Public Enum IRQSource
    irqSystemVIA = 1
    irqUserVIA = 2
    irqACIA6850 = 4
End Enum

Private mlIRQLine As Long

Public Sub SetIRQLine(ByVal irqsSource As IRQSource, ByVal lValue As Long)
    mlIRQLine = (mlIRQLine Or irqsSource) Xor ((1 - lValue) * irqsSource)
    'Processor6502.IRQRaisedBy = mlIRQLine
    
    If mlIRQLine > 0 Then
        Processor6502.IRQFlag = True
    Else
        Processor6502.IRQFlag = False
    End If
End Sub
