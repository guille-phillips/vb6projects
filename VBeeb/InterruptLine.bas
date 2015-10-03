Attribute VB_Name = "InterruptLine"
Option Explicit

Public Enum IRQSource
    irqSystemVIA = 1
    irqUserVIA = 2
    irqACIA6850 = 4
End Enum

Public Enum NMISource
    nmiFDC8271 = 1
End Enum

Private mlIRQLine As Long
Private mlNMILine As Long

Public Sub InitialiseInterruptLine()
    mlIRQLine = 0
    mlNMILine = 0
End Sub

Public Sub SetIRQLine(ByVal irqsSource As IRQSource, ByVal lValue As Long)
    ' Debugging.WriteString "InterruptLine.SetIRQLine"
    
    mlIRQLine = (mlIRQLine Or irqsSource) Xor ((1 - lValue) * irqsSource)
    'Processor6502.IRQRaisedBy = mlIRQLine
    
    If mlIRQLine > 0 Then
        Processor6502.IRQFlag = True
    Else
        Processor6502.IRQFlag = False
    End If
End Sub

Public Sub SetNMILine(ByVal nmisSource As NMISource, ByVal lValue As Long)
    Dim lOriginalNMILine As Long
    
    ' Debugging.WriteString "InterruptLine.SetNMILine"
    
    lOriginalNMILine = Sgn(mlNMILine And nmisSource)
    mlNMILine = (mlNMILine Or nmisSource) Xor ((1 - lValue) * nmisSource)
    
    If lOriginalNMILine = 0 And lValue = 1 Then   ' negative edge but we inverse for convenience
        Processor6502.NMIFlag = True
    End If
End Sub

'Public Sub SetNMILineWithDescription(ByVal nmisSource As NMISource, ByVal lValue As Long, ByVal sDescription As String)
'    Dim lOriginalNMILine As Long
'
'    ' Debugging.WriteString "InterruptLine.SetNMILine"
'
'    lOriginalNMILine = Sgn(mlNMILine And nmisSource)
'    mlNMILine = (mlNMILine Or nmisSource) Xor ((1 - lValue) * nmisSource)
'
'    If lOriginalNMILine = 0 And lValue = 1 Then   ' negative edge but we inverse for convenience
'        Processor6502.NMIFlag = True
'        Processor6502.NMISource = sDescription
'    End If
'End Sub
