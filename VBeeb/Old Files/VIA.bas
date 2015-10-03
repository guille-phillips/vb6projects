Attribute VB_Name = "UserVIA6522"
Option Explicit

Public ORA As Integer
Public ORB As Integer
Public IRA As Integer
Public IRB As Integer
Public DDRA As Integer
Public DDRB As Integer
Public SHR As Integer
Public ACR As Integer
Public PCR As Integer
Public IFR As Integer
Public IER As Integer

Public mlTimer1 As Long
Public mlTimer1Latch As Long
Public mbTimer1HasInterrupted As Boolean
Public mlTimer2 As Long
Public mlTimer2Latch As Long
Public mbTimer2HasInterrupted As Boolean

Private mlLatchAddress As Long

' IFR
Private CA2 As Long
Private CA1 As Long
Private SHIFTREG As Long
Private CB2 As Long
Private CB1 As Long
Private TIMER2 As Long
Private TIMER1 As Long
Private MASTER As Long

' IER
Private CA2Enable As Long
Private CA1Enable As Long
Private SHIFTREGEnable As Long
Private CB2Enable As Long
Private CB1Enable As Long
Private TIMER2Enable As Long
Private TIMER1Enable As Long
Private MASTEREnable As Long

' ACR
Private mlTimer1Control As Long
Private mlTimer2Control As Long
Private mlShiftRegisterControl As Long
Private mlLatchEnable As Long

Public Sub InitialiseUserVIA()
    Dim lAddress As Long
    
    For lAddress = &HFE60& To &HFE6F&
        gyMem(lAddress) = 0
    Next
    
    mbTimer1HasInterrupted = False
    mbTimer2HasInterrupted = False
End Sub

Public Sub AssertCA2(Optional ByVal lBit As Long = 1)
    If lBit = 1 Then
        IFR = IFR Or &H81
        gyMem(&HFE4D&) = IFR
        If CA2Enable And MASTEREnable Then
            Processor6502.IRQFlag = True
        End If
    Else
        IFR = IFR And &HFE&
        If IFR = &H80 Then IFR = 0
        gyMem(&HFE4D&) = IFR
    End If
End Sub

Public Sub AssertCA1(Optional ByVal lBit As Long = 1)
    If lBit = 1 Then
        IFR = IFR Or &H82
        gyMem(&HFE4D&) = IFR
        If CA1Enable And MASTEREnable Then
            Processor6502.IRQFlag = True
        End If
    Else
        IFR = IFR And &HFD&
        If IFR = &H80 Then IFR = 0
        gyMem(&HFE4D&) = IFR
    End If
End Sub

Public Sub AssertSHIFTREG()
    IFR = IFR Or &H84
    gyMem(&HFE6D&) = IFR
    If SHIFTREGEnable And MASTEREnable Then
        Processor6502.IRQFlag = True
    End If
End Sub

Public Sub AssertCB2()
    IFR = IFR Or &H88
    gyMem(&HFE6D&) = IFR
    If CB2Enable And MASTEREnable Then
        Processor6502.IRQFlag = True
    End If
End Sub

Public Sub AssertCB1()
    IFR = IFR Or &H90
    gyMem(&HFE6D&) = IFR
    If CB1Enable And MASTEREnable Then
        Processor6502.IRQFlag = True
    End If
End Sub

Public Sub AssertTIMER2()
    IFR = IFR Or &HA0
    gyMem(&HFE6D&) = IFR
    If TIMER2Enable And MASTEREnable Then
        Processor6502.IRQFlag = True
    End If
End Sub

Public Sub AssertTIMER1()
    IFR = IFR Or &HC0
    gyMem(&HFE6D&) = IFR
    If TIMER1Enable And MASTEREnable Then
        Processor6502.IRQFlag = True
    End If
End Sub

Public Sub TimersTick(ByVal lCycles As Long)
    mlTimer1 = mlTimer1 - lCycles
    mlTimer2 = mlTimer2 - lCycles
    
    If mlTimer1 <= 0 Then
        If mlTimer1Control And 1 Then ' Continuous Interrupts
            mlTimer1 = (mlTimer1 + mlTimer1Latch) And &HFFFF&
            AssertTIMER1
        Else
            mlTimer1 = mlTimer1 + &H10000
            
            If Not mbTimer1HasInterrupted Then
                mbTimer1HasInterrupted = True
                AssertTIMER1
            End If
        End If
    End If
    
    CopyMemory gyMem(&HFE64&), mlTimer1, 2&

    
    If mlTimer2 <= 0 Then
        mlTimer2 = mlTimer1 + &H10000
        If Not mbTimer2HasInterrupted Then
            mbTimer2HasInterrupted = True
            AssertTIMER2
        End If
    End If
    
    CopyMemory gyMem(&HFE68&), mlTimer2, 2&
End Sub


Public Sub WriteRegister(ByVal lRegister As Long, ByVal lValue As Long)
    Dim lLatchAddress As Long
    Dim lSlowDataBit As Long
    
    Select Case lRegister
        Case &H0 ' IORB
            ORB = lValue And DDRB
            gyMem(&HFE60&) = ORB
        Case &H1, &HF ' IORA
            ORA = lValue And DDRA
            gyMem(&HFE61&) = ORA
        Case &H2 ' DDRB
            DDRB = lValue
            gyMem(&HFE62&) = DDRB
        Case &H3 ' DDRA
            DDRA = lValue
            gyMem(&HFE63&) = DDRA
        Case &H4, &H6 ' T1CL / T1CH
            mlTimer1Latch = (mlTimer1Latch And &HFF00&) + lValue
        Case &H5 ' T1CH
            mlTimer1Latch = (mlTimer1Latch And &HFF&) + lValue * 256&
            mlTimer1 = mlTimer1Latch
            gyMem(&HFE6D&) = gyMem(&HFE6D&) And &HBF& ' CLEAR IFR FOR TIMER1
            mbTimer1HasInterrupted = False ' RESET ONE SHOT
        Case &H7 ' T1LH
            mlTimer1Latch = (mlTimer1Latch And &HFF&) + lValue * 256&
        Case &H8 ' T2CL
            mlTimer2Latch = lValue
        Case &H9 ' T2CH
            mlTimer2 = mlTimer2Latch + lValue * 256&
            mbTimer2HasInterrupted = False
            IFR = IFR And &HDF&
            gyMem(&HFE6D&) = IFR ' CLEAR IFR FOR TIMER2
        Case &HA ' SR
        Case &HB ' ACR
            mlTimer1Control = (lValue And &HC0&) \ 64&
            mlTimer2Control = (lValue And &H20&) \ 32&
            mlShiftRegisterControl = (lValue And &H1C&) \ 4&
            mlLatchEnable = lValue And &H3&
            ACR = lValue
            gyMem(&HFE6B&) = lValue
        Case &HC ' PCR
            PCR = lValue
            gyMem(&HFE6C&) = lValue
        Case &HD ' IFR
            If lValue And 128& Then
                IFR = IFR Or (lValue And &H7F&)
            Else
                IFR = IFR And (lValue Xor &H7F&)
            End If
            gyMem(&HFE6D&) = IFR
            If False Then Stop
        Case &HE ' IER
            If lValue And 128& Then
                IER = IER Or (lValue And &H7F&)
            Else
                IER = IER And (lValue Xor &H7F&)
            End If
            
            CA2Enable = IER And 1
            CA1Enable = -((IER And 2) <> 0)
            SHIFTREGEnable = -((IER And 4) <> 0)
            CB2Enable = -((IER And 8) <> 0)
            CB1Enable = -((IER And 16) <> 0)
            TIMER2Enable = -((IER And 32) <> 0)
            TIMER1Enable = -((IER And 64) <> 0)
            MASTEREnable = -((IER And 128) <> 0)
            MASTEREnable = 1
            gyMem(&HFE6E&) = IER Or 128&
    End Select
End Sub

Public Function ReadRegister(ByVal lRegister As Long) As Long
    Select Case lRegister
        Case &H4&
            WriteRegister &HD&, &H40& ' Reset timer 1 interrrupt
            ReadRegister = gyMem(&HFE64&)
        Case &H8&
            WriteRegister &HD&, &H20& ' Reset timer 2 interrupt
            ReadRegister = gyMem(&HFE68&)
    End Select
End Function

