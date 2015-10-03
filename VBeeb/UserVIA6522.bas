Attribute VB_Name = "UserVIA6522"
Option Explicit

Public ORA As Long
Public ORB As Long
Public IRA As Long
Public IRB As Long
Public DDRA As Long
Public DDRB As Long
Public SHR As Long
Public ACR As Long
Public PCR As Long
Public IFR As Long
Public IER As Long

Public mlTimer1 As Long
Public mlTimer1Latch As Long
Public mbTimer1HasInterrupted As Boolean
Private mlTimer1Fraction As Long

Public mlTimer2 As Long
Public mlTimer2Latch As Long
Public mbTimer2HasInterrupted As Boolean
Private mlTimer2Fraction As Long

' ACR
Private mlTimer1Control As Long
Private mlTimer2Control As Long
Private mlShiftRegisterControl As Long
Private mlLatchEnable As Long

Public Sub ResetUserVIA()
    Dim lAddress As Long
    
    ' Debugging.WriteString "UserVIA6522.ResetUserVIA"
    
    For lAddress = &HFE60& To &HFE7F&
        gyMem(lAddress) = 0
    Next
    
    ORB = &HFF&
    ORA = &HFF&
    IRB = &HFF&
    IRA = &HFF&
    
    DDRA = &HFF&
    DDRB = &HFF&
    
    WriteRegister &HD&, &H7F&
    WriteRegister &HE&, &H7F&
    
    mlTimer1 = 0
    mlTimer2 = 0
    mlTimer1Latch = &HFFFF&
    mlTimer2Latch = &HFFFF&

    mlTimer1Control = 0
    mlTimer2Control = 0
    mbTimer1HasInterrupted = True
    mbTimer2HasInterrupted = True
End Sub


Public Sub AssertCA2(Optional ByVal lBit As Long = 1)
    ' Debugging.WriteString "UserVIA6522.AssertCA2"

    WriteIFR 1& Or lBit * &H80&
End Sub

Public Sub AssertCA1(Optional ByVal lBit As Long = 1)
    ' Debugging.WriteString "UserVIA6522.AssertCA1"

    WriteIFR 2& Or lBit * &H80&
End Sub

Public Sub AssertSHIFTREG(Optional ByVal lBit As Long = 1)
    ' Debugging.WriteString "UserVIA6522.AssertSHIFTREG"

    WriteIFR 4& Or lBit * &H80&
End Sub

Public Sub AssertCB2(Optional ByVal lBit As Long = 1)
    ' Debugging.WriteString "UserVIA6522.AssertCB2"

    WriteIFR 8& Or lBit * &H80&
End Sub

Public Sub AssertCB1(Optional ByVal lBit As Long = 1)
    ' Debugging.WriteString "UserVIA6522.AssertCB1"

    WriteIFR &H10& Or lBit * &H80&
End Sub

Public Sub AssertTIMER2(Optional ByVal lBit As Long = 1)
    ' Debugging.WriteString "UserVIA6522.AssertTIMER2"

    WriteIFR &H20& Or lBit * &H80&
End Sub

Public Sub AssertTIMER1(Optional ByVal lBit As Long = 1)
    ' Debugging.WriteString "UserVIA6522.AssertTIMER1"

    WriteIFR &H40& Or lBit * &H80&
End Sub


Public Sub TimersTick(ByVal lCycles As Long)
    ' Debugging.WriteString "UserVIA6522.TimersTick"
    Dim lHalfCycle As Long
    Dim lHalfCycles As Long
    
    lHalfCycle = lCycles And &H1&
    mlTimer1Fraction = mlTimer1Fraction + lHalfCycle
    mlTimer2Fraction = mlTimer2Fraction + lHalfCycle
    
    lHalfCycles = lCycles \ 2
    mlTimer1 = mlTimer1 - lHalfCycles
    mlTimer2 = mlTimer2 - lHalfCycles
    
    If mlTimer1Fraction = 2 Then
        mlTimer1Fraction = 0
        mlTimer1 = mlTimer1 - 1
    End If
    
    If mlTimer2Fraction = 2 Then
        mlTimer2Fraction = 0
        mlTimer2 = mlTimer2 - 1
    End If
    
    If mlTimer1 <= 0 Then
        If mlTimer1Control And 1 Then ' Continuous Interrupts
            mlTimer1 = (mlTimer1 + mlTimer1Latch) And &HFFFF&
            If mlTimer1Control And &H2& Then
                ORB = ORB Xor &H80&
            End If
            AssertTIMER1
        Else
            mlTimer1 = mlTimer1 + &H10000
            
            If Not mbTimer1HasInterrupted Then
                mbTimer1HasInterrupted = True
                If mlTimer1Control And &H2& Then
                    ORB = ORB Or &H80&
                End If
                AssertTIMER1
            End If
        End If
    End If
    
    CopyMemory gyMem(&HFE64&), mlTimer1, 2&
    CopyMemory gyMem(&HFE74&), mlTimer1, 2&
    
    If mlTimer2 <= 0 Then
        mlTimer2 = mlTimer2 + &H10000
        If Not mbTimer2HasInterrupted Then
            mbTimer2HasInterrupted = True
            AssertTIMER2
        End If
    End If
    
    CopyMemory gyMem(&HFE68&), mlTimer2, 2&
    CopyMemory gyMem(&HFE78&), mlTimer2, 2&
End Sub


Public Sub WriteRegister(ByVal lRegister As Long, ByVal lValue As Long)
    Dim lLatchAddress As Long
    Dim lSlowDataBit As Long
    Dim lIRQMask As Long
    
    ' Debugging.WriteString "UserVIA6522.WriteRegister"
    
    Select Case lRegister
        Case &H0 ' ORB
            ORB = lValue
            gyMem(&HFE60&) = ORB
            gyMem(&HFE70&) = ORB
        Case &H1, &HF ' ORA
            ORA = lValue
            gyMem(&HFE61&) = ORA
            gyMem(&HFE71&) = ORA
        Case &H2 ' DDRB
            DDRB = lValue
            gyMem(&HFE62&) = DDRB
            gyMem(&HFE72&) = DDRB
        Case &H3 ' DDRA
            DDRA = lValue
            gyMem(&HFE63&) = DDRA
            gyMem(&HFE73&) = DDRA
        Case &H4, &H6 ' T1CL / T1CH
            mlTimer1Latch = (mlTimer1Latch And &HFF00&) + lValue
        Case &H5 ' T1CH
            mlTimer1Latch = (mlTimer1Latch And &HFF&) + lValue * 256&
            mlTimer1 = mlTimer1Latch
            mlTimer1Fraction = 0
            If mlTimer1Control And &H2 Then
                ORB = ORB And &H7F&
            End If
            AssertTIMER1 0
            mbTimer1HasInterrupted = False ' RESET ONE SHOT
        Case &H7 ' T1LH
            mlTimer1Latch = (mlTimer1Latch And &HFF&) + lValue * 256&
        Case &H8 ' T2CL
            mlTimer2Latch = lValue
        Case &H9 ' T2CH
            mlTimer2 = mlTimer2Latch + lValue * 256&
            mlTimer2Fraction = 0
            mbTimer2HasInterrupted = False
            AssertTIMER2 0
        Case &HA ' SR
        Case &HB ' ACR
            mlTimer1Control = (lValue And &HC0&) \ 64&
            mlTimer2Control = (lValue And &H20&) \ 32&
            mlShiftRegisterControl = (lValue And &H1C&) \ 4&
            mlLatchEnable = lValue And &H3&
            ACR = lValue
            gyMem(&HFE6B&) = lValue
            gyMem(&HFE7B&) = lValue
        Case &HC ' PCR
            PCR = lValue
            gyMem(&HFE6C&) = lValue
            gyMem(&HFE7C&) = lValue
        Case &HD ' IFR
            WriteIFR lValue And &H7F& ' always clear bits - interrupts
        Case &HE ' IER
            ' Set/reset bits
            If lValue And 128& Then
                IER = IER Or (lValue And &H7F&)
            Else
                IER = IER And (lValue Xor &H7F&)
            End If
            gyMem(&HFE6E&) = IER Or &H80&
            gyMem(&HFE7E&) = IER Or &H80&
            
            ' Set IRQ
            lIRQMask = IFR And IER
            
            If (lIRQMask And &H7F&) = 0 Then
                InterruptLine.SetIRQLine irqUserVIA, 0
            Else
                InterruptLine.SetIRQLine irqUserVIA, 1
            End If
    End Select
End Sub

Private Function WriteIFR(ByVal lValue As Long)
    Dim lIRQMask As Long
    
    ' Set/reset bits
    If lValue And 128& Then
        IFR = IFR Or (lValue And &H7F&)
    Else
        IFR = IFR And (lValue Xor &H7F&)
    End If
    
    ' Set master IRQ condition
    If (IFR And &H7F&) = 0 Then
        IFR = 0
    Else
        IFR = IFR Or &H80&
    End If
    gyMem(&HFE6D&) = IFR
    gyMem(&HFE7D&) = IFR
    
    ' Set IRQ
    lIRQMask = IFR And IER
    
    If (lIRQMask And &H7F&) = 0 Then
        InterruptLine.SetIRQLine irqUserVIA, 0
    Else
        InterruptLine.SetIRQLine irqUserVIA, 1
    End If
End Function

Public Function ReadRegister(ByVal lRegister As Long) As Long
    ' Debugging.WriteString "UserVIA6522.ReadRegister"
    
    Select Case lRegister
        Case 0
            ReadRegister = (ORB And DDRB) Or (IRB And (255 - DDRB))
        Case &H4&
            WriteRegister &HD&, &H40& ' Reset timer 1 interrrupt
            ReadRegister = gyMem(&HFE64&)
        Case &H8&
            WriteRegister &HD&, &H20& ' Reset timer 2 interrupt
            ReadRegister = gyMem(&HFE68&)
    End Select
End Function

