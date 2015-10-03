Attribute VB_Name = "SystemVIA6522"
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
Public mlTimer2Fraction As Long

Public mlLatchAddress As Long

' ACR
Private mlTimer1Control As Long
Private mlTimer2Control As Long
Private mlShiftRegisterControl As Long
Private mlLatchEnable As Long

Public LatchValue As Long

Public Sub ResetSystemVIA()
    Dim lAddress As Long
    
    ' Debugging.WriteString "SystemVIA6522.ResetSystemVIA"
    
    For lAddress = &HFE40& To &HFE5F&
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
    ' Debugging.WriteString "SystemVIA6522.AssertCA2"

    WriteIFR 1& Or lBit * &H80&
End Sub

Public Sub AssertCA1(Optional ByVal lBit As Long = 1)
    ' Debugging.WriteString "SystemVIA6522.AssertCA1"

    WriteIFR 2& Or lBit * &H80&
End Sub

Public Sub AssertSHIFTREG(Optional ByVal lBit As Long = 1)
    ' Debugging.WriteString "SystemVIA6522.AssertSHIFTREG"

    WriteIFR 4& Or lBit * &H80&
End Sub

Public Sub AssertCB2(Optional ByVal lBit As Long = 1)
    ' Debugging.WriteString "SystemVIA6522.AssertCB2"

    WriteIFR 8& Or lBit * &H80&
End Sub

Public Sub AssertCB1(Optional ByVal lBit As Long = 1)
    ' Debugging.WriteString "SystemVIA6522.AssertCB1"

    WriteIFR &H10& Or lBit * &H80&
End Sub

Public Sub AssertTIMER2(Optional ByVal lBit As Long = 1)
    ' Debugging.WriteString "SystemVIA6522.AssertTIMER2"

    WriteIFR &H20& Or lBit * &H80&
End Sub

Public Sub AssertTIMER1(Optional ByVal lBit As Long = 1)
    ' Debugging.WriteString "SystemVIA6522.AssertTIMER1"

    WriteIFR &H40& Or lBit * &H80&
End Sub

Public Sub TimersTick(ByVal lCycles As Long)
    ' Debugging.WriteString "SystemVIA6522.TimersTick"
    Dim lHalfCylcle As Long
    
    lHalfCylcle = lCycles And &H1&
    mlTimer1Fraction = mlTimer1Fraction + lHalfCylcle
    mlTimer2Fraction = mlTimer2Fraction + lHalfCylcle
    
    mlTimer1 = mlTimer1 - lCycles \ 2
    mlTimer2 = mlTimer2 - lCycles \ 2
    
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
    
    CopyMemory gyMem(&HFE44&), mlTimer1, 2&
    CopyMemory gyMem(&HFE54&), mlTimer1, 2&

    
    If mlTimer2 <= 0 Then
        mlTimer2 = mlTimer2 + &H10000
        If Not mbTimer2HasInterrupted Then
            mbTimer2HasInterrupted = True
            AssertTIMER2
        End If
    End If
    
    CopyMemory gyMem(&HFE48&), mlTimer2, 2&
    CopyMemory gyMem(&HFE58&), mlTimer2, 2&
End Sub


Public Sub WriteRegister(ByVal lRegister As Long, ByVal lValue As Long)
    Dim lLatchAddress As Long
    Dim lSlowDataBit As Long
    Dim lBitValue As Long
    Dim lIRQMask As Long
    
    ' Debugging.WriteString "SystemVIA6522.WriteRegister"
    
    Select Case lRegister
        Case &H0 ' IORB
            ORB = lValue And DDRB
            mlLatchAddress = ORB And &H7&
            lSlowDataBit = (ORB And &H8&) \ 8&
            
            lBitValue = 2 ^ mlLatchAddress
            LatchValue = LatchValue And (&HFF& Xor lBitValue) Or lBitValue * lSlowDataBit
            
            Select Case mlLatchAddress
                Case 0 ' Sound Generator enable
                    If lSlowDataBit = 1& And Sound.Enabled = 0& Then
                        Sound.WriteByte ORA
                    End If
                    Sound.Enabled = lSlowDataBit
                Case 1
                    ' Speech processor read enable
                Case 2
                    ' Speech processor write enable
                Case 3
                    Keyboard.EnableScan = lSlowDataBit
                Case 4 ' Hardware scroll low
                    VideoULA.HardwareScrollSize = (VideoULA.HardwareScrollSize And -2&) Or lSlowDataBit
                    VideoULA.HardwareScrollBytes = Array(&H4000&, &H2000&, &H5000&, &H2800&)(VideoULA.HardwareScrollSize)
                Case 5 ' Hardware scroll high
                    VideoULA.HardwareScrollSize = (VideoULA.HardwareScrollSize And -3&) Or lSlowDataBit * 2
                    VideoULA.HardwareScrollBytes = Array(&H4000&, &H2000&, &H5000&, &H2800&)(VideoULA.HardwareScrollSize)
                Case 6 ' Caps Lock
                    KeyboardIndicators.SetLED 1, 1 - lSlowDataBit
                Case 7 ' Shift Lock
                    KeyboardIndicators.SetLED 2, 1 - lSlowDataBit
            End Select
            WriteRegister &HD&, &H18& ' Clear CB1, CB2 interrupts
        Case &H1, &HF ' IORA
            ORA = lValue And DDRA
            If Keyboard.EnableScan = 0& Then
                Keyboard.WriteRegister ORA
            End If
            If lRegister = 1 Then
                WriteRegister &HD&, &H3& ' Clear CA1, CA2 interrupts
            End If
        Case &H2 ' DDRB
            DDRB = lValue
            gyMem(&HFE40&) = &H80& ' No speech system
            gyMem(&HFE50&) = &H80&
        Case &H3 ' DDRA
            DDRA = lValue
            gyMem(&HFE43&) = DDRA
            gyMem(&HFE53&) = DDRA
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
            'Stop
        Case &HB ' ACR
            mlTimer1Control = (lValue And &HC0&) \ 64&
            mlTimer2Control = (lValue And &H20&) \ 32&
            mlShiftRegisterControl = (lValue And &H1C&) \ 4&
            mlLatchEnable = lValue And &H3&
            ACR = lValue
            gyMem(&HFE4B&) = lValue
            gyMem(&HFE5B&) = lValue
        Case &HC ' PCR
            PCR = lValue
            gyMem(&HFE4C&) = PCR
            gyMem(&HFE5C&) = PCR
        Case &HD ' IFR
            WriteIFR lValue And &H7F& ' always clear bits - interrupts
        Case &HE ' IER
            ' Set/reset bits
            If lValue And 128& Then
                IER = IER Or (lValue And &H7F&)
            Else
                IER = IER And (lValue Xor &H7F&)
            End If
            gyMem(&HFE4E&) = IER Or &H80&
            gyMem(&HFE5E&) = IER Or &H80&
            
            ' Set IRQ
            lIRQMask = IFR And IER
            
            If (lIRQMask And &H7F&) = 0 Then
                InterruptLine.SetIRQLine irqSystemVIA, 0
            Else
                InterruptLine.SetIRQLine irqSystemVIA, 1
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
    gyMem(&HFE4D&) = IFR
    gyMem(&HFE5D&) = IFR
    
    ' Set IRQ
    lIRQMask = IFR And IER
    
    If (lIRQMask And &H7F&) = 0 Then
        InterruptLine.SetIRQLine irqSystemVIA, 0
    Else
        InterruptLine.SetIRQLine irqSystemVIA, 1
    End If
End Function

Public Function ReadRegister(ByVal lRegister As Long) As Long
    ' Debugging.WriteString "SystemVIA6522.ReadRegister"
    
    Select Case lRegister
        Case &H0&
            ReadRegister = &HB0& And (DDRB Xor &HFF) ' Joystick buttons not pressed (b4-5), no speech system (b7)
            If (PCR And &H20&) = 0& Then
                WriteRegister &HD&, &H18& ' Clear CB1, CB2 interrupts
            End If
            
        Case &H1&, &HF&
            If lRegister = 1 Then
                WriteRegister &HD&, &H2& ' Clear CA1, CA2 interrupts
            End If
        Case &H4&
            WriteRegister &HD&, &H40& ' Reset timer 1 interrrupt
            ReadRegister = gyMem(&HFE44&)
        Case &H8&
            WriteRegister &HD&, &H20& ' Reset timer 2 interrupt
            ReadRegister = gyMem(&HFE48&)
    End Select
End Function
