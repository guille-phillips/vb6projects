Attribute VB_Name = "SystemVIA6522"
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

Public mlLatchAddress As Long

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

Public LatchValue As Long

Public Sub InitialiseSystemVIA()
    Dim lAddress As Long
    
    For lAddress = &HFE40& To &HFE4F&
        gyMem(lAddress) = 0
    Next
    
    IFR = 0

    WriteRegister &HE&, &H7F&
    IER = 128
    gyMem(&HFE6E&) = IER
    
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
    If lBit = 1 Then
        IFR = IFR Or &H81
        gyMem(&HFE4D&) = IFR
        If CA2Enable And MASTEREnable Then
            Processor6502.IRQFlag = True
            'Processor6502.IRQRaisedBy = "CA2"
        End If
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
    gyMem(&HFE4D&) = IFR
    If SHIFTREGEnable And MASTEREnable Then
        Processor6502.IRQFlag = True
    End If
End Sub

Public Sub AssertCB2()
    IFR = IFR Or &H88
    gyMem(&HFE4D&) = IFR
    If CB2Enable And MASTEREnable Then
        Processor6502.IRQFlag = True
    End If
End Sub

Public Sub AssertCB1()
    IFR = IFR Or &H90
    gyMem(&HFE4D&) = IFR
    If CB1Enable And MASTEREnable Then
        Processor6502.IRQFlag = True
    End If
End Sub

Public Sub AssertTIMER2()
    IFR = IFR Or &HA0
    gyMem(&HFE4D&) = IFR
    If TIMER2Enable And MASTEREnable Then
        Processor6502.IRQFlag = True
    End If
End Sub

Public Sub AssertTIMER1()
    IFR = IFR Or &HC0
    gyMem(&HFE4D&) = IFR
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
    
    CopyMemory gyMem(&HFE44&), mlTimer1, 2&

    
    If mlTimer2 <= 0 Then
        mlTimer2 = mlTimer1 + &H10000
        If Not mbTimer2HasInterrupted Then
            mbTimer2HasInterrupted = True
            AssertTIMER2
        End If
    End If
    
    CopyMemory gyMem(&HFE48&), mlTimer2, 2&
End Sub


Public Sub WriteRegister(ByVal lRegister As Long, ByVal lValue As Long)
    Dim lLatchAddress As Long
    Dim lSlowDataBit As Long
    Dim lBitValue As Long
    
    Select Case lRegister
        Case &H0 ' IORB
            ORB = lValue And DDRB
            mlLatchAddress = ORB And &H7&
            lSlowDataBit = (ORB And &H8&) \ 8&
            
            lBitValue = 2 ^ mlLatchAddress
            LatchValue = LatchValue And (&HFF& Xor lBitValue) Or lBitValue * lSlowDataBit
            
            Select Case mlLatchAddress
                Case 0
                    ' Sound Generator enable
                    'Debug.Print "0:" & lSlowDataBit
                    Sound.Enabled = lSlowDataBit
                    If lSlowDataBit = 1& Then
                        Sound.WriteByte ORA
                    End If
                Case 1
                    ' Speech processor read enable
                Case 2
                    ' Speech processor write enable
                Case 3
                    Keyboard.EnableScan = lSlowDataBit
                Case 4 ' Hardware scroll low
                    VideoULA.HardwareScrollSize = (VideoULA.HardwareScrollSize And -2&) Or lSlowDataBit
                    VideoULA.HardwareScrollBytes = Array(&H4000&, &H2000&, &H5000&, &H2800&)(VideoULA.HardwareScrollSize)
                Case 5 ' Hardware scroll hi
                    VideoULA.HardwareScrollSize = (VideoULA.HardwareScrollSize And -3&) Or lSlowDataBit * 2
                    VideoULA.HardwareScrollBytes = Array(&H4000&, &H2000&, &H5000&, &H2800&)(VideoULA.HardwareScrollSize)
                Case 6 ' Caps Lock
                    KeyboardIndicators.SetLED 1, 1 - lSlowDataBit
                Case 7 ' Shift Lock
                    KeyboardIndicators.SetLED 2, 1 - lSlowDataBit
            End Select
        Case &H1, &HF ' IORA
            ORA = lValue And DDRA
            If Keyboard.EnableScan = 0& Then
                Keyboard.WriteRegister ORA
            End If
            If Sound.Enabled = 1& Then
                'Sound.WriteByte ORA
            End If
        Case &H2 ' DDRB
            DDRB = lValue
            gyMem(&HFE40&) = &H80& ' No speech system
        Case &H3 ' DDRA
            DDRA = lValue
        Case &H4, &H6 ' T1CL / T1CH
            mlTimer1Latch = (mlTimer1Latch And &HFF00&) + lValue
        Case &H5 ' T1CH
            mlTimer1Latch = (mlTimer1Latch And &HFF&) + lValue * 256&
            mlTimer1 = mlTimer1Latch
            gyMem(&HFE4D&) = gyMem(&HFE4D&) And &HBF& ' CLEAR IFR FOR TIMER1
            mbTimer1HasInterrupted = False ' RESET ONE SHOT
        Case &H7 ' T1LH
            mlTimer1Latch = (mlTimer1Latch And &HFF&) + lValue * 256&
        Case &H8 ' T2CL
            mlTimer2Latch = lValue
        Case &H9 ' T2CH
            mlTimer2 = mlTimer2Latch + lValue * 256&
            mbTimer2HasInterrupted = False
            IFR = IFR And &HDF&
            gyMem(&HFE4D&) = IFR ' CLEAR IFR FOR TIMER2
        Case &HA ' SR
            'Stop
        Case &HB ' ACR
            mlTimer1Control = (lValue And &HC0&) \ 64&
            mlTimer2Control = (lValue And &H20&) \ 32&
            mlShiftRegisterControl = (lValue And &H1C&) \ 4&
            mlLatchEnable = lValue And &H3&
            ACR = lValue
            gyMem(&HFE4B&) = lValue
        Case &HC ' PCR
            PCR = lValue
            gyMem(&HFE4C&) = PCR
        Case &HD ' IFR
            If lValue And 128& Then
                IFR = IFR Or (lValue And &H7F&)
            Else
                IFR = IFR And (lValue Xor &H7F&)
            End If
            gyMem(&HFE4D&) = IFR
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
            gyMem(&HFE4E&) = IER Or 128&
    End Select
End Sub

Public Function ReadRegister(ByVal lRegister As Long) As Long
    Select Case lRegister
        Case &H0&
            ReadRegister = &HB0& And (DDRB Xor &HFF) ' Joystick buttons not pressed (b4-5), no speech system (b7)
        Case &H4&
            WriteRegister &HD&, &H40& ' Reset timer 1 interrrupt
            ReadRegister = gyMem(&HFE44&)
        Case &H8&
            WriteRegister &HD&, &H20& ' Reset timer 2 interrupt
            ReadRegister = gyMem(&HFE48&)
    End Select
End Function
