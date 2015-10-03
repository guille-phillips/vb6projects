Attribute VB_Name = "ACIA6850"
Option Explicit

Private mlDivideRatio As Long
Private mlDivideRatioRate As Long

Private mlDataBits As Long
Private mlParity As Long
Private mlStopBits As Long
Private mlTransmitterControl As Long
Public mlReceiverInterruptEnable As Long


Public mlDataIn As Long
Public mlDataOut As Long

Public mlStatus As Long

Private mlTotalCycles As Long

Private Const TestChars = ":HELLO THIS IS A TEST"
Private Const TestTapeChars = "tape test"

Public mlCyclesPerByteReceive As Long
Public mlCyclesPerByteTransmit As Long

Private mbNoCarrierInterrupt As Boolean

Private mlIndex As Long

Public Sub InitialiseACIA6850()
    ' Debugging.WriteString "ACIA6850.InitialiseACIA6850"
    
    mlCyclesPerByteReceive = 2000000#
    mlCyclesPerByteTransmit = 2000000#

    mlStatus = -(SerialULA.SelectRS432 = 0&) * 4& ' TDREmpty, DCD High = no carrier from cassette
    mlTotalCycles = 0
End Sub

Public Sub Tick(ByVal lCycles As Long)
    ' Debugging.WriteString "ACIA6850.Tick"
    
    mlTotalCycles = mlTotalCycles + lCycles
    If mlTotalCycles >= mlCyclesPerByteReceive Then
        mlTotalCycles = mlTotalCycles - mlCyclesPerByteReceive
        
        If SerialULA.SelectRS432 = 1& Then
            ReadByteFromStream
        Else
'            If TapeHandler.mbRecording Then
'                WriteByteToCassette
'            Else
                ReadByteFromCassette
'            End If
        End If
    End If
End Sub

Public Sub WriteRegister(lRegister As Long, yValue As Byte)
    ' Debugging.WriteString "ACIA6850.WriteRegister"
    
    Select Case lRegister
        Case 0 ' Control
            mlDivideRatio = yValue And &H3&
            If mlDivideRatio = 3& Then
                mlStatus = -(SerialULA.SelectRS432 = 0&) * 4&  ' TDREmpty, DCD High = no carrier from cassette
                InterruptLine.SetIRQLine irqACIA6850, 0
                Exit Sub
            End If
            mlDivideRatioRate = Array(1, 16, 64, 1)(mlDivideRatio)
            
            mlParity = (yValue And &H4&) \ 4& ' e/o
            mlStopBits = (yValue And &H8&) \ 8& ' 2/1
            mlDataBits = (yValue And &H10&) \ 16& ' 7/8
            
            If mlDataBits = &H10& Then
                If mlStopBits = 0& Then
                    mlStopBits = &H8&
                    mlParity = 0&
                End If
            End If
            
            mlTransmitterControl = (yValue And &H60&) \ 32&
            mlReceiverInterruptEnable = (yValue And &H80&) \ 128&
            SetCyclesPerByte
        Case 1 ' Data
            mlDataOut = yValue
            mlStatus = mlStatus And &HFD& ' clear TDEmpty interrupt flag
            InterruptLine.SetIRQLine irqACIA6850, 0
    End Select
End Sub

Public Function ReadRegister(lRegister As Long) As Byte
    ' Debugging.WriteString "ACIA6850.ReadRegister"
    
    Select Case lRegister
        Case 0
            ReadRegister = mlStatus
        Case 1
            ReadRegister = mlDataIn
            mlStatus = mlStatus And &H7E& ' clear RDFull interrupt flag
            InterruptLine.SetIRQLine irqACIA6850, 0
    End Select
End Function

Private Function ReadByteFromStream()
    ' Debugging.WriteString "ACIA6850.ReadByteFromStream"
    
    mlDataIn = Asc(Mid$(TestChars, mlIndex + 1, 1))
    mlIndex = (mlIndex + 1) Mod Len(TestChars)
    
    If mlReceiverInterruptEnable = 1& Then
        mlStatus = mlStatus Or &H81&
        InterruptLine.SetIRQLine irqACIA6850, 1
    End If
End Function

Public Sub SetCyclesPerByte()
    Dim lByteSize As Long
    Dim nBaudReceiveBaud As Single
    Dim nBaudTransmitBaud As Single
    
    ' Debugging.WriteString "ACIA6850.SetCyclesPerByte"
    
    lByteSize = IIf(mlDataBits = 0&, 7&, 8&) + IIf(mlStopBits = 0&, 2&, 1&)
    
    If mlDivideRatioRate <> 0 Then
        nBaudReceiveBaud = SerialULA.ReceiveBaudRate / mlDivideRatioRate / lByteSize
        nBaudTransmitBaud = SerialULA.TransmitBaudRate / mlDivideRatioRate / lByteSize
        
        mlCyclesPerByteReceive = 2000000# / nBaudReceiveBaud
        mlCyclesPerByteTransmit = 2000000# / nBaudTransmitBaud
        mlTotalCycles = 0
        
        mlCyclesPerByteReceive = mlCyclesPerByteReceive / 4
    End If
End Sub


Public Sub LoadBlankTape(ByVal sPath As String)
    ' Debugging.WriteString "ACIA6850.LoadBlankTape"
    
    NoCarrierDetectedInterrrupt
End Sub


Public Sub InitialiseCassette()
    ' Debugging.WriteString "ACIA6850.InitialiseCassette"
    
    NoCarrierDetectedInterrrupt
End Sub

Public Sub EjectCassette()
    ' Debugging.WriteString "ACIA6850.EjectCassette"
    
    NoCarrierDetectedInterrrupt
End Sub

Private Sub NoCarrierDetectedInterrrupt()
    ' Debugging.WriteString "ACIA6850.NoCarrierDetectedInterrrupt"
    
    ACIA6850.mlStatus = ACIA6850.mlStatus Or &H84& ' DCD=1 No Carrier Detected
    If ACIA6850.mlReceiverInterruptEnable = 1& Then
        InterruptLine.SetIRQLine irqACIA6850, 1
    End If
End Sub

Public Sub WriteByteToCassette()
    ' Debugging.WriteString "ACIA6850.WriteByteToCassette"
    
    If StorageMedia.CassetteStorage Is Nothing Or SerialULA.CassetteMotor = 0& Then
        ACIA6850.mlStatus = ACIA6850.mlStatus Or 4& ' DCD=1 No Carrier Detected
        Exit Sub
    End If
    
    If (ACIA6850.mlStatus And &H2&) = 0& Then
'        ReDim Preserve TapeStream(mlTapeFileIndex)
'        TapeStream(mlTapeFileIndex) = ACIA6850.mlDataOut
'        mlTapeFileIndex = mlTapeFileIndex + 1
        ACIA6850.mlStatus = mlStatus Or &H82&
    End If
End Sub

Public Sub ReadByteFromCassette()
    Dim lByte As Long
    
    ' Debugging.WriteString "ACIA6850.ReadByteFromCassette"
    
    If StorageMedia.CassetteStorage Is Nothing Then
        ACIA6850.mlStatus = ACIA6850.mlStatus Or 4& ' DCD=1 No Carrier Detected
        Exit Sub
    End If
    
    lByte = StorageMedia.CassetteStorage.NextCassetteByte
    
    If lByte = 256 Then
        ACIA6850.mlStatus = ACIA6850.mlStatus And &HFB&   ' DCD=0 Carrier Detected
        mbNoCarrierInterrupt = True
    ElseIf lByte = 257 Then
        If ACIA6850.mlReceiverInterruptEnable = 1& And mbNoCarrierInterrupt Then
            ACIA6850.mlStatus = ACIA6850.mlStatus Or &H84& ' DCD=1 No Carrier Detected
            InterruptLine.SetIRQLine irqACIA6850, 1
            mbNoCarrierInterrupt = False
        End If
    Else
        mlDataIn = lByte
        If ACIA6850.mlReceiverInterruptEnable = 1& Then
            ACIA6850.mlStatus = (ACIA6850.mlStatus Or &H81&) And &HFB&  ' Carrier detected, RDRFull
            InterruptLine.SetIRQLine irqACIA6850, 1
        End If
    End If
End Sub




