Attribute VB_Name = "TapeHandler"
Option Explicit

Public TapeStream() As Byte
Private mlTapeFileIndex As Long

Const LeadInLengthShort As Long = 100
Const LeadInLengthLong As Long = 150
Private mlLeadIn As Long
Private mbStartLeadIn As Boolean
Private mbNoCarrierInterrupt As Boolean

Private mlNextBlockPosition As Long
Private mlNextBlockNumber As Long

Private mbTapeLoaded As Boolean
Public mbRecording As Boolean

Private Enum BlockTypes
    btStart = 1
    btEnd = 2
    btMid = 4
End Enum

Private mbiBlocks() As BlockInformation
Private mlTotalBlocks As Long
Private mlBlockIndex As Long
Private Type BlockInformation
    Position As Long
    BlockType As Long
End Type

Public Sub LoadBlankTape(ByVal sPath As String)
    ACIA6850.mlStatus = ACIA6850.mlStatus Or &H84& ' DCD=1 No Carrier Detected
    If ACIA6850.mlReceiverInterruptEnable = 1& Then
        InterruptLine.SetIRQLine irqACIA6850, 1
    End If

    mbTapeLoaded = False
    Erase TapeStream
    mlTapeFileIndex = 0
    mbTapeLoaded = True
End Sub

'    Five seconds of 2400Hz tone.
'    One synchronisation byte (&2A).
'    File name (one to ten characters).
'00 One end of file name marker byte (&00).
'01 Load address of file, four bytes, low byte first.
'05 Execution address of file, four bytes, low byte first.
'09 Block number, two bytes, low byte first.
'11 Data block length, two bytes, low byte first.
'13 Block flag, one byte.
'14 Spare, four bytes, currently &00.
'18 CRC on header, two bytes.
'20 Data, 0 to 256 bytes.
'    CRC on data, two bytes.

Public Sub LoadTape(ByVal sPath As String)
    Dim lIndex As Long
    
    ACIA6850.mlStatus = ACIA6850.mlStatus Or &H84& ' DCD=1 No Carrier Detected
    If ACIA6850.mlReceiverInterruptEnable = 1& Then
        InterruptLine.SetIRQLine irqACIA6850, 1
    End If
    
    mbTapeLoaded = False
    Erase TapeStream
    
    If Not UEFHandler.LoadCompressedUEFFile(sPath) Then
        'If Not UEFHandler.LoadUEFFile(sPath) Then
            MsgBox "Cassette format not recognised", vbExclamation
            Exit Sub
        'End If
    End If
    While UEFHandler.LoadNextBlock
'        Debug.Print HexNum(UEFHandler.BlockIdentifier, 4)
        Select Case UEFHandler.BlockIdentifier
            Case &H100& ' data block
                If UEFHandler.BlockData(0) = &H2A& Then ' sync byte
                    ReDim Preserve TapeStream(mlTapeFileIndex + UEFHandler.BlockLength - 1)
                    CopyMemory TapeStream(mlTapeFileIndex), UEFHandler.BlockData(0), UEFHandler.BlockLength
                    mlTapeFileIndex = mlTapeFileIndex + UEFHandler.BlockLength
                End If
        End Select
    Wend
    mlTapeFileIndex = 0
    FindNextBlock
    mlLeadIn = LeadInLengthLong
    mbStartLeadIn = True
    mbTapeLoaded = True
    SetupBlocks
End Sub

Private Sub SetupBlocks()
    Dim lBlockNumber As Long
    Dim bLastBlock As Long
    Dim lIndex As Long
    Dim biInfo As BlockInformation
    Dim bOk As Boolean
    
    mlTapeFileIndex = 0
    
    bOk = True
    
    Erase mbiBlocks
    
    Do
        lIndex = FindZero(mlTapeFileIndex)
        lBlockNumber = TapeStream(lIndex + 9)
        
        bLastBlock = TapeStream(lIndex + 13) <> 0
        
        biInfo.Position = mlTapeFileIndex
        biInfo.BlockType = IIf(lBlockNumber = 0, btStart, 0)
        biInfo.BlockType = biInfo.BlockType Or IIf(bLastBlock, btEnd, 0)
        biInfo.BlockType = biInfo.BlockType Or IIf(lBlockNumber <> 0 And Not bLastBlock, btMid, 0)
        
'        Debug.Print HexNum(lBlockNumber, 2) & ":" & biInfo.BlockType
        
        ReDim Preserve mbiBlocks(mlBlockIndex)
        mbiBlocks(mlBlockIndex) = biInfo
        mlBlockIndex = mlBlockIndex + 1
        mlTotalBlocks = mlBlockIndex
        
        FindNextBlock
        mlTapeFileIndex = mlNextBlockPosition
        If mlTapeFileIndex = -1 Then
            bOk = False
        End If
    Loop While bOk
    mlTapeFileIndex = 0
    mlBlockIndex = 1
    FindNextBlock
End Sub

Public Sub EjectTape()
    mbTapeLoaded = False
    ACIA6850.mlStatus = ACIA6850.mlStatus Or &H84& ' DCD=1 No Carrier Detected
    If ACIA6850.mlReceiverInterruptEnable = 1& Then
        InterruptLine.SetIRQLine irqACIA6850, 1
    End If
End Sub

Private Sub FindNextBlock()
    Dim lIndex As Long
    Dim lBlockLength As Long
    
    lIndex = FindZero(mlTapeFileIndex)
    If lIndex = -1 Then
        mlNextBlockPosition = -1
        mlNextBlockNumber = -1
    Else
        CopyMemory lBlockLength, TapeStream(lIndex + 11), 2&
        mlNextBlockPosition = lIndex + lBlockLength + 22
        mlNextBlockNumber = TapeStream(FindZero(mlNextBlockPosition) + 9)
        If mlNextBlockPosition > UBound(TapeStream) Then
            mlNextBlockPosition = -1
            mlNextBlockNumber = -1
        End If
    End If
End Sub

Private Function FindZero(ByVal lPosition) As Long
    If lPosition > UBound(TapeStream) Then
        FindZero = -1
        Exit Function
    End If
    While TapeStream(lPosition) <> 0
        lPosition = lPosition + 1
    Wend
    FindZero = lPosition
End Function

Public Sub WriteByteToCassette()
    If Not mbTapeLoaded Or SerialULA.CassetteMotor = 0& Then
        ACIA6850.mlStatus = ACIA6850.mlStatus Or 4& ' DCD=1 No Carrier Detected
        Exit Sub
    End If
    
    If (ACIA6850.mlStatus And &H2&) = 0& Then
        ReDim Preserve TapeStream(mlTapeFileIndex)
        TapeStream(mlTapeFileIndex) = ACIA6850.mlDataOut
        mlTapeFileIndex = mlTapeFileIndex + 1
        ACIA6850.mlStatus = mlStatus Or &H82&
    End If
End Sub

Public Sub ReadByteFromCassette()
    If Not mbTapeLoaded Or SerialULA.CassetteMotor = 0& Then
        ACIA6850.mlStatus = ACIA6850.mlStatus Or 4& ' DCD=1 No Carrier Detected
        Exit Sub
    End If
    
    If mlLeadIn = -1 Then
        mlDataIn = TapeStream(mlTapeFileIndex)
        If ACIA6850.mlReceiverInterruptEnable = 1& Then
            ACIA6850.mlStatus = (ACIA6850.mlStatus Or &H81&) And &HFB&  ' Carrier detected, RDRFull
            InterruptLine.SetIRQLine irqACIA6850, 1
        End If
        mlTapeFileIndex = mlTapeFileIndex + 1
        If mlBlockIndex < mlTotalBlocks Then
            If mlTapeFileIndex = mbiBlocks(mlBlockIndex).Position Then
                If (mbiBlocks(mlBlockIndex).BlockType And 2&) = 2& Then
                    mlLeadIn = LeadInLengthLong
                Else
                    mlLeadIn = LeadInLengthShort
                End If
                mbStartLeadIn = True
                mlBlockIndex = mlBlockIndex + 1
            End If
        Else
            If mlTapeFileIndex > UBound(TapeStream) Then
                mlLeadIn = -2
                mbNoCarrierInterrupt = True
            End If
        End If
    ElseIf mlLeadIn = -2 Then ' end of tape
        If ACIA6850.mlReceiverInterruptEnable = 1& And mbNoCarrierInterrupt Then
            ACIA6850.mlStatus = ACIA6850.mlStatus Or &H84& ' DCD=1 No Carrier Detected
            InterruptLine.SetIRQLine irqACIA6850, 1
            mbNoCarrierInterrupt = False
        End If
    ElseIf mlLeadIn <= (LeadInLengthShort \ 2) Then
        If ACIA6850.mlReceiverInterruptEnable = 1& And mbNoCarrierInterrupt Then
            ACIA6850.mlStatus = ACIA6850.mlStatus Or &H84& ' DCD=1 No Carrier Detected
            InterruptLine.SetIRQLine irqACIA6850, 1
            mbNoCarrierInterrupt = False
        End If
        mlLeadIn = mlLeadIn - 1
    Else
        ACIA6850.mlStatus = ACIA6850.mlStatus And &HFB&   ' DCD=0 Carrier Detected
        mbNoCarrierInterrupt = True
        mlLeadIn = mlLeadIn - 1
    End If
End Sub


Public Sub RewindOneBlock()
    mlBlockIndex = mlBlockIndex - 1
    If mlBlockIndex < 0 Then
        mlBlockIndex = 0
    End If
    mlTapeFileIndex = mbiBlocks(mlBlockIndex).Position
    FindNextBlock
    mlLeadIn = LeadInLengthLong
'    ACIA6850.mlStatus = ACIA6850.mlStatus Or 4&
End Sub

Public Sub RewindStart()
    mlBlockIndex = 0
    mlTapeFileIndex = 0
    mlLeadIn = LeadInLengthLong
'    ACIA6850.mlStatus = ACIA6850.mlStatus Or 4&
End Sub

Public Sub RewindStartBlocks()
    Dim lBlockIndex As Long
    
    For lBlockIndex = mlBlockIndex To 0 Step -1
        If mbiBlocks(lBlockIndex).BlockType And btStart Then
            mlBlockIndex = lBlockIndex
            mlTapeFileIndex = mbiBlocks(lBlockIndex).Position
            Exit For
        End If
    Next
    FindNextBlock
    mlLeadIn = LeadInLengthLong
'    ACIA6850.mlStatus = ACIA6850.mlStatus Or 4&
End Sub

Public Sub ForwardOneBlock()
    mlBlockIndex = mlBlockIndex + 1
    mlTapeFileIndex = mbiBlocks(mlBlockIndex).Position
    FindNextBlock
    mlLeadIn = LeadInLengthLong
'    ACIA6850.mlStatus = ACIA6850.mlStatus Or 4&
End Sub

Public Sub ForwardEndBlocks()
    mlLeadIn = LeadInLengthLong
'    ACIA6850.mlStatus = ACIA6850.mlStatus Or 4&
    FindNextBlock
End Sub

Public Sub ForwardEnd()
    mlLeadIn = LeadInLengthLong
'    ACIA6850.mlStatus = ACIA6850.mlStatus Or 4&
    FindNextBlock
End Sub


'Private Function CRC(lStart As Long, lEnd As Long) As Long
'    Dim lIndex As Long
'    Dim yByte As Byte
'    Dim lBit As Long
'    Dim lT As Long
'    Dim yTemp(1) As Byte
'
'    For lIndex = lStart To lEnd
'        CRC = CRC Xor CLng(mlDummyBlock(lIndex)) * 256
'        For lBit = 1 To 8
'            lT = 0
'            If (CRC And &H8000&) <> 0 Then
'                CRC = CRC Xor &H810&
'                lT = 1
'            End If
'            CRC = (CRC * 2 + lT) And &HFFFF&
'        Next
'    Next
'    CopyMemory yTemp(0), CRC, 2&
'    CRC = CLng(yTemp(0)) * 256 + yTemp(1)
'End Function

