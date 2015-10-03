Attribute VB_Name = "Snapshot"
Option Explicit

Private mlBaseAddress As Long

Public Sub LoadSnapshot(ByVal sPath As String)
    ' Debugging.WriteString "Snapshot.LoadSnapshot"
    
    If UEFHandler.LoadUEFFile(sPath) Then
        LoadRegisters
        LoadRomSelect
        LoadMemory
        LoadSidewaysRam
        LoadSystemVia
        LoadUserVia
        LoadVideo
        LoadSound
        LoadMemorySegments
    Else
        MsgBox "Snapshot format not recognised."
    End If
End Sub


Private Sub LoadMemory()
    Dim lAddress As Long
    
    ' Debugging.WriteString "Snapshot.LoadMemory"
    
    UEFHandler.ResetUEF
    If UEFHandler.FindBlock(&H462&) Then
        CopyMemory gyMem(0), UEFHandler.BlockData(0), UEFHandler.BlockLength
    End If
End Sub

Private Sub LoadRomSelect()
    ' Debugging.WriteString "Snapshot.LoadRomSelect"
    
    UEFHandler.ResetUEF
    If UEFHandler.FindBlock(&H461&) Then
        RomSelect.SetRom UEFHandler.BlockData(0)
    End If
End Sub

Private Sub LoadVideo()
    Dim yReg(0) As Byte
    Dim lRegister As Long
    Dim lPalletteIndex As Long
    Dim vPalletteOrder As Variant
    
    ' Debugging.WriteString "Snapshot.LoadVideo"
    
    UEFHandler.ResetUEF
    If UEFHandler.FindBlock(&H468&) Then
        For lRegister = 0 To 17
            Memory.Mem(&HFE00&) = lRegister
            Memory.Mem(&HFE01&) = UEFHandler.BlockData(lRegister)
        Next
        
        Memory.Mem(&HFE20&) = UEFHandler.BlockData(18)
            
        For lPalletteIndex = 0 To 15
            Memory.Mem(&HFE21&) = UEFHandler.BlockData(19 + lPalletteIndex) + lPalletteIndex * 16
        Next
    End If
    VideoULA.mlCurrentRow = 1000 ' force vertical sync
End Sub

Private Sub LoadRegisters()
    ' Debugging.WriteString "Snapshot.LoadRegisters"
    
    UEFHandler.ResetUEF
    If UEFHandler.FindBlock(&H460&) Then
        CopyMemory Processor6502Debug.PC, UEFHandler.BlockData(0), 2&

        Processor6502.A = BlockData(2&)
        Processor6502.X = BlockData(3&)
        Processor6502.Y = BlockData(4&)
        Processor6502.S = BlockData(5&) + &H100&
        
        Processor6502.N = BlockData(6&) And 128&
        Processor6502.V = BlockData(6&) And 64&
        Processor6502.B = BlockData(6&) And 16&
        Processor6502.D = BlockData(6&) And 8&
        Processor6502.I = BlockData(6&) And 4&
        Processor6502.Z = BlockData(6&) And 2&
        Processor6502.C = BlockData(6&) And 1&
    End If
End Sub

Private Sub LoadSystemVia()
    Dim lT1 As Long
    Dim lT1Latch As Long
    Dim lT2 As Long
    Dim lT2Latch As Long
    Dim lBitIndex As Long
    Dim bBlockFound As Boolean
    
    ' Debugging.WriteString "Snapshot.LoadSystemVia"
    
    UEFHandler.ResetUEF
    Do
        If UEFHandler.FindBlock(&H467&) Then
            If BlockData(0) = 0 Then
                bBlockFound = True
                Exit Do
            End If
        Else
            Exit Do
        End If
    Loop
    
    If bBlockFound Then
        Memory.Mem(&HFE40&) = UEFHandler.BlockData(1&) ' ORB
        gyMem(&HFE40&) = UEFHandler.BlockData(2&) ' IRB
        Memory.Mem(&HFE41&) = UEFHandler.BlockData(3&) ' ORA
        gyMem(&HFE41&) = UEFHandler.BlockData(4&) ' IRB
        
        Memory.Mem(&HFE42&) = UEFHandler.BlockData(5&) ' DDRB
        Memory.Mem(&HFE43&) = UEFHandler.BlockData(6&) ' DDRA
        
        lT1 = (UEFHandler.BlockData(8&) * 256& + UEFHandler.BlockData(7&))
        lT1Latch = (UEFHandler.BlockData(10&) * 256& + UEFHandler.BlockData(9&))
        lT2 = (UEFHandler.BlockData(12&) * 256& + UEFHandler.BlockData(11&))
        lT2Latch = (UEFHandler.BlockData(14&) * 256& + UEFHandler.BlockData(13&))
        
        SystemVIA6522.mlTimer1 = lT1
        SystemVIA6522.mlTimer1Latch = lT1Latch
        SystemVIA6522.mlTimer2 = lT2
        SystemVIA6522.mlTimer2Latch = lT2Latch
        
        CopyMemory gyMem(&HFE44&), lT1, 2&
        CopyMemory gyMem(&HFE46&), lT1Latch, 2&
        CopyMemory gyMem(&HFE48&), lT2, 2&
        
        Memory.Mem(&HFE4B&) = UEFHandler.BlockData(15&) ' ACR
        Memory.Mem(&HFE4C&) = UEFHandler.BlockData(16&) ' PCR
        
        Memory.Mem(&HFE4D&) = &H7F&
        Memory.Mem(&HFE4D&) = UEFHandler.BlockData(17&) Or &H80& 'IFR
        Memory.Mem(&HFE4E&) = &H7F&
        Memory.Mem(&HFE4E&) = UEFHandler.BlockData(18&) Or &H80& ' IER
        
        SystemVIA6522.mbTimer1HasInterrupted = UEFHandler.BlockData(19&) * True
        SystemVIA6522.mbTimer2HasInterrupted = UEFHandler.BlockData(20&) * True
        
        Dim lMask As Long
        
        lMask = 1
        For lBitIndex = 0 To 7
            Memory.Mem(&HFE40&) = -((UEFHandler.BlockData(21&) And lMask) <> 0) * 8 + lBitIndex
            lMask = lMask * 2
        Next
    End If
End Sub

Private Sub LoadUserVia()
    Dim lT1 As Long
    Dim lT1Latch As Long
    Dim lT2 As Long
    Dim lT2Latch As Long
    Dim bBlockFound As Boolean
    
    ' Debugging.WriteString "Snapshot.LoadUserVia"
    
    UEFHandler.ResetUEF
    Do
        If UEFHandler.FindBlock(&H467&) Then
            If BlockData(0) = 1 Then
                bBlockFound = True
                Exit Do
            End If
        Else
            Exit Do
        End If
    Loop
    
    If bBlockFound Then
        Memory.Mem(&HFE60&) = BlockData(1&) ' ORB
        gyMem(&HFE60&) = BlockData(2&) ' IRB
        Memory.Mem(&HFE61&) = BlockData(3&) ' ORA
        gyMem(&HFE61&) = BlockData(4&) ' IRB
        
        Memory.Mem(&HFE62&) = BlockData(5&) ' DDRB
        Memory.Mem(&HFE63&) = BlockData(6&) ' DDRA
        
        lT1 = (BlockData(8&) * 256& + BlockData(7&))
        lT1Latch = (BlockData(10&) * 256& + BlockData(9&))
        lT2 = (BlockData(12&) * 256& + BlockData(11&))
        lT2Latch = (BlockData(14&) * 256& + BlockData(13&))
        
        UserVIA6522.mlTimer1 = lT1
        UserVIA6522.mlTimer1Latch = lT1Latch
        UserVIA6522.mlTimer2 = lT2
        UserVIA6522.mlTimer2Latch = lT2Latch
        
        CopyMemory gyMem(&HFE64&), lT1, 2&
        CopyMemory gyMem(&HFE66&), lT1Latch, 2&
        CopyMemory gyMem(&HFE68&), lT2, 2&
        
        Memory.Mem(&HFE6B&) = BlockData(15&) ' ACR
        Memory.Mem(&HFE6C&) = BlockData(16&) ' PCR
        
        Memory.Mem(&HFE6D&) = &H7F&
        Memory.Mem(&HFE6D&) = BlockData(17&) Or &H80& 'IFR
        Memory.Mem(&HFE6E&) = &H7F&
        Memory.Mem(&HFE6E&) = BlockData(18&) Or &H80& ' IER
        
        UserVIA6522.mbTimer1HasInterrupted = BlockData(19&) * True
        UserVIA6522.mbTimer2HasInterrupted = BlockData(20&) * True
    End If
End Sub

Private Sub LoadMemorySegments()
    Dim lAddress As Long
    
    ' Debugging.WriteString "Snapshot.LoadMemorySegments"
    
    UEFHandler.ResetUEF
    While UEFHandler.FindBlock(&HFF00&)
        CopyMemory lAddress, UEFHandler.BlockData(0), 2&
        Debug.Print HexNum(lAddress, 4)
        CopyMemory gyMem(lAddress), UEFHandler.BlockData(2), UEFHandler.BlockLength - 2
    Wend
End Sub

Public Sub SaveSnapshot(ByVal sPath As String)
    ' Debugging.WriteString "Snapshot.SaveSnapshot"
    
    UEFHandler.CreateUEFFile
    SaveRegisters
    SaveRomSelect
    SaveMemory
    SaveSidewaysRam
    SaveSystemVia
    SaveUserVia
    SaveVideo
    SaveSound
    UEFHandler.SaveUEFFile sPath
End Sub

Private Sub SaveRegisters()
    ' Debugging.WriteString "Snapshot.SaveRegisters"
    
    UEFHandler.ResetBlock 7&

    CopyMemory UEFHandler.BlockData(0), Processor6502.PC, 2&
    
    BlockData(2&) = Processor6502.A
    BlockData(3&) = Processor6502.X
    BlockData(4&) = Processor6502.Y
    BlockData(5&) = Processor6502.S And &HFF
    
    BlockData(6&) = Processor6502.N + Processor6502.V + Processor6502.B + Processor6502.D + Processor6502.I + Processor6502.Z + Processor6502.C

    UEFHandler.SaveBlock &H460&
End Sub

Private Sub SaveMemory()
    ' Debugging.WriteString "Snapshot.SaveMemory"
    
    UEFHandler.ResetBlock 32768
    CopyMemory UEFHandler.BlockData(0), gyMem(0), 32768
    UEFHandler.SaveBlock &H462&
End Sub

Private Sub SaveSidewaysRam()
    ' Debugging.WriteString "Snapshot.SaveSidewaysRam"
    Dim lRomSlot As Long
    
    For lRomSlot = 0 To 15
        If RomBankWriteable(lRomSlot) Then
            UEFHandler.ResetBlock 16384 + 1
            UEFHandler.BlockData(0) = lRomSlot
            CopyMemory UEFHandler.BlockData(1), RomBank(0, lRomSlot), 16384
            UEFHandler.SaveBlock &H362&
        End If
    Next
End Sub

Private Sub LoadSidewaysRam()
    Dim lAddress As Long
    Dim lRomSlot As Long
    
    ' Debugging.WriteString "Snapshot.LoadSidewaysRam"
    
    UEFHandler.ResetUEF
    If UEFHandler.FindBlock(&H362&) Then
        lRomSlot = UEFHandler.BlockData(0)
        CopyMemory RomBank(0, lRomSlot), UEFHandler.BlockData(1), UEFHandler.BlockLength - 1
    End If
End Sub

Private Sub SaveRomSelect()
    ' Debugging.WriteString "Snapshot.SaveRomSelect"
    
    UEFHandler.ResetBlock 2
    UEFHandler.BlockData(0) = Memory.Mem(&HFE30&)
    UEFHandler.SaveBlock &H461&
End Sub

Private Sub SaveVideo()
    ' Debugging.WriteString "Snapshot.SaveVideo"
    
    UEFHandler.ResetBlock 35
    
    CopyMemory BlockData(0), CRTC6845.Register(0), 18&
    CopyMemory BlockData(18&), VideoULA.Register(0), 17&
    
    UEFHandler.SaveBlock &H468&
End Sub

Private Sub SaveSystemVia()
    ' Debugging.WriteString "Snapshot.SaveSystemVia"
    
    UEFHandler.ResetBlock 22
    
    BlockData(0&) = 0 ' SystemVIA
    
    BlockData(1&) = SystemVIA6522.ORB
    BlockData(2&) = gyMem(&HFE40&)
    BlockData(3&) = SystemVIA6522.ORA
    BlockData(4&) = gyMem(&HFE41&)
    
    
    BlockData(5&) = SystemVIA6522.DDRB
    BlockData(6&) = SystemVIA6522.DDRA
    
    CopyMemory BlockData(7&), SystemVIA6522.mlTimer1, 2&
    CopyMemory BlockData(9&), SystemVIA6522.mlTimer1Latch, 2&
    CopyMemory BlockData(11&), SystemVIA6522.mlTimer2, 2&
    CopyMemory BlockData(13&), SystemVIA6522.mlTimer2Latch, 2&
    
    BlockData(15&) = SystemVIA6522.ACR
    BlockData(16&) = SystemVIA6522.PCR
    
    BlockData(17&) = SystemVIA6522.IFR
    BlockData(18&) = SystemVIA6522.IER
    BlockData(19&) = SystemVIA6522.mbTimer1HasInterrupted And 1&
    BlockData(20&) = SystemVIA6522.mbTimer2HasInterrupted And 1&
    
    BlockData(21&) = SystemVIA6522.LatchValue ' IC32 Latch
    
    UEFHandler.SaveBlock &H467&
End Sub

Private Sub SaveUserVia()
    ' Debugging.WriteString "Snapshot.SaveUserVia"
    
    UEFHandler.ResetBlock 22
    
    BlockData(0) = 1 ' SystemVIA
    
    BlockData(1) = UserVIA6522.ORB
    BlockData(2) = gyMem(&HFE60&)
    BlockData(3) = UserVIA6522.ORA
    BlockData(4) = gyMem(&HFE61&)
    
    
    BlockData(5) = UserVIA6522.DDRB
    BlockData(6) = UserVIA6522.DDRA
    
    CopyMemory BlockData(7), UserVIA6522.mlTimer1, 2&
    CopyMemory BlockData(9), UserVIA6522.mlTimer1Latch, 2&
    CopyMemory BlockData(11), UserVIA6522.mlTimer2, 2&
    CopyMemory BlockData(13), UserVIA6522.mlTimer2Latch, 2&
    
    BlockData(15) = UserVIA6522.ACR
    BlockData(16) = UserVIA6522.PCR
    
    BlockData(17) = UserVIA6522.IFR
    BlockData(18) = UserVIA6522.IER
    BlockData(19) = UserVIA6522.mbTimer1HasInterrupted And 1&
    BlockData(20) = UserVIA6522.mbTimer2HasInterrupted And 1&
    
    BlockData(21) = 0 ' IC32 Latch
    
    UEFHandler.SaveBlock &H467&
End Sub

Private Sub SaveSound()
    ' Debugging.WriteString "Snapshot.SaveSound"
    
    UEFHandler.ResetBlock 20
    CopyMemory BlockData(0), Sound.chChannels(3).Frequency, 2&
    CopyMemory BlockData(2), Sound.chChannels(2).Frequency, 2&
    CopyMemory BlockData(4), Sound.chChannels(1).Frequency, 2&
    BlockData(6) = Sound.chChannels(3).Volume
    BlockData(7) = Sound.chChannels(2).Volume
    BlockData(8) = Sound.chChannels(1).Volume
    BlockData(9) = 0 ' Sound.chChannels(0).Frequency ' not quite right
    BlockData(10) = Sound.chChannels(1).Volume
    BlockData(11) = 0 ' Most recent tone frequency lo byte
    CopyMemory BlockData(12), 0, 2& ' Speech chip
    CopyMemory BlockData(14), 0, 2& ' Speech chip
    CopyMemory BlockData(16), 0, 2& ' Speech chip
    CopyMemory BlockData(18), 0, 2& ' Speech chip
    UEFHandler.SaveBlock &H46B&
End Sub

Private Sub LoadSound()
    ' Debugging.WriteString "Snapshot.LoadSound"
    
    UEFHandler.ResetUEF

    If Not UEFHandler.FindBlock(&H46B&) Then
        Exit Sub
    End If
    
    Sound.WriteByte (BlockData(0) And &HF&) + 0 * &H20& + &H80&
    Sound.WriteByte (BlockData(0) And &H3F0&) \ 16&
    Sound.WriteByte (BlockData(1) And &HF&) + 1 * &H20& + &H80&
    Sound.WriteByte (BlockData(1) And &H3F0&) \ 16&
    Sound.WriteByte (BlockData(2) And &HF&) + 2 * &H20& + &H80&
    Sound.WriteByte (BlockData(2) And &H3F0&) \ 16&
    
    Sound.WriteByte BlockData(6) + 0 * &H20& + &H80& + &H10&
    Sound.WriteByte BlockData(7) + 1 * &H20& + &H80& + &H10&
    Sound.WriteByte BlockData(8) + 2 * &H20& + &H80& + &H10&
    Sound.WriteByte BlockData(10) + 3 * &H20& + &H80& + &H10&
End Sub


Public Sub StartTransfer(ByVal sFile As String)
    Dim dTime As Date
    
    mlBaseAddress = &H7700&
    
    InitialiseComPort
    TransferBasicProgram
    
    dTime = Time
    While Time < TimeSerial(Hour(dTime), Minute(dTime), Second(dTime) + 2)
        DoEvents
    Wend
    InitialiseFastComPort
    TransferSnapshot sFile
    Console.Com.PortOpen = False
End Sub

Private Sub InitialiseComPort()
    Console.Com.CommPort = 1
    Console.Com.Settings = "9600,N,8,1"
    Console.Com.InputLen = 0
    Console.Com.OutBufferSize = 32767
    Console.Com.PortOpen = True
End Sub


Private Sub InitialiseFastComPort()
    Console.Com.CommPort = 1
    Console.Com.Settings = "38400,N,8,1"
    Console.Com.InputLen = 0
    Console.Com.OutBufferSize = 32767
    Console.Com.PortOpen = True
End Sub

Private Sub SendWord(ByVal lWord As Long)
    Dim yOut(1) As Byte
    
    yOut(0) = lWord And &HFF
    yOut(1) = (lWord And &HFF00&) \ 256

    Console.Com.Output = yOut
End Sub

Private Sub SendBlockDetails(ByVal lAddress As Long, ByVal lLength As Long)
    lLength = 65536 - lLength
    SendWord lAddress - (lLength And &HFF&)
    SendWord lLength
End Sub

Private Sub SendByte(ByVal lAddress, ByVal yValue As Long)
    Dim yByte(0) As Byte
    
    SendBlockDetails lAddress, 1
    yByte(0) = yValue
    Console.Com.Output = yByte
End Sub

Private Sub SendData(ByVal lAddress As Long, ByVal lLength As Long, yData() As Byte)
    SendBlockDetails lAddress, lLength
    Console.Com.Output = yData
End Sub

Private Sub TransferBasicProgram()
    Dim oFSO As New FileSystemObject
    Dim sFile As String
    Dim lLineNumber As Long
    Dim vSplit As Variant
    Dim vLine As Variant
    Dim dTime As Date
    Dim lSlow As Long
    
    sFile = oFSO.OpenTextFile(App.path & "\ReadSerialBBC.txt").ReadAll
    vSplit = Split(sFile, vbCrLf)
    lLineNumber = 10
    For Each vLine In vSplit
        If vLine <> "" Then
            vLine = Replace$(vLine, "&7F00", "&" & Hex$(mlBaseAddress))
            Console.Com.Output = lLineNumber & vLine & vbCr
            lLineNumber = lLineNumber + 10
        Else
            Console.Com.Output = vbCr
        End If
        For lSlow = 0 To 10000
            DoEvents
        Next
    Next
    Console.Com.Output = "RUN" & vbCr
    'Console.Com.Output = "*FX2" & vbCr
    Console.Com.PortOpen = False
End Sub

Private Sub TransferSnapshot(ByVal sPath As String)

    UEFHandler.LoadUEFFile sPath
    
    Console.Com.Output = "*"
    
    TransferMemory
    TransferRomSelect
    TransferVideo
    TransferRegisters
    TransferSound
    TransferUserVIA
    TransferSystemVIA

    SendWord 0 ' Dummy start address
    SendWord 0 ' Dummy end address : Load registers and jump
End Sub

Private Sub TransferMemory()
    Dim lIndex As Long
    Dim sMemString As String
    Dim yMem() As Byte
    Dim lTotal As Long
    Dim lBlock As Long
    Dim lBlockSize As Long
    Dim lBlockEnd As Long
    Dim lEndAddress As Long
    
    Dim lSlow As Long
    
    UEFHandler.ResetUEF
    If Not UEFHandler.FindBlock(&H462&) Then
        Exit Sub
    End If
    
    lBlockSize = 2000
    ReDim yMem(lBlockSize - 1)
    
    lEndAddress = mlBaseAddress - 1
    For lBlock = &H0& To lEndAddress Step lBlockSize
        If (lBlock + lBlockSize - 1) <= lEndAddress Then
            For lIndex = 0 To lBlockSize - 1
                yMem(lIndex) = UEFHandler.BlockData(lBlock + lIndex)
            Next
            SendData lBlock, lBlockSize, yMem
        Else
            ReDim yMem(lEndAddress - lBlock)
            For lIndex = 0 To lEndAddress - lBlock
                yMem(lIndex) = UEFHandler.BlockData(lBlock + lIndex)
            Next
            SendData lBlock, lEndAddress - lBlock + 1, yMem
        End If
    Next
    
    lBlockSize = 2000
    ReDim yMem(lBlockSize - 1)
    
    lEndAddress = &H7FFF&
    For lBlock = mlBaseAddress + 200 To lEndAddress Step lBlockSize
        If (lBlock + lBlockSize - 1) <= lEndAddress Then
            For lIndex = 0 To lBlockSize - 1
                yMem(lIndex) = UEFHandler.BlockData(lBlock + lIndex)
            Next
            SendData lBlock, lBlockSize, yMem
        Else
            ReDim yMem(lEndAddress - lBlock)
            For lIndex = 0 To lEndAddress - lBlock
                yMem(lIndex) = UEFHandler.BlockData(lBlock + lIndex)
            Next
            SendData lBlock, lEndAddress - lBlock + 1, yMem
        End If
    Next
    
End Sub

Private Sub TransferRomSelect()
    UEFHandler.ResetUEF
    If UEFHandler.FindBlock(&H461&) Then
        SendByte &HFE30&, UEFHandler.BlockData(0)
    End If
End Sub

Private Sub TransferVideo()
    Dim yReg(0) As Byte
    Dim lRegister As Long
    Dim lPalletteIndex As Long
    Dim vPalletteOrder As Variant
    
    UEFHandler.ResetUEF
    If Not UEFHandler.FindBlock(&H468&) Then
        Exit Sub
    End If
    
    For lRegister = 0 To 17
        SendByte &HFE00&, lRegister
        SendByte &HFE01&, UEFHandler.BlockData(lRegister)
    Next
    
    SendByte &HFE20&, UEFHandler.BlockData(18)
        
    For lPalletteIndex = 0 To 15
        SendByte &HFE21&, UEFHandler.BlockData(19 + lPalletteIndex) + lPalletteIndex * 16
    Next
End Sub

Private Sub TransferRegisters()
    Dim lC As Long
    Dim lZ As Long
    Dim lInt As Long
    Dim lD As Long
    Dim lB As Long
    Dim lV As Long
    Dim lN As Long
    Dim lP As Long
    
    UEFHandler.ResetUEF
    If Not UEFHandler.FindBlock(&H460&) Then
        Exit Sub
    End If
    
    SendByte mlBaseAddress + 1& + 30&, UEFHandler.BlockData(5) ' S Reg
    SendByte mlBaseAddress + 4& + 30&, UEFHandler.BlockData(6) ' P
    SendByte mlBaseAddress + 7& + 30&, UEFHandler.BlockData(3) ' X Reg
    SendByte mlBaseAddress + 9& + 30&, UEFHandler.BlockData(4) ' Y Reg
    SendByte mlBaseAddress + 11& + 30&, UEFHandler.BlockData(2) ' A Reg
    SendByte mlBaseAddress + 14& + 30&, UEFHandler.BlockData(0) ' PC Reg Lo
    SendByte mlBaseAddress + 15& + 30&, UEFHandler.BlockData(1) ' PC Reg Hi
 
'    txtA.Text = Hex2(UEFHandler.BlockData(2))
'    txtX.Text = Hex2(UEFHandler.BlockData(3))
'    txtY.Text = Hex2(UEFHandler.BlockData(4))
'    txtS.Text = Hex2(UEFHandler.BlockData(5))
'    txtP.Text = Hex2(UEFHandler.BlockData(6))
'    txtPC.Text = Hex4(UEFHandler.BlockData(0) + UEFHandler.BlockData(1) * 256&)
End Sub

Private Sub TransferSystemVIA()
    Dim lByte As Long
    Dim bBlockFound As Boolean
    
    UEFHandler.ResetUEF
    Do
        If UEFHandler.FindBlock(&H467&) Then
            If BlockData(0) = 0 Then
                bBlockFound = True
                Exit Do
            End If
        Else
            Exit Do
        End If
    Loop
    
    If Not bBlockFound Then
        Exit Sub
    End If

    ' IC 32 state
    Dim lValue As Long
    
    SendByte &HFE42&, &HFF&
    lValue = UEFHandler.BlockData(21)
    For lByte = 0 To 7
        SendByte &HFE40&, lByte + Sgn(lValue And 2 ^ lByte) * 8
    Next

    SendByte &HFE4B&, UEFHandler.BlockData(15) ' ACR
    SendByte &HFE4C&, UEFHandler.BlockData(16) ' PCR

    SendByte &HFE42&, &HFF& ' Output B
    SendByte &HFE40&, UEFHandler.BlockData(1) ' ORB

    SendByte &HFE43&, &HFF& ' Output A
    SendByte &HFE4F&, UEFHandler.BlockData(3) ' ORA

    SendByte &HFE42&, UEFHandler.BlockData(5) ' DDRB
    SendByte &HFE43&, UEFHandler.BlockData(6) ' DDRA

    SendByte &HFE44&, &HFF ' T1-L
    SendByte &HFE45&, &HFF ' T1-H

    SendByte mlBaseAddress + 1, UEFHandler.BlockData(7)
    SendByte mlBaseAddress + 3, UEFHandler.BlockData(8)
    SendByte mlBaseAddress + 11, UEFHandler.BlockData(9)
    SendByte mlBaseAddress + 13, UEFHandler.BlockData(10)
    SendByte mlBaseAddress + 21, UEFHandler.BlockData(13)
    SendByte mlBaseAddress + 23, UEFHandler.BlockData(14)
    
'    SendByte &HFE46&, UEFHandler.BlockData(9)
'    SendByte &HFE47&, UEFHandler.BlockData(10)
    
    SendByte &HFE48&, UEFHandler.BlockData(13)
    SendByte &HFE49&, UEFHandler.BlockData(14)
    
'    SendByte &HFE48&, &HFF ' T2-L
'    SendByte &HFE49&, &HFF ' T2-H

'    SendByte mlBaseAddress + 11, UEFHandler.BlockData(13)
'    SendByte mlBaseAddress + 16, UEFHandler.BlockData(14)

    SendByte &HFE4D&, &H7F ' IFR Clear all bits
    SendByte &HFE4D&, UEFHandler.BlockData(17) Or &H80 ' IFR

    SendByte &HFE4E&, &H7F ' IER Clear all bits
    SendByte &HFE4E&, UEFHandler.BlockData(18) Or &H80 ' IER set bits
End Sub

Private Sub TransferUserVIA()
    Dim bBlockFound As Boolean
    
    UEFHandler.ResetUEF
    Do
        If UEFHandler.FindBlock(&H467&) Then
            If BlockData(0) = 1 Then
                bBlockFound = True
                Exit Do
            End If
        Else
            Exit Do
        End If
    Loop
    
    If Not bBlockFound Then
        Exit Sub
    End If

    SendByte &HFE6B&, UEFHandler.BlockData(15)  ' ACR
    SendByte &HFE6C&, UEFHandler.BlockData(16) ' PCR
    
    SendByte &HFE62&, &HFF& ' DDRB
    SendByte &HFE60&, UEFHandler.BlockData(1) ' ORB
    
    SendByte &HFE63&, &HFF&  ' DDRA
    SendByte &HFE6F&, UEFHandler.BlockData(3)  ' ORA
    
    SendByte &HFE62&, UEFHandler.BlockData(5) ' DDRB
    SendByte &HFE63&, UEFHandler.BlockData(6)  ' DDRA
    
    SendByte &HFE64&, UEFHandler.BlockData(7) ' T1-L
    SendByte &HFE65&, UEFHandler.BlockData(8)  ' T1-H
    
    SendByte &HFE68&, UEFHandler.BlockData(11) ' T2-L
    SendByte &HFE69&, UEFHandler.BlockData(12) ' T2-H
    
    SendByte &HFE6E&, &H7F ' IER Clear all bits
    SendByte &HFE6E&, UEFHandler.BlockData(18) Or &H80  ' IER set bits
    
    SendByte &HFE6D&, &H7F ' IFR
    SendByte &HFE6D&, UEFHandler.BlockData(17) Or &H80  ' IFR
End Sub

Private Sub TransferSound()
    ' Debugging.WriteString "Snapshot.TransferSound"
    
    UEFHandler.ResetUEF

    If Not UEFHandler.FindBlock(&H46B&) Then
        Exit Sub
    End If
    
    SendByte &HFE43&, &HFF& ' DDRA
    SendByte &HFE42&, &HF&  ' DDRB
    
    SendByte &HFE4F&, (BlockData(0) And &HF&) + 0 * &H20& + &H80&: SendByte &HFE40&, &H0&: SendByte &HFE40&, &H8&
    SendByte &HFE4F&, (BlockData(0) And &H3F0&) \ 16&: SendByte &HFE40&, &H0&: SendByte &HFE40&, &H8&
    SendByte &HFE4F&, (BlockData(1) And &HF&) + 1 * &H20& + &H80&: SendByte &HFE40&, &H0&: SendByte &HFE40&, &H8&
    SendByte &HFE4F&, (BlockData(1) And &H3F0&) \ 16&: SendByte &HFE40&, &H0&: SendByte &HFE40&, &H8&
    SendByte &HFE4F&, (BlockData(2) And &HF&) + 2 * &H20& + &H80&: SendByte &HFE40&, &H0&: SendByte &HFE40&, &H8&
    SendByte &HFE4F&, (BlockData(2) And &H3F0&) \ 16&: SendByte &HFE40&, &H0&: SendByte &HFE40&, &H8&
    
    SendByte &HFE4F&, BlockData(6) + 0 * &H20& + &H80& + &H10&: SendByte &HFE40&, &H0&: SendByte &HFE40&, &H8&
    SendByte &HFE4F&, BlockData(7) + 1 * &H20& + &H80& + &H10&: SendByte &HFE40&, &H0&: SendByte &HFE40&, &H8&
    SendByte &HFE4F&, BlockData(8) + 2 * &H20& + &H80& + &H10&: SendByte &HFE40&, &H0&: SendByte &HFE40&, &H8&
    SendByte &HFE4F&, BlockData(10) + 3 * &H20& + &H80& + &H10&: SendByte &HFE40&, &H0&: SendByte &HFE40&, &H8&
End Sub


