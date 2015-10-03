Attribute VB_Name = "Processor6502Debug"
Option Explicit

Public A As Long
Public X As Long
Public Y As Long
Public S As Long

Public PC As Long
Public PC1 As Long
Public PC2 As Long

Public N As Long ' 128
Public V As Long ' 64
Public B As Long  ' 16
Public D As Long ' 8
Public I As Long ' 4
Public Z As Long ' 2
Public C As Long '1

Public IRQFlag As Boolean
Public NMIFlag As Boolean
Public StopReason As Long
Public mbResetSystemVia As Boolean

Public IRQRaisedBy As String

Public mlTotalCycles As Long

Private Type TwoByte
    Lo As Byte
    Hi As Byte
End Type

Private oFSO As New FileSystemObject
Private oTS As TextStream
    
Private mbHexOn As Boolean
Private mbTraceOn As Boolean
Private mbTraceBrances As Boolean

Private sAddresses() As String
Private lAddressIndex As Long

Private mlInsCount(255) As Long

Private mlCycleTable(255) As Long
Private mlRelative(255) As Long

Public lPCPrevious As Long
    
Public NMISource As String

Public Sub Initialise6502()
    mbTraceOn = False
    mbTraceBrances = False
    
    InitialiseCycleTable
    InitialiseRelative
    IRQFlag = False
    NMIFlag = False
    RES
    StopReason = srNone
End Sub

Private Sub InitialiseCycleTable()
    Dim lInstruction As Long
    Dim lOpCode As Long
    Dim lMode As Long
    Dim lSwitch As Long
    Dim lCycles As Long
    
    ' Debugging.WriteString "Processor6502.InitialiseCycleTable"
    
    ' Regular instructions and addressing modes
    For lInstruction = 0 To 255
        lCycles = 0
        Select Case lInstruction
            Case &H0  ' BRK 4.800
                lCycles = 7&
            Case &H10, &H30, &H50, &H70, &H90, &HB0, &HD0, &HF0 ' BPL' BMI' BVC' BVS' BCC' BCS ' BNE ' BEQ
                lCycles = 2&
            Case &H20, &H60, &H40 ' JSR abs' RTS' RTI
                lCycles = 6&
            Case &H8, &H48  ' PHP' PHA
                lCycles = 3&
            Case &H28, &H68 ' PLP' PLA
                lCycles = 4&
            Case &H88, &HC8, &HE8, &HCA, &HA8, &H98, &H8A, &H9A, &HAA, &HBA ' DEY ' INY INX' DEX ' TAY' TYA ' TXA ' TXS ' TAX' TSX
                lCycles = 2&
            Case &H18, &H38, &H58, &H78, &HB8, &HD8, &HF8 ' CLC' SEC' CLI' SEI ' CLV' CLD' SED
                lCycles = 2&
            Case &HEA ' NOP
                lCycles = 2&
            Case &H2, &H12, &H22, &H32, &H42, &H52, &H62, &H72, &H82, &H92, &HB2, &HC2, &HD2, &HE2, &HF2 ' unassigned
                lCycles = 2&
            Case Else
                lOpCode = lInstruction And &HE0
                lMode = lInstruction And &H1C
                lSwitch = lInstruction And &H3
                
                Select Case lSwitch
                    Case &H0
                        Select Case lMode
                            Case &H0    ' #immediate
                                lCycles = 0&
                            Case &H4  ' zero page
                                lCycles = 1&
                            Case &HC, &H14, &H1C  ' absolute 'zero page,X ' absolute,X
                                lCycles = 2&
                        End Select
                        Select Case lOpCode
                            Case &H40  ' JMP 3.407
                                lCycles = 3&
                            Case &H60  ' JMP (abs)
                                lCycles = 5&
                            Case &H20, &H80, &HA0, &HC0, &HE0 ' BIT  ' STY' LDY ' CPY ' CPX
                                lCycles = lCycles + 2&
                        End Select
                            
                    Case &H1
                        Select Case lMode
                            Case &H0  ' (zero page,X)
                                lCycles = 4&
                            Case &H4  ' zero page
                                lCycles = 1&
                            Case &H8  ' #immediate
                                lCycles = 0&
                            Case &HC, &H14, &H18, &H1C  ' absolute' zero page,X' absolute,Y ' absolute,X
                                lCycles = 2&
                            Case &H10 ' (zero page),Y
                                lCycles = 3&
                        End Select
                        Select Case lOpCode
                            Case &H0, &H20, &H40, &H60, &H80, &HA0, &HC0, &HE0 ' ORA ' AND' EOR' ADC' STA' LDA' CMP' SBC
                                lCycles = lCycles + 2&
                        End Select
                    Case &H2
                        Select Case lMode
                            Case &H0  ' #immediate
                                lCycles = 0&
                            Case &H4 ' zero page'
                                lCycles = 1&
                            Case &H8  ' accumulator
                                lCycles = -2&
                            Case &HC, &H1C, &H14 '  absolute absolute,X / absolute,Y ' zero page,X / zero page,Y
                                lCycles = 2&
                        End Select
                        Select Case lOpCode
                            Case &H0, &H20, &H40, &H60 ' ASL ' ROL ' LSR ' ROR
                                lCycles = lCycles + 4&
                                If lMode = &H1C Then
                                    lCycles = lCycles + 1
                                End If
                            Case &H80, &HA0 ' STX' LDX
                                lCycles = lCycles + 2&
                            Case &HC0, &HE0 ' DEC ' INC
                                lCycles = lCycles + 4&
                                If lMode = &H1C Then
                                    lCycles = lCycles + 1
                                End If
                        End Select
                    Case &H3
                        lCycles = 2&
                End Select
        End Select
        If lCycles = 0 Then
            lCycles = 2
        End If
        mlCycleTable(lInstruction) = lCycles
    Next
End Sub

Private Sub InitialiseRelative()
    Dim lTemp As Long
    
    For lTemp = 0 To 255
        If lTemp < 128 Then
            mlRelative(lTemp) = lTemp
        Else
            mlRelative(lTemp) = lTemp - 256&
        End If
    Next
End Sub

Public Sub DumpInsCountTo127()
    Dim lIndex As Long
    
    For lIndex = 0 To 127
        Debug.Print HexNum(lIndex, 2) & vbTab & mlInsCount(lIndex)
    Next
End Sub

Public Sub DumpInsCountTo255()
    Dim lIndex As Long
    
    For lIndex = 128 To 255
        Debug.Print HexNum(lIndex, 2) & vbTab & mlInsCount(lIndex)
    Next
End Sub


Private Function AddAddress(ByVal sAddress As String) As Long
    Dim lTest As Long
    
    For lTest = 0 To lAddressIndex - 1
        If sAddresses(lTest) = sAddress Then
            AddAddress = lTest
            Exit Function
        End If
    Next
    ReDim Preserve sAddresses(lAddressIndex)
    sAddresses(lAddressIndex) = sAddress
    lAddressIndex = lAddressIndex + 1
    AddAddress = -1
End Function

Public Sub LogPC(ByVal lPC As Long, Optional ByVal bFirst As Boolean = False, Optional ByVal sTitle As String)
    If mbTraceOn Then
        If bFirst Then
            oTS.Write HexNum(RomSelect.SelectedBank, 2) & ":" & HexNum(lPC, 4) & ":"
        Else
            oTS.WriteLine HexNum(RomSelect.SelectedBank, 2) & ":" & HexNum(lPC, 4) & " " & sTitle
        End If
    End If
End Sub
    
Public Sub DumpStack()
    Dim lMem As Long
    
    If mbTraceOn Then
        oTS.Write "Stack:"
        For lMem = 255 To ((S + 1) And &HFF&) Step -1
            oTS.Write HexNum(gyMem(&H100& + lMem), 2)
        Next
        oTS.WriteLine ""
    End If
End Sub

Public Sub Execute()
    Dim lInstruction As Long
    Dim lOpCode As Long
    Dim lMode As Long
    Dim lSwitch As Long
    Dim lLocation As Long
    Dim lLocationTemp As Long
    Dim lTemp As Long
    Dim lCycles As Long
    Dim lCycleCount As Long
    Dim tbTemp As TwoByte

    Dim lSum As Long
    Dim lTemp1 As Long
    Dim lTemp2 As Long
    
    Dim lALoNybble As Long
    Dim lLoNybble As Long
    Dim lHalfCarry As Long
    Dim lHiNybble As Long
    
    Dim lOffsetCycles As Long
    
    Do
        PC1 = (PC + 1&) And &HFFFF&
        PC2 = (PC + 2&) And &HFFFF&
        
        lInstruction = CLng(Mem(PC))
        lOffsetCycles = 0

        If mbHexOn Then
            Debug.Print HexNum$(PC, 4) & ":" & HexNum$(lPCPrevious, 4) & " A:" & A & " X:" & X & " Y:" & Y
            'Stop
        End If

        If HexNum(PC, 4) >= "7B00" And HexNum(PC, 4) <= "7BFF" Then
            'Debug.Print HexNum$(PC, 4) & ":" & HexNum$(lPCPrevious, 4) & " A:" & A & " X:" & X & " Y:" & Y & " NVZC:" & N \ 128 & V \ 64 & Z \ 2 & C & " MEM:" & Mem(Mem(7) + Mem(8) * 256)
'            Stop
'            mbHexOn = True
        Else
            'mbHexOn = False
        End If
        
        lPCPrevious = PC
       
        Select Case lInstruction
            Case &H0  ' BRK 4.800
                LogPC PC, True
                CopyMemory tbTemp, PC2, &H2
                gyMem(S) = tbTemp.Hi
                S = S - 1&: If S < &H100& Then S = &H1FF&
                gyMem(S) = tbTemp.Lo
                S = S - 1&: If S < &H100& Then S = &H1FF&
                gyMem(S) = N + V + 48 + D + I + Z + C ' PSR OR &H10 (break flag)
                S = S - 1&: If S < &H100& Then S = &H1FF&
                CopyMemory PC, gyMem(&HFFFE&), 2&
                PC = PC - 1&
                B = 16
                I = 4
                LogPC PC + 1, False, "BRK"
                DumpStack
            Case &H10 ' BPL
                If mbTraceBrances Then LogPC PC, True
                If N = 0& Then PC = PC1 + mlRelative(Mem(PC1)): lOffsetCycles = 1& + Sgn((PC2 Xor (PC + 1)) And 256&) Else PC = PC1 ' +1 page boundary
                If mbTraceBrances Then LogPC PC + 1, False, "BPL"
            Case &H30 ' BMI
                If mbTraceBrances Then LogPC PC, True
                If N = 0& Then PC = PC1 Else PC = PC1 + mlRelative(Mem(PC1)): lOffsetCycles = 1& + Sgn((PC2 Xor (PC + 1)) And 256&)
                If mbTraceBrances Then LogPC PC + 1, False, "BMI"
            Case &H50 ' BVC
                If mbTraceBrances Then LogPC PC, True
                If V = 0& Then PC = PC1 + mlRelative(Mem(PC1)): lOffsetCycles = 1& + Sgn((PC2 Xor (PC + 1)) And 256&) Else PC = PC1
                If mbTraceBrances Then LogPC PC + 1, False, "BVC"
            Case &H70 ' BVS
                If mbTraceBrances Then LogPC PC, True
                If V = 0& Then PC = PC1 Else PC = PC1 + mlRelative(Mem(PC1)): lOffsetCycles = 1& + Sgn((PC2 Xor (PC + 1)) And 256&)
                If mbTraceBrances Then LogPC PC + 1, False, "BVS"
            Case &H90 ' BCC
                If mbTraceBrances Then LogPC PC, True
                If C = 0& Then PC = PC1 + mlRelative(Mem(PC1)): lOffsetCycles = 1& + Sgn((PC2 Xor (PC + 1)) And 256&) Else PC = PC1
                If mbTraceBrances Then LogPC PC + 1, False, "BCC"
            Case &HB0 ' BCS
                If mbTraceBrances Then LogPC PC, True
                If C = 0& Then PC = PC1 Else PC = PC1 + mlRelative(Mem(PC1)): lOffsetCycles = 1& + Sgn((PC2 Xor (PC + 1)) And 256&)
                If mbTraceBrances Then LogPC PC + 1, False, "BCS"
            Case &HD0 ' BNE
                If mbTraceBrances Then LogPC PC, True
                If Z = 0& Then PC = PC1 + mlRelative(Mem(PC1)): lOffsetCycles = 1& + Sgn((PC2 Xor (PC + 1)) And 256&) Else PC = PC1
                If mbTraceBrances Then LogPC PC + 1, False, "BNE"
            Case &HF0 ' BEQ
                If mbTraceBrances Then LogPC PC, True
                If Z = 0& Then PC = PC1 Else PC = PC1 + mlRelative(Mem(PC1)): lOffsetCycles = 1& + Sgn((PC2 Xor (PC + 1)) And 256&)
                If mbTraceBrances Then LogPC PC + 1, False, "BEQ"
            Case &H20 ' JSR abs
                LogPC PC, True
                CopyMemory tbTemp, PC2, 2&
                gyMem(S) = tbTemp.Hi
                S = S - 1&: If S < &H100& Then S = &H1FF&
                gyMem(S) = tbTemp.Lo
                S = S - 1&: If S < &H100& Then S = &H1FF&
                PC = Mem(PC1) + Mem(PC2) * 256& - 1&
                LogPC PC + 1, , "JSR"
                DumpStack
            Case &H8  ' PHP
                gyMem(S) = N + V + 32& + B + D + I + Z + C
                S = S - 1&: If S < &H100& Then S = &H1FF&
                DumpStack
            Case &H28 ' PLP
                S = S + 1&: If S > &H1FF& Then S = &H100&
                lTemp = gyMem(S)
                N = lTemp And 128&
                V = lTemp And 64&
                B = lTemp And 16&
                D = lTemp And 8&
                I = lTemp And 4&
                Z = lTemp And 2&
                DumpStack
                C = lTemp And 1&
            Case &H48 ' PHA
                gyMem(S) = A
                S = S - 1&: If S < &H100& Then S = &H1FF&
                DumpStack
            Case &H68 ' PLA
                S = S + 1&: If S > &H1FF& Then S = &H100&
                A = gyMem(S)
                N = A And 128&
                Z = (A = 0&) * -2&
                DumpStack
            Case &H88 ' DEY
                Y = (Y - 1&) And 255&
                N = Y And 128&
                Z = (Y = 0&) * -2&
            Case &HC8 ' INY
                Y = (Y + 1&) And 255&
                N = Y And 128&
                Z = (Y = 0&) * -2&
            Case &HE8 ' INX
                X = (X + 1&) And 255&
                N = X And 128&
                Z = (X = 0&) * -2&
            Case &HCA ' DEX
                X = (X - 1&) And 255&
                N = X And 128&
                Z = (X = 0&) * -2&
            Case &HA8 ' TAY
                Y = A
                N = Y And 128&
                Z = (Y = 0&) * -2&
            Case &H18 ' CLC
                C = 0&
            Case &H38 ' SEC
                C = 1&
            Case &H58 ' CLI
                I = 0&
            Case &H78 ' SEI
                I = 4&
            Case &H98 ' TYA
                A = Y
                N = A And 128&
                Z = (A = 0&) * -2&
            Case &HB8 ' CLV
                V = 0&
            Case &HD8 ' CLD
                D = 0&
            Case &HF8 ' SED
                D = 8&
            Case &H8A ' TXA
                A = X
                N = A And 128
                Z = (A = 0&) * -2&
            Case &H9A ' TXS
                S = X + &H100&
            Case &HAA ' TAX
                X = A
                N = X And 128&
                Z = (X = 0&) * -2&
            Case &HBA ' TSX
                X = S And &HFF&
                N = X And 128&
                Z = (X = 0&) * -2&
            Case &HEA ' NOP
                ' do nothing
            Case &H40 ' RTI
                LogPC PC, True
                S = S + 1&: If S > &H1FF& Then S = &H100&
                lTemp = gyMem(S) ' pull PSR
                N = lTemp And 128&
                V = lTemp And 64&
                B = lTemp And 16&
                D = lTemp And 8&
                I = lTemp And 4&
                Z = lTemp And 2&
                C = lTemp And 1&
                S = S + 1&: If S > &H1FF& Then S = &H100&
                PC = gyMem(S) - 1&
                S = S + 1&: If S > &H1FF& Then S = &H100&
                PC = PC + gyMem(S) * 256&
                LogPC PC + 1, False, "RTI"
                DumpStack
            Case &H60 ' RTS
                LogPC PC, True
                S = S + 1&: If S > &H1FF& Then S = &H100&
                PC = gyMem(S)
                S = S + 1&: If S > &H1FF& Then S = &H100&
                PC = PC + gyMem(S) * 256&
                LogPC PC + 1, False, "RTS"
                DumpStack
            Case &H2, &H12, &H22, &H32, &H42, &H52, &H62, &H72, &H82, &H92, &HB2, &HC2, &HD2, &HE2, &HF2
                PC = PC - 1
                NMIFlag = False
                IRQFlag = False
                LogPC PC, True
                LogPC PC + 1, False, "HALT"
            Case Else
                lOpCode = lInstruction And &HE0&
                lMode = lInstruction And &H1C&
                lSwitch = lInstruction And &H3&
                Select Case lSwitch
                    Case &H0
                        Select Case lMode
                            Case &H0&    ' #immediate
                                lLocation = PC1
                                PC = PC1
                            Case &H4&  ' zero page
                                lLocation = Mem(PC1)
                                PC = PC1
                            Case &HC&  ' absolute
                                lLocation = Mem(PC1) + Mem(PC2) * 256&
                                PC = PC2
                            Case &H14& ' zero page,X
                                lLocation = (Mem(PC1) + X) And &HFF&
                                PC = PC1
                            Case &H1C&  ' absolute,X
                                lLocationTemp = Mem(PC1) + Mem(PC2) * 256&
                                lLocation = (lLocationTemp + X) And &HFFFF&
                                lOffsetCycles = Sgn((lLocationTemp Xor lLocation) And 256&) ' +1 page boundary
                                PC = PC2
                        End Select
                        Select Case lOpCode
                            Case &H20  ' BIT zp:3.686 abs:3.815
                                Z = ((A And Mem(lLocation)) = 0&) * -2&
                                N = Mem(lLocation) And 128&
                                V = Mem(lLocation) And 64&
                            Case &H40  ' JMP 3.407
                                LogPC PC, True
                                PC = lLocation - 1&
                                LogPC PC + 1, False, "JMP"
                            Case &H60  ' JMP (abs)
                                LogPC PC, True
                                lTemp = Mem(lLocation)
                                If (lLocation And &HFF&) <> &HFF& Then
                                    PC = lTemp + Mem(lLocation + 1&) * 256& - 1&
                                Else
                                    PC = lTemp + Mem(lLocation - 255&) * 256& - 1&
                                End If
                                LogPC PC + 1, False, "JMP (" & HexNum(lLocation, 4) & ")"
                            Case &H80  ' STY
                                Mem(lLocation) = Y
                            Case &HA0  ' LDY
                                Y = Mem(lLocation)
                                N = Y And 128&
                                Z = (Y = 0&) * -2&
                            Case &HC0  ' CPY
                                lTemp = Y - Mem(lLocation)
                                N = lTemp And 128&
                                Z = (lTemp = 0&) * -2&
                                C = Abs(lTemp >= 0&)
                            Case &HE0  ' CPX
                                lTemp = X - Mem(lLocation)
                                N = lTemp And 128&
                                Z = (lTemp = 0&) * -2&
                                C = Abs(lTemp >= 0&)
                        End Select
                        
                    Case &H1
                        Select Case lMode
                            Case &H0  ' (zero page,X)
                                lLocation = (Mem(PC1) + X) And 255&
                                lLocation = gyMem(lLocation) + gyMem((lLocation + 1&) And 255&) * 256&
                                PC = PC1
                            Case &H4  ' zero page
                                lLocation = Mem(PC1)
                                PC = PC1
                            Case &H8  ' #immediate
                                PC = PC1
                                lLocation = PC
                            Case &HC  ' absolute
                                lLocation = Mem(PC1) + Mem(PC2) * 256&
                                PC = PC2
                            Case &H10 ' (zero page),Y
                                lTemp = Mem(PC1)
                                lLocationTemp = gyMem(lTemp) + gyMem((lTemp + 1&) And 255&) * 256&
                                lLocation = (lLocationTemp + Y) And &HFFFF&
                                If lOpCode <> &H80 Then
                                    lOffsetCycles = Sgn((lLocationTemp Xor lLocation) And 256&) ' +1 page boundary
                                Else
                                    lOffsetCycles = 1
                                End If
                                PC = PC1
                            Case &H14  ' zero page,X
                                lLocation = (Mem(PC1) + X) And 255&
                                PC = PC1
                            Case &H18  ' absolute,Y
                                lLocationTemp = Mem(PC1) + Mem(PC2) * 256&
                                lLocation = (lLocationTemp + Y) And &HFFFF&
                                If lOpCode <> &H80 Then
                                    lOffsetCycles = Sgn((lLocationTemp Xor lLocation) And 256&) ' +1 page boundary
                                Else
                                    lOffsetCycles = 1
                                End If
                                PC = PC2
                            Case &H1C  ' absolute,X
                                lLocationTemp = Mem(PC1) + Mem(PC2) * 256&
                                lLocation = (lLocationTemp + X) And &HFFFF&
                                If lOpCode <> &H80 Then
                                    lOffsetCycles = Sgn((lLocationTemp Xor lLocation) And 256&) ' +1 page boundary
                                Else
                                    lOffsetCycles = 1
                                End If
                                PC = PC2
                        End Select
                        Select Case lOpCode
                            Case &H0 ' ORA
                                A = A Or Mem(lLocation)
                                N = A And 128&
                                Z = (A = 0&) * -2&
                            Case &H20 ' AND
                                A = A And Mem(lLocation)
                                N = A And 128&
                                Z = (A = 0&) * -2&
                            Case &H40 ' EOR
                                A = A Xor Mem(lLocation)
                                N = A And 128&
                                Z = (A = 0&) * -2&
                            Case &H60 ' ADC
                                If D = 0 Then
                                    Select Case mlRelative(A) + mlRelative(Mem(lLocation)) + C
                                        Case Is > 127&, Is < -128&
                                            V = 64
                                        Case Else
                                            V = 0
                                    End Select
                                    lSum = A + Mem(lLocation) + C
                                    C = Abs(lSum > 255&)
                                    A = lSum And 255&
                                    N = A And 128&
                                    Z = (A = 0&) * -2
                                Else
                                    lALoNybble = A And &HF
                                    lSum = A + Mem(lLocation)
                                    Z = (((lSum + C) And &HFF&) = 0&) * -2
                                    lSum = lSum + (lALoNybble >= 10&) * (lALoNybble - 10&) - 250&
                                    V = ((lSum >= -128&) And (lSum <= 127&)) * -64&
                                    V = V And (((A Xor Mem(lLocation)) And &H80&) = 0&)
                                    lLoNybble = lALoNybble + (Mem(lLocation) And &HF&) + C
                                    lHalfCarry = lLoNybble >= 10&
                                    lHiNybble = ((A And &HF0&) + (Mem(lLocation) And &HF0&)) - lHalfCarry * 16&
                                    C = Abs(lHiNybble >= 160&)
                                    A = (((lLoNybble + lHalfCarry * 10&) And &HF)) + ((lHiNybble - C * 160&) And &HF0&)
                                    N = lSum And 128&
                                End If
                            Case &H80 ' STA
                                Mem(lLocation) = A
                            Case &HA0 ' LDA
                                A = Mem(lLocation)
                                N = A And 128&
                                Z = (A = 0&) * -2&
                            Case &HC0 ' CMP
                                lTemp = A - Mem(lLocation)
                                N = lTemp And 128&
                                Z = (lTemp = 0&) * -2&
                                C = Abs(lTemp >= 0&)
                            Case &HE0 ' SBC
                                Dim lMemLoNybble As Long
                                Dim lAHiNybble As Long
                                Dim lMemHiNybble As Long
                                Dim lNegC As Long
                                Dim lLoDifference As Long
                                Dim lHiDifference As Long
                                Dim lLoDifferenceNegC As Long
                                Dim lLoDifferenceNegC10 As Long
                                
                                lNegC = 1& - C
                                
                                Select Case mlRelative(A) - mlRelative(Mem(lLocation)) - lNegC
                                    Case Is > 127&, Is < -128&
                                        V = 64&
                                    Case Else
                                        V = 0&
                                End Select
                                lSum = A - Mem(lLocation) - lNegC
                                C = Abs(lSum >= 0&)
                                
                                If D = 0& Then
                                    A = lSum And 255&
                                Else
                                    lALoNybble = A And &HF&
                                    lMemLoNybble = Mem(lLocation) And &HF&
                                    
                                    lAHiNybble = A And &HF0&
                                    lMemHiNybble = Mem(lLocation) And &HF0&
                                    
                                    lLoDifference = lALoNybble - lMemLoNybble
                                    lHiDifference = lAHiNybble - lMemHiNybble
                                    
                                    lLoDifferenceNegC = lLoDifference < lNegC
                                    lLoDifferenceNegC10 = lLoDifference < (lNegC - 10&)
                                    lLoNybble = lLoDifference - lNegC + ((lLoDifferenceNegC And lALoNybble < 10&) + lLoDifferenceNegC10) * 6&
                                    
                                    lHalfCarry = (lLoDifferenceNegC + lLoDifferenceNegC10) * 16&
                                    
                                    lHiNybble = lHiDifference + lHalfCarry + (lHiDifference < -lHalfCarry) * 6& * 16&
                                    A = (lLoNybble And &HF&) + (lHiNybble And &HF0&)
                                End If
                                N = lSum And 128&
                                Z = (lSum = 0&) * -2&
                        End Select
                    Case &H2
                        Select Case lMode
                            Case &H0&  ' #immediate
                                PC = PC1
                                lLocation = PC
                            Case &H4&  ' zero page
                                lLocation = Mem(PC1)
                                PC = PC1
                            Case &H8&  ' accumulator
                            Case &HC&  ' absolute
                                lLocation = Mem(PC1) + Mem(PC2) * 256&
                                PC = PC2
                            Case &H14&  ' zero page,X / zero page,Y
                                If lOpCode <> &H80& And lOpCode <> &HA0& Then 'STX
                                    lLocation = (Mem(PC1) + X) And 255&
                                Else
                                    lLocation = (Mem(PC1) + Y) And 255&
                                End If
                                PC = PC1
                            Case &H1C&  ' absolute,X / absolute,Y
                                lLocationTemp = Mem(PC1) + Mem(PC2) * 256&
                                
                                If lOpCode <> &HA0& Then
                                    lLocation = (lLocationTemp + X) And &HFFFF&
                                Else
                                    lLocation = (lLocationTemp + Y) And &HFFFF&
                                End If
                                If lOpCode = &HA0& Then ' LDX
                                    lOffsetCycles = Sgn((lLocationTemp Xor lLocation) And 256&) ' +1 page boundary
                                Else
                                    lOffsetCycles = 0
                                End If
                                PC = PC2
                        End Select
                        Select Case lOpCode
                            Case &H0&  ' ASL
                                If lMode <> &H8& Then
                                    C = Sgn(Mem(lLocation) And 128&)
                                    Mem(lLocation) = (Mem(lLocation) * 2&) And &HFE&
                                    N = Mem(lLocation) And 128&
                                    Z = (Mem(lLocation) = 0&) * -2&
                                Else
                                    C = Sgn(A And 128&)
                                    A = (A * 2&) And &HFE&
                                    N = A And 128&
                                    Z = (A = 0&) * -2&
                                End If
                            Case &H20&  ' ROL
                                lTemp = C
                                If lMode <> &H8& Then
                                    C = Sgn(Mem(lLocation) And 128&)
                                    lTemp = (Mem(lLocation) * 2) And &HFE& Or lTemp
                                    Mem(lLocation) = lTemp
                                    N = lTemp And 128&
                                    Z = (lTemp = 0&) * -2&
                                Else
                                    C = Sgn(A And 128&)
                                    A = (A * 2&) And &HFE& Or lTemp
                                    N = A And 128&
                                    Z = (A = 0&) * -2&
                                End If
                            Case &H40&  ' LSR
                                If lMode <> &H8& Then
                                    C = Mem(lLocation) And 1&
                                    Mem(lLocation) = Mem(lLocation) \ 2&
                                    N = 0&
                                    Z = (Mem(lLocation) = 0&) * -2&
                                Else
                                    C = A And 1&
                                    A = A \ 2&
                                    N = 0&
                                    Z = (A = 0&) * -2&
                                End If
                            Case &H60&  ' ROR
                                N = C * 128&
                                If lMode <> &H8 Then
                                    C = Mem(lLocation) And 1&
                                    lTemp = (Mem(lLocation) \ 2&) Or N
                                    Mem(lLocation) = lTemp
                                    Z = (lTemp = 0&) * -2&
                                Else
                                    C = A And 1&
                                    A = (A \ 2) Or N
                                    Z = (A = 0) * -2&
                                End If
                            Case &H80&  ' STX
                                Mem(lLocation) = X
                                lCycles = lCycles + 2&
                            Case &HA0&  ' LDX
                                X = Mem(lLocation)
                                N = X And 128&
                                Z = (X = 0&) * -2&
                            Case &HC0&  ' DEC
                                lTemp = (Mem(lLocation) - 1&) And 255&
                                Mem(lLocation) = lTemp
                                N = lTemp And 128&
                                Z = (lTemp = 0&) * -2&
                            Case &HE0&  ' INC
                                lTemp = (Mem(lLocation) + 1&) And 255&
                                Mem(lLocation) = lTemp
                                N = lTemp And 128&
                                Z = (lTemp = 0) * -2&
                        End Select
                    Case &H3&
                        'Exit Do
                End Select
        End Select
        PC = PC + 1&
        
        lCycles = mlCycleTable(lInstruction) + lOffsetCycles  ' Cycles at 2 MHz
        
        mlTotalCycles = mlTotalCycles + lCycles

        If mlTotalCycles > 20 Then
            If mbResetSystemVia Then
                IRQFlag = False
                NMIFlag = False
                ResetSystemVIA
                mbResetSystemVia = False
            End If
        End If

        VideoULA.Tick lCycles
        Keyboard.Tick lCycles
        KeyboardIndicators.Tick lCycles
        SystemVIA6522.TimersTick lCycles
        UserVIA6522.TimersTick lCycles
        ACIA6850.Tick lCycles
        FDC8271.Tick lCycles
        
        mlTotalCycles = mlTotalCycles + ProcessInterrupt ' Cycles at 2 MHz
        
        Throttle.ThrottleTick lCycles
    Loop While StopReason = srNone
End Sub

Private Function ProcessInterrupt() As Long
    If NMIFlag Then
        NMIFlag = False
        ProcessInterrupt = NMI
    ElseIf IRQFlag And I = 0 Then
        ProcessInterrupt = IRQ
    End If
End Function

Private Function IRQ() As Long
    Dim tbTemp As TwoByte
    
    LogPC PC, True
    CopyMemory tbTemp, PC, &H2
    gyMem(S) = tbTemp.Hi
    S = S - 1&: If S < &H100& Then S = &H1FF&
    gyMem(S) = tbTemp.Lo
    S = S - 1&: If S < &H100& Then S = &H1FF&
    gyMem(S) = N + V + 32& + D + I + Z + C
    S = S - 1&: If S < &H100& Then S = &H1FF&
    B = 0
    I = 4 ' disabled
    PC = gyMem(&HFFFE&) + gyMem(&HFFFF&) * 256&
    IRQ = 7
    LogPC PC, False, "IRQ:" & IRQRaisedBy & ":"
    DumpStack
End Function

Private Function NMI() As Long
    Dim tbTemp As TwoByte
    
    LogPC PC - 1, True
    CopyMemory tbTemp, PC, &H2
    gyMem(S) = tbTemp.Hi
    S = S - 1&: If S < &H100& Then S = &H1FF&
    gyMem(S) = tbTemp.Lo
    S = S - 1&: If S < &H100& Then S = &H1FF&
    gyMem(S) = N + V + 32& + D + I + Z + C
    S = S - 1&: If S < &H100& Then S = &H1FF&
    PC = gyMem(&HFFFA&) + gyMem(&HFFFB&) * 256&
    I = 4 ' disabled
    NMI = 7 ' Cycles needs checking
    LogPC PC, False, "NMI (" & NMISource & "):"
    DumpStack
End Function

Public Function RES() As Long
    If Not oTS Is Nothing Then
        oTS.Close
        Set oTS = Nothing
    End If
    Set oTS = oFSO.OpenTextFile(App.path & "\trace\trace.txt", ForWriting, True)
    
    LogPC PC, True
    A = 0&
    X = 0&
    Y = 0&
    S = &H1FF&
    N = 0&
    C = 0&
    Z = 0&
    B = 0&
    D = 0&
    I = 4&
    PC = gyMem(&HFFFC&) + gyMem(&HFFFD&) * 256&
    IRQFlag = False
    NMIFlag = False
    mlTotalCycles = 0
    mbResetSystemVia = True
    LogPC PC, False, "RES"
End Function
