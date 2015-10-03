Attribute VB_Name = "Processor6502"
Option Explicit

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Public A As Long
Public X As Long
Public Y As Long
Public S As Long

Public PC As Long
Public PC1 As Long
Public PC2 As Long

Public N As Long ' 128
Public V As Long ' 64
Public B As Long ' 16
Public D As Long ' 8
Public I As Long ' 4
Public Z As Long ' 2
Public C As Long '1

Public IRQFlag As Boolean
Public NMIFlag As Boolean
Public RESFlag As Boolean
Public STOPFlag As Boolean

Private Type TwoByte
    Lo As Byte
    Hi As Byte
End Type


' The following code is optimised for speed, change at your own peril
' To do: BCD for SBC opcode.
Public Sub Execute()
    Dim iInstruction As Integer
    Dim iOpcode As Integer
    Dim iMode As Integer
    Dim iSwitch As Integer
    Dim lLocation As Long
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
    
    STOPFlag = False
    Do
        PC1 = PC + 1&
        PC2 = PC + 2&
        
        iInstruction = gyMem(PC)
        
        Select Case iInstruction
            Case &H0  ' BRK 4.800
                CopyMemory tbTemp, PC2, &H2
                Mem(S) = tbTemp.Hi
                S = S - 1&: If S < &H100& Then S = &H1FF&
                Mem(S) = tbTemp.Lo
                S = S - 1&: If S < &H100& Then S = &H1FF&
                Mem(S) = N + V + 48 + D + I + Z + C ' PSR OR &H10 (break flag)
                S = S - 1&: If S < &H100& Then S = &H1FF&
                PC = gyMem(&HFFFE&) + gyMem(&HFFFF&) * 256& - 1&
                B = 16
                I = 4
                lCycles = 7&
            Case &H10 ' BPL
                If N Then PC = PC1: lCycles = 2& Else PC = PC1 + gyMem(PC1) + (gyMem(PC1) > 127) * 256&: lCycles = 3& ' +1 page boundary
            Case &H30 ' BMI
                If N Then PC = PC1 + gyMem(PC1) + (gyMem(PC1) > 127) * 256&: lCycles = 3& Else PC = PC1: lCycles = 2&
            Case &H50 ' BVC
                If V Then PC = PC1: lCycles = 2& Else PC = PC1 + gyMem(PC1) + (gyMem(PC1) > 127) * 256&: lCycles = 3&
            Case &H70 ' BVS
                If V Then PC = PC1 + gyMem(PC1) + (gyMem(PC1) > 127) * 256&: lCycles = 3& Else PC = PC1: lCycles = 2&
            Case &H90 ' BCC
                If C Then PC = PC1: lCycles = 2& Else PC = PC1 + gyMem(PC1) + (gyMem(PC1) > 127) * 256&: lCycles = 3&
            Case &HB0 ' BCS
                If C Then PC = PC1 + gyMem(PC1) + (gyMem(PC1) > 127) * 256&: lCycles = 3& Else PC = PC1: lCycles = 2&
            Case &HD0 ' BNE
                If Z Then PC = PC1: lCycles = 2& Else PC = PC1 + gyMem(PC1) + (gyMem(PC1) > 127) * 256&: lCycles = 3&
            Case &HF0 ' BEQ
                If Z Then PC = PC1 + gyMem(PC1) + (gyMem(PC1) > 127) * 256&: lCycles = 3& Else PC = PC1: lCycles = 2&
            Case &H20 ' JSR abs
                CopyMemory tbTemp, PC2, 2&
                Mem(S) = tbTemp.Hi
                S = S - 1&: If S < &H100& Then S = &H1FF&
                Mem(S) = tbTemp.Lo
                S = S - 1&: If S < &H100& Then S = &H1FF&
                PC = gyMem(PC1) + gyMem(PC2) * 256& - 1&
                lCycles = 6&
            Case &H8  ' PHP
                Mem(S) = N + V + 32 + B + D + I + Z + C
                S = S - 1&: If S < &H100& Then S = &H1FF&
                lCycles = 3&
            Case &H28 ' PLP
                S = S + 1&: If S > &H1FF& Then S = &H100&
                lTemp = gyMem(S)
                N = lTemp And 128
                V = lTemp And 64
                B = lTemp And 16
                D = lTemp And 8
                I = lTemp And 4
                Z = lTemp And 2
                C = lTemp And 1
                lCycles = 4&
            Case &H48 ' PHA
                Mem(S) = A
                S = S - 1&: If S < &H100& Then S = &H1FF&
                lCycles = 3&
            Case &H68 ' PLA
                S = S + 1&: If S > &H1FF& Then S = &H100&
                A = gyMem(S)
                N = A And 128
                Z = (A = 0) * -2
                lCycles = 4&
            Case &H88 ' DEY
                Y = (Y - 1) And 255
                N = Y And 128
                Z = (Y = 0) * -2
                lCycles = 2&
            Case &HC8 ' INY
                Y = (Y + 1) And 255
                N = Y And 128
                Z = (Y = 0) * -2
                lCycles = 2&
            Case &HE8 ' INX
                X = (X + 1) And 255
                N = X And 128
                Z = (X = 0) * -2
                lCycles = 2&
            Case &HCA ' DEX
                X = (X - 1) And 255
                N = X And 128
                Z = (X = 0) * -2
                lCycles = 2&
            Case &HA8 ' TAY
                Y = A
                N = Y And 128
                Z = (Y = 0) * -2
                lCycles = 2&
            Case &H18 ' CLC
                C = 0&
                lCycles = 2&
            Case &H38 ' SEC
                C = 1&
                lCycles = 2&
            Case &H58 ' CLI
                I = 0&
                lCycles = 2&
            Case &H78 ' SEI
                I = 4&
                lCycles = 2&
            Case &H98 ' TYA
                A = Y
                N = A And 128
                Z = (A = 0) * -2
                lCycles = 2&
            Case &HB8 ' CLV
                V = 0&
                lCycles = 2&
            Case &HD8 ' CLD
                D = 0&
                lCycles = 2&
            Case &HF8 ' SED
                D = 8&
                lCycles = 2&
            Case &H8A ' TXA
                A = X
                N = A And 128
                Z = (A = 0) * -2
                lCycles = 2&
            Case &H9A ' TXS
                S = X + &H100&
                lCycles = 2&
            Case &HAA ' TAX
                X = A
                N = X And 128
                Z = (X = 0) * -2
                lCycles = 2&
            Case &HBA ' TSX
                X = S And &HFF&
                N = X And 128
                Z = (X = 0) * -2
                lCycles = 2&
            Case &HEA ' NOP
                lCycles = 2&
            Case &H40 ' RTI
                S = S + 1&: If S > &H1FF& Then S = &H100&
                lTemp = gyMem(S) ' pull PSR
                N = lTemp And 128
                V = lTemp And 64
                B = lTemp And 16
                D = lTemp And 8
                I = lTemp And 4
                Z = lTemp And 2
                C = lTemp And 1
                S = S + 1&: If S > &H1FF& Then S = &H100&
                PC = gyMem(S) - 1&
                S = S + 1&: If S > &H1FF& Then S = &H100&
                PC = PC + gyMem(S) * 256&
                lCycles = 6&
            Case &H60 ' RTS
                S = S + 1&: If S > &H1FF& Then S = &H100&
                PC = gyMem(S)
                S = S + 1&: If S > &H1FF& Then S = &H100&
                PC = PC + gyMem(S) * 256&
                lCycles = 6&
            Case &H2, &H12, &H32, &H42, &H52, &H62, &H72, &H82, &H92, &HB2, &HC2, &HD2, &HE2, &HF2
                lCycles = 2&
                Stop
                'Exit Do
            Case Else
                iOpcode = iInstruction And &HE0
                iMode = iInstruction And &H1C
                iSwitch = iInstruction And &H3
                Select Case iSwitch
                    Case &H0
                        Select Case iMode
                            Case &H0    ' #immediate
                                lLocation = PC1
                                PC = PC1
                                lCycles = 0&
                            Case &H4  ' zero page
                                lLocation = gyMem(PC1)
                                PC = PC1
                                lCycles = 1&
                            Case &HC  ' absolute
                                lLocation = gyMem(PC1) + gyMem(PC2) * 256&
                                PC = PC2
                                lCycles = 2&
                            Case &H14 ' zero page,X
                                lLocation = (gyMem(PC1) + X) And &HFF
                                PC = PC1
                                lCycles = 2&
                            Case &H1C  ' absolute,X
                                lLocation = gyMem(PC1) + X
                                lCycles = 2& + IIf(lLocation >= 256&, 1&, 0&) ' +1 page boundary
                                lLocation = lLocation + gyMem(PC2) * 256&
                                PC = PC2
                        End Select
                        Select Case iOpcode
                            Case &H20  ' BIT zp:3.686 abs:3.815
                                Z = ((A And gyMem(lLocation)) = 0) * -2
                                N = gyMem(lLocation) And 128
                                V = gyMem(lLocation) And 64
                                lCycles = lCycles + 2&
                            Case &H40  ' JMP 3.407
                                PC = lLocation - 1&
                                lCycles = 3&
                            Case &H60  ' JMP (abs)
                                lTemp = gyMem(lLocation)
                                If lTemp <> &HFF& Then
                                    PC = lTemp + gyMem(lLocation + 1&) * 256& - 1&
                                Else
                                    PC = lTemp + gyMem(lLocation - 255) * 256& - 1&
                                End If
                                lCycles = 5&
                            Case &H80  ' STY
                                Mem(lLocation) = Y
                                lCycles = lCycles + 2&
                            Case &HA0  ' LDY
                                Y = gyMem(lLocation)
                                N = Y And 128
                                Z = (Y = 0) * -2
                                lCycles = lCycles + 2&
                            Case &HC0  ' CPY
                                lTemp = Y - gyMem(lLocation)
                                N = lTemp And 128
                                Z = (lTemp = 0) * -2
                                C = Abs(lTemp >= 0&)
                                lCycles = lCycles + 2&
                            Case &HE0  ' CPX
                                lTemp = X - gyMem(lLocation)
                                N = lTemp And 128
                                Z = (lTemp = 0) * -2
                                C = Abs(lTemp >= 0&)
                                lCycles = lCycles + 2&
                        End Select
                        
                    Case &H1
                        Select Case iMode
                            Case &H0  ' (zero page,X)
                                lLocation = (gyMem(PC1) + X) And 255
                                lLocation = gyMem(lLocation) + gyMem((lLocation + 1&) And 255) * 256&
                                PC = PC1
                                lCycles = 4&
                            Case &H4  ' zero page
                                lLocation = gyMem(PC1)
                                PC = PC1
                                lCycles = 1&
                            Case &H8  ' #immediate
                                PC = PC1
                                lLocation = PC
                                lCycles = 0&
                            Case &HC  ' absolute
                                lLocation = gyMem(PC1) + gyMem(PC2) * 256&
                                PC = PC2
                                lCycles = 2&
                            Case &H10 ' (zero page),Y
                                lTemp = gyMem(PC1)
                                lLocation = gyMem(lTemp) + Y
                                lCycles = 3& + IIf(lLocation >= 256&, 1&, 0&) ' +1 page boundary
                                lLocation = (lLocation + gyMem((lTemp + 1) And 255) * 256&) And &HFFFF&
                                PC = PC1
                            Case &H14  ' zero page,X
                                lLocation = (gyMem(PC1) + X) And 255
                                PC = PC1
                                lCycles = 2&
                            Case &H18  ' absolute,Y
                                lLocation = gyMem(PC1) + Y
                                lCycles = 2& + IIf(lLocation >= 256&, 1&, 0&) ' +1 page boundary
                                lLocation = lLocation + gyMem(PC2) * 256&
                                PC = PC2
                                lCycles = 2& ' +1 page boundary
                            Case &H1C  ' absolute,X
                                lLocation = gyMem(PC1) + X
                                lCycles = 2& + IIf(lLocation >= 256&, 1&, 0&) ' +1 page boundary
                                lLocation = lLocation + gyMem(PC2) * 256&
                                PC = PC2
                        End Select
                        Select Case iOpcode
                            Case &H0 ' ORA
                                A = A Or gyMem(lLocation)
                                N = A And 128
                                Z = (A = 0) * -2
                                lCycles = lCycles + 2&
                            Case &H20 ' AND
                                A = A And gyMem(lLocation)
                                N = A And 128
                                Z = (A = 0) * -2
                                lCycles = lCycles + 2&
                            Case &H40 ' EOR
                                A = A Xor gyMem(lLocation)
                                N = A And 128
                                Z = (A = 0) * -2
                                lCycles = lCycles + 2&
                            Case &H60 ' ADC
                                If D = 0 Then
                                    Select Case (A + (A > 127) * 256&) + (gyMem(lLocation) + (gyMem(lLocation) > 127) * 256&) + C
                                        Case Is > 127, Is < -128
                                            V = 64
                                        Case Else
                                            V = 0
                                    End Select
                                    lSum = A + gyMem(lLocation) + C
                                    C = Abs(lSum > 255&)
                                    A = lSum And 255&
                                Else
                                    lALoNybble = A And &HF
                                    lSum = A + gyMem(lLocation) + (lALoNybble >= 10) * (lALoNybble - 10) - 250&
                                    V = ((lSum >= -128) And (lSum <= 127)) * -64
                                    V = V And (((A Xor gyMem(lLocation)) And &H80) = 0)
                                    lLoNybble = lALoNybble + (gyMem(lLocation) And &HF) + C
                                    lHalfCarry = lLoNybble >= 10
                                    lHiNybble = ((A And &HF0) + (gyMem(lLocation) And &HF0)) - lHalfCarry * 16
                                    C = Abs(lHiNybble >= 160)
                                    A = (((lLoNybble + lHalfCarry * 10) And &HF)) + ((lHiNybble - C * 160) And &HF0)
                                End If
                                N = A And 128
                                Z = (A = 0) * -2
                                lCycles = lCycles + 2&
                            Case &H80 ' STA
                                Mem(lLocation) = A
                                lCycles = lCycles + 2&
                            Case &HA0 ' LDA
                                A = gyMem(lLocation)
                                N = A And 128
                                Z = (A = 0) * -2
                                lCycles = lCycles + 2&
                            Case &HC0 ' CMP
                                lTemp = A - gyMem(lLocation)
                                N = lTemp And 128
                                Z = (lTemp = 0) * -2
                                C = Abs(lTemp >= 0&)
                                lCycles = lCycles + 2&
                            Case &HE0 ' SBC
                                Dim lMemLoNybble As Long
                                Dim lAHiNybble As Long
                                Dim lMemHiNybble As Long
                                Dim lNegC As Long
                                Dim lLoDifference As Long
                                Dim lHiDifference As Long
                                Dim lLoDifferenceNegC As Long
                                Dim lLoDifferenceNegC10 As Long
                                
                                lNegC = 1 - C
                                If D = 0 Then
                                    Select Case (A + (A > 127) * 256&) - (gyMem(lLocation) + (gyMem(lLocation) > 127) * 256&) - lNegC
                                        Case Is > 127, Is < -128
                                            V = 64
                                        Case Else
                                            V = 0
                                    End Select
                                    lSum = A - gyMem(lLocation) - lNegC
                                    C = Abs(lSum >= 0&)
                                    A = lSum And 255&
                                Else
                                    lALoNybble = A And &HF
                                    lMemLoNybble = gyMem(lLocation) And &HF
                                    
                                    lAHiNybble = A And &HF0
                                    lMemHiNybble = gyMem(lLocation) And &HF0
                                    
                                    lLoDifference = lALoNybble - lMemLoNybble
                                    lHiDifference = lAHiNybble - lMemHiNybble
                                    
                                    lLoDifferenceNegC = lLoDifference < lNegC
                                    lLoDifferenceNegC10 = lLoDifference < (lNegC - 10)
                                    lLoNybble = lLoDifference - lNegC + ((lLoDifferenceNegC And lALoNybble < 10) + lLoDifferenceNegC10) * 6
                                    
                                    lHalfCarry = (lLoDifferenceNegC + lLoDifferenceNegC10) * 16
                                    
                                    lHiNybble = lHiDifference + lHalfCarry + (lHiDifference < -lHalfCarry) * 6 * 16
                                    C = Abs(lHiDifference >= (-16 * lLoDifferenceNegC))
                                    V = -64 * (((lAHiNybble >= 128 And lMemHiNybble < 128) And lHiDifference < (128 + lLoDifferenceNegC * -16)) Or ((lAHiNybble < 128 And lMemHiNybble >= 128) And lHiDifference >= (-128 + lLoDifferenceNegC * -16)))
                                    A = (lLoNybble And &HF) + (lHiNybble And &HF0)
                                End If
                                N = A And 128
                                Z = (A = 0) * -2
                                lCycles = lCycles + 2&
                        End Select
                    Case &H2
                        Select Case iMode
                            Case &H0  ' #immediate
                                PC = PC1
                                lLocation = PC
                                lCycles = 0&
                            Case &H4  ' zero page
                                lLocation = gyMem(PC1)
                                PC = PC1
                                lCycles = 1&
                            Case &H8  ' accumulator
                            Case &HC  ' absolute
                                lLocation = gyMem(PC1) + gyMem(PC2) * 256&
                                PC = PC2
                                lCycles = 1&
                            Case &H14  ' zero page,X / zero page,Y
                                If iOpcode <> &H80 Then  'STX
                                    lLocation = (gyMem(PC1) + X) And 255
                                Else
                                    lLocation = (gyMem(PC1) + Y) And 255
                                End If
                                PC = PC1
                                lCycles = 1&
                            Case &H1C  ' absolute,X
                                If iOpcode <> &HA0 Then
                                    lLocation = gyMem(PC1) + X
                                Else
                                    lLocation = gyMem(PC1) + Y
                                End If
                                lCycles = 2& + IIf(lLocation >= 256&, 1&, 0&) ' +1 page boundary
                                lLocation = lLocation + gyMem(PC2) * 256&
                                PC = PC2
                        End Select
                        Select Case iOpcode
                            Case &H0  ' ASL
                                If iMode <> &H8 Then
                                    C = (gyMem(lLocation) And 128) \ 128
                                    Mem(lLocation) = (gyMem(lLocation) * 2) And 255
                                    N = gyMem(lLocation) And 128
                                    Z = (gyMem(lLocation) = 0) * -2
                                    lCycles = lCycles + 4&
                                Else
                                    C = Sgn(A And 128)
                                    A = (A * 2) And 255
                                    N = A And 128
                                    Z = (A = 0) * -2
                                    lCycles = 2&
                                End If
                            Case &H20  ' ROL
                                lTemp = C
                                If iMode <> &H8 Then
                                    C = (gyMem(lLocation) And 128) \ 128
                                    lTemp = (gyMem(lLocation) * 2 + lTemp) And 255
                                    Mem(lLocation) = lTemp
                                    N = lTemp And 128
                                    Z = (lTemp = 0) * -2
                                    lCycles = lCycles + 4&
                                Else
                                    C = Sgn(A And 128)
                                    A = (A * 2 + lTemp) And 255
                                    N = A And 128
                                    Z = (A = 0) * -2
                                    lCycles = 2&
                                End If
                            Case &H40  ' LSR
                                If iMode <> &H8 Then
                                    C = gyMem(lLocation) And 1
                                    Mem(lLocation) = gyMem(lLocation) \ 2
                                    N = 0
                                    Z = (gyMem(lLocation) = 0) * -2
                                    lCycles = lCycles + 4&
                                Else
                                    C = A And 1
                                    A = A \ 2
                                    N = 0
                                    Z = (A = 0) * -2
                                    lCycles = 2&
                                End If
                            Case &H60  ' ROR
                                lTemp = C
                                N = C * 128&
                                If iMode <> &H8 Then
                                    C = gyMem(lLocation) And 1
                                    lTemp = (gyMem(lLocation) + lTemp * 256&) \ 2
                                    Mem(lLocation) = lTemp
                                    Z = (lTemp = 0) * -2
                                    lCycles = lCycles + 4&
                                Else
                                    C = A And 1
                                    A = (A + lTemp * 256&) \ 2
                                    Z = (A = 0) * -2
                                    lCycles = 2&
                                End If
                            Case &H80  ' STX
                                Mem(lLocation) = X
                                lCycles = lCycles + 2&
                            Case &HA0  ' LDX
                                X = gyMem(lLocation)
                                N = X And 128
                                Z = (X = 0) * -2
                                lCycles = lCycles + 2&
                            Case &HC0  ' DEC
                                lTemp = (gyMem(lLocation) - 1) And 255
                                Mem(lLocation) = lTemp
                                N = lTemp And 128
                                Z = (lTemp = 0) * -2
                                lCycles = lCycles + 4&
                            Case &HE0  ' INC
                                lTemp = (gyMem(lLocation) + 1) And 255
                                Mem(lLocation) = lTemp
                                N = lTemp And 128
                                Z = (lTemp = 0) * -2
                                lCycles = lCycles + 4&
                        End Select
                    Case &H3
                        lCycles = 2&
                        'Exit Do
                End Select
        End Select
        PC = PC + 1&
        
        VideoULA.Tick lCycles
        Keyboard.Tick lCycles
        SystemVIA6522.TimersTick lCycles
        UserVIA6522.TimersTick lCycles
        ProcessInterrupt
        Throttle.ThrottleTick lCycles
    Loop Until STOPFlag
    STOPFlag = False
End Sub

Private Sub ProcessInterrupt()
    If RESFlag Then
        RESFlag = False
        RES
    ElseIf NMIFlag Then
        NMIFlag = False
        NMI
    ElseIf IRQFlag And I = 0 Then
        IRQFlag = False
        IRQ
    End If
End Sub

Private Sub IRQ()
    Dim tbTemp As TwoByte
    
    CopyMemory tbTemp, PC, &H2
    Mem(S) = tbTemp.Hi
    S = S - 1&: If S < &H100& Then S = &H1FF&
    Mem(S) = tbTemp.Lo
    S = S - 1&: If S < &H100& Then S = &H1FF&
    Mem(S) = N + V + 32 + D + I + Z + C
    S = S - 1&: If S < &H100& Then S = &H1FF&
    B = 0
    I = 4 ' disabled
    PC = gyMem(&HFFFE&) + gyMem(&HFFFF&) * 256&
End Sub

Private Sub NMI()
    Dim tbTemp As TwoByte
    
    CopyMemory tbTemp, PC, &H2
    Mem(S) = tbTemp.Hi
    S = S - 1&: If S < &H100& Then S = &H1FF&
    Mem(S) = tbTemp.Lo
    S = S - 1&: If S < &H100& Then S = &H1FF&
    Mem(S) = N + V + 32 + D + I + Z + C
    S = S - 1&: If S < &H100& Then S = &H1FF&
    PC = gyMem(&HFFFA&) + gyMem(&HFFFB&) * 256&
End Sub

Public Sub RES()
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
End Sub
