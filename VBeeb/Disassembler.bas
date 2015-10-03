Attribute VB_Name = "Disassembler"
Option Explicit

Private moLabels As New Dictionary

Private Enum AddressingModes
    amImplied = -1
    amZeroPageXIndirect
    amZeroPage
    amImmediate
    amAbsolute
    amZeroPageYIndirect
    amZeroPageX
    amAbsoluteY
    amAbsoluteX
    amAccumulator
    amZeroPageY
    amAbsoluteIndirect
End Enum

Private lLabelIndex As Long

Private mbDisplay As Boolean

Public Function DisassembleInstruction(lInstructionAddress As Long, yInstructions() As Byte, ByVal bResolveLabels As Boolean) As Variant
    Dim yInstruction As Byte
    Dim yType As Byte
    Dim yMode As Byte
    Dim ySwitch As Byte
    Dim yOperand1 As Byte
    Dim lOperand2 As Long
    Dim amAddressingMode As AddressingModes
    Dim sMessage As String
    Dim lMessageAddress As Long
    Dim lInstructionLength As Long
    Dim sInstructionText As String
    Dim lOperand As Long
    
    mbDisplay = bResolveLabels
    mbDisplay = True
    
    yInstruction = yInstructions(0)
    yType = (yInstruction And &HE0) \ &H20
    yMode = (yInstruction And &H1C) \ &H4
    ySwitch = yInstruction And &H3
    yOperand1 = yInstructions(1)
    lOperand2 = yInstructions(1) + yInstructions(2) * 256&
    
    lInstructionLength = 1
    Select Case yInstruction
        Case &H2, &H12, &H32, &H42, &H52, &H62, &H72, &H82, &H92, &HB2, &HC2, &HD2, &HE2, &HF2
            sInstructionText = InstructionText("???")
        Case &H10 ' BPL
            lOperand = lInstructionAddress + 2 + yOperand1 + (yOperand1 > 127) * 256&
            sInstructionText = InstructionText("BPL", amAbsolute, lOperand)
            lInstructionLength = 2
        Case &H30 ' BMI
            lOperand = lInstructionAddress + 2 + yOperand1 + (yOperand1 > 127) * 256&
            sInstructionText = InstructionText("BMI", amAbsolute, lOperand)
            lInstructionLength = 2
        Case &H50 ' BVC
            lOperand = lInstructionAddress + 2 + yOperand1 + (yOperand1 > 127) * 256&
            sInstructionText = InstructionText("BVC", amAbsolute, lOperand)
            lInstructionLength = 2
        Case &H70 ' BVS
            lOperand = lInstructionAddress + 2 + yOperand1 + (yOperand1 > 127) * 256&
            sInstructionText = InstructionText("BVS", amAbsolute, lOperand)
            lInstructionLength = 2
        Case &H90 ' BCC
            lOperand = lInstructionAddress + 2 + yOperand1 + (yOperand1 > 127) * 256&
            sInstructionText = InstructionText("BCC", amAbsolute, lOperand)
            lInstructionLength = 2
        Case &HB0 ' BCS
            lOperand = lInstructionAddress + 2 + yOperand1 + (yOperand1 > 127) * 256&
            sInstructionText = InstructionText("BCS", amAbsolute, lOperand)
            lInstructionLength = 2
        Case &HD0 ' BNE
            lOperand = lInstructionAddress + 2 + yOperand1 + (yOperand1 > 127) * 256&
            sInstructionText = InstructionText("BNE", amAbsolute, lOperand)
            lInstructionLength = 2
        Case &HF0 ' BEQ
            lOperand = lInstructionAddress + 2 + yOperand1 + (yOperand1 > 127) * 256&
            sInstructionText = InstructionText("BEQ", amAbsolute, lOperand)
            lInstructionLength = 2
        Case &H8  ' PHP
            sInstructionText = InstructionText("PHP")
        Case &H28 ' PLP
            sInstructionText = InstructionText("PLP")
        Case &H48 ' PHA
            sInstructionText = InstructionText("PHA")
        Case &H68 ' PLA
            sInstructionText = InstructionText("PLA")
        Case &H88 ' DEY
            sInstructionText = InstructionText("DEY")
        Case &HA8 ' TAY
            sInstructionText = InstructionText("TAY")
        Case &HC8 ' INY
            sInstructionText = InstructionText("INY")
        Case &HE8 ' INX
            sInstructionText = InstructionText("INX")
        Case &H18 ' CLC ' 2.514 us
            sInstructionText = InstructionText("CLC")
        Case &H38 ' SEC
            sInstructionText = InstructionText("SEC")
        Case &H58 ' CLI
            sInstructionText = InstructionText("CLI")
        Case &H78 ' SEI
            sInstructionText = InstructionText("SEI")
        Case &H98 ' TYA
            sInstructionText = InstructionText("TYA")
        Case &HB8 ' CLV
            sInstructionText = InstructionText("CLV")
        Case &HD8 ' CLD
            sInstructionText = InstructionText("CLD")
        Case &HF8 ' SED
            sInstructionText = InstructionText("SED")
        Case &H8A ' TXA
            sInstructionText = InstructionText("TXA")
        Case &H9A ' TXS
            sInstructionText = InstructionText("TXS")
        Case &HAA ' TAX
            sInstructionText = InstructionText("TAX")
        Case &HBA ' TSX
            sInstructionText = InstructionText("TSX")
        Case &HCA ' DEX
            sInstructionText = InstructionText("DEX")
        Case &HEA ' NOP
            sInstructionText = InstructionText("NOP")
        Case &H0 ' BRK
            sInstructionText = InstructionText("BRK")
        Case &H20 ' JSR abs
            sInstructionText = InstructionText("JSR", amAbsolute, lOperand2)
            lInstructionLength = 3
        Case &H40 ' RTI
            sInstructionText = InstructionText("RTI")
        Case &H60 ' RTS
            sInstructionText = InstructionText("RTS")
        Case Else
            Select Case ySwitch
                Case &H0
                    Select Case yMode
                        Case 0  ' #immediate
                            amAddressingMode = amImmediate
                            lInstructionLength = 2
                            lOperand = yOperand1
                        Case 1  ' zero page
                            amAddressingMode = amZeroPage
                            lInstructionLength = 2
                            lOperand = yOperand1
                        Case 3  ' absolute
                            amAddressingMode = amAbsolute
                            lInstructionLength = 3
                            lOperand = lOperand2
                        Case 5  ' zero page,X
                            amAddressingMode = amZeroPageX
                            lInstructionLength = 2
                            lOperand = yOperand1
                        Case 7  ' absolute,X
                            amAddressingMode = amAbsoluteX
                            lInstructionLength = 3
                            lOperand = lOperand2
                    End Select
                    Select Case yType
                        Case 1  ' BIT
                            sInstructionText = InstructionText("BIT", amAddressingMode, lOperand2)
                        Case 2  ' JMP
                            sInstructionText = InstructionText("JMP", amAddressingMode, lOperand2)
                        Case 3  ' JMP (abs)
                            sInstructionText = InstructionText("JMP", amAbsoluteIndirect, lOperand2)
                        Case 4  ' STY
                            sInstructionText = InstructionText("STY", amAddressingMode, lOperand2)
                        Case 5  ' LDY
                            sInstructionText = InstructionText("LDY", amAddressingMode, lOperand2)
                        Case 6  ' CPY
                            sInstructionText = InstructionText("CPY", amAddressingMode, lOperand2)
                        Case 7  ' CPX
                            sInstructionText = InstructionText("CPX", amAddressingMode, lOperand2)
                    End Select
                    
                Case &H1
                    Select Case yMode
                        Case 0  ' (zero page,X)
                            amAddressingMode = amZeroPageXIndirect
                            lInstructionLength = 2
                            lOperand = yOperand1
                        Case 1  ' zero page
                            amAddressingMode = amZeroPage
                            lInstructionLength = 2
                            lOperand = yOperand1
                        Case 2  ' #immediate
                            amAddressingMode = amImmediate
                            lInstructionLength = 2
                            lOperand = yOperand1
                        Case 3  ' absolute
                            amAddressingMode = amAbsolute
                            lInstructionLength = 3
                            lOperand = lOperand2
                        Case 4  ' (zero page),Y
                            amAddressingMode = amZeroPageYIndirect
                            lInstructionLength = 2
                            lOperand = yOperand1
                        Case 5  ' zero page,X
                            amAddressingMode = amZeroPageX
                            lInstructionLength = 2
                            lOperand = yOperand1
                        Case 6  ' absolute,Y
                            amAddressingMode = amAbsoluteY
                            lInstructionLength = 3
                            lOperand = lOperand2
                        Case 7  ' absolute,X
                            amAddressingMode = amAbsoluteX
                            lInstructionLength = 3
                            lOperand = lOperand2
                    End Select
                    Select Case yType
                        Case 0 ' ORA
                            sInstructionText = InstructionText("ORA", amAddressingMode, lOperand2)
                        Case 1 ' AND
                            sInstructionText = InstructionText("AND", amAddressingMode, lOperand2)
                        Case 2 ' EOR
                            sInstructionText = InstructionText("EOR", amAddressingMode, lOperand2)
                        Case 3 ' ADC
                            sInstructionText = InstructionText("ADC", amAddressingMode, lOperand2)
                        Case 4 ' STA
                            sInstructionText = InstructionText("STA", amAddressingMode, lOperand2)
                        Case 5 ' LDA
                            sInstructionText = InstructionText("LDA", amAddressingMode, lOperand2)
                        Case 6 ' CMP
                            sInstructionText = InstructionText("CMP", amAddressingMode, lOperand2)
                        Case 7 ' SBC
                            sInstructionText = InstructionText("SBC", amAddressingMode, lOperand2)
                    End Select
                Case &H2
                    Select Case yMode
                        Case 0  ' #immediate
                            amAddressingMode = amImmediate
                            lInstructionLength = 2
                            lOperand = yOperand1
                        Case 1  ' zero page
                            amAddressingMode = amZeroPage
                            lInstructionLength = 2
                            lOperand = yOperand1
                        Case 2  ' accumulator
                            amAddressingMode = amAccumulator
                            lInstructionLength = 1
                        Case 3  ' absolute
                            amAddressingMode = amAbsolute
                            lInstructionLength = 3
                            lOperand = lOperand2
                        Case 5  ' zero page,X / zero page,Y
                            If yType <> 4 Then  'STX
                                amAddressingMode = amZeroPageX
                            Else
                                amAddressingMode = amZeroPageY
                            End If
                            lInstructionLength = 2
                            lOperand = yOperand1
                        Case 7  ' absolute,X
                            amAddressingMode = amAbsoluteX
                            lInstructionLength = 3
                            lOperand = lOperand2
                    End Select
                    Select Case yType
                        Case 0  ' ASL
                            sInstructionText = InstructionText("ASL", amAddressingMode, lOperand2)
                        Case 1  ' ROL
                            sInstructionText = InstructionText("ROL", amAddressingMode, lOperand2)
                        Case 2  ' LSR
                            sInstructionText = InstructionText("LSR", amAddressingMode, lOperand2)
                        Case 3  ' ROR
                            sInstructionText = InstructionText("ROR", amAddressingMode, lOperand2)
                        Case 4  ' STX
                            sInstructionText = InstructionText("STX", amAddressingMode, lOperand2)
                        Case 5  ' LDX
                            sInstructionText = InstructionText("LDX", amAddressingMode, lOperand2)
                        Case 6  ' DEC
                            sInstructionText = InstructionText("DEC", amAddressingMode, lOperand2)
                        Case 7  ' INC
                            sInstructionText = InstructionText("INC", amAddressingMode, lOperand2)
                    End Select
                Case &H3
                    sInstructionText = InstructionText("???", amImplied)
            End Select
    End Select
    DisassembleInstruction = Array(sInstructionText, lInstructionLength, lOperand)
End Function

Private Function InstructionText(sInstruction As String, Optional ByVal amAddressingMode As AddressingModes = amImplied, Optional lOperand As Long) As String
    Dim sInstructionLine As String
    
    'sInstructionLine = ShowLabelAtAddress(lInstructionAddress)
    If mbDisplay Then
        If sInstructionLine <> "" Then
            InstructionText = sInstructionLine
        End If
    End If
    sInstructionLine = sInstruction & " "
    sInstructionLine = sInstructionLine & ShowAddressingMode(amAddressingMode, lOperand)
    
    If mbDisplay Then
        InstructionText = sInstructionLine
    End If
End Function

Private Function ShowAddressingMode(amType As AddressingModes, lOperand As Long) As String
    Dim sLabel As String
    Select Case amType
        Case amImplied
            ShowAddressingMode = ""
        Case amZeroPageXIndirect
            sLabel = GetLabel(lOperand And 255, 1)
            ShowAddressingMode = "(" & sLabel & ",X)"
        Case amZeroPage
            sLabel = GetLabel(lOperand And 255, 1)
            ShowAddressingMode = sLabel
        Case amImmediate
            ShowAddressingMode = "#" & HexNum$(lOperand And 255, 2) & "h"
        Case amAbsolute
            sLabel = GetLabel(lOperand)
            ShowAddressingMode = sLabel
        Case amZeroPageYIndirect
            sLabel = GetLabel(lOperand And 255, 1)
            ShowAddressingMode = "(" & sLabel & "),Y"
        Case amZeroPageX
            sLabel = GetLabel(lOperand And 255, 1)
            ShowAddressingMode = sLabel & ",X"
        Case amAbsoluteY
            sLabel = GetLabel(lOperand)
            ShowAddressingMode = sLabel & ",Y"
        Case amAbsoluteX
            sLabel = GetLabel(lOperand)
            ShowAddressingMode = sLabel & ",X"
        Case amAccumulator
            ShowAddressingMode = "A"
        Case amZeroPageY
            sLabel = GetLabel(lOperand And 255, 1)
            ShowAddressingMode = sLabel & ",Y"
        Case amAbsoluteIndirect
            sLabel = GetLabel(lOperand)
            ShowAddressingMode = "(" & sLabel & ")"
    End Select
End Function

Private Function GetLabel(ByVal lAddress As Long, Optional ByVal lSize As Long = 2) As String
    Dim sAddress As String
    
    sAddress = HexNum$(lAddress, lSize * 2)
    If Not moLabels.Exists(sAddress) Then
        GetLabel = sAddress & "h"
        moLabels.Add sAddress, GetLabel
    Else
        GetLabel = moLabels.Item(sAddress)
    End If
End Function

Private Function NewLabel() As String
    NewLabel = "label" & lLabelIndex
    lLabelIndex = lLabelIndex + 1
End Function

'Public Function Disassemble(ByVal lStartAddress As Long, ByVal lEndAddress As Long) As String
'    mvLines = Array()
'    DisassembleCode lStartAddress, lEndAddress, False
'    ShowLabelsOutsideRange lStartAddress, lEndAddress
'    DisassembleCode lStartAddress, lEndAddress, True
'    Disassemble = Join(mvLines, vbCrLf)
'End Function

'Private Function ShowLabelsOutsideRange(ByVal lStartAddress As Long, ByVal lEndAddress As Long)
'    Dim vLabel As Variant
'    Dim vKey As Variant
'    Dim lIndex As Long
'    Dim lKeyAddress As Long
'
'    For lIndex = 0 To moLabels.Count - 1
'        vLabel = moLabels.Items(lIndex)
'        vKey = moLabels.Keys(lIndex)
'        lKeyAddress = Base16(vKey)
'        If lKeyAddress < lStartAddress Or lKeyAddress > lEndAddress Then
'            AddLine vKey & " " & vLabel
'        End If
'    Next
'End Function

Private Function ShowLabelAtAddress(ByVal lAddress As Long) As String
    Dim sKey As String
    
    sKey = HexNum$(lAddress, 4)
    If moLabels.Exists(sKey) Then
        ShowLabelAtAddress = sKey & " " & moLabels(sKey)
    End If
End Function

