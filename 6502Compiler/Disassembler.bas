Attribute VB_Name = "Disassembler"
Option Explicit

Private mlInstructionAddress As Long
Private mlNewInstructionAddress As Long

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
Private mvLines As Variant

Private Sub AddLine(ByVal sLine As String)
    Dim lUbound As Long
    
    lUbound = UBound(mvLines) + 1
    ReDim Preserve mvLines(lUbound)
    mvLines(lUbound) = sLine
End Sub

Private Function ShowInstruction(sInstruction As String, Optional ByVal amAddressingMode As AddressingModes = amImplied, Optional lOperand As Long)
    Dim sInstructionLine As String
    
    sInstructionLine = ShowLabelAtAddress(mlInstructionAddress)
    If mbDisplay Then
        If sInstructionLine <> "" Then
            AddLine sInstructionLine
        End If
    End If
    sInstructionLine = HexNum$(mlInstructionAddress, 4) & " " & sInstruction & " "
    sInstructionLine = sInstructionLine & ShowAddressingMode(amAddressingMode, lOperand)
    
    If mbDisplay Then
        AddLine sInstructionLine
    End If
End Function

Private Function ShowAddressingMode(amType As AddressingModes, lOperand As Long) As String
    Dim sLabel As String
    Select Case amType
        Case amImplied
            ShowAddressingMode = ""
        Case amZeroPageXIndirect
            sLabel = GetLabel(lOperand And 255)
            ShowAddressingMode = "(" & sLabel & ",X)"
        Case amZeroPage
            sLabel = GetLabel(lOperand And 255)
            ShowAddressingMode = sLabel
        Case amImmediate
            ShowAddressingMode = "#" & HexNum$(lOperand And 255, 2) & "h"
        Case amAbsolute
            sLabel = GetLabel(lOperand)
            ShowAddressingMode = sLabel
        Case amZeroPageYIndirect
            sLabel = GetLabel(lOperand And 255)
            ShowAddressingMode = "(" & sLabel & "),Y"
        Case amZeroPageX
            sLabel = GetLabel(lOperand And 255)
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
            sLabel = GetLabel(lOperand And 255)
            ShowAddressingMode = sLabel & ",Y"
        Case amAbsoluteIndirect
            sLabel = GetLabel(lOperand)
            ShowAddressingMode = "(" & sLabel & ")"
    End Select
End Function

Private Function GetLabel(ByVal lAddress As Long) As String
    Dim sAddress As String
    
    sAddress = HexNum$(lAddress, 4)
    If Not moLabels.Exists(sAddress) Then
        GetLabel = "label_" & sAddress
        moLabels.Add sAddress, GetLabel
    Else
        GetLabel = moLabels.Item(sAddress)
    End If
End Function

Private Function NewLabel() As String
    NewLabel = "label" & lLabelIndex
    lLabelIndex = lLabelIndex + 1
End Function

Public Function Disassemble(ByVal lStartAddress As Long, ByVal lEndAddress As Long) As String
    mvLines = Array()
    DisassembleCode lStartAddress, lEndAddress, False
    ShowLabelsOutsideRange lStartAddress, lEndAddress
    DisassembleCode lStartAddress, lEndAddress, True
    Disassemble = Join(mvLines, vbCrLf)
End Function

Private Function ShowLabelsOutsideRange(ByVal lStartAddress As Long, ByVal lEndAddress As Long)
    Dim vLabel As Variant
    Dim vKey As Variant
    Dim lIndex As Long
    Dim lKeyAddress As Long
    
    For lIndex = 0 To moLabels.Count - 1
        vLabel = moLabels.Items(lIndex)
        vKey = moLabels.Keys(lIndex)
        lKeyAddress = Base16(vKey)
        If lKeyAddress < lStartAddress Or lKeyAddress > lEndAddress Then
            AddLine vKey & " " & vLabel
        End If
    Next
End Function

Private Function ShowLabelAtAddress(ByVal lAddress As Long) As String
    Dim sKey As String
    
    sKey = HexNum$(lAddress, 4)
    If moLabels.Exists(sKey) Then
        ShowLabelAtAddress = sKey & " " & moLabels(sKey)
    End If
End Function

Public Sub DisassembleCode(ByVal lStartAddress As Long, ByVal lEndAddress As Long, ByVal bResolveLabels As Boolean)
    Dim iInstruction As Byte
    Dim iType As Byte
    Dim iMode As Byte
    Dim iSwitch As Byte
    Dim yOperand1 As Byte
    Dim yOperand2 As Long
    Dim amAddressingMode As AddressingModes
    Dim sMessage As String
    Dim lMessageAddress As Long
    
    mbDisplay = bResolveLabels
    
    mlInstructionAddress = lStartAddress
    
    Do
        iInstruction = gyMem(mlInstructionAddress)
        iType = (iInstruction And &HE0) \ &H20
        iMode = (iInstruction And &H1C) \ &H4
        iSwitch = iInstruction And &H3
        yOperand1 = gyMem(mlInstructionAddress + 1)
        yOperand2 = gyMem(mlInstructionAddress + 1) + gyMem((mlInstructionAddress + 2) And &HFFFF&) * 256&
        
        mlNewInstructionAddress = mlInstructionAddress + 1
        Select Case iInstruction
            Case &H2, &H12, &H32, &H42, &H52, &H62, &H72, &H82, &H92, &HB2, &HC2, &HD2, &HE2, &HF2
                ShowInstruction "???"
            Case &H10 ' BPL
                ShowInstruction "BPL", amAbsolute, mlInstructionAddress + 2 + yOperand1 + (yOperand1 > 127) * 256&
                mlNewInstructionAddress = mlInstructionAddress + 2
            Case &H30 ' BMI
                ShowInstruction "BMI", amAbsolute, mlInstructionAddress + 2 + yOperand1 + (yOperand1 > 127) * 256&
                mlNewInstructionAddress = mlInstructionAddress + 2
            Case &H50 ' BVC
                ShowInstruction "BVC", amAbsolute, mlInstructionAddress + 2 + yOperand1 + (yOperand1 > 127) * 256&
                mlNewInstructionAddress = mlInstructionAddress + 2
            Case &H70 ' BVS
                ShowInstruction "BVS", amAbsolute, mlInstructionAddress + 2 + yOperand1 + (yOperand1 > 127) * 256&
                mlNewInstructionAddress = mlInstructionAddress + 2
            Case &H90 ' BCC
                ShowInstruction "BCC", amAbsolute, mlInstructionAddress + 2 + yOperand1 + (yOperand1 > 127) * 256&
                mlNewInstructionAddress = mlInstructionAddress + 2
            Case &HB0 ' BCS
                ShowInstruction "BCS", amAbsolute, mlInstructionAddress + 2 + yOperand1 + (yOperand1 > 127) * 256&
                mlNewInstructionAddress = mlInstructionAddress + 2
            Case &HD0 ' BNE
                ShowInstruction "BNE", amAbsolute, mlInstructionAddress + 2 + yOperand1 + (yOperand1 > 127) * 256&
                mlNewInstructionAddress = mlInstructionAddress + 2
            Case &HF0 ' BEQ
                ShowInstruction "BEQ", amAbsolute, mlInstructionAddress + 2 + yOperand1 + (yOperand1 > 127) * 256&
                mlNewInstructionAddress = mlInstructionAddress + 2
            Case &H8  ' PHP
                ShowInstruction "PHP"
            Case &H28 ' PLP
                ShowInstruction "PLP"
            Case &H48 ' PHA
                ShowInstruction "PHA"
            Case &H68 ' PLA
                ShowInstruction "PLA"
            Case &H88 ' DEY
                ShowInstruction "DEY"
            Case &HA8 ' TAY
                ShowInstruction "TAY"
            Case &HC8 ' INY
                ShowInstruction "INY"
            Case &HE8 ' INX
                ShowInstruction "INX"
            Case &H18 ' CLC ' 2.514 us
                ShowInstruction "CLC"
            Case &H38 ' SEC
                ShowInstruction "SEC"
            Case &H58 ' CLI
                ShowInstruction "CLI"
            Case &H78 ' SEI
                ShowInstruction "SEI"
            Case &H98 ' TYA
                ShowInstruction "TYA"
            Case &HB8 ' CLV
                ShowInstruction "CLV"
            Case &HD8 ' CLD
                ShowInstruction "CLD"
            Case &HF8 ' SED
                ShowInstruction "SED"
            Case &H8A ' TXA
                ShowInstruction "TXA"
            Case &H9A ' TXS
                ShowInstruction "TXS"
            Case &HAA ' TAX
                ShowInstruction "TAX"
            Case &HBA ' TSX
                ShowInstruction "TSX"
            Case &HCA ' DEX
                ShowInstruction "DEX"
            Case &HEA ' NOP
                ShowInstruction "NOP"
            Case &H0 ' BRK
                ShowInstruction "BRK"
'                mlInstructionAddress = mlInstructionAddress + 1
'                If mbDisplay Then AddLine HexNum$(mlInstructionAddress, 4) & " DB " & HexNum$(gyMem(mlInstructionAddress + 1), 2) & "h"
'                mlInstructionAddress = mlInstructionAddress + 1
'                sMessage = ""
'                lMessageAddress = mlInstructionAddress
'                Do
'                    sMessage = sMessage & Chr$(gyMem(lMessageAddress))
'                    lMessageAddress = lMessageAddress + 1
'                Loop Until gyMem(lMessageAddress) = 0
'
'                If mbDisplay Then AddLine HexNum$(mlInstructionAddress, 4) & " DS " & sMessage
'                If mbDisplay Then AddLine HexNum$(lMessageAddress, 4) & " DB 00h"
'                mlNewInstructionAddress = lMessageAddress + 1
            Case &H20 ' JSR abs
                ShowInstruction "JSR", amAbsolute, yOperand2
                mlNewInstructionAddress = mlInstructionAddress + 3
            Case &H40 ' RTI
                ShowInstruction "RTI"
            Case &H60 ' RTS
                ShowInstruction "RTS"
            Case Else
                Select Case iSwitch
                    Case &H0
                        Select Case iMode
                            Case 0  ' #immediate
                                amAddressingMode = amImmediate
                                mlNewInstructionAddress = mlInstructionAddress + 2
                            Case 1  ' zero page
                                amAddressingMode = amZeroPage
                                mlNewInstructionAddress = mlInstructionAddress + 2
                            Case 3  ' absolute
                                amAddressingMode = amAbsolute
                                mlNewInstructionAddress = mlInstructionAddress + 3
                            Case 5  ' zero page,X
                                amAddressingMode = amZeroPageX
                                mlNewInstructionAddress = mlInstructionAddress + 2
                            Case 7  ' absolute,X
                                amAddressingMode = amAbsoluteX
                                mlNewInstructionAddress = mlInstructionAddress + 3
                        End Select
                        Select Case iType
                            Case 1  ' BIT
                                ShowInstruction "BIT", amAddressingMode, yOperand2
                            Case 2  ' JMP
                                ShowInstruction "JMP", amAddressingMode, yOperand2
                            Case 3  ' JMP (abs)
                                ShowInstruction "JMP", amAbsoluteIndirect, yOperand2
                            Case 4  ' STY
                                ShowInstruction "STY", amAddressingMode, yOperand2
                            Case 5  ' LDY
                                ShowInstruction "LDY", amAddressingMode, yOperand2
                            Case 6  ' CPY
                                ShowInstruction "CPY", amAddressingMode, yOperand2
                            Case 7  ' CPX
                                ShowInstruction "CPX", amAddressingMode, yOperand2
                        End Select
                        
                    Case &H1
                        Select Case iMode
                            Case 0  ' (zero page,X)
                                amAddressingMode = amZeroPageXIndirect
                                mlNewInstructionAddress = mlInstructionAddress + 2
                            Case 1  ' zero page
                                amAddressingMode = amZeroPage
                                mlNewInstructionAddress = mlInstructionAddress + 2
                            Case 2  ' #immediate
                                amAddressingMode = amImmediate
                                mlNewInstructionAddress = mlInstructionAddress + 2
                            Case 3  ' absolute
                                amAddressingMode = amAbsolute
                                mlNewInstructionAddress = mlInstructionAddress + 3
                            Case 4  ' (zero page),Y
                                amAddressingMode = amZeroPageYIndirect
                                mlNewInstructionAddress = mlInstructionAddress + 2
                            Case 5  ' zero page,X
                                amAddressingMode = amZeroPageX
                                mlNewInstructionAddress = mlInstructionAddress + 2
                            Case 6  ' absolute,Y
                                amAddressingMode = amAbsoluteY
                                mlNewInstructionAddress = mlInstructionAddress + 3
                            Case 7  ' absolute,X
                                amAddressingMode = amAbsoluteX
                                mlNewInstructionAddress = mlInstructionAddress + 3
                        End Select
                        Select Case iType
                            Case 0 ' ORA
                                ShowInstruction "ORA", amAddressingMode, yOperand2
                            Case 1 ' AND
                                ShowInstruction "AND", amAddressingMode, yOperand2
                            Case 2 ' EOR
                                ShowInstruction "EOR", amAddressingMode, yOperand2
                            Case 3 ' ADC
                                ShowInstruction "ADC", amAddressingMode, yOperand2
                            Case 4 ' STA
                                ShowInstruction "STA", amAddressingMode, yOperand2
                            Case 5 ' LDA
                                ShowInstruction "LDA", amAddressingMode, yOperand2
                            Case 6 ' CMP
                                ShowInstruction "CMP", amAddressingMode, yOperand2
                            Case 7 ' SBC
                                ShowInstruction "SBC", amAddressingMode, yOperand2
                        End Select
                    Case &H2
                        Select Case iMode
                            Case 0  ' #immediate
                                amAddressingMode = amImmediate
                                mlNewInstructionAddress = mlInstructionAddress + 2
                            Case 1  ' zero page
                                amAddressingMode = amZeroPage
                                mlNewInstructionAddress = mlInstructionAddress + 2
                            Case 2  ' accumulator
                                amAddressingMode = amAccumulator
                                mlNewInstructionAddress = mlInstructionAddress + 1
                            Case 3  ' absolute
                                amAddressingMode = amAbsolute
                                mlNewInstructionAddress = mlInstructionAddress + 3
                            Case 5  ' zero page,X / zero page,Y
                                If iType <> 4 Then  'STX
                                    amAddressingMode = amZeroPageX
                                Else
                                    amAddressingMode = amZeroPageY
                                End If
                                mlNewInstructionAddress = mlInstructionAddress + 2
                            Case 7  ' absolute,X
                                amAddressingMode = amAbsoluteX
                                mlNewInstructionAddress = mlInstructionAddress + 3
                        End Select
                        Select Case iType
                            Case 0  ' ASL
                                ShowInstruction "ASL", amAddressingMode, yOperand2
                            Case 1  ' ROL
                                ShowInstruction "ROL", amAddressingMode, yOperand2
                            Case 2  ' LSR
                                ShowInstruction "LSR", amAddressingMode, yOperand2
                            Case 3  ' ROR
                                ShowInstruction "ROR", amAddressingMode, yOperand2
                            Case 4  ' STX
                                ShowInstruction "STX", amAddressingMode, yOperand2
                            Case 5  ' LDX
                                ShowInstruction "LDX", amAddressingMode, yOperand2
                            Case 6  ' DEC
                                ShowInstruction "DEC", amAddressingMode, yOperand2
                            Case 7  ' INC
                                ShowInstruction "INC", amAddressingMode, yOperand2
                        End Select
                    Case &H3
                        ShowInstruction "???", amImplied
                End Select
        End Select
        mlInstructionAddress = mlNewInstructionAddress
    Loop Until mlInstructionAddress >= lEndAddress
End Sub

Public Function MemoryDump(ByVal lStartAddress As Long, ByVal lEndAddress As Long) As String
    Dim lAddress As Long
    
    For lAddress = lStartAddress To lEndAddress
        Debug.Print " " & HexNum$(lAddress, 4) & "    " & HexNum$(gyMem(lAddress), 2) & " " & HexNum$(gyMem(lAddress + 1), 2) & HexNum$(gyMem(lAddress), 2) & " " & Chr$(gyMem(lAddress))
    Next
    
End Function
