Attribute VB_Name = "Compiler6502"
Option Explicit

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Private Type MemoryLocation
    Location As Long
    ZeroPage As Boolean
End Type

Private moInstruction As ISaffronObject
Private moLabels As Dictionary

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

Public Sub LoadAssembly(ByVal sFilePath As String)
    Dim sAssembly As String
    Dim sPath As String

    sPath = sFilePath
    Close #1
    Open sPath For Binary As #1
    sAssembly = Space$(FileLen(sFilePath))
    Get #1, , sAssembly
    Close #1
    Compile sAssembly
End Sub

Public Sub LoadRoms()
    Dim lLen As Long
    Dim lMem As Long
    Dim yValue As Byte
    
'    RomSelect.LoadRom 15, "basic2.rom"
'    RomSelect.SetRom 15
    
    lLen = FileLen(App.path & "\Roms\os12.rom")
    Open App.path & "\Roms\os12.rom" For Binary As #1
    For lMem = 0 To lLen - 1
        Get #1, , yValue
        gyMem(lMem + &HC000&) = yValue
    Next
    Close #1
End Sub

Public Sub LoadRam(ByVal sFile As String)
    Dim lLen As Long
    Dim lMem As Long
    Dim yValue As Byte
    
    lLen = FileLen(App.path & "\rams\" & sFile)
    Open App.path & "\" & sFile For Binary As #1
    For lMem = 0 To lLen - 1
        Get #1, , yValue
        gyMem(lMem) = yValue
    Next
    Close #1
End Sub

Private Function Compile(sText As String) As String
    InitialiseCompiler
    
    Set moLabels = New Dictionary
    
    If CompileCodeUEF(sText, False) Then
        CompileCodeUEF sText, True
    End If
End Function

Private Sub InitialiseCompiler()
    Dim oParser As SaffronClasses.ISaffronObject
    'Dim oParser As Object
    Dim SaffronCompiler As Object
    
    Dim sRules As String
    sRules = Space$(FileLen(App.path & "\P6502Compiler.saf"))
    
    Open App.path & "\P6502Compiler.saf" For Binary As #1
    Get #1, , sRules
    Close #1
    
    If Not SaffronObject.CreateRules(sRules) Then
        MsgBox "Bad Saf"
        End
    End If
    Set moInstruction = SaffronObject.Rules("instruction")
End Sub

Private Sub ErrorMessage(sCaption As String, sMessage As String)
    MsgBox sCaption & " : " & vbCrLf & sMessage
End Sub

Private Function CompileCode(sText As String, bResolveLabels As Boolean) As Boolean
    Dim memLocation As MemoryLocation
    Dim oTree As SaffronTree
    Dim lInstruction As Long
    Dim lMode As Long
    Dim lValue As Long
    Dim bFinished As Boolean
    Dim sPreIndexRegister As String
    Dim sPostIndexRegister As String
    Dim lPosition As Long
    Dim sInstruction As String
'    Dim lMaximum As Long
    Dim bInstructionOk As Boolean
    Dim lTempMode As Long
    Dim lInstructionOffset As Long
    Dim sLiteral As String
    Dim bStringLiteral As Boolean
    Dim lIndex As Long
    Dim lChar As Long
    
    If bResolveLabels Then Open App.path & "\object.txt" For Binary As #1
    
    CompileCode = True
    
    SaffronStream.Text = sText
    If bResolveLabels Then Put #1, , vbCrLf
    
    Do
        lPosition = SaffronStream.Position
        'Set oTree = New SaffronTree
        If SaffronStream.Position > Len(sText) Then
            Exit Function
        End If
        If Not moInstruction.Parse(oTree) Then
            ErrorMessage "Error in compilation", Mid$(sText, SaffronStream.Position, 100)
            CompileCode = False
            Exit Function
        End If
        
        bStringLiteral = False
        Select Case oTree.Index
            Case 1 ' comment
                ' do nothing
            Case 2 ' instruction
                sInstruction = UCase$(oTree(1)(1).Text)
                
                lMode = amImplied
                lValue = -2
                sPreIndexRegister = ""
                sPostIndexRegister = ""
                If oTree(1)(2).Index = 1 Then
                    Select Case oTree(1)(2)(1).Index
                        Case 1 ' immediate
                            'lValue = ResolveExpression(oTree(1)(2)(1)(1)(1))
                            If bResolveLabels And lValue = -1 And sInstruction <> "DS" Then
                                ErrorMessage "Unresolved Label", Mid$(sText, lPosition, 100)
                                CompileCode = False
                                Exit Function
                            End If
                            lMode = amImmediate
                            If lValue > 255 Then
                                ErrorMessage "Number out of bounds", Mid$(sText, lPosition, 100)
                                CompileCode = False
                                Exit Function
                            End If
                        Case 2 ' indexed
                            If UCase$(oTree(1)(2)(1)(1)(1).Text) <> "A" Then
                                lValue = ResolveExpression(oTree(1)(2)(1)(1)(1)).Location
                                If bResolveLabels And lValue = -1 And sInstruction <> "DS" Then
                                    ErrorMessage "Unresolved Label", Mid$(sText, lPosition, 100)
                                    CompileCode = False
                                    Exit Function
                                End If
'                                lMaximum = MaximumNumber(oTree(1)(2)(1)(1)(1))
                                sPreIndexRegister = UCase$(oTree(1)(2)(1)(1)(2).Text)
                                Select Case sPreIndexRegister
                                    Case "X"
'                                        If lMaximum < 256 Then
'                                            lMode = amZeroPageX
'                                        Else
                                            lMode = amAbsoluteX
'                                        End If
                                    Case "Y"
'                                        If lMaximum < 256 Then
'                                            lMode = amZeroPageY
'                                        Else
                                            lMode = amAbsoluteY
'                                        End If
                                    Case Else
'                                        If lMaximum < 256 Then
'                                            lMode = amZeroPage
'                                        Else
                                            lMode = amAbsolute
'                                        End If
                                End Select
                            Else
                                lMode = amAccumulator
                            End If

                        Case 3 ' bracket
                            lValue = ResolveExpression(oTree(1)(2)(1)(1)(1)(1)).Location
                            sPreIndexRegister = UCase$(oTree(1)(2)(1)(1)(1)(2).Text)
                            sPostIndexRegister = UCase$(oTree(1)(2)(1)(1)(2).Text)

                            If sPreIndexRegister = "X" Then
                                If sPostIndexRegister = "" Then
                                    If lValue > 256 Then
                                        ErrorMessage "Number out of bounds", Mid$(sText, lPosition, 100)
                                        CompileCode = False
                                        Exit Function
                                    End If
                                    lMode = amZeroPageXIndirect
                                Else
                                    ErrorMessage "Incorrect post index register", Mid$(sText, lPosition, 100)
                                    CompileCode = False
                                    Exit Function
                                End If
                            ElseIf sPreIndexRegister = "" Then
                                If sPostIndexRegister = "Y" Then
                                    If lValue > 256 Then
                                        ErrorMessage "Number out of bounds", Mid$(sText, lPosition, 100)
                                        CompileCode = False
                                        Exit Function
                                    End If
                                    lMode = amZeroPageYIndirect
                                ElseIf sPostIndexRegister = "" Then
                                    lMode = amAbsoluteIndirect
                                Else
                                    ErrorMessage "Incorrect post index register", Mid$(sText, lPosition, 100)
                                    CompileCode = False
                                    Exit Function
                                End If
                            Else
                                ErrorMessage "Incorrect pre index register", Mid$(sText, lPosition, 100)
                                CompileCode = False
                                Exit Function
                            End If
                    End Select
                End If
                
                Select Case sInstruction
                    Case "ORA", "AND", "EOR", "ADC", "STA", "LDA", "CMP", "SBC"
                        Select Case lMode
                            Case amImplied, amAccumulator, amZeroPageY, amAbsoluteIndirect
                                ErrorMessage "Addressing mode not supported:", Mid$(sText, lPosition, 100)
                                CompileCode = False
                                Exit Function
                        End Select
                        If sInstruction = "STA" And lMode = amImmediate Then
                            ErrorMessage "Addressing mode not supported:", Mid$(sText, lPosition, 100)
                            CompileCode = False
                            Exit Function
                        End If
                        lInstruction = (oTree(1)(1)(1).Index - 1) * 32 + 1 + lMode * 4
                        
                    Case "ASL", "ROL", "LSR", "ROR", "STX", "LDX", "DEC", "INC"
                        Select Case lMode
                            Case amZeroPageXIndirect, amZeroPageYIndirect, amAbsoluteY, amAbsoluteIndirect
                                bInstructionOk = False
                            Case Else
                                bInstructionOk = True
                                Select Case sInstruction
                                    Case "STX"
                                        Select Case lMode
                                            Case amImmediate, amAccumulator, amZeroPageX, amAbsoluteX
                                                bInstructionOk = False
                                        End Select
                                    Case "LDX"
                                        Select Case lMode
                                            Case amAccumulator
                                                bInstructionOk = False
                                        End Select
                                    Case "DEC", "INC"
                                        Select Case lMode
                                            Case amImmediate, amAccumulator, amZeroPageY
                                                bInstructionOk = False
                                        End Select
                                    Case Else
                                        Select Case lMode
                                            Case amZeroPageY, amImmediate
                                                bInstructionOk = False
                                        End Select
                                End Select
                        End Select
                        
                        If Not bInstructionOk Then
                            ErrorMessage "Addressing mode not supported:", Mid$(sText, lPosition, 100)
                            CompileCode = False
                            Exit Function
                        Else
                            Select Case lMode
                                Case amImmediate
                                    lTempMode = 0
                                Case amZeroPage
                                    lTempMode = 1
                                Case amAccumulator
                                    lTempMode = 2
                                Case amAbsolute
                                    lTempMode = 3
                                Case amZeroPageX, amZeroPageY
                                    lTempMode = 5
                                Case amAbsoluteX
                                    lTempMode = 7
                            End Select
                            lInstruction = (oTree(1)(1)(1).Index - 9) * 32 + 2 + lTempMode * 4
                        End If
                                               
                    Case "BIT", "JMP", "STY", "LDY", "CPY", "CPX"
                        Select Case lMode
                            Case amZeroPageXIndirect, amZeroPageYIndirect, amAbsoluteY, amAccumulator, amZeroPageY
                                bInstructionOk = False
                            Case Else
                                bInstructionOk = True
                                lInstructionOffset = 0
                                Select Case sInstruction
                                    Case "BIT"
                                        Select Case lMode
                                            Case amImmediate, amZeroPageX, amAbsoluteX, amAbsoluteIndirect
                                                bInstructionOk = False
                                        End Select
                                    Case "JMP"
                                        Select Case lMode
                                            Case amImmediate, amZeroPageX, amAbsoluteX
                                                bInstructionOk = False
                                            Case amZeroPage
                                                lMode = amAbsolute
                                            Case amAbsoluteIndirect
                                                lInstructionOffset = 1
                                        End Select
                                    Case "STY"
                                        Select Case lMode
                                            Case amImmediate, amAbsoluteX, amAbsoluteIndirect
                                                bInstructionOk = False
                                        End Select
                                        lInstructionOffset = 1
                                    Case "LDY"
                                        lInstructionOffset = 1
                                    Case "CPY", "CPX"
                                        lInstructionOffset = 1
                                        Select Case lMode
                                            Case amZeroPageX, amAbsoluteX, amAbsoluteIndirect
                                                bInstructionOk = False
                                        End Select
                                End Select
                        End Select
                        
                        If Not bInstructionOk Then
                            ErrorMessage "Addressing mode not supported:", Mid$(sText, lPosition, 100)
                            CompileCode = False
                            Exit Function
                        Else
                            Select Case lMode
                                Case amImmediate
                                    lTempMode = 0
                                Case amZeroPage
                                    lTempMode = 1
                                Case amAbsolute, amAbsoluteIndirect
                                    lTempMode = 3
                                Case amZeroPageX
                                    lTempMode = 5
                                Case amAbsoluteX
                                    lTempMode = 7
                            End Select
                            lInstruction = (oTree(1)(1)(1).Index - 16 + lInstructionOffset) * 32 + lTempMode * 4
                        End If
                    Case "BPL", "BMI", "BVC", "BVS", "BCC", "BCS", "BNE", "BEQ"
                        Select Case lMode
                            Case amZeroPage, amAbsolute
                                lInstruction = (oTree(1)(1)(1).Index - 23) * 32 + 16
                                If bResolveLabels Then
                                    lValue = lValue - OffsetLocation(memLocation, 2).Location
                                    If lValue > 127 Or lValue < -128 Then
                                        ErrorMessage "Address out of bounds:", Mid$(sText, lPosition, 100)
                                        CompileCode = False
                                        Exit Function
                                    End If
                                End If
                                lMode = amZeroPage
                            Case Else
                                ErrorMessage "Addressing mode not supported:", Mid$(sText, lPosition, 100)
                                CompileCode = False
                                Exit Function
                        End Select
                    Case "PHP", "CLC", "PLP", "SEC", "PHA", "CLI", "PLA", "SEI", "DEY", "TYA", "TAY", "CLV", "INY", "CLD", "INX", "SED"
                        lInstruction = (oTree(1)(1)(1).Index - 31) * 16 + 8
                        If lMode <> -1 Then
                            ErrorMessage "Implied instruction:", Mid$(sText, lPosition, 100)
                            CompileCode = False
                            Exit Function
                        End If
                    Case "TXA", "TXS", "TAX", "TSX", "DEX", "NOP"
                        lInstruction = (oTree(1)(1)(1).Index - 39) * 16 + 10
                        If lMode <> -1 Then
                            ErrorMessage "Implied instruction:", Mid$(sText, lPosition, 100)
                            CompileCode = False
                            Exit Function
                        End If
                    Case "BRK", "JSR", "RTI", "RTS"
                        If sInstruction <> "JSR" Then
                            If lMode <> -1 Then
                                ErrorMessage "Implied instruction:", Mid$(sText, lPosition, 100)
                                CompileCode = False
                                Exit Function
                            End If
                        Else
                            Select Case lMode
                                Case amZeroPage
                                    lMode = amAbsolute
                                Case amAbsolute
                                Case Else
                                    ErrorMessage "Addressing mode not supported:", Mid$(sText, lPosition, 100)
                                    CompileCode = False
                                    Exit Function
                            End Select
                        End If
                        lInstruction = (oTree(1)(1)(1).Index - 53) * 32
                    Case "DB"
                        lInstruction = lValue And 255
                        lMode = -1
                    Case "DW"
                        lInstruction = lValue And 255
                        lValue = Hi(lValue)
                        lMode = 0
                    Case "DS"
                        bStringLiteral = True
                        sLiteral = Replace$(oTree(1)(2).Text, "\ ", " ")
                        If bResolveLabels Then Put #1, , HexNum(memLocation.Location, 4) & " "
                        For lIndex = 1 To Len(sLiteral)
                            lChar = Asc(Mid$(sLiteral, lIndex, 1))
                            If bResolveLabels Then Put #1, , HexNum(lChar, 2) & " "
                            'Mem(lLocation + lIndex - 1) = lChar And 255
                        Next
                        If bResolveLabels Then Put #1, , vbCrLf
                        memLocation.Location = memLocation.Location + Len(sLiteral)
                        lMode = -1
                    Case "HALT"
                        lInstruction = &H2
                End Select
                
                If Not bStringLiteral Then
                    If bResolveLabels Then Put #1, , HexNum(memLocation.Location, 4) & " "
                    
                    If bResolveLabels Then Put #1, , HexNum(lInstruction, 2) & " "
                    Select Case lMode
                        Case amImmediate, amZeroPage, amZeroPageX, amZeroPageY, amZeroPageXIndirect, amZeroPageYIndirect
                            If bResolveLabels Then Put #1, , HexNum(lValue And 255, 2) & " " & vbCrLf
                            memLocation.Location = memLocation.Location + 2
                        Case amAbsolute, amAbsoluteX, amAbsoluteY, amAbsoluteIndirect
                            If bResolveLabels Then Put #1, , HexNum(lValue And 255, 2) & " " & HexNum(Hi(lValue), 2) & " " & vbCrLf
                            memLocation.Location = memLocation.Location + 3
                        Case Else
                            If bResolveLabels Then Put #1, , vbCrLf
                            memLocation.Location = memLocation.Location + 1
                    End Select
                End If
                
            Case 3 ' header
                Dim lHeaderIndex As Long
                Dim lValue2 As Long
                
                lValue2 = ResolveExpression(oTree(1)(1), True, memLocation.Location).Location
                
                If lValue2 <> -1 Then
                    memLocation.Location = lValue2
                End If
        End Select
    Loop Until bFinished
    
    Close #1
End Function

Private Function CompileCodeUEF(sText As String, bResolveLabels As Boolean) As Boolean
    Dim memLocation As MemoryLocation
    Dim oTree As SaffronTree
    Dim lInstruction As Long
    Dim lMode As Long
    Dim memValue As MemoryLocation
    Dim bFinished As Boolean
    Dim sPreIndexRegister As String
    Dim sPostIndexRegister As String
    Dim lPosition As Long
    Dim sInstruction As String
    Dim lMaximum As Long
    Dim bInstructionOk As Boolean
    Dim lTempMode As Long
    Dim lInstructionOffset As Long
    Dim sLiteral As String
    Dim bStringLiteral As Boolean
    Dim lIndex As Long
    Dim lChar As Long
    
    Dim yBlock() As Byte
    Dim lBlockSize As Long
    Dim lBlockAddress As Long
    
    SaffronStream.Reset
    
    If bResolveLabels Then
        UEFHandler.CreateUEFFile
    End If
    
    CompileCodeUEF = True
    
    SaffronStream.Text = sText
    
    Do
        lPosition = SaffronStream.Position

        If SaffronStream.Position > Len(sText) Then
            Exit Do
        End If
        
        Set oTree = New SaffronTree
        If Not moInstruction.Parse(oTree) Then
            ErrorMessage "Error in compilation", Mid$(sText, SaffronStream.Position, 100)
            CompileCodeUEF = False
            Exit Function
        End If
        
        bStringLiteral = False
        Select Case oTree.Index
            Case 1 ' comment
                ' do nothing
            Case 2 ' instruction
                sInstruction = UCase$(oTree(1)(1).Text)
                
                lMode = amImplied
                memValue.Location = -2
                sPreIndexRegister = ""
                sPostIndexRegister = ""
                If oTree(1)(2).Index = 1 Then
                    Select Case oTree(1)(2)(1).Index
                        Case 1 ' immediate
                            memValue = ResolveExpression(oTree(1)(2)(1)(1)(1))
                            If bResolveLabels And memValue.Location = -1 And sInstruction <> "DS" Then
                                ErrorMessage "Unresolved Label", Mid$(sText, lPosition, 100)
                                CompileCodeUEF = False
                                Exit Function
                            End If
                            lMode = amImmediate
                            If memValue.Location > 255 Then
                                ErrorMessage "Number out of bounds", Mid$(sText, lPosition, 100)
                                CompileCodeUEF = False
                                Exit Function
                            End If
                        Case 2 ' indexed
                            If UCase$(oTree(1)(2)(1)(1)(1).Text) <> "A" Then
                                memValue = ResolveExpression(oTree(1)(2)(1)(1)(1))
                                If bResolveLabels And memValue.Location = -1 And sInstruction <> "DS" Then
                                    ErrorMessage "Unresolved Label", Mid$(sText, lPosition, 100)
                                    CompileCodeUEF = False
                                    Exit Function
                                End If
                                sPreIndexRegister = UCase$(oTree(1)(2)(1)(1)(2).Text)
                                Select Case sPreIndexRegister
                                    Case "X"
                                        If memValue.ZeroPage Then
                                            lMode = amZeroPageX
                                        Else
                                            lMode = amAbsoluteX
                                        End If
                                    Case "Y"
                                        If memValue.ZeroPage Then
                                            lMode = amZeroPageY
                                        Else
                                            lMode = amAbsoluteY
                                        End If
                                    Case Else
                                        If memValue.ZeroPage Then
                                            lMode = amZeroPage
                                        Else
                                            lMode = amAbsolute
                                        End If
                                End Select
                            Else
                                lMode = amAccumulator
                            End If

                        Case 3 ' bracket
                            memValue = ResolveExpression(oTree(1)(2)(1)(1)(1)(1))
                            sPreIndexRegister = UCase$(oTree(1)(2)(1)(1)(1)(2).Text)
                            sPostIndexRegister = UCase$(oTree(1)(2)(1)(1)(2).Text)

                            If sPreIndexRegister = "X" Then
                                If sPostIndexRegister = "" Then
                                    If Not memValue.ZeroPage Then
                                        ErrorMessage "Number out of bounds", Mid$(sText, lPosition, 100)
                                        CompileCodeUEF = False
                                        Exit Function
                                    End If
                                    lMode = amZeroPageXIndirect
                                Else
                                    ErrorMessage "Incorrect post index register", Mid$(sText, lPosition, 100)
                                    CompileCodeUEF = False
                                    Exit Function
                                End If
                            ElseIf sPreIndexRegister = "" Then
                                If sPostIndexRegister = "Y" Then
                                    If Not memValue.ZeroPage Then
                                        ErrorMessage "Number out of bounds", Mid$(sText, lPosition, 100)
                                        CompileCodeUEF = False
                                        Exit Function
                                    End If
                                    lMode = amZeroPageYIndirect
                                ElseIf sPostIndexRegister = "" Then
                                    lMode = amAbsoluteIndirect
                                Else
                                    ErrorMessage "Incorrect post index register", Mid$(sText, lPosition, 100)
                                    CompileCodeUEF = False
                                    Exit Function
                                End If
                            Else
                                ErrorMessage "Incorrect pre index register", Mid$(sText, lPosition, 100)
                                CompileCodeUEF = False
                                Exit Function
                            End If
                    End Select
                End If
                
                Select Case sInstruction
                    Case "ORA", "AND", "EOR", "ADC", "STA", "LDA", "CMP", "SBC"
                        Select Case lMode
                            Case amImplied, amAccumulator, amZeroPageY, amAbsoluteIndirect
                                ErrorMessage "Addressing mode not supported:", Mid$(sText, lPosition, 100)
                                CompileCodeUEF = False
                                Exit Function
                        End Select
                        If sInstruction = "STA" And lMode = amImmediate Then
                            ErrorMessage "Addressing mode not supported:", Mid$(sText, lPosition, 100)
                            CompileCodeUEF = False
                            Exit Function
                        End If
                        lInstruction = (oTree(1)(1)(1).Index - 1) * 32 + 1 + lMode * 4
                        
                    Case "ASL", "ROL", "LSR", "ROR", "STX", "LDX", "DEC", "INC"
                        Select Case lMode
                            Case amZeroPageXIndirect, amZeroPageYIndirect, amAbsoluteIndirect
                                bInstructionOk = False
                            Case Else
                                bInstructionOk = True
                                Select Case sInstruction
                                    Case "STX"
                                        Select Case lMode
                                            Case amImmediate, amAccumulator, amZeroPageX, amAbsoluteX, amAbsoluteY
                                                bInstructionOk = False
                                        End Select
                                    Case "LDX"
                                        Select Case lMode
                                            Case amAccumulator, amZeroPageX, amAbsoluteX
                                                bInstructionOk = False
                                        End Select
                                    Case "DEC", "INC"
                                        Select Case lMode
                                            Case amImmediate, amAccumulator, amZeroPageY, amAbsoluteY
                                                bInstructionOk = False
                                        End Select
                                    Case Else
                                        Select Case lMode
                                            Case amZeroPageY, amImmediate, amAbsoluteY
                                                bInstructionOk = False
                                        End Select
                                End Select
                        End Select
                        
                        If Not bInstructionOk Then
                            ErrorMessage "Addressing mode not supported:", Mid$(sText, lPosition, 100)
                            CompileCodeUEF = False
                            Exit Function
                        Else
                            Select Case lMode
                                Case amImmediate
                                    lTempMode = 0
                                Case amZeroPage
                                    lTempMode = 1
                                Case amAccumulator
                                    lTempMode = 2
                                Case amAbsolute
                                    lTempMode = 3
                                Case amZeroPageX, amZeroPageY
                                    lTempMode = 5
                                Case amAbsoluteX, amAbsoluteY
                                    lTempMode = 7
                            End Select
                            lInstruction = (oTree(1)(1)(1).Index - 9) * 32 + 2 + lTempMode * 4
                        End If
                                               
                    Case "BIT", "JMP", "STY", "LDY", "CPY", "CPX"
                        Select Case lMode
                            Case amZeroPageXIndirect, amZeroPageYIndirect, amAbsoluteY, amAccumulator, amZeroPageY
                                bInstructionOk = False
                            Case Else
                                bInstructionOk = True
                                lInstructionOffset = 0
                                Select Case sInstruction
                                    Case "BIT"
                                        Select Case lMode
                                            Case amImmediate, amZeroPageX, amAbsoluteX, amAbsoluteIndirect
                                                bInstructionOk = False
                                        End Select
                                    Case "JMP"
                                        Select Case lMode
                                            Case amImmediate, amZeroPageX, amAbsoluteX
                                                bInstructionOk = False
                                            Case amZeroPage
                                                lMode = amAbsolute
                                            Case amAbsoluteIndirect
                                                lInstructionOffset = 1
                                        End Select
                                    Case "STY"
                                        Select Case lMode
                                            Case amImmediate, amAbsoluteX, amAbsoluteIndirect
                                                bInstructionOk = False
                                        End Select
                                        lInstructionOffset = 1
                                    Case "LDY"
                                        lInstructionOffset = 1
                                    Case "CPY", "CPX"
                                        lInstructionOffset = 1
                                        Select Case lMode
                                            Case amZeroPageX, amAbsoluteX, amAbsoluteIndirect
                                                bInstructionOk = False
                                        End Select
                                End Select
                        End Select
                        
                        If Not bInstructionOk Then
                            ErrorMessage "Addressing mode not supported:", Mid$(sText, lPosition, 100)
                            CompileCodeUEF = False
                            Exit Function
                        Else
                            Select Case lMode
                                Case amImmediate
                                    lTempMode = 0
                                Case amZeroPage
                                    lTempMode = 1
                                Case amAbsolute, amAbsoluteIndirect
                                    lTempMode = 3
                                Case amZeroPageX
                                    lTempMode = 5
                                Case amAbsoluteX
                                    lTempMode = 7
                            End Select
                            lInstruction = (oTree(1)(1)(1).Index - 16 + lInstructionOffset) * 32 + lTempMode * 4
                        End If
                    Case "BPL", "BMI", "BVC", "BVS", "BCC", "BCS", "BNE", "BEQ"
                        Select Case lMode
                            Case amZeroPage, amAbsolute
                                lInstruction = (oTree(1)(1)(1).Index - 23) * 32 + 16
                                If bResolveLabels Then
                                    memValue.Location = memValue.Location - (memLocation.Location + 2)
                                    If memValue.Location > 127 Or memValue.Location < -128 Then
                                        ErrorMessage "Address out of bounds:", Mid$(sText, lPosition, 100)
                                        CompileCodeUEF = False
                                        Exit Function
                                    End If
                                End If
                                lMode = amZeroPage
                            Case Else
                                ErrorMessage "Addressing mode not supported:", Mid$(sText, lPosition, 100)
                                CompileCodeUEF = False
                                Exit Function
                        End Select
                    Case "PHP", "CLC", "PLP", "SEC", "PHA", "CLI", "PLA", "SEI", "DEY", "TYA", "TAY", "CLV", "INY", "CLD", "INX", "SED"
                        lInstruction = (oTree(1)(1)(1).Index - 31) * 16 + 8
                        If lMode <> -1 Then
                            ErrorMessage "Implied instruction:", Mid$(sText, lPosition, 100)
                            CompileCodeUEF = False
                            Exit Function
                        End If
                    Case "TXA", "TXS", "TAX", "TSX", "DEX", "NOP"
                        lInstruction = (oTree(1)(1)(1).Index - 39) * 16 + 10 + IIf(oTree(1)(1)(1).Index = 52, 16, 0)
                    
                        If lMode <> -1 Then
                            ErrorMessage "Implied instruction:", Mid$(sText, lPosition, 100)
                            CompileCodeUEF = False
                            Exit Function
                        End If
                    Case "BRK", "JSR", "RTI", "RTS"
                        If sInstruction <> "JSR" Then
                            If lMode <> -1 Then
                                ErrorMessage "Implied instruction:", Mid$(sText, lPosition, 100)
                                CompileCodeUEF = False
                                Exit Function
                            End If
                        Else
                            Select Case lMode
                                Case amZeroPage
                                    lMode = amAbsolute
                                Case amAbsolute
                                Case Else
                                    ErrorMessage "Addressing mode not supported:", Mid$(sText, lPosition, 100)
                                    CompileCodeUEF = False
                                    Exit Function
                            End Select
                        End If
                        lInstruction = (oTree(1)(1)(1).Index - 53) * 32
                    Case "DB"
                        lInstruction = memValue.Location And 255
                        lMode = -1
                    Case "DW"
                        lInstruction = memValue.Location And 255
                        memValue.Location = Hi(memValue.Location)
                        lMode = amZeroPage
                    Case "DS"
                        bStringLiteral = True
                        sLiteral = Replace$(oTree(1)(2).Text, "\ ", " ")
                        For lIndex = 1 To Len(sLiteral)
                            lChar = Asc(Mid$(sLiteral, lIndex, 1))
                            If bResolveLabels Then
                                ReDim Preserve yBlock(lBlockSize)
                                yBlock(lBlockSize) = lChar And &HFF&
                                lBlockSize = lBlockSize + 1
                            End If
                        Next
                        memLocation.Location = memLocation.Location + Len(sLiteral)
                        lMode = -1
                    Case "HALT"
                        lInstruction = &H2
                End Select
                
                If Not bStringLiteral Then
                    If bResolveLabels Then
                        ReDim Preserve yBlock(lBlockSize)
                        yBlock(lBlockSize) = lInstruction And &HFF&
                        lBlockSize = lBlockSize + 1
                    End If
                    Select Case lMode
                        Case amImmediate, amZeroPage, amZeroPageX, amZeroPageY, amZeroPageXIndirect, amZeroPageYIndirect
                            If bResolveLabels Then
                                ReDim Preserve yBlock(lBlockSize)
                                yBlock(lBlockSize) = memValue.Location And &HFF&
                                lBlockSize = lBlockSize + 1
                            End If
                            memLocation.Location = memLocation.Location + 2
                        Case amAbsolute, amAbsoluteX, amAbsoluteY, amAbsoluteIndirect
                            If bResolveLabels Then
                                ReDim Preserve yBlock(lBlockSize + 1)
                                yBlock(lBlockSize) = memValue.Location And &HFF&
                                yBlock(lBlockSize + 1) = Hi(memValue.Location)
                                lBlockSize = lBlockSize + 2
                            End If
                            memLocation.Location = memLocation.Location + 3
                        Case Else
                            memLocation.Location = memLocation.Location + 1
                    End Select
                End If
                
            Case 3 ' header
                Dim memValue2 As MemoryLocation
                
                memValue2 = ResolveExpression(oTree(1)(1), True, memLocation.Location, memLocation.ZeroPage)
                
                If memValue2.Location <> -1 Then
                    memLocation = memValue2
                End If
                
                If bResolveLabels Then
                    If lBlockSize > 0 Then
                        Snapshot.SaveMemorySegment lBlockAddress, yBlock
                        Debug.Print HexNum(lBlockAddress, 4)
                    End If
                    lBlockAddress = memLocation.Location
                    Erase yBlock
                    lBlockSize = 0
                End If
        End Select
    Loop Until bFinished
    
    If bResolveLabels Then
        If lBlockSize > 0 Then
            Snapshot.SaveMemorySegment lBlockAddress, yBlock
            Debug.Print HexNum(lBlockAddress, 4)
        End If
    End If
    
    If bResolveLabels Then
        UEFHandler.SaveUEFFile App.path & "\Object Code\object.uef"
    End If
End Function

Private Function ResolveExpression(oTree As SaffronTree, Optional bDefineLabels As Boolean, Optional lMemLocation As Long = -1, Optional bMemZeroPage As Boolean = True) As MemoryLocation
    Dim sLabel As String
    Dim lIndex As Long
    Dim lValue As Long
    Dim lTerm As Long
    Dim sOperator As String
    Dim lSize As Long
    Dim bZeroPage As Boolean
    
    bZeroPage = True
    
    lValue = -1
        
    For lIndex = 1 To oTree.SubTree.Count
        If lIndex Mod 2 = 1 Then
            lSize = Len(oTree(lIndex).Text)
            Select Case oTree(lIndex).Index
                Case 1 ' binary
                    lTerm = Base2(oTree(lIndex).Text)
                    If lSize > 8 Then
                        bZeroPage = False
                    End If
                Case 2 ' decimal
                    lTerm = Base10(oTree(lIndex).Text)
                    If lSize > 3 Then
                        bZeroPage = False
                    End If
                Case 3 ' hex
                    lTerm = Base16(oTree(lIndex).Text)
                    If lSize > 2 Then
                        bZeroPage = False
                    End If
                Case 4 ' label
                    sLabel = oTree(lIndex).Text
                    If moLabels.Exists(sLabel) Then
                        lTerm = moLabels.Item(sLabel)(0)
                        If moLabels.Item(sLabel)(1) = False Then
                            bZeroPage = False
                        End If
                    Else
                        If bDefineLabels Then
                            DefineLabel oTree(lIndex).Text, lMemLocation, bMemZeroPage
                        Else
                            ResolveExpression.Location = -1
                            Exit Function
                        End If
                        lTerm = -1
                    End If
            End Select
            Select Case sOperator
                Case "+"
                    lValue = lValue + lTerm
                Case "-"
                    lValue = lValue - lTerm
                Case "*"
                    lValue = lValue * lTerm
                Case "\"
                    lValue = lValue \ lTerm
                Case Else
                    lValue = lTerm
            End Select
        Else
            sOperator = oTree(lIndex).Text
        End If
    Next
    If lValue >= 256 Then
        bZeroPage = False
    End If
    
    ResolveExpression.Location = lValue
    ResolveExpression.ZeroPage = bZeroPage
End Function


Private Sub DefineLabel(ByVal sLabel As String, ByVal lLocation As Long, ByVal bZeroPage As Boolean)
    If moLabels.Exists(sLabel) Then
        moLabels.Remove sLabel
    End If
    moLabels.Add sLabel, Array(lLocation, bZeroPage)
End Sub

Public Function Base2(sNumber As String) As Long
    Dim lIndex As Long
    Dim sChar As String * 1
    
    For lIndex = 1 To Len(sNumber)
        sChar = Mid$(sNumber, lIndex, 1)
        Base2 = 2 * Base2 + Val(sChar)
    Next
End Function

Public Function Base10(sNumber As String) As Long
    Dim lIndex As Long
    Dim sChar As String * 1
    
    For lIndex = 1 To Len(sNumber)
        sChar = Mid$(sNumber, lIndex, 1)
        Base10 = 10 * Base10 + Val(sChar)
    Next
End Function

Public Function Base16(ByVal sNumber As String) As Long
    Dim lIndex As Long
    Dim sChar As String * 1
    Dim sDig As String
    
    sNumber = UCase$(sNumber)
    sDig = "0123456789ABCDEF"
    
    For lIndex = 1 To Len(sNumber)
        sChar = Mid$(sNumber, lIndex, 1)
        Base16 = 16 * Base16 + InStr(sDig, sChar) - 1
    Next
End Function

Public Function HexNum(ByVal lNumber As Long, ByVal iPlaces As Integer) As String
    HexNum = VBA.Conversion.Hex(lNumber)
    If Len(HexNum) <= iPlaces Then
        HexNum = String$(iPlaces - Len(HexNum), "0") & HexNum
    Else
        HexNum = Right$(HexNum, iPlaces)
    End If
End Function

Public Function Hi(lValue As Long) As Long
    CopyMemory Hi, ByVal VarPtr(lValue) + 1, 1&
End Function

Private Function OffsetLocation(memLocation As MemoryLocation, ByVal lOffset As Long) As MemoryLocation
    memLocation.Location = memLocation.Location + lOffset
End Function

Private Function AddOffsetLocation(memLocation As MemoryLocation, ByVal lOffset As Long)
    memLocation.Location = memLocation.Location + lOffset
End Function
