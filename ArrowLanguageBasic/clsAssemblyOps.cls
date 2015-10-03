VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAssemblyOps"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Count As Long
Private moAssemblyOps() As clsAssemblyOp

Private mvOutput As Variant

Public Function AddAssemblyOp(oMember As clsAssemblyOp)
    ReDim Preserve moAssemblyOps(Count)
    Set moAssemblyOps(Count) = oMember
    Count = Count + 1
End Function

Public Property Get AssemblyOps(ByVal lIndex As Long) As clsAssemblyOp
Attribute AssemblyOps.VB_UserMemId = 0
    Set AssemblyOps = moAssemblyOps(lIndex)
End Property

Public Property Set AssemblyOps(ByVal lIndex As Long, oMember As clsAssemblyOp)
    Set moAssemblyOps(lIndex) = oMember
End Property

Private Sub AddLine(sLine As String)
    ReDim Preserve mvOutput(UBound(mvOutput) + 1)
    mvOutput(UBound(mvOutput)) = sLine
End Sub

Private Function Offset(ByVal sIdentifier, ByVal lOffset As Long)
    If lOffset > 0 Then
        Offset = sIdentifier & "+" & lOffset & "d"
    Else
        Offset = sIdentifier
    End If
End Function


Private Function ExGenIdentOffset(ByVal sInstruction As String, ByVal sIdentifier As Variant, ByVal lSize As Long, ByVal lActualSize As Long, ByVal lOffset As Long, Optional ByVal bReverse As Boolean = False)
    Dim lIndex As Long
    
    If Not bReverse Then
        If lOffset <= (lSize - 1) Then
            AddLine sInstruction & " " & Offset(sIdentifier, lOffset)
        ElseIf lOffset <= (lActualSize - 1) Then
            AddLine sInstruction & " #00h"
        End If
    Else
        If (lActualSize - lOffset - 1) <= (lSize - 1) Then
            AddLine sInstruction & " " & Offset(sIdentifier, lActualSize - lOffset - 1)
        Else
            AddLine sInstruction & " #00h"
        End If
    End If
End Function

Private Function ExGenConstOffset(ByVal sInstruction As String, ByVal lConstant As Long, ByVal lSize As Long, ByVal lActualSize As Long, ByVal lOffset As Long, Optional ByVal bReverse As Boolean = False)
    Dim lIndex As Long
    
    If Not bReverse Then
        For lIndex = 1 To lOffset
            lConstant = lConstant \ 256
        Next
        If lOffset <= (lSize - 1) Then
            AddLine sInstruction & " #" & HexNum(lConstant And &HFF, 2)
        ElseIf lOffset <= (lActualSize - 1) Then
            AddLine sInstruction & " #00h"
        End If
    Else
        For lIndex = 1 To (lActualSize - lOffset - 1)
            lConstant = lConstant \ 256
        Next
        If (lActualSize - lOffset - 1) <= (lSize - 1) Then
            AddLine sInstruction & " #" & HexNum(lConstant And &HFF, 2)
        Else
            AddLine sInstruction & " #00h"
        End If
    End If
End Function

Private Function ExGenIdent(ByVal sInstruction As String, ByVal sIdentifier, ByVal lSize As Long, ByVal lActualSize As Long, Optional ByVal bReverse As Boolean = False)
    Dim lIndex As Long
    
    If Not bReverse Then
        For lIndex = 0 To lActualSize - 1
            If lIndex <= lSize Then
                AddLine sInstruction & " " & Offset(sIdentifier, lIndex)
            Else
                AddLine sInstruction & " #00h"
            End If
        Next
    Else
        For lIndex = lActualSize - 1 To 0 Step -1
            If lIndex <= lSize Then
                AddLine sInstruction & " " & Offset(sIdentifier, lIndex)
            Else
                AddLine sInstruction & " #00h"
            End If
        Next
    End If
End Function

Private Sub Optimise()
    Dim lIndex As Long
    Dim lIndex2 As Long
    Dim oAssemblyOp(2) As clsAssemblyOp
Exit Sub
    For lIndex = 0 To Count - 1
        For lIndex2 = 0 To UBound(oAssemblyOp)
            If (lIndex + lIndex2) < Count Then
                Set oAssemblyOp(lIndex2) = moAssemblyOps(lIndex + lIndex2)
            Else
                Set oAssemblyOp(lIndex2) = Nothing
            End If
        Next
        
        Select Case oAssemblyOp(0).Instruction
            Case ioBRC, ioBRS
                Select Case oAssemblyOp(1).Instruction
                    Case ioLabel
                        If oAssemblyOp(0).Operand = oAssemblyOp(1).Operand Then
                            oAssemblyOp(0).Instruction = ioNOP
                        End If
                End Select
            Case ioST
                Select Case oAssemblyOp(1).Instruction
                    Case ioLD
                        If oAssemblyOp(0).Register = oAssemblyOp(1).Register Then
                            If oAssemblyOp(0).Operand = oAssemblyOp(1).Operand Then
                                oAssemblyOp(1).Instruction = ioNOP
                            End If
                        End If
                End Select
        End Select
    Next
End Sub

Public Function Compile() As String
    Dim lIndex As Long
    Dim oAO As clsAssemblyOp
    Dim lInterleaveIndex As Long
    Dim lOffsetIndex As Long
    Dim lInterleaveCount As Long
    Dim bInterleaveReverse As Boolean
    
    Optimise
    
    mvOutput = Array()
            
    For lIndex = 0 To Count - 1
        Set oAO = moAssemblyOps(lIndex)
        lInterleaveCount = oAO.InterleaveCount
        bInterleaveReverse = oAO.InterleaveReverse
        
        For lOffsetIndex = 0 To oAO.OperandActualSize - 1
            For lInterleaveIndex = 0 To lInterleaveCount - 1
                Set oAO = moAssemblyOps(lIndex + lInterleaveIndex)
                Select Case oAO.Instruction
                    Case ioNOP
                        ' Do nothing
                    Case ioLD
                        Select Case oAO.Register
                            Case irA, irX, irY, irAX, irAY, irXY
                                Select Case oAO.Mode
                                    Case imAddress
                                        ExGenIdentOffset "LD" & RegisterName(oAO.Register), oAO.Operand, oAO.OperandSize, oAO.OperandActualSize, lOffsetIndex, bInterleaveReverse
                                    Case imConstant
                                        ExGenConstOffset "LD" & RegisterName(oAO.Register), oAO.Operand, oAO.OperandSize, oAO.OperandActualSize, lOffsetIndex, bInterleaveReverse
                                End Select
                                
                            Case irC, irI, irD
                                If oAO.Operand = 0 Then
                                    AddLine "CL" & RegisterName(oAO.Register)
                                ElseIf oAO.Operand = 1 Then
                                    AddLine "SE" & RegisterName(oAO.Register)
                                End If
                        End Select
                    Case ioST
                        ExGenIdentOffset "ST" & RegisterName(oAO.Register), oAO.Operand, oAO.OperandSize, oAO.OperandActualSize, lOffsetIndex, bInterleaveReverse
                    Case ioADC
                        Select Case oAO.Mode
                            Case imAddress
                                ExGenIdentOffset "ADC", oAO.Operand, oAO.OperandSize, oAO.OperandActualSize, lOffsetIndex, bInterleaveReverse
                            Case imConstant
                                ExGenConstOffset "ADC", oAO.Operand, oAO.OperandSize, oAO.OperandActualSize, lOffsetIndex, bInterleaveReverse
                        End Select
                    Case ioSBC
                        Select Case oAO.Mode
                            Case imAddress
                                ExGenIdentOffset "SBC", oAO.Operand, oAO.OperandSize, oAO.OperandActualSize, lOffsetIndex, bInterleaveReverse
                            Case imConstant
                                ExGenConstOffset "SBC", oAO.Operand, oAO.OperandSize, oAO.OperandActualSize, lOffsetIndex, bInterleaveReverse
                        End Select
                        
                    Case ioCMP
                        Select Case oAO.Mode
                            Case imAddress
                                ExGenIdentOffset "CMP", oAO.Operand, oAO.OperandSize, oAO.OperandActualSize, lOffsetIndex, bInterleaveReverse
                            Case imConstant
                                ExGenConstOffset "CMP", oAO.Operand, oAO.OperandSize, oAO.OperandActualSize, lOffsetIndex, bInterleaveReverse
                        End Select
                    Case ioAND
                        Select Case oAO.Mode
                            Case imAddress
                                ExGenIdentOffset "AND", oAO.Operand, oAO.OperandSize, oAO.OperandActualSize, lOffsetIndex, bInterleaveReverse
                            Case imConstant
                                ExGenConstOffset "AND", oAO.Operand, oAO.OperandSize, oAO.OperandActualSize, lOffsetIndex, bInterleaveReverse
                        End Select
                    Case ioOR
                        Select Case oAO.Mode
                            Case imAddress
                                ExGenIdentOffset "ORA", oAO.Operand, oAO.OperandSize, oAO.OperandActualSize, lOffsetIndex, bInterleaveReverse
                            Case imConstant
                                ExGenConstOffset "ORA", oAO.Operand, oAO.OperandSize, oAO.OperandActualSize, lOffsetIndex, bInterleaveReverse
                        End Select
                    Case ioEOR
                        Select Case oAO.Mode
                            Case imAddress
                                ExGenIdentOffset "EOR", oAO.Operand, oAO.OperandSize, oAO.OperandActualSize, lOffsetIndex, bInterleaveReverse
                            Case imConstant
                                ExGenConstOffset "EOR", oAO.Operand, oAO.OperandSize, oAO.OperandActualSize, lOffsetIndex, bInterleaveReverse
                        End Select
                    Case ioPH
                    Case ioPL
                    Case ioJMP
                        AddLine "JMP " & oAO.Operand
                    Case ioBRS
                        Select Case oAO.Register
                            Case irZ
                                AddLine "BEQ " & oAO.Operand
                            Case irV
                                AddLine "BVS " & oAO.Operand
                            Case irC
                                AddLine "BCS " & oAO.Operand
                            Case irN
                                AddLine "BMI " & oAO.Operand
                        End Select
                    Case ioBRC
                        Select Case oAO.Register
                            Case irZ
                                AddLine "BNE " & oAO.Operand
                            Case irV
                                AddLine "BVC " & oAO.Operand
                            Case irC
                                AddLine "BCC " & oAO.Operand
                            Case irN
                                AddLine "BPL " & oAO.Operand
                        End Select
                    Case ioDEC
                        Select Case oAO.Mode
                            Case imAddress
                                ExGenIdentOffset "DEC", oAO.Operand, oAO.OperandSize, oAO.OperandActualSize, lOffsetIndex, bInterleaveReverse
                            Case imRegister
                                AddLine "DE" & RegisterName(oAO.Register)
                        End Select
                    Case ioINC
                        Select Case oAO.Mode
                            Case imAddress
                                ExGenIdentOffset "INC", oAO.Operand, oAO.OperandSize, oAO.OperandActualSize, lOffsetIndex, bInterleaveReverse
                            Case imRegister
                                AddLine "IN" & RegisterName(oAO.Register)
                        End Select
                    Case ioROL
                        Select Case oAO.Mode
                            Case imAddress
                                ExGenIdentOffset "ROL", oAO.Operand, oAO.OperandSize, oAO.OperandActualSize, lOffsetIndex, bInterleaveReverse
                            Case imRegister
                                AddLine "ROL " & RegisterName(oAO.Register)
                        End Select
                    Case ioROR
                        Select Case oAO.Mode
                            Case imAddress
                                ExGenIdentOffset "ROR", oAO.Operand, oAO.OperandSize, oAO.OperandActualSize, lOffsetIndex, bInterleaveReverse
                            Case imRegister
                                AddLine "ROR " & RegisterName(oAO.Register)
                        End Select
                    Case ioASL
                        Select Case oAO.Mode
                            Case imAddress
                                ExGenIdentOffset "ASL", oAO.Operand, oAO.OperandSize, oAO.OperandActualSize, lOffsetIndex, bInterleaveReverse
                            Case imRegister
                                AddLine "ASL " & RegisterName(oAO.Register)
                        End Select
                    Case ioLSR
                        Select Case oAO.Mode
                            Case imAddress
                                ExGenIdentOffset "LSR", oAO.Operand, oAO.OperandSize, oAO.OperandActualSize, lOffsetIndex, bInterleaveReverse
                            Case imRegister
                                AddLine "LSR " & RegisterName(oAO.Register)
                        End Select
                    Case ioAddress
                        AddLine AddressValue(oAO.Operand, oAO.OperandSize)
                    Case ioLabel
                        AddLine oAO.Operand
                    Case ioAddressLabel
                        AddLine AddressValue(oAO.Operand, oAO.OperandSize) & " " & oAO.Operand2
                    Case ioConstant
                        BytesConstant oAO.Operand, oAO.OperandSize
                    Case ioLDAFlag
                        Select Case oAO.Register
                            Case irZ
                                AddLine "PHP"
                                AddLine "PLA"
                                AddLine "AND #02h"
                                AddLine "LSR A"
                            Case irV
                                AddLine "PHP"
                                AddLine "PLA"
                                AddLine "AND #40h"
                                AddLine "ASL A"
                                AddLine "ASL A"
                                AddLine "ROL A"
                            Case irC
                                AddLine "PHP"
                                AddLine "PLA"
                                AddLine "AND #01h"
                            Case irN
                                AddLine "PHP"
                                AddLine "PLA"
                                AddLine "AND #80h"
                                AddLine "ASL A"
                                AddLine "ROL A"
                        End Select
                    Case ioLDAInvFlag
                        Select Case oAO.Register
                            Case irZ
                                AddLine "PHP"
                                AddLine "PLA"
                                AddLine "AND #02h"
                                AddLine "EOR #02h"
                                AddLine "LSR A"
                            Case irV
                                AddLine "PHP"
                                AddLine "PLA"
                                AddLine "AND #40h"
                                AddLine "EOR #40h"
                                AddLine "ASL A"
                                AddLine "ASL A"
                                AddLine "ROL A"
                            Case irC
                                AddLine "PHP"
                                AddLine "PLA"
                                AddLine "AND #01h"
                                AddLine "EOR #01h"
                            Case irN
                                AddLine "PHP"
                                AddLine "PLA"
                                AddLine "AND #80h"
                                AddLine "EOR #80h"
                                AddLine "ASL A"
                                AddLine "ROL A"
                        End Select
                    Case ioJSR
                        AddLine "JSR " & oAO.Operand
                    Case ioRTS
                        AddLine "RTS"
                    Case ioSeparator
                        AddLine ""
                End Select
            Next
        Next
        lIndex = lIndex + lInterleaveCount - 1
    Next
    
    Compile = Join(mvOutput, vbCrLf)
End Function

Private Function RegisterName(ByVal irRegister As AssemblyOpRegister) As String
    Select Case irRegister
        Case irA
            RegisterName = "A"
        Case irX
            RegisterName = "X"
        Case irY
            RegisterName = "Y"
        Case irP
            RegisterName = "P"
        Case irS
            RegisterName = "S"
        Case irC
            RegisterName = "C"
        Case irN
            RegisterName = "N"
        Case irZ
            RegisterName = "Z"
        Case irI
            RegisterName = "I"
        Case irV
            RegisterName = "V"
        Case irD
            RegisterName = "D"
        Case irAX
            RegisterName = "AX"
        Case irAY
            RegisterName = "AY"
        Case irXY
            RegisterName = "XY"
    End Select
End Function

Private Function HexNum(ByVal lNumber As Long, Optional ByVal iPlaces As Integer = 4) As String
    HexNum = Hex$(lNumber)
    If Len(HexNum) <= iPlaces Then
        HexNum = String$(iPlaces - Len(HexNum), "0") & HexNum
    Else
        HexNum = Right$(HexNum, iPlaces)
    End If
    HexNum = HexNum & "h"
End Function

Private Function AddressValue(ByVal lAddress As Long, ByVal lSize As Long) As String
    If lAddress >= 256 Then
        AddressValue = HexNum(lAddress, 4)
    Else
        AddressValue = HexNum(lAddress, 2)
    End If
End Function

Private Function BytesConstant(ByVal lConstant As Long, ByVal lSize As Long)
    Dim lIndex As Long
    Dim sLine As String
    
    For lIndex = 1 To lSize
        sLine = sLine & "DB " & HexNum(lConstant And 255, 2) & " "
        lConstant = lConstant \ 256
    Next
    AddLine sLine
End Function

'Private Sub ExNegate(ByVal sIdentifier1 As String, ByVal lSize1 As Long, ByVal sIdentifier2 As String, ByVal lSize2 As Long)
'    Dim lOffset As Long
'
'    AddLine "LDA " & Offset(sIdentifier1, lSize1 - 1) ' Sign bit
'    AddLine "ROL A"
'    AddLine "LDA #00h"
'    AddLine "SBC #00h"
'    AddLine "CLC"
'    For lOffset = 1 To Largest(lSize1, lSize2)
'        AddLine "LDA " & Offset(sIdentifier1, lOffset - 1)
'        AddLine "EOR #FFh"
'        AddLine "ADC #00h"
'        AddLine "STA " & Offset(sIdentifier2, lOffset - 1)
'    Next
'End Sub

