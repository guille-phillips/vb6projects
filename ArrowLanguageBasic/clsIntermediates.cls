VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIntermediates"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Scope As clsScope

Public Count As Long
Private moIntermediates() As clsIntermediate

Private lBaseAddress As Long

Private moAssemblyOps As New clsAssemblyOps

Private msLabelStack() As String
Private mlLabelStackPointer As Long

Public Function Add(oMember As clsIntermediate)
    ReDim Preserve moIntermediates(Count)
    Set moIntermediates(Count) = oMember
    Count = Count + 1
End Function

Public Property Get Intermediates(ByVal lIndex As Long) As clsIntermediate
Attribute Intermediates.VB_UserMemId = 0
    Set Intermediates = moIntermediates(lIndex)
End Property

Public Property Set Intermediates(ByVal lIndex As Long, oMember As clsIntermediate)
    Set moIntermediates(lIndex) = oMember
End Property


Public Function MergeIntermediates(oIntermediates As clsIntermediates)
    Dim lIndex As Long
    
    For lIndex = 0 To oIntermediates.Count - 1
        Me.Add oIntermediates.Intermediates(lIndex)
    Next
End Function

Private Sub Optimise()
    Dim lIndex As Long
    Dim oIntermediate(2) As clsIntermediate
    Dim lIndex2 As Long
    Dim oObject As clsObject
Exit Sub
    For lIndex = 0 To Count - 1
        For lIndex2 = 0 To UBound(oIntermediate)
            If (lIndex + lIndex2) < Count Then
                Set oIntermediate(lIndex2) = moIntermediates(lIndex + lIndex2)
            Else
                Set oIntermediate(lIndex2) = Nothing
            End If
        Next
    
        Select Case oIntermediate(0).Operator
            Case opEqual, opNotEqual, opLessThan, opGreaterThan, opLessThanEqual, opGreaterThanEqual
                Select Case oIntermediate(1).Operator
                    Case opIsZero
                        If oIntermediate(1).Operand1 Is oIntermediate(0).Operand3 Then
                            Select Case oIntermediate(2).Operator
                                Case opIf, opWhile, opEndDoUntil
                                    oIntermediate(0).Options = oIntermediate(0).Options Or osFlagOnly
                                    oIntermediate(1).Operator = opNone
                                    Select Case oIntermediate(0).Operator
                                        Case opEqual
                                            oIntermediate(2).Options = osUseZFlag
                                        Case opNotEqual
                                            oIntermediate(2).Options = osUseZFlag Or osFlagSet
                                        Case opLessThan
                                            oIntermediate(2).Options = osUseCFlag Or osFlagSet
                                        Case opGreaterThan
                                            oIntermediate(2).Options = osUseCFlag Or osFlagSet
                                        Case opLessThanEqual
                                            oIntermediate(2).Options = osUseCFlag
                                        Case opGreaterThanEqual
                                            oIntermediate(2).Options = osUseCFlag
                                    End Select
                                Case opUntil, opEndDoWhile
                                    oIntermediate(0).Options = oIntermediate(0).Options Or osFlagOnly
                                    oIntermediate(1).Operator = opNone
                                    Select Case oIntermediate(0).Operator
                                        Case opEqual
                                            oIntermediate(2).Options = osUseZFlag Or osFlagSet
                                        Case opNotEqual
                                            oIntermediate(2).Options = osUseZFlag
                                        Case opLessThan
                                            oIntermediate(2).Options = osUseCFlag
                                        Case opGreaterThan
                                            oIntermediate(2).Options = osUseCFlag
                                        Case opLessThanEqual
                                            oIntermediate(2).Options = osUseCFlag Or osFlagSet
                                        Case opGreaterThanEqual
                                            oIntermediate(2).Options = osUseCFlag Or osFlagSet
                                    End Select
                            End Select
                        End If
                End Select
        End Select
    Next
End Sub

Public Function Compile() As clsAssemblyOps
    Dim vOutput As Variant
    Dim oIntermediate As clsIntermediate
    Dim lOperand1Address As Long
    Dim lOperand2Address As Long
    Dim lOperand3Address As Long
    Dim lIndex As Long
    
    Optimise
    
    For lIndex = 0 To Count - 1
        Set oIntermediate = moIntermediates(lIndex)
        Select Case oIntermediate.Operator
            Case opNone
                ' Do nothing
            Case opCopy
                OperatorCopy oIntermediate
            Case opCopyToRegister
                OperatorCopyToRegister oIntermediate
            Case opAdd
                OperatorAdd oIntermediate
            Case opSub
                OperatorSubtract oIntermediate
            Case opMultiply
                OperatorMultiply oIntermediate
            Case opDivide
                OperatorDivide oIntermediate
            Case opModulus
                OperatorModulus oIntermediate
            Case opEqual
                OperatorEqual oIntermediate
            Case opNotEqual
                OperatorNotEqual oIntermediate
            Case opLessThan
                OperatorLessThan oIntermediate
            Case opGreaterThan
                OperatorGreaterThan oIntermediate
            Case opLessThanEqual
                OperatorLessThanEqual oIntermediate
            Case opGreaterThanEqual
                OperatorGreaterThanEqual oIntermediate
            Case opAnd
                OperatorAnd oIntermediate
            Case opOr
                OperatorOr oIntermediate
            Case opEor
                OperatorEor oIntermediate
            Case opIsZero
                OperatorIsZero oIntermediate
            Case opIf
                OperatorIf oIntermediate
            Case opEndIf
                OperatorEndIf
            Case opStartWhile
                OperatorStartWhile
            Case opWhile
                OperatorWhile oIntermediate
            Case opEndWhile
                OperatorEndWhile
            Case opStartUntil
                OperatorStartUntil
            Case opUntil
                OperatorUntil oIntermediate
            Case opEndUntil
                OperatorEndUntil
            Case opStartDoWhile
                OperatorStartDoWhile
            Case opEndDoWhile
                OperatorEndDoWhile oIntermediate
            Case opStartDoUntil
                OperatorStartDoUntil
            Case opEndDoUntil
                OperatorEndDoUntil oIntermediate
            Case opReturn
                AddOp ioRTS, ""
            Case opOrigin
                AddOp ioAddress, oIntermediate.Operand1.Hydrate.UnitClass.Ranges(0).Constant, 2
            Case opObject
                OperatorObject oIntermediate
            Case opFunctionCall
                OperatorFunctionCall oIntermediate
            Case opFunction
                OperatorFunction oIntermediate
            Case opEndFunction
                OperatorEndFunction oIntermediate
            Case opSeparator
                AddOp ioSeparator, ""
        End Select
    Next
    
    Set Compile = moAssemblyOps
End Function

Private Sub OperatorObject(oIntermediate As clsIntermediate)
    Dim oOp As Object
    
    Select Case oIntermediate.Operand1.ObjectType
        Case otConst
            Set oOp = oIntermediate.Operand1.Hydrate
            AddOp ioAddressLabel, oOp.MemAddress, 2, , , , , , , oOp.UniqueIdentifier
            AddOp ioConstant, oOp.UnitClass.Ranges(0).Constant, oOp.UnitClass.Ranges.Size
        Case otVar, otVirt, otFunc
            Set oOp = oIntermediate.Operand1.Hydrate
            AddOp ioAddressLabel, oOp.MemAddress, 2, , , , , , , oOp.UniqueIdentifier
    End Select
End Sub

Private Sub AddOp(ByVal aoInstruction As AssemblyOps, ByVal vOperand As Variant, Optional ByVal lOperandSize As Long = 1, Optional ByVal lOperandActualSize As Long = 1, Optional ByVal aoMode As AssemblyOpMode = imImplied, Optional ByVal aoRegister As AssemblyOpRegister = irNone, Optional ByVal aoModeRegister As AssemblyOpRegister = irNone, Optional ByVal lInterleaveCount As Long = 1, Optional ByVal bInterleaveReverse As Boolean = False, Optional ByVal vOperand2 As Variant)
    Dim oAssemblyOp As New clsAssemblyOp
    
    oAssemblyOp.Instruction = aoInstruction
    oAssemblyOp.Operand = vOperand
    oAssemblyOp.Operand2 = vOperand2
    oAssemblyOp.OperandSize = lOperandSize
    oAssemblyOp.OperandActualSize = lOperandActualSize
    oAssemblyOp.Mode = aoMode
    oAssemblyOp.Register = aoRegister
    oAssemblyOp.ModeRegister = aoModeRegister
    oAssemblyOp.InterleaveCount = lInterleaveCount
    oAssemblyOp.InterleaveReverse = bInterleaveReverse
    
    moAssemblyOps.AddAssemblyOp oAssemblyOp
End Sub

Private Function Label(sIdentify As String) As String
    Dim lIndex As Long
    
    For lIndex = 1 To 5
        Label = Label & Chr$(Int(Rnd() * 26) + 65)
    Next
    Label = sIdentify & "_" & Label
End Function

Private Function OperatorCopy(oIntermediate As clsIntermediate)
    Dim oOp1 As Object
    Dim oOp2 As Object
    Dim sOperand1 As String
    Dim sOperand2 As String
    Dim lOperand1Size As Long
    Dim lOperand2Size As Long
    Dim bOperand2Const As Boolean
    Dim lOperand2 As Long
    
    Set oOp1 = oIntermediate.Operand1.Hydrate
    Set oOp2 = oIntermediate.Operand2.Hydrate
    sOperand1 = oOp1.UniqueIdentifier
    sOperand2 = oOp2.UniqueIdentifier
    lOperand1Size = oOp1.UnitClass.Ranges.Size
    lOperand2Size = oOp2.UnitClass.Ranges.Size
    
    bOperand2Const = oIntermediate.Operand2.ObjectType = otConst
    lOperand2 = oOp2.UnitClass.Ranges(0).Starting
    
    If Not bOperand2Const Then
        AddOp ioLD, sOperand2, lOperand2Size, lOperand1Size, imAddress, irA, , 2
    Else
        AddOp ioLD, lOperand2, lOperand2Size, lOperand1Size, imConstant, irA, , 2
    End If
    AddOp ioST, sOperand1, lOperand1Size, lOperand1Size, imAddress, irA
End Function

Private Function OperatorCopyToRegister(oIntermediate As clsIntermediate)
    Dim sOperand1 As String
    Dim sOperand2 As String
    Dim lOperand1Size As Long
    Dim lOperand2Size As Long
    Dim lLargestSize As Long
    Dim bOperand2Const As Boolean
    Dim lOperand2 As Long
    Dim rtRegisterTarget As RegisterTargets
    
    rtRegisterTarget = oIntermediate.Operand1.RegisterTarget
    sOperand2 = oIntermediate.Operand2.UniqueIdentifier
    lOperand2Size = oIntermediate.Operand2.UnitClass.Ranges.Size
    
    bOperand2Const = oIntermediate.Operand2.ObjectType = otConst
    lOperand2 = oIntermediate.Operand2.UnitClass.Ranges(0).Starting
    
    Select Case rtRegisterTarget
        Case rtA
            If Not bOperand2Const Then
                AddOp ioLD, sOperand2, lOperand2Size, 1, imAddress, irA
            Else
                AddOp ioLD, lOperand2, lOperand2Size, 1, imConstant, irA
            End If
        Case rtX
            If Not bOperand2Const Then
                AddOp ioLD, sOperand2, lOperand2Size, 1, imAddress, irX
            Else
                AddOp ioLD, lOperand2, lOperand2Size, 1, imConstant, irX
            End If
        Case rtY
            If Not bOperand2Const Then
                AddOp ioLD, sOperand2, lOperand2Size, 1, imAddress, irY
            Else
                AddOp ioLD, lOperand2, lOperand2Size, 1, imConstant, irY
            End If
        Case rtXY
            If Not bOperand2Const Then
                AddOp ioLD, sOperand2, lOperand2Size, 2, imAddress, irXY
            Else
                AddOp ioLD, lOperand2, lOperand2Size, 2, imConstant, irXY
            End If
        Case rtAX
            If Not bOperand2Const Then
                AddOp ioLD, sOperand2, lOperand2Size, 2, imAddress, irAX
            Else
                AddOp ioLD, lOperand2, lOperand2Size, 2, imConstant, irAX
            End If
        Case rtAY
            If Not bOperand2Const Then
                AddOp ioLD, sOperand2, lOperand2Size, 2, imAddress, irAY, , 2
            Else
                AddOp ioLD, lOperand2, lOperand2Size, 2, imConstant, irAY, , 2
            End If
    End Select
End Function

Private Function Offset(ByVal sIdentifier As String, ByVal lOffset As Long) As String
    If lOffset <> 0 Then
        Offset = sIdentifier & "+" & lOffset & "d"
    Else
        Offset = sIdentifier
    End If
End Function

Private Function OperatorAdd(oIntermediate As clsIntermediate)
    Dim oOp1 As Object
    Dim oOp2 As Object
    Dim oOp3 As Object
    Dim sOperand1 As String
    Dim sOperand2 As String
    Dim sOperand3 As String
    Dim lOperand1Size As Long
    Dim lOperand2Size As Long
    Dim lOperand3Size As Long
    Dim lOffset As Long
    Dim lLargestSize As Long
    Dim bOperand2Const As Boolean
    Dim lOperand2 As Long
    
    Set oOp1 = oIntermediate.Operand1.Hydrate
    Set oOp2 = oIntermediate.Operand2.Hydrate
    Set oOp3 = oIntermediate.Operand3.Hydrate
    
    sOperand1 = oOp1.UniqueIdentifier
    sOperand2 = oOp2.UniqueIdentifier
    sOperand3 = oOp3.UniqueIdentifier

    lOperand1Size = oOp1.UnitClass.Ranges.Size
    lOperand2Size = oOp2.UnitClass.Ranges.Size
    lOperand3Size = oOp3.UnitClass.Ranges.Size
    
    bOperand2Const = oOp2.ObjectType = otConst
    lOperand2 = oOp2.UnitClass.Ranges(0).Constant
    
    lLargestSize = Largest(lOperand1Size, lOperand2Size, lOperand3Size)
    
    AddOp ioLD, 0, , , imConstant, irC
    AddOp ioLD, sOperand1, lOperand1Size, lOperand3Size, imAddress, irA, , 3
    If Not bOperand2Const Then
        AddOp ioADC, sOperand2, lOperand2Size, lOperand3Size, imAddress
    Else
        AddOp ioADC, lOperand2, lOperand2Size, lOperand3Size, imConstant
    End If
    AddOp ioST, sOperand3, lOperand3Size, lOperand3Size, imAddress, irA
    
End Function

Private Function OperatorAnd(oIntermediate As clsIntermediate)
    Dim sOperand1 As String
    Dim sOperand2 As String
    Dim sOperand3 As String
    Dim lOperand1Size As Long
    Dim lOperand2Size As Long
    Dim lOperand3Size As Long
    Dim lLargestSize As Long
    Dim bOperand2Const As Boolean
    Dim lOperand2 As Long
        
    sOperand1 = oIntermediate.Operand1.UniqueIdentifier
    sOperand2 = oIntermediate.Operand2.UniqueIdentifier
    sOperand3 = oIntermediate.Operand3.UniqueIdentifier

    lOperand1Size = oIntermediate.Operand1.UnitClass.Ranges.Size
    lOperand2Size = oIntermediate.Operand2.UnitClass.Ranges.Size
    lOperand3Size = oIntermediate.Operand3.UnitClass.Ranges.Size
    
    lLargestSize = Largest(lOperand1Size, lOperand2Size, lOperand3Size)
    
    bOperand2Const = oIntermediate.Operand2.ObjectType = otConst
    lOperand2 = oIntermediate.Operand2.UnitClass.Ranges(0).Starting
            
    AddOp ioLD, sOperand1, lOperand1Size, lOperand3Size, imAddress, irA, , 3
    If Not bOperand2Const Then
        AddOp ioAND, sOperand2, lOperand2Size, lOperand3Size, imAddress
    Else
        AddOp ioAND, lOperand2, lOperand2Size, lOperand3Size, imConstant
    End If
    AddOp ioST, sOperand3, lOperand3Size, lOperand3Size, imAddress, irA
End Function

Private Function OperatorOr(oIntermediate As clsIntermediate)
    Dim sOperand1 As String
    Dim sOperand2 As String
    Dim sOperand3 As String
    Dim lOperand1Size As Long
    Dim lOperand2Size As Long
    Dim lOperand3Size As Long
    Dim lLargestSize As Long
    Dim bOperand2Const As Boolean
    Dim lOperand2 As Long
        
    sOperand1 = oIntermediate.Operand1.UniqueIdentifier
    sOperand2 = oIntermediate.Operand2.UniqueIdentifier
    sOperand3 = oIntermediate.Operand3.UniqueIdentifier

    lOperand1Size = oIntermediate.Operand1.UnitClass.Ranges.Size
    lOperand2Size = oIntermediate.Operand2.UnitClass.Ranges.Size
    lOperand3Size = oIntermediate.Operand3.UnitClass.Ranges.Size
    
    lLargestSize = Largest(lOperand1Size, lOperand2Size, lOperand3Size)
    
    bOperand2Const = oIntermediate.Operand2.ObjectType = otConst
    lOperand2 = oIntermediate.Operand2.UnitClass.Ranges(0).Starting
        
    AddOp ioLD, sOperand1, lOperand1Size, lOperand3Size, imAddress, irA, , 3
    If Not bOperand2Const Then
        AddOp ioOR, sOperand2, lOperand2Size, lOperand3Size, imAddress
    Else
        AddOp ioOR, lOperand2, lOperand2Size, lOperand3Size, imConstant
    End If
    AddOp ioST, sOperand3, lOperand3Size, lOperand3Size, imAddress, irA
End Function

Private Function OperatorEor(oIntermediate As clsIntermediate)
    Dim sOperand1 As String
    Dim sOperand2 As String
    Dim sOperand3 As String
    Dim lOperand1Size As Long
    Dim lOperand2Size As Long
    Dim lOperand3Size As Long
    Dim lLargestSize As Long
    Dim bOperand2Const As Boolean
    Dim lOperand2 As Long
        
    sOperand1 = oIntermediate.Operand1.UniqueIdentifier
    sOperand2 = oIntermediate.Operand2.UniqueIdentifier
    sOperand3 = oIntermediate.Operand3.UniqueIdentifier

    lOperand1Size = oIntermediate.Operand1.UnitClass.Ranges.Size
    lOperand2Size = oIntermediate.Operand2.UnitClass.Ranges.Size
    lOperand3Size = oIntermediate.Operand3.UnitClass.Ranges.Size
    
    lLargestSize = Largest(lOperand1Size, lOperand2Size, lOperand3Size)
    
    bOperand2Const = oIntermediate.Operand2.ObjectType = otConst
    lOperand2 = oIntermediate.Operand2.UnitClass.Ranges(0).Starting
        
    AddOp ioLD, sOperand1, lOperand1Size, lOperand3Size, imAddress, irA, , 3
    If Not bOperand2Const Then
        AddOp ioEOR, sOperand2, lOperand2Size, lOperand3Size, imAddress
    Else
        AddOp ioEOR, lOperand2, lOperand2Size, lOperand3Size, imConstant
    End If
    AddOp ioST, sOperand3, lOperand3Size, lOperand3Size, imAddress, irA
End Function

Private Function OperatorSubtract(oIntermediate As clsIntermediate)
    Dim sOperand1 As String
    Dim sOperand2 As String
    Dim sOperand3 As String
    Dim lOperand1Size As Long
    Dim lOperand2Size As Long
    Dim lOperand3Size As Long
    Dim lLargestSize As Long
    Dim bOperand2Const As Boolean
    Dim lOperand2 As Long
        
    sOperand1 = oIntermediate.Operand1.UniqueIdentifier
    sOperand2 = oIntermediate.Operand2.UniqueIdentifier
    sOperand3 = oIntermediate.Operand3.UniqueIdentifier

    lOperand1Size = oIntermediate.Operand1.UnitClass.Ranges.Size
    lOperand2Size = oIntermediate.Operand2.UnitClass.Ranges.Size
    lOperand3Size = oIntermediate.Operand3.UnitClass.Ranges.Size
    
    lLargestSize = Largest(lOperand1Size, lOperand2Size, lOperand3Size)
    
    bOperand2Const = oIntermediate.Operand2.ObjectType = otConst
    lOperand2 = oIntermediate.Operand2.UnitClass.Ranges(0).Starting
        
    AddOp ioLD, 1, , , imConstant, irC
    AddOp ioLD, sOperand1, lOperand1Size, lOperand3Size, imAddress, irA, , 3
    If Not bOperand2Const Then
        AddOp ioSBC, sOperand2, lOperand2Size, lOperand3Size, imAddress
    Else
        AddOp ioSBC, lOperand2, lOperand2Size, lOperand3Size, imConstant
    End If
    AddOp ioST, sOperand3, lOperand3Size, lOperand3Size, imAddress, irA
End Function

'Private Function OperatorMultiplyX(oIntermediate As clsIntermediate)
'    Dim sOperand1 As String
'    Dim sOperand2 As String
'    Dim sOperand3 As String
'    Dim lOperand1Size As Long
'    Dim lOperand2Size As Long
'    Dim lOperand3Size As Long
'    Dim lOffset As Long
'    Dim sLabel1 As String
'    Dim sLabel2 As String
'
'    sOperand1 = oIntermediate.Operand1.UniqueIdentifier
'    sOperand2 = oIntermediate.Operand2.UniqueIdentifier
'    sOperand3 = oIntermediate.Operand3.UniqueIdentifier
'
'    lOperand1Size = oIntermediate.Operand1.UnitClass.Ranges.Size
'    lOperand2Size = oIntermediate.Operand2.UnitClass.Ranges.Size
'    lOperand3Size = oIntermediate.Operand3.UnitClass.Ranges.Size
'
'    sLabel1 = Label("mult_next")
'    sLabel2 = Label("mult_noadd")
'
'    ExCopy sOperand1, lOperand1Size, sOperand3, lOperand3Size
'    AddLine "LDX #" & lOperand1Size * 8 & "d"
'    'AddLine "LDA #00h"
'    AddLine "CLC"
'    AddLine sLabel1
'    AddLine "BCC " & sLabel2
'    ExAdd Offset(sOperand3, lOperand1Size), lOperand2Size, sOperand2, lOperand2Size, Offset(sOperand3, lOperand1Size), lOperand2Size
'    'AddLine "CLC"
'    'AddLine "ADC " & sOperand2
'    AddLine sLabel2
'    'AddLine "ROR A"
'    ExROR sOperand3, lOperand3Size
'    AddLine "DEX"
'    AddLine "BPL " & sLabel1
'    'AddLine "STA " & Offset(sOperand3, lOperand3Size - 1)
'
'End Function

Private Function OperatorMultiply(oIntermediate As clsIntermediate)
    Dim sOperand1 As String
    Dim sOperand2 As String
    Dim sOperand3 As String
    Dim lOperand1Size As Long
    Dim lOperand2Size As Long
    Dim lOperand3Size As Long
    Dim lOffset As Long
    Dim sLabel1 As String
    Dim sLabel2 As String
    Dim bOperand2Const As Boolean
    Dim lOperand2 As Long
    
    sOperand1 = oIntermediate.Operand1.UniqueIdentifier
    sOperand2 = oIntermediate.Operand2.UniqueIdentifier
    sOperand3 = oIntermediate.Operand3.UniqueIdentifier

    lOperand1Size = oIntermediate.Operand1.UnitClass.Ranges.Size
    lOperand2Size = oIntermediate.Operand2.UnitClass.Ranges.Size
    lOperand3Size = oIntermediate.Operand3.UnitClass.Ranges.Size
    
    bOperand2Const = oIntermediate.Operand2.ObjectType = otConst
    lOperand2 = oIntermediate.Operand2.UnitClass.Ranges(0).Starting
    
    sLabel1 = Label("mult_next")
    sLabel2 = Label("mult_noadd")
    
    AddOp ioLD, sOperand1, lOperand1Size, lOperand3Size, imAddress, irA, , 2
    AddOp ioST, sOperand3, lOperand3Size, lOperand3Size, imAddress, irA
    AddOp ioLD, lOperand1Size * 8, 1, 1, imConstant, irX
    AddOp ioLD, 0, , , , irC
    AddOp ioLabel, sLabel1
    AddOp ioBRC, sLabel2, , , , irC
    AddOp ioLD, 0, , , , irC
    
    AddOp ioLD, Offset(sOperand3, lOperand1Size), lOperand2Size, lOperand2Size, imAddress, irA, , 3
    
    If Not bOperand2Const Then
        AddOp ioADC, sOperand2, lOperand2Size, lOperand2Size, imAddress
    Else
        AddOp ioADC, lOperand2, lOperand2Size, lOperand2Size, imConstant
    End If
    
    AddOp ioST, Offset(sOperand3, lOperand1Size), lOperand2Size, lOperand2Size, imAddress, irA
    
    AddOp ioLabel, sLabel2
    AddOp ioROR, sOperand3, lOperand3Size, lOperand3Size, imAddress, , , , True
    
    AddOp ioDEC, "", , , imRegister, irX
    AddOp ioBRC, sLabel1, , , imAddress, irN
End Function

Private Function OperatorDivide(oIntermediate As clsIntermediate)
    Dim sOperand1 As String
    Dim sOperand2 As String
    Dim sOperand3 As String
    Dim sOperand4 As String
    Dim lOperand1Size As Long
    Dim lOperand2Size As Long
    Dim lOperand3Size As Long
    Dim lOperand4Size As Long
    Dim lOffset As Long
    Dim sLabel1 As String
    Dim sLabel2 As String
    Dim sLabel3 As String
    Dim bOperand2Const As Boolean
    Dim lOperand2 As Long
    
    sOperand1 = oIntermediate.Operand1.UniqueIdentifier
    sOperand2 = oIntermediate.Operand2.UniqueIdentifier
    sOperand3 = oIntermediate.Operand3.UniqueIdentifier
    sOperand4 = oIntermediate.Operand4.UniqueIdentifier

    lOperand1Size = oIntermediate.Operand1.UnitClass.Ranges.Size
    lOperand2Size = oIntermediate.Operand2.UnitClass.Ranges.Size
    lOperand3Size = oIntermediate.Operand3.UnitClass.Ranges.Size
    lOperand4Size = oIntermediate.Operand4.UnitClass.Ranges.Size
    
    bOperand2Const = oIntermediate.Operand2.ObjectType = otConst
    lOperand2 = oIntermediate.Operand2.UnitClass.Ranges(0).Starting
        
    sLabel1 = Label("divide_next")
    sLabel2 = Label("divide_nosub")
    sLabel3 = Label("divide_compare")
    
    AddOp ioLD, sOperand1, lOperand1Size, lOperand1Size, imAddress, irA, , 2
    AddOp ioST, sOperand3, lOperand3Size, lOperand3Size, imAddress, irA
    
    AddOp ioLD, 0, lOperand2Size, lOperand2Size, imConstant, irA, , 2
    AddOp ioST, sOperand4, lOperand4Size, lOperand4Size, imAddress, irA
    
    AddOp ioLD, lOperand1Size * 8, 1, 1, imConstant, irX
    AddOp ioLabel, sLabel1
    
    AddOp ioLD, sOperand4, lOperand4Size, lOperand4Size, imAddress, irA, , 3, True
    If Not bOperand2Const Then
        AddOp ioCMP, sOperand2, lOperand2Size, lOperand4Size, imAddress
    Else
        AddOp ioCMP, lOperand2, lOperand2Size, lOperand4Size, imConstant
    End If
    AddOp ioBRC, sLabel3, , , , irZ
    AddOp ioLabel, sLabel3
    AddOp ioBRC, sLabel2, , , , irC
    
    AddOp ioLD, sOperand4, lOperand4Size, lOperand4Size, imAddress, irA, , 3
    If Not bOperand2Const Then
        AddOp ioSBC, sOperand2, lOperand2Size, lOperand4Size, imAddress
    Else
        AddOp ioSBC, lOperand2, lOperand2Size, lOperand4Size, imConstant
    End If
    AddOp ioST, sOperand4, lOperand4Size, lOperand4Size, imAddress, irA
    AddOp ioLabel, sLabel2

    AddOp ioROL, sOperand3, lOperand3Size, lOperand3Size, imAddress
    AddOp ioROL, sOperand4, lOperand4Size, lOperand4Size, imAddress
    
    AddOp ioDEC, "", , , imRegister, irX
    AddOp ioBRC, sLabel1, , , imAddress, irN
End Function

Private Function OperatorModulus(oIntermediate As clsIntermediate)
    Dim sOperand1 As String
    Dim sOperand2 As String
    Dim sOperand3 As String
    Dim sOperand4 As String
    Dim lOperand1Size As Long
    Dim lOperand2Size As Long
    Dim lOperand3Size As Long
    Dim lOperand4Size As Long
    Dim lOffset As Long
    Dim sLabel1 As String
    Dim sLabel2 As String
    Dim sLabel3 As String
    Dim bOperand2Const As Boolean
    Dim lOperand2 As Long
    
    sOperand1 = oIntermediate.Operand1.UniqueIdentifier
    sOperand2 = oIntermediate.Operand2.UniqueIdentifier
    sOperand3 = oIntermediate.Operand3.UniqueIdentifier
    sOperand4 = oIntermediate.Operand4.UniqueIdentifier

    lOperand1Size = oIntermediate.Operand1.UnitClass.Ranges.Size
    lOperand2Size = oIntermediate.Operand2.UnitClass.Ranges.Size
    lOperand3Size = oIntermediate.Operand3.UnitClass.Ranges.Size
    lOperand4Size = oIntermediate.Operand4.UnitClass.Ranges.Size
    
    bOperand2Const = oIntermediate.Operand2.ObjectType = otConst
    lOperand2 = oIntermediate.Operand2.UnitClass.Ranges(0).Starting
        
    sLabel1 = Label("divide_next")
    sLabel2 = Label("divide_nosub")
    
    sLabel1 = Label("divide_next")
    sLabel2 = Label("divide_nosub")
    sLabel3 = Label("divide_compare")
    
    AddOp ioLD, sOperand1, lOperand1Size, lOperand1Size, imAddress, irA, , 2
    AddOp ioST, sOperand4, lOperand4Size, lOperand4Size, imAddress, irA
    
    AddOp ioLD, 0, lOperand2Size, lOperand2Size, imConstant, irA, , 2
    AddOp ioST, sOperand3, lOperand3Size, lOperand3Size, imAddress, irA
    
    AddOp ioLD, lOperand1Size * 8, 1, 1, imConstant, irX
    AddOp ioLabel, sLabel1
    
    AddOp ioLD, sOperand3, lOperand3Size, lOperand3Size, imAddress, irA, , 3, True
    If Not bOperand2Const Then
        AddOp ioCMP, sOperand2, lOperand2Size, lOperand3Size, imAddress
    Else
        AddOp ioCMP, lOperand2, lOperand2Size, lOperand3Size, imConstant
    End If
    AddOp ioBRC, sLabel3, , , , irZ
    AddOp ioLabel, sLabel3
    AddOp ioBRC, sLabel2, , , , irC
    
    AddOp ioLD, sOperand3, lOperand3Size, lOperand3Size, imAddress, irA, , 3
    If Not bOperand2Const Then
        AddOp ioSBC, sOperand2, lOperand2Size, lOperand3Size, imAddress
    Else
        AddOp ioSBC, lOperand2, lOperand2Size, lOperand3Size, imConstant
    End If
    AddOp ioST, sOperand3, lOperand3Size, lOperand3Size, imAddress, irA
    AddOp ioLabel, sLabel2

    AddOp ioROL, sOperand4, lOperand4Size, lOperand4Size, imAddress
    AddOp ioROL, sOperand3, lOperand3Size, lOperand3Size, imAddress
    
    AddOp ioDEC, "", , , imRegister, irX
    AddOp ioBRC, sLabel1, , , imAddress, irN
    AddOp ioROR, sOperand3, lOperand3Size, lOperand3Size, imAddress, , , , True
End Function

Private Function OperatorEqual(oIntermediate As clsIntermediate)
    Dim sOperand1 As String
    Dim sOperand2 As String
    Dim sOperand3 As String
    Dim lOperand1Size As Long
    Dim lOperand2Size As Long
    Dim lOperand3Size As Long
    Dim lLargestSize As Long
    Dim sLabel As String
    Dim bOperand2Const As Boolean
    Dim lOperand2 As Long
        
    sOperand1 = oIntermediate.Operand1.UniqueIdentifier
    sOperand2 = oIntermediate.Operand2.UniqueIdentifier
    sOperand3 = oIntermediate.Operand3.UniqueIdentifier

    lOperand1Size = oIntermediate.Operand1.UnitClass.Ranges.Size
    lOperand2Size = oIntermediate.Operand2.UnitClass.Ranges.Size
    lOperand3Size = oIntermediate.Operand3.UnitClass.Ranges.Size
    
    lLargestSize = Largest(lOperand1Size, lOperand2Size, lOperand3Size)
    
    bOperand2Const = oIntermediate.Operand2.ObjectType = otConst
    lOperand2 = oIntermediate.Operand2.UnitClass.Ranges(0).Starting
        
    sLabel = Label("compare_exit")
    AddOp ioLD, sOperand1, lOperand1Size, lLargestSize, imAddress, irA, , 3, True
    If Not bOperand2Const Then
        AddOp ioCMP, sOperand2, lOperand2Size, lLargestSize, imAddress
    Else
        AddOp ioCMP, lOperand2, lOperand2Size, lLargestSize, imConstant
    End If
    AddOp ioBRC, sLabel, , , imAddress, irZ
    AddOp ioLabel, sLabel
    
    If (oIntermediate.Options And osFlagOnly) = 0 Then
        AddOp ioLDAFlag, "", , , , irZ
        AddOp ioST, sOperand3, 1, lOperand3Size, imAddress, irA
    End If
End Function

Private Function OperatorNotEqual(oIntermediate As clsIntermediate)
    Dim sOperand1 As String
    Dim sOperand2 As String
    Dim sOperand3 As String
    Dim lOperand1Size As Long
    Dim lOperand2Size As Long
    Dim lOperand3Size As Long
    Dim sLabel As String
    Dim lLargestSize As Long
    Dim bOperand2Const As Boolean
    Dim lOperand2 As Long
    
    sOperand1 = oIntermediate.Operand1.UniqueIdentifier
    sOperand2 = oIntermediate.Operand2.UniqueIdentifier
    sOperand3 = oIntermediate.Operand3.UniqueIdentifier

    lOperand1Size = oIntermediate.Operand1.UnitClass.Ranges.Size
    lOperand2Size = oIntermediate.Operand2.UnitClass.Ranges.Size
    lOperand3Size = oIntermediate.Operand3.UnitClass.Ranges.Size
    
    lLargestSize = Largest(lOperand1Size, lOperand2Size, lOperand3Size)
    
    bOperand2Const = oIntermediate.Operand2.ObjectType = otConst
    lOperand2 = oIntermediate.Operand2.UnitClass.Ranges(0).Starting
        
    sLabel = Label("compare_exit")
    AddOp ioLD, sOperand1, lOperand1Size, lLargestSize, imAddress, irA, , 3, True
    If Not bOperand2Const Then
        AddOp ioCMP, sOperand2, lOperand2Size, lLargestSize, imAddress
    Else
        AddOp ioCMP, lOperand2, lOperand2Size, lLargestSize, imConstant
    End If
    AddOp ioBRC, sLabel, , , imAddress, irZ
    AddOp ioLabel, sLabel
    If (oIntermediate.Options And osFlagOnly) = 0 Then
        AddOp ioLDAInvFlag, "", , , , irZ
        AddOp ioST, sOperand3, 1, lOperand3Size, imAddress, irA
    End If
End Function

Private Function OperatorLessThan(oIntermediate As clsIntermediate)
    Dim sOperand1 As String
    Dim sOperand2 As String
    Dim sOperand3 As String
    Dim lOperand1Size As Long
    Dim lOperand2Size As Long
    Dim lOperand3Size As Long
    Dim sLabel As String
    Dim lLargestSize As Long
    Dim bOperand2Const As Boolean
    Dim lOperand2 As Long
        
    sOperand1 = oIntermediate.Operand1.UniqueIdentifier
    sOperand2 = oIntermediate.Operand2.UniqueIdentifier
    sOperand3 = oIntermediate.Operand3.UniqueIdentifier

    lOperand1Size = oIntermediate.Operand1.UnitClass.Ranges.Size
    lOperand2Size = oIntermediate.Operand2.UnitClass.Ranges.Size
    lOperand3Size = oIntermediate.Operand3.UnitClass.Ranges.Size
    
    lLargestSize = Largest(lOperand1Size, lOperand2Size, lOperand3Size)
    
    bOperand2Const = oIntermediate.Operand2.ObjectType = otConst
    lOperand2 = oIntermediate.Operand2.UnitClass.Ranges(0).Starting
        
    sLabel = Label("compare_exit")
    AddOp ioLD, sOperand1, lOperand1Size, lLargestSize, imAddress, irA, , 3, True
    If Not bOperand2Const Then
        AddOp ioCMP, sOperand2, lOperand2Size, lLargestSize, imAddress
    Else
        AddOp ioCMP, lOperand2, lOperand2Size, lLargestSize, imConstant
    End If
    AddOp ioBRC, sLabel, , lLargestSize - 1, imAddress, irZ
    AddOp ioLabel, sLabel
    If (oIntermediate.Options And osFlagOnly) = 0 Then
        AddOp ioLDAInvFlag, "", , , , irC
        AddOp ioST, sOperand3, 1, lOperand3Size, imAddress, irA
    End If
End Function

Private Function OperatorGreaterThan(oIntermediate As clsIntermediate)
    Dim sOperand1 As String
    Dim sOperand2 As String
    Dim sOperand3 As String
    Dim lOperand1Size As Long
    Dim lOperand2Size As Long
    Dim lOperand3Size As Long
    Dim sLabel As String
    Dim lLargestSize As Long
    Dim bOperand2Const As Boolean
    Dim lOperand2 As Long
    
    sOperand1 = oIntermediate.Operand1.UniqueIdentifier
    sOperand2 = oIntermediate.Operand2.UniqueIdentifier
    sOperand3 = oIntermediate.Operand3.UniqueIdentifier

    lOperand1Size = oIntermediate.Operand1.UnitClass.Ranges.Size
    lOperand2Size = oIntermediate.Operand2.UnitClass.Ranges.Size
    lOperand3Size = oIntermediate.Operand3.UnitClass.Ranges.Size
    
    lLargestSize = Largest(lOperand1Size, lOperand2Size, lOperand3Size)
    
    bOperand2Const = oIntermediate.Operand2.ObjectType = otConst
    lOperand2 = oIntermediate.Operand2.UnitClass.Ranges(0).Starting
        
    sLabel = Label("compare_exit")
    If Not bOperand2Const Then
        AddOp ioLD, sOperand2, lOperand2Size, lLargestSize, imAddress, irA, , 3, True
    Else
        AddOp ioLD, lOperand2, lOperand2Size, lLargestSize, imConstant, irA, , 3, True
    End If
    
    AddOp ioCMP, sOperand1, lOperand1Size, lLargestSize, imAddress
    AddOp ioBRC, sLabel, , , imAddress, irZ
    AddOp ioLabel, sLabel
    
    If (oIntermediate.Options And osFlagOnly) = 0 Then
        AddOp ioLDAInvFlag, "", , , , irC
        AddOp ioST, sOperand3, 1, lOperand3Size, imAddress, irA
    End If
End Function

Private Function OperatorLessThanEqual(oIntermediate As clsIntermediate)
    Dim sOperand1 As String
    Dim sOperand2 As String
    Dim sOperand3 As String
    Dim lOperand1Size As Long
    Dim lOperand2Size As Long
    Dim lOperand3Size As Long
    Dim sLabel As String
    Dim lLargestSize As Long
    Dim bOperand2Const As Boolean
    Dim lOperand2 As Long
    
    sOperand1 = oIntermediate.Operand1.UniqueIdentifier
    sOperand2 = oIntermediate.Operand2.UniqueIdentifier
    sOperand3 = oIntermediate.Operand3.UniqueIdentifier

    lOperand1Size = oIntermediate.Operand1.UnitClass.Ranges.Size
    lOperand2Size = oIntermediate.Operand2.UnitClass.Ranges.Size
    lOperand3Size = oIntermediate.Operand3.UnitClass.Ranges.Size
    
    lLargestSize = Largest(lOperand1Size, lOperand2Size, lOperand3Size)
    
    bOperand2Const = oIntermediate.Operand2.ObjectType = otConst
    lOperand2 = oIntermediate.Operand2.UnitClass.Ranges(0).Starting
        
    sLabel = Label("compare_exit")
    If Not bOperand2Const Then
        AddOp ioLD, sOperand2, lOperand2Size, lLargestSize, imAddress, irA, , 3, True
    Else
        AddOp ioLD, lOperand2, lOperand2Size, lLargestSize, imConstant, irA, , 3, True
    End If
    AddOp ioCMP, sOperand1, lOperand1Size, lLargestSize, imAddress
    AddOp ioBRC, sLabel, , , imAddress, irZ
    AddOp ioLabel, sLabel
    If (oIntermediate.Options And osFlagOnly) = 0 Then
        AddOp ioLDAFlag, "", , , , irC
        AddOp ioST, sOperand3, 1, lOperand3Size, imAddress, irA
    End If
End Function

Private Function OperatorGreaterThanEqual(oIntermediate As clsIntermediate)
    Dim sOperand1 As String
    Dim sOperand2 As String
    Dim sOperand3 As String
    Dim lOperand1Size As Long
    Dim lOperand2Size As Long
    Dim lOperand3Size As Long
    Dim sLabel As String
    Dim lLargestSize As Long
    Dim bOperand2Const As Boolean
    Dim lOperand2 As Long
    
    sOperand1 = oIntermediate.Operand1.UniqueIdentifier
    sOperand2 = oIntermediate.Operand2.UniqueIdentifier
    sOperand3 = oIntermediate.Operand3.UniqueIdentifier

    lOperand1Size = oIntermediate.Operand1.UnitClass.Ranges.Size
    lOperand2Size = oIntermediate.Operand2.UnitClass.Ranges.Size
    lOperand3Size = oIntermediate.Operand3.UnitClass.Ranges.Size
    
    lLargestSize = Largest(lOperand1Size, lOperand2Size, lOperand3Size)
    
    bOperand2Const = oIntermediate.Operand2.ObjectType = otConst
    lOperand2 = oIntermediate.Operand2.UnitClass.Ranges(0).Constant
        
    sLabel = Label("compare_exit")
    AddOp ioLD, sOperand1, lOperand1Size, lLargestSize, imAddress, irA, , 3, True
    If Not bOperand2Const Then
        AddOp ioCMP, sOperand2, lOperand2Size, lLargestSize, imAddress
    Else
        AddOp ioCMP, lOperand2, lOperand2Size, lLargestSize, imConstant
    End If
    AddOp ioBRC, sLabel, , , imAddress, irZ
    AddOp ioLabel, sLabel
    If (oIntermediate.Options And osFlagOnly) = 0 Then
        AddOp ioLDAFlag, "", , , , irC
        AddOp ioST, sOperand3, 1, lOperand3Size, imAddress, irA
    End If
End Function
                
Private Function OperatorIsZero(oIntermediate As clsIntermediate)
    Dim sOperand1 As String
    Dim lOperand1Size As Long
    
    sOperand1 = oIntermediate.Operand1.UniqueIdentifier
    lOperand1Size = oIntermediate.Operand1.UnitClass.Ranges.Size
    
    AddOp ioLD, sOperand1, 1, 1, imAddress, irA
    AddOp ioOR, Offset(sOperand1, 1), lOperand1Size - 1, lOperand1Size - 1, imAddress
End Function

Private Function OperatorFunctionCall(oIntermediate As clsIntermediate)
    AddOp ioJSR, oIntermediate.Operand1.Identifier
End Function

Private Function OperatorFunction(oIntermediate As clsIntermediate)
    If oIntermediate.Operand1.HasFunctionAddress Then
        AddOp ioAddress, oIntermediate.Operand1.FunctionAddress, 2
    End If
    AddOp ioLabel, oIntermediate.Operand1.Identifier
End Function

Private Function OperatorEndFunction(oIntermediate As clsIntermediate)
    AddOp ioRTS, ""
End Function

Private Function OperatorIf(oIntermediate As clsIntermediate)
    BranchOnFlags oIntermediate, PushLabel("exit_if")
End Function

Private Function BranchOnFlags(oIntermediate As clsIntermediate, sLabel As String)
    If (oIntermediate.Options And osUseZFlag) <> 0 Then
        If (oIntermediate.Options And osFlagSet) <> 0 Then
            AddOp ioBRS, sLabel, , , imAddress, irZ
        Else
            AddOp ioBRC, sLabel, , , imAddress, irZ
        End If
    ElseIf (oIntermediate.Options And osUseCFlag) <> 0 Then
        If (oIntermediate.Options And osFlagSet) <> 0 Then
            AddOp ioBRS, sLabel, , , imAddress, irC
        Else
            AddOp ioBRC, sLabel, , , imAddress, irC
        End If
    ElseIf (oIntermediate.Options And osUseVFlag) <> 0 Then
        If (oIntermediate.Options And osFlagSet) <> 0 Then
            AddOp ioBRS, sLabel, , , imAddress, irV
        Else
            AddOp ioBRC, sLabel, , , imAddress, irV
        End If
    ElseIf (oIntermediate.Options And osUseNFlag) <> 0 Then
        If (oIntermediate.Options And osFlagSet) <> 0 Then
            AddOp ioBRS, sLabel, , , imAddress, irN
        Else
            AddOp ioBRC, sLabel, , , imAddress, irN
        End If
    End If
End Function

Private Function OperatorEndIf()
    AddOp ioLabel, PullLabel()
End Function

Private Function OperatorStartWhile()
    Dim sLabel As String
    
    sLabel = PushLabel("while")
    AddOp ioLabel, "start_" & sLabel
End Function

Private Function OperatorWhile(ByVal oIntermediate As clsIntermediate)
    'AddOp ioBRC, "exit_" & GetLabel(), , , imAddress, irZ
    BranchOnFlags oIntermediate, "exit_" & GetLabel()
End Function

Private Function OperatorEndWhile()
    Dim sLabel As String
    
    sLabel = PullLabel()
    
    AddOp ioJMP, "start_" & sLabel, , , imAddress
    AddOp ioLabel, "exit_" & sLabel
End Function

Private Function OperatorStartUntil()
    Dim sLabel As String
    
    sLabel = PushLabel("until")
    AddOp ioLabel, "start_" & sLabel
End Function

Private Function OperatorUntil(ByVal oIntermediate As clsIntermediate)
'    AddOp ioBRS, "exit_" & GetLabel(), , , imAddress, irZ
    BranchOnFlags oIntermediate, "exit_" & GetLabel()
End Function

Private Function OperatorEndUntil()
    Dim sLabel As String
    
    sLabel = PullLabel()
    
    AddOp ioJMP, "start_" & sLabel, , , imAddress
    AddOp ioLabel, "exit_" & sLabel
End Function

Private Function OperatorStartDoWhile()
    AddOp ioLabel, PushLabel("start_dowhile"), , , imAddress, irZ
End Function

Private Function OperatorEndDoWhile(ByVal oIntermediate As clsIntermediate)
'    AddOp ioBRS, PullLabel(), , , imAddress, irZ
    BranchOnFlags oIntermediate, PullLabel()
End Function

Private Function OperatorStartDoUntil()
    AddOp ioLabel, PushLabel("start_dountil"), , , imAddress, irZ
End Function

Private Function OperatorEndDoUntil(ByVal oIntermediate As clsIntermediate)
'    AddOp ioBRC, PullLabel(), , , imAddress, irZ
    BranchOnFlags oIntermediate, PullLabel()
End Function

 

'Private Function Offset(ByVal sIdentifier, ByVal lOffset As Long)
'    If lOffset > 0 Then
'        Offset = sIdentifier & "+" & lOffset & "d"
'    Else
'        Offset = sIdentifier
'    End If
'End Function


Private Function PushLabel(ByVal sLabel As String) As String
    Dim sQualifiedLabel As String
    
    sQualifiedLabel = Label(sLabel)
    ReDim Preserve msLabelStack(mlLabelStackPointer)
    msLabelStack(mlLabelStackPointer) = sQualifiedLabel
    mlLabelStackPointer = mlLabelStackPointer + 1
    PushLabel = sQualifiedLabel
End Function

Private Function PullLabel() As String
    mlLabelStackPointer = mlLabelStackPointer - 1
    PullLabel = msLabelStack(mlLabelStackPointer)
End Function

Private Function GetLabel() As String
    GetLabel = msLabelStack(mlLabelStackPointer - 1)
End Function

Private Sub AppendAssemblyOps(ByVal oAssemblyOps As clsAssemblyOps)
    Dim lIndex As Long
    
    For lIndex = 0 To oAssemblyOps.Count - 1
        moAssemblyOps.AddAssemblyOp oAssemblyOps.AssemblyOps(lIndex)
    Next
End Sub
 
