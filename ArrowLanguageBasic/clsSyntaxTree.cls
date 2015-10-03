VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSyntaxTree"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public moIntermediates As New clsIntermediates
Public Scope As clsScope

Private mlLineStarts() As Long
Private mlLineStartsCount As Long

Private Sub CalculateLineStarts(ByVal sText As String)
    Dim lIndex As Long
    
    ReDim mlLineStarts(0)
    mlLineStarts(0) = 1
    mlLineStartsCount = 1
    For lIndex = 1 To Len(sText)
        If Mid$(sText, lIndex, 2) = vbCrLf Then
            ReDim Preserve mlLineStarts(mlLineStartsCount)
            mlLineStarts(mlLineStartsCount) = lIndex + 2
            mlLineStartsCount = mlLineStartsCount + 1
            lIndex = lIndex + 1
        End If
    Next
End Sub

Public Function FindLineNumber(ByVal lPosition As Long) As Long
    Dim lIndex As Long
    
    For lIndex = mlLineStartsCount - 1 To 0 Step -1
        If lPosition >= mlLineStarts(lIndex) Then
            FindLineNumber = lIndex + 1
            Exit For
        End If
    Next
End Function

Public Function LoadProgram(oParser As ISaffronObject) As clsIntermediates
    Dim oFSO As New FileSystemObject
    Dim oTree As New SaffronTree
    Dim oIntermediate As clsIntermediate
    
    SaffronStream.Text = oFSO.OpenTextFile(App.Path & "\testprogram.txt", ForReading).ReadAll
    CalculateLineStarts SaffronStream.Text
    
    Set oTree = New SaffronTree
    If Not oParser.Parse(oTree) Then
        MsgBox "Syntax error"
        End
    End If
    
    Set Scope = New clsScope
    Set Scope.moSyntaxTree = Me
    
    ' Assemby Origin
    moIntermediates.Add IntermediateFactory.Create(opOrigin, ObjectFactory.Create(otConst, ClassFactory.Create(, RangeFactory.Create(&H7100))))
        
    DoStatements oTree
    
    moIntermediates.Add IntermediateFactory.Create(opReturn)
    
    Set LoadProgram = moIntermediates
End Function

Public Sub DoStatements(oTree As SaffronTree)
    Dim oStatement As SaffronTree
    
    For Each oStatement In oTree.SubTree
        DoStatement oStatement(1)
    Next
End Sub

Private Sub DoStatement(oTree As SaffronTree)
    Select Case oTree.Index
        Case 1 ' return statement
            DoReturnStatement oTree(1)
        Case 2 ' function statement
            DoFunctionStatement oTree(1)
        Case 3 ' if statement
            DoIfStatement oTree(1)
        Case 4 ' while statement
            DoWhileStatement oTree(1)
        Case 5 ' until statement
            DoUntilStatement oTree(1)
        Case 6 ' do while statement
            DoDoWhileStatement oTree(1)
        Case 7 ' do until statement
            DoDoUntilStatement oTree(1)
        Case 8 ' Class Declaration
            DoClassDeclaration oTree(1)
        Case 9 ' Object Instance
            DoObject oTree(1)
        Case 10 ' Expression
            DoExpression oTree(1)
        Case 11 ' white space
            ' do nothing
    End Select
End Sub

Private Sub DoClassDeclaration(oTree As SaffronTree)
    Dim oClass As New clsClass
    Dim lIndex As Long
    
    If Not Scope.moClasses.GetByIdentifier(oTree(1).Text) Is Nothing Then
        MsgBox "Class: " & oTree(1).Text & " already defined on line " & Scope.FindLineNumber(oTree(1).TextStart), vbOKOnly
        End
    End If
    
    Set oClass = GetCreateRanges(oTree(2)(1))
    oClass.Identifier = oTree(1).Text
    
    Scope.moClasses.Add oClass
End Sub

Private Function GetCreateRanges(oTree As SaffronTree) As clsClass
    Dim oClass As New clsClass
    Dim lIndex As Long
    
    oClass.AccessType = atByte
    oClass.Ranges.AccessType = oClass.AccessType
    For lIndex = 1 To oTree.SubTree.Count
        oClass.Ranges.Add RangeFactory.Create(oTree(lIndex)(1).Text, oTree(lIndex)(2).Text)
    Next
    
    Set GetCreateRanges = oClass
End Function

Private Sub DoObject(oTree As SaffronTree)
    Dim oObject As Object
    Dim oClass As clsClass
    Dim lType As Long
    
    lType = oTree(1).Index
    Select Case lType
        Case 1 ' var
            Set oObject = New clsObjectVar
        Case 2 ' param
            Set oObject = New clsObjectParam
    End Select
    
    oObject.Identifier = oTree(2).Text
    
    If Not Scope.moObjects.GetByIdentifier(oTree(2).Text) Is Nothing Then
        MsgBox "Object: " & oTree(2).Text & " already defined on line " & Scope.FindLineNumber(oTree(2).TextStart), vbOKOnly
        End
    End If
    
    If oTree(3).Index = 1 Then
        If oTree(3)(1).Index = 1 Then
            oObject.HasHardAddress = True
            oObject.MemAddress = Base16(oTree(3).Text)
        Else
            If lType = 2 Then ' param
                oObject.HasRegisterTarget = True
                Select Case oTree(3)(1)(1).Text
                    Case "A"
                        oObject.RegisterTarget = rtA
                    Case "X"
                        oObject.RegisterTarget = rtX
                    Case "Y"
                        oObject.RegisterTarget = rtY
                    Case "XY"
                        oObject.RegisterTarget = rtXY
                    Case "AX"
                        oObject.RegisterTarget = rtAX
                    Case "AY"
                        oObject.RegisterTarget = rtAY
                End Select
            Else
                MsgBox "Register target @" & oTree(3)(1)(1).Text & " can only be used for param types on line " & Scope.FindLineNumber(oTree(3)(1)(1).TextStart), vbOKOnly
                End
            End If
        End If
    End If
    
    Select Case oTree(4).Index
        Case 1 ' class reference
            Set oObject.UnitClass = Scope.GetClassByIdentifier(oTree(4)(1).Text)
            If oObject.UnitClass Is Nothing Then
                MsgBox "Class: " & oTree(4)(1).Text & " not found on line " & Scope.FindLineNumber(oTree(4)(1).TextStart), vbOKOnly
                End
            End If
        Case 2 ' straight unit
            Set oObject.UnitClass = GetCreateRanges(oTree(4)(1)(1))
    End Select
    
    Scope.moObjects.AddObject oObject
    If lType = 2 Then ' param
        Scope.moParamObjects.AddObject oObject
    End If
End Sub

Private Function Base16(ByVal sNumber As String) As Long
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

Private Function DoExpression(oTree As SaffronTree, Optional bChild As Boolean) As Object
    Dim bTerm As Boolean
    Dim oSubTree As SaffronTree
    Dim oPreviousObject As Object
    Dim oObject As clsObject
    Dim opOperator As ILOperators
    Dim oIntermediate As clsIntermediate
    Dim oVirtObject As clsObjectVirt
    Dim oVirtClass As clsClass
    Dim oVirtObject2 As clsObject
    Dim oVirtClass2 As clsClass
    Dim oConstantObject As clsObject
    Dim bIsTarget As Boolean
    Dim lIndex As Long
    Dim oIntermediateSubroutine As clsIntermediate
    
    bTerm = True
    
    Dim oTargets() As clsObjectVar
    Dim lTargetsCount As Long
    
    For Each oSubTree In oTree.SubTree
        Select Case bTerm
            Case True ' Term
                If oPreviousObject Is Nothing Then
                    Select Case oSubTree.Index
                        Case 1 ' function call
                            Set oPreviousObject = DoFunctionCallOrIndex(oSubTree(1))
                            oPreviousObject.IncreaseReferenceCount
                            Set DoExpression = oPreviousObject
                            Set oIntermediateSubroutine = New clsIntermediate
                            oIntermediateSubroutine.Create opFunctionCall, oPreviousObject
                            moIntermediates.Add oIntermediateSubroutine
                        Case 2 ' identifier
                            bIsTarget = oSubTree(1)(1).Index = 1
                            Set oPreviousObject = Scope.GetObjectByIdentifier(oSubTree(1)(2).Text)
                            If oPreviousObject Is Nothing Then
                                MsgBox "Variable: " & oSubTree(1)(2).Text & " not found on line " & Scope.FindLineNumber(oSubTree(1)(2).TextStart), vbOKOnly
                                End
                            End If
                            Set DoExpression = oPreviousObject
                            oPreviousObject.IncreaseReferenceCount
                            
                            If bIsTarget Then
                                ReDim Preserve oTargets(lTargetsCount)
                                Set oTargets(lTargetsCount) = oPreviousObject
                                lTargetsCount = lTargetsCount + 1
                            End If
                        Case 3 ' constant
                            Set oObject = ObjectFactory.Create(otConst, ClassFactory.Create("", RangeFactory.Create(oSubTree(1).Text)))
                            oObject.IncreaseReferenceCount
                            Scope.moObjects.AddObject oObject
                            Set oPreviousObject = oObject
                            Set DoExpression = oPreviousObject
                        Case 4 ' brackets
                            Set oPreviousObject = DoExpression(oSubTree(1)(1), True)
                            Set DoExpression = oPreviousObject
                    End Select
                Else
                    Select Case oSubTree.Index
                        Case 1 ' function call
                            Set oObject = DoFunctionCallOrIndex(oSubTree(1))
                            oObject.IncreaseReferenceCount
                            
                            Set oIntermediateSubroutine = New clsIntermediate
                            oIntermediateSubroutine.Create opFunctionCall, oObject
                            moIntermediates.Add oIntermediateSubroutine
                        Case 2 ' identifier
                            bIsTarget = oSubTree(1)(1).Index = 1
                            Set oObject = Scope.GetObjectByIdentifier(oSubTree(1)(2).Text)
                            If oObject Is Nothing Then
                                MsgBox "Variable: " & oSubTree(1)(2).Text & " not found on line " & Scope.FindLineNumber(oSubTree(1)(2).TextStart), vbOKOnly
                                End
                            End If
                            oObject.IncreaseReferenceCount
                            If bIsTarget Then
                                ReDim Preserve oTargets(lTargetsCount)
                                Set oTargets(lTargetsCount) = oObject
                                lTargetsCount = lTargetsCount + 1
                            End If
                        Case 3 ' constant
                            Set oObject = ObjectFactory.Create(otConst, ClassFactory.Create("", RangeFactory.Create(oSubTree(1).Text)))
                            oObject.IncreaseReferenceCount
                            'Scope.moObjects.AddObject oObject
                        Case 4 ' brackets
                            Set oObject = DoExpression(oSubTree(1)(1), True)
                    End Select
                    
                    Dim lOperand1Starting As Long
                    Dim lOperand1Ending As Long
                    Dim lOperand2Starting As Long
                    Dim lOperand2Ending As Long
                    Dim lOperand1ByteSize As Long
                    Dim lOperand2ByteSize As Long
                    
                    lOperand1Starting = oPreviousObject.UnitClass.Ranges(0).Starting
                    lOperand1Ending = oPreviousObject.UnitClass.Ranges(0).Ending
                    lOperand2Starting = oObject.UnitClass.Ranges(0).Starting
                    lOperand2Ending = oObject.UnitClass.Ranges(0).Ending
                    
                    lOperand1ByteSize = oPreviousObject.UnitClass.Ranges.Size
                    lOperand2ByteSize = oObject.UnitClass.Ranges.Size
                    
                    If TypeOf oPreviousObject Is clsObjectConst And TypeOf oObject Is clsObjectConst Then
                        Dim lConstantResult As Long
                        
                        Select Case opOperator
                            Case opAdd
                                lConstantResult = lOperand1Starting + lOperand2Starting
                            Case opSub
                                lConstantResult = lOperand1Starting - lOperand2Starting
                            Case opMultiply
                                lConstantResult = lOperand1Starting * lOperand2Starting
                            Case opDivide
                                lConstantResult = lOperand1Starting \ lOperand2Starting
                            Case opModulus
                                lConstantResult = lOperand1Starting Mod lOperand2Starting
                            Case opEqual
                                lConstantResult = -(lOperand1Starting = lOperand2Starting)
                            Case opNotEqual
                                lConstantResult = -(lOperand1Starting <> lOperand2Starting)
                            Case opLessThan
                                lConstantResult = -(lOperand1Starting < lOperand2Starting)
                            Case opGreaterThan
                                lConstantResult = -(lOperand1Starting > lOperand2Starting)
                            Case opLessThanEqual
                                lConstantResult = -(lOperand1Starting <= lOperand2Starting)
                            Case opGreaterThanEqual
                                lConstantResult = -(lOperand1Starting >= lOperand2Starting)
                            Case opAnd
                                lConstantResult = -(lOperand1Starting And lOperand2Starting)
                            Case opOr
                                lConstantResult = -(lOperand1Starting Or lOperand2Starting)
                            Case opEor
                                lConstantResult = -(lOperand1Starting Xor lOperand2Starting)
                        End Select
                        oPreviousObject.UnitClass.Ranges(0).Constant = lConstantResult
                    Else
                        Set oIntermediate = IntermediateFactory.Create(opOperator, oPreviousObject, oObject)
                        Set oPreviousObject = oObject
                        Set DoExpression = oPreviousObject

                        Select Case opOperator
                            Case opAdd, opSub, opMultiply, opDivide, opModulus, opEqual, opNotEqual, opLessThan, opGreaterThan, opLessThanEqual, opGreaterThanEqual, opAnd, opOr, opEor
                                Set oVirtClass = New clsClass
                                
                                Select Case opOperator
                                    Case opAdd
                                        oVirtClass.Ranges.Add RangeFactory.Create(lOperand1Starting + lOperand2Starting, lOperand1Ending + lOperand2Ending)
                                    Case opSub
                                        oVirtClass.Ranges.Add RangeFactory.Create(0, lOperand1Ending - lOperand2Starting)
                                    Case opMultiply
                                        oVirtClass.Ranges.Add RangeFactory.Create(lOperand1Starting * lOperand2Starting, lOperand1Ending * lOperand2Ending)
'                                        oVirtClass.ByteSize = lOperand1ByteSize + lOperand2ByteSize
                                    Case opDivide
                                        oVirtClass.Ranges.Add RangeFactory.Create(lOperand1Starting, lOperand1Ending)
                                        
                                        Set oVirtClass2 = New clsClass
                                        oVirtClass2.Ranges(0).Create lOperand2Starting, lOperand2Ending
                                        
                                        Set oVirtObject2 = New clsObject
                                        Set oVirtObject2.UnitClass = oVirtClass2
                                        oVirtObject2.ObjectType = otVirt
                                        oVirtObject2.IncreaseReferenceCount
                                        Scope.moObjects.AddObject oVirtObject2
                                        Set oIntermediate.Operand4 = oVirtObject2
                                        
                                        oVirtObject2.PrePopSize = IIf(oIntermediate.Operand1.ObjectType <> otVirt, -lOperand1ByteSize, 0)
                                        If oIntermediate.Operand1.ObjectType <> otVirt Then
                                            If oIntermediate.Operand2.ObjectType <> otVirt Then
                                                oVirtObject2.PostPopSize = lOperand1ByteSize + lOperand2ByteSize
                                            Else
                                                oVirtObject2.PostPopSize = lOperand1ByteSize
                                            End If
                                        Else
                                            oVirtObject2.PostPopSize = lOperand2ByteSize
                                        End If
                                        
                                    Case opModulus
                                        oVirtClass.Ranges.Add RangeFactory.Create(lOperand2Starting, lOperand2Ending)
                                        
                                        Set oVirtClass2 = New clsClass
                                        oVirtClass2.Ranges(0).Create lOperand1Starting, lOperand1Ending
                                        
                                        Set oVirtObject2 = New clsObject
                                        Set oVirtObject2.UnitClass = oVirtClass2
                                        
                                        oVirtObject2.ObjectType = otVirt
                                        oVirtObject2.IncreaseReferenceCount
                                        Scope.moObjects.AddObject oVirtObject2
                                        Set oIntermediate.Operand4 = oVirtObject2
                                    
                                        If oIntermediate.Operand1.ObjectType <> otVirt Then
                                            If oIntermediate.Operand2.ObjectType <> otVirt Then
                                                oVirtObject2.PrePopSize = -lOperand2ByteSize
                                                oVirtObject2.PostPopSize = lOperand1ByteSize + lOperand2ByteSize
                                            Else
                                                oVirtObject2.PrePopSize = 0
                                                oVirtObject2.PostPopSize = -lOperand2ByteSize
                                            End If
                                        Else
                                            If oIntermediate.Operand2.ObjectType <> otVirt Then
                                                oVirtObject2.PrePopSize = lOperand1ByteSize
                                                oVirtObject2.PostPopSize = -lOperand1ByteSize
                                            Else
                                                oVirtObject2.PrePopSize = lOperand1ByteSize + lOperand2ByteSize
                                                oVirtObject2.PostPopSize = -2 * lOperand2ByteSize - lOperand1ByteSize
                                            End If
                                        End If
                                        
                                    Case opEqual, opNotEqual, opLessThan, opGreaterThan, opLessThanEqual, opGreaterThanEqual
                                        oVirtClass.Ranges.Add RangeFactory.Create(0, 1)
                                        
                                    Case opAnd, opOr, opEor
                                        oVirtClass.Ranges.Add RangeFactory.Create(0, Largest(lOperand1Ending, lOperand2Ending))
                                End Select
                                
                                Set oVirtObject = New clsObjectVirt
                                Set oVirtObject.UnitClass = oVirtClass
                                oVirtObject.IncreaseReferenceCount
                                Scope.moObjects.AddObject oVirtObject
                                Set oIntermediate.Operand3 = oVirtObject
                                Set oPreviousObject = oVirtObject
                                Set DoExpression = oPreviousObject
                                oVirtObject.PrePopSize = IIf(oIntermediate.Operand1.ObjectType = otVirt, oIntermediate.Operand1.Hydrate.UnitClass.Ranges.Size, 0) + IIf(oIntermediate.Operand2.ObjectType = otVirt, oIntermediate.Operand2.Hydrate.UnitClass.Ranges.Size, 0)
                                
                                Set oVirtObject2 = Nothing
                        End Select
                        moIntermediates.Add oIntermediate
                    End If
                    
                End If
                bTerm = Not bTerm
            Case False ' Operator
                Select Case oSubTree(1).Index
                    Case 1 ' none
                        opOperator = opNone
                    Case 2 ' add
                        opOperator = opAdd
                    Case 3 ' subtract
                        opOperator = opSub
                    Case 4 ' multiply
                        opOperator = opMultiply
                    Case 5 ' divide
                        opOperator = opDivide
                    Case 6 ' modulus
                        opOperator = opModulus
                    Case 7 ' not equal
                        opOperator = opNotEqual
                    Case 8 ' less than equal
                        opOperator = opLessThanEqual
                    Case 9 ' greater than equal
                        opOperator = opGreaterThanEqual
                    Case 10 ' equal
                        opOperator = opEqual
                    Case 11 ' less than
                        opOperator = opLessThan
                    Case 12 ' greater than
                        opOperator = opGreaterThan
                    Case 13 ' and
                        opOperator = opAnd
                    Case 14 ' or
                        opOperator = opOr
                    Case 15 ' eor
                        opOperator = opEor
                End Select
                bTerm = Not bTerm
        End Select
    Next

    If Not bChild Then
        oPreviousObject.PostPopSize = IIf(TypeOf oPreviousObject Is clsObjectVirt, oPreviousObject.UnitClass.Ranges(0).CalcBytes, 0)
    End If
    For lIndex = 0 To lTargetsCount - 1
        Set oIntermediate = New clsIntermediate
        moIntermediates.Add oIntermediate.Create(opCopy, oTargets(lIndex), oPreviousObject)
    Next
End Function

Private Function DoReturnStatement(oTree As SaffronTree)
    Dim oIntermediate As New clsIntermediate
    
    moIntermediates.Add oIntermediate.Create(opCopy, Scope.moReturnObject, DoExpression(oTree(1)))
End Function

Private Function DoFunctionStatement(oTree As SaffronTree)
    Dim oFunctionObject As New clsObjectFunc

    Dim oIntermediate1 As New clsIntermediate
    Dim oIntermediate2 As New clsIntermediate
    Dim oIntermeidate3 As New clsIntermediate
    
    Dim oSubSyntaxTree As New clsSyntaxTree
        
    oFunctionObject.Identifier = oTree(1).Text
    oFunctionObject.UniqueIdentifier = "function_" & oTree(1).Text & "_" & Scope.ScopeIdentifier
    If oTree(2).Index = 1 Then
        oFunctionObject.HasFunctionAddress = True
        oFunctionObject.FunctionAddress = Base16(oTree(2).Text)
    End If
    If oTree(3).Index = 1 Then
        If oTree(3)(1).Index = 1 Then ' identifier
            Set oFunctionObject.UnitClass = Scope.GetClassByIdentifier(oTree(3)(1).Text)
            If oFunctionObject.UnitClass Is Nothing Then
                MsgBox "Class " & oTree(3)(1).Text & " not found on line " & Scope.FindLineNumber(oTree(3)(1).TextStart)
                End
            End If
        Else ' fields
            Set oFunctionObject.UnitClass = GetCreateRanges(oTree(3)(1)(1)(1))
        End If
    End If
    
    Set oSubSyntaxTree.Scope = New clsScope
    Set oSubSyntaxTree.Scope.ParentScope = Scope
    Set oSubSyntaxTree.Scope.moReturnObject = oFunctionObject
    Scope.moChildScopes.Add oSubSyntaxTree.Scope
    
    oSubSyntaxTree.DoStatements oTree(4)
    
    oFunctionObject.Intermediates.Add oIntermediate1.Create(opFunction, oFunctionObject)
    oFunctionObject.Intermediates.MergeIntermediates oSubSyntaxTree.moIntermediates
    oFunctionObject.Intermediates.Add oIntermediate2.Create(opEndFunction)
    Set oFunctionObject.Intermediates.Scope = oSubSyntaxTree.Scope
    
    Scope.moObjects.AddObject oFunctionObject
End Function

Private Function DoIfStatement(oTree As SaffronTree)
    Dim oIntermediate1 As New clsIntermediate
    Dim oIntermediate2 As New clsIntermediate
    Dim oIntermediate3 As New clsIntermediate
    Dim oSubSyntaxTree As New clsSyntaxTree
    
    moIntermediates.Add oIntermediate1.Create(opIsZero, DoExpression(oTree(1)))
    moIntermediates.Add oIntermediate2.Create(opIf, , , , , osUseZFlag + osFlagSet)
    
    Set oSubSyntaxTree.Scope = Scope
    oSubSyntaxTree.DoStatements oTree(2)
    moIntermediates.MergeIntermediates oSubSyntaxTree.moIntermediates

    moIntermediates.Add oIntermediate3.Create(opEndIf)
End Function

Private Function DoWhileStatement(oTree As SaffronTree)
    Dim oIntermediate1 As New clsIntermediate
    Dim oIntermediate2 As New clsIntermediate
    Dim oIntermediate3 As New clsIntermediate
    Dim oIntermediate4 As New clsIntermediate
    Dim oSubSyntaxTree As New clsSyntaxTree
    
    moIntermediates.Add oIntermediate1.Create(opStartWhile)
    moIntermediates.Add oIntermediate2.Create(opIsZero, DoExpression(oTree(1)))  ' condition
    moIntermediates.Add oIntermediate3.Create(opWhile, , , , , osUseZFlag + osFlagSet)
    
    Set oSubSyntaxTree.Scope = Scope
    oSubSyntaxTree.DoStatements oTree(2)
    moIntermediates.MergeIntermediates oSubSyntaxTree.moIntermediates
    
    moIntermediates.Add oIntermediate4.Create(opEndWhile)
End Function
        
Private Function DoUntilStatement(oTree As SaffronTree)
    Dim oIntermediate1 As New clsIntermediate
    Dim oIntermediate2 As New clsIntermediate
    Dim oIntermediate3 As New clsIntermediate
    Dim oIntermediate4 As New clsIntermediate
    Dim oSubSyntaxTree As New clsSyntaxTree
    
    moIntermediates.Add oIntermediate1.Create(opStartUntil)
    moIntermediates.Add oIntermediate2.Create(opIsZero, DoExpression(oTree(1)))  ' condition
    moIntermediates.Add oIntermediate3.Create(opUntil, , , , , osUseZFlag)
    
    Set oSubSyntaxTree.Scope = Scope
    oSubSyntaxTree.DoStatements oTree(2)
    moIntermediates.MergeIntermediates oSubSyntaxTree.moIntermediates
    
    moIntermediates.Add oIntermediate4.Create(opEndUntil)
End Function

Private Function DoDoWhileStatement(oTree As SaffronTree)
    Dim oIntermediate1 As New clsIntermediate
    Dim oIntermediate2 As New clsIntermediate
    Dim oIntermediate3 As New clsIntermediate
    Dim oSubSyntaxTree As New clsSyntaxTree
    
    moIntermediates.Add oIntermediate1.Create(opStartDoWhile)

    Set oSubSyntaxTree.Scope = Scope
    oSubSyntaxTree.DoStatements oTree(1)
    moIntermediates.MergeIntermediates oSubSyntaxTree.moIntermediates
    
    moIntermediates.Add oIntermediate2.Create(opIsZero, DoExpression(oTree(2)))  ' condition
    
    moIntermediates.Add oIntermediate3.Create(opEndDoWhile, , , , , osUseZFlag)
End Function

Private Function DoDoUntilStatement(oTree As SaffronTree)
    Dim oIntermediate1 As New clsIntermediate
    Dim oIntermediate2 As New clsIntermediate
    Dim oIntermediate3 As New clsIntermediate
    Dim oSubSyntaxTree As New clsSyntaxTree
    
    moIntermediates.Add oIntermediate1.Create(opStartDoUntil)

    Set oSubSyntaxTree.Scope = Scope
    oSubSyntaxTree.DoStatements oTree(1)
    moIntermediates.MergeIntermediates oSubSyntaxTree.moIntermediates
    
    moIntermediates.Add oIntermediate2.Create(opIsZero, DoExpression(oTree(2)))  ' condition
    moIntermediates.Add oIntermediate3.Create(opEndDoUntil, , , , , osUseZFlag Or osFlagSet)
End Function


Private Function DoFunctionCallOrIndex(oTree As SaffronTree) As clsObject
    Dim oFunctionObject As clsObject
    Dim lIndex As Long
    Dim oParam As clsObject

    Dim lCount As Long
    Dim oIntermediate As clsIntermediate
    
    Dim oObject As clsObject
    
    Set oObject = Scope.GetObjectByIdentifier(oTree(1).Text)
    
    If oObject Is Nothing Then
        MsgBox "Object or function: " & oTree(1).Text & " not found on line " & Scope.FindLineNumber(oTree(1).TextStart), vbOKOnly
        End
    End If
    
    Select Case oObject.ObjectType
        Case otFunc
            DoFunctionCall oTree(2), oObject
            Set DoFunctionCallOrIndex = oObject
        Case Else
            Set DoFunctionCallOrIndex = DoObjectIndex(oTree(2), oObject)
    End Select
End Function

Private Function DoFunctionCall(oTree As SaffronTree, oFunctionObject As clsObject)
    Dim oExpressionObjects() As clsObject
    Dim lIndex As Long
    Dim oParam As clsObject
    Dim oIntermediate As New clsIntermediate
    
    If oTree.SubTree.Count <> oFunctionObject.Intermediates.Scope.moParamObjects.Count Then
        MsgBox "Function " & oTree(1).Text & " expects " & oFunctionObject.Intermediates.Scope.moParamObjects.Count & " parameters at line " & Scope.FindLineNumber(oTree(1).TextStart)
        End
    End If
    
    For lIndex = 1 To oTree.SubTree.Count
        ReDim Preserve oExpressionObjects(lIndex - 1)
        Set oExpressionObjects(lIndex - 1) = DoExpression(oTree(lIndex))
    Next
    
    For lIndex = 0 To oFunctionObject.Intermediates.Scope.moParamObjects.Count - 1
        Set oParam = oFunctionObject.Intermediates.Scope.moParamObjects.Objects(lIndex)
        Set oIntermediate = New clsIntermediate
        If oParam.HasRegisterTarget Then
            oIntermediate.Create opCopyToRegister, oParam, oExpressionObjects(lIndex)
        Else
            oIntermediate.Create opCopy, oParam, oExpressionObjects(lIndex)
        End If
        moIntermediates.Add oIntermediate
    Next
End Function

Private Function DoObjectIndex(oTree As SaffronTree, oObject As clsObject) As clsObject
    Dim oExpressionObjects() As clsObject
    Dim lIndex As Long
    Dim oIndexedObject As clsObject
    Dim lObjectIndexValue As Long
    
    If oTree.SubTree.Count <> 1 Then
        MsgBox "Syntax error " & oObject.Identifier & " expects 1 argument at line " & Scope.FindLineNumber(oTree(1).TextStart)
        End
    End If
    
    For lIndex = 1 To oTree.SubTree.Count
        ReDim Preserve oExpressionObjects(lIndex - 1)
        Set oExpressionObjects(lIndex - 1) = DoExpression(oTree(lIndex))
    Next
    
    If oExpressionObjects(0).ObjectType = otConst Then
        Set oIndexedObject = ObjectFactory.Create(otVirt, ClassFactory.Create("", oObject.UnitClass.Ranges(lObjectIndexValue)))

        oIndexedObject.IdentifierOffset = oObject.UnitClass.Ranges.Size(atByte, oExpressionObjects(0).UnitClass.Ranges(0).Constant - 1)
    Else
        MsgBox "Object index for " & oObject.Identifier & " must be constant at line " & Scope.FindLineNumber(oTree(1).TextStart)
    End If
    
    Set DoObjectIndex = oIndexedObject
End Function









