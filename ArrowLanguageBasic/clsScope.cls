VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsScope"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public ParentScope As clsScope

Public moClasses As New clsClasses
Public moObjects As New clsObjects
Public moParamObjects As New clsObjects
Public moReturnObject As New clsObjectFunc
Public moChildScopes As New clsScopes

Public moSyntaxTree As clsSyntaxTree

Public ScopeIdentifier As String

Private mlVirtualStack As Long

Public Function PushVirtual(ByVal lSize As Long)
    mlVirtualStack = mlVirtualStack + lSize
End Function

Public Function PullVirtual(ByVal lSize As Long)
    mlVirtualStack = mlVirtualStack - lSize
End Function

Public Function GetObjectByIdentifier(ByVal sIdentifier As String) As Object
    Dim oObject As Object
    
    Set oObject = moObjects.GetByIdentifier(sIdentifier)
    If oObject Is Nothing Then
        If Not ParentScope Is Nothing Then
            Set oObject = ParentScope.GetObjectByIdentifier(sIdentifier)
        End If
    End If
    
    Set GetObjectByIdentifier = oObject
End Function

Public Function GetClassByIdentifier(ByVal sIdentifier As String) As clsClass
    Dim oClass As clsClass
    
    Set oClass = moClasses.GetByIdentifier(sIdentifier)
    If oClass Is Nothing Then
        If Not ParentScope Is Nothing Then
            Set oClass = ParentScope.GetClassByIdentifier(sIdentifier)
        End If
    End If
    
    Set GetClassByIdentifier = oClass
End Function

Private Sub Class_Initialize()
    Dim lIndex As Long
    
    For lIndex = 1 To 5
        ScopeIdentifier = ScopeIdentifier & Chr$(Int(Rnd() * 26) + 65)
    Next
End Sub

Private Function RandomIdent() As String
    Dim lIndex As Long
    
    For lIndex = 1 To 5
        RandomIdent = RandomIdent & Chr$(Int(Rnd() * 26) + 65)
    Next
End Function

Public Function CompileFunctions(ByVal bHasFunctionAddress As Boolean) As clsIntermediates
    Dim oIntermediates As New clsIntermediates
    Dim lIndex As Long
    Dim oObjectFunc As clsObjectFunc
    Dim oSeparator As New clsIntermediate
    
    oSeparator.Create opSeparator
    For lIndex = 0 To moObjects.Count - 1
        If moObjects.Objects(lIndex).ObjectType = otFunc Then
            Set oObjectFunc = moObjects.Objects(lIndex)
            If oObjectFunc.HasFunctionAddress = bHasFunctionAddress And oObjectFunc.ReferenceCount > 0 Then
                oIntermediates.Add oSeparator
                oIntermediates.MergeIntermediates oObjectFunc.Intermediates
            End If
        End If
    Next
    
    For lIndex = 0 To moChildScopes.Count - 1
        oIntermediates.MergeIntermediates moChildScopes.Scopes(lIndex).CompileFunctions(bHasFunctionAddress)
    Next
    
    Set CompileFunctions = oIntermediates
End Function



Public Function CompileVirtualStorage(Optional lVirtBaseAddress As Long = &H70) As clsIntermediates
    Dim lIndex As Long
    Dim oObject As clsIObject
    Dim oObjectVirt As clsObjectVirt
    Dim lOffset As Long
    Dim lHighestOffset As Long
    Dim oVirtualObjectBase As New clsObject
    Dim oIntermediate As clsIntermediate
    Dim oIntermediates As New clsIntermediates
    
    For lIndex = 0 To moObjects.Count - 1
        Set oObject = moObjects.Objects(lIndex)
        If oObject.ObjectType = otVirt Then
            Set oObjectVirt = oObject.Hydrate
            lOffset = lOffset - oObjectVirt.PrePopSize
            oObjectVirt.UniqueIdentifier = "vir_" & ScopeIdentifier & "+" & lOffset & "d"
            lOffset = lOffset + oObjectVirt.UnitClass.Ranges.Size
            lOffset = lOffset - oObjectVirt.PostPopSize
            
            If lOffset > lHighestOffset Then
                lHighestOffset = lOffset
            End If

        End If
    Next
    

    oVirtualObjectBase.MemAddress = lVirtBaseAddress
    oVirtualObjectBase.UniqueIdentifier = "vir_" & ScopeIdentifier
    Set oIntermediate = New clsIntermediate
    oIntermediate.Operator = opObject
    Set oIntermediate.Operand1 = oVirtualObjectBase
    oIntermediates.Add oIntermediate

    For lIndex = 0 To moChildScopes.Count - 1
        oIntermediates.MergeIntermediates moChildScopes.Scopes(lIndex).CompileVirtualStorage(lVirtBaseAddress + lHighestOffset)
    Next

    Set CompileVirtualStorage = oIntermediates
End Function

Public Function CompileStorage(Optional lVarAddress As Long = &H7000) As clsIntermediates
    Dim lIndex As Long
    Dim oObject As clsIObject
    Dim lIndex2 As Long
    Dim lIndex3 As Long
    Dim oIntermediates As New clsIntermediates
    Dim oIntermediate As clsIntermediate
    Dim lAddress As Long
    Dim oObjectVar As clsObjectVar
    Dim oObjectConst As clsObjectConst
    Dim oObjectFunc As clsObjectFunc
    Dim oObjectParam As clsObjectParam
    
    For lIndex = 0 To moObjects.Count - 1
        Set oObject = moObjects.Objects(lIndex)
        Select Case oObject.ObjectType
            Case otVar
                Set oObjectVar = oObject.Hydrate
                oObjectVar.UniqueIdentifier = "var_" & oObjectVar.Identifier & "_" & ScopeIdentifier
                If oObjectVar.ReferenceCount > 0 Then
                    If Not oObjectVar.HasHardAddress Then
                        oObjectVar.MemAddress = lVarAddress
                        lVarAddress = lVarAddress + oObjectVar.UnitClass.Ranges(0).CalcBytes
                    End If
                End If
            Case otParam
                Set oObjectParam = oObject.Hydrate
                oObjectParam.UniqueIdentifier = "var_" & oObjectParam.Identifier & "_" & ScopeIdentifier
                If Not oObjectParam.HasRegisterTarget Then
                    If oObjectParam.ReferenceCount > 0 Then
                        If Not oObjectParam.HasHardAddress Then
                            oObjectParam.MemAddress = lVarAddress
                            lVarAddress = lVarAddress + oObjectParam.UnitClass.Ranges(0).CalcBytes
                        End If
                    End If
                End If
            Case otConst
                Set oObjectConst = oObject.Hydrate
                If oObjectConst.ReferenceCount > 0 Then
                    oObjectConst.UniqueIdentifier = "const_" & RandomIdent & "_" & ScopeIdentifier
                    oObjectConst.MemAddress = lVarAddress
                    lVarAddress = lVarAddress + oObjectConst.UnitClass.Ranges(0).CalcBytes
                End If
            Case otFunc
                Set oObjectFunc = oObject.Hydrate
                If oObjectFunc.ReferenceCount > 0 Then
                    oObjectFunc.MemAddress = lVarAddress
                    lVarAddress = lVarAddress + oObjectFunc.UnitClass.Ranges(0).CalcBytes
                End If
        End Select
    Next
    
    For lIndex = 0 To moObjects.Count - 1
        Set oObject = moObjects.Objects(lIndex)
        Select Case oObject.ObjectType
            Case otVar, otConst, otFunc
                If oObject.Hydrate.ReferenceCount > 0 Then
                    Set oIntermediate = New clsIntermediate
                    oIntermediate.Operator = opObject
                    Set oIntermediate.Operand1 = oObject
                    oIntermediates.Add oIntermediate
                End If
            Case otParam
                If Not oObject.Hydrate.HasRegisterTarget Then
                    If oObject.Hydrate.ReferenceCount > 0 Then
                        Set oIntermediate = New clsIntermediate
                        oIntermediate.Operator = opObject
                        Set oIntermediate.Operand1 = oObject
                        oIntermediates.Add oIntermediate
                    End If
                End If
        End Select
    Next
    
    For lIndex = 0 To moChildScopes.Count - 1
        oIntermediates.MergeIntermediates moChildScopes.Scopes(lIndex).CompileStorage(lVarAddress)
    Next
    
    Set CompileStorage = oIntermediates
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


Public Function FindLineNumber(ByVal lPosition As Long) As Long
    Dim lIndex As Long
    
    If ParentScope Is Nothing Then
        FindLineNumber = moSyntaxTree.FindLineNumber(lPosition)
    Else
        FindLineNumber = ParentScope.FindLineNumber(lPosition)
    End If
End Function
