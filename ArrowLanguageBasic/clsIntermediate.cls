VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIntermediate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Operator As ILOperators
Public Operand1 As clsIObject
Public Operand2 As clsIObject
Public Operand3 As clsIObject
Public Operand4 As clsIObject

Public Options As Long

Public MemoryAddress As Long

Public Function Create(ByVal ilOperator As ILOperators, Optional ByVal oOperand1 As clsIObject, Optional ByVal oOperand2 As clsIObject, Optional ByVal oOperand3 As clsIObject, Optional ByVal oOperand4 As clsIObject, Optional ByVal lOptions As Long) As clsIntermediate
    Set Create = New clsIntermediate
    
    Create.Operator = ilOperator
    Set Create.Operand1 = oOperand1
    Set Create.Operand2 = oOperand2
    Set Create.Operand3 = oOperand3
    Set Create.Operand4 = oOperand4
    Create.Options = lOptions
End Function

Public Property Get Operand(ByVal lIndex As Long) As Object
    Select Case lIndex
        Case 1
            Set Operand = Operand1.Hydrate
        Case 2
            Set Operand = Operand2.Hydrate
        Case 3
            Set Operand = Operand3.Hydrate
        Case 4
            Set Operand = Operand4.Hydrate
    End Select
End Property
