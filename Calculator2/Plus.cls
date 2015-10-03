VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Plus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IAlgebraic

Public Operand1 As IAlgebraic
Public Operand2 As IAlgebraic

Private ntNumberType As NumberTypes

Private Property Get IAlgebraic_OperandA() As IAlgebraic

End Property

Public Property Set IAlgebraic_OperandA(oOperand As IAlgebraic)
    'Set OperandA = oOperand
End Property

Public Property Get OperandA(oOperand As IAlgebraic)
    'Set OperandA = oOperandA
End Property

Public Property Set OperandB(oOperand As IAlgebraic)
    'Set OperandB = oOperand
End Property

Public Property Get OperandB() As IAlgebraic
   ' Set OperandAB = oOperandB
End Property

Private Property Let IAlgebraic_NumberType(RHS As NumberTypes)
    ntNumberType = RHS
End Property

Private Property Get IAlgebraic_NumberType() As NumberTypes
    IAlgebraic_NumberType = ntNumberType
End Property

Private Function IAlgebraic_ImaginaryValue() As Double
    'ImaginaryValue = Operand1.ImaginaryValue + Operand2.ImaginaryValue
End Function

Private Property Set IAlgebraic_OperandB(RHS As IAlgebraic)

End Property

Private Property Get IAlgebraic_OperandB() As IAlgebraic

End Property

Private Function IAlgebraic_RealValue() As Double
    'RealValue = Operand1.RealValue + Operand2.RealValue
End Function

Private Function IAlgebraic_Simplify() As IAlgebraic
    Dim oOperand1 As IAlgebraic
    Dim oOperand2 As IAlgebraic
    
    Set oOperand1 = Operand1.Simplify
    Set oOperand2 = Operand2.Simplify
    
    ' Simple add
'    If Operand1 Is Number And Operand2 Is Number Then
'        Set IAlgebraic_Simplify = oOperand1.RealValue + oOperand2.RealValue
'    End If
    
End Function

Private Sub Class_Initialize()
    ntNumberType = NumberTypes.Plus
End Sub

