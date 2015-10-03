VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Multiply"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IAlgebraic

Private oOperandA As IAlgebraic
Private oOperandB As IAlgebraic

Private ntNumberType As NumberTypes

Public Property Set IAlgebraic_OperandA(oOperand As IAlgebraic)
    Set oOperandA = oOperand
End Property

Public Property Get IAlgebraic_OperandA() As IAlgebraic
'    Set OperandA = oOperandA
End Property

Public Property Set IAlgebraic_OperandB(oOperand As IAlgebraic)
'    Set oOperandB = oOperand
End Property

Public Property Get IAlgebraic_OperandB() As IAlgebraic
'    Set OperandAB = oOperandB
End Property


Private Property Let IAlgebraic_NumberType(RHS As NumberTypes)
    ntNumberType = RHS
End Property

Private Property Get IAlgebraic_NumberType() As NumberTypes
    IAlgebraic_NumberType = ntNumberType
End Property

Private Function IAlgebraic_RealValue() As Double
'    RealValue = Operand1.RealValue * Operand2.RealValue
End Function

Private Function IAlgebraic_ImaginaryValue() As Double
'    ImaginaryValue = Operand1.ImaginaryValue ^ Operand2.ImaginaryValue
End Function



Private Sub Class_Initialize()
    ntNumberType = NumberTypes.Multiply
End Sub

Private Function IAlgebraic_Simplify() As IAlgebraic

End Function
