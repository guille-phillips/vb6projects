VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IAlgebraic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum NumberTypes
    Number
    Plus
    Multiply
    power
End Enum

Public Property Set OperandA(oOperand As IAlgebraic)
'
End Property

Public Property Get OperandA() As IAlgebraic
'
End Property

Public Property Set OperandB(oOperand As IAlgebraic)
'
End Property

Public Property Get OperandB() As IAlgebraic
'
End Property


Public Property Let NumberType(lType As NumberTypes)
'
End Property

Public Property Get NumberType() As NumberTypes
'
End Property

Public Function RealValue() As Double
'
End Function

Public Function ImaginaryValue() As Double
'
End Function

Public Function Simplify() As IAlgebraic
'
End Function

