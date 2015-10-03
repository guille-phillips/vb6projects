VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAssemblyOp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Instruction As AssemblyOps
Public Operand As Variant
Public Operand2 As Variant
Public OperandSize As Long
Public OperandActualSize As Long
Public Register As AssemblyOpRegister
Public Mode As AssemblyOpMode
Public ModeRegister As AssemblyOpRegister
Public InterleaveCount As Long
Public InterleaveReverse As Boolean

Private Sub Class_Initialize()
    Instruction = ioNOP
    Operand = 0
    OperandSize = 1
    Register = irNone
    Mode = imImplied
    ModeRegister = irNone
    InterleaveCount = 1
    InterleaveReverse = False
End Sub
