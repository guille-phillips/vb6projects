VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsStatements"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Count As Long
Private moStatements() As clsIStatement

Public Function AddStatement(oMember As clsIStatement)
    ReDim Preserve moStatements(Count)
    Set moStatements(Count) = oMember
    Count = Count + 1
End Function

Public Property Get Statements(ByVal lIndex As Long) As clsIStatement
    Set Statements = moStatements(lIndex)
End Property

Public Property Set Statements(ByVal lIndex As Long, oMember As clsIStatement)
    Set moStatements(lIndex) = oMember
End Property

