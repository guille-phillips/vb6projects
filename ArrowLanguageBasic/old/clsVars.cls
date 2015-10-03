VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsVars"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Count As Long
Private moVars() As clsVar

Public Function AddVar(oMember As clsVar)
    ReDim Preserve moVars(Count)
    Set moVars(Count) = oMember
    Count = Count + 1
End Function

Public Property Get Vars(ByVal lIndex As Long) As clsVar
    Set Vars = moVars(lIndex)
End Property

Public Property Set Vars(ByVal lIndex As Long, oMember As clsVar)
    Set moVars(lIndex) = oMember
End Property


Public Function GetByIdentifier(ByVal sIdentifier As String) As clsVars
    Dim lIndex As Long
    
    For lIndex = 0 To Count - 1
        If moVars(lIndex).Identifier = sIdentifier Then
            Set GetByIdentifier = moVars(lIndex)
            Exit Function
        End If
    Next
End Function
