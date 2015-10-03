VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsScopes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Count As Long
Private moScopes() As clsScope

Public Function Add(oMember As clsScope)
    ReDim Preserve moScopes(Count)
    Set moScopes(Count) = oMember
    Count = Count + 1
End Function

Public Property Get Scopes(ByVal lIndex As Long) As clsScope
Attribute Scopes.VB_UserMemId = 0
    Set Scopes = moScopes(lIndex)
End Property

Public Property Set Scopes(ByVal lIndex As Long, oMember As clsScope)
    Set moScopes(lIndex) = oMember
End Property
