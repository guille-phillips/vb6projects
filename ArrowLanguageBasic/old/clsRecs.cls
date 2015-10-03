VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRecs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Count As Long
Private moRecs() As clsRec

Public Function AddRec(oMember As clsRec)
    ReDim Preserve moRecs(Count)
    Set moRecs(Count) = oMember
    Count = Count + 1
End Function

Public Property Get Recs(ByVal lIndex As Long) As clsRec
    Set Recs = moRecs(lIndex)
End Property

Public Property Set Recs(ByVal lIndex As Long, oMember As clsRec)
    Set moRecs(lIndex) = oMember
End Property


Public Function GetByIdentifier(ByVal sIdentifier As String) As clsRec
    Dim lIndex As Long
    
    For lIndex = 0 To Count - 1
        If moRecs(lIndex).Identifier = sIdentifier Then
            Set GetByIdentifier = moRecs(lIndex)
            Exit Function
        End If
    Next
End Function
