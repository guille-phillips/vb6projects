VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsExtents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Count As Long
Private moExtents() As clsExtent
Private mlSize As Long

Public Function AddExtent(oMember As clsExtent)
    ReDim Preserve moExtents(Count)
    Set moExtents(Count) = oMember
    
    If Count = 0 Then
        mlSize = oMember.Size
    Else
        mlSize = mlSize * oMember.Size
    End If
    Count = Count + 1
End Function

Public Property Get Extents(ByVal lIndex As Long) As clsExtent
    Set Extents = moExtents(lIndex)
End Property

Public Property Set Extents(ByVal lIndex As Long, oMember As clsExtent)
    Set moExtents(lIndex) = oMember
End Property

Public Property Get Size() As Long
    Size = mlSize
End Property

Private Sub Class_Initialize()
    mlSize = 1
End Sub
