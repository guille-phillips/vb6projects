VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsExtent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Count As Long
Private moMembers() As clsIMember

Private mlSize As Long

Public Function AddMember(oMember As clsIMember)
    ReDim Preserve moMembers(Count)
    Set moMembers(Count) = oMember
    mlSize = mlSize + oMember.Size
    Count = Count + 1
End Function

Public Property Get Members(ByVal lIndex As Long) As clsIMember
    Set Members = moMembers(lIndex)
End Property

Public Property Set Members(ByVal lIndex As Long, oMember As clsIMember)
    Set moMembers(lIndex) = oMember
End Property

Public Property Get Size() As Long
    Size = mlSize
End Property
