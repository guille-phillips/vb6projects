VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SubTrees"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private oItems As New Collection

Public Sub Add(oTree As SaffronTree)
    oItems.Add oTree
End Sub

Public Property Get Item(ByVal vIndex As Variant) As SaffronTree
Attribute Item.VB_UserMemId = 0
    Set Item = oItems(vIndex)
End Property

Public Property Get Count() As Long
    Count = oItems.Count
End Property

Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    Set NewEnum = oItems.[_NewEnum]
End Property

Public Sub Remove(lIndex As Long)
    oItems.Remove lIndex
End Sub
