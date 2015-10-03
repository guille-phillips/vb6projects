VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ListCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private ListItems() As ListItem
Private NotEmpty As Boolean

Public Property Get Item(index As Long) As ListItem
Attribute Item.VB_UserMemId = 0
    Set Item = ListItems(index)
End Property

Public Property Set Item(index As Long, oFile As ListItem)
    Set ListItems(index) = oFile
End Property

Public Function Add(Thing As ListItem)
    If NotEmpty Then
        ReDim Preserve ListItems(1 To UBound(ListItems) + 1) As ListItem
        Set ListItems(UBound(ListItems)) = Thing
    Else
        ReDim ListItems(1 To 1) As ListItem
        Set ListItems(1) = Thing
        NotEmpty = True
    End If
End Function

Public Property Get Count() As Long
    If NotEmpty Then
        Count = UBound(ListItems)
    End If
End Property

Public Function Remove(index As Long)
    Dim x As Long
    If index = 1 And UBound(ListItems) = 1 Then
        Erase ListItems()
        NotEmpty = False
    Else
        For x = index To UBound(ListItems)
            Set ListItems(x) = ListItems(x + 1)
        Next
        ReDim Preserve ListItems(1 To UBound(ListItems) - 1) As ListItem
    End If
End Function

Public Function RemoveListItem(Thing As ListItem)
    Dim x As Long
    Dim y As Long
    If ListItems(1) Is Thing And UBound(ListItems) = 1 Then
        Erase ListItems()
        NotEmpty = False
    Else
        For x = 1 To UBound(ListItems) - 1
            If Thing Is ListItems(x) Then
                For y = x To UBound(ListItems) - 1
                    Set ListItems(y) = ListItems(y + 1)
                Next
                ReDim Preserve ListItems(1 To UBound(ListItems) - 1) As ListItem
            End If
        Next
    End If
End Function

