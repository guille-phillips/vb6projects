VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SaffronTree"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public SubTree As New SubTrees
Attribute SubTree.VB_VarMemberFlags = "200"
Attribute SubTree.VB_VarUserMemId = 0
Public Index As Long
Public ParentIndex As Long

Public TextStart As Long
Public TextEnd As Long

Public Error As Boolean
Public Leaf As Boolean

Public Property Get FullText() As String
    FullText = Mid$(TextString, TextStart, TextEnd - TextStart + 1)
End Property

Public Property Get Text() As String
    Dim oTree As SaffronTree
    
    If Not Leaf Then
        If SubTree.Count > 0 Then
            For Each oTree In SubTree
                Text = Text & oTree.Text
            Next
        End If
    Else
        If TextStart <= TextEnd Then
            Text = Mid$(TextString, TextStart, TextEnd - TextStart + 1)
        End If
    End If
End Property

