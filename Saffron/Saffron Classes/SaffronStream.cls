VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SaffronStream"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Property Let Text(sTextString As String)
    TextString = sTextString
    TextPosition = 1
    LenTextString = Len(sTextString)
End Property

Public Property Get Text() As String
    Text = TextString
End Property

Public Sub Reset()
    TextPosition = 1
End Sub

Public Property Let Position(lPosition As Long)
    TextPosition = lPosition
End Property

Public Property Get Position() As Long
    Position = TextPosition
End Property

Public Property Set External(oExternal As Object)
    Set ExternalClass = oExternal
End Property

