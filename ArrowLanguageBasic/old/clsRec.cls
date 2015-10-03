VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Implements clsIMember

Public Identifier As String
Public Extent As New clsExtent

Private msIdentifier As String

Private Property Let clsIMember_Identifier(ByVal sIdentifier As String)
    msIdentifier = sIdentifier
End Property

Private Property Get clsIMember_Identifier() As String
    clsIMember_Identifier = msIdentifier
End Property

Private Property Get clsIMember_MemberType() As Long
    clsIMember_MemberType = REC
End Property

Private Property Get clsIMember_Size() As Long
'
End Property
