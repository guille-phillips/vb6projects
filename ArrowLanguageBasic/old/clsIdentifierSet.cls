VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIdentifierSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Implements clsIMember

Private moIdentifierSet() As String

Private msIdentifier As String
Private mlSize As Long

Private Property Let clsIMember_Identifier(ByVal sIdentifier As String)
    msIdentifier = sIdentifier
End Property

Private Property Get clsIMember_Identifier() As String
    clsIMember_Identifier = msIdentifier
End Property

Private Property Get clsIMember_MemberType() As Long
    clsIMember_MemberType = IDENTIFIER_SET
End Property

Public Sub AddIdentifier(ByVal sIdentifier As String)
    ReDim Preserve moIdentifierSet(mlSize)
    moIdentifierSet(mlSize) = sIdentifier
    mlSize = mlSize + 1
End Sub

Public Property Get IdentifierSet(ByVal lIndex As Long) As String
    IdentifierSet = moIdentifierSet(lIndex)
End Property

Public Property Let IdentifierSet(ByVal lIndex As Long, sIdentifier As String)
    moIdentifierSet(lIndex) = oMember
End Property

Private Property Get clsIMember_Size() As Long
    clsIMember_Size = mlSize
End Property
