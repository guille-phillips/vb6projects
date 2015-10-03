VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements clsIObject

Public Identifier As String
Public UniqueIdentifier As String
Public UnitClass As clsClass

Public ObjectType As ObjectTypes

Public MemAddress As Long

Public ReferenceCount As Long
Public HasHardAddress As Boolean
Public HasRegisterTarget As Boolean
Public RegisterTarget As RegisterTargets
Public HasFunctionAddress As Boolean
Public FunctionAddress As Long

Public Intermediates As New clsIntermediates

Public PrePopSize As Long
Public PostPopSize As Long

Public IdentifierOffset As Long

Private Sub Class_Initialize()
    MemAddress = -1
End Sub

Public Sub IncreaseReferenceCount()
    ReferenceCount = ReferenceCount + 1
End Sub

Public Function Create(ByVal otObjectType As ObjectTypes, ByVal oUnitClass As clsClass, Optional ByVal sIdentifier As String) As clsObject
    Set Create = New clsObject
    Create.ObjectType = otObjectType
    Set Create.UnitClass = oUnitClass
    Create.Identifier = sIdentifier
End Function


Private Function clsIObject_Hydrate() As Object
    Set clsIObject_Hydrate = Me
End Function

Private Property Get clsIObject_ObjectType() As ObjectTypes
    clsIObject_ObjectType = otVar
End Property
