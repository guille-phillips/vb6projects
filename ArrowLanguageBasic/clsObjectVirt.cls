VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsObjectVirt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements clsIObject

Public UnitClass As clsClass
Public UniqueIdentifier As String
Public ReferenceCount As Long
Public PrePopSize As Long
Public PostPopSize As Long

Private Function clsIObject_Hydrate() As Object
    Set clsIObject_Hydrate = Me
End Function

Private Property Get clsIObject_ObjectType() As ObjectTypes
    clsIObject_ObjectType = otVirt
End Property

Public Sub IncreaseReferenceCount()
    ReferenceCount = ReferenceCount + 1
End Sub
