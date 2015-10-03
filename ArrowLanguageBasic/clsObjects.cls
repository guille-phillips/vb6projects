VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsObjects"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Count As Long
Private moObjects() As clsIObject
Private moAssemblyOps As New clsAssemblyOps

Public Function AddObject(oMember As clsIObject)
    ReDim Preserve moObjects(Count)
    Set moObjects(Count) = oMember
    Count = Count + 1
End Function

Public Property Get Objects(ByVal lIndex As Long) As clsIObject
Attribute Objects.VB_UserMemId = 0
    Set Objects = moObjects(lIndex)
End Property

Public Property Set Objects(ByVal lIndex As Long, oMember As clsIObject)
    Set moObjects(lIndex) = oMember
End Property


Public Function GetByIdentifier(ByVal sIdentifier As String, Optional ByVal otObjectType As ObjectTypes = otAny) As Object
    Dim lIndex As Long
    Dim oObject As Object
    
    For lIndex = 0 To Count - 1
        If otObjectType = otAny Or moObjects(lIndex).ObjectType = otObjectType Then
            Set oObject = moObjects(lIndex).Hydrate
            If oObject.Identifier = sIdentifier Then
                Set GetByIdentifier = oObject
                Exit Function
            End If
        End If
    Next
End Function
