VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsClasses"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Count As Long
Private moClasses() As clsClass

Public Function Add(oMember As clsClass)
    ReDim Preserve moClasses(Count)
    Set moClasses(Count) = oMember
    Count = Count + 1
End Function

Public Property Get Classes(ByVal lIndex As Long) As clsClass
Attribute Classes.VB_UserMemId = 0
    Set Classes = moClasses(lIndex)
End Property

Public Property Set Classes(ByVal lIndex As Long, oMember As clsClass)
    Set moClasses(lIndex) = oMember
End Property


Public Function GetByIdentifier(ByVal sIdentifier As String) As clsClass
    Dim lIndex As Long
    
    For lIndex = 0 To Count - 1
        If moClasses(lIndex).Identifier = sIdentifier Then
            Set GetByIdentifier = moClasses(lIndex)
            Exit Function
        End If
    Next
End Function

