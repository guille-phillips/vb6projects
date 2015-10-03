VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRanges"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Count As Long
Private moRanges() As clsRange

Public AccessType As AccessTypes

Public Function Add(oMember As clsRange)
    ReDim Preserve moRanges(Count)
    Set moRanges(Count) = oMember
    Count = Count + 1
End Function


Public Property Get Ranges(ByVal lIndex As Long) As clsRange
Attribute Ranges.VB_UserMemId = 0
    Set Ranges = moRanges(lIndex)
End Property

Public Property Set Ranges(ByVal lIndex As Long, oMember As clsRange)
    Set moRanges(lIndex) = oMember
End Property

Public Property Get Size(Optional ByVal atAccessType As AccessTypes = atDefault, Optional ByVal lFinalRange As Long = -2) As Currency
    Dim lIndex As Long
    Dim cSize As Currency
    
    If atAccessType = atDefault Then
        atAccessType = AccessType
    End If
    
    Select Case atAccessType
        Case atCompact
            cSize = 1
        Case atBit, atByte
            cSize = 0
    End Select
    
    If lFinalRange = -2 Then
        lFinalRange = Count - 1
    End If
    
    For lIndex = 0 To lFinalRange
        Select Case atAccessType
            Case atCompact
                cSize = cSize * moRanges(lIndex).StorageRange
            Case atBit
                cSize = cSize + moRanges(lIndex).StorageBits
            Case atByte
                cSize = cSize + moRanges(lIndex).StorageBytes
        End Select
    Next
    
    Size = cSize
End Property

Private Sub Class_Initialize()
    AccessType = atByte
End Sub
