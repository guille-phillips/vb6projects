VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Part"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Enum PartTypes
    ptAssignment
    ptSetVar
    ptParallelVar
    ptRange
    ptNumber
    ptSetDef
    ptEscape
    ptChar
    ptBlank
End Enum

Public Parallel As Boolean
Public VarName As String

Public PartSetDef As SetDef
Public ReferenceSet As SetDef

Public Range As Boolean
Public RangeStart As Long
Public RangeEnd As Long
Public LeadingZeroes As Boolean

Public Character As String

Public PartType As PartTypes

Public MemberLink As Long

Public Property Get Text(ByVal lIndex As Long) As String
    Select Case PartType
        Case ptAssignment
            Text = PartSetDef.Text(lIndex)
        Case ptSetVar
            If Not PartSetDef Is Nothing Then
                Text = PartSetDef.Text(lIndex)
            End If
        Case ptParallelVar
            Text = PartSetDef.Text(lIndex)
        Case ptSetDef
            Text = PartSetDef.Text(lIndex)
        Case ptRange
            If Not LeadingZeroes Then
                Text = lIndex + RangeStart
            Else
                Text = Format$(lIndex + RangeStart, String$(Len(CStr(RangeEnd)), "0"))
            End If
        Case ptNumber
            If Not LeadingZeroes Then
                Text = lIndex
            Else
                Text = Format$(lIndex, String$(Len(CStr(RangeStart)), "0"))
            End If
        Case ptEscape, ptChar, ptBlank
            Text = Character
    End Select
End Property

Public Property Get SetSize() As Long
    SetSize = 1
    Select Case PartType
        Case ptAssignment
            SetSize = PartSetDef.SetSize
        Case ptSetVar
            If Not PartSetDef Is Nothing Then
                SetSize = PartSetDef.SetSize
            End If
        Case ptParallelVar
            SetSize = PartSetDef.SetSize
        Case ptSetDef
            SetSize = PartSetDef.SetSize
        Case ptRange
            SetSize = RangeEnd - RangeStart + 1
        Case ptNumber
            SetSize = RangeStart + 1
        Case ptEscape, ptChar, ptBlank
            SetSize = 1
    End Select
End Property
