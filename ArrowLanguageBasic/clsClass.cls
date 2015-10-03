VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Identifier As String
Public Ranges As New clsRanges

Public AccessType As AccessTypes

'Private mlByteSize As Long
'Private mlBitSize As Long
'
'Private mbByteSizeSet As Boolean
'Private mbBitSizeSet As Boolean

Function Create(Optional ByVal sIdentifier As String, Optional ByVal oRange As clsRange) As clsClass
    Set Create = New clsClass
    Create.Identifier = sIdentifier
    If Not oRange Is Nothing Then
        Create.Ranges.Add oRange
    End If
End Function

'Public Property Get ByteSize() As Long
'    Dim cTotalRange As Currency
'
'    If mbByteSizeSet Then
'        ByteSize = mlByteSize
'    Else
'        cTotalRange = Range.Ending - Range.Starting
'
'        If cTotalRange > 0 Then
'            mlByteSize = Int(Log(cTotalRange) / Log(256) + 1)
'        Else
'            If Range.Starting = 0 Then
'                mlByteSize = 1
'            Else
'                mlByteSize = Int(Log(Range.Starting) / Log(256) + 1)
'            End If
'        End If
'        mbByteSizeSet = True
'        ByteSize = mlByteSize
'    End If
'End Property
'
'Public Property Let ByteSize(ByVal lByteSize As Long)
'    mbByteSizeSet = True
'    mlByteSize = lByteSize
'End Property
'
'Public Property Get BitSize() As Long
'    Dim cTotalRange As Currency
'
'    If mbBitSizeSet Then
'        BitSize = mlBitSize
'    Else
'        cTotalRange = Range.Ending - Range.Starting
'
'        If cTotalRange > 0 Then
'            mlBitSize = Int(Log(cTotalRange) / Log(2) + 1)
'        Else
'            mlBitSize = Int(Log(Range.Starting) / Log(2) + 1)
'        End If
'        mbBitSizeSet = True
'        BitSize = mlBitSize
'    End If
'End Property
'
'Public Property Let BitSize(lBitSize As Long)
'    mbBitSizeSet = True
'    mlBitSize = lBitSize
'End Property

Private Sub Class_Initialize()
    AccessType = atByte
End Sub
