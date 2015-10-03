VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRangeSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Option Explicit
'
'Implements clsIMember
'
'Private moRangeSet() As clsRange
'Private mlCount As Long
'
'Private msIdentifier As String
'Private mlSize As Long
'
'Private Property Let clsIMember_Identifier(ByVal sIdentifier As String)
'    msIdentifier = sIdentifier
'End Property
'
'Private Property Get clsIMember_Identifier() As String
'    clsIMember_Identifier = msIdentifier
'End Property
'
'Private Property Get clsIMember_MemberType() As Long
'    clsIMember_MemberType = RANGE_SET
'End Property
'
'Public Sub AddRange(ByVal oRange As clsRange)
'    ReDim Preserve moRangeSet(mlCount)
'    Set moRangeSet(mlCount) = oRange
'    mlCount = mlCount + 1
'    mlSize = mlSize + Abs(oRange.Ending - oRange.Starting) + 1
'End Sub
'
'Public Property Get RangeSet(ByVal lIndex As Long) As clsRange
'    Set RangeSet = moRangeSet(lIndex)
'End Property
'
'Public Property Let RangeSet(ByVal lIndex As Long, oRange As clsRange)
'    moRangeSet(lIndex) = oRange
'End Property
'
'Private Property Get clsIMember_Size() As Long
'    clsIMember_Size = mlSize
'End Property
'
'Public Property Get Count() As Long
'    Count = mlSize
'End Property
