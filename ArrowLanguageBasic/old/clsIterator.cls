VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIterator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public moIMember As clsIMember
Public mlIndex As Long
Public mlRealValue As Long

Private mlLowerLimit As Long
Private mlUpperLimit As Long

Public Function NextValue() As Boolean
    Dim oRangeSet As clsRangeSet
    Dim oIdentifierSet As clsIdentifierSet
    Dim oSymbolSet As clsSymbolSet
    
    mlIndex = mlIndex + 1
    If mlIndex <= mlUpperLimit Then
        Select Case moIMember.MemberType
            Case RANGE_SET
                Set oRangeSet = moIMember
                mlRealValue = GetRealValue(mlIndex, oRangeSet)
            Case IDENTIFIER_SET
                Set oIdentifierSet = moIMember
            Case SYMBOL_SET
                Set oSymbolSet = moIMember
        End Select
        
        NextValue = True
    End If
End Function

Public Property Set Member(ByVal oIMember As clsIMember)
    Dim oRangeSet As clsRangeSet
    Dim oIdentifierSet As clsIdentifierSet
    Dim oSymbolSet As clsSymbolSet
    
    Set moIMember = oIMember
    Select Case moIMember.MemberType
        Case RANGE_SET
            Set oRangeSet = oIMember
            SetRangeLimits oRangeSet
        Case IDENTIFIER_SET
            Set oIdentifierSet = oIMember
        Case SYMBOL_SET
            Set oSymbolSet = oIMember
    End Select
    
    mlIndex = -1
    NextValue
End Property

Private Sub SetRangeLimits(ByVal oRangeSet As clsRangeSet)
    Dim oRange As clsRange
    Dim lIndex As Long
    
    mlLowerLimit = 0
    For lIndex = 0 To oRangeSet.Count - 1
        Set oRange = oRangeSet.RangeSet(lIndex)
        mlUpperLimit = mlUpperLimit + Abs(oRange.Ending - oRange.Starting) + 1
    Next
    mlUpperLimit = mlUpperLimit - 1
End Sub

Private Function GetRealValue(ByVal lIteratorIndex As Long, ByVal oRangeSet As clsRangeSet) As Long
    Dim lThisRangeStart As Long
    Dim oRange As clsRange
    Dim lIndex As Long
    
    For lIndex = 0 To oRangeSet.Count - 1
        Set oRange = oRangeSet.RangeSet(lIndex)
        If lIteratorIndex >= lThisRangeStart And lIteratorIndex <= (lThisRangeStart + Abs(oRange.Ending - oRange.Starting)) Then
            GetRealValue = (lIteratorIndex - lThisRangeStart) * Sgn(oRange.Ending - oRange.Starting) + oRange.Starting
            Exit Function
        End If
        lThisRangeStart = lThisRangeStart + Abs(oRange.Ending - oRange.Starting) + 1
    Next
End Function
