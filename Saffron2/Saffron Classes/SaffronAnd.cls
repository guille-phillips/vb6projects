VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SaffronAnd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements ISaffronObject

' Generic variables
Private mbOmit As Boolean
Private mbNonConsuming As Boolean

' Object specific variables
Private moObjects() As ISaffronObject
Private mlObjectCount As Long

Private Property Get ISaffronObject_Omitted() As Boolean
    ISaffronObject_Omitted = mbOmit
End Property

' Initialises parameters used for parsing
Public Function ISaffronObject_Initialise(bNonConsuming As Boolean, bOmit As Boolean, ParamArray pInitParam() As Variant) As ISaffronObject
    Dim lOffset As Long
    Dim lAndObjectIndex As Long
    
    mbNonConsuming = bNonConsuming
    mbOmit = bOmit

    If TypeOf pInitParam(0) Is Collection Then
        mlObjectCount = pInitParam(0).Count
        ReDim moObjects(1 To mlObjectCount) As ISaffronObject

        For lAndObjectIndex = 1 To mlObjectCount
            Set moObjects(lAndObjectIndex) = pInitParam(0)(lAndObjectIndex)
        Next
    Else
        mlObjectCount = (UBound(pInitParam) - LBound(pInitParam) + 1)
        ReDim moObjects(1 To mlObjectCount) As ISaffronObject
        
        lOffset = LBound(pInitParam)
        For lAndObjectIndex = lOffset To UBound(pInitParam)
            Set moObjects(lAndObjectIndex - lOffset + 1) = pInitParam(lAndObjectIndex)
        Next
    End If

    Set ISaffronObject_Initialise = Me
End Function


' Will perform the parsing function on the object - if parsing fails will return FALSE.
Public Function ISaffronObject_Parse(oResult As SaffronTree) As Boolean
    Dim oAndObject As ISaffronObject
    Dim oObjectResult As SaffronTree
    Dim lPosition As Long
    Dim lObjectIndex As Long
    
    Dim lControl As Boolean
    
    lPosition = TextPosition
    oResult.TextStart = lPosition
    
    ISaffronObject_Parse = True
    For lObjectIndex = 1 To mlObjectCount
        Set oObjectResult = New SaffronTree

        If Not moObjects(lObjectIndex).Parse(oObjectResult) Then
            ISaffronObject_Parse = False
            oResult.Error = True
            oResult.Index = lObjectIndex
            TextPosition = lPosition
            Exit Function
        End If

        oResult.TextEnd = oObjectResult.TextEnd
        If Not moObjects(lObjectIndex).Omitted Then
            oResult.SubTree.Add oObjectResult
        End If
    Next
    If mbNonConsuming Then
        TextPosition = lPosition
    End If
End Function
