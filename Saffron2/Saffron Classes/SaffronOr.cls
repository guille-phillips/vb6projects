VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SaffronOr"
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
    Dim lOrObjectIndex As Long
    
    mbNonConsuming = bNonConsuming
    mbOmit = bOmit

    If TypeOf pInitParam(0) Is Collection Then
        mlObjectCount = pInitParam(0).Count
        ReDim moObjects(1 To mlObjectCount) As ISaffronObject

        For lOrObjectIndex = 1 To mlObjectCount
            Set moObjects(lOrObjectIndex) = pInitParam(0)(lOrObjectIndex)
        Next
    Else
        mlObjectCount = (UBound(pInitParam) - LBound(pInitParam) + 1)
        ReDim moObjects(1 To mlObjectCount) As ISaffronObject
        
        lOffset = LBound(pInitParam)
        For lOrObjectIndex = lOffset To UBound(pInitParam)
            Set moObjects(lOrObjectIndex - lOffset + 1) = pInitParam(lOrObjectIndex)
        Next
    End If

    Set ISaffronObject_Initialise = Me
End Function

' Will perform the parsing function on the object - if parsing fails will return FALSE.
Public Function ISaffronObject_Parse(oResult As SaffronTree) As Boolean
    Dim lPosition As Long
    Dim oOrObject As ISaffronObject
    Dim lObjectIndex As Long
    Dim oObjectResult As SaffronTree
    Dim lOrObjectIndex As Long
    Dim lControl As Boolean

    lPosition = TextPosition
    For lOrObjectIndex = 1 To mlObjectCount
        Set oObjectResult = New SaffronTree

        If moObjects(lOrObjectIndex).Parse(oObjectResult) Then
            ISaffronObject_Parse = True

            oResult.TextStart = lPosition
            oResult.TextEnd = oObjectResult.TextEnd

            If Not moObjects(lOrObjectIndex).Omitted Then
                oResult.SubTree.Add oObjectResult
            End If

            oResult.Index = lOrObjectIndex
            Exit For
        End If
        TextPosition = lPosition
    Next
    
    If Not ISaffronObject_Parse Then
        oResult.Error = True
        oResult.TextStart = TextPosition
    End If
    
    If mbNonConsuming Then
        TextPosition = lPosition
    End If
    
End Function
