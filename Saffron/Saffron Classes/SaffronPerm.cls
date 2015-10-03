VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SaffronPerm"
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
Private mbAny As Boolean

Private Function ISaffronObject_Copy(Optional ByVal bNonConsuming As Variant, Optional ByVal bOmit As Variant) As ISaffronObject
    Dim oObjects As New Collection
    Dim lIndex As Long

    If IsMissing(bNonConsuming) Then
        bNonConsuming = mbNonConsuming
    End If
    If IsMissing(bOmit) Then
        bOmit = mbOmit
    End If
    
    For lIndex = 1 To UBound(moObjects)
        oObjects.Add moObjects(lIndex).Copy
    Next
    
    Set ISaffronObject_Copy = New SaffronPerm
    ISaffronObject_Copy.Initialise CBool(bNonConsuming), CBool(bOmit), mbAny, oObjects
End Function

Private Property Get ISaffronObject_Omitted() As Boolean
    ISaffronObject_Omitted = mbOmit
End Property

' Initialises parameters used for parsing
Public Function ISaffronObject_Initialise(bNonConsuming As Boolean, bOmit As Boolean, ParamArray pInitParam() As Variant) As ISaffronObject
    Dim lOffset As Long
    Dim lObjectIndex As Long
    
    mbNonConsuming = bNonConsuming
    mbOmit = bOmit

    If TypeOf pInitParam(0) Is Collection Then
        mbAny = pInitParam(0)(0)
        mlObjectCount = pInitParam(0).Count - 1
        ReDim moObjects(1 To mlObjectCount) As ISaffronObject

        For lObjectIndex = 1 To mlObjectCount
            Set moObjects(lObjectIndex) = pInitParam(0)(lObjectIndex + 1)
        Next
    Else
        mbAny = pInitParam(0)
        
        mlObjectCount = (UBound(pInitParam) - LBound(pInitParam))
        ReDim moObjects(1 To mlObjectCount) As ISaffronObject
        
        lOffset = LBound(pInitParam) + 1
        For lObjectIndex = lOffset To UBound(pInitParam)
            Set moObjects(lObjectIndex - lOffset + 1) = pInitParam(lObjectIndex)
        Next
    End If

    Set ISaffronObject_Initialise = Me
End Function

' Will perform the parsing function on the object - if parsing fails will return FALSE.
Public Function ISaffronObject_Parse(oResult As SaffronTree) As Boolean
    Dim bParsed() As Boolean
    Dim oObjectResult As SaffronTree
    Dim lPosition As Long
    Dim lObjectIndex As Long
    Dim lParsedCount As Long
    Dim lPositionIndex As Long
    
    ReDim bParsed(1 To mlObjectCount)
    
    lPosition = TextPosition
    ISaffronObject_Parse = True
    
    oResult.TextStart = lPosition

    For lObjectIndex = 1 To mlObjectCount
        If Not bParsed(lObjectIndex) Then
            Set oObjectResult = New SaffronTree
            
            If moObjects(lObjectIndex).Parse(oObjectResult) Then
                bParsed(lObjectIndex) = True
                lParsedCount = lParsedCount + 1
                
                If Not moObjects(lObjectIndex).Omitted Then
                    oResult.TextEnd = TextPosition - 1
                    oResult.SubTree.Add oObjectResult
                End If
                lObjectIndex = 0
            End If
        End If
    Next
    
    oResult.Index = lParsedCount
    
    If Not mbAny Then
        If lParsedCount < mlObjectCount Then
            ISaffronObject_Parse = False
            TextPosition = lPosition
            oResult.Error = True
        End If
    End If
End Function
