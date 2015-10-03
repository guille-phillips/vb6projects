VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SaffronEach"
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
    
    Set ISaffronObject_Copy = New SaffronEach
    ISaffronObject_Copy.Initialise CBool(bNonConsuming), CBool(bOmit), moObjects
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
        mlObjectCount = pInitParam(0).Count
        ReDim moObjects(1 To mlObjectCount) As ISaffronObject

        For lObjectIndex = 1 To mlObjectCount
            Set moObjects(lObjectIndex) = pInitParam(0)(lObjectIndex)
        Next
    Else
        mlObjectCount = (UBound(pInitParam) - LBound(pInitParam) + 1)
        ReDim moObjects(1 To mlObjectCount) As ISaffronObject
        
        lOffset = LBound(pInitParam)
        For lObjectIndex = lOffset To UBound(pInitParam)
            Set moObjects(lObjectIndex - lOffset + 1) = pInitParam(lObjectIndex)
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
    Dim myOrObjectIndex As Long
    
    Dim lControl As Boolean
    
    lPosition = TextPosition
    lObjectIndex = 1
    For myOrObjectIndex = 1 To mlObjectCount
        Set oResult = New SaffronTree
        If moObjects(myOrObjectIndex).Parse(oResult) Then
            ISaffronObject_Parse = True
                            
            'lControl = myResultControls(lObjectIndex)
'            If lControl = pcLeave Then
                'lControl = oResult.Control
'            End If
            
'            If lControl <> pcOmit Then
                oResult.TextStart = lPosition
                oResult.TextEnd = TextPosition - 1

                oResult.ParentIndex = lObjectIndex
                oResult.SubTree.Add oResult
'            End If
            lObjectIndex = lObjectIndex + 1
        End If
        TextPosition = lPosition
    Next
    
    oResult.Index = lObjectIndex
    
    'oResult.Control = myResultControl
    If Not ISaffronObject_Parse Then
        oResult.Error = True
        oResult.TextStart = TextPosition
    End If
End Function

