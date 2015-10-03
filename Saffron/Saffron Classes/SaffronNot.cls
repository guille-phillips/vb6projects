VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SaffronNot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements ISaffronObject

Private mbNonConsuming As Boolean
Private mbOmit As Boolean
Private moNotObject As ISaffronObject

Private Function ISaffronObject_Copy(Optional ByVal bNonConsuming As Variant, Optional ByVal bOmit As Variant) As ISaffronObject
    If IsMissing(bNonConsuming) Then
        bNonConsuming = mbNonConsuming
    End If
    If IsMissing(bOmit) Then
        bOmit = mbOmit
    End If
    
    Set ISaffronObject_Copy = New SaffronNot
    ISaffronObject_Copy.Initialise CBool(bNonConsuming), CBool(bOmit), moNotObject.Copy
End Function

Private Property Get ISaffronObject_Omitted() As Boolean
    ISaffronObject_Omitted = mbOmit
End Property

' Initialises parameters used for parsing
Public Function ISaffronObject_Initialise(bNonConsuming As Boolean, bOmit As Boolean, ParamArray pInitParam() As Variant) As ISaffronObject
    mbNonConsuming = bNonConsuming
    mbOmit = bOmit
    
    Set moNotObject = pInitParam(LBound(pInitParam))
    
    Set ISaffronObject_Initialise = Me
End Function

' Will perform the parsing function on the object - if parsing fails will return FALSE.
Public Function ISaffronObject_Parse(oResult As SaffronTree) As Boolean
    Dim oObjectResult As New SaffronTree
    
    oResult.TextStart = TextPosition
    oResult.TextEnd = TextPosition - 1
    
    ISaffronObject_Parse = Not moNotObject.Parse(oObjectResult)
    If Not ISaffronObject_Parse Then
        oResult.Error = True
    End If
End Function
