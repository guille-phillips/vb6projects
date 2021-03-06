VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SaffronExternal"
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
Private moFunctionName As String
Private mvParameters As Variant

Private Property Get ISaffronObject_Omitted() As Boolean
    ISaffronObject_Omitted = mbOmit
End Property


' Initialises parameters used for parsing
Public Function ISaffronObject_Initialise(bNonConsuming As Boolean, bOmit As Boolean, ParamArray pInitParam() As Variant) As ISaffronObject
    mbNonConsuming = bNonConsuming
    mbOmit = bOmit

    moFunctionName = pInitParam(0)
    mvParameters = pInitParam(1)
    
    Set ISaffronObject_Initialise = Me
End Function

' Will perform the parsing function on the object - if parsing fails will return FALSE.
Public Function ISaffronObject_Parse(oResult As SaffronTree) As Boolean
    Dim lPosition As Long
    Dim lIndex As Long
    
    lPosition = TextPosition
    oResult.TextStart = lPosition
    oResult.Leaf = True
    
    If CallByName(ExternalClass, moFunctionName, VbMethod, TextString, TextPosition, lIndex, mvParameters) Then
        oResult.Index = lIndex
        oResult.TextEnd = TextPosition - 1
        If mbNonConsuming Then
            TextPosition = lPosition
        End If
        ISaffronObject_Parse = True
    Else
        TextPosition = lPosition
        oResult.Error = True
    End If
End Function

