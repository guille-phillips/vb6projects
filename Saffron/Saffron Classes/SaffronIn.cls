VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SaffronIn"
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
Private msInString As String

Private Function ISaffronObject_Copy(Optional ByVal bNonConsuming As Variant, Optional ByVal bOmit As Variant) As ISaffronObject
    If IsMissing(bNonConsuming) Then
        bNonConsuming = mbNonConsuming
    End If
    If IsMissing(bOmit) Then
        bOmit = mbOmit
    End If
    
    Set ISaffronObject_Copy = New SaffronIn
    ISaffronObject_Copy.Initialise CBool(bNonConsuming), CBool(bOmit), msInString
End Function

Private Property Get ISaffronObject_Omitted() As Boolean
    ISaffronObject_Omitted = mbOmit
End Property


' Initialises parameters used for parsing
Public Function ISaffronObject_Initialise(bNonConsuming As Boolean, bOmit As Boolean, ParamArray pInitParam() As Variant) As ISaffronObject
    mbNonConsuming = bNonConsuming
    mbOmit = bOmit
    
    msInString = pInitParam(0)
    
    Set ISaffronObject_Initialise = Me
End Function

' Will perform the parsing function on the object - if parsing fails will return FALSE.
Public Function ISaffronObject_Parse(oResult As SaffronTree) As Boolean
    Dim sTestChar As String
    Dim sGetChar As String
    Dim mlPosition As Long
    
    oResult.Leaf = True
        
    mlPosition = TextPosition
    oResult.TextStart = mlPosition
    
    If TextPosition > LenTextString Then
        oResult.TextStart = mlPosition
        oResult.Error = True
        Exit Function
    End If
    
    sGetChar = Mid$(TextString, TextPosition, 1):
    
    If InStr(1, msInString, sGetChar, vbBinaryCompare) <> 0 Then
        oResult.TextStart = mlPosition
        oResult.TextEnd = mlPosition
        If Not mbNonConsuming Then
            TextPosition = TextPosition + 1
        End If
        ISaffronObject_Parse = True
    Else
        oResult.Error = True
    End If
End Function
