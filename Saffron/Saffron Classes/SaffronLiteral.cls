VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SaffronLiteral"
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
Private msLiteralString As String
Private msOriginalLiteralString As String
Private mlLiteralStringLength As Long
Private mbCaseSensitive As Boolean

Private Function ISaffronObject_Copy(Optional ByVal bNonConsuming As Variant, Optional ByVal bOmit As Variant) As ISaffronObject

    If IsMissing(bNonConsuming) Then
        bNonConsuming = mbNonConsuming
    End If
    If IsMissing(bOmit) Then
        bOmit = mbOmit
    End If
    
    Set ISaffronObject_Copy = New SaffronLiteral
    ISaffronObject_Copy.Initialise CBool(bNonConsuming), CBool(bOmit), msLiteralString, mbCaseSensitive
End Function

Private Property Get ISaffronObject_Omitted() As Boolean
    ISaffronObject_Omitted = mbOmit
End Property


' Initialises parameters used for parsing
Public Function ISaffronObject_Initialise(bNonConsuming As Boolean, bOmit As Boolean, ParamArray pInitParam() As Variant) As ISaffronObject
    mbNonConsuming = bNonConsuming
    mbOmit = bOmit
    
    msLiteralString = pInitParam(0)
    mlLiteralStringLength = Len(msLiteralString)
    msOriginalLiteralString = msLiteralString
    If UBound(pInitParam) = 1 Then
        mbCaseSensitive = pInitParam(UBound(pInitParam))
    End If
    
    If Not mbCaseSensitive Then
        msLiteralString = UCase$(msLiteralString)
    End If
    Set ISaffronObject_Initialise = Me
End Function

' Will perform the parsing function on the object - if parsing fails will return FALSE.
Public Function ISaffronObject_Parse(oResult As SaffronTree) As Boolean
    Dim lPosition As Long
    Dim sGetLiteral As String
    
    oResult.Leaf = True
    
    lPosition = TextPosition
    oResult.TextStart = lPosition
    
    If (TextPosition + mlLiteralStringLength - 1) > LenTextString Then
        oResult.Error = True
        Exit Function
    End If
    
    sGetLiteral = Mid$(TextString, TextPosition, mlLiteralStringLength)

    If Not mbCaseSensitive Then
        sGetLiteral = UCase$(sGetLiteral)
    End If
    If msLiteralString = sGetLiteral Then
        oResult.TextEnd = lPosition + mlLiteralStringLength - 1
        If Not mbNonConsuming Then
            TextPosition = TextPosition + mlLiteralStringLength
        End If
        ISaffronObject_Parse = True
    Else
        oResult.Error = True
    End If
End Function
