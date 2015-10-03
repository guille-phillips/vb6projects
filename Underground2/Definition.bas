Attribute VB_Name = "Definition"
Option Explicit

Public oParsePosition As IParseObject
Public oParseRelationship As IParseObject
Public oParseLine As IParseObject

Public Const pi2  As Double = 6.28318530717959

Public Const LineWidth As Double = 12
Public Const OuterCircleRadius As Double = 10
Public Const InnerCircleRadius As Double = 8
Public Const TextSeparation As Double = 15
Public Const LineGap As Double = 17
Public Const FontSize As Double = 10

Public Sub DoDefinition()
    Dim sDef As String
    
    sDef = sDef & "pos := REPEAT IN '0' TO '9', '.', '-';"
    sDef = sDef & "name := REPEAT IN 32 TO 255, NOT '|' MIN 0;"
    sDef = sDef & "position := AND ['P:'], reference, ['|'], name, ['|'], pos, ['|'], pos, ['|'], pos, ['|'], pos;"
    sDef = sDef & "reference := REPEAT IN CASE 'A' TO 'F', '0' TO '9';"
    sDef = sDef & "relationship := AND ['R:'], reference, ['|'], reference, ['|'], pos, ['|'], (LIST pos, ['|']);"
    sDef = sDef & "offset := AND ['O:'], pos, ['|'], pos;"
    sDef = sDef & "zoom := AND ['Z:'], pos;"
    sDef = sDef & "colours := AND ['C:'], (LIST pos, ['|']);"
    sDef = sDef & "line := OR position, relationship, offset, zoom, colours;"
    
    If Not SetNewDefinition(sDef) Then
        MsgBox "Bad Def"
        End
    End If
    
    Set oParsePosition = ParserObjects("position")
    Set oParseRelationship = ParserObjects("relationship")
    Set oParseLine = ParserObjects("line")
End Sub


Public Sub AddToArray(ByRef vArray As Variant, lColour As Long)
    Dim lUbound As Long
    
    lUbound = UBound(vArray)
    ReDim Preserve vArray(lUbound + 1)
    vArray(lUbound + 1) = lColour
End Sub

Public Sub RemoveFromArray(vArray As Variant, lColour As Long)
    Dim lUbound As Long
    Dim lIndex As Long
    Dim lIndex2 As Long
    
    lIndex = 0
    lUbound = UBound(vArray)
    For lIndex = 0 To lUbound
        If vArray(lIndex) = lColour Then
            For lIndex2 = lIndex To lUbound - 1
                vArray(lIndex2) = vArray(lIndex2 + 1)
            Next
            If lUbound > 0 Then
                ReDim Preserve vArray(lUbound - 1)
            Else
                vArray = Array()
            End If
            Exit For
        End If
    Next
End Sub

Public Function InArray(ByRef vArray As Variant, lColour As Long) As Boolean
    Dim lUbound As Long
    Dim lIndex As Long
    
    lIndex = 0
    lUbound = UBound(vArray)
    For lIndex = 0 To lUbound
        If vArray(lIndex) = lColour Then
            InArray = True
            Exit For
        End If
    Next
End Function
