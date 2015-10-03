Attribute VB_Name = "Module1"
Option Explicit
Private oParser As ISaffronObject

Sub main()
    InitParser
    'TestObject
   ' InitialiseNameParser
End Sub

Private Sub InitParser()
    Dim oTree As SaffronTree
    
    If Not SetDefinition("test in |123 | |") Then
        MsgBox "Bad Def"
    End If
    
    Set oParser = ParserObjects("test")
    SaffronStream.Text = "123Aexp|delexpdelexpq"
    Set oTree = New SaffronTree
    
    Set SaffronStream.External = New Dictionary.Lookup
    If oParser.Parse(oTree) Then
        Stop
    End If
End Sub

Private Sub TestObject()
    Dim x As New SaffronList
    Dim o As ISaffronObject
    Dim t As New SaffronTree
    
    Dim l1 As New SaffronLiteral
    Dim l2 As New SaffronLiteral
    
    Dim lo1 As ISaffronObject
    Dim lo2 As ISaffronObject
    
    Set lo1 = l1.ISaffronObject_Initialise(False, False, ",")
    Set lo2 = l2.ISaffronObject_Initialise(False, False, "a")
    
    Set o = x.ISaffronObject_Initialise(False, False, l2, l1, 0, 3)
    
    Stream.Text = "a,"
    Debug.Print o.Parse(t)
    'Debug.Print t.Text
End Sub


'Public Sub InitialiseNameParser()
'    Dim Definition As String
'    Dim oStripQuotes As IParseObject
'
'    Definition = "strip_quotes := in 32, 123 to 124, 'x';"
'
'    If Not SetNewDefinition(Definition) Then
'        Debug.Print ErrorString
'        End
'    End If
'
'    Set oStripQuotes = ParserObjects("strip_quotes")
'End Sub
