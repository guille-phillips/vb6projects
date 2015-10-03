Attribute VB_Name = "Formatting"
Option Explicit

Public goHTMLParser As IParseObject

Private mbNotIndented As Boolean

Public Sub Main()
    InitialiseParser
    ParseFiles
End Sub

Private Sub InitialiseParser()
    Dim sFile As String
    Dim sDefinition As String
    
    sFile = App.Path & "/WebFormat.pdl"
    sDefinition = String$(FileLen(sFile), Chr(0))
    Open sFile For Binary As #1
    Get #1, , sDefinition
    Close #1
    
    If Not SetNewDefinition(sDefinition) Then
        MsgBox "Bad Def"
        End
    End If
    
    Set goHTMLParser = ParserObjects("html")
    
'    Dim oTree As ParseTree
'    Set goHTMLParser = ParserObjects("css")
'
'    sFile = App.Path & "/test.php"
'    sDefinition = String$(FileLen(sFile), Chr(0))
'    Open sFile For Binary As #1
'    Get #1, , sDefinition
'    Close #1
'
'    Stream.Text = sDefinition
'    Set oTree = New ParseTree
'    If goHTMLParser.Parse(oTree) Then
'        Stop
'    Else
'        Stop
'    End If
End Sub

Private Sub ParseFiles()
    Dim sFile As String
    Dim sContents As String
    Dim sFormattedFile As String
    Dim vExtensions As Variant
    Dim vExtension As Variant
    vExtensions = Array("php", "htm", "html", "shtml")
    
    For Each vExtension In vExtensions
        sFile = Dir(App.Path & "\*." & vExtension)
        On Error Resume Next
        MkDir App.Path & "\cleaned"
        On Error GoTo 0
        While Not sFile = ""
            sContents = String$(FileLen(sFile), Chr(0))
            Open sFile For Binary As #1
            Get #1, , sContents
            Close #1
        
            sFormattedFile = ParseFile(sContents)
            On Error Resume Next
            Kill App.Path & "\cleaned\" & sFile
            On Error GoTo 0
            Open App.Path & "\cleaned\" & sFile For Binary As #1
            Put #1, , sFormattedFile
            Close #1
            sFile = Dir
        Wend
    Next
    MsgBox "Files have been formatted.", vbOKOnly
End Sub

Private Function ParseFile(sContents As String) As String
    Dim oTree As New ParseTree
    
    Stream.Text = sContents
    
    If goHTMLParser.Parse(oTree) Then
        ParseFile = FormatHTML(oTree, 0)
    Else
        ParseFile = sContents
    End If
    
    'Debug.Print ParseFile
End Function

Private Function LineCount(ByVal sText As String) As Long
    LineCount = UBound(Split(sText, vbCrLf)) + 1
End Function

Private Function FormatHTML(oTree As ParseTree, ByVal lIndentationLevel As Long) As String
    Dim oLine As ParseTree
    Dim sIndentation As String
    Dim bIndent As Boolean
    Dim sLastTag As String
    Dim oPHPTree As ParseTree
    Dim sTrimmed As String
    Dim sFormatted As String
    Dim bMultilineText As Boolean
    
    For Each oLine In oTree.SubTree
        sIndentation = Indentation(lIndentationLevel)
        Select Case oLine.Index
            Case 1 ' php
                Select Case oLine!php(1).Index
                    Case 1 ' php
                        sFormatted = FormatPHP(oLine!php, lIndentationLevel)
                        If LineCount(sFormatted) = 3 Then
                            mbNotIndented = True
                            sFormatted = TrimText(FormatPHP(oLine!php, lIndentationLevel))
                            FormatHTML = FormatHTML & IIf(Left$(sLastTag, 1) = "/", vbCrLf & sIndentation, "") & sFormatted
                            bMultilineText = False
                            mbNotIndented = False
                        Else
                            bMultilineText = True
                            FormatHTML = FormatHTML & NewLine & sFormatted
                        End If
                    Case 2 ' text
                        If LineCount(oLine.Text) > 1 Then
                            sTrimmed = "<?php " & NewLine & oLine.Text & NewLine & sIndentation & "?>"
                        Else
                            sTrimmed = "<?php " & TrimText(oLine.Text) & "?>"
                        End If
                        If Len(sTrimmed) > 80 Or LineCount(sTrimmed) > 1 Or Left$(sLastTag, 1) = "/" Then
                            FormatHTML = FormatHTML & NewLine & sIndentation & sTrimmed
                            sLastTag = ""
                            bMultilineText = True
                        Else
                            FormatHTML = FormatHTML & sTrimmed
                            bMultilineText = False
                        End If
                End Select

            Case 2 ' javascript
                Select Case oLine!javascript(2).Index
                    Case 1 ' javascript
                        FormatHTML = FormatHTML & NewLine & FormatJavascript(oLine!javascript, lIndentationLevel)
                    Case 2 ' text
                        sTrimmed = FormatOpenTag(oLine(1)!open_tag, lIndentationLevel) & NewLine & oLine(1)(2).Text & NewLine & sIndentation & "</script>"
                        If Len(sTrimmed) > 80 Or LineCount(sTrimmed) > 3 Then
                            FormatHTML = FormatHTML & NewLine & sIndentation & sTrimmed
                            sLastTag = ""
                            bMultilineText = True
                        Else
                            FormatHTML = FormatHTML & sTrimmed
                            bMultilineText = False
                        End If
                End Select
            Case 3 ' css
                Select Case oLine!css(2).Index
                    Case 1 ' css
                        FormatHTML = FormatHTML & NewLine & FormatCSS(oLine!css, lIndentationLevel)
                    Case 2 'text
                        sTrimmed = FormatOpenTag(oLine(1)!open_tag, lIndentationLevel) & NewLine & oLine(1)(2).Text & NewLine & sIndentation & "</style>"
                        If Len(sTrimmed) > 80 Or LineCount(sTrimmed) > 3 Then
                            FormatHTML = FormatHTML & NewLine & sIndentation & sTrimmed
                            sLastTag = ""
                            bMultilineText = True
                        Else
                            FormatHTML = FormatHTML & sTrimmed
                            bMultilineText = False
                        End If
                End Select
            Case 4 ' open tag
                FormatHTML = FormatHTML & NewLine & sIndentation & FormatOpenTag(oLine!open_tag, lIndentationLevel)
                sLastTag = oLine!open_tag(1).Text
            Case 5 ' close tag
                lIndentationLevel = lIndentationLevel - 1
                sIndentation = Indentation(lIndentationLevel)
                If oLine!close_tag.Text <> sLastTag Or bMultilineText Then
                    FormatHTML = FormatHTML & NewLine & sIndentation
                End If
                FormatHTML = FormatHTML & "</" & oLine!close_tag.Text & ">"
                sLastTag = "/" & oLine!close_tag.Text
                bMultilineText = False
            Case 6 ' text
                sTrimmed = TrimText(oLine.Text)
                If Len(sTrimmed) > 80 Or LineCount(sTrimmed) > 1 Then
                    FormatHTML = FormatHTML & NewLine & sIndentation & sTrimmed
                    sLastTag = ""
                    bMultilineText = True
                Else
                    FormatHTML = FormatHTML & sTrimmed
                    'bMultilineText = False
                End If
        End Select
    Next
    FormatHTML = TrimText(FormatHTML)
End Function

Private Function TrimText(ByVal sText As String) As String
    Dim lPosition As Long
    Dim sChar As String
    
    Do
        lPosition = lPosition + 1
        sChar = Mid$(sText, lPosition, 1)
    Loop While Asc(sChar & "X") < 33
    sText = Mid$(sText, lPosition)
    
    If sText = "" Then
        Exit Function
    End If
    lPosition = Len(sText) + 1
    Do
        lPosition = lPosition - 1
        sChar = Mid$(sText, lPosition, 1)
    Loop While Asc(sChar) < 33
    sText = Left$(sText, lPosition)
    TrimText = sText
End Function

Private Function FormatOpenTag(oTree As ParseTree, lIndentationLevel As Long) As String
    Dim oAttribute As ParseTree
    Dim bIndentation As Boolean
    
    Select Case oTree(1).Index
        Case 1 ' non indenting
        Case 2 ' indenting
            lIndentationLevel = lIndentationLevel + 1
    End Select
    FormatOpenTag = "<" & oTree(1).Text
    
    For Each oAttribute In oTree(2).SubTree
        Select Case oAttribute.Index
            Case 1 ' standard
                FormatOpenTag = FormatOpenTag & " " & oAttribute(1)!Tag.Text
                If oAttribute(1)(2).Index = 1 Then
                    FormatOpenTag = FormatOpenTag & "=" & oAttribute(1)(2)(1)!attribute_value.Text
                End If
            Case 2 ' php
                bIndentation = mbNotIndented
                mbNotIndented = True
                FormatOpenTag = FormatOpenTag & FormatPHP(oAttribute(1), 0)
                mbNotIndented = bIndentation
        End Select
    Next
    FormatOpenTag = FormatOpenTag & ">"
End Function

Private Function FormatPHP(oTree As ParseTree, ByVal lIndentationLevel As Long) As String
    Dim oStatement As ParseTree
    Dim sIndentation As String
    Dim sIndentation1 As String
    Dim lLineCount As Long
    
    sIndentation = Indentation(lIndentationLevel)
    sIndentation1 = Indentation(lIndentationLevel + 1)
    
    FormatPHP = sIndentation & "<?php" & NewLine
    FormatPHP = FormatPHP & FormatStatements(oTree(1)(1)!statements, lIndentationLevel)
    FormatPHP = FormatPHP & sIndentation & "?>"
End Function

Private Function FormatJavascript(oTree As ParseTree, lIndentationLevel As Long) As String
    Dim oStatement As ParseTree
    Dim sIndentation As String
    Dim sIndentation1 As String
    
    sIndentation = Indentation(lIndentationLevel)
    sIndentation1 = Indentation(lIndentationLevel + 1)
    
    FormatJavascript = sIndentation & FormatOpenTag(oTree!open_tag, lIndentationLevel) & NewLine
    FormatJavascript = FormatJavascript & FormatStatements(oTree(2)(1)!statements, lIndentationLevel)
    FormatJavascript = FormatJavascript & sIndentation & "</script>"
End Function

Private Function FormatStatements(oTree As ParseTree, ByVal lIndentationLevel As Long) As String
    Dim oStatement As ParseTree
    Dim sIndentation As String
    Dim sIndentation1 As String
    Dim bOneLine As Boolean
    
    bOneLine = lIndentationLevel = -1
        
    sIndentation = Indentation(lIndentationLevel)
    sIndentation1 = Indentation(lIndentationLevel + 1)

    For Each oStatement In oTree.SubTree
        FormatStatements = FormatStatements & IIf(bOneLine, " ", Mid$(oStatement!line_ws.Text, 4))
        FormatStatements = FormatStatements & FormatStatement(oStatement!statement, lIndentationLevel)
    Next
End Function

Private Function FormatStatement(oTree As ParseTree, ByVal lIndentationLevel As Long) As String
    Dim oStatement As ParseTree
    Dim sIndentation As String
    Dim sIndentation1 As String
    Dim sNewLine As String
    
    sIndentation = Indentation(lIndentationLevel)
    sIndentation1 = Indentation(lIndentationLevel + 1)
    
    sNewLine = IIf(lIndentationLevel = -1, "", NewLine)
    
    Set oStatement = oTree
    Select Case oStatement.Index
        Case 1 ' comment
            FormatStatement = FormatStatement & sIndentation1 & oStatement.Text & sNewLine
        Case 2 ' comment long
            FormatStatement = FormatStatement & sIndentation1 & oStatement.Text & sNewLine
        Case 3 ' switch statement
            FormatStatement = FormatStatement & FormatSwitchStatement(oStatement!switch_statement, lIndentationLevel + 1)
        Case 4 ' while statement
            FormatStatement = FormatStatement & FormatWhileStatement(oStatement!while_statement, lIndentationLevel + 1)
        Case 5 ' do while statement
            FormatStatement = FormatStatement & FormatDoWhileStatement(oStatement!do_while_statement, lIndentationLevel + 1)
        Case 6 ' foreach statement
            FormatStatement = FormatStatement & FormatForeachStatement(oStatement!foreach_statement, lIndentationLevel + 1)
        Case 7 ' for statement
            FormatStatement = FormatStatement & FormatForStatement(oStatement!for_statement, lIndentationLevel + 1)
        Case 8 ' with statement
            FormatStatement = FormatStatement & FormatWithStatement(oStatement!with_statement, lIndentationLevel + 1)
        Case 9 ' if statement
            FormatStatement = FormatStatement & FormatIfStatement(oStatement!if_statement, lIndentationLevel + 1)
        Case 10 ' try catch statement
            FormatStatement = FormatStatement & FormatTryCatchStatement(oStatement!try_catch_statement, lIndentationLevel + 1)
        Case 11 ' return statement
            FormatStatement = FormatStatement & sIndentation1 & "return " & FormatExpression(oStatement(1)(1)) & ";" & sNewLine
        Case 12 ' implicit keyword
            FormatStatement = FormatStatement & sIndentation1 & oStatement!implicit_keyword.Text & ";" & sNewLine
        Case 13 ' simple keyword
            FormatStatement = FormatStatement & sIndentation1 & oStatement!simple_keyword(1).Text & " " & FormatExpression(oStatement!simple_keyword(2)) & ";" & sNewLine
        Case 14 ' function declaration
            FormatStatement = FormatStatement & FormatFunctionDeclaration(oStatement!function_declaration, lIndentationLevel + 1)
        Case 15 ' class declaration
            FormatStatement = FormatStatement & FormatClassDeclaration(oStatement!class_declaration, lIndentationLevel + 1)
        Case 16 ' variable declaration
            FormatStatement = FormatStatement & FormatVariableDeclaration(oStatement(0), lIndentationLevel + 1)
        Case 17 ' expression
            FormatStatement = FormatStatement & sIndentation1 & FormatExpression(oStatement(1)(1)) & ";" & sNewLine
        Case 18 ' semicolon
        Case 19 ' ws
    End Select
End Function

Private Function FormatVariableDeclaration(oTree As ParseTree, ByVal lIndentationLevel As Long) As String
    Dim sIndentation As String
    Dim sIndentation1 As String
    
    sIndentation = Indentation(lIndentationLevel)
    sIndentation1 = Indentation(lIndentationLevel + 1)
    
    FormatVariableDeclaration = sIndentation
    FormatVariableDeclaration = FormatVariableDeclaration & LCase$(oTree(1).Text) & " "
    FormatVariableDeclaration = FormatVariableDeclaration & oTree(2).Text
End Function

Private Function FormatWhileStatement(oTree As ParseTree, ByVal lIndentationLevel As Long) As String
    Dim sIndentation As String
    Dim sIndentation1 As String
    Dim oSwitchBlock As ParseTree
    Dim oStatement As ParseTree
    
    sIndentation = Indentation(lIndentationLevel)
    sIndentation1 = Indentation(lIndentationLevel + 1)
    FormatWhileStatement = sIndentation & "while ("
    FormatWhileStatement = FormatWhileStatement & FormatExpression(oTree!expression) & ") "
    FormatWhileStatement = FormatWhileStatement & FormatBlock(oTree!block, lIndentationLevel) & NewLine
End Function


Private Function FormatDoWhileStatement(oTree As ParseTree, ByVal lIndentationLevel As Long) As String
    Dim sIndentation As String
    Dim sIndentation1 As String
    Dim oSwitchBlock As ParseTree
    Dim oStatement As ParseTree
    
    sIndentation = Indentation(lIndentationLevel)
    sIndentation1 = Indentation(lIndentationLevel + 1)
    FormatDoWhileStatement = sIndentation & "do "
    FormatDoWhileStatement = FormatDoWhileStatement & FormatBlock(oTree!block, lIndentationLevel)
    FormatDoWhileStatement = FormatDoWhileStatement & " while (" & FormatExpression(oTree(2)(1)) & ");" & NewLine
End Function


Private Function FormatSwitchStatement(oTree As ParseTree, ByVal lIndentationLevel As Long) As String
    Dim sIndentation As String
    Dim sIndentation1 As String
    Dim oSwitchBlock As ParseTree
    Dim oStatement As ParseTree
    
    sIndentation = Indentation(lIndentationLevel)
    sIndentation1 = Indentation(lIndentationLevel + 1)
    FormatSwitchStatement = sIndentation & "switch ("
    FormatSwitchStatement = FormatSwitchStatement & FormatExpression(oTree!expression) & ") {" & NewLine
    
    Set oSwitchBlock = oTree!switch_block
    For Each oStatement In oSwitchBlock(1).SubTree
        Select Case oStatement(2).Index
            Case 1 ' default
                FormatSwitchStatement = FormatSwitchStatement & sIndentation1 & "default:" & NewLine
            Case 2 ' case
                FormatSwitchStatement = FormatSwitchStatement & sIndentation1 & "case " & FormatExpression(oStatement(2)(1)!expression) & ":" & NewLine
            Case 3 ' statement
                FormatSwitchStatement = FormatSwitchStatement & FormatStatement(oStatement(2)!statement, lIndentationLevel + 1)
        End Select
    Next
    FormatSwitchStatement = FormatSwitchStatement & sIndentation & "}" & NewLine
End Function

Private Function FormatForeachStatement(oTree As ParseTree, ByVal lIndentationLevel As Long) As String
    Dim sIndentation As String
    Dim sIndentation1 As String
    
    sIndentation = Indentation(lIndentationLevel)
    sIndentation1 = Indentation(lIndentationLevel + 1)
    FormatForeachStatement = sIndentation & "foreach ("
    FormatForeachStatement = FormatForeachStatement & FormatExpression(oTree!expression) & " as "
    FormatForeachStatement = FormatForeachStatement & FormatExpression(oTree!expression1) & ")"
    FormatForeachStatement = FormatForeachStatement & FormatBlock(oTree!block, lIndentationLevel) & NewLine
End Function

Private Function FormatForStatement(oTree As ParseTree, ByVal lIndentationLevel As Long) As String
    Dim sIndentation As String
    Dim sIndentation1 As String
    
    sIndentation = Indentation(lIndentationLevel)
    sIndentation1 = Indentation(lIndentationLevel + 1)
    
    FormatForStatement = sIndentation & "for ("
    If oTree(1)(1).Index = 1 Then
        FormatForStatement = FormatForStatement & FormatExpression(oTree(1)(1)(1)) & "; "
    Else
        FormatForStatement = FormatForStatement & "; "
    End If
    If oTree(1)(2).Index = 1 Then
        FormatForStatement = FormatForStatement & FormatExpression(oTree(1)(2)(1)) & "; "
    Else
        FormatForStatement = FormatForStatement & "; "
    End If
    If oTree(1)(3).Index = 1 Then
        FormatForStatement = FormatForStatement & FormatExpression(oTree(1)(3)(1))
    End If
    FormatForStatement = FormatForStatement & ") "
    FormatForStatement = FormatForStatement & FormatBlock(oTree!block, lIndentationLevel) & NewLine
End Function

Private Function FormatIfStatement(oTree As ParseTree, ByVal lIndentationLevel As Long) As String
    Dim sIndentation As String
    Dim sIndentation1 As String
    
    sIndentation = Indentation(lIndentationLevel)
    sIndentation1 = Indentation(lIndentationLevel + 1)
    
    FormatIfStatement = sIndentation & "if ("
    FormatIfStatement = FormatIfStatement & FormatExpression(oTree!expression)
    FormatIfStatement = FormatIfStatement & ") "
    FormatIfStatement = FormatIfStatement & FormatBlock(oTree!block, lIndentationLevel)
    If oTree(3).Index = 1 Then
        FormatIfStatement = FormatIfStatement & IIf(Right$(FormatIfStatement, 2) = NewLine, sIndentation, "") & " else "
        Select Case oTree(3)(1)(1).Index
            Case 1 ' if statement
                FormatIfStatement = FormatIfStatement & TrimText(FormatIfStatement(oTree(3)(1)(1)!if_statement, lIndentationLevel)) & NewLine
            Case 2 ' block
                FormatIfStatement = FormatIfStatement & FormatBlock(oTree(3)(1)(1)!block, lIndentationLevel) & NewLine
        End Select
    Else
        FormatIfStatement = FormatIfStatement & NewLine
    End If
End Function

Private Function FormatTryCatchStatement(oTree As ParseTree, ByVal lIndentationLevel As Long) As String
    Dim sIndentation As String
    Dim sIndentation1 As String
    
    sIndentation = Indentation(lIndentationLevel)
    sIndentation1 = Indentation(lIndentationLevel + 1)
    
    FormatTryCatchStatement = sIndentation & "try "
    FormatTryCatchStatement = FormatTryCatchStatement & FormatBlock(oTree!block, lIndentationLevel)
    FormatTryCatchStatement = FormatTryCatchStatement & " catch ("
    FormatTryCatchStatement = FormatTryCatchStatement & FormatExpression(oTree!expression)
    FormatTryCatchStatement = FormatTryCatchStatement & ") "
    FormatTryCatchStatement = FormatTryCatchStatement & FormatBlock(oTree(3), lIndentationLevel) & NewLine
End Function

Private Function FormatWithStatement(oTree As ParseTree, ByVal lIndentationLevel As Long) As String
    Dim sIndentation As String
    Dim sIndentation1 As String
    
    sIndentation = Indentation(lIndentationLevel)
    sIndentation1 = Indentation(lIndentationLevel + 1)
    
    FormatWithStatement = sIndentation & "with ("
    FormatWithStatement = FormatWithStatement & FormatExpression(oTree!expression)
    FormatWithStatement = FormatWithStatement & ") "
    FormatWithStatement = FormatWithStatement & FormatBlock(oTree!block, lIndentationLevel) & NewLine
End Function

Private Function FormatAnonymousFunctionDeclaration(oTree As ParseTree, ByVal lIndentationLevel As Long) As String
    Dim sIndentation As String
    Dim sIndentation1 As String
    
    sIndentation = Indentation(lIndentationLevel)
    sIndentation1 = Indentation(lIndentationLevel + 1)
    
    FormatAnonymousFunctionDeclaration = FormatAnonymousFunctionDeclaration & sIndentation
    
    FormatAnonymousFunctionDeclaration = FormatAnonymousFunctionDeclaration & "function "
    If oTree(2).Index = 1 Then
        FormatAnonymousFunctionDeclaration = FormatAnonymousFunctionDeclaration & "("
        FormatAnonymousFunctionDeclaration = FormatAnonymousFunctionDeclaration & FormatParameterList(oTree(2)(1)!parameter_list)
        FormatAnonymousFunctionDeclaration = FormatAnonymousFunctionDeclaration & ") "
    End If
    FormatAnonymousFunctionDeclaration = FormatAnonymousFunctionDeclaration & FormatBlock(oTree(3), lIndentationLevel) & IIf(lIndentationLevel = -1, "", NewLine)
End Function

Private Function FormatFunctionDeclaration(oTree As ParseTree, ByVal lIndentationLevel As Long) As String
    Dim sIndentation As String
    Dim sIndentation1 As String
    
    sIndentation = Indentation(lIndentationLevel)
    sIndentation1 = Indentation(lIndentationLevel + 1)
    
    FormatFunctionDeclaration = FormatFunctionDeclaration & sIndentation
    
    If oTree(1).Index = 1 Then
        FormatFunctionDeclaration = FormatFunctionDeclaration & LCase$(oTree(1).Text) & " "
    End If
    FormatFunctionDeclaration = FormatFunctionDeclaration & "function "
    FormatFunctionDeclaration = FormatFunctionDeclaration & oTree(3).Text
    If oTree(4).Index = 1 Then
        FormatFunctionDeclaration = FormatFunctionDeclaration & "("
        FormatFunctionDeclaration = FormatFunctionDeclaration & FormatParameterList(oTree(4)(1)!parameter_list)
        FormatFunctionDeclaration = FormatFunctionDeclaration & ") "
    End If
    FormatFunctionDeclaration = FormatFunctionDeclaration & FormatBlock(oTree(5), lIndentationLevel) & NewLine
End Function

Private Function FormatClassDeclaration(oTree As ParseTree, ByVal lIndentationLevel As Long) As String
    Dim sIndentation As String
    Dim sIndentation1 As String
    
    sIndentation = Indentation(lIndentationLevel)
    sIndentation1 = Indentation(lIndentationLevel + 1)
    
    FormatClassDeclaration = sIndentation & "class "
    FormatClassDeclaration = FormatClassDeclaration & oTree(2).Text & " "
    If oTree(3).Index = 1 Then
        FormatClassDeclaration = FormatClassDeclaration & "extends " & oTree(3)(1)(2).Text & " "
    End If
    FormatClassDeclaration = FormatClassDeclaration & FormatBlock(oTree(4), lIndentationLevel) & NewLine
End Function



Private Function FormatParameterList(oTree As ParseTree) As String
    Dim oParameter As ParseTree
    Dim vParameters As Variant
    
    vParameters = Array()

    For Each oParameter In oTree.SubTree
        ReDim Preserve vParameters(UBound(vParameters) + 1)
        vParameters(UBound(vParameters)) = IIf(oParameter(1).Index = 1, "& ", "") & FormatExpression(oParameter(2))
    Next
    FormatParameterList = Join(vParameters, ", ")
End Function

Private Function FormatBlock(oTree As ParseTree, ByVal lIndentationLevel As Long) As String
    Dim sIndentation As String
    Dim bOneLine As Boolean
    
    bOneLine = lIndentationLevel = -1
    
    sIndentation = Indentation(lIndentationLevel)

    Select Case oTree.Index
        Case 2 ' statement
            FormatBlock = FormatBlock & IIf(bOneLine, "", NewLine) & FormatStatement(oTree!statement, lIndentationLevel)
        Case 1 ' block
            FormatBlock = "{" & IIf(bOneLine, "", NewLine)
            FormatBlock = FormatBlock & FormatStatements(oTree(1)!statements, lIndentationLevel)
            FormatBlock = FormatBlock & sIndentation & "}"
    End Select
End Function


Private Function FormatCSS(oTree As ParseTree, ByVal lIndentationLevel As Long) As String
    Dim oStatement As ParseTree
    Dim sIndentation As String
    Dim sIndentation1 As String
    
    sIndentation = Indentation(lIndentationLevel)
    sIndentation1 = Indentation(lIndentationLevel + 1)
    
    FormatCSS = sIndentation & FormatOpenTag(oTree!open_tag, lIndentationLevel) & NewLine
    FormatCSS = FormatCSS & FormatCSSStatements(oTree(2)(1)!css_statements, lIndentationLevel)
    FormatCSS = FormatCSS & sIndentation & "</style>"
End Function

Private Function FormatCSSStatements(oTree As ParseTree, ByVal lIndentationLevel As Long) As String
    Dim oStatement As ParseTree
    Dim sIndentation As String
    Dim sIndentation1 As String
    Dim vList As Variant
    
    sIndentation = Indentation(lIndentationLevel)
    sIndentation1 = Indentation(lIndentationLevel + 1)

    vList = Array()
    For Each oStatement In oTree.SubTree
        Select Case oStatement.Index
            Case 1
                AddArray vList, oStatement.Text
            Case 2
                AddArray vList, FormatCSSStatement(oStatement(1), lIndentationLevel)
        End Select
    Next
    
    FormatCSSStatements = Join(vList, NewLine)
End Function


Private Function FormatCSSStatement(oTree As ParseTree, ByVal lIndentationLevel As Long) As String
    Dim oStatement As ParseTree
    Dim sIndentation As String
    Dim sIndentation1 As String
    
    sIndentation = Indentation(lIndentationLevel)
    sIndentation1 = Indentation(lIndentationLevel + 1)
    
    FormatCSSStatement = FormatCSSStatement & sIndentation & FormatCSSSelector(oTree!css_selector)
    If oTree(2).Index = 1 Then
        FormatCSSStatement = FormatCSSStatement & " (" & FormatCSSAttributeList(oTree(2)(1)(1), 0, False) & ")"
    End If
    FormatCSSStatement = FormatCSSStatement & " {" & FormatCSSAttributeList(oTree(4), lIndentationLevel)
    FormatCSSStatement = FormatCSSStatement & NewLine & sIndentation & "}" & NewLine
End Function

Private Function FormatCSSSelector(oTree As ParseTree) As String
    Dim vList As Variant
    Dim oPart As ParseTree
    
    vList = Array()
    For Each oPart In oTree.SubTree
        If oPart.Text <> "" Then
            AddArray vList, oPart.Text
        End If
    Next
    FormatCSSSelector = Join(vList, " ")
End Function

Private Function FormatCSSWhiteSpace(oTree As ParseTree) As String
    Dim oChar As ParseTree
    Dim vList As Variant
    
    vList = Array()
    For Each oChar In oTree.SubTree
        AddArray vList, oChar.Text
    Next
    FormatCSSWhiteSpace = Join(vList, "")
End Function

Private Function FormatCSSAttributeList(oTree As ParseTree, ByVal lIndentationLevel As Long, Optional ByVal bIndent As Boolean = True) As String
    Dim oAttribute As ParseTree
    Dim sIndentation As String
    Dim sIndentation1 As String
    Dim vList As Variant
    
    sIndentation = Indentation(lIndentationLevel)
    sIndentation1 = Indentation(lIndentationLevel + 1)
    
    vList = Array()
    
    For Each oAttribute In oTree.SubTree
        If oAttribute.Description = "OR" Then
            Select Case oAttribute.Index
                Case 1 ' comment
                    AddArray vList, IIf(bIndent, NewLine & sIndentation1, "") & oAttribute.Text
                Case 2 ' css statement
                    AddArray vList, IIf(bIndent, NewLine, "") & FormatCSSStatement(oAttribute(1), lIndentationLevel + 1)
                Case 3 ' css attribute
                    AddArray vList, IIf(bIndent, NewLine & sIndentation1, "") & FormatCSSAttribute(oAttribute(1)) & ";"
                Case 4 ' semicolon
                Case 5 ' whitespace
            End Select
            
        End If
    Next
    FormatCSSAttributeList = Join(vList, "")
End Function

Private Function FormatCSSAttribute(oTree As ParseTree) As String
    FormatCSSAttribute = FormatCSSAttribute & oTree(1).Text & ":" & FormatCSSAttribute & oTree(3).Text
End Function

Private Function FormatExpression(oTree As ParseTree) As String
    Select Case oTree.Index
        Case 1 ' ternary
            FormatExpression = FormatExpressionSub(oTree(1)(1)) & "?" & FormatExpressionSub(oTree(1)(2)) & ":" & FormatExpressionSub(oTree(1)(3))
        Case 2 ' expression sub
            FormatExpression = FormatExpressionSub(oTree(1))
    End Select
End Function

Private Function FormatExpressionSub(oTree As ParseTree) As String
    Dim oPart As ParseTree
    
    For Each oPart In oTree.SubTree
        Select Case oPart.Index
            Case 0 ' operator
                Select Case oPart.Text
                    Case "."
                        FormatExpressionSub = FormatExpressionSub & oPart.Text
                    Case Else
                        FormatExpressionSub = FormatExpressionSub & " " & oPart.Text & " "
                End Select
            Case 1 ' number
                FormatExpressionSub = FormatExpressionSub & oPart.Text
            Case 2 ' array literal
                FormatExpressionSub = FormatExpressionSub & "[" & FormatExpressionList(oPart(1)(1)) & "]"
            Case 3 ' anonymous function declaration
                FormatExpressionSub = FormatExpressionSub & FormatAnonymousFunctionDeclaration(oPart(1), -1)
            Case 4 ' assignment
                FormatExpressionSub = FormatExpressionSub & FormatAssignment(oPart!assignment)
            Case 5 ' function call
                FormatExpressionSub = FormatExpressionSub & FormatFunctionCall(oPart(1))
            Case 6 ' unary post
                FormatExpressionSub = FormatExpressionSub & FormatVariable(oPart(1)(1)) & oPart(1)(2).Text
            Case 7  ' unary
                FormatExpressionSub = FormatExpressionSub & oPart(1)(1).Text & FormatVariable(oPart(1)(2))
            Case 8 ' variable
                FormatExpressionSub = FormatExpressionSub & FormatVariable(oPart!variable)
            Case 9 ' string
                FormatExpressionSub = FormatExpressionSub & oPart.Text
            Case 10 ' reg exp
                FormatExpressionSub = FormatExpressionSub & oPart.Text
            Case 11 ' unary other
                FormatExpressionSub = FormatExpressionSub & FormatUnaryOther(oPart!unary_other)
            Case 12 ' bracketed
                FormatExpressionSub = FormatExpressionSub & "(" & FormatExpression(oPart!bracketed!expression) & ")"
        End Select
    Next
End Function



Private Function FormatUnaryOther(oTree As ParseTree) As String
    Select Case oTree(1).Index
        Case 1 ' !
            FormatUnaryOther = oTree(1).Text
        Case 2 ' (cast)
            FormatUnaryOther = "(" & oTree(1).Text & ") "
    End Select
    
    FormatUnaryOther = FormatUnaryOther & FormatExpressionSub(oTree!expression_sub)
End Function

Private Function FormatFunctionCall(oTree As ParseTree) As String
    Dim oParameter As ParseTree
    Dim vParameters As Variant
    
    vParameters = Array()
    If oTree(1).Index = 1 Then
        FormatFunctionCall = "new "
    End If
    If oTree(2).Index = 1 Then
        FormatFunctionCall = FormatFunctionCall & oTree(2)(1).Text & "::"
    End If
    
    FormatFunctionCall = FormatFunctionCall & oTree(3).Text & "("
    FormatFunctionCall = FormatFunctionCall & FormatExpressionList(oTree(4))
    FormatFunctionCall = FormatFunctionCall & Join(vParameters, ", ") & ")"
End Function

Private Function FormatAssignment(oTree As ParseTree) As String

    FormatAssignment = IIf(oTree(1).Index = 1, LCase$(oTree(1).Text) & " ", "")
    Select Case oTree(2).Index
        Case 1 ' function assignment
            FormatAssignment = FormatAssignment & FormatFunctionAssignment(oTree(2)(1))
        Case 2 ' variable assignment
            FormatAssignment = FormatAssignment & FormatVariable(oTree(2)(1))
    End Select
    FormatAssignment = FormatAssignment & " " & oTree!assignment_operator.Text & " "
    Select Case oTree(4).Index
        Case 1
            FormatAssignment = FormatAssignment & FormatFunctionDeclaration(oTree(4)(1), 0)
        Case 2
            FormatAssignment = FormatAssignment & FormatExpression(oTree(4)(1))
    End Select
End Function

Private Function FormatFunctionAssignment(oTree As ParseTree) As String
    FormatFunctionAssignment = oTree(1).Text & "("
    FormatFunctionAssignment = FormatFunctionAssignment & FormatExpressionList(oTree(2))
    FormatFunctionAssignment = FormatFunctionAssignment & ")"
End Function

Private Function FormatVariable(oTree As ParseTree) As String
    Dim oIndex As ParseTree
    Dim oExpression As ParseTree
    Dim vExpression As Variant
    
    FormatVariable = oTree(1).Text
    If oTree!Index.Index > 0 Then
        For Each oIndex In oTree!Index.SubTree
            FormatVariable = FormatVariable & "[" & FormatExpressionList(oIndex(1)!expression_list) & "]"
        Next
    End If
End Function

Private Function FormatExpressionList(oTree As ParseTree) As String
    Dim oParameter As ParseTree
    Dim vParameters As Variant
    Dim sKey As String
    
    vParameters = Array()

    For Each oParameter In oTree.SubTree
        If oParameter.Index = 1 Then
            ReDim Preserve vParameters(UBound(vParameters) + 1)
            If oParameter(1).Index = 1 Then
                sKey = oParameter(1)(1).Text & " => "
            Else
                sKey = ""
            End If
            vParameters(UBound(vParameters)) = sKey & FormatExpression(oParameter(1)(2))
        Else
            ReDim Preserve vParameters(UBound(vParameters) + 1)
            vParameters(UBound(vParameters)) = ""
        End If
    Next
    FormatExpressionList = Join(vParameters, ", ")
End Function

Private Function Indentation(ByVal lDepth As Long) As String
    If lDepth < 0 Then
        lDepth = 0
    End If
    If Not mbNotIndented Then
        Indentation = String$(lDepth, vbTab)
    Else
        Indentation = " "
    End If
End Function

Private Function NewLine() As String
    If Not mbNotIndented Then
        NewLine = vbCrLf
    End If
End Function

Private Sub AddArray(vArray As Variant, sItem As String)
    Dim lUbound As Long
    
    lUbound = UBound(vArray)
    
    ReDim Preserve vArray(lUbound + 1)
    vArray(lUbound + 1) = sItem
End Sub
