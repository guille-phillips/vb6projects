VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CLanguageLexAn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private moObjects As New Collection
Private moRules As ISaffronObject

Friend Sub CreateStructure()
    With moObjects
        ' Single chars
        Create("Pipe", "Literal").Initialise False, False, "|"
        Create("Blank", "In").Initialise False, False, " " & vbCr & vbLf & vbTab
        
        ' Repeat chars
        Create("WS", "List").Initialise False, True, !Blank
        Create("WSNon", "List").Initialise True, True, !Blank
        Create("EscapeWSOmit", "And").Initialise False, True, !Pipe, Create(Empty, "Or").Initialise(False, False, Create(Empty, "EOS").Initialise(False, False), !WSNon)
        Create("EscapeWS", "And").Initialise False, False, !Pipe, Create(Empty, "Or").Initialise(False, False, Create(Empty, "EOS").Initialise(False, False), !WSNon)
        
        ' Keywords (with non-alpha)
        Create("AndKeywordWS", "And").Initialise False, True, Create("AndKeyword", "Literal").Initialise(False, False, "AND"), !WSNon
        Create("OrKeywordWS", "And").Initialise False, True, Create("OrKeyword", "Literal").Initialise(False, False, "OR"), !WSNon
        Create("InKeywordWS", "And").Initialise False, True, Create("InKeyword", "Literal").Initialise(False, False, "IN"), !WSNon
        Create("SensitiveKeywordWS", "And").Initialise False, True, Create("SensitiveKeyword", "Literal").Initialise(False, False, "SEN"), !WSNon
        Create("ToKeywordWS", "And").Initialise False, True, Create("ToKeyword", "Literal").Initialise(False, False, "TO"), !WSNon
        Create("UntilKeywordWS", "And").Initialise False, True, Create("UntilKeyword", "Literal").Initialise(False, False, "UNTIL"), !WSNon
        Create("ListKeywordWS", "And").Initialise False, True, Create("ListKeyword", "Literal").Initialise(False, False, "LIST"), !WSNon
        Create("DelimitKeywordWS", "And").Initialise False, True, Create("DelimitKeyword", "Literal").Initialise(False, False, "DEL"), !WSNon
        Create("MinKeywordWS", "And").Initialise False, True, Create("MinKeyword", "Literal").Initialise(False, False, "MIN"), !WSNon
        Create("MaxKeywordWS", "And").Initialise False, True, Create("MaxKeyword", "Literal").Initialise(False, False, "MAX"), !WSNon
        Create("NotKeywordWS", "And").Initialise False, True, Create("NotKeyword", "Literal").Initialise(False, False, "NOT"), !WSNon
        Create("OptionalKeywordWS", "And").Initialise False, True, Create("OptionalKeyword", "Literal").Initialise(False, False, "OPT"), !WSNon
        Create("EOSKeywordWS", "And").Initialise False, True, Create("EOSKeyword", "Literal").Initialise(False, False, "EOS"), !WSNon
        Create("BOSKeywordWS", "And").Initialise False, True, Create("BOSKeyword", "Literal").Initialise(False, False, "BOS"), !WSNon
        Create("FailKeywordWS", "And").Initialise False, True, Create("FailKeyword", "Literal").Initialise(False, False, "FAIL"), !WSNon
        Create("PassKeywordWS", "And").Initialise False, True, Create("PassKeyword", "Literal").Initialise(False, False, "PASS"), !WSNon
        Create("PermKeywordWS", "And").Initialise False, True, Create("PermKeyword", "Literal").Initialise(False, False, "PERM"), !WSNon
        Create("ExternalKeywordWS", "And").Initialise False, True, Create("ExternalKeyword", "Literal").Initialise(False, False, "EXT"), !WSNon
        Create("EachKeywordWS", "And").Initialise False, True, Create("EachKeyword", "Literal").Initialise(False, False, "EACH"), !WSNon
        Create("SkipKeywordWS", "And").Initialise False, True, Create("SkipKeyword", "Literal").Initialise(False, False, "SKIP"), !WSNon
        Create("OmitKeywordWS", "And").Initialise False, True, Create("OmitKeyword", "Literal").Initialise(False, False, "OMIT"), !WSNon
        Create("NonKeywordWS", "And").Initialise False, False, Create("NonKeyword", "Literal").Initialise(False, False, "NON"), !WSNon
        
        ' RuleName
        Create("RuleNameChar", "In").Initialise False, False, "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz_0123456789"
        Create("RuleNameStartChar", "In").Initialise False, False, "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz"
        Create("RepeatRuleNameChar", "List").Initialise False, False, !RuleNameChar, , , 0
        Create("RuleName", "And").Initialise False, False, !RuleNameStartChar, !RepeatRuleNameChar, !WSNon
        
        ' Number
        Create("NumberChar", "In").Initialise False, False, "0123456789"
        Create("Figure", "List").Initialise False, False, !NumberChar
        
        ' IN Phrases
        Create("SingleLiteral", "And").Initialise False, False, Create(Empty, "Skip").Initialise(False, False, 1), !WSNon
        Create("CharacterCode", "List").Initialise False, False, !NumberChar, , !WSNon, 3, 3
        Create("SingleLiteralOrCharacterCode", "Or").Initialise False, False, !CharacterCode, !SingleLiteral
        Create("RangeStatement", "And").Initialise False, False, !SingleLiteralOrCharacterCode, !WS, !ToKeywordWS, !WS, !SingleLiteralOrCharacterCode

        ' Literal
        Create("Character", "Or").Initialise False, False, !EscapeWS, Create(Empty, "Skip").Initialise(False, False, 1)
        Create("LiteralList", "List").Initialise False, False, !Character, , !WSNon
        Create("Literal", "And").Initialise False, False, Create(Empty, "List").Initialise(False, False, Create(Empty, "And").Initialise(False, False, !SensitiveKeywordWS, !WS), , , 0, 1), Create(Empty, "Or").Initialise(False, False, !CharacterCode, !LiteralList)
        
        ' In Phrases
        Create("InLiteralExp", "Or").Initialise False, False, !RangeStatement, !CharacterCode, !Literal
        Create("OptionalSensitive", "List").Initialise False, False, Create(Empty, "And").Initialise(False, False, !SensitiveKeywordWS, !WS), , , 0, 1
        Create("OptionalNot", "List").Initialise False, False, Create(Empty, "And").Initialise(False, False, !NotKeywordWS, !WS), , , 0, 1
        Create("InSubExp", "And").Initialise False, False, !OptionalNot, !OptionalSensitive, !InLiteralExp
        Create("InPhrase", "List").Initialise False, False, Create(Empty, "And").Initialise(False, False, !InSubExp, !WS), , !EscapeWSOmit

'        ' Phrase parts
        Create "Phrase", "And"
        Create("PhraseList", "List").Initialise False, False, Create(Empty, "And").Initialise(False, False, !Phrase, !WS), , !EscapeWSOmit, 0
        
'        ' Functions
        Create("AndFunction", "And").Initialise False, False, !AndKeywordWS, !WS, !PhraseList
        Create("OrFunction", "And").Initialise False, False, !OrKeywordWS, !WS, !PhraseList
        Create("EachFunction", "And").Initialise False, False, !EachKeywordWS, !WS, !PhraseList
        Create("PermFunction", "And").Initialise False, False, !PermKeywordWS, !WS, !PhraseList
        
        Create("OptionalDelimitClause", "List").Initialise False, False, Create(Empty, "And").Initialise(False, False, !DelimitKeywordWS, !WS, !Phrase, !WS), , , 0, 1
        Create("OptionalUntilClause", "List").Initialise False, False, Create(Empty, "And").Initialise(False, False, !UntilKeywordWS, !WS, !Phrase, !WS), , , 0, 1
        Create("OptionalMinClause", "List").Initialise False, False, Create(Empty, "And").Initialise(False, False, !MinKeywordWS, !WS, !Figure, !WS), , , 0, 1
        Create("OptionalMaxClause", "List").Initialise False, False, Create(Empty, "And").Initialise(False, False, !MaxKeywordWS, !WS, !Figure, !WS), , , 0, 1
                                
        Create("ListFunction", "And").Initialise False, False, !ListKeywordWS, !WS, !Phrase, !WS, !OptionalDelimitClause, !OptionalUntilClause, !OptionalMinClause, !OptionalMaxClause, !EscapeWSOmit
        Create("OptionalFunction", "And").Initialise False, False, !OptionalKeywordWS, !WS, !Phrase, !WS, !EscapeWSOmit
        Create("ExternalFunction", "And").Initialise False, False, !ExternalKeywordWS, !WS, !RuleName, !WS, Create(Empty, "List").Initialise(False, False, Create(Empty, "And").Initialise(False, False, !Literal, !WS), , !EscapeWSOmit, 0)
        Create("InFunction", "And").Initialise False, False, !InKeywordWS, !WS, !InPhrase
        Create("NotFunction", "And").Initialise False, False, !NotKeywordWS, !WS, !Phrase, !WS, !EscapeWSOmit
        Create("SkipFunction", "And").Initialise False, False, !SkipKeywordWS, !WS, Create(Empty, "List").Initialise(False, False, Create(Empty, "And").Initialise(False, False, !Figure, !WS), , , 0, 1), !EscapeWSOmit
        
        Create("OptionalNon", "List").Initialise False, False, Create(Empty, "And").Initialise(False, False, !NonKeywordWS, !WS), , , 0, 1
        Create("OptionalOmit", "List").Initialise False, False, Create(Empty, "And").Initialise(False, False, !OmitKeywordWS, !WS), , , 0, 1
        
'        ' Phrases
        Create("PhraseSub", "Or").Initialise False, False, !AndFunction, !OrFunction, !PermFunction, !EachFunction, !ListFunction, !OptionalFunction, !ExternalFunction, !InFunction, !NotFunction, !SkipFunction, !EOSKeywordWS, !BOSKeywordWS, !PassKeywordWS, !FailKeywordWS, !Literal
        !Phrase.Initialise False, False, !OptionalNon, !OptionalOmit, !PhraseSub
        
'        ' Statements
        Create("Rule", "And").Initialise False, False, !RuleName, !WS, !Phrase, !WS, !EscapeWSOmit
        Create("Rules", "List").Initialise False, False, !Rule, !WS
        
        Set moRules = !Rules

'        SaffronStream.Text = "nonext skip  |"
'        Dim oResult As New SaffronTree
'
'        Set oResult = New SaffronTree
'        If !Phrase.Parse(oResult) Then
'            Stop
'        Else
'            Debug.Print Left$(SaffronStream.Text, SaffronStream.Position)
'        End If
        
        Set moObjects = Nothing
    End With
End Sub

Friend Function Parse(oResult As SaffronTree) As Boolean
    Set oResult = New SaffronTree
    Parse = moRules.Parse(oResult)
End Function

Public Function Create(vName As Variant, sType As String) As ISaffronObject
    Set Create = CreateObject("SaffronClasses.Saffron" & sType)
    If Not IsMissing(vName) And Not IsEmpty(vName) Then
        moObjects.Add Create, CStr(vName)
    End If
End Function

