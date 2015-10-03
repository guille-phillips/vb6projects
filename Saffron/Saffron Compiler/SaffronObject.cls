VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SaffronObject"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public ErrorString As String

Private oRulesObjects As Collection

Public Property Get Rules() As Collection
    Set Rules = oRulesObjects
End Property

Public Function CreateRules(sLanguageDefinitionString As String) As Boolean
    Dim oLanguageLexicalAnalyser As New CLanguageLexAn
    Dim oLanguageCompiler As New CCompileLanguage
    Dim oParsedLanguage As New SaffronTree

    ErrorString = ""
    
    Set oRulesObjects = Nothing
    Set oRulesObjects = New Collection
    
    ' Load in the lexical analysis description
    SaffronStream.Text = sLanguageDefinitionString
    oLanguageLexicalAnalyser.CreateStructure
    If oLanguageLexicalAnalyser.Parse(oParsedLanguage) Then
        Set oRulesObjects = oLanguageCompiler.CompileLanguage(oParsedLanguage)
        CreateRules = True
    Else
        ErrorString = "Compile error at position " & oParsedLanguage.TextStart & vbCrLf
        ErrorString = ErrorString & Left(sLanguageDefinitionString, oParsedLanguage.TextStart)
    End If
End Function

