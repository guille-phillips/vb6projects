Attribute VB_Name = "SharedFunctions"
Option Explicit

Public goNet As New clsNode

Public Function OpenTextFile(ByVal sName As String) As String
    OpenTextFile = String$(FileLen(App.Path & "\" & sName), Chr$(0))
    Open App.Path & "\" & sName For Binary As #1
    Get #1, , OpenTextFile
    Close #1
End Function

Public Function SaveTextFile(ByVal sName As String, ByVal sContents As String) As String
    Kill App.Path & "\" & sName
    Open App.Path & "\" & sName For Binary As #1
    Put #1, , sContents
    Close #1
End Function
