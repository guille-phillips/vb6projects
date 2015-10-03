Attribute VB_Name = "Debugging"
Option Explicit

Public Sub OpenDebugFile()
    Open App.path & "\debug.txt" For Binary As #4
End Sub

Public Sub WriteString(ByVal sString As String)
    Put #4, , sString & vbCrLf
End Sub

Public Sub CloseDebugFile()
    Close #4
End Sub
