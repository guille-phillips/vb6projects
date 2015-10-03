Attribute VB_Name = "NewCompiler6502"
Option Explicit

Private moProgram As ISaffronObject
Private mlPC As Long

Public Sub LoadAssembly(ByVal sFile As String)
    Dim sAssembly As String
    Dim sPath As String

    sPath = App.Path & "\" & sFile
    Open sPath For Binary As #1
    sAssembly = Space$(FileLen(sPath))
    Get #1, , sAssembly
    Close #1
    Compile sAssembly
End Sub

Private Function Compile(sText As String) As String
    InitialiseCompiler
    
    If CompileCode(sText, False) Then
        CompileCode sText, True
    End If
End Function

Private Sub InitialiseCompiler()
    Dim oParser As SaffronClasses.ISaffronObject
    
    Dim sRules As String
    sRules = Space$(FileLen(App.Path & "\NewCompiler.saf"))
    
    Open App.Path & "\NewCompiler.saf" For Binary As #1
    Get #1, , sRules
    Close #1
    
    If Not SaffronCompiler.CreateRules(sRules) Then
        MsgBox "Bad Saf"
        End
    End If

End Sub

Private Function CompileCode(sText As String, bResolveLabels As Boolean) As Boolean
    Dim oTree As New SaffronTree
    
    Set moProgram = SaffronCompiler.Rules("level0")

    SaffronStream.Text = sText
    If moProgram.Parse(oTree) Then
        CompileLevel0 oTree
    End If
    Stop
    
    
End Function

Private Function CompileLevel0(oTree As SaffronTree) As String
    Dim moSymbols As New Dictionary
    Dim lBase As Long
    
    If oTree.SubTree.Count = 1 Then
        Select Case oTree(1)(1).Index
            Case 1 ' base
                lBase = 10
                Select Case LCase$(oTree(1)(1)(1)(4).Text)
                    Case "h"
                        lBase = 16
                    Case "d"
                        lBase = 10
                    Case "b"
                        lBase = 2
                End Select
                mlPC = ConvertBase(oTree(1)(1)(1)(3).Text, lBase)
            Case 2 ' identifier
                moSymbols.Add oTree(1)(1).Text, mlPC
        End Select
    End If
End Function

Private Function Number(oTree) As Variant
    Dim lBase As Long
    
    lBase = 10
    Select Case LCase$(oTree(1)(1)(1)(4).Text)
        Case "h"
            lBase = 16
        Case "d"
            lBase = 10
        Case "b"
            lBase = 2
    End Select
    mlPC = ConvertBase(oTree(1)(1)(1)(3).Text, lBase)
End Function

Private Function ConvertBase(sNumber As String, lBase As Long) As Long
    Dim sDigits As String
    Dim lSum As Long
    Dim lPos As Long
    
    sDigits = "0123456789ABCDEF"
    For lPos = 1 To Len(sNumber)
        lSum = lSum * lBase
        lSum = lSum + InStr(sDigits, Mid$(sNumber, lPos, 1)) - 1
    Next
End Function

Private Function NumberSize(sNumber As String, lBase As Long) As Long
    
End Function
