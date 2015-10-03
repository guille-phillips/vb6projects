VERSION 5.00
Begin VB.Form frmConsole 
   BackColor       =   &H00000000&
   Caption         =   "Adventure Engine"
   ClientHeight    =   9435
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10560
   ForeColor       =   &H00FFFFFF&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   9435
   ScaleWidth      =   10560
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtConsole 
      Height          =   6615
      Left            =   0
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   0
      Width           =   6015
   End
End
Attribute VB_Name = "frmConsole"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private msErrorMessage As String
Private moLanguageLex As ISaffronObject

Private moLocation As clsNode
Private moPlayer As clsNode
Private msLanguage As String

Private Sub Form_Load()
    msErrorMessage = "I don't understand."
    LoadLanguageCore
    goNet.LoadFile
    
    Set moLocation = goNet("the clearing")
    Set moPlayer = goNet("player")
    DescribeLocation
End Sub


Private Sub DescribeLocation()
    InsertText moLocation("description").DescriptorList & vbCrLf & vbCrLf
End Sub

Private Sub Form_Resize()
    txtConsole.Width = Me.Width
    txtConsole.Height = Me.Height
End Sub

Private Sub LoadLanguageCore()
    Dim sDef As String
    
    msLanguage = OpenTextFile("language.txt")
End Sub

Private Sub CompileLanguage()
    Dim sLanguage As String
    
    sLanguage = msLanguage & CompileVerbs & CompileNouns
    If Not CreateRules(sLanguage) Then
        MsgBox "Bad Def!"
        End
    End If
    Set moLanguageLex = SaffronObject.Rules("sentence")
End Sub

Private Function CompileVerbs() As String
    Dim sVerbs As String
    
    sVerbs = "verb or "
    
    sVerbs = sVerbs & "take "
    sVerbs = sVerbs & "| | "
    CompileVerbs = sVerbs
End Function

Private Function CompileNouns() As String
    Dim sNouns As String
    
    sNouns = "noun or "
    
    sNouns = sNouns & "torch cat "
    sNouns = sNouns & "| | "
    CompileNouns = sNouns
End Function


Private Sub SaveGameNet()
    Dim sDescriptor As String
    
    SaveTextFile "game_net.txt", goNet.Descriptor(False)
End Sub


Private Sub Form_KeyPress(KeyAscii As Integer)
    Dim lCursorPos As Long
    Dim vLines As Variant
    Dim sCommand As String
    Dim oInputTree As SaffronTree
    
    Select Case KeyAscii
        Case 13
            CompileLanguage
            
            lCursorPos = txtConsole.SelStart
            vLines = Split(Left(txtConsole.Text, lCursorPos), vbCrLf)
            If UBound(vLines) <> -1 Then
                sCommand = vLines(UBound(vLines))
            End If
            KeyAscii = 0
            
            SaffronStream.Text = sCommand
            Set oInputTree = New SaffronTree
            
            If moLanguageLex.Parse(oInputTree) Then
                InsertText PerformAction(oInputTree), True
            Else
                If Trim$(sCommand) <> "" Then
                    InsertText msErrorMessage, True
                Else
                    InsertText vbCrLf
                End If
            End If
        Case 27
            Unload Me
    End Select
End Sub

Private Function InsertText(ByVal sText As String, Optional ByVal bNewLine As Boolean)
    Dim iTextPos As Integer
    
    If bNewLine Then
        sText = vbCrLf & sText & vbCrLf & vbCrLf
    End If
    iTextPos = txtConsole.SelStart
    txtConsole.Text = Left$(txtConsole.Text, iTextPos) & sText & Mid$(txtConsole.Text, iTextPos + 1)
    
    txtConsole.SelStart = iTextPos + Len(sText)
End Function

Private Function PerformAction(oInputTree As SaffronTree) As String
    Dim sVerb As String
    Dim sNoun As String
    Dim oNoun As clsNode
    Dim oConcept As clsNode
    Dim oVerbs As clsNode
    
    sVerb = LCase$(oInputTree(1).Text)
    sNoun = LCase$(oInputTree(2).Text)
    
    ' Is the noun available in this context?
    Set oNoun = goNet("context").Search(sNoun)
    
    If Not oNoun Is Nothing Then
        Set oConcept = goNet(sNoun)
        
        If Not oConcept Is Nothing Then
            ' Is this verb applicable?
            Set oVerbs = oConcept("verbs")
            If Not oVerbs Is Nothing Then
                If Not oVerbs(sVerb) Is Nothing Then
                    Stop
                End If
            Else
                PerformAction = ""
            End If
            PerformAction = "What is a(n) " & sNoun & "?"
        End If
    Else
        Set oConcept = goNet(sNoun)
        If Not oConcept Is Nothing Then
            PerformAction = "What " & sNoun & "?"
        Else
            PerformAction = msErrorMessage
        End If
    End If
End Function

