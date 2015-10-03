VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.MDIForm MDIRelationships 
   BackColor       =   &H8000000C&
   Caption         =   "Network"
   ClientHeight    =   9615
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   13245
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog comFile 
      Left            =   480
      Top             =   480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuNew 
         Caption         =   "&New..."
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open..."
      End
      Begin VB.Menu mnuClose 
         Caption         =   "&Close"
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Enabled         =   0   'False
      Begin VB.Menu mnuZoom 
         Caption         =   "&Zoom"
      End
   End
   Begin VB.Menu mnuWindow 
      Caption         =   "&Window"
      Begin VB.Menu mnuTileVertically 
         Caption         =   "Tile Vertically"
      End
      Begin VB.Menu mnuTileHorizontally 
         Caption         =   "Tile Horizontally"
      End
   End
End
Attribute VB_Name = "MDIRelationships"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub MDIForm_Load()
    DoDefinition
End Sub

Private Sub mnuClose_Click()
    Unload Me.ActiveForm
End Sub

Private Sub mnuExit_Click()
    Unload Me
End Sub

Private Sub mnuNew_Click()
    Dim oFSO As FileSystemObject
    Dim oDiagram As Paper
    Dim sExtension As String
    Dim iDot As Integer
    Dim sFile As String
    Dim oNewFile As TextStream
    
    comFile.Filter = "*.rel"
    comFile.ShowOpen

    Set oFSO = New FileSystemObject
    If oFSO.FileExists(comFile.FileName) Then
        If MsgBox("Would you like to overwrite this file?", vbYesNo) = vbNo Then
            Exit Sub
        End If
    End If
    iDot = InStrRev(comFile.FileTitle, ".")
    sFile = comFile.FileName
    If iDot > 0 Then
        sExtension = UCase$(Right$(comFile.FileTitle, Len(comFile.FileTitle) - iDot))
        If sExtension = "" Then
            sFile = comFile.FileName & ".rel"
        End If
    Else
        sFile = comFile.FileName & ".rel"
    End If
    Set oNewFile = oFSO.CreateTextFile(sFile, True, False)
    oNewFile.WriteLine "C:0|255|65280|65535|16711680|16711935|16776960|16777215|8421504|33023|16576|8388736|16744703|16777215|16777215|16777215|16777215|16777215|16777215|16777215|16777215"
    oNewFile.WriteLine "O:0|0"
    oNewFile.WriteLine "Z:1"
    oNewFile.Close
    
    Set oDiagram = New Paper
    oDiagram.DiagramRef.Initialise sFile
    oDiagram.Caption = "Relationships - " & comFile.FileTitle
    oDiagram.Show
End Sub

Private Sub mnuOpen_Click()
    Dim oDiagram As Paper
    
    comFile.Filter = "*.rel"
    comFile.ShowOpen
    If comFile.Flags And 1024 = 1024 Then
        Set oDiagram = New Paper
        oDiagram.DiagramRef.Initialise comFile.FileName
        oDiagram.Caption = "Relationships - " & comFile.FileTitle
        oDiagram.Show
    End If
End Sub

Private Sub mnuTileHorizontally_Click()
    Me.Arrange vbTileHorizontal
End Sub

Private Sub mnuTileVertically_Click()
    Me.Arrange vbTileVertical
End Sub
