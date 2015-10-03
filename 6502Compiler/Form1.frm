VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Compiler"
   ClientHeight    =   2175
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   2310
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2175
   ScaleWidth      =   2310
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCompileAssembly 
      Caption         =   "Compile Assembly"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   1815
   End
   Begin VB.CommandButton cmdDisassembleROM 
      Caption         =   "Disassemble ROM"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   1560
      Width           =   1815
   End
   Begin VB.CommandButton cmdDisassemble 
      Caption         =   "Disassemble Snapshot"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   840
      Width           =   1815
   End
   Begin MSComDlg.CommonDialog comFile 
      Left            =   1320
      Top             =   360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Function GetFilePath(ByVal sFileName As String) As String
    Dim lSlash As Long
    
    lSlash = InStrRev(sFileName, "\")
    GetFilePath = Left$(sFileName, lSlash) & "*.*"
End Function

Private Sub cmdCompileAssembly_Click()
    Dim sDisassembled As String
    Dim sPath As String
    
    On Error GoTo exit_cmdCA
    
    sPath = GetSetting("R6502Compiler", "Paths", "Assembly", App.path & "\Assembly\*.*")
    
    comFile.CancelError = True
    comFile.FileName = sPath
    comFile.Filter = "*.txt"
    comFile.ShowOpen

    SaveSetting "R6502Compiler", "Paths", "Assembly", GetFilePath(comFile.FileName)

    Compiler6502.LoadAssembly (comFile.FileName)
    MsgBox "Finished"
    
exit_cmdCA:
End Sub

Private Sub cmdDisassemble_Click()
    Dim sDisassembled As String
    
    On Error GoTo exit_cmdD
        
    comFile.CancelError = True
    comFile.FileName = App.path & "\Snapshots\*.*"
    comFile.Filter = "*.uef"
    comFile.ShowOpen
    
    Erase gyMem
    
    Snapshot.LoadSnapshot comFile.FileName
    
    sDisassembled = Disassembler.Disassemble(0, &H7FFF&)
    
    If Dir(App.path & "\Disassembly\disassembled.txt") <> "" Then
        Kill App.path & "\Disassembly\disassembled.txt"
    End If
    Open App.path & "\Disassembly\disassembled.txt" For Binary As #1
    Put #1, , sDisassembled
    Close #1
    MsgBox "Finished"
    
exit_cmdD:
End Sub

Private Sub cmdDisassembleROM_Click()
    Dim sDisassembled As String
    
    On Error GoTo exit_cmdR
        
    comFile.CancelError = True
    comFile.FileName = App.path & "\Snapshots\*.*"
    comFile.Filter = "*.uef"
    comFile.ShowOpen

    Roms.LoadRom comFile.FileName
    
    sDisassembled = Disassembler.Disassemble(&H8000&, &HBFFF&)
    
    If Dir(App.path & "\Disassembly\disassembled.txt") <> "" Then
        Kill App.path & "\Disassembly\disassembled.txt"
    End If
    Open App.path & "\Disassembly\disassembled.txt" For Binary As #1
    Put #1, , sDisassembled
    Close #1
    MsgBox "Finished"
    
exit_cmdR:
End Sub
