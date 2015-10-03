VERSION 5.00
Begin VB.Form SaveMemory 
   Caption         =   "Save Memory"
   ClientHeight    =   1485
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   2940
   LinkTopic       =   "Form1"
   ScaleHeight     =   1485
   ScaleWidth      =   2940
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtEndAddress 
      Height          =   285
      Left            =   1560
      TabIndex        =   4
      Text            =   "7FFF"
      Top             =   600
      Width           =   615
   End
   Begin VB.TextBox txtStartAddress 
      Height          =   285
      Left            =   1560
      TabIndex        =   3
      Text            =   "0000"
      Top             =   240
      Width           =   615
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Default         =   -1  'True
      Height          =   375
      Left            =   1920
      TabIndex        =   2
      Top             =   960
      Width           =   855
   End
   Begin VB.Label lblEndAddress 
      Caption         =   "End Address (hex):"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   1455
   End
   Begin VB.Label lblStartAddress 
      Caption         =   "Start Address (hex):"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   1455
   End
End
Attribute VB_Name = "SaveMemory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public mlStartAddress As Long
Public mlEndAddress As Long

Public Enum ModeTypes
    mtSave
    mtLoad
End Enum

Private Sub cmdSave_Click()
    mlStartAddress = HexToDec(txtStartAddress.Text)
    mlEndAddress = HexToDec(txtEndAddress.Text)
    
    If mlStartAddress = -1 Then
        MsgBox "Start address not valid"
        txtStartAddress.SetFocus
        Exit Sub
    End If
    
    If mlEndAddress = -1 Then
        MsgBox "End address not valid"
        txtEndAddress.SetFocus
        Exit Sub
    End If
    
    Me.Hide
End Sub

Private Sub Form_Initialize()
    mlStartAddress = -1
    mlEndAddress = -1
End Sub

Public Property Let Mode(lMode As ModeTypes)
    Select Case lMode
        Case mtSave
            Me.Caption = "Save Memory"
            lblStartAddress = "Start Address (hex):"
            lblEndAddress.Visible = True
            txtEndAddress.Visible = True
            cmdSave.Caption = "Save"
        Case mtLoad
            Me.Caption = "Load Memory"
            lblStartAddress = "Load Address (hex):"
            lblEndAddress.Visible = False
            txtEndAddress.Visible = False
            cmdSave.Caption = "Load"
    End Select
End Property
