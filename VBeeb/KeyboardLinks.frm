VERSION 5.00
Begin VB.Form KeyboardLinks 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Keyboard Links"
   ClientHeight    =   3060
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   1635
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3060
   ScaleWidth      =   1635
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkBit 
      Caption         =   "Bit 0"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   7
      Top             =   2640
      Width           =   975
   End
   Begin VB.CheckBox chkBit 
      Caption         =   "Bit 1"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   6
      Top             =   2280
      Width           =   975
   End
   Begin VB.CheckBox chkBit 
      Caption         =   "Bit 2"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   5
      Top             =   1920
      Width           =   975
   End
   Begin VB.CheckBox chkBit 
      Caption         =   "Bit 3"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   4
      Top             =   1560
      Width           =   975
   End
   Begin VB.CheckBox chkBit 
      Caption         =   "Bit 4"
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   975
   End
   Begin VB.CheckBox chkBit 
      Caption         =   "Bit 5"
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   975
   End
   Begin VB.CheckBox chkBit 
      Caption         =   "Bit 6"
      Height          =   255
      Index           =   6
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   975
   End
   Begin VB.CheckBox chkBit 
      Caption         =   "Bit 7"
      Height          =   255
      Index           =   7
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "KeyboardLinks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mbNoUpdate As Boolean

Private Sub chkBit_Click(Index As Integer)
    If Not mbNoUpdate Then
        Keyboard.KeyboardLinks = Keyboard.KeyboardLinks Xor 2 ^ Index
        SaveSetting "VBeeb", "KeyboardLinks", "KeyboardLinks", Keyboard.KeyboardLinks
        Keyboard.InitialiseKeyboardLinks
    End If
End Sub

Private Sub Form_Load()
    Dim lBit As Long
    Dim lMask As Long
    
    mbNoUpdate = True
    lMask = 1
    For lBit = 0 To 7
        If (Keyboard.KeyboardLinks And lMask) = 0 Then
            chkBit(lBit).Value = vbUnchecked
        Else
            chkBit(lBit).Value = vbChecked
        End If
        lMask = lMask * 2
    Next
    mbNoUpdate = False
End Sub
