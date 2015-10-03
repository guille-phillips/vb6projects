VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Height          =   2415
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "Form1.frx":0000
      Top             =   240
      Width           =   4335
   End
   Begin MSCommLib.MSComm Com 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
      OutBufferSize   =   32767
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Com_OnComm()
    Debug.Print Com.Input
End Sub

Private Sub Form_Load()
    InitialiseComPort
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    Com.Output = Chr$(KeyAscii)
    
End Sub

Private Sub InitialiseComPort()
    Com.CommPort = 1
    Com.Settings = "9600,N,8,1"
    Com.InputLen = 0
    Com.OutBufferSize = 32767
    Com.PortOpen = True
End Sub
