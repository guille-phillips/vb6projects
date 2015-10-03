VERSION 5.00
Begin VB.Form NewString 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00800080&
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtString 
      Height          =   285
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4455
   End
End
Attribute VB_Name = "NewString"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.Hide
    End If
End Sub

Private Sub Form_Resize()
    txtString.Width = Me.ScaleWidth
    txtString.Height = Me.ScaleHeight
End Sub
