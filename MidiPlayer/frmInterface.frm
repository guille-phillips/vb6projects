VERSION 5.00
Begin VB.Form frmInterface 
   Caption         =   "Midi Player"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   206
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrClock 
      Interval        =   1
      Left            =   240
      Top             =   240
   End
End
Attribute VB_Name = "frmInterface"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private lTick As Long

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim lError As Long
    
    Debug.Print midiOutClose(hMidiOutput)
    Debug.Print midiInStop(hMidiInput)
    Debug.Print midiInClose(hMidiInput)
End Sub

