VERSION 5.00
Begin VB.Form Speed 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Speed"
   ClientHeight    =   885
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   885
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrSpeed 
      Interval        =   500
      Left            =   3000
      Top             =   0
   End
   Begin VB.HScrollBar scrSpeed 
      Height          =   255
      LargeChange     =   100
      Left            =   240
      Max             =   30000
      Min             =   100
      SmallChange     =   100
      TabIndex        =   0
      Top             =   480
      Value           =   10000
      Width           =   4215
   End
   Begin VB.Label lblSpeed 
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   120
      Width           =   2175
   End
End
Attribute VB_Name = "Speed"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mbUserChangedSpeed As Boolean

Private Sub Form_Activate()
    scrSpeed.Value = Throttle.DefinedSpeedControl
    mbUserChangedSpeed = False
    scrSpeed_Change
End Sub

Private Sub scrSpeed_Change()
    lblSpeed.Caption = Format$(scrSpeed.Value / 100, "0.00") & " %"
    Throttle.SpeedControlStep = 1024
    Throttle.SpeedControl = scrSpeed.Value
    If mbUserChangedSpeed Then
        Throttle.DefinedSpeedControl = scrSpeed.Value
    End If
    mbUserChangedSpeed = True
End Sub

Private Sub scrSpeed_Scroll()
    scrSpeed_Change
End Sub

Private Sub tmrSpeed_Timer()
    mbUserChangedSpeed = False
    scrSpeed.Value = Throttle.SpeedControl
End Sub
