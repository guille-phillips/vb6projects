VERSION 5.00
Begin VB.Form frmEditAlarm 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Edit Alarm"
   ClientHeight    =   3750
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   3270
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3750
   ScaleWidth      =   3270
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtAlarm 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   5
      Left            =   1320
      TabIndex        =   6
      Top             =   2400
      Width           =   375
   End
   Begin VB.ComboBox cboRecurPeriod 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "frmEditAlarm.frx":0000
      Left            =   1800
      List            =   "frmEditAlarm.frx":0031
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   2400
      Width           =   1335
   End
   Begin VB.CheckBox chkThisComputerOnly 
      Alignment       =   1  'Right Justify
      Caption         =   "This Computer Only"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1560
      TabIndex        =   8
      Top             =   2880
      Width           =   225
   End
   Begin VB.TextBox txtAlarm 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   4
      Left            =   1320
      Locked          =   -1  'True
      ScrollBars      =   1  'Horizontal
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1320
      Width           =   1815
   End
   Begin VB.CommandButton cmdExpire 
      Cancel          =   -1  'True
      Caption         =   "Cancel Alarm"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   960
      TabIndex        =   9
      Top             =   3240
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtAlarm 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   1320
      TabIndex        =   3
      Top             =   1680
      Width           =   1815
   End
   Begin VB.ComboBox cboRemindPeriod 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "frmEditAlarm.frx":00BC
      Left            =   1800
      List            =   "frmEditAlarm.frx":00ED
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   2040
      Width           =   1335
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "Set Alarm"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      TabIndex        =   10
      Top             =   3240
      Width           =   975
   End
   Begin VB.TextBox txtAlarm 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Index           =   3
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   360
      Width           =   3015
   End
   Begin VB.TextBox txtAlarm 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   2
      Left            =   1320
      TabIndex        =   4
      Top             =   2040
      Width           =   375
   End
   Begin VB.TextBox txtAlarm 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   1320
      TabIndex        =   1
      Top             =   960
      Width           =   1815
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Recur Every"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   2400
      Width           =   1695
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "This Computer Only"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   2880
      Width           =   1695
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Time Remaining"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   1320
      Width           =   1695
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Reminder From"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   1680
      Width           =   1695
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Event Description"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Remind Every"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   2040
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Event Time"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   960
      Width           =   1695
   End
End
Attribute VB_Name = "frmEditAlarm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const SWP_NOMOVE = 2
Private Const SWP_NOSIZE = 1
Private Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2

Private Declare Function SetWindowPos Lib "user32" _
      (ByVal hwnd As Long, _
      ByVal hWndInsertAfter As Long, _
      ByVal X As Long, _
      ByVal Y As Long, _
      ByVal cx As Long, _
      ByVal cy As Long, _
      ByVal wFlags As Long) As Long
      
Public Callback As frmAlarms
Private moAlarm As clsAlarm
Public Index As Long

Private Enum Fields
    EventTime
    ReminderFrom
    RemindEvery
    EventDescription
    TimeRemaining
    RecurEvery
End Enum

Public Sub ShowUrgency()
    Dim nPercentage As Single
    Dim lDiff As Long
    Dim lDiff2 As Long
    
    If Not moAlarm.Expired Then
        If moAlarm.ReminderFrom <> 0 Then
            If Now < moAlarm.ReminderFrom Then
                BackColor = RGB(217, 255, 210) ' green
            Else
                BackColor = RGB(255, 255, 179) ' yellow
            End If
        Else
            BackColor = RGB(217, 255, 210) ' green
        End If
        
        If Now >= moAlarm.EventTime Then
            BackColor = RGB(255, 182, 182)
        End If
    End If
End Sub

Private Sub Form_Initialize()
    Index = -1
End Sub

Private Sub Form_Load()
    SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS
End Sub



Private Sub txtAlarm_Change(Index As Integer)
    txtAlarm(Index).ForeColor = vbBlack
End Sub

Private Sub txtAlarm_LostFocus(Index As Integer)
    Dim dDate As Date
    
    If Index <= ReminderFrom Then
        dDate = ParseDateExpression(txtAlarm(Index).Text)
        If dDate <> 0 Then
            txtAlarm(Index).Text = Format$(dDate, "hh:nn ddd dd/mm/yyyy")
            txtAlarm(Index).ForeColor = vbBlack
        Else
            txtAlarm(Index).Text = Trim$(txtAlarm(Index).Text)
            txtAlarm(Index).ForeColor = vbRed
        End If
    End If
End Sub

Public Sub ShowAlarm(oAlarm As clsAlarm)
    Set moAlarm = oAlarm
    txtAlarm(EventTime).Text = Format$(oAlarm.EventTime, "hh:nn ddd dd/mm/yyyy")
    If moAlarm.ReminderFrom <> 0 Then
        txtAlarm(ReminderFrom).Text = Format$(oAlarm.ReminderFrom, "hh:nn ddd dd/mm/yyyy")
    End If
    If moAlarm.RemindEvery <> 0 Then
        txtAlarm(RemindEvery).Text = oAlarm.RemindEvery
    End If
    cboRemindPeriod.ListIndex = oAlarm.RemindEveryType
    
    If moAlarm.RecurEvery <> 0 Then
        txtAlarm(RecurEvery).Text = oAlarm.RecurEvery
    End If
    cboRecurPeriod.ListIndex = oAlarm.RecurEveryType
    
    txtAlarm(EventDescription).Text = oAlarm.EventDescription
    chkThisComputerOnly.Value = IIf(oAlarm.OwnerID <> "", vbChecked, vbUnchecked)
    ShowUrgency
End Sub

Private Function AdjustTime(ByVal dTime As Date) As Date
    Dim dNow As Date
    
    If dTime < 1 Then
        dNow = Now
        If (dNow - Int(dNow)) <= dTime Then
            AdjustTime = Int(dNow) + dTime
        Else
            AdjustTime = Int(dNow) + dTime + 1
        End If
    Else
        AdjustTime = dTime
    End If
End Function

Private Sub cmdExpire_Click()
    moAlarm.Expired = True
    moAlarm.OnHold = False
    If Index <> -1 Then
        Callback.moOpenAlarms.Remove "k" & Index
    End If
    Callback.SaveAlarms
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim bAddAlarm As Boolean
    Dim dDate As Date
    Dim dReminderDate As Date
    
    If ValidateFields Then
        If moAlarm Is Nothing Then
            Set moAlarm = New clsAlarm
            bAddAlarm = True
        End If
        
        dDate = ParseDateExpression(txtAlarm(EventTime).Text)
        moAlarm.EventTime = dDate
        If txtAlarm(ReminderFrom).Text <> "" Then
            dReminderDate = ParseDateExpression(txtAlarm(ReminderFrom).Text)
            moAlarm.ReminderFrom = dReminderDate
            If moAlarm.ReminderFrom > Now Then
                moAlarm.MostRecentAlarmTime = moAlarm.ReminderFrom
            End If
        Else
            moAlarm.ReminderFrom = 0
            If moAlarm.EventTime > Now Then
                moAlarm.MostRecentAlarmTime = moAlarm.EventTime
            End If
        End If
        moAlarm.RemindEvery = Val(txtAlarm(RemindEvery).Text)
        moAlarm.RemindEveryType = cboRemindPeriod.ListIndex
        moAlarm.RecurEvery = Val(txtAlarm(RecurEvery).Text)
        moAlarm.RecurEveryType = cboRecurPeriod.ListIndex
        moAlarm.EventDescription = txtAlarm(EventDescription).Text
        moAlarm.Expired = False
        moAlarm.OnHold = False
        moAlarm.Duplicated = False
        If chkThisComputerOnly.Value = vbChecked Then
            moAlarm.OwnerID = gsComputerIdentifier
        Else
            moAlarm.OwnerID = ""
        End If

        If bAddAlarm Then
            Callback.AddAlarm moAlarm
        End If
        Callback.SaveAlarms
        Unload Me
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error Resume Next
    If Index <> -1 Then
        Callback.moOpenAlarms.Remove "k" & Index
    End If
    moAlarm.OnHold = False
End Sub

Private Function ValidateFields() As Boolean
    Dim dDate As Date
    Dim dReminderDate As Date
    
    dDate = ParseDateExpression(txtAlarm(EventTime).Text)
        
    If dDate = 0 Then
        MsgBox "Invalid Event Time"
        Exit Function
    End If

    If txtAlarm(ReminderFrom).Text <> "" Then
        dReminderDate = ParseDateExpression(txtAlarm(ReminderFrom).Text)
        If dReminderDate = 0 Then
            MsgBox "Invalid Reminder From Time"
            Exit Function
        End If
        If dReminderDate >= dDate Then
            MsgBox "Reminder Date must be before Event Date"
            Exit Function
        End If
    End If
    
    If txtAlarm(RemindEvery).Text <> "" Then
        If cboRemindPeriod.ListIndex = -1 Then
            MsgBox "Choose a remind period"
            Exit Function
        End If
    End If
    
    If txtAlarm(RecurEvery).Text <> "" Then
        If cboRecurPeriod.ListIndex = -1 Then
            MsgBox "Choose a recur period"
            Exit Function
        End If
    End If
    ValidateFields = True
End Function
