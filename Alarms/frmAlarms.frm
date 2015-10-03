VERSION 5.00
Begin VB.Form frmAlarms 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Alarms"
   ClientHeight    =   2700
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   7695
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAlarms.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2700
   ScaleWidth      =   7695
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOld 
      Caption         =   "Old"
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   2280
      Width           =   975
   End
   Begin VB.CommandButton cmdExpire 
      Caption         =   "Expire"
      Height          =   375
      Left            =   3360
      TabIndex        =   7
      Top             =   2280
      Width           =   975
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "Refresh"
      Height          =   375
      Left            =   1200
      TabIndex        =   6
      Top             =   2280
      Width           =   975
   End
   Begin VB.PictureBox pctList 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3975
      Left            =   7680
      ScaleHeight     =   3915
      ScaleWidth      =   5475
      TabIndex        =   5
      Top             =   0
      Visible         =   0   'False
      Width           =   5535
   End
   Begin VB.CommandButton cmdHide 
      Caption         =   "Hide"
      Height          =   375
      Left            =   2280
      TabIndex        =   4
      Top             =   2280
      Width           =   975
   End
   Begin VB.Timer tmrAlarm 
      Interval        =   1000
      Left            =   120
      Top             =   2400
   End
   Begin VB.CommandButton cmdChange 
      Caption         =   "Change"
      Default         =   -1  'True
      Height          =   375
      Left            =   6600
      TabIndex        =   3
      Top             =   2280
      Width           =   975
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "Remove"
      Height          =   375
      Left            =   5520
      TabIndex        =   2
      Top             =   2280
      Width           =   975
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Height          =   375
      Left            =   4440
      TabIndex        =   1
      Top             =   2280
      Width           =   975
   End
   Begin VB.ListBox lstAlarms 
      Height          =   2205
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7695
   End
   Begin VB.Menu mnuOptions 
      Caption         =   "Options"
      Visible         =   0   'False
      Begin VB.Menu mnuClose 
         Caption         =   "Close"
      End
   End
End
Attribute VB_Name = "frmAlarms"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const BM_SETSTATE = &HF3

Private Declare Function SendMessageBynum Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Private Declare Function Shell_NotifyIcon Lib "shell32.dll" (ByVal dwMessage As Long, lpData As NOTIFYICONDATA) As Long

Private Const NIM_ADD = &H0
Private Const NIM_DELETE = &H2
Private Const NIM_MODIFY = &H1

Private Const NIF_ICON = &H2
Private Const NIF_MESSAGE = &H1
Private Const NIF_TIP = &H4

Private Const WM_LBUTTONDBLCLK = &H203
Private Const WM_LBUTTONDOWN = &H201
Private Const WM_LBUTTONUP = &H202
Private Const WM_MOUSEMOVE = &H200
Private Const WM_RBUTTONDBLCLK = &H206
Private Const WM_RBUTTONDOWN = &H204
Private Const WM_RBUTTONUP = &H205

Private mlTopIndex As Long

Private Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    uID As Long
    uFlags As Long
    uCallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type

Private IconData As NOTIFYICONDATA

Private moAlarms As New clsNode
Private moFiles As New clsNode
Private moAlarmList As New clsNode

Private msListHeight As Single
Private msButtonTop As Single
Private msHideLeft As Single
Private msAddLeft As Single
Private msRemoveLeft As Single
Private msChangeLeft As Single
Private msRefreshLeft As Single
Private msExpireLeft As Single
Private mbViewRemoved As Boolean
Private msOldLeft As Single

Public moOpenAlarms As New Collection

Private Sub cmdExpire_Click()
    Dim lPhysicalIndex As Long
    
    If lstAlarms.ListIndex <> -1 Then
        lPhysicalIndex = moAlarmList.ItemPhysical(lstAlarms.ListIndex).Value
        moAlarms.ItemPhysical(lPhysicalIndex).Value.Expired = True
    End If
End Sub

Private Sub cmdHide_Click()
    Me.Hide
End Sub

Private Sub cmdOld_Click()
    lstAlarms.ListIndex = -1
    mbViewRemoved = Not mbViewRemoved
    
    If mbViewRemoved Then
        cmdRemove.Caption = "Delete Removed"
    Else
        cmdRemove.Caption = "Remove"
    End If
    
    'SendMessageBynum cmdOld.hwnd, BM_SETSTATE, CLng(mbViewRemoved), 0
End Sub

Private Sub cmdRefresh_Click()
    LoadAlarms
End Sub

Private Sub Form_Load()
    msListHeight = Me.Height - lstAlarms.Height
    msButtonTop = Me.Height - cmdAdd.Top
    msHideLeft = Me.Width - cmdHide.Left
    msExpireLeft = Me.Width - cmdExpire.Left
    msAddLeft = Me.Width - cmdAdd.Left
    msRemoveLeft = Me.Width - cmdRemove.Left
    msChangeLeft = Me.Width - cmdChange.Left
    msRefreshLeft = Me.Width - cmdRefresh.Left
    msOldLeft = Me.Width - cmdOld.Left
    
    Me.Width = GetSetting("Alarms", "Dimensions", "Width", Me.Width)
    Me.Height = GetSetting("Alarms", "Dimensions", "Height", Me.Height)
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then
        Me.Show
    ElseIf Button = vbRightButton Then
        Me.PopupMenu mnuOptions
    End If
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    lstAlarms.Width = Me.ScaleWidth
    lstAlarms.Height = Me.Height - msListHeight
    cmdHide.Top = Me.Height - msButtonTop
    cmdExpire.Top = Me.Height - msButtonTop
    cmdAdd.Top = Me.Height - msButtonTop
    cmdRemove.Top = Me.Height - msButtonTop
    cmdChange.Top = Me.Height - msButtonTop
    cmdRefresh.Top = Me.Height - msButtonTop
    cmdOld.Top = Me.Height - msButtonTop
    cmdHide.Left = Me.Width - msHideLeft
    cmdExpire.Left = Me.Width - msExpireLeft
    cmdAdd.Left = Me.Width - msAddLeft
    cmdRemove.Left = Me.Width - msRemoveLeft
    cmdChange.Left = Me.Width - msChangeLeft
    cmdRefresh.Left = Me.Width - msRefreshLeft
    cmdOld.Left = Me.Width - msOldLeft
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveSetting "Alarms", "Dimensions", "Width", Me.Width
    SaveSetting "Alarms", "Dimensions", "Height", Me.Height
End Sub

Private Sub Form_Initialize()
    Me.Hide
    gsComputerIdentifier = GetSetting("alarms", "identifier", "computer", NewGUID)
    SaveSetting "alarms", "identifier", "computer", gsComputerIdentifier
    
    LoadRules
    LoadAlarms
    
    With IconData
        .cbSize = Len(IconData)
        .hIcon = Me.Icon
        .hwnd = Me.hwnd
        .szTip = "Alarms" & Chr(0)
        .uCallbackMessage = WM_MOUSEMOVE
        .uFlags = NIF_ICON Or NIF_MESSAGE  'Or NIF_TIP Or
        .uID = vbNull
    End With
    Shell_NotifyIcon NIM_ADD, IconData
End Sub

Private Sub Form_Terminate()
    SaveAlarms
    Shell_NotifyIcon NIM_DELETE, IconData
End Sub

Private Sub cmdAdd_Click()
    Dim oEditAlarm As New frmEditAlarm
    
    Set oEditAlarm.Callback = Me
    oEditAlarm.Show
End Sub

Private Sub cmdChange_Click()
    Dim oEditAlarm As frmEditAlarm
    Dim oAlarm As clsAlarm
    Dim lLogicalIndex As Long
    Dim lPhysicalIndex As Long
    
    If lstAlarms.ListIndex <> -1 Then
        lPhysicalIndex = moAlarmList.ItemPhysical(lstAlarms.ListIndex).Value
        Set oAlarm = moAlarms.ItemPhysical(lPhysicalIndex).Value
        lLogicalIndex = moAlarms.ItemPhysical(lPhysicalIndex).LogicalKey
        
        If Not oAlarm.OnHold Then
            oAlarm.OnHold = True
            Set oEditAlarm = New frmEditAlarm
            Set oEditAlarm.Callback = Me
            oEditAlarm.Index = lLogicalIndex
            oEditAlarm.ShowAlarm moAlarms.ItemPhysical(lPhysicalIndex).Value, True
            oEditAlarm.Show
            moOpenAlarms.Add oEditAlarm, "k" & lLogicalIndex
        End If
    End If
End Sub

Private Sub cmdRemove_Click()
    Dim lPhysicalIndex As Long
    Dim lListIndex As Long
    
    If lstAlarms.ListIndex <> -1 Then
        If Not mbViewRemoved Then
            lPhysicalIndex = moAlarmList.ItemPhysical(lstAlarms.ListIndex).Value
            moAlarms.ItemPhysical(lPhysicalIndex).Value.Deleted = True
        Else
            lPhysicalIndex = moAlarmList.ItemPhysical(lstAlarms.ListIndex).Value
            moAlarms.RemovePhysical (lPhysicalIndex)
        End If
    Else
        If mbViewRemoved Then
            lPhysicalIndex = 0
            While lPhysicalIndex < moAlarms.Count
                If moAlarms.ItemPhysical(lPhysicalIndex).Value.Deleted Then
                    moAlarms.RemovePhysical (lPhysicalIndex)
                    lPhysicalIndex = lPhysicalIndex - 1
                End If
                lPhysicalIndex = lPhysicalIndex + 1
            Wend
        Else
            lPhysicalIndex = 0
            While lPhysicalIndex < moAlarms.Count
                If moAlarms.ItemPhysical(lPhysicalIndex).Value.Expired Then
                    moAlarms.ItemPhysical(lPhysicalIndex).Value.Deleted = True
                End If
                lPhysicalIndex = lPhysicalIndex + 1
            Wend
        End If
    End If
    SaveAlarms
End Sub


Private Sub lstAlarms_Click()
    cmdRemove.Caption = "Remove"
End Sub

Private Sub lstAlarms_DblClick()
    cmdChange_Click
End Sub

Private Sub lstAlarms_Scroll()
    mlTopIndex = lstAlarms.TopIndex
End Sub

Private Sub mnuClose_Click()
    Unload Me
End Sub

Private Sub tmrAlarm_Timer()
    Dim oAlarm As clsAlarm
    Dim oNewAlarm As clsAlarm
    Dim oAlarmListItem As clsNode
        
    Dim lListIndex As Long
    Dim sLineDescription As String
    Dim sDescription As String
    Dim sDescriptionShort As String
    Dim oMessage As frmEditAlarm
    Dim sTimeLeft As String
    Dim dNow As Date
    
    Dim oAlarmsSorted As New Collection
    Dim dEarliestEventTime As Date
    Dim dCutoff As Date
    Dim oThisAlarm As clsAlarm
    Dim bSorted As Boolean
    
    Dim lIndex As Long
    Dim lLogicalIndex As Long
    
    While Not bSorted
        bSorted = True
        For lIndex = 0 To moAlarms.Count - 2
            If moAlarms.ItemPhysical(lIndex).Value.EventTime > moAlarms.ItemPhysical(lIndex + 1).Value.EventTime Then
                moAlarms.Move lIndex, lIndex + 1
                bSorted = False
            End If
        Next
    Wend

    lListIndex = lstAlarms.ListIndex
    
    lstAlarms.Clear
    
    dNow = Now
    pctList.Cls
    moAlarmList.RemoveAll

    For lIndex = 0 To moAlarms.Count - 1
        Set oAlarm = moAlarms.ItemPhysical(lIndex).Value
        lLogicalIndex = moAlarms.ItemPhysical(lIndex).LogicalKey
        
        If Not (oAlarm.Deleted) Or mbViewRemoved Then
            Set oAlarmListItem = moAlarmList.AddNew
            oAlarmListItem.Value = lIndex
                
            sLineDescription = IIf(oAlarm.Deleted, "REMOVED: ", IIf(oAlarm.Expired, "EXPIRED: ", ""))
            sDescriptionShort = sLineDescription
            sLineDescription = sLineDescription & oAlarm.EventDescription & " "
            sLineDescription = sLineDescription & Format$(oAlarm.EventTime, "ddd dd mmm yyyy")
            
            sTimeLeft = oAlarm.TimeLeft(Now)
            If sTimeLeft <> "" Then
                sLineDescription = sLineDescription & " - " & sTimeLeft
                sDescriptionShort = sDescriptionShort & sTimeLeft
            End If
            lstAlarms.AddItem sLineDescription
            DisplayRow lIndex, "status", oAlarm.EventDescription, oAlarm.EventTime
            On Error Resume Next
            moOpenAlarms.Item("k" & lLogicalIndex).txtAlarm(4).Text = sDescriptionShort
            On Error GoTo 0
        End If
    Next

    lstAlarms.ListIndex = lListIndex
    For lIndex = 0 To moAlarms.Count - 1
        Set oAlarm = moAlarms.ItemPhysical(lIndex).Value
        lLogicalIndex = moAlarms.ItemPhysical(lIndex).LogicalKey
        If Not oAlarm.Deleted Then
            sDescription = oAlarm.CheckAlarm(Now)
            If sDescription <> "" Then
                oAlarm.OnHold = True
                Set oMessage = New frmEditAlarm
                moOpenAlarms.Add oMessage, "k" & lLogicalIndex
                oMessage.Index = lLogicalIndex
                oMessage.cmdExpire.Visible = True
                oMessage.cmdOK.Caption = "Snooze"
                oMessage.Caption = "Alarm Reminder"
                oMessage.ShowAlarm oAlarm
                Set oMessage.Callback = Me
                oMessage.Show
            End If
            If oAlarm.Expired And oAlarm.RecurEvery <> 0 And Not oAlarm.Duplicated Then
                oAlarm.Duplicated = True
                Set oNewAlarm = oAlarm.Duplicate
                Set moAlarms.AddNew.Value = oNewAlarm
            End If
        End If
    Next
    
    lstAlarms.TopIndex = mlTopIndex
End Sub

Private Sub DisplayRow(ByVal lIndex As Long, ByVal sStatus As String, ByVal sDescription, ByVal dEventTime As Date)
    Dim nVerticalPos As Single
    Const nOffsetTop As Single = 0
    Const nOffsetLeft As Single = 0
    
    nVerticalPos = lIndex * pctList.TextHeight("X") + nOffsetTop
    pctList.CurrentY = nVerticalPos
    pctList.CurrentX = nOffsetLeft
    pctList.Print sDescription
End Sub

Public Sub AddAlarm(oAlarm As clsAlarm)
    Dim oNewNode As clsNode
    
    Set oNewNode = moAlarms.AddNew
    Set oNewNode.Value = oAlarm
    SaveAlarms
End Sub

Public Sub ExpireAlarm(oAlarm As clsAlarm)
    oAlarm.MostRecentAlarmTime = DateAdd("y", 100, oAlarm.MostRecentAlarmTime)
    SaveAlarms
End Sub

Public Sub LoadAlarms()
    Dim oFSO As New FileSystemObject
    Dim oTS As TextStream
    Dim vAlarm As Variant
    Dim oAlarm As clsAlarm
    Dim oNewNode As clsNode
    Dim sGUID As String
    Dim dLastUpdatedDate As Date
    Dim sFilename As String
    Dim sOwnerId As String
    Dim lRecurIndex As Long
    
    Dim lLetter As Long
    Dim lIndex As Long
    
    Set moFiles = Nothing
    
    For lLetter = Asc("A") To Asc("Z")
        sFilename = Chr$(lLetter) & ":\alarms.txt"
        If oFSO.FileExists(sFilename) Then
            Set oNewNode = moFiles.AddNew
            oNewNode.Value = sFilename
        End If
    Next
    sFilename = App.Path & "\alarms.txt"
    If oFSO.FileExists(sFilename) Then
        Set oNewNode = moFiles.AddNew
        oNewNode.Value = sFilename
    End If

    For lIndex = 0 To moFiles.Count - 1
        sFilename = moFiles.ItemPhysical(lIndex).Value
        Set oTS = oFSO.OpenTextFile(sFilename, ForReading)
        
        Do While Not oTS.AtEndOfStream
            vAlarm = Split(oTS.ReadLine, "|")
            If UBound(vAlarm) < 6 Then
                Exit Do
            End If
            sOwnerId = vAlarm(12)
            If sOwnerId = "" Or sOwnerId = gsComputerIdentifier Then
                dLastUpdatedDate = CDate(vAlarm(1))
                sGUID = vAlarm(0)
                lRecurIndex = vAlarm(14)
                
                Set oAlarm = FindAlarm(sGUID, lRecurIndex)
                
                If oAlarm Is Nothing Then
                    Set oAlarm = New clsAlarm
                    If sGUID <> "" Then
                        oAlarm.GUID = sGUID
                    End If
                    oAlarm.LastUpdatedDate = dLastUpdatedDate
                    oAlarm.EventTime = CDate(vAlarm(2))
                    oAlarm.ReminderFrom = CDate(vAlarm(3))
                    oAlarm.RemindEvery = vAlarm(4)
                    oAlarm.RemindEveryType = vAlarm(5)
                    oAlarm.RecurEvery = vAlarm(6)
                    oAlarm.RecurEveryType = vAlarm(7)
                    oAlarm.EventDescription = vAlarm(8)
                    oAlarm.MostRecentAlarmTime = CDate(vAlarm(9))
                    oAlarm.Expired = vAlarm(10)
                    oAlarm.Duplicated = vAlarm(11)
                    oAlarm.OwnerID = vAlarm(12)
                    oAlarm.Deleted = vAlarm(13)
                    oAlarm.RecurIndex = vAlarm(14)
                    Set oNewNode = moAlarms.AddNew
                    Set oNewNode.Value = oAlarm
                Else
                    If dLastUpdatedDate > oAlarm.LastUpdatedDate Then
                        oAlarm.LastUpdatedDate = dLastUpdatedDate
                        oAlarm.EventTime = CDate(vAlarm(2))
                        oAlarm.ReminderFrom = CDate(vAlarm(3))
                        oAlarm.RemindEvery = vAlarm(4)
                        oAlarm.RemindEveryType = vAlarm(5)
                        oAlarm.RecurEvery = vAlarm(6)
                        oAlarm.RecurEveryType = vAlarm(7)
                        oAlarm.EventDescription = vAlarm(8)
                        oAlarm.MostRecentAlarmTime = CDate(vAlarm(9))
                        oAlarm.Expired = vAlarm(10)
                        oAlarm.Duplicated = vAlarm(11)
                        oAlarm.OwnerID = vAlarm(12)
                        oAlarm.Deleted = vAlarm(13)
                        oAlarm.RecurIndex = vAlarm(14)
                    End If
                End If
            End If
        Loop
    Next
End Sub

Private Function FindAlarm(sGUID As String, Optional ByVal lRecurIndex As Long) As clsAlarm
    Dim oNode As clsNode
    Dim oAlarm As clsAlarm
    Dim lIndex As Long
    
    For lIndex = 0 To moAlarms.Count - 1
        Set oNode = moAlarms.ItemPhysical(lIndex)
        Set oAlarm = oNode.Value
        
        If oAlarm.GUID = sGUID And oAlarm.RecurIndex = lRecurIndex Then
            Set FindAlarm = oAlarm
        End If
    Next
End Function

Public Sub SaveAlarms()
    Dim vAlarm As Variant
    Dim oAlarm As clsAlarm
    Dim oTS As TextStream
    Dim oFSO As New FileSystemObject
    Dim lIndex As Long
    Dim lFileIndex As Long
    Dim sFilename As String
    
    For lFileIndex = 0 To moFiles.Count - 1
        sFilename = moFiles.ItemPhysical(lFileIndex).Value
        If oFSO.FileExists(sFilename) Then
            Set oTS = oFSO.CreateTextFile(sFilename, True)
            
            For lIndex = 0 To moAlarms.Count - 1
                Set oAlarm = moAlarms.ItemPhysical(lIndex).Value
                vAlarm = Array(oAlarm.GUID, Format$(Now, "hh:nn:ss dd/mm/yyyy"), Format$(oAlarm.EventTime, "hh:nn dd/mm/yyyy"), Format$(oAlarm.ReminderFrom, "hh:nn dd/mm/yyyy"), oAlarm.RemindEvery, oAlarm.RemindEveryType, oAlarm.RecurEvery, oAlarm.RecurEveryType, oAlarm.EventDescription, Format$(oAlarm.MostRecentAlarmTime, "hh:nn dd/mm/yyyy"), oAlarm.Expired, oAlarm.Duplicated, oAlarm.OwnerID, oAlarm.Deleted, oAlarm.RecurIndex)
                oTS.WriteLine Join(vAlarm, "|")
            Next
            oTS.Close
        End If
    Next
End Sub

Private Sub LoadRules()
    Dim sDefinition As String
    
    sDefinition = Space$(FileLen(App.Path & "\date.saf"))
    Open App.Path & "\date.saf" For Binary As #1
    Get #1, , sDefinition
    Close #1

    If Not CreateRules(sDefinition) Then
        Debug.Print ErrorString
        End
    End If
    
'    Dim oTree As SaffronTree
'
'    Set oTree = New SaffronTree
'    SaffronStream.Text = "  16:50   "
'    If Rules("full_date_string").Parse(oTree) Then
'        Stop
'    Else
'        Stop
'    End If

    Set goDateParser = Rules("full_date_string")
'    Debug.Print ParseDateExpression("thu 3 oct 2009")
End Sub

