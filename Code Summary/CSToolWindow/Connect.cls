VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Connect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "Code Summary"
Option Explicit

Implements IDTExtensibility

Public VBInstance As VBIDE.VBE
Private AddInInstance As VBIDE.AddIn

Private mWin As VBIDE.Window
Attribute mWin.VB_VarHelpID = -1
Public mobjDoc As Object

Private MyMenuItem As Office.CommandBarButton
Private MyWindowItem As Office.CommandBarButton
Private MyReplaceItem As Office.CommandBarButton
Private WithEvents ClassClick As VBIDE.VBComponentsEvents
Attribute ClassClick.VB_VarHelpID = -1
Private WithEvents CodeSummaryClick As VBIDE.CommandBarEvents
Attribute CodeSummaryClick.VB_VarHelpID = -1
Private WithEvents WindowCloseClick As VBIDE.CommandBarEvents
Attribute WindowCloseClick.VB_VarHelpID = -1
Private WithEvents ReplaceClick As VBIDE.CommandBarEvents
Attribute ReplaceClick.VB_VarHelpID = -1
Private WithEvents ProjectEvent As VBIDE.VBProjectsEvents
Attribute ProjectEvent.VB_VarHelpID = -1

Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long

Private Sub Class_Terminate()
    Call UnhookWindowsHookEx(hHook1)
    'Call UnhookWindowsHookEx(hHook2)
End Sub


Private Sub ClassClick_ItemActivated(ByVal VBComponent As VBIDE.VBComponent)
    If Not mobjDoc Is Nothing Then
        mobjDoc.DisplayList VBComponent.CodeModule
        'mobjDoc.UpdateSelection
    End If
End Sub

Private Sub ClassClick_ItemSelected(ByVal VBComponent As VBIDE.VBComponent)
    If Not mobjDoc Is Nothing Then
        mobjDoc.DisplayList VBComponent.CodeModule
        'mobjDoc.UpdateSelection
    End If

    If Not mWin Is Nothing Then
        Dim winvis As Boolean
        winvis = GetSetting("ActiveProcedure", "DisplayOptions", "Visible", False)
        If mWin.Visible <> winvis Then
            mWin.Visible = winvis
        End If
    End If
End Sub

Private Sub CodeSummaryClick_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    CreateWindow
End Sub

Public Sub CreateWindow()
    If mWin Is Nothing Then
        Set mWin = AddInInstance.VBE.Windows.CreateToolWindow(AddInInstance, "CSWindow.UserDocument1", "Code Summary", "blah", mobjDoc)
    
        DoEvents
        DoEvents
        mWin.Visible = False
    
        Set mobjDoc.VBInstance = VBInstance
    
        mobjDoc.InitialiseSettings
        If Not VBInstance.ActiveCodePane Is Nothing Then
            mobjDoc.DisplayList VBInstance.ActiveCodePane.CodeModule
        End If
    End If
    mWin.Visible = True
    SaveSetting "ActiveProcedure", "DisplayOptions", "Visible", True
End Sub

'------------------------------------------------------
'this method adds the Add-In to VB
'------------------------------------------------------
Private Sub IDTExtensibility_OnConnection(ByVal VBInst As Object, ByVal ConnectMode As vbext_ConnectMode, ByVal AddInInst As VBIDE.AddIn, custom() As Variant)
    
    On Error GoTo error_handler
    
    'save the vb instance
    Set VBInstance = VBInst
    Set AddInInstance = AddInInst

    If VBInstance.CommandBars.Item("View").Controls.Count > 10 Then
        Set MyMenuItem = VBInstance.CommandBars.Item("View").Controls.Add(, , VBIDE.vbextMenuShortcutCtrlW, 10, False)
    Else
        Set MyMenuItem = VBInstance.CommandBars.Item("View").Controls.Add(, , VBIDE.vbextMenuShortcutCtrlW, , False)
    End If
    
    MyMenuItem.Caption = "Code Summary &Window"
    MyMenuItem.ShortcutText = "Ctrl+E"

    Set CodeSummaryClick = VBInstance.Events.CommandBarEvents(MyMenuItem)
    
    Set MyWindowItem = VBInstance.CommandBars.Item("Window").Controls.Add(, , , 1, True)
    MyWindowItem.Caption = "Close All Windows"
    
    Set WindowCloseClick = VBInstance.Events.CommandBarEvents(MyWindowItem)
    
    Set ProjectEvent = VBInstance.Events.VBProjectsEvents
    
'    Set MyReplaceItem = VBInstance.CommandBars.Item("View").Controls.Add(, , , 11, True)
'    MyReplaceItem.Caption = "Replace Blank Lines"
'
'    Set ReplaceClick = VBInstance.Events.CommandBarEvents(MyReplaceItem)
        
    CreateWindow
    
    Set ConnectObject = Me

    Exit Sub
    
error_handler:
    
    MsgBox Err.Description
    
End Sub

'------------------------------------------------------
'this method removes the Add-In from VB
'------------------------------------------------------
Private Sub IDTExtensibility_OnDisconnection(ByVal RemoveMode As vbext_DisconnectMode, custom() As Variant)
    On Error Resume Next
    MyMenuItem.Delete
    MyWindowItem.Delete
    Call UnhookWindowsHookEx(hHook1)
    'Call UnhookWindowsHookEx(hHook2)
    'MyReplaceItem.Delete
End Sub

Private Sub IDTExtensibility_OnStartupComplete(custom() As Variant)
    Set ProjectEvent = VBInstance.Events.VBProjectsEvents
        
    CreateWindow
    hHook1 = SetWindowsHookEx(WH_KEYBOARD, AddressOf KeyboardProc, 0&, App.ThreadID)
    'hHook2 = SetWindowsHookEx(WH_GETMESSAGE, AddressOf WindowProc, 0&, App.ThreadID)
End Sub

Private Sub IDTExtensibility_OnAddInsUpdate(custom() As Variant)

End Sub

Private Sub ProjectEvent_ItemActivated(ByVal VBProject As VBIDE.VBProject)
    If Not VBProject Is Nothing Then
        Open "C:\LOG.TXT" For Append As #1
        Print #1, Now & " " & VBProject.Name
        Close #1
    End If
    
    Set ClassClick = VBInstance.Events.VBComponentsEvents(VBProject)
End Sub

Private Sub ProjectEvent_ItemAdded(ByVal VBProject As VBIDE.VBProject)
    If Not VBProject Is Nothing Then
        Open "C:\LOG.TXT" For Append As #1
        Print #1, Now & " " & VBProject.Name
        Close #1
    End If
    
    Set ClassClick = VBInstance.Events.VBComponentsEvents(VBProject)
End Sub

Private Sub WindowCloseClick_Click(ByVal CommandBarControl As Object, handled As Boolean, CancelDefault As Boolean)
    Dim oWin As VBIDE.Window
    For Each oWin In VBInstance.Windows
        Select Case oWin.Type
            Case vbext_wt_CodeWindow, vbext_wt_Designer
                oWin.Close
        End Select
    Next
End Sub




 
