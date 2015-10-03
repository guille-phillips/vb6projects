Attribute VB_Name = "KeyboardProcedure"
Option Explicit

Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal ncode As Long, ByVal wParam As Long, lParam As Any) As Long
Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long

Private Type POINTAPI
        x As Long
        y As Long
End Type

Private Type MSG
    hwnd As Long
    message As Long
    wParam As Long
    lParam As Long
    time As Long
    pt As POINTAPI
End Type

Public Const KBH_MASK = &H20000000
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202

' SetWindowsHook() codes
Public Const WH_MIN = (-1)
Public Const WH_MSGFILTER = (-1)
Public Const WH_JOURNALRECORD = 0
Public Const WH_JOURNALPLAYBACK = 1
Public Const WH_KEYBOARD = 2
Public Const WH_GETMESSAGE = 3
Public Const WH_CALLWNDPROC = 4
Public Const WH_CBT = 5
Public Const WH_SYSMSGFILTER = 6
Public Const WH_MOUSE = 7
Public Const WH_HARDWARE = 8
Public Const WH_DEBUG = 9
Public Const WH_SHELL = 10
Public Const WH_FOREGROUNDIDLE = 11
Public Const WH_MAX = 11

Private Const MSGF_DIALOGBOX = 0

Global hHook1 As Long
Global hHook2 As Long

Public ConnectObject As Connect

Private CtrlDown As Boolean

Public Function KeyboardProc(ByVal ncode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    If ncode >= 0 Then
        If wParam = 17 Then
            If (lParam And &HC0000000) = 0 Then
                CtrlDown = True
            ElseIf (lParam And &HC0000000) = &HC0000000 Then
                CtrlDown = False
            End If
        End If
        
        If wParam = Asc("E") And CtrlDown Then
            If (lParam And &HC0000000) = 0 Then
                If Not ConnectObject Is Nothing Then
                    ConnectObject.CreateWindow
                End If
                KeyboardProc = 1
                Exit Function
            End If
        End If

    End If
    
    KeyboardProc = CallNextHookEx(hHook1, ncode, wParam, lParam)
End Function

Public Function WindowProc(ByVal ncode As Long, ByVal wParam As Long, ByRef lParam As MSG) As Long
    If lParam.message = &H118 Then
        If Not ConnectObject Is Nothing Then
            If Not ConnectObject.mobjDoc Is Nothing Then
                'ConnectObject.mobjDoc.UpdateSelection
            End If
        End If
    End If
    WindowProc = CallNextHookEx(hHook2, ncode, wParam, lParam)
End Function

