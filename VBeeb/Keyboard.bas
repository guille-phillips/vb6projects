Attribute VB_Name = "Keyboard"
Option Explicit

Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal ncode As Long, ByVal wParam As Long, lParam As Any) As Long
Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long


Private Declare Function GetKeyboardState Lib "user32" (pbKeyState As Byte) As Long
Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer

Private myKeyState(0 To 255) As Byte
    
Public KeyboardLinks As Long

Private mlWindowsKeyPressed As Long

Public lMapping(511) As Long

Private mlColumn As Long

Private mlRow As Long
Public EnableScan As Long
Private mlORA As Long
Public mbInitialised As Boolean

Private Const VK_LSHIFT = &HA0
Private Const VK_RSHIFT = &HA1
Private Const VK_LCONTROL = &HA2
Private Const VK_RCONTROL = &HA3
Private Const VK_LMENU = &HA4
Private Const VK_RMENU = &HA5

Private mbKeyDown(255) As Boolean
Private mlColumnKeyDown(9) As Long
Private mlColumnKeyDownMask As Long
Private mlColumnKeyMasks(15, 15) As Long
Private mlPowers(15) As Long

Private mbLeftShiftDown As Boolean
Private mbRightShiftDown As Boolean
Private mbLeftControlDown As Boolean
Private mbRightControlDown As Boolean
Private mbLeftMenuDown As Boolean
Private mbRightMenuDown As Boolean

Private hHook1 As Long

Public Sub ClearPressedKeys()
    Erase mbKeyDown
    Erase mlColumnKeyDown
    mlColumnKeyDownMask = 0
End Sub

Private Sub InitialisePowers()
    Dim lIndex As Long
    Dim lPower As Long
    
    ' Debugging.WriteString "Keyboard.InitialisePowers"
    
    lPower = 1
    For lIndex = 0 To 15
        mlPowers(lIndex) = lPower
        lPower = lPower * 2
    Next
End Sub

Private Sub InitialiseColumnKeyMasks()
    Dim lStartBit As Long
    Dim lSizeBit As Long
    Dim lStartMask As Long
    Dim lMask As Long
    Dim lValue As Long
    
    ' Debugging.WriteString "Keyboard.InitialiseColumnKeyMasks"
    
    lStartMask = 1
    For lStartBit = 0 To 15
        lValue = 0
        lMask = lStartMask
        For lSizeBit = 0 To 15
            lValue = lValue + lMask
            mlColumnKeyMasks(lStartBit, lSizeBit) = lValue
            If (lStartBit + lSizeBit) = 15 Then
                lMask = 1
            Else
                lMask = lMask * 2
            End If
        Next
        lStartMask = lStartMask * 2
    Next
End Sub

Public Sub InitialiseKeyboardLinks()
    Dim lBitMask As Long
    Dim lLinkBit As Long

    ' Debugging.WriteString "Keyboard.InitialiseKeyboardLinks"
    
    KeyboardLinks = GetSetting("VBeeb", "KeyboardLinks", "KeyboardLinks", 0)
    lBitMask = 1
    For lLinkBit = 0 To 7
        If KeyboardLinks And lBitMask Then
            mbKeyDown(9 - lLinkBit) = True
        Else
            mbKeyDown(9 - lLinkBit) = False
        End If
        lBitMask = lBitMask * 2
    Next
End Sub

Private Sub InitialiseKeyMappings()
    Dim lIndex As Long
    
    ' Debugging.WriteString "Keyboard.InitialiseKeyMappings"
    
    For lIndex = 0 To UBound(lMapping)
        lMapping(lIndex) = -1
    Next
    
    lMapping(vbKeyShift) = 0
    lMapping(vbKeyControl) = 1

    lMapping(vbKeyA) = 65
    lMapping(vbKeyB) = 100
    lMapping(vbKeyC) = 82
    lMapping(vbKeyD) = 50
    lMapping(vbKeyE) = 34
    lMapping(vbKeyF) = 67
    lMapping(vbKeyG) = 83
    lMapping(vbKeyH) = 84
    lMapping(vbKeyI) = 37
    lMapping(vbKeyJ) = 69
    lMapping(vbKeyK) = 70
    lMapping(vbKeyL) = 86
    lMapping(vbKeyM) = 101
    lMapping(vbKeyN) = 85
    lMapping(vbKeyO) = 54
    lMapping(vbKeyP) = 55
    lMapping(vbKeyQ) = 16
    lMapping(vbKeyR) = 51
    lMapping(vbKeyS) = 81
    lMapping(vbKeyT) = 35
    lMapping(vbKeyU) = 53
    lMapping(vbKeyV) = 99
    lMapping(vbKeyW) = 33
    lMapping(vbKeyX) = 66
    lMapping(vbKeyY) = 68
    lMapping(vbKeyZ) = 97
    
    lMapping(vbKeySpace) = 98
    
    lMapping(vbKey0) = 39
    lMapping(vbKey1) = 48
    lMapping(vbKey2) = 49
    lMapping(vbKey3) = 17
    lMapping(vbKey4) = 18
    lMapping(vbKey5) = 19
    lMapping(vbKey6) = 52
    lMapping(vbKey7) = 36
    lMapping(vbKey8) = 21
    lMapping(vbKey9) = 38
    
    lMapping(vbKeyF1) = 32
    lMapping(vbKeyF2) = 113
    lMapping(vbKeyF3) = 114
    lMapping(vbKeyF4) = 115
    lMapping(vbKeyF5) = 20
    lMapping(vbKeyF6) = 116
    lMapping(vbKeyF7) = 117
    lMapping(vbKeyF8) = 22
    lMapping(vbKeyF9) = 118
    lMapping(vbKeyF10) = 119

    lMapping(vbKeyReturn) = 73
    lMapping(vbKeyEscape) = 112
    lMapping(223) = 112 ' back tick = escape
    
    lMapping(vbKeyTab) = 96
    
    lMapping(273) = 105 ' COPY
    lMapping(107) = 105 ' COPY
    lMapping(20) = 64 ' CAPSLOCK
    lMapping(220) = 80 ' SHIFTLOCK
    
    lMapping(vbKeyUp) = 57
    lMapping(vbKeyRight) = 121
    lMapping(vbKeyDown) = 41
    lMapping(vbKeyLeft) = 25
    
    lMapping(vbKeyBack) = 89
    lMapping(93) = 89
    
    lMapping(vbKeyInsert) = 120
    lMapping(vbKeyDelete) = 40

    
    lMapping(vbKeyHome) = 25
    lMapping(vbKeyEnd) = 57
    lMapping(vbKeyPageUp) = 121
    lMapping(vbKeyPageDown) = 41
    
    lMapping(189) = 23 ' MINUS
    lMapping(187) = 24 ' CIRCUMFLEX
    
    lMapping(219) = 71 ' LEFT SQUARE
    lMapping(221) = 56 ' RIGHT SQUARE
    lMapping(222) = 88 ' BACKSLASH
    
    lMapping(186) = 87 ' SEMI COLON
    lMapping(192) = 72 ' COLON
    
    lMapping(188) = 102 ' COMMA
    lMapping(190) = 103 ' DOT
    lMapping(191) = 104 ' SLASH
    
    lMapping(vbKeyF12) = 256 ' BREAK
    lMapping(vbKeyPause) = 256 ' BREAK
    
    mbInitialised = True
End Sub

Public Sub InitialiseKeyboard()
    ' Debugging.WriteString "Keyboard.InitialiseKeyboard"
    
    InitialisePowers
    InitialiseColumnKeyMasks
    InitialiseKeyboardLinks
    InitialiseKeyMappings
End Sub

Public Sub TerminateKeyboard()
    UnhookWindowsHookEx hHook1
End Sub

Public Sub WriteRegister(ByVal lValue As Long)
    ' Debugging.WriteString "Keyboard.WriteRegister"
    
    mlORA = lValue
    If EnableScan = 0 Then
        mlColumn = lValue And &HF&
        If mlColumn < 10 Then
            If mlColumnKeyDown(mlColumn) > 0 Then
                SystemVIA6522.AssertCA2
            End If
        End If
        If mbKeyDown(lMapping(vbKeyTab)) Then ' tab key kludge, key up event not triggered
            If GetKeyState(vbKeyTab) >= 0 Then
                WindowsKeyUp vbKeyTab
                Console.mlPreviousKey = 0
            End If
        End If
        
        If mbKeyDown(lValue) Then
            gyMem(&HFE4F&) = lValue Or &H80
        Else
            gyMem(&HFE4F&) = lValue
        End If
    End If
End Sub

Public Sub Tick(ByVal lCycles As Long)
    ' Debugging.WriteString "Keyboard.Tick"
    
    If EnableScan = 1 Then
        If lCycles < 16 Then
            If (mlColumnKeyDownMask And mlColumnKeyMasks(mlColumn, lCycles)) <> 0 Then
                SystemVIA6522.AssertCA2
            End If
        Else
            If mlColumnKeyDownMask <> 0 Then
                SystemVIA6522.AssertCA2
            End If
        End If
        mlColumn = (mlColumn + lCycles) And &HF&
    End If
End Sub

Public Sub WindowsKeyDown(ByVal lScanCode As Long)
    Dim lIndex As Long
    Dim lScanCodeConverted As Long
    Dim bRightControl As Boolean
    Dim bRightShift As Boolean
    Dim mlKeyColumn As Long
    
    ' Debugging.WriteString "Keyboard.WindowsKeyDown"
    
    Select Case lScanCode
        Case vbKeyShift
            If GetKeyState(VK_RSHIFT) < 0 Then
                lScanCode = lScanCode + 256
                mbRightShiftDown = True
            Else
                mbLeftShiftDown = True
            End If
        Case vbKeyControl
            If GetKeyState(VK_RCONTROL) < 0 Then
                lScanCode = lScanCode + 256
                mbRightControlDown = True
            Else
                mbLeftControlDown = True
            End If
        Case vbKeyMenu
            If GetKeyState(VK_RMENU) < 0 Then
                lScanCode = lScanCode + 256
                mbRightMenuDown = True
            Else
                mbLeftMenuDown = True
            End If
    End Select
    
    'Debug.Print "Down:" & lScanCode
    
    lScanCodeConverted = lMapping(lScanCode)
    
    If lScanCodeConverted <> -1 And lScanCodeConverted <> 256 Then
        If Not mbKeyDown(lScanCodeConverted) Then
            If lScanCodeConverted > 1 Then
                mlKeyColumn = lScanCodeConverted And &HF&
                mlColumnKeyDown(mlKeyColumn) = mlColumnKeyDown(mlKeyColumn) + 1
                mlColumnKeyDownMask = mlColumnKeyDownMask Or mlPowers(mlKeyColumn)
            End If
        End If
        mbKeyDown(lScanCodeConverted) = True
    End If
End Sub

Public Sub WindowsKeyUp(ByVal lScanCode As Long)
    Dim lIndex As Long
    Dim lScanCodeConverted As Long
    Dim lIndex2 As Long
    Dim mlKeyColumn As Long
    
    ' Debugging.WriteString "Keyboard.WindowsKeyUp"

    Select Case lScanCode
        Case vbKeyShift
            If Not mbLeftShiftDown Then
                lScanCode = lScanCode + 256
                mbRightShiftDown = False
            ElseIf Not mbRightShiftDown Then
                mbLeftShiftDown = False
            End If
        Case vbKeyControl
            If Not mbLeftControlDown Then
                lScanCode = lScanCode + 256
                mbRightControlDown = False
            ElseIf Not mbRightControlDown Then
                mbLeftControlDown = False
            End If
        Case vbKeyMenu
            If Not mbLeftMenuDown Then
                lScanCode = lScanCode + 256
                mbRightMenuDown = False
            ElseIf Not mbRightShiftDown Then
                mbLeftMenuDown = False
            End If
    End Select
    
    'Debug.Print "Up:" & lScanCode
            
    lScanCodeConverted = lMapping(lScanCode)
    'Debug.Print "Up:" & lScanCodeConverted
    
    If lScanCodeConverted <> -1 And lScanCodeConverted <> 256 Then
        If mbKeyDown(lScanCodeConverted) Then
            If lScanCodeConverted > 1 Then
                mlKeyColumn = lScanCodeConverted And &HF&
                
                mlColumnKeyDown(mlKeyColumn) = mlColumnKeyDown(mlKeyColumn) - 1
        
                If mlColumnKeyDown(mlKeyColumn) <= 0 Then
                    mlColumnKeyDownMask = mlColumnKeyDownMask Xor mlPowers(mlKeyColumn)
                    mlColumnKeyDown(mlKeyColumn) = 0
                End If
            End If
        End If
        mbKeyDown(lScanCodeConverted) = False
    End If
End Sub
