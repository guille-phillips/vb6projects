VERSION 5.00
Begin VB.Form Monitor 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Code Monitor"
   ClientHeight    =   5430
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   8415
   DrawWidth       =   10
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   362
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   561
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExitDebugMode 
      Caption         =   "Exit Debug Mode"
      Height          =   495
      Left            =   5760
      TabIndex        =   2
      Top             =   4800
      Width           =   2055
   End
   Begin VB.TextBox txtExpressionEdit 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFC0&
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   3600
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   360
      Width           =   2055
   End
   Begin VB.VScrollBar scrMem 
      Height          =   5175
      LargeChange     =   64
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   255
   End
   Begin VB.Image imgArrow 
      Height          =   255
      Left            =   960
      Picture         =   "Monitor.frx":0000
      Top             =   1320
      Visible         =   0   'False
      Width           =   180
   End
End
Attribute VB_Name = "Monitor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function DrawText Lib "user32.dll" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, ByRef lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32.dll" (ByVal crColor As Long) As Long
Private Declare Function FillRect Lib "user32.dll" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
Private Declare Function GetTextExtentPoint32 Lib "gdi32" Alias "GetTextExtentPoint32A" (ByVal hdc As Long, ByVal lpsz As String, ByVal cbString As Long, lpSize As SIZE) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function DrawFocusRect Lib "user32.dll" (ByVal hdc As Long, lpRect As RECT) As Long

Private Declare Function SetWindowPos Lib "user32" _
      (ByVal hwnd As Long, _
      ByVal hWndInsertAfter As Long, _
      ByVal X As Long, _
      ByVal Y As Long, _
      ByVal cx As Long, _
      ByVal cy As Long, _
      ByVal wFlags As Long) As Long
      
Private Const SWP_NOSIZE = 1
Private Const SWP_NOMOVE = 2
Private Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE
Private Const HWND_TOPMOST = -1

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type SIZE
    Width As Long
    Height As Long
End Type


Private Const DT_LEFT = &H0
Private Const DT_TOP = &H0
Private Const DT_CENTER = &H1
Private Const DT_RIGHT = &H2
Private Const DT_VCENTER = &H4
Private Const DT_BOTTOM = &H8
Private Const DT_WORDBREAK = &H10
Private Const DT_SINGLELINE = &H20
Private Const DT_NOPREFIX = &H800

Private mhBrush As Long
Private mhBrushEmpty As Long
Private mhBrushHighlightText As Long
Private mhBrushSelected As Long


Private Const TraceWindowMarginTop As Long = 8
Private Const TraceWindowLineHeight As Long = 15
Private Const TraceWindowLines As Long = 20
Private Const TraceWindowMarginLeft As Long = 32

Private mlTraceTopLocation As Long
Private mlTraceBottomLocation As Long
Private mlTraceInstructionLocation As Long

Private mlTraceAddresses(TraceWindowLines) As Long
Private mlTraceOperandWidths(TraceWindowLines) As Long
Private mlTraceOperands(TraceWindowLines) As Long

Private Const msPSRFlags = "NVBD IZC"

'Private mbiBreakpoints() As BreakpointInfo
'Private mlBreakpointCount As Long

Private Enum BreakpointTypes
    btNormal
    btOnceOnly
End Enum

Private Type BreakpointInfo
    Location As Long
    RomSelect As Long
    BreakType As BreakpointTypes
End Type

Private mbNoUpdate As Boolean

Private Type BoxInfo
    Top As Long
    Right As Long
    Bottom As Long
    Left As Long
    BackColour As Long
    ForeColour As Long
    Text As String
    Focus As Boolean
    Tag As Long
    Column As Long
    Row As Long
    Centred As Boolean
End Type

Private mbiBoxes() As BoxInfo
Private mlBoxesCount As Long
Private mlLastBoxIndex As Long
Private mlSelectedBoxIndex As Long

Private Const DigitBoxWidth As Long = 12
Private Const DigitBoxHeight As Long = 15

Private Const mlBackColour As Long = &HD8E9EC
Private Const mlSelectedBackColour As Long = &HF9BB7B
Private Const mlDisabledBackColour As Long = &HC8C8C8
Private Const mlDisabledForeColour As Long = &H787878

Private mlWatchCursorAddress As Long
Private mlWatchCursorRow As Long

Private Type Pair
    Device As Long
    Value As Long
End Type

Private moExpression As ISaffronObject
Private Type Operation
    OperationIndex As Long
    ResultPos As Long
    OperandPos As Long
    Level As Long
End Type
Private Type Expression
    Constants() As Pair
    ConstantsCount As Long
    Operations() As Operation
    OperationsCount As Long
    Reads() As Pair
    ReadsCount As Long
    Writes() As Pair
    WritesCount As Long
End Type

Private mexpTempExpression As Expression
Private mlSelectedWatchRow As Long

Private Enum BreakTypes
    None
    Break
    BreakOnce
End Enum

Private Type WatchInfo2
    WatchExpressionText As String
    WatchExpression As Expression
    ExpressionOk As Boolean
    WidthBytes As Long
    Base As Long
    BitNames As String
    Break As BreakTypes
    ValueUpdateable As Boolean
End Type

Private Const WatchWindowMarginTop As Long = 8
Private Const WatchWindowLineHeight As Long = 15
Private Const WatchWindowSpacer As Long = 5
Private Const WatchWindowMarginLeft As Long = 250

Private mwiWatchLocations() As WatchInfo2
Private mlTotalWatchLocations As Long
Private mlBreakWatchRow As Long

Private mlPreviousPC As Long

Private Sub cmdExitDebugMode_Click()
    Me.Hide
    Processor6502.StopReason = srNormal
    ProcessorStopped
End Sub

Private Sub Form_Activate()
    SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS
End Sub

Private Sub Form_Initialize()
    InitialiseMonitor
End Sub

Public Sub InitialiseMonitor()
    InitialiseCompiler
    
    txtExpressionEdit.Width = DigitBoxWidth * 10
    txtExpressionEdit.Height = DigitBoxHeight
    txtExpressionEdit.Visible = False
    txtExpressionEdit.Left = WatchWindowMarginLeft + DigitBoxWidth * 3 + WatchWindowSpacer
    
    mlLastBoxIndex = -1
    mlSelectedBoxIndex = -1
    mlWatchCursorAddress = -1
    mlBreakWatchRow = -1
    
    AddWatch "A", 1, 3, BreakTypes.None
    AddWatch "X", 1, 3, BreakTypes.None
    AddWatch "Y", 1, 3, BreakTypes.None
    AddWatch "S", 1, 3, BreakTypes.None
    AddWatch "P", 1, 3, BreakTypes.None
    AddWatch "PC", 2, 3, BreakTypes.None
    AddWatch "PPC", 2, 3, BreakTypes.None
    AddWatch "ROM", 1, 3, BreakTypes.None
    
End Sub

'Private Sub AddBreakpoint(ByVal lLocation As Long, ByVal lRomSelect As Long, ByVal btBreakpointType As BreakpointTypes)
'    ReDim Preserve mbiBreakpoints(mlBreakpointCount)
'    With mbiBreakpoints(mlBreakpointCount)
'        .Location = lLocation
'        .RomSelect = lRomSelect
'        .BreakType = btBreakpointType
'    End With
'    mlBreakpointCount = mlBreakpointCount + 1
'End Sub

'Private Sub RemoveBreakpointAtAddress(ByVal lLocation As Long, ByVal lRomSelect As Long)
'    Dim lFindIndex As Long
'
'    For lFindIndex = 0 To mlBreakpointCount - 1
'        If mbiBreakpoints(lFindIndex).Location = lLocation Then
'            If mbiBreakpoints(lFindIndex).RomSelect = lRomSelect Then
'                RemoveBreakpoint lFindIndex
'                Exit Sub
'            End If
'        End If
'    Next
'End Sub

Private Sub RemoveWatch(ByVal lIndex As Long)
    Dim lWatchIndex As Long
    
    For lWatchIndex = lIndex To mlTotalWatchLocations - 2
        mwiWatchLocations(lWatchIndex) = mwiWatchLocations(lWatchIndex + 1)
    Next
    
    mlTotalWatchLocations = mlTotalWatchLocations - 1
    If mlTotalWatchLocations > 0 Then
        ReDim Preserve mwiWatchLocations(mlTotalWatchLocations - 1)
    Else
        Erase mwiWatchLocations
    End If
    
End Sub

'Private Sub RemoveBreakpoint(ByVal lIndex As Long)
'    Dim lFindIndex As Long
'
'    For lFindIndex = lIndex To mlBreakpointCount - 2
'        mbiBreakpoints(lFindIndex) = mbiBreakpoints(lFindIndex + 1)
'    Next
'
'    mlBreakpointCount = mlBreakpointCount - 1
'    If mlBreakpointCount > 0 Then
'        ReDim Preserve mbiBreakpoints(mlBreakpointCount - 1)
'    Else
'        Erase mbiBreakpoints
'    End If
'End Sub

Private Sub UpdateMonitor(Optional ByVal bUpdateScroll As Boolean = True)
    Erase mbiBoxes
    mlBoxesCount = 0
    
    ShowTraceWindow
    ShowWatchWindow
    
    If bUpdateScroll Then
        mbNoUpdate = True
        scrMem.Value = Processor6502.PC \ 2
        mbNoUpdate = False
    End If
End Sub

Public Sub ShowTraceWindow()
    Dim lRow As Long
    Dim lCurrentLocation As Long
    Dim lCurrentInstructionLength As Long
    Dim lType As Long
    Dim lNext As Long
    Dim lFindBreakpointIndex As Long
    Dim vDisplayedInstruction As Variant
    
    Me.Cls
    
    lCurrentLocation = mlTraceTopLocation
    For lRow = 0 To TraceWindowLines - 1
        If lCurrentLocation = mlTraceInstructionLocation Then
            lType = 1
        Else
            lType = 0
        End If
        
        mlTraceBottomLocation = lCurrentLocation
        
        mlTraceAddresses(lRow) = lCurrentLocation
        
'        For lFindBreakpointIndex = 0 To mlBreakpointCount - 1
'            If mbiBreakpoints(lFindBreakpointIndex).Location = lCurrentLocation Then
'                If mbiBreakpoints(lFindBreakpointIndex).RomSelect > 0 Then
'                    If mbiBreakpoints(lFindBreakpointIndex).RomSelect = RomSelect.SelectedBank Then
'                        lType = lType Or 2
'                    End If
'                Else
'                    lType = lType Or 2
'                End If
'            End If
'        Next
        
        vDisplayedInstruction = DisplayInstruction(lCurrentLocation, lRow, True, lType)
        lCurrentInstructionLength = vDisplayedInstruction(0)
        mlTraceOperandWidths(lRow) = vDisplayedInstruction(0) - 1
        mlTraceOperands(lRow) = vDisplayedInstruction(1)
        lCurrentLocation = lCurrentLocation + lCurrentInstructionLength
    Next
    
    lNext = AddBox(TraceWindowMarginLeft, 430, DigitBoxWidth * 5, DigitBoxHeight, mlBackColour, vbBlack, "ROMSEL", 2, 0, 0)
    lNext = AddBox(lNext, 430, DigitBoxWidth, DigitBoxHeight, vbWhite, vbBlack, HexNum$(RomSelect.SelectedBank, 1), 2, 1, 0)
    
End Sub

Private Property Let InstructionLocation(ByVal lMemoryLocation As Long)
    mlTraceInstructionLocation = lMemoryLocation
    If mlTraceInstructionLocation < mlTraceTopLocation Then
        mlTraceTopLocation = mlTraceInstructionLocation
    End If
    If mlTraceInstructionLocation > mlTraceBottomLocation Then
        mlTraceTopLocation = mlTraceInstructionLocation
    End If
End Property

Private Function DisplayInstruction(ByVal lMemoryLocation As Long, ByVal lRow As Long, ByVal bAsInstruction As Boolean, ByVal lType As Long) As Variant
    Dim sMemoryLocation As String
    Dim sMemHex As String
    Dim yInstruction(2) As Byte
    Dim vInstruction As Variant
    Dim lBackColour As Long
    Dim lNext As Long
    
    If lType And 1& Then
        lBackColour = RGB(255, 255, 0)
    Else
        lBackColour = RGB(255, 255, 255)
    End If
    
    If lType And 2& Then
        lNext = AddBox(TraceWindowMarginLeft, lRow * TraceWindowLineHeight + TraceWindowMarginTop, 16, TraceWindowLineHeight, lBackColour, RGB(0, 0, 0), "*", 0, 0, lRow)
    Else
        lNext = AddBox(TraceWindowMarginLeft, lRow * TraceWindowLineHeight + TraceWindowMarginTop, 16, TraceWindowLineHeight, lBackColour, RGB(0, 0, 0), "", 0, 0, lRow)
    End If
    
    sMemoryLocation = BaseNum(lMemoryLocation, 4, 16)
    lNext = AddBox(lNext, lRow * TraceWindowLineHeight + TraceWindowMarginTop, 45, TraceWindowLineHeight, lBackColour, RGB(0, 0, 0), sMemoryLocation, 0, 1, lRow)
    
    If Not bAsInstruction Then
        sMemHex = BaseNum(gyMem(lMemoryLocation), 2, 16)
        AddBox lNext, lRow * TraceWindowLineHeight + TraceWindowMarginTop, 200, TraceWindowLineHeight, lBackColour, RGB(0, 0, 0), sMemHex, 0, 2, lRow
        DisplayInstruction = Array(1, 0)
    Else
        If lMemoryLocation <= &HFFFD& Then
            CopyMemory yInstruction(0), gyMem(lMemoryLocation), 3&
        Else
            Dim lLocation As Long
            For lLocation = 0 To 2
                yInstruction(lLocation) = gyMem((lLocation + lMemoryLocation) And &HFFFF&)
            Next
        End If
        
        vInstruction = Disassembler.DisassembleInstruction(lMemoryLocation, yInstruction, False)
        AddBox lNext, lRow * TraceWindowLineHeight + TraceWindowMarginTop, 120, TraceWindowLineHeight, lBackColour, RGB(0, 0, 0), vInstruction(0), 0, 2, lRow
        DisplayInstruction = Array(vInstruction(1), vInstruction(2))
    End If
End Function

Private Sub ShowWatchWindow()
    Dim lBases(3) As Long
    Dim lBaseWidths(3) As Long
    Dim lWatchIndex As Long
    Dim lBoxTop As Long
    Dim lNext As Long
    Dim lValue As Long
    
    lBases(3) = 16
    lBases(2) = 10
    lBases(1) = 8
    lBases(0) = 2
    
    lBaseWidths(3) = 2
    lBaseWidths(2) = 3
    lBaseWidths(1) = 3
    lBaseWidths(0) = 8
    
    For lWatchIndex = 0 To mlTotalWatchLocations - 1
        With mwiWatchLocations(lWatchIndex)
            lBoxTop = lWatchIndex * WatchWindowLineHeight + WatchWindowMarginTop
            
            ' Break Arrow
            If lWatchIndex = mlBreakWatchRow Then
                imgArrow.Left = WatchWindowMarginLeft - 18
                imgArrow.Top = lBoxTop - 2
                mlBreakWatchRow = -1
                imgArrow.Visible = True
            End If
            
            ' Break
            lNext = AddBox(WatchWindowMarginLeft, lBoxTop, DigitBoxWidth * 3, DigitBoxHeight, vbWhite, IIf(.Break > None, vbRed, mlDisabledForeColour), Choose(.Break + 1, "Break", "Break", "Once"), 1, 0, lWatchIndex, True) + WatchWindowSpacer
            
            ' Expression
            lNext = AddBox(lNext, lBoxTop, DigitBoxWidth * 10, DigitBoxHeight, vbWhite, IIf(.ExpressionOk, vbBlack, vbRed), .WatchExpressionText, 1, 1, lWatchIndex) + WatchWindowSpacer
            
            ' Width
            lNext = AddBox(lNext, lBoxTop, DigitBoxWidth * 1, DigitBoxHeight, vbWhite, vbBlack, .WidthBytes, 1, 2, lWatchIndex, True) + WatchWindowSpacer
            
            ' Base
            lNext = AddBox(lNext, lBoxTop, DigitBoxWidth * 2, DigitBoxHeight, vbWhite, vbBlack, lBases(.Base), 1, 3, lWatchIndex, True) + WatchWindowSpacer
            
            ' Value
            If .ExpressionOk Then
                lValue = EvaluateExpression(lWatchIndex)
                lNext = AddNumber(lNext, lBoxTop, lValue, lBases(.Base), lBaseWidths(.Base) * .WidthBytes, 1, 4, lWatchIndex) + WatchWindowSpacer
            End If
            
            ' Remove
            lNext = AddBox(lNext, lBoxTop, DigitBoxWidth, DigitBoxHeight, vbRed, vbWhite, "x", 1, 5, lWatchIndex, True) + WatchWindowSpacer
            
'            If .BitNames = "" Then
'                If Not .ShowValueOrMask Then
'                    lNext = AddNumber(lNext, lBoxTop, lValue, lBases(.Base), lBaseWidths(.Base) * .WidthBytes, 1, 6, lWatchIndex)
'                Else
'                    lNext = AddMaskNumber(lNext, lBoxTop, .BreakWhenRead.Value, .BreakWhenRead.Mask, lBases(.Base), lBaseWidths(.Base) * .WidthBytes, 1, 6, lWatchIndex)
'                End If
'            Else
'                lNext = AddLabelledNumber(lNext, lBoxTop, lValue, lBases(.Base), lBaseWidths(.Base) * .WidthBytes, .BitNames, 1, 6, lWatchIndex)
'            End If
        End With
    Next
End Sub

Private Sub AddWatch(ByVal sExpression As String, ByVal lWidthBytes As Long, ByVal lBase As Long, btBreak As BreakTypes)
    ReDim Preserve mwiWatchLocations(mlTotalWatchLocations)
    With mwiWatchLocations(mlTotalWatchLocations)
        .WatchExpressionText = sExpression
        If ParseUserExpression(sExpression) Then
            .WatchExpression = mexpTempExpression
            .ExpressionOk = True
        Else
            .ExpressionOk = False
        End If
        .Base = lBase
        .Break = btBreak
        .WidthBytes = lWidthBytes
    End With
    mlTotalWatchLocations = mlTotalWatchLocations + 1
End Sub

Private Function AddBox(ByVal lX As Long, ByVal lY As Long, ByVal lWidth As Long, ByVal lHeight As Long, ByVal lBackColour As Long, ByVal lForeColour As Long, ByVal sText As String, ByVal lTag As Long, ByVal lColumn As Long, ByVal lRow As Long, Optional ByVal bCentre As Boolean) As Long
    ReDim Preserve mbiBoxes(mlBoxesCount)
    
    With mbiBoxes(mlBoxesCount)
        .Left = lX
        .Right = lX + lWidth
        .Top = lY
        .Bottom = lY + lHeight
        .BackColour = lBackColour
        .ForeColour = lForeColour
        .Text = sText
        .Tag = lTag
        .Column = lColumn
        .Row = lRow
        .Centred = bCentre
    End With
    
    AddBox = RenderBox(mlBoxesCount, bCentre)
    
    mlBoxesCount = mlBoxesCount + 1
End Function

Private Function FindBox(ByVal lX As Long, ByVal lY As Long) As Long
    Dim lBoxIndex As Long
    
    FindBox = -1
    
    For lBoxIndex = 0 To mlBoxesCount - 1
        With mbiBoxes(lBoxIndex)
            If lX >= .Left Then
                If lX <= .Right Then
                    If lY >= .Top Then
                        If lY <= .Bottom Then
                            FindBox = lBoxIndex
                            Exit Function
                        End If
                    End If
                End If
            End If
        End With
    Next
End Function

Private Function RenderBox(ByVal lBoxIndex As Long, Optional ByVal bCentre As Boolean, Optional ByVal bFocus As Boolean) As Long
    Dim rectArea As RECT
    Dim hBrush As Long
    Dim hPen As Long
    
    With mbiBoxes(lBoxIndex)
        hBrush = CreateSolidBrush(.BackColour)
        SetTextColor Me.hdc, .ForeColour
                
        rectArea.Top = .Top
        rectArea.Left = .Left
        rectArea.Bottom = .Bottom
        rectArea.Right = .Right
        
        RenderBox = .Right + 1
        
        FillRect Me.hdc, rectArea, hBrush
        If bFocus Then
            DrawFocusRect Me.hdc, rectArea
        End If
        DrawText Me.hdc, .Text, Len(.Text), rectArea, DT_SINGLELINE Or DT_VCENTER Or DT_NOPREFIX Or (DT_CENTER * -bCentre)
        
        DeleteObject hBrush
    End With
End Function

Private Sub UpdateWatch(ByVal lWatchIndex As Long, ByVal lColumn As Long, ByVal sNewDigit As String)
    Dim lBases(3) As Long
    Dim lBaseWidths(3) As Long
    Dim sDigits As String
    Dim lWidthIndex As Long
    Dim lValue As Long
    
    lBases(3) = 16
    lBases(2) = 10
    lBases(1) = 8
    lBases(0) = 2
    lBaseWidths(3) = 2
    lBaseWidths(2) = 3
    lBaseWidths(1) = 3
    lBaseWidths(0) = 8
    
'    With mwiWatchLocations(lWatchIndex)
'        If ConvertBase(sNewDigit, 16) > lBases(.Base) Then
'            Exit Sub
'        End If
'
'        sDigits = BaseNum(.CurrentValue, lBaseWidths(.Base) * .WidthBytes, lBases(.Base))
'        Mid$(sDigits, lColumn + 1, 1) = sNewDigit
'        lValue = ConvertBase(sDigits, lBases(.Base))
'        .CurrentValue = lValue
'        Select Case .Device
'            Case 0 ' mem
'                For lWidthIndex = 1 To .WidthBytes
'                    Memory.Mem(.Location + lWidthIndex - 1) = lValue And &HFF&
'                    lValue = lValue \ 256
'                Next
'            Case 1 ' processor
'
'        End Select
'    End With
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim vInstruction As Variant
    Dim lNextInstruction As Long
    Dim yInstruction(2) As Byte
    Dim lMemoryLocation As Long
    Dim lStack As Long
    Dim lWatchIndex As Long
    Dim bExit As Boolean
    Dim lFindBreakpointIndex As Long
    
    Select Case KeyCode
        Case vbKeyA To vbKeyF, vbKey0 To vbKey9
            If mlSelectedBoxIndex <> -1 Then
                Select Case mbiBoxes(mlSelectedBoxIndex).Tag
                    Case 0 ' Trace
                    Case 1 ' Watch
                        Select Case mbiBoxes(mlSelectedBoxIndex).Column And &HF&
                            Case 6 ' Number
                                UpdateWatch mbiBoxes(mlSelectedBoxIndex).Row, mbiBoxes(mlSelectedBoxIndex).Column \ &H10&, UCase$(Chr$(KeyCode))
                                UpdateMonitor False
                        End Select
                    Case 2 ' Rom Select
                        If mbiBoxes(mlSelectedBoxIndex).Column = 1 Then
                            RomSelect.SetRom HexToDec(UCase$(Chr$(KeyCode)))
                            Cls
                            UpdateMonitor False
                        End If
                End Select
            End If
        Case vbKeyS
            If mlWatchCursorAddress <> -1 Then
                Processor6502.PC = mlWatchCursorAddress
                InstructionLocation = mlWatchCursorAddress
                UpdateMonitor False
            End If
        Case vbKeyR
            If mlWatchCursorRow <> -1 Then
                If mlTraceOperandWidths(mlWatchCursorRow) > 0 Then
                    AddWatch "?" & HexNum(mlTraceOperands(mlWatchCursorRow), 4) & "H", 1, 3, None
                    Cls
                    UpdateMonitor False
                End If
            End If
        Case vbKeyM
            If mlWatchCursorRow <> -1 Then
                AddWatch "?" & HexNum(mlTraceAddresses(mlWatchCursorRow), 4) & "H", 2, 3, None
                Cls
                UpdateMonitor False
            End If
        Case 119 'F8
            Select Case Shift
                Case 0, 1 ' Step Into / Step Over
                    If Shift = 1 And gyMem(Processor6502.PC) = &H20 Then
                        If Processor6502.PC > &H8000& And Processor6502.PC < &HBFFF& Then
                            AddWatch "(PC=" & HexNum(Processor6502.PC + 3, 4) & "H)&(ROM=H)", 2, 3, BreakOnce
                        Else
                            AddWatch "PC=" & HexNum(Processor6502.PC + 3, 4) & "H", 2, 3, BreakOnce
                        End If
                        Processor6502.StopReason = srDebugRun
                        Keyboard.ClearPressedKeys
                        Me.Hide
                        bExit = False
                        Do
                            Processor6502.Execute
'                            For lFindBreakpointIndex = 0 To mlBreakpointCount - 1
'                                If mbiBreakpoints(lFindBreakpointIndex).Location = Processor6502.PC Then
'                                    If mbiBreakpoints(lFindBreakpointIndex).RomSelect > 0 Then
'                                        If mbiBreakpoints(lFindBreakpointIndex).RomSelect = RomSelect.SelectedBank Then
'                                            If mbiBreakpoints(lFindBreakpointIndex).BreakType = btOnceOnly Then
'                                                RemoveBreakpoint lFindBreakpointIndex
'                                            End If
'                                            bExit = True
'                                        End If
'                                    Else
'                                        If mbiBreakpoints(lFindBreakpointIndex).BreakType = btOnceOnly Then
'                                            RemoveBreakpoint lFindBreakpointIndex
'                                        End If
'                                        bExit = True
'                                    End If
'                                End If
'                            Next
                        Loop Until bExit Or Processor6502.StopReason <> srDebugRun
                        If Processor6502.StopReason = srDebugRun Then
                            VideoULA.UpdateWholeDisplay
                            Me.Show
                            mlTraceInstructionLocation = Processor6502.PC
                            UpdateMonitor
                            Processor6502.StopReason = srDebugBreak
                        Else
                            ProcessorStopped
                        End If
                    Else
                        Processor6502.Execute
                        If Processor6502.StopReason = srDebugBreak Then
                            VideoULA.UpdateWholeDisplay
                            InstructionLocation = Processor6502.PC
                            UpdateMonitor
                        Else
                            Keyboard.ClearPressedKeys
                            Me.Hide
                            ProcessorStopped
                        End If
                    End If

                Case 2 ' Run to cursor
                    If mlWatchCursorAddress <> -1 Then
                        If Processor6502.PC > &H8000& And Processor6502.PC < &HBFFF& Then
                            AddWatch "(PC=" & HexNum(mlWatchCursorAddress + 3, 4) & "H)&(ROM=H)", 2, 3, BreakOnce
                        Else
'                            AddBreakpoint Processor6502.PC + 3, 0, btOnceOnly
                            AddWatch "PC=" & HexNum(mlWatchCursorAddress + 3, 4) & "H", 2, 3, BreakOnce
                        End If
                        
                        DoDebugRun
                    End If
                Case 3 ' Step Out
                    lStack = Processor6502.S
                    Processor6502.StopReason = srDebugRun
                    Keyboard.ClearPressedKeys
                    Me.Hide
                    Do
                        Processor6502.Execute
'                        For lFindBreakpointIndex = 0 To mlBreakpointCount - 1
'                            If mbiBreakpoints(lFindBreakpointIndex).Location = Processor6502.PC Then
'                                If mbiBreakpoints(lFindBreakpointIndex).RomSelect > 0 Then
'                                    If mbiBreakpoints(lFindBreakpointIndex).RomSelect = RomSelect.SelectedBank Then
'                                        If mbiBreakpoints(lFindBreakpointIndex).BreakType = btOnceOnly Then
'                                            RemoveBreakpoint lFindBreakpointIndex
'                                        End If
'                                        bExit = True
'                                    End If
'                                Else
'                                    If mbiBreakpoints(lFindBreakpointIndex).BreakType = btOnceOnly Then
'                                        RemoveBreakpoint lFindBreakpointIndex
'                                    End If
'                                    bExit = True
'                                End If
'                            End If
'                        Next
                    Loop Until bExit Or ((Processor6502.StopReason <> srDebugRun Or gyMem(Processor6502.PC) = &H60 Or gyMem(Processor6502.PC) = &H40 Or gyMem(Processor6502.PC) = &H6C) And Processor6502.S >= lStack)
                    If Processor6502.StopReason = srDebugRun Then
                        Processor6502.Execute
                        Me.Show
                        InstructionLocation = Processor6502.PC
                        UpdateMonitor
                        Processor6502.StopReason = srDebugBreak
                    Else
                        ProcessorStopped
                    End If
            End Select

        Case 116 'F5
            Keyboard.ClearPressedKeys
            Me.Hide
            DoDebugRun

    End Select
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lFindBreakpointIndex As Long
    Dim lAddress As Long
    Dim bBreakpointFound  As Boolean
    
    mlSelectedBoxIndex = FindBox(X, Y)
    
    mlWatchCursorAddress = -1
    
    If mlSelectedBoxIndex > -1 Then
        Select Case mbiBoxes(mlSelectedBoxIndex).Tag
            Case 0 ' Trace
                With mbiBoxes(mlSelectedBoxIndex)
                    If .Column = 0 Then
                        lAddress = mlTraceAddresses(.Row)
                        If lAddress > &H8000& And lAddress < &HBFFF& Then
                            AddWatch "(PC=" & HexNum(lAddress, 4) & "H)&(ROM=" & HexNum(RomSelect.SelectedBank, 1) & "H)", 2, 3, BreakTypes.Break
                        Else
                            AddWatch "PC=" & HexNum(lAddress, 4) & "H", 2, 3, BreakTypes.Break
                        End If
                    Else
                        mlWatchCursorAddress = mlTraceAddresses(.Row)
                        mlWatchCursorRow = .Row
                    End If
                End With
                UpdateMonitor False
            Case 1 ' Watch
                Select Case mbiBoxes(mlSelectedBoxIndex).Column
                    Case 0 ' Break Switch
                        mwiWatchLocations(mbiBoxes(mlSelectedBoxIndex).Row).Break = (mwiWatchLocations(mbiBoxes(mlSelectedBoxIndex).Row).Break + 1) Mod 3
                        Cls
                        UpdateMonitor False
                    Case 1 ' Expression
                        txtExpressionEdit.Top = mbiBoxes(mlSelectedBoxIndex).Row * WatchWindowLineHeight + WatchWindowMarginTop
                        txtExpressionEdit.Text = mwiWatchLocations(mbiBoxes(mlSelectedBoxIndex).Row).WatchExpressionText
                        txtExpressionEdit.Visible = True
                        mlSelectedWatchRow = mbiBoxes(mlSelectedBoxIndex).Row
                    Case 2 ' Width Switch
                        mwiWatchLocations(mbiBoxes(mlSelectedBoxIndex).Row).WidthBytes = (mwiWatchLocations(mbiBoxes(mlSelectedBoxIndex).Row).WidthBytes Mod 4) + 1
                        Cls
                        UpdateMonitor False
                    Case 3 ' Base Switch
                        mwiWatchLocations(mbiBoxes(mlSelectedBoxIndex).Row).Base = (mwiWatchLocations(mbiBoxes(mlSelectedBoxIndex).Row).Base + 1) Mod 4
                        Cls
                        UpdateMonitor False
                    Case 5 ' Delete watch
                        RemoveWatch mbiBoxes(mlSelectedBoxIndex).Row
                        Cls
                        UpdateMonitor False
                        Button = 0
                        mlLastBoxIndex = -1
                End Select
        End Select
        If Button = vbLeftButton Then
            RenderBox mlSelectedBoxIndex, mbiBoxes(mlSelectedBoxIndex).Centred, True
        End If
    End If
End Sub

Private Sub Form_Paint()
    UpdateMonitor False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Cancel = True
    Keyboard.ClearPressedKeys
    If Processor6502.StopReason = srDebugBreak Then
        Processor6502.StopReason = srDebugRun
        Controller.ProcessorStopped
    Else
        Processor6502.StopReason = srDebugRun
    End If
End Sub

Private Sub scrMem_Change()
    If Not mbNoUpdate Then
        mlTraceTopLocation = CLng(scrMem.Value) * 2&
        UpdateMonitor False
    End If
End Sub

Private Sub scrMem_Scroll()
    mlTraceTopLocation = CLng(scrMem.Value) * 2&
    UpdateMonitor False
End Sub

Private Function AddNumber(ByVal lX As Long, ByVal lY As Long, ByVal lNumber As Long, ByVal lBase As Long, ByVal lNumberOfDigits As Long, ByVal lTag As Long, ByVal lColumn As Long, ByVal lRow As Long) As Long
    Dim sDigits As String
    Dim lCharIndex As Long
    
    sDigits = BaseNum(lNumber, lNumberOfDigits, lBase)
    
    For lCharIndex = 1 To Len(sDigits)
        AddNumber = AddBox(lX, lY, DigitBoxWidth, DigitBoxHeight, vbWhite, vbBlack, Mid$(sDigits, lCharIndex, 1), lTag, lColumn + (lCharIndex - 1) * 16, lRow, True)
        lX = lX + DigitBoxWidth
    Next
End Function

Private Function AddMaskNumber(ByVal lX As Long, ByVal lY As Long, ByVal lNumber As Long, ByVal lMask As Long, ByVal lBase As Long, ByVal lNumberOfDigits As Long, ByVal lTag As Long, ByVal lColumn As Long, ByVal lRow As Long) As Long
    Dim sDigits As String
    Dim lCharIndex As Long
    Dim sDigitsMask As String
    Dim sChar As String
    
    sDigits = BaseNum(lNumber, lNumberOfDigits, lBase)
    sDigitsMask = BaseNum(lMask, lNumberOfDigits, lBase)
    
    For lCharIndex = 1 To Len(sDigits)
        sChar = IIf(Mid$(sDigitsMask, lCharIndex, 1) = "1", "X", Mid$(sDigits, lCharIndex, 1))
        
        If sChar = "X" Then
            AddMaskNumber = AddBox(lX, lY, DigitBoxWidth, DigitBoxHeight, mlDisabledBackColour, mlDisabledForeColour, sChar, lTag, lColumn + (lCharIndex - 1) * 16, lRow)
        Else
            AddMaskNumber = AddBox(lX, lY, DigitBoxWidth, DigitBoxHeight, vbWhite, vbBlack, sChar, lTag, lColumn + (lCharIndex - 1) * 16, lRow)
        End If
        lX = lX + DigitBoxWidth
    Next
End Function

Private Function AddLabelledNumber(ByVal lX As Long, ByVal lY As Long, ByVal lNumber As Long, ByVal lBase As Long, ByVal lNumberOfDigits As Long, ByVal sMask As String, ByVal lTag As Long, ByVal lColumn As Long, ByVal lRow As Long) As Long
    Dim sDigits As String
    Dim lCharIndex As Long
    Dim sMaskChar As String
    Dim lBackColour As Long
    Dim lForeColour As Long
    
    sDigits = BaseNum(lNumber, lNumberOfDigits, lBase)
    
    For lCharIndex = 1 To Len(sDigits)
        sMaskChar = Mid$(sMask, lCharIndex, 1)
        If Mid$(sDigits, lCharIndex, 1) = "0" Then
            lBackColour = RGB(200, 200, 200)
            lForeColour = RGB(120, 120, 120)
        Else
            lBackColour = vbWhite
            lForeColour = vbBlack
        End If
        
        AddLabelledNumber = AddBox(lX, lY, DigitBoxWidth, DigitBoxHeight, lBackColour, lForeColour, sMaskChar, lTag, lColumn + (lCharIndex - 1) * 16, lRow)
        lX = lX + DigitBoxWidth
    Next
End Function

Public Sub DoBreak()
    Console.DebugOn = True
    Me.Show

    mlTraceTopLocation = Processor6502.PC
    mlTraceInstructionLocation = Processor6502.PC
    UpdateMonitor True
End Sub

Public Sub DoDebugRun()
    Dim bExit As Boolean
    Dim lWatchIndex As Long
    Dim lFindBreakpointIndex As Long
    Dim lBreakIndex As Long
    
    Console.DebugOn = True
    
    Processor6502.StopReason = srDebugRun
    Keyboard.ClearPressedKeys
    Me.Hide
    bExit = False
    Do
        Memory.RecentlyRead = -1
        Memory.RecentlyWritten = -1
        mlPreviousPC = Processor6502.PC
        Processor6502.Execute
        
        If Memory.RecentlyRead <> -1 Then
            For lWatchIndex = 0 To mlTotalWatchLocations - 1
                If mwiWatchLocations(lWatchIndex).ExpressionOk Then
                    For lBreakIndex = 0 To mwiWatchLocations(lWatchIndex).WatchExpression.ReadsCount - 1
                        If mwiWatchLocations(lWatchIndex).WatchExpression.Reads(lBreakIndex).Device = 0 Then
                            If mwiWatchLocations(lWatchIndex).WatchExpression.Reads(lBreakIndex).Value = Memory.RecentlyRead Then
                                If mwiWatchLocations(lWatchIndex).Break = None Then
                                    bExit = True
                                    mlBreakWatchRow = lWatchIndex
                                Else
                                    If EvaluateExpression(lWatchIndex) <> 0 Then
                                        If mwiWatchLocations(lWatchIndex).Break = BreakOnce Then
                                            mwiWatchLocations(lWatchIndex).Break = None
                                        End If
                                        bExit = True
                                        mlBreakWatchRow = lWatchIndex
                                    End If
                                End If
                            End If
                        End If
                    Next
                End If
            Next
        End If


        If Memory.RecentlyWritten <> -1 Then
            For lWatchIndex = 0 To mlTotalWatchLocations - 1
                If mwiWatchLocations(lWatchIndex).ExpressionOk Then
                    For lBreakIndex = 0 To mwiWatchLocations(lWatchIndex).WatchExpression.WritesCount - 1
                        If mwiWatchLocations(lWatchIndex).WatchExpression.Writes(lBreakIndex).Device = 0 Then
                            If mwiWatchLocations(lWatchIndex).WatchExpression.Writes(lBreakIndex).Value = Memory.RecentlyWritten Then
                                If mwiWatchLocations(lWatchIndex).Break = None Then
                                    bExit = True
                                    mlBreakWatchRow = lWatchIndex
                                Else
                                    If EvaluateExpression(lWatchIndex) <> 0 Then
                                        If mwiWatchLocations(lWatchIndex).Break = BreakOnce Then
                                            mwiWatchLocations(lWatchIndex).Break = None
                                        End If
                                        bExit = True
                                        mlBreakWatchRow = lWatchIndex
                                    End If
                                End If
                            End If
                        End If
                    Next
                End If
            Next
        End If
        
        
        For lWatchIndex = 0 To mlTotalWatchLocations - 1
            If mwiWatchLocations(lWatchIndex).Break > None Then
                If mwiWatchLocations(lWatchIndex).ExpressionOk Then
                    If mwiWatchLocations(lWatchIndex).WatchExpression.ReadsCount = 0 And mwiWatchLocations(lWatchIndex).WatchExpression.WritesCount = 0 Then
                        If EvaluateExpression(lWatchIndex) <> 0 Then
                            If mwiWatchLocations(lWatchIndex).Break = BreakOnce Then
                                mwiWatchLocations(lWatchIndex).Break = None
                            End If
                            bExit = True
                            mlBreakWatchRow = lWatchIndex
                        End If
                    End If
                End If
            End If
        Next
        

    Loop Until bExit Or Processor6502.StopReason <> srDebugRun
    If Processor6502.StopReason = srDebugRun Then
        Processor6502.StopReason = srDebugBreak
        Me.Show
        InstructionLocation = Processor6502.PC
        UpdateMonitor
    Else
        ProcessorStopped
    End If
End Sub


Private Sub InitialiseCompiler()
    Dim oParser As SaffronClasses.ISaffronObject
    Dim SaffronCompiler As Object
    Dim sRules As String
    
    sRules = Space$(FileLen(App.path & "\debug.saf"))
    
    Open App.path & "\debug.saf" For Binary As #1
    Get #1, , sRules
    Close #1
    
    If Not SaffronObject.CreateRules(sRules) Then
        MsgBox "Bad Saf"
        End
    End If
    Set moExpression = SaffronObject.Rules("expression")
    
'    Dim oResult As New SaffronTree
'    SaffronStream.Text = "PC"
'    If moExpression.Parse(oResult) Then
'        Stop
'    Else
'        Stop
'    End If
End Sub

Private Function ParseUserExpression(ByVal sExpression As String) As Boolean
    Dim oTree As SaffronTree
    
    mexpTempExpression.ConstantsCount = 0
    mexpTempExpression.OperationsCount = 0
    mexpTempExpression.ReadsCount = 0
    mexpTempExpression.WritesCount = 0
    
    Erase mexpTempExpression.Constants
    Erase mexpTempExpression.Operations
    Erase mexpTempExpression.Reads
    Erase mexpTempExpression.Writes
    
    SaffronClasses.SaffronStream.Text = sExpression

    
    Set oTree = New SaffronClasses.SaffronTree
    If Not moExpression.Parse(oTree) Then
        Exit Function
    End If
    
    ParseExpression oTree, 0
    SortOperationsByLevel
    ParseUserExpression = True
End Function

Private Sub ParseExpression(oTree As SaffronTree, ByVal lLevel As Long)
    Dim lIndex As Long
    Dim oSub As SaffronTree
    Dim lPreviousReultPos As Long
    
    For lIndex = 1 To oTree.SubTree.Count
        Set oSub = oTree(lIndex)
        If (lIndex Mod 2) = 1 Then
            If lIndex = 1 Then
                lPreviousReultPos = mexpTempExpression.OperationsCount
            End If
            ParseTerm oSub, lLevel + 1
        Else
            ReDim Preserve mexpTempExpression.Operations(mexpTempExpression.OperationsCount)
            With mexpTempExpression.Operations(mexpTempExpression.OperationsCount)
                .OperationIndex = InStr("+-*/=<>&|^", oSub.Text)
                .ResultPos = lPreviousReultPos
                .OperandPos = mexpTempExpression.ConstantsCount
                .Level = lLevel
            End With
            mexpTempExpression.OperationsCount = mexpTempExpression.OperationsCount + 1
        End If
    Next
End Sub

Private Sub ParseTerm(oTerm As SaffronTree, ByVal lLevel As Long)
    Dim lConstant As Long
    Dim bBreakOnRead As Boolean
    Dim bBreakOnWrite As Boolean
    Dim bLabelOk As Boolean
    Dim pLabel As Pair
    
    Select Case oTerm.Index
        Case 1 ' Break Number
            bBreakOnRead = oTerm(1)(1).Index = 1
            bBreakOnWrite = oTerm(1)(2).Index = 1
            Select Case oTerm(1)(3).Index
                Case 1 ' binary
                    lConstant = ConvertBase(UCase$(oTerm(1)(3)(1)(1).Text), 2)
                Case 2 ' decimal
                    lConstant = ConvertBase(UCase$(oTerm(1)(3)(1)(1).Text), 10)
                Case 3 ' hex
                    lConstant = ConvertBase(UCase$(oTerm(1)(3)(1)(1).Text), 16)
            End Select
            ReDim Preserve mexpTempExpression.Constants(mexpTempExpression.ConstantsCount)
            mexpTempExpression.Constants(mexpTempExpression.ConstantsCount).Value = lConstant
            mexpTempExpression.Constants(mexpTempExpression.ConstantsCount).Device = 0
            mexpTempExpression.ConstantsCount = mexpTempExpression.ConstantsCount + 1
            If bBreakOnRead Then
                ReDim Preserve mexpTempExpression.Reads(mexpTempExpression.ReadsCount)
                mexpTempExpression.Reads(mexpTempExpression.ReadsCount).Value = lConstant
                mexpTempExpression.Reads(mexpTempExpression.ReadsCount).Device = 0
                mexpTempExpression.ReadsCount = mexpTempExpression.ReadsCount + 1
            End If
            If bBreakOnWrite Then
                ReDim Preserve mexpTempExpression.Writes(mexpTempExpression.WritesCount)
                mexpTempExpression.Writes(mexpTempExpression.WritesCount).Value = lConstant
                mexpTempExpression.Writes(mexpTempExpression.WritesCount).Device = 0
                mexpTempExpression.WritesCount = mexpTempExpression.WritesCount + 1
            End If
        Case 2 ' Bracketed
            ParseExpression oTerm(1)(1), lLevel + 1
        Case 3 ' Label
            bLabelOk = False
'            bBreakOnRead = oTerm(1)(1).Index = 1
'            bBreakOnWrite = oTerm(1)(2).Index = 1
            Select Case UCase$(oTerm.Text)
                Case "A"
                    pLabel.Device = 1
                    pLabel.Value = 0
                    bLabelOk = True
                Case "X"
                    pLabel.Device = 1
                    pLabel.Value = 1
                    bLabelOk = True
                Case "Y"
                    pLabel.Device = 1
                    pLabel.Value = 2
                    bLabelOk = True
                Case "S"
                    pLabel.Device = 1
                    pLabel.Value = 3
                    bLabelOk = True
                Case "P"
                    pLabel.Device = 1
                    pLabel.Value = 4
                    bLabelOk = True
                Case "PC"
                    pLabel.Device = 1
                    pLabel.Value = 5
                    bLabelOk = True
                Case "PPC"
                    pLabel.Device = 1
                    pLabel.Value = 13
                    bLabelOk = True
                Case "N"
                    pLabel.Device = 1
                    pLabel.Value = 6
                    bLabelOk = True
                Case "V"
                    pLabel.Device = 1
                    pLabel.Value = 7
                    bLabelOk = True
                Case "B"
                    pLabel.Device = 1
                    pLabel.Value = 8
                    bLabelOk = True
                Case "D"
                    pLabel.Device = 1
                    pLabel.Value = 9
                    bLabelOk = True
                Case "I"
                    pLabel.Device = 1
                    pLabel.Value = 10
                    bLabelOk = True
                Case "Z"
                    pLabel.Device = 1
                    pLabel.Value = 11
                    bLabelOk = True
                Case "C"
                    pLabel.Device = 1
                    pLabel.Value = 12
                    bLabelOk = True
                Case "ROM"
                    pLabel.Device = 2
                    pLabel.Value = 0
                    bLabelOk = True
            End Select
            If bLabelOk Then
                ReDim Preserve mexpTempExpression.Constants(mexpTempExpression.ConstantsCount)
                mexpTempExpression.Constants(mexpTempExpression.ConstantsCount) = pLabel
                mexpTempExpression.ConstantsCount = mexpTempExpression.ConstantsCount + 1
            End If
        Case 4 ' Unary
            ParseTerm oTerm(1)(2), lLevel + 1
            ReDim Preserve mexpTempExpression.Operations(mexpTempExpression.OperationsCount)
            With mexpTempExpression.Operations(mexpTempExpression.OperationsCount)
                .OperationIndex = (InStr("?x!x%x??!!%%", oTerm(1)(1).Text) - 1) \ 2 + 1 + 10
                .ResultPos = mexpTempExpression.ConstantsCount - 1
                .Level = lLevel
            End With
            mexpTempExpression.OperationsCount = mexpTempExpression.OperationsCount + 1
    End Select
End Sub

Private Sub SortOperationsByLevel()
    Dim bSorted As Boolean
    Dim lIndex As Long
    Dim lTempOperation As Operation
    
    While Not bSorted
        bSorted = True
        With mexpTempExpression
            For lIndex = 0 To .OperationsCount - 2
                If .Operations(lIndex).Level < .Operations(lIndex + 1).Level Then
                    lTempOperation = .Operations(lIndex)
                    .Operations(lIndex) = .Operations(lIndex + 1)
                    .Operations(lIndex + 1) = lTempOperation
                    bSorted = False
                End If
            Next
        End With
    Wend
End Sub

Private Function EvaluateExpression(ByVal lWatchLocation As Long) As Long
    Dim lOperationIndex As Long
    Dim lValues() As Pair
    Dim lOperandValue As Long
    Dim lResultValue As Long
    Dim lConstantIndex As Long
    
    If mwiWatchLocations(lWatchLocation).WatchExpression.ConstantsCount = 0 Then
        Exit Function
    End If
    
    ReDim lValues(mwiWatchLocations(lWatchLocation).WatchExpression.ConstantsCount - 1)
    CopyMemory lValues(0), mwiWatchLocations(lWatchLocation).WatchExpression.Constants(0), 8 * UBound(lValues) + 8
    
    For lConstantIndex = 0 To mwiWatchLocations(lWatchLocation).WatchExpression.ConstantsCount - 1
        Select Case lValues(lConstantIndex).Device
            Case 0
            Case 1 ' Processor
                Select Case lValues(lConstantIndex).Value
                    Case 0 ' A
                        lValues(lConstantIndex).Value = Processor6502.A
                    Case 1 ' X
                        lValues(lConstantIndex).Value = Processor6502.X
                    Case 2 ' Y
                        lValues(lConstantIndex).Value = Processor6502.Y
                    Case 3 ' S
                        lValues(lConstantIndex).Value = Processor6502.S And &HFF&
                    Case 4 ' P
                        lValues(lConstantIndex).Value = Processor6502.N + Processor6502.V + Processor6502.B + Processor6502.D + Processor6502.I + Processor6502.Z + Processor6502.C
                    Case 5 ' PC
                        lValues(lConstantIndex).Value = Processor6502.PC
                    Case 6 ' N
                        lValues(lConstantIndex).Value = Processor6502.N \ 128&
                    Case 7 ' V
                        lValues(lConstantIndex).Value = Processor6502.V \ 64&
                    Case 8 ' B
                        lValues(lConstantIndex).Value = Processor6502.B \ 16&
                    Case 9 ' D
                        lValues(lConstantIndex).Value = Processor6502.D \ 8&
                    Case 10 ' I
                        lValues(lConstantIndex).Value = Processor6502.I \ 4&
                    Case 11 ' Z
                        lValues(lConstantIndex).Value = Processor6502.Z \ 2&
                    Case 12 ' C
                        lValues(lConstantIndex).Value = Processor6502.C
                    Case 13 ' PPC Previous PC value
                        lValues(lConstantIndex).Value = mlPreviousPC
                End Select
            Case 2 ' ROM SELECT
                Select Case lValues(lConstantIndex).Value
                    Case 0 ' ROM select
                        lValues(lConstantIndex).Value = RomSelect.SelectedBank
                End Select
        End Select
    Next
    
    For lOperationIndex = 0 To mwiWatchLocations(lWatchLocation).WatchExpression.OperationsCount - 1
        With mwiWatchLocations(lWatchLocation).WatchExpression.Operations(lOperationIndex)
            lResultValue = lValues(.ResultPos).Value
            lOperandValue = lValues(.OperandPos).Value

            Select Case .OperationIndex
                Case 1 ' +
                    lValues(.ResultPos).Value = lResultValue + lOperandValue
                Case 2 ' -
                    lValues(.ResultPos).Value = lResultValue - lOperandValue
                Case 3 ' *
                    lValues(.ResultPos).Value = lResultValue * lOperandValue
                Case 4 ' /
                    lValues(.ResultPos).Value = lResultValue / lOperandValue
                Case 5 ' =
                    lValues(.ResultPos).Value = Abs(lResultValue = lOperandValue)
                Case 6 ' <
                    lValues(.ResultPos).Value = Abs(lResultValue < lOperandValue)
                Case 7 ' >
                    lValues(.ResultPos).Value = Abs(lResultValue > lOperandValue)
                Case 8 ' &
                    lValues(.ResultPos).Value = lResultValue And lOperandValue
                Case 9 ' |
                    lValues(.ResultPos).Value = lResultValue Or lOperandValue
                Case 10 ' ^
                    lValues(.ResultPos).Value = lResultValue Xor lOperandValue
                Case 11 ' ?
                    lValues(.ResultPos).Value = gyMem(lResultValue)
                Case 12 ' !
                    lValues(.ResultPos).Value = gyMem(lResultValue And &HFFFF&) + gyMem((lResultValue + 1) And &HFFFF&) * 256
                Case 13 ' %
                    lValues(.ResultPos).Value = gyMem(lResultValue)
                Case 14 ' ??
                    lValues(.ResultPos).Value = gyMem(lResultValue) + 256 * (gyMem(lResultValue) > 127)
                Case 15 ' !!
                    lValues(.ResultPos).Value = gyMem(lResultValue)
                Case 16 ' %%
                    lValues(.ResultPos).Value = gyMem(lResultValue)

            End Select
        End With
    Next
    EvaluateExpression = lValues(0).Value
End Function

Private Sub txtExpressionEdit_Change()
    txtExpressionEdit.ForeColor = vbBlack
End Sub

Private Sub txtExpressionEdit_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If ParseUserExpression(txtExpressionEdit.Text) Then
            txtExpressionEdit.ForeColor = vbBlack
            txtExpressionEdit.Visible = False
            With mwiWatchLocations(mlSelectedWatchRow)
                .ExpressionOk = True
                .WatchExpressionText = txtExpressionEdit.Text
                .WatchExpression = mexpTempExpression
            End With
            ShowWatchWindow
        Else
            txtExpressionEdit.ForeColor = vbRed
            With mwiWatchLocations(mlSelectedWatchRow)
                .ExpressionOk = False
                .WatchExpressionText = txtExpressionEdit.Text
            End With
        End If
    End If
End Sub

Public Sub AddMemory(ByVal lAddress As Long)
    AddWatch "W" & HexNum(lAddress, 4) & "h", 1, 3, None
    UpdateMonitor
End Sub


