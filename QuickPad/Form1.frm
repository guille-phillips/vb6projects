VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Quick Pad"
   ClientHeight    =   4005
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4590
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   267
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   306
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtList 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   840
      MultiLine       =   -1  'True
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   3360
      Width           =   3495
   End
   Begin VB.TextBox txtReplace 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   840
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   3000
      Width           =   3495
   End
   Begin VB.TextBox txtFind 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   840
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   2640
      Width           =   3495
   End
   Begin VB.TextBox txtText 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   0
      MultiLine       =   -1  'True
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   150
      Width           =   4335
   End
   Begin VB.Label lblList 
      Caption         =   "List:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   3360
      Width           =   735
   End
   Begin VB.Label lblReplace 
      Caption         =   "Replace:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   105
      TabIndex        =   4
      Top             =   3030
      Width           =   735
   End
   Begin VB.Label lblFind 
      Caption         =   "Find:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   105
      TabIndex        =   3
      Top             =   2670
      Width           =   495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function ShowScrollBar Lib "user32" (ByVal hwnd As Long, ByVal wBar As Long, ByVal bShow As Long) As Long

Const SB_HORZ = 0
Const SB_VERT = 1
Const SB_CTL = 2
Const SB_BOTH = 3

Private Const EM_SETTABSTOPS = &HCB
Private mlTabStops(100) As Long
Private mlTabSize As Long

Private mnBoxWidth As Single
Private mnTextHeight As Single
Private mnFindTop As Single
Private mnReplaceTop As Single
Private mnListTop As Single

Private moReplaceLex As ISaffronObject
Private moSortLex As ISaffronObject
Private moSetLex As ISaffronObject

Private msChanges As Variant

Private miShift As Integer
Private mlDragStop As Long

Private Sub Form_Load()
    Dim sTabStops As String
    
    mnBoxWidth = Me.Width / Screen.TwipsPerPixelX - txtFind.Width
    mnTextHeight = Me.Height / Screen.TwipsPerPixelY - txtText.Height
    mnFindTop = Me.Height / Screen.TwipsPerPixelY - txtFind.Top
    mnReplaceTop = Me.Height / Screen.TwipsPerPixelY - txtReplace.Top
    mnListTop = Me.Height / Screen.TwipsPerPixelY - txtList.Top
    
    msChanges = Array()
    InitialiseParser
    
    Me.Width = GetSetting("QuickPad", "Dimensions", "Width", Me.Width)
    Me.Height = GetSetting("QuickPad", "Dimensions", "Height", Me.Height)
    
    mlTabSize = GetSetting("QuickPad", "Tabs", "Size", 30)
    sTabStops = GetSetting("QuickPad", "Tabs", "TabStops", "")
    
    If sTabStops <> "" Then
        DecodeTabStops sTabStops
    Else
        UpdateTabs mlTabSize
    End If
    
    ShowScrollBar txtText.hwnd, SB_VERT, 1
    ShowTabStops
    mlDragStop = -1
End Sub

Private Sub UpdateTabs(ByVal lSize As Long)
    Dim lTabIndex As Long
    
    If lSize <> 0 Then
        For lTabIndex = 0 To UBound(mlTabStops)
            mlTabStops(lTabIndex) = lTabIndex * lSize
        Next
    End If
    
    SortTabs
    ShowTabStops
    SetTextBoxTabs
End Sub

Private Sub SortTabs()
    Dim lIndex As Long
    Dim bFinished As Boolean
    Dim lTemp As Long
    
    While Not bFinished
        bFinished = True
        For lIndex = 0 To UBound(mlTabStops) - 1
            If mlTabStops(lIndex) > mlTabStops(lIndex + 1) Then
                lTemp = mlTabStops(lIndex + 1)
                mlTabStops(lIndex) = mlTabStops(lIndex + 1)
                mlTabStops(lIndex + 1) = lTemp
                bFinished = False
            End If
        Next
    Wend
End Sub

Private Sub ShowTabStops()
    Dim lIndex As Long
    
    Me.Line (0, 0)-Step(Me.Width, 10), vbWhite, BF
    For lIndex = 0 To UBound(mlTabStops)
        Me.CurrentX = mlTabStops(lIndex) * 1.754
        Me.CurrentY = 0
        Me.Line (mlTabStops(lIndex) * 1.754, 0)-Step(0, 10)
    Next
End Sub

Private Sub SetTextBoxTabs()
    SendMessage txtText.hwnd, EM_SETTABSTOPS, UBound(mlTabStops) + 1, mlTabStops(0)
End Sub

Private Function Compare(ByVal sString1 As String, ByVal sString2 As String, ByVal bIntelligentCompare As Boolean) As Boolean
    Dim oString1 As New SaffronTree
    Dim oString2 As New SaffronTree
    Dim bSame As Boolean
    Dim lIndex As Long
    Dim sText1 As String
    Dim sText2 As String
    
    If Not bIntelligentCompare Then
        If sString1 > sString2 Then
            Compare = True
        End If
        Exit Function
    End If
    
    SaffronStream.Text = sString1
    moSortLex.Parse oString1
    
    SaffronStream.Text = sString2
    moSortLex.Parse oString2
     
    bSame = True
    lIndex = 1
    
    Do
        If lIndex > oString1.SubTree.Count And lIndex <= oString2.SubTree.Count Then
            Compare = False
            Exit Function
        ElseIf lIndex > oString2.SubTree.Count And lIndex <= oString1.SubTree.Count Then
            Compare = True
            Exit Function
        ElseIf lIndex > oString1.SubTree.Count And lIndex > oString2.SubTree.Count Then
            Compare = False
            Exit Function
        End If
        
        SaffronStream.Text = sString1
        sText1 = oString1.SubTree(lIndex).Text
        SaffronStream.Text = sString2
        sText2 = oString2.SubTree(lIndex).Text
        
        If oString1.SubTree(lIndex).Index = 1 And oString2.SubTree(lIndex).Index = 1 Then
            If Val(sText1) > Val(sText2) Then
                Compare = True
                Exit Function
            ElseIf Val(sText1) < Val(sText2) Then
                Compare = False
                Exit Function
            End If
        Else
            If sText1 > sText2 Then
                Compare = True
                Exit Function
            ElseIf sText1 < sText2 Then
                Compare = False
                Exit Function
            End If
        End If
        lIndex = lIndex + 1
    Loop
End Function

Private Sub InitialiseParser()
    Dim sDef As String
    Dim sPath As String
    
    sPath = App.Path & "\quickpad.saf"
    sDef = Space$(FileLen(sPath))
    Open sPath For Binary As #1
    Get #1, , sDef
    Close #1
    
    If Not CreateRules(sDef) Then
        MsgBox "Bad Def"
        End
    End If
    Set moReplaceLex = Rules("text")
    Set moSortLex = Rules("string")
    Set moSetLex = Rules("set")
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lIndex As Long
    
    If Button = vbLeftButton Then
        If Y < 8 Then
            X = X / 1.754
            For lIndex = 0 To UBound(mlTabStops)
                If Abs(X - mlTabStops(lIndex)) < 10 Then
                    mlDragStop = lIndex
                    Exit For
                End If
            Next
        End If
    End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If mlDragStop > -1 Then
        If Button = vbLeftButton Then
            mlTabStops(mlDragStop) = X / 1.754
            SetTextBoxTabs
            ShowTabStops
        Else
            mlDragStop = -1
        End If
    End If
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    mlDragStop = -1
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    txtText.Width = Me.Width / Screen.TwipsPerPixelX - 10
    txtText.Height = Me.Height / Screen.TwipsPerPixelY - mnTextHeight
    txtFind.Width = Me.Width / Screen.TwipsPerPixelX - mnBoxWidth
    txtReplace.Width = Me.Width / Screen.TwipsPerPixelX - mnBoxWidth
    txtList.Width = Me.Width / Screen.TwipsPerPixelX - mnBoxWidth
    txtFind.Top = Me.Height / Screen.TwipsPerPixelY - mnFindTop
    txtReplace.Top = Me.Height / Screen.TwipsPerPixelY - mnReplaceTop
    txtList.Top = Me.Height / Screen.TwipsPerPixelY - mnListTop
    lblFind.Top = Me.Height / Screen.TwipsPerPixelY - mnFindTop + 3
    lblReplace.Top = Me.Height / Screen.TwipsPerPixelY - mnReplaceTop + 3
    lblList.Top = Me.Height / Screen.TwipsPerPixelY - mnListTop + 3
    
    ShowTabStops
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveSetting "QuickPad", "Dimensions", "Width", Me.Width
    SaveSetting "QuickPad", "Dimensions", "height", Me.Height
    SaveSetting "QuickPad", "Tabs", "Size", mlTabSize
    SaveSetting "QuickPad", "Tabs", "TabStops", EncodeTabStops
End Sub

Private Function EncodeTabStops() As String
    Dim vStops As Variant
    Dim lIndex As Long
    
    vStops = Array()
    ReDim vStops(UBound(mlTabStops))
    For lIndex = 0 To UBound(mlTabStops)
        vStops(lIndex) = CStr(mlTabStops(lIndex))
    Next
    EncodeTabStops = Join(vStops, "|")
End Function

Private Function DecodeTabStops(ByVal sStops As String)
    Dim vStops As Variant
    Dim lIndex As Long
    
    vStops = Split(sStops, "|")
    For lIndex = 0 To UBound(vStops)
        mlTabStops(lIndex) = vStops(lIndex)
    Next
    SetTextBoxTabs
End Function

Private Sub txtFind_GotFocus()
    txtText.TabStop = False
    txtFind.TabStop = True
    txtReplace.TabStop = True
    txtList.TabStop = False
End Sub

Private Sub txtList_Change()
    txtList.ForeColor = vbBlack
End Sub

Private Sub txtText_GotFocus()
    txtText.TabStop = False
    txtFind.TabStop = False
    txtReplace.TabStop = False
    txtList.TabStop = False
End Sub

Private Sub txtFind_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtReplace.SetFocus
    End If
End Sub

Private Sub txtList_KeyPress(iKeyAscii As Integer)
    Dim oSaffronTree As SaffronTree
    Dim oSetDef As SetDef
    
    If iKeyAscii = 10 Then
        iKeyAscii = 0
        SaffronStream.Text = "{" & Replace$(txtList.Text, vbCrLf, "") & "}"
        Set oSaffronTree = New SaffronTree
        If moSetLex.Parse(oSaffronTree) Then
            Set goAllSets = New Dictionary
            Set oSetDef = New SetDef
            oSetDef.CreateSet oSaffronTree
            txtText.Text = oSetDef.AllText
        Else
            txtList.ForeColor = vbRed
        End If
    End If
End Sub

Private Sub txtReplace_KeyPress(iKeyAscii As Integer)
    If iKeyAscii = 13 Then
        TextReplace
    End If
End Sub

Private Sub TextReplace()
    Dim sFind As String
    Dim sReplace As String
    Dim sLeft As String
    Dim sMiddle As String
    Dim sRight As String
    Dim iSelStart As Long
    Dim iSelLength As Long
    
    iSelStart = txtText.SelStart
    iSelLength = txtText.SelLength
    
    sFind = Decode(txtFind.Text)
    sReplace = Decode(txtReplace.Text)

    sLeft = Left$(txtText.Text, txtText.SelStart)
    sMiddle = txtText.SelText
    sRight = Mid$(txtText.Text, txtText.SelStart + txtText.SelLength + 1)
    
    If iSelLength <> 0 Then
        txtText.Text = sLeft & Replace(sMiddle, sFind, sReplace) & sRight
    Else
        txtText.Text = Replace(sLeft & sRight, sFind, sReplace)
    End If
    
    txtText.SelStart = iSelStart
    txtText.SelLength = iSelLength
End Sub

Private Function Decode(ByVal sCode As String) As String
    Dim oTree As SaffronTree
    Dim oSub As SaffronTree
    
    SaffronStream.Text = sCode
    Set oTree = New SaffronTree
    If moReplaceLex.Parse(oTree) Then
        For Each oSub In oTree.SubTree
            Select Case oSub.Index
                Case 1
                    Decode = Decode & "#"
                Case 2 ' #t
                    Decode = Decode & vbTab
                Case 3 ' #n
                    Decode = Decode & vbCrLf
                Case 4
                    Decode = Decode & Chr(Val(oSub.Text))
                Case 5
                    Decode = Decode & oSub.Text
            End Select
        Next
    Else
        Decode = sCode
    End If
End Function

Private Sub txtText_Change()
    If txtText.Tag = "" Then
        LogChange txtText.Text, txtText.SelStart, txtText.SelLength
    End If
End Sub

Private Function LogChange(sText As String, ByVal lStart As Long, ByVal lLength As Long)
    ReDim Preserve msChanges(UBound(msChanges) + 1)
    msChanges(UBound(msChanges)) = Array(sText, lStart, lLength)
End Function

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 10 Then
        TextReplace
        KeyAscii = 0
    End If
End Sub

Private Sub txtText_KeyPress(KeyAscii As Integer)
    If KeyAscii = 26 Then
        txtText.Tag = "x"
        If UBound(msChanges) > 0 Then
            ReDim Preserve msChanges(UBound(msChanges) - 1)
            txtText.Text = msChanges(UBound(msChanges))(0)
            txtText.SelStart = msChanges(UBound(msChanges))(1)
            txtText.SelLength = msChanges(UBound(msChanges))(2)
        ElseIf UBound(msChanges) = 0 Then
            txtText.Text = ""
            txtText.SelStart = 0
            txtText.SelLength = 0
            msChanges = Array()
        End If
        txtText.Tag = ""
        KeyAscii = 0
    ElseIf KeyAscii = 27 Then
        msChanges = Array()
        txtText.Text = ""
    End If
End Sub

Private Sub CompileList()
    Dim vLines As Variant
    Dim lLineCount As Long
    Dim sPart As String
    Dim vLine As Variant
    Dim lSize As Long
    Dim sTopLine As String
    Dim lPosition As Long
    Dim bInAll As Boolean
    Dim lLineIndex As Long
    Dim lFoundPosition As Long
    Dim lCharIndex As Long
    Dim vParts As Variant
    Dim bCovereds() As Boolean
    Dim bCovered As Boolean
    Dim vGaps As Variant
    Dim vGap As Variant
    
    vLines = Split(txtText.Text, vbCrLf)
    lLineCount = UBound(vLines) + 1
    
    If lLineCount <= 1 Then
        Exit Sub
    End If
    
    vParts = Array()
    
    sTopLine = vLines(0)
    ReDim bCovereds(Len(sTopLine))
    
    For lSize = Len(sTopLine) To 1 Step -1
        For lPosition = 1 To Len(sTopLine) - lSize + 1
            bCovered = False
            For lCharIndex = lPosition To lPosition + lSize - 1
                If bCovereds(lCharIndex) Then
                    bCovered = True
                    Exit For
                End If
            Next
            
            If Not bCovered Then
                sPart = Mid$(sTopLine, lPosition, lSize)
                bInAll = True
                For lLineIndex = 1 To UBound(vLines)
                    lFoundPosition = InStr(vLines(lLineIndex), sPart)
                    If lFoundPosition = 0 Then
                        bInAll = False
                        Exit For
                    End If
                Next
                If bInAll Then
                    For lCharIndex = lPosition To lPosition + lSize - 1
                        bCovereds(lCharIndex) = True
                    Next
                    ReDim Preserve vParts(UBound(vParts) + 1)
                    vParts(UBound(vParts)) = sPart
                End If
            End If
        Next
    Next
   
    vGaps = Array()
    ReDim vGaps(lLineCount - 1)
   
    For lLineIndex = 0 To lLineCount - 1
        ReDim vGap(Len(vLines(lLineIndex))) As Boolean
        vGaps(lLineIndex) = vGap
    Next
   
    AssignGaps vLines(0), vParts(0), vGaps(0)
   Debug.Print FindGap(vLines(0), "Keyword", vGaps(0))
End Sub

Private Sub AssignGaps(ByVal sText As String, ByVal sSubText As String, ByRef vGap As Variant)
    Dim lPos As Long
    Dim lLen As Long
    Dim lStart As Long
    Dim lIndex As Long
    Dim bOk As Boolean
    
    lPos = 1
    lStart = 1
    While Not bOk And lStart <= Len(sText) And lPos > 0
        lPos = InStr(lStart, sText, sSubText)
        lLen = Len(sSubText)
        
        If lPos > 0 Then
            bOk = True
            For lIndex = lPos To lPos + lLen - 1
                If vGap(lIndex) Then
                    lStart = lPos + lLen
                    bOk = False
                    Exit For
                End If
            Next
        End If
    Wend
    
    If lPos > 0 Then
        For lIndex = lPos To lPos + lLen - 1
            vGap(lIndex) = True
        Next
    End If
End Sub

Private Function FindGap(ByVal sText As String, ByVal sSubText As String, vGap As Variant) As Long
    Dim lPos As Long
    Dim lLen As Long
    Dim lStart As Long
    Dim lIndex As Long
    Dim bOk As Boolean
    Dim bPrevious As Boolean
    Dim lGapCount As Long
    Dim lGaps() As Long
    Dim lOffset As Long
    
    ReDim lGaps(UBound(vGap))
    
    bPrevious = False
    For lIndex = 0 To UBound(vGap)
        If vGap(lIndex) < bPrevious Then
            lGapCount = lGapCount + 1
        End If
        lGaps(lIndex) = lGapCount
        bPrevious = vGap(lIndex)
    Next
    lOffset = (lGapCount + 1) * lGapCount \ 2
    
    lPos = 1
    lStart = 1
    While Not bOk And lStart <= Len(sText) And lPos > 0
        lPos = InStr(lStart, sText, sSubText)
        lLen = Len(sSubText)
        
        If lPos > 0 Then
            bOk = True
            For lIndex = lPos To lPos + lLen - 1
                If vGap(lIndex) Then
                    lStart = lPos + lLen
                    bOk = False
                    Exit For
                End If
            Next
        End If
    Wend
    
    If lPos > 0 Then
        FindGap = lOffset + lGaps(lPos)
    End If
End Function

Private Function GetSelectedText() As String
    If txtText.SelLength = 0 Then
        GetSelectedText = txtText.Text
    Else
        GetSelectedText = Mid$(txtText.Text, txtText.SelStart, txtText.SelLength)
    End If
End Function

Private Function InsertTextIntoSelection(ByVal sInsert As String)
    Dim sLeft As String
    Dim sRight As String
    
    sLeft = Left$(txtText.Text, txtText.SelStart)
    sRight = Mid$(txtText.Text, txtText.SelStart + txtText.SelLength + 1)
    txtText.Text = sLeft & sInsert & sRight
    txtText.SelLength = Len(sInsert)
End Function

Private Sub txtText_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim vPosition As Variant
    Dim vTable As Variant
    Dim lTabIndex As Long
    
    If (Shift And 4) = 4 Then
        Select Case KeyCode
            Case 39 ' Column Right
                LogChange txtText.Text, txtText.SelStart, txtText.SelLength
                vPosition = FindPosition
                vTable = Pivot(PadTable(ConvertTextToTable(txtText.Text)))
                MoveRow vTable, vPosition(1), 1
                txtText.Text = ConvertTableToText(Pivot(vTable))
                MovePosition Array(vPosition(0), vPosition(1) + 1, vPosition(2))
            Case 37 ' Column Left
                LogChange txtText.Text, txtText.SelStart, txtText.SelLength
                vPosition = FindPosition
                vTable = Pivot(PadTable(ConvertTextToTable(txtText.Text)))
                MoveRow vTable, vPosition(1), -1
                txtText.Text = ConvertTableToText(Pivot(vTable))
                MovePosition Array(vPosition(0), vPosition(1) - 1, vPosition(2))
            Case 40 ' Row Down
                LogChange txtText.Text, txtText.SelStart, txtText.SelLength
                vPosition = FindPosition
                vTable = ConvertTextToTable(txtText.Text)
                MoveRow vTable, vPosition(0), 1
                txtText.Text = ConvertTableToText(vTable)
                MovePosition Array(vPosition(0) + 1, vPosition(1), vPosition(2))
            Case 38 ' Row Up
                LogChange txtText.Text, txtText.SelStart, txtText.SelLength
                vPosition = FindPosition
                vTable = ConvertTextToTable(txtText.Text)
                MoveRow vTable, vPosition(0), -1
                txtText.Text = ConvertTableToText(vTable)
                MovePosition Array(vPosition(0) - 1, vPosition(1), vPosition(2))
            Case 13 ' Pivot
                LogChange txtText.Text, txtText.SelStart, txtText.SelLength
                vPosition = FindPosition
                txtText.Text = ConvertTableToText(Pivot(PadTable(ConvertTextToTable(txtText.Text))))
                MovePosition Array(vPosition(1), vPosition(0), vPosition(2))
            Case 46 ' Delete Column
                LogChange txtText.Text, txtText.SelStart, txtText.SelLength
                vPosition = FindPosition
                vTable = Pivot(PadTable(ConvertTextToTable(txtText.Text)))
                DeleteRow vTable, vPosition(1)
                txtText.Text = ConvertTableToText(Pivot(vTable))
                MovePosition Array(vPosition(0), vPosition(1), 0)
            Case 45 ' Insert Column
                LogChange txtText.Text, txtText.SelStart, txtText.SelLength
                vPosition = FindPosition
                vTable = Pivot(PadTable(ConvertTextToTable(txtText.Text)))
                InsertRow vTable, vPosition(1)
                txtText.Text = ConvertTableToText(Pivot(vTable))
                MovePosition Array(vPosition(0), vPosition(1), 0)
            Case 67 ' Duplicate Column
                LogChange txtText.Text, txtText.SelStart, txtText.SelLength
                vPosition = FindPosition
                vTable = Pivot(PadTable(ConvertTextToTable(txtText.Text)))
                DuplicateRow vTable, vPosition(1)
                txtText.Text = ConvertTableToText(Pivot(vTable))
                MovePosition Array(vPosition(0), vPosition(1), 0)
            Case 89 ' Delete Row
                LogChange txtText.Text, txtText.SelStart, txtText.SelLength
                vPosition = FindPosition
                vTable = ConvertTextToTable(txtText.Text)
                DeleteRow vTable, vPosition(0)
                txtText.Text = ConvertTableToText(vTable)
                MovePosition Array(vPosition(0), vPosition(1), 0)
            Case 73 ' Insert Row
                LogChange txtText.Text, txtText.SelStart, txtText.SelLength
                vPosition = FindPosition
                vTable = ConvertTextToTable(txtText.Text)
                InsertRow vTable, vPosition(0)
                txtText.Text = ConvertTableToText(vTable)
                MovePosition Array(vPosition(0), vPosition(1), 0)
            Case 109 ' Reduce tab stops
                mlTabSize = mlTabSize - 1
                UpdateTabs mlTabSize
            Case 107 ' Increase tab stops
                mlTabSize = mlTabSize + 1
                UpdateTabs mlTabSize
                
            Case 83 ' Sort Ascending
                LogChange txtText.Text, txtText.SelStart, txtText.SelLength
                vPosition = FindPosition
                vTable = PadTable(ConvertTextToTable(txtText.Text))
                SortByColumn vTable, vPosition(1), False, -(Shift And 1)
                txtText.Text = ConvertTableToText(vTable)
                MovePosition Array(vPosition(0), vPosition(1), 0)
            Case 68 ' Sort Descending
                LogChange txtText.Text, txtText.SelStart, txtText.SelLength
                vPosition = FindPosition
                vTable = PadTable(ConvertTextToTable(txtText.Text))
                SortByColumn vTable, vPosition(1), True, -(Shift And 1)
                txtText.Text = ConvertTableToText(vTable)
                MovePosition Array(vPosition(0), vPosition(1), 0)
                
            Case 85 ' Upper case
                LogChange txtText.Text, txtText.SelStart, txtText.SelLength
                vPosition = FindPosition
                vTable = Pivot(PadTable(ConvertTextToTable(txtText.Text)))
                ChangeCaseRow vTable, vPosition(1), True
                txtText.Text = ConvertTableToText(Pivot(vTable))
                MovePosition Array(vPosition(0), vPosition(1), 0)
            Case 76 ' Lower case
                LogChange txtText.Text, txtText.SelStart, txtText.SelLength
                vPosition = FindPosition
                vTable = Pivot(PadTable(ConvertTextToTable(txtText.Text)))
                ChangeCaseRow vTable, vPosition(1), False
                txtText.Text = ConvertTableToText(Pivot(vTable))
                MovePosition Array(vPosition(0), vPosition(1), 0)
            Case 84 ' Insert Tab
                Dim lSelStart As Long
                lSelStart = txtText.SelStart
                txtText.Text = Left$(txtText.Text, txtText.SelStart) & vbTab & Mid$(txtText.Text, txtText.SelStart + txtText.SelLength + 1)
                txtText.SelStart = lSelStart + 1
            Case 87 ' Remove Duplicates
                txtText.Text = ConvertTableToText(RemoveDuplicates(ConvertTextToTable(txtText.Text)))
            Case 82 'Reset tabs
                UpdateTabs mlTabSize
        End Select
    End If
End Sub

Private Function Pivot(vTable As Variant) As Variant
    Dim lColumn As Long
    Dim vPivotTable As Variant
    Dim vPivotRow As Variant
    
    Dim lRow As Long
    
    Dim lRows As Long
    Dim lColumns As Long
    
    vPivotTable = Array()
    lRows = UBound(vTable)
    
    If UBound(vTable) > -1 Then
        lColumns = UBound(vTable(0))
        
        vPivotTable = Array()
        ReDim vPivotTable(lColumns)
        
        For lColumn = 0 To lColumns
            vPivotRow = Array()
            ReDim vPivotRow(lRows)
            For lRow = 0 To lRows
                vPivotRow(lRow) = vTable(lRow)(lColumn)
            Next
            vPivotTable(lColumn) = vPivotRow
        Next
    End If
    Pivot = vPivotTable
End Function

Private Function ConvertTextToTable(sText As String) As Variant
    Dim vTable As Variant
    Dim vRows As Variant
    Dim vRow As Variant
    
    vTable = Array()
    
    vRows = Split(sText, vbCrLf)
    For Each vRow In vRows
        ReDim Preserve vTable(UBound(vTable) + 1)
        vTable(UBound(vTable)) = Split(vRow, vbTab)
    Next
    ConvertTextToTable = vTable
End Function

Private Function ConvertTableToText(vTable As Variant) As String
    Dim vRow As Variant
    Dim vRows As Variant
    
    vRows = Array()
    For Each vRow In vTable
        ReDim Preserve vRows(UBound(vRows) + 1)
        vRows(UBound(vRows)) = Join(vRow, vbTab)
    Next
    ConvertTableToText = Join(vRows, vbCrLf)
End Function

Private Function PadTable(vTable As Variant) As Variant
    Dim lRow As Long
    Dim lMaxWidth As Long
    Dim vRow As Variant
    Dim vPaddedTable As Variant
    
    For lRow = 0 To UBound(vTable)
        vRow = vTable(lRow)
        If UBound(vRow) > lMaxWidth Then
            lMaxWidth = UBound(vRow)
        End If
    Next
    
    vPaddedTable = Array()
    ReDim vPaddedTable(UBound(vTable))
    
    For lRow = 0 To UBound(vTable)
        vRow = vTable(lRow)
        ReDim Preserve vRow(lMaxWidth)
        vPaddedTable(lRow) = vRow
    Next
    
    PadTable = vPaddedTable
End Function

Private Function FindPosition() As Variant
    Dim lPosition As Long
    Dim vRows As Variant
    Dim vRow As Variant
    Dim lRow As Long
    Dim lColumn As Long
    Dim lOffset As Long
    Dim vRowSplit As Variant
    
    lPosition = txtText.SelStart
    
    vRows = Split(Left$(txtText.Text, lPosition), vbCrLf)

    lRow = UBound(vRows)
    If lRow > -1 Then
        vRow = vRows(lRow)
    
        vRowSplit = Split(vRow, vbTab)
        lColumn = UBound(vRowSplit)
        
        If lColumn > -1 Then
            lOffset = Len(vRowSplit(lColumn))
        Else
            lColumn = 0
        End If
    End If
    FindPosition = Array(lRow, lColumn, lOffset)
End Function

Private Sub MovePosition(vPosition As Variant)
    Dim vRows As Variant
    Dim lRow As Long
    Dim sText As String
    Dim vRow As Variant
    Dim lColumn As Long
    
    vRows = Split(txtText.Text, vbCrLf)
    If vPosition(0) < 0 Then
        vPosition(0) = 0
    End If
    If vPosition(0) > UBound(vRows) Then
        vPosition(0) = UBound(vRows)
    End If
    If vPosition(1) < 0 Then
        vPosition(1) = 0
    End If
    For lRow = 0 To vPosition(0) - 1
        sText = sText & vRows(lRow) & vbCrLf
    Next
    If UBound(vRows) > -1 Then
        vRow = Split(vRows(vPosition(0)), vbTab)
        If UBound(vRow) > -1 Then
            For lColumn = 0 To vPosition(1) - 1
                sText = sText & vRow(lColumn) & vbTab
            Next
        End If
        If UBound(vRow) > -1 Then
            If vPosition(1) > UBound(vRow) Then
                sText = sText & Left$(vRow(UBound(vRow)), vPosition(2))
            Else
                sText = sText & Left$(vRow(vPosition(1)), vPosition(2))
            End If
        End If
        txtText.SelStart = Len(sText)
    End If
End Sub

Private Sub MoveRow(vTable As Variant, ByVal lFrom As Long, ByVal lDirection As Long)
    Dim lRows As Long
    Dim vTemp As Variant
    
    lRows = UBound(vTable)
    
    If (lFrom = lRows And lDirection = 1) Or (lFrom = 0 And lDirection = -1) Or lFrom < 0 Or lFrom > lRows Then
        Exit Sub
    End If
    
    vTemp = vTable(lFrom)
    vTable(lFrom) = vTable(lFrom + lDirection)
    vTable(lFrom + lDirection) = vTemp
End Sub

Private Sub DeleteRow(vTable As Variant, ByVal lRow As Long)
    Dim lRowIndex As Long
    
    If lRow < 0 Then
        lRow = 0
    End If
    For lRowIndex = lRow To UBound(vTable) - 1
        vTable(lRowIndex) = vTable(lRowIndex + 1)
    Next
    If UBound(vTable) > 0 Then
        ReDim Preserve vTable(UBound(vTable) - 1)
    Else
        vTable = Array()
    End If
End Sub

Private Sub InsertRow(vTable As Variant, ByVal lRow As Long)
    Dim lRowIndex As Long
    
    If lRow = -1 Then
        lRow = 0
    End If
    ReDim Preserve vTable(UBound(vTable) + 1)
    For lRowIndex = UBound(vTable) To lRow + 1 Step -1
        vTable(lRowIndex) = vTable(lRowIndex - 1)
    Next
    vTable(lRow) = Array()
    vTable = PadTable(vTable)
End Sub

Private Sub ChangeCaseRow(vTable As Variant, ByVal lRow As Long, ByVal bUpper As Boolean)
    Dim lColumnIndex As Long
    
    If lRow = -1 Then
        lRow = 0
    End If
    For lColumnIndex = LBound(vTable(lRow)) To UBound(vTable(lRow))
        If bUpper Then
            vTable(lRow)(lColumnIndex) = UCase$(vTable(lRow)(lColumnIndex))
        Else
            vTable(lRow)(lColumnIndex) = LCase$(vTable(lRow)(lColumnIndex))
        End If
    Next
End Sub

Private Sub DuplicateRow(vTable As Variant, ByVal lRow As Long)
    Dim lRowIndex As Long
    
    If lRow = -1 Then
        lRow = 0
    End If
    ReDim Preserve vTable(UBound(vTable) + 1)
    For lRowIndex = UBound(vTable) To lRow + 1 Step -1
        vTable(lRowIndex) = vTable(lRowIndex - 1)
    Next
    vTable(lRow) = vTable(lRow + 1)
    vTable = PadTable(vTable)
End Sub

' Removes consecutive duplicate rows
Private Function RemoveDuplicates(vTable As Variant) As Variant
    Dim bDuplicate As Boolean
    Dim lColumn As Long
    Dim lRow As Long
    Dim vRow As Variant
    Dim vPreviousRow As Variant
    Dim vOutputTable As Variant
    
    If UBound(vTable) < 1 Then
        RemoveDuplicates = vTable
        Exit Function
    Else
        vOutputTable = Array()
    End If
    
    For lRow = 1 To UBound(vTable)
        vRow = vTable(lRow)
        vPreviousRow = vTable(lRow - 1)
        
        bDuplicate = True
        For lColumn = 0 To UBound(vRow)
            If vRow(lColumn) <> vPreviousRow(lColumn) Then
                bDuplicate = False
                Exit For
            End If
        Next
        If Not bDuplicate Then
            ReDim Preserve vOutputTable(UBound(vOutputTable) + 1)
            vOutputTable(UBound(vOutputTable)) = vPreviousRow
        End If
    Next
    ReDim Preserve vOutputTable(UBound(vOutputTable) + 1)
    vOutputTable(UBound(vOutputTable)) = vRow
    
    RemoveDuplicates = vOutputTable
End Function

Private Sub SortByColumn(vTable As Variant, ByVal lColumn As Long, Optional ByVal bDescending As Boolean, Optional ByVal bIntelligentCompare As Boolean)
    Dim vTemp As Variant
    Dim bSorted As Boolean
    Dim lRow As Long
    
    While Not bSorted
        bSorted = True
        For lRow = 0 To UBound(vTable) - 1
            If Not bDescending Then
                If Compare(vTable(lRow)(lColumn), vTable(lRow + 1)(lColumn), bIntelligentCompare) Then
                    vTemp = vTable(lRow)
                    vTable(lRow) = vTable(lRow + 1)
                    vTable(lRow + 1) = vTemp
                    bSorted = False
                End If
            Else
                If Compare(vTable(lRow + 1)(lColumn), vTable(lRow)(lColumn), bIntelligentCompare) Then
                    vTemp = vTable(lRow)
                    vTable(lRow) = vTable(lRow + 1)
                    vTable(lRow + 1) = vTemp
                    bSorted = False
                End If
            End If
        Next
    Wend
End Sub


