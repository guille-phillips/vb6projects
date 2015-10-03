VERSION 5.00
Begin VB.Form frmQuickPad 
   AutoRedraw      =   -1  'True
   Caption         =   "Quick Pad"
   ClientHeight    =   4530
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   9360
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
   ScaleHeight     =   302
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   624
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdConvertToList 
      Caption         =   "Convert To List"
      Height          =   375
      Left            =   7920
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   4080
      Width           =   1335
   End
   Begin VB.CommandButton cmdReduceNewlines 
      Caption         =   "Reduce #n"
      Height          =   375
      Left            =   6480
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   4080
      Width           =   1335
   End
   Begin VB.CommandButton cmdReduceSpaces 
      Caption         =   "Reduce Spaces"
      Height          =   375
      Left            =   5040
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   4080
      Width           =   1335
   End
   Begin VB.CommandButton cmdReduceTabs 
      Caption         =   "Reduce Tabs"
      Height          =   375
      Left            =   3720
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   4080
      Width           =   1215
   End
   Begin VB.CommandButton cmdNoEndSpace 
      Caption         =   "No End Spaces"
      Height          =   375
      Left            =   2280
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   4080
      Width           =   1335
   End
   Begin VB.CommandButton cmdNoSpaces 
      Caption         =   "No Spaces"
      Height          =   375
      Left            =   1200
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   4080
      Width           =   975
   End
   Begin VB.CommandButton cmdNoTabs 
      Caption         =   "No Tabs"
      Height          =   375
      Left            =   120
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   4080
      Width           =   975
   End
   Begin VB.TextBox txtRowDel 
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
      Left            =   4815
      MultiLine       =   -1  'True
      TabIndex        =   9
      TabStop         =   0   'False
      Text            =   "frmQuickPad.frx":0000
      Top             =   3390
      Width           =   3120
   End
   Begin VB.TextBox txtColDel 
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
      TabIndex        =   7
      TabStop         =   0   'False
      Text            =   "frmQuickPad.frx":0003
      Top             =   3390
      Width           =   3120
   End
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
      Left            =   960
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   5520
      Visible         =   0   'False
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
      Width           =   8415
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
      Width           =   8415
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
   Begin VB.Label lblRowDel 
      Caption         =   "Row Del.:"
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
      Left            =   4080
      TabIndex        =   10
      Top             =   3420
      Width           =   735
   End
   Begin VB.Label lblColDel 
      Caption         =   "Col. Del.:"
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
      Left            =   105
      TabIndex        =   8
      Top             =   3420
      Width           =   735
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
      Left            =   240
      TabIndex        =   6
      Top             =   5520
      Visible         =   0   'False
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
      TabIndex        =   5
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
      TabIndex        =   4
      Top             =   2670
      Width           =   495
   End
End
Attribute VB_Name = "frmQuickPad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function ShowScrollBar Lib "user32" (ByVal hWnd As Long, ByVal wBar As Long, ByVal bShow As Long) As Long
Private Declare Function SetClipboardData Lib "user32" (ByVal wFormat As Long, ByVal hMem As Long) As Long
Private Declare Function EmptyClipboard Lib "user32" () As Long
Private Declare Function OpenClipboard Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function CloseClipboard Lib "user32" () As Long

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
Private mnDelimitersTop As Single
Private mnListTop As Single
Private mnButtonsTop As Single

Private moReplaceLex As ISaffronObject
Private moSortLex As ISaffronObject
Private moSetLex As ISaffronObject
Private moCamelLex As ISaffronObject

Private msChanges As Variant

Private miShift As Integer
Private mlDragStop As Long

Private moNewSetLex As ISaffronObject

Private mstRange As SelectionTypes

Private msColumnDelimiter As String
Private msRowDelimiter As String

Private msOriginalText As String
Private mbCodeToggle As Boolean

Private Enum SelectionTypes
    Column
    Row
    Table
    Cell
End Enum

Private Enum StringFunctionTypes
    LowerCase
    UpperCase
    CamelCase
    Underscored
    SpaceCase
End Enum

Private Sub cmdConvertToList_Click()
    txtText.Text = Replace$(txtText.Text, vbCrLf, ",")
End Sub

Private Sub cmdNoEndSpace_Click()
    Dim sOriginal As String
    
    sOriginal = txtText.Text
    
    Do
        txtText.Text = Replace$(txtText.Text, " " & vbCrLf, vbCrLf)
        txtText.Text = Replace$(txtText.Text, vbTab & vbCrLf, vbCrLf)
        If txtText.Text = sOriginal Then
            Exit Sub
        End If
        sOriginal = txtText.Text
    Loop
End Sub

Private Sub cmdNoSpaces_Click()
    txtText.Text = Replace$(txtText.Text, " ", "")
End Sub

Private Sub cmdNoTabs_Click()
    txtText.Text = Replace$(txtText.Text, vbTab, "")
End Sub

Private Sub cmdReduceNewlines_Click()
    Dim sOriginal As String
    
    sOriginal = txtText.Text
    
    Do
        txtText.Text = Replace$(txtText.Text, vbCrLf & vbCrLf, vbCrLf)
        If txtText.Text = sOriginal Then
            Exit Sub
        End If
        sOriginal = txtText.Text
    Loop
End Sub

Private Sub cmdReduceSpaces_Click()
    Dim sOriginal As String
    
    sOriginal = txtText.Text
    
    Do
        txtText.Text = Replace$(txtText.Text, "  ", " ")
        If txtText.Text = sOriginal Then
            Exit Sub
        End If
        sOriginal = txtText.Text
    Loop
End Sub

Private Sub cmdReduceTabs_Click()
    Dim sOriginal As String
    
    sOriginal = txtText.Text
    
    Do
        txtText.Text = Replace$(txtText.Text, vbTab & vbTab, vbTab)
        If txtText.Text = sOriginal Then
            Exit Sub
        End If
        sOriginal = txtText.Text
    Loop
End Sub

Private Sub Form_Initialize()
    Range = Column
    msColumnDelimiter = vbTab
    msRowDelimiter = vbCrLf
End Sub

Private Sub Form_Load()
    Dim sTabStops As String
    
    mnBoxWidth = Me.Width / Screen.TwipsPerPixelX - txtFind.Width
    mnTextHeight = Me.Height / Screen.TwipsPerPixelY - txtText.Height
    mnFindTop = Me.Height / Screen.TwipsPerPixelY - txtFind.Top
    mnReplaceTop = Me.Height / Screen.TwipsPerPixelY - txtReplace.Top
    mnDelimitersTop = Me.Height / Screen.TwipsPerPixelY - txtColDel.Top
    mnListTop = Me.Height / Screen.TwipsPerPixelY - txtList.Top
    mnButtonsTop = Me.Height / Screen.TwipsPerPixelY - cmdNoSpaces.Top
    
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
    
    ShowScrollBar txtText.hWnd, SB_VERT, 1
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
    SendMessage txtText.hWnd, EM_SETTABSTOPS, UBound(mlTabStops) + 1, mlTabStops(0)
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
    Set moNewSetLex = Rules("element2")
    Set moCamelLex = Rules("camel_text")
    

'    Dim oResult As SaffronTree
'
'    Set oResult = New SaffronTree
'    SaffronStream.Text = "AlphaNumeric23| djjk"
'
'    If moCamelLex.Parse(oResult) Then
'        Stop
'    End If

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
    Dim lTabIndex As Long
    
    If mlDragStop > -1 Then
        If Button = vbLeftButton Then
            If Shift = 0 Then
                mlTabStops(mlDragStop) = X / 1.754
            Else
                For lTabIndex = 0 To UBound(mlTabStops)
                    mlTabStops(lTabIndex) = (CSng(lTabIndex) / CSng(mlDragStop)) * X / 1.754
                Next
            End If
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
    txtColDel.Top = Me.Height / Screen.TwipsPerPixelY - mnDelimitersTop
    txtRowDel.Top = Me.Height / Screen.TwipsPerPixelY - mnDelimitersTop
    txtList.Top = Me.Height / Screen.TwipsPerPixelY - mnListTop
    cmdNoEndSpace.Top = Me.Height / Screen.TwipsPerPixelY - mnButtonsTop
    cmdNoSpaces.Top = Me.Height / Screen.TwipsPerPixelY - mnButtonsTop
    cmdNoTabs.Top = Me.Height / Screen.TwipsPerPixelY - mnButtonsTop
    cmdReduceNewlines.Top = Me.Height / Screen.TwipsPerPixelY - mnButtonsTop
    cmdReduceSpaces.Top = Me.Height / Screen.TwipsPerPixelY - mnButtonsTop
    cmdReduceTabs.Top = Me.Height / Screen.TwipsPerPixelY - mnButtonsTop
    cmdConvertToList.Top = Me.Height / Screen.TwipsPerPixelY - mnButtonsTop
    lblFind.Top = Me.Height / Screen.TwipsPerPixelY - mnFindTop + 3
    lblReplace.Top = Me.Height / Screen.TwipsPerPixelY - mnReplaceTop + 3
    lblList.Top = Me.Height / Screen.TwipsPerPixelY - mnListTop + 3
    lblColDel.Top = Me.Height / Screen.TwipsPerPixelY - mnDelimitersTop + 3
    lblRowDel.Top = Me.Height / Screen.TwipsPerPixelY - mnDelimitersTop + 3
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

Private Sub txtColDel_Change()
    msColumnDelimiter = Decode(txtColDel.Text)
End Sub

Private Sub txtRowDel_Change()
    msRowDelimiter = Decode(txtRowDel.Text)
End Sub

Private Sub txtFind_GotFocus()
    txtText.TabStop = False
    txtFind.TabStop = True
    txtReplace.TabStop = True
    txtList.TabStop = False
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

Private Sub txtList_Change()
    txtList.ForeColor = vbBlack
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
            InsertTextIntoSelection oSetDef.AllText
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
    
'    iSelStart = txtText.SelStart
'    iSelLength = txtText.SelLength
'
    sFind = Decode(txtFind.Text)
    sReplace = Decode(txtReplace.Text)

'    sLeft = Left$(txtText.Text, txtText.SelStart)
'    sMiddle = txtText.SelText
'    sRight = Mid$(txtText.Text, txtText.SelStart + txtText.SelLength + 1)
'
'    If iSelLength <> 0 Then
'        txtText.Text = sLeft & Replace$(sMiddle, sFind, sReplace) & sRight
'    Else
'        txtText.Text = Replace$(sLeft & sRight, sFind, sReplace)
'    End If
'
'    txtText.SelStart = iSelStart
'    txtText.SelLength = iSelLength
    
    txtText.Text = Replace$(txtText.Text, sFind, sReplace)
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
                Case 4 ' #b
                    Decode = Decode & vbCrLf
                Case 5
                    Decode = Decode & Chr(Val(oSub.Text))
                Case 6
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

Private Sub txtText_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim sTextCodes As String
    
    If KeyCode = 17 Then
        Exit Sub
    End If
    
    Select Case KeyCode
        Case vbKeyV
            KeyCode = 0
    End Select
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
        txtColDel.Text = "#t"
        txtRowDel.Text = "#n"
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
    
    vLines = Split(txtText.Text, msRowDelimiter)
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
    
    On Error Resume Next
    sLeft = Left$(txtText.Text, txtText.SelStart)
    sRight = Mid$(txtText.Text, txtText.SelStart + txtText.SelLength + 1)
    txtText.Text = sLeft & sInsert & sRight
    txtText.SelLength = Len(sInsert)
End Function

Private Property Let Range(ByVal stRange As SelectionTypes)
    mstRange = stRange
    Me.Caption = "Quick Pad - [" & Array("Column", "Row", "Table", "Cell")(stRange) & "]"
End Property


Private Sub txtText_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim vPosition As Variant
    Dim vTable As Variant
    Dim lTabIndex As Long
    Dim vColumn As Variant
    Dim vArray As Variant
    Dim sTextCodes As String
    Dim sJoin As String
    
    If KeyCode = vbKeyInsert And Shift = 0 Then
        Range = (mstRange + 1) Mod 4
    End If
    
'    If Shift = 2 Then
'        Select Case KeyCode
'            Case vbKeyC
'                vPosition = FindPosition
'                Select Case mstRange
'                    Case SelectionTypes.Column
'                        vColumn = ArrayFromTableColumn(PadTable(ConvertTextToTable(txtText.Text)), vPosition(1))
'                        Clipboard.SetText Join(vColumn, vbCrLf)
'                    Case SelectionTypes.Row
'                        vColumn = ArrayFromTableRow(PadTable(ConvertTextToTable(txtText.Text)), vPosition(0))
'                        Clipboard.SetText Join(vColumn, vbTab)
'                    Case SelectionTypes.Table
'                        Clipboard.SetText txtText.Text
'                    Case SelectionTypes.Cell
'                        vColumn = PadTable(ConvertTextToTable(txtText.Text))
'                        Clipboard.SetText vColumn(vPosition(0))(vPosition(1))
'                End Select
'        End Select
'    End If
    
    If (Shift And 4) = 4 Then
        Select Case KeyCode
            Case vbKeyRight, vbKeyLeft, vbKeyDown, vbKeyUp, vbKeyReturn, vbKeyDelete, vbKeyInsert, vbKeyC, vbKeyY, vbKeyI, vbKeyS, vbKeyD, vbKeyU, vbKeyL, vbKeyN, vbKeyM, vbKeyB, vbKeyH, vbKeyJ, vbKeyPageDown
                LogChange txtText.Text, txtText.SelStart, txtText.SelLength
                vPosition = FindPosition
                vTable = PadTable(ConvertTextToTable(txtText.Text))
        End Select
        
        Select Case KeyCode
            Case 223
                If Not mbCodeToggle Then
                    msOriginalText = txtText.Text
                    sTextCodes = txtText.Text
                    sTextCodes = Replace$(sTextCodes, vbCrLf, "¶" & vbCrLf)
                    sTextCodes = Replace$(sTextCodes, vbTab, "»" & vbTab)
                    sTextCodes = Replace$(sTextCodes, " ", "·")
                    txtText.Text = sTextCodes
                    txtText.Locked = True
                Else
                    txtText.Text = msOriginalText
                    txtText.Locked = False
                End If
                mbCodeToggle = Not mbCodeToggle
                
            ' table functions
            Case vbKeyRight ' Column Right
                vTable = Pivot(vTable)
                If MoveRow(vTable, vPosition(1), 1) Then
                    txtText.Text = ConvertTableToText(Pivot(vTable))
                    MovePosition Array(vPosition(0), vPosition(1) + 1, vPosition(2))
                End If
            Case vbKeyLeft ' Column Left
                vTable = Pivot(vTable)
                If MoveRow(vTable, vPosition(1), -1) Then
                    txtText.Text = ConvertTableToText(Pivot(vTable))
                    MovePosition Array(vPosition(0), vPosition(1) - 1, vPosition(2))
                End If
            Case vbKeyDown ' Row Down
                If MoveRow(vTable, vPosition(0), 1) Then
                    txtText.Text = ConvertTableToText(vTable)
                    MovePosition Array(vPosition(0) + 1, vPosition(1), vPosition(2))
                End If
            Case vbKeyUp ' Row Up
                If MoveRow(vTable, vPosition(0), -1) Then
                    txtText.Text = ConvertTableToText(vTable)
                    MovePosition Array(vPosition(0) - 1, vPosition(1), vPosition(2))
                End If
            Case vbKeyReturn ' Pivot
                txtText.Text = ConvertTableToText(Pivot(vTable))
                MovePosition Array(vPosition(1), vPosition(0), vPosition(2))
            Case vbKeyH ' Convert to HTML table
                txtText.Text = TableToHTML(txtText.Text)
            Case vbKeyJ ' Convert from HTML table
                txtText.Text = HTMLToTable(txtText.Text)
            Case vbKeyW ' Remove Duplicates
                txtText.Text = ConvertTableToText(RemoveDuplicates(ConvertTextToTable(txtText.Text)))
                
            ' range functions
            Case vbKeyDelete ' Delete Column
                Select Case mstRange
                    Case SelectionTypes.Column
                        vTable = Pivot(vTable)
                        DeleteRow vTable, vPosition(1)
                        vTable = Pivot(vTable)
                        txtText.Text = ConvertTableToText(vTable)
                        MovePosition Array(vPosition(0), vPosition(1), 0)
                    Case SelectionTypes.Row
                        DeleteRow vTable, vPosition(0)
                        txtText.Text = ConvertTableToText(vTable)
                        MovePosition Array(vPosition(0), vPosition(1), 0)
                    Case SelectionTypes.Cell
                        vTable(vPosition(0))(vPosition(1)) = ""
                        txtText.Text = ConvertTableToText(vTable)
                        MovePosition Array(vPosition(0), vPosition(1), 0)
                    Case SelectionTypes.Table
                        txtText.Text = ""
                End Select

            Case vbKeyInsert ' Insert Column/Row
                Select Case mstRange
                    Case SelectionTypes.Column
                        vTable = Pivot(vTable)
                        InsertRow vTable, vPosition(1)
                        txtText.Text = ConvertTableToText(Pivot(vTable))
                        MovePosition Array(vPosition(0), vPosition(1), 0)
                    Case SelectionTypes.Row
                        InsertRow vTable, vPosition(0)
                        txtText.Text = ConvertTableToText(vTable)
                        MovePosition Array(vPosition(0), vPosition(1), 0)
                    Case SelectionTypes.Cell
                    Case SelectionTypes.Table
                End Select

            Case vbKeyC ' Duplicate Column/Row
                Select Case mstRange
                    Case SelectionTypes.Column
                        vTable = Pivot(vTable)
                        DuplicateRow vTable, vPosition(1)
                        txtText.Text = ConvertTableToText(Pivot(vTable))
                        MovePosition Array(vPosition(0), vPosition(1), 0)
                    Case SelectionTypes.Row
                        DuplicateRow vTable, vPosition(0)
                        txtText.Text = ConvertTableToText(vTable)
                        MovePosition Array(vPosition(0), vPosition(1), 0)
                    Case SelectionTypes.Cell
                    Case SelectionTypes.Table
                        txtText.Text = ConvertTableToText(vTable) & msRowDelimiter & ConvertTableToText(vTable)
                End Select

            Case vbKeyS ' Sort Ascending
                Select Case mstRange
                    Case SelectionTypes.Column
                        vColumn = ConvertArrayToTable(ArrayFromTableColumn(vTable, vPosition(1)))
                        vColumn = SortByColumn(vColumn, 0, False, -(Shift And 1))
                        vTable = ArrayIntoTableColumn(vTable, vPosition(1), ArrayFromTableColumn(vColumn, 0))
                        txtText.Text = ConvertTableToText(vTable)
                        MovePosition Array(vPosition(0), vPosition(1), 0)
                    Case SelectionTypes.Row
                        vTable = Pivot(vTable)
                        vColumn = ConvertArrayToTable(ArrayFromTableColumn(vTable, vPosition(0)))
                        vColumn = SortByColumn(vColumn, 0, False, -(Shift And 1))
                        vTable = ArrayIntoTableColumn(vTable, vPosition(0), ArrayFromTableColumn(vColumn, 0))
                        txtText.Text = ConvertTableToText(Pivot(vTable))
                        MovePosition Array(vPosition(0), vPosition(1), 0)
                    Case SelectionTypes.Cell
                    Case SelectionTypes.Table
                        vTable = SortByColumn(vTable, vPosition(1), False, -(Shift And 1))
                        txtText.Text = ConvertTableToText(vTable)
                        MovePosition Array(vPosition(0), vPosition(1), 0)
                End Select

            Case vbKeyD ' Sort Descending
                Select Case mstRange
                    Case SelectionTypes.Column
                        vColumn = ConvertArrayToTable(ArrayFromTableColumn(vTable, vPosition(1)))
                        vColumn = SortByColumn(vColumn, 0, True, -(Shift And 1))
                        vTable = ArrayIntoTableColumn(vTable, vPosition(1), ArrayFromTableColumn(vColumn, 0))
                        txtText.Text = ConvertTableToText(vTable)
                        MovePosition Array(vPosition(0), vPosition(1), 0)
                    Case SelectionTypes.Row
                        vTable = Pivot(vTable)
                        vColumn = ConvertArrayToTable(ArrayFromTableColumn(vTable, vPosition(0)))
                        vColumn = SortByColumn(vColumn, 0, True, -(Shift And 1))
                        vTable = ArrayIntoTableColumn(vTable, vPosition(0), ArrayFromTableColumn(vColumn, 0))
                        txtText.Text = ConvertTableToText(Pivot(vTable))
                        MovePosition Array(vPosition(0), vPosition(1), 0)
                    Case SelectionTypes.Cell
                    Case SelectionTypes.Table
                        vTable = SortByColumn(vTable, vPosition(1), True, -(Shift And 1))
                        txtText.Text = ConvertTableToText(vTable)
                        MovePosition Array(vPosition(0), vPosition(1), 0)
                End Select

            Case vbKeyU ' Upper case
                Select Case mstRange
                    Case SelectionTypes.Column
                        vArray = ArrayFromTableColumn(vTable, vPosition(1))
                        vArray = ApplyStringFunction(vArray, UpperCase)
                        vTable = ArrayIntoTableColumn(vTable, vPosition(1), vArray)
                        txtText.Text = ConvertTableToText(vTable)
                        MovePosition Array(vPosition(0), vPosition(1), vPosition(2))
                    Case SelectionTypes.Row
                        vArray = ArrayFromTableRow(vTable, vPosition(0))
                        vArray = ApplyStringFunction(vArray, UpperCase)
                        vTable = ArrayIntoTableRow(vTable, vPosition(0), vArray)
                        txtText.Text = ConvertTableToText(vTable)
                        MovePosition Array(vPosition(0), vPosition(1), vPosition(2))
                    Case SelectionTypes.Cell
                        vTable(vPosition(0))(vPosition(1)) = ApplyStringFunction(vTable(vPosition(0))(vPosition(1)), UpperCase)
                        txtText.Text = ConvertTableToText(vTable)
                        MovePosition Array(vPosition(0), vPosition(1), vPosition(2))
                    Case SelectionTypes.Table
                        vTable = ApplyStringFunction(vTable, UpperCase)
                        txtText.Text = ConvertTableToText(vTable)
                        MovePosition Array(vPosition(0), vPosition(1), vPosition(2))
                End Select

            Case vbKeyL ' Lower case
                Select Case mstRange
                    Case SelectionTypes.Column
                        vArray = ArrayFromTableColumn(vTable, vPosition(1))
                        vArray = ApplyStringFunction(vArray, LowerCase)
                        vTable = ArrayIntoTableColumn(vTable, vPosition(1), vArray)
                        txtText.Text = ConvertTableToText(vTable)
                        MovePosition Array(vPosition(0), vPosition(1), vPosition(2))
                    Case SelectionTypes.Row
                        vArray = ArrayFromTableRow(vTable, vPosition(0))
                        vArray = ApplyStringFunction(vArray, LowerCase)
                        vTable = ArrayIntoTableRow(vTable, vPosition(0), vArray)
                        txtText.Text = ConvertTableToText(vTable)
                        MovePosition Array(vPosition(0), vPosition(1), vPosition(2))
                    Case SelectionTypes.Cell
                        vTable(vPosition(0))(vPosition(1)) = ApplyStringFunction(vTable(vPosition(0))(vPosition(1)), LowerCase)
                        txtText.Text = ConvertTableToText(vTable)
                        MovePosition Array(vPosition(0), vPosition(1), vPosition(2))
                    Case SelectionTypes.Table
                        vTable = ApplyStringFunction(vTable, LowerCase)
                        txtText.Text = ConvertTableToText(vTable)
                        MovePosition Array(vPosition(0), vPosition(1), vPosition(2))
                End Select
                
            Case vbKeyN ' To lower case with underscores from camel case
                Select Case mstRange
                    Case SelectionTypes.Column
                        vArray = ArrayFromTableColumn(vTable, vPosition(1))
                        vArray = ApplyStringFunction(vArray, Underscored)
                        vTable = ArrayIntoTableColumn(vTable, vPosition(1), vArray)
                        txtText.Text = ConvertTableToText(vTable)
                        MovePosition Array(vPosition(0), vPosition(1), vPosition(2))
                    Case SelectionTypes.Row
                        vArray = ArrayFromTableRow(vTable, vPosition(0))
                        vArray = ApplyStringFunction(vArray, Underscored)
                        vTable = ArrayIntoTableRow(vTable, vPosition(0), vArray)
                        txtText.Text = ConvertTableToText(vTable)
                        MovePosition Array(vPosition(0), vPosition(1), vPosition(2))
                    Case SelectionTypes.Cell
                        vTable(vPosition(0))(vPosition(1)) = ApplyStringFunction(vTable(vPosition(0))(vPosition(1)), Underscored)
                        txtText.Text = ConvertTableToText(vTable)
                        MovePosition Array(vPosition(0), vPosition(1), vPosition(2))
                    Case SelectionTypes.Table
                        vTable = ApplyStringFunction(vTable, Underscored)
                        txtText.Text = ConvertTableToText(vTable)
                        MovePosition Array(vPosition(0), vPosition(1), vPosition(2))
                End Select
                
            Case vbKeyM ' To camel case from lower case with underscores
                Select Case mstRange
                    Case SelectionTypes.Column
                        vArray = ArrayFromTableColumn(vTable, vPosition(1))
                        vArray = ApplyStringFunction(vArray, CamelCase)
                        vTable = ArrayIntoTableColumn(vTable, vPosition(1), vArray)
                        txtText.Text = ConvertTableToText(vTable)
                        MovePosition Array(vPosition(0), vPosition(1), vPosition(2))
                    Case SelectionTypes.Row
                        vArray = ArrayFromTableRow(vTable, vPosition(0))
                        vArray = ApplyStringFunction(vArray, CamelCase)
                        vTable = ArrayIntoTableRow(vTable, vPosition(0), vArray)
                        txtText.Text = ConvertTableToText(vTable)
                        MovePosition Array(vPosition(0), vPosition(1), vPosition(2))
                    Case SelectionTypes.Cell
                        vTable(vPosition(0))(vPosition(1)) = ApplyStringFunction(vTable(vPosition(0))(vPosition(1)), CamelCase)
                        txtText.Text = ConvertTableToText(vTable)
                        MovePosition Array(vPosition(0), vPosition(1), vPosition(2))
                    Case SelectionTypes.Table
                        vTable = ApplyStringFunction(vTable, CamelCase)
                        txtText.Text = ConvertTableToText(vTable)
                        MovePosition Array(vPosition(0), vPosition(1), vPosition(2))
                End Select
                
            Case vbKeyB ' Auto insert spaces at Upper case
                Select Case mstRange
                    Case SelectionTypes.Column
                        vArray = ArrayFromTableColumn(vTable, vPosition(1))
                        vArray = ApplyStringFunction(vArray, SpaceCase)
                        vTable = ArrayIntoTableColumn(vTable, vPosition(1), vArray)
                        txtText.Text = ConvertTableToText(vTable)
                        MovePosition Array(vPosition(0), vPosition(1), vPosition(2))
                    Case SelectionTypes.Row
                        vArray = ArrayFromTableRow(vTable, vPosition(0))
                        vArray = ApplyStringFunction(vArray, SpaceCase)
                        vTable = ArrayIntoTableRow(vTable, vPosition(0), vArray)
                        txtText.Text = ConvertTableToText(vTable)
                        MovePosition Array(vPosition(0), vPosition(1), vPosition(2))
                    Case SelectionTypes.Cell
                        vTable(vPosition(0))(vPosition(1)) = ApplyStringFunction(vTable(vPosition(0))(vPosition(1)), SpaceCase)
                        txtText.Text = ConvertTableToText(vTable)
                        MovePosition Array(vPosition(0), vPosition(1), vPosition(2))
                    Case SelectionTypes.Table
                        vTable = ApplyStringFunction(vTable, SpaceCase)
                        txtText.Text = ConvertTableToText(vTable)
                        MovePosition Array(vPosition(0), vPosition(1), vPosition(2))
                End Select
                
            Case vbKeyPageDown ' Fill with selected value
                Select Case mstRange
                    Case SelectionTypes.Column
                        vArray = ArrayFromTableColumn(vTable, vPosition(1))
                        vArray = FillArray(vArray, vTable(vPosition(0))(vPosition(1)))
                        vTable = ArrayIntoTableColumn(vTable, vPosition(1), vArray)
                        txtText.Text = ConvertTableToText(vTable)
                        MovePosition Array(vPosition(0), vPosition(1), vPosition(2))
                    Case SelectionTypes.Row
                        vArray = ArrayFromTableRow(vTable, vPosition(0))
                        vArray = FillArray(vArray, vTable(vPosition(0))(vPosition(1)))
                        vTable = ArrayIntoTableRow(vTable, vPosition(0), vArray)
                        txtText.Text = ConvertTableToText(vTable)
                        MovePosition Array(vPosition(0), vPosition(1), vPosition(2))
                    Case SelectionTypes.Cell
                    Case SelectionTypes.Table
                        vTable = FillArray(vTable, vTable(vPosition(0))(vPosition(1)))
                        txtText.Text = ConvertTableToText(vTable)
                        MovePosition Array(vPosition(0), vPosition(1), vPosition(2))
                End Select
                
            Case vbKeySubtract ' Reduce tab stops
                mlTabSize = mlTabSize - 1
                UpdateTabs mlTabSize
            Case vbKeyAdd ' Increase tab stops
                mlTabSize = mlTabSize + 1
                UpdateTabs mlTabSize
            Case vbKeyR 'Reset tabs
                UpdateTabs mlTabSize
            Case vbKeyE ' Derive list expression
                txtList.Text = DeriveExpression(PadTable(ConvertTextToTable(txtText.Text)))

        End Select
    End If
End Sub

Private Function TableToHTML(ByVal sTable As String) As String
    Dim vRows As Variant
    Dim vRow As Variant
    Dim vTable As Variant
    
    vRows = Array()
    
    vTable = ConvertTextToTable(sTable)
    For Each vRow In vTable
        ArrayAppend vRows, "<td>" & Join(vRow, "</td><td>") & "</td>"
    Next
    TableToHTML = "<table>" & vbCrLf & vbTab & "<tr>" & vbCrLf & vbTab & vbTab & Join(vRows, vbCrLf & vbTab & "</tr>" & vbCrLf & vbTab & "<tr>" & vbCrLf & vbTab & vbTab) & vbCrLf & vbTab & "</tr>" & vbCrLf & "</table>"
End Function

Private Function HTMLToTable(ByVal sString As String) As String
    sString = Replace$(sString, "</td></tr><tr><td>", msRowDelimiter)
    sString = Replace$(sString, "</td><td>", msColumnDelimiter)
    sString = Replace$(sString, "<table><tr><td>", "")
    sString = Replace$(sString, "</td></tr></table>", "")
    HTMLToTable = sString
End Function

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
    
    vRows = Split(sText, msRowDelimiter)
    For Each vRow In vRows
        ReDim Preserve vTable(UBound(vTable) + 1)
        vTable(UBound(vTable)) = Split(vRow, msColumnDelimiter)
    Next
    ConvertTextToTable = vTable
End Function

Private Function ConvertTableToText(vTable As Variant) As String
    Dim vRow As Variant
    Dim vRows As Variant
    
    vRows = Array()
    For Each vRow In vTable
        ReDim Preserve vRows(UBound(vRows) + 1)
        vRows(UBound(vRows)) = Join(vRow, msColumnDelimiter)
    Next
    ConvertTableToText = Join(vRows, msRowDelimiter)
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
    If lPosition = 0 Then
        FindPosition = Array(0, 0, 0)
        Exit Function
    End If
    vRows = Split(Left$(txtText.Text, lPosition), msRowDelimiter)

    lRow = UBound(vRows)
    If lRow > -1 Then
        vRow = vRows(lRow)
    
        vRowSplit = Split(vRow, msColumnDelimiter)
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
    
    vRows = Split(txtText.Text, msRowDelimiter)
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
        sText = sText & vRows(lRow) & msRowDelimiter
    Next
    If UBound(vRows) > -1 Then
        vRow = Split(vRows(vPosition(0)), msColumnDelimiter)
        If UBound(vRow) > -1 Then
            For lColumn = 0 To vPosition(1) - 1
                sText = sText & vRow(lColumn) & msColumnDelimiter
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

Private Function MoveRow(vTable As Variant, ByVal lFrom As Long, ByVal lDirection As Long) As Boolean
    Dim lRows As Long
    Dim vTemp As Variant
    
    lRows = UBound(vTable)
    
    If (lFrom = lRows And lDirection = 1) Or (lFrom = 0 And lDirection = -1) Or lFrom < 0 Or lFrom > lRows Then
        Exit Function
    End If
    
    MoveRow = True
    
    vTemp = vTable(lFrom)
    vTable(lFrom) = vTable(lFrom + lDirection)
    vTable(lFrom + lDirection) = vTemp
End Function

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


Private Function FillArray(ByVal vArray As Variant, ByVal sString As String) As Variant
    Dim vModified As Variant
    Dim lIndex As Long
    Dim lIndex2 As Long
    
    vModified = vArray
    For lIndex = 0 To UBound(vArray)
        If IsArray(vArray(lIndex)) Then
            For lIndex2 = 0 To UBound(vArray(lIndex))
                vModified(lIndex)(lIndex2) = sString
            Next
        Else
            vModified(lIndex) = sString
        End If
    Next
    FillArray = vModified
End Function

Private Function ApplyStringFunction(ByVal vArray As Variant, ByVal sftFunction As StringFunctionTypes) As Variant
    Dim lIndex1 As Long
    Dim lIndex2 As Long
    Dim vRow As Variant
    Dim vModified As Variant
    
    vModified = vArray
    If IsArray(vArray) Then
        For lIndex1 = 0 To UBound(vArray)
            vRow = vArray(lIndex1)
            If IsArray(vRow) Then
                For lIndex2 = 0 To UBound(vRow)
                    ' apply function
                    Select Case sftFunction
                        Case StringFunctionTypes.UpperCase
                            vModified(lIndex1)(lIndex2) = UCase$(vModified(lIndex1)(lIndex2))
                        Case StringFunctionTypes.LowerCase
                            vModified(lIndex1)(lIndex2) = LCase$(vModified(lIndex1)(lIndex2))
                        Case StringFunctionTypes.CamelCase
                            vModified(lIndex1)(lIndex2) = ApplyCamelCase(vModified(lIndex1)(lIndex2))
                        Case StringFunctionTypes.Underscored
                            vModified(lIndex1)(lIndex2) = ApplyUnderscoredCase(vModified(lIndex1)(lIndex2))
                        Case StringFunctionTypes.SpaceCase
                            vModified(lIndex1)(lIndex2) = ApplySpaceCase(vModified(lIndex1)(lIndex2))
                    End Select
                    
                Next
            Else
                ' apply function
                Select Case sftFunction
                    Case StringFunctionTypes.UpperCase
                        vModified(lIndex1) = UCase$(vModified(lIndex1))
                    Case StringFunctionTypes.LowerCase
                        vModified(lIndex1) = LCase$(vModified(lIndex1))
                    Case StringFunctionTypes.CamelCase
                        vModified(lIndex1) = ApplyCamelCase(vModified(lIndex1))
                    Case StringFunctionTypes.Underscored
                        vModified(lIndex1) = ApplyUnderscoredCase(vModified(lIndex1))
                    Case StringFunctionTypes.SpaceCase
                        vModified(lIndex1) = ApplySpaceCase(vModified(lIndex1))
                End Select
            End If
        Next
    Else
        ' apply function
        vModified = vModified
        Select Case sftFunction
            Case StringFunctionTypes.UpperCase
                vModified = UCase$(vModified)
            Case StringFunctionTypes.LowerCase
                vModified = LCase$(vModified)
            Case StringFunctionTypes.CamelCase
                vModified = ApplyCamelCase(vModified)
            Case StringFunctionTypes.Underscored
                vModified = ApplyUnderscoredCase(vModified)
            Case StringFunctionTypes.SpaceCase
                vModified = ApplySpaceCase(vModified)
        End Select
    End If
    ApplyStringFunction = vModified
End Function


Private Function ApplyCamelCase(ByVal sString As String) As String
    Dim bUpper As Boolean
    Dim sChar As String
    Dim lIndex As Long
    
    bUpper = True
        
    For lIndex = 1 To Len(sString)
        sChar = Mid$(sString, lIndex, 1)
        
        If sChar = "_" Then
            bUpper = True
        Else
            If bUpper Then
                ApplyCamelCase = ApplyCamelCase & UCase$(sChar)
                If LCase$(sChar) <> UCase$(sChar) Then
                    bUpper = False
                End If
            Else
                ApplyCamelCase = ApplyCamelCase & sChar
            End If
        End If
    Next
End Function

Private Function ApplyUnderscoredCase(ByVal sString As String) As String
    Dim bUnderscore As Boolean
    Dim sChar As String
    Dim sChar2 As String
    Dim lIndex As Long
    
    bUnderscore = False
        
    For lIndex = 1 To Len(sString)
        sChar = Mid$(sString, lIndex, 1)
        sChar2 = Mid$(sString, lIndex + 1, 1)
        
        If (UCase$(sChar) = sChar And LCase$(sChar) <> sChar) And LCase$(sChar2) = sChar2 And lIndex <> 1 Then
            bUnderscore = True
        Else
            bUnderscore = False
        End If
        
        ApplyUnderscoredCase = ApplyUnderscoredCase & IIf(bUnderscore, "_", "") & LCase$(sChar)
    Next
End Function

Private Function ApplySpaceCase(ByVal sString As String) As String
    Dim bSpace As Boolean
    Dim sChar As String
    Dim sChar2 As String
    Dim lIndex As Long
    
    bSpace = False
        
    For lIndex = 1 To Len(sString)
        sChar = Mid$(sString, lIndex, 1)
        sChar2 = Mid$(sString, lIndex + 1, 1)
        
        If (UCase$(sChar) = sChar And LCase$(sChar) <> sChar) And LCase$(sChar2) = sChar2 And lIndex <> 1 Then
            bSpace = True
        Else
            bSpace = False
        End If
        
        ApplySpaceCase = ApplySpaceCase & IIf(bSpace, " ", "") & sChar
    Next
End Function

Private Sub ChangeCaseRow(vTable As Variant, ByVal lRow As Long, ByVal bUpper As Boolean)
    Dim lColumnIndex As Long
    
    If UBound(vTable) = -1 Then
        Exit Sub
    End If
    
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

Private Sub ChangeToCamelCaseRow(vTable As Variant, ByVal lRow As Long)
    Dim lColumnIndex As Long
    Dim bUpper As Boolean
    Dim sCellText As String
    Dim lIndex As Long
    Dim sChar As String
    Dim sNewCellText As String
    
    If UBound(vTable) = -1 Then
        Exit Sub
    End If
    
    If lRow = -1 Then
        lRow = 0
    End If
    For lColumnIndex = LBound(vTable(lRow)) To UBound(vTable(lRow))
        sCellText = vTable(lRow)(lColumnIndex)
        sNewCellText = ""
        bUpper = True
            
        For lIndex = 1 To Len(sCellText)
            sChar = Mid$(sCellText, lIndex, 1)
            
            If sChar = "_" Then
                bUpper = True
            Else
                If bUpper Then
                    sNewCellText = sNewCellText & UCase$(sChar)
                    If LCase$(sChar) <> UCase$(sChar) Then
                        bUpper = False
                    End If
                Else
                    sNewCellText = sNewCellText & sChar
                End If
            End If
        Next
        
        vTable(lRow)(lColumnIndex) = sNewCellText
    Next
End Sub

Private Sub ChangeToLCUnderscoredRow(vTable As Variant, ByVal lRow As Long)
    Dim lColumnIndex As Long
    Dim bUnderscore As Boolean
    Dim sCellText As String
    Dim lIndex As Long
    Dim sChar As String
    Dim sChar2 As String
    
    Dim sNewCellText As String
    
    If UBound(vTable) = -1 Then
        Exit Sub
    End If
    
    If lRow = -1 Then
        lRow = 0
    End If
    For lColumnIndex = LBound(vTable(lRow)) To UBound(vTable(lRow))
        sCellText = vTable(lRow)(lColumnIndex)
        sNewCellText = ""
        bUnderscore = False
            
        For lIndex = 1 To Len(sCellText)
            sChar = Mid$(sCellText, lIndex, 1)
            sChar2 = Mid$(sCellText, lIndex + 1, 1)
            
            If (UCase$(sChar) = sChar And LCase$(sChar) <> sChar) And LCase$(sChar2) = sChar2 And lIndex <> 1 Then
                bUnderscore = True
            Else
                bUnderscore = False
            End If
            
            sNewCellText = sNewCellText & IIf(bUnderscore, "_", "") & LCase$(sChar)
        Next
        
        vTable(lRow)(lColumnIndex) = sNewCellText
    Next
End Sub

Private Sub ChangeToSpaceCaseRow(vTable As Variant, ByVal lRow As Long)
    Dim lColumnIndex As Long
    Dim bSpace As Boolean
    Dim sCellText As String
    Dim lIndex As Long
    Dim sChar As String
    Dim sChar2 As String
    
    Dim sNewCellText As String
    
    If UBound(vTable) = -1 Then
        Exit Sub
    End If
    
    If lRow = -1 Then
        lRow = 0
    End If
    For lColumnIndex = LBound(vTable(lRow)) To UBound(vTable(lRow))
        sCellText = vTable(lRow)(lColumnIndex)
        sNewCellText = ""
        bSpace = False
            
        For lIndex = 1 To Len(sCellText)
            sChar = Mid$(sCellText, lIndex, 1)
            sChar2 = Mid$(sCellText, lIndex + 1, 1)
            
            If (UCase$(sChar) = sChar And LCase$(sChar) <> sChar) And LCase$(sChar2) = sChar2 And lIndex <> 1 Then
                bSpace = True
            Else
                bSpace = False
            End If
            
            sNewCellText = sNewCellText & IIf(bSpace, " ", "") & sChar
        Next
        
        vTable(lRow)(lColumnIndex) = sNewCellText
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

Private Function SortByColumn(ByVal vTable As Variant, ByVal lColumn As Long, Optional ByVal bDescending As Boolean, Optional ByVal bIntelligentCompare As Boolean) As Variant
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
    
    SortByColumn = vTable
End Function

