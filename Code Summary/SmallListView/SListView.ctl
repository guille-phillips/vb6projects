VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.UserControl SListView 
   ClientHeight    =   4695
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5070
   ScaleHeight     =   4695
   ScaleWidth      =   5070
   Begin VB.PictureBox PicContainer 
      BackColor       =   &H80000005&
      Height          =   4335
      Left            =   120
      ScaleHeight     =   4275
      ScaleWidth      =   4635
      TabIndex        =   0
      Top             =   120
      Width           =   4695
      Begin VB.HScrollBar Scroll 
         Height          =   255
         Left            =   0
         TabIndex        =   2
         Top             =   3960
         Width           =   4215
      End
      Begin VB.PictureBox picList 
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H00E0E0E0&
         Height          =   3015
         Left            =   0
         OLEDropMode     =   1  'Manual
         ScaleHeight     =   3015
         ScaleWidth      =   4455
         TabIndex        =   1
         Top             =   0
         Width           =   4455
      End
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   4680
      Top             =   3000
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   5
      ImageHeight     =   20
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   1
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "SListView.ctx":0000
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "SListView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private ListItems As New ListCollection
Private CurrentHighlight As Long
Private LastHighlight As Long
Private ColumnWidth As Single
Private RowHeight As Single
Private Const Margin = 200
Private Const HighlightColour = &HC0E0FF ' &HE0E0E0

Event DblClick()
Event Click()
Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Event OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)
Event OLEStartDrag(Data As DataObject, AllowedEffects As Long)
Event OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)

Public IconPicture As StdPicture

Private PageLineCount As Long

Public Function HitTest(x As Single, y As Single) As Object
    Dim xpos As Single
    Dim ypos As Single
    Dim index As Long
    
    xpos = x \ ColumnWidth
    ypos = y \ RowHeight
    
    index = 1 + xpos * PageLineCount + ypos
    If index > 0 And index <= ListItems.Count Then
        Set HitTest = ListItems(index)
    End If
End Function


Public Sub SelectItem(x As Single, y As Single)
    Dim xpos As Single
    Dim ypos As Single
    Dim index As Long
    
    xpos = x \ ColumnWidth
    ypos = y \ RowHeight
    
    index = 1 + xpos * PageLineCount + ypos
    Highlight = index
End Sub

Public Sub OLEDrag()
    picList.OLEDrag
End Sub

Public Property Get SelectedItem() As Object
    If CurrentHighlight <> 0 Then
        Set SelectedItem = ListItems(CurrentHighlight)
    End If
End Property

Public Sub Add(Text As String, Bold As Boolean, Color As Long, Optional Key As String, Optional Tag As String, Optional Image As Boolean)
    Dim aListItem As ListItem
    Set aListItem = New ListItem
    
    With aListItem
        .Text = Text
        .Bold = Bold
        .Color = Color
        .Tag = Tag
        .Key = Key
        .Image = Image
    End With
    ListItems.Add aListItem
End Sub

Public Property Get Highlight() As Long
    Highlight = CurrentHighlight
End Property

Public Property Let Highlight(index As Long)
    If index <= ListItems.Count Then
        If index <> CurrentHighlight Then
            DisplayItem CurrentHighlight
            DisplayItem index, HighlightColour
            CurrentHighlight = index
        End If
    Else
        DisplayItem CurrentHighlight
        CurrentHighlight = 0
    End If
End Property

Private Sub Display()
    Dim index As Long
    Dim ypos As Single
    Dim xpos As Single
    Dim maxcount As Long
    Dim maxwidth As Single
    Dim extra As Single
    
    maxwidth = ((ListItems.Count \ PageLineCount) + 1) * ColumnWidth
    If maxwidth < PicContainer.Width Then
        picList.Width = PicContainer.Width
    Else
        picList.Width = maxwidth
    End If
    
    If picList.Width > PicContainer.Width Then
        Scroll.Visible = True
        Scroll.Max = (ListItems.Count \ PageLineCount) * 40
        Scroll.SmallChange = 40
        Scroll.LargeChange = 40
    Else
        Scroll.Visible = False
    End If
    
    For index = 1 To ListItems.Count
        DisplayItem (index)
    Next
    
    maxcount = ((picList.Height - Scroll.Height) \ RowHeight) * (picList.Width \ ColumnWidth + 1)
    
    For index = ListItems.Count + 1 To maxcount - 1
        xpos = ((index - 1) \ PageLineCount) * ColumnWidth
        ypos = ((index - 1) Mod PageLineCount) * RowHeight
        picList.Line (xpos, ypos)-Step(ColumnWidth, RowHeight), vbWhite, BF
    Next
    
    extra = (PicContainer.ScaleHeight - Scroll.Height) / RowHeight
    extra = (extra - Int(extra)) * RowHeight
    picList.Line (0, PicContainer.ScaleHeight - Scroll.Height - extra)-Step(Scroll.Width, Scroll.Height + extra), vbWindowBackground, BF
End Sub


Private Sub picList_Click()
    RaiseEvent Click
End Sub

Private Sub picList_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub picList_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim xpos As Long
    Dim ypos As Long
    Dim index As Long
    
    If Button = vbLeftButton Then
        xpos = x \ ColumnWidth
        ypos = y \ RowHeight
        
        index = 1 + xpos * ((PicContainer.ScaleHeight - Scroll.Height) \ RowHeight) + ypos
        Highlight = index
    End If
    
    RaiseEvent MouseDown(Button, Shift, x, y)
End Sub

Private Sub DisplayItem(index As Long, Optional ByVal Highlight As Long = vbWhite)
    Dim xpos As Single
    Dim ypos As Single
    
    xpos = ((index - 1) \ PageLineCount) * ColumnWidth
    ypos = ((index - 1) Mod PageLineCount) * RowHeight
    picList.Line (xpos, ypos)-Step(ColumnWidth, RowHeight), RGB(255, 255, 255), BF

    If index <= ListItems.Count And index > 0 Then
        picList.ForeColor = ListItems(index).Color
        picList.FontBold = ListItems(index).Bold
        
        picList.Line (xpos + 3 * Margin / 4, ypos)-Step(picList.TextWidth(ListItems(index).Text) + Margin / 4, RowHeight), Highlight, BF
        picList.Line (xpos + Margin, ypos)-Step(1, 0)
        picList.Print ListItems.Item(index).Text
        If ListItems.Item(index).Image Then
            picList.PaintPicture IconPicture, xpos + Margin - IconPicture.Width, ypos
        End If
    End If
End Sub

Private Sub SetColumnWidth()
    Dim x As Long
    Dim w As Single
    
    ColumnWidth = 0
    For x = 1 To ListItems.Count
        picList.FontBold = ListItems(x).Bold
        w = picList.TextWidth(ListItems(x))
        If w > ColumnWidth Then
            ColumnWidth = w
        End If
    Next
    ColumnWidth = ColumnWidth + Margin
End Sub

Private Sub picList_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseMove(Button, Shift, x, y)
End Sub

Private Sub picList_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent OLEDragDrop(Data, Effect, Button, Shift, x, y)
End Sub

Private Sub picList_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, x As Single, y As Single, State As Integer)
    RaiseEvent OLEDragOver(Data, Effect, Button, Shift, x, y, State)
End Sub

Private Sub picList_OLEStartDrag(Data As DataObject, AllowedEffects As Long)
    RaiseEvent OLEStartDrag(Data, AllowedEffects)
End Sub

Private Sub Scroll_Change()
    picList.Left = -(Scroll.Value / 40) * ColumnWidth
End Sub

Private Sub Scroll_Scroll()
    picList.Left = -(Scroll.Value / 40) * ColumnWidth
End Sub

Private Sub UserControl_Initialize()
    RowHeight = picList.TextHeight("T")
    ColumnWidth = 1000
    'Set IconPicture = LoadPicture("c:\projects\smalllistview\public3.bmp")
    Set IconPicture = ImageList1.ListImages(1).Picture
End Sub

Public Sub ShowList()
    SetColumnWidth
    Display
    DisplayItem CurrentHighlight, HighlightColour
End Sub

Public Sub ClearList()
    Set ListItems = Nothing
    Set ListItems = New ListCollection
    CurrentHighlight = 0
End Sub

Private Sub UserControl_Resize()
    PicContainer.Move 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight
    picList.Move 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight

    Scroll.Width = PicContainer.ScaleWidth
    Scroll.Top = PicContainer.ScaleHeight - Scroll.Height
    
    PageLineCount = (PicContainer.ScaleHeight - Scroll.Height) \ RowHeight
    Display
    DisplayItem CurrentHighlight, HighlightColour
End Sub
