Attribute VB_Name = "modGraph"
Option Explicit

Private Declare Function DrawText Lib "user32.dll" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, ByRef lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32.dll" (ByVal crColor As Long) As Long
Private Declare Function FillRect Lib "user32.dll" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
Private Declare Function GetTextExtentPoint32 Lib "gdi32" Alias "GetTextExtentPoint32A" (ByVal hdc As Long, ByVal lpsz As String, ByVal cbString As Long, lpSize As SIZE) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function DrawFocusRect Lib "user32.dll" (ByVal hdc As Long, lpRect As RECT) As Long

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

Private Enum btBoxTypes
    btNote
End Enum

Private Type BoxInfo
    Top As Long
    Right As Long
    Bottom As Long
    Left As Long
    BackColour As Long
    ForeColour As Long
    Text As String
    Focus As Boolean
    BoxType As btBoxTypes
    Column As Long
    Row As Long
    Centred As Boolean
    Index As Long
End Type

Private mbiBoxes() As BoxInfo
Private mlBoxesCount As Long

Private Type NoteGraphType
    NoteRef As Long
    Column As Long
    Row As Long
End Type

Private Function AddBox(ByVal lX As Long, ByVal lY As Long, ByVal lWidth As Long, ByVal lHeight As Long, ByVal lBackColour As Long, ByVal lForeColour As Long, ByVal sText As String, ByVal btBoxType As btBoxTypes, ByVal lColumn As Long, ByVal lRow As Long, Optional ByVal bCentre As Boolean) As Long
    ReDim Preserve mbiBoxes(mlBoxesCount)
    
    With mbiBoxes(mlBoxesCount)
        .Left = lX
        .Right = lX + lWidth
        .Top = lY
        .Bottom = lY + lHeight
        .BackColour = lBackColour
        .ForeColour = lForeColour
        .Text = sText
        .BoxType = btBoxType
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
'        SetTextColor Me.hdc, .ForeColour
                
        rectArea.Top = .Top
        rectArea.Left = .Left
        rectArea.Bottom = .Bottom
        rectArea.Right = .Right
        
        RenderBox = .Right + 1
        
'        FillRect Me.hdc, rectArea, hBrush
        If bFocus Then
'            DrawFocusRect Me.hdc, rectArea
        End If
'        DrawText Me.hdc, .Text, Len(.Text), rectArea, DT_SINGLELINE Or DT_VCENTER Or DT_NOPREFIX Or (DT_CENTER * -bCentre)
        
        DeleteObject hBrush
    End With
End Function

Private Function ConvertSequence()
    Dim lIndex As Long
    
    
End Function
