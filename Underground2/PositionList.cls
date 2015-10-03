VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "PositionList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public List As New Collection

Public Function FindPosition(X As Single, Y As Single) As Position
    Dim oPosition As Position
    Dim lPositionIndex As Long
    
    For lPositionIndex = List.Count To 1 Step -1
        Set oPosition = List(lPositionIndex)
        If X >= (oPosition.Pos.X - oPosition.Radius) And X <= (oPosition.Pos.X + oPosition.Radius) Then
            If Y >= (oPosition.Pos.Y - oPosition.Radius) And Y <= (oPosition.Pos.Y + oPosition.Radius) Then
                Set FindPosition = oPosition
                Exit For
            End If
        End If
    Next
End Function

Public Function FindPositions(X1 As Single, Y1 As Single, X2 As Single, Y2 As Single) As Collection
    Dim oPosition As Position
    Dim MidX As Single
    Dim MidY As Single
    Dim WidthX As Single
    Dim WidthY As Single
    Dim nTemp As Single
    Dim lPositionIndex As Long
    
    If X2 < X1 Then
        nTemp = X1
        X1 = X2
        X2 = nTemp
    End If
    
    If Y2 < Y1 Then
        nTemp = Y1
        Y1 = Y2
        Y2 = nTemp
    End If
    
    Set FindPositions = New Collection
    
    For lPositionIndex = List.Count To 1 Step -1
        Set oPosition = List(lPositionIndex)
        If X1 <= oPosition.Pos.X And X2 >= oPosition.Pos.X Then
            If Y1 <= oPosition.Pos.Y And Y2 >= oPosition.Pos.Y Then
                    FindPositions.Add oPosition
            End If
        End If
    Next
End Function

Public Function FindPositionFromReference(sRef As String) As Position
    Dim oPosition As Position
    
    For Each oPosition In List
        If oPosition.Reference = sRef Then
            Set FindPositionFromReference = oPosition
            Exit For
        End If
    Next
End Function

Public Sub RemovePosition(oPosition As Position)
    Dim iIndex As Long
    
    For iIndex = 1 To List.Count
        If List(iIndex) Is oPosition Then
            List.Remove iIndex
            Exit Sub
        End If
    Next
End Sub

Public Sub RenderAll()
    Dim oPosition As Position
    
    For Each oPosition In List
        oPosition.RenderName
    Next
End Sub

Public Sub SendToFront(ByVal oPosition As Position)
    Dim lIndex As Long
    
    For lIndex = 1 To List.Count
        If List(lIndex).Reference = oPosition.Reference Then
            List.Remove (lIndex)
            List.Add oPosition
        End If
    Next
End Sub

Public Sub SendToBack(ByVal oPosition As Position)
    Dim lIndex As Long
    
    For lIndex = 1 To List.Count
        If List(lIndex).Reference = oPosition.Reference Then
            List.Remove (lIndex)
            List.Add oPosition, , 1
        End If
    Next
End Sub
