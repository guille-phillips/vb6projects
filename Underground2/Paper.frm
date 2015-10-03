VERSION 5.00
Begin VB.Form Paper 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   Caption         =   "Relationships"
   ClientHeight    =   8760
   ClientLeft      =   165
   ClientTop       =   255
   ClientWidth     =   11865
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "Arial Black"
      Size            =   18
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FFFFFF&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   584
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   791
   WindowState     =   2  'Maximized
   Begin VB.HScrollBar scrColourComponent 
      Height          =   255
      Index           =   2
      Left            =   240
      Max             =   255
      TabIndex        =   2
      Top             =   600
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.HScrollBar scrColourComponent 
      Height          =   255
      Index           =   1
      Left            =   240
      Max             =   255
      TabIndex        =   1
      Top             =   360
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.HScrollBar scrColourComponent 
      Height          =   255
      Index           =   0
      Left            =   240
      Max             =   255
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   1575
   End
End
Attribute VB_Name = "Paper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)


Private mnMouseX As Single
Private mnMouseY As Single

Private oSelectedPositions As Collection

Private oSelectedPosition As Position
Private oSelectedPosOffsetX As Single
Private oSelectedPosOffsetY As Single

Private nBoxStartX As Single
Private nBoxStartY As Single

Private oInitialLink As Position
Private oFinalLink As Position

Public BackColour As Long

Public DiagramRef As New Diagram

Private bDragGroupSelected As Boolean

Private mlRecentColourIndex As Long

Public mbGraticuleOn As Boolean
Public mbShowCircles As Boolean

Private Sub Form_Activate()
    Dim c As New CirclePrimitive
    Dim oCentre As New Vector
    Dim ostart As New Vector
    Dim X As Long
    Dim Y As Long
    Dim lPixelOn As Long
    
    oCentre.SetVector 20, 20
    c.Initialise oCentre, 5
    c.SetStart ostart
    
    For Y = 0 To 40
        For X = 0 To 40
            If Y Mod 2 = 0 Then
                lPixelOn = lPixelOn Xor c.MoveRight
            Else
                lPixelOn = lPixelOn Xor c.MoveLeft
            End If
            
        Next
        lPixelOn = lPixelOn Xor c.MoveDown
    Next
End Sub

Private Sub Form_Initialize()
    Set DiagramRef.PaperRef = Me
    DiagramRef.Colours = Array(RGB(0, 0, 0), RGB(255, 0, 0), RGB(0, 255, 0), RGB(255, 255, 0), RGB(0, 0, 255), RGB(255, 0, 255), RGB(0, 255, 255), RGB(255, 255, 255), RGB(128, 128, 128), &H80FF&, 16576, &H800080, RGB(255, 128, 255), RGB(255, 255, 255), RGB(255, 255, 255), RGB(255, 255, 255), RGB(255, 255, 255), RGB(255, 255, 255), RGB(255, 255, 255), RGB(255, 255, 255), RGB(255, 255, 255))
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim oPosition As Position
    Dim oRelationship As Relationship
    Dim vRelColours As Variant
    Dim lColourIndex As Long
    Dim lColour As Long
    Dim yColour(3) As Byte
    
    Select Case KeyCode
        Case vbKey1, vbKey2, vbKey3, vbKey4, vbKey5, vbKey6, vbKey7, vbKey8, vbKey9, vbKey0
            mlRecentColourIndex = (KeyCode - 48 - ((Shift And 1) = 1) * 10)

            lColour = DiagramRef.Colours(mlRecentColourIndex)
            CopyMemory yColour(0), lColour, &H4&
            
            scrColourComponent(0).Value = yColour(0)
            scrColourComponent(1).Value = yColour(1)
            scrColourComponent(2).Value = yColour(2)
            
            If oInitialLink Is Nothing Then
                Set oInitialLink = DiagramRef.Positions.FindPosition((mnMouseX - DiagramRef.TopLeft.X) / DiagramRef.Zoom, (mnMouseY - DiagramRef.TopLeft.Y) / DiagramRef.Zoom)
                If Not oInitialLink Is Nothing Then
                    If (Shift And 2) = 2 Then
                        oInitialLink.ColourIndex = KeyCode - 48 - ((Shift And 1) = 1) * 10
                        oInitialLink.RenderName
                        Set oInitialLink = Nothing
                    Else
                        oInitialLink.RenderName &HFFC0C0
                    End If
                Else
                    If (Shift And 2) = 0 Then
                        NewString.txtString = ""
                        NewString.Show vbModal
                        Set oPosition = New Position
                        oPosition.Name = NewString.txtString
                        oPosition.Pos.X = Int(((mnMouseX - DiagramRef.TopLeft.X) / DiagramRef.Zoom + 10) / 20) * 20
                        oPosition.Pos.Y = Int(((mnMouseY - DiagramRef.TopLeft.Y) / DiagramRef.Zoom + 10) / 20) * 20
                        oPosition.ColourIndex = (KeyCode - 48 - (Shift = 1) * 10)
                        Set oPosition.DiagramRef = DiagramRef
                        Set oPosition.ParserRef = oParsePosition
                        DiagramRef.Positions.List.Add oPosition
                        oPosition.RenderName
                        DiagramRef.FileIOs.WriteFile
                    End If
                End If
            Else
                Set oFinalLink = DiagramRef.Positions.FindPosition((mnMouseX - DiagramRef.TopLeft.X) / DiagramRef.Zoom, (mnMouseY - DiagramRef.TopLeft.Y) / DiagramRef.Zoom)
                If Not oFinalLink Is Nothing Then
                    If oFinalLink.Reference = oInitialLink.Reference Then
                        oInitialLink.RenderName
                        Set oInitialLink = Nothing
                        Set oFinalLink = Nothing
                    End If
                Else
                    oInitialLink.RenderName
                    Set oInitialLink = Nothing
                    Set oFinalLink = Nothing
                End If
                
                If Not oFinalLink Is Nothing Then
                    oFinalLink.RenderName &HFFC0C0
                    
                    Set oRelationship = DiagramRef.Relationships.FindRelationship(oInitialLink, oFinalLink)
                    If oRelationship Is Nothing Then
                        Set oRelationship = New Relationship
                        Set oRelationship.DiagramRef = DiagramRef
                        Set oRelationship.FromPos = oInitialLink
                        Set oRelationship.ToPos = oFinalLink
                        oRelationship.ColourIndeces = Array((KeyCode - 48 - (Shift = 1) * 10))
                        oRelationship.RenderRelationship
                        DiagramRef.Relationships.List.Add oRelationship
                        DiagramRef.Positions.RenderAll
                    Else
                        lColourIndex = (KeyCode - 48 - (Shift = 1) * 10)
                        vRelColours = oRelationship.ColourIndeces
                        If InArray(vRelColours, lColourIndex) Then
                            RemoveFromArray vRelColours, lColourIndex
                        Else
                            AddToArray vRelColours, lColourIndex
                        End If
                        oRelationship.ColourIndeces = vRelColours
                        If UBound(vRelColours) = -1 Then
                            DiagramRef.Relationships.RemoveRelationship oRelationship
                        End If
                        DiagramRef.RenderAndSave
                    End If
                    Set oInitialLink = Nothing
                    Set oFinalLink = Nothing
                    DiagramRef.FileIOs.WriteFile
                End If
            End If
        Case vbKeyB
            If oInitialLink Is Nothing Then
                Set oInitialLink = DiagramRef.Positions.FindPosition((mnMouseX - DiagramRef.TopLeft.X) / DiagramRef.Zoom, (mnMouseY - DiagramRef.TopLeft.Y) / DiagramRef.Zoom)
                If Not oInitialLink Is Nothing Then
                    oInitialLink.RenderName &HFFC0C0
                End If
            Else
                Set oFinalLink = DiagramRef.Positions.FindPosition((mnMouseX - DiagramRef.TopLeft.X) / DiagramRef.Zoom, (mnMouseY - DiagramRef.TopLeft.Y) / DiagramRef.Zoom)
                If Not oFinalLink Is Nothing Then
                    oFinalLink.RenderName &HFFC0C0
                    Set oRelationship = DiagramRef.Relationships.FindRelationship(oInitialLink, oFinalLink)
                    If Not oRelationship Is Nothing Then
                        oRelationship.Style = 1 - oRelationship.Style
                        DiagramRef.RenderAndSave
                        Set oInitialLink = Nothing
                        Set oFinalLink = Nothing
                    End If
                End If
            End If
        Case vbKeyL
            If oInitialLink Is Nothing Then
                Set oInitialLink = DiagramRef.Positions.FindPosition((mnMouseX - DiagramRef.TopLeft.X) / DiagramRef.Zoom, (mnMouseY - DiagramRef.TopLeft.Y) / DiagramRef.Zoom)
                If Not oInitialLink Is Nothing Then
                    oInitialLink.RenderName &HFFC0C0
                End If
            Else
                Set oFinalLink = DiagramRef.Positions.FindPosition((mnMouseX - DiagramRef.TopLeft.X) / DiagramRef.Zoom, (mnMouseY - DiagramRef.TopLeft.Y) / DiagramRef.Zoom)
                If Not oFinalLink Is Nothing Then
                    oFinalLink.RenderName &HFFC0C0
                    Set oRelationship = DiagramRef.Relationships.FindRelationship(oInitialLink, oFinalLink)
                    If Not oRelationship Is Nothing Then
                        oRelationship.Style = 2
                        DiagramRef.RenderAndSave
                        Set oInitialLink = Nothing
                        Set oFinalLink = Nothing
                    End If
                End If
            End If
        Case vbKeyR
            Set oPosition = DiagramRef.Positions.FindPosition((mnMouseX - DiagramRef.TopLeft.X) / DiagramRef.Zoom, (mnMouseY - DiagramRef.TopLeft.Y) / DiagramRef.Zoom)
            If Not oPosition Is Nothing Then
                oPosition.Orientation = (oPosition.Orientation + 2 * (Shift = 1) + 1 + 8) Mod 8
                DiagramRef.RenderAndSave
            End If
        Case vbKeyDelete
            Set oPosition = DiagramRef.Positions.FindPosition((mnMouseX - DiagramRef.TopLeft.X) / DiagramRef.Zoom, (mnMouseY - DiagramRef.TopLeft.Y) / DiagramRef.Zoom)
            If Not oPosition Is Nothing Then
                DiagramRef.Relationships.RemoveRelationshipWithReference oPosition
            End If
            DiagramRef.Positions.RemovePosition oPosition
            DiagramRef.RenderAndSave

        Case vbKeyE
            Set oPosition = DiagramRef.Positions.FindPosition((mnMouseX - DiagramRef.TopLeft.X) / DiagramRef.Zoom, (mnMouseY - DiagramRef.TopLeft.Y) / DiagramRef.Zoom)
            If Not oPosition Is Nothing Then
                oPosition.ClearName
                NewString.txtString = oPosition.Name
                NewString.Show vbModal
                oPosition.Name = NewString.txtString
                oPosition.RenderName
                DiagramRef.FileIOs.WriteFile
            End If
        Case vbKeyG
            mbGraticuleOn = Not mbGraticuleOn
            DiagramRef.RenderAndSave
        Case vbKeyAdd
            DiagramRef.Zoom = DiagramRef.Zoom * 4 / 3
            DiagramRef.RenderAndSave
        Case vbKeySubtract
            DiagramRef.Zoom = Int(DiagramRef.Zoom * 3 + 0.5) / 4
            DiagramRef.RenderAndSave
        Case vbKeyS
            mbShowCircles = Not mbShowCircles
            DiagramRef.Relationships.RemoveDuplicates
            DiagramRef.RenderAndSave
        Case vbKeyF
            Set oPosition = DiagramRef.Positions.FindPosition((mnMouseX - DiagramRef.TopLeft.X) / DiagramRef.Zoom, (mnMouseY - DiagramRef.TopLeft.Y) / DiagramRef.Zoom)
            If Not oPosition Is Nothing Then
                If Shift <> 2 Then
                    DiagramRef.Positions.SendToFront oPosition
                    DiagramRef.Relationships.SendToFront oPosition
                Else
                    DiagramRef.Positions.SendToBack oPosition
                    DiagramRef.Relationships.SendToBack oPosition
                End If
                DiagramRef.RenderAndSave
            End If
        Case vbKeyC
            scrColourComponent(0).Visible = Not scrColourComponent(0).Visible
            scrColourComponent(1).Visible = Not scrColourComponent(1).Visible
            scrColourComponent(2).Visible = Not scrColourComponent(2).Visible
    End Select
End Sub

Private Sub Form_Load()
    BackColour = Me.BackColor
    
    DiagramRef.RenderAndSave
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim oPosition As Position
    
    If Button = vbLeftButton Then
        Set oSelectedPosition = DiagramRef.Positions.FindPosition((X - DiagramRef.TopLeft.X) / DiagramRef.Zoom, (Y - DiagramRef.TopLeft.Y) / DiagramRef.Zoom)
        If Not oSelectedPosition Is Nothing Then
            oSelectedPosOffsetX = (oSelectedPosition.Pos.X * DiagramRef.Zoom + DiagramRef.TopLeft.X) - X
            oSelectedPosOffsetY = (oSelectedPosition.Pos.Y * DiagramRef.Zoom + DiagramRef.TopLeft.Y) - Y
            If Not oSelectedPositions Is Nothing Then
                bDragGroupSelected = False
                For Each oPosition In oSelectedPositions
                    If oPosition Is oSelectedPosition Then
                        bDragGroupSelected = True
                    End If
                Next
            End If
        Else
            bDragGroupSelected = False
            If Shift = 0 Then
                Set oSelectedPositions = Nothing
                Set oInitialLink = Nothing
                Set oFinalLink = Nothing
                DiagramRef.Positions.RenderAll
            Else
                nBoxStartX = X
                nBoxStartY = Y
            End If
        End If
    End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Static PreviousX As Single
    Static PreviousY As Single
    Dim oPositionA As Position
    Dim nSelectedX As Single
    Dim nSelectedY As Single
    
    mnMouseX = X
    mnMouseY = Y
    If Button = vbLeftButton Then
        If Not oSelectedPosition Is Nothing Then
            If bDragGroupSelected Then
                If Not oSelectedPositions Is Nothing Then
                    nSelectedX = oSelectedPosition.Pos.X
                    nSelectedY = oSelectedPosition.Pos.Y
                    For Each oPositionA In oSelectedPositions
                        oPositionA.ClearName
                        oPositionA.Pos.X = oPositionA.Pos.X + ((mnMouseX - DiagramRef.TopLeft.X) / DiagramRef.Zoom + oSelectedPosOffsetX - nSelectedX)
                        oPositionA.Pos.Y = oPositionA.Pos.Y + ((mnMouseY - DiagramRef.TopLeft.Y) / DiagramRef.Zoom + oSelectedPosOffsetY - nSelectedY)
                        oPositionA.RenderName &HFFC0C0
                    Next
                End If
            Else
                oSelectedPosition.ClearName
                oSelectedPosition.Pos.X = (mnMouseX - DiagramRef.TopLeft.X) / DiagramRef.Zoom + oSelectedPosOffsetX
                oSelectedPosition.Pos.Y = (mnMouseY - DiagramRef.TopLeft.Y) / DiagramRef.Zoom + oSelectedPosOffsetY
                oSelectedPosition.RenderName &HFFC0C0
            End If
        Else
            If Shift = 0 Then
                If PreviousX <> 0 Or PreviousY <> 0 Then
                    DiagramRef.TopLeft.X = DiagramRef.TopLeft.X - (PreviousX - X)
                    DiagramRef.TopLeft.Y = DiagramRef.TopLeft.Y - (PreviousY - Y)
                    DiagramRef.RenderAndSave
                End If
            Else
                Me.DrawMode = vbXorPen
                Me.DrawStyle = vbDash
                Me.FillStyle = 1
                Me.Line (nBoxStartX, nBoxStartY)-(PreviousX, PreviousY), , B
                Me.Line (nBoxStartX, nBoxStartY)-(X, Y), , B

                Me.DrawMode = vbCopyPen
                Me.DrawStyle = vbSolid
                
                Set oSelectedPositions = DiagramRef.Positions.FindPositions((nBoxStartX - DiagramRef.TopLeft.X) / DiagramRef.Zoom, (nBoxStartY - DiagramRef.TopLeft.Y) / DiagramRef.Zoom, (X - DiagramRef.TopLeft.X) / DiagramRef.Zoom, (Y - DiagramRef.TopLeft.Y) / DiagramRef.Zoom)
                For Each oPositionA In oSelectedPositions
                    oPositionA.RenderName &HFFC0C0
                Next
            End If
        End If
    End If
    PreviousX = X
    PreviousY = Y
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim oPosition As Position
    
    If Not oSelectedPosition Is Nothing Then
        If Shift <> 1 Then
            oSelectedPosition.Pos.X = Int((oSelectedPosition.Pos.X + 10) / 20) * 20
            oSelectedPosition.Pos.Y = Int((oSelectedPosition.Pos.Y + 10) / 20) * 20
                    
            If Not oSelectedPositions Is Nothing Then
                For Each oPosition In oSelectedPositions
                    oPosition.Pos.X = Int((oPosition.Pos.X + 10) / 20) * 20
                    oPosition.Pos.Y = Int((oPosition.Pos.Y + 10) / 20) * 20
                Next
            End If
        End If
        
        If Shift = 2 Then
            If oSelectedPositions Is Nothing Then
                Set oSelectedPositions = New Collection
            End If
            
            oSelectedPositions.Add oSelectedPosition
            oSelectedPosition.RenderName &HFFC0C0
        Else
            Set oSelectedPosition = Nothing
        End If
        
        DiagramRef.RenderAndSave

        If Not oSelectedPositions Is Nothing Then
            For Each oPosition In oSelectedPositions
                oPosition.RenderName &HFFC0C0
            Next
        End If
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    DiagramRef.FileIOs.WriteFile
End Sub

Private Sub scrColourComponent_Change(Index As Integer)
    Dim yColour(3) As Byte
    Dim vColours As Variant
    Dim lColour As Long
    
    yColour(0) = scrColourComponent(0).Value
    yColour(1) = scrColourComponent(1).Value
    yColour(2) = scrColourComponent(2).Value
    
    CopyMemory lColour, yColour(0), &H4&

    vColours = DiagramRef.Colours
    vColours(mlRecentColourIndex) = lColour
    DiagramRef.Colours = vColours
    DiagramRef.RenderAndSave
End Sub

Public Sub Watermark()
    Dim nX As Single
    Dim nY As Single
    
    Me.ForeColor = DiagramRef.Colours(17)
    For nX = 0 To 1600 Step 160
        For nY = 0 To 1200 Step 15
            CurrentX = nX + ((nY \ 15) Mod 2) * 80
            CurrentY = nY
            
            Print "Copyright " & Chr$(169) & " Guillermo Phillips 2007"
            Print
        Next
    Next
    
End Sub
