VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Graphics"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function Polygon Lib "gdi32" (ByVal hdc As Long, lpPoint As POINTAPI, ByVal nCount As Long) As Long
Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
Private Declare Function SelectObject Lib "gdi32.dll" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long

Private Type POINTAPI
   X As Long
   Y As Long
End Type

Public DiagramRef As Diagram

Public Sub StripedLine(ostart As Vector, oEnd As Vector, vColours As Variant)
    Dim lTotalColours As Long
    Dim lColourIndex As Long
    Dim dStep As Double
    Dim oPosition As New Vector
    Dim oDX As Vector
    Dim dLength As Long
    
    dStep = 25 * DiagramRef.Zoom
    lTotalColours = UBound(vColours) + 1
    dLength = oEnd.Subs(ostart).Size
    oPosition.Copy ostart
    Set oDX = oEnd.Subs(ostart).Normal.Scalar(dStep)
    While dLength >= dStep
        FatLine oPosition, oPosition.Add(oDX), LineWidth / 2 * DiagramRef.Zoom, DiagramRef.Colours(lColourIndex)
        Set oPosition = oPosition.Add(oDX)
        lColourIndex = (lColourIndex + 1) Mod lTotalColours
        dLength = dLength - dStep
    Wend
    FatLine oPosition, oEnd, LineWidth / 2 * DiagramRef.Zoom, vColours(lColourIndex)
End Sub

Public Sub NormalLine(ostart As Vector, oEnd As Vector, ByVal lColour As Long)
    DiagramRef.PaperRef.DrawWidth = LineWidth * DiagramRef.Zoom
    DiagramRef.PaperRef.Line (ostart.X, ostart.Y)-(oEnd.X, oEnd.Y), lColour
End Sub

Public Sub FatLine(ostart As Vector, oEnd As Vector, dWidth As Double, ByVal lColour As Long)
    Dim hPen As Long
    Dim hBrush As Long
    Dim hOldPen As Long
    Dim hOldBrush As Long
    Dim paPoints(3) As POINTAPI
    Dim oPerp As Vector
    
    hPen = CreatePen(0, 0, lColour)
    hBrush = CreateSolidBrush(lColour)
    
    hOldPen = SelectObject(DiagramRef.PaperRef.hdc, hPen)
    hOldBrush = SelectObject(DiagramRef.PaperRef.hdc, hBrush)
    
    If oEnd.Subs(ostart).Size = 0 Then
        Exit Sub
    End If
    Set oPerp = oEnd.Subs(ostart).Perpendicular.Normal.Scalar(dWidth)
    
    paPoints(0).X = ostart.Add(oPerp).X
    paPoints(0).Y = ostart.Add(oPerp).Y
    paPoints(1).X = ostart.Add(oPerp.Scalar(-1)).X
    paPoints(1).Y = ostart.Add(oPerp.Scalar(-1)).Y
    paPoints(2).X = oEnd.Add(oPerp.Scalar(-1)).X
    paPoints(2).Y = oEnd.Add(oPerp.Scalar(-1)).Y
    paPoints(3).X = oEnd.Add(oPerp).X
    paPoints(3).Y = oEnd.Add(oPerp).Y

    Polygon DiagramRef.PaperRef.hdc, paPoints(0), UBound(paPoints) + 1
    
    Call SelectObject(DiagramRef.PaperRef.hdc, hOldPen)
    Call SelectObject(DiagramRef.PaperRef.hdc, hOldBrush)
    DeleteObject hPen
    DeleteObject hBrush
End Sub

Public Sub Graticule()
    Dim nX As Single
    Dim nY As Single
    Dim nOffsetX As Single
    Dim nOffsetY As Single
    
    If Not DiagramRef.PaperRef.mbGraticuleOn Then
        Exit Sub
    End If
    
    nOffsetX = DiagramRef.TopLeft.X - Int(DiagramRef.TopLeft.X / (20 * DiagramRef.Zoom)) * (20 * DiagramRef.Zoom)
    nOffsetY = DiagramRef.TopLeft.Y - Int(DiagramRef.TopLeft.Y / (20 * DiagramRef.Zoom)) * (20 * DiagramRef.Zoom)
    
    DiagramRef.PaperRef.DrawStyle = vbDot
    For nX = 0 To 1600 Step 20 * DiagramRef.Zoom
        DiagramRef.PaperRef.Line (nX + nOffsetX, 0)-Step(0, 1024), vbWhite
    Next
    
    For nY = 0 To 1200 Step 20 * DiagramRef.Zoom
        DiagramRef.PaperRef.Line (0, nY + nOffsetY)-Step(1280, 0), vbWhite
    Next
    
    DiagramRef.PaperRef.DrawStyle = vbSolid
    DiagramRef.PaperRef.Line (DiagramRef.TopLeft.X - 20 * DiagramRef.Zoom, DiagramRef.TopLeft.Y)-Step(40 * DiagramRef.Zoom, 0), vbRed
    DiagramRef.PaperRef.Line (DiagramRef.TopLeft.X, DiagramRef.TopLeft.Y - 20 * DiagramRef.Zoom)-Step(0, 40 * DiagramRef.Zoom), vbRed
End Sub


Public Sub PartialCircle(ByVal oCentre As Vector, nRadius As Single, ByVal lColour As Long, ByVal oPos1 As Vector, ByVal oPos2 As Vector, ByVal bSwap As Boolean)
    Dim nAngle1 As Single
    Dim nAngle2 As Single
    
    nAngle1 = pi2 - Atn2(oPos1.X - oCentre.X, oPos1.Y - oCentre.Y)
    nAngle2 = pi2 - Atn2(oPos2.X - oCentre.X, oPos2.Y - oCentre.Y)
    
    
    DiagramRef.PaperRef.FillStyle = vbFSTransparent
    DiagramRef.PaperRef.DrawWidth = LineWidth * DiagramRef.Zoom
    
    If Not bSwap Then
        DiagramRef.PaperRef.Circle (oCentre.X, oCentre.Y), nRadius, lColour, nAngle1, nAngle2
    Else
        DiagramRef.PaperRef.Circle (oCentre.X, oCentre.Y), nRadius, lColour, nAngle2, nAngle1
    End If
End Sub

Public Sub FilledCircle(ByVal oCentre As Vector, ByVal nOuterRadius As Single, ByVal nInnerRadius As Single, ByVal lOuterColour As Long, ByVal lInnerColour As Long)
    DiagramRef.PaperRef.FillStyle = vbFSSolid
    DiagramRef.PaperRef.DrawWidth = 1
    
    DiagramRef.PaperRef.FillColor = lOuterColour
    DiagramRef.PaperRef.Circle (oCentre.X, oCentre.Y), nOuterRadius
    DiagramRef.PaperRef.FillColor = lInnerColour
    DiagramRef.PaperRef.Circle (oCentre.X, oCentre.Y), nInnerRadius
End Sub
    

Private Function Atn2(ByVal X As Double, ByVal Y As Double) As Double
    Dim lSegment As Long
    Dim vOffset As Variant
    
    If X = 0 Then
        Atn2 = (2 - Sgn(Y)) * pi2 / 4
        Exit Function
    End If
    
    vOffset = Array(pi2 / 2, pi2 / 2, pi2, pi2 / 2, 0, 0, pi2 / 2, 0, 0)
    
    lSegment = Sgn(X) + 1 + 3 * (Sgn(Y) + 1)
    Atn2 = Atn(Y / X) + CDbl(vOffset(lSegment))
End Function

