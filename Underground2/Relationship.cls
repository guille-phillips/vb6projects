VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Relationship"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public FromPos As Position
Public ToPos As Position

Public DiagramRef As Diagram

Public ColourIndeces As Variant
Public Description As String
Public Style As Long

Public Sub RenderRelationship()
    DrawRelationship
End Sub

Public Sub ClearRelationship()
    DrawRelationship DiagramRef.PaperRef.BackColour
End Sub

Public Function WriteDef() As String
    WriteDef = FromPos.Reference & "|" & ToPos.Reference
End Function

Private Sub DrawRelationship(Optional lForeColour As Long = vbBlack)
    Dim ostart As New Vector
    Dim oEnd As New Vector
    Dim oVector As New Vector
    
    Dim oFromPosPrevious As Position
    Dim oToPosNext As Position
    
    Dim oFromRelations As Collection
    Dim oToRelations As Collection
    Dim oFromRelation As Relationship
    Dim oToRelation As Relationship
    
    Dim oRelation As Relationship
    Dim lIndex As Long
    
    Dim oFromVector As Vector
    Dim oToVector As Vector
    Dim oCentreFrom As Vector
    Dim oCentreTo As Vector
    
    Dim oDirection As Vector
    Dim oCentres As Vector
    Dim oStartLine As Vector
    
    Dim vColours As Variant
    
    Dim oLineGap As Vector
    Dim oStartGap As Vector
    Dim oEndGap As Vector
    Dim dLineSize As Long
    
    Const nCircleRadius As Single = 30
    
    If lForeColour = 0 Then
        lForeColour = DiagramRef.Colours(ColourIndeces(0))
    End If

    ostart.SetVector FromPos.Pos.X * DiagramRef.Zoom + DiagramRef.TopLeft.X, FromPos.Pos.Y * DiagramRef.Zoom + DiagramRef.TopLeft.Y
    oEnd.SetVector ToPos.Pos.X * DiagramRef.Zoom + DiagramRef.TopLeft.X, ToPos.Pos.Y * DiagramRef.Zoom + DiagramRef.TopLeft.Y
    Set oVector = oEnd.Subs(ostart)
        
    If Not ostart.Equal(oEnd) Then
        Select Case Style
            Case 0 ' normal
                If UBound(ColourIndeces) = 0 Then
                    vColours = Array(lForeColour)
                Else
                    vColours = ColourIndeces
                End If
                Set oLineGap = oEnd.Subs(ostart).Normal.Scalar(LineGap * DiagramRef.Zoom)
                dLineSize = oEnd.Subs(ostart).Size
                
                'If oEnd.Subs(oStart).Size > (2 * LineGap) And Not PaperRef.mbShowCircles Then
                If FromPos.Name <> "" And Not DiagramRef.PaperRef.mbShowCircles Then
                    Set oStartGap = ostart.Add(oLineGap)
                    dLineSize = dLineSize - LineGap * DiagramRef.Zoom
                Else
                    Set oStartGap = ostart
                End If
                If ToPos.Name <> "" And Not DiagramRef.PaperRef.mbShowCircles Then
                    Set oEndGap = oEnd.Subs(oLineGap)
                    dLineSize = dLineSize - LineGap * DiagramRef.Zoom
                Else
                    Set oEndGap = oEnd
                End If
            
                'PaperRef.StripedLine oStartGap, oEndGap, vColours
                
                If dLineSize > 0 Or DiagramRef.PaperRef.mbShowCircles Then
                    DiagramRef.PaperRef.DrawWidth = LineWidth * DiagramRef.Zoom
                    DiagramRef.PaperRef.Line (oStartGap.X, oStartGap.Y)-(oEndGap.X, oEndGap.Y), DiagramRef.Colours(ColourIndeces(0))
                End If
            Case 1 ' bend
                Set oFromRelations = DiagramRef.Relationships.FindRelationships(FromPos, ColourIndeces)
                Set oToRelations = DiagramRef.Relationships.FindRelationships(ToPos, ColourIndeces)
                
                lIndex = 1
                For Each oRelation In oFromRelations
                    If oRelation.FromPos Is FromPos And oRelation.ToPos Is ToPos Then
                        oFromRelations.Remove lIndex
                    End If
                    lIndex = lIndex + 1
                Next
                
                lIndex = 1
                For Each oRelation In oToRelations
                    If oRelation.FromPos Is FromPos And oRelation.ToPos Is ToPos Then
                        oToRelations.Remove lIndex
                    End If
                    lIndex = lIndex + 1
                Next
                
                If oFromRelations.Count = 1 And oToRelations.Count = 1 Then
                    Set oFromRelation = oFromRelations(1)
                    Set oToRelation = oToRelations(1)
                    
                    Dim oIntersection As Vector
                    Dim oLine1 As New LineVector
                    Dim oLine2 As New LineVector
                    Dim oInsCircle As Variant
                    
                    If oFromRelation.ToPos.Reference = FromPos.Reference Then
                        oLine1.SetLineVector oFromRelation.FromPos.Pos, oFromRelation.ToPos.Pos
                    Else
                        oLine1.SetLineVector oFromRelation.ToPos.Pos, oFromRelation.FromPos.Pos
                    End If
                    
                    If oToRelation.ToPos.Reference = ToPos.Reference Then
                        oLine2.SetLineVector oToRelation.FromPos.Pos, oToRelation.ToPos.Pos
                    Else
                        oLine2.SetLineVector oToRelation.ToPos.Pos, oToRelation.FromPos.Pos
                    End If
                    
                    oInsCircle = oLine1.InscribedCircle(oLine2)
    
                    If Not IsEmpty(oInsCircle) Then
                        DiagramRef.Graphics.PartialCircle oInsCircle(0).Scalar(DiagramRef.Zoom).Add(DiagramRef.TopLeft), oInsCircle(1) * DiagramRef.Zoom, DiagramRef.Colours(ColourIndeces(0)), oInsCircle(2).Scalar(DiagramRef.Zoom).Add(DiagramRef.TopLeft), oInsCircle(3).Scalar(DiagramRef.Zoom).Add(DiagramRef.TopLeft), False
                        
                        If UBound(ColourIndeces) = 0 Then
                            'PaperRef.StripedLine oInsCircle(4).Scalar(DiagramRef.Zoom).Add(DiagramRef.TopLeft), oInsCircle(5).Scalar(DiagramRef.Zoom).Add(DiagramRef.TopLeft), Array(lForeColour)
                            DiagramRef.Graphics.NormalLine oInsCircle(4).Scalar(DiagramRef.Zoom).Add(DiagramRef.TopLeft), oInsCircle(5).Scalar(DiagramRef.Zoom).Add(DiagramRef.TopLeft), lForeColour
                        Else
                            'PaperRef.StripedLine oInsCircle(4).Scalar(DiagramRef.Zoom).Add(DiagramRef.TopLeft), oInsCircle(5).Scalar(DiagramRef.Zoom).Add(DiagramRef.TopLeft), Colours
                            DiagramRef.Graphics.NormalLine oInsCircle(4).Scalar(DiagramRef.Zoom).Add(DiagramRef.TopLeft), oInsCircle(5).Scalar(DiagramRef.Zoom).Add(DiagramRef.TopLeft), DiagramRef.Colours(ColourIndeces(0))
                        End If
                    Else
                        If UBound(ColourIndeces) = 0 Then
                            DiagramRef.Graphics.StripedLine ostart, oEnd, Array(lForeColour)
                        Else
                            DiagramRef.Graphics.StripedLine ostart, oEnd, ColourIndeces
                        End If
                    End If
                    
                Else
                    If UBound(ColourIndeces) = 0 Then
                        DiagramRef.Graphics.StripedLine ostart, oEnd, Array(lForeColour)
                    Else
                        DiagramRef.Graphics.StripedLine ostart, oEnd, ColourIndeces
                    End If
                End If
            Case 2 ' link
                DiagramRef.Graphics.FilledCircle ostart, OuterCircleRadius * DiagramRef.Zoom * 1.2, InnerCircleRadius * DiagramRef.Zoom * 1#, vbBlack, DiagramRef.Colours(ColourIndeces(0))
                DiagramRef.Graphics.FilledCircle oEnd, OuterCircleRadius * DiagramRef.Zoom * 1.2, InnerCircleRadius * DiagramRef.Zoom * 1#, vbBlack, DiagramRef.Colours(ColourIndeces(0))
                DiagramRef.Graphics.FatLine ostart, oEnd, LineWidth * 1.2, vbBlack
                DiagramRef.Graphics.FatLine ostart, oEnd, LineWidth, DiagramRef.Colours(ColourIndeces(0))
        End Select
    Else
        DiagramRef.PaperRef.FillStyle = 1
        DiagramRef.PaperRef.DrawWidth = 1
        DiagramRef.PaperRef.Circle (FromPos.Pos.Scalar(DiagramRef.Zoom).Add(DiagramRef.TopLeft).X, FromPos.Pos.Scalar(DiagramRef.Zoom).Add(DiagramRef.TopLeft).Y), 30, vbRed
    End If
End Sub

Private Sub Class_Initialize()
    ColourIndeces = Array()
End Sub

