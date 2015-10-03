VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LineVector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public StartPos As Vector
Public EndPos As Vector

Public Function Intersection(oLineVector As LineVector) As Vector
    Dim oDiff As Vector
    Dim dBSize As Double
    Dim oA As Vector
    Dim oB1 As Vector
    Dim oC1 As Vector
    
    Set oA = oLineVector.StartPos.Subs(StartPos)
    Set oB1 = Direction
    Set oC1 = oLineVector.Direction
    
    dBSize = oA.Cross(oC1) * oB1.Size / oB1.Cross(oC1)
    Set Intersection = StartPos.Add(oB1.Normal.Scalar(dBSize))
End Function

Public Sub SetLineVector(oStartPos As Vector, oEndPos As Vector)
    Set StartPos = oStartPos
    Set EndPos = oEndPos
End Sub

Public Function Direction() As Vector
    Set Direction = EndPos.Subs(StartPos)
End Function

Public Function InscribedCircle(oLineVector As LineVector) As Variant
    Dim oDiff As Vector
    Dim dBSize As Double
    Dim dCSize As Double
    Dim oA As Vector
    Dim oB1 As Vector
    Dim oC1 As Vector
    Dim dBaseB As Double
    Dim dBaseC As Double
    Dim dHalfSine As Double
    Dim dRadius As Double
    Dim dTan As Double
    Dim dHalfTan As Double
    Dim oCentre As Vector
    Dim oCircleEnd As Vector
    Dim oParity As Vector
    Dim bSwap As Boolean
    
    On Error Resume Next
    
    Set oA = oLineVector.StartPos.Subs(StartPos)
    Set oB1 = Direction
    Set oC1 = oLineVector.Direction
    
    Set oParity = oLineVector.EndPos.Subs(StartPos)
    bSwap = oParity.Cross(oB1) < 0
    
    If oB1.Cross(oC1) = 0 Then ' parallel
        InscribedCircle = Empty
        Exit Function
    End If
    
    dBSize = oA.Cross(oC1) * oB1.Size / oB1.Cross(oC1)
    dCSize = -oA.Cross(oB1) * oC1.Size / oC1.Cross(oB1)
    
    dBaseB = dBSize - oB1.Size
    dBaseC = dCSize - oC1.Size
    
    If dBaseB < dBaseC Then
        If oB1.Dot(oC1) <> 0 Then
            dTan = -oB1.Cross(oC1) / oB1.Dot(oC1)
            dHalfTan = (Sqr(1 + dTan * dTan) - 1) / dTan
            dRadius = dBaseB * dHalfTan
        Else
            If bSwap Then
                dRadius = dBaseB
            Else
                dRadius = -dBaseB
            End If
        End If
        
        Set oCentre = EndPos.Add(oB1.Perpendicular.Normal.Scalar(dRadius))
        Set oCircleEnd = oCentre.Add(oC1.Perpendicular.Normal.Scalar(dRadius))
        If Not bSwap Then
            InscribedCircle = Array(oCentre, Abs(dRadius), EndPos, oCircleEnd, oCircleEnd, oLineVector.EndPos)
        Else
            InscribedCircle = Array(oCentre, Abs(dRadius), oCircleEnd, EndPos, oCircleEnd, oLineVector.EndPos)
        End If
    Else
        If oC1.Dot(oB1) <> 0 Then
            dTan = -oC1.Cross(oB1) / oC1.Dot(oB1)
            dHalfTan = (Sqr(1 + dTan * dTan) - 1) / dTan
            dRadius = dBaseC * dHalfTan
        Else
            If bSwap Then
                dRadius = -dBaseC
            Else
                dRadius = dBaseC
            End If
        End If
        Set oCentre = oLineVector.EndPos.Add(oC1.Perpendicular.Normal.Scalar(dRadius))
        Set oCircleEnd = oCentre.Add(oB1.Perpendicular.Normal.Scalar(dRadius))
        
        If Not bSwap Then
            InscribedCircle = Array(oCentre, Abs(dRadius), oCircleEnd, oLineVector.EndPos, oCircleEnd, EndPos)
        Else
            InscribedCircle = Array(oCentre, Abs(dRadius), oLineVector.EndPos, oCircleEnd, oCircleEnd, EndPos)
        End If
    End If
End Function
