VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CirclePrimitive"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mlValue As Long
Private moPosition As Vector
Private moCentre As Vector
Private mlSize As Long

Private mlPreviousSign As Long

Public Sub Initialise(oCentre As Vector, lSize As Long)
    Set moCentre = oCentre
    mlSize = lSize
End Sub

Public Sub SetStart(oPoint As Vector)
    Set moPosition = oPoint.Subs(moCentre)
    mlValue = moPosition.Dot(moPosition) - mlSize * mlSize
    mlPreviousSign = Sgn(mlValue)
End Sub

Public Function MoveRight() As Long
    Dim lSign As Long
    
    mlValue = mlValue + 2 * moPosition.X + 1
    moPosition.X = moPosition.X + 1
    lSign = Sgn(mlValue)
    
    If lSign <> mlPreviousSign Then
        MoveRight = 1
    End If
    mlPreviousSign = lSign
End Function

Public Function MoveLeft() As Long
    Dim lSign As Long
    
    mlValue = mlValue - 2 * moPosition.X + 1
    moPosition.X = moPosition.X - 1
    lSign = Sgn(mlValue)
    
    If lSign <> mlPreviousSign Then
        MoveLeft = 1
    End If
    mlPreviousSign = lSign
End Function

Public Function MoveDown() As Long
    Dim lSign As Long
    
    mlValue = mlValue + 2 * moPosition.Y + 1
    moPosition.Y = moPosition.Y + 1
    lSign = Sgn(mlValue)
    
    If lSign <> mlPreviousSign Then
        MoveDown = 1
    End If
    mlPreviousSign = lSign
End Function

