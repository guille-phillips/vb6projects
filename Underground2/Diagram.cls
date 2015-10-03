VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Diagram"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Relationships As New RelationshipList
Public Positions As New PositionList
Public FileIOs As New FileIO

Private moPaperRef As Form
Public Graphics As New Graphics

Public TopLeft As New Vector
Public Zoom As Double
Public Colours As Variant

Public Property Set PaperRef(oPaper As Form)
    Set moPaperRef = oPaper
    Set Graphics.DiagramRef = Me
End Property

Public Property Get PaperRef() As Form
    Set PaperRef = moPaperRef
End Property

Public Sub RenderAndSave()
    PaperRef.Cls
    PaperRef.BackColor = Colours(18)
    'PaperRef.Watermark
    Relationships.RenderAll
    Positions.RenderAll
    Graphics.Graticule
    FileIOs.WriteFile
End Sub

Public Sub Render()
    PaperRef.Cls
    Relationships.RenderAll
    Positions.RenderAll
    Graphics.Graticule
End Sub

Private Sub Class_Initialize()
    Set FileIOs.RelationsRef = Relationships
    Set FileIOs.PositionsRef = Positions
    Set FileIOs.DiagramRef = Me
End Sub

Public Sub Initialise(sFilePath As String)
    If sFilePath <> "" Then
        FileIOs.FileStore = sFilePath
    Else
        FileIOs.FileStore = App.Path & "\diagram.txt"
    End If
    FileIOs.ReadFile
End Sub

Public Function FindColourIndex(ByVal lColour As Long) As Long
    Dim lIndex As Long
    
    FindColourIndex = -1
    For lIndex = 0 To UBound(Colours)
        If Colours(lIndex) = lColour Then
            FindColourIndex = lIndex
            Exit Function
        End If
    Next
End Function

