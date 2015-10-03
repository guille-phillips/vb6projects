VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FileIO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public PositionsRef As PositionList
Public RelationsRef As RelationshipList

Public FileStore As String
Public DiagramRef As Diagram

Public Sub ReadFile()
    Dim oPosition As Position
    Dim oRelationship As Relationship
    Dim oFile As TextStream
    Dim sLine As String
    Dim oFSO As New FileSystemObject
    Dim oTree As ParseTree
    Dim vColours As Variant
    Dim vColourIndeces As Variant
    Dim lColourIndex As Long
    Dim oColourItem As ParseTree
    
    If oFSO.FileExists(FileStore) Then
        Set oFile = oFSO.OpenTextFile(FileStore, ForReading)
        While Not oFile.AtEndOfStream
            sLine = oFile.ReadLine
            
            Set oTree = New ParseTree
            Stream.Text = sLine
            If oParseLine.Parse(oTree) Then
                Select Case oTree.Index
                    Case 1 ' Position
                        Set oPosition = New Position
                        With oPosition
                            .Reference = oTree(1)(1).Text
                            .Name = oTree(1)(2).Text
                            .Pos.X = oTree(1)(3).Text
                            .Pos.Y = oTree(1)(4).Text
                            .ColourIndex = oTree(1)(5).Text
                            .Orientation = oTree(1)(6).Text
                            Set .DiagramRef = DiagramRef
                        End With
                        DiagramRef.Positions.List.Add oPosition
                    Case 2 ' Relation
                        Set oRelationship = New Relationship
                        With oRelationship
                            Set .FromPos = DiagramRef.Positions.FindPositionFromReference(oTree(1)(1).Text)
                            Set .ToPos = DiagramRef.Positions.FindPositionFromReference(oTree(1)(2).Text)
                            If .FromPos Is Nothing Then
                                Stop
                            End If
                            If .ToPos Is Nothing Then
                                Stop
                            End If
                            .Style = Val(oTree(1)(3).Text)
                            vColours = .ColourIndeces
                            vColourIndeces = .ColourIndeces
                            For lColourIndex = 1 To oTree(1)(4).Index
                                AddToArray vColourIndeces, oTree(1)(4)(lColourIndex).Text
                            Next
                            .ColourIndeces = vColours
                            .ColourIndeces = vColourIndeces
                            Set .DiagramRef = DiagramRef
                        End With
                        DiagramRef.Relationships.List.Add oRelationship
                    Case 3 ' Offset
                        DiagramRef.TopLeft.SetVector oTree(1)(1).Text, oTree(1)(2).Text
                    Case 4 ' Zoom
                        DiagramRef.Zoom = oTree(1).Text
                    Case 5 ' Colours
                        lColourIndex = 0
                        For Each oColourItem In oTree(1)(1).SubTree
                            vColours = DiagramRef.Colours
                            vColours(lColourIndex) = CLng(oColourItem.Text)
                            DiagramRef.Colours = vColours
                            lColourIndex = lColourIndex + 1
                        Next
                End Select
            End If
            
        Wend
    End If
End Sub

Public Sub WriteFile()
    Dim oPosition As Position
    Dim oRelationship As Relationship
    Dim oFSO As New FileSystemObject
    Dim oFile As TextStream
    
    Dim vColour As Variant
    
    Set oFile = oFSO.OpenTextFile(FileStore, ForWriting, True)
    oFile.WriteLine "C:" & Join(DiagramRef.Colours, "|")
    oFile.WriteLine "O:" & DiagramRef.TopLeft.X & "|" & DiagramRef.TopLeft.Y
    oFile.WriteLine "Z:" & DiagramRef.Zoom
    
    For Each oPosition In PositionsRef.List
        With oPosition
            oFile.WriteLine "P:" & .Reference & "|" & .Name & "|" & .Pos.X & "|" & .Pos.Y & "|" & .ColourIndex & "|" & .Orientation
        End With
    Next
    
    For Each oRelationship In RelationsRef.List
        With oRelationship
            oFile.WriteLine "R:" & .FromPos.Reference & "|" & .ToPos.Reference & "|" & .Style & "|" & Join(.ColourIndeces, "|")
        End With
    Next
    oFile.Close
End Sub


