VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RelationshipList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public List As New Collection

Public Function FindRelationship(oFirst As Position, oSecond As Position) As Relationship
    Dim oRelationship As Relationship
    
    For Each oRelationship In List
        If oRelationship.FromPos Is oFirst Then
            If oRelationship.ToPos Is oSecond Then
                Set FindRelationship = oRelationship
                Exit Function
            End If
        End If
        If oRelationship.FromPos Is oSecond Then
            If oRelationship.ToPos Is oFirst Then
                Set FindRelationship = oRelationship
                Exit Function
            End If
        End If
    Next
End Function

Public Function FindRelationships(oPosition As Position, Optional vColours As Variant) As Collection
    Dim oRelationship As Relationship
    Dim vColour As Variant
    
    Set FindRelationships = New Collection
    For Each oRelationship In List
        If oRelationship.FromPos Is oPosition Or oRelationship.ToPos Is oPosition Then
            If IsMissing(vColours) Then
                FindRelationships.Add oRelationship
            Else
                For Each vColour In vColours
                    If InArray(oRelationship.ColourIndeces, CLng(vColour)) Then
                        FindRelationships.Add oRelationship
                        Exit For
                    End If
                Next
            End If
        End If
    Next
End Function

Public Function RemoveRelationship(oRelationship As Relationship)
    Dim oRelation As Relationship
    Dim iIndex As Long
    
    For iIndex = 1 To List.Count
        If List(iIndex) Is oRelationship Then
            List.Remove iIndex
            Exit Function
        End If
    Next
End Function

Public Function RemoveRelationshipWithReference(oPosition As Position)
    Dim oRelation As Relationship
    Dim bFinished As Boolean
    Dim iIndex As Long
    
    Do
        For iIndex = 1 To List.Count
            If List(iIndex).FromPos Is oPosition Or List(iIndex).ToPos Is oPosition Then
                List.Remove iIndex
                Exit For
            End If
        Next
        If iIndex = List.Count + 1 Then
            bFinished = True
        End If
    Loop While Not bFinished
End Function

Public Sub RemoveDuplicates()
    Dim oRelation As Relationship
    Dim iIndex As Long
    
    iIndex = 1
    While iIndex <= List.Count
        If List(iIndex).FromPos Is List(iIndex).ToPos Then
            List.Remove iIndex
        End If
        iIndex = iIndex + 1
    Wend
End Sub

Public Sub ClearAll()
    Dim oRelationship As Relationship
    
    For Each oRelationship In List
        oRelationship.ClearRelationship
    Next
End Sub

Public Sub RenderAll()
    Dim oRelationship As Relationship
    
    For Each oRelationship In List
        oRelationship.RenderRelationship
    Next
End Sub


Public Sub SendToFront(ByVal oPosition As Position)
    Dim lIndex As Long
    Dim oRelationship As Relationship
    
    For lIndex = 1 To List.Count
        Set oRelationship = List(lIndex)
        If oRelationship.FromPos.Reference = oPosition.Reference Or oRelationship.ToPos.Reference = oPosition.Reference Then
            List.Remove (lIndex)
            List.Add oRelationship
        End If
    Next
End Sub

Public Sub SendToBack(ByVal oPosition As Position)
    Dim lIndex As Long
    Dim oRelationship As Relationship
    
    For lIndex = 1 To List.Count
        Set oRelationship = List(lIndex)
        If oRelationship.FromPos.Reference = oPosition.Reference Or oRelationship.ToPos.Reference = oPosition.Reference Then
            List.Remove (lIndex)
            List.Add oRelationship, , 1
        End If
    Next
End Sub
