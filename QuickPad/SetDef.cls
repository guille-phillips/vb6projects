VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SetDef"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Invisible As Boolean
Public Reversed As Boolean
Public Elements As New Collection

Public Sub CreateSet(oTree As SaffronTree)
    Dim oSubTree As SaffronTree
    Dim oElement As Element
    
    If oTree(1).Index = 1 Then
        Select Case oTree(1)(1).Index
            Case 1, 2
                Invisible = True
                Reversed = True
            Case 3
                Invisible = True
            Case 4
                Reversed = True
        End Select
    End If

    For Each oSubTree In oTree(2).SubTree
        Set oElement = New Element
        oElement.CreateElement oSubTree
        Elements.Add oElement
    Next
End Sub

Public Property Get Text(ByVal lIndex As Long) As String
    Dim lPosition As Long
    
    If Not Invisible Then
        If Reversed Then
            lIndex = SetSize - lIndex - 1
        End If
        lPosition = 1
        While lIndex > (Elements(lPosition).SetSize - 1)
            lIndex = lIndex - Elements(lPosition).SetSize
            lPosition = lPosition + 1
            If lPosition > Elements.Count Then
                Exit Property
            End If
        Wend
        Text = Elements(lPosition).Text(lIndex)
    End If
End Property

Public Function SetSize() As Long
    Dim oElement As Element
    
    For Each oElement In Elements
        SetSize = SetSize + oElement.SetSize
    Next
End Function

Public Property Get AllText() As String
    Dim lIndex As Long
    
    If Not Invisible Then
        If Not Reversed Then
            For lIndex = 0 To SetSize - 1
                AllText = AllText & Text(lIndex)
            Next
        Else
            For lIndex = SetSize - 1 To 0 Step -1
                AllText = AllText & Text(lIndex)
            Next
        End If
    End If
End Property
