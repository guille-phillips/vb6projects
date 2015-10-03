VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsNode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private msIdentifier As String
Private moLinks() As clsNode
Private moReference As clsNode

Private bLinksNotEmpty As Boolean

Private bSearched As Boolean

Private goNetLex As ISaffronObject

Private mbRemove As Boolean

' Straight through properties
Public Property Get Identifier() As String
    Identifier = msIdentifier
End Property

Public Property Let Identifier(ByVal sIdentifier As String)
    msIdentifier = sIdentifier
End Property

Public Property Get Remove() As Boolean
    Remove = mbRemove
End Property

Public Property Let Remove(ByVal bRemove As Boolean)
    mbRemove = bRemove
End Property

Public Property Get Reference() As clsNode
    Set Reference = mbReference
End Property

Public Property Set Reference(ByVal oReference As clsNode)
    Set moReference = oReference
End Property

' Node management
Public Function AddNode(oNode As clsNode)
    If bLinksNotEmpty Then
        ReDim Preserve moLinks(UBound(moLinks) + 1)
    Else
        ReDim moLinks(0)
        bLinksNotEmpty = True
    End If
    Set moLinks(UBound(moLinks)) = oNode
End Function

Public Function AddReference(oNode As clsNode)
    If bReferencesNotEmpty Then
        ReDim Preserve moReferences(UBound(moReferences) + 1)
    Else
        ReDim moReferences(0)
        bReferencesNotEmpty = True
    End If
    Set moReferences(UBound(moReferences)) = oNode
End Function

Public Function RemoveNode(oNode As clsNode)
    Dim lPosition As Long
    Dim lPosition2 As Long
    
    If bLinksNotEmpty Then
        For lPosition = 0 To UBound(moLinks)
            If moLinks(lPosition) Is oNode Then
                For lPosition2 = lPosition To UBound(moLinks) - 1
                    Set moLinks(lPosition2) = moLinks(lPosition2 + 1)
                Next
                ReDim Preserve moLinks(UBound(moLinks) - 1)
            End If
        Next
    End If
End Function

' Node creation
Public Function Create(ByVal sIdentifier) As clsNode
    Set Create = New clsNode
    Create.Identifier = sIdentifier
End Function

Public Function CreateChild(ByVal sIdentifier) As clsNode
    Set CreateChild = New clsNode
    CreateChild.Identifier = sIdentifier
    AddNode CreateChild
End Function


Public Property Get Node(ByVal sIdentifier As String) As clsNode
    Dim lPosition As Long
    
    If bLinksNotEmpty Then
        For lPosition = 0 To UBound(moLinks)
            If moLinks(lPosition).Identifier = sIdentifier Then
                Set Node = moLinks(lPosition)
                Exit Property
            End If
        Next
    End If
End Property


Public Property Get NodeByPath(ByVal sPath As String) As clsNode
    Dim vPath As Variant
    Dim oNode As clsNode
    Dim lIndex As Long
    Dim oNextNode As clsNode
    
    Set oNode = Me
    vPath = Split(sPath, "|")
    For lIndex = 0 To UBound(vPath)
        Set oNode = oNode.Node(vPath(lIndex))
        If oNode Is Nothing Then
            Exit For
        End If
    Next

    Set NodeByPath = oNode
End Property


' File management
Public Sub LoadFile()
    Dim sGameNet As String
    Dim oGameTree As New SaffronTree
    
    LoadNetCore
    
    SaffronStream.Text = OpenTextFile("game_net.txt")
    
    If Not goNetLex.Parse(oGameTree) Then
        MsgBox "Invalid Game Net"
        End
    End If
    
    CreateNetList oGameTree, Me
    
    Debug.Print Me.NodeByPath("player|torch|z").Identifier
End Sub

Private Sub LoadNetCore()
    Dim sDef As String
    sDef = OpenTextFile("net.txt")
    If Not CreateRules(sDef) Then
        MsgBox "Bad Def!"
        End
    End If
    Set goNetLex = SaffronObject.Rules("netlist")
    
'    Dim oT As New SaffronTree
'    SaffronStream.Text = "a,c{"
'    If goNetLex.Parse(oT) Then
'        Stop
'    End If
    
End Sub

Private Sub CreateNetList(oTree As SaffronTree, oParent As clsNode, Optional ByVal bOverride As Boolean)
    Dim sNodeName As String
    Dim oSubTree As SaffronTree
    Dim oNewNode As clsNode
    Dim oReference As clsNode
    Dim oSearch As clsNode
    Dim bReference As Boolean
    Dim bRemove As Boolean

    For Each oSubTree In oTree.SubTree
        sNodeName = oSubTree(1).Text
        
        ClearSearch
        If Not bOverride Then
            Set oReference = goNet.Search(sNodeName)
        End If
        If oReference Is Nothing Then
            Set oNewNode = oParent.CreateChild(sNodeName)
            If oSubTree(2).Index = 1 Then
                If oSubTree(2)(1)(1).SubTree.Count > 0 Then
                    CreateNetList oSubTree(2)(1)(1), oNewNode
                Else
                    Stop
                End If
            End If
        Else
            If oSubTree(2).Index = 1 Then
                Set oNewNode = oParent.CreateChild(sNodeName)
                Set oNewNode.Reference = oReference
                If oSubTree(2).Index = 1 Then
                    CreateNetList oSubTree(2)(1)(1), oNewNode, True
                End If
            Else
                oParent.AddNode oReference
            End If
        End If
    Next
End Sub

Public Sub SaveFile()

End Sub


' Node searching
Public Sub ClearSearch()
    Dim lPosition As Long
    
    bSearched = False
    If bLinksNotEmpty Then
        For lPosition = 0 To UBound(moLinks)
            moLinks(lPosition).ClearSearch
        Next
    End If
End Sub


Public Property Get Search(ByVal sIdentifier As String) As clsNode
    Dim oNode As clsNode
    Dim lPosition As Long
    
    If Not bSearched Then
        bSearched = True
        If msIdentifier = sIdentifier Then
            Set Search = Me
            Exit Property
        End If
        
        If bLinksNotEmpty Then
            For lPosition = 0 To UBound(moLinks)
                Set Search = moLinks(lPosition).Search(sIdentifier)
                If Not Search Is Nothing Then
                    Exit Property
                End If
            Next
        End If
    End If

End Property


' Network description
Public Property Get Descriptor(Optional ByVal bBrackets As Boolean = True) As String
    Dim vNode As Variant
    Dim sDescriptorList As String
    
    Descriptor = msIdentifier
    If bLinksNotEmpty Then
        If bBrackets Then
            Descriptor = Descriptor & "{"
        End If
        For Each vNode In moLinks
            sDescriptorList = sDescriptorList & "|" & vNode.Descriptor
        Next
        Descriptor = Descriptor & Mid$(sDescriptorList, 2)
        If bBrackets Then
            Descriptor = Descriptor & "}"
        End If
    End If
End Property

Public Property Get DescriptorList() As String
    Dim vNode As Variant
    Dim sDescriptorList As String

    For Each vNode In moLinks
        DescriptorList = DescriptorList & "|" & vNode.Descriptor(False)
    Next
    DescriptorList = Mid$(DescriptorList, 2)
End Property
