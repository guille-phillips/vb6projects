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

Private mvValue As Variant
Private moParent As clsNode
Private mlLogicalKey As Long
Private mlPhysicalKey As Long

Private moItems() As clsNode
Private mlPhysicalItemCount As Long ' Last item in list
Private mlLogicalItemCount As Long ' Key of new item

Private moKeys As New clsKey
Private msTextKey As String

Public Property Get Keys() As clsKey
    Set Keys = moKeys
End Property

Public Property Set Keys(oKeys As clsKey)
    Set moKeys = oKeys
End Property

Public Property Get Value() As Variant
    If VarType(mvValue) = vbObject Then
        Set Value = mvValue
    Else
        Value = mvValue
    End If
End Property

Public Property Set Value(ByVal vValue As Variant)
    Set mvValue = vValue
End Property

Public Property Let Value(ByVal vValue As Variant)
    mvValue = vValue
End Property

Public Property Get Parent() As clsNode
    Set Parent = moParent
End Property

Public Property Set Parent(ByVal oParent As clsNode)
    Set moParent = oParent
End Property

Public Property Let LogicalKey(ByVal lLogicalKey As Long)
    mlLogicalKey = lLogicalKey
End Property

Public Property Get LogicalKey() As Long
    LogicalKey = mlLogicalKey
End Property

Public Property Let PhysicalKey(ByVal lPhysicalKey As Long)
    mlPhysicalKey = lPhysicalKey
End Property

Public Property Get PhysicalKey() As Long
    PhysicalKey = mlPhysicalKey
End Property

Public Property Get TextKey() As String
    TextKey = msTextKey
End Property

Public Property Let TextKey(ByVal sTextKey As String)
    Dim oKey As clsKey
    
    If sTextKey = msTextKey Then
        Exit Property
    End If
    
    If Not moKeys.Word(sTextKey) Is Nothing Then
        Err.Raise -1, "clsKey", "Key '" & sTextKey & "' already assigned."
        Exit Property
    End If
    
    msTextKey = sTextKey
    Set oKey = moKeys.AddWord(sTextKey)
    Set oKey.Value = Me
End Property

Public Sub Move(lPhysicalPositionFrom As Long, lPhysicalPositionTo As Long)
    Dim oMovee As clsNode
    Dim lDirection As Long
    Dim lStart As Long
    Dim lEnd As Long
    Dim lIndex As Long
    
    If lPhysicalPositionFrom = lPhysicalPositionTo Then
        Exit Sub
    ElseIf lPhysicalPositionFrom > lPhysicalPositionTo Then
        lDirection = -1
    Else
        lDirection = 1
    End If
    
    Set oMovee = moItems(lPhysicalPositionFrom)
    For lIndex = lPhysicalPositionFrom To lPhysicalPositionTo - lDirection Step lDirection
        Set moItems(lIndex) = moItems(lIndex + lDirection)
        moItems(lIndex).PhysicalKey = lIndex
    Next
    Set moItems(lPhysicalPositionTo) = oMovee
    moItems(lPhysicalPositionTo).PhysicalKey = lPhysicalPositionTo
End Sub

Public Function AddNew(Optional ByVal lPhysicalKey As Long = -1, Optional ByVal sTextKey As String, Optional ByVal lLogicalKey As Long = -1) As clsNode
    Set AddNew = New clsNode
    
    If lLogicalKey = -1 Then
        AddNew.LogicalKey = mlLogicalItemCount
    Else
        AddNew.LogicalKey = lLogicalKey
        If lLogicalKey > mlLogicalItemCount Then
            mlLogicalItemCount = lLogicalKey + 1
        End If
    End If
    
    AddNew.PhysicalKey = mlPhysicalItemCount
    Set AddNew.Parent = Me
    Set AddNew.Keys = moKeys
    AddNew.TextKey = sTextKey
    ReDim Preserve moItems(mlPhysicalItemCount)
    Set moItems(mlPhysicalItemCount) = AddNew
    mlPhysicalItemCount = mlPhysicalItemCount + 1
    
    If lLogicalKey = -1 Then
        mlLogicalItemCount = mlLogicalItemCount + 1
    End If
    
    If lPhysicalKey > -1 Then
        Move mlPhysicalItemCount - 1, lPhysicalKey
    End If
End Function

Public Sub RemoveTextKey(ByVal sKey As String)
    Dim lIndex As Long
    Dim oKey As clsKey
    
    Set oKey = Keys.Word(sKey)
    If Not oKey Is Nothing Then
        If TypeOf oKey.Value Is clsNode Then
            If oKey.Value.Parent Is Me Then
                RemovePhysical oKey.Value.PhysicalKey
            End If
        End If
    End If
End Sub

Public Sub RemoveLogical(ByVal lLogicalKey As Long)
    Dim lIndex As Long
    
    For lIndex = 0 To mlPhysicalItemCount - 1
        If moItems(lIndex).LogicalKey = lLogicalKey Then
            RemovePhysical lIndex
            Exit For
        End If
    Next
End Sub

Public Sub RemovePhysical(ByVal lPhysicalKey As Long)
    Dim lIndex As Long
    
    If moItems(lPhysicalKey).TextKey <> "" Then
        moKeys.RemoveWord moItems(lPhysicalKey).TextKey
    End If
    Move lPhysicalKey, mlPhysicalItemCount - 1
    mlPhysicalItemCount = mlPhysicalItemCount - 1
    If mlPhysicalItemCount > 0 Then
        ReDim Preserve moItems(mlPhysicalItemCount - 1)
    Else
        Erase moItems
    End If
End Sub

Public Sub RemoveAll()
    mlPhysicalItemCount = 0
    mlLogicalItemCount = 0
    mlPhysicalKey = 0
    mlLogicalKey = 0
    Erase moItems
End Sub

Public Function FindLogicalItem(ByVal vValue As Variant, Optional ByVal lPhysicalKey = -1) As clsNode
    Dim lIndex As Long
    
    Select Case VarType(vValue)
        Case vbObject
            If vValue Is Nothing Then
                Exit Function
            End If
            For lIndex = lPhysicalKey + 1 To mlPhysicalItemCount - 1
                If VarType(moItems(lIndex).Value) = vbObject Then
                    If moItems(lIndex).Value Is vValue Then
                        Set FindLogicalItem = moItems(lIndex)
                        Exit For
                    End If
                End If
            Next
        Case Else
            For lIndex = lPhysicalKey + 1 To mlPhysicalItemCount - 1
                If VarType(moItems(lIndex).Value) <> vbObject Then
                    If moItems(lIndex).Value = vValue Then
                        Set FindLogicalItem = moItems(lIndex)
                        Exit For
                    End If
                End If
            Next
    End Select
End Function

Public Function FindPhysicalItem(ByVal oObject As Object, Optional ByVal lPhysicalKey = -1) As clsNode
    Dim lIndex As Long
    
    For lIndex = lPhysicalKey + 1 To mlPhysicalItemCount - 1
        If moItems(lIndex) Is oObject Then
            Set FindPhysicalItem = moItems(lIndex)
            Exit For
        End If
    Next
End Function

Public Property Get ItemPhysical(ByVal lPhysicalKey) As clsNode
    Set ItemPhysical = moItems(lPhysicalKey)
End Property

Public Property Get ItemLogical(ByVal lLogicalKey As Long) As clsNode
    Dim lIndex As Long
    
    For lIndex = 0 To mlPhysicalItemCount - 1
        If moItems(lIndex).LogicalKey = lLogicalKey Then
            Set ItemLogical = moItems(lIndex)
            Exit For
        End If
    Next
End Property

Public Property Get Count() As Long
    Count = mlPhysicalItemCount
End Property

Public Function Create() As clsNode
    Set Create = New clsNode
End Function

