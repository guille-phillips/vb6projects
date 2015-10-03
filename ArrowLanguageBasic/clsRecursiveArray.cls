VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRecursiveArray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Count As Long
Private moArray() As Variant

Public Function SetElement(vValue As Variant, ParamArray vIndecesParams() As Variant)
    Dim lIndex As Long
    Dim lIndecesLast As Long
    Dim vIndeces As Variant
    
    If IsArray(vIndecesParams(0)) Then
        vIndeces = vIndecesParams(0)
    Else
        vIndeces = Array()
        ReDim vIndeces(UBound(vIndecesParams))
        For lIndex = 0 To UBound(vIndecesParams)
            vIndeces(lIndex) = vIndecesParams(lIndex)
        Next
    End If
    
    lIndecesLast = UBound(vIndeces)
    
    lIndex = vIndeces(0)
    
    If lIndex >= Count Then
        Count = vIndeces(0) + 1
        ReDim Preserve moArray(Count - 1)
        If lIndecesLast = 0 Then
            moArray(lIndex) = vValue
        Else
            If IsEmpty(moArray(lIndex)) Then
                Set moArray(lIndex) = New clsRecursiveArray
                ShiftArray vIndeces
                moArray(lIndex).SetElement vValue, vIndeces
            ElseIf TypeOf moArray(lIndex) Is clsRecursiveArray Then
                Stop
            End If
        End If
    Else
        If lIndecesLast = 0 Then
            If TypeOf moArray(lIndex) Is clsRecursiveArray Then
                Err.Raise -2, "RecursiveArray", "Element already set"
            Else
                moArray(lIndex) = vValue
            End If
        Else
            If IsEmpty(moArray(lIndex)) Then
                Set moArray(lIndex) = New clsRecursiveArray
                ShiftArray vIndeces
                moArray(lIndex).SetElement vValue, vIndeces
            ElseIf TypeOf moArray(lIndex) Is clsRecursiveArray Then
                ShiftArray vIndeces
                moArray(lIndex).SetElement vValue, vIndeces
            Else
                Err.Raise -1, "Recursive Array", "Element already set"
            End If
        End If
    End If
End Function

Private Sub ShiftArray(vArray As Variant)
    Dim lIndex As Long
    
    For lIndex = 0 To UBound(vArray) - 1
        vArray(lIndex) = vArray(lIndex + 1)
    Next
    ReDim Preserve vArray(UBound(vArray) - 1)
End Sub

Public Function GetElement(ParamArray vIndecesParams() As Variant) As Variant
    Dim lIndex As Long
    Dim lIndecesLast As Long
    Dim vIndeces As Variant
    
    If IsArray(vIndecesParams(0)) Then
        vIndeces = vIndecesParams(0)
    Else
        vIndeces = Array()
        ReDim vIndeces(UBound(vIndecesParams))
        For lIndex = 0 To UBound(vIndecesParams)
            vIndeces(lIndex) = vIndecesParams(lIndex)
        Next
    End If
    
    lIndecesLast = UBound(vIndeces)
    
    lIndex = vIndeces(0)
    
    If lIndecesLast = 0 Then
        If TypeOf moArray(lIndex) Is Object  Then
            Set GetElement = moArray(lIndex)
        Else
            GetElement = moArray(lIndex)
        End If
    Else
        ShiftArray vIndeces
        If lIndex >= Count Then
            Exit Function
        Else
            If IsEmpty(moArray(lIndex)) Then
                Exit Function
            Else
                GetElement = moArray(lIndex).GetElement(vIndeces)
            End If
        End If
    End If
End Function


