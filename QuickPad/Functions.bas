Attribute VB_Name = "Functions"
Option Explicit

Public Sub test()
'    Dim vTable As Variant
'    Dim vLists As Variant
'
'    vTable = Array(Array("a", "1"), Array("a", "2"), Array("b", "2"), Array("b", "1"))
'    vLists = FactorColumns(vTable)
End Sub

Public Function TableMaxWidth(vTable As Variant) As Long
    Dim lRow As Long
    Dim lMaxWidth As Long
    Dim vRow As Variant
    
    For lRow = 0 To UBound(vTable)
        vRow = vTable(lRow)
        If UBound(vRow) > lMaxWidth Then
            lMaxWidth = UBound(vRow)
        End If
    Next
    TableMaxWidth = lMaxWidth
End Function

Public Function DeriveExpression(ByVal vTable) As String
    Dim vFactors As Variant
    Dim lIndex As Long
    Dim sSubFactor As String
    Dim lFactorIndex As Long
    Dim lSubFactorIndex As Long
    Dim vSubFactor As Variant
    Dim vMainFactor As Variant
    
    Dim sFactors As String
    
    vFactors = FactorTable(vTable)

    DeriveExpression = FactorsAsString(vFactors) & "^"
End Function

Public Function ArrayAsString(ByVal vFactors As Variant) As String
    Dim lIndex As Long
    
    For lIndex = 0 To UBound(vFactors)
        If VarType(vFactors(lIndex)) <> (vbArray Or vbVariant) Then
            ArrayAsString = ArrayAsString & "," & EscapeExpression(vFactors(lIndex))
        Else
            ArrayAsString = ArrayAsString & "," & ArrayAsString(vFactors(lIndex))
        End If
    Next
    ArrayAsString = "{" & Mid$(ArrayAsString, 2) & "}"
End Function

Public Function FactorsAsString(ByVal vFactors As Variant) As String
    Dim sFactors As String
    Dim lFactorIndex As Long
    
    If ArrayDepth(vFactors) = 0 Then
        FactorsAsString = ArrayAsString(vFactors)
        Exit Function
    End If
    For lFactorIndex = 0 To UBound(vFactors)
        sFactors = sFactors & "," & ArrayAsString(vFactors(lFactorIndex)(0)) & "%" & FactorsAsString(vFactors(lFactorIndex)(1))
    Next
    FactorsAsString = "{" & Mid$(sFactors, 2) & "}"
End Function

Private Function ArrayDepth(ByVal vArray As Variant) As Long
    Dim lIndex As Long
    Dim lDepth As Long
    
    For lIndex = 0 To UBound(vArray)
        If VarType(vArray(lIndex)) = (vbArray Or vbVariant) Then
            lDepth = ArrayDepth(vArray(lIndex)) + 1
            If lDepth > ArrayDepth Then
                ArrayDepth = lDepth
            End If
        End If
    Next
End Function

Private Function EscapeExpression(ByVal sExpression As String) As String
    sExpression = Replace$(sExpression, "|", "||")
    sExpression = Replace$(sExpression, ",", "|,")
    sExpression = Replace$(sExpression, "%", "|%")
    sExpression = Replace$(sExpression, "^", "|^")
    sExpression = Replace$(sExpression, "{", "|{")
    sExpression = Replace$(sExpression, "}", "|{")
    sExpression = Replace$(sExpression, "@", "|@")
    sExpression = Replace$(sExpression, "~", "|~")
    sExpression = Replace$(sExpression, "#", "|#")
    sExpression = Replace$(sExpression, ":", "|:")
    sExpression = Replace$(sExpression, "$", "|$")
    EscapeExpression = sExpression
End Function

Private Function EscapeExpressionArray(ByVal vArray As Variant) As Variant
    Dim lIndex As Long
    
    For lIndex = 0 To UBound(vArray)
        vArray(lIndex) = EscapeExpression(vArray(lIndex))
    Next
    EscapeExpressionArray = vArray
End Function

Public Function FactorTable(ByVal vTable As Variant) As Variant
    Dim lRows As Long
    Dim vResult As Variant
    Dim lIndex As Long
    Dim sColumn1 As String
    Dim lSearchIndex As Long
    Dim bFound As Boolean
    Dim lResultRows As Long
    Dim lSearchFound As Long
    Dim vMembers As Variant
    Dim lTotalMembers As Long
    Dim bEqual As Boolean
    Dim vMembers0 As Variant
    Dim lTotalMembers0 As Long
    Dim vEqualRows As Variant
    Dim lEqualityIndex As Long
    Dim lEqualityMatch As Long
    Dim vFactor As Variant
    Dim vSubFactors As Variant
    Dim vFactors As Variant
    Dim lTableWidth As Long
    
    lTableWidth = TableMaxWidth(vTable)
    
    If lTableWidth = 0 Then
        FactorTable = ArrayFromTableColumn(vTable, 0)
        Exit Function
    End If
    
    lRows = UBound(vTable)
    
    ' Group repeating members in first column
    vResult = Array()
    
    For lIndex = 0 To lRows
        sColumn1 = vTable(lIndex)(0)
        lResultRows = UBound(vResult)
        
        ' Search for exisiting group member
        bFound = False
        For lSearchIndex = 0 To lResultRows
            If vResult(lSearchIndex)(0) = sColumn1 Then
                bFound = True
                lSearchFound = lSearchIndex
                Exit For
            End If
        Next
        
        ' Create group member if not found
        If Not bFound Then
            lSearchFound = lResultRows + 1
            ReDim Preserve vResult(lSearchFound)
            vResult(lSearchFound) = Array(sColumn1, Array())
        End If
        
        vMembers = vResult(lSearchFound)(1)
        lTotalMembers = UBound(vMembers) + 1
        ReDim Preserve vMembers(lTotalMembers)
        vMembers(lTotalMembers) = ExtractElements(vTable(lIndex), 1, UBound(vTable(lIndex)))
        vResult(lSearchFound)(1) = vMembers
    Next
    
    ' Sort the groups
    For lIndex = 0 To UBound(vResult)
        vResult(lIndex)(1) = SortArray(vResult(lIndex)(1))
    Next
    
    ' Uniquely index the groups with equal rows
    vEqualRows = Array()
    ReDim vEqualRows(UBound(vResult))
    
    For lEqualityIndex = 0 To UBound(vResult)
        If IsEmpty(vEqualRows(lEqualityIndex)) Then
            vEqualRows(lEqualityIndex) = lEqualityIndex

            For lIndex = lEqualityIndex + 1 To UBound(vResult)
                bEqual = True
                If IsEmpty(vEqualRows(lIndex)) Then
                    If ArraysEqual(vResult(lEqualityIndex)(1), vResult(lIndex)(1)) Then
                        vEqualRows(lIndex) = lEqualityIndex
                    End If
                End If
            Next
        End If
    Next

    ' Factor out each unique group index
    vFactors = Array()
    For lEqualityIndex = 0 To UBound(vResult)
        If Not IsEmpty(vEqualRows(lEqualityIndex)) Then
            lEqualityMatch = vEqualRows(lEqualityIndex)
            vFactor = Array(vResult(lEqualityIndex)(0))
            vSubFactors = FactorTable(vResult(lEqualityIndex)(1))
            For lSearchIndex = lEqualityIndex + 1 To UBound(vResult)
                If vEqualRows(lSearchIndex) = lEqualityMatch Then ' unique indeces match
                    ArrayAppend vFactor, vResult(lSearchIndex)(0)
                    vEqualRows(lSearchIndex) = Empty
                End If
            Next
            ArrayAppend vFactors, Array(vFactor, vSubFactors)
        End If
    Next
    FactorTable = vFactors
End Function

Public Function SubTable(vTable As Variant, vEqualities As Variant, lEqualitySelect As Long) As Variant
    Dim vSubTable As Variant
    Dim lIndex As Long
    Dim lColumn As Long
    Dim vRow As Variant
    
    vSubTable = Array()
    
    For lIndex = 0 To UBound(vEqualities)
        If vEqualities(lIndex) = lEqualitySelect Then
            vRow = Array()
            For lColumn = 1 To UBound(vTable(lIndex))
                ArrayAppend vRow, vTable(lIndex)(lColumn)
            Next
            ArrayAppend vSubTable, vRow
        End If
    Next
    SubTable = vSubTable
End Function

Public Function ArrayAppend(vArray As Variant, vItem As Variant)
    Dim lUbound As Long
    lUbound = UBound(vArray) + 1
    ReDim Preserve vArray(lUbound)
    vArray(lUbound) = vItem
End Function

Private Function SortArray(ByVal vArray As Variant) As Variant
    Dim vTemp As Variant
    Dim bSorted As Boolean
    Dim lIndex As Long
    
    While Not bSorted
        bSorted = True
        For lIndex = 0 To UBound(vArray) - 1
            If ArrayGreater(vArray(lIndex), vArray(lIndex + 1)) Then
                vTemp = vArray(lIndex)
                vArray(lIndex) = vArray(lIndex + 1)
                vArray(lIndex + 1) = vTemp
                bSorted = False
            End If
        Next
    Wend
    
    SortArray = vArray
End Function

Public Function ArrayFromTableColumn(ByVal vTable As Variant, ByVal lColumn As Long) As Variant
    Dim lIndex As Long
    Dim vResult As Variant
    
    vResult = Array()
    For lIndex = 0 To UBound(vTable)
        ReDim Preserve vResult(lIndex)
        vResult(lIndex) = vTable(lIndex)(lColumn)
    Next
    ArrayFromTableColumn = vResult
End Function


Public Function ArrayIntoTableColumn(ByVal vTable As Variant, ByVal lColumn As Long, ByVal vColumn As Variant) As Variant
    Dim lIndex As Long
    
    If UBound(vTable) < UBound(vColumn) Then
        ReDim vTable(UBound(vColumn))
    End If
    vTable = PadTable(vTable)
    For lIndex = 0 To UBound(vColumn)
        vTable(lIndex)(lColumn) = vColumn(lIndex)
    Next
    ArrayIntoTableColumn = vTable
End Function

Public Function ArrayFromTableRow(ByVal vTable As Variant, ByVal lRow As Long) As Variant
    Dim lIndex As Long
    Dim vResult As Variant
    
    vResult = Array()
    For lIndex = 0 To UBound(vTable(lRow))
        ReDim Preserve vResult(lIndex)
        vResult(lIndex) = vTable(lRow)(lIndex)
    Next
    ArrayFromTableRow = vResult
End Function


Public Function ArrayIntoTableRow(ByVal vTable As Variant, ByVal lRow As Long, ByVal vRow As Variant) As Variant
    Dim lIndex As Long
    Dim vTableRow As Variant
    
    vTableRow = vTable(lRow)
    If UBound(vTableRow) < UBound(vRow) Then
        ReDim vTableRow(UBound(vRow))
    End If
    vTable = PadTable(vTable)
    For lIndex = 0 To UBound(vRow)
        vTableRow(lIndex) = vRow(lIndex)
    Next
    vTable(lRow) = vTableRow
    ArrayIntoTableRow = vTable
End Function

Public Function ConvertArrayToTable(ByVal vArray As Variant) As Variant
    Dim vTable As Variant
    Dim lIndex As Long
    
    vTable = Array()
    
    ReDim vTable(UBound(vArray))
    For lIndex = 0 To UBound(vArray)
        vTable(lIndex) = Array(vArray(lIndex))
    Next
    ConvertArrayToTable = vTable
End Function


Public Function PadTable(vTable As Variant) As Variant
    Dim lRow As Long
    Dim lMaxWidth As Long
    Dim vRow As Variant
    Dim vPaddedTable As Variant
    
    If UBound(vTable) = -1 Then
        PadTable = vTable
        Exit Function
    End If
    
    lMaxWidth = TableMaxWidth(vTable)
    
    vPaddedTable = Array()
    ReDim vPaddedTable(UBound(vTable))
    
    For lRow = 0 To UBound(vTable)
        vRow = vTable(lRow)
        ReDim Preserve vRow(lMaxWidth)
        vPaddedTable(lRow) = vRow
    Next
    
    PadTable = vPaddedTable
End Function

Private Function RowFromTable(ByVal vTable As Variant, ByVal lRow As Long) As Variant
    Dim lIndex As Long
    Dim vResult As Variant
    Dim vRow As Variant
    
    vResult = Array()
    vRow = vTable(lRow)
    For lIndex = 0 To UBound(vRow)
        ReDim Preserve vResult(lIndex)
        vResult(lIndex) = vTable(lRow)(lIndex)
    Next
    RowFromTable = vResult
End Function

Private Function ExtractElements(ByVal vArray As Variant, ByVal lStart As Long, ByVal lEnd As Long) As Variant
    Dim vResult As Variant
    Dim lIndex As Long
    
    vResult = Array()
    ReDim vResult(lEnd - lStart)
    For lIndex = lStart To lEnd
        vResult(lIndex - lStart) = vArray(lIndex)
    Next
    ExtractElements = vResult
End Function

Private Function ArraysEqual(ByVal vArray1 As Variant, ByVal vArray2 As Variant) As Boolean
    Dim lIndex As Long
    
    If UBound(vArray1) <> UBound(vArray2) Then
        Exit Function
    End If
    For lIndex = 0 To UBound(vArray1)
        If VarType(vArray1(lIndex)) <> (vbArray Or vbVariant) Then
            If VarType(vArray2(lIndex)) <> (vbArray Or vbVariant) Then
                If vArray1(lIndex) <> vArray2(lIndex) Then
                    Exit Function
                End If
            Else
                Exit Function
            End If
        Else
            If VarType(vArray2(lIndex)) <> (vbArray Or vbVariant) Then
                Exit Function
            Else
                ArraysEqual = ArraysEqual(vArray1(lIndex), vArray2(lIndex))
                If Not ArraysEqual Then
                    Exit Function
                End If
            End If
        End If
    Next
    ArraysEqual = True
End Function

Private Function ArrayGreater(ByVal vArray1 As Variant, ByVal vArray2 As Variant) As Boolean
    Dim lIndex As Long
    
    For lIndex = 0 To UBound(vArray1)
        If lIndex > UBound(vArray2) Then
            ArrayGreater = True
            Exit Function
        End If
        If VarType(vArray1(lIndex)) <> (vbArray Or vbVariant) Then
            If VarType(vArray2(lIndex)) <> (vbArray Or vbVariant) Then
                If vArray1(lIndex) < vArray2(lIndex) Then
                    Exit Function
                ElseIf vArray1(lIndex) > vArray2(lIndex) Then
                    ArrayGreater = True
                    Exit Function
                End If
            Else
                Exit Function
            End If
        Else
            If VarType(vArray2(lIndex)) <> (vbArray Or vbVariant) Then
                ArrayGreater = True
                Exit Function
            Else
                ArrayGreater = ArrayGreater(vArray1(lIndex), vArray2(lIndex))
                If ArrayGreater Then
                    Exit Function
                End If
            End If
        End If
    Next
End Function



