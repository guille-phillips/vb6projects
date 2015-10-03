Attribute VB_Name = "Semantic"
Option Explicit

Private gmSymbolTable() As Collection

Private Enum ElementTypes
    etIdentifier = 1
    etNumber
    etSymbol
    etSet
End Enum

Public Sub Analyse(oTree As SaffronTree)
    Debug.Print RuleAssign(oTree)

End Sub

Private Sub FindRule(oTree)
    
End Sub

Private Function RuleDefine() As Long

End Function

Private Function RuleAssign(oTree As SaffronTree) As Long
    Dim lCursor As Long
    
    lCursor = 1
    If oTree(lCursor).Index <> etIdentifier Then Exit Function Else lCursor = lCursor + 1
    If oTree(lCursor).Text <> ":=" Then Exit Function Else lCursor = lCursor + 1
    If oTree(lCursor).Index <> etNumber Then Exit Function Else lCursor = lCursor + 1
    RuleAssign = lCursor
End Function

