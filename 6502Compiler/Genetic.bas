Attribute VB_Name = "Genetic"
Option Explicit

Private vInstructionCategories As Variant

Public Sub InitialiseCategories()
    vInstructionCategories = Array()
    AddArray vInstructionCategories, Array("INX", "XREG INC ARITH")
    AddArray vInstructionCategories, Array("INY", "YREG INC ARITH")
    AddArray vInstructionCategories, Array("DEX", "XREG DEC ARITH")
    AddArray vInstructionCategories, Array("DEY", "YREG DEC ARITH")
    AddArray vInstructionCategories, Array("LDA", "AREG MOVE LOAD")
    AddArray vInstructionCategories, Array("LDX", "XREG MOVE LOAD")
    AddArray vInstructionCategories, Array("LDY", "XREG MOVE LOAD")
    AddArray vInstructionCategories, Array("STA", "AREG MOVE STORE")
    AddArray vInstructionCategories, Array("STX", "XREG MOVE STORE")
    AddArray vInstructionCategories, Array("STY", "XREG MOVE STORE")
    
    AddArray vInstructionCategories, Array("AND", "AREG LOGIC AND")
    AddArray vInstructionCategories, Array("ORA", "AREG LOGIC OR")
    AddArray vInstructionCategories, Array("EOR", "AREG LOGIC XOR")
    
    AddArray vInstructionCategories, Array("CMP", "AREG ARITH SUB")
    AddArray vInstructionCategories, Array("ADC", "AREG ARITH ADD")
    AddArray vInstructionCategories, Array("SBC", "AREG ARITH SUB")
    AddArray vInstructionCategories, Array("BIT", "AND LOGIC")
    
End Sub

Private Sub AddArray(vArray As Variant, vItem As Variant)
    Dim lUbound As Long
    
    lUbound = UBound(vArray) + 1
    ReDim Preserve vArray(lUbound)
    vArray(lUbound) = vItem
End Sub

Public Function FindInstructionByCategories(ByVal sCategories As String) As Variant
    Dim lIndex As Long
    Dim vCategories As Variant
    Dim lIndex2 As Long
    Dim sInstructionCats As String
    Dim lCategoryMatches As Long
    
    FindInstructionByCategories = Array()
    vCategories = Split(sCategories, " ")
    
    For lIndex = 0 To UBound(vInstructionCategories)
        sInstructionCats = " " & vInstructionCategories(lIndex)(1) & " "
        lCategoryMatches = 0
        For lIndex2 = 0 To UBound(vCategories)
            If InStr(sInstructionCats, " " & vCategories(lIndex2) & " ") > 0 Then
                lCategoryMatches = lCategoryMatches + 1
            End If
        Next
        If lCategoryMatches > 0 Then
            AddArray FindInstructionByCategories, vInstructionCategories(lIndex)(0)
        End If
    Next
End Function

Public Sub test()
    Dim vIns As Variant
    
    InitialiseCategories
    vIns = FindInstructionByCategories("XREG")
End Sub
