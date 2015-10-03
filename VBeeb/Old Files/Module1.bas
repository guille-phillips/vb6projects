Attribute VB_Name = "Module1"
Option Explicit

Public Sub Test2()
    Dim a(99) As Byte
    Dim b(4, 19) As Byte
    
    Dim lx As Long
    
    For lx = 1 To 100
        a(lx - 1) = lx
    Next
    
    CopyMemory b(0, 0), a(0), 100&
End Sub
