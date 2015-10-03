Attribute VB_Name = "Module1"
Option Explicit

Sub MAIN()
    RomCompare
End Sub

Private Sub RomCompare()
    Dim f1 As Byte
    Dim f2 As Byte
    Dim x As Long
    Dim bEqual As Boolean
    
    Const path1 As String = "D:\Emulators\Beebem\BeebFile\UnusedRoms\DFS-1_20.ROM"
    Const path2 As String = "D:\Emulators\Beebem\BeebFile\UnusedRoms\DNFS-1_20.rom"
    
    Open path1 For Random As 1
    Open path2 For Random As 2
        
    If FileLen(path1) <> FileLen(path2) Then
        MsgBox "Files not the same"
        Close #1
        Close #2
        Exit Sub
    End If
    
    bEqual = True
    For x = 0 To FileLen(path1)
        Get #1, , f1
        Get #2, , f2
        If f1 <> f2 Then
            bEqual = False
            Exit For
        End If
    Next
    
    If bEqual Then
        MsgBox "Files the same"
    Else
        MsgBox "Files not the same"
    End If
    Close #1
    Close #2
End Sub
