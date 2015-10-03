Attribute VB_Name = "Snapshot"
Option Explicit

Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Private mlBaseAddress As Long

Public Sub SaveMemorySegment(lAddress As Long, yBlock() As Byte)
    ' Debugging.WriteString "Snapshot.SaveMemory"
    
    UEFHandler.ResetBlock 2 + UBound(yBlock) + 1
    CopyMemory UEFHandler.BlockData(0), lAddress, 2&
    CopyMemory UEFHandler.BlockData(2), yBlock(0), UBound(yBlock) + 1
    UEFHandler.SaveBlock &HFF00&
End Sub

Public Sub LoadSnapshot(ByVal sPath As String)
    ' Debugging.WriteString "Snapshot.LoadSnapshot"
    
    If UEFHandler.LoadUEFFile(sPath) Then
        LoadMemory
        LoadMemorySegments
    Else
        MsgBox "Snapshot format not recognised."
    End If
End Sub

Private Sub LoadMemory()
    Dim lAddress As Long
    
    ' Debugging.WriteString "Snapshot.LoadMemory"
    
    UEFHandler.ResetUEF
    If UEFHandler.FindBlock(&H462&) Then
        CopyMemory gyMem(0), UEFHandler.BlockData(0), UEFHandler.BlockLength
    End If
End Sub

Private Sub LoadMemorySegments()
    Dim lAddress As Long
    
    ' Debugging.WriteString "Snapshot.LoadMemorySegments"
    
    UEFHandler.ResetUEF
    While UEFHandler.FindBlock(&HFF00&)
        CopyMemory lAddress, UEFHandler.BlockData(0), 2&
        Debug.Print HexNum(lAddress, 4)
        CopyMemory gyMem(lAddress), UEFHandler.BlockData(2), UEFHandler.BlockLength - 2
    Wend
End Sub
