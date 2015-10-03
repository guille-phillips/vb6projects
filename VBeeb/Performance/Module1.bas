Attribute VB_Name = "Module1"
Option Explicit

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Private N As Long
Private PC As Long
Private PC1 As Long
Private lCycles As Long
Private mlRelative(255) As Long

Sub Main()
    Dim lCount As Long
    Dim lCount2 As Long
    Dim dTime As Double
    Dim dFastestTime As Double
    Const lIterations As Long = 1000000
    Const bit As Long = 128&
    
    Dim lValue As Long
    Dim lLocation As Long
    
    For lCount = 0 To 255
        If lCount < 128 Then
            mlRelative(lCount) = lCount
        Else
            mlRelative(lCount) = lCount - 256&
        End If
    Next
    
    N = 0
    
    dFastestTime = 1000000#
    For lCount2 = 1 To 1000
        StartCounter
        For lCount = 1 To lIterations

            'If N = bit Then PC = PC1: lCycles = 2& Else PC = PC1 + gyMem(PC1) + (gyMem(PC1) > 127) * 256&: lCycles = 3&
            PC = PC1 + mlRelative(gyMem(PC1))
            'lValue = gyMem(lLocation) + gyMem(lLocation + 1) * 256&
            'CopyMemory lValue, gyMem(lLocation), 2&
        Next
        dTime = CDbl(GetCounter) / CDbl(lIterations)
        If dTime < dFastestTime Then
            dFastestTime = dTime
        End If
    Next

    Debug.Print Scientific(dFastestTime)
    If Dir(App.Path & "\perf.txt") <> "" Then Kill App.Path & "\perf.txt"
    Open App.Path & "\perf.txt" For Binary As #2
    Put #2, , Scientific(dFastestTime)
    Close #2
End Sub
