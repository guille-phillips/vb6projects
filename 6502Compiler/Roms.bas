Attribute VB_Name = "Roms"
Option Explicit

Public Sub LoadRom(ByVal sRomFilePath As String)
    Dim lLen As Long
    Dim lMem As Long
    Dim yValue As Byte
    
    Dim sFileName As String
    Dim vSplitPath As Variant
    Dim lDot As Long
    
    lLen = FileLen(sRomFilePath)
    Open sRomFilePath For Binary As #1
    For lMem = 0 To lLen - 1
        Get #1, , yValue
        gyMem(&H8000& + lMem) = yValue
    Next
    Close #1
End Sub
