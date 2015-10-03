Attribute VB_Name = "RomSelect"
Option Explicit

Public RomBank(16384, 15) As Byte
Public RomBankWriteable(15) As Boolean
Public SelectedBank As Long

Public Sub SetRom(ByVal lRomSocket As Long)
    ' Debugging.WriteString "RomSelect.SetRom"
    
    lRomSocket = lRomSocket And &HF&
    If RomBankWriteable(SelectedBank) Then
      CopyMemory RomBank(0, SelectedBank), gyMem(&H8000&), 16384&
    End If
    
    SelectedBank = lRomSocket
    CopyMemory gyMem(&H8000&), RomBank(0, lRomSocket), 16384&
End Sub
