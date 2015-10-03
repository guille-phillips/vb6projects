Attribute VB_Name = "Memory"
Option Explicit

Public gyMem(65535) As Byte
Public RecentlyRead As Long
Public RecentlyWritten As Long

Public Sub ResetRam()
    Dim lClearMem As Long
    
    ' Debugging.WriteString "Memory.ResetRam"
    
    For lClearMem = 0 To &H7FFF&
        gyMem(lClearMem) = 0
    Next
    RecentlyRead = -1
    RecentlyWritten = -1
End Sub

' Perform processing on Write not Read
Public Property Let Mem(ByVal lLocation As Long, ByVal yValue As Byte)
    ' Debugging.WriteString "Memory.Let Mem"
    
    If lLocation >= &HFE00& Then
        Select Case lLocation
            Case &HFC00& To &HFDFF&
                ' do nothing
            Case &HFE00&, &HFE02&, &HFE04&, &HFE06&
                CRTC6845.SelectedRegister = yValue
                gyMem(lLocation) = yValue
            Case &HFE01&, &HFE03&, &HFE05&, &HFE07&
                CRTC6845.WriteRegister yValue
                gyMem(lLocation) = yValue
            Case &HFE08& To &HFE0F&
                ACIA6850.WriteRegister (lLocation - &HFE08&) And &H1&, yValue
            Case &HFE10& To &HFE1F&
                SerialULA.WriteRegister yValue
            Case &HFE20& To &HFE21&
                VideoULA.WriteRegister lLocation - &HFE20&, yValue
                gyMem(lLocation) = yValue
            Case &HFE30&
                RomSelect.SetRom yValue
                gyMem(lLocation) = yValue
            Case &HFE40& To &HFE4F&
                SystemVIA6522.WriteRegister lLocation - &HFE40&, yValue
            Case &HFE50& To &HFE5F&
                SystemVIA6522.WriteRegister lLocation - &HFE50&, yValue
            Case &HFE60& To &HFE6F&
                UserVIA6522.WriteRegister lLocation - &HFE60&, yValue
            Case &HFE70& To &HFE7F&
                UserVIA6522.WriteRegister lLocation - &HFE70&, yValue
            Case &HFE80& To &HFE9F&
                FDC8271.WriteRegister (lLocation - &HFE80&) And &H7&, yValue
        End Select
    ElseIf lLocation >= &HC000& Then
    ElseIf lLocation >= &H8000& Then
        If RomSelect.RomBankWriteable(RomSelect.SelectedBank) Then
            gyMem(lLocation) = yValue
        End If
    Else
        gyMem(lLocation) = yValue
    End If
    
'   If (lLocation > &H6000&) And (lLocation <= &H7FFF&) Then
'        Debug.Print HexNum(Processor6502.PC1 - 1, 4) & ":" & HexNum(lLocation, 4) & ":" & yValue
'        Stop
'    End If

'    If (lLocation = &H206& Or lLocation = &H207&) Then
''        Debug.Print HexNum(Processor6502.PC1 - 1, 4) & ":" & HexNum(lLocation, 4) & ":" & yValue
'        'Debug.Print Processor6502.lTotalCycles
'        Stop
'    End If

'    If gyMem(510) = &H26& And gyMem(511) = &H40& Then
'        Stop
'    End If

    RecentlyWritten = lLocation
End Property

Public Property Get Mem(ByVal lLocation As Long) As Byte
    ' Debugging.WriteString "Memory.Get Mem"
    
'   If (lLocation = &H3000&) Then
'        Debug.Print HexNum(Processor6502.PC1 - 1, 4) & ":" & HexNum(lLocation, 4)
'        'Stop
'    End If

    If (lLocation And &HFE00&) = &HFE00& Then
        Select Case lLocation
            Case &HFE01&
                Mem = CRTC6845.Register(CRTC6845.SelectedRegister)
            Case &HFE08& To &HFE09& ' reset data received flag
                Mem = ACIA6850.ReadRegister(lLocation - &HFE08&)
            Case &HFE40&, &HFE41&, &HFE4F, &HFE44&, &HFE48& ' reset timer interrupts
                Mem = SystemVIA6522.ReadRegister(lLocation - &HFE40&)
            Case &HFE50&, &HFE51&, &HFE5F&, &HFE54&, &HFE58& ' reset timer interrupts
                Mem = SystemVIA6522.ReadRegister(lLocation - &HFE50&)
            Case &HFE60&, &HFE64&, &HFE68& ' reset timer 2 interrupt
                Mem = UserVIA6522.ReadRegister(lLocation - &HFE60&)
            Case &HFE70&, &HFE74&, &HFE78& ' reset timer 2 interrupt
                Mem = UserVIA6522.ReadRegister(lLocation - &HFE70&)
            Case &HFE80& To &HFE84&
                Mem = FDC8271.ReadRegister(lLocation - &HFE80&)
            Case Else
                Mem = gyMem(lLocation)
        End Select
    Else
        Mem = gyMem(lLocation)
    End If
    
'   If (lLocation = &HC7FD&) Then
''        Debug.Print HexNum(Processor6502.PC1 - 1, 4) & ":" & HexNum(lLocation, 4) & ":" & yValue
'        'Debug.Print Processor6502.lTotalCycles
'        'Stop
'    End If

    RecentlyRead = lLocation
End Property

Public Sub SaveMemory(ByVal sPath As String, ByVal lStartAddress As Long, ByVal lEndAddress As Long)
    Dim oFSO As New FileSystemObject
    Dim oMemory() As Byte
    
    If oFSO.FileExists(sPath) Then
        Kill sPath
    End If
    
    ReDim oMemory(lEndAddress - lStartAddress)
    
    CopyMemory oMemory(0), gyMem(lStartAddress), lEndAddress - lStartAddress + 1
    
    Open sPath For Binary As #1
    Put #1, , oMemory
    Close #1
End Sub


Public Sub LoadMemory(ByVal sPath As String, ByVal lStartAddress As Long)
    Dim oFSO As New FileSystemObject
    Dim yMemory() As Byte
    
    ReDim yMemory(FileLen(sPath) - 1)
    
    Open sPath For Binary As #1
    Get #1, , yMemory
    Close #1
    
    CopyMemory gyMem(lStartAddress), yMemory(0), UBound(yMemory) + 1
End Sub
