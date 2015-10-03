Attribute VB_Name = "RomConfigure"
Option Explicit


Public Sub InitialiseRoms()
    Dim lBank As Long
    Dim lMem As Long
    Dim sRomPath As String
    
    Dim sFileName As String
    Dim vSplitPath As Variant
    Dim lDot As Long
    Dim lRomCount As Long
    
    ' Debugging.WriteString "RomConfigure.InitialiseRoms"
    
    On Error GoTo PathNotFound
    
    For lBank = 0 To 15
       RomBank(3, lBank) = &H60
    Next
    
    LoadOSRom
    
    If GetSetting("VBeeb", "Preferences", "FirstTime", "y") = "y" Then
        SaveSetting "VBeeb", "ROMs", "15", App.path & "\Roms\basic2.rom"
        SaveSetting "VBeeb", "ROMs", "14", App.path & "\Roms\dnfs.rom"
        SaveSetting "VBeeb", "Preferences", "FirstTime", "n"
    End If
    
    For lBank = 12 To 15
        sRomPath = GetSetting("VBeeb", "ROMs", CStr(lBank), "")

        If sRomPath <> "" Then
            If Dir$(sRomPath) <> "" Then
                LoadRom lBank, sRomPath
                
                vSplitPath = Split(sRomPath, "\")
                sFileName = vSplitPath(UBound(vSplitPath))
                lDot = InStrRev(sFileName, ".")
                If lDot > 0 Then
                    Console.mnuSlots(lBank - 12).Caption = Left$(sFileName, lDot - 1)
                Else
                    Console.mnuSlots(lBank - 12).Caption = sFileName
                End If
                
                Console.mnuSlots(lBank - 12).Caption = sFileName
                lRomCount = lRomCount + 1
            Else
                Console.mnuSlots(lBank - 12).Caption = "Socket " & lBank - 11
            End If
        End If
        RomBankWriteable(lBank) = GetSetting("VBeeb", "ROMs", "Writeable" & CStr(lBank), False)
        Select Case lBank
            Case 12
                Console.mnuSlots0RAM.Checked = RomBankWriteable(lBank)
            Case 13
                Console.mnuSlots1RAM.Checked = RomBankWriteable(lBank)
            Case 14
                Console.mnuSlots2RAM.Checked = RomBankWriteable(lBank)
            Case 15
                Console.mnuSlots3RAM.Checked = RomBankWriteable(lBank)
        End Select
        If RomBankWriteable(lBank) Then
            Console.mnuSlots(lBank - 12).Caption = "RAM"
        End If
    Next
    RomSelect.SetRom 15
    Exit Sub
PathNotFound:
End Sub

Public Sub LoadRom(ByVal lSocketNumber As Long, ByVal sRomFilePath As String)
    Dim lLen As Long
    Dim lMem As Long
    Dim yValue As Byte
    
    Dim sFileName As String
    Dim vSplitPath As Variant
    Dim lDot As Long
    
    ' Debugging.WriteString "RomConfigure.LoadRom"
    
    SaveSetting "VBeeb", "ROMs", CStr(lSocketNumber), sRomFilePath
    vSplitPath = Split(sRomFilePath, "\")
    sFileName = vSplitPath(UBound(vSplitPath))
    lDot = InStrRev(sFileName, ".")
    If lDot > 0 Then
        Console.mnuSlots(lSocketNumber - 12).Caption = Left$(sFileName, lDot - 1)
    Else
        Console.mnuSlots(lSocketNumber - 12).Caption = sFileName
    End If
    
    lLen = FileLen(sRomFilePath)
    Open sRomFilePath For Binary As #1
    For lMem = 0 To lLen - 1
        Get #1, , yValue
        RomSelect.RomBank(lMem, lSocketNumber) = yValue
    Next
    Close #1
End Sub

Private Sub LoadOSRom()
    Dim lLen As Long
    Dim lMem As Long
    Dim yValue As Byte

    ' Debugging.WriteString "RomConfigure.LoadOSRom"
    
    lLen = FileLen(App.path & "\Roms\os12.rom")
    Open App.path & "\Roms\os12.rom" For Binary As #1
    For lMem = 0 To lLen - 1
        Get #1, , yValue
        gyMem(lMem + &HC000&) = yValue
    Next
    Close #1
    
'    Dim sCode As String
'    sCode = Disassemble(&HC000&, &HFFFF&)
'    Open App.Path & "\Disassembly\electron_os.txt" For Binary As #1
'    Put #1, , sCode
'    Close #1
    
    For lMem = &HFC00& To &HFDFF&
        gyMem(lMem) = lMem And &HFF&
    Next
End Sub

Public Sub EmptyRom(ByVal lSocketNumber As Long)
    Dim lMem As Long
    
    ' Debugging.WriteString "RomConfigure.EmptyRom"
    
    On Error Resume Next
    
    For lMem = 0 To 16383
        RomSelect.RomBank(lMem, lSocketNumber) = 0
    Next
    SaveSetting "VBeeb", "ROMs", CStr(lSocketNumber), ""
    Console.mnuSlots(lSocketNumber - 12).Caption = "Socket " & lSocketNumber - 11
End Sub

Public Sub SetROMWritable(ByVal lSocketNumber As Long, ByVal bWriteable As Boolean)
    SaveSetting "VBeeb", "ROMs", "Writeable" & CStr(lSocketNumber), CStr(CLng(bWriteable))
    RomSelect.RomBankWriteable(lSocketNumber) = bWriteable
    If bWriteable Then
        EmptyRom lSocketNumber
    End If
End Sub
