Attribute VB_Name = "Controller"
Option Explicit

Public msSnapshotFilePath As String

Public Enum StopReasons
    srNone
    srHardReset
    srSoftReset
    srLoadSnapshot
    srDebugBreak
    srDebugRun
    srPause
    srNormal
End Enum

Public Sub ProcessorStopped()
    ' Debugging.WriteString "Controller.ProcessorStopped"
    
    Console.mbFrozen = False
    Select Case Processor6502.StopReason
        Case srHardReset
            HardResetDo
        Case srSoftReset
            SoftResetDo
        Case srLoadSnapshot
            Snapshot.LoadSnapshot msSnapshotFilePath
            Console.LoadedSnapshotPath = msSnapshotFilePath
            
            Console.DebugOn = False
            Processor6502.StopReason = srNone
            Processor6502.Execute
            ProcessorStopped
        Case srDebugBreak
            Console.mbFrozen = True
            Monitor.DoBreak
        Case srDebugRun
            Console.mbFrozen = True
            Monitor.DoDebugRun
        Case srPause
            Console.mbFrozen = True
        Case srNormal
            Console.DebugOn = False
            Processor6502.StopReason = srNone
            Processor6502.Execute
            ProcessorStopped
    End Select
End Sub

Private Sub HardResetDo()
    ' Debugging.WriteString "Controller.HardResetDo"
    LoadPreferences

    Memory.ResetRam
    InitialiseInterruptLine
    InitialiseSound
    'InitialiseSound2
    InitialiseThrottle
    InitialiseRoms
    InitialiseKeyboard
    ResetSystemVIA
    ResetUserVIA
    Initialise6845
    InitialiseVideoULA
    InitialiseSerialULA
    InitialiseACIA6850
    InitialiseFDC8271
    Processor6502.Initialise6502
    'Processor6502.Initialise6502
    
    Console.DebugOn = False
    Processor6502.Execute
    ProcessorStopped
End Sub

Private Sub SoftResetDo()
    ' Debugging.WriteString "Controller.SoftResetDo"
    
    UserVIA6522.ResetUserVIA
    Processor6502.StopReason = srNone
    Processor6502.RES
    
    Console.DebugOn = False
    Processor6502.Execute
    ProcessorStopped
End Sub

Public Function HexNum(ByVal lNumber As Long, ByVal iPlaces As Integer) As String
    HexNum = Hex$(lNumber)
    If Len(HexNum) <= iPlaces Then
        HexNum = String$(iPlaces - Len(HexNum), "0") & HexNum
    Else
        HexNum = Right$(HexNum, iPlaces)
    End If
End Function

Public Function BaseNum(ByVal lNumber As Long, ByVal iPlaces As Integer, Optional ByVal lBase As Long = 2) As String
    Dim lIndex As Long
    Const sChars As String = "0123456789ABCDEF"
    
    For lIndex = 0 To iPlaces - 1
        BaseNum = Mid$(sChars, (IIf(lNumber < 0, lBase + (lNumber Mod lBase), lNumber Mod lBase)) + 1, 1) & BaseNum
        lNumber = lNumber \ lBase
    Next
End Function

Public Function ConvertBase(ByVal sNumber As String, ByVal lBase As Long) As Long
    Dim lIndex As Long
    Const sChars As String = "0123456789ABCDEF"
    Dim lValue As Long
    
    For lIndex = 1 To Len(sNumber)
        lValue = lValue * lBase
        lValue = lValue + InStr(sChars, Mid$(sNumber, lIndex, 1)) - 1
    Next
    ConvertBase = lValue
End Function

Public Function HexToDec(ByVal sHex As String)
    Dim lIndex As Long
    Dim lValue As Long
    Dim lDigit As Long
    
    Const sChars As String = "0123456789ABCDEF"
        
    HexToDec = -1
    
    For lIndex = 1 To Len(sHex)
        lDigit = InStr(sChars, Mid$(sHex, lIndex, 1))
        If lDigit = 0 Then
            Exit Function
        End If
        
        lValue = lValue * 16
        lValue = lValue + lDigit - 1
    Next
    HexToDec = lValue
End Function

Private Sub LoadPreferences()
    ' Debugging.WriteString "Controller.LoadPreferences"
    
    StorageMedia.ArchiveAsCassette = GetSetting("VBeeb", "Preferences", "ArchiveAsCassette", False)
    Console.mnuCassetteArchive.Checked = StorageMedia.ArchiveAsCassette
    StorageMedia.DiscUnitPreference = GetSetting("VBeeb", "Preferences", "DiscUnitPreference", 0)
    Console.mnuDiscSelectDrive(StorageMedia.DiscUnitPreference).Checked = True
    Console.mnuDiscSelectDrive(1 - StorageMedia.DiscUnitPreference).Checked = False
End Sub
