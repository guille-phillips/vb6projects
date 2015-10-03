VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form Console 
   BackColor       =   &H00000000&
   ClientHeight    =   8190
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   9900
   ForeColor       =   &H00FFFFFF&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MouseIcon       =   "Console.frx":0000
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   546
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   660
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrCursor 
      Interval        =   3000
      Left            =   2040
      Top             =   480
   End
   Begin MSCommLib.MSComm Com 
      Left            =   1080
      Top             =   480
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   0   'False
      OutBufferSize   =   32767
   End
   Begin MSComctlLib.StatusBar staBar 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   7935
      Width           =   9900
      _ExtentX        =   17463
      _ExtentY        =   450
      Style           =   1
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog comFile 
      Left            =   360
      Top             =   480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileLoad 
         Caption         =   "&Load"
      End
      Begin VB.Menu mnuFileSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSaveDisc 
         Caption         =   "Save &Disc"
      End
      Begin VB.Menu mnuFileSaveCassette 
         Caption         =   "Save &Cassette"
      End
      Begin VB.Menu mnuSaveSnapshot 
         Caption         =   "Save &Snapshot"
      End
      Begin VB.Menu mnuFileSaveMemory 
         Caption         =   "Save &Memory"
      End
      Begin VB.Menu mnuFileSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSendSnapshot 
         Caption         =   "Send Snapshot"
      End
      Begin VB.Menu mnuFileSendDisc 
         Caption         =   "Send Disc"
      End
   End
   Begin VB.Menu mnuDisc 
      Caption         =   "Disc"
      Begin VB.Menu mnuDiscInsertBlank 
         Caption         =   "Insert Blank"
      End
      Begin VB.Menu mnuDiscEject 
         Caption         =   "Eject"
      End
      Begin VB.Menu mnuDiscFlipSides 
         Caption         =   "Flip Sides"
      End
      Begin VB.Menu mnuDiscSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDiscSelectDrive 
         Caption         =   "Select Drive 0"
         Index           =   0
      End
      Begin VB.Menu mnuDiscSelectDrive 
         Caption         =   "Select Drive 1"
         Index           =   1
      End
   End
   Begin VB.Menu mnuCassette 
      Caption         =   "Cassette"
      Begin VB.Menu mnuInsertBlank 
         Caption         =   "Insert Blank"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuCassetteEject 
         Caption         =   "Eject"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuCassetteSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCassetteControl 
         Caption         =   "Control"
      End
      Begin VB.Menu mnuCassetteSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCassetteArchive 
         Caption         =   "Load Archive as Cassette"
      End
   End
   Begin VB.Menu mnuSystem 
      Caption         =   "System"
      Begin VB.Menu mnuSlotsAll 
         Caption         =   "Rom Sockets"
         Begin VB.Menu mnuSlots 
            Caption         =   "Socket 1"
            Index           =   0
            Begin VB.Menu mnuSlot0Load 
               Caption         =   "Load"
            End
            Begin VB.Menu mnuSlot0Empty 
               Caption         =   "Empty"
            End
            Begin VB.Menu mnuSlots0RAM 
               Caption         =   "Sideways RAM"
            End
         End
         Begin VB.Menu mnuSlots 
            Caption         =   "Socket 2"
            Index           =   1
            Begin VB.Menu mnuSlot1Load 
               Caption         =   "Load"
            End
            Begin VB.Menu mnuSlot1Empty 
               Caption         =   "Empty"
            End
            Begin VB.Menu mnuSlots1RAM 
               Caption         =   "Sideways RAM"
            End
         End
         Begin VB.Menu mnuSlots 
            Caption         =   "Socket 3"
            Index           =   2
            Begin VB.Menu mnuSlot2Load 
               Caption         =   "Load"
            End
            Begin VB.Menu mnuSlot2Empty 
               Caption         =   "Empty"
            End
            Begin VB.Menu mnuSlots2RAM 
               Caption         =   "Sideways RAM"
            End
         End
         Begin VB.Menu mnuSlots 
            Caption         =   "Socket 4"
            Index           =   3
            Begin VB.Menu mnuSlot3Load 
               Caption         =   "Load"
            End
            Begin VB.Menu mnuSlot3Empty 
               Caption         =   "Empty"
            End
            Begin VB.Menu mnuSlots3RAM 
               Caption         =   "Sideways RAM"
            End
         End
      End
      Begin VB.Menu mnuSystemKeyboardLinks 
         Caption         =   "Keyboard Links"
      End
      Begin VB.Menu mnuSystemSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSystemReboot 
         Caption         =   "Reboot"
      End
   End
   Begin VB.Menu mnuDebug 
      Caption         =   "&Debug"
      Begin VB.Menu mnuDebugBreak 
         Caption         =   "Break"
      End
      Begin VB.Menu mnuDebugMode 
         Caption         =   "Debug Mode"
      End
      Begin VB.Menu mnuDebugSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDebugRAM 
         Caption         =   "RAM"
      End
   End
   Begin VB.Menu mnuVBeeb 
      Caption         =   "VBeeb"
      Begin VB.Menu mnuVBeebSpeed 
         Caption         =   "Speed"
      End
      Begin VB.Menu mnuVBeebPause 
         Caption         =   "Pause"
      End
      Begin VB.Menu mnuVBeebSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuVBeebKeyboardMappings 
         Caption         =   "Keyboard Mappings"
      End
      Begin VB.Menu mnuVBeebSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
      Begin VB.Menu mnuVBeebSep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "Console"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public mlPreviousKey As Long

Public mlMeHDC As Long

Private mbActivated As Boolean

Private moDiscStorage As Storage
Private mlCursorHandle As Long

Private mbCursorOff As Boolean
Private mbDontSaveDimensions As Boolean

Private Const EmulatorVersion = "VBeeb 2.1"
Private msLoadedSnapshotPath As String
Private mbDebugOn As Boolean

Public mbFrozen As Boolean

Private Sub Form_Activate()
    ' Debugging.WriteString "Console.Form_Activate"
      
    If Not mbActivated Then
        mbActivated = True
        Processor6502.StopReason = srHardReset
        ProcessorStopped
    End If
End Sub

Private Sub Form_Initialize()
    'Debugging.OpenDebugFile
    ' Debugging.WriteString "Console.Form_Initialise"
    
    mbDontSaveDimensions = True
    Me.Width = GetSetting("VBeeb", "Dimensions", "Width", Me.Width)
    Me.Height = GetSetting("VBeeb", "Dimensions", "Height", Me.Height)
    mbDontSaveDimensions = False
    
    'Shell "regsvr32 /s " & App.path & "\MSCOMM32.OCX"
    Shell "regsvr32 -u /s " & App.path & "\DLLs\SaffronCompiler.dll"
    Shell "regsvr32 -u /s " & App.path & "\DLLs\SaffronClasses.dll"
    Shell "regsvr32 /s " & App.path & "\DLLs\SaffronCompiler.dll"
    Shell "regsvr32 /s " & App.path & "\DLLs\SaffronClasses.dll"

    Shell "regsvr32 /s " & App.path & "\DLLs\comdlg32.OCX"
    mlMeHDC = Me.hdc
    UpdateCaption
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = mlPreviousKey Then
        Exit Sub
    End If
    mlPreviousKey = KeyCode
    
    If Keyboard.lMapping(KeyCode) = 256 Then ' BREAK
        Processor6502.StopReason = srSoftReset
    Else
        Keyboard.WindowsKeyDown KeyCode
    End If
    
    KeyCode = 0
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If Keyboard.lMapping(KeyCode) <> 256 And Keyboard.mbInitialised Then
        Keyboard.WindowsKeyUp KeyCode
    End If
    mlPreviousKey = 0
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Monitor.Visible Then
        Monitor.AddMemory VideoULA.GetAddress(X, Y)
    End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Static nOldX As Single
    Static nOldY As Single
    
    Dim bMoved As Boolean
    
    If X <> nOldX Then
        nOldX = X
        bMoved = True
    End If
    
    If Y <> nOldY Then
        nOldY = Y
        bMoved = True
    End If
    
    If bMoved Then
        If mbCursorOff Then
            Me.MousePointer = vbDefault
            tmrCursor.Enabled = False
            tmrCursor.Enabled = True
            mbCursorOff = False
        End If
    
        VideoULA.UpdateLightPenPosition X, Y
    End If
End Sub

Private Sub Form_Paint()
    If mbFrozen Then
        VideoULA.UpdateWholeDisplay
    End If
End Sub

Private Sub mnuDebugRAM_Click()
    Search.Show vbModal
End Sub

Private Sub mnuFileSaveMemory_Click()
    Dim sPath As String
    Dim oSaveMemory As New SaveMemory
    
    On Error GoTo Exit_mnuFileSaveMemory_Click
    
    oSaveMemory.Mode = mtSave
    oSaveMemory.Show vbModal
    
    If oSaveMemory.mlStartAddress = -1 Or oSaveMemory.mlEndAddress = -1 Then
        Exit Sub
    End If
    
    sPath = GetSetting("VBeeb", "Paths", "Memory", App.path)

    comFile.FileName = sPath
    comFile.Filter = "Memory (*.mem)|*.mem"
    comFile.CancelError = True
    comFile.ShowSave
    SaveSetting "VBeeb", "Paths", "Memory", GetFilePath(comFile.FileName)
    
    If oSaveMemory.mlStartAddress <> -1 And oSaveMemory.mlEndAddress <> -1 Then
        Memory.SaveMemory comFile.FileName, oSaveMemory.mlStartAddress, oSaveMemory.mlEndAddress
    End If

Exit_mnuFileSaveMemory_Click:
End Sub



Private Sub mnuVBeebKeyboardMappings_Click()
    KeyboardMappings.Show vbModal
End Sub

Private Sub mnuVBeebPause_Click()
    mnuVBeebPause.Checked = Not mnuVBeebPause.Checked
    If mnuVBeebPause.Checked Then
        Processor6502.StopReason = srPause
    Else
        Processor6502.StopReason = srNormal
        Controller.ProcessorStopped
    End If
End Sub

Private Sub mnuVBeebSpeed_Click()
    Speed.Show
End Sub

Private Sub tmrCursor_Timer()
    If Not mbCursorOff Then
        Me.MousePointer = vbCustom
        mbCursorOff = True
    End If
End Sub

Private Sub Form_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim vFile As Variant
    Dim oStorage As New Storage
    
    If Data.GetFormat(vbCFFiles) Then
        For Each vFile In Data.Files
            Select Case oStorage.LoadFile(vFile, StorageMedia.ArchiveAsCassette)
                Case ftCassette
                    Set StorageMedia.CassetteStorage = oStorage
                Case ftDisc, ftDiscDSD
                    Set StorageMedia.DiscStorage(StorageMedia.DiscUnitPreference) = oStorage
                Case ftArchive
                    If Not StorageMedia.ArchiveAsCassette Then
                       Set StorageMedia.DiscStorage(StorageMedia.DiscUnitPreference) = oStorage
                    Else
                        oStorage.BuildCassetteImage True
                        Set StorageMedia.CassetteStorage = oStorage
                    End If
            End Select
        Next
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Snapshot.SaveSnapshot App.path & "\snapshots\quick.uef"
    Keyboard.TerminateKeyboard
    Unload Search
End Sub

Private Sub Form_Resize()
    ' Debugging.WriteString "Console.Form_Resize"
    CRTC6845.DisplayDimensionsChanged = True
    VideoULA.NewScreenWidth = Me.ScaleWidth
    VideoULA.NewScreenHeight = Me.ScaleHeight
    
    If Not mbDontSaveDimensions Then
        SaveSetting "VBeeb", "Dimensions", "Width", Me.Width
        SaveSetting "VBeeb", "Dimensions", "Height", Me.Height
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload Me
    End
End Sub

Private Sub mnuAbout_Click()
    About.Show vbModal
End Sub

Private Sub mnuCassetteArchive_Click()
    mnuCassetteArchive.Checked = Not mnuCassetteArchive.Checked
    StorageMedia.ArchiveAsCassette = mnuCassetteArchive.Checked
    SaveSetting "VBeeb", "Preferences", "ArchiveAsCassette", StorageMedia.ArchiveAsCassette
End Sub

Private Sub mnuCassetteControl_Click()
    If StorageMedia.CassetteStorage Is Nothing Then
        MsgBox "No Cassette Loaded"
    Else
        StorageMedia.CassetteStorage.ShowCassetteCatalogue
    End If
End Sub

Private Sub mnuEjectCassette_Click()
    ACIA6850.EjectCassette
End Sub

Private Sub mnuCassetteEject_Click()
    Set StorageMedia.CassetteStorage = Nothing
End Sub

Private Sub mnuDebugBreak_Click()
    DebugOn = -True
    Processor6502.StopReason = srDebugBreak
End Sub

Private Sub mnuDebugMode_Click()
    mnuDebugMode.Checked = Not mnuDebugMode.Checked
    If mnuDebugMode.Checked Then
        Processor6502.StopReason = srDebugRun
    Else
        Processor6502.StopReason = srNormal
    End If
End Sub

Private Sub mnuDiscEject_Click()
'    FDC8271.InitialiseDisc
'    Set StorageMedia.DiscStorage(StorageMedia.DiscUnitPreference) = Nothing
End Sub

Private Sub mnuDiscFlipSides_Click()
    If StorageMedia.DiscStorage(StorageMedia.DiscUnitPreference) Is Nothing Then
        MsgBox "Disc is not present", vbOKOnly Or vbInformation, "VBeeb Flip Sides"
    Else
        StorageMedia.DiscStorage(StorageMedia.DiscUnitPreference).FlipSides
    End If
End Sub

Private Sub mnuDiscInsertBlank_Click()
    If StorageMedia.DiscStorage(StorageMedia.DiscUnitPreference) Is Nothing Then
        Set StorageMedia.DiscStorage(StorageMedia.DiscUnitPreference) = New Storage
    End If
    StorageMedia.DiscStorage(StorageMedia.DiscUnitPreference).ClearDiskImage
    FDC8271.InitialiseDisc
End Sub

Private Sub mnuDiscSelectDrive_Click(Index As Integer)
    StorageMedia.DiscUnitPreference = Index
    mnuDiscSelectDrive(Index).Checked = True
    mnuDiscSelectDrive(1 - Index).Checked = False
    SaveSetting "VBeeb", "Preferences", "DiscUnitPreference", Index
End Sub

Private Sub mnuExit_Click()
    Unload Me
    End
End Sub

Private Sub mnuFileLoad_Click()
    Dim oStorage As New Storage
    Dim sPath As String
    Dim yImage() As Byte
    
    On Error GoTo mnuFileLoadExit
    sPath = GetSetting("VBeeb", "Paths", "Files", App.path)
    
    comFile.CancelError = True
    comFile.FileName = sPath
    comFile.Filter = "*.uef; *.img; *.ssd; *.dsd; *.mem;"
    comFile.ShowOpen

    SaveSetting "VBeeb", "Paths", "Files", GetFilePath(comFile.FileName)
    
    Select Case oStorage.LoadFile(comFile.FileName, StorageMedia.ArchiveAsCassette)
        Case ftCassette
            Set StorageMedia.CassetteStorage = oStorage
            SaveSetting "VBeeb", "Paths", "Cassettes", GetFilePath(comFile.FileName)
        Case ftDisc, ftDiscDSD
            Set StorageMedia.DiscStorage(StorageMedia.DiscUnitPreference) = oStorage
            SaveSetting "VBeeb", "Paths", "Discs", GetFilePath(comFile.FileName)
        Case ftArchive
            If Not StorageMedia.ArchiveAsCassette Then
               Set StorageMedia.DiscStorage(StorageMedia.DiscUnitPreference) = oStorage
               SaveSetting "VBeeb", "Paths", "Cassettes", GetFilePath(comFile.FileName)
            Else
                oStorage.BuildCassetteImage True
                Set StorageMedia.CassetteStorage = oStorage
                SaveSetting "VBeeb", "Paths", "Discs", GetFilePath(comFile.FileName)
            End If
        Case ftSnapshot
            SaveSetting "VBeeb", "Paths", "Snapshots", GetFilePath(comFile.FileName)
    End Select

mnuFileLoadExit:
End Sub

Private Sub mnuFileSaveCassette_Click()
    Dim sPath As String
    
    On Error GoTo Exit_mnuFileSaveCassette_Click
    
    If StorageMedia.CassetteStorage Is Nothing Then
        MsgBox "Cassette is not present", vbOKOnly Or vbInformation, "VBeeb Save Cassette"
        Exit Sub
    End If
    
    sPath = GetSetting("VBeeb", "Paths", "Cassettes", App.path)

    comFile.FileName = sPath
    comFile.Filter = "Cassette (*.uef)|*.uef|Disc (*.ssd; *.dsd)|*.ssd|Archive (*.inf)|*.inf"
    comFile.CancelError = True
    comFile.ShowSave
    SaveSetting "VBeeb", "Paths", "Cassettes", GetFilePath(comFile.FileName)
    Select Case comFile.FilterIndex
        Case 1
            StorageMedia.CassetteStorage.SaveCassetteUEF comFile.FileName
            '    If Right$(LCase$(comFile.FileName), 4) <> ".uef" Then
            
            '        comFile.FileName = comFile.FileName + ".uef"
            '    End If
        Case 2
        Case 3
    End Select
    
Exit_mnuFileSaveCassette_Click:
End Sub

Private Sub mnuFileSaveDisc_Click()
    Dim sPath As String
    
    On Error GoTo Exit_mnuFileSaveDisc_Click
    
    If StorageMedia.DiscStorage(StorageMedia.DiscUnitPreference) Is Nothing Then
        MsgBox "Disc is not present", vbOKOnly Or vbInformation, "VBeeb Save Disc"
        Exit Sub
    End If
    
    sPath = GetSetting("VBeeb", "Paths", "Discs", App.path)

    comFile.FileName = sPath
    comFile.Filter = "Disc (*.ssd)|*.ssd|Disc (*.dsd)|*.dsd|Archive (*.inf)|*.inf"
    comFile.CancelError = True
    comFile.ShowSave
    SaveSetting "VBeeb", "Paths", "Discs", GetFilePath(comFile.FileName)
    Select Case comFile.FilterIndex
        Case 1 ' ssd
            StorageMedia.DiscStorage(StorageMedia.DiscUnitPreference).SaveDiscImage comFile.FileName, False
        Case 2 ' dsd
            StorageMedia.DiscStorage(StorageMedia.DiscUnitPreference).SaveDiscImage comFile.FileName, True
        Case 3 ' inf
            StorageMedia.DiscStorage(StorageMedia.DiscUnitPreference).SaveArchive comFile.FileName, False
    End Select
    
Exit_mnuFileSaveDisc_Click:
End Sub

Private Sub mnuSaveSnapshot_Click()
    Dim sPath As String
    
    On Error GoTo Exit_mnuSaveSnapshot
    
    sPath = GetSetting("VBeeb", "Paths", "Snapshots", App.path & "\Snapshots\*.*")
    
    comFile.CancelError = True
    comFile.FileName = sPath
    comFile.Filter = "*.uef"
    comFile.ShowSave
    If Right$(LCase$(comFile.FileName), 4) <> ".uef" Then
        SaveSetting "VBeeb", "Paths", "Snapshots", GetFilePath(comFile.FileName)
        comFile.FileName = comFile.FileName & ".uef"
    End If
    Snapshot.SaveSnapshot comFile.FileName
Exit_mnuSaveSnapshot:
End Sub

Private Sub mnuFileSendSnapshot_Click()
    Dim sPath As String
    
    On Error GoTo Exit_mnuFileSendSnapshot
    
    sPath = GetSetting("VBeeb", "Paths", "Snapshots", App.path & "\Snapshots\*.*")
    comFile.CancelError = True
    comFile.FileName = sPath
    comFile.Filter = "*.uef"
    comFile.ShowSave
    If Right$(LCase$(comFile.FileName), 4) <> ".uef" Then
        comFile.FileName = comFile.FileName & ".uef"
    End If
    
    SaveSetting "VBeeb", "Paths", "Snapshots", GetFilePath(comFile.FileName)
    Snapshot.StartTransfer comFile.FileName
    Snapshot.LoadSnapshot comFile.FileName
Exit_mnuFileSendSnapshot:
End Sub

Private Sub mnuFileSendDisc_Click()
    Dim sPath As String
    
    On Error GoTo Exit_mnuFileSendDisc
    
'    sPath = GetSetting("VBeeb", "Paths", "Discs", App.path & "\Discs\*.*")
'    comFile.CancelError = True
'    comFile.FileName = sPath
'    comFile.Filter = "*.uef"
'    comFile.ShowSave
'    Select Case Right$(LCase$(comFile.FileName), 4)
'        Case ".ssd", ".dsd"
'        Case Else
'            comFile.FileName = comFile.FileName & ".ssd"
'    End Select
'    SaveSetting "VBeeb", "Paths", "Discs", GetFilePath(comFile.FileName)

    StorageMedia.DiscStorage(StorageMedia.DiscUnitPreference).StartTransfer comFile.FileName
Exit_mnuFileSendDisc:
End Sub

Private Sub mnuInsertBlank_Click()
'    comFile.flags = 0
'    comFile.FileName = App.path & "\Tapes\Blank.UEF"
'    comFile.Filter = "*.uef"
'    comFile.ShowSave
'    If comFile.flags And 1024 = 1024 Then
'        ACIA6850.LoadBlankTape comFile.FileName
'    End If
End Sub

Private Sub mnuSlots0RAM_Click()
    mnuSlots0RAM.Checked = True
    RomConfigure.SetROMWritable 12, True
    mnuSlot0Empty.Checked = False
    mnuSlots(0).Caption = "RAM"
End Sub

Private Sub mnuSlots1RAM_Click()
    mnuSlots1RAM.Checked = True
    RomConfigure.SetROMWritable 13, True
    mnuSlot1Empty.Checked = False
    mnuSlots(1).Caption = "RAM"
End Sub

Private Sub mnuSlots2RAM_Click()
    mnuSlots2RAM.Checked = True
    RomConfigure.SetROMWritable 14, True
    mnuSlot2Empty.Checked = False
    mnuSlots(2).Caption = "RAM"
End Sub

Private Sub mnuSlots3RAM_Click()
    mnuSlots3RAM.Checked = True
    RomConfigure.SetROMWritable 15, True
    mnuSlot3Empty.Checked = False
    mnuSlots(3).Caption = "RAM"
End Sub

Private Sub mnuSystemKeyboardLinks_Click()
    KeyboardLinks.Show vbModal
End Sub

Private Function GetFilePath(ByVal sFileName As String) As String
    Dim lSlash As Long
    
    lSlash = InStrRev(sFileName, "\")
    GetFilePath = Left$(sFileName, lSlash) & "*.*"
End Function




Private Sub mnuSystemReboot_Click()
    Keyboard.ClearPressedKeys
    If Processor6502.StopReason = srDebugBreak Then
        Monitor.Hide
        DebugOn = False
        Processor6502.StopReason = srHardReset
        Controller.ProcessorStopped
    ElseIf Processor6502.StopReason = srPause Then
        mnuVBeebPause.Checked = False
        Processor6502.StopReason = srHardReset
        Controller.ProcessorStopped
    Else
        Processor6502.StopReason = srHardReset
    End If
End Sub



Private Sub mnuSlot0Empty_Click()
    RomConfigure.EmptyRom 12
    RomConfigure.SetROMWritable 12, False
    mnuSlot0Empty.Checked = True
    mnuSlots0RAM.Checked = False
End Sub

Private Sub mnuSlot1Empty_Click()
    RomConfigure.EmptyRom 13
    RomConfigure.SetROMWritable 13, False
    mnuSlot1Empty.Checked = True
    mnuSlots1RAM.Checked = False
End Sub

Private Sub mnuSlot2Empty_Click()
    RomConfigure.EmptyRom 14
    RomConfigure.SetROMWritable 14, False
    mnuSlot2Empty.Checked = True
    mnuSlots2RAM.Checked = False
End Sub

Private Sub mnuSlot3Empty_Click()
    RomConfigure.EmptyRom 15
    RomConfigure.SetROMWritable 15, False
    mnuSlot3Empty.Checked = True
    mnuSlots3RAM.Checked = False
End Sub


Private Sub mnuSlot0Load_Click()
    If SlotLoad(12) Then
        mnuSlot0Empty.Checked = False
        mnuSlots0RAM.Checked = False
    End If
End Sub

Private Sub mnuSlot1Load_Click()
    If SlotLoad(13) Then
        mnuSlot0Empty.Checked = False
        mnuSlots0RAM.Checked = False
    End If
End Sub

Private Sub mnuSlot2Load_Click()
    If SlotLoad(14) Then
        mnuSlot0Empty.Checked = False
        mnuSlots0RAM.Checked = False
    End If
End Sub

Private Sub mnuSlot3Load_Click()
    If SlotLoad(15) Then
        mnuSlot0Empty.Checked = False
        mnuSlots0RAM.Checked = False
    End If
End Sub


Private Sub mnuSlotsEmpty_Click(Index As Integer)
    RomConfigure.EmptyRom Index
End Sub

Private Function SlotLoad(ByVal lSocketNumber) As Boolean
    On Error GoTo SlotLoadExit
    
    comFile.CancelError = True
    comFile.FileName = App.path & "\Roms\*.*"
    comFile.Filter = "*.rom"
    comFile.ShowOpen
    RomConfigure.SetROMWritable lSocketNumber, False
    RomConfigure.LoadRom lSocketNumber, comFile.FileName

    SlotLoad = True
SlotLoadExit:
End Function


Private Sub mnuFile_Click()
    Sound.PauseSound
End Sub

Private Sub mnuDisc_Click()
    Sound.PauseSound
End Sub

Private Sub mnuCassette_Click()
    Sound.PauseSound
End Sub

Private Sub mnuSystem_Click()
    Sound.PauseSound
End Sub

Private Sub mnuVBeeb_Click()
    Sound.PauseSound
End Sub



Public Property Let LoadedSnapshotPath(ByVal sPath As String)
    msSnapshotFilePath = sPath
    UpdateCaption
End Property

Public Property Let DebugOn(ByVal bDebugOn As Boolean)
    mbDebugOn = bDebugOn
    mnuDebugMode.Checked = bDebugOn
    UpdateCaption
End Property

Private Sub UpdateCaption()
    Me.Caption = App.ProductName & IIf(msSnapshotFilePath <> "", " [" & msSnapshotFilePath & "]", "") & IIf(mbDebugOn, " - DEBUG", "")
End Sub
