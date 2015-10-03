Attribute VB_Name = "Beebem8271"
'Option Explicit
'
'Private TorchTube As Long
'
'Private Disc8271Trigger As Long  ' ' Cycle based time Disc8271Trigger
'Private ResultReg As Byte
'Private StatusReg As Byte
'Private DataReg As Byte
'Private Internal_Scan_SectorNum As Byte
'Private Internal_Scan_Count As Long  ' Read as two bytes
'Private Internal_ModeReg As Byte ';
'Private Internal_CurrentTrack(1) As Byte '; ' 0/1 for surface number
'Private Internal_DriveControlOutputPort As Byte ';
'Private Internal_DriveControlInputPort As Byte ';
'Private Internal_BadTracks(1, 1) As Byte ' ' 1st subscript is surface 0/1 and second subscript is badtrack 0/1
'
'Private ThisCommand As Long
'Private NParamsInThisCommand As Long
'Private PresentParam As Long ' ' From 0
'Private Params(15) As Byte  ' Wildly more than we need
'Private Selects(2) As Long ' Drive selects
'Private Writeable(1) As Long ' ' True if the drives are writeable
'
'Private FirstWriteInt As Long ' ' Indicates the start of a write operation
'
'Private NextInterruptIsErr As Long  ' ' none 0 causes error and drops this value into result reg
'Const TRACKSPERDRIVE As Long = 80
'
'' Note Head select is done from bit 5 of the drive output register
''#define CURRENTHEAD ((Internal_DriveControlOutputPort>>5) & 1)
'
'' Note: reads/writes one byte every 80us
'Const TIMEBETWEENBYTES = 160
'
'Type IDField
'    CylinderNum As Long ' 7
'    RecordNum As Long '5
'    HeadNum As Long '1
'    PhysRecLength As Long
'End Type
'
'Type SectorType
'    IDFieldx As IDField
'    Deleted As Long ' 1 ' If non-zero the sector is deleted
'    Data As Variant
'End Type
'
'Type TrackType
'    LogicalSectors As Long
'    NSectors As Long
'    Sectors As SectorType
'    Gap1Size As Long
'    Gap3Size As Long
'    Gap5Size As Long ' From format command
'End Type
'
'' All data on the disc - first param is drive number, then head. then physical track id
'Private DiscStore(1, 1, TRACKSPERDRIVE) As TrackType
'
'' File names of loaded disc images
'Private FileNames(1, 255) As Byte
'
'' Number of sides of loaded disc images
'Private NumHeads(1) As Long
'
''private sub SaveTrackImage(DriveNum as long, int HeadNum, int TrackNum);
'
''typedef void (*CommandFunc)();
'
'Type CommandStatus
'    TrackAddr As Long
'    CurrentSector As Long
'    SectorLength As Long ' ' In bytes
'    SectorsToGo As Long
'    CurrentSectorPtr As SectorType
'    CurrentTrackPtr As TrackType
'    ByteWithinSector  As Long ' ' Next byte in sector or ID field
'End Type
'
''
'Type PrimaryCommandLookupType
'    CommandNum As Byte
'    Mask As Byte ' ' Mask command with this before comparing with CommandNum - allows drive ID to be removed
'    NParams As Long ' ' Number of parameters to follow
'    ToCall As CommandFunc '; ' Called after all paameters have arrived
'    IntHandler As CommandFunc '; ' Called when interrupt requested by command is about to happen
'    Ident As Variant ' char *Ident; ' Mainly for debugging
'End Type
'
'
''#define UPDATENMISTATUS if (StatusReg & 8) NMIStatus |=1<<nmi_floppy; else NMIStatus &= ~(1<<nmi_floppy);
'Private Sub UpdateNMIStatus()
'    If (StatusReg & 8) Then
'        NMIStatus = NMIStatus Or 2 ^ nmi_floppy
'    Else
'        NMIStatus = NMIStatus and  ~(2 ^ nmi_floppy)
'    End If
'End Sub
'
'' For appropriate commands checks the select bits in the command code and
'' selects the appropriate drive.
'Private Sub DoSelects()
'    Selects(0) = (ThisCommand And &H40&) <> 0
'    Selects(1) = (ThisCommand And &H80&) <> 0
'    Internal_DriveControlOutputPort& = &H3F&
'
'    If Selects(0) Then
'       Internal_DriveControlOutputPort = Internal_DriveControlOutputPort Or &H40&
'    End If
'
'    If Selects(1) Then
'        Internal_DriveControlOutputPort = Internal_DriveControlOutputPort Or &H80&
'    End If
'End Sub
'
'Private Sub NotImp(NotImpCom As String)
'    MsgBox "Disc operation " & NotImpCom & " not supported", vbOKOnly Or vbExclamation
'End Sub
'
'' Load the head - ignore for the moment
'Private Sub DoLoadHead()
'End Sub
'
'' Initialise our disc structures
'Private Sub InitDiscStore()
'  Dim Head As Long
'  Dim Track As Long
'  Dim Drive As Long
'  Dim blank As TrackType
'
'  'TrackType blank={0,0,NULL,0,0,0};
'
'    For Drive = 0 To 1
'        For Head = 0 To 1
'            For Track = 0 To TRACKSPERDRIVE
'                DiscStore(Drive, Head, Track) = blank
'            Next
'        Next
'    Next
'End Sub
'
'' Given a logical track number accounts for bad tracks
'Private Function SkipBadTracks(Unit As Long, trackin As Long) As Long
'    Dim offset As Long
'
'    If Not TorchTube Then 'If running under Torch Z80, ignore bad tracks
'        If Internal_BadTracks(Unit, 0) <= trackin Then
'            offset = offset + 1
'        End If
'
'        If Internal_BadTracks(Unit, 1) <= trackin Then
'            offset = offset + 1
'        End If
'    End If
'  SkipBadTracks = trackin + offset
'End Function
'
'' Returns a pointer to the data structure for a particular track.  You
'' pass the logical track number, it takes into account bad tracks and the
'' drive select and head select etc.  It always returns a valid ptr - if
'' there aren't that many tracks then it uses the last one.
'' The one exception!!!! is that if no drives are selected it returns NULL
'Private Function GetTrackPtr(LogicalTrackID As Long) As TrackType
'    Dim UnitID As Long
'
'    UnitID = -1
'
'    If (Selects(0)) Then
'        UnitID = 0
'    End If
'
'    If (Selects(1)) Then
'        UnitID = 1
'    End If
'
'    If (UnitID < 0) Then
'        GetTrackPtr = Null
'        Exit Function
'    End If
'
'    LogicalTrackID = SkipBadTracks(UnitID, LogicalTrackID)
'
'    If (LogicalTrackID >= TRACKSPERDRIVE) Then
'        LogicalTrackID = TRACKSPERDRIVE - 1
'    End If
'
'    GetTrackPtr = DiscStore(UnitID, CURRENTHEAD, LogicalTrackID)
'End Function
'
'
'' Returns a pointer to the data structure for a particular sector. Returns
'' NULL for Sector not found. Doesn't check cylinder/head ID
'Private Function GetSectorPtr(Track As TrackType, LogicalSectorID As Long, FindDeleted As Long) As SectorType
'    Dim CurrentSector As Long
'
'    If (Track.Sectors = Null) Then
'        GetSectorPtr = Null
'        Exit Function
'    End If
'
'    For CurrentSector = 0 To Track.NSectors - 1
'        If ((Track.Sectors(CurrentSector).IDField.RecordNum = LogicalSectorID) And ((Not Track.Sectors(CurrentSector).Deleted) Or (Not FindDeleted))) Then
'            GetSectorPtr = Track.Sectors(CurrentSector)
'        End If
'    Next
'
'    GetSectorPtr = Null
'End Function
'
'
'' Returns true if the drive signified by the current selects is ready
'Private Function CheckReady() As Long
'    CheckReady = 0
'
'    If (Selects(0)) Then
'        CheckReady = 1
'    End If
'
'    If (Selects(1)) Then
'        CheckReady = 1
'    End If
'End Function
'
'
'' Cause an error - pass err num
'Private Sub DoErr(ErrNum As Long)
'  SetTrigger 50, Disc8271Trigger ' Give it a bit of time
'  NextInterruptIsErr = ErrNum
'  StatusReg = &H80& ' Command is busy - come back when I have an interrupt
'  UpdateNMIStatus
'End Sub
'
'
'' Checks a few things in the sector - returns true if OK
'Private Function ValidateSector(SecToVal As SectorType, Track As Long, SecLength As Long) As Long
'    ValidateSector = 1
'
'    If (SecToVal.IDFieldx.CylinderNum <> Track) Then
'        ValidateSector = 0
'    End If
'
'    If (SecToVal.IDFieldx.PhysRecLength <> SecLength) Then
'        ValidateSector = 0
'    End If
'End Function
'
'
'Private Sub DoVarLength_ScanDataCommand()
'  DoSelects
'  NotImp ("DoVarLength_ScanDataCommand")
'End Sub
'
'
'Private Sub DoVarLength_ScanDataAndDeldCommand()
'  DoSelects
'  NotImp ("DoVarLength_ScanDataAndDeldCommand")
'End Sub
'
'
'Private Sub Do128ByteSR_WriteDataCommand()
'  DoSelects
'  NotImp ("Do128ByteSR_WriteDataCommand")
'End Sub
'
'
'Private Sub DoVarLength_WriteDataCommand()
'    Dim Drive As Long
'    Drive = -1
'
'    DoSelects
'    DoLoadHead
'
'  If (!CheckReady()) Then
'    DoErr (&H10) ' Drive not ready
'    Exit Sub
'  End If
'
'  If (Selects(0)) Then
'    Drive = 0
'  End If
'
'  If (Selects(1)) Then
'    Drive = 1
'  End If
'
'  If (!Writeable(Drive)) Then
'    DoErr (&H12) ' Drive write protected
'    Exit Sub
'  End If
'
'  Internal_CurrentTrack [Drive] = Params(0)
'  CommandStatus.CurrentTrackPtr = GetTrackPtr(Params(0))
'  If (CommandStatus.CurrentTrackPtr = Null) Then
'    DoErr (&H10)
'    Exit Sub
'  End If
'
'  CommandStatus.CurrentSectorPtr = GetSectorPtr(CommandStatus.CurrentTrackPtr, Params(1), 0)
'  If (CommandStatus.CurrentSectorPtr = Null) Then
'    DoErr (&H1E) ' Sector not found
'    Exit Sub
'  End If
'
'  CommandStatus.TrackAddr = Params(0)
'  CommandStatus.CurrentSector = Params(1)
'  CommandStatus.SectorsToGo = Params(2) & 31
'  CommandStatus.SectorLength=2^(7+((Params(2) >> 5) and 7))
'
'  If (ValidateSector(CommandStatus.CurrentSectorPtr, CommandStatus.TrackAddr, CommandStatus.SectorLength)) Then
'    CommandStatus.ByteWithinSector = 0
'    SetTrigger TIMEBETWEENBYTES, Disc8271Trigger
'    StatusReg = &H80 ' Command busy
'    UpdateNMIStatus
'    CommandStatus.ByteWithinSector = 0
'    FirstWriteInt = 1
'  Else
'    DoErr (&H1E) ' Sector not found
'  End If
'End Sub
'
'
'Private Sub WriteInterrupt()
'    Dim LastByte As Long
'
'    If (CommandStatus.SectorsToGo < 0) Then
'        StatusReg = &H18 ' Result and interrupt
'        UpdateNMIStatus
'        Exit Sub
'    End If
'
'    If (!FirstWriteInt) Then
'        CommandStatus.CurrentSectorPtr.Data(CommandStatus.ByteWithinSector) = DataReg
'        CommandStatus.ByteWithinSector = CommandStatus.ByteWithinSector + 1
'    Else
'        FirstWriteInt = 0
'    End If
'
'  ResultReg = 0
'  If (CommandStatus.ByteWithinSector >= CommandStatus.SectorLength) Then
'    CommandStatus.ByteWithinSector = 0
'    If (--CommandStatus.SectorsToGo) Then
'      CommandStatus.CurrentSector = CommandStatus.CurrentSector + 1
'      CommandStatus.CurrentSectorPtr = GetSectorPtr(CommandStatus.CurrentTrackPtr, CommandStatus.CurrentSector, 0)
'      If (CommandStatus.CurrentSectorPtr = Null) Then
'        DoErr (&H1E) ' Sector not found
'        Exit Sub
'      End If
'    Else
'      ' Last sector done, write the track back to disc
'      SaveTrackImage(Selects(0) ? 0 : 1, CURRENTHEAD, CommandStatus.TrackAddr)
'      StatusReg = &H10
'      UpdateNMIStatus
'      LastByte = 1
'      CommandStatus.SectorsToGo = -1 ' To let us bail out
'      SetTrigger 0, Disc8271Trigger ' To pick up result
'    End If
'  End If
'
'    If (!LastByte) Then
'        StatusReg = &H8C ' Command busy,
'        UpdateNMIStatus
'        SetTrigger TIMEBETWEENBYTES, Disc8271Trigger
'    End If
'End Sub
'
'
'Private Sub Do128ByteSR_WriteDeletedDataCommand()
'  DoSelects
'  NotImp ("Do128ByteSR_WriteDeletedDataCommand")
'End Sub
'
'
'Private Sub DoVarLength_WriteDeletedDataCommand()
'  DoSelects
'  NotImp ("DoVarLength_WriteDeletedDataCommand")
'End Sub
'
'
'Private Sub Do128ByteSR_ReadDataCommand()
'  DoSelects
'  NotImp ("Do128ByteSR_ReadDataCommand")
'End Sub
'
'
'Private Sub DoVarLength_ReadDataCommand()
'    Dim Drive As Long
'    Drive = -1
'
'    DoSelects
'    DoLoadHead
'
'    If (!CheckReady()) Then
'        DoErr (&H10) ' Drive not ready
'        Exit Sub
'    End If
'
'    If (Selects(0)) Then
'        Drive = 0
'    End If
'
'    If (Selects(1)) Then
'        Drive = 1
'    End If
'
'  Internal_CurrentTrack [Drive] = Params(0)
'  CommandStatus.CurrentTrackPtr = GetTrackPtr(Params(0))
'  If (CommandStatus.CurrentTrackPtr = Null) Then
'    DoErr (&H10)
'    Exit Sub
'  End If
'
'  CommandStatus.CurrentSectorPtr = GetSectorPtr(CommandStatus.CurrentTrackPtr, Params(1), 0)
'  If (CommandStatus.CurrentSectorPtr = Null) Then
'    DoErr (&H1E) ' Sector not found
'    Exit Sub
'  End If
'
'  CommandStatus.TrackAddr = Params(0)
'  CommandStatus.CurrentSector = Params(1)
'  CommandStatus.SectorsToGo = Params(2) & 31
'  CommandStatus.SectorLength=2^(7+((Params(2) >> 5) and 7&))
'
'  If (ValidateSector(CommandStatus.CurrentSectorPtr, CommandStatus.TrackAddr, CommandStatus.SectorLength)) Then
'    CommandStatus.ByteWithinSector = 0
'    SetTrigger TIMEBETWEENBYTES, Disc8271Trigger
'    StatusReg = &H80 ' Command busy
'    UpdateNMIStatus
'    CommandStatus.ByteWithinSector = 0
'  Else
'    DoErr (&H1E) ' Sector not found
'  End If
'End Sub
'
'
'Private Sub ReadInterrupt()
'    Dim DumpAfterEach As Long
'    Dim LastByte As Long
'
'    If (CommandStatus.SectorsToGo < 0) Then
'        StatusReg = &H18 ' Result and interrupt
'        UpdateNMIStatus
'        Exit Sub
'    End If
'
'    DataReg = CommandStatus.CurrentSectorPtr.Data(CommandStatus.ByteWithinSector)
'    CommandStatus.ByteWithinSector = CommandStatus.ByteWithinSector + 1
'    'cerr << "ReadInterrupt called - DataReg=&h" << hex << int(DataReg) << dec << "ByteWithinSector=" << CommandStatus.ByteWithinSector << "\n";
'
'    DumpAfterEach = 1
'    ResultReg = 0
'    If (CommandStatus.ByteWithinSector >= CommandStatus.SectorLength) Then
'        CommandStatus.ByteWithinSector = 0
'        ' I don't know if this can cause the thing to step - I presume not for the moment
'        If (--CommandStatus.SectorsToGo) Then
'            CommandStatus.CurrentSector = CommandStatus.CurrentSector + 1
'            CommandStatus.CurrentSectorPtr = GetSectorPtr(CommandStatus.CurrentTrackPtr, CommandStatus.CurrentSector, 0)
'            If (CommandStatus.CurrentSectorPtr = Null) Then
'                DoErr (&H1E) ' Sector not found
'                Exit Sub
'                '}' else cerr << "all ptr for sector " << CommandStatus.CurrentSector << "\n";
'            End If
'        Else
'            ' Last sector done
'            StatusReg = &H9C
'            UpdateNMIStatus
'            LastByte = 1
'            CommandStatus.SectorsToGo = -1 ' To let us bail out
'            SetTrigger TIMEBETWEENBYTES, Disc8271Trigger ' To pick up result
'        End If
'    End If
'
'    If (!LastByte) Then
'        StatusReg = &H8C ' Command busy
'        UpdateNMIStatus
'        SetTrigger TIMEBETWEENBYTES, Disc8271Trigger
'    End If
'End Sub
'
'
'Private Sub Do128ByteSR_ReadDataAndDeldCommand()
'  DoSelects
'  NotImp ("Do128ByteSR_ReadDataAndDeldCommand")
'End Sub
'
'
'Private Sub DoVarLength_ReadDataAndDeldCommand()
'  ' Use normal read command for now - deleted data not supported
'  DoVarLength_ReadDataCommand
'End Sub
'
'
'Private Sub DoReadIDCommand()
'    Dim Drive As Long
'    Drive = -1
'    DoSelects
'    DoLoadHead
'
'    If (Not CheckReady()) Then
'        DoErr (&H10) ' Drive not ready
'        Exit Sub
'    End If
'
'    If (Selects(0)) Then
'        Drive = 0
'    End If
'
'    If (Selects(1)) Then
'        Drive = 1
'    End If
'
'    Internal_CurrentTrack [Drive] = Params(0)
'    CommandStatus.CurrentTrackPtr = GetTrackPtr(Params(0))
'    If (CommandStatus.CurrentTrackPtr = Null) Then
'        DoErr (&H10)
'        Exit Sub
'    End If
'
'    CommandStatus.CurrentSectorPtr = GetSectorPtr(CommandStatus.CurrentTrackPtr, 0, 0)
'    If (CommandStatus.CurrentSectorPtr = Null) Then
'        DoErr (&H1E) ' Sector not found
'        Exit Sub
'    End If
'
'    CommandStatus.TrackAddr = Params(0)
'    CommandStatus.CurrentSector = 0
'    CommandStatus.SectorsToGo = Params(2)
'
'    CommandStatus.ByteWithinSector = 0
'    SetTrigger TIMEBETWEENBYTES, Disc8271Trigger
'    StatusReg = &H80& ' Command busy
'    UpdateNMIStatus
'    CommandStatus.ByteWithinSector = 0
'End Sub
'
'
'Private Sub ReadIDInterrupt()
'  Dim LastByte As Long
'
'  If (CommandStatus.SectorsToGo < 0) Then
'    StatusReg = &H18& ' Result and interrupt
'    UpdateNMIStatus
'    Exit Sub
'  End If
'
'    If (CommandStatus.ByteWithinSector = 0) Then
'        'DataReg=CommandStatus.CurrentSectorPtr->IDField.CylinderNum
'    ElseIf (CommandStatus.ByteWithinSector = 1) Then
'        'DataReg=CommandStatus.CurrentSectorPtr->IDField.HeadNum
'    ElseIf (CommandStatus.ByteWithinSector = 2) Then
'        'DataReg=CommandStatus.CurrentSectorPtr->IDField.RecordNum
'    Else
'        DataReg = 1 ' 1=256 byte sector length
'    End If
'
'  CommandStatus.ByteWithinSector = CommandStatus.ByteWithinSector + 1
'
'  ResultReg = 0
'  If (CommandStatus.ByteWithinSector >= 4) Then
'    CommandStatus.ByteWithinSector = 0
'    If (--CommandStatus.SectorsToGo) Then
'      CommandStatus.CurrentSector = CommandStatus.CurrentSector + 1
'      CommandStatus.CurrentSectorPtr = GetSectorPtr(CommandStatus.CurrentTrackPtr, CommandStatus.CurrentSector, 0)
'      If (CommandStatus.CurrentSectorPtr = Null) Then
'        DoErr (&H1E) ' Sector not found
'        Exit Sub
'      End If
'    Else
'      ' Last sector done
'      StatusReg = &H9C&
'      UpdateNMIStatus
'      LastByte = 1
'      CommandStatus.SectorsToGo = -1 ' To let us bail out
'      SetTrigger TIMEBETWEENBYTES, Disc8271Trigger ' To pick up result
'    End If
'  End If
'
'  If (Not LastByte) Then
'    StatusReg = &H8C& ' Command busy
'    UpdateNMIStatus
'    SetTrigger TIMEBETWEENBYTES, Disc8271Trigger
'  End If
'End Sub
'
'
'Private Sub Do128ByteSR_VerifyDataAndDeldCommand()
'  DoSelects
'  NotImp ("Do128ByteSR_VerifyDataAndDeldCommand")
'End Sub
'
'
'Private Sub DoVarLength_VerifyDataAndDeldCommand()
'    Dim Drive As Long
'
'    Drive = -1
'
'    DoSelects
'
'    If (!CheckReady()) Then
'      DoErr (&H10) ' Drive not ready
'      Exit Sub
'    End If
'
'    If (Selects(0)) Then
'        Drive = 0
'    End If
'
'    If (Selects(1)) Then
'        Drive = 1
'    End If
'
'    Internal_CurrentTrack [Drive] = Params(0)
'    CommandStatus.CurrentTrackPtr = GetTrackPtr(Params(0))
'    If (CommandStatus.CurrentTrackPtr = Null) Then
'      DoErr (&H10)
'      Exit Sub
'    End If
'
'    CommandStatus.CurrentSectorPtr = GetSectorPtr(CommandStatus.CurrentTrackPtr, Params(1), 0)
'    If (CommandStatus.CurrentSectorPtr = Null) Then
'      DoErr (&H1E) ' Sector not found
'      Exit Sub
'    End If
'
'    StatusReg = &H80 ' Command busy
'    UpdateNMIStatus
'    SetTrigger 100, Disc8271Trigger ' A short delay to causing an interrupt
'End Sub
'
'
'Private Sub VerifyInterrupt()
'  StatusReg = &H18 ' Result with interrupt
'  UpdateNMIStatus
'  ResultReg = 0 ' All OK
'End Sub
'
'
'Private Sub DoFormatCommand()
'    Dim Drive As Long
'
'    Drive = -1
'
'    DoSelects
'    DoLoadHead
'
'    If (!CheckReady()) Then
'        DoErr (&H10) ' Drive not ready
'        Exit Sub
'    End If
'
'    If (Selects(0)) Then
'        Drive = 0
'    End If
'
'    If (Selects(1)) Then
'        Drive = 1
'    End If
'
'    If (Not Writeable(Drive)) Then
'        DoErr (&H12) ' Drive write protected
'        Exit Sub
'    End If
'
'    Internal_CurrentTrack [Drive] = Params(0)
'    CommandStatus.CurrentTrackPtr = GetTrackPtr(Params(0))
'    If (CommandStatus.CurrentTrackPtr = Null) Then
'        DoErr (&H10)
'        Exit Sub
'    End If
'
'    CommandStatus.CurrentSectorPtr = GetSectorPtr(CommandStatus.CurrentTrackPtr, 0, 0)
'    If (CommandStatus.CurrentSectorPtr = Null) Then
'        DoErr (&H1E) ' Sector not found
'        Exit Sub
'    End If
'
'  CommandStatus.TrackAddr = Params(0)
'  CommandStatus.CurrentSector = 0
'  CommandStatus.SectorsToGo = Params(2) & 31
'  CommandStatus.SectorLength=2^(7+((Params(2) >> 5) and 7&))
'
'  If (CommandStatus.SectorsToGo = 10 And CommandStatus.SectorLength = 256) Then
'    CommandStatus.ByteWithinSector = 0
'    SetTrigger TIMEBETWEENBYTES, Disc8271Trigger
'    StatusReg = &H80 ' Command busy
'    UpdateNMIStatus
'    CommandStatus.ByteWithinSector = 0
'    FirstWriteInt = 1
'  Else
'    DoErr (&H1E) ' Sector not found
'  End If
'End Sub
'
'
'Private Sub FormatInterrupt()
'    Dim I As Long
'    Dim LastByte As Long
'
'  If (CommandStatus.SectorsToGo < 0) Then
'    StatusReg = &H18 ' Result and interrupt
'    UpdateNMIStatus
'    Exit Sub
'  End If
'
'  If (!FirstWriteInt) Then
'    ' Ignore the ID data for now - just count the bytes
'    CommandStatus.ByteWithinSector = CommandStatus.ByteWithinSector + 1
'  Else
'    FirstWriteInt = 0
'
'  ResultReg = 0
'  If (CommandStatus.ByteWithinSector >= 4) Then
'    ' Fill sector with &he5 chars
'    For I = 0 To 255
'      CommandStatus.CurrentSectorPtr.Data(I) = &HE5&
'    Next
'
'    CommandStatus.ByteWithinSector = 0
'    If (--CommandStatus.SectorsToGo) Then
'      CommandStatus.CurrentSector = CommandStatus.CurrentSector + 1
'      CommandStatus.CurrentSectorPtr = GetSectorPtr(CommandStatus.CurrentTrackPtr, CommandStatus.CurrentSector, 0)
'      If (CommandStatus.CurrentSectorPtr = Null) Then
'        DoErr (&H1E&) ' Sector not found
'        Exit Sub
'      End If
'    Else
'      ' Last sector done, write the track back to disc
'      SaveTrackImage(Selects(0) ? 0 : 1, CURRENTHEAD, CommandStatus.TrackAddr)
'      StatusReg = &H10&
'      UpdateNMIStatus
'      LastByte = 1
'      CommandStatus.SectorsToGo = -1 ' To let us bail out
'      SetTrigger 0, Disc8271Trigger ' To pick up result
'    End If
'  End If
'
'  If (!LastByte) Then
'    StatusReg = &H8C ' Command busy
'    UpdateNMIStatus
'    SetTrigger TIMEBETWEENBYTES, Disc8271Trigger
'  End If
'End Sub
'
'
'Private Sub DoSeekInt()
'  StatusReg = &H18 ' Result with interrupt
'  UpdateNMIStatus
'  ResultReg = 0 ' All OK
'End Sub
'
'
'Private Sub DoSeekCommand()
'    Dim Drive As Long
'
'    Drive = -1
'    DoSelects
'
'    DoLoadHead
'
'    If (Selects(0)) Then
'        Drive = 0
'    End If
'
'    If (Selects(1)) Then
'        Drive = 1
'    End If
'
'    If (Drive < 0) Then
'        DoErr (&H10&)
'        Exit Sub
'    End If
'
'    Internal_CurrentTrack(Drive) = Params(0)
'
'    StatusReg = &H80& ' Command busy
'    UpdateNMIStatus
'    SetTrigger 100, Disc8271Trigger ' A short delay to causing an interrupt
'End Sub
'
'
'Private Sub DoReadDriveStatusCommand()
'  Dim Track0 As Long
'  Dim WriteProt As Long
'
'  DoSelects
'
'  If (Selects(0)) Then
'    Track0 = (Internal_CurrentTrack(0) = 0)
'    WriteProt = Not Writeable(0)
'  End If
'
'  If (Selects(1)) Then
'    Track0 = 0
'    Internal_CurrentTrack(1) = 0
'    WriteProt = Not Writeable(1)
'  End If
'
'  ResultReg=&h80& or (Selects(1)?&h40:0) or (Selects(0)?&h4:0) or (Track0?2:0) or (WriteProt?8:0)
'  StatusReg = StatusReg Or &H10& ' Result
'  UpdateNMIStatus
'End Sub
'
'
'Private Sub DoSpecifyCommand()
'  ' Should set stuff up here
'End Sub
'
'
'Private Sub DoWriteSpecialCommand()
'  DoSelects
'
'  Select Case Params(0)
'    Case &H6&
'      Internal_Scan_SectorNum = Params(1)
'    Case &H14&
'      Internal_Scan_Count& = &HFF
'      Internal_Scan_Count = Internal_Scan_Count Or Params(1) * 256
'    Case &H13&
'      Internal_Scan_Count& = &HFF00
'      Internal_Scan_Count = Internal_Scan_Count Or Params(1)
'    Case &H12&
'      Internal_CurrentTrack(0) = Params(1)
'    Case &H1A&
'      Internal_CurrentTrack(1) = Params(1)
'    Case &H17&
'      Internal_ModeReg = Params(1)
'    Case &H23&
'      Internal_DriveControlOutputPort = Params(1)
'      Selects(0) = (Params(1) & &H40) <> 0
'      Selects(1) = (Params(1) & &H80) <> 0
'    Case &H22&
'      Internal_DriveControlInputPort = Params(1)
'    Case &H10&
'      Internal_BadTracks(0)(0) = Params(1)
'    Case &H11&
'      Internal_BadTracks(0)(1) = Params(1)
'    Case &H18&
'      Internal_BadTracks(1)(0) = Params(1)
'    Case &H19&
'      Internal_BadTracks(1)(1) = Params(1)
'    Case Else
'      ' cerr << "Write to bad special register\n";
'    End Select
'End Sub
'
'
'Private Sub DoReadSpecialCommand()
'  DoSelects
'
'  Select Case Params(0)
'    Case &H6&
'      ResultReg = Internal_Scan_SectorNum
'    Case &H14&
'      ResultReg = (Internal_Scan_Count \ 256) And 255
'    Case &H13&
'      ResultReg = Internal_Scan_Count And 255
'    Case &H12&
'      ResultReg = Internal_CurrentTrack(0)
'    Case &H1A&
'      ResultReg = Internal_CurrentTrack(1)
'    Case &H17&
'      ResultReg = Internal_ModeReg
'    Case &H23&
'      ResultReg = Internal_DriveControlOutputPort
'    Case &H22&
'      ResultReg = Internal_DriveControlInputPort
'    Case &H10&
'      ResultReg = Internal_BadTracks(0, 0)
'    Case &H11&
'      ResultReg = Internal_BadTracks(0, 1)
'    Case &H18&
'      ResultReg = Internal_BadTracks(1, 0)
'    Case &H19&
'      ResultReg = Internal_BadTracks(1, 1)
'    Case Else
'      ' cerr << "Read of bad special register\n";
'        Exit Sub
'    End Select
'
'  StatusReg = StatusReg Or 16 ' Result reg full
'  UpdateNMIStatus
'End Sub
'
'Private Sub DoBadCommand()
'End Sub
'
'
'' The following table is used to parse commands from the command number written into
''the command register - it can't distinguish between subcommands selected from the
''First parameter
'Private Function PrimaryCommandLookupType()
''PrimaryCommandLookup[]={
''  {&h00, &h3f, 3, DoVarLength_ScanDataCommand, NULL,  "Scan Data (Variable Length/Multi-Record)"},
''  {&h04, &h3f, 3, DoVarLength_ScanDataAndDeldCommand, NULL,  "Scan Data & deleted data (Variable Length/Multi-Record)"},
''  {&h0a, &h3f, 2, Do128ByteSR_WriteDataCommand, NULL, "Write Data (128 byte/single record)"},
''  {&h0b, &h3f, 3, DoVarLength_WriteDataCommand, WriteInterrupt, "Write Data (Variable Length/Multi-Record)"},
''  {&h0e, &h3f, 2, Do128ByteSR_WriteDeletedDataCommand, NULL, "Write Deleted Data (128 byte/single record)"},
''  {&h0f, &h3f, 3, DoVarLength_WriteDeletedDataCommand, NULL, "Write Deleted Data (Variable Length/Multi-Record)"},
''  {&h12, &h3f, 2, Do128ByteSR_ReadDataCommand, NULL, "Read Data (128 byte/single record)"},
''  {&h13, &h3f, 3, DoVarLength_ReadDataCommand, ReadInterrupt, "Read Data (Variable Length/Multi-Record)"},
''  {&h16, &h3f, 2, Do128ByteSR_ReadDataAndDeldCommand, NULL, "Read Data & deleted data (128 byte/single record)"},
''  {&h17, &h3f, 3, DoVarLength_ReadDataAndDeldCommand, ReadInterrupt, "Read Data & deleted data (Variable Length/Multi-Record)"},
''  {&h1b, &h3f, 3, DoReadIDCommand, ReadIDInterrupt, "ReadID" },
''  {&h1e, &h3f, 2, Do128ByteSR_VerifyDataAndDeldCommand, NULL, "Verify Data and Deleted Data (128 byte/single record)"},
''  {&h1f, &h3f, 3, DoVarLength_VerifyDataAndDeldCommand, VerifyInterrupt, "Verify Data and Deleted Data (Variable Length/Multi-Record)"},
''  {&h23, &h3f, 5, DoFormatCommand, FormatInterrupt, "Format"},
''  {&h29, &h3f, 1, DoSeekCommand, DoSeekInt,    "Seek"},
''  {&h2c, &h3f, 0, DoReadDriveStatusCommand, NULL, "Read drive status"},
''  {&h35, &hff, 4, DoSpecifyCommand, NULL, "Specify" },
''  {&h3a, &h3f, 2, DoWriteSpecialCommand, NULL, "Write special registers" },
''  {&h3d, &h3f, 1, DoReadSpecialCommand, NULL, "Read special registers" },
''  {0,    0,    0, DoBadCommand, NULL, "Unknown command"} ' Terminator due to 0 mask matching all
'End Function
'
'
'' returns a pointer to the data structure for the given command
'' If no matching command is given, the pointer points to an entry with a 0
'' mask, with a sensible function to call.
'Private Function CommandPtrFromNumber(CommandNumber As Long) As PrimaryCommandLookupType
'  PrimaryCommandLookupType *presptr=PrimaryCommandLookup;
'
'  for(;presptr->CommandNum<>(presptr->Mask & CommandNumber);presptr++);
'
'  CommandPtrFromNumber = presptr
'End Function
'
'
'' Address is in the range 0-7 - with the fe80 etc stripped out
'Private Function Disc8271_read(Address As Long) As Long
'    Dim Value As Long
'  Select Case Address
'    Case 0
'      'cerr << "8271 Status register read (&h" << hex << int(StatusReg) << dec << ")\n";
'      Value = StatusReg
'
'    Case 1
'      'cerr << "8271 Result register read (&h" << hex << int(ResultReg) << dec << ")\n";
'      StatusReg = StatusReg and ~18 ' Clear interrupt request  and result reg full flag
'      UpdateNMIStatus
'      Value = ResultReg
'      ResultReg = 0 ' Register goes to 0 after its read
'
'    Case 4
'      'cerr << "8271 data register read\n";
'      StatusReg = StatusReg and ~&hc ' Clear interrupt and non-dma request - not stated but DFS never looks at result reg!
'      UpdateNMIStatus
'      Value = DataReg
'
'    Case Else:
'      ' cerr << "8271: Read to unknown register address=" << Address << "\n";
'    End Select
'
'  Disc8271_read = Value
'End Function
'
'
'Private Sub CommandRegWrite(Value As Long)
'  PrimaryCommandLookupType *ptr=CommandPtrFromNumber(Value);
'  'cerr << "8271: Command register write value=&h" << hex << Value << dec << "(Name=" << ptr->Ident << ")\n";
'  ThisCommand = Value
'  NParamsInThisCommand = ptr.NParams
'  PresentParam = 0
'
'  StatusReg = StatusReg Or &H90& ' Observed on beeb for read special
'  UpdateNMIStatus
'
'  ' No parameters then call routine immediatly
'  If (NParamsInThisCommand = 0) Then
'    StatusReg = StatusReg And &H7E&
'    UpdateNMIStatus
'    ptr.ToCall();
'  End If
'End Sub
'
'
'Private Sub ParamRegWrite(Value As Long)
'    Dim tmp As Long
'
'  If (PresentParam >= NParamsInThisCommand) Then
'    ' cerr << "8271: Unwanted parameter register write value=&h" << hex << Value << dec << "\n";
'  Else
'    Params(PresentParam) = Value
'    PresentParam = PresentParam + 1
'
'    StatusReg& = &HFE ' Observed on beeb
'    UpdateNMIStatus
'
'    If (PresentParam >= NParamsInThisCommand) Then
'
'      StatusReg& = &H7E& ' Observed on beeb
'      UpdateNMIStatus
'
'      PrimaryCommandLookupType *ptr=CommandPtrFromNumber(ThisCommand)
'    ' cerr << "<Disc access>";
'    '  cerr << "8271: All parameters arrived for '" << ptr->Ident;
'        For tmp = 0 To PresentParam - 1
'            cerr << " &h" << hex << int(Params[tmp]);
'          cerr << dec << "\n";
'
'          ptr->ToCall();
'        Next
'    End If
'  End If
'End Sub
'
'
'' Address is in the range 0-7 - with the fe80 etc stripped out
'Private Sub Disc8271_write(Address As Long, Value As Long)
'  Select Case Address
'    Case 0
'      CommandRegWrite (Value)
'
'    Case 1
'      ParamRegWrite (Value)
'
'    Case 2
'      ' cerr << "8271: Reset register write, value=&h" << hex << Value << dec << "\n";
'      ' The caller should write a 1 and then >11 cycles later a 0 - but I'm just going
'      'to reset on both edges
'      Disc8271_reset
'
'    Case 4
'      ' cerr << "8271: data register write, value=&h" << hex << Value << dec << "\n";
'      StatusReg =StatusReg and  ~&hc;
'      UpdateNMIStatus
'      DataReg = Value
'
'    Case Else
'      ' cerr << "8271: Write to unknown register address=" << Address << ", value=&h" << hex << Value << dec << "\n";
'    End Select
'End Sub
'
'
'Private Sub Disc8271_poll_real()
'  ClearTrigger Disc8271Trigger
'  PrimaryCommandLookupType *comptr;
'  ' Set the interrupt flag in the status register
'  StatusReg = StatusReg Or 8
'  UpdateNMIStatus
'
'  If (NextInterruptIsErr) Then
'    ResultReg = NextInterruptIsErr
'    StatusReg = &H18 ' ResultReg full and interrupt
'    UpdateNMIStatus
'    NextInterruptIsErr = 0
'  Else
'    ' Should only happen while a command is still active
'    comptr = CommandPtrFromNumber(ThisCommand)
'    If (comptr.IntHandler <> Null) Then
'        comptr.IntHandler
'    End If
'  End If
'End Sub
'
'
'' Checks it the sectors passed in look like a valid disc catalogue. Returns:
''      1 - looks like a catalogue
''      0 - does not look like a catalogue
''     -1 - cannot tell
'Private Function CheckForCatalogue(Sec1() As Byte, Sec2() As Byte) As Long
'    Dim Valid As Long
'    Dim CatEntries As Long
'    Dim file As Long
'    Dim C As Byte
'
'    Valid = 1
'
'  ' First check the number of sectors (cannot be > &h320)
'  If (((Sec2(6) And 3) * 256) + Sec2(7) > &H320&) Then
'    Valid = 0
'    End If
'
'  ' Check the number of catalogue entries (must be multiple of 8)
'  If (Valid) Then
'    If (Sec2(5) Mod 8) Then
'      Valid = 0
'    Else
'      CatEntries = Sec2(5) / 8
'    End If
'  End If
'
'  ' Check that the catalogue file names are all printable characters.
'  for (File=0; Valid AND  File<CatEntries )
'    for (int i=0; Valid AND  i<8)
'      C = Sec1(8 + file * 8 + I)
'
'      If (I = 7) Then ' Remove lock bit
'        C& = &H7F&
'     End If
'
'        If (C < &H20 Or C > &H7F) Then
'            Valid = 0 ' not printable
'        End If
'    Next
'  Next
'
'    ' Check that all the bytes after the file names are 0
'    for (File=CatEntries; Valid AND  File<31)
'        for (int i=0; Valid AND  i<8)
'            C = Sec1(8 + file * 8 + I)
'
'            If (C <> 0) Then
'                Valid = 0
'            End If
'        Next
'    Next
'
'    ' If still valid but there are no catalogue entries then we cannot tell
'    '  if its a catalog
'    If (Valid And CatEntries = 0) Then
'        Valid = -1
'    End If
'
'    CheckForCatalogue = Valid
'End Function
'
'
'Private Sub FreeDiscImage(DriveNum As Long)
'    Dim Track As Long
'    Dim Head As Long
'    Dim Sector As Long
'    Dim SectorType As SecPtr
'
'    For Track = 0 To TRACKSPERDRIVE - 1
'        For Head = 0 To 1
'        SecPtr = DiscStore(DriveNum, Head, Track).Sectors
'        If (SecPtr <> Null) Then
'            For Sector = 0 To 9
'            If SecPtr(Sector).Data <> Null Then
'              free SecPtr(Sector).Data
'              SecPtr(Sector).Data = Null
'            End If
'          Next
'          free SecPtr
'          DiscStore(DriveNum, Head, Track).Sectors = Null
'        End If
'        Next
'    Next
'End Sub
'
'
'Private Sub LoadSimpleDiscImage(FileName As String, DriveNum As Long, HeadNum As Long, Tracks As Long)
'  Dim CurrentTrack As Long
'  Dim CurrentSector As Long
'  Dim SecPtr As SectorType
'
'  'FILE *infile=fopen(FileName,"rb");
'  If (Not infile) Then
'    MsgBox "Could not open disc file " & FileName, vbOK Or vbExclamation
'    Exit Sub
'  End If
'
'  'mainWin->SetImageName(FileName,DriveNum,0);
'
'  FileNames(DriveNum) = FileName
'  NumHeads(DriveNum) = 1
'
'  FreeDiscImage (DriveNum)
'
'  for(CurrentTrack=0;CurrentTrack<Tracks;CurrentTrack++) {
'    DiscStore[DriveNum][HeadNum][CurrentTrack].LogicalSectors=10;
'    DiscStore[DriveNum][HeadNum][CurrentTrack].NSectors=10;
'    SecPtr=DiscStore[DriveNum][HeadNum][CurrentTrack].Sectors=(SectorType*)calloc(10,sizeof(SectorType));
'    DiscStore[DriveNum][HeadNum][CurrentTrack].Gap1Size=0; ' Don't bother for the mo
'    DiscStore[DriveNum][HeadNum][CurrentTrack].Gap3Size=0;
'    DiscStore[DriveNum][HeadNum][CurrentTrack].Gap5Size=0;
'
'    for(CurrentSector=0;CurrentSector<10;CurrentSector++) {
'      SecPtr[CurrentSector].IDField.CylinderNum=CurrentTrack;
'      SecPtr[CurrentSector].IDField.RecordNum=CurrentSector;
'      SecPtr[CurrentSector].IDField.HeadNum=HeadNum;
'      SecPtr[CurrentSector].IDField.PhysRecLength=256;
'      SecPtr[CurrentSector].Deleted=0;
'      SecPtr[CurrentSector].Data=(unsigned char *)calloc(1,256);
'      fread(SecPtr[CurrentSector].Data,1,256,infile);
'    }; ' Sector
'  }; ' Track
'
'  fclose (infile)
'
'  ' Check if the sectors that would be the disc catalogue of a double sized
'     image look like a disc catalogue - give a warning if they do.
'  if (CheckForCatalogue(DiscStore[DriveNum][HeadNum](1).Sectors(0).Data,
'                        DiscStore[DriveNum][HeadNum](1).Sectors(1).Data) = 1) {
'#ifdef WIN32
'    MessageBox(GETHWND,"WARNING - Incorrect disc type selected?\n\n"
'                       "This disc file looks like a double sided\n"
'                       "disc image. Check files before copying them.\n",
'                       "BBC Emulator",MB_OK|MB_ICONWARNING);
'#Else
'    cerr << "WARNING - Incorrect disc type selected(?) in drive " << DriveNum << "\n";
'    cerr << "This disc file looks like a double sided disc image.\n";
'    cerr << "Check files before copying them.\n";
'#End If
'  }
'End Sub
'
'
'private sub LoadSimpleDSDiscImage(char *FileName, DriveNum as long,int Tracks)
'  FILE *infile=fopen(FileName,"rb");
'  int CurrentTrack,CurrentSector,HeadNum;
'  SectorType *SecPtr;
'
'  if (!infile) {
'#ifdef WIN32
'    char errstr[200];
'    sprintf(errstr, "Could not open disc file:\n  %s", FileName);
'    MessageBox(GETHWND,errstr,"BBC Emulator",MB_OK|MB_ICONERROR);
'#Else
'    cerr << "Could not open disc file " << FileName << "\n";
'#End If
'    Exit Sub
'  };
'
'  mainWin->SetImageName(FileName,DriveNum,1);
'
'  strcpy(FileNames[DriveNum], FileName);
'  NumHeads[DriveNum] = 2;
'
'  FreeDiscImage(DriveNum);
'
'  for(CurrentTrack=0;CurrentTrack<Tracks;CurrentTrack++) {
'    for(HeadNum=0;HeadNum<2;HeadNum++) {
'      DiscStore[DriveNum][HeadNum][CurrentTrack].LogicalSectors=10;
'      DiscStore[DriveNum][HeadNum][CurrentTrack].NSectors=10;
'      SecPtr=DiscStore[DriveNum][HeadNum][CurrentTrack].Sectors=(SectorType *)calloc(10,sizeof(SectorType));
'      DiscStore[DriveNum][HeadNum][CurrentTrack].Gap1Size=0; ' Don't bother for the mo
'      DiscStore[DriveNum][HeadNum][CurrentTrack].Gap3Size=0;
'      DiscStore[DriveNum][HeadNum][CurrentTrack].Gap5Size=0;
'
'      for(CurrentSector=0;CurrentSector<10;CurrentSector++) {
'        SecPtr[CurrentSector].IDField.CylinderNum=CurrentTrack;
'        SecPtr[CurrentSector].IDField.RecordNum=CurrentSector;
'        SecPtr[CurrentSector].IDField.HeadNum=HeadNum;
'        SecPtr[CurrentSector].IDField.PhysRecLength=256;
'        SecPtr[CurrentSector].Deleted=0;
'        SecPtr[CurrentSector].Data=(unsigned char *)calloc(1,256);
'        fread(SecPtr[CurrentSector].Data,1,256,infile);
'      }; ' Sector
'    }; ' Head
'  }; ' Track
'
'  fclose(infile);
'
'  ' Check if the side 2 catalogue sectors look OK - give a warning if they do not.
'  if (CheckForCatalogue(DiscStore[DriveNum](1)(0).Sectors(0).Data,
'                        DiscStore[DriveNum](1)(0).Sectors(1).Data) = 0) {
'#ifdef WIN32
'    MessageBox(GETHWND,"WARNING - Incorrect disc type selected?\n\n"
'                       "This disc file looks like a single sided\n"
'                       "disc image. Check files before copying them.\n",
'                       "BBC Emulator",MB_OK|MB_ICONWARNING);
'#Else
'    cerr << "WARNING - Incorrect disc type selected(?) in drive " << DriveNum << "\n";
'    cerr << "This disc file looks like a single sided disc image.\n";
'    cerr << "Check files before copying them.\n";
'#End If
'  }
'End Sub
'
'
'private sub SaveTrackImage(DriveNum as long, int HeadNum, int TrackNum)
'  int Success=1;
'  int CurrentSector;
'  long FileOffset;
'  long FileLength;
'  SectorType *SecPtr;
'
'  FILE *outfile=fopen(FileNames[DriveNum],"r+b");
'
'  if (!outfile) {
'#ifdef WIN32
'    char errstr[200];
'    sprintf(errstr, "Could not open disc file for write:\n  %s", FileNames[DriveNum]);
'    MessageBox(GETHWND,errstr,"BBC Emulator",MB_OK|MB_ICONERROR);
'#Else
'    cerr << "Could not open disc file for write " << FileNames[DriveNum] << "\n";
'#End If
'    Exit Sub
'  };
'
'  FileOffset=(NumHeads[DriveNum]*TrackNum+HeadNum)*2560;
'
'  ' Get the file length to check if the file needs extending
'  Success = !fseek(outfile, 0L, SEEK_END);
'  if (Success)
'  {
'    FileLength=ftell(outfile);
'    if (FileLength = -1L)
'      Success=0;
'  }
'  While (Success And FileOffset > FileLength)
'  {
'    if (fputc(0, outfile) = EOF)
'      Success=0;
'    FileLength++;
'  }
'  if (Success)
'  {
'    Success = !fseek(outfile, FileOffset, SEEK_SET);
'
'    SecPtr=DiscStore[DriveNum][HeadNum][TrackNum].Sectors;
'    for(CurrentSector=0;Success AND  CurrentSector<10;CurrentSector++) {
'      if (fwrite(SecPtr[CurrentSector].Data,1,256,outfile) <> 256)
'        Success=0;
'    }
'  }
'
'  if (fclose(outfile) <> 0)
'    Success=0;
'
'  if (!Success) {
'#ifdef WIN32
'    char errstr[200];
'    sprintf(errstr, "Failed writing to disc file:\n  %s", FileNames[DriveNum]);
'    MessageBox(GETHWND,errstr,"BBC Emulator",MB_OK|MB_ICONERROR);
'#Else
'    cerr << "Failed writing to disc file " << FileNames[DriveNum] << "\n";
'#End If
'  };
'End Sub
'
'
'Private Function IsDiscWritable(DriveNum As Long) As Long
'  IsDiscWritable = Writeable(DriveNum)
'End Function
'
'
'Private Sub DiscWriteEnable(DriveNum As Long, WriteEnable As Long)
'  int HeadNum;
'  SectorType *SecPtr;
'  unsigned char *Data;
'  int File;
'  int Catalogue, NumCatalogues;
'  int NumSecs;
'  int StartSec, LastSec;
'  int DiscOK=1;
'
'  Writeable[DriveNum]=WriteEnable;
'
'  ' If disc is being made writable then check that the disc catalogue will
'     not get corrupted if new files are added.  The files in the disc catalogue
'     must be in descending sector order otherwise the DFS ROMs write over
'     files at the start of the disc.  The sector count in the catalogue must
'     also be correct.
'  if (WriteEnable) {
'    for(HeadNum=0; DiscOK AND  HeadNum<NumHeads[DriveNum]; HeadNum++) {
'      SecPtr=DiscStore[DriveNum][HeadNum](0).Sectors;
'      if (SecPtr=NULL)
'        Exit Sub ' No disc image!
'
'      Data=SecPtr(1).Data;
'
'      ' Check for a Watford DFS 62 file catalogue
'      NumCatalogues=2;
'      Data=SecPtr(2).Data;
'      for (int i=0; i<8; ++i)
'        if (Data[i]<>(unsigned char)&haa) {
'          NumCatalogues=1;
'          break;
'        }
'
'      for (Catalogue=0; DiscOK AND  Catalogue<NumCatalogues; ++Catalogue) {
'        Data=SecPtr[Catalogue*2+1].Data;
'
'        ' First check the number of sectors
'        NumSecs=((Data[6]&3)<<8)+Data[7];
'        if (NumSecs <> &h320 AND  NumSecs <> &h190) {
'          DiscOK=0;
'        } else {
'
'          ' Now check the start sectors of each file
'          LastSec=&h320;
'          for (File=0; DiscOK AND  File<Data[5]/8; ++File) {
'            StartSec=((Data[File*8+14]&3)<<8)+Data[File*8+15];
'            if (LastSec < StartSec)
'              DiscOK=0;
'            LastSec=StartSec;
'          }
'        } ' if num sectors OK
'      } ' for catalogue
'    } ' for disc head
'
'    if (!DiscOK)
'    {
'#ifdef WIN32
'      MessageBox(GETHWND,"WARNING - Invalid Disc Catalogue\n\n"
'                       "This disc image will get corrupted if\n"
'                       "files are written to it.  Copy all the\n"
'                       "files to a new image to fix it.",
'                       "BBC Emulator",MB_OK|MB_ICONWARNING);
'#Else
'      cerr << "WARNING - Invalid Disc Catalogue in drive " << DriveNum << "\n";
'      cerr << "This disc image will get corrupted if files are written to it.\n";
'      cerr << "Copy all the files to a new image to fix it.\n";
'#End If
'    }
'
'  } ' if write enabled
'
'End Sub
'
'
'Private Sub CreateDiscImage(FileName As String, DriveNum As Long, Heads As Long, Tracks As Long)
'    Dim Success As Long
'    Dim Sector As Long
'    Dim NumSectors As Long
'    Dim I As Long
'    Dim outfile As file
'    Dim SecData(255) As Byte
'
'    Success = 1
'
'    ' First check if file already exists
'    outfile = fopen(FileName, "rb")
'    If (outfile <> Null) Then
'        fclose (outfile)
'        MsgBox "File " & FileName & " already exists. Overwrite file?", vbYesNo Or vbQuestion
'        Exit Sub
'    End If
'
'    outfile = fopen(FileName, "wb")
'    If Not outfile Then
'        MsgBox "Could not create disc file " & FileName, vbOKOnly
'        Exit Sub
'    End If
'
'    NumSectors = Tracks * 10
'
'    ' Create the first two secotrs on each side - the rest will get created when
'    ' data is written to it.
'  'for(Sector=0;Success AND  Sector<(Heads=1?2:12);Sector++) {
'    For I = 0 To 255
'      SecData(I) = 0
'    Next
'        If (Sector = 1 Or Sector = 11) Then
'          SecData(6) = NumSectors \ 256
'          SecData(7) = NumSectors And &HFF
'        End If
'
'    If (fwrite(SecData, 1, 256, outfile) <> 256) Then
'      Success = 0
'    End If
'
'  If (fclose(outfile) <> 0) Then
'    Success = 0
'  End If
'
'  If Not Success Then
'    MsgBox "Failed writing to disc file: " & FileNames, vbOKCancel
'  Else
'    ' Now load the new image into the correct drive
'    If (Heads = 1) Then
'      If ((MachineType = 3) Or (Not NativeFDC)) Then
'        Load1770DiscImage FileName, DriveNum, 0, mainWin.m_hMenu
'      Else
'        LoadSimpleDiscImage FileName, DriveNum, 0, Tracks
'        End If
'    Else
'      If ((MachineType = 3) Or (Not NativeFDC)) Then
'        Load1770DiscImage FileName, DriveNum, 1, mainWin.m_hMenu
'      Else
'        LoadSimpleDSDiscImage FileName, DriveNum, Tracks
'    End If
'  End If
'End Sub
'
'
'Private Sub LoadStartupDisc(DriveNum As Long, DiscString As String)
'    Dim DoubleSided As Byte
'    Dim Tracks As Long
'    Dim Name(1023) As Byte
'    Dim scanfres As Long
'
'    scanfres = sscanf(DiscString, "%c:%d:%s", DoubleSided, Tracks, Name)
'    If scanfres <> 3 Then
'        MsgBox "Incorrect format for BeebDiscLoad, correct format is D|S|A:tracks:filename", vbOK Or vbExclamation
'    Else
'        Select Case DoubleSided
'            Case "d", "D"
'                If ((MachineType = 3) Or (Not NativeFDC)) Then
'                 Load1770DiscImage Name, DriveNum, 1, mainWin.m_hMenu
'                Else
'                 LoadSimpleDSDiscImage Name, DriveNum, Tracks
'                End If
'
'            Case "S", "s"
'              If ((MachineType = 3) Or (Not NativeFDC)) Then
'                Load1770DiscImage Name, DriveNum, 0, mainWin.m_hMenu
'              Else
'                LoadSimpleDiscImage Name, DriveNum, 0, Tracks
'              End If
'
'            Case "A", "a"
'              If ((MachineType = 3) Or (Not NativeFDC)) Then
'                Load1770DiscImage Name, DriveNum, 2, mainWin.m_hMenu
'              Else
'                MessageBox GETHWND,"The 8271 FDC Cannot load the ADFS disc image specified in the BeebDiscLoad environment variable","BeebEm",MB_ICONERROR|MB_OK
'              End If
'
'            Case Else
'                MsgBox "BeebDiscLoad disc type incorrect, use S for single sided, D for double sided and A for ADFS", vbOK Or vbExclamation
'        End Select
'    End If
'End Sub
'
'
'Private Sub Disc8271_reset()
'    Dim onetime_initdisc As Long
'    Dim DiscString As String
'
'    ResultReg = 0
'    StatusReg = 0
'    UpdateNMIStatus
'    Internal_Scan_SectorNum = 0
'    Internal_Scan_Count = 0 ' Read as two bytes
'    Internal_ModeReg = 0
'    Internal_CurrentTrack(0) = 0 ' 0/1 for surface number
'    Internal_CurrentTrack(1) = 0 ' 0/1 for surface number
'    Internal_DriveControlOutputPort = 0
'    Internal_DriveControlInputPort = 0
'    Internal_BadTracks(0)(0) = &HFF&
'    Internal_BadTracks(0)(1) = &HFF&
'    Internal_BadTracks(1)(0) = &HFF&
'    Internal_BadTracks(1)(1) = &HFF& ' 1st subscript is surface 0/1 and second subscript is badtrack 0/1
'    ClearTrigger (Disc8271Trigger) ' No Disc8271Triggered events yet
'
'    ThisCommand = -1
'    NParamsInThisCommand = 0
'    PresentParam = 0
'    Selects(0) = 0
'    Selects(1) = 0
'
'    If Not onetime_initdisc Then
'        onetime_initdisc = onetime_initdisc + 1
'        InitDiscStore
'
'        DiscString = getenv("BeebDiscLoad")
'        If (DiscString = Null) Then
'            DiscString = getenv("BeebDiscLoad0")
'        End If
'        If (DiscString <> Null) Then
'            LoadStartupDisc 0, DiscString
'        Else
'            LoadStartupDisc 0, "S:80:discims/test.ssd"
'        End If
'
'        DiscString = getenv("BeebDiscLoad1")
'        If (DiscString <> Null) Then
'            LoadStartupDisc 1, DiscString
'        End If
'
'        If (getenv("BeebDiscWrites") <> Null) Then
'            DiscWriteEnable 0, 1
'            DiscWriteEnable 1, 1
'        End If
'    End If
'End Sub
'
'Private Sub Save8271UEF(SUEF As file)
'    Dim blank(255) As Byte
'    Dim memset(255, 0, 256) As Byte
'
'    fput16 &H46E, SUEF
'    fput32 613, SUEF
'
'    If DiscStore(0)(0)(0).Sectors = Null Then
'        ' No disc in drive 0
'        fwrite blank, 1, 256, SUEF
'    Else
'        fwrite FileNames(0), 1, 256, SUEF
'    End If
'
'    If DiscStore(1, 0, 0).Sectors = Null Then
'        ' No disc in drive 1
'        fwrite blank, 1, 256, SUEF
'    Else
'        fwrite FileNames(1), 1, 256, SUEF
'    End If
'
'    If Disc8271Trigger = CycleCountTMax Then
'        fput32 Disc8271Trigger, SUEF
'    Else
'        fput32 Disc8271Trigger - TotalCycles, SUEF
'    End If
'
'    fputc ResultReg, SUEF
'    fputc StatusReg, SUEF
'    fputc DataReg, SUEF
'    fputc Internal_Scan_SectorNum, SUEF
'    fput32 Internal_Scan_Count, SUEF
'    fputc Internal_ModeReg, SUEF
'    fputc Internal_CurrentTrack(0), SUEF
'    fputc Internal_CurrentTrack(1), SUEF
'    fputc Internal_DriveControlOutputPort, SUEF
'    fputc Internal_DriveControlInputPort, SUEF
'    fputc Internal_BadTracks(0)(0), SUEF
'    fputc Internal_BadTracks(0)(1), SUEF
'    fputc Internal_BadTracks(1)(0), SUEF
'    fputc Internal_BadTracks(1)(1), SUEF
'    fput32 ThisCommand, SUEF
'    fput32 NParamsInThisCommand, SUEF
'    fput32 PresentParam, SUEF
'    fwrite Params, 1, 16, SUEF
'    fput32 NumHeads(0), SUEF
'    fput32 NumHeads(1), SUEF
'    fput32 Selects(0), SUEF
'    fput32 Selects(1), SUEF
'    fput32 Writeable(0), SUEF
'    fput32 Writeable(1), SUEF
'    fput32 FirstWriteInt, SUEF
'    fput32 NextInterruptIsErr, SUEF
'    fput32 CommandStatus.TrackAddr, SUEF
'    fput32 CommandStatus.CurrentSector, SUEF
'    fput32 CommandStatus.SectorLength, SUEF
'    fput32 CommandStatus.SectorsToGo, SUEF
'    fput32 CommandStatus.ByteWithinSector, SUEF
'End Sub
'
'Private Sub Load8271UEF(SUEF As file)
'    Dim DiscLoaded(2) As Boolean
'    Dim sFileName As String
'    Dim ext As String
'    Dim Loaded As Long
'    Dim LoadFailed As Long
'
'    ' Clear out current images, don't want them corrupted if
'    ' saved state was in middle of writing to disc.
'    FreeDiscImage (0)
'    FreeDiscImage (1)
'    DiscLoaded(0) = False
'    DiscLoaded(1) = False
'
'    fread FileName, 1, 256, SUEF
'    If (FileName(0)) Then
'        ' Load drive 0
'        Loaded = 1
'        ext = strrchr(FileName, ".")
'        If (ext <> Null And stricmp(ext + 1, "dsd") = 0) Then
'            LoadSimpleDSDiscImage FileName, 0, 80
'        Else
'            LoadSimpleDiscImage FileName, 0, 0, 80
'        End If
'
'        If (DiscStore(0)(0)(0).Sectors = Null) Then
'            LoadFailed = 1
'        End If
'    End If
'
'    fread FileName, 1, 256, SUEF
'    If (FileName(0)) Then
'        ' Load drive 1
'        Loaded = 1
'        ext = strrchr(FileName, ".")
'        If (ext <> Null And stricmp(ext + 1, "dsd") = 0) Then
'            LoadSimpleDSDiscImage FileName, 1, 80
'        Else
'            LoadSimpleDiscImage FileName, 1, 0, 80
'        End If
'        If (DiscStore(1)(0)(0).Sectors = Null) Then
'            LoadFailed = 1
'        End If
'    End If
'
'    If (Loaded And Not LoadFailed) Then
'        Disc8271Trigger = fget32(SUEF)
'        If (Disc8271Trigger <> CycleCountTMax) Then
'            Disc8271Trigger = Disc8271Trigger + TotalCycles
'        End If
'        ResultReg = fgetc(SUEF)
'        StatusReg = fgetc(SUEF)
'        DataReg = fgetc(SUEF)
'        Internal_Scan_SectorNum = fgetc(SUEF)
'        Internal_Scan_Count = fget32(SUEF)
'        Internal_ModeReg = fgetc(SUEF)
'        Internal_CurrentTrack(0) = fgetc(SUEF)
'        Internal_CurrentTrack(1) = fgetc(SUEF)
'        Internal_DriveControlOutputPort = fgetc(SUEF)
'        Internal_DriveControlInputPort = fgetc(SUEF)
'        Internal_BadTracks(0)(0) = fgetc(SUEF)
'        Internal_BadTracks(0)(1) = fgetc(SUEF)
'        Internal_BadTracks(1)(0) = fgetc(SUEF)
'        Internal_BadTracks(1)(1) = fgetc(SUEF)
'        ThisCommand = fget32(SUEF)
'        NParamsInThisCommand = fget32(SUEF)
'        PresentParam = fget32(SUEF)
'        fread Params, 1, 16, SUEF
'        NumHeads(0) = fget32(SUEF)
'        NumHeads(1) = fget32(SUEF)
'        Selects(0) = fget32(SUEF)
'        Selects(1) = fget32(SUEF)
'        Writeable(0) = fget32(SUEF)
'        Writeable(1) = fget32(SUEF)
'        FirstWriteInt = fget32(SUEF)
'        NextInterruptIsErr = fget32(SUEF)
'        CommandStatus.TrackAddr = fget32(SUEF)
'        CommandStatus.CurrentSector = fget32(SUEF)
'        CommandStatus.SectorLength = fget32(SUEF)
'        CommandStatus.SectorsToGo = fget32(SUEF)
'        CommandStatus.ByteWithinSector = fget32(SUEF)
'
'        CommandStatus.CurrentTrackPtr = GetTrackPtr(CommandStatus.TrackAddr)
'        If (CommandStatus.CurrentTrackPtr <> Null) Then
'            CommandStatus.CurrentSectorPtr = GetSectorPtr(CommandStatus.CurrentTrackPtr, CommandStatus.CurrentSector, 0)
'        Else
'            CommandStatus.CurrentSectorPtr = Null
'        End If
'End Sub
'
'
'Private Sub disc8271_dumpstate()
''  cerr << "8271:\n";
''  cerr << "  ResultReg=" << int(ResultReg)<< "\n";
''  cerr << "  StatusReg=" << int(StatusReg)<< "\n";
''  cerr << "  DataReg=" << int(DataReg)<< "\n";
''  cerr << "  Internal_Scan_SectorNum=" << int(Internal_Scan_SectorNum)<< "\n";
''  cerr << "  Internal_Scan_Count=" << Internal_Scan_Count<< "\n";
''  cerr << "  Internal_ModeReg=" << int(Internal_ModeReg)<< "\n";
''  cerr << "  Internal_CurrentTrack=" << int(Internal_CurrentTrack(0)) << "," << int(Internal_CurrentTrack(1)) << "\n";
''  cerr << "  Internal_DriveControlOutputPort=" << int(Internal_DriveControlOutputPort)<< "\n";
''  cerr << "  Internal_DriveControlInputPort=" << int(Internal_DriveControlInputPort)<< "\n";
''  cerr << "  Internal_BadTracks=" << "(" << int(Internal_BadTracks(0)(0)) << "," << int(Internal_BadTracks(0)(1)) << ") (";
''  cerr <<                                   int(Internal_BadTracks(1)(0)) << "," << int(Internal_BadTracks(1)(1)) << ")\n";
''  cerr << "  Disc8271Trigger=" << Disc8271Trigger << "\n";
''  cerr << "  ThisCommand=" << ThisCommand<< "\n";
''  cerr << "  NParamsInThisCommand=" << NParamsInThisCommand<< "\n";
''  cerr << "  PresentParam=" << PresentParam<< "\n";
''  cerr << "  Selects=" << Selects(0) << "," << Selects(1) << "\n";
''  cerr << "  NextInterruptIsErr=" << NextInterruptIsErr<< "\n";
'End Sub
'
