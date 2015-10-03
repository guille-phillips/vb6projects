Attribute VB_Name = "FDC8271"
Option Explicit

Private mlSurface As Long
Private mlCommand As Long
Private mlParameters(4) As Long
Private mlParametersPending As Long
Private mlSpecialRegisters(&H2F&) As Long
Private mlStatus As Long
Private mlResult As Long
Private mlReadData As Long
Private mlWriteData As Long

Private mlInternalScanSectorNum As Long
Private mlInternalScanCount As Long
Private mlInternalModeReg As Long
Private mlInternalCurrentTrack(1) As Long
Private mlInternalDriveControlOutputPort As Long
Private mlInternalDriveControlInputPort As Long
Private mlInternalBadTracks(1, 1) As Long
Private mbDriveHeadLoaded As Boolean
Private mbDriveHeadUnloadPending As Boolean
Private mlDisc8271Trigger As Long
Private mlPresentParam As Long
Private mlSelects(1) As Long
Private mbWriteable(1) As Boolean
Private mlThisCommand As Long
Private mlNumberOfParamsInThisCommand As Long

Private mbCheckReady As Boolean
Private mlNextInterrruptIsError As Long
Private mlTotalCycles As Long


Private mlCSTrackAddr As Long
Private mlCSTrackNumber As Long
Private msiSectorInfo As SectorInfo
Private mlCSSectorNumber As Long
Private mlCSSectorLength As Long
Private mlCSSectorsToGo As Long
Private mlCSByteWithinSector As Long
Private mtiCSCurrentTrackInfo As TrackInfo
Private msiCSCurrentSectorInfo As SectorInfo

Private mbFirstWriteInt As Boolean

Private Enum SpecialRegisters
    spScansectornumber = &H6&
    spScanMSBofcount = &H14&
    spScanLSBofcount = &H13&
    spSurface0CurrentTrack = &H12&
    spSurface1CurrentTrack = &H1A&
    spMode = &H17&
    spDriveControlOutputPort = &H23&
    spDriveControlInputPort = &H22&
    spSurface0BadTrack1 = &H10&
    spSurface0BadTrack2 = &H11&
    spSurface1BadTrack1 = &H18&
    spSurface1BadTrack2 = &H19&
End Enum

Public Type SectorInfo
    DriveUnit As Long
    IDFieldCylinderNumber As Long
    IDFieldRecordNumber As Long
    IDFieldHeadNumber As Long
    IDFieldPhysicalRecordLength As Long
    Deleted As Boolean
    Data As Byte
    ValidSector As Boolean
End Type

Public Type TrackInfo
    LogicalSectors As Long
    NumberOfSectors As Long
    Sectors(9) As SectorInfo
    TrackNumber As Long
    ValidTrack As Boolean
End Type

Public mtiDiscStore(1, 1, 79) As TrackInfo  ' Unit Head Track

Private Const TIME_BETWEEN_BYTES As Long = 160&
Private Const SHORT_PAUSE = 330&


'Sheila Address  Read function      Write function
'FE80               Status register    Command Register
'FE81               Result register     Parameter register
'FE82                                       Reset register
'FE83               Not Used            Not used
'FE84               Read data          Write data
'                       (DMA Ack. set)  (DMA Ack. set)
    
' Command Register:
' SSCCCCCC Surface/Command
' 000000 Scan data
' 000100 Scan Data and Deleted Data
' 001001 ? DMA Access?
' 001010 Write Data (Single Byte)
' 001011 Write Data (Variable Length/Multi Record)
' 001110 Write Deleted Data (Single Byte)
' 001111 Write Deleted Data (Variable Length/Multi Record)
' 010010 Read Data (Single Byte)
' 010011 Read Data (Variable Length/Multi Record)
' 010110 Read Data and Deleted Data (Single Byte)
' 010111 Read Data and Deleted Data (Variable Length/Multi Record)
' 011011 Read ID
' 011110 Verify Data and Deleted Data (Single Byte)
' 011111 Verify Data and Deleted Data (Variable Length/Multi Record)
' 100011 Format Track
' 101001 Seek
' 101100 Read Drive Status
' 110101 Specify
' 111010 Write Special Register
' 111101 Read Special Register


' 000001 Reset Start (10 cycles)
' 000000 Reset End



' Status Register:
' BFPRID00
' B Command Busy
' F Command Register Full
' P Parameter Register Full
' R Result Register Full
' I Interrupt Request
' D Non DMA Data Request

' Result Register:
' 00DTTCC0 Deleted Data Found/Condition Type/Condition Code
' TTCC Values:
' 0000 Good Completion/Scan not met
' 0001 Scan Met Equal
' 0010 Scan Met Not Equal
' 0011 -
' 0100 Clock Error
' 0101 Late DMA
' 0110 ID CRC Error
' 0111 Data CRC Error
' 1000 Drive Not Ready
' 1001 Write Protect
' 1010 Track 0 Not Found
' 1011 Write Fault
' 1100 Sector Not Found
' 1101 -
' 1110 -
' 1111 -

' Reset Register
' 00000001 For 10 cycles
' 00000000 end of reset

Public Sub InitialiseFDC8271()
    ' Debugging.WriteString "FDC8271.InitialiseFDC8271"
    
    mlResult = 0
    mlStatus = 0
    UpdateNMIStatus
    'UpdateNMIStatusWithSource "Initialise"
    
    mlInternalScanSectorNum = 0
    mlInternalScanCount = 0
    mlInternalModeReg = 0
    mlInternalCurrentTrack(0) = 0
    mlInternalCurrentTrack(1) = 0
    mlInternalDriveControlOutputPort = 0
    mlInternalDriveControlOutputPort = 0
    mlInternalBadTracks(0, 0) = &HFF&
    mlInternalBadTracks(0, 1) = &HFF&
    mlInternalBadTracks(1, 0) = &HFF&
    mlInternalBadTracks(1, 1) = &HFF&
    
    If mbDriveHeadLoaded Then
        mbDriveHeadUnloadPending = True
    End If
    
    mlCommand = -1
    mlThisCommand = -1
    mlNumberOfParamsInThisCommand = 0
    mlPresentParam = 0
    mlSelects(0) = 0
    mlSelects(1) = 0
    
    mlDisc8271Trigger = -1 ' reset-value
End Sub

Public Sub Tick(ByVal lCycles As Long)
    ' Debugging.WriteString "FDC8271.Tick"
    
    mlTotalCycles = mlTotalCycles + lCycles \ 2
    If mlDisc8271Trigger >= 0 Then
        If mlTotalCycles > mlDisc8271Trigger Then
            mlTotalCycles = 0
            mlDisc8271Trigger = -1
            
            mlStatus = mlStatus Or &H8&
            UpdateNMIStatus
            'UpdateNMIStatusWithSource "Tick"

            If mlNextInterrruptIsError > 0 Then
                mlResult = mlNextInterrruptIsError
                mlStatus = &H18&
                UpdateNMIStatus
                'UpdateNMIStatusWithSource "TickError"
                mlNextInterrruptIsError = 0
            Else
                Select Case mlCommand
                    Case &HB& 'Write Data (Variable Length/Multi Record)
                        WriteInterrupt
                    Case &H13& 'Read Data (Variable Length/Multi Record)
                        ReadInterrupt
                    Case &H17& 'Read Data and Deleted Data (Variable Length/Multi Record)
                        ReadInterrupt
                    Case &H1B& 'Read ID
                        ReadIDInterrupt
                    Case &H1F& 'Verify Data and Deleted Data (Variable Length/Multi Record)
                        VerifyInterrupt
                    Case &H23& 'Format Track
                        FormatTrackInterrupt
                    Case &H29& 'Seek
                        SeekInterrupt
                    Case Else
                        Debug.Print "Command: " & HexNum(mlCommand, 2)
                End Select
            
            End If
            
        End If
    End If
End Sub

Private Sub WriteInterrupt()
    ' Debugging.WriteString "FDC8271.WriteInterrupt"
    
    Dim bLastByte As Boolean

    If mlCSSectorsToGo < 0 Then
        mlStatus = &H18&
        UpdateNMIStatus
        'UpdateNMIStatusWithSource "WriteInterruptNoSectorsToGo"
        Exit Sub
    End If
    
    If Not mbFirstWriteInt Then
        With msiCSCurrentSectorInfo
            StorageMedia.DiscStorage(.DriveUnit).WriteDiscByte mlCSByteWithinSector, .IDFieldRecordNumber, .IDFieldCylinderNumber, .IDFieldHeadNumber, mlWriteData
            mlCSByteWithinSector = mlCSByteWithinSector + 1
        End With
    Else
        mbFirstWriteInt = False
    End If
    
    mlResult = 0
    If mlCSByteWithinSector >= mlCSSectorLength Then
        mlCSByteWithinSector = 0
        mlCSSectorsToGo = mlCSSectorsToGo - 1
        If mlCSSectorsToGo > 0 Then
            mlCSSectorNumber = mlCSSectorNumber + 1
             
            msiCSCurrentSectorInfo = GetSectorInfo(mtiCSCurrentTrackInfo, mlCSSectorNumber, 0)
            If Not msiCSCurrentSectorInfo.ValidSector Then
                DoError &H1E&
                Exit Sub
            End If
        Else
            mlStatus = &H10&
            UpdateNMIStatus
            'UpdateNMIStatusWithSource "WriteInterruptNoSectorsLeft"
            bLastByte = True
            mlCSSectorsToGo = -1
            mlDisc8271Trigger = 0
        End If
    End If
    
    If Not bLastByte Then
        mlStatus = &H8C&
        UpdateNMIStatus
        'UpdateNMIStatusWithSource "WriteInterruptNotLastByte"
        mlDisc8271Trigger = TIME_BETWEEN_BYTES
    End If
End Sub

Private Sub ReadInterrupt()
    Dim bLastByte As Boolean
    
    ' Debugging.WriteString "FDC8271.ReadInterrupt"
    
    On Error Resume Next
    
    If mlCSSectorsToGo < 0 Then
        mlStatus = &H18&
        UpdateNMIStatus
        'UpdateNMIStatusWithSource "ReadInterrupt"
        Exit Sub
    End If
    
    With msiCSCurrentSectorInfo
        mlReadData = StorageMedia.DiscStorage(.DriveUnit).ReadDiscByte(mlCSByteWithinSector, .IDFieldRecordNumber, .IDFieldCylinderNumber, .IDFieldHeadNumber)
        'Debug.Print mlCSByteWithinSector & ":" & .IDFieldRecordNumber & ":" & .IDFieldCylinderNumber & " " & HexNum(mlReadData, 2) & " " & Chr$(mlReadData)
        mlCSByteWithinSector = mlCSByteWithinSector + 1
    End With
    
    mlResult = 0
    If mlCSByteWithinSector >= mlCSSectorLength Then
        mlCSByteWithinSector = 0
        mlCSSectorsToGo = mlCSSectorsToGo - 1
        If mlCSSectorsToGo > 0 Then
            mlCSSectorNumber = mlCSSectorNumber + 1
             
            msiCSCurrentSectorInfo = GetSectorInfo(mtiCSCurrentTrackInfo, mlCSSectorNumber, 0)
            If Not msiCSCurrentSectorInfo.ValidSector Then
                DoError &H1E&
                Exit Sub
            End If
        Else
            mlStatus = &H9C&
            UpdateNMIStatus
            'UpdateNMIStatusWithSource "ReadInterruptNoSectors"
            bLastByte = True
            mlCSSectorsToGo = -1
            mlDisc8271Trigger = TIME_BETWEEN_BYTES
        End If
    End If
    
    If Not bLastByte Then
        mlStatus = &H8C&
        UpdateNMIStatus
        'UpdateNMIStatusWithSource "ReadInterruptNotLastByte"
        mlDisc8271Trigger = TIME_BETWEEN_BYTES
    End If
End Sub

Private Sub ReadIDInterrupt()
    ' Debugging.WriteString "FDC8271.ReadIDInterrupt"
End Sub

Private Sub VerifyInterrupt()
    ' Debugging.WriteString "FDC8271.VerifyInterrupt"
    
    mlStatus = &H18&
    UpdateNMIStatus
    'UpdateNMIStatusWithSource "VerifyInterrupt"
    mlResult = 0
End Sub

Private Sub FormatTrackInterrupt()
    Dim lIndex As Long
    Dim bLastByte As Boolean
    
    ' Debugging.WriteString "FDC8271.FormatTrackInterrupt"
    
    If mlCSSectorsToGo < 0 Then
        mlStatus = &H18&
        UpdateNMIStatus
        'UpdateNMIStatusWithSource "FormatTrackInterruptNoSectors"
        Exit Sub
    End If
    
    If Not mbFirstWriteInt Then
        mlCSByteWithinSector = mlCSByteWithinSector + 1
    Else
        mbFirstWriteInt = False
    End If
    
        
    mlResult = 0
    
    If mlCSByteWithinSector >= 4 Then
        For lIndex = 0 To 255
            With msiCSCurrentSectorInfo
                StorageMedia.DiscStorage(.DriveUnit).WriteDiscByte lIndex, .IDFieldRecordNumber, .IDFieldCylinderNumber, .IDFieldHeadNumber, &H5E&
            End With
        Next
        
        mlCSByteWithinSector = 0
        
        mlCSSectorsToGo = mlCSSectorsToGo - 1
        If mlCSSectorsToGo <> 0 Then
            mlCSSectorNumber = mlCSSectorNumber + 1
             
            msiCSCurrentSectorInfo = GetSectorInfo(mtiCSCurrentTrackInfo, mlCSSectorNumber, 0)
            If Not msiCSCurrentSectorInfo.ValidSector Then
                DoError &H1E&
                Exit Sub
            End If
        Else
            mlStatus = &H10&
            UpdateNMIStatus
            'UpdateNMIStatusWithSource "FormatTrackInterruptNoSectorsToGo"
            bLastByte = True
            mlCSSectorsToGo = -1
            mlDisc8271Trigger = 0
        End If
        
    End If
    
    If Not bLastByte Then
        mlStatus = &H8C&
        UpdateNMIStatus
        'UpdateNMIStatusWithSource "FormatTrackInterruptNotLastByte"
        mlDisc8271Trigger = TIME_BETWEEN_BYTES
    End If
    
End Sub

Private Sub SeekInterrupt()
    ' Debugging.WriteString "FDC8271.SeekInterrupt"
    
    mlStatus = &H18&
    UpdateNMIStatus
    'UpdateNMIStatusWithSource "SeekInterrupt"
    mlResult = 0
End Sub

Private Sub UpdateNMIStatus()
    ' Debugging.WriteString "FDC8271.UpdateNMIStatus"
    InterruptLine.SetNMILine nmiFDC8271, Sgn(mlStatus And &H8&)
End Sub

'Private Sub UpdateNMIStatusWithSource(ByVal sSource As String)
'    ' Debugging.WriteString "FDC8271.UpdateNMIStatus"
'    InterruptLine.SetNMILineWithDescription nmiFDC8271, Sgn(mlStatus And &H8&), sSource
'End Sub

Public Sub WriteRegister(ByVal lRegister As Long, ByVal lValue As Long)
    'Debug.Print "8271W:" & lRegister & ":" & HexNum(lValue, 2)

'    If mbDriveHeadUnloadPending Then
'        mbDriveHeadUnloadPending = False
'        mlDisc8271Trigger = 10000
'    End If
    
    ' Debugging.WriteString "FDC8271.WriteRegister"
    
    Select Case lRegister
        Case 0 ' Command
            mlSurface = lValue \ 64
            mlCommand = lValue And &H3F&
            mlStatus = mlStatus Or &H90&
            UpdateNMIStatus
            'UpdateNMIStatusWithSource "WriteCommand"
            
            'Debug.Print mlSurface & ":" & HexNum(mlCommand, 2) & ":" & HexNum(lValue, 2)
            Select Case mlCommand
                Case &H0& 'Scan data
                    mlParametersPending = 3
                Case &H4& 'Scan Data and Deleted Data
                    mlParametersPending = 3
                Case &HA& 'Write Data (Single Byte)
                    mlParametersPending = 2
                Case &HB& 'Write Data (Variable Length/Multi Record)
                    mlParametersPending = 3
                Case &HE& 'Write Deleted Data (Single Byte)
                    mlParametersPending = 2
                Case &HF& 'Write Deleted Data (Variable Length/Multi Record)
                    mlParametersPending = 3
                Case &H12& 'Read Data (Single Byte)
                    mlParametersPending = 2
                Case &H13& 'Read Data (Variable Length/Multi Record)
                    mlParametersPending = 3
                Case &H16& 'Read Data and Deleted Data (Single Byte)
                    mlParametersPending = 2
                Case &H17& 'Read Data and Deleted Data (Variable Length/Multi Record)
                    mlParametersPending = 3
                Case &H1B& 'Read ID
                    mlParametersPending = 3
                Case &H1E& 'Verify Data and Deleted Data (Single Byte)
                    mlParametersPending = 3
                Case &H1F& 'Verify Data and Deleted Data (Variable Length/Multi Record)
                    mlParametersPending = 3
                Case &H23& 'Format Track
                    mlParametersPending = 5
                Case &H29& 'Seek
                    mlParametersPending = 1
                Case &H2C& 'Read Drive Status
                    mlParametersPending = 0
                    mlStatus = mlStatus And &H7E&
                    UpdateNMIStatus
                    'UpdateNMIStatusWithSource "ReadDriveStatus"
                    ReadDriveStatus
                Case &H35& 'Specify
                    mlParametersPending = 4
                Case &H3A& 'Write Special Register
                    mlParametersPending = 2
                Case &H3D& 'Read Special Register
                    mlParametersPending = 1
            End Select

        Case 1 ' Parameter
            'Debug.Print ":" & HexNum(lValue, 2)
            If mlParametersPending = 0 Then
                Exit Sub
            End If
            mlParameters(mlParametersPending - 1) = lValue
            mlParametersPending = mlParametersPending - 1
            mlStatus = mlStatus And &HFE&
            UpdateNMIStatus
            'UpdateNMIStatusWithSource "WriteParameter"
            If mlParametersPending = 0 Then
                mlStatus = mlStatus And &H7E&
                Select Case mlCommand
                    Case &H0& 'Scan data
                        ScanData
                    Case &H4& 'Scan Data and Deleted Data
                        ScanDataAndDeletedData
                    Case &HA& 'Write Data (Single Byte)
                        WriteData
                    Case &HB& 'Write Data (Variable Length/Multi Record)
                        WriteDataVariable
                    Case &HE& 'Write Deleted Data (Single Byte)
                        WriteDeletedData
                    Case &HF& 'Write Deleted Data (Variable Length/Multi Record)
                        WriteDeletedDataVariable
                    Case &H12& 'Read Data (Single Byte)
                        ReadData
                    Case &H13& 'Read Data (Variable Length/Multi Record)
                        ReadDataVariable
                    Case &H16& 'Read Data and Deleted Data (Single Byte)
                        ReadDataAndDeletedData
                    Case &H17& 'Read Data and Deleted Data (Variable Length/Multi Record)
                        ReadDataAndDeletedDataVariable
                    Case &H1B& 'Read ID
                        ReadID
                    Case &H1E& 'Verify Data and Deleted Data (Single Byte)
                        VerifyDataAndDeletedData
                    Case &H1F& 'Verify Data and Deleted Data (Variable Length/Multi Record)
                        VerifyDataAndDeletedDataVariable
                    Case &H23& 'Format Track
                        FormatTrack
                    Case &H29& 'Seek
                        SeekPosition
                    Case &H2C& 'Read Drive Status
                        ReadDriveStatus
                    Case &H35& 'Specify
                        Specify
                    Case &H3A& 'Write Special Register
                        WriteSpecialRegister
                    Case &H3D& 'Read Special Register
                        ReadSpecialRegister
                End Select
            End If
        Case 2 ' Reset
            InitialiseFDC8271
        Case 4 ' Write DMA Data
            mlStatus = mlStatus And &HF3&
            UpdateNMIStatus
            'UpdateNMIStatusWithSource "WriteDMAData"
            mlWriteData = lValue
    End Select
End Sub

Public Function ReadRegister(ByVal lRegister As Long) As Long
    ' Debugging.WriteString "FDC8271.ReadRegister"
    
    'Debug.Print "8271R:" & lRegister
    ReadRegister = 0
    
    Select Case lRegister
        Case 0 ' Status
            ReadRegister = mlStatus
        Case 1 ' Result
            mlStatus = mlStatus And &HE7&
            UpdateNMIStatus
            'UpdateNMIStatusWithSource "ReadResult"
            ReadRegister = mlResult
            mlResult = 0
        Case 4 ' Read DMA Data
            mlStatus = mlStatus And &HF3&
            UpdateNMIStatus
            'UpdateNMIStatusWithSource "ReadDMAData"
            ReadRegister = mlReadData
    End Select
End Function


Private Sub ScanData()
    ' Debugging.WriteString "FDC8271.ScanData"
End Sub

Private Sub ScanDataAndDeletedData()
    ' Debugging.WriteString "FDC8271.ScanDataAndDeletedData"
End Sub

Private Sub WriteData()
    ' Debugging.WriteString "FDC8271.WriteData"
End Sub

Private Sub WriteDataVariable()
    ' Debugging.WriteString "FDC8271.WriteDataVariable"

    Dim lDrive As Long
       
    lDrive = -1
    DoSelects
    
    If Not CheckReady Then
        DoError &H10&
        Exit Sub
    End If
    
    If mlSelects(0) = 1 Then
        lDrive = 0
    End If
    
    If mlSelects(1) = 1 Then
        lDrive = 1
    End If
    
'    If Not Writeable(lDrive) Then
'        DoError &H12&
'        Exit Sub
'    End If

    mlInternalCurrentTrack(lDrive) = mlParameters(2)
    mtiCSCurrentTrackInfo = GetTrackInfo(mlParameters(2))
    If Not mtiCSCurrentTrackInfo.ValidTrack Then
        DoError &H10&
        Exit Sub
    End If
    
    msiCSCurrentSectorInfo = GetSectorInfo(mtiCSCurrentTrackInfo, mlParameters(1), 0)
    If Not msiCSCurrentSectorInfo.ValidSector Then
        DoError &H1E&
        Exit Sub
    End If
    
    mlCSTrackAddr = mlParameters(2)
    mlCSSectorNumber = mlParameters(1)
    
    mlCSSectorsToGo = mlParameters(0) And 31&
    mlCSSectorLength = 2& ^ (((mlParameters(0) \ 32&) And 7&) + 7&)
        
    If ValidateSector(mlCSTrackAddr, msiCSCurrentSectorInfo, mlCSSectorLength) Then
        mlCSByteWithinSector = 0
        mlDisc8271Trigger = TIME_BETWEEN_BYTES
        mlStatus = &H80&
        UpdateNMIStatus
        'UpdateNMIStatusWithSource "WriteDataVariable"
        mbFirstWriteInt = True
    Else
        DoError &H1E& ' sector not found
    End If
End Sub

Private Sub WriteDeletedData()
    ' Debugging.WriteString "FDC8271.WriteDeletedData"
End Sub

Private Sub WriteDeletedDataVariable()
    ' Debugging.WriteString "FDC8271.WriteDeletedDataVariable"
End Sub

Private Sub ReadData()
    ' Debugging.WriteString "FDC8271.ReadData"
End Sub

Private Sub ReadDataVariable()
    Dim lDrive As Long
    
    ' Debugging.WriteString "FDC8271.ReadDataVariable"
    
    lDrive = -1
    DoSelects
    'DoLoadHead
    
    If Not CheckReady Then
        DoError &H10&
        Exit Sub
    End If
    
    If mlSelects(0) = 1 Then
        lDrive = 0
    End If
    
    If mlSelects(1) = 1 Then
        lDrive = 1
    End If
    
    mlInternalCurrentTrack(lDrive) = mlParameters(2)
    mtiCSCurrentTrackInfo = GetTrackInfo(mlParameters(2))
    If Not mtiCSCurrentTrackInfo.ValidTrack Then
        DoError &H10&
        Exit Sub
    End If
    
    msiCSCurrentSectorInfo = GetSectorInfo(mtiCSCurrentTrackInfo, mlParameters(1), 0)
    If Not msiCSCurrentSectorInfo.ValidSector Then
        DoError &H1E&
        Exit Sub
    End If
    
    mlCSTrackAddr = mlParameters(2)
    mlCSSectorNumber = mlParameters(1)
    
    mlCSSectorsToGo = mlParameters(0) And 31&
    mlCSSectorLength = 2& ^ (((mlParameters(0) \ 32&) And 7&) + 7&)
        
    If ValidateSector(mlCSTrackAddr, msiCSCurrentSectorInfo, mlCSSectorLength) Then
        mlCSByteWithinSector = 0
        mlDisc8271Trigger = TIME_BETWEEN_BYTES
        mlStatus = &H80&
        UpdateNMIStatus
        'UpdateNMIStatusWithSource "ReadDataVariable"
    Else
        DoError &H1E& ' sector not found
    End If
End Sub

Private Function ValidateSector(ByVal lTrack As Long, siSectorInfo As SectorInfo, ByVal lSectorLength As Long) As Boolean
    ' Debugging.WriteString "FDC8271.ValidateSector"
    
    If siSectorInfo.IDFieldCylinderNumber <> lTrack Then
        Exit Function
    End If
    
    If siSectorInfo.IDFieldPhysicalRecordLength <> lSectorLength Then
        Exit Function
    End If
    ValidateSector = True
End Function

Private Function GetSectorInfo(tiTrackInfo As TrackInfo, ByVal lLogicalSector As Long, ByVal bFindDeleted As Boolean) As SectorInfo
    Dim lCurrentSector As Long
    
    ' Debugging.WriteString "FDC8271.GetSectorInfo"
    
    For lCurrentSector = 0 To tiTrackInfo.NumberOfSectors - 1
        If tiTrackInfo.Sectors(lCurrentSector).IDFieldRecordNumber = lLogicalSector And (Not tiTrackInfo.Sectors(lCurrentSector).Deleted Or Not bFindDeleted) Then
            GetSectorInfo = tiTrackInfo.Sectors(lCurrentSector)
            GetSectorInfo.ValidSector = True
            Exit Function
        End If
    Next
End Function

Private Function GetTrackInfo(ByVal lLogicalTrackNumber) As TrackInfo
    Dim lUnit As Long
    
    ' Debugging.WriteString "FDC8271.GetTrackInfo"
    
    lUnit = -1
    
    If mlSelects(0) = 1 Then
        lUnit = 0
    End If
    
    If mlSelects(1) = 1 Then
        lUnit = 1
    End If
    
    If lUnit = -1 Then
        GetTrackInfo.ValidTrack = False
        Exit Function
    End If
    
    lLogicalTrackNumber = SkipBadTracks(lLogicalTrackNumber, lUnit)

    If lLogicalTrackNumber >= 80 Then
        lLogicalTrackNumber = 79
    End If
    GetTrackInfo = mtiDiscStore(lUnit, GetCurrentHead, lLogicalTrackNumber)
    GetTrackInfo.TrackNumber = lLogicalTrackNumber
    GetTrackInfo.ValidTrack = True
End Function

Private Function GetCurrentHead() As Long
    ' Debugging.WriteString "FDC8271.GetCurrentHead"
    
    GetCurrentHead = (mlInternalDriveControlOutputPort \ 32) And &H1&
End Function

Private Function SkipBadTracks(ByVal lLogicalTrackNumber, ByVal lUnit As Long) As Long
    ' Debugging.WriteString "FDC8271.SkipBadTracks"
    
    SkipBadTracks = lLogicalTrackNumber
    
    If mlInternalBadTracks(lUnit, 0) <= lLogicalTrackNumber Then
        SkipBadTracks = SkipBadTracks + 1
    End If
    If mlInternalBadTracks(lUnit, 1) <= lLogicalTrackNumber Then
        SkipBadTracks = SkipBadTracks + 1
    End If
End Function

Private Function CheckReady() As Boolean
    ' Debugging.WriteString "FDC8271.CheckReady"
    
    If mlSelects(0) = 1 Or mlSelects(1) = 1 Then
        CheckReady = True
    End If
End Function

Private Sub ReadDataAndDeletedData()
    ' Debugging.WriteString "FDC8271.ReadDataAndDeletedData"
    DoSelects
End Sub

Private Sub ReadDataAndDeletedDataVariable()
    ' Debugging.WriteString "FDC8271.ReadDataAndDeletedDataVariable"
    DoSelects
End Sub

Private Sub ReadID()
    ' Debugging.WriteString "FDC8271.ReadID"
    DoSelects
End Sub

Private Sub VerifyDataAndDeletedData()
    ' Debugging.WriteString "FDC8271.VerifyDataAndDeletedData"
        
    DoSelects
End Sub

Private Sub VerifyDataAndDeletedDataVariable()
    ' Debugging.WriteString "FDC8271.VerifyDataAndDeletedDataVariable"
    
    Dim lDrive As Long
    
    lDrive = -1
    DoSelects
    
    If Not CheckReady Then
        DoError &H10&
        Exit Sub
    End If
    
    If mlSelects(0) = 1 Then
        lDrive = 0
    End If
    
    If mlSelects(1) = 1 Then
        lDrive = 1
    End If
    
    mlInternalCurrentTrack(lDrive) = mlParameters(2)
    mtiCSCurrentTrackInfo = GetTrackInfo(mlParameters(2))
    If Not mtiCSCurrentTrackInfo.ValidTrack Then
        DoError &H10&
        Exit Sub
    End If
    
    msiCSCurrentSectorInfo = GetSectorInfo(mtiCSCurrentTrackInfo, mlParameters(1), 0)
    If Not msiCSCurrentSectorInfo.ValidSector Then
        DoError &H1E&
        Exit Sub
    End If
    
    mlStatus = &H80&
    UpdateNMIStatus
    'UpdateNMIStatusWithSource "VeryiyDataAndDeletedDataVariable"
    mlDisc8271Trigger = SHORT_PAUSE
End Sub

Private Sub FormatTrack()
    Dim lDrive As Long
    
    ' Debugging.WriteString "FDC8271.FormatTrack"
    
    lDrive = -1
    
    DoSelects
    
    ' DoLoadHead - not needed
    
    If mlSelects(0) = 1 Then
        lDrive = 0
    End If
    
    If mlSelects(1) = 1 Then
        lDrive = 1
    End If
    
    If lDrive = -1 Then
        DoError &H10&
        Exit Sub
    End If
    
    mlInternalCurrentTrack(lDrive) = mlParameters(4)
    mtiCSCurrentTrackInfo = GetTrackInfo(mlParameters(4))
    If Not mtiCSCurrentTrackInfo.ValidTrack Then
        DoError &H10&
        Exit Sub
    End If
    
    msiCSCurrentSectorInfo = GetSectorInfo(mtiCSCurrentTrackInfo, 0, 0)
    If Not msiCSCurrentSectorInfo.ValidSector Then
        DoError &H1E&
        Exit Sub
    End If
    
    mlCSTrackAddr = mlParameters(4)
    mlCSSectorNumber = 0
    
    mlCSSectorsToGo = mlParameters(2) And 31&
    mlCSSectorLength = 2& ^ (((mlParameters(2) \ 32&) And 7&) + 7&)
    
    If mlCSSectorsToGo = 10 And mlCSSectorLength = 256 Then
        mlCSByteWithinSector = 0
        mlDisc8271Trigger = TIME_BETWEEN_BYTES
        mlStatus = &H80&
        UpdateNMIStatus
        'UpdateNMIStatusWithSource "FormatTrack"
        mbFirstWriteInt = True
    Else
        DoError &H1E& ' sector not found
    End If
End Sub

Private Sub SeekPosition()
    Dim lDrive As Long
    
    ' Debugging.WriteString "FDC8271.SeekPosition"
    
    lDrive = -1
    
    DoSelects
    
    ' DoLoadHead - not needed
    
    If mlSelects(0) = 1 Then
        lDrive = 0
    End If
    
    If mlSelects(1) = 1 Then
        lDrive = 1
    End If
    
    If lDrive = -1 Then
        DoError &H10&
        Exit Sub
    End If
    
    mlInternalCurrentTrack(lDrive) = mlParameters(0)
    
    mlStatus = &H80&
    UpdateNMIStatus
    'UpdateNMIStatusWithSource "SeekPosition"
    mlDisc8271Trigger = SHORT_PAUSE
End Sub

Private Sub ReadDriveStatus()
    Dim lTrack0 As Long
    Dim lWriteProtect As Long
    
    ' Debugging.WriteString "FDC8271.ReadDriveStatus"
    
    DoSelects
    
    If mlSelects(0) = 1 Then
        lTrack0 = Abs(mlInternalCurrentTrack(0) = 0)
        lWriteProtect = Abs(Not mbWriteable(0))
    End If
    
    If mlSelects(1) = 1 Then
        lTrack0 = Abs(mlInternalCurrentTrack(1) = 0)
        lWriteProtect = Abs(Not mbWriteable(1))
    End If
    
    mlResult = &H80& Or mlSelects(1) * &H40& Or mlSelects(0) * &H4& Or lTrack0 * &H2& Or lWriteProtect * &H8&
    mlStatus = mlStatus Or &H10&
    UpdateNMIStatus
    'UpdateNMIStatusWithSource "ReadDriveStatus"
End Sub

Private Sub Specify()
    ' Debugging.WriteString "FDC8271.Specify"
    
'    Select Case mlParameters(3)
'        Case &HD& ' Initialisation
'        Case &H10& ' Load Bad Tracks Surface 0
'        Case &H18& ' Load Bad Tracks Surface 1
'    End Select
End Sub

Private Sub WriteSpecialRegister()
    Dim lSpecialRegister As Long
    Dim lValue As Long
    
    ' Debugging.WriteString "FDC8271.WriteSpecialRegister"
    
    DoSelects
    'mlSpecialRegisters(mlParameters(1)) = mlParameters(0)
    lSpecialRegister = mlParameters(1)
    Select Case lSpecialRegister
        Case &H6&
            mlInternalScanSectorNum = mlParameters(0)
        Case &H14&
            mlInternalScanCount = mlInternalScanCount And &HFF&
            mlInternalScanCount = mlInternalScanCount Or mlParameters(0) * 256
        Case &H13&
            mlInternalScanCount = mlInternalScanCount And &HFF00&
            mlInternalScanCount = mlInternalScanCount Or mlParameters(0)
        Case &H12&
            mlInternalCurrentTrack(0) = mlParameters(0)
        Case &H1A&
            mlInternalCurrentTrack(1) = mlParameters(0)
        Case &H17&
            mlInternalModeReg = mlParameters(0) 'done
        Case &H23&
            mlInternalDriveControlOutputPort = mlParameters(0) 'done
            mlSelects(0) = Sgn(mlParameters(0) And &H40&)
            mlSelects(1) = Sgn(mlParameters(0) And &H80&)
        Case &H22&
            mlInternalDriveControlInputPort = mlParameters(0)
        Case &H10&
            mlInternalBadTracks(0, 0) = mlParameters(0)
        Case &H11&
            mlInternalBadTracks(0, 1) = mlParameters(0)
        Case &H18&
            mlInternalBadTracks(1, 0) = mlParameters(0)
        Case &H19&
            mlInternalBadTracks(1, 1) = mlParameters(0)
    End Select
End Sub

Private Sub DoSelects()
    ' Debugging.WriteString "FDC8271.DoSelects"
    
    mlSelects(0) = Sgn(mlSurface And &H1&)
    mlSelects(1) = Sgn(mlSurface And &H2&)
    mlInternalDriveControlOutputPort = mlInternalDriveControlOutputPort And &H3F&
    If mlSelects(0) = 1 Then
        mlInternalDriveControlOutputPort = mlInternalDriveControlOutputPort Or &H40&
    End If
    If mlSelects(1) = 1 Then
        mlInternalDriveControlOutputPort = mlInternalDriveControlOutputPort Or &H80&
    End If
End Sub

Private Sub DoError(ByVal lErrorNumber As Long)
    ' Debugging.WriteString "FDC8271.DoError"
    
    mlDisc8271Trigger = 50&
    mlNextInterrruptIsError = lErrorNumber
    mlStatus = mlStatus Or &H80&
    
    UpdateNMIStatus
    'UpdateNMIStatusWithSource "DoError"
End Sub

Private Sub ReadSpecialRegister()
    Dim lSpecialRegister As Long
    
    ' Debugging.WriteString "FDC8271.ReadSpecialRegister"
    
    lSpecialRegister = mlParameters(1)

    DoSelects
    Select Case lSpecialRegister
        Case &H6&
            mlResult = mlInternalScanSectorNum
        Case &H14&
            mlResult = mlInternalScanCount \ 256
        Case &H13&
            mlResult = mlInternalScanCount And &HFF&
        Case &H12&
            mlResult = mlInternalCurrentTrack(0)
        Case &H1A&
            mlResult = mlInternalCurrentTrack(1)
        Case &H17&
            mlResult = mlInternalModeReg
        Case &H23&
            mlResult = mlInternalDriveControlOutputPort
        Case &H22&
            mlResult = mlInternalDriveControlInputPort
        Case &H10&
            mlResult = mlInternalBadTracks(0, 0)
        Case &H11&
            mlResult = mlInternalBadTracks(0, 1)
        Case &H18&
            mlResult = mlInternalBadTracks(1, 0)
        Case &H19&
            mlResult = mlInternalBadTracks(1, 1)
    End Select
    
    mlStatus = mlStatus Or &H10&
    UpdateNMIStatus
    'UpdateNMIStatusWithSource "ReadSpecialRegister"
End Sub


Public Sub InitialiseDisc()
    Dim lTrack As Long
    Dim lSector As Long
    Dim lDriveNumber As Long
    Dim lHead As Long
    
    ' Debugging.WriteString "FDC8271.InitialiseDisc"
    
    For lDriveNumber = 0 To 1
        For lHead = 0 To 1
            For lTrack = 0 To 79
                With FDC8271.mtiDiscStore(lDriveNumber, lHead, lTrack)
                    .LogicalSectors = 10
                    .NumberOfSectors = 10
                End With
                For lSector = 0 To 9
                    With FDC8271.mtiDiscStore(lDriveNumber, lHead, lTrack).Sectors(lSector)
                        .DriveUnit = lDriveNumber
                        .IDFieldHeadNumber = lHead
                        .IDFieldCylinderNumber = lTrack
                        .IDFieldRecordNumber = lSector
                        .IDFieldPhysicalRecordLength = 256
                        .Deleted = False
                    End With
                Next
            Next
        Next
    Next
    mbWriteable(0) = True
    mbWriteable(1) = True
End Sub
