VERSION 5.00
Begin VB.Form Storage 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "Files"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstFiles 
      Height          =   2010
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2055
   End
End
Attribute VB_Name = "Storage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Type fiFile
    FileName As String
    Name As String
    Load As Long
    Execution As Long
    Length As Long
    Locked As Boolean
    Side As Long
    HeaderCRC As Long
    DataCRC As Long
    Data() As Byte
    BlockData() As Byte
    BlockLengths() As Long
    BlockDataStarts() As Long
    DriveNumber As Long
    NextFile As String
End Type

Private mfiCatalogue() As fiFile
Private mlCatalogueCount As Long
Private msCatalogueName(1) As String
Private mlBootOption(1) As Long

Public Enum FileTypes
    ftUnknown
    ftSnapshot
    ftCassette
    ftDisc
    ftDiscDSD
    ftArchive
    ftMemory
End Enum

Private Type CassettePosition
    FileIndex As Long
    BlockIndex As Long
    DataIndex As Long
    CarrierTimer As Long
End Type

Private cpCassetteIndex As CassettePosition

Private myDiscImage(255, 9, 79, 1) As Byte ' DATUM/SECTOR/TRACK/HEAD*2+DRIVE
Private mlBaseAddress As Long

Private Sub AddFile(fiAddFile As fiFile)
    ' Debugging.WriteString "Storage.AddFile"
    
    ReDim Preserve mfiCatalogue(mlCatalogueCount)
    mfiCatalogue(mlCatalogueCount) = fiAddFile
    mlCatalogueCount = mlCatalogueCount + 1
End Sub

Public Sub ClearDiskImage()
    Erase myDiscImage
    myDiscImage(6, 1, 0, 0) = 2560 \ &H100&
    myDiscImage(7, 1, 0, 0) = 2560 And &HFF&
End Sub

Public Sub FlipSides()
    Dim yImage(204799) As Byte
    
    CopyMemory yImage(0), myDiscImage(0, 0, 0, 0), 204800
    CopyMemory myDiscImage(0, 0, 0, 0), myDiscImage(0, 0, 0, 1), 204800
    CopyMemory myDiscImage(0, 0, 0, 1), yImage(0), 204800
End Sub

Public Function LoadFile(ByVal sPath As String, ByVal bLoadAsCassette As Boolean) As FileTypes
    Dim yImage() As Byte
    Dim lFileType As FileTypes
    Dim oSaveMemory As SaveMemory
    
    ' Debugging.WriteString "Storage.AddFile"
    
    lFileType = Me.DetermineFileType(sPath)
    Select Case lFileType
        Case ftSnapshot
            Controller.msSnapshotFilePath = sPath
            If Processor6502.StopReason = srDebugBreak Then
                Monitor.Hide
                Console.DebugOn = False
                Processor6502.StopReason = srLoadSnapshot
                Controller.ProcessorStopped
            Else
                Processor6502.StopReason = srLoadSnapshot
            End If
        Case ftCassette
            LoadCassette sPath
            ACIA6850.InitialiseCassette
        Case ftDisc, ftDiscDSD
            LoadDisc sPath, lFileType = ftDiscDSD
            FDC8271.InitialiseDisc
        Case ftArchive
            Me.LoadArchive sPath
            If bLoadAsCassette Then
            Else
                yImage = BuildDiscImage
                CopyMemory myDiscImage(0, 0, 0, 0), yImage(0, 0), 256& * 10& * 80& * 2&
                FDC8271.InitialiseDisc
            End If
        Case ftMemory
            Set oSaveMemory = New SaveMemory
            oSaveMemory.Mode = mtLoad
            oSaveMemory.Show vbModal
            If oSaveMemory.mlStartAddress <> -1 Then
                Memory.LoadMemory sPath, oSaveMemory.mlStartAddress
            End If
        Case ftUnknown
            MsgBox "File format not recognised", vbOKOnly Or vbExclamation
    End Select
    
    LoadFile = lFileType
End Function

Public Function DetermineFileType(ByVal sPath As String) As Long
    Dim lDot As Long
    Dim sExt As String
    Dim bDisc As Boolean
    Dim oFSO As New FileSystemObject
    Dim oFile As file
    
    ' Debugging.WriteString "Storage.DetermineFileType"
    
    DetermineFileType = ftUnknown
    If UEFHandler.LoadCompressedUEFFile(sPath) Then
        ' Tape or Snapshot
        If UEFHandler.FindBlock(&H462&) Then  ' memory block present
            DetermineFileType = ftSnapshot
            Exit Function
        End If
        
        UEFHandler.ResetUEF
        If UEFHandler.FindBlock(&HFF00&) Then  ' memory block present
            DetermineFileType = ftSnapshot
            Exit Function
        End If
        
        UEFHandler.ResetUEF
        If UEFHandler.FindBlock(&H100&) Then ' data block present
            DetermineFileType = ftCassette
            Exit Function
        End If
    Else
        ' Disc or Archive or something else
        lDot = InStrRev(sPath, ".")
        If lDot > 0 Then
            sExt = LCase$(Mid$(sPath, lDot + 1))
            Select Case sExt
                Case "img", "ssd"
                    DetermineFileType = ftDisc
                    bDisc = True
                Case "dsd"
                    DetermineFileType = ftDiscDSD
                    bDisc = True
                Case "mem"
                    DetermineFileType = ftMemory
                    bDisc = True
            End Select
        End If
        
        If bDisc Then
            Exit Function
        Else
            For Each oFile In oFSO.GetFile(sPath).ParentFolder.Files
                lDot = InStrRev(oFile.Name, ".")
                If lDot > 0 Then
                    sExt = LCase$(Mid$(oFile.Name, lDot + 1))
                    If sExt = "inf" Then
                        DetermineFileType = ftArchive
                        Exit Function
                    End If
                End If
            Next
        End If
    End If
End Function

Public Sub LoadArchive(ByVal sPath As String)
    Dim oFSO As New FileSystemObject
    Dim oFile As file
    Dim oSearchFile As file
    Dim lPoint As Long
    Dim sExt As String
    Dim fiFileInfo As fiFile
    Dim sContents As String
    Dim sPreviousContents As String
    Dim vParts As Variant
    Dim vChecksum As Variant
    Dim lPos As Long
    
    ' Debugging.WriteString "Storage.LoadArchive"
    
    Set oFile = oFSO.GetFile(sPath)
    
    For Each oSearchFile In oFile.ParentFolder.Files
        lPoint = InStrRev(oSearchFile.Name, ".")
        If lPoint > 0 Then
            sExt = LCase$(Mid$(oSearchFile.Name, lPoint + 1))
                    
            If sExt = "inf" Then
                sContents = oSearchFile.OpenAsTextStream.ReadAll
                While sContents <> sPreviousContents
                    sPreviousContents = sContents
                    sContents = Replace$(sContents, "  ", " ")
                Wend
                vParts = Split(sContents, " ")
                With fiFileInfo
                    .Locked = False
                    .NextFile = ""
                    .DataCRC = 0
                    
                    .FileName = Left$(oSearchFile.Name, lPoint - 1)
                    .Name = vParts(0)
                    .Load = FromHex(Right$(vParts(1), 5))
                    .Execution = FromHex(Right$(vParts(2), 5))
                    If UBound(vParts) = 3 Then
                        If LCase$(vParts(3)) = "locked" Then
                            .Locked = True
                            vChecksum = Split(vParts(4), "=")
                            .DataCRC = FromHex(vChecksum(1))
                        ElseIf LCase$(Left$(vParts(3), 4)) = "next" Then
                            vChecksum = Split(vParts(3), "=")
                            If UBound(vChecksum) = 1 Then
                                .NextFile = StripWhiteSpace(vChecksum(1))
                            End If
                        Else
                            vChecksum = Split(vParts(3), "=")
                            If UBound(vChecksum) = 1 Then
                                .DataCRC = FromHex(vChecksum(1))
                            End If
                        End If
                    End If
                End With
                AddFile fiFileInfo
            End If
        End If
    Next
    
    For Each oSearchFile In oFile.ParentFolder.Files
        For lPos = 0 To mlCatalogueCount - 1
            If mfiCatalogue(lPos).Length = 0 Then
                If LCase$(mfiCatalogue(lPos).FileName) = LCase$(oSearchFile.Name) Then
                    mfiCatalogue(lPos).Length = oSearchFile.SIZE
                    ReDim mfiCatalogue(lPos).Data(mfiCatalogue(lPos).Length - 1)
                    Open oSearchFile For Binary As #1
                    Get #1, , mfiCatalogue(lPos).Data
                    Close #1
                    Exit For
                End If
            End If
        Next
    Next
    
    Dim bSorted As Boolean
    Dim lIndex1 As Long
    Dim lIndex2 As Long
    
    While Not bSorted
        bSorted = True
        For lIndex1 = 0 To mlCatalogueCount - 1
            If mfiCatalogue(lIndex1).NextFile <> "" Then
                For lIndex2 = lIndex1 + 2 To mlCatalogueCount - 1
                    If LCase$(mfiCatalogue(lIndex1).NextFile) = LCase$(mfiCatalogue(lIndex2).Name) Then
                        fiFileInfo = mfiCatalogue(lIndex1 + 1)
                        mfiCatalogue(lIndex1 + 1) = mfiCatalogue(lIndex2)
                        mfiCatalogue(lIndex2) = fiFileInfo
                        bSorted = False
                        Exit For
                    End If
                Next
            End If
        Next
    Wend
End Sub

Private Function StripWhiteSpace(ByVal sString As String) As String
    Dim lPos As Long
    
    For lPos = 1 To Len(sString)
        If Asc(Mid$(sString, lPos, 1)) > 32 Then
            StripWhiteSpace = StripWhiteSpace & Mid$(sString, lPos, 1)
        End If
    Next
End Function

Public Sub LoadCassette(ByVal sPath As String)
    Dim sFileName As String
    Dim lChecksum As Long
    Dim fiFileInfo As fiFile
    Dim lPosition As Long
    Dim lDataPosition As Long
    Dim sPreviousFileName As String
    Dim lBlockLength As Long
    
    Dim lCRC As Long
    
    ' Debugging.WriteString "Storage.LoadCassette"
    
    If Not UEFHandler.LoadCompressedUEFFile(sPath) Then
        MsgBox "Cassette format not recognised", vbExclamation
        Exit Sub
    End If
    
    While UEFHandler.LoadNextBlock
'        Debug.Print HexNum(UEFHandler.BlockIdentifier, 4)
        Select Case UEFHandler.BlockIdentifier
            Case &H100& ' data block
                If UEFHandler.BlockData(0) = &H2A& Then ' sync byte
                    lPosition = 1
                    sFileName = ""
                    While Not UEFHandler.BlockData(lPosition) = 0
                        sFileName = sFileName & Chr$(UEFHandler.BlockData(lPosition))
                        lPosition = lPosition + 1
                    Wend
                    lPosition = lPosition + 1
                    
                    If sFileName <> sPreviousFileName Then
                        If sPreviousFileName <> "" Then
                            AddFile fiFileInfo
                        End If
                        fiFileInfo.FileName = sFileName
                        CopyMemory fiFileInfo.Load, UEFHandler.BlockData(lPosition), 4&
                        CopyMemory fiFileInfo.Execution, UEFHandler.BlockData(lPosition + 4), 4&
                        CopyMemory lBlockLength, UEFHandler.BlockData(lPosition + 10), 2&
                        
                        
                        fiFileInfo.Length = lBlockLength
                        If lBlockLength > 0 Then
                            ReDim fiFileInfo.Data(lBlockLength - 1)
                            CopyMemory fiFileInfo.Data(0), UEFHandler.BlockData(lPosition + 19), lBlockLength
                        Else
                            Erase fiFileInfo.Data
                        End If
                        lDataPosition = lBlockLength
                        ReDim fiFileInfo.BlockLengths(0)
                        fiFileInfo.BlockLengths(0) = lBlockLength
                        sPreviousFileName = sFileName
                    Else
                        ReDim Preserve fiFileInfo.Data(lDataPosition + lBlockLength - 1)
                        CopyMemory fiFileInfo.Data(lDataPosition), UEFHandler.BlockData(lPosition + 19), lBlockLength
                        fiFileInfo.Length = fiFileInfo.Length + lBlockLength
                        ReDim Preserve fiFileInfo.BlockLengths(UBound(fiFileInfo.BlockLengths) + 1)
                        fiFileInfo.BlockLengths(UBound(fiFileInfo.BlockLengths)) = lBlockLength
                        lDataPosition = lDataPosition + lBlockLength
                    End If
                End If
        End Select
    Wend
    If sFileName <> "" Then
        AddFile fiFileInfo
    End If
    
    BuildCassetteImage
    cpCassetteIndex.FileIndex = 0
    cpCassetteIndex.BlockIndex = 0
    cpCassetteIndex.DataIndex = 0
    cpCassetteIndex.CarrierTimer = 150
End Sub

Public Sub LoadDisc(ByVal sPath As String, Optional ByVal bInterleaved As Boolean)
    ' Debugging.WriteString "Storage.LoadDisc"

    Dim lImagePagePos As Long
    Dim lImagePos(1) As Long
    Dim lChunkSize As Long
    Dim lSideIndex As Long
    Dim lFileSize As Long
    Dim yImagePage() As Byte
    
    ReDim yImage(204799, 1) As Byte
    
    lFileSize = FileLen(sPath)
    If Not bInterleaved Then
        ReDim yImagePage(lFileSize - 1)
        Open sPath For Binary As #1
        Get #1, , yImagePage()
        Close #1
        CopyMemory yImage(0, 0), yImagePage(0), lFileSize
        CopyMemory myDiscImage(0, 0, 0, 0), yImage(0, 0), 204800
    Else
        Open sPath For Binary As #1
        lChunkSize = 2560&
        While lChunkSize = 2560&
            lChunkSize = lFileSize - lImagePagePos
            If lChunkSize > 2560& Then
                lChunkSize = 2560&
            End If
            If lChunkSize > 0 Then
                ReDim yImagePage(lChunkSize - 1)
                Get #1, , yImagePage()
                lImagePagePos = lImagePagePos + lChunkSize
                CopyMemory yImage(lImagePos(lSideIndex), lSideIndex), yImagePage(0), lChunkSize
                lImagePos(lSideIndex) = lImagePos(lSideIndex) + lChunkSize
                lSideIndex = 1 - lSideIndex
            End If
        Wend
        Close #1
        CopyMemory myDiscImage(0, 0, 0, 0), yImage(0, 0), 409600
    End If
End Sub


'    Five seconds of 2400Hz tone.
'    One synchronisation byte (&2A).
'    File name (one to ten characters).
'00 One end of file name marker byte (&00).
'01 Load address of file, four bytes, low byte first.
'05 Execution address of file, four bytes, low byte first.
'09 Block number, two bytes, low byte first.
'11 Data block length, two bytes, low byte first.
'13 Block flag, one byte.
'14 Spare, four bytes, currently &00.
'18 CRC on header, two bytes.
'20 Data, 0 to 256 bytes.
'    CRC on data, two bytes.
Public Sub BuildCassetteImage(Optional ByVal bBuildBlockLengths As Boolean)
    Dim lFileNumber As Long
    Dim yImage() As Byte
    Dim lDataPosition As Long
    Dim lStartPosition As Long
    Dim lBlockNumber As Long
    Dim fiFileInfo As fiFile
    Dim lBlockFlag As Long
    Dim lFileDataPos As Long
    Dim lTotalSize As Long
    Dim lBlockLast As Long
    Dim sShortName As String
    
    ' Debugging.WriteString "Storage.BuildCassetteImage"
    
    If bBuildBlockLengths Then
        For lFileNumber = 0 To mlCatalogueCount - 1
            ReDim mfiCatalogue(lFileNumber).BlockLengths(mfiCatalogue(lFileNumber).Length \ 256&)
        Next
    End If
    
    For lFileNumber = 0 To mlCatalogueCount - 1
        fiFileInfo = mfiCatalogue(lFileNumber)
        lFileDataPos = 0
        lBlockLast = UBound(fiFileInfo.BlockLengths)
        ReDim mfiCatalogue(lFileNumber).BlockDataStarts(lBlockLast)
        Erase yImage
        lDataPosition = 0
        For lBlockNumber = 0 To lBlockLast
            mfiCatalogue(lFileNumber).BlockDataStarts(lBlockNumber) = lDataPosition
            
            lTotalSize = lTotalSize + Len(fiFileInfo.FileName) + 23 + fiFileInfo.BlockLengths(lBlockNumber)
            
            ReDim Preserve yImage(lTotalSize - 1)
            
            lStartPosition = lDataPosition
            yImage(lDataPosition) = &H2A&
            sShortName = ShortFileName(fiFileInfo.Name)
            CopyMemory yImage(lDataPosition + 1), ByVal sShortName, Len(fiFileInfo.FileName)
            lDataPosition = lDataPosition + Len(sShortName) + 1
            yImage(lDataPosition) = 0
            CopyMemory yImage(lDataPosition + 1), fiFileInfo.Load, 4&
            CopyMemory yImage(lDataPosition + 5), fiFileInfo.Execution, 4&
            CopyMemory yImage(lDataPosition + 9), lBlockNumber, 2&
            CopyMemory yImage(lDataPosition + 11), fiFileInfo.BlockLengths(lBlockNumber), 2&
            lBlockFlag = -&H80& * (lBlockNumber = UBound(fiFileInfo.BlockLengths))
            CopyMemory yImage(lDataPosition + 13), lBlockFlag, 1& ' block flag
            CopyMemory yImage(lDataPosition + 18), CRC(yImage, lStartPosition + 1, lDataPosition + 17 - lStartPosition), 2& ' block flag
            If fiFileInfo.BlockLengths(lBlockNumber) > 0 Then
                CopyMemory yImage(lDataPosition + 20), fiFileInfo.Data(lFileDataPos), fiFileInfo.BlockLengths(lBlockNumber)
            End If
            lDataPosition = lDataPosition + fiFileInfo.BlockLengths(lBlockNumber) + 20
            CopyMemory yImage(lDataPosition), CRC(fiFileInfo.Data, lFileDataPos, fiFileInfo.BlockLengths(lBlockNumber)), 2&
            
            lFileDataPos = lFileDataPos + fiFileInfo.BlockLengths(lBlockNumber)
            lDataPosition = lDataPosition + 2
        Next
        mfiCatalogue(lFileNumber).BlockData = yImage
    Next
End Sub

Private Function ShortFileName(ByVal sFullFileName As String)
    Dim lDot As Long
    
    lDot = InStrRev(sFullFileName, ".")
    ShortFileName = Mid$(sFullFileName, lDot + 1)
End Function

Public Sub SaveCassetteUEF(ByVal sFileName As String)
    Dim lFileIndex As Long
    Dim lBlockIndex As Long
    Dim lNextBlockStart As Long
    Dim lBlockStart As Long
    Dim lLastBlock As Long
    
    ' Debugging.WriteString "Storage.SaveCassetteUEF"
    
    BuildCassetteImage True
    
    UEFHandler.CreateUEFFile
    
    For lFileIndex = 0 To mlCatalogueCount - 1
        lLastBlock = UBound(mfiCatalogue(lFileIndex).BlockDataStarts)
        For lBlockIndex = 0 To lLastBlock
            lBlockStart = mfiCatalogue(lFileIndex).BlockDataStarts(lBlockIndex)
            If lBlockIndex < lLastBlock Then
                lNextBlockStart = mfiCatalogue(lFileIndex).BlockDataStarts(lBlockIndex + 1)
            Else
                lNextBlockStart = UBound(mfiCatalogue(lFileIndex).BlockData) + 1
            End If

            UEFHandler.ResetBlock lNextBlockStart - lBlockStart
            CopyMemory UEFHandler.BlockData(0), mfiCatalogue(lFileIndex).BlockData(lBlockStart), lNextBlockStart - lBlockStart
            UEFHandler.SaveBlock &H100&
        Next
    Next
    
    'UEFHandler.SaveCompressedUEFFile sFileName
    UEFHandler.SaveUEFFile sFileName
End Sub

Public Sub SaveDiscImage(ByVal sFileName As String, bInterleaved As Boolean)
    Dim yImage() As Byte
    Dim lSideLength As Long
    Dim lIndex As Long
    Dim lSide As Long
    Dim lTrack As Long
    Dim oFSO As New FileSystemObject
    
    ' Debugging.WriteString "Storage.SaveDiscImage"
    
    lSideLength = 256& * 10& * 80&
    If Not bInterleaved Then
        ReDim yImage(lSideLength - 1)
        CopyMemory yImage(0), myDiscImage(0, 0, 0, 0), lSideLength
        lIndex = lSideLength - 1
        While yImage(lIndex) = 0
            lIndex = lIndex - 1
        Wend
        ReDim Preserve yImage(lIndex)
        If oFSO.FileExists(sFileName) Then
            oFSO.DeleteFile sFileName, True
        End If
        Open sFileName For Binary As #1
        Put #1, , yImage
        Close #1
    Else
        ReDim yImage(lSideLength * 2 - 1)
        lIndex = 0
        For lTrack = 0 To 79
            For lSide = 0 To 1
                CopyMemory yImage(lIndex), myDiscImage(0, 0, lTrack, lSide), 2560&
                lIndex = lIndex + 2560&
            Next
        Next
        lIndex = lSideLength * 2 - 1
        While yImage(lIndex) = 0
            lIndex = lIndex - 1
        Wend
        ReDim Preserve yImage(lIndex)
        If oFSO.FileExists(sFileName) Then
            oFSO.DeleteFile sFileName, True
        End If
        Open sFileName For Binary As #1
        Put #1, , yImage
        Close #1
    End If
End Sub

Public Sub SaveArchive(ByVal sFileName As String, bCassette As Boolean)
    ' Debugging.WriteString "Storage.SaveArchive"
    Dim lIndex As Long
    Dim oFSO As New FileSystemObject
    Dim oParentFolder As Folder
    Dim oFolder As Folder
    Dim oTS As TextStream
    Dim lSlash As Long
    Dim lDot As Long
    Dim sArchiveFolder As String
    Dim sArchiveFileName As String
    
    If True Then
        CreateCatalogeFromDiscImage
            
        lSlash = InStrRev(sFileName, "\")
        sArchiveFolder = Mid$(sFileName, lSlash + 1)
        lDot = InStrRev(sArchiveFolder, ".")
        If lDot > 0 Then
            sArchiveFolder = Left$(sArchiveFolder, lDot - 1)
        End If
        
        Set oParentFolder = oFSO.GetFolder(oFSO.GetParentFolderName(sFileName))
        Set oFolder = oFSO.CreateFolder(oParentFolder.path & "\" & sArchiveFolder)
        For lIndex = 0 To mlCatalogueCount - 1
            sArchiveFileName = sArchiveFolder & "." & mfiCatalogue(lIndex).Name
             Set oTS = oFolder.CreateTextFile(sArchiveFileName & ".INF")
             oTS.Write mfiCatalogue(lIndex).Name & " "
             oTS.Write HexNum(mfiCatalogue(lIndex).Load, 6) & " "
             oTS.Write HexNum(mfiCatalogue(lIndex).Execution, 6) & " "
             oTS.Write IIf(mfiCatalogue(lIndex).Locked, "Locked ", "")
             If lIndex < (mlCatalogueCount - 1) Then
                oTS.Write "Next=" & mfiCatalogue(lIndex + 1).Name
             End If
             oTS.Close
             
            If oFSO.FileExists(oFolder.path & "\" & sArchiveFileName) Then
                oFSO.DeleteFile oFolder.path & "\" & sArchiveFileName, True
            End If
            
            Open oFolder.path & "\" & sArchiveFileName For Binary As #1
            Put #1, , mfiCatalogue(lIndex).Data
            Close #1
        Next
    Else
    End If
End Sub

Private Sub CreateCatalogeFromDiscImage()
    Dim lFileCount As Long
    Dim lFileNumber As Long
    Dim sFileName As String * 8
    Dim sName As String * 8
    Dim sDirectory As String
    Dim fiFileInfo As fiFile
    Dim lDataAddress As Long
    Dim lCatalogueInfoAdr As Long
    Dim yImage(204799) As Byte
    Dim lSideIndex As Long
    Dim lFileInfoAdr As Long
    Dim lLogicalSide As Long
    Dim sName1 As String * 8
    Dim sName2 As String * 4
    Dim lHi As Long
    
    Erase mfiCatalogue
    mlCatalogueCount = 0
    
    For lSideIndex = 0 To 0
        CopyMemory yImage(0), myDiscImage(0, 0, 0, lSideIndex), 204800
        
        mlBootOption(lSideIndex) = (yImage(&H100& + 6) \ 16) And 3&
        CopyMemory ByVal sName1, yImage(0&), 8&
        CopyMemory ByVal sName2, yImage(&H100&), 4&
        msCatalogueName(lSideIndex) = Trim$(sName1 & sName2)

        lFileCount = yImage(&H100& + 5) \ 8

        For lFileNumber = lFileCount - 1 To 0 Step -1
            fiFileInfo.DriveNumber = lLogicalSide

            lFileInfoAdr = 8 + 8 * lFileNumber

            CopyMemory ByVal sFileName, yImage(lFileInfoAdr), 7&
            sName = Trim$(Left$(sFileName, 7))
            'Debug.Print sName
            sDirectory = Chr$(yImage(lFileInfoAdr + 7) And &H7F&)
            fiFileInfo.Locked = (yImage(lFileInfoAdr + 7) And &H80&) <> 0
            fiFileInfo.Name = Trim$(sDirectory & "." & sName)

            lCatalogueInfoAdr = &H100& + 8 + lFileNumber * 8

            fiFileInfo.Load = 0
            fiFileInfo.Execution = 0
            fiFileInfo.Length = 0

            CopyMemory fiFileInfo.Load, yImage(lCatalogueInfoAdr + 0), 2&
            CopyMemory fiFileInfo.Execution, yImage(lCatalogueInfoAdr + 2), 2&
            CopyMemory fiFileInfo.Length, yImage(lCatalogueInfoAdr + 4), 2&

            lHi = (yImage(lCatalogueInfoAdr + 6) And &HC&) \ 4&
            fiFileInfo.Load = fiFileInfo.Load + IIf(lHi > 2, &HFF0000, lHi * &H10000)
            
            lHi = (yImage(lCatalogueInfoAdr + 6) And &HC0&) \ 64&
            fiFileInfo.Execution = fiFileInfo.Execution + IIf(lHi > 2, &HFF0000, lHi * &H10000)
                        
            lHi = (yImage(lCatalogueInfoAdr + 6) And &H30&) \ 16&
            fiFileInfo.Length = fiFileInfo.Length + lHi * &H10000
            
            lDataAddress = 256& * (256& * (yImage(lCatalogueInfoAdr + 6) And &H3&) + yImage(lCatalogueInfoAdr + 7))
            'Debug.Print fiFileInfo.DriveNumber & ":" & fiFileInfo.Name & ": " & HexNum(lDataAddress, 8) & ":" & HexNum(fiFileInfo.Length, 8)

            ReDim fiFileInfo.Data(fiFileInfo.Length - 1)
            CopyMemory fiFileInfo.Data(0), yImage(lDataAddress), fiFileInfo.Length
            AddFile fiFileInfo
        Next
    Next
End Sub

Private Function CRC(lData() As Byte, lStart As Long, lLength As Long) As Long
    Dim lIndex As Long
    Dim yByte As Byte
    Dim lBit As Long
    Dim lT As Long
    Dim yTemp(1) As Byte
    
    ' Debugging.WriteString "Storage.CRC"
    
    For lIndex = lStart To lStart + lLength - 1
        CRC = CRC Xor CLng(lData(lIndex)) * 256
        For lBit = 1 To 8
            lT = 0
            If (CRC And &H8000&) <> 0 Then
                CRC = CRC Xor &H810&
                lT = 1
            End If
            CRC = (CRC * 2 + lT) And &HFFFF&
        Next
    Next
    CopyMemory yTemp(0), CRC, 2&
    CRC = CLng(yTemp(0)) * 256 + yTemp(1)
End Function

Private Function FromHex(ByVal sHex As String) As Long
    Dim lValue As Long
    Dim lPos As Long
    Dim sChars As String
    
    sChars = "0123456789ABCDEF"
    sHex = UCase$(sHex)
    For lPos = 1 To Len(sHex)
        FromHex = FromHex * 16&
        FromHex = FromHex + InStr(sChars, Mid$(sHex, lPos, 1)) - 1
    Next
End Function

Public Function BuildDiscImage(Optional ByVal bInterleaved As Boolean) As Byte()
    Dim lDot As Long
    Dim sName As String
    Dim fiFileInfo As fiFile
    Dim sFileName As String
    Dim sFileDir As String
    Dim yImage() As Byte
    Dim lAddress(1) As Long
    Dim lFileNumber(1) As Long
    Dim lSector As Long
    Dim lCatalogueInfoAdr As Long
    Dim lSideIndex As Long
    Dim lCatalogueIndex As Long
    Dim yInterleavedImage() As Byte
    Dim lDataPosition As Long
    Dim lTrack As Long
    Dim lHiBits As Long
    
    ' Debugging.WriteString "Storage.BuildDiscImage"
    
    ReDim yImage(256& * 10& * 80& - 1, 1)
    
    For lSideIndex = 0 To 1
        ' build catalogue
        yImage(&H100& + 5, lSideIndex) = CatalogueCount(lSideIndex) * 8
        yImage(&H100& + 6, lSideIndex) = (800& And &H300&) \ 256 + mlBootOption(lSideIndex) * 16
        yImage(&H100& + 7, lSideIndex) = (800& And &HFF&)
        
        lAddress(lSideIndex) = &H200&
        
        For lCatalogueIndex = mlCatalogueCount - 1 To 0 Step -1
            fiFileInfo = mfiCatalogue(lCatalogueIndex)
            If fiFileInfo.DriveNumber = lSideIndex Then
                lDot = InStr(fiFileInfo.Name, ".")
                
                If lDot > 0 Then
                    sFileDir = Left$(fiFileInfo.Name, lDot - 1)
                    sFileName = Mid$(fiFileInfo.Name, lDot + 1)
                Else
                    sFileDir = "$"
                    sFileName = Left$(fiFileInfo.Name, 7)
                End If
                
                sName = "        "
        
                Mid$(sName, 1, Len(sFileName)) = sFileName
                Mid$(sName, 8, 1) = sFileDir
                
                CopyMemory yImage(8 + lFileNumber(lSideIndex) * 8, lSideIndex), ByVal sName, 8&
                yImage(8 + lFileNumber(lSideIndex) * 8 + 7, lSideIndex) = yImage(8 + lFileNumber(lSideIndex) * 8 + 7, lSideIndex) Or -fiFileInfo.Locked * &H80&
                
                lCatalogueInfoAdr = &H100& + 8 + lFileNumber(lSideIndex) * 8
                
                CopyMemory yImage(lCatalogueInfoAdr + 0, lSideIndex), fiFileInfo.Load, 2&
                CopyMemory yImage(lCatalogueInfoAdr + 2, lSideIndex), fiFileInfo.Execution, 2&
                CopyMemory yImage(lCatalogueInfoAdr + 4, lSideIndex), fiFileInfo.Length, 2&
                        
                lSector = lAddress(lSideIndex) \ 256
                lHiBits = (lSector And &H300&) \ 256&
                lHiBits = lHiBits + IIf(fiFileInfo.Load > &H2FFFF, &HC&, 4 * ((fiFileInfo.Load \ &H10000) And 3&))
                lHiBits = lHiBits + ((fiFileInfo.Length \ &H10000) And 3&) * 16
                lHiBits = lHiBits + IIf(fiFileInfo.Execution > &H2FFFF, &HC0&, 64 * ((fiFileInfo.Execution \ &H10000) And 3&))

                yImage(lCatalogueInfoAdr + 6, lSideIndex) = lHiBits
                yImage(lCatalogueInfoAdr + 7, lSideIndex) = lSector And &HFF&
                
                CopyMemory yImage(lAddress(lSideIndex), lSideIndex), fiFileInfo.Data(0), fiFileInfo.Length
                
                lAddress(lSideIndex) = (lAddress(lSideIndex) + fiFileInfo.Length + 255) And &HFFFFFF00
                lFileNumber(lSideIndex) = lFileNumber(lSideIndex) + 1
            End If
        Next
    Next
    
    Erase lAddress
    If bInterleaved Then
        ReDim yInterleavedImage(256& * 10& * 80& * 2& - 1)
        For lTrack = 0 To 79
            For lSideIndex = 0 To 1
                CopyMemory yInterleavedImage(lDataPosition), yImage(lAddress(lSideIndex), lSideIndex), &HA00&
                lDataPosition = lDataPosition + &HA00&
                lAddress(lSideIndex) = lAddress(lSideIndex) + &HA00&
            Next
        Next
        BuildDiscImage = yInterleavedImage
    Else
        BuildDiscImage = yImage
    End If
End Function

Private Function CatalogueCount(ByVal lDriveNumber As Long) As Long
    Dim lFileNumber As Long
    
    ' Debugging.WriteString "Storage.CatalogueCount"
    
    For lFileNumber = 0 To mlCatalogueCount - 1
        If mfiCatalogue(lFileNumber).DriveNumber = lDriveNumber Then
            CatalogueCount = CatalogueCount + 1
        End If
    Next
End Function

Public Sub ShowCassetteCatalogue()
    Dim lFileNumber As Long
    Dim lBlockNumber As Long
    
    ' Debugging.WriteString "Storage.ShowCassetteCatalogue"
    
    lstFiles.Clear
    For lFileNumber = 0 To mlCatalogueCount - 1
        For lBlockNumber = 0 To UBound(mfiCatalogue(lFileNumber).BlockLengths)
            lstFiles.AddItem ShortFileName(mfiCatalogue(lFileNumber).Name) & vbTab & HexNum(lBlockNumber, 2)
        Next
    Next
    Me.Show
End Sub


' Carrier, No Carrier, Data
Public Function NextCassetteByte() As Long
    ' Debugging.WriteString "Storage.NextCassetteByte"
    
    If cpCassetteIndex.CarrierTimer > 0 Then
        If cpCassetteIndex.CarrierTimer > 50 Then
            NextCassetteByte = 256 ' carrier
        Else
            NextCassetteByte = 257 ' no carrier
        End If
        cpCassetteIndex.CarrierTimer = cpCassetteIndex.CarrierTimer - 1
    Else
        If cpCassetteIndex.FileIndex < mlCatalogueCount Then
            If cpCassetteIndex.BlockIndex < UBound(mfiCatalogue(cpCassetteIndex.FileIndex).BlockDataStarts) Then
                If cpCassetteIndex.DataIndex < mfiCatalogue(cpCassetteIndex.FileIndex).BlockDataStarts(cpCassetteIndex.BlockIndex + 1) Then
                    NextCassetteByte = mfiCatalogue(cpCassetteIndex.FileIndex).BlockData(cpCassetteIndex.DataIndex)
                    cpCassetteIndex.DataIndex = cpCassetteIndex.DataIndex + 1
                Else
                    cpCassetteIndex.BlockIndex = cpCassetteIndex.BlockIndex + 1
                    cpCassetteIndex.CarrierTimer = 100
                    NextCassetteByte = 256 ' carrier
                End If
            Else
                If cpCassetteIndex.DataIndex <= UBound(mfiCatalogue(cpCassetteIndex.FileIndex).BlockData) Then
                    NextCassetteByte = mfiCatalogue(cpCassetteIndex.FileIndex).BlockData(cpCassetteIndex.DataIndex)
                    cpCassetteIndex.DataIndex = cpCassetteIndex.DataIndex + 1
                Else
                    cpCassetteIndex.FileIndex = cpCassetteIndex.FileIndex + 1
                    cpCassetteIndex.BlockIndex = 0
                    cpCassetteIndex.DataIndex = 0
                    cpCassetteIndex.CarrierTimer = 150
                    NextCassetteByte = 256 ' carrier
                End If
            End If
        Else
            NextCassetteByte = 257 ' no carrier
        End If
    End If
End Function

Public Function ReadDiscByte(ByVal lDataPos As Long, ByVal lSector As Long, ByVal lTrack As Long, ByVal lHead As Long) As Byte
    ' Debugging.WriteString "Storage.ReadDiscByte"
    
    ReadDiscByte = myDiscImage(lDataPos, lSector, lTrack, lHead)
End Function

Public Function WriteDiscByte(ByVal lDataPos As Long, ByVal lSector As Long, ByVal lTrack As Long, ByVal lHead As Long, ByVal yData As Byte)
    ' Debugging.WriteString "Storage.WriteDiscByte"
    
    myDiscImage(lDataPos, lSector, lTrack, lHead) = yData
    If lDataPos = 0 Then
        'Debug.Print lDataPos & " " & lSector & " " & lTrack & " " & lHead
    End If
End Function


Private Sub lstFiles_Click()
    Dim lListIndex As Long
    Dim lFileIndex As Long
    Dim lBlocks As Long
    
    ' Debugging.WriteString "Storage.lstFiles_Click"
    
    lListIndex = lstFiles.ListIndex
    For lFileIndex = 0 To mlCatalogueCount
        lBlocks = UBound(mfiCatalogue(lFileIndex).BlockLengths) + 1
        If (lListIndex - lBlocks) >= 0 Then
            lListIndex = lListIndex - lBlocks
        Else
            Exit For
        End If
    Next
    
    cpCassetteIndex.FileIndex = lFileIndex
    cpCassetteIndex.BlockIndex = lListIndex
    cpCassetteIndex.DataIndex = 0
    cpCassetteIndex.CarrierTimer = 150
End Sub

Private Sub Form_Resize()
    ' Debugging.WriteString "Storage.Form_Resize"
    
    lstFiles.Width = Me.ScaleWidth
    lstFiles.Height = Me.ScaleHeight
End Sub



Public Sub StartTransfer(ByVal sFile As String)
    Dim dTime As Date
    
    mlBaseAddress = &H7700&
    
    InitialiseComPort
    TransferBasicProgram
    
    dTime = Time
    While Time < TimeSerial(Hour(dTime), Minute(dTime), Second(dTime) + 2)
        DoEvents
    Wend
    InitialiseFastComPort
    TransferDisc sFile
    Console.Com.PortOpen = False
End Sub

Private Sub InitialiseComPort()
    Console.Com.CommPort = 1
    Console.Com.Settings = "9600,N,8,1"
    Console.Com.InputLen = 0
    Console.Com.OutBufferSize = 32767
    Console.Com.PortOpen = True
End Sub


Private Sub InitialiseFastComPort()
    Console.Com.CommPort = 1
    Console.Com.Settings = "38400,N,8,1"
    Console.Com.InputLen = 0
    Console.Com.OutBufferSize = 32767
    Console.Com.PortOpen = True
End Sub

Private Sub SendWord(ByVal lWord As Long)
    Dim yOut(1) As Byte
    
    yOut(0) = lWord And &HFF
    yOut(1) = (lWord And &HFF00&) \ 256

    Console.Com.Output = yOut
End Sub

Private Sub SendBlockDetails(ByVal lAddress As Long, ByVal lLength As Long)
    lLength = 65536 - lLength
    SendWord lAddress - (lLength And &HFF&)
    SendWord lLength
End Sub

Private Sub SendByte(ByVal lAddress, ByVal yValue As Long)
    Dim yByte(0) As Byte
    
    SendBlockDetails lAddress, 1
    yByte(0) = yValue
    Console.Com.Output = yByte
End Sub

Private Sub SendData(ByVal lAddress As Long, ByVal lLength As Long, yData() As Byte)
    SendBlockDetails lAddress, lLength
    Console.Com.Output = yData
End Sub

Private Sub TransferBasicProgram()
    Dim oFSO As New FileSystemObject
    Dim sFile As String
    Dim lLineNumber As Long
    Dim vSplit As Variant
    Dim vLine As Variant
    Dim dTime As Date
    Dim lSlow As Long
    
    sFile = oFSO.OpenTextFile(App.path & "\ReadSerialBBC.txt").ReadAll
    vSplit = Split(sFile, vbCrLf)
    lLineNumber = 10
    For Each vLine In vSplit
        If vLine <> "" Then
            vLine = Replace$(vLine, "&7F00", "&" & Hex$(mlBaseAddress))
            Console.Com.Output = lLineNumber & vLine & vbCr
            lLineNumber = lLineNumber + 10
        Else
            Console.Com.Output = vbCr
        End If
        For lSlow = 0 To 10000
            DoEvents
        Next
    Next
    Console.Com.Output = "RUN" & vbCr
    'Console.Com.Output = "*FX2" & vbCr
    Console.Com.PortOpen = False
End Sub

Private Sub TransferDisc(ByVal sPath As String)
    Dim lSector As Long
    Dim lTrack As Long
    Dim ySector(255) As Byte
    
    Console.Com.Output = "*"
    
    ' Disc NMI routine
    '   Load NMI routine
    SendData &HD00&, 256, ySector
    
    For lTrack = 0 To 79
        For lSector = 0 To 9
            CopyMemory ySector(0), myDiscImage(0, lSector, lTrack, 0), 256&
            SendData &H3000&, 256, ySector
            SendByte &HFE80&, &H4B&   ' Start Disc Write
            SendByte &HFE81&, lTrack ' Track
            SendByte &HFE81&, lSector ' Sector
            SendByte &HFE81&, 1 ' Sectors to go +
        Next
    Next
    
    SendWord 0 ' Dummy start address
    SendWord 0 ' Dummy end address : Load registers and jump
End Sub
