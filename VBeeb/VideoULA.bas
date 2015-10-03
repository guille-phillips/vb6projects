Attribute VB_Name = "VideoULA"
Option Explicit

Private Declare Function StretchDIBits Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal dx As Long, ByVal dy As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal wSrcWidth As Long, ByVal wSrcHeight As Long, lpBits As Any, lpBitsInfo As BITMAPINFO, ByVal wUsage As Long, ByVal dwRop As Long) As Long
Private Declare Function FillRect Lib "user32.dll" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32.dll" (ByVal crColor As Long) As Long

Private Declare Function SetDIBitsToDevice Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal dx As Long, ByVal dy As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal Scan As Long, ByVal NumScans As Long, Bits As Any, BitsInfo As BITMAPINFO, ByVal wUsage As Long) As Long


Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type AREA
    Left As Long
    Top As Long
    Width As Long
    Height As Long
End Type

Private Type BITMAPINFOHEADER '40 bytes
    biSize As Long
    biWidth As Long
    biHeight As Long
    biPlanes As Integer
    biBitCount As Integer
    biCompression As Long
    biSizeImage As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed As Long
    biClrImportant As Long
End Type

Private Type RGBQUAD
    rgbBlue As Byte
    rgbGreen As Byte
    rgbRed As Byte
    rgbReserved As Byte
End Type

Private Type BITMAPINFO
    bmiHeader As BITMAPINFOHEADER
    bmiColors(255) As Long
End Type

Private Const BI_RGB = 0&
Private Const DIB_RGB_COLORS = 0 '  color table in RGBs
Private bmiCharacterRow As BITMAPINFO
Private bmiCharacterRowTeletext As BITMAPINFO
Private bmiCursor As BITMAPINFO
Private Const SRCCOPY = &HCC0020 ' (DWORD) dest = source


Private myDisplayMemory(40960) As Byte

Private mlField As Long

Public mlCurrentRow As Long
Public mlCharacterRowBottomScanline As Long
Public mlCharacterRowTopScanline As Long

Private mlColours(7) As Long
Private mlPhysicalColours(16) As Long

Private mlFlashColourSelect As Long
Private mlPreviousFlashColourSelect As Long

Private mlNextTeletext As Long
Private mlTeletext As Long

Private mlModeIndex As Long
Private mlPreviousModeIndex As Long

Private mlDisplayedWidth As Long

Public HardwareScrollSize As Long
Public HardwareScrollBytes As Long

Private mlConsoleHDC As Long

Private myTeletextCharsetPattern(15, 19, 255, 2, 7, 7) As Byte ' Bit/Scanline/Char/Graphic/Fore Colour/Back Colour
Private myTeletextDHCharsetPattern(15, 39, 255, 2, 7, 7) As Byte ' Bit/Scanline/Char/Graphic/Fore Colour/Back Colour

Private mlBlackBrush As Long
Private mlGreyBrush As Long
Private mlWhiteBrush As Long

Private mlTotalCycles As Long

Private mlCharacterRowAddress As Long
Private mlCharacterRow As Long
Public CharacterRowDisplayedHeight As Long
Public CharacterRowHeight As Long
Public CharacterRowTotalBytes As Long

Public ScreenBaseAddress As Long
Public ScreenTeletextBaseAddress As Long
Public CursorAddress As Long
Public CursorTeletextAddress As Long
Public CursorBlinkFieldRate As Long
Private mlCursorWidthColumns As Long

Private mlLookUpColour(32, 255, 7) As Byte
Private mlScreenByteWidth As Long
Private mlScreenPixelsPerRow As Long
Private mlBytesPerRow As Long

Public NewScreenWidth As Long
Public NewScreenHeight As Long
Private mlScreenLeft As Long
Private mlScreenTop As Long

Private mareaScreenRect As AREA
Private mareaNewScreenRect As AREA

Private mlCursorBlinkTimer As Long
Private mlCursorInvertMask As Byte

Public Register(16) As Byte

Public Sub InitialiseVideoULA()
    ' Debugging.WriteString "VideoULA.InitialiseVideoULA"

    mlPreviousFlashColourSelect = -1
    mlPreviousModeIndex = -1
    mlTeletext = 0
    
    'BuildMode7.CreateDoubleHeightTeletextFontFile
'    BuildMode7.CreateTeletextCharsetFile
'    BuildMode7.CreateTeletextDoubleHeightCharsetFile
    
    CreateColourLookups
    CreateBlackBrush
    ReadTeletextCharset
    InitialiseLogicalColours
    InitialiseBMIHeader
    InitialisePallette
    MapColours
    InitialiseBitmap
    mlConsoleHDC = Console.hdc
    
    mlCharacterRow = 0
    mlCharacterRowAddress = 0
    ScreenBaseAddress = 0
    ScreenTeletextBaseAddress = 0
    CursorAddress = 0
    CursorTeletextAddress = 0
    mlCurrentRow = 256
End Sub

' We need to map individual bit patterns in memory to pixels for efficient rendering
Private Sub CreateColourLookups()
    Dim lIndex As Long
    Dim lPattern1 As Long
    Dim lPattern2 As Long
    Dim lBitMask As Long
    Dim lColour(8) As Long
    Dim lBitIndex As Long
    Dim lColourIndex As Long
        
    ' Debugging.WriteString "VideoULA.CreateColourLookups"
    
    For lIndex = 0 To 255
        lBitMask = 128
        For lBitIndex = 0 To 7
            lColour(lBitIndex) = -((lIndex And lBitMask) <> 0)
            mlLookUpColour(lBitIndex, lIndex, 7) = lColour(lBitIndex) ' 7
            mlLookUpColour(lBitIndex * 2, lIndex, 2) = lColour(lBitIndex) ' 2
            mlLookUpColour(lBitIndex * 2 + 1, lIndex, 2) = lColour(lBitIndex)
            
            lBitMask = lBitMask \ 2
        Next
    
        lPattern1 = (lIndex And &H5) Or (lIndex And &H50) \ 8
        lPattern2 = (lIndex And &HA) \ 2 Or (lIndex And &HA0) \ 16
        lColour(0) = lPattern2 \ 4
        lColour(1) = lPattern1 \ 4
        lColour(2) = lPattern2 And 3&
        lColour(3) = lPattern1 And 3&
        
        For lColourIndex = 0 To 3
            mlLookUpColour(lColourIndex * 2, lIndex, 6) = lColour(lColourIndex) ' 6
            mlLookUpColour(lColourIndex * 2 + 1, lIndex, 6) = lColour(lColourIndex)
            
            mlLookUpColour(lColourIndex * 4, lIndex, 1) = lColour(lColourIndex) ' 1
            mlLookUpColour(lColourIndex * 4 + 1, lIndex, 1) = lColour(lColourIndex)
            mlLookUpColour(lColourIndex * 4 + 2, lIndex, 1) = lColour(lColourIndex)
            mlLookUpColour(lColourIndex * 4 + 3, lIndex, 1) = lColour(lColourIndex)
        Next
        
    
        lBitMask = 128
        lColour(0) = 0
        lColour(1) = 0
        For lBitIndex = 0 To 7
            lColour(lBitIndex And 1&) = 2 * lColour(lBitIndex And 1&) - ((lIndex And lBitMask) <> 0)
            lBitMask = lBitMask \ 2
        Next
        
        For lColourIndex = 0 To 7
            mlLookUpColour(lColourIndex * 2, lIndex, 5) = lColour(lColourIndex \ 2) '5
            mlLookUpColour(lColourIndex * 2 + 1, lIndex, 5) = lColour(lColourIndex \ 2)
            
            mlLookUpColour(lColourIndex * 4, lIndex, 0) = lColour(lColourIndex \ 2)   '0
            mlLookUpColour(lColourIndex * 4 + 1, lIndex, 0) = lColour(lColourIndex \ 2)
            mlLookUpColour(lColourIndex * 4 + 2, lIndex, 0) = lColour(lColourIndex \ 2) '0
            mlLookUpColour(lColourIndex * 4 + 3, lIndex, 0) = lColour(lColourIndex \ 2)
        Next
    Next
End Sub

Private Sub CreateBlackBrush()
    ' Debugging.WriteString "VideoULA.CreateBlackBrush"
    
    mlBlackBrush = CreateSolidBrush(vbBlack)
    mlGreyBrush = CreateSolidBrush(RGB(128, 128, 128))
    mlWhiteBrush = CreateSolidBrush(vbWhite)
End Sub

Private Sub ReadTeletextCharset()
    ' Debugging.WriteString "VideoULA.ReadTeletextCharset"
    
    Dim sSingleHeightPath As String
    Dim sDoubleHeightPath As String
    
    sSingleHeightPath = App.path & "\mode7charset.dat"
    sDoubleHeightPath = App.path & "\mode7DHcharset.dat"
    
    Open sSingleHeightPath For Binary As #1
    Get #1, , myTeletextCharsetPattern
    Close #1
    
    Open sDoubleHeightPath For Binary As #1
    Get #1, , myTeletextDHCharsetPattern
    Close #1
    
    If FileLen(sSingleHeightPath) <> 491520 Then
        Exit Sub
    End If
    
    Dim lBackColour As Long
    Dim lForeColour As Long
    Dim lGraphicIndex As Long
    Dim lChar As Long
    Dim lColumn As Long
    Dim lRow As Long
    
    Dim yOriginalBackColour As Byte
    Dim yOriginalForeColour As Byte
    Dim yPalletteIndex As Byte
    
    Dim lBackColourBlockSize As Long
    
    lBackColourBlockSize = 16& * 20& * 256& * 3& * 8&
    
    For lBackColour = 0 To 7
        If lBackColour > 0 Then
            FillMemory myTeletextCharsetPattern(0, 0, 0, 0, 0, lBackColour), lBackColourBlockSize, CByte(lBackColour)
        End If
        For lForeColour = 0 To 7
            For lGraphicIndex = 0 To 2
                For lChar = 0 To 255
                    For lColumn = 0 To 15
                        For lRow = 0 To 19
'                            If myTeletextCharsetPattern(lColumn, lRow, lChar, lGraphicIndex, 1, 0) = 1& Then
'                                myTeletextCharsetPattern(lColumn, lRow, lChar, lGraphicIndex, lForeColour, lBackColour) = lForeColour
'                            End If
                            yPalletteIndex = myTeletextCharsetPattern(lColumn, lRow, lChar, lGraphicIndex, 1, 0)
                            If yPalletteIndex <> 64& Then
                                yOriginalForeColour = yPalletteIndex And &H7&
                                yOriginalBackColour = (yPalletteIndex \ 8&) And &H7&
                                myTeletextCharsetPattern(lColumn, lRow, lChar, lGraphicIndex, lForeColour, lBackColour) = 64 + IIf(yOriginalForeColour = 1, lForeColour, lBackColour) + 8 * IIf(yOriginalBackColour = 1, lForeColour, lBackColour)
                            End If
                        Next
                    Next
                Next
            Next
        Next
    Next

    lBackColourBlockSize = 16& * 40& * 256& * 3& * 8&
    For lBackColour = 0 To 7
        If lBackColour > 0 Then
            FillMemory myTeletextDHCharsetPattern(0, 0, 0, 0, 0, lBackColour), lBackColourBlockSize, CByte(lBackColour)
        End If
        For lForeColour = 0 To 7
            For lGraphicIndex = 0 To 2
                For lChar = 0 To 255
                    For lColumn = 0 To 15
                        For lRow = 0 To 39
'                            If myTeletextDHCharsetPattern(lColumn, lRow, lChar, lGraphicIndex, 1, 0) = 1& Then
'                                myTeletextDHCharsetPattern(lColumn, lRow, lChar, lGraphicIndex, lForeColour, lBackColour) = lForeColour
'                            End If
                            yPalletteIndex = myTeletextDHCharsetPattern(lColumn, lRow, lChar, lGraphicIndex, 1, 0)
                            If yPalletteIndex <> 64& Then
                                yOriginalForeColour = yPalletteIndex And &H7&
                                yOriginalBackColour = (yPalletteIndex \ 8&) And &H7&
                                myTeletextDHCharsetPattern(lColumn, lRow, lChar, lGraphicIndex, lForeColour, lBackColour) = 64 + IIf(yOriginalForeColour = 1, lForeColour, lBackColour) + 8 * IIf(yOriginalBackColour = 1, lForeColour, lBackColour)
                            End If
                        Next
                    Next
                Next
            Next
        Next
    Next
    
    Open sSingleHeightPath For Binary As #1
    Put #1, , myTeletextCharsetPattern
    Close #1
    
    Open sDoubleHeightPath For Binary As #1
    Put #1, , myTeletextDHCharsetPattern
    Close #1
End Sub


Public Sub InitialiseLogicalColours()
    ' Debugging.WriteString "VideoULA.InitialiseLogicalColours"
    
    ' Note that the colours are back to front  BGR not RGB
    mlColours(0) = RGB(0, 0, 0)
    mlColours(1) = RGB(0, 0, 255)
    mlColours(2) = RGB(0, 255, 0)
    mlColours(3) = RGB(0, 255, 255)
    mlColours(4) = RGB(255, 0, 0)
    mlColours(5) = RGB(255, 0, 255)
    mlColours(6) = RGB(255, 255, 0)
    mlColours(7) = RGB(255, 255, 255)
End Sub

Public Sub InitialisePallette()
    ' Debugging.WriteString "VideoULA.InitialisePallette"
    
    Dim lPhysicalColourIndex As Long
    
    For lPhysicalColourIndex = 0 To 15
        mlPhysicalColours(lPhysicalColourIndex) = lPhysicalColourIndex
    Next
End Sub

Public Sub InitialiseBMIHeader()
    ' Debugging.WriteString "VideoULA.InitialiseBMIHeader"
    
    With bmiCharacterRow.bmiHeader
        .biSize = Len(bmiCharacterRow.bmiHeader)
        .biHeight = 16&
        .biPlanes = 1
        .biCompression = BI_RGB
        .biBitCount = 8
    End With
    
    With bmiCharacterRowTeletext.bmiHeader
        .biSize = Len(bmiCharacterRow.bmiHeader)
        .biHeight = 20&
        .biPlanes = 1
        .biCompression = BI_RGB
        .biBitCount = 8
    End With
    
    Dim lColourIndex As Long
    For lColourIndex = 0 To 7
        bmiCharacterRowTeletext.bmiColors(lColourIndex) = mlColours(lColourIndex)
    Next
    
    Dim lBackColourIndex As Long
    Dim lForeColourIndex As Long
    Dim lMixedColour As Long
    Dim lSplitColour(3, 2) As Byte
    Dim lRGBIndex As Long
    
    For lBackColourIndex = 0 To 7
        For lForeColourIndex = 0 To 7
            CopyMemory lSplitColour(0, 0), mlColours(lForeColourIndex), 4&
            CopyMemory lSplitColour(0, 1), mlColours(lBackColourIndex), 4&
            
            For lRGBIndex = 0 To 2
                lSplitColour(lRGBIndex, 2) = (lSplitColour(lRGBIndex, 0) + 2 * lSplitColour(lRGBIndex, 1)) \ 3
            Next
            CopyMemory lMixedColour, lSplitColour(0, 2), 4&
            bmiCharacterRowTeletext.bmiColors(64 + lForeColourIndex + 8 * lBackColourIndex) = lMixedColour
        Next
    Next
    
    With bmiCursor.bmiHeader
        .biSize = Len(bmiCharacterRow.bmiHeader)
        .biHeight = 16&
        .biPlanes = 1
        .biCompression = BI_RGB
        .biBitCount = 8
    End With
End Sub


Public Sub WriteRegister(ByVal lRegister As Long, ByVal lRegisterValue As Long)
    Dim lLogicalColour As Long
    Dim bDimensionsChanged As Boolean
    
    ' Debugging.WriteString "VideoULA.WriteRegister"
    
'    If lRegisterValue = 224 Then
'        'Stop
'    End If
    Select Case lRegister
        Case 0
            Register(0) = lRegisterValue
            
            mlFlashColourSelect = lRegisterValue And 1&
            If mlPreviousFlashColourSelect <> mlFlashColourSelect Then
                mlPreviousFlashColourSelect = mlFlashColourSelect
                MapColours
            End If
            
            mlNextTeletext = lRegisterValue And 2&
            
            mlCursorWidthColumns = Array(0&, 1&, 1&, 1&, 1&, 1&, 2&, 4&)((lRegisterValue And &HE0&) \ 32&)
            mlModeIndex = (lRegisterValue And &H1C&) \ 4&
            
            If mlModeIndex = 5 Then
                mlCursorInvertMask = &H3F&
            Else
                mlCursorInvertMask = &HFF&
            End If
            
            If mlTeletext = 0& Then
                If mlModeIndex <> mlPreviousModeIndex Then
                    mlPreviousModeIndex = mlModeIndex
                    InitialiseBitmap
                    mareaScreenRect.Width = mlScreenPixelsPerRow
                End If
            End If
        Case 1
            Select Case mlModeIndex
                Case 1, 6 ' mode 5 / 1
                    lLogicalColour = (lRegisterValue And &HF0&) \ 32&
                    lLogicalColour = (lLogicalColour And 1&) + (lLogicalColour And 4&) \ 2&
                    mlPhysicalColours(lLogicalColour) = lRegisterValue And &HF& Xor 7&
                    MapColours
                    Register((lRegisterValue And &HF0&) \ 16& + 1) = lRegisterValue And &HF&
                Case 0, 2, 7 ' mode 0 / 3 / 4 / 6
                    lLogicalColour = (lRegisterValue And &HF0&) \ 128&
                    mlPhysicalColours(lLogicalColour) = lRegisterValue And &HF& Xor 7&
                    Register((lRegisterValue And &HF0&) \ 16& + 1) = lRegisterValue And &HF&
                    MapColours
                Case 5 ' mode 2
                    lLogicalColour = (lRegisterValue And &HF0&) \ 16&
                    mlPhysicalColours(lLogicalColour) = lRegisterValue And &HF& Xor 7&
                    Register((lRegisterValue And &HF0&) \ 16& + 1) = lRegisterValue And &HF&
                    MapColours
            End Select
    End Select
End Sub

Public Sub MapColours()
    Dim lPhysicalColourIndex As Long
    Dim lPhysicalColour As Long
    
    ' Debugging.WriteString "VideoULA.MapColours"
    
    For lPhysicalColourIndex = 0 To 15
        lPhysicalColour = mlPhysicalColours(lPhysicalColourIndex)
        If mlPhysicalColours(lPhysicalColourIndex) < 8 Then
            bmiCharacterRow.bmiColors(lPhysicalColourIndex) = mlColours(lPhysicalColour)
            bmiCursor.bmiColors(lPhysicalColourIndex) = mlColours(lPhysicalColour)
        Else
            If mlFlashColourSelect = 1& Then
                bmiCharacterRow.bmiColors(lPhysicalColourIndex) = mlColours(15 - lPhysicalColour)
                bmiCursor.bmiColors(lPhysicalColourIndex) = mlColours(15 - lPhysicalColour)
            Else
                bmiCharacterRow.bmiColors(lPhysicalColourIndex) = mlColours(lPhysicalColour - 8)
                bmiCursor.bmiColors(lPhysicalColourIndex) = mlColours(lPhysicalColour - 8)
            End If
        End If
    Next
End Sub

Public Sub InitialiseBitmap()
    ' Debugging.WriteString "VideoULA.InitialiseBitmap"
    
    If mlTeletext = 0& Then
        mlScreenByteWidth = Array(32, 16, 16, 0, 0, 8, 8, 8)(mlModeIndex)
        bmiCharacterRow.bmiHeader.biHeight = 16&
        mlBytesPerRow = 8& * CRTC6845.Columns
        mlScreenPixelsPerRow = mlScreenByteWidth * CRTC6845.Columns
        bmiCharacterRow.bmiHeader.biWidth = mlScreenPixelsPerRow
        
        bmiCursor.bmiHeader.biWidth = 16&
    Else
        mlScreenByteWidth = 16&
        mlBytesPerRow = CRTC6845.Columns
        
        mlScreenPixelsPerRow = CRTC6845.Columns * mlScreenByteWidth
        bmiCharacterRowTeletext.bmiHeader.biWidth = mlScreenPixelsPerRow
        
        bmiCursor.bmiHeader.biWidth = 16&
    End If
End Sub

Public Sub Tick(ByVal lCycles As Long)
    ' Debugging.WriteString "VideoULA.Tick"
    
    mlTotalCycles = mlTotalCycles + lCycles
    If mlTotalCycles >= 0& Then
        mlTotalCycles = mlTotalCycles - (2 - (mlModeIndex \ 4&)) * (CRTC6845.Register(0) + 1) * (CRTC6845.Register(9) + 1) \ (1 + mlTeletext \ 2)  ' cycles per character row
        
        If mlCurrentRow = 0 Then
            'mlTotalCycles = mlTotalCycles - (2 - (mlModeIndex \ 4&)) * (CRTC6845.Register(0) + 1) * CRTC6845.Register(5)
        End If
        UpdateDisplay
    End If
End Sub

Public Sub UpdateDisplay()
    ' Debugging.WriteString "VideoULA.UpdateDisplay"
    
    If mlTeletext = 0& Then
'        Do
            UpdateDisplayAnyColour
'        Loop Until mlCurrentRow = 0
'        Do
'            UpdateDisplayAnyColour
'        Loop Until mlCurrentRow = 0
    Else
        UpdateDisplayTeletext
    End If
End Sub

Public Sub UpdateWholeDisplay()
    ' Debugging.WriteString "VideoULA.UpdateWholeDisplay"
    Dim lCurrentRow As Long
    Dim lCursorBlinkTimer As Long
    
    lCurrentRow = mlCurrentRow
    lCursorBlinkTimer = mlCursorBlinkTimer
    If mlTeletext = 0& Then
        Do
            UpdateDisplayAnyColour False
        Loop Until mlCurrentRow = lCurrentRow
    Else
        Do
            UpdateDisplayTeletext False
        Loop Until mlCurrentRow = lCurrentRow
    End If
    mlCursorBlinkTimer = lCursorBlinkTimer
End Sub

Private Sub CheckDimensionsChangedTeletext()
    ' Debugging.WriteString "VideoULA.CheckDimensionsChangedTeletext"
    
    If CRTC6845.DisplayDimensionsChanged Then
        mlScreenLeft = (NewScreenWidth - 47& * 16&) \ 2& - CRTC6845.StartX * mlScreenByteWidth + 847& + 32&
        mlScreenTop = (NewScreenHeight - 38& * 16&) \ 2& + (CRTC6845.StartY - CRTC6845.VerticalSyncPosition) * (CRTC6845.ScanlinesPerRow + 1) + CRTC6845.ScanAdjust
        
        With mareaNewScreenRect
            .Left = mlScreenLeft
            .Top = mlScreenTop
            .Width = mlScreenPixelsPerRow
            .Height = CRTC6845.Rows * (CRTC6845.ScanlinesPerRow + 2)
        End With
        
        BlankDisplay mareaScreenRect, mareaNewScreenRect
        mareaScreenRect = mareaNewScreenRect
        CRTC6845.DisplayDimensionsChanged = False
    End If
End Sub

Private Sub CheckDimensionsChangedAnyColour()
    ' Debugging.WriteString "VideoULA.CheckDimensionsChangedAnyColour"
    
    If CRTC6845.DisplayDimensionsChanged Then
        mlScreenLeft = (NewScreenWidth - 48& * 16&) \ 2& - CRTC6845.StartX * mlScreenByteWidth + 847&
        mlScreenTop = NewScreenHeight \ 2& - 316& + (CRTC6845.StartY - CRTC6845.VerticalSyncPosition) * ((CRTC6845.ScanlinesPerRow + 1) * 2) + CRTC6845.ScanAdjust * 2
        
        With mareaNewScreenRect
            .Left = mlScreenLeft
            .Top = mlScreenTop
            .Width = mlScreenPixelsPerRow
            .Height = CRTC6845.Rows * (CRTC6845.ScanlinesPerRow + 1) * 2
        End With
        
        BlankDisplay mareaScreenRect, mareaNewScreenRect
        mareaScreenRect = mareaNewScreenRect
        CRTC6845.DisplayDimensionsChanged = False
    End If
End Sub

Public Sub UpdateLightPenPosition(ByVal lX As Single, ByVal lY As Single)
    Dim lCharX As Long
    Dim lCharY As Long
    Dim lScreenAddress As Long
    Dim lMem As Long
    
    Dim lOffsetX As Long
    Dim lOffsetY As Long
    Dim lByteRow As Long
    
    lOffsetX = lX - mlScreenLeft
    lOffsetY = lY - mlScreenTop
    
    If lOffsetX < 0 Then
        Exit Sub
    End If
    
    If lOffsetY < 0 Then
        Exit Sub
    End If
    lCharX = lOffsetX \ mlScreenByteWidth
    If lCharX >= CRTC6845.Columns Then
        Exit Sub
    End If
    
    CRTC6845.LightPenCharX = lOffsetX \ mlScreenByteWidth
    If mlTeletext = 2& Then
        lCharY = lOffsetY \ (CRTC6845.ScanlinesPerRow + 2)
        If lCharY >= CRTC6845.Rows Then
            Exit Sub
        End If
        
        CRTC6845.LightPenCharY = lCharY
        lMem = ScreenTeletextBaseAddress + CRTC6845.LightPenCharX + CRTC6845.LightPenCharY * CRTC6845.Columns
    Else
        lCharY = lOffsetY \ ((CRTC6845.ScanlinesPerRow + 1) * 2)
        If lCharY >= CRTC6845.Rows Then
            Exit Sub
        End If
        CRTC6845.LightPenCharY = lCharY
        lByteRow = ((lOffsetY \ 2) Mod (ScanlinesPerRow + 2))
        If lByteRow >= 8 Then
            Exit Sub
        End If
        lMem = ScreenBaseAddress + CRTC6845.LightPenCharX * 8& + CRTC6845.LightPenCharY * CRTC6845.Columns * 8& + lByteRow
    End If
    
    'Console.Caption = CRTC6845.LightPenCharX & "," & CRTC6845.LightPenCharY & "=" & gyMem(lMem) & " " & Chr$(gyMem(lMem))
    'Console.Caption = HexNum(lMem, 4)
End Sub

Public Function GetAddress(ByVal lX As Single, ByVal lY As Single)
    Dim lCharX As Long
    Dim lCharY As Long
    Dim lScreenAddress As Long
    Dim lMem As Long
    
    Dim lOffsetX As Long
    Dim lOffsetY As Long
    Dim lByteRow As Long
    
    lOffsetX = lX - mlScreenLeft
    lOffsetY = lY - mlScreenTop
    
    If lOffsetX < 0 Then
        Exit Function
    End If
    
    If lOffsetY < 0 Then
        Exit Function
    End If
    
    lCharX = lOffsetX \ mlScreenByteWidth
    If mlTeletext = 2& Then
        lCharY = lOffsetY \ (CRTC6845.ScanlinesPerRow + 2)
        If lCharY >= CRTC6845.Rows Then
            Exit Function
        End If
        lMem = ScreenTeletextBaseAddress + lCharX + lCharY * CRTC6845.Columns
    Else
        lCharY = lOffsetY \ ((CRTC6845.ScanlinesPerRow + 1) * 2)
        If lCharY >= CRTC6845.Rows Then
            Exit Function
        End If
        lByteRow = ((lOffsetY \ 2) Mod (ScanlinesPerRow + 2))
        If lByteRow >= 8 Then
            Exit Function
        End If
        lMem = ScreenBaseAddress + lCharX * 8& + lCharY * CRTC6845.Columns * 8& + lByteRow
    End If
    
    'Console.Caption = lCharX & "," & lCharY & "=" & gyMem(lMem) & " " & Chr$(gyMem(lMem))
    Console.Caption = HexNum(lMem, 4)
    GetAddress = lMem
End Function

Public Sub UpdateDisplayTeletext(Optional ByVal bTriggerVerticalSync As Boolean = True)
    Dim lColumn As Long
    Dim yValue As Byte
    Dim lScanLine As Long
    Dim lCharX As Long
    Dim lCharXAddScanlineByColumnsAdjusted As Long
    Dim lBitPatternIndex As Long
    Dim lAddress As Long
    
    Dim lPattern As Long
    Dim lBitIndex As Long
    Dim lMask As Long
    
    Dim lForeColour As Long
    Dim lBackColour As Long
    Dim lGraphicIndex As Long
    Dim lSubGraphicIndex As Long
    Dim lFlashOn As Long
    Dim lHoldGraphic As Long
    Dim bHoldGraphics As Long
    Dim bDoubleHeight As Boolean
    Static bThisRowDoubleHeightBottom As Boolean
    Static bThisRowDoubleHeight As Boolean
    
    Static lFlashMark As Long
    Static lFlashSpace As Long
    Static lFlashIndex As Long
    
    Dim lCursorColumn As Long
    Dim lCursorCompareAddress As Long

    
    ' Debugging.WriteString "VideoULA.UpdateDisplayTeletext"
    
    If mlCurrentRow > CRTC6845.TotalRows Then
        mlTeletext = mlNextTeletext
        If bTriggerVerticalSync Then
            CRTC6845.VerticalSync
        End If
        mlCurrentRow = CRTC6845.TotalRows
        
        If mlTeletext = 0& Then
            CRTC6845.DisplayDimensionsChanged = True
            InitialiseBitmap
            UpdateDisplayAnyColour
            Exit Sub
        End If
        
        CheckDimensionsChangedTeletext

        mlCurrentRow = 0
        mlTotalCycles = -2 * (8 - CRTC6845.VerticalSyncWidth) * (CRTC6845.TotalColumns + 1)
        
        mlCharacterRowAddress = ScreenTeletextBaseAddress
        mlCharacterRowTopScanline = mlScreenTop

        If lFlashIndex = 0 Then
            lFlashMark = lFlashMark - 1
            If lFlashMark = -50 Then
                lFlashMark = 0
                lFlashIndex = 1 - lFlashIndex
            End If
        Else
            lFlashSpace = lFlashSpace - 1
            If lFlashSpace = -25 Then
                lFlashSpace = 0
                lFlashIndex = 1 - lFlashIndex
            End If
        End If
            
        If mlCursorWidthColumns > 0 Then
            If CRTC6845.CursorBlankingDelay = &HC0& Then
                mlCursorBlinkTimer = CursorBlinkFieldRate ' cursor off
            ElseIf CRTC6845.CursorBlink Then
                mlCursorBlinkTimer = mlCursorBlinkTimer - 1
                If mlCursorBlinkTimer <= (-CursorBlinkFieldRate + 1) Then
                    mlCursorBlinkTimer = CursorBlinkFieldRate
                End If
            Else
                If CursorBlinkFieldRate = 16 Then
                    mlCursorBlinkTimer = CursorBlinkFieldRate ' cursor off
                Else
                    mlCursorBlinkTimer = 0 ' cursor on
                End If
            End If
        Else
            mlCursorBlinkTimer = CursorBlinkFieldRate
        End If
        
        bThisRowDoubleHeightBottom = False
        
        ' Vertical Sync
        If bTriggerVerticalSync Then
            SystemVIA6522.AssertCA1
        End If
        DoEvents
        Exit Sub
    End If
    
    If mlCurrentRow < CRTC6845.RowDisplayStart Then
        mlCurrentRow = mlCurrentRow + 1
        Exit Sub
    End If
    
    If mlCurrentRow >= (CRTC6845.Rows + CRTC6845.RowDisplayStart) Then
        mlCurrentRow = mlCurrentRow + 1
        Exit Sub
    End If
    
    If mlCurrentRow = (CRTC6845.RowDisplayStart + CRTC6845.LightPenCharY) Then
        CRTC6845.UpdateLightPenRegisters (((mlCharacterRowAddress - &H7400& + &H10000) And &HFFFF&) Xor &H2000&)
    End If
    
    lCursorColumn = -10
    
    lAddress = mlCharacterRowAddress
    lCursorCompareAddress = mlCharacterRowAddress
    
    If mlCursorBlinkTimer <= 0 Then
        For lCharX = 0& To CRTC6845.Columns - 1
            If lCursorCompareAddress = CursorTeletextAddress Then
                lCursorColumn = lCharX
                Exit For
            End If
            lCursorCompareAddress = lCursorCompareAddress + 1
        Next
    End If
    
    bThisRowDoubleHeight = False
    
    Dim lDisplayBlankingDelay As Long
    Dim bDisplayDisabled As Boolean
    
    lDisplayBlankingDelay = (CRTC6845.Interlace \ 16& And 3&) - 1
    bDisplayDisabled = lDisplayBlankingDelay = 2
    
    lBitPatternIndex = mlScreenPixelsPerRow * 19
    For lScanLine = 0& To 19&
        lForeColour = 7
        lBackColour = 0
        lGraphicIndex = 0
        lSubGraphicIndex = 0
        lFlashOn = 0
        bHoldGraphics = False
        lHoldGraphic = 32
        bDoubleHeight = False
        
        For lCharX = 0& To CRTC6845.Columns - 1
            If bDisplayDisabled Then
                yValue = 32
            Else
                yValue = gyMem((lAddress + lDisplayBlankingDelay) And &HFFFF&)
            End If

            Select Case yValue
                Case 32 To 127, 160 To 255 ' printable chars
                    If lFlashIndex = 1 And lFlashOn = 1 Then
                        lForeColour = lBackColour
                    End If
                    If bHoldGraphics Then
                        lHoldGraphic = yValue
                    End If
                    ' Blank if single height of double height bottom
                    If Not bDoubleHeight Then
                        If bThisRowDoubleHeightBottom Then
                            yValue = 32
                        End If
                    End If
                Case 129 To 135 ' text colour
                    If lFlashIndex = 0 Or lFlashOn = 0 Then
                        lForeColour = yValue - 128
                    Else
                        lForeColour = lBackColour
                    End If
                    lGraphicIndex = 0
                Case 136 ' flash on
                    lFlashOn = 1
                    If bHoldGraphics Then
                        yValue = lHoldGraphic
                    End If
                Case 137 ' flash off
                    lFlashOn = 0
                    If bHoldGraphics Then
                        yValue = lHoldGraphic
                    End If
                Case 140 ' double height off
                    bDoubleHeight = False
                    If bHoldGraphics Then
                        yValue = lHoldGraphic
                    End If
                Case 141 ' double height on
                    bDoubleHeight = True
                    bThisRowDoubleHeight = True
                    If bHoldGraphics Then
                        yValue = lHoldGraphic
                    End If
                Case 145 To 151 ' graphics colour
                    lForeColour = yValue - 144
                    lGraphicIndex = lSubGraphicIndex + 1
                    If bHoldGraphics Then
                        yValue = lHoldGraphic
                    End If
                Case 152 ' conceal
                    lForeColour = lBackColour
                    If bHoldGraphics Then
                        yValue = lHoldGraphic
                    End If
                Case 153 ' contiguous graphics
                    lSubGraphicIndex = 0
                    If bHoldGraphics Then
                        yValue = lHoldGraphic
                    End If
                Case 155 ' nothing
                    If bHoldGraphics Then
                        yValue = lHoldGraphic
                    End If
                Case 154 ' separated graphics
                    lSubGraphicIndex = 1
                    If lGraphicIndex > 0 Then
                        lGraphicIndex = 2
                    End If
                    If bHoldGraphics Then
                        yValue = lHoldGraphic
                    End If
                Case 156 ' black background
                    lBackColour = 0
                    If bHoldGraphics Then
                        yValue = lHoldGraphic
                    End If
                Case 157 ' New background
                    lBackColour = lForeColour
                    If bHoldGraphics Then
                        yValue = lHoldGraphic
                    End If
                Case 158 ' hold graphics
                    bHoldGraphics = True
                    If bHoldGraphics Then
                        yValue = lHoldGraphic
                    End If
                Case 159 ' release graphics
                    bHoldGraphics = False
            End Select
            
            If lCharX = lCursorColumn Then
                If lScanLine >= CRTC6845.CursorStart And lScanLine <= CRTC6845.CursorEnd Then
                    CopyMemory myDisplayMemory(lBitPatternIndex), myTeletextCharsetPattern(0, lScanLine, yValue, lGraphicIndex, 7 - lForeColour, 7 - lBackColour), 16&
                Else
                    CopyMemory myDisplayMemory(lBitPatternIndex), myTeletextCharsetPattern(0, lScanLine, yValue, lGraphicIndex, lForeColour, lBackColour), 16&
                End If
           Else
                If Not bDoubleHeight Then
                    CopyMemory myDisplayMemory(lBitPatternIndex), myTeletextCharsetPattern(0, lScanLine, yValue, lGraphicIndex, lForeColour, lBackColour), 16&
                Else
                    If Not bThisRowDoubleHeightBottom Then
                        CopyMemory myDisplayMemory(lBitPatternIndex), myTeletextDHCharsetPattern(0, lScanLine, yValue, lGraphicIndex, lForeColour, lBackColour), 16&
                    Else
                        CopyMemory myDisplayMemory(lBitPatternIndex), myTeletextDHCharsetPattern(0, lScanLine + 20, yValue, lGraphicIndex, lForeColour, lBackColour), 16&
                    End If
                End If
            End If
            
            lBitPatternIndex = lBitPatternIndex + mlScreenByteWidth
            
            lAddress = lAddress + 1
            If lAddress >= &H8000& Then lAddress = lAddress - &H400&
        Next
        lBitPatternIndex = lBitPatternIndex - 2 * mlScreenPixelsPerRow
        lAddress = mlCharacterRowAddress
    Next
    
    If bThisRowDoubleHeight And Not bThisRowDoubleHeightBottom Then
        bThisRowDoubleHeightBottom = True
    Else
        bThisRowDoubleHeightBottom = False
    End If
    
    SetDIBitsToDevice mlConsoleHDC, mlScreenLeft, mlCharacterRowTopScanline, mlScreenPixelsPerRow, 20&, 0&, 0&, 0&, 20&, myDisplayMemory(0), bmiCharacterRowTeletext, DIB_RGB_COLORS

    mlCurrentRow = mlCurrentRow + 1
    mlCharacterRowTopScanline = mlCharacterRowTopScanline + 20&

    mlCharacterRowAddress = mlCharacterRowAddress + CRTC6845.Columns
    If mlCharacterRowAddress >= &H8000& Then mlCharacterRowAddress = mlCharacterRowAddress - &H400&
End Sub


Public Sub UpdateDisplayAnyColour(Optional ByVal bTriggerVerticalSync As Boolean = True)
    Dim lColumn As Long
    Dim yValue As Byte
    Dim lScanLine As Long
    Dim lCharX As Long
    Dim lBitPatternIndex As Long
    Dim lBitPatternIndex2 As Long
    Dim lAddress As Long
    
    Dim lScanlineAddress As Long
    Dim rRect As RECT
    
    Dim lCursorColumn As Long
    Dim lCursorCompareAddress As Long

    
    Static lInterlaceFrame As Long
    
    ' Debugging.WriteString "VideoULA.UpdateDisplayAnyColour"
    
    If mlCurrentRow > CRTC6845.TotalRows Then
        If bTriggerVerticalSync Then
            CRTC6845.VerticalSync
        End If
        mlCurrentRow = TotalRows
        
        mlTeletext = mlNextTeletext
        If mlTeletext = 2& Then
            CRTC6845.DisplayDimensionsChanged = True
            InitialiseBitmap
            UpdateDisplayTeletext
            Exit Sub
        End If
        
        CheckDimensionsChangedAnyColour
        
        mlCurrentRow = 0
        mlTotalCycles = -2 * (8 - CRTC6845.VerticalSyncWidth) * (CRTC6845.TotalColumns + 1) - 2 * (CRTC6845.Register(5) * (CRTC6845.TotalColumns + 1))
        
        
        mlCharacterRowAddress = ScreenBaseAddress
        mlCharacterRowTopScanline = mlScreenTop

        If mlCursorWidthColumns > 0 Then
            If CRTC6845.CursorBlankingDelay = &HC0& Then
                mlCursorBlinkTimer = CursorBlinkFieldRate ' cursor off
            ElseIf CRTC6845.CursorBlink Then
                mlCursorBlinkTimer = mlCursorBlinkTimer - 1
                If mlCursorBlinkTimer <= (-CursorBlinkFieldRate + 1) Then
                    mlCursorBlinkTimer = CursorBlinkFieldRate
                End If
            Else
                If CursorBlinkFieldRate = 16 Then
                    mlCursorBlinkTimer = CursorBlinkFieldRate
                Else
                    mlCursorBlinkTimer = 0
                End If
            End If
        Else
            mlCursorBlinkTimer = CursorBlinkFieldRate
        End If
                
        lInterlaceFrame = 1 - lInterlaceFrame
        
        ' Vertical Sync
        If bTriggerVerticalSync Then
            SystemVIA6522.AssertCA1
        End If
        DoEvents
        Exit Sub
    End If
    
    If mlCurrentRow < CRTC6845.RowDisplayStart Then
        mlCurrentRow = mlCurrentRow + 1
        Exit Sub
    End If
    
    If mlCurrentRow >= (CRTC6845.Rows + CRTC6845.RowDisplayStart) Then
        mlCurrentRow = mlCurrentRow + 1
        Exit Sub
    End If
    
    If mlCurrentRow = (CRTC6845.RowDisplayStart + CRTC6845.LightPenCharY) Then
        CRTC6845.UpdateLightPenRegisters mlCharacterRowAddress \ 8
    End If
    
    lCursorColumn = -1
    
    lAddress = mlCharacterRowAddress
    lScanlineAddress = mlCharacterRowAddress
    
    lBitPatternIndex = mlScreenPixelsPerRow * 15
    lBitPatternIndex2 = mlScreenPixelsPerRow * 14
    
    lCursorCompareAddress = mlCharacterRowAddress
    lCursorColumn = -100
    
    If mlCursorBlinkTimer <= 0 Then
        If CRTC6845.CursorEnd >= CRTC6845.CursorStart Then
            For lCharX = 0& To CRTC6845.Columns - 1
                If lCursorCompareAddress = CursorAddress Then
                    lCursorColumn = lCharX
                    Exit For
                End If
                lCursorCompareAddress = lCursorCompareAddress + 8
            Next
        End If
    End If
    
    If CRTC6845.ScanlinesPerRow > 7& Then
        rRect.Top = mlCharacterRowTopScanline + CharacterRowDisplayedHeight
        rRect.Bottom = rRect.Top + CharacterRowHeight - CharacterRowDisplayedHeight
        rRect.Left = mlScreenLeft
        rRect.Right = mlScreenLeft + mlScreenPixelsPerRow
        FillRect mlConsoleHDC, rRect, mlBlackBrush
        If lCursorColumn > -1 Then
            If CRTC6845.CursorStart <= CRTC6845.ScanlinesPerRow Then
                If CRTC6845.CursorEnd > 7& Then
                    If CRTC6845.CursorStart > 8& Then
                        rRect.Top = rRect.Top + (CRTC6845.CursorStart - 8&) * 2
                    End If
                    If CRTC6845.CursorEnd < CRTC6845.ScanlinesPerRow Then
                        rRect.Bottom = rRect.Bottom - (CRTC6845.ScanlinesPerRow - CRTC6845.CursorEnd) * 2
                    End If
                    rRect.Left = rRect.Left + lCursorColumn * mlScreenByteWidth
                    rRect.Right = rRect.Left + mlScreenByteWidth
                    FillRect mlConsoleHDC, rRect, mlWhiteBrush
                End If
            End If
        End If
    End If
    
    Dim lDisplayBlankingDelay As Long
    Dim bDisplayDisabled As Boolean
    
    lDisplayBlankingDelay = (CRTC6845.Interlace \ 16& And 3&)
    bDisplayDisabled = lDisplayBlankingDelay = 3
    
    For lScanLine = 0& To 7&
        For lCharX = 0& To CRTC6845.Columns - 1
            If bDisplayDisabled Then
                yValue = 0
            Else
                yValue = gyMem(lAddress + lDisplayBlankingDelay * 8)
            End If
            
'            If (lScanLine And 1&) = lInterlaceFrame Then
'                yValue = 0
'            End If
            
            CopyMemory myDisplayMemory(lBitPatternIndex), mlLookUpColour(0, yValue, mlModeIndex), mlScreenByteWidth
            CopyMemory myDisplayMemory(lBitPatternIndex2), mlLookUpColour(0, yValue, mlModeIndex), mlScreenByteWidth
            
            If lCharX >= lCursorColumn And lCharX < (lCursorColumn + mlCursorWidthColumns) Then
                If lScanLine >= CRTC6845.CursorStart And lScanLine <= CRTC6845.CursorEnd Then
                    CopyMemory myDisplayMemory(lBitPatternIndex), mlLookUpColour(0, yValue Xor mlCursorInvertMask, mlModeIndex), mlScreenByteWidth
                    CopyMemory myDisplayMemory(lBitPatternIndex2), mlLookUpColour(0, yValue Xor mlCursorInvertMask, mlModeIndex), mlScreenByteWidth
                End If
            End If
            
            lBitPatternIndex = lBitPatternIndex + mlScreenByteWidth
            lBitPatternIndex2 = lBitPatternIndex2 + mlScreenByteWidth
            
            lAddress = lAddress + 8
            If lAddress >= &H8000& Then lAddress = lAddress - HardwareScrollBytes
        Next
        lScanlineAddress = lScanlineAddress + 1
        
        lAddress = lScanlineAddress
        
        lBitPatternIndex = lBitPatternIndex - mlScreenPixelsPerRow * 3
        lBitPatternIndex2 = lBitPatternIndex2 - mlScreenPixelsPerRow * 3
    Next
    
    SetDIBitsToDevice mlConsoleHDC, mlScreenLeft, mlCharacterRowTopScanline, mlScreenPixelsPerRow, CharacterRowDisplayedHeight, 0&, 0&, 0&, 16&, myDisplayMemory(0), bmiCharacterRow, DIB_RGB_COLORS

    mlCurrentRow = mlCurrentRow + 1
    mlCharacterRowTopScanline = mlCharacterRowTopScanline + CharacterRowHeight

    mlCharacterRowAddress = mlCharacterRowAddress + CharacterRowTotalBytes
    If mlCharacterRowAddress >= &H8000& Then mlCharacterRowAddress = mlCharacterRowAddress - HardwareScrollBytes
End Sub

Public Sub BlankDisplay(recOldArea As AREA, recNewArea As AREA)
    Dim recOverlap As RECT
    Dim lOverlapIndex As Long
    Dim recBlock As RECT
    
    ' Debugging.WriteString "VideoULA.BlankDisplay"
    
    mlGreyBrush = mlBlackBrush
    
    If recOldArea.Left < recNewArea.Left Then
        lOverlapIndex = 1
    End If
    If recOldArea.Top < recNewArea.Top Then
        lOverlapIndex = lOverlapIndex + 2
    End If
    If (recOldArea.Left + recOldArea.Width) > (recNewArea.Left + recNewArea.Width) Then
        lOverlapIndex = lOverlapIndex + 4
    End If
    If (recOldArea.Top + recOldArea.Height) > (recNewArea.Top + recNewArea.Height) Then
        lOverlapIndex = lOverlapIndex + 8
    End If
    
    'NW
    If (lOverlapIndex And 3&) = 3& Then
        recBlock.Top = recOldArea.Top
        recBlock.Left = recOldArea.Left
        recBlock.Right = recNewArea.Left
        recBlock.Bottom = recNewArea.Top
        
        FillRect mlConsoleHDC, recBlock, mlGreyBrush
    End If
    
    'N
    If (lOverlapIndex And 2&) = 2& Then
        recBlock.Left = recNewArea.Left
        recBlock.Right = recNewArea.Left + recNewArea.Width
        recBlock.Top = recOldArea.Top
        recBlock.Bottom = recNewArea.Top
        
        FillRect mlConsoleHDC, recBlock, mlGreyBrush
    End If

    'NE
    If (lOverlapIndex And 6&) = 6& Then
        recBlock.Left = recNewArea.Left + recNewArea.Width
        recBlock.Right = recOldArea.Left + recOldArea.Width
        recBlock.Top = recOldArea.Top
        recBlock.Bottom = recNewArea.Top

        FillRect mlConsoleHDC, recBlock, mlGreyBrush
    End If

    'E
    If (lOverlapIndex And 4&) = 4& Then
        recBlock.Left = recNewArea.Left + recNewArea.Width
        recBlock.Right = recOldArea.Left + recOldArea.Width
        recBlock.Top = recNewArea.Top
        recBlock.Bottom = recNewArea.Top + recNewArea.Height

        FillRect mlConsoleHDC, recBlock, mlGreyBrush
    End If


    'SE
    If (lOverlapIndex And 12&) = 12& Then
        recBlock.Left = recNewArea.Left + recNewArea.Width
        recBlock.Right = recOldArea.Left + recOldArea.Width
        recBlock.Top = recNewArea.Top + recNewArea.Height
        recBlock.Bottom = recOldArea.Top + recOldArea.Height

        FillRect mlConsoleHDC, recBlock, mlGreyBrush
    End If

    'S
    If (lOverlapIndex And 8&) = 8& Then
        recBlock.Left = recNewArea.Left
        recBlock.Right = recNewArea.Left + recNewArea.Width
        recBlock.Top = recNewArea.Top + recNewArea.Height
        recBlock.Bottom = recOldArea.Top + recOldArea.Height

        FillRect mlConsoleHDC, recBlock, mlGreyBrush
    End If

    'SW
    If (lOverlapIndex And 9&) = 9& Then
        recBlock.Left = recOldArea.Left
        recBlock.Right = recNewArea.Left
        recBlock.Top = recNewArea.Top + recNewArea.Height
        recBlock.Bottom = recOldArea.Top + recOldArea.Height

        FillRect mlConsoleHDC, recBlock, mlGreyBrush
    End If

    'W
    If (lOverlapIndex And 1&) = 1& Then
        recBlock.Left = recOldArea.Left
        recBlock.Right = recNewArea.Left
        recBlock.Top = recNewArea.Top
        recBlock.Bottom = recNewArea.Top + recNewArea.Height

        FillRect mlConsoleHDC, recBlock, mlGreyBrush
    End If
End Sub
