Attribute VB_Name = "CRTC6845"
Option Explicit

Public ScreenStart As Long

Public Columns As Long
Public TotalColumns As Long
Public Rows As Long
Public TotalRows As Long
Public ScanAdjust As Long
Public VerticalSyncPosition As Long
Public VerticalSyncWidth As Long
Public CursorAddress As Long
Public RefreshCount As Long
Public Interlace As Long
Public ScanlinesPerRow As Long
Public CursorStart As Long
Public CursorEnd As Long
Public CursorBlink As Long
Public CursorBlankingDelay As Long
Public DisplayBlankingDelay As Long
Public StartX As Long
Public StartY As Long
Public RowDisplayStart As Long
Public HorizontalSyncWidth As Long

Public LightPenCharX As Long
Public LightPenCharY As Long

Public DisplayDimensionsChanged As Boolean

Public SelectedRegister As Long

Private Type ChangePair
    Register As Long
    Value As Long
End Type

Private mcpRegisterChanged() As ChangePair
Private mlRegisterChangedIndex As Long
Public Register(17) As Byte

Public Sub Initialise6845()
    ' Debugging.WriteString "CRTC6845.Initialise6845"
    
    CRTC6845.ScanAdjust = 0
    CRTC6845.Interlace = 0
End Sub

Public Sub WriteRegister(ByVal lRegisterValue As Long)
    Dim cpEntry As ChangePair
    
    ' Debugging.WriteString "CRTC6845.WriteRegister"
    
    cpEntry.Register = SelectedRegister
    cpEntry.Value = lRegisterValue
    
    ReDim Preserve mcpRegisterChanged(mlRegisterChangedIndex)
    mcpRegisterChanged(mlRegisterChangedIndex) = cpEntry
    mlRegisterChangedIndex = mlRegisterChangedIndex + 1
End Sub

Public Sub VerticalSync()
    Dim lRegisterChangedIndex As Long
    Dim lRegisterValue As Long
    
    ' Debugging.WriteString "CRTC6845.VerticalSync"
    
    For lRegisterChangedIndex = 0 To mlRegisterChangedIndex - 1
        lRegisterValue = mcpRegisterChanged(lRegisterChangedIndex).Value
            
        Select Case mcpRegisterChanged(lRegisterChangedIndex).Register
            Case 0 ' h total
                'Debug.Print mcpRegisterChanged(lRegisterChangedIndex).Register & ":" & lRegisterValue
                CRTC6845.TotalColumns = lRegisterValue
            Case 1 ' h displayed
                CRTC6845.Columns = lRegisterValue
                VideoULA.CharacterRowTotalBytes = 8& * lRegisterValue
                VideoULA.InitialiseBitmap
                DisplayDimensionsChanged = True
            Case 2 ' sync pos
                CRTC6845.StartX = lRegisterValue
                DisplayDimensionsChanged = True
            Case 3 ' sync width
                CRTC6845.HorizontalSyncWidth = lRegisterValue And &HF&
                CRTC6845.VerticalSyncWidth = lRegisterValue \ 16&
                'Debug.Print mcpRegisterChanged(lRegisterChangedIndex).Register & ":" & lRegisterValue
            Case 4 ' v total
                lRegisterValue = lRegisterValue And &H7F&
                CRTC6845.StartY = lRegisterValue
                CRTC6845.TotalRows = lRegisterValue
                CRTC6845.RowDisplayStart = lRegisterValue - CRTC6845.VerticalSyncPosition
                DisplayDimensionsChanged = True
                'Debug.Print mcpRegisterChanged(lRegisterChangedIndex).Register & ":" & lRegisterValue
            Case 5 ' v scanline adjust
                lRegisterValue = lRegisterValue And &H1F&
                ScanAdjust = lRegisterValue
                DisplayDimensionsChanged = True
                DisplayDimensionsChanged = True
                'Debug.Print mcpRegisterChanged(lRegisterChangedIndex).Register & ":" & lRegisterValue
            Case 6 ' v display
                lRegisterValue = lRegisterValue And &H7F&
                CRTC6845.Rows = lRegisterValue
                DisplayDimensionsChanged = True
            Case 7 ' v sync position
                lRegisterValue = lRegisterValue And &H7F&
                CRTC6845.VerticalSyncPosition = lRegisterValue
                CRTC6845.RowDisplayStart = CRTC6845.TotalRows - lRegisterValue
                DisplayDimensionsChanged = True
                'Debug.Print mcpRegisterChanged(lRegisterChangedIndex).Register & ":" & lRegisterValue
            Case 8 ' interlace mode
                lRegisterValue = lRegisterValue And &H3F&
                CRTC6845.Interlace = lRegisterValue
                CRTC6845.CursorBlankingDelay = lRegisterValue And &HC0&
                CRTC6845.DisplayBlankingDelay = lRegisterValue And &H30&
            Case 9 ' max scan line
                lRegisterValue = lRegisterValue And &H1F&
                CRTC6845.ScanlinesPerRow = lRegisterValue
                VideoULA.CharacterRowDisplayedHeight = IIf(lRegisterValue <= 7, lRegisterValue + 1, 7 + 1) * 2
                VideoULA.CharacterRowHeight = lRegisterValue * 2 + 2
                DisplayDimensionsChanged = True
                'Debug.Print mcpRegisterChanged(lRegisterChangedIndex).Register & ":" & lRegisterValue
            Case 10 ' cursor start
                CursorBlink = lRegisterValue And &H40&
                If (lRegisterValue And &H20&) Then
                    VideoULA.CursorBlinkFieldRate = 16
                Else
                    VideoULA.CursorBlinkFieldRate = 8
                End If
                CursorStart = lRegisterValue And &H1F&
            Case 11 ' cursor end
                CursorEnd = lRegisterValue And &H1F&
            Case 12 ' start hi
                ScreenStart = (ScreenStart And &HFF&) Or lRegisterValue * 256&
                VideoULA.ScreenBaseAddress = (ScreenStart * 8) And &HFFFF&
                VideoULA.ScreenTeletextBaseAddress = &H7400& + (CRTC6845.ScreenStart Xor &H2000&)
            Case 13 ' start lo
                ScreenStart = (ScreenStart And &HFF00&) Or lRegisterValue
                VideoULA.ScreenBaseAddress = (ScreenStart * 8) And &HFFFF&
                VideoULA.ScreenTeletextBaseAddress = &H7400& + (CRTC6845.ScreenStart Xor &H2000&)
            Case 14 ' cursor hi
                CRTC6845.CursorAddress = (CursorAddress And &HFF&) Or lRegisterValue * 256&
                VideoULA.CursorAddress = CursorAddress * 8
                VideoULA.CursorTeletextAddress = &H7400& + (CRTC6845.CursorAddress Xor &H2000&)
                If VideoULA.CursorAddress >= &H8000& Then
                    VideoULA.CursorAddress = VideoULA.CursorAddress - HardwareScrollBytes
                End If
                If VideoULA.CursorTeletextAddress >= &H8000& Then
                    VideoULA.CursorTeletextAddress = VideoULA.CursorTeletextAddress - &H400&
                End If
            Case 15 ' cursor lo
                CRTC6845.CursorAddress = (CursorAddress And &HFF00&) Or lRegisterValue
                VideoULA.CursorAddress = CursorAddress * 8
                VideoULA.CursorTeletextAddress = &H7400& + (CRTC6845.CursorAddress Xor &H2000&)
                If VideoULA.CursorAddress >= &H8000& Then
                    VideoULA.CursorAddress = VideoULA.CursorAddress - HardwareScrollBytes
                End If
                If VideoULA.CursorTeletextAddress >= &H8000& Then
                    VideoULA.CursorTeletextAddress = VideoULA.CursorTeletextAddress - &H400&
                End If
            Case 16 ' Light pen hi
            Case 17 ' Light pen lo
        End Select
        If mcpRegisterChanged(lRegisterChangedIndex).Register <= 17 Then
            Register(mcpRegisterChanged(lRegisterChangedIndex).Register) = lRegisterValue
        End If
    Next
    
    Erase mcpRegisterChanged
    mlRegisterChangedIndex = 0
End Sub

Public Sub UpdateLightPenRegisters(ByVal lAddress As Long)
    If CRTC6845.LightPenCharX >= 0 And CRTC6845.LightPenCharX < CRTC6845.Columns And CRTC6845.LightPenCharY >= 0 And CRTC6845.LightPenCharY < CRTC6845.Rows Then
        lAddress = lAddress + CRTC6845.LightPenCharX
        Register(16) = lAddress And &HFF&
        Register(17) = (lAddress \ 256&) And &HFF&
    End If
End Sub
