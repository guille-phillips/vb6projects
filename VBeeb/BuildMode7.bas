Attribute VB_Name = "BuildMode7"
Option Explicit

Private myTeletextCharsetPattern(15, 19, 255, 2, 1, 0) As Byte ' Bit/Scanline/Char/Graphic/Fore Colour/Back Colour
Private myTeletextDHCharsetPattern(15, 39, 255, 2, 1, 0) As Byte ' Bit/Scanline/Char/Graphic/Fore Colour/Back Colour


Public Sub CreateDoubleHeightTeletextFontFile()
    Dim oFSO As New FileSystemObject
    Dim oTS As TextStream
    Dim oTS2 As TextStream
    Dim sPattern As String
    
    Set oTS = oFSO.OpenTextFile(App.path & "\Mode7Charset\Mode7Charset.txt")
    Set oTS2 = oFSO.CreateTextFile(App.path & "\Mode7Charset\Mode7DHCharset.txt")
    
    While Not oTS.AtEndOfStream
        sPattern = oTS.ReadLine
        If sPattern <> "" Then
            oTS2.WriteLine sPattern
            oTS2.WriteLine sPattern
        Else
            oTS2.WriteLine ""
        End If
    Wend
    oTS.Close
    oTS2.Close
End Sub

Public Sub CreateTeletextCharsetFile()
    Dim oFSO As New FileSystemObject
    Dim oTS As TextStream
    Dim oTSDH As TextStream
    Dim sPattern As String
    Dim lScanLine As Long
    Dim lBit As Long
    Dim lChar As Long
    
    Static bLoaded As Boolean
    
    If bLoaded Then
        Exit Sub
    End If
    bLoaded = True
    
    Set oTS = oFSO.OpenTextFile(App.path & "\Mode7Charset\Mode7Charset.txt")
    
    lChar = 32
    lScanLine = 2
    While Not oTS.AtEndOfStream
        sPattern = oTS.ReadLine
        If sPattern <> "" Then
            For lBit = 0 To 9
                myTeletextCharsetPattern(lBit + 2, lScanLine, lChar, 0, 1, 0) = IIf(Mid$(sPattern, lBit + 1, 1) = "O", 1, 0)
                myTeletextCharsetPattern(lBit + 2, lScanLine, lChar + 128, 0, 1, 0) = myTeletextCharsetPattern(lBit + 2, lScanLine, lChar, 0, 1, 0)
                
                If (lChar >= 64 And lChar <= 95) Then
                    myTeletextCharsetPattern(lBit + 2, lScanLine, lChar, 1, 1, 0) = myTeletextCharsetPattern(lBit + 2, lScanLine, lChar, 0, 1, 0)
                    myTeletextCharsetPattern(lBit + 2, lScanLine, lChar, 2, 1, 0) = myTeletextCharsetPattern(lBit + 2, lScanLine, lChar, 0, 1, 0)
                    myTeletextCharsetPattern(lBit + 2, lScanLine, lChar + 128, 1, 1, 0) = myTeletextCharsetPattern(lBit + 2, lScanLine, lChar, 0, 1, 0)
                    myTeletextCharsetPattern(lBit + 2, lScanLine, lChar + 128, 2, 1, 0) = myTeletextCharsetPattern(lBit + 2, lScanLine, lChar, 0, 1, 0)
                End If
            Next
            lScanLine = lScanLine + 1
            If lScanLine = 20 Then
                lChar = lChar + 1
                lScanLine = 2
            End If
        End If
    Wend
    oTS.Close
    
    Dim lGraphicIndex As Long
    Dim lBitValue As Long
    Dim lMask As Long
    Dim lBitColumn As Long
    Dim lBitRow As Long
    Dim lOffset As Long
    Dim lRow As Long
    Dim lColumn As Long
    
    Dim lBitRowStart(3) As Long
    lBitRowStart(0) = 0
    lBitRowStart(1) = 6
    lBitRowStart(2) = 14
    lBitRowStart(3) = 20

    For lGraphicIndex = 0 To 63
        lMask = 1
        For lRow = 0 To 2
            For lColumn = 0 To 1
                lBitValue = -1 * ((lGraphicIndex And lMask) <> 0)
                            
                For lBitRow = lBitRowStart(lRow) To lBitRowStart(lRow + 1) - 1
                    For lBitColumn = lColumn * 6 To lColumn * 6 + 5
                        myTeletextCharsetPattern(lBitColumn, lBitRow, lGraphicIndex + 32 + lOffset, 1, 1, 0) = lBitValue
                        myTeletextCharsetPattern(lBitColumn, lBitRow, lGraphicIndex + 160 + lOffset, 1, 1, 0) = lBitValue
                    Next
                Next
                
                For lBitRow = lBitRowStart(lRow) To lBitRowStart(lRow + 1) - 3
                    For lBitColumn = lColumn * 6 + 2 To lColumn * 6 + 5
                        myTeletextCharsetPattern(lBitColumn, lBitRow, lGraphicIndex + 32 + lOffset, 2, 1, 0) = lBitValue
                        myTeletextCharsetPattern(lBitColumn, lBitRow, lGraphicIndex + 160 + lOffset, 2, 1, 0) = lBitValue
                    Next
                Next
                
                lMask = lMask * 2
            Next
        Next
        If lGraphicIndex >= 31 Then
            lOffset = 32
        End If
    Next
    
    Dim yScanLine(15&) As Byte
    Dim lSubBit As Long
    
    For lGraphicIndex = 0 To 2
        For lChar = 0 To 255
            For lScanLine = 0 To 19
                lSubBit = 0
                For lBit = 0 To 11 Step 3
                    yScanLine(lSubBit * 4 + 0) = ColourIndex(myTeletextCharsetPattern(lBit, lScanLine, lChar, lGraphicIndex, 1, 0), myTeletextCharsetPattern(lBit, lScanLine, lChar, lGraphicIndex, 1, 0))
                    yScanLine(lSubBit * 4 + 1) = ColourIndex(myTeletextCharsetPattern(lBit, lScanLine, lChar, lGraphicIndex, 1, 0), myTeletextCharsetPattern(lBit + 1, lScanLine, lChar, lGraphicIndex, 1, 0))
                    yScanLine(lSubBit * 4 + 2) = ColourIndex(myTeletextCharsetPattern(lBit + 2, lScanLine, lChar, lGraphicIndex, 1, 0), myTeletextCharsetPattern(lBit + 1, lScanLine, lChar, lGraphicIndex, 1, 0))
                    yScanLine(lSubBit * 4 + 3) = ColourIndex(myTeletextCharsetPattern(lBit + 2, lScanLine, lChar, lGraphicIndex, 1, 0), myTeletextCharsetPattern(lBit + 2, lScanLine, lChar, lGraphicIndex, 1, 0))
                
                    lSubBit = lSubBit + 1
                Next
                CopyMemory myTeletextCharsetPattern(0, lScanLine, lChar, lGraphicIndex, 1, 0), yScanLine(0), 16&
            Next
        Next
    Next
    
    Kill App.path & "\mode7charset.dat"
    Open App.path & "\mode7charset.dat" For Binary As #1
    Put #1, , myTeletextCharsetPattern
    Close #1
End Sub

Private Function ColourIndex(ByVal yColour1 As Byte, ByVal yColour2 As Byte) As Byte
    ColourIndex = 64 + yColour1 + yColour2 * 8
End Function

Public Sub CreateTeletextDoubleHeightCharsetFile()
    Dim oFSO As New FileSystemObject
    Dim oTS As TextStream
    Dim sPattern As String
    Dim lScanLine As Long
    Dim lBit As Long
    Dim lChar As Long
    
    Static bLoaded As Boolean
    
    If bLoaded Then
        Exit Sub
    End If
    bLoaded = True
    
    Set oTS = oFSO.OpenTextFile(App.path & "\Mode7Charset\Mode7DHCharset.txt")
    
    lChar = 32
    lScanLine = 4
    While Not oTS.AtEndOfStream
        sPattern = oTS.ReadLine
        If sPattern <> "" Then
            For lBit = 0 To 9
                myTeletextDHCharsetPattern(lBit + 2, lScanLine, lChar, 0, 1, 0) = IIf(Mid$(sPattern, lBit + 1, 1) = "O", 1, 0)
                myTeletextDHCharsetPattern(lBit + 2, lScanLine, lChar + 128, 0, 1, 0) = myTeletextDHCharsetPattern(lBit + 2, lScanLine, lChar, 0, 1, 0)
                
                If (lChar >= 64 And lChar <= 95) Then
                    myTeletextDHCharsetPattern(lBit + 2, lScanLine, lChar, 1, 1, 0) = myTeletextDHCharsetPattern(lBit + 2, lScanLine, lChar, 0, 1, 0)
                    myTeletextDHCharsetPattern(lBit + 2, lScanLine, lChar, 2, 1, 0) = myTeletextDHCharsetPattern(lBit + 2, lScanLine, lChar, 0, 1, 0)
                    myTeletextDHCharsetPattern(lBit + 2, lScanLine, lChar + 128, 1, 1, 0) = myTeletextDHCharsetPattern(lBit + 2, lScanLine, lChar, 0, 1, 0)
                    myTeletextDHCharsetPattern(lBit + 2, lScanLine, lChar + 128, 2, 1, 0) = myTeletextDHCharsetPattern(lBit + 2, lScanLine, lChar, 0, 1, 0)
                End If
            Next
            lScanLine = lScanLine + 1
            If lScanLine = 40 Then
                lChar = lChar + 1
                lScanLine = 4
            End If
        End If
    Wend
    oTS.Close
    
    Dim lGraphicIndex As Long
    Dim lBitValue As Long
    Dim lMask As Long
    Dim lBitColumn As Long
    Dim lBitRow As Long
    Dim lOffset As Long
    Dim lRow As Long
    Dim lColumn As Long
    
    For lGraphicIndex = 0 To 63
        lMask = 1
        For lRow = 0 To 2
            For lColumn = 0 To 1
                lBitValue = -1 * ((lGraphicIndex And lMask) <> 0)
                            
                For lBitRow = lRow * 14 To lRow * 14 + 13
                    If lBitRow < 40 Then
                        For lBitColumn = lColumn * 6 To lColumn * 6 + 5
                            myTeletextDHCharsetPattern(lBitColumn, lBitRow, lGraphicIndex + 32 + lOffset, 1, 1, 0) = lBitValue
                            myTeletextDHCharsetPattern(lBitColumn, lBitRow, lGraphicIndex + 160 + lOffset, 1, 1, 0) = lBitValue
                            If lBitRow < (lRow * 14 + 12) And lBitColumn < (lColumn * 6 + 4) Then
                               myTeletextDHCharsetPattern(lBitColumn, lBitRow, lGraphicIndex + 32 + lOffset, 2, 1, 0) = lBitValue
                               myTeletextDHCharsetPattern(lBitColumn, lBitRow, lGraphicIndex + 160 + lOffset, 2, 1, 0) = lBitValue
                            End If
                        Next
                    End If
                Next
                
                lMask = lMask * 2
            Next
        Next
        If lGraphicIndex >= 31 Then
            lOffset = 32
        End If
    Next

    Dim yScanLine(15&) As Byte
    Dim lSubBit As Long
    
    For lGraphicIndex = 0 To 2
        For lChar = 0 To 255
            For lScanLine = 0 To 39
                lSubBit = 0
                For lBit = 0 To 11 Step 3
                    yScanLine(lSubBit * 4 + 0) = ColourIndex(myTeletextDHCharsetPattern(lBit, lScanLine, lChar, lGraphicIndex, 1, 0), myTeletextDHCharsetPattern(lBit, lScanLine, lChar, lGraphicIndex, 1, 0))
                    yScanLine(lSubBit * 4 + 1) = ColourIndex(myTeletextDHCharsetPattern(lBit, lScanLine, lChar, lGraphicIndex, 1, 0), myTeletextDHCharsetPattern(lBit + 1, lScanLine, lChar, lGraphicIndex, 1, 0))
                    yScanLine(lSubBit * 4 + 2) = ColourIndex(myTeletextDHCharsetPattern(lBit + 2, lScanLine, lChar, lGraphicIndex, 1, 0), myTeletextDHCharsetPattern(lBit + 1, lScanLine, lChar, lGraphicIndex, 1, 0))
                    yScanLine(lSubBit * 4 + 3) = ColourIndex(myTeletextDHCharsetPattern(lBit + 2, lScanLine, lChar, lGraphicIndex, 1, 0), myTeletextDHCharsetPattern(lBit + 2, lScanLine, lChar, lGraphicIndex, 1, 0))
                
                    lSubBit = lSubBit + 1
                Next
                CopyMemory myTeletextDHCharsetPattern(0, lScanLine, lChar, lGraphicIndex, 1, 0), yScanLine(0), 16&
            Next
        Next
    Next
    
    Kill App.path & "\mode7DHcharset.dat"
    Open App.path & "\mode7DHcharset.dat" For Binary As #1
    Put #1, , myTeletextDHCharsetPattern
    Close #1
End Sub
