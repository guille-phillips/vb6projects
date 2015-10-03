VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "BBC Snapshot Transfer"
   ClientHeight    =   900
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3885
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   900
   ScaleWidth      =   3885
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOpenSnapshot 
      Caption         =   "Run Snapshot"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1455
   End
   Begin MSComDlg.CommonDialog ComDlg 
      Left            =   1680
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSCommLib.MSComm Com 
      Left            =   2280
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
      OutBufferSize   =   0
      InputMode       =   1
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mlBaseAddress As Long

Private Sub Form_Load()
    mlBaseAddress = &HC00&
    InitialiseComPort
    TransferBasicProgram
    InitialiseFastComPort
End Sub

Private Sub cmdOpenSnapshot_Click()
    ComDlg.Filter = "*.uef"
    ComDlg.ShowOpen
    If ComDlg.Flags And 1024 = 1024 Then
        mlBaseAddress = Val("&h" & txtLoaderAddress & "&")
        StartTransfer ComDlg.FileName
    End If
End Sub


Private Sub InitialiseComPort()
    Com.CommPort = 1
    Com.Settings = "9600,N,8,1"
    'Com.InputLen = 0
    'Com.OutBufferSize = 32767
    Com.PortOpen = True
End Sub


Private Sub InitialiseFastComPort()
    Com.CommPort = 1
    Com.Settings = "38400,N,8,1"
    'Com.InputLen = 0
    'Com.OutBufferSize = 32767
    Com.PortOpen = True
End Sub


Private Sub TransferBasicProgram()
    Dim oFSO As New FileSystemObject
    Dim sFile As String
    Dim lLineNumber As Long
    Dim vSplit As Variant
    Dim vLine As Variant
    Dim dTime As Date
    Dim lSlow As Long
    
    sFile = oFSO.OpenTextFile(App.Path & "\BBCFileSystem.txt").ReadAll
    vSplit = Split(sFile, vbCrLf)
    lLineNumber = 10
    For Each vLine In vSplit
        vLine = Replace$(vLine, "&7F00", "&" & Hex$(mlBaseAddress))
        Com.Output = lLineNumber & vLine & vbCr
        lLineNumber = lLineNumber + 10
        
        For lSlow = 0 To 20000
            DoEvents
        Next
    Next
    Com.Output = "RUN" & vbCr
    Com.PortOpen = False
End Sub


Private Function Hex2(ByVal lValue As Long)
    Hex2 = Hex$(lValue)
    Hex2 = String$(2 - Len(Hex2), "0") & Hex2
End Function

Private Function Hex4(ByVal lValue As Long)
    Hex4 = Hex$(lValue)
    Hex4 = String$(4 - Len(Hex4), "0") & Hex4
End Function

Private Sub Com_OnComm()
    Dim yBuffer() As Byte
    Dim lBufferCount As Long
    
    
    Debug.Print Com.CommEvent & " " & Com.InBufferCount

    Select Case Com.CommEvent
            Case comEventBreak   ' A Break was received.
                If Com.InBufferCount > 0 Then
                    ReDim yBuffer(Com.InBufferCount - 1) As Byte
                    yBuffer = Com.Input
                    Select Case yBuffer(0)
                        Case &HFF ' Load file
                            LoadFile yBuffer
                        Case 0 ' Save file
                        Case 1
                        Case 2
                        Case 3
                        Case 4
                        Case 5
                        Case 6 ' Delete file
                    End Select
                End If
            Case comEventCDTO    ' CD (RLSD) Timeout.
            Case comEventCTSTO   ' CTS Timeout.
            Case comEventDSRTO   ' DSR Timeout.
            Case comEventFrame   ' Framing Error.
            Case comEventOverrun ' Data Lost.
            Case comEventRxOver  ' Receive buffer overflow.
            Case comEventRxParity   ' Parity Error.
            Case comEventTxFull  ' Transmit buffer full.
            Case comEventDCB     ' Unexpected error retrieving DCB]

         ' Events
            Case comEvCD   ' Change in the CD line.

            Case comEvCTS  ' Change in the CTS line.
            Case comEvDSR  ' Change in the DSR line.
            Case comEvRing ' Change in the Ring Indicator.
            Case comEvSend ' There are SThreshold number of
                           ' characters in the transmit buffer.
            Case comEvEOF  ' An EOF character was found in the
                           ' input stream.
    
            Case comEvReceive ' Received RThreshold # of chars.

    End Select
End Sub
' 0|00|101|01
' 0|10|101|11
Private Function StripZero(yBuffer() As Byte)
    Dim lPos As Long
    
    For lPos = 0 To UBound(yBuffer) Step 2
        yBuffer(lPos \ 2) = yBuffer(lPos)
    Next
    ReDim Preserve yBuffer((UBound(yBuffer) + 1) \ 2 - 1)
End Function

Private Sub LoadFile(yInfo() As Byte)
    Dim sFilepath As String
    Dim lPos As Long
    Dim lLoadAddress As Long
    Dim lExecutionAddress As Long
    Dim lLength As Long
    Dim lAttributes As Long
    Dim lFileLength As Long
    Dim yFile() As Byte

    lPos = 1
    While yInfo(lPos) <> 13
        sFilepath = sFilepath & Chr$(yInfo(lPos))
        lPos = lPos + 1
    Wend
    lPos = lPos + 1
    
    lLoadAddress = GetWord(yInfo, lPos)
    lExecutionAddress = GetWord(yInfo, lPos + 4)
    lLength = GetWord(yInfo, lPos + 8)
    lAttributes = yInfo(lPos + 12)
    
    lFileLength = FileLen(sFilepath)
    ReDim yFile(lFileLength - 1)
    Open sFilepath For Binary As #1
    Get #1, , yFile
    Close #1
    Delay
    SendData lLoadAddress, lFileLength, yFile
End Sub

Private Function GetWord(yInfo() As Byte, lOffset As Long)
    Dim lIndex As Long
    Dim lMultiplier As Long
    
    lMultiplier = 1
    For lIndex = 0 To 1
        GetWord = GetWord + yInfo(lOffset + lIndex) * lMultiplier
        lMultiplier = lMultiplier * 256
    Next
End Function

Private Sub Delay()
    Dim lDelay As Long
    For lDelay = 1 To 100000#
        DoEvents
    Next
End Sub

Private Sub SendWord(ByVal lWord As Long)
    Dim yOut(1) As Byte
    
    yOut(0) = lWord And &HFF
    yOut(1) = lWord \ 256

    Com.Output = yOut
End Sub

Private Sub SendBlockDetails(ByVal lAddress As Long, ByVal lLength As Long)
    SendWord lAddress - ((65536 - lLength) And &HFF&)
    SendWord (65536 - lLength) And &HFFFF&
End Sub

Private Sub SendByte(ByVal lAddress, ByVal yValue As Long)
    Dim yByte(0) As Byte
    
    SendBlockDetails lAddress, 1
    yByte(0) = yValue
    Com.Output = yByte
End Sub

Private Sub SendData(ByVal lAddress As Long, ByVal lLength As Long, yData() As Byte)
    SendBlockDetails lAddress, lLength
    Com.Output = yData
End Sub
