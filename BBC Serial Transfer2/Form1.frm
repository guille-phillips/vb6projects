VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "BBC Snapshot Transfer"
   ClientHeight    =   5940
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3885
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5940
   ScaleWidth      =   3885
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtA 
      Height          =   285
      Left            =   1440
      TabIndex        =   9
      Top             =   1200
      Width           =   615
   End
   Begin VB.TextBox txtY 
      Height          =   285
      Left            =   1440
      TabIndex        =   8
      Top             =   1920
      Width           =   615
   End
   Begin VB.TextBox txtS 
      Height          =   285
      Left            =   1440
      TabIndex        =   7
      Top             =   2280
      Width           =   615
   End
   Begin VB.TextBox txtP 
      Height          =   285
      Left            =   1440
      TabIndex        =   6
      Top             =   2640
      Width           =   615
   End
   Begin VB.TextBox txtPC 
      Height          =   285
      Left            =   1440
      TabIndex        =   5
      Top             =   3000
      Width           =   615
   End
   Begin VB.TextBox txtX 
      Height          =   285
      Left            =   1440
      TabIndex        =   4
      Top             =   1560
      Width           =   615
   End
   Begin VB.TextBox txtStack 
      Height          =   5295
      Left            =   2400
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   480
      Width           =   1335
   End
   Begin VB.TextBox txtLoaderAddress 
      Height          =   285
      Left            =   1440
      TabIndex        =   2
      Text            =   "7C00"
      Top             =   840
      Width           =   615
   End
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
      OutBufferSize   =   32767
   End
   Begin VB.Label Label7 
      Caption         =   "A"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "X"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "Y"
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   1920
      Width           =   1215
   End
   Begin VB.Label Label4 
      Caption         =   "S"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "P"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Label Label6 
      Caption         =   "PC"
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Loader Address"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mlBaseAddress As Long

Private Sub cmdOpenSnapshot_Click()
    ComDlg.Filter = "*.uef"
    ComDlg.ShowOpen
    If ComDlg.Flags And 1024 = 1024 Then
        mlBaseAddress = Val("&h" & txtLoaderAddress & "&")
        StartTransfer ComDlg.FileName
    End If
End Sub


Private Sub StartTransfer(ByVal sFile As String)
    Dim dTime As Date
    
    InitialiseComPort
    TransferBasicProgram
    dTime = Time
    While Time < TimeSerial(Hour(dTime), Minute(dTime), Second(dTime) + 2)
        DoEvents
    Wend
    InitialiseFastComPort
    TransferSnapshot sFile
    Com.PortOpen = False
End Sub

Private Sub InitialiseComPort()
    Com.CommPort = 1
    Com.Settings = "9600,N,8,1"
    Com.InputLen = 0
    Com.OutBufferSize = 32767
    Com.PortOpen = True
End Sub


Private Sub InitialiseFastComPort()
    Com.CommPort = 1
    Com.Settings = "38400,N,8,1"
    Com.InputLen = 0
    Com.OutBufferSize = 32767
    Com.PortOpen = True
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

Private Sub TransferBasicProgram()
    Dim oFSO As New FileSystemObject
    Dim sFile As String
    Dim lLineNumber As Long
    Dim vSplit As Variant
    Dim vLine As Variant
    Dim dTime As Date
    Dim lSlow As Long
    
    sFile = oFSO.OpenTextFile(App.Path & "\ReadSerialBBC.txt").ReadAll
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
    Com.Output = "*FX2,1" & vbCr
    For lSlow = 0 To 20000
        DoEvents
    Next
    Com.Output = "RUN" & vbCr
    Com.PortOpen = False
End Sub

Private Sub TransferSnapshot(ByVal sPath As String)
    Dim yFile() As Byte
    
    ReDim yFile(FileLen(sPath) - 1)
    Open sPath For Binary As #1
    Get #1, , yFile
    Close #1
    
    DisplayStack yFile
    TransferRegisters yFile
    TransferMemory yFile
    'TransferUserVIA yFile
    'TransferSystemVIA yFile
    TransferVideo yFile
    SendBlockDetails 0, 0  ' Load registers and jump
End Sub

Private Sub TransferMemory(yFile() As Byte)
    Dim lIndex As Long
    Dim yMem() As Byte

    'SendByte &HFE30&, yFile(&H51&) 'paged rom
    
    TransferMemoryBlock yFile, &H0&, mlBaseAddress - 1
    TransferMemoryBlock yFile, mlBaseAddress + 256, &H7FFF&
    
    ReDim yMem(&H7FFF&)
    
    For lIndex = 0 To &H7FFF&
        yMem(lIndex) = yFile(lIndex + &H5C&)
    Next
    
    On Error Resume Next
    Kill App.Path & "\memory.rom"
    Open App.Path & "\memory.rom" For Binary As #1
    Put #1, , yMem
    Close #1
End Sub

Private Sub TransferMemoryBlock(yFile() As Byte, ByVal lStart As Long, ByVal lEnd As Long)
    Dim lBlockSize As Long
    Dim yMem() As Byte
    Dim lEndAddress
    Dim lBlock As Long
    Dim lIndex As Long
    
    lBlockSize = 4800
    ReDim yMem(lBlockSize - 1)
    
    For lBlock = lStart To lEnd Step lBlockSize
        If (lBlock + lBlockSize - 1) <= lEnd Then
            For lIndex = 0 To lBlockSize - 1
                yMem(lIndex) = yFile(lBlock + lIndex + &H5C&)
            Next
            'Debug.Print Hex4(lBlock + lIndex - 1)
            SendData lBlock, lBlockSize, yMem
        Else
            ReDim yMem(lEnd - lBlock)
            For lIndex = 0 To lEnd - lBlock
                yMem(lIndex) = yFile(lBlock + lIndex + &H5C&)
            Next
            SendData lBlock, lEnd - lBlock + 1, yMem
            'Debug.Print Hex4(lBlock + lIndex - 1)
        End If
    Next
End Sub

Private Sub TransferVideo(yFile() As Byte)
    Dim yReg(0) As Byte
    Dim lRegister As Long
    Dim lPalletteIndex As Long
    Dim vPalletteOrder As Variant
    
    For lRegister = 0 To 15
        SendByte &HFE00&, lRegister
        SendByte &HFE01&, yFile(&H1807E + lRegister)
    Next
    
    SendByte &HFE20&, yFile(&H18090)
        
    For lPalletteIndex = 0 To 15
        SendByte &HFE21&, yFile(&H18091 + lPalletteIndex) + lPalletteIndex * 16
    Next
End Sub

Private Sub TransferRegisters(yFile() As Byte)
    Dim lC As Long
    Dim lZ As Long
    Dim lInt As Long
    Dim lD As Long
    Dim lB As Long
    Dim lV As Long
    Dim lN As Long
    Dim lP As Long
    
    
    txtA.Text = Hex2(yFile(&H40&))
    txtX.Text = Hex2(yFile(&H41&))
    txtY.Text = Hex2(yFile(&H42&))
    txtS.Text = Hex2(yFile(&H43&))
    txtP.Text = Hex2(yFile(&H44&))
    txtPC.Text = Hex4(CLng(yFile(&H3E&)) + CLng(yFile(&H3F&)) * 256)
    

    lP = yFile(&H44&)
    SendByte mlBaseAddress + &H0&, lP And &HC0& ' P: NV flags
    SendByte mlBaseAddress + &H2&, yFile(&H43&)  ' S Reg
    SendByte mlBaseAddress + &H5&, yFile(&H41&)  ' X Reg
    SendByte mlBaseAddress + &H7&, yFile(&H42&)  ' Y Reg
    SendByte mlBaseAddress + &H9&, yFile(&H40&)  ' A Reg
    
    lC = lP And 1&
    lZ = lP And 2&
    lInt = lP And 4&
    lD = lP And 8&
    lB = lP And 16&

'    If lC = 0 Then
'        SendByte mlBaseAddress + &HD&, &H18&
'    Else
'        SendByte mlBaseAddress + &HD&, &H38&
'    End If
'
'    If lD = 0 Then
'        SendByte mlBaseAddress + &HE&, &HD8&
'    Else
'        SendByte mlBaseAddress + &HE&, &HF8&
'    End If
'
'    lInt = 0
'    If lInt = 0 Then
'        SendByte mlBaseAddress + &HF&, &H58&
'    Else
'        SendByte mlBaseAddress + &HF&, &H78&
'    End If

    SendByte mlBaseAddress + &H11&, yFile(&H3E&)  ' PC Reg Lo
    SendByte mlBaseAddress + &H12&, yFile(&H3F&)  ' PC Reg Hi
End Sub

Private Sub TransferSystemVIA(yFile() As Byte)
    Dim lByte As Long
    Dim lValue As Long
    
    SendByte &HFE42&, yFile(&H180B8) ' DDRB
    SendByte &HFE43&, yFile(&H180B9) ' DDRA
    
    ' IC 32 state
    lValue = yFile(&H180C8)
    For lByte = 4 To 5
        SendByte &HFE40&, lByte + Sgn(lValue And 2 ^ lByte) * 8
    Next

    SendByte &HFE40&, yFile(&H180B4) ' ORB
    SendByte &HFE41&, yFile(&H180B6) ' ORA

    SendByte &HFE46&, &HFF ' T1-LL
    SendByte &HFE47&, &HFF ' T1-HL
    SendByte &HFE45&, &HFF ' T1-HL


    
    SendByte &HFE4E&, &H7F ' IER Clear all bits
    SendByte &HFE4E&, yFile(&H180C5) Or &H80  ' IER set bits
    
    SendByte &HFE4D&, yFile(&H180C4) ' IFR
    SendByte &HFE4D&, 0 ' IFR
    
    SendByte &HFE4B&, yFile(&H180C2) ' ACR
    SendByte &HFE4C&, yFile(&H180C3) ' PCR

'    SendByte &HFE44&, yFile(&H180BA) ' T1-LC
'    SendByte &HFE45&, yFile(&H180BB) ' T1-HC
'
'    SendByte &HFE46&, yFile(&H180BC) ' T1-LL

    SendByte &HFE48&, yFile(&H180BE) ' T2-LC
    SendByte &HFE49&, yFile(&H180BF) ' T2-HC
    
    SendByte &HFE48&, &HFF ' T2-LC
    SendByte &HFE49&, &HFF ' T2-HC
    
    SendByte &HFE46&, yFile(&H180BC) ' T1-LL
    SendByte &HFE47&, yFile(&H180BD) ' T1-HL
    SendByte &HFE45&, yFile(&H180BD) ' T1-HL
End Sub

Private Sub TransferUserVIA(yFile() As Byte)

    SendByte &HFE6E&, &H7F ' IER Clear all bits
    'SendByte &HFE6E&, yFile(&H180E1) Or &H80  ' IER set bits
    
    SendByte &HFE6D&, yFile(&H180C4) ' IFR
    SendByte &HFE6D&, 0 ' IFR
    
    SendByte &HFE6B&, yFile(&H180DE) ' ACR
    SendByte &HFE6C&, yFile(&H180DF) ' PCR
    
    SendByte &HFE60&, yFile(&H180D0) ' ORB
    SendByte &HFE61&, yFile(&H180D2) ' ORA
    SendByte &HFE62&, yFile(&H180B4) ' DDRB
    SendByte &HFE63&, yFile(&H180B5) ' DDRA
    
    SendByte &HFE64&, yFile(&H180D6) ' T1-L
    SendByte &HFE65&, yFile(&H180D7) ' T1-H
    
    SendByte &HFE66&, yFile(&H180D8) ' T1-L
    SendByte &HFE67&, yFile(&H180D9) ' T1-H
    
    SendByte &HFE68&, yFile(&H180DA) ' T2-L
    SendByte &HFE69&, yFile(&H180DB) ' T2-H

    
End Sub

Private Sub TransferSound(yFile() As Byte)

End Sub


Private Sub DisplayStack(yFile() As Byte)
    Dim lMem As Long
    Dim lS As Long
    
    txtStack.Text = ""
    lS = yFile(&H43&)
    For lMem = 0 To 20
        txtStack.Text = txtStack.Text & Hex4(yFile(&H5C& + ((lS + lMem) Mod 256&) + &H100&) + yFile(&H5C& + ((lS + lMem + 1) Mod 256&) + &H100&) * 256&) & vbCrLf
    Next
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
    Dim buffer() As Byte
    
    'If Com.InBufferCount > 0 Then
        'Debug.Print Com.CommEvent & " " & Com.InBufferCount
        'ReDim buffer(Com.InBufferCount) As Byte
        'buffer = Com.Input
    'End If

End Sub

Private Sub BufferIn()
'    Dim ov As Long
'    Dim nv As Long
'    Do
'        nv = Com.InBufferCount
'        If nv <> ov Then
'            Debug.Print nv
'            nv = ov
'        End If
'        ov = Com.InBufferCount
'        DoEvents
'        DoEvents
'    Loop
End Sub


Private Sub Text1_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 8
            KeyAscii = 127
    End Select
    Com.Output = Chr$(KeyAscii)
End Sub

