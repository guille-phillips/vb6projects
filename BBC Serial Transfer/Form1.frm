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
      Text            =   "7780"
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
    'Exit Sub
    
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
    yOut(1) = (lWord And &HFF00&) \ 256

    Com.Output = yOut
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
    
    sFile = oFSO.OpenTextFile(App.path & "\ReadSerialBBC.txt").ReadAll
    vSplit = Split(sFile, vbCrLf)
    lLineNumber = 10
    For Each vLine In vSplit
        vLine = Replace$(vLine, "&7F00", "&" & Hex$(mlBaseAddress))
        Com.Output = lLineNumber & vLine & vbCr
        lLineNumber = lLineNumber + 10
        
        For lSlow = 0 To 10000
            DoEvents
        Next
    Next
    Com.Output = "RUN" & vbCr
    'Com.Output = "*FX2" & vbCr
    Com.PortOpen = False
End Sub

Private Sub TransferSnapshot(ByVal sPath As String)

    UEFHandler.LoadUEFFile sPath
    
    TransferMemory
    TransferRomSelect
    TransferVideo
    TransferRegisters
    DisplayStack
    TransferUserVIA
    TransferSystemVIA

    SendWord 0 ' Dummy start address
    SendWord 0 ' Dummy end address : Load registers and jump
End Sub

Private Sub TransferMemory()
    Dim lIndex As Long
    Dim sMemString As String
    Dim yMem() As Byte
    Dim lTotal As Long
    Dim lBlock As Long
    Dim lBlockSize As Long
    Dim lBlockEnd As Long
    Dim lEndAddress As Long
    
    Dim lSlow As Long
    
    UEFHandler.ResetUEF
    If Not UEFHandler.FindBlock(&H462&) Then
        Exit Sub
    End If
    
    lBlockSize = 2000
    ReDim yMem(lBlockSize - 1)
    
    lEndAddress = mlBaseAddress - 1
    For lBlock = &H0& To lEndAddress Step lBlockSize
        If (lBlock + lBlockSize - 1) <= lEndAddress Then
            For lIndex = 0 To lBlockSize - 1
                yMem(lIndex) = UEFHandler.BlockData(lBlock + lIndex)
            Next
            SendData lBlock, lBlockSize, yMem
        Else
            ReDim yMem(lEndAddress - lBlock)
            For lIndex = 0 To lEndAddress - lBlock
                yMem(lIndex) = UEFHandler.BlockData(lBlock + lIndex)
            Next
            SendData lBlock, lEndAddress - lBlock + 1, yMem
        End If
    Next
    
    lBlockSize = 2000
    ReDim yMem(lBlockSize - 1)
    
    lEndAddress = &H7FFF&
    For lBlock = mlBaseAddress + 150 To lEndAddress Step lBlockSize
        If (lBlock + lBlockSize - 1) <= lEndAddress Then
            For lIndex = 0 To lBlockSize - 1
                yMem(lIndex) = UEFHandler.BlockData(lBlock + lIndex)
            Next
            SendData lBlock, lBlockSize, yMem
        Else
            ReDim yMem(lEndAddress - lBlock)
            For lIndex = 0 To lEndAddress - lBlock
                yMem(lIndex) = UEFHandler.BlockData(lBlock + lIndex)
            Next
            SendData lBlock, lEndAddress - lBlock + 1, yMem
        End If
    Next
    

'    ReDim yMem(&H7FFF&)
'
'    For lIndex = 0 To &H7FFF&
'        yMem(lIndex) = yFile(lIndex + &H5C&)
'    Next
'
'    On Error Resume Next
'    Kill App.path & "\memory.rom"
'    Open App.path & "\memory.rom" For Binary As #1
'    Put #1, , yMem
'    Close #1
End Sub

Private Sub TransferRomSelect()
    UEFHandler.ResetUEF
    If UEFHandler.FindBlock(&H461&) Then
        SendByte &HFE30&, UEFHandler.BlockData(0)
    End If
End Sub

Private Sub TransferVideo()
    Dim yReg(0) As Byte
    Dim lRegister As Long
    Dim lPalletteIndex As Long
    Dim vPalletteOrder As Variant
    
    UEFHandler.ResetUEF
    If Not UEFHandler.FindBlock(&H468&) Then
        Exit Sub
    End If
    
    For lRegister = 0 To 17
        SendByte &HFE00&, lRegister
        SendByte &HFE01&, UEFHandler.BlockData(lRegister)
    Next
    
    SendByte &HFE20&, UEFHandler.BlockData(18)
        
    For lPalletteIndex = 0 To 15
        SendByte &HFE21&, UEFHandler.BlockData(19 + lPalletteIndex) + lPalletteIndex * 16
    Next
End Sub

Private Sub TransferRegisters()
    Dim lC As Long
    Dim lZ As Long
    Dim lInt As Long
    Dim lD As Long
    Dim lB As Long
    Dim lV As Long
    Dim lN As Long
    Dim lP As Long
    
    UEFHandler.ResetUEF
    If Not UEFHandler.FindBlock(&H460&) Then
        Exit Sub
    End If
    
    SendByte mlBaseAddress + 1&, UEFHandler.BlockData(5)  ' S Reg
    SendByte mlBaseAddress + 4&, UEFHandler.BlockData(6)  ' P
    SendByte mlBaseAddress + 7&, UEFHandler.BlockData(3) ' X Reg
    SendByte mlBaseAddress + 9&, UEFHandler.BlockData(4) ' Y Reg
    SendByte mlBaseAddress + 11&, UEFHandler.BlockData(2) ' A Reg
    SendByte mlBaseAddress + 14&, UEFHandler.BlockData(0)  ' PC Reg Lo
    SendByte mlBaseAddress + 15&, UEFHandler.BlockData(1)  ' PC Reg Hi
 
    txtA.Text = Hex2(UEFHandler.BlockData(2))
    txtX.Text = Hex2(UEFHandler.BlockData(3))
    txtY.Text = Hex2(UEFHandler.BlockData(4))
    txtS.Text = Hex2(UEFHandler.BlockData(5))
    txtP.Text = Hex2(UEFHandler.BlockData(6))
    txtPC.Text = Hex4(UEFHandler.BlockData(0) + UEFHandler.BlockData(1) * 256&)
End Sub

Private Sub TransferSystemVIA()
    Dim lByte As Long
    Dim bBlockFound As Boolean
    
    UEFHandler.ResetUEF
    Do
        If UEFHandler.FindBlock(&H467&) Then
            If BlockData(0) = 0 Then
                bBlockFound = True
                Exit Do
            End If
        Else
            Exit Do
        End If
    Loop
    
    If Not bBlockFound Then
        Exit Sub
    End If

    ' IC 32 state
    Dim lValue As Long
    
    SendByte &HFE42&, &HFF&
    lValue = UEFHandler.BlockData(21)
    For lByte = 0 To 7
        SendByte &HFE40&, lByte + Sgn(lValue And 2 ^ lByte) * 8
    Next

    SendByte &HFE4B&, UEFHandler.BlockData(15) ' ACR
    SendByte &HFE4C&, UEFHandler.BlockData(16) ' PCR
    
    SendByte &HFE40&, UEFHandler.BlockData(1) ' ORB
    SendByte &HFE41&, UEFHandler.BlockData(3) ' ORA
    SendByte &HFE42&, UEFHandler.BlockData(4) ' DDRB
    SendByte &HFE43&, UEFHandler.BlockData(6) ' DDRA
    
    SendByte &HFE44&, UEFHandler.BlockData(7) ' T1-L
    SendByte &HFE45&, UEFHandler.BlockData(8) ' T1-H

    SendByte &HFE48&, UEFHandler.BlockData(11) ' T2-L
    SendByte &HFE49&, UEFHandler.BlockData(12) ' T2-H

    SendByte &HFE4D&, &H7F ' IFR Clear all bits
    SendByte &HFE4D&, UEFHandler.BlockData(17) Or &H80 ' IFR

    SendByte &HFE4E&, &H7F ' IER Clear all bits
    SendByte &HFE4E&, UEFHandler.BlockData(18) Or &H80 ' IER set bits
End Sub

Private Sub TransferUserVIA()
    Dim bBlockFound As Boolean
    
    UEFHandler.ResetUEF
    Do
        If UEFHandler.FindBlock(&H467&) Then
            If BlockData(0) = 1 Then
                bBlockFound = True
                Exit Do
            End If
        Else
            Exit Do
        End If
    Loop
    
    If Not bBlockFound Then
        Exit Sub
    End If

    SendByte &HFE6B&, UEFHandler.BlockData(15)  ' ACR
    SendByte &HFE6C&, UEFHandler.BlockData(16) ' PCR
    
    SendByte &HFE60&, UEFHandler.BlockData(1) ' ORB
    SendByte &HFE61&, UEFHandler.BlockData(3)  ' ORA
    SendByte &HFE62&, UEFHandler.BlockData(5) ' DDRB
    SendByte &HFE63&, UEFHandler.BlockData(6)  ' DDRA
    
    SendByte &HFE64&, UEFHandler.BlockData(7) ' T1-L
    SendByte &HFE65&, UEFHandler.BlockData(8)  ' T1-H
    
    SendByte &HFE68&, UEFHandler.BlockData(11) ' T2-L
    SendByte &HFE69&, UEFHandler.BlockData(12) ' T2-H

    SendByte &HFE6D&, &H7F ' IFR
    SendByte &HFE6D&, UEFHandler.BlockData(17) Or &H80  ' IFR
    
    SendByte &HFE6E&, &H7F ' IER Clear all bits
    SendByte &HFE6E&, UEFHandler.BlockData(18) Or &H80  ' IER set bits
End Sub

Private Sub TransferSound()

End Sub


Private Sub DisplayStack()
    Dim lMem As Long
    Dim lS As Long
    
    UEFHandler.ResetUEF
    If Not UEFHandler.FindBlock(&H460&) Then
        Exit Sub
    End If
    
    lS = UEFHandler.BlockData(5)
    
    UEFHandler.ResetUEF
    If Not UEFHandler.FindBlock(&H462&) Then
        Exit Sub
    End If
        
    txtStack.Text = ""
    For lMem = 0 To 20
        txtStack.Text = txtStack.Text & Hex4(UEFHandler.BlockData(&H100& + ((lMem + lS) And &HFF&)) + UEFHandler.BlockData(&H100 + ((lMem + lS + 1) And &HFF&)) * 256&) & vbCrLf
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
        Debug.Print Com.CommEvent & " " & Com.InBufferCount
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

