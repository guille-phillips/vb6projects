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
   Begin VB.CommandButton cmdTransferROM 
      Caption         =   "Transfer ROM"
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
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub StartTransfer(ByVal sFile As String)
    Dim dTime As Date
    
    InitialiseComPort
    TransferBasicProgram
    
    dTime = Time
    While Time < TimeSerial(Hour(dTime), Minute(dTime), Second(dTime) + 2)
        DoEvents
    Wend
    InitialiseFastComPort
    TransferROM sFile
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

Private Sub TransferROM(ByVal sPath As String)
    Dim yData() As Byte
    
    ReDim yData(FileLen(sPath))
    Open sPath For Binary As #1
    Get #1, , yData
    Close #1
    
    SendByte &HFE30&, 0 ' Select Sideways RAM
    TransferMemory yData
    SendByte &HFE30&, 15 ' Select Basic
    
    SendWord 0 ' Dummy start address
    SendWord 0 ' Dummy end address : Load registers and jump
End Sub

Private Sub TransferMemory(yData() As Byte)
    Dim lIndex As Long
    Dim sMemString As String
    Dim yMem() As Byte
    Dim lTotal As Long
    Dim lBlock As Long
    Dim lBlockSize As Long
    Dim lBlockEnd As Long
    Dim lEndAddress As Long
    
    Dim lSlow As Long
    
    lBlockSize = 2000
    ReDim yMem(lBlockSize - 1)
    
    lEndAddress = &HBFFF&
    For lBlock = &H8000& To lEndAddress Step lBlockSize
        If (lBlock + lBlockSize - 1) <= lEndAddress Then
            For lIndex = 0 To lBlockSize - 1
                yMem(lIndex) = yData(lBlock + lIndex - &H8000&)
            Next
            SendData lBlock, lBlockSize, yMem
        Else
            ReDim yMem(lEndAddress - lBlock)
            For lIndex = 0 To lEndAddress - lBlock
                yMem(lIndex) = yData(lBlock + lIndex - &H8000&)
            Next
            SendData lBlock, lEndAddress - lBlock + 1, yMem
        End If
    Next
End Sub

Private Sub TransferRomSelect()
    UEFHandler.ResetUEF
    If UEFHandler.FindBlock(&H461&) Then

    End If
End Sub

Private Sub cmdTransferROM_Click()
    ComDlg.Filter = "*.rom"
    ComDlg.ShowOpen
    If ComDlg.Flags And 1024 = 1024 Then
        StartTransfer ComDlg.FileName
    End If
End Sub

