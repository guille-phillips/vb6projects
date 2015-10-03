VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRangeNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mcStarting As Currency
Private mcEnding As Currency

' meta data
Private mcOffset As Currency
Private mcConstant As Currency
Private mbSigned As Boolean
Private mcStorageRange As Currency
Private mlStorageBits As Long
Private mlStorageBytes As Long

Public Property Get Starting() As Currency
    Starting = mcStarting
End Property

Public Property Get Offset() As Currency
    Offset = mcOffset
End Property

Public Property Get Constant() As Currency
    Constant = mcStarting
End Property

Public Property Get Ending() As Currency
    Ending = mcEnding
End Property

Public Property Let Constant(ByVal cConstant As Currency)
    mcStarting = cConstant
    mcEnding = cConstant
    mcOffset = cConstant
    mcConstant = cConstant
    mcStorageRange = 0
    mlStorageBits = 0
    mlStorageBytes = 0
    mbSigned = cStarting < 0
End Property

Public Property Let Starting(ByVal cStarting As Currency)
    mcStarting = cStarting
    mcEnding = cStarting
    mcConstant = cStarting
    mcOffset = cStarting
    mcStorageRange = 0
    mlStorageBits = 0
    mlStorageBytes = 0
    mbSigned = cStarting < 0
End Property

Public Property Let Ending(ByVal cEnding As Currency)
    If cEnding < mcStarting Then
        Err.Raise -1, , "Cannot have end range before start range"
        End
    End If
    mcEnding = cEnding
    mcStorageRange = mcEnding - mcStarting
    If mcEnding < 0 Then
        mbSigned = True
    End If
    
    If mcStorageRange > 0 Then
        mlStorageBits = Int(Log(mcStorageRange) / Log(2)) + 1
        mlStorageBytes = ((mlStorageBits - 1) \ 8) + 1
    Else
        mcStorageBits = 0
        mlStorageBytes = 0
    End If
End Property

Public Property Get StorageRange() As Currency
    StorageRange = mcStorageRange
End Property

Public Property Get StorageBits() As Long
    StorageBits = mlStorageBits
End Property

Public Property Get StorageBytes() As Long
    StorageBytes = mlStorageBytes
End Property

Public Property Get Signed() As Boolean
    Signed = mbSigned
End Property

Public Property Get CalcBytes() As Long
    If Not mbSigned Then
        CalcBytes = Int(Log(mcEnding) / Log(256)) + 1
    Else
        If Abs(mcStarting) > Abs(mcEnding) Then
            CalcBytes = Int(Log(mcStarting * 2) / Log(256)) + 1
        Else
            CalcBytes = Int(Log(mcEnding * 2) / Log(256)) + 1
        End If
    End If
End Property

Private Sub Class_Initialize()
    mcStarting = 0
    mcEnding = -1
    mcStorageRange = -1
    mlStorageBits = -1
    mlStorageBytes = -1
    mlCalcRange = -1
    mlCalcBytes = -1
End Sub


