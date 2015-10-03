VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRange"
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
Private mbIsConstant As Boolean
Private mbSigned As Boolean
Private mcStorageRange As Currency
Private mlStorageBits As Long
Private mlStorageBytes As Long

Public Function Create(ByVal cStarting As Currency, Optional ByVal cEnding As Currency = -0.5) As clsRange
    Set Create = New clsRange
    
    Create.Starting = cStarting
    If cEnding <> -0.5 Then
        Create.Ending = cEnding
    End If
End Function

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
    mbIsConstant = True
    mcOffset = cConstant
    mcConstant = cConstant
    mcStorageRange = 0
    mlStorageBits = 0
    mlStorageBytes = 0
    mbSigned = cConstant < 0
End Property

Public Property Let Starting(ByVal cStarting As Currency)
    mcStarting = cStarting
    mcEnding = cStarting
    mcConstant = cStarting
    mbIsConstant = True
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
        mbIsConstant = False
    Else
        mlStorageBits = 0
        mlStorageBytes = 0
        mbIsConstant = True
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
    Dim lBitsPositive As Long
    Dim lBitsNegative As Long
    
    If Not mbIsConstant Then
        If Not mbSigned Then
            CalcBytes = mlStorageBytes
        Else
            If mcEnding > -1 Then
                lBitsPositive = Int(Log(mcEnding) / Log(2)) + 1
            End If
            If mcStarting < -1 Then
                lBitsNegative = Int(Log(-mcStarting - 1) / Log(2)) + 1
            End If
            If lBitsNegative > lBitsPositive Then
                CalcBytes = lBitsNegative \ 8 + 1
            Else
                CalcBytes = lBitsPositive \ 8 + 1
            End If
        End If
    Else
        If mcConstant >= 0 Then
            CalcBytes = Int(Log(mcConstant) / Log(256)) + 1
        Else
            CalcBytes = Int(Log(-mcConstant - 1) / Log(256)) + 1
        End If
    End If
End Property

Public Property Get HasOffset() As Boolean
    HasOffset = mcOffset <> 0
End Property

Private Sub Class_Initialize()
    mcStarting = 0
    mcEnding = -1
    mcStorageRange = -1
    mlStorageBits = -1
    mlStorageBytes = -1
End Sub

