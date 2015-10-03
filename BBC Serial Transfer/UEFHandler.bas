Attribute VB_Name = "UEFHandler"
Option Explicit

Private Declare Function SetDllDirectory Lib "kernel32" Alias "SetDllDirectoryA" (ByVal path As String) As Long

Private Declare Function gzopen Lib "zlibwapi.dll" (ByVal filePath As String, ByVal mode As String) As Long
Private Declare Function gzread Lib "zlibwapi.dll" (ByVal file As Long, ByVal uncompr As String, ByVal uncomprLen As Integer) As Integer
'Private Declare Function gzwrite Lib "zlibwapi.dll" (ByVal file As Long, ByVal uncompr As String, ByVal uncomprLen As Integer) As Integer
Private Declare Function gzwrite Lib "zlibwapi.dll" (ByVal file As Long, uncompr As Any, ByVal uncomprLen As Long) As Integer
Private Declare Function gzclose Lib "zlibwapi.dll" (ByVal file As Long) As Integer
Private Declare Function gzeof Lib "zlibwapi.dll" (ByVal file As Long) As Integer
Private Declare Function gzgetc Lib "zlibwapi.dll" (ByVal file As Long) As Integer
Private Declare Function gzputc Lib "zlibwapi.dll" (ByVal file As Long, ByVal C As Integer) As Integer

Private myFile() As Byte
Private mlTotalLength As Long
Private mlBlockPointer As Long
Public BlockIdentifier As Long
Public BlockLength As Long
Public BlockData() As Byte
Public MinorVersion As Byte
Public MajorVersion As Byte
Private myDummy As Byte

Public Function LoadCompressedUEFFile(ByVal sPath As String) As Boolean
    Dim sIdentifier As String * 9
    Dim lGZFileHandle As Long
    Dim lByte As Long
    
    SetDllDirectory App.path
    lGZFileHandle = gzopen(sPath, "rb")
    gzread lGZFileHandle, sIdentifier, 9
    
    If sIdentifier <> "UEF File!" Then
        gzclose lGZFileHandle
        Exit Function
    End If
    LoadCompressedUEFFile = True
    
    gzread lGZFileHandle, myDummy, 1
    gzread lGZFileHandle, MinorVersion, 1
    gzread lGZFileHandle, MajorVersion, 1
    
    ResetUEF
    
    mlTotalLength = 0
    
    Do While gzeof(lGZFileHandle) = 0
        ReDim Preserve myFile(mlTotalLength)
        lByte = gzgetc(lGZFileHandle)
        If lByte >= 0 Then
            myFile(mlTotalLength) = lByte
            mlTotalLength = mlTotalLength + 1
        Else
            Exit Do
        End If
    Loop
        
    gzclose lGZFileHandle
End Function

Public Function LoadUEFFile(ByVal sPath As String) As Boolean
    Dim sIdentifier As String * 9
    
    Open sPath For Binary As #1
    
    Get #1, , sIdentifier
    If sIdentifier <> "UEF File!" Then
        Exit Function
    End If
    LoadUEFFile = True
    
    Get #1, , myDummy
    Get #1, , MinorVersion
    Get #1, , MajorVersion
    
    ResetUEF
    
    mlTotalLength = FileLen(sPath) - 12&
    ReDim myFile(mlTotalLength)
    Get #1, , myFile
    Close #1
End Function

Public Function LoadNextBlock() As Boolean
    If mlBlockPointer = mlTotalLength Then
        Exit Function
    End If
    
    CopyMemory BlockIdentifier, myFile(mlBlockPointer), 2&
    CopyMemory BlockLength, myFile(mlBlockPointer + 2), 4&
    ReDim Preserve BlockData(BlockLength - 1)
    CopyMemory BlockData(0), myFile(mlBlockPointer + 6), BlockLength
    mlBlockPointer = mlBlockPointer + 6& + BlockLength
    
    LoadNextBlock = True
End Function

Public Function FindBlock(lBlockIdentifier As Long) As Boolean
    Dim mlBlockIdentifier As Long
    Dim mlBlockLength As Long
    
    Do
        CopyMemory mlBlockIdentifier, myFile(mlBlockPointer), 2&
        CopyMemory mlBlockLength, myFile(mlBlockPointer + 2), 4&
        
        If mlBlockIdentifier = lBlockIdentifier Then
            LoadNextBlock
            FindBlock = True
            Exit Function
        Else
            mlBlockPointer = mlBlockPointer + 6& + mlBlockLength
        End If
    Loop Until mlBlockPointer = mlTotalLength
End Function

Public Sub ResetUEF()
    mlBlockPointer = 0
    BlockIdentifier = 0
    BlockLength = 0
    Erase BlockData
End Sub

Public Sub CreateUEFFile()
    ResetUEF
    mlBlockPointer = 0
    mlTotalLength = 0
    Erase myFile
End Sub

Public Sub SaveBlock(ByVal lBlockIdentifier As Long)
    BlockLength = UBound(BlockData) + 1
    ReDim Preserve myFile(mlTotalLength - 1 + 6 + BlockLength)
    CopyMemory myFile(mlTotalLength), lBlockIdentifier, 2&
    CopyMemory myFile(mlTotalLength + 2), BlockLength, 4&
    CopyMemory myFile(mlTotalLength + 6), BlockData(0), BlockLength
    mlTotalLength = mlTotalLength + 6 + BlockLength
End Sub

Public Sub ResetBlock(lSize As Long)
    ReDim BlockData(lSize - 1)
    BlockLength = lSize
End Sub

Public Sub SaveUEFFile(ByVal sPath As String)
    If Dir(sPath) <> "" Then
        Kill sPath
    End If
    Open sPath For Binary As #2
    Put #2, , "UEF File!" & Chr$(0)
    Put #2, , CByte(10)
    Put #2, , CByte(0)
    Put #2, , myFile
    Close #2
End Sub

Public Sub SaveCompressedUEFFile(ByVal sPath As String)
    Dim lGZFileHandle As Long
    Dim sName As String
    Dim lIndex As Long
    
    If Dir(sPath) <> "" Then
        Kill sPath
    End If
    lGZFileHandle = gzopen(sPath, "w")
    
    sName = "UEF File!" & Chr$(0)
    
    gzwrite lGZFileHandle, ByVal sName, Len(sName)
    gzputc lGZFileHandle, CByte(10)
    gzputc lGZFileHandle, CByte(0)
    For lIndex = 0 To UBound(myFile)
        gzputc lGZFileHandle, myFile(lIndex)
    Next
    'gzwrite lGZFileHandle, myFile(0), UBound(myFile) + 1
    gzclose lGZFileHandle
End Sub

