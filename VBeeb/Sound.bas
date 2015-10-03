Attribute VB_Name = "Sound"
Option Explicit

Private dx As DirectX8  'Our DirectX object
Private ds As DirectSound8 'Our DirectSound object
Private mdsBuf As DSBUFFERDESC
Private mdsBuffer(4) As DirectSoundSecondaryBuffer8 'Our SoundBuffer

Private myToneBuffer(99) As Byte
Private myNoiseBuffer(32767) As Byte
Private myPeriodicNoiseBuffer(15) As Byte

Private mlWhiteNoiseOn As Long
Private mbLinkChannel As Boolean

Public Type Channel
    Frequency As Long
    Volume As Long
End Type

Public chChannels(3) As Channel
Private mlRegister As Long
Private mlVolume As Long

Public Enabled As Long

Private mbNoDirectX As Boolean

Private Sub SetUpToneBuffer()
    Dim lIndex As Long

    ' Debugging.WriteString "Sound.SetUpToneBuffer"
    
    For lIndex = 0 To 99
        'myToneBuffer(lIndex) = ((lIndex \ 5) * 2 - 1) * 8 + 128
        myToneBuffer(lIndex) = ((lIndex \ 50) * 2 - 1) * 8 + 128
    Next
End Sub

Private Sub SetUpPeriodicNoiseBuffer()
    Dim lIndex As Long
    Dim lStep As Long
    
    ' Debugging.WriteString "Sound.SetUpPeriodicNoiseBuffer"
    
    For lIndex = 0 To 7
        myPeriodicNoiseBuffer(lIndex) = 128 - (7 - lIndex) + 4
    Next
    For lIndex = 8 To 15
        myPeriodicNoiseBuffer(lIndex) = 128 + 11 - lIndex
    Next
End Sub

Private Sub SetUpNoiseBuffer()
    Dim lSample As Long
    Dim lValue As Long
    Dim lNewBit As Long
    
    ' Debugging.WriteString "Sound.SetUpNoiseBuffer"
    
    lValue = 32768
    
    Do
        myNoiseBuffer(lSample) = ((lValue And 1&) * 2 - 1) * 8 + 128
        lNewBit = ((lValue And 32768) \ 32768) Xor ((lValue And 4&) \ 4&) Xor (lValue And 1&)
        lValue = lValue \ 2 Or lNewBit * 32768
        lSample = lSample + 1
    Loop Until lValue = 32768
    
End Sub

Public Sub InitialiseSound()
    Dim lFrequency As Long
    Dim lBit As Long
    Dim lMask As Long
    Dim lValue As Long
    Dim lIndex As Long
    Dim lChannel As Long
    Dim dssbBuffer As DirectSoundSecondaryBuffer8
    
    ' Debugging.WriteString "Sound.InitialiseSound"
    
    On Error Resume Next
    
    Set dx = New DirectX8
    
    SetUpToneBuffer
    SetUpPeriodicNoiseBuffer
    SetUpNoiseBuffer
    
    Set ds = dx.DirectSoundCreate(vbNullString) 'Create a default DirectSound object
    ds.SetCooperativeLevel Console.hWnd, DSSCL_PRIORITY

    mdsBuf.lFlags = DSBCAPS_CTRLVOLUME Or DSBCAPS_CTRLFREQUENCY

    With mdsBuf.fxFormat
        .nFormatTag = WAVE_FORMAT_PCM
        .nChannels = 1
        .lSamplesPerSec = 44100
        .nBitsPerSample = 8
        .nBlockAlign = 1
        .lAvgBytesPerSec = .lSamplesPerSec * .nBlockAlign
    End With
    

    mdsBuf.lBufferBytes = 100
    Set dssbBuffer = ds.CreateSoundBuffer(mdsBuf)
    dssbBuffer.WriteBuffer 0, mdsBuf.lBufferBytes, myToneBuffer(0), 0
    
    For lChannel = 0 To 2
        Set mdsBuffer(lChannel) = ds.DuplicateSoundBuffer(dssbBuffer)
    Next


    mdsBuf.lBufferBytes = 32768
    Set dssbBuffer = ds.CreateSoundBuffer(mdsBuf)
    dssbBuffer.WriteBuffer 0, mdsBuf.lBufferBytes, myNoiseBuffer(0), 0
    Set mdsBuffer(3) = dssbBuffer
    
    mdsBuf.lBufferBytes = 15
    Set dssbBuffer = ds.CreateSoundBuffer(mdsBuf)
    dssbBuffer.WriteBuffer 0, mdsBuf.lBufferBytes, myPeriodicNoiseBuffer(0), 0
    Set mdsBuffer(4) = dssbBuffer
    
    For lChannel = 0 To 4
        mdsBuffer(lChannel).SetFrequency 1200000 / 6
        mdsBuffer(lChannel).SetVolume -10000
        mdsBuffer(lChannel).Play DSBPLAY_LOOPING
    Next
    
    Exit Sub
    
InitialiseSoundNoDirectX:
    mbNoDirectX = True
    MsgBox "No direct X"
End Sub

Private Sub SetNoiseFrequency(ByVal lValue As Long)
    ' Debugging.WriteString "Snapshot.SetNoiseFrequency"
    
    On Error Resume Next
    
    If lValue >= 6 Then
        mdsBuffer(3).SetCurrentPosition 0
        mdsBuffer(3).SetFrequency 112000 / lValue
    Else
        mdsBuffer(3).SetFrequency 200000
    End If
End Sub

Private Sub SetPeriodicNoiseFrequency(ByVal lValue As Long)
    ' Debugging.WriteString "Snapshot.SetPeriodicNoiseFrequency"
    
    On Error Resume Next
    
    If lValue >= 6 Then
        mdsBuffer(4).SetCurrentPosition 0
        mdsBuffer(4).SetFrequency 112000 / lValue
    Else
        mdsBuffer(4).SetFrequency 200000
    End If
End Sub

Private Sub SetTone(ByVal lChannel As Long, ByVal lValue As Long)
    ' Debugging.WriteString "Snapshot.SetTone"
    
    On Error Resume Next
    
    If lValue >= 6 Then
        mdsBuffer(lChannel).SetCurrentPosition 0
        mdsBuffer(lChannel).SetFrequency 12500000 / lValue
    Else
        mdsBuffer(lChannel).SetFrequency 200000
    End If
End Sub

Private Sub SetVolume(ByVal lChannel As Long, ByVal lValue As Long)
    ' Debugging.WriteString "Snapshot.SetVolume"
    
    On Error Resume Next
    
    If lValue <> 15 Then
        mdsBuffer(lChannel).SetVolume lValue * -200
    Else
        mdsBuffer(lChannel).SetVolume -10000
    End If
End Sub

Public Function WriteByte(ByVal lByte As Long)
    Dim lRegister As Long
    Dim mlVolume As Long
    Dim lData As Long
    Dim dFrequency As Double
    Dim lType As Long
    Dim lNoiseFrequencyType As Long
    Dim lWhiteNoise As Long
    
    ' Debugging.WriteString "Snapshot.WriteByte"
    
    'Debug.Print HexNum(lByte, 2)
    
    On Error Resume Next
    
    If (lByte And &H80&) = &H80& Then
        mlRegister = (lByte And &H60&) \ 32& Xor 3
        mlVolume = (lByte And &H10&)
        lData = (lByte And &HF&)
        
        If mlVolume = &H10& Then
            chChannels(mlRegister).Volume = lData
            If mlRegister <> 0 Then
                SetVolume mlRegister - 1, chChannels(mlRegister).Volume
            Else
                SetVolume 4 - mlWhiteNoiseOn, chChannels(mlRegister).Volume
            End If
'                Debug.Print "Volume " & mlRegister & ":" & lData
        Else
            If mlRegister <> 0 Then
                chChannels(mlRegister).Frequency = (chChannels(mlRegister).Frequency And &H3F00&) Or lData
                SetTone mlRegister - 1, chChannels(mlRegister).Frequency
                If mlRegister = 1 And mbLinkChannel Then
                    SetPeriodicNoiseFrequency chChannels(1).Frequency
                End If
'                    Debug.Print "Frequency " & mlRegister & ":" & chChannels(mlRegister).Frequency
            Else
                lNoiseFrequencyType = lData And 3&
                lWhiteNoise = lData And 4&
                
'                If lWhiteNoise = 4& Then
'                    If mlWhiteNoiseOn = 0& Then
'                        SetVolume 4, 15
'                    End If
'                Else
'                    If mlWhiteNoiseOn = 1& Then
'                        SetVolume 3, 15
'                    End If
'                End If
                
                mlWhiteNoiseOn = Abs(Sgn(lWhiteNoise))
                SetVolume 4 - mlWhiteNoiseOn, chChannels(mlRegister).Volume
                SetVolume 3 + mlWhiteNoiseOn, 15
                
                If lWhiteNoise = 4& Then
                    Select Case lNoiseFrequencyType
                        Case 0
                            mbLinkChannel = False
                            SetNoiseFrequency 15
                        Case 1
                            mbLinkChannel = False
                            SetNoiseFrequency 30
                        Case 2
                            mbLinkChannel = False
                            SetNoiseFrequency 60
                        Case Else
                            mbLinkChannel = True
                            SetNoiseFrequency chChannels(1).Frequency
                    End Select
                Else
                    Select Case lNoiseFrequencyType
                        Case 0, 1, 2
                            mbLinkChannel = False
                            SetPeriodicNoiseFrequency (lNoiseFrequencyType + 1) * 16
                        Case Else
                            mbLinkChannel = True
                            SetPeriodicNoiseFrequency chChannels(1).Frequency
                    End Select
                End If
            End If
        End If
    Else
        If mlVolume = &H10& Then
            lData = lByte And &HF&
            chChannels(mlRegister).Volume = lData
            
            If mlRegister <> 0 Then
                SetVolume mlRegister - 1, chChannels(mlRegister).Volume
            Else
                SetVolume 4 - mlWhiteNoiseOn, chChannels(mlRegister).Volume
            End If
'                Debug.Print "Update Volume " & mlRegister & ":" & lData
        Else
            If mlRegister <> 0 Then
                lData = lByte And &H3F&
                chChannels(mlRegister).Frequency = (chChannels(mlRegister).Frequency And &HF&) + lData * 16&
                
                SetTone mlRegister - 1, chChannels(mlRegister).Frequency
                If mlRegister = 1 And mbLinkChannel Then
                    SetPeriodicNoiseFrequency chChannels(1).Frequency
                End If
                'Debug.Print "Update Frequency " & mlRegister & ":" & chChannels(mlRegister).Frequency
            Else
            End If
        End If
    End If
End Function

Public Sub PauseSound()
    SetVolume 0, 15
    SetVolume 1, 15
    SetVolume 2, 15
    SetVolume 3, 15
End Sub

Public Sub ResumeSound()
    SetVolume 0, chChannels(0).Volume
    SetVolume 1, chChannels(1).Volume
    SetVolume 2, chChannels(2).Volume
    SetVolume 3, chChannels(3).Volume
End Sub
