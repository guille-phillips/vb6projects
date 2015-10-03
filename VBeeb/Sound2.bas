Attribute VB_Name = "Sound2"
Option Explicit

Private dx As DirectX8  'Our DirectX object
Private ds As DirectSound8 'Our DirectSound object
Private mdsBuf As DSBUFFERDESC
Private mdsBuffer(3) As DirectSoundSecondaryBuffer8 'Our SoundBuffer

Private myToneBufferOn(99) As Byte
Private myToneBufferOff(99) As Byte
Private mlTotalCycles As Long
Private mlToggle As Long

Public Sub InitialiseSound2()
    Dim dssbBuffer As DirectSoundSecondaryBuffer8
    Dim lChannel As Long
    Dim lIndex As Long
    
    Set dx = New DirectX8
        
    For lIndex = 0 To 99
        myToneBufferOn(lIndex) = 128 + 8
        
        If lIndex <= 49 Then
            myToneBufferOff(lIndex) = 128 + 40
        Else
            myToneBufferOff(lIndex) = 128
        End If
    Next
    
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
    dssbBuffer.WriteBuffer 0, mdsBuf.lBufferBytes, myToneBufferOff(0), 0
    
    For lChannel = 0 To 0
        Set mdsBuffer(lChannel) = ds.DuplicateSoundBuffer(dssbBuffer)
        mdsBuffer(lChannel).WriteBuffer 0, mdsBuf.lBufferBytes, myToneBufferOff(0), 0
        mdsBuffer(lChannel).SetFrequency 10000
        mdsBuffer(lChannel).SetVolume -1000
        mdsBuffer(lChannel).Play DSBPLAY_LOOPING
    Next
End Sub

Public Sub Tick(ByVal lCycles As Long)
    mlTotalCycles = mlTotalCycles - lCycles
    
    If mlTotalCycles < 0 Then
        If mlToggle = 0 Then
            'mdsBuffer(0).WriteBuffer 0, 100, myToneBufferOff(0), 0
            'mdsBuffer(0).Play DSBPLAY_LOOPING
            mdsBuffer(0).SetCurrentPosition (49)
        Else
            'mdsBuffer(0).WriteBuffer 0, 100, myToneBufferOn(1), 0
            'mdsBuffer(0).Stop
        End If
        
        mlToggle = 1 - mlToggle
        mlTotalCycles = mlTotalCycles + 1000
    End If
End Sub
