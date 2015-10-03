Attribute VB_Name = "modSound"
Option Explicit

Private dx As DirectX8  'Our DirectX object
Private ds As DirectSound8 'Our DirectSound object
Private mdsBuf As DSBUFFERDESC
Private mdsBuffer(7) As DirectSoundSecondaryBuffer8 'Our SoundBuffer

Private myToneBufferHi(37) As Byte
Private myToneBufferLo(255) As Byte
Private myNoiseBuffer(32766) As Byte
Private myPeriodicNoiseBuffer(14) As Byte


Private mlChannelNote(3) As Long

Private mbNoDirectX As Boolean

Private mlFrequency(255) As Long

Public Sub InitialiseSound(oInterface As frmInterface)
    Dim lFrequency As Long
    Dim lBit As Long
    Dim lMask As Long
    Dim lValue As Long
    Dim lIndex As Long
    Dim lChannel As Long
    Dim dssbBuffer As DirectSoundSecondaryBuffer8
    
    ' Debugging.WriteString "Sound.InitialiseSound"
    
    'On Error Resume Next
    On Error GoTo InitialiseSoundNoDirectX
    
    Set dx = New DirectX8
    
    For lIndex = 0 To 255
        mlFrequency(lIndex) = 440 * 2 ^ (CDbl(lIndex) / 12)
    Next
    
    For lIndex = 0 To UBound(mlChannelNote)
        mlChannelNote(lIndex) = -1
    Next
    
    SetUpToneBuffers
    SetUpPeriodicNoiseBuffer
    SetUpNoiseBuffer
    
    Set ds = dx.DirectSoundCreate(vbNullString) 'Create a default DirectSound object
    ds.SetCooperativeLevel oInterface.hWnd, DSSCL_PRIORITY

    mdsBuf.lFlags = DSBCAPS_CTRLVOLUME Or DSBCAPS_CTRLFREQUENCY

    With mdsBuf.fxFormat
        .nFormatTag = WAVE_FORMAT_PCM
        .nChannels = 1
        .lSamplesPerSec = 44100
        .nBitsPerSample = 8
        .nBlockAlign = 1
        .lAvgBytesPerSec = .lSamplesPerSec * .nBlockAlign
    End With
    

    ' Lo End Tone
    mdsBuf.lBufferBytes = 256
    Set dssbBuffer = ds.CreateSoundBuffer(mdsBuf)
    dssbBuffer.WriteBuffer 0, mdsBuf.lBufferBytes, myToneBufferLo(0), 0

    For lChannel = 0 To 3
        Set mdsBuffer(lChannel) = ds.DuplicateSoundBuffer(dssbBuffer)
    Next
    

    For lChannel = 0 To 3
        mdsBuffer(lChannel).SetFrequency 4704025 / 1008
        mdsBuffer(lChannel).SetVolume -1000
        mdsBuffer(lChannel).Play DSBPLAY_LOOPING
    Next
    

    Exit Sub
    
InitialiseSoundNoDirectX:
    mbNoDirectX = True
    MsgBox "No direct X." & vbCrLf & "Sound will be unavailable."
End Sub


Private Sub SetUpToneBuffers()
    Dim lIndex As Long

    ' Debugging.WriteString "Sound.SetUpToneBuffer"
    
    For lIndex = 0 To 37
        myToneBufferHi(lIndex) = ((lIndex \ 19) * 2 - 1) * 8 + 128
    Next
    
    For lIndex = 0 To 255
        myToneBufferLo(lIndex) = ((lIndex \ 128) * 2 - 1) * 8 + 128
    Next
End Sub

Private Sub SetUpPeriodicNoiseBuffer()
    Dim lIndex As Long
    Dim lStep As Long
    
    ' Debugging.WriteString "Sound.SetUpPeriodicNoiseBuffer"
    
    myPeriodicNoiseBuffer(0) = 128 + 8
    For lIndex = 1 To 14
        myPeriodicNoiseBuffer(lIndex) = 128 - 8
    Next
End Sub

Private Sub SetUpNoiseBuffer()
    Dim lSample As Long
    Dim lValue As Long
    Dim lNewBit As Long
    
    ' Debugging.WriteString "Sound.SetUpNoiseBuffer"
    
    lValue = 16384
    
    Do
        myNoiseBuffer(lSample) = ((lValue And 1&) * 2 - 1) * 8 + 128
        lNewBit = (lValue \ 2&) Xor (lValue \ 1&)
        
        lValue = lValue \ 2 Or (lNewBit And &H1&) * 16384
        lSample = lSample + 1
    Loop Until lValue = 16384
      
End Sub




Public Sub PlaySound(lNote As Long, lVolume As Long)
    Dim lIndex As Long
    
    For lIndex = 0 To UBound(mlChannelNote)
        If mlChannelNote(lIndex) = -1 Then
        Debug.Print "Donw:" & lIndex
            mdsBuffer(lIndex).SetFrequency mlFrequency(lNote)
            mdsBuffer(lIndex).SetVolume -1000 / lVolume
'            mdsBuffer(lIndex).Play DSBPLAY_LOOPING
            mlChannelNote(lIndex) = lNote
            Exit For
        End If
    Next
End Sub

Public Sub StopSound(lNote As Long)
    Dim lIndex As Long
    
    For lIndex = 0 To UBound(mlChannelNote)
        If mlChannelNote(lIndex) = lNote Then
        Debug.Print "Up:" & lIndex
            mdsBuffer(lIndex).SetVolume -10000
            mlChannelNote(lIndex) = -1
            Exit For
        End If
    Next

End Sub
