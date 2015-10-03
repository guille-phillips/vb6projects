Attribute VB_Name = "modMidi"
Option Explicit

Private Declare Function midiOutGetNumDevs Lib "winmm" () As Integer
Private Declare Function midiOutGetDevCaps Lib "winmm.dll" Alias "midiOutGetDevCapsA" (ByVal uDeviceID As Long, lpCaps As MIDIOUTCAPS, ByVal uSize As Long) As Long
Private Declare Function midiInGetNumDevs Lib "winmm.dll" () As Long
Private Declare Function midiInGetDevCaps Lib "winmm.dll" Alias "midiInGetDevCapsA" (ByVal uDeviceID As Long, lpCaps As MIDIINCAPS, ByVal uSize As Long) As Long
Public Declare Function midiOutShortMsg Lib "winmm.dll" (ByVal hMidiOUT As Long, ByVal dwMsg As Long) As Long
Private Declare Function midiOutOpen Lib "winmm.dll" (lphMidiOut As Long, ByVal uDeviceID As Long, ByVal dwCallback As Long, ByVal dwInstance As Long, ByVal dwFlags As Long) As Long
Public Declare Function midiInOpen Lib "winmm.dll" (lphMidiIn As Long, ByVal uDeviceID As Long, ByVal dwCallback As Long, ByVal dwInstance As Long, ByVal dwFlags As Long) As Long
Public Declare Function midiOutClose Lib "winmm.dll" (ByVal hMidiOUT As Long) As Long
Public Declare Function midiInClose Lib "winmm.dll" (ByVal hMidiIN As Long) As Long
Private Declare Function midiInStart Lib "winmm.dll" (ByVal hMidiIN As Long) As Long
Public Declare Function midiInStop Lib "winmm.dll" (ByVal hMidiIN As Long) As Long

Private Const MAXPNAMELEN = 32       '  max product name length (including NULL)

Private Type MIDIINCAPS
        wMid As Integer
        wPid As Integer
        vDriverVersion As Long
        szPname As String * MAXPNAMELEN
End Type

Private Type MIDIOUTCAPS
        wMid As Integer
        wPid As Integer
        vDriverVersion As Long
        szPname As String * MAXPNAMELEN
        wTechnology As Integer
        wVoices As Integer
        wNotes As Integer
        wChannelMask As Integer
        dwSupport As Long
End Type

Private Const MMSYSERR_NOERROR = 0               '  no error

Public Const CALLBACK_NULL = &H0                '  no callback
Public Const CALLBACK_FUNCTION = &H30000        '  dwCallback is a FARPROC

' MIDI status messages
Public Const NOTE_OFF = &H80
Public Const NOTE_ON = &H90

Public lInputPortDevice As Long
Public lOutputPortDevice As Long
Public hMidiOutput As Long
Public hMidiInput As Long

Public Const MM_MIM_OPEN = &H3C1                '  MIDI input
Public Const MM_MIM_CLOSE = &H3C2
Public Const MM_MIM_DATA = &H3C3
Public Const MM_MIM_LONGDATA = &H3C4
Public Const MM_MIM_ERROR = &H3C5
Public Const MM_MIM_LONGERROR = &H3C6

Private Type mmMidiMessageType
    Byte0 As Byte
    Byte1 As Byte
    Byte2 As Byte
    Byte3 As Byte
End Type

Public Sub Initialise()
    Dim iOutputDevicesCount As Integer
    Dim iInputDevicesCount As Integer
    Dim lIndex As Long
    Dim mocMidiOutCaps As MIDIOUTCAPS
    Dim micMidiInCaps As MIDIINCAPS
    Dim lError As Long
    Dim sName As String
    
    lOutputPortDevice = -1
    iOutputDevicesCount = midiOutGetNumDevs()
    For lIndex = 0 To iOutputDevicesCount - 1
        lError = midiOutGetDevCaps(lIndex, mocMidiOutCaps, 52)
        If lError <> MMSYSERR_NOERROR Then MsgBox lError, vbOKOnly, "Output Port Device Caps": Exit Sub
        sName = Left$(mocMidiOutCaps.szPname, InStr(mocMidiOutCaps.szPname, Chr$(0)) - 1)
        If sName = "USB Audio Device" Then
            lOutputPortDevice = lIndex
            Exit For
        End If
    Next
    
    lInputPortDevice = -1
    iInputDevicesCount = midiInGetNumDevs()
    For lIndex = 0 To iInputDevicesCount - 1
        lError = midiInGetDevCaps(lIndex, micMidiInCaps, 40)
        If lError <> MMSYSERR_NOERROR Then MsgBox lError, vbOKOnly, "Input Port Device Caps": Exit Sub
        sName = Left$(micMidiInCaps.szPname, InStr(micMidiInCaps.szPname, Chr$(0)) - 1)
        If sName = "USB Audio Device" Then
            lInputPortDevice = lIndex
            Exit For
        End If
    Next
    
    hMidiOutput = GetSetting("Midiplayer", "Ports", "OutputPortHandle", 0)
    midiOutClose hMidiOutput
    
    hMidiInput = GetSetting("Midiplayer", "Ports", "InputPortHandle", 0)
    midiInStop hMidiInput
    midiInClose hMidiInput
    
    lError = midiOutOpen(hMidiOutput, lOutputPortDevice, vbNull, 0, CALLBACK_NULL)
    If lError <> MMSYSERR_NOERROR Then MsgBox lError, vbOKOnly, "Output Port Open Error": Exit Sub
    SaveSetting "MidiPlayer", "Ports", "OutputPortHandle", CStr(hMidiOutput)
    
    lError = midiInOpen(hMidiInput, lInputPortDevice, AddressOf MidiMessageIn, 0, CALLBACK_FUNCTION)
    If lError <> MMSYSERR_NOERROR Then MsgBox lError, vbOKOnly, "Input Port Open Error": Exit Sub
    SaveSetting "MidiPlayer", "Ports", "InputPortHandle", CStr(hMidiInput)
    lError = midiInStart(hMidiInput)
    If lError <> MMSYSERR_NOERROR Then MsgBox lError, vbOKOnly, "Input Port Start Error": Exit Sub
End Sub

Public Sub MidiMessageIn(ByVal hmIN As Long, ByVal wMsg As Long, ByVal dwInstance As Long, ByVal dwParam1 As Long, ByVal dwParam2 As Long)
    Dim mmMessage As mmMidiMessageType
    Dim ntNote As ntNoteType
    Dim fTiming As Double
    
    Dim x As Long
    Dim y As Long
    
    x = 23
    y = x
'    CopyMemory y, x, 4&
'    MsgBox y
'    End
    
    Select Case wMsg
        Case MM_MIM_OPEN
            'Stop
        Case MM_MIM_CLOSE
            'Stop
        Case MM_MIM_DATA
'            CopyMemory mmMessage, VarPtr(dwParam1), 4&
            Select Case mmMessage.Byte0
'                Case &HF8 ' Clock
'                Case &HFE ' Active Sensing
                Case &H90  ' Note event
'                    fTiming = GetCounter
'                    ntNote.Pitch = mmMessage.Byte1
'                    ntNote.Volume = mmMessage.Byte2
'                    ntNote.Position = fTiming / 0.002
'                    Debug.Print "Pitch:" & mmMessage.Byte1 & "Volume:" & mmMessage.Byte2
'                    If ntNote.Volume = 0 Then
'                        ReceiveNoteUp ntNote
'                    Else
'                        ReceiveNoteDown ntNote
'                    End If
'                Case Else
'                    Debug.Print hmIN & " " & HexNum(wMsg, 8) & " " & dwInstance & " " & HexNum(dwParam1, 8) & " " & HexNum(dwParam2, 8)
            End Select
            'Stop
        Case MM_MIM_LONGDATA
            'Stop
        Case MM_MIM_ERROR
            'Stop
        Case MM_MIM_LONGERROR
            'Stop
    End Select
End Sub

Public Sub ReceiveNoteDown(ntNote As ntNoteType)
    modSound.PlaySound ntNote.Pitch, ntNote.Volume
End Sub

Public Sub ReceiveNoteUp(ntNote As ntNoteType)
    modSound.StopSound ntNote.Pitch
End Sub

Public Sub PlayNoteDown(ntNote As ntNoteType)
    Dim mmMessage As mmMidiMessageType
    Dim lMessage As Long
    
    mmMessage.Byte0 = NOTE_ON
    mmMessage.Byte1 = ntNote.Pitch
    mmMessage.Byte2 = ntNote.Volume
    mmMessage.Byte3 = 0
    
    CopyMemory lMessage, mmMessage, 4&
    
    midiOutShortMsg hMidiOutput, lMessage
End Sub

Public Sub PlayNoteUp(ntNote As ntNoteType)
    Dim mmMessage As mmMidiMessageType
    Dim lMessage As Long
    
    mmMessage.Byte0 = NOTE_OFF
    mmMessage.Byte1 = ntNote.Pitch
    mmMessage.Byte2 = 0
    mmMessage.Byte3 = 0
    
    CopyMemory lMessage, mmMessage, 4&
    
    midiOutShortMsg hMidiOutput, lMessage
End Sub
