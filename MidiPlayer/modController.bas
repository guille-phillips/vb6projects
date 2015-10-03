Attribute VB_Name = "modController"
Option Explicit

Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Public oInterface As frmInterface

Public Enum mdModes
    mdPlaying
    mdPracticing
    mdEditing
    mdRecording
End Enum

Public mdMode As mdModes

Sub Main()
    Initialise
    InitialiseSequence

    InitialiseCounters
    
    'PlaySequence
    
    Set oInterface = New frmInterface
    oInterface.Show
    
    InitialiseSound oInterface
End Sub
