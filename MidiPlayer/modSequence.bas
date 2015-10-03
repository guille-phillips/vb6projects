Attribute VB_Name = "modSequence"
Option Explicit

Public Type ntNoteType
    Pitch As Long
    Duration As Long ' 960 parts of bar (4/4)
    Volume As Long
    Length As Long
    Position As Long ' bar = 960
End Type

Public ntNotes() As ntNoteType
Public lNotesCount As Long
Public lNotesPointer As Long

Public Type nrNoteReferencesType
    NoteRef(5) As Byte
End Type
Public nrNoteReferences() As nrNoteReferencesType
Public lNoteReferencesCount As Long


Public Sub InitialiseSequence()
    ReDim ntNotes(1)
    
    With ntNotes(0)
        .Pitch = 40
        .Volume = 80
        .Position = 0
        .Duration = 960 / 4
    End With
    
    With ntNotes(1)
        .Pitch = 44
        .Volume = 80
        .Position = 0
        .Duration = 960 / 4
    End With
    
    lNotesPointer = 0
    lNotesCount = 2
    
    ConvertSequence
End Sub

Public Sub ConvertSequence()
    Dim lNoteIndex As Long
    Dim lNoteUpPosition As Long
    
    For lNoteIndex = 0 To lNotesCount - 1
        lNoteUpPosition = ntNotes(lNoteIndex).Position + ntNotes(lNoteIndex).Duration
        If lNoteUpPosition > lNoteReferencesCount Then
            lNoteReferencesCount = lNoteUpPosition
            ReDim Preserve nrNoteReferences(lNoteReferencesCount)
        End If
        
        PushNoteReference ntNotes(lNoteIndex).Position, lNoteIndex + 1
        PushNoteReference lNoteUpPosition, lNoteIndex + 1 + 128
    Next
End Sub

Private Sub PushNoteReference(lIndex, lValue)
    Dim lCheck As Long
    
    While nrNoteReferences(lIndex).NoteRef(lCheck) <> 0 And lCheck <= 5
        lCheck = lCheck + 1
    Wend
    If lCheck <= 5 Then
        nrNoteReferences(lIndex).NoteRef(lCheck) = lValue
    End If
End Sub

Public Sub PlaySequence()
    Dim lTick As Long
    Dim fNext As Double
    Dim fTime As Double
    Dim lIndex As Long
    
    StartCounter
    
    While lTick <= lNoteReferencesCount
        fTime = GetCounter
        While fTime < fNext
            DoEvents
            fTime = GetCounter
        Wend
        
        If lNoteReferencesCount > 0 Then
            If lTick <= lNoteReferencesCount Then
                For lIndex = 0 To 5
                    If nrNoteReferences(lTick).NoteRef(lIndex) >= 128 Then
                        PlayNoteUp ntNotes(nrNoteReferences(lTick).NoteRef(lIndex) - 129)
                    ElseIf nrNoteReferences(lTick).NoteRef(lIndex) > 0 Then
                        PlayNoteDown ntNotes(nrNoteReferences(lTick).NoteRef(lIndex) - 1)
                    End If
                Next
            End If
        End If
        
        fNext = fNext + 0.002
        lTick = lTick + 1
        'Debug.Print fTime & " " & fNext & " " & lTick
    Wend
End Sub
