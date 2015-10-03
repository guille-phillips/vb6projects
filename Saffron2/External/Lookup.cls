VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Lookup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private oConnection As New Connection

Public Function Verb(sText As String, lPosition As Long, lIndex As Long, vParams As Variant) As Boolean
    Dim lLength As Long

    lLength = PartOfSpeech("Verbs", sText, lPosition)
    
    If lLength > 0 Then
        lPosition = lPosition + lLength
        Verb = True
    End If
End Function

Public Function Noun(sText As String, lPosition As Long, lIndex As Long, vParams As Variant) As Boolean
    Dim lLength As Long

    lLength = PartOfSpeech("Nouns", sText, lPosition)
    
    If lLength > 0 Then
        lPosition = lPosition + lLength
        Noun = True
    End If
End Function

Public Function Adjective(sText As String, lPosition As Long, lIndex As Long, vParams As Variant) As Boolean
    Dim lLength As Long

    lLength = PartOfSpeech("Adjectives", sText, lPosition)
    
    If lLength > 0 Then
        lPosition = lPosition + lLength
        Adjective = True
    End If
End Function

Public Function Article(sText As String, lPosition As Long, lIndex As Long, vParams As Variant) As Boolean
    Dim lLength As Long

    lLength = PartOfSpeech("Articles", sText, lPosition)
    
    If lLength > 0 Then
        lPosition = lPosition + lLength
        Article = True
    End If
End Function

Public Function Preposition(sText As String, lPosition As Long, lIndex As Long, vParams As Variant) As Boolean
    Dim lLength As Long

    lLength = PartOfSpeech("Prepositions", sText, lPosition)
    
    If lLength > 0 Then
        lPosition = lPosition + lLength
        Preposition = True
    End If
End Function

Private Sub Class_Initialize()
    oConnection.Open "Provider=Microsoft.Jet.OLEDB.4.0;User ID=Admin;Data Source=" & App.Path & "\Dictionary.mdb;Mode=Share Deny Read|Share Deny Write;Extended Properties="""";Jet OLEDB:System database=C:\PROGRA~1\COMMON~1\System\SYSTEM.MDW;Jet OLEDB:Registry Path="""";Jet OLEDB:Database Password="""";Jet OLEDB:Engine Type=5;Jet OLEDB:Database Locking Mode=0;Jet OLEDB:Global Partial Bulk Ops=2;Jet OLEDB:Global Bulk Transactions=1;Jet OLEDB:New Database Password="""";Jet OLEDB:Create System Database=False;Jet OLEDB:Encrypt Database=False;Jet OLEDB:Don't Copy Locale on Compact=False;Jet OLEDB:Compact Without Replica Repair=False;Jet OLEDB:SFP=False"
End Sub

Private Sub Class_Terminate()
    oConnection.Close
End Sub

Private Function PartOfSpeech(sTable As String, sText As String, lPosition As Long) As Long
    Dim oRS As New Recordset
    Dim lLength As Long
    Dim sWord As String
    Dim lSpace As Long
    
    lSpace = InStr(lPosition, sText, " ")
    
    If lSpace = 0 Then
        sWord = Mid$(sText, lPosition)
    Else
        sWord = Mid$(sText, lPosition, lSpace - lPosition)
    End If
    sWord = UCase$(sWord)
    
    oRS.Open sTable, oConnection, adOpenForwardOnly, , adCmdTableDirect
    
    With oRS
        While Not .EOF
            If sWord = UCase$(!Word) Then
                PartOfSpeech = Len(sWord)
                Exit Function
            End If
            .MoveNext
        Wend
    End With
End Function
