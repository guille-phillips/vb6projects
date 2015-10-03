VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Type STARTUPINFO
cb As Long
lpReserved As String
lpDesktop As String
lpTitle As String
dwX As Long
dwY As Long
dwXSize As Long
dwYSize As Long
dwXCountChars As Long
dwYCountChars As Long
dwFillAttribute As Long
dwFlags As Long
wShowWindow As Integer
cbReserved2 As Integer
lpReserved2 As Long
hStdInput As Long
hStdOutput As Long
hStdError As Long
End Type

Private Type PROCESS_INFORMATION
hProcess As Long
hThread As Long
dwProcessID As Long
dwThreadID As Long
End Type

Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal _
hHandle As Long, ByVal dwMilliseconds As Long) As Long

Private Declare Function CreateProcessA Lib "kernel32" (ByVal _
lpApplicationName As String, ByVal lpCommandLine As String, ByVal _
lpProcessAttributes As Long, ByVal lpThreadAttributes As Long, _
ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, _
ByVal lpEnvironment As Long, ByVal lpCurrentDirectory As String, _
lpStartupInfo As STARTUPINFO, lpProcessInformation As _
PROCESS_INFORMATION) As Long

Private Declare Function CloseHandle Lib "kernel32" _
(ByVal hObject As Long) As Long

Private Declare Function GetExitCodeProcess Lib "kernel32" _
(ByVal hProcess As Long, lpExitCode As Long) As Long

Private Const NORMAL_PRIORITY_CLASS = &H20&
Private Const INFINITE = -1&

   
Private Sub Form_Load()
    'CreateImgFiles
    ImageMaker
    'RemoveBBCIM
End Sub

Private Sub CreateImgFiles()
    Dim fso As New FileSystemObject
    Dim oFile As File
    Dim iDot As Integer
    Dim sFileName As String
    Dim sExt As String
    Dim sBBCIMPath As String
    Dim sCommand As String
    Dim oCompanyFolder As Folder
    Dim oFolder As Folder
    Dim proc As PROCESS_INFORMATION
    Dim start As STARTUPINFO
    Dim oFolderName As String
    Dim oGameFolder As Folder
    Dim bContainsIMG As Boolean
    Dim wait As Long
    
    start.cb = Len(start)
      
    Set oGameFolder = fso.GetFolder("c:\emulators\beebem\games")
    
    For Each oCompanyFolder In oGameFolder.SubFolders
        For Each oFolder In oCompanyFolder.SubFolders
            If InStr(oFolder.Name, " (ZIP)") > 0 Then
                oFolderName = Left$(oFolder.Name, InStr(oFolder.Name, " (ZIP") - 1)
                sBBCIMPath = oFolder.Path & "\bbcim.exe"
                bContainsIMG = False
                For Each oFile In oFolder.Files
                    iDot = InStrRev(oFile.Name, ".")
                    If iDot > 0 Then
                        sExt = UCase$(Mid$(oFile.Name, iDot + 1))
                        If sExt = "IMG" Then
                            bContainsIMG = True
                        End If
                    End If
                Next
                If Not bContainsIMG Then
                    For Each oFile In oFolder.Files
                        iDot = InStrRev(oFile.Name, ".")
                        If iDot > 0 Then
                            sFileName = Left$(oFile.Name, iDot - 1)
                            sExt = UCase$(Mid$(oFile.Name, iDot + 1))
                            If sExt = "INF" Then
                                sCommand = " -a " & oFolderName & ".img " & sFileName
                                If CreateProcessA(sBBCIMPath, sCommand, 0&, 0&, 1&, NORMAL_PRIORITY_CLASS, 0&, ByVal oFolder.Path, start, proc) <> 0 Then
                                    WaitForSingleObject proc.hProcess, INFINITE
                                    Call CloseHandle(proc.hThread)
                                    Call CloseHandle(proc.hProcess)
                                    For wait = 0 To 100
                                        DoEvents
                                    Next
                                End If
                            End If
                        End If
                    Next
                End If
            End If
        Next
    Next

End Sub

Private Sub DeleteZips()
    Dim fso As New FileSystemObject
    Dim vFolders As Folders
    Dim vFolder As Folder
    Dim vAFolder As Folder
    Dim oZipFile As File
    Dim sExt As String
    Dim vSubFolder As Folder
    Dim sName As String
    Dim oGameFolder As Folder
    Dim oBBCImFile As File
    Dim iDot As Integer
    Dim sZipName As String
    Dim oCreatedFolder As Folder
    
    Set oGameFolder = fso.GetFolder("c:\emulators\beebem\games")
    
    For Each vFolder In oGameFolder.SubFolders
        For Each oZipFile In vFolder.Files
            iDot = InStr(oZipFile.Name, ".")
            If iDot > 0 Then
                sZipName = Left$(oZipFile.Name, iDot - 1)
                sExt = UCase(Mid$(oZipFile.Name, iDot + 1))
                Select Case sExt
                    Case "ZIP"
                        fso.DeleteFile oZipFile.Path
                End Select
            End If
        Next
    Next

End Sub

Private Sub CopyZips()
    Dim fso As New FileSystemObject
    Dim vFolders As Folders
    Dim vFolder As Folder
    Dim vAFolder As Folder
    Dim oZipFile As File
    Dim sExt As String
    Dim vSubFolder As Folder
    Dim sName As String
    Dim oGameFolder As Folder
    Dim oBBCImFile As File
    Dim iDot As Integer
    Dim sZipName As String
    Dim oCreatedFolder As Folder
    
    Set oGameFolder = fso.GetFolder("c:\emulators\beebem\games")
    Set oBBCImFile = fso.GetFile(oGameFolder.Path & "\bbcim.exe")
    For Each vFolder In oGameFolder.SubFolders
        For Each oZipFile In vFolder.Files
            iDot = InStr(oZipFile.Name, ".")
            If iDot > 0 Then
                sZipName = Left$(oZipFile.Name, iDot - 1)
                sExt = UCase(Mid$(oZipFile.Name, iDot + 1))
                Select Case sExt
                    Case "ZIP"
                        Set oCreatedFolder = fso.CreateFolder(vFolder.Path & "\" & sZipName & " (ZIP)")
                        fso.CopyFile oZipFile.Path, oCreatedFolder.Path & "\" & oZipFile.Name
                        fso.CopyFile oBBCImFile.Path, oCreatedFolder.Path & "\bbcim.exe"
                End Select
            End If
        Next
    Next
End Sub

Private Sub CleanDirectory()
    Dim fso As New FileSystemObject
    Dim vFolders As Folders
    Dim vFolder As Folder
    Dim vAFolder As Folder
    Dim vFile As File
    Dim sExt As String
    Dim vSubFolder As Folder
    Dim sName As String
    Dim oCreatedFolder As Folder
    
    Set oCreatedFolder = fso.CreateFolder("c:\emulators\beebem\Games1")
    Set vFolders = fso.GetFolder("c:\emulators\beebem\games").SubFolders
    For Each vFolder In vFolders
        For Each vSubFolder In vFolder.SubFolders
            For Each vFile In vSubFolder.Files
                If InStr(vFile.Name, ".") > 0 Then
                    sName = Left$(vFile.Name, InStr(vFile.Name, ".") - 1)
                    sExt = UCase(Mid$(vFile.Name, InStr(vFile.Name, ".") + 1))
                    Select Case sExt
                        Case "IMG", "SSD", "DSD"
                            fso.CopyFile vFile.Path, oCreatedFolder.Path & "\" & UCase$(Left$(sName, 1)) & LCase$(Mid$(sName, 2)) & ".img"
                    End Select
                End If
            Next
        Next
    Next

End Sub

Private Sub ImgCopierToDisc()
    Dim fso As New FileSystemObject
    Dim vFolders As Folders
    Dim vFolder As Folder
    Dim vAFolder As Folder
    Dim vFile As File
    Dim sExt As String
    Dim vSubFolder As Folder
    Dim sName As String
    
    Set vFolders = fso.GetFolder("c:\emulators\beebem\games").SubFolders
    For Each vFolder In vFolders
        Set vAFolder = fso.CreateFolder("a:\" & vFolder.Name)
        For Each vFile In vFolder.Files
            sExt = UCase(Mid$(vFile.Name, InStr(vFile.Name, ".") + 1))
            Select Case sExt
                Case "IMG", "SSD", "DSD"
                    fso.CopyFile vFile.Path, vAFolder.Path & "\" & vFile.Name
            End Select
        Next
        For Each vSubFolder In vFolder.SubFolders
            For Each vFile In vSubFolder.Files
                sName = Left$(vFile.Name, InStr(vFile.Name, ".") - 1)
                sExt = UCase(Mid$(vFile.Name, InStr(vFile.Name, ".") + 1))
                Select Case sExt
                    Case "IMG", "SSD", "DSD"
                        fso.CopyFile vFile.Path, vAFolder.Path & "\" & UCase(Left$(sName, 1)) & LCase(Mid$(sName, 2)) & ".img"
                End Select
            Next
        Next
    Next
End Sub

Private Sub DirectoryMaker()
    Dim fso As New FileSystemObject
    Dim vFile As File
    Dim vFiles As Files
    Dim fname As String
    
    Set vFiles = fso.GetFolder("c:\emulators\beebem\discims\superior").Files
    For Each vFile In vFiles
        fname = Left$(vFile.Name, InStr(vFile.Name, ".") - 1)
        If fname <> "bbcim" Then
            fso.CreateFolder "c:\emulators\beebem\discims\superior\" & fname
            fso.MoveFile vFile.Path, "c:\emulators\beebem\discims\superior\" & fname & "\" & vFile.Name
            fso.CopyFile "c:\emulators\beebem\discims\superior\bbcim.exe", "c:\emulators\beebem\discims\superior\" & fname & "\"
        End If
    Next
End Sub

Private Sub DirectoryMaker2()
    Dim fso As New FileSystemObject
    Dim vFile As File
    Dim fname As String
    Dim sExt As String
    Dim vFolder As Folder
    Dim sNewFolder As String
    
    For Each vFolder In fso.GetFolder("c:\emulators\beebem\games").SubFolders
        For Each vFile In vFolder.Files
            fname = UCase$(Left$(vFile.Name, InStr(vFile.Name, ".") - 1))
            sExt = UCase$(Mid$(vFile.Name, InStr(vFile.Name, ".") + 1))
            Select Case sExt
                Case "IMG", "SSD", "DSD"
                    sNewFolder = vFile.ParentFolder.Path & "\" & UCase$(Left$(fname, 1)) & LCase$(Mid$(fname, 2))
                    fso.CreateFolder sNewFolder
                    fso.MoveFile vFile.Path, sNewFolder & "\" & vFile.Name
            End Select
        Next
    Next
End Sub

Private Sub bbcimexecute()
    Dim fso As New FileSystemObject
    Dim vFile As File
    Dim vFiles As Files
    Dim vFolders As Folders
    Dim vFolder As Folder
    Dim fname As String
    
    Set vFolders = fso.GetFolder("c:\emulators\beebem\discims\micro power").SubFolders
    For Each vFolder In vFolders
        If vFolder.Files.Count = 2 Then
            For Each vFile In vFolder.Files
                If vFile.Name <> "bbcim.exe" Then
                    ChDir vFolder.Path
                    Shell "bbcim.exe -e " & vFile.Name
                End If
            Next
        End If
    Next
End Sub

Private Sub ImageMaker()
    Dim fso As New FileSystemObject
    Dim oBaseFolder As Folder
    Dim oCompany As Folder
    Dim oGame As Folder
    Dim sINF As String
    
    Set oBaseFolder = fso.GetFolder("C:\Main\Emulators\Beebem\Games")
    
    For Each oCompany In oBaseFolder.SubFolders
        For Each oGame In oCompany.SubFolders
            ' Does it already contain an IMG SSD or DSD file?
            If Dir(oGame.Path & "\*.img") = "" And Dir(oGame.Path & "\*.ssd") = "" And Dir(oGame.Path & "\*.dsd") = "" Then
                ' Do we have INF files?
                If Dir(oGame.Path & "\*.inf") <> "" Then
                    ' Do we have BBCIM? No then copy it
                    If Dir(oGame.Path & "\bbcim.exe") = "" Then
                        fso.CopyFile "C:\Main\Emulators\Beebem\Games\bbcim.exe", oGame.Path & "\bbcim.exe"
                    End If
                    
                    sINF = Dir(oGame.Path & "\*.inf")
                    While sINF <> ""
                        ExecuteBBCIM oGame.Path & "\bbcim.exe", oGame.Path, oGame.Name, Left$(sINF, Len(sINF) - 4)
                        sINF = Dir
                    Wend
                    fso.DeleteFile oGame.Path & "\bbcim.exe", True
                End If
            End If
        Next
    Next
End Sub

Private Sub ExecuteBBCIM(ByVal sBBCIMPath As String, ByVal sFolderPath As String, ByVal sFolderName As String, ByVal sINFFile As String)
    Dim proc As PROCESS_INFORMATION
    Dim start As STARTUPINFO
    Dim sCommand As String
    Dim wait As Long
    
    start.cb = Len(start)
        
    sCommand = " -a """ & sFolderName & ".img"" " & sINFFile
    If CreateProcessA(sBBCIMPath, sCommand, 0&, 0&, 1&, NORMAL_PRIORITY_CLASS, 0&, ByVal sFolderPath, start, proc) <> 0 Then
        WaitForSingleObject proc.hProcess, INFINITE
        Call CloseHandle(proc.hThread)
        Call CloseHandle(proc.hProcess)
        For wait = 0 To 100
            DoEvents
        Next
    End If
End Sub

Private Sub RemoveBBCIM()
    Dim fso As New FileSystemObject
    Dim oBaseFolder As Folder
    Dim oCompany As Folder
    Dim oGame As Folder
    Dim sINF As String
    
    Set oBaseFolder = fso.GetFolder("C:\Main\Emulators\Beebem\Games")
    
    For Each oCompany In oBaseFolder.SubFolders
        For Each oGame In oCompany.SubFolders
            If Dir(oGame.Path & "\bbcim.exe") <> "" Then
                fso.DeleteFile oGame.Path & "\bbcim.exe", True
            End If
        Next
    Next
End Sub
