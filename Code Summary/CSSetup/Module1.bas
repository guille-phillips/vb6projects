Attribute VB_Name = "Module1"
Option Explicit

Sub Main()
    Shell "regsvr32 -u -s smalllistview.ocx"
    Shell "regsvr32 -u -s cswindow.dll"
    Shell "regsvr32 -u -s cstoolwindow.dll"
    
    Shell "regsvr32 -s smalllistview.ocx"
    Shell "regsvr32 -s cswindow.dll"
    Shell "regsvr32 -s cstoolwindow.dll"
    Open "c:\winnt\vbaddin.ini" For Append As #1
        Print #1, "CSToolWindow.Connect=0"
    Close #1
    
    MsgBox "Code Summary Installation Completed", vbOKOnly, "Code Summary Install"
End Sub
