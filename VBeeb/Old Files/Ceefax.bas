Attribute VB_Name = "Ceefax"
Option Explicit

Public Sub LoadPage(ByVal lPage As Long)
    Dim sText As String
    
    sText = Console.Inet1.OpenURL("http://www.ceefax.tv/txtmaster.php?page=100&subpage=0&channel=bbc1")

End Sub
