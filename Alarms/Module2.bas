Attribute VB_Name = "Module2"
Option Explicit

Public Type GUID
    Data(15) As Byte
End Type

Public Declare Function CoCreateGuid Lib "ole32.dll" (pGuid As GUID) As Long
Public Declare Function StringFromGUID2 Lib "ole32.dll" (rclsid As GUID, ByVal lpsz As Long, ByVal cbMax As Long) As Long

Public Function NewGUID() As String
    Dim uGUID As GUID
    

    CoCreateGuid uGUID
    
    NewGUID = StringFromGUID(uGUID)
End Function

Public Function GUIDFromString(ByVal sGUID As String) As GUID
    Dim lIndex As Long
    
    With GUIDFromString
        For lIndex = 8 To 15
            .Data(lIndex) = CByte("&h" & Mid$(sGUID, lIndex * 2 + 1, 2))
        Next
        
        For lIndex = 6 To 7
            .Data(13 - lIndex) = CByte("&h" & Mid$(sGUID, lIndex * 2 + 1, 2))
        Next
        
        For lIndex = 4 To 5
            .Data(9 - lIndex) = CByte("&h" & Mid$(sGUID, lIndex * 2 + 1, 2))
        Next
        
        For lIndex = 0 To 3
            .Data(3 - lIndex) = CByte("&h" & Mid$(sGUID, lIndex * 2 + 1, 2))
        Next
    End With
End Function

Public Function StringFromGUID(gGUID As GUID) As String
    Dim sGUID As String
    Dim bGUID() As Byte
    Dim lLen As Long
    Dim RetVal As Long
    
    bGUID = String(40, 0)
    
    RetVal = StringFromGUID2(gGUID, VarPtr(bGUID(0)), 40)
    
    sGUID = bGUID
    If (Asc(Mid$(sGUID, RetVal, 1)) = 0) Then RetVal = RetVal - 1
    StringFromGUID = Replace$(Replace$(Replace$(Left$(sGUID, RetVal), "-", ""), "{", ""), "}", "")
End Function

Public Function StepGUID(gGUID As GUID) As GUID
    StepGUID = gGUID
    StepGUID.Data(0) = (gGUID.Data(0) + 1) Mod 256
End Function
