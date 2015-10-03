Attribute VB_Name = "Module2"
Option Explicit

Public Type GUID
    Data(15) As Byte
End Type

Public Declare Function CoCreateGuid Lib "ole32.dll" (pGuid As GUID) As Long
Public Declare Function StringFromGUID2 Lib "ole32.dll" (rclsid As GUID, ByVal lpsz As Long, ByVal cbMax As Long) As Long

Public Function NewGUID() As String
  Dim uGUID As GUID
  Dim sGUID As String
  Dim bGUID() As Byte
  Dim lLen As Long
  Dim RetVal As Long

  bGUID = String(40, 0)

  CoCreateGuid uGUID

  RetVal = StringFromGUID2(uGUID, VarPtr(bGUID(0)), 40)

  sGUID = bGUID
  If (Asc(Mid$(sGUID, RetVal, 1)) = 0) Then RetVal = RetVal - 1
  NewGUID = Replace$(Replace$(Replace$(Left$(sGUID, RetVal), "-", ""), "{", ""), "}", "")
End Function
