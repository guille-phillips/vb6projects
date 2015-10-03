VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Storage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type fiFile
    Name As String
    Load As Long
    Execution As Long
    Length As Long
    Locked As Boolean
    CheckSum As Long
    Data() As Byte
End Type

Private mfiCatalogue() As fiFile
Private mlCatalogueCount As Long
Private msName As String

Private Sub AddFile(fiAddFile As fiFile)
    ReDim Preserve mfiCatalogue(mlCatalogueCount)
    mfiCatalogue(mlCatalogueCount) = fiAddFile
    mlCatalogueCount = mlCatalogueCount + 1
End Sub

Public Sub FromDisc(yImage() As Byte)

End Sub

Public Sub FromCassette(yImage() As Byte)

End Sub

Public Sub SaveDiscImage()

End Sub

Public Sub SaveCassetteImage()

End Sub

Public Sub SaveArchiveFormat()

End Sub
