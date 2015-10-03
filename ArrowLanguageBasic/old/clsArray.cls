VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsArray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private mlValues() As Long
Private mlSize As Long
Private moExtents As clsExtents

Public Property Set Extents(ByVal oExtents As clsExtents)
    Set moExtents = oExtents
    ReDim mlValues(moExtents.Size - 1)
End Property

Public Property Get Extents() As clsExtents
    Set Extents = moExtents
End Property
