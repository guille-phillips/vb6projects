VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ISaffronObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' Initialises parameters used for parsing
Public Function Initialise(bNonConsuming As Boolean, bOmit As Boolean, ParamArray pInitParam() As Variant) As ISaffronObject

End Function

' Will perform the parsing function on the object - if parsing fails will return FALSE.
Public Function Parse(oResult As SaffronTree) As Boolean

End Function

' True if the object is to be ommited
Public Property Get Omitted() As Boolean

End Property

' Makes a deep copy
Public Function Copy(Optional ByVal bNonConsuming As Variant, Optional ByVal bOmit As Variant) As ISaffronObject

End Function
