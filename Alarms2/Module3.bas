Attribute VB_Name = "Module3"
Option Explicit

Public gsComputerIdentifier As String

Public Type Coord
    X As Long
    Y As Long
End Type

Public Type Dimension
    Width As Long
    Height As Long
End Type

Public Type BoxAbsolute
    NW As Coord
    SE As Coord
End Type

Public Type BoxRelative
    NW As Coord
    Box As Dimension
End Type

Public Type Compass
    W As Long
    N As Long
    E As Long
    S As Long
End Type

Public Type Theme
    BackColour As Long
    TextColour As Long
End Type

Public Enum HAlign
    AlignLeft
    AlignCentre
    AlignRight
End Enum

Public Enum VAlign
    Top
    Centre = 4
    Bottom = 8
End Enum
