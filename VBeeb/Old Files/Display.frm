VERSION 5.00
Begin VB.Form Display 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Form1"
   ClientHeight    =   9855
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10095
   ForeColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   9855
   ScaleWidth      =   10095
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   20
      Left            =   600
      Top             =   240
   End
End
Attribute VB_Name = "Display"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function StretchDIBits Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal dx As Long, ByVal dy As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal wSrcWidth As Long, ByVal wSrcHeight As Long, lpBits As Any, lpBitsInfo As BITMAPINFO, ByVal wUsage As Long, ByVal dwRop As Long) As Long

Private Type BITMAPINFOHEADER '40 bytes
        biSize As Long
        biWidth As Long
        biHeight As Long
        biPlanes As Integer
        biBitCount As Integer
        biCompression As Long
        biSizeImage As Long
        biXPelsPerMeter As Long
        biYPelsPerMeter As Long
        biClrUsed As Long
        biClrImportant As Long
End Type

Private Type RGBQUAD
        rgbBlue As Byte
        rgbGreen As Byte
        rgbRed As Byte
        rgbReserved As Byte
End Type

Private Type BITMAPINFO
        bmiHeader As BITMAPINFOHEADER
        bmiColors(1) As RGBQUAD
End Type

Private Const BI_RGB = 0&
Private Const DIB_RGB_COLORS = 0 '  color table in RGBs
Private bmiBuffer As BITMAPINFO
Private Const SRCCOPY = &HCC0020 ' (DWORD) dest = source
Private colors(3) As Long
Private mlMeHDC As Long

Public Sub UpdateDisplay()
    Dim X As Long
'    StartCounter
    StretchDIBits mlMeHDC, 0, 512, 512, -512, 0, 0, 256, 256, DisplayMem(Word(&H1000)), bmiBuffer, DIB_RGB_COLORS, SRCCOPY
'    Debug.Print GetCounter
End Sub

Private Sub Form_Initialize()
    mlMeHDC = Me.hdc
    Initialise
End Sub

Public Sub Initialise()
    colors(0) = vbBlack
    colors(1) = vbWhite
    CreateScreenBuffer
End Sub

Public Sub CreateScreenBuffer()
    bmiBuffer.bmiHeader.biSize = Len(bmiBuffer.bmiHeader)
    bmiBuffer.bmiHeader.biWidth = 256
    bmiBuffer.bmiHeader.biHeight = 256
    bmiBuffer.bmiHeader.biPlanes = 1
    bmiBuffer.bmiHeader.biBitCount = 1
    bmiBuffer.bmiHeader.biCompression = BI_RGB
    bmiBuffer.bmiHeader.biSizeImage = 0
    bmiBuffer.bmiHeader.biXPelsPerMeter = 200
    bmiBuffer.bmiHeader.biYPelsPerMeter = 200
    bmiBuffer.bmiHeader.biClrUsed = 0
    bmiBuffer.bmiHeader.biClrImportant = 0
    
    InitBitmapColors
End Sub

Public Sub InitBitmapColors()
    SetBitmapColor 0, colors(0)
    SetBitmapColor 1, colors(1)
End Sub

Public Sub SetBitmapColor(lIndex As Long, lColor As Long)
    If lIndex > 1 Then Exit Sub
    
    bmiBuffer.bmiColors(lIndex).rgbRed = (lColor And &HFF&)
    bmiBuffer.bmiColors(lIndex).rgbGreen = (lColor And &HFF00&) \ 256
    bmiBuffer.bmiColors(lIndex).rgbBlue = (lColor And &HFF0000) \ 65536
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    Mem(Word(&HFE00)) = KeyAscii
    Processor6502.IRQ
End Sub

Private Sub Timer1_Timer()
    UpdateDisplay
End Sub
