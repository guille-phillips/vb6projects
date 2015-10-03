VERSION 5.00
Begin VB.Form Search 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Memory"
   ClientHeight    =   7920
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   2415
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7920
   ScaleWidth      =   2415
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdAddLocation 
      Caption         =   "Add Location"
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   1320
      Width           =   2175
   End
   Begin VB.CommandButton cmdRemoveLocation 
      Caption         =   "Remove Location"
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   1800
      Width           =   2175
   End
   Begin VB.CheckBox chkDiffers 
      Caption         =   "Differs"
      Height          =   255
      Left            =   1200
      TabIndex        =   7
      Top             =   480
      Width           =   975
   End
   Begin VB.CheckBox chkHex 
      Caption         =   "Hex"
      Height          =   195
      Left            =   120
      TabIndex        =   6
      Top             =   960
      Width           =   855
   End
   Begin VB.CheckBox chkWord 
      Caption         =   "Word"
      Height          =   195
      Left            =   120
      TabIndex        =   5
      Top             =   720
      Width           =   855
   End
   Begin VB.CheckBox chkSigned 
      Caption         =   "Signed"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   480
      Width           =   855
   End
   Begin VB.CommandButton cmdUpdateMemory 
      Caption         =   "Update Memory"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   2280
      Width           =   2175
   End
   Begin VB.ListBox lstMem 
      Height          =   4545
      ItemData        =   "Search.frx":0000
      Left            =   120
      List            =   "Search.frx":0002
      TabIndex        =   2
      Top             =   3240
      Width           =   2175
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear Results"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   2760
      Width           =   2175
   End
   Begin VB.TextBox txtValue 
      Height          =   285
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2175
   End
End
Attribute VB_Name = "Search"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mlMemoryLocations() As MemoryValue
Private mlMemoryLocationsCount As Long

Private Type MemoryValue
    Location As Long
    UnsignedSingle As Long
    SignedSingle As Long
    UnsignedDouble As Long
    SignedDouble As Long
End Type

Private Sub chkHex_Click()
    ReRenderList
End Sub

Private Sub chkSigned_Click()
    ReRenderList
End Sub

Private Sub chkWord_Click()
    ReRenderList
End Sub

Private Sub cmdAddLocation_Click()
    Dim lValue As Long
    Dim lIndex As Long
    Dim lMemIndex As Long
    
    lValue = -1
    
    If UCase$(Right$(txtValue.Text, 1)) = "H" Then
        lValue = ConvertBase(Left$(txtValue.Text, Len(txtValue.Text) - 1), 16)
        txtValue.Text = HexNum(lValue, 4) & "h"
    Else
        lValue = Val(txtValue.Text)
        txtValue.Text = Val(txtValue.Text)
    End If
    
    If lValue <> -1 Then
        For lIndex = 0 To mlMemoryLocationsCount - 1
            If mlMemoryLocations(lIndex).Location = lValue Then
                Exit Sub
            ElseIf mlMemoryLocations(lIndex).Location > lValue Or lIndex = mlMemoryLocationsCount - 1 Then
                lstMem.AddItem HexNum$(lValue, 4) & vbTab & "x", lIndex
                ReDim Preserve mlMemoryLocations(mlMemoryLocationsCount)
                
                mlMemoryLocationsCount = mlMemoryLocationsCount + 1
                For lMemIndex = mlMemoryLocationsCount - 1 To lIndex + 1
                    mlMemoryLocations(lMemIndex) = mlMemoryLocations(lMemIndex - 1)
                Next
                With mlMemoryLocations(lIndex)
                    .Location = lValue
                    .UnsignedSingle = gyMem(lValue)
                    .SignedSingle = gyMem(lValue) + (gyMem(lValue) >= 128&) * 256&
                    .UnsignedDouble = gyMem(lValue + 1&) * 256& + gyMem(lValue)
                    .SignedDouble = .UnsignedDouble + (.UnsignedDouble >= &H8000&) * &H10000
                End With
                Exit Sub
            End If
        Next
        
        lstMem.AddItem HexNum$(lValue, 4) & vbTab & RenderValue(CLng(gyMem(lValue)))
        
        ReDim Preserve mlMemoryLocations(mlMemoryLocationsCount)
        
        mlMemoryLocationsCount = mlMemoryLocationsCount + 1
        With mlMemoryLocations(lIndex)
            .Location = lValue
            .UnsignedSingle = gyMem(lValue)
            .SignedSingle = gyMem(lValue) + (gyMem(lValue) >= 128&) * 256&
            .UnsignedDouble = gyMem(lValue + 1&) * 256& + gyMem(lValue)
            .SignedDouble = .UnsignedDouble + (.UnsignedDouble >= &H8000&) * &H10000
        End With
    End If
End Sub

Private Sub cmdClear_Click()
    lstMem.Clear
    Erase mlMemoryLocations
    mlMemoryLocationsCount = 0
End Sub

Private Sub cmdRemoveLocation_Click()
    Dim lIndex As Long
    
    If lstMem.ListIndex <> -1 Then
        For lIndex = lstMem.ListIndex To mlMemoryLocationsCount - 2
            mlMemoryLocations(lIndex) = mlMemoryLocations(lIndex + 1)
        Next
        lstMem.RemoveItem lstMem.ListIndex
    End If
End Sub

Private Sub cmdUpdateMemory_Click()
    Dim lValue As Long
    
    If UCase$(Right$(txtValue.Text, 1)) = "H" Then
        lValue = ConvertBase(Left$(txtValue.Text, Len(txtValue.Text) - 1), 16)
        txtValue.Text = HexNum(lValue, 4) & "h"
    Else
        lValue = Val(txtValue.Text)
        txtValue.Text = Val(txtValue.Text)
    End If
    
    If lstMem.ListIndex > -1 Then
        If chkWord.Value <> vbChecked Then
            gyMem(mlMemoryLocations(lstMem.ListIndex).Location) = lValue And &HFF&
        Else
            gyMem(mlMemoryLocations(lstMem.ListIndex).Location) = lValue And &HFF&
            gyMem(mlMemoryLocations(lstMem.ListIndex).Location + 1) = (lValue \ &H100&) And &HFF&
        End If
    End If
    
    ReRenderList lstMem.ListIndex
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If UnloadMode <> 1 Then
        Cancel = True
        Me.Hide
    End If
End Sub

Private Sub lstMem_Click()
    If lstMem.ListIndex <> -1 Then
        txtValue.Text = gyMem(mlMemoryLocations(lstMem.ListIndex).Location)
    End If
End Sub

Private Sub txtValue_KeyPress(KeyAscii As Integer)
    Dim lValue As Long
    
    If KeyAscii = 13 Then
        txtValue.Text = Trim$(txtValue.Text)
        If txtValue.Text = "" Then
            SearchAll
        ElseIf UCase$(Right$(txtValue.Text, 1)) = "H" Then
            lValue = ConvertBase(Left$(txtValue.Text, Len(txtValue.Text) - 1), 16)
            txtValue.Text = HexNum(lValue, 4) & "h"
            SearchValue lValue
        Else
            lValue = Val(txtValue.Text)
            txtValue.Text = Val(txtValue.Text)
            SearchValue lValue
        End If
    End If
End Sub

Private Sub SearchAll()
    Dim lLocation As Long
    Dim lIndex As Long
    Dim lMemoryLocationsCopy() As MemoryValue
    Dim lMemoryLocationsCopyCount As Long
    
    Dim lUnsignedSingle As Long
    Dim lSignedSingle As Long
    Dim lSignedDouble As Long
    Dim lUnsignedDouble As Long
    
    Dim bChkDiffers As Boolean
    
    If mlMemoryLocationsCount = 0 Then
        ReDim Preserve mlMemoryLocations(&H7FFF&)
        mlMemoryLocationsCount = &H8000&
        For lLocation = 0 To &H7FFF&
            lUnsignedSingle = gyMem(lLocation)
            lSignedSingle = gyMem(lLocation) + (gyMem(lLocation) >= 128&) * 256&
            lUnsignedDouble = gyMem(lLocation + 1&) * 256& + gyMem(lLocation)
            lSignedDouble = lUnsignedDouble + (lUnsignedDouble >= &H8000&) * &H10000
            
            With mlMemoryLocations(lLocation)
                .Location = lLocation
                .UnsignedSingle = lUnsignedSingle
                .SignedSingle = lSignedSingle
                .UnsignedDouble = lUnsignedDouble
                .SignedDouble = lSignedDouble
            End With
        Next
    Else
        bChkDiffers = chkDiffers.Value = vbChecked
        For lIndex = 0 To mlMemoryLocationsCount - 1
            lUnsignedSingle = gyMem(mlMemoryLocations(lIndex).Location)
            lSignedSingle = gyMem(mlMemoryLocations(lIndex).Location) + (gyMem(mlMemoryLocations(lIndex).Location) >= 128&) * 256&
            lUnsignedDouble = gyMem(mlMemoryLocations(lIndex).Location + 1&) * 256& + gyMem(mlMemoryLocations(lIndex).Location)
            lSignedDouble = lUnsignedDouble + (lUnsignedDouble >= &H8000&) * &H10000

            If bChkDiffers Then
                If mlMemoryLocations(lIndex).UnsignedSingle <> lUnsignedSingle Then
                    ReDim Preserve lMemoryLocationsCopy(lMemoryLocationsCopyCount)
                    With lMemoryLocationsCopy(lMemoryLocationsCopyCount)
                        .Location = mlMemoryLocations(lIndex).Location
                        .UnsignedSingle = lUnsignedSingle
                        .SignedSingle = lSignedSingle
                        .UnsignedDouble = lUnsignedDouble
                        .SignedDouble = lSignedDouble
                    End With
                    
                    lMemoryLocationsCopyCount = lMemoryLocationsCopyCount + 1
                End If
            Else
                If mlMemoryLocations(lIndex).UnsignedSingle = lUnsignedSingle Then
                    ReDim Preserve lMemoryLocationsCopy(lMemoryLocationsCopyCount)
                    lMemoryLocationsCopy(lMemoryLocationsCopyCount) = mlMemoryLocations(lIndex)
                    lMemoryLocationsCopyCount = lMemoryLocationsCopyCount + 1
                End If
            End If
        Next
        
        If lMemoryLocationsCopyCount = 0 Then
            Erase mlMemoryLocations
        Else
            ReDim mlMemoryLocations(lMemoryLocationsCopyCount - 1)
        End If
        For lIndex = 0 To lMemoryLocationsCopyCount - 1
            mlMemoryLocations(lIndex) = lMemoryLocationsCopy(lIndex)
        Next
        mlMemoryLocationsCount = lMemoryLocationsCopyCount
    End If
    
    lstMem.ListIndex = -1
    ReRenderList
End Sub

Private Sub SearchValue(ByVal lValue As Long)
    Dim lLocation As Long
    Dim lIndex As Long
    Dim lMemoryLocationsCopy() As MemoryValue
    Dim lMemoryLocationsCopyCount As Long
    
    Dim lUnsignedSingle As Long
    Dim lSignedSingle As Long
    Dim lSignedDouble As Long
    Dim lUnsignedDouble As Long
    
    If mlMemoryLocationsCount = 0 Then
        For lLocation = 0 To &H7FFF&
            lUnsignedSingle = gyMem(lLocation)
            lSignedSingle = gyMem(lLocation) + (gyMem(lLocation) >= 128&) * 256&
            lUnsignedDouble = gyMem(lLocation + 1&) * 256& + gyMem(lLocation)
            lSignedDouble = lUnsignedDouble + (lUnsignedDouble >= &H8000&) * &H10000
            
            If lValue = lUnsignedSingle Or lValue = lSignedSingle Or lValue = lSignedDouble Or lValue = lUnsignedDouble Then
                ReDim Preserve mlMemoryLocations(mlMemoryLocationsCount)
                mlMemoryLocations(mlMemoryLocationsCount).Location = lLocation
                mlMemoryLocationsCount = mlMemoryLocationsCount + 1
            End If
        Next
    Else
        For lIndex = 0 To mlMemoryLocationsCount - 1
            lUnsignedSingle = gyMem(mlMemoryLocations(lIndex).Location)
            lSignedSingle = gyMem(mlMemoryLocations(lIndex).Location) + (gyMem(mlMemoryLocations(lIndex).Location) >= 128&) * 256&
            lUnsignedDouble = gyMem(mlMemoryLocations(lIndex).Location + 1&) * 256& + gyMem(mlMemoryLocations(lIndex).Location)
            lSignedDouble = lUnsignedDouble + (lUnsignedDouble >= &H8000&) * &H10000

            If lValue = lUnsignedSingle Or lValue = lSignedSingle Or lValue = lSignedDouble Or lValue = lUnsignedDouble Then
                ReDim Preserve lMemoryLocationsCopy(lMemoryLocationsCopyCount)
                lMemoryLocationsCopy(lMemoryLocationsCopyCount) = mlMemoryLocations(lIndex)
                lMemoryLocationsCopyCount = lMemoryLocationsCopyCount + 1
            End If
        Next
        
        If lMemoryLocationsCopyCount = 0 Then
            Erase mlMemoryLocations
        Else
            ReDim mlMemoryLocations(lMemoryLocationsCopyCount - 1)
        End If
        For lIndex = 0 To lMemoryLocationsCopyCount - 1
            mlMemoryLocations(lIndex) = lMemoryLocationsCopy(lIndex)
        Next
        mlMemoryLocationsCount = lMemoryLocationsCopyCount
    End If
    
    lstMem.ListIndex = -1
    ReRenderList

End Sub

Private Sub ReRenderList(Optional ByVal lListIndex As Long = -1)
    Dim lIndex As Long
    
    If lListIndex = -1 Then
        lListIndex = lstMem.ListIndex
        lstMem.Clear
        
        For lIndex = 0 To mlMemoryLocationsCount - 1
            lstMem.AddItem HexNum(mlMemoryLocations(lIndex).Location, 4) & vbTab & RenderValue(gyMem(mlMemoryLocations(lIndex).Location + 1) * 256& + gyMem(mlMemoryLocations(lIndex).Location))
        Next
        lstMem.ListIndex = lListIndex
    Else
        lstMem.List(lListIndex) = HexNum(mlMemoryLocations(lListIndex).Location, 4) & vbTab & RenderValue(gyMem(mlMemoryLocations(lListIndex).Location + 1) * 256& + gyMem(mlMemoryLocations(lListIndex).Location))
    End If
End Sub

Private Function RenderValue(lValue As Long)
    If chkWord.Value <> vbChecked Then
        lValue = lValue And &HFF&
        If chkHex.Value <> vbChecked Then
            If chkSigned.Value = vbChecked Then
                lValue = lValue + (lValue >= 128&) * 256&
            End If
            RenderValue = lValue
        Else
            RenderValue = HexNum(lValue, 2)
        End If
    Else
        If chkHex.Value <> vbChecked Then
            If chkSigned.Value = vbChecked Then
                lValue = lValue + (lValue >= &H8000&) * &H10000
            End If
            RenderValue = lValue
        Else
            RenderValue = HexNum(lValue, 4)
        End If
    End If
End Function
