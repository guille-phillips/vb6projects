VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H80000005&
   Caption         =   "Calculator"
   ClientHeight    =   4185
   ClientLeft      =   165
   ClientTop       =   525
   ClientWidth     =   4860
   Icon            =   "Form1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   279
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   324
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Entry 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3855
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   4575
   End
   Begin VB.Menu mnuFunctions 
      Caption         =   "Functions"
      Visible         =   0   'False
      Begin VB.Menu mnuFunc 
         Caption         =   "SIN"
         Index           =   0
      End
      Begin VB.Menu mnuFunc 
         Caption         =   "COS"
         Index           =   1
      End
      Begin VB.Menu mnuFunc 
         Caption         =   "TAN"
         Index           =   2
      End
      Begin VB.Menu mnuFunc 
         Caption         =   "SQR"
         Index           =   3
      End
      Begin VB.Menu mnuFunc 
         Caption         =   "DEG"
         Index           =   4
      End
      Begin VB.Menu mnuFunc 
         Caption         =   "RAD"
         Index           =   5
      End
      Begin VB.Menu mnuFunc 
         Caption         =   "EXP"
         Index           =   6
      End
      Begin VB.Menu mnuFunc 
         Caption         =   "LOG"
         Index           =   7
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
    
Private oVariables As New Dictionary
Private oStatement As IParseObject

Private Enum FunctionTypes
    ftLoop = 1
    ftFactorial
    ftPercent
    ftTime
    ftRadixNumber
    ftNumber
    ftFunctionBitExpression
    ftFunctionExpressionX
    ftConstant
    ftFunctionVariableCall
    ftUnaryExpression
    ftBracketExpression
End Enum

Private ErrorSource As String

Private Sub Form_Load()
    Dim sDefinition As String
    Dim oFS As New FileSystemObject
    
    Me.Width = GetSetting("Calculator", "Dimensions", "Width", Me.Width)
    Me.Height = GetSetting("Calculator", "Dimensions", "Height", Me.Height)
    
    sDefinition = oFS.OpenTextFile(App.Path & "\calculator.pdl").ReadAll

    If Not SetNewDefinition(sDefinition) Then
        Debug.Print ErrorString
        Stop
    End If
    
    Set oStatement = ParserObjects("level0")
    
    ReadDefinitions
End Sub

Private Sub Entry_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = vbRightButton Then
        Me.PopupMenu mnuFunctions
    End If
End Sub

Private Sub Form_Activate()
    SetWindowPos Me.hwnd, -1, 0, 0, 0, 0, 3
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    On Error GoTo ShowError
    
    Dim sCommand As String
    Dim vLines As Variant
    Dim oResult As New ParseTree
    Dim lCursorPos As Long
    Dim answer As String
    
    Select Case KeyAscii
        Case 13
            lCursorPos = Entry.SelStart
            vLines = Split(Left(Entry.Text, lCursorPos), vbCrLf)
            If UBound(vLines) <> -1 Then
                sCommand = vLines(UBound(vLines))
            End If
            KeyAscii = 0
            Stream.Text = sCommand
            If oStatement.Parse(oResult) Then
                If Stream.Position = Len(sCommand) + 1 Then
                    answer = EvaluateLevel0(oResult)
                    If answer <> "" Then
                        InsertText vbCrLf & answer & vbCrLf & vbCrLf
                    Else
                        InsertText vbCrLf
                    End If
                Else
                    InsertText vbCrLf & "Syntax Error" & vbCrLf & vbCrLf
                End If
            Else
                If Trim$(sCommand) <> "" Then
                    InsertText vbCrLf & "Syntax Error" & vbCrLf & vbCrLf
                Else
                    InsertText vbCrLf
                End If
            End If
        Case 27
            Unload Me
    End Select
    
    Exit Sub
    
ShowError:
    If Err.Description <> "" Then
        InsertText vbCrLf & Err.Description & vbCrLf & vbCrLf
    Else
        InsertText vbCrLf & "Error" & vbCrLf & vbCrLf
    End If
    Err.Clear
    ErrorSource = ""
End Sub


Private Function EvaluateLevel0(ByVal oExpression As ParseTree, Optional oParameters As Dictionary) As String
    Dim runningvalue As String
    Dim vPart As Variant
    Dim lPartIndex As Long
    Dim tempvalue As String
    
    On Error GoTo ExitPoint
    Dim sErrorSource As String
    
    Select Case oExpression.Index
        Case 1
            For Each vPart In oExpression(1)(1).SubTree
                runningvalue = EvaluateLevel0(vPart, oParameters)
            Next
        Case 2
            runningvalue = EvaluateLevel1(oExpression(1), oParameters)
    End Select
    
    EvaluateLevel0 = runningvalue
    
    Exit Function
ExitPoint:
    Err.Raise -1, , ErrorHandler(sErrorSource)
End Function

Private Function EvaluateLevel1(ByVal oExpression As ParseTree, Optional oParameters As Dictionary) As String
    On Error GoTo ExitPoint
    Dim sErrorSource As String
    
    Dim newvar As New Variable
    
    Select Case oExpression.Index
        Case 1 ' Keyword
            Select Case LCase$(oExpression.Text)
                Case "list"
                    ListDefinitions
            End Select
        Case 2 ' Value assignment
            If oVariables.Exists(oExpression(1)(1).Text) Then
                newvar.Value = CStr(EvaluateLevel0(oExpression(1)(3)))
                oVariables.Remove CStr(oExpression(1)(1).Text)
                oVariables.Add CStr(oExpression(1)(1).Text), newvar
            Else
                newvar.Value = CStr(EvaluateLevel0(oExpression(1)(3)))
                oVariables.Add CStr(oExpression(1)(1).Text), newvar
            End If
        
        Case 3 ' Function Assigment
            Dim sVariableName As String
            sVariableName = oExpression(1)(1)(1).Text
            If oVariables.Exists(sVariableName) Then
                oVariables.Remove CStr(sVariableName)
            End If
            
            ' Check Expression
            Dim sExpression As String
            Dim oLevel0 As IParseObject
            Dim oResult As New ParseTree
            
            Set oLevel0 = ParserObjects("level0")
            sExpression = oExpression(1)(3).Text
            Stream.Text = sExpression
        
            If Not oLevel0.Parse(oResult) Then
                Err.Raise -1
                Exit Function
            ElseIf Stream.Position <> Len(sExpression) + 1 Then
                Err.Raise -1
                Exit Function
            End If
            
            newvar.Expression = oExpression(1)(3).Text
            
            ' Any parameters?
            Dim oParameterVar As Variable
            Dim oParameter As ParseTree
            If oExpression(1)(1)(2).Index > 0 Then
                Set newvar.Parameters = New Dictionary
                For Each oParameter In oExpression(1)(1)(2)(1)(1).SubTree
                    Set oParameterVar = New Variable
                    newvar.Parameters.Add CStr(oParameter.Text), oParameterVar
                Next
            End If
        
            oVariables.Add CStr(sVariableName), newvar
                
        Case 4
            EvaluateLevel1 = EvaluateLevel2(oExpression(1), oParameters)
    End Select
    Exit Function
ExitPoint:
    Err.Raise -1, , ErrorHandler(sErrorSource)
End Function

Private Function EvaluateLevel2(ByVal oExpression As ParseTree, Optional oParameters As Dictionary) As String
    Dim runningvalue As String
    Dim vPart As Variant
    Dim lPartIndex As Long
    Dim tempvalue As String
    
    On Error GoTo ExitPoint
    Dim sErrorSource As String
    
    runningvalue = EvaluateLevel3(oExpression(1), oParameters)
    For lPartIndex = 2 To oExpression.SubTree.Count - 1 Step 2
        Set vPart = oExpression(lPartIndex)
        Select Case UCase(vPart(1).Text)
            Case ">"
                runningvalue = IIf(Val(runningvalue) > Val(EvaluateLevel3(oExpression(lPartIndex + 1), oParameters)), 1, 0)
            Case "<"
                runningvalue = IIf(Val(runningvalue) < Val(EvaluateLevel3(oExpression(lPartIndex + 1), oParameters)), 1, 0)
            Case "=="
                runningvalue = IIf(Val(runningvalue) = Val(EvaluateLevel3(oExpression(lPartIndex + 1), oParameters)), 1, 0)
            Case "<>"
                runningvalue = IIf(Val(runningvalue) <> Val(EvaluateLevel3(oExpression(lPartIndex + 1), oParameters)), 1, 0)
            Case ">="
                runningvalue = IIf(Val(runningvalue) >= Val(EvaluateLevel3(oExpression(lPartIndex + 1), oParameters)), 1, 0)
            Case "<="
                runningvalue = IIf(Val(runningvalue) <= Val(EvaluateLevel3(oExpression(lPartIndex + 1), oParameters)), 1, 0)
        End Select
    Next
    EvaluateLevel2 = runningvalue
    
    Exit Function
ExitPoint:
    Err.Raise -1, , ErrorHandler(sErrorSource)
End Function

Private Function EvaluateLevel3(ByVal oExpression As ParseTree, Optional oParameters As Dictionary) As String
    Dim runningvalue As String
    Dim vPart As Variant
    Dim lPartIndex As Long
    Dim tempvalue As String
    
    On Error GoTo ExitPoint
    Dim sErrorSource As String
    
    runningvalue = EvaluateLevel4(oExpression(1), oParameters)
    For lPartIndex = 2 To oExpression.SubTree.Count - 1 Step 2
        Set vPart = oExpression(lPartIndex)
        Select Case UCase(vPart(1).Text)
            Case "AND"
                runningvalue = runningvalue And EvaluateLevel4(oExpression(lPartIndex + 1), oParameters)
            Case "OR"
                runningvalue = runningvalue Or EvaluateLevel4(oExpression(lPartIndex + 1), oParameters)
            Case "XOR"
                runningvalue = runningvalue Xor EvaluateLevel4(oExpression(lPartIndex + 1), oParameters)
            Case "MOD"
                tempvalue = EvaluateLevel4(oExpression(lPartIndex + 1), oParameters)
                runningvalue = ((runningvalue / tempvalue) - Int(runningvalue / tempvalue)) * tempvalue
        End Select
    Next
    
    EvaluateLevel3 = runningvalue
    
    Exit Function
ExitPoint:
    Err.Raise -1, , ErrorHandler(sErrorSource)
End Function

Private Function EvaluateLevel4(ByVal oExpression As ParseTree, Optional oParameters As Dictionary) As String
    Dim runningvalue As String
    Dim vPart As Variant
    Dim lPartIndex As Long
    
    On Error GoTo ExitPoint
    Dim sErrorSource As String
    
    runningvalue = EvaluateLevel5(oExpression(1), oParameters)
    For lPartIndex = 2 To oExpression.SubTree.Count - 1 Step 2
        Set vPart = oExpression(lPartIndex)
        Select Case UCase(vPart(1).Text)
            Case "+"
                runningvalue = runningvalue + CDbl(EvaluateLevel5(oExpression(lPartIndex + 1), oParameters))
            Case "-"
                runningvalue = runningvalue - CDbl(EvaluateLevel5(oExpression(lPartIndex + 1), oParameters))
            Case "&"
                runningvalue = runningvalue & CDbl(EvaluateLevel5(oExpression(lPartIndex + 1), oParameters))
        End Select
    Next
    
    EvaluateLevel4 = runningvalue
        
    Exit Function
ExitPoint:
    Err.Raise -1, , ErrorHandler(sErrorSource)
End Function

Private Function EvaluateLevel5(ByVal oExpression As ParseTree, Optional oParameters As Dictionary) As String
    Dim runningvalue As String
    Dim vPart As Variant
    Dim lPartIndex As Long
    
    On Error GoTo ExitPoint
    Dim sErrorSource As String
    
    runningvalue = EvaluateLevel6(oExpression(1), oParameters)
    For lPartIndex = 2 To oExpression.SubTree.Count - 1 Step 2
        Set vPart = oExpression(lPartIndex)
        Select Case UCase(vPart(1).Text)
            Case "*"
                runningvalue = runningvalue * EvaluateLevel6(oExpression(lPartIndex + 1), oParameters)
            Case "/"
                sErrorSource = "DIVIDE"
                runningvalue = runningvalue / EvaluateLevel6(oExpression(lPartIndex + 1), oParameters)
            Case "\"
                runningvalue = runningvalue \ EvaluateLevel6(oExpression(lPartIndex + 1), oParameters)
        End Select
    Next
    
    EvaluateLevel5 = runningvalue
        
    Exit Function
ExitPoint:
    Err.Raise -1, , ErrorHandler(sErrorSource)
End Function

Private Function EvaluateLevel6(ByVal oExpression As ParseTree, Optional oParameters As Dictionary) As String
    Dim runningvalue As String
    Dim vPart As Variant
    Dim lPartIndex As Long
    
    On Error GoTo ExitPoint
    Dim sErrorSource As String
    
    runningvalue = EvaluateLevel7(oExpression(1), oParameters)
    For lPartIndex = 2 To oExpression.SubTree.Count - 1 Step 2
        Set vPart = oExpression(lPartIndex)
        Select Case UCase(vPart(1).Text)
            Case "^"
                runningvalue = runningvalue ^ EvaluateLevel7(oExpression(lPartIndex + 1), oParameters)
        End Select
    Next
        
    EvaluateLevel6 = runningvalue
        
    Exit Function
ExitPoint:
    Err.Raise -1, , ErrorHandler(sErrorSource)
End Function

Private Function EvaluateLevel7(ByVal oExpression As ParseTree, Optional oParameters As Dictionary) As String
    Dim runningvalue As String
    Dim runningvalue1 As String
    Dim runningvalue2 As String
    Dim runningvalue3 As String
    Dim vPart As Variant
    Dim temp As Double
    
    Dim hour As Double
    Dim minute As Integer
    Dim second As Integer
    Dim fraction As Double
    Dim sign As Long
    
    Dim iDirection As Integer
    Dim oSubExpression As ParseTree
    
    On Error GoTo ExitPoint
    Dim sErrorSource As String
    
    Select Case oExpression.Index
        Case ftLoop
            Dim iLower As Double
            Dim iUpper As Double
            Dim sOperator As String
            Dim sVariable As String
            Dim iLoopIndex As Double
            Dim bShowResult As Boolean
            Dim oVariable As Variable
            
            iLower = EvaluateLevel0(oExpression(1)(2)(1), oParameters)
            iUpper = EvaluateLevel0(oExpression(1)(2)(2), oParameters)
            sOperator = oExpression(1)(1)(2).Text
            sVariable = oExpression(1)(4).Text
            Select Case LCase$(oExpression(1)(1)(1).Text)
                Case "loop"
                    bShowResult = False
                Case "show"
                    bShowResult = True
            End Select
                    
            If oVariables.Exists(sVariable) Then
                oVariables.Remove sVariable
            End If
            
            Set oVariable = New Variable
            
            oVariables.Add sVariable, oVariable
            
            Select Case sOperator
                Case "+", "-"
                    runningvalue = 0
                Case "*", "/"
                    runningvalue = 1
                Case "&"
                    runningvalue = ""
            End Select
            
            If iUpper < iLower Then
                iDirection = -1
            Else
                iDirection = 1
            End If
            For iLoopIndex = iLower To iUpper Step iDirection
                oVariables.Item(sVariable).Value = iLoopIndex
                Select Case sOperator
                    Case ""
                        runningvalue = EvaluateLevel0(oExpression(1)(3), oParameters)
                    Case "+"
                        runningvalue = Val(runningvalue) + EvaluateLevel0(oExpression(1)(3), oParameters)
                    Case "-"
                        runningvalue = Val(runningvalue) - EvaluateLevel0(oExpression(1)(3), oParameters)
                    Case "*"
                        runningvalue = Val(runningvalue) * EvaluateLevel0(oExpression(1)(3), oParameters)
                    Case "/"
                        runningvalue = Val(runningvalue) / EvaluateLevel0(oExpression(1)(3), oParameters)
                    Case "&"
                        runningvalue = runningvalue & EvaluateLevel0(oExpression(1)(3), oParameters)
                End Select
                If bShowResult Then
                    InsertText vbCrLf & runningvalue
                End If
            Next
            If bShowResult Then
                runningvalue = ""
            End If
            
        Case ftFactorial
            runningvalue = Factorial(EvaluateLevel7(oExpression(1)(1), oParameters))
        
        Case ftPercent
            runningvalue = CDbl(oExpression.Text) / 100
        
        Case ftTime
            Select Case oExpression(1)(1).Name
                Case "minsec"
                    sign = (oExpression(1)(1)(1).Text = "-") * 2 + 1
                    hour = oExpression(1)(1)(2).Text
                    minute = oExpression(1)(1)(4).Text
                    If oExpression(1)(1)(5).Index = 1 Then
                        fraction = "0" & oExpression(1)(1)(5).Text
                    End If
                    runningvalue = sign * (hour + (minute + fraction) / 60)
                Case "hourminsec"
                    sign = (oExpression(1)(1)(1).Text = "-") * 2 + 1
                    hour = oExpression(1)(1)(2).Text
                    minute = oExpression(1)(1)(4).Text
                    second = oExpression(1)(1)(6).Text
                    If oExpression(1)(1)(7).Index = 1 Then
                        fraction = "0" & oExpression(1)(1)(7).Text
                    End If
                    runningvalue = sign * (hour + minute / 60 + (second + fraction) / 3600)
                
            End Select
            
        Case ftRadixNumber
            runningvalue = Base(oExpression(1)(1).Text, CLng(oExpression(1)(3).Text), 10)

        Case ftNumber
            sErrorSource = "NUMBER"
            runningvalue = CDbl(oExpression.Text)
            
        Case ftFunctionExpressionX
            Dim iNoOfParams As Integer
            iNoOfParams = oExpression(1)(2).Index
            sErrorSource = "FUNCTION"
            
            Select Case UCase$(oExpression(1)(1).Text)
                Case "SIN", "COS", "TAN", "EXP", "LOG", "DEG", "RAD", "INT", "FRAC", "SQR", "NOT", "ATN", "ACS", "ASN", "GAM", "COT", "CSC", "SEC", "DMS", "LEN"
                    If iNoOfParams <> 1 Then
                        Err.Raise -1
                    End If
            End Select
            Set oSubExpression = oExpression(1)(2)(1)
            
            Select Case UCase$(oExpression(1)(1).Text)
                Case "SIN"
                    runningvalue = Sin(EvaluateLevel0(oSubExpression, oParameters))
                Case "COS"
                    runningvalue = Cos(EvaluateLevel0(oSubExpression, oParameters))
                Case "TAN"
                    runningvalue = Tan(EvaluateLevel0(oSubExpression, oParameters))
                Case "EXP"
                    runningvalue = Exp(EvaluateLevel0(oSubExpression, oParameters))
                Case "LOG"
                    runningvalue = Log(EvaluateLevel0(oSubExpression, oParameters))
                Case "DEG"
                    runningvalue = 360 * (EvaluateLevel0(oSubExpression, oParameters)) / 8 / Atn(1)
                Case "RAD"
                    runningvalue = (EvaluateLevel0(oSubExpression, oParameters)) * 8 * Atn(1) / 360
                Case "INT"
                    runningvalue = Int(EvaluateLevel0(oSubExpression, oParameters))
                Case "FRAC"
                    temp = EvaluateLevel0(oSubExpression, oParameters)
                    runningvalue = temp - Int(temp)
                Case "SQR"
                    sErrorSource = "SQR"
                    runningvalue = Sqr(EvaluateLevel0(oSubExpression, oParameters))
                Case "NOT"
                    runningvalue = Not (EvaluateLevel0(oSubExpression, oParameters))
                Case "ATN"
                    runningvalue = Atn(EvaluateLevel0(oSubExpression, oParameters))
                Case "ACS"
                    temp = (EvaluateLevel0(oSubExpression, oParameters))
                    runningvalue = 2 * Atn(1) - Atn(temp / Sqr(1 - temp * temp))
                Case "ASN"
                    temp = (EvaluateLevel0(oSubExpression, oParameters))
                    runningvalue = Atn(temp / Sqr(1 - temp * temp))
                Case "GAM"
                    '   G[z] = Integral[t^(z-1) Exp[-t] dt, {t, 0, Infinity}]
                    Dim oldrunningvalue As Double
                    Dim t As Double
                    temp = (EvaluateLevel0(oSubExpression, oParameters))
                    runningvalue = 0
                    oldrunningvalue = -1
                    t = 0.0001
                    While oldrunningvalue <> runningvalue
                        oldrunningvalue = runningvalue
                        runningvalue = runningvalue + CDbl((t ^ temp * Exp(-t)) * 0.0001)
                        t = t + 0.0001
                    Wend
                Case "COT"
                    runningvalue = 1 / Tan(EvaluateLevel0(oSubExpression, oParameters))
                Case "CSC"
                    runningvalue = 1 / Sin(EvaluateLevel0(oSubExpression, oParameters))
                Case "SEC"
                    runningvalue = 1 / Cos(EvaluateLevel0(oSubExpression, oParameters))
                Case "DMS"
                    runningvalue = EvaluateLevel0(oSubExpression, oParameters)
                    hour = Int(runningvalue) Mod 24
                    hour = IIf(runningvalue < 0, 24 * (1 - (runningvalue + 1) \ 24) + runningvalue, (Int(runningvalue)) Mod 24)
                    runningvalue = Format(hour, "00") & ":" & Format((runningvalue - Int(runningvalue)) * 60, "00")
                Case "DM"
                    runningvalue = EvaluateLevel0(oSubExpression, oParameters)
                    runningvalue = Format(Int(runningvalue), "00") & ":" & PadNumber((runningvalue - Int(runningvalue)) * 60, 1)
                Case "LEN"
                    runningvalue = EvaluateLevel0(oSubExpression, oParameters)
                    If Int(runningvalue) = runningvalue Then
                        runningvalue = Len(CStr(runningvalue))
                    Else
                        runningvalue = 0
                    End If
                Case "IF"
                    runningvalue = Int(EvaluateLevel0(oSubExpression, oParameters))
                    If runningvalue < 0 Or runningvalue > oExpression(1)(2).Index Then
                        Err.Raise -1
                    Else
                        runningvalue = EvaluateLevel0(oExpression(1)(2)(Val(runningvalue) + 2), oParameters)
                    End If
                    
                Case Else
                    If Left(UCase(oExpression(1)(1).Text), 3) = "LOG" Then
                        runningvalue = Log(EvaluateLevel0(oSubExpression, oParameters)) / Log(oExpression(1)(1)(1)(2).Text)
                    End If
                    If Left(UCase(oExpression(1)(1).Text), 5) = "RADIX" Then
                        runningvalue = Base(EvaluateLevel0(oSubExpression, oParameters), 10, Val(oExpression(1)(1)(1)(2).Text))
                    End If
                    If Left(UCase(oExpression(1)(1).Text), 3) = "FIX" Then
                        Dim iDecimalPlaces As Long
                        iDecimalPlaces = Val(oExpression(1)(1)(1)(2).Text)
                        runningvalue = Int(EvaluateLevel0(oSubExpression, oParameters) * 10 ^ iDecimalPlaces + 0.5) / 10 ^ iDecimalPlaces
                    End If
                    If iNoOfParams = 1 Then
                    Else
                        Err.Raise -1
                    End If
            End Select
            
'Secant Sec(X) = 1 / Cos(X)
'Cosecant Cosec(X) = 1 / Sin(X)
'Cotangent Cotan(X) = 1 / Tan(X)
'Inverse Sine Arcsin(X) = Atn(X / Sqr(-X * X + 1))
'Inverse Cosine Arccos(X) = Atn(-X / Sqr(-X * X + 1)) + 2 * Atn(1)
'Inverse Secant Arcsec(X) = 2 * Atn(1) – Atn(Sgn(X) / Sqr(X * X – 1))
'Inverse Cosecant Arccosec(X) = Atn(Sgn(X) / Sqr(X * X – 1))
'Inverse Cotangent Arccotan(X) = 2 * Atn(1) - Atn(X)
'Hyperbolic Sine HSin(X) = (Exp(X) – Exp(-X)) / 2
'Hyperbolic Cosine HCos(X) = (Exp(X) + Exp(-X)) / 2
'Hyperbolic Tangent HTan(X) = (Exp(X) – Exp(-X)) / (Exp(X) + Exp(-X))
'Hyperbolic Secant HSec(X) = 2 / (Exp(X) + Exp(-X))
'Hyperbolic Cosecant HCosec(X) = 2 / (Exp(X) – Exp(-X))
'Hyperbolic Cotangent HCotan(X) = (Exp(X) + Exp(-X)) / (Exp(X) – Exp(-X))
'Inverse Hyperbolic Sine HArcsin(X) = Log(X + Sqr(X * X + 1))
'Inverse Hyperbolic Cosine HArccos(X) = Log(X + Sqr(X * X – 1))
'Inverse Hyperbolic Tangent HArctan(X) = Log((1 + X) / (1 – X)) / 2
'Inverse Hyperbolic Secant HArcsec(X) = Log((Sqr(-X * X + 1) + 1) / X)
'Inverse Hyperbolic Cosecant HArccosec(X) = Log((Sgn(X) * Sqr(X * X + 1) + 1) / X)
'Inverse Hyperbolic Cotangent HArccotan(X) = Log((X + 1) / (X – 1)) / 2
'Logarithm to base N LogN(X) = Log(X) / Log(N)
        
        Case ftFunctionBitExpression
            runningvalue1 = CStr(EvaluateLevel0(oExpression(1)(2), oParameters))
            For Each vPart In oExpression(1)(3).SubTree
                Select Case vPart.Index
                    Case 2 ' single
                        runningvalue2 = EvaluateLevel0(vPart(1), oParameters)
                        runningvalue = runningvalue & Mid$(runningvalue1, runningvalue2, 1)
                    Case 1 ' range
                        runningvalue2 = EvaluateLevel0(vPart(1)(1)(1), oParameters)
                        runningvalue3 = EvaluateLevel0(vPart(1)(2)(1), oParameters)
                        runningvalue = runningvalue & Mid$(runningvalue1, runningvalue2, runningvalue3 - runningvalue2 + 1)
                End Select
            Next
            
        Case ftConstant
            Select Case UCase(oExpression.Text)
                Case "PI"
                    runningvalue = Atn(1) * 4
                Case "E"
                    runningvalue = Exp(1)
            End Select
            
        Case ftFunctionVariableCall
            Dim sVariableName As String
            Dim iParameterIndex As Long
            Dim oParameterValue As ParseTree
            Dim oParameterVariable As Dictionary
            
            sErrorSource = "VARIABLE"
            
            sVariableName = oExpression(1)(1).Text

            If oExpression(1)(2).Index = 0 Then ' Has no parameters
                If Not oParameters Is Nothing Then
                    If oParameters.Exists(sVariableName) Then
                        runningvalue = oParameters.Item(sVariableName).Value
                        EvaluateLevel7 = runningvalue
                        Exit Function
                    End If
                End If
            End If
            
            If oVariables.Exists(sVariableName) Then
                If oExpression(1)(2).Index = 0 Then
                    If Not oVariables.Item(sVariableName).Parameters Is Nothing Then
                        Err.Raise -1
                    End If
                Else
                    If oExpression(1)(2)(1)(1).Index <> oVariables.Item(sVariableName).Parameters.Count Then
                        Err.Raise -1
                    End If

                    iParameterIndex = 0
                    Set oParameterVariable = oVariables.Item(sVariableName).Parameters
                    For Each oParameterValue In oExpression(1)(2)(1)(1).SubTree
                        oParameterVariable.Items(iParameterIndex).Value = EvaluateLevel0(oParameterValue)
                        iParameterIndex = iParameterIndex + 1
                    Next
                End If
                
                Dim thevar As Variable
                Set thevar = oVariables.Item(sVariableName)
                If thevar.Expression = "" Then
                    runningvalue = thevar.Value
                Else
                    Dim subdecode As IParseObject
                    Dim othisResult As New ParseTree
                    Dim savetext As String
                    Dim savepos As String

                    savetext = Stream.Text
                    savepos = Stream.Position

                    Set subdecode = ParserObjects("level0")

                    Stream.Text = thevar.Expression
                    subdecode.Parse othisResult
                    runningvalue = EvaluateLevel0(othisResult, oVariables.Item(sVariableName).Parameters)
                    Stream.Text = savetext
                    Stream.Position = savepos
                End If
            Else
                sErrorSource = "VARIABLE"
                Err.Raise -3
            End If

        Case ftUnaryExpression
            Select Case oExpression(1)(1).Text
                Case "+"
                    runningvalue = EvaluateLevel0(oExpression(1)(2), oParameters)

                Case "-"
                    runningvalue = -EvaluateLevel0(oExpression(1)(2), oParameters)
            End Select
            
        Case ftBracketExpression
            runningvalue = EvaluateLevel0(oExpression(1)(2), oParameters)
            
    End Select
    
    EvaluateLevel7 = runningvalue
    
    Exit Function
ExitPoint:
    Err.Raise -1, , ErrorHandler(sErrorSource)
End Function

Private Function PadNumber(dNum As Double, lPlaces As Long) As String
    Dim lZeroes As Long
    
    If dNum > 0 Then
        lZeroes = lPlaces - Int(Log(dNum) / Log(10))
    ElseIf dNum < 0 Then
    Else
        lZeroes = lPlaces
    End If
        
    PadNumber = String(lZeroes, "0") & dNum
End Function

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    If Button = vbRightButton Then
        Me.PopupMenu mnuFunctions
    End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    SaveSetting "Calculator", "Dimensions", "Width", Me.Width
    SaveSetting "Calculator", "Dimensions", "Height", Me.Height
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    Entry.Width = Me.ScaleWidth - 16
    Entry.Height = Me.ScaleHeight - 16
End Sub

Private Function InsertText(ByVal sText As String)
    Dim iTextPos As Integer
    
    iTextPos = Entry.SelStart
    Entry.Text = Left$(Entry.Text, iTextPos) & sText & Mid$(Entry.Text, iTextPos + 1)
    
    Entry.SelStart = iTextPos + Len(sText)
End Function

Private Function Factorial(ByVal lnum As Double) As Double
    On Error GoTo ExitFunction
    Dim sErrorSource As String
    
    Dim q As Long
    
    sErrorSource = "FACTORIAL"
    
    If (lnum < 0) Or (Int(lnum) <> lnum) Then
        Err.Raise -1
    End If
    
    Factorial = 1
    For q = 2 To lnum
        Factorial = Factorial * q
    Next
    Exit Function
ExitFunction:
    Err.Raise -1, , ErrorHandler(sErrorSource)
End Function

Private Function Factor(ByVal lnum As Double) As String
    Dim f As Double
    lnum = Int(lnum)
    f = 2
    While lnum > 1
        If (lnum / f) = Int(lnum / f) Then
            Factor = Factor & CStr(f) & " "
            lnum = lnum / Factor
            f = 2
        Else
            f = f + 1
        End If
    Wend
End Function

Private Function Base(ByVal lnum As String, ByVal frombase As Integer, ByVal tobase As Integer) As String
    Const digits As String = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    Dim digitpos As Long
    Dim power As Long
    Dim result As Long
    Dim digitvalue As Long
    
    power = 1
    For digitpos = Len(lnum) To 1 Step -1
        digitvalue = InStr(digits, Mid$(lnum, digitpos, 1)) - 1
        result = result + digitvalue * power
        power = power * frombase
    Next
    
    power = tobase
    While result > 0
        digitvalue = ((result / power) - Int(result / power)) * power
        Base = Mid$(digits, digitvalue + 1, 1) & Base
        result = Int(result / power)
    Wend

End Function

Private Function ErrorHandler(sSource As String) As String
    If ErrorSource = "" Then
        ErrorSource = sSource
    Else
        ErrorHandler = Err.Description
        Exit Function
    End If
    
    Select Case sSource
        Case "DIVIDE"
            Select Case Err.Number
                Case 11
                    ErrorHandler = "Division by Zero"
            End Select
        Case "SQR"
            Select Case Err.Number
                Case 5
                    ErrorHandler = "Negative Square Root"
            End Select
        Case "NUMBER"
            Select Case Err.Number
                Case 6
                    ErrorHandler = "Number too Large"
            End Select
        Case "VARIABLE"
            Select Case Err.Number
                Case -1
                    ErrorHandler = "Wrong Number of Arguments"
                Case -2
                    ErrorHandler = "Variable not Assigned"
                Case -3
                    ErrorHandler = "Variable not Assigned"
            End Select
        Case "FACTORIAL"
            Select Case Err.Number
                Case -1
                    ErrorHandler = "Factorial not Positive Integer"
                Case Else
                    ErrorHandler = "Factorial too Large"
            End Select
        Case "FUNCTION"
            ErrorHandler = "Wrong Number of Arguments"
        Case Else
            ErrorHandler = Err.Description
    End Select
End Function

Private Function WriteDefinitions()
    Dim oFSO As New FileSystemObject
    Dim oTS As TextStream
    Dim oFunction As Variant
    Dim iIndex As Long
    Dim iIndex2 As Long
    Dim sName As String
    Dim sParameters As String
    Dim x As Dictionary
    
    Set oTS = oFSO.CreateTextFile(App.Path & "\calculator.ini", True)
    
    For iIndex = 0 To UBound(oVariables.Items)
        If Not IsEmpty(oVariables.Items(iIndex)) Then
            Set oFunction = oVariables.Items(iIndex)
            sName = oVariables.Keys(iIndex)
            oTS.Write sName
            
            If Not oFunction.Parameters Is Nothing Then
                oTS.Write "("
                sParameters = ""
                For iIndex2 = 0 To UBound(oFunction.Parameters.Items)
                    Set x = oFunction.Parameters
                    sParameters = sParameters & "," & x.Keys(iIndex2)
                Next
                oTS.Write Mid$(sParameters, 2)
                oTS.Write ")"
            End If
            oTS.Write ":=" & oFunction.Expression
            oTS.WriteLine
        End If
    Next
End Function

Private Function ReadDefinitions()
    Dim oFSO As New FileSystemObject
    Dim oTS As TextStream
    Dim oResult As ParseTree
    Dim sCommand As String
    
    Set oTS = oFSO.OpenTextFile(App.Path & "\calculator.ini")
    
'    While Not oTS.AtEndOfStream
'        sCommand = oTS.ReadLine
'        ParserText = sCommand
'        Set oResult = New ParseTree
'        If oStatement.Parse(oResult) Then
'            If ParserTextPosition = Len(sCommand) + 1 Then
'                 EvaluateLevel0 oResult
'            End If
'        End If
'    Wend
End Function

Private Sub Form_Terminate()
    WriteDefinitions
End Sub

Private Function ListDefinitions()
    Dim oFunction As Variant
    Dim iIndex As Long
    Dim iIndex2 As Long
    Dim sName As String
    Dim sParameters As String
    Dim x As Dictionary
    
    InsertText vbCrLf
    For iIndex = 0 To UBound(oVariables.Items)
        If Not IsEmpty(oVariables.Items(iIndex)) Then
            Set oFunction = oVariables.Items(iIndex)
            sName = oVariables.Keys(iIndex)
            InsertText sName
            
            If Not oFunction.Parameters Is Nothing Then
                InsertText "("
                sParameters = ""
                For iIndex2 = 0 To UBound(oFunction.Parameters.Items)
                    Set x = oFunction.Parameters
                    sParameters = sParameters & "," & x.Keys(iIndex2)
                Next
                InsertText Mid$(sParameters, 2)
                InsertText ")"
            End If
            If oFunction.Expression <> "" Then
                InsertText ":=" & oFunction.Expression
            Else
                InsertText "=" & oFunction.Value
            End If
            InsertText vbCrLf
        End If
    Next
End Function

