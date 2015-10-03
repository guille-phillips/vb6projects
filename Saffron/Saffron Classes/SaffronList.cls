VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SaffronList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements ISaffronObject

' Generic variables

Private mbOmit As Boolean
Private mbNonConsuming As Boolean

' Object specific variables
Private moCondition As ISaffronObject
Private moDelimiter As ISaffronObject
Private moTerminator As ISaffronObject
Private mlLowerLimit As Long
Private mlUpperLimit As Long

Private Function ISaffronObject_Copy(Optional ByVal bNonConsuming As Variant, Optional ByVal bOmit As Variant) As ISaffronObject
    If IsMissing(bNonConsuming) Then
        bNonConsuming = mbNonConsuming
    End If
    If IsMissing(bOmit) Then
        bOmit = mbOmit
    End If
    
    Set ISaffronObject_Copy = New SaffronList
    
    If moDelimiter Is Nothing And moTerminator Is Nothing Then
        ISaffronObject_Copy.Initialise CBool(bNonConsuming), CBool(bOmit), moCondition, , , mlLowerLimit, mlUpperLimit
    ElseIf moDelimiter Is Nothing And Not moTerminator Is Nothing Then
        ISaffronObject_Copy.Initialise CBool(bNonConsuming), CBool(bOmit), moCondition, , moTerminator.Copy, mlLowerLimit, mlUpperLimit
    ElseIf Not moDelimiter Is Nothing And moTerminator Is Nothing Then
        ISaffronObject_Copy.Initialise CBool(bNonConsuming), CBool(bOmit), moCondition, moDelimiter.Copy, , mlLowerLimit, mlUpperLimit
    Else
        ISaffronObject_Copy.Initialise CBool(bNonConsuming), CBool(bOmit), moCondition, moDelimiter.Copy, moTerminator.Copy, mlLowerLimit, mlUpperLimit
    End If
    
End Function

Private Property Get ISaffronObject_Omitted() As Boolean
    ISaffronObject_Omitted = mbOmit
End Property


' Initialises parameters used for parsing
Public Function ISaffronObject_Initialise(bNonConsuming As Boolean, bOmit As Boolean, ParamArray pInitParam() As Variant) As ISaffronObject
    mbNonConsuming = bNonConsuming
    mbOmit = bOmit
    
    mlLowerLimit = 1
    mlUpperLimit = 0
    Set moCondition = Nothing
    Set moDelimiter = Nothing
    Set moTerminator = Nothing
    
    Set moCondition = pInitParam(0)
    
    If UBound(pInitParam) >= 1 Then
        If Not IsMissing(pInitParam(1)) Then
            Set moDelimiter = pInitParam(1)
        End If
    End If
    
    If UBound(pInitParam) >= 2 Then
        If Not IsMissing(pInitParam(2)) Then
            Set moTerminator = pInitParam(2)
        End If
    End If
    
    If UBound(pInitParam) >= 3 Then
        If Not IsMissing(pInitParam(3)) Then
            mlLowerLimit = pInitParam(3)
            If mlLowerLimit < 0 Then
                Err.Raise 1, "SaffronList: ISaffronObject_Initialise", "Lower Bound cannot be less than zero"
            End If
        End If
    End If
    If UBound(pInitParam) >= 4 Then
        If Not IsMissing(pInitParam(4)) Then
            mlUpperLimit = pInitParam(4)
            If mlUpperLimit < 0 Then
                Err.Raise 2, "SaffronList: ISaffronObject_Initialise", "Upper Bound cannot be less than zero"
            ElseIf mlUpperLimit < mlLowerLimit And mlUpperLimit <> 0 Then
                Err.Raise 3, "SaffronList: ISaffronObject_Initialise", "Upper Bound cannot be less than Lower Bound"
            End If
        End If
    End If
    Set ISaffronObject_Initialise = Me
End Function

' Will perform the parsing function on the object - if parsing fails will return FALSE.
Public Function ISaffronObject_Parse(oResult As SaffronTree) As Boolean
    Dim lCounter As Long
    Dim oObjectResult As SaffronTree
    Dim oDelimiterResult As SaffronTree
    Dim lPosition As Long
    Dim lDelimiterPosition As Long
    Dim bDelimiterParsed As Boolean
    
    ISaffronObject_Parse = True
    
    lPosition = TextPosition
    oResult.TextStart = lPosition
    oResult.TextEnd = oResult.TextStart - 1
    
    If Not moTerminator Is Nothing Then
        Set oObjectResult = New SaffronTree
        If moTerminator.Parse(oObjectResult) Then
            If mlLowerLimit = 0 Then
                If Not moTerminator.Omitted Then
                    oResult.Index = 1
                    oResult.TextEnd = oObjectResult.TextEnd
                    oResult.SubTree.Add oObjectResult
                End If
                If mbNonConsuming Then TextPosition = lPosition
                Exit Function
            Else
                ISaffronObject_Parse = False
                oResult.Error = True
                If mbNonConsuming Then TextPosition = lPosition
                Exit Function
            End If
        End If
    End If

    Set oObjectResult = New SaffronTree
    If moCondition.Parse(oObjectResult) Then
        lCounter = 1
        oResult.Index = 1
        If Not moCondition.Omitted Then
            oResult.TextEnd = oObjectResult.TextEnd
            oResult.SubTree.Add oObjectResult
        End If
        lCounter = lCounter + 1

        If moTerminator Is Nothing Then
            If mlUpperLimit = 1 Then
                If mbNonConsuming Then TextPosition = lPosition
                Exit Function
            End If
        End If
    Else
        If mlLowerLimit = 0 And moTerminator Is Nothing Then
            If mbNonConsuming Then TextPosition = lPosition
            Exit Function
        Else
            ISaffronObject_Parse = False
            oResult.Error = True
            If mbNonConsuming Then TextPosition = lPosition
            Exit Function
        End If
    End If
    
    Do
        If Not moTerminator Is Nothing Then
            If Not oDelimiterResult Is Nothing Then
                ISaffronObject_Parse = False
                oResult.Error = True
                If mbNonConsuming Then TextPosition = lPosition
                Exit Function
            End If
            Set oObjectResult = New SaffronTree
            If moTerminator.Parse(oObjectResult) Then
                If (lCounter - 1) >= mlLowerLimit And ((lCounter - 1) <= mlUpperLimit Or mlUpperLimit = 0) Then
                    If Not moTerminator.Omitted Then
                        oResult.Index = lCounter
                        oResult.TextEnd = oObjectResult.TextEnd
                        oResult.SubTree.Add oObjectResult
                    End If
                    If mbNonConsuming Then TextPosition = lPosition
                    Exit Function
                Else
                    ISaffronObject_Parse = False
                    oResult.Error = True
                    If mbNonConsuming Then TextPosition = lPosition
                    Exit Function
                End If
            Else
                If (lCounter - 1) = mlUpperLimit Then
                    ISaffronObject_Parse = False
                    oResult.Error = True
                    If mbNonConsuming Then TextPosition = lPosition
                    Exit Function
                End If
            End If
        Else
            If (lCounter - 1) = mlUpperLimit Then
                If mbNonConsuming Then TextPosition = lPosition
                Exit Function
            End If
        End If
        
        If Not moDelimiter Is Nothing Then
            Set oDelimiterResult = New SaffronTree
            lDelimiterPosition = TextPosition
            If Not moDelimiter.Parse(oDelimiterResult) Then
                bDelimiterParsed = False
                If Not moTerminator Is Nothing Then
                    ISaffronObject_Parse = False
                    oResult.Error = True
                    If mbNonConsuming Then TextPosition = lPosition
                    Exit Function
                Else
                    If (lCounter - 1) >= mlLowerLimit And ((lCounter - 1) <= mlUpperLimit Or mlUpperLimit = 0) Then
                        If mbNonConsuming Then TextPosition = lPosition
                        Exit Function
                    Else
                        ISaffronObject_Parse = False
                        oResult.Error = True
                        If mbNonConsuming Then TextPosition = lPosition
                        Exit Function
                    End If
                End If
            Else
                bDelimiterParsed = True
            End If
        End If
        
        Set oObjectResult = New SaffronTree
        If moCondition.Parse(oObjectResult) Then
            If Not oDelimiterResult Is Nothing Then
                If Not moDelimiter.Omitted Then
                    oResult.TextEnd = oDelimiterResult.TextEnd
                    oResult.SubTree.Add oDelimiterResult
                End If
            End If
            Set oDelimiterResult = Nothing
            oResult.Index = lCounter
            If Not moCondition.Omitted Then
                oResult.TextEnd = oObjectResult.TextEnd
                oResult.SubTree.Add oObjectResult
            End If
            lCounter = lCounter + 1
        Else
            If moTerminator Is Nothing Then
                If (lCounter - 1) >= mlLowerLimit And ((lCounter - 1) <= mlUpperLimit Or mlUpperLimit = 0) Then
                    If Not moDelimiter Is Nothing And bDelimiterParsed Then
                        TextPosition = lDelimiterPosition
                    End If
                    If mbNonConsuming Then TextPosition = lPosition
                    Exit Function
                Else
                    ISaffronObject_Parse = False
                    oResult.Error = True
                    If mbNonConsuming Then TextPosition = lPosition
                    Exit Function
                End If
            Else
                ISaffronObject_Parse = False
                oResult.Error = True
                If mbNonConsuming Then TextPosition = lPosition
                Exit Function
            End If
        End If
    Loop
End Function


