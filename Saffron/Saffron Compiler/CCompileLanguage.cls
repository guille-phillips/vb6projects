VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CCompileLanguage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private moObjects As New Collection
Private mvObjectNames As Variant

Private Enum RuleTypes
    stAnd = 1
    stOr
    stPerm
    stEach
    stList
    stOptional
    stExternal
    stIn
    stNot
    stSkip
    stEOS
    stBOS
    stPass
    stFail
    stLiteral
End Enum

Friend Function CompileLanguage(oReturnProgram As SaffronTree) As Collection
    Dim sObjectName As String
    Dim oObject As ISaffronObject
    Dim oRule As SaffronTree
    Dim lRuleIndex As Long
    
    mvObjectNames = Array()
    
    ' Create variables
    For Each oRule In oReturnProgram.SubTree
        Create oRule(1).Text, Choose(oRule(2)(2).Index, "And", "Or", "Perm", "Each", "List", "List", "External", "In", "Not", "Skip", "EOS", "BOS", "Null", "Null", "Literal", "Literal")
    Next
    
    For Each oRule In oReturnProgram.SubTree
        sObjectName = oRule(1).Text
        CompileFunction oRule(2), moObjects(sObjectName), sObjectName
    Next
    
    Set CompileLanguage = moObjects
End Function

Private Function Create(vName As Variant, sType As String) As ISaffronObject
    Set Create = CreateObject("SaffronClasses.Saffron" & sType)
    If Not IsMissing(vName) And Not IsEmpty(vName) Then
        moObjects.Add Create, CStr(vName)
    End If
    If Not IsEmpty(vName) Then
        If Not Exists(vName) Then
            ReDim Preserve mvObjectNames(UBound(mvObjectNames) + 1)
            mvObjectNames(UBound(mvObjectNames)) = CStr(vName)
        End If
    End If
End Function

Private Function Exists(ByVal sObjectName As String) As Boolean
    Dim lIndex As Long
    
    For lIndex = 0 To UBound(mvObjectNames)
        If mvObjectNames(lIndex) = sObjectName Then
            Exists = True
        End If
    Next
End Function

Private Function CompileFunction(ByVal oFunctionIn As SaffronTree, Optional WorkingObject As ISaffronObject, Optional sObjectName As String) As ISaffronObject
    Dim lRuleType As Long
    Dim ObjectType As String
    Dim oObject As ISaffronObject
    Dim pcParseControl As Boolean
    Dim iType As Integer
    Dim oObj As SaffronTree
    Dim oObj1 As ISaffronObject
    Dim oParams As Collection

    
    Dim oFunction As SaffronTree
    Dim lObjIndex As Long
    Dim bNullControl As Boolean
    Dim lSkipNo As Long
     
    Dim bNon As Boolean
    Dim bOmit As Boolean
    
    Select Case oFunctionIn(1).Index
        Case 2
            Select Case LCase$(oFunctionIn(1)(1).Text)
                Case "non"
                    bNon = True
                Case "omit"
                    bOmit = True
            End Select
            Select Case LCase$(oFunctionIn(1)(2).Text)
                Case "non"
                    bNon = True
                Case "omit"
                    bOmit = True
            End Select
        Case 1
            Select Case LCase$(oFunctionIn(1)(1).Text)
                Case "non"
                    bNon = True
                Case "omit"
                    bOmit = True
            End Select
    End Select

    Set oFunction = oFunctionIn(2)

    lRuleType = CInt(oFunction.Index)
        
    Select Case lRuleType
        Case stAnd, stOr, stPerm, stEach, stList, stOptional, stExternal, stIn, stNot, stSkip, stEOS, stBOS, stPass, stFail, stLiteral
            If WorkingObject Is Nothing Then
                Set CompileFunction = Create(Empty, Choose(lRuleType, "And", "Or", "Perm", "Each", "List", "List", "External", "In", "Not", "Skip", "EOS", "BOS", "Null", "Null", "Literal", "Literal"))
            Else
                Set CompileFunction = WorkingObject
            End If
    End Select
    
    Select Case lRuleType
        Case stAnd, stPerm, stEach, stOr
            Set oParams = New Collection
            For Each oObj In oFunction(1)(1).SubTree
                oParams.Add CompileFunction(oObj(1))
            Next

            CompileFunction.Initialise bNon, bOmit, oParams
    
        Case stLiteral
            Dim bSensitive As Boolean
            Dim sLiteral As String
            
            bSensitive = oFunction(1)(1).Index = 1
            
            On Error Resume Next
            Set CompileFunction = moObjects(oFunction.Text)
            If Err.Number = 5 Then
                sLiteral = EscapeLiteral(oFunction(1)(2))
                CompileFunction.Initialise bNon, bOmit, sLiteral, bSensitive
            Else
                If Not bNon And Not bOmit Then
                    Set CompileFunction = moObjects(oFunction.Text)
                ElseIf Not bNon And bOmit Then
                    Set CompileFunction = moObjects(oFunction.Text).Copy(, bOmit)
                ElseIf bNon And Not bOmit Then
                    Set CompileFunction = moObjects(oFunction.Text).Copy(bNon)
                ElseIf bNon And bOmit Then
                    Set CompileFunction = moObjects(oFunction.Text).Copy(bNon, bOmit)
                End If
            End If
            
            On Error GoTo 0
               
        Case stNot
            CompileFunction.Initialise bNon, bOmit, CompileFunction(oFunction(1)(1))
            
        Case stIn
            CompileFunction.Initialise bNon, bOmit, ResolveString(oFunction(1)(1))
                                  
        Case stOptional
            CompileFunction.Initialise bNon, bOmit, CompileFunction(oFunction(1)(1)), , , 0, 1
        
        Case stList
            Dim bDelimiter As Boolean
            Dim bTerminator As Boolean
            Dim bMin As Boolean
            Dim bMax As Boolean
            Dim lMin As Long
            Dim lMax As Long
            Dim oCondition As ISaffronObject
            Dim oDelimiter As ISaffronObject
            Dim oTerminator As ISaffronObject
            
            lMin = 1
            lMax = 0
            
            bDelimiter = oFunction(1)(2).Index = 1
            bTerminator = oFunction(1)(3).Index = 1
            bMin = oFunction(1)(4).Index = 1
            bMax = oFunction(1)(5).Index = 1
            
            Set oCondition = CompileFunction(oFunction(1)(1))
            If bDelimiter Then
                Set oDelimiter = CompileFunction(oFunction(1)(2)(1)(1))
            End If
            
            If bTerminator Then
                Set oTerminator = CompileFunction(oFunction(1)(3)(1)(1))
            End If
            
            If bMin Then
                lMin = oFunction(1)(4).Text
            End If
            
            If bMax Then
                lMax = oFunction(1)(5).Text
            End If
            
            CompileFunction.Initialise bNon, bOmit, oCondition, oDelimiter, oTerminator, lMin, lMax
            
        Case stSkip
            Dim lSkipCount As Long
            Dim bSkip As Long
            
            If oFunction(1)(1).Index = 1 Then
                lSkipCount = oFunction.Text
            Else
                lSkipCount = 1
            End If
            CompileFunction.Initialise bNon, bOmit, lSkipCount
        
        Case stEOS, stBOS
            CompileFunction.Initialise bNon, bOmit
            
        Case stPass
            CompileFunction.Initialise bNon, bOmit, True
            
        Case stFail
            CompileFunction.Initialise bNon, bOmit, False
            
        Case stExternal
            Dim vParams As Variant
            Dim vString As Variant
            
            vParams = Array()
            For Each vString In oFunction(1)(2).SubTree
                ReDim Preserve vParams(UBound(vParams) + 1)
                vParams(UBound(vParams)) = CStr(vString.Text)
            Next
            CompileFunction.Initialise bNon, bOmit, oFunction(1)(1).Text, vParams

    End Select
End Function


Private Function ResolveString(ByVal oFunction As SaffronTree) As String
    Dim oSubFunction As SaffronTree
    Dim iSubCount As Integer
    Dim bCharSet(255) As Boolean
    Dim bNot As Boolean
    Dim bSensitive As Boolean
    Dim sCharList As String
    Dim iPos As Integer
    Dim iStartRange As Integer
    Dim iEndRange As Integer
    Dim iChar As Integer
    Dim iUpperCase As Integer
    Dim iLowerCase As Integer
    
    For Each oSubFunction In oFunction.SubTree
        bNot = oSubFunction(1)(1).Index = 1
        bSensitive = oSubFunction(1)(2).Index = 1
        
        Select Case oSubFunction(1)(3).Index
            Case 1 ' Range
                If oSubFunction(1)(3)(1)(1).Index = 1 Then
                    iStartRange = oSubFunction(1)(3)(1)(1).Text
                Else
                    iStartRange = Asc(oSubFunction(1)(3)(1)(1).Text)
                End If
                If oSubFunction(1)(3)(1)(2).Index = 1 Then
                    iEndRange = oSubFunction(1)(3)(1)(2).Text
                Else
                    iEndRange = Asc(oSubFunction(1)(3)(1)(2).Text)
                End If
                
                For iPos = iStartRange To iEndRange
                    If Not bSensitive Then
                        iLowerCase = Asc(LCase(Chr(iPos)))
                        iUpperCase = Asc(UCase(Chr(iPos)))
                        bCharSet(iLowerCase) = Not bNot
                        bCharSet(iUpperCase) = Not bNot
                    Else
                        bCharSet(iPos) = Not bNot
                    End If
                Next
                
            Case 2 ' Literal
                sCharList = EscapeLiteral(oSubFunction(1)(3)(1))
                For iPos = 1 To Len(sCharList)
                    If Not bSensitive Then
                        iLowerCase = Asc(LCase$(Mid$(sCharList, iPos, 1)))
                        iUpperCase = Asc(UCase$(Mid$(sCharList, iPos, 1)))
                        bCharSet(iLowerCase) = Not bNot
                        bCharSet(iUpperCase) = Not bNot
                    Else
                        bCharSet(Asc(Mid$(sCharList, iPos, 1))) = Not bNot
                    End If
                Next
        End Select
    Next

ExitPoint:
    For iPos = 0 To 255
        If bCharSet(iPos) Then
            ResolveString = ResolveString & Chr(iPos)
        End If
    Next
End Function

Private Function EscapeLiteral(oLiteral As SaffronTree) As String
    Dim oObj As SaffronTree
    Dim sLiteral As String
    
    For Each oObj In oLiteral.SubTree
        Select Case oObj.Index
            Case 1 ' Code
                sLiteral = sLiteral & Chr$(oObj.Text)
            Case 2 ' Literal
                If oObj.Text = "||" Then
                    sLiteral = sLiteral & "|"
                Else
                    sLiteral = sLiteral & Replace$(Mid$(oObj.Text, 2, Len(oObj.Text) - 2), "||", "|")
                End If
            Case 3 ' Other
                sLiteral = sLiteral & oObj.Text
        End Select
    Next
    
    EscapeLiteral = sLiteral
End Function
