Attribute VB_Name = "Definition"
Option Explicit

Public Sub Main()
    Dim oFSO As New FileSystemObject
    Dim oParser As ISaffronObject
    Dim oSyntaxTree As New clsSyntaxTree
    Dim oIntermediates As New clsIntermediates
    Dim oObjectIntermediates As clsIntermediates
    Dim sAssembly As String
    Dim oAssembly As clsAssemblyOps
    Dim oVariables As clsAssemblyOps
    Dim oSeparator As New clsIntermediate
    
    Dim oVirtualIntermediates As clsIntermediates
    Dim oStorageIntermediates As clsIntermediates
    Dim oMainProgramIntermediates As clsIntermediates
    Dim oFunctionIntermediates As clsIntermediates
    Dim oFixedFunctionIntermediates As clsIntermediates
    
    Set oParser = InitialiseParser

    Set oMainProgramIntermediates = oSyntaxTree.LoadProgram(oParser)
    Set oVirtualIntermediates = oSyntaxTree.Scope.CompileVirtualStorage
    Set oStorageIntermediates = oSyntaxTree.Scope.CompileStorage
    Set oFunctionIntermediates = oSyntaxTree.Scope.CompileFunctions(False)
    Set oFixedFunctionIntermediates = oSyntaxTree.Scope.CompileFunctions(True)
    
    oIntermediates.MergeIntermediates oVirtualIntermediates
    oIntermediates.MergeIntermediates oStorageIntermediates
    oIntermediates.Add oSeparator.Create(opSeparator)
    oIntermediates.MergeIntermediates oMainProgramIntermediates
    oIntermediates.MergeIntermediates oFunctionIntermediates
    oIntermediates.MergeIntermediates oFixedFunctionIntermediates
    
    Set oAssembly = oIntermediates.Compile()
    sAssembly = oAssembly.Compile
    
    oFSO.CreateTextFile(App.Path & "\assembly.txt").Write sAssembly
End Sub

Public Function InitialiseParser() As ISaffronObject
    Dim Definition As String
    Dim oFSO As New FileSystemObject

    Definition = oFSO.OpenTextFile(App.Path & "\arrow.saf").ReadAll

    If Not SaffronObject.CreateRules(Definition) Then
        Debug.Print ErrorString
        End
    End If

    Set InitialiseParser = SaffronObject.Rules("statements")

'    Set InitialiseParser = SaffronObject.Rules("statements")
'
'    SaffronStream.Text = "class int  [ 0 .. 25 ]" & vbCrLf & "class uint [0..255]"
'
'    Dim oTree As New SaffronTree
'
'    If InitialiseParser.Parse(oTree) Then
'        Stop
'    Else
'        Stop
'    End If
End Function
