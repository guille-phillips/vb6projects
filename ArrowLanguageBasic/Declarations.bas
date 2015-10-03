Attribute VB_Name = "Declarations"
Option Explicit

Public ObjectFactory As New clsObject
Public IntermediateFactory As New clsIntermediate
Public ClassFactory As New clsClass
Public AssemblyOpFactory As New clsAssemblyOp
Public RangeFactory As New clsRange
Public ScopeFactory As New clsScope

Public Const RANGE_SET As Long = 1
Public Const IDENTIFIER_SET As Long = 2
Public Const SYMBOL_SET As Long = 3
Public Const REC As Long = 4
Public Const STR As Long = 5

Public Const FOR_LOOP_STATEMENT = 0

Public Enum ILOperators
    opNone
    opCopy
    opCopyToRegister
    opAdd
    opSub
    opMultiply
    opDivide
    opModulus
    opEqual
    opNotEqual
    opLessThan
    opGreaterThan
    opLessThanEqual
    opGreaterThanEqual
    opAnd
    opOr
    opEor
    opIf
    opEndIf
    opStartWhile
    opWhile
    opEndWhile
    opStartUntil
    opUntil
    opEndUntil
    opStartDoWhile
    opEndDoWhile
    opStartDoUntil
    opEndDoUntil
    opIsZero
    opFunction
    opEndFunction
    opReturn
    opFunctionCall
    opOrigin
    opObject
    opSeparator
End Enum

Public Enum AssemblyOps
    ioNOP
    ioLD
    ioST
    ioADC
    ioSBC
    ioCMP
    ioAND
    ioOR
    ioEOR
    ioPH
    ioPL
    ioBRS
    ioBRC
    ioDEC
    ioINC
    ioROL
    ioROR
    ioASL
    ioLSR
    ioAddress
    ioLabel
    ioAddressLabel
    ioConstant
    ioLDAFlag
    ioLDAInvFlag
    ioJSR
    ioRTS
    ioJMP
    ioSeparator
End Enum

Public Enum AssemblyOpRegister
    irNone
    irA
    irX
    irY
    irS
    irP
    irC
    irZ
    irN
    irV
    irD
    irI
    irAX
    irAY
    irXY
End Enum

Public Enum AssemblyOpMode
    imImplied
    imAddress
    imConstant
    imIndexed
    imIndirectPre
    imIndirectPost
    imIndirect
    imRegister
End Enum

Public Enum ObjectTypes
    otVar
    otVarIndexed
    otParam
    otConst
    otVirt
    otFunc
    otAny
End Enum


Public Enum IntermediateOptions
    osFlagOnly = 1
    osFlagSet = 2
    osUseZFlag = 4
    osUseCFlag = 8
    osUseVFlag = 16
    osUseNFlag = 32
End Enum

Public Enum RegisterTargets
    rtA
    rtX
    rtY
    rtXY
    rtAX
    rtAY
End Enum

Public Enum AccessTypes
    atByte
    atBit
    atCompact
    atDefault
End Enum

Public Function Largest(ByVal lValue1 As Long, ByVal lValue2 As Long, Optional ByVal lValue3 As Long) As Long
    If lValue1 > lValue2 Then
        Largest = lValue1
    Else
        Largest = lValue2
    End If
    
    If lValue3 > 0 Then
        If lValue3 > Largest Then
            Largest = lValue3
        End If
    End If
End Function


