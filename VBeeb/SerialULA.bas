Attribute VB_Name = "SerialULA"
Option Explicit

Public TransmitBaud As Long
Public ReceiveBaud As Long
Public SelectRS432 As Long
Public CassetteMotor As Long

Public TransmitBaudRate As Long
Public ReceiveBaudRate As Long

Public Sub InitialiseSerialULA()
    ' Debugging.WriteString "SerialULA.InitialiseSerialULA"
    
    TransmitBaudRate = 19200
    ReceiveBaudRate = 19200
    SelectRS432 = 1
End Sub

Public Sub WriteRegister(yValue As Byte)
    ' Debugging.WriteString "SerialULA.WriteRegister"
    
    TransmitBaud = yValue And &H7&
    ReceiveBaud = (yValue And &H38&) \ 8&
    SelectRS432 = (yValue And &H40&) \ 64&

    CassetteMotor = (yValue And &H80&) \ 128&

    TransmitBaudRate = Array(19200, 1200, 4800, 150, 9600, 300, 2400, 75)(TransmitBaud)
    ReceiveBaudRate = Array(19200, 1200, 4800, 150, 9600, 300, 2400, 75)(ReceiveBaud)
    
    ACIA6850.SetCyclesPerByte
    
    KeyboardIndicators.SetLED 4, CassetteMotor
End Sub

