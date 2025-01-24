Attribute VB_Name = "Module1"
Public Declare Function MF_GetDLL_Ver Lib "MF_API.dll" (ByRef rVER As Byte) As Integer

Public Declare Function MF_InitComm Lib "MF_API.dll" (ByVal PortName As String, ByVal baud As Long) As Long

Public Declare Function MF_ControlBuzzer Lib "MF_API.dll" (ByVal DeviceAddr As Integer, BeepTime As Integer) As Long

Public Declare Function MF_DeviceReset Lib "MF_API.dll" (ByVal DeviceAddr As Integer) As Long
Public Declare Function MF_ExitComm Lib "MF_API.dll" () As Long
Public Declare Function MF_GetDevice_Ver Lib "MF_API.dll" (ByVal DeviceAddr As Integer, ByRef ver As Byte) As Long
Public Declare Function MF_SetDeviceBaud Lib "MF_API.dll" (ByVal DeviceAddr As Integer, ByVal baud As Long) As Long
Public Declare Function MF_SetDeviceAddr Lib "MF_API.dll" (ByVal DeviceAddr As Integer, ByVal Addr As Integer) As Long
Public Declare Function MF_ControlLED Lib "MF_API.dll" (ByVal DeviceAddr As Integer, ByVal LED1 As Integer, ByVal LED2 As Integer) As Long
Public Declare Function MF_GetDeviceAddr Lib "MF_API.dll" (ByVal DeviceAddr As Integer, ByRef Addr As Byte) As Long
Public Declare Function MF_SetDeviceSNR Lib "MF_API.dll" (ByVal DeviceAddr As Integer, ByVal snr As String) As Long
Public Declare Function MF_GetDeviceSNR Lib "MF_API.dll" (ByVal DeviceAddr As Integer, ByRef snr As Byte) As Long
Public Declare Function MF_SetRF_ON Lib "MF_API.dll" (ByVal DeviceAddr As Integer) As Long
Public Declare Function MF_SetRF_OFF Lib "MF_API.dll" (ByVal DeviceAddr As Integer) As Long
Public Declare Function MF_SetWiegandMode Lib "MF_API.dll" (ByVal DeviceAddr As Integer, ByVal Mode As Integer, ByVal alarm As Integer) As Long
'''''''''''''''''''''''''''''''''''card reading functions''''''''''''''''''''''''''''''''''''''''''
Public Declare Function MF_Request Lib "MF_API.dll" (ByVal DeviceAddr As Integer, ByVal Mode As Integer, ByRef CardType As Byte) As Long
Public Declare Function MF_Anticoll Lib "MF_API.dll" (ByVal DeviceAddr As Integer, ByRef snr As Byte) As Long
Public Declare Function MF_Halt Lib "MF_API.dll" (ByVal DeviceAddr As Integer) As Long
Public Declare Function MF_Select Lib "MF_API.dll" (ByVal DeviceAddr As Integer, ByRef snr As Byte) As Long
Public Declare Function MF_LoadKey Lib "MF_API.dll" (ByVal DeviceAddr As Integer, ByRef Key As Byte) As Long

Public Declare Function MF_LoadKeyFromEE Lib "MF_API.dll" (ByVal DeviceAddr As Integer, ByVal KeyType As Integer, ByVal KeyNum As Integer) As Long
Public Declare Function MF_StoreKeyToEE Lib "MF_API.dll" (ByVal DeviceAddr As Integer, ByVal KeyAB As Integer, ByVal KeyAdd As Integer, ByRef Key As Byte) As Long
Public Declare Function MF_Authentication Lib "MF_API.dll" (ByVal DeviceAddr As Integer, ByVal AuthType As Integer, ByVal block As Integer, ByRef snr As Byte) As Long
Public Declare Function MF_Read Lib "MF_API.dll" (ByVal DeviceAddr As Integer, ByVal block As Integer, ByVal numbers As Integer, ByRef databuff As Byte) As Long
Public Declare Function MF_Write Lib "MF_API.dll" (ByVal DeviceAddr As Integer, ByVal block As Integer, ByVal numbers As Integer, ByRef databuff As Byte) As Long
Public Declare Function MF_Value Lib "MF_API.dll" (ByVal DeviceAddr As Integer, ByVal valoption As Integer, ByRef Value As Byte) As Long
Public Declare Function MF_transfer Lib "MF_API.dll" (ByVal DeviceAddr As Integer, ByVal block As Integer) As Long

Public DLL_version(32) As Byte
Public portN(3) As Byte
Public Dver(32) As Byte
Public Daddress As Byte
Public Dsn(7) As Byte
Public cardT(2) As Byte
Public cardSN(3) As Byte
Public Ckey(5) As Byte
Public databuffer(255) As Byte
Public Value(3) As Byte
Public Dbuffer(64) As Byte
Public Function hex2dec(ByVal inpt As String) As Integer
'On Error Resume Next
If Len(inpt) = 1 Then inpt = "0" & inpt
Select Case (Mid(inpt, 1, 1))
Case "A": hex2dec = hex2dec + 10 * 16
Case "a": hex2dec = hex2dec + 10 * 16
Case "B": hex2dec = hex2dec + 11 * 16
Case "b": hex2dec = hex2dec + 11 * 16
Case "C": hex2dec = hex2dec + 12 * 16
Case "c": hex2dec = hex2dec + 12 * 16
Case "D": hex2dec = hex2dec + 13 * 16
Case "d": hex2dec = hex2dec + 13 * 16
Case "E": hex2dec = hex2dec + 14 * 16
Case "e": hex2dec = hex2dec + 14 * 16
Case "F": hex2dec = hex2dec + 15 * 16
Case "f": hex2dec = hex2dec + 15 * 16
Case Else: hex2dec = hex2dec + Mid(inpt, 1, 1) * 16
End Select
Select Case (Mid(inpt, 2, 1))
Case "A": hex2dec = hex2dec + 10
Case "a": hex2dec = hex2dec + 10
Case "B": hex2dec = hex2dec + 11
Case "b": hex2dec = hex2dec + 11
Case "C": hex2dec = hex2dec + 12
Case "c": hex2dec = hex2dec + 12
Case "D": hex2dec = hex2dec + 13
Case "d": hex2dec = hex2dec + 13
Case "E": hex2dec = hex2dec + 14
Case "e": hex2dec = hex2dec + 14
Case "F": hex2dec = hex2dec + 15
Case "f": hex2dec = hex2dec + 15
Case Else: hex2dec = hex2dec + Mid(inpt, 2, 1)
End Select
End Function


