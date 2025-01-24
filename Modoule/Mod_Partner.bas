Attribute VB_Name = "Mod_Partner"

'Partner Modolue

Public Declare Function Inp32 Lib "inpout32.dll" _
(ByVal PortAddress As Integer) As Integer
Public Declare Sub Out32 Lib "inpout32.dll" _
(ByVal PortAddress As Integer, ByVal Value As Integer)



Const GPIO10 = &H1
Const GPIO10MASK = &HFE
Const GPIO11 = &H2
Const GPIO12 = &H4
Const GPIO12MASK = &HFB
Const GPIO13 = &H8
Const GPIO14 = &H10
Const GPIO15 = &H20
Const GPIO16 = &H40
Const GPIO17 = &H80
Public Const GPIOAllIn = &HFF


Public Const Cash1Out = GPIO10
Public Const Cash2Out = GPIO11
Public Const Cash1In = GPIO12
Public Const Cash2In = GPIO13
Public Const GPIOInOutSetReg = &HF0
Public Const GPIOFuncSelAddress = &H2A
Public Const GPIOFuncSelSetDefault = &H7C
Public Const GPIOInOutDataReg = &HF1
Public Const Cash1OutMASK = GPIO10MASK
Public Const Cash2OutMASK = GPIO12MASK

Public MutilSelectDefault As Integer
Public GPIOSetRegDefault As Integer
Public GPIOSetDataDefault As Integer

Public IsCashA As Boolean
Public Counter As Long
Public TotalDelay As Integer
Public IsFirst As Boolean

Public Sub Enter_Config()

Out32 &H2E, &H87
Sleep 10
Out32 &H2E, &H87
Sleep 10
End Sub
Public Sub SelectLD7()

'PowerOnDefault
SendData &H7, &H7

End Sub
Public Sub MutilpinSelGPIO()
    'SendData GPIOInOutSetReg, &HFF
    SendData &H2A, &HFC
    'SendData GPIOFuncSelAddress, &HFC
    End Sub
Public Sub MutilpinSelDefault()

    SendData &H2A, &H2E
    End Sub
Public Sub DefineInOut()
    ReadData (GPIOInOutSetReg)
    ReadData (GPIOInOutDataReg)
    SendData GPIOInOutDataReg, &HFF
    ReadData (GPIOInOutDataReg)
    SendData GPIOInOutSetReg, &HFC
    'SendData &HF0, &HFA
    'MsgBox "FC--->F0"
End Sub
Public Sub End_Config()

Out32 &H2E, &HAA
End Sub

Public Sub BackToDefault()
    Enter_Config
    SelectLD7
    MutilpinSelDefault
    'MsgBox "MutilpinSelDefault"
    PowerOnSet
    'MsgBox "GPIO Set"
    'PowerOnData
    'MsgBox "GPIO Data"
    End_Config

End Sub

Public Sub ReadMutilSel()
    MutilSelectDefault = ReadData(&H2A)
    End Sub
Public Sub ReadGPIOSet()
    GPIOSetRegDefault = ReadData(GPIOInOutSetReg)
    End Sub
Public Sub ReadGPIOData()
    GPIOSetDataDefault = ReadData(GPIOInOutDataReg)
    End Sub

Public Sub PowerOnSet()
    SendData GPIOInOutSetReg, &HFF
    SendData GPIOInOutDataReg, &HCF
End Sub
Public Sub PowerOnData()
    SendData GPIOInOutDataReg, 0 'GPIOSetDataDefault
End Sub

Public Sub SendData(ByVal Addr As Byte, ByVal SendVal As Byte)
    ''List1.AddItem "Send F0h to " & Hex(Addr) & "h"
    Out32 &H2E, Addr
    Sleep 10
    ''List1.AddItem "Send F0h to " & Hex(SendVal) & "h"
    Out32 &H2F, SendVal
    Sleep 10
    'List1.AddItem "Send " & Hex(SendVal) & " to " & Hex(Addr)
End Sub

Public Function ReadData(ByVal Addr As Byte) As Byte
    ''List1.AddItem "Send F0h to " & Hex(Addr) & "h"
    Out32 &H2E, Addr
    Sleep 10
    ''List1.AddItem "Send F0h to " & Hex(SendVal) & "h"
    ReadData = Inp32(&H2F)
    Sleep 10
    'List1.AddItem "Get Data " & Hex(ReadData) & " from " & Hex(Addr)
End Function







