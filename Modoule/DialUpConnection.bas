Attribute VB_Name = "DialUpConnection"
Option Explicit
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, ByVal pSrc As String, ByVal ByteLen As Long)
Private Declare Sub ZeroMemory Lib "kernel32.dll" Alias "RtlZeroMemory" (Destination As Any, ByVal Length As Long)
Public Declare Function RasEnumConnections Lib _
"rasapi32.dll" Alias "RasEnumConnectionsA" (lpRasConn As _
Any, lpcb As Long, lpcConnections As Long) As Long

Public Declare Function RasHangUp Lib "rasapi32.dll" Alias _
"RasHangUpA" (ByVal hRasConn As Long) As Long
Public gstrISPName As String
Public ReturnCode As Long

Const RAS95_MaxEntryName = 256
Const RAS_MaxPhoneNumber = 128
Const RAS_MaxCallbackNumber = RAS_MaxPhoneNumber

Const UNLEN = 256
Const PWLEN = 256
Const DNLEN = 12
Private Type RASDIALPARAMS
   dwSize As Long ' 1052
   szEntryName(RAS95_MaxEntryName) As Byte
   szPhoneNumber(RAS_MaxPhoneNumber) As Byte
   szCallbackNumber(RAS_MaxCallbackNumber) As Byte
   szUserName(UNLEN) As Byte
   szPassword(PWLEN) As Byte
   szDomain(DNLEN) As Byte
End Type

Private Type RASENTRYNAME95
    'set dwsize to 264
    dwSize As Long
    szEntryName(RAS95_MaxEntryName) As Byte
End Type
Public Const RAS_MAXENTRYNAME As Integer = 256
Public Const RAS_MAXDEVICETYPE As Integer = 16
Public Const RAS_MAXDEVICENAME As Integer = 128
Public Const RAS_RASCONNSIZE As Integer = 412

Public Type RasEntryName
    dwSize As Long
    szEntryName(RAS_MAXENTRYNAME) As Byte
End Type

Public Type RasConn
    dwSize As Long
    hRasConn As Long
    szEntryName(RAS_MAXENTRYNAME) As Byte
    szDeviceType(RAS_MAXDEVICETYPE) As Byte
    szDeviceName(RAS_MAXDEVICENAME) As Byte
End Type

Private Declare Function RasDial Lib "rasapi32.dll" Alias "RasDialA" (ByVal lprasdialextensions As Long, ByVal lpcstr As String, ByRef lprasdialparamsa As RASDIALPARAMS, ByVal dword As Long, lpvoid As Any, ByRef lphrasconn As Long) As Long
Private Declare Function RasEnumEntries Lib "rasapi32.dll" Alias "RasEnumEntriesA" (ByVal reserved As String, ByVal lpszPhonebook As String, lprasentryname As Any, lpcb As Long, lpcEntries As Long) As Long
Private Declare Function RasGetEntryDialParams Lib "rasapi32.dll" Alias "RasGetEntryDialParamsA" (ByVal lpcstr As String, ByRef lprasdialparamsa As RASDIALPARAMS, ByRef lpbool As Long) As Long

Public Function Dial(ByVal PhoneNumber As String, ByVal UserName As String, ByVal Password As String) As Long
    Dim rp As RASDIALPARAMS, h As Long, resp As Long
    rp.dwSize = Len(rp) + 6
    ChangeBytes "", rp.szEntryName
    ChangeBytes PhoneNumber, rp.szPhoneNumber 'Phone number stored for the connection
    ChangeBytes "", rp.szCallbackNumber 'Callback number stored for the connection
    ChangeBytes UserName, rp.szUserName
    ChangeBytes Password, rp.szPassword
    ChangeBytes "", rp.szDomain 'Domain stored for the connection
    'Dial
    resp = RasDial(ByVal 0, ByVal 0, rp, 0, ByVal 0, h)   'AddressOf RasDialFunc
    Dial = resp
End Function

Private Function ChangeToStringUni(Bytes() As Byte) As String
    'Changes an byte array to a Visual Basic unicode string
    Dim temp As String
    temp = StrConv(Bytes, vbUnicode)
    ChangeToStringUni = Left(temp, InStr(temp, Chr(0)) - 1)
End Function

Private Function ChangeBytes(ByVal str As String, Bytes() As Byte) As Boolean
    'Changes a Visual Basic unicode string to an byte array
    'Returns True if it truncates str
    Dim lenBs As Long 'length of the byte array
    Dim lenStr As Long 'length of the string
    lenBs = UBound(Bytes) - LBound(Bytes)
    lenStr = LenB(StrConv(str, vbFromUnicode))
    If lenBs > lenStr Then
        CopyMemory Bytes(0), str, lenStr
        ZeroMemory Bytes(lenStr), lenBs - lenStr
    ElseIf lenBs = lenStr Then
        CopyMemory Bytes(0), str, lenStr
    Else
        CopyMemory Bytes(0), str, lenBs 'Queda truncado
        ChangeBytes = True
    End If
End Function


Public Function HangUp()
    Dim i As Long
    Dim lpRasConn(255) As RasConn
    Dim lpcb As Long
    Dim lpcConnections As Long
    Dim hRasConn As Long
    lpRasConn(0).dwSize = RAS_RASCONNSIZE
    lpcb = RAS_MAXENTRYNAME * lpRasConn(0).dwSize
    lpcConnections = 0
    ReturnCode = RasEnumConnections(lpRasConn(0), lpcb, _
    lpcConnections)
    HangUp = False
    Dim ERROR_SUCCESS As Long
    
    If ReturnCode = ERROR_SUCCESS Then
        For i = 0 To lpcConnections - 1
            If Trim(ByteToString(lpRasConn(i).szEntryName)) = Trim(gstrISPName) Then
                hRasConn = lpRasConn(i).hRasConn
                ReturnCode = RasHangUp(ByVal hRasConn)
                HangUp = True
            End If
        Next i
    End If
End Function

Public Function ByteToString(bytString() As Byte) As String
    Dim i As Integer
    ByteToString = ""
    i = 0
    While bytString(i) = 0&
        ByteToString = ByteToString & Chr(bytString(i))
        i = i + 1
    Wend
End Function

