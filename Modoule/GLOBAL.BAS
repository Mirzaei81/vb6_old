Attribute VB_Name = "GLOBAL"
' SMARTKEY-NET fields'' length

Global Const MAX_LEN_LABEL = 16
Global Const MAX_LEN_PASSWORD = 16
Global Const MAX_LEN_DATA = 64
Global Const MAX_LEN_EXTENDED = 352
Global Const MAX_LEN_SCRAMBLE = 8


'Memory allocation for all the SMARTKEY-NET 's fields
Global sKey(231) As Integer

'Offset about each files, expressed with integer (2 bytes)

Global Const INI_NET_COM = 1    ' Net command field
Global Const END_NET_COM = 1
Global Const INI_NET_PSW = 2    ' Net password filed
Global Const END_NET_PSW = 3
Global Const INI_LPT = 4        ' Lpt field
Global Const END_LPT = 4
Global Const INI_COMMAND = 5    ' Command filed
Global Const END_COMMAND = 5
Global Const INI_LABEL = 6
Global Const END_LABEL = 13
Global Const INI_PASSWORD = 14
Global Const END_PASSWORD = 21
Global Const INI_DATA = 22
Global Const END_DATA = 53
Global Const SKEY_FAIL = 54
Global Const SKEY_STATUS = 55
Global Const INI_EXTENDED = 56
Global Const END_EXTENDED = 231



'Network STATUS error

Global Const ST_OK = 0
Global Const ST_KEY_NONE = -1
Global Const ST_SYNT_ERR = -2
Global Const ST_LABEL_FAILED = -3
Global Const ST_PW_DATA_FAILED = -4
Global Const ST_NET_ERROR = -5
Global Const ST_CLOSE_ERROR = -6
Global Const ST_ACCESS_ERROR = -7
Global Const ST_USER_ERROR = -8
Global Const ST_NET_PWD_ERROR = -9
Global Const ST_TSR_NOT_INST = -10
Global Const ST_INIT_ERROR = -11
Global Const ST_HW_FAILUR = -20



'STATUS error
'MODLES
Global Const SKEY_NONE = 0
Global Const SKEY_FX = 1
Global Const SKEY_PR = 2
Global Const SKEY_EP = 3
Global Const SKEY_NET_5 = 4
Global Const SKEY_NET_10 = 5
Global Const SKEY_NET_25 = 6
Global Const SKEY_NET_50 = 7
Global Const SKEY_NET_NL = 8
Global Const SKEY_SP = 9
Global Const SKEY_NET = &HA


' String values of Smartkey's memory fields
Global KeyLabel As String
Global KeyPassword As String
Global KeyData As String
Global KeyPas As String


'DLL library declaration
Declare Function smartlink Lib "SKEYLINK.dll" (sKey As Integer) As Integer

''''''''''
Global Const EOS = "@@END@@"
Global Const seperator = ";"
Global Const RemotePortSock = 1144
Global Const LocalPortSock = 0

Global Const RemotePort_Pos_Eghtesade = 8002 'Eghtesade Novin
Global Const RemotePort_Pos_Saderat = 8012  'saderat

