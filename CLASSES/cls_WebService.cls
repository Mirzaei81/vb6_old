VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls_WebService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Private mvarSOAP As MSSOAPLib30.SoapClient30
Private mvarSOAP As MSSOAPLib.SoapClient
Private mvarURL As String
Private DialUpConnectionName  As String
Private DialUpConnectionUserName As String
Private DialUpConnectionPassword As String
Private DialUpConnectionTell As String
Private DialUpConnectionLog As String

Public Function Connect(p_URL As String) As Long
    On Error GoTo ErrorHandler
    Dim RetValue As Long
    RetValue = ConnectDialUp
    If RetValue <> 0 Then
        Connect = RetValue
        Exit Function
    End If
    'Set mvarSOAP = New MSSOAPLib30.SoapClient30
    Set mvarSOAP = New MSSOAPLib.SoapClient
    mvarSOAP.mssoapinit p_URL

    If mvarSOAP.faultstring = "" Then
        Connect = RetValue
        mvarURL = p_URL
    Else
        Connect = -1
        mvarURL = ""
    End If
    Exit Function
ErrorHandler:
    MsgBox err.Description
    Connect = -1
    DisconnectDialUp

End Function

Private Sub Class_Initialize()
    DialUpConnectionUserName = "mkt2000"
    DialUpConnectionPassword = "2000mkt"
    If clsStation.CityCode = "21" Then
        DialUpConnectionTell = "88506928"
    Else
        DialUpConnectionTell = "02188506928"
    End If
End Sub

Private Sub Class_Terminate()
    Set mvarSOAP = Nothing
End Sub

Public Function GetCodeRegisterClientWS(Code As String) As String
    GetCodeRegisterClientWS = mvarSOAP.GetCodeRegisterClientWS(Code)
    DisconnectDialUp
End Function
Private Function ConnectDialUp() As Long
    ConnectDialUp = DialUpConnection.Dial(DialUpConnectionTell, DialUpConnectionUserName, DialUpConnectionPassword)
End Function

Private Function DisconnectDialUp() As Boolean
    DisconnectDialUp = DialUpConnection.HangUp
End Function

