Attribute VB_Name = "modsock"
Public Sub ClearDataStationSock(strCommand As String)
    If mdifrm.Winsock1.State = sckConnected Then
        modgl.strSockRecive = ""
        mdifrm.Winsock1.SendData Operations.ClearDataStation & seperator & strCommand & seperator & EOS
        While (strSockRecive = "")
            DoEvents
        Wend
    End If
End Sub

Public Sub GetCodeRegisterSock(strCommand As String)
    If mdifrm.Winsock1.State = sckConnected Then
        modgl.strSockRecive = ""
        mdifrm.Winsock1.SendData Operations.GetCodeRegister & seperator & strCommand & seperator & EOS
        While (strSockRecive = "")
            DoEvents
        Wend
    End If
End Sub
Public Sub GetDataStationSock(strCommand As String)
    If mdifrm.Winsock1.State = sckConnected Then
        modgl.strSockRecive = ""
        mdifrm.Winsock1.SendData Operations.GetDataStation & seperator & strCommand & seperator & EOS
        While (strSockRecive = "")
            DoEvents
        Wend
    End If
End Sub
Public Sub SetDataExpireDateSock(strCommand As String)
    If mdifrm.Winsock1.State = sckConnected Then
        modgl.strSockRecive = ""
        mdifrm.Winsock1.SendData Operations.SetDataExpireDate & seperator & strCommand & seperator & EOS
        While (strSockRecive = "")
            DoEvents
        Wend
    End If
End Sub
Public Sub ValidateSetDataStationSock(strCommand As String)
    If mdifrm.Winsock1.State = sckConnected Then
        modgl.strSockRecive = ""
        mdifrm.Winsock1.SendData Operations.ValidateSetDataStation & seperator & strCommand & seperator & EOS
        While (strSockRecive = "")
            DoEvents
        Wend
    End If
End Sub

Public Sub SetDefaultServerDataRegisterSock(strCommand As String)
    If mdifrm.Winsock1.State = sckConnected Then
        modgl.strSockRecive = ""
        mdifrm.Winsock1.SendData Operations.SetDefaultServerDataRegister & seperator & strCommand & seperator & EOS
        While (strSockRecive = "")
            DoEvents
        Wend
    End If
End Sub

Public Sub SetDataStationSock(strCommand As String)
    If mdifrm.Winsock1.State = sckConnected Then
        modgl.strSockRecive = ""
        mdifrm.Winsock1.SendData Operations.SetDataStation & seperator & strCommand & seperator & EOS
        While (strSockRecive = "")
            DoEvents
        Wend
    End If
End Sub

Public Sub ConnectSock()
    'On Error Resume Next
    mdifrm.Winsock1.Close
    mdifrm.Winsock1.RemoteHost = clsArya.ServerName
    mdifrm.Winsock1.RemotePort = RemotePortSock
    mdifrm.Winsock1.LocalPort = LocalPortSock
    
    If mdifrm.Winsock1.State <> sckConnected Then
            bolSockIsConnected = False
            mdifrm.Winsock1.Connect
            While (bolSockIsConnected = False)
                DoEvents
            Wend
    End If
    
    If err.Number > 0 Then
        MsgBox err.Description
    End If
End Sub
Public Sub ConnectToClient()
    With mdifrm.WinsockUdp
        .Close
        .Protocol = sckUDPProtocol
        .RemoteHost = "255.255.255.255"
        .LocalPort = clsStation.DiscoveryPort
        .RemotePort = clsStation.ResponsePort
        .Bind clsStation.DiscoveryPort
     '   .Listen
    End With
'    With mdifrm.WinsockUdp
'        .Close
'        .RemoteHost = "255.255.255.255" 'clsArya.ServerName
'        .RemotePort = clsStation.ResponsePort 'RemotePortSock
'        .LocalPort = clsStation.DiscoveryPort 'LocalPortSock
'
'        If .State <> sckConnected Then
'                bolSockIsConnected = False
'                .Connect
'                While (bolSockIsConnected = False)
'                    DoEvents
'                Wend
'        End If
'
'    End With
'    If err.Number > 0 Then
'        MsgBox err.Description
'    End If

End Sub
Public Sub ConnectToClient2()
    With mdifrm.Winsock_Print
        .Close
        .Protocol = sckUDPProtocol
        .RemoteHost = "255.255.255.255"
        .LocalPort = 4000 'clsStation.DiscoveryPort
        .RemotePort = 5000 'clsStation.ResponsePort
        .Bind 4000 'clsStation.DiscoveryPort
     '   .Listen
    End With
'    With mdifrm.WinsockUdp
'        .Close
'        .RemoteHost = "255.255.255.255" 'clsArya.ServerName
'        .RemotePort = clsStation.ResponsePort 'RemotePortSock
'        .LocalPort = clsStation.DiscoveryPort 'LocalPortSock
'
'        If .State <> sckConnected Then
'                bolSockIsConnected = False
'                .Connect
'                While (bolSockIsConnected = False)
'                    DoEvents
'                Wend
'        End If
'
'    End With
'    If err.Number > 0 Then
'        MsgBox err.Description
'    End If

End Sub
Public Sub ConnectToClient_Farabin()
    
    With mdifrm.Winsock_Farabin
        .Close
        .RemoteHost = .LocalIP
        .RemotePort = 5200
        .Connect
    End With
'    wsock.Close ''Close connection
'    wsock.RemoteHost = IP 'Get IP address of PC
'    wsock.RemotePort = Port 'Port number - in Port.Text (TextBox)
'    wsock.Connect ''Set Connection

End Sub

Public Sub DiscounectSock()
    If mdifrm.Winsock1.State = sckConnected Then
        modgl.strSockRecive = ""
        mdifrm.Winsock1.SendData Operations.LogOutStation & seperator & EOS
        While (strSockRecive = "")
            DoEvents
        Wend
    End If
End Sub
Public Sub PrintReport(strCommand As String)
    If mdifrm.Winsock1.State = sckConnected Then
        modgl.strSockRecive = ""
        mdifrm.Winsock1.SendData Operations.PrintReport & seperator & strCommand & seperator & EOS
'        While (strSockRecive = "")
'            DoEvents
'        Wend
    End If
End Sub

