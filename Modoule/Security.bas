Attribute VB_Name = "Security"
Option Explicit

Public strFarabin As String
Public IsFarabin As Boolean
Public IsClientString As Boolean
Public strDataLock As String
Public IsClient As Boolean
Public KarbarKey As String
Public CustomerRegisterFlag As Boolean
Private DateRemain As Long
'Get the computer Name
Public Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public MachineLocalIp As String
Public SystemFolderName As String
Public HardLockFlag As Boolean
Public HardLockFlagTrial As Boolean
Public AutoHavale As Long
Public RegRec As Long   'Amount Records in Registry
Public Hhhh As String   'HardDiskId
Public SanadCountingRecord As Long  'For Limited Version
Public TrialCountFlag As Long
Public RemaindateFlag As Boolean
Public maxRecordCountFlag As Boolean
Private clsDate As New clsDate
Private rctmp As New ADODB.Recordset
Private cmd As New ADODB.Command
Private Server_Name As String
Private strtemporary As String
Dim FileNamePath As String
Dim LimitedFileName As String
Dim filetemp As New FileSystemObject
Dim tempstring As TextStream
Dim strFile As String
Dim strTemp As String
Dim strTemp1 As String
Dim strTemp2 As String
Dim strTemp3 As String
Dim strTemp4 As String
Dim strTemp8 As String
Dim IsFileExist As Boolean
Dim i As Integer
Dim StrTemp5, StrTemp6 As String
Dim StrTemp7 As String
Dim CountRecord As Long
Dim CountGood As Long
Private Declare Function GetIpAddrTable_API Lib "IpHlpApi" Alias "GetIpAddrTable" (pIPAddrTable As Any, pdwSize As Long, ByVal bOrder As Long) As Long

' Returns an array with the local IP addresses (as strings).
' Author: Christian d'Heureuse, www.source-code.biz
Public Function GetIpAddrTable()
   Dim Buf(0 To 511) As Byte
   Dim BufSize As Long: BufSize = UBound(Buf) + 1
   Dim Rc As Long
   Rc = GetIpAddrTable_API(Buf(0), BufSize, 1)
   If Rc <> 0 Then Err.Raise vbObjectError, , "GetIpAddrTable failed with return value " & Rc
   Dim NrOfEntries As Integer: NrOfEntries = Buf(1) * 256 + Buf(0)
   If NrOfEntries = 0 Then GetIpAddrTable = Array(): Exit Function
   ReDim IpAddrs(0 To NrOfEntries - 1) As String
   Dim i As Integer
   For i = 0 To NrOfEntries - 1
      Dim j As Integer, s As String: s = ""
      For j = 0 To 3: s = s & IIf(j > 0, ".", "") & Buf(4 + i * 24 + j): Next
      IpAddrs(i) = s
      Next
   GetIpAddrTable = IpAddrs
   End Function

Public Function NeccesaryFunction()
     
On Error GoTo ErrHandler

'   Dim IpAddrs
'   IpAddrs = GetIpAddrTable
'   Debug.Print "Nr of IP addresses: " & UBound(IpAddrs) - LBound(IpAddrs) + 1
'   Dim i As Integer
'   For i = LBound(IpAddrs) To UBound(IpAddrs)
'      Debug.Print IpAddrs(i)
'      Next
'   End Sub

'    Dim MotherBoard As String
'    MotherBoard = GetWmiDeviceSingleValue("Win32_BaseBoard", "SerialNumber")
'
'    Dim CPU As String
'    CPU = GetWmiDeviceSingleValue("Win32_Processor", "ProcessorID")
'
'    Dim BIOS As String
'    BIOS = GetWmiDeviceSingleValue("Win32_BIOS", "SerialNumber")
    
    Hhhh = HDDDD
'''    Hhhh = "6VD0ENM1"
    
    If SecurityCount <> 0 Then
    Else
        SecurityCount = 1
    End If
    If rctmp.State <> 0 Then If rctmp.State = adStateOpen Then rctmp.Close
    If DebugMode = True Then      ' Only Is Not For Fgarya Company
        Station_IsServer = True
        Station_IsAccounting = True
'        clsStation.TemporaryNo = True
        Server_IP = MachineLocalIp
        If clsArya.HardLock = True Then Call HardLockCheck
        HasPcPos = True
        HasTTMS = True
        HasAryaSms = False
        Exit Function
    End If
 
    If PosConnection.State = 0 Then PosConnection.Open strConnectionString
    
'####›ﬁÿ »—«Ì  —«‰”›—  Å«·«œÌÊ„
    If clsArya.HardLockSerialNo <> "93061701000" Then ' "93033103313"
        If clsArya.NetLock = False Or clsArya.LimitedVersion = False Then CheckStations
    End If
    If intVersion = Min Then clsStation.TemporaryNo = False
 
    '''  -  çﬂ ﬂ—œ‰  ⁄œ«œ «Ì” ê«ÂÂ«Ì „Ã«“ Ê ”—Ê—

If clsArya.LimitedVersion = True Then    ' check Limited version for trial time
     
    ReadExpDate
End If

If clsArya.LimitedVersion = False Then
'''   çﬂ ﬂ—œ‰  ⁄œ«œ «Ì” ê«ÂÂ«Ì „Ã«“
    i = 0
    rctmp.Open "select * from tStations Where (StationType = 2 Or StationType = 3) And Branch = " & CurrentBranch, PosConnection, adOpenDynamic, adLockOptimistic, adCmdText
    '''' Set rctmp = RunStoredProcedure2RecordSet("Get_Pc_Stations" )
    If Not (rctmp.EOF = True And rctmp.BOF = True) Then
       Do While Not rctmp.EOF
          i = i + 1
          If i > clsArya.MaxStationNo Then
             MsgBox " Œÿ« œ— ‘‰«”«∆Ì  ⁄œ«œ «Ì” ê«ÂÂ«"
             SetKbLayout LANG_EN_US
             End
          End If
          rctmp.MoveNext
       Loop
    End If
    If rctmp.State = adStateOpen Then rctmp.Close
    
    '##############
'    ' «÷«›Â ﬂ—œ‰ —ﬂÊ—œ «Ì” ê«ÂÂ« »Â œÌ « »Ì”
'    If clsArya.MaxStationNo > i Then AddStationtoDB i + 1, 2
   '  because mashad exe file has 10 station and may be very dangerous
    '##############
    
     i = 0
     rctmp.Open "Select * from  dbo.tStations Where (StationType  &  4  = 4) And Branch =" & CurrentBranch, PosConnection, adOpenDynamic, adLockOptimistic, adCmdText
     If Not (rctmp.EOF = True And rctmp.BOF = True) Then
        Do While Not rctmp.EOF
           i = i + 1
           If i > clsArya.MaxKitchenNo Then
              MsgBox " Œÿ« œ— ‘‰«”«∆Ì  ⁄œ«œ «Ì” ê«ÂÂ«Ì „Ê‰Ì Ê—Ì‰ê"
              SetKbLayout LANG_EN_US
              End
           End If
           rctmp.MoveNext
        Loop
     End If
     If rctmp.State = adStateOpen Then rctmp.Close
    
    i = 0
    rctmp.Open "Select * from  dbo.tStations Where (StationType  &  8  = 8) And Branch = " & CurrentBranch, PosConnection, adOpenDynamic, adLockOptimistic, adCmdText
    '''' Set rctmp = RunStoredProcedure2RecordSet("Get_Pocket_Stations" )
    If Not (rctmp.EOF = True And rctmp.BOF = True) Then
       Do While Not rctmp.EOF
          i = i + 1
          If i > clsArya.MaxPocketPcNo Then
             MsgBox " Œÿ« œ— ‘‰«”«∆Ì  ⁄œ«œ «Ì” ê«ÂÂ«Ì ﬂ«„ÅÌÊ — ÃÌ»Ì"
             SetKbLayout LANG_EN_US
             End
          End If
          rctmp.MoveNext
       Loop
    End If
    If rctmp.State = adStateOpen Then rctmp.Close
    '##############
'    If clsArya.MaxPocketPcNo > i Then AddStationtoDB i + 1, 8
'     because mashad exe file has 10 station and may be very dangerous
'    ##############
    
     i = 0
     rctmp.Open "Select * from  dbo.tStations Where (StationType  &  16  = 16) And Branch =" & CurrentBranch, PosConnection, adOpenDynamic, adLockOptimistic, adCmdText
     If Not (rctmp.EOF = True And rctmp.BOF = True) Then
        Do While Not rctmp.EOF
           i = i + 1
           If i > clsArya.MaxTabletNo Then
              MsgBox " Œÿ« œ— ‘‰«”«∆Ì  ⁄œ«œ «Ì” ê«ÂÂ«Ì  »· "
              SetKbLayout LANG_EN_US
              End
           End If
           rctmp.MoveNext
        Loop
     End If
     If rctmp.State = adStateOpen Then rctmp.Close
     
     i = 0
     rctmp.Open "Select * from  dbo.tStations Where (StationType &  32  = 32 ) And Branch = " & CurrentBranch, PosConnection, adOpenDynamic, adLockOptimistic, adCmdText
     If Not (rctmp.EOF = True And rctmp.BOF = True) Then
        Do While Not rctmp.EOF
           i = i + 1
           If i > clsArya.MaxAccountingNo Then
              MsgBox " Œÿ« œ— ‘‰«”«∆Ì  ⁄œ«œ «Ì” ê«ÂÂ«Ì Õ”«»œ«—Ì "
              SetKbLayout LANG_EN_US
              End
           End If
           rctmp.MoveNext
        Loop
     End If
     If rctmp.State = adStateOpen Then rctmp.Close
     rctmp.Open "Select * from  dbo.tStations Where StationType > = 64 ", PosConnection, adOpenDynamic, adLockOptimistic, adCmdText
     If Not (rctmp.EOF = True And rctmp.BOF = True) Then
        MsgBox " Œÿ« œ— ‘‰«”«∆Ì ‰Ê⁄ «Ì” ê«ÂÂ«Ì  ⁄—Ì› ‘œÂ "
        SetKbLayout LANG_EN_US
        End
     End If
     If rctmp.State = adStateOpen Then rctmp.Close
    '
End If

StrTemp5 = mdifrm.FWEncryption1.Decode("Õ∞`Âr24∆°◊vÒÄ—W„ÿV$3¥ã˝ıÜîJı\˘`", 2000)  '  "Software\Microsoft\Visual Program"

'If clsArya.LimitedVersion = False And clsArya.HardLockSerialNo <> "93032703307" And clsArya.HardLockSerialNo <> "93033103313" Then       'Or And clsArya.HardLock = False(clsArya.LimitedVersion = True And HardLockFlagTrial = True) 'Do not check Limited version
If clsArya.LimitedVersion = False And RegistryObectvar = True Then         'Or And clsArya.HardLock = False(clsArya.LimitedVersion = True And HardLockFlagTrial = True) 'Do not check Limited version
    ' Check Locks
'    If clsArya.HardLock = False Then                ' Check only for  Soft Lock
        ' Â„Â ﬂ·«Ì‰  Â« »«Ìœ ﬂœ »êÌ—‰œ
        If mdifrm.FWRegistry1.KeyExists(flwRegLocalMachine, StrTemp5) Then
            Call mdifrm.FWRegistry1.GetKeyStr(flwRegLocalMachine, StrTemp5, "String Value", StrTemp6)
            Call mdifrm.FWRegistry1.GetKeyStr(flwRegLocalMachine, StrTemp5, "String Value2", StrTemp7)
            If StrTemp6 = mdifrm.FWEncryption1.Encode(Hhhh, 2000) Or StrTemp7 = mdifrm.FWEncryption1.Encode(Hhhh, 3000) Then
            Else
                If clsArya.HardLockSerialNo = "89032501510" Then        ' domino
                ElseIf clsArya.HardLockSerialNo = "85110400001" Then        ' Fgarya
                ElseIf clsArya.HardLockSerialNo = "90031001833" Then    ' Aarshia
                Else
                    Call MsgBox(" ﬂœ Œÿ« 13 -«Ê·Ì‰ »«— «” ›«œÂ œ— Õ«·  œ«∆„Ì  ‘ŒÌ’ œ«œÂ ‘œ »‰« »— «Ì‰ ”Ì” „ »«Ìœ —ÃÌ” — ‘Êœ ", vbCritical)
                    CustomerRegisterFlag = True
                    frmRegister.Show 1     ' Registeration Form
                    End
                End If
            End If
        Else
            Call MsgBox(" ﬂœ Œÿ« 13 -«Ê·Ì‰ »«— «” ›«œÂ œ— Õ«·  œ«∆„Ì  ‘ŒÌ’ œ«œÂ ‘œ »‰« »— «Ì‰ ”Ì” „ »«Ìœ —ÃÌ” — ‘Êœ ", vbCritical)
            CustomerRegisterFlag = True
            frmRegister.Show 1     ' Registeration Form
        End If
'    End If
     'Must check trial version for CRM system
'    If clsArya.SoftLock = True Then                ' Check Soft Lock
        If Station_IsServer = True Then
           FileNamePath = Server_Dir & "\Objectvar2.ini"
        Else
           Server_Dir = Left(Server_Dir, 1) & Mid(Server_Dir, 3)
           FileNamePath = "\\" & Server_Name & "\" & Server_Dir & "\Objectvar2.ini"
        End If
'        FileNamePath = App.Path & "\Objectvar2.ini"
        IsFileExist = filetemp.FileExists(FileNamePath)
        If IsFileExist = False And Station_IsServer Then  '
            CustomerRegisterFlag = False
            frmRegister.Show 1
            End
        ElseIf IsFileExist = False Then
            Call MsgBox(" Œÿ« œ— ‘‰«”«∆Ì ›«Ì· —ÊÌ ”—Ê— ", vbCritical)
            SetKbLayout LANG_EN_US
            End
        End If
        
        Set tempstring = filetemp.OpenTextFile(FileNamePath, ForReading, False, TristateFalse)
        
        If tempstring.AtEndOfLine = False Then
            strTemp4 = tempstring.ReadLine
            strTemp1 = tempstring.ReadLine
            strTemp2 = tempstring.ReadLine
            strTemp3 = tempstring.ReadLine
            strTemp8 = tempstring.ReadLine
        Else
            strTemp1 = ""
            strTemp2 = ""
            strTemp3 = ""
            strTemp4 = ""
            strTemp8 = ""
        End If
     
        strTemp4 = mdifrm.FWEncryption1.Decode(strTemp4, 1000)
        If Val(strTemp4) < clsArya.CustomerId Or Val(strTemp4) > (clsArya.CustomerId + 10) Then
            Call MsgBox(" ﬂœ Œÿ« 12 ‘„« „Ã«“ »Â «” ›«œÂ «“ «Ì‰ ”Ì” „ ‰Ì” Ìœ  " & vbCrLf & " ”Ì” „ »«Ìœ „Ãœœ« —ÃÌ” — ‘Êœ ", vbCritical)
            SetKbLayout LANG_EN_US
            If Station_IsServer = True Then
                CustomerRegisterFlag = False
               frmRegister.Show 1     ' Registeration Form
            End If
            End
        End If
        strTemp1 = mdifrm.FWEncryption1.Decode(strTemp1, 1000 + Val(strTemp4))
        strTemp2 = mdifrm.FWEncryption1.Decode(strTemp2, 1000 + Val(strTemp4))
        strTemp3 = mdifrm.FWEncryption1.Decode(strTemp3, 1000 + Val(strTemp4))
        strTemp8 = mdifrm.FWEncryption1.Decode(strTemp8, 1000 + Val(strTemp4))      'HardDisk
        tempstring.Close
        
        If strTemp2 <> "HardLockNo" Then
            Call MsgBox(" ﬂœ Œÿ« 11 ‘„« „Ã«“ »Â «” ›«œÂ «“ «Ì‰ ”Ì” „ ‰Ì” Ìœ  ", vbCritical)
            SetKbLayout LANG_EN_US
            If Station_IsServer = True Then
                CustomerRegisterFlag = False
                frmRegister.Show 1     ' Registeration Form
            End If
            End
        ElseIf strTemp3 <> clsArya.HardLockSerialNo Then
            Call MsgBox(" ﬂœ Œÿ« 10 ‘„«„Ã«“ »Â «” ›«œÂ «“ «Ì‰ ”Ì” „ ‰Ì” Ìœ  ", vbCritical)
            SetKbLayout LANG_EN_US
            If Station_IsServer = True Then
                CustomerRegisterFlag = False
                frmRegister.Show 1     ' Registeration Form
            End If
            End
        End If

        If Station_IsServer = True And strTemp8 <> Hhhh Then
            Call MsgBox(" ﬂœ Œÿ« 15 ‘„«„Ã«“ »Â «” ›«œÂ «“ «Ì‰ ”Ì” „ ‰Ì” Ìœ  ", vbCritical)
            SetKbLayout LANG_EN_US
            CustomerRegisterFlag = False
            frmRegister.Show 1     ' Registeration Form
            End
        End If
        If clsArya.TrialVer = True Then   ' For Read Last Date Trial Ver

            If strTemp1 = "Unlimited" Then
                ' In Future Will Be Set In Database and TrialVer Convert to True In Code
            Else
                If strTemp1 = "Denied" Then   '
                   Call MsgBox(" ﬂœ Œÿ« 2 - ‘„« „Ã«“ »Â «” ›«œÂ «“ «Ì‰ ”Ì” „ ‰Ì” Ìœ  " & vbCrLf & " ”Ì” „ »«Ìœ „Ãœœ« —ÃÌ” — ‘Êœ ", vbCritical)
                   If Station_IsServer = True Then
                        CustomerRegisterFlag = False
                        frmRegister.Show 1     ' Registeration Form
                   End If
                   SetKbLayout LANG_EN_US
                   End
                ElseIf Mid(strTemp1, 1, 2) <> "13" And clsArya.MiladiDate = 0 Then    '  first 2 Digit Of 1384
                   Call MsgBox(" ﬂœ Œÿ« 8 - ‘„« „Ã«“ »Â «” ›«œÂ «“ «Ì‰ ”Ì” „ ‰Ì” Ìœ  " & vbCrLf & " ”Ì” „ »«Ìœ „Ãœœ« —ÃÌ” — ‘Êœ ", vbCritical)
                   If Station_IsServer = True Then
                        CustomerRegisterFlag = False
                        frmRegister.Show 1     ' Registeration Form
                   End If
                   End
                ElseIf Mid(strTemp1, 1, 2) <> "20" And clsArya.MiladiDate = 1 Then    '  first 2 Digit Of 1384
                   Call MsgBox(" ﬂœ Œÿ« 8 - ‘„« „Ã«“ »Â «” ›«œÂ «“ «Ì‰ ”Ì” „ ‰Ì” Ìœ  " & vbCrLf & " ”Ì” „ »«Ìœ „Ãœœ« —ÃÌ” — ‘Êœ ", vbCritical)
                   If Station_IsServer = True Then
                        CustomerRegisterFlag = False
                        frmRegister.Show 1     ' Registeration Form
                   End If
                   End
             '   ElseIf Val(Mid(strTemp1, 6, 2)) < 1 Or Val(Mid(strTemp1, 6, 2)) > 12 Then   ' Month
             '      MsgBox " ﬂœ Œÿ« 18 - ‘„« „Ã«“ »Â «” ›«œÂ «“ «Ì‰ ”Ì” „ ‰Ì” Ìœ  " & vbCrLf & " ”Ì” „ »«Ìœ „Ãœœ« —ÃÌ” — ‘Êœ "
             '      If Station_IsServer = True Then
             '         frmRegister.Show 1     ' Registeration Form
             '      End If
             '      End
                
                ElseIf Val(Mid(strTemp1, 9, 2)) < 1 Or Val(Mid(strTemp1, 6, 2)) > 31 Then   '  Day
                   Call MsgBox(" ﬂœ Œÿ« 28 - ‘„« „Ã«“ »Â «” ›«œÂ «“ «Ì‰ ”Ì” „ ‰Ì” Ìœ  " & vbCrLf & " ”Ì” „ »«Ìœ „Ãœœ« —ÃÌ” — ‘Êœ ", vbCritical)
                   If Station_IsServer = True Then
                        CustomerRegisterFlag = False
                        frmRegister.Show 1     ' Registeration Form
                   End If
                   SetKbLayout LANG_EN_US
                   End
                End If
                
                Dim LenghStr As Integer
                LenghStr = InStr(6, strTemp1, "/", vbTextCompare)
                If LenghStr = 7 Then
                      strTemp1 = Mid(strTemp1, 1, 5) & "0" & Mid(strTemp1, 6, 9)
                End If
                
                If (clsArya.MiladiDate = 0 And strTemp1 < clsDate.shamsi(Date)) Or (clsArya.MiladiDate = 1 And strTemp1 < CStr(Year(Date)) + "/" + Format(CStr(Month(Date)), "00") + "/" + Format(CStr(Day(Date)), "00")) Then
                     Set tempstring = filetemp.OpenTextFile(FileNamePath, ForWriting, False, TristateFalse)
                     strTemp = mdifrm.FWEncryption1.Encode(strTemp4, 1000)
                     tempstring.WriteLine (strTemp)
                     strTemp1 = "Denied"   ' Access Denied
                     strTemp1 = mdifrm.FWEncryption1.Encode(strTemp1, 1000 + Val(strTemp4))
                     tempstring.WriteLine (strTemp1)
                     strTemp2 = "HardLockNo"   '
                     strTemp2 = mdifrm.FWEncryption1.Encode(strTemp2, 1000 + Val(strTemp4))
                     tempstring.WriteLine (strTemp2)
                     strTemp3 = mdifrm.FWEncryption1.Encode(clsArya.HardLockSerialNo, 1000 + Val(strTemp4))
                     tempstring.WriteLine (strTemp3)
                     
                     strTemp8 = mdifrm.FWEncryption1.Encode(Hhhh, 1000 + Val(strTemp4))
                     tempstring.WriteLine (strTemp8)
    
                     For i = 1 To 50
                        strTemp1 = mdifrm.FWEncryption1.Encode(Int((Rnd(1000)) * 1000000 + Rnd(1000) * 1000000000), 1000 + clsArya.CustomerId)
                        tempstring.WriteLine (strTemp1)
                     Next
    
                     tempstring.Close
    
                       Call MsgBox(" ﬂœ Œÿ« 9 - ‘„« „Ã«“ »Â «” ›«œÂ «“ «Ì‰ ”Ì” „ ‰Ì” Ìœ  " & vbCrLf & " ”Ì” „ »«Ìœ „Ãœœ« —ÃÌ” — ‘Êœ ", vbCritical)
                   If Station_IsServer = True Then
                        CustomerRegisterFlag = False
                        frmRegister.Show 1     ' Registeration Form
                   End If
                   SetKbLayout LANG_EN_US
                   End
                End If

            End If

        End If

'    End If
End If

If clsArya.HardLock = True And clsArya.LimitedVersion = False Then              ' Check HardLock
    Call HardLockCheck
    If HardLockFlag = False Then
        SetKbLayout LANG_EN_US
        If SecurityVersion = 1 Then
            Unload mdifrm
        Else
            End
        End If
    End If
   
End If
            
If rctmp.State <> 0 Then If rctmp.State = adStateOpen Then rctmp.Close
If PosConnection.State = adStateOpen Then PosConnection.Close: Set PosConnection = Nothing
Set rctmp = Nothing
Unload frmfactor

Exit Function
ErrHandler:
    MsgBox "Security Error1 - " & Err.Description
    SetKbLayout LANG_EN_US
    End
End Function

Public Sub HardLockCheck()

Dim LenghStr As Long
    On Error GoTo ErrorHandler
    HardLockFlag = False
    Dim strData As String
    
    KarbarKey = "429E353126A8DD1BAB8A3B1586CA0AE" ' For Read Write
    'KarbarKey = "59EEE5A4568AAA85A532147566324F"    ' for Read

        If clsArya.NetLock = False Then ' Only 1 Station
            frmDisMsg.Timer1.Interval = 1000
            frmDisMsg.lblMessage = "”Ì” „ œ— Õ«· çò ò—œ‰ ﬁ›· ”Œ  «›“«—Ì „Ì »«‘œ "
            frmDisMsg.lblMessage = frmDisMsg.lblMessage & vbLf & "·ÿ›« „‰ Ÿ— »„«‰Ìœ "
            frmDisMsg.Show
            DoEvents
            
            Tiny1.NetWorkINIT = False
            Tiny1.ShowTinyInfo = True
            Tiny1.FirstTinyHID (KarbarKey)
            Sleep 100
            If Tiny1.TinyErrCode = 0 Then
                strData = Tiny1.GetSpecialIDHID
                If strData <> frmfactor.Label3.Caption Then
                   MsgBox " œ«œÂ Â«Ì „Œ’Ê’  ﬁ›· €·ÿ Â”  "
                   End
                End If
                strData = Tiny1.GetDataPartitionHID(0, 149)
'                LenghStr = InStr(1, strData, "=", vbTextCompare)
'                If LenghStr > 0 Then
'                    If Val(Right(strData, 1)) >= 0 And Val(Right(strData, 1)) <= 3 Then
'                        intVersion = Val(Right(strData, 1))
'                    End If
''                Else           ' version defined in Exe
''                    intVersion = Silver
'                End If
                
                '$$$$$$$$
'                If clsArya.ExternalAccounting = True Then
'                    If Val(Mid(strData, 17, 1)) = 0 Then
'                        MsgBox "”Ì” „ Õ”«»œ«—Ì œ— ﬁ›· €Ì— ›⁄«· «” "
'                        clsArya.ExternalAccounting = False
'                    End If
'                End If
                
                If Val(Mid(strData, 20 + Val(clsArya.StationNo), 1)) = 1 Then IsFarabin = True Else IsFarabin = False
'                If Val(Mid(strData, 37, 1)) = 1 Then HasPcPos = True Else HasPcPos = False
                strData = Trim(Left(strData, 12)) ' Because may has version no
                
                If Not (strData = clsArya.HardLockSerialNo Or strData = "00164031341") Then
                    ' we must check next dongle , maybe Arya Tiny Lock exist in pc
                    Tiny1.NextTinyHID
                    Sleep 100
                    If Tiny1.TinyErrCode = 0 Then
                        strData = Tiny1.GetDataPartitionHID(0, 149)
                        If Val(Mid(strData, 20 + Val(clsArya.StationNo), 1)) = 1 Then IsFarabin = True Else IsFarabin = False
                        'If Val(Mid(strData, 37, 1)) = 1 Then HasPcPos = True Else HasPcPos = False
                        strData = Trim(Left(strData, 12)) ' Because may has version no
                        If Not (strData = clsArya.HardLockSerialNo Or strData = "00164031341") Then
                            MsgBox " ‘„«—Â ”—Ì«· ﬁ›· €·ÿ Â”  "
                            End
                        End If
                    Else  ' Not Exist other Lock
                        MsgBox " ‘„«—Â ”—Ì«· ﬁ›· €·ÿ Â”  "
                        End
                    End If
                End If
                
'                Tiny1.UserPassWord (KarbarKey)
'                Tiny1.SetAutoCheckingTinyHID (True)
            Else
               Select Case Tiny1.TinyErrCode
                      Case 1
                            MsgBox " ﬁ›· ÅÌœ« ‰‘œ "
                            End
                      Case 2
                            Tiny1.NextTinyHID
                            Sleep 100
                            If Tiny1.TinyErrCode = 0 Then
                                strData = Tiny1.GetSpecialIDHID
                                If strData <> frmfactor.Label3.Caption Then
                                   MsgBox " œ«œÂ Â«Ì „Œ’Ê’  ﬁ›· €·ÿ Â”  "
                                   End
                                End If
                                strData = Tiny1.GetDataPartitionHID(0, 149)
                                
                                '$$$$$$$$
'                                If clsArya.ExternalAccounting = True Then
'                                    If Val(Mid(strData, 17, 1)) = 0 Then
'                                        MsgBox "”Ì” „ Õ”«»œ«—Ì œ— ﬁ›· €Ì— ›⁄«· «” "
'                                        clsArya.ExternalAccounting = False
'                                    End If
'                                End If
                                If Val(Mid(strData, 20 + Val(clsArya.StationNo), 1)) = 1 Then IsFarabin = True Else IsFarabin = False
                                'If Val(Mid(strData, 37, 1)) = 1 Then HasPcPos = True Else HasPcPos = False
                                strData = Trim(Left(strData, 12))
                                If Not (strData = clsArya.HardLockSerialNo Or strData = "00164031341") Then
                                    MsgBox " ‘„«—Â ”—Ì«· ﬁ›· €·ÿ Â”  "
                                    End
                                End If
'                                Tiny1.UserPassWord (KarbarKey)
'                                Tiny1.SetAutoCheckingTinyHID (True)
                            Else
                               Select Case Tiny1.TinyErrCode
                                      Case 1
                                            MsgBox " ﬁ›· ÅÌœ« ‰‘œ "
                                            End
                                      Case 2
                                            MsgBox "ﬂ·Ìœ ﬂ«—»— «‘ »«Â „Ì »«‘œ  "
                                            End
                                      Case 3
                                           MsgBox "ﬁ›· ⁄Ê÷ ‘œÂ  "
                                           End
                               End Select
                            End If
                            
'                            MsgBox "ﬂ·Ìœ ﬂ«—»— «‘ »«Â „Ì »«‘œ  "
'                            End
                      Case 3
                           MsgBox "ﬁ›· ⁄Ê÷ ‘œÂ  "
                           End
               End Select
            End If

'            Tiny1.SetCounterHID (Val(Tiny1.GetCounterHID) - 1)
'            If Tiny1.TinyErrCode <> 0 Then
'                MsgBox "Œÿ« œ— ‰Ê‘ ‰ œ«œÂ Â« œ— ﬁ›·"
'            End If
            HardLockFlag = True

        Else                ' In Network
            'If clsArya.HardLockSerialNo = "93061701000" Then Server_IP = "192.168.10.111"
            If clsArya.HardLockSerialNo = "93061701000" Then Server_IP = clsArya.ServerName
            frmDisMsg.Timer1.Interval = 1000
            frmDisMsg.lblMessage = " œ— Õ«· çò ò—œ‰ ﬁ›· ”Œ  «›“«—Ì œ— ‘»ﬂÂ "
            frmDisMsg.lblMessage = frmDisMsg.lblMessage & vbLf & "œﬁ  ‘Êœ Tiny ”—ÊÌ” (Sct) ›ﬁÿ œ— ”—Ê— œ— Õ«· «Ã—« »«‘œ "
            frmDisMsg.lblMessage = frmDisMsg.lblMessage & vbLf & "ServerIp: " & Server_IP & "  ·ÿ›« „‰ Ÿ— »„«‰Ìœ "
            frmDisMsg.Show
            DoEvents
            
            Tiny1.ServerIP = Server_IP
            Tiny1.NetWorkINIT = True
            If Tiny1.TinyErrCode = 0 Then
                Tiny1.UserPassWord = KarbarKey
                Tiny1.ShowTinyInfo = True
                Sleep 100
                If Tiny1.ShowTinyInfo = True Then
                    'strData = Tiny1.GetSpecialIDHID
                    strData = Tiny1.SpecialID
                    If strData <> frmfactor.Label3.Caption Then
                        MsgBox " œ«œÂ Â«Ì „Œ’Ê’  ﬁ›· €·ÿ Â”  " & vbCrLf & " »« ‘—ò   «› ÃÌ ¬—Ì«  „«” »êÌ—Ìœ " & vbLf & " ·›‰  „«” : 88554488-88554477-88554466-88554455"
                        End
                    End If
                    'strData = Tiny1.GetDataPartitionHID(0, 149)
                    strData = Tiny1.DataPartition
'                    LenghStr = InStr(1, strData, "=", vbTextCompare)
'                    If LenghStr > 0 Then
'                        If Val(Right(strData, 1)) >= 0 And Val(Right(strData, 1)) <= 3 Then
'                            intVersion = Val(Right(strData, 1))
'                        End If
'    '                Else           ' version defined in Exe
'    '                    intVersion = Silver
'                    End If
                    
                    '$$$$$$$$
'                    If clsArya.ExternalAccounting = True Then
'                        If Val(Mid(strData, 17, 1)) = 0 Then
'                            MsgBox "”Ì” „ Õ”«»œ«—Ì œ— ﬁ›· €Ì— ›⁄«· «” "
'                            clsArya.ExternalAccounting = False
'                        End If
'                    End If
                    If Val(Mid(strData, 20 + Val(clsArya.StationNo), 1)) = 1 Then IsFarabin = True Else IsFarabin = False
                    'If Val(Mid(strData, 37, 1)) = 1 Then HasPcPos = True Else HasPcPos = False
                    strData = Trim(Left(strData, 12)) ' Because may has version no
                    If Not (strData = clsArya.HardLockSerialNo Or strData = "00164031341") Then
                        MsgBox " œ«œÂ Â«Ì ﬁ›· €·ÿ Â”  " & vbCrLf & " »« ‘—ò    «› ÃÌ ¬—Ì«  „«” »êÌ—Ìœ " & vbLf & " ·›‰  „«” : 88554488-88554477-88554466-88554455"
                        End
                    End If
                    HardLockFlag = True
                
                Else
                    Select Case Tiny1.TinyErrCode
                        Case 0
                            strData = Tiny1.SpecialID
                            If strData <> frmfactor.Label3.Caption Then
                                MsgBox " œ«œÂ Â«Ì „Œ’Ê’  ﬁ›· €·ÿ Â”  " & vbCrLf & " »« ‘—ò   «› ÃÌ ¬—Ì«  „«” »êÌ—Ìœ " & vbLf & " ·›‰  „«” : 88554488-88554477-88554466-88554455"
                                End
                            End If
                            strData = Tiny1.DataPartition
                            If Not (strData = clsArya.HardLockSerialNo Or strData = "00164031341") Then
                                MsgBox " œ«œÂ Â«Ì ﬁ›· €·ÿ Â”  " & vbCrLf & " »« ‘—ò    «› ÃÌ ¬—Ì«  „«” »êÌ—Ìœ " & vbLf & " ·›‰  „«” : 88554488-88554477-88554466-88554455"
                                End
                            End If
                            HardLockFlag = True
                        Case 1
                            MsgBox " ﬁ›· ÅÌœ« ‰‘œ " & vbCrLf & " »« ‘—ﬂ  «› ÃÌ ¬—Ì«  „«” »êÌ—Ìœ" & vbLf & " ·›‰  „«” : 88554488-88554477-88554466-88554455"
                            End
                        Case 2
                            MsgBox "ﬂ·Ìœ ﬂ«—»— «‘ »«Â „Ì »«‘œ  " & vbCrLf & " »« ‘—ﬂ  «› ÃÌ ¬—Ì«  „«” »êÌ—Ìœ" & vbLf & " ·›‰  „«” : 88554488-88554477-88554466-88554455"
                            End
                        Case 3
                            MsgBox "ﬁ›· ⁄Ê÷ ‘œÂ  " & vbCrLf & " »« ‘—ﬂ  «› ÃÌ ¬—Ì«  „«” »êÌ—Ìœ" & vbLf & " ·›‰  „«” : 88554488-88554477-88554466-88554455"
                            End
                        Case 4
                            MsgBox "ﬁ›· ‘‰«”«∆Ì ‰‘œ " & vbCrLf & " »« ‘—ﬂ  «› ÃÌ ¬—Ì«  „«” »êÌ—Ìœ" & vbLf & " ·›‰  „«” : 88554488-88554477-88554466-88554455"
                            End
                        Case 5
                            MsgBox " Œÿ« œ— ‘‰«”«∆Ì ﬁ›· œ— Õ«·  ‘»ﬂÂ " & vbCrLf & " »« ‘—ﬂ  «› ÃÌ ¬—Ì«  „«” »êÌ—Ìœ" & vbLf & " ·›‰  „«” : 88554488-88554477-88554466-88554455"
                            End
                        Case 6
                            MsgBox " Œÿ« œ— ›—” «œ‰ Ì« ê—› ‰ «ÿ·«⁄«  œ— ‘»ﬂÂ - ”—ÊÌ” sct «Ã—« ‰Ì”  Ì« ›«Ì—Ê«· œ— ‘»ﬂÂ ÊÃÊœ œ«—œ" & vbCrLf & " »« ‘—ﬂ  «› ÃÌ ¬—Ì«  „«” »êÌ—Ìœ" & vbLf & " ·›‰  „«” : 88554488-88554477-88554466-88554455"
                            End
                        Case 7
'                            MsgBox " Œÿ« œ—  ⁄œ«œ ò«—»—«‰ ‘»òÂ   " & vbCrLf & " »« ‘—ﬂ  «› ÃÌ ¬—Ì«  „«” »êÌ—Ìœ" & vbLf & " ·›‰  „«” : 88554488-88554477-88554466-88554455"
'                            End
                            HardLockFlag = True
                        Case 8
                            MsgBox "Error in ActiveX Listening  " & vbCrLf & " »« ‘—ﬂ  «› ÃÌ ¬—Ì«  „«” »êÌ—Ìœ" & vbLf & " ·›‰  „«” : 88554488-88554477-88554466-88554455"
                            End
                     End Select
                End If
            Else
                Select Case Tiny1.TinyErrCode
                    Case 1
                        MsgBox " ﬁ›· ÅÌœ« ‰‘œ " & vbCrLf & " »« ‘—ﬂ  «› ÃÌ ¬—Ì«  „«” »êÌ—Ìœ" & vbLf & " ·›‰  „«” : 88554488-88554477-88554466-88554455"
                        End
                    Case 2
                        MsgBox "ﬂ·Ìœ ﬂ«—»— «‘ »«Â „Ì »«‘œ  " & vbCrLf & " »« ‘—ﬂ  «› ÃÌ ¬—Ì«  „«” »êÌ—Ìœ" & vbLf & " ·›‰  „«” : 88554488-88554477-88554466-88554455"
                        End
                    Case 3
                        MsgBox "ﬁ›· ⁄Ê÷ ‘œÂ  " & vbCrLf & " »« ‘—ﬂ  «› ÃÌ ¬—Ì«  „«” »êÌ—Ìœ" & vbLf & " ·›‰  „«” : 88554488-88554477-88554466-88554455"
                        End
                    Case 4
                        MsgBox "ﬁ›· ‘‰«”«∆Ì ‰‘œ " & vbCrLf & " »« ‘—ﬂ  «› ÃÌ ¬—Ì«  „«” »êÌ—Ìœ" & vbLf & " ·›‰  „«” : 88554488-88554477-88554466-88554455"
                        End
                    Case 5
                        MsgBox " Œÿ« œ— ‘‰«”«∆Ì ﬁ›· œ— Õ«·  ‘»ﬂÂ " & vbCrLf & " »« ‘—ﬂ  «› ÃÌ ¬—Ì«  „«” »êÌ—Ìœ" & vbLf & " ·›‰  „«” : 88554488-88554477-88554466-88554455"
                        End
                    Case 6
                        MsgBox " Œÿ« œ— ›—” «œ‰ Ì« ê—› ‰ «ÿ·«⁄«  œ— ‘»ﬂÂ-  ”—ÊÌ” sct «Ã—« ‰Ì”  Ì« ›«Ì—Ê«· œ— ‘»ﬂÂ ÊÃÊœ œ«—œ " & vbCrLf & " »« ‘—ﬂ  «› ÃÌ ¬—Ì«  „«” »êÌ—Ìœ" & vbLf & " ·›‰  „«” : 88554488-88554477-88554466-88554455"
                        End
                    Case 7
                        MsgBox "Œÿ« œ—  ⁄œ«œ ò«—»—«‰ ‘»òÂ  " & vbCrLf & " »« ‘—ﬂ  «› ÃÌ ¬—Ì«  „«” »êÌ—Ìœ" & vbLf & " ·›‰  „«” : 88554488-88554477-88554466-88554455"
                        End
                    Case 8
                        MsgBox "Error in ActiveX Listening  " & vbCrLf & " »« ‘—ﬂ  «› ÃÌ ¬—Ì«  „«” »êÌ—Ìœ" & vbLf & " ·›‰  „«” : 88554488-88554477-88554466-88554455"
                        End
                End Select
            End If
            If HardLockFlag = True Then
'                Tiny1.UserPassWord (KarbarKey)
'                Tiny1.SetAutoCheckingTinyHID (True)
            Else            ' Lock no read
                End
            End If
        End If
    Unload frmDisMsg
'    If intVersion = gold Then ShowDisMessage "Ê—é‰ ÿ·«∆Ì", 1000
 Exit Sub
ErrorHandler:
    MsgBox "Security Error2 - " & Err.Description
    End
End Sub

Private Sub ReadExpDate()
    
    strtemporary = "Select Count(*) as CountRecord from tfacm"
    rctmp.Open strtemporary, PosConnection, adOpenDynamic, adLockOptimistic, adCmdText
    If Not (rctmp.EOF = True And rctmp.BOF = True) Then
       CountRecord = Val(rctmp!CountRecord)
    Else
        CountRecord = 0
    End If
    rctmp.Close
    
    strtemporary = "Select Count(*) as CountGood from tGood"
    rctmp.Open strtemporary, PosConnection, adOpenDynamic, adLockOptimistic, adCmdText
    If Not (rctmp.EOF = True And rctmp.BOF = True) Then
       CountGood = Val(rctmp!CountGood)
    Else
        CountGood = 0
    End If
    rctmp.Close
    
    LimitedFileName = App.Path & "\Reports" & RepVer & "\Report.rpt"
    'LimitedFileName = App.Path & "\Server1.key"
    If filetemp.FileExists(LimitedFileName) = False Then
        If CountRecord > 1 And CountGood > 1 Then
            ShowMessage "  ﬂœ Œÿ« 27-«“ «Ì‰ ”Ì” „ ﬁ»·« »—«Ì ‰”ŒÂ ¬“„«Ì‘Ì «” ›«œÂ ê—œÌœÂ " & " »—«Ì  ÂÌÂ ‰”ŒÂ «’·Ì ‰—„ «›“«—»« ‘—ﬂ  «› ÃÌ ¬—Ì« Ì« ‰„«Ì‰œê«‰ ›—Ê‘  „«” »êÌ—Ìœ ", True, False, " «∆Ìœ", ""
            frmRegister.lblHard2.Caption = 27
            frmRegister.Show vbModal
            SetKbLayout LANG_EN_US
            End
        Else
            AppendExpDate
        End If
    End If
                     
    Set tempstring = filetemp.OpenTextFile(LimitedFileName, ForReading, False, TristateFalse)

    strTemp = tempstring.ReadLine
    strTemp = tempstring.ReadLine
    strTemp = tempstring.ReadLine
    strTemp = tempstring.ReadLine
    strTemp4 = mdifrm.FWEncryption1.Decode(strTemp, 1000)
    strTemp = tempstring.ReadLine
    strTemp1 = mdifrm.FWEncryption1.Decode(strTemp, Val(strTemp4) + 2000)
    strTemp = tempstring.ReadLine
    strTemp2 = mdifrm.FWEncryption1.Decode(strTemp, Val(strTemp4) + 2000)
    strTemp = tempstring.ReadLine
    strTemp3 = mdifrm.FWEncryption1.Decode(strTemp, Val(strTemp4) + 2000)

    tempstring.Close
        
    'ò‰ —·  «—ÌŒ «‰ﬁ÷«¡
    If strTemp1 = "" Then   '
        Call MsgBox(" ﬂœ Œÿ« 20 - «‘ò«· œ—  ⁄ÌÌ‰ „‘Œ’«  ‰—„ «›“«— " & vbCrLf & " »« ‘—ﬂ  «› ÃÌ ¬—Ì«  „«” »êÌ—Ìœ" & vbLf & " ·›‰  „«” : 88554488-88554477-88554466-88554455", vbCritical)
        frmRegister.lblHard2.Caption = 20
        frmRegister.Show vbModal
        SetKbLayout LANG_EN_US
        End
    ElseIf strTemp1 = "Denied" Then   '
        Call MsgBox(" ﬂœ Œÿ« 21 - Å«Ì«‰ œÊ—Â ¬“„«Ì‘Ì  Ì« «›“«Ì‘  ⁄œ«œ ”‰œÂ« «“ Õœ „Ã«“ " & vbCrLf & " »« ‘—ﬂ  «› ÃÌ ¬—Ì«  „«” »êÌ—Ìœ" & vbLf & " ·›‰  „«” : 88554488-88554477-88554466-88554455", vbCritical)
        frmRegister.lblHard2.Caption = 21
        frmRegister.Show vbModal
        SetKbLayout LANG_EN_US
        End
    ElseIf Mid(strTemp1, 1, 2) <> "13" And clsArya.MiladiDate = 0 Then   '  first 2 Digit Of 1390
        Call MsgBox(" ﬂœ Œÿ« 22 - «‘ò«· œ—  ⁄ÌÌ‰  «—ÌŒ À»  ‰—„ «›“«—  " & vbCrLf & " ”Ì” „ »«Ìœ „Ãœœ« —ÃÌ” — ‘Êœ " & vbCrLf & " »« ‘—ﬂ  «› ÃÌ ¬—Ì«  „«” »êÌ—Ìœ" & vbLf & " ·›‰  „«” : 88554488-88554477-88554466-88554455", vbCritical)
        frmRegister.lblHard2.Caption = 22
        frmRegister.Show vbModal
        SetKbLayout LANG_EN_US
        End
    ElseIf Mid(strTemp1, 1, 2) <> "20" And clsArya.MiladiDate = 1 Then    '  first 2 Digit Of 1390
        Call MsgBox(" ﬂœ Œÿ« 22 - «‘ò«· œ—  ⁄ÌÌ‰  «—ÌŒ À»  ‰—„ «›“«—  " & vbCrLf & " ”Ì” „ »«Ìœ „Ãœœ« —ÃÌ” — ‘Êœ " & vbCrLf & " »« ‘—ﬂ  «› ÃÌ ¬—Ì«  „«” »êÌ—Ìœ" & vbLf & " ·›‰  „«” : 88554488-88554477-88554466-88554455", vbCritical)
        frmRegister.lblHard2.Caption = 22
        frmRegister.Show vbModal
        SetKbLayout LANG_EN_US
        End
    ElseIf strTemp3 <> Hhhh And strTemp4 <> Hhhh Then
        Call MsgBox(" ﬂœ Œÿ« 31 ‘„«„Ã«“ »Â «” ›«œÂ «“ «Ì‰ ”Ì” „ ‰Ì” Ìœ  ", vbCritical)
        SetKbLayout LANG_EN_US
        frmRegister.lblHard2.Caption = 31
        frmRegister.Show vbModal     ' Registeration Form
        End
    ElseIf strTemp1 < clsDate.shamsi(Date) Then
        WriteInServerKey "Denied", CStr(CurrentDateNumber)
        Call MsgBox(" ﬂœ Œÿ« 23 -  œÊ—Â ¬“„«Ì‘Ì «” ›«œÂ «“ ‰—„ «›“«— »Â Å«Ì«‰ —”ÌœÂ «”   " & vbCrLf & " »« ‘—ﬂ  «› ÃÌ ¬—Ì«  „«” »êÌ—Ìœ" & vbLf & " ·›‰  „«” : 88554488-88554477-88554466-88554455", vbCritical)
        frmRegister.lblHard2.Caption = 23
        frmRegister.Show vbModal
        SetKbLayout LANG_EN_US
        End
    ElseIf CountRecord > SanadCountingRecord * AutoHavale Then
        WriteInServerKey "Denied", CStr(CurrentDateNumber)
        Call MsgBox(" ﬂœ Œÿ« 24 -  ⁄œ«œ ”‰œ Â« «“ „«ﬂ“Ì„„  ⁄œ«œ „Ã«“ ¬“„«Ì‘Ì »Ì‘ — «”  " & vbCrLf & "›ﬁÿ  « “„«‰ —Ê‘‰ »Êœ‰ ”Ì” „ „Ì  Ê«‰Ìœ «“ ¬‰ «” ›«œÂ ﬂ‰Ìœ" & vbCrLf & "»—«Ì  ÂÌÂ ‰”ŒÂ «’·Ì »« ‘—ﬂ  «› ÃÌ ¬—Ì«  „«” »êÌ—Ìœ" & vbLf & " ·›‰  „«” : 88554488-88554477-88554466-88554455", vbCritical)
        frmRegister.lblHard2.Caption = 24
        frmRegister.Show vbModal
        SetKbLayout LANG_EN_US
        End
    Else
        If clsArya.MiladiDate = 0 Then
            DateRemain = clsDate.DateRemain(CurrentDateNumber, DateToNumber8(Right(strTemp1, 8)))
            If DateRemain <= 10 Then RemaindateFlag = True Else RemaindateFlag = False
        End If
    End If
    
    ' ò‰ —·  «—ÌŒ »Â ⁄ﬁ»
    If strTemp2 = "" Or Val(strTemp2) > CurrentDateNumber Then    '
        WriteInServerKey "Denied", CStr(CurrentDateNumber)
        Call MsgBox(" ﬂœ Œÿ« 25 -  «—ÌŒ ”Ì” „ œ” ò«—Ì ‘œÂ «”  " & vbCrLf & " »« ‘—ﬂ  «› ÃÌ ¬—Ì«  „«” »êÌ—Ìœ" & vbLf & " ·›‰  „«” : 88554488-88554477-88554466-88554455", vbCritical)
        frmRegister.lblHard2.Caption = 25
        frmRegister.Show vbModal
        SetKbLayout LANG_EN_US
        End
    Else
        WriteInServerKey strTemp1, CStr(CurrentDateNumber)
    End If
    
    ' ò‰ —·  ⁄œ«œ «”‰«œ
'        Call mdifrm.FWRegistry1.GetKeyStr(flwRegLocalMachine, StrTemp5, "String Value6", strTemp3)
'        strTemp3 = mdifrm.FWEncryption1.Decode(strTemp3, 1000 + Val(strTemp))
'        If strTemp3 = "" Or Val(strTemp3) = 0 Or Val(strTemp3) > (SanadCountingRecord * AutoHavale) + 10 Or Val(strTemp3) > CountRecord + 10 Then    '
'           Call MsgBox(" ﬂœ Œÿ« 26 - «‘ò«· œ— œÌ «»Ì” - —ﬂÊ—œÂ« Õ–› ‘œÂ «‰œ" & vbCrLf & " »« ‘—ﬂ  «› ÃÌ ¬—Ì«  „«” »êÌ—Ìœ" & vbLf & " ·›‰  „«” : 88554488-88554477-88554466-88554455", vbCritical)
'           frmRegister.lblHard2.Caption = 26
'           frmRegister.Show vbModal
'           SetKbLayout LANG_EN_US
'           End
'        End If
'
'        RegRec = Val(strTemp3)  '


    ' Because se must save denied in Registry
    If CountRecord > SanadCountingRecord * AutoHavale Then
        MsgBox " ﬂœ Œÿ« 28 -  ⁄œ«œ”‰œ Â« «“ Õœ „Ã«“ »Ì‘ — «”   " & vbCrLf & " »« ‘—ﬂ  «› ÃÌ ¬—Ì«  „«” »êÌ—Ìœ" & vbLf & " ·›‰  „«” : 88554488-88554477-88554466-88554455"
        frmRegister.lblHard2.Caption = 28
        frmRegister.Show vbModal
        SetKbLayout LANG_EN_US
        End
    Else
        If CountRecord > SanadCountingRecord - 1000 Then maxRecordCountFlag = True
    End If
    'MsgBox " ‘„« œ— Õ«· «” ›«œÂ «“ ‰”ŒÂ ¬“„«Ì‘Ì ”Ì” „ Â«Ì ›—Ê‘ê«ÂÌ «› ÃÌ ¬—Ì« „Ì »«‘Ìœ" & vbCrLf & " »—«Ì  ÂÌÂ ‰”ŒÂ «’·Ì ‰—„ «›“«—»« ‘—ﬂ  «› ÃÌ ¬—Ì«  „«” »êÌ—Ìœ" & vbLf & " ·›‰  „«” : 88554488-88554477-88554466-88554455"
    If clsArya.MiladiDate = 0 Then
        ShowMessage " ‘„« œ— Õ«· «” ›«œÂ «“ ‰”ŒÂ ¬“„«Ì‘Ì ”Ì” „ Â«Ì ›—Ê‘ê«ÂÌ «› ÃÌ ¬—Ì« „Ì »«‘Ìœ" & " »—«Ì  ÂÌÂ ‰”ŒÂ «’·Ì ‰—„ «›“«—»« ‘—ﬂ  «› ÃÌ ¬—Ì« Ì« ‰„«Ì‰œê«‰ ›—Ê‘  „«” »êÌ—Ìœ ", True, False, " «∆Ìœ", ""
        ShowMessage "    ⁄œ«œ —Ê“Â«Ì ”Å—Ì ‘œÂ : " & (30 - DateRemain) & " —Ê“ " & vbLf & "  ⁄œ«œ —Ê“Â«Ì »«ﬁÌ„«‰œÂ : " & DateRemain & " —Ê“  ", True, False, " «∆Ìœ", ""
    Else
        ShowMessage " ‘„« œ— Õ«· «” ›«œÂ «“ ‰”ŒÂ ¬“„«Ì‘Ì ”Ì” „ Â«Ì ›—Ê‘ê«ÂÌ  „Ì »«‘Ìœ" & " »—«Ì  ÂÌÂ ‰”ŒÂ «’·Ì ‰—„ «›“«—»« ‰„«Ì‰œê«‰ ›—Ê‘  „«” »êÌ—Ìœ ", True, False, " «∆Ìœ", ""
    End If
     

End Sub
Public Sub AppendExpDate()
    Dim strExist As Boolean
    Dim IsFileExist As Boolean
    Dim f As New FileSystemObject
    Dim tempstring As TextStream
    Dim ExpireDate As String
    Dim i, ii As Integer
    On Error GoTo ErrorHandler
    
    IsFileExist = f.FileExists(LimitedFileName)
    If IsFileExist = False Then
        f.CreateTextFile LimitedFileName
    End If
    
    ExpireDate = clsDate.shamsi(DateAdd("d", 30, Now))    ' clsDate.shamsiAddedDate(Date, 30)
    
    For i = 1 To 50
        
        Set tempstring = f.OpenTextFile(LimitedFileName, ForWriting, False, TristateFalse)
        For ii = 1 To 3
           strTemp1 = mdifrm.FWEncryption1.Encode(Int((Rnd(1000)) * 1000000 + Rnd(1000) * 1000000000), 2000)
           tempstring.WriteLine (strTemp1)
        Next
        strTemp = mdifrm.FWEncryption1.Encode(i, 1000)
        tempstring.WriteLine (strTemp)
        strTemp1 = mdifrm.FWEncryption1.Encode(ExpireDate, i + 2000)
        tempstring.WriteLine (strTemp1)
        strTemp2 = mdifrm.FWEncryption1.Encode(CStr(CurrentDateNumber), i + 2000)
        tempstring.WriteLine (strTemp2)
        strTemp3 = mdifrm.FWEncryption1.Encode(Hhhh, i + 2000) 'Lock No
        tempstring.WriteLine (strTemp3)
                            
        tempstring.Close

        Set tempstring = f.OpenTextFile(LimitedFileName, ForReading, False, TristateFalse)

        strTemp = tempstring.ReadLine
        strTemp = tempstring.ReadLine
        strTemp = tempstring.ReadLine
        strTemp = tempstring.ReadLine
        ii = mdifrm.FWEncryption1.Decode(strTemp, 1000)
        strTemp = tempstring.ReadLine
        strTemp1 = mdifrm.FWEncryption1.Decode(strTemp, i + 2000)
        strTemp = tempstring.ReadLine
        strTemp2 = mdifrm.FWEncryption1.Decode(strTemp, i + 2000)
        strTemp = tempstring.ReadLine
        strTemp3 = mdifrm.FWEncryption1.Decode(strTemp, i + 2000)

        tempstring.Close
        If ii = i And ExpireDate = strTemp1 And CStr(CurrentDateNumber) = strTemp2 And Hhhh = strTemp3 Then
            Exit For
        End If
                
    Next
    
    Set tempstring = f.OpenTextFile(LimitedFileName, ForAppending, False, TristateFalse)
    For i = 1 To 50
       strTemp1 = mdifrm.FWEncryption1.Encode(Int((Rnd(1000)) * 1000000 + Rnd(1000) * 1000000000), 1000)
       tempstring.WriteLine (strTemp1)

    Next

    tempstring.Close

 Exit Sub
ErrorHandler:
    MsgBox "Security Error4 - " & Err.Description
    End

End Sub
Private Sub WriteInServerKey(ExpireDate As String, CurrentDateNumber As String)
    
    Dim i, ii As Integer
    For i = 1 To 50
        Set tempstring = filetemp.OpenTextFile(LimitedFileName, ForWriting, False, TristateFalse)
        For ii = 1 To 3
           strTemp = mdifrm.FWEncryption1.Encode(Int((Rnd(1000)) * 1000000 + Rnd(1000) * 1000000000), 2200)
           tempstring.WriteLine (strTemp)
        Next
        strTemp = mdifrm.FWEncryption1.Encode(i, 1000)
        tempstring.WriteLine (strTemp)
    '    strTemp1 = "Denied"   ' Access Denied
        strTemp1 = mdifrm.FWEncryption1.Encode(ExpireDate, i + 2000)
        tempstring.WriteLine (strTemp1)
    '    strTemp2 = CStr(CurrentDateNumber)   '
        strTemp2 = mdifrm.FWEncryption1.Encode(CurrentDateNumber, i + 2000)
        tempstring.WriteLine (strTemp2)
        
        strTemp3 = mdifrm.FWEncryption1.Encode(Hhhh, i + 2000)
        tempstring.WriteLine (strTemp3)

        tempstring.Close

        Set tempstring = filetemp.OpenTextFile(LimitedFileName, ForReading, False, TristateFalse)

        strTemp = tempstring.ReadLine
        strTemp = tempstring.ReadLine
        strTemp = tempstring.ReadLine
        strTemp = tempstring.ReadLine
        ii = mdifrm.FWEncryption1.Decode(strTemp, 1000)
        strTemp = tempstring.ReadLine
        strTemp1 = mdifrm.FWEncryption1.Decode(strTemp, i + 2000)
        strTemp = tempstring.ReadLine
        strTemp2 = mdifrm.FWEncryption1.Decode(strTemp, i + 2000)
        strTemp = tempstring.ReadLine
        strTemp3 = mdifrm.FWEncryption1.Decode(strTemp, i + 2000)

        tempstring.Close
        If ii = i And ExpireDate = strTemp1 And CStr(CurrentDateNumber) = strTemp2 And Hhhh = strTemp3 Then
            Exit For
        End If
    
    Next
    
    Set tempstring = filetemp.OpenTextFile(LimitedFileName, ForAppending, False, TristateFalse)
    For i = 1 To 50
       strTemp1 = mdifrm.FWEncryption1.Encode(Int((Rnd(1000)) * 1000000 + Rnd(1000) * 1000000000), 1000)
       tempstring.WriteLine (strTemp1)

    Next

    tempstring.Close

End Sub

Private Sub CheckStations()
    
    Dim ComputerNameText As String
    ComputerNameText = String(200, Chr$(0))
    GetComputerName ComputerNameText, 200
    ComputerNameText = Left$(ComputerNameText, InStr(ComputerNameText, Chr$(0)) - 1)
    MachineName = ComputerNameText
    
    If rctmp.State <> 0 Then rctmp.Close
    strtemporary = "∏≤ª≥îßX^Qõ§üπe¡¢¬íß°£ü®Rß¥™ø¥nÑßô®ö§†y∞eäo"
    strtemporary = DText((strtemporary), frmfactor.Label1.Caption)

    rctmp.Open strtemporary & clsArya.StationNo & "And ((StationType & 2 = 2 ) Or (StationType & 3 = 3)) And Branch = " & CurrentBranch, PosConnection, adOpenDynamic, adLockOptimistic, adCmdText

    If Not (rctmp.EOF = True And rctmp.BOF = True) Then
        
        Dim UpdateIsOk As Boolean
        If rctmp.Fields("StationId") = clsArya.StationNo Then
            If IsNull(rctmp.Fields("IP")) Or rctmp.Fields("IP") = "" Then
                UpdateIsOk = True
                rctmp.Fields("IP") = MachineLocalIp
            ElseIf MachineLocalIp <> rctmp.Fields("IP") Then
               ' MsgBox " Œÿ« œ— ‘‰«”«∆Ì ¬œ—” ﬂ«„ÅÌÊ —" & vbCrLf & " »« ‘—ﬂ  «› ÃÌ ¬—Ì«  „«” »êÌ—Ìœ" & vbLf & " ·›‰  „«” : 88554488-88554477-88554466-88554455"
               ' SetKbLayout LANG_EN_US
                'End
            End If
            If IsNull(rctmp.Fields("Dir")) Or rctmp.Fields("Dir") = "" Or rctmp.Fields("Dir") <> App.Path Then
                rctmp.Fields("Dir") = App.Path   'SystemFolderName
                UpdateIsOk = True
            Else
'                If (Trim(UCase(rctmp.Fields("Dir"))) <> Trim(UCase(SystemFolderName))) Then    ' MachineIP
'''''                    MsgBox " Œÿ« œ— ‘‰«”«∆Ì œ«Ì—ﬂ Ê—Ì ÊÌ‰œÊ“" & vbCrLf & " »« ‘—ﬂ   «› ÃÌ ¬—Ì«  „«” »êÌ—Ìœ"
'''''                    End
'                End If
            End If
            
            If IsNull(rctmp.Fields("Machine_Name")) Or rctmp.Fields("Machine_Name") = "" Then
                rctmp.Fields("Machine_Name") = MachineName
                UpdateIsOk = True
            Else
' «ê— ﬁ›· ‘»ﬂÂ »«‘œ‰«„ ﬂ«„ÅÌÊ — „Â„ «” 
                If LCase(rctmp.Fields("Machine_Name")) <> LCase(MachineName) And (clsArya.LimitedVersion = False Or (clsArya.LimitedVersion = True And clsArya.NetLock = True)) Then         ' MachineIP
                    MsgBox " Œÿ« œ— ‘‰«”«∆Ì ‰«„ ﬂ«„ÅÌÊ —"
                    SetKbLayout LANG_EN_US
                    End
                End If
            End If
            
            If (rctmp.Fields("StationType") And 1) = 1 Then        ' MachineIP
                Station_IsServer = True
            End If
            
            If (rctmp.Fields("StationType") And 32) = 32 Then        ' MachineIP
                Station_IsAccounting = True
            End If
            
            If rctmp.Fields("TemporaryNo") = True Then         ' MachineIP
                clsStation.TemporaryNo = True
            End If
            
            If UpdateIsOk = True Then
                rctmp.Update
                UpdateIsOk = False
            End If
        
        End If
    Else
        MsgBox "«Ì‰ «Ì” ê«Â »—«Ì «Ì‰ »—‰«„Â ¬„«œÂ ”«“Ì ‰‘œÂ «” "
        SetKbLayout LANG_EN_US
        End
    End If
    If rctmp.State = adStateOpen Then rctmp.Close

'Else    ' For Limited Version
'    Station_IsServer = True
'    Server_IP = "127.0.0.1"
'    Server_Name = MachineName
'    Server_Dir = SystemFolderName
'    FileNamePath = Server_Dir & "\Objectvar2.ini"
'End If
    i = 0
    strtemporary = "ò≤ª≥îßX^Qõ§üπe¡¢¬íß°£ü®Rá¥™ø¥nYÜ¨ï•û°û†æΩ¥nQYXTbURmlvmxní°úTz®sì¿Æ√¥nndXuüôRræ¶ª≤∂QpX"
    strtemporary = DText((strtemporary), frmfactor.Label1.Caption)
    ''''    rctmp.Open "Select * from tStations Where (StationType  &  1  = 1 ) and IsActive =1 And Branch = " & CurrentBranch , adOpenDynamic, adLockOptimistic, adCmdText
    rctmp.Open strtemporary & CurrentBranch, PosConnection, adOpenDynamic, adLockOptimistic, adCmdText
    '''' Set rctmp = RunStoredProcedure2RecordSet("Get_Pc_Stations" )
    i = 0
    If Not (rctmp.EOF = True And rctmp.BOF = True) Then
       Do While Not rctmp.EOF
          i = i + 1
           If i > 1 Then
               MsgBox " Œÿ« œ— ‘‰«”«∆Ì ”—Ê—"
               End
           End If
          Server_IP = IIf(IsNull(rctmp.Fields("IP")), "127.0.0.1", rctmp.Fields("IP"))
          Server_Name = IIf(IsNull(rctmp.Fields("Machine_Name")), ".", rctmp.Fields("Machine_Name"))
          Server_Dir = IIf(IsNull(rctmp.Fields("Dir")), App.Path, rctmp.Fields("Dir"))  'rctmp.Fields("Dir")
          
          rctmp.MoveNext
       Loop
    Else
        ' «÷«›Â ﬂ—œ‰ ”—Ê—
        Call AddServertoDB
    End If
    If rctmp.State = adStateOpen Then rctmp.Close

End Sub






