VERSION 5.00
Object = "{7AEDC602-D94C-11D1-BB7A-00E0290EA3C9}#1.0#0"; "ResizeKit.ocx"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "mscomm32.ocx"
Begin VB.Form frmSms 
   Caption         =   "                                               Arya GSM Modem "
   ClientHeight    =   9000
   ClientLeft      =   555
   ClientTop       =   945
   ClientWidth     =   9045
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSms.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9000
   ScaleWidth      =   9045
   Begin VB.Frame Frame1 
      Caption         =   " ( Step 1 ) Init && Start IR1.0 Board "
      ForeColor       =   &H00C00000&
      Height          =   1575
      Left            =   120
      TabIndex        =   33
      Top             =   120
      Width           =   2775
      Begin VB.ComboBox cmbPort 
         Height          =   315
         Left            =   1380
         Style           =   2  'Dropdown List
         TabIndex        =   38
         Top             =   240
         Width           =   1215
      End
      Begin VB.CheckBox chkCLRonStart 
         Caption         =   "Clear IR1.0 Temp on Start"
         Height          =   285
         Left            =   150
         TabIndex        =   35
         Top             =   570
         Width           =   2205
      End
      Begin VB.CheckBox chkForceUpDateTime 
         Caption         =   "Up DateTime&&Restart Before"
         Height          =   285
         Left            =   150
         TabIndex        =   34
         Top             =   810
         Width           =   2595
      End
      Begin VB.CommandButton cmdAbout 
         Caption         =   "&About"
         Height          =   390
         Left            =   180
         TabIndex        =   36
         Top             =   1110
         Visible         =   0   'False
         Width           =   885
      End
      Begin VB.CommandButton cmdInit 
         Caption         =   "&Init && Start"
         Height          =   390
         Left            =   1290
         TabIndex        =   37
         Top             =   1110
         Visible         =   0   'False
         Width           =   1275
      End
      Begin VB.Label Label1 
         Caption         =   ": Ê÷⁄Ì  «— »«ÿ"
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   1200
         TabIndex        =   43
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label lblStatus 
         Caption         =   "›⁄«· ”«“Ì"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   240
         TabIndex        =   42
         Top             =   1200
         Width           =   855
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Com Port :"
         Height          =   195
         Left            =   180
         TabIndex        =   39
         Top             =   270
         Width           =   765
      End
   End
   Begin RESIZEKITLibCtl.ResizeKit ResizeKit1 
      Height          =   480
      Left            =   120
      OleObjectBlob   =   "frmSms.frx":A4C2
      TabIndex        =   27
      Top             =   360
      Width           =   480
   End
   Begin MSCommLib.MSComm mscSerial 
      Left            =   720
      Top             =   600
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.Frame Frame5 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   1545
      Left            =   6840
      TabIndex        =   22
      Top             =   120
      Width           =   2115
      Begin VB.TextBox txtDial 
         Height          =   315
         Left            =   120
         TabIndex        =   41
         Text            =   "*141*1#"
         Top             =   1080
         Width           =   1335
      End
      Begin VB.CommandButton cmdDial 
         Caption         =   "Dial"
         Height          =   315
         Left            =   1560
         TabIndex        =   40
         Top             =   1080
         Width           =   435
      End
      Begin VB.CommandButton cmdReject 
         Caption         =   "(Hang Up) ﬁÿ⁄  „«”"
         Height          =   315
         Left            =   120
         TabIndex        =   24
         Top             =   600
         Width           =   1695
      End
      Begin VB.CheckBox chkAutoReject 
         Alignment       =   1  'Right Justify
         Caption         =   "ﬁÿ⁄  „«” « Ê„« Ìﬂ"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   240
         Width           =   1635
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         Caption         =   " „«” Â«Ì œ—Ì«› Ì"
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   240
         TabIndex        =   25
         Top             =   0
         Width           =   1455
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   " Communication's Log view "
      Enabled         =   0   'False
      ForeColor       =   &H00C00000&
      Height          =   2985
      Left            =   120
      TabIndex        =   14
      Top             =   6000
      Width           =   8775
      Begin VB.CheckBox ChkRTLRead 
         Caption         =   "(—«”  çÌ‰ (ÃÂ  ŒÊ«‰œ‰ ›«—”Ì"
         Height          =   285
         Left            =   5100
         TabIndex        =   16
         Top             =   0
         Width           =   2835
      End
      Begin VB.TextBox txtLog 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   2325
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   15
         Top             =   480
         Width           =   8295
      End
   End
   Begin VB.Frame Frame3 
      Enabled         =   0   'False
      ForeColor       =   &H00C00000&
      Height          =   4185
      Left            =   120
      TabIndex        =   6
      Top             =   1680
      Width           =   8775
      Begin VB.CommandButton cmdOk 
         Caption         =   "«—”«·Ì Â«"
         Height          =   435
         Left            =   1440
         TabIndex        =   31
         Top             =   3000
         Width           =   1095
      End
      Begin VB.CommandButton CmdFindNotice 
         Caption         =   "«‰ Œ«» „ ‰"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   360
         MaskColor       =   &H80000017&
         TabIndex        =   30
         Top             =   1200
         Width           =   1815
      End
      Begin VB.CommandButton CmdCustFind 
         Caption         =   "«‰ Œ«» „‘ —ﬂ"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   360
         MaskColor       =   &H80000017&
         TabIndex        =   29
         Top             =   360
         Width           =   1815
      End
      Begin VB.CommandButton Note 
         Caption         =   "?"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4920
         MaskColor       =   &H80000017&
         TabIndex        =   28
         Top             =   1200
         Width           =   375
      End
      Begin VB.CheckBox chkFlashSMS 
         Alignment       =   1  'Right Justify
         Caption         =   "Flash SMS «—”«· »Â ’Ê— "
         Height          =   315
         Left            =   360
         TabIndex        =   19
         Top             =   2400
         Width           =   2235
      End
      Begin VB.CheckBox chkDelivery 
         Alignment       =   1  'Right Justify
         Caption         =   " (Delivered)œ—Ì«›   «ÌÌœ «—”«· "
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   2040
         Width           =   2475
      End
      Begin VB.CheckBox ChkRTL 
         Caption         =   "(—«”  çÌ‰ (ÃÂ   «ÌÅ ›«—”Ì"
         Height          =   285
         Left            =   4920
         TabIndex        =   12
         Top             =   3840
         Width           =   2835
      End
      Begin VB.CommandButton cmdSendSMS 
         Caption         =   "«—”«·"
         Height          =   435
         Left            =   240
         TabIndex        =   11
         Top             =   3000
         Width           =   1095
      End
      Begin VB.TextBox txtSMSMessage 
         Height          =   2115
         Left            =   2760
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   9
         Top             =   1560
         Width           =   5535
      End
      Begin VB.TextBox txtSMSDest 
         Height          =   885
         Left            =   2760
         TabIndex        =   8
         Top             =   240
         Width           =   4635
      End
      Begin VB.Label LBLSMSSendStatus 
         AutoSize        =   -1  'True
         ForeColor       =   &H000000FF&
         Height          =   315
         Left            =   240
         TabIndex        =   17
         Top             =   3600
         Width           =   2355
      End
      Begin VB.Label LBLMsgLen 
         AutoSize        =   -1  'True
         Caption         =   "0 ﬂ«—«ﬂ —"
         Height          =   195
         Left            =   3240
         TabIndex        =   13
         Top             =   3840
         Width           =   570
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   ":(ÅÌ«„(»Â ’Ê—  ›«—”Ì Ê Ì« «‰ê·Ì”Ì"
         Height          =   195
         Left            =   5400
         TabIndex        =   10
         Top             =   1200
         Width           =   2520
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   ": ‘„«—Â"
         Height          =   255
         Left            =   7440
         TabIndex        =   7
         Top             =   240
         Width           =   555
      End
   End
   Begin VB.Frame Frame2 
      Enabled         =   0   'False
      ForeColor       =   &H00C00000&
      Height          =   1545
      Left            =   3000
      TabIndex        =   0
      Top             =   120
      Width           =   3735
      Begin VB.TextBox txtSMSCName 
         Height          =   315
         Left            =   120
         TabIndex        =   5
         Top             =   720
         Width           =   1785
      End
      Begin VB.CommandButton cmdSetSMSC 
         Caption         =   "–ŒÌ—Â œ— ”Ì„ ﬂ«— "
         Height          =   735
         Left            =   2880
         TabIndex        =   3
         Top             =   360
         Width           =   795
      End
      Begin VB.TextBox txtSMSCNumber 
         Height          =   315
         Left            =   120
         TabIndex        =   2
         Top             =   300
         Width           =   1425
      End
      Begin VB.Label Label11 
         Caption         =   ":  ‰ŸÌ„«  ”Ì„ ﬂ«— "
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   960
         TabIndex        =   26
         Top             =   0
         Width           =   1575
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Warn :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   90
         TabIndex        =   21
         Top             =   1170
         Width           =   540
      End
      Begin VB.Label Label9 
         Caption         =   "Change these if Requered."
         Height          =   285
         Left            =   840
         TabIndex        =   20
         Top             =   1170
         Width           =   1995
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   ": ‰«„ œ” ê«Â"
         Height          =   195
         Left            =   1680
         TabIndex        =   4
         ToolTipText     =   "SMS Center Name"
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   ": ‘„«—Â ”—ÊÌ”"
         Height          =   195
         Left            =   1440
         TabIndex        =   1
         ToolTipText     =   "SMS Center Number"
         Top             =   330
         Width           =   1290
      End
   End
   Begin VB.Label LblEdit 
      AutoSize        =   -1  'True
      ForeColor       =   &H000000FF&
      Height          =   315
      Left            =   3000
      TabIndex        =   32
      Top             =   2880
      Width           =   1755
   End
End
Attribute VB_Name = "frmSms"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''Option Explicit
''Private LastCallSEQ As Byte
''Dim rctmp As New ADODB.Recordset
''Dim i As Long
''
''Private Sub chkAutoReject_Click()
''    IR1GSM.AutoRejectInCall = chkAutoReject.Value
''    cmdReject.Enabled = Not IR1GSM.AutoRejectInCall
''End Sub
''
''
''
''Private Sub ChkRTL_Click()
''    txtSMSMessage.RightToLeft = ChkRTL.Value
''    If txtSMSMessage.RightToLeft Then
''        txtSMSMessage.Alignment = 1
''    Else
''        txtSMSMessage.Alignment = 0
''    End If
''    txtSMSMessage.SetFocus
''End Sub
''
''Private Sub ChkRTLRead_Click()
''    txtLog.RightToLeft = ChkRTLRead.Value
''    If txtLog.RightToLeft Then
''        txtLog.Alignment = 1
''    Else
''        txtLog.Alignment = 0
''    End If
''End Sub
''
''Private Sub cmdAbout_Click()
''    IR1GSM.ShowAbout
''End Sub
''Public Sub ExitForm()
''    Unload Me
''End Sub
''
''Private Sub InitModem()
''Dim s As String
''Dim i As Long
''Screen.MousePointer = 11
''Dim rctmp As New ADODB.Recordset
''
''ReDim Parameter(0) As Parameter
''Parameter(0) = GenerateInputParameter("@StationId", adInteger, 4, clsArya.StationNo)
''Set rctmp = RunParametricStoredProcedure2Rec("Get_DeviceSetting", Parameter)
''While (rctmp.EOF <> True)
''
''    If rctmp.Fields("PortNo").Value <> 0 And rctmp.Fields("PortNo").Value <> 20 Then    ' Not Lpt Port
''        If rctmp.Fields("DeviceCode").Value = EnumDevice.SmsCenter Then
''
''            If lblStatus.Caption = "›⁄«· ”«“Ì" Then
''                lblStatus.Caption = "...ç‰œ ·ÕŸÂ ’»— ﬂ‰Ìœ"
''                IR1GSM.Close_IR1
''               ' s = IR1GSM.Init_IR1(rctmp.Fields("PortNo").Value, 1)  'Old Ocx
''                cmbPort.ListIndex = rctmp.Fields("PortNo").Value - 1
''                s = IR1GSM.Init_IR1(rctmp.Fields("PortNo").Value, 0, 0)
''                If s = "Ok" Then
''                    txtSMSCName = IR1GSM.CurSMSCName
''                    txtSMSCNumber = IR1GSM.CurSMSCNumber
''                    Frame2.Enabled = True
''                    Frame3.Enabled = True
''                    Frame4.Enabled = True
''                    Frame5.Enabled = True
''                    lblStatus.Caption = "œ—Õ«· «— »«ÿ"
''                Else
''                    txtSMSCName = ""
''                    txtSMSCNumber = ""
''                    lblStatus.Caption = "⁄œ„ «— »«ÿ"
''                    MsgBox "«‘ﬂ«· œ— «— »«ÿ »« ”Œ  «›“«—- »« ‘—ﬂ  ¬—Ì«  „«” »êÌ—Ìœ" & vbCrLf & s, vbInformation, ""
''                End If
''            Else
''                IR1GSM.Close_IR1
''                Frame2.Enabled = False
''                Frame3.Enabled = False
''                Frame4.Enabled = False
''                Frame5.Enabled = False
''                txtSMSCName = ""
''                txtSMSCNumber = ""
''                lblStatus.Caption = "›⁄«· ”«“Ì"
''            End If
''
''            Screen.MousePointer = 0
''            Exit Sub
''        End If
''
''    End If
''    rctmp.MoveNext
''Wend
''Screen.MousePointer = 0
''
''End Sub
''
''Private Sub cmdDial_Click()
''    IR1GSM.Dial txtDial
''End Sub
''
''Private Sub CmdFindNotice_Click()
''    frmNotice.Show
''    frmNotice.CmdNoticForSms.Visible = True
''End Sub
''
''Private Sub cmdOK_Click()
''    frmSmsNotSend.Show vbModal
''End Sub
''
''Private Sub cmdReject_Click()
''    IR1GSM.RejectCall LastCallSEQ
''End Sub
''
''Private Sub cmdSendSMS_Click()
''        If txtSMSDest = "" Then
''            LBLSMSSendStatus.Caption = "Õœ«ﬁ· Ìﬂ ‘„«—Â »«Ìœ Ê«—œ ﬂ—œ"
''            Exit Sub
''        End If
''
''    If txtSMSDest <> "" And IR1GSM.CurSMSCNumber <> "" Then
''        LBLSMSSendStatus.Caption = "...·ÿ›« „‰ Ÿ— »„«‰Ìœ"
''
''        ReDim Parameter(2) As Parameter
''        Dim lentxtSMSDest As Integer
''        Dim h As Integer
''        lentxtSMSDest = Len(frmSms.txtSMSDest.Text) / 12
''        h = InStr(1, frmSms.txtSMSDest.Text, ";")
''
''            Parameter(0) = GenerateInputParameter("@Tel", adVarWChar, 100, frmSms.txtSMSDest.Text)
''            Parameter(1) = GenerateInputParameter("@description", adVarWChar, 400, frmSms.txtSMSMessage.Text)
''            Parameter(2) = GenerateOutputParameter("@id", adInteger, 4)
''
''           Dim Result As Integer
''           Result = RunParametricStoredProcedure("Insert_tblPubSms", Parameter)
''
''        If chkDelivery.Value = 1 Then
''            IR1GSM.DeliveryReportRequest = True
''        Else
''            IR1GSM.DeliveryReportRequest = False
''        End If
''
''        If chkFlashSMS.Value = 1 Then
''            IR1GSM.SendSMSasFlash = True
''        Else
''            IR1GSM.SendSMSasFlash = False
''        End If
''
''        If IR1GSM.SendSMS(txtSMSDest, txtSMSMessage) = 0 Then
''            LBLSMSSendStatus.Caption = "SMS «—”«· ‰‘œ!"
''        Else
''            LBLSMSSendStatus.Caption = " »« „Ê›ﬁÌ  «—”«· ‘œ SMS"
''        End If
''
''        If LBLSMSSendStatus.Caption = " »« „Ê›ﬁÌ  «—”«· ‘œ SMS" Then
''                ReDim Parameter(0) As Parameter
''
''                    Parameter(0) = GenerateInputParameter("@id", adInteger, 100, Val(Result))
''
''                    RunParametricStoredProcedure "Update_tblPubSms", Parameter
''        End If
''
''    End If
''
''End Sub
''
''Private Sub cmdSetSMSC_Click()
''    If txtSMSCName = "" Or txtSMSCNumber = "" Then Exit Sub
''    If IR1GSM.SetSMSCenter(txtSMSCNumber, txtSMSCName) = 0 Then
''        MsgBox "SMS Center Change fail", vbCritical, ""
''    Else
''        MsgBox "SMS Center Changed Sucessfully", vbInformation, ""
''    End If
''End Sub
''
''Private Sub CmdCustFind_Click()
''    frmFindCust2.Show vbModal
''End Sub
''
''Private Sub Form_Activate()
''    VarActForm = Me.Name
''End Sub
''
''Private Sub Form_Load()
''
''    Dim i As Long
''    cmbPort.Clear
''    For i = 1 To 16
''       cmbPort.AddItem "Com" & i
''    Next
''    cmbPort.ListIndex = 0
''    InitModem
''    formloadFlag = False
''    Me.Left = Val(GetSetting(strMainKey, Me.Name, "Left"))
''    If Val(GetSetting(strMainKey,Me.Name, "Height")) > 5000Then
''        Me.Height = Val(GetSetting(strMainKey, Me.Name, "Height"))
''    End If
''    If Val(GetSetting(strMainKey, Me.Name, "Width")) > 5000 Then
''        Me.Width = Val(GetSetting(strMainKey, Me.Name, "Width"))
''    End If
''    Me.Top = Val(GetSetting(strMainKey, Me.Name, "Top"))
''    formloadFlag = True
''
''End Sub
''Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
''    Call PresetScreenSaver
''End Sub
''
''Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
''    Call PresetScreenSaver
''End Sub
''
''
''Private Sub Note_Click()
''    frmMsg.fwlblMsg.Caption = " ⁄œ«œ ﬂ«—«ﬂ —Â«Ì „ ‰ ÅÌ«„ ﬂÊ «Â »Â ›«—”Ì 70 Ê »Â «‰ê·Ì”Ì 160 ﬂ«—«ﬂ — „Ì  Ê«‰œ »«‘œ"
''    frmMsg.fwBtn(0).Caption = "ﬁ»Ê·"
''    frmMsg.Show vbModal
''End Sub
''
''Private Sub ResizeKit1_ExitResize(ByVal XScale As Double, ByVal YScale As Double)
''
''    If formloadFlag = True Then
''        SaveSetting strMainKey, Me.Name, "Height", Me.Height
''        SaveSetting strMainKey, Me.Name, "Width", Me.Width
''    End If
''
''
''End Sub
''
''Private Sub Form_Unload(Cancel As Integer)
''    If PosConnection.State = adStateOpen Then PosConnection.Close: Set PosConnection = Nothing
''    IR1GSM.Close_IR1
''    Unload frmSms
''    'mdifrm.Arrange 0
''    mdifrm.Toolbar1.Buttons(20).Enabled = False
''    mdifrm.Toolbar1.Buttons(21).Enabled = False
''    mdifrm.Toolbar1.Buttons(23).Enabled = True
''    mdifrm.Toolbar1.Buttons(24).Enabled = True
''    mdifrm.Toolbar1.Buttons(25).Enabled = True
''    mdifrm.Toolbar1.Buttons(26).Enabled = True
''    mdifrm.Toolbar1.Buttons(27).Enabled = True
''
''    'mdifrm.PicKeyBoard.Visible = False
''
''    SaveSetting strMainKey, Me.Name, "Left", Me.Left
''    SaveSetting strMainKey, Me.Name, "Top", Me.Top
''
''    VarActForm = ""
''
''End Sub
''
''Private Sub IR1GSM_DeliveryReport(Dest As String, sTime As String, sDate As String)
''    AddToLog Dest & " Delivered at " & sTime & "   " & sDate
''End Sub
''
''Private Sub IR1GSM_LocalEndCall(Seq As Byte)
''    AddToLog "Local End," & " SEQ= " & Seq
''End Sub
''
''Private Sub IR1GSM_NewIncomingCall(CallerIDNumber As String, Seq As Byte)
''    LastCallSEQ = Seq
''    AddToLog "Incomming Call, CID= " & CallerIDNumber & ", SEQ= " & Seq
''End Sub
''
''Private Sub IR1GSM_RemoteEndCall(Seq As Byte)
''    AddToLog "Remote End," & " SEQ= " & Seq
''End Sub
''
''''''Private Sub IR1GSM_SMSRecive(nSMSInBuffer As Long)
''''''    Dim From As String, MSG As String
''''''    While IR1GSM.SMSRecive(From, MSG)
''''''        AddToLog "New SMS Recived From: " & From & vbCrLf & MSG
''''''        DoEvents
''''''    Wend
''''''End Sub
''
''Private Sub IR1GSM_SMSRecive(ByVal nSMSInBuffer As Long)
''    Dim From As String, Msg As String
''    While IR1GSM.SMSRecive(From, Msg)
''        AddToLog "New SMS Recived From: " & From & vbCrLf & Msg
''        DoEvents
''    Wend
''
''End Sub
''
''Private Sub txtSMSMessage_Change()
''    LBLMsgLen.Caption = Len(txtSMSMessage.Text) & " ﬂ«—«ﬂ — "
''End Sub
''
''Private Sub AddToLog(txt As String)
''    If txtLog <> "" Then
''        txtLog = txt & vbCrLf & txtLog
''    Else
''        txtLog = txt
''    End If
''End Sub
