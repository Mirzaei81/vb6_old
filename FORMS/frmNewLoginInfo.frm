VERSION 5.00
Begin VB.Form frmNewLoginInfo 
   BorderStyle     =   0  'None
   Caption         =   "Login"
   ClientHeight    =   1965
   ClientLeft      =   2790
   ClientTop       =   3045
   ClientWidth     =   4500
   BeginProperty Font 
      Name            =   "Nazanin"
      Size            =   12
      Charset         =   178
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   1160.987
   ScaleMode       =   0  'User
   ScaleWidth      =   4225.257
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOK 
      Caption         =   " «ÌÌœ"
      Default         =   -1  'True
      Height          =   390
      Left            =   1680
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   1440
      Width           =   1140
   End
   Begin VB.TextBox txtUserName 
      Alignment       =   2  'Center
      Height          =   480
      Left            =   439
      TabIndex        =   1
      Top             =   157
      Width           =   2325
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "«‰’—«›"
      Height          =   390
      Left            =   420
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   1440
      Width           =   1140
   End
   Begin VB.TextBox txtPassword 
      Alignment       =   2  'Center
      Height          =   480
      IMEMode         =   3  'DISABLE
      Left            =   439
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   682
      Width           =   2325
   End
   Begin VB.Label lblLabels 
      Caption         =   "‰«„ ﬂ«—»—:"
      Height          =   375
      Index           =   0
      Left            =   3199
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   157
      Width           =   1095
   End
   Begin VB.Label lblLabels 
      Caption         =   "—„“ ⁄»Ê—:"
      Height          =   375
      Index           =   1
      Left            =   3206
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   757
      Width           =   1095
   End
End
Attribute VB_Name = "frmNewLoginInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim clsDate As New clsDate

Private Sub cmdCancel_Click()
    'set the global var to false
    'to denote a failed login
    LoginSucceeded = False
    'frmRestore.cmbDataBase.Clear
    strConnectionString = ""
    Unload Me
'    Me.Hide
End Sub

Private Sub cmdOK_Click()
    Dim L_Rst As New ADODB.Recordset
    Set L_Rst = modgl.GetPerInfo(TxtUserName.Text, txtPassword.Text, CurrentBranch)
    
    If Not (L_Rst.BOF = True And L_Rst.EOF = True) Then
        If L_Rst.Fields("ActDeAct") = 0 Then
          frmMsg.fwlblMsg.Caption = "ò«—»— €Ì— ›⁄«· «”  "
          frmMsg.fwBtn(0).Visible = False
          frmMsg.fwBtn(1).ButtonType = flwButtonOk
          frmMsg.fwBtn(1).Caption = "ﬁ»Ê·"
          frmMsg.Show vbModal
          Exit Sub
        ElseIf L_Rst.Fields("Job") = 9 Then
           frmMsg.fwlblMsg.Caption = "ê«—”Ê‰ ‰„Ì  Ê«‰œ Ê«—œ ”Ì” „ ‘Êœ"
           frmMsg.fwBtn(0).Visible = False
           frmMsg.fwBtn(1).ButtonType = flwButtonOk
           frmMsg.fwBtn(1).Caption = "ﬁ»Ê·"
           frmMsg.Show vbModal
           Exit Sub
        End If
        mvarCurrentLoggedInUserName = Trim(Me.TxtUserName.Text)
        mvarCurUserNo = L_Rst.Fields("UID")
        mvarPPNo = L_Rst.Fields("pPno")
        mVarAccessLevel = L_Rst.Fields("intAccessLevel")
        mvarCountRePrint = L_Rst.Fields("CountRePrint")
        mvarCountInvoicePrint = L_Rst.Fields("CountInvoicePrint")
                 
        Dim aa, bb As String
        aa = L_Rst.Fields("Description")
        bb = L_Rst.Fields("nvcFirstName") + " " + L_Rst.Fields("nvcSurName")
    
        If clsStation.Language = Farsi Then
            mdifrm.StatusBar1.Panels(1).Text = " ‘⁄»Â :" & CurrentBranchName & " | " & "«Ã—« : ê—ÊÂ ‘—ﬂ  Â«Ì ¬—Ì« _  ·›‰ :" & "  +982188554455,+982188554466,+982188554477,+982188554488"
            mdifrm.StatusBar1.Panels(3).Text = clsDate.shamsi(Date)
            mdifrm.StatusBar1.Panels(4).Text = "”„ " & " = " & aa
            mdifrm.StatusBar1.Panels(5).Text = "‰«„ ﬂ«—»—:" & " = " & bb
'            mdifrm.StatusBar1.Panels(6).Text = mdifrm.StatusBar1.Panels(6).Text & DatabaseVersion & "_" & SoftwareVersion & "_" & CurrentScriptNo
        Else
            mdifrm.StatusBar1.Panels(1).Text = " Branch :" & CurrentBranchName & " | " & "WWW.FGARYA.COM / TEl :" & "  +982188554455,+982188554466,+982188554477,+982188554488"
            mdifrm.StatusBar1.Panels(3).Text = clsDate.shamsi(Date)
            mdifrm.StatusBar1.Panels(4).Text = "Position:" & " = " & aa
            mdifrm.StatusBar1.Panels(5).Text = "User name:" & " = " & bb
'            mdifrm.StatusBar1.Panels(6).Text = mdifrm.StatusBar1.Panels(6).Text & DatabaseVersion & "_" & SoftwareVersion & "_" & CurrentScriptNo
        End If
        LoginSucceeded = True
        'frmRestore.CmdPreSet.Enabled = True
        Unload Me
        Unload frmGroupMenu
        frmGroupMenu.Show  'Reload with new access

    Else
        ShowMessage "‰«„ ﬂ«—»—Ì Ì« —„“ ⁄»Ê— «‘ »«Â «” ", True, False, " «ÌÌœ", ""
    End If
    
    If L_Rst.State = adStateOpen Then L_Rst.Close: Set L_Rst = Nothing
    
'    'check for correct password
'    If txtPassword = "password" Then
'        'place code to here to pass the
'        'success to the calling sub
'        'setting a global var is the easiest
'        LoginSucceeded = True
'        Me.Hide
'    Else
'        MsgBox "Invalid Password, try again!", , "Login"
'        txtPassword.SetFocus
'        SendKeys "{Home}+{End}"
'    End If
End Sub

