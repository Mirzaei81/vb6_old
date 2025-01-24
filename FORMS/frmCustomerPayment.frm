VERSION 5.00
Object = "{7AEDC602-D94C-11D1-BB7A-00E0290EA3C9}#1.0#0"; "ResizeKit.ocx"
Object = "{75D4F767-8785-11D3-93AD-0000832EF44D}#2.14#0"; "FAST2001.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmCustomerPayment 
   ClientHeight    =   3645
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10425
   Icon            =   "frmCustomerPayment.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3645
   ScaleWidth      =   10425
   StartUpPosition =   3  'Windows Default
   Begin RESIZEKITLibCtl.ResizeKit ResizeKit1 
      Height          =   480
      Left            =   0
      OleObjectBlob   =   "frmCustomerPayment.frx":A4C2
      TabIndex        =   10
      Top             =   0
      Width           =   480
   End
   Begin FLWCtrls.FWButton FWBtnOK 
      Height          =   615
      Left            =   1800
      TabIndex        =   7
      Top             =   3000
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   1085
      Caption         =   "À» "
      FontName        =   "B Homa"
      FontBold        =   -1  'True
      FontSize        =   12
      Alignment       =   1
   End
   Begin FLWCtrls.FWButton FWBtnCancel 
      Height          =   615
      Left            =   240
      TabIndex        =   8
      ToolTipText     =   "«‰’—«› Ê Œ—ÊÃ"
      Top             =   3000
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   1085
      ButtonType      =   1
      Caption         =   "«‰’—«›"
      BackColor       =   12632256
      FontName        =   "B Homa"
      FontSize        =   12
      Alignment       =   1
      Object.ToolTipText     =   "«‰’—«› Ê Œ—ÊÃ"
   End
   Begin VB.Frame FrameCheque 
      Height          =   2295
      Left            =   240
      TabIndex        =   11
      Top             =   480
      Width           =   9975
      Begin VB.TextBox txtChequeDescription 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Nazanin"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3720
         TabIndex        =   6
         Top             =   1680
         Width           =   5175
      End
      Begin VB.ComboBox cmbtBank 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Nazanin"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         ItemData        =   "frmCustomerPayment.frx":A548
         Left            =   7200
         List            =   "frmCustomerPayment.frx":A54A
         RightToLeft     =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   960
         Width           =   1695
      End
      Begin VB.TextBox txtBranch 
         BeginProperty Font 
            Name            =   "Nazanin"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4320
         TabIndex        =   4
         Top             =   840
         Width           =   1455
      End
      Begin VB.TextBox txtChequeAmount 
         BeginProperty Font 
            Name            =   "Nazanin"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1560
         TabIndex        =   5
         Top             =   840
         Width           =   1395
      End
      Begin VB.TextBox txtChequeAcc 
         BeginProperty Font 
            Name            =   "Nazanin"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4320
         TabIndex        =   1
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox txtChequeSerial 
         BeginProperty Font 
            Name            =   "Nazanin"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   7200
         TabIndex        =   0
         Top             =   240
         Width           =   1695
      End
      Begin MSMask.MaskEdBox mskChequeDate 
         Height          =   495
         Left            =   1560
         TabIndex        =   2
         Top             =   240
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   873
         _Version        =   393216
         MaxLength       =   8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Nazanin"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/##"
         PromptChar      =   " "
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         Caption         =   " Ê÷ÌÕ« :"
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   525
         Left            =   9000
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   1680
         Width           =   885
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "‘⁄»Â:"
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   525
         Left            =   5880
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   960
         Width           =   645
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "”—Ì«·:"
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   525
         Left            =   9240
         RightToLeft     =   -1  'True
         TabIndex        =   17
         Top             =   240
         Width           =   645
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "‘„«—Â Õ”«»:"
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   525
         Left            =   5760
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Top             =   240
         Width           =   1245
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   " «—ÌŒ ”— —”Ìœ:"
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   525
         Left            =   3000
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   240
         Width           =   1245
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "»«‰ﬂ:"
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   525
         Left            =   9240
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   960
         Width           =   645
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "„»·€:"
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   525
         Left            =   3000
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   960
         Width           =   645
      End
   End
   Begin VB.Label lblTitel 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "œ—Ì«›  çﬂ"
      BeginProperty Font 
         Name            =   "B Homa"
         Size            =   14.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   495
      Left            =   4080
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   0
      Width           =   2775
   End
End
Attribute VB_Name = "frmCustomerPayment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private MyFormAddEditMode As EnumAddEditMode
Private CustomerPaymentType As EnumCustomerPaymentType
Dim Parameter() As Parameter
Dim CustCode As Long

Private Sub Form_Activate()
    VarActForm = Me.Name
    txtChequeSerial.SetFocus
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call PresetScreenSaver
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call PresetScreenSaver
End Sub

Private Sub Form_Load()
   CenterCenterinSecondScreen Me
   MyFormAddEditMode = ViewMode
    
    CustomerPaymentType = frmCreditCustomerAccount.CustomerPaymentType
    CustCode = frmCreditCustomerAccount.fwBtnCustFind.Tag
    DefaultSetting
    CenterCenter Me
    
    formloadFlag = False
    Me.Left = Val(GetSetting(strMainKey, Me.Name, "Left"))
    If Val(GetSetting(strMainKey, Me.Name, "Height")) > 5000 Then
        Me.Height = Val(GetSetting(strMainKey, Me.Name, "Height"))
    End If
    If Val(GetSetting(strMainKey, Me.Name, "Width")) > 5000 Then
        Me.Width = Val(GetSetting(strMainKey, Me.Name, "Width"))
    End If
    Me.Top = Val(GetSetting(strMainKey, Me.Name, "Top"))
    If Me.Left < 0 Then Me.Left = 0
    If Me.Top < 0 Then Me.Top = 0
    If Me.Top > Me.ScaleHeight Then Me.Top = 0

    formloadFlag = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
 Set Rc = Nothing
    Set rctmp = Nothing
    
    If PosConnection.State = adStateOpen Then PosConnection.Close: Set PosConnection = Nothing
        
    VarActForm = frmCreditCustomerAccount.Name
    Unload frmCustomerPayment
   
    SaveSetting strMainKey, Me.Name, "Left", Me.Left
    SaveSetting strMainKey, Me.Name, "Top", Me.Top
End Sub
Public Sub CenterCenterinSecondScreen(ByRef MyForm As Form)
     
    If MyForm.MDIChild = True Then
        MyForm.Left = mdifrm.Left + (mdifrm.Width - MyForm.Width) / 2
        MyForm.Top = (Screen.Height - MyForm.Height) / 4
    ElseIf LCase(strMainKey) = "total2" Then
        MyForm.Left = Screen.Width + (Screen.Width - MyForm.Width) / 2
        MyForm.Top = (Screen.Height - MyForm.Height) / 4
    ElseIf LCase(strMainKey) = "total" Then
        MyForm.Left = (Screen.Width - MyForm.Width) / 2
        MyForm.Top = (Screen.Height - MyForm.Height) / 4
    End If
End Sub

Private Sub FWBtnCancel_Click()
Unload Me
End Sub

Private Sub FWBtnOK_Click()
    Me.Update
End Sub

Private Sub ResizeKit1_ExitResize(ByVal XScale As Double, ByVal YScale As Double)

    If formloadFlag = True Then
        SaveSetting strMainKey, Me.Name, "Height", Me.Height
        SaveSetting strMainKey, Me.Name, "Width", Me.Width
    End If

End Sub
Public Sub DefaultSetting()
Dim cn As New ADODB.Connection
Dim rctmp As New ADODB.Recordset

    
    FrameCheque.Visible = True
    lblTitel.Caption = "œ—Ì«›  çﬂ"
    txtChequeSerial.Text = ""
    txtChequeAcc.Text = ""
    mskChequeDate.Text = "  /  /  "
    txtBranch.Text = ""
    txtChequeAmount.Text = ""
    txtChequeDescription.Text = ""
    
    Select Case clsArya.ExternalAccounting
    
        Case True
            cn.Open AccstrConnectionString
        
        Set rctmp = RunStoredProcedure2RecordSet("Get_All_tBanks", cn)
        Case False
            Set rctmp = RunStoredProcedure2RecordSet("Get_All_tBanks")
    End Select
    If Not (rctmp.EOF = True And rctmp.BOF = True) Then
        While Not rctmp.EOF
            cmbtBank.AddItem rctmp!nvcBankName
            cmbtBank.ItemData(cmbtBank.NewIndex) = rctmp!tintBank
            rctmp.MoveNext
        Wend
    Else
        cmbtBank.AddItem " "
        cmbtBank.ItemData(0) = 0
    End If
    Me.cmbtBank.ListIndex = -1
    rctmp.Close
    Set rctmp = Nothing
    If clsArya.ExternalAccounting = True Then
        If cn.State = 1 Then cn.Close: Set cn = Nothing
    End If
End Sub
Public Sub Update()

Select Case CustomerPaymentType
    Case EnumCustomerPaymentType.Cheque
        If Trim(txtChequeSerial.Text) = "" Or Trim(txtChequeAcc.Text) = "" Or mskChequeDate.Text = "  /  /  " Or Trim(txtBranch.Text) = "" Or Trim(txtChequeAmount.Text) = "" Or cmbtBank.ListIndex = -1 Then
            frmMsg.fwlblMsg.Caption = "·ÿ›« «ÿ·«⁄«  —« »’Ê—  ﬂ«„· Ê«—œ ‰„«∆Ìœ "
            frmMsg.fwBtn(1).Visible = False
            frmMsg.fwBtn(0).ButtonType = flwButtonCancel
            frmMsg.fwBtn(0).Caption = "ﬁ»Ê·"
            frmMsg.Show vbModal
            Exit Sub
        End If
        
        ReDim Parameter(0) As Parameter
                Parameter(0) = GenerateInputParameter("@intChequeSerial", adBigInt, 8, Val(txtChequeSerial.Text))
                Set rctmp = RunParametricStoredProcedure2Rec("Get_Check_tblAcc_Recieved_Cheque_Serial", Parameter)
                    If rctmp!CountSerial = 1 Then
                        frmMsg.fwlblMsg.Caption = " «Ì‰ ‘„«—Â ”—Ì«· ﬁ»·« œ— ”Ì” „ À»  ‘œÂ «” "
                        frmMsg.fwBtn(0).Visible = False
                        frmMsg.fwBtn(1).ButtonType = flwButtonOk
                        frmMsg.fwBtn(1).Caption = "ﬁ»Ê·"
                        frmMsg.Show vbModal
                        txtChequeSerial.SetFocus
                        Exit Sub
                    End If
        ReDim Parameter(9) As Parameter
        Parameter(0) = GenerateInputParameter("@intChequeSerial", adBigInt, 8, Val(txtChequeSerial.Text))
        Parameter(1) = GenerateInputParameter("@intChequeAcc", adBigInt, 8, Val(txtChequeAcc.Text))
        Parameter(2) = GenerateInputParameter("@ChequeDate", adVarChar, 50, mskChequeDate.Text)
        Parameter(3) = GenerateInputParameter("@tintBank", adTinyInt, 1, cmbtBank.ItemData(cmbtBank.ListIndex))
        Parameter(4) = GenerateInputParameter("@nvcBranch", adVarWChar, 50, txtBranch.Text)
        Parameter(5) = GenerateInputParameter("@intChequeAmount", adInteger, 4, Val(txtChequeAmount.Text))
        Parameter(6) = GenerateInputParameter("@UID", adInteger, 4, mvarCurUserNo)
        Parameter(7) = GenerateInputParameter("@Description", adVarWChar, 50, txtChequeDescription.Text)
        Parameter(8) = GenerateInputParameter("@Code_Bes", adBigInt, 8, CustCode)
        Parameter(9) = GenerateInputParameter("@AccountYear", adSmallInt, 2, AccountYear)
        RunParametricStoredProcedure "Insert_tblAcc_Recieved_Cheque", Parameter
            
  
  
  End Select
Unload Me
End Sub
Private Sub txtChequeAmount_KeyPress(KeyAscii As Integer)
 If IsNumeric(Chr(KeyAscii)) = False And (KeyAscii <> 8 And KeyAscii <> 13) Then
    KeyAscii = 0
End If
End Sub

