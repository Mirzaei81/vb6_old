VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmCashClose 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "                  "
   ClientHeight    =   3960
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8310
   Icon            =   "frmCashClose.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   3960
   ScaleWidth      =   8310
   Begin VB.Frame frmCashClose 
      Height          =   3135
      Left            =   240
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   720
      Width           =   7935
      Begin VB.ComboBox cmbShift 
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
         Left            =   4680
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Text            =   "cmbShift"
         Top             =   1320
         Width           =   1815
      End
      Begin VB.CommandButton cmdCashOpen 
         BackColor       =   &H000000FF&
         Caption         =   "»«“ ﬂ—œ‰ Õ”«» Â«"
         Default         =   -1  'True
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   825
         Left            =   720
         TabIndex        =   7
         Top             =   1440
         Width           =   1755
      End
      Begin VB.CommandButton cmdCashClose 
         BackColor       =   &H00008000&
         Caption         =   "»” ‰ Õ”«» Â« œ— «Ì‰  «—ÌŒ"
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   825
         Left            =   720
         TabIndex        =   6
         Top             =   480
         Width           =   1755
      End
      Begin VB.CommandButton CmdStatus 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Ê÷⁄Ì  Õ”«»Â« "
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   825
         Left            =   2880
         TabIndex        =   2
         Top             =   1440
         Width           =   1455
      End
      Begin MSMask.MaskEdBox mdDate 
         Height          =   555
         Left            =   4680
         TabIndex        =   3
         Top             =   480
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   979
         _Version        =   393216
         MaxLength       =   8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
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
      Begin VB.Label lblStatus 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   855
         Left            =   2760
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   " ‘Ì›  :"
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6480
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "  «—ÌŒ :"
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6360
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   480
         Width           =   975
      End
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "»” ‰ Ê »«“ ﬂ—œ‰  Õ”«»Â«Ì ’‰œÊﬁ"
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
      Left            =   2040
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   0
      Width           =   4215
   End
End
Attribute VB_Name = "frmCashClose"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim clsDate As New clsDate
Dim i As Integer
Dim Parameter() As Parameter
Dim Rst As New ADODB.Recordset


Private Sub cmbShift_Change()
    CmdStatus_Click
End Sub

Private Sub cmbShift_Click()
    CmdStatus_Click
End Sub

Private Sub Form_Activate()
    
    SetFirstToolBar
    VarActForm = Me.Name
    
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case Shift
          Case 0
              Select Case KeyCode
                  Case 27  ' Esc
                     Me.ExitForm
              End Select
          Case 2
               Select Case KeyCode
                  Case 123  'Exit
                     If clsStation.KeyboardType = EnumKeyBoardType.Rb2 Then
                        Me.ExitForm
                     End If
              End Select

    End Select
End Sub

Private Sub cmdCashClose_Click()
    If Trim(mdDate.text) = "" Then
         frmDisMsg.lblMessage = "  «—ÌŒ „⁄ »— Ê«—œ ﬂ‰Ìœ "
         frmDisMsg.Timer1.Enabled = True
         frmDisMsg.Show vbModal
    Else
            ReDim Parameter(2) As Parameter
            Parameter(0) = GenerateInputParameter("@Date", adVarChar, 50, Trim(mdDate.text))
            Parameter(1) = GenerateInputParameter("@ShiftNo", adInteger, 4, cmbShift.ItemData(cmbShift.ListIndex))
            Parameter(2) = GenerateInputParameter("@CashActive", adBoolean, 1, 0)
            Set Rst = RunParametricStoredProcedure2Rec("Update_tblAcc_CashClose", Parameter)
            CmdStatus_Click
    End If
End Sub

Private Sub cmdCashOpen_Click()
    If Trim(mdDate.text) = "" Then
         frmDisMsg.lblMessage = "  «—ÌŒ „⁄ »— Ê«—œ ﬂ‰Ìœ "
         frmDisMsg.Timer1.Enabled = True
         frmDisMsg.Show vbModal
    Else
        ReDim Parameter(2) As Parameter
        Parameter(0) = GenerateInputParameter("@Date", adVarChar, 50, Trim(mdDate.text))
        Parameter(1) = GenerateInputParameter("@ShiftNo", adInteger, 4, cmbShift.ItemData(cmbShift.ListIndex))
        Parameter(2) = GenerateInputParameter("@CashActive", adBoolean, 1, 1)
        Set Rst = RunParametricStoredProcedure2Rec("Update_tblAcc_CashClose", Parameter)
        CmdStatus_Click
    End If

End Sub

Private Sub Form_Load()
    If ClsFormAccess.frmCashClose = False Then
        Unload Me
        Exit Sub
    End If
    CenterCenter Me
        
    mdDate.text = Mid(clsDate.shamsi(Date), 3)
    
    FillShift
    cmdCashClose.Enabled = False
    CmdCashOpen.Enabled = False
    CmdStatus_Click
    
End Sub

Private Sub CmdStatus_Click()
    If Trim(mdDate.text) = "" Then
         frmDisMsg.lblMessage = "  «—ÌŒ „⁄ »— Ê«—œ ﬂ‰Ìœ "
         frmDisMsg.Timer1.Enabled = True
         frmDisMsg.Show vbModal
         Exit Sub
    End If
    ReDim Parameter(1) As Parameter
    
    Parameter(0) = GenerateInputParameter("@Date", adVarChar, 50, CStr(IIf(Trim(mdDate.ClipText) = "", "", Trim(mdDate.text))))
    Parameter(1) = GenerateInputParameter("@ShiftNo", adInteger, 4, cmbShift.ItemData(cmbShift.ListIndex))
    Set Rst = RunParametricStoredProcedure2Rec("Get_tblAcc_CashClose", Parameter)

    If Rst.EOF <> True And Rst.BOF <> True Then
        If Rst!CashActive = 0 Then
            cmdCashClose.Enabled = False
            CmdCashOpen.Enabled = True
            lblStatus.ForeColor = vbRed
            lblStatus.Caption = " œ— «Ì‰  «—ÌŒ ’‰œÊﬁ  »” Â Â”  "
        Else
            cmdCashClose.Enabled = True
            CmdCashOpen.Enabled = False
            lblStatus.ForeColor = &H8000&
            lblStatus.Caption = " œ— «Ì‰  «—ÌŒ ’‰œÊﬁ »«“ Â”  "
        End If
    Else
        cmdCashClose.Enabled = True
        CmdCashOpen.Enabled = False
        lblStatus.ForeColor = &H8000&
        lblStatus.Caption = " œ— «Ì‰  «—ÌŒ ’‰œÊﬁ »«“ Â”  "
    End If

End Sub


Public Sub ExitForm()

    Unload Me
    
End Sub
Private Sub Form_Unload(Cancel As Integer)
    
    If Rst.State = 1 Then Rst.Close
    Set Rst = Nothing
    If PosConnection.State = adStateOpen Then PosConnection.Close: Set PosConnection = Nothing
       
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call PresetScreenSaver
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call PresetScreenSaver
End Sub

Private Sub SetFirstToolBar()

    AllButton vbOff, True
    
    mdifrm.Toolbar1.Buttons(23).Enabled = True
    mdifrm.Toolbar1.Buttons(24).Enabled = True
    mdifrm.Toolbar1.Buttons(25).Enabled = True
    mdifrm.Toolbar1.Buttons(26).Enabled = True
    mdifrm.Toolbar1.Buttons(27).Enabled = True

End Sub

Private Sub mdDate_Change()
    cmdCashClose.Enabled = False
    CmdCashOpen.Enabled = False
    lblStatus.Caption = ""
End Sub
Private Sub FillShift()
    
    cmbShift.Clear
'    ReDim Parameter(0) As Parameter
'    Parameter(0) = GenerateInputParameter("@Branch", adInteger, 4, CurrentBranch)
'    Set Rst = RunParametricStoredProcedure2Rec("Get_All_tShift", Parameter)
    Set Rst = RunStoredProcedure2RecordSet("Get_All_tShift")
    If Not (Rst.EOF = True And Rst.BOF = True) Then
        While Not Rst.EOF
            cmbShift.AddItem Rst!Description
            cmbShift.ItemData(cmbShift.NewIndex) = Rst!Code
            Rst.MoveNext
        Wend
        Me.cmbShift.ListIndex = 0
    Else
        cmbShift.AddItem " "
        cmbShift.ItemData(0) = -1
    End If

End Sub
