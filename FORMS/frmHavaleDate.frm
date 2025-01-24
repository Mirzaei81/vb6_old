VERSION 5.00
Object = "{7AEDC602-D94C-11D1-BB7A-00E0290EA3C9}#1.0#0"; "ResizeKit.ocx"
Object = "{75D4F767-8785-11D3-93AD-0000832EF44D}#2.14#0"; "FAST2001.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmHavaleDate 
   BackColor       =   &H80000016&
   Caption         =   "                           ⁄ÌÌ‰  «—ÌŒ  Ê·Ìœ ÕÊ«·Â              "
   ClientHeight    =   2325
   ClientLeft      =   2775
   ClientTop       =   3870
   ClientWidth     =   5055
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Nazanin"
      Size            =   9.75
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   RightToLeft     =   -1  'True
   ScaleHeight     =   2325
   ScaleWidth      =   5055
   Begin RESIZEKITLibCtl.ResizeKit ResizeKit1 
      Height          =   480
      Left            =   0
      OleObjectBlob   =   "frmHavaleDate.frx":0000
      TabIndex        =   1
      Top             =   0
      Width           =   480
   End
   Begin MSMask.MaskEdBox mskDate1 
      Height          =   495
      Left            =   2400
      TabIndex        =   0
      Top             =   360
      Width           =   1275
      _ExtentX        =   2249
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
   Begin FLWCtrls.FWButton FWBtnOK 
      Height          =   495
      Left            =   360
      TabIndex        =   3
      Top             =   1200
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   873
      Caption         =   "F2 À»   "
      FontBold        =   -1  'True
      Alignment       =   1
   End
   Begin MSMask.MaskEdBox mskDate2 
      Height          =   495
      Left            =   2400
      TabIndex        =   4
      Top             =   1200
      Width           =   1275
      _ExtentX        =   2249
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
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "  «  «—ÌŒ "
      BeginProperty Font 
         Name            =   "B Homa"
         Size            =   12
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3840
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "«“  «—ÌŒ "
      BeginProperty Font 
         Name            =   "B Homa"
         Size            =   12
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3840
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   360
      Width           =   1095
   End
End
Attribute VB_Name = "frmHavaleDate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Form_Activate()
     mskDate1.SetFocus
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
 Select Case Shift
      Case 0
          Select Case KeyCode
            Case 13
              SendKeys "{Tab}", 12
                Case 113  ' F2
                           FWBtnOK_Click
                  
                 End Select
    End Select
End Sub

Private Sub Form_Load()

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
    
    On Error Resume Next
    mskDate1.Text = GetSetting(strMainKey, "HavaleDate", "Date1")
    mskDate2.Text = GetSetting(strMainKey, "HavaleDate", "Date2")
    On Error GoTo 0

End Sub
Private Sub Form_Unload(Cancel As Integer)
    
    SaveSetting strMainKey, "HavaleDate", "Date1", mskDate1.Text
    SaveSetting strMainKey, "HavaleDate", "Date2", mskDate2.Text
    
    SaveSetting strMainKey, Me.Name, "Left", Me.Left
    SaveSetting strMainKey, Me.Name, "Top", Me.Top

End Sub
Private Sub FWBtnOK_Click()
    If mskDate1.Text = "  /  /  " Or mskDate2.Text = "  /  /  " Then
        frmMsg.fwlblMsg.Caption = "·ÿ›« «ÿ·«⁄«  —« ﬂ«„· Ê«—œ ﬂ‰Ìœ"
        frmMsg.fwBtn(0).Visible = False
        frmMsg.fwBtn(1).ButtonType = flwButtonOk
        frmMsg.fwBtn(1).Caption = "ﬁ»Ê·"
        frmMsg.Show vbModal
        If mskDate1.Text = "  /  /  " Then
            mskDate1.SetFocus
        Else
            mskDate2.SetFocus
        End If
    Else
        If mvarStatus = fromStore Then
            frmPurchase.txtDescription.Text = "ÕÊ«·Â  ›—Ê‘ «“ " & frmPurchase.cmbInventory.Text & "  «—ÌŒ " & mskDate1.Text & "  « " & mskDate2.Text
        Else
            frmPurchase.txtDescription.Text = "—”Ìœ »—ê‘  ›—Ê‘ «“" & frmPurchase.cmbInventory.Text & "  «—ÌŒ " & mskDate1.Text & "  « " & mskDate2.Text
        End If
        frmPurchase.FromDate = mskDate1.Text
        frmPurchase.ToDate = mskDate2.Text
        
        Unload Me
    End If
End Sub


Private Sub ResizeKit1_ExitResize(ByVal XScale As Double, ByVal YScale As Double)

    If formloadFlag = True Then
        SaveSetting strMainKey, Me.Name, "Height", Me.Height
        SaveSetting strMainKey, Me.Name, "Width", Me.Width
    End If

End Sub
