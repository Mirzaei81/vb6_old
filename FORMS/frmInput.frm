VERSION 5.00
Object = "{75D4F767-8785-11D3-93AD-0000832EF44D}#2.14#0"; "FAST2001.ocx"
Begin VB.Form frmInput 
   BackColor       =   &H00808080&
   BorderStyle     =   0  'None
   ClientHeight    =   2205
   ClientLeft      =   3000
   ClientTop       =   10005
   ClientWidth     =   6240
   FillStyle       =   0  'Solid
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   ScaleHeight     =   2205
   ScaleWidth      =   6240
   ShowInTaskbar   =   0   'False
   Begin FLWCtrls.FWButton btnCancel 
      Cancel          =   -1  'True
      Height          =   525
      Left            =   120
      TabIndex        =   5
      Top             =   1560
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   926
      ButtonType      =   1
      Caption         =   "«‰’—«›"
      BackColor       =   10654626
      FontName        =   "B Homa"
      FontSize        =   12
      Alignment       =   1
   End
   Begin FLWCtrls.FWLabel3D fwlblInput 
      Height          =   495
      Left            =   120
      Top             =   120
      Width           =   6045
      _ExtentX        =   10663
      _ExtentY        =   873
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "B Homa"
         Size            =   12
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Deep            =   100
      ForeColor1      =   16761024
      ForeColor2      =   0
      BackColor       =   12632256
      Caption         =   ""
      Alignment       =   1
   End
   Begin FLWCtrls.FWButton btnOk 
      Default         =   -1  'True
      Height          =   525
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   926
      Caption         =   "ﬁ»Ê·"
      BackColor       =   -2147483633
      FontName        =   "B Homa"
      FontSize        =   12
      Alignment       =   1
   End
   Begin VB.TextBox txtInput 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Nazanin"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   1680
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   960
      Width           =   4485
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      Height          =   1485
      Left            =   1650
      ScaleHeight     =   1485
      ScaleWidth      =   4515
      TabIndex        =   3
      Top             =   630
      Visible         =   0   'False
      Width           =   4515
      Begin VB.OptionButton OptionLevel 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "ç«Å ›«ò Ê— ›—Ê‘"
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   390
         Index           =   2
         Left            =   0
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   960
         Visible         =   0   'False
         Width           =   4395
      End
      Begin VB.OptionButton OptionLevel 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "ç«Å „Ãœœ"
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   390
         Index           =   0
         Left            =   0
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   0
         Value           =   -1  'True
         Width           =   4395
      End
      Begin VB.OptionButton OptionLevel 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Caption         =   "ç«Å ›«ò Ê— ›—Ê‘"
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   390
         Index           =   1
         Left            =   0
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   480
         Width           =   4395
      End
   End
End
Attribute VB_Name = "frmInput"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private iHeight As Integer
Private iWidth As Integer
Private mvarIndex As Integer
Private mvarMyForm As String

Private Sub btnCancel_Click()
    mvarInput = ""
    If MyForm = "frmInvoice" Then
        Form_KeyDown 27, 0
    End If
    Unload Me
End Sub

Private Sub Form_Activate()
    'mdifrm.PicKeyBoard.Visible = False
    If txtInput.Visible = True Then
        txtInput.SelStart = 0
        txtInput.SelLength = Len(txtInput)
        On Error Resume Next
        txtInput.SetFocus
        
    Else
        If OptionLevel(0).Visible = True Then
            If OptionLevel(0).Value = True Then
                OptionLevel(0).SetFocus
            Else
                OptionLevel(1).SetFocus
            End If
        End If
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    KeyCode = 0
End Sub


Private Sub Form_Load()


    CenterCenterinSecondScreen Me

    
    mvarIndex = mvarBtnIndex
    txtInput.Text = ""
    mvarInput = ""
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

'    VarActForm = ""

    'mdifrm.PicKeyBoard.Visible = False
End Sub

Private Sub btnok_Click()

    mvarBtnIndex = mvarIndex
    mvarInput = Trim(txtInput.Text)
    If Picture1.Visible = True Then  'LCase(Me.MyForm) = "frminvoice" And
        Dim i As Integer
        For i = 0 To OptionLevel.Count - 1
            If OptionLevel(i).Value = True Then
                mvarInput = CStr(i)
            End If
        Next i
    Else  'If LCase(Me.MyForm) = "frmkeyboardtz1" Or LCase(Me.MyForm) = "frmmenu" Then
        mvarInput = Trim(txtInput.Text)
    End If
    
    Unload Me
    
End Sub


Public Property Let MyForm(mydata As String)
    mvarMyForm = mydata
End Property

Public Property Get MyForm() As String
    MyForm = mvarMyForm
End Property



