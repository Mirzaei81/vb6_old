VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmCashSharzh 
   ClientHeight    =   7275
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   4545
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
   LinkTopic       =   "Form1"
   ScaleHeight     =   7275
   ScaleWidth      =   4545
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture2 
      Height          =   1935
      Left            =   0
      ScaleHeight     =   1875
      ScaleWidth      =   4395
      TabIndex        =   10
      Top             =   4680
      Width           =   4455
      Begin VB.TextBox txtPrice2 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         IMEMode         =   3  'DISABLE
         Left            =   360
         TabIndex        =   11
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "«ê— „«Ì·Ìœ ÊÃÂ ‘«—é ’‰œÊﬁ «“ »Ì—Ê‰ Ê »Â ⁄‰Ê«‰ „«‰œÂ «Ê·ÌÂ œ— ê“«—‘ ’‰œÊﬁ ·Õ«Ÿ ‘Êœ «Ì‰ ﬁ”„  —« Å— ﬂ‰Ìœ"
         BeginProperty Font 
            Name            =   "Nazanin"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   1095
         Left            =   0
         TabIndex        =   13
         Top             =   0
         Width           =   4215
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "„»·€ (—Ì«·)"
         BeginProperty Font 
            Name            =   "Nazanin"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   435
         Left            =   2400
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   1320
         Width           =   1515
      End
   End
   Begin VB.TextBox TxtPrice 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      IMEMode         =   3  'DISABLE
      Left            =   480
      TabIndex        =   6
      Top             =   3000
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   " «∆Ìœ"
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
      Index           =   0
      Left            =   2640
      TabIndex        =   1
      Top             =   6720
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "«‰’—«›"
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
      Index           =   1
      Left            =   480
      TabIndex        =   0
      Top             =   6720
      Width           =   1095
   End
   Begin MSMask.MaskEdBox txtDate1 
      Height          =   465
      Left            =   480
      TabIndex        =   8
      Top             =   1680
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   820
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
   Begin MSMask.MaskEdBox txtDate2 
      Height          =   465
      Left            =   480
      TabIndex        =   9
      Top             =   2280
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   820
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
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "ÃÂ  ÊÌ—«Ì‘ »⁄œÌ «“ ›—„ œ—Ì«›  Ê Å—œ«Œ  «” ›«œÂ ‰„«∆Ìœ"
      BeginProperty Font 
         Name            =   "Nazanin"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   735
      Left            =   120
      TabIndex        =   7
      Top             =   3840
      Width           =   4215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "„»·€ (—Ì«·)"
      BeginProperty Font 
         Name            =   "Nazanin"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   435
      Left            =   2400
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   3000
      Width           =   1515
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      Caption         =   "¬Ì« „«Ì·Ìœ «“ „ÊÃÊœÌ ’‰œÊﬁ —Ê“ ﬁ»· »Â ’‰œÊﬁ «„—Ê“ Å—œ«Œ Ì œ«‘ Â »«‘Ìœ  « œ— ê“«—‘ ’‰œÊﬁ ·Õ«Ÿ ‘Êœø"
      BeginProperty Font 
         Name            =   "Nazanin"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   1215
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   4335
   End
   Begin VB.Shape Shape1 
      BorderWidth     =   2
      Height          =   2175
      Left            =   240
      Top             =   1560
      Width           =   4035
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "»Â ’‰œÊﬁ —Ê“"
      BeginProperty Font 
         Name            =   "Nazanin"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   435
      Left            =   2400
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   2280
      Width           =   1515
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "«“ ’‰œÊﬁ —Ê“"
      BeginProperty Font 
         Name            =   "Nazanin"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   435
      Left            =   2400
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   1680
      Width           =   1515
   End
End
Attribute VB_Name = "frmCashSharzh"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim clsDate As New clsDate

Private Sub Command1_Click(Index As Integer)
If Index = 0 Then
     Dim Result As Long
    If Val(txtPrice2) > 0 Then
        ReDim Parameter(11) As Parameter
        Parameter(0) = GenerateInputParameter("@No", adInteger, 4, 0)
        Parameter(1) = GenerateInputParameter("@List", adTinyInt, 1, 1)
        Parameter(2) = GenerateInputParameter("@Date", adVarChar, 50, Trim(txtDate2.Text))
        Parameter(3) = GenerateInputParameter("@Uid", adInteger, 4, mvarCurUserNo) '
        Parameter(4) = GenerateInputParameter("@Description", adVarChar, 300, "»«»  ‘«—é «Ê·ÌÂ ’‰œÊﬁ")
        Parameter(5) = GenerateInputParameter("@Bestankar", adInteger, 4, Val(TxtPrice.Text))
        Parameter(6) = GenerateInputParameter("@RecieveType", adInteger, 4, EnumRecieveType.CashRemain)
        Parameter(7) = GenerateInputParameter("@Code_Bes", adBigInt, 8, 0)
        Parameter(8) = GenerateInputParameter("@AddUser", adInteger, 4, mvarCurUserNo)
        Parameter(9) = GenerateInputParameter("@AccountYear", adSmallInt, 2, AccountYear)
        Parameter(10) = GenerateInputParameter("@intSerialNo", adBigInt, 8, 0)
        Parameter(11) = GenerateInputParameter("@Branch", adInteger, 4, CurrentBranch)
        
        Result = RunParametricStoredProcedure("Insert_tblAcc_Recieved", Parameter)
        ShowDisMessage "‘«—é «Ê·ÌÂ ’‰œÊﬁ »« „Ê›ﬁÌ  «‰Ã«„ ‘œ", 2000
        Unload Me
    
    Else
        If Trim(txtDate1.ClipText) = "" Or Trim(txtDate2.ClipText) = "" Or Val(TxtPrice.Text) = 0 Then
            ShowDisMessage "·ÿ›« «ÿ·«⁄«  ÷—Ê—Ì —« Ê«—œ ‰„«ÌÌœ", 2000
            Exit Sub
        End If
        ReDim Parameter(10) As Parameter
        Parameter(0) = GenerateInputParameter("@No", adInteger, 4, 0)
        Parameter(1) = GenerateInputParameter("@List", adTinyInt, 1, 1)
        Parameter(2) = GenerateInputParameter("@Date", adVarChar, 50, Trim(txtDate1.Text))
        Parameter(3) = GenerateInputParameter("@Uid", adInteger, 4, mvarCurUserNo)
        Parameter(4) = GenerateInputParameter("@Description", adVarChar, 300, "»«»  ‘«—é ’‰œÊﬁ")
        Parameter(5) = GenerateInputParameter("@Bestankar", adInteger, 4, Val(TxtPrice.Text))
        Parameter(6) = GenerateInputParameter("@PaymentType", adInteger, 4, EnumPaymentType.CashRemain)
        Parameter(7) = GenerateInputParameter("@Uid_Bede", adInteger, 4, 0)
        Parameter(8) = GenerateInputParameter("@AddUser", adInteger, 4, mvarCurUserNo)
        Parameter(9) = GenerateInputParameter("@AccountYear", adSmallInt, 2, AccountYear)
        Parameter(10) = GenerateInputParameter("@Branch", adInteger, 4, CurrentBranch)
        
        Result = RunParametricStoredProcedure("Insert_tblAcc_Cash", Parameter)

        ReDim Parameter(11) As Parameter
        Parameter(0) = GenerateInputParameter("@No", adInteger, 4, 0)
        Parameter(1) = GenerateInputParameter("@List", adTinyInt, 1, 1)
        Parameter(2) = GenerateInputParameter("@Date", adVarChar, 50, Trim(txtDate2.Text))
        Parameter(3) = GenerateInputParameter("@Uid", adInteger, 4, mvarCurUserNo) '
        Parameter(4) = GenerateInputParameter("@Description", adVarChar, 300, "»«»  ‘«—é ’‰œÊﬁ")
        Parameter(5) = GenerateInputParameter("@Bestankar", adInteger, 4, Val(TxtPrice.Text))
        Parameter(6) = GenerateInputParameter("@RecieveType", adInteger, 4, EnumRecieveType.CashRemain)
        Parameter(7) = GenerateInputParameter("@Code_Bes", adBigInt, 8, 0)
        Parameter(8) = GenerateInputParameter("@AddUser", adInteger, 4, mvarCurUserNo)
        Parameter(9) = GenerateInputParameter("@AccountYear", adSmallInt, 2, AccountYear)
        Parameter(10) = GenerateInputParameter("@intSerialNo", adBigInt, 8, 0)
        Parameter(11) = GenerateInputParameter("@Branch", adInteger, 4, CurrentBranch)
        
        Result = RunParametricStoredProcedure("Insert_tblAcc_Recieved", Parameter)
        ShowDisMessage "«‰ ﬁ«· ÊÃÂ »Â ’‰œÊﬁ »« „Ê›ﬁÌ  «‰Ã«„ ‘œ", 2000
        Unload Me
    End If
Else
    ShowDisMessage "«‰ ﬁ«· ÊÃÂ »Â ’‰œÊﬁ «‰Ã«„ ‰‘œ", 2000
    Unload Me
End If
End Sub
        
Private Sub Form_Activate()
    txtDate1.Text = Right(clsDate.shamsi(Date), 8)
    txtDate2.Text = Right(clsDate.shamsi(Date), 8)
End Sub

Private Sub TxtPrice_GotFocus()
    txtPrice2.Text = ""
End Sub

Private Sub txtPrice2_GotFocus()
    TxtPrice.Text = ""
End Sub

