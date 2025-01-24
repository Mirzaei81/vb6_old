VERSION 5.00
Begin VB.Form frmTestPos 
   Caption         =   "                     ⁄„·Ì«  »«‰òÌ »« ŒÊœÅ—œ«“"
   ClientHeight    =   8250
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6360
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8250
   ScaleWidth      =   6360
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame4 
      Caption         =   "»«“ò—œ‰"
      BeginProperty Font 
         Name            =   "Nazanin"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   1095
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   22
      Top             =   7080
      Visible         =   0   'False
      Width           =   6255
      Begin VB.CommandButton CmdUnlock 
         Caption         =   "»«“ ò—œ‰"
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
         Left            =   4680
         TabIndex        =   24
         Top             =   360
         Width           =   1215
      End
      Begin VB.TextBox TxtUnlock 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   23
         Top             =   480
         Width           =   3975
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "»—ê‘  «“ Œ—Ìœ"
      BeginProperty Font 
         Name            =   "Nazanin"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   2415
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   14
      Top             =   5160
      Width           =   6255
      Begin VB.TextBox txtRefundResult 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   18
         Top             =   1920
         Width           =   3975
      End
      Begin VB.TextBox txtRefundAmount 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1440
         TabIndex        =   17
         Top             =   840
         Width           =   3015
      End
      Begin VB.CommandButton cmdRefund 
         Caption         =   "»—ê‘  «“ Õ”«»"
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
         Left            =   2040
         TabIndex        =   16
         Top             =   1320
         Width           =   1695
      End
      Begin VB.TextBox txtRefundRef 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1440
         TabIndex        =   15
         Top             =   360
         Width           =   3015
      End
      Begin VB.Label Label7 
         Caption         =   "‰ ÌÃÂ"
         BeginProperty Font 
            Name            =   "Nazanin"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   375
         Left            =   4680
         TabIndex        =   21
         Top             =   1920
         Width           =   855
      End
      Begin VB.Label Label6 
         Caption         =   "„»·€"
         BeginProperty Font 
            Name            =   "Nazanin"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   375
         Left            =   4560
         TabIndex        =   20
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "‘„«—Â ÅÌ êÌ—Ì"
         BeginProperty Font 
            Name            =   "Nazanin"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   375
         Index           =   0
         Left            =   4560
         TabIndex        =   19
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "„«‰œÂ"
      BeginProperty Font 
         Name            =   "Nazanin"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   1455
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   3720
      Width           =   6255
      Begin VB.TextBox txtBalanceResult 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   12
         Top             =   960
         Width           =   4095
      End
      Begin VB.CommandButton cmdBlanace 
         Caption         =   "„«‰œÂ êÌ—Ì"
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
         Left            =   2160
         TabIndex        =   11
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label5 
         Caption         =   "‰ ÌÃÂ"
         BeginProperty Font 
            Name            =   "Nazanin"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   375
         Left            =   4680
         TabIndex        =   13
         Top             =   960
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Œ—Ìœ"
      BeginProperty Font 
         Name            =   "Nazanin"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   3735
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   0
      Width           =   6255
      Begin VB.TextBox txtBuyResult 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   120
         TabIndex        =   5
         Top             =   2520
         Width           =   5295
      End
      Begin VB.TextBox txtMurchant 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1320
         TabIndex        =   4
         Top             =   1440
         Width           =   3015
      End
      Begin VB.TextBox txtCustomString 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1320
         TabIndex        =   3
         Top             =   960
         Width           =   3015
      End
      Begin VB.CommandButton CmdBuy 
         Caption         =   "Œ—Ìœ"
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
         Left            =   2160
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   1920
         Width           =   1215
      End
      Begin VB.TextBox txtBuyAmount 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1320
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   480
         Width           =   3015
      End
      Begin VB.Label Label3 
         Caption         =   "‰ ÌÃÂ"
         BeginProperty Font 
            Name            =   "Nazanin"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404000&
         Height          =   375
         Index           =   1
         Left            =   5520
         TabIndex        =   9
         Top             =   2520
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "‰Ê⁄ Õ”«» „‘ —Ì"
         BeginProperty Font 
            Name            =   "Nazanin"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   0
         Left            =   4320
         TabIndex        =   8
         Top             =   1440
         Width           =   1815
      End
      Begin VB.Label Label2 
         Caption         =   "—‘ Â «Œ Ì«—Ì"
         BeginProperty Font 
            Name            =   "Nazanin"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4440
         TabIndex        =   7
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "„»·€"
         BeginProperty Font 
            Name            =   "Nazanin"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   375
         Left            =   4560
         TabIndex        =   6
         Top             =   480
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmTestPos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
' Private Declare Function Buy Lib "PCPOS.dll" ( _
'                       ByVal amount As String, ByVal custom_string As String, _
'                       ByVal murchant As String) As String
'
' Private Declare Function Balance Lib "PCPOS.dll" () As String
' 'Private Declare Function UnLock Lib "PCPOS.dll" () As String
'
' Private Declare Function Refund Lib "PCPOS.dll" (ByVal RefNumber As String, ByVal amount As String) As String
'Dim WithEvents Pos1 As PCPos
'
'
'Private Sub cmdBlanace_Click()
'    txtBuyResult.Text = ""
'    If clsStation.PosModel = PasargadPos Then
'        Pos1.getBalance ""
'    Else
'        txtBalanceResult.Text = Balance()
'    End If
'End Sub
'
'Private Sub CmdBuy_Click()
'    If Val(txtBuyAmount) = 0 Then txtBuyResult = "„»·€ —« Ê«—œ ò‰Ìœ": Exit Sub
'    txtBuyResult.Text = ""
'    If clsStation.PosModel = PasargadPos Then
'        Pos1.doTransaction CCur(txtBuyAmount.Text), txtCustomString.Text
'    Else
'        txtBuyResult.Text = Buy(txtBuyAmount.Text, txtCustomString.Text, txtMurchant.Text)
'    End If
'End Sub
'Private Sub pos1_ActionCompleted(ByVal ActionRes As PEPPCPos.ActionResult)
'    'process result here
'    If ActionRes.Result = True Then
'        txtBuyResult.Text = "„»·€ " + ActionRes.TransactionRes.amount
'        txtBuyResult.Text = txtBuyResult.Text + vbNewLine + "‘„«—Â ÅÌêÌ—Ì " + ActionRes.TransactionRes.SequenceNumber
'        txtBuyResult.Text = txtBuyResult.Text + vbNewLine + "‘„«—Â ò«—  " + ActionRes.TransactionRes.CardNumber
'        txtBuyResult.Text = txtBuyResult.Text + vbNewLine + " «—ÌŒ " + ActionRes.TransactionRes.ShmasiDate
'        txtBuyResult.Text = txtBuyResult.Text + vbNewLine + "  òœœ—ŒÊ«”  " + ActionRes.TransactionRes.RequestNumber
'        txtBuyResult.Text = txtBuyResult.Text + vbNewLine + "Ê÷⁄Ì  " + ActionRes.TransactionRes.Status
'    Else
'        txtBuyResult = "⁄„·Ì«  ‘ò”  ŒÊ—œ" + vbNewLine + ActionRes.Description
'    End If
'End Sub
'
'Private Sub cmdRefund_Click()
'    txtBuyResult.Text = ""
'    txtRefundResult.Text = Refund(txtRefundRef.Text, txtRefundAmount.Text)
'End Sub
'
'Private Sub CmdUnlock_Click()
'    'TxtUnlock.Text = UnLock()
'End Sub
'
'
'Private Sub Form_Activate()
'    VarActForm = Me.Name
'End Sub
'
