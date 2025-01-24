VERSION 5.00
Object = "{7AEDC602-D94C-11D1-BB7A-00E0290EA3C9}#1.0#0"; "ResizeKit.ocx"
Begin VB.Form FrmFinalCheck 
   ClientHeight    =   7545
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7815
   Icon            =   "FrmFinalCheck.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7545
   ScaleWidth      =   7815
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton FWBtnOK 
      BackColor       =   &H00C0C000&
      Caption         =   "À» "
      BeginProperty Font 
         Name            =   "Nazanin"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6120
      TabIndex        =   35
      Top             =   6720
      Width           =   1335
   End
   Begin VB.CommandButton FWBtnPrint 
      BackColor       =   &H00004080&
      Caption         =   "(F6)  ç«Å  "
      BeginProperty Font 
         Name            =   "Nazanin"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4320
      OLEDropMode     =   1  'Manual
      TabIndex        =   34
      Top             =   6720
      Width           =   1335
   End
   Begin VB.CommandButton FWBtnCancel 
      BackColor       =   &H00004080&
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
      Height          =   615
      Left            =   360
      TabIndex        =   33
      Top             =   6720
      Width           =   1335
   End
   Begin VB.Frame Frame2 
      Height          =   1575
      Left            =   360
      TabIndex        =   20
      Top             =   4920
      Width           =   7095
      Begin VB.CommandButton FWBtnYes 
         BackColor       =   &H00004080&
         Caption         =   "»·Ì"
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
         Left            =   2280
         TabIndex        =   32
         Top             =   960
         Width           =   1095
      End
      Begin VB.CommandButton FWBtnNo 
         BackColor       =   &H00004080&
         Caption         =   "ŒÌ—"
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
         Left            =   480
         TabIndex        =   31
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label LblTipYes 
         Alignment       =   1  'Right Justify
         Caption         =   " «∆Ìœ ê—œÌœ . "
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   465
         Left            =   4920
         RightToLeft     =   -1  'True
         TabIndex        =   28
         Top             =   840
         Visible         =   0   'False
         Width           =   1425
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   " ¬Ì« „«Ì·Ìœ „«»Â «· ›«Ê  „»·€ œ—Ì«› Ì «“ „‘ —Ì »Â Õ”«» «‰⁄«„ „‰ŸÊ— ê—œœ ø"
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   240
         RightToLeft     =   -1  'True
         TabIndex        =   21
         Top             =   240
         Width           =   6585
      End
   End
   Begin VB.Frame Frame1 
      Height          =   4335
      Left            =   360
      TabIndex        =   15
      Top             =   480
      Width           =   3615
      Begin VB.TextBox TxtPayment 
         BeginProperty Font 
            Name            =   "B Nazanin"
            Size            =   18
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   840
         TabIndex        =   17
         Top             =   600
         Width           =   1935
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "—Ì«·"
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
         Height          =   615
         Left            =   0
         TabIndex        =   27
         Top             =   2640
         Width           =   495
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "—Ì«·"
         BeginProperty Font 
            Name            =   "Nazanin"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   615
         Left            =   0
         TabIndex        =   26
         Top             =   1920
         Width           =   495
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "—Ì«·"
         BeginProperty Font 
            Name            =   "Nazanin"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   615
         Left            =   240
         TabIndex        =   25
         Top             =   600
         Width           =   495
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "ò”—"
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   555
         Left            =   2640
         RightToLeft     =   -1  'True
         TabIndex        =   24
         Top             =   2760
         Width           =   465
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "«÷«›Â"
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   585
         Left            =   2520
         RightToLeft     =   -1  'True
         TabIndex        =   23
         Top             =   2040
         Width           =   585
      End
      Begin VB.Label LblRemainMinus 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "B Nazanin"
            Size            =   18
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   585
         Left            =   600
         RightToLeft     =   -1  'True
         TabIndex        =   22
         Top             =   2640
         Width           =   1905
      End
      Begin VB.Label LblRemainPlus 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "B Nazanin"
            Size            =   18
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   585
         Left            =   600
         RightToLeft     =   -1  'True
         TabIndex        =   19
         Top             =   1920
         Width           =   1905
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "»«ﬁÌ„«‰œÂ :"
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   585
         Left            =   2160
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   1320
         Width           =   945
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         Caption         =   " œ—Ì«›   :  "
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   2160
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Top             =   240
         Width           =   945
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00E0E0E0&
      Height          =   4335
      Left            =   4080
      TabIndex        =   2
      Top             =   480
      Width           =   3375
      Begin VB.Label LblTax 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   345
         Left            =   240
         RightToLeft     =   -1  'True
         TabIndex        =   37
         Top             =   3120
         Width           =   1335
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "„«·Ì« "
         BeginProperty Font 
            Name            =   "Nazanin"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   495
         Left            =   2280
         TabIndex        =   36
         Top             =   3120
         Width           =   735
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "⁄Ê«—÷"
         BeginProperty Font 
            Name            =   "Nazanin"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   495
         Left            =   2280
         TabIndex        =   30
         Top             =   2640
         Width           =   735
      End
      Begin VB.Label LblDuty 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   345
         Left            =   240
         RightToLeft     =   -1  'True
         TabIndex        =   29
         Top             =   2760
         Width           =   1335
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Ã„⁄     "
         BeginProperty Font 
            Name            =   "Nazanin"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   615
         Left            =   2640
         TabIndex        =   14
         Top             =   240
         Width           =   495
      End
      Begin VB.Label LblSubTotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1065
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   435
         Left            =   240
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   240
         Width           =   1995
      End
      Begin VB.Label lblCarryFeeTotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   240
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   1300
         Width           =   1335
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Ã„⁄  ò·   "
         BeginProperty Font 
            Name            =   "Nazanin"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   495
         Left            =   2160
         TabIndex        =   11
         Top             =   3720
         Width           =   855
      End
      Begin VB.Label LblPacking 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "»” Â »‰œÌ"
         BeginProperty Font 
            Name            =   "Nazanin"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000C0C0&
         Height          =   375
         Left            =   1920
         TabIndex        =   10
         Top             =   1720
         Width           =   1095
      End
      Begin VB.Label LblCarryFee 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "ò—«ÌÂ Õ„·"
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
         Height          =   375
         Left            =   2280
         TabIndex        =   9
         Top             =   1280
         Width           =   855
      End
      Begin VB.Label LblDiscount 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   " Œ›Ì›"
         BeginProperty Font 
            Name            =   "Nazanin"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   375
         Left            =   2160
         TabIndex        =   8
         Top             =   840
         Width           =   855
      End
      Begin VB.Label lblSumPrice 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1065
            SubFormatType   =   1
         EndProperty
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   555
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   3720
         Width           =   1875
      End
      Begin VB.Label lblPackingTotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   240
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   1730
         Width           =   1335
      End
      Begin VB.Label lblDiscountTotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   375
         Left            =   240
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   840
         Width           =   1335
      End
      Begin VB.Label lblServiceTotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   240
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   2160
         Width           =   1335
      End
      Begin VB.Label LblService 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "”—ÊÌ”"
         BeginProperty Font 
            Name            =   "Nazanin"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   495
         Left            =   2280
         TabIndex        =   3
         Top             =   2160
         Width           =   735
      End
   End
   Begin RESIZEKITLibCtl.ResizeKit ResizeKit1 
      Height          =   480
      Left            =   240
      OleObjectBlob   =   "FrmFinalCheck.frx":A4C2
      TabIndex        =   0
      Top             =   0
      Width           =   480
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "çò ‰Â«∆Ì"
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
      Left            =   3360
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   0
      Width           =   1695
   End
End
Attribute VB_Name = "FrmFinalCheck"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Rst As New ADODB.Recordset
Dim p() As Parameter
Dim Flag As Boolean

Public intSerialNo As Double

Private Sub Form_Activate()
    lblCarryFeeTotal.Caption = frmInvoice.lblCarryFeeTotal
    lblDiscountTotal.Caption = frmInvoice.lblDiscountTotal
    lblPackingTotal.Caption = frmInvoice.lblPackingTotal
    lblServiceTotal.Caption = frmInvoice.lblServiceTotal
    LblSubTotal.Caption = frmInvoice.LblSubTotal
    LblTax.Caption = frmInvoice.lblTaxTotal
    lblDuty.Caption = frmInvoice.LblDutyTotal
'''    If clsStation.TaxView = True Then
'''       lblSumPrice.Caption = Val(frmInvoice.lblSumPrice.tag)
'''       mvarTaxAmount = 0
'''       LblTax.Caption = "0"
'''    Else
'''        mvarTaxAmount = Val(LblTax.Caption)
'''        lblSumPrice.Caption = Val(frmInvoice.lblSumPrice.tag) + Val(LblTax.Caption)
'''    End If
    lblSumPrice.Caption = frmInvoice.lblSumPrice.Tag
    Frame2.Visible = Not clsStation.AutoTip
    
    TxtPayment.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case Shift
          Case 0
              Select Case KeyCode
                  Case 13  ' Esc
                     FWBtnOK_Click
                  Case 27  ' Esc
                     FWBtnCancel_Click
                  Case 117  ' Esc
                    FWBtnPrint_Click
              End Select

    End Select
End Sub

Private Sub Form_Load()
    CenterTop Me
    
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
    If PosConnection.State = adStateOpen Then PosConnection.Close: Set PosConnection = Nothing
    SaveSetting strMainKey, Me.Name, "Left", Me.Left
    SaveSetting strMainKey, Me.Name, "Top", Me.Top

End Sub

Private Sub FWBtnCancel_Click()
''''    If flag = True Then
''''        frmMsg.fwlblMsg.Caption = "¬Ì« »—«Ì Œ—ÊÃ «“ ›—„ «ÿ„Ì‰«‰ œ«—Ìœ"
''''        frmMsg.fwBtn(0).ButtonType = flwButtonOk
''''        frmMsg.fwBtn(0).Caption = "»·Â"
''''        frmMsg.fwBtn(1).ButtonType = flwButtonCancel
''''        frmMsg.fwBtn(1).Caption = "ŒÌ—"
''''        frmMsg.Show vbModal
''''        If mvarMsgIdx <> vbYes Then Exit Sub
''''    End If
    mvarIndexNo = 0
    Unload Me
End Sub

Private Sub FWBtnNo_Click()
    If Val(LblRemainPlus.Caption) > 0 Then
        mvarTipAmount = 0
        FWBtnYes.Enabled = True
        FWBtnNo.Enabled = False
        LblTipYes.Visible = False
    End If
End Sub

Private Sub FWBtnOK_Click()
    mvarIndexNo = 1
    Unload Me

End Sub

Private Sub FWBtnPrint_Click()

    mvarIndexNo = 2
    Unload Me

End Sub

Private Sub DoCalculate()

    If Val(lblSumPrice.Caption) - Val(TxtPayment.Text) > 0 Then
        LblRemainMinus.Caption = Val(lblSumPrice.Caption) - Val(TxtPayment.Text)
        LblRemainPlus.Caption = ""
    Else
        LblRemainPlus.Caption = Val(TxtPayment.Text) - Val(lblSumPrice.Caption)
        LblRemainMinus.Caption = ""
        If clsStation.AutoTip = True Then
            mvarTipAmount = Val(LblRemainPlus.Caption)
        End If
    End If
End Sub

Private Sub FWBtnYes_Click()
    If Val(LblRemainPlus.Caption) > 0 Then
        mvarTipAmount = Val(LblRemainPlus.Caption)
        FWBtnYes.Enabled = False
        LblTipYes.Visible = True
        FWBtnNo.Enabled = True
    End If
    
End Sub


Private Sub ResizeKit1_ExitResize(ByVal XScale As Double, ByVal YScale As Double)

    If formloadFlag = True Then
        SaveSetting strMainKey, Me.Name, "Height", Me.Height
        SaveSetting strMainKey, Me.Name, "Width", Me.Width
    End If

End Sub

Private Sub TxtPayment_Change()
    DoCalculate
End Sub

Private Sub TxtPayment_KeyDown(KeyCode As Integer, Shift As Integer)
   Form_KeyDown KeyCode, Shift
End Sub
