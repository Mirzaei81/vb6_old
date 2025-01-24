VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{7AEDC602-D94C-11D1-BB7A-00E0290EA3C9}#1.0#0"; "ResizeKit.ocx"
Object = "{75D4F767-8785-11D3-93AD-0000832EF44D}#2.14#0"; "FAST2001.ocx"
Begin VB.Form frmCustComp 
   ClientHeight    =   9150
   ClientLeft      =   5055
   ClientTop       =   450
   ClientWidth     =   11835
   Icon            =   "frmCustComp.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   Picture         =   "frmCustComp.frx":A4C2
   RightToLeft     =   -1  'True
   ScaleHeight     =   9150
   ScaleMode       =   0  'User
   ScaleWidth      =   11835
   Begin VB.Frame Frame8 
      Caption         =   "‘⁄»Â"
      BeginProperty Font 
         Name            =   "B Homa"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   3960
      RightToLeft     =   -1  'True
      TabIndex        =   45
      Top             =   3360
      Width           =   3135
      Begin VB.CheckBox ChkDistance 
         Alignment       =   1  'Right Justify
         Caption         =   "Â“Ì‰Â Õ„· Ê ‰ﬁ·"
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000001&
         Height          =   615
         Left            =   600
         RightToLeft     =   -1  'True
         TabIndex        =   47
         Top             =   840
         Width           =   1815
      End
      Begin VB.ComboBox cmbBranch 
         Enabled         =   0   'False
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
         Left            =   480
         RightToLeft     =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   46
         Top             =   360
         Width           =   1965
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   495
      Left            =   3960
      RightToLeft     =   -1  'True
      ScaleHeight     =   435
      ScaleWidth      =   3075
      TabIndex        =   42
      Top             =   840
      Width           =   3135
      Begin VB.OptionButton OptionActive 
         Alignment       =   1  'Right Justify
         Caption         =   "›⁄«·"
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000001&
         Height          =   330
         Index           =   0
         Left            =   1560
         RightToLeft     =   -1  'True
         TabIndex        =   44
         Top             =   80
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.OptionButton OptionActive 
         Alignment       =   1  'Right Justify
         Caption         =   "€Ì— ›⁄«·"
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000001&
         Height          =   330
         Index           =   1
         Left            =   360
         RightToLeft     =   -1  'True
         TabIndex        =   43
         Top             =   80
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "B Homa"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4420
      Left            =   240
      RightToLeft     =   -1  'True
      TabIndex        =   25
      Top             =   650
      Width           =   3600
      Begin VB.TextBox txtCarryFee 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   240
         RightToLeft     =   -1  'True
         TabIndex        =   33
         Top             =   3300
         Width           =   2175
      End
      Begin VB.TextBox txtPaykFee 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   240
         RightToLeft     =   -1  'True
         TabIndex        =   32
         Top             =   3795
         Width           =   2175
      End
      Begin VB.TextBox txtCredit 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   240
         RightToLeft     =   -1  'True
         TabIndex        =   31
         Top             =   2790
         Width           =   2175
      End
      Begin VB.TextBox txtUnit 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   240
         RightToLeft     =   -1  'True
         TabIndex        =   30
         Top             =   2295
         Width           =   2175
      End
      Begin VB.TextBox txtFlour 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   240
         RightToLeft     =   -1  'True
         TabIndex        =   29
         Top             =   1800
         Width           =   2175
      End
      Begin VB.TextBox txtMobile 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   240
         RightToLeft     =   -1  'True
         TabIndex        =   28
         Top             =   795
         Width           =   2175
      End
      Begin VB.TextBox txtTel 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   240
         RightToLeft     =   -1  'True
         TabIndex        =   27
         Top             =   300
         Width           =   2175
      End
      Begin VB.TextBox txtintTel 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   240
         RightToLeft     =   -1  'True
         TabIndex        =   26
         Top             =   1305
         Width           =   2175
      End
      Begin VB.Label lblCarryFee 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "ﬂ—«ÌÂ Õ„·"
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000001&
         Height          =   315
         Left            =   2430
         RightToLeft     =   -1  'True
         TabIndex        =   41
         Top             =   3360
         Width           =   915
      End
      Begin VB.Label lblPaykFee 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "ﬂ—«ÌÂ ÅÌﬂ"
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000001&
         Height          =   315
         Left            =   2430
         RightToLeft     =   -1  'True
         TabIndex        =   40
         Top             =   3795
         Width           =   915
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "«⁄ »«—"
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000001&
         Height          =   315
         Left            =   2430
         RightToLeft     =   -1  'True
         TabIndex        =   39
         Top             =   2865
         Width           =   915
      End
      Begin VB.Label lblUnit 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Ê«Õœ"
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000001&
         Height          =   315
         Left            =   2430
         RightToLeft     =   -1  'True
         TabIndex        =   38
         Top             =   2355
         Width           =   915
      End
      Begin VB.Label lblPart 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "ÿ»ﬁÂ"
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000001&
         Height          =   315
         Left            =   2430
         RightToLeft     =   -1  'True
         TabIndex        =   37
         Top             =   1860
         Width           =   915
      End
      Begin VB.Label lblMobile 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "„Ê»«Ì·"
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000001&
         Height          =   315
         Left            =   2430
         RightToLeft     =   -1  'True
         TabIndex        =   36
         Top             =   855
         Width           =   915
      End
      Begin VB.Label lblTel 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   " ·›‰"
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000001&
         Height          =   315
         Left            =   2430
         RightToLeft     =   -1  'True
         TabIndex        =   35
         Top             =   360
         Width           =   915
      End
      Begin VB.Label lblintTel 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "œ«Œ·Ì"
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000001&
         Height          =   315
         Left            =   2430
         RightToLeft     =   -1  'True
         TabIndex        =   34
         Top             =   1365
         Width           =   915
      End
   End
   Begin VB.Frame Frame3 
      Height          =   4335
      Left            =   7215
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   740
      Width           =   4400
      Begin VB.TextBox txtMembershipId 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFC0C0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Nazanin"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   810
         RightToLeft     =   -1  'True
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   240
         Width           =   2175
      End
      Begin VB.ComboBox cmbGender 
         BackColor       =   &H00FFC0C0&
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
         ItemData        =   "frmCustComp.frx":A804
         Left            =   810
         List            =   "frmCustComp.frx":A806
         RightToLeft     =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   804
         Width           =   2175
      End
      Begin VB.TextBox txtName 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFC0C0&
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
         Left            =   810
         RightToLeft     =   -1  'True
         TabIndex        =   19
         Top             =   1960
         Width           =   2175
      End
      Begin VB.ComboBox cmbPrefix 
         BackColor       =   &H00FFC0C0&
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
         ItemData        =   "frmCustComp.frx":A808
         Left            =   1290
         List            =   "frmCustComp.frx":A80A
         RightToLeft     =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   1393
         Width           =   1695
      End
      Begin VB.TextBox txtFamily 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFC0C0&
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
         Left            =   75
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Top             =   2550
         Width           =   2910
      End
      Begin VB.TextBox txtDiscount 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2280
         RightToLeft     =   -1  'True
         TabIndex        =   21
         Top             =   3240
         Width           =   705
      End
      Begin VB.ComboBox cmbSellPrice 
         BeginProperty Font 
            Name            =   "Nazanin"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         ItemData        =   "frmCustComp.frx":A80C
         Left            =   600
         List            =   "frmCustComp.frx":A80E
         RightToLeft     =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   3750
         Width           =   2385
      End
      Begin VB.Label lblPrefix 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "⁄‰Ê«‰"
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000001&
         Height          =   495
         Left            =   2985
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Top             =   1413
         Width           =   1155
      End
      Begin VB.Label lblCode 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "«‘ —«ò"
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000001&
         Height          =   495
         Left            =   2985
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   195
         Width           =   1155
      End
      Begin VB.Label lblSex 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Ã‰”Ì "
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000001&
         Height          =   495
         Left            =   2985
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   804
         Width           =   1155
      End
      Begin VB.Label lblName 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "‰«„"
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000001&
         Height          =   495
         Left            =   2985
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   2022
         Width           =   1155
      End
      Begin VB.Label lblFamily 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "‰«„ Œ«‰Ê«œêÌ"
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000001&
         Height          =   495
         Left            =   2985
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   2631
         Width           =   1155
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   " Œ›Ì›"
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000001&
         Height          =   345
         Left            =   3435
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   3240
         Width           =   705
      End
      Begin VB.Label lblDiscount2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "œ—’œ"
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
         Height          =   405
         Left            =   1545
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   3255
         Width           =   705
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "‰—Œ ÅÌ‘ ›—÷"
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   405
         Left            =   3015
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   3795
         Width           =   1125
      End
   End
   Begin VB.Frame InvoiceFactor 
      Caption         =   "Õ”«»œ«—Ì"
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
      Height          =   1935
      Left            =   3960
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   1335
      Width           =   3135
      Begin VB.TextBox txtSanadNo 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   49
         Top             =   1320
         Width           =   1455
      End
      Begin VB.TextBox txtAtf 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   48
         Top             =   360
         Width           =   2175
      End
      Begin VB.TextBox txtTafsiliCode 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   24
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "‘„«—Â ”‰œ"
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
         Height          =   525
         Left            =   1920
         RightToLeft     =   -1  'True
         TabIndex        =   50
         Top             =   1320
         Width           =   1155
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblAtf 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   " ⁄ÿ›"
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
         Height          =   405
         Left            =   2280
         RightToLeft     =   -1  'True
         TabIndex        =   23
         Top             =   360
         Width           =   795
      End
      Begin VB.Label lblTafsili 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "  ›÷Ì·Ì"
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
         Height          =   405
         Left            =   2160
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   960
         Width           =   795
      End
   End
   Begin VSFlex7LCtl.VSFlexGrid vsMembers 
      Height          =   3825
      Left            =   240
      TabIndex        =   2
      Top             =   5160
      Width           =   11385
      _cx             =   20082
      _cy             =   6747
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Nazanin"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   16761024
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   11
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   500
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmCustComp.frx":A810
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   -1  'True
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   0
      ShowComboButton =   -1  'True
      WordWrap        =   0   'False
      TextStyle       =   0
      TextStyleFixed  =   0
      OleDragMode     =   0
      OleDropMode     =   0
      ComboSearch     =   3
      AutoSizeMouse   =   -1  'True
      FrozenRows      =   0
      FrozenCols      =   0
      AllowUserFreezing=   0
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
   Begin FLWCtrls.FWLabel3D fwlblMode 
      Height          =   480
      Left            =   10560
      Top             =   0
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   847
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "B Homa"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor1      =   32896
      ForeColor2      =   128
      BackColor       =   9412754
      Caption         =   "„—Ê—"
      Alignment       =   2
   End
   Begin VB.TextBox txtCode 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   600
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   2055
   End
   Begin RESIZEKITLibCtl.ResizeKit ResizeKit1 
      Height          =   480
      Left            =   120
      OleObjectBlob   =   "frmCustComp.frx":A900
      TabIndex        =   4
      Top             =   600
      Width           =   480
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "«⁄÷«¡ «‘ —«ﬂ"
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
      Left            =   4560
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   0
      Width           =   2175
   End
   Begin VB.Label lblNo 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "—œÌ›"
      BeginProperty Font 
         Name            =   "B Homa"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   315
      Left            =   2640
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   705
   End
End
Attribute VB_Name = "frmCustComp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cn As New ADODB.Connection
Dim MyFormAddEditMode As EnumAddEditMode
Dim i As Integer
Dim Parameter() As Parameter
Dim OldTafsili As Long
Dim OldAtf As Long
Dim frmact As Form

Public Sub Add()
    MyFormAddEditMode = AddMode
    
    DefaultSettings
    SetFirstToolBar
    FillvsMembers
End Sub

Public Sub Cancel()
    Add
    
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call PresetScreenSaver
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call PresetScreenSaver
End Sub

Public Sub DefaultSettings()

    txtFamily.Text = ""
    txtFlour.Text = ""
    txtintTel.Text = ""
    txtCarryFee.Text = 0
    txtPaykFee.Text = 0
   ' txtMembershipId = ""
    txtMobile.Text = ""
    TxtName.Text = ""
    txtTel.Text = ""
    txtUnit.Text = ""
    txtCredit.Text = ""
    txtDiscount.Text = ""
    cmbSellPrice.ListIndex = 0
    OptionActive(0).Value = True
    cmbGender.ListIndex = 0
    cmbPrefix.ListIndex = 0
    txtTafsiliCode.Text = ""
    ChkDistance.Value = Unchecked
    ChkDistance.Enabled = False
    If MyFormAddEditMode <> ViewMode Then
            Dim Rst As New ADODB.Recordset
            ReDim Parameter(0) As Parameter
            Parameter(0) = GenerateInputParameter("@Code", adInteger, 4, Val(txtCode.Text))
            Set Rst = RunParametricStoredProcedure2Rec("Get_tblTotal_tDistance_ByCustNo", Parameter)
            If Not (Rst.EOF = True And Rst.BOF = True) Then
                txtCarryFee.Text = Rst!carryfee
                txtPaykFee.Text = Rst!PaykFee
                ChkDistance.Enabled = True
             End If
             Rst.Close
             Set Rst = Nothing
      End If
    OldTafsili = 0
    OldAtf = 0
End Sub
Public Sub Delete()
    On Error GoTo ErrHandler
    ReDim Parameter(1) As Parameter
    Parameter(0) = GenerateInputParameter("@Code", adInteger, 4, txtMembershipId.Tag)
    Parameter(1) = GenerateInputParameter("@Branch", adInteger, 4, CurrentBranch)
    RunParametricStoredProcedure "Delete_Cust", Parameter
    FillvsMembers
    Add

ErrHandler:
    Select Case err.Number
        Case -2147217873
            
            frmMsg.fwlblMsg.Caption = "›«ò Ê—Â«ÌÌ œ— —«»ÿÂ »« «Ì‰ „‘ —Ì ÊÃÊœ œ«—œ" + vbCrLf + " ‘„« ‰„Ì  Ê«‰Ìœ «Ì‰ „‘ —Ì —« Õ–› ò‰Ìœ "
            frmMsg.fwBtn(0).Visible = False
            frmMsg.fwBtn(1).ButtonType = flwButtonOk
            frmMsg.fwBtn(1).Caption = "ﬁ»Ê·"
            frmMsg.Show vbModal
    
    End Select
    
End Sub
Public Sub Edit()
    MyFormAddEditMode = EditMode
    SetFirstToolBar

End Sub
Public Sub ExitForm()

    Unload Me
    
End Sub

Public Sub GetDataDetail()
    On Error GoTo ErrHandler
    DefaultSettings
    With vsMembers
    
        If .Rows > 1 Then
        
            If .Row > 0 Then
                ReDim Parameter(1) As Parameter
                Dim Rst As ADODB.Recordset
                
                Parameter(0) = GenerateInputParameter("@Code", adInteger, 4, Val(.TextMatrix(.Row, 1)))
                Parameter(1) = GenerateInputParameter("@Branch", adInteger, 4, CurrentBranch)
                Set Rst = RunParametricStoredProcedure2Rec("Get_Cust_info", Parameter)
                
                txtMembershipId.Tag = Rst!Code '.TextMatrix(.Row, 1)
                TxtName.Text = Rst!Name ' .TextMatrix(.Row, 2)
                txtFamily.Text = Rst!Family ' .TextMatrix(.Row, 3)
                txtTel.Text = Rst!Tel1 ' .TextMatrix(.Row, 5)
                txtMobile.Text = IIf(IsNull(Rst!Mobile), "", Rst!Mobile)   ' .TextMatrix(.Row, 7)
                txtintTel.Text = IIf(IsNull(Rst!internalNo), "", Rst!internalNo) '.TextMatrix(.Row, 8)
                txtFlour.Text = IIf(IsNull(Rst!Flour), "", Rst!Flour) ' .TextMatrix(.Row, 9)
                txtUnit.Text = IIf(IsNull(Rst!Unit), "", Rst!Unit) ' .TextMatrix(.Row, 10)
                txtDiscount.Text = Rst!Discount
                txtCredit.Text = Rst!Credit
                txtCarryFee = Rst!carryfee
                txtPaykFee = Rst!PaykFee
                txtTafsiliCode.Text = IIf(IsNull(Rst!Tafsili), "", Rst!Tafsili)
                OldTafsili = Val(txtTafsiliCode.Text)
                txtSanadNo.Text = IIf(IsNull(Rst!SanadNo), "", Rst!SanadNo)
'                If clsArya.ExternalAccounting = True And cmbAtf.ListIndex <> -1 Then
'                    OldAtf = cmbAtf.ItemData(cmbAtf.ListIndex)
'                End If
                
                For i = 0 To cmbGender.ListCount - 1
                    If cmbGender.ItemData(i) = Rst!Sex Then  ' 'Val(.TextMatrix(.Row, 6))
                        cmbGender.ListIndex = i
                        Exit For
                    End If
                Next i
                
                For i = 0 To cmbPrefix.ListCount - 1
                    If cmbPrefix.ItemData(i) = Rst!Prefix Then ' Val(.TextMatrix(.Row, 4))
                        cmbPrefix.ListIndex = i
                        Exit For
                    End If
                Next i
            
                For i = 0 To cmbSellPrice.ListCount - 1
                     If cmbSellPrice.ItemData(i) = Rst!SellPrice Then
                         cmbSellPrice.ListIndex = i
                         Exit For
                     End If
                 Next i
            
                If Rst!ActDeAct Then
                    OptionActive(0).Value = True
                Else
                    OptionActive(1).Value = True
                End If
                
            End If
            
        End If
    End With
Exit Sub

ErrHandler:
    ShowDisMessage err.Description, 2000

End Sub

Sub SetFirstToolBar()
    
    Dim Obj As Object
    
    AllButton vbOff, True

    mdifrm.Toolbar1.Buttons(1).Enabled = True   'Home
    mdifrm.Toolbar1.Buttons(2).Enabled = True   'PageUp
    mdifrm.Toolbar1.Buttons(3).Enabled = True   'PageDown
    mdifrm.Toolbar1.Buttons(4).Enabled = True   'End
    
    mdifrm.Toolbar1.Buttons(15).Enabled = True   'Print
 
    mdifrm.Toolbar1.Buttons(23).Enabled = True
    mdifrm.Toolbar1.Buttons(24).Enabled = True
    mdifrm.Toolbar1.Buttons(25).Enabled = True
    mdifrm.Toolbar1.Buttons(26).Enabled = True
    mdifrm.Toolbar1.Buttons(27).Enabled = True
 
    If MyFormAddEditMode = ViewMode Then  ' View Mode
 
        On Error Resume Next
        For Each Obj In Me
            If TypeOf Obj Is OptionButton Or TypeOf Obj Is FWCheck Or TypeOf Obj Is CheckBox Then
                Obj.Enabled = False
            ElseIf TypeOf Obj Is TextBox Or TypeOf Obj Is ComboBox Then
                Obj.Locked = True
            End If
        Next Obj
        On Error GoTo 0
        
        mdifrm.Toolbar1.Buttons(6).Enabled = True  'Add
        mdifrm.Toolbar1.Buttons(7).Enabled = True  'Edit
        mdifrm.Toolbar1.Buttons(8).Enabled = False  'Enter
        mdifrm.Toolbar1.Buttons(9).Enabled = False   'Esc
        mdifrm.Toolbar1.Buttons(10).Enabled = True   'Delete
        txtTafsiliCode.Enabled = False
        Frame1.Enabled = False
        Frame3.Enabled = False
'        fwlblMode.Caption = "„—Ê—"
        
    ElseIf MyFormAddEditMode = AddMode Then    'Add Mode
        
        On Error Resume Next
        For Each Obj In Me
            If TypeOf Obj Is OptionButton Or TypeOf Obj Is FWCheck Or TypeOf Obj Is CheckBox Then
                Obj.Enabled = True
            Else
                Obj.Locked = False
            End If
        Next Obj
        On Error GoTo 0
        
        mdifrm.Toolbar1.Buttons(6).Enabled = False  'Add
        mdifrm.Toolbar1.Buttons(7).Enabled = False  'Edit
        mdifrm.Toolbar1.Buttons(8).Enabled = True   'Enter
        mdifrm.Toolbar1.Buttons(9).Enabled = True   'Esc
        mdifrm.Toolbar1.Buttons(10).Enabled = False   'Delete
        txtTafsiliCode.Enabled = True
        Frame1.Enabled = True
        Frame3.Enabled = True
'        fwlblMode.Caption = "ÃœÌœ"
    
    ElseIf MyFormAddEditMode = EditMode Then     'Edit
    
        On Error Resume Next
        For Each Obj In Me
            If TypeOf Obj Is OptionButton Or TypeOf Obj Is FWCheck Or TypeOf Obj Is CheckBox Then
                Obj.Enabled = True
            Else
                Obj.Locked = False
            End If
        Next Obj
        On Error GoTo 0
        
        mdifrm.Toolbar1.Buttons(6).Enabled = True  'Add
        mdifrm.Toolbar1.Buttons(7).Enabled = False  'Edit
        mdifrm.Toolbar1.Buttons(8).Enabled = True   'Enter
        mdifrm.Toolbar1.Buttons(9).Enabled = True   'Esc
        mdifrm.Toolbar1.Buttons(10).Enabled = False   'Delete
        txtTafsiliCode.Enabled = True
        
        Frame1.Enabled = True
        Frame3.Enabled = True
'        fwlblMode.Caption = "«’·«Õ"
    
    End If
    HeaderLabel Val(MyFormAddEditMode), fwlblMode
End Sub

Public Sub Update()
    
    If Trim(txtMembershipId.Text) = "" Or Trim(cmbGender.Text) = "" Or Trim(txtFamily.Text) = "" Then
        frmMsg.fwlblMsg.Caption = "·ÿ›« «ÿ·«⁄«  —« »Â ÿÊ— ò«„· Ê«—œ ‰„«ÌÌœ"
        frmMsg.fwBtn(0).Visible = False
        frmMsg.fwBtn(1).ButtonType = flwButtonOk
        frmMsg.fwBtn(1).Caption = "ﬁ»Ê·"
        frmMsg.Show vbModal
        Exit Sub
    End If
    
    If Val(txtDiscount.Text) < 0 Or Val(txtDiscount.Text) > 100 Then
        frmMsg.fwlblMsg.Caption = "„ﬁœ«—  Œ›Ì› ‰„Ì  Ê«‰œ ò„ — «“ ’›— Ì« »Ì‘ — «“ ’œ œ—’œ »«‘œ "
        frmMsg.fwBtn(0).Visible = False
        frmMsg.fwBtn(1).ButtonType = flwButtonOk
        frmMsg.fwBtn(1).Caption = "ﬁ»Ê·"
        frmMsg.Show vbModal
        Exit Sub
    End If
    Select Case MyFormAddEditMode
    
        Case AddMode
'            If clsArya.ExternalAccounting = True And cmbAtf.ListIndex <> -1 Then
'                clsAccounting.CosumerCompanyAtf = cmbAtf.ItemData(cmbAtf.ListIndex)
'                setAccountingSettingFile
'            End If
        
            ReDim Parameter(42) As Parameter
            Parameter(0) = GenerateInputParameter("@MembershipId", adInteger, 4, 0)
            Parameter(1) = GenerateInputParameter("@MasterCode", adInteger, 4, txtCode.Text)
            Parameter(2) = GenerateInputParameter("@Owner", adInteger, 4, 1)
            Parameter(3) = GenerateInputParameter("@Name", adVarWChar, 50, TxtName.Text)
            Parameter(4) = GenerateInputParameter("@Family", adVarWChar, 50, txtFamily.Text)
            Parameter(5) = GenerateInputParameter("@Sex", adInteger, 4, cmbGender.ItemData(cmbGender.ListIndex))
            Parameter(6) = GenerateInputParameter("@WorkName", adVarWChar, 50, "")
            Parameter(7) = GenerateInputParameter("@InternalNo", adVarWChar, 50, txtintTel.Text)
            Parameter(8) = GenerateInputParameter("@Unit", adVarWChar, 50, txtUnit.Text)
            Parameter(9) = GenerateInputParameter("@City", adInteger, 4, 0)
            Parameter(10) = GenerateInputParameter("@ActKind", adInteger, 4, 0)
            Parameter(11) = GenerateInputParameter("@ActDeAct", adInteger, 4, IIf(OptionActive(0).Value = True, 1, 0))
            Parameter(12) = GenerateInputParameter("@Prefix", adInteger, 4, cmbPrefix.ItemData(cmbPrefix.ListIndex))
            Parameter(13) = GenerateInputParameter("@Assansor", adInteger, 4, 0)
            Parameter(14) = GenerateInputParameter("@Address", adVarWChar, 255, "")
            Parameter(15) = GenerateInputParameter("@PostalCode", adVarWChar, 50, "")
            Parameter(16) = GenerateInputParameter("@Tel1", adVarWChar, 50, Trim(txtTel.Text))
            Parameter(17) = GenerateInputParameter("@Tel2", adVarWChar, 50, "")
            Parameter(18) = GenerateInputParameter("@Tel3", adVarWChar, 50, "")
            Parameter(19) = GenerateInputParameter("@Tel4", adVarWChar, 50, "")
            Parameter(20) = GenerateInputParameter("@Mobile", adVarWChar, 50, Trim(txtMobile.Text))
            Parameter(21) = GenerateInputParameter("@Fax", adVarWChar, 50, "")
            Parameter(22) = GenerateInputParameter("@Email", adVarWChar, 50, "")
            Parameter(23) = GenerateInputParameter("@Flour", adVarWChar, 50, txtFlour.Text)
            Parameter(24) = GenerateInputParameter("@CarryFee", adDouble, 8, Val(txtCarryFee.Text))
            Parameter(25) = GenerateInputParameter("@PaykFee", adDouble, 8, Val(txtPaykFee.Text))
            Parameter(26) = GenerateInputParameter("@Distance", adInteger, 4, 0)
            Parameter(27) = GenerateInputParameter("@Credit", adDouble, 8, Val(txtCredit.Text))
            Parameter(28) = GenerateInputParameter("@Discount", adDouble, 8, Val(txtDiscount.Text))
            Parameter(29) = GenerateInputParameter("@BuyState", adInteger, 4, 15)
            Parameter(30) = GenerateInputParameter("@Description", adVarWChar, 255, "")
            Parameter(31) = GenerateInputParameter("@User", adInteger, 4, mvarCurUserNo)
            Parameter(32) = GenerateInputParameter("@FamilyNo", adInteger, 4, 0)
            Parameter(33) = GenerateInputParameter("@Member", adBoolean, 1, 0)
            Parameter(34) = GenerateInputParameter("@State", adInteger, 4, 0)
            Parameter(35) = GenerateInputParameter("@Central", adBoolean, 1, 0)
            Parameter(36) = GenerateInputParameter("@Sellprice", adSmallInt, 2, cmbSellPrice.ItemData(cmbSellPrice.ListIndex))
            Parameter(37) = GenerateInputParameter("@EconomicCode", adVarWChar, 20, Trim(frmact.TxtEconomicalCode.Text))
            Parameter(38) = GenerateInputParameter("@nvcRFID", adVarWChar, 20, Trim(frmact.TxtRfid.Text))
            Parameter(39) = GenerateInputParameter("@nvcBirthDate", adVarWChar, 10, CStr(IIf(Trim(frmact.txtBirthDate.ClipText) = "", "", Trim(frmact.txtBirthDate.Text))))
            Parameter(40) = GenerateInputParameter("@TotalRemainingAmount", adDouble, 8, 0)
            Parameter(41) = GenerateInputParameter("@Branch", adInteger, 4, cmbBranch.ItemData(cmbBranch.ListIndex))
            Parameter(42) = GenerateOutputParameter("@Code", adBigInt, 8)
            
            Dim LastCode As Long
            LastCode = RunParametricStoredProcedure("Insert_Cust", Parameter)
            If LastCode <> -1 Then
                frmMsg.fwlblMsg.Caption = "À»  „‘ —ò ÃœÌœ »« „Ê›ﬁÌ  «‰Ã«„ ‘œ"
                frmMsg.fwBtn(0).Visible = False
                frmMsg.fwBtn(1).ButtonType = flwButtonOk
                frmMsg.fwBtn(1).Caption = "ﬁ»Ê·"
                frmMsg.Show vbModal
                If clsArya.ExternalAccounting = True Or HasMiniAcc = True Then
                   Insert_Tafsili LastCode, True
                End If
                
           Else
                frmMsg.fwlblMsg.Caption = "À»  «‰Ã«„ ‰‘œ" + vbCrLf + "‘„«—Â «‘ —«ò —« »——”Ì ‰„«ÌÌœ."
                frmMsg.fwBtn(0).Visible = False
                frmMsg.fwBtn(1).ButtonType = flwButtonOk
                frmMsg.fwBtn(1).Caption = "ﬁ»Ê·"
                frmMsg.Show vbModal
'                txtMembershipId.SetFocus
                Exit Sub
            End If
            
            
        Case EditMode
        
'            If clsArya.ExternalAccounting = True And cmbAtf.ListIndex <> -1 Then
'                clsAccounting.CosumerCompanyAtf = cmbAtf.ItemData(cmbAtf.ListIndex)
'                setAccountingSettingFile
'            End If
'
            ReDim Parameter(43) As Parameter
            Parameter(0) = GenerateInputParameter("@MembershipId", adInteger, 4, 0)
            Parameter(1) = GenerateInputParameter("@MasterCode", adInteger, 4, txtCode.Text)
            Parameter(2) = GenerateInputParameter("@Owner", adInteger, 4, 1)
            Parameter(3) = GenerateInputParameter("@Name", adVarWChar, 50, TxtName.Text)
            Parameter(4) = GenerateInputParameter("@Family", adVarWChar, 50, txtFamily.Text)
            Parameter(5) = GenerateInputParameter("@Sex", adInteger, 4, cmbGender.ItemData(cmbGender.ListIndex))
            Parameter(6) = GenerateInputParameter("@WorkName", adVarWChar, 50, "")
            Parameter(7) = GenerateInputParameter("@InternalNo", adVarWChar, 50, txtintTel.Text)
            Parameter(8) = GenerateInputParameter("@Unit", adVarWChar, 50, txtUnit.Text)
            Parameter(9) = GenerateInputParameter("@City", adInteger, 4, 0)
            Parameter(10) = GenerateInputParameter("@ActKind", adInteger, 4, 0)
            Parameter(11) = GenerateInputParameter("@ActDeAct", adInteger, 4, IIf(OptionActive(0).Value = True, 1, 0))
            Parameter(12) = GenerateInputParameter("@Prefix", adInteger, 4, cmbPrefix.ItemData(cmbPrefix.ListIndex))
            Parameter(13) = GenerateInputParameter("@Assansor", adInteger, 4, 0)
            Parameter(14) = GenerateInputParameter("@Address", adVarWChar, 255, "")
            Parameter(15) = GenerateInputParameter("@PostalCode", adVarWChar, 50, "")
            Parameter(16) = GenerateInputParameter("@Tel1", adVarWChar, 50, Trim(txtTel.Text))
            Parameter(17) = GenerateInputParameter("@Tel2", adVarWChar, 50, "")
            Parameter(18) = GenerateInputParameter("@Tel3", adVarWChar, 50, "")
            Parameter(19) = GenerateInputParameter("@Tel4", adVarWChar, 50, "")
            Parameter(20) = GenerateInputParameter("@Mobile", adVarWChar, 50, Trim(txtMobile.Text))
            Parameter(21) = GenerateInputParameter("@Fax", adVarWChar, 50, "")
            Parameter(22) = GenerateInputParameter("@Email", adVarWChar, 50, "")
            Parameter(23) = GenerateInputParameter("@Flour", adVarWChar, 50, txtFlour.Text)
            Parameter(24) = GenerateInputParameter("@CarryFee", adDouble, 8, Val(txtCarryFee.Text))
            Parameter(25) = GenerateInputParameter("@PaykFee", adDouble, 8, Val(txtPaykFee.Text))
            Parameter(26) = GenerateInputParameter("@Distance", adInteger, 4, 0)
            Parameter(27) = GenerateInputParameter("@Credit", adDouble, 8, Val(txtCredit.Text))
            Parameter(28) = GenerateInputParameter("@Discount", adDouble, 8, Val(txtDiscount.Text))
            Parameter(29) = GenerateInputParameter("@BuyState", adInteger, 4, 15)
            Parameter(30) = GenerateInputParameter("@Description", adVarWChar, 255, "")
            Parameter(31) = GenerateInputParameter("@User", adInteger, 4, mvarCurUserNo)
            Parameter(32) = GenerateInputParameter("@Code", adInteger, 4, Val(txtMembershipId.Tag))
            Parameter(33) = GenerateInputParameter("@FamilyNo", adInteger, 4, 0)
            Parameter(34) = GenerateInputParameter("@Member", adBoolean, 1, 0)
            Parameter(35) = GenerateInputParameter("@State", adInteger, 4, 0)
            Parameter(36) = GenerateInputParameter("@Central", adBoolean, 1, 0)
            Parameter(37) = GenerateInputParameter("@Sellprice", adSmallInt, 2, cmbSellPrice.ItemData(cmbSellPrice.ListIndex))
            Parameter(38) = GenerateInputParameter("@EconomicCode", adVarWChar, 20, Trim(frmact.TxtEconomicalCode.Text))
            Parameter(39) = GenerateInputParameter("@nvcRFID", adVarWChar, 20, Trim(frmact.TxtRfid.Text))
            Parameter(40) = GenerateInputParameter("@nvcBirthDate", adVarWChar, 10, CStr(IIf(Trim(frmact.txtBirthDate.ClipText) = "", "", Trim(frmact.txtBirthDate.Text))))
            Parameter(41) = GenerateInputParameter("@TotalRemainingAmount", adDouble, 8, 0)
            Parameter(42) = GenerateInputParameter("@Branch", adInteger, 4, cmbBranch.ItemData(cmbBranch.ListIndex))
            Parameter(43) = GenerateOutputParameter("@Updated", adBigInt, 8)
            Dim Updated As Long
            Updated = RunParametricStoredProcedure("Update_Cust", Parameter)
            If Updated > 0 Then
                frmMsg.fwlblMsg.Caption = " €ÌÌ—«  »« „Ê›ﬁÌ  «‰Ã«„ ‘œ"
                frmMsg.fwBtn(0).Visible = False
                frmMsg.fwBtn(1).ButtonType = flwButtonOk
                frmMsg.fwBtn(1).Caption = "ﬁ»Ê·"
                frmMsg.Show vbModal
                
                If (clsArya.ExternalAccounting = True Or HasMiniAcc = True) And Val(txtCredit.Text) > 0 Then
                   Insert_Tafsili Updated, True
                End If
            Else
                frmMsg.fwlblMsg.Caption = " €ÌÌ—«  «‰Ã«„ ‰‘œ" + vbCrLf + "‘„«—Â «‘ —«ò —« »——”Ì ‰„«ÌÌœ."
                frmMsg.fwBtn(0).Visible = False
                frmMsg.fwBtn(1).ButtonType = flwButtonOk
                frmMsg.fwBtn(1).Caption = "ﬁ»Ê·"
                frmMsg.Show vbModal
                txtMembershipId.SetFocus
                Exit Sub
            End If
    
    End Select
    
    MyFormAddEditMode = AddMode
    SetFirstToolBar
    Add
    
    Exit Sub
RollBack:
    

End Sub

Private Sub ChkDistance_Click()
    If ChkDistance.Value = Checked Then
        txtCarryFee.Enabled = True
        txtPaykFee.Enabled = True
    Else
        txtCarryFee.Enabled = False
        txtPaykFee.Enabled = False
    End If
End Sub

Private Sub ResizeKit1_ExitResize(ByVal XScale As Double, ByVal YScale As Double)

    If formloadFlag = True Then
        SaveSetting strMainKey, Me.Name, "Height", Me.Height
        SaveSetting strMainKey, Me.Name, "Width", Me.Width
    End If

End Sub

Private Sub txtCarryFee_GotFocus()
    txtCarryFee.Text = ""
End Sub
Private Sub txtCarryFee_LostFocus()
    If txtCarryFee.Text = "" Then
        txtCarryFee.Text = "0"
    End If

End Sub

Private Sub FillvsMembers()
    On Error GoTo ErrHandler
    
    Dim Rst As New ADODB.Recordset
    ReDim Parameter(0) As Parameter
    
    With vsMembers
        .Rows = 1
        Parameter(0) = GenerateInputParameter("@MasterCode", adInteger, 4, Val(txtCode.Text))
        
        Set Rst = RunParametricStoredProcedure2Rec("Get_MemberCustomers", Parameter)
        
        'txtCode.Text = frmact.mvarcode
        If Not (Rst.EOF = True And Rst.BOF = True) Then
            txtMembershipId.Text = Rst!MembershipId
        End If
        i = 0
        While Rst.EOF <> True
            i = i + 1
            .Rows = .Rows + 1
            .TextMatrix(i, 0) = i
            .TextMatrix(i, 1) = Rst.Fields("Code").Value
            .TextMatrix(i, 2) = Rst.Fields("Name").Value
            .TextMatrix(i, 3) = Rst.Fields("Family").Value
            .TextMatrix(i, 4) = Rst.Fields("Prefix").Value
            .TextMatrix(i, 5) = Rst.Fields("Tel1").Value
            .TextMatrix(i, 6) = Rst.Fields("Sex").Value
            .TextMatrix(i, 7) = IIf(IsNull(Rst!Mobile), "", Rst!Mobile)
            .TextMatrix(i, 8) = IIf(IsNull(Rst!internalNo), "", Rst!internalNo)
            .TextMatrix(i, 9) = IIf(IsNull(Rst!Flour), "", Rst!Flour) '
            .TextMatrix(i, 10) = IIf(IsNull(Rst!Unit), "", Rst!Unit) '
            Rst.MoveNext
        Wend
    End With
    
    Set Rst = Nothing
Exit Sub
ErrHandler:
    MsgBox err.Description
End Sub

Private Sub Form_Activate()
    VarActForm = Me.Name
    SetFirstToolBar
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

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0

End Sub

Private Sub Form_Load()

    If ClsFormAccess.frmCustComp = False Then
        Unload Me
    End If
    
    VarActForm = Me.Name
    
    CenterTop Me
    
    
    Dim varForm As Form
    For Each varForm In Forms
        If LCase(varForm.Name) = "frmcust" Then
            Set frmact = varForm
            Exit For
        End If
    Next
    frmact.Hide
    txtCode.Text = frmact.mvarcode2
    txtMembershipId.Text = frmact.txtMembershipId.Text
    
    vsMembers.Cell(flexcpAlignment, 0, 0, 0, vsMembers.Cols - 1) = flexAlignCenterCenter
    vsMembers.Cell(flexcpAlignment, 0, 0, vsMembers.Rows - 1, 0) = flexAlignCenterCenter
    
    cmbGender.Clear
    Select Case clsStation.Language
    
        Case EnumLanguage.Farsi
        
            cmbGender.AddItem "¬ﬁ«"
            cmbGender.ItemData(cmbGender.ListCount - 1) = 1
            cmbGender.AddItem "Œ«‰„"
            cmbGender.ItemData(cmbGender.ListCount - 1) = 0
            vsMembers.ColComboList(6) = "#1;¬ﬁ«|#0;Œ«‰„"
            
        Case EnumLanguage.English
        
            cmbGender.AddItem "Male"
            cmbGender.ItemData(cmbGender.ListCount - 1) = 1
            cmbGender.AddItem "Female"
            cmbGender.ItemData(cmbGender.ListCount - 1) = 0
            vsMembers.ColComboList(6) = "#1;Male|#2;Female"
            
    End Select
    
    Dim Rst As New ADODB.Recordset
    Set Rst = RunStoredProcedure2RecordSet("Get_All_tPrefix")
    Dim tmpStr As String
    tmpStr = vsMembers.BuildComboList(Rst, "Description", "Code")
    vsMembers.ColComboList(4) = tmpStr
    cmbPrefix.Clear
    Rst.MoveFirst
    While Rst.EOF <> True
        cmbPrefix.AddItem Rst!Description
        cmbPrefix.ItemData(cmbPrefix.ListCount - 1) = Rst!Code
        Rst.MoveNext
    Wend
    If cmbPrefix.ListCount > 0 Then cmbPrefix.ListIndex = 0
     Set Rst = RunStoredProcedure2RecordSet("Get_All_tblPub_SellPrice")
    If Not (Rst.EOF = True And Rst.BOF = True) Then
        While Not Rst.EOF
            cmbSellPrice.AddItem Rst!Description
            cmbSellPrice.ItemData(cmbSellPrice.NewIndex) = Rst!Code
            Rst.MoveNext
        Wend
    Else
        cmbSellPrice.AddItem " ‰—Œ «Ê·"
        cmbSellPrice.ItemData(0) = 1
    End If
    Me.cmbSellPrice.ListIndex = 0
    Rst.Close
    
    Set Rst = Nothing
    
    With vsMembers
        .TextMatrix(0, 1) = "òœ"
        .TextMatrix(0, 2) = "‰«„"
        .TextMatrix(0, 3) = "‰«„ Œ«‰Ê«œêÌ"
        .TextMatrix(0, 4) = "⁄‰Ê«‰"
        .TextMatrix(0, 5) = " ·›‰"
        .TextMatrix(0, 6) = "Ã‰”Ì "
        .TextMatrix(0, 7) = "„Ê»«Ì·"
        .TextMatrix(0, 8) = "œ«Œ·Ì"
        .TextMatrix(0, 9) = "ÿ»ﬁÂ"
        .TextMatrix(0, 10) = "Ê«Õœ"

         For i = 0 To .Cols - 1
            .ColWidth(i) = Val(GetSetting(strMainKey, "frmCustComp_vsMembers", "Col" & i))
            If .ColWidth(i) = 0 Then
                .ColWidth(i) = 1000       '
            End If
         Next i
    End With
    
    FillBranch
    If clsArya.ExternalAccounting = True Or HasMiniAcc = True Then
        FillAtf
    Else
    End If
    
    formloadFlag = False
    Me.left = Val(GetSetting(strMainKey, Me.Name, "Left"))
    If Val(GetSetting(strMainKey, Me.Name, "Height")) > 5000 Then
        Me.Height = Val(GetSetting(strMainKey, Me.Name, "Height"))
    End If
    If Val(GetSetting(strMainKey, Me.Name, "Width")) > 5000 Then
        Me.Width = Val(GetSetting(strMainKey, Me.Name, "Width"))
    End If
    Me.top = Val(GetSetting(strMainKey, Me.Name, "Top"))
    If Me.left < 0 Then Me.left = 0
    If Me.top < 0 Then Me.top = 0
    If Me.top > Me.ScaleHeight Then Me.top = 0

    formloadFlag = True


    Add
    
End Sub
Private Sub FillBranch()
    Dim rctmp As New ADODB.Recordset
    Dim i As Long
    cmbBranch.Clear
    Set rctmp = RunStoredProcedure2RecordSet("Get_All_Branches")
    Do While rctmp.EOF = False
        cmbBranch.AddItem rctmp!nvcBranchName
        cmbBranch.ItemData(cmbBranch.NewIndex) = rctmp!Branch
        rctmp.MoveNext
    Loop
    rctmp.Close
    For i = 0 To cmbBranch.ListCount - 1
        If CurrentBranch = cmbBranch.ItemData(i) Then
            cmbBranch.ListIndex = i
            Exit For
        End If
    Next

End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    If PosConnection.State = adStateOpen Then PosConnection.Close: Set PosConnection = Nothing
    
    SaveSetting strMainKey, Me.Name, "Left", Me.left
    SaveSetting strMainKey, Me.Name, "Top", Me.top

   frmact.Show
End Sub



Private Sub vsMembers_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    If Col = -1 Then Exit Sub
    For i = 0 To vsMembers.Cols - 1
        SaveSetting strMainKey, "frmCustComp_vsMembers", "Col" & i, vsMembers.ColWidth(i)
    Next

End Sub

Private Sub vsMembers_Click()

    MyFormAddEditMode = ViewMode
    GetDataDetail
    SetFirstToolBar
    HeaderLabel Val(MyFormAddEditMode), Me.fwlblMode
   
End Sub
Private Sub Insert_Tafsili(CustCode As Long, ShowMessageflag As Boolean)
    On Error GoTo ErrHandler
    Dim rs As New ADODB.Recordset
    Dim TafsiliName As String
    TafsiliName = Trim(TxtName.Text) & " " & Trim(txtFamily.Text) & Trim(txtWorkName)
    If txtTafsiliCode.Text = "" Then
        txtTafsiliCode.Text = Accounting.Insert_TafsiliDll(ShowMessageflag, cmbBranch.ItemData(cmbBranch.ListIndex), TafsiliName, EnumAtf.Companies)
    Else
        Accounting.Update_TafsiliDll ShowMessageflag, cmbBranch.ItemData(cmbBranch.ListIndex), Val(txtTafsiliCode.Text), TafsiliName, EnumAtf.Companies
    End If
   
    If (Val(txtPrimaryBedehi) <> 0 Or Val(txtPrimaryTalab) <> 0) And Val(txtSanadNo) = 0 Then
        Accounting.Insert_PrimarySand_Cust CustCode, Val(txtTafsiliCode.Text), Val(txtPrimaryBedehi), Val(txtPrimaryTalab), 0, 0
            
    End If
    If Val(txtTafsiliCode.Text) > 0 Then
        ReDim Parameter(1) As Parameter
        Parameter(0) = GenerateInputParameter("@TafsiliId", adInteger, 4, Val(txtTafsiliCode.Text))
        Parameter(1) = GenerateInputParameter("@Code", adInteger, 4, CustCode)
        RunParametricStoredProcedure "Update_tCust_tafsili", Parameter
    End If
    If ShowMessageflag = False Then Exit Sub
    
Exit Sub
ErrHandler:
    MsgBox err.Description & "External Accountig"
    Resume Next
End Sub


Private Sub FillAtf()
'    On Error GoTo ErrHandler
'    Dim Rst As New ADODB.Recordset
'    If cn.State = 0 Then cn.Open AccstrConnectionString
'
'    Set Rst = RunStoredProcedure2RecordSet("Get_All_tblAcc_Atfs", cn)
'    cmbAtf.Clear
'    If Rst.EOF <> True And Rst.BOF <> True Then
'        Do While Rst.EOF = False
'            cmbAtf.AddItem Rst!AtfName
'            cmbAtf.ItemData(cmbAtf.NewIndex) = Rst!AtfID
'            Rst.MoveNext
'        Loop
'    End If
'    Rst.Close
'    If cmbAtf.ListCount > 0 Then
'        For i = 0 To cmbAtf.ListCount - 1
'            If cmbAtf.ItemData(i) = clsAccounting.CosumerCompanyAtf Then
'                 cmbAtf.ListIndex = i
'                 Exit For
'            End If
'        Next
'    End If
'    Exit Sub
'ErrHandler:
'MsgBox err.Description
'modgl.LogSave "frmPer", err, "FillAtf"

    txtAtf.Text = "«‘Œ«’ Ê ‘—ò Â«"

End Sub

