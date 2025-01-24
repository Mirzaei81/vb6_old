VERSION 5.00
Object = "{7AEDC602-D94C-11D1-BB7A-00E0290EA3C9}#1.0#0"; "ResizeKit.ocx"
Object = "{75D4F767-8785-11D3-93AD-0000832EF44D}#2.14#0"; "FAST2001.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmTarazSoodZian 
   BorderStyle     =   5  'Sizable ToolWindow
   Caption         =   "”Êœ Ê “Ì«‰"
   ClientHeight    =   9420
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13650
   ForeColor       =   &H00000080&
   Icon            =   "frmTaraz.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9420
   ScaleWidth      =   13650
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame4 
      Caption         =   "›—Ê‘ Ê  Œ›Ì›«  ﬂ«·«"
      BeginProperty Font 
         Name            =   "Nazanin"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6120
      Left            =   6720
      RightToLeft     =   -1  'True
      TabIndex        =   46
      Top             =   2400
      Width           =   6855
      Begin VB.Frame Frame2 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1005
         Index           =   1
         Left            =   120
         TabIndex        =   47
         Top             =   4680
         Width           =   6615
         Begin VB.Label LblSoodZianNavizhe 
            Alignment       =   1  'Right Justify
            Caption         =   "”Êœ (“Ì«‰ ) ‰«ÊÌéÂ"
            BeginProperty Font 
               Name            =   "Nazanin"
               Size            =   14.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   495
            Left            =   3600
            RightToLeft     =   -1  'True
            TabIndex        =   49
            Top             =   315
            Width           =   2175
         End
         Begin VB.Label LblTotalBenefitLoss 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "B Nazanin"
               Size            =   14.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   495
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   48
            Top             =   270
            Width           =   2595
         End
      End
      Begin VB.Label LblTotalFinalSellReturnAmount 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "B Nazanin"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   165
         RightToLeft     =   -1  'True
         TabIndex        =   59
         Top             =   4155
         Width           =   2025
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         Caption         =   "«÷«›Â „Ì ‘Êœ ﬁÌ„  »—ê‘  «“ ›—Ê‘"
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
         Left            =   2610
         RightToLeft     =   -1  'True
         TabIndex        =   58
         Top             =   4140
         Width           =   3735
      End
      Begin VB.Label LblTotalSellReturnAmount 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "B Nazanin"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   480
         Left            =   135
         RightToLeft     =   -1  'True
         TabIndex        =   57
         Top             =   3210
         Width           =   2025
      End
      Begin VB.Label LblTotalSellReturnAmountLabel 
         Alignment       =   1  'Right Justify
         Caption         =   "ﬂ”— „Ì ‘Êœ ﬁÌ„   „«„ ‘œÂ ›—Ê‘ ò«·«Â«"
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
         Height          =   495
         Left            =   2520
         RightToLeft     =   -1  'True
         TabIndex        =   56
         Top             =   2310
         Width           =   3855
      End
      Begin VB.Label LblTotalSellAmount 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "B Nazanin"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   135
         RightToLeft     =   -1  'True
         TabIndex        =   55
         Top             =   420
         Width           =   2025
      End
      Begin VB.Label LblTotalSellAmountLabel 
         Alignment       =   1  'Right Justify
         Caption         =   "›—Ê‘ ﬂ· "
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
         Left            =   3690
         RightToLeft     =   -1  'True
         TabIndex        =   54
         Top             =   555
         Width           =   2655
      End
      Begin VB.Label LblTotalSaleDiscount 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "B Nazanin"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   480
         Left            =   135
         RightToLeft     =   -1  'True
         TabIndex        =   53
         Top             =   1305
         Width           =   2025
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "ﬂ”— „Ì ‘Êœ  Œ›Ì›«  ›—Ê‘"
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
         Height          =   495
         Left            =   3360
         RightToLeft     =   -1  'True
         TabIndex        =   52
         Top             =   1395
         Width           =   3015
      End
      Begin VB.Label lblTotalFinalSellAmount 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "B Nazanin"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   480
         Left            =   135
         RightToLeft     =   -1  'True
         TabIndex        =   51
         Top             =   2250
         Width           =   2025
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         Caption         =   "ﬂ”— „Ì ‘Êœ »—ê‘  «“ ›—Ê‘ ò«·«Â«"
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
         Height          =   495
         Left            =   2280
         RightToLeft     =   -1  'True
         TabIndex        =   50
         Top             =   3225
         Width           =   4095
      End
   End
   Begin VB.Frame Frame6 
      Height          =   975
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   31
      Top             =   8520
      Width           =   13455
      Begin VB.CommandButton cmdClose 
         Caption         =   " Œ—ÊÃ  (Esc)"
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
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   33
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton cmdReCalculate 
         Caption         =   "„Õ«”»Â ”Êœ Ê “Ì«‰"
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
         Left            =   11520
         RightToLeft     =   -1  'True
         TabIndex        =   32
         Top             =   240
         Width           =   1815
      End
      Begin FLWCtrls.FWCoolButton FWBtnPrint 
         Height          =   615
         Left            =   1560
         TabIndex        =   40
         Top             =   240
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   1085
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Nazanin"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "frmTaraz.frx":A4C2
         DownPicture     =   "frmTaraz.frx":A7DC
         PictureAlign    =   3
         Caption         =   "F6+ç«Å"
         MaskColor       =   -2147483633
         Style           =   2
      End
      Begin VB.Label Label11 
         Alignment       =   2  'Center
         Caption         =   "›ﬁÿ «ÿ·«⁄«  Â“Ì‰Â ÕﬁÊﬁ »«Ìœ Ê«—œ ‘Êœ . »ﬁÌÂ «ÿ·«⁄«  »Â ’Ê—  « Ê„« Ìò «“ ”Ì” „ ›—Ê‘ „Õ«”»Â Ê Ê«—œ ‘œÂ «”  "
         BeginProperty Font 
            Name            =   "Nazanin"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   645
         Left            =   3735
         RightToLeft     =   -1  'True
         TabIndex        =   37
         Top             =   120
         Width           =   6570
      End
   End
   Begin VB.Frame Frame5 
      Height          =   1815
      Left            =   6720
      TabIndex        =   24
      Top             =   480
      Width           =   6855
      Begin VB.ComboBox cmbSalMali 
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
         Left            =   3960
         RightToLeft     =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   25
         Top             =   480
         Width           =   1455
      End
      Begin MSMask.MaskEdBox txtDateTo 
         Height          =   465
         Left            =   240
         TabIndex        =   26
         Top             =   1080
         Width           =   1395
         _ExtentX        =   2461
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
      Begin MSMask.MaskEdBox txtDateFrom 
         Height          =   465
         Left            =   240
         TabIndex        =   27
         Top             =   480
         Width           =   1395
         _ExtentX        =   2461
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
      Begin VB.Label lblDate 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "«„—Ê“"
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
         Height          =   495
         Left            =   3240
         RightToLeft     =   -1  'True
         TabIndex        =   34
         Top             =   1080
         Width           =   3495
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "”«· „«·Ì"
         BeginProperty Font 
            Name            =   "B Traffic"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5400
         RightToLeft     =   -1  'True
         TabIndex        =   30
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "«“  «—ÌŒ"
         BeginProperty Font 
            Name            =   "B Traffic"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   1680
         RightToLeft     =   -1  'True
         TabIndex        =   29
         Top             =   600
         Width           =   1065
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   " «  «—ÌŒ"
         BeginProperty Font 
            Name            =   "B Traffic"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   1800
         RightToLeft     =   -1  'True
         TabIndex        =   28
         Top             =   1080
         Width           =   945
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "”Êœ Ê “Ì«‰"
      BeginProperty Font 
         Name            =   "Nazanin"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   15
      Top             =   6120
      Width           =   6495
      Begin VB.Label lblTotalOtherSale2 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "B Nazanin"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   23
         Top             =   1200
         Width           =   2625
      End
      Begin VB.Label Label22 
         Alignment       =   1  'Right Justify
         Caption         =   "«÷«›Â „Ì ‘Êœ ”«Ì— œ—¬„œÂ«"
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
         Left            =   2880
         RightToLeft     =   -1  'True
         TabIndex        =   22
         Top             =   1200
         Width           =   3255
      End
      Begin VB.Label lblTotalHazineha2 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "B Nazanin"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   480
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   21
         Top             =   720
         Width           =   2625
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         Caption         =   "ﬂ”— „Ì ‘Êœ  Ã„⁄ ﬂ· Â“Ì‰Â Â«"
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
         Height          =   495
         Left            =   2880
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Top             =   720
         Width           =   3255
      End
      Begin VB.Label lblTotalSoodZian 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "B Nazanin"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   495
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   19
         Top             =   1800
         Width           =   2595
      End
      Begin VB.Label lblSoodZian 
         Alignment       =   1  'Right Justify
         Caption         =   "”Êœ (“Ì«‰ ) ⁄„·Ì« Ì"
         BeginProperty Font 
            Name            =   "Nazanin"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   495
         Left            =   3360
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   1800
         Width           =   2175
      End
      Begin VB.Label LblTotalBenefitLoss2 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "B Nazanin"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   17
         Top             =   240
         Width           =   2625
      End
      Begin VB.Label LblSoodZianNavizhe2 
         Alignment       =   1  'Right Justify
         Caption         =   "”Êœ (“Ì«‰ ) ‰«ÊÌéÂ"
         BeginProperty Font 
            Name            =   "Nazanin"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3840
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Top             =   360
         Width           =   2295
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "”«Ì— œ—¬„œÂ«"
      BeginProperty Font 
         Name            =   "Nazanin"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Index           =   0
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   360
      Width           =   6495
      Begin VB.Label Label24 
         Alignment       =   1  'Right Justify
         Caption         =   "«÷«›Â „Ì ‘Êœ œ—¬„œ „«·Ì«  Ê ⁄Ê«—÷"
         BeginProperty Font 
            Name            =   "Nazanin"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2760
         RightToLeft     =   -1  'True
         TabIndex        =   44
         Top             =   1800
         Width           =   3615
      End
      Begin VB.Label lblTotalTax 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "B Nazanin"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   43
         Top             =   1680
         Width           =   2025
      End
      Begin VB.Label Label21 
         Alignment       =   1  'Right Justify
         Caption         =   "«÷«›Â „Ì ‘Êœ œ—¬„œ ”«Ì— ”—ÊÌ” Â«"
         BeginProperty Font 
            Name            =   "Nazanin"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2760
         RightToLeft     =   -1  'True
         TabIndex        =   42
         Top             =   1320
         Width           =   3615
      End
      Begin VB.Label lblTotalService 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "B Nazanin"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   41
         Top             =   1200
         Width           =   2025
      End
      Begin VB.Label lblTotalPacking 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "B Nazanin"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   720
         Width           =   2025
      End
      Begin VB.Label Label23 
         Alignment       =   1  'Right Justify
         Caption         =   "«÷«›Â „Ì ‘Êœ œ—¬„œ ò—«ÌÂ Õ„·"
         BeginProperty Font 
            Name            =   "Nazanin"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2760
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   840
         Width           =   3615
      End
      Begin VB.Label lblTotalOtherSale 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "B Nazanin"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   495
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   2280
         Width           =   2595
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         Caption         =   "Ã„⁄ ”«Ì— œ—¬„œÂ«"
         BeginProperty Font 
            Name            =   "Nazanin"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   495
         Left            =   3360
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   2295
         Width           =   3015
      End
      Begin VB.Label lblTotalCarreeFee 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "B Nazanin"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   240
         Width           =   2025
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "«÷«›Â „Ì ‘Êœ œ—¬„œ »” Â »‰œÌ"
         BeginProperty Font 
            Name            =   "Nazanin"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3120
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   360
         Width           =   3255
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Â“Ì‰Â Â«"
      BeginProperty Font 
         Name            =   "Nazanin"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Index           =   0
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   3240
      Width           =   6495
      Begin VB.TextBox txtHoghough 
         Alignment       =   2  'Center
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "#,##0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1065
            SubFormatType   =   2
         EndProperty
         BeginProperty Font 
            Name            =   "B Nazanin"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   105
         TabIndex        =   45
         Text            =   "0"
         Top             =   735
         Width           =   2010
      End
      Begin VB.Label lblTotalHazineTax 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "B Nazanin"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   480
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   39
         Top             =   1770
         Width           =   2025
      End
      Begin VB.Label Label18 
         Alignment       =   1  'Right Justify
         Caption         =   "Ã„⁄ Â“Ì‰Â Â«Ì „«·Ì«  Ê ⁄Ê«—÷"
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
         Height          =   495
         Left            =   3000
         RightToLeft     =   -1  'True
         TabIndex        =   38
         Top             =   1785
         Width           =   2895
      End
      Begin VB.Label lblTotalLosses 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "B Nazanin"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   465
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   36
         Top             =   240
         Width           =   2025
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Ã„⁄ Â“Ì‰Â Â«Ì ÷«Ì⁄« "
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
         Height          =   495
         Left            =   2160
         RightToLeft     =   -1  'True
         TabIndex        =   35
         Top             =   240
         Width           =   3735
      End
      Begin VB.Label lblTotalHazineTolid 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "B Nazanin"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   480
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   1260
         Width           =   2025
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         Caption         =   "Ã„⁄ Â“Ì‰Â Â«Ì ﬁÌ„   „«„ ‘œÂ"
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
         Height          =   495
         Left            =   2160
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   1275
         Width           =   3735
      End
      Begin VB.Label lblTotalHazineha 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "B Nazanin"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   495
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   2310
         Width           =   2595
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         Caption         =   "Ã„⁄ ﬂ· Â“Ì‰Â Â«"
         BeginProperty Font 
            Name            =   "Nazanin"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   495
         Left            =   3360
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   2310
         Width           =   2175
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Ã„⁄ Â“Ì‰Â Â«Ì ÕﬁÊﬁ Ê œ” „“œ"
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
         Height          =   495
         Left            =   2520
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   720
         Width           =   3375
      End
   End
   Begin RESIZEKITLibCtl.ResizeKit ResizeKit1 
      Height          =   480
      Left            =   0
      OleObjectBlob   =   "frmTaraz.frx":AC2E
      TabIndex        =   0
      Top             =   0
      Width           =   480
   End
   Begin FLWCtrls.FWToolTip FWToolTip 
      Left            =   0
      Top             =   0
      _ExtentX        =   926
      _ExtentY        =   926
      ForeColor       =   -2147483625
      BackColor       =   65535
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "”Êœ Ê “Ì«‰ ⁄„·Ì« Ì (ÊÌéÂ)"
      BeginProperty Font 
         Name            =   "B Homa"
         Size            =   15.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   495
      Left            =   4920
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   -120
      Width           =   4455
   End
End
Attribute VB_Name = "frmTarazSoodZian"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private MyFormAddEditMode As EnumAddEditMode
Dim tmpTextMatrix As String
Dim Parameter() As Parameter

Dim TotalSellAmount As Currency
Dim TotalSellReturnAmount As Currency
Dim TotalFinalSellAmount As Currency
Dim TotalFinalSellReturnAmount As Currency
Dim TotalFirstPrice As Currency
Dim TotalSaleDiscount As Currency
 
Dim TotalBenefitLoss As Currency

Dim TotalCareeFee As Currency
Dim TotalPacking As Currency
Dim TotalService As Currency
Dim TotalTax As Currency
Dim TotalOtherSale As Currency
 
Dim TotalLosses As Currency
Dim TotalGeneral As Currency
Dim TotalHazineTolid As Currency
Dim TotalHazineTax As Currency
Dim TotalHazineha As Currency
 
Dim TotalSoodZian As Currency
Dim CrystalReport1

Dim i As Integer
    
Public Sub ExitForm()
    frmMsg.fwlblMsg.Caption = "¬Ì« »—«Ì Œ—ÊÃ «“  ›—„ «ÿ„Ì‰«‰ œ«—Ìœø"
    frmMsg.fwBtn(0).ButtonType = flwButtonOk
    frmMsg.fwBtn(1).ButtonType = flwButtonCancel
    frmMsg.fwBtn(0).Caption = "»·Ì"
    frmMsg.fwBtn(1).Caption = "ŒÌ—"
    frmMsg.Show vbModal
    If mvarMsgIdx = vbYes Then
        Unload Me
    End If
End Sub

Public Sub SetFirstToolBar()
    
End Sub


Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdReCalculate_Click()
    CalculateTotalLabels
End Sub

Private Sub Form_Activate()
    VarActForm = Me.Name
    
    SetFirstToolBar
End Sub

Private Sub cmbSalMali_Click()
    CalculateTotalLabels
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case Shift
          Case 0
              Select Case KeyCode
                  Case 27  ' Esc
                        Me.ExitForm
                  Case vbKeyF6  '
                        Printing
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

Private Sub Form_Load()
    On Error GoTo Err_Handler
    
    If ClsFormAccess.frmTarazSoodZian = False Then
        Unload Me
        Exit Sub
    End If
    
    If intVersion <> Diamond Then
        ShowDisMessage "ﬁ«»·Ì  ê—› ‰ ”Êœ Ê“Ì«‰ ›ﬁÿ œ— ‰”ŒÂ «·„«” ÊÃÊœ œ«—œ", 1500
        Unload Me
        Exit Sub
    End If
    
    formloadFlag = False
  
    CenterTop Me
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

    
    txtDateFrom.Text = Mid(AccountYear, 3, 2) & "/01" & "/01"
    txtDateTo.Text = Mid(clsDate.shamsi(Date), 3)
    
'    SetTooltipText
    FillSalMali
    
    lblDate.Caption = "«„—Ê“  " & clsDate.Find_DayOfWeek(Weekday(Date, vbSaturday)) & "   " & clsDate.shamsi(Date)
    
    Set CrystalReport1 = CreateObject("Crystal.CrystalReport")
    
    formloadFlag = True
    
    Exit Sub
    
Err_Handler:
    LogSaveNew "frmTaraz => ", err.Description, err.Number, err.Source, "Form_Load"
    ShowErrorMessage
    err.Clear
    Resume Next
End Sub
Private Sub FillSalMali()
    On Error GoTo Err_Handler
    
    cmbSalMali.Clear
    Set rs = RunStoredProcedure2RecordSet("Get_All_tAccountYears")
    Do While rs.EOF = False
        cmbSalMali.AddItem rs!AccountYear
        cmbSalMali.ItemData(cmbSalMali.ListCount - 1) = Val(rs.Fields("AccountYear"))
        rs.MoveNext
    Loop
    rs.Close
    Dim i As Integer
    For i = 0 To cmbSalMali.ListCount - 1
        If AccountYear = cmbSalMali.ItemData(i) Then
            cmbSalMali.ListIndex = i
            Exit For
        End If
    Next
    'If cmbSalMali.ListCount > 0 Then cmbSalMali.ListIndex = 0
    
    Exit Sub
    
Err_Handler:
    LogSaveNew "frmTaraz => ", err.Description, err.Number, err.Source, "FillSalMali"
    ShowErrorMessage
    err.Clear
End Sub
Private Sub FillInventory()
    On Error GoTo Err_Handler
    
    cmbInventory.Clear
    ReDim Parameter(1) As Parameter
    Parameter(0) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
    Parameter(1) = GenerateInputParameter("@Type", adInteger, 4, 0)
    Set rs = RunParametricStoredProcedure2Rec("GetInventory", Parameter)
    Do While rs.EOF = False
        cmbInventory.AddItem rs!Tafsili
        cmbInventory.ItemData(cmbInventory.ListCount - 1) = Val(rs.Fields("Description"))
        rs.MoveNext
    Loop
    rs.Close
    If cmbInventory.ListCount > 0 Then cmbInventory.ListIndex = 0
    
    Exit Sub
    
Err_Handler:
    LogSaveNew "frmTaraz => ", err.Description, err.Number, err.Source, "FillInventory"
    ShowErrorMessage
    err.Clear
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    VarActForm = ""
    If PosConnection.State = adStateOpen Then PosConnection.Close: Set PosConnection = Nothing

    Dim i As Integer
    
    Set CrystalReport1 = Nothing
    
    SaveSetting strMainKey, Me.Name, "Left", Me.left
    SaveSetting strMainKey, Me.Name, "Top", Me.top
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call PresetScreenSaver
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call PresetScreenSaver
End Sub

Private Sub FWBtnPrint_Click()
    Printing
End Sub

Private Sub Label19_Click()

End Sub

Private Sub LblTotalBenefitLoss_Change()
    LblTotalBenefitLoss2 = LblTotalBenefitLoss
End Sub

Private Sub LblTotalBenefitLoss2_Change()
    CalculateSoodZianVizhe
End Sub

Private Sub lblTotalHazineha2_Change()
    CalculateSoodZianVizhe
End Sub

Private Sub lblTotalOtherSale2_Change()
    CalculateSoodZianVizhe
End Sub

Private Sub ResizeKit1_ExitResize(ByVal XScale As Double, ByVal YScale As Double)

    If formloadFlag = True Then
        SaveSetting strMainKey, Me.Name, "Height", Me.Height
        SaveSetting strMainKey, Me.Name, "Width", Me.Width
    End If

End Sub
Public Sub Printing()
    
    On Error GoTo Err_Handler

    If Len(txtDateFrom.ClipText) < 6 Or Len(txtDateTo.ClipText) < 6 Then
        frmDisMsg.lblMessage = "  «—ÌŒ „⁄ »— Ê«—œ ﬂ‰Ìœ "
        frmDisMsg.Timer1.Enabled = True
        frmDisMsg.Show vbModal
        Exit Sub
    End If
    ReDim Parameter(6) As Parameter

    Parameter(0) = GenerateInputParameter("@SystemDate", adVarWChar, 10, clsDate.shamsi(Date))
    Parameter(1) = GenerateInputParameter("@SystemDay", adVarWChar, 10, clsDate.Find_DayOfWeek(Weekday(Date, vbSaturday)))
    Parameter(2) = GenerateInputParameter("@SystemTime", adVarWChar, 5, Mid(Str(time), 1, 5))
    Parameter(3) = GenerateInputParameter("@DateBefore", adInteger, 4, CLng(DateToNumber("13" & txtDateFrom.Text)))
    Parameter(4) = GenerateInputParameter("@DateAfter", adInteger, 4, CLng(DateToNumber("13" & txtDateTo.Text)))
    Parameter(5) = GenerateInputParameter("@AccountYear", adSmallInt, 2, Val(cmbSalMali.Text))
    Parameter(6) = GenerateInputParameter("@Branch", adInteger, 4, CurrentBranch)
'    Parameter(7) = GenerateInputParameter("@MojodiPrice", adBigInt, 8, Val(txtTotalMojodiPrice))
    
    CrystalReport1.ReportFileName = App_Path & "\AccountingReport" & "\RepSoodZian.rpt"

    Dim fileSystem As New FileSystemObject
    Dim IsFileExist As Boolean
    IsFileExist = fileSystem.FileExists(CrystalReport1.ReportFileName)
    If IsFileExist = False Then
        frmDisMsg.lblMessage = " ›«Ì· ê“«—‘ ”ÊœÊ “Ì«‰ (RepSoodZian) ÅÌœ« ‰‘œ "
        frmDisMsg.Timer1.Interval = 2000
        frmDisMsg.Timer1.Enabled = True
        frmDisMsg.Show vbModal
        Exit Sub
    End If

    CrystalReport1.ReportTitle = "  ê“«—‘ ”Êœ Ê “Ì«‰ ”«· " & cmbSalMali.Text
    CrystalReport1.Destination = crptToWindow 'crptToPrinter '
    Dim intIndex As Integer

    For intIndex = 0 To UBound(Parameter) - LBound(Parameter)
        CrystalReport1.ParameterFields(intIndex) = CStr(Parameter(intIndex).Name) & ";" & CStr(Parameter(intIndex).Value) & ";" & "True"
    Next intIndex

    CrystalReport1.WindowState = crptMaximized
    ODBCSetting clsArya.ServerName, clsArya.DbName
    CrystalReport1.Connect = CrystallConnection
    CrystalReport1.Action = 1
    CrystalReport1.RetrieveDataFiles

    If Screen.Width > 12000 Then
        CrystalReport1.PageZoom (100)
    Else
        CrystalReport1.PageZoom (75)
    End If

Exit Sub

Err_Handler:
    LogSaveNew "frmTaraz => ", err.Description, err.Number, err.Source, "Printing"
    ShowErrorMessage
    err.Clear

End Sub

Private Sub CalculateTotalLabels()
    On Error GoTo Err_Handler
    
    If cmbSalMali.ListIndex = -1 Then Exit Sub
    
    If Len(txtDateFrom.ClipText) < 6 Or Len(txtDateTo.ClipText) < 6 Then
        frmDisMsg.lblMessage = "  «—ÌŒ „⁄ »— Ê«—œ ﬂ‰Ìœ "
        frmDisMsg.Timer1.Enabled = True
        frmDisMsg.Show vbModal
        Exit Sub
    End If
    MyFormAddEditMode = EnumAddEditMode.ViewMode 'VIEW Mode
    
    SetFirstToolBar
    
    TotalBenefitLoss = 0
    TotalSellReturnAmount = 0
    TotalFinalSellReturnAmount = 0
    TotalSellAmount = 0
    TotalFinalSellAmount = 0
    TotalFirstPrice = 0
    
    TotalSaleDiscount = 0
    TotalBuyDiscount = 0
    TotalHazineTolid = 0
    TotalHazineTax = 0
    TotalLosses = 0
    TotalGeneral = 0
    TotalHazineha = 0
    TotalSoodZian = 0
    
    LblTotalSellAmount = ""
    LblTotalSellReturnAmount = ""
    lblTotalFinalSellAmount = ""
    LblTotalFinalSellReturnAmount = ""
    LblTotalSaleDiscount = ""
    
    LblTotalBenefitLoss = ""
    
    lblTotalCarreeFee = ""
    lblTotalPacking = ""
    lblTotalService = ""
    lblTotalTax = ""
    lblTotalOtherSale = ""
    
    lblTotalLosses = ""
    lblTotalHazineTolid = ""
    lblTotalHazineTax = ""
    lblTotalHazineha = ""
    
    LblTotalBenefitLoss2 = ""
    lblTotalOtherSale2 = ""
    lblTotalHazineha2 = ""
    
    Me.MousePointer = vbHourglass
    DoEvents
    Dim L_Rst As New ADODB.Recordset
    
    ReDim Parameter(3)
    Parameter(0) = GenerateInputParameter("@DateBefore", adVarWChar, 8, txtDateFrom.Text)
    Parameter(1) = GenerateInputParameter("@DateAfter", adVarWChar, 8, txtDateTo.Text)
    Parameter(2) = GenerateInputParameter("@AccountYear", adSmallInt, 2, Val(cmbSalMali.Text))
    Parameter(3) = GenerateInputParameter("@Branch", adInteger, 4, CurrentBranch)
    
    Set L_Rst = RunParametricStoredProcedure2Rec("Get_TarazSoodZian_Sale", Parameter)
    
    If L_Rst.BOF = True And L_Rst.EOF = True Then
        Set L_Rst = Nothing
    Else
        TotalSellAmount = Val(L_Rst!TotalSellAmount)
        TotalSellReturnAmount = Val(L_Rst!TotalSellReturnAmount)
        TotalFirstPrice = Val(L_Rst!TotalFirstPrice)
        TotalFinalSellAmount = Val(L_Rst!TotalFinalSellAmount)
        TotalFinalSellReturnAmount = Val(L_Rst!TotalFinalSellReturnAmount)
        TotalSaleDiscount = Val(L_Rst!TotalSaleDiscount)
        TotalBuyDiscount = Val(L_Rst!TotalBuyDiscount)
        TotalCareeFee = Val(L_Rst!TotalCareeFee)
        TotalPacking = Val(L_Rst!TotalPacking)
        TotalService = Val(L_Rst!TotalService)
        TotalTax = Val(L_Rst!TotalTax)
        TotalLosses = Val(L_Rst!TotalLosses)
        TotalGeneral = Val(L_Rst!TotalGeneral)
        TotalHazineTolid = Val(L_Rst!TotalHazineTolid)
        TotalHazineTax = Val(L_Rst!TotalHazineTax)
        L_Rst.Close
    End If
    Set L_Rst = Nothing
        
    LblTotalSellAmount.Caption = Format(TotalSellAmount, "#,##")
    
    lblTotalFinalSellAmount.Caption = Format(TotalFinalSellAmount, "#,##")
    lblTotalFinalSellAmount.Caption = "(" & lblTotalFinalSellAmount.Caption & ")"
    
    LblTotalSellReturnAmount.Caption = Format(TotalSellReturnAmount, "#,##")
    LblTotalSellReturnAmount.Caption = "(" & LblTotalSellReturnAmount.Caption & ")"
    
    LblTotalFinalSellReturnAmount.Caption = Format(TotalFinalSellReturnAmount, "#,##")
    
    LblTotalSaleDiscount.Caption = Format(TotalSaleDiscount, "#,##")
    LblTotalSaleDiscount.Caption = "(" & LblTotalSaleDiscount.Caption & ")"
    
    CalculateBenefitLoss
    
    lblTotalCarreeFee.Caption = Format(TotalCareeFee, "#,##")
    lblTotalPacking.Caption = Format(TotalPacking, "#,##")
    lblTotalService.Caption = Format(TotalService, "#,##")
    lblTotalTax.Caption = Format(TotalTax, "#,##")
    TotalOtherSale = TotalCareeFee + TotalPacking + TotalService + TotalTax
    lblTotalOtherSale.Caption = Format(TotalOtherSale, "#,##") & clsArya.UnitPrice
    lblTotalOtherSale2.Caption = Format(TotalOtherSale, "#,##") & clsArya.UnitPrice

    lblTotalHazineTolid.Caption = Format(TotalHazineTolid, "#,##")
    lblTotalHazineTax.Caption = Format(TotalHazineTax, "#,##")
''    txtHoghough = Format(TotalGeneral, "#,##")
    txtHoghough = TotalGeneral
    lblTotalLosses.Caption = Format(TotalLosses, "#,##")
    TotalHazineha = Val(txtHoghough) + TotalLosses + TotalHazineTolid + TotalHazineTax
    lblTotalHazineha.Caption = Format(TotalHazineha, "#,##") & clsArya.UnitPrice
    lblTotalHazineha2.Caption = "(" & Format(TotalHazineha, "#,##") & ")" & clsArya.UnitPrice

'    LblTotalBenefitLoss2 = TotalBenefitLoss
'    lblTotalHazineha2 = TotalHazineha
'    lblTotalOtherSale2 = TotalOtherSale
    
    Me.MousePointer = vbDefault
Exit Sub
    
Err_Handler:
    ShowErrorMessage
    LogSaveNew "frmTaraz => ", err.Description, err.Number, err.Source, "CalculateTotalLabels"
    Me.MousePointer = vbDefault
    Me.MousePointer = vbDefault
    Resume Next
End Sub
Private Sub SetTooltipText()
    
    With FWToolTip
        .BackColor = vbYellow
        .Ballon = True
        .Margin(flwToolTipMarginLeft) = 20
        .MaxWidth = 500
        .DelayTime(flwToolTipDelayDefault) = 100
        '.DelayTime(flwToolTipDelayInitial) = 100
        .DelayTime(flwToolTipDelayShow) = 3000
        .DelayTime(flwToolTipDelayReshow) = 1500
        .Text(frmTaraz) = "›—„Ê· „Õ«”»Â ”Êœ- “Ì«‰ ﬂ·:" & vbCrLf & vbCrLf & "((ﬂ· ›—Ê‘ „‰Â«Ì ﬂ· »—ê‘  «“ ›—Ê‘)-(ﬂ· »Â«¡  „«„ ‘œÂ ›—Ê‘ Ê » «“ ›—Ê‘))"
       End With
End Sub

Private Sub txtHoghough_Change()
    CalculateSoodZianVizhe
End Sub

Private Sub CalculateBenefitLoss()
    
    TotalBenefitLoss = TotalSellAmount - TotalSellReturnAmount - TotalFinalSellAmount - TotalSaleDiscount + TotalFinalSellReturnAmount
    LblTotalBenefitLoss.Caption = Format(TotalBenefitLoss, "#,##") & clsArya.UnitPrice
    
    If TotalBenefitLoss >= 0 Then
        LblSoodZianNavizhe = " ”Êœ ‰«ÊÌéÂ"
        LblSoodZianNavizhe2 = " ”Êœ ‰«ÊÌéÂ"
    Else
        LblSoodZianNavizhe = "“Ì«‰ ‰«ÊÌéÂ"
        LblSoodZianNavizhe2 = "“Ì«‰ ‰«ÊÌéÂ"
        LblTotalBenefitLoss.Caption = "(" & Format(Abs(TotalBenefitLoss), "#,##") & ")" & clsArya.UnitPrice
    End If

End Sub
Private Sub CalculateSoodZianVizhe()
    TotalHazineha = Val(txtHoghough) + TotalLosses + TotalHazineTolid + TotalHazineTax
    lblTotalHazineha.Caption = Format(TotalHazineha, "#,##") & clsArya.UnitPrice
    TotalSoodZian = TotalBenefitLoss - TotalHazineha + TotalOtherSale
    lblTotalSoodZian = TotalSoodZian
    lblTotalSoodZian.Caption = Format(TotalSoodZian, "#,##") & clsArya.UnitPrice
    If TotalSoodZian >= 0 Then
        lblSoodZian = " ”Êœ ⁄„·Ì« Ì"
    Else
        lblSoodZian = "“Ì«‰ ⁄„·Ì« Ì"
        lblTotalSoodZian.Caption = "(" & Format(Abs(TotalSoodZian), "#,##") & ")" & clsArya.UnitPrice
    End If
End Sub

Private Sub txtTotalMojodiPrice_KeyPress(KeyAscii As Integer)
    If IsNumeric(ChrW(KeyAscii)) = False And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub
