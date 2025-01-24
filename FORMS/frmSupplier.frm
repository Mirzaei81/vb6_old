VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{7AEDC602-D94C-11D1-BB7A-00E0290EA3C9}#1.0#0"; "ResizeKit.ocx"
Object = "{75D4F767-8785-11D3-93AD-0000832EF44D}#2.14#0"; "FAST2001.ocx"
Begin VB.Form frmSupplier 
   ClientHeight    =   10095
   ClientLeft      =   60
   ClientTop       =   465
   ClientWidth     =   11910
   BeginProperty Font 
      Name            =   "B Homa"
      Size            =   9.75
      Charset         =   178
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSupplier.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   10095
   ScaleWidth      =   11910
   Begin VB.Frame frameAccounting 
      Caption         =   "Õ”«»œ«—Ì"
      ForeColor       =   &H00008000&
      Height          =   3195
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   52
      Top             =   480
      Width           =   3975
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
         Left            =   240
         RightToLeft     =   -1  'True
         TabIndex        =   58
         Top             =   360
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
         Left            =   240
         RightToLeft     =   -1  'True
         TabIndex        =   57
         Top             =   240
         Visible         =   0   'False
         Width           =   2175
      End
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
         Left            =   240
         RightToLeft     =   -1  'True
         TabIndex        =   56
         Top             =   1920
         Width           =   1455
      End
      Begin VB.TextBox txtPrimaryTalab 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   240
         RightToLeft     =   -1  'True
         TabIndex        =   55
         ToolTipText     =   "„»·€ »œÂÌ »œÊ‰ ⁄·«„  Ê „»·€ »” «‰ﬂ«—Ì »« ⁄·«„  „‰›Ì Ê«—œ ‘Êœ"
         Top             =   1320
         Width           =   1455
      End
      Begin VB.TextBox txtPrimaryBedehi 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   480
         Left            =   240
         RightToLeft     =   -1  'True
         TabIndex        =   54
         ToolTipText     =   "„»·€ »œÂÌ »œÊ‰ ⁄·«„  Ê „»·€ »” «‰ﬂ«—Ì »« ⁄·«„  „‰›Ì Ê«—œ ‘Êœ"
         Top             =   840
         Width           =   1455
      End
      Begin VB.CommandButton cmdAddTafsili 
         Caption         =   "«÷«›Â ﬂ—œ‰ ﬂ·ÌÂ „‘ —Ì«‰ ÃœÌœ »Â Õ”«»œ«—Ì"
         BeginProperty Font 
            Name            =   "Nazanin"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   240
         RightToLeft     =   -1  'True
         TabIndex        =   53
         Top             =   2400
         Width           =   3495
      End
      Begin VB.Label lblTafsili 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "ﬂœ  ›÷Ì·Ì"
         ForeColor       =   &H00008000&
         Height          =   525
         Left            =   1680
         RightToLeft     =   -1  'True
         TabIndex        =   62
         Top             =   360
         Width           =   915
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "‘„«—Â ”‰œ"
         ForeColor       =   &H00008000&
         Height          =   525
         Left            =   2400
         RightToLeft     =   -1  'True
         TabIndex        =   61
         Top             =   1920
         Width           =   1155
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "„«‰œÂ «Ê·ÌÂ-ÿ·»"
         ForeColor       =   &H00008000&
         Height          =   525
         Left            =   1920
         TabIndex        =   60
         Top             =   1320
         Width           =   1635
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "„«‰œÂ «Ê·ÌÂ - »œÂÌ"
         ForeColor       =   &H000000C0&
         Height          =   525
         Left            =   1920
         TabIndex        =   59
         Top             =   840
         Width           =   1635
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Frame frameCompany 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   2340
      Left            =   7800
      RightToLeft     =   -1  'True
      TabIndex        =   44
      Top             =   1395
      Width           =   4040
      Begin VB.ComboBox cmbActKind 
         BackColor       =   &H0080C0FF&
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
         ItemData        =   "frmSupplier.frx":A4C2
         Left            =   240
         List            =   "frmSupplier.frx":A4C4
         RightToLeft     =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   46
         Top             =   1200
         Width           =   2400
      End
      Begin VB.TextBox txtWorkName 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0080C0FF&
         BeginProperty Font 
            Name            =   "Nazanin"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   210
         RightToLeft     =   -1  'True
         TabIndex        =   45
         Top             =   360
         Width           =   2535
      End
      Begin FLWCtrls.FWButton FWBtnActKind 
         Height          =   405
         Left            =   360
         TabIndex        =   47
         ToolTipText     =   " €ÌÌ— Ê«ÕœÂ«"
         Top             =   1200
         Visible         =   0   'False
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   714
         ButtonType      =   5
         Caption         =   "FWButton1"
         BackColor       =   12632256
         FontName        =   "MS Sans Serif"
         FontBold        =   0   'False
         FontSize        =   8.25
         Object.ToolTipText     =   " €ÌÌ— Ê«ÕœÂ«"
      End
      Begin VB.Label lblActKind 
         Alignment       =   1  'Right Justify
         Caption         =   "‰Ê⁄ ›⁄«·Ì "
         ForeColor       =   &H80000002&
         Height          =   405
         Left            =   2640
         RightToLeft     =   -1  'True
         TabIndex        =   49
         Top             =   1200
         Width           =   1155
      End
      Begin VB.Label lblWorkName 
         Alignment       =   1  'Right Justify
         Caption         =   "* ‰«„ „Õ·"
         ForeColor       =   &H80000002&
         Height          =   405
         Left            =   2760
         RightToLeft     =   -1  'True
         TabIndex        =   48
         Top             =   360
         Width           =   1155
      End
   End
   Begin VB.Frame Frame6 
      Height          =   3255
      Left            =   4120
      TabIndex        =   29
      Top             =   480
      Width           =   3600
      Begin VB.TextBox txtPostalCode 
         Alignment       =   1  'Right Justify
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
         Left            =   160
         RightToLeft     =   -1  'True
         TabIndex        =   70
         Top             =   2760
         Width           =   1905
      End
      Begin VB.TextBox txtTel1 
         Alignment       =   1  'Right Justify
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
         Left            =   160
         RightToLeft     =   -1  'True
         TabIndex        =   36
         Top             =   240
         Width           =   1900
      End
      Begin VB.TextBox txtTel2 
         Alignment       =   1  'Right Justify
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
         Left            =   160
         RightToLeft     =   -1  'True
         TabIndex        =   35
         Top             =   660
         Width           =   1900
      End
      Begin VB.TextBox txtTel3 
         Alignment       =   1  'Right Justify
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
         Left            =   160
         RightToLeft     =   -1  'True
         TabIndex        =   34
         Top             =   1080
         Width           =   1900
      End
      Begin VB.TextBox txtTel4 
         Alignment       =   1  'Right Justify
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
         Left            =   160
         RightToLeft     =   -1  'True
         TabIndex        =   33
         Top             =   1500
         Width           =   1900
      End
      Begin VB.TextBox txtFax 
         Alignment       =   1  'Right Justify
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
         Left            =   160
         RightToLeft     =   -1  'True
         TabIndex        =   32
         Top             =   2340
         Width           =   1900
      End
      Begin VB.TextBox txtMobile 
         Alignment       =   1  'Right Justify
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
         Left            =   160
         RightToLeft     =   -1  'True
         TabIndex        =   31
         Top             =   1920
         Width           =   1900
      End
      Begin VB.TextBox txtEmail 
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
         Left            =   160
         TabIndex        =   30
         Top             =   2760
         Visible         =   0   'False
         Width           =   1900
      End
      Begin VB.Label lblPostalCode 
         Alignment       =   1  'Right Justify
         Caption         =   "ﬂœ Å” Ì"
         ForeColor       =   &H80000002&
         Height          =   405
         Left            =   2400
         RightToLeft     =   -1  'True
         TabIndex        =   71
         Top             =   2760
         Width           =   1035
      End
      Begin VB.Label lblTel1 
         Alignment       =   1  'Right Justify
         Caption         =   " ·›‰1"
         ForeColor       =   &H80000002&
         Height          =   405
         Left            =   2055
         RightToLeft     =   -1  'True
         TabIndex        =   43
         Top             =   240
         Width           =   1395
      End
      Begin VB.Label lblFax 
         Alignment       =   1  'Right Justify
         Caption         =   "›«ﬂ”"
         ForeColor       =   &H80000002&
         Height          =   405
         Left            =   2055
         RightToLeft     =   -1  'True
         TabIndex        =   42
         Top             =   2355
         Width           =   1395
      End
      Begin VB.Label lblMobile 
         Alignment       =   1  'Right Justify
         Caption         =   "„Ê»«Ì·"
         ForeColor       =   &H80000002&
         Height          =   405
         Left            =   2055
         RightToLeft     =   -1  'True
         TabIndex        =   41
         Top             =   1935
         Width           =   1395
      End
      Begin VB.Label lblTel4 
         Alignment       =   1  'Right Justify
         Caption         =   " ·›‰4"
         ForeColor       =   &H80000002&
         Height          =   405
         Left            =   2055
         RightToLeft     =   -1  'True
         TabIndex        =   40
         Top             =   1515
         Width           =   1395
      End
      Begin VB.Label lblTel3 
         Alignment       =   1  'Right Justify
         Caption         =   " ·›‰3"
         ForeColor       =   &H80000002&
         Height          =   405
         Left            =   2055
         RightToLeft     =   -1  'True
         TabIndex        =   39
         Top             =   1080
         Width           =   1395
      End
      Begin VB.Label lblTel2 
         Alignment       =   1  'Right Justify
         Caption         =   " ·›‰2"
         ForeColor       =   &H80000002&
         Height          =   405
         Left            =   2055
         RightToLeft     =   -1  'True
         TabIndex        =   38
         Top             =   660
         Width           =   1395
      End
      Begin VB.Label lblEmail 
         Alignment       =   1  'Right Justify
         Caption         =   "Å”  «·ﬂ —Ê‰ÌﬂÌ"
         ForeColor       =   &H80000002&
         Height          =   405
         Left            =   2055
         RightToLeft     =   -1  'True
         TabIndex        =   37
         Top             =   2775
         Visible         =   0   'False
         Width           =   1395
      End
   End
   Begin VB.Frame Frame5 
      ForeColor       =   &H8000000D&
      Height          =   2055
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   24
      Top             =   3600
      Width           =   7665
      Begin VB.TextBox txtNationalCode 
         Alignment       =   1  'Right Justify
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
         Left            =   3720
         RightToLeft     =   -1  'True
         TabIndex        =   69
         Top             =   840
         Width           =   2265
      End
      Begin VB.TextBox txtEconomicCode 
         Alignment       =   1  'Right Justify
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
         Left            =   3720
         RightToLeft     =   -1  'True
         TabIndex        =   68
         Top             =   360
         Width           =   2265
      End
      Begin VB.TextBox txtDescription 
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
         Height          =   660
         Left            =   3720
         MultiLine       =   -1  'True
         RightToLeft     =   -1  'True
         TabIndex        =   66
         Text            =   "frmSupplier.frx":A4C6
         Top             =   1320
         Width           =   3825
      End
      Begin VB.TextBox txtDiscount 
         Alignment       =   1  'Right Justify
         ForeColor       =   &H00000080&
         Height          =   405
         Left            =   2760
         RightToLeft     =   -1  'True
         TabIndex        =   63
         ToolTipText     =   "„Ì“«‰ œ—’œ »«Ìœ ﬂ„ — «“ 100 »«‘œ"
         Top             =   240
         Visible         =   0   'False
         Width           =   735
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
         Left            =   1800
         RightToLeft     =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   51
         Top             =   1440
         Width           =   1725
      End
      Begin VB.ComboBox cmbCity 
         Height          =   465
         ItemData        =   "frmSupplier.frx":A4D0
         Left            =   120
         List            =   "frmSupplier.frx":A4D2
         RightToLeft     =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   26
         Top             =   780
         Width           =   1935
      End
      Begin VB.ComboBox CmbState 
         Height          =   465
         ItemData        =   "frmSupplier.frx":A4D4
         Left            =   120
         List            =   "frmSupplier.frx":A4D6
         RightToLeft     =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   25
         Top             =   255
         Width           =   1935
      End
      Begin FLWCtrls.FWButton FWBtnPerson 
         Height          =   585
         Left            =   120
         TabIndex        =   50
         Top             =   1320
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   1032
         ButtonType      =   8
         Caption         =   "ÊÌ“Ì Ê—"
         BackColor       =   12632256
         ForeColor       =   16384
         Alignment       =   1
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "òœ «ﬁ ’«œÌ"
         ForeColor       =   &H80000002&
         Height          =   405
         Left            =   6000
         RightToLeft     =   -1  'True
         TabIndex        =   73
         Top             =   240
         Width           =   1395
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "òœ/‘‰«”Â „·Ì"
         ForeColor       =   &H80000002&
         Height          =   405
         Left            =   6000
         RightToLeft     =   -1  'True
         TabIndex        =   72
         Top             =   840
         Width           =   1395
      End
      Begin VB.Label lblDiscount 
         Alignment       =   1  'Right Justify
         Caption         =   "„Ì“«‰  Œ›Ì›"
         ForeColor       =   &H00004080&
         Height          =   405
         Left            =   3360
         RightToLeft     =   -1  'True
         TabIndex        =   65
         Top             =   240
         Visible         =   0   'False
         Width           =   1395
      End
      Begin VB.Label lblDiscount2 
         Alignment       =   1  'Right Justify
         Caption         =   "œ—’œ"
         ForeColor       =   &H00000080&
         Height          =   405
         Left            =   2640
         RightToLeft     =   -1  'True
         TabIndex        =   64
         Top             =   240
         Visible         =   0   'False
         Width           =   705
      End
      Begin VB.Label lblCity 
         Alignment       =   1  'Right Justify
         Caption         =   "‘Â—"
         ForeColor       =   &H80000002&
         Height          =   360
         Left            =   2040
         RightToLeft     =   -1  'True
         TabIndex        =   28
         Top             =   855
         Width           =   555
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "«” «‰"
         ForeColor       =   &H80000002&
         Height          =   405
         Left            =   2160
         RightToLeft     =   -1  'True
         TabIndex        =   27
         Top             =   345
         Width           =   555
      End
   End
   Begin VB.PictureBox Frame2 
      Height          =   855
      Left            =   9840
      ScaleHeight     =   795
      ScaleWidth      =   1875
      TabIndex        =   21
      Top             =   600
      Width           =   1935
      Begin VB.OptionButton OptionActDeAct 
         Alignment       =   1  'Right Justify
         Caption         =   "›⁄«·"
         ForeColor       =   &H80000002&
         Height          =   285
         Index           =   0
         Left            =   390
         RightToLeft     =   -1  'True
         TabIndex        =   23
         Top             =   480
         Width           =   1095
      End
      Begin VB.OptionButton OptionActDeAct 
         Alignment       =   1  'Right Justify
         Caption         =   "€Ì—›⁄«·"
         ForeColor       =   &H80000002&
         Height          =   345
         Index           =   1
         Left            =   360
         RightToLeft     =   -1  'True
         TabIndex        =   22
         Top             =   10
         Width           =   1125
      End
   End
   Begin VB.PictureBox frameOwner 
      Height          =   855
      Left            =   7800
      ScaleHeight     =   795
      ScaleWidth      =   1875
      TabIndex        =   18
      Top             =   600
      Width           =   1935
      Begin VB.OptionButton OptionOwner 
         Alignment       =   1  'Right Justify
         Caption         =   "›—œ"
         ForeColor       =   &H80000002&
         Height          =   315
         Index           =   0
         Left            =   570
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Top             =   10
         Width           =   765
      End
      Begin VB.OptionButton OptionOwner 
         Alignment       =   1  'Right Justify
         Caption         =   "‘—ò "
         ForeColor       =   &H80000002&
         Height          =   285
         Index           =   1
         Left            =   240
         RightToLeft     =   -1  'True
         TabIndex        =   19
         Top             =   480
         Width           =   1095
      End
   End
   Begin VB.TextBox txtCode 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   8760
      RightToLeft     =   -1  'True
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   3075
      Visible         =   0   'False
      Width           =   1290
   End
   Begin VB.TextBox txtMembershipId 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   13
      Top             =   0
      Width           =   1425
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
      Height          =   2085
      Left            =   7800
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   3600
      Width           =   4035
      Begin VB.TextBox txtAddress 
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
         Height          =   1260
         Left            =   120
         MultiLine       =   -1  'True
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   600
         Width           =   3795
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "¬œ—”*"
         ForeColor       =   &H80000002&
         Height          =   285
         Left            =   2760
         RightToLeft     =   -1  'True
         TabIndex        =   67
         Top             =   240
         Width           =   1035
      End
   End
   Begin VB.Frame framePerson 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   2340
      Left            =   7800
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   1395
      Width           =   4040
      Begin VB.TextBox txtFamily 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0080C0FF&
         BeginProperty Font 
            Name            =   "Nazanin"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   165
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   1800
         Width           =   2415
      End
      Begin VB.TextBox txtName 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0080C0FF&
         BeginProperty Font 
            Name            =   "Nazanin"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   165
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   1270
         Width           =   2415
      End
      Begin VB.ComboBox cmbGender 
         BackColor       =   &H0080C0FF&
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
         ItemData        =   "frmSupplier.frx":A4D8
         Left            =   165
         List            =   "frmSupplier.frx":A4DA
         RightToLeft     =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   210
         Width           =   2415
      End
      Begin VB.ComboBox cmbPrefix 
         BackColor       =   &H0080C0FF&
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
         ItemData        =   "frmSupplier.frx":A4DC
         Left            =   645
         List            =   "frmSupplier.frx":A4DE
         RightToLeft     =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   740
         Width           =   1935
      End
      Begin FLWCtrls.FWButton FWBtnPrefix 
         Height          =   405
         Left            =   165
         TabIndex        =   5
         ToolTipText     =   " €ÌÌ— Ê«ÕœÂ«"
         Top             =   750
         Visible         =   0   'False
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   714
         ButtonType      =   5
         Caption         =   "FWButton1"
         BackColor       =   12632256
         FontName        =   "MS Sans Serif"
         FontBold        =   0   'False
         FontSize        =   8.25
         Object.ToolTipText     =   " €ÌÌ— Ê«ÕœÂ«"
      End
      Begin VB.Label lblFamily 
         Alignment       =   1  'Right Justify
         Caption         =   "* ‰«„ Œ«‰Ê«œêÌ"
         ForeColor       =   &H80000002&
         Height          =   405
         Left            =   2550
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   1860
         Width           =   1365
      End
      Begin VB.Label lblName 
         Alignment       =   1  'Right Justify
         Caption         =   " ‰«„"
         ForeColor       =   &H80000002&
         Height          =   405
         Left            =   2520
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   1305
         Width           =   1245
      End
      Begin VB.Label lblSex 
         Alignment       =   1  'Right Justify
         Caption         =   "Ã‰”Ì "
         ForeColor       =   &H80000002&
         Height          =   405
         Left            =   2520
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   240
         Width           =   1245
      End
      Begin VB.Label lblPrefix 
         Alignment       =   1  'Right Justify
         Caption         =   "⁄‰Ê«‰"
         ForeColor       =   &H80000002&
         Height          =   405
         Left            =   2400
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   765
         Width           =   1365
      End
   End
   Begin VSFlex7LCtl.VSFlexGrid vsCustomer 
      Height          =   4275
      Left            =   240
      TabIndex        =   12
      Top             =   5760
      Width           =   11595
      _cx             =   20452
      _cy             =   7541
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
      Cols            =   7
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   500
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmSupplier.frx":A4E0
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   1
      AutoSearchDelay =   2
      MultiTotals     =   -1  'True
      SubtotalPosition=   1
      OutlineBar      =   0
      OutlineCol      =   0
      Ellipsis        =   0
      ExplorerBar     =   3
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
      Height          =   525
      Left            =   10440
      Top             =   0
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   926
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
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
   Begin RESIZEKITLibCtl.ResizeKit ResizeKit1 
      Height          =   480
      Left            =   0
      OleObjectBlob   =   "frmSupplier.frx":A5A4
      TabIndex        =   17
      Top             =   0
      Width           =   480
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "«ÿ·«⁄«   «„Ì‰ ò‰‰œê«‰"
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
      Left            =   4920
      RightToLeft     =   -1  'True
      TabIndex        =   16
      Top             =   0
      Width           =   2775
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "* ﬂœ «‘ —«ﬂ"
      ForeColor       =   &H000000C0&
      Height          =   405
      Left            =   1680
      RightToLeft     =   -1  'True
      TabIndex        =   15
      Top             =   0
      Width           =   1335
   End
End
Attribute VB_Name = "frmSupplier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private clsDate As New clsDate
Private cn As New ADODB.Connection
Private Rc As New ADODB.Recordset
Private rctmp As New ADODB.Recordset
Private MyFormAddEditMode As EnumAddEditMode
Dim Parameter() As Parameter
Dim i As Integer
Dim OldTafsili As Long

Public Sub Delete()
    
End Sub

Private Sub FillvsCustomer()

    Dim Rst As New ADODB.Recordset
    ReDim Parameter(1) As Parameter
    
    
    Parameter(0) = GenerateInputParameter("@MainCust", adBoolean, 1, 1)
    Parameter(1) = GenerateInputParameter("@Branch", adInteger, 4, CurrentBranch)
    Set Rst = RunParametricStoredProcedure2Rec("Get_Supplier", Parameter)
    
    With vsCustomer
        .Rows = 1
        i = 0
        While Rst.EOF <> True
            .Rows = .Rows + 1
            i = .Rows - 1
            .TextMatrix(i, 0) = i
            .TextMatrix(i, 1) = Rst!Code
            .TextMatrix(i, 2) = Rst![Full Name]
            .TextMatrix(i, 3) = Rst!MembershipId
            If Rst!WorkName.Value <> "" Then
                .TextMatrix(i, 4) = -1
            End If
            .TextMatrix(i, 5) = Rst!address
            .TextMatrix(i, 6) = Rst!Discount
            Rst.MoveNext
        Wend
    End With
    Set Rst = Nothing
End Sub


Private Sub cmdAddTafsili_Click()
    ShowMessage "¬Ì« „Ì ŒÊ«ÂÌœ »—«Ì ﬂ·ÌÂ „‘ —Ì«‰ Ê «‘Œ«’ '  ›÷Ì·Ì ÃœÌœ œ— ”Ì” „ Õ”«»œ«—Ì «ÌÕ«œ ò‰Ìœø ", True, True, "»·Ì", "ŒÌ—"
    If modgl.mvarMsgIdx = vbNo Then
        Exit Sub
    End If
    
    ReDim Parameter(0) As Parameter
    Dim Rst As New ADODB.Recordset
    
    Parameter(0) = GenerateInputParameter("@Branch", adInteger, 4, cmbBranch.ItemData(cmbBranch.ListIndex))
    Set Rst = RunParametricStoredProcedure2Rec("Get_All_Supplier", Parameter)
    
    txtFamily.Text = ""
    txtWorkName.Text = ""
    While Rst.EOF <> True
        txtTafsiliCode.Text = ""
        TxtName.Text = Rst!Name
        If IsNull(Rst!Tafsili) = True Or Trim(Rst!Tafsili) = "" Then Insert_Tafsili Rst!Code, False
        Rst.MoveNext
    Wend
    If Rst.State = 1 Then Rst.Close
    If cn.State = 1 Then cn.Close
    Set Rst = Nothing
    Set cn = Nothing
    
    ShowDisMessage " ⁄—Ì› „‘ —Ì«‰ œ— ”Ì” „ Õ”«»œ«—Ì «‰Ã«„ ê—œÌœ", 1000
    DefaultSettings

End Sub

Private Sub Form_Activate()

    VarActForm = Me.Name
    SetFirstToolBar

End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call PresetScreenSaver
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call PresetScreenSaver
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

    CenterTop Me
    
    If ClsFormAccess.frmSupplier = False Then
        Unload Me
        Exit Sub
    End If
    
    If intVersion = Min Or intVersion = Normal Then
        ShowDisMessage " ⁄—Ì›  «„Ì‰ ﬂ‰‰œê«‰ œ— ‰”ŒÂ ÅÌ‘—› Â Ê »«·« — ÊÃÊœ œ«—œ", 1500
        Unload Me
        Exit Sub
    End If
    
    VarActForm = Me.Name
    
    cmbGender.Clear
    Select Case clsStation.Language
    
        Case 0
        
            cmbGender.AddItem "¬ﬁ«"
            cmbGender.ItemData(cmbGender.ListCount - 1) = 1
            cmbGender.AddItem "Œ«‰„"
            cmbGender.ItemData(cmbGender.ListCount - 1) = 0
        
        Case 1
        
            cmbGender.AddItem "Male"
            cmbGender.ItemData(cmbGender.ListCount - 1) = 1
            cmbGender.AddItem "Female"
            cmbGender.ItemData(cmbGender.ListCount - 1) = 0
    
    End Select
    cmbGender.ListIndex = 0
    
    Set rctmp = RunStoredProcedure2RecordSet("Get_All_tPrefix")
    If Not (rctmp.EOF = True And rctmp.BOF = True) Then
        While Not rctmp.EOF
            cmbPrefix.AddItem rctmp!Description
            cmbPrefix.ItemData(cmbPrefix.NewIndex) = rctmp!Code
            rctmp.MoveNext
        Wend
    Else
        cmbPrefix.AddItem " "
        cmbPrefix.ItemData(0) = 0
    End If
    Me.cmbPrefix.ListIndex = 0
    rctmp.Close

    Set rctmp = RunStoredProcedure2RecordSet("Get_All_tWorkType")
    If Not (rctmp.EOF = True And rctmp.BOF = True) Then
        While Not rctmp.EOF
            cmbActKind.AddItem rctmp!Description
            cmbActKind.ItemData(cmbActKind.NewIndex) = rctmp!Code
            rctmp.MoveNext
        Wend
    Else
        cmbActKind.AddItem " "
        cmbActKind.ItemData(0) = 0
    End If
    Me.cmbActKind.ListIndex = 0
    rctmp.Close
    
    Set rctmp = RunStoredProcedure2RecordSet("Get_All_tState")
    If Not (rctmp.EOF = True And rctmp.BOF = True) Then
        While Not rctmp.EOF
            CmbState.AddItem rctmp!Description
            CmbState.ItemData(CmbState.NewIndex) = rctmp!Code
            rctmp.MoveNext
        Wend
    Else
        CmbState.AddItem " "
        CmbState.ItemData(0) = 0
    End If
    Me.CmbState.ListIndex = 0
    rctmp.Close
    
    Set rctmp = RunStoredProcedure2RecordSet("Get_All_tCity")
    If Not (rctmp.EOF = True And rctmp.BOF = True) Then
        While Not rctmp.EOF
            cmbCity.AddItem rctmp!Description
            cmbCity.ItemData(cmbCity.NewIndex) = rctmp!Code
            rctmp.MoveNext
        Wend
    Else
        cmbCity.AddItem " "
        cmbCity.ItemData(0) = 0
    End If
    Me.cmbCity.ListIndex = 0
    rctmp.Close
    
    
    vsCustomer.ColHidden(6) = True
    
    FillBranch
    
    FillAtf
    If (clsArya.ExternalAccounting = True And ClsFormAccess.frmAccount = True) Or HasMiniAcc = True Then
        cmdAddTafsili.Enabled = True
    Else
        cmdAddTafsili.Enabled = False
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

    With vsCustomer
         For i = 0 To .Cols - 1
            .ColWidth(i) = Val(GetSetting(strMainKey, "frmSupplier_vsCustomer", "Col" & i))
            If .ColWidth(i) = 0 Then
                .ColWidth(i) = 1000       '
            End If
         Next i
    End With
    Add

    If ClsFormAccess.ChangeTotalRemainingAmount = True Then
        txtPrimaryBedehi.Enabled = True
        txtPrimaryTalab.Enabled = True
    Else
        txtPrimaryBedehi.Enabled = False
        txtPrimaryTalab.Enabled = False
    End If
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

    Set Rc = Nothing
    Set rctmp = Nothing
    Set cn = Nothing
    If PosConnection.State = adStateOpen Then PosConnection.Close: Set PosConnection = Nothing
    
    Set clsDate = Nothing
    Set mdifrm.FileCls = Nothing
        
    VarActForm = ""
    Unload frmSupplierComp
    Dim Obj As Object
    Dim Exit_Form As Boolean
    For Each Obj In Forms
        If LCase(Obj.Name) = "frmpurchase" Then
            If ClsFormAccess.frmPurchase = True Then
                frmPurchase.Show
                frmPurchase.SetFirstToolBar
                Exit_Form = True
            End If
        End If
    Next Obj
''''
    
    If Exit_Form = False Then
        mdifrm.Toolbar1.Buttons(20).Enabled = False
        mdifrm.Toolbar1.Buttons(21).Enabled = False
        mdifrm.Toolbar1.Buttons(23).Enabled = True
        mdifrm.Toolbar1.Buttons(24).Enabled = True
        mdifrm.Toolbar1.Buttons(25).Enabled = True
        mdifrm.Toolbar1.Buttons(26).Enabled = True
        mdifrm.Toolbar1.Buttons(27).Enabled = True
    End If


    SaveSetting strMainKey, Me.Name, "Left", Me.left
    SaveSetting strMainKey, Me.Name, "Top", Me.top



End Sub

Public Sub BeforeFirstKey()
    If MyFormAddEditMode <> ViewMode Then
        Cancel
    End If
End Sub

Public Sub FirstKey()
    Dim i As Long
    MyFormAddEditMode = ViewMode  'View Mode
    SetFirstToolBar
    
    ReDim Parameter(2) As Parameter
    Parameter(0) = GenerateInputParameter("@CurrentSupplierCode", adInteger, 4, 0)
    Parameter(1) = GenerateInputParameter("@Direction", adInteger, 4, 0)
    Parameter(2) = GenerateInputParameter("@Branch", adInteger, 4, CurrentBranch)
    Set rctmp = RunParametricStoredProcedure2Rec("NavigateInSupplier", Parameter)
    If rctmp.EOF <> True And rctmp.BOF <> True Then
        Me.txtCode.Text = rctmp.Fields("code").Value
        i = vsCustomer.FindRow(txtCode.Text, 1, 1, True, True)
        If i > 0 Then
            vsCustomer.Row = i
            vsCustomer.ShowCell i, 1
        End If
    End If
    rctmp.Close
    
    GetDataDetail
End Sub

Public Sub BeforePreviousKey()
End Sub

Public Sub PreviousKey()
    Dim i As Long
    MyFormAddEditMode = ViewMode  'View Mode
    SetFirstToolBar
    
    ReDim Parameter(2) As Parameter
    Parameter(0) = GenerateInputParameter("@CurrentSupplierCode", adInteger, 4, Val(txtCode.Text))
    Parameter(1) = GenerateInputParameter("@Direction", adInteger, 4, 1)
    Parameter(2) = GenerateInputParameter("@Branch", adInteger, 4, CurrentBranch)
    Set rctmp = RunParametricStoredProcedure2Rec("NavigateInSupplier", Parameter)
    If rctmp.EOF <> True And rctmp.BOF <> True Then
         Me.txtCode.Text = rctmp.Fields("code").Value
        i = vsCustomer.FindRow(txtCode.Text, 1, 1, True, True)
        If i > 0 Then
            vsCustomer.Row = i
            vsCustomer.ShowCell i, 1
        End If
    End If
    rctmp.Close
    
    GetDataDetail
    HeaderLabel Val(MyFormAddEditMode), fwlblMode
End Sub

Public Sub BeforeNextKey()

End Sub

Public Sub NextKey()
    Dim i As Long
    MyFormAddEditMode = ViewMode  'View Mode
    SetFirstToolBar
    
    ReDim Parameter(2) As Parameter
    Parameter(0) = GenerateInputParameter("@CurrentSupplierCode", adInteger, 4, Val(txtCode.Text))
    Parameter(1) = GenerateInputParameter("@Direction", adInteger, 4, 2)
    Parameter(2) = GenerateInputParameter("@Branch", adInteger, 4, CurrentBranch)
    Set rctmp = RunParametricStoredProcedure2Rec("NavigateInSupplier", Parameter)
    If rctmp.EOF <> True And rctmp.BOF <> True Then
         Me.txtCode.Text = rctmp.Fields("code").Value
        i = vsCustomer.FindRow(txtCode.Text, 1, 1, True, True)
        If i > 0 Then
            vsCustomer.Row = i
            vsCustomer.ShowCell i, 1
        End If
    End If
    rctmp.Close
 
    MyFormAddEditMode = ViewMode
    GetDataDetail
    SetFirstToolBar
    HeaderLabel Val(MyFormAddEditMode), Me.fwlblMode
End Sub

Public Sub BeforeLastKey()
    If MyFormAddEditMode <> ViewMode Then
    Cancel
    End If
End Sub
Public Sub Cancel()
    Select Case MyFormAddEditMode
        Case AddMode 'new
            DefaultSettings
            MyFormAddEditMode = AddMode
            SetFirstToolBar
        Case EditMode 'edit
            GetDataDetail
            MyFormAddEditMode = ViewMode
            SetFirstToolBar
    End Select
End Sub

Public Sub LastKey()
    Dim i As Long
    MyFormAddEditMode = ViewMode  'View Mode
    SetFirstToolBar
    
    ReDim Parameter(2) As Parameter
    Parameter(0) = GenerateInputParameter("@CurrentSupplierCode", adInteger, 4, 0)
    Parameter(1) = GenerateInputParameter("@Direction", adInteger, 4, 3)
    Parameter(2) = GenerateInputParameter("@Branch", adInteger, 4, CurrentBranch)
    Set rctmp = RunParametricStoredProcedure2Rec("NavigateInSupplier", Parameter)
    If rctmp.EOF <> True And rctmp.BOF <> True Then
         Me.txtCode.Text = rctmp.Fields("code").Value
        i = vsCustomer.FindRow(txtCode.Text, 1, 1, True, True)
        If i > 0 Then
            vsCustomer.Row = i
            vsCustomer.ShowCell i, 1
            
        End If
    End If
    rctmp.Close
    
    GetDataDetail
End Sub
Public Sub DefaultSettings()

    On Error Resume Next
    
    txtPrimaryBedehi = ""
    txtPrimaryTalab = ""
    cmbActKind.ListIndex = 0
    CmbState.ListIndex = 0
    cmbCity.ListIndex = 0
    cmbPrefix.ListIndex = 0
    cmbGender.ListIndex = 0
    
    On Error GoTo 0
    
    TxtAddress.Text = ""
    txtDescription.Text = ""
    txtDiscount.Text = 0
    txtEmail.Text = ""
    txtFamily.Text = ""
    txtFax.Text = ""
    txtMobile.Text = ""
    TxtName.Text = ""
    txtPostalCode.Text = ""
    txtTel1.Text = ""
    txtTel2.Text = ""
    txtTel3.Text = ""
    txtTel4.Text = ""
    txtWorkName.Text = ""
    txtMembershipId.Text = ""
    
    OptionActDeAct(0).Value = True
    OptionOwner(0).Value = True
    
    txtTafsiliCode.Text = ""
    OldTafsili = 0
    txtEconomicCode = ""
    txtNationalCode = ""
End Sub

Public Sub Add()
On Error GoTo ErrHandler
    If MyFormAddEditMode = EditMode Then
        DefaultSettings
    End If
    MyFormAddEditMode = AddMode
    DefaultSettings

    ReDim Parameter(0) As Parameter
    Parameter(0) = GenerateInputParameter("@Branch", adInteger, 4, CurrentBranch)
    Set rctmp = RunParametricStoredProcedure2Rec("Get_New_Supplier_Code", Parameter)
    txtCode.Text = rctmp.Fields("Code").Value
    txtMembershipId.Text = rctmp.Fields("MembershipId").Value
    If OptionOwner(0).Value Then
        Me.OptionOwnerValue 0
    Else
        Me.OptionOwnerValue 1
    End If
    
    SetFirstToolBar
    FillvsCustomer
    Exit Sub
ErrHandler:
    MsgBox err.Description
    modgl.LogSave "frmSupplier", err, "Add"
End Sub

Public Sub ExitSub()
If PosConnection.State = adStateOpen Then PosConnection.Close: Set PosConnection = Nothing
    Unload Me
End Sub

Public Sub Update()
On Error GoTo ErrHandler
    If MyFormAddEditMode = ViewMode Then Exit Sub
    Dim strBinBuyState As String
    Dim intBuyState As Integer
    
    If Val(txtDiscount.Text) < 0 Or Val(txtDiscount.Text) > 100 Then
        frmMsg.fwlblMsg.Caption = "„ﬁœ«—  Œ›Ì› ‰„Ì  Ê«‰œ ò„ — «“ ’›— Ì« »Ì‘ — «“ ’œ œ—’œ »«‘œ "
        frmMsg.fwBtn(0).Visible = False
        frmMsg.fwBtn(1).ButtonType = flwButtonOk
        frmMsg.fwBtn(1).Caption = "ﬁ»Ê·"
        frmMsg.Show vbModal
        Exit Sub
    End If
    
    If framePerson.Visible = True Then
        If Trim(txtMembershipId.Text) = "" Or Trim(txtFamily.Text) = "" Or Trim(TxtAddress.Text) = "" Or Trim(txtMembershipId.Text) = "" Then
            frmMsg.fwlblMsg.Caption = "·ÿ›« «ÿ·«⁄«  —« »Â ÿÊ— ò«„· Ê«—œ ‰„«ÌÌœ"
            frmMsg.fwBtn(0).Visible = False
            frmMsg.fwBtn(1).ButtonType = flwButtonOk
            frmMsg.fwBtn(1).Caption = "ﬁ»Ê·"
            frmMsg.Show vbModal
            Exit Sub
        End If
    ElseIf frameCompany.Visible = True Then
        If Trim(txtWorkName.Text) = "" Or Trim(TxtAddress.Text) = "" Or Trim(txtMembershipId.Text) = "" Then
            frmMsg.fwlblMsg.Caption = "·ÿ›« «ÿ·«⁄«  —« »Â ÿÊ— ò«„· Ê«—œ ‰„«ÌÌœ"
            frmMsg.fwBtn(0).Visible = False
            frmMsg.fwBtn(1).ButtonType = flwButtonOk
            frmMsg.fwBtn(1).Caption = "ﬁ»Ê·"
            frmMsg.Show vbModal
            Exit Sub
        End If
    End If
     
    If OptionOwner(0).Value = True Then
        txtWorkName.Text = ""
        cmbActKind.ListIndex = 0
        
    Else
        TxtName.Text = ""
        txtFamily.Text = ""
        cmbGender.ListIndex = 0
        cmbPrefix.ListIndex = 0
    
    End If
    
    
    Select Case MyFormAddEditMode
        Case AddMode
        
            ReDim Parameter(30) As Parameter
            Parameter(0) = GenerateInputParameter("@MembershipId", adVarWChar, 50, Val(txtMembershipId.Text))
            Parameter(1) = GenerateInputParameter("@MasterCode", adInteger, 4, 0)
            Parameter(2) = GenerateInputParameter("@Owner", adInteger, 4, IIf(OptionOwner(0).Value = True, 0, 1))
            Parameter(3) = GenerateInputParameter("@Name", adVarWChar, 50, Trim(TxtName.Text))
            Parameter(4) = GenerateInputParameter("@Family", adVarWChar, 50, Trim(txtFamily.Text))
            Parameter(5) = GenerateInputParameter("@Sex", adInteger, 4, cmbGender.ItemData(cmbGender.ListIndex))
            Parameter(6) = GenerateInputParameter("@WorkName", adVarWChar, 50, txtWorkName.Text)
            Parameter(7) = GenerateInputParameter("@InternalNo", adVarWChar, 50, "")
            Parameter(8) = GenerateInputParameter("@Unit", adVarWChar, 50, "")
            Parameter(9) = GenerateInputParameter("@State", adInteger, 4, CmbState.ItemData(CmbState.ListIndex))
            Parameter(10) = GenerateInputParameter("@City", adInteger, 4, cmbCity.ItemData(cmbCity.ListIndex))
            Parameter(11) = GenerateInputParameter("@ActKind", adInteger, 4, cmbActKind.ItemData(cmbActKind.ListIndex))
            Parameter(12) = GenerateInputParameter("@ActDeAct", adInteger, 4, IIf(OptionActDeAct(0).Value = True, 1, 0))
            Parameter(13) = GenerateInputParameter("@Prefix", adInteger, 4, cmbPrefix.ItemData(cmbPrefix.ListIndex))
            Parameter(14) = GenerateInputParameter("@Address", adVarWChar, 255, Trim(TxtAddress.Text))
            Parameter(15) = GenerateInputParameter("@PostalCode", adVarWChar, 50, Trim(txtPostalCode.Text))
            Parameter(16) = GenerateInputParameter("@Tel1", adVarWChar, 50, Trim(txtTel1.Text))
            Parameter(17) = GenerateInputParameter("@Tel2", adVarWChar, 50, Trim(txtTel2.Text))
            Parameter(18) = GenerateInputParameter("@Tel3", adVarWChar, 50, Trim(txtTel3.Text))
            Parameter(19) = GenerateInputParameter("@Tel4", adVarWChar, 50, Trim(txtTel4.Text))
            Parameter(20) = GenerateInputParameter("@Mobile", adVarWChar, 50, Trim(txtMobile.Text))
            Parameter(21) = GenerateInputParameter("@Fax", adVarWChar, 50, Trim(txtFax.Text))
            Parameter(22) = GenerateInputParameter("@Email", adVarWChar, 50, Trim(txtEmail.Text))
            Parameter(23) = GenerateInputParameter("@Flour", adVarWChar, 50, "")
            Parameter(24) = GenerateInputParameter("@Discount", adDouble, 8, Val(txtDiscount.Text))
            Parameter(25) = GenerateInputParameter("@Description", adVarWChar, 255, Trim(txtDescription.Text))
            Parameter(26) = GenerateInputParameter("@User", adInteger, 4, mvarCurUserNo)
            Parameter(27) = GenerateInputParameter("@TotalRemainingAmount", adDouble, 8, IIf(Val(txtPrimaryBedehi) > 0, Val(txtPrimaryBedehi), -1 * Val(txtPrimaryTalab)))
            Parameter(28) = GenerateInputParameter("@EconomicCode", adVarWChar, 20, Trim(txtEconomicCode.Text))
            Parameter(29) = GenerateInputParameter("@NationalCode", adVarWChar, 20, Trim(txtNationalCode.Text))
            Parameter(30) = GenerateOutputParameter("@Code", adBigInt, 8)
            
            Dim LastCode As Long
            LastCode = RunParametricStoredProcedure("Insert_Supplier", Parameter)
            If LastCode <> -1 Then
                frmMsg.fwlblMsg.Caption = "À»   «„Ì‰ ò‰‰œÂ ÃœÌœ »« „Ê›ﬁÌ  «‰Ã«„ ‘œ"
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
                txtMembershipId.SetFocus
                Exit Sub
            End If
            
            
        Case EditMode
        
            ReDim Parameter(31) As Parameter
            Parameter(0) = GenerateInputParameter("@MembershipId", adVarWChar, 50, Val(txtMembershipId.Text))
            Parameter(1) = GenerateInputParameter("@MasterCode", adInteger, 4, 0)
            Parameter(2) = GenerateInputParameter("@Owner", adInteger, 4, IIf(OptionOwner(0).Value = True, 0, 1))
            Parameter(3) = GenerateInputParameter("@Name", adVarWChar, 50, Trim(TxtName.Text))
            Parameter(4) = GenerateInputParameter("@Family", adVarWChar, 50, Trim(txtFamily.Text))
            Parameter(5) = GenerateInputParameter("@Sex", adInteger, 4, cmbGender.ItemData(cmbGender.ListIndex))
            Parameter(6) = GenerateInputParameter("@WorkName", adVarWChar, 50, txtWorkName.Text)
            Parameter(7) = GenerateInputParameter("@InternalNo", adVarWChar, 50, "")
            Parameter(8) = GenerateInputParameter("@Unit", adVarWChar, 50, "")
            Parameter(9) = GenerateInputParameter("@State", adInteger, 4, CmbState.ItemData(CmbState.ListIndex))
            Parameter(10) = GenerateInputParameter("@City", adInteger, 4, cmbCity.ItemData(cmbCity.ListIndex))
            Parameter(11) = GenerateInputParameter("@ActKind", adInteger, 4, cmbActKind.ItemData(cmbActKind.ListIndex))
            Parameter(12) = GenerateInputParameter("@ActDeAct", adInteger, 4, IIf(OptionActDeAct(0).Value = True, 1, 0))
            Parameter(13) = GenerateInputParameter("@Prefix", adInteger, 4, cmbPrefix.ItemData(cmbPrefix.ListIndex))
            Parameter(14) = GenerateInputParameter("@Address", adVarWChar, 255, Trim(TxtAddress.Text))
            Parameter(15) = GenerateInputParameter("@PostalCode", adVarWChar, 50, Trim(txtPostalCode.Text))
            Parameter(16) = GenerateInputParameter("@Tel1", adVarWChar, 50, Trim(txtTel1.Text))
            Parameter(17) = GenerateInputParameter("@Tel2", adVarWChar, 50, Trim(txtTel2.Text))
            Parameter(18) = GenerateInputParameter("@Tel3", adVarWChar, 50, Trim(txtTel3.Text))
            Parameter(19) = GenerateInputParameter("@Tel4", adVarWChar, 50, Trim(txtTel4.Text))
            Parameter(20) = GenerateInputParameter("@Mobile", adVarWChar, 50, Trim(txtMobile.Text))
            Parameter(21) = GenerateInputParameter("@Fax", adVarWChar, 50, Trim(txtFax.Text))
            Parameter(22) = GenerateInputParameter("@Email", adVarWChar, 50, Trim(txtEmail.Text))
            Parameter(23) = GenerateInputParameter("@Flour", adVarWChar, 50, "")
            Parameter(24) = GenerateInputParameter("@Discount", adDouble, 8, Val(txtDiscount.Text))
            Parameter(25) = GenerateInputParameter("@Description", adVarWChar, 255, Trim(txtDescription.Text))
            Parameter(26) = GenerateInputParameter("@User", adInteger, 4, mvarCurUserNo)
            Parameter(27) = GenerateInputParameter("@Code", adBigInt, 8, Val(txtCode.Text))
            Parameter(28) = GenerateInputParameter("@TotalRemainingAmount", adDouble, 8, IIf(Val(txtPrimaryBedehi) > 0, Val(txtPrimaryBedehi), -1 * Val(txtPrimaryTalab)))
            Parameter(29) = GenerateInputParameter("@EconomicCode", adVarWChar, 20, Trim(txtEconomicCode.Text))
            Parameter(30) = GenerateInputParameter("@NationalCode", adVarWChar, 20, Trim(txtNationalCode.Text))
            Parameter(31) = GenerateOutputParameter("@Updated", adBigInt, 8)
            Dim Updated As Long
            Updated = RunParametricStoredProcedure("Update_Supplier", Parameter)
            If Updated > 0 Then
                frmMsg.fwlblMsg.Caption = " €ÌÌ—«  »« „Ê›ﬁÌ  «‰Ã«„ ‘œ"
                frmMsg.fwBtn(0).Visible = False
                frmMsg.fwBtn(1).ButtonType = flwButtonOk
                frmMsg.fwBtn(1).Caption = "ﬁ»Ê·"
                frmMsg.Show vbModal
                If clsArya.ExternalAccounting = True Or HasMiniAcc = True Then
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
ErrHandler:
    MsgBox err.Description
    modgl.LogSave "frmSupplier", err, "Update"

End Sub


Public Sub Edit()

    If OptionOwner(0).Value Then
        Me.OptionOwnerValue 0
    Else
        Me.OptionOwnerValue 1
    End If
    
    MyFormAddEditMode = EditMode
    SetFirstToolBar
    
End Sub
Public Sub Find()

        frmFindSupplier.Show vbModal
        
        If mvarcode <> 0 Then
'            txtCode.Text = vsCustomer.TextMatrix(vsCustomer.Row, 1)
            txtCode.Text = mvarcode
            MyFormAddEditMode = ViewMode
            GetDataDetail
            SetFirstToolBar
            HeaderLabel Val(MyFormAddEditMode), Me.fwlblMode
            
            mvarcode = 0
        End If
End Sub

Public Sub ExitForm()
    Unload Me
End Sub



Private Sub FWBtnPerson_Click()

    If MyFormAddEditMode <> AddMode And OptionOwner(1).Value = True Then
    
        mvarcode = Val(txtCode.Text)
        frmSupplierComp.Show
        
    ElseIf MyFormAddEditMode = AddMode Then
    
        frmMsg.fwlblMsg.Caption = "«» œ«  «„Ì‰ ò‰‰œÂ ›Êﬁ —« À»  Ê ”Å” Ê«—œ „—Õ·Â ÊÌ“Ì Ê—Â« ‘ÊÌœ"
        frmMsg.fwBtn(0).Visible = False
        frmMsg.fwBtn(1).ButtonType = flwButtonOk
        frmMsg.fwBtn(1).Caption = "ﬁ»Ê·"
        frmMsg.Show vbModal
        
    ElseIf OptionOwner(1).Value = False Then
    
        frmMsg.fwlblMsg.Caption = "»Â  «„Ì‰ ò‰‰œÂ ›Êﬁ ‰„Ì  Ê«‰Ìœ ÊÌ“Ì Ê— «÷«›Â ‰„«ÌÌœ"
        frmMsg.fwBtn(0).Visible = False
        frmMsg.fwBtn(1).ButtonType = flwButtonOk
        frmMsg.fwBtn(1).Caption = "ﬁ»Ê·"
        frmMsg.Show vbModal
        
    End If

End Sub


Private Sub OptionOwner_Click(index As Integer)

    Me.OptionOwnerValue index
    
End Sub

Private Sub ResizeKit1_ExitResize(ByVal XScale As Double, ByVal YScale As Double)


    If formloadFlag = True Then
        SaveSetting strMainKey, Me.Name, "Height", Me.Height
        SaveSetting strMainKey, Me.Name, "Width", Me.Width
    End If

End Sub


Private Sub txtDescription_GotFocus()
    txtDescription = ""
End Sub

Private Sub txtDiscount_Change()
    If Val(txtDiscount.Text) > 100 Then
        frmMsg.fwlblMsg.Caption = "„Ì“«‰  Œ›Ì› »«Ìœ ﬂ„ — «“ 100 »«‘œ"
        frmMsg.fwBtn(0).Visible = False
        frmMsg.fwBtn(1).ButtonType = flwButtonOk
        frmMsg.fwBtn(1).Caption = "ﬁ»Ê·"
        frmMsg.Show vbModal
    End If
End Sub

Public Sub OptionOwnerValue(index As Integer)
    Select Case index
        Case 0:
             framePerson.Visible = True
             frameCompany.Visible = False
        Case 1:
             framePerson.Visible = False
             frameCompany.Visible = True
    End Select
    
End Sub

Sub SetFirstToolBar()
    
    Dim Obj As Object
    
    AllButton vbOff, True
    
    mdifrm.Toolbar1.Buttons(1).Enabled = True   'Home
    mdifrm.Toolbar1.Buttons(2).Enabled = True   'PageUp
    mdifrm.Toolbar1.Buttons(3).Enabled = True   'PageDown
    mdifrm.Toolbar1.Buttons(4).Enabled = True   'End
    mdifrm.Toolbar1.Buttons(13).Enabled = True   'Find
    
    mdifrm.Toolbar1.Buttons(15).Enabled = True   'Print
    
    mdifrm.Toolbar1.Buttons(23).Enabled = True
    mdifrm.Toolbar1.Buttons(24).Enabled = True
    mdifrm.Toolbar1.Buttons(25).Enabled = True
    mdifrm.Toolbar1.Buttons(26).Enabled = True
    mdifrm.Toolbar1.Buttons(27).Enabled = True
    
 
    If MyFormAddEditMode = ViewMode Then  ' View Mode
 
        On Error Resume Next
        For Each Obj In Me
            If TypeOf Obj Is OptionButton Or TypeOf Obj Is FWCheck Then
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
        txtTafsiliCode.Enabled = False
    ElseIf MyFormAddEditMode = AddMode Then    'Add Mode
        
        On Error Resume Next
        For Each Obj In Me
            If TypeOf Obj Is OptionButton Or TypeOf Obj Is FWCheck Then
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
        
    ElseIf MyFormAddEditMode = EditMode Then     'Edit
    
        On Error Resume Next
        For Each Obj In Me
            If TypeOf Obj Is OptionButton Or TypeOf Obj Is FWCheck Then
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
        
    End If
    
    HeaderLabel Val(MyFormAddEditMode), fwlblMode
    
End Sub
Sub GetDataDetail()
    On Error GoTo ErrHandler
    Dim rctmp As New ADODB.Recordset
    DefaultSettings
    
    Dim TempStr As String
    
    ReDim Parameter(1) As Parameter
    Parameter(0) = GenerateInputParameter("@Code", adInteger, 4, Val(txtCode.Text))
    Parameter(1) = GenerateInputParameter("@Branch", adInteger, 4, CurrentBranch)
    Set rctmp = RunParametricStoredProcedure2Rec("Get_Supplier_info", Parameter)
    Dim ii As Integer
    If Not (rctmp.BOF Or rctmp.EOF) Then
        If rctmp!Owner = 0 Then
            OptionOwner(0).Value = True
        Else
            OptionOwner(1).Value = True
        End If
        
        If rctmp!ActDeAct = True Then
            OptionActDeAct(0).Value = True
        Else
            OptionActDeAct(1).Value = True
        End If
        
        txtCode = rctmp!Code
        txtMembershipId = rctmp!MembershipId
        TxtName = rctmp!Name
        txtFamily = rctmp!Family
        txtTel1 = rctmp!Tel1
        txtTel2 = rctmp!Tel2
        txtTel3 = rctmp!Tel3
        txtTel4 = rctmp!Tel4
        txtWorkName = rctmp!WorkName
        txtFax = rctmp!Fax
        txtMobile = rctmp!Mobile
        txtDiscount = rctmp!Discount
        txtEmail = rctmp!Email
        txtDescription = rctmp!Description
        TxtAddress = rctmp!address
        txtTafsiliCode.Text = IIf(IsNull(rctmp!Tafsili), "", rctmp!Tafsili)
        OldTafsili = Val(txtTafsiliCode.Text)
        txtSanadNo.Text = IIf(IsNull(rctmp!SanadNo), "", rctmp!SanadNo)
        
        If IsNull(rctmp!TotalRemainingAmount) = False Then
            If Val(rctmp!TotalRemainingAmount) > 0 Then
                txtPrimaryBedehi.Text = Val(rctmp!TotalRemainingAmount)
            Else
                txtPrimaryTalab.Text = -1 * Val(rctmp!TotalRemainingAmount)
            End If
        End If
        For i = 0 To cmbActKind.ListCount - 1
            If cmbActKind.ItemData(i) = rctmp!ActKind Then
                cmbActKind.ListIndex = i
                Exit For
            End If
        Next i
        
        For i = 0 To CmbState.ListCount - 1
            If CmbState.ItemData(i) = rctmp!State Then
                CmbState.ListIndex = i
                Exit For
            End If
        Next i
        
        For i = 0 To cmbCity.ListCount - 1
            If cmbCity.ItemData(i) = rctmp!City Then
                cmbCity.ListIndex = i
                Exit For
            End If
        Next i
        
        
        For i = 0 To cmbPrefix.ListCount - 1
            If cmbPrefix.ItemData(i) = rctmp!Prefix Then
                cmbPrefix.ListIndex = i
                Exit For
            End If
        Next i
        
        For i = 0 To cmbGender.ListCount - 1
            If cmbGender.ItemData(i) = rctmp!Sex Then
                cmbGender.ListIndex = i
                Exit For
            End If
        Next i
        txtPostalCode.Text = IIf(IsNull(rctmp!PostalCode), "", rctmp!PostalCode)
        txtEconomicCode.Text = IIf(IsNull(rctmp!EconomicCode), "", rctmp!EconomicCode)
        txtNationalCode.Text = IIf(IsNull(rctmp!NationalCode), "", rctmp!NationalCode)
        
               
    End If
    rctmp.Close
    Exit Sub
ErrHandler:
    MsgBox err.Description
    modgl.LogSave "frmSupplier", err, "GetDataDetail"
End Sub



Private Sub vsCustomer_AfterSort(ByVal Col As Long, Order As Integer)
    With vsCustomer
        If Col = 3 And .Rows > 1 Then
            For i = 1 To .Rows - 2
                If (Val(.TextMatrix(i, 3)) - Val(.TextMatrix(i + 1, 3)) > 1 And Order = 2) Or (Val(.TextMatrix(i + 1, 3)) - Val(.TextMatrix(i, 3)) > 1 And Order = 1) Then
                    .Cell(flexcpBackColor, i, .FixedCols, i, .Cols - 1) = 8421631
                Else
                    .Cell(flexcpBackColor, i, .FixedCols, i, .Cols - 1) = &H80000005
                End If
            Next i
        End If
    End With
End Sub

Private Sub vsCustomer_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    If Col = -1 Then Exit Sub
    For i = 0 To vsCustomer.Cols - 1
        SaveSetting strMainKey, "frmSupplier_vsCustomer", "Col" & i, vsCustomer.ColWidth(i)
    Next

End Sub

Private Sub vsCustomer_Click()
    
    txtCode.Text = vsCustomer.TextMatrix(vsCustomer.Row, 1)
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
        Accounting.Insert_PrimarySand_Cust CustCode, Val(txtTafsiliCode.Text), Val(txtPrimaryBedehi), Val(txtPrimaryTalab), 1, 1
            
    End If
    If Val(txtTafsiliCode.Text) > 0 Then
        ReDim Parameter(1) As Parameter
        Parameter(0) = GenerateInputParameter("@TafsiliId", adInteger, 4, Val(txtTafsiliCode.Text))
        Parameter(1) = GenerateInputParameter("@Code", adInteger, 4, CustCode)
        RunParametricStoredProcedure "Update_tsupplier_tafsili", Parameter
    End If
    If ShowMessageflag = False Then Exit Sub
Exit Sub
ErrHandler:
    MsgBox err.Description
    modgl.LogSave "frmSupplier", err, "Insert_Tafsili"
    Resume Next
End Sub

Private Sub FillAtf()
    txtAtf.Text = "«‘Œ«’ Ê ‘—ò Â«"
End Sub

Private Sub vsCustomer_RowColChange()
    txtCode.Text = vsCustomer.TextMatrix(vsCustomer.Row, 1)
    MyFormAddEditMode = ViewMode
    GetDataDetail

End Sub
