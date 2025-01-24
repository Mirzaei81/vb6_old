VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{7AEDC602-D94C-11D1-BB7A-00E0290EA3C9}#1.0#0"; "ResizeKit.ocx"
Object = "{75D4F767-8785-11D3-93AD-0000832EF44D}#2.14#0"; "FAST2001.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmCust 
   ClientHeight    =   10275
   ClientLeft      =   5055
   ClientTop       =   450
   ClientWidth     =   14550
   Icon            =   "frmCust.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   10275
   ScaleMode       =   0  'User
   ScaleWidth      =   14550
   Begin VB.Frame Frame21 
      Caption         =   "Frame2"
      Height          =   2295
      Left            =   360
      TabIndex        =   108
      Top             =   600
      Visible         =   0   'False
      Width           =   3255
      Begin VB.ComboBox Bcount 
         Height          =   315
         Left            =   1560
         TabIndex        =   119
         Text            =   "1"
         Top             =   240
         Visible         =   0   'False
         Width           =   615
      End
      Begin VB.TextBox blockNtxt 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   480
         TabIndex        =   118
         Text            =   "4"
         Top             =   240
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CheckBox KeyAorB 
         BackColor       =   &H00FF8080&
         Caption         =   "KEY A"
         Height          =   255
         Left            =   840
         TabIndex        =   117
         Top             =   240
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox keyTXT 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   0
         Left            =   -240
         TabIndex        =   116
         Text            =   "FF"
         Top             =   1920
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.TextBox keyTXT 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   1
         Left            =   0
         TabIndex        =   115
         Text            =   "FF"
         Top             =   1920
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.TextBox keyTXT 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   2
         Left            =   240
         TabIndex        =   114
         Text            =   "FF"
         Top             =   1920
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.TextBox keyTXT 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   3
         Left            =   480
         TabIndex        =   113
         Text            =   "FF"
         Top             =   1920
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.TextBox keyTXT 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   4
         Left            =   720
         TabIndex        =   112
         Text            =   "FF"
         Top             =   1920
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.TextBox keyTXT 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   5
         Left            =   960
         TabIndex        =   111
         Text            =   "FF"
         Top             =   1920
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.TextBox BufferTXT 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   720
         TabIndex        =   110
         Top             =   1440
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Timer TimerRFID 
         Interval        =   1000
         Left            =   0
         Top             =   600
      End
      Begin VB.TextBox ResultTXT 
         Appearance      =   0  'Flat
         Height          =   255
         Left            =   -840
         Locked          =   -1  'True
         TabIndex        =   109
         Top             =   1920
         Visible         =   0   'False
         Width           =   735
      End
   End
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
      Height          =   4335
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   83
      Top             =   360
      Width           =   1455
      Begin VB.Frame Frame7 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   102
         Top             =   2280
         Width           =   1215
         Begin VB.TextBox TxtFamilyNo 
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
            ForeColor       =   &H00000080&
            Height          =   405
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   103
            ToolTipText     =   "„Ì“«‰ œ—’œ »«Ìœ ﬂ„ — «“ 100 »«‘œ"
            Top             =   360
            Width           =   435
         End
         Begin VB.Label Label6 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "‰›— "
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
            Left            =   0
            RightToLeft     =   -1  'True
            TabIndex        =   105
            Top             =   840
            Width           =   315
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   " ﬂ›·"
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
            Left            =   720
            RightToLeft     =   -1  'True
            TabIndex        =   104
            Top             =   360
            Width           =   435
         End
      End
      Begin VB.PictureBox Picture1 
         Height          =   855
         Left            =   0
         RightToLeft     =   -1  'True
         ScaleHeight     =   795
         ScaleWidth      =   1305
         TabIndex        =   85
         Top             =   1200
         Width           =   1365
         Begin VB.CheckBox ChkMember 
            Alignment       =   1  'Right Justify
            Caption         =   "⁄÷Ê"
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
            Height          =   375
            Left            =   240
            RightToLeft     =   -1  'True
            TabIndex        =   87
            Top             =   0
            Value           =   1  'Checked
            Width           =   855
         End
         Begin VB.CheckBox ChkCentral 
            Alignment       =   1  'Right Justify
            Caption         =   "‘Â—” «‰Ì"
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
            Height          =   375
            Left            =   0
            RightToLeft     =   -1  'True
            TabIndex        =   86
            Top             =   360
            Value           =   1  'Checked
            Width           =   1095
         End
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
         Left            =   30
         RightToLeft     =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   84
         Top             =   480
         Width           =   1365
      End
      Begin FLWCtrls.FWProgressBar FWProgressBar1 
         Height          =   375
         Left            =   0
         Top             =   3840
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   661
         BorderStyle     =   10
      End
   End
   Begin VB.Frame frameAccounting 
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
      Height          =   2475
      Left            =   1560
      RightToLeft     =   -1  'True
      TabIndex        =   77
      Top             =   2205
      Width           =   5055
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
         Height          =   1455
         Left            =   3600
         RightToLeft     =   -1  'True
         TabIndex        =   95
         Top             =   840
         Width           =   1335
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
         TabIndex        =   91
         ToolTipText     =   "„»·€ »œÂÌ »œÊ‰ ⁄·«„  Ê „»·€ »” «‰ﬂ«—Ì »« ⁄·«„  „‰›Ì Ê«—œ ‘Êœ"
         Top             =   840
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
         TabIndex        =   90
         ToolTipText     =   "„»·€ »œÂÌ »œÊ‰ ⁄·«„  Ê „»·€ »” «‰ﬂ«—Ì »« ⁄·«„  „‰›Ì Ê«—œ ‘Êœ"
         Top             =   1320
         Width           =   1455
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
         TabIndex        =   89
         Top             =   1920
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
         Left            =   2760
         RightToLeft     =   -1  'True
         TabIndex        =   88
         Top             =   360
         Width           =   2175
      End
      Begin VB.TextBox txtTafsiliCode 
         BeginProperty Font 
            Name            =   "Arial"
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
         TabIndex        =   78
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "„«‰œÂ «Ê·ÌÂ - »œÂÌ"
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   525
         Left            =   1920
         TabIndex        =   94
         Top             =   840
         Width           =   1635
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "„«‰œÂ «Ê·ÌÂ-ÿ·»"
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
         TabIndex        =   93
         Top             =   1320
         Width           =   1635
         WordWrap        =   -1  'True
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
         Left            =   2400
         RightToLeft     =   -1  'True
         TabIndex        =   92
         Top             =   1920
         Width           =   1155
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblTafsili 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "ﬂœ  ›÷Ì·Ì"
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
         Left            =   1680
         RightToLeft     =   -1  'True
         TabIndex        =   79
         Top             =   360
         Width           =   915
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Frame Frame4 
      Height          =   5175
      Left            =   6720
      RightToLeft     =   -1  'True
      TabIndex        =   51
      Top             =   360
      Width           =   3615
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
         Left            =   120
         TabIndex        =   61
         Top             =   3045
         Width           =   1900
      End
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
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   60
         Top             =   3510
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
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   59
         Top             =   2115
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
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   58
         Top             =   2580
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
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   57
         Top             =   1650
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
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   56
         Top             =   1170
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
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   55
         Top             =   705
         Width           =   1900
      End
      Begin VB.ComboBox cmbCity 
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
         ItemData        =   "frmCust.frx":A4C2
         Left            =   120
         List            =   "frmCust.frx":A4C4
         RightToLeft     =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   54
         Top             =   4515
         Width           =   1900
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
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   53
         Top             =   240
         Width           =   1900
      End
      Begin VB.ComboBox CmbState 
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
         ItemData        =   "frmCust.frx":A4C6
         Left            =   120
         List            =   "frmCust.frx":A4C8
         RightToLeft     =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   52
         Top             =   3990
         Width           =   1900
      End
      Begin VB.Label lblEmail 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Å”  «·ﬂ —Ê‰ÌﬂÌ"
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   405
         Left            =   2055
         RightToLeft     =   -1  'True
         TabIndex        =   71
         Top             =   3090
         Width           =   1395
      End
      Begin VB.Label lblPostalCode 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "ﬂœ Å” Ì"
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   405
         Left            =   2055
         RightToLeft     =   -1  'True
         TabIndex        =   70
         Top             =   3570
         Width           =   1395
      End
      Begin VB.Label lblTel2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   " ·›‰2"
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   405
         Left            =   2055
         RightToLeft     =   -1  'True
         TabIndex        =   69
         Top             =   720
         Width           =   1395
      End
      Begin VB.Label lblTel3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   " ·›‰3"
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   405
         Left            =   2055
         RightToLeft     =   -1  'True
         TabIndex        =   68
         Top             =   1185
         Width           =   1395
      End
      Begin VB.Label lblTel4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   " ·›‰4"
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   405
         Left            =   2055
         RightToLeft     =   -1  'True
         TabIndex        =   67
         Top             =   1665
         Width           =   1395
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
         ForeColor       =   &H80000002&
         Height          =   405
         Left            =   2055
         RightToLeft     =   -1  'True
         TabIndex        =   66
         Top             =   2145
         Width           =   1395
      End
      Begin VB.Label lblFax 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "›«ﬂ”"
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   405
         Left            =   2055
         RightToLeft     =   -1  'True
         TabIndex        =   65
         Top             =   2610
         Width           =   1395
      End
      Begin VB.Label lblCity 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "‘Â—"
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   405
         Left            =   2055
         RightToLeft     =   -1  'True
         TabIndex        =   64
         Top             =   4515
         Width           =   1395
      End
      Begin VB.Label lblTel1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   " ·›‰1"
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
         Left            =   2055
         RightToLeft     =   -1  'True
         TabIndex        =   63
         Top             =   240
         Width           =   1395
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "«” «‰"
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   405
         Left            =   2055
         RightToLeft     =   -1  'True
         TabIndex        =   62
         Top             =   4035
         Width           =   1395
      End
   End
   Begin VB.Frame frameDescription 
      Caption         =   " Ê÷ÌÕ« "
      BeginProperty Font 
         Name            =   "B Homa"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1095
      Left            =   6720
      RightToLeft     =   -1  'True
      TabIndex        =   49
      Top             =   5520
      Width           =   3615
      Begin VB.TextBox txtDescription 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Nazanin"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   660
         Left            =   120
         MultiLine       =   -1  'True
         RightToLeft     =   -1  'True
         TabIndex        =   50
         Top             =   360
         Width           =   3345
      End
   End
   Begin VB.Frame Frame3 
      Height          =   1800
      Left            =   1560
      RightToLeft     =   -1  'True
      TabIndex        =   41
      Top             =   360
      Width           =   5055
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
         ItemData        =   "frmCust.frx":A4CA
         Left            =   840
         List            =   "frmCust.frx":A4CC
         RightToLeft     =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   47
         Top             =   1200
         Width           =   1935
      End
      Begin VB.TextBox txtDiscount 
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
         ForeColor       =   &H00000080&
         Height          =   405
         Left            =   2160
         RightToLeft     =   -1  'True
         TabIndex        =   43
         ToolTipText     =   "„Ì“«‰ œ—’œ »«Ìœ ﬂ„ — «“ 100 »«‘œ"
         Top             =   720
         Width           =   615
      End
      Begin VB.TextBox txtCredit 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   405
         Left            =   840
         RightToLeft     =   -1  'True
         TabIndex        =   42
         ToolTipText     =   "«⁄ »«— »“—ê —«“’›— »—«Ì „‘ —Ì«‰ œ«—«Ì Õ”«» œ› —Ì «” ›«œÂ „Ì ê—œœ"
         Top             =   240
         Width           =   1935
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
         Left            =   3660
         RightToLeft     =   -1  'True
         TabIndex        =   48
         Top             =   1200
         Width           =   1125
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
         Left            =   1080
         RightToLeft     =   -1  'True
         TabIndex        =   46
         Top             =   720
         Width           =   705
      End
      Begin VB.Label lblDiscount 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "„Ì“«‰  Œ›Ì›"
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
         Left            =   3660
         RightToLeft     =   -1  'True
         TabIndex        =   45
         Top             =   720
         Width           =   1125
      End
      Begin VB.Label lblCredit 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "„Ì“«‰ «⁄ »«—"
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
         Left            =   3660
         RightToLeft     =   -1  'True
         TabIndex        =   44
         Top             =   240
         Width           =   1125
      End
   End
   Begin VB.PictureBox frameOwner 
      Height          =   495
      Left            =   10440
      RightToLeft     =   -1  'True
      ScaleHeight     =   435
      ScaleWidth      =   1905
      TabIndex        =   38
      Top             =   480
      Width           =   1960
      Begin VB.OptionButton OptionOwner 
         Alignment       =   1  'Right Justify
         Caption         =   "„Õ· ﬂ«—"
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   375
         Index           =   1
         Left            =   40
         RightToLeft     =   -1  'True
         TabIndex        =   40
         Top             =   80
         Width           =   975
      End
      Begin VB.OptionButton OptionOwner 
         Alignment       =   1  'Right Justify
         Caption         =   "„‰“·"
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   450
         Index           =   0
         Left            =   1050
         RightToLeft     =   -1  'True
         TabIndex        =   39
         Top             =   50
         Width           =   765
      End
   End
   Begin VB.PictureBox Frame2 
      Height          =   495
      Left            =   12510
      RightToLeft     =   -1  'True
      ScaleHeight     =   435
      ScaleWidth      =   1875
      TabIndex        =   35
      Top             =   480
      Width           =   1935
      Begin VB.OptionButton OptionActDeAct 
         Alignment       =   1  'Right Justify
         Caption         =   "»«ÿ·"
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   270
         Index           =   1
         Left            =   960
         RightToLeft     =   -1  'True
         TabIndex        =   37
         Top             =   120
         Width           =   765
      End
      Begin VB.OptionButton OptionActDeAct 
         Alignment       =   1  'Right Justify
         Caption         =   "›⁄«·"
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   375
         Index           =   0
         Left            =   0
         RightToLeft     =   -1  'True
         TabIndex        =   36
         Top             =   75
         Width           =   855
      End
   End
   Begin VB.Frame frameBuyState 
      Caption         =   "Õ«· Â«Ì Œ—Ìœ"
      BeginProperty Font 
         Name            =   "B Homa"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   1320
      Left            =   1560
      RightToLeft     =   -1  'True
      TabIndex        =   30
      Top             =   840
      Visible         =   0   'False
      Width           =   2625
      Begin FLWCtrls.FWCheck FWCheckBuy 
         Height          =   375
         Index           =   0
         Left            =   1680
         TabIndex        =   31
         Top             =   360
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   661
         Caption         =   "Õ÷Ê—Ì"
         Color           =   255
         ForeColor       =   -2147483646
         FontName        =   "B Homa"
         FontBold        =   -1  'True
         FontSize        =   9.75
         Alignment       =   1
      End
      Begin FLWCtrls.FWCheck FWCheckBuy 
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   32
         Top             =   360
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         Caption         =   "Õ÷Ê—Ì «—”«·Ì"
         Color           =   255
         ForeColor       =   -2147483646
         FontName        =   "B Homa"
         FontBold        =   -1  'True
         FontSize        =   9.75
         Alignment       =   1
      End
      Begin FLWCtrls.FWCheck FWCheckBuy 
         Height          =   375
         Index           =   2
         Left            =   1680
         TabIndex        =   33
         Top             =   840
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   661
         Caption         =   " ·›‰Ì"
         Color           =   255
         ForeColor       =   -2147483646
         FontName        =   "B Homa"
         FontBold        =   -1  'True
         FontSize        =   9.75
         Alignment       =   1
      End
      Begin FLWCtrls.FWCheck FWCheckBuy 
         Height          =   375
         Index           =   3
         Left            =   0
         TabIndex        =   34
         Top             =   840
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   661
         Caption         =   " ·›‰Ì «—”«·Ì"
         Color           =   255
         ForeColor       =   -2147483646
         FontName        =   "B Homa"
         FontBold        =   -1  'True
         FontSize        =   9.75
         Alignment       =   1
      End
   End
   Begin VSFlex7LCtl.VSFlexGrid VsCustomer2 
      Height          =   3915
      Left            =   120
      TabIndex        =   28
      Top             =   7320
      Visible         =   0   'False
      Width           =   14355
      _cx             =   25321
      _cy             =   6906
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483643
      ForeColor       =   128
      BackColorFixed  =   8454143
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   16761087
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
      Rows            =   1
      Cols            =   14
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   500
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmCust.frx":A4CE
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
   Begin VB.Frame frameDistance 
      Height          =   2025
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   21
      Top             =   4680
      Width           =   6495
      Begin VB.TextBox TxtRfid 
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
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   97
         Top             =   840
         Width           =   1900
      End
      Begin VB.TextBox TxtEconomicalCode 
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
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   96
         Top             =   1425
         Width           =   1900
      End
      Begin VB.TextBox txtCarryFee 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "B Nazanin"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   3240
         RightToLeft     =   -1  'True
         TabIndex        =   80
         Top             =   825
         Width           =   1900
      End
      Begin VB.CheckBox ChkDistance 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   6120
         RightToLeft     =   -1  'True
         TabIndex        =   27
         Top             =   200
         Width           =   255
      End
      Begin VB.TextBox txtPaykFee 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "B Nazanin"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   3240
         RightToLeft     =   -1  'True
         TabIndex        =   26
         Top             =   1440
         Width           =   1900
      End
      Begin VB.ComboBox cmbDistance 
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
         ItemData        =   "frmCust.frx":A64B
         Left            =   3240
         List            =   "frmCust.frx":A64D
         RightToLeft     =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   25
         Top             =   240
         Width           =   1900
      End
      Begin MSMask.MaskEdBox txtBirthDate 
         Height          =   465
         Left            =   120
         TabIndex        =   100
         Top             =   240
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   820
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "####/##/##"
         PromptChar      =   " "
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   "  «—ÌŒ  Ê·œ:"
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   12
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   2160
         RightToLeft     =   -1  'True
         TabIndex        =   101
         Top             =   240
         Width           =   945
      End
      Begin VB.Label LblRfid 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   " ﬂœRFID   "
         BeginProperty Font 
            Name            =   "Nazanin"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   285
         Left            =   2280
         RightToLeft     =   -1  'True
         TabIndex        =   99
         Top             =   960
         Width           =   1035
      End
      Begin VB.Label LblEconomicalCode 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "ﬂœ «ﬁ ’«œÌ"
         BeginProperty Font 
            Name            =   "Nazanin"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   405
         Left            =   1680
         RightToLeft     =   -1  'True
         TabIndex        =   98
         Top             =   1395
         Width           =   1515
      End
      Begin VB.Label lblPaykFee 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H80000002&
         Height          =   405
         Left            =   5220
         RightToLeft     =   -1  'True
         TabIndex        =   24
         Top             =   1380
         Width           =   855
      End
      Begin VB.Label lblCarryFee 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H80000002&
         Height          =   405
         Left            =   5160
         RightToLeft     =   -1  'True
         TabIndex        =   23
         Top             =   855
         Width           =   915
      End
      Begin VB.Label lblDistance 
         Alignment       =   1  'Right Justify
         Caption         =   "„ÕœÊœÂ"
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
         Left            =   4440
         RightToLeft     =   -1  'True
         TabIndex        =   22
         Top             =   195
         Width           =   1395
      End
   End
   Begin VB.Frame framePerson 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   2340
      Left            =   10440
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   2160
      Width           =   4035
      Begin VB.ComboBox cmbPrefix 
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
         ItemData        =   "frmCust.frx":A64F
         Left            =   840
         List            =   "frmCust.frx":A651
         RightToLeft     =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   720
         Width           =   1695
      End
      Begin VB.ComboBox cmbGender 
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
         ItemData        =   "frmCust.frx":A653
         Left            =   405
         List            =   "frmCust.frx":A655
         RightToLeft     =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   210
         Width           =   2175
      End
      Begin VB.TextBox txtName 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0E0FF&
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
         Left            =   405
         RightToLeft     =   -1  'True
         TabIndex        =   0
         Top             =   1270
         Width           =   2175
      End
      Begin VB.TextBox txtFamily 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0E0FF&
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
         Left            =   405
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   1700
         Width           =   2175
      End
      Begin FLWCtrls.FWButton FWBtnPrefix 
         Height          =   405
         Left            =   360
         TabIndex        =   10
         ToolTipText     =   " €ÌÌ— Ê«ÕœÂ«"
         Top             =   720
         Visible         =   0   'False
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   714
         ButtonType      =   5
         Caption         =   "FWButton1"
         BackColor       =   12632256
         FontName        =   "Arial"
         Object.ToolTipText     =   " €ÌÌ— Ê«ÕœÂ«"
      End
      Begin VB.Label lblPrefix 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H80000002&
         Height          =   405
         Left            =   2760
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   760
         Width           =   1005
      End
      Begin VB.Label lblSex 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H80000002&
         Height          =   405
         Left            =   2760
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   240
         Width           =   1005
      End
      Begin VB.Label lblName 
         Alignment       =   1  'Right Justify
         Caption         =   " ‰«„"
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   405
         Left            =   2880
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   1280
         Width           =   885
      End
      Begin VB.Label lblFamily 
         Alignment       =   1  'Right Justify
         Caption         =   "* ‰«„ Œ«‰Ê«œêÌ"
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
         Left            =   2640
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   1800
         Width           =   1245
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "* ¬œ—” "
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
      Height          =   2205
      Left            =   10440
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   4410
      Width           =   4035
      Begin VB.TextBox TxtMemberCode 
         Alignment       =   2  'Center
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
         Left            =   2340
         RightToLeft     =   -1  'True
         TabIndex        =   106
         Text            =   "òœ «‘ —«ò"
         Top             =   1560
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.TextBox txtAddress 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "Nazanin"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   900
         Left            =   120
         MultiLine       =   -1  'True
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   300
         Width           =   3795
      End
      Begin FLWCtrls.FWCheck fwCheckAssansor 
         Height          =   270
         Left            =   2100
         TabIndex        =   8
         Top             =   1305
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   476
         Value           =   0   'False
         Caption         =   "¬”«‰”Ê—"
         Color           =   255
         ForeColor       =   -2147483646
         FontName        =   "B Homa"
         FontBold        =   -1  'True
         FontSize        =   9.75
         Alignment       =   1
      End
      Begin FLWCtrls.FWButton FWBtnPerson 
         Height          =   495
         Left            =   120
         TabIndex        =   107
         Top             =   1230
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   873
         ButtonType      =   8
         Caption         =   "«⁄÷«Ì «‘ —«ﬂ"
         BackColor       =   49152
         ForeColor       =   16384
         FontName        =   "B Homa"
         FontBold        =   -1  'True
         FontSize        =   9.75
         Alignment       =   1
      End
      Begin FLWCtrls.FWButton FWBtnpicture 
         Height          =   405
         Left            =   90
         TabIndex        =   120
         Top             =   1755
         Width           =   2025
         _ExtentX        =   3572
         _ExtentY        =   714
         ButtonType      =   8
         Caption         =   " ’ÊÌ— „‘ —ﬂ"
         BackColor       =   16576
         ForeColor       =   16576
         FontName        =   "B Homa"
         FontBold        =   -1  'True
         FontSize        =   9.75
         Alignment       =   1
      End
   End
   Begin VSFlex7LCtl.VSFlexGrid vsCustomer 
      Height          =   3435
      Left            =   120
      TabIndex        =   6
      Top             =   6720
      Width           =   14355
      _cx             =   25321
      _cy             =   6059
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   8438015
      ForeColor       =   12582912
      BackColorFixed  =   16744576
      ForeColorFixed  =   4194304
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   8421631
      BackColorAlternate=   8438015
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
      Rows            =   1
      Cols            =   14
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   500
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmCust.frx":A657
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
      Height          =   405
      Left            =   13080
      Top             =   0
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   714
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
   Begin VB.Frame frameCompany 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   2340
      Left            =   10440
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   2160
      Width           =   4035
      Begin VB.TextBox txtCode 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
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
         Left            =   600
         RightToLeft     =   -1  'True
         TabIndex        =   81
         TabStop         =   0   'False
         Top             =   840
         Visible         =   0   'False
         Width           =   1290
      End
      Begin VB.ComboBox cmbActKind 
         BackColor       =   &H00C0E0FF&
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
         ItemData        =   "frmCust.frx":A7D4
         Left            =   480
         List            =   "frmCust.frx":A7D6
         RightToLeft     =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   1440
         Width           =   2160
      End
      Begin VB.TextBox txtWorkName 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0E0FF&
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
         Left            =   570
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   360
         Width           =   2055
      End
      Begin FLWCtrls.FWButton FWBtnActKind 
         Height          =   405
         Left            =   120
         TabIndex        =   17
         ToolTipText     =   " €ÌÌ— Ê«ÕœÂ«"
         Top             =   1440
         Visible         =   0   'False
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   714
         ButtonType      =   5
         Caption         =   "FWButton1"
         BackColor       =   12632256
         FontName        =   "Arial"
         Object.ToolTipText     =   " €ÌÌ— Ê«ÕœÂ«"
      End
      Begin VB.Label lblActKind 
         Alignment       =   2  'Center
         Caption         =   "‰Ê⁄ ›⁄«·Ì "
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   405
         Left            =   2640
         RightToLeft     =   -1  'True
         TabIndex        =   19
         Top             =   1440
         Width           =   1155
      End
      Begin VB.Label lblWorkName 
         Alignment       =   1  'Right Justify
         Caption         =   "* ‰«„ „Õ·"
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   405
         Left            =   2640
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   360
         Width           =   1155
      End
   End
   Begin RESIZEKITLibCtl.ResizeKit ResizeKit1 
      Height          =   480
      Left            =   0
      OleObjectBlob   =   "frmCust.frx":A7D8
      TabIndex        =   29
      Top             =   0
      Width           =   480
   End
   Begin VB.Frame Frame5 
      Height          =   1215
      Left            =   10440
      RightToLeft     =   -1  'True
      TabIndex        =   72
      Top             =   960
      Width           =   4035
      Begin VB.TextBox txtMembershipId 
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
         Height          =   420
         Left            =   480
         RightToLeft     =   -1  'True
         TabIndex        =   74
         Top             =   600
         Width           =   1425
      End
      Begin VB.TextBox txtMaxMembershipId 
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
         Height          =   420
         Left            =   480
         RightToLeft     =   -1  'True
         TabIndex        =   73
         Top             =   150
         Width           =   1425
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "ﬂœ «‘ —«ﬂ"
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   405
         Left            =   2160
         RightToLeft     =   -1  'True
         TabIndex        =   76
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   " „«ﬂ“Ì„„ ﬂœ «‘ —«ﬂ"
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
         Left            =   2160
         RightToLeft     =   -1  'True
         TabIndex        =   75
         Top             =   120
         Width           =   1695
      End
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "‰«„ ‘⁄»Â"
      BeginProperty Font 
         Name            =   "B Homa"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   3240
      RightToLeft     =   -1  'True
      TabIndex        =   82
      Top             =   1320
      Width           =   375
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "«ÿ·«⁄«  „‘ —ﬂÌ‰"
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
      Left            =   6000
      RightToLeft     =   -1  'True
      TabIndex        =   20
      Top             =   -120
      Width           =   2175
   End
End
Attribute VB_Name = "frmCust"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private clsDate As New clsDate
Private cn As New ADODB.Connection
Private Rc As New ADODB.Recordset
Private rctmp As New ADODB.Recordset
Private iHeight As Integer
Private iWidth As Integer
Private capfld As New Collection
Private objfld As New Collection
Public mvarcode2 As String
Private mvarcmbActKind As Boolean
Private mvarcmbPrefix As Boolean
Private mvarLast As Boolean
Private mvarArrowKey As Boolean
Private varclick As Boolean
Private MyFormAddEditMode As EnumAddEditMode
Dim Parameter() As Parameter
Dim i As Long
Dim mvarGender As Integer
Dim Exit_Form As Integer
Dim OldTafsili As Long
Dim OldMembershipId As String
Dim OldSwitchName As String
Dim OldTell As String
Public mvarCustName As String
Public mvarCustFamily As String
Public mvarMemberShipId As String
Private RfidReaderIsActive As Boolean
Dim RFIDStatus As String

Public Sub Find()

        frmFindCust.Show vbModal
        
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

Public Sub Delete()
    Select Case clsStation.Language
        Case 0
            frmMsg.fwlblMsg.Caption = "¬Ì« „ÿ„∆‰Ìœ „Ì ŒÊ«ÂÌœ «‘ —«ò " & "'" & txtMembershipId.Text & "'" & " —« Õ–› ﬂ‰Ìœø"
            frmMsg.fwBtn(0).Caption = "»·Ì"
            frmMsg.fwBtn(1).Caption = "ŒÌ—"
        Case 1
            frmMsg.fwlblMsg.Caption = "You are going to delete '" & txtMembershipId.Text & "'" + vbNewLine + "Are you sure ?"
            frmMsg.fwBtn(0).Caption = "Yes"
            frmMsg.fwBtn(1).Caption = "No"
            frmMsg.fwlblMsg.Alignment = vbLeftJustify
    End Select
    
    frmMsg.Show vbModal
    
    If modgl.mvarMsgIdx = vbNo Then
        Exit Sub
    End If
    ReDim Parameter(2) As Parameter
    Parameter(0) = GenerateInputParameter("@Code", adInteger, 4, txtCode.Text)
    Parameter(1) = GenerateInputParameter("@Branch", adInteger, 4, CurrentBranch)
    Parameter(2) = GenerateOutputParameter("@Result", adInteger, 4)
    Dim Result As Integer
    Result = RunParametricStoredProcedure("Delete_tCust", Parameter)
    
    If Result = 0 Then
    
        Select Case clsStation.Language

            Case 0
                frmMsg.fwlblMsg.Caption = "„‘ò·Ì œ—Õ–› «Ì‰ „‘ —Ì ÊÃÊœ œ«—œ ‘„« ‰„Ì  Ê«‰Ìœ «Ì‰ «‘ —«ò —« Õ–› ò‰Ìœ"
                frmMsg.fwBtn(0).Caption = "ﬁ»Ê·"
            Case 1
                frmMsg.fwlblMsg.Caption = "There are some factors related to this MembershipId , you cant delete it"
                frmMsg.fwBtn(0).Caption = "Ok"
                frmMsg.fwlblMsg.Alignment = vbLeftJustify
        End Select
        frmMsg.fwBtn(1).Visible = False
        frmMsg.fwBtn(0).ButtonType = flwButtonOk
        frmMsg.Show vbModal
    
    Else
    
        Select Case clsStation.Language
            Case 0
                frmMsg.fwlblMsg.Caption = "‘„« Ìò «‘ —«ò —« Õ–› ò—œÌœ"
                frmMsg.fwBtn(0).Caption = "ﬁ»Ê·"
            Case 1
                frmMsg.fwlblMsg.Caption = "You have deleted one MembershipId"
                frmMsg.fwBtn(0).Caption = "Ok"
                frmMsg.fwlblMsg.Alignment = vbLeftJustify
        End Select
        
        frmMsg.fwBtn(1).Visible = False
        frmMsg.fwBtn(0).ButtonType = flwButtonOk
        frmMsg.Show vbModal
        
    End If
    
    MyFormAddEditMode = AddMode
    SetFirstToolBar
    If vsCustomer.Rows > 1 Then
        vsCustomer.RemoveItem (vsCustomer.Row)
    End If
    Add
    
End Sub

Private Sub FillvsCustomer(MembershipId As Double, SwitchName As String, Tel1 As String, Status As Integer, SwitchValue As Integer)
    vsCustomer.Rows = 1
    Dim Rst As New ADODB.Recordset
    ReDim Parameter(6) As Parameter
    
    
    Parameter(0) = GenerateInputParameter("@MainCust", adBoolean, 1, 1)
    Parameter(1) = GenerateInputParameter("@MembershipId", adBigInt, 8, MembershipId)
    Parameter(2) = GenerateInputParameter("@SwitchName", adVarWChar, 50, SwitchName)
    Parameter(3) = GenerateInputParameter("@Tel1", adVarWChar, 50, Tel1)
    Parameter(4) = GenerateInputParameter("@Status", adInteger, 4, Status)
    Parameter(5) = GenerateInputParameter("@SwitchValue", adInteger, 4, SwitchValue)
    Parameter(6) = GenerateInputParameter("@Branch", adInteger, 4, CurrentBranch)
    Set Rst = RunParametricStoredProcedure2Rec("Get_Customer", Parameter)
    
    If Rst.EOF = True And Rst.BOF = True Then Exit Sub
    
    If Status = 0 Then
        With vsCustomer
            .Rows = 1
            i = 0
            FWProgressBar1.Value = 0
            MousePointer = 11
            
            Static arr As Variant
            arr = Rst.GetRows
            
            ' reset the control
            .BindToArray Null
            '  SetDefaults fa
            
            ' set the properties we want
            
'            While Rst.EOF <> True
'                .Rows = .Rows + 1
'                i = .Rows - 1
'                .TextMatrix(i, 0) = i
'                .TextMatrix(i, 1) = Rst!Code
'                .TextMatrix(i, 2) = Rst![Full Name]
'                .TextMatrix(i, 3) = Rst!MembershipId
'                .TextMatrix(i, 4) = Rst!Tel1
'                If Rst!WorkName.Value <> "" Then
'                .TextMatrix(i, 5) = -1
'                End If
'                .TextMatrix(i, 6) = Rst!address
'                .TextMatrix(i, 7) = Rst!Discount
'                .TextMatrix(i, 8) = Rst!Credit
'                .TextMatrix(i, 9) = Rst!carryfee
'                .TextMatrix(i, 10) = Rst!PaykFee
'                .Cell(flexcpText, i, 11) = CStr(Rst!Distance)
'                .TextMatrix(i, 12) = Rst!FamilyNo
'                .TextMatrix(i, 13) = IIf(Rst!Member = True, -1, 0)
'
'                If i Mod 1000 = 0 Then DoEvents
'                If i Mod 100 = 0 Then
'                    FWProgressBar1.Value = FWProgressBar1 + 1
'                    If FWProgressBar1.Value = 100 Then
'                        FWProgressBar1.Value = 1
'                    End If
'                End If
'
'                Rst.MoveNext
'            Wend
            
            .LoadArray arr
              
            For i = 1 To .Rows - 1
                .TextMatrix(i, 0) = i
            Next
            .Cols = 14
            MousePointer = 0
        End With
     Else
        With VsCustomer2
            .Rows = 1
            i = 0
            While Rst.EOF <> True
                .Rows = .Rows + 1
                i = .Rows - 1
                .TextMatrix(i, 0) = i
                .TextMatrix(i, 1) = Rst!Code
                .TextMatrix(i, 2) = Rst![Full Name]
                .TextMatrix(i, 3) = Rst!MembershipId
                .TextMatrix(i, 4) = Rst!Tel1
                If Rst!WorkName.Value <> "" Then
                .TextMatrix(i, 5) = -1
                End If
                .TextMatrix(i, 6) = Rst!address
                .TextMatrix(i, 7) = Rst!Discount
                .TextMatrix(i, 8) = Rst!Credit
                .TextMatrix(i, 9) = Rst!carryfee
                .TextMatrix(i, 10) = Rst!PaykFee
                .Cell(flexcpText, i, 11) = CStr(Rst!Distance)
                .TextMatrix(i, 12) = Rst!FamilyNo
                .TextMatrix(i, 13) = IIf(Rst!Member = True, -1, 0)
                
                
                Rst.MoveNext
            Wend
        End With
    End If
    Set Rst = Nothing
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

Private Sub cmbDistance_Click()
    ChkDistance.Value = Unchecked
    ChkDistance.Enabled = False
    If MyFormAddEditMode <> ViewMode Then
        If cmbDistance.ListIndex > 0 Then
            Dim Rst As New ADODB.Recordset
            ReDim Parameter(0) As Parameter
            Parameter(0) = GenerateInputParameter("@Code", adInteger, 4, cmbDistance.ItemData(cmbDistance.ListIndex))
            Set Rst = RunParametricStoredProcedure2Rec("Get_tblTotal_tDistance_ByCode", Parameter)
            If Not (Rst.EOF = True And Rst.BOF = True) Then
                txtCarryFee.Text = Rst!carryfee
                txtPaykFee.Text = Rst!PaykFee
                ChkDistance.Enabled = True
             End If
             Rst.Close
             Set Rst = Nothing
        Else
            txtCarryFee.Text = ""
            txtPaykFee.Text = ""
            ChkDistance.Enabled = False
        End If
      End If
End Sub

Private Sub cmdAddTafsili_Click()
    ShowMessage "¬Ì« „Ì ŒÊ«ÂÌœ »—«Ì ﬂ·ÌÂ „‘ —Ì«‰ Ê «‘Œ«’ œ«—«Ì «⁄ »«—'  ›÷Ì·Ì ÃœÌœ œ— ”Ì” „ Õ”«»œ«—Ì «ÌÕ«œ ò‰Ìœø ", True, True, "»·Ì", "ŒÌ—"
    If modgl.mvarMsgIdx = vbNo Then
        Exit Sub
    End If
    
    ReDim Parameter(1) As Parameter
    Dim Rst As New ADODB.Recordset
    
    Parameter(0) = GenerateInputParameter("@ActDeact", adInteger, 4, 0) ' Only Active
    Parameter(1) = GenerateInputParameter("@Branch", adInteger, 4, cmbBranch.ItemData(cmbBranch.ListIndex))
    Set Rst = RunParametricStoredProcedure2Rec("Get_All_Customers", Parameter)
    
    txtFamily.Text = ""
    txtWorkName.Text = ""
    While Rst.EOF <> True
        If Rst!Credit > 0 Then
            txtTafsiliCode.Text = ""
            txtName.Text = Rst!Name
            If IsNull(Rst!Tafsili) = True Or Trim(Rst!Tafsili) = "" Then Insert_Tafsili Rst!Code, False
        End If
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
    
    Dim hMenu As Long

    hMenu = GetSystemMenu(Me.hWnd, False)

    DeleteMenu hMenu, 6, MF_BYPOSITION
    
    VarActForm = Me.Name
    SetFirstToolBar
    varclick = False

    Add
    txtTel1 = NewCallNumber         ' from Caller Id- new call
    
    OnTopMe Me, True
    Me.ZOrder (0)
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case Shift
          Case 0
              Select Case KeyCode
                  Case 27  ' Esc
            
                    Me.ExitForm
                  Case 13  ' Esc
            
                    Me.Update
                  
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

    On Error GoTo ErrHandler
    
    CenterTop Me
'    Me.Top = 0
    If ClsFormAccess.frmCust = False Then
        Unload Me
        Exit Sub
    End If
    
    If intVersion = Min Then
        txtDiscount.Enabled = False
        cmbSellPrice.Enabled = False
    End If
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

     
     Set rctmp = RunStoredProcedure2RecordSet("Get_All_tblPub_SellPrice")
    If Not (rctmp.EOF = True And rctmp.BOF = True) Then
        While Not rctmp.EOF
            cmbSellPrice.AddItem rctmp!Description
            cmbSellPrice.ItemData(cmbSellPrice.NewIndex) = rctmp!Code
            rctmp.MoveNext
        Wend
    Else
        cmbSellPrice.AddItem " ‰—Œ «Ê·"
        cmbSellPrice.ItemData(0) = 1
    End If
    Me.cmbSellPrice.ListIndex = 0
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
    cmbDistance.AddItem ""
    cmbDistance.ItemData(cmbDistance.NewIndex) = -1
    Set rctmp = RunStoredProcedure2RecordSet("Get_All_tDistance")
    vsCustomer.ColComboList(11) = "#0;|" & vsCustomer.BuildComboList(rctmp, "Description", "Code")
    VsCustomer2.ColComboList(11) = "#0;|" & VsCustomer2.BuildComboList(rctmp, "Description", "Code")
    
    Set rctmp = RunStoredProcedure2RecordSet("Get_All_tDistance")
    If Not (rctmp.EOF = True And rctmp.BOF = True) Then
        While Not rctmp.EOF
            cmbDistance.AddItem rctmp!Description
            cmbDistance.ItemData(cmbDistance.NewIndex) = rctmp!Code
            rctmp.MoveNext
        Wend
    Else
        cmbDistance.AddItem " "
        cmbDistance.ItemData(0) = 0
    End If

    rctmp.Close
    
    FillBranch
    
    If mvarCategory = Club Then 'Bank Of Tejarat
        vsCustomer.ColHidden(4) = True
        vsCustomer.ColHidden(5) = True
        vsCustomer.ColHidden(6) = True
        vsCustomer.ColHidden(7) = True
        vsCustomer.ColHidden(8) = True
        vsCustomer.ColHidden(9) = True
        vsCustomer.ColHidden(10) = True
        vsCustomer.ColHidden(12) = False
    Else
        Label3.Visible = False
        Frame7.Visible = False
        Label6.Visible = False
        ChkMember.Visible = False
        Picture1.Visible = False
        ChkCentral.Visible = False
        TxtFamilyNo.Visible = False
        vsCustomer.ColHidden(11) = False    ' Distance
        vsCustomer.ColHidden(12) = True     'FamilyNo
        vsCustomer.ColHidden(13) = True     'FamilyNo
    End If
    
'    If Not (mvarCategory = Restaurant Or mvarCategory = Club) Then
'        frameBuyState.Enabled = False
'    Else
'        LblEconomicalCode.Visible = False
'        TxtEconomicalCode.Visible = False
'    End If
    
    If clsArya.Delivery = False Then
        LblCarryFee.Visible = False
        lblPaykFee.Visible = False
        lblDistance.Visible = False
        txtCarryFee.Visible = False
        txtPaykFee.Visible = False
        cmbDistance.Visible = False
        ChkDistance.Visible = False
    End If
    
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
            .ColWidth(i) = Val(GetSetting(strMainKey, "frmCust_vsCustomer", "Col" & i))
            If .ColWidth(i) = 0 Then
                .ColWidth(i) = 1000       '
            End If
         Next i
    End With
    If ClsFormAccess.ChangeTotalRemainingAmount = True Then
        txtPrimaryBedehi.Enabled = True
        txtPrimaryTalab.Enabled = True
    Else
        txtPrimaryBedehi.Enabled = False
        txtPrimaryTalab.Enabled = False
    End If

    If clsStation.RfidReader = True Then GetProperController

Exit Sub
ErrHandler:
    ShowDisMessage err.Description, 1500
End Sub
Private Sub GetProperController()
    
    RfidReaderIsActive = False
    Dim rctmp As New ADODB.Recordset
    ReDim Parameter(0) As Parameter
    Parameter(0) = GenerateInputParameter("@StationId", adInteger, 4, clsArya.StationNo)
    Set rctmp = RunParametricStoredProcedure2Rec("Get_DeviceSetting", Parameter)
    
    Do While (rctmp.EOF <> True)
        If rctmp.Fields("DeviceCode").Value = EnumDevice.RFT230 And rctmp.Fields("PortNo").Value <> 20 Then    ' Printer Port
            MF_ExitComm
            RFIDStatus = MF_InitComm("Com" & rctmp.Fields("PortNo").Value, rctmp.Fields("BaudRate").Value)
            If RFIDStatus = 0 Then RfidReaderIsActive = True: ShowDisMessage "”Ì” „ ò«—  ŒÊ«‰ „«Ì›— ›⁄«· ‘œ", 1000 Else ShowDisMessage "«‘ò«· œ— « ’«· ”Ì” „ ò«—  ŒÊ«‰", 1500
            Exit Do
        ElseIf rctmp.Fields("DeviceCode").Value = EnumDevice.RFT230 And rctmp.Fields("PortNo").Value = 20 Then    ' Printer Port
            MF_ExitComm
            RFIDStatus = MF_InitComm("USB", rctmp.Fields("BaudRate").Value)
            If RFIDStatus = 0 Then RfidReaderIsActive = True: ShowDisMessage "”Ì” „ ò«—  ŒÊ«‰ „«Ì›— ›⁄«· ‘œ", 1000 Else ShowDisMessage "«‘ò«· œ— « ’«· ”Ì” „ ò«—  ŒÊ«‰", 1500
            Exit Do
        End If
        rctmp.MoveNext
    Loop

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
    
    If PosConnection.State = adStateOpen Then PosConnection.Close: Set PosConnection = Nothing
    
    Set clsDate = Nothing
    Set mdifrm.FileCls = Nothing
    mdifrm.Toolbar1.Buttons(20).Enabled = False
    mdifrm.Toolbar1.Buttons(21).Enabled = False
    modgl.mvarDeleteMsg = ""
    Unload frmCust
    Exit_Form = 0
    Dim Obj As Object
    For Each Obj In Forms
        If LCase(Obj.Name) = "frminvoice" Then
            Exit_Form = 2
        End If
    Next Obj

    SaveSetting strMainKey, Me.Name, "Left", Me.left
    SaveSetting strMainKey, Me.Name, "Top", Me.top

    If RfidReaderIsActive = True Then MF_ExitComm

End Sub

'Private Sub Form_Resize()
'    On Error Resume Next
'     If Me.ScaleHeight > 0 Then
'        Me.Height = iHeight
'        Me.Width = iWidth
'     End If
'End Sub
Public Sub BeforeFirstKey()
''''    If MyFormAddEditMode <> ViewMode Then
''''        Cancel
''''    End If
End Sub

Public Sub FirstKey()
    MyFormAddEditMode = ViewMode  'View Mode
    SetFirstToolBar
    Dim i As Long
    ReDim Parameter(2) As Parameter
    Parameter(0) = GenerateInputParameter("@CurrentCustomerCode", adInteger, 4, 0)
    Parameter(1) = GenerateInputParameter("@Direction", adInteger, 4, 0)
    Parameter(2) = GenerateInputParameter("@Branch", adInteger, 4, CurrentBranch)
    Set rctmp = RunParametricStoredProcedure2Rec("NavigateInCustomer", Parameter)
    If rctmp.EOF <> True And rctmp.BOF <> True Then
        Me.txtCode.Text = rctmp.Fields("code").Value
        If clsStation.CustomerSearchDefault = False Then
            i = vsCustomer.FindRow(txtCode.Text, 1, 1, True, True)
            If i > 0 Then
                vsCustomer.Row = i
                 vsCustomer.ShowCell i, 0
            End If
        Else
            GetDataDetail
        End If
    End If
    rctmp.Close
    
    HeaderLabel Val(MyFormAddEditMode), fwlblMode
    
End Sub

Public Sub BeforePreviousKey()
End Sub

Public Sub PreviousKey()
    MyFormAddEditMode = ViewMode  'View Mode
    SetFirstToolBar
    Dim i As Long
    
    ReDim Parameter(2) As Parameter
    Parameter(0) = GenerateInputParameter("@CurrentCustomerCode", adInteger, 4, Val(txtCode.Text))
    Parameter(1) = GenerateInputParameter("@Direction", adInteger, 4, 1)
    Parameter(2) = GenerateInputParameter("@Branch", adInteger, 4, CurrentBranch)
    Set rctmp = RunParametricStoredProcedure2Rec("NavigateInCustomer", Parameter)
    If rctmp.EOF <> True And rctmp.BOF <> True Then
        txtCode.Text = rctmp.Fields("code").Value
        If clsStation.CustomerSearchDefault = False Then
            i = vsCustomer.FindRow(txtCode.Text, 1, 1, True, True)
            If i > 0 Then
                vsCustomer.Row = i
                 vsCustomer.ShowCell i, 1
            End If
        Else
            GetDataDetail
        End If
    End If
    rctmp.Close
    
    HeaderLabel Val(MyFormAddEditMode), fwlblMode
End Sub

Public Sub BeforeNextKey()

End Sub

Public Sub NextKey()
    MyFormAddEditMode = ViewMode  'View Mode
    SetFirstToolBar
    Dim i As Long
    
    ReDim Parameter(2) As Parameter
    Parameter(0) = GenerateInputParameter("@CurrentCustomerCode", adInteger, 4, Val(txtCode.Text))
    Parameter(1) = GenerateInputParameter("@Direction", adInteger, 4, 2)
    Parameter(2) = GenerateInputParameter("@Branch", adInteger, 4, CurrentBranch)
    Set rctmp = RunParametricStoredProcedure2Rec("NavigateInCustomer", Parameter)
    If rctmp.EOF <> True And rctmp.BOF <> True Then
        txtCode.Text = rctmp.Fields("code").Value
        If clsStation.CustomerSearchDefault = False Then
            i = vsCustomer.FindRow(txtCode.Text, 1, 1, True, True)
            If i > 0 Then
                vsCustomer.Row = i
                 vsCustomer.ShowCell i, 0
            End If
        Else
            GetDataDetail
        End If
    End If
    rctmp.Close
 
    HeaderLabel Val(MyFormAddEditMode), fwlblMode
End Sub

Public Sub BeforeLastKey()
''''    If MyFormAddEditMode <> ViewMode Then
''''        Cancel
''''    End If
End Sub
Public Sub Cancel()
    Select Case MyFormAddEditMode
        Case AddMode 'new
            MyFormAddEditMode = AddMode
            SetFirstToolBar
            Add
            
        Case EditMode 'edit
            GetDataDetail
            MyFormAddEditMode = ViewMode
            SetFirstToolBar
    End Select
''''    vsCustomerClear
End Sub

Public Sub LastKey()
    MyFormAddEditMode = ViewMode  'View Mode
    SetFirstToolBar
    Dim i As Long
    
    ReDim Parameter(2) As Parameter
    Parameter(0) = GenerateInputParameter("@CurrentCustomerCode", adInteger, 4, 0)
    Parameter(1) = GenerateInputParameter("@Direction", adInteger, 4, 3)
    Parameter(2) = GenerateInputParameter("@Branch", adInteger, 4, CurrentBranch)
    Set rctmp = RunParametricStoredProcedure2Rec("NavigateInCustomer", Parameter)
    If rctmp.EOF <> True And rctmp.BOF <> True Then
        txtCode.Text = rctmp.Fields("code").Value
        If clsStation.CustomerSearchDefault = False Then
            i = vsCustomer.FindRow(txtCode.Text, 1, 1, True, True)
            If i > 0 Then
                vsCustomer.Row = i
                 vsCustomer.ShowCell i, 0
            End If
        Else
            GetDataDetail
        End If
    End If
    rctmp.Close
    
    HeaderLabel Val(MyFormAddEditMode), fwlblMode
End Sub
Public Sub DefaultSettings()

    On Error Resume Next
    
    txtPrimaryBedehi = ""
    txtPrimaryTalab = ""
    OldMembershipId = ""
    OldSwitchName = ""
    OldTell = ""
    cmbActKind.ListIndex = 0
    CmbState.ListIndex = 0
    cmbCity.ListIndex = 0
    If cmbDistance.ListCount > 2 Then cmbDistance.ListIndex = 1 Else cmbDistance.ListIndex = 0
    cmbPrefix.ListIndex = 0
    cmbGender.ListIndex = 0
    cmbSellPrice.ListIndex = 0
    
    On Error GoTo 0
    
    
    TxtAddress.Text = ""
    'txtCarryFee.Text = 0
    txtCredit.Text = 0
    txtDescription.Text = ""
    txtDiscount.Text = 0
    txtEmail.Text = ""
    txtFamily.Text = ""
    txtFax.Text = ""
    txtMobile.Text = ""
    txtName.Text = ""
    txtPostalCode.Text = ""
    txtTel1.Text = ""
    txtTel2.Text = ""
    txtTel3.Text = ""
    txtTel4.Text = ""
    txtWorkName.Text = ""
    txtMembershipId.Text = ""
    TxtFamilyNo.Text = ""
    txtCarryFee.Text = ""
    txtPaykFee.Text = ""
    TxtEconomicalCode.Text = ""
    TxtRfid.Text = ""
    ChkMember.Value = Checked
    ChkCentral.Value = Unchecked
    fwCheckAssansor.Value = 0
    
    OptionActDeAct(0).Value = True
    OptionOwner(0).Value = True
    
    txtTafsiliCode.Text = ""
    OldTafsili = 0
    txtBirthDate.Text = "    /  /  "
End Sub

Public Sub Add()
    
    VsCustomer2.Visible = False
    
    If MyFormAddEditMode = EditMode Then
        DefaultSettings
    End If
    MyFormAddEditMode = AddMode
    DefaultSettings

    ReDim Parameter(0) As Parameter
    Parameter(0) = GenerateInputParameter("@Branch", adInteger, 4, CurrentBranch)
    Set rctmp = RunParametricStoredProcedure2Rec("Get_New_Cust_Code", Parameter)
    txtMaxMembershipId.Text = rctmp.Fields("MembershipId").Value - 1
    
    txtCode.Text = rctmp.Fields("Code").Value
    txtMembershipId.Text = rctmp.Fields("MembershipId").Value
    
    If clsStation.CustomerSearchDefault = False And vsCustomer.Rows = 1 Then
        FillvsCustomer 0, "", "", 0, 0
        vsCustomer_AfterSort 3, 1
        TxtMemberCode.Visible = True
        vsCustomer.ShowCell 0, 0
    End If
    
    If OptionOwner(0).Value Then
        Me.OptionOwnerValue 0
    Else
        Me.OptionOwnerValue 1
    End If
    
    SetFirstToolBar
End Sub

Public Sub ExitSub()
If PosConnection.State = adStateOpen Then PosConnection.Close: Set PosConnection = Nothing
    Unload Me
End Sub

Public Sub Update()
    If MyFormAddEditMode = ViewMode Then Exit Sub
    Dim strBinBuyState As String
    Dim intBuyState As Integer
    
    If Val(txtDiscount.Text) < 0 Or Val(txtDiscount.Text) > 100 Then
        ShowMessage "„ﬁœ«—  Œ›Ì› ‰„Ì  Ê«‰œ ò„ — «“ ’›— Ì« »Ì‘ — «“ ’œ œ—’œ »«‘œ ", True, False, " «ÌÌœ", ""
        Exit Sub
    End If
    
    If framePerson.Visible = True Then
        If Trim(txtMembershipId.Text) = "" Or Trim(txtFamily.Text) = "" Or Trim(TxtAddress.Text) = "" Then
            ShowMessage "·ÿ›« «ÿ·«⁄«  ‰«„ Ê «‘ —«ﬂ —« »Â ÿÊ— ò«„· Ê«—œ ‰„«ÌÌœ", True, False, " «ÌÌœ", ""
            Exit Sub
        End If
    ElseIf frameCompany.Visible = True Then
        If Trim(txtWorkName.Text) = "" Or Trim(TxtAddress.Text) = "" Or Trim(txtMembershipId.Text) = "" Then
            ShowMessage "·ÿ›« «ÿ·«⁄«  „Õ· ﬂ«——« »Â ÿÊ— ò«„· Ê«—œ ‰„«ÌÌœ", True, False, " «ÌÌœ", ""
            Exit Sub
        End If
    End If
    If Trim(txtTel1.Text) = "" Then
        ShowMessage "·ÿ›« ‘„«—Â  ·›‰ —« Ê«—œ ﬂ‰Ìœ", True, False, " «ÌÌœ", ""
        Exit Sub
    End If
    If cmbDistance.ListIndex < 1 Then
        If clsArya.Delivery = True Then
            ShowMessage "·ÿ›« „ÕœÊœÂ —« «‰ Œ«» ﬂ‰Ìœ", True, False, " «ÌÌœ", ""
            Exit Sub
        Else
            cmbDistance.ListIndex = 4   'Only In Place
        End If
    End If
    If OptionOwner(0).Value = True Then
        txtWorkName.Text = ""
        cmbActKind.ListIndex = 0
        mvarGender = cmbGender.ItemData(cmbGender.ListIndex)
    Else
        txtName.Text = ""
        txtFamily.Text = ""
 '       cmbGender.ListIndex = 0
        mvarGender = 2
        cmbPrefix.ListIndex = 0
    
    End If
    
    For i = 0 To 3
        If FWCheckBuy(i).Value = True Then
            strBinBuyState = strBinBuyState & "1"
        Else
            strBinBuyState = strBinBuyState & "0"
        End If
    Next i
    
    If Trim(txtBirthDate.ClipText) <> "" Then
        If Len(Trim(txtBirthDate.ClipText)) < 8 Then
            ShowMessage " ›Ì·œ  «—ÌŒ  Ê·œ —« ﬂ«„· Å— ﬂ‰Ìœ Ì« Œ«·Ì ﬂ‰Ìœ", True, False, " «ÌÌœ", ""
            Exit Sub
        End If
    End If
    intBuyState = ConvertBinToInt(strBinBuyState)
    
    'Mifare Card
    If RfidReaderIsActive = True Then 'And TxtRfid = ""
        If ReadCard = False Then Exit Sub
    End If
    Dim SwitchName As String, SwitchValue As Integer
                
    Select Case MyFormAddEditMode
        Case AddMode
            ReDim Parameter(5) As Parameter
                Parameter(0) = GenerateInputParameter("@intMode", adInteger, 4, 0)
                Parameter(1) = GenerateInputParameter("@MembershipId", adBigInt, 8, Val(txtMembershipId.Text))
                If OptionOwner(0).Value = True Then
                    SwitchName = Trim(txtFamily.Text)
                    SwitchValue = 0
                Else
                    SwitchName = Trim(txtWorkName.Text)
                    SwitchValue = 1
                End If
                Parameter(2) = GenerateInputParameter("@SwitchName", adVarWChar, 50, SwitchName)
                Parameter(3) = GenerateInputParameter("@Tel1", adVarWChar, 50, Trim(txtTel1.Text))
                Parameter(4) = GenerateInputParameter("@SwitchValue", adInteger, 4, SwitchValue)
                Parameter(5) = GenerateInputParameter("@Branch", adInteger, 4, CurrentBranch)
                Set rctmp = RunParametricStoredProcedure2Rec("Get_Check_tblTotal_tCust", Parameter)
                    If rctmp!intMember = 1 Then
                        VsCustomer2.Visible = True
                        FillvsCustomer Val(txtMembershipId.Text), SwitchName, Trim(txtTel1.Text), 1, SwitchValue
                        ShowMessage " «Ì‰ ‘„«—Â «‘ —«ﬂ ﬁ»·« œ— ”Ì” „  ⁄—Ì› ‘œÂ «” . ‘„«—Â «‘ —«ﬂ œÌê—Ì  ⁄—Ì› ò‰Ìœ ", True, False, " «ÌÌœ", ""
                        VsCustomer2.Visible = False
                        txtMembershipId.SetFocus
                        Exit Sub
                    End If
                    If rctmp!intSwitchName = 1 Then
                        VsCustomer2.Visible = True
                        FillvsCustomer Val(txtMembershipId.Text), SwitchName, Trim(txtTel1.Text), 1, SwitchValue
                        ShowMessage " «Ì‰ ‰«„ ﬁ»·« œ— ”Ì” „  ⁄—Ì› ‘œÂ «” .¬Ì« „Ì ŒÊ«ÂÌœ «Ì‰ «‘ —«ﬂ —« À»  ò‰Ìœ ", True, True, "»·Ì", "ŒÌ—"
                        If modgl.mvarMsgIdx = vbNo Then
                          Select Case SwitchValue
                            Case 0:
                                txtFamily.SetFocus
                            Case 1:
                                txtWorkName.SetFocus
                          End Select
                          Exit Sub
                        End If
                        VsCustomer2.Visible = False
                        
                        
                    End If
                    If rctmp!intTel1 = 1 Then
                            VsCustomer2.Visible = True
                            FillvsCustomer Val(txtMembershipId.Text), SwitchName, Trim(txtTel1.Text), 1, SwitchValue
                            ShowMessage " «Ì‰  ·›‰ ﬁ»·« œ— ”Ì” „  ⁄—Ì› ‘œÂ «” .¬Ì« „Ì ŒÊ«ÂÌœ «Ì‰ «‘ —«ﬂ —« À»  ò‰Ìœ ", True, True, "»·Ì", "ŒÌ—"
                            If modgl.mvarMsgIdx = vbNo Then
                                txtTel1.SetFocus
                                Exit Sub
                            End If
                            VsCustomer2.Visible = False
                    End If
            ReDim Parameter(42) As Parameter
            Parameter(0) = GenerateInputParameter("@MembershipId", adVarWChar, 50, Val(txtMembershipId.Text))
            Parameter(1) = GenerateInputParameter("@MasterCode", adInteger, 4, 0)
            Parameter(2) = GenerateInputParameter("@Owner", adInteger, 4, IIf(OptionOwner(0).Value = True, 0, 1))
            Parameter(3) = GenerateInputParameter("@Name", adVarWChar, 50, Trim(txtName.Text))
            Parameter(4) = GenerateInputParameter("@Family", adVarWChar, 50, Trim(txtFamily.Text))
            Parameter(5) = GenerateInputParameter("@Sex", adInteger, 4, mvarGender)
            Parameter(6) = GenerateInputParameter("@WorkName", adVarWChar, 50, txtWorkName.Text)
            Parameter(7) = GenerateInputParameter("@InternalNo", adVarWChar, 50, "")
            Parameter(8) = GenerateInputParameter("@Unit", adVarWChar, 50, "")
            Parameter(9) = GenerateInputParameter("@City", adInteger, 4, cmbCity.ItemData(cmbCity.ListIndex))
            Parameter(10) = GenerateInputParameter("@ActKind", adInteger, 4, cmbActKind.ItemData(cmbActKind.ListIndex))
            Parameter(11) = GenerateInputParameter("@ActDeAct", adInteger, 4, IIf(OptionActDeAct(0).Value = True, 1, 0))
            Parameter(12) = GenerateInputParameter("@Prefix", adInteger, 4, cmbPrefix.ItemData(cmbPrefix.ListIndex))
            Parameter(13) = GenerateInputParameter("@Assansor", adInteger, 4, IIf(fwCheckAssansor.Value = True, 1, 0))
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
            Parameter(24) = GenerateInputParameter("@CarryFee", adDouble, 8, Val(txtCarryFee.Text))
            Parameter(25) = GenerateInputParameter("@PaykFee", adDouble, 8, Val(txtPaykFee.Text))
            Parameter(26) = GenerateInputParameter("@Distance", adInteger, 4, cmbDistance.ItemData(cmbDistance.ListIndex))
            Parameter(27) = GenerateInputParameter("@Credit", adDouble, 8, Val(txtCredit.Text))
            Parameter(28) = GenerateInputParameter("@Discount", adDouble, 8, Val(txtDiscount.Text))
            Parameter(29) = GenerateInputParameter("@BuyState", adInteger, 4, intBuyState)
            Parameter(30) = GenerateInputParameter("@Description", adVarWChar, 255, Trim(txtDescription.Text))
            Parameter(31) = GenerateInputParameter("@User", adInteger, 4, mvarCurUserNo)
            Parameter(32) = GenerateInputParameter("@FamilyNo", adInteger, 4, IIf(TxtFamilyNo.Text = "", 0, TxtFamilyNo.Text))
            Parameter(33) = GenerateInputParameter("@Member", adBoolean, 1, IIf(ChkMember = Checked, 1, 0))
            Parameter(34) = GenerateInputParameter("@State", adInteger, 4, CmbState.ItemData(CmbState.ListIndex))
            Parameter(35) = GenerateInputParameter("@Central", adBoolean, 1, IIf(ChkCentral = Checked, 0, 1))
            Parameter(36) = GenerateInputParameter("@Sellprice", adSmallInt, 2, cmbSellPrice.ItemData(cmbSellPrice.ListIndex))
            Parameter(37) = GenerateInputParameter("@EconomicCode", adVarWChar, 20, Trim(TxtEconomicalCode.Text))
            Parameter(38) = GenerateInputParameter("@nvcRFID", adVarWChar, 20, Trim(TxtRfid.Text))
            Parameter(39) = GenerateInputParameter("@nvcBirthDate", adVarWChar, 10, CStr(IIf(Trim(txtBirthDate.ClipText) = "", "", Trim(txtBirthDate.Text))))
            Parameter(40) = GenerateInputParameter("@TotalRemainingAmount", adDouble, 8, IIf(Val(txtPrimaryBedehi) > 0, Val(txtPrimaryBedehi), -1 * Val(txtPrimaryTalab)))
            Parameter(41) = GenerateInputParameter("@Branch", adInteger, 4, cmbBranch.ItemData(cmbBranch.ListIndex))
            Parameter(42) = GenerateOutputParameter("@Code", adBigInt, 8)
            
            Dim LastCode As Long
            LastCode = RunParametricStoredProcedure("Insert_Cust", Parameter)
            If LastCode > 0 Then
               '''' vsCustomerClear
                ShowMessage "À»  „‘ —ò ÃœÌœ »« „Ê›ﬁÌ  «‰Ã«„ ‘œ", True, False, " «ÌÌœ", ""
                
                If (clsArya.ExternalAccounting = True Or HasMiniAcc = True) And Val(txtCredit) > 0 Then
                   Insert_Tafsili LastCode, True
                End If
            Else
                ShowMessage "À»  «‰Ã«„ ‰‘œ" + vbCrLf + "«ÿ·«⁄«  —« »——”Ì ‰„«ÌÌœ.", True, False, " «ÌÌœ", ""
                txtMembershipId.SetFocus
                Exit Sub
            End If
            
        Case EditMode
                
                ReDim Parameter(5) As Parameter
                Parameter(0) = GenerateInputParameter("@intMode", adInteger, 4, 1)
                Parameter(1) = GenerateInputParameter("@MembershipId", adBigInt, 8, Val(txtMembershipId.Text))
                If OptionOwner(0).Value = True Then
                    SwitchName = Trim(txtFamily.Text)
                    SwitchValue = 0
                Else
                    SwitchName = Trim(txtWorkName.Text)
                    SwitchValue = 1
                End If
                Parameter(2) = GenerateInputParameter("@SwitchName", adVarWChar, 50, SwitchName)
                Parameter(3) = GenerateInputParameter("@Tel1", adVarWChar, 50, Trim(txtTel1.Text))
                Parameter(4) = GenerateInputParameter("@SwitchValue", adInteger, 4, SwitchValue)
                Parameter(5) = GenerateInputParameter("@Branch", adInteger, 4, CurrentBranch)
                Set rctmp = RunParametricStoredProcedure2Rec("Get_Check_tblTotal_tCust", Parameter)
                If Trim(OldMembershipId) <> Trim(txtMembershipId.Text) Then
                    If rctmp!intMember = 1 Then
                        VsCustomer2.Visible = True
                        FillvsCustomer Val(txtMembershipId.Text), "-999", "-999", 1, -999
                        frmMsg.fwlblMsg.Caption = " «Ì‰ ‘„«—Â «‘ —«ﬂ ﬁ»·« œ— ”Ì” „  ⁄—Ì› ‘œÂ «” . ‘„«—Â «‘ —«ﬂ œÌê—Ì  ⁄—Ì› ò‰Ìœ "
                        frmMsg.fwBtn(0).Visible = False
                        frmMsg.fwBtn(1).ButtonType = flwButtonOk
                        frmMsg.fwBtn(1).Caption = "ﬁ»Ê·"
                        frmMsg.Show vbModal
                        VsCustomer2.Visible = False
                        txtMembershipId.SetFocus
                        Exit Sub
                    End If
                End If
                If Trim(OldSwitchName) <> SwitchName Then
                    If rctmp!intSwitchName = 1 Then
                        VsCustomer2.Visible = True
                        FillvsCustomer -999, SwitchName, "-999", 1, SwitchValue
                        frmMsg.fwlblMsg.Caption = " «Ì‰ ›«„Ì·Ì ﬁ»·« œ— ”Ì” „  ⁄—Ì› ‘œÂ «” .¬Ì« „Ì ŒÊ«ÂÌœ «Ì‰ «‘ —«ﬂ —« À»  ò‰Ìœ "
                        frmMsg.fwBtn(0).ButtonType = flwButtonOk
                        frmMsg.fwBtn(0).Caption = "»·Ì"
                        frmMsg.fwBtn(1).ButtonType = flwButtonNo
                        frmMsg.fwBtn(1).Caption = "ŒÌ—"
                        frmMsg.Show vbModal
                        If modgl.mvarMsgIdx = vbNo Then
                            Select Case SwitchValue
                            Case 0:
                                txtFamily.SetFocus
                            Case 1:
                                txtWorkName.SetFocus
                            End Select
                            Exit Sub
                        End If
                        VsCustomer2.Visible = False
                    End If
                End If
                If Trim(OldTell) <> Trim(txtTel1.Text) Then
                    If rctmp!intTel1 = 1 Then
                            VsCustomer2.Visible = True
                            FillvsCustomer -999, "-999", Trim(txtTel1.Text), 1, -999
                            frmMsg.fwlblMsg.Caption = " «Ì‰  ·›‰ ﬁ»·« œ— ”Ì” „  ⁄—Ì› ‘œÂ «” .¬Ì« „Ì ŒÊ«ÂÌœ «Ì‰ «‘ —«ﬂ —« À»  ò‰Ìœ "
                            frmMsg.fwBtn(0).ButtonType = flwButtonOk
                            frmMsg.fwBtn(0).Caption = "»·Ì"
                            frmMsg.fwBtn(1).ButtonType = flwButtonNo
                            frmMsg.fwBtn(1).Caption = "ŒÌ—"
                            frmMsg.Show vbModal
                            If modgl.mvarMsgIdx = vbNo Then
                                txtTel1.SetFocus
                                Exit Sub
                            End If
                            VsCustomer2.Visible = False
                    End If
                End If
                
            ReDim Parameter(43) As Parameter
            Parameter(0) = GenerateInputParameter("@MembershipId", adVarWChar, 50, Val(txtMembershipId.Text))
            Parameter(1) = GenerateInputParameter("@MasterCode", adInteger, 4, 0)
            Parameter(2) = GenerateInputParameter("@Owner", adInteger, 4, IIf(OptionOwner(0).Value = True, 0, 1))
            Parameter(3) = GenerateInputParameter("@Name", adVarWChar, 50, Trim(txtName.Text))
            Parameter(4) = GenerateInputParameter("@Family", adVarWChar, 50, Trim(txtFamily.Text))
            Parameter(5) = GenerateInputParameter("@Sex", adInteger, 4, mvarGender)
            Parameter(6) = GenerateInputParameter("@WorkName", adVarWChar, 50, txtWorkName.Text)
            Parameter(7) = GenerateInputParameter("@InternalNo", adVarWChar, 50, "")
            Parameter(8) = GenerateInputParameter("@Unit", adVarWChar, 50, "")
            Parameter(9) = GenerateInputParameter("@City", adInteger, 4, cmbCity.ItemData(cmbCity.ListIndex))
            Parameter(10) = GenerateInputParameter("@ActKind", adInteger, 4, cmbActKind.ItemData(cmbActKind.ListIndex))
            Parameter(11) = GenerateInputParameter("@ActDeAct", adInteger, 4, IIf(OptionActDeAct(0).Value = True, 1, 0))
            Parameter(12) = GenerateInputParameter("@Prefix", adInteger, 4, cmbPrefix.ItemData(cmbPrefix.ListIndex))
            Parameter(13) = GenerateInputParameter("@Assansor", adInteger, 4, IIf(fwCheckAssansor.Value = True, 1, 0))
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
            Parameter(24) = GenerateInputParameter("@CarryFee", adDouble, 8, Val(txtCarryFee.Text))
            Parameter(25) = GenerateInputParameter("@PaykFee", adDouble, 8, Val(txtPaykFee.Text))
            Parameter(26) = GenerateInputParameter("@Distance", adInteger, 4, cmbDistance.ItemData(cmbDistance.ListIndex))
            Parameter(27) = GenerateInputParameter("@Credit", adDouble, 8, Val(txtCredit.Text))
            Parameter(28) = GenerateInputParameter("@Discount", adDouble, 8, Val(txtDiscount.Text))
            Parameter(29) = GenerateInputParameter("@BuyState", adInteger, 4, intBuyState)
            Parameter(30) = GenerateInputParameter("@Description", adVarWChar, 255, Trim(txtDescription.Text))
            Parameter(31) = GenerateInputParameter("@User", adInteger, 4, mvarCurUserNo)
            Parameter(32) = GenerateInputParameter("@Code", adBigInt, 8, Val(txtCode.Text))
            Parameter(33) = GenerateInputParameter("@FamilyNo", adInteger, 4, IIf(TxtFamilyNo.Text = "", 0, TxtFamilyNo.Text))
            Parameter(34) = GenerateInputParameter("@Member", adBoolean, 1, IIf(ChkMember = Checked, 1, 0))
            Parameter(35) = GenerateInputParameter("@State", adInteger, 4, CmbState.ItemData(CmbState.ListIndex))
            Parameter(36) = GenerateInputParameter("@Central", adBoolean, 1, IIf(ChkCentral = Checked, 0, 1))
            Parameter(37) = GenerateInputParameter("@Sellprice", adSmallInt, 2, cmbSellPrice.ItemData(cmbSellPrice.ListIndex))
            Parameter(38) = GenerateInputParameter("@EconomicCode", adVarWChar, 20, Trim(TxtEconomicalCode.Text))
            Parameter(39) = GenerateInputParameter("@nvcRFID", adVarWChar, 20, Trim(TxtRfid.Text))
            Parameter(40) = GenerateInputParameter("@nvcBirthDate", adVarWChar, 10, CStr(IIf(Trim(txtBirthDate.ClipText) = "", "", Trim(txtBirthDate.Text))))
            Parameter(41) = GenerateInputParameter("@TotalRemainingAmount", adDouble, 8, IIf(Val(txtPrimaryBedehi) > 0, Val(txtPrimaryBedehi), -1 * Val(txtPrimaryTalab)))
            Parameter(42) = GenerateInputParameter("@Branch", adInteger, 4, cmbBranch.ItemData(cmbBranch.ListIndex))
            Parameter(43) = GenerateOutputParameter("@Updated", adBigInt, 8)
            Dim Updated As Long
            Updated = RunParametricStoredProcedure("Update_Cust", Parameter)
            If Updated > 0 Then
                ''''vsCustomerClear
                frmMsg.fwlblMsg.Caption = " €ÌÌ—«  »« „Ê›ﬁÌ  «‰Ã«„ ‘œ"
                frmMsg.fwBtn(0).Visible = False
                frmMsg.fwBtn(1).ButtonType = flwButtonOk
                frmMsg.fwBtn(1).Caption = "ﬁ»Ê·"
                frmMsg.Show vbModal
                
                If (clsArya.ExternalAccounting = True Or HasMiniAcc = True) And Val(txtCredit) > 0 Then
                   Insert_Tafsili Updated, True
                End If
                
            Else
                frmMsg.fwlblMsg.Caption = " €ÌÌ—«  «‰Ã«„ ‰‘œ" + vbCrLf + "«ÿ·«⁄«  —« »——”Ì ‰„«ÌÌœ."
                frmMsg.fwBtn(1).Visible = False
                frmMsg.fwBtn(0).ButtonType = flwButtonCancel
                frmMsg.fwBtn(0).Caption = "ﬁ»Ê·"
                frmMsg.Show vbModal
                txtMembershipId.SetFocus
                Exit Sub
            End If

        End Select
    
    If vsCustomer.Rows > 1 Then
        With vsCustomer
            If MyFormAddEditMode = AddMode Then
                .Rows = .Rows + 1
                i = .Rows - 1
                .TextMatrix(i, 0) = i
                .TextMatrix(i, 1) = LastCode
            ElseIf MyFormAddEditMode = EditMode Then
                i = .Row
            End If
            If OptionOwner(0).Value = True Then
                .TextMatrix(i, 2) = txtFamily.Text & " " & txtName.Text
            Else
                .TextMatrix(i, 2) = txtWorkName.Text
                .TextMatrix(i, 5) = -1
            End If
            .TextMatrix(i, 3) = txtMembershipId.Text
            .TextMatrix(i, 4) = txtTel1.Text
            .TextMatrix(i, 6) = TxtAddress.Text
            .TextMatrix(i, 7) = txtDiscount.Text
            .TextMatrix(i, 8) = txtCredit.Text
            .TextMatrix(i, 9) = txtCarryFee.Text
            .TextMatrix(i, 10) = txtPaykFee.Text
            .TextMatrix(i, 11) = cmbDistance.ListIndex
            .TextMatrix(i, 12) = TxtFamilyNo.Text
            .TextMatrix(i, 13) = IIf(ChkCentral.Value = 1, -1, 0)
            
        End With
    
    End If
    MyFormAddEditMode = AddMode
    SetFirstToolBar
    Add
    If clsArya.LimitedVersion = True And HardLockFlagTrial = False And (RemaindateFlag = True Or maxRecordCountFlag = True) Then
        TrialCountFlag = TrialCountFlag + 1
        If TrialCountFlag Mod 2 = 0 Then
            ShowMessage " ‘„« œ— Õ«· «” ›«œÂ «“ ‰”ŒÂ ¬“„«Ì‘Ì ”Ì” „ Â«Ì ›—Ê‘ê«ÂÌ «› ÃÌ ¬—Ì« „Ì »«‘Ìœ" & " »—«Ì  ÂÌÂ ‰”ŒÂ «’·Ì ‰—„ «›“«—»« ‘—ﬂ  «› ÃÌ ¬—Ì« Ì« ‰„«Ì‰œê«‰ ›—Ê‘  „«” »êÌ—Ìœ ", True, False, " «∆Ìœ", ""
            Sleep 1000 * TrialCountFlag / 2
        End If
    End If
    
    
    Exit Sub
RollBack:
    
'    cnn.RollbackTrans
    
End Sub

Private Function ReadCard() As Boolean

    On Error GoTo ErrHandler
    ShowMessage "ò«—  —« —ÊÌ œ” ê«Â ò«—  ŒÊ«‰ ﬁ—«—œÂÌœ", True, False, "ﬁ»Ê·", ""

    Dim Status As Integer
    Status = 1
    If MF_Request(0, 1, cardT(0)) = 0 Then
        RFIDStatus = MF_Anticoll(0, cardSN(0))
        Sleep 1000
        If RFIDStatus = 0 Then
            For i = 0 To 5
                Ckey(i) = hex2dec(keyTXT(i))
            Next i
            If MF_Select(0, cardSN(0)) = 0 Then
                If MF_LoadKey(0, Ckey(0)) = 0 Then
                    If MF_Authentication(0, IIf(KeyAorB.Value, 1, 0), Val(blockNtxt), cardSN(0)) = 0 Then
    '                    For i = 0 To 64
    '                        Dbuffer(i) = 0
    '                    Next i
    '                    If MF_Read(0, blockNtxt, Bcount.Text, Dbuffer(0)) = 0 Then
    '                        For i = 0 To 64
    '                            BufferTXT = BufferTXT & Chr(Dbuffer(i))
    '                        Next i
    '                    Else
    '                        status = 0
    '                    End If
                    Else
                       Status = 0
                    End If
                Else
                    Status = 0
                End If
            Else
                Status = 0
            End If
        Else
            Status = 0
        End If
    Else
        Status = 0
    End If

    If Status = 0 Then
        ShowDisMessage "⁄œ„ ‘‰«”«∆Ì ﬂ«— ", 1000
        ReadCard = True  ' Save Continue
        Exit Function
    End If
    Dim L_Rst As New ADODB.Recordset
    Dim serial As String
    serial = CStr(Hex(cardSN(0)) & Hex(cardSN(1)) & Hex(cardSN(2)) & Hex(cardSN(3)))
    ReDim Parameter(0)
    Parameter(0) = GenerateInputParameter("@Serial", adVarWChar, 50, serial)
    
    Set L_Rst = RunParametricStoredProcedure2Rec("Check_RFIDSerialExist", Parameter)
    If Not (L_Rst.BOF Or L_Rst.EOF) Then
        ShowDisMessage "«Ì‰ ﬂ«—  ﬁ»·« »Â " & L_Rst!Name & " " & L_Rst!Family & L_Rst!WorkName & " »« «‘ —«ò " & L_Rst!MembershipId & " «Œ ’«’ œ«œÂ ‘œÂ «” ", 2000
        Set L_Rst = Nothing
        ReadCard = False
        Exit Function
    Else
        ShowDisMessage "ò«—  ‘‰«”«ÌÌ ‘œ", 1000
        TxtRfid = serial
        Set L_Rst = Nothing
        ReadCard = True
    End If
    
    Set L_Rst = Nothing
    
Exit Function
ErrHandler:
    MsgBox err.Description
    Set L_Rst = Nothing
End Function

Public Sub Edit()

    If OptionOwner(0).Value Then
        Me.OptionOwnerValue 0
    Else
        Me.OptionOwnerValue 1
    End If
    
    MyFormAddEditMode = EditMode
    SetFirstToolBar
    OldMembershipId = txtMembershipId.Text
    If OptionOwner(0).Value = True Then
        OldSwitchName = Trim(txtFamily.Text)
    Else
        OldSwitchName = Trim(txtWorkName.Text)
    End If
     OldTell = txtTel1.Text
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call PresetScreenSaver
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call PresetScreenSaver
End Sub


Public Sub ExitForm()
    Unload Me
    
    If Exit_Form = 0 Then
        mdifrm.Toolbar1.Buttons(20).Enabled = False
        mdifrm.Toolbar1.Buttons(21).Enabled = False
        mdifrm.Toolbar1.Buttons(23).Enabled = True
        mdifrm.Toolbar1.Buttons(24).Enabled = True
        mdifrm.Toolbar1.Buttons(25).Enabled = True
        mdifrm.Toolbar1.Buttons(26).Enabled = True
        mdifrm.Toolbar1.Buttons(27).Enabled = True
        varclick = False
        VarActForm = ""
    ElseIf Exit_Form = 2 Then
        If ClsFormAccess.frmInvoice = True Then
            frmInvoice.Show
            
            VarActForm = "frmInvoice"
            If FindCustFlag = True Then
                frmInvoice.FindCust
                FindCustFlag = False
            Else
                frmInvoice.SetFirstToolBar
            End If
        End If
    
    End If
   
    
    
End Sub

Public Sub Person()

'mdifrm.PicKeyBoard.Visible = True
If MyFormAddEditMode <> 1 Then
    mvarcode2 = txtCode.Text
    frmCustComp.Show
    ''mdifrm.Arrange 0
Else
   ' load frmMsg
    frmMsg.fwlblMsg.Caption = "«» œ« „‘ —Ì ›Êﬁ —« À»  Ê ”Å” Ê«—œ „—Õ·Â «⁄÷«¡ „‘ —ﬂ ‘ÊÌœ"
    frmMsg.fwBtn(0).Visible = False
    frmMsg.fwBtn(1).ButtonType = flwButtonOk
    frmMsg.fwBtn(1).Caption = "ﬁ»Ê·"
    frmMsg.Show vbModal
End If
End Sub

Private Sub FWBtnPerson_Click()
If MyFormAddEditMode <> AddMode And OptionOwner(1).Value = True Then
    mvarcode2 = txtCode.Text
    frmCustComp.Show
    frmCustComp.SetFocus
ElseIf MyFormAddEditMode = AddMode Then
    frmMsg.fwlblMsg.Caption = "«» œ« „‘ —Ì ›Êﬁ —« À»  Ê ”Å” Ê«—œ „—Õ·Â «⁄÷«¡ „‘ —ﬂ ‘ÊÌœ"
    frmMsg.fwBtn(0).Visible = False
    frmMsg.fwBtn(1).ButtonType = flwButtonOk
    frmMsg.fwBtn(1).Caption = "ﬁ»Ê·"
    frmMsg.Show vbModal
ElseIf OptionOwner(1).Value = False Then
    frmMsg.fwlblMsg.Caption = "»Â „‘ —Ì ›Êﬁ ‰„Ì  Ê«‰Ìœ «⁄÷«¡ «÷«›Â ‰„«ÌÌœ"
    frmMsg.fwBtn(0).Visible = False
    frmMsg.fwBtn(1).ButtonType = flwButtonOk
    frmMsg.fwBtn(1).Caption = "ﬁ»Ê·"
    frmMsg.Show vbModal
End If
End Sub


Private Sub FWBtnpicture_Click()
  If MyFormAddEditMode <> AddMode Then
      mvarcode2 = txtCode.Text
      mvarCustName = txtName.Text
      mvarCustFamily = txtFamily
      mvarMemberShipId = txtMembershipId.Text
      frmCustPicture.Show
      frmCustPicture.SetFocus
  ElseIf MyFormAddEditMode = AddMode Then
      frmMsg.fwlblMsg.Caption = "«» œ« „‘ —Ì ›Êﬁ —« À»  Ê ”Å” Ê«—œ „—Õ·Â À»   ’ÊÌ— ‘ÊÌœ"
      frmMsg.fwBtn(0).Visible = False
      frmMsg.fwBtn(1).ButtonType = flwButtonOk
      frmMsg.fwBtn(1).Caption = "ﬁ»Ê·"
      frmMsg.Show vbModal
  End If

End Sub

Private Sub FWBtnPrefix_Click()
    mvarcmbPrefix = True
   ' load frmCodingGeneral
'        frmCodingGeneral.SSTab.Tab = 10
'        frmCodingGeneral.Show
    
    ''mdifrm.Arrange 0
End Sub




Private Sub FWButton1_Click()

End Sub


'Private Sub FWBtnPrefix_GotFocus()
'    FWBtnPrefix.SetFocus
'    Set objName = FWBtnPrefix
'End Sub
'

Private Sub OptionOwner_Click(index As Integer)
    Me.OptionOwnerValue index
End Sub

Private Sub ResizeKit1_ExitResize(ByVal XScale As Double, ByVal YScale As Double)

    If formloadFlag = True Then
        If Me.Height < mdifrm.Height / 3 Then Me.Height = mdifrm.Height / 3
        SaveSetting strMainKey, Me.Name, "Height", Me.Height
        If Me.Width < mdifrm.Width / 3 Then Me.Width = mdifrm.Width / 3
        SaveSetting strMainKey, Me.Name, "Width", Me.Width
    End If

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
'             FWBtnPerson.Enabled = False
             framePerson.Visible = True
             frameCompany.Visible = False
        Case 1:
'             FWBtnPerson.Enabled = True
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
            If TypeOf Obj Is OptionButton Or TypeOf Obj Is FWCheck Or TypeOf Obj Is MaskEdBox Then
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
        TxtMemberCode.Enabled = True
'        fwlblMode.Caption = "„—Ê—"
        
    ElseIf MyFormAddEditMode = AddMode Then    'Add Mode
        
        On Error Resume Next
        For Each Obj In Me
            If TypeOf Obj Is OptionButton Or TypeOf Obj Is FWCheck Or TypeOf Obj Is MaskEdBox Then
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
    
    ElseIf MyFormAddEditMode = EditMode Then     'Edit
    
        On Error Resume Next
        For Each Obj In Me
            If TypeOf Obj Is OptionButton Or TypeOf Obj Is FWCheck Or TypeOf Obj Is MaskEdBox Then
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
'        fwlblMode.Caption = "«’·«Õ"
    
    End If
    
    HeaderLabel Val(MyFormAddEditMode), fwlblMode
    
End Sub
Sub GetDataDetail()
    
    On Error GoTo Err_Handler
    
    Dim L_Rst As New ADODB.Recordset
    
    DefaultSettings
    
    Dim TempStr As String
    
    ReDim Parameter(1) As Parameter
    Parameter(0) = GenerateInputParameter("@Code", adInteger, 4, Val(txtCode.Text))
    Parameter(1) = GenerateInputParameter("@Branch", adInteger, 4, CurrentBranch)
    Set L_Rst = RunParametricStoredProcedure2Rec("Get_Cust_info", Parameter)
    
    Dim ii As Integer
    
    If (L_Rst.BOF Or L_Rst.EOF) Then
        Set L_Rst = Nothing
        Exit Sub
    End If
    
'        OptionOwner(0) = Not (L_Rst!Owner)
'        OptionOwner(1) = Not (OptionOwner(0))
        If L_Rst!Owner = 0 Then
            OptionOwner(0).Value = True
        Else
            OptionOwner(1).Value = True
        End If
        
'        OptionActDeAct(0) = Not (L_Rst!ActDeAct)
'        OptionActDeAct(1) = Not (OptionActDeAct(0))
        If L_Rst!ActDeAct = True Then
            OptionActDeAct(0).Value = True
        Else
            OptionActDeAct(1).Value = True
        End If
        
        fwCheckAssansor.Value = L_Rst!Assansor
        txtMembershipId.Text = L_Rst!MembershipId
        txtName.Text = L_Rst!Name
        txtFamily.Text = L_Rst!Family
        txtTel1.Text = L_Rst!Tel1
        txtTel2.Text = L_Rst!Tel2
        txtTel3.Text = L_Rst!Tel3
        txtTel4.Text = L_Rst!Tel4
        '         txtTel5 = L_Rst!Tel5
        '         txtTel6 = L_Rst!Tel6
        txtWorkName.Text = L_Rst!WorkName
        txtFax.Text = L_Rst!Fax
        txtMobile = L_Rst!Mobile
        txtCredit.Text = L_Rst!Credit
        txtDiscount.Text = L_Rst!Discount
        txtEmail.Text = IIf(IsNull(L_Rst!Email), "", L_Rst!Email)
        txtCarryFee.Text = L_Rst!carryfee
        txtPaykFee.Text = L_Rst!PaykFee
        TxtEconomicalCode.Text = L_Rst!EconomicCode
        TxtRfid.Text = L_Rst!nvcRFID
        txtDescription.Text = IIf(IsNull(L_Rst!Description), "", L_Rst!Description)
        TxtAddress.Text = L_Rst!address
        ' txtDate , "Date"
        ' txtTime , "Time"
        ' txtUser , "User"
        txtTafsiliCode.Text = IIf(IsNull(L_Rst!Tafsili), "", L_Rst!Tafsili)
        TxtFamilyNo.Text = IIf(L_Rst!FamilyNo = 0, "", L_Rst!FamilyNo)
        ChkMember.Value = IIf(L_Rst!Member = True, Checked, Unchecked)
        ChkCentral.Value = IIf(L_Rst!Central = 1, 0, 1)
        OldTafsili = Val(txtTafsiliCode.Text)
        txtSanadNo.Text = IIf(IsNull(L_Rst!SanadNo), "", L_Rst!SanadNo)
 
        If IsNull(L_Rst!TotalRemainingAmount) = False Then
            If Val(L_Rst!TotalRemainingAmount) > 0 Then
                txtPrimaryBedehi.Text = Val(L_Rst!TotalRemainingAmount)
            Else
                txtPrimaryTalab.Text = -1 * Val(L_Rst!TotalRemainingAmount)
            End If
        End If
        For i = 0 To cmbActKind.ListCount - 1
            If cmbActKind.ItemData(i) = L_Rst!ActKind Then
                cmbActKind.ListIndex = i
                Exit For
            End If
        Next i
        
        For i = 0 To CmbState.ListCount - 1
            If CmbState.ItemData(i) = L_Rst!State Then
                CmbState.ListIndex = i
                Exit For
            End If
        Next i
        For i = 0 To cmbCity.ListCount - 1
            If cmbCity.ItemData(i) = L_Rst!City Then
                cmbCity.ListIndex = i
                Exit For
            End If
        Next i
        
        For i = 0 To cmbDistance.ListCount - 1
            If cmbDistance.ItemData(i) = L_Rst!Distance Then
                cmbDistance.ListIndex = i
                Exit For
            End If
        Next i
        
        For i = 0 To cmbPrefix.ListCount - 1
            If cmbPrefix.ItemData(i) = L_Rst!Prefix Then
                cmbPrefix.ListIndex = i
                Exit For
            End If
        Next i
        
        For i = 0 To cmbGender.ListCount - 1
            If cmbGender.ItemData(i) = L_Rst!Sex Then
                cmbGender.ListIndex = i
                Exit For
            End If
        Next i
        
        TempStr = ConvertToBin(L_Rst!BuyState, 4)
        For i = 0 To 3
            If Mid(TempStr, i + 1, 1) = 1 Then
                FWCheckBuy(i).Value = True
            Else
                FWCheckBuy(i).Value = False
            End If
        
        Next i
       
        For i = 0 To cmbSellPrice.ListCount - 1
            If cmbSellPrice.ItemData(i) = L_Rst!SellPrice Then
                cmbSellPrice.ListIndex = i
                Exit For
            End If
        Next i
        If L_Rst!nvcBirthDate <> "" Then
            txtBirthDate.Text = L_Rst!nvcBirthDate
        End If
    L_Rst.Close
    Set L_Rst = Nothing
    
'    TxtMemberCode.Text = txtMembershipId.Text
    Exit Sub
    
Err_Handler:
    ShowErrorMessage
    LogSaveNew "frmCust => ", err.Description, err.Number, err.Source, "GetDataDetail"
    Set L_Rst = Nothing
End Sub

Private Sub TxtFamilyNo_Change()
  If TxtFamilyNo.Text = "" Then Exit Sub
  TxtFamilyNo.Text = Val(TxtFamilyNo.Text)
  If TxtFamilyNo.Text = "0" Then TxtFamilyNo.Text = ""
End Sub

Private Sub TxtMemberCode_Change()
    Dim i As Long
    i = vsCustomer.FindRow(TxtMemberCode.Text, 1, 3, True, True)
    If i > 0 Then
        vsCustomer.Row = i
        vsCustomer.ShowCell i, 3
    Else
        vsCustomer.Row = 0
        vsCustomer.ShowCell 0, 0
    End If

End Sub

Private Sub TxtMemberCode_GotFocus()
    TxtMemberCode.Text = ""
End Sub

Private Sub vsCustomer_AfterSort(ByVal Col As Long, Order As Integer)
    Dim ExitFlag     As Boolean
    With vsCustomer
        If Col = 3 And .Rows > 1 Then
            For i = 1 To .Rows - 2
                If (Val(.TextMatrix(i, 3)) - Val(.TextMatrix(i + 1, 3)) > 1 And Order = 2) Or (Val(.TextMatrix(i + 1, 3)) - Val(.TextMatrix(i, 3)) > 1 And Order = 1) Then
                    .Cell(flexcpBackColor, i, .FixedCols, i, .Cols - 1) = 8421631
                    If Order = 1 And ExitFlag = False Then
                          txtMembershipId.Text = Val(.TextMatrix(i, 3)) + 1
                          ExitFlag = True
                    End If
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
        SaveSetting strMainKey, "frmCust_vsCustomer", "Col" & i, vsCustomer.ColWidth(i)
    Next

End Sub

Private Sub vsCustomer_Click()
    If clsStation.CustomerSearchDefault = False Then
        SetFirstToolBar
        HeaderLabel Val(MyFormAddEditMode), Me.fwlblMode
    End If
End Sub

Private Sub Insert_Tafsili(CustCode As Long, ShowMessageflag As Boolean)
    On Error GoTo ErrHandler
    Dim rs As New ADODB.Recordset
    Dim TafsiliName As String
    TafsiliName = Trim(txtName.Text) & " " & Trim(txtFamily.Text) & Trim(txtWorkName)
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
Private Sub vsCustomerClear()
    For i = 0 To vsCustomer.Rows - 2
        vsCustomer.RemoveItem
    Next i
End Sub

Private Sub FillAtf()
    txtAtf.Text = "«‘Œ«’ Ê ‘—ò Â«"
End Sub

Private Sub vsCustomer_RowColChange()
    txtCode.Text = vsCustomer.TextMatrix(vsCustomer.Row, 1)
    MyFormAddEditMode = ViewMode
    GetDataDetail

End Sub

