VERSION 5.00
Object = "{7AEDC602-D94C-11D1-BB7A-00E0290EA3C9}#1.0#0"; "ResizeKit.ocx"
Object = "{75D4F767-8785-11D3-93AD-0000832EF44D}#2.14#0"; "FAST2001.ocx"
Begin VB.Form frmPer 
   ClientHeight    =   9810
   ClientLeft      =   300
   ClientTop       =   420
   ClientWidth     =   8640
   Icon            =   "frmPer.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   9810
   ScaleWidth      =   8640
   Begin VB.Frame Frame4 
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
      Height          =   1125
      Left            =   240
      RightToLeft     =   -1  'True
      TabIndex        =   47
      Top             =   5760
      Width           =   8355
      Begin VB.TextBox txtAtf 
         Alignment       =   1  'Right Justify
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
         Left            =   5280
         RightToLeft     =   -1  'True
         TabIndex        =   53
         Top             =   360
         Width           =   2295
      End
      Begin VB.CommandButton cmdAddTafsili 
         Caption         =   "«÷«›Â ﬂ—œ‰ ﬂ·ÌÂ Å—”‰· ÃœÌœ »Â Õ”«»œ«—Ì"
         BeginProperty Font 
            Name            =   "Nazanin"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   2640
         RightToLeft     =   -1  'True
         TabIndex        =   51
         Top             =   240
         Width           =   2415
      End
      Begin VB.TextBox txtTafsiliCode 
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
         Left            =   240
         RightToLeft     =   -1  'True
         TabIndex        =   49
         Top             =   350
         Width           =   1215
      End
      Begin VB.Label lblTafsili 
         Alignment       =   1  'Right Justify
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
         Height          =   405
         Left            =   1350
         RightToLeft     =   -1  'True
         TabIndex        =   50
         Top             =   345
         Width           =   1035
      End
      Begin VB.Label lblAtf 
         Alignment       =   1  'Right Justify
         Caption         =   "⁄ÿ›"
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
         Left            =   7680
         RightToLeft     =   -1  'True
         TabIndex        =   48
         Top             =   345
         Width           =   495
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "‘⁄»Â"
      BeginProperty Font 
         Name            =   "B Homa"
         Size            =   11.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   850
      Left            =   240
      RightToLeft     =   -1  'True
      TabIndex        =   42
      Top             =   480
      Width           =   8355
      Begin VB.ComboBox cmbBranch 
         BackColor       =   &H8000000F&
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
         Left            =   120
         RightToLeft     =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   46
         Top             =   280
         Width           =   2235
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         Caption         =   "* Å—”‰·  ⁄—Ì› ‘œÂ œ— Â— ‘⁄»Â ﬁ«œ— »Â «‰ ﬁ«· »Â ‘⁄»Â Â«Ì œÌê— Â„ Â” ‰œ"
         BeginProperty Font 
            Name            =   "Nazanin"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   345
         Left            =   2400
         RightToLeft     =   -1  'True
         TabIndex        =   52
         Top             =   360
         Width           =   5805
      End
   End
   Begin VB.Frame Frame2 
      Height          =   4365
      Left            =   240
      RightToLeft     =   -1  'True
      TabIndex        =   20
      Top             =   1320
      Width           =   8355
      Begin VB.PictureBox Picture1 
         Height          =   480
         Left            =   4920
         RightToLeft     =   -1  'True
         ScaleHeight     =   420
         ScaleWidth      =   3195
         TabIndex        =   43
         Top             =   300
         Width           =   3255
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
            ForeColor       =   &H00000080&
            Height          =   375
            Index           =   0
            Left            =   360
            RightToLeft     =   -1  'True
            TabIndex        =   45
            Top             =   0
            Width           =   855
         End
         Begin VB.OptionButton OptionActDeAct 
            Alignment       =   1  'Right Justify
            Caption         =   "€Ì—›⁄«·"
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
            Height          =   450
            Index           =   1
            Left            =   1680
            RightToLeft     =   -1  'True
            TabIndex        =   44
            Top             =   0
            Width           =   1245
         End
      End
      Begin VB.TextBox txtMaxCredit 
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
         Height          =   480
         Left            =   840
         RightToLeft     =   -1  'True
         TabIndex        =   32
         Top             =   2670
         Width           =   1575
      End
      Begin VB.TextBox txtPersonnelNumber 
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
         Height          =   480
         Left            =   375
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   300
         Width           =   2055
      End
      Begin VB.TextBox txtIdNumber 
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
         Height          =   480
         Left            =   375
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   900
         Width           =   2055
      End
      Begin VB.TextBox txtInsuranceNo 
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
         Height          =   480
         Left            =   375
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   1485
         Width           =   2055
      End
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
         Height          =   735
         Left            =   360
         MultiLine       =   -1  'True
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   3360
         Width           =   6435
      End
      Begin VB.ComboBox cmbJob 
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
         Left            =   675
         RightToLeft     =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   2070
         Width           =   1755
      End
      Begin VB.TextBox txtFirstName 
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
         Height          =   480
         Left            =   4515
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   1485
         Width           =   2295
      End
      Begin VB.TextBox txtSurName 
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
         Height          =   480
         Left            =   4515
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   2070
         Width           =   2295
      End
      Begin VB.TextBox txtTel 
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
         Height          =   480
         Left            =   4515
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   2670
         Width           =   2295
      End
      Begin VB.ComboBox cmbGender 
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
         ItemData        =   "frmPer.frx":A4C2
         Left            =   5055
         List            =   "frmPer.frx":A4C4
         RightToLeft     =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   900
         Width           =   1755
      End
      Begin VB.Label lblMaxCredit 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H00000080&
         Height          =   405
         Left            =   2790
         RightToLeft     =   -1  'True
         TabIndex        =   31
         Top             =   2670
         Width           =   1035
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "* ‘„«—Â Å—”‰·Ì"
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
         Left            =   2430
         RightToLeft     =   -1  'True
         TabIndex        =   29
         Top             =   300
         Width           =   1395
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "‘„«—Â ‘‰«”‰«„Â"
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2430
         RightToLeft     =   -1  'True
         TabIndex        =   28
         Top             =   900
         Width           =   1395
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "‘„«—Â »Ì„Â"
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2430
         RightToLeft     =   -1  'True
         TabIndex        =   27
         Top             =   1485
         Width           =   1395
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "¬œ—”"
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   6720
         RightToLeft     =   -1  'True
         TabIndex        =   26
         Top             =   3480
         Width           =   1395
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         Caption         =   "‘€·"
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2430
         RightToLeft     =   -1  'True
         TabIndex        =   25
         Top             =   2070
         Width           =   1395
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "* ‰«„"
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
         Left            =   6720
         RightToLeft     =   -1  'True
         TabIndex        =   24
         Top             =   1485
         Width           =   1395
      End
      Begin VB.Label Label4 
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
         Left            =   6720
         RightToLeft     =   -1  'True
         TabIndex        =   23
         Top             =   2070
         Width           =   1395
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   " ·›‰ "
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   6720
         RightToLeft     =   -1  'True
         TabIndex        =   22
         Top             =   2715
         Width           =   1395
      End
      Begin VB.Label Label7 
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
         Height          =   405
         Left            =   6720
         RightToLeft     =   -1  'True
         TabIndex        =   21
         Top             =   900
         Width           =   1395
      End
   End
   Begin VB.CheckBox chkUser 
      Alignment       =   1  'Right Justify
      Caption         =   "ò«—»— ”Ì” „"
      BeginProperty Font 
         Name            =   "B Homa"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   345
      Left            =   7080
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   6840
      Width           =   1485
   End
   Begin VB.Frame Frame1 
      Height          =   2505
      Left            =   240
      RightToLeft     =   -1  'True
      TabIndex        =   14
      Top             =   7200
      Visible         =   0   'False
      Width           =   8385
      Begin VB.TextBox txtCountInvoiceEditable 
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
         Height          =   480
         Left            =   360
         RightToLeft     =   -1  'True
         TabIndex        =   38
         Top             =   1275
         Width           =   1395
      End
      Begin VB.TextBox txtCountInvoiceRefferable 
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
         Height          =   480
         Left            =   360
         RightToLeft     =   -1  'True
         TabIndex        =   37
         Top             =   1800
         Width           =   1395
      End
      Begin VB.TextBox txtCountInvoicePrint 
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
         Height          =   480
         Left            =   360
         RightToLeft     =   -1  'True
         TabIndex        =   36
         Top             =   765
         Width           =   1395
      End
      Begin VB.TextBox txtCountRePrint 
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
         Height          =   480
         Left            =   360
         RightToLeft     =   -1  'True
         TabIndex        =   35
         Top             =   240
         Width           =   1395
      End
      Begin VB.ComboBox cmbAccessLevel 
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
         Left            =   4320
         RightToLeft     =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   1800
         Width           =   2505
      End
      Begin VB.TextBox txtConfirm 
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
         Height          =   480
         IMEMode         =   3  'DISABLE
         Left            =   4350
         PasswordChar    =   "@"
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   1305
         Width           =   2475
      End
      Begin VB.TextBox txtPassword 
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
         Height          =   480
         IMEMode         =   3  'DISABLE
         Left            =   4350
         PasswordChar    =   "@"
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   765
         Width           =   2475
      End
      Begin VB.TextBox txtUserName 
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
         Height          =   480
         Left            =   4350
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   240
         Width           =   2475
      End
      Begin VB.Label lblCountInvoiceEditable 
         Alignment       =   1  'Right Justify
         Caption         =   " ⁄œ«œ›Ì‘ ﬁ«»· ÊÌ—«Ì‘"
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   525
         Left            =   1800
         RightToLeft     =   -1  'True
         TabIndex        =   40
         Top             =   1395
         Width           =   1965
      End
      Begin VB.Label lblCountInvoiceRefferable 
         Alignment       =   1  'Right Justify
         Caption         =   " ⁄œ«œ›Ì‘ ﬁ«»· „—ÃÊ⁄"
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   405
         Left            =   1800
         RightToLeft     =   -1  'True
         TabIndex        =   39
         Top             =   1920
         Width           =   1965
      End
      Begin VB.Label lblCountInvoicePrint 
         Alignment       =   1  'Right Justify
         Caption         =   " ⁄œ«œ ç«Å ›«ﬂ Ê— ›—Ê‘"
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   525
         Left            =   1800
         RightToLeft     =   -1  'True
         TabIndex        =   34
         Top             =   885
         Width           =   1965
      End
      Begin VB.Label lblCountRePrint 
         Alignment       =   1  'Right Justify
         Caption         =   " ⁄œ«œ ç«Å „Ãœœ"
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   525
         Left            =   1800
         RightToLeft     =   -1  'True
         TabIndex        =   33
         Top             =   360
         Width           =   1965
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         Caption         =   "*”ÿÕ œ” —”Ì"
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   465
         Left            =   6720
         RightToLeft     =   -1  'True
         TabIndex        =   19
         Top             =   1920
         Width           =   1455
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         Caption         =   "*  ò—«— ò·„Â ⁄»Ê—"
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   525
         Left            =   6630
         RightToLeft     =   -1  'True
         TabIndex        =   17
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         Caption         =   "* ‰«„ ò«—»—"
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   405
         Left            =   6960
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Top             =   360
         Width           =   1245
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         Caption         =   "* ò·„Â ⁄»Ê—"
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   405
         Left            =   6960
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   840
         Width           =   1245
      End
   End
   Begin FLWCtrls.FWLabel3D fwlblMode 
      Height          =   525
      Left            =   7200
      Top             =   0
      Width           =   1545
      _ExtentX        =   2725
      _ExtentY        =   926
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Titr"
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
      OleObjectBlob   =   "frmPer.frx":A4C6
      TabIndex        =   41
      Top             =   0
      Width           =   480
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   " ⁄—Ì› Å—”‰·"
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
      Height          =   450
      Left            =   3360
      RightToLeft     =   -1  'True
      TabIndex        =   30
      Top             =   0
      Width           =   2175
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      Caption         =   "*  ò„Ì· «ÿ·«⁄«  »—«Ì ⁄‰«ÊÌ‰ ” «—Â œ«— «Ã»«—Ì „Ì »«‘œ"
      BeginProperty Font 
         Name            =   "Nazanin"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   345
      Left            =   1200
      RightToLeft     =   -1  'True
      TabIndex        =   18
      Top             =   6840
      Width           =   4245
   End
End
Attribute VB_Name = "frmPer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private MyFormAddEditMode As EnumAddEditMode
Dim cn As New ADODB.Connection
Dim Parameter() As Parameter
Dim CurrentAccessIndex As Integer
Dim OldTafsili As Long

Public Sub ChangeLanguage()
    
    DefaultSetting
    
End Sub

Public Sub Add()
    cmbBranch.Enabled = True
    MyFormAddEditMode = AddMode
    DefaultSetting
    SetFirstToolBar
    
End Sub
Public Sub Find()
        
        frmFindPerson.Show vbModal
        
        If mvarcode <> 0 Then
            For i = 0 To cmbBranch.ListCount - 1
                cmbBranch.ListIndex = i
                If mvarBranch = cmbBranch.ItemData(cmbBranch.ListIndex) Then
                    Exit For
                End If
            Next
            
            Dim Rst As New ADODB.Recordset
            
            ReDim Parameter(1) As Parameter
            Parameter(0) = GenerateInputParameter("@CurrentpPno", adInteger, 4, mvarcode)
            Parameter(1) = GenerateInputParameter("@Branch", adInteger, 4, cmbBranch.ItemData(cmbBranch.ListIndex))
            Set Rst = RunParametricStoredProcedure2Rec("Get_PersonelInfo", Parameter)
            If Rst.State = 1 Then
                If Not (Rst.EOF = True And Rst.BOF = True) Then
                    GetRecrdsetDetail Rst
                End If
            End If
            Set Rst = Nothing
    
            MyFormAddEditMode = ViewMode
            SetFirstToolBar
            mvarcode = 0
        
        Else
            Exit Sub
            
        End If
    
End Sub

Public Sub Printing()

    ReDim Parameter(0) As Parameter
    
    Dim MyCrystalReport
    Set MyCrystalReport = CreateObject("Crystal.CrystalReport")
    
    Parameter(0) = GenerateInputParameter("@PPNO", adInteger, 4, txtPersonnelNumber.Tag)
    
    MyCrystalReport.ReportFileName = App.Path & "\Reports" & RepVer & "\RepIdCard.rpt"
    MyCrystalReport.Destination = crptToPrinter 'crptToWindow ' '
    MyCrystalReport.ParameterFields(0) = CStr(Parameter(0).Name) & ";" & CStr(Parameter(0).Value) & ";" & "True"
    
    MyCrystalReport.RetrieveDataFiles
    ODBCSetting clsArya.ServerName, clsArya.DbName
    MyCrystalReport.Connect = CrystallConnection
    MyCrystalReport.Action = 1
    MyCrystalReport.PageZoom (150)
   
   
End Sub
Public Sub Cancel()

    MyFormAddEditMode = AddMode
    SetFirstToolBar
    DefaultSetting
    
End Sub

Public Sub DefaultSetting()

    Dim Rst As New ADODB.Recordset
    ReDim Parameter(0) As Parameter
    
    Parameter(0) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
    
    
    If Rst.State <> 0 Then Rst.Close
    
    Set Rst = RunParametricStoredProcedure2Rec("GetPostInfo", Parameter)
    cmbJob.Clear
    If Not (Rst.EOF = True And Rst.BOF = True) Then
        i = 1
        While Rst.EOF <> True
            cmbJob.AddItem Rst.Fields("Description").Value
            cmbJob.ItemData(cmbJob.ListCount - 1) = Rst.Fields("code").Value
            Rst.MoveNext
        Wend
    End If
    cmbJob.ListIndex = 0
    
    
     CurrentAccessIndex = -1
    If Rst.State <> 0 Then Rst.Close
    
    Set Rst = RunParametricStoredProcedure2Rec("GetAccessLevel", Parameter)
    cmbAccessLevel.Clear
    If Not (Rst.EOF = True And Rst.BOF = True) Then
        i = 1
        While Rst.EOF <> True
            cmbAccessLevel.AddItem Rst.Fields("Description").Value
            cmbAccessLevel.ItemData(cmbAccessLevel.ListCount - 1) = Rst.Fields("intAccessLevel").Value
            Rst.MoveNext
        Wend
    End If
    cmbAccessLevel.ListIndex = 0
    
''''    Set rctmp = RunStoredProcedure2RecordSet("Get_New_Per_PersonnelNumber", cnn)
''''    txtPersonnelNumber.Text = rctmp.Fields("Code").Value
''''    txtPersonnelNumber.Tag = rctmp.Fields("Code").Value
    
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
     
    Dim Obj As Object
    For Each Obj In Me
        If TypeOf Obj Is TextBox Then
            Obj.Text = ""
            Obj.Tag = 0
        End If
    Next Obj
    
    ReDim Parameter(1) As Parameter
    Parameter(0) = GenerateInputParameter("@CurrentpPno", adInteger, 4, 0)
    Parameter(1) = GenerateInputParameter("@Branch", adInteger, 4, cmbBranch.ItemData(cmbBranch.ListIndex))
    Set Rst = RunParametricStoredProcedure2Rec("Get_PersonelInfo", Parameter)
    If Not (Rst.EOF = True And Rst.BOF = True) Then
        txtPersonnelNumber.Tag = Rst!ppno
        txtPersonnelNumber.Text = Rst!PersonnelNumber
    End If
    chkUser.Value = 0
    txtTafsiliCode.Text = ""
    OldTafsili = 0
    OldAtf = 0
    txtMaxCredit.Visible = False
    lblMaxCredit.Visible = False
    OptionActDeAct(0).Value = True
End Sub

Public Sub Edit()
    ''cmbBranch.Enabled = False
    MyFormAddEditMode = EditMode
    SetFirstToolBar
End Sub

Private Sub GetRecrdsetDetail(tempRst As ADODB.Recordset)

    DefaultSetting
    
    If tempRst.EOF = True And tempRst.BOF = True Then Exit Sub
    
    txtMaxCredit.Text = IIf(IsNull(tempRst.Fields("Maxcredit").Value), " ", tempRst.Fields("Maxcredit").Value)
    txtPersonnelNumber.Tag = tempRst.Fields("pPno").Value
    txtPersonnelNumber.Text = tempRst.Fields("PersonnelNumber").Value
    txtFirstName.Text = tempRst.Fields("nvcFirstName").Value
    txtSurName.Text = tempRst.Fields("nvcSurName").Value
    txtIdNumber.Text = IIf(IsNull(tempRst.Fields("IdNumber").Value), 0, tempRst.Fields("IdNumber").Value)
    txtInsuranceNo.Text = IIf(IsNull(tempRst.Fields("InsuranceNo").Value), 0, tempRst.Fields("InsuranceNo").Value)
    TxtAddress.Text = tempRst.Fields("Address").Value
    txtTel.Text = IIf(IsNull(tempRst.Fields("Tel").Value), "", tempRst.Fields("Tel").Value)
    txtTafsiliCode.Text = IIf(IsNull(tempRst.Fields("Tafsili").Value), "", tempRst.Fields("Tafsili").Value)
    OldTafsili = Val(txtTafsiliCode.Text)
    
    If tempRst!ActDeAct = True Then
            OptionActDeAct(0).Value = True
        Else
            OptionActDeAct(1).Value = True
    End If
    If IsNull(tempRst.Fields("UID").Value) = False Then
        chkUser.Value = 1
        TxtUserName.Text = tempRst.Fields("Username").Value
        TxtUserName.Tag = tempRst.Fields("UID").Value
        txtPassword.Text = tempRst.Fields("Password").Value
        txtConfirm.Text = txtPassword.Text
        txtCountRePrint.Text = tempRst.Fields("CountRePrint").Value
        txtCountInvoicePrint.Text = tempRst.Fields("CountInvoicePrint").Value
        txtCountInvoiceEditable = tempRst.Fields("CountInvoiceEditable").Value
        txtCountInvoiceRefferable.Text = tempRst.Fields("CountInvoiceRefferable").Value
        Frame1.Enabled = False
        Frame1.Visible = True
       If IsNull(tempRst.Fields("intAccesslevel").Value) = False Then
          cmbAccessLevel.ListIndex = tempRst.Fields("intAccesslevel").Value - 1
       End If
    Else
        chkUser.Value = 0
    End If
    
    For i = 0 To cmbGender.ListCount - 1
        If tempRst.Fields("Gender").Value = cmbGender.ItemData(i) Then
            cmbGender.ListIndex = i
            Exit For
        End If
    Next i
    
    For i = 0 To cmbJob.ListCount - 1
        If tempRst.Fields("job").Value = cmbJob.ItemData(i) Then
            cmbJob.ListIndex = i
            Exit For
        End If
    Next i
    
End Sub

Public Sub ExitForm()

    Unload Me
    
End Sub

Public Sub FirstKey()
    cmbBranch.Enabled = True
    ReDim Parameter(3) As Parameter
    Dim Rst As New ADODB.Recordset
    
    Parameter(0) = GenerateInputParameter("@CurrentpPno", adInteger, 4, Val(txtPersonnelNumber.Tag))
    Parameter(1) = GenerateInputParameter("@Direction", adInteger, 4, EnumDirection.FirstRecord)
    Parameter(2) = GenerateInputParameter("@AccessLevelCurrentUser", adInteger, 4, mVarAccessLevel)
    Parameter(3) = GenerateInputParameter("@Branch", adInteger, 4, cmbBranch.ItemData(cmbBranch.ListIndex))
    Set Rst = RunParametricStoredProcedure2Rec("NavigateInPersonel", Parameter)
    If Rst.State = 1 Then
        If Not (Rst.EOF = True And Rst.BOF = True) Then
            GetRecrdsetDetail Rst
        End If
    End If
    Set Rst = Nothing
    
    MyFormAddEditMode = ViewMode
    SetFirstToolBar
End Sub

Public Sub PreviousKey()
    cmbBranch.Enabled = True
    ReDim Parameter(3) As Parameter
    Dim Rst As New ADODB.Recordset
    
    Parameter(0) = GenerateInputParameter("@CurrentpPno", adInteger, 4, Val(txtPersonnelNumber.Tag))
    Parameter(1) = GenerateInputParameter("@Direction", adInteger, 4, EnumDirection.PreviousRecord)
    Parameter(2) = GenerateInputParameter("@AccessLevelCurrentUser", adInteger, 4, mVarAccessLevel)
    Parameter(3) = GenerateInputParameter("@Branch", adInteger, 4, cmbBranch.ItemData(cmbBranch.ListIndex))
    Set Rst = RunParametricStoredProcedure2Rec("NavigateInPersonel", Parameter)
    If Rst.State = 1 Then
        If Not (Rst.EOF = True And Rst.BOF = True) Then
            GetRecrdsetDetail Rst
        End If
    End If
    Set Rst = Nothing
    
    MyFormAddEditMode = ViewMode
    SetFirstToolBar
End Sub

Public Sub NextKey()
    cmbBranch.Enabled = True
    ReDim Parameter(3) As Parameter
    Dim Rst As New ADODB.Recordset
    
    Parameter(0) = GenerateInputParameter("@CurrentpPno", adInteger, 4, Val(txtPersonnelNumber.Tag))
    Parameter(1) = GenerateInputParameter("@Direction", adInteger, 4, EnumDirection.NextRecord)
    Parameter(2) = GenerateInputParameter("@AccessLevelCurrentUser", adInteger, 4, mVarAccessLevel)
    Parameter(3) = GenerateInputParameter("@Branch", adInteger, 4, cmbBranch.ItemData(cmbBranch.ListIndex))
    Set Rst = RunParametricStoredProcedure2Rec("NavigateInPersonel", Parameter)
    If Rst.State = 1 Then
        If Not (Rst.EOF = True And Rst.BOF = True) Then
            GetRecrdsetDetail Rst
        End If
    End If
    Set Rst = Nothing
    
    MyFormAddEditMode = ViewMode
    SetFirstToolBar
End Sub

Public Sub LastKey()
    cmbBranch.Enabled = True
    ReDim Parameter(3) As Parameter
    Dim Rst As New ADODB.Recordset
    
    Parameter(0) = GenerateInputParameter("@CurrentpPno", adInteger, 4, Val(txtPersonnelNumber.Tag))
    Parameter(1) = GenerateInputParameter("@Direction", adInteger, 4, EnumDirection.LastRecord)
    Parameter(2) = GenerateInputParameter("@AccessLevelCurrentUser", adInteger, 4, mVarAccessLevel)
    Parameter(3) = GenerateInputParameter("@Branch", adInteger, 4, cmbBranch.ItemData(cmbBranch.ListIndex))
    Set Rst = RunParametricStoredProcedure2Rec("NavigateInPersonel", Parameter)
    If Rst.State = 1 Then
        If Not (Rst.EOF = True And Rst.BOF = True) Then
            GetRecrdsetDetail Rst
        End If
    End If
    Set Rst = Nothing
    
    MyFormAddEditMode = ViewMode
    SetFirstToolBar
End Sub


Public Sub SetFirstToolBar()
    
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
    
        mdifrm.Toolbar1.Buttons(6).Enabled = True  'Add
        mdifrm.Toolbar1.Buttons(7).Enabled = True  'Edit
        mdifrm.Toolbar1.Buttons(8).Enabled = False  'Enter
        mdifrm.Toolbar1.Buttons(9).Enabled = False   'Esc
        Frame1.Enabled = False
        Frame2.Enabled = False
        txtTafsiliCode.Enabled = False
        For Each Obj In Me
             If TypeOf Obj Is OptionButton Or TypeOf Obj Is FWCheck Then
                 Obj.Enabled = False
              End If
        Next Obj
          
         
    ElseIf MyFormAddEditMode = AddMode Then    'Add Mode
    
        mdifrm.Toolbar1.Buttons(6).Enabled = False  'Add
        mdifrm.Toolbar1.Buttons(7).Enabled = False  'Edit
        mdifrm.Toolbar1.Buttons(8).Enabled = True   'Enter
        mdifrm.Toolbar1.Buttons(9).Enabled = True   'Esc
        Frame1.Enabled = True
        Frame2.Enabled = True
        For Each Obj In Me
            If TypeOf Obj Is OptionButton Or TypeOf Obj Is FWCheck Then
                Obj.Enabled = True
            End If
        Next Obj
       
    ElseIf MyFormAddEditMode = EditMode Then     'Edit
        
        mdifrm.Toolbar1.Buttons(6).Enabled = True  'Add
        mdifrm.Toolbar1.Buttons(7).Enabled = False  'Edit
        mdifrm.Toolbar1.Buttons(8).Enabled = True   'Enter
        mdifrm.Toolbar1.Buttons(9).Enabled = True   'Esc
        Frame1.Enabled = True
        Frame2.Enabled = True
        For Each Obj In Me
            If TypeOf Obj Is OptionButton Or TypeOf Obj Is FWCheck Then
                Obj.Enabled = True
            End If
        Next Obj
        
    End If
    
    HeaderLabel Val(MyFormAddEditMode), fwlblMode

End Sub
Public Sub Update()
    
    Dim Result As Long
    Dim Obj As Object
    
    If Trim(txtPersonnelNumber.Text) = "" Or Trim(txtFirstName.Text) = "" Or Trim(txtSurName.Text) = "" Then
            
            frmMsg.fwlblMsg.Caption = "·ÿ›« «ÿ·«⁄«  ÷—Ê—Ì —« Ê«—œ ‰„«ÌÌœ"
            frmMsg.fwBtn(0).ButtonType = flwButtonOk
            frmMsg.fwBtn(0).Caption = "ﬁ»Ê·"
            frmMsg.Show vbModal
            Exit Sub

    ElseIf chkUser.Value = 1 Then
    
        If mVarAccessLevel > cmbAccessLevel.ItemData(cmbAccessLevel.ListIndex) Then
            frmMsg.fwlblMsg.Caption = "”ÿÕ œ” —”Ì «‰ Œ«» ‘œÂ »«·« — «“ œ” —”Ì ›⁄·Ì „Ì »«‘œ. À»  «‰Ã«„ ‰‘œ"
            frmMsg.fwBtn(0).ButtonType = flwButtonOk
            frmMsg.fwBtn(0).Caption = "ﬁ»Ê·"
            frmMsg.Show vbModal
            Exit Sub
        End If
        Select Case MyFormAddEditMode
            Case AddMode
                
                If TxtUserName.Text = "" Or txtPassword.Text = "" Or txtConfirm.Text = "" Then
                    frmMsg.fwlblMsg.Caption = "·ÿ›« «ÿ·«⁄«  ÷—Ê—Ì —« Ê«—œ ‰„«ÌÌœ"
                    frmMsg.fwBtn(0).ButtonType = flwButtonOk
                    frmMsg.fwBtn(0).Caption = "ﬁ»Ê·"
                    frmMsg.Show vbModal
                    Exit Sub
                End If
            
            Case Else
                If TxtUserName.Text = "" Or txtPassword.Text = "" Then
                    frmMsg.fwlblMsg.Caption = "·ÿ›« «ÿ·«⁄«  ÷—Ê—Ì —« Ê«—œ ‰„«ÌÌœ"
                    frmMsg.fwBtn(0).ButtonType = flwButtonOk
                    frmMsg.fwBtn(0).Caption = "ﬁ»Ê·"
                    frmMsg.Show vbModal
                    Exit Sub
                End If
        End Select
        If txtPassword.Text <> txtConfirm.Text Then
            frmMsg.fwlblMsg.Caption = " ﬂ—«— —„“ »« —„“ Â„ŒÊ«‰Ì ‰œ«—œ"
            frmMsg.fwBtn(0).ButtonType = flwButtonOk
            frmMsg.fwBtn(0).Caption = "ﬁ»Ê·"
            frmMsg.Show vbModal
            Exit Sub
        End If
    End If
    
     If Trim(txtMaxCredit.Text) = "" And txtMaxCredit.Visible = True Then
            frmMsg.fwlblMsg.Caption = "·ÿ›« «ÿ·«⁄«  ÷—Ê—Ì —« Ê«—œ ‰„«ÌÌœ"
            frmMsg.fwBtn(0).ButtonType = flwButtonOk
            frmMsg.fwBtn(0).Caption = "ﬁ»Ê·"
            frmMsg.Show vbModal
            Exit Sub
     End If
    
    ReDim Parameter(20) As Parameter
    Select Case MyFormAddEditMode
        Case AddMode
        
            
            Parameter(0) = GenerateInputParameter("@PersonnelNumber", adVarChar, 50, Trim(txtPersonnelNumber.Text))
            Parameter(1) = GenerateInputParameter("@nvcFirstName", adVarChar, 50, Trim(txtFirstName.Text))
            Parameter(2) = GenerateInputParameter("@nvcSurName", adVarChar, 50, Trim(txtSurName.Text))
            Parameter(3) = GenerateInputParameter("@Gender", adBoolean, 1, cmbGender.ItemData(cmbGender.ListIndex))
            Parameter(4) = GenerateInputParameter("@IdNumber", adVarChar, 50, Trim(txtIdNumber.Text))
            Parameter(5) = GenerateInputParameter("@Job", adInteger, 4, cmbJob.ItemData(cmbJob.ListIndex))
            Parameter(6) = GenerateInputParameter("@InsuranceNo", adVarChar, 50, Trim(txtInsuranceNo.Text))
            Parameter(7) = GenerateInputParameter("@Address", adVarChar, 300, Trim(TxtAddress.Text))
            Parameter(8) = GenerateInputParameter("@Tel", adVarChar, 30, Trim(txtTel.Text))
            Parameter(9) = GenerateInputParameter("@User", adInteger, 4, mvarCurUserNo)
            Parameter(13) = GenerateInputParameter("@MaxCredit", adInteger, 4, Val(Trim(txtMaxCredit.Text)))
            Parameter(14) = GenerateInputParameter("@ActDeAct", adInteger, 4, IIf(OptionActDeAct(0).Value = True, 1, 0))
            Parameter(15) = GenerateInputParameter("@CountRePrint", adInteger, 4, Val(Trim(txtCountRePrint.Text)))
            Parameter(16) = GenerateInputParameter("@CountInvoicePrint", adInteger, 4, Val(Trim(txtCountInvoicePrint.Text)))
            Parameter(17) = GenerateInputParameter("@CountInvoiceEditable", adInteger, 4, Val(Trim(txtCountInvoiceEditable.Text)))
            Parameter(18) = GenerateInputParameter("@CountInvoiceRefferable", adInteger, 4, Val(Trim(txtCountInvoiceRefferable.Text)))
            Parameter(19) = GenerateInputParameter("@Branch", adInteger, 4, cmbBranch.ItemData(cmbBranch.ListIndex))
            Parameter(20) = GenerateOutputParameter("@pPno", adInteger, 4)
            
            If chkUser.Value = 1 Then
                Parameter(10) = GenerateInputParameter("@UserName", adVarChar, 50, TxtUserName.Text)
                Parameter(11) = GenerateInputParameter("@Password", adVarChar, 50, txtPassword.Text)
                Parameter(12) = GenerateInputParameter("@intAccessLevel", adInteger, 4, cmbAccessLevel.ItemData(cmbAccessLevel.ListIndex))
            Else
                Parameter(10) = GenerateInputParameter("@UserName", adVarChar, 50, "")
                Parameter(11) = GenerateInputParameter("@Password", adVarChar, 50, "")
                Parameter(12) = GenerateInputParameter("@intAccessLevel", adInteger, 4, 0)
            End If
            On Error GoTo ErrHandler
            
            Result = RunParametricStoredProcedure("InsertPersonel", Parameter)
            If Result <= 0 Then GoTo ErrHandler
            
            frmMsg.fwlblMsg.Caption = "À»  «ÿ·«⁄«  ÃœÌœ »« „Ê›ﬁÌ  Å«Ì«‰ Ì«› "
            frmMsg.fwBtn(0).ButtonType = flwButtonOk
            frmMsg.fwBtn(0).Caption = "ﬁ»Ê·"
            frmMsg.fwBtn(1).Visible = False
            frmMsg.Show vbModal
            
            If clsArya.ExternalAccounting = True Or HasMiniAcc = True Then
                Insert_Tafsili Result, True
            End If
            Add
            
        Case EditMode
            Dim Parameter2(22) As Parameter
            
            Parameter2(0) = GenerateInputParameter("@CurrentPPNO", adInteger, 50, Val(txtPersonnelNumber.Tag))
            Parameter2(1) = GenerateInputParameter("@PersonnelNumber", adVarChar, 50, Trim(txtPersonnelNumber.Text))
            Parameter2(2) = GenerateInputParameter("@nvcFirstName", adVarChar, 50, Trim(txtFirstName.Text))
            Parameter2(3) = GenerateInputParameter("@nvcSurName", adVarChar, 50, Trim(txtSurName.Text))
            Parameter2(4) = GenerateInputParameter("@Gender", adBoolean, 1, cmbGender.ItemData(cmbGender.ListIndex))
            Parameter2(5) = GenerateInputParameter("@IdNumber", adVarChar, 50, Trim(txtIdNumber.Text))
            Parameter2(6) = GenerateInputParameter("@Job", adInteger, 4, cmbJob.ItemData(cmbJob.ListIndex))
            Parameter2(7) = GenerateInputParameter("@InsuranceNo", adVarChar, 50, Trim(txtInsuranceNo.Text))
            Parameter2(8) = GenerateInputParameter("@Address", adVarChar, 300, Trim(TxtAddress.Text))
            Parameter2(9) = GenerateInputParameter("@Tel", adVarChar, 30, Trim(txtTel.Text))
            Parameter2(10) = GenerateInputParameter("@User", adInteger, 4, mvarCurUserNo)
            Parameter2(15) = GenerateInputParameter("@MaxCredit", adInteger, 4, Val(Trim(txtMaxCredit.Text)))
            Parameter2(16) = GenerateInputParameter("@ActDeAct", adInteger, 4, IIf(OptionActDeAct(0).Value = True, 1, 0))
            Parameter2(17) = GenerateInputParameter("@CountRePrint", adInteger, 4, Val(Trim(txtCountRePrint.Text)))
            Parameter2(18) = GenerateInputParameter("@CountInvoicePrint", adInteger, 4, Val(Trim(txtCountInvoicePrint.Text)))
            Parameter2(19) = GenerateInputParameter("@CountInvoiceEditable", adInteger, 4, Val(Trim(txtCountInvoiceEditable.Text)))
            Parameter2(20) = GenerateInputParameter("@CountInvoiceRefferable", adInteger, 4, Val(Trim(txtCountInvoiceRefferable.Text)))
            Parameter2(21) = GenerateInputParameter("@Branch", adInteger, 4, cmbBranch.ItemData(cmbBranch.ListIndex))
            Parameter2(22) = GenerateOutputParameter("@pPno", adInteger, 4)
            
            If chkUser.Value = 1 Then
                Parameter2(11) = GenerateInputParameter("@UID", adInteger, 4, Val(TxtUserName.Tag))
                Parameter2(12) = GenerateInputParameter("@UserName", adVarChar, 50, TxtUserName.Text)
                Parameter2(13) = GenerateInputParameter("@Password", adVarChar, 50, txtPassword.Text)
                Parameter2(14) = GenerateInputParameter("@intAccessLevel", adInteger, 4, cmbAccessLevel.ItemData(cmbAccessLevel.ListIndex))
            Else
                Parameter2(11) = GenerateInputParameter("@UID", adInteger, 4, 0)
                Parameter2(12) = GenerateInputParameter("@UserName", adVarChar, 50, "")
                Parameter2(13) = GenerateInputParameter("@Password", adVarChar, 50, "")
                Parameter2(14) = GenerateInputParameter("@intAccessLevel", adInteger, 4, 0)
            End If
            
            On Error GoTo ErrHandler
            Result = RunParametricStoredProcedure("UpdatePersonel", Parameter2)
            If Result <= 0 Then GoTo ErrHandler
            
            frmMsg.fwlblMsg.Caption = " €ÌÌ— «ÿ·«⁄«  »« „Ê›ﬁÌ  Å«Ì«‰ Ì«› "
            frmMsg.fwBtn(0).ButtonType = flwButtonOk
            frmMsg.fwBtn(0).Caption = "ﬁ»Ê·"
            frmMsg.fwBtn(1).Visible = False
            frmMsg.Show vbModal
            
            If clsArya.ExternalAccounting = True Or HasMiniAcc = True Then
                Insert_Tafsili Result, True
            End If
            
            Add
    End Select

Exit Sub
ErrHandler:
    If err.Number = -2147217873 Or Result <= 0 Then
        frmMsg.fwlblMsg.Caption = "À»  «‰Ã«„ ‰‘œ" + vbCrLf + "«ÿ·«⁄«   ò—«—Ì „Ì »«‘œ"
        frmMsg.fwBtn(0).ButtonType = flwButtonOk
        frmMsg.fwBtn(0).Caption = "ﬁ»Ê·"
        frmMsg.fwBtn(1).Visible = False
        frmMsg.Show vbModal
     Else
        'MsgBox Err.Description
     End If
    
End Sub

Private Sub chkUser_Click()

    Select Case MyFormAddEditMode
    
        Case ViewMode
        
            If chkUser.Value = 1 Then
                Frame1.Visible = True
            ElseIf chkUser.Value = 0 Then
                Frame1.Visible = False
            End If
            
        Case Else
        
            If chkUser.Value = 1 Then
                Frame1.Visible = True
            ElseIf chkUser.Value = 0 Then
                Frame1.Visible = False
            End If
            
    End Select
    
End Sub

Private Sub cmbAccessLevel_Click()
    If cmbAccessLevel.ListIndex > -1 And MyFormAddEditMode = EditMode Then
        If CurrentAccessIndex <> -1 And cmbAccessLevel.ItemData(cmbAccessLevel.ListIndex) < mVarAccessLevel Then
            cmbAccessLevel.ListIndex = CurrentAccessIndex
            frmMsg.fwlblMsg.Caption = "‘„« ‰„Ì  Ê«‰Ìœ «Ì‰ ”ÿÕ œ” —”Ì —«  €ÌÌ— œÂÌœ"
            frmMsg.fwBtn(1).Visible = False
            frmMsg.Show vbModal
        End If
    End If
End Sub
Private Sub cmbAccessLevel_GotFocus()
    CurrentAccessIndex = cmbAccessLevel.ListIndex
End Sub

Private Sub cmbBranch_Click()
    'DefaultSetting
End Sub

Private Sub cmbJob_Click()
    If cmbJob.ItemData(cmbJob.ListIndex) = 3 Then
        txtMaxCredit.Visible = True
        lblMaxCredit.Visible = True
    Else
        txtMaxCredit.Visible = False
        lblMaxCredit.Visible = False
    End If
End Sub



Private Sub cmdAddTafsili_Click()
    ShowMessage "¬Ì« „Ì ŒÊ«ÂÌœ »—«Ì ﬂ·ÌÂ Å—”‰· '  ›÷Ì·Ì ÃœÌœ œ— ”Ì” „ Õ”«»œ«—Ì «ÌÕ«œ ò‰Ìœø ", True, True, "»·Ì", "ŒÌ—"
    If modgl.mvarMsgIdx = vbNo Then
        Exit Sub
    End If
    
    ReDim Parameter(1) As Parameter
    Dim Rst As New ADODB.Recordset
    
    Parameter(0) = GenerateInputParameter("@AccessLevelCurrentUser", adInteger, 4, 1)
    Parameter(1) = GenerateInputParameter("@Branch", adInteger, 4, cmbBranch.ItemData(cmbBranch.ListIndex))
    Set Rst = RunParametricStoredProcedure2Rec("Get_tblTotal_tPer_ByAccessLevel", Parameter)
    
    While Rst.EOF <> True
        txtTafsiliCode.Text = ""
        txtFirstName.Text = ""
        txtSurName.Text = Rst!PersonName
        If IsNull(Rst!Tafsili) = True Or Trim(Rst!Tafsili) = "" Then Insert_Tafsili Rst!ppno, False
        Rst.MoveNext
    Wend
    If Rst.State = 1 Then Rst.Close
    If cn.State = 1 Then cn.Close
    Set Rst = Nothing
    Set cn = Nothing
    
    ShowDisMessage " ⁄—Ì› Å—”‰· œ— ”Ì” „ Õ”«»œ«—Ì «‰Ã«„ ê—œÌœ", 1000
    DefaultSetting

End Sub

Private Sub Form_Activate()
    VarActForm = Me.Name
    SetFirstToolBar
    
    If TempPerFlag = True Then
    End If
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call PresetScreenSaver
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call PresetScreenSaver
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case Shift
        Case 0
            Select Case KeyCode
                
                Case 33
                    NextKey
                Case 34
                    PreviousKey
                Case 35
                    LastKey
                Case 36
                    FirstKey
            End Select
    
    End Select
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

Private Sub Form_Load()
    If TempPerFlag = True Then
        If ClsFormAccess.frmPer = False Then
            Unload Me
            Exit Sub
        End If
'        Dim obj As Object
'        For Each obj In Forms
'            If TypeOf obj Is Form Then
'                If obj.Name <> "mdifrm" And obj.Name <> Me.Name And obj.Name <> "frmAbout" Then
'                    obj.Hide
'                End If
'            End If
'
'        Next obj
    
    End If
    
    CenterTop Me
    
    VarActForm = Me.Name
    
    MyFormAddEditMode = AddMode
    FillBranch
    DefaultSetting
    If TempPerFlag = True Then
        SetFirstToolBar
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


End Sub
Private Sub FillBranch()
    
    cmbBranch.Clear
    Set rctmp = RunStoredProcedure2RecordSet("Get_All_Branches")
    Do While rctmp.EOF = False
        cmbBranch.AddItem rctmp!nvcBranchName
        cmbBranch.ItemData(cmbBranch.NewIndex) = rctmp!Branch
        rctmp.MoveNext
    Loop
    rctmp.Close
    For i = 0 To cmbBranch.ListCount - 1
        cmbBranch.ListIndex = i
        If CurrentBranch = cmbBranch.ItemData(cmbBranch.ListIndex) Then
            Exit For
        End If
    Next

End Sub

Private Sub Form_Unload(Cancel As Integer)
    VarActForm = ""
    If PosConnection.State = adStateOpen Then PosConnection.Close: Set PosConnection = Nothing
    
    If TempPerFlag = True Then
    End If

    SaveSetting strMainKey, Me.Name, "Left", Me.left
    SaveSetting strMainKey, Me.Name, "Top", Me.top

End Sub

Private Sub ResizeKit1_ExitResize(ByVal XScale As Double, ByVal YScale As Double)


    If formloadFlag = True Then
        SaveSetting strMainKey, Me.Name, "Height", Me.Height
        SaveSetting strMainKey, Me.Name, "Width", Me.Width
    End If

End Sub

Private Sub txtConfirm_GotFocus()
    SetKbLayout LANG_EN_US
    If TempPerFlag = True Then
        mdifrm.Toolbar1.Buttons(23).Value = tbrPressed
    End If
End Sub

Private Sub txtConfirm_LostFocus()
       If clsStation.Language = Farsi Then
           SetKbLayout LANG_Pr_IR
       End If

End Sub

Private Sub txtPassword_GotFocus()
    SetKbLayout LANG_EN_US
    If TempPerFlag = True Then
        mdifrm.Toolbar1.Buttons(23).Value = tbrPressed
    End If
End Sub

Private Sub txtPassword_LostFocus()
       If clsStation.Language = Farsi Then
           SetKbLayout LANG_Pr_IR
       End If

End Sub


Private Sub txtUserName_GotFocus()
    SetKbLayout LANG_EN_US
    If TempPerFlag = True Then
        mdifrm.Toolbar1.Buttons(23).Value = tbrPressed
    End If
End Sub

Private Sub txtUserName_LostFocus()
       If clsStation.Language = Farsi Then
           SetKbLayout LANG_Pr_IR
       End If
End Sub

Private Sub Insert_Tafsili(intPpno As Long, ShowMessageflag As Boolean)
    
On Error GoTo ErrHandler
    Dim rs As New ADODB.Recordset
    Dim TafsiliName As String
    TafsiliName = Trim(txtFirstName.Text) & " " & Trim(txtSurName.Text)
    If txtTafsiliCode.Text = "" Then
        txtTafsiliCode.Text = Accounting.Insert_TafsiliDll(ShowMessageflag, cmbBranch.ItemData(cmbBranch.ListIndex), TafsiliName, EnumAtf.Perssonel)
    Else
        Accounting.Update_TafsiliDll ShowMessageflag, cmbBranch.ItemData(cmbBranch.ListIndex), Val(txtTafsiliCode.Text), TafsiliName, EnumAtf.Perssonel
    End If
   
    If Val(txtTafsiliCode.Text) > 0 Then
        ReDim Parameter(1) As Parameter
        Parameter(0) = GenerateInputParameter("@TafsiliId", adInteger, 4, Val(txtTafsiliCode.Text))
        Parameter(1) = GenerateInputParameter("@Ppno", adInteger, 4, intPpno)
        RunParametricStoredProcedure "Update_tper_tafsili", Parameter
    End If
    If ShowMessageflag = False Then Exit Sub
Exit Sub
ErrHandler:
    MsgBox err.Description & "External Accountig"
    Resume Next


End Sub

Private Sub FillAtf()
    txtAtf.Text = "Å—”‰· Ê ”Â«„œ«—«‰"
End Sub

