VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{7AEDC602-D94C-11D1-BB7A-00E0290EA3C9}#1.0#0"; "ResizeKit.ocx"
Object = "{75D4F767-8785-11D3-93AD-0000832EF44D}#2.14#0"; "FAST2001.ocx"
Begin VB.Form frmFindCust 
   BackColor       =   &H00C0FFF0&
   Caption         =   "                                                                            Ã” ÃÊÌ „‘ —òÌ‰"
   ClientHeight    =   9150
   ClientLeft      =   2775
   ClientTop       =   3870
   ClientWidth     =   13515
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmFindCust.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   9150
   ScaleWidth      =   13515
   Begin VB.Frame FrameAddCust 
      BackColor       =   &H00C0E0FF&
      Caption         =   "             «ÿ·«⁄«  „‘ —ò ÃœÌœ                                      "
      BeginProperty Font 
         Name            =   "Nazanin"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6375
      Left            =   5190
      MouseIcon       =   "frmFindCust.frx":A4C2
      RightToLeft     =   -1  'True
      TabIndex        =   68
      Top             =   2175
      Visible         =   0   'False
      Width           =   4335
      Begin VB.CommandButton cmdAddCust 
         Caption         =   "À»  „‘ —ﬂ ÃœÌœ"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   2520
         Picture         =   "frmFindCust.frx":A7CC
         Style           =   1  'Graphical
         TabIndex        =   79
         Top             =   5280
         Width           =   1635
      End
      Begin VB.PictureBox frameOwner 
         BackColor       =   &H00C0E0FF&
         Height          =   495
         Left            =   120
         RightToLeft     =   -1  'True
         ScaleHeight     =   435
         ScaleWidth      =   3945
         TabIndex        =   85
         Top             =   480
         Width           =   4005
         Begin VB.CheckBox Check1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0E0FF&
            Caption         =   "„‰“·"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1440
            RightToLeft     =   -1  'True
            TabIndex        =   87
            Top             =   0
            Value           =   1  'Checked
            Width           =   2295
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00C0E0FF&
         Height          =   3135
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   71
         Top             =   960
         Width           =   4095
         Begin VB.TextBox txtAddFamily 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFF0&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   74
            Top             =   1100
            Width           =   2595
         End
         Begin VB.TextBox txtAddMobile 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFF0&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   480
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   76
            Top             =   2000
            Width           =   2595
         End
         Begin VB.TextBox txtAddName 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFF0&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   73
            Top             =   675
            Width           =   2595
         End
         Begin VB.TextBox txtAddMembershipId 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFF0&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   72
            Top             =   240
            Width           =   2595
         End
         Begin VB.TextBox txtAddAdress 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFF0&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   585
            Left            =   120
            MultiLine       =   -1  'True
            RightToLeft     =   -1  'True
            TabIndex        =   77
            Top             =   2450
            Width           =   2595
         End
         Begin VB.TextBox txtAddTel1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFF0&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   480
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   75
            Top             =   1530
            Width           =   2595
         End
         Begin VB.Label Label18 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "‰«„ Œ«‰Ê«œêÌ*"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   435
            Left            =   2760
            RightToLeft     =   -1  'True
            TabIndex        =   86
            Top             =   1080
            Width           =   1185
         End
         Begin VB.Label Label17 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "„Ê»«Ì·"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   2760
            RightToLeft     =   -1  'True
            TabIndex        =   84
            Top             =   2040
            Width           =   1185
         End
         Begin VB.Label Label16 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "‰«„"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   2760
            RightToLeft     =   -1  'True
            TabIndex        =   83
            Top             =   615
            Width           =   1185
         End
         Begin VB.Label Label15 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "òœ «‘ —«ò"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Left            =   2760
            RightToLeft     =   -1  'True
            TabIndex        =   82
            Top             =   255
            Width           =   1185
         End
         Begin VB.Label Label14 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "¬œ—” *"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   555
            Left            =   2760
            RightToLeft     =   -1  'True
            TabIndex        =   81
            Top             =   2505
            Width           =   1155
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "  ·›‰*"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   435
            Left            =   2760
            RightToLeft     =   -1  'True
            TabIndex        =   80
            Top             =   1575
            Width           =   1185
         End
      End
      Begin VB.Frame Frame10 
         BackColor       =   &H00C0E0FF&
         Caption         =   " Ê÷ÌÕ«  :"
         BeginProperty Font 
            Name            =   "Nazanin"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   70
         Top             =   4080
         Width           =   4095
         Begin VB.TextBox txtAddDescription 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFF0&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   585
            Left            =   120
            MultiLine       =   -1  'True
            RightToLeft     =   -1  'True
            TabIndex        =   78
            Top             =   360
            Width           =   3915
         End
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "»” ‰"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   960
         Left            =   120
         Picture         =   "frmFindCust.frx":E35E
         Style           =   1  'Graphical
         TabIndex        =   69
         Top             =   5280
         Width           =   1635
      End
   End
   Begin VB.CommandButton OKButton 
      BackColor       =   &H0000C000&
      Caption         =   "ﬁ»Ê·"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   10080
      RightToLeft     =   -1  'True
      TabIndex        =   57
      Top             =   8400
      Width           =   1215
   End
   Begin VB.CommandButton CancelButton 
      BackColor       =   &H000000C0&
      Cancel          =   -1  'True
      Caption         =   "«‰’—«›"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   56
      Top             =   8400
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   2100
      Top             =   240
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00C0FFF0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6480
      RightToLeft     =   -1  'True
      TabIndex        =   51
      Top             =   2520
      Width           =   2895
      Begin VB.TextBox TxtTimer 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFF0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   960
         RightToLeft     =   -1  'True
         TabIndex        =   52
         Text            =   "500"
         Top             =   120
         Width           =   615
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "„Ì·Ì À«‰ÌÂ"
         Height          =   495
         Left            =   0
         RightToLeft     =   -1  'True
         TabIndex        =   54
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "“„«‰  «ŒÌ— »Ì‰ ﬂ·ÌœÂ«Ì Ê—ÊœÌ"
         Height          =   615
         Left            =   1560
         RightToLeft     =   -1  'True
         TabIndex        =   53
         Top             =   120
         Width           =   1215
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00C0FFF0&
      Height          =   1800
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   47
      Top             =   3360
      Width           =   1725
      Begin VB.Label LblMaxCustomerNo 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   50
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label LblMaxCustomerActive 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   375
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   49
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label LblMaxCustomerInActive 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   375
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   48
         Top             =   1200
         Width           =   1455
      End
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00C0FFF0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   9480
      RightToLeft     =   -1  'True
      TabIndex        =   36
      Top             =   0
      Width           =   3975
      Begin VB.TextBox txtPhone 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFF0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   0
         RightToLeft     =   -1  'True
         TabIndex        =   41
         Top             =   970
         Width           =   2595
      End
      Begin VB.TextBox txtAddress 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFF0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Left            =   0
         MultiLine       =   -1  'True
         RightToLeft     =   -1  'True
         TabIndex        =   40
         Top             =   1420
         Width           =   2595
      End
      Begin VB.TextBox txtMembershipId 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFF0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   0
         RightToLeft     =   -1  'True
         TabIndex        =   39
         Top             =   120
         Width           =   2595
      End
      Begin VB.TextBox txtName 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFF0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   0
         RightToLeft     =   -1  'True
         TabIndex        =   38
         Top             =   555
         Width           =   2595
      End
      Begin VB.ComboBox CmbPrefix 
         BackColor       =   &H00C0FFF0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   0
         RightToLeft     =   -1  'True
         TabIndex        =   37
         Text            =   "Combo1"
         Top             =   2000
         Width           =   2595
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   " ·›‰"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   2640
         RightToLeft     =   -1  'True
         TabIndex        =   46
         Top             =   975
         Width           =   1065
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "¬œ—”"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Left            =   2640
         RightToLeft     =   -1  'True
         TabIndex        =   45
         Top             =   1425
         Width           =   1035
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "òœ «‘ —«ò"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   2640
         RightToLeft     =   -1  'True
         TabIndex        =   44
         Top             =   135
         Width           =   1065
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "‰«„"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   2640
         RightToLeft     =   -1  'True
         TabIndex        =   43
         Top             =   495
         Width           =   1065
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "‘€·"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   2640
         RightToLeft     =   -1  'True
         TabIndex        =   42
         Top             =   1920
         Width           =   1065
      End
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00C0FFF0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   15
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   0
      Width           =   6375
      Begin VB.CommandButton cmdTurnOver 
         Caption         =   "ê—œ‘ Õ”«» «Ì‰ „‘ —Ì"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2400
         TabIndex        =   90
         Top             =   2640
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.CommandButton CmdNewCust 
         BackColor       =   &H00C0FFE0&
         Caption         =   "„‘ —ﬂ ÃœÌœ"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   800
         Left            =   120
         Picture         =   "frmFindCust.frx":15A28
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   12
         ToolTipText     =   "«„ﬂ«‰ À»  «‘ —«ﬂ »Â ’Ê—  ”—Ì⁄ "
         Top             =   2450
         Width           =   1695
      End
      Begin VB.Label LblRecieve1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Height          =   375
         Left            =   255
         RightToLeft     =   -1  'True
         TabIndex        =   93
         Top             =   255
         Width           =   855
      End
      Begin VB.Label LblTotalCreditDebitLabel 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "»œÂÌ- ÿ·» ﬂ·:"
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   4920
         RightToLeft     =   -1  'True
         TabIndex        =   92
         Top             =   2040
         Width           =   1335
      End
      Begin VB.Label LblTotalCreditDebit 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H000000FF&
         Height          =   450
         Left            =   2400
         RightToLeft     =   -1  'True
         TabIndex        =   91
         Top             =   2040
         Width           =   2415
      End
      Begin VB.Label AddedDate1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Height          =   375
         Left            =   4200
         RightToLeft     =   -1  'True
         TabIndex        =   89
         Top             =   240
         Width           =   975
      End
      Begin VB.Label AddedDate 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   " «—ÌŒ ⁄÷ÊÌ "
         Height          =   375
         Left            =   5280
         RightToLeft     =   -1  'True
         TabIndex        =   88
         Top             =   240
         Width           =   975
      End
      Begin VB.Label MaxPrice 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "»Ì‘ —Ì‰ Œ—Ìœ"
         Height          =   375
         Left            =   5160
         RightToLeft     =   -1  'True
         TabIndex        =   35
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label LastDate 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "¬Œ—Ì‰ Œ—Ìœ"
         Height          =   375
         Left            =   3240
         RightToLeft     =   -1  'True
         TabIndex        =   34
         Top             =   240
         Width           =   975
      End
      Begin VB.Label BuyAverage 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "„Ì«‰êÌ‰ Œ—Ìœ"
         Height          =   375
         Left            =   3240
         RightToLeft     =   -1  'True
         TabIndex        =   33
         Top             =   720
         Width           =   975
      End
      Begin VB.Label LastNo 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "¬Œ—Ì‰ ›Ì‘"
         Height          =   375
         Left            =   5340
         RightToLeft     =   -1  'True
         TabIndex        =   32
         Top             =   1590
         Width           =   975
      End
      Begin VB.Label BuyCount 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   " ⁄œ«œ œ›⁄«  Œ—Ìœ"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   780
         RightToLeft     =   -1  'True
         TabIndex        =   31
         Top             =   1650
         Width           =   1215
      End
      Begin VB.Label MinPrice 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "ﬂ„ —Ì‰ Œ—Ìœ"
         Height          =   375
         Left            =   5280
         RightToLeft     =   -1  'True
         TabIndex        =   30
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label LastPrice 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "„»·€ ¬Œ—Ì‰ Œ—Ìœ"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3120
         RightToLeft     =   -1  'True
         TabIndex        =   29
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label MaxPrice1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Height          =   375
         Left            =   4260
         RightToLeft     =   -1  'True
         TabIndex        =   28
         Top             =   720
         Width           =   975
      End
      Begin VB.Label LastDate1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Height          =   375
         Left            =   2280
         RightToLeft     =   -1  'True
         TabIndex        =   27
         Top             =   240
         Width           =   975
      End
      Begin VB.Label BuyAverage1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Height          =   375
         Left            =   2280
         RightToLeft     =   -1  'True
         TabIndex        =   26
         Top             =   660
         Width           =   975
      End
      Begin VB.Label LastNo1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Height          =   375
         Left            =   4350
         RightToLeft     =   -1  'True
         TabIndex        =   25
         Top             =   1545
         Width           =   975
      End
      Begin VB.Label BuyCount1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Height          =   375
         Left            =   -405
         RightToLeft     =   -1  'True
         TabIndex        =   24
         Top             =   1605
         Width           =   975
      End
      Begin VB.Label MinPrice1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Height          =   375
         Left            =   4200
         RightToLeft     =   -1  'True
         TabIndex        =   23
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label LastPrice1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Height          =   375
         Left            =   2160
         RightToLeft     =   -1  'True
         TabIndex        =   22
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label LastCredit1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Height          =   375
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   21
         Top             =   660
         Width           =   1020
      End
      Begin VB.Label LastCredit 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "„«‰œÂ "
         Height          =   375
         Left            =   1530
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Top             =   705
         Width           =   375
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "«⁄ »«—Ì"
         Height          =   495
         Left            =   1320
         RightToLeft     =   -1  'True
         TabIndex        =   19
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "œ—Ì«› "
         Height          =   375
         Left            =   1470
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   330
         Width           =   495
      End
      Begin VB.Label LblBuy1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Height          =   375
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   17
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "‰—Œ"
         Height          =   495
         Left            =   1320
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Top             =   2040
         Width           =   375
      End
      Begin VB.Label LbCustomerlRate 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Height          =   375
         Left            =   720
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   2040
         Width           =   615
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Œ—Ìœ Â«Ì «„—Ê“"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3120
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   1680
         Width           =   1095
      End
      Begin VB.Label lblCountCurrentBuy 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Height          =   375
         Left            =   2280
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   1680
         Width           =   615
      End
   End
   Begin VB.TextBox TxtBarcode 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFF0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3720
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Text            =   "»«—ﬂœ"
      Top             =   8520
      Width           =   2055
   End
   Begin VB.PictureBox Frame1 
      BackColor       =   &H00C0FFF0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   7920
      RightToLeft     =   -1  'True
      ScaleHeight     =   2355
      ScaleWidth      =   1395
      TabIndex        =   7
      Top             =   120
      Width           =   1455
      Begin VB.OptionButton optInPerson 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFF0&
         Caption         =   "Õ÷Ê—Ì"
         Height          =   420
         Left            =   240
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   1800
         Width           =   1035
      End
      Begin VB.OptionButton optByPhone 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFF0&
         Caption         =   " ·›‰Ì"
         Height          =   435
         Left            =   240
         Picture         =   "frmFindCust.frx":1B97A
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   600
         Value           =   -1  'True
         Width           =   1035
      End
      Begin VB.Shape Shape1 
         BorderStyle     =   2  'Dash
         Height          =   1095
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   30
         Width           =   1215
      End
      Begin VB.Image Image2 
         Height          =   480
         Left            =   360
         Picture         =   "frmFindCust.frx":1C244
         Top             =   1320
         Width           =   480
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   360
         Picture         =   "frmFindCust.frx":1C54E
         Top             =   120
         Width           =   480
      End
      Begin VB.Shape Shape2 
         BorderStyle     =   2  'Dash
         Height          =   1095
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   1200
         Width           =   1215
      End
   End
   Begin VB.PictureBox Frame2 
      BackColor       =   &H00C0FFF0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   6480
      RightToLeft     =   -1  'True
      ScaleHeight     =   2355
      ScaleWidth      =   1275
      TabIndex        =   3
      Top             =   120
      Width           =   1335
      Begin VB.OptionButton OptTable 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFF0&
         Caption         =   "„Ì“  "
         Height          =   435
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   1680
         Width           =   1035
      End
      Begin VB.OptionButton OptDelivery 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFF0&
         Caption         =   "«—”«·Ì"
         Height          =   435
         Left            =   240
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   120
         Value           =   -1  'True
         Width           =   915
      End
      Begin VB.OptionButton OptSaloon 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFF0&
         Caption         =   "”«·‰"
         Height          =   435
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   960
         Width           =   1035
      End
   End
   Begin VB.PictureBox Frame3 
      BackColor       =   &H00C0FFF0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   9480
      RightToLeft     =   -1  'True
      ScaleHeight     =   675
      ScaleWidth      =   3795
      TabIndex        =   0
      Top             =   2640
      Width           =   3855
      Begin VB.OptionButton Option1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFF0&
         Caption         =   "Ã” ÃÊÌ ”—Ì⁄"
         Height          =   375
         Index           =   0
         Left            =   2040
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   120
         Width           =   1695
      End
      Begin VB.OptionButton Option1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFF0&
         Caption         =   "Ã” ÃÊÌ „⁄„Ê·Ì"
         Height          =   375
         Index           =   1
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   120
         Width           =   1815
      End
   End
   Begin VSFlex7LCtl.VSFlexGrid vsCustomer 
      Height          =   4905
      Left            =   1695
      TabIndex        =   55
      Top             =   3465
      Width           =   11835
      _cx             =   20876
      _cy             =   8652
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483644
      ForeColor       =   12582912
      BackColorFixed  =   16761024
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   16744576
      BackColorAlternate=   -2147483644
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
      SelectionMode   =   3
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   8
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   500
      RowHeightMax    =   500
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmFindCust.frx":1CE18
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
      ExplorerBar     =   5
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
   Begin FLWCtrls.FWProgressBar FWProgressBar1 
      Height          =   375
      Left            =   0
      Top             =   7920
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      BackColor       =   -2147483626
      BorderStyle     =   10
   End
   Begin RESIZEKITLibCtl.ResizeKit ResizeKit1 
      Height          =   480
      Left            =   0
      OleObjectBlob   =   "frmFindCust.frx":1CEF7
      TabIndex        =   58
      Top             =   0
      Width           =   480
   End
   Begin VB.PictureBox Frame_7 
      BackColor       =   &H00C0FFF0&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   0
      RightToLeft     =   -1  'True
      ScaleHeight     =   2715
      ScaleWidth      =   1635
      TabIndex        =   62
      Top             =   5130
      Width           =   1695
      Begin VB.CommandButton CmdFamilySet 
         BackColor       =   &H000000C0&
         Caption         =   "Set"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   30
         RightToLeft     =   -1  'True
         TabIndex        =   64
         Top             =   600
         Width           =   1575
      End
      Begin VB.TextBox TxtFamilyNo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   63
         Top             =   15
         Width           =   615
      End
      Begin VB.Label LblLastBuy 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0E0FF&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   240
         RightToLeft     =   -1  'True
         TabIndex        =   67
         Top             =   1680
         Width           =   1455
      End
      Begin VB.Label LblResult 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0E0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "  €ÌÌ— «‰Ã«„         ‘œ"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   495
         Left            =   165
         RightToLeft     =   -1  'True
         TabIndex        =   66
         Top             =   1110
         Width           =   1095
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   " ò›·"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   495
         Left            =   600
         RightToLeft     =   -1  'True
         TabIndex        =   65
         Top             =   0
         Width           =   615
      End
   End
   Begin VB.Image Image5 
      Height          =   1455
      Left            =   15
      Picture         =   "frmFindCust.frx":1CF7D
      Stretch         =   -1  'True
      Top             =   5145
      Width           =   1695
   End
   Begin VB.Image Image4 
      Height          =   1455
      Left            =   15
      Stretch         =   -1  'True
      Top             =   5175
      Width           =   1695
   End
   Begin VB.Image Image3 
      Height          =   1470
      Left            =   -75
      Picture         =   "frmFindCust.frx":1FFE2
      Stretch         =   -1  'True
      Top             =   6525
      Width           =   3120
   End
   Begin VB.Label LblFindCust 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
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
      Height          =   495
      Left            =   1320
      RightToLeft     =   -1  'True
      TabIndex        =   61
      Top             =   120
      Width           =   3255
   End
   Begin VB.Label LblCount 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
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
      TabIndex        =   60
      Top             =   8520
      Width           =   3015
   End
   Begin VB.Label lablkharid4 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1200
      RightToLeft     =   -1  'True
      TabIndex        =   59
      Top             =   1920
      Width           =   975
   End
End
Attribute VB_Name = "frmFindCust"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim CountCustomer As Long
Dim i, j As Long
Dim SearchType As Integer
Dim Rst As New ADODB.Recordset
Dim tmpflag As Boolean
Dim mvarbarcode As Boolean
Dim clsDate As New clsDate
Dim CountDailyBuy As Integer
Dim CountShiftBuy As Integer
Dim TotalCreditDebit As Currency

Private Sub ShowCustPicture()
    
    On Error GoTo ErrHandler
    Dim Rst As New ADODB.Recordset
    Dim rctmp As New ADODB.Recordset
    ReDim Parameter(0) As Parameter
   
    Parameter(0) = GenerateInputParameter("@code", adInteger, 4, vsCustomer.TextMatrix(vsCustomer.Row, 1))
    Set Rst = RunParametricStoredProcedure2Rec("Get_All_TCust_Picture", Parameter)
    If Not (Rst.EOF = True And Rst.BOF = True) Then
        ReDim Parameter(0) As Parameter
        Parameter(0) = GenerateInputParameter("@intCode", adInteger, 4, Rst!PictureNo)
        Set rctmp = RunParametricStoredProcedure2Rec("Get_TCust_Picture", Parameter)
        If Not (rctmp.BOF Or rctmp.EOF) Then
            Image4.Picture = LoadPicture(rctmp!PicturePath)
        End If
        Image5.Visible = False
        Image4.Visible = True
    Else
        Image5.Visible = True
        Image4.Visible = False
        Image4.Picture = LoadPicture("")
    End If
    Set Rst = Nothing
    Set rctmp = Nothing

Exit Sub
ErrHandler:
 If err.Number = 53 Then
    
    ShowDisMessage "⁄ﬂ” „‘ —ò „Ê—œ ‰Ÿ— Å«ﬂ ‘œÂ «” ", 1000
 End If
    Image4.Picture = LoadPicture("")
    Image5.Visible = True
    Image4.Visible = False
    Set Rst = Nothing
    Set rctmp = Nothing

End Sub

Private Sub CancelButton_Click()
    mvarcode = 0
    txtMembershipId.Text = ""
    txtPhone.Text = ""
    txtName.Text = ""
    TxtAddress.Text = ""
    CreditCode = 0
    cmbPrefix.ListIndex = 0
    ''Me.Hide
    Unload Me
End Sub

Private Sub CmbPrefix_Change()
''''    If Option1(0).Value = False Then
''''        i = vsCustomer.FindRow(CmbPrefix.ItemData(CmbPrefix.ListIndex), 1, 6, False, False)
''''        If i > 0 Then
''''            vsCustomer.Row = i
''''            vsCustomer.ShowCell i, 0
''''            LblFindCust.Caption = ""
''''        Else
''''            vsCustomer.Row = 0
''''            vsCustomer.ShowCell 0, 0
''''            labelClear
''''            If Trim(CmbPrefix.Text) <> "" Then
''''               LblFindCust.Caption = " ‰«„ ( " & CmbPrefix.Text & " ) œ— ”Ì” „  ⁄—Ì› ‰‘œÂ"
''''             Else
''''                LblFindCust.Caption = "‰«„ „‘ —ﬂ —« Ê«—œ ﬂ‰Ìœ  "
''''             End If
''''
''''        End If
''''    Else
''''       If Trim(CmbPrefix.Text) <> "" Then
''''            SearchType = 5
''''            Timer1.Interval = Val(TxtTimer.Text)
''''            Timer1.Enabled = True
''''       Else
''''          vsCustomer.Rows = 1
''''          labelClear
''''
''''       End If
''''
''''   End If

End Sub

Private Sub CmbPrefix_Click()
'    txtName.Text = ""
'    TxtAddress.Text = ""
'    txtMembershipId.Text = ""
'    txtPhone.Text = ""
    If Option1(0).Value = False Then
        i = vsCustomer.FindRow(cmbPrefix.ItemData(cmbPrefix.ListIndex), 1, 7, False, False)
        If i > 0 Then
            vsCustomer.Row = i
            vsCustomer.ShowCell i, 0
            LblFindCust.Caption = ""
        Else
            vsCustomer.Row = 0
            vsCustomer.ShowCell 0, 0
            labelClear
            If Trim(cmbPrefix.Text) <> "" Then
               LblFindCust.Caption = " ‰«„ ( " & cmbPrefix.Text & " ) œ— ”Ì” „  ⁄—Ì› ‰‘œÂ"
             Else
                LblFindCust.Caption = "‰«„ „‘ —ﬂ —« Ê«—œ ﬂ‰Ìœ  "
             End If
    
        End If
    Else
       If cmbPrefix.ListIndex > 0 Then
            SearchType = 5
            Timer1.Interval = Val(TxtTimer.Text)
            Timer1.Enabled = True
       Else
          vsCustomer.Rows = 1
          labelClear

       End If
  
   End If

End Sub

Private Sub cmdAddCust_Click()
        
    On Error GoTo ErrHandler
    If txtAddFamily = "" Or txtAddTel1 = "" Or txtAddAdress = "" Then
        ShowDisMessage "«ÿ·«⁄«  Ê—ÊœÌ ﬂ«„· ‰Ì” ", 1500
        Exit Sub
    End If
    ReDim Parameter(8) As Parameter
    Parameter(0) = GenerateInputParameter("@MembershipId", adBigInt, 8, Val(txtAddMembershipId.Text))
    Parameter(1) = GenerateInputParameter("@Name", adVarWChar, 50, Trim(txtAddName.Text))
    Parameter(2) = GenerateInputParameter("@Family", adVarWChar, 50, Trim(txtAddFamily.Text))
    Parameter(3) = GenerateInputParameter("@Address", adVarWChar, 255, Trim(txtAddAdress.Text))
    Parameter(4) = GenerateInputParameter("@Tel1", adVarWChar, 50, Trim(txtAddTel1.Text))
    Parameter(5) = GenerateInputParameter("@Mobile", adVarWChar, 50, Trim(txtAddMobile.Text))
    Parameter(6) = GenerateInputParameter("@Description", adVarWChar, 255, Trim(txtAddDescription.Text))
    Parameter(7) = GenerateInputParameter("@User", adInteger, 4, mvarCurUserNo)
    Parameter(8) = GenerateOutputParameter("@Code", adBigInt, 8)
    
    Dim LastCode As Long
    LastCode = RunParametricStoredProcedure("Insert_CustomerFast", Parameter)
    If LastCode > 0 Then
        ShowMessage "À»  „‘ —ò ÃœÌœ »« „Ê›ﬁÌ  «‰Ã«„ ‘œ", True, False, " «ÌÌœ", ""
        vsCustomer.Rows = 1
        labelClear
'        Option1(0).Value = True
        If clsStation.DefaultCustSearch = EnumDefaultCustSearch.MembershipId Then
            txtMembershipId = txtAddMembershipId
        ElseIf clsStation.DefaultCustSearch = EnumDefaultCustSearch.Phone Then
            txtPhone = ""
            txtPhone = txtAddTel1
        ElseIf clsStation.DefaultCustSearch = EnumDefaultCustSearch.Name Then
            txtName = txtAddName
        ElseIf clsStation.DefaultCustSearch = EnumDefaultCustSearch.address Then
            txtAddAdress = txtAddAdress
        End If
        ClearFrameAddCust
        FrameAddCust.Visible = False
'        txtMembershipId.SetFocus
    Else
        ShowMessage "À»  «‰Ã«„ ‰‘œ" + vbCrLf + "«ÿ·«⁄«  —« »——”Ì ‰„«ÌÌœ.", True, False, " «ÌÌœ", ""
    End If
Exit Sub
ErrHandler:
    ShowDisMessage err.Description, 1500
End Sub

Private Sub cmdClose_Click()
    FrameAddCust.Visible = False
End Sub

Private Sub ClearFrameAddCust()
    txtAddAdress = ""
    txtAddDescription = ""
    txtAddFamily = ""
    txtAddMembershipId = ""
    txtAddMobile = ""
    txtAddName = ""
    txtAddTel1 = ""
End Sub
Private Sub CmdFamilySet_Click()
If Val(txtMembershipId.Text) = 0 Or Val(vsCustomer.Rows) > 2 Then Exit Sub
    ReDim Parameter(1) As Parameter
        Parameter(0) = GenerateInputParameter("@FamilyNo", adInteger, 4, Val(TxtFamilyNo.Text))
        Parameter(1) = GenerateInputParameter("@Code", adBigInt, 8, Val(txtMembershipId.Text))
        Set Rst = RunParametricStoredProcedure2Rec("Update_Customer_FamilyNo", Parameter)
        LblResult.Visible = True
        CmdFamilySet.Enabled = False
        vsCustomer.SetFocus
        tmpflag = True
End Sub

Private Sub CmdNewCust_Click()
    NewCallNumber = ""
    Dim rctmp As New ADODB.Recordset
    
    If ClsFormAccess.frmCust = True Then
        If Len(txtPhone.Text) > 0 Then
            ShowMessage "¬Ì« „Ì ŒÊ«ÂÌœ «Ì‰  ·›‰ œ— ﬁ”„  „‘ —Ì«‰ ÃœÌœ À»  ê—œœø", True, True, "»·Ì", "ŒÌ—"
            If mvarMsgIdx = vbYes Then
                NewCallNumber = txtPhone.Text
            Else
                NewCallNumber = ""
            End If
            If clsStation.FastCustSave = True Then
            
                FrameAddCust.Visible = True
                ClearFrameAddCust
'                ReDim Parameter(0) As Parameter
'                Parameter(0) = GenerateInputParameter("@Branch", adInteger, 4, CurrentBranch)
'                Set rctmp = RunParametricStoredProcedure2Rec("Get_New_Cust_Code", Parameter)
'                txtAddMembershipId.Text = rctmp.Fields("MembershipId").Value
                txtAddTel1 = txtPhone
'                rctmp.Close
'                Set rctmp = Nothing
            Else
                FindCustFlag = True
                txtMembershipId.Text = ""
                txtPhone.Text = ""
                txtName.Text = ""
                TxtAddress.Text = ""
                CreditCode = 0
                cmbPrefix.ListIndex = 0
                ''Me.Hide
                Unload Me
                Unload frmCallerIdView
                frmCust.Show
            End If
        Else
            If clsStation.FastCustSave = True Then
                FrameAddCust.Visible = True
                ClearFrameAddCust
                ReDim Parameter(0) As Parameter
                Parameter(0) = GenerateInputParameter("@Branch", adInteger, 4, CurrentBranch)
                Set rctmp = RunParametricStoredProcedure2Rec("Get_New_Cust_Code", Parameter)
                txtAddMembershipId.Text = rctmp.Fields("MembershipId").Value
                rctmp.Close
                Set rctmp = Nothing
            Else
                FindCustFlag = True
                txtMembershipId.Text = ""
                txtPhone.Text = ""
                txtName.Text = ""
                TxtAddress.Text = ""
                CreditCode = 0
                cmbPrefix.ListIndex = 0
                ''Me.Hide
                Unload Me
                Unload frmCallerIdView
                frmCust.Show
            End If
        End If
    Else
        ShowDisMessage "œ” —”Ì »—«Ì À»  „‘ —ﬂ ÃœÌœ ÊÃÊœ ‰œ«—œ", 1500
    End If

End Sub

Private Sub cmdTurnOver_Click()
    If vsCustomer.Row < 1 Then Exit Sub
    
    If ClsFormAccess.AccfrmKartHesabReport = True Then
        If vsCustomer.ValueMatrix(vsCustomer.Row, 8) > 0 Then
            Accounting.KartHesabShowDll "KolBedehkaran", CStr(vsCustomer.TextMatrix(vsCustomer.Row, 8)), CStr(vsCustomer.TextMatrix(vsCustomer.Row, 3)), Mid(AccountYear & "/01/01", 3), Mid(clsDate.shamsi(Date), 3)
        Else
            ShowDisMessage "«Ì‰ „‘ —Ì œ— ”Ì” „ Õ”«»œ«—Ì œ«—«Ì ﬂœ  ›÷Ì·Ì ‰Ì” ", 2000
        End If
    Else
        ShowDisMessage "‘„« »Â «Ì‰ «„ﬂ«‰ œ” —”Ì ‰œ«—Ìœ", 1500
    End If

End Sub

Private Sub Form_Activate()
    
    On Error Resume Next
    Dim tmpSearch As Boolean
    
    Dim hMenu As Long

    hMenu = GetSystemMenu(Me.hWnd, False)

    DeleteMenu hMenu, 6, MF_BYPOSITION
    
    CmdNewCust.Visible = False
    CancelButton.Visible = False
    OKButton.Visible = False
    Frame3.Visible = False
    
    If Val(TxtTimer.Text) < 1 Then TxtTimer.Text = "100"
    TxtTimer.Text = clsStation.SrarchInputDelayKeyboard
    
    If clsStation.CustomerSearchDefault = True Then
        Option1(0).Value = True
    Else
        Option1(1).Value = True
    End If
'    Option1(0).Value = clsStation.CustomerSearchDefault
'    Option1(1).Value = Not (clsStation.CustomerSearchDefault)
'    Option1(1).Value = True
    
    Select Case clsStation.DefaultCustSearch
        Case EnumDefaultCustSearch.MembershipId
            txtMembershipId.SetFocus
            LblFindCust.Caption = "ﬂœ «‘ —«ﬂ —« Ê«—œ ﬂ‰Ìœ  "
        Case EnumDefaultCustSearch.address
            TxtAddress.SetFocus
            LblFindCust.Caption = "¬œ—” „‘ —ﬂ —« Ê«—œ ﬂ‰Ìœ  "
        Case EnumDefaultCustSearch.Name
            txtName.SetFocus
            LblFindCust.Caption = "‰«„ „‘ —ﬂ —« Ê«—œ ﬂ‰Ìœ  "
        Case EnumDefaultCustSearch.Phone
            txtPhone.SetFocus
            LblFindCust.Caption = " ·›‰ „‘ —ﬂ —« Ê«—œ ﬂ‰Ìœ  "
    End Select
    If Not (Val(strCategory) >= 0 And Val(strCategory) <= 7) Then
       OptTable.Visible = False
       OptSaloon.Caption = "›—Ê‘ê«Â"
    End If
    
    If clsStation.CustomerOrderDefault = True Then
       optInPerson = True
    Else
       optByPhone = True
    End If
    If clsStation.CustomerServeplace = 0 Then
       OptDelivery = True
    ElseIf clsStation.CustomerServeplace = 1 Then
       OptSaloon = True
    ElseIf clsStation.CustomerServeplace = 2 Then
       OptTable = True
    End If
    
''''    tmpSearch = clsStation.CustomerSearchDefault    ' For Speed Search
''''    clsStation.CustomerSearchDefault = True
    If CreditCode > 0 Then
        txtMembershipId.SetFocus
        txtMembershipId.Text = CreditCode
''''        Sleep 500
        If vsCustomer.Row > 0 Then
            mvarcode = vsCustomer.TextMatrix(vsCustomer.Row, 1)
            mvarName = vsCustomer.TextMatrix(vsCustomer.Row, 3)
            mvarPublicOrderType = inPerson
            mvarServePlace = EnumServePlace.Salon
            CreditCode = 0
'            clsStation.CustomerSearchDefault = tmpSearch
            txtMembershipId.Text = ""
            txtPhone.Text = ""
            txtName.Text = ""
            TxtAddress.Text = ""
            cmbPrefix.ListIndex = 0
            ''Me.Hide
            Unload Me
        End If
    Else        'Not Credit Card
'        clsStation.CustomerSearchDefault = tmpSearch
''''        If clsStation.CustomerSearchDefault = False Then
''''            FillvsCustomer
''''        End If
    End If
    If Len(Call_RealNumber) > 0 Then
        txtPhone.SetFocus
        txtPhone.Text = Call_RealNumber
    ElseIf Len(NewCallNumber) > 0 Then
        txtPhone.SetFocus
        txtPhone.Text = NewCallNumber
        NewCallNumber = ""
    Else        'Not Caller Id
'        clsStation.CustomerSearchDefault = tmpSearch
'        If clsStation.CustomerSearchDefault = False Then
'            FillvsCustomer
'        End If
    End If
    CreditCode = 0

    ReDim Parameter(0) As Parameter
    Parameter(0) = GenerateInputParameter("@Branch", adInteger, 4, CurrentBranch)
    Set Rst = RunParametricStoredProcedure2Rec("Get_All_Customers_Count", Parameter)
    LblMaxCustomerNo.Caption = "ﬂ· : " & Rst!MaxCustomerNo
    LblMaxCustomerActive.Caption = "›⁄«·:" & Rst!MaxCustomerActive
    LblMaxCustomerInActive.Caption = "»«ÿ· :" & Rst!MaxCustomerInActive
    CountCustomer = Rst!MaxCustomerActive

    CancelButton.Visible = True
    OKButton.Visible = True
    Frame3.Visible = True
    If VarActForm = "frmCust" Then
      CmdNewCust.Visible = False
    Else
        CmdNewCust.Visible = True
    End If
    
    mvarbarcode = False
    If clsArya.LimitedVersion = True And HardLockFlagTrial = False Then
        If CountCustomer > 100 Then
            MsgBox " ‰”ŒÂ ¬“„«Ì‘Ì - ‘„« »Ì‘ «“ «Ì‰ „Ã«“ »Â «” ›«œÂ «“ «Ì‰ ”Ì” „ ‰Ì” Ìœ  ", vbCritical
            On Error Resume Next
            Dim Obj As Object
            For Each Obj In Me
                If TypeOf Obj Is VSFlexGrid Or TypeOf Obj Is MaskEdBox Or TypeOf Obj Is Label Then
                    Obj.Enabled = False
                ElseIf TypeOf Obj Is TextBox Then     ' Or TypeOf obj Is ComboBox
                    Obj.Locked = True
                End If
            Next Obj
            On Error GoTo 0
        End If
    End If

End Sub

''''Public Sub barcode()
''''    txtMembershipId.Text = Mid(txtBarcode.Text, 9, 4)
''''    Timer1_Timer
''''    mvarcode = vsCustomer.TextMatrix(vsCustomer.Row, 1)
''''    ReDim Parameter(1) As Parameter
''''    Parameter(0) = GenerateInputParameter("@Code", adInteger, 4, Val(vsCustomer.TextMatrix(vsCustomer.Row, 1)))
''''    Parameter(1) = GenerateInputParameter("@Discount", adInteger, 4, Val(Mid(txtBarcode.Text, 7, 2)))
''''    Set Rst = RunParametricStoredProcedure2Rec("Update_Customer_Discount", Parameter)
''''
''''    If vsCustomer.Rows >= 2 Then
''''        OKButton_Click
''''    End If
''''End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 115 Then   'F4 Key
        CmdNewCust_Click
    End If
''''    If mvarbarcode = True Then
''''
''''
''''    Select Case KeyCode
''''
''''            Case 111, 191: '/ Barcode
''''             '   txtBarcode_Change
''''                Me.barcode
''''
''''        End Select
''''
''''    Else
''''
''''        Select Case KeyCode
''''
''''            Case 111, 191: '/ Barcode
''''
''''                mvarbarcode = True
''''                txtBarcode.Text = ""
''''                txtBarcode.SetFocus
''''
''''        End Select
''''
''''    End If

End Sub

Private Sub Form_Load()
    CenterCenterinSecondScreen Me
    
    mvarcode = 0
    If VarActForm = "frmCust" Then
      CmdNewCust.Visible = False
    End If
    TxtFamilyNo.Text = ""
    LblResult.Visible = False
    If strCategory <> "07" Then
        Label10.Visible = False
        TxtFamilyNo.Visible = False
        CmdFamilySet.Visible = False
        LblLastBuy.Visible = False
        Frame_7.Visible = False
    End If
    LblLastBuy.Caption = ""
    txtBarcode.Text = ""
    cmbPrefix.Clear
    Set Rst = RunStoredProcedure2RecordSet("Get_All_tPrefix")
    With cmbPrefix
        .AddItem ""
        .ItemData(.NewIndex) = 0
        If Not (Rst.EOF = True And Rst.BOF = True) Then
            While Rst.EOF <> True
                .AddItem Rst!Description
                .ItemData(.NewIndex) = Rst!Code
                Rst.MoveNext
            Wend
        End If
    End With
    With vsCustomer
        .Cols = 9
        .TextMatrix(0, 7) = "⁄‰Ê«‰"
        Set Rst = RunStoredProcedure2RecordSet("Get_All_tPrefix")
        .ColComboList(7) = vsCustomer.BuildComboList(Rst, "Description", "Code")
        .ColHidden(6) = True
         For i = 0 To .Cols - 1
            .ColWidth(i) = Val(GetSetting(strMainKey, "frmFindCust_vsCustomer", "Col" & i))
            If .ColWidth(i) = 0 Then
                .ColWidth(i) = 1000       '
            End If
         Next i
    End With
    
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


    ChangeLanguage
    
    If clsArya.ExternalAccounting = True Then cmdTurnOver.Visible = True Else cmdTurnOver.Visible = False

End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call PresetScreenSaver
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call PresetScreenSaver
End Sub

Public Sub ChangeLanguage()
If clsStation.Language = English Then
 AddedDate.Caption = "Registery Date"
 BuyAverage.Caption = "Buy Average"
 BuyCount.Caption = "Buy Count"
 CancelButton.Caption = "Cancel"
 CmdNewCust.Caption = "New Customer"
 Label1.Caption = "Name"
 Label11.Caption = "Receive"
 Label12.Caption = "Today buys"
 Label13.Caption = "Career"
 Label2.Caption = "Customer ID"
 Label3.Caption = "Address"
 Label4.Caption = "Phone NO"
 Label5.Caption = "Credit"
 Label6.Caption = "Rate"
 Label7.Caption = "Delay between input keys"
 Label8.Caption = "miliSecond"
 LastCredit.Caption = "Remained"
 LastDate.Caption = "Last buy"
 LastNo.Caption = "Last invoice"
 LastPrice.Caption = "Cost of last buy"
 MaxPrice.Caption = "Maximum buy"
 MinPrice.Caption = "Minimum buy"
 OKButton.Caption = "Accept"
 optByPhone.Caption = "By phone"
 OptDelivery.Caption = "Transmited"
 optInPerson.Caption = "Inside"
 Option1(0).Caption = "Fast Search"
 Option1(1).Caption = "Regular Search"
 OptSaloon.Caption = "Saloon"
 OptTable.Caption = "Table"
 With vsCustomer
    .TextMatrix(0, 0) = "Row"
    .TextMatrix(0, 2) = "Customer ID"
    .TextMatrix(0, 3) = "Name"
    .TextMatrix(0, 4) = "Phone"
    .TextMatrix(0, 5) = "Address"
    .TextMatrix(0, 7) = "Title"
    .TextMatrix(0, 8) = "Tafsili"
 End With
Else
 AddedDate.Caption = " «—Œ ⁄÷ÊÌ "
 BuyAverage.Caption = "„Ì«‰êÌ‰ Œ—Ìœ"
 BuyCount.Caption = " ⁄œ«œ œ›⁄«  Œ—Ìœ"
 CancelButton.Caption = "«‰’—«›"
 CmdNewCust.Caption = "„‘ —ﬂ ÃœÌœ"
 Label1.Caption = "‰«„"
 Label11.Caption = "œ—Ì«› "
 Label12.Caption = "Œ—ÌœÂ«Ì «„—Ê“"
 Label13.Caption = "‘€·"
 Label2.Caption = "«‘ —«ﬂ"
 Label3.Caption = "¬œ—”"
 Label4.Caption = " ·›‰"
 Label5.Caption = "«⁄ »«—Ì"
 Label6.Caption = "‰—Œ"
 Label7.Caption = "“„«‰  «ŒÌ— »Ì‰ ﬂ·ÌœÂ«Ì Ê—ÊœÌ"
 Label8.Caption = "„Ì·Ì À«‰ÌÂ"
 LastCredit.Caption = "„«‰œÂ"
 LastDate.Caption = "¬Œ—Ì‰ Œ—Ìœ"
 LastNo.Caption = "¬Œ—Ì‰ ›Ì‘"
 LastPrice.Caption = "„»·€ ¬Œ—Ì‰ Œ—Ìœ"
 MaxPrice.Caption = "»Ì‘ —Ì‰ Œ—Ìœ"
 MinPrice.Caption = "ﬂ„ —Ì‰ Œ—Ìœ"
 OKButton.Caption = "ﬁ»Ê·"
 optByPhone.Caption = " ·›‰Ì"
 OptDelivery.Caption = "«—”«·Ì"
 optInPerson.Caption = "Õ÷Ê—Ì"
 Option1(0).Caption = "Ã” ÃÊÌ ”—Ì⁄"
 Option1(1).Caption = "Ã” ÃÊÌ „⁄„Ê·Ì"
 OptSaloon.Caption = "”«·‰"
 OptTable.Caption = "„Ì“"
 With vsCustomer
    .TextMatrix(0, 0) = "—œÌ›"
    .TextMatrix(0, 2) = "«‘ —«ﬂ"
    .TextMatrix(0, 3) = "‰«„"
    .TextMatrix(0, 4) = " ·›‰"
    .TextMatrix(0, 5) = "¬œ—”"
    .TextMatrix(0, 7) = "⁄‰Ê«‰"
    .TextMatrix(0, 8) = " ›÷Ì·Ì"
 End With
End If
End Sub





Private Sub Option1_Click(index As Integer)
    If Option1(0).Value = False Then
        FillvsCustomer
    Else
        vsCustomer.Rows = 1
        labelClear
    End If
'    If txtMembershipId.Visible Then txtMembershipId.SetFocus
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    txtMembershipId.Text = ""
    txtPhone.Text = ""
    txtName.Text = ""
    TxtAddress.Text = ""
    CreditCode = 0
    cmbPrefix.ListIndex = 0
    
    If PosConnection.State = adStateOpen Then PosConnection.Close: Set PosConnection = Nothing
    Set Rst = Nothing
    SaveSetting strMainKey, Me.Name, "Left", Me.left
    SaveSetting strMainKey, Me.Name, "Top", Me.top
End Sub

Private Sub OKButton_Click()
 
    If vsCustomer.Row > 0 Then
               If (VarActForm <> "frmCust" And VarActForm <> "frmCreditCustomerAccount") Then
                    If Not (CountDailyBuy < clsStation.CountCustomerDailyBuy Or clsStation.CountCustomerDailyBuy = 0) Then 'For Restaurant
        
                        frmDisMsg.lblMessage.Caption = "  ⁄œ«œ œ›⁄«  Œ—Ìœ ‘„« œ— —Ê“ »Â Å«Ì«‰ —”ÌœÂ  «”  "
                        frmDisMsg.Timer1.Interval = 1000
                        frmDisMsg.Timer1.Enabled = True
                        frmDisMsg.Show vbModal
                        Exit Sub
                     ElseIf Not (CountShiftBuy < clsStation.CountCustomerShiftBuy Or clsStation.CountCustomerShiftBuy = 0) Then 'For Restaurant
        
                        frmDisMsg.lblMessage.Caption = "  ⁄œ«œ œ›⁄«  Œ—Ìœ ‘„« œ— ‘Ì›  »Â Å«Ì«‰ —”ÌœÂ  «”  "
                        frmDisMsg.Timer1.Interval = 1000
                        frmDisMsg.Timer1.Enabled = True
                        frmDisMsg.Show vbModal
                        Exit Sub
                    End If
                End If
                mvarcode = vsCustomer.TextMatrix(vsCustomer.Row, 1)
                mvarMemberShipId = vsCustomer.TextMatrix(vsCustomer.Row, 2)
                mvarName = vsCustomer.TextMatrix(vsCustomer.Row, 3)
                If optByPhone.Value = True Then
                    mvarPublicOrderType = ByPhone
                Else
                    mvarPublicOrderType = inPerson
                End If
                If OptDelivery.Value = True Then
                   mvarServePlace = EnumServePlace.Delivery
                ElseIf OptSaloon.Value = True Then
                   mvarServePlace = EnumServePlace.Salon
                ElseIf OptTable.Value = True Then
                   mvarServePlace = EnumServePlace.Table
                End If
    Else
        mvarcode = 0
        mvarName = ""
        mvarPublicOrderType = inPerson
    End If
''''    BeforeCustomerSellPrice = clsStation.PriceType
''''    clsStation.PriceType = CustomerSellPrice
    txtMembershipId.Text = ""
    txtPhone.Text = ""
    txtName.Text = ""
    TxtAddress.Text = ""
    CreditCode = 0
    cmbPrefix.ListIndex = 0
'''    CustomerSellPrice = 1
    ''Me.Hide
    Unload Me
End Sub


Private Sub ResizeKit1_ExitResize(ByVal XScale As Double, ByVal YScale As Double)

    If formloadFlag = True Then
        SaveSetting strMainKey, Me.Name, "Height", Me.Height
        SaveSetting strMainKey, Me.Name, "Width", Me.Width
    End If

End Sub

Private Sub Timer1_Timer()
    Timer1.Enabled = False
 '   DoEvents
    Define_Customer
    If vsCustomer.Rows > 1 Then
        vsCustomer.Row = 1
        vsCustomer.ShowCell 1, 0
        LblFindCust.Caption = ""
    Else
        vsCustomer.Row = 0
        vsCustomer.ShowCell 0, 0
        labelClear
        Select Case SearchType
            Case 1:
                 If Val(txtMembershipId.Text) > 0 Then
                   LblFindCust.Caption = " «‘ —«ﬂ ( " & txtMembershipId.Text & " ) œ— ”Ì” „  ⁄—Ì› ‰‘œÂ"
                 Else
                    LblFindCust.Caption = "ﬂœ «‘ —«ﬂ —« Ê«—œ ﬂ‰Ìœ  "
                 End If
            Case 2:
                 If Len(txtName.Text) > 0 Then
                   LblFindCust.Caption = " ‰«„ ( " & txtName.Text & " ) œ— ”Ì” „  ⁄—Ì› ‰‘œÂ"
                 Else
                    LblFindCust.Caption = "‰«„ „‘ —ﬂ —« Ê«—œ ﬂ‰Ìœ  "
                 End If
            Case 3:
                 If Len(txtPhone.Text) > 0 Then
                    LblFindCust.Caption = "  ·›‰ ( " & txtPhone.Text & " ) œ— ”Ì” „  ⁄—Ì› ‰‘œÂ"
                 Else
                    LblFindCust.Caption = " ·›‰ „‘ —ﬂ —« Ê«—œ ﬂ‰Ìœ  "
                 End If
            Case 4:
                 If Len(TxtAddress.Text) > 0 Then
                   LblFindCust.Caption = " ¬œ—” ( " & TxtAddress.Text & " ) œ— ”Ì” „  ⁄—Ì› ‰‘œÂ"
                 Else
                    LblFindCust.Caption = "¬œ—” «‘ —«ﬂ —« Ê«—œ ﬂ‰Ìœ  "
                 End If
        End Select
    End If
            
End Sub

Private Sub txtAddress_Change()

    If Option1(0).Value = False Then
        i = vsCustomer.FindRow(TxtAddress.Text, 1, 5, False, False)
        If i > 0 Then
            vsCustomer.Row = i
            vsCustomer.ShowCell i, 0
            LblFindCust.Caption = ""
        Else
            vsCustomer.Row = 0
            vsCustomer.ShowCell 0, 0
            labelClear
            If TxtAddress.Text <> "" Then
               LblFindCust.Caption = " ¬œ—” ( " & TxtAddress.Text & " )œ— ”Ì” „  ⁄—Ì› ‰‘œÂ"
             Else
                LblFindCust.Caption = "¬œ—” „‘ —ﬂ —« Ê«—œ ﬂ‰Ìœ  "
             End If
    
        End If
   Else
       If Len(TxtAddress.Text) > 0 Then
            SearchType = 4
            Timer1.Interval = Val(TxtTimer.Text)
            Timer1.Enabled = True
       Else
          vsCustomer.Rows = 1
          labelClear
      
       End If
  
   End If

End Sub

Private Sub txtAddress_GotFocus()

    txtPhone.Text = ""
    txtName.Text = ""
    txtMembershipId.Text = ""
    cmbPrefix.ListIndex = 0
    vsCustomer.Row = 0
    labelClear
    vsCustomer.Select vsCustomer.Row, 5
    vsCustomer.Sort = flexSortGenericAscending
    vsCustomer_AfterSort 5, 1
    LblFindCust.Caption = "¬œ—” „‘ —ﬂ —« Ê«—œ ﬂ‰Ìœ  "

End Sub

Private Sub TxtAddress_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 40 Then
    If vsCustomer.Row >= 1 Then
         vsCustomer.SetFocus
         If vsCustomer.Rows > 2 Then
            vsCustomer.Row = 2
         End If
    End If
End If

End Sub

Private Sub txtBarcode_Change()
    If Right(txtBarcode.Text, 1) = "/" Then
        txtBarcode.Text = left(txtBarcode.Text, Len(txtBarcode.Text) - 1)
    ElseIf left(txtBarcode.Text, 1) = "/" Then
        txtBarcode.Text = Right(txtBarcode.Text, Len(txtBarcode.Text) - 1)
    End If

End Sub


Private Sub txtMembershipId_Change()
    CmdFamilySet.Enabled = False
    If Option1(0).Value = False Then
        i = vsCustomer.FindRow(txtMembershipId.Text, 1, 2, True, True)
        If i > 0 Then
            vsCustomer.Row = i
            vsCustomer.ShowCell i, 0
            LblFindCust.Caption = ""
        Else
            vsCustomer.Row = 0
            vsCustomer.ShowCell 0, 0
            labelClear
            If Val(txtMembershipId.Text) > 0 Then
               LblFindCust.Caption = " «‘ —«ﬂ ( " & txtMembershipId.Text & " ) œ— ”Ì” „  ⁄—Ì› ‰‘œÂ"
             Else
                LblFindCust.Caption = "ﬂœ «‘ —«ﬂ —« Ê«—œ ﬂ‰Ìœ  "
             End If
        
        End If
   Else
       If Val(txtMembershipId.Text) > 0 Then
            SearchType = 1
            Timer1.Interval = Val(TxtTimer.Text)
            Timer1.Enabled = True
       Else
          vsCustomer.Rows = 1
          labelClear
        
       End If
  
   End If
    
End Sub

Private Sub txtMembershipId_GotFocus()
    txtPhone.Text = ""
    TxtAddress.Text = ""
    txtName.Text = ""
    cmbPrefix.ListIndex = 0
    vsCustomer.Row = 0
    labelClear
    vsCustomer.Select vsCustomer.Row, 2
    vsCustomer.Sort = flexSortGenericAscending
    vsCustomer_AfterSort 2, 1
    LblFindCust.Caption = "ﬂœ «‘ —«ﬂ —« Ê«—œ ﬂ‰Ìœ  "
    CmdFamilySet.Enabled = False

End Sub
Private Sub txtName_Change()

    If Option1(0).Value = False Then
        i = vsCustomer.FindRow(txtName.Text, 1, 3, False, False)
        If i > 0 Then
            vsCustomer.Row = i
            vsCustomer.ShowCell i, 0
            LblFindCust.Caption = ""
        Else
            vsCustomer.Row = 0
            vsCustomer.ShowCell 0, 0
            labelClear
            If txtName.Text <> "" Then
               LblFindCust.Caption = " ‰«„ ( " & txtName.Text & " ) œ— ”Ì” „  ⁄—Ì› ‰‘œÂ"
             Else
                LblFindCust.Caption = "‰«„ „‘ —ﬂ —« Ê«—œ ﬂ‰Ìœ  "
             End If
    
        End If
    Else
       If Len(txtName.Text) > 0 Then
            SearchType = 2
            Timer1.Interval = Val(TxtTimer.Text)
            Timer1.Enabled = True
       Else
          vsCustomer.Rows = 1
          labelClear

       End If
  
   End If

End Sub

Private Sub txtName_GotFocus()
    txtPhone.Text = ""
    TxtAddress.Text = ""
    txtMembershipId.Text = ""
    cmbPrefix.ListIndex = 0
    vsCustomer.Row = 0
    labelClear
    vsCustomer.Select vsCustomer.Row, 3
    vsCustomer.Sort = flexSortGenericAscending
    vsCustomer_AfterSort 3, 1
    LblFindCust.Caption = "‰«„ „‘ —ﬂ —« Ê«—œ ﬂ‰Ìœ  "

End Sub

Private Sub FillvsCustomer()
    Dim TmpTel As String
    Dim jj As Integer
    vsCustomer.Rows = 1
    LblCount = ""
    ReDim Parameter(1) As Parameter
    If VarActForm = "frmCust" Then
        Parameter(0) = GenerateInputParameter("@ActDeact", adInteger, 4, 2) ' All Customers
    Else
        Parameter(0) = GenerateInputParameter("@ActDeact", adInteger, 4, 0) ' Only Active
    End If
    Parameter(1) = GenerateInputParameter("@Branch", adInteger, 4, CurrentBranch)
    Set Rst = RunParametricStoredProcedure2Rec("Get_All_Customers", Parameter)
    
    If Rst.EOF = True And Rst.BOF = True Then Exit Sub
    
    With vsCustomer
        .ColHidden(1) = True
        .Rows = 1
        i = 0
        j = 0
        FWProgressBar1.Value = 0
        MousePointer = 11
        Static arr As Variant
        arr = Rst.GetRows
        
        ' reset the control
        .BindToArray Null
        '  SetDefaults fa
        
        ' set the properties we want
        
'        While Rst.EOF <> True
'
'            TmpTel = Rst!Tel1
'            j = j + 1
'            For jj = 1 To 5
'                If TmpTel <> "" Or (jj = 1) Then
'                    i = i + 1
'                    .Rows = .Rows + 1
'                    .TextMatrix(i, 0) = j
'                    .TextMatrix(i, 1) = Rst!Code
'                    .TextMatrix(i, 2) = Rst!MembershipId
'                    .TextMatrix(i, 3) = Rst![Name]
'                    .TextMatrix(i, 4) = Trim(TmpTel)
'                    .TextMatrix(i, 5) = IIf(IsNull(Rst!address), " ", Rst!address)
'                    .TextMatrix(i, 7) = Rst!Prefix
'                    TmpTel = ""
'                End If
'                If jj = 1 And Trim(Rst!Tel2) <> "" Then
'                    TmpTel = Rst!Tel2
'                ElseIf jj = 2 And Trim(Rst!Tel3) <> "" Then
'                    TmpTel = Rst!Tel3
'                ElseIf jj = 3 And Trim(Rst!Tel4) <> "" Then
'                    TmpTel = Rst!Tel4
'                ElseIf jj = 4 And Trim(Rst!Mobile) <> "" Then
'                    TmpTel = Rst!Mobile
'                End If
'            Next jj
''''            CustomerSellPrice = Rst!SellPrice
'            If i Mod 1000 = 0 Then DoEvents
'
'            Rst.MoveNext
'            If i Mod 100 = 0 Then
'                FWProgressBar1.Value = FWProgressBar1 + 1
'                If FWProgressBar1.Value = 100 Then
'                    FWProgressBar1.Value = 1
'                End If
'            End If
'            LblCount.Caption = " ⁄œ«œ —ﬂÊ—œÂ«   -  " & j
'        Wend
        
        .LoadArray arr
          
        For i = 1 To .Rows - 1
            .TextMatrix(i, 0) = i
        Next
        .Cols = 9
        .TextMatrix(0, 8) = " ›÷Ì·Ì"
        LblCount.Caption = " ⁄œ«œ —ﬂÊ—œÂ«   -  " & .Rows - 1
        MousePointer = 0
    
        Set Rst = RunStoredProcedure2RecordSet("Get_All_tPrefix")
        .ColComboList(7) = vsCustomer.BuildComboList(Rst, "Description", "Code")
        
        Set Rst = Nothing
    
    End With
    
'    vsCustomer.MergeCompare = flexMCTrimNoCase
'    vsCustomer.MergeCells = flexMergeRestrictRows
'    vsCustomer.MergeRow(vsCustomer.Rows - 1) = True
'    vsCustomer.MergeCol(0) = True
'    vsCustomer.MergeCol(1) = True
'    vsCustomer.MergeCol(2) = True
'    vsCustomer.MergeCol(3) = True
'    vsCustomer.MergeCol(5) = True
'    vsCustomer.MergeCol(7) = True
'    vsCustomer.ColWidth(3) = vsCustomer.ColWidth(3) * 1.1
'    vsCustomer.AutoSizeMode = flexAutoSizeColWidth
'    vsCustomer.AutoSize 0, vsCustomer.Cols - 1
    
'    If vsCustomer.ColWidth(2) < 800 Then
'        vsCustomer.ColWidth(2) = 800
'    End If
'    If vsCustomer.ColWidth(3) < 3000 Then
'        vsCustomer.ColWidth(3) = 3000
'    End If
'    If vsCustomer.ColWidth(4) < 1500 Then
'        vsCustomer.ColWidth(4) = 1500
'    End If
'    If vsCustomer.ColWidth(5) < 4000 Then
'        vsCustomer.ColWidth(5) = 4000
'    End If
   ' vsCustomer.ColIndent(0) = 1

End Sub


Private Sub txtName_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 40 Then
    If vsCustomer.Row >= 1 Then
         vsCustomer.SetFocus
         If vsCustomer.Rows > 2 Then
            vsCustomer.Row = 2
         End If
    End If
End If
End Sub

Private Sub txtPhone_Change()

    If Option1(0).Value = False Then
        i = vsCustomer.FindRow(txtPhone.Text, 1, 4, False, False)
        If i > 0 Then
            vsCustomer.Row = i
            vsCustomer.ShowCell i, 0
            LblFindCust.Caption = ""
        Else
            vsCustomer.Row = 0
            vsCustomer.ShowCell 0, 0
            labelClear
            If txtPhone.Text <> "" Then
               LblFindCust.Caption = "  ·›‰ ( " & txtPhone.Text & " ) œ— ”Ì” „  ⁄—Ì› ‰‘œÂ"
             Else
                LblFindCust.Caption = " ·›‰ „‘ —ﬂ —« Ê«—œ ﬂ‰Ìœ  "
             End If
    
        End If
    Else
       If Len(txtPhone.Text) > 0 Then
            SearchType = 3
            Timer1.Interval = Val(TxtTimer.Text)
            Timer1.Enabled = True
       Else
          vsCustomer.Rows = 1
          labelClear
      
       End If
  
   End If

'If txtPhone.Text = "" Then
'    MsgBox ""
'
'End If
End Sub

Private Sub txtPhone_GotFocus()
    If tmpflag = True Then
        tmpflag = False
        Exit Sub
    End If
    txtName.Text = ""
    TxtAddress.Text = ""
    txtMembershipId.Text = ""
    cmbPrefix.ListIndex = 0
    vsCustomer.Row = 0
    labelClear
    vsCustomer.Select vsCustomer.Row, 4
    vsCustomer.Sort = flexSortGenericAscending
    vsCustomer_AfterSort 4, 1
    LblFindCust.Caption = " ·›‰ „‘ —ﬂ —« Ê«—œ ﬂ‰Ìœ  "
    If Len(Call_RealNumber) > 0 Then
        txtPhone.Text = Call_RealNumber
        
    End If

End Sub

Private Sub txtPhone_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 40 Then
    If vsCustomer.Row >= 1 Then
         vsCustomer.SetFocus
         If vsCustomer.Rows > 2 Then
            vsCustomer.Row = 2
         End If
    End If
End If

End Sub

Private Sub TxtTimer_Change()
    If Val(TxtTimer.Text) < 1 Then TxtTimer.Text = "100"
    clsStation.SrarchInputDelayKeyboard = Val(TxtTimer.Text)
    SetStationSettingFile
End Sub

Private Sub vsCustomer_AfterSort(ByVal Col As Long, Order As Integer)
    j = 1
    For i = 1 To vsCustomer.Rows - 1
'        If i = vsCustomer.Rows - 1 Then
'            vsCustomer.TextMatrix(i, 0) = j
'            Exit For
'        End If
        vsCustomer.TextMatrix(i, 0) = i
'        If vsCustomer.TextMatrix(i, 3) <> vsCustomer.TextMatrix(i + 1, 3) Then
'            j = j + 1
'        End If
    Next
'    vsCustomer.MergeCells = flexMergeRestrictRows
'    vsCustomer.MergeRow(vsCustomer.Rows - 1) = True
'    vsCustomer.MergeCol(0) = True
'    vsCustomer.MergeCol(0) = True
    
End Sub

Private Sub vsCustomer_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    
    If Col = -1 Then Exit Sub
    For i = 0 To vsCustomer.Cols - 1
        SaveSetting strMainKey, "frmFindCust_vsCustomer", "Col" & i, vsCustomer.ColWidth(i)
    Next

End Sub

Private Sub vsCustomer_DblClick()
    
    If OKButton.Visible = False Then Exit Sub
    If vsCustomer.Row > 0 Then
        OKButton_Click
    End If
End Sub

Private Sub Define_Customer()
    
    TxtFamilyNo.Text = ""
    LblLastBuy.Caption = "¬Œ—Ì‰ Œ—Ìœ"
    labelClear
    LblCount = ""
    vsCustomer.Rows = 1
    ReDim Parameter(1) As Parameter
    
    If VarActForm = "frmCust" Then
        Parameter(0) = GenerateInputParameter("@ActDeact", adInteger, 4, 2) ' All Customers
    Else
        Parameter(0) = GenerateInputParameter("@ActDeact", adInteger, 4, 0) ' Only Active
    End If
    
    Select Case SearchType
        Case 1
            Parameter(1) = GenerateInputParameter("@Code", adBigInt, 8, Val(txtMembershipId.Text))
            Set Rst = RunParametricStoredProcedure2Rec("Get_Customer_Code", Parameter)
        Case 2
            Parameter(1) = GenerateInputParameter("@Name", adVarWChar, 50, left(txtName.Text, 50))
            Set Rst = RunParametricStoredProcedure2Rec("Get_Customer_Name", Parameter)
        Case 3
            Parameter(1) = GenerateInputParameter("@Tel", adVarWChar, 20, left(txtPhone.Text, 20))
            Set Rst = RunParametricStoredProcedure2Rec("Get_Customer_Tel", Parameter)
        Case 4
            Parameter(1) = GenerateInputParameter("@Addresse", adVarWChar, 100, left(TxtAddress.Text, 100))
            Set Rst = RunParametricStoredProcedure2Rec("Get_Customer_Address", Parameter)
        Case 5
            Parameter(1) = GenerateInputParameter("@Code", adInteger, 4, cmbPrefix.ItemData(cmbPrefix.ListIndex))
            Set Rst = RunParametricStoredProcedure2Rec("Get_Customer_Prefix", Parameter)
    End Select
    
'Debug.Print txtPhone.Text & "1*"
    CmdFamilySet.Enabled = False
    If Rst.EOF = True And Rst.BOF = True Then Exit Sub
    Dim TmpTel As String
    Dim jj As Integer
    
    With vsCustomer
        .Rows = 1
        i = 0
        j = 0
        Static arr As Variant
        arr = Rst.GetRows
        
        ' reset the control
        .BindToArray Null
        '  SetDefaults fa
        
        ' set the properties we want
        
        CmdFamilySet.Enabled = True
'        Do While Rst.EOF <> True
'            If SearchType = 1 Then
'                TxtFamilyNo.Text = IIf(IsNull(Rst!FamilyNo), "", Rst!FamilyNo)
'            End If
'
'            TmpTel = Rst!Tel1
'            j = j + 1
'
'            For jj = 1 To 5
'                If TmpTel <> "" Or (jj = 1) Then
'                    i = i + 1
'                    .Rows = .Rows + 1
'                    .TextMatrix(i, 0) = j
'                    .TextMatrix(i, 1) = Rst!Code
'                    .TextMatrix(i, 2) = Rst!MembershipId
'                    .TextMatrix(i, 3) = Rst![Name]
'                    .TextMatrix(i, 4) = Left(Trim(TmpTel), 11)
'                    .TextMatrix(i, 5) = IIf(IsNull(Rst!address), " ", Rst!address)
'                    .TextMatrix(i, 7) = Rst!Prefix
'''''                    .TextMatrix(i, 7) = Rst!Credit
'''''                    .TextMatrix(i, 8) = Rst!carryfee
'''''                    .TextMatrix(i, 9) = Rst!PaykFee
'''''                    .TextMatrix(i, 10) = Rst!Distance
'                    TmpTel = ""
'                End If
'                If jj = 1 And Trim(Rst!Tel2) <> "" Then
'                    TmpTel = Rst!Tel2
'                ElseIf jj = 2 And Trim(Rst!Tel3) <> "" Then
'                    TmpTel = Rst!Tel3
'                ElseIf jj = 3 And Trim(Rst!Tel4) <> "" Then
'                    TmpTel = Rst!Tel4
'                ElseIf jj = 4 And Trim(Rst!Mobile) <> "" Then
'                    TmpTel = Rst!Mobile
'                End If
'            Next jj
             
            
'            If i > Val(txtMaxRecord.Text) Then Exit Do
'            Rst.MoveNext
'        Loop
'
'        vsCustomer.MergeCompare = flexMCTrimNoCase
'        vsCustomer.MergeCells = flexMergeRestrictRows
'        vsCustomer.MergeRow(vsCustomer.Rows - 1) = True
'        vsCustomer.MergeCol(0) = True
'        vsCustomer.MergeCol(1) = True
'        vsCustomer.MergeCol(2) = True
'        vsCustomer.MergeCol(3) = True
'    ''''    vsCustomer.MergeCol(4) = True
'        vsCustomer.MergeCol(5) = True
'        vsCustomer.MergeCol(7) = True
'
'        If i > 0 Then
'            LblCount.Caption = " ⁄œ«œ —ﬂÊ—œÂ«   -  " & i
'        Else
'            LblCount.Caption = ""
'        End If
'
'        vsCustomer.AutoSizeMode = flexAutoSizeColWidth
'        vsCustomer.AutoSize 0, .Cols - 1
'    If vsCustomer.ColWidth(3) < 3000 Then
'        vsCustomer.ColWidth(3) = 3000
'    End If
'
'    If vsCustomer.ColWidth(4) < 1500 Then
'        vsCustomer.ColWidth(4) = 1500
'    End If
'
'    If vsCustomer.ColWidth(5) < 4000 Then
'        vsCustomer.ColWidth(5) = 4000
'    End If
'
'     If vsCustomer.ColWidth(2) < 800 Then
'        vsCustomer.ColWidth(2) = 800
'    End If
        .LoadArray arr
          
        For i = 1 To .Rows - 1
            .TextMatrix(i, 0) = i
        Next
        
        LblCount.Caption = " ⁄œ«œ —ﬂÊ—œÂ«   -  " & .Rows - 1
        .Cols = 9
        .TextMatrix(0, 8) = " ›÷Ì·Ì"
        LblResult.Visible = False
        If .Rows > 1 Then .ShowCell 1, 0: vsCustomer_RowColChange
        Set Rst = RunStoredProcedure2RecordSet("Get_All_tPrefix")
        .ColComboList(7) = vsCustomer.BuildComboList(Rst, "Description", "Code")
        
        Set Rst = Nothing
    End With

End Sub


Private Sub vsCustomer_RowColChange()
    If vsCustomer.Row < 1 Then Exit Sub
    If strCategory = "07" Then
        ReDim Parameter(1) As Parameter
        Parameter(0) = GenerateInputParameter("@Code", adBigInt, 8, vsCustomer.TextMatrix(vsCustomer.Row, 1))
        Parameter(1) = GenerateOutputParameter("@Result", adBigInt, 8)
        LblLastBuy.Caption = LblLastBuy.Caption & " " & vbLf & RunParametricStoredProcedure("Get_LastBuy", Parameter)
    End If
    If vsCustomer.Row = 0 Or Val(vsCustomer.TextMatrix(vsCustomer.Row, 1)) = -1 Then Exit Sub
    ReDim Parameter(0) As Parameter
    Parameter(0) = GenerateInputParameter("@Code", adBigInt, 8, vsCustomer.TextMatrix(vsCustomer.Row, 1))
    Set Rst = RunParametricStoredProcedure2Rec("Get_BuyCustomer", Parameter)
    If Not (Rst.EOF = True And Rst.BOF = True) Then
        LastDate1.Caption = Rst!LastDate
        LastNo1.Caption = Rst!LastNo
        BuyAverage1.Caption = Rst!BuyAverage
        BuyCount1.Caption = Rst!BuyCount
        MaxPrice1.Caption = Rst!MaxPrice
        MinPrice1.Caption = Rst!MinPrice
        AddedDate1.Caption = Rst!AddedDate
        LastPrice1.Caption = Rst!LastPrice
        LblBuy1.Caption = Rst!CreditBuy
        LblRecieve1.Caption = Rst!RecievedAmount
        LastCredit1.Caption = Rst!CreditBuy - Rst!RecievedAmount
        LbCustomerlRate.Caption = Rst!SellPrice
        lblCountCurrentBuy.Caption = Rst!CountCurrentDayBuy
        If Val(LastCredit1.Caption) > 0 Then
            LastCredit.Caption = "»œÂÌ"
        Else
            LastCredit.Caption = "ÿ·»"
            LastCredit1.Caption = -1 * LastCredit1.Caption
        End If
        Set Rst = Nothing
        
        ReDim Parameter(1) As Parameter
        Parameter(0) = GenerateInputParameter("@Code", adBigInt, 8, vsCustomer.TextMatrix(vsCustomer.Row, 1))
        Parameter(1) = GenerateInputParameter("@Date", adVarChar, 50, Mid(clsDate.shamsi(Date), 3))
        Set Rst = RunParametricStoredProcedure2Rec("Get_BuyDailyCustomer", Parameter)
        If Not (Rst.EOF = True And Rst.BOF = True) Then
           CountDailyBuy = Rst!CountDailyBuy
           CountShiftBuy = Rst!CountShiftBuy
        End If
    End If
    
    Call ShowCustPicture
    Call GetCreditDebit

End Sub

Private Sub labelClear()
    LastDate1.Caption = ""
    LastNo1.Caption = ""
    BuyAverage1.Caption = ""
    BuyCount1.Caption = ""
    MaxPrice1.Caption = ""
    MinPrice1.Caption = ""
    AddedDate1.Caption = ""
    LastPrice1.Caption = ""
    LblBuy1.Caption = ""
    LblRecieve1.Caption = ""
    LastCredit1.Caption = ""
    lblCountCurrentBuy.Caption = ""
    LbCustomerlRate.Caption = ""
    
    Image5.Visible = True
    Image4.Visible = False
    Image4.Picture = LoadPicture("")

End Sub


Private Sub GetCreditDebit()
    On Error GoTo Err_Handler
    Dim TotalBedehkar, TotalBestankar As Double
    Dim L_Rst As New ADODB.Recordset
    
    Me.LblTotalCreditDebit.Caption = ""
    If vsCustomer.Row < 1 Then Exit Sub
    If clsArya.ExternalAccounting = True And Val(vsCustomer.TextMatrix(vsCustomer.Row, 8)) > 0 Then
        Set L_Rst = Accounting.GetCreditDebitDll(Val(vsCustomer.TextMatrix(vsCustomer.Row, 8)), 0)
        If L_Rst.BOF = True And L_Rst.EOF = True Then
            Set L_Rst = Nothing
            Exit Sub
        Else
            TotalCreditDebit = 0: TotalBedehkar = 0: TotalBestankar = 0
            While L_Rst.EOF = False
                TotalBedehkar = TotalBedehkar + L_Rst.Fields("Bedehkar").Value
                TotalBestankar = TotalBestankar + L_Rst.Fields("Bestankar").Value
                L_Rst.MoveNext
            Wend
        
        End If
        TotalCreditDebit = TotalBedehkar - TotalBestankar
        L_Rst.Close
        Set L_Rst = Nothing
        If TotalCreditDebit > 0 Then
            LblTotalCreditDebitLabel = "»œÂÌ ﬂ·: "
            LblTotalCreditDebit = Format(TotalCreditDebit, "#,##") & clsArya.UnitPrice
            LblTotalCreditDebitLabel.ForeColor = vbRed
            LblTotalCreditDebit.ForeColor = vbRed
        ElseIf TotalCreditDebit = 0 Then
            LblTotalCreditDebitLabel.Caption = "»œÂÌ- ÿ·» ﬂ·: "
            LblTotalCreditDebit.Caption = Format(TotalCreditDebit, "#,##") & clsArya.UnitPrice
            LblTotalCreditDebitLabel.ForeColor = vbGreen
            LblTotalCreditDebit.ForeColor = vbGreen
        Else
            TotalCreditDebit = Abs(TotalCreditDebit)
            LblTotalCreditDebitLabel.Caption = "ÿ·» ﬂ·: "
            LblTotalCreditDebit.Caption = Format((TotalCreditDebit), "#,##") & clsArya.UnitPrice
            LblTotalCreditDebitLabel.ForeColor = vbGreen
            LblTotalCreditDebit.ForeColor = vbGreen
            
        End If
        
    End If
    
Exit Sub

Err_Handler:
    LogSaveNew "frmFindCust => ", err.Description, err.Number, err.Source, "GetCreditDebit"
    ShowErrorMessage
    err.Clear
    Resume Next
    Set L_Rst = Nothing
End Sub


