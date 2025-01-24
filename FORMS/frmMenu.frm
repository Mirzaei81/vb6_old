VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{7AEDC602-D94C-11D1-BB7A-00E0290EA3C9}#1.0#0"; "ResizeKit.ocx"
Object = "{75D4F767-8785-11D3-93AD-0000832EF44D}#2.14#0"; "FAST2001.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmMenu 
   Caption         =   "             ⁄—Ì› „‰ÊÂ«Ì ﬂ«·«  "
   ClientHeight    =   9870
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14025
   Icon            =   "frmMenu.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9870
   ScaleWidth      =   14025
   Begin VB.CommandButton cmdCopyMenu 
      Caption         =   "òÅÌ „‰ÊÂ« »Â"
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
      Left            =   8040
      TabIndex        =   72
      Top             =   240
      Width           =   1335
   End
   Begin VB.Frame FrameGroup 
      BackColor       =   &H0080FFFF&
      Caption         =   "‰«„ ê—ÊÂÂ«"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6135
      Left            =   5400
      RightToLeft     =   -1  'True
      TabIndex        =   62
      Top             =   1200
      Visible         =   0   'False
      Width           =   3735
      Begin VB.CommandButton CmdDoneGroup 
         Caption         =   " «∆Ìœ"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2520
         TabIndex        =   66
         Top             =   3720
         Width           =   1095
      End
      Begin VB.CommandButton Command2 
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
         Height          =   615
         Left            =   120
         TabIndex        =   65
         Top             =   3720
         Width           =   1095
      End
      Begin VB.ListBox lstGroups 
         BeginProperty Font 
            Name            =   "Nazanin"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   2760
         Left            =   120
         RightToLeft     =   -1  'True
         Style           =   1  'Checkbox
         TabIndex        =   64
         Top             =   840
         Width           =   3555
      End
      Begin VB.TextBox txtGroupName 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Nazanin"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   540
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   63
         Top             =   360
         Width           =   3555
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0080FFFF&
         Caption         =   $"frmMenu.frx":A4C2
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   67
         Top             =   4560
         Width           =   3495
      End
   End
   Begin VB.PictureBox frameMenu 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   5595
      Left            =   7320
      RightToLeft     =   -1  'True
      ScaleHeight     =   5565
      ScaleWidth      =   6600
      TabIndex        =   68
      Top             =   960
      Width           =   6630
      Begin VB.CommandButton BtnMenu 
         BeginProperty Font 
            Name            =   "Nazanin"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Index           =   1
         Left            =   120
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   69
         Top             =   600
         Width           =   1455
      End
      Begin TabDlg.SSTab SSTab1 
         Height          =   540
         Left            =   0
         TabIndex        =   70
         Top             =   0
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   953
         _Version        =   393216
         Tabs            =   1
         TabsPerRow      =   1
         TabHeight       =   882
         ForeColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Nazanin"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "ê—ÊÂ 1"
         TabPicture(0)   =   "frmMenu.frx":A5AC
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).ControlCount=   0
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   " ⁄—Ì› ‰«„ ê—ÊÂÂ«"
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
      Left            =   9480
      TabIndex        =   61
      Top             =   240
      Width           =   1455
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   3135
      Left            =   7200
      ScaleHeight     =   3105
      ScaleWidth      =   6705
      TabIndex        =   8
      Top             =   6720
      Width           =   6735
      Begin VB.Frame Frame3 
         Appearance      =   0  'Flat
         Caption         =   "‰„Ê‰Â"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   1650
         Left            =   4320
         RightToLeft     =   -1  'True
         TabIndex        =   59
         Top             =   0
         Width           =   1215
         Begin VB.CommandButton cmdPatern 
            Height          =   1050
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   60
            Top             =   360
            Width           =   930
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFC0C0&
         Height          =   3015
         Left            =   0
         TabIndex        =   9
         Top             =   0
         Width           =   3840
         Begin VB.OptionButton OptColor 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00400040&
            Height          =   450
            Index           =   48
            Left            =   3360
            Style           =   1  'Graphical
            TabIndex        =   57
            Top             =   2520
            Width           =   450
         End
         Begin VB.OptionButton OptColor 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00800080&
            Height          =   450
            Index           =   47
            Left            =   3360
            Style           =   1  'Graphical
            TabIndex        =   56
            Top             =   2040
            Width           =   450
         End
         Begin VB.OptionButton OptColor 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C000C0&
            Height          =   450
            Index           =   46
            Left            =   3360
            Style           =   1  'Graphical
            TabIndex        =   55
            Top             =   1560
            Width           =   450
         End
         Begin VB.OptionButton OptColor 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FF00FF&
            Height          =   450
            Index           =   45
            Left            =   3360
            Style           =   1  'Graphical
            TabIndex        =   54
            Top             =   1080
            Width           =   450
         End
         Begin VB.OptionButton OptColor 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FF80FF&
            Height          =   450
            Index           =   44
            Left            =   3360
            Style           =   1  'Graphical
            TabIndex        =   53
            Top             =   600
            Width           =   450
         End
         Begin VB.OptionButton OptColor 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFC0FF&
            Height          =   450
            Index           =   43
            Left            =   3360
            Style           =   1  'Graphical
            TabIndex        =   52
            Top             =   120
            Width           =   450
         End
         Begin VB.OptionButton OptColor 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00400000&
            Height          =   450
            Index           =   42
            Left            =   2880
            Style           =   1  'Graphical
            TabIndex        =   51
            Top             =   2520
            Width           =   450
         End
         Begin VB.OptionButton OptColor 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00800000&
            Height          =   450
            Index           =   41
            Left            =   2880
            Style           =   1  'Graphical
            TabIndex        =   50
            Top             =   2040
            Width           =   450
         End
         Begin VB.OptionButton OptColor 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C00000&
            Height          =   450
            Index           =   40
            Left            =   2880
            Style           =   1  'Graphical
            TabIndex        =   49
            Top             =   1560
            Width           =   450
         End
         Begin VB.OptionButton OptColor 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FF0000&
            Height          =   450
            Index           =   39
            Left            =   2880
            Style           =   1  'Graphical
            TabIndex        =   48
            Top             =   1080
            Width           =   450
         End
         Begin VB.OptionButton OptColor 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FF8080&
            Height          =   450
            Index           =   38
            Left            =   2880
            Style           =   1  'Graphical
            TabIndex        =   47
            Top             =   600
            Width           =   450
         End
         Begin VB.OptionButton OptColor 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFC0C0&
            Height          =   450
            Index           =   37
            Left            =   2880
            Style           =   1  'Graphical
            TabIndex        =   46
            Top             =   120
            Width           =   450
         End
         Begin VB.OptionButton OptColor 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Height          =   450
            Index           =   1
            Left            =   0
            Style           =   1  'Graphical
            TabIndex        =   45
            Top             =   120
            Width           =   450
         End
         Begin VB.OptionButton OptColor 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            Height          =   450
            Index           =   2
            Left            =   0
            Style           =   1  'Graphical
            TabIndex        =   44
            Top             =   600
            Width           =   450
         End
         Begin VB.OptionButton OptColor 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C0C0&
            Height          =   450
            Index           =   3
            Left            =   0
            Style           =   1  'Graphical
            TabIndex        =   43
            Top             =   1080
            Width           =   450
         End
         Begin VB.OptionButton OptColor 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00808080&
            Height          =   450
            Index           =   4
            Left            =   0
            Style           =   1  'Graphical
            TabIndex        =   42
            Top             =   1560
            Width           =   450
         End
         Begin VB.OptionButton OptColor 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00404040&
            Height          =   450
            Index           =   5
            Left            =   0
            Style           =   1  'Graphical
            TabIndex        =   41
            Top             =   2040
            Width           =   450
         End
         Begin VB.OptionButton OptColor 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00000000&
            Height          =   450
            Index           =   6
            Left            =   0
            Style           =   1  'Graphical
            TabIndex        =   40
            Top             =   2520
            Width           =   450
         End
         Begin VB.OptionButton OptColor 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C0FF&
            Height          =   450
            Index           =   7
            Left            =   480
            Style           =   1  'Graphical
            TabIndex        =   39
            Top             =   120
            Width           =   450
         End
         Begin VB.OptionButton OptColor 
            Alignment       =   1  'Right Justify
            BackColor       =   &H008080FF&
            Height          =   450
            Index           =   8
            Left            =   480
            Style           =   1  'Graphical
            TabIndex        =   38
            Top             =   600
            Width           =   450
         End
         Begin VB.OptionButton OptColor 
            Alignment       =   1  'Right Justify
            BackColor       =   &H000000FF&
            Height          =   450
            Index           =   9
            Left            =   480
            Style           =   1  'Graphical
            TabIndex        =   37
            Top             =   1080
            Width           =   450
         End
         Begin VB.OptionButton OptColor 
            Alignment       =   1  'Right Justify
            BackColor       =   &H000000C0&
            Height          =   450
            Index           =   10
            Left            =   480
            Style           =   1  'Graphical
            TabIndex        =   36
            Top             =   1560
            Width           =   450
         End
         Begin VB.OptionButton OptColor 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00000080&
            Height          =   450
            Index           =   11
            Left            =   480
            Style           =   1  'Graphical
            TabIndex        =   35
            Top             =   2040
            Width           =   450
         End
         Begin VB.OptionButton OptColor 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00000040&
            Height          =   450
            Index           =   12
            Left            =   480
            Style           =   1  'Graphical
            TabIndex        =   34
            Top             =   2520
            Width           =   450
         End
         Begin VB.OptionButton OptColor 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0E0FF&
            Height          =   450
            Index           =   13
            Left            =   960
            Style           =   1  'Graphical
            TabIndex        =   33
            Top             =   120
            Width           =   450
         End
         Begin VB.OptionButton OptColor 
            Alignment       =   1  'Right Justify
            BackColor       =   &H0080C0FF&
            Height          =   450
            Index           =   14
            Left            =   960
            Style           =   1  'Graphical
            TabIndex        =   32
            Top             =   600
            Width           =   450
         End
         Begin VB.OptionButton OptColor 
            Alignment       =   1  'Right Justify
            BackColor       =   &H000080FF&
            Height          =   450
            Index           =   15
            Left            =   960
            Style           =   1  'Graphical
            TabIndex        =   31
            Top             =   1080
            Width           =   450
         End
         Begin VB.OptionButton OptColor 
            Alignment       =   1  'Right Justify
            BackColor       =   &H000040C0&
            Height          =   450
            Index           =   16
            Left            =   960
            Style           =   1  'Graphical
            TabIndex        =   30
            Top             =   1560
            Width           =   450
         End
         Begin VB.OptionButton OptColor 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00004080&
            Height          =   450
            Index           =   17
            Left            =   960
            Style           =   1  'Graphical
            TabIndex        =   29
            Top             =   2040
            Width           =   450
         End
         Begin VB.OptionButton OptColor 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00404080&
            Height          =   450
            Index           =   18
            Left            =   960
            Style           =   1  'Graphical
            TabIndex        =   28
            Top             =   2520
            Width           =   450
         End
         Begin VB.OptionButton OptColor 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFFF&
            Height          =   450
            Index           =   19
            Left            =   1440
            Style           =   1  'Graphical
            TabIndex        =   27
            Top             =   120
            Width           =   450
         End
         Begin VB.OptionButton OptColor 
            Alignment       =   1  'Right Justify
            BackColor       =   &H0080FFFF&
            Height          =   450
            Index           =   20
            Left            =   1440
            Style           =   1  'Graphical
            TabIndex        =   26
            Top             =   600
            Width           =   450
         End
         Begin VB.OptionButton OptColor 
            Alignment       =   1  'Right Justify
            BackColor       =   &H0000FFFF&
            Height          =   450
            Index           =   21
            Left            =   1440
            Style           =   1  'Graphical
            TabIndex        =   25
            Top             =   1080
            Width           =   450
         End
         Begin VB.OptionButton OptColor 
            Alignment       =   1  'Right Justify
            BackColor       =   &H0000C0C0&
            Height          =   450
            Index           =   22
            Left            =   1440
            Style           =   1  'Graphical
            TabIndex        =   24
            Top             =   1560
            Width           =   450
         End
         Begin VB.OptionButton OptColor 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00008080&
            Height          =   450
            Index           =   23
            Left            =   1440
            Style           =   1  'Graphical
            TabIndex        =   23
            Top             =   2040
            Width           =   450
         End
         Begin VB.OptionButton OptColor 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00004040&
            Height          =   450
            Index           =   24
            Left            =   1440
            Style           =   1  'Graphical
            TabIndex        =   22
            Top             =   2520
            Width           =   450
         End
         Begin VB.OptionButton OptColor 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFC0&
            Height          =   450
            Index           =   25
            Left            =   1920
            Style           =   1  'Graphical
            TabIndex        =   21
            Top             =   120
            Width           =   450
         End
         Begin VB.OptionButton OptColor 
            Alignment       =   1  'Right Justify
            BackColor       =   &H0080FF80&
            Height          =   450
            Index           =   26
            Left            =   1920
            Style           =   1  'Graphical
            TabIndex        =   20
            Top             =   600
            Width           =   450
         End
         Begin VB.OptionButton OptColor 
            Alignment       =   1  'Right Justify
            BackColor       =   &H0000FF00&
            Height          =   450
            Index           =   27
            Left            =   1920
            Style           =   1  'Graphical
            TabIndex        =   19
            Top             =   1080
            Width           =   450
         End
         Begin VB.OptionButton OptColor 
            Alignment       =   1  'Right Justify
            BackColor       =   &H0000C000&
            Height          =   450
            Index           =   28
            Left            =   1920
            Style           =   1  'Graphical
            TabIndex        =   18
            Top             =   1560
            Width           =   450
         End
         Begin VB.OptionButton OptColor 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00008000&
            Height          =   450
            Index           =   29
            Left            =   1920
            Style           =   1  'Graphical
            TabIndex        =   17
            Top             =   2040
            Width           =   450
         End
         Begin VB.OptionButton OptColor 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00004000&
            Height          =   450
            Index           =   30
            Left            =   1920
            Style           =   1  'Graphical
            TabIndex        =   16
            Top             =   2520
            Width           =   450
         End
         Begin VB.OptionButton OptColor 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Height          =   450
            Index           =   31
            Left            =   2400
            Style           =   1  'Graphical
            TabIndex        =   15
            Top             =   120
            Width           =   450
         End
         Begin VB.OptionButton OptColor 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFF80&
            Height          =   450
            Index           =   32
            Left            =   2400
            Style           =   1  'Graphical
            TabIndex        =   14
            Top             =   600
            Width           =   450
         End
         Begin VB.OptionButton OptColor 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFF00&
            Height          =   450
            Index           =   33
            Left            =   2400
            Style           =   1  'Graphical
            TabIndex        =   13
            Top             =   1080
            Width           =   450
         End
         Begin VB.OptionButton OptColor 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0C000&
            Height          =   450
            Index           =   34
            Left            =   2400
            Style           =   1  'Graphical
            TabIndex        =   12
            Top             =   1560
            Width           =   450
         End
         Begin VB.OptionButton OptColor 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00808000&
            Height          =   450
            Index           =   35
            Left            =   2400
            Style           =   1  'Graphical
            TabIndex        =   11
            Top             =   2040
            Width           =   450
         End
         Begin VB.OptionButton OptColor 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00404000&
            Height          =   450
            Index           =   36
            Left            =   2400
            Style           =   1  'Graphical
            TabIndex        =   10
            Top             =   2520
            Width           =   450
         End
      End
      Begin FLWCtrls.FWRealButton FWBOK 
         Height          =   855
         Left            =   3840
         TabIndex        =   58
         Top             =   1920
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   1508
         Caption         =   "À»  —‰ê ê—ÊÂ Ã«—Ì"
         ForeColor       =   -2147483641
         FontName        =   "B Homa"
         FontSize        =   12
      End
   End
   Begin VB.ComboBox cboStations 
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
      Left            =   11040
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   240
      Width           =   1755
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   90
      Top             =   1680
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   1320
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":A5C8
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":AEA4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":B780
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":C05C
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":C938
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":D214
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":DAF0
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":E1AC
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMenu.frx":E4C8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Height          =   840
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7305
      _ExtentX        =   12885
      _ExtentY        =   1482
      ButtonWidth     =   2249
      ButtonHeight    =   1429
      Wrappable       =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   8
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "«Œ ’«’ ﬂ«·«"
            Object.ToolTipText     =   "«Œ ’«’ ﬂ«·« »Â „‰ÊÂ«"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Ã«»Ã«ÌÌ „‰ÊÂ«"
            Object.ToolTipText     =   "Ã«»Ã«ÌÌ „‰ÊÂ«"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Õ–› ﬂ«·« «“ „‰Ê"
            Object.ToolTipText     =   "Õ–› ﬂ«·« «“ „‰Ê"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   " €ÌÌ— ‰«„ „‰Ê"
            Object.ToolTipText     =   " €ÌÌ— ‰«„ „‰Ê"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "«Œ ’«’ ¬ÌﬂÊ‰ "
            ImageIndex      =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Õ–› ¬ÌòÊ‰"
            ImageIndex      =   5
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.Frame Frame1 
      Height          =   9015
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   7095
      Begin VB.CommandButton cmdDone 
         BackColor       =   &H00008000&
         Caption         =   "À»  ﬂ«·« —ÊÌ „‰Ê"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         MaskColor       =   &H000000FF&
         TabIndex        =   2
         Top             =   4200
         Width           =   1905
      End
      Begin VSFlex7LCtl.VSFlexGrid vsNoKeyBoardDefined 
         Height          =   3585
         Left            =   0
         TabIndex        =   3
         Top             =   5280
         Width           =   6885
         _cx             =   12144
         _cy             =   6324
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
         BackColor       =   -2147483624
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483643
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483643
         BackColorAlternate=   12648447
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
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   ""
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
         ExplorerBar     =   1
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
      Begin VSFlex7LCtl.VSFlexGrid vsKeyBoardDefined 
         Height          =   3345
         Left            =   120
         TabIndex        =   4
         Top             =   600
         Width           =   6765
         _cx             =   11933
         _cy             =   5900
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
         BackColor       =   -2147483624
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483643
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483643
         BackColorAlternate=   12648447
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
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   ""
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
         ExplorerBar     =   1
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
      Begin FLWCtrls.FWLabel3D FWLabel3D1 
         Height          =   375
         Left            =   2520
         Top             =   120
         Width           =   3630
         _ExtentX        =   6403
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor1      =   128
         Caption         =   "        : ﬂ«·«Â«Ì „⁄—›Ì ‘œÂ —ÊÌ „‰ÊÌ ﬂ«·« "
         Alignment       =   1
      End
      Begin FLWCtrls.FWLabel3D FWLabel3D2 
         Height          =   435
         Left            =   3360
         Top             =   4800
         Width           =   3330
         _ExtentX        =   5874
         _ExtentY        =   767
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor1      =   32768
         Caption         =   "    : ﬂ«·«Â«Ì „⁄—›Ì ‰‘œÂ —ÊÌ „‰ÊÌ ﬂ«·« "
         Alignment       =   1
      End
      Begin FLWCtrls.FWLabel3D FWLabel3D3 
         Height          =   375
         Left            =   3600
         Top             =   4200
         Width           =   3030
         _ExtentX        =   5345
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor1      =   128
         Caption         =   "Õœ«ﬂÀ—  « 40 ﬂ«·« —ÊÌ Â— „‰Ê"
         Alignment       =   1
      End
   End
   Begin RESIZEKITLibCtl.ResizeKit ResizeKit1 
      Height          =   480
      Left            =   9720
      OleObjectBlob   =   "frmMenu.frx":F75D
      TabIndex        =   7
      Top             =   0
      Width           =   480
   End
   Begin FLWCtrls.FWNumericTextBox FWStationNo 
      Height          =   495
      Left            =   7560
      TabIndex        =   71
      Top             =   240
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      Max             =   2
      Min             =   1
      Value           =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "‰«„ «Ì” ê«Â"
      BeginProperty Font 
         Name            =   "B Homa"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   12720
      TabIndex        =   6
      Top             =   240
      Width           =   1125
   End
   Begin VB.Shape Shape4 
      BackColor       =   &H00008000&
      FillColor       =   &H8000000F&
      FillStyle       =   0  'Solid
      Height          =   795
      Left            =   7440
      Shape           =   4  'Rounded Rectangle
      Top             =   50
      Width           =   6465
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'################# MenuBar #############################
Dim lastPosition As Position
Dim MaxBtnMenu As Long
Dim BtnMenuPerFrame As Long
Private Type Position
    X As Single
    Y As Single
End Type
Dim IndexMenuTab

Dim MyFormAddEditMode As EnumMenuEditMode
Dim MyFormAddEditMode2 As EnumAddEditMode
Dim TempButton As Integer 'for keeping the number of the last pressed key
Dim i As Integer
Dim NotSuportedGoodType As EnumGoodType
Dim tmpStationNo As Integer
Dim Parameter() As Parameter

Public Sub Cancel()

    For i = 1 To Toolbar1.Buttons.Count
        Toolbar1.Buttons(i).Enabled = True
    Next i
    
    MyFormAddEditMode = EnumMenuEditMode.ViewButton
    TempButton = 0
    UpdateToolbars
    
    DefaultSetting
    MyFormAddEditMode2 = ViewMode
    SetFirstToolBar
    
End Sub
Public Sub Add()

    MyFormAddEditMode2 = AddMode
    SetFirstToolBar
    
End Sub

Public Sub ExitForm()

    Unload Me
    
End Sub
Public Sub ChangeLanguage()

   UpdateToolbars

End Sub

Public Sub FillvsNoKeyBoardDefined()

    Dim Rst As New ADODB.Recordset
   
    With vsNoKeyBoardDefined
        .Rows = 1
        
    ReDim Parameter(3) As Parameter
    
    Parameter(0) = GenerateInputParameter("@FactorType", adInteger, 4, EnumFactorType.Invoice)
    Parameter(1) = GenerateInputParameter("@StationId", adInteger, 4, tmpStationNo)
    Parameter(2) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
    Parameter(3) = GenerateInputParameter("@NotSuportedGoodType", adInteger, 4, NotSuportedGoodType)
    
    Set Rst = RunParametricStoredProcedure2Rec("Get_Good_Level", Parameter)
        
        If Rst.EOF = True And Rst.BOF = True Then
            Set Rst = Nothing
            Exit Sub
        End If
'        Rst.moveFirst
        While Rst.EOF <> True
            .Rows = .Rows + 1
            i = .Rows - 1
            .TextMatrix(i, 0) = i
            .TextMatrix(i, 1) = Rst.Fields("Code").Value
            .TextMatrix(i, 2) = Rst.Fields("Level1").Value
            .TextMatrix(i, 3) = Rst.Fields("Level2").Value
            .TextMatrix(i, 4) = 0
            .TextMatrix(i, 5) = Rst.Fields("Deslevel1").Value
            .TextMatrix(i, 6) = Rst.Fields("Deslevel2").Value
            .TextMatrix(i, 7) = Rst.Fields("Name").Value
            
            
            Rst.MoveNext
        Wend
        
        .Row = 0
        .AutoSizeMode = flexAutoSizeColWidth
        .AutoSize 0, .Cols - 1
        
    End With
    Set Rst = Nothing
End Sub

Public Sub FillvsKeyBoardDefined(Optional index As Integer)
    Dim Rst As New ADODB.Recordset
   
    With vsKeyBoardDefined
        .Rows = 1

    ReDim Parameter(3) As Parameter
    
    Parameter(0) = GenerateInputParameter("@FactorType", adInteger, 4, EnumFactorType.Invoice)
    Parameter(1) = GenerateInputParameter("@StationId", adInteger, 4, tmpStationNo)
    Parameter(2) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
    Parameter(3) = GenerateInputParameter("@BtnNum", adInteger, 4, index)
    
      
    Set Rst = RunParametricStoredProcedure2Rec("Get_Good_Menu", Parameter)
        
        If Rst.EOF = True And Rst.BOF = True Then
            Set Rst = Nothing
            Exit Sub
        End If
'        Rst.moveFirst
        While Rst.EOF <> True
            .Rows = .Rows + 1
            i = .Rows - 1
            .TextMatrix(i, 0) = i
            .TextMatrix(i, 1) = Rst.Fields("GoodCode").Value
            .TextMatrix(i, 2) = Rst.Fields("Level1").Value
            .TextMatrix(i, 3) = Rst.Fields("Level2").Value
            .TextMatrix(i, 4) = 0
            .TextMatrix(i, 5) = Rst.Fields("Deslevel1").Value
            .TextMatrix(i, 6) = Rst.Fields("Deslevel2").Value
            .TextMatrix(i, 7) = Rst.Fields("Name").Value
            
            
            Rst.MoveNext
        Wend
        
        .Row = 0
        .AutoSizeMode = flexAutoSizeColWidth
        .AutoSize 0, .Cols - 1
        
    End With
    Set Rst = Nothing
End Sub

Public Sub UpdateToolbars()
    
    UpdateButtons
    FillvsKeyBoardDefined
    FillvsNoKeyBoardDefined
    
'    AllButton vbOff, True
'
'    mdifrm.Toolbar1.Buttons(23).Enabled = True
'    mdifrm.Toolbar1.Buttons(24).Enabled = True
'    mdifrm.Toolbar1.Buttons(25).Enabled = True
'    mdifrm.Toolbar1.Buttons(26).Enabled = True
'    mdifrm.Toolbar1.Buttons(27).Enabled = True
'    mdifrm.Toolbar1.Buttons(9).Enabled = True
'
    For i = 1 To Toolbar1.Buttons.Count
         Toolbar1.Buttons.Item(i).Enabled = True
    Next i
    
    
    Select Case MyFormAddEditMode
        Case EnumMenuEditMode.CodeToButton
            Toolbar1.Buttons.Item(1).Enabled = False
        Case EnumMenuEditMode.ExchangeButton
            Toolbar1.Buttons.Item(2).Enabled = False
        Case EnumMenuEditMode.DeleteButton
            Toolbar1.Buttons.Item(3).Enabled = False
        Case EnumMenuEditMode.RenameButton
            Toolbar1.Buttons.Item(5).Enabled = False
        Case EnumMenuEditMode.PictureButton
            Toolbar1.Buttons.Item(7).Enabled = False
        Case EnumMenuEditMode.DeletePicture
            Toolbar1.Buttons.Item(8).Enabled = False
    End Select
    
    TempButton = 0
    
End Sub

Public Sub UpdateButtons()

    ValueBtnMenu2
    
''    On Error Resume Next
''    Dim Rst As New ADODB.Recordset
''
''    ReDim Parameter(3) As Parameter
''
''    Parameter(0) = GenerateInputParameter("@FactorType", adInteger, 4, EnumFactorType.Invoice)
''    Parameter(1) = GenerateInputParameter("@StationId", adInteger, 4, tmpStationNo)
''    Parameter(2) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
''    Parameter(3) = GenerateInputParameter("@BtnNum", adInteger, 4, 0)
''
''    Set Rst = RunParametricStoredProcedure2Rec("Get_Good_Menu", Parameter)
''
''    For i = 1 To BtnMenu.Count
''        If i < 21 Then
''            BtnMenu(i).BackColor = IIf(Invoice_BackColorBtn0 = 0, 12640511, Invoice_BackColorBtn0)
''        ElseIf i > 20 And i < 41 Then
''            BtnMenu(i).BackColor = IIf(Invoice_BackColorBtn1 = 0, 12640511, Invoice_BackColorBtn1)
''        ElseIf i > 40 And i < 61 Then
''            BtnMenu(i).BackColor = IIf(Invoice_BackColorBtn2 = 0, 12640511, Invoice_BackColorBtn2)
''        ElseIf i > 60 And i < 81 Then
''            BtnMenu(i).BackColor = IIf(Invoice_BackColorBtn3 = 0, 12640511, Invoice_BackColorBtn3)
''        ElseIf i > 80 And i < 101 Then
''            BtnMenu(i).BackColor = IIf(Invoice_BackColorBtn4 = 0, 12640511, Invoice_BackColorBtn4)
''        ElseIf i > 100 And i < 121 Then
''            BtnMenu(i).BackColor = IIf(Invoice_BackColorBtn5 = 0, 12640511, Invoice_BackColorBtn5)
''        ElseIf i > 120 And i < 141 Then
''            BtnMenu(i).BackColor = IIf(Invoice_BackColorBtn6 = 0, 12640511, Invoice_BackColorBtn6)
''        ElseIf i > 140 Then
''            BtnMenu(i).BackColor = IIf(Invoice_BackColorBtn7 = 0, 12640511, Invoice_BackColorBtn7)
''        End If
''        If BtnMenu(i).Enabled = True Then
''            BtnMenu(i).Caption = ""
''            BtnMenu(i).Picture = LoadPicture("")
''            BtnMenu(i).Tag = 0
''            Select Case clsStation.Language
''                Case EnumLanguage.Farsi
''                    BtnMenu(i).Font.Name = Invoice_FontMenuName
''                    BtnMenu(i).Font.Size = Val(Invoice_FontMenuSize)
''                    BtnMenu(i).Font.Bold = Invoice_FontMenuBold
''
''    '                FWRealButton1(i).Font.Name = Invoice_FontMenuName
''    '                FWRealButton1(i).Font.Size = Val(Invoice_FontMenuSize)
''    '                FWRealButton1(i).Font.Bold = Invoice_FontMenuBold
''
''                Case EnumLanguage.English
''                     BtnMenu(i).Font = "TimesNewRoman"
''                    BtnMenu(i).Font.Size = 10
''                    BtnMenu(i).Font.Bold = True
''
''    '                FWRealButton1(i).Font = "TimesNewRoman"
''    '                FWRealButton1(i).Font.Size = 10
''    '                FWRealButton1(i).Font.Bold = True
''
''           End Select
'''            BtnMenu(i).BackColor = &HFFC0C0       '&H137D7D
''        End If
''    Next i
''
''    If Rst.EOF = True And Rst.BOF = True Then
''
''    Else
''
''        While Rst.EOF <> True
''            BtnMenu(Rst.Fields("BtnNum").Value).Caption = Rst.Fields("NameDisp").Value
''            BtnMenu(Rst.Fields("BtnNum").Value).Tag = Rst.Fields("goodcode").Value
'''            BtnMenu(Rst.Fields("BtnNum").Value).BackColor = &HFF8080    '&HC0C0&
''
''            If Not IsNull(Rst.Fields("PicturePath").Value) And Trim(Rst.Fields("PicturePath").Value) <> "" Then
''              '  BtnMenu(Rst.Fields("BtnNum").Value).BackStyle = fmBackStyleOpaque
''                BtnMenu(Rst.Fields("BtnNum").Value).Picture = LoadPicture(App.Path & Rst.Fields("PicturePath").Value)
''              '  BtnMenu(Rst.Fields("BtnNum").Value).PicturePosition = fmPicturePositionAboveCenter
''            Else
''                BtnMenu(Rst.Fields("BtnNum").Value).Picture = LoadPicture("")
''            End If
''
''            Rst.MoveNext
''        Wend
''    End If
''
''    Set Rst = Nothing
    
End Sub

Private Sub BtnMenu_Click(index As Integer)

    Dim Rst As New ADODB.Recordset
    Dim s As String
    
    FillvsKeyBoardDefined index
    
    frmMsg.fwBtn(0).ButtonType = flwButtonOk
    frmMsg.fwBtn(0).Caption = "ﬁ»Ê·"
    frmMsg.fwBtn(0).Visible = True
    frmMsg.fwBtn(1).Visible = False
    
    Select Case MyFormAddEditMode
        
        Case EnumMenuEditMode.CodeToButton
            
                TempButton = index
                frmMsg.fwlblMsg.Caption = "·ÿ›« Ìò ò«·« «“ ·Ì”  ò«·«Â«Ì „⁄—›Ì ‰‘œÂ —ÊÌ „‰Ê «‰ Œ«» ‰„«ÌÌœ"
                frmMsg.fwBtn(1).ButtonType = flwButtonOk
                frmMsg.fwBtn(1).Caption = "ﬁ»Ê·"
                frmMsg.Show vbModal
        
        Case EnumMenuEditMode.ExchangeButton
        
            If TempButton = 0 Then 'first key
            
                TempButton = index
                
                frmMsg.fwlblMsg.Caption = "ﬂ·Ìœ «‰ Œ«»Ì »Â ﬂœ«„ ﬂ·Ìœ «‰ ﬁ«· Ì«»œ"
                frmMsg.fwBtn(1).ButtonType = flwButtonOk
                frmMsg.fwBtn(1).Caption = "ﬁ»Ê·"
                frmMsg.Show vbModal
                
            
            Else 'second key
                ReDim Parameter(4) As Parameter
                
                Parameter(0) = GenerateInputParameter("@FactorType", adInteger, 4, EnumFactorType.Invoice)
                Parameter(1) = GenerateInputParameter("@StationId", adInteger, 4, tmpStationNo)
                Parameter(2) = GenerateInputParameter("@BtnNum1", adInteger, 4, TempButton)
                Parameter(3) = GenerateInputParameter("@BtnNum2", adInteger, 4, index)
                Parameter(4) = GenerateOutputParameter("@Result", adInteger, 4)
                
                RunParametricStoredProcedure "ExchangeButtons", Parameter
                
                MyFormAddEditMode = EnumMenuEditMode.ViewButton
                UpdateToolbars
            
            End If
        
        
        Case EnumMenuEditMode.DeleteButton
        
            If BtnMenu(index).Caption = "" Then
                TempButton = 0
                frmMsg.fwlblMsg.Caption = "·ÿ›« Ìò ò·Ìœ ò«·« œ«— «‰ Œ«» ‰„«ÌÌœ"
                frmMsg.Show vbModal
            Else
            
                TempButton = index
                
                frmMsg.fwBtn(1).Caption = "ŒÌ—"
                frmMsg.fwBtn(0).Caption = "»·Â"
                frmMsg.fwBtn(1).Visible = True
                frmMsg.fwBtn(0).Visible = True
                
                frmMsg.fwlblMsg.Caption = "¬Ì« „Ì ŒÊ«ÂÌœ ﬂ· ﬂ«·«Â«Ì ﬂ·Ìœ —« Å«ﬂ ﬂ‰Ìœø "
                frmMsg.Show vbModal
                
                If modgl.mvarMsgIdx = vbYes Then
                
                    On Error GoTo RollBack
                    ReDim Parameter(2) As Parameter
                    Parameter(0) = GenerateInputParameter("@FactorType", adInteger, 4, EnumFactorType.Invoice)
                    Parameter(1) = GenerateInputParameter("@BtnNum", adInteger, 4, TempButton)
                    Parameter(2) = GenerateInputParameter("@StationId", adInteger, 4, tmpStationNo)
                    RunParametricStoredProcedure "Delete_tGood_Menu_By_BtnNum", Parameter
                    On Error GoTo 0
                    
                    frmMsg.fwlblMsg.Caption = " . ﬂ·ÌÂ ﬂ«·«Â« «“ ﬂ·Ìœ ›Êﬁ Å«ﬂ ‘œ "
                    frmMsg.fwBtn(0).Visible = False
                    frmMsg.fwBtn(1).ButtonType = flwButtonOk
                    frmMsg.fwBtn(1).Caption = "ﬁ»Ê·"
                    frmMsg.Show vbModal
                    
                    MyFormAddEditMode = ViewButton
                    UpdateToolbars
                    
                Else
                
                    frmMsg.fwlblMsg.Caption = "ﬂ«·«Ì „Ê—œ ‰Ÿ— —« «“ ÃœÊ· ﬂ«·«Â«Ì „⁄—›Ì ‘œÂ —ÊÌ ﬂ·Ìœ ﬂ«·« «‰ Œ«» ﬂ‰Ìœ  "
                    frmMsg.fwBtn(0).Visible = False
                    frmMsg.fwBtn(1).ButtonType = flwButtonOk
                    frmMsg.fwBtn(1).Caption = "ﬁ»Ê·"
                    frmMsg.Show vbModal
            
                End If
                
            End If
            
        Case EnumMenuEditMode.RenameButton
    
            If BtnMenu(index).Caption = "" Then
                TempButton = 0
                frmMsg.fwlblMsg.Caption = "·ÿ›« Ìò ò·Ìœ ò«·« œ«— «‰ Œ«» ‰„«ÌÌœ"
                frmMsg.Show vbModal
            Else
                frmInput.fwlblInput.Caption = " ‰«„ ÃœÌœ »—«Ì ‰„«Ì‘ ò«·« " & BtnMenu(index).Caption & " «‰ Œ«» ‰„«ÌÌœ "
                frmInput.Picture1.Visible = False
                frmInput.txtInput.Text = ""
                frmInput.MyForm = Me.Name
                frmInput.Show vbModal
'                mvarInput = frmInput.txtInput.Text
                mvarInput = Trim(mvarInput)
                If mvarInput <> "" And InStr(1, mvarInput, "'") = 0 Then
                    
                    ReDim Parameter(4) As Parameter
                    
                    Parameter(0) = GenerateInputParameter("@FactorType", adInteger, 4, EnumFactorType.Invoice)
                    Parameter(1) = GenerateInputParameter("@BtnNum", adInteger, 4, index)
                    Parameter(2) = GenerateInputParameter("@StationId", adInteger, 4, tmpStationNo)
                    Parameter(3) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
                    Parameter(4) = GenerateInputParameter("@Namedisp", adVarWChar, 50, mvarInput)
                    
                    RunParametricStoredProcedure2String "EditNameDisp", Parameter
                    
                    MyFormAddEditMode = ViewButton
                    UpdateToolbars
                Else
                    frmMsg.fwlblMsg.Caption = "·ÿ›« Ìò  ‰«„ „⁄ »— »—«Ì ò·Ìœ «‰ Œ«» ‰„«ÌÌœ"
                    frmMsg.Show vbModal
                End If
            End If
            
        Case EnumMenuEditMode.PictureButton
        
            If BtnMenu(index).Caption = "" Then
                frmMsg.fwlblMsg.Caption = "·ÿ›« Ìò ò·Ìœ ò«·« œ«— «‰ Œ«» ‰„«ÌÌœ"
                frmMsg.Show vbModal
            Else
                CommonDialog1.InitDir = App.Path & "\IMAGE\FOOD_PIC"
                CommonDialog1.Filter = "Pictures (*.bmp;*.ico;*.gif;*.jpg;*.jpeg)|*.bmp;*.ico;*.gif;*.jpg;*.jpeg"
                CommonDialog1.CancelError = True
                
                On Error GoTo RollBack
                CommonDialog1.ShowOpen
                On Error GoTo 0
                Dim fso As New FileSystemObject
                If fso.FileExists(CommonDialog1.InitDir & "\" & fso.GetBaseName(CommonDialog1.Filename) & "." & fso.GetExtensionName(CommonDialog1.Filename)) = True And LCase(CommonDialog1.Filename) <> LCase(CommonDialog1.InitDir & "\" & fso.GetBaseName(CommonDialog1.Filename) & "." & fso.GetExtensionName(CommonDialog1.Filename)) Then
                    Dim f As file
                    
                    Set f = fso.GetFile(CommonDialog1.InitDir & "\" & fso.GetBaseName(CommonDialog1.Filename) & "." & fso.GetExtensionName(CommonDialog1.Filename))
                    If Mid(ConvertToBin(f.Attributes, 8), 8, 1) = "1" Then
                        'If f.Attributes = ReadOnly Then
                        frmMsg.fwBtn(1).Caption = "ŒÌ—"
                        frmMsg.fwBtn(0).Caption = "»·Â"
                        frmMsg.fwBtn(1).Visible = True
                        frmMsg.fwBtn(0).Visible = True
                        frmMsg.fwlblMsg.Caption = "„ÊÃÊœ „Ì »«‘œ" & CommonDialog1.InitDir & "«Ì‰ ›«Ì· œ— " & vbCrLf & "¬Ì« „«Ì·Ìœ «“ ¬‰ «” ›«œÂ ‰„«ÌÌœø"
                        frmMsg.Show vbModal
                        f.Attributes = Normal
                        If mvarMsgIdx = vbYes Then
                    
                        Else
                            fso.CopyFile CommonDialog1.Filename, CommonDialog1.InitDir & "\" & fso.GetBaseName(CommonDialog1.Filename) & "." & fso.GetExtensionName(CommonDialog1.Filename), True
                        End If
                    End If
                ElseIf LCase(CommonDialog1.Filename) <> LCase(CommonDialog1.InitDir & "\" & fso.GetBaseName(CommonDialog1.Filename) & "." & fso.GetExtensionName(CommonDialog1.Filename)) Then
                    fso.CopyFile CommonDialog1.Filename, CommonDialog1.InitDir & "\" & fso.GetBaseName(CommonDialog1.Filename) & "." & fso.GetExtensionName(CommonDialog1.Filename), True
                End If
                ReDim Parameter(3) As Parameter
                
                Parameter(0) = GenerateInputParameter("@FactorType", adInteger, 4, EnumFactorType.Invoice)
                Parameter(1) = GenerateInputParameter("@BtnNum", adInteger, 4, index)
                Parameter(2) = GenerateInputParameter("@StationId", adInteger, 4, tmpStationNo)
                Parameter(3) = GenerateInputParameter("@PicturePath", adVarWChar, 50, "\IMAGE\FOOD_PIC\" & fso.GetBaseName(CommonDialog1.Filename) & "." & fso.GetExtensionName(CommonDialog1.Filename))
                
                RunParametricStoredProcedure2String "EditNameDispPicture", Parameter
                
                MyFormAddEditMode = ViewButton
                UpdateToolbars
                
            End If
            
        Case EnumMenuEditMode.DeletePicture
        
            If BtnMenu(index).Picture = 0 Then
                frmMsg.fwlblMsg.Caption = "·ÿ›« Ìò ò·Ìœ ⁄ò” œ«— «‰ Œ«» ‰„«ÌÌœ"
                frmMsg.Show vbModal
            Else
                ReDim Parameter(3) As Parameter
                
                Parameter(0) = GenerateInputParameter("@FactorType", adInteger, 4, EnumFactorType.Invoice)
                Parameter(1) = GenerateInputParameter("@BtnNum", adInteger, 4, index)
                Parameter(2) = GenerateInputParameter("@StationId", adInteger, 4, tmpStationNo)
                Parameter(3) = GenerateInputParameter("@PicturePath", adVarWChar, 50, "")
                
                RunParametricStoredProcedure2String "EditNameDispPicture", Parameter
                
                MyFormAddEditMode = ViewButton
                UpdateToolbars
            End If
        
    End Select
    
    On Error Resume Next
    Unload frmMsg
    On Error GoTo 0
    Exit Sub
    
RollBack:
    Select Case err.Number
        Case 32755
            MyFormAddEditMode = ViewButton
            UpdateToolbars
        Case Else
            MyFormAddEditMode = ViewButton
            UpdateToolbars
            MsgBox "„ «”›«‰Â  €ÌÌ—«  ﬁ«»· «⁄„«· ‰Ì” "
    
    End Select
End Sub


Private Sub cboStations_Click()
    tmpStationNo = cboStations.ItemData(cboStations.ListIndex)
    MyFormAddEditMode = EnumMenuEditMode.ViewButton
    UpdateToolbars


    Dim Rst As New ADODB.Recordset
    ReDim Parameter(1) As Parameter
    
    
    If cboStations.ListIndex > -1 Then
        tmpStationNo = cboStations.ItemData(cboStations.ListIndex)
    Else
        tmpStationNo = 0
    End If
    
    MenuBarDescription
    ValueBtnMenu2
    SetBtnMenuPosition
        
'    Parameter(0) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
'    Parameter(1) = GenerateInputParameter("@StationId", adInteger, 4, tmpStationNo)
'
'    If Rst.State <> 0 Then Rst.Close
'    Set Rst = RunParametricStoredProcedure2Rec("Get_tPocketPCGroup", Parameter)
'    Dim ii As Integer
'    ii = 0
'    While Rst.EOF <> True
'        Select Case clsStation.Language
'            Case Farsi
'                lstGroups.AddItem Rst.Fields("Description").Value
'            Case English
'                lstGroups.AddItem Rst.Fields("LatinDescription").Value
'        End Select
'
'        lstGroups.ItemData(lstGroups.ListCount - 1) = Rst.Fields("PocketPCGroupCode").Value
'        If IsNull(Rst.Fields("StationId").Value) <> True Then
'
'            lstGroups.Selected(lstGroups.ListCount - 1) = True
'            If ii <= SSTab1.Tabs - 1 Then
'                SSTab1.TabCaption(ii) = Rst.Fields("Description").Value
'                ii = ii + 1
'            End If
'        End If
'        Rst.MoveNext
'    Wend
'
'    Set Rst = Nothing

End Sub
Private Sub MenuBarDefine()
    If clsStation.TouchScreen = True Then
        MaxBtnMenu = 160
        BtnMenuPerFrame = 20
    Else
        MaxBtnMenu = 320
        BtnMenuPerFrame = 40
    End If
    
    For i = 2 To MaxBtnMenu
        Load BtnMenu(i)
    Next
    
End Sub


Private Sub cmdCopyMenu_Click()
    If tmpStationNo = FWStationNo.Value Then ShowDisMessage "„‰ÊÂ«Ì „»œ« Ê „ﬁ’œ ÌòÌ Â” ‰œ", 1500: Exit Sub
    ShowMessage "¬Ì« »—«Ì òÅÌ „‰ÊÂ« »Â «Ì” ê«Â œÌê— „ÿ„∆‰ Â” Ìœø „‰ÊÂ«Ì „ﬁ’œ Õ–› Ê „‰ÊÂ«Ì ÃœÌœ Ã«Ìê“Ì‰ ¬‰Â« ŒÊ«Â‰œ ‘œ", True, True, "»·Ì", "ŒÌ—"
    Dim Result As Long
    If modgl.mvarMsgIdx = vbYes Then
        
        ReDim Parameter(2) As Parameter
        Parameter(0) = GenerateInputParameter("@StationId", adInteger, 4, tmpStationNo)
        Parameter(1) = GenerateInputParameter("@NewStationId", adInteger, 4, FWStationNo.Value)
        Parameter(2) = GenerateOutputParameter("@intStatus", adInteger, 4)
        
        Result = RunParametricStoredProcedure("Copy_tGood_Menu", Parameter)
        If Result > 0 Then
            ShowDisMessage " €ÌÌ—«  „‰ÊÂ« «‰Ã«„ ‘œ", 1500
        Else
            ShowDisMessage " œ— À»   €ÌÌ—«  „‰ÊÂ« „‘ò· ÊÃÊœ œ«—œ ", 1500
        End If
    End If

End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
            
    Dim ii As Integer
    For ii = 0 To SSTab1.Tabs - 1
        SSTab1.TabPicture(ii) = LoadPicture("")
    Next
    ValueBtnMenu2
    SetBtnMenuPosition
    If SSTab1.TabsPerRow < 5 Then SSTab1.TabPicture(SSTab1.Tab) = ImageList1.ListImages(4).Picture

End Sub

Public Sub MenuBarDescription()
Dim ii As Long
On Error Resume Next
    ReDim Parameter(1) As Parameter
    Parameter(0) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
    Parameter(1) = GenerateInputParameter("@StationId", adInteger, 4, tmpStationNo)
    
    If rctmp.State <> 0 Then rctmp.Close
    Set rctmp = RunParametricStoredProcedure2Rec("Get_tPocketPCGroup", Parameter)
    ii = 0
    SSTab1.Tabs = 1
    SSTab1.TabCaption(0) = "ê—ÊÂ 1"
    While rctmp.EOF <> True
        If IsNull(rctmp.Fields("StationId").Value) <> True Then
            If rctmp.Fields("PocketPCGroupCode").Value <= 8 Then
                SSTab1.Tabs = ii + 1
                SSTab1.TabCaption(ii) = rctmp.Fields("Description").Value
                ii = ii + 1
            End If
        End If
        rctmp.MoveNext
    Wend
    Dim jj As Long
    'If SSTab1.Tabs > 10 Then SSTab1.TabsPerRow = CInt(SSTab1.Tabs / 2) Else SSTab1.TabsPerRow = SSTab1.Tabs
    If clsStation.NoRowMenu = 2 Then SSTab1.TabsPerRow = CLng((SSTab1.Tabs + 0.5) / 2) Else SSTab1.TabsPerRow = SSTab1.Tabs
    If SSTab1.TabsPerRow > SSTab1.Tabs Then SSTab1.TabsPerRow = SSTab1.Tabs
    If SSTab1.TabsPerRow > 5 Then SSTab1.Font.Size = SSTab1.Font.Size - 2: SSTab1.TabHeight = 600 Else SSTab1.TabHeight = 500
    SSTab1.Height = SSTab1.TabHeight + 50
    Set rctmp = Nothing
    
End Sub

Public Sub SetBtnMenuPosition()
'txtTxtWidth = 0
'TxtHeight = 0
    Dim xHeight As Double
    Dim xWidth As Double
    If clsStation.TouchScreen = True Then
        xHeight = frameMenu.Height - SSTab1.Height - 50
        If xHeight < 100 Then xHeight = 100
        BtnMenu(1).Height = xHeight / 5
    Else
        xHeight = frameMenu.Height - SSTab1.Height - 100
        If xHeight < 100 Then xHeight = 100
        BtnMenu(1).Height = xHeight / 10
    End If
    xWidth = (frameMenu.Width - 250)
    If xWidth < 100 Then xWidth = 100
    BtnMenu(1).Width = xWidth / 4
    SSTab1.Width = xWidth + 50
    SSTab1.Left = 100
    lastPosition.X = 120
    lastPosition.Y = SSTab1.Height + 50
    
    IndexMenuTab = SSTab1.Tab  '(RowTab * 4) + Column - 2 ' Index of group menu
    Dim i As Long
    For i = (IndexMenuTab * BtnMenuPerFrame) + 1 To (IndexMenuTab * BtnMenuPerFrame) + BtnMenuPerFrame
'        Debug.Print (i - 1) Mod 4
'        If (i - 1) Mod 4 = 0 And i > 1 Then
        If lastPosition.X > frameMenu.Width - (Val(BtnMenu(1).Width)) Then
            lastPosition.X = 120
            lastPosition.Y = lastPosition.Y + Val(BtnMenu(1).Height)
        End If
        BtnMenu(i).Width = BtnMenu(1).Width
        BtnMenu(i).Height = BtnMenu(1).Height
        BtnMenu(i).Left = lastPosition.X
        BtnMenu(i).Top = lastPosition.Y
        lastPosition.X = lastPosition.X + Val(BtnMenu(1).Width)
    Next
End Sub

Public Sub ValueBtnMenu2()
    Dim i As Long
    For i = 1 To MaxBtnMenu
        BtnMenu(i).Visible = False
        BtnMenu(i).Enabled = True       ' all key enabled
        BtnMenu(i).Tag = ""
        BtnMenu(i).Caption = ""
        BtnMenu(i).Picture = LoadPicture("")
    Next
    IndexMenuTab = SSTab1.Tab   '(RowTab * 4) + Column - 2   ' Index of group menu
    For i = (IndexMenuTab * BtnMenuPerFrame) + 1 To (IndexMenuTab * BtnMenuPerFrame) + BtnMenuPerFrame
        BtnMenu(i).Visible = True
        Select Case IndexMenuTab
            Case 0
                BtnMenu(i).BackColor = IIf(Invoice_BackColorBtn0 = 0, 12640511, Invoice_BackColorBtn0)
            Case 1
                BtnMenu(i).BackColor = IIf(Invoice_BackColorBtn1 = 0, 12640511, Invoice_BackColorBtn1)
            Case 2
                BtnMenu(i).BackColor = IIf(Invoice_BackColorBtn2 = 0, 12640511, Invoice_BackColorBtn2)
            Case 3
                BtnMenu(i).BackColor = IIf(Invoice_BackColorBtn3 = 0, 12640511, Invoice_BackColorBtn3)
            Case 4
                BtnMenu(i).BackColor = IIf(Invoice_BackColorBtn4 = 0, 12640511, Invoice_BackColorBtn4)
            Case 5
                BtnMenu(i).BackColor = IIf(Invoice_BackColorBtn5 = 0, 12640511, Invoice_BackColorBtn5)
            Case 6
                BtnMenu(i).BackColor = IIf(Invoice_BackColorBtn6 = 0, 12640511, Invoice_BackColorBtn6)
            Case 7
                BtnMenu(i).BackColor = IIf(Invoice_BackColorBtn7 = 0, 12640511, Invoice_BackColorBtn7)
        End Select

'        Select Case clsStation.Language
'            Case EnumLanguage.Farsi
'                BtnMenu(i).Font.Name = Invoice_FontMenuName
'                BtnMenu(i).Font.Size = Val(Invoice_FontMenuSize)
'                BtnMenu(i).Font.Bold = Invoice_FontMenuBold
'            Case EnumLanguage.English
'                BtnMenu(i).Font = "TimesNewRoman"
'                BtnMenu(i).Font.Size = 10
'                BtnMenu(i).Font.Bold = True
'       End Select
    Next
    
    ReDim Parameter(1) As Parameter
    Parameter(0) = GenerateInputParameter("@StationId", adInteger, 4, tmpStationNo)
    Parameter(1) = GenerateInputParameter("@FactorType", adInteger, 4, EnumFactorType.Invoice)
    
    Set rctmp = RunParametricStoredProcedure2Rec("GetPictureButton", Parameter)
    
    Do While Not rctmp.EOF
        If Not IsNull(rctmp.Fields("PicturePath")) And rctmp.Fields("BtnNum") <= BtnMenu.Count Then
           ' BtnMenu(rctmp.Fields("BtnNum")).BackStyle = fmBackStyleOpaque
            If rctmp.Fields("PicturePath") <> "" And rctmp.Fields("BtnNum") >= (IndexMenuTab * BtnMenuPerFrame) + 1 And rctmp.Fields("BtnNum") <= (IndexMenuTab * BtnMenuPerFrame) + BtnMenuPerFrame Then
                BtnMenu(rctmp.Fields("BtnNum")).Picture = LoadPicture(App.Path & rctmp.Fields("PicturePath"))
                BtnMenu(rctmp.Fields("BtnNum")).Caption = rctmp.Fields("NameDisp")
                
            End If
           ' BtnMenu(rctmp.Fields("BtnNum")).PicturePosition = fmPicturePositionAboveCenter
           ' BtnMenu(rctmp.Fields("BtnNum")).WordWrap = False  ' Single Line If Has Picture
        
        End If
       ' BtnMenu(rctmp.Fields("BtnNum")).WordWrap = True  ' Double Line If No Picture
        rctmp.MoveNext
    Loop
    rctmp.Cancel
    
    ReDim Parameter(2) As Parameter
    
    Parameter(0) = GenerateInputParameter("@StationId", adInteger, 4, tmpStationNo)
    Parameter(1) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
    Parameter(2) = GenerateInputParameter("@FactorType", adInteger, 4, EnumFactorType.Invoice)
    
    Set rctmp = RunParametricStoredProcedure2Rec("GetButtonMenu", Parameter)
    
    Do While Not rctmp.EOF
        If Not IsNull(rctmp.Fields("BtnNum")) And rctmp.Fields("BtnNum") <= BtnMenu.Count And rctmp.Fields("BtnNum") >= (IndexMenuTab * BtnMenuPerFrame) + 1 And rctmp.Fields("BtnNum") <= (IndexMenuTab * BtnMenuPerFrame) + BtnMenuPerFrame Then
            BtnMenu(rctmp.Fields("BtnNum")).Tag = BtnMenu(rctmp.Fields("BtnNum")).Tag & rctmp.Fields("Code") & ";"
'            FWRealButton1(rctmp.Fields("BtnNum")).Tag = BtnMenu(rctmp.Fields("BtnNum")).Tag & rctmp.Fields("Code") & ";"
            BtnMenu(rctmp.Fields("BtnNum")).Visible = True
            BtnMenu(rctmp.Fields("BtnNum")).Enabled = True
            If Not IsNull(rctmp.Fields("NameDisp")) Then
                BtnMenu(rctmp.Fields("BtnNum")).Caption = rctmp.Fields("NameDisp")
'                FWRealButton1(rctmp.Fields("BtnNum")).Caption = rctmp.Fields("NameDisp")
            Else
                BtnMenu(rctmp.Fields("BtnNum")).Caption = rctmp.Fields("Name")
'                FWRealButton1(rctmp.Fields("BtnNum")).Caption = rctmp.Fields("Name")
            End If
        End If
        rctmp.MoveNext
    Loop
    rctmp.Cancel
'    For i = (IndexMenuTab * BtnMenuPerFrame) + 1 To (IndexMenuTab * BtnMenuPerFrame) + BtnMenuPerFrame
'        If i <= BtnMenu.Count Then
'           If Len(BtnMenu(i).Tag) > 0 Then
'               BtnMenu(i).Tag = Left(BtnMenu(i).Tag, Len(BtnMenu(i).Tag) - 1)
'        '        FWRealButton1(i).Tag = Left(FWRealButton1(i).Tag, Len(FWRealButton1(i).Tag) - 1)
'               If BtnMenu(i).Tag = "" And BtnMenu(i).Caption = "" Then
'                   BtnMenu(i).Enabled = False
'        '            FWRealButton1(i).Enabled = False
'
'               End If
'           Else
'               BtnMenu(i).Enabled = False
'           End If
'        End If
'    Next i

Exit Sub
Err1:
Resume Next
End Sub

Private Sub CmdDoneGroup_Click()
    
    Dim i As Integer
    
    If lstGroups.SelCount = 0 Then Exit Sub
    
    Dim SelectedGroups As String
    ReDim Parameter(1) As Parameter
    For i = 0 To lstGroups.ListCount - 1
        If lstGroups.Selected(i) = True Then
            SelectedGroups = SelectedGroups & lstGroups.ItemData(i) & ","
        End If
    Next i
    SelectedGroups = Left(SelectedGroups, Len(SelectedGroups) - 1)
    If cboStations.ListIndex > -1 Then
        Parameter(0) = GenerateInputParameter("@StationId", adInteger, 4, cboStations.ItemData(cboStations.ListIndex))
        Parameter(1) = GenerateInputParameter("@PocketPCGroupCode", adVarWChar, 4000, SelectedGroups)
        
        RunParametricStoredProcedure "Update_tPocketPC_StationGroups", Parameter
        ShowDisMessage "«Œ ’«’ ê—ÊÂÂ« »Â «Ì” ê«Â «‰Ã«„ ‘œ", 1500
        cboStations_Click
    End If
End Sub

Private Sub CmdDone_Click()
    
    Dim strTemp As String
    With vsNoKeyBoardDefined

        Select Case MyFormAddEditMode
            Case CodeToButton
                If TempButton <> 0 And .Rows > 1 Then
                    
                    For i = 1 To .Rows - 1
                        If Val(.TextMatrix(i, 4)) = -1 Then
                            strTemp = strTemp & .TextMatrix(i, 1) & ","
                            
                        End If
                    Next i
                    
                    If strTemp <> "" Then
                        strTemp = Left(strTemp, Len(strTemp) - 1)
                        On Error GoTo RollBack
                        ReDim Parameter(3) As Parameter
                        Parameter(0) = GenerateInputParameter("@SelectedGoodCode", adVarWChar, 4000, strTemp)
                        Parameter(1) = GenerateInputParameter("@StationId", adInteger, 4, tmpStationNo)
                        Parameter(2) = GenerateInputParameter("@BTNNUM", adInteger, 4, TempButton)
                        Parameter(3) = GenerateInputParameter("@FactorType", adInteger, 4, EnumFactorType.Invoice)
                        RunParametricStoredProcedure "InsertGoodMenu", Parameter
                        On Error GoTo 0
                    
                        TempButton = 0
                        MyFormAddEditMode = ViewButton
                        UpdateToolbars
                        
                    End If
                End If
        End Select
    End With

    Exit Sub
    
RollBack:

    MyFormAddEditMode = ViewButton
    UpdateToolbars
    MsgBox "„ «”›«‰Â  €ÌÌ—«  ﬁ«»· «⁄„«· ‰Ì” "
    
End Sub


Private Sub Command1_Click()
    If FrameGroup.Visible = False Then Cancel
    FrameGroup.Visible = True
End Sub

Private Sub Command2_Click()
    FrameGroup.Visible = False
End Sub
Public Sub Update()
    
    On Error GoTo ErrHandler
    Dim intResult As Integer
    
    Select Case MyFormAddEditMode2
    
        Case AddMode
        
            If txtGroupName.Text = "" Or InStr(txtGroupName.Text, "'") <> 0 Then
                Exit Sub
            End If
            
            ReDim Parameter(1) As Parameter
            
            Parameter(0) = GenerateInputParameter("@Description", adVarWChar, 50, txtGroupName.Text)
            Parameter(1) = GenerateOutputParameter("@Result", adInteger, 4)
            intResult = RunParametricStoredProcedure("Insert_PocketPCGroup", Parameter)
            If intResult <> -1 Then
            
            Else
            
            End If
            
        Case EditMode
        
            If txtGroupName.Text = "" Or InStr(txtGroupName.Text, "'") <> 0 Then
                Exit Sub
            End If
            Dim Parameter2(3) As Parameter
            Parameter2(0) = GenerateInputParameter("@PocketPCGroup", adInteger, 4, lstGroups.ItemData(lstGroups.ListIndex))
            Parameter2(1) = GenerateInputParameter("@Description", adVarWChar, 50, txtGroupName.Text)
            Parameter2(2) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
            Parameter2(3) = GenerateOutputParameter("@Result", adInteger, 4)
            intResult = RunParametricStoredProcedure("Update_PocketPCGroup", Parameter2)
            If intResult <> -1 Then
            
            Else
            
            End If
    End Select
    cboStations_Click
    DefaultSetting
'    MyFormAddEditMode2 = ViewMode
'    SetFirstToolBar
''    HeaderLabel CInt(MyFormAddEditMode), Me.fwlblMode
Exit Sub
ErrHandler:
    ShowDisMessage err.Description, 3000
End Sub
Private Sub DefaultSetting()

    lstGroups.Clear
    txtGroupName.Text = ""
    txtGroupName.Locked = True
    
    Dim Rst As New ADODB.Recordset
    
    If cboStations.ListIndex > -1 Then
        tmpStationNo = cboStations.ItemData(cboStations.ListIndex)
    Else
        tmpStationNo = 0
    End If
    
    ReDim Parameter(1) As Parameter
    
    Parameter(0) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
    Parameter(1) = GenerateInputParameter("@StationId", adInteger, 4, tmpStationNo)
    
    If Rst.State <> 0 Then Rst.Close
    Set Rst = RunParametricStoredProcedure2Rec("Get_tPocketPCGroup", Parameter)
    
    If Not (Rst.EOF = True And Rst.BOF = True) Then
        While Rst.EOF <> True
            If Rst.Fields("PocketPCGroupCode").Value <= 8 Then
                Select Case clsStation.Language
                    Case Farsi
                        lstGroups.AddItem Rst.Fields("Description").Value
                    Case English
                        lstGroups.AddItem Rst.Fields("LatinDescription").Value
                End Select
                lstGroups.ItemData(lstGroups.ListCount - 1) = Rst.Fields("PocketPCGroupCode").Value
                If IsNull(Rst.Fields("StationId").Value) <> True Then
                            lstGroups.Selected(lstGroups.ListCount - 1) = True
        
                End If
            End If
            Rst.MoveNext
        Wend
    End If
    Set Rst = Nothing
    
End Sub
Private Sub SetFirstToolBar()

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
    
    If MyFormAddEditMode2 = ViewMode Then  ' View Mode
    
        mdifrm.Toolbar1.Buttons(6).Enabled = True  'Add
        mdifrm.Toolbar1.Buttons(7).Enabled = True  'Edit
        mdifrm.Toolbar1.Buttons(8).Enabled = False  'Enter
        mdifrm.Toolbar1.Buttons(9).Enabled = False   'Esc
        mdifrm.Toolbar1.Buttons(10).Enabled = True 'Delete
        txtGroupName.Locked = True
        
    ElseIf MyFormAddEditMode2 = AddMode Then    'Add Mode
    
        mdifrm.Toolbar1.Buttons(6).Enabled = False  'Add
        mdifrm.Toolbar1.Buttons(7).Enabled = False  'Edit
        mdifrm.Toolbar1.Buttons(8).Enabled = True   'Enter
        mdifrm.Toolbar1.Buttons(9).Enabled = True   'Esc
        mdifrm.Toolbar1.Buttons(10).Enabled = False 'Delete
        txtGroupName.Locked = False
        
    ElseIf MyFormAddEditMode2 = EditMode Then     'Edit
        
        mdifrm.Toolbar1.Buttons(6).Enabled = False  'Add
        mdifrm.Toolbar1.Buttons(7).Enabled = False  'Edit
        mdifrm.Toolbar1.Buttons(8).Enabled = True   'Enter
        mdifrm.Toolbar1.Buttons(9).Enabled = True   'Esc
        mdifrm.Toolbar1.Buttons(10).Enabled = False 'Delete
        txtGroupName.Locked = False
    
    End If
    
'    HeaderLabel Val(MyFormAddEditMode), fwlblMode


End Sub
Public Sub Edit()
    
    MyFormAddEditMode2 = EditMode
    SetFirstToolBar

End Sub

Public Sub Delete()
    
    On Error GoTo ErrHandler
   
    If lstGroups.SelCount = 0 Then Exit Sub
    Dim a, b
    a = lstGroups.List(lstGroups.ListIndex)
    b = lstGroups.ListIndex
    If b = "" Then Exit Sub
    Dim SelectedGroups As String
    ReDim Parameter(1) As Parameter
        If lstGroups.Selected(b) = True Then
            ShowDisMessage "«» œ« «Œ ’«’ ‰«„ ê—ÊÂ »Â „‰Ê Â« —« »—œ«—Ìœ", 1500
            Exit Sub
    End If
        SelectedGroups = lstGroups.ItemData(b)
'    If SelectedGroups = "" Then Exit Sub
    
    frmMsg.fwlblMsg.Caption = "¬Ì« „ÿ„∆‰Ìœ „Ì ŒÊ«ÂÌœ ê—ÊÂ " & "'" & lstGroups.List(lstGroups.ListIndex) & "'" & " —« Õ–› ﬂ‰Ìœø"
    frmMsg.fwBtn(0).Caption = "»·Ì"
    frmMsg.fwBtn(1).Caption = "ŒÌ—"
    
    frmMsg.Show vbModal
    
    If modgl.mvarMsgIdx = vbNo Then
        Exit Sub
    End If
    ReDim Parameter(2) As Parameter
    Parameter(0) = GenerateInputParameter("@Code", adInteger, 4, SelectedGroups)
    Parameter(1) = GenerateInputParameter("@Branch", adInteger, 4, CurrentBranch)
    Parameter(2) = GenerateOutputParameter("@Result", adInteger, 4)
    Dim Result As Integer
    Result = RunParametricStoredProcedure("Delete_tPocketPCGroup", Parameter)
    
    If Result = 0 Then
    
        frmMsg.fwlblMsg.Caption = "„‘ò·Ì œ—Õ–› «Ì‰ ê—ÊÂ ÊÃÊœ œ«—œ «»‰œ« «Œ ’«’ ‰«„ ê—ÊÂÂ« »Â „‰ÊÂ« —« »—œ«—Ìœ"
        frmMsg.fwBtn(0).Caption = "ﬁ»Ê·"
        frmMsg.fwBtn(1).Visible = False
        frmMsg.fwBtn(0).ButtonType = flwButtonOk
        frmMsg.Show vbModal
        Exit Sub
    Else
    
        frmMsg.fwlblMsg.Caption = "‘„« Ìò ê—ÊÂ —« Õ–› ò—œÌœ"
        frmMsg.fwBtn(0).Caption = "ﬁ»Ê·"
        frmMsg.fwBtn(1).Visible = False
        frmMsg.fwBtn(0).ButtonType = flwButtonOk
        frmMsg.Show vbModal
        
    End If
    
    cboStations_Click
'    MyFormAddEditMode2 = AddMode
'    SetFirstToolBar
Exit Sub
ErrHandler:
    ShowDisMessage err.Description, 3000
    
End Sub

Private Sub lstGroups_Click()

    If MyFormAddEditMode2 = EditMode Then
    
        txtGroupName.Text = lstGroups.List(lstGroups.ListIndex)
    End If

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call PresetScreenSaver
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call PresetScreenSaver
End Sub
Private Sub Form_Activate()

    VarActForm = Me.Name
    AllButton vbOff, True
    
    mdifrm.Toolbar1.Buttons(20).Enabled = False
    mdifrm.Toolbar1.Buttons(21).Enabled = False
    mdifrm.Toolbar1.Buttons(23).Enabled = True
    mdifrm.Toolbar1.Buttons(24).Enabled = True
    mdifrm.Toolbar1.Buttons(25).Enabled = True
    mdifrm.Toolbar1.Buttons(26).Enabled = True
    mdifrm.Toolbar1.Buttons(27).Enabled = True
    mdifrm.Toolbar1.Buttons(9).Enabled = True
     
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

    If ClsFormAccess.frmMenu = False Then
        Unload Me
        Exit Sub
    End If

    VarActForm = Me.Name
    CenterTop Me
    
    NotSuportedGoodType = forBuy
            
    Dim Rst As New ADODB.Recordset
    Set Rst = RunStoredProcedure2RecordSet("Get_Pc_Stations")

    i = 0
    cboStations.Clear
    If Not (Rst.EOF = True And Rst.BOF = True) Then
        While Not Rst.EOF
            i = i + 1
            cboStations.AddItem Rst.Fields("Description").Value
            cboStations.ItemData(cboStations.ListCount - 1) = Rst.Fields("StationID").Value
            Rst.MoveNext
        Wend
    End If
    If Rst.State <> 0 Then Rst.Close
    If i > clsArya.MaxStationNo And DebugMode = False And HardLockFlagTrial = False Then
       MsgBox "Œÿ« œ—  ⁄œ«œ «Ì” ê«ÂÂ«Ì Pc"
       End
    End If
    
    FWStationNo.Max = clsArya.MaxStationNo
    MenuBarDefine
    
    If cboStations.ListCount > 0 Then
        For i = 0 To cboStations.ListCount - 1
            If clsArya.StationNo = cboStations.ItemData(i) Then
                cboStations.ListIndex = i
                Exit For
            End If
        Next
    Else
        Unload Me
        Exit Sub
    End If
    Set Rst = Nothing
    
            
            
    With vsKeyBoardDefined
        
        .Rows = 1
        .Cols = 8
        .TextMatrix(0, 0) = "—œÌ›"
        .TextMatrix(0, 1) = "òœ ò«·«"
        .TextMatrix(0, 2) = "òœ ”ÿÕ «Ê· ò«·«"
        .TextMatrix(0, 3) = "òœ ”ÿÕ œÊ„ ò«·«"
        .TextMatrix(0, 4) = "«‰ Œ«»"
        .TextMatrix(0, 5) = "ê—ÊÂ «’·Ì"
        .TextMatrix(0, 6) = "“Ì— ê—ÊÂ"
        .TextMatrix(0, 7) = "‰«„ ò«·«"
        
        .ColDataType(4) = flexDTBoolean
      '  .ColHidden(1) = True
        .ColHidden(2) = True
        .ColHidden(3) = True
'        .ColHidden(4) = True
        .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter
        .ColAlignment(5) = flexAlignRightCenter
        .ColAlignment(6) = flexAlignRightCenter
        .ColAlignment(7) = flexAlignRightCenter
        .ColAlignment(0) = flexAlignCenterCenter
       
        .AutoSearch = flexSearchFromCursor
    End With
    
    
    With vsNoKeyBoardDefined
        
        .Rows = 1
        .Cols = 8
        .TextMatrix(0, 0) = "—œÌ›"
        .TextMatrix(0, 1) = "òœ ò«·«"
        .TextMatrix(0, 2) = "òœ ”ÿÕ «Ê· ò«·«"
        .TextMatrix(0, 3) = "òœ ”ÿÕ œÊ„ ò«·«"
        .TextMatrix(0, 4) = "«‰ Œ«»"
        .TextMatrix(0, 5) = "ê—ÊÂ «’·Ì"
        .TextMatrix(0, 6) = "“Ì— ê—ÊÂ"
        .TextMatrix(0, 7) = "‰«„ ò«·«"
        
        .ColDataType(4) = flexDTBoolean
      '  .ColHidden(1) = True
        .ColHidden(2) = True
        .ColHidden(3) = True
        .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter
        .ColAlignment(5) = flexAlignRightCenter
        .ColAlignment(6) = flexAlignRightCenter
        .ColAlignment(7) = flexAlignRightCenter
        .ColAlignment(0) = flexAlignCenterCenter
       
        .AutoSearch = flexSearchFromCursor
    End With

    
    MyFormAddEditMode = ViewButton
    UpdateToolbars
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

    Call SetColor
    
    MyFormAddEditMode = ViewButton
    DefaultSetting
    SetFirstToolBar
    Add
    End Sub

Private Sub Form_Unload(Cancel As Integer)
    If PosConnection.State = adStateOpen Then PosConnection.Close: Set PosConnection = Nothing
    AllButton vbOff, True

    VarActForm = ""
    SaveSetting strMainKey, Me.Name, "Left", Me.Left
    SaveSetting strMainKey, Me.Name, "Top", Me.Top

    
End Sub


Private Sub FWBOK_Click()
    Dim varAnswer As Integer
    Load frmMsg
    frmMsg.fwlblMsg.Caption = " ¬Ì«  €ÌÌ—«  –ŒÌ—Â ‘Êœ ø"
    frmMsg.Show vbModal
    varAnswer = modgl.mvarMsgIdx
    If varAnswer = vbYes Then
    
        Call SetUserSettingFile(cmdPatern.BackColor, Val(2 & IndexMenuTab))
        ShowDisMessage ".  ‰ŸÌ„«  «‰Ã«„ ‘œ", 1000
        Call SetColor
        UpdateButtons
    End If
  
Exit Sub

Err1:
Resume Next

End Sub

Private Sub optColor_Click(index As Integer)
    cmdPatern.BackColor = OptColor(index).BackColor

'    publngBackColorKlydKala = optColor(Index).BackColor

End Sub

Private Sub ResizeKit1_ExitResize(ByVal XScale As Double, ByVal YScale As Double)


    If formloadFlag = True Then
        SaveSetting strMainKey, Me.Name, "Height", Me.Height
        SaveSetting strMainKey, Me.Name, "Width", Me.Width
    
    
        SetBtnMenuPosition
    
    End If

End Sub


Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    
    For i = 1 To Toolbar1.Buttons.Count
        If i <> Button.index Then
            Toolbar1.Buttons.Item(i).Enabled = True
        Else
            Toolbar1.Buttons.Item(i).Enabled = False
        End If
    Next i
    
    TempButton = 0
    
    frmMsg.fwBtn(0).ButtonType = flwButtonOk
    frmMsg.fwBtn(0).Caption = "ﬁ»Ê·"
    frmMsg.fwBtn(0).Visible = True
    frmMsg.fwBtn(1).Visible = False
    
    Select Case Button.index
    
        Case 1
            MyFormAddEditMode = CodeToButton
            frmMsg.fwlblMsg.Caption = "·ÿ›« Ìò ò·Ìœ «‰ Œ«» ‰„«ÌÌœ"
            frmMsg.Show vbModal
            
        Case 2
            MyFormAddEditMode = ExchangeButton
            frmMsg.fwlblMsg.Caption = "·ÿ›« ò·Ìœ «Ê· —« «‰ Œ«» ‰„«ÌÌœ"
            frmMsg.Show vbModal
            TempButton = 0
            
        Case 3
            MyFormAddEditMode = DeleteButton
            frmMsg.fwlblMsg.Caption = " ·ÿ›« ﬂ·Ìœ Õ–›Ì —«  ⁄ÌÌ‰  ‰„«ÌÌœ"
            frmMsg.Show vbModal
        Case 5
            MyFormAddEditMode = RenameButton
            frmMsg.fwlblMsg.Caption = "·ÿ›« Ìò ò·Ìœ ò«·« œ«— «‰ Œ«» ‰„«ÌÌœ"
            frmMsg.Show vbModal
        Case 7
            MyFormAddEditMode = PictureButton
            frmMsg.fwlblMsg.Caption = "·ÿ›« Ìò ò·Ìœ ò«·« œ«— «‰ Œ«» ‰„«ÌÌœ"
            frmMsg.Show vbModal
        Case 8
            MyFormAddEditMode = DeletePicture
            frmMsg.fwlblMsg.Caption = "·ÿ›« Ìò ò·Ìœ ⁄ò” œ«— «‰ Œ«» ‰„«ÌÌœ"
            frmMsg.Show vbModal
    End Select
    
    Unload frmMsg

End Sub


Private Sub vsKeyBoardDefined_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    With vsKeyBoardDefined
        Select Case MyFormAddEditMode
            Case DeleteButton
                If Val(.TextMatrix(Row, Col)) = -1 And TempButton <> 0 Then
                    
                    On Error GoTo RollBack
                    ReDim Parameter(3) As Parameter
                    Parameter(0) = GenerateInputParameter("@FactorType", adInteger, 4, EnumFactorType.Invoice)
                    Parameter(1) = GenerateInputParameter("@GoodCode", adInteger, 4, .TextMatrix(Row, 1))
                    Parameter(2) = GenerateInputParameter("@StationId", adInteger, 4, tmpStationNo)
                    Parameter(3) = GenerateInputParameter("@btnNum", adInteger, 4, TempButton)
                    
                    RunParametricStoredProcedure "Delete_tGood_Menu_By_GoodCode", Parameter
                    
                    On Error GoTo 0
                    
                    TempButton = 0
                    MyFormAddEditMode = ViewButton
                    UpdateToolbars
                End If
        End Select
    End With
    
    Exit Sub
    
RollBack:

    MyFormAddEditMode = ViewButton
    UpdateToolbars
    MsgBox "„ «”›«‰Â  €ÌÌ—«  ﬁ«»· «⁄„«· ‰Ì” "

End Sub

Private Sub vsKeyBoardDefined_KeyDown(KeyCode As Integer, Shift As Integer)

    With vsKeyBoardDefined
    
        If KeyCode <> 32 Or .Rows < 2 Or .Row < 1 Or .Col <> 4 Then Exit Sub
        
        Select Case MyFormAddEditMode
        
            Case DeleteButton
            
                If TempButton <> 0 Then
                    .Select .Row, .Col
                    .EditCell
                End If
                If Val(.TextMatrix(.Row, .Col)) = -1 Then
                    For i = 1 To .Rows - 1
                        If i <> .Row Then
                            .TextMatrix(i, 4) = 0
                        End If
                    Next i
                End If
                
        End Select
        
    End With


End Sub

Private Sub vsKeyBoardDefined_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    With vsKeyBoardDefined
    
        If Button <> 1 Or .Rows < 2 Or .Row < 1 Or .Col <> 4 Then Exit Sub
        
        Select Case MyFormAddEditMode
            Case DeleteButton
            
                If TempButton <> 0 Then
                    .Select .Row, .Col
                    .EditCell
                End If
                
                If Val(.TextMatrix(.Row, .Col)) = -1 Then
                    For i = 1 To .Rows - 1
                        If i <> .Row Then
                            .TextMatrix(i, 4) = 0
                        End If
                    Next i
                End If
        End Select
        
    End With

End Sub


Private Sub vsNoKeyBoardDefined_KeyDown(KeyCode As Integer, Shift As Integer)

    With vsNoKeyBoardDefined
    
        If KeyCode <> 32 Or .Rows < 2 Or .Row < 1 Or .Col <> 4 Then Exit Sub
        
        Select Case MyFormAddEditMode
        
            Case CodeToButton
            
                If TempButton <> 0 Then
                    .Select .Row, .Col
                    .EditCell
                End If
                
                                
        End Select
        
    End With

End Sub

Private Sub vsNoKeyBoardDefined_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    With vsNoKeyBoardDefined
    
        If Button <> 1 Or .Rows < 2 Or .Row < 1 Or .Col <> 4 Then Exit Sub
        
        Select Case MyFormAddEditMode
            Case CodeToButton
            
                If TempButton <> 0 Then
                    .Select .Row, .Col
                    .EditCell
                End If
                
        End Select
        
    End With
End Sub


