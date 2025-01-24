VERSION 5.00
Object = "{7AEDC602-D94C-11D1-BB7A-00E0290EA3C9}#1.0#0"; "ResizeKit.ocx"
Object = "{75D4F767-8785-11D3-93AD-0000832EF44D}#2.14#0"; "FAST2001.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmTables 
   BackColor       =   &H00C0FFC0&
   Caption         =   "                                                                                                                 áíÓÊ ãíÒåÇ"
   ClientHeight    =   8835
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11775
   BeginProperty Font 
      Name            =   "Nazanin"
      Size            =   9.75
      Charset         =   178
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmTables.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   8835
   ScaleWidth      =   11775
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame FrameHeader 
      BackColor       =   &H00C0FFC0&
      Height          =   735
      Left            =   240
      RightToLeft     =   -1  'True
      TabIndex        =   26
      Top             =   -120
      Width           =   11415
      Begin VB.CheckBox chkViewSpecial 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0FFC0&
         Caption         =   "ÝíáÊÑ ÈÑ ÇÓÇÓ äÝÑ"
         Height          =   420
         Left            =   9480
         RightToLeft     =   -1  'True
         TabIndex        =   36
         Top             =   240
         Width           =   1815
      End
      Begin FLWCtrls.FWNumericTextBox txtInterval 
         Height          =   480
         Left            =   840
         TabIndex        =   35
         Top             =   200
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   847
         Min             =   1
         BackColor       =   12648384
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "B Nazanin"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.TextBox TxtFontSize 
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
         Height          =   420
         Left            =   3600
         RightToLeft     =   -1  'True
         TabIndex        =   34
         Text            =   "12"
         Top             =   240
         Width           =   720
      End
      Begin VB.TextBox TxtHeight 
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
         Height          =   420
         Left            =   6240
         RightToLeft     =   -1  'True
         TabIndex        =   30
         Text            =   "850"
         Top             =   225
         Width           =   720
      End
      Begin VB.TextBox TxtWidth 
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
         Height          =   420
         Left            =   7920
         RightToLeft     =   -1  'True
         TabIndex        =   28
         Text            =   "1400"
         Top             =   225
         Width           =   720
      End
      Begin VB.Label Label32 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "ÓÇíÒ ÝæäÊ"
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
         Left            =   4440
         RightToLeft     =   -1  'True
         TabIndex        =   33
         Top             =   240
         Width           =   975
      End
      Begin VB.Label lblIntervalTitle 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "ÒãÇä ÈÑæÒ ÑÓÇäí :"
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
         Left            =   1560
         RightToLeft     =   -1  'True
         TabIndex        =   32
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label lblSecondTitle 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "ËÇäíå"
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
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   31
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label31 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "ÚÑÖ"
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
         Left            =   7200
         RightToLeft     =   -1  'True
         TabIndex        =   29
         Top             =   240
         Width           =   495
      End
      Begin VB.Label Label30 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Øæá"
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
         Left            =   8760
         RightToLeft     =   -1  'True
         TabIndex        =   27
         Top             =   330
         Width           =   495
      End
   End
   Begin VB.Timer Timer1 
      Left            =   360
      Top             =   0
   End
   Begin VB.Frame frameMap 
      Height          =   600
      Left            =   240
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   8280
      Width           =   11415
      Begin VB.PictureBox Picture4 
         BackColor       =   &H0000FFFF&
         Height          =   255
         Left            =   4800
         RightToLeft     =   -1  'True
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   19
         Top             =   240
         Width           =   255
      End
      Begin VB.PictureBox Picture3 
         BackColor       =   &H000000C0&
         Height          =   255
         Left            =   2160
         RightToLeft     =   -1  'True
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   11
         Top             =   240
         Width           =   255
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H000080FF&
         Height          =   255
         Left            =   7560
         RightToLeft     =   -1  'True
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   8
         Top             =   240
         Width           =   255
      End
      Begin VB.PictureBox Picture1 
         BackColor       =   &H0000C000&
         Height          =   255
         Left            =   10320
         RightToLeft     =   -1  'True
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   7
         Top             =   240
         Width           =   255
      End
      Begin VB.Label LblOverTime 
         Alignment       =   1  'Right Justify
         Caption         =   "ãíÒ ÈÇ ÊÇíã ÇÖÇÝí"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   3240
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label lblTableWithInvoicePrint 
         Alignment       =   1  'Right Justify
         Caption         =   "Ç ÝÇßÊæÑÑÝÊå ÔÏå"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label lblReservedTables 
         Alignment       =   1  'Right Justify
         Caption         =   "ãíÒ Ñ"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6480
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   240
         Width           =   975
      End
      Begin VB.Label lblEmptyTables 
         Alignment       =   1  'Right Justify
         Caption         =   "ãíÒ ÎÇáí"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   9480
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   240
         Width           =   735
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7575
      Left            =   240
      TabIndex        =   0
      Top             =   720
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   13361
      _Version        =   393216
      Tabs            =   20
      Tab             =   1
      TabsPerRow      =   10
      TabHeight       =   520
      BackColor       =   16777088
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Ñæå Çæá"
      TabPicture(0)   =   "frmTables.frx":A4C2
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame1"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Ñæå Ïæã"
      TabPicture(1)   =   "frmTables.frx":A4DE
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Frame2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Ñæå Óæã"
      TabPicture(2)   =   "frmTables.frx":A4FA
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame3"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Ñæå åÇÑã"
      TabPicture(3)   =   "frmTables.frx":A516
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame4"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "Ñæå äÌã"
      TabPicture(4)   =   "frmTables.frx":A532
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "Frame5"
      Tab(4).ControlCount=   1
      TabCaption(5)   =   "Tab 5"
      TabPicture(5)   =   "frmTables.frx":A54E
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "Frame6"
      Tab(5).ControlCount=   1
      TabCaption(6)   =   "Tab 6"
      TabPicture(6)   =   "frmTables.frx":A56A
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "Frame7"
      Tab(6).ControlCount=   1
      TabCaption(7)   =   "Tab 7"
      TabPicture(7)   =   "frmTables.frx":A586
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "Frame8"
      Tab(7).ControlCount=   1
      TabCaption(8)   =   "Tab 8"
      TabPicture(8)   =   "frmTables.frx":A5A2
      Tab(8).ControlEnabled=   0   'False
      Tab(8).Control(0)=   "Frame9"
      Tab(8).ControlCount=   1
      TabCaption(9)   =   "Tab 9"
      TabPicture(9)   =   "frmTables.frx":A5BE
      Tab(9).ControlEnabled=   0   'False
      Tab(9).Control(0)=   "Frame10"
      Tab(9).ControlCount=   1
      TabCaption(10)  =   "Tab 10"
      TabPicture(10)  =   "frmTables.frx":A5DA
      Tab(10).ControlEnabled=   0   'False
      Tab(10).Control(0)=   "Frame11"
      Tab(10).ControlCount=   1
      TabCaption(11)  =   "Tab 11"
      TabPicture(11)  =   "frmTables.frx":A5F6
      Tab(11).ControlEnabled=   0   'False
      Tab(11).Control(0)=   "Frame12"
      Tab(11).ControlCount=   1
      TabCaption(12)  =   "Tab 12"
      TabPicture(12)  =   "frmTables.frx":A612
      Tab(12).ControlEnabled=   0   'False
      Tab(12).Control(0)=   "Frame13"
      Tab(12).ControlCount=   1
      TabCaption(13)  =   "Tab 13"
      TabPicture(13)  =   "frmTables.frx":A62E
      Tab(13).ControlEnabled=   0   'False
      Tab(13).Control(0)=   "Frame14"
      Tab(13).ControlCount=   1
      TabCaption(14)  =   "Tab 14"
      TabPicture(14)  =   "frmTables.frx":A64A
      Tab(14).ControlEnabled=   0   'False
      Tab(14).Control(0)=   "Frame15"
      Tab(14).ControlCount=   1
      TabCaption(15)  =   "Tab 15"
      TabPicture(15)  =   "frmTables.frx":A666
      Tab(15).ControlEnabled=   0   'False
      Tab(15).Control(0)=   "Frame16"
      Tab(15).ControlCount=   1
      TabCaption(16)  =   "Tab 16"
      TabPicture(16)  =   "frmTables.frx":A682
      Tab(16).ControlEnabled=   0   'False
      Tab(16).Control(0)=   "Frame17"
      Tab(16).ControlCount=   1
      TabCaption(17)  =   "Tab 17"
      TabPicture(17)  =   "frmTables.frx":A69E
      Tab(17).ControlEnabled=   0   'False
      Tab(17).Control(0)=   "Frame18"
      Tab(17).ControlCount=   1
      TabCaption(18)  =   "Tab 18"
      TabPicture(18)  =   "frmTables.frx":A6BA
      Tab(18).ControlEnabled=   0   'False
      Tab(18).Control(0)=   "Frame19"
      Tab(18).ControlCount=   1
      TabCaption(19)  =   "Tab 19"
      TabPicture(19)  =   "frmTables.frx":A6D6
      Tab(19).ControlEnabled=   0   'False
      Tab(19).Control(0)=   "Frame20"
      Tab(19).ControlCount=   1
      Begin VB.Frame Frame20 
         BackColor       =   &H00FF8080&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6615
         Left            =   -74880
         RightToLeft     =   -1  'True
         TabIndex        =   79
         Top             =   840
         Width           =   11175
         Begin FLWCtrls.FWCoolButton cmd19 
            Height          =   850
            Index           =   0
            Left            =   360
            TabIndex        =   80
            Top             =   360
            Visible         =   0   'False
            Width           =   1400
            _ExtentX        =   2461
            _ExtentY        =   1508
            BackColor       =   -2147483633
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Nazanin"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   ""
            MaskColor       =   -2147483633
         End
         Begin VB.Label Label19 
            Alignment       =   2  'Center
            Caption         =   "Label19"
            BeginProperty Font 
               Name            =   "Nazanin"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Index           =   0
            Left            =   360
            RightToLeft     =   -1  'True
            TabIndex        =   81
            Top             =   1320
            Visible         =   0   'False
            Width           =   1335
            WordWrap        =   -1  'True
         End
      End
      Begin VB.Frame Frame19 
         BackColor       =   &H00FF8080&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6615
         Left            =   -74880
         RightToLeft     =   -1  'True
         TabIndex        =   76
         Top             =   840
         Width           =   11175
         Begin FLWCtrls.FWCoolButton cmd18 
            Height          =   850
            Index           =   0
            Left            =   360
            TabIndex        =   77
            Top             =   360
            Visible         =   0   'False
            Width           =   1400
            _ExtentX        =   2461
            _ExtentY        =   1508
            BackColor       =   -2147483633
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Nazanin"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   ""
            MaskColor       =   -2147483633
         End
         Begin VB.Label Label18 
            Alignment       =   2  'Center
            Caption         =   "Label18"
            BeginProperty Font 
               Name            =   "Nazanin"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Index           =   0
            Left            =   360
            RightToLeft     =   -1  'True
            TabIndex        =   78
            Top             =   1320
            Visible         =   0   'False
            Width           =   1335
            WordWrap        =   -1  'True
         End
      End
      Begin VB.Frame Frame18 
         BackColor       =   &H00FF8080&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6615
         Left            =   -74880
         RightToLeft     =   -1  'True
         TabIndex        =   73
         Top             =   840
         Width           =   11175
         Begin FLWCtrls.FWCoolButton cmd17 
            Height          =   850
            Index           =   0
            Left            =   360
            TabIndex        =   74
            Top             =   360
            Visible         =   0   'False
            Width           =   1400
            _ExtentX        =   2461
            _ExtentY        =   1508
            BackColor       =   -2147483633
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Nazanin"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   ""
            MaskColor       =   -2147483633
         End
         Begin VB.Label Label17 
            Alignment       =   2  'Center
            Caption         =   "Label17"
            BeginProperty Font 
               Name            =   "Nazanin"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Index           =   0
            Left            =   360
            RightToLeft     =   -1  'True
            TabIndex        =   75
            Top             =   1320
            Visible         =   0   'False
            Width           =   1335
            WordWrap        =   -1  'True
         End
      End
      Begin VB.Frame Frame17 
         BackColor       =   &H00FF8080&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6615
         Left            =   -74880
         RightToLeft     =   -1  'True
         TabIndex        =   70
         Top             =   840
         Width           =   11175
         Begin FLWCtrls.FWCoolButton cmd16 
            Height          =   850
            Index           =   0
            Left            =   360
            TabIndex        =   71
            Top             =   360
            Visible         =   0   'False
            Width           =   1400
            _ExtentX        =   2461
            _ExtentY        =   1508
            BackColor       =   -2147483633
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Nazanin"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   ""
            MaskColor       =   -2147483633
         End
         Begin VB.Label Label16 
            Alignment       =   2  'Center
            Caption         =   "Label16"
            BeginProperty Font 
               Name            =   "Nazanin"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Index           =   0
            Left            =   360
            RightToLeft     =   -1  'True
            TabIndex        =   72
            Top             =   1320
            Visible         =   0   'False
            Width           =   1335
            WordWrap        =   -1  'True
         End
      End
      Begin VB.Frame Frame16 
         BackColor       =   &H00FF8080&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6615
         Left            =   -74880
         RightToLeft     =   -1  'True
         TabIndex        =   67
         Top             =   840
         Width           =   11175
         Begin FLWCtrls.FWCoolButton cmd15 
            Height          =   850
            Index           =   0
            Left            =   360
            TabIndex        =   68
            Top             =   360
            Visible         =   0   'False
            Width           =   1400
            _ExtentX        =   2461
            _ExtentY        =   1508
            BackColor       =   -2147483633
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Nazanin"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   ""
            MaskColor       =   -2147483633
         End
         Begin VB.Label Label15 
            Alignment       =   2  'Center
            Caption         =   "Label15"
            BeginProperty Font 
               Name            =   "Nazanin"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Index           =   0
            Left            =   360
            RightToLeft     =   -1  'True
            TabIndex        =   69
            Top             =   1320
            Visible         =   0   'False
            Width           =   1335
            WordWrap        =   -1  'True
         End
      End
      Begin VB.Frame Frame15 
         BackColor       =   &H00FF8080&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6615
         Left            =   -74880
         RightToLeft     =   -1  'True
         TabIndex        =   64
         Top             =   840
         Width           =   11175
         Begin FLWCtrls.FWCoolButton cmd14 
            Height          =   850
            Index           =   0
            Left            =   360
            TabIndex        =   65
            Top             =   360
            Visible         =   0   'False
            Width           =   1400
            _ExtentX        =   2461
            _ExtentY        =   1508
            BackColor       =   -2147483633
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Nazanin"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   ""
            MaskColor       =   -2147483633
         End
         Begin VB.Label Label14 
            Alignment       =   2  'Center
            Caption         =   "Label14"
            BeginProperty Font 
               Name            =   "Nazanin"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Index           =   0
            Left            =   360
            RightToLeft     =   -1  'True
            TabIndex        =   66
            Top             =   1320
            Visible         =   0   'False
            Width           =   1335
            WordWrap        =   -1  'True
         End
      End
      Begin VB.Frame Frame14 
         BackColor       =   &H00FF8080&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6615
         Left            =   -74880
         RightToLeft     =   -1  'True
         TabIndex        =   61
         Top             =   840
         Width           =   11175
         Begin FLWCtrls.FWCoolButton cmd13 
            Height          =   850
            Index           =   0
            Left            =   360
            TabIndex        =   62
            Top             =   360
            Visible         =   0   'False
            Width           =   1400
            _ExtentX        =   2461
            _ExtentY        =   1508
            BackColor       =   -2147483633
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Nazanin"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   ""
            MaskColor       =   -2147483633
         End
         Begin VB.Label Label13 
            Alignment       =   2  'Center
            Caption         =   "Label13"
            BeginProperty Font 
               Name            =   "Nazanin"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Index           =   0
            Left            =   360
            RightToLeft     =   -1  'True
            TabIndex        =   63
            Top             =   1320
            Visible         =   0   'False
            Width           =   1335
            WordWrap        =   -1  'True
         End
      End
      Begin VB.Frame Frame13 
         BackColor       =   &H00FF8080&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6615
         Left            =   -74880
         RightToLeft     =   -1  'True
         TabIndex        =   58
         Top             =   840
         Width           =   11175
         Begin FLWCtrls.FWCoolButton cmd12 
            Height          =   850
            Index           =   0
            Left            =   360
            TabIndex        =   59
            Top             =   360
            Visible         =   0   'False
            Width           =   1400
            _ExtentX        =   2461
            _ExtentY        =   1508
            BackColor       =   -2147483633
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Nazanin"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   ""
            MaskColor       =   -2147483633
         End
         Begin VB.Label Label12 
            Alignment       =   2  'Center
            Caption         =   "Label12"
            BeginProperty Font 
               Name            =   "Nazanin"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Index           =   0
            Left            =   360
            RightToLeft     =   -1  'True
            TabIndex        =   60
            Top             =   1320
            Visible         =   0   'False
            Width           =   1335
            WordWrap        =   -1  'True
         End
      End
      Begin VB.Frame Frame12 
         BackColor       =   &H00FF8080&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6615
         Left            =   -74880
         RightToLeft     =   -1  'True
         TabIndex        =   55
         Top             =   840
         Width           =   11175
         Begin FLWCtrls.FWCoolButton cmd11 
            Height          =   850
            Index           =   0
            Left            =   360
            TabIndex        =   56
            Top             =   360
            Visible         =   0   'False
            Width           =   1400
            _ExtentX        =   2461
            _ExtentY        =   1508
            BackColor       =   -2147483633
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Nazanin"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   ""
            MaskColor       =   -2147483633
         End
         Begin VB.Label Label11 
            Alignment       =   2  'Center
            Caption         =   "Label11"
            BeginProperty Font 
               Name            =   "Nazanin"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Index           =   0
            Left            =   360
            RightToLeft     =   -1  'True
            TabIndex        =   57
            Top             =   1320
            Visible         =   0   'False
            Width           =   1335
            WordWrap        =   -1  'True
         End
      End
      Begin VB.Frame Frame11 
         BackColor       =   &H00FF8080&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6615
         Left            =   -74880
         RightToLeft     =   -1  'True
         TabIndex        =   52
         Top             =   840
         Width           =   11175
         Begin FLWCtrls.FWCoolButton cmd10 
            Height          =   850
            Index           =   0
            Left            =   360
            TabIndex        =   53
            Top             =   360
            Visible         =   0   'False
            Width           =   1400
            _ExtentX        =   2461
            _ExtentY        =   1508
            BackColor       =   -2147483633
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Nazanin"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   ""
            MaskColor       =   -2147483633
         End
         Begin VB.Label Label10 
            Alignment       =   2  'Center
            Caption         =   "Label10"
            BeginProperty Font 
               Name            =   "Nazanin"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Index           =   0
            Left            =   360
            RightToLeft     =   -1  'True
            TabIndex        =   54
            Top             =   1320
            Visible         =   0   'False
            Width           =   1335
            WordWrap        =   -1  'True
         End
      End
      Begin VB.Frame Frame10 
         BackColor       =   &H00FF8080&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6615
         Left            =   -74880
         RightToLeft     =   -1  'True
         TabIndex        =   49
         Top             =   840
         Width           =   11175
         Begin FLWCtrls.FWCoolButton cmd9 
            Height          =   850
            Index           =   0
            Left            =   360
            TabIndex        =   50
            Top             =   360
            Visible         =   0   'False
            Width           =   1400
            _ExtentX        =   2461
            _ExtentY        =   1508
            BackColor       =   -2147483633
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Nazanin"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   ""
            MaskColor       =   -2147483633
         End
         Begin VB.Label Label9 
            Alignment       =   2  'Center
            Caption         =   "Label9"
            BeginProperty Font 
               Name            =   "Nazanin"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Index           =   0
            Left            =   360
            RightToLeft     =   -1  'True
            TabIndex        =   51
            Top             =   1320
            Visible         =   0   'False
            Width           =   1335
            WordWrap        =   -1  'True
         End
      End
      Begin VB.Frame Frame9 
         BackColor       =   &H00FF8080&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6615
         Left            =   -74880
         RightToLeft     =   -1  'True
         TabIndex        =   46
         Top             =   840
         Width           =   11175
         Begin FLWCtrls.FWCoolButton cmd8 
            Height          =   850
            Index           =   0
            Left            =   360
            TabIndex        =   47
            Top             =   360
            Visible         =   0   'False
            Width           =   1400
            _ExtentX        =   2461
            _ExtentY        =   1508
            BackColor       =   -2147483633
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Nazanin"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   ""
            MaskColor       =   -2147483633
         End
         Begin VB.Label Label8 
            Alignment       =   2  'Center
            Caption         =   "Label8"
            BeginProperty Font 
               Name            =   "Nazanin"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Index           =   0
            Left            =   360
            RightToLeft     =   -1  'True
            TabIndex        =   48
            Top             =   1320
            Visible         =   0   'False
            Width           =   1335
            WordWrap        =   -1  'True
         End
      End
      Begin VB.Frame Frame8 
         BackColor       =   &H00FF8080&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6615
         Left            =   -74880
         RightToLeft     =   -1  'True
         TabIndex        =   43
         Top             =   840
         Width           =   11175
         Begin FLWCtrls.FWCoolButton cmd7 
            Height          =   850
            Index           =   0
            Left            =   360
            TabIndex        =   44
            Top             =   360
            Visible         =   0   'False
            Width           =   1400
            _ExtentX        =   2461
            _ExtentY        =   1508
            BackColor       =   -2147483633
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Nazanin"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   ""
            MaskColor       =   -2147483633
         End
         Begin VB.Label Label7 
            Alignment       =   2  'Center
            Caption         =   "Label7"
            BeginProperty Font 
               Name            =   "Nazanin"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Index           =   0
            Left            =   360
            RightToLeft     =   -1  'True
            TabIndex        =   45
            Top             =   1320
            Visible         =   0   'False
            Width           =   1335
            WordWrap        =   -1  'True
         End
      End
      Begin VB.Frame Frame7 
         BackColor       =   &H00FF8080&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6615
         Left            =   -74880
         RightToLeft     =   -1  'True
         TabIndex        =   40
         Top             =   840
         Width           =   11175
         Begin FLWCtrls.FWCoolButton cmd6 
            Height          =   850
            Index           =   0
            Left            =   360
            TabIndex        =   41
            Top             =   360
            Visible         =   0   'False
            Width           =   1400
            _ExtentX        =   2461
            _ExtentY        =   1508
            BackColor       =   -2147483633
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Nazanin"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   ""
            MaskColor       =   -2147483633
         End
         Begin VB.Label Label6 
            Alignment       =   2  'Center
            Caption         =   "Label6"
            BeginProperty Font 
               Name            =   "Nazanin"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Index           =   0
            Left            =   360
            RightToLeft     =   -1  'True
            TabIndex        =   42
            Top             =   1320
            Visible         =   0   'False
            Width           =   1335
            WordWrap        =   -1  'True
         End
      End
      Begin VB.Frame Frame6 
         BackColor       =   &H00FF8080&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6615
         Left            =   -74880
         RightToLeft     =   -1  'True
         TabIndex        =   37
         Top             =   840
         Width           =   11175
         Begin FLWCtrls.FWCoolButton cmd5 
            Height          =   850
            Index           =   0
            Left            =   360
            TabIndex        =   38
            Top             =   360
            Visible         =   0   'False
            Width           =   1400
            _ExtentX        =   2461
            _ExtentY        =   1508
            BackColor       =   -2147483633
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Nazanin"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   ""
            MaskColor       =   -2147483633
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            Caption         =   "Label5"
            BeginProperty Font 
               Name            =   "Nazanin"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Index           =   0
            Left            =   360
            RightToLeft     =   -1  'True
            TabIndex        =   39
            Top             =   1320
            Visible         =   0   'False
            Width           =   1335
            WordWrap        =   -1  'True
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00FF8080&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6615
         Left            =   -74880
         RightToLeft     =   -1  'True
         TabIndex        =   17
         Top             =   780
         Width           =   11175
         Begin FLWCtrls.FWCoolButton cmd4 
            Height          =   850
            Index           =   0
            Left            =   360
            TabIndex        =   18
            Top             =   360
            Visible         =   0   'False
            Width           =   1400
            _ExtentX        =   2461
            _ExtentY        =   1508
            BackColor       =   -2147483633
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Nazanin"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   ""
            MaskColor       =   -2147483633
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            Caption         =   "Label4"
            BeginProperty Font 
               Name            =   "Nazanin"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Index           =   0
            Left            =   360
            RightToLeft     =   -1  'True
            TabIndex        =   25
            Top             =   1320
            Visible         =   0   'False
            Width           =   1335
            WordWrap        =   -1  'True
         End
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6615
         Left            =   -74880
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   780
         Width           =   11175
         Begin FLWCtrls.FWCoolButton cmd3 
            Height          =   850
            Index           =   0
            Left            =   360
            TabIndex        =   16
            Top             =   360
            Visible         =   0   'False
            Width           =   1400
            _ExtentX        =   2461
            _ExtentY        =   1508
            BackColor       =   -2147483633
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Nazanin"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   ""
            MaskColor       =   -2147483633
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            Caption         =   "Label3"
            BeginProperty Font 
               Name            =   "Nazanin"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Index           =   0
            Left            =   360
            RightToLeft     =   -1  'True
            TabIndex        =   24
            Top             =   1320
            Visible         =   0   'False
            Width           =   1335
            WordWrap        =   -1  'True
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00C0FFC0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6615
         Left            =   -74880
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   780
         Width           =   11175
         Begin FLWCtrls.FWCoolButton cmd2 
            Height          =   850
            Index           =   0
            Left            =   360
            TabIndex        =   5
            Top             =   360
            Visible         =   0   'False
            Width           =   1400
            _ExtentX        =   2461
            _ExtentY        =   1508
            BackColor       =   -2147483633
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Nazanin"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   ""
            MaskColor       =   -2147483633
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            Caption         =   "Label2"
            BeginProperty Font 
               Name            =   "Nazanin"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Index           =   0
            Left            =   360
            RightToLeft     =   -1  'True
            TabIndex        =   23
            Top             =   1320
            Visible         =   0   'False
            Width           =   1335
            WordWrap        =   -1  'True
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6615
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   780
         Width           =   11175
         Begin FLWCtrls.FWCoolButton cmd1 
            Height          =   855
            Index           =   0
            Left            =   360
            TabIndex        =   14
            Top             =   360
            Visible         =   0   'False
            Width           =   1400
            _ExtentX        =   2461
            _ExtentY        =   1508
            BackColor       =   -2147483633
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Nazanin"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   ""
            MaskColor       =   -2147483633
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Caption         =   "Label1"
            BeginProperty Font 
               Name            =   "Nazanin"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Index           =   0
            Left            =   360
            RightToLeft     =   -1  'True
            TabIndex        =   22
            Top             =   1320
            Visible         =   0   'False
            Width           =   1335
            WordWrap        =   -1  'True
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H0080C0FF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6615
         Left            =   -74880
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   780
         Width           =   11295
         Begin FLWCtrls.FWCoolButton cmd 
            Height          =   855
            Index           =   0
            Left            =   360
            TabIndex        =   4
            Top             =   360
            Visible         =   0   'False
            Width           =   1400
            _ExtentX        =   2461
            _ExtentY        =   1508
            BackColor       =   -2147483633
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Nazanin"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   ""
            MaskColor       =   -2147483633
         End
         Begin VB.Label Label 
            Alignment       =   2  'Center
            Caption         =   "Label"
            BeginProperty Font 
               Name            =   "Nazanin"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   855
            Index           =   0
            Left            =   360
            RightToLeft     =   -1  'True
            TabIndex        =   21
            Top             =   1320
            Visible         =   0   'False
            Width           =   1335
            WordWrap        =   -1  'True
         End
      End
   End
   Begin RESIZEKITLibCtl.ResizeKit ResizeKit1 
      Height          =   480
      Left            =   0
      OleObjectBlob   =   "frmTables.frx":A6F2
      TabIndex        =   13
      Top             =   240
      Width           =   480
   End
End
Attribute VB_Name = "frmTables"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Parameter() As Parameter
Dim Rst As Recordset
Private Type Position
    X As Long
    Y As Long
End Type
Dim lastPosition(0 To 19) As Position
Dim i, j, k, m, l, i0, i1, i2, i3, i4, i5, i6, i7, i8, i9, i10, i11, i12, i13, i14, i15, i16, i17, i18, i19 As Integer
Dim varCmd As CommandButton
Dim varlbl As Label
Dim tablesCount As Integer
Dim Part(0 To 19) As Integer

Private Sub chkViewSpecial_Click()
    GetTables CurrentBranch
End Sub

Private Sub Form_Activate()
    GetCount CurrentBranch
    If tablesCount > 0 Then
        If clsArya.HardLockSerialNo = "92072202873" Then
            chkViewSpecial.Value = 1
        Else
            GetTables CurrentBranch
        End If
'''''        DetectReservedTables CurrentBranch
        DetectBusyTables CurrentBranch
        DetectOtherTables CurrentBranch
    End If
    For i = 0 To 19
        If Part(i) = clsStation.PartitionID Then
            SSTab1.Tab = i
            Exit For
        End If
    Next
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case Shift
          Case 0
              Select Case KeyCode
                  Case 27  ' Esc
            
                  Unload Me
              End Select
          Case 2
               Select Case KeyCode
                  Case 123  'Exit
                     If clsStation.KeyboardType = EnumKeyBoardType.Rb2 Then
                        Unload Me
                     End If
              End Select

    End Select

End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call PresetScreenSaver
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call PresetScreenSaver
End Sub
Private Sub EmptyTable(TableNo As Integer)
    
    ReDim Parameter(0) As Parameter
    Parameter(0) = GenerateInputParameter("@intTableNo", adInteger, 4, TableNo)
    RunParametricStoredProcedure "Update_tTable_By_Empty", Parameter

End Sub
Private Sub label_Click(index As Integer)
    If Part(0) = clsStation.PartitionID Or clsStation.OtherPartition = False Then
        If CheckTableState(Label(index).Tag) = True Then      'If empty
            mvarTable = Val(Label(index).Tag)
        Else
            If clsArya.HardLockSerialNo = "92072202873" Then
                ShowMessage "ÂíÇ ãí ÎæÇåíÏ ãíÒ ãæÑÏ äÙÑ ÎÇáí ÔæÏ ¿", True, True, "Èáí", "ÎíÑ"
                If mvarMsgIdx = vbYes Then
                    mvarTable = Val(Label(index).Tag)
                    EmptyTable mvarTable
                    Unload Me
                    Exit Sub
                Else
                End If
            End If
            mvarTable = 0
            FindInvoice Label(index).Tag, CurrentBranch
        End If
        Unload Me
    Else
        ShowDisMessage "ÈÎÔ ãæÑÏ äÙÑ ÈÇ ÈÎÔ Çíä ÇíÓÊÇå åãÎæÇäí äÏÇÑÏ", 1000
    End If
End Sub

Private Sub Label1_Click(index As Integer)
    If Part(1) = clsStation.PartitionID Or clsStation.OtherPartition = False Then
        If CheckTableState(Label1(index).Tag) = True Then
            mvarTable = Val(Label1(index).Tag)
        Else
            If clsArya.HardLockSerialNo = "92072202873" Then
                ShowMessage "ÂíÇ ãí ÎæÇåíÏ ãíÒ ãæÑÏ äÙÑ ÎÇáí ÔæÏ ¿", True, True, "Èáí", "ÎíÑ"
                If mvarMsgIdx = vbYes Then
                    mvarTable = Val(Label1(index).Tag)
                    EmptyTable mvarTable
                    Unload Me
                    Exit Sub
                Else
                End If
            End If
            mvarTable = 0
            FindInvoice Label1(index).Tag, CurrentBranch
        End If
        Unload Me
    Else
        ShowDisMessage "ÈÎÔ ãæÑÏ äÙÑ ÈÇ ÈÎÔ Çíä ÇíÓÊÇå åãÎæÇäí äÏÇÑÏ", 1000
    End If
End Sub

Private Sub label2_Click(index As Integer)
    If Part(2) = clsStation.PartitionID Or clsStation.OtherPartition = False Then
        If CheckTableState(Label2(index).Tag) = True Then
            mvarTable = Val(Label2(index).Tag)
        Else
            If clsArya.HardLockSerialNo = "92072202873" Then
                ShowMessage "ÂíÇ ãí ÎæÇåíÏ ãíÒ ãæÑÏ äÙÑ ÎÇáí ÔæÏ ¿", True, True, "Èáí", "ÎíÑ"
                If mvarMsgIdx = vbYes Then
                    mvarTable = Val(Label2(index).Tag)
                    EmptyTable mvarTable
                    Unload Me
                    Exit Sub
                Else
                End If
            End If
            mvarTable = 0
            FindInvoice Label2(index).Tag, CurrentBranch
        End If
        Unload Me
    Else
        ShowDisMessage "ÈÎÔ ãæÑÏ äÙÑ ÈÇ ÈÎÔ Çíä ÇíÓÊÇå åãÎæÇäí äÏÇÑÏ", 1000
    End If
End Sub
Private Sub label3_Click(index As Integer)
    If Part(3) = clsStation.PartitionID Or clsStation.OtherPartition = False Then
        If CheckTableState(Label3(index).Tag) = True Then
            mvarTable = Val(Label3(index).Tag)
        Else
            If clsArya.HardLockSerialNo = "92072202873" Then
                ShowMessage "ÂíÇ ãí ÎæÇåíÏ ãíÒ ãæÑÏ äÙÑ ÎÇáí ÔæÏ ¿", True, True, "Èáí", "ÎíÑ"
                If mvarMsgIdx = vbYes Then
                    mvarTable = Val(Label3(index).Tag)
                    EmptyTable mvarTable
                    Unload Me
                    Exit Sub
                Else
                End If
            End If
            mvarTable = 0
            FindInvoice Label3(index).Tag, CurrentBranch
        End If
        Unload Me
    Else
        ShowDisMessage "ÈÎÔ ãæÑÏ äÙÑ ÈÇ ÈÎÔ Çíä ÇíÓÊÇå åãÎæÇäí äÏÇÑÏ", 1000
    End If
End Sub
Private Sub label4_Click(index As Integer)
    If Part(4) = clsStation.PartitionID Or clsStation.OtherPartition = False Then
        If CheckTableState(Label4(index).Tag) = True Then
            mvarTable = Val(Label4(index).Tag)
        Else
            If clsArya.HardLockSerialNo = "92072202873" Then
                ShowMessage "ÂíÇ ãí ÎæÇåíÏ ãíÒ ãæÑÏ äÙÑ ÎÇáí ÔæÏ ¿", True, True, "Èáí", "ÎíÑ"
                If mvarMsgIdx = vbYes Then
                    mvarTable = Val(Label4(index).Tag)
                    EmptyTable mvarTable
                    Unload Me
                    Exit Sub
                Else
                End If
            End If
            mvarTable = 0
            FindInvoice Label4(index).Tag, CurrentBranch
        End If
        Unload Me
    Else
        ShowDisMessage "ÈÎÔ ãæÑÏ äÙÑ ÈÇ ÈÎÔ Çíä ÇíÓÊÇå åãÎæÇäí äÏÇÑÏ", 1000
    End If
End Sub
Private Sub Label5_Click(index As Integer)
    If Part(5) = clsStation.PartitionID Or clsStation.OtherPartition = False Then
        If CheckTableState(Label5(index).Tag) = True Then
            mvarTable = Val(Label5(index).Tag)
        Else
            If clsArya.HardLockSerialNo = "92072202873" Then
                ShowMessage "ÂíÇ ãí ÎæÇåíÏ ãíÒ ãæÑÏ äÙÑ ÎÇáí ÔæÏ ¿", True, True, "Èáí", "ÎíÑ"
                If mvarMsgIdx = vbYes Then
                    mvarTable = Val(Label5(index).Tag)
                    EmptyTable mvarTable
                    Unload Me
                    Exit Sub
                Else
                End If
            End If
            mvarTable = 0
            FindInvoice Label5(index).Tag, CurrentBranch
        End If
        Unload Me
    Else
        ShowDisMessage "ÈÎÔ ãæÑÏ äÙÑ ÈÇ ÈÎÔ Çíä ÇíÓÊÇå åãÎæÇäí äÏÇÑÏ", 1000
    End If
End Sub
Private Sub label6_Click(index As Integer)
    If Part(6) = clsStation.PartitionID Or clsStation.OtherPartition = False Then
        If CheckTableState(Label6(index).Tag) = True Then
            mvarTable = Val(Label6(index).Tag)
        Else
            If clsArya.HardLockSerialNo = "92072202873" Then
                ShowMessage "ÂíÇ ãí ÎæÇåíÏ ãíÒ ãæÑÏ äÙÑ ÎÇáí ÔæÏ ¿", True, True, "Èáí", "ÎíÑ"
                If mvarMsgIdx = vbYes Then
                    mvarTable = Val(Label6(index).Tag)
                    EmptyTable mvarTable
                    Unload Me
                    Exit Sub
                Else
                End If
            End If
            mvarTable = 0
            FindInvoice Label6(index).Tag, CurrentBranch
        End If
        Unload Me
    Else
        ShowDisMessage "ÈÎÔ ãæÑÏ äÙÑ ÈÇ ÈÎÔ Çíä ÇíÓÊÇå åãÎæÇäí äÏÇÑÏ", 1000
    End If
End Sub
Private Sub label7_Click(index As Integer)
    If Part(7) = clsStation.PartitionID Or clsStation.OtherPartition = False Then
        If CheckTableState(Label7(index).Tag) = True Then
            mvarTable = Val(Label7(index).Tag)
        Else
            If clsArya.HardLockSerialNo = "92072202873" Then
                ShowMessage "ÂíÇ ãí ÎæÇåíÏ ãíÒ ãæÑÏ äÙÑ ÎÇáí ÔæÏ ¿", True, True, "Èáí", "ÎíÑ"
                If mvarMsgIdx = vbYes Then
                    mvarTable = Val(Label7(index).Tag)
                    EmptyTable mvarTable
                    Unload Me
                    Exit Sub
                Else
                End If
            End If
            mvarTable = 0
            FindInvoice Label7(index).Tag, CurrentBranch
        End If
        Unload Me
    Else
        ShowDisMessage "ÈÎÔ ãæÑÏ äÙÑ ÈÇ ÈÎÔ Çíä ÇíÓÊÇå åãÎæÇäí äÏÇÑÏ", 1000
    End If
End Sub
Private Sub label8_Click(index As Integer)
    If Part(8) = clsStation.PartitionID Or clsStation.OtherPartition = False Then
        If CheckTableState(Label8(index).Tag) = True Then
            mvarTable = Val(Label8(index).Tag)
        Else
            If clsArya.HardLockSerialNo = "92072202873" Then
                ShowMessage "ÂíÇ ãí ÎæÇåíÏ ãíÒ ãæÑÏ äÙÑ ÎÇáí ÔæÏ ¿", True, True, "Èáí", "ÎíÑ"
                If mvarMsgIdx = vbYes Then
                    mvarTable = Val(Label8(index).Tag)
                    EmptyTable mvarTable
                    Unload Me
                    Exit Sub
                Else
                End If
            End If
            mvarTable = 0
            FindInvoice Label8(index).Tag, CurrentBranch
        End If
        Unload Me
    Else
        ShowDisMessage "ÈÎÔ ãæÑÏ äÙÑ ÈÇ ÈÎÔ Çíä ÇíÓÊÇå åãÎæÇäí äÏÇÑÏ", 1000
    End If
End Sub
Private Sub label9_Click(index As Integer)
    If Part(9) = clsStation.PartitionID Or clsStation.OtherPartition = False Then
        If CheckTableState(Label9(index).Tag) = True Then
            mvarTable = Val(Label9(index).Tag)
        Else
            If clsArya.HardLockSerialNo = "92072202873" Then
                ShowMessage "ÂíÇ ãí ÎæÇåíÏ ãíÒ ãæÑÏ äÙÑ ÎÇáí ÔæÏ ¿", True, True, "Èáí", "ÎíÑ"
                If mvarMsgIdx = vbYes Then
                    mvarTable = Val(Label9(index).Tag)
                    EmptyTable mvarTable
                    Unload Me
                    Exit Sub
                Else
                End If
            End If
            mvarTable = 0
            FindInvoice Label9(index).Tag, CurrentBranch
        End If
        Unload Me
    Else
        ShowDisMessage "ÈÎÔ ãæÑÏ äÙÑ ÈÇ ÈÎÔ Çíä ÇíÓÊÇå åãÎæÇäí äÏÇÑÏ", 1000
    End If
End Sub
Private Sub label10_Click(index As Integer)
    If Part(10) = clsStation.PartitionID Or clsStation.OtherPartition = False Then
        If CheckTableState(Label10(index).Tag) = True Then
            mvarTable = Val(Label10(index).Tag)
        Else
            If clsArya.HardLockSerialNo = "92072202873" Then
                ShowMessage "ÂíÇ ãí ÎæÇåíÏ ãíÒ ãæÑÏ äÙÑ ÎÇáí ÔæÏ ¿", True, True, "Èáí", "ÎíÑ"
                If mvarMsgIdx = vbYes Then
                    mvarTable = Val(Label10(index).Tag)
                    EmptyTable mvarTable
                    Unload Me
                    Exit Sub
                Else
                End If
            End If
            mvarTable = 0
            FindInvoice Label10(index).Tag, CurrentBranch
        End If
        Unload Me
    Else
        ShowDisMessage "ÈÎÔ ãæÑÏ äÙÑ ÈÇ ÈÎÔ Çíä ÇíÓÊÇå åãÎæÇäí äÏÇÑÏ", 1000
    End If
End Sub
Private Sub label11_Click(index As Integer)
    If Part(11) = clsStation.PartitionID Or clsStation.OtherPartition = False Then
        If CheckTableState(Label11(index).Tag) = True Then
            mvarTable = Val(Label11(index).Tag)
        Else
            If clsArya.HardLockSerialNo = "92072202873" Then
                ShowMessage "ÂíÇ ãí ÎæÇåíÏ ãíÒ ãæÑÏ äÙÑ ÎÇáí ÔæÏ ¿", True, True, "Èáí", "ÎíÑ"
                If mvarMsgIdx = vbYes Then
                    mvarTable = Val(Label11(index).Tag)
                    EmptyTable mvarTable
                    Unload Me
                    Exit Sub
                Else
                End If
            End If
            mvarTable = 0
            FindInvoice Label11(index).Tag, CurrentBranch
        End If
        Unload Me
    Else
        ShowDisMessage "ÈÎÔ ãæÑÏ äÙÑ ÈÇ ÈÎÔ Çíä ÇíÓÊÇå åãÎæÇäí äÏÇÑÏ", 1000
    End If
End Sub
Private Sub label12_Click(index As Integer)
    If Part(4) = clsStation.PartitionID Or clsStation.OtherPartition = False Then
        If CheckTableState(Label12(index).Tag) = True Then
            mvarTable = Val(Label12(index).Tag)
        Else
            If clsArya.HardLockSerialNo = "92072202873" Then
                ShowMessage "ÂíÇ ãí ÎæÇåíÏ ãíÒ ãæÑÏ äÙÑ ÎÇáí ÔæÏ ¿", True, True, "Èáí", "ÎíÑ"
                If mvarMsgIdx = vbYes Then
                    mvarTable = Val(Label12(index).Tag)
                    EmptyTable mvarTable
                    Unload Me
                    Exit Sub
                Else
                End If
            End If
            mvarTable = 0
            FindInvoice Label12(index).Tag, CurrentBranch
        End If
        Unload Me
    Else
        ShowDisMessage "ÈÎÔ ãæÑÏ äÙÑ ÈÇ ÈÎÔ Çíä ÇíÓÊÇå åãÎæÇäí äÏÇÑÏ", 1000
    End If
End Sub
Private Sub label13_Click(index As Integer)
    If Part(4) = clsStation.PartitionID Or clsStation.OtherPartition = False Then
        If CheckTableState(Label13(index).Tag) = True Then
            mvarTable = Val(Label13(index).Tag)
        Else
            If clsArya.HardLockSerialNo = "92072202873" Then
                ShowMessage "ÂíÇ ãí ÎæÇåíÏ ãíÒ ãæÑÏ äÙÑ ÎÇáí ÔæÏ ¿", True, True, "Èáí", "ÎíÑ"
                If mvarMsgIdx = vbYes Then
                    mvarTable = Val(Label13(index).Tag)
                    EmptyTable mvarTable
                    Unload Me
                    Exit Sub
                Else
                End If
            End If
            mvarTable = 0
            FindInvoice Label13(index).Tag, CurrentBranch
        End If
        Unload Me
    Else
        ShowDisMessage "ÈÎÔ ãæÑÏ äÙÑ ÈÇ ÈÎÔ Çíä ÇíÓÊÇå åãÎæÇäí äÏÇÑÏ", 1000
    End If
End Sub
Private Sub label14_Click(index As Integer)
    If Part(4) = clsStation.PartitionID Or clsStation.OtherPartition = False Then
        If CheckTableState(Label14(index).Tag) = True Then
            mvarTable = Val(Label14(index).Tag)
        Else
            If clsArya.HardLockSerialNo = "92072202873" Then
                ShowMessage "ÂíÇ ãí ÎæÇåíÏ ãíÒ ãæÑÏ äÙÑ ÎÇáí ÔæÏ ¿", True, True, "Èáí", "ÎíÑ"
                If mvarMsgIdx = vbYes Then
                    mvarTable = Val(Label14(index).Tag)
                    EmptyTable mvarTable
                    Unload Me
                    Exit Sub
                Else
                End If
            End If
            mvarTable = 0
            FindInvoice Label14(index).Tag, CurrentBranch
        End If
        Unload Me
    Else
        ShowDisMessage "ÈÎÔ ãæÑÏ äÙÑ ÈÇ ÈÎÔ Çíä ÇíÓÊÇå åãÎæÇäí äÏÇÑÏ", 1000
    End If
End Sub
Private Sub label15_Click(index As Integer)
    If Part(4) = clsStation.PartitionID Or clsStation.OtherPartition = False Then
        If CheckTableState(Label15(index).Tag) = True Then
            mvarTable = Val(Label15(index).Tag)
        Else
            If clsArya.HardLockSerialNo = "92072202873" Then
                ShowMessage "ÂíÇ ãí ÎæÇåíÏ ãíÒ ãæÑÏ äÙÑ ÎÇáí ÔæÏ ¿", True, True, "Èáí", "ÎíÑ"
                If mvarMsgIdx = vbYes Then
                    mvarTable = Val(Label15(index).Tag)
                    EmptyTable mvarTable
                    Unload Me
                    Exit Sub
                Else
                End If
            End If
            mvarTable = 0
            FindInvoice Label15(index).Tag, CurrentBranch
        End If
        Unload Me
    Else
        ShowDisMessage "ÈÎÔ ãæÑÏ äÙÑ ÈÇ ÈÎÔ Çíä ÇíÓÊÇå åãÎæÇäí äÏÇÑÏ", 1000
    End If
End Sub
Private Sub label16_Click(index As Integer)
    If Part(4) = clsStation.PartitionID Or clsStation.OtherPartition = False Then
        If CheckTableState(Label16(index).Tag) = True Then
            mvarTable = Val(Label16(index).Tag)
        Else
            If clsArya.HardLockSerialNo = "92072202873" Then
                ShowMessage "ÂíÇ ãí ÎæÇåíÏ ãíÒ ãæÑÏ äÙÑ ÎÇáí ÔæÏ ¿", True, True, "Èáí", "ÎíÑ"
                If mvarMsgIdx = vbYes Then
                    mvarTable = Val(Label16(index).Tag)
                    EmptyTable mvarTable
                    Unload Me
                    Exit Sub
                Else
                End If
            End If
            mvarTable = 0
            FindInvoice Label16(index).Tag, CurrentBranch
        End If
        Unload Me
    Else
        ShowDisMessage "ÈÎÔ ãæÑÏ äÙÑ ÈÇ ÈÎÔ Çíä ÇíÓÊÇå åãÎæÇäí äÏÇÑÏ", 1000
    End If
End Sub
Private Sub label17_Click(index As Integer)
    If Part(4) = clsStation.PartitionID Or clsStation.OtherPartition = False Then
        If CheckTableState(Label17(index).Tag) = True Then
            mvarTable = Val(Label17(index).Tag)
        Else
            If clsArya.HardLockSerialNo = "92072202873" Then
                ShowMessage "ÂíÇ ãí ÎæÇåíÏ ãíÒ ãæÑÏ äÙÑ ÎÇáí ÔæÏ ¿", True, True, "Èáí", "ÎíÑ"
                If mvarMsgIdx = vbYes Then
                    mvarTable = Val(Label17(index).Tag)
                    EmptyTable mvarTable
                    Unload Me
                    Exit Sub
                Else
                End If
            End If
            mvarTable = 0
            FindInvoice Label17(index).Tag, CurrentBranch
        End If
        Unload Me
    Else
        ShowDisMessage "ÈÎÔ ãæÑÏ äÙÑ ÈÇ ÈÎÔ Çíä ÇíÓÊÇå åãÎæÇäí äÏÇÑÏ", 1000
    End If
End Sub
Private Sub label18_Click(index As Integer)
    If Part(4) = clsStation.PartitionID Or clsStation.OtherPartition = False Then
        If CheckTableState(Label18(index).Tag) = True Then
            mvarTable = Val(Label18(index).Tag)
        Else
            If clsArya.HardLockSerialNo = "92072202873" Then
                ShowMessage "ÂíÇ ãí ÎæÇåíÏ ãíÒ ãæÑÏ äÙÑ ÎÇáí ÔæÏ ¿", True, True, "Èáí", "ÎíÑ"
                If mvarMsgIdx = vbYes Then
                    mvarTable = Val(Label18(index).Tag)
                    EmptyTable mvarTable
                    Unload Me
                    Exit Sub
                Else
                End If
            End If
            mvarTable = 0
            FindInvoice Label18(index).Tag, CurrentBranch
        End If
        Unload Me
    Else
        ShowDisMessage "ÈÎÔ ãæÑÏ äÙÑ ÈÇ ÈÎÔ Çíä ÇíÓÊÇå åãÎæÇäí äÏÇÑÏ", 1000
    End If
End Sub
Private Sub Label19_Click(index As Integer)
    If Part(4) = clsStation.PartitionID Or clsStation.OtherPartition = False Then
        If CheckTableState(Label19(index).Tag) = True Then
            mvarTable = Val(Label19(index).Tag)
        Else
            If clsArya.HardLockSerialNo = "92072202873" Then
                ShowMessage "ÂíÇ ãí ÎæÇåíÏ ãíÒ ãæÑÏ äÙÑ ÎÇáí ÔæÏ ¿", True, True, "Èáí", "ÎíÑ"
                If mvarMsgIdx = vbYes Then
                    mvarTable = Val(Label19(index).Tag)
                    EmptyTable mvarTable
                    Unload Me
                    Exit Sub
                Else
                End If
            End If
            mvarTable = 0
            FindInvoice Label19(index).Tag, CurrentBranch
        End If
        Unload Me
    Else
        ShowDisMessage "ÈÎÔ ãæÑÏ äÙÑ ÈÇ ÈÎÔ Çíä ÇíÓÊÇå åãÎæÇäí äÏÇÑÏ", 1000
    End If
End Sub

Private Sub Form_Load()
    On Error GoTo ErrHandler
    
    formloadFlag = False
    
    If Val(GetSetting(strMainKey, Me.Name, "TimerInterval")) > 0 Then
        Timer1.Interval = Val(GetSetting(strMainKey, Me.Name, "TimerInterval"))
    Else
         Timer1.Interval = 5000
    End If
    If Val(GetSetting(strMainKey, Me.Name, "TxtWidth")) > 0 Then
        TxtWidth = Val(GetSetting(strMainKey, Me.Name, "TxtWidth"))
    Else
         TxtWidth = 1600
    End If
    If Val(GetSetting(strMainKey, Me.Name, "TxtHeight")) > 0 Then
        TxtHeight = Val(GetSetting(strMainKey, Me.Name, "TxtHeight"))
    Else
         TxtHeight = 850
    End If
    If Val(GetSetting(strMainKey, Me.Name, "TxtFontSize")) > 0 Then
        TxtFontSize = Val(GetSetting(strMainKey, Me.Name, "TxtFontSize"))
    Else
         TxtFontSize = 12
    End If
    txtInterval.Value = CStr(Timer1.Interval / 1000)
    
    For i = 0 To 19
        SSTab1.TabVisible(i) = False
    Next
    
    Dim Rst As New ADODB.Recordset
    
    ReDim Parameter(1) As Parameter
    Parameter(0) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
    Parameter(1) = GenerateInputParameter("@Branch", adInteger, 4, CurrentBranch)
    Set Rst = RunParametricStoredProcedure2Rec("Get_All_tPartitions", Parameter)
    
    If Not (Rst.EOF = True And Rst.BOF = True) Then
        i = 0
        While Rst.EOF <> True
            SSTab1.TabVisible(i) = True
            SSTab1.TabCaption(i) = Rst.Fields("PartitionDescription").Value
            Part(i) = Rst.Fields("PartitionID").Value
            i = i + 1
            Rst.MoveNext
        Wend
    End If
    

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

Exit Sub
ErrHandler:
    formloadFlag = True
    ShowDisMessage err.Description, 2000
End Sub
Private Sub GetCount(Branch As Integer)
    On Error GoTo ErrHandler
        ReDim Parameter(0)
        Parameter(0) = GenerateInputParameter("@Branch", adInteger, 4, Branch)
        
        Set Rst = RunParametricStoredProcedure2Rec("GetTableCountByBranch", Parameter)
        
        If Not (Rst.EOF And Rst.BOF) Then
            Do While Rst.EOF = False
                tablesCount = Rst!ct
                Rst.MoveNext
            Loop
        End If
        
    Exit Sub
ErrHandler:
    MsgBox err.Description
    LogSave "frmTables", err, "GetCount"
End Sub



Private Sub GetTables(Branch As Integer)
    On Error GoTo ErrHandler
    
    For i = 1 To Label.Count - 1
        Unload Label(i)
    Next
    For i = 1 To Label1.Count - 1
        Unload Label1(i)
    Next
    For i = 1 To Label2.Count - 1
        Unload Label2(i)
    Next
    For i = 1 To Label3.Count - 1
        Unload Label3(i)
    Next
    For i = 1 To Label4.Count - 1
        Unload Label4(i)
    Next
    For i = 1 To Label5.Count - 1
        Unload Label5(i)
    Next
    For i = 1 To Label6.Count - 1
        Unload Label6(i)
    Next
    For i = 1 To Label7.Count - 1
        Unload Label7(i)
    Next
    For i = 1 To Label8.Count - 1
        Unload Label8(i)
    Next
    For i = 1 To Label9.Count - 1
        Unload Label9(i)
    Next
    For i = 1 To Label10.Count - 1
        Unload Label10(i)
    Next
    For i = 1 To Label11.Count - 1
        Unload Label11(i)
    Next
    For i = 1 To Label12.Count - 1
        Unload Label12(i)
    Next
    For i = 1 To Label13.Count - 1
        Unload Label13(i)
    Next
    For i = 1 To Label14.Count - 1
        Unload Label14(i)
    Next
    For i = 1 To Label15.Count - 1
        Unload Label15(i)
    Next
    For i = 1 To Label16.Count - 1
        Unload Label16(i)
    Next
    For i = 1 To Label17.Count - 1
        Unload Label17(i)
    Next
    For i = 1 To Label18.Count - 1
        Unload Label18(i)
    Next
    For i = 1 To Label19.Count - 1
        Unload Label19(i)
    Next
    
    Label(0).Width = TxtWidth
    Label(0).Height = TxtHeight
    Label(0).Font.Size = TxtFontSize
    Label1(0).Width = TxtWidth
    Label1(0).Height = TxtHeight
    Label1(0).Font.Size = TxtFontSize
    Label2(0).Width = TxtWidth
    Label2(0).Height = TxtHeight
    Label2(0).Font.Size = TxtFontSize
    Label3(0).Width = TxtWidth
    Label3(0).Height = TxtHeight
    Label3(0).Font.Size = TxtFontSize
    Label4(0).Width = TxtWidth
    Label4(0).Height = TxtHeight
    Label4(0).Font.Size = TxtFontSize
    Label5(0).Width = TxtWidth
    Label5(0).Height = TxtHeight
    Label5(0).Font.Size = TxtFontSize
    Label6(0).Width = TxtWidth
    Label6(0).Height = TxtHeight
    Label6(0).Font.Size = TxtFontSize
    Label7(0).Width = TxtWidth
    Label7(0).Height = TxtHeight
    Label7(0).Font.Size = TxtFontSize
    Label8(0).Width = TxtWidth
    Label8(0).Height = TxtHeight
    Label8(0).Font.Size = TxtFontSize
    Label9(0).Width = TxtWidth
    Label9(0).Height = TxtHeight
    Label9(0).Font.Size = TxtFontSize
    Label10(0).Width = TxtWidth
    Label10(0).Height = TxtHeight
    Label10(0).Font.Size = TxtFontSize
    Label11(0).Width = TxtWidth
    Label11(0).Height = TxtHeight
    Label11(0).Font.Size = TxtFontSize
    Label12(0).Width = TxtWidth
    Label12(0).Height = TxtHeight
    Label12(0).Font.Size = TxtFontSize
    Label13(0).Width = TxtWidth
    Label13(0).Height = TxtHeight
    Label13(0).Font.Size = TxtFontSize
    Label14(0).Width = TxtWidth
    Label14(0).Height = TxtHeight
    Label14(0).Font.Size = TxtFontSize
    Label15(0).Width = TxtWidth
    Label15(0).Height = TxtHeight
    Label15(0).Font.Size = TxtFontSize
    Label16(0).Width = TxtWidth
    Label16(0).Height = TxtHeight
    Label16(0).Font.Size = TxtFontSize
    Label17(0).Width = TxtWidth
    Label17(0).Height = TxtHeight
    Label17(0).Font.Size = TxtFontSize
    Label18(0).Width = TxtWidth
    Label18(0).Height = TxtHeight
    Label18(0).Font.Size = TxtFontSize
    Label19(0).Width = TxtWidth
    Label19(0).Height = TxtHeight
    Label19(0).Font.Size = TxtFontSize
    
    For i = 0 To 19
      lastPosition(i).X = 360
      lastPosition(i).Y = 360
    Next

'    Label(0).Height = 1000
'    Label(1).Height = 1000
'    Label(2).Height = 1000
'    Label(3).Height = 1000
'    Label(4).Height = 1000
    
    ReDim Parameter(1)
    Parameter(0) = GenerateInputParameter("@Branch", adInteger, 4, Branch)
    Parameter(1) = GenerateInputParameter("@TableControl", adInteger, 4, 0)
    
    Set Rst = RunParametricStoredProcedure2Rec("RetriveTable_Branch", Parameter)
    i = 1
    i0 = 1
    i1 = 1
    i2 = 1
    i3 = 1
    i4 = 1
    i5 = 1
    i6 = 1
    i7 = 1
    i8 = 1
    i9 = 1
    i10 = 1
    i11 = 1
    i12 = 1
    i13 = 1
    i14 = 1
    i15 = 1
    i16 = 1
    i17 = 1
    i18 = 1
    i19 = 1
    If Not (Rst.EOF And Rst.BOF) Then
        Do While Rst.EOF = False
            If Not (chkViewSpecial.Value = 1 And Val(frmInvoice.TxtGuestNo) <> Rst!NumberOfChair) Then
                Select Case Rst!PartitionID
                    Case Part(0)
                        If lastPosition(0).X > Frame1.Width - (Val(TxtWidth) + 105) Then
                            lastPosition(0).X = 360
                            lastPosition(0).Y = lastPosition(0).Y + Val(TxtHeight) + 105
                        End If
                        Load Label(i0)
                        Label(i0).Left = lastPosition(0).X
                        Label(i0).Top = lastPosition(0).Y
                        Label(i0).Visible = True
                        If clsArya.HardLockSerialNo = "92072202873" Then
                            Label(i0).Caption = Rst!TableDescription & " _ äÝÑ" & Rst!NumberOfChair
                        Else
                            Label(i0).Caption = Rst!TableDescription
                        End If
                        Label(i0).Tag = Rst!No
                        Label(i0).Width = Val(TxtWidth)
                        Label(i0).Height = Val(TxtHeight)
                        Label(i0).BackColor = IIf(Rst!Empty = False, &HB0FF&, &HC000&)                           'Not empty
                        'Label(i).Enabled = False
                        lastPosition(0).X = lastPosition(0).X + Val(TxtWidth) + 105
                        i0 = i0 + 1
                    Case Part(1)
                        If lastPosition(1).X > Frame2.Width - (Val(TxtWidth) + 105) Then
                            lastPosition(1).X = 360
                            lastPosition(1).Y = lastPosition(1).Y + Val(TxtHeight) + 105
                        End If
                        Load Label1(i1)
                        Label1(i1).Left = lastPosition(1).X
                        Label1(i1).Top = lastPosition(1).Y
                        Label1(i1).Visible = True
                        If clsArya.HardLockSerialNo = "92072202873" Then
                            Label1(i1).Caption = Rst!TableDescription & " _ äÝÑ" & Rst!NumberOfChair
                        Else
                            Label1(i1).Caption = Rst!TableDescription
                        End If
                        Label1(i1).Tag = Rst!No
                        Label1(i1).BackColor = IIf(Rst!Empty = False, &HB0FF&, &HC000&)                   'Not empty
                        'Label(i).Enabled = False
                        lastPosition(1).X = lastPosition(1).X + Val(TxtWidth) + 105
                        i1 = i1 + 1
                    Case Part(2)
                        If lastPosition(2).X > Frame3.Width - (Val(TxtWidth) + 105) Then
                            lastPosition(2).X = 360
                            lastPosition(2).Y = lastPosition(2).Y + Val(TxtHeight) + 105
                        End If
                        Load Label2(i2)
                        Label2(i2).Left = lastPosition(2).X
                        Label2(i2).Top = lastPosition(2).Y
                        Label2(i2).Visible = True
                        If clsArya.HardLockSerialNo = "92072202873" Then
                            Label2(i2).Caption = Rst!TableDescription & " _ äÝÑ" & Rst!NumberOfChair
                        Else
                            Label2(i2).Caption = Rst!TableDescription
                        End If
                        Label2(i2).Tag = Rst!No
                        Label2(i2).BackColor = IIf(Rst!Empty = False, &HB0FF&, &HC000&)                    'Not empty
                        'Label(i).Enabled = False
                        lastPosition(2).X = lastPosition(2).X + Val(TxtWidth) + 105
                        i2 = i2 + 1
                    Case Part(3)
                        If lastPosition(3).X > Frame4.Width - (Val(TxtWidth) + 105) Then
                            lastPosition(3).X = 360
                            lastPosition(3).Y = lastPosition(3).Y + Val(TxtHeight) + 105
                        End If
                        Load Label3(i3)
                        Label3(i3).Left = lastPosition(3).X
                        Label3(i3).Top = lastPosition(3).Y
                        Label3(i3).Visible = True
                        If clsArya.HardLockSerialNo = "92072202873" Then
                            Label3(i3).Caption = Rst!TableDescription & " _ äÝÑ" & Rst!NumberOfChair
                        Else
                            Label3(i3).Caption = Rst!TableDescription
                        End If
                        Label3(i3).Tag = Rst!No
                        Label3(i3).BackColor = IIf(Rst!Empty = False, &HB0FF&, &HC000&)                    'Not empty
                        'Label(i).Enabled = False
                        lastPosition(3).X = lastPosition(3).X + Val(TxtWidth) + 105
                        i3 = i3 + 1
                     Case Part(4)
                        If lastPosition(4).X > Frame5.Width - (Val(TxtWidth) + 105) Then
                            lastPosition(4).X = 360
                            lastPosition(4).Y = lastPosition(4).Y + Val(TxtHeight) + 105
                        End If
                        Load Label4(i4)
                        Label4(i4).Left = lastPosition(4).X
                        Label4(i4).Top = lastPosition(4).Y
                        Label4(i4).Visible = True
                        If clsArya.HardLockSerialNo = "92072202873" Then
                            Label4(i4).Caption = Rst!TableDescription & " _ äÝÑ" & Rst!NumberOfChair
                        Else
                            Label4(i4).Caption = Rst!TableDescription
                        End If
                        Label4(i4).Tag = Rst!No
                        Label4(i4).BackColor = IIf(Rst!Empty = False, &HB0FF&, &HC000&)                    'Not empty
                        'Label(i).Enabled = False
                        lastPosition(4).X = lastPosition(4).X + Val(TxtWidth) + 105
                        i4 = i4 + 1
                     Case Part(5)
                        If lastPosition(5).X > Frame5.Width - (Val(TxtWidth) + 105) Then
                            lastPosition(5).X = 360
                            lastPosition(5).Y = lastPosition(5).Y + Val(TxtHeight) + 105
                        End If
                        Load Label5(i5)
                        Label5(i5).Left = lastPosition(5).X
                        Label5(i5).Top = lastPosition(5).Y
                        Label5(i5).Visible = True
                        If clsArya.HardLockSerialNo = "92072202873" Then
                            Label5(i5).Caption = Rst!TableDescription & " _ äÝÑ" & Rst!NumberOfChair
                        Else
                            Label5(i5).Caption = Rst!TableDescription
                        End If
                        Label5(i5).Tag = Rst!No
                        Label5(i5).BackColor = IIf(Rst!Empty = False, &HB0FF&, &HC000&)                    'Not empty
                        'Label(i).Enabled = False
                        lastPosition(5).X = lastPosition(5).X + Val(TxtWidth) + 105
                        i5 = i5 + 1
                     Case Part(6)
                        If lastPosition(6).X > Frame5.Width - (Val(TxtWidth) + 105) Then
                            lastPosition(6).X = 360
                            lastPosition(6).Y = lastPosition(6).Y + Val(TxtHeight) + 105
                        End If
                        Load Label6(i6)
                        Label6(i6).Left = lastPosition(6).X
                        Label6(i6).Top = lastPosition(6).Y
                        Label6(i6).Visible = True
                        If clsArya.HardLockSerialNo = "92072202873" Then
                            Label6(i6).Caption = Rst!TableDescription & " _ äÝÑ" & Rst!NumberOfChair
                        Else
                            Label6(i6).Caption = Rst!TableDescription
                        End If
                        Label6(i6).Tag = Rst!No
                        Label6(i6).BackColor = IIf(Rst!Empty = False, &HB0FF&, &HC000&)                    'Not empty
                        'Label(i).Enabled = False
                        lastPosition(6).X = lastPosition(6).X + Val(TxtWidth) + 105
                        i6 = i6 + 1
                     Case Part(7)
                        If lastPosition(7).X > Frame5.Width - (Val(TxtWidth) + 105) Then
                            lastPosition(7).X = 360
                            lastPosition(7).Y = lastPosition(7).Y + Val(TxtHeight) + 105
                        End If
                        Load Label7(i7)
                        Label7(i7).Left = lastPosition(7).X
                        Label7(i7).Top = lastPosition(7).Y
                        Label7(i7).Visible = True
                        If clsArya.HardLockSerialNo = "92072202873" Then
                            Label7(i7).Caption = Rst!TableDescription & " _ äÝÑ" & Rst!NumberOfChair
                        Else
                            Label7(i7).Caption = Rst!TableDescription
                        End If
                        Label7(i7).Tag = Rst!No
                        Label7(i7).BackColor = IIf(Rst!Empty = False, &HB0FF&, &HC000&)                    'Not empty
                        'Label(i).Enabled = False
                        lastPosition(7).X = lastPosition(7).X + Val(TxtWidth) + 105
                        i7 = i7 + 1
                     Case Part(8)
                        If lastPosition(8).X > Frame5.Width - (Val(TxtWidth) + 105) Then
                            lastPosition(8).X = 360
                            lastPosition(8).Y = lastPosition(8).Y + Val(TxtHeight) + 105
                        End If
                        Load Label8(i8)
                        Label8(i8).Left = lastPosition(8).X
                        Label8(i8).Top = lastPosition(8).Y
                        Label8(i8).Visible = True
                        If clsArya.HardLockSerialNo = "92072202873" Then
                            Label8(i8).Caption = Rst!TableDescription & " _ äÝÑ" & Rst!NumberOfChair
                        Else
                            Label8(i8).Caption = Rst!TableDescription
                        End If
                        Label8(i8).Tag = Rst!No
                        Label8(i8).BackColor = IIf(Rst!Empty = False, &HB0FF&, &HC000&)                    'Not empty
                        'Label(i).Enabled = False
                        lastPosition(8).X = lastPosition(8).X + Val(TxtWidth) + 105
                        i8 = i8 + 1
                     Case Part(9)
                        If lastPosition(9).X > Frame5.Width - (Val(TxtWidth) + 105) Then
                            lastPosition(9).X = 360
                            lastPosition(9).Y = lastPosition(9).Y + Val(TxtHeight) + 105
                        End If
                        Load Label9(i9)
                        Label9(i9).Left = lastPosition(9).X
                        Label9(i9).Top = lastPosition(9).Y
                        Label9(i9).Visible = True
                        If clsArya.HardLockSerialNo = "92072202873" Then
                            Label9(i9).Caption = Rst!TableDescription & " _ äÝÑ" & Rst!NumberOfChair
                        Else
                            Label9(i9).Caption = Rst!TableDescription
                        End If
                        Label9(i9).Tag = Rst!No
                        Label9(i9).BackColor = IIf(Rst!Empty = False, &HB0FF&, &HC000&)                    'Not empty
                        'Label(i).Enabled = False
                        lastPosition(9).X = lastPosition(9).X + Val(TxtWidth) + 105
                        i9 = i9 + 1
                     Case Part(10)
                        If lastPosition(10).X > Frame5.Width - (Val(TxtWidth) + 105) Then
                            lastPosition(10).X = 360
                            lastPosition(10).Y = lastPosition(10).Y + Val(TxtHeight) + 105
                        End If
                        Load Label10(i10)
                        Label10(i10).Left = lastPosition(10).X
                        Label10(i10).Top = lastPosition(10).Y
                        Label10(i10).Visible = True
                        If clsArya.HardLockSerialNo = "92072202873" Then
                            Label10(i10).Caption = Rst!TableDescription & " _ äÝÑ" & Rst!NumberOfChair
                        Else
                            Label10(i10).Caption = Rst!TableDescription
                        End If
                        Label10(i10).Tag = Rst!No
                        Label10(i10).BackColor = IIf(Rst!Empty = False, &HB0FF&, &HC000&)                    'Not empty
                        'Label(i).Enabled = False
                        lastPosition(10).X = lastPosition(10).X + Val(TxtWidth) + 105
                        i10 = i10 + 1
                     Case Part(11)
                        If lastPosition(11).X > Frame5.Width - (Val(TxtWidth) + 105) Then
                            lastPosition(11).X = 360
                            lastPosition(11).Y = lastPosition(11).Y + Val(TxtHeight) + 105
                        End If
                        Load Label11(i11)
                        Label11(i11).Left = lastPosition(11).X
                        Label11(i11).Top = lastPosition(11).Y
                        Label11(i11).Visible = True
                        If clsArya.HardLockSerialNo = "92072202873" Then
                            Label11(i11).Caption = Rst!TableDescription & " _ äÝÑ" & Rst!NumberOfChair
                        Else
                            Label11(i11).Caption = Rst!TableDescription
                        End If
                        Label11(i11).Tag = Rst!No
                        Label11(i11).BackColor = IIf(Rst!Empty = False, &HB0FF&, &HC000&)                    'Not empty
                        'Label(i).Enabled = False
                        lastPosition(11).X = lastPosition(11).X + Val(TxtWidth) + 105
                        i11 = i11 + 1
                     Case Part(12)
                        If lastPosition(12).X > Frame5.Width - (Val(TxtWidth) + 105) Then
                            lastPosition(12).X = 360
                            lastPosition(12).Y = lastPosition(12).Y + Val(TxtHeight) + 105
                        End If
                        Load Label12(i12)
                        Label12(i12).Left = lastPosition(12).X
                        Label12(i12).Top = lastPosition(12).Y
                        Label12(i12).Visible = True
                        If clsArya.HardLockSerialNo = "92072202873" Then
                            Label12(i12).Caption = Rst!TableDescription & " _ äÝÑ" & Rst!NumberOfChair
                        Else
                            Label12(i12).Caption = Rst!TableDescription
                        End If
                        Label12(i12).Tag = Rst!No
                        Label12(i12).BackColor = IIf(Rst!Empty = False, &HB0FF&, &HC000&)                    'Not empty
                        'Label(i).Enabled = False
                        lastPosition(12).X = lastPosition(12).X + Val(TxtWidth) + 105
                        i12 = i12 + 1
                     Case Part(13)
                        If lastPosition(13).X > Frame5.Width - (Val(TxtWidth) + 105) Then
                            lastPosition(13).X = 360
                            lastPosition(13).Y = lastPosition(13).Y + Val(TxtHeight) + 105
                        End If
                        Load Label13(i13)
                        Label13(i13).Left = lastPosition(13).X
                        Label13(i13).Top = lastPosition(13).Y
                        Label13(i13).Visible = True
                        If clsArya.HardLockSerialNo = "92072202873" Then
                            Label13(i13).Caption = Rst!TableDescription & " _ äÝÑ" & Rst!NumberOfChair
                        Else
                            Label13(i13).Caption = Rst!TableDescription
                        End If
                        Label13(i13).Tag = Rst!No
                        Label13(i13).BackColor = IIf(Rst!Empty = False, &HB0FF&, &HC000&)                    'Not empty
                        'Label(i).Enabled = False
                        lastPosition(13).X = lastPosition(13).X + Val(TxtWidth) + 105
                        i13 = i13 + 1
                     Case Part(14)
                        If lastPosition(14).X > Frame5.Width - (Val(TxtWidth) + 105) Then
                            lastPosition(14).X = 360
                            lastPosition(14).Y = lastPosition(14).Y + Val(TxtHeight) + 105
                        End If
                        Load Label14(i14)
                        Label14(i14).Left = lastPosition(14).X
                        Label14(i14).Top = lastPosition(14).Y
                        Label14(i14).Visible = True
                        If clsArya.HardLockSerialNo = "92072202873" Then
                            Label14(i14).Caption = Rst!TableDescription & " _ äÝÑ" & Rst!NumberOfChair
                        Else
                            Label14(i14).Caption = Rst!TableDescription
                        End If
                        Label14(i14).Tag = Rst!No
                        Label14(i14).BackColor = IIf(Rst!Empty = False, &HB0FF&, &HC000&)                    'Not empty
                        'Label(i).Enabled = False
                        lastPosition(14).X = lastPosition(14).X + Val(TxtWidth) + 105
                        i14 = i14 + 1
                     Case Part(15)
                        If lastPosition(15).X > Frame5.Width - (Val(TxtWidth) + 105) Then
                            lastPosition(15).X = 360
                            lastPosition(15).Y = lastPosition(15).Y + Val(TxtHeight) + 105
                        End If
                        Load Label15(i15)
                        Label15(i15).Left = lastPosition(15).X
                        Label15(i15).Top = lastPosition(15).Y
                        Label15(i15).Visible = True
                        If clsArya.HardLockSerialNo = "92072202873" Then
                            Label15(i15).Caption = Rst!TableDescription & " _ äÝÑ" & Rst!NumberOfChair
                        Else
                            Label15(i15).Caption = Rst!TableDescription
                        End If
                        Label15(i15).Tag = Rst!No
                        Label15(i15).BackColor = IIf(Rst!Empty = False, &HB0FF&, &HC000&)                    'Not empty
                        'Label(i).Enabled = False
                        lastPosition(15).X = lastPosition(15).X + Val(TxtWidth) + 105
                        i15 = i15 + 1
                     Case Part(16)
                        If lastPosition(16).X > Frame5.Width - (Val(TxtWidth) + 105) Then
                            lastPosition(16).X = 360
                            lastPosition(16).Y = lastPosition(16).Y + Val(TxtHeight) + 105
                        End If
                        Load Label16(i16)
                        Label16(i16).Left = lastPosition(16).X
                        Label16(i16).Top = lastPosition(16).Y
                        Label16(i16).Visible = True
                        If clsArya.HardLockSerialNo = "92072202873" Then
                            Label16(i16).Caption = Rst!TableDescription & " _ äÝÑ" & Rst!NumberOfChair
                        Else
                            Label16(i16).Caption = Rst!TableDescription
                        End If
                        Label16(i16).Tag = Rst!No
                        Label16(i16).BackColor = IIf(Rst!Empty = False, &HB0FF&, &HC000&)                    'Not empty
                        'Label(i).Enabled = False
                        lastPosition(16).X = lastPosition(16).X + Val(TxtWidth) + 105
                        i16 = i16 + 1
                     Case Part(17)
                        If lastPosition(17).X > Frame5.Width - (Val(TxtWidth) + 105) Then
                            lastPosition(17).X = 360
                            lastPosition(17).Y = lastPosition(17).Y + Val(TxtHeight) + 105
                        End If
                        Load Label17(i17)
                        Label17(i17).Left = lastPosition(17).X
                        Label17(i17).Top = lastPosition(17).Y
                        Label17(i17).Visible = True
                        If clsArya.HardLockSerialNo = "92072202873" Then
                            Label17(i17).Caption = Rst!TableDescription & " _ äÝÑ" & Rst!NumberOfChair
                        Else
                            Label17(i17).Caption = Rst!TableDescription
                        End If
                        Label17(i17).Tag = Rst!No
                        Label17(i17).BackColor = IIf(Rst!Empty = False, &HB0FF&, &HC000&)                    'Not empty
                        'Label(i).Enabled = False
                        lastPosition(17).X = lastPosition(17).X + Val(TxtWidth) + 105
                        i17 = i17 + 1
                     Case Part(18)
                        If lastPosition(18).X > Frame5.Width - (Val(TxtWidth) + 105) Then
                            lastPosition(18).X = 360
                            lastPosition(18).Y = lastPosition(18).Y + Val(TxtHeight) + 105
                        End If
                        Load Label18(i18)
                        Label18(i18).Left = lastPosition(18).X
                        Label18(i18).Top = lastPosition(18).Y
                        Label18(i18).Visible = True
                        If clsArya.HardLockSerialNo = "92072202873" Then
                            Label18(i18).Caption = Rst!TableDescription & " _ äÝÑ" & Rst!NumberOfChair
                        Else
                            Label18(i18).Caption = Rst!TableDescription
                        End If
                        Label18(i18).Tag = Rst!No
                        Label18(i18).BackColor = IIf(Rst!Empty = False, &HB0FF&, &HC000&)                    'Not empty
                        'Label(i).Enabled = False
                        lastPosition(18).X = lastPosition(18).X + Val(TxtWidth) + 105
                        i18 = i18 + 1
                     Case Part(19)
                        If lastPosition(19).X > Frame5.Width - (Val(TxtWidth) + 105) Then
                            lastPosition(19).X = 360
                            lastPosition(19).Y = lastPosition(19).Y + Val(TxtHeight) + 105
                        End If
                        Load Label19(i19)
                        Label19(i19).Left = lastPosition(19).X
                        Label19(i19).Top = lastPosition(19).Y
                        Label19(i19).Visible = True
                        If clsArya.HardLockSerialNo = "92072202873" Then
                            Label19(i19).Caption = Rst!TableDescription & " _ äÝÑ" & Rst!NumberOfChair
                        Else
                            Label19(i19).Caption = Rst!TableDescription
                        End If
                        Label19(i19).Tag = Rst!No
                        Label19(i19).BackColor = IIf(Rst!Empty = False, &HB0FF&, &HC000&)                    'Not empty
                        'Label(i).Enabled = False
                        lastPosition(19).X = lastPosition(19).X + Val(TxtWidth) + 105
                        i19 = i19 + 1
                End Select
            End If
            Rst.MoveNext
        Loop
    
    End If

    Exit Sub
ErrHandler:
   MsgBox err.Description
   LogSave "frmTables", err, "GetNumberTables"
End Sub

Private Sub DetectEmptyTables(Branch As Integer)
    On Error GoTo ErrHandler
    ReDim Parameter(1)
    Parameter(0) = GenerateInputParameter("@Branch", adInteger, 4, Branch)
    Parameter(1) = GenerateInputParameter("@TableControl", adInteger, 4, 0)
    
    Set Rst = RunParametricStoredProcedure2Rec("RetriveTable_Branch", Parameter)

    If Not (Rst.EOF And Rst.BOF) Then
        For Each varlbl In Label    'Enable the empty tables
            If varlbl.Tag = CStr(Rst!No) Then
                If Rst!Empty = False Then varlbl.BackColor = &HB0FF& Else varlbl.BackColor = &HC000& '': varlbl.Caption = Rst!TableDescription
                Rst.MoveNext
                If Rst.EOF Then Exit For
            End If
        Next
        If Rst.EOF = False Then
            For Each varlbl In Label1
                If varlbl.Tag = CStr(Rst!No) Then
                    If Rst!Empty = False Then varlbl.BackColor = &HB0FF& Else varlbl.BackColor = &HC000& '': varlbl.Caption = Rst!TableDescription
                    Rst.MoveNext
                    If Rst.EOF Then Exit For
                End If
            Next
        End If
        If Rst.EOF = False Then
             For Each varlbl In Label2
                 If varlbl.Tag = CStr(Rst!No) Then
                     If Rst!Empty = False Then varlbl.BackColor = &HB0FF& Else varlbl.BackColor = &HC000& '': varlbl.Caption = Rst!TableDescription
                     Rst.MoveNext
                     If Rst.EOF Then Exit For
                 End If
             Next
         End If
        If Rst.EOF = False Then
             For Each varlbl In Label3
                 If varlbl.Tag = CStr(Rst!No) Then
                     If Rst!Empty = False Then varlbl.BackColor = &HB0FF& Else varlbl.BackColor = &HC000& '': varlbl.Caption = Rst!TableDescription
                     Rst.MoveNext
                     If Rst.EOF Then Exit For
                 End If
             Next
        End If
        If Rst.EOF = False Then
             For Each varlbl In Label4
                 If varlbl.Tag = CStr(Rst!No) Then
                     If Rst!Empty = False Then varlbl.BackColor = &HB0FF& Else varlbl.BackColor = &HC000& '': varlbl.Caption = Rst!TableDescription
                     Rst.MoveNext
                     If Rst.EOF Then Exit For
                 End If
             Next
         End If
        If Rst.EOF = False Then
             For Each varlbl In Label5
                 If varlbl.Tag = CStr(Rst!No) Then
                     If Rst!Empty = False Then varlbl.BackColor = &HB0FF& Else varlbl.BackColor = &HC000& '': varlbl.Caption = Rst!TableDescription
                     Rst.MoveNext
                     If Rst.EOF Then Exit For
                 End If
             Next
         End If
        If Rst.EOF = False Then
             For Each varlbl In Label6
                 If varlbl.Tag = CStr(Rst!No) Then
                     If Rst!Empty = False Then varlbl.BackColor = &HB0FF& Else varlbl.BackColor = &HC000& '': varlbl.Caption = Rst!TableDescription
                     Rst.MoveNext
                     If Rst.EOF Then Exit For
                 End If
             Next
         End If
        If Rst.EOF = False Then
             For Each varlbl In Label7
                 If varlbl.Tag = CStr(Rst!No) Then
                     If Rst!Empty = False Then varlbl.BackColor = &HB0FF& Else varlbl.BackColor = &HC000& '': varlbl.Caption = Rst!TableDescription
                     Rst.MoveNext
                     If Rst.EOF Then Exit For
                 End If
             Next
         End If
        If Rst.EOF = False Then
             For Each varlbl In Label8
                 If varlbl.Tag = CStr(Rst!No) Then
                     If Rst!Empty = False Then varlbl.BackColor = &HB0FF& Else varlbl.BackColor = &HC000& '': varlbl.Caption = Rst!TableDescription
                     Rst.MoveNext
                     If Rst.EOF Then Exit For
                 End If
             Next
         End If
        If Rst.EOF = False Then
             For Each varlbl In Label9
                 If varlbl.Tag = CStr(Rst!No) Then
                     If Rst!Empty = False Then varlbl.BackColor = &HB0FF& Else varlbl.BackColor = &HC000& '': varlbl.Caption = Rst!TableDescription
                     Rst.MoveNext
                     If Rst.EOF Then Exit For
                 End If
             Next
         End If
        If Rst.EOF = False Then
             For Each varlbl In Label10
                 If varlbl.Tag = CStr(Rst!No) Then
                     If Rst!Empty = False Then varlbl.BackColor = &HB0FF& Else varlbl.BackColor = &HC000& '': varlbl.Caption = Rst!TableDescription
                     Rst.MoveNext
                     If Rst.EOF Then Exit For
                 End If
             Next
         End If
        If Rst.EOF = False Then
             For Each varlbl In Label11
                 If varlbl.Tag = CStr(Rst!No) Then
                     If Rst!Empty = False Then varlbl.BackColor = &HB0FF& Else varlbl.BackColor = &HC000& '': varlbl.Caption = Rst!TableDescription
                     Rst.MoveNext
                     If Rst.EOF Then Exit For
                 End If
             Next
         End If
        If Rst.EOF = False Then
             For Each varlbl In Label12
                 If varlbl.Tag = CStr(Rst!No) Then
                     If Rst!Empty = False Then varlbl.BackColor = &HB0FF& Else varlbl.BackColor = &HC000& '': varlbl.Caption = Rst!TableDescription
                     Rst.MoveNext
                     If Rst.EOF Then Exit For
                 End If
             Next
         End If
        If Rst.EOF = False Then
             For Each varlbl In Label13
                 If varlbl.Tag = CStr(Rst!No) Then
                     If Rst!Empty = False Then varlbl.BackColor = &HB0FF& Else varlbl.BackColor = &HC000& '': varlbl.Caption = Rst!TableDescription
                     Rst.MoveNext
                     If Rst.EOF Then Exit For
                 End If
             Next
         End If
        If Rst.EOF = False Then
             For Each varlbl In Label14
                 If varlbl.Tag = CStr(Rst!No) Then
                     If Rst!Empty = False Then varlbl.BackColor = &HB0FF& Else varlbl.BackColor = &HC000& '': varlbl.Caption = Rst!TableDescription
                     Rst.MoveNext
                     If Rst.EOF Then Exit For
                 End If
             Next
         End If
        If Rst.EOF = False Then
             For Each varlbl In Label15
                 If varlbl.Tag = CStr(Rst!No) Then
                     If Rst!Empty = False Then varlbl.BackColor = &HB0FF& Else varlbl.BackColor = &HC000& '': varlbl.Caption = Rst!TableDescription
                     Rst.MoveNext
                     If Rst.EOF Then Exit For
                 End If
             Next
         End If
        If Rst.EOF = False Then
             For Each varlbl In Label16
                 If varlbl.Tag = CStr(Rst!No) Then
                     If Rst!Empty = False Then varlbl.BackColor = &HB0FF& Else varlbl.BackColor = &HC000& '': varlbl.Caption = Rst!TableDescription
                     Rst.MoveNext
                     If Rst.EOF Then Exit For
                 End If
             Next
         End If
        If Rst.EOF = False Then
             For Each varlbl In Label17
                 If varlbl.Tag = CStr(Rst!No) Then
                     If Rst!Empty = False Then varlbl.BackColor = &HB0FF& Else varlbl.BackColor = &HC000& '': varlbl.Caption = Rst!TableDescription
                     Rst.MoveNext
                     If Rst.EOF Then Exit For
                 End If
             Next
         End If
        If Rst.EOF = False Then
             For Each varlbl In Label18
                 If varlbl.Tag = CStr(Rst!No) Then
                     If Rst!Empty = False Then varlbl.BackColor = &HB0FF& Else varlbl.BackColor = &HC000& '': varlbl.Caption = Rst!TableDescription
                     Rst.MoveNext
                     If Rst.EOF Then Exit For
                 End If
             Next
         End If
        If Rst.EOF = False Then
             For Each varlbl In Label19
                 If varlbl.Tag = CStr(Rst!No) Then
                     If Rst!Empty = False Then varlbl.BackColor = &HB0FF& Else varlbl.BackColor = &HC000& '': varlbl.Caption = Rst!TableDescription
                     Rst.MoveNext
                     If Rst.EOF Then Exit For
                 End If
             Next
         End If

    End If
'    SSTab1.Tab = 0
    Exit Sub
ErrHandler:
    MsgBox err.Description
    LogSave "frmTables", err, "DetectResevedTables"
End Sub

Private Sub DetectBusyTables(Branch As Integer)
    On Error GoTo ErrHandler
    ReDim Parameter(0)
    Parameter(0) = GenerateInputParameter("@Branch", adInteger, 4, Branch)

    Set Rst = RunParametricStoredProcedure2Rec("Get_tblSamar_TableUsage_BusyTable", Parameter)

    If Not (Rst.EOF And Rst.BOF) Then
        For Each varlbl In Label    'Enable the empty tables
             If varlbl.Tag = CStr(Rst!No) Then
                 varlbl.Caption = " : " & Rst!TableDescription & vbLf & Rst!MinuteUseDiff & "  ÏÞíÞå"            'green for empty
                 If Not IsNull(Rst!nvcMaxUseTime) Then
                    If Val(Rst!MinuteUseDiff) >= Val(Rst!nvcMaxUseTime) Then varlbl.BackColor = &HFFFF&: varlbl.Caption = varlbl.Caption & vbLf & "ÇÖÇÝí" & Val(Rst!MinuteUseDiff) - Val(Rst!nvcMaxUseTime) & "  ÏÞíÞå"
                    
                 End If
                 Rst.MoveNext
                 If Rst.EOF Then Exit For
            End If
        Next
        If Rst.EOF = False Then
            For Each varlbl In Label1
                 If varlbl.Tag = CStr(Rst!No) Then
                    varlbl.Caption = " : " & Rst!TableDescription & vbLf & " : " & Rst!MinuteUseDiff & "  ÏÞíÞå"             'green for empty
                    If Not IsNull(Rst!nvcMaxUseTime) Then
                       If Rst!MinuteUseDiff >= Val(Rst!nvcMaxUseTime) Then varlbl.BackColor = &HFFFF&: varlbl.Caption = varlbl.Caption & vbLf & "ÇÖÇÝí" & Val(Rst!MinuteUseDiff) - Val(Rst!nvcMaxUseTime) & "  ÏÞíÞå"
                    End If
                    Rst.MoveNext
                    If Rst.EOF Then Exit For
                End If
            Next
        End If
        If Rst.EOF = False Then
             For Each varlbl In Label2
                  If varlbl.Tag = CStr(Rst!No) Then
                     varlbl.Caption = " : " & Rst!TableDescription & vbLf & " : " & Rst!MinuteUseDiff & "  ÏÞíÞå"            'green for empty
                     If Not IsNull(Rst!nvcMaxUseTime) Then
                        If Rst!MinuteUseDiff >= Val(Rst!nvcMaxUseTime) Then varlbl.BackColor = &HFFFF&: varlbl.Caption = varlbl.Caption & vbLf & "ÇÖÇÝí" & Val(Rst!MinuteUseDiff) - Val(Rst!nvcMaxUseTime) & "  ÏÞíÞå"
                     End If
                     Rst.MoveNext
                     If Rst.EOF Then Exit For
                 End If
             Next
         End If
        If Rst.EOF = False Then
             For Each varlbl In Label3
                 If varlbl.Tag = CStr(Rst!No) Then
                     varlbl.Caption = " : " & Rst!TableDescription & vbLf & " : " & Rst!MinuteUseDiff & "  ÏÞíÞå"           'green for empty
                     If Not IsNull(Rst!nvcMaxUseTime) Then
                        If Rst!MinuteUseDiff >= Val(Rst!nvcMaxUseTime) Then varlbl.BackColor = &HFFFF&: varlbl.Caption = varlbl.Caption & vbLf & "ÇÖÇÝí" & Val(Rst!MinuteUseDiff) - Val(Rst!nvcMaxUseTime) & "  ÏÞíÞå"
                     End If
                     Rst.MoveNext
                     If Rst.EOF Then Exit For
                 End If
             Next
         End If
        If Rst.EOF = False Then
             For Each varlbl In Label4
                 If varlbl.Tag = CStr(Rst!No) Then
                     varlbl.Caption = " : " & Rst!TableDescription & vbLf & " : " & Rst!MinuteUseDiff & "  ÏÞíÞå"            'green for empty
                     If Not IsNull(Rst!nvcMaxUseTime) Then
                        If Rst!MinuteUseDiff >= Val(Rst!nvcMaxUseTime) Then varlbl.BackColor = &HFFFF&: varlbl.Caption = varlbl.Caption & vbLf & "ÇÖÇÝí" & Val(Rst!MinuteUseDiff) - Val(Rst!nvcMaxUseTime) & "  ÏÞíÞå"
                     End If
                     Rst.MoveNext
                     If Rst.EOF Then Exit For
                 End If
             Next
         End If
        If Rst.EOF = False Then
             For Each varlbl In Label5
                 If varlbl.Tag = CStr(Rst!No) Then
                     varlbl.Caption = " : " & Rst!TableDescription & vbLf & " : " & Rst!MinuteUseDiff & "  ÏÞíÞå"            'green for empty
                     If Not IsNull(Rst!nvcMaxUseTime) Then
                        If Rst!MinuteUseDiff >= Val(Rst!nvcMaxUseTime) Then varlbl.BackColor = &HFFFF&: varlbl.Caption = varlbl.Caption & vbLf & "ÇÖÇÝí" & Val(Rst!MinuteUseDiff) - Val(Rst!nvcMaxUseTime) & "  ÏÞíÞå"
                     End If
                     Rst.MoveNext
                     If Rst.EOF Then Exit For
                 End If
             Next
         End If
        If Rst.EOF = False Then
             For Each varlbl In Label6
                 If varlbl.Tag = CStr(Rst!No) Then
                     varlbl.Caption = " : " & Rst!TableDescription & vbLf & " : " & Rst!MinuteUseDiff & "  ÏÞíÞå"            'green for empty
                     If Not IsNull(Rst!nvcMaxUseTime) Then
                        If Rst!MinuteUseDiff >= Val(Rst!nvcMaxUseTime) Then varlbl.BackColor = &HFFFF&: varlbl.Caption = varlbl.Caption & vbLf & "ÇÖÇÝí" & Val(Rst!MinuteUseDiff) - Val(Rst!nvcMaxUseTime) & "  ÏÞíÞå"
                     End If
                     Rst.MoveNext
                     If Rst.EOF Then Exit For
                 End If
             Next
         End If
        If Rst.EOF = False Then
             For Each varlbl In Label7
                 If varlbl.Tag = CStr(Rst!No) Then
                     varlbl.Caption = " : " & Rst!TableDescription & vbLf & " : " & Rst!MinuteUseDiff & "  ÏÞíÞå"            'green for empty
                     If Not IsNull(Rst!nvcMaxUseTime) Then
                        If Rst!MinuteUseDiff >= Val(Rst!nvcMaxUseTime) Then varlbl.BackColor = &HFFFF&: varlbl.Caption = varlbl.Caption & vbLf & "ÇÖÇÝí" & Val(Rst!MinuteUseDiff) - Val(Rst!nvcMaxUseTime) & "  ÏÞíÞå"
                     End If
                     Rst.MoveNext
                     If Rst.EOF Then Exit For
                 End If
             Next
         End If
        If Rst.EOF = False Then
             For Each varlbl In Label8
                 If varlbl.Tag = CStr(Rst!No) Then
                     varlbl.Caption = " : " & Rst!TableDescription & vbLf & " : " & Rst!MinuteUseDiff & "  ÏÞíÞå"            'green for empty
                     If Not IsNull(Rst!nvcMaxUseTime) Then
                        If Rst!MinuteUseDiff >= Val(Rst!nvcMaxUseTime) Then varlbl.BackColor = &HFFFF&: varlbl.Caption = varlbl.Caption & vbLf & "ÇÖÇÝí" & Val(Rst!MinuteUseDiff) - Val(Rst!nvcMaxUseTime) & "  ÏÞíÞå"
                     End If
                     Rst.MoveNext
                     If Rst.EOF Then Exit For
                 End If
             Next
         End If
        If Rst.EOF = False Then
             For Each varlbl In Label9
                 If varlbl.Tag = CStr(Rst!No) Then
                     varlbl.Caption = " : " & Rst!TableDescription & vbLf & " : " & Rst!MinuteUseDiff & "  ÏÞíÞå"            'green for empty
                     If Not IsNull(Rst!nvcMaxUseTime) Then
                        If Rst!MinuteUseDiff >= Val(Rst!nvcMaxUseTime) Then varlbl.BackColor = &HFFFF&: varlbl.Caption = varlbl.Caption & vbLf & "ÇÖÇÝí" & Val(Rst!MinuteUseDiff) - Val(Rst!nvcMaxUseTime) & "  ÏÞíÞå"
                     End If
                     Rst.MoveNext
                     If Rst.EOF Then Exit For
                 End If
             Next
         End If
        If Rst.EOF = False Then
             For Each varlbl In Label10
                 If varlbl.Tag = CStr(Rst!No) Then
                     varlbl.Caption = " : " & Rst!TableDescription & vbLf & " : " & Rst!MinuteUseDiff & "  ÏÞíÞå"            'green for empty
                     If Not IsNull(Rst!nvcMaxUseTime) Then
                        If Rst!MinuteUseDiff >= Val(Rst!nvcMaxUseTime) Then varlbl.BackColor = &HFFFF&: varlbl.Caption = varlbl.Caption & vbLf & "ÇÖÇÝí" & Val(Rst!MinuteUseDiff) - Val(Rst!nvcMaxUseTime) & "  ÏÞíÞå"
                     End If
                     Rst.MoveNext
                     If Rst.EOF Then Exit For
                 End If
             Next
         End If
        If Rst.EOF = False Then
             For Each varlbl In Label11
                 If varlbl.Tag = CStr(Rst!No) Then
                     varlbl.Caption = " : " & Rst!TableDescription & vbLf & " : " & Rst!MinuteUseDiff & "  ÏÞíÞå"            'green for empty
                     If Not IsNull(Rst!nvcMaxUseTime) Then
                        If Rst!MinuteUseDiff >= Val(Rst!nvcMaxUseTime) Then varlbl.BackColor = &HFFFF&: varlbl.Caption = varlbl.Caption & vbLf & "ÇÖÇÝí" & Val(Rst!MinuteUseDiff) - Val(Rst!nvcMaxUseTime) & "  ÏÞíÞå"
                     End If
                     Rst.MoveNext
                     If Rst.EOF Then Exit For
                 End If
             Next
         End If
        If Rst.EOF = False Then
             For Each varlbl In Label12
                 If varlbl.Tag = CStr(Rst!No) Then
                     varlbl.Caption = " : " & Rst!TableDescription & vbLf & " : " & Rst!MinuteUseDiff & "  ÏÞíÞå"            'green for empty
                     If Not IsNull(Rst!nvcMaxUseTime) Then
                        If Rst!MinuteUseDiff >= Val(Rst!nvcMaxUseTime) Then varlbl.BackColor = &HFFFF&: varlbl.Caption = varlbl.Caption & vbLf & "ÇÖÇÝí" & Val(Rst!MinuteUseDiff) - Val(Rst!nvcMaxUseTime) & "  ÏÞíÞå"
                     End If
                     Rst.MoveNext
                     If Rst.EOF Then Exit For
                 End If
             Next
         End If
        If Rst.EOF = False Then
             For Each varlbl In Label13
                 If varlbl.Tag = CStr(Rst!No) Then
                     varlbl.Caption = " : " & Rst!TableDescription & vbLf & " : " & Rst!MinuteUseDiff & "  ÏÞíÞå"            'green for empty
                     If Not IsNull(Rst!nvcMaxUseTime) Then
                        If Rst!MinuteUseDiff >= Val(Rst!nvcMaxUseTime) Then varlbl.BackColor = &HFFFF&: varlbl.Caption = varlbl.Caption & vbLf & "ÇÖÇÝí" & Val(Rst!MinuteUseDiff) - Val(Rst!nvcMaxUseTime) & "  ÏÞíÞå"
                     End If
                     Rst.MoveNext
                     If Rst.EOF Then Exit For
                 End If
             Next
         End If
        If Rst.EOF = False Then
             For Each varlbl In Label14
                 If varlbl.Tag = CStr(Rst!No) Then
                     varlbl.Caption = " : " & Rst!TableDescription & vbLf & " : " & Rst!MinuteUseDiff & "  ÏÞíÞå"            'green for empty
                     If Not IsNull(Rst!nvcMaxUseTime) Then
                        If Rst!MinuteUseDiff >= Val(Rst!nvcMaxUseTime) Then varlbl.BackColor = &HFFFF&: varlbl.Caption = varlbl.Caption & vbLf & "ÇÖÇÝí" & Val(Rst!MinuteUseDiff) - Val(Rst!nvcMaxUseTime) & "  ÏÞíÞå"
                     End If
                     Rst.MoveNext
                     If Rst.EOF Then Exit For
                 End If
             Next
         End If
        If Rst.EOF = False Then
             For Each varlbl In Label15
                 If varlbl.Tag = CStr(Rst!No) Then
                     varlbl.Caption = " : " & Rst!TableDescription & vbLf & " : " & Rst!MinuteUseDiff & "  ÏÞíÞå"            'green for empty
                     If Not IsNull(Rst!nvcMaxUseTime) Then
                        If Rst!MinuteUseDiff >= Val(Rst!nvcMaxUseTime) Then varlbl.BackColor = &HFFFF&: varlbl.Caption = varlbl.Caption & vbLf & "ÇÖÇÝí" & Val(Rst!MinuteUseDiff) - Val(Rst!nvcMaxUseTime) & "  ÏÞíÞå"
                     End If
                     Rst.MoveNext
                     If Rst.EOF Then Exit For
                 End If
             Next
         End If
        If Rst.EOF = False Then
             For Each varlbl In Label16
                 If varlbl.Tag = CStr(Rst!No) Then
                     varlbl.Caption = " : " & Rst!TableDescription & vbLf & " : " & Rst!MinuteUseDiff & "  ÏÞíÞå"            'green for empty
                     If Not IsNull(Rst!nvcMaxUseTime) Then
                        If Rst!MinuteUseDiff >= Val(Rst!nvcMaxUseTime) Then varlbl.BackColor = &HFFFF&: varlbl.Caption = varlbl.Caption & vbLf & "ÇÖÇÝí" & Val(Rst!MinuteUseDiff) - Val(Rst!nvcMaxUseTime) & "  ÏÞíÞå"
                     End If
                     Rst.MoveNext
                     If Rst.EOF Then Exit For
                 End If
             Next
         End If
        If Rst.EOF = False Then
             For Each varlbl In Label17
                 If varlbl.Tag = CStr(Rst!No) Then
                     varlbl.Caption = " : " & Rst!TableDescription & vbLf & " : " & Rst!MinuteUseDiff & "  ÏÞíÞå"            'green for empty
                     If Not IsNull(Rst!nvcMaxUseTime) Then
                        If Rst!MinuteUseDiff >= Val(Rst!nvcMaxUseTime) Then varlbl.BackColor = &HFFFF&: varlbl.Caption = varlbl.Caption & vbLf & "ÇÖÇÝí" & Val(Rst!MinuteUseDiff) - Val(Rst!nvcMaxUseTime) & "  ÏÞíÞå"
                     End If
                     Rst.MoveNext
                     If Rst.EOF Then Exit For
                 End If
             Next
         End If
        If Rst.EOF = False Then
             For Each varlbl In Label18
                 If varlbl.Tag = CStr(Rst!No) Then
                     varlbl.Caption = " : " & Rst!TableDescription & vbLf & " : " & Rst!MinuteUseDiff & "  ÏÞíÞå"            'green for empty
                     If Not IsNull(Rst!nvcMaxUseTime) Then
                        If Rst!MinuteUseDiff >= Val(Rst!nvcMaxUseTime) Then varlbl.BackColor = &HFFFF&: varlbl.Caption = varlbl.Caption & vbLf & "ÇÖÇÝí" & Val(Rst!MinuteUseDiff) - Val(Rst!nvcMaxUseTime) & "  ÏÞíÞå"
                     End If
                     Rst.MoveNext
                     If Rst.EOF Then Exit For
                 End If
             Next
         End If
        If Rst.EOF = False Then
             For Each varlbl In Label19
                 If varlbl.Tag = CStr(Rst!No) Then
                     varlbl.Caption = " : " & Rst!TableDescription & vbLf & " : " & Rst!MinuteUseDiff & "  ÏÞíÞå"            'green for empty
                     If Not IsNull(Rst!nvcMaxUseTime) Then
                        If Rst!MinuteUseDiff >= Val(Rst!nvcMaxUseTime) Then varlbl.BackColor = &HFFFF&: varlbl.Caption = varlbl.Caption & vbLf & "ÇÖÇÝí" & Val(Rst!MinuteUseDiff) - Val(Rst!nvcMaxUseTime) & "  ÏÞíÞå"
                     End If
                     Rst.MoveNext
                     If Rst.EOF Then Exit For
                 End If
             Next
         End If

    End If
    
    Exit Sub
ErrHandler:
    MsgBox err.Description
    LogSave "frmTables", err, "DetectResevedTables"
    Resume Next
End Sub

Private Sub DetectOtherTables(Branch As Integer)
    On Error GoTo ErrHandler
        ReDim Parameter(0)
        Parameter(0) = GenerateInputParameter("@branch", adInteger, 4, Branch)
            
        Set Rst = RunParametricStoredProcedure2Rec("GetTablesWithInvoicePrint", Parameter)
        
        If Not (Rst.EOF And Rst.BOF) Then
            Do While Rst.EOF = False
                For Each varlbl In Label    'Enable the empty tables
                     If varlbl.Tag = CStr(Rst!No) Then
                         varlbl.BackColor = vbRed
                         Rst.MoveNext
                         If Rst.EOF Then Exit For
                    End If
                Next
                If Rst.EOF = False Then
                For Each varlbl In Label1
                     If varlbl.Tag = CStr(Rst!No) Then
                         varlbl.BackColor = vbRed
                         varlbl.Enabled = True
                         Rst.MoveNext
                         If Rst.EOF Then Exit For
                    End If
                Next
                End If
                If Rst.EOF = False Then
                     For Each varlbl In Label2
                          If varlbl.Tag = CStr(Rst!No) Then
                              varlbl.BackColor = vbRed
                              varlbl.Enabled = True
                              Rst.MoveNext
                              If Rst.EOF Then Exit For
                         End If
                     Next
                 End If
                If Rst.EOF = False Then
                     For Each varlbl In Label3
                          If varlbl.Tag = CStr(Rst!No) Then
                              varlbl.BackColor = vbRed
                              varlbl.Enabled = True
                              Rst.MoveNext
                              If Rst.EOF Then Exit For
                         End If
                     Next
                 End If
                If Rst.EOF = False Then
                     For Each varlbl In Label4
                          If varlbl.Tag = CStr(Rst!No) Then
                              varlbl.BackColor = vbRed
                              varlbl.Enabled = True
                              Rst.MoveNext
                              If Rst.EOF Then Exit For
                         End If
                     Next
                 End If
                If Rst.EOF = False Then
                     For Each varlbl In Label5
                          If varlbl.Tag = CStr(Rst!No) Then
                              varlbl.BackColor = vbRed
                              varlbl.Enabled = True
                              Rst.MoveNext
                              If Rst.EOF Then Exit For
                         End If
                     Next
                 End If
                If Rst.EOF = False Then
                     For Each varlbl In Label6
                          If varlbl.Tag = CStr(Rst!No) Then
                              varlbl.BackColor = vbRed
                              varlbl.Enabled = True
                              Rst.MoveNext
                              If Rst.EOF Then Exit For
                         End If
                     Next
                 End If
                If Rst.EOF = False Then
                     For Each varlbl In Label7
                          If varlbl.Tag = CStr(Rst!No) Then
                              varlbl.BackColor = vbRed
                              varlbl.Enabled = True
                              Rst.MoveNext
                              If Rst.EOF Then Exit For
                         End If
                     Next
                 End If
                If Rst.EOF = False Then
                     For Each varlbl In Label8
                          If varlbl.Tag = CStr(Rst!No) Then
                              varlbl.BackColor = vbRed
                              varlbl.Enabled = True
                              Rst.MoveNext
                              If Rst.EOF Then Exit For
                         End If
                     Next
                 End If
                If Rst.EOF = False Then
                     For Each varlbl In Label9
                          If varlbl.Tag = CStr(Rst!No) Then
                              varlbl.BackColor = vbRed
                              varlbl.Enabled = True
                              Rst.MoveNext
                              If Rst.EOF Then Exit For
                         End If
                     Next
                 End If
                If Rst.EOF = False Then
                     For Each varlbl In Label10
                          If varlbl.Tag = CStr(Rst!No) Then
                              varlbl.BackColor = vbRed
                              varlbl.Enabled = True
                              Rst.MoveNext
                              If Rst.EOF Then Exit For
                         End If
                     Next
                 End If
                If Rst.EOF = False Then
                     For Each varlbl In Label11
                          If varlbl.Tag = CStr(Rst!No) Then
                              varlbl.BackColor = vbRed
                              varlbl.Enabled = True
                              Rst.MoveNext
                              If Rst.EOF Then Exit For
                         End If
                     Next
                 End If
                If Rst.EOF = False Then
                     For Each varlbl In Label12
                          If varlbl.Tag = CStr(Rst!No) Then
                              varlbl.BackColor = vbRed
                              varlbl.Enabled = True
                              Rst.MoveNext
                              If Rst.EOF Then Exit For
                         End If
                     Next
                 End If
                If Rst.EOF = False Then
                     For Each varlbl In Label13
                          If varlbl.Tag = CStr(Rst!No) Then
                              varlbl.BackColor = vbRed
                              varlbl.Enabled = True
                              Rst.MoveNext
                              If Rst.EOF Then Exit For
                         End If
                     Next
                 End If
                If Rst.EOF = False Then
                     For Each varlbl In Label14
                          If varlbl.Tag = CStr(Rst!No) Then
                              varlbl.BackColor = vbRed
                              varlbl.Enabled = True
                              Rst.MoveNext
                              If Rst.EOF Then Exit For
                         End If
                     Next
                 End If
                If Rst.EOF = False Then
                     For Each varlbl In Label15
                          If varlbl.Tag = CStr(Rst!No) Then
                              varlbl.BackColor = vbRed
                              varlbl.Enabled = True
                              Rst.MoveNext
                              If Rst.EOF Then Exit For
                         End If
                     Next
                 End If
                If Rst.EOF = False Then
                     For Each varlbl In Label16
                          If varlbl.Tag = CStr(Rst!No) Then
                              varlbl.BackColor = vbRed
                              varlbl.Enabled = True
                              Rst.MoveNext
                              If Rst.EOF Then Exit For
                         End If
                     Next
                 End If
                If Rst.EOF = False Then
                     For Each varlbl In Label17
                          If varlbl.Tag = CStr(Rst!No) Then
                              varlbl.BackColor = vbRed
                              varlbl.Enabled = True
                              Rst.MoveNext
                              If Rst.EOF Then Exit For
                         End If
                     Next
                 End If
                If Rst.EOF = False Then
                     For Each varlbl In Label18
                          If varlbl.Tag = CStr(Rst!No) Then
                              varlbl.BackColor = vbRed
                              varlbl.Enabled = True
                              Rst.MoveNext
                              If Rst.EOF Then Exit For
                         End If
                     Next
                 End If
                If Rst.EOF = False Then
                     For Each varlbl In Label19
                          If varlbl.Tag = CStr(Rst!No) Then
                              varlbl.BackColor = vbRed
                              varlbl.Enabled = True
                              Rst.MoveNext
                              If Rst.EOF Then Exit For
                         End If
                     Next
                 End If
                 If Rst.EOF = False Then Rst.MoveNext
            Loop
        End If
    
    Exit Sub
ErrHandler:
    MsgBox err.Description
    LogSave "frmTables", err, "DetectOtherTables"
End Sub


Private Sub FindInvoice(TableNo As Integer, Branch As Integer)
    On Error GoTo ErrHandler
    ReDim Parameter(1)
    Parameter(0) = GenerateInputParameter("@Branch", adInteger, 4, Branch)
    Parameter(1) = GenerateInputParameter("@TableNO", adInteger, 4, TableNo)
    
    Set Rst = RunParametricStoredProcedure2Rec("GetInvoiceByTable", Parameter)
    mvarInvoiceNO = 0
    If Not (Rst.EOF And Rst.BOF) Then
        mvarInvoiceNO = Rst!No
    End If
    Exit Sub
ErrHandler:
    MsgBox err.Description
    LogSave "frmTables", err, "FindInvoice"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveSetting strMainKey, Me.Name, "TimerInterval", CStr(Val(txtInterval.Value) * 1000)
    SaveSetting strMainKey, Me.Name, "TxtWidth", CStr(Val(TxtWidth.Text))
    SaveSetting strMainKey, Me.Name, "TxtHeight", CStr(Val(TxtHeight.Text))
    SaveSetting strMainKey, Me.Name, "TxtFontSize", CStr(Val(TxtFontSize.Text))

    SaveSetting strMainKey, Me.Name, "Left", Me.Left
    SaveSetting strMainKey, Me.Name, "Top", Me.Top

End Sub


Private Sub ResizeKit1_ExitResize(ByVal XScale As Double, ByVal YScale As Double)
    If formloadFlag = True Then
        SaveSetting strMainKey, Me.Name, "Height", Me.Height
        SaveSetting strMainKey, Me.Name, "Width", Me.Width
    End If

End Sub

Private Sub Timer1_Timer()
    Timer1.Interval = txtInterval.Value * 1000
    If tablesCount > 0 Then
        DetectEmptyTables CurrentBranch
        DetectBusyTables CurrentBranch
        DetectOtherTables CurrentBranch
    End If
End Sub


'Private Sub FindTable(strTableName As String)
'    Dim Result As Boolean
'    Result = False
'    On Error GoTo ErrHandler
'        For Each varlbl In Label
'            If varlbl.Caption = strTableName Then
'                If varlbl.Enabled Then
'                    varlbl.SetFocus
'                End If
''                varlbl.BackColor = vbBlue
'                SSTab1.Tab = 0
'                Result = True
'                Exit For
'            End If
'        Next
'        If Result = False Then
'            For Each varlbl In Label1
'                If varlbl.Caption = strTableName Then
'                    If varlbl.Enabled Then varlbl.SetFocus
''                    varlbl.BackColor = vbBlue
'                    SSTab1.Tab = 1
'                    Result = True
'                    Exit For
'                End If
'            Next
'        End If
'        If Result = False Then
'            For Each varlbl In Label2
'                If varlbl.Caption = strTableName Then
'                    If varlbl.Enabled Then varlbl.SetFocus
''                    varlbl.BackColor = vbBlue
'                    SSTab1.Tab = 2
'                    Result = True
'                    Exit For
'                End If
'            Next
'        End If
'        If Result = False Then
'            For Each varlbl In Label3
'                If varlbl.Caption = strTableName Then
'                    If varlbl.Enabled Then varlbl.SetFocus
''                    varlbl.BackColor = vbBlue
'                    SSTab1.Tab = 3
'                    Result = True
'                    Exit For
'                End If
'            Next
'        End If
'        If Result = False Then
'            For Each varlbl In Label4
'                If varlbl.Caption = strTableName Then
'                    If varlbl.Enabled Then varlbl.SetFocus
''                    varlbl.BackColor = vbBlue
'                    SSTab1.Tab = 4
'                    Result = True
'                    Exit For
'                End If
'            Next
'        End If
'    Exit Sub
'ErrHandler:
'    MsgBox err.Description
'    LogSave "frmTables", err, "FindTable"
'End Sub
Private Function CheckTableState(TableNo As String) As Boolean
    On Error GoTo ErrHandler
    CheckTableState = True
    ReDim Parameter(1)
    Parameter(0) = GenerateInputParameter("@TableNO", adInteger, 4, TableNo)
    Parameter(1) = GenerateInputParameter("@nvcDate", adWChar, 8, mvarDate)
    
    Set Rst = RunParametricStoredProcedure2Rec("CheckTableStatus", Parameter)
    If Not (Rst.EOF And Rst.BOF) Then  ' Not Empty
        CheckTableState = False 'IIf(Rst!Empty = False, False, True)
    End If
Exit Function
ErrHandler:
    MsgBox err.Description
    LogSave "FrmTables", err, "CheckTableState"
End Function


