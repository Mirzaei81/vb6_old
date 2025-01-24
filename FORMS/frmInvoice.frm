VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{7AEDC602-D94C-11D1-BB7A-00E0290EA3C9}#1.0#0"; "ResizeKit.ocx"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Object = "{75D4F767-8785-11D3-93AD-0000832EF44D}#2.14#0"; "FAST2001.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "mscomm32.ocx"
Object = "{158336E7-3FF3-456E-912C-5985E9BBED24}#1.0#0"; "MTUSBHIDSwipe.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{A32FAC36-847B-4323-AC23-038F9131C74A}#6.0#0"; "USBCID.ocx"
Begin VB.Form frmInvoice 
   BackColor       =   &H00C0E0FF&
   Caption         =   "›«ò Ê— ›—Ê‘"
   ClientHeight    =   9615
   ClientLeft      =   120
   ClientTop       =   900
   ClientWidth     =   15360
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   9
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9615
   ScaleWidth      =   15360
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer_PersonIdCheck 
      Enabled         =   0   'False
      Interval        =   7000
      Left            =   2250
      Top             =   165
   End
   Begin VB.Frame Frame21 
      Caption         =   "Frame2"
      Height          =   2295
      Left            =   360
      TabIndex        =   171
      Top             =   960
      Visible         =   0   'False
      Width           =   3255
      Begin VB.TextBox ResultTXT 
         Appearance      =   0  'Flat
         Height          =   255
         Left            =   -840
         Locked          =   -1  'True
         TabIndex        =   182
         Top             =   1920
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Timer TimerRFID 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   0
         Top             =   600
      End
      Begin VB.TextBox BufferTXT 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   720
         TabIndex        =   181
         Top             =   1440
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.TextBox keyTXT 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   5
         Left            =   960
         TabIndex        =   180
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
         TabIndex        =   179
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
         TabIndex        =   178
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
         TabIndex        =   177
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
         TabIndex        =   176
         Text            =   "FF"
         Top             =   1920
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.TextBox keyTXT 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   0
         Left            =   -240
         TabIndex        =   175
         Text            =   "FF"
         Top             =   1920
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.CheckBox KeyAorB 
         BackColor       =   &H00FF8080&
         Caption         =   "KEY A"
         Height          =   255
         Left            =   840
         TabIndex        =   174
         Top             =   240
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.TextBox blockNtxt 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   480
         TabIndex        =   173
         Text            =   "4"
         Top             =   240
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.ComboBox Bcount 
         Height          =   330
         Left            =   1560
         TabIndex        =   172
         Text            =   "1"
         Top             =   240
         Visible         =   0   'False
         Width           =   615
      End
   End
   Begin VB.Timer Timer_Printers 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   240
      Top             =   0
   End
   Begin VB.Frame FrameCustInfo 
      BackColor       =   &H00C0E0FF&
      Caption         =   "                             «ÿ·«⁄«  „‘ —ò                                         "
      BeginProperty Font 
         Name            =   "Nazanin"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6615
      Left            =   8220
      MouseIcon       =   "frmInvoice.frx":0000
      RightToLeft     =   -1  'True
      TabIndex        =   15
      Top             =   705
      Visible         =   0   'False
      Width           =   5535
      Begin VB.CommandButton cmdDeleteCustomer 
         Caption         =   "Õ–› „‘ —Ì «“ ›«ﬂ Ê—"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1000
         Left            =   3600
         Picture         =   "frmInvoice.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   151
         Top             =   5400
         Width           =   1755
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
         Height          =   1000
         Left            =   120
         Picture         =   "frmInvoice.frx":3E9C
         Style           =   1  'Graphical
         TabIndex        =   148
         Top             =   5400
         Width           =   1275
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
         Height          =   975
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   49
         Top             =   4320
         Width           =   5295
         Begin VB.Label LblDescription 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000013&
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Nazanin"
               Size            =   11.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000040&
            Height          =   495
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   50
            Top             =   360
            Width           =   4965
         End
      End
      Begin VB.Frame Frame9 
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3015
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   1320
         Width           =   5295
         Begin VB.Label lblCountMonthBuy 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   480
            RightToLeft     =   -1  'True
            TabIndex        =   45
            Top             =   2040
            Width           =   615
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Œ—Ìœ Â«Ì „«Â"
            BeginProperty Font 
               Name            =   "Nazanin"
               Size            =   9
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1200
            RightToLeft     =   -1  'True
            TabIndex        =   44
            Top             =   2040
            Width           =   1095
         End
         Begin VB.Label lblCountCurrentBuy 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3240
            RightToLeft     =   -1  'True
            TabIndex        =   43
            Top             =   2040
            Width           =   615
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Œ—Ìœ Â«Ì «„—Ê“"
            BeginProperty Font 
               Name            =   "Nazanin"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3600
            RightToLeft     =   -1  'True
            TabIndex        =   42
            Top             =   2040
            Width           =   1335
         End
         Begin VB.Label MaxPrice 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "»Ì‘ —Ì‰ Œ—Ìœ"
            BeginProperty Font 
               Name            =   "Nazanin"
               Size            =   9
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3840
            RightToLeft     =   -1  'True
            TabIndex        =   40
            Top             =   1080
            Width           =   1095
         End
         Begin VB.Label LastDate 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "¬Œ—Ì‰ Œ—Ìœ"
            BeginProperty Font 
               Name            =   "Nazanin"
               Size            =   9
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3960
            RightToLeft     =   -1  'True
            TabIndex        =   39
            Top             =   120
            Width           =   975
         End
         Begin VB.Label BuyAverage 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "„Ì«‰êÌ‰ Œ—Ìœ"
            BeginProperty Font 
               Name            =   "Nazanin"
               Size            =   9
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3960
            RightToLeft     =   -1  'True
            TabIndex        =   38
            Top             =   600
            Width           =   975
         End
         Begin VB.Label LastNo 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "¬Œ—Ì‰ ›Ì‘"
            BeginProperty Font 
               Name            =   "Nazanin"
               Size            =   9
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1200
            RightToLeft     =   -1  'True
            TabIndex        =   37
            Top             =   120
            Width           =   975
         End
         Begin VB.Label AddedDate 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   " «—ÌŒ ⁄÷ÊÌ "
            BeginProperty Font 
               Name            =   "Nazanin"
               Size            =   9
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3960
            RightToLeft     =   -1  'True
            TabIndex        =   36
            Top             =   1560
            Width           =   975
         End
         Begin VB.Label BuyCount 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   " ⁄œ«œ œ›⁄«  Œ—Ìœ"
            BeginProperty Font 
               Name            =   "Nazanin"
               Size            =   9
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1200
            RightToLeft     =   -1  'True
            TabIndex        =   35
            Top             =   600
            Width           =   975
         End
         Begin VB.Label MinPrice 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "ﬂ„ —Ì‰ Œ—Ìœ"
            BeginProperty Font 
               Name            =   "Nazanin"
               Size            =   9
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1200
            RightToLeft     =   -1  'True
            TabIndex        =   34
            Top             =   1080
            Width           =   975
         End
         Begin VB.Label LastPrice 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "„»·€ ¬Œ—Ì‰ Œ—Ìœ"
            BeginProperty Font 
               Name            =   "Nazanin"
               Size            =   9
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   960
            RightToLeft     =   -1  'True
            TabIndex        =   33
            Top             =   1560
            Width           =   1335
         End
         Begin VB.Label MaxPrice1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2880
            RightToLeft     =   -1  'True
            TabIndex        =   32
            Top             =   1080
            Width           =   975
         End
         Begin VB.Label LastDate1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2880
            RightToLeft     =   -1  'True
            TabIndex        =   31
            Top             =   120
            Width           =   975
         End
         Begin VB.Label BuyAverage1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2880
            RightToLeft     =   -1  'True
            TabIndex        =   30
            Top             =   600
            Width           =   975
         End
         Begin VB.Label LastNo1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   29
            Top             =   120
            Width           =   975
         End
         Begin VB.Label AddedDate1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2880
            RightToLeft     =   -1  'True
            TabIndex        =   28
            Top             =   1560
            Width           =   975
         End
         Begin VB.Label BuyCount1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   27
            Top             =   600
            Width           =   975
         End
         Begin VB.Label MinPrice1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   26
            Top             =   1080
            Width           =   975
         End
         Begin VB.Label LastPrice1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   25
            Top             =   1560
            Width           =   975
         End
         Begin VB.Label LastCredit1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   360
            RightToLeft     =   -1  'True
            TabIndex        =   24
            Top             =   2400
            Width           =   855
         End
         Begin VB.Label LastCredit 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "„«‰œÂ "
            BeginProperty Font 
               Name            =   "Nazanin"
               Size            =   9
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1320
            RightToLeft     =   -1  'True
            TabIndex        =   23
            Top             =   2400
            Width           =   375
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "«⁄ »«—Ì"
            BeginProperty Font 
               Name            =   "Nazanin"
               Size            =   9
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   4320
            RightToLeft     =   -1  'True
            TabIndex        =   22
            Top             =   2400
            Width           =   615
         End
         Begin VB.Label Label11 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "œ—Ì«› "
            BeginProperty Font 
               Name            =   "Nazanin"
               Size            =   9
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2760
            RightToLeft     =   -1  'True
            TabIndex        =   21
            Top             =   2400
            Width           =   495
         End
         Begin VB.Label LblBuy1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3480
            RightToLeft     =   -1  'True
            TabIndex        =   20
            Top             =   2400
            Width           =   855
         End
         Begin VB.Label LblRecieve1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1800
            RightToLeft     =   -1  'True
            TabIndex        =   19
            Top             =   2400
            Width           =   855
         End
      End
      Begin VB.Frame Frame8 
         BackColor       =   &H00C0E0FF&
         Caption         =   "«⁄ »«—"
         BeginProperty Font 
            Name            =   "Nazanin"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Top             =   480
         Width           =   5295
         Begin VB.Label lblCredit 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000013&
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
            ForeColor       =   &H00000040&
            Height          =   570
            Left            =   360
            RightToLeft     =   -1  'True
            TabIndex        =   17
            Top             =   360
            Width           =   4485
         End
      End
      Begin FLWCtrls.FWButton cmdTurnOver 
         Height          =   1000
         Left            =   1440
         TabIndex        =   158
         TabStop         =   0   'False
         Tag             =   "-"
         Top             =   5400
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   1773
         ButtonType      =   5
         Caption         =   "ê—œ‘ Õ”«» «Ì‰ „‘ —Ì"
         BackColor       =   12632319
         ForeColor       =   255
         FontName        =   "B Traffic"
         FontBold        =   -1  'True
         FontSize        =   9.75
      End
   End
   Begin VB.PictureBox frameMenu 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6195
      Left            =   8160
      RightToLeft     =   -1  'True
      ScaleHeight     =   6135
      ScaleWidth      =   6090
      TabIndex        =   155
      Top             =   120
      Width           =   6150
      Begin TabDlg.SSTab SSTab1 
         Height          =   615
         Left            =   0
         TabIndex        =   170
         Top             =   0
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   1085
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
         TabPicture(0)   =   "frmInvoice.frx":B566
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).ControlCount=   0
      End
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
         Height          =   600
         Index           =   1
         Left            =   120
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   166
         Top             =   600
         Width           =   1455
      End
      Begin MSComctlLib.StatusBar MenuBar 
         Height          =   615
         Left            =   0
         TabIndex        =   167
         Top             =   0
         Visible         =   0   'False
         Width           =   5925
         _ExtentX        =   10451
         _ExtentY        =   1085
         _Version        =   393216
         BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
            NumPanels       =   6
            BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               Object.Width           =   353
               MinWidth        =   353
               Picture         =   "frmInvoice.frx":B582
            EndProperty
            BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               AutoSize        =   2
               Object.Width           =   2302
               MinWidth        =   2293
            EndProperty
            BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               AutoSize        =   2
               Object.Width           =   2302
               MinWidth        =   2293
            EndProperty
            BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               AutoSize        =   2
               Object.Width           =   2302
               MinWidth        =   2293
            EndProperty
            BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               AutoSize        =   2
               Object.Width           =   2302
               MinWidth        =   2293
            EndProperty
            BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               Object.Width           =   706
               MinWidth        =   706
               Picture         =   "frmInvoice.frx":B89C
            EndProperty
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Nazanin"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.Timer TimerScale 
      Enabled         =   0   'False
      Interval        =   400
      Left            =   0
      Top             =   0
   End
   Begin VB.Frame FrameBascule 
      Caption         =   " —«“Ê"
      Height          =   975
      Left            =   5640
      RightToLeft     =   -1  'True
      TabIndex        =   152
      ToolTipText     =   "„Ì  Ê«‰Ìœ «Ì‰ ›—Ì„ —« »Â Â— Ã«Ì ’›ÕÂ «‰ ﬁ«· œÂÌœ"
      Top             =   6390
      Width           =   2415
      Begin VB.Label BascoleLabel 
         Alignment       =   2  'Center
         BackColor       =   &H80000013&
         BackStyle       =   0  'Transparent
         Caption         =   " —«“ÊÌ 1 :"
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
         Index           =   0
         Left            =   1560
         RightToLeft     =   -1  'True
         TabIndex        =   154
         Top             =   360
         Width           =   855
      End
      Begin VB.Label lblScale 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00A0E0FF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   585
         Index           =   0
         Left            =   120
         TabIndex        =   153
         Top             =   240
         Width           =   1425
      End
   End
   Begin VB.ComboBox cmbServePlace 
      BackColor       =   &H0070C0FF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   150
      TabStop         =   0   'False
      Text            =   "cmbServeplace"
      ToolTipText     =   " ⁄ÌÌ‰ „Õ· ”—Ê "
      Top             =   600
      Width           =   1890
   End
   Begin Total.CallerIDMonitor UCCallerIDMonitor1 
      Height          =   855
      Left            =   2880
      TabIndex        =   149
      Top             =   240
      Visible         =   0   'False
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   1508
      RemoveLen       =   "0"
   End
   Begin VB.ListBox lstDifference 
      BeginProperty DataFormat 
         Type            =   4
         Format          =   ""
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1065
         SubFormatType   =   8
      EndProperty
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   960
      ItemData        =   "frmInvoice.frx":BBB6
      Left            =   480
      List            =   "frmInvoice.frx":BBB8
      RightToLeft     =   -1  'True
      Style           =   1  'Checkbox
      TabIndex        =   10
      Top             =   1800
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2775
      Left            =   8160
      ScaleHeight     =   2745
      ScaleWidth      =   6120
      TabIndex        =   106
      Top             =   6360
      Width           =   6150
      Begin VB.CommandButton BtnKeypad 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000D&
         Caption         =   "Å«ﬂ ﬂ‰"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Index           =   12
         Left            =   4440
         RightToLeft     =   -1  'True
         TabIndex        =   169
         Tag             =   "-"
         Top             =   2040
         Width           =   1560
      End
      Begin FLWCtrls.FWCheck ChkCallerId 
         Height          =   285
         Left            =   360
         TabIndex        =   137
         TabStop         =   0   'False
         ToolTipText     =   "‰„«Ì‘  „«” Â«Ì «‰Ã«„ ‘œÂ œ— ’Ê—  ÊÃÊœ ﬂ«·— ¬ÌœÌ"
         Top             =   2265
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   503
         Value           =   0   'False
         CheckType       =   7
         Caption         =   "·Ì”   „«” Â«"
         Color           =   32768
         BackColor       =   8438015
         ForeColor       =   32768
         FontSize        =   8.25
         Alignment       =   1
         Object.ToolTipText     =   "‰„«Ì‘  „«” Â«Ì «‰Ã«„ ‘œÂ œ— ’Ê—  ÊÃÊœ ﬂ«·— ¬ÌœÌ"
      End
      Begin VB.TextBox txtDescription 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
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
         ForeColor       =   &H00FF0000&
         Height          =   510
         Left            =   80
         RightToLeft     =   -1  'True
         TabIndex        =   138
         Text            =   "        ÅÌ€«„      "
         ToolTipText     =   "ê–«‘ ‰ ÅÌ€«„ —ÊÌ ›Ì‘ »—«Ì Ì«œ¬Ê—Ì Ê À»  œ— ›Ì‘ ¬‘Å“Œ«‰Â"
         Top             =   1065
         Width           =   2000
      End
      Begin VB.Frame Frame_CallerId 
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   900
         Left            =   80
         TabIndex        =   126
         ToolTipText     =   "‰„«Ì‘ ŒÿÊÿ ﬂ«·— ¬ÌœÌ œ— ’Ê—  ÊÃÊœ"
         Top             =   120
         Width           =   1935
         Begin VB.Frame Frame7 
            Caption         =   "Frame7"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1095
            Left            =   0
            TabIndex        =   127
            Top             =   0
            Width           =   15
         End
         Begin FLWCtrls.FWCoolButton FWModem 
            Height          =   420
            Index           =   0
            Left            =   0
            TabIndex        =   128
            ToolTipText     =   "‰„«Ì‘ ŒÿÊÿ ﬂ«·— ¬ÌœÌ"
            Top             =   0
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   741
            BackColor       =   16776960
            ForeColor       =   -2147483635
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "B Nazanin"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PictureAlign    =   4
            Caption         =   "1"
            MaskColor       =   -2147483633
         End
         Begin FLWCtrls.FWCoolButton FWModem 
            Height          =   420
            Index           =   1
            Left            =   480
            TabIndex        =   129
            ToolTipText     =   "‰„«Ì‘ ŒÿÊÿ ﬂ«·— ¬ÌœÌ"
            Top             =   0
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   741
            BackColor       =   16776960
            ForeColor       =   -2147483635
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "B Nazanin"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "2"
            MaskColor       =   -2147483633
         End
         Begin FLWCtrls.FWCoolButton FWModem 
            Height          =   420
            Index           =   2
            Left            =   960
            TabIndex        =   130
            ToolTipText     =   "‰„«Ì‘ ŒÿÊÿ ﬂ«·— ¬ÌœÌ"
            Top             =   0
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   741
            BackColor       =   16776960
            ForeColor       =   -2147483635
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "B Nazanin"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "3"
            MaskColor       =   -2147483633
         End
         Begin FLWCtrls.FWCoolButton FWModem 
            Height          =   420
            Index           =   3
            Left            =   1440
            TabIndex        =   131
            ToolTipText     =   "‰„«Ì‘ ŒÿÊÿ ﬂ«·— ¬ÌœÌ"
            Top             =   0
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   741
            BackColor       =   16776960
            ForeColor       =   -2147483635
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "B Nazanin"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "4"
            MaskColor       =   -2147483633
         End
         Begin FLWCtrls.FWCoolButton FWModem 
            Height          =   420
            Index           =   5
            Left            =   480
            TabIndex        =   132
            ToolTipText     =   "‰„«Ì‘ ŒÿÊÿ ﬂ«·— ¬ÌœÌ"
            Top             =   420
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   741
            BackColor       =   16776960
            ForeColor       =   -2147483635
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "B Nazanin"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "6"
            MaskColor       =   -2147483633
         End
         Begin FLWCtrls.FWCoolButton FWModem 
            Height          =   420
            Index           =   6
            Left            =   960
            TabIndex        =   145
            ToolTipText     =   "‰„«Ì‘ ŒÿÊÿ ﬂ«·— ¬ÌœÌ"
            Top             =   420
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   741
            BackColor       =   16776960
            ForeColor       =   -2147483635
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "B Nazanin"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "7"
            MaskColor       =   -2147483633
         End
         Begin FLWCtrls.FWCoolButton FWModem 
            Height          =   420
            Index           =   7
            Left            =   1440
            TabIndex        =   146
            ToolTipText     =   "‰„«Ì‘ ŒÿÊÿ ﬂ«·— ¬ÌœÌ"
            Top             =   420
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   741
            BackColor       =   16776960
            ForeColor       =   -2147483635
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "B Nazanin"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "8"
            MaskColor       =   -2147483633
         End
         Begin FLWCtrls.FWCoolButton FWModem 
            Height          =   420
            Index           =   4
            Left            =   0
            TabIndex        =   147
            ToolTipText     =   "‰„«Ì‘ ŒÿÊÿ ﬂ«·— ¬ÌœÌ"
            Top             =   420
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   741
            BackColor       =   16776960
            ForeColor       =   -2147483635
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "B Nazanin"
               Size            =   8.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "5"
            MaskColor       =   -2147483633
         End
      End
      Begin VB.CommandButton BtnKeypad 
         BackColor       =   &H8000000D&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "B Nazanin"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Index           =   0
         Left            =   2880
         TabIndex        =   119
         Tag             =   "0"
         Top             =   2110
         Width           =   675
      End
      Begin VB.CommandButton BtnKeypad 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000D&
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "B Nazanin"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Index           =   1
         Left            =   2160
         TabIndex        =   118
         Tag             =   "1"
         Top             =   1440
         UseMaskColor    =   -1  'True
         Width           =   675
      End
      Begin VB.CommandButton BtnKeypad 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000D&
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "B Nazanin"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Index           =   2
         Left            =   2880
         TabIndex        =   117
         Tag             =   "2"
         Top             =   1440
         UseMaskColor    =   -1  'True
         Width           =   675
      End
      Begin VB.CommandButton BtnKeypad 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000D&
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "B Nazanin"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Index           =   3
         Left            =   3600
         TabIndex        =   116
         Tag             =   "3"
         Top             =   1440
         UseMaskColor    =   -1  'True
         Width           =   675
      End
      Begin VB.CommandButton BtnKeypad 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000D&
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "B Nazanin"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Index           =   4
         Left            =   2160
         TabIndex        =   115
         Tag             =   "4"
         Top             =   750
         Width           =   675
      End
      Begin VB.CommandButton BtnKeypad 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000D&
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "B Nazanin"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Index           =   5
         Left            =   2880
         TabIndex        =   114
         Tag             =   "5"
         Top             =   750
         Width           =   675
      End
      Begin VB.CommandButton BtnKeypad 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000D&
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "B Nazanin"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Index           =   6
         Left            =   3600
         TabIndex        =   113
         Tag             =   "6"
         Top             =   750
         Width           =   675
      End
      Begin VB.CommandButton BtnKeypad 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000D&
         Caption         =   "7"
         BeginProperty Font 
            Name            =   "B Nazanin"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Index           =   7
         Left            =   2160
         TabIndex        =   112
         Tag             =   "7"
         Top             =   50
         Width           =   675
      End
      Begin VB.CommandButton BtnKeypad 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000D&
         Caption         =   "8"
         BeginProperty Font 
            Name            =   "B Nazanin"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Index           =   8
         Left            =   2880
         TabIndex        =   111
         Tag             =   "8"
         Top             =   50
         Width           =   675
      End
      Begin VB.CommandButton BtnKeypad 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000D&
         Caption         =   "9"
         BeginProperty Font 
            Name            =   "B Nazanin"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Index           =   9
         Left            =   3600
         TabIndex        =   110
         Tag             =   "9"
         Top             =   50
         Width           =   675
      End
      Begin VB.CommandButton BtnKeypad 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000D&
         Caption         =   "."
         BeginProperty Font 
            Name            =   "B Nazanin"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Index           =   10
         Left            =   2160
         RightToLeft     =   -1  'True
         TabIndex        =   109
         Tag             =   "."
         Top             =   2110
         Width           =   675
      End
      Begin VB.CommandButton BtnKeypad 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000D&
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "B Nazanin"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Index           =   11
         Left            =   3600
         RightToLeft     =   -1  'True
         TabIndex        =   108
         Tag             =   "%"
         Top             =   2110
         Width           =   675
      End
      Begin FLWCtrls.FWLabel FwPartition 
         Height          =   435
         Left            =   1440
         Top             =   0
         Visible         =   0   'False
         Width           =   2220
         _ExtentX        =   3916
         _ExtentY        =   767
         Enabled         =   -1  'True
         Caption         =   ""
         FillType        =   3
         FirstColor      =   16711680
         SecondColor     =   0
         Angle           =   0
         ForeColor       =   0
         BackColor       =   128
         FontName        =   "Nazanin"
         FontBold        =   -1  'True
         FontSize        =   11.25
         Alignment       =   2
         Picture         =   "frmInvoice.frx":BBBA
         BorderStyle     =   1
      End
      Begin FLWCtrls.FWLabel fwCash 
         Height          =   375
         Left            =   1080
         Top             =   0
         Visible         =   0   'False
         Width           =   2220
         _ExtentX        =   3916
         _ExtentY        =   661
         Enabled         =   -1  'True
         Caption         =   ""
         FillType        =   3
         FirstColor      =   9981440
         SecondColor     =   16777215
         Angle           =   0
         ForeColor       =   0
         BackColor       =   8421631
         FontName        =   "Nazanin"
         FontBold        =   -1  'True
         FontSize        =   11.25
         Alignment       =   2
         Picture         =   "frmInvoice.frx":BBD6
      End
      Begin FLWCtrls.FWLabel FWlblAcc 
         Height          =   315
         Left            =   1920
         Top             =   0
         Visible         =   0   'False
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   556
         Enabled         =   -1  'True
         Caption         =   "Õ”«» Â« »” Â"
         FillType        =   3
         FirstColor      =   16711680
         SecondColor     =   0
         Angle           =   0
         ForeColor       =   0
         BackColor       =   128
         FontName        =   "Nazanin"
         FontSize        =   11.25
         Alignment       =   2
         Picture         =   "frmInvoice.frx":BBF2
         BorderStyle     =   1
      End
      Begin FLWCtrls.FWLabel FWlblCash 
         Height          =   315
         Left            =   2040
         Top             =   0
         Visible         =   0   'False
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   556
         Enabled         =   -1  'True
         Caption         =   "’‰œÊﬁ »” Â"
         FillType        =   3
         FirstColor      =   16711680
         SecondColor     =   0
         Angle           =   0
         ForeColor       =   0
         BackColor       =   128
         FontName        =   "Nazanin"
         FontBold        =   -1  'True
         FontSize        =   9.75
         Alignment       =   2
         Picture         =   "frmInvoice.frx":BC0E
         BorderStyle     =   1
      End
      Begin FLWCtrls.FWCheck FWChkHavale 
         Height          =   60
         Left            =   360
         TabIndex        =   136
         TabStop         =   0   'False
         ToolTipText     =   "‰„«Ì‘ Ê÷⁄Ì  ÕÊ«·Â «‰»«— «‰Ã«„ "
         Top             =   2040
         Visible         =   0   'False
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   106
         Value           =   0   'False
         CheckType       =   7
         Caption         =   "«‰ ﬁ«·  ÕÊ«·Â"
         Enabled         =   0   'False
         Color           =   -2147483639
         BackColor       =   8438015
         ForeColor       =   0
         FontSize        =   9.75
         Alignment       =   1
         Object.ToolTipText     =   "‰„«Ì‘ Ê÷⁄Ì  ÕÊ«·Â «‰»«— «‰Ã«„ "
      End
      Begin FLWCtrls.FWButton FWBtnSplit 
         Height          =   465
         Left            =   2400
         TabIndex        =   140
         TabStop         =   0   'False
         Tag             =   "-"
         Top             =   0
         Visible         =   0   'False
         Width           =   840
         _ExtentX        =   1482
         _ExtentY        =   820
         ButtonType      =   3
         Caption         =   "„⁄„Ê·Ì"
         BackColor       =   16384
         ForeColor       =   255
         FontName        =   "Nazanin"
         FontSize        =   9.75
      End
      Begin FLWCtrls.FWCheck FWChkAccount 
         Height          =   75
         Left            =   0
         TabIndex        =   157
         Top             =   2040
         Visible         =   0   'False
         Width           =   2190
         _ExtentX        =   3863
         _ExtentY        =   132
         Value           =   0   'False
         CheckType       =   5
         Caption         =   "«‰ ﬁ«·  »Â Õ”«»œ«—Ì"
         Enabled         =   0   'False
         Color           =   49152
         BackColor       =   16765183
         ForeColor       =   4194304
         Alignment       =   1
      End
      Begin VB.Label LblAccNo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   ": ”‰œ ‘„«—Â "
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   120
         TabIndex        =   156
         Top             =   1680
         Width           =   1695
      End
      Begin VB.Shape Shape1 
         Height          =   2655
         Left            =   4320
         Shape           =   4  'Rounded Rectangle
         Top             =   45
         Width           =   1770
      End
      Begin VB.Label LblTip 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   405
         Left            =   4485
         RightToLeft     =   -1  'True
         TabIndex        =   139
         ToolTipText     =   "‰„«Ì‘ «‰⁄«„ Â«Ì Å—œ«Œ  ‘œÂ —ÊÌ Â— ›Ì‘"
         Top             =   1455
         Width           =   1455
      End
      Begin VB.Label lblBarCode 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   405
         Left            =   4485
         TabIndex        =   135
         ToolTipText     =   "‰„«Ì‘ »«—ﬂœ"
         Top             =   975
         Width           =   1455
      End
      Begin VB.Label lblNum 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   405
         Left            =   4485
         RightToLeft     =   -1  'True
         TabIndex        =   134
         ToolTipText     =   "‰„«Ì‘ «—ﬁ«„ Ê—ÊœÌ  Ê”ÿ ﬂÌ Åœ Ê ﬂÌ»Ê—œ"
         Top             =   510
         Width           =   1455
      End
      Begin VB.Label LblRate 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   330
         Left            =   4680
         TabIndex        =   133
         Top             =   120
         Width           =   1095
      End
      Begin VB.Shape Shape10 
         Height          =   2655
         Left            =   0
         Shape           =   4  'Rounded Rectangle
         Top             =   45
         Width           =   2130
      End
      Begin VB.Label LblInvoicePrint 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "B Nazanin"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   450
         Left            =   1800
         TabIndex        =   125
         ToolTipText     =   "‰„«Ì‘  ⁄œ«œç«Å ›«ﬂ Ê— —ÊÌ Â— ›Ì‘"
         Top             =   1710
         Width           =   255
      End
      Begin VB.Label LblRemain 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   450
         Left            =   75
         RightToLeft     =   -1  'True
         TabIndex        =   124
         ToolTipText     =   "»«ﬁÌ„«‰œÂ ÊÃÂ ›«ﬂ Ê— —« ‰‘«‰ „Ì œÂœ"
         Top             =   1710
         Visible         =   0   'False
         Width           =   1695
      End
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3615
      Left            =   0
      ScaleHeight     =   3555
      ScaleWidth      =   2715
      TabIndex        =   70
      Top             =   5640
      Width           =   2775
      Begin VB.Frame Frame6 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   3615
         Left            =   0
         TabIndex        =   71
         Top             =   0
         Width           =   2655
         Begin VB.PictureBox framelastFich 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1515
            Left            =   80
            RightToLeft     =   -1  'True
            ScaleHeight     =   1455
            ScaleWidth      =   2445
            TabIndex        =   141
            Top             =   2025
            Visible         =   0   'False
            Width           =   2500
            Begin VB.Label lblLastPrice 
               Alignment       =   2  'Center
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
                  Name            =   "Titr"
                  Size            =   14.25
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   675
               Left            =   120
               RightToLeft     =   -1  'True
               TabIndex        =   143
               Top             =   600
               Width           =   2200
            End
            Begin VB.Label Label15 
               Alignment       =   2  'Center
               BackColor       =   &H00EACCEC&
               BackStyle       =   0  'Transparent
               Caption         =   "„»·€ ¬Œ—Ì‰ ›Ì‘"
               BeginProperty Font 
                  Name            =   "Titr"
                  Size            =   12
                  Charset         =   178
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00808000&
               Height          =   615
               Left            =   120
               RightToLeft     =   -1  'True
               TabIndex        =   142
               Top             =   0
               Width           =   2280
            End
            Begin VB.Shape Shape9 
               BackColor       =   &H00008000&
               FillColor       =   &H00FFC0C0&
               FillStyle       =   0  'Solid
               Height          =   855
               Left            =   45
               Shape           =   4  'Rounded Rectangle
               Top             =   480
               Width           =   2300
            End
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "Ã„⁄     "
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
            Height          =   375
            Left            =   1620
            RightToLeft     =   -1  'True
            TabIndex        =   87
            Top             =   120
            Width           =   975
         End
         Begin VB.Label LblSubTotal 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
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
               Size            =   15.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   345
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   86
            Top             =   120
            Width           =   1455
         End
         Begin VB.Label lblCarryFeeTotal 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   15.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   345
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   85
            Top             =   885
            Width           =   1455
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Ã„⁄  ò·   "
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9.75
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   1620
            RightToLeft     =   -1  'True
            TabIndex        =   84
            Top             =   2760
            Width           =   975
         End
         Begin VB.Label LblPacking 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "»” Â »‰œÌ"
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
            Height          =   375
            Left            =   1620
            RightToLeft     =   -1  'True
            TabIndex        =   83
            Top             =   1260
            Width           =   975
         End
         Begin VB.Label LblCarryFee 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "ò—«ÌÂ Õ„·"
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
            Height          =   375
            Left            =   1620
            RightToLeft     =   -1  'True
            TabIndex        =   82
            Top             =   880
            Width           =   975
         End
         Begin VB.Label LblDiscount 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   " Œ›Ì›"
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
            Height          =   375
            Left            =   1620
            RightToLeft     =   -1  'True
            TabIndex        =   81
            Top             =   500
            Width           =   975
         End
         Begin VB.Label lblSumPrice 
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
               Name            =   "Tahoma"
               Size            =   20.25
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   500
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   80
            ToolTipText     =   "Ã„⁄ ﬂ· „»·€ ‰Â«∆Ì ›Ì‘"
            Top             =   3050
            Width           =   2355
         End
         Begin VB.Label lblPackingTotal 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   15.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   345
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   79
            Top             =   1260
            Width           =   1455
         End
         Begin VB.Label lblDiscountTotal 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   345
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   78
            Top             =   495
            Width           =   1455
         End
         Begin VB.Label lblServiceTotal 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   15.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   345
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   77
            Top             =   1635
            Width           =   1455
         End
         Begin VB.Label LblService 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "”—ÊÌ”"
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
            Height          =   375
            Left            =   1620
            RightToLeft     =   -1  'True
            TabIndex        =   76
            Top             =   1640
            Width           =   975
         End
         Begin VB.Label lblDuty 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "⁄Ê«—÷"
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
            Height          =   375
            Left            =   1620
            RightToLeft     =   -1  'True
            TabIndex        =   75
            Top             =   2020
            Width           =   975
         End
         Begin VB.Label lblTaxTotal 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   15.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   345
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   74
            Top             =   2400
            Width           =   1455
         End
         Begin VB.Label LblDutyTotal 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   15.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   345
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   73
            Top             =   2025
            Width           =   1455
         End
         Begin VB.Label LblTax 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "„«·Ì« "
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
            Height          =   375
            Left            =   1620
            RightToLeft     =   -1  'True
            TabIndex        =   72
            Top             =   2400
            Width           =   975
         End
         Begin VB.Shape Shape4 
            BackColor       =   &H00008000&
            FillColor       =   &H00FFC0C0&
            FillStyle       =   0  'Solid
            Height          =   495
            Left            =   120
            Shape           =   4  'Rounded Rectangle
            Top             =   3050
            Width           =   2385
         End
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   3550
      Left            =   2700
      ScaleHeight     =   3525
      ScaleWidth      =   5325
      TabIndex        =   66
      Top             =   5640
      Width           =   5360
      Begin VB.Frame Frame_Printers 
         Height          =   500
         Left            =   120
         TabIndex        =   159
         Top             =   3000
         Visible         =   0   'False
         Width           =   5145
         Begin VB.Label lblPrinter 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "0"
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
            Index           =   5
            Left            =   4680
            TabIndex        =   165
            Top             =   150
            Width           =   135
         End
         Begin VB.Label lblPrinter 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "0"
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
            Index           =   4
            Left            =   3840
            TabIndex        =   164
            Top             =   150
            Width           =   135
         End
         Begin VB.Label lblPrinter 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "0"
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
            Index           =   3
            Left            =   3000
            TabIndex        =   163
            Top             =   150
            Width           =   135
         End
         Begin VB.Label lblPrinter 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "0"
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
            Index           =   2
            Left            =   2040
            TabIndex        =   162
            Top             =   150
            Width           =   135
         End
         Begin VB.Label lblPrinter 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "0"
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
            Index           =   1
            Left            =   1200
            TabIndex        =   161
            Top             =   150
            Width           =   135
         End
         Begin VB.Label lblPrinter 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "0"
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
            Index           =   0
            Left            =   300
            TabIndex        =   160
            Top             =   150
            Width           =   135
         End
         Begin VB.Shape PrinterShape 
            FillColor       =   &H000000FF&
            FillStyle       =   0  'Solid
            Height          =   375
            Index           =   2
            Left            =   1896
            Shape           =   3  'Circle
            Top             =   120
            Width           =   495
         End
         Begin VB.Shape PrinterShape 
            FillColor       =   &H000000FF&
            FillStyle       =   0  'Solid
            Height          =   375
            Index           =   5
            Left            =   4560
            Shape           =   3  'Circle
            Top             =   120
            Width           =   495
         End
         Begin VB.Shape PrinterShape 
            FillColor       =   &H000000FF&
            FillStyle       =   0  'Solid
            Height          =   375
            Index           =   4
            Left            =   3672
            Shape           =   3  'Circle
            Top             =   120
            Width           =   495
         End
         Begin VB.Shape PrinterShape 
            FillColor       =   &H000000FF&
            FillStyle       =   0  'Solid
            Height          =   375
            Index           =   3
            Left            =   2784
            Shape           =   3  'Circle
            Top             =   120
            Width           =   495
         End
         Begin VB.Shape PrinterShape 
            FillColor       =   &H000000FF&
            FillStyle       =   0  'Solid
            Height          =   375
            Index           =   1
            Left            =   1008
            Shape           =   3  'Circle
            Top             =   120
            Width           =   495
         End
         Begin VB.Shape PrinterShape 
            FillColor       =   &H000000FF&
            FillStyle       =   0  'Solid
            Height          =   375
            Index           =   0
            Left            =   120
            Shape           =   3  'Circle
            Top             =   120
            Width           =   495
         End
      End
      Begin VB.ComboBox CmbPayk 
         BackColor       =   &H80000003&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   360
         RightToLeft     =   -1  'True
         TabIndex        =   144
         TabStop         =   0   'False
         Text            =   "cmbPayk"
         ToolTipText     =   "«‰ Œ«» ÅÌﬂ —ÊÌ ›«ﬂ Ê— "
         Top             =   2500
         Width           =   1650
      End
      Begin VB.ComboBox cmbTable 
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         IntegralHeight  =   0   'False
         ItemData        =   "frmInvoice.frx":BC2A
         Left            =   120
         List            =   "frmInvoice.frx":BC2C
         RightToLeft     =   -1  'True
         TabIndex        =   100
         ToolTipText     =   "«‰ Œ«» „Ì“ »—«Ì ”›«—‘"
         Top             =   120
         Width           =   1215
      End
      Begin VB.ComboBox cmbGarson 
         BackColor       =   &H80000003&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   99
         TabStop         =   0   'False
         Text            =   "cmbGarson"
         ToolTipText     =   "«‰ Œ«» ê«—”Ê‰ Ì« ‰„«Ì‘ ê«—”Ê‰ „— »ÿ »« „Ì“ «‰ Œ«» ‘œÂ"
         Top             =   680
         Width           =   1935
      End
      Begin VB.CommandButton cmdTables 
         Caption         =   "„Ì“"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1320
         RightToLeft     =   -1  'True
         TabIndex        =   98
         ToolTipText     =   "F12- ‰„«Ì‘ „Ì“Â«"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   735
      End
      Begin VB.CommandButton cmdPayFactor 
         Caption         =   "œ—Ì«› "
         Height          =   450
         Left            =   1200
         TabIndex        =   93
         ToolTipText     =   "«» œ« „»·€  —« »« «—ﬁ«„ Ê«—œ Ê »« ›‘«— «Ì‰ ﬂ·Ìœ „»·€  Ê«—œ ‘œÂ  œ— ﬁ”„  œ—Ì«›  ‰„«Ì‘ Ê À»  „Ì ê—œœ"
         Top             =   1200
         Width           =   735
      End
      Begin VB.TextBox TxtTempAddress 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   585
         Left            =   2160
         RightToLeft     =   -1  'True
         TabIndex        =   92
         ToolTipText     =   "¬œ—” „Êﬁ  —« Ê«—œ ò‰Ìœ"
         Top             =   2400
         Width           =   3090
      End
      Begin VB.TextBox TxtAddress 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   615
         Left            =   2160
         RightToLeft     =   -1  'True
         TabIndex        =   89
         ToolTipText     =   " —« »“‰Ìœ  « À»  ê—œœ(Enter) ¬œ—” „‘ —Ì —«  Ê«—œ ò‰Ìœ Ê ò·Ìœ   "
         Top             =   1755
         Width           =   3090
      End
      Begin VB.TextBox TxtCustDescription 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   585
         Left            =   2160
         RightToLeft     =   -1  'True
         TabIndex        =   88
         ToolTipText     =   "—« »—‰Ìœ  « À»  ê—œœ(Enter)  Ê÷ÌÕ«  „‘ —Ì —«  Ê«—œ ò‰Ìœ Ê ò·Ìœ  "
         Top             =   2400
         Visible         =   0   'False
         Width           =   3090
      End
      Begin FLWCtrls.FWCoolButton lblCustomer 
         Height          =   555
         Left            =   2160
         TabIndex        =   90
         TabStop         =   0   'False
         ToolTipText     =   "«‰ Œ«» „‘ —Ì —ÊÌ ›«ﬂ Ê—"
         Top             =   600
         Width           =   3090
         _ExtentX        =   5450
         _ExtentY        =   979
         BackColor       =   -2147483648
         ForeColor       =   128
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "frmInvoice.frx":BC2E
         PictureAlign    =   3
         Caption         =   "„‘‰—Ì"
         MaskColor       =   -2147483633
      End
      Begin MSMask.MaskEdBox TxtGuestNo 
         Height          =   405
         Left            =   2280
         TabIndex        =   94
         ToolTipText     =   " ⁄œ«œ ‰›—« Ì ﬂÂ «“ „Ì“ «” ›«œÂ „Ìﬂ‰‰œ"
         Top             =   120
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   714
         _Version        =   393216
         PromptInclude   =   0   'False
         AllowPrompt     =   -1  'True
         MaxLength       =   3
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "B Nazanin"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "999"
         PromptChar      =   "_"
      End
      Begin FLWCtrls.FWScrollText FWScrolltextPay 
         Height          =   500
         Left            =   150
         TabIndex        =   101
         ToolTipText     =   "Å«ﬂ ﬂ—œ‰ œ—Ì«›  Â«Ì «‰Ã«„ ‘œÂ —ÊÌ «Ì‰ ›Ì‘ »« ﬂ·Ìﬂ —ÊÌ ¬‰ Ê œ«‘ ‰ œ” —”Ì(Ctrl+F10)"
         Top             =   1850
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   873
         Caption         =   ""
         BorderStyle     =   0
         BackColor       =   -2147483629
         FontSize        =   9.75
         Object.ToolTipText     =   "Å«ﬂ ﬂ—œ‰ œ—Ì«›  Â«Ì «‰Ã«„ ‘œÂ —ÊÌ «Ì‰ ›Ì‘ »« ﬂ·Ìﬂ —ÊÌ ¬‰ Ê œ«‘ ‰ œ” —”Ì(Ctrl+F10)"
      End
      Begin FLWCtrls.FWScrollText FWScrollSend 
         Height          =   500
         Left            =   1120
         TabIndex        =   102
         Top             =   1850
         Visible         =   0   'False
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   873
         Caption         =   ""
         BorderStyle     =   0
         BackColor       =   -2147483644
         FontSize        =   9.75
      End
      Begin FLWCtrls.FWCheck chKTax 
         Height          =   405
         Left            =   4395
         TabIndex        =   183
         Top             =   75
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   714
         Value           =   0   'False
         CheckType       =   5
         Caption         =   "—”„Ì"
         Color           =   4210688
         BackColor       =   16765183
         ForeColor       =   4194304
         FontName        =   "Nazanin"
         FontBold        =   -1  'True
         FontSize        =   11.25
         Alignment       =   1
      End
      Begin VB.Shape Shape8 
         Height          =   495
         Left            =   2160
         Top             =   60
         Width           =   3090
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "‰›—« "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   375
         Left            =   2880
         TabIndex        =   105
         Top             =   150
         Width           =   495
      End
      Begin VB.Shape Shape7 
         Height          =   585
         Left            =   120
         Top             =   1150
         Width           =   1935
      End
      Begin VB.Shape Shape6 
         Height          =   585
         Left            =   120
         Top             =   1800
         Width           =   1935
      End
      Begin VB.Shape Shape5 
         Height          =   580
         Left            =   120
         Top             =   2400
         Width           =   1935
      End
      Begin VB.Label LblFacpayment 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   104
         Top             =   2400
         Width           =   255
      End
      Begin VB.Label lblDeliveryFullName 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Nazanin"
            Size            =   11.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   525
         Left            =   915
         TabIndex        =   103
         Top             =   2400
         Width           =   1575
      End
      Begin VB.Label txtSumCountNo 
         Alignment       =   2  'Center
         BackColor       =   &H0080C0FF&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3420
         RightToLeft     =   -1  'True
         TabIndex        =   97
         ToolTipText     =   "‰„«Ì‘  ⁄œ« «ﬁ·«„  ﬂ«·«Â«"
         Top             =   120
         Width           =   495
      End
      Begin VB.Label lblPayFactorTotal 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000013&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   435
         Left            =   170
         RightToLeft     =   -1  'True
         TabIndex        =   96
         ToolTipText     =   "œ—Ì«›  —ÊÌ ›Ì‘ »Â ’Ê—  ﬂ«„· Ì« ⁄·Ì «·Õ”«»"
         Top             =   1200
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "«ﬁ·«„"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   375
         Left            =   3945
         TabIndex        =   95
         Top             =   135
         Width           =   375
      End
      Begin VB.Label fwStatusBarCust 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   525
         Left            =   2160
         TabIndex        =   91
         ToolTipText     =   "„‘Œ’«    ﬂ„Ì·Ì „‘ —Ì"
         Top             =   1200
         Width           =   3090
      End
      Begin VB.Label lblTemporary 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   360
         Left            =   120
         TabIndex        =   69
         Top             =   3050
         Visible         =   0   'False
         Width           =   1515
      End
      Begin VB.Label lblDailyDelivered 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   360
         Left            =   1680
         TabIndex        =   68
         Top             =   3050
         Visible         =   0   'False
         Width           =   1755
      End
      Begin VB.Label lblDailyDelivery 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   405
         Left            =   3360
         TabIndex        =   67
         Top             =   3050
         Visible         =   0   'False
         Width           =   1875
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00008000&
         FillColor       =   &H0080FFFF&
         FillStyle       =   0  'Solid
         Height          =   480
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   3000
         Width           =   5145
      End
   End
   Begin VB.Timer TimerAlmP6 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   1680
      Top             =   360
   End
   Begin VB.TextBox txtScale 
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
      Left            =   6120
      TabIndex        =   46
      Text            =   "TxtScale"
      Top             =   4680
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox txtDate 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   120
      Locked          =   -1  'True
      RightToLeft     =   -1  'True
      TabIndex        =   41
      TabStop         =   0   'False
      Top             =   150
      Width           =   1245
   End
   Begin VB.Timer TimerALM 
      Enabled         =   0   'False
      Interval        =   300
      Left            =   120
      Top             =   0
   End
   Begin USBCID.USBCallerID USBCallerID1 
      Left            =   9600
      Top             =   0
      _ExtentX        =   1296
      _ExtentY        =   1296
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   9135
      Left            =   14320
      TabIndex        =   6
      Top             =   0
      Width           =   975
      Begin VB.CommandButton BtnFindGood 
         BackColor       =   &H0000FFFF&
         Caption         =   "Ã” ÃÊÌ ò«·«"
         Height          =   900
         Left            =   75
         Picture         =   "frmInvoice.frx":C508
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   168
         Top             =   120
         Width           =   855
      End
      Begin VB.CommandButton CmdColor 
         BackColor       =   &H000080C0&
         Caption         =   " €ÌÌ— —‰ê"
         Height          =   765
         Left            =   80
         Picture         =   "frmInvoice.frx":DFE2
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   123
         Top             =   4800
         Width           =   800
      End
      Begin VB.CommandButton CmdPager 
         BackColor       =   &H0000A0C0&
         Caption         =   "ÅÌÃ—"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   780
         Left            =   80
         Picture         =   "frmInvoice.frx":E2EC
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   121
         ToolTipText     =   "›—«ŒÊ«‰ 3 —ﬁ„Ì Â„—«Â »« 5 ¬·«—„  ÊÃÂ"
         Top             =   5643
         Width           =   800
      End
      Begin VB.CheckBox ChkIsLocked 
         Alignment       =   1  'Right Justify
         Caption         =   "ﬁ›·"
         CausesValidation=   0   'False
         DownPicture     =   "frmInvoice.frx":10E7E
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   900
         Left            =   80
         MaskColor       =   &H00FFFFFF&
         Picture         =   "frmInvoice.frx":11D48
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   51
         Top             =   6466
         Width           =   800
      End
      Begin VB.CommandButton cmdTempFich 
         BackColor       =   &H0080C0FF&
         Caption         =   "›Ì‘ „Êﬁ "
         Height          =   900
         Left            =   80
         Picture         =   "frmInvoice.frx":12C12
         Style           =   1  'Graphical
         TabIndex        =   48
         Top             =   2949
         Width           =   855
      End
      Begin VB.CommandButton CmdStationSaleSummery 
         BackColor       =   &H008080FF&
         Caption         =   "ê“«—‘ ’‰œÊﬁ"
         Height          =   900
         Left            =   80
         Picture         =   "frmInvoice.frx":134DC
         Style           =   1  'Graphical
         TabIndex        =   47
         Top             =   2006
         Width           =   855
      End
      Begin FLWCtrls.FWCoolButton FWBtnPayk 
         Height          =   900
         Left            =   75
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   3892
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   1588
         BackColor       =   16576
         ForeColor       =   12640511
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "frmInvoice.frx":13DA6
         Caption         =   "ÅÌﬂ"
         MaskColor       =   -2147483633
      End
      Begin FLWCtrls.FWCoolButton cmdPay 
         Height          =   900
         Left            =   75
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   1063
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   1588
         BackColor       =   49344
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "frmInvoice.frx":140C0
         Caption         =   "œ—Ì«› "
         MaskColor       =   -2147483633
      End
      Begin FLWCtrls.FWButton FWMojodiControl 
         Height          =   825
         Left            =   75
         TabIndex        =   120
         TabStop         =   0   'False
         Tag             =   "-"
         Top             =   8295
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   1455
         ButtonType      =   6
         Caption         =   "ﬂ‰ —· „ÊÃÊœÌ"
         Enabled         =   0   'False
         BackColor       =   16384
         ForeColor       =   255
         FontSize        =   8.25
      End
      Begin FLWCtrls.FWButton BtnKalaDelete 
         Height          =   840
         Left            =   75
         TabIndex        =   122
         TabStop         =   0   'False
         Tag             =   "-"
         Top             =   7409
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   1482
         ButtonType      =   2
         Caption         =   "Õ–› ò«·«"
         BackColor       =   -2147483643
         ForeColor       =   255
         Alignment       =   1
      End
   End
   Begin VSFlex7LCtl.VSFlexGrid FlxDetail 
      Height          =   4545
      Left            =   0
      TabIndex        =   2
      Top             =   1080
      Width           =   8085
      _cx             =   14261
      _cy             =   8017
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483629
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483644
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483645
      BackColorAlternate=   -2147483628
      GridColor       =   -2147483632
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   3
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   3
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   12
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   17
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   500
      RowHeightMax    =   500
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmInvoice.frx":1499A
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
      WordWrap        =   -1  'True
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
   Begin ctlUSBHID.USBHID USBHID1 
      Left            =   7320
      Top             =   1200
      _ExtentX        =   873
      _ExtentY        =   873
   End
   Begin FLWCtrls.FWLabel fwlblRecursive 
      Height          =   405
      Left            =   5520
      Top             =   150
      Visible         =   0   'False
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   714
      Enabled         =   -1  'True
      Caption         =   "„—ÃÊ⁄Ì"
      FirstColor      =   12632319
      SecondColor     =   192
      Angle           =   0
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      FontName        =   "Nazanin"
      FontBold        =   -1  'True
      FontItalic      =   -1  'True
      FontSize        =   15.75
      Alignment       =   2
      Picture         =   "frmInvoice.frx":14A96
   End
   Begin FLWCtrls.FWLabel FWLblEdit 
      Height          =   405
      Left            =   5520
      Top             =   530
      Visible         =   0   'False
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   714
      Enabled         =   -1  'True
      Caption         =   "«’·«ÕÌ"
      FirstColor      =   16777088
      SecondColor     =   4210688
      Angle           =   0
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      FontName        =   "Nazanin"
      FontBold        =   -1  'True
      FontItalic      =   -1  'True
      FontSize        =   15.75
      Alignment       =   2
      Picture         =   "frmInvoice.frx":14AB2
   End
   Begin VB.Timer TimerReader 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   3000
      Top             =   5400
   End
   Begin VB.Timer TimerNumber 
      Interval        =   8000
      Left            =   1440
      Top             =   120
   End
   Begin MSCommLib.MSComm mscSerial 
      Index           =   3
      Left            =   1320
      Top             =   5160
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin MSCommLib.MSComm mscSerial 
      Index           =   4
      Left            =   1920
      Top             =   5160
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin MSCommLib.MSComm mscSerial 
      Index           =   1
      Left            =   120
      Top             =   5160
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin MSCommLib.MSComm mscSerial 
      Index           =   2
      Left            =   750
      Top             =   5160
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin Crystal.CrystalReport CrystalReport2 
      Left            =   180
      Top             =   4980
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin FLWCtrls.FWLed FWLed1 
      Height          =   615
      Left            =   6840
      TabIndex        =   0
      Top             =   360
      Visible         =   0   'False
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1085
      BorderStyle     =   9
      ColorOn         =   192
      ColorOff        =   16777215
      BackColor       =   16777215
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   7230
      Top             =   2040
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   11400
      Top             =   100
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInvoice.frx":14ACE
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInvoice.frx":14DE8
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInvoice.frx":15102
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmInvoice.frx":159DC
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar 
      Height          =   390
      Left            =   45
      TabIndex        =   1
      Top             =   9210
      Width           =   7935
      _ExtentX        =   13996
      _ExtentY        =   688
      SimpleText      =   "\"
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   7
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   970
            MinWidth        =   970
            Picture         =   "frmInvoice.frx":15DDC
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   2831
            MinWidth        =   2822
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   2831
            MinWidth        =   2822
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   2831
            MinWidth        =   2822
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   2831
            MinWidth        =   2822
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Object.Width           =   2831
            MinWidth        =   2822
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   2
            Object.Width           =   1058
            MinWidth        =   1058
            Picture         =   "frmInvoice.frx":160F6
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin FLWCtrls.FWLabel LblOrder 
      Height          =   495
      Left            =   2040
      Top             =   480
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   873
      Enabled         =   -1  'True
      Caption         =   "Õ÷Ê—Ì"
      FillType        =   4
      FirstColor      =   8388608
      SecondColor     =   192
      Angle           =   0
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      FontName        =   "Nazanin"
      FontBold        =   -1  'True
      FontSize        =   15.75
      Alignment       =   2
      Picture         =   "frmInvoice.frx":16410
      BorderStyle     =   1
   End
   Begin FLWCtrls.FWLabel fwlblMode 
      Height          =   495
      Left            =   1320
      Top             =   120
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   873
      Enabled         =   -1  'True
      Caption         =   "„—Ê—"
      FillType        =   4
      FirstColor      =   12582912
      SecondColor     =   10070188
      Angle           =   0
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      FontName        =   "Nazanin"
      FontBold        =   -1  'True
      FontSize        =   15.75
      Alignment       =   2
      Picture         =   "frmInvoice.frx":1642C
   End
   Begin SHDocVwCtl.WebBrowser wbsrPrint 
      Height          =   255
      Left            =   480
      TabIndex        =   11
      Top             =   7440
      Visible         =   0   'False
      Width           =   1815
      ExtentX         =   3201
      ExtentY         =   450
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin RESIZEKITLibCtl.ResizeKit ResizeKit1 
      Height          =   480
      Left            =   480
      OleObjectBlob   =   "frmInvoice.frx":16448
      TabIndex        =   14
      Top             =   0
      Width           =   480
   End
   Begin FLWCtrls.FWLed FWLedTemp 
      Height          =   615
      Left            =   6840
      TabIndex        =   52
      Top             =   360
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   1085
      BorderStyle     =   9
      ColorOn         =   192
      ColorOff        =   16777215
      BackColor       =   16777215
   End
   Begin VB.Frame Frame11 
      Caption         =   "Frame11"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4575
      Left            =   3720
      TabIndex        =   54
      Top             =   1080
      Visible         =   0   'False
      Width           =   3015
      Begin VB.TextBox txtTax 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00B9B8C5&
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
         Height          =   375
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   65
         TabStop         =   0   'False
         Top             =   360
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.TextBox txtPacking 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00B9B8C5&
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
         Height          =   375
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   64
         TabStop         =   0   'False
         Top             =   0
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.TextBox txtCarryFeePercent 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00B9B8C5&
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
         Height          =   480
         Left            =   360
         RightToLeft     =   -1  'True
         TabIndex        =   63
         TabStop         =   0   'False
         Top             =   3600
         Visible         =   0   'False
         Width           =   2550
      End
      Begin VB.TextBox txtTaxPercent 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00B9B8C5&
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
         Height          =   480
         Left            =   240
         RightToLeft     =   -1  'True
         TabIndex        =   62
         TabStop         =   0   'False
         Top             =   1920
         Visible         =   0   'False
         Width           =   2550
      End
      Begin VB.TextBox txtPackingPercent 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00B9B8C5&
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
         Height          =   480
         Left            =   180
         RightToLeft     =   -1  'True
         TabIndex        =   61
         TabStop         =   0   'False
         Top             =   1080
         Visible         =   0   'False
         Width           =   2550
      End
      Begin VB.TextBox txtSumFeeTotal 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   480
         RightToLeft     =   -1  'True
         TabIndex        =   60
         TabStop         =   0   'False
         Top             =   3120
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.TextBox txtDiscount 
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
         Height          =   405
         Left            =   195
         RightToLeft     =   -1  'True
         TabIndex        =   59
         TabStop         =   0   'False
         Top             =   3045
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.TextBox txtNo 
         Alignment       =   1  'Right Justify
         BackColor       =   &H000080FF&
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
         Left            =   195
         TabIndex        =   58
         TabStop         =   0   'False
         ToolTipText     =   "ﬂœ „«‘Ì‰ ¬·«  —« œ— »— „ÌêÌ—œ"
         Top             =   2640
         Visible         =   0   'False
         Width           =   2535
      End
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
         Height          =   405
         Left            =   195
         RightToLeft     =   -1  'True
         TabIndex        =   57
         TabStop         =   0   'False
         Top             =   2235
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.TextBox txtDiscountPercent 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00B9B8C5&
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
         Height          =   375
         Left            =   0
         RightToLeft     =   -1  'True
         TabIndex        =   56
         TabStop         =   0   'False
         Top             =   1440
         Visible         =   0   'False
         Width           =   2535
      End
      Begin VB.TextBox txtRecursive 
         Alignment       =   1  'Right Justify
         BackColor       =   &H000080FF&
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
         TabIndex        =   55
         TabStop         =   0   'False
         ToolTipText     =   "ﬂœ „«‘Ì‰ ¬·«  —« œ— »— „ÌêÌ—œ"
         Top             =   720
         Visible         =   0   'False
         Width           =   2535
      End
   End
   Begin MSComctlLib.StatusBar sbrFactorProp 
      Height          =   375
      Left            =   8100
      TabIndex        =   107
      Top             =   9170
      Width           =   7120
      _ExtentX        =   12568
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   2117
            MinWidth        =   2117
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   2302
            MinWidth        =   2293
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   2117
            MinWidth        =   2117
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   2117
            MinWidth        =   2117
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   2117
            MinWidth        =   2117
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label LblLimited 
      Alignment       =   2  'Center
      BackColor       =   &H000000C0&
      Caption         =   "‰”ŒÂ ¬“„«Ì‘Ì"
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
      Left            =   3600
      RightToLeft     =   -1  'True
      TabIndex        =   53
      Top             =   600
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label lblServePlace 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "”«·‰"
      BeginProperty Font 
         Name            =   "Nazanin"
         Size            =   15.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   615
      Left            =   3000
      TabIndex        =   13
      Top             =   120
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   405
      Left            =   0
      TabIndex        =   12
      Top             =   20640
      Visible         =   0   'False
      Width           =   1755
   End
   Begin VB.Label LblInvoice 
      BackStyle       =   0  'Transparent
      Caption         =   "›«ﬂ Ê— ›—Ê‘"
      BeginProperty Font 
         Name            =   "Nazanin"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   495
      Left            =   3720
      TabIndex        =   9
      ToolTipText     =   "«„ﬂ«‰  ⁄—Ì› ”›«—‘«  Ê ÷«Ì⁄«  »« ﬂ·Ìﬂ —ÊÌ «Ì‰ ﬁ”„  ÊÃÊœ œ«—œ"
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label lblDate 
      Alignment       =   2  'Center
      BackColor       =   &H80000013&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   405
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   0
      Width           =   1575
   End
   Begin VB.Label BascoleLabel 
      Alignment       =   2  'Center
      BackColor       =   &H80000013&
      BackStyle       =   0  'Transparent
      Caption         =   " —«“ÊÌ 1 :"
      BeginProperty Font 
         Name            =   "Nazanin"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Index           =   1
      Left            =   7440
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   8520
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Label lblScale 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00CCAD84&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   585
      Index           =   1
      Left            =   6120
      TabIndex        =   3
      Top             =   8400
      Visible         =   0   'False
      Width           =   1425
   End
   Begin VB.Shape ShapeScale 
      BackColor       =   &H80000013&
      FillColor       =   &H80000013&
      FillStyle       =   0  'Solid
      Height          =   690
      Index           =   1
      Left            =   5880
      Shape           =   4  'Rounded Rectangle
      Top             =   8400
      Visible         =   0   'False
      Width           =   2100
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00008000&
      BackStyle       =   1  'Opaque
      FillColor       =   &H0070C0FF&
      FillStyle       =   0  'Solid
      Height          =   900
      Left            =   0
      Shape           =   4  'Rounded Rectangle
      Top             =   120
      Width           =   8085
   End
End
Attribute VB_Name = "frmInvoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'#################   ‘ŒÌ’ ÂÊÌ  ”«“„«‰Ì #############################
Dim WithEvents clsfinger As CheckFingerPrintDll.CFP
Attribute clsfinger.VB_VarHelpID = -1
Dim WithEvents clsfinger2 As CheckFingerPrintDll.CFP
Attribute clsfinger2.VB_VarHelpID = -1
Dim PersonIdqueue As Long
Dim SubFolder As String
'#################  Stimul Printing #############################
Dim StimulPrn As AryaPrinting.StimulPrint
Public MoveToCredit As Boolean
'################# Rfid Reader #############################
Dim RfidReaderIsActive As Boolean
Dim strStream
Dim RFIDStatus As String
'################# MenuBar #############################
Dim lastPosition As Position
Dim MaxBtnMenu As Long
Dim BtnMenuPerFrame As Long
'################# CRM #############################
Dim WithEvents clsdiscount As AryaCRMDiscountCalculator.clsdiscount
Attribute clsdiscount.VB_VarHelpID = -1
Dim GoodItem As AryaCRMDiscountCalculator.GoodView
Dim InvInfo As AryaCRMDiscountCalculator.InvoiceItem
Dim IsLoyaltyCustomer As Boolean
'###################################################
Dim AdminEdit As Boolean
Dim ServeChangeFlag As Boolean
Private Type Position
    x As Single
    y As Single
End Type
Dim Clicked As Position
Dim Moveflag As Boolean
Dim TempAddressEdit As Boolean
Dim TempArrowKeyServe As Boolean
Dim filetemp As New FileSystemObject
Dim AutoDiscountValue As Long
Dim PayClick As Boolean
Dim CallerIdformshow As Boolean
Dim PreReceived As Long
Dim BlnPosResponsed As Boolean
Dim BlnPosApprovedWait As Boolean
Dim PosTempFactorNo As Long
Dim PosTempPrice As Long
Dim ReceiveTypeFlag As Boolean
Dim aaaa As Long  '###################
Dim MainDevice As Boolean
Dim LineNumber As Integer
Dim OldAmount As Long
Dim GoodAmount(30, 1) As String
Dim ViewFlag As Boolean
Dim TempFactorNo As Long
Dim OldCostDifference As Long
Dim Ins As String   'Jahat Daryaft VorodiHay CallerID as Port RS232
Dim AlmPort As Integer
Dim UpdateFromFinalCheck As Boolean
Dim ClsPrint As New Printing
Dim TempCreditCode As Long
Dim textDescriptionFlag As Boolean
Dim textDescription As Boolean
Dim MyFormAddEditMode As EnumAddEditMode
Dim ActionMode As EnumActionLog
Dim mVarOrderType As EnumOrderType
Dim clsDate As New clsDate
Dim ClsCnvKeyBoard As New ClsCnvKeyBoard

Dim DeviceCode(0 To 4) As Integer
Dim RThreshold(1 To 4) As Integer
Dim DeviceType(1 To 4) As Integer
Dim UsbCallerIdIndex As Integer
Dim TimeReaderPort As Integer
Dim intCountGood As Long
Dim rctmp As New ADODB.Recordset
Dim RstTemp As New ADODB.Recordset
Dim Rst As New ADODB.Recordset

Dim mvarEmpty As Boolean
Dim mvarbarcode As Boolean
Public blnCreditCust As Boolean
Dim boolPayment As Boolean
Dim BalancePayment As Boolean
Dim BlnFormLoaded As Boolean

Public MaxRowFlexGrid As Integer

Dim i As Long
Dim intSumOfCurrentServePlaces As Integer

Dim dblFichUser As Double
Dim intSerialNo As Double
Dim dblBasFichNo As Long
Dim mvarCustCredit As Double
Dim intTempFich As Double

Dim Parameter() As Parameter
Dim DefaultServicePercent As Single
Dim DefaultTaxPercent As Single
Dim DropDownFlag As Boolean
Dim mvarStationNo As Integer
Dim mvarEditedFich As Double
Dim MvarUserDefine As Boolean
Dim mvarKeyCode, MvarShiftKey As Integer
Dim Exit_Keypress_Flag As Boolean
Public OldSumPrice As Currency
Dim TmpGoodDiscount As Long
Dim ServiceRate As Double
Dim SplitFlag As Boolean
Dim tmpUserNo As Integer
Dim Repeatbarcode As Integer
Dim ChanceBarcodeQuantity As Integer
Dim TodayFlag As Boolean
Dim IsPrinting As Boolean
Dim EnableDefaultServiceRate As Boolean
Dim Current_PosFacNo As Long
Dim BeforEditInvoice As New ClsInvoice
Dim BeforShowDifferenceFlxRow As Integer
Dim EnableBeforShowDifferenceFlxRow As Boolean
Dim ArrCostDifferences() As Long
Dim cmbTableName As String
Dim cmbTableData As Integer
Dim AddressFlag As Boolean
Dim CustDescriptionFlag As Boolean
Dim textTempAddressFlag As Boolean
Public FindFlag As Boolean
Dim BuyCountTimes1, BuyCountTimes2, BuyCountTimes3 As Integer
Dim RoundDiscount As Double
Public GoodCode As Long

Dim LineDial(8, 8) As String

Const IndexColRow As Integer = 0
Const IndexColAmount As Integer = 1
Const IndexColGoodName As Integer = 2
Const IndexColFee As Integer = 3
Const IndexColTotalFee As Integer = 4
Const IndexColGoodCode As Integer = 5

Const IndexColServePalce As Integer = 8
Const IndexColDifferences As Integer = 10
Const IndexColDiscountPercent As Integer = 11
Const IndexColRate As Integer = 12
Const IndexColChair As Integer = 13
Const IndexColInventory As Integer = 14
Const IndexColLevel1 As Integer = 15
Const IndexColStock As Integer = 16
Const IndexColDuty As Integer = 17
Const IndexColTax As Integer = 18


Private Sub UpdatelblCustomer()
    On Error GoTo ErrHandler
    Dim rctmp2 As New ADODB.Recordset

    If lblCustomer.Tag <> "" Then
        Dim mvarMemberShipId, mvarTel, mvarAddress, mvarDescription As String
        Dim TmpDate1, TmpDate2 As String
     
   '     fwScrollTextCust.Caption = ""
        fwStatusBarCust.Caption = ""
        TxtAddress.Text = ""
        lblCredit.Caption = ""
        lblCredit.BackColor = Me.BackColor
   '     fwScrollTextCust.BackColor = Me.BackColor
        If lblCustomer.Tag <> "-1" Then
            TmpDate1 = Mid(clsDate.shamsi(Date), 3, 6) & "01"
            TmpDate2 = Mid(clsDate.shamsi(Date), 3)
            ReDim Parameter(4) As Parameter
            Parameter(0) = GenerateInputParameter("@Code", adBigInt, 8, Val(lblCustomer.Tag))
            Parameter(1) = GenerateInputParameter("@DateBefore", adVarWChar, 50, CStr(TmpDate1))
            Parameter(2) = GenerateInputParameter("@DateAfter", adVarWChar, 50, CStr(TmpDate2))
            Parameter(3) = GenerateInputParameter("@Branch", adInteger, 4, CurrentBranch)
            Parameter(4) = GenerateOutputParameter("@BuyCountTimes", adInteger, 4)
            BuyCountTimes1 = RunParametricStoredProcedure("BuyCustomerCount", Parameter)
        
            TmpDate1 = Mid(clsDate.shamsi(Date), 3, 3) & "01/01"
            TmpDate2 = Mid(clsDate.shamsi(Date), 3)
            ReDim Parameter(4) As Parameter
            Parameter(0) = GenerateInputParameter("@Code", adBigInt, 8, Val(lblCustomer.Tag))
            Parameter(1) = GenerateInputParameter("@DateBefore", adVarWChar, 50, CStr(TmpDate1))
            Parameter(2) = GenerateInputParameter("@DateAfter", adVarWChar, 50, CStr(TmpDate2))
            Parameter(3) = GenerateInputParameter("@Branch", adInteger, 4, CurrentBranch)
            Parameter(4) = GenerateOutputParameter("@BuyCountTimes", adInteger, 4)
            BuyCountTimes2 = RunParametricStoredProcedure("BuyCustomerCount", Parameter)
            
            If clsArya.MiladiDate = 0 Then
                TmpDate1 = "70/01/01"
            Else
                TmpDate1 = "01/01/01"
            End If
            TmpDate2 = Mid(clsDate.shamsi(Date), 3)
            ReDim Parameter(4) As Parameter
            Parameter(0) = GenerateInputParameter("@Code", adBigInt, 8, Val(lblCustomer.Tag))
            Parameter(1) = GenerateInputParameter("@DateBefore", adVarWChar, 50, CStr(TmpDate1))
            Parameter(2) = GenerateInputParameter("@DateAfter", adVarWChar, 50, CStr(TmpDate2))
            Parameter(3) = GenerateInputParameter("@Branch", adInteger, 4, CurrentBranch)
            Parameter(4) = GenerateOutputParameter("@BuyCountTimes", adInteger, 4)
            BuyCountTimes3 = RunParametricStoredProcedure("BuyCustomerCount", Parameter)
            
        Else
            BuyCountTimes1 = 0
            BuyCountTimes2 = 0
            BuyCountTimes3 = 0
        End If
        If Val(lblCustomer.Tag) <> -1 Then
            ReDim Parameter(1) As Parameter
            Parameter(0) = GenerateInputParameter("@Code", adInteger, 4, Val(lblCustomer.Tag))
            Parameter(1) = GenerateInputParameter("@Branch", adInteger, 4, CurrentBranch)
            Set rctmp2 = RunParametricStoredProcedure2Rec("Get_vw_Customers", Parameter)
            
            If rctmp2.EOF = False And rctmp2.BOF = False Then
                    
                Tafsili = Val(IIf(IsNull(rctmp2!Tafsili), 0, rctmp2!Tafsili))
                If clsArya.ExternalAccounting = True And Tafsili = 0 And Val(rctmp2.Fields("Credit")) > 0 Then
                    ShowMessage "«» œ« »—«Ì «Ì‰ „‘ —Ì œ— ”Ì” „ Õ”«»œ«—Ì  ›÷Ì·Ì «ÌÃ«œ ﬂ‰Ìœ", True, False, " «∆Ìœ", ""
                    lblCustomer.Tag = 0
                    Exit Sub
                End If
                If clsStation.CustomerFeeDataBase = True Then
                    clsStation.CustomerRate = Val(rctmp2.Fields("Sellprice")) - 1
                End If
                
                If clsStation.CustomerRate = 0 Then
                    LblRate.Caption = "‰—Œ «Ê·"
                ElseIf clsStation.CustomerRate = 1 Then
                    LblRate.Caption = "‰—Œ œÊ„"
                ElseIf clsStation.CustomerRate = 2 Then
                    LblRate.Caption = "‰—Œ ”Ê„"
                ElseIf clsStation.CustomerRate = 3 Then
                    LblRate.Caption = "‰—Œ çÂ«—„"
                ElseIf clsStation.CustomerRate = 4 Then
                    LblRate.Caption = "‰—Œ Å‰Ã„"
                ElseIf clsStation.CustomerRate = 5 Then
                    LblRate.Caption = "‰—Œ ‘‘„"
                End If
                
                mvarTel = ""
                If rctmp2.Fields("tel1") <> "" Then
                        mvarTel = " ...  ·›‰ : " + rctmp2.Fields("tel1")
                End If
                If rctmp2.Fields("tel2") <> "" Then
                        mvarTel = mvarTel + " ; " + rctmp2.Fields("tel2")
                End If
                If rctmp2.Fields("FullAddress") <> "" Then
                        mvarAddress = rctmp2.Fields("FullAddress")
                End If
                txtDiscountPercent = rctmp2.Fields("Discount")
                If mvarServePlace = Delivery Then
                    txtCarryFee.Text = rctmp2.Fields("CarryFee")
                    lblCarryFeeTotal = rctmp2.Fields("CarryFee")
                Else
                    txtCarryFee.Text = 0
                    lblCarryFeeTotal = 0
                End If
                lblCustomer.Caption = rctmp2.Fields("FullName")
                mvarCustCredit = rctmp2.Fields("Credit")
                mvarDescription = rctmp2.Fields("Description")
                
                If strCategory = "07" Then 'Bank Of Tejarat
                    mvarDescription = ""
                    mvarDescription = "     Õ   ﬂ›· = " & rctmp2.Fields("FamilyNo").Value & " ‰›—  " & mvarDescription
                    If rctmp2.Fields("Member") = True Then
                        mvarMemberShipId = "⁄÷Ê : " & rctmp2.Fields("MemberShipId") '& " " & rctmp2.Fields("State")
                    Else
                        mvarMemberShipId = "€Ì— ⁄÷Ê : " & rctmp2.Fields("MemberShipId") '& " " & rctmp2.Fields("State")
                    End If
                Else
                    If clsStation.Language = Farsi Then
                        mvarMemberShipId = "«‘ —«ﬂ : " & rctmp2.Fields("MemberShipId")
                    Else
                        mvarMemberShipId = "Customer Id :" & rctmp2.Fields("MemberShipId")
                    End If
                End If
                
''''                mvarDescription = mvarDescription & "  œ›⁄«  Œ—Ìœ œ— «Ì‰ „«Â = " & BuyCountTimes1 & " »«—"
''''                mvarDescription = mvarDescription & " - œ— «„”«· = " & BuyCountTimes2 & " »«—"
''''                mvarDescription = mvarDescription & " - œ— ﬂ· = " & BuyCountTimes3 & " »«—"
''''                mvarDescription = "  œ›⁄«  Œ—Ìœ œ— «Ì‰ „«Â = " & BuyCountTimes1 & " »«—" & " - œ— «„”«· = " & BuyCountTimes2 & " »«—" & " - œ— ﬂ· = " & BuyCountTimes3 & " »«—    " & mvarDescription
''                If rctmp2.Fields("Code") <> -1 And clsStation.ViewTempAddress = -1 Then
                  If rctmp2.Fields("Code") <> -1 Then
                    If clsStation.ViewTempAddress = -1 Then
                        TxtTempAddress.Visible = False
                        TxtCustDescription.Visible = True
                    Else
                        TxtTempAddress.Visible = True
                        TxtCustDescription.Visible = False
                     End If
                    TxtAddress.Visible = True
                    TxtCustDescription.Text = mvarDescription
                   ' fwScrollTextCust.Caption = mvarDescription
                    fwStatusBarCust.Caption = mvarMemberShipId & mvarTel
                    TxtAddress.Text = mvarAddress
                    If mvarDescription <> "" And Me.txtRecursive <> 1 And mvarEditedFich <> 1 Then
                     '  fwScrollTextCust.Visible = True
                    Else
                     '  fwScrollTextCust.Visible = False
                    End If
                 '   fwStatusBarCust.Visible = True
                    lblCredit.Visible = True
                    blnCreditCust = IIf(rctmp2!Credit > 0, True, False)
                    If clsArya.ExternalAccounting = False Then
                        If blnCreditCust And clsArya.ExternalAccounting = False Then
                            If MyFormAddEditMode = EditMode Or MyFormAddEditMode = ManipulateMode Then
                                If clsStation.CreditCalculate = True Then
                                    mvarCustCredit = rctmp2!Credit + rctmp2!Bestankar - rctmp2!Price + lblSumPrice.Tag
                                Else
                                    mvarCustCredit = rctmp2!Bestankar - rctmp2!Price + lblSumPrice.Tag
                                End If
                            Else
                                If clsStation.CreditCalculate = True Then
                                    mvarCustCredit = rctmp2!Credit + rctmp2!Bestankar - rctmp2!Price
                                Else
                                    mvarCustCredit = rctmp2!Bestankar - rctmp2!Price
                                End If
                            End If
                            If mvarCustCredit < 0 Then
                               ' fwScrollTextCust.Visible = True
                              '  fwScrollTextCust.Caption = fwScrollTextCust.Caption & "  -  »œÂÌ œ«—œ : œﬁ  ‘Êœ"
                                lblCredit.ForeColor = vbRed
                                If clsStation.CreditCalculate = True Then
                                    lblCredit.Caption = " »œÂÌ »« «Õ ”«» «⁄ »«—: " & Abs(mvarCustCredit)
                                Else
                                    lblCredit.Caption = " »œÂÌ :     " & Abs(mvarCustCredit)
                                End If
                            ElseIf mvarCustCredit > 0 Then
                                lblCredit.ForeColor = 0
                                lblCredit.BackColor = Me.BackColor
                                If clsStation.CreditCalculate = True Then
                                    lblCredit.Caption = " «⁄ »«— »«ﬁÌ„«‰œÂ :   " & mvarCustCredit
                                Else
                                    lblCredit.Caption = " ÿ·»   :   " & mvarCustCredit
                                End If
                            End If
                            If clsArya.Accounting = True Then
                                'lblCredit.ForeColor = &H8000&  ' &HC0C0C0
                                'lblCredit.Caption = "        €Ì— ‰ﬁœÌ Ê «⁄ »«—Ì "
                                If FWScrolltextPay.Caption = " ”ÊÌÂ ‰‘œÂ" Then
                                    FWScrolltextPay.Caption = " €Ì— ‰ﬁœÌ "
                                End If
                            End If
                        Else
                            lblCredit.BackColor = Me.BackColor
                            mvarCustCredit = 0
                            lblCredit.Caption = ""
                        End If
                    End If
                
                    '' After Change Rate
                    If MyFormAddEditMode = AddMode Or MyFormAddEditMode = EditMode Then RateChanged True
                Else
                    If strCategory = "07" Then
                        lblCustomer.Caption = "€Ì— ⁄÷Ê"
                    End If
                    blnCreditCust = False
                 '   fwScrollTextCust.Caption = ""
                    fwStatusBarCust.Caption = ""
                    TxtAddress.Text = "¬œ—” : "
                    TxtCustDescription.Text = ""
                 '   fwScrollTextCust.Visible = False
                  '  fwStatusBarCust.Visible = False
                    TxtTempAddress.Visible = True
                    TxtCustDescription.Visible = False
                   ' TxtAddress.Visible = False
                    lblCredit.Visible = False
                End If
                
                
            End If
        Else
             mvarTel = ""
             mvarAddress = ""
'''
'''             txtDiscountPercent = 0
'''             txtCarryFee.Text = 0
'''             lblCarryFeeTotal = 0
             If clsStation.Language = Farsi Then
                lblCustomer.Caption = "€Ì— „‘ —ò"
             Else
                lblCustomer.Caption = "Non-customer"
             End If
             mvarCustCredit = 0
             mvarDescription = ""
             TxtTempAddress.Visible = True
             TxtCustDescription.Visible = False
             blnCreditCust = False
          '   fwScrollTextCust.Caption = ""
             fwStatusBarCust.Caption = ""
             TxtAddress.Text = "¬œ—” : "
             TxtCustDescription.Text = ""
          '   fwScrollTextCust.Visible = False
          '   fwStatusBarCust.Visible = False
             lblCredit.Visible = False
            
        End If
        
        Set rctmp2 = Nothing
    End If

Exit Sub

ErrHandler:
    LogSaveNew "frmInvoice => ", err.Description, err.Number, err.Source, "UpdateLbLCustomer"
'    MsgBox err.Description
    Resume Next
End Sub

Private Sub BascoleLabel_Click(index As Integer)
    If (lblScale(index).Caption <> txtScale.Text) Then
        lblScale(index).Enabled = IIf(lblScale(index).Enabled, False, True)
    End If
End Sub

Private Sub BtnKalaDelete_GotFocus()
    On Error Resume Next
    FlxDetail.SetFocus
End Sub


Private Sub BtnKeypad_Click(index As Integer)
    Select Case BtnKeypad(index).Tag
        Case "1", "2", "3", "4", "5", "6", "7", "8", "9", "0":
            lblNum.Caption = lblNum.Caption + BtnKeypad(index).Tag
            If Len(lblNum.Caption) > 12 Then
               lblNum.Caption = ""
            End If
        Case "%":
         ' If BtnKeypad(11).Enabled Then
             If lblNum.Caption <> "" Then
                lblNum.Caption = lblNum.Caption + BtnKeypad(index).Tag
'                BtnKeypad(11).Enabled = False     '"%"
            End If
         '  End If
        Case ".":
            lblNum.Caption = lblNum.Caption + BtnKeypad(index).Tag
'            BtnKeypad(10).Enabled = False      '"."
    
        Case "-":
            If Len(Trim(lblNum.Caption)) >= 1 Then
                lblNum.Caption = left(lblNum.Caption, Len(Trim(lblNum.Caption)) - 1)
            End If
    End Select
End Sub

Private Sub UpdateLock(invoiceNumber As Long, CurrentBranch As Integer, BitLock As Boolean)
    
    On Error GoTo ErrHandler
    ReDim Parameter(4) As Parameter
    Parameter(0) = GenerateInputParameter("@invoiceNO", adBigInt, 8, Val(txtNo.Text))
    Parameter(1) = GenerateInputParameter("@Status", adInteger, 4, mvarStatus)
    Parameter(2) = GenerateInputParameter("@AccountYear", adSmallInt, 2, AccountYear)
    Parameter(3) = GenerateInputParameter("@branch", adInteger, 4, CurrentBranch)
    Parameter(4) = GenerateInputParameter("@bitLock", adBoolean, 1, BitLock)
    RunParametricStoredProcedure "Update_bitLock", Parameter
     
    Exit Sub
ErrHandler:
    MsgBox err.Description
    modgl.LogSave "frmInvoice ", err, "UpdateLock"
End Sub

Private Sub ChkIsLocked_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If MyFormAddEditMode = AddMode Then Exit Sub
    If ClsFormAccess.LockCheck = False Then
        frmMsg.fwlblMsg.Caption = " «„ﬂ«‰ œ” —”Ì »Â «Ì‰ «„ﬂ«‰ »—«Ì ‘„« ÊÃÊœ ‰œ«—œ "
        frmMsg.fwBtn(0).Visible = True
        frmMsg.fwBtn(0).Caption = "ﬁ»Ê·"
        frmMsg.fwBtn(1).Visible = False
        frmMsg.Show vbModal
        Exit Sub
    End If
    
    UpdateLock Val(txtNo.Text), CurrentBranch, ChkIsLocked.Value
    If ChkIsLocked.Value = 1 Then
        frmDisMsg.lblMessage.Caption = "”‰œ ﬁ›· ‘œ."
    Else
        frmDisMsg.lblMessage.Caption = "”‰œ «“ Õ«·  ﬁ›· Œ«—Ã ‘œ."
    End If
    frmDisMsg.Timer1.Interval = 1500
    frmDisMsg.Timer1.Enabled = True
    frmDisMsg.Show vbModal
    Exit Sub

End Sub

Private Sub chKTax_Click()
Dim ii As Long
With FlxDetail
    For ii = 1 To MaxRowFlexGrid - 1
        .TextMatrix(ii, IndexColTax) = chKTax.Value
        .TextMatrix(ii, IndexColDuty) = chKTax.Value
    Next
End With
RefreshLables

End Sub


Private Sub cmbServePlace_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub


Private Sub CmdColor_Click()
    If IsFarabin = True Then If MaxRowFlexGrid > 1 Then ShowMonitor 1
    
    frmColor.Show vbModal

    'OpenCashDrawer
End Sub

Private Sub cmdTurnOver_Click()
    If ClsFormAccess.AccfrmKartHesabReport = True Then
        If Tafsili > 0 Then
            Accounting.KartHesabShowDll "KolBedehkaran", Val(Tafsili), lblCustomer.Caption, Right(AccountYear, 2) & "/01/01", mvarDate
        End If
    Else
        ShowDisMessage "‘„« »Â «Ì‰ «„ﬂ«‰ œ” —”Ì ‰œ«—Ìœ", 1500
    End If
End Sub

Private Sub cmbServePlace_Click()
        If TempArrowKeyServe = True Then TempArrowKeyServe = False:   Exit Sub
        If MyFormAddEditMode = ViewMode Then Exit Sub
        Dim OldServePlace As Integer
        OldServePlace = mvarServePlace
        mvarServePlace = cmbServePlace.ItemData(cmbServePlace.ListIndex)
        If mvarServePlace = OldServePlace Then Exit Sub
        If mvarServePlace = Out Then
            clsStation.PriceType = clsStation.OutPrice
        Else
            clsStation.PriceType = MainPriceType
        End If
        
        If clsStation.PriceType = 1 Then
            LblRate.Caption = "‰—Œ «Ê·"
            mvarRate = 1
        ElseIf clsStation.PriceType = 2 Then
            LblRate.Caption = "‰—Œ œÊ„"
            mvarRate = 2
        ElseIf clsStation.PriceType = 3 Then
            LblRate.Caption = "‰—Œ ”Ê„"
            mvarRate = 3
        ElseIf clsStation.PriceType = 4 Then
            LblRate.Caption = "‰—Œ çÂ«—„"
            mvarRate = 4
        ElseIf clsStation.PriceType = 5 Then
            LblRate.Caption = "‰—Œ Å‰Ã„"
            mvarRate = 5
        ElseIf clsStation.PriceType = 6 Then
            LblRate.Caption = "‰—Œ ‘‘„"
            mvarRate = 6
        End If
        
        UpdatelblServePlace
        FlxDetail.ColHidden(8) = False
        mvarMsgIdx = vbNo
        If MaxRowFlexGrid > 1 Then ShowMessage "¬Ì« „«Ì·Ìœ Â„Â „Ê«—œ »Â «Ì‰ Õ«·  ”—Ê  »œÌ· ‘Ê‰œø", True, True, "»·Ì", "ŒÌ—"
        If mvarMsgIdx = vbYes Then
          If clsStation.HasOptionPrice = False Then
            For i = 1 To FlxDetail.Rows - 1
                If FlxDetail.TextMatrix(i, 5) <> "" Then
                    ReDim Parameter(4) As Parameter
                    Parameter(0) = GenerateInputParameter("@Code", adInteger, 4, FlxDetail.ValueMatrix(i, 5))
                    Parameter(1) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
                    Parameter(2) = GenerateInputParameter("@StationId", adInteger, 4, clsArya.StationNo)
                    Parameter(3) = GenerateInputParameter("@Flag", adBoolean, 1, 1)
                    Parameter(4) = GenerateInputParameter("@AccountYear", adSmallInt, 2, AccountYear)
                    Set rctmp = RunParametricStoredProcedure2Rec("Get_Good_Code", Parameter)
                    If Not (rctmp.BOF Or rctmp.EOF) Then
                        FlxDetail.TextMatrix(i, 12) = mvarRate
                        FlxDetail.TextMatrix(i, 8) = mvarServePlace
                        If clsStation.PriceType = 1 Then
                           FlxDetail.TextMatrix(i, 3) = rctmp.Fields("SellPrice").Value
                        ElseIf clsStation.PriceType = 2 Then
                           FlxDetail.TextMatrix(i, 3) = rctmp.Fields("SellPrice2").Value
                        ElseIf clsStation.PriceType = 3 Then
                           FlxDetail.TextMatrix(i, 3) = rctmp.Fields("SellPrice3").Value
                        ElseIf clsStation.PriceType = 4 Then
                           FlxDetail.TextMatrix(i, 3) = rctmp.Fields("SellPrice4").Value
                        ElseIf clsStation.PriceType = 5 Then
                           FlxDetail.TextMatrix(i, 3) = rctmp.Fields("SellPrice5").Value
                        ElseIf clsStation.PriceType = 6 Then
                           FlxDetail.TextMatrix(i, 3) = rctmp.Fields("SellPrice6").Value
                        End If
                    End If
                End If
            Next
          End If
          intSumOfCurrentServePlaces = CalculateSumOfServeplace
        End If
        RefreshLables
       ''' FlxDetail.SetFocus
        TempArrowKeyServe = False
End Sub

Private Sub cmbServePlace_KeyDown(KeyCode As Integer, Shift As Integer)
    Debug.Print KeyCode
    If Shift = 0 And (KeyCode = 40 Or KeyCode = 38) Then TempArrowKeyServe = True
    If Shift = 0 And KeyCode = 13 Then cmbServePlace_Click
'    KeyCode = 0
End Sub

Private Sub cmdClose_Click()
    FrameCustInfo.Visible = False
End Sub

Private Sub cmdDeleteCustomer_Click()
    lblCustomer.Tag = -1
    clsStation.CustomerRate = 0
    mvarPublicOrderType = inPerson
    mVarOrderType = inPerson
    If clsStation.Language = Farsi Then
       LblOrder.Caption = "Õ÷Ê—Ì"
    Else
        LblOrder.Caption = "Inside"
    End If
    
    mvarServePlace = clsStation.ServePlaceDefault
    If mvarServePlace = EnumServePlace.Table Or mvarServePlace = EnumServePlace.Salon Then
        ServiceRate = DefaultServicePercent
    Else
        ServiceRate = 0
    End If
    For i = 0 To cmbServePlace.ListCount - 1
        If mvarServePlace = cmbServePlace.ItemData(i) Then
            cmbServePlace.ListIndex = i
            Exit For
        End If
    Next i
    UpdatelblCustomer
    UpdatelblServePlace
    RefreshLables
    FrameCustInfo.Visible = False
End Sub


Private Sub FrameBascule_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Clicked.x = x
    Clicked.y = y
    Moveflag = True
End Sub

Private Sub FrameBascule_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Moveflag = False Then Exit Sub
    If Clicked.x <> 0 Or Clicked.y <> 0 Then
        FrameBascule.left = FrameBascule.left + (x - Clicked.x)
        FrameBascule.top = FrameBascule.top + (y - Clicked.y)
        Clicked.x = x
        Clicked.y = y
    End If
End Sub

Private Sub FrameBascule_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    SaveSetting strMainKey, "FrameBascule", "Left", FrameBascule.left
    SaveSetting strMainKey, "FrameBascule", "Top", FrameBascule.top
    Moveflag = False
End Sub



Private Sub FWScrolltextPay_DblClick()
    If MyFormAddEditMode = AddMode Then Exit Sub
    If ClsFormAccess.UnBalance = False Then
        frmMsg.fwlblMsg.Caption = " «„ﬂ«‰ œ” —”Ì »Â «Ì‰ «„ﬂ«‰ »—«Ì ‘„« ÊÃÊœ ‰œ«—œ "
        frmMsg.fwBtn(0).Visible = True
        frmMsg.fwBtn(0).Caption = "ﬁ»Ê·"
        frmMsg.fwBtn(1).Visible = False
        frmMsg.Show vbModal
        Exit Sub
    End If
    If FWScrolltextPay.BackColor = vbGreen Then
        frmMsg.fwlblMsg.Caption = "¬Ì« ‰”»  »Â Õ–› œ—Ì«› Â«Ì «‰Ã«„ ‘œÂ —ÊÌ «Ì‰ ›Ì‘ „ÿ„∆‰ Â” Ìœø "
        frmMsg.fwBtn(0).ButtonType = flwButtonOk
        frmMsg.fwBtn(0).Caption = "»·Ì"
        frmMsg.fwBtn(1).Visible = flwButtonCancel
        frmMsg.fwBtn(1).Caption = "ŒÌ—"
        frmMsg.fwBtn(1).Default = True
        frmMsg.Show vbModal
        If mvarMsgIdx = vbYes Then
            ReDim Parameter(1) As Parameter
            Parameter(0) = GenerateInputParameter("@intSerialNo", adBigInt, 8, intSerialNo)
            Parameter(1) = GenerateInputParameter("@branch", adInteger, 4, CurrentBranch)
            RunParametricStoredProcedure "Update_tfacM_UnBalance", Parameter
            ShowDisMessage "Õ–› œ—Ì«›  Â« «‰Ã«„ ‘œ.", 1000
            GetDataDetail
            RefreshLables
        End If
    End If

End Sub

Private Sub Timer_PersonIdCheck_Timer()
If lblCustomer.Tag <> "-1" Then
     Timer_PersonIdCheck.Enabled = False
     GetFirstPersonFromList
End If
End Sub

Private Sub Timer_Printers_Timer()
    
    Dim PrinterStatus As String
    Dim JobStatus As String
    Dim JobQuantity As Long
    Dim ErrorInfo As String
    Dim ii As Long
    For ii = 1 To 6
        If Val(PrinterNo(ii)) > 0 Then
   
            'Clear the status info for new info/status.
            PrinterShape(ii - 1).FillColor = vbGreen
            lblPrinter(ii - 1) = 0
            'Call sub to perform check.
            Call CheckPrinter(PrinterName(ii), PrinterStatus, JobStatus, JobQuantity)
            'Text2.Text = PrinterStatus
            'Text3.Text = JobStatus
            lblPrinter(ii - 1) = JobQuantity
            If JobQuantity > 0 Then
                PrinterShape(ii - 1).FillColor = vbRed
            End If
        Else
            lblPrinter(ii - 1).Visible = False
            PrinterShape(ii - 1).Visible = False
        End If
    Next

End Sub

Private Sub TxtGuestNo_Change()
    lblNum.Caption = ""
End Sub

Private Sub TxtGuestNo_GotFocus()
    TxtGuestNo.Text = lblNum.Caption
End Sub


Private Sub TxtTempAddress_KeyPress(KeyAscii As Integer)
    TempAddressEdit = True
End Sub

Private Sub UCCallerIDMonitor1_CallerIDDetect(Line As Byte, Number As String)
    
On Error GoTo Err_Handler
    Call mdifrm.FWMMedia1.PlayWaveFile(App.Path & "\Sound\Ringin.wav", True, False)
    FWModem(Line - 1).BackColor = vbRed ' &H80000003&
    FWModem(Line - 1).ToolTipText = Number 'Left(LTrim(Mid(Inputstr, jj + 9)), 15)
    
    DeviceCode(0) = EnumDevice.CallerIdInterface2_AlmP6
    GetCallerInfo 0, Number, Line

Exit Sub
Err_Handler:
    ShowDisMessage err.Description, 1500
End Sub

Private Sub ChkCallerId_Click()
Dim lR As Long
    If ChkCallerId.Value = True Then
        frmCallerIdView.Show vbModal
        ChkCallerId.Value = False
'        lR = SetTopMostWindow(frmCallerIdView.hwnd, True)
    End If

End Sub

Private Sub cmbGarson_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim IsInList As Boolean
    Dim intLenght As Integer
    
    If cmbGarson.Text = "" Then Exit Sub       ' Or KeyCode = 8
    
    For i = 0 To cmbGarson.ListCount - 1
        If cmbGarson.List(i) Like cmbGarson.Text & "*" Then
            intLenght = Len(cmbGarson.Text)
            cmbGarson.ListIndex = i
         '   cmbGarson.SelStart = intLenght
         '   cmbGarson.SelLength = Len(cmbGarson.Text) - intLenght
            
            IsInList = True
            Exit For
        End If
        
    Next i
    
    If IsInList = False Then
        KeyCode = 0
        cmbGarson.Text = Mid(cmbGarson.Text, 1, Len(cmbGarson.Text) - 1)
        cmbGarson.SetFocus
        SendKeys "{Left}"
    End If
    lblNum = ""
End Sub

Private Sub cmdPayFactor_Click()
     'Case "œ—Ì«› ":
If MyFormAddEditMode <> ViewMode Then

    If Val(lblNum.Caption) < 0 Then
        ShowMessage "œ—Ì«›  „‰›Ì ﬁ«»· ﬁ»Ê· ‰Ì”  ", True, False, "ﬁ»Ê·", ""
        lblNum.Caption = 0
        Exit Sub
    End If
''''    If blnCreditCust = False Then
''''        frmMsg.fwlblMsg.Caption = " . „‘ —ﬂ €Ì— «⁄ »«—Ì «”  ‰„Ì  Ê«‰Ìœ œ—Ì«›  ﬂ‰Ìœ "
''''        frmMsg.fwBtn(1).Visible = False
''''        frmMsg.Show vbModal
''''        lblNum.Caption = 0
''''        Exit Sub
''''    End If
    
    If Right(lblNum.Caption, 1) <> "%" Then
        lblPayFactorTotal = Val(lblNum.Caption)
    Else
        ShowMessage "œ—Ì«›  œ—’œÌ ﬁ«»· ﬁ»Ê· ‰Ì”  ", True, False, "ﬁ»Ê·", ""
        lblNum.Caption = 0
        Exit Sub
    End If
    
    BtnKeypad(10).Enabled = True
    BtnKalaDelete.Enabled = True
    BtnKeypad(11).Enabled = True
    'BtnKalaDelete.ForeColor = &H404080
Else
End If
    lblNum.Caption = ""
    FlxDetail.SetFocus
End Sub

Private Sub cmbTable_Click()
    Dim rctmp2 As New ADODB.Recordset
    
    If cmbTable.ListIndex = -1 Then Exit Sub
    
    If (cmbGarson.ListIndex = -1 Or cmbGarson.ListIndex = 0) Then
        ReDim Parameter(0) As Parameter
        Parameter(0) = GenerateInputParameter("@No", adInteger, 4, cmbTable.ItemData(cmbTable.ListIndex))
        Set rctmp2 = RunParametricStoredProcedure2Rec("Get_Table_Incharge", Parameter)
        
        If Not (rctmp2.EOF = True And rctmp2.BOF = True) Then
            For i = 1 To cmbGarson.ListCount - 1
                If rctmp2.Fields("Person").Value = cmbGarson.ItemData(i) Then
                    cmbGarson.ListIndex = i
                    Exit For
                End If
            Next i
        Else
            cmbGarson.ListIndex = 0
        End If
    End If
    Dim L_Rst As New ADODB.Recordset
    If cmbTable.ListIndex <> 0 Then
         ReDim Parameter(1) As Parameter
         Parameter(0) = GenerateInputParameter("@Branch", adInteger, 4, CurrentBranch)
         Parameter(1) = GenerateInputParameter("@TableControl", adBoolean, 1, 0)
         Set L_Rst = RunParametricStoredProcedure2Rec("RetriveTable_Branch", Parameter)
         
        
         If Not (L_Rst.EOF = True And L_Rst.BOF = True) Then
             Do While L_Rst.EOF <> True
                 If cmbTable.ItemData(cmbTable.ListIndex) = L_Rst!No Then
                    If MyFormAddEditMode <> EditMode Then
                        ServiceRate = L_Rst!DefaultServicePercent
                        RefreshLables
                    End If
                    Exit Do
                 End If
                 L_Rst.MoveNext
             Loop
         End If
         
         L_Rst.Close: Set L_Rst = Nothing
        
        mvarServePlace = EnumServePlace.Table
        For i = 0 To cmbServePlace.ListCount - 1
            If mvarServePlace = cmbServePlace.ItemData(i) Then
                cmbServePlace.ListIndex = i
                Exit For
            End If
        Next i
     '''''   UpdatelblServePlace
'        If ServiceRate = 0 Then
'             ServiceRate = DefaultServicePercent
'        End If
        
    End If
End Sub

Private Sub cmbTable_KeyUp(KeyCode As Integer, Shift As Integer)
    
    lblNum = ""
    If MyFormAddEditMode = ViewMode Then Exit Sub
'    FillsTableCombo
    
    Dim IsInList As Boolean
    Dim intLenght As Integer
    
    If cmbTable.Text = "" Then Exit Sub  ' Or KeyCode = 8
    
'    For i = 0 To cmbTable.ListCount - 1
'        If cmbTable.List(i) Like cmbTable.Text & "*" Then
'            intLenght = Len(cmbTable.Text)
'            cmbTable.ListIndex = i
'            cmbTable.SelStart = intLenght
'            cmbTable.SelLength = Len(cmbTable.Text) - intLenght
'
'            IsInList = True
'            Exit For
'        End If
'
'    Next i

'    If IsInList = False Then
'        'KeyCode = 0
'        cmbTable.Text = Mid(cmbTable.Text, 1, Len(cmbTable.Text) - 1)
'        cmbTable.SetFocus
'        SendKeys "{Left}"
'    End If
'    If mvarServePlace = Table Then
'        Dim strSelectedTable As String
'        strSelectedTable = cmbTable.Text
'        IsInList = False
'
'        For i = 0 To cmbTable.ListCount - 1
'            If strSelectedTable = cmbTable.List(i) Then
'                IsInList = True
'            End If
'        Next i
'
'        If IsInList = False Then
'            ShowMessage "„Ì“ œ— ·Ì”  ÊÃÊœ ‰œ«—œ. ·ÿ›« Ìﬂ „Ì“ „⁄ »— «‰ Œ«» ﬂ‰Ìœ", True, False, " «ÌÌœ", ""
'            cmbTable.Text = ""
'            cmbTable.ListIndex = -1
'        End If
'    End If
End Sub

Private Sub cmbTable_Validate(Cancel As Boolean)
    Dim IsInList As Boolean
    IsInList = False
    
    If mvarServePlace = Table Then
        Dim strSelectedTable As String
        strSelectedTable = cmbTable.Text
        IsInList = False
        
        For i = 0 To cmbTable.ListCount - 1
            If strSelectedTable = cmbTable.List(i) Then
                IsInList = True
            End If
        Next i
        
        If IsInList = False Then
            ShowMessage "„Ì“ œ— ·Ì”  ÊÃÊœ ‰œ«—œ. ·ÿ›« Ìﬂ „Ì“ „⁄ »— «‰ Œ«» ﬂ‰Ìœ", True, False, " «ÌÌœ", ""
            cmbTable.Text = ""
            cmbTable.ListIndex = -1
            Cancel = True
        'End If
        Else
            Cancel = False
        End If
    End If
End Sub


Private Sub CmdPager_Click()
    
If clsStation.PersonIdCheck = True Then
     
    ' clsfinger.ShowForm

Else
    If clsStation.Pager = False Then ShowDisMessage " ‰ŸÌ„«  ÅÌÃ— «‰Ã«„ ‰‘œÂ", 1500: Exit Sub
    MousePointer = 11
    DoEvents
    PagerAction Val(lblNum.Caption)

    lblNum.Caption = ""
    MousePointer = 0
End If
End Sub

Private Sub PagerAction(No As Integer)
    
On Error GoTo ErrHandler

    If No <= 0 Then
        If clsStation.SoundAlarm = True Then
    '        Call mdifrm.FWMMedia1.PlayWaveFile(App.Path & "\Sound\winAquariumError.wav", True, False)
            Call mdifrm.FWMMedia1.PlayWaveFile(App.Path & "\Pager\Notify.wav", True, False)
        End If
        Exit Sub
    End If
    
    PagerNo = No
    If PagerNo < 1000 Then frmPager.UpdateNumber
    
    DoEvents
''''
    Dim i, j As Integer
    i = No Mod 10
    j = No Mod 100

    Select Case No
        Case Is >= 1000
            Call mdifrm.FWMMedia1.PlayWaveFile(App.Path & "\Pager\" & No & ".wav", False, False)

        Case Is < 100
            If No < 21 Or i = 0 Then
                Call mdifrm.FWMMedia1.PlayWaveFile(App.Path & "\Pager\" & "Number.wav", False, False)
                Call mdifrm.FWMMedia1.PlayWaveFile(App.Path & "\Pager\" & No & ".wav", False, False)
            Else
                Call mdifrm.FWMMedia1.PlayWaveFile(App.Path & "\Pager\" & "Number.wav", False, False)
                Call mdifrm.FWMMedia1.PlayWaveFile(App.Path & "\Pager\" & (No - i) / 10 & "x.wav", False, False)
                Call mdifrm.FWMMedia1.PlayWaveFile(App.Path & "\Pager\" & i & ".wav", False, False)
            End If
        Case Else
            If j = 0 Then
                Call mdifrm.FWMMedia1.PlayWaveFile(App.Path & "\Pager\" & "Number.wav", False, False)
                Call mdifrm.FWMMedia1.PlayWaveFile(App.Path & "\Pager\" & No & ".wav", False, False)
            ElseIf j < 21 Or i = 0 Then
                Call mdifrm.FWMMedia1.PlayWaveFile(App.Path & "\Pager\" & "Number.wav", False, False)
                Call mdifrm.FWMMedia1.PlayWaveFile(App.Path & "\Pager\" & (No - j) / 100 & "xx.wav", False, False)
                Call mdifrm.FWMMedia1.PlayWaveFile(App.Path & "\Pager\" & j & ".wav", False, False)
            Else
                Call mdifrm.FWMMedia1.PlayWaveFile(App.Path & "\Pager\" & "Number.wav", False, False)
                Call mdifrm.FWMMedia1.PlayWaveFile(App.Path & "\Pager\" & (No - j) / 100 & "xx.wav", False, False)
                Call mdifrm.FWMMedia1.PlayWaveFile(App.Path & "\Pager\" & (j - i) / 10 & "x.wav", False, False)
                Call mdifrm.FWMMedia1.PlayWaveFile(App.Path & "\Pager\" & i & ".wav", False, False)
            End If
    End Select
''''''''''''''
Exit Sub
ErrHandler:
''    MsgBox err.Description & "  =  " & App.Path & "\Pager\Pager.exe "
    MsgBox err.Description & "  =  ÅÌÃ— ›—«ŒÊ«‰"
End Sub

Private Sub cmdPay_Click()
    Dim Balance As Integer
    Dim temp As Long
    Dim Price As Long
    
    If BlnPosApprovedWait = True Then
        ShowDisMessage "”Ì” „ „‰ Ÿ— œ—Ì«›  Å«”Œ «“ ŒÊœÅ—œ«“ „Ì »«‘œ", 1000
        Exit Sub
    End If
    If MyFormAddEditMode <> ViewMode And mvarStatus = Order Then Exit Sub
 
    ReceiveTypeFlag = False
    If MyFormAddEditMode = ViewMode And mvarStatus = Order Then
        ReDim Parameter(4) As Parameter
        Parameter(0) = GenerateInputParameter("@No", adBigInt, 8, Val(txtNo.Text))
        Parameter(1) = GenerateInputParameter("@Status", adInteger, 4, mvarStatus)
        Parameter(2) = GenerateInputParameter("@AccountYear", adSmallInt, 2, AccountYear)
        Parameter(3) = GenerateInputParameter("@Branch", adInteger, 4, CurrentBranch)
        Parameter(4) = GenerateOutputParameter("@Balance", adInteger, 4)

        Balance = RunParametricStoredProcedure("Get_tfacm_Balance", Parameter)
        If Balance = 1 Then
        frmDisMsg.lblMessage = " ”›«—‘ ﬁ»·«  ÕÊÌ· ‘œÂ "
        frmDisMsg.Timer1.Enabled = True
        frmDisMsg.Show vbModal
        Exit Sub
       End If
       OldSumPrice = Val(lblPayFactorTotal.Caption)
       MyFormAddEditMode = EditMode
       OrderNo = Val(txtNo.Text)
       
'        frmMsg.fwlblMsg.Caption = "¬Ì« ”›«—‘  ”ÊÌÂ „Ì ‘Êœø "
'        frmMsg.fwBtn(0).ButtonType = flwButtonOk
'        frmMsg.fwBtn(0).Caption = "»·Ì"
'        frmMsg.fwBtn(1).Visible = flwButtonCancel
'        frmMsg.fwBtn(1).Caption = "ŒÌ—"
'        frmMsg.fwBtn(1).Default = True
'        frmMsg.Show vbModal
'        If mvarMsgIdx = vbYes Then
            frmFactorReceived.FWBtnPrint.Visible = True
            frmFactorReceived.Show vbModal
'            ReDim Parameter(4) As Parameter
'            Parameter(0) = GenerateInputParameter("@No", adBigInt, 8, OrderNo)
'            Parameter(1) = GenerateInputParameter("@Status", adInteger, 4, EnumFactorType.Order)
'            Parameter(2) = GenerateInputParameter("@UID", adInteger, 4, mvarCurUserNo)
'            Parameter(3) = GenerateInputParameter("@AccountYear", adSmallInt, 2, AccountYear)
'            Parameter(4) = GenerateInputParameter("@ds", adWChar, 4000, sFactorReceived)
'
'            RunParametricStoredProcedure "Update_tfacm_Balance", Parameter
'        Else
'            sFactorReceived = ""
'        End If
          Dim st As String
          DetailsString1 = ""
          DetailsString2 = ""
          DetailsString3 = ""
          DetailsString4 = ""
          st = ""
          i = 1
          With FlxDetail
             While i <= MaxRowFlexGrid - 1
                While Len(st) + 255 < 4000 And i <= MaxRowFlexGrid - 1
                    If Val(.TextMatrix(i, 1)) > 0 Then
                        st = GenerateDetailsString3(st, .TextMatrix(i, 1), .TextMatrix(i, 5), Val(.TextMatrix(i, 3)), Val(.TextMatrix(i, 11)), Val(.TextMatrix(i, 12)), .TextMatrix(i, 13), " ", .TextMatrix(i, 14), "", .TextMatrix(i, 8), .TextMatrix(i, 9))
                    End If
                    i = i + 1
                Wend
                If DetailsString1 = "" Then
                    DetailsString1 = st
                    st = ""
                ElseIf DetailsString2 = "" Then
                    DetailsString2 = st
                    st = ""
                ElseIf DetailsString3 = "" Then
                    DetailsString3 = st
                    st = ""
                ElseIf DetailsString4 = "" Then
                     DetailsString4 = st
                     st = ""
                End If
             Wend
          End With

        ReDim Parameter(28) As Parameter
        
        Parameter(0) = GenerateInputParameter("@Status", adInteger, 4, 2)
        Parameter(1) = GenerateInputParameter("@Owner", adInteger, 4, 0)
        If (Val(lblCustomer.Tag) > -1) Then
            Parameter(2) = GenerateInputParameter("@Customer", adInteger, 4, Val(lblCustomer.Tag))
        Else
            Parameter(2) = GenerateInputParameter("@Customer", adInteger, 4, -1)
        End If
        Parameter(3) = GenerateInputParameter("@DiscountTotal", adDouble, 8, Val(Me.lblDiscountTotal) - AutoDiscountValue)
        Parameter(4) = GenerateInputParameter("@CarryFeeTotal", adDouble, 8, Val(Me.lblCarryFeeTotal))
        Parameter(5) = GenerateInputParameter("@Recursive", adInteger, 4, Val(Me.txtRecursive.Text))
        Parameter(6) = GenerateInputParameter("@InCharge", adInteger, 4, 0)
        Parameter(7) = GenerateInputParameter("@FacPayment", adBoolean, 1, 1)
        Parameter(8) = GenerateInputParameter("@OrderType", adInteger, 4, mVarOrderType)
        Parameter(9) = GenerateInputParameter("@StationId", adInteger, 4, clsArya.StationNo)
        Parameter(10) = GenerateInputParameter("@ServiceTotal", adDouble, 8, ServiceRate)
        Parameter(11) = GenerateInputParameter("@PackingTotal", adDouble, 8, Val(Me.lblPackingTotal))
        Parameter(12) = GenerateInputParameter("@TableNo", adInteger, 4, cmbTable.ItemData(cmbTable.ListIndex))
        Parameter(13) = GenerateInputParameter("@User", adInteger, 4, mvarCurUserNo)
        Parameter(14) = GenerateInputParameter("@Date", adVarWChar, 50, txtDate.Text) 'mvarDate
        Parameter(15) = GenerateInputParameter("@DetailsString", adVarWChar, 4000, DetailsString1)
        Parameter(16) = GenerateInputParameter("@ds", adVarWChar, 4000, sFactorReceived)
        Parameter(17) = GenerateInputParameter("@Balance", adBoolean, 1, Abs(CInt(BalancePayment)))
        Parameter(18) = GenerateInputParameter("@AccountYear", adSmallInt, 2, AccountYear)
        Parameter(19) = GenerateInputParameter("@NvcDescription", adVarWChar, 150, " ”›«—‘ -" & Val(txtNo.Text))
        Parameter(20) = GenerateInputParameter("@HavaleNo", adInteger, 4, 0)
        Parameter(21) = GenerateInputParameter("@TempAddress", adVarWChar, 255, Trim(TxtTempAddress.Text))
        Parameter(22) = GenerateInputParameter("@GuestNo", adSmallInt, 2, Val(TxtGuestNo.Text))
        Parameter(23) = GenerateInputParameter("@DetailsString2", adVarWChar, 4000, DetailsString2)
        Parameter(24) = GenerateInputParameter("@DetailsString3", adVarWChar, 4000, DetailsString3)
        Parameter(25) = GenerateInputParameter("@DetailsString4", adVarWChar, 4000, DetailsString4)
        Parameter(26) = GenerateInputParameter("@AddedTotal", adInteger, 4, IIf(chKTax.Value = 1, 0, Val(lblTaxTotal)))
        Parameter(27) = GenerateInputParameter("@Rasmi", adBoolean, 1, chKTax)
        Parameter(28) = GenerateOutputParameter("@LastFacMNo", adInteger, 4)
        Dim NewFich As Long
        NewFich = RunParametricStoredProcedure("InsertFactorMasterDetails", Parameter)
        If NewFich <= 0 Then GoTo ErrHandler
                                
        ReDim Parameter(0) As Parameter
        Parameter(0) = GenerateInputParameter("@intSerialNo", adBigInt, 8, NewFich)
        Set rctmp = RunParametricStoredProcedure2Rec("Get_RowCount_FactorDetail", Parameter)
        NewFich = rctmp!No
        ReDim Parameter(3) As Parameter
        
        Parameter(0) = GenerateInputParameter("@No", adBigInt, 8, NewFich)
        Parameter(1) = GenerateInputParameter("@AccountYear", adSmallInt, 2, AccountYear)
        Parameter(2) = GenerateInputParameter("@OrderNo", adInteger, 4, Val(txtNo.Text))
        Parameter(3) = GenerateInputParameter("@Branch", adInteger, 4, CurrentBranch)
        Set Rst = RunParametricStoredProcedure2Rec("UpdateOrderRefrence", Parameter)
        
        If clsStation.PrintAfterDeliver Then
            mvarStatus = Invoice
            ClsPrint.Printing NewFich, clsArya.StationNo, InvoiceFactor, EnumActionLog.InvoicePrint
            mvarStatus = Order
        End If
        If clsStation.PrintAfterOrder Then
            mvarStatus = Invoice
            ClsPrint.Printing NewFich, clsArya.StationNo, EnumAddEditMode.AddMode, EnumActionLog.InvoicePrint
            mvarStatus = Order
        End If
        sFactorReceived = ""
        If clsStation.StopOnEditFich = False Or MyFormAddEditMode = AddMode Then
            If clsStation.InvoiceStatusDefault = True Then
                mvarStatus = EnumFactorType.Invoice
                If clsStation.Language = Farsi Then
                    LblInvoice.Caption = "›«ﬂ Ê— ›—Ê‘"
                Else
                    LblInvoice.Caption = "Invoice"
                End If
                If clsStation.PayFactorView = True Then
                    cmdPayFactor.Visible = True
                    lblPayFactorTotal.Visible = True
                Else
                    cmdPayFactor.Visible = False
                    lblPayFactorTotal.Visible = False
                End If
            End If
            Add
        Else
            MyFormAddEditMode = ViewMode
            GetDataDetail
            RefreshLables
            SetFirstToolBar
        End If
        Exit Sub
    End If

If clsStation.CashClose = True Then
    frmDisMsg.lblMessage.Caption = "’‰œÊﬁ »” Â «”  Ê «„ﬂ«‰ œ—Ì«›  ÊÃÊœ ‰œ«—œ"
    frmDisMsg.Timer1.Interval = 3000
    frmDisMsg.Timer1.Enabled = True
    frmDisMsg.Show vbModal
    Exit Sub
End If
If ClsFormAccess.frmFactorReceived = False Then
    frmDisMsg.lblMessage.Caption = "‘„« »Â «Ì‰ «„ò«‰ œ” —”Ì ‰œ«—Ìœ"
    frmDisMsg.Timer1.Interval = 3000
    frmDisMsg.Timer1.Enabled = True
    frmDisMsg.Show vbModal
    Exit Sub
End If
   
'Invoice
If MaxRowFlexGrid < 2 Then
'    If ClsFormAccess.frmGroupBoxTo = True Then
'        Unload frmFindGoods
'        Unload frmFindCust
'
'        frmGroupBoxTo.Show
'    End If
Else

    If MyFormAddEditMode <> AddMode Then
         ReDim Parameter(4) As Parameter
         Parameter(0) = GenerateInputParameter("@No", adBigInt, 8, Val(txtNo.Text))
         Parameter(1) = GenerateInputParameter("@Status", adInteger, 4, mvarStatus)
         Parameter(2) = GenerateInputParameter("@AccountYear", adSmallInt, 2, AccountYear)
         Parameter(3) = GenerateInputParameter("@Branch", adInteger, 4, CurrentBranch)
         Parameter(4) = GenerateOutputParameter("@Balance", adInteger, 4)

        Balance = RunParametricStoredProcedure("Get_tfacm_Balance", Parameter)
    Else
        Balance = 0
    End If
    If ChkIsLocked <> 0 Then
        frmDisMsg.lblMessage = " ›Ì‘ ﬁ›· «” . Ê «„ò«‰ œ—Ì«›  ¬‰ ÊÃÊœ ‰œ«—œ"
        frmDisMsg.Timer1.Enabled = True
        frmDisMsg.Show vbModal
        Exit Sub
    ElseIf Me.txtRecursive = 1 Then
        frmDisMsg.lblMessage = " ›Ì‘ ﬁ»·« „—ÃÊ⁄ ‘œÂ "
        frmDisMsg.Timer1.Enabled = True
        frmDisMsg.Show vbModal
        Exit Sub
    ElseIf Balance = 1 And MyFormAddEditMode = ViewMode Then
        frmDisMsg.lblMessage = " ›Ì‘ ﬁ»·«  ”ÊÌÂ ‘œÂ «”  "
        frmDisMsg.Timer1.Enabled = True
        frmDisMsg.Show vbModal
        If mvarServePlace = Table And cmbTable.ListIndex >= 0 Then
            ReDim Parameter(0) As Parameter
            Parameter(0) = GenerateInputParameter("@intTableNo", adInteger, 4, cmbTable.ItemData(cmbTable.ListIndex))
            RunParametricStoredProcedure "Update_tTable_By_Empty", Parameter
        End If
        Exit Sub
    End If
    If BlnPosResponsed = False Then
'        If strCategory = "00" And strDelegate = "00" And clsArya.CustomerId = 221 Then  'Naghsh_Jahan
        If clsArya.HardLockSerialNo = "86021800382" Then  'Naghsh_Jahan
           mvarInput = 0
        Else
'            If (MyFormAddEditMode = AddMode Or MyFormAddEditMode = EditMode Or MyFormAddEditMode = ManipulateMode) And clsStation.PosPayment = True Then ReceiveTypeFlag = True  ''And clsStation.PosModel > 0
'            If ReceiveTypeFlag = True Then
''                    ShowInputForm True, True, False, "œ—Ì«›  ‰ﬁœÌ", "œ—Ì«›  «“ ÿ—Ìﬁ ŒÊœÅ—œ«“ »«‰òÌ", "", "‰ÕÊÂ œ—Ì«›  —« „‘Œ’ ‰„«ÌÌœ", True, True, False, 1
''                Else
''                    mvarInput = "0"
''                End If
'                ShowInputForm True, True, False, "œ—Ì«›  ‰ﬁœÌ", "œ—Ì«›  œ” Ì «“ ÿ—Ìﬁ ŒÊœÅ—œ«“ »«‰òÌ", "œ—Ì«›  « Ê„« Ìò «“ ÿ—Ìﬁ ŒÊœÅ—œ«“ »«‰òÌ", "‰ÕÊÂ œ—Ì«›  —« „‘Œ’ ‰„«ÌÌœ", True, True, False, 1
'            Else
'                ShowInputForm True, True, False, "œ—Ì«›  ‰ﬁœÌ", "œ—Ì«›  ‰ﬁœÌ Ì««“ ÿ—Ìﬁ ŒÊœÅ—œ«“ »«‰òÌ", "", "‰ÕÊÂ œ—Ì«›  —« „‘Œ’ ‰„«ÌÌœ", True, True, False, 0
'            End If
        End If
'        If mvarInput = "" Then
'            Exit Sub
'        ElseIf mvarInput = "2" Then
'
''            Price = Val(Me.lblSumPrice.Tag)
''            SendPriceToPOS txtNo, Price
'            Exit Sub
'        ElseIf mvarInput = "1" Then
            MoveToCredit = False
            frmFactorReceived.intSerialNo = intSerialNo
            frmFactorReceived.FWBtnPrint.Visible = True
            frmFactorReceived.Show vbModal
'        ElseIf mvarInput = "0" Then
'            sFactorReceived = GenerateDetailsStringFactorReceived("", 1, 0, 0, 0, 0, "", 0, Val(lblSumPrice.Tag), "", "")   '- Val(lblPayFactorTotal.Caption)
'        End If
'        If mvarInput = "1" And mvarIndexNo = 0 Then Exit Sub
         If mvarIndexNo = 0 Then Exit Sub
        If MyFormAddEditMode = AddMode Or MyFormAddEditMode = EditMode Or MyFormAddEditMode = ManipulateMode Then
            UpdateFromFinalCheck = False
'            If clsStation.FinalCheck = True And mvarInput = "0" Then
'                FrmFinalCheck.FWBtnPrint.Visible = False
'                FrmFinalCheck.Show vbModal
'                If mvarIndexNo = 0 Then Exit Sub
'            End If
            '«‰ ﬁ«· »Â Õ”«» „‘ —Ì «⁄ »«—Ì
            If MoveToCredit = False Then
                boolPayment = True
                BalancePayment = True
            Else
                boolPayment = True
                BalancePayment = False
                sFactorReceived = ""
            End If
            PayClick = True
            If mvarIndexNo = 1 Then
                Update
            ElseIf mvarIndexNo = 2 Then
                Printing
            End If
            UpdateFromFinalCheck = True
        Else    'View
             temp = Val(txtNo.Text)
             If mvarStatus = Order Then
                BalancePayment = True
                boolPayment = True
                PayClick = True
                temp = Update
'                    temp = Val(txtNo.Text)
                ReDim Parameter(5) As Parameter
                Parameter(0) = GenerateInputParameter("@No", adBigInt, 8, Val(txtNo.Text))
                Parameter(1) = GenerateInputParameter("@Status", adInteger, 4, mvarStatus)
                Parameter(2) = GenerateInputParameter("@UID", adInteger, 4, mvarCurUserNo)
                Parameter(3) = GenerateInputParameter("@AccountYear", adSmallInt, 2, AccountYear)
                Parameter(4) = GenerateInputParameter("@ds", adWChar, 4000, sFactorReceived)
                Parameter(5) = GenerateInputParameter("@Branch", adInteger, 4, CurrentBranch)
                RunParametricStoredProcedure "Update_tfacm_Balance", Parameter
             Else   'Invoice
'                If clsStation.FinalCheck = True And mvarInput = "0" Then
'                    FrmFinalCheck.FWBtnPrint.Visible = False
'                    FrmFinalCheck.Show vbModal
'                    If mvarIndexNo = 1 Or mvarIndexNo = 2 Then
'                        ReDim Parameter(5) As Parameter
'                        Parameter(0) = GenerateInputParameter("@No", adBigInt, 8, Val(txtNo.Text))
'                        Parameter(1) = GenerateInputParameter("@Status", adInteger, 4, mvarStatus)
'                        Parameter(2) = GenerateInputParameter("@UID", adInteger, 4, mvarCurUserNo)
'                        Parameter(3) = GenerateInputParameter("@AccountYear", adSmallInt, 2, AccountYear)
'                        Parameter(4) = GenerateInputParameter("@ds", adWChar, 4000, sFactorReceived)
'                        Parameter(5) = GenerateInputParameter("@Branch", adInteger, 4, CurrentBranch)
'                        RunParametricStoredProcedure "Update_tfacm_Balance", Parameter
'                        If mvarTipAmount > 0 Then
'                             ReDim Parameter(4) As Parameter
'
'                             Parameter(0) = GenerateInputParameter("@No", adBigInt, 8, temp)
'                             Parameter(1) = GenerateInputParameter("@intServiceNo", adInteger, 4, mvarServiceStatus.Tip)
'                             Parameter(2) = GenerateInputParameter("@Amount", adBigInt, 8, mvarTipAmount)
'                             Parameter(3) = GenerateInputParameter("@AccountYear", adSmallInt, 2, AccountYear)
'                             Parameter(4) = GenerateInputParameter("@Branch", adInteger, 4, CurrentBranch)
'                             Set Rst = RunParametricStoredProcedure2Rec("InsertFactorAdditionalServices", Parameter)
'                             mvarTipAmount = 0
'                         End If
'                    Else
'                        Exit Sub
'                    End If
'                Else
                    If MoveToCredit = False Then
                    Else
                        sFactorReceived = ""
                    End If
                    ReDim Parameter(5) As Parameter
                    Parameter(0) = GenerateInputParameter("@No", adBigInt, 8, Val(txtNo.Text))
                    Parameter(1) = GenerateInputParameter("@Status", adInteger, 4, mvarStatus)
                    Parameter(2) = GenerateInputParameter("@UID", adInteger, 4, mvarCurUserNo)
                    Parameter(3) = GenerateInputParameter("@AccountYear", adSmallInt, 2, AccountYear)
                    Parameter(4) = GenerateInputParameter("@ds", adWChar, 4000, sFactorReceived)
                    Parameter(5) = GenerateInputParameter("@Branch", adInteger, 4, CurrentBranch)
                    RunParametricStoredProcedure "Update_tfacm_Balance", Parameter
'                End If
                If clsStation.StopOnEditFich = False Then
                    Add
                Else
                    GetDataDetail
                    RefreshLables
                    
                End If
            End If
        End If
        If clsStation.PrintAfterTasvieh Then
            If mvarStatus = Order Then
                mvarStatus = Invoice
                ClsPrint.Printing temp, clsArya.StationNo, InvoiceFactor, EnumActionLog.InvoicePrint
                mvarStatus = Order
            Else
                ClsPrint.Printing temp, clsArya.StationNo, InvoiceFactor, EnumActionLog.InvoicePrint
            End If
        End If
        If MoveToCredit = False Then
            frmDisMsg.lblMessage = " ›Ì‘   ”ÊÌÂ ‘œ "
        Else
            frmDisMsg.lblMessage = " ›Ì‘ »Â Õ”«» „‘ —Ì «⁄ »«—Ì „‰ ﬁ· ‘œ "
        End If
        frmDisMsg.Timer1.Enabled = True
        frmDisMsg.Show vbModal
'    Else
'        sFactorReceived = GenerateDetailsStringFactorReceived("", 5, 0, 0, 0, 0, "", 0, Val(lblSumPrice.Tag), "", "")          '- Val(lblPayFactorTotal.Caption)
'        ReDim Parameter(5) As Parameter
'        Parameter(0) = GenerateInputParameter("@No", adBigInt, 8, Val(txtNo.Text))
'        Parameter(1) = GenerateInputParameter("@Status", adInteger, 4, mvarStatus)
'        Parameter(2) = GenerateInputParameter("@UID", adInteger, 4, mvarCurUserNo)
'        Parameter(3) = GenerateInputParameter("@AccountYear", adSmallInt, 2, AccountYear)
'        Parameter(4) = GenerateInputParameter("@ds", adWChar, 4000, sFactorReceived)
'        Parameter(5) = GenerateInputParameter("@Branch", adInteger, 4, CurrentBranch)
'        RunParametricStoredProcedure "Update_tfacm_Balance", Parameter
'        If clsStation.PrintAfterTasvieh Then
'            ClsPrint.Printing temp, clsArya.StationNo, InvoiceFactor, EnumActionLog.InvoicePrint
'        End If
'        If clsStation.StopOnEditFich = False Then
'            Add
'        End If
    End If
End If

Exit Sub

ErrHandler:
    LogSaveNew "frmInvoice => ", err.Description, err.Number, err.Source, "cmdPay_Click"
    MsgBox err.Description, vbOKOnly, err.Number

End Sub

Private Sub CmdStationSaleSummery_Click()
    
    If clsArya.ExternalAccounting = True Then
        ShowMessage "¬Ì« „«Ì·Ìœ Œ·«’Â ê“«—‘ ŒÊœ —« À»  ò‰Ìœø", True, True, "»·Ì", "ŒÌ—"
        If mvarMsgIdx = vbYes Then
            Load frmReceivedSummary
            frmReceivedSummary.AccessUser = False
            frmReceivedSummary.frame_Change.Visible = False
    '        frmReceivedSummary.Height = frmReceivedSummary.Height - frmReceivedSummary.frame_Change.Height
            frmReceivedSummary.Show vbModal
        End If
    End If
    If ClsFormAccess.DailyReport = True Then
        Dim ArrayUbound  As Integer
        ReDim Parameter(12) As Parameter
        
       
        Parameter(0) = GenerateInputParameter("@SystemDate", adVarWChar, 50, Mid(clsDate.shamsi(Date), 3, 8))
        Parameter(1) = GenerateInputParameter("@SystemDay", adVarWChar, 50, clsDate.Find_DayOfWeek(Weekday(Date, vbSaturday)))
        Parameter(2) = GenerateInputParameter("@SystemTime", adVarWChar, 50, Mid(Str(time), 1, 5))
        Parameter(3) = GenerateInputParameter("@Date1", adVarWChar, 50, txtDate.Text)
        Parameter(4) = GenerateInputParameter("@Date2", adVarWChar, 50, Mid(clsDate.shamsi(Date), 3, 8))
        Parameter(5) = GenerateInputParameter("@Time1", adVarWChar, 50, "00:00")
        Parameter(6) = GenerateInputParameter("@Time2", adVarWChar, 50, "24:00") ' FormatDateTime(Time, vbShortTime)) 'Mid(str(Time), 1, 5))
        Parameter(7) = GenerateInputParameter("@User1", adInteger, 4, mvarCurUserNo)
        Parameter(8) = GenerateInputParameter("@User2", adInteger, 4, mvarCurUserNo)
        Parameter(9) = GenerateInputParameter("@Station1", adInteger, 4, 1)
        Parameter(10) = GenerateInputParameter("@Station2", adInteger, 4, 100)    '
        Parameter(11) = GenerateInputParameter("@Branch1", adInteger, 4, CurrentBranch)
        Parameter(12) = GenerateInputParameter("@Branch2", adInteger, 4, CurrentBranch)
        
        CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\RepStationSaleSummaryByUser.rpt"
        Dim fileSystem As New FileSystemObject
        Dim IsFileExist As Boolean
        IsFileExist = fileSystem.FileExists(CrystalReport1.ReportFileName)
        If IsFileExist = False Then
            frmDisMsg.lblMessage = " ›«Ì· Œ·«’Â ê“«—‘ ›—Ê‘ ÅÌœ« ‰‘œ "
            frmDisMsg.Timer1.Interval = 1000
            frmDisMsg.Timer1.Enabled = True
            frmDisMsg.Show vbModal
            Exit Sub
        End If
        CrystalReport1.ReportTitle = clsArya.StationName
        CrystalReport1.Destination = crptToWindow 'crptToPrinter '
        Dim intIndex As Integer
       
        For intIndex = 0 To UBound(Parameter) - LBound(Parameter)
            CrystalReport1.ParameterFields(intIndex) = CStr(Parameter(intIndex).Name) & ";" & CStr(Parameter(intIndex).Value) & ";" & "True"
        Next intIndex
    
        CrystalReport1.WindowShowGroupTree = True
        CrystalReport1.WindowShowSearchBtn = True
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
    Else
        frmDisMsg.lblMessage = "‘„« »Â «Ì‰ ﬁ«»·Ì  œ” —”Ì ‰œ«—Ìœ "
        frmDisMsg.Timer1.Interval = 1000
        frmDisMsg.Timer1.Enabled = True
        frmDisMsg.Show vbModal
    End If

End Sub

Private Sub cmdTables_Click()
    
    If intVersion <> gold And intVersion <> Diamond Then
        ShowDisMessage " ›ﬁÿ Ê—é‰ Â«Ì ÿ·«∆Ì „Ì  Ê«‰‰œ «“ «Ì‰ ﬁ«»·Ì  «” ›«œÂ ﬂ‰‰œ ", 1000
        Exit Sub
    End If
    If ClsFormAccess.GraphicTables = False Then
        ShowDisMessage " œ” —”Ì »—«Ì «Ì‰ ﬁ«»·Ì  ò«›Ì ‰Ì”  ", 1000
        Exit Sub
    End If
    If MyFormAddEditMode = AddMode Then mvarTable = 0
    mvarInvoiceNO = 0
    
    frmTables.Show vbModal
    If mvarInvoiceNO <> 0 Then
        Cancel
        txtNo.Text = mvarInvoiceNO
        mvarInvoiceNO = 0
        MyFormAddEditMode = ViewMode   'view Mode
        GetDataDetail
        RefreshLables
        SetFirstToolBar
        MyFormAddEditMode = ViewMode   'view Mode
    ElseIf mvarTable <> 0 And MyFormAddEditMode <> ViewMode Then
        'Cancel
        FillsFullTableCombo
        For i = 1 To cmbTable.ListCount - 1
            If cmbTable.ItemData(i) = mvarTable Then
                cmbTable.ListIndex = i
                Exit For
            End If
       Next
    ElseIf mvarTable <> 0 And MyFormAddEditMode = ViewMode And mvarStatus = Order Then
        'Cancel
        FillsFullTableCombo
        For i = 1 To cmbTable.ListCount - 1
            If cmbTable.ItemData(i) = mvarTable Then
                cmbTable.ListIndex = i
                Exit For
            End If
       Next
    Else
        If MyFormAddEditMode = AddMode Then
             cmbTable.ListIndex = -1
             cmbGarson.ListIndex = -1
        End If
    End If

End Sub

Private Sub cmdTempFich_Click()
    framelastFich.Visible = False
    If BlnPosApprovedWait = True Then
        ShowDisMessage "”Ì” „ „‰ Ÿ— œ—Ì«›  Å«”Œ «“ ŒÊœÅ—œ«“ „Ì »«‘œ", 1000
        Exit Sub
    End If

    If MyFormAddEditMode <> AddMode Then Exit Sub
    If cmbGarson.ListIndex = -1 Then
       cmbGarson.ListIndex = 0
    End If
    If MaxRowFlexGrid > 1 Then
        
     Dim st As String
     DetailsString1 = ""
     DetailsString2 = ""
     DetailsString3 = ""
     DetailsString4 = ""
     st = ""
     i = 1
     
     With FlxDetail
        While i <= MaxRowFlexGrid - 1
           While Len(st) + 255 < 4000 And i <= MaxRowFlexGrid - 1
               If Val(.TextMatrix(i, 1)) > 0 Then
                    st = GenerateDetailsString3(st, .TextMatrix(i, 1), .TextMatrix(i, 5), Val(.TextMatrix(i, 3)), Val(.TextMatrix(i, 11)), Val(.TextMatrix(i, 12)), .TextMatrix(i, 13), " ", .TextMatrix(i, 14), "", .TextMatrix(i, 8), .TextMatrix(i, 9))
               End If
               i = i + 1
           Wend
           If DetailsString1 = "" Then
               DetailsString1 = st
               st = ""
           ElseIf DetailsString2 = "" Then
               DetailsString2 = st
               st = ""
           ElseIf DetailsString3 = "" Then
               DetailsString3 = st
               st = ""
           
           ElseIf DetailsString4 = "" Then
                DetailsString4 = st
                st = ""
           End If
        Wend
     End With
     If Len(st) > 0 Then
         frmMsg.fwlblMsg.Caption = " ⁄œ«œ ”ÿ—Â« «“ „ﬁœ«— „Ã«“ »Ì‘ — „Ì »«‘œ. «„ò«‰ À»  ÊÃÊœ ‰œ«—œ"
         frmMsg.fwBtn(0).Visible = False
         frmMsg.fwBtn(1).ButtonType = flwButtonOk
         frmMsg.fwBtn(1).Caption = "ﬁ»Ê·"
         frmMsg.Show vbModal
         Exit Sub
     End If
        If intTempFich = 0 Then
            
            If cmbTable.ListIndex = -1 Then
               cmbTable.ListIndex = 0
            End If
            
            ReDim Parameter(24) As Parameter
            
            Parameter(0) = GenerateInputParameter("@Status", adInteger, 4, mvarStatus)
            Parameter(1) = GenerateInputParameter("@Owner", adInteger, 4, 0)
            If (Val(lblCustomer.Tag) > -1) Then
                Parameter(2) = GenerateInputParameter("@Customer", adInteger, 4, Me.lblCustomer.Tag)
            Else
                Parameter(2) = GenerateInputParameter("@Customer", adInteger, 4, -1)
            End If
            
            Parameter(3) = GenerateInputParameter("@DiscountTotal", adDouble, 8, Val(Me.lblDiscountTotal) - AutoDiscountValue)
            Parameter(4) = GenerateInputParameter("@CarryFeeTotal", adDouble, 8, Val(Me.lblCarryFeeTotal))
            Parameter(5) = GenerateInputParameter("@SumPrice", adDouble, 8, Val(Me.lblSumPrice.Tag))
            Parameter(6) = GenerateInputParameter("@Recursive", adInteger, 4, Val(Me.txtRecursive.Text))
            Parameter(7) = GenerateInputParameter("@InCharge", adInteger, 4, cmbGarson.ItemData(cmbGarson.ListIndex))
            Parameter(8) = GenerateInputParameter("@FacPayment", adBoolean, 1, Abs(CInt(boolPayment)))
            Parameter(9) = GenerateInputParameter("@OrderType", adInteger, 4, mVarOrderType)
            Parameter(10) = GenerateInputParameter("@StationId", adInteger, 4, clsArya.StationNo)
            Parameter(11) = GenerateInputParameter("@ServiceTotal", adDouble, 8, Val(Me.lblServiceTotal))
            Parameter(12) = GenerateInputParameter("@PackingTotal", adDouble, 8, Val(Me.lblPackingTotal))
            Parameter(13) = GenerateInputParameter("@TableNo", adInteger, 4, cmbTable.ItemData(cmbTable.ListIndex))
            Parameter(14) = GenerateInputParameter("@User", adInteger, 4, mvarCurUserNo)
            Parameter(15) = GenerateInputParameter("@Date", adVarWChar, 50, txtDate.Text)
            
            Parameter(16) = GenerateInputParameter("@DetailsString", adVarWChar, 4000, DetailsString1)
            Parameter(17) = GenerateInputParameter("@NvcDescription", adVarWChar, 150, IIf(textDescriptionFlag = True, Right(txtDescription.Text, 150), " "))
            Parameter(18) = GenerateInputParameter("@GuestNo", adSmallInt, 2, Val(TxtGuestNo.Text))
            Parameter(19) = GenerateInputParameter("@Branch", adInteger, 4, CurrentBranch)
            Parameter(20) = GenerateInputParameter("@TempAddress", adVarWChar, 255, IIf(TempAddressEdit = True, Trim(Right(TxtTempAddress.Text, 255)), " "))
            Parameter(21) = GenerateInputParameter("@DetailsString2", adVarWChar, 4000, DetailsString2)
            Parameter(22) = GenerateInputParameter("@DetailsString3", adVarWChar, 4000, DetailsString3)
            Parameter(23) = GenerateInputParameter("@DetailsString4", adVarWChar, 4000, DetailsString4)
            Parameter(24) = GenerateOutputParameter("@LastFacMNo", adInteger, 4)
         
                    
            RunParametricStoredProcedure "InsertFactorMasterDetailsTemp", Parameter
            
        Else
            ReDim Parameter(24) As Parameter
            Parameter(0) = GenerateInputParameter("@No", adBigInt, 8, intTempFich)
            Parameter(1) = GenerateInputParameter("@Status", adInteger, 4, mvarStatus)
            Parameter(2) = GenerateInputParameter("@Owner", adInteger, 4, 0)
            If Val(lblCustomer.Tag) > -1 Then
                Parameter(3) = GenerateInputParameter("@Customer", adInteger, 4, Me.lblCustomer.Tag)
            Else
                Parameter(3) = GenerateInputParameter("@Customer", adInteger, 4, -1)
            End If
            Parameter(4) = GenerateInputParameter("@DiscountTotal", adDouble, 8, Val(Me.lblDiscountTotal) - AutoDiscountValue)
            Parameter(5) = GenerateInputParameter("@CarryFeeTotal", adDouble, 8, Val(Me.lblCarryFeeTotal))
            Parameter(6) = GenerateInputParameter("@SumPrice", adDouble, 8, Val(Me.lblSumPrice.Tag))
            Parameter(7) = GenerateInputParameter("@Recursive", adInteger, 4, Val(Me.txtRecursive.Text))
            Parameter(8) = GenerateInputParameter("@InCharge", adInteger, 4, cmbGarson.ItemData(cmbGarson.ListIndex))
            Parameter(9) = GenerateInputParameter("@FacPayment", adBoolean, 1, Abs(CInt(boolPayment)))
            Parameter(10) = GenerateInputParameter("@OrderType", adInteger, 4, mVarOrderType)
            Parameter(11) = GenerateInputParameter("@StationId", adInteger, 4, clsArya.StationNo)
            Parameter(12) = GenerateInputParameter("@ServiceTotal", adDouble, 8, Val(Me.lblServiceTotal))
            Parameter(13) = GenerateInputParameter("@PackingTotal", adDouble, 8, Val(Me.lblPackingTotal))
            Parameter(14) = GenerateInputParameter("@TableNo", adInteger, 4, cmbTable.ItemData(cmbTable.ListIndex))
            Parameter(15) = GenerateInputParameter("@Date", adVarWChar, 50, txtDate.Text)

            Parameter(16) = GenerateInputParameter("@DetailsString", adVarWChar, 4000, DetailsString1)
            Parameter(17) = GenerateInputParameter("@NvcDescription", adVarWChar, 150, IIf(textDescriptionFlag = True, Right(txtDescription.Text, 150), " "))
            Parameter(18) = GenerateInputParameter("@GuestNo", adSmallInt, 2, Val(TxtGuestNo.Text))
            Parameter(19) = GenerateInputParameter("@Branch", adInteger, 4, CurrentBranch)
            Parameter(20) = GenerateInputParameter("@TempAddress", adVarWChar, 255, IIf(TempAddressEdit = True, Trim(Right(TxtTempAddress.Text, 255)), " "))
            Parameter(21) = GenerateInputParameter("@DetailsString2", adVarWChar, 4000, DetailsString2)
            Parameter(22) = GenerateInputParameter("@DetailsString3", adVarWChar, 4000, DetailsString3)
            Parameter(23) = GenerateInputParameter("@DetailsString4", adVarWChar, 4000, DetailsString4)
            Parameter(24) = GenerateOutputParameter("@LastFacMNo", adInteger, 4)
                    
            RunParametricStoredProcedure "EditFactorMasterDetailsTemp", Parameter
        End If
        Add
    Else
        frmTempFich.Show vbModal
        If Val(frmTempFich.mvarcode) <> 0 Then
        
            LoadTempData (frmTempFich.mvarcode)
            RefreshLables
            SetFirstToolBar
            intTempFich = Val(frmTempFich.mvarcode)
        End If
    End If

    CalculateTemporary
    
End Sub

Private Sub FlxDetail_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim mvarSellPriceNew As Long
    Dim ReturnValue As Boolean
    ReturnValue = True
    With FlxDetail
        If Val(.TextMatrix(.Row, IndexColAmount)) <> 0 And Val(.TextMatrix(.Row, IndexColGoodCode)) <> 0 Then
'''
            If clsStation.RowMojodiControl = True And MojodiControlFlag = True And mvarStatus = Invoice Then
                DetailsString1 = ""
                With FlxDetail
                    DetailsString1 = GenerateDetailsString3(DetailsString1, IIf(Val(lblNum.Caption) = 0, 1, Val(lblNum.Caption)), .TextMatrix(.Row, IndexColGoodCode), CStr(mvarSellPrice), CStr(mvarDisCount), CStr(mvarRate), "", " ", .TextMatrix(.Row, IndexColInventory), "", .TextMatrix(.Row, IndexColServePalce), "")
                End With
            
                If MyFormAddEditMode = AddMode Then
                    ReDim Parameter(3) As Parameter
                    Parameter(0) = GenerateInputParameter("@DetailsString", adVarWChar, 4000, DetailsString1)
                    Parameter(1) = GenerateInputParameter("@DetailsString2", adVarWChar, 4000, "")
                    Parameter(2) = GenerateInputParameter("@DetailsString3", adVarWChar, 4000, "")
                    Parameter(3) = GenerateInputParameter("@DetailsString4", adVarWChar, 4000, "")
                    Set Rst = RunParametricStoredProcedure2Rec("CheckPreSave", Parameter)
                    If Not (Rst.BOF Or Rst.EOF) Then
                        mvarAddeditMode = MyFormAddEditMode
                        frmMojodiReduce.Show vbModal
                        If frmMojodiReduce.Result = False Then
                           ReturnValue = False
                        End If
                    End If
                Else
                    ReDim Parameter(5) As Parameter
                    mvarNo = Val(txtNo.Text)
                    Parameter(0) = GenerateInputParameter("@DetailsString", adVarWChar, 4000, DetailsString1)
                    Parameter(1) = GenerateInputParameter("@Status", adInteger, 4, mvarStatus)
                    Parameter(2) = GenerateInputParameter("@No", adBigInt, 8, mvarNo)
                    Parameter(3) = GenerateInputParameter("@DetailsString2", adVarWChar, 4000, "")
                    Parameter(4) = GenerateInputParameter("@DetailsString3", adVarWChar, 4000, "")
                    Parameter(5) = GenerateInputParameter("@DetailsString4", adVarWChar, 4000, "")
                    Set Rst = RunParametricStoredProcedure2Rec("CheckPreSave_Edit", Parameter)
                    Dim ss As String
                    If Not (Rst.BOF Or Rst.EOF) Then
                        mvarAddeditMode = MyFormAddEditMode
                        frmMojodiReduce.Show vbModal
                        If frmMojodiReduce.Result = False Then
                           ReturnValue = False
                        End If
                    End If
    
                End If
                lblNum.Caption = ""
'''
            End If
            If ReturnValue = True Then
                FlxDetail.TextMatrix(FlxDetail.Row, IndexColTotalFee) = Val(FlxDetail.TextMatrix(FlxDetail.Row, IndexColAmount)) * Val(FlxDetail.TextMatrix(FlxDetail.Row, IndexColFee))
            Else
                FlxDetail.TextMatrix(FlxDetail.Row, IndexColAmount) = OldAmount
            End If
            If Col = IndexColRate Or Col = IndexColServePalce Then
                ReDim Parameter(4) As Parameter
                Parameter(0) = GenerateInputParameter("@Code", adInteger, 4, FlxDetail.TextMatrix(Row, IndexColGoodCode))
                Parameter(1) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
                Parameter(2) = GenerateInputParameter("@StationId", adInteger, 4, clsArya.StationNo)
                Parameter(3) = GenerateInputParameter("@Flag", adBoolean, 1, 1)
                Parameter(4) = GenerateInputParameter("@AccountYear", adSmallInt, 2, AccountYear)
                Set rctmp = RunParametricStoredProcedure2Rec("Get_Good_Code", Parameter)
                
                If Not (rctmp.BOF Or rctmp.EOF) Then
                    If Col = IndexColServePalce Then
                        Dim TempRate As Integer
                        If FlxDetail.ValueMatrix(Row, IndexColServePalce) = Out Then
                            TempRate = clsStation.OutPrice
                        Else
                            TempRate = MainPriceType
                        End If
                        FlxDetail.TextMatrix(Row, IndexColRate) = TempRate
                        If TempRate = 1 Then
                           FlxDetail.TextMatrix(Row, IndexColFee) = rctmp.Fields("SellPrice").Value
                        ElseIf TempRate = 2 Then
                           FlxDetail.TextMatrix(Row, IndexColFee) = rctmp.Fields("SellPrice2").Value
                        ElseIf TempRate = 3 Then
                           FlxDetail.TextMatrix(Row, IndexColFee) = rctmp.Fields("SellPrice3").Value
                        ElseIf TempRate = 4 Then
                           FlxDetail.TextMatrix(Row, IndexColFee) = rctmp.Fields("SellPrice4").Value
                        ElseIf TempRate = 5 Then
                           FlxDetail.TextMatrix(Row, IndexColFee) = rctmp.Fields("SellPrice5").Value
                        ElseIf TempRate = 6 Then
                           FlxDetail.TextMatrix(Row, IndexColFee) = rctmp.Fields("SellPrice6").Value
                        End If
                    Else
                        If FlxDetail.TextMatrix(Row, IndexColRate) = 1 Then
                           FlxDetail.TextMatrix(Row, IndexColFee) = rctmp.Fields("SellPrice").Value
                        ElseIf FlxDetail.TextMatrix(Row, IndexColRate) = 2 Then
                           FlxDetail.TextMatrix(Row, IndexColFee) = rctmp.Fields("SellPrice2").Value
                        ElseIf FlxDetail.TextMatrix(Row, IndexColRate) = 3 Then
                           FlxDetail.TextMatrix(Row, IndexColFee) = rctmp.Fields("SellPrice3").Value
                        ElseIf FlxDetail.TextMatrix(Row, IndexColRate) = 4 Then
                           FlxDetail.TextMatrix(Row, IndexColFee) = rctmp.Fields("SellPrice4").Value
                        ElseIf FlxDetail.TextMatrix(Row, IndexColRate) = 5 Then
                           FlxDetail.TextMatrix(Row, IndexColFee) = rctmp.Fields("SellPrice5").Value
                        ElseIf FlxDetail.TextMatrix(Row, IndexColRate) = 6 Then
                           FlxDetail.TextMatrix(Row, IndexColFee) = rctmp.Fields("SellPrice6").Value
                        End If
                    End If
                End If    ''
            End If
            
            FlxDetail.TextMatrix(FlxDetail.Row, IndexColTotalFee) = Val(FlxDetail.TextMatrix(FlxDetail.Row, IndexColAmount)) * Val(FlxDetail.TextMatrix(FlxDetail.Row, IndexColFee))
            
            If Col = IndexColFee And clsStation.UpdateSellprice = True Then
                ReDim Parameter(2) As Parameter
                Parameter(0) = GenerateInputParameter("@Goodcode", adInteger, 4, FlxDetail.TextMatrix(.Row, 5))
                Parameter(1) = GenerateInputParameter("@SellpriceNO", adInteger, 4, Val(FlxDetail.TextMatrix(FlxDetail.Row, 12)))
                Parameter(2) = GenerateInputParameter("@NewSellPrice", adInteger, 4, Val(FlxDetail.TextMatrix(FlxDetail.Row, 3)))
                RunParametricStoredProcedure "UpdateSellPrice", Parameter
            End If
            If Col = IndexColDifferences And Len(FlxDetail.TextMatrix(FlxDetail.Row, IndexColDifferences)) > 0 And FlxDetail.TextMatrix(FlxDetail.Row, 9) = "" Then
               If SaveDifferences > 0 Then FlxDetail.TextMatrix(FlxDetail.Row, IndexColDifferences) = "": FlxDetail.Select FlxDetail.Row, IndexColDifferences: FlxDetail_Click   '''ShowDisMessage "À»   €ÌÌ—«  - " & FlxDetail.TextMatrix(FlxDetail.Row, IndexColGoodName) & "  -Êò«·«Â«Ì „‘«»Â ¬‰ «‰Ã«„ ‘œ", 1000:
            End If
        ElseIf Val(.TextMatrix(.Row, IndexColGoodCode)) <> 0 Then
''            If CheekGoodAmount() = False Then
''                frmAccess.lblTitle = "ﬂ«·« ‰„Ì  Ê«‰Ìœ ﬂ„ ﬂ‰Ìœ —„“ »« œ” —”Ì »«·« — Ê«—œ ò‰Ìœ"
''                frmAccess.AccessStatus = UpperAmountGood
''                frmAccess.Show vbModal
''                If frmAccess.ReturnAccess = False Then
''                    frmMsg.fwlblMsg.Caption = "ﬂ«·« ‰„Ì  Ê«‰Ìœ ﬂ„ ﬂ‰Ìœ"
''                    frmMsg.fwBtn(0).Visible = False
''                    frmMsg.fwBtn(1).ButtonType = flwButtonOk
''                    frmMsg.fwBtn(1).Caption = "ﬁ»Ê·"
''                    frmMsg.Show vbModal
''                    Exit Sub
''                End If
''            End If
            
            FlxDetail.RemoveItem (FlxDetail.Row)
            If FlxDetail.Rows < MaxInvoiceRows Then
                AddEmptyRow     'add row Instead of Remove
            End If
            MaxRowFlexGrid = MaxRowFlexGrid - 1
            RefreshFlxDetailRowNumber
            frmMsg.fwlblMsg.Caption = " .ò«·«Ì „Ê—œ ‰Ÿ— «“ ·Ì”  Õ–› ‘œ "
            frmMsg.fwBtn(0).Visible = False
            frmMsg.fwBtn(1).ButtonType = flwButtonOk
            frmMsg.fwBtn(1).Caption = "ﬁ»Ê·"
            frmMsg.Show vbModal
            FlxDetail.Row = MaxRowFlexGrid     'Last Row
''''            FlxDetail.TopRow = FlxDetail.Rows - 7
            txtScale.Text = ""
        End If
        BtnKeypad(11).Enabled = True     '"%"
        BtnKeypad(10).Enabled = True      '"."
        BtnKalaDelete.Enabled = True
        lblNum.Caption = ""
        If clsInvoiceValue.ShowInvoiceMenu = True Then
            frmShowInvoiceMenu.UpdateGridValue
        End If
        RefreshLables
    End With
End Sub


Private Sub FlxDetail_AfterSort(ByVal Col As Long, Order As Integer)
    
    With FlxDetail
        If .Rows < MaxInvoiceRows Then
            MaxRowFlexGrid = .Rows
            .Rows = MaxInvoiceRows
            .Row = MaxRowFlexGrid
        Else
            MaxRowFlexGrid = .Rows
            .Rows = .Rows + 1
            .Row = MaxRowFlexGrid
        End If
        
    End With
    
End Sub

Private Sub FlxDetail_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    For i = 0 To FlxDetail.Cols - 1
        SaveSetting strMainKey, Me.Name & "_ResFlexgrid", "Col" & i, FlxDetail.ColWidth(i)
    Next
End Sub

Private Sub FlxDetail_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With FlxDetail
        If Col = IndexColAmount Then
            OldAmount = .TextMatrix(Row, IndexColAmount)
        End If
    End With
End Sub

Private Sub FlxDetail_BeforeSort(ByVal Col As Long, Order As Integer)
    
    With FlxDetail
        i = .Rows - 1
        While i >= 1
            If .TextMatrix(i, 2) = "" Then
                .RemoveItem (i)
'                i = i - 1
            End If
            i = i - 1
        Wend
    End With
    
End Sub

Private Sub FlxDetail_Click()
Dim ReturnValue As Boolean
ReturnValue = True
    If MyFormAddEditMode = ViewMode Then Exit Sub
    
    
    With FlxDetail
        
        If .Row > 0 And .TextMatrix(.Row, 5) <> "" And (.Col = 0 Or .Col = 2) And EnableBeforShowDifferenceFlxRow = False Then    'And Val(.TextMatrix(.Row, 8)) = mvarServePlace
            mvarGoodCode = .TextMatrix(.Row, 5)
            mvarUnitGood = .TextMatrix(.Row, 7)
            mvarServePlace = .TextMatrix(.Row, 8)
            mvarGoodName = .TextMatrix(.Row, 2)
            If clsStation.HasOptionPrice = False Then
                mvarSellPrice = .TextMatrix(.Row, 3)
            Else
                ReDim Parameter(4) As Parameter
                Parameter(0) = GenerateInputParameter("@Code", adInteger, 4, mvarGoodCode)
                Parameter(1) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
                Parameter(2) = GenerateInputParameter("@StationId", adInteger, 4, clsArya.StationNo)
                Parameter(3) = GenerateInputParameter("@Flag", adBoolean, 1, 1)
                Parameter(4) = GenerateInputParameter("@AccountYear", adSmallInt, 2, AccountYear)
                Set rctmp = RunParametricStoredProcedure2Rec("Get_Good_Code", Parameter)
                If Not (rctmp.BOF Or rctmp.EOF) Then
                    If clsStation.PriceType = 1 Then
                       mvarSellPrice = rctmp.Fields("SellPrice").Value
                    ElseIf clsStation.PriceType = 2 Then
                       mvarSellPrice = rctmp.Fields("SellPrice2").Value
                    ElseIf clsStation.PriceType = 3 Then
                       mvarSellPrice = rctmp.Fields("SellPrice3").Value
                    ElseIf clsStation.PriceType = 4 Then
                       mvarSellPrice = rctmp.Fields("SellPrice4").Value
                    ElseIf clsStation.PriceType = 5 Then
                       mvarSellPrice = rctmp.Fields("SellPrice5").Value
                    ElseIf clsStation.PriceType = 6 Then
                       mvarSellPrice = rctmp.Fields("SellPrice6").Value
                    End If
                End If
            End If
            If clsStation.RowMojodiControl = True And MojodiControlFlag = True And mvarStatus = Invoice And .Col <> 0 Then
                DetailsString1 = ""
                With FlxDetail
                    DetailsString1 = GenerateDetailsString3(DetailsString1, IIf(Val(lblNum.Caption) = 0, Val(.TextMatrix(.Row, IndexColAmount)) + 1, Val(.TextMatrix(.Row, IndexColAmount)) + Val(lblNum.Caption)), .TextMatrix(.Row, IndexColGoodCode), CStr(mvarSellPrice), CStr(mvarDisCount), CStr(mvarRate), "", " ", .TextMatrix(.Row, IndexColInventory), "", .TextMatrix(.Row, IndexColServePalce), "")
                End With
            
                If MyFormAddEditMode = AddMode Then
                    ReDim Parameter(3) As Parameter
                    Parameter(0) = GenerateInputParameter("@DetailsString", adVarWChar, 4000, DetailsString1)
                    Parameter(1) = GenerateInputParameter("@DetailsString2", adVarWChar, 4000, "")
                    Parameter(2) = GenerateInputParameter("@DetailsString3", adVarWChar, 4000, "")
                    Parameter(3) = GenerateInputParameter("@DetailsString4", adVarWChar, 4000, "")
                    Set Rst = RunParametricStoredProcedure2Rec("CheckPreSave", Parameter)
                    If Not (Rst.BOF Or Rst.EOF) Then
                        mvarAddeditMode = MyFormAddEditMode
                        frmMojodiReduce.Show vbModal
                        If frmMojodiReduce.Result = False Then
                           ReturnValue = False
                        End If
                    End If
                Else
                    ReDim Parameter(5) As Parameter
                    mvarNo = Val(txtNo.Text)
                    Parameter(0) = GenerateInputParameter("@DetailsString", adVarWChar, 4000, DetailsString1)
                    Parameter(1) = GenerateInputParameter("@Status", adInteger, 4, mvarStatus)
                    Parameter(2) = GenerateInputParameter("@No", adBigInt, 8, mvarNo)
                    Parameter(3) = GenerateInputParameter("@DetailsString2", adVarWChar, 4000, "")
                    Parameter(4) = GenerateInputParameter("@DetailsString3", adVarWChar, 4000, "")
                    Parameter(5) = GenerateInputParameter("@DetailsString4", adVarWChar, 4000, "")
                    Set Rst = RunParametricStoredProcedure2Rec("CheckPreSave_Edit", Parameter)
                    Dim ss As String
                    If Not (Rst.BOF Or Rst.EOF) Then
                        mvarAddeditMode = MyFormAddEditMode
                        frmMojodiReduce.Show vbModal
                        If frmMojodiReduce.Result = False Then
                           ReturnValue = False
                        End If
                    End If
    
                End If
            End If
            If .Col = 0 Then lblNum.Caption = -1
            If ReturnValue = True Then ChangeGoodquantity
            
            
        ElseIf .Row > 0 And .TextMatrix(.Row, 5) <> "" And (.Col = 3) And EnableBeforShowDifferenceFlxRow = False Then
            If ClsFormAccess.CustomizeFee = True Then
              .Select .Row, .Col
              .EditCell
            ElseIf MyFormAddEditMode <> EnumAddEditMode.ViewMode Then
              frmDisMsg.lblMessage = "  ‘„« »Â «„ﬂ«‰   €ÌÌ— ›Ì œ” —”Ì ‰œ«—Ìœ "
              frmDisMsg.Timer1.Enabled = True
              frmDisMsg.Show vbModal
            End If
         ElseIf .Row > 0 And .TextMatrix(.Row, 5) <> "" And .Col = 11 Then
            If ClsFormAccess.Discount = True Then
              .Select .Row, .Col
              .EditCell
            ElseIf MyFormAddEditMode <> EnumAddEditMode.ViewMode Then
              frmDisMsg.lblMessage = "  ‘„« »Â «„ﬂ«‰ œ«œ‰  Œ›Ì› œ” —”Ì ‰œ«—Ìœ "
              frmDisMsg.Timer1.Enabled = True
              frmDisMsg.Show vbModal
            End If
       ElseIf .Row > 0 And .TextMatrix(.Row, 5) <> "" And (.Col = 1 Or .Col = 8 Or .Col = 13) Then
            .Select .Row, .Col
            .EditCell
        ElseIf .Row > 0 And .TextMatrix(.Row, 5) <> "" And (.Col = 10) Then
            
           If lstDifference.Visible = True Then  '' Save Diffence without press Enter and with touch in same column
                Dim SelectedDifferenceFlag As Boolean
                SelectedDifferenceFlag = False
                For i = 0 To lstDifference.ListCount - 1
                    If lstDifference.Selected(i) = True Then
                        SelectedDifferenceFlag = True
                        Exit For
                    End If
                Next i
                If SelectedDifferenceFlag = True Then
                    lstDifference_KeyUp 13, 0
                    FlxDetail.ShowCell MaxRowFlexGrid, 1
                    FlxDetail.SetFocus
                    FlxDetail.Select MaxRowFlexGrid, 1
                    If clsInvoiceValue.ShowInvoiceMenu = True Then
                        frmShowInvoiceMenu.UpdateGridValue
                    End If
                    Exit Sub
                End If
           End If
            
            .ShowCell .Row, .Col
            
            If (lstDifference.top = .RowPos(.Row) + .RowHeight(.Row)) And lstDifference.left = .CellLeft And lstDifference.Visible = True Then
                lstDifference.Clear
                lstDifference.Visible = False
                
                Exit Sub
            End If
            lstDifference.top = (.RowPos(.Row) + .RowHeight(.Row)) + .RowHeight(.Row) + 200
            lstDifference.left = .CellLeft + .CellWidth
            
            ReDim Parameter(1) As Parameter
            
            Parameter(0) = GenerateInputParameter("@GoodCode", adInteger, 4, .TextMatrix(.Row, 5))
            Parameter(1) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
            
            Set Rst = RunParametricStoredProcedure2Rec("Get_Good_Difference", Parameter)
            
            Dim ArrDifferences() As String
            
                
            ArrDifferences = Split(FlxDetail.TextMatrix(FlxDetail.Row, 9), ";")

            lstDifference.Clear
         
            OldCostDifference = 0
            ReDim ArrCostDifferences(0)
            While Rst.EOF <> True
                lstDifference.AddItem Rst!Difference
                lstDifference.ItemData(lstDifference.ListCount - 1) = Rst!Code
                On Error GoTo ErrHandler
                ReDim Preserve ArrCostDifferences(UBound(ArrCostDifferences) + 1)
                On Error GoTo 0
                ArrCostDifferences(UBound(ArrCostDifferences)) = Rst!CostDifference
                
                For i = LBound(ArrDifferences) To UBound(ArrDifferences)
                    If ArrDifferences(i) = lstDifference.ItemData(lstDifference.ListCount - 1) Then
                        lstDifference.Selected(lstDifference.ListCount - 1) = True
                          OldCostDifference = OldCostDifference + ArrCostDifferences(UBound(ArrCostDifferences))
                        
                    End If
                Next i
                Rst.MoveNext
            Wend
            EnableBeforShowDifferenceFlxRow = False
            Set Rst = Nothing
            If lstDifference.ListCount <> 0 Then
            
                If lstDifference.ListCount = 1 Then
                    lstDifference.Height = 1000
                ElseIf lstDifference.ListCount >= 10 Then
                    lstDifference.Height = 500 + 400 * Invoice_FontDifferencesSize
                Else
                    lstDifference.Height = 500 + (lstDifference.ListCount - 1) * 40 * Invoice_FontDifferencesSize
                End If
                lstDifference.Visible = True
                lstDifference.SetFocus
                BeforShowDifferenceFlxRow = FlxDetail.Row
                EnableBeforShowDifferenceFlxRow = True
                .Select .Row, 10: .EditCell
                Exit Sub
            Else
                 frmDisMsg.lblMessage = " €ÌÌ—« Ì »—«Ì «Ì‰ ò«·« „‰ŸÊ— ‰‘œÂ "
                 frmDisMsg.Timer1.Enabled = True
                 frmDisMsg.Show vbModal
                .Select .Row, 10: .EditCell
                
                
            End If
            
        End If
        .ShowCell .Row, 1
        
       If .Row > 0 And .TextMatrix(.Row, 5) <> "" And EnableBeforShowDifferenceFlxRow = True Then 'And Val(.TextMatrix(.Row, 8)) = mvarServePlace
            EnableBeforShowDifferenceFlxRow = False
            lstDifference_KeyUp vbKeyReturn, 0
'            .TextMatrix(.Row, 3) = .TextMatrix(.Row, 3) + GetCostDifferences
'            RefreshLables  'Set Lables
'            HideLstBoxes 27
        End If
    End With

    If clsInvoiceValue.ShowInvoiceMenu = True Then
        frmShowInvoiceMenu.UpdateGridValue
    End If

Exit Sub

ErrHandler:
    If err.Number = 9 Then
        ReDim ArrDifferences(0)
        Resume Next
    End If

End Sub


Private Sub FlexGridActive()

    With FlxDetail
        .Rows = MaxInvoiceRows
        .Cols = 19
        .ForeColor = &H40&
        
        
'        .ColAlignment(2) = flexAlignRightCenter
        .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter
        .ColAlignment(-1) = flexAlignCenterCenter
        .ColAlignment(2) = flexAlignRightCenter
        
         For i = 0 To .Cols - 1
            .ColWidth(i) = Val(GetSetting(strMainKey, Me.Name & "_ResFlexgrid", "Col" & i))
            If .ColWidth(i) = 0 Then
                .ColWidth(i) = 1000
            End If
         Next i
         
               
        SetHiddenCols
''''        .AutoSizeMode = flexAutoSizeColWidth
''''        .AutoSize 8, .Cols - 1
        .ColDataType(17) = flexDTBoolean
        .ColDataType(18) = flexDTBoolean
   
        .RowHeightMax = .Height / (MaxInvoiceRows * 1.08) '8.2
        .RowHeightMin = .Height / (MaxInvoiceRows * 1.11)  '8.5
        .ScrollBars = flexScrollBarBoth
        .Row = 1
        MaxRowFlexGrid = 1
    
    End With


End Sub
Private Sub SetHiddenCols()
    With FlxDetail
        .ColHidden(0) = Not clsInvoiceValue.ColRow
        .ColHidden(3) = Not clsInvoiceValue.ColFee
        .ColHidden(4) = Not clsInvoiceValue.ColTotal
        .ColHidden(5) = Not clsInvoiceValue.ColGoodCode
        .ColHidden(10) = Not clsInvoiceValue.ColChanges
        .ColHidden(11) = Not clsInvoiceValue.ColDiscount
        .ColHidden(12) = Not clsInvoiceValue.ColRate
        .ColHidden(14) = Not clsInvoiceValue.ColStore
        .ColHidden(16) = Not clsInvoiceValue.ColMojodi
        .ColHidden(6) = True   ' weight
        .ColHidden(7) = True   ' unit
        .ColHidden(8) = True   ' Serve
        .ColHidden(9) = True   ' changesCode
        .ColHidden(13) = True   ' Chair
        .ColHidden(15) = True   'Main Group
        .ColHidden(17) = Not clsInvoiceValue.ColDuty
        .ColHidden(18) = Not clsInvoiceValue.ColTax
'        If strCategory = "00" Or strCategory = "02" Or strCategory = "05" Or strCategory = "06" Or strCategory = "07" Then
'            .ColHidden(13) = False   '  Chair
'        Else
'            .ColHidden(13) = True   '  Chair
'        End If
    End With
End Sub
Private Sub BtnKalaDelete_Click()
'Case "-":
    If lblNum.Caption = "" Then
        lblNum.Caption = lblNum.Caption + BtnKalaDelete.Tag
        'BtnKalaDelete.ForeColor = &H80&
        BtnKeypad(11).Enabled = False     '"%"
        BtnKeypad(10).Enabled = True      '"."
    Else
        If left(lblNum.Caption, 1) = "-" Then
        lblNum.Caption = ""
        'BtnKalaDelete.ForeColor = &H404080
        BtnKeypad(11).Enabled = True     '"%"
        BtnKeypad(10).Enabled = True      '"."
        End If
    End If

End Sub

Private Sub BtnMenu_MouseMove(index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    BtnMenu(index).ToolTipText = BtnMenu(index).Caption

End Sub



Private Sub lblCustomer_Click()
    If BlnPosApprovedWait = True Then
        ShowDisMessage "”Ì” „ „‰ Ÿ— œ—Ì«›  Å«”Œ «“ ŒÊœÅ—œ«“ „Ì »«‘œ", 1000
        Exit Sub
    End If
If lblCustomer.Tag <> "-1" Then
    FrameCustInfo.Visible = True

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
    lblCountMonthBuy.Caption = ""
    FrameCustInfo = ""
    ReDim Parameter(0) As Parameter
    Parameter(0) = GenerateInputParameter("@Code", adBigInt, 8, Val(lblCustomer.Tag))
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
        lblCountCurrentBuy.Caption = Rst!CountCurrentDayBuy
        lblCountMonthBuy.Caption = BuyCountTimes1
    '    LblDescription = Rst!Description
        If Val(LastCredit1.Caption) > 0 Then
            LastCredit.Caption = "»œÂÌ"
        Else
            LastCredit.Caption = "ÿ·»"
            LastCredit1.Caption = -1 * LastCredit1.Caption
        End If
        
    End If
Else
    Me.FindCust
End If

End Sub

Private Sub LblDiscount_Click()
    If BlnPosApprovedWait = True Then
        ShowDisMessage "”Ì” „ „‰ Ÿ— œ—Ì«›  Å«”Œ «“ ŒÊœÅ—œ«“ „Ì »«‘œ", 1000
        Exit Sub
    End If
Dim ii As Integer
'Case " Œ›Ì›":
If ClsFormAccess.Discount <> True And AdminEdit = False Then
    frmAccess.MyFormAddEditMode = EditMode
    frmAccess.lblTitle.Caption = ". ‘„« «Ã«“Â  Œ›Ì› œ«œ‰ ‰œ«—Ìœ..»—«Ì «œ«„Â —„“ »« œ” —”Ì »«·« »“‰Ìœ"
    frmAccess.AccessStatus = EnumAccessStatus.Edit
    frmAccess.Show vbModal
    If frmAccess.ReturnAccess = False Then
        lblNum.Caption = ""
        Exit Sub
    End If
    AdminEdit = True
'    frmMsg.fwlblMsg.Caption = " . ‘„« «Ã«“Â  Œ›Ì› œ«œ‰ ‰œ«—Ìœ "
'    frmMsg.fwBtn(1).Visible = False
'    frmMsg.Show vbModal
'    Exit Sub
End If

If MyFormAddEditMode <> ViewMode Then
    frmInput.fwlblInput.Caption = "òœ«„ Õ«·   Œ›Ì› „Ê—œ ‰Ÿ— ‘„«”  "
    frmInput.OptionLevel(0).Visible = True
    frmInput.OptionLevel(1).Visible = True
    frmInput.OptionLevel(0).Caption = " Œ›Ì› —ÊÌ ›«ﬂ Ê—"
    frmInput.OptionLevel(1).Caption = " Œ›Ì› œ—’œÌ —ÊÌ ﬂ«·«Â«"
    frmInput.btnCancel.Visible = True
    frmInput.Picture1.Visible = True
    frmInput.txtInput.Visible = False
    frmInput.MyForm = Me.Name

    If clsStation.DiscountDefault = 0 Then
       frmInput.OptionLevel(0).Value = True
    Else
       frmInput.OptionLevel(1).Value = True
    End If

    frmInput.Show vbModal
    If mvarInput = "" Then
        Exit Sub
    End If
'  mvarInput = "1"
    Dim Str As String
    If Right(lblNum.Caption, 1) = "%" Then
        Str = left(lblNum.Caption, Len(lblNum.Caption) - 1)
    Else
        Str = lblNum.Caption
    End If
    If Val(Str) < 0 Then
        frmMsg.fwlblMsg.Caption = " .  Œ›Ì› „‰›Ì ﬁ«»· ﬁ»Ê· ‰Ì”  "
        frmMsg.fwBtn(1).Visible = False
        frmMsg.Show vbModal
        lblNum.Caption = 0
        Exit Sub
    End If
    
    
    If lblNum.Caption = "" Then
        Load frmMsg
        frmMsg.fwlblMsg.Caption = " . „ﬁœ«—  Œ›Ì› —« Ê«—œ ﬂ‰Ìœ "
        frmMsg.fwBtn(0).Visible = False
        frmMsg.fwBtn(1).ButtonType = flwButtonOk
        frmMsg.fwBtn(1).Caption = "ﬁ»Ê·"
        frmMsg.Show vbModal
'        frameMenu.Enabled = True
        Exit Sub
    End If
    If mvarInput = "0" Then
''''        Load frmMsg
''''        If Right(lblNum.Caption, 1) <> "%" Then
''''            frmMsg.fwlblMsg.Caption = "¬Ì« „»·€   " & Val(lblNum.Caption) & " —Ì«· »—«Ì  Œ›Ì› „Ê—œ  «∆Ìœ «” ø "
''''        Else
''''            frmMsg.fwlblMsg.Caption = "¬Ì« „Ì“«‰   " & Val(lblNum.Caption) & " œ—’œ »—«Ì  Œ›Ì› „Ê—œ  «∆Ìœ «” ø "
''''        End If
''''        frmMsg.Show vbModal
''''        If modgl.mvarMsgIdx = vbNo Then
''''            lblNum.Caption = ""
''''            Exit Sub
''''        End If
        If Right(lblNum.Caption, 1) <> "%" Then
            txtDiscount.Text = Val(lblNum.Caption)
            txtDiscountPercent.Text = "0"
        Else
            txtDiscountPercent.Text = Val(left(lblNum.Caption, Len(lblNum.Caption) - 1))
            txtDiscount.Text = "0"
            BtnKeypad(10).Enabled = True
            BtnKalaDelete.Enabled = True
            BtnKeypad(11).Enabled = True
        End If
        RefreshLables
        If CCur(lblSumPrice.Tag) < 0 Then
            Load frmMsg
            frmMsg.fwlblMsg.Caption = " . „»·€  Œ›Ì› »Ì‘ «“ „»·€ ›Ì‘ „Ì »«‘œ "
            frmMsg.fwBtn(0).Visible = False
            frmMsg.fwBtn(1).ButtonType = flwButtonOk
            frmMsg.fwBtn(1).Caption = "ﬁ»Ê·"
            frmMsg.Show vbModal
            lblDiscountTotal = 0
            txtDiscount.Text = 0
            RefreshLables
        End If
    End If
    If mvarInput = "1" Then
        If Val(lblNum.Caption) < 0 And Val(lblNum.Caption) > 100 Then
            frmDisMsg.lblMessage = " œ—’œ  Œ›Ì› »«Ìœ »Ì‰ 0 , 100 »«‘œ "
            frmDisMsg.Timer1.Enabled = True
            frmDisMsg.Show vbModal
            Exit Sub
        Else
            BtnKeypad(10).Enabled = True
            BtnKalaDelete.Enabled = True
            BtnKeypad(11).Enabled = True
            If Right(lblNum.Caption, 1) <> "%" Then
                txtDiscount.Text = Val(lblNum.Caption)
            Else
                For ii = 1 To MaxRowFlexGrid - 1
                    FlxDetail.TextMatrix(ii, IndexColDiscountPercent) = Val(lblNum.Caption)
                Next ii
            End If
        End If
        RefreshLables
        If CCur(lblSumPrice.Tag) < 0 Then
            Load frmMsg
            frmMsg.fwlblMsg.Caption = " . „»·€  Œ›Ì› »Ì‘ «“ „»·€ ›Ì‘ „Ì »«‘œ "
            frmMsg.fwBtn(0).Visible = False
            frmMsg.fwBtn(1).ButtonType = flwButtonOk
            frmMsg.fwBtn(1).Caption = "ﬁ»Ê·"
            frmMsg.Show vbModal
            For ii = 1 To MaxRowFlexGrid - 1
                FlxDetail.TextMatrix(ii, 11) = 0
            Next ii
            RefreshLables
        End If
    End If
Else    ' View mode

End If
lblNum.Caption = ""
FlxDetail.SetFocus

End Sub


Private Sub lblPacking_Click()
'Case "»” Â »‰œÌ":
    If BlnPosApprovedWait = True Then
        ShowDisMessage "”Ì” „ „‰ Ÿ— œ—Ì«›  Å«”Œ «“ ŒÊœÅ—œ«“ „Ì »«‘œ", 1000
        Exit Sub
    End If
If MyFormAddEditMode <> ViewMode Then

    If Val(lblNum.Caption) < 0 Then
        frmMsg.fwlblMsg.Caption = " . Â“Ì‰Â »” Â »‰œÌ „‰›Ì ﬁ«»· ﬁ»Ê· ‰Ì”  "
        frmMsg.fwBtn(1).Visible = False
        frmMsg.Show vbModal
        lblNum.Caption = 0
        Exit Sub
    End If
    If ClsFormAccess.packing <> True Then
        frmMsg.fwlblMsg.Caption = " . ‘„« «Ã«“Â ê—› ‰ Â“Ì‰Â »” Â »‰œÌ ‰œ«—Ìœ "
        frmMsg.fwBtn(1).Visible = False
        frmMsg.Show vbModal
        Exit Sub
    End If

    If Right(lblNum.Caption, 1) <> "%" Then
        txtPacking.Text = Val(lblNum.Caption)
      '  txtPackingPercent = 0
    Else
     '   txtPacking.Text = 0
        txtPackingPercent = Val(lblNum.Caption)
    End If
    BtnKeypad(10).Enabled = True
    BtnKalaDelete.Enabled = True
    BtnKeypad(11).Enabled = True
    'BtnKalaDelete.ForeColor = &H404080
Else
End If
    lblNum.Caption = ""
    FlxDetail.SetFocus
    RefreshLables
'    lblPackingTotal = (Val(txtSumFeeTotal.Text) * Val(txtPackingPercent.Text) / 100) + Val(txtPacking.Text)
'    lblSumPrice.Caption = Val(txtSumFeeTotal.Text) + Val(lblCarryFeeTotal.Caption) + Val(lblServiceTotal.Caption) + Val(lblPackingTotal.Caption) - Val(lblDiscountTotal.Caption)
'    lblSumPrice.Tag = lblSumPrice.Caption
'    lblSumPrice.Caption = Format(lblSumPrice.Caption, "#,##")
End Sub

Private Sub lblCarryFee_Click()
'Case "ﬂ—«ÌÂ Õ„·":
    If BlnPosApprovedWait = True Then
        ShowDisMessage "”Ì” „ „‰ Ÿ— œ—Ì«›  Å«”Œ «“ ŒÊœÅ—œ«“ „Ì »«‘œ", 1000
        Exit Sub
    End If
If MyFormAddEditMode <> ViewMode Then

    If Val(lblNum.Caption) < 0 Then
        frmMsg.fwlblMsg.Caption = " . ﬂ—«ÌÂ Õ„· „‰›Ì ﬁ«»· ﬁ»Ê· ‰Ì”  "
        frmMsg.fwBtn(1).Visible = False
        frmMsg.Show vbModal
        lblNum.Caption = 0
        Exit Sub
    End If
    If ClsFormAccess.carryfee <> True Then
        frmMsg.fwlblMsg.Caption = " . ‘„« «Ã«“Â ê—› ‰ ò—«ÌÂ Õ„· ‰œ«—Ìœ "
        frmMsg.fwBtn(1).Visible = False
        frmMsg.Show vbModal
        Exit Sub
    End If

    If Right(lblNum.Caption, 1) <> "%" Then
        txtCarryFee.Text = Val(lblNum.Caption)
    Else
        txtCarryFeePercent = Val(lblNum.Caption)
    End If
    BtnKeypad(10).Enabled = True
    BtnKalaDelete.Enabled = True
    BtnKeypad(11).Enabled = True
    'BtnKalaDelete.ForeColor = &H404080
    If Val(lblCustomer.Tag) <> 1000 And clsStation.UpDateCarryFee = True Then
        ReDim Parameter(2) As Parameter
        Parameter(0) = GenerateInputParameter("@Code", adInteger, 4, Val(lblCustomer.Tag))
        Parameter(1) = GenerateInputParameter("@NewCarryFee", adDouble, 8, Val(txtCarryFee.Text))
        Parameter(2) = GenerateInputParameter("@User", adInteger, 4, mvarCurUserNo)
        Dim Update As Long
        Update = RunParametricStoredProcedure("Update_Cust_By_NewCarryFee_FromFactor", Parameter)
         
    End If

Else
End If
    lblNum.Caption = ""
    FlxDetail.SetFocus
    RefreshLables
'    lblCarryFeeTotal = CLng((Val(txtSumFeeTotal.Text) * Val(txtCarryFeePercent.Text) / 100) + Val(txtCarryFee.Text)) ' + Val(txtCarryFeeCust.Text)
'    lblSumPrice.Caption = CLng(Val(txtSumFeeTotal.Text) + Val(lblCarryFeeTotal.Caption) + Val(lblServiceTotal.Caption) + Val(lblPackingTotal.Caption) - Val(lblDiscountTotal.Caption))
'    lblSumPrice.Tag = lblSumPrice.Caption
'    lblSumPrice.Caption = Format(lblSumPrice.Caption, "#,##")

End Sub

Private Sub lblPayFactorTotal_Change()
    
    LblRemain.Caption = CCur(lblSumPrice.Tag) - CCur(lblPayFactorTotal.Caption)
    LblRemain.Caption = "„«‰œÂ: " & Format(LblRemain.Caption, "#,## —Ì«·")
    
End Sub

Private Sub lblService_Click()
'Case "”—ÊÌ”":
    If BlnPosApprovedWait = True Then
        ShowDisMessage "”Ì” „ „‰ Ÿ— œ—Ì«›  Å«”Œ «“ ŒÊœÅ—œ«“ „Ì »«‘œ", 1000
        Exit Sub
    End If
If MyFormAddEditMode <> ViewMode Then

    If Val(lblNum.Caption) < 0 Then
        frmMsg.fwlblMsg.Caption = " .  ”—ÊÌ” „‰›Ì ﬁ«»· ﬁ»Ê· ‰Ì”  "
        frmMsg.fwBtn(1).Visible = False
        frmMsg.Show vbModal
        lblNum.Caption = 0
        Exit Sub
    End If
    If ClsFormAccess.service <> True Then
        frmMsg.fwlblMsg.Caption = " . ‘„« «Ã«“Â ê—› ‰ ”—ÊÌ” ‰œ«—Ìœ "
        frmMsg.fwBtn(1).Visible = False
        frmMsg.Show vbModal
        Exit Sub
    End If

    If Right(lblNum.Caption, 1) <> "%" Then
        frmMsg.fwlblMsg.Caption = " . ”—ÊÌ” »«Ìœ »Â ’Ê—  œ—’œÌ Ê«—œ ‘Êœ "
        frmMsg.fwBtn(1).Visible = False
        frmMsg.Show vbModal
        Exit Sub
    Else
        ServiceRate = Val(lblNum.Caption)
        EnableDefaultServiceRate = False
    End If
    BtnKeypad(10).Enabled = True
    BtnKalaDelete.Enabled = True
    BtnKeypad(11).Enabled = True
    'BtnKalaDelete.ForeColor = &H404080
Else
End If
    lblNum.Caption = ""
    FlxDetail.SetFocus
'    lblServiceTotal.Caption = CLng((Val(txtSumFeeTotal.Text) * ServiceRate / 100))
'    lblSumPrice.Caption = CLng(Val(txtSumFeeTotal.Text) + Val(lblCarryFeeTotal.Caption) + Val(lblServiceTotal.Caption) + Val(lblPackingTotal.Caption) - Val(lblDiscountTotal.Caption))
'    lblSumPrice.Tag = lblSumPrice.Caption
'    lblSumPrice.Caption = Format(lblSumPrice.Caption, "#,##")
    RefreshLables
End Sub

Private Sub FlxDetail_EnterCell()
    With FlxDetail
        If .Row > 0 And .TextMatrix(.Row, 5) <> "" And (.Col = 3) Then
            If ClsFormAccess.CustomizeFee = True Then
               .Select .Row, .Col
               .EditCell
            ElseIf MyFormAddEditMode <> EnumAddEditMode.ViewMode Then
               frmDisMsg.lblMessage = "  ‘„« »Â «„ﬂ«‰   €ÌÌ— ›Ì œ” —”Ì ‰œ«—Ìœ "
               frmDisMsg.Timer1.Enabled = True
               frmDisMsg.Show vbModal
            End If
               
        End If
        
        If .Row > 0 And .TextMatrix(.Row, 5) <> "" And .Col = 12 Then
               .Select .Row, .Col
               .EditCell
        End If
    End With

End Sub


Private Sub FlxDetail_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    
    If (IsNumeric(Chr(KeyAscii)) = False And KeyAscii <> 8 And KeyAscii <> 45 And FlxDetail.Col <> 10) Or mvarbarcode = True Then
       KeyAscii = 0
     ElseIf KeyAscii = 13 And Me.ActiveControl.Name = FlxDetail.Name And FlxDetail.Col = IndexColDifferences And Len(FlxDetail.TextMatrix(FlxDetail.Row, IndexColDifferences)) > 0 And FlxDetail.TextMatrix(FlxDetail.Row, 9) = "" Then
     
         If SaveDifferences > 0 Then FlxDetail.TextMatrix(FlxDetail.Row, 10) = "": FlxDetail.Select FlxDetail.Row, IndexColDifferences: FlxDetail_Click  ''':ShowDisMessage "À»   €ÌÌ—«  - " & FlxDetail.TextMatrix(FlxDetail.Row, IndexColGoodName) & "  -Ê ò«·«Â«Ì „‘«»Â ¬‰ «‰Ã«„ ‘œ", 1000
    End If

End Sub

Private Sub FlxDetail_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode > 32 And KeyCode < 37 Then
        KeyActi vbtxtbox, KeyCode, Shift, Me
        FlxDetail.ShowCell 1, 1
    End If
End Sub

'''Private Sub FlxDetail_KeyUpEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
'''    With FlxDetail
'''        .Row = Row
'''        .Col = Col
'''        If Val(FlxDetail.TextMatrix(FlxDetail.Row, 1)) <> 0 Then
'''            FlxDetail.TextMatrix(FlxDetail.Row, 4) = CLng(Val(FlxDetail.TextMatrix(FlxDetail.Row, 1)) * Val(FlxDetail.TextMatrix(FlxDetail.Row, 3)))
'''        End If
'''        RefreshLables
'''    End With
'''End Sub

Private Sub FlxDetail_LeaveCell()
    With FlxDetail
        If .Row > 0 And .Row < MaxRowFlexGrid And .Col = IndexColFee Then
            If IsNumeric(.TextMatrix(.Row, .Col)) <> True Or Trim(.TextMatrix(.Row, .Col)) = "" Then
                .TextMatrix(.Row, .Col) = 0
            End If
        End If
    End With
End Sub

Private Sub FlxDetail_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With FlxDetail
        .Row = Row
        .Col = Col
        If Val(FlxDetail.TextMatrix(FlxDetail.Row, IndexColAmount)) <> 0 Then
            FlxDetail.TextMatrix(FlxDetail.Row, IndexColTotalFee) = CCur(Val(FlxDetail.TextMatrix(FlxDetail.Row, IndexColAmount)) * Val(FlxDetail.TextMatrix(FlxDetail.Row, IndexColFee)))
        End If
        RefreshLables
    End With
End Sub

Private Sub Form_Activate()
       
    IsPrinting = False
    VarActForm = Me.Name
    
    If clsArya.Customers = False Then
        lblCustomer.Enabled = False
    End If
    If clsArya.Delivery = False Then
        FWBtnPayk.Enabled = False
    End If
    Dim mm As Integer
    For mm = 0 To 7
        FWModem(mm).BackColor = &H80000016  '&H808000
    Next mm
    
    mvarStatus = Invoice
    Add

End Sub

Public Sub Printing()
    On Error GoTo Err_Handler
    
    If Me.ChkIsLocked.Value <> 0 And Not (strCategory = "04" And strDelegate = "00" And clsArya.CustomerId = 102) And Not (strCategory = "00" And strDelegate = "00" And clsArya.CustomerId = 221) Then 'Naghsh_Jahan & Spu Then
        frmDisMsg.lblMessage.Caption = "”‰œ ﬁ›· ‘œÂ «”  Ê «„ò«‰ ç«Å ¬‰ ÊÃÊœ ‰œ«—œ."
        frmDisMsg.Timer1.Interval = 2000
        frmDisMsg.Timer1.Enabled = True
        frmDisMsg.Show vbModal
        Exit Sub
    End If
    
    Dim CountPrinting As Integer, CountRePrint As Integer, CountInvoicePrint As Integer
    
    If ClsFormAccess.Printing = True Then

    Dim intPrintFichNo As Long
    Dim strCommand As String
    Dim s As String
    Dim tempTxtNo As Long
    Dim TempMyFormAddEditMode As EnumAddEditMode
    
    TempMyFormAddEditMode = MyFormAddEditMode
    
    If tempTxtNo = 0 Then
        tempTxtNo = txtNo.Text
    End If
    
    ReDim Parameter(2) As Parameter
    Parameter(0) = GenerateInputParameter("@No", adBigInt, 8, Val(txtNo.Text))
    Parameter(1) = GenerateInputParameter("@Status", adInteger, 4, mvarStatus)
    Parameter(2) = GenerateInputParameter("@AccountYear", adSmallInt, 2, AccountYear)
    
    Set Rst = RunParametricStoredProcedure2Rec("Get_CountPrint_tAction", Parameter)
    
    CountPrinting = Rst!CountPrinting
    CountRePrint = Rst!CountRePrint
    CountInvoicePrint = Rst!CountInvoicePrint
    
    frmInput.OptionLevel(0).Visible = True
    frmInput.OptionLevel(1).Visible = True
    frmInput.OptionLevel(0).Caption = "ç«Å „Ãœœ"
    frmInput.OptionLevel(1).Caption = "ç«Å ›«ò Ê— ›—Ê‘"
    
    If MaxRowFlexGrid >= 2 Then     ' Fich Is Not Empty
       Select Case MyFormAddEditMode
           Case ViewMode
           
                If mvarStatus = Invoice Then
                    frmInput.fwlblInput.Caption = "òœ«„ Õ«·  ç«Å „Ê—œ ‰Ÿ— ‘„«”  "
                    frmInput.btnCancel.Visible = True
                    frmInput.Picture1.Visible = True
                    frmInput.txtInput.Visible = False
                    frmInput.MyForm = Me.Name
                    If clsStation.ReprintDefault = 0 Then
                       frmInput.OptionLevel(0).Value = True
                    Else
                       frmInput.OptionLevel(1).Value = True
                    End If
                    
                    frmInput.Show vbModal
                    If mvarInput = "" Then
                        Exit Sub
                    End If
                Else
                    mvarInput = "0"
                End If
                If mvarInput = "0" Then
                    tempTxtNo = txtNo.Text
                    If ClsFormAccess.Reprint = False Then
                        frmDisMsg.lblMessage = "  ‘„« »Â «„ﬂ«‰ ç«Å „Ãœœ œ” —”Ì ‰œ«—Ìœ "
                        frmDisMsg.Timer1.Enabled = True
                        frmDisMsg.Show vbModal
                        Exit Sub
                    End If
                    If mvarCountRePrint <= CountRePrint Then
                        frmMsg.fwlblMsg.Caption = " ⁄œ«œ ç«Å „Ãœœ ‘„« —ÊÌ «Ì‰ ›«ﬂ Ê—  „«„ ‘œÂ"
                        frmMsg.fwBtn(0).Visible = False
                        frmMsg.fwBtn(1).ButtonType = flwButtonOk
                        frmMsg.fwBtn(1).Caption = "ﬁ»Ê·"
                        frmMsg.Show vbModal
                        Exit Sub
                    End If
                    ActionMode = EnumActionLog.Reprint
'                    If clsArya.PrintServer = False Then  ' Check in Printing Routine
                         If clsArya.NewPrinting = False Then
                              IsPrinting = ClsPrint.Printing(Val(txtNo.Text), clsArya.StationNo, MyFormAddEditMode, ActionMode)
                         Else
                              IsPrinting = StimulPrn.Printing(Val(txtNo.Text), clsArya.StationNo, MyFormAddEditMode, clsStation.Language, AccountYear, clsStation.PartitionID, CurrentBranch, ActionMode, mvarCurUserNo, "Reports" & RepVer)      '
     
                         End If
'                    Else
'                        intPrintFichNo = ClsPrint.InsertPrintFich(CLng(Val(txtNo.Text)), clsArya.StationNo, MyFormAddEditMode, ActionMode)
'                        strCommand = Str(intPrintFichNo)
'                        mdifrm.Winsock_Print.SendData strCommand
'                    End If
                    If clsStation.StopOnEditFich = False Then
                        Add
                    End If
                    Exit Sub
                Else
                    If ClsFormAccess.InvoicePrint = False Then
                        frmDisMsg.lblMessage = "  ‘„« »Â «„ﬂ«‰ ç«Å ›«ﬂ Ê—›—Ê‘ œ” —”Ì ‰œ«—Ìœ "
                        frmDisMsg.Timer1.Enabled = True
                        frmDisMsg.Show vbModal
                        Exit Sub
                    End If
                    If mvarCountInvoicePrint <= CountInvoicePrint Then
                        frmMsg.fwlblMsg.Caption = " ⁄œ«œ ç«Å ›«ﬂ Ê—›—Ê‘ ‘„« —ÊÌ «Ì‰ ›«ﬂ Ê—  „«„ ‘œÂ"
                        frmMsg.fwBtn(0).Visible = False
                        frmMsg.fwBtn(1).ButtonType = flwButtonOk
                        frmMsg.fwBtn(1).Caption = "ﬁ»Ê·"
                        frmMsg.Show vbModal
                        Exit Sub
                    End If
                    ActionMode = EnumActionLog.InvoicePrint
'                    If clsArya.PrintServer = False Then ' Check in Printing Routine
                        IsPrinting = ClsPrint.Printing(Val(txtNo.Text), clsArya.StationNo, InvoiceFactor, ActionMode)
'                    Else
'                        intPrintFichNo = ClsPrint.InsertPrintFich(CLng(Val(txtNo.Text)), clsArya.StationNo, InvoiceFactor, ActionMode)
'                        strCommand = Str(intPrintFichNo)
'                        mdifrm.Winsock_Print.SendData strCommand
'                    End If
                    If clsStation.StopOnEditFich = False Then
                        Add
                    End If
                    Exit Sub
                        
                End If ' end of Print Type Selection (Reprint, Print, Bijak)
                
            Case AddMode
                tempTxtNo = Update
                If tempTxtNo = -1 Then
                    Exit Sub
                End If
                ActionMode = EnumActionLog.Printing
            Case EditMode, ManipulateMode, RefferedMode
                Dim tempReffered As Integer
                tempReffered = txtRecursive.Text
                tempTxtNo = Update
                If tempTxtNo = -1 Then
                    Exit Sub
                End If
                
                If TempMyFormAddEditMode = RefferedMode Then
                
                    If tempTxtNo <> -1 And tempReffered = 1 Then
                    
'                        fwlblRecursive.Visible = True
                        frmDisMsg.lblMessage = "›Ì‘ „—ÃÊ⁄Ì "
                        frmDisMsg.Timer1.Enabled = True
                        frmDisMsg.Show vbModal
                        
                    ElseIf tempTxtNo = -1 And tempReffered = 1 Then
                    
                        frmMsg.fwlblMsg.Caption = "›Ì‘ „—ÃÊ⁄ ‰‘œ"
                        frmMsg.fwBtn(0).Visible = False
                        frmMsg.fwBtn(1).ButtonType = flwButtonOk
                        frmMsg.fwBtn(1).Caption = "ﬁ»Ê·"
                        frmMsg.Show vbModal
                        
                    ElseIf tempTxtNo <> -1 And tempReffered = 0 Then
                    
'                        fwlblRecursive.Visible = False
                        frmDisMsg.lblMessage = " »«“ê—œ«‰Ì ›Ì‘ „—ÃÊ⁄Ì "
                        frmDisMsg.Timer1.Enabled = True
                        frmDisMsg.Show vbModal
                        
                    ElseIf tempTxtNo = -1 And tempReffered = 0 Then
                    
                        frmMsg.fwlblMsg.Caption = "›Ì‘ „—ÃÊ⁄Ì »«“ê—œ«‰Ì ‰‘œ"
                        frmMsg.fwBtn(0).Visible = False
                        frmMsg.fwBtn(1).ButtonType = flwButtonOk
                        frmMsg.fwBtn(1).Caption = "ﬁ»Ê·"
                        frmMsg.Show vbModal
                        
                    End If
                    
                End If
                ActionMode = EnumActionLog.Printing
        End Select
      
    Else
        Exit Sub
    End If
    
    If tempTxtNo > 0 Then
        DoEvents
'        If clsArya.PrintServer = False Then ' Check in Printing Routine
            If clsArya.NewPrinting = False Then
               IsPrinting = ClsPrint.Printing(tempTxtNo, clsArya.StationNo, TempMyFormAddEditMode, ActionMode)
            Else
               IsPrinting = StimulPrn.Printing(tempTxtNo, clsArya.StationNo, TempMyFormAddEditMode, clsStation.Language, AccountYear, clsStation.PartitionID, CurrentBranch, ActionMode, mvarCurUserNo, "Reports" & RepVer)                 ', ActionMode, mvarCurUserNo, "\Reports" & RepVer
            End If
            If clsStation.LabelPrint = True Then
                IsPrinting = ClsPrint.Printing(tempTxtNo, clsArya.StationNo, EnumAddEditMode.Perfrage, ActionMode)
            End If
'        Else
'           intPrintFichNo = ClsPrint.InsertPrintFich(CLng(tempTxtNo), clsArya.StationNo, TempMyFormAddEditMode, ActionMode)
'           strCommand = Str(intPrintFichNo)
'           mdifrm.Winsock_Print.SendData strCommand
'        End If
            ' Because in view mode has add and exit sub
'            If MyFormAddEditMode = ViewMode Then    ' Because set in update mode
'               If clsStation.StopOnEditFich = False Then
'                   Add
'               Else
'                   MyFormAddEditMode = ViewMode
'                   SetFirstToolBar
'                   GetDataDetail
'                   RefreshLables
'               End If
'            End If
'        End If
    End If
    If clsStation.InvoiceStatusDefault = True Then
        mvarStatus = EnumFactorType.Invoice
        If clsStation.Language = Farsi Then
            LblInvoice.Caption = "›«ﬂ Ê— ›—Ê‘"
        Else
            LblInvoice.Caption = "Invoice"
        End If
        If clsStation.PayFactorView = True Then
            cmdPayFactor.Visible = True
            lblPayFactorTotal.Visible = True
        Else
            cmdPayFactor.Visible = False
            lblPayFactorTotal.Visible = False
        End If
    End If
    
Else
    frmDisMsg.lblMessage = "  ‘„« »Â «„ﬂ«‰  ç«Å œ” —”Ì ‰œ«—Ìœ "
    frmDisMsg.Timer1.Enabled = True
    frmDisMsg.Show vbModal

End If

If clsStation.Frame_Printers = True Then
    Timer_Printers_Timer
End If
    
Exit Sub

Err_Handler:
    LogSaveNew "frmInvoice => ", err.Description, err.Number, err.Source, "Printing"
    ShowErrorMessage
    
End Sub
''Private Sub PrintLable(FichNo As Long)
''
''    On Error GoTo ErrHandler
''
''    If Printerprint(FichNo, clsArya.StationNo) = False Then Exit Sub
''
''    ReDim Parameter(4)
''    Parameter(0) = GenerateInputParameter("@No", adBigInt, 8, FichNo)
''    Parameter(1) = GenerateInputParameter("@Status", adInteger, 4, mvarStatus)
''    Parameter(2) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
''    Parameter(3) = GenerateInputParameter("@AccountYear", adSmallInt, 2, AccountYear)
''    Parameter(4) = GenerateInputParameter("@Branch", adInteger, 4, CurrentBranch)
''
''    Set Rst = RunParametricStoredProcedure2Rec("Get_FacMD_Good", Parameter)
''
''    Dim ii, kk, jj, MaxCount, UsedCount As Long
''    Dim strGood, strDescription As String
''    ii = 0
''    If Not (Rst.BOF Or Rst.EOF) Then
''        Do While Not (Rst.EOF)
''
''            MaxCount = Rst!SumAmount
''            For kk = 1 To Rst!amount
''                If clsStation.LableUsedGood = False Then
''                    ii = ii + 1
''                    strGood = Rst!nvcName & " " & IIf(IsNull(Rst!DifferencesDescription), "", Rst!DifferencesDescription)
''                    strDescription = ii & " of " & MaxCount
''                    CrystalPrint FichNo, strGood, strDescription
''                Else
''                    ReDim Parameter(0)
''                    Parameter(0) = GenerateInputParameter("@GoodCode", adInteger, 4, Rst!GoodCode)
''                    Set rctmp = RunParametricStoredProcedure2Rec("Get_Good_Used", Parameter)
''                    If Not (rctmp.BOF Or rctmp.EOF) Then
''                        jj = 0
''                        ii = ii + 1
''                        UsedCount = rctmp!UsedCount
''                        Do While Not (rctmp.EOF)
''                            jj = jj + 1
''                            strGood = rctmp!nvcName
''                            strDescription = ii & " of " & MaxCount & "-" & jj & " of " & UsedCount
''                            CrystalPrint FichNo, strGood, strDescription
''                            rctmp.MoveNext
''                        Loop
''                    Else        '''ç«Å Â„«‰ ò«·«Ì «’·Ì
''                        ii = ii + 1
''                        strGood = Rst!nvcName & " " & IIf(IsNull(Rst!DifferencesDescription), "", Rst!DifferencesDescription)
''                        strDescription = ii & " of " & MaxCount
''                        CrystalPrint FichNo, strGood, strDescription
''                    End If
''                End If
''            Next
''            Rst.MoveNext
''        Loop
''    End If
''Exit Sub
''ErrHandler:
''    ShowDisMessage err.Description, 2000
''End Sub
''
'' Private Sub CrystalPrint(ByVal FichNo As Long, ByVal strGood As String, ByVal strDescription As String)
''
''    On Error GoTo ErrHandler
''        Dim ArrayUbound  As Integer
''        ReDim Parameter(2) As Parameter
''
'''        Parameter(0) = GenerateInputParameter("@SystemDate", adVarWChar, 50, Mid(clsDate.shamsi(Date), 3, 8))
'''        Parameter(1) = GenerateInputParameter("@SystemDay", adVarWChar, 50, clsDate.Find_DayOfWeek(Weekday(Date, vbSaturday)))
'''        Parameter(2) = GenerateInputParameter("@SystemTime", adVarWChar, 50, Mid(Str(time), 1, 5))
''        Parameter(0) = GenerateInputParameter("@FichNo", adInteger, 4, FichNo)
''        Parameter(1) = GenerateInputParameter("@strGood", adVarWChar, Len(strGood) + 1, strGood)
''        Parameter(2) = GenerateInputParameter("@strDescription", adVarWChar, Len(strDescription) + 1, strDescription)
''
''        CrystalReport1.ReportTitle = clsArya.StationName
''        Dim intIndex As Integer
''
''        For intIndex = 0 To UBound(Parameter) - LBound(Parameter)
''            CrystalReport1.ParameterFields(intIndex) = CStr(Parameter(intIndex).Name) & ";" & CStr(Parameter(intIndex).Value) & ";" & "True"
''        Next intIndex
''        ODBCSetting clsArya.ServerName, clsArya.DbName
''        CrystalReport1.Connect = CrystallConnection
''        CrystalReport1.RetrieveDataFiles
''        CrystalReport1.ProgressDialog = False
''        CrystalReport1.Action = 1
''
''Exit Sub
''ErrHandler:
''    ShowDisMessage err.Description, 1500
''End Sub
''
''Private Function Printerprint(FichNumber As Long, intStationId As Integer) As Boolean
''    On Error GoTo ErrHandler
''    Dim RstTemp As New ADODB.Recordset
''    Dim RstTemp2 As New ADODB.Recordset
''    Dim j As Long
''    Dim PartitionNo As Long
''    Dim ParametersTmp(3) As Parameter
''    ParametersTmp(0) = GenerateInputParameter("@Status", adInteger, 4, mvarStatus)
''    ParametersTmp(1) = GenerateInputParameter("@No", adBigInt, 8, FichNumber)
''    ParametersTmp(2) = GenerateInputParameter("@AccountYear", adSmallInt, 2, AccountYear)
''    ParametersTmp(3) = GenerateInputParameter("@Branch", adInteger, 4, CurrentBranch)
''
''    Set RstTemp2 = RunParametricStoredProcedure2Rec("GetFacMinfo", ParametersTmp)
''
''    If Not (RstTemp2.EOF = True And RstTemp2.BOF = True) Then
''        If RstTemp2!PartitionID > 0 Then
''            PartitionNo = RstTemp2!PartitionID
''        Else
''            PartitionNo = clsStation.PartitionID
''        End If
''        Dim Parameters(5) As Parameter
''        Parameters(0) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
''        Parameters(1) = GenerateInputParameter("@ServePlace", adInteger, 4, RstTemp2!ServePlace)
''        Parameters(2) = GenerateInputParameter("@CurrentStationId", adInteger, 4, intStationId)
''        Parameters(3) = GenerateInputParameter("@PartitionId", adInteger, 4, PartitionNo)
''        Parameters(4) = GenerateInputParameter("@Mode", adInteger, 4, Perfrage)
''        Parameters(5) = GenerateInputParameter("@Status", adInteger, 4, mvarStatus)
''        If RstTemp.State <> 0 Then RstTemp.Close
''        Set RstTemp = RunParametricStoredProcedure2Rec("GetPrintInfo", Parameters)
''
''        Dim ReportFileName  As String
''
''        Dim fileSystem As New FileSystemObject
''        Dim IsFileExist As Boolean
''
''        If Not (RstTemp.BOF Or RstTemp.EOF) Then
''            If RstTemp.Fields("PermittedModes").Value = Perfrage Then
''
''                ReportFileName = App.Path & "\Reports" & RepVer & "\" & RstTemp!rptFilePath
''                If fileSystem.FileExists(ReportFileName) = False Then
''                    ShowDisMessage " ›«Ì· »—«Ì ç«Å Å—›—«é -" & CrystalReport1.ReportFileName & "ÅÌœ« ‰‘œ ", 2000
''                    Printerprint = False
''                Else
''                    Printerprint = Doprint(RstTemp!PrinterName, ReportFileName)
''                End If
''            End If
''
''        Else
''            ShowDisMessage " „Ê—œ ç«Å »—«Ì Õ«·  Å—›—«é  ⁄—Ì› ‰‘œÂ «”   ", 2000
''            Printerprint = False
''        End If
''    End If
''    Set RstTemp = Nothing
''    Set RstTemp2 = Nothing
''Exit Function
''ErrHandler:
''    ShowDisMessage err.Description, 1500
''
''End Function
''
''Private Function Doprint(PassedPrinterName As String, ReportFileName As String) As Boolean
''
''    On Error GoTo ErrHandler
''    Dim prnPrinter
''    Dim IsPrinterNameOk As Boolean
''    For Each prnPrinter In Printers
''        If InStr(1, prnPrinter.DeviceName, PassedPrinterName, 1) Then
''            Set Printer = prnPrinter
''            IsPrinterNameOk = True
''            Exit For
''        End If
''    Next
''    If IsPrinterNameOk = False Then
''        MsgBox PassedPrinterName & "Å—Ì‰ — Å—›—«é œ—·Ì”  Å—Ì‰ —Â« „ÊÃÊœ ‰Ì” "
''
''        Doprint = False
''        Exit Function
''    End If
''    CrystalReport1.ReportFileName = ReportFileName
''    CrystalReport1.PrinterDriver = Printer.DriverName
''    CrystalReport1.PrinterName = Printer.DeviceName
''    CrystalReport1.PrinterPort = Printer.Port
''
''    CrystalReport1.PrinterCopies = 1
''    'CrystalReport1.Destination = crptowindow    'crptowindow
''    CrystalReport1.Destination = crptToPrinter
''
''    Doprint = True
''Exit Function
''ErrHandler:
''    ShowDisMessage err.Description, 2000
''    Doprint = False
''End Function
Public Sub GetCustBarCode()
    If clsArya.Customers = True Then
        If DropDownFlag = False Then
            On Error GoTo ErrorHandler
            frmGetCustBarcode.Show vbModal
            
            If mvarcode <> 0 Then
                lblCustomer.Tag = mvarcode
                mvarcode = 0
                mVarOrderType = mvarPublicOrderType
                mvarPublicOrderType = inPerson
            Else
                lblCustomer.Tag = -1
                mvarPublicOrderType = inPerson
                mVarOrderType = inPerson
            End If
            If mVarOrderType = ByPhone Then
               If clsStation.Language = Farsi Then
                    LblOrder.Caption = " ·›‰Ì"
               Else
                    LblOrder.Caption = "By Phone"
               End If
            Else
              If clsStation.Language = Farsi Then
               LblOrder.Caption = "Õ÷Ê—Ì"
              Else
               LblOrder.Caption = "Inside"
               End If
            End If
            UpdatelblCustomer
            RefreshLables
       End If
     Else
                    
        frmDisMsg.lblMessage.Caption = " ‘„« »Â «Ì‰ «„ﬂ«‰ œ” —”Ì ‰œ«—Ìœ "
        frmDisMsg.Timer1.Interval = 1000
        frmDisMsg.Timer1.Enabled = True
        frmDisMsg.Show vbModal
       
     End If
Exit Sub
ErrorHandler:
    frmFindCust.txtMembershipId.Text = CreditCode
    VarActForm = "frmInvoice"
End Sub

Private Sub BtnMenu_Click(index As Integer)
    Call PresetScreenSaver
    If BlnPosApprovedWait = True Then
        ShowDisMessage "”Ì” „ „‰ Ÿ— œ—Ì«›  Å«”Œ «“ ŒÊœÅ—œ«“ „Ì »«‘œ", 1000
        Exit Sub
    End If
    If MyFormAddEditMode = ViewMode Then
        Exit Sub
    End If
    
    Dim var1, var2 As Double
    Dim j As Double
            
        
    BtnKeypad(10).Enabled = True
    BtnKalaDelete.Enabled = True
    BtnKeypad(11).Enabled = True
    
    If InStr(1, BtnMenu(index).Tag, ";", 1) > 0 Then
        Call frmFindGoods_Menu.SendVariables(index)
        frmFindGoods_Menu.Show vbModal
    Else
        If GetGoodCode(Val(BtnMenu(index).Tag)) = True Then
            ChangeGoodquantity
            If clsStation.ShowOption = True Then
                Call frmFindGoods_Difference.SendVariables(Val(BtnMenu(index).Tag))
                frmFindGoods_Difference.Show vbModal
            End If
        End If
    End If

    FlxDetail.SetFocus
    
End Sub
Private Sub BtnFindGood_Click()
    Call PresetScreenSaver
    If BlnPosApprovedWait = True Then
        ShowDisMessage "”Ì” „ „‰ Ÿ— œ—Ì«›  Å«”Œ «“ ŒÊœÅ—œ«“ „Ì »«‘œ", 1000
        Exit Sub
    End If
    If MyFormAddEditMode = ViewMode Then
        Exit Sub
    End If
    
    BtnKeypad(10).Enabled = True
    BtnKalaDelete.Enabled = True
    BtnKeypad(11).Enabled = True
    
    frmFindGoods.Show vbModal

    FlxDetail.SetFocus
    
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Call PresetScreenSaver
    If Me.ActiveControl.Name = TxtGuestNo.Name Then Exit_Keypress_Flag = True: Exit Sub
    If textDescription = True Then Exit_Keypress_Flag = True: Exit Sub
    If AddressFlag = True Then Exit_Keypress_Flag = True: Exit Sub
    If CustDescriptionFlag = True Then Exit_Keypress_Flag = True: Exit Sub
    If textTempAddressFlag = True Then Exit_Keypress_Flag = True: Exit Sub
    
    mvarKeyCode = KeyCode
    MvarShiftKey = Shift
    Exit_Keypress_Flag = False
    
    Dim BarcodeLengh As Integer
    BarcodeLengh = 15
    
    KeyActi vbtxtbox, KeyCode, Shift, Me
    
    If (KeyCode < 32 And KeyCode <> 13 And KeyCode <> 27 And KeyCode <> 8 And KeyCode <> 16) Then
        Exit_Keypress_Flag = True
        Exit Sub
    End If
    'For Sedasima Mashad (; in start string)
    If (Shift = 0 And KeyCode = 186) Then
        Form_KeyPress 59
        Exit_Keypress_Flag = True
        Exit Sub
    End If
    
    Select Case Shift
    
        Case 0
            Select Case KeyCode
                
                Case 187          '= Beep Key

                    Exit_Keypress_Flag = True

                Case vbKeyF3        'Edit Mode
                
                    Exit_Keypress_Flag = True
                Case vbKeyF4

                   lblCustomer_Click
                    Exit_Keypress_Flag = True
                  
                Case vbKeyF6    'Printing
                    Exit_Keypress_Flag = True
                
                Case vbKeyF7
                    
                Case vbKeyF8
                
                    cmdPay_Click
                    Exit_Keypress_Flag = True
                    
                Case vbKeyF9        'Customize Fee
                
                    With FlxDetail

                        If .Col <> 3 Then
                            .Col = 3
                        End If
                        If .TextMatrix(.Row, 1) = "" Then
                            .Row = MaxRowFlexGrid - 1
                        End If
                        .ShowCell .Row, .Col

                    End With
                    FlxDetail_Click

                    Exit_Keypress_Flag = True
                
                Case vbKeyF10
                
                    If ClsFormAccess.frmPayk = True And clsArya.Delivery = True Then
                        frmPayk.Show
                    Else
                        frmDisMsg.lblMessage.Caption = " ‘„« »Â «Ì‰ «„ﬂ«‰ œ” —”Ì ‰œ«—Ìœ "
                        frmDisMsg.Timer1.Interval = 1000
                        frmDisMsg.Timer1.Enabled = True
                        frmDisMsg.Show vbModal

                    End If
                    Sendkey "{Tab}", True
                    Exit_Keypress_Flag = True
                
                Case vbKeyF11   'Screen Saver
                    
                    mdifrm.CallScreenSaver
                
                Case vbKeyF12   '
                     
                    cmdTables_Click
                    Exit_Keypress_Flag = True
                
                Case vbKeySubtract, 189 '-
                    If MyFormAddEditMode = ViewMode Then
                       UndoRedo
                    Else
                       BtnKalaDelete_Click
                    End If
                    Exit_Keypress_Flag = True
                
                Case vbKeyDecimal, 190          ' .
                
                    If BtnKeypad(10).Enabled Then
                        BtnKeypad_Click (10)
                    End If
                    Exit_Keypress_Flag = True
                    
                Case vbKeyDivide, 191 '/ Barcode
                    If strCategory = "00" And clsArya.CustomerId = 259 And strDelegate = "00" Then
                        GetCustBarCode
                    Else
                        If mvarbarcode = False Then
                            mvarbarcode = True
                        Else
                            barcode
                        End If
                    End If
                    Exit_Keypress_Flag = True
                    
                Case vbKey0, vbKeyNumpad0                '0
                
                    If Not (mvarbarcode) Then
                        BtnKeypad_Click (0)
                    Else
                        lblBarCode = lblBarCode & "0"
                        If Len(lblBarCode) > BarcodeLengh Then
                            lblBarCode = ""
                            mvarbarcode = False
                        End If
                    End If
                    Exit_Keypress_Flag = True
                    
                Case vbKey1, vbKeyNumpad1                ' 1
                
                    If Not (mvarbarcode) Then
                        BtnKeypad_Click (1)
                    Else
                        lblBarCode = lblBarCode & "1"
                        If Len(lblBarCode) > BarcodeLengh Then
                            lblBarCode = ""
                            mvarbarcode = False
                        End If
                    End If
                    Exit_Keypress_Flag = True
                    
                Case vbKey2, vbKeyNumpad2                '2
                
                    If Not (mvarbarcode) Then
                        BtnKeypad_Click (2)
                    Else
                        lblBarCode = lblBarCode & "2"
                        If Len(lblBarCode) > BarcodeLengh Then
                            lblBarCode = ""
                            mvarbarcode = False
                        End If
                    End If
                    Exit_Keypress_Flag = True
                    
                Case vbKey3, vbKeyNumpad3               '3
                
                    If Not (mvarbarcode) Then
                        BtnKeypad_Click (3)
                    Else
                        lblBarCode = lblBarCode & "3"
                        If Len(lblBarCode) > BarcodeLengh Then
                            lblBarCode = ""
                            mvarbarcode = False
                        End If
                    End If
                    Exit_Keypress_Flag = True
                    
                Case vbKey4, vbKeyNumpad4               '4
                
                    If Not (mvarbarcode) Then
                        BtnKeypad_Click (4)
                    Else
                        lblBarCode = lblBarCode & "4"
                        If Len(lblBarCode) > BarcodeLengh Then
                            lblBarCode = ""
                            mvarbarcode = False
                        End If
                    End If
                    Exit_Keypress_Flag = True
                    
                Case vbKey5, vbKeyNumpad5  '5
            
                    If Not (mvarbarcode) Then
                        BtnKeypad_Click (5)
                    Else
                        lblBarCode = lblBarCode & "5"
                        If Len(lblBarCode) > BarcodeLengh Then
                            lblBarCode = ""
                            mvarbarcode = False
                        End If
                    End If
                    Exit_Keypress_Flag = True
                    
                Case vbKey6, vbKeyNumpad6   '6
                
                    If Not (mvarbarcode) Then
                        BtnKeypad_Click (6)
                    Else
                        lblBarCode = lblBarCode & "6"
                        If Len(lblBarCode) > BarcodeLengh Then
                            lblBarCode = ""
                            mvarbarcode = False
                        End If
                    End If
                    Exit_Keypress_Flag = True
                    
                Case vbKey7, vbKeyNumpad7     '7
                
                    If Not (mvarbarcode) Then
                        BtnKeypad_Click (7)
                    Else
                        lblBarCode = lblBarCode & "7"
                        If Len(lblBarCode) > BarcodeLengh Then
                            lblBarCode = ""
                            mvarbarcode = False
                        End If
                    End If
                    Exit_Keypress_Flag = True
                    
                Case vbKey8, vbKeyNumpad8      '8
                
                    If Not (mvarbarcode) Then
                        BtnKeypad_Click (8)
                    Else
                        lblBarCode = lblBarCode & "8"
                        If Len(lblBarCode) > BarcodeLengh Then
                            lblBarCode = ""
                            mvarbarcode = False
                        End If
                    End If
                    Exit_Keypress_Flag = True
                    
                Case vbKey9, vbKeyNumpad9      '9
                    
                    If Not (mvarbarcode) Then
                        BtnKeypad_Click (9)
                    Else
                        lblBarCode = lblBarCode & "9"
                        If Len(lblBarCode) > BarcodeLengh Then
                            lblBarCode = ""
                            mvarbarcode = False
                        End If
                    End If
                    Exit_Keypress_Flag = True
                    
                Case vbKeyBack
                
                    If Len(Trim(lblBarCode.Caption)) >= 1 Then
                        lblBarCode.Caption = left(lblBarCode.Caption, Len(Trim(lblBarCode.Caption)) - 1)
                    ElseIf Len(Trim(lblNum.Caption)) >= 1 Then
                        If Right(lblNum.Caption, 1) = "." Then
                           BtnKeypad(10).Enabled = True
                        End If
                        If Right(lblNum.Caption, 1) = "%" Then
                           BtnKeypad(11).Enabled = True
                        End If
                        lblNum.Caption = left(lblNum.Caption, Len(Trim(lblNum.Caption)) - 1)
                    End If
                    Exit_Keypress_Flag = True
                    
                Case vbKeyEscape
                
                    Exit_Keypress_Flag = True
                     
                    If FrameCustInfo.Visible = True Then
                        FrameCustInfo.Visible = False
                        Exit Sub
                    ElseIf (lstDifference.Visible = False) Then
                        If cmbTable.ListIndex > 0 Or cmbGarson.ListIndex > 0 Then
                            Cancel
                            Exit Sub
                        ElseIf MyFormAddEditMode = AddMode And MaxRowFlexGrid <= 1 And CCur(lblSumPrice.Tag) = 0 And Val(lblCustomer.Tag) = -1 Then
                           If clsStation.KeyboardType = EnumKeyBoardType.Rb2 Then
                               If clsArya.CustomerId <> 10 And clsArya.CustomerId <> 1000 Then       ' Naghsh_Jahan Do Not Use
                                    ExitForm
                                    Exit Sub
                               End If
                           End If
                        End If
                        
                        Cancel
                        HideLstBoxes KeyCode
                        Exit Sub
                    Else
                        HideLstBoxes KeyCode
                        Exit Sub
                    End If
                    Exit Sub
                    
                Case vbKeyReturn
                       
                    Exit_Keypress_Flag = True
                   If lstDifference.Visible = False And Me.ActiveControl.Name <> cmbTable.Name And Me.ActiveControl.Name <> cmbGarson.Name And Me.ActiveControl.Name <> CmbPayk.Name And Me.ActiveControl.Name <> cmbServePlace.Name And Not (Me.ActiveControl.Name = FlxDetail.Name And FlxDetail.Col = 10) Then
                        If Not MyFormAddEditMode = ViewMode Then   ' add & Edit
                            If ClsFormAccess.SaveWithoutPrint = True Then
                                BeforeUpdate
                            Else
                                ShowDisMessage "‘„« »Â «„ò«‰ À»  »œÊ‰ ç«Å œ” —”Ì ‰œ«—Ìœ", 2000
                            End If
                       End If
                    ElseIf lstDifference.Visible = True Then
                        lstDifference_KeyUp KeyCode, 0
                        
                       ' RefreshLables  'Set Lables
                        FlxDetail.ShowCell MaxRowFlexGrid, 1
                        FlxDetail.SetFocus
                        FlxDetail.Select MaxRowFlexGrid, 1
                        If clsInvoiceValue.ShowInvoiceMenu = True Then
                            frmShowInvoiceMenu.UpdateGridValue
                        End If
                        Exit Sub
''                    ElseIf Me.ActiveControl.Name = FlxDetail.Name And FlxDetail.Col = 10 Then
''
''                        If SaveDifferences > 0 Then ShowDisMessage "", 1000: FlxDetail.Select FlxDetail.Row, 1
                        
                    ElseIf Me.ActiveControl.Name = cmbTable.Name Or Me.ActiveControl.Name = cmbGarson.Name Or Me.ActiveControl.Name = CmbPayk.Name Or Me.ActiveControl.Name = cmbServePlace.Name Then
                        FlxDetail.SetFocus
                        
                        Exit Sub
                    End If
                
                End Select
                
        Case 1     'With Shift Key
           
            Select Case KeyCode
            
                
                Case vbKey4, vbKeyNumpad4
                    
                    If clsStation.KeyboardType <> EnumKeyBoardType.S1 Then   ' Good Difference
                        With FlxDetail
                            
                            If .Col <> 10 Then
                                .Col = 10
                            End If
                            If .TextMatrix(.Row, 1) = "" Then
                                .Row = MaxRowFlexGrid - 1
                            End If
                            .ShowCell .Row, .Col
                            
                        End With
                        FlxDetail_Click
                        Exit_Keypress_Flag = True
                     End If
                
                Case vbKey5, vbKeyNumpad5    '%
                    
                    If BtnKeypad(11).Enabled = True Then
                        BtnKeypad_Click (11)
                    End If
                    'Exit_Keypress_Flag = True
                    
                Case 222        'Shift + ' (")
                    If ClsFormAccess.frmTempFich = True Then
                        If clsStation.KeyboardType <> EnumKeyBoardType.S1 Then
                           cmdTempFich_Click
                           Exit_Keypress_Flag = True
                        End If
                    End If
                Case 112, 113, 114, 115, 116, 117, 118, 119, 120, 121, 122, 123 'Shift + F1 ~ Shift + F12
                       MvarUserDefine = True
                       Form_KeyPress 112
            End Select
   
        Case 2
        
            Select Case KeyCode
            
               Case 17
                    If textDescription = False Then
                        txtDescription.SetFocus
                    Else
                        FlxDetail.SetFocus
                    End If
                    Exit_Keypress_Flag = True
               Case vbKeyO
                    If clsStation.KeyboardType <> EnumKeyBoardType.S1 Then
                        KeyCode = 0
                        Shift = 0
                        FWBtnGarsoon_Click
                        Exit_Keypress_Flag = True
                      End If
                
                Case vbKeyZ
                    If clsStation.KeyboardType <> EnumKeyBoardType.S1 And (Val(strCategory) >= 0 And Val(strCategory) <= 7) Then
                        KeyCode = 0
                        Shift = 0
                       FWBtnTable_Click
                       cmbTable.ListIndex = 0
                       Exit_Keypress_Flag = True
                   End If
                
                Case vbKeyF3       'Discount
                    
                    LblDiscount_Click
                
                    Exit_Keypress_Flag = True
                Case vbKeyF4      'Caree Fee
                    
                    If DropDownFlag = False Then
                       lblCarryFee_Click
                       Exit_Keypress_Flag = True
                    End If
                Case vbKeyF5       'Service
                    
                    If (Val(strCategory) >= 0 And Val(strCategory) <= 7) Then
                       lblService_Click
                    End If
                    Exit_Keypress_Flag = True
                    
                Case vbKeyF6       'Packing
                    
                    lblPacking_Click
                
                    Exit_Keypress_Flag = True
                Case vbKeyF7       ' Good Find
                    
                    BtnFindGood_Click
                    
                    Exit_Keypress_Flag = True
                Case vbKeyF8
                
                    Call OpenCashDrawer
                    
                    Exit_Keypress_Flag = True
                Case vbKeyF9
                    If clsStation.KeyboardType = EnumKeyBoardType.S1 Then
                        With FlxDetail
                            
                            If .Col <> 3 Then
                                .Col = 3
                            End If
                            If .TextMatrix(.Row, 1) = "" Then
                                .Row = MaxRowFlexGrid - 1
                            End If
                            .ShowCell .Row, .Col
                            
                        End With
                        FlxDetail_Click
                        
                    End If
                    Exit_Keypress_Flag = True
                       
                Case vbKeyF10
                
                    FWScrolltextPay_DblClick
                    Exit_Keypress_Flag = True
                
                Case vbKeyF11  'Customer Barcode
                    GetCustBarCode
                    Exit_Keypress_Flag = True
                Case vbKeyF12
                
                Case 221   ' Control + ]
                    If clsStation.KeyboardType <> EnumKeyBoardType.S1 Then
                        Shift = 0
                        KeyCode = 0
                        DropDownFlag = True
                        cmbServePlace.SetFocus
                        SendKeys "{F4}", True
                        DropDownFlag = False
                       '''''lblServePlace_Click
                       Exit_Keypress_Flag = True
                   End If
                 Case 220   ' Control + \
                    If textDescription = False Then
                        txtDescription.SetFocus
                    Else
                        FlxDetail.SetFocus
                    End If
                       
                    Exit_Keypress_Flag = True
           
            End Select
    End Select
    If IsUserDefinedKey(KeyCode, Shift) = True Then
       MvarUserDefine = True
    Else
       MvarUserDefine = False
    End If
  
End Sub

Private Function SaveDifferences() As Long
    SaveDifferences = 0
    On Error GoTo ErrHandler
    
    If Trim(FlxDetail.TextMatrix(FlxDetail.Row, IndexColDifferences)) = "" Then Exit Function
    Dim Rst As New ADODB.Recordset
    With FlxDetail
        ReDim Parameter(4) As Parameter
        Parameter(0) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
        Parameter(1) = GenerateInputParameter("@Defference", adVarWChar, 200, Trim(.TextMatrix(.Row, IndexColDifferences)))
        Parameter(2) = GenerateInputParameter("@NegativeDefference", adVarWChar, 200, "")
        Parameter(3) = GenerateInputParameter("@CostDifference", adInteger, 4, 0)
        Parameter(4) = GenerateOutputParameter("@LastCode", adInteger, 4)
        Dim Result As Long
        Result = RunParametricStoredProcedure("Insert_Differences", Parameter)
        
        ReDim Parameter(3) As Parameter
        Parameter(0) = GenerateInputParameter("@Level1", adInteger, 4, left(.TextMatrix(.Row, IndexColGoodCode), 2))
        Parameter(1) = GenerateInputParameter("@strSelectedLevels", adVarWChar, 4000, left(.TextMatrix(.Row, 5), 4))
        Parameter(2) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
        Parameter(3) = GenerateInputParameter("@ProductCompany", adInteger, 4, -1)
        Set Rst = RunParametricStoredProcedure2Rec("Get_Good_In_Levels", Parameter)
        Dim ExistCode As Boolean
        ExistCode = False
        Dim GoodStr As String
        GoodStr = ""
        If Not (Rst.BOF Or Rst.EOF) Then
            Do While Rst.EOF <> True
                If Rst!Code = .TextMatrix(.Row, 5) Then ExistCode = True
                GoodStr = GoodStr & "," & Rst!Code
                Rst.MoveNext
            Loop
        End If
        If ExistCode = True Then
            ReDim Parameter(1) As Parameter
            Parameter(0) = GenerateInputParameter("@GoodCode", adBSTR, 4000, GoodStr)
            Parameter(1) = GenerateInputParameter("@DifferenceCode", adInteger, 4, Result)
            RunParametricStoredProcedure "Insert_Goods_Difference", Parameter
        Else
            ReDim Parameter(1) As Parameter
            Parameter(0) = GenerateInputParameter("@GoodCode", adBSTR, 4000, .TextMatrix(.Row, IndexColGoodCode))
            Parameter(1) = GenerateInputParameter("@DifferenceCode", adInteger, 4, Result)
            RunParametricStoredProcedure "Insert_Goods_Difference", Parameter
        End If
    End With
    SaveDifferences = Result
    Set Rst = Nothing
Exit Function
ErrHandler:
    ShowDisMessage err.Description, 2000
End Function

Private Sub Form_KeyPress(KeyAscii As Integer)
        
    If Exit_Keypress_Flag = True Then Exit Sub
    If textDescription = True Then Exit Sub
    If AddressFlag = True Then Exit Sub
    If CustDescriptionFlag = True Then Exit Sub
    If textTempAddressFlag = True Then Exit Sub
     
    Dim j As Double
    Dim mvarstr As String
    If mvarbarcode Then
        Exit Sub
    End If
     
    If lblNum.Caption = "" And clsStation.StartCharacter <> "" Then
        If KeyAscii = Asc(clsStation.StartCharacter) Then
            GetCustBarCode
            Exit Sub
        End If
    End If
     
     If Me.ActiveControl.Name = cmbTable.Name Or Me.ActiveControl.Name = cmbGarson.Name Or Me.ActiveControl.Name = CmbPayk.Name Or (Me.ActiveControl.Name = FlxDetail.Name And FlxDetail.Col = 10) Then
         Exit Sub
     End If
     
    If KeyAscii = 27 Then        'Esc Key
''''    If (KeyAscii > 26 And KeyAscii <= 34) Then       'Control Key
''''    ElseIf KeyAscii = 36 And clsStation.KeyboardType = EnumKeyBoardType.Rb2 Then         'Control Key & .& -
''''    ElseIf KeyAscii = 37 Or KeyAscii = 144 Then         'Control Key & .& -
''''    ElseIf KeyAscii >= 47 And KeyAscii < 58 Then     'Numeric Key (Keycode47=/ for Barcode)
''''    ElseIf Val(GetKbLayout) = Val(LANG_EN_US) And KeyAscii = 39 Then
''''
''''    ElseIf KeyAscii = 13 And mvarShiftKey = 0 Then   ' Enter Key
''''    ElseIf KeyAscii = 8 Then
''''
''''    ElseIf MvarUserDefine = True Then
    ElseIf MvarUserDefine = True Then
        FlxDetail.Row = MaxRowFlexGrid
        
        ReDim Parameter(2) As Parameter
        
        Parameter(0) = GenerateInputParameter("@StationId", adInteger, 4, clsArya.StationNo)
        Parameter(1) = GenerateInputParameter("@KeyCode", adInteger, 4, mvarKeyCode)
        Parameter(2) = GenerateInputParameter("@ShiftKey", adInteger, 4, MvarShiftKey)
        
        Set rctmp = RunParametricStoredProcedure2Rec("Get_KB_Good", Parameter)
        
        mvarstr = ""
        Do While Not (rctmp.EOF)
            mvarstr = mvarstr & rctmp.Fields("GoodCode")
            rctmp.MoveNext
        Loop
        rctmp.Close
        If mvarstr <> "" Then
            Me.KeyPress KeyAscii
        Else
            frmDisMsg.Timer1.Interval = 500
            frmDisMsg.lblMessage = "ﬂ·Ìœ ›Êﬁ  ⁄—Ì› ‰‘œÂ «” "
            frmDisMsg.Timer1.Enabled = True
            frmDisMsg.Show vbModal
            MvarUserDefine = False
            Exit Sub
        End If
        
''
  Else
     Dim temp As Integer
     temp = KeyAscii
     If KeyAscii >= 127 Or Val(GetKbLayout) = Val(LANG_Pr_IR) Then
         KeyAscii = ClsCnvKeyBoard.CnvKeyBoard(KeyAscii)
     End If
        
    KeyAscii = temp
    ReDim Parameter(0) As Parameter
    Parameter(0) = GenerateInputParameter("@KeyAscii", adInteger, 4, KeyAscii)
    Set rctmp = RunParametricStoredProcedure2Rec("Get_DefaultKB_Good", Parameter)

    mvarstr = ""
    Do While Not (rctmp.EOF)
        mvarstr = mvarstr & rctmp.Fields("Code")
        rctmp.MoveNext
    Loop
    rctmp.Close
    If mvarstr <> "" Then
        Me.KeyPress KeyAscii
    Else
        frmDisMsg.lblMessage = "ﬂ·Ìœ ›Êﬁ  ⁄—Ì› ‰‘œÂ «” "
        frmDisMsg.Timer1.Enabled = True
        frmDisMsg.Show vbModal
    End If
 End If
MvarUserDefine = False

End Sub

Private Sub Form_Load()
    
On Error GoTo ErrHandler
    If ClsFormAccess.frmInvoice = False Then
        Unload Me
        Exit Sub
    End If

''''  BeforeCustomerSellPrice = clsStation.PriceType
    If clsStation.NumberOfId = 0 Then
        clsStation.NumberOfId = 8
    End If
    If clsStation.CityCode = "" Then
        clsStation.CityCode = "21"
    End If
   ' SetParent Me.Hwnd, mdifrm.Hwnd
'    If clsStation.CallerId8Port = True Then
'        FrameCallerId8Port.Visible = True
'    Else
'        FrameCallerId8Port.Visible = False
'    End If
    
    If clsStation.InvoiceRows = 0 Then
        MaxInvoiceRows = 8
    Else
        MaxInvoiceRows = Val(clsStation.InvoiceRows) + 1
    End If
   
    If clsStation.NoCurrentDay = True Then
        txtDate.Locked = False
    Else
        txtDate.Locked = True
    End If
    VarActForm = Me.Name
    
'    mvarStatus = Invoice
    
    Me.frameMenu.Visible = True
    
    FlexGridActive
    
'    If clsStation.FichStatusBar = True Then
'        Me.StatusBar.Visible = True
'    Else
'        Me.StatusBar.Visible = False
'    End If
    
''''    fwScrollTextCust.BackColor = Me.BackColor
    
    mvarServePlace = clsStation.ServePlaceDefault
    mVarOrderType = inPerson
    
'    ReDim arrBarcode(0)
    
    FillsGarsonCombo
    
    FillsTableCombo
    
    FillsPaykCombo
    
    Call ColorSetting
    
    ChangeLanguage
    
    MenuBarDefine
    
    PortClose
    
    GetProperController
    
    BlnFormLoaded = True
    
    clsStation.BasculeOn = False
   
    If clsStation.ShiftRate = True Then
        LblRate.Visible = True
    '    LblRate.Enabled = False
        clsStation.PriceType = mvarShiftNo
    Else
        If clsStation.MultiPrice = True Then
            LblRate.Visible = True
            If clsStation.PriceType < 1 Or clsStation.PriceType > 6 Then
                clsStation.PriceType = 1
                MsgBox " «‘ﬂ«· œ—  ⁄ÌÌ‰ ‰—Œ " & vbLf & "‰—Œ ÅÌ‘ ›—÷ «” ›«œÂ „Ì ‘Êœ"
            End If
        Else
            LblRate.Visible = False
            clsStation.PriceType = 1
        End If
    End If
    If clsStation.OutPrice = 0 Then clsStation.OutPrice = 1
    If clsArya.Accounting = True Then
        'FWlblAcc.Visible = True
    End If
   
   If clsArya.MojodiControl = True Then FWMojodiControl.Enabled = True
   If MojodiControlFlag = True Then
      FWMojodiControl.ButtonType = flwButtonOk
      FWMojodiControl.ForeColor = &H4000&  'vbGreen
   Else
      FWMojodiControl.ButtonType = flwButtonDelete
      FWMojodiControl.ForeColor = vbRed
   End If
    
    FWMojodiControl.Caption = "»«ﬁÌ„«‰œÂ ﬂ«·«"
        
    
    If strCategory = "07" Then
        FWBtnSplit.Visible = True
    End If
    ReDim ArrCostDifferences(0)
    
    ''''CenterTop Me
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

    If Val(GetSetting(strMainKey, "FrameBascule", "Left")) > 0 Then FrameBascule.left = Val(GetSetting(strMainKey, "FrameBascule", "Left"))
    If Val(GetSetting(strMainKey, "FrameBascule", "Top")) > 0 Then FrameBascule.top = Val(GetSetting(strMainKey, "FrameBascule", "Top"))

    
    If clsStation.PayFactorView = True Then
        cmdPayFactor.Visible = True
        lblPayFactorTotal.Visible = True
    Else
        cmdPayFactor.Visible = False
        lblPayFactorTotal.Visible = False
    End If
    If intVersion = Min Then
        cmbGarson.Enabled = False
        cmbTable.Enabled = False
        cmdTables.Enabled = False
        CmdPager.Enabled = False
        cmdPay.Enabled = False
        txtDescription.Enabled = False
        LblRate.Enabled = False
        cmdTempFich.Enabled = False
        ChkCallerId.Enabled = False
        Frame_CallerId.Enabled = False
        clsStation.TemporaryNo = False
    ElseIf intVersion = Normal Then
        CmdPager.Enabled = False
    ElseIf intVersion = Silver Then
        CmdPager.Enabled = True
        cmdTables.Enabled = False
    ElseIf intVersion = gold Or intVersion = Diamond Then
        CmdPager.Enabled = True
        cmdTables.Enabled = True
    End If
    
    If ClsFormAccess.frmTempFich = False Then cmdTempFich.Enabled = False
    ReDim Parameter(1) As Parameter
    Parameter(0) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
    Parameter(1) = GenerateInputParameter("@Branch", adInteger, 4, CurrentBranch)
    i = 0
    Set Rst = RunParametricStoredProcedure2Rec("Get_All_tPartitions", Parameter)
    If Not (Rst.BOF Or Rst.EOF) Then
        Do While Rst.EOF <> True
            i = i + 1
            Rst.MoveNext
        Loop
    End If
    If i > 1 And clsStation.OtherPartition = True And cmdTables.Enabled = True Then cmbTable.Enabled = False
    FrameCustInfo.Visible = False
    mvarStatus = Invoice

    Dim AutoHavale As Long
    ReDim Parameter(0) As Parameter
    Parameter(0) = GenerateOutputParameter("@AutoHavale", adInteger, 4)
    AutoHavale = RunParametricStoredProcedure("Get_AutoHavale", Parameter)
    If Val(AutoHavale) = 1 Then
        FWChkHavale.Visible = False
    End If

    If clsStation.Pager = True Then frmPager.Show '(intVersion = gold Or intVersion = Silver)
    
    If clsInvoiceValue.ShowInvoiceMenu = True And intVersion <> Min Then
        frmShowInvoiceMenu.Show vbModeless
    End If
    If clsInvoiceValue.ShowLogo = True And intVersion <> Min Then
        frmShowLogo.Show vbModeless
    End If
    If clsInvoiceValue.GoodMenuView = True And intVersion <> Min Then
        If filetemp.FileExists(App.Path & "\" & clsInvoiceValue.GoodMenuFileName & ".exe") Then
            Shell App.Path & "\" & clsInvoiceValue.GoodMenuFileName & ".exe", vbMinimizedFocus
          '  Sleep 200
        Else
            MsgBox " ›«Ì· ÅÌœ« ‰‘œ " & App.Path & "\" & clsInvoiceValue.GoodMenuFileName & ".exe"
        End If
    End If
    
    If clsStation.TemporaryNo = True Then FWLed1.Visible = False: FWLedTemp.Visible = True Else FWLed1.Visible = True: FWLedTemp.Visible = False

    If clsArya.LimitedVersion = True And HardLockFlagTrial = False Then LblLimited.Visible = True Else LblLimited.Visible = False

    If clsStation.DirectBascule Then FrameBascule.Visible = True Else FrameBascule.Visible = False
    
    If clsStation.Frame_Printers = True Then Frame_Printers.Visible = True: GetPrintersInDataBase
    If clsArya.MaxStationNo > 1 And clsStation.RefreshFichNo = True Then
        TimerNumber.Enabled = True
    Else
        TimerNumber.Enabled = False
    End If
    If clsStation.Frame_Printers = True And TimerNumber.Enabled = False Then Timer_Printers.Enabled = True
    
'################# CRM #############################
    If clsStation.LoyaltyCustomers = True Or clsStation.LoyaltyAllCustomers = True Then
        Set clsdiscount = New AryaCRMDiscountCalculator.clsdiscount
        Set InvInfo = New AryaCRMDiscountCalculator.InvoiceItem
        Set GoodItem = New AryaCRMDiscountCalculator.GoodView

        'clsdiscount.SetSMSDatabase "ServerName", "DbName", "DBLogin", "lemon7430"
        clsdiscount.SetDatabase clsArya.ServerName, clsArya.DbName, clsArya.DBLogin, SqlPass
        
        'set size of factor
        clsdiscount.InitFactorView 0.7 * Screen.Width \ Screen.TwipsPerPixelX, 0.7 * Screen.Height \ Screen.TwipsPerPixelY
        
        Dim Result As String
        ' init Rfid
        
        Result = clsdiscount.InitRFID(CurrentBranch, clsArya.StationNo)
        If Result <> "" Then ShowDisMessage Result, 1200
        ' stop start read card
        clsdiscount.ChangeCardReadStatus (True)
    End If
'###################################################
'''StimulPrinting
     
     If clsArya.NewPrinting = True Then Set StimulPrn = New AryaPrinting.StimulPrint
     
'#################  ‘ŒÌ’ ÂÊÌ  ”«“„«‰Ì #############################
     If clsStation.PersonIdCheck = True Then
          
          CmdPager.Caption = "»«—ê–«—Ì ·Ì” "
          PersonIdFolder    '   make folder for each day
          Set clsfinger = New CheckFingerPrintDll.CFP
          
          ' clsfinger.SetDevice clsStation.DeviceIP, clsStation.DeviceID
          
          'set database
          clsfinger.SetDatabase clsArya.ServerName, clsArya.DbName, clsArya.DBLogin, SqlPass
          
          'set form
          ' clsfinger.StartFingerPrintView clsStation.PersonIdRefreshTime, SubFolder & "\", "”Ì” „ „ò«‰Ì“Â  ‘ŒÌ’ ÂÊÌ  ¬—Ì«", clsStation.ListFont
          ' clsfinger.StartReadData
          If clsStation.Device2Id > 0 Then
               Set clsfinger2 = New CheckFingerPrintDll.CFP
               ' clsfinger2.SetDevice clsStation.Device2IP, clsStation.Device2Id
               
               'set database
               clsfinger2.SetDatabase clsArya.ServerName, clsArya.DbName, clsArya.DBLogin, SqlPass
               
               'set form
               ' clsfinger2.StartFingerPrintView clsStation.PersonIdRefreshTime, SubFolder & "\", "”Ì” „ „ò«‰Ì“Â  ‘ŒÌ’ ÂÊÌ  ¬—Ì«", clsStation.ListFont
               ' clsfinger2.StartReadData
          End If

     End If
'############################
 Exit Sub
ErrHandler:
     ShowDisMessage err.Description, 1500

End Sub
Private Sub PersonIdFolder()
     
     Dim filetemp As New FileSystemObject
     
     If Not filetemp.FolderExists(App.Path & "\FingerPrint") Then
         filetemp.CreateFolder App.Path & "\FingerPrint"
     End If
     SubFolder = App.Path & "\FingerPrint\" & DateToNumber8(Right(clsDate.shamsi(Date), 8))
     If Not filetemp.FolderExists(SubFolder) Then
         filetemp.CreateFolder SubFolder
     End If
    
End Sub
Private Sub lblTax_Click()
If lblNum.Caption = "" Then Exit Sub

If MyFormAddEditMode <> ViewMode Then

    Dim Str As String
    Dim Value As Double
    
    If Right(lblNum.Caption, 1) = "%" Then
        ShowMessage "  œ—’œÌ ﬁ«»· ﬁ»Ê· ‰Ì”  . ›ﬁÿ ⁄œœÌ Ê«—œ ò‰Ìœ " & LblTax.Caption, True, False, " «ÌÌœ", ""
        Exit Sub
    Else
        Str = lblNum.Caption
    End If
    
    Value = Val(Str)
    
     If Value < 0 Then
         ShowMessage "  „‰›Ì ﬁ«»· ﬁ»Ê· ‰Ì”  " & LblTax.Caption, True, False, " «ÌÌœ", ""
         lblNum.Caption = 0
         Exit Sub
     End If
    
'        If ClsFormAccess.CarryFee <> True Then
'            ShowMessage " ‘„« «Ã«“Â ê—› ‰  " & AddedDefaultTotal & " —« ‰œ«—Ìœ ", True, False, " «ÌÌœ", ""
'            Exit Sub
'        End If
        
     lblTaxTotal.Caption = Value
     lblTaxTotal = Val(Format(lblTaxTotal, "##"))
    
    Call RefreshLables
    BtnKeypad(10).Enabled = True
    BtnKalaDelete.Enabled = True
    BtnKeypad(11).Enabled = True
    'BtnKalaDelete.ForeColor = &H404080
End If
    lblNum.Caption = ""
    FlxDetail.SetFocus

End Sub

Private Sub TimerAlmP6_Timer()
    TimerAlmP6.Enabled = False
    LastRecordshow = True
    
    Dim ActiveFormModal As Boolean
    Dim varForm As Form
    Dim frmact As Form
    ActiveFormModal = False
    For Each varForm In Forms
        If varForm.Name = "frmFindCust" Or varForm.Name = "frmFindGoods" Or varForm.Name = "frmFindGoods_Menu" Or varForm.Name = "frmFindGoods_Kb" Then   'frmCallerIdView
            ActiveFormModal = True
            Exit For
        End If
    Next
    If ActiveFormModal = False And MaxRowFlexGrid = 1 And MyFormAddEditMode = AddMode And lblCustomer.Tag = "-1" Then
         
        If mdifrm.WindowState = 1 Then mdifrm.WindowState = 2    ' minimize to Maximizt
        
        frmCallerIdView.Show vbModal
    End If
'    lR = SetTopMostWindow(frmCallerIdView.hwnd, True)
End Sub

Private Sub txtDate_KeyDown(KeyCode As Integer, Shift As Integer)
    If txtDate.Locked = True Then Exit Sub
    If KeyCode = vbKeyDelete Then
        If Mid(txtDate.Text, txtDate.SelStart + 1, 1) = "/" Then
            KeyCode = 0
        End If
    End If
If Shift = 0 And ((KeyCode >= 48 And KeyCode <= 57) Or (KeyCode >= 96 And KeyCode <= 105)) Then
        If Len(Trim(lblNum.Caption)) >= 1 Then
            lblNum.Caption = left(lblNum.Caption, Len(Trim(lblNum.Caption)) - 1)
        End If
    End If

End Sub

Private Sub txtDate_KeyPress(KeyAscii As Integer)
'  FlxDetail.SetFocus
'On Error Resume Next
If Len(txtDate.Text) >= 8 And (KeyAscii >= 48 And KeyAscii <= 57) Then
    KeyAscii = 0
    Exit Sub
End If
If txtDate.SelStart = 0 Then Exit Sub
If KeyAscii = 8 Then
    If Len(txtDate.Text) = txtDate.SelStart Then
        Exit Sub
    End If
    If Mid(txtDate.Text, txtDate.SelStart, 1) = "/" Then
        KeyAscii = 0
        Exit Sub
    End If
    Exit Sub
End If
If KeyAscii < 48 Or KeyAscii > 57 Then
    KeyAscii = 0
    Exit Sub
End If
If Len(txtDate.Text) <> txtDate.SelStart Then
    Exit Sub
End If
Select Case Len(txtDate.Text)
Case 2
    txtDate.Text = txtDate.Text & "/"
    txtDate.SelStart = Len(txtDate.Text) + 1
Case 5
    txtDate.Text = txtDate.Text & "/"
    txtDate.SelStart = Len(txtDate.Text) + 1
End Select
End Sub

Private Sub FillsGarsonCombo()
    If rctmp.State = 1 Then rctmp.Close
    cmbGarson.Clear
    ReDim Parameter(0) As Parameter
    Parameter(0) = GenerateInputParameter("@Branch", adInteger, 4, CurrentBranch)
    Set rctmp = RunParametricStoredProcedure2Rec("Get_vw_Garson", Parameter)
    cmbGarson.AddItem ""
    cmbGarson.ItemData(0) = 0
    If Not (rctmp.EOF = True And rctmp.BOF = True) Then
        Do While rctmp.EOF <> True
        
            cmbGarson.AddItem CStr(rctmp.Fields("nvcFirstName")) & " " & CStr(rctmp.Fields("nvcSurName"))
            cmbGarson.ItemData(cmbGarson.ListCount - 1) = Val(rctmp.Fields("pPNo"))
            rctmp.MoveNext
            
        Loop
         
    End If
    cmbGarson.ListIndex = 0
    rctmp.Close

End Sub
Private Sub FillsPaykCombo()
    If rctmp.State = 1 Then rctmp.Close
    CmbPayk.Clear
    
    ReDim Parameter(1) As Parameter
    Parameter(0) = GenerateInputParameter("@Job", adInteger, 4, EnumIncharge.Payk)
    Parameter(1) = GenerateInputParameter("@Branch", adInteger, 4, CurrentBranch)
    Set rctmp = RunParametricStoredProcedure2Rec("Get_vw_Per_BY_Job", Parameter)
    CmbPayk.AddItem ""
    CmbPayk.ItemData(0) = 0
    If Not (rctmp.EOF = True And rctmp.BOF = True) Then
        Do While rctmp.EOF <> True
        
            CmbPayk.AddItem CStr(rctmp.Fields("nvcFirstName")) & " " & CStr(rctmp.Fields("nvcSurName"))
            CmbPayk.ItemData(CmbPayk.ListCount - 1) = Val(rctmp.Fields("pPNo"))
            rctmp.MoveNext
            
        Loop
         
    End If
    CmbPayk.ListIndex = -1
    rctmp.Close

End Sub

Private Sub FillsTableCombo()
    Dim L_Rst As New ADODB.Recordset
    
'    If L_Rst.State = adStateOpen Then L_Rst.Close
    cmbTable.Clear
    
    ReDim Parameter(1) As Parameter
    Parameter(0) = GenerateInputParameter("@Branch", adInteger, 4, CurrentBranch)
    Parameter(1) = GenerateInputParameter("@TableControl", adBoolean, 1, IIf(clsStation.TableControl = True, 1, 0))
'    Set L_Rst = RunParametricStoredProcedure2Rec("RetriveTable", Parameter)
    Set L_Rst = RunParametricStoredProcedure2Rec("RetriveTable_Branch", Parameter)
    
    cmbTable.AddItem ""
    cmbTable.ItemData(0) = 0
    
    If Not (L_Rst.EOF = True And L_Rst.BOF = True) Then
        '''L_Rst.moveFirst
        Do While L_Rst.EOF <> True
            cmbTable.AddItem L_Rst!TableDescription
            '''cmbTable.AddItem L_Rst.Fields("Name")
            cmbTable.ItemData(cmbTable.NewIndex) = Val(L_Rst!No)
            L_Rst.MoveNext
        Loop
    Else
        ReDim Parameter(0) As Parameter
        Parameter(0) = GenerateInputParameter("@Branch", adInteger, 4, CurrentBranch)
        Set L_Rst = RunParametricStoredProcedure2Rec("Get_Tables", Parameter)
'        If Not (L_Rst.EOF = True And L_Rst.BOF = True) Then ' If Table Exist Then
'            ShowDisMessage " ﬂ·ÌÂ „Ì“Â« Å— Â” ‰œ", 2000
'        End If
    End If
    
    If cmbTable.Text = "" Then cmbTable.ListIndex = 0
    L_Rst.Close: Set L_Rst = Nothing
    
    If cmbTableData <> 0 Then  'And MyFormAddEditMode <> ViewMode
        Dim ExistTableNo As Boolean
        ExistTableNo = False
        For i = 0 To cmbTable.ListCount - 1
            If cmbTable.ItemData(i) = cmbTableData Then
                ExistTableNo = True
                cmbTable.ListIndex = i
                Exit For
            End If
        Next i
        If ExistTableNo = False Then
            cmbTable.AddItem cmbTableName
            cmbTable.ItemData(cmbTable.ListCount - 1) = cmbTableData
            cmbTable.ListIndex = cmbTable.ListCount - 1
''''            For i = 0 To cmbTable.ListCount - 1
''''                If cmbTable.ItemData(i) = cmbTableData Then
''''                    cmbTable.ListIndex = i
''''                    Exit For
''''                End If
''''            Next i
        End If
    End If

End Sub
Private Sub FillsFullTableCombo()
    Dim L_Rst As New ADODB.Recordset
    
    cmbTable.Clear
    
    ReDim Parameter(1) As Parameter
    Parameter(0) = GenerateInputParameter("@Branch", adInteger, 4, CurrentBranch)
    Parameter(1) = GenerateInputParameter("@TableControl", adBoolean, 1, 0)
'    Set L_Rst = RunParametricStoredProcedure2Rec("RetriveTable", Parameter)
    Set L_Rst = RunParametricStoredProcedure2Rec("RetriveTable_Branch", Parameter)
    
    cmbTable.AddItem ""
    cmbTable.ItemData(0) = 0
    
    If Not (L_Rst.EOF = True And L_Rst.BOF = True) Then
        Do While L_Rst.EOF <> True
            cmbTable.AddItem L_Rst!TableDescription
            cmbTable.ItemData(cmbTable.NewIndex) = L_Rst!No
            
            L_Rst.MoveNext
        Loop
    End If
    
    cmbTable.ListIndex = 0
    L_Rst.Close: Set L_Rst = Nothing
End Sub

Private Sub cmbTable_DropDown()
  FillsTableCombo
  ViewFlag = False
End Sub

Private Sub PortClose()
    
    For i = mscSerial.LBound To mscSerial.UBound
        If Me.mscSerial(i).PortOpen Then
            Me.mscSerial(i).PortOpen = False
        End If
    Next i
    
    USBCallerID1.CloseDevice ' Close and stop the device
    
End Sub


'Private Sub Form_LostFocus()
'MsgBox "safsdf"
'End Sub

Private Sub Form_Unload(Cancel As Integer)

    modgl.mvarDeleteMsg = ""
    ClearDataFlexGrid
    
    BlnFormLoaded = False
    
    VarActForm = ""
    
    Set mdifrm.FileCls = Nothing
    Set clsDate = Nothing
    Set ClsCnvKeyBoard = Nothing
    
    Set rctmp = Nothing
    If PosConnection.State = adStateOpen Then PosConnection.Close: Set PosConnection = Nothing
    Set RstTemp = Nothing
    
    PortClose
    Set MahakScaleOCX3 = Nothing
    TimerScale.Enabled = False
    
    If clsInvoiceValue.GoodMenuView = True Then
        CloseWindow "„‰ÊœÌÃÌ «· " & Trim(clsArya.Company)
    End If
    If clsInvoiceValue.ShowLogo = True Then
        Unload frmShowLogo
    End If
    
    SaveSetting strMainKey, Me.Name, "Left", Me.left
    SaveSetting strMainKey, Me.Name, "Top", Me.top
    
'################# CRM #############################
    If clsStation.LoyaltyCustomers = True Or clsStation.LoyaltyAllCustomers = True Then
        clsdiscount.ChangeCardReadStatus (False)
        Set clsdiscount = Nothing
        Set InvInfo = Nothing
        Set GoodItem = Nothing
    End If


'#################  ‘ŒÌ’ ÂÊÌ  #############################
    If clsStation.PersonIdCheck = True Then
'        clsdiscount.ChangeCardReadStatus (False)
        ' clsfinger.StopReadData
        Set clsfinger = Nothing
        If clsStation.Device2Id > 0 Then
          ' clsfinger2.StopReadData
          Set clsfinger2 = Nothing
        End If
    End If
'###################################################
    If RfidReaderIsActive = True Then MF_ExitComm

'    If clsInvoiceValue.ShowInvoiceMenu = True Then
'        Unload frmShowInvoiceMenu
'    End If
''    Dim nodX
''    For Each nodX In frmGroupMenu.trMenu.Nodes
''        nodX.Expanded = False
''       ' nodX.EnsureVisible
''    Next nodX

End Sub


Public Sub FirstKey()

    If ClsFormAccess.NavigateFactor = False Then
        ShowDisMessage "‘„« »Â «Ì‰ «„ò«‰ œ” —”Ì ‰œ«—Ìœ", 1500
        Exit Sub
    End If
    MyFormAddEditMode = ViewMode
    DefaultValueLables

    ArrowkeyStatusbar FirstRecord
    If StatusBar.Panels(2).Tag <> "" Then
        Me.txtNo.Text = StatusBar.Panels(2).Tag
    Else
        For i = 2 To 5
            If StatusBar.Panels(i).Tag <> "" Then
                Me.txtNo.Text = StatusBar.Panels(i).Tag
                Exit For
            End If
        Next i
    End If
    GetDataDetail
    RefreshLables
    SetFirstToolBar

End Sub


Public Sub PreviousKey()
    
    If ClsFormAccess.NavigateFactor = False Then
        ShowDisMessage "‘„« »Â «Ì‰ «„ò«‰ œ” —”Ì ‰œ«—Ìœ", 1500
        Exit Sub
    End If
    If MyFormAddEditMode = AddMode Then
        LastKey
        Exit Sub
    End If
    MyFormAddEditMode = ViewMode  'View Mode
    
    ArrowkeyStatusbar PreviousRecord, Val(Me.txtNo.Text)
    
    Dim j As Integer
    If StatusBar.Panels(6).Tag = "" Then
        For i = 5 To 3
            If StatusBar.Panels(i).Tag <> "" Then
                j = i
                Exit For
            End If
        Next i
    Else
        j = 6
    End If
    If j = 0 Then Exit Sub
    If Val(Me.txtNo.Text) <= Val(StatusBar.Panels(j).Tag) Then
        For i = j To 3 Step -1
            If Me.txtNo.Text = StatusBar.Panels(i).Tag Then
                If StatusBar.Panels(i - 1).Tag <> "" Then
                    Me.txtNo.Text = StatusBar.Panels(i - 1).Tag
                End If
                Exit For
            End If
        Next i
    Else
        Me.txtNo.Text = StatusBar.Panels(6).Tag
    End If
    
    GetDataDetail
    RefreshLables
    SetFirstToolBar
    
End Sub


Public Sub NextKey()

    If ClsFormAccess.NavigateFactor = False Then
        ShowDisMessage "‘„« »Â «Ì‰ «„ò«‰ œ” —”Ì ‰œ«—Ìœ", 1500
        Exit Sub
    End If
    If MyFormAddEditMode = AddMode Then
        LastKey
        Exit Sub
    End If
    MyFormAddEditMode = ViewMode  'View Mode
    
    ArrowkeyStatusbar EnumDirection.NextRecord, Val(Me.txtNo.Text)
    
    Dim j As Integer
    If StatusBar.Panels(2).Tag = "" Then
        For i = 3 To 6
            If StatusBar.Panels(i).Tag <> "" Then
                j = i
                Exit For
            End If
        Next i
    Else
        j = 2
    End If
    If j = 0 Then Exit Sub
    If Val(Me.txtNo.Text) >= Val(StatusBar.Panels(j).Tag) Then
        For i = j To 5
        
            If Me.txtNo.Text = StatusBar.Panels(i).Tag Then
                If StatusBar.Panels(i + 1).Tag <> "" Then
                    Me.txtNo.Text = StatusBar.Panels(i + 1).Tag
                End If
                Exit For
            End If
        Next i
    Else
        Me.txtNo.Text = StatusBar.Panels(j).Tag
    End If
    
    GetDataDetail
    RefreshLables
    SetFirstToolBar
    
End Sub


Public Sub LastKey()

    If ClsFormAccess.NavigateFactor = False Then
        ShowDisMessage "‘„« »Â «Ì‰ «„ò«‰ œ” —”Ì ‰œ«—Ìœ", 1500
        Exit Sub
    End If
    MyFormAddEditMode = ViewMode  'View Mode
    
    DefaultValueLables
    ArrowkeyStatusbar LastRecord
    If StatusBar.Panels(6).Tag <> "" Then
        Me.txtNo.Text = StatusBar.Panels(6).Tag
    Else
        For i = 6 To 3
            If StatusBar.Panels(i).Tag <> "" Then
                Me.txtNo.Text = StatusBar.Panels(i).Tag
            End If
        Next i
    End If
    If StatusBar.Panels(6).Tag <> "" Then
       GetDataDetail
       RefreshLables
       SetFirstToolBar
    End If
End Sub

Public Sub Add()

    On Error Resume Next
    PayClick = False
    UpdateFromFinalCheck = False
    FrameCustInfo.Visible = False
    Dim AutoValue As Integer
    HideLstBoxes 27
    intTempFich = 0
    RoundDiscount = 0
    intSerialNo = 0
    If clsStation.Language = Farsi Then
        LblOrder.Caption = "Õ÷Ê—Ì"
    Else
        LblOrder.Caption = "Inside"
    End If
    boolPayment = False
    BalancePayment = False
    BtnKeypad(10).Enabled = True
    BtnKeypad(11).Enabled = True
    mvarTipAmount = 0
    ClearDataFlexGrid
    txtDate.Text = mvarDate  ' Right(clsDate.shamsi(Date), 8)
    Me.Number
    DefaultValueLables       'Set Default Value Lables
    Me.ValueLabel
    For i = 2 To 5
        Me.StatusBar.Panels(i).Bevel = sbrInset
    Next i
    
    ArrowkeyStatusbar LastRecord         'Display 5 Last Fich
    

    Me.txtRecursive = 0
    fwlblRecursive.Visible = False
    FWLblEdit.Visible = False

'    fwScrollTextCust.Visible = True
    lblNum = ""
    lblBarCode = ""
    mvarbarcode = False
    
    FillsTableCombo
    If cmbTable.ListCount > 0 Then cmbTable.ListIndex = 0
    cmbGarson.ListIndex = 0
    CmbPayk.ListIndex = 0
    
    mVarOrderType = inPerson
    
    mvarServePlace = clsStation.ServePlaceDefault
    If mvarServePlace = EnumServePlace.Table Or mvarServePlace = EnumServePlace.Salon Then
        ServiceRate = DefaultServicePercent
    Else
        ServiceRate = 0
    End If
    EnableDefaultServiceRate = True
    For i = 0 To cmbServePlace.ListCount - 1
        If mvarServePlace = cmbServePlace.ItemData(i) Then
            cmbServePlace.ListIndex = i
            Exit For
        End If
    Next i

    If clsStation.ServePlaceDefault = EnumServePlace.Delivery And MyFormAddEditMode = AddMode And BlnFormLoaded Then
       '' FindCust   ''  Â„Ì‘Â »«“ „Ì ‘œ Ê „‘ò· «ÌÃ«œ „Ì ò—œ
    End If
    
    MyFormAddEditMode = AddMode       'Add Mode
    SetFirstToolBar
    
    FlxDetail.ColHidden(8) = True
   ' FlxDetail.ColWidth(10) = FlxDetail.Width / 6   'ServePlace
    FlxDetail.ColWidth(10) = FlxDetail.Width / 4.5      'Diffrence
    FWScrolltextPay.Visible = False
    FWScrollSend.Visible = False
     
    intSumOfCurrentServePlaces = mvarServePlace

    Call CashCloseStatus
    mvarEditedFich = 0
    FlxDetail.Select 1, 1
    FWBtnSplit.Caption = "„Õ«”»Â"
'''     FWBtnSplit.Caption = "„⁄„Ê·Ì"
    FWBtnSplit.ForeColor = vbRed
    SplitFlag = False
    If clsStation.DeliveryNoView Then
       CalculateDelivery
       CalculateTemporary
    End If
    InventoryNo = 0
    
    LblSubTotal.Caption = 0

Dim mm As Integer
If AlmPort = 0 Then
    For mm = 0 To 7
        If FWModem(mm).BackColor = vbGreen Then
            FWModem(mm).BackColor = &H80000016  '&H808000
            FWModem(mm).ToolTipText = ""
        End If
    Next mm
End If
If clsStation.AutoCallerId = True Then
    If Val(ModemPriority(1)) <> 0 Then
        
        FWModem(Call_Priority - 1).BackColor = vbGreen   ' &H80000003&
        Call_RealNumber = Call_Number(Call_Priority)
        ModemPriority(1) = 0
        Call_Number(Call_Priority) = ""
        For mm = 2 To 8
            ModemPriority(mm - 1) = ModemPriority(mm)
        Next mm
        If Val(ModemPriority(1)) = 0 Then Call_Priority = 0 Else Call_Priority = Val(ModemPriority(1))
        
        FindCust
            
    
    Else
           
'        For i = 0 To FWModem.Count - 1
'            FWModem(i).BackColor = &H808000
'        Next i
    End If
End If
    If clsStation.Language = Farsi Then
        txtDescription.Text = "               ÅÌ€«„     "
    Else
        txtDescription.Text = "             Message     "
    End If
    ChanceBarcodeQuantity = 0
    Repeatbarcode = 0
    
    textDescriptionFlag = False
    textTempAddressFlag = False
    FlxDetail.SetFocus

    cmbTableName = ""
    cmbTableData = 0
    flgShowOrderDetail = True

    If mvarServePlace = Out And clsStation.FixRateChange = False Then
        clsStation.PriceType = clsStation.OutPrice
    ElseIf clsStation.FixRateChange = False Then
        clsStation.PriceType = MainPriceType
    End If
    
    If clsStation.ShiftRate = True Then
        clsStation.PriceType = mvarShiftNo
    End If
    
   If clsStation.FixRateChange = False Then ''And mvarRateFlag = True Then
        If clsStation.PriceType = 1 Then
            If clsStation.Language = Farsi Then
                LblRate.Caption = "‰—Œ «Ê·"
            Else
                LblRate.Caption = "Rate 1"
            End If
        ElseIf clsStation.PriceType = 2 Then
            If clsStation.Language = Farsi Then
                LblRate.Caption = "‰—Œ œÊ„"
            Else
               LblRate.Caption = "Rate 2"
            End If
        ElseIf clsStation.PriceType = 3 Then
            If clsStation.Language = Farsi Then
                LblRate.Caption = "‰—Œ ”Ê„"
            Else
                LblRate.Caption = "Rate 3"
            End If
        ElseIf clsStation.PriceType = 4 Then
            If clsStation.Language = Farsi Then
                LblRate.Caption = "‰—Œ çÂ«—„"
            Else
                LblRate.Caption = "Rate 4"
            End If
        ElseIf clsStation.PriceType = 5 Then
            If clsStation.Language = Farsi Then
                LblRate.Caption = "‰—Œ Å‰Ã„"
            Else
                LblRate.Caption = "Rate 5"
            End If
        ElseIf clsStation.PriceType = 6 Then
            If clsStation.Language = Farsi Then
                LblRate.Caption = "‰—Œ ‘‘„"
            Else
                LblRate.Caption = "Rate 6"
            End If
        End If
   Else
            If mvarStartRate = 1 Then
                If clsStation.Language = Farsi Then
                    LblRate.Caption = "‰—Œ «Ê·"
                Else
                    LblRate.Caption = "Rate 1"
                End If
            ElseIf mvarStartRate = 2 Then
                If clsStation.Language = Farsi Then
                    LblRate.Caption = "‰—Œ œÊ„"
                Else
                    LblRate.Caption = "Rate 2"
                End If
            ElseIf mvarStartRate = 3 Then
                If clsStation.Language = Farsi Then
                    LblRate.Caption = "‰—Œ ”Ê„"
                Else
                    LblRate.Caption = "Rate 3"
                End If
            ElseIf mvarStartRate = 4 Then
                If clsStation.Language = Farsi Then
                    LblRate.Caption = "‰—Œ çÂ«—„"
                Else
                    LblRate.Caption = "Rate 4"
                End If
            ElseIf mvarStartRate = 5 Then
                If clsStation.Language = Farsi Then
                    LblRate.Caption = "‰—Œ Å‰Ã„"
                Else
                    LblRate.Caption = "Rate 5"
                End If
            ElseIf mvarStartRate = 6 Then
               If clsStation.Language = Farsi Then
                    LblRate.Caption = "‰—Œ ‘‘„"
                Else
                    LblRate.Caption = "Rate 6"
                End If
            End If
            
        
     End If
   
    If TimeReaderPort > 0 Then
        TimerReader.Enabled = True
    End If
    
    TxtTempAddress.Text = "¬œ—” „Êﬁ  : "
    LblInvoicePrint.Caption = ""
    ViewFlag = False
    FWChkHavale.Value = False
    
    ReDim ArrCostDifferences(0)
    ReDim ArrDifferences(0)

    ChkCallerId.Visible = True

'    LblTip.Visible = False
    
    ChkIsLocked.Value = 0
    ChkIsLocked.Enabled = False
    
    LblRemain.Caption = "»«ﬁÌ„«‰œÂ"
    
    FWChkAccount.Value = False
    
    ServeChangeFlag = False

    If IsFarabin = True Then ShowMonitor 0

    If mvarStatus = Invoice And clsStation.RfidReader = True And HasRfidReader = True Then
        EnableRFID
    Else
        DisableRFID
    End If

''################  ‘ŒÌ’ ÂÊÌ  - »«—ê–«—Ì « Ê„« Ìò   #####################################
     If clsStation.PersonIdCheck = True And clsStation.ListAutoLoad = True Then
          GetFirstPersonFromList
     End If
'####################
End Sub
Private Sub GetFirstPersonFromList()
     On Error GoTo ErrorHandler
     If RstTemp.State <> 0 Then RstTemp.Close
     Set RstTemp = RunStoredProcedure2RecordSet("Arya_Kitchen_GetFirstPerson")
     If Not (RstTemp.EOF = True And RstTemp.BOF = True) Then
          mvarcode = RstTemp!Code
          PersonIdqueue = RstTemp!Pk_Id
          FindCustomerByCode mvarcode
     Else
          Timer_PersonIdCheck.Enabled = True
     End If
Exit Sub
ErrorHandler:
     ShowDisMessage err.Description, 1500
End Sub


Private Sub FindCustomerByCode(Code As Double)
    On Error GoTo ErrorHandler
     lblCustomer.Tag = mvarcode
     mvarcode = 0
     If clsStation.CustomerOrderDefault = True Then
        mVarOrderType = inPerson
     Else
        mVarOrderType = ByPhone
     End If
     mvarPublicOrderType = inPerson
     If mVarOrderType = ByPhone Then
        If clsStation.Language = Farsi Then
             LblOrder.Caption = " ·›‰Ì"
        Else
             LblOrder.Caption = "By phone"
        End If
     End If
     If clsStation.Language = Farsi Then
        LblOrder.Caption = "Õ÷Ê—Ì"
     Else
         LblOrder.Caption = "Inside"
     End If
     For i = 0 To cmbServePlace.ListCount - 1
         If mvarServePlace = cmbServePlace.ItemData(i) Then
             cmbServePlace.ListIndex = i
             Exit For
         End If
     Next i
     UpdatelblCustomer
     UpdatelblServePlace
     RefreshLables
Exit Sub
ErrorHandler:
     ShowDisMessage err.Description, 1500
End Sub
Private Sub clsfinger_DataRecieved(ByVal IDs As String)
'      MsgBox " :  œ—Ì«›  ‘œ " & " " & IDs
'
    On Error GoTo ErrorHandler
    PersonIdqueue = Val(IDs)
    Dim rctmp2 As New ADODB.Recordset
    ReDim Parameter(0) As Parameter
    Parameter(0) = GenerateInputParameter("@PK_ID", adInteger, 4, Val(IDs))
    Set rctmp2 = RunParametricStoredProcedure2Rec("Arya_Kitchen_GetPerson_ById", Parameter)
    If rctmp2.EOF <> True And rctmp2.BOF <> True Then
          mvarcode = rctmp2!Code
          FindCustomerByCode mvarcode
    Else
    End If
Exit Sub
ErrorHandler:
     ShowDisMessage err.Description, 1500

End Sub
Private Sub PersonIdChangeStatus()
    On Error GoTo ErrorHandler
    ReDim Parameter(1) As Parameter
    Parameter(0) = GenerateInputParameter("@PK_ID", adInteger, 4, PersonIdqueue)
    Parameter(1) = GenerateInputParameter("@StatusNo", adInteger, 4, 1)
    RunParametricStoredProcedure "Arya_Kitchen_ChangeStatus", Parameter

Exit Sub
ErrorHandler:
     ShowDisMessage err.Description, 1500

End Sub
Private Sub TimerRFID_Timer()
    Dim serial As String
    BufferTXT.Text = ""
    serial = ""
    '''###
    If mvarStatus <> Invoice Then Exit Sub
    If MyFormAddEditMode <> AddMode Then Exit Sub
    Dim Status As Integer
    Status = 1
    If MF_Request(0, 0, cardT(0)) = 0 Then
        RFIDStatus = MF_Anticoll(0, cardSN(0))
       
        serial = CStr(Hex(cardSN(0)) & Hex(cardSN(1)) & Hex(cardSN(2)) & Hex(cardSN(3)))
    
        Sleep 100
        If RFIDStatus = 0 Then
            For i = 0 To 5
                Ckey(i) = hex2dec(keyTXT(i))
            Next i
            If MF_Select(0, cardSN(0)) = 0 Then
                If MF_LoadKey(0, Ckey(0)) = 0 Then
                    If MF_Authentication(0, IIf(KeyAorB.Value, 1, 0), Val(blockNtxt), cardSN(0)) = 0 Then
                        For i = 0 To 64
                            Dbuffer(i) = 0
                        Next i
                        If MF_Read(0, blockNtxt, Bcount.Text, Dbuffer(0)) = 0 Then
                            For i = 0 To 64
                                BufferTXT = BufferTXT & Chr(Dbuffer(i))
                            Next i
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
    Else
        Status = 0
    End If
    
    If Status = 0 Then
        Exit Sub
    End If

'    Dim Position As Integer
'    Position = 0
'    Position = InStr(BufferTXT.Text, "+")
'    If Position = 0 Then ShowDisMessage "  ò«—  Œ«„ „Ì »«‘œ ", 1500: Exit Sub
'
'    Dim RFIDCart As String
'    RFIDCart = Mid(BufferTXT.Text, 1, Position - 1)

    
    ReDim Parameter(1) As Parameter
    Parameter(0) = GenerateInputParameter("@ActDeact", adInteger, 4, 0) ' Only Active
    Parameter(1) = GenerateInputParameter("@nvcRfid", adWChar, 20, serial)
    Set Rst = RunParametricStoredProcedure2Rec("Get_Customer_Rfid", Parameter)
    If Rst.EOF <> True Then
        mvarcode = Rst!Code
        mvarName = Rst![Name]
    Else
        If clsStation.RfidLongBuzzer = True Then i = MF_ControlBuzzer(0, 1.5)
        mvarcode = 0
        mvarName = ""
        frmDisMsg.lblMessage.Caption = " «Ì‰ ò«—  œ— œÌ «»Ì” „‘ —Ì«‰  ⁄—Ì› ‰‘œÂ "
        frmDisMsg.Timer1.Interval = 2000
        frmDisMsg.Timer1.Enabled = True
        frmDisMsg.Show vbModal
    End If

    If mvarcode = 0 Then Exit Sub
    
    i = MF_DeviceReset(0)

    If mvarcode <> 0 Then
        lblCustomer.Tag = mvarcode
        mvarcode = 0
    Else
        lblCustomer.Tag = -1
    End If
    mvarPublicOrderType = inPerson
    mVarOrderType = inPerson
    If clsStation.Language = Farsi Then
        LblOrder.Caption = "Õ÷Ê—Ì"
    Else
        LblOrder.Caption = "Inside"
    End If
    UpdatelblCustomer
    RefreshLables
    
    DisableRFID

End Sub
Private Sub EnableRFID()
    TimerRFID.Enabled = True
'    chkRFID.Value = 1
End Sub
Private Sub DisableRFID()
    TimerRFID.Enabled = False
'    chkRFID.Value = 0
End Sub

Private Sub CashCloseStatus()
    Dim rctmp2 As New ADODB.Recordset
    ReDim Parameter(1) As Parameter
    Parameter(0) = GenerateInputParameter("@Date", adVarChar, 50, txtDate.Text)
    Parameter(1) = GenerateInputParameter("@ShiftNo", adInteger, 4, -1)
    Set rctmp2 = RunParametricStoredProcedure2Rec("Get_tblAcc_CashClose", Parameter)
    If rctmp2.EOF <> True And rctmp2.BOF <> True Then
        If rctmp2!CashActive = 0 Then
            clsStation.CashClose = True
            FWlblCash.Visible = True
            FWlblCash.BackColor = vbRed
            FWlblCash.Caption = " ’‰œÊﬁ »” Â "
        Else
            clsStation.CashClose = False
            FWlblCash.BackColor = &H8000&
            FWlblCash.Caption = " ’‰œÊﬁ »«“ "
            FWlblCash.Visible = False
        End If
    Else
        clsStation.CashClose = False
        FWlblCash.BackColor = &H8000&
        FWlblCash.Caption = " ’‰œÊﬁ »«“ "
        FWlblCash.Visible = False
    End If

End Sub
Public Sub Cancel()
    If MyFormAddEditMode = AddMode Then
    
        If Val(lblCustomer.Tag) > 0 Then
            If MaxRowFlexGrid <> 1 Then
                ClearDataFlexGrid
            Else
                If clsStation.InvoiceStatusDefault = True Then
                    mvarStatus = EnumFactorType.Invoice
                    If clsStation.Language = Farsi Then
                        LblInvoice.Caption = "›«ﬂ Ê— ›—Ê‘"
                    Else
                        LblInvoice.Caption = "Invoice"
                    End If
                    If clsStation.PayFactorView = True Then
                        cmdPayFactor.Visible = True
                        lblPayFactorTotal.Visible = True
                    Else
                        cmdPayFactor.Visible = False
                        lblPayFactorTotal.Visible = False
                    End If
                End If
                Add
            End If
        Else
            If clsStation.InvoiceStatusDefault = True Then
                mvarStatus = EnumFactorType.Invoice
                If clsStation.Language = Farsi Then
                    LblInvoice.Caption = "›«ﬂ Ê— ›—Ê‘"
                Else
                    LblInvoice.Caption = "Invoice"
                End If
                If clsStation.PayFactorView = True Then
                    cmdPayFactor.Visible = True
                    lblPayFactorTotal.Visible = True
                Else
                    cmdPayFactor.Visible = False
                    lblPayFactorTotal.Visible = False
                End If
            End If
            Add
        End If
    Else
        MyFormAddEditMode = AddMode
        If clsStation.InvoiceStatusDefault = True Then
            mvarStatus = EnumFactorType.Invoice
            If clsStation.Language = Farsi Then
                LblInvoice.Caption = "›«ﬂ Ê— ›—Ê‘"
            Else
                LblInvoice.Caption = "Invoice"
            End If
            If clsStation.PayFactorView = True Then
                cmdPayFactor.Visible = True
                lblPayFactorTotal.Visible = True
            Else
                cmdPayFactor.Visible = False
                lblPayFactorTotal.Visible = False
            End If
        End If
        Add
    End If

End Sub

Public Sub Edit()
    
    If FWChkAccount.Value = True And Val(lblCustomer.Tag) < 1 Then ShowDisMessage "”‰œ Õ”«»œ«—Ì »—«Ì «Ì‰ ›«ò Ê— ﬁ»·« ’«œ— ‘œÂ Ê ﬁ«»· ÊÌ—«Ì‘ ‰Ì”  . «“›«ò Ê— ÃœÌœ Ê »—ê‘  «“ ›—Ê‘ «” ›«œÂ ò‰Ìœ ", 1500: Exit Sub
    If (MyFormAddEditMode = ViewMode And clsStation.EditCompatibleSamar1 = True) Then
            Find
            If Me.FindFlag = False Then Exit Sub
    ElseIf MyFormAddEditMode = AddMode And clsStation.EditCompatibleSamar1 = True And MaxRowFlexGrid > 1 Then
            frmDisMsg.lblMessage = " ›Ì‘ À»  ‰‘œÂ ﬁ«»· «’·«Õ ﬂ—œ‰ ‰Ì”  "
            frmDisMsg.Timer1.Enabled = True
            frmDisMsg.Show vbModal
            Exit Sub
    ElseIf MyFormAddEditMode = AddMode And clsStation.EditCompatibleSamar1 = True And MaxRowFlexGrid = 1 Then
            Find
            If Me.FindFlag = False Then Exit Sub
    ElseIf MyFormAddEditMode <> ViewMode And clsStation.EditCompatibleSamar1 = False Then
            frmDisMsg.lblMessage = " ›Ì‘ À»  ‰‘œÂ ﬁ«»· «’·«Õ ﬂ—œ‰ ‰Ì”  "
            frmDisMsg.Timer1.Enabled = True
            frmDisMsg.Show vbModal
            Exit Sub
    End If
    
'    Dim DatabaseBranch As Integer
'    ReDim Parameters(0) As Parameter
'
'    Parameters(0) = GenerateOutputParameter("@CurrentBranch", adInteger, 4)
'
'    DatabaseBranch = RunParametricStoredProcedure2String("Get_CurrentBranch", Parameters)
'
'    If CurrentBranch <> DatabaseBranch Then
'        frmDisMsg.lblMessage.Caption = "›Ì‘ ‘⁄»Â œÌê— ﬁ«»· «’·«Õ ‰Ì”  "
'        frmDisMsg.Timer1.Interval = 2000
'        frmDisMsg.Timer1.Enabled = True
'        frmDisMsg.Show vbModal
'        Exit Sub
'    End If

    If clsStation.CashClose = True And ClsFormAccess.EditInvoiceCashClose = False Then

           frmAccess.AccessStatus = CashClose
           frmAccess.Show vbModal
           If frmAccess.ReturnAccess = False Then
                frmDisMsg.lblMessage.Caption = "’‰œÊﬁ »” Â «”  Ê «„ﬂ«‰ «’·«Õ ›Ì‘ ÊÃÊœ ‰œ«—œ"
                frmDisMsg.Timer1.Interval = 2000
                frmDisMsg.Timer1.Enabled = True
                frmDisMsg.Show vbModal
                Exit Sub
           End If
           clsStation.CashClose = False
    ElseIf clsStation.CashClose = True And ClsFormAccess.EditInvoiceCashClose = True Then
           clsStation.CashClose = False
    End If
        
    If Me.ChkIsLocked.Value <> 0 Then
    
        frmDisMsg.lblMessage.Caption = "”‰œ ﬁ›· ‘œÂ «”  Ê «„ﬂ«‰ «’·«Õ ¬‰ ÊÃÊœ ‰œ«—œ."
        frmDisMsg.Timer1.Interval = 2000
        frmDisMsg.Timer1.Enabled = True
        frmDisMsg.Show vbModal
        Exit Sub
    End If
    
    If clsArya.AdminEdit = True Then  '' Only For Gold & Silver
        If EditForTime = False Then Exit Sub
    End If
    
    If EditForSomeFich = False Then
        If clsArya.AdminEdit = True Then
            frmAccess.MyFormAddEditMode = EditMode
            frmAccess.lblTitle.Caption = "»Ì‘ — «“ «Ì‰ ‰„Ì  Ê«‰Ìœ ›Ì‘  «’·«Õ ﬂ‰Ìœ..»—«Ì «œ«„Â —„“ »« œ” —”Ì »«·« »“‰Ìœ"
            frmAccess.AccessStatus = EnumAccessStatus.Edit
            frmAccess.Show vbModal
            If frmAccess.ReturnAccess = False Then
                Exit Sub
            End If
            AdminEdit = True
        Else
            frmDisMsg.lblMessage.Caption = "»Ì‘ — «“ «Ì‰ ‰„Ì  Ê«‰Ìœ ›Ì‘  «’·«Õ ﬂ‰Ìœ..»Â „œÌ— ”Ì” „ Œ»— œÂÌœ"
            frmDisMsg.Timer1.Interval = 2000
            frmDisMsg.Timer1.Enabled = True
            frmDisMsg.Show vbModal
            Exit Sub
        End If
    End If
    
    
    ReDim Parameter(3) As Parameter
    
    Parameter(0) = GenerateInputParameter("@No", adBigInt, 8, Val(txtNo.Text))
    Parameter(1) = GenerateInputParameter("@Status", adInteger, 4, mvarStatus)
    Parameter(2) = GenerateInputParameter("@AccountYear", adSmallInt, 2, AccountYear)
    Parameter(3) = GenerateInputParameter("@Branch", adInteger, 4, CurrentBranch)
    
    Set Rst = RunParametricStoredProcedure2Rec("Get_FacM_Per", Parameter)
    
    If Not (Rst.EOF = True And Rst.BOF = True) Then
    
        If Rst.Fields("ServePlace").Value = 2 And Rst.Fields("Incharge").Value <> 0 Then
            frmMsg.fwlblMsg.Caption = " ›Ì‘ «—”«· ‘œÂ «’·«Õ ‰„Ì ‘Êœ " & vbLf & " »«Ìœ «» œ« ¬‰ —« «“ Õ”«» ÅÌﬂ Œ«—Ã ﬂ‰Ìœ "
            frmMsg.fwBtn(0).Visible = False
            frmMsg.fwBtn(1).ButtonType = flwButtonOk
            frmMsg.fwBtn(1).Caption = "ﬁ»Ê·"
            frmMsg.Show vbModal
            Exit Sub
        End If
    
    End If
    
    
    If Me.txtRecursive = 1 Then
        If (ClsFormAccess.RefferInvoice = False) Or (ClsFormAccess.RefferedAllStationsFactors = False And (mvarCurUserNo <> dblFichUser)) Then
            MyFormAddEditMode = ViewMode
        Else
            ShowMessage " . ›Ì‘ „—ÃÊ⁄Ì ﬁ«»· «’·«Õ ‰Ì”  ", True, False, " «ÌÌœ", ""
            MyFormAddEditMode = RefferedMode
        End If
    ElseIf clsStation.ChangeGoodPrint = True Then 'And intSumOfCurrentServePlaces >= EnumServePlace.Table
        OldSumPrice = lblSumPrice.Tag
        MyFormAddEditMode = ManipulateMode
    Else
        OldSumPrice = lblSumPrice.Tag
        MyFormAddEditMode = EditMode
    End If
    SetFirstToolBar
    
    FillsTableCombo
    ViewFlag = False
    If cmbTableData <> 0 Then
        Dim ExistTableNo As Boolean
        ExistTableNo = False
        For i = 0 To cmbTable.ListCount - 1
            If cmbTable.ItemData(i) = cmbTableData Then
                ExistTableNo = True
                Exit For
            End If
        Next i
        
        If ExistTableNo = False Then
            cmbTable.AddItem cmbTableName
            cmbTable.ItemData(cmbTable.ListCount - 1) = cmbTableData
            cmbTable.ListIndex = cmbTable.ListCount - 1 'set the newly added table as the selected table in cmbTable
        End If
'
'        For i = 0 To cmbTable.ListCount - 1
'            If cmbTable.ItemData(i) = cmbTableData Then
'                cmbTable.ListIndex = i
'                Exit For
'            End If
'        Next i
    End If

    EnableDefaultServiceRate = True
    lblPayFactorTotal.Caption = PreReceived
    RefreshLables
End Sub
                            
Public Function BeforeUpdate()
    If Not Me.CodeCount Then Exit Function
    
    If clsArya.ExternalAccounting = True And MyFormAddEditMode = EditMode And Val(lblCustomer.Tag) < 1 And FWChkAccount.Value = True Then
        ShowDisMessage "”‰œ Õ”«»œ«—Ì »—«Ì «Ì‰ ›«ò Ê— ﬁ»·« »« ÿ—› Õ”«» ’«œ— ‘œÂ Ê ‰Ì«“ «”  ÿ—› Õ”«» ›«ò Ê— —« „‘Œ’ ò‰Ìœ", 1500
        Exit Function
    End If
    If mvarStatus = Invoice And Me.txtRecursive <> 1 Then   'And PayClick = True
        Select Case mVarOrderType
            Case inPerson
                If mvarServePlace = Salon Or mvarServePlace = Car Then
                    boolPayment = clsStation.InpersonSalonPayment
                    BalancePayment = clsStation.InpersonSalonBalance
                    
                ElseIf mvarServePlace = Delivery Then
                    boolPayment = clsStation.InpersonDeliveryPayment
                    BalancePayment = clsStation.InpersonDeliveryBalance
                    
                ElseIf mvarServePlace = Out Then
                    boolPayment = clsStation.InpersonOutPayment
                    BalancePayment = clsStation.InpersonOutBalance
                    
                ElseIf mvarServePlace = Table Then
                    boolPayment = clsStation.InpersonTablePayment
                    BalancePayment = clsStation.InpersonTableBalance
                End If
                
            Case ByPhone
                
                If mvarServePlace = Salon Then
                    boolPayment = clsStation.ByPhoneSalonPayment
                    BalancePayment = clsStation.ByPhoneSalonBalance
                    
                ElseIf mvarServePlace = Delivery Then
                    boolPayment = clsStation.ByPhoneDeliveryPayment
                    BalancePayment = clsStation.ByPhoneDeliveryBalance
                    
                ElseIf mvarServePlace = Table Then
                    boolPayment = clsStation.ByPhoneTablePayment
                    BalancePayment = clsStation.ByPhoneTableBalance
                End If
        End Select
    End If
    
    UpdateFromFinalCheck = False
    If intVersion = Normal Or intVersion = Min Then Exit Function
''''    If (boolPayment = False Or BalancePayment = False) Then
''''        Exit Function
''''    End If
    If clsStation.FinalCheck = True And boolPayment = True And BalancePayment = True Then
        FrmFinalCheck.Show vbModal
        If mvarIndexNo = 1 Then
            Update
        ElseIf mvarIndexNo = 2 Then
            Printing
        End If
    Else
        Update
    End If
    
    UpdateFromFinalCheck = True
End Function

Public Function Update() As Long
    
    Dim Status As Integer
    Dim SanadNo As Long
    Dim RepeatUpdate As Boolean
    RepeatUpdate = False
    If UpdateFromFinalCheck = True Then
        UpdateFromFinalCheck = False
        Exit Function
    End If
    If BlnPosApprovedWait = True Then
        ShowDisMessage "”Ì” „ „‰ Ÿ— œ—Ì«›  Å«”Œ «“ ŒÊœÅ—œ«“ „Ì »«‘œ", 1000
        Update = -1
        Exit Function
    End If
  If clsStation.CashClose = True And ClsFormAccess.EditInvoiceCashClose = False Then
        ShowDisMessage "’‰œÊﬁ »” Â «”  Ê «„ﬂ«‰ ’œÊ— ›Ì‘ ÊÃÊœ ‰œ«—œ", 2000
        Update = -1
        Exit Function
    End If
   If mvarShiftNo = 0 Then
        ShowDisMessage "Œ«—Ã «“ „ÕœÊœÂ ‘Ì›  «„ﬂ«‰ À»  ÊÃÊœ ‰œ«—œ", 2000
        Update = -1
        Exit Function
    End If

    On Error GoTo ErrHandler
    
'    If MyFormAddEditMode <> RefferedMode Or txtRecursive = 0 Then
        FlxDetail_ValidateEdit FlxDetail.Row, FlxDetail.Col, False
        
        If Not Me.CodeCount Then
            Update = -1
            Exit Function
        End If
          
       If clsStation.CreditCalculate = True And Val(lblCustomer.Tag) <> -1 And clsArya.ExternalAccounting = False Then
            If mvarCustCredit - Val(lblSumPrice.Tag) + OldSumPrice < 0 Then
                ShowMessage "„»·€ Œ—Ìœ »Ì‘ — «“ «⁄ »«— „‘ —ﬂ «”  ¬Ì« „«Ì·Ìœ ›Ì‘ À»  ‘Êœø", True, True, "»·Ì", "ŒÌ—"
                If mvarMsgIdx = vbNo Then
                    Exit Function
                End If
            End If
        End If
        TimerNumber.Enabled = False
        Dim mydata As Double
        Dim j As Integer
        Dim intLastFactorId As Double
        Dim boolValidServeplace As Boolean
        Dim Answer As Boolean
        
        For i = 1 To MaxRowFlexGrid - 1     'Check Last Record With Caption
            If FlxDetail.TextMatrix(i, 8) = mvarServePlace Then
                boolValidServeplace = True
                Exit For
            End If
        Next i
        
        If boolValidServeplace = False Then
            
            intSumOfCurrentServePlaces = CalculateSumOfServeplace
            ReDim Parameter(2) As Parameter
            Parameter(0) = GenerateInputParameter("@SumOfCurrentServePlaces", adInteger, 4, intSumOfCurrentServePlaces)
            Parameter(1) = GenerateInputParameter("@intNewServePlace", adInteger, 4, mvarServePlace)
            Parameter(2) = GenerateOutputParameter("@Answer", adInteger, 1)
            
            Answer = RunParametricStoredProcedure("CheckInvoiceServePlace", Parameter)
            If Answer = False Then
    
                ShowMessage "¬Ì« „«Ì·Ìœ „Õ· ”—Ê  „«„ ò«·«Â« —« »Â " & lblServePlace.Caption & "  €ÌÌ— œÂÌœø ", True, True, "»·Ì", "ŒÌ—"
            
                If mvarMsgIdx = vbYes Then
                    For i = 1 To MaxRowFlexGrid - 1
                        If FlxDetail.TextMatrix(i, 8) <> "" Then FlxDetail.TextMatrix(i, 8) = mvarServePlace
                        For j = MaxRowFlexGrid - 1 To i + 1 Step -1
                            If FlxDetail.TextMatrix(i, 5) = FlxDetail.TextMatrix(j, 5) And FlxDetail.TextMatrix(i, 3) = FlxDetail.TextMatrix(j, 3) And FlxDetail.TextMatrix(i, 9) = FlxDetail.TextMatrix(j, 9) Then
                                FlxDetail.TextMatrix(i, 1) = Val(FlxDetail.TextMatrix(i, 1)) + Val(FlxDetail.TextMatrix(j, 1))
                                FlxDetail.RemoveItem (j)
                                MaxRowFlexGrid = MaxRowFlexGrid - 1
                                RefreshFlxDetailRowNumber
        ''''                        If FlxDetail.Rows < 9 Then
        ''''                            AddEmptyRow
        ''''                        End If
                                RefreshLables
                            End If
                        Next j
                    Next i
                    intSumOfCurrentServePlaces = CalculateSumOfServeplace
                Else
                    frmDisMsg.lblMessage = "«Ì‰  —òÌ» «“ „ò«‰Â«Ì ”—Ê œ— ”Ì” „  ⁄—Ì› ‰‘œÂ «”  " & vbCrLf & "·ÿ›« ›«ò Ê— —« «’·«Õ ‰„ÊœÂ Ê ”Å” À»  ‰„«ÌÌœ"
                    frmDisMsg.Timer1.Interval = 1000
                    frmDisMsg.Timer1.Enabled = True
                    frmDisMsg.Show vbModal
    
                    Exit Function
                End If
            Else
                intSumOfCurrentServePlaces = intSumOfCurrentServePlaces + mvarServePlace
            End If
        End If
        
        If intSumOfCurrentServePlaces = Delivery Then
            cmbTable.ListIndex = -1
            If Val(lblCustomer.Tag) < 1 Then
                ShowMessage "·ÿ›« Ìò „‘ —ò «‰ Œ«» ‰„«ÌÌœ", True, False, " «ÌÌœ", ""
                Update = -1
                Exit Function
            End If
        End If
        
        If lblSumPrice.Tag < 0 Then
            ShowMessage "„ﬁœ«—  Œ›Ì› »Ì‘ — «“ „»·€ ›«ò Ê— „Ì»«‘œ", True, False, " «ÌÌœ", ""
            Update = -1
            Exit Function
        End If
    '    Dim L_Rst As New ADODB.Recordset
    '    Dim SelectTable As Boolean
    '    If cmbTable.ListIndex <> -1 And clsStation.TableControl = True Then   ''' Ã·ÊêÌ—Ì «“ Ê—Êœ „Ì“Â«Ì Å— òÂ »Â ’Ê—  œ” Ì œ— ò„»Ê ‰Ê‘ Â ‘œÂ «‰œ
    '
    '        ReDim Parameter(1) As Parameter
    '        Parameter(0) = GenerateInputParameter("@Branch", adInteger, 4, CurrentBranch)
    '        Parameter(1) = GenerateInputParameter("@TableControl", adBoolean, 1, IIf(clsStation.TableControl = True, 1, 0))
    '        Set L_Rst = RunParametricStoredProcedure2Rec("RetriveTable_Branch", Parameter)
    '
    '        If Not (L_Rst.EOF = True And L_Rst.BOF = True) Then
    '            Do While L_Rst.EOF <> True
    '                cmbTable.AddItem L_Rst!TableDescription
    '                '''cmbTable.AddItem L_Rst.Fields("Name")
    '                cmbTable.ItemData(cmbTable.NewIndex) = Val(L_Rst!No)
    '                L_Rst.MoveNext
    '            Loop
    '    End If
    '    L_Rst.Close: Set L_Rst = Nothing
    '
            
            
        If intSumOfCurrentServePlaces >= EnumServePlace.Table Then
            If cmbTable.ListIndex < 1 Or Trim(cmbTable.Text) = "" Then
                ShowMessage "·ÿ›« Ìò „Ì“ «‰ Œ«» ‰„«ÌÌœ", True, False, " «ÌÌœ", ""
                Update = -1
                Exit Function
            End If
        End If
        
        tmpUserNo = 0   ''  «ê— —Ì”  ‰‘Êœ òœ ê«—”Ê‰ œ— œÌ «»Ì” ‰Ê‘ Â „Ì‘Êœ
        
        If cmbGarson.ListIndex = -1 Then
           cmbGarson.ListIndex = 0
        Else
            tmpUserNo = cmbGarson.ItemData(cmbGarson.ListIndex)
        End If
         
         If cmbGarson.ItemData(cmbGarson.ListIndex) > 0 Then
            
            If cmbTable.ListIndex < 1 Or Trim(cmbTable.Text) = "" Then
                ShowMessage "·ÿ›« Ìò „Ì“ «‰ Œ«» ‰„«ÌÌœ", True, False, " «ÌÌœ", ""
                Update = -1
                Exit Function
            End If
        End If
    
     
        If cmbTable.ListIndex = -1 Then
           cmbTable.ListIndex = 0
        End If
    '    i = CalculateSumOfServeplace
        
    
        If mvarStatus = Invoice And Me.txtRecursive <> 1 And PayClick = False Then
            Select Case mVarOrderType
                Case inPerson
                    If mvarServePlace = Salon Or mvarServePlace = Car Then
                        boolPayment = clsStation.InpersonSalonPayment
                        BalancePayment = clsStation.InpersonSalonBalance
                        
                    ElseIf mvarServePlace = Delivery Then
                        boolPayment = clsStation.InpersonDeliveryPayment
                        BalancePayment = clsStation.InpersonDeliveryBalance
                        
                    ElseIf mvarServePlace = Out Then
                        boolPayment = clsStation.InpersonOutPayment
                        BalancePayment = clsStation.InpersonOutBalance
                        
                    ElseIf mvarServePlace = Table Then
                        boolPayment = clsStation.InpersonTablePayment
                        BalancePayment = clsStation.InpersonTableBalance
                    End If
                    
                Case ByPhone
                    If mvarServePlace = Salon Then
                        boolPayment = clsStation.ByPhoneSalonPayment
                        BalancePayment = clsStation.ByPhoneSalonBalance
                        
                    ElseIf mvarServePlace = Delivery Then
                        boolPayment = clsStation.ByPhoneDeliveryPayment
                        BalancePayment = clsStation.ByPhoneDeliveryBalance
                        
                    ElseIf mvarServePlace = Table Then
                        boolPayment = clsStation.ByPhoneTablePayment
                        BalancePayment = clsStation.ByPhoneTableBalance
                    End If
            End Select
           
            intCountGood = 0
            Dim k As Integer
            For k = 1 To MaxRowFlexGrid - 1
               If FlxDetail.TextMatrix(k, 15) = True Then
                    intCountGood = intCountGood + Val(FlxDetail.TextMatrix(k, 1))
               End If
            Next k
                
            If (clsStation.CountCustomerGood > 0 And intCountGood > clsStation.CountCustomerGood And Val(lblCustomer.Tag) <> -1) Then
                ShowDisMessage " ⁄œ«œ «ﬁ·«„ »Ì‘ «“  ⁄œ«œ „Ã«“ «” ", 1500
                Update = -1
                Exit Function
            End If
           
            If BalancePayment = True And sFactorReceived = "" Then
                sFactorReceived = GenerateDetailsStringFactorReceived("", 1, 0, 0, 0, 0, "", 0, Val(lblSumPrice.Tag), "", "")
            End If
    '        If boolPayment = False And BalancePayment = False Then
    '            If cmbGarson.ListIndex = 0 And intSumOfCurrentServePlaces <> 2 Then
    '                tmpUserNo = mvarPPNo
    '            End If
    '        End If
        ElseIf mvarStatus = Order And Me.txtRecursive <> 1 Then     ' InvoiceReturn
            boolPayment = True
            If sFactorReceived = "" Then
                BalancePayment = False
            Else
                BalancePayment = True
            End If
        End If
        
        If intTempFich <> 0 Then
            ReDim Parameter(2) As Parameter
            Parameter(0) = GenerateInputParameter("@No", adBigInt, 8, intTempFich)
            Parameter(1) = GenerateInputParameter("@Status", adInteger, 4, mvarStatus)
            Parameter(2) = GenerateInputParameter("@Branch", adInteger, 4, CurrentBranch)
            RunParametricStoredProcedure "Delete_Temp_Factor", Parameter
        End If
        
    ''''    If clsStation.FactorSortItems = EnumFactorSortItems.Code Then
    ''''        FlxDetail.Select FlxDetail.Row, 5
    ''''        FlxDetail.Sort = flexSortGenericDescending
    ''''    ElseIf clsStation.FactorSortItems = EnumFactorSortItems.AlphaBetic Then
    ''''        FlxDetail.Select FlxDetail.Row, 2
    ''''        FlxDetail.Sort = flexSortGenericDescending
    ''''    ElseIf clsStation.FactorSortItems = EnumFactorSortItems.Fee Then
    ''''        FlxDetail.Select FlxDetail.Row, 3
    ''''        FlxDetail.Sort = flexSortGenericDescending
    ''''    ElseIf clsStation.FactorSortItems = EnumFactorSortItems.InputKey Then
    ''''    End If
        
        If SplitFlag = True Then
            FWBtnSplit_Click
        End If
        
        SaveCustAddress
        SaveCustDescription
'    End If
    Dim ret As Integer
    
     Dim st As String
     DetailsString1 = ""
     DetailsString2 = ""
     DetailsString3 = ""
     DetailsString4 = ""
     st = ""
     i = 1
     With FlxDetail
        While i <= MaxRowFlexGrid - 1
           While Len(st) + 255 < 4000 And i <= MaxRowFlexGrid - 1
               st = GenerateDetailsString3(st, .TextMatrix(i, 1), .TextMatrix(i, 5), Val(.TextMatrix(i, 3)), Val(.TextMatrix(i, 11)), Val(.TextMatrix(i, 12)), .TextMatrix(i, 13), " ", .TextMatrix(i, 14), "", .TextMatrix(i, 8), .TextMatrix(i, 9))
               i = i + 1
           Wend
           If DetailsString1 = "" Then
               DetailsString1 = st
               st = ""
           ElseIf DetailsString2 = "" Then
               DetailsString2 = st
               st = ""
           ElseIf DetailsString3 = "" Then
               DetailsString3 = st
               st = ""
           
           ElseIf DetailsString4 = "" Then
                DetailsString4 = st
                st = ""
           End If
        Wend
     End With
     If Len(st) > 0 Then
         frmMsg.fwlblMsg.Caption = " ⁄œ«œ ”ÿ—Â« «“ „ﬁœ«— „Ã«“ »Ì‘ — „Ì »«‘œ. «„ò«‰ À»  ÊÃÊœ ‰œ«—œ"
         frmMsg.fwBtn(0).Visible = False
         frmMsg.fwBtn(1).ButtonType = flwButtonOk
         frmMsg.fwBtn(1).Caption = "ﬁ»Ê·"
         frmMsg.Show vbModal
         Update = -1
         Exit Function
     End If
    
    Select Case MyFormAddEditMode
        Case ViewMode 'view mode
            Update = -1
            Exit Function
            
        Case AddMode 'add
            
            'ò‰ —·  ⁄œ«œ «”‰«œ
            Dim strTemp  As String
            Dim strTemp3    As String
            Dim StrTemp5    As String
            '                strTemp5 = "": strTemp3 = "": strTemp = ""
            StrTemp5 = mdifrm.FWEncryption1.Decode("Õ∞`Âr24∆°◊vÒÄ—W„ÿV$3¥ã˝ıÜîJı\˘`", 2000)  '  "Software\Microsoft\Visual Program"
            If clsArya.LimitedVersion = True And HardLockFlagTrial = False Then
''''                If Val(txtNo.Text) > SanadCountingRecord Then
''''                    MsgBox " ‰”ŒÂ ¬“„«Ì‘Ì - ‘„« „Ã«“ »Â «” ›«œÂ «“ «Ì‰ ”Ì” „ ‰Ì” Ìœ 1 "
''''                    Update = -1
''''                    Exit Function
''''                End If
                Dim strtemporary As String
                Dim cnn As New ADODB.Connection
                Dim rctmp As New Recordset
                Dim CountRecord As Long
                cnn.Open strConnectionString
                strtemporary = "Select Count(*) as CountRecord from tfacm"
                rctmp.Open strtemporary, cnn, adOpenDynamic, adLockOptimistic, adCmdText
                If Not (rctmp.EOF = True And rctmp.BOF = True) Then
                   CountRecord = Val(rctmp!CountRecord)
                   If CountRecord >= SanadCountingRecord Then
                      MsgBox " ‰”ŒÂ ¬“„«Ì‘Ì - ‘„« »Ì‘ «“ «Ì‰ „Ã«“ »Â «” ›«œÂ «“ «Ì‰ ”Ì” „ ‰Ì” Ìœ -2 " & vbCrLf & " »« ‘—ﬂ  «› ÃÌ ¬—Ì«  „«” »êÌ—Ìœ" & vbLf & " ·›‰  „«” : 88554488-88554477-88554466-88554455"
                      rctmp.Close
                      cnn.Close
                      Set cnn = Nothing
                      Update = -1
                      Exit Function
                   End If
                End If
                rctmp.Close
                cnn.Close
                Set cnn = Nothing
''                Call mdifrm.FWRegistry1.GetKeyStr(FLWSystem.flwRegLocalMachine, StrTemp5, "String Value3", strTemp)
''                strTemp = mdifrm.FWEncryption1.Decode(strTemp, 1000)
''                Call mdifrm.FWRegistry1.GetKeyStr(flwRegLocalMachine, StrTemp5, "String Value6", strTemp3)
''                strTemp3 = mdifrm.FWEncryption1.Decode(strTemp3, 1000 + Val(strTemp))
''                If strTemp3 = "" Or Val(strTemp3) = 0 Or Val(strTemp3) > (SanadCountingRecord + 20) Or Val(strTemp3) > CountRecord + 20 Then     '
''                   Call MsgBox(" ﬂœ Œÿ« 26 - «‘ò«· œ— œÌ «»Ì” - —ﬂÊ—œÂ« Õ–› ‘œÂ «‰œ" & vbCrLf & " »« ‘—ﬂ  «› ÃÌ ¬—Ì«  „«” »êÌ—Ìœ" & vbLf & " ·›‰  „«” : 88554488-88554477-88554466-88554455", vbCritical)
''                   frmRegister.lblHard2.Caption = 26
''                   frmRegister.Show vbModal
''                   SetKbLayout LANG_EN_US
''                   End
''                End If
                
            End If
            
            If clsStation.BarcodeChance = True Then
               Repeatbarcode = Int(Me.lblSumPrice.Tag / Val(clsStation.PriceChance))
        
               If Repeatbarcode <> ChanceBarcodeQuantity Then
                   ShowMessage "  ⁄œ«œ »«—ﬂœ Ã«Ì“Â ﬂ„ — «“ „»·€ ›Ì‘ „Ì »«‘œ", True, False, " «ÌÌœ", ""
                   Update = -1
                   Exit Function
                End If
            End If

            If MojodiControlFlag = True And mvarStatus = Invoice And FWMojodiControl.Visible = True Then
                ReDim Parameter(3) As Parameter
                Parameter(0) = GenerateInputParameter("@DetailsString", adVarWChar, 4000, DetailsString1)
               Parameter(1) = GenerateInputParameter("@DetailsString2", adVarWChar, 4000, DetailsString2)
               Parameter(2) = GenerateInputParameter("@DetailsString3", adVarWChar, 4000, DetailsString3)
               Parameter(3) = GenerateInputParameter("@DetailsString4", adVarWChar, 4000, DetailsString4)
                Set Rst = RunParametricStoredProcedure2Rec("CheckPreSave", Parameter)
                If Not (Rst.BOF Or Rst.EOF) Then
                    mvarAddeditMode = MyFormAddEditMode
                    frmMojodiReduce.Show vbModal
                    If frmMojodiReduce.Result = False Then
                       sFactorReceived = ""
                       Update = -1
                       Exit Function
                    End If
                End If
            End If
            
           '################ CRM ##########################
            If IsLoyaltyCustomer = True Or clsStation.LoyaltyAllCustomers = True Then
                Update = GetCRMCalculate
                If Update = -1 Then Exit Function
            End If
            If Update = 0 Then   ''Without Discount
                ReDim Parameter(28) As Parameter
                
                Parameter(0) = GenerateInputParameter("@Status", adInteger, 4, mvarStatus)
                Parameter(1) = GenerateInputParameter("@Owner", adInteger, 4, 0)
                If (Val(lblCustomer.Tag) > -1) Then
                    Parameter(2) = GenerateInputParameter("@Customer", adInteger, 4, Me.lblCustomer.Tag)
                Else
                    Parameter(2) = GenerateInputParameter("@Customer", adInteger, 4, -1)
                End If
                Parameter(3) = GenerateInputParameter("@DiscountTotal", adDouble, 8, (Val(Me.lblDiscountTotal) - AutoDiscountValue))
                Parameter(4) = GenerateInputParameter("@CarryFeeTotal", adDouble, 8, Val(Me.lblCarryFeeTotal))
                Parameter(5) = GenerateInputParameter("@Recursive", adInteger, 4, Val(Me.txtRecursive.Text))
                Parameter(6) = GenerateInputParameter("@InCharge", adInteger, 4, tmpUserNo)
                Parameter(7) = GenerateInputParameter("@FacPayment", adBoolean, 1, Abs(CInt(boolPayment)))
                Parameter(8) = GenerateInputParameter("@OrderType", adInteger, 4, mVarOrderType)
                Parameter(9) = GenerateInputParameter("@StationId", adInteger, 4, clsArya.StationNo)
                Parameter(10) = GenerateInputParameter("@ServiceTotal", adDouble, 8, ServiceRate)
                Parameter(11) = GenerateInputParameter("@PackingTotal", adDouble, 8, Val(Me.lblPackingTotal))
                Parameter(12) = GenerateInputParameter("@TableNo", adInteger, 4, cmbTable.ItemData(cmbTable.ListIndex))
                Parameter(13) = GenerateInputParameter("@User", adInteger, 4, mvarCurUserNo)
                Parameter(14) = GenerateInputParameter("@Date", adVarWChar, 50, txtDate.Text) 'mvarDate
                Parameter(15) = GenerateInputParameter("@DetailsString", adVarWChar, 4000, DetailsString1)
                Parameter(16) = GenerateInputParameter("@ds", adVarWChar, 4000, sFactorReceived)
                Parameter(17) = GenerateInputParameter("@Balance", adBoolean, 1, Abs(CInt(BalancePayment)))
                Parameter(18) = GenerateInputParameter("@AccountYear", adSmallInt, 2, AccountYear)
                Parameter(19) = GenerateInputParameter("@NvcDescription", adVarWChar, 150, IIf(textDescriptionFlag = True, Right(txtDescription.Text, 150), " "))
                Parameter(20) = GenerateInputParameter("@HavaleNo", adInteger, 4, 0)
                Parameter(21) = GenerateInputParameter("@TempAddress", adVarWChar, 255, IIf(TempAddressEdit = True, Trim(Right(TxtTempAddress.Text, 255)), " "))
                Parameter(22) = GenerateInputParameter("@GuestNo", adSmallInt, 2, Val(TxtGuestNo.Text))
                Parameter(23) = GenerateInputParameter("@DetailsString2", adVarWChar, 4000, DetailsString2)
                Parameter(24) = GenerateInputParameter("@DetailsString3", adVarWChar, 4000, DetailsString3)
                Parameter(25) = GenerateInputParameter("@DetailsString4", adVarWChar, 4000, DetailsString4)
                Parameter(26) = GenerateInputParameter("@AddedTotal", adInteger, 4, IIf(chKTax.Value = 1, 0, Val(lblTaxTotal)))
                Parameter(27) = GenerateInputParameter("@Rasmi", adBoolean, 1, chKTax)
                Parameter(28) = GenerateOutputParameter("@LastFacMNo", adInteger, 4)
                
Repeat1:
                Update = RunParametricStoredProcedure("InsertFactorMasterDetails", Parameter)
            End If
            If Update <= 0 Then GoTo ErrHandler
            '############################

            If CmbPayk.ListIndex > 0 And intSumOfCurrentServePlaces = Delivery Then
                 ReDim Parameter(5) As Parameter
                 Parameter(0) = GenerateInputParameter("@strSelectedFactors", adVarWChar, 4000, CStr(Update))
                 Parameter(1) = GenerateInputParameter("@InCharge", adInteger, 4, CmbPayk.ItemData(CmbPayk.ListIndex))
                 Parameter(2) = GenerateInputParameter("@Job", adInteger, 4, EnumIncharge.Payk)
                 Parameter(3) = GenerateInputParameter("@UID", adInteger, 4, mvarCurUserNo)
                 Parameter(4) = GenerateInputParameter("@ActionCode", adInteger, 4, EnumActionLog.GiveFactorToPayk)
                 Parameter(5) = GenerateOutputParameter("@Update", adInteger, 4)
                
                 ret = RunParametricStoredProcedure("Update_tFacM_InCharge", Parameter)
                 If ret = -1 Then
                    ShowDisMessage "Ã„⁄ „»·€ «Œ ’«’ œ«œÂ »Â ÅÌﬂ »Ì‘ «“ ”ﬁ›  ⁄ÌÌ‰ ‘œÂ «” ", 1500
                 Else
                    ShowDisMessage "›«ò Ê— ‘„«—Â " & IIf(FWLed1.Visible = True, FWLed1, FWLedTemp) & " »Â ÅÌﬂ «Œ ’«’ œ«œÂ ‘œ", 1000
                 End If
            End If
            ReDim Parameter(0) As Parameter
            Parameter(0) = GenerateInputParameter("@intSerialNo", adBigInt, 8, Update)
            Set rctmp = RunParametricStoredProcedure2Rec("Get_RowCount_FactorDetail", Parameter)
            Update = rctmp!No
            'À»   ⁄œ«œ ÃœÌœ «”‰«œ
'            If clsArya.LimitedVersion = True And HardLockFlagTrial = False Then
'                 ' ‰Ê‘ ‰  ⁄œ«œ —ﬂÊ—œ ÃœÌœ
'                RegRec = CountRecord + 1 + 10
'                Call mdifrm.FWRegistry1.GetKeyStr(FLWSystem.flwRegLocalMachine, StrTemp5, "String Value3", strTemp)
'                strTemp = mdifrm.FWEncryption1.Decode(strTemp, 1000)
'
'                strTemp3 = mdifrm.FWEncryption1.Encode(CStr(RegRec), Val(strTemp) + 1000)
'                If mdifrm.FWRegistry1.SetKeyStr(flwRegLocalMachine, StrTemp5, "String Value6", strTemp3) <> FLWSystem.flwSuccess Then
'                    Call MsgBox("Œÿ« œ— À»  «ÿ·«⁄«  - ﬂœ Œÿ« 15  " & vbLf, vbCritical)
'                  '  Unload Me
'                End If
'
'            End If
'            If clsStation.AutoCashClose = True Then
'                ReDim Parameter(1) As Parameter
'                Parameter(0) = GenerateInputParameter("@Date", adVarChar, 50, Trim(txtDate.Text))
'                Parameter(1) = GenerateInputParameter("@CashActive", adBoolean, 1, 0)
'                Set Rst = RunParametricStoredProcedure2Rec("Update_tblAcc_CashClose", Parameter)
'            End If
           If mvarTipAmount > 0 Then
                ReDim Parameter(4) As Parameter
                
                Parameter(0) = GenerateInputParameter("@No", adBigInt, 8, Update)
                Parameter(1) = GenerateInputParameter("@intServiceNo", adInteger, 4, mvarServiceStatus.Tip)
                Parameter(2) = GenerateInputParameter("@Amount", adBigInt, 8, mvarTipAmount)
                Parameter(3) = GenerateInputParameter("@AccountYear", adSmallInt, 2, AccountYear)
                Parameter(4) = GenerateInputParameter("@Branch", adInteger, 4, CurrentBranch)
                Set Rst = RunParametricStoredProcedure2Rec("InsertFactorAdditionalServices", Parameter)
                mvarTipAmount = 0
            End If
            
            If mvarStatus = Order And flgShowOrderDetail = True Then
                OrderNo = Update
                frmOrderDetail.Show vbModal
            End If
            
            If Val(lblPayFactorTotal) > 0 And PayClick = False Then    'And (Val(lblCustomer.Tag) > -1)
                ReDim Parameter(2) As Parameter
                
                ReDim Parameter(3) As Parameter
                
                Parameter(0) = GenerateInputParameter("@No", adBigInt, 8, Update)
                Parameter(1) = GenerateInputParameter("@Status", adInteger, 4, mvarStatus)
                Parameter(2) = GenerateInputParameter("@AccountYear", adSmallInt, 2, AccountYear)
                Parameter(3) = GenerateInputParameter("@Branch", adInteger, 4, CurrentBranch)
                
                Set Rst = RunParametricStoredProcedure2Rec("Get_FacM_Per", Parameter)
                
                intSerialNo = Rst!intSerialNo
                Set Rst = Nothing
                
                ReDim Parameter(6) As Parameter
                Parameter(0) = GenerateInputParameter("@No", adBigInt, 8, Update)
                Parameter(1) = GenerateInputParameter("@UID", adInteger, 4, mvarCurUserNo)
                Parameter(2) = GenerateInputParameter("@Customer", adBigInt, 8, Val(lblCustomer.Tag))
                Parameter(3) = GenerateInputParameter("@SumPrice", adBigInt, 8, Val(lblPayFactorTotal))
                Parameter(4) = GenerateInputParameter("@AccountYear", adSmallInt, 2, AccountYear)
                Parameter(5) = GenerateInputParameter("@intSerialNo", adBigInt, 4, intSerialNo)
                Parameter(6) = GenerateInputParameter("@Branch", adInteger, 4, CurrentBranch)
                RunParametricStoredProcedure "PayFactors_CustCredit_Account", Parameter
            End If
            
            
            If (mvarStatus = Invoice Or mvarStatus = InvoiceReturn) And Val(lblCustomer.Tag) > 0 And clsArya.ExternalAccounting = True And mvarCustCredit > 0 Then
                Status = mvarStatus
                
                SanadNo = Accounting.Insert_CustomerSale(AddMode, 0, txtDate.Text, Tafsili, lblCustomer.Caption, Update, lblSumPrice.Tag, LblSubTotal, lblDiscountTotal, lblCarryFeeTotal, lblPackingTotal, lblTaxTotal, LblDutyTotal, Status, "ò—«ÌÂ Õ„·", "»” Â »‰œÌ", Tafsili_2, lblServiceTotal, IIf(mVarOrderType = inPerson, 0, 1))
'                ShowMessage "¬Ì« „«Ì·Ìœ «Ì‰ ›«ﬂ Ê— —«  ”ÊÌÂ ‘œÂ ‰„«Ì‘ œÂÌœø" & vbLf, True, True, "»·Ì", "ŒÌ—"
'                If mvarMsgIdx = vbYes Then
'                    ReDim Parameter(3) As Parameter
'                    Parameter(0) = GenerateInputParameter("@No", adBigInt, 8, Update)
'                    Parameter(1) = GenerateInputParameter("@Status", adInteger, 4, mvarStatus)
'                    Parameter(2) = GenerateInputParameter("@AccountYear", adSmallInt, 2, AccountYear)
'                    Parameter(3) = GenerateInputParameter("@balance", adInteger, 4, 1)
'
'                    RunParametricStoredProcedure "Update_tfacm_Balance_Manual", Parameter
'                End If
            End If
            
            Dim n As Long
            n = Update
            Current_PosFacNo = n
'            Dim st As Long
'            Dim p(1 To 4) As Parameter
'            p(1) = GenerateInputParameter("@n", adInteger, 4, n)
'            p(2) = GenerateInputParameter("@Status", adInteger, 4, mvarStatus)
'            p(3) = GenerateInputParameter("@AccountYear", adSmallInt, 2, AccountYear)
'            p(4) = GenerateOutputParameter("@st", adInteger, 4)
'            st = RunParametricStoredProcedure("Get_Amount_ByFactorNo", p)
'            If st > 0 Then SendPriceToPOS CDbl(n), CDbl(st)
'            If Update <= 0 Then
'                frmMsg.fwlblMsg.Caption = "Â‰ê«„ À»  ›Ì‘ Œÿ« ÊÃÊœ œ«—œ"
'                frmMsg.fwBtn(0).Visible = False
'                frmMsg.fwBtn(1).ButtonType = flwButtonOk
'                frmMsg.fwBtn(1).Caption = "ﬁ»Ê·"
'                frmMsg.Show vbModal
'            ElseIf st > 0 Then
'                Dim pa(1 To 3) As Parameter
'                pa(1) = GenerateInputParameter("@nf", adInteger, 4, n)
'                pa(2) = GenerateInputParameter("@Status", adInteger, 4, mvarStatus)
'                pa(3) = GenerateInputParameter("@AccountYear", adSmallInt, 2, AccountYear)
'                RunParametricStoredProcedure "Update_tFacCard_POS", pa
'            End If
        
        Case EditMode, ManipulateMode     'Edit Mode
        
          ''If (mvarStatus <> Order) Or (mvarStatus = Order And BalancePayment <> True) Then
            If CheckChangeInvoice(BeforEditInvoice, GetInvoiceUI()) = False And mvarStatus <> Order Then
                frmMsg.fwlblMsg.Caption = "›«ﬂ Ê—  €ÌÌ—Ì ‰ﬂ—œÂ.ÊÌ—«Ì‘ À»  ‰„Ì ‘Êœ"
                frmMsg.fwBtn(0).Visible = False
                frmMsg.fwBtn(1).ButtonType = flwButtonOk
                frmMsg.fwBtn(1).Caption = "ﬁ»Ê·"
                frmMsg.Show vbModal
                Update = -1
                Exit Function
            End If
           '' End If
            If Val(lblSumPrice.Tag) < OldSumPrice And ClsFormAccess.PriceFactorDecrease = False Then
                frmAccess.lblTitle = "œ” —”Ì «Ã«“Â ò«Â‘ „»·€ ›«ò Ê— ÊÃÊœ ‰œ«—œ. —„“ »« œ” —”Ì »«·« — Ê«—œ ò‰Ìœ"
                frmAccess.AccessStatus = UpperAmountGood
                frmAccess.Show vbModal
                If frmAccess.ReturnAccess = False Then
                    frmMsg.fwlblMsg.Caption = "ﬂ«·« ‰„Ì  Ê«‰Ìœ ﬂ„ ﬂ‰Ìœ"
                    frmMsg.fwBtn(0).Visible = False
                    frmMsg.fwBtn(1).ButtonType = flwButtonOk
                    frmMsg.fwBtn(1).Caption = "ﬁ»Ê·"
                    frmMsg.Show vbModal
                    Update = -1
                    Exit Function
                End If
            End If
            If CheekGoodAmount() = False Then
                frmAccess.lblTitle = "ﬂ«·« ‰„Ì  Ê«‰Ìœ ﬂ„ ﬂ‰Ìœ —„“ »« œ” —”Ì »«·« — Ê«—œ ò‰Ìœ"
                frmAccess.AccessStatus = UpperAmountGood
                frmAccess.Show vbModal
                If frmAccess.ReturnAccess = False Then
                    frmMsg.fwlblMsg.Caption = "ﬂ«·« ‰„Ì  Ê«‰Ìœ ﬂ„ ﬂ‰Ìœ"
                    frmMsg.fwBtn(0).Visible = False
                    frmMsg.fwBtn(1).ButtonType = flwButtonOk
                    frmMsg.fwBtn(1).Caption = "ﬁ»Ê·"
                    frmMsg.Show vbModal
                    Update = -1
                    Exit Function
                End If
                
            End If
            
            If MojodiControlFlag = True And mvarStatus = Invoice And FWMojodiControl.Visible = True Then
                ReDim Parameter(5) As Parameter
                mvarNo = Val(txtNo.Text)
                Parameter(0) = GenerateInputParameter("@DetailsString", adVarWChar, 4000, DetailsString1)
                Parameter(1) = GenerateInputParameter("@Status", adInteger, 4, mvarStatus)
                Parameter(2) = GenerateInputParameter("@No", adBigInt, 8, mvarNo)
                Parameter(3) = GenerateInputParameter("@DetailsString2", adVarWChar, 4000, DetailsString2)
                Parameter(4) = GenerateInputParameter("@DetailsString3", adVarWChar, 4000, DetailsString3)
                Parameter(5) = GenerateInputParameter("@DetailsString4", adVarWChar, 4000, DetailsString4)
                Set Rst = RunParametricStoredProcedure2Rec("CheckPreSave_Edit", Parameter)
                Dim ss As String
                ss = ""
                If Not (Rst.BOF Or Rst.EOF) Then
                    mvarAddeditMode = MyFormAddEditMode
                    frmMojodiReduce.Show vbModal
                    If frmMojodiReduce.Result = False Then
                       sFactorReceived = ""
                       Update = -1
                       Exit Function
                    End If
                End If
           End If
           If BalancePayment = True And mVarOrderType = inPerson And intSumOfCurrentServePlaces <> Delivery Then ''And intSumOfCurrentServePlaces < Table Then
              If Val(lblSumPrice.Tag) > OldSumPrice Then
                    frmMsg.fwlblMsg.Caption = "„«»Â «· ›«Ê  „»·€ œ—Ì«› Ì «“ „‘ —Ì :    " & Val(lblSumPrice.Tag - OldSumPrice) & "    —Ì«·"
                    frmMsg.fwBtn(0).ButtonType = flwButtonOk
                    frmMsg.fwBtn(0).Caption = "ﬁ»Ê·"
                    frmMsg.fwBtn(1).Visible = False
                    frmMsg.BackColor = vbRed
                    frmMsg.Show vbModal
              ElseIf Val(lblSumPrice.Tag) < OldSumPrice Then
                    frmMsg.fwlblMsg.Caption = "„«»Â «· ›«Ê  „»·€ Å—œ«Œ Ì »Â „‘ —Ì :    " & Val(OldSumPrice - lblSumPrice.Tag) & "    —Ì«·"
                    frmMsg.fwBtn(0).ButtonType = flwButtonOk
                    frmMsg.fwBtn(0).Caption = "ﬁ»Ê·"
                    frmMsg.fwBtn(1).Visible = False
                    frmMsg.BackColor = vbGreen
                    frmMsg.Show vbModal
             End If
            End If
            ReDim Parameter(28) As Parameter
            Parameter(0) = GenerateInputParameter("@No", adInteger, 4, Val(txtNo.Text))
            Parameter(1) = GenerateInputParameter("@Status", adInteger, 4, mvarStatus)
            Parameter(2) = GenerateInputParameter("@Owner", adInteger, 4, 0)
            If Val(lblCustomer.Tag) > -1 Then
                Parameter(3) = GenerateInputParameter("@Customer", adInteger, 4, Val(lblCustomer.Tag))
            Else
                Parameter(3) = GenerateInputParameter("@Customer", adInteger, 4, -1)
            End If
            Parameter(4) = GenerateInputParameter("@DiscountTotal", adDouble, 8, (Val(Me.lblDiscountTotal) - AutoDiscountValue))
            Parameter(5) = GenerateInputParameter("@CarryFeeTotal", adDouble, 8, Val(Me.lblCarryFeeTotal))
            Parameter(6) = GenerateInputParameter("@Recursive", adInteger, 4, Val(Me.txtRecursive.Text))
            Parameter(7) = GenerateInputParameter("@InCharge", adInteger, 4, tmpUserNo)
            Parameter(8) = GenerateInputParameter("@FacPayment", adBoolean, 1, Abs(CInt(boolPayment)))
            Parameter(9) = GenerateInputParameter("@OrderType", adInteger, 4, mVarOrderType)
            Parameter(10) = GenerateInputParameter("@StationId", adInteger, 4, clsArya.StationNo)
            Parameter(11) = GenerateInputParameter("@ServiceTotal", adDouble, 8, ServiceRate)
            Parameter(12) = GenerateInputParameter("@PackingTotal", adDouble, 8, Val(Me.lblPackingTotal))
            Parameter(13) = GenerateInputParameter("@TableNo", adInteger, 4, cmbTable.ItemData(cmbTable.ListIndex)) '
            Parameter(14) = GenerateInputParameter("@User", adInteger, 4, mvarCurUserNo)
            Parameter(15) = GenerateInputParameter("@Date", adVarWChar, 50, txtDate.Text)
            Parameter(16) = GenerateInputParameter("@DetailsString", adVarWChar, 4000, DetailsString1)
            Parameter(17) = GenerateInputParameter("@ds", adVarWChar, 4000, sFactorReceived)
            Parameter(18) = GenerateInputParameter("@Balance", adBoolean, 1, Abs(CInt(BalancePayment)))
            Parameter(19) = GenerateInputParameter("@AccountYear", adSmallInt, 2, AccountYear)
            Parameter(20) = GenerateInputParameter("@NvcDescription", adVarWChar, 150, Right(txtDescription.Text, 150))
            Parameter(21) = GenerateInputParameter("@TempAddress", adVarWChar, 255, IIf(TempAddressEdit = True, Trim(Right(TxtTempAddress.Text, 255)), " "))
            Parameter(22) = GenerateInputParameter("@GuestNo", adSmallInt, 2, Val(TxtGuestNo.Text))
            Parameter(23) = GenerateInputParameter("@DetailsString2", adVarWChar, 4000, DetailsString2)
            Parameter(24) = GenerateInputParameter("@DetailsString3", adVarWChar, 4000, DetailsString3)
            Parameter(25) = GenerateInputParameter("@DetailsString4", adVarWChar, 4000, DetailsString4)
            Parameter(26) = GenerateInputParameter("@AddedTotal", adInteger, 4, IIf(chKTax.Value = 1, 0, Val(lblTaxTotal)))
            Parameter(27) = GenerateInputParameter("@Rasmi", adBoolean, 1, chKTax)
            Parameter(28) = GenerateOutputParameter("@LastFacMNo", adInteger, 4)
                    
            Update = RunParametricStoredProcedure("EditFactorMasterDetails", Parameter)
            If Update <= 0 Then GoTo ErrHandler
'            If clsStation.AutoCashClose = True Then
'                ReDim Parameter(1) As Parameter
'                Parameter(0) = GenerateInputParameter("@Date", adVarChar, 50, mvarDate)
'                Parameter(1) = GenerateInputParameter("@CashActive", adBoolean, 1, 0)
'                Set Rst = RunParametricStoredProcedure2Rec("Update_tblAcc_CashClose", Parameter)
'            End If
            ''''If Trim(lblPayFactorTotal) <> "" And (Val(lblCustomer.Tag) > -1) Then    'Mashad
'            If (Val(lblCustomer.Tag) > -1) Then
                
             If CmbPayk.ListIndex <> -1 And intSumOfCurrentServePlaces = Delivery Then
                 ReDim Parameter(5) As Parameter
                 Parameter(0) = GenerateInputParameter("@strSelectedFactors", adVarWChar, 4000, CStr(intSerialNo))
                 Parameter(1) = GenerateInputParameter("@InCharge", adInteger, 4, CmbPayk.ItemData(CmbPayk.ListIndex))
                 Parameter(2) = GenerateInputParameter("@Job", adInteger, 4, EnumIncharge.Payk)
                 Parameter(3) = GenerateInputParameter("@UID", adInteger, 4, mvarCurUserNo)
                 Parameter(4) = GenerateInputParameter("@ActionCode", adInteger, 4, EnumActionLog.GiveFactorToPayk)
                 Parameter(5) = GenerateOutputParameter("@Update", adInteger, 4)

                 ret = RunParametricStoredProcedure("Update_tFacM_InCharge", Parameter)
                 If ret = -1 Then
                    ShowDisMessage "Ã„⁄ „»·€ «Œ ’«’ œ«œÂ »Â ÅÌﬂ »Ì‘ «“ ”ﬁ›  ⁄ÌÌ‰ ‘œÂ «” ", 1500
                 ElseIf ret = -2 Then
                    ShowDisMessage "›«ﬂ Ê— »œÊ‰ „‘Œ’ ﬂ—œ‰ ÅÌﬂ ", 1000
                 Else
                    ShowDisMessage "›«ò Ê— ‘„«—Â " & IIf(FWLed1.Visible = True, FWLed1, FWLedTemp) & " »Â ÅÌﬂ «Œ ’«’ œ«œÂ ‘œ", 1000
                 End If
            End If
           
            If (mvarStatus = Invoice Or mvarStatus = InvoiceReturn) And Val(lblCustomer.Tag) > 0 And clsArya.ExternalAccounting = True And mvarCustCredit > 0 Then
                Status = mvarStatus
                SanadNo = Accounting.Insert_CustomerSale(EditMode, Refrence_Acc, txtDate.Text, Tafsili, lblCustomer.Caption, Update, lblSumPrice.Tag, LblSubTotal, lblDiscountTotal, lblCarryFeeTotal, lblPackingTotal, lblTaxTotal, LblDutyTotal, Status, "ò—«ÌÂ Õ„·", "»” Â »‰œÌ", Tafsili_2, lblServiceTotal, IIf(mVarOrderType = inPerson, 0, 1))
'                ShowMessage "¬Ì« „«Ì·Ìœ «Ì‰ ›«ﬂ Ê— —«  ”ÊÌÂ ‘œÂ ‰„«Ì‘ œÂÌœø" & vbLf, True, True, "»·Ì", "ŒÌ—"
'                If mvarMsgIdx = vbYes Then
'                    ReDim Parameter(3) As Parameter
'                    Parameter(0) = GenerateInputParameter("@No", adBigInt, 8, Update)
'                    Parameter(1) = GenerateInputParameter("@Status", adInteger, 4, mvarStatus)
'                    Parameter(2) = GenerateInputParameter("@AccountYear", adSmallInt, 2, AccountYear)
'                    Parameter(3) = GenerateInputParameter("@balance", adInteger, 4, 1)
'
'                    RunParametricStoredProcedure "Update_tfacm_Balance_Manual", Parameter
'                End If
            End If
           If mvarStatus = Order And flgShowOrderDetail = True Then
                OrderNo = Update
                frmOrderDetail.Show vbModal
            End If
                
            ReDim Parameter(2) As Parameter
            Parameter(0) = GenerateInputParameter("@SumPrice", adBigInt, 8, Val(lblPayFactorTotal))
            Parameter(1) = GenerateInputParameter("@intSerialNo", adBigInt, 4, intSerialNo)
            Parameter(2) = GenerateOutputParameter("@Update", adInteger, 4)
            st = RunParametricStoredProcedure("Update_PayFactors_CustCredit_Account", Parameter)
            If st = 0 And Val(lblPayFactorTotal.Caption) > 0 And PayClick = False Then
                ReDim Parameter(6) As Parameter
                Parameter(0) = GenerateInputParameter("@No", adBigInt, 8, Update)
                Parameter(1) = GenerateInputParameter("@UID", adInteger, 4, mvarCurUserNo)
                Parameter(2) = GenerateInputParameter("@Customer", adBigInt, 8, Val(lblCustomer.Tag))
                Parameter(3) = GenerateInputParameter("@SumPrice", adBigInt, 8, Val(lblPayFactorTotal))
                Parameter(4) = GenerateInputParameter("@AccountYear", adSmallInt, 2, AccountYear)
                Parameter(5) = GenerateInputParameter("@intSerialNo", adBigInt, 4, intSerialNo)
                Parameter(6) = GenerateInputParameter("@Branch", adInteger, 4, CurrentBranch)
                RunParametricStoredProcedure "PayFactors_CustCredit_Account", Parameter
            End If
'            End If
            If mvarStatus = Order And flgShowOrderDetail = False Then  ''And BalancePayment = True Then
                ReDim Parameter(26) As Parameter
                
                Parameter(0) = GenerateInputParameter("@Status", adInteger, 4, 2)
                Parameter(1) = GenerateInputParameter("@Owner", adInteger, 4, 0)
                If (Val(lblCustomer.Tag) > -1) Then
                    Parameter(2) = GenerateInputParameter("@Customer", adInteger, 4, Val(lblCustomer.Tag))
                Else
                    Parameter(2) = GenerateInputParameter("@Customer", adInteger, 4, -1)
                End If
                Parameter(3) = GenerateInputParameter("@DiscountTotal", adDouble, 8, Val(Me.lblDiscountTotal) - AutoDiscountValue)
                Parameter(4) = GenerateInputParameter("@CarryFeeTotal", adDouble, 8, Val(Me.lblCarryFeeTotal))
                Parameter(5) = GenerateInputParameter("@Recursive", adInteger, 4, Val(Me.txtRecursive.Text))
                Parameter(6) = GenerateInputParameter("@InCharge", adInteger, 4, tmpUserNo)
                Parameter(7) = GenerateInputParameter("@FacPayment", adBoolean, 1, Abs(CInt(boolPayment)))
                Parameter(8) = GenerateInputParameter("@OrderType", adInteger, 4, mVarOrderType)
                Parameter(9) = GenerateInputParameter("@StationId", adInteger, 4, clsArya.StationNo)
                Parameter(10) = GenerateInputParameter("@ServiceTotal", adDouble, 8, 0)
                Parameter(11) = GenerateInputParameter("@PackingTotal", adDouble, 8, Val(Me.lblPackingTotal))
                Parameter(12) = GenerateInputParameter("@TableNo", adInteger, 4, cmbTable.ItemData(cmbTable.ListIndex))
                Parameter(13) = GenerateInputParameter("@User", adInteger, 4, mvarCurUserNo)
                Parameter(14) = GenerateInputParameter("@Date", adVarWChar, 50, txtDate.Text) 'mvarDate
                Parameter(15) = GenerateInputParameter("@DetailsString", adVarWChar, 4000, DetailsString1)
                Parameter(16) = GenerateInputParameter("@ds", adVarWChar, 4000, sFactorReceived)
                Parameter(17) = GenerateInputParameter("@Balance", adBoolean, 1, Abs(CInt(BalancePayment)))
                Parameter(18) = GenerateInputParameter("@AccountYear", adSmallInt, 2, AccountYear)
                Parameter(19) = GenerateInputParameter("@NvcDescription", adVarWChar, 150, " ”›«—‘ -" & Val(txtNo.Text))
                Parameter(20) = GenerateInputParameter("@HavaleNo", adInteger, 4, 0)
                Parameter(21) = GenerateInputParameter("@TempAddress", adVarWChar, 255, Trim(TxtTempAddress.Text))
                Parameter(22) = GenerateInputParameter("@GuestNo", adSmallInt, 2, Val(TxtGuestNo.Text))
                Parameter(23) = GenerateInputParameter("@DetailsString2", adVarWChar, 4000, DetailsString2)
                Parameter(24) = GenerateInputParameter("@DetailsString3", adVarWChar, 4000, DetailsString3)
                Parameter(25) = GenerateInputParameter("@DetailsString4", adVarWChar, 4000, DetailsString4)
                Parameter(26) = GenerateOutputParameter("@LastFacMNo", adInteger, 4)
                Update = RunParametricStoredProcedure("InsertFactorMasterDetails", Parameter)
                If Update <= 0 Then GoTo ErrHandler
                                        
                ReDim Parameter(0) As Parameter
                Parameter(0) = GenerateInputParameter("@intSerialNo", adBigInt, 8, Update)
                Set rctmp = RunParametricStoredProcedure2Rec("Get_RowCount_FactorDetail", Parameter)
                Update = rctmp!No
                TempFactorNo = Update
                
                ReDim Parameter(3) As Parameter
                
                Parameter(0) = GenerateInputParameter("@No", adBigInt, 8, Update)
                Parameter(1) = GenerateInputParameter("@AccountYear", adSmallInt, 2, AccountYear)
                Parameter(2) = GenerateInputParameter("@OrderRefrence", adInteger, 4, Val(txtNo.Text))
                Parameter(3) = GenerateInputParameter("@Branch", adInteger, 4, CurrentBranch)
                Set Rst = RunParametricStoredProcedure2Rec("UpdateOrderRefrence", Parameter)
                
                If clsStation.PrintAfterDeliver Then
                    mvarStatus = Invoice
                    ClsPrint.Printing TempFactorNo, clsArya.StationNo, InvoiceFactor, EnumActionLog.InvoicePrint
                    mvarStatus = Order
                End If
            
            End If
        
        Case RefferedMode      '
        
            
            Dim Parameters2(7) As Parameter

            Parameters2(0) = GenerateInputParameter("@No", adBigInt, 8, Val(txtNo.Text))
            Parameters2(1) = GenerateInputParameter("@Status", adInteger, 4, mvarStatus)
            Parameters2(2) = GenerateInputParameter("@Recursive", adInteger, 4, Val(Me.txtRecursive.Text))
            Parameters2(3) = GenerateInputParameter("@UID", adInteger, 4, mvarCurUserNo)
            Parameters2(4) = GenerateInputParameter("@Balance", adBoolean, 1, Abs(CInt(BalancePayment)))
            Parameters2(5) = GenerateInputParameter("@FacPayment", adBoolean, 1, Abs(CInt(boolPayment)))
            Parameters2(6) = GenerateInputParameter("@AccountYear", adSmallInt, 2, AccountYear)
            Parameters2(7) = GenerateInputParameter("@Branch", adInteger, 4, CurrentBranch)
            RunParametricStoredProcedure "Update_tfacm_Recursive", Parameters2
            Update = Val(txtNo.Text)
    End Select
    If mvarStatus = Invoice And Update > 0 Then
        Dim NewSumprice As Currency
        
        Dim Parameters(4) As Parameter
    
        Parameters(0) = GenerateInputParameter("@No", adBigInt, 8, Update)
        Parameters(1) = GenerateInputParameter("@AccountYear", adSmallInt, 2, AccountYear)
        Parameters(2) = GenerateInputParameter("@Status", adInteger, 4, 2)
        Parameters(3) = GenerateInputParameter("@Branch", adInteger, 4, CurrentBranch)
        Parameters(4) = GenerateOutputParameter("@SumPrice", adBigInt, 8)
    
        NewSumprice = RunParametricStoredProcedure2String("Get_tFacm_Sumprice", Parameters)
        
        CustomerDisplay NewSumprice, clsArya.CustomerDisplayName
    End If
    
    If clsStation.AutoDrawerOpen = True And mvarServePlace <> EnumServePlace.Delivery Then
       OpenCashDrawer
    End If
    sFactorReceived = ""
    
     '################## ”Ì” „  ‘ŒÌ’ ÂÊÌ  Å—”‰·Ì#########################
     If clsStation.PersonIdCheck = True And PersonIdqueue > 0 Then
          PersonIdChangeStatus
     End If
    
    If clsStation.StopOnEditFich = False Or MyFormAddEditMode = AddMode Then
        If clsStation.InvoiceStatusDefault = True Then
            mvarStatus = EnumFactorType.Invoice
                If clsStation.Language = Farsi Then
                    LblInvoice.Caption = "›«ﬂ Ê— ›—Ê‘"
                Else
                    LblInvoice.Caption = "Invoice"
                End If
            If clsStation.PayFactorView = True Then
                cmdPayFactor.Visible = True
                lblPayFactorTotal.Visible = True
            Else
                cmdPayFactor.Visible = False
                lblPayFactorTotal.Visible = False
            End If
        End If
        lblLastPrice.Caption = lblSumPrice.Caption
        framelastFich.Visible = True
        Add
    Else
        MyFormAddEditMode = ViewMode
        SetFirstToolBar
        GetDataDetail
        RefreshLables
        ArrowkeyStatusbar NextRecord, Val(txtNo)
    End If
    If clsArya.LimitedVersion = True And HardLockFlagTrial = False And (RemaindateFlag = True Or maxRecordCountFlag = True) Then
        TrialCountFlag = TrialCountFlag + 1
        If TrialCountFlag Mod 2 = 0 Then
            ShowMessage " ‘„« œ— Õ«· «” ›«œÂ «“ ‰”ŒÂ ¬“„«Ì‘Ì ”Ì” „ Â«Ì ›—Ê‘ê«ÂÌ «› ÃÌ ¬—Ì« „Ì »«‘Ìœ" & " »—«Ì  ÂÌÂ ‰”ŒÂ «’·Ì ‰—„ «›“«—»« ‘—ﬂ  «› ÃÌ ¬—Ì« Ì« ‰„«Ì‰œê«‰ ›—Ê‘  „«” »êÌ—Ìœ ", True, False, " «∆Ìœ", ""
            Sleep 1000 * TrialCountFlag
        End If
    End If
    
    
Exit Function
ErrHandler:
    LogSaveNew "frmInvoice => ", err.Description, err.Number, err.Source, "Update"
    Select Case err.Number
        Case 0
        
        Case -2147217873
'            If RepeatUpdate = False Then
'                ShowDisMessage "œ— œÌ « »Ì”  Â„“„«‰Ì ÊÃÊœ œ«‘  òÂ »—ÿ—› ê—œÌœ ", 300
'              '  Sleep 200
'                RepeatUpdate = True
'                GoTo Repeat1
'            'ShowDisMessage "œ— œÌ « »Ì”  Œÿ« ÊÃÊœ œ«—œ " & vbCrLf & "·ÿ›« ¬‰ —« «’·«Õ ‰„ÊœÂ Ê ”Å” À»  ‰„«ÌÌœ", 3000
'            Else
                ShowDisMessage "œ— œÌ « »Ì”  Œÿ« ÊÃÊœ œ«—œ " & vbCrLf, 3000
'            End If
            
        Case Else
    End Select
    
    MsgBox err.Description, vbOKOnly, err.Number
    sFactorReceived = ""
    Update = -1
End Function
Private Function GetCRMCalculate() As Long
    
    Dim SerialNo As Long
    Dim ErrorMsg As String
    Dim Result As Boolean
    With FlxDetail
        clsdiscount.ClearGoodList
        For i = 1 To MaxRowFlexGrid - 1
            Set GoodItem = New AryaCRMDiscountCalculator.GoodView
    
            GoodItem.GoodCode = .TextMatrix(i, 5)
            GoodItem.RowNo = i
            GoodItem.GoodAmount = .TextMatrix(i, 1)
            GoodItem.GoodName = .TextMatrix(i, 2)
            GoodItem.SellPrice = Val(.TextMatrix(i, 3))
            GoodItem.Totalprice = Val(.TextMatrix(i, 4))
            GoodItem.Differences = .TextMatrix(i, 10)
            GoodItem.DifferencesCodes = .TextMatrix(i, 9)
            GoodItem.InventoryNo = .TextMatrix(i, 14)
            GoodItem.ServePlace = .TextMatrix(i, 8)
            GoodItem.DutySale = .TextMatrix(i, 17)
            GoodItem.TaxSale = .TextMatrix(i, 18)
            
            clsdiscount.AddGoodList GoodItem
    
        Next i
    End With
              
    InvInfo.InvoiceNo = txtNo
    InvInfo.DiscountTotal = 0 ' (Val(Me.lblDiscountTotal) - AutoDiscountValue)
    InvInfo.DutyTotal = LblDutyTotal
    InvInfo.TaxTotal = lblTaxTotal
    InvInfo.PackingTotal = lblPackingTotal
    InvInfo.CarryFeeTotal = lblCarryFeeTotal
    InvInfo.sumPrice = lblSumPrice.Tag
    InvInfo.TaxTotal = lblTaxTotal
    InvInfo.DutyTotal = LblDutyTotal
           
    InvInfo.Customer = Me.lblCustomer.Tag
    InvInfo.ServePlace = mvarServePlace
    InvInfo.InvoiceStatus = 2
    InvInfo.Owner = 0
    InvInfo.Incharge = tmpUserNo
    InvInfo.OrderType = mVarOrderType
    InvInfo.StationId = clsArya.StationNo
    InvInfo.ServiceTotal = ServiceRate
    InvInfo.Date = txtDate.Text
    InvInfo.TableNo = cmbTable.ItemData(cmbTable.ListIndex)
    
    InvInfo.User = mvarCurUserNo
    InvInfo.DetailsString = DetailsString1
    InvInfo.GuestNo = Val(TxtGuestNo.Text)
    InvInfo.AccountYear = AccountYear
    
    InvInfo.NvcDescription = IIf(textDescriptionFlag = True, Right(txtDescription.Text, 150), " ")
    InvInfo.FacPayment = Abs(CInt(boolPayment))
    InvInfo.Balance = Abs(CInt(BalancePayment))
    InvInfo.PaymentString = sFactorReceived
        
    ErrorMsg = ""
    Result = clsdiscount.GetDiscountValue(ServiceRate, InvInfo, ErrorMsg)
        
    If Result = False Then
        MsgBox " :  Œÿ« œ— „Õ«”»Â Êœ—Ì«›   Œ›Ì› " & " " & ErrorMsg
        GetCRMCalculate = -1
        Exit Function
    Else
        ShowDisMessage " :„ﬁœ«—   Œ›Ì›  „Õ«”»Â ‘œÂ ›«ò Ê— «Ì‰ „‘ —Ì " & " " & CStr(InvInfo.DiscountTotal) & " —Ì«·  ", 1500
    End If
     
'    If InvInfo.DiscountTotal = 0 Then
'        GetCRMCalculate = 0   ''Without Discount and system continue in normal mode
'        Exit Function
'    End If
    
    SerialNo = 0
    ErrorMsg = ""
    Result = clsdiscount.InsertFactor(CurrentBranch, ErrorMsg, SerialNo)
    
    If Result = False Then
        ShowDisMessage " :  Œÿ« œ— À»  ›«ò Ê— " & " " & ErrorMsg, 1500
        GetCRMCalculate = -1
    Else
      GetCRMCalculate = SerialNo
     '   MsgBox " : À»  ›«ò Ê— " & " " & CStr(SerialNo)
    End If

End Function
Private Sub clsdiscount_CardRecieved(ByVal nvcRFID As String)

    Dim strtemporary As String
    Dim cnn As New ADODB.Connection
    Dim rctmp As New Recordset
    cnn.Open strConnectionString
    strtemporary = "SELECT Code , (Name + ' ' + Family) AS nvcName  FROM dbo.tCust WHERE Branch = " & CurrentBranch & " AND nvcRFID = " & nvcRFID
    rctmp.Open strtemporary, cnn, adOpenDynamic, adLockOptimistic, adCmdText
    If Not (rctmp.EOF = True And rctmp.BOF = True) Then
        IsLoyaltyCustomer = True
        mvarcode = rctmp!Code
        ShowDisMessage "  ò«—  »Â ‰«„ " & rctmp!nvcName & "  ‘‰«”«ÌÌ ‘œ ", 1200
        lblCustomer.Tag = mvarcode
        mvarcode = 0
        mVarOrderType = inPerson
        mvarServePlace = Salon
        If mVarOrderType = ByPhone Then
           If clsStation.Language = Farsi Then
                LblOrder.Caption = " ·›‰Ì"
           Else
                LblOrder.Caption = "By phone"
           End If
        Else
            If clsStation.Language = Farsi Then
               LblOrder.Caption = "Õ÷Ê—Ì"
            Else
                LblOrder.Caption = "Inside"
            End If
        End If
        For i = 0 To cmbServePlace.ListCount - 1
            If mvarServePlace = cmbServePlace.ItemData(i) Then
                cmbServePlace.ListIndex = i
                Exit For
            End If
        Next i
        UpdatelblCustomer
        UpdatelblServePlace
        RefreshLables
    Else
        MsgBox " : ò«—  ‘‰«”«ÌÌ ‰‘œ " & " " & nvcRFID
    End If
    rctmp.Close
    cnn.Close
    Set cnn = Nothing
      
End Sub

Public Sub UpdatelblServePlace()

    ReDim Parameter(2) As Parameter
    
    Parameter(0) = GenerateInputParameter("@CurrentServePlace", adInteger, 4, mvarServePlace)
    Parameter(1) = GenerateInputParameter("@intLangugae", adInteger, 4, clsStation.Language)
    Parameter(2) = GenerateOutputParameter("@Caption", adVarWChar, 50)
    
    lblServePlace.Caption = RunParametricStoredProcedure2String("GetServePlaceCaption", Parameter)
    
    If MyFormAddEditMode = ViewMode Then
        EnableDefaultServiceRate = False
    Else
        If ServiceRate = 0 And EnableDefaultServiceRate = True Then
            If mvarServePlace = EnumServePlace.Salon Then ' mvarServePlace = EnumServePlace.Table Or  “Ì—« „Ì“ Ãœ«ê«‰Â ”—ÊÌ” „Ì êÌ—œ Ê œ— Å«— Ì‘‰ ŒÊœ‘ „ﬁœ«— œÂÌ „Ì ‘Êœ
                ServiceRate = DefaultServicePercent
                RefreshLables
            End If
        Else
            If mvarServePlace = EnumServePlace.Car Or mvarServePlace = EnumServePlace.Delivery Or mvarServePlace = EnumServePlace.Out Then
                ServiceRate = 0
                RefreshLables
            End If
        End If
    End If
    If mvarServePlace = EnumServePlace.Delivery Then CmbPayk.Enabled = True Else CmbPayk.Enabled = False
End Sub
Public Sub Find()
 
''''    If mvarStatus = invoice Then
 
''''        frmInput.Picture1.Visible = True
''''        frmInput.txtInput.Visible = False
''''        frmInput.btnCancel.Visible = True
''''
''''        frmInput.fwlblInput.Caption = "Ã” ÃÊ »— «”«” "
''''        frmInput.OptionLevel(0).Caption = "›Ì‘"
''''        frmInput.OptionLevel(1).Caption = " „Ì“"
''''        frmInput.OptionLevel(2).Caption = " «—”«·Ì"
''''
''''        frmInput.OptionLevel(0).Visible = True
''''        frmInput.OptionLevel(1).Visible = True
''''        frmInput.OptionLevel(2).Visible = True
        
''''        If clsStation.SearchType = 0 Then
''''            frmInput.OptionLevel(0).Value = True
''''        ElseIf clsStation.SearchType = 1 Then
''''            frmInput.OptionLevel(1).Value = True
''''        ElseIf clsStation.SearchType = 2 Then
''''            frmInput.OptionLevel(2).Value = True
''''        End If
        
''''        frmInput.Show vbModal
''''    Else
''''        mvarInput = "0"
''''    End If
''''
''''    If mvarInput = "" Then
''''        Exit Sub
''''    ElseIf mvarInput = "0" Then
''''        If mvarStatus = invoice Then
''''            frmFindFactor.Show vbModal
''''        Else
''''            frmFindOrder.Show vbModal
''''        End If
''''        If mvarcode <> 0 Then
''''            txtNo.Text = mvarcode
''''            mvarcode = 0
''''            MyFormAddEditMode = ViewMode   'view Mode
''''            GetDataDetail
''''            RefreshLables
''''            SetFirstToolbar
''''
''''        Else
''''            Exit Sub
''''
''''        End If
''''    ElseIf mvarInput = "1" Then
''''        frmFindTable.Show vbModal
''''
''''        If mvarcode <> 0 Then
''''            txtNo.Text = mvarcode
''''            mvarcode = 0
''''            MyFormAddEditMode = ViewMode   'view Mode
''''            GetDataDetail
''''            RefreshLables
''''            SetFirstToolbar
''''            MyFormAddEditMode = ViewMode   'view Mode
''''
''''        Else
''''            Exit Sub
''''
''''        End If
''''    ElseIf mvarInput = "2" Then
''''        frmFindDeliveries.Show vbModal
''''
''''        If mvarcode <> 0 Then
''''            txtNo.Text = mvarcode
''''            mvarcode = 0
''''            MyFormAddEditMode = ViewMode   'view Mode
''''            GetDataDetail
''''            RefreshLables
''''            SetFirstToolbar
''''            MyFormAddEditMode = ViewMode   'view Mode
''''
''''        Else
''''            Exit Sub
''''
''''        End If
''''
''''    End If
''''    frmInput.OptionLevel(2).Visible = False
   On Error GoTo err_hh
        
        If ClsFormAccess.frmFindFactor = False Then
            ShowDisMessage "‘„« »Â «Ì‰ «„ò«‰ œ” —”Ì ‰œ«—Ìœ", 1500
            Exit Sub
        End If
        
        If mvarStatus = Invoice Then
'            Load frmFindTableDeliveryFich
            frmFindTableDeliveryFich.Show vbModal
        Else
            frmFindOrder.Show vbModal
        End If
        If mvarcode <> 0 Then
            txtNo.Text = mvarcode
            mvarcode = 0
            MyFormAddEditMode = ViewMode   'view Mode
            GetDataDetail
            RefreshLables
            SetFirstToolBar
            MyFormAddEditMode = ViewMode   'view Mode

        Else
            Exit Sub

        End If
   Exit Sub
   
err_hh:
   MsgBox err.Description
End Sub

Public Sub UndoRedo()
'    Dim DatabaseBranch As Integer
'    ReDim Parameters(0) As Parameter
'
'    Parameters(0) = GenerateOutputParameter("@CurrentBranch", adInteger, 4)
'
'    DatabaseBranch = RunParametricStoredProcedure2String("Get_CurrentBranch", Parameters)
'
'    If CurrentBranch <> DatabaseBranch Then
'        frmDisMsg.lblMessage.Caption = "›Ì‘ ‘⁄»Â œÌê— „—ÃÊ⁄ ‰„Ì ‘Êœ "
'        frmDisMsg.Timer1.Interval = 2000
'        frmDisMsg.Timer1.Enabled = True
'        frmDisMsg.Show vbModal
'        Exit Sub
'    End If

    If FWChkAccount.Value = True Then ShowDisMessage "”‰œ Õ”«»œ«—Ì »—«Ì «Ì‰ ›«ò Ê— ﬁ»·« ’«œ— ‘œÂ Ê ﬁ«»· „—ÃÊ⁄ ‰Ì”  . «“»—ê‘  «“ ›—Ê‘ «” ›«œÂ ò‰Ìœ ", 1500: Exit Sub
    If Me.ChkIsLocked.Value <> 0 Then
        frmDisMsg.lblMessage.Caption = "”‰œ ﬁ›· «”  Ê «„ﬂ«‰ „—ÃÊ⁄ Ê »«“ê—œ«‰Ì ¬‰ ÊÃÊœ ‰œ«—œ"
        frmDisMsg.Timer1.Interval = 2000
        frmDisMsg.Timer1.Enabled = True
        frmDisMsg.Show vbModal
        Exit Sub
    End If
    If RefferedForSomeFich = False Then
            frmAccess.MyFormAddEditMode = RefferedMode
            frmAccess.lblTitle.Caption = "»Ì‘ — «“ «Ì‰ ‰„Ì  Ê«‰Ìœ ›Ì‘  „—ÃÊ⁄ ﬂ‰Ìœ..»—«Ì «œ«„Â —„“ »« œ” —”Ì »«·« »“‰Ìœ"
            frmAccess.Show vbModal
            If frmAccess.ReturnAccess = False Then
                Exit Sub
            End If
    End If

    If clsStation.CashClose = True And ClsFormAccess.EditInvoiceCashClose = False Then
           frmAccess.AccessStatus = CashClose
           frmAccess.Show vbModal
           If frmAccess.ReturnAccess = False Then
                Exit Sub
           End If
            clsStation.CashClose = False
     End If

    If clsStation.CashClose = True Then
        frmDisMsg.lblMessage.Caption = "’‰œÊﬁ »” Â «”  Ê «„ﬂ«‰ „—ÃÊ⁄ Ê »«“ê—œ«‰Ì ›Ì‘ ÊÃÊœ ‰œ«—œ"
        frmDisMsg.Timer1.Interval = 2000
        frmDisMsg.Timer1.Enabled = True
        frmDisMsg.Show vbModal
        Exit Sub
    End If
    If (MyFormAddEditMode = ViewMode And clsStation.UndoRedoCompatibleSamar1 = True) Then
            Find
            If Me.FindFlag = False Then Exit Sub
    ElseIf MyFormAddEditMode = AddMode And clsStation.UndoRedoCompatibleSamar1 = True And MaxRowFlexGrid > 1 Then
            frmDisMsg.lblMessage = " ›Ì‘ À»  ‰‘œÂ ﬁ«»· „—ÃÊ⁄ ﬂ—œ‰ ‰Ì”  "
            frmDisMsg.Timer1.Enabled = True
            frmDisMsg.Show vbModal
            Exit Sub
    ElseIf MyFormAddEditMode = AddMode And clsStation.UndoRedoCompatibleSamar1 = True And MaxRowFlexGrid = 1 Then
            Find
            If Me.FindFlag = False Then Exit Sub
    ElseIf (MyFormAddEditMode <> ViewMode And MyFormAddEditMode <> RefferedMode) And clsStation.UndoRedoCompatibleSamar1 = False Then
            frmDisMsg.lblMessage = " ›Ì‘ À»  ‰‘œÂ ﬁ«»· „—ÃÊ⁄ ﬂ—œ‰ ‰Ì”  "
            frmDisMsg.Timer1.Enabled = True
            frmDisMsg.Show vbModal
            Exit Sub
    End If

If mvarCurUserNo = 0 Then

    Set Rst = Nothing
    frmDisMsg.lblMessage = "‘„« «Ã«“Â „—ÃÊ⁄ ò—œ‰ ›Ì‘ —« ‰œ«—Ìœ"
    frmDisMsg.Timer1.Enabled = True
    frmDisMsg.Show vbModal
    Exit Sub
End If

If (ClsFormAccess.RefferInvoice = False) Or (ClsFormAccess.RefferedAllStationsFactors = False And (mvarCurUserNo <> dblFichUser)) Then

    Set Rst = Nothing
    frmAccess.MyFormAddEditMode = RefferedMode
    frmAccess.lblTitle.Caption = "‘„« «Ã«“Â „—ÃÊ⁄ ò—œ‰ ›Ì‘ —« ‰œ«—Ìœ..»—«Ì «œ«„Â —„“ »« œ” —”Ì »«·« »“‰Ìœ"
    frmAccess.Show vbModal
    If frmAccess.ReturnAccess = False Then
        Exit Sub
    End If
    
End If

    Select Case Val(txtRecursive.Text)
    
        Case 0

            ReDim Parameter(0) As Parameter
            Parameter(0) = GenerateInputParameter("@intSerialNo", adBigInt, 8, intSerialNo)
            Set rctmp = RunParametricStoredProcedure2Rec("Get_RowCount_FactorDetail", Parameter)
            
            If rctmp.Fields("NoOfRows") = 0 Or IsNull(rctmp.Fields("NoOfRows")) Then
                frmMsg.fwlblMsg.Caption = " . ›Ì‘ Œ«·Ì „—ÃÊ⁄ ‰„Ì ê—œœ"
                frmMsg.fwBtn(0).Visible = False
                frmMsg.fwBtn(1).ButtonType = flwButtonOk
                frmMsg.fwBtn(1).Caption = "ﬁ»Ê·"
                frmMsg.Show vbModal
                Exit Sub
            End If
            
            
            ReDim Parameter(3) As Parameter
            
            Parameter(0) = GenerateInputParameter("@No", adBigInt, 8, Val(txtNo.Text))
            Parameter(1) = GenerateInputParameter("@Status", adInteger, 4, mvarStatus)
            Parameter(2) = GenerateInputParameter("@AccountYear", adSmallInt, 2, AccountYear)
            Parameter(3) = GenerateInputParameter("@Branch", adInteger, 4, CurrentBranch)
            
            Set Rst = RunParametricStoredProcedure2Rec("Get_FacM_Per", Parameter)
            
            If Not (Rst.EOF = True And Rst.BOF = True) Then
            
                If Rst.Fields("ServePlace").Value = 2 And Rst.Fields("Incharge").Value <> 0 Then
                    frmMsg.fwlblMsg.Caption = " ›Ì‘ «—”«· ‘œÂ „—ÃÊ⁄ ‰„Ì ‘Êœ " & vbLf & " »«Ìœ «» œ« ¬‰ —« «“ Õ”«» ÅÌﬂ Œ«—Ã ﬂ‰Ìœ "
                    frmMsg.fwBtn(0).Visible = False
                    frmMsg.fwBtn(1).ButtonType = flwButtonOk
                    frmMsg.fwBtn(1).Caption = "ﬁ»Ê·"
                    frmMsg.Show vbModal
                    Exit Sub
                End If
            
            End If
            If clsStation.Language = Farsi Then
                frmMsg.fwlblMsg.Caption = "¬Ì« „ÿ„∆‰Ìœ „Ì ŒÊ«ÂÌœ ›Ì‘ —« „—ÃÊ⁄ ﬂ‰Ìœø "
            Else
                frmMsg.fwlblMsg.Caption = "Do you want Refund this invoice ? "
            End If
            frmMsg.Show vbModal
            If modgl.mvarMsgIdx = vbYes Then
               MyFormAddEditMode = RefferedMode
               txtRecursive.Text = 1
              '  Edit
                Printing
                
            End If
            
        Case 1
        
            If clsStation.Language = Farsi Then
                frmMsg.fwlblMsg.Caption = "¬Ì« „ÿ„∆‰Ìœ „Ì ŒÊ«ÂÌœ ›Ì‘ „—ÃÊ⁄ ‘œÂ —« »—ê—œ«‰Ìœø "
            Else
                frmMsg.fwlblMsg.Caption = "Do you want undo to Refund this invoice ? "
            End If
            frmMsg.Show vbModal
            If modgl.mvarMsgIdx = vbYes Then
                fwlblRecursive.Visible = False
              '  fwScrollTextCust.Visible = True
                MyFormAddEditMode = RefferedMode
                txtRecursive.Text = 0
                Printing
                Add
            End If
    End Select
        
End Sub
Public Sub ExitForm()
    If clsArya.ExternalAccounting = True And ClsFormAccess.frmCreateSanad = True Then
        ReDim Parameter(3) As Parameter
        Parameter(0) = GenerateInputParameter("@Branch", adInteger, 4, CurrentBranch)
        Parameter(1) = GenerateInputParameter("@DateBefore", adVarWChar, 8, txtDate)
        Parameter(2) = GenerateInputParameter("@DateAfter", adVarWChar, 8, txtDate)
        Parameter(3) = GenerateInputParameter("@Uid", adInteger, 4, 0)
        Set Rst = RunParametricStoredProcedure2Rec("Get_SaleSummary", Parameter)
    '''''''''''   ›—Ê‘
        If Rst.EOF <> True And Rst.BOF <> True Then
            ShowMessage " Ê·Ìœ ”‰œ Õ”«»œ«—Ì »—«Ì ›«ò Ê—Â«Ì „ ›—ﬁÂ «‰Ã«„ ‘Êœø ", True, True, "»·Ì", "ŒÌ—"
            If mvarMsgIdx = vbYes Then
                Unload frmCreateSanad
                frmCreateSanad.cmbUsers.Enabled = False
                Load frmCreateSanad
                If frmCreateSanad.cmbUsers.ListCount > 0 Then
                    For i = 0 To frmCreateSanad.cmbUsers.ListCount - 1
                        If frmCreateSanad.cmbUsers.ItemData(i) = mvarCurUserNo Then
                             frmCreateSanad.cmbUsers.ListIndex = i
                             Exit For
                        End If
                    Next
                End If
                frmCreateSanad.Show
            End If
            Unload Me
            Exit Sub
        End If
    End If
        
'''  »—«Ì Õ«·  »œÊ‰ Õ”«»œ«—Ì
        frmMsg.fwlblMsg.Caption = "¬Ì« »—«Ì Œ—ÊÃ «“ ›«ﬂ Ê—›—Ê‘ «ÿ„Ì‰«‰ œ«—Ìœ"
        frmMsg.fwBtn(0).ButtonType = flwButtonOk
        frmMsg.fwBtn(1).ButtonType = flwButtonCancel
        frmMsg.fwBtn(0).Caption = "»·Ì"
        frmMsg.fwBtn(1).Caption = "ŒÌ—"
        frmMsg.Show vbModal
        If mvarMsgIdx = vbYes Then
           USBHID1.PortOpen = False
'            If clsStation.PosPayment = True Then
'                frmMsg.fwlblMsg.Caption = "¬Ì« „ÌŒÊ«ÂÌœ  «ÌÌœÌÂ  —«ﬂ‰‘ Â«Ì «‰Ã«„ ‘œÂ »Â „—ﬂ“ «—”«· ê—œœø"
'                frmMsg.fwBtn(0).ButtonType = flwButtonOk
'                frmMsg.fwBtn(1).ButtonType = flwButtonCancel
'                frmMsg.fwBtn(0).Caption = "»·Ì"
'                frmMsg.fwBtn(1).Caption = "ŒÌ—"
'                frmMsg.Show vbModal
'
'                If mvarMsgIdx = vbYes Then
'                    IsClosingInvoiceForm = True
'                    SendSettelmentMessageToPos "90"
'                    frmDisMsg.lblMessage.Caption = "«ÿ·«⁄«  »Â Å«Ì«‰Â ›—Ê‘ «—”«· ê—œÌœ. œ— Õ«· «‰ Ÿ«— »—«Ì Å«”Œ..."
'                    frmDisMsg.Timer1.Interval = 2000
'                    frmDisMsg.Timer1.Enabled = True
'                    frmDisMsg.Show vbModal
'                Else
'                    Unload Me
'                End If
'            Else
               Unload Me
'            End If
        
        End If
End Sub
Private Sub FWBtnGarsoon_Click()
    On Error Resume Next
    DropDownFlag = True
    cmbGarson.SetFocus
    SendKeys "{F4}", True
'    cmbGarson_KeyUp 0, 0
    DropDownFlag = False
End Sub

Private Sub FWBtnPayk_Click()
    If BlnPosApprovedWait = True Then
        ShowDisMessage "”Ì” „ „‰ Ÿ— œ—Ì«›  Å«”Œ «“ ŒÊœÅ—œ«“ „Ì »«‘œ", 1000
        Exit Sub
    End If
    If ClsFormAccess.frmPayk = True And clsArya.Delivery = True Then
        frmPayk.Show
    Else
        frmDisMsg.lblMessage.Caption = " ‘„« »Â «Ì‰ «„ﬂ«‰ œ” —”Ì ‰œ«—Ìœ "
        frmDisMsg.Timer1.Interval = 1000
        frmDisMsg.Timer1.Enabled = True
        frmDisMsg.Show vbModal
    
    End If
End Sub

Private Sub FWBtnSplit_Click()
If MaxRowFlexGrid = 1 Then Exit Sub
Dim ii As Integer
If SplitFlag = False Then
    FWBtnSplit.Caption = "„⁄„Ê·Ì"
''    FWBtnSplit.Caption = "„Õ«”»Â"
    FWBtnSplit.ForeColor = vbGreen
    SplitFlag = True
    DetailsString1 = ""
    With FlxDetail
        For i = 1 To MaxRowFlexGrid - 1
            DetailsString1 = GenerateDetailsString3(DetailsString1, .TextMatrix(i, 1), .TextMatrix(i, 5), Val(.TextMatrix(i, 3)), Val(.TextMatrix(i, 4)) * Val(.TextMatrix(i, 11)) / 100, Val(.TextMatrix(i, 12)), .TextMatrix(i, 13), " ", .TextMatrix(i, 14), "", .TextMatrix(i, 8), .TextMatrix(i, 9))
        Next i
    End With
    
    ClearDataFlexGrid
    
    ReDim Parameter(4) As Parameter
    
    Parameter(0) = GenerateInputParameter("@DetailsString", adVarWChar, 4000, DetailsString1)
    If (Val(lblCustomer.Tag) > -1) Then
        Parameter(1) = GenerateInputParameter("@Customer", adInteger, 4, Val(lblCustomer.Tag))
    Else
        Parameter(1) = GenerateInputParameter("@Customer", adInteger, 4, -1)
    End If
    Parameter(2) = GenerateInputParameter("@intlanguage", adInteger, 4, clsStation.Language)
    Parameter(3) = GenerateInputParameter("@Split", adBoolean, 1, 1)
    Parameter(4) = GenerateInputParameter("@CustomerSoftwareCode", adInteger, 4, clsArya.CustomerId)
    
    Set Rst = RunParametricStoredProcedure2Rec("Get_Split_Special", Parameter)
    
    
    If Not (Rst.BOF Or Rst.EOF) Then
        TmpGoodDiscount = 0
        Do While Not (Rst.EOF)
            ii = ii + 1
            FlxDetail.TextMatrix(ii, 0) = ii 'Number
            FlxDetail.TextMatrix(ii, 1) = Rst!amount
            FlxDetail.TextMatrix(ii, 2) = Rst!Name 'GoodName
            FlxDetail.TextMatrix(ii, 3) = Rst!FeeUnit
            FlxDetail.TextMatrix(ii, 4) = Rst!amount * Rst!FeeUnit ' rst!FeeTotal
            FlxDetail.TextMatrix(ii, 5) = Rst!GoodCode
            FlxDetail.TextMatrix(ii, 6) = Rst!Weight ' rst!WeightUnit
            FlxDetail.TextMatrix(ii, 7) = Rst!Unit
            FlxDetail.TextMatrix(ii, 8) = Rst!ServePlace
            FlxDetail.TextMatrix(ii, 9) = IIf(IsNull(Rst!DifferencesCode), "", Rst!DifferencesCode)
            FlxDetail.TextMatrix(ii, 10) = IIf(IsNull(Rst!DifferencesDescription), "", Rst!DifferencesDescription)
            If Rst!FeeUnit <> 0 Then
                FlxDetail.TextMatrix(ii, 11) = Rst!Discount * 100 / (Rst!amount * Rst!FeeUnit)
            End If
            FlxDetail.TextMatrix(ii, 12) = Rst!Rate
            FlxDetail.TextMatrix(ii, 13) = IIf(IsNull(Rst!ChairName), "", Rst!ChairName)
            FlxDetail.TextMatrix(ii, 14) = Rst!intInventoryNo
            FlxDetail.TextMatrix(ii, 15) = Rst!mainType
            TmpGoodDiscount = TmpGoodDiscount + Rst!Discount
            
            Rst.MoveNext
            
            MaxRowFlexGrid = ii + 1
            If ii >= FlxDetail.Rows - 2 And Rst.EOF = False Then
                AddEmptyRow
            End If

        Loop
        
        FlxDetail.Row = MaxRowFlexGrid - 1
        mvarServePlace = FlxDetail.TextMatrix(MaxRowFlexGrid - 1, 8)
        
    End If
    If Rst.State <> 0 Then Rst.Close
    
Else
    FWBtnSplit.Caption = "„Õ«”»Â"
'''    FWBtnSplit.Caption = "„⁄„Ê·Ì"
    FWBtnSplit.ForeColor = vbRed
    SplitFlag = False
    DetailsString1 = ""
    With FlxDetail
        For i = 1 To MaxRowFlexGrid - 1
            DetailsString1 = GenerateDetailsString3(DetailsString1, .TextMatrix(i, 1), .TextMatrix(i, 5), Val(.TextMatrix(i, 3)), Val(.TextMatrix(i, 4)) * Val(.TextMatrix(i, 11)) / 100, Val(.TextMatrix(i, 12)), .TextMatrix(i, 13), " ", .TextMatrix(i, 14), "", .TextMatrix(i, 8), .TextMatrix(i, 9))
        Next i
    End With
    
    ClearDataFlexGrid
    
    ReDim Parameter(4) As Parameter
    
    Parameter(0) = GenerateInputParameter("@DetailsString", adVarWChar, 4000, DetailsString1)
    If (Val(lblCustomer.Tag) > -1) Then
        Parameter(1) = GenerateInputParameter("@Customer", adInteger, 4, Val(lblCustomer.Tag))
    Else
        Parameter(1) = GenerateInputParameter("@Customer", adInteger, 4, -1)
    End If
    Parameter(2) = GenerateInputParameter("@intlanguage", adInteger, 4, clsStation.Language)
    Parameter(3) = GenerateInputParameter("@Split", adBoolean, 1, 0)
    Parameter(4) = GenerateInputParameter("@CustomerSoftwareCode", adInteger, 4, clsArya.CustomerId)
    
    Set Rst = RunParametricStoredProcedure2Rec("Get_Split_Special", Parameter)
    
    If Not (Rst.BOF Or Rst.EOF) Then
        TmpGoodDiscount = 0
        Do While Not (Rst.EOF)
            ii = ii + 1
            FlxDetail.TextMatrix(ii, 0) = ii 'Number
            FlxDetail.TextMatrix(ii, 1) = Rst!amount
            FlxDetail.TextMatrix(ii, 2) = Rst!Name 'GoodName
            FlxDetail.TextMatrix(ii, 3) = Rst!FeeUnit
            FlxDetail.TextMatrix(ii, 4) = Rst!amount * Rst!FeeUnit ' rst!FeeTotal
            FlxDetail.TextMatrix(ii, 5) = Rst!GoodCode
            FlxDetail.TextMatrix(ii, 6) = Rst!Weight ' rst!WeightUnit
            FlxDetail.TextMatrix(ii, 7) = Rst!Unit
            FlxDetail.TextMatrix(ii, 8) = Rst!ServePlace
            FlxDetail.TextMatrix(ii, 9) = IIf(IsNull(Rst!DifferencesCode), "", Rst!DifferencesCode)
            FlxDetail.TextMatrix(ii, 10) = IIf(IsNull(Rst!DifferencesDescription), "", Rst!DifferencesDescription)
            If Rst!FeeUnit <> 0 Then
                FlxDetail.TextMatrix(ii, 11) = Rst!Discount * 100 / (Rst!amount * Rst!FeeUnit)
            End If
            FlxDetail.TextMatrix(ii, 12) = Rst!Rate
            FlxDetail.TextMatrix(ii, 13) = "" 'IIf(IsNull(Rst!ChairName), "", Rst!ChairName)
            FlxDetail.TextMatrix(ii, 14) = Rst!intInventoryNo
            FlxDetail.TextMatrix(ii, 15) = Rst!mainType
            TmpGoodDiscount = TmpGoodDiscount + Rst!Discount
            
            Rst.MoveNext
            
            MaxRowFlexGrid = ii + 1
            If ii >= FlxDetail.Rows - 2 And Rst.EOF = False Then
                AddEmptyRow
            End If

        Loop
        
        FlxDetail.Row = MaxRowFlexGrid - 1
        mvarServePlace = FlxDetail.TextMatrix(MaxRowFlexGrid - 1, 8)
        
    End If
    If Rst.State <> 0 Then Rst.Close
End If
RefreshLables
End Sub

Private Sub FWBtnTable_Click()
    DropDownFlag = True
    
    cmbTable.SetFocus
    cmbTable.ListIndex = 0
    SendKeys "{F4}", True
    DropDownFlag = False
End Sub

Public Function ColorSetting()
On Error Resume Next
    
    Call SetColor
    If Invoice_FontDifferencesName <> "" Then
        lstDifference.Font.Name = Invoice_FontDifferencesName
        lstDifference.Font.size = Invoice_FontDifferencesSize
        lstDifference.Font.Bold = Invoice_FontDifferencesBold
    Else
        lstDifference.Font.size = "11"
        lstDifference.Font.Bold = "True"
        Invoice_FontDifferencesSize = "11"
        Invoice_FontDifferencesBold = "True"
    End If
    FWChkAccount.BackColor = Invoice_BackColorForm
    Me.BackColor = Invoice_BackColorForm
    frameMenu.BackColor = Invoice_BackColorForm
    Frame1.BackColor = Invoice_BackColorForm
    Frame6.BackColor = Invoice_BackColorForm
    Picture1.BackColor = Invoice_BackColorForm
    Picture3.BackColor = Invoice_BackColorForm
    ChkCallerId.BackColor = Invoice_BackColorForm
    txtDescription.BackColor = Invoice_BackColorForm
    Shape3.FillColor = Invoice_BackColorForm
    
'    Frame2.BackColor = Invoice_BackColorForm
'    Frame3.BackColor = Invoice_BackColorForm
    Frame_CallerId.BackColor = Invoice_BackColorForm
'    Frame5.BackColor = Invoice_BackColorForm
    FWLed1.BackColor = Invoice_BackColorForm
    FWLed1.ColorOff = Invoice_BackColorForm
    FWLedTemp.BackColor = Invoice_BackColorForm
    FWLedTemp.ColorOff = Invoice_BackColorForm
'    cmbTable.BackColor = Invoice_BackColorForm
'    cmbGarson.BackColor = Invoice_BackColorForm
    
    'KeyPadMenu.BackColor = Invoice_BackColorForm
    FlxDetail.BackColor = Invoice_BackColorFlexGrid
    txtDate.BackColor = Invoice_BackColorForm

    FWChkHavale.BackColor = Invoice_BackColorForm
'    txtDescription.BackColor = Invoice_BackColorForm
'    FWChkHavale.BackColor = Invoice_BackColorForm
'    ChkCallerId.BackColor = Invoice_BackColorForm
    
    cmbServePlace.BackColor = Shape3.FillColor
    
End Function

Private Sub FWMojodiControl_Click()
    If ClsFormAccess.MojodiControl = False Then
       FWMojodiControl.Enabled = False
       frmMsg.fwlblMsg.Caption = ". ‘„« »Â «Ì‰ «„ﬂ«‰ œ” —”Ì ‰œ«—Ìœ"
       frmMsg.fwBtn(1).Visible = False
       frmMsg.Show vbModal
    Else
        If MojodiControlFlag = False Then
            MojodiControlFlag = True
            FWMojodiControl.ButtonType = flwButtonOk
            FWMojodiControl.ForeColor = &H4000&     'vbGreen
        Else
            MojodiControlFlag = False
            FWMojodiControl.ButtonType = flwButtonDelete
            FWMojodiControl.ForeColor = vbRed
        End If
   End If
End Sub

Private Sub FWMojodiControl_GotFocus()
    On Error Resume Next
    FlxDetail.SetFocus
End Sub

Private Sub LblInvoice_Click()
        
    If intVersion = Min Or intVersion = Normal Then
        ShowDisMessage "«„ﬂ«‰ «” ›«œÂ «“ ”›«—‘«  Ê ÷«Ì⁄«  œ— ‰”ŒÂ Â«Ì ÅÌ‘—› Â Ê »«·« — ÊÃÊœ œ«—œ", 1500
        Exit Sub
    End If
    If ClsFormAccess.frmOrder = True Or ClsFormAccess.frmLosses = True Then
        framelastFich.Visible = False
        frmMsg.fwlblMsg.Caption = "Õ«·  ”‰œ Ã«—Ì  €ÌÌ— „Ì ﬂ‰œ "
        frmMsg.fwBtn(0).ButtonType = flwButtonOk
        frmMsg.fwBtn(0).Caption = " «∆Ìœ"
        frmMsg.fwBtn(1).Visible = False
        frmMsg.Show vbModal
        If mvarStatus = Invoice And ClsFormAccess.frmOrder = True Then
            mvarStatus = EnumFactorType.Order
            If clsStation.Language = Farsi Then
                LblInvoice.Caption = "”›«—‘"
                cmdPay.Caption = "(F8) ÕÊÌ·"
            Else
                LblInvoice.Caption = "Order"
                cmdPay.Caption = "Recieve(F8)"
            End If
            Add
        ElseIf (mvarStatus = Invoice Or mvarStatus = Order) And ClsFormAccess.SaleReturn = True Then
            mvarStatus = EnumFactorType.InvoiceReturn
            LblInvoice.Caption = "»—ê‘  «“ ›—Ê‘"
            Add
        ElseIf ((mvarStatus = Invoice Or mvarStatus = InvoiceReturn Or mvarStatus = EnumFactorType.Order) And ClsFormAccess.frmLosses = True) Then
            mvarStatus = Losses
            If clsStation.Language = Farsi Then
                LblInvoice.Caption = "÷«Ì⁄« "
                cmdPay.Caption = "(F8)œ—Ì«› "
            Else
                LblInvoice.Caption = "Losses"
                cmdPay.Caption = "Recieve(F8)"
           End If
            Add
        ElseIf mvarStatus = EnumFactorType.Losses Then
            mvarStatus = Invoice
            If clsStation.Language = Farsi Then
                LblInvoice.Caption = "›«ﬂ Ê— ›—Ê‘"
                cmdPay.Caption = "(F8)œ—Ì«› "
            Else
                LblInvoice.Caption = "Invoice"
                cmdPay.Caption = "Recieve(F8)"
           End If
            Add
        End If
    Else
        ShowDisMessage "‘„« »Â «Ì‰ ﬁ«»·Ì  œ” —”Ì ‰œ«—Ìœ", 1500
        Exit Sub
    End If
    
    If clsStation.PayFactorView = True Or mvarStatus = Order Then
        cmdPayFactor.Visible = True
        lblPayFactorTotal.Visible = True
    Else
        cmdPayFactor.Visible = False
        lblPayFactorTotal.Visible = False
    End If

End Sub

Private Sub LblRate_Click()
    
    If clsStation.ShiftRate = True Then Exit Sub

    If ClsFormAccess.MultiPrice = True Then
        frmMsg.fwlblMsg.Caption = "¬Ì« „«Ì·Ìœ ‰—Œ ﬂ«·«Â« —«  €ÌÌ— œÂÌœø "
        frmMsg.fwBtn(0).ButtonType = flwButtonOk
        frmMsg.fwBtn(0).Caption = "»·Ì"
        frmMsg.fwBtn(1).ButtonType = flwButtonCancel
        frmMsg.fwBtn(1).Caption = "ŒÌ—"
        frmMsg.fwBtn(1).Default = True
        frmMsg.Show vbModal
        
        If mvarMsgIdx = vbYes Then
            If clsStation.PriceType = 6 Then
               clsStation.PriceType = 1
               mvarStartRate = 1
               LblRate.Caption = "‰—Œ «Ê·"
            ElseIf clsStation.PriceType = 1 Then
               If clsStation.MaxPrices > 1 Then
                    clsStation.PriceType = 2
                    mvarStartRate = 2
                    LblRate.Caption = "‰—Œ œÊ„"
               End If
            ElseIf clsStation.PriceType = 2 Then
               If clsStation.MaxPrices > 2 Then
                    clsStation.PriceType = 3
                    mvarStartRate = 3
                    LblRate.Caption = "‰—Œ ”Ê„"
               Else
                    clsStation.PriceType = 1
                    LblRate.Caption = "‰—Œ «Ê·"
                    mvarStartRate = 1
               End If
            ElseIf clsStation.PriceType = 3 Then
               If clsStation.MaxPrices > 3 Then
                    clsStation.PriceType = 4
                    mvarStartRate = 4
                    LblRate.Caption = "‰—Œ çÂ«—„"
               Else
                    clsStation.PriceType = 1
                    LblRate.Caption = "‰—Œ «Ê·"
                    mvarStartRate = 1
               End If
            ElseIf clsStation.PriceType = 4 Then
               If clsStation.MaxPrices > 4 Then
                    clsStation.PriceType = 5
                    mvarStartRate = 5
                    LblRate.Caption = "‰—Œ Å‰Ã„"
               Else
                    clsStation.PriceType = 1
                    LblRate.Caption = "‰—Œ «Ê·"
                    mvarStartRate = 1
               End If
            ElseIf clsStation.PriceType = 5 Then
               If clsStation.MaxPrices > 5 Then
                    clsStation.PriceType = 6
                    mvarStartRate = 6
                    LblRate.Caption = "‰—Œ ‘‘„"
               Else
                    clsStation.PriceType = 1
                    mvarStartRate = 1
                    LblRate.Caption = "‰—Œ «Ê·"
                    mvarStartRate = 1
               End If
            End If
            frmDisMsg.lblMessage.Caption = " ‰—Œ ﬂ«·«Â« »Â ‰—Œ " & clsStation.PriceType & "  €ÌÌ— Ì«›  "
            frmDisMsg.Timer1.Interval = 1000
            frmDisMsg.Timer1.Enabled = True
            frmDisMsg.Show vbModal
                    
            ServeChangeFlag = True
            RateChanged False
        End If
     Else
        frmDisMsg.lblMessage.Caption = " ‘„« »Â «Ì‰ ﬁ«»·Ì   œ” —”Ì ‰œ«—Ìœ "
        frmDisMsg.Timer1.Interval = 1000
        frmDisMsg.Timer1.Enabled = True
        frmDisMsg.Show vbModal
     End If
End Sub

Private Sub RateChanged(CustomerRate As Boolean)
    
    Dim cnn As New ADODB.Connection
    Dim rctmp As New Recordset
    cnn.Open strConnectionString
    Dim strtemporary As String
    With FlxDetail
        For i = 1 To MaxRowFlexGrid - 1
            mvarGoodCode = Val(.TextMatrix(i, 5))
            strtemporary = "SELECT * FROM dbo.tGood WHERE Code = " & mvarGoodCode
            rctmp.Open strtemporary, cnn, adOpenDynamic, adLockOptimistic, adCmdText
            If Not (rctmp.EOF = True And rctmp.BOF = True) Then
                If CustomerRate = False Then
                    If clsStation.PriceType = 1 Then
                       .TextMatrix(i, 3) = rctmp.Fields("SellPrice").Value
                       .TextMatrix(i, 12) = 1
                    ElseIf clsStation.PriceType = 2 Then
                       .TextMatrix(i, 3) = rctmp.Fields("SellPrice2").Value
                       .TextMatrix(i, 12) = 2
                    ElseIf clsStation.PriceType = 3 Then
                       .TextMatrix(i, 3) = rctmp.Fields("SellPrice3").Value
                       .TextMatrix(i, 12) = 3
                    ElseIf clsStation.PriceType = 4 Then
                       .TextMatrix(i, 3) = rctmp.Fields("SellPrice4").Value
                       .TextMatrix(i, 12) = 4
                    ElseIf clsStation.PriceType = 5 Then
                       .TextMatrix(i, 3) = rctmp.Fields("SellPrice5").Value
                       .TextMatrix(i, 12) = 5
                    ElseIf clsStation.PriceType = 6 Then
                       .TextMatrix(i, 3) = rctmp.Fields("SellPrice6").Value
                       .TextMatrix(i, 12) = 6
                    End If
                Else
                    If clsStation.CustomerRate = 0 Then
                       .TextMatrix(i, 3) = rctmp.Fields("SellPrice").Value
                       .TextMatrix(i, 12) = 1
                    ElseIf clsStation.CustomerRate = 1 Then
                       .TextMatrix(i, 3) = rctmp.Fields("SellPrice2").Value
                       .TextMatrix(i, 12) = 2
                    ElseIf clsStation.CustomerRate = 2 Then
                       .TextMatrix(i, 3) = rctmp.Fields("SellPrice3").Value
                       .TextMatrix(i, 12) = 3
                    ElseIf clsStation.CustomerRate = 3 Then
                       .TextMatrix(i, 3) = rctmp.Fields("SellPrice4").Value
                       .TextMatrix(i, 12) = 4
                    ElseIf clsStation.CustomerRate = 4 Then
                       .TextMatrix(i, 3) = rctmp.Fields("SellPrice5").Value
                       .TextMatrix(i, 12) = 5
                    ElseIf clsStation.CustomerRate = 5 Then
                       .TextMatrix(i, 3) = rctmp.Fields("SellPrice6").Value
                       .TextMatrix(i, 12) = 6
                    End If
                
                End If
            End If
            rctmp.Close
                    
        Next
        cnn.Close
        Set cnn = Nothing
        RefreshLables
    End With

End Sub
Private Sub LblOrder_Click()
    
    If mVarOrderType = inPerson Then
        mVarOrderType = ByPhone
        If clsStation.Language = Farsi Then
            LblOrder.Caption = " ·›‰Ì"
        Else
            LblOrder.Caption = "By phone"
        End If
    Else
        mVarOrderType = inPerson
        If clsStation.Language = Farsi Then
            LblOrder.Caption = "Õ÷Ê—Ì"
        Else
            LblOrder.Caption = "Inside"
        End If
    End If

End Sub

Private Sub lblScale_Click(index As Integer)
    lblScale(index).RightToLeft = IIf(lblScale(index).RightToLeft, False, True)

End Sub

Private Sub lblServePlace_Click()
 ' If clsArya.Restaurant = True Then
        ReDim Parameter(0) As Parameter
        
        Parameter(0) = GenerateInputParameter("@CurrentServePlace", adInteger, 4, mvarServePlace)
        
        Set Rst = RunParametricStoredProcedure2Rec("GetValidServePlace", Parameter)
        
        If Not (Rst.EOF = True And Rst.BOF = True) Then
            mvarServePlace = Rst.Fields("intServeplace").Value
        Else
            Set Rst = RunStoredProcedure2RecordSet("GetFirstValidServePlace")
            mvarServePlace = Rst.Fields("intServeplace").Value
        End If
        If mvarServePlace = Out Then
            clsStation.PriceType = clsStation.OutPrice
        Else
            clsStation.PriceType = MainPriceType
        End If
        
        If clsStation.PriceType = 1 Then
            LblRate.Caption = "‰—Œ «Ê·"
        ElseIf clsStation.PriceType = 2 Then
            LblRate.Caption = "‰—Œ œÊ„"
        ElseIf clsStation.PriceType = 3 Then
            LblRate.Caption = "‰—Œ ”Ê„"
        ElseIf clsStation.PriceType = 4 Then
            LblRate.Caption = "‰—Œ çÂ«—„"
        ElseIf clsStation.PriceType = 5 Then
            LblRate.Caption = "‰—Œ Å‰Ã„"
        ElseIf clsStation.PriceType = 6 Then
            LblRate.Caption = "‰—Œ ‘‘„"
        End If
        
        UpdatelblServePlace
        RefreshLables
 '  End If

End Sub

'Private Sub lblServePlace_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    lblServePlace_Click
'End Sub

Private Sub LblSubTotal_Click()
    LblSubTotal.RightToLeft = IIf(LblSubTotal.RightToLeft, False, True)
End Sub

Private Sub lblSumPrice_Click()
    lblSumPrice.RightToLeft = IIf(lblSumPrice.RightToLeft, False, True)
End Sub
Public Sub SplitByOptions(CurrentRow As Long)
    If IsNumeric(lblNum.Caption) Then
        If Val(lblNum.Caption) < Val(FlxDetail.TextMatrix(CurrentRow, 1)) And lblNum.Caption <> 0 Then
            With FlxDetail
                .TextMatrix(MaxRowFlexGrid, 0) = .TextMatrix(CurrentRow, 0)
                .TextMatrix(MaxRowFlexGrid, 1) = .TextMatrix(CurrentRow, 1) - Val(lblNum.Caption)
                .TextMatrix(MaxRowFlexGrid, 2) = .TextMatrix(CurrentRow, 2)
                .TextMatrix(MaxRowFlexGrid, 3) = .TextMatrix(CurrentRow, 3)
                .TextMatrix(MaxRowFlexGrid, 4) = Val(.TextMatrix(MaxRowFlexGrid, 3)) * Val(.TextMatrix(MaxRowFlexGrid, 1))
                .TextMatrix(MaxRowFlexGrid, 5) = .TextMatrix(CurrentRow, 5)
                .TextMatrix(MaxRowFlexGrid, 6) = .TextMatrix(CurrentRow, 6)
                .TextMatrix(MaxRowFlexGrid, 7) = .TextMatrix(CurrentRow, 7)
                .TextMatrix(MaxRowFlexGrid, 8) = .TextMatrix(CurrentRow, 8)
                .TextMatrix(MaxRowFlexGrid, 9) = .TextMatrix(CurrentRow, 9)
                .TextMatrix(MaxRowFlexGrid, 10) = .TextMatrix(CurrentRow, 10)
                .TextMatrix(MaxRowFlexGrid, 12) = mvarRate
                .TextMatrix(MaxRowFlexGrid, 14) = .TextMatrix(CurrentRow, 14)
                .TextMatrix(MaxRowFlexGrid, 15) = .TextMatrix(CurrentRow, 15)
                .TextMatrix(CurrentRow, 1) = Val(lblNum.Caption)
                .TextMatrix(CurrentRow, 4) = Val(.TextMatrix(CurrentRow, 3)) * Val(.TextMatrix(CurrentRow, 1))
                MaxRowFlexGrid = MaxRowFlexGrid + 1
            End With
        End If
    End If

End Sub
Private Sub lstDifference_KeyUp(KeyCode As Integer, Shift As Integer)
    FlxDetail.Row = BeforShowDifferenceFlxRow
    
     If KeyCode = 13 Then
        
        If MaxRowFlexGrid > 1 And Val(FlxDetail.TextMatrix(FlxDetail.Row, 5)) <> 0 Then
            SplitByOptions FlxDetail.Row
            lblNum.Caption = ""
            'End If
            
            If lstDifference.SelCount > 0 Then
            
                'Dim ArrDifferences() As String
                ReDim ArrDifferences(0) As String
                For i = 0 To lstDifference.ListCount - 1
                    If lstDifference.Selected(i) = True Then
                        On Error GoTo ErrHandler
                        ReDim Preserve ArrDifferences(UBound(ArrDifferences) + 1)
                        On Error GoTo 0
                        ArrDifferences(UBound(ArrDifferences)) = lstDifference.ItemData(i)
                        FlxDetail.TextMatrix(FlxDetail.Row, 9) = ""
                        FlxDetail.TextMatrix(FlxDetail.Row, 10) = ""
                    Else
                        ArrCostDifferences(i + 1) = 0
                    End If
                Next i
                
                Dim j As Integer
                For i = LBound(ArrDifferences) To UBound(ArrDifferences) - 1
                    'Debug.Print ArrDifferences(i)
                    For j = i + 1 To UBound(ArrDifferences)
                        If Val(ArrDifferences(i)) = Val(ArrDifferences(j)) Then
                            ArrDifferences(j) = 0
                            ArrCostDifferences(j + 1) = 0
                        ElseIf Val(ArrDifferences(i)) + Val(ArrDifferences(j)) = 0 Then
                            ArrDifferences(i) = 0
                            ArrDifferences(j) = 0
                            ArrCostDifferences(i + 1) = 0
                            ArrCostDifferences(j + 1) = 0
                        End If
                    Next j
                    
                Next i
                
                'Debug.Print ArrDifferences(i)
                
                For i = LBound(ArrDifferences) To UBound(ArrDifferences)
                    If Val(ArrDifferences(i)) <> 0 Then
                        For j = 0 To lstDifference.ListCount - 1
                            If Val(ArrDifferences(i)) = lstDifference.ItemData(j) Then
                                FlxDetail.TextMatrix(FlxDetail.Row, 9) = FlxDetail.TextMatrix(FlxDetail.Row, 9) & lstDifference.ItemData(j) & ";"
                                FlxDetail.TextMatrix(FlxDetail.Row, 10) = FlxDetail.TextMatrix(FlxDetail.Row, 10) & lstDifference.List(j) & ","
                                Exit For
                            End If
                        Next j
                        
                    End If
                Next i
                
                If Right(FlxDetail.TextMatrix(FlxDetail.Row, 9), 1) = ";" Then
                    
                    FlxDetail.TextMatrix(FlxDetail.Row, 9) = left(FlxDetail.TextMatrix(FlxDetail.Row, 9), Len(FlxDetail.TextMatrix(FlxDetail.Row, 9)) - 1)
                    FlxDetail.TextMatrix(FlxDetail.Row, 10) = left(FlxDetail.TextMatrix(FlxDetail.Row, 10), Len(FlxDetail.TextMatrix(FlxDetail.Row, 10)) - 1)
                
                End If
                
            Else
            
                FlxDetail.TextMatrix(FlxDetail.Row, 9) = ""
                FlxDetail.TextMatrix(FlxDetail.Row, 10) = ""
                ReDim ArrCostDifferences(0)
            End If
            
        
    End If
    
    Me.lstDifference.Visible = False
        'HideLstBoxes KeyCode
    If FlxDetail.TextMatrix(FlxDetail.Row, 3) <> "" And clsStation.HasOptionPrice Then
        FlxDetail.TextMatrix(FlxDetail.Row, 3) = FlxDetail.TextMatrix(FlxDetail.Row, 3) - OldCostDifference
        FlxDetail.TextMatrix(FlxDetail.Row, 3) = FlxDetail.TextMatrix(FlxDetail.Row, 3) + GetCostDifferences
        FlxDetail.TextMatrix(FlxDetail.Row, 4) = Val(FlxDetail.TextMatrix(FlxDetail.Row, 3)) * Val(FlxDetail.TextMatrix(FlxDetail.Row, 1))
        EnableBeforShowDifferenceFlxRow = False
        RefreshLables
    End If
    FlxDetail.Select MaxRowFlexGrid, 1
    ElseIf KeyCode = 27 Then
        
        HideLstBoxes KeyCode
    End If
    
    KeyCode = 0
    
    Exit Sub
    
ErrHandler:
    If err.Number = 9 Then
        ReDim ArrDifferences(0)
        Resume Next
    End If
End Sub

Private Sub lstDifference_LostFocus()
    
'    lstDifference_KeyUp vbKeyReturn, 0

End Sub
Private Sub Acknowledge()
    
    Sleep 150
    Dim j, k As Integer
    For j = 1 To 2
 '       StartTimeReader
        For k = 1 To 30
            mscSerial(TimeReaderPort).Output = Chr$(113)
        Next k
        
        
        mscSerial(TimeReaderPort).Output = Chr$(221) + Chr$(29) + Chr$(1) + Chr$(0) + Chr$(48) + Chr$(0) + Chr$(44) + Chr$(1) + Chr$(29) _
                         + Chr$(41) + Chr$(1) + Chr$(0) + Chr$(127) + Chr$(47) + Chr$(2) + Chr$(32) + Chr$(45) + Chr$(17) + Chr$(46) _
                         + Chr$(15) + Chr$(45) + Chr$(16) + Chr$(43) + Chr$(1) + Chr$(42) + Chr$(0) + Chr$(255) + Chr$(0) _
                         + Chr$(159) + Chr$(3) + Chr$(0) + Chr$(0)
      
    Next j

End Sub
Private Sub TimerReader_Timer()
    StartTimeReader
End Sub
Private Sub StartTimeReader()
    mscSerial(TimeReaderPort).Output = Chr$(238) + Chr$(238) + Chr$(238) + Chr$(238) + Chr$(238) + Chr$(238) + Chr$(238) + Chr$(238) + Chr$(238) + Chr$(238) + Chr$(238) + Chr$(238) + Chr$(238) + Chr$(238) + Chr$(238) _
                   + Chr$(238) + Chr$(238) + Chr$(238) + Chr$(238) + Chr$(238) + Chr$(238) + Chr$(238) + Chr$(238) + Chr$(238) + Chr$(238) _
                   + Chr$(221) + Chr$(1) + Chr$(49) + Chr$(0) + Chr$(10) + Chr$(48) + Chr$(10) + Chr$(0) + Chr$(118) _
                   + Chr$(255) + Chr$(255) + Chr$(124) + Chr$(187)


End Sub

Private Sub TimerALM_Timer()
On Error Resume Next

    Dim PEIns As Long         'pointer for end of packet
    Dim TIns As String        'Len Packet
    Dim LiN As Integer        'Line Number
    Dim LFind As Integer      'for find next line busy for record voice
    Dim vData$                'Voice Data for record
    Dim Zm As String
    Dim RngC As Integer
    Dim VoiceLine As String
    Dim VoicePath As String
    Dim fs As New FileSystemObject
    
    VoicePath = App.Path & "\Records"
    If clsStation.VoiceRecord Then
        If Not (fs.FolderExists(VoicePath)) Then
            fs.CreateFolder (VoicePath)
        End If
        VoicePath = VoicePath & "\" & DateToNumber8(Right(clsDate.shamsi(Date), 8)) & "\"
'        If Not (FS.FolderExists(VoicePath)) Then
'            FS.CreateFolder (VoicePath)
'        End If
    End If
    
    
    If mscSerial(AlmPort).InBufferCount < 1 And Ins = "" Then Exit Sub
    Dim aa As String
    aa = mscSerial(AlmPort).Input
    Ins = Ins & aa
    
              
StartDetect:
    
    If clsStation.AlmLogFile = True Then
        LogSave (aa)
    End If
    
    If UCase(Mid(Ins, InStr(1, Ins, ".", vbTextCompare) + 1, 5)) = "START" Then
        Sleep 100
        Ins = ""
        Exit Sub
    End If
    
    If Asc(left$(Ins, 1)) <> 13 Then   'not report
        If Asc(left$(Ins, 1)) = 253 Then 'Enter Command from ALM
            PEIns = InStr(1, Ins, Chr$(254), vbBinaryCompare)
            If PEIns = 0 Then
                PEIns = InStr(1, Ins, Chr$(13), vbBinaryCompare)
                If PEIns = 0 Then Ins = "": Exit Sub Else PEIns = PEIns - 1: GoTo ExS
            End If
            GoTo ExS
        End If
        PEIns = InStr(1, Ins, Chr$(10), vbBinaryCompare)
        If PEIns = 0 Then
            PEIns = InStr(1, Ins, Chr$(13), vbBinaryCompare)
            If PEIns = 0 Then Ins = "": Exit Sub Else PEIns = PEIns - 1: GoTo ExS
        End If
        GoTo ExS 'not complete voice or report or command
    End If
    
    
    PEIns = InStr(1, Ins, Chr$(10), vbBinaryCompare) 'A Pointer to End of First Packet
    If PEIns = 0 Then Exit Sub        'Data not compelet (not closed with chr 10) goto get
    
    If Asc(Mid(Ins, 2, 1)) = 200 Then  'Record Permision AND have voice
    If LineDial(2, 7) = "" Then
         mscSerial(AlmPort).Output = Chr$(252) & "L9CO" & Chr$(253) 'goto 9 = Stop Record Voice
    Else
        vData = Space(PEIns - 4)
        vData = Mid$(Ins, 4, PEIns - 4) 'Clear VoiceData
        Put Val(LineDial(2, 7)), , vData 'Put voice to File (FileHandle,,VoiceData)
    End If
    GoTo ExS
    End If
    
    Zm = Zaman
    TIns = left(Ins, InStr(1, Ins, "EEE", vbTextCompare)): TIns = CutOut(TIns, Chr(10) & Chr(13))
    LiN = Mid(TIns, InStr(1, TIns, ":", vbTextCompare) - 1, 1)
    
    
    
    If InStr(1, TIns, "RING", vbTextCompare) > 0 Then   'if report about RING
        FWModem(LiN - 1).BackColor = vbRed ' &H80000003&
        GoTo ExS
    End If
    
    If InStr(1, TIns, ":CallerID", vbTextCompare) > 0 Then  'If Detect CallerID in Packet
        Dim InID As String
        InID = Trim(Mid(TIns, InStr(1, TIns, ":CallerID", vbTextCompare) + 10, 22))
        
        If Trim$(LineDial(LiN, 3)) = "" Then LineDial(LiN, 3) = Zm
        LineDial(LiN, 4) = InID
        InID = FNSplitt(InID)
        
        LineDial(LiN, 4) = FixTel(FNSplitt(LineDial(LiN, 4)))
        If Len(LineDial(LiN, 4)) < 7 Then GoTo ExS
        If Trim$(LineDial(LiN, 3)) = "" Then LineDial(LiN, 3) = Zm
        
        
        FWModem(LiN - 1).ToolTipText = InID 'Val(Left(InID, 15))
        GetCallerInfo AlmPort, TIns, LiN
        GoTo ExS
    End If
    
    If UCase(Mid(TIns, InStr(1, TIns, ":", vbTextCompare) + 1, 6)) = "HOOKON" Then  'if HangUp report
        FWModem(LiN - 1).BackColor = vbGreen ' &H80000003&
        If Trim$(LineDial(LiN, 3)) = "" Then LineDial(LiN, 3) = Zm
        If LineDial(LiN, 2) <> "" Then
            TIns = Space(20)
            Mid(TIns, 1, 20) = Trim$(LineDial(LiN, 2))
        End If
        If LineDial(1, 7) = "" Then 'If Active Main Voice System and free voice channel
            mscSerial(AlmPort).Output = Chr$(252) & "L" & Trim$(Str$(LiN)) & "CO" & Chr$(253) 'Start Record Line x
            LineDial(1, 7) = Trim$(Str$(LiN))
            LineDial(2, 7) = Trim$(Str$(FreeFile))
            LineDial(3, 7) = VoicePath & Mid$(LineDial(LiN, 3), 12, 2) & Mid$(LineDial(LiN, 3), 15, 2) & Mid$(LineDial(LiN, 3), 18, 2) & "_" & LineDial(1, 7)
            Open LineDial(3, 7) For Binary Access Write As Val(LineDial(2, 7))
            vData = Space(1000)
                vData = Right$(LineDial(3, 7), 26) & Chr$(13) & Chr$(10) & Chr$(13) & Chr$(10) & Chr$(13) & Chr$(10) & Chr$(13) & Chr$(10) & left$(LineDial(3, 7), Len(LineDial(3, 7)) - 26)
                vData = vData & Space(998 - Len(vData)) & Chr$(13) & Chr$(10)
                If Len(vData) > 1000 Then vData = left$(vData, 1000)
            Put Val(LineDial(2, 7)), , vData
        End If
                
        GoTo ExS
    End If
    
    If UCase(Mid(TIns, InStr(1, TIns, ":", vbTextCompare) + 1, 7)) = "HOOKOFF" Then
        FWModem(LiN - 1).BackColor = &H80000016  '&H808000 '
        
        If LineDial(LiN, 1) <> "" Then
        LineDial(LiN, 4) = Trim$(Str$(B2DS(LineDial(LiN, 3), Zm))) 'Mid(TIns, InStr(1, TIns, "tion:", vbTextCompare) + 5, 4)
        TIns = Space(24)
            Mid(TIns, 1, 20) = Trim$(LineDial(LiN, 2))
            Mid(TIns, 21, 4) = LineDial(LiN, 4)
        End If
        If Val(LineDial(1, 7)) = LiN Then
             mscSerial(AlmPort).Output = Chr$(252) & "L9CO" & Chr$(253) 'End Record Voice
            Close Val(LineDial(2, 7)) 'close Voice file !
            If Not (fs.FolderExists(VoicePath)) Then
                fs.CreateFolder (VoicePath)
            End If
            CopyFile2Wav LineDial(3, 7), LineDial(3, 7) & ".WAV"
            LineDial(1, 7) = "": LineDial(2, 7) = "": LineDial(3, 7) = "": LineDial(4, 7) = "": LineDial(5, 7) = "": LineDial(6, 7) = "": LineDial(7, 7) = "": LineDial(8, 7) = ""
            For LFind = 1 To 8
                If LineDial(LFind, 3) <> "" And LFind <> LiN And _
                InStr(1, VoiceLine, Trim$(Str$(LFind)), vbTextCompare) > 0 Then
                     mscSerial(AlmPort).Output = Chr$(252) & "L" & Trim$(Str$(LFind)) & "CO" & Chr$(253) 'Start Record Line x
                    LineDial(1, 7) = Trim$(Str$(LFind))
                    LineDial(2, 7) = FreeFile
                    LineDial(3, 7) = VoicePath & Mid$(LineDial(LFind, 3), 12, 2) & Mid$(LineDial(LFind, 3), 15, 2) & Mid$(LineDial(LFind, 3), 18, 2) & "_" & LineDial(1, 7)
                    Open LineDial(3, 7) For Binary Access Write As LineDial(2, 7)
                    vData = Space(1000)
                        vData = Right$(LineDial(3, 7), 26) & Chr$(13) & Chr$(10) & Chr$(13) & Chr$(10) & Chr$(13) & Chr$(10) & Chr$(13) & Chr$(10) & left$(LineDial(3, 7), Len(LineDial(3, 7)) - 26)
                        vData = vData & Space(998 - Len(vData)) & Chr$(13) & Chr$(10)
                        If Len(vData) > 1000 Then vData = left$(vData, 1000)
                    Put Val(LineDial(2, 7)), , vData
                    Exit For
                End If
            Next
        End If
        LineDial(LiN, 4) = Zm
        If Trim$(LineDial(LiN, 2)) <> "" And Trim$(LineDial(LiN, 1)) = "" Then 'Dial and outgoing
            'Cmd.ActiveConnection = Conn
            'Cmd.CommandText = "INSERT INTO TelOut (TimeS,TimeE,Tel,Ln,Opr) VALUES ('" & Trim$(LineDial(Lin, 3)) & "','" & B2D(Trim$(LineDial(Lin, 3)), Trim$(LineDial(Lin, 4))) & "','" & LineDial(Lin, 2) & "','" & Lin & "','" & UserName & "')"
            'Cmd.Execute
        End If
        LineDial(LiN, 1) = "": LineDial(LiN, 2) = "": LineDial(LiN, 3) = "": LineDial(LiN, 4) = ""
    
        
        GoTo ExS
    End If
    
    If UCase(Mid(TIns, InStr(1, TIns, ":", vbTextCompare) + 1, 6)) = "DIALED" Then  'If Dial from inside to outside
        If LineDial(LiN, 1) <> "" Then GoTo ExS
        GoTo ExS
    End If

ExS:
'    Dim EEEnd As Integer
'    EEEnd = InStr(1, Ins, "EEE", vbTextCompare)
'    If Len(Ins) < EEEnd + 7 Or EEEnd = 0 Then Ins = "": Exit Sub
'    Ins = Right(Ins, Len(Ins) - EEEnd - 2)


''''    Debug.Print Ins
    DoEvents
    If Len(Ins) > PEIns Then Ins = Right$(Ins, Len(Ins) - PEIns): GoTo StartDetect Else Ins = ""
    Exit Sub
    
ErrHandler:
    modgl.LogSaveNew "frmInvoice", err.Description, err.Number, err.Source, "TimerALM_Timer"
    MsgBox err.Description
End Sub
Private Function CutOut(ByVal str1 As String, Str2 As String) As String
On Error Resume Next
'Cut out any chr in (Str2) FROM (Str1) : ?CutOut("Mohammad","ma") => [Mohd]
Dim i1, i2 As Long
Dim SCut As Boolean
Dim StrOut As String
StrOut = ""
For i1 = 1 To Len(str1)
    SCut = False
    For i2 = 1 To Len(Str2)
        If Mid(str1, i1, 1) = Mid(Str2, i2, 1) Then SCut = True
    Next
    If Not SCut Then StrOut = StrOut & Mid(str1, i1, 1)
Next
CutOut = StrOut
End Function
Private Function FNSplitt(ByVal Str As String) As String
On Error Resume Next
Dim s As Integer: s = 0
Dim StrOut As String: StrOut = ""

For i = 1 To Len(Str)
    If Asc(Mid(Str, i, 1)) > 47 And Asc(Mid(Str, i, 1)) < 58 Then
        s = 1
        StrOut = StrOut & Mid(Str, i, 1)
    Else
        If s = 1 Then FNSplitt = StrOut: Exit Function
    End If
Next
End Function
'Private Sub LogSave(InputString As String)
'    Dim filetemp As New FileSystemObject
'    Dim tempstring As TextStream
'    Dim CallerIDFile As String
'
'
'    CallerIDFile = App.Path & "\CallerID" & DateToNumber8(Right(ClsDate.shamsi(Date), 8)) & ".Log"
'
'    If filetemp.FileExists(CallerIDFile) Then
'        Set tempstring = filetemp.OpenTextFile(CallerIDFile, ForAppending, False, TristateFalse)
'    Else
'        filetemp.CreateTextFile CallerIDFile
'        Set tempstring = filetemp.OpenTextFile(CallerIDFile, ForWriting, False, TristateFalse)
'    End If
'    tempstring.WriteLine (InputString)
'    tempstring.Close
'
'End Sub

Private Sub mscSerial_OnComm(index As Integer)
On Error GoTo ErrorHandler

    Select Case mscSerial(index).CommEvent
    
        Case comEvReceive   ' Received RThreshold # of
            
            mscSerial(index).RThreshold = 0
            Dim InputString, TempStr As String
            Dim kk As Integer
            Dim jj As Integer
            
            Select Case DeviceType(index)
              Case EnumDeviceType.Pager
                    Sleep 100
                     
                    InputString = mscSerial(index).Input
                    If Asc(Mid(InputString, 5, 1)) = 13 Then
                        If PagerNo > 0 Then
                            PagerAction PagerNo
                            PagerNo = 0
                        End If
                    ElseIf Asc(Mid(InputString, 5, 1)) >= 48 And Asc(Mid(InputString, 5, 1)) <= 57 Then
                        PagerNo = Val(CStr(PagerNo) + CStr(Mid(InputString, 5, 1)))
                    ElseIf Asc(Mid(InputString, 5, 1)) = 12 Then
                        PagerNo = 0
                    End If
              Case EnumDeviceType.CardReader
                    If DeviceCode(index) = EnumDevice.BarcodeTimeReader Then
                    
                        TimerReader.Enabled = False
                        Sleep 200
                        InputString = mscSerial(index).Input
                        
                        kk = InStr(1, InputString, "pppppppppppppppppppppppppppp", 1)
                        If kk > 0 Then
                            TempStr = Mid(InputString, 29 + kk, 7)
                            If Asc(Mid(TempStr, 1, 1)) = 51 Then
                                CreditCode = Val(CLng(Asc(Mid(TempStr, 4, 1))) + CLng(Asc(Mid(TempStr, 5, 1))) * 256) + 256 * ((CLng(Asc(Mid(TempStr, 6, 1)) * 256) + CLng(Asc(Mid(TempStr, 7, 1))) * 256 * 256))
                                Acknowledge
                                If CreditCode <> TempCreditCode Then
                                    TempCreditCode = CreditCode
                                    FindCust
                                End If
                            End If
                        End If
                        TimerReader.Enabled = True
                    Else
                        Sleep 500
                        CreditCode = Val(Mid(mscSerial(index).Input, clsStation.StartNumberCartReader, clsStation.NumberOfCardReader))
                        FindCust
                    End If
                Case EnumDeviceType.Bascule
                    
                    Dim strFinal As String
                    If clsStation.BasculeModel = EnumDevice.Pand Then
                        InputString = mscSerial(index).Input
                        If Trim(Right(InputString, 8)) <> "" Then
                            InputString = Right(InputString, 8)
                        ElseIf Trim(left(InputString, 8)) <> "" Then
                            InputString = left(InputString, 8)
                        Else
                            Exit Sub
                        End If
                        If Len(InputString) < 4 Then Exit Sub
                        jj = 1
                        Do While Not jj >= 8
                            If Asc(Mid(InputString, jj, 1)) = 187 Then
                                jj = jj + 1
                                Exit Do
                            End If
                            jj = jj + 1
                        Loop
                        If (jj >= 5) Then
    ''                        frmDisMsg.lblMessage = "«ÿ·«⁄«  «—”«·Ì »Â ÅÊ—  ‘„«—Â  " & Me.PortNo & "’ÕÌÕ ‰Ì”  "
                        End If
                        If Asc(Mid(InputString, jj, 1)) = 240 Then
                            ShowDisMessage "Ê“‰ «÷«›Ì", 1200
                            Exit Sub
                        End If
                        If Asc(Mid(InputString, jj, 1)) = 224 Then
    ''''                        frmDisMsg.lblMessage = "«ÿ·«⁄«  Ê“‰ „—»Êÿ »Â ÅÊ—  ‘„«—Â  " & Me.PortNo & "’ÕÌÕ ‰Ì”  "
                            strFinal = "0.00"
                            txtScale.Text = strFinal
                            lblScale(0).Caption = txtScale.Text
                            Exit Sub
                        End If
                        For i = 1 To 6
                            Select Case i
                                Case 1:
                                    strFinal = strFinal + Trim(Str(Int(Asc(Mid(InputString, jj, 1)) / 16)))
                                    strFinal = Str(Val(strFinal))
                                Case 2:
                                    strFinal = strFinal + Trim(Str(Int(Asc(Mid(InputString, jj, 1)) And 15)))
                                    strFinal = Str(Val(strFinal))
                                    jj = jj + 1
                                Case 3:
                                    strFinal = strFinal + Trim(Str(Int(Asc(Mid(InputString, jj, 1)) / 16)))
                                    strFinal = Str(Val(strFinal)) + "."
                                Case 4:
                                    strFinal = strFinal + Trim(Str(Int(Asc(Mid(InputString, jj, 1)) And 15)))
                                    jj = jj + 1
                                Case 5:
                                    strFinal = strFinal + Trim(Str(Int(Asc(Mid(InputString, jj, 1)) / 16)))
                                Case 6:
                                    strFinal = strFinal + Trim(Str(Int(Asc(Mid(InputString, jj, 1)) And 15)))
                            End Select
                        Next i
                        txtScale.Text = strFinal
                        If txtScale.Text <> "" Then txtScale.Text = Format(txtScale.Text, "00.000")
                        lblScale(0).Caption = txtScale.Text
                        If Val(txtScale.Text) > 10 Then
                            lblScale(0).Font.size = 18
                        Else
                            lblScale(0).Font.size = 20
                        End If
                
                    ElseIf clsStation.BasculeModel = EnumDevice.Sairan Then
                        If SairanFlag = True Then Exit Sub
                        SairanFlag = True
                        Sleep 100
                        InputString = mscSerial(index).Input
                        If Len(InputString) < 12 Then
                            Exit Sub
                        End If
                
                        strFinal = Mid(InputString, 7, 2) & "." & Mid(InputString, 9, 3)
                        txtScale.Text = strFinal
                        lblScale(0).Caption = txtScale.Text
                        If Val(txtScale.Text) > 10 Then
                            lblScale(0).Font.size = 18
                        Else
                            lblScale(0).Font.size = 20
                        End If
                    
                    End If
                Case EnumDeviceType.Pos
                    Sleep 1000
                    'ProcessPOSTrain mscSerial(index).Input
                Case EnumDeviceType.Modem
                    Call mdifrm.FWMMedia1.PlayWaveFile(App.Path & "\Sound\Ringin.wav", True, False)
                  ''''  Call mdifrm.FWMMedia1.PlayWaveSystem(FLWMMedia.flwMMWaveIn, True, False)
                    'Call mdifrm.FWMMedia1.StopWaveFile
                    Sleep 200
                    Dim Inputstr As String
                    Inputstr = mscSerial(index).Input
                    LogSaveNew Inputstr, "", "", "", ""
                  '  Debug.Print InputStr
                    If (DeviceCode(index) <> EnumDevice.CallerIdInterface1 And DeviceCode(index) <> EnumDevice.CallerIdInterface2_AlmP3 And DeviceCode(index) <> EnumDevice.CallerIdInterface2_AlmP1 And DeviceCode(index) <> EnumDevice.CallerIdInterface2_AlmP6) Then
                        FWModem(1).BackColor = vbRed     ' &H80000003&
                        
                    ElseIf DeviceCode(index) = EnumDevice.CallerIdInterface1 Then
                        kk = InStr(1, Inputstr, "L", 1)
                        LineNumber = Val(Mid(Inputstr, kk + 1, 1))
                        If LineNumber > 0 And LineNumber < 9 And kk > 0 Then
                            jj = InStr(1, Inputstr, "@", 1)
                            If jj > kk Then
                                FWModem(LineNumber - 1).BackColor = vbRed ' &H80000003&
                                FWModem(LineNumber - 1).ToolTipText = Val(Mid(Inputstr, kk + 1, jj - kk - 2)) '
                                
                                GetCallerInfo index, Inputstr, LineNumber
                            End If
                        End If
                    ElseIf DeviceCode(index) = EnumDevice.CallerIdInterface2_AlmP3 Then
                        kk = InStr(1, Inputstr, "L", 1)
                        LineNumber = Val(Mid(Inputstr, kk + 1, 1))
                        If LineNumber > 0 And LineNumber < 9 And kk > 0 Then
                            jj = InStr(1, Inputstr, "@", 1)
                            If jj > kk Then
                                FWModem(LineNumber - 1).BackColor = vbRed ' &H80000003&
                                FWModem(LineNumber - 1).ToolTipText = Val(Mid(Inputstr, kk + 1, jj - kk - 2)) '
                                GetCallerInfo index, Inputstr, LineNumber
                            End If
                        End If
                    ElseIf DeviceCode(index) = EnumDevice.CallerIdInterface2_AlmP6 Then  ' New protocol with Danzhe
''''                        kk = InStr(1, Inputstr, "L", 1)
''''                        LineNumber = Val(Mid(Inputstr, kk + 1, 1))
''''                        If LineNumber > 0 And LineNumber < 9 And kk > 0 Then
''''                                jj = InStr(1, LCase(Inputstr), "callerid:", 1)
''''                                If jj > kk Then
''''                                    FWModem(LineNumber - 1).BackColor = vbRed ' &H80000003&
''''                                    FWModem(LineNumber - 1).ToolTipText = Left(LTrim(Mid(Inputstr, jj + 9)), 15)
'''''                                    If clsStation.NetworkCallerId = True Then
'''''                                         mdifrm.WinsockUdp.SendData str(LineNumber) & str(Index) & InputStr
'''''                                    End If
''''                                    'if the station is not currently serving any other call, serve the incoming call
''''                                        GetCallerInfo Index, Inputstr, LineNumber
''''
''''                                End If
''''                        End If
                     
                     ElseIf DeviceCode(index) = EnumDevice.CallerIdInterface2_AlmP1 Then
''''                        Dim ID As String
''''                        Dim Date_Location, Start_Of_ID, End_Of_ID, tmp As Integer
''''                        Dim Has_No_Date As Boolean
''''                        kk = InStr(1, InputStr, "L", 1)
''''                        If Val(Mid(InputStr, kk + 1, 1)) > 0 And Val(Mid(InputStr, kk + 1, 1)) < 9 And kk > 0 Then
''''
''''                              '--- Find location of ID ---
''''                                Start_Of_ID = InStr(LCase(InputStr), "callerid:") + 9
''''                                If Start_Of_ID = 9 Then Exit Sub
''''                                Date_Location = InStr(Start_Of_ID, InputStr, "/")
''''                                If Date_Location > 0 Then Has_No_Date = True
''''
''''                                End_Of_ID = Date_Location - 4
''''
''''
''''                                '--- Take Out Caller Information ---
''''                                If Has_No_Date = True Then
''''                                   ID = Mid(InputStr, Start_Of_ID, (End_Of_ID - Start_Of_ID))
''''                                ElseIf Has_No_Date = False Then
''''                                   ID = Mid(InputStr, Start_Of_ID, (Len(InputStr) - Start_Of_ID))
''''                                End If
''''                                jj = InStr(1, LCase(InputStr), "callerid:", 1)
''''                                If jj > kk Then
''''                                    FWModem(Val(Mid(InputStr, kk + 1, 1)) - 1).BackColor = vbRed ' &H80000003&
'''''                                    FWModem(Val(Mid(InputStr, kk + 1, 1)) - 1).ToolTipText = Mid(LTrim(Mid(InputStr, jj + 9)), 10, 15)  '
''''                                    FWModem(Val(Mid(InputStr, kk + 1, 1)) - 1).ToolTipText = Val(Left(ID, 15))
''''                                    GetCallerInfo Index, InputStr
''''                                End If
''''                        End If
                    End If
                
            End Select
            If index <> AlmPort Then
                mscSerial(index).RThreshold = RThreshold(index)
            End If
            If mscSerial(index).PortOpen = False Then
                mscSerial(index).PortOpen = True
            End If
    End Select

Exit Sub

ErrorHandler:
    MsgBox "mscSerial_OnComm" & err.Description
    LogSaveNew "frmInvoice => ", err.Description, err.Number, err.Source, "mscSerial_OnComm"
    err.Clear
End Sub

Private Sub ResizeKit1_ExitResize(ByVal XScale As Double, ByVal YScale As Double)

    If formloadFlag = True Then
        SaveSetting strMainKey, Me.Name, "Height", Me.Height
        SaveSetting strMainKey, Me.Name, "Width", Me.Width
        FlxDetail.RowHeightMax = FlxDetail.Height / (MaxInvoiceRows * 1.08) '8.2
        FlxDetail.RowHeightMin = FlxDetail.Height / (MaxInvoiceRows * 1.11)  '8.5
'        MenuBar.Width = frameMenu.Width - 50
'        Dim i As Long
'        For i = 2 To 5
'            MenuBar.Panels(i).Width = (frameMenu.Width - 1200) / 4
'        Next
'        MenuBar.Panels(1).Width = 500
'        MenuBar.Panels(6).Width = 500
        
        SetBtnMenuPosition
        LblAccNo.Height = LblInvoicePrint.Height
        LblInvoicePrint.top = LblAccNo.top
    End If
End Sub

Private Sub StatusBar_PanelClick(ByVal Panel As MSComctlLib.Panel)
    If ClsFormAccess.NavigateFactor = False Then
        ShowDisMessage "‘„« »Â «Ì‰ «„ò«‰ œ” —”Ì ‰œ«—Ìœ", 1500
        Exit Sub
    End If
    PanelClick Panel.index
End Sub

Private Sub TimerNumber_Timer()
    If MyFormAddEditMode = AddMode Then
       Me.Number
    End If
    If clsStation.DeliveryNoView Then
       CalculateDelivery
       CalculateTemporary
    End If
End Sub

Private Sub lblCarryFeeTotal_Click()
    lblCarryFeeTotal.RightToLeft = IIf(lblCarryFeeTotal.RightToLeft, False, True)
End Sub

Private Sub lblDiscountTotal_Click()
    lblDiscountTotal.RightToLeft = IIf(lblDiscountTotal.RightToLeft, False, True)
End Sub
Private Sub lblPayFactorTotal_Click()
    lblPayFactorTotal.RightToLeft = IIf(lblPayFactorTotal.RightToLeft, False, True)
End Sub

Private Sub txtAddress_GotFocus()
    If AddressFlag = False Then AddressFlag = True 'TxtAddress = "":
End Sub

Private Sub TxtAddress_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 And Shift = 0 And Val(lblCustomer.Tag) <> -1 Then
        SaveCustAddress
    End If
End Sub

Private Sub SaveCustAddress()
    ReDim Parameter(1) As Parameter
    Parameter(0) = GenerateInputParameter("@Code", adInteger, 4, Val(lblCustomer.Tag))
    Parameter(1) = GenerateInputParameter("@Address", adVarWChar, 255, Trim(TxtAddress.Text))
    
    Set rctmp = RunParametricStoredProcedure2Rec("Update_Cust_By_Address", Parameter)
End Sub
Private Sub SaveCustDescription()
    ReDim Parameter(1) As Parameter
    Parameter(0) = GenerateInputParameter("@Code", adInteger, 4, Val(lblCustomer.Tag))
    Parameter(1) = GenerateInputParameter("@Description", adVarWChar, 255, Trim(TxtCustDescription.Text))
    
    Set rctmp = RunParametricStoredProcedure2Rec("Update_Cust_By_Description", Parameter)
End Sub
Private Sub txtAddress_LostFocus()
    AddressFlag = False
End Sub


Private Sub TxtCustDescription_GotFocus()
    CustDescriptionFlag = True
End Sub

Private Sub TxtCustDescription_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 And Shift = 0 Then
        SaveCustDescription
    End If
End Sub

Private Sub TxtCustDescription_LostFocus()
    CustDescriptionFlag = False
End Sub

Private Sub txtDescription_Change()
    BtnKeypad(11).Enabled = True     '"%"
    BtnKeypad(10).Enabled = True      '"."
    BtnKalaDelete.Enabled = True
    lblNum.Caption = ""
    lblBarCode.Caption = ""
End Sub

Private Sub txtDescription_GotFocus()
    If Trim(txtDescription) = "ÅÌ€«„" Then txtDescription.Text = ""
    textDescription = True
    textDescriptionFlag = True
End Sub

Private Sub txtDescription_LostFocus()
    textDescription = False
End Sub

Private Sub txtNo_Change()
    If Trim(txtNo.Text) <> "" Then
        If Right(txtNo.Text, 3) <> "000" Then
            FWLed1.Value = Right(txtNo.Text, 3)
        Else
            FWLed1.Value = "1" + Right(txtNo.Text, 3)
        End If
    End If
End Sub


Private Sub lblPackingTotal_Click()
    lblPackingTotal.RightToLeft = IIf(lblPackingTotal.RightToLeft, False, True)
End Sub


Private Sub lblServiceTotal_Click()
    lblServiceTotal.RightToLeft = IIf(lblServiceTotal.RightToLeft, False, True)
End Sub


Private Sub txtSumCountNo_Click()
    txtSumCountNo.RightToLeft = IIf(txtSumCountNo.RightToLeft, False, True)
End Sub

Public Sub RefreshLables()    'For Refresh Lables When Edit

    Dim rctmp As New ADODB.Recordset
    Dim ValueCountNo, ValueCountWeight, ValueSumWeight, ValueWeightTotal, ValueFeeTotal, ValueGoodDiscount, ValueGoodsDuty, ValueGoodsTax As Double
    Dim a As TextBox
    
    On Error Resume Next
    Dim ii As Integer
    Dim amount As Double
    Dim Fee As Double
    Dim discountPercent As Double
    Dim TotalFee As Double
    Dim SumTotalFee As Double
    txtSumFeeTotal.Text = 0
    txtSumCountNo.Caption = 0
'    txtSumCountWeight.Caption = 0
    SumTotalFee = 0#
    ValueGoodsDuty = 0
    ValueGoodsTax = 0
    ValueGoodDiscount = 0
    With FlxDetail
        For ii = 1 To MaxRowFlexGrid - 1
            amount = Val(.TextMatrix(ii, IndexColAmount))
            Fee = Val(.TextMatrix(ii, IndexColFee))
            discountPercent = Val(.TextMatrix(ii, IndexColDiscountPercent))
            
            TotalFee = amount * Fee
            FlxDetail.TextMatrix(ii, 4) = TotalFee
            Me.txtSumFeeTotal.Text = Me.txtSumFeeTotal.Text + TotalFee
            ValueGoodDiscount = ValueGoodDiscount + (TotalFee * discountPercent / 100)
            If ValueGoodDiscount <> 0 Then ValueGoodDiscount = Format(ValueGoodDiscount, "##")
            If chKTax.Value = True Then
                '„Õ«”»Â ⁄Ê«—÷ ﬂ«·«Â«
                If .TextMatrix(ii, IndexColDuty) = True Then ValueGoodsDuty = ValueGoodsDuty + TotalFee
                If .TextMatrix(ii, IndexColTax) = True Then ValueGoodsTax = ValueGoodsTax + TotalFee
            End If
'            If Val(FlxDetail.TextMatrix(ii, 7)) <> 1 Then        'Numeric
            
                txtSumCountNo.Caption = Val(txtSumCountNo.Caption) + amount
'            Else
'                txtSumCountWeight.Caption = Val(txtSumCountWeight.Caption) + 1
'            End If
        Next ii
    End With
    LblSubTotal.Caption = Me.txtSumFeeTotal.Text
'    ValueGoodsDuty = CLng(ValueGoodsDuty)
'    ValueGoodsTax = CLng(ValueGoodsTax)
        
Select Case MyFormAddEditMode
    Case EnumAddEditMode.AddMode, EnumAddEditMode.EditMode, EnumAddEditMode.ManipulateMode, EnumAddEditMode.RefferedMode
        lblServiceTotal = CLng(Val(txtSumFeeTotal.Text) * ServiceRate / 100)
        lblDiscountTotal = CLng((Val(txtSumFeeTotal.Text) * Val(txtDiscountPercent.Text) / 100) + ValueGoodDiscount + Val(txtDiscount.Text))
        If Val(lblDiscountTotal) <> 0 Then lblDiscountTotal = Format(lblDiscountTotal, "##")
        lblCarryFeeTotal = CLng(Val(txtCarryFee.Text) + (Val(txtSumFeeTotal.Text) * Val(txtCarryFeePercent.Text) / 100)) '+ Val(txtCarryFee.Text)
        If Val(lblCarryFeeTotal) <> 0 Then lblCarryFeeTotal = Format(lblCarryFeeTotal, "##")
        
        lblPackingTotal = CLng(Val(txtPacking.Text) + (Val(txtSumFeeTotal.Text) * Val(txtPackingPercent.Text) / 100))
        If Val(lblPackingTotal) <> 0 Then lblPackingTotal = Format(lblPackingTotal, "##")
        
        If chKTax = True Then
          ReDim Parameter(5) As Parameter
          Parameter(0) = GenerateInputParameter("@ValueGoodsDuty", adDouble, 8, ValueGoodsDuty)
          Parameter(1) = GenerateInputParameter("@ValueGoodsTax", adDouble, 8, ValueGoodsTax)
          Parameter(2) = GenerateInputParameter("@DiscountTotal", adDouble, 8, Val(lblDiscountTotal.Caption))
          Parameter(3) = GenerateInputParameter("@ServiceTotal", adDouble, 8, Val(lblServiceTotal.Caption))
          Parameter(4) = GenerateInputParameter("@CarryFeeTotal", adDouble, 8, Val(lblCarryFeeTotal.Caption))
          Parameter(5) = GenerateInputParameter("@PackingTotal", adDouble, 8, Val(lblPackingTotal.Caption))
          
          Set rctmp = RunParametricStoredProcedure2Rec("Get_DutyTax", Parameter)
          
          If Not (rctmp.BOF Or rctmp.EOF) Then
              lblTaxTotal.Caption = rctmp!TaxTotal
              LblDutyTotal.Caption = rctmp!DutyTotal
          End If
        Else
              lblTaxTotal.Caption = 0
              LblDutyTotal.Caption = 0
        End If
        lblSumPrice = Val(txtSumFeeTotal.Text) + Val(lblCarryFeeTotal.Caption) + Val(lblServiceTotal.Caption) + Val(lblPackingTotal.Caption) + Val(lblTaxTotal.Caption) + Val(LblDutyTotal.Caption) - Val(lblDiscountTotal.Caption)
''===

        ReDim Parameter(1) As Parameter
        Parameter(0) = GenerateInputParameter("@SumPrice", adBigInt, 8, Val(lblSumPrice.Caption))
        Parameter(1) = GenerateOutputParameter("@Remain", adInteger, 4)
        
        AutoDiscountValue = RunParametricStoredProcedure2String("Get_RoundSumPrice", Parameter)
        If Val(lblSumPrice.Caption) <> 0 Then
            lblSumPrice.Caption = Val(lblSumPrice.Caption) - AutoDiscountValue
        End If
        If AutoDiscountValue <> 0 Then lblDiscountTotal.Caption = Format(Val(lblDiscountTotal.Caption) + AutoDiscountValue, "##")
            
''===
        If lblPayFactorTotal.Visible = True And Val(lblPayFactorTotal.Caption) > 0 Then
            LblRemain.Caption = Val(lblSumPrice.Caption) - Val(lblPayFactorTotal.Caption)
            If Val(LblRemain.Caption) > 0 Then
                LblRemain.Caption = "„«‰œÂ: " & Format(LblRemain.Caption, "#,## —Ì«·")
            Else
                LblRemain.Caption = ""
            End If
        Else
            LblRemain.Caption = ""
        End If
        lblSumPrice.Tag = lblSumPrice.Caption
        lblSumPrice.Caption = Format(lblSumPrice, "#,## —Ì«·")

    Case Else
       If strCategory = "07" And SplitFlag = False Then
            lblServiceTotal = CLng(Val(txtSumFeeTotal.Text) * ServiceRate / 100)
            lblDiscountTotal = CLng((Val(txtSumFeeTotal.Text) * Val(txtDiscountPercent.Text) / 100) + ValueGoodDiscount + Val(txtDiscount.Text))
            If Val(lblDiscountTotal) <> 0 Then lblDiscountTotal = Format(lblDiscountTotal, "##")
            lblCarryFeeTotal = CLng(Val(txtCarryFee.Text) + (Val(txtSumFeeTotal.Text) * Val(txtCarryFeePercent.Text) / 100)) '+ Val(txtCarryFee.Text)
            If Val(lblCarryFeeTotal) <> 0 Then lblCarryFeeTotal = Format(lblCarryFeeTotal, "##")
            lblPackingTotal = CLng(Val(txtPacking.Text) + (Val(txtSumFeeTotal.Text) * Val(txtPackingPercent.Text) / 100))
            If Val(lblPackingTotal) <> 0 Then lblPackingTotal = Format(lblPackingTotal, "##")
'            LblDutyTotal.Caption = ValueDuty
'            LblTaxTotal.Caption = ValueTax
            lblSumPrice.Caption = Val(txtSumFeeTotal.Text) + Val(lblCarryFeeTotal.Caption) + Val(lblServiceTotal.Caption) + Val(lblPackingTotal.Caption) + Val(lblTaxTotal.Caption) + Val(LblDutyTotal.Caption) - Val(lblDiscountTotal.Caption)
''===

            ReDim Parameter(1) As Parameter
            Parameter(0) = GenerateInputParameter("@SumPrice", adBigInt, 8, Val(lblSumPrice.Caption))
            Parameter(1) = GenerateOutputParameter("@Remain", adInteger, 4)
            
            AutoDiscountValue = RunParametricStoredProcedure2String("Get_RoundSumPrice", Parameter)
            If Val(lblSumPrice.Caption) <> 0 Then
                lblSumPrice.Caption = Val(lblSumPrice.Caption) - AutoDiscountValue
            End If
            
            lblDiscountTotal.Caption = Format(Val(lblDiscountTotal.Caption) + AutoDiscountValue, "##")
            
''===
            lblSumPrice.Tag = lblSumPrice.Caption
            lblSumPrice.Caption = Format(lblSumPrice.Caption, "#,##")
        
            ReDim Parameter(3) As Parameter
            Parameter(0) = GenerateInputParameter("@No", adBigInt, 8, Val(txtNo.Text))
            Parameter(1) = GenerateInputParameter("@Status", adInteger, 4, mvarStatus)
            Parameter(2) = GenerateInputParameter("@AccountYear", adSmallInt, 2, AccountYear)
            Parameter(3) = GenerateInputParameter("@Branch", adInteger, 4, CurrentBranch)
            
            Set rctmp = RunParametricStoredProcedure2Rec("Get_tFacM_By_No_Status", Parameter)
            
            If Not (rctmp.BOF Or rctmp.EOF) Then
                   
                txtDiscount.Text = 0
                
                If Not IsNull(rctmp!Customer) Then
                    lblCustomer.Tag = rctmp!Customer
                    UpdatelblCustomer           ''''
                End If
                
                                  
                
                If Not IsNull(rctmp!Date) Then
                   Me.txtDate.Text = rctmp!Date
                   Me.txtDate.Tag = rctmp!Date
                End If
                
                
                If Not IsNull(rctmp!Recursive) Then
                    Me.txtRecursive = rctmp!Recursive
                End If
                
                If Not IsNull(rctmp!ServePlace) Then
                    intSumOfCurrentServePlaces = rctmp!ServePlace
                    
                End If
            End If
        
        Else
               
            ReDim Parameter(3) As Parameter
            Parameter(0) = GenerateInputParameter("@No", adBigInt, 8, Val(txtNo.Text))
            Parameter(1) = GenerateInputParameter("@Status", adInteger, 4, mvarStatus)
            Parameter(2) = GenerateInputParameter("@AccountYear", adSmallInt, 2, AccountYear)
            Parameter(3) = GenerateInputParameter("@Branch", adInteger, 4, CurrentBranch)
            
            Set rctmp = RunParametricStoredProcedure2Rec("Get_tFacM_By_No_Status", Parameter)
            
            If Not (rctmp.BOF Or rctmp.EOF) Then
                   
                txtDiscount.Text = 0
                
                If Not IsNull(rctmp!Customer) Then
                    lblCustomer.Tag = rctmp!Customer
                    UpdatelblCustomer           ''''
                End If
                
                If Not IsNull(rctmp!DiscountTotal) Then
                    lblDiscountTotal.Caption = rctmp!DiscountTotal
                    If Val(txtDiscountPercent.Text) = 0 Then   ' Not Customer Discount
                        txtDiscount.Text = rctmp!DiscountTotal - TmpGoodDiscount - rctmp!RoundDiscount
                        'txtDiscountPercent.Text = (Val(lblDiscountTotal) - rctmp!RoundDiscount - ValueGoodDiscount) * 100 / Val(txtSumFeeTotal.Text)  '(Val(lblDiscountTotal) - rctmp!RoundDiscount -
                    End If
                    
                End If
                
               
                If Not IsNull(rctmp!CarryFeeTotal) Then
                    lblCarryFeeTotal.Caption = rctmp!CarryFeeTotal
                    txtCarryFee.Text = rctmp!CarryFeeTotal
                End If
                
                If Not IsNull(rctmp!sumPrice) Then
                    lblSumPrice.Caption = rctmp!sumPrice
                    If lblPayFactorTotal.Visible = True And Val(lblPayFactorTotal.Caption) > 0 Then
                        LblRemain.Caption = Val(lblSumPrice.Caption) - Val(lblPayFactorTotal.Caption)
                        If Val(LblRemain.Caption) > 0 Then
                            LblRemain.Caption = "„«‰œÂ: " & Format(LblRemain.Caption, "#,## —Ì«·")
                        Else
                            LblRemain.Caption = ""
                        End If
                    Else
                        LblRemain.Caption = ""
                    End If
                    lblSumPrice.Tag = lblSumPrice.Caption
                    lblSumPrice.Caption = Format(lblSumPrice.Caption, "#,## —Ì«·")
                End If
                
                If Not IsNull(rctmp!ServiceTotal) Then
                   lblServiceTotal.Caption = rctmp!ServiceTotal
                    ServiceRate = rctmp!ServiceTotal * 100 / (Val(lblSumPrice.Tag) - rctmp!ServiceTotal - rctmp!PackingTotal - rctmp!CarryFeeTotal - rctmp!DutyTotal - rctmp!TaxTotal + rctmp!DiscountTotal)  '- rctmp!RoundDiscount
                End If
                
                If Not IsNull(rctmp!TaxTotal) Then
                   lblTaxTotal.Caption = rctmp!TaxTotal
                End If
                
                If Not IsNull(rctmp!DutyTotal) Then
                   LblDutyTotal.Caption = rctmp!DutyTotal
                End If
                
                If Not IsNull(rctmp!PackingTotal) Then
                   lblPackingTotal.Caption = rctmp!PackingTotal
                   txtPacking.Text = rctmp!PackingTotal
                End If
                
                If Not IsNull(rctmp!Date) Then
                   Me.txtDate.Text = rctmp!Date
                   Me.txtDate.Tag = rctmp!Date
                End If
                
                
                If Not IsNull(rctmp!Recursive) Then
                    Me.txtRecursive = rctmp!Recursive
                End If
                
                If Not IsNull(rctmp!ServePlace) Then
                    intSumOfCurrentServePlaces = rctmp!ServePlace
                    UpdatelblServePlace
                End If
               
                If Not IsNull(rctmp!GuestNo) Then
                    TxtGuestNo.Text = rctmp!GuestNo
                Else
                    TxtGuestNo.Text = "0"
                End If
               
            End If
        End If
        If Me.txtRecursive = 1 Then
            fwlblRecursive.Visible = True
            LblRemain.Caption = ""
            If (mvarCurUserNo = dblFichUser And ClsFormAccess.RefferInvoice = True) Or (ClsFormAccess.RefferedAllStationsFactors = True) Then
                MyFormAddEditMode = RefferedMode
            End If
        '    fwScrollTextCust.Visible = False
        Else
            fwlblRecursive.Visible = False
        '    fwScrollTextCust.Visible = True
        End If
               
        rctmp.Close
        BeforEditInvoice.Customer = Me.lblCustomer.Tag
        BeforEditInvoice.DiscountTotal = Val(Me.lblDiscountTotal)
        BeforEditInvoice.CarryFeeTotal = Val(Me.lblCarryFeeTotal)
        BeforEditInvoice.PackingTotal = Val(Me.lblPackingTotal)
        BeforEditInvoice.ServiceTotal = Val(Me.lblServiceTotal)
        BeforEditInvoice.GuestNo = Val(TxtGuestNo.Text)
        BeforEditInvoice.TaxTotal = Val(Me.lblTaxTotal)
        BeforEditInvoice.DutyTotal = Val(Me.LblDutyTotal)
 
  End Select
'  UpdatelblCustomer
    If MyFormAddEditMode = ViewMode And mvarStatus = Invoice Then
        Dim NewSumprice As Long
        
        Dim Parameters(4) As Parameter
    
        Parameters(0) = GenerateInputParameter("@No", adBigInt, 8, Val(txtNo.Text))
        Parameters(1) = GenerateInputParameter("@AccountYear", adSmallInt, 2, AccountYear)
        Parameters(2) = GenerateInputParameter("@Status", adInteger, 4, 2)
        Parameters(3) = GenerateInputParameter("@Branch", adInteger, 4, CurrentBranch)
        Parameters(4) = GenerateOutputParameter("@SumPrice", adBigInt, 8)
    
        NewSumprice = RunParametricStoredProcedure2String("Get_tFacm_Sumprice", Parameters)
        LblTip.Caption = ""
        If NewSumprice - Val(lblSumPrice.Tag) > 0 Then
            If mVarAccessLevel = 1 Then: LblTip.Caption = NewSumprice - Val(lblSumPrice.Tag) & " —Ì«·"
        End If
  End If
  
    If (MyFormAddEditMode = EditMode Or MyFormAddEditMode = RefferedMode Or MyFormAddEditMode = ManipulateMode) And frmAccess.ReturnAccess = True Then
       LblTip.Caption = ""
       Exit Sub
    End If

    Call CashCloseStatus
    If clsInvoiceValue.ShowInvoiceMenu = True Then
        frmShowInvoiceMenu.UpdateLblValue
    End If

    If IsFarabin = True Then ShowMonitor 1
    If clsInvoiceValue.ShowLogo = True Then
        frmShowLogo.UpdateLblValue
    End If
    If mvarStatus = Invoice And MyFormAddEditMode = ViewMode Then
        CustomerDisplay Val(lblSumPrice.Tag), clsArya.CustomerDisplayName
    End If

End Sub


Public Function CodeCount() As Boolean

    If MaxRowFlexGrid <= 1 Then  'Or lblSumPrice = 0
        frmMsg.fwlblMsg.Caption = " . ›Ì‘ Œ«·Ì «”  Ê À»  ‰„Ì ê—œœ"
        frmMsg.fwBtn(0).Visible = False
        frmMsg.fwBtn(1).ButtonType = flwButtonOk
        frmMsg.fwBtn(1).Caption = "ﬁ»Ê·"
        frmMsg.Show vbModal
        CodeCount = False
        mvarEmpty = True
    Else
        CodeCount = True
        mvarEmpty = False
    End If

End Function

Public Sub Number()
    Dim Rst As New ADODB.Recordset
    ReDim Parameter(2) As Parameter
    Parameter(0) = GenerateInputParameter("@Status", adInteger, 4, mvarStatus)
    Parameter(1) = GenerateInputParameter("@AccountYear", adSmallInt, 2, AccountYear)
    Parameter(2) = GenerateInputParameter("@Branch", adInteger, 4, CurrentBranch)
    
    Set Rst = RunParametricStoredProcedure2Rec("Get_New_FacM_No", Parameter)
    txtNo.Text = Rst!No
    FWLedTemp.Value = Rst!tempNo
    Rst.Close: Set Rst = Nothing

    If clsStation.Frame_Printers = True Then
        Timer_Printers_Timer
    End If

End Sub

Public Sub ArrowkeyStatusbar(intDirection As EnumDirection, Optional CurrentintSerialNo As Double)                'Display 5 Last Fich
    
    Dim L_Rst As New ADODB.Recordset
    Dim j As Integer
    Dim str1 As String
    
    ReDim Parameter(6) As Parameter
    Parameter(0) = GenerateInputParameter("@No", adBigInt, 8, CurrentintSerialNo)
    Parameter(1) = GenerateInputParameter("@Direction", adInteger, 4, intDirection)
    Parameter(2) = GenerateInputParameter("@User", adInteger, 4, mvarCurUserNo)
    Parameter(3) = GenerateInputParameter("@Status", adInteger, 4, mvarStatus)
    Parameter(4) = GenerateInputParameter("@Date", adWChar, 10, txtDate.Text)
    Parameter(5) = GenerateInputParameter("@AccountYear", adSmallInt, 2, AccountYear)
    Parameter(6) = GenerateInputParameter("@Branch", adSmallInt, 2, CurrentBranch)
    
    Set L_Rst = RunParametricStoredProcedure2Rec("NavigateFacM", Parameter)
    
    For i = 1 To 7
        Me.StatusBar.Panels(i).Tag = ""
        Me.StatusBar.Panels(i).Text = ""
    Next i
    
    If Not (L_Rst.EOF = True And L_Rst.BOF = True) Then
        i = 7
        Do While Not (L_Rst.EOF = True)
            str1 = L_Rst.Fields("SumPrice")
            i = i - 1
            If i = 1 Then
                Exit Do
            End If
            If i <> 1 And i <> 7 Then
                Me.StatusBar.Panels(i).Tag = L_Rst.Fields("No").Value
                If L_Rst!BascoleNo = 0 Then
                    If clsStation.TemporaryNo = True Then
                        Me.StatusBar.Panels(i).Text = IIf(Right(Str(L_Rst.Fields("TempNo")), 3) <> "000", "", "1") & Right(Str(L_Rst.Fields("TempNo")), 3) & ")" & str1
                    Else
                        Me.StatusBar.Panels(i).Text = IIf(Right(Str(L_Rst.Fields("No")), 3) <> "000", "", "1") & Right(Str(L_Rst.Fields("No")), 3) & ")" & str1
                    End If
                Else
                    If clsStation.TemporaryNo = True Then
                        Me.StatusBar.Panels(i).Text = Right(Str(L_Rst.Fields("BascoleNo")), 1) & ")" & IIf(Right(Str(L_Rst.Fields("TempNo")), 3) <> "000", "", "1") & Right(Str(L_Rst.Fields("TempNo")), 3) & ")" & str1
                    Else
                        Me.StatusBar.Panels(i).Text = Right(Str(L_Rst.Fields("BascoleNo")), 1) & ")" & IIf(Right(Str(L_Rst.Fields("No")), 3) <> "000", "", "1") & Right(Str(L_Rst.Fields("No")), 3) & ")" & str1
                    End If
                End If
                If (Val(strCategory) >= 0 And Val(strCategory) <= 7) Then
                    If L_Rst.Fields("OrderType").Value = 1 Then
                        Me.StatusBar.Panels(i).Picture = ImageList1.ListImages(3).Picture  'LoadPicture(App.Path & "\Image\Key\Tel.Ico")
                    Else
                        Me.StatusBar.Panels(i).Picture = ImageList1.ListImages(1).Picture 'LoadPicture(App.Path & "\Image\Key\Hozor.Ico")
                    End If
                End If
             End If
            L_Rst.MoveNext
        Loop
    End If
    If L_Rst.State = adStateOpen Then L_Rst.Close: Set L_Rst = Nothing
End Sub

Public Sub ValueBtnMenu()

'On Error Resume Next
'
'
'    SSTab1.Tab = 0
'
'    ReDim Parameter(1) As Parameter
'    Parameter(0) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
'    Parameter(1) = GenerateInputParameter("@StationId", adInteger, 4, clsArya.StationNo)
'
'    If rctmp.State <> 0 Then rctmp.Close
'    Set rctmp = RunParametricStoredProcedure2Rec("Get_tPocketPCGroup", Parameter)
'    Dim ii As Integer
'    ii = 0
'    While rctmp.EOF <> True
'        If IsNull(rctmp.Fields("StationId").Value) <> True Then
'            If ii <= SSTab1.Tabs - 1 Then
'                SSTab1.TabCaption(ii) = rctmp.Fields("Description").Value
'                ii = ii + 1
'            End If
'        End If
'        rctmp.MoveNext
'    Wend
'
'    Set rctmp = Nothing
'
'    ReDim Parameter(1) As Parameter
'    Parameter(0) = GenerateInputParameter("@StationId", adInteger, 4, clsArya.StationNo)
'    Parameter(1) = GenerateInputParameter("@FactorType", adInteger, 4, EnumFactorType.Invoice)
'
'    Set rctmp = RunParametricStoredProcedure2Rec("GetPictureButton", Parameter)
'
'    Do While Not rctmp.EOF
'        If Not IsNull(rctmp.Fields("PicturePath")) Then
'           ' BtnMenu(rctmp.Fields("BtnNum")).BackStyle = fmBackStyleOpaque
'            If rctmp.Fields("PicturePath") <> "" Then
'                BtnMenu(rctmp.Fields("BtnNum")).Picture = LoadPicture(App.Path & rctmp.Fields("PicturePath"))
'            End If
'           ' BtnMenu(rctmp.Fields("BtnNum")).PicturePosition = fmPicturePositionAboveCenter
'            BtnMenu(rctmp.Fields("BtnNum")).Caption = rctmp.Fields("NameDisp")
'           ' BtnMenu(rctmp.Fields("BtnNum")).WordWrap = False  ' Single Line If Has Picture
'
'
'        End If
'       ' BtnMenu(rctmp.Fields("BtnNum")).WordWrap = True  ' Double Line If No Picture
'        rctmp.MoveNext
'    Loop
'    rctmp.Cancel
'
'    ReDim Parameter(2) As Parameter
'
'    Parameter(0) = GenerateInputParameter("@StationId", adInteger, 4, clsArya.StationNo)
'    Parameter(1) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
'    Parameter(2) = GenerateInputParameter("@FactorType", adInteger, 4, EnumFactorType.Invoice)
'
'    Set rctmp = RunParametricStoredProcedure2Rec("GetButtonMenu", Parameter)
'
'     For i = 0 To BtnMenu.Count - 1
'        BtnMenu(i).Tag = ""
'    Next i
'
'    Do While Not rctmp.EOF
'        If Not IsNull(rctmp.Fields("BtnNum")) And rctmp.Fields("BtnNum") < BtnMenu.Count Then
'            BtnMenu(rctmp.Fields("BtnNum")).Tag = BtnMenu(rctmp.Fields("BtnNum")).Tag & rctmp.Fields("Code") & ";"
''            FWRealButton1(rctmp.Fields("BtnNum")).Tag = BtnMenu(rctmp.Fields("BtnNum")).Tag & rctmp.Fields("Code") & ";"
'            If Not IsNull(rctmp.Fields("NameDisp")) Then
'                BtnMenu(rctmp.Fields("BtnNum")).Caption = rctmp.Fields("NameDisp")
''                FWRealButton1(rctmp.Fields("BtnNum")).Caption = rctmp.Fields("NameDisp")
'            Else
'                BtnMenu(rctmp.Fields("BtnNum")).Caption = rctmp.Fields("Name")
''                FWRealButton1(rctmp.Fields("BtnNum")).Caption = rctmp.Fields("Name")
'            End If
'        End If
'        rctmp.MoveNext
'    Loop
'    rctmp.Cancel
'    For i = 1 To BtnMenu.Count - 1
'        If Len(BtnMenu(i).Tag) > 0 Then
'            BtnMenu(i).Tag = Left(BtnMenu(i).Tag, Len(BtnMenu(i).Tag) - 1)
'    '        FWRealButton1(i).Tag = Left(FWRealButton1(i).Tag, Len(FWRealButton1(i).Tag) - 1)
'            If BtnMenu(i).Tag = "" And BtnMenu(i).Caption = "" Then
'                BtnMenu(i).Enabled = False
'    '            FWRealButton1(i).Enabled = False
'
'            End If
'        Else
'            BtnMenu(i).Enabled = False
'        End If
'    Next i
'
'Exit Sub
'Err1:
'Resume Next
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
    
    MenuBarDescription
    ValueBtnMenu2
    SetBtnMenuPosition
    If SSTab1.TabsPerRow < 5 Then SSTab1.TabPicture(SSTab1.Tab) = ImageList1.ListImages(4).Picture
    SSTab1.BackColor = Invoice_BackColorForm
End Sub
Public Sub MenuBarDescription()
Dim ii As Long
On Error Resume Next
'    For ii = 2 To 5
'        MenuBar.Panels(ii).Text = ""
'        MenuBar.Panels(ii).Tag = 0
'    Next
    ReDim Parameter(1) As Parameter
    Parameter(0) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
    Parameter(1) = GenerateInputParameter("@StationId", adInteger, 4, clsArya.StationNo)
    
    If rctmp.State <> 0 Then rctmp.Close
    Set rctmp = RunParametricStoredProcedure2Rec("Get_tPocketPCGroup", Parameter)
    ii = 0
'    While rctmp.EOF <> True
'        If IsNull(rctmp.Fields("StationId").Value) <> True Then
'            If index = 0 Then
'                If ii <= 3 And rctmp.Fields("PocketPCGroupCode").Value < 5 Then
'                    MenuBar.Panels(ii + 2).Text = rctmp.Fields("Description").Value
'                    MenuBar.Panels(ii + 2).Tag = ii + 1
'                    ii = ii + 1
'                End If
'            Else
'                If ii <= 3 And rctmp.Fields("PocketPCGroupCode").Value >= 5 And rctmp.Fields("PocketPCGroupCode").Value < 9 Then
'                    MenuBar.Panels(ii + 2).Text = rctmp.Fields("Description").Value
'                    MenuBar.Panels(ii + 2).Tag = ii + 1
'                    ii = ii + 1
'                End If
'            End If
'        End If
'        rctmp.MoveNext
'    Wend
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
    If SSTab1.TabsPerRow > 5 Then SSTab1.Font.size = SSTab1.Font.size - 2: SSTab1.TabHeight = 600 Else SSTab1.TabHeight = 500
    SSTab1.Height = SSTab1.TabHeight + 50
    Set rctmp = Nothing
        
    
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
Public Sub SetBtnMenuPosition()
'txtTxtWidth = 0
'TxtHeight = 0
    If formloadFlag = False Then Exit Sub
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
    SSTab1.left = 100
    lastPosition.x = 120
    lastPosition.y = SSTab1.Height + 50
    
    Dim mm As Long
   ' mm = (RowTab * 4) + Column - 2 ' Index of group menu
    mm = SSTab1.Tab ' Index of group menu
   
    For i = (mm * BtnMenuPerFrame) + 1 To (mm * BtnMenuPerFrame) + BtnMenuPerFrame
'        Debug.Print (i - 1) Mod 4
'        If (i - 1) Mod 4 = 0 And i > 1 Then
        If lastPosition.x > frameMenu.Width - (Val(BtnMenu(1).Width)) Then
            lastPosition.x = 120
            lastPosition.y = lastPosition.y + Val(BtnMenu(1).Height)
        End If
        BtnMenu(i).Width = BtnMenu(1).Width
        BtnMenu(i).Height = BtnMenu(1).Height
        BtnMenu(i).left = lastPosition.x
        BtnMenu(i).top = lastPosition.y
'        BtnMenu(mm).Width = Val(TxtWidth)
'        BtnMenu(mm).Height = Val(TxtHeight)
        lastPosition.x = lastPosition.x + Val(BtnMenu(1).Width)
    Next
End Sub

Public Sub ValueBtnMenu2()
    
    Dim i As Long
    For i = 1 To MaxBtnMenu
        BtnMenu(i).Visible = False
        BtnMenu(i).Enabled = False
        BtnMenu(i).Tag = ""
        BtnMenu(i).Caption = ""
        BtnMenu(i).Picture = LoadPicture("")
    Next
    Dim mm As Long
'    mm = (RowTab * 4) + Column - 2  ' Index of group menu
    mm = SSTab1.Tab  ' Index of group menu
    
    For i = (mm * BtnMenuPerFrame) + 1 To (mm * BtnMenuPerFrame) + BtnMenuPerFrame
        BtnMenu(i).Visible = True
        Select Case mm
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

        Select Case clsStation.Language
            Case EnumLanguage.Farsi
                BtnMenu(i).Font.Name = Invoice_FontMenuName
                BtnMenu(i).Font.size = Val(Invoice_FontMenuSize)
                BtnMenu(i).Font.Bold = Invoice_FontMenuBold
            Case EnumLanguage.English
                BtnMenu(i).Font = "TimesNewRoman"
                BtnMenu(i).Font.size = 10
                BtnMenu(i).Font.Bold = True
       End Select
    Next
    
    ReDim Parameter(1) As Parameter
    Parameter(0) = GenerateInputParameter("@StationId", adInteger, 4, clsArya.StationNo)
    Parameter(1) = GenerateInputParameter("@FactorType", adInteger, 4, EnumFactorType.Invoice)
    
    Set rctmp = RunParametricStoredProcedure2Rec("GetPictureButton", Parameter)
    
    Do While Not rctmp.EOF
        If Not IsNull(rctmp.Fields("PicturePath")) And rctmp.Fields("BtnNum") <= BtnMenu.Count Then
           ' BtnMenu(rctmp.Fields("BtnNum")).BackStyle = fmBackStyleOpaque
            If rctmp.Fields("PicturePath") <> "" And rctmp.Fields("BtnNum") >= (mm * BtnMenuPerFrame) + 1 And rctmp.Fields("BtnNum") <= (mm * BtnMenuPerFrame) + BtnMenuPerFrame Then
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
    
    Parameter(0) = GenerateInputParameter("@StationId", adInteger, 4, clsArya.StationNo)
    Parameter(1) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
    Parameter(2) = GenerateInputParameter("@FactorType", adInteger, 4, EnumFactorType.Invoice)
    
    Set rctmp = RunParametricStoredProcedure2Rec("GetButtonMenu", Parameter)
    
    Do While Not rctmp.EOF
        If Not IsNull(rctmp.Fields("BtnNum")) And rctmp.Fields("BtnNum") <= BtnMenu.Count And rctmp.Fields("BtnNum") >= (mm * BtnMenuPerFrame) + 1 And rctmp.Fields("BtnNum") <= (mm * BtnMenuPerFrame) + BtnMenuPerFrame Then
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
    For i = (mm * BtnMenuPerFrame) + 1 To (mm * BtnMenuPerFrame) + BtnMenuPerFrame
        If i <= BtnMenu.Count Then
           If Len(BtnMenu(i).Tag) > 0 Then
               BtnMenu(i).Tag = left(BtnMenu(i).Tag, Len(BtnMenu(i).Tag) - 1)
        '        FWRealButton1(i).Tag = Left(FWRealButton1(i).Tag, Len(FWRealButton1(i).Tag) - 1)
               If BtnMenu(i).Tag = "" And BtnMenu(i).Caption = "" Then
                   BtnMenu(i).Enabled = False
        '            FWRealButton1(i).Enabled = False
              
               End If
           Else
               BtnMenu(i).Enabled = False
           End If
        End If
    Next i

Exit Sub
Err1:
Resume Next
End Sub

Public Sub ValueLabel()
Select Case Val(txtRecursive.Text)
    Case 1:
        fwlblRecursive.Visible = True
        
    '    fwScrollTextCust.Visible = False
    Case Else:
        fwlblRecursive.Visible = False
    '    fwScrollTextCust.Visible = True
End Select
        

End Sub

Public Sub PanelClick(Panel As Integer)

    Dim j As Integer
    Dim str1 As String
    
    If Not (Me.StatusBar.Panels(Panel).Enabled) Then
        Exit Sub
    End If
    Select Case Panel
        Case 1
            If Val(Me.StatusBar.Panels(6).Tag) = 0 Or Me.StatusBar.Panels(6).Tag < 6 Then Exit Sub
            For i = 2 To 6
                Me.StatusBar.Panels(i).Bevel = sbrInset
                Me.StatusBar.Panels(i).Enabled = True
            Next i
            If Val(Me.StatusBar.Panels(6).Tag) <> 0 Then
                ArrowkeyStatusbar PreviousRecord, Val(Me.StatusBar.Panels(6).Tag)
            Else
                ArrowkeyStatusbar FirstRecord
            End If
        Case 7
            If Val(Me.StatusBar.Panels(2).Tag) = 0 Then Exit Sub
            For i = 2 To 6
                Me.StatusBar.Panels(i).Bevel = sbrInset
                Me.StatusBar.Panels(i).Enabled = True
            Next i
            
            If Val(Me.StatusBar.Panels(2).Tag) <> 0 Then
                ArrowkeyStatusbar NextRecord, Val(Me.StatusBar.Panels(2).Tag)
            Else
                ArrowkeyStatusbar LastRecord
            End If
            
        Case Else
        
            If Val(Me.StatusBar.Panels(Panel).Tag) <> 0 Then
                Me.StatusBar.Enabled = False
                str1 = Me.StatusBar.Panels(Panel).Tag
                txtNo.Text = str1
                MyFormAddEditMode = ViewMode   'view Mode
                SetFirstToolBar
                GetDataDetail
                RefreshLables
                For i = 2 To 6
                    Me.StatusBar.Panels(i).Bevel = sbrInset
                    Me.StatusBar.Panels(i).Enabled = True
                Next i
''                Edit
                Me.StatusBar.Enabled = True
            End If
    End Select
    Exit Sub
    
ErrorHandler:
    Exit Sub
End Sub

Public Function StatusPic(ServePlace As Integer) As String
    Select Case ServePlace
        Case 1:
            StatusPic = App.Path & "\Image\HozorSaloon.ICO"
        Case 2:
            StatusPic = App.Path & "\Image\HozorErsal.ICO"
        Case 3:
            StatusPic = App.Path & "\Image\TelSaloon.ICO"
        Case 4:
            StatusPic = App.Path & "\Image\TelErsal.ICO"
    End Select
End Function


Public Sub FindCust()
   
    If clsArya.Customers = True Then
        If DropDownFlag = False Then
            On Error GoTo ErrorHandler
            frmFindCust.Show vbModal
            Call_RealNumber = ""
            If mvarcode > 0 Then
                lblCustomer.Tag = mvarcode
                mvarcode = 0
                mVarOrderType = mvarPublicOrderType
                mvarPublicOrderType = inPerson
            Else
                lblCustomer.Tag = -1
                mvarPublicOrderType = inPerson
                mVarOrderType = inPerson
            End If
            If mVarOrderType = ByPhone Then
               If clsStation.Language = Farsi Then
                    LblOrder.Caption = " ·›‰Ì"
               Else
                    LblOrder.Caption = "By phone"
               End If
            Else
            If clsStation.Language = Farsi Then
               LblOrder.Caption = "Õ÷Ê—Ì"
            Else
                LblOrder.Caption = "Inside"
            End If
            End If
            For i = 0 To cmbServePlace.ListCount - 1
                If mvarServePlace = cmbServePlace.ItemData(i) Then
                    cmbServePlace.ListIndex = i
                    Exit For
                End If
            Next i
            UpdatelblCustomer
            UpdatelblServePlace
            RefreshLables
       End If
     Else
                    
        frmDisMsg.lblMessage.Caption = " ‘„« »Â «Ì‰ «„ﬂ«‰ œ” —”Ì ‰œ«—Ìœ "
        frmDisMsg.Timer1.Interval = 1000
        frmDisMsg.Timer1.Enabled = True
        frmDisMsg.Show vbModal
       
     End If
Exit Sub

ErrorHandler:
    frmFindCust.txtMembershipId.Text = CreditCode
    VarActForm = "frmInvoice"
End Sub
Public Function GetGoodBarcode(Code As String)
    
    Dim ReturnValue As Boolean
    ReturnValue = False
    If Code = "" Then Exit Function
    
    ReDim Parameter(1) As Parameter
    Parameter(0) = GenerateInputParameter("@Barcode", adVarWChar, 50, Code)
    Parameter(1) = GenerateInputParameter("@GoodCode", adInteger, 4, 0)
    Set rctmp = RunParametricStoredProcedure2Rec("Get_Good_Barcode_Check", Parameter)
        
    If (rctmp.BOF Or rctmp.EOF) Then
        frmDisMsg.lblMessage.Caption = " «Ì‰ »«—ﬂœ œ— ò«·«Â«  ⁄—Ì› ‰‘œÂ «”  "
        frmDisMsg.Timer1.Interval = 1000
        frmDisMsg.Timer1.Enabled = True
        frmDisMsg.Show vbModal
        rctmp.Close
        Exit Function
    End If
    ReDim Parameter(3) As Parameter
    Parameter(0) = GenerateInputParameter("@Barcode", adVarWChar, 50, Code)
    Parameter(1) = GenerateInputParameter("@StationId", adInteger, 4, clsArya.StationNo)
    Parameter(2) = GenerateInputParameter("@Flag", adBoolean, 1, 1)
    Parameter(3) = GenerateInputParameter("@AccountYear", adSmallInt, 2, AccountYear)
    Set rctmp = RunParametricStoredProcedure2Rec("Get_Good_Barcode", Parameter)

    If Not (rctmp.BOF Or rctmp.EOF) Then
        i = 0
        Do While Not rctmp.EOF
            i = i + 1
            mvarGoodCode = rctmp.Fields("Code")
            GoodCode = rctmp.Fields("Code")
            mvarUnitGood = rctmp.Fields("Unit")
            mvarGoodName = rctmp.Fields("Name")
            mvarGoodWeight = rctmp.Fields("Weight")
            mvarDisCount = rctmp.Fields("Discount")
            mvarInventoryNo = rctmp.Fields("InventoryNo")
            mvarMojodi = rctmp.Fields("Mojodi")
            If chKTax = True Then
               mvarDuty = True
               mvarTax = True
            Else
               mvarDuty = rctmp.Fields("DutySale")
               mvarTax = rctmp.Fields("TaxSale")
            End If
            Tafsili_2 = rctmp.Fields("Tafsili")
            If Val(lblCustomer.Tag) = -1 Or ServeChangeFlag = True Then  '' ‰—Œ œ” Ì «Ê·ÊÌ  œ«—œ »— ‰—Œ „‘ —òÌ‰
                If clsStation.PriceType = 1 Then
                   mvarSellPrice = rctmp.Fields("SellPrice").Value
                   mvarRate = 1
                ElseIf clsStation.PriceType = 2 Then
                   mvarSellPrice = rctmp.Fields("SellPrice2").Value
                   mvarRate = 2
                ElseIf clsStation.PriceType = 3 Then
                   mvarSellPrice = rctmp.Fields("SellPrice3").Value
                   mvarRate = 3
                ElseIf clsStation.PriceType = 4 Then
                   mvarSellPrice = rctmp.Fields("SellPrice4").Value
                   mvarRate = 4
                ElseIf clsStation.PriceType = 5 Then
                   mvarSellPrice = rctmp.Fields("SellPrice5").Value
                   mvarRate = 5
                ElseIf clsStation.PriceType = 6 Then
                   mvarSellPrice = rctmp.Fields("SellPrice6").Value
                   mvarRate = 6
                End If
            Else
                If clsStation.CustomerRate = 0 Then
                   mvarSellPrice = rctmp.Fields("SellPrice").Value
                   mvarRate = 1
                ElseIf clsStation.CustomerRate = 1 Then
                   mvarSellPrice = rctmp.Fields("SellPrice2").Value
                   mvarRate = 2
                ElseIf clsStation.CustomerRate = 2 Then
                   mvarSellPrice = rctmp.Fields("SellPrice3").Value
                   mvarRate = 3
                ElseIf clsStation.CustomerRate = 3 Then
                   mvarSellPrice = rctmp.Fields("SellPrice4").Value
                   mvarRate = 4
                ElseIf clsStation.CustomerRate = 4 Then
                   mvarSellPrice = rctmp.Fields("SellPrice5").Value
                   mvarRate = 5
                ElseIf clsStation.CustomerRate = 5 Then
                   mvarSellPrice = rctmp.Fields("SellPrice6").Value
                   mvarRate = 6
                End If

            End If
            InventoryNo = rctmp.Fields("InventoryNo").Value
            rctmp.MoveNext
        Loop
        If i > 1 Then
            frmDisMsg.lblMessage.Caption = "»Ì‘ «“ Ìò «‰»«— »—«Ì «Ì‰ ò«·« œ— «Ì‰ «Ì” ò«Â  ⁄—Ì› ‘œÂ «”  "
            frmDisMsg.Timer1.Interval = 1000
            frmDisMsg.Timer1.Enabled = True
            frmDisMsg.Show vbModal
            rctmp.Close
            Exit Function
        End If
        ReturnValue = True
    End If
    If ReturnValue = True And clsStation.RowMojodiControl = True Then
        DetailsString1 = ""
        With FlxDetail
            DetailsString1 = GenerateDetailsString3(DetailsString1, IIf(Val(lblNum.Caption) = 0, 1, Val(lblNum.Caption)), CStr(mvarGoodCode), CStr(mvarSellPrice), CStr(mvarDisCount), CStr(mvarRate), "", " ", CStr(mvarInventoryNo), "", 1, "")
        End With
        If MojodiControlFlag = True And mvarStatus = Invoice Then   'And FWMojodiControl.Visible = True Then
            If MyFormAddEditMode = AddMode Then
                ReDim Parameter(3) As Parameter
                Parameter(0) = GenerateInputParameter("@DetailsString", adVarWChar, 4000, DetailsString1)
               Parameter(1) = GenerateInputParameter("@DetailsString2", adVarWChar, 4000, "")
               Parameter(2) = GenerateInputParameter("@DetailsString3", adVarWChar, 4000, "")
               Parameter(3) = GenerateInputParameter("@DetailsString4", adVarWChar, 4000, "")
                Set Rst = RunParametricStoredProcedure2Rec("CheckPreSave", Parameter)
                If Not (Rst.BOF Or Rst.EOF) Then
                    mvarAddeditMode = MyFormAddEditMode
                    frmMojodiReduce.Show vbModal
                    If frmMojodiReduce.Result = False Then
                       ReturnValue = False
                       lblNum.Caption = ""
                    End If
                End If
            Else
                ReDim Parameter(5) As Parameter
                mvarNo = Val(txtNo.Text)
                Parameter(0) = GenerateInputParameter("@DetailsString", adVarWChar, 4000, DetailsString1)
                Parameter(1) = GenerateInputParameter("@Status", adInteger, 4, mvarStatus)
                Parameter(2) = GenerateInputParameter("@No", adBigInt, 8, mvarNo)
               Parameter(3) = GenerateInputParameter("@DetailsString2", adVarWChar, 4000, "")
               Parameter(4) = GenerateInputParameter("@DetailsString3", adVarWChar, 4000, "")
               Parameter(5) = GenerateInputParameter("@DetailsString4", adVarWChar, 4000, "")
                Set Rst = RunParametricStoredProcedure2Rec("CheckPreSave_Edit", Parameter)
                Dim ss As String
                If Not (Rst.BOF Or Rst.EOF) Then
                    mvarAddeditMode = MyFormAddEditMode
                    frmMojodiReduce.Show vbModal
                    If frmMojodiReduce.Result = False Then
                       ReturnValue = False
                       lblNum.Caption = ""
                    End If
                End If

            End If
        End If

    End If
    GetGoodBarcode = ReturnValue
    If ReturnValue = True And clsInvoiceValue.ShowPictureGood = True Then
       
        Dim Result As Integer
        ReDim Parameters(1) As Parameter
        Parameters(0) = GenerateInputParameter("@GoodCode", adBigInt, 8, GoodCode)
        Parameters(1) = GenerateOutputParameter("@Result", adInteger, 1)
        Result = RunParametricStoredProcedure2String("Get_CounttblTotal_GoodPic_byCode", Parameters)
        If Result = 1 Then
            frmShowPictureGood.Show vbModal
        End If
   End If
End Function

Public Function CheckDataBarcode()
    Dim ReturnValue As Boolean
    ReturnValue = False

    Dim Parameter(3) As Parameter
    Parameter(0) = GenerateInputParameter("@No", adBigInt, 8, Val(Mid(lblBarCode.Caption, 5, 9)))
    Parameter(1) = GenerateInputParameter("@Status", adInteger, 4, IIf(Val(Mid(lblBarCode.Caption, 4, 1)) = 0, 2, 10))
    Parameter(2) = GenerateInputParameter("@AccountYear", adSmallInt, 2, AccountYear)
    Parameter(3) = GenerateInputParameter("@Branch", adInteger, 4, CurrentBranch)
    Set rctmp = RunParametricStoredProcedure2Rec("Get_No_Status", Parameter)
    If Not (rctmp.BOF Or rctmp.EOF) Then
        ReturnValue = True
    End If
    CheckDataBarcode = ReturnValue

End Function
Public Function CheckExistBarcode()
    Dim ReturnValue As Boolean
    ReturnValue = False

    ReDim Parameter(0) As Parameter
    Parameter(0) = GenerateInputParameter("@BarcodeString", adWChar, 50, lblBarCode.Caption)
    Set rctmp = RunParametricStoredProcedure2Rec("Get_ExistChanceBarcode", Parameter)
    If Not (rctmp.BOF Or rctmp.EOF) Then
        ReturnValue = True
        intSerialNo = rctmp!intSerialNo
        If mvarDate = rctmp!Date Then
            TodayFlag = True
        Else
            TodayFlag = False
        End If
    End If
    CheckExistBarcode = ReturnValue
    rctmp.Close

End Function
Public Function CheckIsUsedBarcode()
    Dim ReturnValue As Boolean
    ReturnValue = False

    
    ReDim Parameter(0) As Parameter
    Parameter(0) = GenerateInputParameter("@intCreditSerial", adBigInt, 8, Val(lblBarCode.Caption))
    Set rctmp = RunParametricStoredProcedure2Rec("Get_tFacCredit_IsUsed", Parameter)
    If rctmp!IsUsed > 0 Then
        ReturnValue = True
    End If
    
    CheckIsUsedBarcode = ReturnValue
    rctmp.Close

End Function

Public Sub barcode()

If clsStation.BarcodeChance = True Then   'Chance  Barcode
    
    If CheckIsUsedBarcode = True Then
        frmMsg.fwlblMsg.Caption = " «Ì‰ ‘„«—Â »«—ﬂœ ﬁ»·« «” ›«œÂ ‘œÂ «”  "
        frmMsg.fwBtn(0).Visible = False
        frmMsg.fwBtn(1).ButtonType = flwButtonCancel
        frmMsg.fwBtn(1).Caption = "ﬁ»Ê·"
        frmMsg.Show vbModal
        lblBarCode = ""
        mvarbarcode = False
        Exit Sub
    End If
    If CheckExistBarcode = True Then
        If TodayFlag = True Then
            If MaxRowFlexGrid <> 1 Then
                ReDim Parameter(0) As Parameter
                Parameter(0) = GenerateInputParameter("@nvcBarCode", adVarWChar, 50, lblBarCode.Caption)
                Set rctmp = RunParametricStoredProcedure2Rec("Get_tFacm_By_nvcBarCode", Parameter)
                If Not (rctmp.EOF = True And rctmp.BOF = True) Then
                    frmMsg.fwlblMsg.Caption = " »«—ﬂœ  ﬂ—«—Ì „Ì »«‘œ"
                    frmMsg.fwBtn(0).Visible = False
                    frmMsg.fwBtn(1).ButtonType = flwButtonOk
                    frmMsg.fwBtn(1).Caption = "ﬁ»Ê·"
                    frmMsg.Show vbModal
                    
                End If
            Else
                frmPrize.intSerialNo = intSerialNo
                frmPrize.nvcPrizeBarCode = lblBarCode.Caption
                frmPrize.Show
                frmPrize.SetFocus
            End If
        Else
            frmMsg.fwlblMsg.Caption = "  «—ÌŒ »«—ﬂœ Ê«—œ ‘œÂ €Ì— «“ «„—Ê“ „Ì »«‘œ "
            frmMsg.fwBtn(0).Visible = False
            frmMsg.fwBtn(1).Visible = True
            frmMsg.fwBtn(1).ButtonType = flwButtonCancel
            frmMsg.fwBtn(1).Caption = "ﬁ»Ê·"
            frmMsg.fwBtn(1).Default = True
            frmMsg.fwBtn(0).Visible = False
            frmMsg.Show vbModal
        End If
    Else
        If clsStation.PriceChance = "" Then clsStation.PriceChance = "50000"
        Repeatbarcode = Int(Me.lblSumPrice.Tag / Val(clsStation.PriceChance))
        If Repeatbarcode > 0 And Repeatbarcode > ChanceBarcodeQuantity Then
            If ChanceBarcodeQuantity = 0 Then
                txtDescription.Text = ""
            End If
            textDescriptionFlag = True
            txtDescription.Text = txtDescription.Text + lblBarCode.Caption + "/"
            ChanceBarcodeQuantity = ChanceBarcodeQuantity + 1
        ElseIf Repeatbarcode = 0 Then
            frmMsg.fwlblMsg.Caption = "»«—ﬂœ Ê«—œ ‘œÂ À»  ‰‘œÂ «” "
            frmMsg.fwBtn(0).Visible = False
            frmMsg.fwBtn(1).ButtonType = flwButtonCancel
            frmMsg.fwBtn(1).Caption = "ﬁ»Ê·"
            frmMsg.Show vbModal
        Else
            frmMsg.fwlblMsg.Caption = "„ﬁœ«— Œ—Ìœ ﬂ„ — «“ „»·€  ⁄ÌÌ‰ ‘œÂ „Ì »«‘œ"
            frmMsg.fwBtn(0).Visible = False
            frmMsg.fwBtn(1).ButtonType = flwButtonCancel
            frmMsg.fwBtn(1).Caption = "ﬁ»Ê·"
            frmMsg.Show vbModal
        End If
    End If
   
ElseIf clsStation.AutoBarcode = True Then   'Auto & Good Barcode
    
    If Len(lblBarCode.Caption) = 12 Then
       lblBarCode.Caption = "0" + lblBarCode.Caption
    ElseIf Len(lblBarCode.Caption) = 13 Then
       lblBarCode.Caption = "0" + left(lblBarCode.Caption, 12)
    End If
    Select Case left(lblBarCode.Caption, 2)
    
        Case 62
        
            If GetGoodBarcode(lblBarCode) = True Then
                ChangeGoodquantity
            End If
        Case Is < 10
                
            Select Case left(lblBarCode.Caption, 3)
    
                Case EnumIncharge.Payk
                
                    Me.Hide
                    If ClsFormAccess.frmPayk = True Then
                        frmPayk.lblBarCode = Me.lblBarCode
                        frmPayk.Show
                    End If
                    
                Case 12     '«—”«·Ì '
                    If clsStation.DeliveryBarcodeDefault = 0 Then
                        MyFormAddEditMode = ViewMode
                        If CheckDataBarcode Then
                            If Val(Mid(lblBarCode.Caption, 4, 1)) = 1 Then
                                mvarStatus = EnumFactorType.Order
                                If clsStation.Language = Farsi Then
                                    LblInvoice.Caption = "”›«—‘"
                                    cmdPay.Caption = "(F8) ÕÊÌ·"
                                Else
                                    LblInvoice.Caption = "Order"
                                    cmdPay.Caption = "Recieve(F8)"
                                End If
                            End If
                            Me.txtNo.Text = Val(Mid(lblBarCode.Caption, 5, 9))
                            GetDataDetail
                            RefreshLables
                            SetFirstToolBar
                            If mvarStatus = Order Then
                                mdifrm.Toolbar1.Buttons(6).Enabled = False  'Add
                                mdifrm.Toolbar1.Buttons(7).Enabled = False  'Edit
                                mdifrm.Toolbar1.Buttons(8).Enabled = False   'Enter
                                mdifrm.Toolbar1.Buttons(15).Enabled = False   'Print
                            End If
                        Else
                            frmDisMsg.lblMessage.Caption = " . «Ì‰ »«—ﬂœ œ— ”Ì” „  ⁄—Ì› ‰‘œÂ «”  "
                            frmDisMsg.Timer1.Interval = 1000
                            frmDisMsg.Timer1.Enabled = True
                            frmDisMsg.Show vbModal
                        End If
                    Else
                        Me.Hide
                        If ClsFormAccess.frmPayk = True Then
                            frmPayk.lblBarCode = Me.lblBarCode
                            frmPayk.Show
                        End If
                    End If
                Case EnumIncharge.Garson
                    
                    Me.Hide
                    If ClsFormAccess.frmGarson = True Then
                        frmGarson.lblBarCode = Me.lblBarCode
                        frmGarson.Show
                    End If
                Case 26, 30   'Table & Table_Out
                    If clsStation.TableBarcodeDefault = 0 Then
                        MyFormAddEditMode = ViewMode
                        If CheckDataBarcode Then
                            If Val(Mid(lblBarCode.Caption, 4, 1)) = 1 Then
                                mvarStatus = EnumFactorType.Order
                                If clsStation.Language = Farsi Then
                                    LblInvoice.Caption = "”›«—‘"
                                    cmdPay.Caption = "(F8) ÕÊÌ·"
                                Else
                                    LblInvoice.Caption = "Order"
                                    cmdPay.Caption = "Recieve(F8)"
                                End If
                            End If
                            Me.txtNo.Text = Val(Mid(lblBarCode.Caption, 5, 9))
                            GetDataDetail
                            RefreshLables
                            SetFirstToolBar
                            If mvarStatus = Order Then
                                mdifrm.Toolbar1.Buttons(6).Enabled = False  'Add
                                mdifrm.Toolbar1.Buttons(7).Enabled = False  'Edit
                                mdifrm.Toolbar1.Buttons(8).Enabled = False   'Enter
                                mdifrm.Toolbar1.Buttons(15).Enabled = False   'Print
                            End If
                        Else
                            frmDisMsg.lblMessage.Caption = " . «Ì‰ »«—ﬂœ œ— ”Ì” „  ⁄—Ì› ‰‘œÂ «”  "
                            frmDisMsg.Timer1.Interval = 1000
                            frmDisMsg.Timer1.Enabled = True
                            frmDisMsg.Show vbModal
                        End If
                    Else
                        Me.Hide
                        If ClsFormAccess.frmGarson = True Then
                            frmGarson.lblBarCode = Me.lblBarCode
                            frmGarson.Show
                        End If
                    End If
                Case Else   ''''' ”«·‰
                    
                    MyFormAddEditMode = ViewMode
                    If CheckDataBarcode Then
                            If Val(Mid(lblBarCode.Caption, 4, 1)) = 1 Then
                                mvarStatus = EnumFactorType.Order
                                If clsStation.Language = Farsi Then
                                    LblInvoice.Caption = "”›«—‘"
                                    cmdPay.Caption = "(F8) ÕÊÌ·"
                                Else
                                    LblInvoice.Caption = "Order"
                                    cmdPay.Caption = "Recieve(F8)"
                                End If
                            End If
                            Me.txtNo.Text = Val(Mid(lblBarCode.Caption, 5, 9))
                            GetDataDetail
                            RefreshLables
                            SetFirstToolBar
                            If mvarStatus = Order Then
                                mdifrm.Toolbar1.Buttons(6).Enabled = False  'Add
                                mdifrm.Toolbar1.Buttons(7).Enabled = False  'Edit
                                mdifrm.Toolbar1.Buttons(8).Enabled = False   'Enter
                                mdifrm.Toolbar1.Buttons(15).Enabled = False   'Print
                            End If
                    Else
                        frmDisMsg.lblMessage.Caption = " . «Ì‰ »«—ﬂœ œ— ”Ì” „  ⁄—Ì› ‰‘œÂ «”  "
                        frmDisMsg.Timer1.Interval = 1000
                        frmDisMsg.Timer1.Enabled = True
                        frmDisMsg.Show vbModal
                    End If
            End Select
                    
           
        Case Else
    
                
            frmDisMsg.lblMessage.Caption = " «Ì‰ »«—ﬂœ œ— ”Ì” „ « Ê„« Ìò  ⁄—Ì› ‰‘œÂ «”  "
            frmDisMsg.Timer1.Interval = 1000
            frmDisMsg.Timer1.Enabled = True
            frmDisMsg.Show vbModal
                    
    End Select
Else
            
    If GetGoodBarcode(lblBarCode) = True Then
        ChangeGoodquantity
    End If

End If
beforeexit:
    lblBarCode = ""
    mvarbarcode = False

Exit Sub
Err1:
    FlxDetail.TextMatrix(FlxDetail.Row, 1) = ""
    frmDisMsg.lblMessage.Caption = " . œ— —Ê Ì‰ »«—ﬂœ «‘ﬂ«· ÊÃÊœ œ«—œ "
    frmDisMsg.Timer1.Interval = 1000
    frmDisMsg.Timer1.Enabled = True
    frmDisMsg.Show vbModal
    
    lblBarCode = ""
    mvarbarcode = False
End Sub
Public Function GetGoodCode(Code As Double)
    Dim ReturnValue As Boolean
    ReturnValue = False
    If Code = 0 Then Exit Function
    ReDim Parameter(4) As Parameter
    Parameter(0) = GenerateInputParameter("@Code", adInteger, 4, Code)
    Parameter(1) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
    Parameter(2) = GenerateInputParameter("@StationId", adInteger, 4, clsArya.StationNo)
    Parameter(3) = GenerateInputParameter("@Flag", adBoolean, 1, 1)
    Parameter(4) = GenerateInputParameter("@AccountYear", adSmallInt, 2, AccountYear)
    Set rctmp = RunParametricStoredProcedure2Rec("Get_Good_Code", Parameter)

    If Not (rctmp.BOF Or rctmp.EOF) Then
        i = 0
        Do While Not rctmp.EOF
            i = i + 1
            mvarGoodCode = rctmp.Fields("Code")
            mvarUnitGood = rctmp.Fields("Unit")
            mvarGoodName = rctmp.Fields("Name")
            mvarGoodWeight = rctmp.Fields("Weight")
            mvarDisCount = rctmp.Fields("Discount")
            mvarInventoryNo = rctmp.Fields("InventoryNo")
            bolMainGroup = rctmp.Fields("MainType")
            mvarMojodi = rctmp.Fields("Mojodi")
            If chKTax = True Then
               mvarDuty = True
               mvarTax = True
            Else
               mvarDuty = rctmp.Fields("DutySale")
               mvarTax = rctmp.Fields("TaxSale")
            End If
            Tafsili_2 = rctmp.Fields("Tafsili")
           If Val(lblCustomer.Tag) = -1 Or ServeChangeFlag = True Then  '' ‰—Œ œ” Ì «Ê·ÊÌ  œ«—œ »— ‰—Œ „‘ —òÌ‰
                If clsStation.PriceType = 1 Then
                   mvarSellPrice = rctmp.Fields("SellPrice").Value
                   mvarRate = 1
                ElseIf clsStation.PriceType = 2 Then
                   mvarSellPrice = rctmp.Fields("SellPrice2").Value
                   mvarRate = 2
                ElseIf clsStation.PriceType = 3 Then
                   mvarSellPrice = rctmp.Fields("SellPrice3").Value
                   mvarRate = 3
                ElseIf clsStation.PriceType = 4 Then
                   mvarSellPrice = rctmp.Fields("SellPrice4").Value
                   mvarRate = 4
                ElseIf clsStation.PriceType = 5 Then
                   mvarSellPrice = rctmp.Fields("SellPrice5").Value
                   mvarRate = 5
                ElseIf clsStation.PriceType = 6 Then
                   mvarSellPrice = rctmp.Fields("SellPrice6").Value
                   mvarRate = 6
                End If
            Else
                If clsStation.CustomerRate = 0 Then
                   mvarSellPrice = rctmp.Fields("SellPrice").Value
                   mvarRate = 1
                ElseIf clsStation.CustomerRate = 1 Then
                   mvarSellPrice = rctmp.Fields("SellPrice2").Value
                   mvarRate = 2
                ElseIf clsStation.CustomerRate = 2 Then
                   mvarSellPrice = rctmp.Fields("SellPrice3").Value
                   mvarRate = 3
                ElseIf clsStation.CustomerRate = 3 Then
                   mvarSellPrice = rctmp.Fields("SellPrice4").Value
                   mvarRate = 4
                ElseIf clsStation.CustomerRate = 4 Then
                   mvarSellPrice = rctmp.Fields("SellPrice5").Value
                   mvarRate = 5
                ElseIf clsStation.CustomerRate = 5 Then
                   mvarSellPrice = rctmp.Fields("SellPrice6").Value
                   mvarRate = 6
                End If
            
            End If
            InventoryNo = rctmp.Fields("InventoryNo").Value

           rctmp.MoveNext
        Loop
        If i > 1 Then
            frmDisMsg.lblMessage.Caption = "»Ì‘ «“ Ìò «‰»«— »—«Ì «Ì‰ ò«·« œ— «Ì‰ «Ì” ò«Â  ⁄—Ì› ‘œÂ «”  "
            frmDisMsg.Timer1.Interval = 1000
            frmDisMsg.Timer1.Enabled = True
            frmDisMsg.Show vbModal
            rctmp.Close
            Exit Function
        End If
        ReturnValue = True

    ElseIf (rctmp.BOF And rctmp.EOF) Then
        frmDisMsg.lblMessage.Caption = " «‰»«— »—«Ì «Ì‰ ò«·«  ⁄—Ì› ‰‘œÂ «”  "
        frmDisMsg.Timer1.Interval = 1000
        frmDisMsg.Timer1.Enabled = True
        frmDisMsg.Show vbModal
    End If
    rctmp.Close
    If ReturnValue = True And clsStation.RowMojodiControl = True Then
        DetailsString1 = ""
        Dim OldAmount As Long
        OldAmount = 0
        With FlxDetail
            For i = 1 To .Rows - 1
                If Val(.TextMatrix(i, IndexColGoodCode)) = mvarGoodCode Then
                    OldAmount = Val(.TextMatrix(i, IndexColAmount))
                End If
            Next
            DetailsString1 = GenerateDetailsString3(DetailsString1, IIf(Val(lblNum.Caption) = 0, 1 + OldAmount, Val(lblNum.Caption) + OldAmount), CStr(mvarGoodCode), CStr(mvarSellPrice), CStr(mvarDisCount), CStr(mvarRate), "", " ", CStr(mvarInventoryNo), "", CStr(mvarServePlace), "")
        End With
        If MojodiControlFlag = True And mvarStatus = Invoice Then   'And FWMojodiControl.Visible = True Then
            If MyFormAddEditMode = AddMode Then
                ReDim Parameter(3) As Parameter
                Parameter(0) = GenerateInputParameter("@DetailsString", adVarWChar, 4000, DetailsString1)
               Parameter(1) = GenerateInputParameter("@DetailsString2", adVarWChar, 4000, "")
               Parameter(2) = GenerateInputParameter("@DetailsString3", adVarWChar, 4000, "")
               Parameter(3) = GenerateInputParameter("@DetailsString4", adVarWChar, 4000, "")
                Set Rst = RunParametricStoredProcedure2Rec("CheckPreSave", Parameter)
                If Not (Rst.BOF Or Rst.EOF) Then
                    mvarAddeditMode = MyFormAddEditMode
                    frmMojodiReduce.Show vbModal
                    If frmMojodiReduce.Result = False Then
                       ReturnValue = False
                       lblNum.Caption = ""
                    End If
                End If
            Else
                ReDim Parameter(5) As Parameter
                mvarNo = Val(txtNo.Text)
                Parameter(0) = GenerateInputParameter("@DetailsString", adVarWChar, 4000, DetailsString1)
                Parameter(1) = GenerateInputParameter("@Status", adInteger, 4, mvarStatus)
                Parameter(2) = GenerateInputParameter("@No", adBigInt, 8, mvarNo)
               Parameter(3) = GenerateInputParameter("@DetailsString2", adVarWChar, 4000, "")
               Parameter(4) = GenerateInputParameter("@DetailsString3", adVarWChar, 4000, "")
               Parameter(5) = GenerateInputParameter("@DetailsString4", adVarWChar, 4000, "")
                Set Rst = RunParametricStoredProcedure2Rec("CheckPreSave_Edit", Parameter)
                Dim ss As String
                If Not (Rst.BOF Or Rst.EOF) Then
                    mvarAddeditMode = MyFormAddEditMode
                    frmMojodiReduce.Show vbModal
                    If frmMojodiReduce.Result = False Then
                       ReturnValue = False
                       lblNum.Caption = ""
                    End If
                End If

            End If
        End If

    End If
    GetGoodCode = ReturnValue
    If ReturnValue = True And clsInvoiceValue.ShowPictureGood = True Then
        GoodCode = Code
        Dim Result As Integer
        ReDim Parameters(1) As Parameter
        Parameters(0) = GenerateInputParameter("@GoodCode", adBigInt, 8, GoodCode)
        Parameters(1) = GenerateOutputParameter("@Result", adInteger, 1)
        Result = RunParametricStoredProcedure2String("Get_CounttblTotal_GoodPic_byCode", Parameters)
        If Result = 1 Then
            frmShowPictureGood.Show vbModal
        End If
   End If
End Function

Public Sub KeyPress(KeyAscii As Integer)

    Dim var1, var2 As Double
    Dim j As Double
    
    
    If MyFormAddEditMode = ViewMode Then
        Exit Sub
    End If
                
    If MvarUserDefine Then
    
        ReDim Parameter(3) As Parameter
        Parameter(0) = GenerateInputParameter("@Keycode", adInteger, 4, mvarKeyCode)
        Parameter(1) = GenerateInputParameter("@ShiftKey", adInteger, 4, MvarShiftKey)
        Parameter(2) = GenerateInputParameter("@StationId", adInteger, 4, clsArya.StationNo)
        Parameter(3) = GenerateInputParameter("@notSupportedType", adInteger, 4, EnumGoodType.forBuy)
        Set rctmp = RunParametricStoredProcedure2Rec("Get_Good_Kb_Count", Parameter)
    Else
    
        ReDim Parameter(1) As Parameter
        Parameter(0) = GenerateInputParameter("@BtnAscDefault", adInteger, 4, KeyAscii)
        Parameter(1) = GenerateInputParameter("@notSupportedType", adInteger, 4, EnumGoodType.forBuy)
        Set rctmp = RunParametricStoredProcedure2Rec("Get_Good_DefaultKb_Count", Parameter)
    End If
    
    i = rctmp.Fields("count")
    rctmp.Close
            
    If i > 1 Then
       
           Call frmFindGoods_Kb.SendVariables(MvarUserDefine, mvarKeyCode, MvarShiftKey, KeyAscii)
           frmFindGoods_Kb.Show vbModal

    ElseIf i = 1 Then
    
        If MvarUserDefine Then
             ReDim Parameter(4) As Parameter
             Parameter(0) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
             Parameter(1) = GenerateInputParameter("@StationId", adInteger, 4, clsArya.StationNo)
             Parameter(2) = GenerateInputParameter("@KeyCode", adInteger, 4, mvarKeyCode)
             Parameter(3) = GenerateInputParameter("@ShiftKey", adInteger, 4, MvarShiftKey)
             Parameter(4) = GenerateInputParameter("@NotSupportedGoodType", adInteger, 4, EnumGoodType.forBuy)
             Set rctmp = RunParametricStoredProcedure2Rec("Get_Good_KB", Parameter)
        Else
             ReDim Parameter(2) As Parameter
             Parameter(0) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
             Parameter(1) = GenerateInputParameter("@BtnAscDefault", adInteger, 4, KeyAscii)
             Parameter(2) = GenerateInputParameter("@NotSupportedGoodType", adInteger, 4, EnumGoodType.forBuy)
             Set rctmp = RunParametricStoredProcedure2Rec("Get_Good_DefaultKB", Parameter)
        End If
        If GetGoodCode(Val(rctmp.Fields("Code"))) = True Then
            ChangeGoodquantity
        End If
       
    End If
    MvarUserDefine = False
 

    
End Sub


Sub DoPrintLogo(PassedPrinterName As String, LogoFileName As String)

End Sub
Private Sub HideLstBoxes(KeyAscii As Integer)

If (KeyAscii = 27) Then
    Me.lstDifference.Visible = False
End If
    
End Sub

Private Function CalculateSumOfServeplace() As Integer
    
    Dim j As Integer
    Dim intServeplaces() As Integer
    
    ReDim Preserve intServeplaces(0)
    
    intServeplaces(0) = Val(FlxDetail.TextMatrix(1, 8))
    For i = 1 To MaxRowFlexGrid - 1
        ReDim Preserve intServeplaces(i)
'        If i <> MaxRowFlexGrid - 1 Then
            intServeplaces(i) = Val(FlxDetail.TextMatrix(i + 1, 8))
'        Else
'            intServeplaces(i) = mVarServePlace
'        End If
        For j = 0 To i - 1
            If Val(FlxDetail.TextMatrix(i, 8)) = intServeplaces(j) Then
                intServeplaces(i) = 0
                Exit For
            End If
        Next j
    
    Next i
    
    CalculateSumOfServeplace = 0
    For i = LBound(intServeplaces) To UBound(intServeplaces)
        CalculateSumOfServeplace = CalculateSumOfServeplace + intServeplaces(i)
    
    Next i

End Function


Public Function ChangeGoodquantity()
      On Error GoTo ErrHandler
 
    framelastFich.Visible = False
    If BlnPosApprovedWait = True Then
        ShowDisMessage "”Ì” „ „‰ Ÿ— œ—Ì«›  Å«”Œ «“ ŒÊœÅ—œ«“ „Ì »«‘œ", 1000
        Exit Function
    End If
'    Dim DatabaseBranch As Integer
'    ReDim Parameters(0) As Parameter
'    Parameters(0) = GenerateOutputParameter("@CurrentBranch", adInteger, 4)
'
'    DatabaseBranch = RunParametricStoredProcedure2String("Get_CurrentBranch", Parameters)
'
'    If CurrentBranch <> DatabaseBranch Then
'        frmDisMsg.lblMessage.Caption = "«„ò«‰ ’œÊ— ›Ì‘ »—«Ì ‘⁄»Â œÌê— ÊÃÊœ ‰œ«—œ "
'        frmDisMsg.Timer1.Interval = 2000
'        frmDisMsg.Timer1.Enabled = True
'        frmDisMsg.Show vbModal
'        Exit Function
'    End If
    If clsStation.CashClose = True And ClsFormAccess.EditInvoiceCashClose = False Then
        frmDisMsg.lblMessage.Caption = "’‰œÊﬁ »” Â «”  Ê «„ﬂ«‰ ’œÊ— ›Ì‘ ÊÃÊœ ‰œ«—œ"
        frmDisMsg.Timer1.Interval = 3000
        frmDisMsg.Timer1.Enabled = True
        frmDisMsg.Show vbModal
        Exit Function
    End If
    Dim Answer As Boolean
    Dim CanAdd As Boolean
    Dim AmountVar As Double

    If lblNum.Caption = "-" Then
        lblNum.Caption = "-1"
    End If
    
    If txtScale.Text = "-" Then
        txtScale.Text = "-1"
    End If
    
       
    
    If MaxRowFlexGrid > 1 And Val(lblNum.Caption) >= 0 Then
        For i = 1 To MaxRowFlexGrid
            If Val(FlxDetail.TextMatrix(i, 8)) = mvarServePlace Then
                CanAdd = True
                Exit For
            End If
        Next i
    
         If CanAdd = False Then
        
            intSumOfCurrentServePlaces = CalculateSumOfServeplace
            ReDim Parameter(2) As Parameter
            Parameter(0) = GenerateInputParameter("@SumOfCurrentServePlaces", adInteger, 4, intSumOfCurrentServePlaces)
            Parameter(1) = GenerateInputParameter("@intNewServePlace", adInteger, 4, mvarServePlace)
            Parameter(2) = GenerateOutputParameter("@Answer", adInteger, 1)
            
            Answer = RunParametricStoredProcedure("CheckInvoiceServePlace", Parameter)
        
            If Answer = False Then
            
                frmMsg.fwlblMsg.Caption = "¬Ì« „«Ì·Ìœ „Õ· ”—Ê  „«„ ò«·«Â« —« »Â " & cmbServePlace.Text & "  €ÌÌ— œÂÌœø "
                frmMsg.fwBtn(0).ButtonType = flwButtonOk
                frmMsg.fwBtn(0).Caption = "»·Ì"
                frmMsg.fwBtn(1).Visible = flwButtonCancel
                frmMsg.fwBtn(1).Caption = "ŒÌ—"
                frmMsg.fwBtn(1).Default = True
                frmMsg.Show vbModal
                
                If mvarMsgIdx = vbYes Then
                    Dim j As Integer
                    For i = 1 To MaxRowFlexGrid - 1
                        If FlxDetail.TextMatrix(i, 8) <> "" Then FlxDetail.TextMatrix(i, 8) = mvarServePlace
                        For j = MaxRowFlexGrid - 1 To i + 1 Step -1
                            If FlxDetail.TextMatrix(i, 5) = FlxDetail.TextMatrix(j, 5) And FlxDetail.TextMatrix(i, 3) = FlxDetail.TextMatrix(j, 3) Then
                                FlxDetail.TextMatrix(i, 1) = Val(FlxDetail.TextMatrix(i, 1)) + Val(FlxDetail.TextMatrix(j, 1))
                                FlxDetail.RemoveItem (j)
                                MaxRowFlexGrid = MaxRowFlexGrid - 1
                                RefreshFlxDetailRowNumber
                                If FlxDetail.Rows < MaxInvoiceRows Then
                                    AddEmptyRow     'add row Instead of Remove
                                End If
                                RefreshLables
                            End If
                        Next j
                    Next i
                   intSumOfCurrentServePlaces = CalculateSumOfServeplace
                    
                Else
                    frmDisMsg.lblMessage = "«Ì‰  —òÌ» «“ „ò«‰Â«Ì ”—Ê œ— ”Ì” „  ⁄—Ì› ‰‘œÂ «”  " & vbCrLf & "·ÿ›« ›«ò Ê— —« «’·«Õ ‰„ÊœÂ Ê ”Å” À»  ‰„«ÌÌœ"
    
                    frmDisMsg.Timer1.Enabled = True
                    frmDisMsg.Show vbModal
    
                    Exit Function
                End If
                
            Else
                intSumOfCurrentServePlaces = intSumOfCurrentServePlaces + mvarServePlace
                FlxDetail.ColHidden(8) = False
                FlxDetail.ColWidth(10) = FlxDetail.Width / 8      'Diffrence
            End If
        End If
    Else
        intSumOfCurrentServePlaces = mvarServePlace
    End If
    

   intCountGood = 0
   If MaxRowFlexGrid > 1 And Val(lblNum.Caption) >= 0 Then
        Dim k As Integer
        If lblNum.Caption = "" Then
            intCountGood = intCountGood + 1
        Else
            intCountGood = intCountGood + Val(lblNum.Caption)
        End If
        For k = 1 To MaxRowFlexGrid - 1
            If FlxDetail.TextMatrix(k, 15) = True Then
                 intCountGood = intCountGood + Val(FlxDetail.TextMatrix(k, 1))
            End If
         Next k
    Else
        If lblNum.Caption = "" Then
             intCountGood = intCountGood + 1
        Else
             intCountGood = intCountGood + Val(lblNum.Caption)
        End If
    End If
    
 If (clsStation.CountCustomerGood > 0 And intCountGood > clsStation.CountCustomerGood And Val(lblCustomer.Tag) <> -1) Then
         frmDisMsg.lblMessage = " ⁄œ«œ «ﬁ·«„ »Ì‘ «“  ⁄œ«œ „Ã«“ «” "

                    frmDisMsg.Timer1.Enabled = True
                    frmDisMsg.Show vbModal

                    Exit Function
    
End If
    Select Case mvarUnitGood
        'Weight Good
        Case 1
        
            If clsStation.DirectBascule Then   'And clsStation.BasculeOn
                If lblNum.Caption <> "" Then
                    AmountVar = Val(lblNum.Caption)
                    
                ElseIf txtScale.Text <> "" Then
                    AmountVar = Val(txtScale.Text)
                    txtScale.Text = ""
                Else
                
                    frmMsg.fwlblMsg.Caption = " .  —«“ÊÌ œÌÃÌ «· ¬„«œÂ ‰Ì”  "
                    frmMsg.fwBtn(0).Visible = False
                    frmMsg.fwBtn(1).ButtonType = flwButtonOk
                    frmMsg.fwBtn(1).Caption = "ﬁ»Ê·"
                    frmMsg.Show vbModal
                    Exit Function
                    
                End If
                
            Else
            
                If lblNum.Caption <> "" Then
                    AmountVar = Val(lblNum.Caption)
                Else
                    AmountVar = 1
                End If
                
            End If
            
        Case Else     'Numeric Good
        
            If lblNum.Caption <> "" Then
                AmountVar = Round(Val(lblNum.Caption), 0)
            Else
                AmountVar = 1
            End If
    End Select
    
    Dim Row_Find As Integer
    
    If FindRecord_FlexGrid(mvarGoodCode) = True Then 'Exist Good In Fich
    
        If left(lblNum.Caption, 1) = "-" And mvarUnitGood = 1 And clsStation.DeletedGood = True Then   'Weight Good & Delete
            FlxDetail.TextMatrix(FlxDetail.Row, 1) = 0
        Else
            FlxDetail.TextMatrix(FlxDetail.Row, 1) = AmountVar + Val(FlxDetail.TextMatrix(FlxDetail.Row, 1))
        End If
        
        If FlxDetail.TextMatrix(FlxDetail.Row, 1) < 0 Then
             frmMsg.fwlblMsg.Caption = " . „ﬁœ«— ﬂ«·« ‰„Ì  Ê«‰œ „‰›Ì »«‘œ"
             frmMsg.fwBtn(0).Visible = False
             frmMsg.fwBtn(1).ButtonType = flwButtonOk
             frmMsg.fwBtn(1).Caption = "ﬁ»Ê·"
             frmMsg.Show vbModal
             FlxDetail.TextMatrix(FlxDetail.Row, 1) = -AmountVar + Val(FlxDetail.TextMatrix(FlxDetail.Row, 1))
             AmountVar = 0
             lblNum.Caption = ""
             txtScale.Text = ""
             Exit Function
        End If
        
        If FlxDetail.TextMatrix(FlxDetail.Row, 1) = 0 Then       '
            FlxDetail.RemoveItem (FlxDetail.Row)
            If FlxDetail.Rows < MaxInvoiceRows Then
                AddEmptyRow     'add row Instead of Remove
            End If
            MaxRowFlexGrid = MaxRowFlexGrid - 1
            RefreshFlxDetailRowNumber
            
''''            frmMsg.fwlblMsg.Caption = " .ò«·«Ì „Ê—œ ‰Ÿ— «“ ·Ì”  Õ–› ‘œ "
''''            frmMsg.Fwbtn(0).Visible = False
''''            frmMsg.Fwbtn(1).ButtonType = flwButtonOk
''''            frmMsg.Fwbtn(1).Caption = "ﬁ»Ê·"
''''            frmMsg.Show vbModal
            frmDisMsg.lblMessage.Caption = " .ò«·«Ì „Ê—œ ‰Ÿ— «“ ·Ì”  Õ–› ‘œ "
            frmDisMsg.Timer1.Interval = 1000
            frmDisMsg.Timer1.Enabled = True
            frmDisMsg.Show vbModal
            
        Else
            FlxDetail.TextMatrix(FlxDetail.Row, 4) = Val(FlxDetail.TextMatrix(FlxDetail.Row, 1) * CCur(FlxDetail.TextMatrix(FlxDetail.Row, 3)))
        End If
        
           
    Else                         'Not Exist in Fich
                      
       If AmountVar <= 0 Then
                frmMsg.fwlblMsg.Caption = " «Ì‰ ﬂ«·« œ— ·Ì”  ‰Ì” .‰„Ì  Ê«‰Ìœ «Ì‰ ﬂ«·« Õ–› ﬂ‰Ìœ"
                frmMsg.fwBtn(0).Visible = False
                frmMsg.fwBtn(1).ButtonType = flwButtonOk
                frmMsg.fwBtn(1).Caption = "ﬁ»Ê·"
                frmMsg.Show vbModal
                AmountVar = 0
                lblNum.Caption = ""
                txtScale.Text = ""
                Exit Function
        End If
            
 
'
'        If clsArya.LimitedVersion = True Then
'            If MaxRowFlexGrid > 5 Then
'                ShowMessage " ‘„« œ— Õ«· «” ›«œÂ «“ ‰”ŒÂ ¬“„«Ì‘Ì ”Ì” „ Â«Ì ›—Ê‘ê«ÂÌ «› ÃÌ ¬—Ì« „Ì »«‘Ìœ" & " »—«Ì  ÂÌÂ ‰”ŒÂ «’·Ì ‰—„ «›“«—»« ‘—ﬂ  «› ÃÌ ¬—Ì«  „«” »êÌ—Ìœ " & " ·›‰  „«”  88554488-88554477-88554466- 88554455", True, False, " «∆Ìœ", ""
'                Exit Function
'            End If
'        End If
        
        FlxDetail.Row = MaxRowFlexGrid
        
        FlxDetail.TextMatrix(FlxDetail.Row, 0) = FlxDetail.Row
        FlxDetail.TextMatrix(FlxDetail.Row, 1) = AmountVar
        
        FlxDetail.TextMatrix(FlxDetail.Row, 5) = mvarGoodCode
        FlxDetail.TextMatrix(FlxDetail.Row, 8) = mvarServePlace
        
        FlxDetail.ShowCell FlxDetail.Row, 0
        
        
        FlxDetail.TextMatrix(FlxDetail.Row, 2) = mvarGoodName
        FlxDetail.TextMatrix(FlxDetail.Row, 6) = mvarGoodWeight
        FlxDetail.TextMatrix(FlxDetail.Row, 3) = mvarSellPrice
        FlxDetail.TextMatrix(FlxDetail.Row, 7) = mvarUnitGood
        FlxDetail.TextMatrix(FlxDetail.Row, 12) = mvarRate
        
        If FlxDetail.TextMatrix(FlxDetail.Row, 3) = "" Then
           FlxDetail.TextMatrix(FlxDetail.Row, 1) = ""
        End If
        On Error GoTo ErrHandler
        
        FlxDetail.TextMatrix(FlxDetail.Row, 4) = CLng(Val(FlxDetail.TextMatrix(FlxDetail.Row, 1)) * Val(FlxDetail.TextMatrix(FlxDetail.Row, 3)))
        FlxDetail.TextMatrix(FlxDetail.Row, 11) = mvarDisCount
        FlxDetail.TextMatrix(FlxDetail.Row, 14) = mvarInventoryNo
        FlxDetail.TextMatrix(FlxDetail.Row, 15) = bolMainGroup
        FlxDetail.TextMatrix(FlxDetail.Row, 16) = mvarMojodi
        If mvarMojodi >= 0 Then
            If mvarMojodi <> Int(mvarMojodi) Then
                FlxDetail.TextMatrix(FlxDetail.Row, 16) = Format(mvarMojodi, "##.000")
                FlxDetail.TextMatrix(FlxDetail.Row, 16) = Val(FlxDetail.TextMatrix(FlxDetail.Row, 16)) ' Delete Last Zeros
            Else
                 FlxDetail.TextMatrix(FlxDetail.Row, 16) = mvarMojodi
            End If
        Else
            If mvarMojodi <> Int(mvarMojodi) Then
                FlxDetail.TextMatrix(FlxDetail.Row, 16) = -Format(mvarMojodi, "##.000")
                FlxDetail.TextMatrix(FlxDetail.Row, 16) = Val(FlxDetail.TextMatrix(FlxDetail.Row, 16)) & "-" ' Delete Last Zeros
            Else
                 FlxDetail.TextMatrix(FlxDetail.Row, 16) = -mvarMojodi & "-"
            End If
        End If
        
        FlxDetail.TextMatrix(FlxDetail.Row, 17) = mvarDuty
        FlxDetail.TextMatrix(FlxDetail.Row, 18) = mvarTax
        On Error GoTo 0
        
        If FlxDetail.Row = (FlxDetail.Rows - 1) Then
           AddEmptyRow
           'FlxDetail.Row = FlxDetail.Row - 1
        End If
        
        FlxDetail.Row = FlxDetail.Row + 1       'Next Row
        MaxRowFlexGrid = FlxDetail.Row            'Last Row

    End If
    
    FlxDetail.Row = MaxRowFlexGrid     'Last Row
    
    BtnKeypad(10).Enabled = True
    BtnKalaDelete.Enabled = True
    BtnKeypad(11).Enabled = True
    lblNum.Caption = ""
    RefreshLables  'Set Lables
    
    FlxDetail.TopRow = FlxDetail.Rows - (MaxInvoiceRows - 1)
    
    txtScale.Text = ""
    
    FlxDetail.Select MaxRowFlexGrid, 1
    FlxDetail.ShowCell MaxRowFlexGrid, 1
    
    If clsStation.CustomerOnlinePrice = True Then
         CustomerDisplay Val(lblSumPrice.Tag), Val(AmountVar * mvarSellPrice), AmountVar
    Else
         CustomerDisplay Val(AmountVar * mvarSellPrice), mvarGoodName, AmountVar
    End If
    If clsInvoiceValue.ShowInvoiceMenu = True Then
        frmShowInvoiceMenu.UpdateGridValue
    End If
'    If clsArya.LimitedVersion = True Then
'        TrialCountFlag = TrialCountFlag + 1
'        If TrialCountFlag = 10 Then
'            ShowMessage " ‘„« œ— Õ«· «” ›«œÂ «“ ‰”ŒÂ ¬“„«Ì‘Ì ”Ì” „ Â«Ì ›—Ê‘ê«ÂÌ «› ÃÌ ¬—Ì« „Ì »«‘Ìœ" & " »—«Ì  ÂÌÂ ‰”ŒÂ «’·Ì ‰—„ «›“«—»« ‘—ﬂ  «› ÃÌ ¬—Ì«  „«” »êÌ—Ìœ " & " ·›‰  „«”  88554488-88554477-88554466- 88554455", True, False, " «∆Ìœ", ""
'            TrialCountFlag = 0
'        End If
'    End If
    
    Exit Function
    
ErrHandler:
    Select Case err.Number
        Case 6
            MsgBox "„ﬁœ«— ò«·«Ì Ê«—œ ‘œÂ »Ì‘ — «“ „ﬁœ«—Ì”  òÂ »—‰«„Â „Ì  Ê«‰œ ﬁ»Ê· ò‰œ " & vbCrLf & "·ÿ›« Ìò ⁄œœ òÊçò — Ê«—œ ‰„«ÌÌœ"
            Add
    End Select
End Function
Private Sub ShowMonitor(RecordType As Long)
'    If mdifrm.Winsock_Farabin.State <> sckConnected Then mdifrm.Winsock_Farabin.Close: mdifrm.Winsock_Farabin.Connect
    strFarabin = ""
    Dim i As Long
    strFarabin = clsArya.StationNo & " ; " & FWLed1 & " ; " & lblDiscountTotal.Caption & " ; " & LblDutyTotal.Caption & " ; " & lblTaxTotal.Caption & " ; " & lblSumPrice.Tag & "$"
    If RecordType = 1 Then      ' Fill Factor Not Init
        With FlxDetail
            For i = 1 To MaxRowFlexGrid - 1
                strFarabin = strFarabin & .TextMatrix(i, 0) & " ; " & .TextMatrix(i, 1) & " ; " & UTF8_Encode_System(.TextMatrix(i, 2)) & " ; " & .TextMatrix(i, 3) & " ; " & .TextMatrix(i, 4) & " ; " & UTF8_Encode_System(.TextMatrix(i, 10)) & "||"
            Next
        End With
    End If
   ' If mdifrm.Winsock_Farabin.State = sckConnected Then mdifrm.Winsock_Farabin.SendData strFarabin
    mdifrm.Winsock_Farabin.Connect
End Sub

Public Function UTF8_Encode_System(ByVal sStr As String) As String
    Dim buffer As String
    Dim Length As Long
     
    'Get the length of the converted data.
    Length = WideCharToMultiByte(65001, 0, StrPtr(sStr), Len(sStr), 0, 0, 0, 0)
     
    'Ensure the buffer is the correct size.
    buffer = String$(Length, 0)
     
    'Convert the string into the buffer.
    Length = WideCharToMultiByte(65001, 0, StrPtr(sStr), Len(sStr), StrPtr(buffer), Len(buffer), 0, 0)
     
    'Access needs it in unicode?
    buffer = StrConv(buffer, vbUnicode)
     
    'Chop of any crap.
    buffer = left$(buffer, Length)
     
    'Return baby.
    UTF8_Encode_System = buffer
End Function
Private Sub RefreshFlxDetailRowNumber()
    Dim i As Integer
    With FlxDetail
        For i = 1 To MaxRowFlexGrid - 1
            .TextMatrix(i, IndexColRow) = i
        Next i
    End With
End Sub

Public Sub ChangeLanguage()
    
    Select Case clsStation.Language
        Case EnumLanguage.Farsi
       
            mdifrm.Caption = clsArya.Company
            cmdPay.Caption = "œ—Ì«› "
            CmdStationSaleSummery.Caption = "ê“«—‘ ’‰œÊﬁ"
            cmdTempFich.Caption = "›Ì‘ „Êﬁ "
'            fwBtnCtrl2.Caption = "‰„«Ì‘"
            FWBtnPayk.Caption = "ÅÌﬂ"
            FWBtnSplit.Caption = "„⁄„Ê·Ì"
            FWlblAcc.Caption = "Õ”«» Â« »” Â"
            FWlblCash.Caption = "’‰œÊﬁ »” Â"
            FWLblEdit.Caption = "«’·«ÕÌ"
            fwlblMode.Caption = "„—Ê—"
            fwlblRecursive.Caption = "„—ÃÊ⁄Ì"
            FWMojodiControl.Caption = "»«ﬁÌ„«‰œÂ ﬂ«·«"
'            FWRealButton1(0).Caption = "Ã” ÃÊÌ ﬂ«·«"
            Label1.Caption = "Ã„⁄"
            Label11.Caption = "œ—Ì«› "
'            Label3.Caption = " : „Ì“ "
            Label4.Caption = "Ã„⁄ ﬂ·"
            Label5.Caption = " ⁄œ«œ"
            Label6.Caption = "Œ—ÌœÂ«Ì «„—Ê“"
            Label7.Caption = "«⁄ »«—Ì"
            Label8.Caption = "Œ—ÌœÂ«Ì „«Â"
            LastCredit.Caption = "„«‰œÂ"
            LastDate.Caption = "¬Œ—Ì‰ Œ—Ìœ"
            LastNo.Caption = "¬Œ—Ì‰ ›Ì‘"
            LastPrice.Caption = "„»·€ ¬Œ—Ì‰ Œ—Ìœ"
            LblInvoice.Caption = "›«ﬂ Ê— ›—Ê‘"
            LblOrder.Caption = "Õ÷Ê—Ì"
            LblPacking.Caption = "»” Â »‰œÌ"
            lblServePlace.Caption = "”«·‰"
            LblService.Caption = "”—ÊÌ”"
            MaxPrice.Caption = "»Ì‘ —Ì‰ Œ—Ìœ"
            MinPrice.Caption = "ﬂ„ —Ì‰ Œ—Ìœ"
            Frame10.Caption = " Ê÷ÌÕ« "
            Frame8.Caption = "«⁄ »«—"
'            frameMenu.Caption = "·Ì”  „‰ÊÂ«Ì ﬂ«·«Â«"
            FrameCustInfo.Caption = "                             «ÿ·«⁄«  „‘ —ﬂ                                         "
            Me.Caption = "›«ﬂ Ê— ›—Ê‘"
'           fwBtnCustFind.Caption = "„‘ —ﬂ"
'           FWBtnGarsoon.Caption = "ê«—”‰"
'           FWBtnTable.Caption = "„Ì“"
           BtnFindGood.Caption = "Ã” ÃÊÌ ﬂ«·«"
           BtnKalaDelete.Caption = "Õ–› ﬂ«·«"
           CmdColor.Caption = " €ÌÌ— —‰ê "
           LblDiscount.Caption = " Œ›Ì›"
           LblCarryFee.Caption = "ﬂ—«ÌÂ Õ„·"
           LblTax.Caption = "„«·Ì« "
           lblDuty.Caption = "⁄Ê«—÷"
            For i = 0 To 9
                BtnKeypad(i).RightToLeft = True
                BtnKeypad(i).Font.Name = "B Nazanin"
              '  BtnKeypad(i).Style = 1
            Next i
           txtDate.RightToLeft = True
            LblSubTotal.RightToLeft = True
            lblDiscountTotal.RightToLeft = True
            lblCarryFeeTotal.RightToLeft = True
            lblPackingTotal.RightToLeft = True
            lblTaxTotal.RightToLeft = True
            LblDutyTotal.RightToLeft = True
            lblServiceTotal.RightToLeft = True
            lblSumPrice.RightToLeft = True
            txtSumCountNo.RightToLeft = True
            
           
            For i = BascoleLabel.LBound To BascoleLabel.UBound
                BascoleLabel(i).Caption = " —«“ÊÌ : " & i
            Next i
            
            ReDim Parameter(1) As Parameter
            Parameter(0) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
            Parameter(1) = GenerateInputParameter("@PartitionID", adInteger, 4, clsStation.PartitionID)
            
            Set Rst = RunParametricStoredProcedure2Rec("RetrivePartitionDescription", Parameter)
            If Not (Rst.BOF Or Rst.EOF) Then

                FwPartition.Caption = "Ê«Õœ " & Rst.Fields("PartitionName")
                DefaultServicePercent = Rst.Fields("DefaultServicePercent").Value
            Else
                FwPartition.Caption = "Ê«Õœ  ----- "
                DefaultServicePercent = 0
            End If
            'FwPartition.Caption = "Ê«Õœ ‘„«—Â¡ " & ClsStation.PartitionID
            If Invoice_FontFlexGridName = "" Then
                 Invoice_FontFlexGridName = "Arial"
                 Invoice_FontFlexGridSize = 12
                 Invoice_FontFlexGridBold = True
            End If
            With FlxDetail
                 
''''                .Font.Name = "Arial"
''''                .Font.Size = 11
''''                .Font.Bold = True
                .Font.Name = Invoice_FontFlexGridName
                .Font.size = Invoice_FontFlexGridSize
                .Font.Bold = Invoice_FontFlexGridBold
                .RightToLeft = True
                .TextMatrix(0, 0) = "—œÌ›(-)"
                .TextMatrix(0, 1) = "„ﬁœ«—"
                .TextMatrix(0, 2) = "‰«„ ò«·« (+)"
                .TextMatrix(0, 3) = "›Ì"
                .TextMatrix(0, 4) = "Ã„⁄"
                .TextMatrix(0, 5) = "ﬂœ ﬂ«·«"
                .TextMatrix(0, 8) = "”—Ê"
                .TextMatrix(0, 10) = " €ÌÌ—« "
                .TextMatrix(0, 11) = "œ—’œ  Œ›Ì›"
                .TextMatrix(0, 12) = "‰—Œ"
                .TextMatrix(0, 13) = "’‰œ·Ì"
                .TextMatrix(0, 14) = "«‰»«—"
                .TextMatrix(0, 15) = "ê—ÊÂ «’·Ì"
                .TextMatrix(0, 16) = "„ÊÃÊœÌ"
                .TextMatrix(0, 17) = "⁄Ê«—÷"
                .TextMatrix(0, 18) = "„«·Ì« "
                .Cell(flexcpFontName, 0, 0, 0, .Cols - 1) = "Tahoma"
                .Cell(flexcpFontBold, 0, 0, 0, .Cols - 1) = False
                .Cell(flexcpFontSize, 0, 0, 0, .Cols - 1) = 10
            
                .ColFormat(3) = "###,###"
                .ColFormat(4) = "###,###"
                .ColFormat(16) = "###,###"
            
            End With
            
            fwCash.Caption = "«Ì” ê«Â ‘„«—Â " & clsArya.StationNo
            lstDifference.RightToLeft = True
        
        Case EnumLanguage.English
        
            mdifrm.Caption = Space(100) & Trim(clsArya.LatinCompany)
            
            
            cmdPay.Caption = "Cash"
            CmdStationSaleSummery.Caption = "Cash report"
            cmdTempFich.Caption = "Temporary Invoice"
'            fwBtnCtrl2.Caption = "Display"
            FWBtnPayk.Caption = "Delivery"
            FWBtnSplit.Caption = "Regular"
            FWlblAcc.Caption = "Closing accounts"
            FWlblCash.Caption = "Closing cash"
            FWLblEdit.Caption = "Edited invoice"
            fwlblMode.Caption = "View"
            fwlblRecursive.Caption = "Refund"
            FWMojodiControl.Caption = "Remaining goods"
'            FWRealButton1(0).Caption = "Searching goods"
            Label1.Caption = "Sum"
            Label11.Caption = "Receive"
'            Label3.Caption = "Table"
            Label4.Caption = "Total price"
            Label5.Caption = "Quantity"
            Label6.Caption = "Today buy"
            Label7.Caption = "Credit buy"
            Label8.Caption = "Month buy"
            LastCredit.Caption = "Remainder"
            LastDate.Caption = "Last buy"
            LastNo.Caption = "Last invoice"
            LastPrice.Caption = "Last buy cost"
            LblOrder.Caption = "Inside"
            lblServePlace.Caption = "Saloon"
            lblDuty.Caption = "Duty"
            LblTax.Caption = "Tax"
            MaxPrice.Caption = "Maximum cost"
            MinPrice.Caption = "Minimum cost"
            Frame10.Caption = "Description"
            Frame8.Caption = "Credit"
'            frameMenu.Caption = "Goods menu"
            FrameCustInfo.Caption = "                            Customer Info                                         "
            LblInvoice.Caption = "Invoice"
            Me.Caption = "Invoice"
            CmdColor.Caption = "Color Change"
            
            txtDate.RightToLeft = False
            LblSubTotal.RightToLeft = False
            lblDiscountTotal.RightToLeft = False
            lblCarryFeeTotal.RightToLeft = False
            lblPackingTotal.RightToLeft = False
            lblTaxTotal.RightToLeft = False
            LblDutyTotal.RightToLeft = False
            lblServiceTotal.RightToLeft = False
            lblSumPrice.RightToLeft = False
            txtSumCountNo.RightToLeft = False
            
            For i = 0 To 9
                BtnKeypad(i).RightToLeft = False
                BtnKeypad(i).Caption = i
                BtnKeypad(i).Font.Name = "Times New Roman"
             '   BtnKeypad(i).Style = 1
            Next i
            
            BtnFindGood.Caption = "GoodSearch"
'            FWBtnGarsoon.Caption = "Garson"
'            FWBtnTable.Caption = "Table"
'            fwBtnCustFind.Caption = "Customer"
            BtnKalaDelete.Caption = "Delete Item"
            BtnKalaDelete.Font.Name = "TimesNewRoman"
            LblDiscount.Caption = "Discount"
            LblCarryFee.Caption = "Shipping"
            LblService.Caption = "Service"
            LblPacking.Caption = "Packing"
            LblDiscount.Font.Name = "TimesNewRoman"
            LblCarryFee.Font.Name = "TimesNewRoman"
            LblPacking.Font.Name = "TimesNewRoman"
            LblService.Font.Name = "TimesNewRoman"
            LblTax.Font.Name = "TimesNewRoman"
            
            For i = BascoleLabel.LBound To BascoleLabel.UBound
                BascoleLabel(i).Caption = "Bascule #" & i
            Next i
            
            ReDim Parameter(1) As Parameter
            Parameter(0) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
            Parameter(1) = GenerateInputParameter("@PartitionID", adInteger, 4, clsStation.PartitionID)
            
            Set Rst = RunParametricStoredProcedure2Rec("RetrivePartitionDescription", Parameter)
            If Not (Rst.BOF Or Rst.EOF) Then

                FwPartition.Caption = "partition:  " & Rst.Fields("PartitionName")
                DefaultServicePercent = Rst.Fields("DefaultServicePercent").Value
            Else
                FwPartition.Caption = "partition  ----- "
                DefaultServicePercent = 0
            End If
            
            With FlxDetail
            
                .Font.Name = "TimesNewRoman"
                .Font.size = 11
                .Font.Bold = True
                
                .RightToLeft = False
                .TextMatrix(0, 0) = "Row"
                .TextMatrix(0, 1) = "Amount"
                .TextMatrix(0, 2) = "Good Name"
                .TextMatrix(0, 3) = "Fee"
                .TextMatrix(0, 4) = "Total"
                .TextMatrix(0, 5) = "GoodCode"
                .TextMatrix(0, 8) = "Serve"
                .TextMatrix(0, 10) = "Changes"
                .TextMatrix(0, 11) = "Discount"
                .TextMatrix(0, 12) = "Rate"
                .TextMatrix(0, 13) = "Chair"
                .TextMatrix(0, 14) = "Store"
                .TextMatrix(0, 15) = "Main Group"
                .TextMatrix(0, 16) = "Remain"
                .TextMatrix(0, 17) = "Duty"
                .TextMatrix(0, 18) = "Tax"
                .ColFormat(3) = "###,###"
                .ColFormat(4) = "###,###"
                .ColFormat(16) = "###,###"
            End With
            fwCash.Caption = "Station #" & clsArya.StationNo
            lstDifference.RightToLeft = False

    
    End Select
    
    ReDim Parameter(1) As Parameter
    Parameter(0) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
    Parameter(1) = GenerateInputParameter("@Branch", adInteger, 4, CurrentBranch)
    Set rctmp = RunParametricStoredProcedure2Rec("GetInventory_Branch", Parameter)
    FlxDetail.ColComboList(14) = FlxDetail.BuildComboList(rctmp, "Description", "InventoryNo")
    rctmp.Close
    
    Dim strTemp As String
    
    If Rst.State = 1 Then Rst.Close
    
    ReDim Parameter(0) As Parameter
    Parameter(0) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
    Set Rst = RunParametricStoredProcedure2Rec("Get_Serveplace", Parameter)
    
    strTemp = FlxDetail.BuildComboList(Rst, "Description", "intServePlace")
    FlxDetail.ColComboList(8) = strTemp
    If Rst.State <> 0 Then Rst.Close
 
    cmbServePlace.Clear
    
    If Rst.State = 1 Then Rst.Close
    
    ReDim Parameter(0) As Parameter
    Parameter(0) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
    Set Rst = RunParametricStoredProcedure2Rec("Get_Serveplace", Parameter)
    If Not (Rst.EOF = True And Rst.BOF = True) Then
        Do While Rst.EOF <> True
            cmbServePlace.AddItem CStr(Rst.Fields("Description"))
            cmbServePlace.ItemData(cmbServePlace.ListCount - 1) = Val(Rst.Fields("intServePlace"))
            Rst.MoveNext
        Loop
    End If
    If Rst.State <> 0 Then Rst.Close
    
    Set Rst = RunStoredProcedure2RecordSet("Get_All_tblPub_SellPrice")
    strTemp = FlxDetail.BuildComboList(Rst, "Description", "Code")
    FlxDetail.ColComboList(12) = strTemp
    
    Rst.Close
    
    'ValueBtnMenu
    
    Set Rst = Nothing
End Sub
Private Sub GetProperController()
    
    On Error GoTo Error_Handler
    
    ReDim Parameter(0) As Parameter
    Parameter(0) = GenerateInputParameter("@StationId", adInteger, 4, clsArya.StationNo)
    
    Set rctmp = RunParametricStoredProcedure2Rec("Get_DeviceSetting", Parameter)
    
    i = 1
    
    While (rctmp.EOF <> True) And (i <= mscSerial.Count)
        If rctmp.Fields("DeviceCode").Value = EnumDevice.MagnetCardReader And clsStation.LoyaltyCustomers = True Then     ' Not Lpt Port
        
        ElseIf rctmp.Fields("PortNo").Value <> 0 And rctmp.Fields("PortNo").Value <> 20 And rctmp.Fields("DeviceCode").Value <> EnumDevice.CallerIdInterface2_AlmP6 And rctmp!DeviceCode <> EnumDevice.SmsCenter And rctmp.Fields("DeviceCode").Value <> EnumDevice.RFT230 Then     ' Not Lpt Port
            mscSerial(i).CommPort = rctmp.Fields("PortNo").Value
            mscSerial(i).Settings = rctmp.Fields("BaudRate").Value & ",N,8,1"
            DeviceCode(i) = rctmp.Fields("DeviceCode").Value
            DeviceType(i) = rctmp.Fields("DeviceTypeCode").Value
            RThreshold(i) = rctmp.Fields("RThreshold").Value
            mscSerial(i).InBufferSize = rctmp.Fields("BufferSize").Value
            mscSerial(i).RThreshold = rctmp.Fields("RThreshold").Value
            
            If Not (mscSerial(i).PortOpen) And (rctmp.Fields("DeviceCode").Value <> EnumDevice.MDS14000 And rctmp.Fields("DeviceCode").Value <> EnumDevice.MDS11000 And rctmp.Fields("DeviceCode").Value <> EnumDevice.Mahak_Serial) Then
                mscSerial(i).PortOpen = True
            End If
          
            If rctmp.Fields("DeviceTypeCode").Value = EnumDeviceType.CustomerDisplay Then
               clsStation.CustDisplayModel = rctmp.Fields("DeviceCode").Value
               clsStation.CustDisplayPort = i

            ElseIf rctmp.Fields("DeviceTypeCode").Value = EnumDeviceType.CashDrawer Then
               clsStation.DrawerModel = rctmp.Fields("DeviceCode").Value
               clsStation.DrawerPort = i
            
            ElseIf rctmp.Fields("DeviceTypeCode").Value = EnumDeviceType.Pos Then
               clsStation.PosModel = rctmp.Fields("DeviceCode").Value
               clsStation.PosPort = i
            
            ElseIf rctmp.Fields("DeviceTypeCode").Value = EnumDeviceType.CardReader Then
                Dim j As Integer
                If rctmp!DeviceCode = EnumDevice.BarcodeTimeReader Then
                    mscSerial(i).Output = Chr$(238) + Chr$(238) + Chr$(238) + Chr$(238) + Chr$(238) + Chr$(238) + Chr$(238) + Chr$(238) + Chr$(238) + Chr$(238) + Chr$(238) + Chr$(238) + Chr$(238) + Chr$(238) + Chr$(238) _
                    + Chr$(238) + Chr$(238) + Chr$(238) + Chr$(238) + Chr$(238) + Chr$(238) + Chr$(238) + Chr$(238) + Chr$(238) + Chr$(238) _
                    + Chr$(221) + Chr$(1) + Chr$(10) + Chr$(0) + Chr$(33) + Chr$(0) + Chr$(1) + Chr$(2) + Chr$(255) _
                    + Chr$(255) + Chr$(255) + Chr$(255) + Chr$(255) + Chr$(255) + Chr$(255) + Chr$(255) + Chr$(255) + Chr$(0) + Chr$(0) _
                    + Chr$(0) + Chr$(96) + Chr$(1) + Chr$(10) + Chr$(0) + Chr$(0) + Chr$(0) + Chr$(69) + Chr$(2) + Chr$(0) _
                    + Chr$(10) + Chr$(9) + Chr$(226) + Chr$(255) + Chr$(255) + Chr$(0) + Chr$(24) + Chr$(0)
                   
                    Sleep 2000
                    For j = 1 To 30
                        mscSerial(i).Output = Chr$(128) + Chr$(0) + Chr$(0)
                    Next j
                    mscSerial(i).Output = Chr$(128) + Chr$(128) + Chr$(128) + Chr$(128) + Chr$(0) + Chr$(0) + Chr$(0) + Chr$(0) + Chr$(128) + Chr$(0)
                    For j = 1 To 50
                          mscSerial(i).Output = Chr$(68)
                    Next j
                    mscSerial(i).Output = Chr$(221) + Chr$(0)
                    TimerReader.Enabled = True
                    TimeReaderPort = i
                 
                End If
                
            ElseIf rctmp.Fields("DeviceTypeCode").Value = EnumDeviceType.Bascule Then
               
'                If rctmp!DeviceCode = EnumDevice.MDS14000 Then
'                    MDS14000CTL1.CommPort = rctmp.Fields("PortNo").Value
'                    MDS14000CTL1.RS485Connection = False
'                    MDS14000CTL1.Id = 1
'                ElseIf rctmp!DeviceCode = EnumDevice.MDS11000 Then
'                    Mahak11000.Model = MDS15000
'                    Mahak11000.CommPort = Val(rctmp.Fields("PortNo").Value)
'                    Mahak11000.Network = False
                If rctmp!DeviceCode = EnumDevice.Mahak_Serial Then
                    'Set MahakScaleOCX3 = CreateObject("SmdOcx2.MahakScaleOCX3")
                    Set MahakScaleOCX3 = CreateObject("SMDOCX3.MahakScaleOCX31")
                   ' MahakScaleOCX3.Model = MDS14000
                    MahakScaleOCX3.ConnectionType = 1 ' RS232
                    MahakScaleOCX3.CommPort = Val(rctmp.Fields("PortNo").Value)
'                    MahakScaleOCX3.ActivateScale (1 , True)
                    MahakScaleOCX3.Network = False
                    MahakScaleOCX3.ActiveTimer
                ElseIf rctmp!DeviceCode = EnumDevice.Mahak Then
                ElseIf rctmp!DeviceCode = EnumDevice.Sairan Then
                End If
                TimerScale.Interval = 400
               
'                ClsBascole(i).BascoleType = rctmp.Fields("DeviceTypeCode").Value
'                ClsBascole(i).BascoleCode = rctmp.Fields("DeviceCode").Value
'                ClsBascole(i).PortNo = rctmp.Fields("PortNo").Value
'                ClsBascole(i).RThreshold = rctmp.Fields("RThreshold").Value
                clsStation.BasculeModel = rctmp.Fields("DeviceCode").Value
                clsStation.BasculePort = i
                
'                If rctmp!DeviceCode = EnumDevice.Pand Then mscSerial(i).RThreshold = 0   ''Deactive Oncomm use fro Timer
'                TimerScaleFlag = True
                
                TimerScale.Enabled = True
                Sleep 200
            
            ElseIf rctmp.Fields("DeviceTypeCode").Value = EnumDeviceType.Modem Then
            
                Dim tempstring As TextStream
                Dim ModemFile As String
                ModemFile = App.Path & "\Modem" & clsArya.StationNo & i & ".Log"
                If filetemp.FileExists(ModemFile) = False Then filetemp.CreateTextFile ModemFile
                If rctmp!DeviceCode <> EnumDevice.SmsCenter Then MainDevice = True
                
                Set tempstring = filetemp.OpenTextFile(ModemFile, ForWriting, False, TristateFalse)
                With mscSerial(i)

                    If rctmp!DeviceCode = EnumDevice.CallerIdInterface2_AlmP3 Then
                        .Output = Chr$(252) + Chr$(112) + Chr$(114) + Chr$(116) + Chr$(3) + Chr$(108) + Chr$(253)    ' INIT Protocol #3
'                    ElseIf rctmp!DeviceCode = EnumDevice.CallerIdInterface2_AlmP6 Then
'                        .Output = Chr$(252) + Chr$(112) + Chr$(114) + Chr$(116) + Chr$(6) + Chr$(108) + Chr$(253)    ' INIT Protocol #3
    
                    ElseIf rctmp!DeviceCode = EnumDevice.CallerIdInterface2_AlmP1 Then
                        .Output = Chr$(252) + Chr$(112) + Chr$(114) + Chr$(116) + Chr$(1) + Chr$(108) + Chr$(253)    ' INIT Protocol #1
                        TimerALM.Enabled = True
                        TimerALM.Interval = 300
                        mscSerial(i).PortOpen = False
                        mscSerial(i).InBufferSize = 1024
                        mscSerial(i).RThreshold = 0
                        mscSerial(i).PortOpen = True
                        AlmPort = i
                    ElseIf rctmp!DeviceCode = EnumDevice.CallerIdModem1 Then
                        .Output = "AT" + vbCrLf               ' INIT
                        tempstring.WriteLine (.Input)
                        Sleep (300)
                        .Output = "at+vcid=1" + vbCrLf 'EnableCallerID  (Cobra Lite 3)
                        tempstring.WriteLine (.Input)
                        Sleep (300)
                        .Output = "ATS0=0" + vbCrLf         'NO Answer    (ATS0=1 Answer After 1 Ring)
                        tempstring.WriteLine (.Input)
                                
                    ElseIf rctmp!DeviceCode = EnumDevice.CallerIdModem2 Then
                        .Output = "AT" + vbCrLf             ' INIT
                        tempstring.WriteLine (.Input)
                        Sleep (300)
                        .Output = "at#cid=1" + vbCrLf 'EnableCallerID  (Smart Spirit)
                        tempstring.WriteLine (.Input)
                        Sleep (300)
                        .Output = "ATS0=0" + vbCrLf         'NO Answer    (ATS0=1 Answer After 1 Ring)
                        tempstring.WriteLine (.Input)
                    End If
                    
                    .DTREnable = True
                End With
                tempstring.Close
            End If
            
            i = i + 1
        
        ElseIf rctmp!DeviceCode = EnumDevice.CallerIdInterface2_AlmP6 Then   ' CallerId ALM & Danzhe
        '   On Error Resume Next
            UCCallerIDMonitor1.PortNumber = Val(rctmp.Fields("PortNo").Value) '
            UCCallerIDMonitor1.Baudrate = rctmp.Fields("BaudRate").Value
            UCCallerIDMonitor1.OpenPort = True
            MainDevice = True
        
        ElseIf rctmp.Fields("PortNo").Value = 0 Then   ' Printer Port
            If rctmp.Fields("DeviceTypeCode").Value = EnumDeviceType.CashDrawer Then
               clsStation.DrawerModel = rctmp.Fields("DeviceCode").Value
               clsStation.DrawerPort = 0
            End If
        ElseIf rctmp.Fields("DeviceCode").Value = EnumDevice.RFT230 And rctmp.Fields("PortNo").Value <> 20 Then    ' Printer Port
            If HasRfidReader = True Then
                TimerRFID.Interval = IIf(clsStation.RfidInterval = "", 2000, clsStation.RfidInterval)
                MF_ExitComm
                RFIDStatus = MF_InitComm("Com" & rctmp.Fields("PortNo").Value, rctmp.Fields("BaudRate").Value)
                If RFIDStatus = 0 Then RfidReaderIsActive = True: ShowDisMessage "”Ì” „ ò«—  ŒÊ«‰ „«Ì›— ›⁄«· ‘œ", 1000 Else ShowDisMessage "«‘ò«· œ— « ’«· ”Ì” „ ò«—  ŒÊ«‰", 1500
                Sleep 500
            Else
                ShowDisMessage " œ— «Ì‰ ‰”ŒÂ «“ ‰—„ «›“«—«„ò«‰ ŒÊ«‰œ‰ ò«—  „«Ì›— ÊÃÊœ ‰œ«—œ ", 1500
            End If
        ElseIf rctmp.Fields("DeviceCode").Value = EnumDevice.RFT230 And rctmp.Fields("PortNo").Value = 20 Then    ' Printer Port
            If HasRfidReader = True Then
                TimerRFID.Interval = IIf(clsStation.RfidInterval = "", 2000, clsStation.RfidInterval)
                MF_ExitComm
                RFIDStatus = MF_InitComm("USB", rctmp.Fields("BaudRate").Value)
                If RFIDStatus = 0 Then RfidReaderIsActive = True: ShowDisMessage "”Ì” „ ò«—  ŒÊ«‰ „«Ì›— ›⁄«· ‘œ", 1000 Else ShowDisMessage "«‘ò«· œ— « ’«· ”Ì” „ ò«—  ŒÊ«‰", 1500
                Sleep 500
            Else
                ShowDisMessage " œ— «Ì‰ ‰”ŒÂ «“ ‰—„ «›“«—«„ò«‰ ŒÊ«‰œ‰ ò«—  „«Ì›— ÊÃÊœ ‰œ«—œ ", 1500
            End If
        ElseIf rctmp.Fields("PortNo").Value = 20 Then   ' USB Port
            UsbCallerIdIndex = i
            DeviceCode(i) = rctmp.Fields("DeviceCode").Value
            DeviceType(i) = rctmp.Fields("DeviceTypeCode").Value
            RThreshold(i) = rctmp.Fields("RThreshold").Value
            If rctmp!DeviceCode = EnumDevice.USBCallerID1 Then
                                        
                    ' this property help you to automatically
                    ' reconnect to the device if a PollingError event ocured
                    ' and the device lost. after enabling this the ocx automatically
                    ' search and try to reconnect to the device each 1500 ms
                    ' after device lost(PollingError) reached.
''''                    USBCallerID1.AutoReconnect = CLng(chkAutoReconnect.Value)
                    USBCallerID1.AutoReconnect = 1
                    
                    ' Recomended value for pooling the channel's is some around 250 ms
                    ' but you can change this value between 50 to 5000 milisecond.
                    USBCallerID1.PollingPeriod = 1000
                    
                    ' Now this metode open the device and the device start to work
                    ' if this methode sucessed it will return 'OK' string
                    ' and else it will return an error string which cused the error
                    ' this error may ocured if the device was not connected to a usb port
                    ' or the driver's was not installed completly.
                    If USBCallerID1.OpenDevice(0) = "Ok" Then
                      '  AddToLog "USB Caller ID1 has been Found and now is working properly."
                    Else
                     MsgBox " ÅÌœ« ‰‘œ USB Caller ID1" & vbLf & " .« ’«·«  Ê œ—«ÌÊ—Â« —« çﬂ ﬂ‰Ìœ "
                    ' MsgBox "Can Not Find USB Caller ID1, check connection's and driver's."
                     '   AddToLog "Can NOT find USB Caller ID1, check connection's and driver's."
                    End If
                    
             ElseIf rctmp!DeviceCode = EnumDevice.MagnetCardReader Then
                USBHID1.PortOpen = True
                If Not (USBHID1.PortOpen) Then
                   ' MsgBox "Couldn't open HID Swipe Reader"
                    MsgBox "«‘ﬂ«· œ— «— »«ÿ »« ﬂ«—  ŒÊ«‰"
                End If
             End If
        End If
        
        rctmp.MoveNext
    Wend
    
    If i <= mscSerial.Count Then
    
        i = i + 1
    End If
    
   
    Exit Sub
    
Error_Handler:

    Select Case err.Number
        Case 8015
            If rctmp.Fields("DeviceTypeCode").Value = EnumDeviceType.CustomerDisplay Then
                clsStation.CustDisplayPrn = True
            ElseIf rctmp.Fields("DeviceTypeCode").Value = EnumDeviceType.CashDrawer Then
                clsStation.CashDrawerPrn = True
            End If
            Resume Next
            
        Case Else
            MsgBox err.Description, , err.Source
             Resume Next
    End Select
End Sub

Private Sub TimerScale_Timer()
    TimerScale.Enabled = False
    On Error GoTo ErrHandler
    TimerScale.Enabled = False
'    If clsStation.BasculeModel = EnumDevice.MDS14000 Then
'        MDS14000CTL1.ReadCurrentWeight   'Open Port
'        txtScale.Text = MDS14000CTL1.CurrentWeight
'      '  txtScale.Text = 1260
'        txtScale.Text = Format(Val(txtScale.Text) / 1000, "##0.000")
'        lblScale(0).Caption = txtScale.Text
'
'        If Val(txtScale.Text) > 10 Then
'            lblScale(0).Font.Size = 18
'        Else
'            lblScale(0).Font.Size = 20
'        End If
'    ElseIf clsStation.BasculeModel = EnumDevice.MDS11000 Then
'        Mahak11000.Model = MDS15000
'        Mahak11000.Network = False
'        If Mahak11000.ReadWeight(1) = True Then   'Open Port
'            txtScale.Text = Mahak11000.Answer(1)
'        Else
'          '  txtScale.Text = ""
'            TimerScale.Enabled = True
'            Exit Sub
'        End If
'        'txtScale.Text = "872"
'        If txtScale.Text <> "" Then txtScale.Text = Format(Val(txtScale.Text) / 1000, "##0.000")
'        'If Val(txtScale.Text) < 0 Then txtScale.Text = 0
'
'        lblScale(1).Caption = txtScale.Text
'
'        If Val(txtScale.Text) > 10 Then
'            lblScale(1).Font.Size = 18
'        Else
'            lblScale(1).Font.Size = 20
'        End If
'        lblScale(1).Visible = True
'        ShapeScale(1).Visible = True
'        BascoleLabel(1).Visible = True
    If clsStation.BasculeModel = EnumDevice.Mahak_Serial Then
        If MahakScaleOCX3.ReadWeight(1) = True Then    'Open Port
            txtScale.Text = MahakScaleOCX3.Answer(1)
        Else
            txtScale.Text = "0"
        End If
        'txtScale.Text = "872"
        If txtScale.Text <> "" Then txtScale.Text = Format(Val(txtScale.Text) / 1000, "##0.000")
        'If Val(txtScale.Text) < 0 Then txtScale.Text = 0

        lblScale(0).Caption = txtScale.Text

        If Val(txtScale.Text) > 10 Then
            lblScale(0).Font.size = 18
        Else
            lblScale(0).Font.size = 20
        End If
'    ElseIf clsStation.BasculeModel = EnumDevice.Mahak Or clsStation.BasculeModel = EnumDevice.Pand Then
'       If Me.mscSerial(clsStation.BasculePort).PortOpen = True Then
'           SetBascoleController ClsBascole(clsStation.BasculePort), Me.mscSerial(clsStation.BasculePort), lblScale(clsStation.BasculePort), BascoleLabel(clsStation.BasculePort), ShapeScale(clsStation.BasculePort)
'       End If
    ElseIf clsStation.BasculeModel = EnumDevice.Sairan Then
       If Me.mscSerial(clsStation.BasculePort).PortOpen = True Then
           mscSerial(clsStation.BasculePort).Output = "*"
           SairanFlag = False
       End If
    End If
    TimerScale.Enabled = True

Exit Sub

ErrHandler:
    
    LogSaveNew "FrmInvoice => ", err.Description, err.Number, err.Source, "TimerScale_Timer"

    Select Case err.Number
        Case 6
'            txtScale.Text = "0.000"
'            lblScale(1).Caption = txtScale.Text
            Sleep 20
            TimerScale.Enabled = True
        Case 8018
'            ShowDisMessage Err.Description, 1000
'            txtScale.Text = "0.000"
'            lblScale(1).Caption = txtScale.Text
            Sleep 20
            TimerScale.Enabled = True
        Case Else
            TimerScale.Enabled = False
            ShowDisMessage err.Description, 1000
            txtScale.Text = "0.000"
            lblScale(0).Caption = txtScale.Text
    End Select
   ' TimerScale.Enabled = False
   ' Resume Next
End Sub

Public Sub SetFWModemSetting(index As Integer)
    
    FWModem(index).BackColor = &H80000016  '&H808000
    FWModem(index).ToolTipText = ""
'    Call_RealNumber = ""
    Call_Number(index + 1) = ""

End Sub
Private Sub FWModem_MouseMove(index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    
'    FWModem(Index).ToolTipText = Call_Number(Index + 1)

End Sub
Public Sub FWModem_MouseDown(index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

    If clsStation.CallerIdTest = False Then  ' ###############
        
        If MaxRowFlexGrid = 1 And MyFormAddEditMode = AddMode And lblCustomer.Tag = "-1" Then
            If FWModem(index).BackColor = vbRed Or (FWModem(index).BackColor = vbGreen And AlmPort = 1) Then
                FWModem(index).BackColor = vbGreen   ' &H80000003&
'                FWModem(Index).ToolTipText = Call_Number(Index + 1)
                Call_RealNumber = Call_Number(index + 1)
              '  ModemPriority(Index + 1) = 0
                If clsStation.NetworkCallerId = True Then
                    mdifrm.WinsockUdp.SendData Str(index) & clsArya.StationNo
'                    Sleep 100
'                    mdifrm.WinsockUdp.SendData Str(Index) & clsArya.StationNo
                End If
                FindCust
                If Val(lblCustomer.Tag) > 0 Then
                    ReDim Parameter(3) As Parameter
                    
                    Parameter(0) = GenerateInputParameter("@nvcCallerId", adWChar, 20, Trim(Call_Number(index + 1)))
                    Parameter(1) = GenerateInputParameter("@intCustomer", adInteger, 4, Val(lblCustomer.Tag))
                    Parameter(2) = GenerateInputParameter("@MembershipId", adBigInt, 8, mvarMemberShipId)
                    Parameter(3) = GenerateInputParameter("@nvcname", adWChar, 50, Trim(lblCustomer.Caption))
                    
                    RunParametricStoredProcedure "Update_tblTotal_CallerId_Number", Parameter
                End If
                Call_Number(index + 1) = ""
            End If
        End If
    
    Else    ' Test CallerId In Network  ###############################
                    
        ShowDisMessage "Caller Id Test", 500
        
        If FWModem(index).BackColor = vbRed Then
            FWModem(index).BackColor = vbGreen   ' &H80000003&
'                FWModem(Index).ToolTipText = Call_Number(Index + 1)
            Call_RealNumber = Call_Number(index + 1)
          '  ModemPriority(Index + 1) = 0
            If clsStation.NetworkCallerId = True Then
                mdifrm.WinsockUdp.SendData Str(index) & clsArya.StationNo
'                Sleep 100
'                mdifrm.WinsockUdp.SendData Str(Index) & clsArya.StationNo
            End If
            FindCust
            If Val(lblCustomer.Tag) > 0 Then
               ReDim Parameter(3) As Parameter
               
               Parameter(0) = GenerateInputParameter("@nvcCallerId", adWChar, 20, Trim(Call_Number(index + 1)))
               Parameter(1) = GenerateInputParameter("@intCustomer", adInteger, 4, Val(lblCustomer.Tag))
               Parameter(2) = GenerateInputParameter("@MembershipId", adBigInt, 8, mvarMemberShipId)
               Parameter(3) = GenerateInputParameter("@nvcname", adWChar, 50, Trim(lblCustomer.Caption))
               
               RunParametricStoredProcedure "Update_tblTotal_CallerId_Number", Parameter
            End If
            Call_Number(index + 1) = ""
        Else
            If aaaa = 0 Then aaaa = 88554455
            aaaa = aaaa + 1
            Sleep 100
            If clsStation.NetworkCallerId = True Then
                mdifrm.WinsockUdp.SendData Str(index + 1) & Str(index) & Str(aaaa)
'                Sleep 100
'                mdifrm.WinsockUdp.SendData Str(Index + 1) & Str(Index) & Str(aaaa)
            End If
            FWModem(index).BackColor = vbRed
            Call_Number(index + 1) = aaaa
            FWModem(index).ToolTipText = Call_Number(index + 1)
            ReDim Parameter(2) As Parameter
            Parameter(0) = GenerateInputParameter("@nvcDate", adWChar, 8, mvarDate)
            Parameter(1) = GenerateInputParameter("@LineNumber", adTinyInt, 1, index + 1)
            Parameter(2) = GenerateInputParameter("@nvcCallerId", adWChar, 20, Call_Number(index + 1))
            RunParametricStoredProcedure "Insert_tblTotal_CallerId", Parameter
        End If
    End If
End Sub

Public Sub GetCallerInfo(ByVal index As Integer, ByVal Caller_Id_string As String, ByVal LineNumber As Integer)

On Error GoTo Err_Handler

If intVersion = Min Then Exit Sub

Dim lR As Long
Dim varForm As Form
Dim frmact As Form
CallerIdformshow = False
For Each varForm In Forms
    If varForm.Name = "frmCallerIdView" Then 'frmCallerIdView
        Set frmact = varForm
        CallerIdformshow = True
        Exit For
    End If
Next
      
Select Case MainDevice
    Case True
            
            Dim Call_NumberTemp As String
            Dim kk As Integer
            
            If InStr(Caller_Id_string, "NMBR") > 0 Then
            
                Call_NumberTemp = Mid(Caller_Id_string, (InStr(Caller_Id_string, "NMBR = ") + 7))
            ElseIf InStr(Caller_Id_string, "L") > 0 Then
                kk = InStr(1, Caller_Id_string, "L", 1)
                Call_NumberTemp = Mid(Caller_Id_string, kk + 2)
                  
            End If
            
            If InStr(LCase(Call_NumberTemp), "callerid:") > 0 Then
            
                If DeviceCode(index) = EnumDevice.CallerIdInterface2_AlmP1 Then
                    Dim InID As String
                    InID = Trim(Mid(Call_NumberTemp, InStr(1, Call_NumberTemp, ":CallerID", vbTextCompare) + 10, 22))
                    InID = FNSplitt(InID)
                    Call_NumberTemp = InID 'Val(Left(ID, 15))
                Else
                
                    Call_NumberTemp = LTrim(Mid(Call_NumberTemp, InStr(1, LCase(Call_NumberTemp), "callerid:") + 9))
                    Call_NumberTemp = Mid(Call_NumberTemp, clsStation.CallerIdSpace, 15)
                End If
            End If
            
            If Val(Call_NumberTemp) = 0 Then Call_NumberTemp = Caller_Id_string
            Call_NumberTemp = Val(Call_NumberTemp)  ' Discard RING AND "0" in First
            
            If left(Call_NumberTemp, 2) = "98" Then  ' Discard Country Code
                Call_NumberTemp = Mid(Call_NumberTemp, 3)
            End If
            
            If Mid(Call_NumberTemp, 1, 2) = "91" Or Mid(Call_NumberTemp, 1, 2) = "93" Then       ' Mobile Phone Number
                Call_NumberTemp = "09" & Mid(Call_NumberTemp, 2, 2) & Right(Call_NumberTemp, 7) ' Define Mobile Phone
            
            ElseIf Mid(Call_NumberTemp, 1, Len(clsStation.CityCode)) = clsStation.CityCode Then       ' Tehran Phone Number
                Call_NumberTemp = Mid(Call_NumberTemp, Len(clsStation.CityCode) + 1, clsStation.NumberOfId) ' Variable  Number And Discard City Code
            
            Else                              ' Other City  Phone Number
                Call_NumberTemp = Right(Call_NumberTemp, clsStation.NumberOfId)  ' Discard City Code
            End If
            If (Len(Call_NumberTemp) > 8 And left(Call_NumberTemp, 1) <> "0") Then Call_NumberTemp = "0" & Call_NumberTemp   ' call from other city
    
    'Save CallerId In Database
            ReDim Parameter(2) As Parameter
            Parameter(0) = GenerateInputParameter("@nvcDate", adWChar, 8, mvarDate)
            Parameter(1) = GenerateInputParameter("@LineNumber", adTinyInt, 1, LineNumber)
            Parameter(2) = GenerateInputParameter("@nvcCallerId", adWChar, 20, Call_NumberTemp)
            RunParametricStoredProcedure "Insert_tblTotal_CallerId", Parameter

            If clsStation.NetworkCallerId = True Then
                mdifrm.WinsockUdp.SendData Str(LineNumber) & Str(index) & Call_NumberTemp
              '  LogSaveNew Inputstr, "", "", "", ""
                LogSaveNew "Network Send: " & Str(LineNumber) & Str(index) & Call_NumberTemp, "", "", ""
'                Sleep 200
'                mdifrm.WinsockUdp.SendData Str(LineNumber) & Str(Index) & Call_NumberTemp
            End If
    Case False
            
        Call_NumberTemp = Val(Caller_Id_string)

End Select

        If index > -1 Then
            If (DeviceCode(index) <> 63 And DeviceCode(index) <> 64 And DeviceCode(index) <> 65 And DeviceCode(index) <> 66 And DeviceCode(index) <> 67 And DeviceCode(index) <> 0) Then
                If MaxRowFlexGrid = 1 And MyFormAddEditMode = AddMode And lblCustomer.Tag = "-1" And Len(Call_NumberTemp) > 1 And Call_Priority = 0 Then
                    FWModem(0).BackColor = vbGreen   ' &H80000003&
                    FWModem(0).ToolTipText = Call_NumberTemp    ' &H80000003&
                    Call_RealNumber = Call_NumberTemp
                    FindCust
                    Call_NumberTemp = ""
                  '  FWModem(Index - 1).BackColor = &H80000003
                ElseIf Len(Call_NumberTemp) > 1 Then
    '''                If clsStation.Callwaiting = True Then
                        If Call_Priority = 0 Then Call_Priority = index
                        Call_Number(LineNumber) = Call_NumberTemp
                        For i = 1 To 4
                            If Val(ModemPriority(i)) = 0 Then
                               ModemPriority(i) = index
                               Exit For
                            End If
                        Next i
    '''                Else
    '''                    FWModem(Index - 1).BackColor = &H808000
    '''                End If
                End If
            Else            ' Segal & Khazama & ALM
                
                If clsStation.AutoCallerId = True And MaxRowFlexGrid = 1 And MyFormAddEditMode = AddMode And lblCustomer.Tag = "-1" And Len(Call_NumberTemp) > 1 And Call_Priority = 0 Then
                    FWModem(LineNumber - 1).ToolTipText = Call_NumberTemp   ' &H80000003&
'                    If DeviceCode(Index) = 66 Then
                        FWModem(LineNumber - 1).BackColor = vbGreen   ' &H80000003&
'                    End If
                    Call_RealNumber = Call_NumberTemp
                    If clsStation.NetworkCallerId = True Then
                        mdifrm.WinsockUdp.SendData Str(index) & clsArya.StationNo
'                        Sleep 100
'                        mdifrm.WinsockUdp.SendData Str(Index) & clsArya.StationNo
                    End If
                    FindCust
                    If Val(lblCustomer.Tag) > 0 Then
                        ReDim Parameter(3) As Parameter
                        
                        Parameter(0) = GenerateInputParameter("@nvcCallerId", adWChar, 20, Trim(Call_NumberTemp))
                        Parameter(1) = GenerateInputParameter("@intCustomer", adInteger, 4, Val(lblCustomer.Tag))
                        Parameter(2) = GenerateInputParameter("@MembershipId", adBigInt, 8, mvarMemberShipId)
                        Parameter(3) = GenerateInputParameter("@nvcname", adWChar, 50, Trim(lblCustomer.Caption))
                        
                        RunParametricStoredProcedure "Update_tblTotal_CallerId_Number", Parameter
                    End If
                    Call_NumberTemp = ""
                ElseIf Len(Call_NumberTemp) > 1 Then
                    FWModem(LineNumber - 1).ToolTipText = Call_NumberTemp   ' &H80000003&
    '''                If clsStation.Callwaiting = True Then
                        If Call_Priority = 0 Then Call_Priority = LineNumber
                        Call_Number(LineNumber) = Call_NumberTemp
                        For i = 1 To 8
                            If Val(ModemPriority(i)) = 0 Then
                               ModemPriority(i) = LineNumber
                               Exit For
                            End If
                        Next i
    '''                Else
    '''                    FWModem(Val(Mid(Caller_Id_string, kk + 1, 1)) - 1).BackColor = &H808000
    '''                End If
                    ' ›—„ ‰„«Ì‘ «ê— ﬁ»·« »«“ «”  »«Ìœ »—«Ì ‰„«Ì‘ „Ãœœ »” Â ‘Êœ
                    If clsStation.CallerIdAutoView = True And CallerIdformshow = True Then
                        Unload frmCallerIdView
                    End If
                    If clsStation.CallerIdAutoView = True Then TimerAlmP6.Enabled = True
                    '  «Ì„— »—«Ì «Ì‰ «”  òÂ ›—’  »—«Ì Œ—ÊÃ «“ —Ê Ì‰ «Ì‰ —«Å  ÅÊ—  ”—Ì«· ÊÃÊœ œ«‘ Â »«‘œ  « ÅÊ—  „Ãœœ« »—«Ì Ê—ÊœÌ »⁄œÌ ¬„«œÂ ê—œœ

                End If
            
            End If
        Else ' if is a Netowrk Caller ID
'            If MaxRowFlexGrid = 1 And MyFormAddEditMode = AddMode And lblCustomer.Tag = "-1" And Len(Call_NumberTemp) > 1 And Call_Priority = 0 Then
'                FWModem(clsArya.StationNo - 1).BackColor = vbGreen   ' &H80000003&
'                FWModem(clsArya.StationNo - 1).ToolTipText = Call_NumberTemp    ' &H80000003&
'                Call_RealNumber = Call_NumberTemp
'                FindCust
'                Call_NumberTemp = ""
'            End If
            If LineNumber < 9 Then
                FWModem(LineNumber - 1).BackColor = vbRed   ' &H80000003&
                FWModem(LineNumber - 1).ToolTipText = Call_NumberTemp    ' &H80000003&
            End If
            Call_Number(LineNumber) = Call_NumberTemp
            Call_NumberTemp = ""
            'Test
            If clsStation.CallerIdAutoView = True And CallerIdformshow = True Then
                Unload frmCallerIdView
            End If
            If clsStation.CallerIdAutoView = True Then TimerAlmP6.Enabled = True

            
'            If clsStation.CallerIdAutoView = True And CallerIdformshow = True Then
'                Unload frmCallerIdView: LastRecordshow = True: frmCallerIdView.Show 'vbModal
'            ElseIf clsStation.CallerIdAutoView = True Then frmCallerIdView.Show 'vbModal
'            End If
'            lR = SetTopMostWindow(frmCallerIdView.hwnd, True)
        End If
        
Exit Sub

Err_Handler:
'    'MsgBox "Callerinfo" & err.Description
'    ShowErrorMessage
'    LogSaveNew "frmInvoice => ", err.Description, err.Number, err.Source, "GetCallerInfo"
'    Resume Next
    ShowDisMessage err.Description, 1500
End Sub

Sub AddEmptyRow()

    With FlxDetail
        .Rows = .Rows + 1
    End With
    
End Sub

Private Function FindRecord_FlexGrid(TempGoodCode As Double) As Boolean

    FindRecord_FlexGrid = False
    Dim FirstRowGood As Integer
    If (TempGoodCode = Val(FlxDetail.TextMatrix(FlxDetail.Row, 5))) And Val(lblNum.Caption) < 0 Then ' for decrease amount good
        FindRecord_FlexGrid = True
        Exit Function
    End If
    Dim jj As Integer
    If (TempGoodCode = Val(FlxDetail.TextMatrix(FlxDetail.Row, 5))) And ((mvarServePlace = Val(FlxDetail.TextMatrix(FlxDetail.Row, 8))) Or Val(lblNum.Caption) < 0) And (Trim(FlxDetail.TextMatrix(FlxDetail.Row, 10)) = "") Then
        FindRecord_FlexGrid = True
        Exit Function
    End If
    Dim flagGood As Boolean
    flagGood = False
    FirstRowGood = -1
     For jj = 1 To FlxDetail.Rows - 1
        If (TempGoodCode = Val(FlxDetail.TextMatrix(jj, 5))) And (mvarServePlace = Val(FlxDetail.TextMatrix(jj, 8))) And (Val(FlxDetail.TextMatrix(jj, 3)) = mvarSellPrice) And (Trim(FlxDetail.TextMatrix(jj, 10)) = "") Then
            If flagGood = False Then
                FirstRowGood = jj
                flagGood = True
            ElseIf FlxDetail.TextMatrix(jj, 10) = "" Then
                FirstRowGood = jj
                Exit For
            End If
        End If
    Next
    If FirstRowGood <> -1 Then
        FindRecord_FlexGrid = True
        FlxDetail.Row = FirstRowGood
    End If
End Function
Sub ClearDataFlexGrid()

    With FlxDetail
        .Rows = 1
        .Rows = MaxInvoiceRows
        .Row = 1
        MaxRowFlexGrid = 1
                
    End With
    If clsInvoiceValue.ShowInvoiceMenu = True Then
        frmShowInvoiceMenu.ClearGridValue
    End If
    If clsInvoiceValue.ShowLogo = True Then
        frmShowLogo.ClearGridValue
    End If
    
End Sub
Sub GetDetails()
    
    framelastFich.Visible = False
    ReDim Parameter(4) As Parameter
    
    Parameter(0) = GenerateInputParameter("@No", adBigInt, 8, Val(txtNo.Text))
    Parameter(1) = GenerateInputParameter("@Status", adInteger, 4, mvarStatus)
    Parameter(2) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
    Parameter(3) = GenerateInputParameter("@AccountYear", adSmallInt, 2, AccountYear)
    Parameter(4) = GenerateInputParameter("@Branch", adInteger, 4, CurrentBranch)
    
    Set Rst = RunParametricStoredProcedure2Rec("Get_FacMD_Good", Parameter)
    
    Dim ii As Integer

    If Not (Rst.BOF Or Rst.EOF) Then
        txtDescription.Text = IIf(IsNull(Rst!NvcDescription), "", Rst!NvcDescription)
        TxtTempAddress.Text = IIf(IsNull(Rst!TempAddress), "", Rst!TempAddress)
        If Len(Trim(TxtTempAddress)) > 0 Then TempAddressEdit = True
        TmpGoodDiscount = 0
'        InventoryNo = Rst!IntInventoryNo
        boolPayment = Rst!FacPayment
        mVarOrderType = Rst!OrderType
        
        With FlxDetail
        Do While Not (Rst.EOF)
            
            ii = ii + 1
            .TextMatrix(ii, 0) = ii 'Number
            .TextMatrix(ii, 1) = Rst!amount
            .TextMatrix(ii, 2) = Rst!nvcName 'GoodName
            .TextMatrix(ii, 3) = Rst!FeeUnit
            .TextMatrix(ii, 4) = Rst!amount * Rst!FeeUnit ' rst!FeeTotal
            .TextMatrix(ii, 5) = Rst!GoodCode
            .TextMatrix(ii, 6) = Rst!Weight ' rst!WeightUnit
            .TextMatrix(ii, 7) = Rst!Unit
            .TextMatrix(ii, 8) = Rst!ServePlace
            .TextMatrix(ii, 9) = IIf(IsNull(Rst!DifferencesCodes), "", Rst!DifferencesCodes)
            .TextMatrix(ii, 10) = IIf(IsNull(Rst!DifferencesDescription), "", Rst!DifferencesDescription)
            If Rst!FeeUnit <> 0 Then
                .TextMatrix(ii, 11) = Rst!Discount
            End If
            .TextMatrix(ii, 12) = Rst!Rate
            .TextMatrix(ii, 13) = IIf(IsNull(Rst!ChairName), "", Rst!ChairName)
            .TextMatrix(ii, 14) = Rst!intInventoryNo
            .TextMatrix(ii, 15) = Rst!mainType
            
            If Rst.Fields("Mojodi").Value >= 0 Then
                If Rst.Fields("Mojodi").Value <> Int(Rst.Fields("Mojodi").Value) Then
                    .TextMatrix(ii, 16) = Format(Rst.Fields("Mojodi").Value, "##.000")
                    .TextMatrix(ii, 16) = Val(.TextMatrix(ii, 16)) ' Delete Last Zeros
                Else
                     .TextMatrix(ii, 16) = Rst.Fields("Mojodi").Value
                End If
            Else
                If Rst.Fields("Mojodi").Value <> Int(Rst.Fields("Mojodi").Value) Then
                    .TextMatrix(ii, 16) = -Format(Rst.Fields("Mojodi").Value, "##.000")
                    .TextMatrix(ii, 16) = Val(.TextMatrix(ii, 16)) & "-" ' Delete Last Zeros
                Else
                     .TextMatrix(ii, 16) = -Rst.Fields("Mojodi").Value & "-"
                End If
            End If
            
            .TextMatrix(ii, 17) = Rst!DutySale
            .TextMatrix(ii, 18) = Rst!TaxSale
            TmpGoodDiscount = TmpGoodDiscount + (Rst!Discount * Rst!amount * Rst!FeeUnit / 100)
            
            Rst.MoveNext
            
            MaxRowFlexGrid = ii + 1
            If ii >= FlxDetail.Rows - 2 And Rst.EOF = False Then
                AddEmptyRow
            End If

        Loop
        End With
        FlxDetail.Row = MaxRowFlexGrid - 1
        mvarServePlace = FlxDetail.TextMatrix(MaxRowFlexGrid - 1, 8)
        For i = 0 To cmbServePlace.ListCount - 1
            If mvarServePlace = cmbServePlace.ItemData(i) Then
                cmbServePlace.ListIndex = i
                Exit For
            End If
        Next i
        
    End If
    
    Dim CountPrinting, CountRePrint, CountInvoicePrint As Integer
    
    ReDim Parameter(2) As Parameter
    Parameter(0) = GenerateInputParameter("@No", adBigInt, 8, Val(txtNo.Text))
    Parameter(1) = GenerateInputParameter("@Status", adInteger, 4, mvarStatus)
    Parameter(2) = GenerateInputParameter("@AccountYear", adSmallInt, 2, AccountYear)
    Set Rst = RunParametricStoredProcedure2Rec("Get_CountPrint_tAction", Parameter)
    
    CountPrinting = Rst!CountPrinting
    CountRePrint = Rst!CountRePrint
    CountInvoicePrint = Rst!CountInvoicePrint
    
    If Rst.State <> 0 Then Rst.Close
    LblInvoicePrint.Caption = CountInvoicePrint
End Sub

Sub GetDataDetail()
    
    If ClsFormAccess.LockCheck = True Then
        Me.ChkIsLocked.Enabled = True
    Else
        Me.ChkIsLocked.Enabled = False
    End If
    
    DefaultValueLables
    
    txtDescription.Text = ""
    sFactorReceived = ""
    ClearDataFlexGrid
    GetDetails
    If strCategory = "07" Then
        SplitFlag = False
        FWBtnSplit_Click
    End If
    If MaxRowFlexGrid = 1 Then
        Exit Sub
    End If
    
    cmbTable.ListIndex = 0
    cmbGarson.ListIndex = 0
    CmbPayk.ListIndex = 0
    
    ReDim Parameter(3) As Parameter
    
    Parameter(0) = GenerateInputParameter("@No", adBigInt, 8, Val(txtNo.Text))
    Parameter(1) = GenerateInputParameter("@Status", adInteger, 4, mvarStatus)
    Parameter(2) = GenerateInputParameter("@AccountYear", adSmallInt, 2, AccountYear)
    Parameter(3) = GenerateInputParameter("@Branch", adInteger, 4, CurrentBranch)
    
    Set Rst = RunParametricStoredProcedure2Rec("Get_FacM_Per", Parameter)
    
    
    sbrFactorProp.Panels(1).Text = ""
    sbrFactorProp.Panels(2).Text = ""
    sbrFactorProp.Panels(3).Text = ""
    sbrFactorProp.Panels(4).Text = ""
    sbrFactorProp.Panels(5).Text = ""
    If Not (Rst.EOF = True And Rst.BOF = True) Then
    
        chKTax.Value = Rst!Rasmi
        chKTax_Click
        If IsNull(Rst!tempNo) = True Then
            FWLedTemp.Value = FWLed1.Value
        Else
            FWLedTemp.Value = Rst!tempNo
        End If
        If Rst.Fields("Serveplace").Value = Salon Or Rst.Fields("Serveplace").Value = Delivery Or Rst.Fields("Serveplace").Value = Car Or Rst.Fields("Serveplace").Value = Table Then        ' For View Serveplace
             FlxDetail.ColHidden(8) = True
             FlxDetail.ColWidth(10) = FlxDetail.Width / 4.5     'Diffrence
        Else
             FlxDetail.ColHidden(8) = False
             FlxDetail.ColWidth(10) = FlxDetail.Width / 8     'Diffrence
        End If
        If mVarOrderType = ByPhone Then         ' ByPhone
           If clsStation.Language = Farsi Then
                LblOrder.Caption = " ·›‰Ì"
           Else
                LblOrder.Caption = "By phone"
           End If
        Else
            If clsStation.Language = Farsi Then
              LblOrder.Caption = "Õ÷Ê—Ì"
           Else
              LblOrder.Caption = "Inside"
           End If
        End If
            
        BalancePayment = Rst.Fields("Balance").Value
        dblFichUser = Rst.Fields("User").Value
        intSerialNo = Rst.Fields("intSerialNo").Value
        mvarStationNo = Rst.Fields("StationId").Value
        FWScrolltextPay.Visible = True
        
        If BalancePayment = True Then
           If clsStation.Language = Farsi Then
                FWScrolltextPay.Caption = "Å—œ«Œ  ‘œÂ"
           Else
                FWScrolltextPay.Caption = "Recieved"
           End If
           FWScrolltextPay.BackColor = vbGreen
        Else
           If clsStation.Language = Farsi Then
                FWScrolltextPay.Caption = " ”ÊÌÂ ‰‘œÂ"
           Else
                FWScrolltextPay.Caption = "Not Recieved"
           End If
           FWScrolltextPay.BackColor = vbRed
        End If
        If Rst!FacPayment = True Then
            LblFacpayment.BackColor = vbGreen
        Else
            LblFacpayment.BackColor = vbRed
        End If
        If Rst.Fields("ServePlace").Value = 2 And Rst.Fields("Incharge").Value = 0 Then
           FWScrollSend.Visible = True
           If clsStation.Language = Farsi Then
                FWScrollSend.Caption = "«—”«· ‰‘œÂ"
           Else
                FWScrollSend.Caption = "Not Send"
           End If
           FWScrollSend.BackColor = vbYellow
        ElseIf Rst.Fields("ServePlace").Value = 2 And Rst.Fields("Incharge").Value <> 0 Then
           If clsStation.Language = Farsi Then
                FWScrollSend.Caption = "«—”«· ‘œÂ"
           Else
                FWScrollSend.Caption = "Sent"
           End If
           FWScrollSend.BackColor = vbBlue
           FWScrollSend.Visible = True
        Else
           FWScrollSend.Visible = False
        End If
        FWChkHavale.Value = Rst!BitHavaleResid
        LblAccNo.Caption = "”‰œ ‘„«—Â : " & IIf(IsNull(Rst!Refrence_Acc), "", Rst!Refrence_Acc)
        Refrence_Acc = IIf(IsNull(Rst!Refrence_Acc), 0, Rst!Refrence_Acc)
        FWChkAccount.Value = IIf(IsNull(Rst!TransferAccounting), 0, Rst!TransferAccounting)
        
        Dim Parameters2(2) As Parameter
        Parameters2(0) = GenerateInputParameter("@intSerialNo", adBigInt, 8, intSerialNo)
        Parameters2(1) = GenerateInputParameter("@Branch", adInteger, 4, CurrentBranch)
        Parameters2(2) = GenerateOutputParameter("@EditedFich", adInteger, 4)

        mvarEditedFich = RunParametricStoredProcedure2String("Get_Edited_Factors", Parameters2)
        
        If mvarEditedFich > 0 Then
            FWLblEdit.Visible = True
        Else
            FWLblEdit.Visible = False
        End If
        
        Select Case clsStation.Language
            Case EnumLanguage.Farsi
                    sbrFactorProp.Panels(1).Text = "ò«—»—" & " : " & Rst.Fields("nvcFirstName").Value & " " & Rst.Fields("nvcSurName").Value
                    sbrFactorProp.Panels(2).Text = " «—ÌŒ : " & Rst.Fields("Regdate").Value
                    sbrFactorProp.Panels(3).Text = "”«⁄  : " & Rst.Fields("Time").Value
                    sbrFactorProp.Panels(4).Text = "‘Ì›  : " & Rst.Fields("ShiftDescription").Value
                    sbrFactorProp.Panels(4).Tag = Rst.Fields("ShiftNo").Value
                    sbrFactorProp.Panels(5).Text = "«Ì” ê«Â : " & Rst.Fields("StationId").Value
            
            Case EnumLanguage.English
            
                sbrFactorProp.Panels(1).Text = "User : " & Rst.Fields("nvcFirstName").Value & " " & Rst.Fields("nvcSurName").Value
                sbrFactorProp.Panels(2).Text = "Date : " & Rst.Fields("Regdate").Value
                sbrFactorProp.Panels(3).Text = "Time : " & Rst.Fields("Time").Value
                sbrFactorProp.Panels(4).Text = "Shift : " & Rst.Fields("ShiftNo").Value
                sbrFactorProp.Panels(4).Tag = Rst.Fields("ShiftNo").Value
                sbrFactorProp.Panels(5).Text = "StationId :" & Rst.Fields("StationId").Value
            
        End Select
                
        If ViewFlag = False Then
            FillsFullTableCombo
            ViewFlag = True
        End If
        cmbTableName = ""
        cmbTableData = 0
        
        For i = 0 To cmbTable.ListCount - 1
            If Rst!TableNo = cmbTable.ItemData(i) Then
                cmbTable.ListIndex = i
                cmbTableName = cmbTable.Text
                cmbTableData = Val(Rst!TableNo)
                Exit For
            End If
        Next i
        
        For i = 0 To cmbGarson.ListCount - 1
            If Rst.Fields("incharge").Value = cmbGarson.ItemData(i) Then
                cmbGarson.ListIndex = i
                Exit For
            End If
        Next i
         
        For i = 0 To CmbPayk.ListCount - 1
            If Rst.Fields("incharge").Value = CmbPayk.ItemData(i) Then
                CmbPayk.ListIndex = i
                Exit For
            End If
        Next i
        
'        If Rst!DeliveryFullName <> "--" Then
'            lblDeliveryFullName.Caption = " ÅÌò : " & Rst!DeliveryFullName
'        Else
'            lblDeliveryFullName.Caption = ""
'        End If
        
        If Rst.Fields("BitLock").Value <> 0 Then
            Me.ChkIsLocked.Value = 1
        Else
            Me.ChkIsLocked.Value = 0
        End If
        
        DetailsString1 = ""
        With FlxDetail
            For i = 1 To MaxRowFlexGrid - 1
                DetailsString1 = GenerateDetailsString3(DetailsString1, .TextMatrix(i, 1), .TextMatrix(i, 5), Val(.TextMatrix(i, 3)), Val(.TextMatrix(i, 11)), Val(.TextMatrix(i, 12)), .TextMatrix(i, 13), " ", .TextMatrix(i, 14), "", .TextMatrix(i, 8), .TextMatrix(i, 9))
            Next i
        End With
        Me.TxtGuestNo.Text = IIf(IsNull(Rst.Fields("GuestNo")) = True, 0, Rst.Fields("GuestNo").Value)
        
        BeforEditInvoice.No = Rst!No
        BeforEditInvoice.Status = Rst!Status
        BeforEditInvoice.Owner = 0
        BeforEditInvoice.Recursive = Val(Me.txtRecursive.Text)
        BeforEditInvoice.Incharge = Rst!Incharge
        BeforEditInvoice.FacPayment = Rst!FacPayment
        BeforEditInvoice.OrderType = mVarOrderType
        BeforEditInvoice.StationId = Rst!StationId
        BeforEditInvoice.BascoleNo = 0
        BeforEditInvoice.TableNo = Rst!TableNo
        BeforEditInvoice.User = Rst!User
        BeforEditInvoice.DateInvoice = txtDate.Text
        BeforEditInvoice.DetailsString = DetailsString1
        BeforEditInvoice.sFactorReceived = sFactorReceived
        BeforEditInvoice.Balance = Abs(CInt(BalancePayment))
        BeforEditInvoice.AccountYear = Rst!AccountYear
        BeforEditInvoice.NvcDescription = Right(txtDescription.Text, 150)
        BeforEditInvoice.TempAddress = TxtTempAddress.Text
        BeforEditInvoice.GuestNo = Val(TxtGuestNo.Text)
        
    End If
    ReDim Parameter(1) As Parameter
    
    Parameter(0) = GenerateInputParameter("@intSerialNo", adBigInt, 8, intSerialNo)
    Parameter(1) = GenerateInputParameter("@Branch", adInteger, 4, CurrentBranch)
    
    Set Rst = RunParametricStoredProcedure2Rec("Get_PayFactors_CustCredit_Account", Parameter)
    
    PreReceived = Rst!Bestankar1
    lblPayFactorTotal.Caption = PreReceived + Rst!Bestankar2
    lblSumPrice.Caption = Val(lblSumPrice.Tag) - Val(lblPayFactorTotal.Caption)
    lblSumPrice.Tag = lblSumPrice.Caption
    lblSumPrice.Caption = Format(lblSumPrice.Caption, "#,##")
        
    Set Rst = Nothing
    If clsInvoiceValue.ShowInvoiceMenu = True Then
        frmShowInvoiceMenu.UpdateGridValue
    End If
End Sub

Sub DefaultValueLables()
        
        Refrence_Acc = 0
        TxtGuestNo.Text = 0
        AutoDiscountValue = 0
        PreReceived = 0
        LblTip.Caption = ""
        AddressFlag = False
        txtSumCountNo.Caption = 0
'        txtSumCountWeight.Caption = 0
        txtSumFeeTotal.Text = 0
        txtDiscount.Text = 0
        lblDiscountTotal.Caption = 0
        txtDiscountPercent.Text = 0
        RoundDiscount = 0
        lblServiceTotal.Caption = 0
        lblDeliveryFullName.Caption = ""
        txtCarryFee.Text = 0
        txtCarryFeePercent.Text = 0
        lblCarryFeeTotal.Caption = 0
        lblPackingTotal.Caption = 0
        txtPacking.Text = 0
        txtPackingPercent.Text = 0
        lblTaxTotal.Caption = 0
        LblDutyTotal.Caption = 0
         
        mvarCustCredit = 0
        
        lblSumPrice.Caption = 0
        lblSumPrice.Tag = 0
        lblCustomer.Tag = -1
        UpdatelblCustomer
'        If lblCustomer.ListCount > 0 Then
'            lblCustomer.ListIndex = 0 ' For Customer = 0 („ ›—ﬁÂ)
'        End If
        If cmbGarson.ListCount > 0 Then
            cmbGarson.ListIndex = 0 ' For Customer = 0 („ ›—ﬁÂ)
        End If
        For i = 1 To sbrFactorProp.Panels.Count
            sbrFactorProp.Panels(i).Text = ""
        Next i
        sbrFactorProp.Panels(2).Text = fwCash.Caption
        sbrFactorProp.Panels(4).Text = FwPartition.Caption
        sbrFactorProp.Panels(6).Text = "”«· „«·Ì :" & CInt(AccountYear)
        lblPayFactorTotal.Caption = "0"
        sFactorReceived = ""
        LblRemain.Caption = ""
        LblFacpayment.BackColor = Me.BackColor
        TempAddressEdit = False
        AdminEdit = False
        chKTax.Value = False

        PersonIdqueue = 0
End Sub

Public Sub DefaultStatusbar()

End Sub
Public Sub SetFirstToolBar()

    AllButton vbOff, True

    mdifrm.Toolbar1.Buttons(1).Enabled = True   'Home
    mdifrm.Toolbar1.Buttons(2).Enabled = True   'PageUp
    mdifrm.Toolbar1.Buttons(3).Enabled = True   'PageDown
    mdifrm.Toolbar1.Buttons(4).Enabled = True   'End
    
    mdifrm.Toolbar1.Buttons(10).Enabled = False   '
    mdifrm.Toolbar1.Buttons(13).Enabled = True   'Find
    mdifrm.Toolbar1.Buttons(15).Enabled = True   'Print
    
    mdifrm.Toolbar1.Buttons(21).Enabled = True
    mdifrm.Toolbar1.Buttons(23).Enabled = True
    mdifrm.Toolbar1.Buttons(24).Enabled = True
    mdifrm.Toolbar1.Buttons(25).Enabled = True
    mdifrm.Toolbar1.Buttons(26).Enabled = True
    mdifrm.Toolbar1.Buttons(27).Enabled = True
    chKTax.Enabled = False
    
If MyFormAddEditMode = ViewMode Or MyFormAddEditMode = RefferedMode Then   ' View Mode
 
    mdifrm.Toolbar1.Buttons(6).Enabled = True  'Add
    mdifrm.Toolbar1.Buttons(7).Enabled = True  'Edit
    mdifrm.Toolbar1.Buttons(8).Enabled = False  'Enter
    mdifrm.Toolbar1.Buttons(9).Enabled = True   'Esc
    mdifrm.Toolbar1.Buttons(18).Enabled = True   'Reffer
    cmbServePlace.Locked = True
    CmbPayk.Locked = True
    cmbTable.Locked = True
ElseIf MyFormAddEditMode = AddMode Then    'Add Mode
    
    mdifrm.Toolbar1.Buttons(6).Enabled = False  'Add
    mdifrm.Toolbar1.Buttons(7).Enabled = False  'Edit
    If ClsFormAccess.SaveWithoutPrint = True Then
        mdifrm.Toolbar1.Buttons(8).Enabled = True   'Enter
    End If
    mdifrm.Toolbar1.Buttons(9).Enabled = True   'Esc
    mdifrm.Toolbar1.Buttons(18).Enabled = False   'Reffer
    cmbServePlace.Locked = False
    CmbPayk.Locked = False
    cmbTable.Locked = False
    If clsStation.ForceTax = False Then chKTax.Enabled = True Else chKTax = True

ElseIf MyFormAddEditMode = EditMode Or MyFormAddEditMode = ManipulateMode Then     'Edit
  
    mdifrm.Toolbar1.Buttons(6).Enabled = True  'Add
    mdifrm.Toolbar1.Buttons(7).Enabled = False  'Edit
    If ClsFormAccess.SaveWithoutPrint = True Then
        mdifrm.Toolbar1.Buttons(8).Enabled = True   'Enter
    End If
    mdifrm.Toolbar1.Buttons(9).Enabled = True   'Esc
    mdifrm.Toolbar1.Buttons(18).Enabled = True   'Reffer
    GetGoodAmount
    cmbServePlace.Locked = False
    CmbPayk.Locked = False
    cmbTable.Locked = False
    If clsStation.ForceTax = False Then chKTax.Enabled = True Else chKTax = True
    
End If

HeaderLabel Val(MyFormAddEditMode), fwlblMode
EnableBeforShowDifferenceFlxRow = False
End Sub

Sub OpenCashDrawer()
    
    On Error GoTo Err_Handler
        
         Select Case clsStation.DrawerModel
            
            Case 0 ' None
''''                frmDisMsg.lblMessage.Caption = "ò‘Ê ÅÊ·  ⁄—Ì› ‰‘œÂ «”  "
''''                frmDisMsg.Timer1.Interval = 1000
''''                frmDisMsg.Timer1.Enabled = True
''''                frmDisMsg.Show vbModal
    
            Case ithacaCashDrawer     'ithaca
                 Select Case clsStation.DrawerPort
                     
                    Case 0 ' Lpt Port
                        Open "Lpt1" For Output As #1
''''                        Print #1, Chr$(27) & "v0"  ' Disable Auto Cutter
''''                        Print #1, Chr$(27) & "J0"  ' Auto Line Feed 0
''''                        Print #1, Chr$(27) & "50"  ' Diable Auto Line Feed
''''                        Print #1, Chr$(27) & "v1"  ' Enable Auto Cutter
''''                        Print #1, Chr$(27) & "d4"  ' Line Feed
                            Print #1, Chr$(27) & "x1"  ' Drawer Open
''''                        Print #1, Chr$(27) & "v1"  ' Enable Auto Cutter
''''                        Print #1, Chr$(27) & "J5"  ' Enable Line Feed 5
''''                        Print #1, Chr$(27) & "51"  ' Enable Auto Line Feed
                        Close #1
                    Case 20 ' Usb Port
                            'From the command prompt type:-
                            'net use LPT2 \\pcname\printername
'                        Open "Lpt2" For Output As #1
'                            Print #1, Chr$(27) & "x1"  ' Drawer Open
'
'                        Close #1

                    Case Else     ' Serial Port
                        If mscSerial(clsStation.DrawerPort).PortOpen = False Then
                           mscSerial(clsStation.DrawerPort).PortOpen = True
                        End If
                        mscSerial(clsStation.DrawerPort).Output = Chr$(27) & "x1"
                        mscSerial(clsStation.DrawerPort).PortOpen = False
                 
                 End Select
            
            Case SamsungCashDrawer      'Samsung
                 Select Case clsStation.DrawerPort
                
                    Case 0 ' Lpt Port
                
                        Open "Lpt1" For Output As #1
                        Print #1, Chr$(27) & Chr$(112) & Chr$(0) & Chr$(10) & Chr$(10)  ' Drawer Open
                        Close #1
                     
                     Case Else     ' Serial Port
                        If mscSerial(clsStation.DrawerPort).PortOpen = False Then
                           mscSerial(clsStation.DrawerPort).PortOpen = True
                        End If
                        mscSerial(clsStation.DrawerPort).Output = Chr$(27) & Chr$(112) & Chr$(0) & Chr$(10) & Chr$(10)  ' Drawer Open
                        mscSerial(clsStation.DrawerPort).PortOpen = False
                 
                 End Select
           
            
            Case AryaCashDrawer      'Serial
                 Select Case clsStation.DrawerPort
                
                    Case 0 ' Lpt Port
                        Open "Lpt1" For Output As #1
                     '   Print #1, Chr$(0) & Chr$(0) & Chr$(0) & Chr$(0) & Chr$(0) & Chr$(0) & Chr$(0) & Chr$(0) & Chr$(0) & Chr$(0) ' Drawer Open
                        Print #1, Chr$(29) & Chr$(18) ' Drawer Open
                        Close #1
                     
                     Case Else     ' Serial Port
                        If mscSerial(clsStation.DrawerPort).PortOpen = False Then
                           mscSerial(clsStation.DrawerPort).PortOpen = True
                        End If
                        For i = 1 To 7
                           mscSerial(clsStation.DrawerPort).Output = Chr$(0) & Chr$(0) & Chr$(0) & Chr$(0) & Chr$(0) & Chr$(0) & Chr$(0) & Chr$(0) & Chr$(0) & Chr$(0) & Chr$(0) & Chr$(0) & Chr$(0) & Chr$(0) & Chr$(0) & Chr$(0) & Chr$(0) & Chr$(0) & Chr$(0) & Chr$(0) & Chr$(0) & Chr$(0)
                        Next i
                 '       mscSerial(clsStation.DrawerPort).Output = Chr$(29) & Chr$(18) ' Drawer Open
                        mscSerial(clsStation.DrawerPort).PortOpen = False
                End Select
                
            Case NCRCashDrawer      'Serial
                 Select Case clsStation.DrawerPort
                
                    Case 0 ' Lpt Port
                        Open "Lpt1" For Output As #1
                        Print #1, Chr$(27) & Chr$(112) & Chr$(0) & Chr$(100) & Chr$(100) ' Drawer Open
                        Close #1
                     
                     Case Else     ' Serial Port
                        If mscSerial(clsStation.DrawerPort).PortOpen = False Then
                           mscSerial(clsStation.DrawerPort).PortOpen = True
                        End If
''''                        mscSerial(ClsStation.DrawerPort).Output = Chr$(248) & Chr$(19) & "jhghg$" 'Test Drawer Open
                        mscSerial(clsStation.DrawerPort).Output = Chr$(27) & Chr$(112) & Chr$(0) & Chr$(100) & Chr$(100) ' Drawer Open
                        mscSerial(clsStation.DrawerPort).PortOpen = False
                End Select
                
            Case ABSCashDrawer      'Serial
                Select Case clsStation.DrawerPort
                  Case 0 ' Lpt Port
                      Open "Lpt1" For Output As #1
                      Print #1, Chr$(0) & Chr$(0) & Chr$(0) & Chr$(0) & Chr$(0) & Chr$(0) & Chr$(0) & Chr$(0) & Chr$(0) & Chr$(0) ' Drawer Open
                      Close #1
                  Case Else     ' Serial Port
                      Select Case clsStation.CashDrawerPrn
                          Case True
                              Open "Com" & clsStation.DrawerPort For Output As #1
                              Print #1, Chr$(0) & Chr$(0) & Chr$(0) & Chr$(0) & Chr$(0) & Chr$(0) & Chr$(0) & Chr$(0) & Chr$(0) & Chr$(0) ' Drawer Open
                              Close #1
                          Case False
                              If mscSerial(clsStation.DrawerPort).PortOpen = False Then
                                 mscSerial(clsStation.DrawerPort).PortOpen = True
                              End If
                ''''                        mscSerial(clsStation.DrawerPort).Output = Chr$(248) & Chr$(19) & "jhghg$" 'Test Drawer Open
                              mscSerial(clsStation.DrawerPort).Output = Chr$(0) & Chr$(0) & Chr$(0) & Chr$(0) & Chr$(0) & Chr$(0) & Chr$(0) & Chr$(0) & Chr$(0) & Chr$(0) & Chr$(0) & Chr$(0) & Chr$(0) & Chr$(0) & Chr$(0) & Chr$(0) & Chr$(0) & Chr$(0) & Chr$(0) & Chr$(0) & Chr$(0) & Chr$(0)
                              mscSerial(clsStation.DrawerPort).PortOpen = False
                        End Select
                End Select
            Case ADPCashDrawer       'Lpt
                Select Case clsStation.DrawerPort
                  Case 0 ' Lpt Port
                      Open "Lpt1" For Output As #1
                      Print #1, Chr$(27) & Chr$(112) & Chr$(48) & Chr$(50) & Chr$(50) '& Chr$(50) ' Drawer Open
                      Close #1
                  Case Else     ' Serial Port
                      Select Case clsStation.CashDrawerPrn
                          Case True
                              Open "Com" & clsStation.DrawerPort For Output As #1
                              Print #1, Chr$(27) & Chr$(112) & Chr$(48) & Chr$(50) & Chr$(50) ' Drawer Open
                              Close #1
                          Case False
                              If mscSerial(clsStation.DrawerPort).PortOpen = False Then
                                 mscSerial(clsStation.DrawerPort).PortOpen = True
                              End If
                ''''                        mscSerial(clsStation.DrawerPort).Output = Chr$(248) & Chr$(19) & "jhghg$" 'Test Drawer Open
                              mscSerial(clsStation.DrawerPort).Output = Chr$(27) & Chr$(112) & Chr$(48) & Chr$(50) & Chr$(50)
                              mscSerial(clsStation.DrawerPort).PortOpen = False
                        End Select
                End Select
            Case EpsonCashDrawer       'Lpt
                Select Case clsStation.DrawerPort
                  Case 0 ' Lpt Port
                      Open "Lpt1" For Output As #1
                      Print #1, Chr$(27) & "p" & Chr$(0)  ' Drawer Open
                      Close #1
                  Case Else     ' Serial Port
                      Select Case clsStation.CashDrawerPrn
                          Case True
                              Open "Com" & clsStation.DrawerPort For Output As #1
                              Print #1, Chr$(27) & "p" & Chr$(0) ' Drawer Open
                              Close #1
                          Case False
                              If mscSerial(clsStation.DrawerPort).PortOpen = False Then
                                 mscSerial(clsStation.DrawerPort).PortOpen = True
                              End If
                ''''                        mscSerial(clsStation.DrawerPort).Output = Chr$(248) & Chr$(19) & "jhghg$" 'Test Drawer Open
                              mscSerial(clsStation.DrawerPort).Output = Chr$(27) & "p" & Chr$(0)
                              mscSerial(clsStation.DrawerPort).PortOpen = False
                        End Select
                End Select
            Case DigiCashDrawer       'Lpt
                
                Select Case clsStation.DrawerPort
                  Case 0 ' Lpt Port
                      Open "Lpt1" For Output As #1
                      Print #1, Chr(&H1B) & Chr(&H3F) & Chr(&H29) & Chr(&HA)
                      Close #1
                  Case Else     ' Serial Port
                      Select Case clsStation.CashDrawerPrn
                          Case True
                              Open "Com" & clsStation.DrawerPort For Output As #1
                              Print #1, Chr(&H1B) & Chr(&H3F) & Chr(&H29) & Chr(&HA)
                              Close #1
                          Case False
                              If mscSerial(clsStation.DrawerPort).PortOpen = False Then
                                 mscSerial(clsStation.DrawerPort).PortOpen = True
                              End If
                              mscSerial(clsStation.DrawerPort).Output = Chr(&H1B) & Chr(&H3F) & Chr(&H29) & Chr(&HA)
                              mscSerial(clsStation.DrawerPort).PortOpen = False
                        End Select
                End Select
                
            Case PartnerCashDrawer      'Internal
                
                Dim Data, Data2 As Integer
                Enter_Config
                SelectLD7
                MutilpinSelGPIO
                DefineInOut
                
                Data = ReadData(GPIOInOutDataReg)
                Data2 = Data
                
                Data = Data Xor Cash1Out
                SendData GPIOInOutDataReg, Data
                'MsgBox "Now should open cash A"
                SendData GPIOInOutDataReg, Data2
                BackToDefault
                
                
         End Select
Exit Sub
    
Err_Handler:

MsgBox err.Number & " " & err.Description

End Sub

Public Sub CustomerDisplay(sumPrice As Currency, StringLine1 As String, Optional amount As Double)
 
    On Error GoTo Err_Handler
    
    Dim strTemp As String
    
    
    If clsStation.CustomerAscii = True Then
        StringLine1 = WinToIranSys(StringLine1)
        strTemp = WinToIranSys("—Ì«·") & "  " & sumPrice
    ElseIf clsStation.CustomerFarsi = True Then
        strTemp = "—Ì«·" & "  " & sumPrice
    Else
       ' StringLine1 = clsArya.CustomerDisplayName
        strTemp = sumPrice & "  " & "Rls"
    End If
   
    Select Case clsStation.CustDisplayModel    '

        Case GigaCustomerDisplay
            
            Select Case clsStation.CustDisplayPrn
                 Case True
                    Open "Com" & clsStation.CustDisplayPort For Output As #1
                    Print #1, Chr(27) & Chr(83) ' Disable AUX
                    Print #1, Chr(4) & Chr(1) & Chr(80) & Chr(49) & Chr(23)
                    Print #1, Space((20 - Len(StringLine1)) / 2) & StringLine1 & Space((20 - Len(StringLine1)) / 2)
        
                    Print #1, Chr(4) & Chr(1) & Chr(80) & Chr(69) & Chr(23)
                    Print #1, Space((20 - Len(strTemp)) / 2) & strTemp & Space((20 - Len(strTemp)) / 2)
        
                    Print #1, Chr(27) & Chr(71) ' enable AUX Device
        
                    Close #1
                
                Case False
                
                    If mscSerial(clsStation.CustDisplayPort).PortOpen = False Then
                       mscSerial(clsStation.CustDisplayPort).PortOpen = True
                    End If
                    mscSerial(clsStation.CustDisplayPort).Output = Chr(27) & Chr(83) ' Disable AUX Device
        
                    mscSerial(clsStation.CustDisplayPort).Output = Chr(4) & Chr(1) & Chr(80) & Chr(49) & Chr(23)
                    mscSerial(clsStation.CustDisplayPort).Output = Space((20 - Len(StringLine1)) / 2) & StringLine1 & Space((20 - Len(StringLine1)) / 2)
        
                    mscSerial(clsStation.CustDisplayPort).Output = Chr(4) & Chr(1) & Chr(80) & Chr(69) & Chr(23)
                    mscSerial(clsStation.CustDisplayPort).Output = Space((20 - Len(strTemp)) / 2) & strTemp & Space((20 - Len(strTemp)) / 2)
        
                    mscSerial(clsStation.CustDisplayPort).Output = Chr(27) & Chr(71) ' enable AUX Device
        
                    mscSerial(clsStation.CustDisplayPort).PortOpen = False

                End Select
        Case AryaCustomerDisplay  'arya
            Select Case clsStation.CustDisplayPrn
                 Case True
                    Open "Com" & clsStation.CustDisplayPort For Output As #1
                    Print #1, Chr(31) & Chr(64) ' Initialize
                    Print #1, Space((20 - Len(StringLine1)) / 2) & StringLine1 & Space((20 - Len(StringLine1)) / 2)
                    Print #1, Space((20 - Len(strTemp)) / 2) & strTemp & Space((20 - Len(strTemp)) / 2)
                    Close #1
                
                Case False
                
                    If mscSerial(clsStation.CustDisplayPort).PortOpen = False Then
                       mscSerial(clsStation.CustDisplayPort).PortOpen = True
                    End If
                    mscSerial(clsStation.CustDisplayPort).Output = Chr(31) & Chr(64) 'Initialize
                    mscSerial(clsStation.CustDisplayPort).Output = Space((20 - Len(StringLine1)) / 2) & StringLine1 & Space((20 - Len(StringLine1)) / 2)
                    mscSerial(clsStation.CustDisplayPort).Output = Space((20 - Len(strTemp)) / 2) & strTemp & Space((20 - Len(strTemp)) / 2)
                    mscSerial(clsStation.CustDisplayPort).PortOpen = False

                End Select
        Case NcrCustomerDisplay
            Select Case clsStation.CustDisplayPrn
                 Case True
                    Open "Com" & clsStation.CustDisplayPort For Output As #1
                    Print #1, Chr(27) & Chr(12) ' Initialize
                    Print #1, Space((20 - Len(StringLine1)) / 2) & StringLine1 & Space((20 - Len(StringLine1)) / 2)
                    Print #1, Space((20 - Len(strTemp)) / 2) & strTemp & Space((20 - Len(strTemp)) / 2)
                    Close #1
                
                Case False
                
                    If mscSerial(clsStation.CustDisplayPort).PortOpen = False Then
                       mscSerial(clsStation.CustDisplayPort).PortOpen = True
                    End If
                    mscSerial(clsStation.CustDisplayPort).Output = Chr(27) & Chr(12) 'Initialize
                    mscSerial(clsStation.CustDisplayPort).Output = Space((20 - Len(StringLine1)) / 2) & StringLine1 & Space((20 - Len(StringLine1)) / 2)
                    mscSerial(clsStation.CustDisplayPort).Output = Space((20 - Len(strTemp)) / 2) & strTemp & Space((20 - Len(strTemp)) / 2)
                    mscSerial(clsStation.CustDisplayPort).PortOpen = False

                End Select
        
        Case ABSCustomerDisplay
            Select Case clsStation.CustDisplayPrn
                 Case True
                    Open "Com" & clsStation.CustDisplayPort For Output As #1
                    Print #1, Chr$(27) & Chr$(73) ' Initialize
                    Print #1, Space((20 - Len(StringLine1)) / 2) & StringLine1 & Space((20 - Len(StringLine1)) / 2)
                    Print #1, Space((20 - Len(strTemp)) / 2) & strTemp & Space((20 - Len(strTemp)) / 2)
                    Close #1
                
                Case False
                
                    If mscSerial(clsStation.CustDisplayPort).PortOpen = False Then
                       mscSerial(clsStation.CustDisplayPort).PortOpen = True
                    End If
                    mscSerial(clsStation.CustDisplayPort).Output = Chr$(27) & Chr$(73)   ' Initialize
                    mscSerial(clsStation.CustDisplayPort).Output = Space((20 - Len(StringLine1)) / 2) & StringLine1 & Space((20 - Len(StringLine1)) / 2)
                    mscSerial(clsStation.CustDisplayPort).Output = Space((20 - Len(strTemp)) / 2) & strTemp & Space((20 - Len(strTemp)) / 2)
                    mscSerial(clsStation.CustDisplayPort).PortOpen = False

            End Select
    
        Case DigiCustomerDisplay
            Select Case clsStation.CustDisplayPrn
                 Case True
                    Open "Com" & clsStation.CustDisplayPort For Output As #1
                    Print #1, Chr(27) & Chr(37) & Chr(43) & Chr(10) 'Clear Screen
                    Print #1, Chr(27) & Chr(35) & Chr(91) & Chr(0) & Chr(0) & Space((20 - Len(StringLine1)) / 2) & StringLine1 & Chr(10)
                    Print #1, Chr(27) & Chr(35) & Chr(91) & Chr(2) & Chr(32) & Space((20 - Len(strTemp)) / 2) & strTemp & Chr(10)
                    Close #1
                
                Case False
                
                    If mscSerial(clsStation.CustDisplayPort).PortOpen = False Then
                       mscSerial(clsStation.CustDisplayPort).PortOpen = True
                    End If
                    mscSerial(clsStation.CustDisplayPort).Output = Chr(27) & Chr(37) & Chr(43) & Chr(10) 'Clear Screen
                    mscSerial(clsStation.CustDisplayPort).Output = Chr(27) & Chr(35) & Chr(91) & Chr(0) & Chr(0) & Space((20 - Len(StringLine1)) / 2) & StringLine1 & Chr(10)
                    mscSerial(clsStation.CustDisplayPort).Output = Chr(27) & Chr(35) & Chr(91) & Chr(2) & Chr(32) & Space((20 - Len(strTemp)) / 2) & strTemp & Chr(10)
                    mscSerial(clsStation.CustDisplayPort).PortOpen = False

            End Select
         Case EpsonCustomerDisplay
            Select Case clsStation.CustDisplayPrn
                 Case True
                    Open "Com" & clsStation.CustDisplayPort For Output As #1
                    Print #1, Chr$(27) & Chr$(64) ' Initialize
                    Print #1, Space((20 - Len(StringLine1)) / 2) & StringLine1 & Space((20 - Len(StringLine1)) / 2)
                    Print #1, Space((20 - Len(strTemp)) / 2) & strTemp & Space((20 - Len(strTemp)) / 2)
                    Close #1
                
                Case False
                
                    If mscSerial(clsStation.CustDisplayPort).PortOpen = False Then
                       mscSerial(clsStation.CustDisplayPort).PortOpen = True
                    End If
                    mscSerial(clsStation.CustDisplayPort).Output = Chr$(27) & Chr$(64)   ' Initialize
                    mscSerial(clsStation.CustDisplayPort).Output = Space((20 - Len(StringLine1)) / 2) & StringLine1 & Space((20 - Len(StringLine1)) / 2)
                    mscSerial(clsStation.CustDisplayPort).Output = Space((20 - Len(strTemp)) / 2) & strTemp & Space((20 - Len(strTemp)) / 2)
                    mscSerial(clsStation.CustDisplayPort).PortOpen = False

            End Select
    
          Case ZonrichCustomerDisplay
            Select Case clsStation.CustDisplayPrn
                Case True
                    Open "Com" & clsStation.CustDisplayPort For Output As #1
                    Print #1, Chr(13) & Chr(64) ' Initialize
                    Print #1, Space((20 - Len(StringLine1)) / 2) & StringLine1 & Space((20 - Len(StringLine1)) / 2)
                    Print #1, Space((20 - Len(strTemp)) / 2) & strTemp & Space((20 - Len(strTemp)) / 2)
                    Close #1
                    
                Case False
                
                    If mscSerial(clsStation.CustDisplayPort).PortOpen = False Then
                       mscSerial(clsStation.CustDisplayPort).PortOpen = True
                    End If
                        mscSerial(clsStation.CustDisplayPort).Output = Chr(13) & Chr(64)
                        mscSerial(clsStation.CustDisplayPort).Output = Space((20 - Len(StringLine1)) / 2) & StringLine1 & Space((20 - Len(StringLine1)) / 2)
                        mscSerial(clsStation.CustDisplayPort).Output = Space((20 - Len(strTemp)) / 2) & strTemp & Space((20 - Len(strTemp)) / 2)
                        mscSerial(clsStation.CustDisplayPort).PortOpen = False
                        
            End Select
    
        Case ZonrichCustomerDisplay_ZQ
          Select Case clsStation.CustDisplayPrn
            Case True
                 Open "Com" & clsStation.CustDisplayPort For Output As #1
                 Print #1, strTemp
                 Print #1, Chr(13) & Chr(64) ' Initialize
                 Close #1
            Case False
                 If mscSerial(clsStation.CustDisplayPort).PortOpen = False Then
                       mscSerial(clsStation.CustDisplayPort).PortOpen = True
                 End If
                 
                 mscSerial(clsStation.CustDisplayPort).Output = strTemp
                 mscSerial(clsStation.CustDisplayPort).Output = Chr(13) & Chr(64)
            End Select
        
        Case StandardEposCustomerDisplay
          Select Case clsStation.CustDisplayPrn
            Case True
                 Open "Com" & clsStation.CustDisplayPort For Output As #1
                 Print #1, Chr(12)
                 Print #1, Chr(27) & "QA" & Space((20 - Len(StringLine1)) / 2) & StringLine1 & Chr(13)
                 Print #1, Chr(27) & "QB" & Space((20 - Len(strTemp)) / 2) & strTemp & Chr(13)
                 Close #1
            Case False
                 If mscSerial(clsStation.CustDisplayPort).PortOpen = False Then
                       mscSerial(clsStation.CustDisplayPort).PortOpen = True
                 End If
                 
                mscSerial(clsStation.CustDisplayPort).Output = Chr(12) 'CLR clear display
    '            mscSerial(clsStation.CustDisplayPort).Output = Chr(CAN) 'CLR clear display
    '            mscSerial(clsStation.CustDisplayPort).Output = Chr(27) & "QD" & strTemp & Chr(CR)  'ESC Q D.......CR msg to scroll in upper line
                mscSerial(clsStation.CustDisplayPort).Output = Chr(27) & "QA" & Space((20 - Len(StringLine1)) / 2) & StringLine1 & Chr(13)  'ESC Q A.......CR msg to upper line
                mscSerial(clsStation.CustDisplayPort).Output = Chr(27) & "QB" & Space((20 - Len(strTemp)) / 2) & strTemp & Chr(13)  'ESC Q B.......CR msg to lower line
            
            End Select
    Case HisenseCustomerDisplay
          strTemp = Val(strTemp)
          Select Case clsStation.CustDisplayPrn
            Case True
                 Open "Com" & clsStation.CustDisplayPort For Output As #1
                 Print #1, Chr(12)
                 Print #1, Chr(27) & Chr(6) & Chr(1)
                 Print #1, Chr(27) & Chr(81) & Chr(64) & Space((15 - Len(StringLine1)) / 2) & StringLine1 & Chr(13)
                 Print #1, Chr(27) & Chr(81) & Chr(66) & "Total:" & strTemp & IIf(Len(strTemp) >= 8, "", Space((9 - Len(strTemp)) / 2)) & "R" & Chr(13)    'ESC Q B.......CR msg to lower line
                 Close #1
            Case False
                 If mscSerial(clsStation.CustDisplayPort).PortOpen = False Then
                       mscSerial(clsStation.CustDisplayPort).PortOpen = True
                 End If

                mscSerial(clsStation.CustDisplayPort).Output = Chr(12) 'CLR clear display
                mscSerial(clsStation.CustDisplayPort).Output = Chr(27) & Chr(6) & Chr(1) 'Font Size
                mscSerial(clsStation.CustDisplayPort).Output = Chr(27) & Chr(81) & Chr(64) & Space((15 - Len(StringLine1)) / 2) & StringLine1 & Chr(13)  'ESC Q A.......CR msg to upper line
                mscSerial(clsStation.CustDisplayPort).Output = Chr(27) & Chr(81) & Chr(66) & "Total:" & strTemp & IIf(Len(strTemp) >= 8, "", Space((9 - Len(strTemp)) / 2)) & "R" & Chr(13)     'ESC Q B.......CR msg to lower line

        End Select
    Case HisensePersianCustomerDisplay
          strTemp = Val(strTemp)
          Select Case clsStation.CustDisplayPrn
            Case True
                 Open "Com" & clsStation.CustDisplayPort For Output As #1
                 Print #1, Chr(12)
                 Print #1, Chr(27) & Chr(6) & Chr(1)
                 Print #1, Chr(27) & Chr(81) & Chr(64) & Space((15 - Len(StringLine1)) / 2) & StringLine1 & Chr(13)
                 Print #1, Chr(27) & Chr(81) & Chr(66) & "Total:" & strTemp & IIf(Len(strTemp) >= 8, "", Space((9 - Len(strTemp)) / 2)) & "R" & Chr(13)    'ESC Q B.......CR msg to lower line
                 Close #1
            Case False
                 If mscSerial(clsStation.CustDisplayPort).PortOpen = False Then
                       mscSerial(clsStation.CustDisplayPort).PortOpen = True
                 End If
                 
                mscSerial(clsStation.CustDisplayPort).Output = Chr(27) & Chr(5)
                mscSerial(clsStation.CustDisplayPort).Output = Chr(12) 'CLR clear display
                If amount <> 0 Then
                    mscSerial(clsStation.CustDisplayPort).Output = Chr(27) & Chr(98) & Chr(64) & GetFarsiHisenseCustomerDisplay(StrReverse(CStr(amount))) & GetFarsiHisenseCustomerDisplay(mvarUnitDescription) & Chr(13)               'ESC Q A.......CR msg to upper line
                    mscSerial(clsStation.CustDisplayPort).Output = Chr(27) & Chr(97) & Chr(64) & GetFarsiHisenseCustomerDisplay(CStr(amount * mvarSellPrice)) & Chr(13)              'ESC Q A.......CR msg to upper line
                    'mscSerial(clsStation.CustDisplayPort).Output = Chr(27) & Chr(6) & Chr(1) 'Font Size
                    mscSerial(clsStation.CustDisplayPort).Output = Chr(27) & Chr(98) & Chr(65) & GetFarsiHisenseCustomerDisplay(mvarGoodName) & Chr(13)             'ESC Q A.......CR msg to upper line

                     mscSerial(clsStation.CustDisplayPort).Output = Chr(27) & Chr(97) & Chr(67) & Chr(150) & Chr(73) & Chr(169) & Chr(102) & GetFarsiHisenseCustomerDisplay(strTemp) & Chr(13)     '& GetFarsiHisenseCustomerDisplay(StrReverse("Ã„⁄")) & Chr(13)     'ESC Q B.......CR msg to lower line
                     mscSerial(clsStation.CustDisplayPort).Output = Chr(27) & Chr(98) & Chr(67) & GetFarsiHisenseCustomerDisplay("Ã„⁄") & Chr(13)     'ESC Q B.......CR msg to lower line
                Else
                
                    mscSerial(clsStation.CustDisplayPort).Output = Chr(12) 'CLR clear display
                    
                    mscSerial(clsStation.CustDisplayPort).Output = Chr(27) & Chr(6) & Chr(0) 'Font Size
                    mscSerial(clsStation.CustDisplayPort).Output = Chr(27) & Chr(3) & Chr(0) & Chr(47)   '& Chr(i) & Chr(10) & StrReverse(GetFarsiHisenseCustomerDisplay(GetFarsiStringFromArabic(clsArya.CustomerDisplayName))) & Chr(13) '& Space((15 - Len(StringLine1)) / 2) & StringLine1 & Chr(13)          'ESC Q A.......CR msg to upper line
                    'mscSerial(clsStation.CustDisplayPort).Output = Chr(27) & Chr(3) & Chr(0) & Chr(1)
                    mscSerial(clsStation.CustDisplayPort).Output = Chr(27) & Chr(4) & Chr(150)
                    mscSerial(clsStation.CustDisplayPort).Output = Chr(27) & Chr(98) & Chr(65) & GetFarsiHisenseCustomerDisplay(GetFarsiStringFromArabic(clsArya.CustomerDisplayName)) & Chr(13)  '& Space((15 - Len(StringLine1)) / 2) & StringLine1 & Chr(13)          'ESC Q A.......CR msg to upper line
                    mscSerial(clsStation.CustDisplayPort).Output = Chr(27) & Chr(97) & Chr(67) & Chr(150) & Chr(73) & Chr(169) & Chr(102) & GetFarsiHisenseCustomerDisplay(strTemp) & Chr(13)     '& GetFarsiHisenseCustomerDisplay(StrReverse("Ã„⁄")) & Chr(13)     'ESC Q B.......CR msg to lower line
                    mscSerial(clsStation.CustDisplayPort).Output = Chr(27) & Chr(98) & Chr(67) & GetFarsiHisenseCustomerDisplay("Ã„⁄") & Chr(13)     'ESC Q B.......CR msg to lower line

                End If
        End Select
    End Select

    Exit Sub
    
Err_Handler:

MsgBox err.Number & " " & err.Description

End Sub


Sub FixPrice()

    
''''    Dim discount_new As Integer
''''    Dim discount_new2 As Integer
''''
''''    discount_new = Me.lblSumPrice.Caption Mod 1000
''''    discount_new2 = Me.lblSumPrice.Caption Mod 100
''''
''''    If (discount_new <= clsStation.MaxAutoDiscount) Then
''''
''''        Me.lblDiscountTotal.Caption = CLng(Val(Me.lblDiscountTotal.Caption) + discount_new)
''''        Me.lblSumPrice.Caption = CLng(Val(Me.lblSumPrice.Caption) - discount_new)
''''
''''    Else
''''        If clsStation.RoundTwoNumber = True Then
''''            If discount_new2 = 0 Then Exit Sub
''''            If discount_new2 > 0 And discount_new2 < 26 Then
''''            ElseIf discount_new2 > 26 And discount_new2 < 76 Then
''''               discount_new2 = discount_new2 - 50
''''            ElseIf discount_new2 > 76 And discount_new2 < 100 Then
''''               discount_new2 = discount_new2 - 100
''''            End If
''''        End If
''''        Me.lblDiscountTotal = CLng(Val(Me.lblDiscountTotal.Caption) + discount_new2)
''''        Me.lblSumPrice.Caption = CLng(Val(Me.lblSumPrice.Caption) - discount_new2)
''''
''''
''''    End If
''''
    
End Sub

Public Sub CalculateDelivery()

    Dim intDelivered As Integer
    Dim intNotDelivered As Integer
    
    lblDailyDelivery.Caption = ""
    lblDailyDelivered.Caption = ""
    
    ReDim Parameter(1) As Parameter
    
    Parameter(0) = GenerateInputParameter("@Date", adVarWChar, 8, mvarDate)
    Parameter(1) = GenerateOutputParameter("@Count", adInteger, 4)
    intDelivered = RunParametricStoredProcedure("Get_DailyCountDeliveredFactors", Parameter)
    
    Parameter(0) = GenerateInputParameter("@Date", adVarWChar, 8, mvarDate)
    Parameter(1) = GenerateOutputParameter("@Count", adInteger, 4)
    intNotDelivered = RunParametricStoredProcedure("Get_DailyCountNotDeliveredFactors", Parameter)
    If (intDelivered + intNotDelivered) = 0 Then
       lblDailyDelivery.Visible = False
    Else
       lblDailyDelivery.Visible = True
    End If
    If intNotDelivered = 0 Then
       lblDailyDelivered.Visible = False
    Else
       lblDailyDelivered.Visible = True
    End If
    lblDailyDelivery.Caption = " ⁄œ«œ ò· «—”«·Ì Â« : " & (intDelivered + intNotDelivered)
    lblDailyDelivered.Caption = " ⁄œ«œ «—”«· ‰‘œÂ Â« : " & intNotDelivered
End Sub
Public Sub CalculateTemporary()

    Dim intTempCount As Integer
    
    lblTemporary.Caption = ""
    
    ReDim Parameter(1) As Parameter
    
    Parameter(0) = GenerateInputParameter("@Branch", adInteger, 4, CurrentBranch)
    Parameter(1) = GenerateOutputParameter("@Count", adInteger, 4)
    intTempCount = RunParametricStoredProcedure("Get_TempFactors_Count", Parameter)
    
    If intTempCount = 0 Then
       lblTemporary.Visible = False
    Else
       lblTemporary.Visible = True
    End If
    lblTemporary.Caption = " ⁄œ«œ ò· „Êﬁ  Â« : " & intTempCount
End Sub

Private Sub LoadTempData(No As String)

    ClearDataFlexGrid
    
    ReDim Parameter(3) As Parameter
    
    Parameter(0) = GenerateInputParameter("@No", adBigInt, 8, Val(No))
    Parameter(1) = GenerateInputParameter("@Status", adInteger, 4, mvarStatus)
    Parameter(2) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
    Parameter(3) = GenerateInputParameter("@Branch", adInteger, 4, CurrentBranch)
    
    Set rctmp = RunParametricStoredProcedure2Rec("Get_FacMD_Good_Temp", Parameter)
    
    Dim ii As Integer
    
    If Not (rctmp.BOF Or rctmp.EOF) Then
    
        TxtTempAddress.Text = IIf(IsNull(rctmp!TempAddress), "", rctmp!TempAddress)
        If Len(Trim(TxtTempAddress)) > 0 Then TempAddressEdit = True
        TxtGuestNo.Text = rctmp!GuestNo
        boolPayment = rctmp!FacPayment
        BalancePayment = False
        cmbTable.ListIndex = 0
        cmbGarson.ListIndex = 0
        txtCarryFee.Text = rctmp!CarryFeeTotal
        ''Because Tax need to calculate and we don't have it
        'ServiceRate = rctmp!ServiceTotal * 100 / (rctmp!sumPrice - rctmp!ServiceTotal - rctmp!CarryFeeTotal - rctmp!PackingTotal + rctmp!DiscountTotal)
        txtPacking.Text = rctmp!PackingTotal
        mVarOrderType = rctmp!OrderType
        mvarServePlace = rctmp!ServePlace
        For i = 0 To cmbServePlace.ListCount - 1
            If mvarServePlace = cmbServePlace.ItemData(i) Then
                cmbServePlace.ListIndex = i
                Exit For
            End If
        Next i

        intSumOfCurrentServePlaces = mvarServePlace
        If mvarServePlace = EnumServePlace.Table Or mvarServePlace = EnumServePlace.Salon Then
            ServiceRate = DefaultServicePercent
        Else
            ServiceRate = 0
        End If
        txtDescription.Text = IIf(IsNull(rctmp!NvcDescription), " ", rctmp!NvcDescription)
        If Trim(txtDescription.Text) <> "" Then textDescriptionFlag = True
'        InventoryNo = rctmp!IntInventoryNo
        
        lblCustomer.Tag = rctmp!Customer
        UpdatelblCustomer
        
        For i = 0 To cmbTable.ListCount - 1
            If rctmp.Fields("TableNo").Value = cmbTable.ItemData(i) Then
                cmbTable.ListIndex = i
            End If
        Next i
        
        For i = 0 To cmbGarson.ListCount - 1
            If rctmp.Fields("incharge").Value = cmbGarson.ItemData(i) Then
                cmbGarson.ListIndex = i
            End If
        Next i
        TmpGoodDiscount = 0
        Do While Not (rctmp.EOF)
            ii = ii + 1
            FlxDetail.TextMatrix(ii, 0) = rctmp!intRow 'Number
            FlxDetail.TextMatrix(ii, 1) = rctmp!amount
            FlxDetail.TextMatrix(ii, 2) = rctmp!nvcName 'GoodName
            FlxDetail.TextMatrix(ii, 3) = rctmp!FeeUnit
            FlxDetail.TextMatrix(ii, 4) = rctmp!amount * rctmp!FeeUnit ' rctmp!FeeTotal
            FlxDetail.TextMatrix(ii, 5) = rctmp!GoodCode
            FlxDetail.TextMatrix(ii, 6) = rctmp!Weight ' rctmp!WeightUnit
            FlxDetail.TextMatrix(ii, 7) = rctmp!Unit
            FlxDetail.TextMatrix(ii, 8) = rctmp!ServePlace
            FlxDetail.TextMatrix(ii, 9) = IIf(IsNull(rctmp!DifferencesCodes), "", rctmp!DifferencesCodes)
            FlxDetail.TextMatrix(ii, 10) = IIf(IsNull(rctmp!DifferencesDescription), "", rctmp!DifferencesDescription)
            FlxDetail.TextMatrix(ii, 11) = rctmp!Discount
            FlxDetail.TextMatrix(ii, 12) = rctmp!Rate
            FlxDetail.TextMatrix(ii, 13) = IIf(IsNull(rctmp!ChairName), "", rctmp!ChairName)
            FlxDetail.TextMatrix(ii, 14) = rctmp!intInventoryNo
            FlxDetail.TextMatrix(ii, 15) = rctmp!mainType
            FlxDetail.TextMatrix(ii, 17) = rctmp!DutySale
            FlxDetail.TextMatrix(ii, 18) = rctmp!TaxSale
            
            TmpGoodDiscount = TmpGoodDiscount + (rctmp!Discount * rctmp!amount * rctmp!FeeUnit / 100)
            txtDiscount.Text = rctmp!DiscountTotal - TmpGoodDiscount
            
            rctmp.MoveNext
            
            
            MaxRowFlexGrid = ii + 1
            If ii >= FlxDetail.Rows - 2 And rctmp.EOF = False Then
                AddEmptyRow
            End If

        Loop
        
        FlxDetail.Row = MaxRowFlexGrid - 1
        'mvarServePlace = FlxDetail.TextMatrix(MaxRowFlexGrid - 1, 8)
    End If
    
    rctmp.Close
    If mVarOrderType = ByPhone Then
      If clsStation.Language = Farsi Then
         LblOrder.Caption = " ·›‰Ì"
       Else
         LblOrder.Caption = "By phone"
       End If
    Else
       If clsStation.Language = Farsi Then
            LblOrder.Caption = "Õ÷Ê—Ì"
       Else
            LblOrder.Caption = "Inside"
       End If
    End If
    RefreshLables
    
    sbrFactorProp.Panels(1).Text = ""
    sbrFactorProp.Panels(2).Text = ""
    sbrFactorProp.Panels(3).Text = ""
    sbrFactorProp.Panels(4).Text = ""
    sbrFactorProp.Panels(5).Text = ""
    
  
End Sub


Private Sub TxtTempAddress_Change()
    BtnKeypad(11).Enabled = True     '"%"
    BtnKeypad(10).Enabled = True      '"."
    BtnKalaDelete.Enabled = True
    lblNum.Caption = ""
    lblBarCode.Caption = ""
End Sub

Private Sub TxtTempAddress_GotFocus()
 textTempAddressFlag = True
 If MyFormAddEditMode = AddMode And TempAddressEdit = False Then TxtTempAddress = ""
End Sub

Private Sub TxtTempAddress_LostFocus()
    textTempAddressFlag = False
End Sub

Private Sub USBHID1_CardDataChanged()
With USBHID1
    If Not (.GetTrack(1) <> "" And .GetTrack(2) <> "" And .GetTrack(3) <> "") Then
        CreditCode = Val(.FindElement(2, ";", 0, "?")) ' & vbCrLf
     '   Text1.Text = Text1.Text & .FindElement(1, "^", 0, "/") & vbCrLf
     '   Text1.Text = Text1.Text & .FindElement(1, "^", 0, "^^") & vbCrLf
     '   Text1.Text = Text1.Text & .GetTrack(1) & vbCrLf
     '   Text1.Text = Text1.Text & .GetTrack(2) & vbCrLf
     '   Text1.Text = Text1.Text & .GetTrack(3) & vbCrLf
     '   Text1.Text = Text1.Text & .GetFName & vbCrLf
     '   Text1.Text = Text1.Text & .GetLName & vbCrLf
    End If
    USBHID1.clearbuffer
    FindCust
End With
End Sub
Private Sub GetGoodAmount()
'    ReDim GoodAmount(MaxRowFlexGrid, 1) As String
Dim i, j As Long
    For i = 1 To 30
        GoodAmount(i, 0) = ""
        GoodAmount(i, 1) = ""
    Next i
    Dim RepeatGood As Boolean
    With FlxDetail
        For i = 1 To MaxRowFlexGrid - 1
            RepeatGood = False
            For j = 1 To 30
                If GoodAmount(j, 0) = .TextMatrix(i, 5) Then
                    RepeatGood = True
                    Exit For
                End If
            Next j
            If RepeatGood = False Then
                GoodAmount(i, 0) = .TextMatrix(i, 5)
                GoodAmount(i, 1) = .TextMatrix(i, 1)
            Else
                GoodAmount(j, 1) = GoodAmount(j, 1) + Val(.TextMatrix(i, 1))
            End If
        Next i
    End With
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call PresetScreenSaver
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call PresetScreenSaver
End Sub

Private Function CheekGoodAmount() As Boolean
    
'    If CBool(clsArya.EnableUpperAmountGood) = False Then
'        CheekGoodAmount = True
'        Exit Function
    If ClsFormAccess.UpperAmountGood = True Then
        CheekGoodAmount = True
        Exit Function
    End If
    
    Dim TempGoodAmount(30, 1) As String
    Dim i, j As Integer
    CheekGoodAmount = True
    If MyFormAddEditMode = AddMode Then
        CheekGoodAmount = True
        Exit Function
    End If
    Dim RepeatGood As Boolean
    With FlxDetail
        For i = 1 To MaxRowFlexGrid - 1
            RepeatGood = False
            For j = 1 To 30
                If TempGoodAmount(j, 0) = .TextMatrix(i, 5) Then
                    RepeatGood = True
                    Exit For
                End If
            Next j
            If RepeatGood = False Then
                TempGoodAmount(i, 0) = .TextMatrix(i, 5)
                TempGoodAmount(i, 1) = .TextMatrix(i, 1)
            Else
                TempGoodAmount(j, 1) = TempGoodAmount(j, 1) + Val(.TextMatrix(i, 1))
            End If
        Next i
    End With
    For i = 1 To 30
        If Val(GoodAmount(i, 0)) > 0 Then
            CheekGoodAmount = False
            For j = 1 To 30
                If Val(GoodAmount(i, 0)) = Val(TempGoodAmount(j, 0)) And Val(GoodAmount(i, 1)) <= Val(TempGoodAmount(j, 1)) Then
                    CheekGoodAmount = True
                    Exit For
                End If
            Next j
        End If
        If CheekGoodAmount = False Then Exit Function
    Next i
End Function
Private Sub wbsrPrint_NavigateComplete2(ByVal pDisp As Object, URL As Variant)
''''    If clsArya.PrintHtml = True And IsPrinting = True Then
''''        wbsrPrint.ExecWB OLECMDID_PRINT, OLECMDEXECOPT_DONTPROMPTUSER
''''        IsPrinting = False
''''    End If
End Sub
Private Function EditForTime() As Boolean
    On Error GoTo ErrHandler

    EditForTime = False
    If rctmp.State = 1 Then If rctmp.State = adStateOpen Then rctmp.Close
    
    ReDim Parameter(1) As Parameter
    Parameter(0) = GenerateInputParameter("@intserialNo", adInteger, 4, intSerialNo)
    Parameter(1) = GenerateInputParameter("@Branch", adInteger, 4, CurrentBranch)
    Set rctmp = RunParametricStoredProcedure2Rec("Get_CurrentEditTime", Parameter)
    If Not (rctmp.EOF = True And rctmp.BOF = True) Then
        If ClsFormAccess.EditAllStationsFactors = True Or ClsFormAccess.EditAllFichUser = True Then
            If rctmp!ManagerDiffTme = 1 Then
                EditForTime = True
            Else
                ShowDisMessage "“„«‰ ÊÌ—«Ì‘ ‰Â«ÌÌ »—«Ì «Ì‰ ›Ì‘ ”Å—Ì ‘œÂ..", 2000
            End If
        ElseIf ClsFormAccess.EditInvoice = True Then
            If rctmp!UserDiffTme = 1 Then
                EditForTime = True
            ElseIf rctmp!ManagerDiffTme = 1 Then
                frmAccess.MyFormAddEditMode = EditMode
                frmAccess.lblTitle.Caption = "“„«‰ ÊÌ—«Ì‘ «Ê·ÌÂ »—«Ì «Ì‰ ›Ì‘ ”Å—Ì ‘œÂ..»—«Ì «œ«„Â —„“ »« œ” —”Ì »«·« »“‰Ìœ"
                frmAccess.AccessStatus = EnumAccessStatus.Edit
                frmAccess.Show vbModal
                If frmAccess.ReturnAccess = True Then
                    EditForTime = True
                End If
            Else
                ShowDisMessage "“„«‰ ÊÌ—«Ì‘ »—«Ì «Ì‰ ›Ì‘ ”Å—Ì ‘œÂ..", 2000
            End If
        Else
            ShowDisMessage "œ” —”Ì »—«Ì ÊÌ—«Ì‘ ›«ò Ê— ÊÃÊœ ‰œ«—œ..", 2000
        End If
    End If
    If rctmp.State = adStateOpen Then If rctmp.State = adStateOpen Then rctmp.Close
Exit Function
ErrHandler:
    ShowDisMessage err.Description, 2000

End Function
Private Function EditForSomeFich() As Boolean
    Dim tempCurNo As String
    EditForSomeFich = False
    If rctmp.State = 1 Then If rctmp.State = adStateOpen Then rctmp.Close
    
    If ClsFormAccess.EditAllStationsFactors = True Then
        EditForSomeFich = True
    ElseIf ClsFormAccess.EditAllFichUser = True And mvarCurUserNo = dblFichUser Then
        EditForSomeFich = True
    ElseIf ClsFormAccess.EditInvoice = True Then
        tempCurNo = txtNo.Text
        
        ReDim Parameter(0) As Parameter
        Parameter(0) = GenerateInputParameter("@CurUserId", adInteger, 4, mvarCurUserNo)
        Set rctmp = RunParametricStoredProcedure2Rec("Get_tblTotal_tFacm_EditForSomeFich_By_User", Parameter)
        Do While Not (rctmp.EOF)
            If tempCurNo = rctmp![No] Then
             EditForSomeFich = True
             Exit Function
            End If
            rctmp.MoveNext
        Loop
        If rctmp.State = adStateOpen Then If rctmp.State = adStateOpen Then rctmp.Close
        EditForSomeFich = False
     Else
        EditForSomeFich = False
     End If


End Function
Private Function RefferedForSomeFich() As Boolean
    Dim tempCurNo As String
    RefferedForSomeFich = False
    If rctmp.State = 1 Then If rctmp.State = adStateOpen Then rctmp.Close
    If ClsFormAccess.RefferedAllStationsFactors = True Then
        RefferedForSomeFich = True
    ElseIf ClsFormAccess.RefferedAllFichUser = True And mvarCurUserNo = dblFichUser Then
        RefferedForSomeFich = True
    ElseIf ClsFormAccess.RefferInvoice = True Then
        tempCurNo = txtNo.Text
        
         ReDim Parameter(0) As Parameter
        Parameter(0) = GenerateInputParameter("@CurUserId", adInteger, 4, mvarCurUserNo)
        Set rctmp = RunParametricStoredProcedure2Rec("Get_tblTotal_tFacm_RefferedForSomeFich_By_User", Parameter)
        Do While Not (rctmp.EOF)
            If tempCurNo = rctmp![No] Then
             RefferedForSomeFich = True
             Exit Function
            End If
            rctmp.MoveNext
        Loop
        If rctmp.State = adStateOpen Then If rctmp.State = adStateOpen Then rctmp.Close
        RefferedForSomeFich = False
     Else
        RefferedForSomeFich = False
     End If
End Function
Private Function GetInvoiceUI() As ClsInvoice
    Dim CInvoice As New ClsInvoice
    
    DetailsString1 = ""
    With FlxDetail
        For i = 1 To MaxRowFlexGrid - 1
            DetailsString1 = GenerateDetailsString3(DetailsString1, .TextMatrix(i, 1), .TextMatrix(i, 5), Val(.TextMatrix(i, 3)), Val(.TextMatrix(i, 11)), Val(.TextMatrix(i, 12)), .TextMatrix(i, 13), " ", .TextMatrix(i, 14), "", .TextMatrix(i, 8), .TextMatrix(i, 9))
        Next i
    End With
            
    CInvoice.No = Val(txtNo.Text)
    CInvoice.Status = mvarStatus
    CInvoice.Owner = 0
    CInvoice.Customer = Me.lblCustomer.Tag
    CInvoice.DiscountTotal = Val(Me.lblDiscountTotal)
    CInvoice.CarryFeeTotal = Val(Me.lblCarryFeeTotal)
    CInvoice.Recursive = Val(Me.txtRecursive.Text)
    If mvarServePlace = Delivery Then
        CInvoice.Incharge = IIf(CmbPayk.ItemData(CmbPayk.ListIndex) = -1, 0, CmbPayk.ItemData(CmbPayk.ListIndex))
    Else
        CInvoice.Incharge = IIf(cmbGarson.ItemData(cmbGarson.ListIndex) = -1, 0, cmbGarson.ItemData(cmbGarson.ListIndex))
    End If
    CInvoice.FacPayment = Abs(CInt(boolPayment))
    CInvoice.OrderType = mVarOrderType
    CInvoice.StationId = clsArya.StationNo
    CInvoice.ServiceTotal = Val(Me.lblServiceTotal) 'ServiceRate
    CInvoice.PackingTotal = Val(Me.lblPackingTotal)
    CInvoice.BascoleNo = 0
    CInvoice.TableNo = IIf(cmbTable.ItemData(cmbTable.ListIndex) = -1, 0, cmbTable.ItemData(cmbTable.ListIndex))
    CInvoice.User = mvarCurUserNo
    CInvoice.DateInvoice = txtDate.Text
    CInvoice.DetailsString = DetailsString1
    CInvoice.sFactorReceived = sFactorReceived
    CInvoice.Balance = Abs(CInt(BalancePayment))
    CInvoice.AccountYear = AccountYear
    CInvoice.NvcDescription = Right(txtDescription.Text, 150)
    CInvoice.TempAddress = TxtTempAddress.Text
    CInvoice.GuestNo = Trim(TxtGuestNo.Text)
    CInvoice.TaxTotal = Val(Me.lblTaxTotal)
    CInvoice.DutyTotal = Val(Me.LblDutyTotal)
    
    Set GetInvoiceUI = CInvoice
    Set CInvoice = Nothing
End Function
Private Function CheckChangeInvoice(FirstInvoce As ClsInvoice, LastInvoce As ClsInvoice) As Boolean
    
    If FirstInvoce.DateInvoice = LastInvoce.DateInvoice And FirstInvoce.No = LastInvoce.No And FirstInvoce.Status = LastInvoce.Status And FirstInvoce.Customer = LastInvoce.Customer And _
        FirstInvoce.DiscountTotal = LastInvoce.DiscountTotal And FirstInvoce.CarryFeeTotal = LastInvoce.CarryFeeTotal And _
        FirstInvoce.Recursive = LastInvoce.Recursive And _
        FirstInvoce.ServiceTotal = LastInvoce.ServiceTotal And _
        FirstInvoce.PackingTotal = LastInvoce.PackingTotal And FirstInvoce.TableNo = LastInvoce.TableNo And FirstInvoce.DetailsString = LastInvoce.DetailsString And _
        FirstInvoce.TaxTotal = LastInvoce.TaxTotal And FirstInvoce.DutyTotal = LastInvoce.DutyTotal And FirstInvoce.NvcDescription = LastInvoce.NvcDescription And _
        FirstInvoce.Incharge = LastInvoce.Incharge And FirstInvoce.TempAddress = LastInvoce.TempAddress And _
        FirstInvoce.User = LastInvoce.User And FirstInvoce.StationId = LastInvoce.StationId And FirstInvoce.OrderType = LastInvoce.OrderType And _
        FirstInvoce.Balance = LastInvoce.Balance And FirstInvoce.GuestNo = LastInvoce.GuestNo And lblPayFactorTotal.Caption = PreReceived Then
        CheckChangeInvoice = False
    Else
        CheckChangeInvoice = True
    End If
End Function
Private Function GetCostDifferences() As Long
    GetCostDifferences = 0
    For i = 0 To UBound(ArrCostDifferences)
        If ArrCostDifferences(i) <> 0 Then
            GetCostDifferences = GetCostDifferences + ArrCostDifferences(i)
        End If
    Next
    
End Function

' This event will fire when a new callerID detect by the device.
' USBCIDNumber : Caller Number i.e 989155714862.

' USBCIDDate   : Date of the new CallerID , this value is send by
'                service provider and may be incorrect.
'                in DTMF based callerID system's date parameter is not exist
'                and hence it will set to 'Not set' string.

' USBCIDTime   : Time of the new CallerID , this value is send by
'                service provider and may be incorrect.
'                in DTMF based callerID system's time parameter is not exist
'                and hence it will set to 'Not set' string.

' USBCIDChannel: this parameter indicate which channel has detected this callerID (or which Tel Line)
'                this value is between 1 to 127 for devices which have 1 to 127 channel.
'                for example for a 4 channel device this value is a number between 1 to 4

' CID_System   : this value indicate which callerID system was used in this callerID(DTMF or FSK)
'                it return 'FSK' string for fsk based system's and 'DTMF' string for dtmf based system's.

' Some exception's:  ======================================================================
'1: if an error ocured during callerID processing all of
'   the above parameter will set to 'Err !' string.
'2: if some piece's of the protocol was not detectable by the device some of the above
'   parameter may set to 'Unknow !' string.
'3: sometime in some callerID system's caller number is set to -private- or -out of area-
'   if this conditions detect by the device USBCIDNumber parameter will set to 'Private Number'
'   or 'Out of Area' string.
Private Sub USBCallerID1_NewCID(USBCIDNumber As String, USBCIDDate As String, USBCIDTime As String, USBCIDChannel As Long, CID_System As String)
        Dim Inputstr As String
        Dim kk As Integer
        If InStr(1, USBCIDNumber, "Unknow", 1) > 0 Then Exit Sub
        Inputstr = "L" & USBCIDChannel & USBCIDNumber & "@"
        kk = InStr(1, Inputstr, "L", 1)
        LineNumber = Val(Mid(Inputstr, kk + 1, 1))
        If LineNumber > 0 And LineNumber < 9 And kk > 0 Then
                Dim jj As Integer
                jj = InStr(1, Inputstr, "@", 1)
                If jj > kk Then
                    FWModem(Val(Mid(Inputstr, kk + 1, 1)) - 1).BackColor = vbRed ' &H80000003&
                    FWModem(Val(Mid(Inputstr, kk + 1, 1)) - 1).ToolTipText = Val(Mid(Inputstr, kk + 1, jj - kk - 2)) '
                    GetCallerInfo 1, Inputstr, LineNumber
                End If
        End If
        
''''        AddToLog "New CallerID(" & USBCIDChannel & ")" & _
''''             " Number:" & USBCIDNumber & _
''''             "  Date:" & USBCIDDate & _
''''             "  Time:" & USBCIDTime & _
''''             " .::" & CID_System
''    txtCID(USBCIDChannel - 1).Text = USBCIDNumber
End Sub

Private Sub DeliveredOrder()
'    flgShowOrderDetail = False
'
'    frmMsg.fwlblMsg.Caption = "¬Ì« ”›«—‘  ”ÊÌÂ „Ì ‘Êœø "
'    frmMsg.fwBtn(0).ButtonType = flwButtonOk
'    frmMsg.fwBtn(0).Caption = "»·Ì"
'    frmMsg.fwBtn(1).Visible = flwButtonCancel
'    frmMsg.fwBtn(1).Caption = "ŒÌ—"
'    frmMsg.fwBtn(1).Default = True
'    frmMsg.Show vbModal
'    If mvarMsgIdx = vbYes Then
'        frmFactorReceived.FWBtnPrint.Visible = True
'        frmFactorReceived.Show vbModal
'        ReDim Parameter(4) As Parameter
'        Parameter(0) = GenerateInputParameter("@No", adBigInt, 8, OrderNo)
'        Parameter(1) = GenerateInputParameter("@Status", adInteger, 4, EnumFactorType.Order)
'        Parameter(2) = GenerateInputParameter("@UID", adInteger, 4, mvarCurUserNo)
'        Parameter(3) = GenerateInputParameter("@AccountYear", adSmallInt, 2, AccountYear)
'        Parameter(4) = GenerateInputParameter("@ds", adWChar, 4000, sFactorReceived)
'
'        RunParametricStoredProcedure "Update_tfacm_Balance", Parameter
'    Else
'        Update
'    End If
       
End Sub

Function Zaman() As String
On Error Resume Next

Dim Str, TDate As String
Dim ym, mm, Dm, Ys, Ms, Ds As Integer
Dim Kb, Kb2 As Boolean
Kb2 = False
TDate = ""
Str = Date 'Zm1
ym = Val(left$(Str, 4))
mm = Val(Mid(Str, 6, 2))
Dm = Val(Right$(Str, 2))

If (mm < 3) Then
    If (ym - 1) Mod 4 = 0 Then Kb = True Else Kb = False
Else
    If (ym - 1) Mod 4 = 0 And mm = 3 And Dm < 21 Then
        Kb = True
        Kb2 = True
    Else
        If (ym Mod 4) = 0 Then Kb = True Else Kb = False
    End If
End If

Ys = ym - 621
Ms = mm - 3
Ds = Dm + 9

Select Case mm
    Case 1
        If Kb Then
            If Ds > 28 Then Ds = Ds - 28: Ms = 11: Ys = Ys - 1 Else Ds = Ds + 2: Ms = 10: Ys = Ys - 1
        Else
            If Ds > 29 Then Ds = Ds - 29: Ms = 11: Ys = Ys - 1 Else Ds = Ds + 1: Ms = 10: Ys = Ys - 1
        End If
    Case 2
        If Kb Then
            If Ds > 27 Then Ds = Ds - 27: Ms = 12: Ys = Ys - 1 Else Ds = Ds + 3: Ms = 11: Ys = Ys - 1
        Else
            If Ds > 28 Then Ds = Ds - 28: Ms = 12: Ys = Ys - 1 Else Ds = Ds + 2: Ms = 11: Ys = Ys - 1
        End If
    Case 3
        If Kb Then
            If Ds > 28 And (Not Kb2) Then Ds = Ds - 28: Ms = 1 Else Ds = Ds + 1: Ms = 12: Ys = Ys - 1
        Else
            If Ds > 29 Then Ds = Ds - 29: Ms = 1 Else Ms = 12: Ys = Ys - 1
        End If
    Case 4
        If Kb Then
            If Ds > 28 Then Ds = Ds - 28: Ms = Ms + 1 Else Ds = Ds + 3
        Else
            If Ds > 29 Then Ds = Ds - 29: Ms = 2 Else Ds = Ds + 2: Ms = 1
        End If
    Case 5
        If Kb Then
            If Ds > 29 Then Ds = Ds - 29: Ms = Ms + 1 Else Ds = Ds + 2
        Else
            If Ds > 30 Then Ds = Ds - 30: Ms = 3 Else Ds = Ds + 1: Ms = 2
        End If
    Case 6
        If Kb Then
            If Ds > 29 Then Ds = Ds - 29: Ms = Ms + 1 Else Ds = Ds + 2
        Else
            If Ds > 30 Then Ds = Ds - 30: Ms = 4 Else Ds = Ds + 1: Ms = 3
        End If
    Case 7
        If Kb Then
            If Ds > 30 Then Ds = Ds - 30: Ms = Ms + 1 Else Ds = Ds + 1
        Else
            If Ds > 31 Then Ds = Ds - 31: Ms = 5
        End If
    Case 8
        If Kb Then
            If Ds > 30 Then Ds = Ds - 30: Ms = Ms + 1 Else Ds = Ds + 1
        Else
            If Ds > 31 Then Ds = Ds - 31: Ms = 6
        End If
    Case 9
        If Kb Then
            If Ds > 30 Then Ds = Ds - 30: Ms = Ms + 1 Else Ds = Ds + 1
        Else
            If Ds > 31 Then Ds = Ds - 31: Ms = 7
        End If
    Case 10
        If Kb Then
            If Ds > 30 Then Ds = Ds - 30: Ms = Ms + 1
        Else
            If Ds > 31 Then Ds = Ds - 31: Ms = 8 Else Ds = Ds - 1
        End If
    Case 11
        If Kb Then
            If Ds > 29 Then Ds = Ds - 29: Ms = Ms + 1 Else Ds = Ds + 1
        Else
            If Ds > 30 Then Ds = Ds - 30: Ms = 9
        End If
    Case 12
        If Kb Then
            If Ds > 29 Then Ds = Ds - 29: Ms = Ms + 1 Else Ds = Ds + 1
        Else
            If Ds > 30 Then Ds = Ds - 30: Ms = 10
        End If
End Select
If Ms < 10 Then TDate = Ys & "/0" & Ms Else TDate = Ys & "/" & Ms
If Ds < 10 Then TDate = TDate & "/0" & Ds Else TDate = TDate & "/" & Ds
Zaman = TDate & " " & time$()
End Function

Private Function FixTel(ByVal Source As String) As String

On Error Resume Next
'»—‰«„Â  ’ÕÌÕ ‘„«—Â  ·›‰ Ê—ÊœÌ(ﬂ«·—¬ÌœÌ)«“„Êœ„ Ê  »œÌ· »’Ê—  —«ÌÃ
'«ê— ÿÊ· —‘ Â Ê—ÊœÌ ﬂ„ — «“ 9 »«‘œ Ì⁄‰Ì œ—”  Ì« ›«ﬁœ ﬂœ «” 
'Ì« ‘„«—Â ‰«ﬁ’ œ—Ì«›  ‘œÂ ﬂÂ Â„«‰ ‘„«—Â ê“«—‘ „Ì‘Êœ
Dim tmp As String
tmp = Trim(Source)
If Len(tmp) < 8 + 1 Then FixTel = tmp: Exit Function
'Õ–› + «Õ „«·Ì «Ê· ‘„«—Â
If left(tmp, 1) = "+" Then tmp = Right(tmp, Len(tmp) - 1)
'Õ–› œÊ ’›— «Õ „«·Ì «Ê· ‘„«—Â
If left(tmp, 2) = "00" Then tmp = Right(tmp, Len(tmp) - 2)
'Õ–› ﬂœ«Õ „«·Ì ò‘Ê— «“ «Ê· ‘„«—Â
'-----If Len(tmp) > 9 And Left(tmp, Len(Trim(Text1.Text))) = TCode1 Then tmp = Right(tmp, Len(tmp) - Len(Trim(Text1.Text)))
If Len(tmp) > 9 And left(tmp, 2) = "98" Then tmp = Right(tmp, Len(tmp) - Len("98"))
'Õ–› ﬂœ «Õ „«·Ì «” «‰ «“ «Ê· ‘„«—Â
If Len(tmp) > 9 And left(tmp, 1) = "0" Then tmp = Right(tmp, Len(tmp) - 1)
'Õ–› ﬂœ «Õ „«·Ì «” «‰ ﬂÂ »œÊ‰ ’›— ¬„œÂ »«‘œ
If Len(tmp) > 9 And left(tmp, 2) = "21" Then tmp = Right(tmp, Len(tmp) - 2)
'œ— ’Ê— Ì ﬂÂ ‘„«—Â «“ 9 —ﬁ„ »Ì‘ — »«‘œ Ê»«ﬂœ 9 »«‘œ
'Ì⁄‰Ì „À· 91 Â„—«Â Ì« 93  «·Ì« ’›—Ì ÃÂ  ﬂ«„· ‘œ‰ ‘„«—Â «÷«›Â „Ìê—œœ
If Len(tmp) > 9 And left(tmp, 1) <> "0" Then tmp = "0" & tmp
FixTel = tmp
End Function
Function B2D(ByVal T1 As String, ByVal T2 As String) As String
On Error Resume Next
'„Õ«”»Â ›«’·Â »Ì‰ œÊ “„«‰ ‘«„·  «—ÌŒ Ê ”«⁄  œﬁÌﬁÂ À«‰ÌÂ
'ıSource to be T1 = yyyy/mm/dd hh:mm:ss  is Upper
'Source to be T2 = yyyy/mm/dd hh:mm:ss  is Lower
'Œ—ÊÃÌ »— Õ”» „ﬁœ«— «Œ ·«› = yyyy/mm/dd hh:mm:ss
Dim tmp As String
Dim Y1, M1, D1, H1, N1, S1 As Integer
Dim Y2, M2, D2, H2, N2, S2 As Integer
Dim y, m, D, h, n, s  As Integer
y = 0: m = 0: D = 0: h = 0: n = 0: s = 0
If T1 = T2 Then B2D = "0000/00/00 00:00:00": Exit Function
If T1 < T2 Then tmp = T1: T1 = T2: T2 = tmp
Y1 = Val(left$(T1, 4)): M1 = Val(Mid(T1, 6, 2)): D1 = Val(Mid(T1, 9, 2)): H1 = Val(Mid(T1, 12, 2)): N1 = Val(Mid(T1, 15, 2)): S1 = Val(Mid(T1, 18, 2))
Y2 = Val(left$(T2, 4)): M2 = Val(Mid(T2, 6, 2)): D2 = Val(Mid(T2, 9, 2)): H2 = Val(Mid(T2, 12, 2)): N2 = Val(Mid(T2, 15, 2)): S2 = Val(Mid(T2, 18, 2))

If S1 = S2 Then
    s = 0
Else
    If S1 > S2 Then s = S1 - S2 Else s = (60 - S2) + S1: n = n - 1
End If

If N1 = N2 Then
    If n < 0 Then h = h - 1: n = n + (60 - N2) + N1 Else n = 0
Else
    If N1 > N2 Then n = n + N1 - N2 Else n = n + (60 - N2) + N1: h = h - 1
End If

If H1 = H2 Then
    If h < 0 Then D = D - 1: h = h + (24 - H2) + H1 Else h = 0
Else
    If H1 > H2 Then h = h + H1 - H2 Else h = h + (24 - H2) + H1: D = D - 1
End If

If D1 = D2 Then
    If D < 0 Then m = m - 1: D = D + (31 - D2) + D1 Else D = 0
Else
    If D1 > D2 Then D = D + D1 - D2 Else D = D + (31 - D2) + D1: m = m - 1
End If

If M1 = M2 Then
    If m < 0 Then y = y - 1: m = m + (12 - M2) + M1 Else m = 0
Else
    If M1 > M2 Then m = m + M1 - M2 + 1 Else m = m + (12 - M2) + M1: y = y - 1
End If

If Y1 = Y2 Then y = 0 Else y = Y1 - Y2 + y

B2D = FitN(y, 4) & "/" & FitN(m, 2) & "/" & FitN(D, 2) & " " & FitN(h, 2) & ":" & FitN(n, 2) & ":" & FitN(s, 2)
End Function

Function B2DS(ByVal Tin1 As String, ByVal Tin2 As String) As Long
On Error Resume Next
Dim T2 As String
    T2 = B2D(Tin1, Tin2)
Dim T3 As Long
    T3 = Val(Mid$(T2, 18, 2)) + (Val(Mid$(T2, 15, 2)) * 60) + _
        (Val(Mid$(T2, 12, 2)) * 3600) + (Val(Mid$(T2, 9, 2)) * 86400)
If T3 > 9999 Then B2DS = 9999 Else B2DS = T3
End Function
Public Sub CopyFile2Wav(FilePathVoiceDataOnly As String, FilePathWAVEwithHeader As String)
On Error GoTo ErrHan

Dim a%, buffer%, temp$, fRead&, fSize&, b%
If clsStation.VoiceRecord Then
    fSize = FileLen(FilePathVoiceDataOnly)
        'If fSize < 10000 Then GoTo MinSize
    a = FreeFile
        Open FilePathVoiceDataOnly For Binary Access Read As a
    b = FreeFile
        Open FilePathWAVEwithHeader For Binary Access Write As b
    Dim hFile(1 To 44) As String
        hFile(1) = "R":  hFile(2) = "I":  hFile(3) = "F":  hFile(4) = "F"
        hFile(5) = Chr$((fSize - 4) Mod 256)
        hFile(6) = Chr$(Int((fSize - 4) / 256) Mod 256)
        hFile(7) = Chr$(Int((fSize - 4) / 65536) Mod 256)
        hFile(8) = Chr$(Int((fSize - 4) / 16777216) Mod 256)
        hFile(9) = "W": hFile(10) = "A": hFile(11) = "V": hFile(12) = "E"
        hFile(13) = "f": hFile(14) = "m": hFile(15) = "t": hFile(16) = " "
        hFile(17) = Chr$(16): hFile(18) = Chr$(0): hFile(19) = Chr$(0): hFile(20) = Chr$(0)
        hFile(21) = Chr$(1): hFile(22) = Chr$(0): hFile(23) = Chr$(1): hFile(24) = Chr$(0)
        hFile(25) = Chr$(49) 'Chr(8000 Mod 256)
        hFile(26) = Chr$(31) 'Chr(Int(8000 / 256))
        hFile(27) = Chr$(0): hFile(28) = Chr$(0)
        hFile(29) = Chr$(49) 'Chr(Int(8000 Mod 256))
        hFile(30) = Chr$(31) 'Chr(Int(8000 / 256))
        hFile(31) = Chr$(0): hFile(32) = Chr$(0): hFile(33) = Chr$(1): hFile(34) = Chr$(0)
        hFile(35) = Chr$(8): hFile(36) = Chr$(0)
        hFile(37) = "d": hFile(38) = "a": hFile(39) = "t": hFile(40) = "a"
        hFile(41) = Chr$(fSize Mod 256)
        hFile(42) = Chr$(Int(fSize / 256) Mod 256)
        hFile(43) = Chr$(Int(fSize / 65536) Mod 256)
        hFile(44) = Chr$(Int(fSize / 16777216) Mod 256)
    For i = 1 To 44: Put b, , hFile(i): Next i
    buffer = 4048
    fRead = 0
    While fRead < fSize
        If buffer > (fSize - fRead) Then buffer = (fSize - fRead)
        temp = Space(buffer)
        Get a, , temp
        Put b, , temp
        fRead = fRead + buffer
        DoEvents
    Wend
    Close a
    Close b
    Kill FilePathVoiceDataOnly
End If
Exit Sub

ErrHan:
    MsgBox "«‘ò«· œ— ”«Œ  ›«Ì· ’Ê Ì" & vbCrLf & err.Description, vbCritical + vbMsgBoxRtlReading
End Sub


Function FitN(ByVal Source As Long, ByVal Fit As Long) As String
On Error Resume Next
'»—ê—œ«‰ „⁄«œ· —‘ Â Ìò ⁄œœ »« ÿÊ· „‘Œ’ . œ—’Ê—  ò”— ’›— »Â «Ê· ⁄œœ «÷«›Â „Ì‘Êœ
If Len(Trim$(Str$(Source))) = Fit Then
    FitN = Fulltrim$(Str$(Source))
ElseIf Len(Trim$(Str$(Source))) > Fit Then
    FitN = Right$(Trim$(Str$(Source)), Fit)
Else
    FitN = Fill("0", Fit - Len(Fulltrim$(Str$(Source)))) & Fulltrim$(Str$(Source))
End If
End Function
Function Fulltrim(ByVal Source As String) As String
On Error Resume Next
'»—ê—œ«‰ —‘ Â Å” «“ Õ–› ò·ÌÂ ›«’·Â Â«Ì ﬁ»· Ê »⁄œ Ê „«»Ì‰ —‘ Â
Dim T As String
T = ""
For i = 1 To Len(Source)
    If Not Mid$(Source, i, 1) = " " Then T = T + Mid$(Source, i, 1)
    Next
Fulltrim = T
End Function
Function Fill(ByVal Source As String, ByVal Num As Integer) As String
On Error Resume Next
'»—ê—œ«‰  ⁄œ«œ „‘Œ’ «“ Ìò ò«—ò — „‘Œ’ ‘œÂ
Dim T As String
For i = 1 To Num
    T = T + Source
Next
Fill = T
End Function

