VERSION 5.00
Object = "{7AEDC602-D94C-11D1-BB7A-00E0290EA3C9}#1.0#0"; "ResizeKit.ocx"
Begin VB.Form frmInvoicesetting 
   BackColor       =   &H00E0E0E0&
   ClientHeight    =   7290
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7620
   Icon            =   "frmInvoicesetting.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7290
   ScaleWidth      =   7620
   Begin VB.Frame Frame_öShowDailyFich 
      Height          =   855
      Left            =   120
      TabIndex        =   51
      Top             =   6360
      Width           =   3495
      Begin VB.CheckBox ChkTemporayNo 
         Alignment       =   1  'Right Justify
         Caption         =   "äãÇíÔ ÔãÇÑå ÝíÔ ÑæÒÇäå "
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
         Left            =   600
         TabIndex        =   52
         Top             =   240
         Width           =   2535
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "ÊäÙíãÇÊ äæÇÑ ÇÝÞí ÈÇáÇí ÕÝÍå"
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
      Height          =   1575
      Left            =   -495
      RightToLeft     =   -1  'True
      TabIndex        =   46
      Top             =   2400
      Visible         =   0   'False
      Width           =   7455
      Begin VB.CheckBox ChkTelephone 
         Alignment       =   1  'Right Justify
         Caption         =   "äãÇíÔ Âíßæä ÇäÊÎÇÈ ÊáÝä"
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   50
         Top             =   960
         Width           =   2775
      End
      Begin VB.CheckBox ChkColor 
         Alignment       =   1  'Right Justify
         Caption         =   "äãÇíÔ Âíßæä ÇäÊÎÇÈ Ñä"
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         TabIndex        =   49
         Top             =   480
         Width           =   2775
      End
      Begin VB.CheckBox ChkKeyboard 
         Alignment       =   1  'Right Justify
         Caption         =   "äãÇíÔ Âíßæä ßíÈæÑÏ"
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4440
         TabIndex        =   48
         Top             =   960
         Width           =   2175
      End
      Begin VB.CheckBox ChkLanguage 
         Alignment       =   1  'Right Justify
         Caption         =   "äãÇíÔ Âíßæä ÇäÊÎÇÈ ÒÈÇä"
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4080
         TabIndex        =   47
         Top             =   480
         Width           =   2535
      End
      Begin VB.Image Image5 
         BorderStyle     =   1  'Fixed Single
         Height          =   540
         Left            =   6720
         Picture         =   "frmInvoicesetting.frx":A4C2
         Top             =   360
         Width           =   540
      End
      Begin VB.Image Image4 
         BorderStyle     =   1  'Fixed Single
         Height          =   540
         Left            =   3120
         Picture         =   "frmInvoicesetting.frx":A7CC
         Top             =   960
         Width           =   540
      End
      Begin VB.Image Image3 
         BorderStyle     =   1  'Fixed Single
         Height          =   540
         Left            =   6720
         Picture         =   "frmInvoicesetting.frx":B096
         Top             =   960
         Width           =   540
      End
      Begin VB.Image Image2 
         BorderStyle     =   1  'Fixed Single
         Height          =   540
         Left            =   3120
         Picture         =   "frmInvoicesetting.frx":B3A0
         Top             =   360
         Width           =   540
      End
   End
   Begin VB.Frame Frame_ScreenSaver 
      Height          =   855
      Left            =   3720
      TabIndex        =   42
      Top             =   6360
      Width           =   3855
      Begin VB.TextBox txtScreenSaverTime 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   720
         MaxLength       =   3
         TabIndex        =   43
         Text            =   "8"
         Top             =   240
         Width           =   495
      End
      Begin VB.Image Image1 
         BorderStyle     =   1  'Fixed Single
         Height          =   540
         Left            =   3240
         Picture         =   "frmInvoicesetting.frx":B6AA
         Top             =   240
         Width           =   540
      End
      Begin VB.Label lblScreenSaver 
         Caption         =   "ÊÑß ÕäÏæÞ ÇÊæãÇÊíß"
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
         Left            =   1320
         RightToLeft     =   -1  'True
         TabIndex        =   45
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label lblMinute 
         Caption         =   "ÏÞíÞå"
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
         TabIndex        =   44
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   " ÊäÙíãÇÊ ãæäíÊæÑ Ïæã"
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
      Height          =   5895
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   28
      Top             =   480
      Width           =   3495
      Begin VB.Frame Frame_Monitor2 
         BeginProperty Font 
            Name            =   "Nazanin"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   53
         Top             =   2400
         Width           =   3255
         Begin VB.CheckBox ChkShowLogo 
            Alignment       =   1  'Right Justify
            Height          =   375
            Left            =   2760
            TabIndex        =   54
            Top             =   360
            Width           =   255
         End
         Begin VB.Label Label8 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "\Image\Logo.jpg"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   375
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   57
            Top             =   720
            Width           =   1815
         End
         Begin VB.Label Label7 
            Caption         =   "æÇÞÚ ÏÑ"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2400
            RightToLeft     =   -1  'True
            TabIndex        =   56
            Top             =   720
            Width           =   615
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            Caption         =   "äãÇíÔ ÞíãÊ ÝÑæÔ ÈÇ áææ ÈÒÑ"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   55
            Top             =   360
            Width           =   2535
         End
      End
      Begin VB.Frame Frame_GoodMenu 
         Height          =   1815
         Left            =   120
         TabIndex        =   37
         Top             =   3960
         Width           =   3255
         Begin VB.TextBox TxtGoodMenuFileName 
            Alignment       =   2  'Center
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
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   40
            Text            =   "Total_GoodMenu"
            Top             =   1080
            Width           =   2925
         End
         Begin VB.CheckBox ChkGoodMenuView 
            Alignment       =   1  'Right Justify
            Height          =   375
            Left            =   2760
            TabIndex        =   38
            Top             =   240
            Width           =   255
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            Caption         =   "äÇã ÝÇíá ãäæåÇ ÏÑ ÏÇíÑ˜ÊæÑí ÌÇÑí"
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
            TabIndex        =   41
            Top             =   720
            Width           =   2895
         End
         Begin VB.Label Label29 
            Alignment       =   1  'Right Justify
            Caption         =   "äãÇíÔ ÞíãÊ åÇí ãäæ  "
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
            Left            =   480
            TabIndex        =   39
            Top             =   240
            Width           =   2055
         End
      End
      Begin VB.Frame Frame_Monitor 
         Height          =   855
         Left            =   120
         TabIndex        =   34
         Top             =   1560
         Width           =   3255
         Begin VB.CheckBox ChkShowInvoice 
            Alignment       =   1  'Right Justify
            Height          =   375
            Left            =   2760
            TabIndex        =   35
            Top             =   240
            Width           =   255
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            Caption         =   "äãÇíÔ ÝÇ˜ÊæÑ ÝÑæÔ"
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
            Left            =   240
            TabIndex        =   36
            Top             =   240
            Width           =   2175
         End
      End
      Begin VB.Frame Frame_Picture 
         Height          =   1215
         Left            =   120
         TabIndex        =   29
         Top             =   360
         Width           =   3255
         Begin VB.TextBox TxtShowGoodTime 
            Alignment       =   2  'Center
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
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   31
            Text            =   "1000"
            Top             =   600
            Width           =   885
         End
         Begin VB.CheckBox ChkShowPictureGood 
            Alignment       =   1  'Right Justify
            Height          =   375
            Left            =   2760
            TabIndex        =   30
            Top             =   240
            Width           =   255
         End
         Begin VB.Label Label107 
            Alignment       =   1  'Right Justify
            Caption         =   "ÒãÇä äãÇíÔ ÊÕæíÑ ßÇáÇ"
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
            Left            =   840
            TabIndex        =   33
            Top             =   720
            Width           =   2295
         End
         Begin VB.Label Label106 
            Alignment       =   1  'Right Justify
            Caption         =   "äãÇíÔ ÊÕæíÑ ßÇáÇ"
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
            Left            =   360
            TabIndex        =   32
            Top             =   240
            Width           =   2175
         End
      End
   End
   Begin VB.Frame Frame13 
      Caption         =   "äãÇíÔ ÓÊæä åÇ"
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
      Height          =   5895
      Left            =   3720
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   465
      Width           =   3855
      Begin VB.CheckBox ChkTax 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   3240
         TabIndex        =   26
         Top             =   5490
         Width           =   255
      End
      Begin VB.CheckBox ChkDuty 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   3240
         TabIndex        =   24
         Top             =   5067
         Width           =   255
      End
      Begin VB.CheckBox ChkMojodi 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   3240
         TabIndex        =   22
         Top             =   4233
         Width           =   255
      End
      Begin VB.CheckBox ChkChanges 
         Alignment       =   1  'Right Justify
         Height          =   255
         Left            =   3240
         TabIndex        =   12
         Top             =   2685
         Width           =   255
      End
      Begin VB.CheckBox ChkRow 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   3240
         TabIndex        =   11
         Top             =   600
         Width           =   255
      End
      Begin VB.CheckBox ChkFee 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   3240
         TabIndex        =   10
         Top             =   1017
         Width           =   255
      End
      Begin VB.CheckBox ChkTotal 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   3240
         TabIndex        =   9
         Top             =   1434
         Width           =   255
      End
      Begin VB.CheckBox ChkUnitGood 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   3240
         TabIndex        =   8
         Top             =   2268
         Width           =   255
      End
      Begin VB.CheckBox ChkSeller 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   3240
         TabIndex        =   7
         Top             =   2982
         Width           =   255
      End
      Begin VB.CheckBox ChkDiscount 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   3240
         TabIndex        =   6
         Top             =   3399
         Width           =   255
      End
      Begin VB.CheckBox ChkGoodCode 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   3240
         TabIndex        =   5
         Top             =   1851
         Width           =   255
      End
      Begin VB.CheckBox ChkStore 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   3240
         TabIndex        =   4
         Top             =   4650
         Width           =   255
      End
      Begin VB.CheckBox ChkRate 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   3240
         TabIndex        =   3
         Top             =   3816
         Width           =   255
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         Caption         =   "ÇäÈÇÑ"
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
         Left            =   1800
         TabIndex        =   58
         Top             =   4680
         Width           =   1335
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "ãÇáíÇÊ"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1800
         TabIndex        =   27
         Top             =   5505
         Width           =   1335
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "ÚæÇÑÖ"
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
         Left            =   1800
         TabIndex        =   25
         Top             =   5088
         Width           =   1335
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "ãæÌæÏí"
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
         Left            =   1830
         TabIndex        =   23
         Top             =   4272
         Width           =   1335
      End
      Begin VB.Label Label39 
         Alignment       =   1  'Right Justify
         Caption         =   "ÊÛííÑÇÊ ßÇáÇ"
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
         TabIndex        =   21
         Top             =   2640
         Width           =   1575
      End
      Begin VB.Label Label27 
         Alignment       =   1  'Right Justify
         Caption         =   "ÑÏíÝ"
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
         Left            =   720
         TabIndex        =   20
         Top             =   600
         Width           =   2415
      End
      Begin VB.Label Label33 
         Alignment       =   1  'Right Justify
         Caption         =   "Ýí"
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
         Left            =   1080
         TabIndex        =   19
         Top             =   1008
         Width           =   2055
      End
      Begin VB.Label Label32 
         Alignment       =   1  'Right Justify
         Caption         =   "ÌãÚ ßá"
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
         Left            =   1080
         TabIndex        =   18
         Top             =   1416
         Width           =   2055
      End
      Begin VB.Label Label20 
         Alignment       =   1  'Right Justify
         Caption         =   "æÇÍÏ ßÇáÇ"
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
         Left            =   1080
         TabIndex        =   17
         Top             =   2232
         Width           =   2055
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "ÝÑæÔäÏå"
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
         Left            =   1680
         TabIndex        =   16
         Top             =   3048
         Width           =   1455
      End
      Begin VB.Label Label41 
         Alignment       =   1  'Right Justify
         Caption         =   "ÊÎÝíÝ"
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
         Left            =   1680
         TabIndex        =   15
         Top             =   3456
         Width           =   1455
      End
      Begin VB.Label Label68 
         Alignment       =   1  'Right Justify
         Caption         =   "ßÏ ßÇáÇ"
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
         Left            =   1080
         TabIndex        =   14
         Top             =   1824
         Width           =   2055
      End
      Begin VB.Label Label89 
         Alignment       =   1  'Right Justify
         Caption         =   "äÑÎ"
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
         Left            =   2160
         TabIndex        =   13
         Top             =   3864
         Width           =   975
      End
   End
   Begin RESIZEKITLibCtl.ResizeKit ResizeKit1 
      Height          =   480
      Left            =   480
      OleObjectBlob   =   "frmInvoicesetting.frx":BF74
      TabIndex        =   0
      Top             =   0
      Width           =   480
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ãÏíÑíÊ æ ÊäÙíãÇÊ ÝÇßÊæÑ ÝÑæÔ"
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
      Left            =   1800
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   -120
      Width           =   4095
   End
End
Attribute VB_Name = "frmInvoicesetting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Parameter() As Parameter



Private Sub Form_Activate()
    
    VarActForm = Me.Name
    
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

Private Sub Form_Load()
 
    SetFirstToolBar
    
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
    
    If intVersion = Min Then
        Frame_ScreenSaver.Enabled = False
        Frame_öShowDailyFich.Enabled = False
        Frame_Picture.Enabled = False
        Frame_GoodMenu.Enabled = False
        Frame_Monitor.Enabled = False
        Frame_Monitor2.Enabled = False
    End If
        
    If clsInvoiceValue.ColChanges = True Then
       ChkChanges.Value = 1
    Else
       ChkChanges.Value = 0
    End If
    
    If clsInvoiceValue.ColDiscount = True Then
       ChkDiscount.Value = 1
    Else
       ChkDiscount.Value = 0
    End If
    
    If clsInvoiceValue.ColFee = True Then
       ChkFee.Value = 1
    Else
       ChkFee.Value = 0
    End If
    
    If clsInvoiceValue.ColGoodCode = True Then
       ChkGoodCode.Value = 1
    Else
       ChkGoodCode.Value = 0
    End If
    
    If clsInvoiceValue.ColMojodi = True Then
       ChkMojodi.Value = 1
    Else
       ChkMojodi.Value = 0
    End If
    
    If clsInvoiceValue.ColRate = True Then
       ChkRate.Value = 1
    Else
       ChkRate.Value = 0
    End If
    
    If clsInvoiceValue.ColRow = True Then
       ChkRow.Value = 1
    Else
       ChkRow.Value = 0
    End If
    
    If clsInvoiceValue.ColSeller = True Then
       ChkSeller.Value = 1
    Else
       ChkSeller.Value = 0
    End If
    
    If clsInvoiceValue.ColStore = True Then
       ChkStore.Value = 1
    Else
       ChkStore.Value = 0
    End If
    
    If clsInvoiceValue.ColTotal = True Then
       ChkTotal.Value = 1
    Else
       ChkTotal.Value = 0
    End If
    
    If clsInvoiceValue.ColUnitGood = True Then
       ChkUnitGood.Value = 1
    Else
       ChkUnitGood.Value = 0
    End If
    
    If clsInvoiceValue.ColDuty = True Then
       ChkDuty.Value = 1
    Else
       ChkDuty.Value = 0
    End If
    
    If clsInvoiceValue.ColTax = True Then
       ChkTax.Value = 1
    Else
       ChkTax.Value = 0
    End If
    
    If clsInvoiceValue.ShowPictureGood = True Then
        ChkShowPictureGood.Value = 1
    Else
        ChkShowPictureGood.Value = 0
    End If
    
    TxtShowGoodTime.Text = Val(clsInvoiceValue.ShowGoodTime)
    
    If clsInvoiceValue.ShowInvoiceMenu = True Then
        ChkShowInvoice.Value = 1
    Else
        ChkShowInvoice.Value = 0
    End If
    
    If clsInvoiceValue.ShowLogo = True Then
        ChkShowLogo.Value = 1
    Else
        ChkShowLogo.Value = 0
    End If
    
    If clsInvoiceValue.GoodMenuView = True Then
        ChkGoodMenuView.Value = 1
    Else
        ChkGoodMenuView.Value = 0
    End If
    
    If clsInvoiceValue.GoodMenuFileName <> "" Then
        TxtGoodMenuFileName.Text = clsInvoiceValue.GoodMenuFileName
    Else
        TxtGoodMenuFileName.Text = "Total_GoodMenu"
    End If
    
    txtScreenSaverTime.Text = Val(clsInvoiceValue.ScreenSaverTime)

    If clsInvoiceValue.LanguageIcon = True Then
        ChkLanguage.Value = 1
    Else
        ChkLanguage.Value = 0
    End If

    If clsInvoiceValue.KeyboardIcon = True Then
        ChkKeyboard.Value = 1
    Else
        ChkKeyboard.Value = 0
    End If

    If clsInvoiceValue.ColorIcon = True Then
        ChkColor.Value = 1
    Else
        ChkColor.Value = 0
    End If

    If clsInvoiceValue.TelephoneIcon = True Then
        ChkTelephone.Value = 1
    Else
        ChkTelephone.Value = 0
    End If
    
    On Error Resume Next
    Dim rctmp As New ADODB.Recordset
    ReDim Parameter(1) As Parameter
    Parameter(0) = GenerateInputParameter("@StationId", adInteger, 4, clsArya.StationNo)
    Parameter(1) = GenerateInputParameter("@Branch", adInteger, 4, CurrentBranch)
    Set rctmp = RunParametricStoredProcedure2Rec("Get_StationId_info", Parameter)
    
    If Not (rctmp.EOF = True And rctmp.BOF = True) Then
        ChkTemporayNo.Value = IIf(rctmp!TemporaryNo = True, 1, 0)
    End If
    rctmp.Close
    
'    If clsInvoiceValue.PrintLable = True Then
'        chkPrintLable.Value = 1
'    Else
'        chkPrintLable.Value = 0
'    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)

    If PosConnection.State = adStateOpen Then PosConnection.Close: Set PosConnection = Nothing
    VarActForm = ""
        
    SaveSetting strMainKey, Me.Name, "Left", Me.left
    SaveSetting strMainKey, Me.Name, "Top", Me.top


End Sub

Public Sub Update()

    clsInvoiceValue.ColChanges = ChkChanges.Value
    clsInvoiceValue.ColDiscount = ChkDiscount.Value
    clsInvoiceValue.ColFee = ChkFee.Value
    clsInvoiceValue.ColGoodCode = ChkGoodCode.Value
    clsInvoiceValue.ColMojodi = ChkMojodi.Value
    clsInvoiceValue.ColRate = ChkRate.Value
    clsInvoiceValue.ColRow = ChkRow.Value
    clsInvoiceValue.ColSeller = ChkSeller.Value
    clsInvoiceValue.ColStore = ChkStore.Value
    clsInvoiceValue.ColTotal = ChkTotal.Value
    clsInvoiceValue.ColUnitGood = ChkUnitGood.Value
    clsInvoiceValue.ColTax = ChkTax.Value
    clsInvoiceValue.ColDuty = ChkDuty.Value
    clsInvoiceValue.ShowPictureGood = ChkShowPictureGood.Value
    clsInvoiceValue.ShowGoodTime = Val(TxtShowGoodTime.Text)
    clsInvoiceValue.ShowInvoiceMenu = ChkShowInvoice.Value
    clsInvoiceValue.GoodMenuView = ChkGoodMenuView.Value
    clsInvoiceValue.GoodMenuFileName = TxtGoodMenuFileName.Text
    clsInvoiceValue.ScreenSaverTime = Val(txtScreenSaverTime.Text)
    clsInvoiceValue.LanguageIcon = ChkLanguage.Value
    clsInvoiceValue.KeyboardIcon = ChkKeyboard.Value
    clsInvoiceValue.ColorIcon = ChkColor.Value
    clsInvoiceValue.TelephoneIcon = ChkTelephone.Value
    clsInvoiceValue.ShowLogo = ChkShowLogo.Value
'    clsInvoiceValue.PrintLable = chkPrintLable.Value
    SetInvoiceSettingFile
   
    Call PresetScreenSaver
    clsStation.TemporaryNo = CBool(ChkTemporayNo.Value)
    ReDim Parameter(1) As Parameter
    Parameter(0) = GenerateInputParameter("@StationId", adInteger, 4, clsArya.StationNo)
    Parameter(1) = GenerateInputParameter("@TemporaryNo", adBoolean, 1, ChkTemporayNo.Value)
    RunParametricStoredProcedure "Update_tstations_TemporaryNo", Parameter
    
    Unload Me
    
End Sub

Public Sub SetFirstToolBar()
    
    AllButton vbOff, True
    
    mdifrm.Toolbar1.Buttons(8).Enabled = True  'Enter
    mdifrm.Toolbar1.Buttons(23).Enabled = True
    mdifrm.Toolbar1.Buttons(24).Enabled = True
    mdifrm.Toolbar1.Buttons(25).Enabled = True
    mdifrm.Toolbar1.Buttons(26).Enabled = True
    mdifrm.Toolbar1.Buttons(27).Enabled = True
    

End Sub

Public Sub ExitForm()

    Unload Me
    
End Sub

Private Sub ResizeKit1_ExitResize(ByVal XScale As Double, ByVal YScale As Double)
    If formloadFlag = True Then
        SaveSetting strMainKey, Me.Name, "Height", Me.Height
        SaveSetting strMainKey, Me.Name, "Width", Me.Width
    End If

End Sub

