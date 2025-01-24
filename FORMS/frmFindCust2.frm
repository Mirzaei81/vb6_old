VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{7AEDC602-D94C-11D1-BB7A-00E0290EA3C9}#1.0#0"; "ResizeKit.ocx"
Begin VB.Form frmFindCust2 
   Caption         =   "                                                                                     Ã” ÃÊÌ „‘ —ﬂ"
   ClientHeight    =   9990
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12270
   Icon            =   "frmFindCust2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9990
   ScaleWidth      =   12270
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   600
      Top             =   0
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   6600
      RightToLeft     =   -1  'True
      TabIndex        =   47
      Top             =   8760
      Width           =   3495
      Begin VB.TextBox txtMaxRecord 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   0
         RightToLeft     =   -1  'True
         TabIndex        =   48
         Text            =   "300"
         Top             =   140
         Width           =   615
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Õœ«ﬂÀ—  ⁄œ«œ Œ—ÊÃÌÂ« »—«Ì ‰„«Ì‘"
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
         Left            =   600
         RightToLeft     =   -1  'True
         TabIndex        =   49
         Top             =   120
         Width           =   2775
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00FFFFFF&
      Height          =   1935
      Left            =   4920
      RightToLeft     =   -1  'True
      TabIndex        =   43
      Top             =   1680
      Width           =   2895
      Begin VB.TextBox TxtTimer 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   1800
         RightToLeft     =   -1  'True
         TabIndex        =   44
         Text            =   "500"
         Top             =   960
         Width           =   615
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "„Ì·Ì À«‰ÌÂ"
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   600
         RightToLeft     =   -1  'True
         TabIndex        =   46
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "“„«‰  «ŒÌ— »Ì‰ ﬂ·ÌœÂ«Ì Ê—ÊœÌ"
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   360
         RightToLeft     =   -1  'True
         TabIndex        =   45
         Top             =   360
         Width           =   2175
      End
   End
   Begin VB.PictureBox Frame3 
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   4920
      RightToLeft     =   -1  'True
      ScaleHeight     =   1155
      ScaleWidth      =   2835
      TabIndex        =   40
      Top             =   360
      Width           =   2895
      Begin VB.OptionButton Option1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "Ã” ÃÊÌ ”—Ì⁄"
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
         Index           =   0
         Left            =   720
         RightToLeft     =   -1  'True
         TabIndex        =   42
         Top             =   600
         Width           =   1695
      End
      Begin VB.OptionButton Option1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         Caption         =   "Ã” ÃÊÌ „⁄„Ê·Ì"
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
         Index           =   1
         Left            =   600
         RightToLeft     =   -1  'True
         TabIndex        =   41
         Top             =   120
         Width           =   1815
      End
   End
   Begin VB.CommandButton CancelButton 
      BackColor       =   &H000000C0&
      Cancel          =   -1  'True
      Caption         =   "«‰’—«›"
      BeginProperty Font 
         Name            =   "B Homa"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   600
      RightToLeft     =   -1  'True
      TabIndex        =   39
      Top             =   8760
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      BackColor       =   &H0000C000&
      Caption         =   "ﬁ»Ê·"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "B Homa"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   10320
      RightToLeft     =   -1  'True
      TabIndex        =   38
      Top             =   8760
      Width           =   1215
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00FFFFFF&
      Height          =   3255
      Left            =   240
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   360
      Width           =   4455
      Begin VB.CommandButton CmdNewCust 
         BackColor       =   &H0000C0C0&
         Caption         =   "„‘ —ﬂ ÃœÌœ"
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   2640
         Width           =   1575
      End
      Begin VB.Label MaxPrice 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "»Ì‘ —Ì‰ Œ—Ìœ"
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3360
         RightToLeft     =   -1  'True
         TabIndex        =   36
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label LastDate 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "¬Œ—Ì‰ Œ—Ìœ"
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3360
         RightToLeft     =   -1  'True
         TabIndex        =   35
         Top             =   120
         Width           =   975
      End
      Begin VB.Label BuyAverage 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "„Ì«‰êÌ‰ Œ—Ìœ"
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3360
         RightToLeft     =   -1  'True
         TabIndex        =   34
         Top             =   600
         Width           =   975
      End
      Begin VB.Label LastNo 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "¬Œ—Ì‰ ›Ì‘"
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1200
         RightToLeft     =   -1  'True
         TabIndex        =   33
         Top             =   120
         Width           =   975
      End
      Begin VB.Label AddedDate 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   " «—ÌŒ ⁄÷ÊÌ "
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3360
         RightToLeft     =   -1  'True
         TabIndex        =   32
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label BuyCount 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   " ⁄œ«œ œ›⁄«  Œ—Ìœ"
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1200
         RightToLeft     =   -1  'True
         TabIndex        =   31
         Top             =   600
         Width           =   975
      End
      Begin VB.Label MinPrice 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "ﬂ„ —Ì‰ Œ—Ìœ"
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1200
         RightToLeft     =   -1  'True
         TabIndex        =   30
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label LastPrice 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "„»·€ ¬Œ—Ì‰ Œ—Ìœ"
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1200
         RightToLeft     =   -1  'True
         TabIndex        =   29
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label MaxPrice1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "B Mitra"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2280
         RightToLeft     =   -1  'True
         TabIndex        =   28
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label LastDate1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "B Mitra"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2520
         RightToLeft     =   -1  'True
         TabIndex        =   27
         Top             =   120
         Width           =   975
      End
      Begin VB.Label BuyAverage1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "B Mitra"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2280
         RightToLeft     =   -1  'True
         TabIndex        =   26
         Top             =   600
         Width           =   975
      End
      Begin VB.Label LastNo1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "B Mitra"
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
         Top             =   120
         Width           =   975
      End
      Begin VB.Label AddedDate1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "B Mitra"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2280
         RightToLeft     =   -1  'True
         TabIndex        =   24
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label BuyCount1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "B Mitra"
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
         TabIndex        =   23
         Top             =   600
         Width           =   975
      End
      Begin VB.Label MinPrice1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "B Mitra"
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
         TabIndex        =   22
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label LastPrice1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "B Mitra"
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
         TabIndex        =   21
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label LastCredit1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "B Mitra"
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
         TabIndex        =   20
         Top             =   2040
         Width           =   855
      End
      Begin VB.Label LastCredit 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "„«‰œÂ "
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1080
         RightToLeft     =   -1  'True
         TabIndex        =   19
         Top             =   2040
         Width           =   375
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "«⁄ »«—Ì"
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3840
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   2040
         Width           =   495
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "œ—Ì«› "
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2400
         RightToLeft     =   -1  'True
         TabIndex        =   17
         Top             =   2040
         Width           =   495
      End
      Begin VB.Label LblBuy1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "B Mitra"
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
         TabIndex        =   16
         Top             =   2040
         Width           =   855
      End
      Begin VB.Label LblRecieve1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "B Mitra"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1440
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   2040
         Width           =   855
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "‰—Œ"
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3840
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   2640
         Width           =   495
      End
      Begin VB.Label LbCustomerlRate 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "B Mitra"
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
         TabIndex        =   13
         Top             =   2640
         Width           =   615
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Œ—Ìœ Â«Ì «„—Ê“"
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2400
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   2640
         Width           =   1095
      End
      Begin VB.Label lblCountCurrentBuy 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "B Mitra"
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
         TabIndex        =   11
         Top             =   2640
         Width           =   615
      End
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00FFFFFF&
      Height          =   2175
      Left            =   8040
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   360
      Width           =   3975
      Begin VB.TextBox txtPhone 
         Alignment       =   1  'Right Justify
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
         Left            =   0
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   970
         Width           =   2595
      End
      Begin VB.TextBox txtAddress 
         Alignment       =   1  'Right Justify
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
         Height          =   700
         Left            =   0
         MultiLine       =   -1  'True
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   1420
         Width           =   2595
      End
      Begin VB.TextBox txtMembershipId 
         Alignment       =   1  'Right Justify
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
         Height          =   435
         Left            =   0
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   120
         Width           =   2595
      End
      Begin VB.TextBox txtName 
         Alignment       =   1  'Right Justify
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
         Height          =   435
         Left            =   0
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   555
         Width           =   2595
      End
      Begin VB.Label Label4 
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
         Height          =   435
         Left            =   2640
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   975
         Width           =   1065
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
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
         Height          =   675
         Left            =   2640
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   1425
         Width           =   1035
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "òœ «‘ —«ò"
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
         Left            =   2640
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   135
         Width           =   1065
      End
      Begin VB.Label Label1 
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
         Height          =   435
         Left            =   2640
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   495
         Width           =   1065
      End
   End
   Begin VSFlex7LCtl.VSFlexGrid vsCustomer 
      Height          =   4845
      Left            =   240
      TabIndex        =   37
      Top             =   3720
      Width           =   11715
      _cx             =   20664
      _cy             =   8546
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
      BackColor       =   -2147483644
      ForeColor       =   12582912
      BackColorFixed  =   16761024
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483633
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
      AllowUserResizing=   0
      SelectionMode   =   3
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   6
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   500
      RowHeightMax    =   500
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmFindCust2.frx":A4C2
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
   Begin RESIZEKITLibCtl.ResizeKit ResizeKit1 
      Height          =   480
      Left            =   0
      OleObjectBlob   =   "frmFindCust2.frx":A587
      TabIndex        =   52
      Top             =   0
      Width           =   480
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
      Left            =   7920
      RightToLeft     =   -1  'True
      TabIndex        =   51
      Top             =   2880
      Width           =   4095
   End
   Begin VB.Label LblCount 
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
      Left            =   2760
      RightToLeft     =   -1  'True
      TabIndex        =   50
      Top             =   8880
      Width           =   3015
   End
End
Attribute VB_Name = "frmFindCust2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim i, j As Long
Dim SearchType As Integer
Dim Rst As New ADODB.Recordset
Dim tmpflag As Boolean
Dim mvarbarcode As Boolean
Dim clsDate As New clsDate
Dim CountDailyBuy As Integer
Dim CountShiftBuy As Integer

Private Sub CancelButton_Click()
    mvarcode = 0
    txtMembershipId.Text = ""
    txtPhone.Text = ""
    TxtName.Text = ""
    TxtAddress.Text = ""
    CreditCode = 0
    Me.Hide
    ''''Unload Me
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call PresetScreenSaver
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call PresetScreenSaver
End Sub

''Private Sub CmdFamilySet_Click()
''If Val(txtMembershipId.Text) = 0 Or Val(vsCustomer.Rows) > 2 Then Exit Sub
''    ReDim Parameter(1) As Parameter
''        Parameter(0) = GenerateInputParameter("@FamilyNo", adInteger, 4, Val(TxtFamilyNo.Text))
''        Parameter(1) = GenerateInputParameter("@Code", adBigInt, 8, Val(txtMembershipId.Text))
''        Set Rst = RunParametricStoredProcedure2Rec("Update_Customer_FamilyNo", Parameter)
'''        LblResult.Visible = True
''        CmdFamilySet.Enabled = False
''        vsCustomer.SetFocus
''        tmpflag = True
''End Sub

Private Sub CmdNewCust_Click()
    
    If ClsFormAccess.frmCust = True Then
        FindCustFlag = True
        txtMembershipId.Text = ""
        txtPhone.Text = ""
        TxtName.Text = ""
        TxtAddress.Text = ""
        CreditCode = 0
        Me.Hide
        ''''Unload Me
        frmCust.Show
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
    Frame3.Visible = True
    Frame4.Visible = True
    Frame5.Visible = True
    
    If Val(txtMaxRecord.Text) < 1 Then txtMaxRecord.Text = "1"
    txtMaxRecord.Text = clsStation.MaxRecordCount
    If Val(TxtTimer.Text) < 1 Then TxtTimer.Text = "100"
    TxtTimer.Text = clsStation.SrarchInputDelayKeyboard
    
    Option1(0).Value = clsStation.CustomerSearchDefault
    Option1(1).Value = Not (clsStation.CustomerSearchDefault)
    
''    Select Case clsStation.DefaultCustSearch
''        Case EnumDefaultCustSearch.MembershipId
''            txtMembershipId.SetFocus
''            LblFindCust.Caption = "ﬂœ «‘ —«ﬂ —« Ê«—œ ﬂ‰Ìœ  "
''        Case EnumDefaultCustSearch.address
''            txtAddress.SetFocus
''            LblFindCust.Caption = "¬œ—” „‘ —ﬂ —« Ê«—œ ﬂ‰Ìœ  "
''        Case EnumDefaultCustSearch.Name
''            txtName.SetFocus
''            LblFindCust.Caption = "‰«„ „‘ —ﬂ —« Ê«—œ ﬂ‰Ìœ  "
''        Case EnumDefaultCustSearch.Phone
''            txtPhone.SetFocus
''            LblFindCust.Caption = " ·›‰ „‘ —ﬂ —« Ê«—œ ﬂ‰Ìœ  "
''    End Select
''    If Not (Val(strCategory) >= 0 And Val(strCategory) <= 7) Then
''       OptTable.Visible = False
''       OptSaloon.Caption = "›—Ê‘ê«Â"
''    End If
    
''    If clsStation.CustomerOrderDefault = True Then
''       optInPerson = True
''    Else
''       optByPhone = True
''    End If
''    If clsStation.CustomerServeplace = 0 Then
''       OptDelivery = True
''    ElseIf clsStation.CustomerServeplace = 1 Then
''       OptSaloon = True
''    ElseIf clsStation.CustomerServeplace = 2 Then
''       OptTable = True
''    End If
    
    tmpSearch = clsStation.CustomerSearchDefault    ' For Speed Search
    clsStation.CustomerSearchDefault = True
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
            clsStation.CustomerSearchDefault = tmpSearch
            txtMembershipId.Text = ""
            txtPhone.Text = ""
            TxtName.Text = ""
            TxtAddress.Text = ""
            Me.Hide
            ''''Unload Me
        End If
    Else        'Not Credit Card
        clsStation.CustomerSearchDefault = tmpSearch
''''        If clsStation.CustomerSearchDefault = False Then
''''            FillvsCustomer
''''        End If
    End If
    If Len(Call_RealNumber) > 0 Then
        txtPhone.SetFocus
        txtPhone.Text = Call_RealNumber
    Else        'Not Caller Id
        clsStation.CustomerSearchDefault = tmpSearch
''''        If clsStation.CustomerSearchDefault = False Then
''''            FillvsCustomer
''''        End If
    End If
''    CreditCode = 0
''
''    Set Rst = RunStoredProcedure2RecordSet("Get_All_Customers_Count")
''    LblMaxCustomerNo.Caption = "ﬂ· : " & Rst!MaxCustomerNo
''    LblMaxCustomerActive.Caption = "›⁄«·:" & Rst!MaxCustomerActive
''    LblMaxCustomerInActive.Caption = "»«ÿ· :" & Rst!MaxCustomerInActive
''
    CancelButton.Visible = True
    OKButton.Visible = True
'    Frame3.Visible = True
''    If VarActForm = "frmCust" Then
''      CmdNewCust.Visible = False
''    Else
''        CmdNewCust.Visible = True
''    End If
''
''    mvarbarcode = False

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
'    TxtFamilyNo.Text = ""
'    LblResult.Visible = False
    If strCategory <> "07" Then
'        Label10.Visible = False
'        TxtFamilyNo.Visible = False
'        CmdFamilySet.Visible = False
'        LblLastBuy.Visible = False
    End If
'    LblLastBuy.Caption = ""
'    txtBarcode.Text = ""
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



End Sub



Private Sub Option1_Click(index As Integer)
    If Option1(0).Value = False Then
        
'        Frame4.Visible = False
'        Frame5.Visible = False
        
        FillvsCustomer
    
    Else
'        Frame4.Visible = True
'        Frame5.Visible = True
        vsCustomer.Rows = 1
        labelClear

    End If
    txtMembershipId.SetFocus
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    txtMembershipId.Text = ""
    txtPhone.Text = ""
    TxtName.Text = ""
    TxtAddress.Text = ""
    CreditCode = 0
    
    If PosConnection.State = adStateOpen Then PosConnection.Close: Set PosConnection = Nothing
    Set Rst = Nothing

    SaveSetting strMainKey, Me.Name, "Left", Me.Left
    SaveSetting strMainKey, Me.Name, "Top", Me.Top



End Sub

Private Sub OKButton_Click()
 
    If vsCustomer.Row > 0 Then
        If frmSms.txtSMSDest.Text <> "" Then
           frmSms.txtSMSDest.Text = frmSms.txtSMSDest.Text & ";" & "0" & Val(vsCustomer.TextMatrix(vsCustomer.Row, 4))
        Else
            frmSms.txtSMSDest.Text = "0" & Val(vsCustomer.TextMatrix(vsCustomer.Row, 4))
'           mvarName = vsCustomer.TextMatrix(vsCustomer.Row, 2)
'           mvarBarcodeName = vsCustomer.TextMatrix(vsGoods.Row, 3)
        End If
    Else
        mvarcode = 0
    End If
'    txtSMSDest.Text = ""
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
                 If Len(TxtName.Text) > 0 Then
                   LblFindCust.Caption = " ‰«„ ( " & TxtName.Text & " ) œ— ”Ì” „  ⁄—Ì› ‰‘œÂ"
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
    TxtName.Text = ""
    txtMembershipId.Text = ""
    vsCustomer.Row = 0
    labelClear
    LblCount = ""
    vsCustomer.Select vsCustomer.Row, 5
    vsCustomer.Sort = flexSortGenericAscending
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
''
''Private Sub txtBarcode_Change()
''    If Right(txtBarcode.Text, 1) = "/" Then
''        txtBarcode.Text = Left(txtBarcode.Text, Len(txtBarcode.Text) - 1)
''    ElseIf Left(txtBarcode.Text, 1) = "/" Then
''        txtBarcode.Text = Right(txtBarcode.Text, Len(txtBarcode.Text) - 1)
''    End If
''
''End Sub

Private Sub txtMaxRecord_Change()
    If Val(txtMaxRecord.Text) < 1 Then txtMaxRecord.Text = "1"
    clsStation.MaxRecordCount = Val(txtMaxRecord.Text)
    SetStationSettingFile
End Sub

Private Sub txtMembershipId_Change()
'    CmdFamilySet.Enabled = False
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
    TxtName.Text = ""
    vsCustomer.Row = 0
    labelClear
    LblCount = ""
    vsCustomer.Select vsCustomer.Row, 2
    vsCustomer.Sort = flexSortGenericAscending
    LblFindCust.Caption = "ﬂœ «‘ —«ﬂ —« Ê«—œ ﬂ‰Ìœ  "
'    CmdFamilySet.Enabled = False

End Sub
Private Sub txtName_Change()

    If Option1(0).Value = False Then
        i = vsCustomer.FindRow(TxtName.Text, 1, 3, False, False)
        If i > 0 Then
            vsCustomer.Row = i
            vsCustomer.ShowCell i, 0
            LblFindCust.Caption = ""
        Else
            vsCustomer.Row = 0
            vsCustomer.ShowCell 0, 0
            labelClear
            If TxtName.Text <> "" Then
               LblFindCust.Caption = " ‰«„ ( " & TxtName.Text & " ) œ— ”Ì” „  ⁄—Ì› ‰‘œÂ"
             Else
                LblFindCust.Caption = "‰«„ „‘ —ﬂ —« Ê«—œ ﬂ‰Ìœ  "
             End If
    
        End If
    Else
       If Len(TxtName.Text) > 0 Then
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
    vsCustomer.Row = 0
    labelClear
    LblCount = ""
    vsCustomer.Select vsCustomer.Row, 3
    vsCustomer.Sort = flexSortGenericAscending
    LblFindCust.Caption = "‰«„ „‘ —ﬂ —« Ê«—œ ﬂ‰Ìœ  "

End Sub

Private Sub FillvsCustomer()
    Dim TmpTel As String
    Dim jj As Integer
    
    ReDim Parameter(1) As Parameter
    If VarActForm = "frmCust" Then
        Parameter(0) = GenerateInputParameter("@ActDeact", adInteger, 4, 2) ' All Customers
    Else
        Parameter(0) = GenerateInputParameter("@ActDeact", adInteger, 4, 0) ' Only Active
    End If
    Parameter(1) = GenerateInputParameter("@Branch", adInteger, 4, CurrentBranch)
    Set Rst = RunParametricStoredProcedure2Rec("Get_All_Customers", Parameter)
    With vsCustomer
        .ColHidden(1) = True
        .Rows = 1
        i = 0
        j = 0
''        FWProgressBar1.Value = 0
        MousePointer = 11
        While Rst.EOF <> True
            
            TmpTel = Rst!Tel1
            j = j + 1
            For jj = 1 To 5
                If TmpTel <> "" Or (jj = 1) Then
                    i = i + 1
                    .Rows = .Rows + 1
                    .TextMatrix(i, 0) = j
                    .TextMatrix(i, 1) = Rst!Code
                    .TextMatrix(i, 2) = Rst!MembershipId
                    .TextMatrix(i, 3) = Rst![Name]
                    .TextMatrix(i, 4) = Rst!Mobile
                    .TextMatrix(i, 5) = IIf(IsNull(Rst!address), " ", Rst!address)
                    TmpTel = ""
                End If
                If jj = 1 And Trim(Rst!Tel2) <> "" Then
                    TmpTel = Rst!Tel2
                ElseIf jj = 2 And Trim(Rst!Tel3) <> "" Then
                    TmpTel = Rst!Tel3
                ElseIf jj = 3 And Trim(Rst!Tel4) <> "" Then
                    TmpTel = Rst!Tel4
                End If
            Next jj
'''            CustomerSellPrice = Rst!SellPrice
            If i Mod 1000 = 0 Then DoEvents

            Rst.MoveNext
''            If i Mod 100 = 0 Then
''                FWProgressBar1.Value = FWProgressBar1 + 1
''                If FWProgressBar1.Value = 100 Then
''                    FWProgressBar1.Value = 1
''                End If
''            End If
            LblCount.Caption = " ⁄œ«œ —ﬂÊ—œÂ«   -  " & j
        Wend
        MousePointer = 0
    End With
    Set Rst = Nothing
    LblCount.Caption = " ⁄œ«œ —ﬂÊ—œÂ«   -  " & j
    vsCustomer.MergeCompare = flexMCTrimNoCase
    vsCustomer.MergeCells = flexMergeRestrictRows
    vsCustomer.MergeRow(vsCustomer.Rows - 1) = True
    vsCustomer.MergeCol(0) = True
    vsCustomer.MergeCol(1) = True
    vsCustomer.MergeCol(2) = True
    vsCustomer.MergeCol(3) = True
    vsCustomer.ColWidth(3) = vsCustomer.ColWidth(3) * 1.1
    vsCustomer.AutoSizeMode = flexAutoSizeColWidth
    vsCustomer.AutoSize 0, vsCustomer.Cols - 1
    If vsCustomer.ColWidth(2) < 800 Then
        vsCustomer.ColWidth(2) = 800
    End If
    If vsCustomer.ColWidth(3) < 3000 Then
        vsCustomer.ColWidth(3) = 3000
    End If
    If vsCustomer.ColWidth(4) < 1500 Then
        vsCustomer.ColWidth(4) = 1500
    End If
    If vsCustomer.ColWidth(5) < 4000 Then
        vsCustomer.ColWidth(5) = 4000
    End If
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

End Sub

Private Sub txtPhone_GotFocus()
    If tmpflag = True Then
        tmpflag = False
        Exit Sub
    End If
    TxtName.Text = ""
    TxtAddress.Text = ""
    txtMembershipId.Text = ""
    vsCustomer.Row = 0
    labelClear
    LblCount = ""
    vsCustomer.Select vsCustomer.Row, 4
    vsCustomer.Sort = flexSortGenericAscending
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
        If i = vsCustomer.Rows - 1 Then
            vsCustomer.TextMatrix(i, 0) = j
            Exit For
        End If
        vsCustomer.TextMatrix(i, 0) = j
        If vsCustomer.TextMatrix(i, 3) <> vsCustomer.TextMatrix(i + 1, 3) Then
            j = j + 1
        End If
    Next
    vsCustomer.MergeCells = flexMergeRestrictRows
    vsCustomer.MergeRow(vsCustomer.Rows - 1) = True
    vsCustomer.MergeCol(0) = True
    vsCustomer.MergeCol(0) = True
    
End Sub
Private Sub vsCustomer_DblClick()
    
    If OKButton.Visible = False Then Exit Sub
    If vsCustomer.Row > 0 Then
        OKButton_Click
    End If
End Sub

Private Sub Define_Customer()
    
'    TxtFamilyNo.Text = ""
'    LblLastBuy.Caption = "¬Œ—Ì‰ Œ—Ìœ"
    labelClear
    
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
            Parameter(1) = GenerateInputParameter("@Name", adVarWChar, 50, Left(TxtName.Text, 50))
            Set Rst = RunParametricStoredProcedure2Rec("Get_Customer_Name", Parameter)
        Case 3
            Parameter(1) = GenerateInputParameter("@Tel", adVarWChar, 20, Left(txtPhone.Text, 20))
            Set Rst = RunParametricStoredProcedure2Rec("Get_Customer_Tel", Parameter)
        Case 4
            Parameter(1) = GenerateInputParameter("@Addresse", adVarWChar, 100, Left(TxtAddress.Text, 100))
            Set Rst = RunParametricStoredProcedure2Rec("Get_Customer_Address", Parameter)
    End Select
'    CmdFamilySet.Enabled = False
    Dim TmpTel As String
    Dim jj As Integer
    
    With vsCustomer
        .Rows = 1
        i = 0
        j = 0
        Do While Rst.EOF <> True
'            If SearchType = 1 Then
'                TxtFamilyNo.Text = IIf(IsNull(Rst!FamilyNo), "", Rst!FamilyNo)
'                CmdFamilySet.Enabled = True
'            End If
            TmpTel = Rst!Tel1
            j = j + 1
            For jj = 1 To 5
                If TmpTel <> "" Or (jj = 1) Then
                    i = i + 1
                    .Rows = .Rows + 1
                    .TextMatrix(i, 0) = j
                    .TextMatrix(i, 1) = Rst!Code
                    .TextMatrix(i, 2) = Rst!MembershipId
                    .TextMatrix(i, 3) = Rst![Name]
                    .TextMatrix(i, 4) = Rst![Mobile]
                    .TextMatrix(i, 5) = IIf(IsNull(Rst!address), " ", Rst!address)
''''                    .TextMatrix(i, 6) = Rst!discount
''''                    .TextMatrix(i, 7) = Rst!Credit
''''                    .TextMatrix(i, 8) = Rst!carryfee
''''                    .TextMatrix(i, 9) = Rst!PaykFee
''''                    .TextMatrix(i, 10) = Rst!Distance
                    TmpTel = ""
                End If
                If jj = 1 And Trim(Rst!Tel2) <> "" Then
                    TmpTel = Rst!Tel2
                ElseIf jj = 2 And Trim(Rst!Tel3) <> "" Then
                    TmpTel = Rst!Tel3
                ElseIf jj = 3 And Trim(Rst!Tel4) <> "" Then
                    TmpTel = Rst!Tel4
                End If
            Next jj
             
''            ReDim Parameter(1) As Parameter
''            Parameter(0) = GenerateInputParameter("@Code", adBigInt, 8, Rst!Code)
''            Parameter(1) = GenerateOutputParameter("@Result", adBigInt, 8)
''            LblLastBuy.Caption = LblLastBuy.Caption & " " & vbLf & RunParametricStoredProcedure("Get_LastBuy", Parameter)
''
            
            
            If i > Val(txtMaxRecord.Text) Then Exit Do
            Rst.MoveNext
        Loop
        vsCustomer.MergeCompare = flexMCTrimNoCase
        vsCustomer.MergeCells = flexMergeRestrictRows
        vsCustomer.MergeRow(vsCustomer.Rows - 1) = True
        vsCustomer.MergeCol(0) = True
        vsCustomer.MergeCol(1) = True
        vsCustomer.MergeCol(2) = True
        vsCustomer.MergeCol(3) = True
    ''''    vsCustomer.MergeCol(4) = True
        vsCustomer.MergeCol(5) = True
        If i > 0 Then
            LblCount.Caption = " ⁄œ«œ —ﬂÊ—œÂ«   -  " & i
        Else
            LblCount.Caption = ""
        End If
        vsCustomer.AutoSizeMode = flexAutoSizeColWidth
        vsCustomer.AutoSize 0, .Cols - 1
    If vsCustomer.ColWidth(3) < 3000 Then
        vsCustomer.ColWidth(3) = 3000
    End If
    If vsCustomer.ColWidth(4) < 1500 Then
        vsCustomer.ColWidth(4) = 1500
    End If
    If vsCustomer.ColWidth(5) < 4000 Then
        vsCustomer.ColWidth(5) = 4000
    End If
    
     If vsCustomer.ColWidth(2) < 800 Then
        vsCustomer.ColWidth(2) = 800
    End If
    
    End With
    Set Rst = Nothing
'    LblResult.Visible = False
    vsCustomer_RowColChange
    
End Sub

Private Sub vsCustomer_RowColChange()
    If vsCustomer.Row = 0 Then Exit Sub
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
End Sub




