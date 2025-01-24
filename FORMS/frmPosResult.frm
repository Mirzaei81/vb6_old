VERSION 5.00
Begin VB.Form frmPosResult 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4935
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4935
   ScaleWidth      =   4560
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdEscape 
      Caption         =   "Œ—ÊÃ"
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
      TabIndex        =   13
      Top             =   4200
      Width           =   1575
   End
   Begin VB.Frame Frame0 
      Caption         =   "Å«”Œ œ—Ì«› Ì «“ ÅÊ“ »«‰òÌ"
      BeginProperty Font 
         Name            =   "B Nazanin"
         Size            =   12
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3855
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   4455
      Begin VB.TextBox txtReplyAmount 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   375
         Left            =   840
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   360
         Width           =   1695
      End
      Begin VB.TextBox txtReplySeq 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   375
         Left            =   840
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   936
         Width           =   1695
      End
      Begin VB.TextBox txtReplyCardNo 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   375
         Left            =   840
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   1512
         Width           =   1695
      End
      Begin VB.TextBox txtReplyDate 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   375
         Left            =   840
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   2088
         Width           =   1695
      End
      Begin VB.TextBox txtReplyRequestCode 
         Enabled         =   0   'False
         Height          =   375
         Left            =   840
         MaxLength       =   4
         TabIndex        =   2
         Top             =   2664
         Width           =   1695
      End
      Begin VB.TextBox txtReplyStatus 
         Enabled         =   0   'False
         Height          =   375
         Left            =   840
         MaxLength       =   4
         TabIndex        =   1
         Top             =   3240
         Width           =   1695
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "„»·€"
         BeginProperty Font 
            Name            =   "B Nazanin"
            Size            =   12
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   1
         Left            =   3000
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "‘„«—Â ÅÌêÌ—Ì"
         BeginProperty Font 
            Name            =   "B Nazanin"
            Size            =   12
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2520
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   1032
         Width           =   1095
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "‘„«—Â ò«— "
         BeginProperty Font 
            Name            =   "B Nazanin"
            Size            =   12
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2640
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   1584
         Width           =   975
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   " «—ÌŒ  —«ò‰‘"
         BeginProperty Font 
            Name            =   "B Nazanin"
            Size            =   12
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2640
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   2136
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "òœ œ—ŒÊ«” "
         BeginProperty Font 
            Name            =   "B Nazanin"
            Size            =   12
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   2640
         TabIndex        =   8
         Top             =   2688
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "Ê÷⁄Ì "
         BeginProperty Font 
            Name            =   "B Nazanin"
            Size            =   12
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   4
         Left            =   3000
         TabIndex        =   7
         Top             =   3240
         Width           =   735
      End
   End
End
Attribute VB_Name = "frmPosResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
