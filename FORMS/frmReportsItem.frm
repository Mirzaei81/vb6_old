VERSION 5.00
Object = "{7AEDC602-D94C-11D1-BB7A-00E0290EA3C9}#1.0#0"; "ResizeKit.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{687FF23F-0E9B-449D-A782-2C5E7116EDB0}#1.0#0"; "Fardate.ocx"
Begin VB.Form frmReportsItem 
   Caption         =   "                                                                                      ê“«—‘«     "
   ClientHeight    =   6525
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12375
   Icon            =   "frmReportsItem.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6525
   ScaleWidth      =   12375
   Begin VB.Frame Frame1 
      Caption         =   "Å«—«„ —Â«Ì ê“«—‘« "
      BeginProperty Font 
         Name            =   "Nazanin"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   4695
      Left            =   240
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   660
      Width           =   5895
      Begin FarDate1.FarDate FarDate2 
         Height          =   345
         Left            =   120
         TabIndex        =   82
         Top             =   360
         Visible         =   0   'False
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   609
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Object.ToolTipText     =   "FarDate"
      End
      Begin FarDate1.FarDate FarDate1 
         Height          =   345
         Left            =   3000
         TabIndex        =   83
         Top             =   360
         Visible         =   0   'False
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   609
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Object.ToolTipText     =   "FarDate"
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   17
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   45
         Top             =   4200
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   16
         Left            =   3000
         RightToLeft     =   -1  'True
         TabIndex        =   44
         Top             =   4200
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   2
         Left            =   3000
         RightToLeft     =   -1  'True
         TabIndex        =   41
         Top             =   840
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   3
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   40
         Top             =   840
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   4
         Left            =   3000
         RightToLeft     =   -1  'True
         TabIndex        =   39
         Top             =   1320
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   5
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   38
         Top             =   1320
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   6
         Left            =   3000
         RightToLeft     =   -1  'True
         TabIndex        =   37
         Top             =   1800
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   7
         Left            =   90
         RightToLeft     =   -1  'True
         TabIndex        =   36
         Top             =   1770
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   8
         Left            =   3000
         RightToLeft     =   -1  'True
         TabIndex        =   35
         Top             =   2280
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   9
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   34
         Top             =   2280
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   10
         Left            =   3000
         RightToLeft     =   -1  'True
         TabIndex        =   33
         Top             =   2760
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   11
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   32
         Top             =   2760
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   12
         Left            =   3000
         RightToLeft     =   -1  'True
         TabIndex        =   31
         Top             =   3240
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   13
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   30
         Top             =   3240
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   14
         Left            =   3000
         RightToLeft     =   -1  'True
         TabIndex        =   29
         Top             =   3720
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   15
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   28
         Top             =   3720
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   2
         Left            =   3000
         RightToLeft     =   -1  'True
         TabIndex        =   25
         Top             =   840
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   3
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   24
         Top             =   840
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   4
         Left            =   3000
         RightToLeft     =   -1  'True
         TabIndex        =   23
         Top             =   1320
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   5
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   22
         Top             =   1320
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   6
         Left            =   3000
         RightToLeft     =   -1  'True
         TabIndex        =   21
         Top             =   1800
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   7
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Top             =   1800
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   8
         Left            =   3000
         RightToLeft     =   -1  'True
         TabIndex        =   19
         Top             =   2280
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   9
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   2280
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   10
         Left            =   3000
         RightToLeft     =   -1  'True
         TabIndex        =   17
         Top             =   2760
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   11
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Top             =   2760
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   12
         Left            =   3000
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   3240
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   13
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   3240
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   14
         Left            =   3000
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   3720
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   15
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   3720
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   16
         Left            =   3000
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   4200
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   17
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   4200
         Visible         =   0   'False
         Width           =   1695
      End
      Begin MSMask.MaskEdBox MaskEdBox1 
         Height          =   360
         Index           =   0
         Left            =   3000
         TabIndex        =   46
         Top             =   375
         Visible         =   0   'False
         Width           =   1680
         _ExtentX        =   2963
         _ExtentY        =   635
         _Version        =   393216
         MaxLength       =   5
         Mask            =   "##:##"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox MaskEdBox1 
         Height          =   360
         Index           =   1
         Left            =   135
         TabIndex        =   47
         Top             =   345
         Visible         =   0   'False
         Width           =   1680
         _ExtentX        =   2963
         _ExtentY        =   635
         _Version        =   393216
         MaxLength       =   5
         Mask            =   "##:##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox MaskEdBox1 
         Height          =   360
         Index           =   2
         Left            =   3015
         TabIndex        =   48
         Top             =   825
         Visible         =   0   'False
         Width           =   1680
         _ExtentX        =   2963
         _ExtentY        =   635
         _Version        =   393216
         MaxLength       =   5
         Mask            =   "##:##"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox MaskEdBox1 
         Height          =   360
         Index           =   3
         Left            =   120
         TabIndex        =   49
         Top             =   825
         Visible         =   0   'False
         Width           =   1680
         _ExtentX        =   2963
         _ExtentY        =   635
         _Version        =   393216
         MaxLength       =   5
         Mask            =   "##:##"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox MaskEdBox1 
         Height          =   360
         Index           =   4
         Left            =   3000
         TabIndex        =   50
         Top             =   1320
         Visible         =   0   'False
         Width           =   1680
         _ExtentX        =   2963
         _ExtentY        =   635
         _Version        =   393216
         MaxLength       =   5
         Mask            =   "##:##"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox MaskEdBox1 
         Height          =   360
         Index           =   5
         Left            =   120
         TabIndex        =   51
         Top             =   1305
         Visible         =   0   'False
         Width           =   1680
         _ExtentX        =   2963
         _ExtentY        =   635
         _Version        =   393216
         MaxLength       =   5
         Mask            =   "##:##"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox MaskEdBox1 
         Height          =   360
         Index           =   6
         Left            =   3000
         TabIndex        =   52
         Top             =   1785
         Visible         =   0   'False
         Width           =   1680
         _ExtentX        =   2963
         _ExtentY        =   635
         _Version        =   393216
         MaxLength       =   5
         Mask            =   "##:##"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox MaskEdBox1 
         Height          =   360
         Index           =   7
         Left            =   120
         TabIndex        =   53
         Top             =   1770
         Visible         =   0   'False
         Width           =   1680
         _ExtentX        =   2963
         _ExtentY        =   635
         _Version        =   393216
         MaxLength       =   5
         Mask            =   "##:##"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox MaskEdBox1 
         Height          =   360
         Index           =   8
         Left            =   2985
         TabIndex        =   54
         Top             =   2280
         Visible         =   0   'False
         Width           =   1680
         _ExtentX        =   2963
         _ExtentY        =   635
         _Version        =   393216
         MaxLength       =   5
         Mask            =   "##:##"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox MaskEdBox1 
         Height          =   360
         Index           =   9
         Left            =   105
         TabIndex        =   55
         Top             =   2280
         Visible         =   0   'False
         Width           =   1680
         _ExtentX        =   2963
         _ExtentY        =   635
         _Version        =   393216
         MaxLength       =   5
         Mask            =   "##:##"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox MaskEdBox1 
         Height          =   360
         Index           =   10
         Left            =   3015
         TabIndex        =   56
         Top             =   2760
         Visible         =   0   'False
         Width           =   1680
         _ExtentX        =   2963
         _ExtentY        =   635
         _Version        =   393216
         MaxLength       =   5
         Mask            =   "##:##"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox MaskEdBox1 
         Height          =   360
         Index           =   11
         Left            =   120
         TabIndex        =   57
         Top             =   2760
         Visible         =   0   'False
         Width           =   1680
         _ExtentX        =   2963
         _ExtentY        =   635
         _Version        =   393216
         MaxLength       =   5
         Mask            =   "##:##"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox MaskEdBox1 
         Height          =   360
         Index           =   12
         Left            =   3000
         TabIndex        =   58
         Top             =   3240
         Visible         =   0   'False
         Width           =   1680
         _ExtentX        =   2963
         _ExtentY        =   635
         _Version        =   393216
         MaxLength       =   5
         Mask            =   "##:##"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox MaskEdBox1 
         Height          =   360
         Index           =   13
         Left            =   120
         TabIndex        =   59
         Top             =   3240
         Visible         =   0   'False
         Width           =   1680
         _ExtentX        =   2963
         _ExtentY        =   635
         _Version        =   393216
         MaxLength       =   5
         Mask            =   "##:##"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox MaskEdBox1 
         Height          =   360
         Index           =   14
         Left            =   3000
         TabIndex        =   60
         Top             =   3720
         Visible         =   0   'False
         Width           =   1680
         _ExtentX        =   2963
         _ExtentY        =   635
         _Version        =   393216
         MaxLength       =   5
         Mask            =   "##:##"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox MaskEdBox1 
         Height          =   360
         Index           =   15
         Left            =   135
         TabIndex        =   61
         Top             =   3720
         Visible         =   0   'False
         Width           =   1680
         _ExtentX        =   2963
         _ExtentY        =   635
         _Version        =   393216
         MaxLength       =   5
         Mask            =   "##:##"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox MaskEdBox1 
         Height          =   360
         Index           =   16
         Left            =   3015
         TabIndex        =   62
         Top             =   4200
         Visible         =   0   'False
         Width           =   1680
         _ExtentX        =   2963
         _ExtentY        =   635
         _Version        =   393216
         MaxLength       =   5
         Mask            =   "##:##"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox MaskEdBox1 
         Height          =   360
         Index           =   17
         Left            =   135
         TabIndex        =   63
         Top             =   4200
         Visible         =   0   'False
         Width           =   1680
         _ExtentX        =   2963
         _ExtentY        =   635
         _Version        =   393216
         MaxLength       =   5
         Mask            =   "##:##"
         PromptChar      =   " "
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   0
         Left            =   3015
         RightToLeft     =   -1  'True
         TabIndex        =   43
         Top             =   360
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox Text1 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   1
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   42
         Top             =   330
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   1
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   26
         Top             =   360
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Index           =   0
         Left            =   3000
         RightToLeft     =   -1  'True
         TabIndex        =   27
         Top             =   360
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label LblIo 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
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
         Index           =   17
         Left            =   1800
         RightToLeft     =   -1  'True
         TabIndex        =   81
         Top             =   4200
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label LblIo 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
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
         Index           =   16
         Left            =   4680
         RightToLeft     =   -1  'True
         TabIndex        =   80
         Top             =   4200
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label LblIo 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
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
         Index           =   15
         Left            =   1800
         RightToLeft     =   -1  'True
         TabIndex        =   79
         Top             =   3720
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label LblIo 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
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
         Index           =   14
         Left            =   4680
         RightToLeft     =   -1  'True
         TabIndex        =   78
         Top             =   3720
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label LblIo 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
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
         Index           =   13
         Left            =   1800
         RightToLeft     =   -1  'True
         TabIndex        =   77
         Top             =   3240
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label LblIo 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
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
         Index           =   12
         Left            =   4680
         RightToLeft     =   -1  'True
         TabIndex        =   76
         Top             =   3240
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label LblIo 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
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
         Index           =   11
         Left            =   1800
         RightToLeft     =   -1  'True
         TabIndex        =   75
         Top             =   2760
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label LblIo 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
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
         Index           =   10
         Left            =   4680
         RightToLeft     =   -1  'True
         TabIndex        =   74
         Top             =   2760
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label LblIo 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
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
         Index           =   9
         Left            =   1800
         RightToLeft     =   -1  'True
         TabIndex        =   73
         Top             =   2280
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label LblIo 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
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
         Index           =   8
         Left            =   4680
         RightToLeft     =   -1  'True
         TabIndex        =   72
         Top             =   2280
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label LblIo 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
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
         Index           =   7
         Left            =   1800
         RightToLeft     =   -1  'True
         TabIndex        =   71
         Top             =   1800
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label LblIo 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
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
         Index           =   6
         Left            =   4680
         RightToLeft     =   -1  'True
         TabIndex        =   70
         Top             =   1800
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label LblIo 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
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
         Index           =   5
         Left            =   1800
         RightToLeft     =   -1  'True
         TabIndex        =   69
         Top             =   1320
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label LblIo 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
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
         Index           =   4
         Left            =   4680
         RightToLeft     =   -1  'True
         TabIndex        =   68
         Top             =   1320
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label LblIo 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
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
         Index           =   3
         Left            =   1800
         RightToLeft     =   -1  'True
         TabIndex        =   67
         Top             =   840
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label LblIo 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
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
         Index           =   2
         Left            =   4680
         RightToLeft     =   -1  'True
         TabIndex        =   66
         Top             =   840
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label LblIo 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
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
         Index           =   1
         Left            =   1800
         RightToLeft     =   -1  'True
         TabIndex        =   65
         Top             =   360
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label LblIo 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
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
         Index           =   0
         Left            =   4680
         RightToLeft     =   -1  'True
         TabIndex        =   64
         Top             =   360
         Visible         =   0   'False
         Width           =   1095
      End
   End
   Begin VB.PictureBox Picture1 
      BeginProperty Font 
         Name            =   "Nazanin"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   915
      Left            =   240
      RightToLeft     =   -1  'True
      ScaleHeight     =   855
      ScaleWidth      =   5865
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   5520
      Width           =   5925
      Begin VB.CommandButton cmd_Esc 
         BackColor       =   &H000040C0&
         Cancel          =   -1  'True
         Caption         =   "Œ—ÊÃ"
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
         Left            =   120
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   120
         Width           =   1365
      End
      Begin VB.CommandButton cmd_Ok 
         BackColor       =   &H0000C000&
         Caption         =   "ﬁ»Ê·"
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
         Left            =   1680
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   120
         Width           =   1365
      End
   End
   Begin VB.ListBox lstItemReports 
      BeginProperty Font 
         Name            =   "Nazanin"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3960
      ItemData        =   "frmReportsItem.frx":A4C2
      Left            =   6240
      List            =   "frmReportsItem.frx":A4C4
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   600
      Width           =   2985
   End
   Begin VB.ListBox lstGroupReports 
      BeginProperty Font 
         Name            =   "Nazanin"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3960
      ItemData        =   "frmReportsItem.frx":A4C6
      Left            =   9360
      List            =   "frmReportsItem.frx":A4C8
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   600
      Width           =   2895
   End
   Begin RESIZEKITLibCtl.ResizeKit ResizeKit1 
      Height          =   480
      Left            =   0
      OleObjectBlob   =   "frmReportsItem.frx":A4CA
      TabIndex        =   4
      Top             =   0
      Width           =   480
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   600
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowGroupTree=   -1  'True
      WindowAllowDrillDown=   -1  'True
      WindowShowCloseBtn=   -1  'True
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin VB.Label LblIo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Nazanin"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   495
      Index           =   18
      Left            =   240
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   120
      Width           =   5895
   End
   Begin VB.Label lblGoodLevel2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H000080FF&
      BackStyle       =   0  'Transparent
      Caption         =   "‰«„ ê“«—‘« "
      BeginProperty Font 
         Name            =   "Titr"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   435
      Left            =   7200
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   120
      Width           =   2025
   End
   Begin VB.Label lblGoodLevel1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H000040C0&
      BackStyle       =   0  'Transparent
      Caption         =   "ê—ÊÂ ê“«—‘« "
      BeginProperty Font 
         Name            =   "Titr"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   435
      Left            =   9720
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   120
      Width           =   2415
   End
End
Attribute VB_Name = "frmReportsItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private MyFormAddEditMode As EnumAddEditMode
Dim Rst As New ADODB.Recordset
Dim Parameter() As Parameter
Dim ReportFileName As String
Dim ParameterData(0 To 17) As Variant
Dim ParameterName(0 To 17) As String
Dim ParameterType(0 To 17) As String
Dim parameterLengh(0 To 17) As Long
Dim ParameterName2(0 To 17) As String
Dim MinValue(0 To 17) As String
Dim MaxValue(0 To 17) As String
Private clsDate As New clsDate
Private keyTXT As New Collection
Dim ii As Integer
Public Sub ExitForm()
    Unload Me
End Sub

Private Sub cmd_Esc_Click()
    ExitForm
End Sub
Private Sub Form_Activate()
    
    VarActForm = Me.Name
    
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call PresetScreenSaver
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call PresetScreenSaver
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case Shift
          Case 0
              Select Case KeyCode
                  Case 13  ' Enter
                    SendKeys "{Left}", True
                  Case 27  ' Esc
                    Me.ExitForm
              End Select
          Case 2
               Select Case KeyCode
                  Case 123  'Exit
              End Select

    End Select

End Sub
Private Sub Text1_KeyDown(index As Integer, KeyCode As Integer, Shift As Integer)
Dim i As Integer
Dim Pass As Boolean
For i = 0 To keyTXT.Count - 1
    If index = keyTXT.Item(i + 1) Then
        Pass = True
    End If
Next i
If Pass Then
    If KeyCode = vbKeyDelete Then
        If Mid(Text1(index).Text, Text1(index).SelStart + 1, 1) = "/" Then
            KeyCode = 0
        End If
    End If
End If
End Sub

Private Sub Text1_KeyPress(index As Integer, KeyAscii As Integer)
On Error Resume Next
Dim n As Integer
Dim i As Integer
Dim Pass As Boolean
For i = 0 To keyTXT.Count - 1
    If index = keyTXT.Item(i + 1) Then
        Pass = True
    End If
Next i
If Pass Then
    n = Len(Text1(index).Text)
    If KeyAscii = 8 Then
    If n = Text1(index).SelStart Then Exit Sub
    If Mid(Text1(index).Text, Text1(index).SelStart, 1) = "/" Then GoTo trap
    Exit Sub: End If
    If KeyAscii < 48 Or KeyAscii > 57 Then GoTo trap
    If n <> Text1(index).SelStart Then Exit Sub
    Select Case n
    Case 2
    Text1(index).Text = Text1(index).Text & "/"
    Text1(index).SelStart = n + 1
    Case 5
    Text1(index).Text = Text1(index).Text & "/"
    Text1(index).SelStart = n + 1
    End Select
End If
Exit Sub
trap:
KeyAscii = 0
Exit Sub
End Sub

Private Sub Form_Load()

'    If clsFormAccess.frm = False Then
'        Unload Me
'        Exit Sub
'    End If
     
     
        
    CenterTop Me
    VarActForm = Me.Name
    
    DefaultSetting
    
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
    
    FarDate1.Visible = False
    FarDate2.Visible = False
    If clsArya.MiladiDate = 0 Then FarDate1.Text = "13" + mvarDate
    If clsArya.MiladiDate = 0 Then FarDate2.Text = clsDate.shamsi(Date) ' Mid(ClsDate.shamsi(Date), 3, 8)
'    FarDate1.Top = Me.UCReportIO1.txt(0).Top + 230
'    FarDate1.Height = Me.UCReportIO1.txt(0).Height + 50
'    FarDate1.Left = Me.UCReportIO1.txt(0).Left
'    FarDate2.Top = Me.UCReportIO1.txt(1).Top + 230
'    FarDate2.Height = Me.UCReportIO1.txt(1).Height + 50
'    FarDate2.Left = Me.UCReportIO1.txt(1).Left
 
End Sub
Private Sub Form_Unload(Cancel As Integer)
    
    VarActForm = ""
    If PosConnection.State = adStateOpen Then PosConnection.Close: Set PosConnection = Nothing

    Dim i As Integer
    
    AllButton vbOff, True
    
    SaveSetting strMainKey, Me.Name, "Left", Me.Left
    SaveSetting strMainKey, Me.Name, "Top", Me.Top
   
End Sub
Public Sub DefaultSetting()

    lstGroupReports.Clear
    lstItemReports.Clear
    FilllstGroupReports
End Sub

Public Sub FilllstGroupReports() ' it fills the lstGroupReports using table tgoodlevel1
    
    lstGroupReports.Clear
    lstItemReports.Clear
    
    Set Rst = RunStoredProcedure2RecordSet("Get_All_GroupReports")
        
    If (Rst.EOF = True And Rst.BOF = True) Then
        Exit Sub
    End If
    
    While Rst.EOF = False
        lstGroupReports.AddItem Rst.Fields("GroupReportName")
        lstGroupReports.ItemData(lstGroupReports.ListCount - 1) = Rst.Fields("intGroupreportId")
        Rst.MoveNext
    Wend
    lstGroupReports.ListIndex = 0
    FilllstItemReports
    Set Rst = Nothing
End Sub

Public Sub FilllstItemReports() ' it fills the lstItemReports using table tgoodlevel2

    lstItemReports.Clear
    ReportFileName = ""
    LblIo(18).Caption = ""
    
    If lstGroupReports.ListIndex = -1 Then
        Set Rst = Nothing
        Exit Sub
    Else
        ReDim Parameter(1) As Parameter
        Parameter(0) = GenerateInputParameter("@intGroupreportId", adInteger, 4, lstGroupReports.ItemData(lstGroupReports.ListIndex))
        Parameter(1) = GenerateInputParameter("@AccessLevel", adInteger, 4, mVarAccessLevel)
        
        If Rst.State <> 0 Then Rst.Close
        Set Rst = RunParametricStoredProcedure2Rec("Get_ItemReports_ByGroupId", Parameter)
        If (Rst.EOF = True And Rst.BOF = True) Then
            Set Rst = Nothing
            Exit Sub
        End If
       ' rst.moveFirst
       Dim i As Long
        While Rst.EOF = False And (i < 5 Or intVersion <> Min)
            lstItemReports.AddItem Rst.Fields("ReportName")
            lstItemReports.ItemData(lstItemReports.ListCount - 1) = Rst.Fields("intReportId")
            i = i + 1
            Rst.MoveNext
        Wend
        
        Set Rst = Nothing
  '      lstItemReports.ListIndex = 0
        
    End If
    
End Sub

Private Sub lstGroupReports_Click()
    FilllstItemReports
End Sub

Private Sub lstItemReports_Click()
    On Error GoTo ErrHandler
    
    If lstItemReports.ListIndex = -1 Then
        Exit Sub
    End If
    
    ReDim Parameter(0) As Parameter
    Parameter(0) = GenerateInputParameter("@intReportId", adInteger, 4, lstItemReports.ItemData(lstItemReports.ListIndex))
    
    Set Rst = RunParametricStoredProcedure2Rec("Get_ItemReports_ByReportId", Parameter)
    If Not (Rst.EOF = True And Rst.BOF = True) Then
        ReportFileName = Rst!latinReportName
        LblIo(18).Caption = Rst!ReportName
    End If

    If Trim(ReportFileName) = "" Then
        MsgBox "‰«„ ê“«—‘ „⁄·Ê„ ‰Ì” "
        Exit Sub
    End If
    
    ClearParameters
    FillParameters
    
    If lstItemReports.ItemData(lstItemReports.ListIndex) = 94 Or lstItemReports.ItemData(lstItemReports.ListIndex) = 95 Then
        If intVersion <> Diamond Then
            cmd_Ok.Enabled = False: ShowDisMessage "«Ì‰ ê“«—‘ ›ﬁÿ œ— Ê—é‰ «·„«” ÊÃÊœ œ«—œ", 1500
        Else
            cmd_Ok.Enabled = True
        End If
    Else
        cmd_Ok.Enabled = True
    End If
    

Exit Sub

ErrHandler:
    LogSave "frmReportsItem => ", err, "LstItemReports_Click()"
    ShowErrorMessage
    err.Clear
    Resume Next
End Sub
Private Sub ClearParameters()
    On Error GoTo ErrHandler

    For i = 0 To 17
        LblIo(i).Visible = False
        Text1(i).Visible = False
        Combo1(i).Visible = False
        MaskEdBox1(i).Visible = False
        LblIo(i).Caption = ""
        Text1(i).Text = ""
        Combo1(i).Clear
        MaskEdBox1(i).SelText = "00:00"
        LblIo(i).Tag = ""
        Text1(i).Tag = ""
        Combo1(i).Tag = ""
        MaskEdBox1(i).Tag = ""
        ParameterData(i) = ""
        ParameterName(i) = ""
        ParameterType(i) = ""
        parameterLengh(i) = 0
        ParameterName2(i) = ""
        MinValue(i) = ""
        MaxValue(i) = ""
    Next i
    
    FarDate1.Visible = False
    FarDate2.Visible = False
    If clsArya.MiladiDate = 0 Then FarDate1.Text = "13" + mvarDate
    If clsArya.MiladiDate = 0 Then FarDate2.Text = clsDate.shamsi(Date) ' Mid(ClsDate.shamsi(Date), 3, 8)

Exit Sub
ErrHandler:
    LogSave "frmReportsItem => ", err, "ClearParameters"
    ShowErrorMessage
    err.Clear
    Resume Next
End Sub
Private Sub FarDate1_Change()
    
    Text1(ii).Text = Mid(FarDate1.Text, 3, 8)
End Sub

Private Sub FarDate2_Change()
    Text1(ii + 1).Text = Mid(FarDate2.Text, 3, 8)
End Sub
Private Sub FillParameters()
    On Error GoTo ErrHandler
    Dim Rst As New ADODB.Recordset
    
    Dim i As Integer
    For i = 0 To keyTXT.Count - 1
        keyTXT.remove (i + 1)
    Next i
    
    ReDim Parameter(0) As Parameter
    Parameter(0) = GenerateInputParameter("@intReportId", adInteger, 4, lstItemReports.ItemData(lstItemReports.ListIndex))
    
    Set Rst = RunParametricStoredProcedure2Rec("Get_ItemReports_Details_ByReportId", Parameter)
    If Not (Rst.EOF = True And Rst.BOF = True) Then
        cmd_Ok.Enabled = False
        While Rst.EOF = False
            LblIo(Rst!Row * 2 - 2).Visible = True
            LblIo(Rst!Row * 2 - 2).Caption = Rst!FromText
            If Rst!Quantity = 2 Then
                LblIo(Rst!Row * 2 - 1).Visible = True
                LblIo(Rst!Row * 2 - 1).Caption = Rst!ToText
            End If
            
            If Rst!ObjectType = EnumObjectType.TextBox Then
                Text1(Rst!Row * 2 - 2).Visible = True
                Text1(Rst!Row * 2 - 2).Text = IIf(IsNull(Rst!MinValue), "", Rst!MinValue)
                Text1(Rst!Row * 2 - 2).RightToLeft = Rst!RightToLeft
             '   Text1(Rst!Row * 2 - 2).Tag = Rst!FromParameter
                MinValue(Rst!Row * 2 - 2) = IIf(IsNull(Rst!MinValue), "", Rst!MinValue)
                If InStr(1, Rst!ParameterName, "Date", vbTextCompare) Then
                    If clsArya.MiladiDate = 0 Then FarDate1.Visible = True
                    ii = Rst!Row * 2 - 2
                    Text1(Rst!Row * 2 - 2).Text = mvarDate
                    keyTXT.Add Rst!Row * 2 - 2
                End If
                ParameterName(Rst!Row * 2 - 2) = Rst!ParameterName
                ParameterType(Rst!Row * 2 - 2) = Rst!ParameterType
                parameterLengh(Rst!Row * 2 - 2) = Rst!parameterLengh
                If Rst!Quantity = 2 Then
                    Text1(Rst!Row * 2 - 1).Visible = True
                    Text1(Rst!Row * 2 - 1).Text = IIf(IsNull(Rst!MaxValue), "", Rst!MaxValue)
                    Text1(Rst!Row * 2 - 1).RightToLeft = Rst!RightToLeft
             '       Text1(Rst!Row * 2 - 1).Tag = Rst!ToParameter
                    MaxValue(Rst!Row * 2 - 1) = IIf(IsNull(Rst!MaxValue), "", Rst!MaxValue)
                    If InStr(1, Rst!ParameterName, "Date", vbTextCompare) Then
                        If clsArya.MiladiDate = 0 Then FarDate2.Visible = True
                        Text1(Rst!Row * 2 - 1).Text = Mid(clsDate.shamsi(Date), 3, 8)
                        keyTXT.Add Rst!Row * 2 - 1
                    End If
                    ParameterName(Rst!Row * 2 - 1) = Rst!ParameterName
                    ParameterType(Rst!Row * 2 - 1) = Rst!ParameterType
                    parameterLengh(Rst!Row * 2 - 1) = Rst!parameterLengh
                End If
           
            ElseIf Rst!ObjectType = EnumObjectType.ComboBox Then
                Combo1(Rst!Row * 2 - 2).Visible = True
               
                If Not IsNull(Rst!ComboQuery) Then
                    FillList Combo1(Rst!Row * 2 - 2), Rst!ComboQuery, Rst!ComboFieldCode, Rst!ComboFieldDescr
                Else
                    MsgBox "«ÿ·«⁄«  »—«Ì Å— ò—œ‰ Å«—«„ —Â«Ì —œÌ›  " & Rst!Row & "ò«„· ‰Ì”  "
                End If
                
                If Rst!ComboFieldCode = "AccountYear" Then
                    For i = o To Combo1(Rst!Row * 2 - 2).ListCount - 1
                        Combo1(Rst!Row * 2 - 2).ListIndex = i
                        If Combo1(Rst!Row * 2 - 2).Text = AccountYear Then Exit For
                        
                    Next i
                Else
                    Combo1(Rst!Row * 2 - 2).ListIndex = 0
                End If
           '     Combo1(Rst!Row * 2 - 2).RightToLeft = Rst!RightToLeft
            '    Combo1(Rst!Row * 2 - 2).Tag = Rst!FromParameter
                
                ParameterName(Rst!Row * 2 - 2) = Rst!ParameterName
                ParameterType(Rst!Row * 2 - 2) = Rst!ParameterType
                parameterLengh(Rst!Row * 2 - 2) = Rst!parameterLengh
                
                If Rst!Quantity = 2 Then
                    Combo1(Rst!Row * 2 - 1).Visible = True
                    FillList Combo1(Rst!Row * 2 - 1), Rst!ComboQuery, Rst!ComboFieldCode, Rst!ComboFieldDescr
                    Combo1(Rst!Row * 2 - 1).ListIndex = Combo1(Rst!Row * 2 - 1).ListCount - 1
            '        Combo1(Rst!Row * 2 - 1).RightToLeft = Rst!RightToLeft
            '        Combo1(Rst!Row * 2 - 1).Tag = Rst!ToParameter
                    ParameterName(Rst!Row * 2 - 1) = Rst!ParameterName
                    ParameterType(Rst!Row * 2 - 1) = Rst!ParameterType
                    parameterLengh(Rst!Row * 2 - 1) = Rst!parameterLengh
                End If
        
            ElseIf Rst!ObjectType = EnumObjectType.MaskEdBox Then
                MaskEdBox1(Rst!Row * 2 - 2).Visible = True
                
                  ParameterName(Rst!Row * 2 - 2) = IIf(IsNull(Rst!MinValue), "  :  ", Rst!MinValue)
                  ''MaskEdBox1(Rst!Row * 2 - 2).Text = IIf(IsNull(Rst!MinValue), "", Rst!MinValue)
           '     MaskEdBox1(Rst!Row * 2 - 2).Tag = Rst!FromParameter
                ParameterName(Rst!Row * 2 - 2) = Rst!ParameterName
                ParameterType(Rst!Row * 2 - 2) = Rst!ParameterType
                parameterLengh(Rst!Row * 2 - 2) = Rst!parameterLengh
                
                If Rst!Quantity = 2 Then
                    MaskEdBox1(Rst!Row * 2 - 1).Visible = True
                    MaskEdBox1(Rst!Row * 2 - 1).Text = Trim(IIf(IsNull(Rst!MaxValue), "", Rst!MaxValue))
             '       MaskEdBox1(Rst!Row * 2 - 1).Tag = Rst!ToParameter
                    ParameterName(Rst!Row * 2 - 1) = Rst!ParameterName
                    ParameterType(Rst!Row * 2 - 1) = Rst!ParameterType
                    parameterLengh(Rst!Row * 2 - 1) = Rst!parameterLengh
                End If
            End If
            Rst.MoveNext
        Wend
        cmd_Ok.Enabled = True
    Else
        cmd_Ok.Enabled = False
    End If
    Set Rst = Nothing
    If PosConnection.State = adStateOpen Then PosConnection.Close: Set PosConnection = Nothing

Exit Sub

ErrHandler:
    LogSave "frmReportsItem=> ", err, "FillParameters"
'    ShowErrorMessage
    err.Clear
    Resume Next
End Sub
Private Sub FillList(Obj As Object, Query As String, Code As String, Name As String)
Dim ii As Long
On Error GoTo ErrorHandler
Dim Rc As New ADODB.Recordset
        Set Rc = RunQuery2RecordSet(Query)
        Obj.Clear
        ii = 0
        If Rc.EOF = False Then
            Rc.MoveFirst
            Do While Not Rc.EOF()
                If ii = 50000 Then
                    MsgBox " «‘ò«· œ— Å— ò—œ‰ " & Obj.Name
                    Exit Sub
                End If
                If Not IsNull(Rc.Fields(Name)) Then
                    Obj.AddItem Rc.Fields(Name)
                    '              If Left(Rc.Fields(Code), 1) <> "0" Then
                        Obj.ItemData(ii) = Right(Rc.Fields(Code), 9)
                    '              Else
                    '                  obj.ItemData(ii) = "9999" & Rc.Fields(Code)
                    '              End If
                    ii = ii + 1
                End If
                Rc.MoveNext
            Loop
        End If
    If Rc.State = adStateOpen Then Rc.Close:    Set Rc = Nothing

Exit Sub

ErrorHandler:
    LogSave "frmReportsItem => ", err, "FillList"
    ShowErrorMessage
    err.Clear
    If Rc.State = adStateOpen Then Rc.Close:    Set Rc = Nothing
    If PosConnection.State = adStateOpen Then PosConnection.Close: Set PosConnection = Nothing
    
End Sub

Private Sub ResizeKit1_ExitResize(ByVal XScale As Double, ByVal YScale As Double)

    If formloadFlag = True Then
        SaveSetting strMainKey, Me.Name, "Height", Me.Height
        SaveSetting strMainKey, Me.Name, "Width", Me.Width
    End If

End Sub
Private Sub ClearDataType()
    For i = 0 To 17
        ParameterName2(i) = ""
    Next i
    mvarPaperType = Receipt
End Sub
Private Sub cmd_Ok_Click()
  If lstItemReports.ListIndex = -1 Then
    frmDisMsg.lblMessage = "‰«„ ê“«—‘ «‰ Œ«» ‰‘œÂ «” "
    frmDisMsg.Timer1.Enabled = True
    frmDisMsg.Show vbModal
    Exit Sub
 End If
    Dim ii, jj As Integer
    ClearDataType
    ii = -1
    jj = 0
    For i = 0 To 17 Step 2
        If LblIo(i).Visible = True Then
''            If InStr(1, ParameterName(i), "papertype", vbTextCompare) Then
''            Else
''                ii = ii + 1
''                ParameterName2(i) = ParameterName(i) & 1
''                If lblIO(i + 1).Visible = True Then
''                    ParameterName2(i + 1) = ParameterName(i + 1) & 2
''                    ii = ii + 1
''                End If
''            End If
''        End If
            If Text1(i).Visible = True Then
                ii = ii + 1
                ParameterName2(jj) = ParameterName(i) & 1
                ParameterData(jj) = Text1(i).Text
                If InStr(1, ParameterName(jj), "Date", vbTextCompare) Then
                    If Trim(ParameterData(jj)) = "" Then
                        ParameterData(jj) = Right(AccountYear, 2) & "/01/01"
                    End If
                End If
                If Trim(ParameterData(jj)) = "" Then
                    ParameterData(jj) = MinValue(jj)
                End If
                jj = jj + 1
                If LblIo(i + 1).Visible = True Then
                    ParameterName2(jj) = ParameterName(i) & 2
                    ParameterData(jj) = Text1(i + 1).Text
                    ii = ii + 1
                    If InStr(1, ParameterName(jj), "Date", vbTextCompare) Then
                        If Trim(ParameterData(jj)) = "" Then
                            ParameterData(jj) = mvarDate '"99/12/29"
                        End If
                    End If
                    If Trim(ParameterData(jj)) = "" Then
                        ParameterData(jj) = MaxValue(jj)
                    End If
                jj = jj + 1
                End If
            ElseIf Combo1(i).Visible = True Then
                If InStr(1, ParameterName(i), "papertype", vbTextCompare) Then
                    mvarPaperType = Combo1(i).ItemData(Combo1(i).ListIndex)
                Else
                    ii = ii + 1
                    ParameterName2(jj) = ParameterName(i) & 1
                    ParameterData(jj) = Combo1(i).ItemData(Combo1(i).ListIndex)
                    If Trim(ParameterData(jj)) = "" Then
                        ParameterData(jj) = MinValue(jj)
                    End If
                    jj = jj + 1
                    If LblIo(i + 1).Visible = True Then
                        ParameterName2(jj) = ParameterName(i) & 2
                        ParameterData(jj) = Combo1(i + 1).ItemData(Combo1(i + 1).ListIndex)
                        If Trim(ParameterData(jj)) = "" Then
                            ParameterData(jj) = MaxValue(jj)
                        End If
                        ii = ii + 1
                        jj = jj + 1
                    End If
                End If
            ElseIf MaskEdBox1(i).Visible = True Then
                ii = ii + 1
                ParameterName2(jj) = ParameterName(i) & 1
                ParameterData(jj) = MaskEdBox1(i).Text
                If Trim(ParameterData(jj)) = "" Then
                    ParameterData(jj) = MinValue(jj)
                End If
                jj = jj + 1
                If LblIo(i + 1).Visible = True Then
                    ParameterName2(jj) = ParameterName(i) & 2
                    ParameterData(jj) = MaskEdBox1(i + 1).Text
                    If Trim(ParameterData(jj)) = "" Then
                        ParameterData(jj) = MaxValue(jj)
                    End If
                    ii = ii + 1
                    jj = jj + 1
                End If
            End If
        End If
    Next i
    
    ReDim Parameter(jj + 2) As Parameter

    Parameter(0) = GenerateInputParameter("@SystemDate", adVarWChar, 50, Mid(clsDate.shamsi(Date), 3, 8))
    Parameter(1) = GenerateInputParameter("@SystemDay", adVarWChar, 50, clsDate.Find_DayOfWeek(Weekday(Date, vbSaturday)))
    Parameter(2) = GenerateInputParameter("@SystemTime", adVarWChar, 50, Mid(Str(time), 1, 5))
    For i = 0 To jj - 1
    
        Parameter(i + 3) = GenerateInputParameter2("@" & ParameterName2(i), ParameterType(i), parameterLengh(i), ParameterData(i))
    Next i
    
    If clsStation.Language = Farsi Then
        If mvarPaperType = Receipt Then
           CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\" & ReportFileName & ".rpt"
        Else
           CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\" & ReportFileName & "_A4.rpt"
        End If
    Else
        If mvarPaperType = Receipt Then
           CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\English\" & ReportFileName & "_En.rpt"
        Else
           CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\English\" & ReportFileName & "_En_A4.rpt"
        End If
    End If
    'CrystalReport1.Status
    'CrystalReport1.PrinterSelect
    Dim fileSystem As New FileSystemObject
    Dim IsFileExist As Boolean
    IsFileExist = fileSystem.FileExists(CrystalReport1.ReportFileName)
    If IsFileExist = False Then
            frmDisMsg.lblMessage = " ›«Ì·  " & CrystalReport1.ReportFileName & "ÅÌœ« ‰‘œ "
            frmDisMsg.Timer1.Interval = 3000
            frmDisMsg.Timer1.Enabled = True
            frmDisMsg.Show vbModal
            Exit Sub
    End If
    ReportShow
    
End Sub

Public Sub ReportShow()
    On Error GoTo ErrorHandler
    '-----------------------
    CrystalReport1.ReportTitle = LblIo(18).Caption  ' ReportHeader
    CrystalReport1.Destination = crptToWindow 'crptToPrinter '
    Dim intIndex As Integer
   
    For intIndex = 0 To UBound(Parameter) - LBound(Parameter)
        CrystalReport1.ParameterFields(intIndex) = CStr(Parameter(intIndex).Name) & ";" & CStr(Parameter(intIndex).Value) & ";" & "True"
    Next intIndex
   
   For intIndex = UBound(Parameter) - LBound(Parameter) + 1 To 30
        CrystalReport1.ParameterFields(intIndex) = ""
   Next intIndex
    CrystalReport1.ProgressDialog = True
    CrystalReport1.RetrieveDataFiles
    ODBCSetting clsArya.ServerName, clsArya.DbName
    CrystalReport1.Connect = CrystallConnection
    CrystalReport1.Action = 1
    If PaperType = 1 Then
       CrystalReport1.PageZoom (100)
    Else
       CrystalReport1.PageZoom (100)
       
    End If
    Exit Sub
ErrorHandler:
   MsgBox err.Description & "  File Name:  " & CrystalReport1.ReportFileName
       Resume Next
   
End Sub



