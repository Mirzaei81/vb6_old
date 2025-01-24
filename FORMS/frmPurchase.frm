VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{7AEDC602-D94C-11D1-BB7A-00E0290EA3C9}#1.0#0"; "ResizeKit.ocx"
Object = "{75D4F767-8785-11D3-93AD-0000832EF44D}#2.14#0"; "FAST2001.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "mscomm32.ocx"
Begin VB.Form frmPurchase 
   BackColor       =   &H00C0C0C0&
   Caption         =   "›«ﬂ Ê— Œ—Ìœ"
   ClientHeight    =   9600
   ClientLeft      =   14745
   ClientTop       =   450
   ClientWidth     =   14925
   FillColor       =   &H00EACCEC&
   Icon            =   "frmPurchase.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9600
   ScaleWidth      =   14925
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0C0C0&
      Height          =   1935
      Left            =   3240
      TabIndex        =   85
      Top             =   6480
      Width           =   3735
      Begin VB.CommandButton BtnMenu 
         BackColor       =   &H0000FFFF&
         Caption         =   "Ã” ÃÊÌ ò«·«"
         BeginProperty Font 
            Name            =   "Nazanin"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   0
         Left            =   2160
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   90
         Top             =   720
         Width           =   1455
      End
      Begin VB.CommandButton BtnKalaDelete 
         BackColor       =   &H000000FF&
         Caption         =   "Õ–› ò«·«"
         BeginProperty Font 
            Name            =   "Nazanin"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2160
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   89
         Tag             =   "-"
         Top             =   1320
         Width           =   1455
      End
      Begin VB.CommandButton cmbNewGood 
         BackColor       =   &H00C0C000&
         Caption         =   "ò«·«Ì ÃœÌœ"
         BeginProperty Font 
            Name            =   "Nazanin"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2160
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   88
         Top             =   120
         Width           =   1455
      End
      Begin VB.CommandButton cmdTempFich 
         BackColor       =   &H00C0C000&
         Caption         =   "”‰œ „Êﬁ "
         BeginProperty Font 
            Name            =   "Nazanin"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   87
         Top             =   240
         Width           =   1935
      End
      Begin VB.CommandButton CmdAnalyze 
         BackColor       =   &H00404080&
         Caption         =   "¬‰«·Ì“ ò«·«Ì ‰«Œ«·’"
         BeginProperty Font 
            Name            =   "Nazanin"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   86
         Top             =   960
         Width           =   1935
      End
   End
   Begin VB.Frame Frame6 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   3240
      TabIndex        =   79
      Top             =   8400
      Width           =   8175
      Begin MSComctlLib.StatusBar sbrFactorProp 
         Height          =   495
         Left            =   120
         TabIndex        =   80
         Top             =   120
         Width           =   5520
         _ExtentX        =   9737
         _ExtentY        =   873
         _Version        =   393216
         BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
            NumPanels       =   4
            BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               AutoSize        =   2
            EndProperty
            BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               AutoSize        =   2
            EndProperty
            BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               Alignment       =   1
               AutoSize        =   2
            EndProperty
            BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               AutoSize        =   2
            EndProperty
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSComctlLib.StatusBar StatusBar 
         Height          =   495
         Left            =   360
         TabIndex        =   81
         Top             =   600
         Width           =   7335
         _ExtentX        =   12938
         _ExtentY        =   873
         SimpleText      =   "\"
         _Version        =   393216
         BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
            NumPanels       =   7
            BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               AutoSize        =   1
               Bevel           =   2
               Object.Width           =   1235
               MinWidth        =   1235
               Picture         =   "frmPurchase.frx":A4C2
            EndProperty
            BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               AutoSize        =   2
               Object.Width           =   2381
               MinWidth        =   2381
            EndProperty
            BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               AutoSize        =   2
               Object.Width           =   2381
               MinWidth        =   2381
            EndProperty
            BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               AutoSize        =   2
               Object.Width           =   2381
               MinWidth        =   2381
            EndProperty
            BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               AutoSize        =   2
               Object.Width           =   2381
               MinWidth        =   2381
            EndProperty
            BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               AutoSize        =   2
               Object.Width           =   2381
               MinWidth        =   2381
            EndProperty
            BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
               Bevel           =   2
               Object.Width           =   1235
               MinWidth        =   1235
               Picture         =   "frmPurchase.frx":A7DC
            EndProperty
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin FLWCtrls.FWCheck FWChkAcc 
         Height          =   435
         Left            =   6840
         TabIndex        =   82
         Top             =   135
         Width           =   1230
         _ExtentX        =   2170
         _ExtentY        =   767
         Value           =   0   'False
         CheckType       =   5
         Caption         =   "Õ”«»œ«—Ì"
         Enabled         =   0   'False
         Color           =   4210688
         BackColor       =   16765183
         ForeColor       =   4194304
         FontName        =   "Nazanin"
         FontBold        =   -1  'True
         FontSize        =   11.25
         Alignment       =   1
      End
      Begin VB.Label LblAccNo 
         Alignment       =   2  'Center
         BackColor       =   &H00FFD0FF&
         BeginProperty Font 
            Name            =   "B Nazanin"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   5640
         RightToLeft     =   -1  'True
         TabIndex        =   83
         Top             =   120
         Width           =   1215
      End
   End
   Begin VB.TextBox txtScale 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   6960
      TabIndex        =   73
      Text            =   "TxtScale"
      Top             =   5160
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Frame KeyPadMenu 
      BorderStyle     =   0  'None
      Height          =   3105
      Left            =   11880
      RightToLeft     =   -1  'True
      TabIndex        =   52
      Top             =   6480
      Width           =   2805
      Begin VB.CommandButton BtnKeypad 
         BackColor       =   &H8000000D&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Index           =   0
         Left            =   960
         TabIndex        =   64
         Tag             =   "0"
         Top             =   2280
         Width           =   795
      End
      Begin VB.CommandButton BtnKeypad 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000D&
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Index           =   1
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   63
         Tag             =   "1"
         Top             =   1560
         Width           =   795
      End
      Begin VB.CommandButton BtnKeypad 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000D&
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Index           =   2
         Left            =   960
         RightToLeft     =   -1  'True
         TabIndex        =   62
         Tag             =   "2"
         Top             =   1560
         Width           =   795
      End
      Begin VB.CommandButton BtnKeypad 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000D&
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Index           =   3
         Left            =   1800
         RightToLeft     =   -1  'True
         TabIndex        =   61
         Tag             =   "3"
         Top             =   1560
         Width           =   795
      End
      Begin VB.CommandButton BtnKeypad 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000D&
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Index           =   4
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   60
         Tag             =   "4"
         Top             =   840
         Width           =   795
      End
      Begin VB.CommandButton BtnKeypad 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000D&
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Index           =   5
         Left            =   960
         RightToLeft     =   -1  'True
         TabIndex        =   59
         Tag             =   "5"
         Top             =   840
         Width           =   795
      End
      Begin VB.CommandButton BtnKeypad 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000D&
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Index           =   6
         Left            =   1800
         RightToLeft     =   -1  'True
         TabIndex        =   58
         Tag             =   "6"
         Top             =   840
         Width           =   795
      End
      Begin VB.CommandButton BtnKeypad 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000D&
         Caption         =   "7"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Index           =   7
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   57
         Tag             =   "7"
         Top             =   120
         Width           =   795
      End
      Begin VB.CommandButton BtnKeypad 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000D&
         Caption         =   "8"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Index           =   8
         Left            =   960
         RightToLeft     =   -1  'True
         TabIndex        =   56
         Tag             =   "8"
         Top             =   120
         Width           =   795
      End
      Begin VB.CommandButton BtnKeypad 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000D&
         Caption         =   "9"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Index           =   9
         Left            =   1800
         RightToLeft     =   -1  'True
         TabIndex        =   55
         Tag             =   "9"
         Top             =   120
         Width           =   795
      End
      Begin VB.CommandButton BtnKeypad 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000D&
         Caption         =   "."
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Index           =   10
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   54
         Tag             =   "."
         Top             =   2280
         Width           =   795
      End
      Begin VB.CommandButton BtnKeypad 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000D&
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Index           =   11
         Left            =   1800
         RightToLeft     =   -1  'True
         TabIndex        =   53
         Tag             =   "%"
         Top             =   2280
         Width           =   795
      End
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00C0C0C0&
      Height          =   1935
      Left            =   9120
      TabIndex        =   0
      Top             =   6480
      Width           =   2295
      Begin VB.CommandButton cmdColor 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000D&
         Caption         =   " €ÌÌ— —‰ê"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   120
         TabIndex        =   92
         Tag             =   "3"
         Top             =   120
         UseMaskColor    =   -1  'True
         Width           =   960
      End
      Begin VB.TextBox txtBarcode 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   120
         MultiLine       =   -1  'True
         RightToLeft     =   -1  'True
         TabIndex        =   1
         ToolTipText     =   "Ê—Êœ »«—ﬂœ ﬂ«·«"
         Top             =   1200
         Width           =   2145
      End
      Begin VB.Label lblBarCode 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   120
         TabIndex        =   67
         ToolTipText     =   "‰„«Ì‘ »«—ﬂœ"
         Top             =   600
         Width           =   2140
      End
      Begin VB.Label lblNum 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   450
         Left            =   1080
         RightToLeft     =   -1  'True
         TabIndex        =   66
         ToolTipText     =   "‰„«Ì‘ «—ﬁ«„ Ê—ÊœÌ"
         Top             =   120
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
      Height          =   2055
      Left            =   240
      TabIndex        =   19
      Top             =   120
      Width           =   14535
      Begin VB.ComboBox cmbDestination 
         BackColor       =   &H00FFC0C0&
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
         Left            =   7920
         RightToLeft     =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   91
         TabStop         =   0   'False
         Top             =   1440
         Width           =   2955
      End
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
         Left            =   5400
         RightToLeft     =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   84
         Top             =   1440
         Width           =   2475
      End
      Begin VB.CommandButton cmdTurnOver 
         BackColor       =   &H0000C0C0&
         Caption         =   "ê—œ‘ Õ”«»   «„Ì‰ ﬂ‰‰œÂ"
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   9
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   2760
         RightToLeft     =   -1  'True
         TabIndex        =   78
         Top             =   1320
         Width           =   1575
      End
      Begin VB.CommandButton fwBtnCustFind 
         BackColor       =   &H00FF8080&
         Caption         =   " «„Ì‰ ò‰‰œÂ"
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   550
         Left            =   9600
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   72
         Top             =   120
         Width           =   1335
      End
      Begin VB.CommandButton fwBtnDailyHavale 
         BackColor       =   &H00FF8080&
         Caption         =   "ÕÊ«·Â —Ê“«‰Â"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   550
         Left            =   9600
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   71
         Top             =   720
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.TextBox txtDescription 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Nazanin"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   570
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   25
         Top             =   600
         Width           =   4215
      End
      Begin VB.ComboBox cmbDestInventory 
         BackColor       =   &H00FFC0C0&
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
         RightToLeft     =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   1440
         Width           =   2355
      End
      Begin VB.ComboBox cmbInventory 
         BackColor       =   &H00FFC0C0&
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
         RightToLeft     =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   840
         Width           =   2355
      End
      Begin VB.ComboBox CmbStatus 
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "Nazanin"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   480
         Left            =   11025
         RightToLeft     =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   240
         Width           =   2355
      End
      Begin VB.TextBox txtCustomer 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "Nazanin"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   480
         Left            =   5400
         Locked          =   -1  'True
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   200
         Width           =   4095
      End
      Begin VB.TextBox txtDate 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   435
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   120
         Width           =   1485
      End
      Begin FLWCtrls.FWScrollText fwScrollTextCust 
         Height          =   525
         Left            =   7920
         TabIndex        =   26
         Top             =   720
         Width           =   2985
         _ExtentX        =   5265
         _ExtentY        =   926
         Caption         =   ""
         BackColor       =   16761024
         FontName        =   "Nazanin"
         FontBold        =   -1  'True
         FontSize        =   11.25
      End
      Begin FLWCtrls.FWLed FWLed1 
         Height          =   615
         Left            =   120
         TabIndex        =   51
         Top             =   1200
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   1085
         BorderStyle     =   10
         ColorOn         =   192
         ColorOff        =   12632256
         BackColor       =   12632256
      End
      Begin FLWCtrls.FWLabel fwlblMode 
         Height          =   735
         Left            =   1800
         Top             =   1200
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   1296
         Enabled         =   -1  'True
         Caption         =   "„—Ê—"
         FillType        =   4
         FirstColor      =   12582912
         SecondColor     =   10070188
         Angle           =   0
         ForeColor       =   -2147483630
         BackColor       =   -2147483633
         FontName        =   "B Homa"
         FontBold        =   -1  'True
         FontSize        =   15.75
         Alignment       =   2
         Picture         =   "frmPurchase.frx":AAF6
      End
      Begin FLWCtrls.FWLabel fwlblRecursive 
         Height          =   525
         Left            =   3600
         Top             =   120
         Visible         =   0   'False
         Width           =   1665
         _ExtentX        =   2937
         _ExtentY        =   926
         Enabled         =   -1  'True
         Caption         =   "„—ÃÊ⁄Ì"
         FirstColor      =   12632319
         SecondColor     =   192
         Angle           =   0
         ForeColor       =   -2147483630
         BackColor       =   -2147483633
         FontName        =   "Nazanin"
         FontBold        =   -1  'True
         FontSize        =   18
         Alignment       =   2
         Picture         =   "frmPurchase.frx":AB12
      End
      Begin VB.Label fwStatusBarCust 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Nazanin"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   555
         Left            =   5400
         TabIndex        =   65
         Top             =   720
         Width           =   2490
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00EACCEC&
         BackStyle       =   0  'Transparent
         Caption         =   "‘—Õ ”‰œ : "
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000001&
         Height          =   375
         Left            =   4320
         RightToLeft     =   -1  'True
         TabIndex        =   31
         Top             =   720
         Width           =   975
      End
      Begin VB.Label lblDestInventory 
         BackColor       =   &H00EACCEC&
         BackStyle       =   0  'Transparent
         Caption         =   "„ﬁ’œ"
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000001&
         Height          =   375
         Left            =   13680
         RightToLeft     =   -1  'True
         TabIndex        =   30
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label lblInventory 
         BackColor       =   &H00EACCEC&
         BackStyle       =   0  'Transparent
         Caption         =   "«‰»«—"
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000001&
         Height          =   375
         Left            =   13680
         RightToLeft     =   -1  'True
         TabIndex        =   29
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00EACCEC&
         BackStyle       =   0  'Transparent
         Caption         =   "Ê÷⁄Ì  ”‰œ:"
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   11.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   375
         Left            =   13320
         RightToLeft     =   -1  'True
         TabIndex        =   28
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label lblDate 
         BackColor       =   &H00EACCEC&
         BackStyle       =   0  'Transparent
         Caption         =   " «—ÌŒ "
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
         Left            =   1680
         RightToLeft     =   -1  'True
         TabIndex        =   27
         Top             =   120
         Width           =   1125
      End
   End
   Begin VSFlex7LCtl.VSFlexGrid FlxDetail 
      Height          =   4095
      Left            =   120
      TabIndex        =   17
      Top             =   2280
      Width           =   14745
      _cx             =   26009
      _cy             =   7223
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Nazanin"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483634
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483644
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483646
      BackColorAlternate=   12640511
      GridColor       =   -2147483632
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   4
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   14
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   500
      RowHeightMax    =   500
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmPurchase.frx":AB2E
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
      Left            =   10155
      TabIndex        =   45
      TabStop         =   0   'False
      ToolTipText     =   "ﬂœ „«‘Ì‰ ¬·«  —« œ— »— „ÌêÌ—œ"
      Top             =   600
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
      Left            =   10155
      RightToLeft     =   -1  'True
      TabIndex        =   44
      TabStop         =   0   'False
      Top             =   2565
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
      Left            =   10155
      RightToLeft     =   -1  'True
      TabIndex        =   43
      TabStop         =   0   'False
      Top             =   2940
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
      Left            =   10155
      RightToLeft     =   -1  'True
      TabIndex        =   42
      TabStop         =   0   'False
      Top             =   3315
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
      Left            =   10155
      TabIndex        =   41
      TabStop         =   0   'False
      ToolTipText     =   "ﬂœ „«‘Ì‰ ¬·«  —« œ— »— „ÌêÌ—œ"
      Top             =   3720
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
      Left            =   10155
      RightToLeft     =   -1  'True
      TabIndex        =   40
      TabStop         =   0   'False
      Top             =   4125
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.TextBox txtRegDate 
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
      Left            =   11085
      RightToLeft     =   -1  'True
      TabIndex        =   39
      TabStop         =   0   'False
      Top             =   4530
      Visible         =   0   'False
      Width           =   1605
   End
   Begin VB.TextBox txtTime 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   10155
      RightToLeft     =   -1  'True
      TabIndex        =   38
      TabStop         =   0   'False
      Top             =   1485
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.TextBox txtUser 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   10155
      RightToLeft     =   -1  'True
      TabIndex        =   37
      TabStop         =   0   'False
      Top             =   2205
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.TextBox Txtservice 
      Alignment       =   1  'Right Justify
      Height          =   360
      Left            =   10155
      RightToLeft     =   -1  'True
      TabIndex        =   36
      Text            =   "txtservice"
      Top             =   1845
      Visible         =   0   'False
      Width           =   2535
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
      Left            =   10155
      RightToLeft     =   -1  'True
      TabIndex        =   35
      TabStop         =   0   'False
      Top             =   5295
      Visible         =   0   'False
      Width           =   2535
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
      Left            =   10140
      RightToLeft     =   -1  'True
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   1005
      Visible         =   0   'False
      Width           =   2550
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
      Left            =   10140
      RightToLeft     =   -1  'True
      TabIndex        =   33
      TabStop         =   0   'False
      Top             =   5700
      Visible         =   0   'False
      Width           =   2550
   End
   Begin VB.TextBox txtServicePercent 
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
      Left            =   17520
      RightToLeft     =   -1  'True
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   5730
      Visible         =   0   'False
      Width           =   2550
   End
   Begin VB.ListBox lstGoodKey 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   630
      Left            =   1080
      RightToLeft     =   -1  'True
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   3840
      Visible         =   0   'False
      Width           =   3660
   End
   Begin VB.Frame Frame1 
      Height          =   3255
      Left            =   0
      TabIndex        =   9
      Top             =   6360
      Width           =   3135
      Begin VB.CommandButton cmbPackingTotal 
         BackColor       =   &H00FF8080&
         Caption         =   "»” Â »‰œÌ"
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
         Left            =   1560
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   68
         Top             =   1440
         Width           =   1095
      End
      Begin VB.CommandButton cmbCarryFeeTotal 
         BackColor       =   &H00FF8080&
         Caption         =   "ò—«ÌÂ Õ„·"
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
         Left            =   1560
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   70
         Top             =   960
         Width           =   1095
      End
      Begin VB.CommandButton CmdDiscountTotal 
         BackColor       =   &H00000080&
         Caption         =   " Œ›Ì›"
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
         Left            =   1560
         RightToLeft     =   -1  'True
         Style           =   1  'Graphical
         TabIndex        =   69
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00404080&
         Caption         =   "⁄Ê«—÷"
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
         Height          =   360
         Left            =   1560
         RightToLeft     =   -1  'True
         TabIndex        =   77
         Top             =   1920
         Width           =   1095
      End
      Begin VB.Label lblTax 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00404080&
         Caption         =   "„«·Ì« "
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
         Height          =   360
         Left            =   1560
         RightToLeft     =   -1  'True
         TabIndex        =   76
         Top             =   2280
         Width           =   1095
      End
      Begin VB.Label LblDutyTotal 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000013&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   75
         Top             =   1920
         Width           =   1140
      End
      Begin VB.Label LblTaxTotal 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000013&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   74
         Top             =   2280
         Width           =   1140
      End
      Begin VB.Label SumPriceLabel 
         Alignment       =   2  'Center
         BackColor       =   &H00EACCEC&
         BackStyle       =   0  'Transparent
         Caption         =   "„»·€ "
         BeginProperty Font 
            Name            =   "Nazanin"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000001&
         Height          =   555
         Left            =   2400
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Top             =   2760
         Width           =   660
      End
      Begin VB.Label lblSumPrice 
         Alignment       =   2  'Center
         BackColor       =   &H0080C0FF&
         BackStyle       =   0  'Transparent
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
            Name            =   "Tahoma"
            Size            =   18
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   500
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   2700
         Width           =   2355
      End
      Begin VB.Label lblPackingTotal 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0080C0FF&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   1410
         Width           =   1140
      End
      Begin VB.Label lblCarryFeeTotal 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0080C0FF&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   945
         Width           =   1170
      End
      Begin VB.Label lblDiscountTotal 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0080C0FF&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   435
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   480
         Width           =   1170
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Ã„⁄     "
         BeginProperty Font 
            Name            =   "Nazanin"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   615
         Left            =   2040
         TabIndex        =   11
         Top             =   120
         Width           =   495
      End
      Begin VB.Label LblSubTotal 
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
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   435
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   120
         Width           =   1635
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0C0C0&
      Height          =   1935
      Left            =   6960
      TabIndex        =   2
      Top             =   6480
      Width           =   2055
      Begin FLWCtrls.FWCheck chKTax 
         Height          =   465
         Left            =   1230
         TabIndex        =   93
         Top             =   195
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   820
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
      Begin VB.Label txtSumWeightTotal 
         Alignment       =   1  'Right Justify
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
         Height          =   435
         Left            =   120
         TabIndex        =   8
         Top             =   1200
         Width           =   1155
      End
      Begin VB.Label txtSumCountWeight 
         Alignment       =   1  'Right Justify
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
         Height          =   435
         Left            =   120
         TabIndex        =   7
         Top             =   720
         Width           =   630
      End
      Begin VB.Label lblSumWeightTotal 
         Alignment       =   2  'Center
         BackColor       =   &H00EACCEC&
         BackStyle       =   0  'Transparent
         Caption         =   "Ã„⁄ Ê“‰"
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000001&
         Height          =   435
         Left            =   1320
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   1200
         Width           =   780
      End
      Begin VB.Label lblSumCountWeight 
         Alignment       =   2  'Center
         BackColor       =   &H00EACCEC&
         BackStyle       =   0  'Transparent
         Caption         =   "Ê“‰Ì"
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000001&
         Height          =   435
         Left            =   795
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   720
         Width           =   435
      End
      Begin VB.Label txtSumCountNo 
         Alignment       =   1  'Right Justify
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
         Height          =   435
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   630
      End
      Begin VB.Label lblSumCountNo 
         Alignment       =   2  'Center
         BackColor       =   &H00EACCEC&
         BackStyle       =   0  'Transparent
         Caption         =   "⁄œœÌ"
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   9
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000001&
         Height          =   495
         Left            =   705
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   240
         Width           =   585
      End
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   9840
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
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   10845
      Top             =   7620
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPurchase.frx":AC54
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPurchase.frx":ACB2
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin FLWCtrls.FWLabel3D lblServePlace 
      Height          =   555
      Left            =   13440
      Top             =   240
      Visible         =   0   'False
      Width           =   1545
      _ExtentX        =   2725
      _ExtentY        =   979
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor1      =   -2147483645
      ForeColor2      =   -2147483643
      BackColor       =   8869511
      Caption         =   ""
      Alignment       =   2
   End
   Begin MSCommLib.MSComm mscSerial 
      Index           =   1
      Left            =   1080
      Top             =   1080
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin FLWCtrls.FWLabel fwCash 
      Height          =   555
      Left            =   11640
      Top             =   240
      Width           =   1740
      _ExtentX        =   3069
      _ExtentY        =   979
      Enabled         =   -1  'True
      Caption         =   ""
      FillType        =   3
      FirstColor      =   9981440
      SecondColor     =   16777215
      Angle           =   0
      ForeColor       =   0
      BackColor       =   128
      FontName        =   "B Homa"
      FontBold        =   -1  'True
      FontSize        =   9.75
      Alignment       =   2
      Picture         =   "frmPurchase.frx":AD10
   End
   Begin RESIZEKITLibCtl.ResizeKit ResizeKit1 
      Height          =   480
      Left            =   3240
      OleObjectBlob   =   "frmPurchase.frx":AD2C
      TabIndex        =   46
      Top             =   0
      Width           =   480
   End
   Begin VB.Label lblPackingPercent 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   " Œ›Ì›"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   285
      Left            =   16680
      RightToLeft     =   -1  'True
      TabIndex        =   50
      Top             =   5730
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Label lblServicePercent 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "”—ÊÌ”"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   285
      Left            =   18480
      RightToLeft     =   -1  'True
      TabIndex        =   49
      Top             =   5850
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Label lblStatus 
      Alignment       =   1  'Right Justify
      Height          =   375
      Left            =   10110
      RightToLeft     =   -1  'True
      TabIndex        =   48
      Top             =   6210
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label LblAccountYear 
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
      ForeColor       =   &H00800000&
      Height          =   495
      Left            =   1560
      TabIndex        =   47
      Top             =   120
      Width           =   1935
   End
End
Attribute VB_Name = "frmPurchase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Dim BitSaveTempReceived As Boolean
Dim SerialTempReceived As Double
Dim BitTempReceived As Boolean
Dim Exit_Keypress_Flag As Boolean
Dim mvarKeyCode, MvarShiftKey As Integer
Dim textDescription As Boolean

Dim MyFormAddEditMode As EnumAddEditMode
Dim mVarOrderType As EnumOrderType

'Dim mycls As FileClass

Dim clsDate As New clsDate
Dim ClsCnvKeyBoard As New ClsCnvKeyBoard

Dim DeviceCode(1 To 4) As Integer
Dim RThreshold(1 To 4) As Integer
Dim DeviceType(1 To 4) As Integer
Dim intTempFich As Double

Dim rctmp As New ADODB.Recordset

Dim RstTemp As New ADODB.Recordset

Dim mvarEmpty As Boolean
Dim mvarbarcode As Boolean
Dim blnCreditCust As Boolean
Dim boolPayment As Boolean
Dim BalancePayment As Boolean
Dim mvarStationNo As Integer
Dim BlnFormLoaded As Boolean

Dim MaxRowFlexGrid As Integer

Dim i As Integer
Dim intSumOfCurrentServePlaces As Integer

Dim dblFichUser As Double
Dim intSerialNo As Double
Dim dblBasFichNo As Double

Dim PosTempFactorNo As Integer
Dim PosTempPrice As Double
Dim Parameter() As Parameter
Dim mvarBuyPrice As Long
Dim TmpGoodDiscount As Long
Dim DestInventoryNo As Integer
Public FromDate, ToDate As String
Dim BitAutoHavale As Boolean

Public Function ColorSetting()
On Error Resume Next
    
    Dim Purchase_BackColorForm As Long
    Dim Purchase_BackColorBtn As Long
    Dim Purchase_BackColorFlexGrid As Long
    Dim filetemp As New FileSystemObject
    Dim tempstring As TextStream
    Dim Str As String
    Dim LenghStr As Integer
    
    Dim IsFileExist As Boolean
    
    If UserSettingFile = "" Then End    'Only  For  Make Exe File
    Set tempstring = filetemp.OpenTextFile(UserSettingFile, ForReading, False, TristateFalse)
    
    Do While tempstring.AtEndOfLine = False
       Str = tempstring.ReadLine
       LenghStr = InStr(1, Str, "=", vbTextCompare)
       
       If InStr(1, Str, "Purchase_BackColorForm", vbTextCompare) Then
          Purchase_BackColorForm = Val(Mid(Str, LenghStr + 1))
       
       ElseIf InStr(1, Str, "Purchase_BackColorBtn", vbTextCompare) Then
          Purchase_BackColorBtn = Mid(Str, LenghStr + 1)
       
       ElseIf InStr(1, Str, "Purchase_BackColorFlexGrid", vbTextCompare) Then
          Purchase_BackColorFlexGrid = Mid(Str, LenghStr + 1)
       
       End If
    Loop
    tempstring.Close
    
    Me.BackColor = Purchase_BackColorForm
  '  frameMenu.BackColor = Purchase_BackColorForm
    KeyPadMenu.BackColor = Purchase_BackColorForm
    FlxDetail.BackColor = Purchase_BackColorFlexGrid
 '   FWScrollText1.BackColor = Purchase_BackColorForm
    txtDate.BackColor = Purchase_BackColorForm
    
    Frame1.BackColor = Purchase_BackColorForm
    Frame2.BackColor = Purchase_BackColorForm
    Frame3.BackColor = Purchase_BackColorForm
    FWLed1.BackColor = Purchase_BackColorForm
    FWLed1.ColorOff = Purchase_BackColorForm
    txtDescription.BackColor = Me.BackColor
    txtBarcode.BackColor = Me.BackColor
    
    For i = 1 To BtnMenu.Count - 1
        BtnMenu(i).BackColor = Purchase_BackColorBtn
    '    BtnMenu(i).ForeColor = &H80000005
    Next
End Function


Private Sub chKTax_Click()
Dim ii As Long
With FlxDetail
    For ii = 1 To MaxRowFlexGrid - 1
        .TextMatrix(ii, 19) = chKTax.Value
        .TextMatrix(ii, 20) = chKTax.Value
    Next
End With
RefreshLables

End Sub

Private Sub cmbDestInventory_Click()
    If cmbDestInventory.ListIndex <> -1 Then
        ClearDataFlexGrid
        DestInventoryNo = cmbDestInventory.ItemData(cmbDestInventory.ListIndex)
        txtDescription.Text = "ÕÊ«·Â «“ " & cmbInventory.Text & " »Â " & cmbDestInventory.Text
      
        For i = 1 To FlxDetail.Rows - 1
            If FlxDetail.TextMatrix(i, 5) <> "" Then
                FlxDetail.TextMatrix(i, 14) = DestInventoryNo
            End If
        Next i
    Else
        DestInventoryNo = 0
    End If
'    If DestInventoryNo > 0 Then
'        fwBtnDailyHavale.Enabled = True
'        Else: fwBtnDailyHavale.Enabled = False
'     End If
    Dim rctmp As New ADODB.Recordset
    ReDim Parameter(1) As Parameter
    Parameter(0) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
    Parameter(1) = GenerateInputParameter("@Type", adInteger, 4, 1) 'All Inventory
    Set rctmp = RunParametricStoredProcedure2Rec("GetInventory", Parameter)
    
    If Not (rctmp.EOF = True And rctmp.BOF = True) Then
        Do While rctmp.EOF <> True
            If DestInventoryNo = rctmp.Fields("InventoryNo") Then
                Tafsili_3 = rctmp.Fields("Tafsili")
                Exit Do
            End If
            rctmp.MoveNext
        Loop
         
    End If
    Set rctmp = Nothing
End Sub

Private Sub cmbInventory_Click()
    If cmbInventory.ListIndex <> -1 Then
        ClearDataFlexGrid
        InventoryNo = cmbInventory.ItemData(cmbInventory.ListIndex)
        If mvarStatus = Purchase Then
            txtDescription.Text = "Œ—Ìœ »Â " & cmbInventory.Text
        ElseIf mvarStatus = fromStore Then
            txtDescription.Text = "ÕÊ«·Â «“ " & cmbInventory.Text & " »Â " & cmbDestInventory.Text
        ElseIf mvarStatus = PurchaseReturn Then
            txtDescription.Text = "»—ê‘  «“ Œ—Ìœ »Â " & cmbInventory.Text
        Else
            txtDescription.Text = ""
        End If
        For i = 1 To FlxDetail.Rows - 1
            If FlxDetail.TextMatrix(i, 5) <> "" Then
                FlxDetail.TextMatrix(i, 13) = InventoryNo
            End If
        Next i
    Else
         InventoryNo = 0
    End If
    Dim rctmp As New ADODB.Recordset
    ReDim Parameter(1) As Parameter
    Parameter(0) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
    Parameter(1) = GenerateInputParameter("@Type", adInteger, 4, 1) 'All Inventory
    Set rctmp = RunParametricStoredProcedure2Rec("GetInventory", Parameter)
    
    If Not (rctmp.EOF = True And rctmp.BOF = True) Then
        Do While rctmp.EOF <> True
            If InventoryNo = rctmp.Fields("InventoryNo") Then
                Tafsili_2 = rctmp.Fields("Tafsili")
                Exit Do
            End If
            rctmp.MoveNext
        Loop
         
    End If
    Set rctmp = Nothing
End Sub

Private Sub cmbNewGood_Click()
    NewGoodFlag = True
    frmGood.Show
End Sub

Private Sub CmbStatus_Click()
    fwScrollTextCust.Visible = True
    cmbDestination.Visible = False
    fwBtnDailyHavale.Visible = False
    fwBtnDailyHavale.Enabled = False
    On Error GoTo ErrorHandler
   mvarStatus = CmbStatus.ItemData(CmbStatus.ListIndex)
   cmbDestInventory.ListIndex = -1
   
   If mvarStatus = 1 Then
    '   FWLabel1.Caption = "›«ﬂ Ê— Œ—Ìœ"
       fwBtnDailyHavale.Visible = False
       mvarStatus = Purchase
       fwBtnCustFind.Enabled = True
              
       CmdDiscountTotal.Visible = True
       lblDiscountTotal.Visible = True
'       If clsArya.CustomerId = 93 Or clsArya.CustomerId = 931 Then
            cmbCarryFeeTotal.Visible = True
            cmbPackingTotal.Visible = True
            lblPackingTotal.Visible = True
            lblCarryFeeTotal.Visible = True
'       Else
'            cmbCarryFeeTotal.Visible = True
'            cmbPackingTotal.Visible = True
'            lblPackingTotal.Visible = True
'            lblCarryFeeTotal.Visible = True
'       End If
       cmbInventory.Enabled = True
       lblInventory.Caption = " «‰»«—"
       cmbInventory.Visible = True
       cmbDestInventory.Visible = True
       lblDestInventory.Visible = True
       lblDestInventory.Caption = " „ﬁ’œ »⁄œÌ"
       FlxDetail.TextMatrix(0, 13) = "„»œ«"
       FlxDetail.TextMatrix(0, 14) = "„ﬁ’œ"
       Add
   ElseIf mvarStatus = 3 Then
    '   FWLabel1.Caption = " ÷«Ì⁄«  "
       fwBtnDailyHavale.Visible = False
       txtCustomer.Tag = -1
       UpdatetxtCustomer
       fwBtnCustFind.Enabled = False
       CmdDiscountTotal.Visible = False
       cmbCarryFeeTotal.Visible = False
       cmbPackingTotal.Visible = False
       lblDiscountTotal.Visible = False
       lblPackingTotal.Visible = False
       lblCarryFeeTotal.Visible = False
       cmbInventory.Enabled = True
       lblInventory.Caption = " «‰»«—"
       cmbDestInventory.Visible = False
       lblDestInventory.Visible = False
       FlxDetail.TextMatrix(0, 13) = "„»œ«"
       FlxDetail.TextMatrix(0, 14) = "„ﬁ’œ"
       Add
   ElseIf mvarStatus = 4 Then
    '   FWLabel1.Caption = "»—ê‘  «“ Œ—Ìœ"
       fwBtnDailyHavale.Visible = False
       fwBtnCustFind.Enabled = True
       CmdDiscountTotal.Visible = True
       cmbCarryFeeTotal.Visible = True
       cmbPackingTotal.Visible = True
       lblDiscountTotal.Visible = True
       lblPackingTotal.Visible = True
       lblCarryFeeTotal.Visible = True
       cmbInventory.Enabled = True
       lblInventory.Caption = " «‰»«—"
       cmbDestInventory.Visible = False
       lblDestInventory.Visible = False
       FlxDetail.TextMatrix(0, 13) = "„»œ«"
       FlxDetail.TextMatrix(0, 14) = "„ﬁ’œ"
       Add
   ElseIf mvarStatus = 6 Then
    '   FWLabel1.Caption = "ÕÊ«·Â «‰ ﬁ«·Ì"
       fwBtnCustFind.Enabled = False
       CmdDiscountTotal.Visible = False
       cmbCarryFeeTotal.Visible = False
       cmbPackingTotal.Visible = False
       lblDiscountTotal.Visible = False
       lblPackingTotal.Visible = False
       lblCarryFeeTotal.Visible = False
       cmbInventory.Enabled = True
       cmbDestInventory.Visible = True
       cmbDestInventory.ListIndex = 0
       lblDestInventory.Visible = True
       lblInventory.Caption = " „»œ«"
       lblDestInventory.Caption = " „ﬁ’œ"
       FlxDetail.TextMatrix(0, 13) = "„»œ«"
       FlxDetail.TextMatrix(0, 14) = "„ﬁ’œ"
       Add
       fwBtnDailyHavale.Visible = True
       fwBtnDailyHavale.Enabled = True
       fwScrollTextCust.Visible = False
       If intVersion = Diamond Then cmbDestination.Visible = True
   ElseIf mvarStatus = 7 Then
    '   FWLabel1.Caption = "ÕÊ«·Â «‰ ﬁ«·Ì"
       fwBtnDailyHavale.Visible = True
       fwBtnDailyHavale.Enabled = False
       fwBtnCustFind.Enabled = False
       CmdDiscountTotal.Visible = False
       cmbCarryFeeTotal.Visible = False
       cmbPackingTotal.Visible = False
       lblDiscountTotal.Visible = False
       lblPackingTotal.Visible = False
       lblCarryFeeTotal.Visible = False
       cmbInventory.Visible = True
       cmbDestInventory.Visible = True
       cmbDestInventory.ListIndex = 0
       
       lblInventory.Visible = True
       lblDestInventory.Visible = True
       lblDestInventory.Caption = " „»œ«"
       lblInventory.Caption = " „ﬁ’œ"
       FlxDetail.TextMatrix(0, 14) = "„»œ«"
       FlxDetail.TextMatrix(0, 13) = "„ﬁ’œ"
       Add
       
   ElseIf mvarStatus = 8 Then
       fwBtnCustFind.Enabled = False
       CmdDiscountTotal.Visible = False
       cmbCarryFeeTotal.Visible = False
       cmbPackingTotal.Visible = False
       lblDiscountTotal.Visible = False
       lblPackingTotal.Visible = False
       lblCarryFeeTotal.Visible = False
       cmbInventory.Enabled = True
       cmbDestInventory.Visible = True
       cmbDestInventory.ListIndex = 0
       lblDestInventory.Visible = True
       lblInventory.Caption = " „»œ«"
       lblDestInventory.Caption = " „ﬁ’œ"
       FlxDetail.TextMatrix(0, 13) = "„»œ«"
       FlxDetail.TextMatrix(0, 14) = "„ﬁ’œ"
       Add
       fwBtnDailyHavale.Visible = False
       fwBtnDailyHavale.Enabled = False
       fwScrollTextCust.Visible = False
       If intVersion = Diamond Then cmbDestination.Visible = True
   
   ElseIf mvarStatus = 9 Then
       fwBtnDailyHavale.Visible = False
       fwBtnCustFind.Enabled = True
       CmdDiscountTotal.Visible = False
       cmbCarryFeeTotal.Visible = False
       cmbPackingTotal.Visible = False
       lblDiscountTotal.Visible = False
       lblPackingTotal.Visible = False
       lblCarryFeeTotal.Visible = False
       cmbInventory.Visible = True
       cmbDestInventory.Visible = False
       cmbDestInventory.ListIndex = 0
       
       lblInventory.Visible = True
       lblDestInventory.Visible = False
       lblInventory.Caption = " «‰»«—"
       FlxDetail.TextMatrix(0, 13) = "„»œ«"
       FlxDetail.TextMatrix(0, 14) = "„ﬁ’œ"
       Add
   End If
Exit Sub
ErrorHandler:

    MsgBox err.Description & "CmbStatus_Click"
End Sub

Private Sub CmdAnalyze_Click()
    If intVersion <> Diamond Then ShowDisMessage " ›ﬁÿ Ê—é‰ Â«Ì «·„«” „Ì  Ê«‰‰œ «“ «Ì‰ ﬁ«»·Ì  «” ›«œÂ ﬂ‰‰œ ", 2000: Exit Sub
    If MyFormAddEditMode = AddMode Then frmAnalyze.Show vbModal
End Sub

Private Sub CmdColor_Click()
    frmColor.Show vbModal
End Sub

Private Sub cmdTempFich_Click()
    If MyFormAddEditMode <> AddMode Then Exit Sub
     
''''    If cmbGarson_Seller.ListIndex = -1 Then
''''       cmbGarson_Seller.ListIndex = 0
''''    End If
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
                     st = GenerateDetailsString3(st, .TextMatrix(i, 1), .TextMatrix(i, 5), Val(.TextMatrix(i, 3)), Val(.TextMatrix(i, 10)), Val(.TextMatrix(i, 11)), " ", Val(.TextMatrix(i, 12)), Val(.TextMatrix(i, 13)), Val(.TextMatrix(i, 14)), .TextMatrix(i, 8), .TextMatrix(i, 9))
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
            
''''            If cmbTable.ListIndex = -1 Then
''''               cmbTable.ListIndex = 0
''''            End If
            
            
            ReDim Parameter(24) As Parameter
            
            Parameter(0) = GenerateInputParameter("@Status", adInteger, 4, mvarStatus)
            If (Me.txtCustomer.Tag > -1) Then
                Parameter(1) = GenerateInputParameter("@Owner", adInteger, 4, Me.txtCustomer.Tag)
            Else
                Parameter(1) = GenerateInputParameter("@Owner", adInteger, 4, -1)
            End If
            Parameter(2) = GenerateInputParameter("@Customer", adInteger, 4, 0)
            
            Parameter(3) = GenerateInputParameter("@DiscountTotal", adDouble, 8, Val(Me.lblDiscountTotal))
            Parameter(4) = GenerateInputParameter("@CarryFeeTotal", adDouble, 8, Val(Me.lblCarryFeeTotal))
            Parameter(5) = GenerateInputParameter("@SumPrice", adDouble, 8, Val(Me.lblSumPrice.Tag))
            Parameter(6) = GenerateInputParameter("@Recursive", adInteger, 4, Val(Me.txtRecursive.Text))
            Parameter(7) = GenerateInputParameter("@InCharge", adInteger, 4, 0)
            Parameter(8) = GenerateInputParameter("@FacPayment", adBoolean, 1, Abs(CInt(boolPayment)))
            Parameter(9) = GenerateInputParameter("@OrderType", adInteger, 4, mVarOrderType)
            Parameter(10) = GenerateInputParameter("@StationId", adInteger, 4, clsArya.StationNo)
            Parameter(11) = GenerateInputParameter("@ServiceTotal", adDouble, 8, 0)
            Parameter(12) = GenerateInputParameter("@PackingTotal", adDouble, 8, Val(Me.lblPackingTotal))
            Parameter(13) = GenerateInputParameter("@TableNo", adInteger, 4, 0)
            Parameter(14) = GenerateInputParameter("@User", adInteger, 4, mvarCurUserNo)
            Parameter(15) = GenerateInputParameter("@Date", adVarWChar, 50, txtDate.Text)
            
            
            Parameter(16) = GenerateInputParameter("@DetailsString", adVarWChar, 4000, DetailsString1)
            Parameter(17) = GenerateInputParameter("@NvcDescription", adVarWChar, 150, Right(txtDescription.Text, 150))
            Parameter(18) = GenerateInputParameter("@GuestNo", adSmallInt, 2, Null)
            Parameter(19) = GenerateInputParameter("@Branch", adInteger, 4, CurrentBranch)
            Parameter(20) = GenerateInputParameter("@TempAddress", adVarWChar, 255, " ")
            Parameter(21) = GenerateInputParameter("@DetailsString2", adVarWChar, 4000, DetailsString2)
            Parameter(22) = GenerateInputParameter("@DetailsString3", adVarWChar, 4000, DetailsString3)
            Parameter(23) = GenerateInputParameter("@DetailsString4", adVarWChar, 4000, DetailsString4)
            Parameter(24) = GenerateOutputParameter("@LastFacMNo", adInteger, 4)
                    
            RunParametricStoredProcedure "InsertFactorMasterDetailsTemp", Parameter
            
        Else
            ReDim Parameter(24) As Parameter
            Parameter(0) = GenerateInputParameter("@No", adBigInt, 8, intTempFich)
            Parameter(1) = GenerateInputParameter("@Status", adInteger, 4, mvarStatus)
            If (Me.txtCustomer.Tag > -1) Then
                Parameter(2) = GenerateInputParameter("@Owner", adInteger, 4, Me.txtCustomer.Tag)
            Else
                Parameter(2) = GenerateInputParameter("@Owner", adInteger, 4, -1)
            End If
            Parameter(3) = GenerateInputParameter("@Customer", adInteger, 4, 0)
            
            Parameter(4) = GenerateInputParameter("@DiscountTotal", adDouble, 8, Val(Me.lblDiscountTotal))
            Parameter(5) = GenerateInputParameter("@CarryFeeTotal", adDouble, 8, Val(Me.lblCarryFeeTotal))
            Parameter(6) = GenerateInputParameter("@SumPrice", adDouble, 8, Val(Me.lblSumPrice.Tag))
            Parameter(7) = GenerateInputParameter("@Recursive", adInteger, 4, Val(Me.txtRecursive.Text))
            Parameter(8) = GenerateInputParameter("@InCharge", adInteger, 4, 0)
            Parameter(9) = GenerateInputParameter("@FacPayment", adBoolean, 1, Abs(CInt(boolPayment)))
            Parameter(10) = GenerateInputParameter("@OrderType", adInteger, 4, mVarOrderType)
            Parameter(11) = GenerateInputParameter("@StationId", adInteger, 4, clsArya.StationNo)
            Parameter(12) = GenerateInputParameter("@ServiceTotal", adDouble, 8, 0)
            Parameter(13) = GenerateInputParameter("@PackingTotal", adDouble, 8, Val(Me.lblPackingTotal))
            Parameter(14) = GenerateInputParameter("@TableNo", adInteger, 4, 0)
            Parameter(15) = GenerateInputParameter("@Date", adVarWChar, 50, txtDate.Text)
            
            Parameter(16) = GenerateInputParameter("@DetailsString", adVarWChar, 4000, DetailsString1)
            Parameter(17) = GenerateInputParameter("@NvcDescription", adVarWChar, 150, Right(txtDescription.Text, 150))
            Parameter(18) = GenerateInputParameter("@GuestNo", adSmallInt, 2, Null)
            Parameter(19) = GenerateInputParameter("@Branch", adInteger, 4, CurrentBranch)
            Parameter(20) = GenerateInputParameter("@TempAddress", adVarWChar, 255, " ")
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
    
        boolPayment = rctmp.Fields("FacPayment").Value
        BalancePayment = False
        txtDiscount.Text = rctmp.Fields("DiscountTotal").Value
        txtCarryFee.Text = rctmp.Fields("CarryFeeTotal").Value
        txtPacking.Text = rctmp.Fields("PackingTotal").Value
        mVarOrderType = rctmp!OrderType
        mvarServePlace = rctmp!ServePlace
        intSumOfCurrentServePlaces = mvarServePlace
 '       InventoryNo = rctmp!intInventoryNo
        
        txtCustomer.Tag = IIf(IsNull(rctmp.Fields("Owner").Value), "-1", rctmp.Fields("Owner").Value)
        UpdatetxtCustomer
        With FlxDetail
        Do While Not (rctmp.EOF)
''''            If Not IsNull(rctmp!intInventoryNo) Then
''''                For ii = 0 To cmbInventory.ListCount - 1
''''                    If cmbInventory.ItemData(ii) = rctmp!intInventoryNo Then
''''                        cmbInventory.ListIndex = ii
''''                        ii = 0
''''                        Exit For
''''                    End If
''''                Next ii
''''            Else
''''                cmbInventory.ListIndex = -1
''''            End If
''''            If Not IsNull(rctmp!DestInventoryNo) Then
''''                cmbDestInventory.ListIndex = rctmp!DestInventoryNo - 1
''''            Else
''''                cmbDestInventory.ListIndex = -1
''''            End If
            ii = ii + 1
            .TextMatrix(ii, 0) = rctmp!intRow 'Number
            .TextMatrix(ii, 1) = rctmp!amount
            .TextMatrix(ii, 2) = rctmp!nvcName 'GoodName
            .TextMatrix(ii, 3) = rctmp!FeeUnit
            .TextMatrix(ii, 4) = rctmp!amount * rctmp!FeeUnit ' rctmp!FeeTotal
            .TextMatrix(ii, 5) = rctmp!GoodCode
            .TextMatrix(ii, 6) = rctmp!Weight ' rctmp!WeightUnit
            .TextMatrix(ii, 7) = rctmp!Unit
            .TextMatrix(ii, 8) = rctmp!ServePlace
            .TextMatrix(ii, 9) = ""
            .TextMatrix(ii, 10) = rctmp!Discount
            .TextMatrix(ii, 11) = rctmp!Rate
            .TextMatrix(ii, 12) = IIf(IsNull(rctmp!ExpireDate), "", rctmp!ExpireDate)
            .TextMatrix(ii, 13) = rctmp!intInventoryNo
            .TextMatrix(ii, 14) = IIf(IsNull(rctmp!DestInventoryNo), "", rctmp!DestInventoryNo)
            .TextMatrix(ii, 15) = rctmp!NumberOfUnit
            If rctmp.Fields("Mojodi").Value >= 0 Then
                If rctmp.Fields("Mojodi").Value <> Int(rctmp.Fields("Mojodi").Value) Then
                    .TextMatrix(ii, 16) = Format(rctmp.Fields("Mojodi").Value, "##.000")
                    .TextMatrix(ii, 16) = Val(.TextMatrix(ii, 16)) ' Delete Last Zeros
                Else
                     .TextMatrix(ii, 16) = rctmp.Fields("Mojodi").Value
                End If
            Else
                If rctmp.Fields("Mojodi").Value <> Int(rctmp.Fields("Mojodi").Value) Then
                    .TextMatrix(ii, 16) = -Format(rctmp.Fields("Mojodi").Value, "##.000")
                    .TextMatrix(ii, 16) = Val(.TextMatrix(ii, 16)) & "-" ' Delete Last Zeros
                Else
                     .TextMatrix(ii, 16) = -rctmp.Fields("Mojodi").Value & "-"
                End If
            End If
            .TextMatrix(ii, 17) = rctmp!SellPrice
            
            rctmp.MoveNext
            
            
            MaxRowFlexGrid = ii + 1
            If ii >= FlxDetail.Rows - 2 And rctmp.EOF = False Then
                AddEmptyRow
            End If

        Loop
        End With
        FlxDetail.Row = MaxRowFlexGrid - 1
        'mvarServePlace = FlxDetail.TextMatrix(MaxRowFlexGrid - 1, 8)
    End If
    
    rctmp.Close
''''    If mVarOrderType = ByPhone Then
''''       LblOrder.Caption = " ·›‰Ì"
''''    Else
''''       LblOrder.Caption = "Õ÷Ê—Ì"
''''    End If
    RefreshLables
    
    sbrFactorProp.Panels(1).Text = ""
    sbrFactorProp.Panels(2).Text = ""
    sbrFactorProp.Panels(3).Text = ""
    sbrFactorProp.Panels(4).Text = ""
    
  
End Sub


Private Sub cmdTurnOver_Click()
    If ClsFormAccess.AccfrmKartHesabReport = True Then
        If Tafsili > 0 Then
            Accounting.KartHesabShowDll "KolBestankaran", CStr(Tafsili), txtCustomer.Text, Right(AccountYear, 2) & "/01/01", mvarDate
        End If
    Else
        ShowDisMessage "‘„« »Â «Ì‰ «„ﬂ«‰ œ” —”Ì ‰œ«—Ìœ", 1500
    End If

End Sub

Private Sub FlxDetail_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    With FlxDetail
        If Val(FlxDetail.TextMatrix(FlxDetail.Row, 1)) <> 0 Or (fwBtnDailyHavale.Visible = True And fwBtnDailyHavale.Enabled = True) Then
            FlxDetail.TextMatrix(FlxDetail.Row, 4) = Val(FlxDetail.TextMatrix(FlxDetail.Row, 1)) * CCur(FlxDetail.TextMatrix(FlxDetail.Row, 3))
            If Col = 3 And clsStation.UpdateBuyPrice = True Then
                ReDim Parameter(1) As Parameter
                Parameter(0) = GenerateInputParameter("@Goodcode", adInteger, 4, FlxDetail.TextMatrix(FlxDetail.Row, 5))
                Parameter(1) = GenerateInputParameter("@NewBuyPrice", adInteger, 4, Val(FlxDetail.TextMatrix(FlxDetail.Row, 3)))
                RunParametricStoredProcedure "UpdateBuyPrice", Parameter
            End If
            If Col = 17 And clsStation.UpdateSellprice = True Then
                ReDim Parameter(2) As Parameter
                Parameter(0) = GenerateInputParameter("@Goodcode", adInteger, 4, FlxDetail.TextMatrix(FlxDetail.Row, 5))
                Parameter(1) = GenerateInputParameter("@SellpriceNO", adInteger, 4, 1) ' Update Sellprice Field
                Parameter(2) = GenerateInputParameter("@NewSellPrice", adInteger, 4, Val(FlxDetail.TextMatrix(FlxDetail.Row, 17)))
                RunParametricStoredProcedure "UpdateSellPrice", Parameter
            End If
        Else
            FlxDetail.RemoveItem (FlxDetail.Row)
            If FlxDetail.Rows < MaxPurchaseRows Then
                AddEmptyRow     'add row Instead of Remove
            End If
            MaxRowFlexGrid = MaxRowFlexGrid - 1
            
            frmMsg.fwlblMsg.Caption = " .ò«·«Ì „Ê—œ ‰Ÿ— «“ ·Ì”  Õ–› ‘œ "
            frmMsg.fwBtn(0).Visible = False
            frmMsg.fwBtn(1).ButtonType = flwButtonOk
            frmMsg.fwBtn(1).Caption = "ﬁ»Ê·"
            frmMsg.Show vbModal
            FlxDetail.Row = MaxRowFlexGrid     'Last Row
            FlxDetail.TopRow = FlxDetail.Rows - (MaxPurchaseRows - 1)
            txtScale.Text = ""
        
        End If
        lblNum.Caption = ""
        RefreshLables
    End With

End Sub

Private Sub FlxDetail_AfterSort(ByVal Col As Long, Order As Integer)
    With FlxDetail
        If .Rows < MaxPurchaseRows Then
            MaxRowFlexGrid = .Rows
            .Rows = MaxPurchaseRows
            .Row = MaxRowFlexGrid
        Else
            MaxRowFlexGrid = .Rows
            .Rows = .Rows + 1
            .Row = MaxRowFlexGrid
        End If
        
    End With
    
End Sub


Private Sub FlxDetail_BeforeSort(ByVal Col As Long, Order As Integer)
        
    With FlxDetail
        i = .Rows - 1
        While i >= 1
            If .TextMatrix(i, 2) = "" Then
                .RemoveItem (i)
            End If
            i = i - 1
        Wend
    End With
    
End Sub

Private Sub FlxDetail_Click()

    If MyFormAddEditMode = ViewMode Then Exit Sub
    
    
    With FlxDetail
        If .Row > 0 And .TextMatrix(.Row, 5) <> "" And .Col = 2 And Val(lblNum.Caption) <> 0 Then
            mvarGoodCode = .TextMatrix(.Row, 5)
            mvarUnitGood = .TextMatrix(.Row, 7)

            ChangeGoodquantity
            
        ElseIf .Row > 0 And .TextMatrix(.Row, 5) <> "" And (.Col = 1 Or .Col = 3 Or .Col = 10 Or .Col = 12 Or .Col = 17) Then
            .Select .Row, .Col
            .EditCell
            
        ElseIf .Row > 0 And .TextMatrix(.Row, 5) <> "" And (.Col = 18) Then
''            ReDim Parameter(1) As Parameter
''            Parameter(0) = GenerateInputParameter("@Goodcode", adInteger, 4, FlxDetail.TextMatrix(FlxDetail.Row, 5))
''            Parameter(1) = GenerateInputParameter("@Supplier", adInteger, 4, txtCustomer.Tag) '
''           ' --RunParametricStoredProcedure "UpdateSellPrice", Parameter
        End If
    
    End With
    

End Sub

Private Sub FlxDetail_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    For i = 0 To FlxDetail.Cols - 1
        SaveSetting strMainKey, "Flexgrid_Purchase", "Col" & i, FlxDetail.ColWidth(i)
    Next

End Sub

Private Sub FlexGridActive()
    
   With FlxDetail
        .Rows = MaxPurchaseRows
        .Cols = 21
        
         For i = 0 To .Cols - 1
             .ColWidth(i) = Val(GetSetting(strMainKey, "FlexGrid_Purchase", "Col" & i))
         Next i
        If .ColWidth(0) = 0 Then
            .ColWidth(0) = .Width / 12     'Row
        End If
        If .ColWidth(1) = 0 Then
            .ColWidth(1) = .Width / 8.5     'Count
        End If
        If .ColWidth(2) = 0 Then
            .ColWidth(2) = .Width / 2.6       'Good Name
        End If
        If .ColWidth(3) = 0 Then
            .ColWidth(3) = .Width / 7.5       'Fee
        End If
        If .ColWidth(4) = 0 Then
            .ColWidth(4) = .Width / 5.5       'FeeTotal
        End If
        If .ColWidth(5) = 0 Then
            .ColWidth(5) = .Width / 12       'GoodCode
        End If
        If .ColWidth(10) = 0 Then
            .ColWidth(10) = .Width / 12       'Discount rate
         End If
        If .ColWidth(13) <= 20 Then
            .ColWidth(13) = .Width / 5       'Rate
        End If
        If .ColWidth(14) <= 20 Then
            .ColWidth(14) = .Width / 5      'Inventory
        End If
        If .ColWidth(15) = 0 Then
            .ColWidth(15) = .Width / 12       'Number Of Unit
        End If
        If .ColWidth(16) = 0 Then
            .ColWidth(16) = .Width / 12       'Mojodi
        End If
        If .ColWidth(7) = 0 Then
            .ColWidth(7) = .Width / 12       'Unit
        End If
        
        If .ColWidth(17) = 0 Then
            .ColWidth(17) = .Width / 12       'SalePrice
        End If
        If .ColWidth(18) = 0 Then
            .ColWidth(18) = .Width / 12       'PreviousBuy
        End If
        
        If .ColWidth(19) = 0 Then
            .ColWidth(19) = .Width / 12       'PreviousBuy
        End If
        
        If .ColWidth(20) = 0 Then
            .ColWidth(20) = .Width / 12       'PreviousBuy
        End If
        
'        .ColHidden(0) = True
      '  .ColHidden(5) = True
        .ColHidden(6) = True
        .ColHidden(8) = True
        .ColHidden(9) = True
        .ColHidden(11) = True
        .ColHidden(12) = True

        .RowHeightMax = .Height / (MaxPurchaseRows * 1.06)   '8.2
        .RowHeightMin = .Height / (MaxPurchaseRows * 1.09)  '8.5
'''        .RowHeightMax = .Height / 8
'''        .RowHeightMin = .Height / 8.3
        
        .ColDataType(19) = flexDTBoolean
        .ColDataType(20) = flexDTBoolean
        .Row = 1
        
        MaxRowFlexGrid = 1
        
    End With

End Sub

Private Sub BtnKalaDelete_Click()

    If lblNum.Caption = "" Then
    
        lblNum.Caption = lblNum.Caption + BtnKalaDelete.Tag
        BtnKeypad(11).Enabled = False     '"%"
        BtnKeypad(10).Enabled = True      '"."
    
    Else
    
        If Left(lblNum.Caption, 1) = "-" Then
            lblNum.Caption = ""
            BtnKeypad(11).Enabled = True     '"%"
            BtnKeypad(10).Enabled = True      '"."
        End If
        
    End If

End Sub

Private Sub BtnKeypad_Click(index As Integer)

    Select Case BtnKeypad(index).Tag
    
        Case "1", "2", "3", "4", "5", "6", "7", "8", "9", "0"
        
            lblNum.Caption = lblNum.Caption + BtnKeypad(index).Tag
        
        Case "%"
        
            If lblNum.Caption <> "" Then
                lblNum.Caption = lblNum.Caption + BtnKeypad(index).Tag
''                BtnKeypad(11).Enabled = False     '"%"
            End If
            
        Case "."
        
            lblNum.Caption = lblNum.Caption + BtnKeypad(index).Tag
''            BtnKeypad(10).Enabled = False      '"."
            
    End Select
    
End Sub

Private Sub BtnMenu_MouseMove(index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    
    BtnMenu(index).ToolTipText = BtnMenu(index).Caption

End Sub


Private Sub CmdDiscountTotal_Click()

    If MyFormAddEditMode <> ViewMode Then
                
        If lblNum.Caption = "" Then
            Load frmMsg
            frmMsg.fwlblMsg.Caption = " . „»·€ ’›— —Ì«·  Œ›Ì› œ—”  ‰Ì”  "
            frmMsg.fwBtn(0).Visible = False
            frmMsg.fwBtn(1).ButtonType = flwButtonOk
            frmMsg.fwBtn(1).Caption = "ﬁ»Ê·"
            frmMsg.Show vbModal
         '   frameMenu.Enabled = True
            Exit Sub
        End If
        Load frmMsg
        If Right(lblNum.Caption, 1) <> "%" Then
            frmMsg.fwlblMsg.Caption = "¬Ì« „»·€   " & Val(lblNum.Caption) & " —Ì«· »—«Ì  Œ›Ì› „Ê—œ  «∆Ìœ «” ø "
        Else
            frmMsg.fwlblMsg.Caption = "¬Ì« „Ì“«‰   " & Val(lblNum.Caption) & " œ—’œ »—«Ì  Œ›Ì› „Ê—œ  «∆Ìœ «” ø "
        End If
        frmMsg.Show vbModal
        If modgl.mvarMsgIdx = vbNo Then
            lblNum.Caption = ""
        End If
        If Right(lblNum.Caption, 1) <> "%" Then
            txtDiscount.Text = Val(lblNum.Caption)
        Else
            txtDiscountPercent.Text = Val(lblNum.Caption)
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
            txtDiscount.Text = 0
            lblDiscountTotal = 0
            RefreshLables
        End If
        
    Else    ' View mode
    
    End If
    
    lblNum.Caption = ""

End Sub

Private Sub CmbCarryFeeTotal_Click()

    If MyFormAddEditMode <> ViewMode Then
    
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
        lblCarryFeeTotal = CLng((Val(txtSumFeeTotal.Text) * Val(txtCarryFeePercent.Text) / 100) + Val(txtCarryFee.Text)) ' + Val(txtCarryFeeCust.Text)
        lblSumPrice = CLng(Val(txtSumFeeTotal.Text) + Val(lblCarryFeeTotal.Caption) + Val(lblPackingTotal.Caption) - Val(lblDiscountTotal.Caption))
        lblSumPrice.Tag = lblSumPrice.Caption
        lblSumPrice.Caption = Format(lblSumPrice.Caption, "#,## —Ì«·")
        BtnKeypad(10).Enabled = True
        BtnKalaDelete.Enabled = True
        BtnKeypad(11).Enabled = True
            
    Else
    
    End If
    
    lblNum.Caption = ""
    
End Sub

Private Sub FlxDetail_EnterCell()
    With FlxDetail
        If .Row > 0 And .TextMatrix(.Row, 5) <> "" And (.Col = 1 Or .Col = 3 Or .Col = 10 Or .Col = 10 Or .Col = 17) Then
            
            .Select .Row, .Col
            .EditCell
        End If
    End With

End Sub

Private Sub FlxDetail_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If (IsNumeric(Chr(KeyAscii)) = False And KeyAscii <> 8) And KeyAscii <> 46 Then
        KeyAscii = 0
    End If
End Sub

Private Sub FlxDetail_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode > 32 And KeyCode < 37 Then
        KeyActi vbtxtbox, KeyCode, Shift, Me
        FlxDetail.ShowCell 1, 1
    End If
End Sub

Private Sub FlxDetail_LeaveCell()
    With FlxDetail
        If .Row > 0 And .Row < MaxRowFlexGrid And (.Col = 3 Or .Col = 17) Then
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
        If Val(FlxDetail.TextMatrix(FlxDetail.Row, 1)) <> 0 Then
            FlxDetail.TextMatrix(FlxDetail.Row, 4) = Val(FlxDetail.TextMatrix(FlxDetail.Row, 1)) * CCur(FlxDetail.TextMatrix(FlxDetail.Row, 3))
        End If
        RefreshLables
    End With
    
End Sub

Private Sub Form_Activate()
    On Error Resume Next
        
    LblAccountYear.Caption = "”«· „«·Ì :" & CInt(AccountYear)
    VarActForm = Me.Name
    SetFirstToolBar
    FlxDetail.SetFocus
    
    mvarStatus = Purchase
    Cancel

End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call PresetScreenSaver
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call PresetScreenSaver
End Sub

Public Sub Printing()

    Dim s As String
    Dim tempTxtNo As Long
    Dim ClsPrint As New Printing
    Dim tmpMyFormAddEditMode As EnumAddEditMode
    Dim ActionMode As EnumActionLog
    If tempTxtNo = 0 Then
        tempTxtNo = txtNo.Text
    End If
    
    tmpMyFormAddEditMode = MyFormAddEditMode
    If MaxRowFlexGrid >= 2 Then     ' Fich Is Not Empty
        Select Case MyFormAddEditMode
        
            Case ViewMode
                    ActionMode = Reprint
                    ClsPrint.Printing Val(txtNo.Text), mvarStationNo, tmpMyFormAddEditMode, ActionMode
                    Cancel

                    Exit Sub
''''                    Dim ArrayUbound As Integer
''''                    ReDim Parameter(5) As Parameter
''''                    Dim intIndex As Integer
''''
''''                    Parameter(0) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
''''                    Parameter(1) = GenerateInputParameter("@SystemDate", adVarWChar, 50, Mid(ClsDate.shamsi(Date), 3, 8))
''''                    Parameter(2) = GenerateInputParameter("@SystemDay", adVarWChar, 50, ClsDate.Find_DayOfWeek(Weekday(Date, vbSaturday)))
''''                    Parameter(3) = GenerateInputParameter("@SystemTime", adVarWChar, 50, Mid(str(Time), 1, 5))
''''                    Parameter(4) = GenerateInputParameter("@intFacNO", adInteger, 4, Val(txtNo.Text))
''''                    Parameter(5) = GenerateInputParameter("@Status", adInteger, 4, 1)
''''
''''                    Set RstTemp = RunParametricStoredProcedure2Rec("GetBuyFactor", Parameter)
''''                    CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\BuyFactor2.rpt"
''''                    CrystalReport1.ReportTitle = "›«ﬂ Ê— Œ—Ìœ"
''''
''''                    CrystalReport1.Destination = crptToPrinter
''''                    CrystalReport1.Destination = crptToWindow
''''
''''                    For intIndex = 0 To ArrayUbound - LBound(Parameter)
''''                        CrystalReport1.ParameterFields(intIndex) = CStr(Parameter(intIndex).Name) & ";" & CStr(Parameter(intIndex).Value) & ";" & "True"
''''                    Next intIndex
''''
''''                    CrystalReport1.RetrieveDataFiles
''''                    CrystalReport1.Action = 1
''''                    CrystalReport1.PageZoom (100)
''''
''''                    Cancel
''''                    Exit Sub
          
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
                
                If tmpMyFormAddEditMode = RefferedMode Then
                
                    If tempTxtNo <> -1 And tempReffered = 1 Then
                    
                        frmDisMsg.lblMessage = "”‰œ „—ÃÊ⁄ ‘œ"
                        frmDisMsg.Timer1.Enabled = True
                        frmDisMsg.Show vbModal
                        
                    ElseIf tempTxtNo = -1 And tempReffered = 1 Then
                    
                        frmMsg.fwlblMsg.Caption = "”‰œ „—ÃÊ⁄ ‰‘œ"
                        frmMsg.fwBtn(0).Visible = False
                        frmMsg.fwBtn(1).ButtonType = flwButtonOk
                        frmMsg.fwBtn(1).Caption = "ﬁ»Ê·"
                        frmMsg.Show vbModal
                        
                    ElseIf tempTxtNo <> -1 And tempReffered = 0 Then
                    
                        frmDisMsg.lblMessage = "”‰œ „—ÃÊ⁄Ì »«“ê—œ«‰Ì ‘œ"
                        frmDisMsg.Timer1.Enabled = True
                        frmDisMsg.Show vbModal
                        
                    ElseIf tempTxtNo = -1 And tempReffered = 0 Then
                    
                        frmMsg.fwlblMsg.Caption = "”‰œ „—ÃÊ⁄Ì »«“ê—œ«‰Ì ‰‘œ"
                        frmMsg.fwBtn(0).Visible = False
                        frmMsg.fwBtn(1).ButtonType = flwButtonOk
                        frmMsg.fwBtn(1).Caption = "ﬁ»Ê·"
                        frmMsg.Show vbModal
                        
                    End If
                    
                End If
                
                Cancel
                
        End Select
      
    Else
       Exit Sub
    End If
    
    If tempTxtNo > 0 Then
        ClsPrint.Printing tempTxtNo, clsArya.StationNo, tmpMyFormAddEditMode, ActionMode
    End If
    
    Cancel
    
End Sub

Private Sub BtnMenu_Click(index As Integer)

    Dim var1, var2 As Double
    Dim j As Double
            
    lstGoodKey.Visible = False
        
    BtnKeypad(10).Enabled = True
    BtnKalaDelete.Enabled = True
    BtnKeypad(11).Enabled = True
    
    If index = 0 Then ' search Good
    
        frmFindGoods.Show vbModal
        If mvarcode <> 0 Then
            BtnMenu(index).Tag = CStr(mvarcode)
        Else
            BtnMenu(index).Tag = ""
        End If
            
    End If
    
    If MyFormAddEditMode = ViewMode Then
    
     '   frameMenu.Enabled = True
        Exit Sub
        
    End If
    
''''    If Len(BtnMenu(Index).Tag) > 8 Then
''''
''''        lstGood.Clear
''''        Dim cnt As Integer
''''        cnt = 0
''''
''''        ReDim Parameter(3) As Parameter
''''
''''        Parameter(0) = GenerateInputParameter("@BtnNum", adInteger, 4, Index)
''''        Parameter(1) = GenerateInputParameter("@StationId", adInteger, 4, clsArya.StationNo)
''''        Parameter(2) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
''''        Parameter(3) = GenerateInputParameter("@FactorType", adInteger, 4, EnumFactorType.Purchase)
''''
''''        Set rctmp = RunParametricStoredProcedure2Rec("GetGoodList", Parameter)
''''
''''        While rctmp.EOF <> True
''''
''''           cnt = cnt + 1
''''           lstGood.AddItem rctmp.Fields("Name")
''''           lstGood.ItemData(lstGood.ListCount - 1) = rctmp.Fields("GoodCode")
''''           rctmp.MoveNext
''''
''''        Wend
''''
''''        rctmp.Close
''''
''''
''''        If (cnt > 10) Then
''''            lstGood.Height = 4000 'cnt * 200
''''        Else
''''            lstGood.Height = cnt * 400
''''        End If
''''
''''        ArrangeLstGood (Index)
''''        lstGood.Visible = True
''''
''''        frameMenu.Enabled = True
''''        Me.lstGood.SetFocus
''''
''''        If Me.lstGood.ListCount > 0 Then
''''
''''            Me.lstGood.ListIndex = 0
''''
''''        End If
''''
''''        Exit Sub
''''
''''    End If
''''
''''
''''    If GetGoodCode(Val(BtnMenu(Index).Tag)) = True Then
''''
''''        ChangeGoodquantity
''''
''''    End If
    
End Sub
Public Function GetGoodCode(Code As Double)
    
    If Code = 0 Then Exit Function
        
'    If mvarStatus = fromStore And cmbDestInventory.ListIndex = -1 Then
'        frmMsg.fwlblMsg.Caption = "«‰»«— „›’œ »«Ìœ «‰ Œ«» ‘Êœ"
'        frmMsg.fwBtn(0).Visible = False
'        frmMsg.fwBtn(1).ButtonType = flwButtonOk
'        frmMsg.fwBtn(1).Caption = "ﬁ»Ê·"
'        frmMsg.Show vbModal
'        Exit Function
'    End If
    If cmbInventory.ListIndex <> -1 And cmbDestInventory.ListIndex <> -1 Then
        If cmbInventory.ItemData(cmbInventory.ListIndex) = cmbDestInventory.ItemData(cmbDestInventory.ListIndex) Then
            frmMsg.fwlblMsg.Caption = "«‰»«—Â« »«Ìœ »« Â„ „ ›«Ê  »«‘‰œ"
            frmMsg.fwBtn(0).Visible = False
            frmMsg.fwBtn(1).ButtonType = flwButtonOk
            frmMsg.fwBtn(1).Caption = "ﬁ»Ê·"
            frmMsg.Show vbModal
            Exit Function
        End If
    End If
    
    Dim ExistfromStore, ExisttoStore, ReturnValue As Boolean
    ReturnValue = False
    ExistfromStore = False
    ExisttoStore = False
    
    ReDim Parameter(4) As Parameter
    Parameter(0) = GenerateInputParameter("@Code", adInteger, 4, Code)
    Parameter(1) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
    Parameter(2) = GenerateInputParameter("@StationId", adInteger, 4, clsArya.StationNo)
    Parameter(3) = GenerateInputParameter("@Flag", adBoolean, 1, 0)
    Parameter(4) = GenerateInputParameter("@AccountYear", adSmallInt, 2, AccountYear)
    Set rctmp = RunParametricStoredProcedure2Rec("Get_Good_Code", Parameter)
    
    If Not (rctmp.BOF Or rctmp.EOF) Then
        
        If cmbDestInventory.ListIndex <> -1 Then
            DestInventoryNo = cmbDestInventory.ItemData(cmbDestInventory.ListIndex)
        Else
            ExisttoStore = True
        End If
        Do While rctmp.EOF <> True
        
    ''''           InventoryNo = rctmp.Fields("InventoryNo").Value
            mvarInventoryNo = cmbInventory.ItemData(cmbInventory.ListIndex) ' rctmp.Fields ("InventoryNo")
            If mvarInventoryNo = rctmp.Fields("InventoryNo").Value Then
                  
                mvarGoodCode = rctmp.Fields("Code")
                mvarUnitGood = rctmp.Fields("Unit")
                mvarGoodName = rctmp.Fields("Name")
                mvarGoodWeight = rctmp.Fields("Weight")
                mvarNumberOfUnit = rctmp.Fields("NumberOfUnit")
                '   mvarDisCount = rctmp.Fields("Discount")   Only For Sale
                mvarMojodi = rctmp.Fields("Mojodi")
                mvarSellPrice = rctmp.Fields("SellPrice")
               If chKTax = True Then
                  mvarDuty = True
                  mvarTax = True
               Else
                  mvarDuty = rctmp.Fields("DutySale")
                  mvarTax = rctmp.Fields("TaxSale")
               End If
                If mvarStatus = 6 Then
                    If clsStation.FromStoreFee = 0 Then
                        ReDim Parameter(3) As Parameter
                        Parameter(0) = GenerateInputParameter("@GoodCode", adInteger, 4, mvarGoodCode)
                        Parameter(1) = GenerateInputParameter("@DateAfter", adWChar, 20, Mid(clsDate.shamsi(Date), 3, 2) & "/01" & "/01")
                        Parameter(2) = GenerateInputParameter("@DateBefore", adWChar, 20, txtDate.Text)
                        Parameter(3) = GenerateInputParameter("@Flag", adBoolean, 1, 0)
                        Set RstTemp = RunParametricStoredProcedure2Rec("AverageCalculateBuyPrice", Parameter)
    
                        mvarBuyPrice = RstTemp!AverageBuyPrice
                        Set RstTemp = Nothing
                    ElseIf clsStation.FromStoreFee = 1 Then
                        mvarBuyPrice = rctmp.Fields("BuyPrice").Value
                    ElseIf clsStation.FromStoreFee = 2 Then
                        mvarBuyPrice = rctmp.Fields("SellPrice").Value
                    ElseIf clsStation.FromStoreFee = 3 Then
                        mvarBuyPrice = rctmp.Fields("SellPrice2").Value
                    ElseIf clsStation.FromStoreFee = 4 Then
                        mvarBuyPrice = rctmp.Fields("SellPrice3").Value
                    ElseIf clsStation.FromStoreFee = 5 Then
                        mvarBuyPrice = rctmp.Fields("SellPrice4").Value
                    ElseIf clsStation.FromStoreFee = 6 Then
                        mvarBuyPrice = rctmp.Fields("SellPrice5").Value
                    ElseIf clsStation.FromStoreFee = 7 Then
                        mvarBuyPrice = rctmp.Fields("SellPrice6").Value
                    End If
                Else
                    mvarBuyPrice = rctmp.Fields("BuyPrice").Value
                End If
                ExistfromStore = True
            End If
            If DestInventoryNo <> 0 Then
                If DestInventoryNo = rctmp.Fields("InventoryNo").Value Then
                    ExisttoStore = True
                End If
            End If

            rctmp.MoveNext
      
        Loop
    End If
    
    If ExistfromStore = False Then
        frmDisMsg.lblMessage.Caption = " «Ì‰ ò«·« »—«Ì «‰»«— „»œ«  ⁄—Ì› ‰‘œÂ «”   "
        frmDisMsg.Timer1.Interval = 1000
        frmDisMsg.Timer1.Enabled = True
        frmDisMsg.Show vbModal
        ReturnValue = False
    End If
    If ExisttoStore = False Then
        frmDisMsg.lblMessage.Caption = " «Ì‰ ò«·« »—«Ì «‰»«— „ﬁ’œ  ⁄—Ì› ‰‘œÂ «”  "
        frmDisMsg.Timer1.Interval = 1000
        frmDisMsg.Timer1.Enabled = True
        frmDisMsg.Show vbModal
        ReturnValue = False
    End If
    rctmp.Close
    If (ExistfromStore = True And ExisttoStore = True) Then ReturnValue = True
    GetGoodCode = ReturnValue
End Function

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Exit_Keypress_Flag = False

    KeyActi vbtxtbox, KeyCode, Shift, Me
    If KeyCode < 48 And KeyCode <> 13 And KeyCode <> 27 And KeyCode <> 8 Then
       Exit_Keypress_Flag = True
       Exit Sub
    End If
    
   If (Shift = 0 And KeyCode >= vbKeyA) And (Shift = 0 And KeyCode <= vbKeyZ) Then
         If mvarbarcode Then
             lblBarCode = lblBarCode & ChrW$(KeyCode)
             If Len(lblBarCode) > clsStation.BarcodeLengh Then
                 lblBarCode = ""
                 mvarbarcode = False
             End If
             Exit_Keypress_Flag = True
             Exit Sub
        End If
    End If
    mvarKeyCode = KeyCode
    MvarShiftKey = Shift
    
    Select Case Shift
    
        Case 0
            Select Case KeyCode
                

                Case vbKeyF3
                
                    If MyFormAddEditMode = ViewMode Then
                       Me.Edit
                    End If
                    Exit_Keypress_Flag = True
                    
                Case vbKeyF4

                   Me.FindCust
                    Exit_Keypress_Flag = True
                   
                Case vbKeyF6
                
                    'Me.Printing
                    Exit_Keypress_Flag = True
                                        
                Case vbKeyF10
                
                    Exit_Keypress_Flag = True
                
                Case vbKeyF11
                
                    Exit_Keypress_Flag = True
                    
                Case vbKeySubtract, 189 '-
                    
                    BtnKalaDelete_Click
                    Exit_Keypress_Flag = True
                    
                Case vbKeyDecimal, 190          ' .
                
                    If BtnKeypad(10).Enabled Then
                        BtnKeypad_Click (10)
                    End If
                    Exit_Keypress_Flag = True
                    
                Case vbKeyDivide, 191 '/ Barcode
                
                    If mvarbarcode = False Then
                        mvarbarcode = True
                    Else
                        barcode
                    End If
                    Exit_Keypress_Flag = True
                    
                Case vbKey0 To vbKey9
                
                    If Not (mvarbarcode) Then
                        BtnKeypad_Click (Chr(KeyCode))
                    Else
                        lblBarCode = lblBarCode & Chr(KeyCode)
                    End If
                    Exit_Keypress_Flag = True
                
                Case vbKeyNumpad0 To vbKeyNumpad9
                    
                    KeyCode = KeyCode - 48
                    If Not (mvarbarcode) Then
                        BtnKeypad_Click (Chr(KeyCode))
                    Else
                        lblBarCode = lblBarCode & Chr(KeyCode)
                    End If
                    KeyCode = KeyCode + 48
                    Exit_Keypress_Flag = True
                   
                Case vbKeyBack
                
                    If Len(Trim(lblNum.Caption)) >= 1 Then
                        If Right(lblNum.Caption, 1) = "." Then
                           BtnKeypad(10).Enabled = True
                        End If
                        If Right(lblNum.Caption, 1) = "%" Then
                           BtnKeypad(11).Enabled = True
                        End If
                        lblNum.Caption = Left(lblNum.Caption, Len(Trim(lblNum.Caption)) - 1)
                    End If
                    If Len(Trim(lblBarCode.Caption)) >= 1 Then
                        lblBarCode.Caption = Left(lblBarCode.Caption, Len(Trim(lblBarCode.Caption)) - 1)
                    End If
                    
                    Exit_Keypress_Flag = True
                Case vbKeyEscape
                    If Len(txtBarcode.Text) > 0 Then Exit Sub
                    If (Me.lstGoodKey.Visible = False) Then
                        Cancel
                        HideLstBoxes KeyCode
                    Else
                        HideLstBoxes KeyCode
                        Exit Sub
                    End If
                    Exit_Keypress_Flag = True
                        
                Case vbKeyReturn
                    If Len(txtBarcode.Text) > 0 Then Exit Sub
                    If Me.lstGoodKey.Visible = False Then
                        If Not MyFormAddEditMode = ViewMode Then   ' add & Edit
                           Update
                        End If
                    End If
                    Exit_Keypress_Flag = True
                                                
                End Select
                
        Case 1     'With Shift Key
           
            Select Case KeyCode
            
                                   
                Case vbKey5, vbKeyNumpad5    '%
                    
                    BtnKeypad_Click (11)
                    
                    Exit_Keypress_Flag = True
                
            End Select
   
        Case 2
        
            Select Case KeyCode
                                
                Case vbKeyF3       'Discount
                    
                    CmdDiscountTotal_Click
                    Exit_Keypress_Flag = True
                
                Case vbKeyF4      'Caree Fee
                    
                    CmbCarryFeeTotal_Click
                    Exit_Keypress_Flag = True
                
                Case vbKeyF5       'Service
                    
                    Exit_Keypress_Flag = True
                
                Case vbKeyF6       'Packing
                    
                    cmbPackingTotal_Click
                    Exit_Keypress_Flag = True
                
                Case vbKeyF7       ' Good Find
                    
                    BtnMenu_Click (0)
                    Exit_Keypress_Flag = True
                    
                Case vbKeyF8
                
                    Exit_Keypress_Flag = True
'                    Call OpenCashDrawer
                    
    '            Case 119:  ' Drawer Open In Modgl Routine
    '                    KeyActi vbtxtbox, 119, 2, me, mycls
    '            Case 120:        'Phone Book in Modgl Routine
    '                    KeyActi vbtxtbox, 120, 2, me, mycls
                Case 221
                
                    lblServePlace_MouseUp 1, 0, 43, 7
                    Exit_Keypress_Flag = True
             
            End Select
    End Select

End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
     If Exit_Keypress_Flag = True Then Exit Sub
   If textDescription = True Then Exit Sub
     
     Dim j As Double
     Dim mvarstr As String
     If mvarbarcode Then
         Exit Sub
     End If
     
     
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

End Sub

Private Sub Form_Load()
   
   On Error GoTo ErrorHandler
   If ClsFormAccess.frmPurchase = False Then
        Unload Me
        Exit Sub
    End If
    
    If intVersion = Min Or intVersion = Normal Then
        ShowDisMessage "›«ﬂ Ê—Œ—Ìœ œ— ‰”ŒÂ Â«Ì ÅÌ‘—› Â Ê »«·« — ÊÃÊœ œ«—œ", 1500
        Unload Me
        Exit Sub
    End If
    
   If ClsFormAccess.frmSavePurchase = False And ClsFormAccess.frmHavaleh = False And ClsFormAccess.frmTempReceived = False And ClsFormAccess.frmLosses = False Then
        ShowDisMessage "œ” —”Ì »Â Õ«· Â«Ì „Œ ·› œ—Ê‰ ›«ò Ê— Œ—Ìœ »” Â «” ", 2000
        Unload Me
        Exit Sub
    End If
    
    If clsStation.PurchaseRows = 0 Then
        MaxPurchaseRows = 8
    Else
        MaxPurchaseRows = Val(clsStation.PurchaseRows) + 1
    End If
    
    FlexGridActive
    
'    fwScrollTextCust.BackColor = Me.BackColor
    ''fwScrollTextCust.Visible = False
    
    mvarServePlace = EnumServePlace.Salon
    mVarOrderType = inPerson
    
    UpdatelblServePlace
    
    txtSumWeightTotal.Visible = True
    lblSumWeightTotal.Visible = True
    txtSumCountWeight.Visible = True
    lblSumCountWeight.Visible = True
    lblSumCountNo.Visible = True
    
    
    If rctmp.State = 1 Then rctmp.Close
    cmbDestination.Clear
    
    Set rctmp = RunStoredProcedure2RecordSet("Get_All_tblPub_Destination")
    
    If Not (rctmp.EOF = True And rctmp.BOF = True) Then
        Do While rctmp.EOF <> True
            cmbDestination.AddItem rctmp.Fields("nvcDestination")
            cmbDestination.ItemData(cmbDestination.ListCount - 1) = Trim(rctmp.Fields("DestinationId"))
            rctmp.MoveNext
        Loop
         
    End If
    rctmp.Close
    
    If rctmp.State = 1 Then rctmp.Close
    cmbInventory.Clear
    
    ReDim Parameter(1) As Parameter
    Parameter(0) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
    Parameter(1) = GenerateInputParameter("@Type", adInteger, 4, 1) 'All Inventory
    Set rctmp = RunParametricStoredProcedure2Rec("GetInventory", Parameter)
'    Parameter(1) = GenerateInputParameter("@Branch", adInteger, 4, CurrentBranch)
'    Set rctmp = RunParametricStoredProcedure2Rec("GetInventory_Branch", Parameter)
    
    If Not (rctmp.EOF = True And rctmp.BOF = True) Then
        Do While rctmp.EOF <> True
            cmbInventory.AddItem rctmp.Fields("Description")
            cmbInventory.ItemData(cmbInventory.ListCount - 1) = Val(rctmp.Fields("InventoryNo"))
            rctmp.MoveNext
        Loop
         
    End If
    cmbInventory.ListIndex = 0
    rctmp.Close
    If rctmp.State = 1 Then rctmp.Close
    cmbDestInventory.Clear
    
    ReDim Parameter(1) As Parameter
    Parameter(0) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
    Parameter(1) = GenerateInputParameter("@Type", adInteger, 4, 1)  ' All Inventory
    Set rctmp = RunParametricStoredProcedure2Rec("GetInventory", Parameter)
    
    If Not (rctmp.EOF = True And rctmp.BOF = True) Then
        Do While rctmp.EOF <> True
            cmbDestInventory.AddItem rctmp.Fields("Description")
            cmbDestInventory.ItemData(cmbDestInventory.ListCount - 1) = Val(rctmp.Fields("InventoryNo"))
            rctmp.MoveNext
        Loop
         
    End If
    cmbDestInventory.ListIndex = 0
    rctmp.Close
    
    FillBranch
    
    With CmbStatus
    
        .Clear
        ReDim Parameter(1) As Parameter
        Parameter(0) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
        Parameter(1) = GenerateInputParameter("@ReturnType", adInteger, 4, 1)
        Set rctmp = RunParametricStoredProcedure2Rec("Get_All_tStatusType", Parameter)
        While rctmp.EOF <> True
            If rctmp!intStatusNo = 6 And ClsFormAccess.frmHavaleh = False Then
            ElseIf rctmp!intStatusNo = 7 And ClsFormAccess.frmHavaleh = False Then
            ElseIf rctmp!intStatusNo = 9 And ClsFormAccess.frmTempReceived = False Then
            ElseIf rctmp!intStatusNo = 1 And ClsFormAccess.frmSavePurchase = False Then
            ElseIf rctmp!intStatusNo = 4 And ClsFormAccess.frmSavePurchase = False Then
            ElseIf rctmp!intStatusNo = 3 And ClsFormAccess.frmLosses = False Then
            
            Else
                .AddItem rctmp!NvcDescription
                .ItemData(.ListCount - 1) = rctmp!intStatusNo
            End If
            rctmp.MoveNext
        
        Wend

''''        .AddItem "—”Ìœ „Êﬁ "
''''        .ItemData(.ListCount - 1) = 9
        If rctmp.State <> 0 Then rctmp.Close
        If .ListCount = 0 Then ShowDisMessage "œ” —”Ì »Â Õ«· Â«Ì „Œ ·› œ—Ê‰ ›«ò Ê— Œ—Ìœ »” Â «” ", 2000: Unload Me: Exit Sub
        .ListIndex = 0
    End With
    
''''    CmbStatus.AddItem "›«ﬂ Ê— Œ—Ìœ"
''''    CmbStatus.AddItem "÷«Ì⁄« "
''''    CmbStatus.AddItem "»—ê‘  «“ Œ—Ìœ"
''''    CmbStatus.ListIndex = 0
    
    PortClose
    
    GetProperController
    
  '  Me.ValueBtnMenu
    
    Call ColorSetting
    
    ChangeLanguage
    
    CenterTop Me
    
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

    BlnFormLoaded = True

Exit Sub
ErrorHandler:

    MsgBox err.Description & "Form_Load"
    
End Sub
Private Sub PortClose()
    
    For i = mscSerial.LBound To mscSerial.UBound
        If Me.mscSerial(i).PortOpen Then
            Me.mscSerial(i).PortOpen = False
        End If
    Next i
    
End Sub


Private Sub fwBtnDailyHavale_Click()
    Dim SumTotal As Currency
    If FlxDetail.TextMatrix(1, 5) <> "" Then Exit Sub

    Dim AutoHavale As Long
    ReDim Parameter(0) As Parameter
    Parameter(0) = GenerateOutputParameter("@AutoHavale", adInteger, 4)
    AutoHavale = RunParametricStoredProcedure("Get_AutoHavale", Parameter)
    If Val(AutoHavale) = 1 Then
        ShowDisMessage "”Ì” „ œ— Õ«· «” ›«œÂ «“ ÕÊ«·Â « Ê„« Ìﬂ „Ì »«‘œ", 2000
        Exit Sub
    End If

    Dim Rst As New ADODB.Recordset
  
    BitAutoHavale = True
    frmHavaleDate.Show vbModal
  
    ReDim Parameter(5) As Parameter
    Parameter(0) = GenerateInputParameter("@Branch", adInteger, 4, CurrentBranch)
    Parameter(1) = GenerateInputParameter("@InventoryNo", adInteger, 4, cmbInventory.ItemData(cmbInventory.ListIndex))
'    Parameter(2) = GenerateInputParameter("@DesInventoryNo", adInteger, 4, cmbDestInventory.ItemData(cmbDestInventory.ListIndex))
    Parameter(2) = GenerateInputParameter("@AccountYear", adSmallInt, 2, AccountYear)
    Parameter(3) = GenerateInputParameter("@Status", adInteger, 4, 2)
    Parameter(4) = GenerateInputParameter("@FromDate", adVarWChar, 8, FromDate)
    Parameter(5) = GenerateInputParameter("@ToDate", adVarWChar, 8, ToDate)
'
    Set Rst = RunParametricStoredProcedure2Rec("Get_DailyGoodForHavale", Parameter)
    Dim ii As Integer
    
    If Rst.BOF = True And Rst.EOF = True Then
        ShowDisMessage "”‰œÌ »—«Ì ÕÊ«·Â ÊÃÊœ ‰œ«—œ", 2000
        Set Rst = Nothing
        Exit Sub
    End If
    
    With FlxDetail
        ii = 0
        Do While Not (Rst.EOF)
            ii = ii + 1
            FlxDetail.TextMatrix(ii, 0) = ii 'Number
            FlxDetail.TextMatrix(ii, 1) = Val(Format(Rst!amount, "##.###"))  'Rst!Amount
            FlxDetail.TextMatrix(ii, 2) = Rst!Name 'GoodName
            FlxDetail.TextMatrix(ii, 3) = Rst!FeeUnit
            FlxDetail.TextMatrix(ii, 4) = Val(Format(Rst!amount * Rst!FeeUnit, "##")) ' rst!FeeTotal
            FlxDetail.TextMatrix(ii, 5) = Rst!GoodCode
            FlxDetail.TextMatrix(ii, 6) = Rst!Weight ' rst!WeightUnit
            FlxDetail.TextMatrix(ii, 7) = Rst!Unit
            FlxDetail.TextMatrix(ii, 8) = 1
            FlxDetail.TextMatrix(ii, 9) = ""
            FlxDetail.TextMatrix(ii, 10) = 0
            FlxDetail.TextMatrix(ii, 11) = 1
            FlxDetail.TextMatrix(ii, 12) = ""
            FlxDetail.TextMatrix(ii, 13) = Rst!intInventoryNo
            FlxDetail.TextMatrix(ii, 14) = IIf(IsNull(Rst!DestInventoryNo), "", Rst!DestInventoryNo)
            FlxDetail.TextMatrix(ii, 15) = Rst!NumberOfUnit
            ''TmpGoodDiscount = TmpGoodDiscount + (Rst!Discount * Rst!Amount * Rst!FeeUnit / 100)
            FlxDetail.TextMatrix(ii, 16) = 0
            SumTotal = SumTotal + CCur(FlxDetail.TextMatrix(ii, 4))
            If SumTotal <> 0 Then
                SumTotal = Val(Format(SumTotal, "##"))
            End If
            Rst.MoveNext
            
            MaxRowFlexGrid = ii + 1
            If ii >= FlxDetail.Rows - 2 And Rst.EOF = False Then
                AddEmptyRow
            End If

        Loop
''
''        FlxDetail.Row = MaxRowFlexGrid - 1
''        mvarServePlace = FlxDetail.TextMatrix(MaxRowFlexGrid - 1, 8)
''    End If
End With
    LblSubTotal.Caption = SumTotal
    lblSumPrice.Caption = SumTotal
    lblSumPrice.Tag = lblSumPrice.Caption
    lblSumPrice.Caption = Format(lblSumPrice.Caption, "#,## —Ì«·")

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

Private Sub mscSerial_OnComm(index As Integer)
    Select Case mscSerial(index).CommEvent
    
        Case comEvReceive   ' Received RThreshold # of
            
            mscSerial(index).RThreshold = 0
            
            Select Case DeviceType(index)
            End Select
            mscSerial(index).RThreshold = RThreshold(index)
            If mscSerial(index).PortOpen = False Then
                mscSerial(index).PortOpen = True
            End If
    End Select
End Sub

Sub GetProperController()
    
    On Error GoTo Err1
    
    
    ReDim Parameter(0) As Parameter
    Parameter(0) = GenerateInputParameter("@StationId", adInteger, 4, clsArya.StationNo)
    
    Set rctmp = RunParametricStoredProcedure2Rec("Get_DeviceSetting", Parameter)
    
    i = 1
    
    If Not (rctmp.BOF Or rctmp.EOF) Then
        
        If rctmp.Fields("DeviceTypeCode").Value = 1 And rctmp.Fields("PortNo").Value <> 11 Then    ' Not Lpt Port
            
            mscSerial(i).CommPort = rctmp.Fields("PortNo").Value
            mscSerial(i).Settings = rctmp.Fields("BaudRate").Value & ",N,8,1"
            DeviceCode(i) = rctmp.Fields("DeviceCode").Value
            DeviceType(i) = rctmp.Fields("DeviceTypeCode").Value
            RThreshold(i) = rctmp.Fields("RThreshold").Value
            mscSerial(i).InBufferSize = rctmp.Fields("BufferSize").Value
            mscSerial(i).RThreshold = rctmp.Fields("RThreshold").Value
            
            If Not (mscSerial(i).PortOpen) Then
                
                mscSerial(i).PortOpen = True
            
            End If
            
        
        End If
    End If
    If i <= mscSerial.Count Then
    
        i = i + 1
    End If
    
   
    Exit Sub
    
Err1:

    MsgBox err.Description, , err.Source
  
End Sub

Private Sub Form_Unload(Cancel As Integer)

    modgl.mvarDeleteMsg = ""
'    ClearDataFlexGrid
    
    BlnFormLoaded = False
    
    VarActForm = ""
    
    Set mdifrm.FileCls = Nothing
    Set clsDate = Nothing
    Set ClsCnvKeyBoard = Nothing
    
    Set rctmp = Nothing
    If PosConnection.State = adStateOpen Then PosConnection.Close: Set PosConnection = Nothing
    Set RstTemp = Nothing
    
    AllButton vbOff, True
    
    Unload frmFindGoods

    SaveSetting strMainKey, Me.Name, "Left", Me.Left
    SaveSetting strMainKey, Me.Name, "Top", Me.Top


    
End Sub


Public Sub FirstKey()

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
    
    If MyFormAddEditMode = AddMode Then
        LastKey
        Exit Sub
    End If
    MyFormAddEditMode = ViewMode  'View Mode
    DefaultValueLables
    
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

    If MyFormAddEditMode = AddMode Then
        LastKey
        Exit Sub
    End If
    MyFormAddEditMode = ViewMode  'View Mode
    DefaultValueLables
    
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
    
    GetDataDetail
    
    RefreshLables
    SetFirstToolBar
    
End Sub

Public Sub Add()

    Dim AutoValue As Integer
    HideLstBoxes 27
    intTempFich = 0
    mvarServePlace = Salon
    boolPayment = 0
    BtnKeypad(10).Enabled = True
   
    ClearDataFlexGrid
    txtDate.Text = mvarDate ' Right(clsDate.shamsi(Date), 8)
    txtRegDate.Text = Right(clsDate.shamsi(Date), 8)
    Me.Number
    DefaultValueLables       'Set Default Value Lables
    Me.ValueLabel
    For i = 2 To 5
        Me.StatusBar.Panels(i).Bevel = sbrInset
    Next i
    
    For i = 0 To cmbBranch.ListCount - 1
        cmbBranch.ListIndex = i
        If CurrentBranch = cmbBranch.ItemData(cmbBranch.ListIndex) Then
            Exit For
        End If
    Next
    
    ArrowkeyStatusbar LastRecord        'Display 5 Last Fich

    Me.txtRecursive = 0
    fwlblRecursive.Visible = False
    'fwScrollTextCust.Visible = False
    lblNum = ""
    For i = 0 To cmbInventory.ListCount - 1
        If cmbInventory.ItemData(i) = clsStation.PurchaseInventoryDefault Then
            cmbInventory.ListIndex = i
            i = 0
            Exit For
        End If
    Next i
'''    cmbDestInventory.ListIndex = 0
    
    MyFormAddEditMode = AddMode       'Add Mode
    SetFirstToolBar
    FlxDetail.Select 1, 1
    txtDescription.Text = ""
    cmbDestInventory.ListIndex = -1
    DestInventoryNo = 0
    
    cmbInventory_Click
    LblSubTotal.Caption = 0
    textDescription = False
    BitAutoHavale = False
End Sub

Public Sub Cancel()

    CmbStatus_Click
    MyFormAddEditMode = AddMode
    Add
    
End Sub

Public Sub Edit()
    
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
    
    If mvarStatus = TempRecieved And BitTempReceived = True Then
        ShowDisMessage "—”Ìœ „Êﬁ  ﬁ»·« »Â —”Ìœ œ«∆„  »œÌ· ‘œÂ Ê ﬁ«»· ÊÌ—«Ì‘ ‰Ì” ", 1500: Exit Sub
    End If
    If mvarStatus = toStore And cmbDestInventory.ListIndex <> -1 Then
        frmMsg.fwlblMsg.Caption = "«Ì‰ —”Ìœ »«Ìœ «“ ÿ—Ìﬁ ÕÊ«·Â „—»ÊÿÂ ÊÌ—«Ì‘ ‘Êœ  "
        frmMsg.fwBtn(0).Visible = False
        frmMsg.fwBtn(1).ButtonType = flwButtonOk
        frmMsg.fwBtn(1).Caption = "ﬁ»Ê·"
        frmMsg.Show vbModal
        MyFormAddEditMode = ViewMode
        SetFirstToolBar

        Exit Sub
    End If
    
    If Me.txtRecursive = 1 Then
        If (ClsFormAccess.RefferInvoice = False) Or (ClsFormAccess.RefferedAllStationsFactors = False And (mvarCurUserNo <> dblFichUser)) Then
            MyFormAddEditMode = ViewMode
        Else
            frmMsg.fwlblMsg.Caption = " . ›Ì‘ „—ÃÊ⁄Ì ﬁ«»· «’·«Õ ‰Ì”  "
            frmMsg.fwBtn(0).Visible = False
            frmMsg.fwBtn(1).ButtonType = flwButtonOk
            frmMsg.fwBtn(1).Caption = "ﬁ»Ê·"
            frmMsg.Show vbModal
            MyFormAddEditMode = RefferedMode
        
        End If
        
    ElseIf ClsFormAccess.EditPurchase <> True Then
    
        frmMsg.fwlblMsg.Caption = " . ‘„« «Ã«“Â «’·«Õ ò—œ‰ ›Ì‘ —« ‰œ«—Ìœ "
        frmMsg.fwBtn(0).Visible = True
        frmMsg.Show vbModal
        MyFormAddEditMode = ViewMode
    
    Else
    
        MyFormAddEditMode = EditMode
        
    End If
    
    UpdatetxtCustomer
    SetFirstToolBar
    
End Sub
Public Sub Delete()
If MyFormAddEditMode = AddMode Then Exit Sub
 If MaxRowFlexGrid <= 1 Then
       Exit Sub
 End If
    
    If FWChkAcc.Value = True Then
        ShowDisMessage "”‰œ Õ”«»œ«—Ì »—«Ì «Ì‰ ›Ì‘ ’«œ— ‘œÂ Ê «„ﬂ«‰ Õ–› ›Ì‘ ÊÃÊœ ‰œ«—œ", 2000
        Exit Sub
    End If
    If mvarStatus = toStore Then
        frmMsg.fwlblMsg.Caption = " —”Ìœ «“ ÿ—Ìﬁ ÕÊ«·Â „—»ÊÿÂ ﬁ«»· Õ–› „Ì »«‘œ  "
        frmMsg.fwBtn(0).Visible = False
        frmMsg.fwBtn(1).ButtonType = flwButtonOk
        frmMsg.fwBtn(1).Caption = "ﬁ»Ê·"
        frmMsg.Show vbModal
        Exit Sub
    End If
    Select Case clsStation.Language
        Case 0
            frmMsg.fwlblMsg.Caption = "¬Ì« „ÿ„∆‰Ìœ „Ì ŒÊ«ÂÌœ ›Ì‘ " & "'" & Val(FWLed1.Value) & "'" & " —« Õ–› ﬂ‰Ìœø"
            frmMsg.fwBtn(0).Caption = "»·Ì"
            frmMsg.fwBtn(1).Caption = "ŒÌ—"
        Case 1
            frmMsg.fwlblMsg.Caption = "You are going to delete '" & Val(FWLed1.Value) & "'" + vbNewLine + "Are you sure ?"
            frmMsg.fwBtn(0).Caption = "Yes"
            frmMsg.fwBtn(1).Caption = "No"
            frmMsg.fwlblMsg.Alignment = vbLeftJustify
    End Select
    
    frmMsg.Show vbModal
    
    If modgl.mvarMsgIdx = vbNo Then
        Exit Sub
    End If
    
    ReDim Parameter(4) As Parameter
    Parameter(0) = GenerateInputParameter("@No", adBigInt, 8, Val(txtNo.Text))
    Parameter(1) = GenerateInputParameter("@Status", adInteger, 4, mvarStatus)
    Parameter(2) = GenerateInputParameter("@AccountYear", adSmallInt, 2, AccountYear)
    Parameter(3) = GenerateInputParameter("@Branch", adInteger, 4, CurrentBranch)
    Parameter(4) = GenerateOutputParameter("@Result", adInteger, 4)
    Dim Result As Integer
    Result = RunParametricStoredProcedure("Delete_tFacmd", Parameter)
    
    If Result = 0 Then
    
        Select Case clsStation.Language

            Case 0
                frmMsg.fwlblMsg.Caption = "„‘ò·Ì œ—Õ–› «Ì‰ ›Ì‘ ÊÃÊœ œ«—œ ‘„« ‰„Ì  Ê«‰Ìœ «Ì‰ ›Ì‘ —« Õ–› ò‰Ìœ"
                frmMsg.fwBtn(0).Caption = "ﬁ»Ê·"
            Case 1
                frmMsg.fwlblMsg.Caption = "There are some factors related to this good , you cant delete it"
                frmMsg.fwBtn(0).Caption = "Ok"
                frmMsg.fwlblMsg.Alignment = vbLeftJustify
        End Select
        frmMsg.fwBtn(1).Visible = False
        frmMsg.fwBtn(0).ButtonType = flwButtonOk
        frmMsg.Show vbModal
    
    Else
    
        Select Case clsStation.Language
            Case 0
                frmMsg.fwlblMsg.Caption = "‘„« Ìò ›Ì‘ —« Õ–› ò—œÌœ"
                frmMsg.fwBtn(0).Caption = "ﬁ»Ê·"
            Case 1
                frmMsg.fwlblMsg.Caption = "You have deleted one good"
                frmMsg.fwBtn(0).Caption = "Ok"
                frmMsg.fwlblMsg.Alignment = vbLeftJustify
        End Select
        
        frmMsg.fwBtn(1).Visible = False
        frmMsg.fwBtn(0).ButtonType = flwButtonOk
        frmMsg.Show vbModal
        
    End If
    
    Add
    
End Sub

Public Function Update() As Long

If mvarStatus = StandardHavale Then
     ShowDisMessage "»—«Ì À»  «” «‰œ«—œ ÕÊ«·Â «“ ›—„ ¬‰«·Ì“ ò«·« «” ›«œÂ ‰„«ÌÌœ", 1500
     Update = -1
     Exit Function
End If

Dim SanadNo As Long
Dim Status As Integer

If BitAutoHavale = True Then
    If MaxRowFlexGrid = 1 Then Exit Function
    Update = InsertAutoHavale
    If clsArya.ExternalAccounting = True And Update > 0 Then
        If intVersion <> Diamond Then
            ShowDisMessage "«„ò«‰  Ê·Ìœ « Ê„« Ìò ”‰œ Õ”«»œ«—Ì ÕÊ«·Â ›ﬁÿ œ— Ê—é‰ «·„«” ÊÃÊœ œ«—œ", 1500
        Else
            Status = mvarStatus
            SanadNo = Accounting.Insert_CustomerSale(AddMode, 0, txtDate.Text, Tafsili, txtCustomer.Text, Update, CCur(lblSumPrice.Tag), Val(LblSubTotal), Val(lblDiscountTotal), Val(lblCarryFeeTotal), Val(lblPackingTotal), Val(lblTaxTotal), Val(LblDutyTotal), Status, "", "", Tafsili_2)
        End If
        Cancel
    End If
    Cancel
    Exit Function
End If

'If mvarStatus = TempRecieved Then
'    frmMsg.fwlblMsg.Caption = "—”Ìœ „Êﬁ  »Â œ«∆„  »œÌ· „Ì ‘Êœ "
'    frmMsg.fwBtn(0).Visible = True
'    frmMsg.fwBtn(1).Visible = True
'    frmMsg.fwBtn(0).ButtonType = flwButtonOk
'    frmMsg.fwBtn(1).ButtonType = flwButtonCancel
'    frmMsg.fwBtn(0).Caption = "ﬁ»Ê·"
'    frmMsg.fwBtn(1).Caption = "«‰’—«›"
'    frmMsg.Show vbModal
'    If mvarMsgIdx = 0 Then
'        Update = -1
'        Exit Function
'    End If
'
'End If
    On Error GoTo ErrHandler
    
    If clsArya.LimitedVersion = True And HardLockFlagTrial = False Then
        If Val(txtNo.Text) > 3000 Then
           MsgBox " ‰”ŒÂ ¬“„«Ì‘Ì - ‘„« „Ã«“ »Â «” ›«œÂ «“ «Ì‰ ”Ì” „ ‰Ì” Ìœ  "
           End
        End If
    End If
    
    FlxDetail_ValidateEdit FlxDetail.Row, FlxDetail.Col, False
    
    If mvarStatus = PurchaseReturn Or mvarStatus = Purchase Then
        If Me.txtCustomer.Tag < 1 Then
            frmMsg.fwlblMsg.Caption = "·ÿ›« Ìò  «„Ì‰ ﬂ‰‰œÂ «‰ Œ«» ‰„«ÌÌœ"
            frmMsg.fwBtn(0).Visible = False
            frmMsg.fwBtn(1).ButtonType = flwButtonOk
            frmMsg.fwBtn(1).Caption = "ﬁ»Ê·"
            frmMsg.Show vbModal
            Update = -1
            Exit Function
        End If
    End If
    If (mvarStatus = fromStore Or mvarStatus = toStore) And cmbDestInventory.ListIndex = -1 And mvarAnalyzeForm = False Then
        ShowMessage "¬Ì« »—«Ì À»  ”‰œ »œÊ‰ „ﬁ’œ »⁄œÌ «ÿ„Ì‰«‰ œ«—Ìœø", True, True, "»·Ì", "ŒÌ—"
        If mvarMsgIdx = vbNo Then
            Exit Function
            Update = -1
        End If
    End If
    If Not Me.CodeCount Then
        Update = -1
        Exit Function
    Else
     If mvarAnalyzeForm = False Then

        frmMsg.fwlblMsg.Caption = "¬Ì« „Ì ŒÊ«ÂÌœ ”‰œ Ã«—Ì À»  ‘Êœ"
        frmMsg.fwBtn(0).Visible = True
        frmMsg.fwBtn(1).Visible = True
        frmMsg.fwBtn(0).ButtonType = flwButtonOk
        frmMsg.fwBtn(0).Caption = "»·Ì"
        frmMsg.fwBtn(1).ButtonType = flwButtonCancel
        frmMsg.fwBtn(1).Caption = "ŒÌ—"
        frmMsg.Show vbModal
        If mvarMsgIdx = vbNo Then
            Exit Function
            Update = -1
        End If
     End If
    End If
    
    Dim mydata As Double
    Dim j As Integer
    Dim intLastFactorId As Double
    Dim boolValidServeplace As Boolean
        
    Update = Val(txtNo.Text)
    txtTime.Text = FormatDateTime(time, vbShortTime)
            
    
    BalancePayment = False
   
    If mvarStatus = PurchaseReturn Then BalancePayment = False
    
    If intTempFich <> 0 Then
        ReDim Parameter(2) As Parameter
        Parameter(0) = GenerateInputParameter("@No", adBigInt, 8, intTempFich)
        Parameter(1) = GenerateInputParameter("@Status", adInteger, 4, mvarStatus)
        Parameter(2) = GenerateInputParameter("@Branch", adInteger, 4, CurrentBranch)
        RunParametricStoredProcedure "Delete_Temp_Factor", Parameter
    End If
    BitSaveTempReceived = False
    If mvarStatus = TempRecieved And MyFormAddEditMode = ViewMode Then
        If ClsFormAccess.frmSaveTempReceived = False Then
            ShowDisMessage "œ” —”Ì »—«Ì À»  —”Ìœ „Êﬁ  »Â œ«∆„ ò«›Ì ‰Ì” ", 1500
            Update = -1
            Exit Function
        End If
        If BitTempReceived = True Then
            ShowDisMessage "—”Ìœ „Êﬁ  ﬁ»·« »Â —”Ìœ œ«∆„  »œÌ· ‘œÂ", 1500
            Update = -1
            Exit Function
        End If
        ShowMessage "¬Ì« —”Ìœ „Êﬁ   »œÌ· »Â —”Ìœ œ«∆„ „Ì ‘Êœ ø", True, True, "»·Ì", "ŒÌ—"
        If mvarMsgIdx = vbYes Then
            mvarStatus = Purchase
            MyFormAddEditMode = AddMode
            txtDescription = "—”Ìœ „Êﬁ  ‘„«—Â  " & txtNo.Text & " »œÌ· »Â —”Ìœ œ«∆„  "
            BitSaveTempReceived = True
        Else
            Update = -1
            Exit Function
        End If
    End If
    Dim temp As Integer
    Dim RepeatState, jjj, kk, tmpmvarStatus, TmpStatus, tmpDestInventory, tmpInventoryNo As Integer
    tmpmvarStatus = mvarStatus
    TmpStatus = mvarStatus
    kk = 1
    
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
                   st = GenerateDetailsString3(st, .TextMatrix(i, 1), .TextMatrix(i, 5), Val(.TextMatrix(i, 3)), Val(.TextMatrix(i, 10)), Val(.TextMatrix(i, 11)), " ", Val(.TextMatrix(i, 12)), Val(.TextMatrix(i, 13)), Val(.TextMatrix(i, 14)), .TextMatrix(i, 8), .TextMatrix(i, 9))
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
         Update = -1
         Exit Function
     End If
    
    Select Case MyFormAddEditMode
    
        Case ViewMode 'view mode
        
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
            If mvarStatus = Purchase And cmbDestInventory.ListIndex <> -1 Then
                RepeatState = 2
                tmpDestInventory = DestInventoryNo
                DestInventoryNo = 0
            Else
                RepeatState = 1
            End If

Repeat1:
            For jjj = kk To RepeatState
                If MojodiControlFlag = True And clsStation.RowMojodiControl = False And mvarStatus = fromStore Then
                     ReDim Parameter(3) As Parameter
                     Parameter(0) = GenerateInputParameter("@DetailsString", adVarWChar, 4000, DetailsString1)
                    Parameter(1) = GenerateInputParameter("@DetailsString2", adVarWChar, 4000, DetailsString2)
                    Parameter(2) = GenerateInputParameter("@DetailsString3", adVarWChar, 4000, DetailsString3)
                    Parameter(3) = GenerateInputParameter("@DetailsString4", adVarWChar, 4000, DetailsString4)
                     Set rctmp = RunParametricStoredProcedure2Rec("CheckPreSave", Parameter)
                     If Not (rctmp.BOF Or rctmp.EOF) Then
                         mvarAddeditMode = MyFormAddEditMode
                         frmMojodiReduce.Show vbModal
                         If frmMojodiReduce.Result = False Then
                            sFactorReceived = ""
                            Update = -1
                            Exit Function
                         End If
                     End If
                 End If
                
                ReDim Parameter(28) As Parameter
                
                Parameter(0) = GenerateInputParameter("@Status", adInteger, 4, tmpmvarStatus)
                If (Me.txtCustomer.Tag > -1) Then
                    Parameter(1) = GenerateInputParameter("@Owner", adInteger, 4, Me.txtCustomer.Tag)
                Else
                    Parameter(1) = GenerateInputParameter("@Owner", adInteger, 4, -1)
                End If
                Parameter(2) = GenerateInputParameter("@Customer", adInteger, 4, 0)
                Parameter(3) = GenerateInputParameter("@DiscountTotal", adDouble, 8, Val(Me.lblDiscountTotal))
                Parameter(4) = GenerateInputParameter("@CarryFeeTotal", adDouble, 8, Val(Me.lblCarryFeeTotal))
                Parameter(5) = GenerateInputParameter("@Recursive", adInteger, 4, Val(Me.txtRecursive.Text))
                Parameter(6) = GenerateInputParameter("@InCharge", adInteger, 4, 0)
                Parameter(7) = GenerateInputParameter("@FacPayment", adBoolean, 1, Abs(CInt(BalancePayment)))
                Parameter(8) = GenerateInputParameter("@OrderType", adInteger, 4, mVarOrderType)
                Parameter(9) = GenerateInputParameter("@StationId", adInteger, 4, clsArya.StationNo)
                Parameter(10) = GenerateInputParameter("@ServiceTotal", adDouble, 8, 0)
                Parameter(11) = GenerateInputParameter("@PackingTotal", adDouble, 8, Val(Me.lblPackingTotal))
                Parameter(12) = GenerateInputParameter("@TableNo", adInteger, 4, 0)
                Parameter(13) = GenerateInputParameter("@User", adInteger, 4, mvarCurUserNo)
                Parameter(14) = GenerateInputParameter("@Date", adVarWChar, 50, txtDate.Text)
                Parameter(15) = GenerateInputParameter("@DetailsString", adVarWChar, 4000, DetailsString1)
                Parameter(16) = GenerateInputParameter("@ds", adVarWChar, 4000, sFactorReceived)
                Parameter(17) = GenerateInputParameter("@Balance", adBoolean, 1, 0)
                Parameter(18) = GenerateInputParameter("@AccountYear", adSmallInt, 2, AccountYear)
                Parameter(19) = GenerateInputParameter("@NvcDescription", adVarWChar, 50, Right(txtDescription.Text, 50))
                Parameter(20) = GenerateInputParameter("@HavaleNo", adInteger, 4, 0)
                Parameter(21) = GenerateInputParameter("@TempAddress", adVarWChar, 255, "")
                Parameter(22) = GenerateInputParameter("@GuestNo", adSmallInt, 2, Null)
                Parameter(23) = GenerateInputParameter("@DetailsString2", adVarWChar, 4000, DetailsString2)
                Parameter(24) = GenerateInputParameter("@DetailsString3", adVarWChar, 4000, DetailsString3)
                Parameter(25) = GenerateInputParameter("@DetailsString4", adVarWChar, 4000, DetailsString4)
                Parameter(26) = GenerateInputParameter("@AddedTotal", adInteger, 4, IIf(chKTax.Value = 1, 0, Val(lblTaxTotal)))
                Parameter(27) = GenerateInputParameter("@Rasmi", adBoolean, 1, chKTax)
                Parameter(28) = GenerateOutputParameter("@LastFacMNo", adInteger, 4)
                        
                                       
                Update = RunParametricStoredProcedure("InsertFactorMasterDetails", Parameter)
                If Update = -1 Then GoTo ErrHandler
                If BitSaveTempReceived = True Then
                    ReDim Parameter(1) As Parameter
                    Parameter(0) = GenerateInputParameter("@intSerialNo", adBigInt, 8, intSerialNo)
                    Parameter(1) = GenerateInputParameter("@Branch", adInteger, 4, CurrentBranch)
                    RunParametricStoredProcedure "Update_BitTempReceived", Parameter
                End If
                Dim DestinationId As Long
                If intVersion = Diamond And mvarStatus = fromStore Then
                    If cmbDestination.ListIndex = -1 Then DestinationId = 0 Else DestinationId = cmbDestination.ItemData(cmbDestination.ListIndex)
                    If DestinationId > 0 Then
                        ReDim Parameter(1) As Parameter
                        Parameter(0) = GenerateInputParameter("@intSerialNo", adBigInt, 8, Update)
                        Parameter(1) = GenerateInputParameter("@DestinationId", adInteger, 4, DestinationId)
                        RunParametricStoredProcedure "Update_tFacM_Destination", Parameter
                    End If
                End If
                ReDim Parameter(0) As Parameter
                Parameter(0) = GenerateInputParameter("@intSerialNo", adBigInt, 8, Update)
                Set rctmp = RunParametricStoredProcedure2Rec("Get_RowCount_FactorDetail", Parameter)
                Update = rctmp!No
                'À»   ⁄œ«œ ÃœÌœ «”‰«œ
'                If clsArya.LimitedVersion = True And HardLockFlagTrial = False Then
'                     ' ‰Ê‘ ‰  ⁄œ«œ —ﬂÊ—œ ÃœÌœ
'                    RegRec = CountRecord + 1 + 10
'                    Call mdifrm.FWRegistry1.GetKeyStr(FLWSystem.flwRegLocalMachine, StrTemp5, "String Value3", strTemp)
'                    strTemp = mdifrm.FWEncryption1.Decode(strTemp, 1000)
'
'                    strTemp3 = mdifrm.FWEncryption1.Encode(CStr(RegRec), Val(strTemp) + 1000)
'                    If mdifrm.FWRegistry1.SetKeyStr(flwRegLocalMachine, StrTemp5, "String Value6", strTemp3) <> FLWSystem.flwSuccess Then
'                        Call MsgBox("Œÿ« œ— À»  «ÿ·«⁄«  - ﬂœ Œÿ« 15  " & vbLf, vbCritical)
'                      '  Unload Me
'                    End If
'
'                End If
                If (mvarStatus = Purchase Or mvarStatus = PurchaseReturn Or mvarStatus = fromStore) And clsArya.ExternalAccounting = True Then
                    If mvarStatus = fromStore And intVersion <> Diamond Then
                        ShowDisMessage "«„ò«‰  Ê·Ìœ « Ê„« Ìò ”‰œ Õ”«»œ«—Ì ÕÊ«·Â ›ﬁÿ œ— Ê—é‰ «·„«” ÊÃÊœ œ«—œ", 1500
                    Else
                        Status = mvarStatus
                        If mvarStatus = fromStore Then Tafsili = Tafsili_3
                        SanadNo = Accounting.Insert_CustomerSale(AddMode, 0, txtDate.Text, Tafsili, txtCustomer.Text, Update, CCur(lblSumPrice.Tag), Val(LblSubTotal), Val(lblDiscountTotal), Val(lblCarryFeeTotal), Val(lblPackingTotal), Val(lblTaxTotal), Val(LblDutyTotal), Status, "", "", Tafsili_2)
                    End If
                End If
                
                If jjj = 2 Then
                    Cancel
                    Exit Function
                End If
                If RepeatState = 2 Then
                    frmMsg.fwlblMsg.Caption = "¬Ì« „Ì ŒÊ«ÂÌœ «Ì‰ «ﬁ·«„ »Â „ﬁ’œ »⁄œÌ ÕÊ«·Â ‘Ê‰œ"
                    frmMsg.fwBtn(0).Visible = True
                    frmMsg.fwBtn(1).Visible = True
                    frmMsg.fwBtn(0).ButtonType = flwButtonOk
                    frmMsg.fwBtn(0).Caption = "»·Ì"
                    frmMsg.fwBtn(1).ButtonType = flwButtonCancel
                    frmMsg.fwBtn(1).Caption = "ŒÌ—"
                    frmMsg.Show vbModal
                    If mvarMsgIdx = vbYes Then
                        tmpmvarStatus = 6
                        DestInventoryNo = tmpDestInventory
                        Me.txtCustomer.Tag = -1
                    Else
                        Cancel
                        Exit Function
                    End If
                Else
                    Cancel
                    Exit Function
                End If

           Next jjj
        
        Case EditMode     'Edit Mode
        
            tmpmvarStatus = mvarStatus
            If mvarStatus = Purchase And cmbDestInventory.ListIndex <> -1 Then
                RepeatState = 2
                tmpDestInventory = DestInventoryNo
                DestInventoryNo = 0
            Else
                RepeatState = 1
            End If
            If MojodiControlFlag = True And clsStation.RowMojodiControl = False And mvarStatus = fromStore Then
                ReDim Parameter(5) As Parameter
                mvarNo = Val(txtNo.Text)
                Parameter(0) = GenerateInputParameter("@DetailsString", adVarWChar, 4000, DetailsString1)
                Parameter(1) = GenerateInputParameter("@Status", adInteger, 4, mvarStatus)
                Parameter(2) = GenerateInputParameter("@No", adBigInt, 8, mvarNo)
               Parameter(3) = GenerateInputParameter("@DetailsString2", adVarWChar, 4000, DetailsString2)
               Parameter(4) = GenerateInputParameter("@DetailsString3", adVarWChar, 4000, DetailsString3)
               Parameter(5) = GenerateInputParameter("@DetailsString4", adVarWChar, 4000, DetailsString4)
                Set rctmp = RunParametricStoredProcedure2Rec("CheckPreSave_Edit", Parameter)
                Dim ss As String
                ss = ""
                If Not (rctmp.BOF Or rctmp.EOF) Then
                    mvarAddeditMode = MyFormAddEditMode
                    frmMojodiReduce.Show vbModal
                    If frmMojodiReduce.Result = False Then
                       sFactorReceived = ""
                       Update = -1
                       Exit Function
                    End If
                End If
             End If
            ReDim Parameter(28) As Parameter
            Parameter(0) = GenerateInputParameter("@No", adInteger, 4, Val(txtNo.Text))
            Parameter(1) = GenerateInputParameter("@Status", adInteger, 4, mvarStatus)
            If Me.txtCustomer.Tag > -1 Then
                Parameter(2) = GenerateInputParameter("@Owner", adInteger, 4, Me.txtCustomer.Tag)
            Else
                Parameter(2) = GenerateInputParameter("@Owner", adInteger, 4, -1)
            End If
            Parameter(3) = GenerateInputParameter("@Customer", adInteger, 4, 0)
            Parameter(4) = GenerateInputParameter("@DiscountTotal", adInteger, 4, Val(Me.lblDiscountTotal))
            Parameter(5) = GenerateInputParameter("@CarryFeeTotal", adInteger, 4, Val(Me.lblCarryFeeTotal))
            Parameter(6) = GenerateInputParameter("@Recursive", adInteger, 4, Val(Me.txtRecursive.Text))
            Parameter(7) = GenerateInputParameter("@InCharge", adInteger, 4, 0)
            Parameter(8) = GenerateInputParameter("@FacPayment", adBoolean, 1, Abs(CInt(boolPayment)))
            Parameter(9) = GenerateInputParameter("@OrderType", adInteger, 4, mVarOrderType)
            Parameter(10) = GenerateInputParameter("@StationId", adInteger, 4, clsArya.StationNo)
            Parameter(11) = GenerateInputParameter("@ServiceTotal", adInteger, 4, 0)
            Parameter(12) = GenerateInputParameter("@PackingTotal", adInteger, 4, Val(Me.lblPackingTotal))
            Parameter(13) = GenerateInputParameter("@TableNo", adInteger, 4, 0)
            Parameter(14) = GenerateInputParameter("@User", adInteger, 4, mvarCurUserNo)
            Parameter(15) = GenerateInputParameter("@Date", adVarWChar, 50, txtDate.Text)

            Parameter(16) = GenerateInputParameter("@DetailsString", adVarWChar, 4000, DetailsString1)
            Parameter(17) = GenerateInputParameter("@ds", adVarWChar, 4000, sFactorReceived)
            Parameter(18) = GenerateInputParameter("@Balance", adBoolean, 1, 0)
            Parameter(19) = GenerateInputParameter("@AccountYear", adSmallInt, 2, AccountYear)
            Parameter(20) = GenerateInputParameter("@NvcDescription", adVarWChar, 50, Right(txtDescription.Text, 50))
            Parameter(21) = GenerateInputParameter("@TempAddress", adVarWChar, 255, "")
            Parameter(22) = GenerateInputParameter("@GuestNo", adSmallInt, 2, Null)
            Parameter(23) = GenerateInputParameter("@DetailsString2", adVarWChar, 4000, DetailsString2)
            Parameter(24) = GenerateInputParameter("@DetailsString3", adVarWChar, 4000, DetailsString3)
            Parameter(25) = GenerateInputParameter("@DetailsString4", adVarWChar, 4000, DetailsString4)
            Parameter(26) = GenerateInputParameter("@AddedTotal", adInteger, 4, IIf(chKTax.Value = 1, 0, Val(lblTaxTotal)))
            Parameter(27) = GenerateInputParameter("@Rasmi", adBoolean, 1, chKTax)
            Parameter(28) = GenerateOutputParameter("@LastFacMNo", adInteger, 4)
                    
            Update = RunParametricStoredProcedure("EditFactorMasterDetails", Parameter)
            If Update = -1 Then GoTo ErrHandler
            If intVersion = Diamond And mvarStatus = fromStore Then
                If cmbDestination.ListIndex = -1 Then DestinationId = 0 Else DestinationId = cmbDestination.ItemData(cmbDestination.ListIndex)
                    
                ReDim Parameter(1) As Parameter
                Parameter(0) = GenerateInputParameter("@intSerialNo", adBigInt, 8, intSerialNo)
                Parameter(1) = GenerateInputParameter("@DestinationId", adInteger, 4, DestinationId)
                RunParametricStoredProcedure "Update_tFacM_Destination", Parameter
            End If
            If (mvarStatus = Purchase Or mvarStatus = PurchaseReturn Or mvarStatus = fromStore) And clsArya.ExternalAccounting = True Then
                If mvarStatus = fromStore And intVersion <> Diamond Then
                    ShowDisMessage "«„ò«‰  Ê·Ìœ « Ê„« Ìò ”‰œ Õ”«»œ«—Ì ÕÊ«·Â ›ﬁÿ œ— Ê—é‰ «·„«” ÊÃÊœ œ«—œ", 1500
                Else
                    Status = mvarStatus
                    If mvarStatus = fromStore Then Tafsili = Tafsili_3
                    SanadNo = Accounting.Insert_CustomerSale(EditMode, Refrence_Acc, txtDate.Text, Tafsili, txtCustomer.Text, Update, CCur(lblSumPrice.Tag), Val(LblSubTotal), Val(lblDiscountTotal), Val(lblCarryFeeTotal), Val(lblPackingTotal), Val(lblTaxTotal), Val(LblDutyTotal), Status, "", "", Tafsili_2)
                End If
            End If
            If RepeatState = 2 Then
                frmMsg.fwlblMsg.Caption = "œﬁ  ‘Êœ ÕÊ«·Â Â«Ì ﬁ»·Ì œ” Ì Å«ò ‘Ê‰œ" & vbLf & "¬Ì« „Ì ŒÊ«ÂÌœ «Ì‰ «ﬁ·«„ »Â „ﬁ’œ »⁄œÌ ÕÊ«·Â ‘Ê‰œ"
                frmMsg.fwBtn(0).Visible = True
                frmMsg.fwBtn(1).Visible = True
                frmMsg.fwBtn(0).ButtonType = flwButtonOk
                frmMsg.fwBtn(0).Caption = "»·Ì"
                frmMsg.fwBtn(1).ButtonType = flwButtonCancel
                frmMsg.fwBtn(1).Caption = "ŒÌ—"
                frmMsg.Show vbModal
                If mvarMsgIdx = vbYes Then
                    tmpmvarStatus = 6
                    DestInventoryNo = tmpDestInventory
                    Me.txtCustomer.Tag = -1
                    kk = 2
                    GoTo Repeat1
                Else
                    Cancel
                    Exit Function
                End If
            Else
                Cancel
                Exit Function
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
                
            
    End Select
    sFactorReceived = ""
    Cancel
    
    If clsArya.LimitedVersion = True And HardLockFlagTrial = False And (RemaindateFlag = True Or maxRecordCountFlag = True) Then
        TrialCountFlag = TrialCountFlag + 1
        If TrialCountFlag Mod 2 = 0 Then
            ShowMessage " ‘„« œ— Õ«· «” ›«œÂ «“ ‰”ŒÂ ¬“„«Ì‘Ì ”Ì” „ Â«Ì ›—Ê‘ê«ÂÌ «› ÃÌ ¬—Ì« „Ì »«‘Ìœ" & " »—«Ì  ÂÌÂ ‰”ŒÂ «’·Ì ‰—„ «›“«—»« ‘—ﬂ  «› ÃÌ ¬—Ì« Ì« ‰„«Ì‰œê«‰ ›—Ê‘  „«” »êÌ—Ìœ ", True, False, " «∆Ìœ", ""
            Sleep 1000 * TrialCountFlag
        End If
    End If
    
    
    Exit Function
    
ErrHandler:
''''    Select Case Err.Number
''''        Case 0
''''
''''        Case -2147217873
''''
''''            frmDisMsg.lblMessage = "«Ì‰  —òÌ» «“ „ò«‰Â«Ì ”—Ê œ— ”Ì” „  ⁄—Ì› ‰‘œÂ «”  " & vbCrLf & "·ÿ›« ›«ò Ê— —« «’·«Õ ‰„ÊœÂ Ê ”Å” À»  ‰„«ÌÌœ"
''''            frmDisMsg.Timer1.Enabled = True
''''            frmDisMsg.Show
''''        Case Else
        
            MsgBox err.Description, vbOKOnly, err.Number
''''    End Select
    Update = -1
    
End Function
Public Sub UpdatelblServePlace()

    ReDim Parameter(2) As Parameter
    
    Parameter(0) = GenerateInputParameter("@CurrentServePlace", adInteger, 4, mvarServePlace)
    Parameter(1) = GenerateInputParameter("@intLangugae", adInteger, 4, clsStation.Language)
    Parameter(2) = GenerateOutputParameter("@Caption", adVarWChar, 50)
    
    lblServePlace.Caption = RunParametricStoredProcedure2String("GetServePlaceCaption", Parameter)
    
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

End Sub
Public Sub Find()
    
    frmFindFactor.Show vbModal
    
'    If MyFormAddEditMode <> ViewMode Then
'        Cancel
'    End If
    Dim tempNo As Long
    Dim ii As Long
    If mvarcode <> 0 Then
        tempNo = mvarcode
        txtNo.Text = mvarcode
        'mvarcode = 0
        If mvarStatus = Purchase And TempStatus = TempRecieved Then
            mvarStatus = TempRecieved
            For ii = 0 To CmbStatus.ListCount - 1
                If CmbStatus.ItemData(ii) = TempRecieved Then
                    CmbStatus.ListIndex = ii
                    ii = 0
                    Exit For
                End If
            Next ii
'            CmbStatus_Click
        End If
        tempNo = mvarcode
        txtNo.Text = mvarcode
        MyFormAddEditMode = ViewMode   'view Mode
        GetDataDetail
        RefreshLables
        SetFirstToolBar
        If mvarStatus = Purchase And TempStatus = TempRecieved Then
            'Number
            txtDescription = "—”Ìœ „Êﬁ  ‘„«—Â  " & tempNo
        End If
    Else
        Exit Sub
        
    End If
    
End Sub

Public Sub UndoRedo()
    Dim DatabaseBranch As Integer
    ReDim Parameters(0) As Parameter

    Parameters(0) = GenerateOutputParameter("@CurrentBranch", adInteger, 4)

    DatabaseBranch = RunParametricStoredProcedure2String("Get_CurrentBranch", Parameters)
    
    If CurrentBranch <> DatabaseBranch Then
        frmDisMsg.lblMessage.Caption = "›Ì‘ ‘⁄»Â œÌê— „—ÃÊ⁄ ‰„Ì ‘Êœ "
        frmDisMsg.Timer1.Interval = 2000
        frmDisMsg.Timer1.Enabled = True
        frmDisMsg.Show vbModal
        Exit Sub
    End If

Dim Rst As New ADODB.Recordset

If mvarStatus = toStore Then
    frmMsg.fwlblMsg.Caption = " —”Ìœ «“ ÿ—Ìﬁ ÕÊ«·Â „—»ÊÿÂ ﬁ«»· „—ÃÊ⁄ «”   "
    frmMsg.fwBtn(0).Visible = False
    frmMsg.fwBtn(1).ButtonType = flwButtonOk
    frmMsg.fwBtn(1).Caption = "ﬁ»Ê·"
    frmMsg.Show vbModal
    Exit Sub
End If
If mvarCurUserNo = 0 Then

    Set Rst = Nothing
    frmDisMsg.lblMessage = "‘„« «Ã«“Â „—ÃÊ⁄ ò—œ‰ ›Ì‘ —« ‰œ«—Ìœ"
    frmDisMsg.Timer1.Enabled = True
    frmDisMsg.Show vbModal
    Exit Sub
End If

If (ClsFormAccess.RefferPurchase = False) Then

    Set Rst = Nothing
    frmDisMsg.lblMessage = "‘„« «Ã«“Â „—ÃÊ⁄ ò—œ‰ ›Ì‘ —« ‰œ«—Ìœ"
    frmDisMsg.Timer1.Enabled = True
    frmDisMsg.Show vbModal
    Exit Sub
    
    
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
            frmMsg.fwlblMsg.Caption = "¬Ì« „ÿ„∆‰Ìœ „Ì ŒÊ«ÂÌœ ›Ì‘ —« „—ÃÊ⁄ ﬂ‰Ìœø "
            frmMsg.Show vbModal
            If modgl.mvarMsgIdx = vbYes Then
               MyFormAddEditMode = RefferedMode
               txtRecursive.Text = 1
              '  Edit
                Printing
            End If
            
        Case 1
        
            frmMsg.fwlblMsg.Caption = "¬Ì« „ÿ„∆‰Ìœ „Ì ŒÊ«ÂÌœ ›Ì‘ „—ÃÊ⁄ ‘œÂ —« »—ê—œ«‰Ìœø "
            frmMsg.Show vbModal
            If modgl.mvarMsgIdx = vbYes Then
                fwlblRecursive.Visible = False
                MyFormAddEditMode = RefferedMode
                txtRecursive.Text = 0
                Printing
                Cancel
            End If
    End Select
        
       
End Sub
Public Sub ExitForm()
''''        frmMsg.fwlblMsg.Caption = "¬Ì« »—«Ì Œ—ÊÃ «“ ›«ﬂ Ê—Œ—Ìœ«ÿ„Ì‰«‰ œ«—Ìœ"
''''      ' frmMsg.fwBtn(0).Visible = True
''''        frmMsg.Fwbtn(1).ButtonType = flwButtonOk
''''        frmMsg.Fwbtn(1).ButtonType = flwButtonCancel
''''        frmMsg.Fwbtn(1).Caption = "ﬁ»Ê·"
''''        frmMsg.Fwbtn(1).Caption = "Œ—ÊÃ"
''''        frmMsg.Show vbModal
''''        If mvarMsgIdx = vbYes Then
    
           Unload Me
''''        End If
End Sub

Private Sub fwBtnCustFind_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
        
    Me.FindCust
    
End Sub

Private Sub cmbPackingTotal_Click()

    If MyFormAddEditMode = AddMode Or MyFormAddEditMode = EditMode Then
        
        If Right(lblNum.Caption, 1) <> "%" Then
            txtPacking.Text = Val(lblNum.Caption)
        Else
            txtPackingPercent = Val(lblNum.Caption)
        End If
        lblPackingTotal = (Val(txtSumFeeTotal.Text) * Val(txtPackingPercent.Text) / 100) + Val(txtPacking.Text)
        lblSumPrice = CCur(txtSumFeeTotal.Text) + Val(lblCarryFeeTotal.Caption) + Val(lblPackingTotal.Caption) - Val(lblDiscountTotal.Caption)
        lblSumPrice.Tag = lblSumPrice.Caption
        lblSumPrice.Caption = Format(lblSumPrice.Caption, "#,## —Ì«·")
        BtnKeypad(10).Enabled = True
        BtnKalaDelete.Enabled = True
        BtnKeypad(11).Enabled = True
    Else
    End If
    lblNum.Caption = ""

End Sub




Private Sub lblServePlace_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    Dim Rst As New ADODB.Recordset
    ReDim Parameter(0) As Parameter
    
    Parameter(0) = GenerateInputParameter("@CurrentServePlace", adInteger, 4, mvarServePlace)
    
    Set Rst = RunParametricStoredProcedure2Rec("GetValidServePlace", Parameter)
    
    If Not (Rst.EOF = True And Rst.BOF = True) Then
        mvarServePlace = Rst.Fields("intServeplace").Value
    Else
        Set Rst = RunStoredProcedure2RecordSet("GetFirstValidServePlace")
        mvarServePlace = Rst.Fields("intServeplace").Value
    End If
    RefreshLables
    
End Sub

Private Sub lblSumPrice_Click()

    lblSumPrice.RightToLeft = IIf(lblSumPrice.RightToLeft, False, True)

End Sub
Private Sub LblSubTotal_Click()
    LblSubTotal.RightToLeft = IIf(LblSubTotal.RightToLeft, False, True)
End Sub
Private Sub txtDescription_Change()
    BtnKeypad(11).Enabled = True     '"%"
    BtnKeypad(10).Enabled = True      '"."
    BtnKalaDelete.Enabled = True
    lblNum.Caption = ""
    lblBarCode.Caption = ""
End Sub

Private Sub txtDescription_GotFocus()
    textDescription = True
End Sub
Private Sub txtDescription_LostFocus()
    textDescription = False
End Sub



Private Sub lstGoodKey_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
        Me.ValueLstGood lstGoodKey
    End If
    
End Sub

Private Sub lstGoodKey_LostFocus()

    lstGoodKey.Visible = False
    
End Sub

Private Sub lstGoodKey_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    Me.ValueLstGood lstGoodKey
    
End Sub


Private Sub ResizeKit1_ExitResize(ByVal XScale As Double, ByVal YScale As Double)

    If formloadFlag = True Then
        SaveSetting strMainKey, Me.Name, "Height", Me.Height
        SaveSetting strMainKey, Me.Name, "Width", Me.Width
    End If

End Sub

Private Sub StatusBar_PanelClick(ByVal Panel As MSComctlLib.Panel)
    
    PanelClick Panel.index
    
End Sub

Private Sub TimerNumber_Timer()

    If MyFormAddEditMode = AddMode Then
       Me.Number
    End If
    
End Sub

Private Sub lblCarryFeeTotal_Click()

    lblCarryFeeTotal.RightToLeft = IIf(lblCarryFeeTotal.RightToLeft, False, True)

End Sub

Private Sub lblDiscountTotal_Click()

    lblDiscountTotal.RightToLeft = IIf(lblDiscountTotal.RightToLeft, False, True)

End Sub


Private Sub txtBarcode_GotFocus()
''''    txtBarcode.Text = ""
''''    mvarbarcode = True
End Sub

Private Sub txtBarcode_KeyDown(KeyCode As Integer, Shift As Integer)
    mvarbarcode = True
    Select Case Shift
        Case 0
            Select Case KeyCode
                Case 13
                    lblBarCode.Caption = txtBarcode.Text
                    barcode
                    FlxDetail.SetFocus
            End Select
    End Select

End Sub

Private Sub txtBarcode_LostFocus()
    mvarbarcode = False
    txtBarcode.Text = ""
End Sub

Private Sub txtNo_Change()

    FWLed1.Value = Val(txtNo.Text)
''''    If Trim(txtNo.Text) <> "" Then
''''            If Val(txtNo.Text) < 1000 Then
''''                FWLed1.Value = Val(txtNo.Text)
''''                FWLed1.Left = 7800 '6890
''''                FWLed1.Width = 1200
''''            ElseIf Val(txtNo.Text) < 10000 Then
''''                FWLed1.Value = Val(txtNo.Text)
''''                FWLed1.Left = 7600 '6690
''''                FWLed1.Width = 1400
''''            ElseIf Val(txtNo.Text) < 100000 Then
''''                FWLed1.Value = Val(txtNo.Text)
''''                FWLed1.Left = 7190 '6390
''''                FWLed1.Width = 1750
''''            Else
''''                FWLed1.Value = Val(txtNo.Text)
''''                FWLed1.Left = 6850 '6050
''''                FWLed1.Width = 2130
''''            End If
''''    End If
    
End Sub

Private Sub txtDate_KeyPress(KeyAscii As Integer)
'On Error Resume Next
'Bypass By Nemat 830513
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
lblNum.Caption = ""

End Sub

Private Sub txtDate_KeyDown(KeyCode As Integer, Shift As Integer)
    If txtDate.Locked = True Then Exit Sub
    If KeyCode = vbKeyDelete Then
        If Mid(txtDate.Text, txtDate.SelStart + 1, 1) = "/" Then
            KeyCode = 0
        End If
    End If

'    If KeyCode = vbKeyDelete Then
'        If Mid(txtDate.Text, txtDate.SelStart + 1, 1) = "/" Then
'            KeyCode = 0
'        End If
'    End If
    lblNum.Caption = ""
End Sub


Private Sub lblPackingTotal_Click()

    lblPackingTotal.RightToLeft = IIf(lblPackingTotal.RightToLeft, False, True)

End Sub


Private Sub txtSumCountNo_Click()
    
    txtSumCountNo.RightToLeft = IIf(txtSumCountNo.RightToLeft, False, True)

End Sub

Private Sub txtSumCountWeight_Click()
    
    txtSumCountWeight.RightToLeft = IIf(txtSumCountWeight.RightToLeft, False, True)

End Sub

Private Sub txtSumWeightTotal_Click()
    
    txtSumWeightTotal.RightToLeft = IIf(txtSumWeightTotal.RightToLeft, False, True)

End Sub

Private Sub RefreshLables()    'For Refresh Lables When Edit

    Dim ValueCountNo, ValueCountWeight, ValueSumWeight, ValueWeightTotal, ValueFeeTotal, ValueGoodDiscount, ValueGoodsDuty, ValueGoodsTax As Double
    Dim a As TextBox
    
    On Error Resume Next
    Dim ii As Integer
    txtSumFeeTotal.Text = 0
    txtSumCountNo.Caption = 0
    txtSumCountWeight.Caption = 0
    txtSumWeightTotal.Caption = 0
    ValueGoodDiscount = 0
    Me.txtSumFeeTotal.Text = 0
    ValueGoodsDuty = 0#
    ValueGoodsTax = 0#
    
    For ii = 1 To MaxRowFlexGrid - 1
        FlxDetail.TextMatrix(ii, 4) = Format(Val(FlxDetail.TextMatrix(ii, 1)) * CCur(FlxDetail.TextMatrix(ii, 3)), "##")
        ValueGoodDiscount = ValueGoodDiscount + Val(FlxDetail.TextMatrix(ii, 4)) * Val(FlxDetail.TextMatrix(ii, 10)) / 100
        If ValueGoodDiscount <> 0 Then ValueGoodDiscount = Format(ValueGoodDiscount, "##")
        FlxDetail.TextMatrix(ii, 4) = CCur(FlxDetail.TextMatrix(ii, 4)) - Format(CCur(FlxDetail.TextMatrix(ii, 4)) * Val(FlxDetail.TextMatrix(ii, 10)) / 100, "##")
        Me.txtSumFeeTotal.Text = Me.txtSumFeeTotal.Text + Val(FlxDetail.TextMatrix(ii, 1)) * CCur(FlxDetail.TextMatrix(ii, 3))
        If chKTax.Value = True And (mvarStatus = Purchase Or mvarStatus = PurchaseReturn) Then
            If FlxDetail.TextMatrix(ii, 19) = True Then ValueGoodsDuty = ValueGoodsDuty + FlxDetail.TextMatrix(ii, 4)
            If FlxDetail.TextMatrix(ii, 20) = True Then ValueGoodsTax = ValueGoodsTax + FlxDetail.TextMatrix(ii, 4)
        End If
        If Val(FlxDetail.TextMatrix(ii, 7)) <> 1 Then        'Numeric
        
            txtSumCountNo.Caption = txtSumCountNo.Caption + Val(FlxDetail.TextMatrix(ii, 1))
            txtSumWeightTotal.Caption = txtSumWeightTotal.Caption + Val(FlxDetail.TextMatrix(ii, 6)) * Val(FlxDetail.TextMatrix(ii, 1))
        Else
            txtSumCountWeight.Caption = txtSumCountWeight.Caption + 1
            txtSumWeightTotal.Caption = txtSumWeightTotal.Caption + Val(FlxDetail.TextMatrix(ii, 1))
        End If
    Next
LblSubTotal.Caption = CCur(Me.txtSumFeeTotal.Text)
'ValueDuty = CLng(ValueDuty)
'ValueTax = CLng(ValueTax)

Select Case MyFormAddEditMode

    Case Is <> ViewMode
        
        lblDiscountTotal = CLng(Val(txtDiscount.Text) + (CCur(txtSumFeeTotal.Text) * Val(txtDiscountPercent.Text) / 100)) + Val(ValueGoodDiscount)
        lblDiscountTotal = Format(lblDiscountTotal, "##")
        lblCarryFeeTotal = CLng(Val(txtCarryFee.Text) + (CCur(txtSumFeeTotal.Text) * Val(txtCarryFeePercent.Text) / 100))  ' + Val(txtCarryFeeCust.Text)
        lblPackingTotal = CLng(Val(txtPacking.Text) + (CCur(txtSumFeeTotal.Text) * Val(txtPackingPercent.Text) / 100))
        
        If (mvarStatus = Purchase Or mvarStatus = PurchaseReturn) And chKTax = True Then
            ReDim Parameter(5) As Parameter
            Parameter(0) = GenerateInputParameter("@ValueGoodsDuty", adDouble, 8, ValueGoodsDuty)
            Parameter(1) = GenerateInputParameter("@ValueGoodsTax", adDouble, 8, ValueGoodsTax)
            Parameter(2) = GenerateInputParameter("@DiscountTotal", adDouble, 8, Val(lblDiscountTotal.Caption))
            Parameter(3) = GenerateInputParameter("@ServiceTotal", adDouble, 8, 0)
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
        lblSumPrice.Caption = CCur(Val(txtSumFeeTotal.Text) + Val(lblCarryFeeTotal.Caption) + Val(lblPackingTotal.Caption) + Val(lblTaxTotal.Caption) + Val(LblDutyTotal.Caption) - Val(lblDiscountTotal.Caption))
        lblSumPrice.Tag = lblSumPrice.Caption
        lblSumPrice.Caption = Format(lblSumPrice.Caption, "#,## —Ì«·")
    
        If TempStatus <> mvarStatus Then
            ReDim Parameter(3) As Parameter
            Parameter(0) = GenerateInputParameter("@No", adBigInt, 8, Val(txtNo.Text))
            Parameter(1) = GenerateInputParameter("@Status", adInteger, 4, TempStatus)
            Parameter(2) = GenerateInputParameter("@AccountYear", adSmallInt, 2, AccountYear)
            Parameter(3) = GenerateInputParameter("@Branch", adInteger, 4, cmbBranch.ItemData(cmbBranch.ListIndex))
            
            Set rctmp = RunParametricStoredProcedure2Rec("Get_tFacM_By_No_Status", Parameter)
            
            If Not (rctmp.BOF Or rctmp.EOF) Then
                
                If Not IsNull(rctmp!Owner) Then
                    txtCustomer.Tag = rctmp!Owner
                    UpdatetxtCustomer           ''''
                End If
                If Not IsNull(rctmp) And rctmp.State = adStateOpen Then
                    rctmp.Close
                    Set rctmp = Nothing
                End If
            End If
        End If
    
    Case Else
               
        ReDim Parameter(3) As Parameter
        Parameter(0) = GenerateInputParameter("@No", adBigInt, 8, Val(txtNo.Text))
        Parameter(1) = GenerateInputParameter("@Status", adInteger, 4, mvarStatus)
        Parameter(2) = GenerateInputParameter("@AccountYear", adSmallInt, 2, AccountYear)
        Parameter(3) = GenerateInputParameter("@Branch", adInteger, 4, cmbBranch.ItemData(cmbBranch.ListIndex))
        
        Set rctmp = RunParametricStoredProcedure2Rec("Get_tFacM_By_No_Status", Parameter)
        
        If Not (rctmp.BOF Or rctmp.EOF) Then
               
            If Not IsNull(rctmp!DiscountTotal) Then
               lblDiscountTotal.Caption = rctmp!DiscountTotal
               txtDiscount.Text = rctmp!DiscountTotal - TmpGoodDiscount
            End If
            
            If Not IsNull(rctmp!CarryFeeTotal) Then
                lblCarryFeeTotal.Caption = rctmp!CarryFeeTotal
                txtCarryFee.Text = rctmp!CarryFeeTotal
            End If
            
            If Not IsNull(rctmp!sumPrice) Then
                lblSumPrice.Caption = rctmp!sumPrice
                lblSumPrice.Tag = lblSumPrice.Caption
                lblSumPrice.Caption = Format(lblSumPrice.Caption, "#,## —Ì«·")
            End If
            
''''            If Not IsNull(rctmp!ServiceTotal) Then
''''               lblServiceTotal.Caption = rctmp!ServiceTotal
''''               Txtservice.Text = rctmp!ServiceTotal
''''            End If
            
            If Not IsNull(rctmp!PackingTotal) Then
               lblPackingTotal.Caption = rctmp!PackingTotal
               txtPacking.Text = rctmp!PackingTotal
            End If
            
            If Not IsNull(rctmp!DutyTotal) Then
               LblDutyTotal.Caption = rctmp!DutyTotal
            End If
            If Not IsNull(rctmp!TaxTotal) Then
               lblTaxTotal.Caption = rctmp!TaxTotal
            End If
            
            If Not IsNull(rctmp!Date) Then
               Me.txtDate.Text = rctmp!Date
               Me.txtDate.Tag = rctmp!Date
            End If
            
            If Not IsNull(rctmp!RegDate) Then
              Me.txtRegDate.Text = rctmp!RegDate
            End If
            
            If Not IsNull(rctmp!Owner) Then
                txtCustomer.Tag = rctmp!Owner
            End If
            
            If Not IsNull(rctmp!time) Then
               Me.txtTime.Text = rctmp!time
            End If
            
            If Not IsNull(rctmp!Recursive) Then
                Me.txtRecursive = rctmp!Recursive
            End If
                         
        End If
        
        If Me.txtRecursive = 1 Then
            fwlblRecursive.Visible = True
            If (mvarCurUserNo = dblFichUser And ClsFormAccess.RefferInvoice = True) Or (ClsFormAccess.RefferedAllStationsFactors = True) Then
                MyFormAddEditMode = RefferedMode
            End If
        Else
            fwlblRecursive.Visible = False
        End If
               
        rctmp.Close
  
  End Select
  
  UpdatetxtCustomer
  UpdatelblServePlace
  
End Sub
Public Function CodeCount() As Boolean

    If MaxRowFlexGrid <= 1 Then    ' Or (lblSumPrice < 1 And lblDiscountTotal.Caption = 0)
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
    If Not (Rst.EOF = True And Rst.BOF = True) Then
        txtNo.Text = Rst!No
        Rst.Close: Set Rst = Nothing
    End If
End Sub

Public Sub ArrowkeyStatusbar(intDirection As EnumDirection, Optional CurrentintSerialNo As Double)                'Display 5 Last Fich

Dim j As Integer
Dim str1 As String
    
    ReDim Parameter(6) As Parameter
    
    Parameter(0) = GenerateInputParameter("@No", adBigInt, 8, CurrentintSerialNo)
    Parameter(1) = GenerateInputParameter("@Direction", adInteger, 4, intDirection)
    Parameter(2) = GenerateInputParameter("@User", adInteger, 4, mvarCurUserNo)
    Parameter(3) = GenerateInputParameter("@Status", adInteger, 4, mvarStatus)
    Parameter(4) = GenerateInputParameter("@Date", adWChar, 10, txtDate.Text)
    Parameter(5) = GenerateInputParameter("@AccountYear", adSmallInt, 2, AccountYear)
    Parameter(6) = GenerateInputParameter("@Branch", adSmallInt, 2, cmbBranch.ItemData(cmbBranch.ListIndex))
    
    Set rctmp = RunParametricStoredProcedure2Rec("NavigateFacm", Parameter)
    
    For i = 1 To 7
        Me.StatusBar.Panels(i).Tag = ""
        Me.StatusBar.Panels(i).Text = ""
    Next i
    
    If Not (rctmp.EOF = True And rctmp.BOF = True) Then
        i = 7
        Do While Not (rctmp.EOF = True)
            str1 = rctmp.Fields("SumPrice")
            i = i - 1
            If i = 1 Then
                Exit Do
            End If
            If i <> 1 And i <> 7 Then
                Me.StatusBar.Panels(i).Tag = rctmp.Fields("No").Value
                If rctmp!BascoleNo = 0 Then
                    Me.StatusBar.Panels(i).Text = IIf(Right(Str(rctmp.Fields("No")), 3) <> "000", "", "1") & Right(Str(rctmp.Fields("No")), 3) & ")" & str1
                Else
                    Me.StatusBar.Panels(i).Text = Right(Str(rctmp.Fields("BascoleNo")), 1) & ")" & IIf(Right(Str(rctmp.Fields("No")), 3) <> "000", "", "1") & Right(Str(rctmp.Fields("No")), 3) & ")" & str1
                End If
            End If
            rctmp.MoveNext
        Loop
    End If
    rctmp.Close

End Sub


Public Sub ValueLabel()

    Select Case Val(txtRecursive.Text)
        Case 1
        
            fwlblRecursive.Visible = True
            
        Case Else
        
            fwlblRecursive.Visible = False
            
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
                GetDataDetail
                RefreshLables
                For i = 2 To 6
                    Me.StatusBar.Panels(i).Bevel = sbrInset
                    Me.StatusBar.Panels(i).Enabled = True
                Next i
                Edit
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

    frmFindSupplier.Show vbModal
    
    If mvarcode <> 0 Then
        txtCustomer.Tag = mvarcode
        mvarcode = 0
    Else
        txtCustomer.Tag = -1
    End If
    
    RefreshLables
End Sub
Public Function GetGoodBarcode(Code As String)
    
    Dim ExistfromStore, ExisttoStore, ReturnValue As Boolean
    ReturnValue = False
    ExistfromStore = False
    ExisttoStore = False
    ReturnValue = False
    If Code = "" Then Exit Function
    
    ReDim Parameter(1) As Parameter
    Parameter(0) = GenerateInputParameter("@Barcode", adVarWChar, 50, Code)
    Parameter(1) = GenerateInputParameter("@GoodCode", adInteger, 4, 0)
    Set rctmp = RunParametricStoredProcedure2Rec("Get_Good_Barcode_Check", Parameter)
        
    If (rctmp.BOF Or rctmp.EOF) Then
        frmDisMsg.lblMessage.Caption = " «Ì‰ »«—ﬂœ œ— ”Ì” „  ⁄—Ì› ‰‘œÂ «”  "
        frmDisMsg.Timer1.Interval = 1000
        frmDisMsg.Timer1.Enabled = True
        frmDisMsg.Show vbModal
        rctmp.Close
        Exit Function
    End If
    
    If mvarStatus = fromStore And cmbDestInventory.ListIndex = -1 Then
        frmMsg.fwlblMsg.Caption = "«‰»«— „›’œ »«Ìœ «‰ Œ«» ‘Êœ"
        frmMsg.fwBtn(0).Visible = False
        frmMsg.fwBtn(1).ButtonType = flwButtonOk
        frmMsg.fwBtn(1).Caption = "ﬁ»Ê·"
        frmMsg.Show vbModal
        Exit Function
    End If
    If cmbInventory.ListIndex <> -1 And cmbDestInventory.ListIndex <> -1 Then
        If cmbInventory.ItemData(cmbInventory.ListIndex) = cmbDestInventory.ItemData(cmbDestInventory.ListIndex) Then
            frmMsg.fwlblMsg.Caption = "«‰»«—Â« »«Ìœ »« Â„ „ ›«Ê  »«‘‰œ"
            frmMsg.fwBtn(0).Visible = False
            frmMsg.fwBtn(1).ButtonType = flwButtonOk
            frmMsg.fwBtn(1).Caption = "ﬁ»Ê·"
            frmMsg.Show vbModal
            Exit Function
        End If
    End If
    
    
    ReDim Parameter(3) As Parameter
    Parameter(0) = GenerateInputParameter("@Barcode", adVarWChar, 50, Code)
    Parameter(1) = GenerateInputParameter("@StationId", adInteger, 4, clsArya.StationNo)
    Parameter(2) = GenerateInputParameter("@Flag", adBoolean, 1, 0)
    Parameter(3) = GenerateInputParameter("@AccountYear", adSmallInt, 2, AccountYear)
    Set rctmp = RunParametricStoredProcedure2Rec("Get_Good_Barcode", Parameter)
    
    If Not (rctmp.BOF Or rctmp.EOF) Then
        If cmbDestInventory.ListIndex <> -1 Then
            DestInventoryNo = cmbDestInventory.ItemData(cmbDestInventory.ListIndex)
        Else
            ExisttoStore = True
        End If
        Do While rctmp.EOF <> True
        
    ''''           InventoryNo = rctmp.Fields("InventoryNo").Value
            mvarInventoryNo = cmbInventory.ItemData(cmbInventory.ListIndex) ' rctmp.Fields ("InventoryNo")
            If mvarInventoryNo = rctmp.Fields("InventoryNo").Value Then
                  
                mvarGoodCode = rctmp.Fields("Code")
                mvarUnitGood = rctmp.Fields("Unit")
                mvarGoodName = rctmp.Fields("Name")
                mvarGoodWeight = rctmp.Fields("Weight")
                mvarNumberOfUnit = rctmp.Fields("NumberOfUnit")
                '   mvarDisCount = rctmp.Fields("Discount")   Only For Sale
                mvarMojodi = rctmp.Fields("Mojodi")
                mvarSellPrice = rctmp.Fields("SellPrice")
               If chKTax = True Then
                  mvarDuty = True
                  mvarTax = True
               Else
                  mvarDuty = rctmp.Fields("DutySale")
                  mvarTax = rctmp.Fields("TaxSale")
               End If
                If mvarStatus = 6 Then
                    If clsStation.FromStoreFee = 0 Then
                        ReDim Parameter(3) As Parameter
                        Parameter(0) = GenerateInputParameter("@GoodCode", adInteger, 4, mvarGoodCode)
                        Parameter(1) = GenerateInputParameter("@DateAfter", adWChar, 20, Mid(clsDate.shamsi(Date), 3, 2) & "/01" & "/01")
                        Parameter(2) = GenerateInputParameter("@DateBefore", adWChar, 20, txtDate.Text)
                        Parameter(3) = GenerateInputParameter("@Flag", adBoolean, 1, 0)
                        Set RstTemp = RunParametricStoredProcedure2Rec("AverageCalculateBuyPrice", Parameter)
    
                        mvarBuyPrice = RstTemp!AverageBuyPrice
                        Set RstTemp = Nothing
                    ElseIf clsStation.FromStoreFee = 1 Then
                        mvarBuyPrice = rctmp.Fields("BuyPrice").Value
                    ElseIf clsStation.FromStoreFee = 2 Then
                        mvarBuyPrice = rctmp.Fields("SellPrice").Value
                    ElseIf clsStation.FromStoreFee = 3 Then
                        mvarBuyPrice = rctmp.Fields("SellPrice2").Value
                    ElseIf clsStation.FromStoreFee = 4 Then
                        mvarBuyPrice = rctmp.Fields("SellPrice3").Value
                    ElseIf clsStation.FromStoreFee = 5 Then
                        mvarBuyPrice = rctmp.Fields("SellPrice4").Value
                    ElseIf clsStation.FromStoreFee = 6 Then
                        mvarBuyPrice = rctmp.Fields("SellPrice5").Value
                    ElseIf clsStation.FromStoreFee = 7 Then
                        mvarBuyPrice = rctmp.Fields("SellPrice6").Value
                    End If
                Else
                    mvarBuyPrice = rctmp.Fields("BuyPrice").Value
                End If
                ExistfromStore = True
            End If
            If DestInventoryNo <> 0 Then
                If DestInventoryNo = rctmp.Fields("InventoryNo").Value Then
                    ExisttoStore = True
                End If
            End If

            rctmp.MoveNext
      
        Loop
    End If
    If ExistfromStore = False Then
        frmDisMsg.lblMessage.Caption = " «Ì‰ ò«·« »—«Ì «‰»«— „»œ«  ⁄—Ì› ‰‘œÂ «”   "
        frmDisMsg.Timer1.Interval = 1000
        frmDisMsg.Timer1.Enabled = True
        frmDisMsg.Show vbModal
        ReturnValue = False
    End If
    If ExisttoStore = False Then
        frmDisMsg.lblMessage.Caption = " «Ì‰ ò«·« »—«Ì «‰»«— „ﬁ’œ  ⁄—Ì› ‰‘œÂ «”  "
        frmDisMsg.Timer1.Interval = 1000
        frmDisMsg.Timer1.Enabled = True
        frmDisMsg.Show vbModal
        ReturnValue = False
    End If
    rctmp.Close
    If (ExistfromStore = True And ExisttoStore = True) Then ReturnValue = True
    GetGoodBarcode = ReturnValue
End Function

Public Sub barcode()
    
    If GetGoodBarcode(lblBarCode.Caption) = True Then
        ChangeGoodquantity
    End If
        
beforeexit:

    lblBarCode = ""
    mvarbarcode = False
    
    
End Sub

Public Sub ValueLstGood(ByRef lstVir As ListBox)

    If GetGoodCode(lstVir.ItemData(lstVir.ListIndex)) = True Then
       ChangeGoodquantity
    End If
    lstVir.Visible = False
      
End Sub

Public Sub KeyPress(KeyAscii As Integer)
    Dim var1, var2 As Double
    Dim j As Double
    
    
    If MyFormAddEditMode = ViewMode Then
        Exit Sub
    End If
                
    
    ReDim Parameter(1) As Parameter
    Parameter(0) = GenerateInputParameter("@BtnAscDefault", adInteger, 4, KeyAscii)
    Parameter(1) = GenerateInputParameter("@notSupportedType", adInteger, 4, EnumGoodType.forSale)
    Set rctmp = RunParametricStoredProcedure2Rec("Get_Good_DefaultKb_Count", Parameter)
    
    i = rctmp.Fields("count")
    rctmp.Close
            
    If i > 1 Then
       
           Call frmFindGoods_Kb.SendVariables(False, mvarKeyCode, MvarShiftKey, KeyAscii)
           frmFindGoods_Kb.Show vbModal

    ElseIf i = 1 Then
    
        ReDim Parameter(2) As Parameter
        Parameter(0) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
        Parameter(1) = GenerateInputParameter("@BtnAscDefault", adInteger, 4, KeyAscii)
        Parameter(2) = GenerateInputParameter("@NotSupportedGoodType", adInteger, 4, EnumGoodType.forSale)
        Set rctmp = RunParametricStoredProcedure2Rec("Get_Good_DefaultKB", Parameter)
        If GetGoodCode(Val(rctmp.Fields("Code"))) = True Then
            ChangeGoodquantity
        End If
    End If


End Sub

Sub DoPrintLogo(PassedPrinterName As String, LogoFileName As String)

End Sub

Private Sub HideLstBoxes(KeyAscii As Integer)

    If (KeyAscii = 27) Then
        Me.lstGoodKey.Visible = False
    End If
    
End Sub

Private Function CalculateSumOfServeplace() As Integer
    
    Dim j As Integer
    Dim intServeplaces() As Integer
    
    ReDim Preserve intServeplaces(0)
    
    intServeplaces(0) = Val(FlxDetail.TextMatrix(1, 8))
    For i = 1 To MaxRowFlexGrid - 1
        ReDim Preserve intServeplaces(i)
        intServeplaces(i) = Val(FlxDetail.TextMatrix(i + 1, 8))
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
    Dim DatabaseBranch As Integer
''''    ReDim Parameters(0) As Parameter
''''
''''    Parameters(0) = GenerateOutputParameter("@CurrentBranch", adInteger, 4)
''''
''''    DatabaseBranch = RunParametricStoredProcedure2String("Get_CurrentBranch", Parameters)
''''
''''    If CurrentBranch <> DatabaseBranch Then
''''        frmDisMsg.lblMessage.Caption = "«„ò«‰ ’œÊ— ›Ì‘ »—«Ì ‘⁄»Â œÌê— ÊÃÊœ ‰œ«—œ "
''''        frmDisMsg.Timer1.Interval = 2000
''''        frmDisMsg.Timer1.Enabled = True
''''        frmDisMsg.Show vbModal
''''        Exit Function
''''    End If
'    If mvarStatus = toStore Then
'        frmMsg.fwlblMsg.Caption = " »—«Ì —”Ìœ »Â «‰»«— «“ —”Ìœ „Êﬁ  «” ›«œÂ ò‰Ìœ "
'        frmMsg.fwBtn(0).Visible = False
'        frmMsg.fwBtn(1).ButtonType = flwButtonOk
'        frmMsg.fwBtn(1).Caption = "ﬁ»Ê·"
'        frmMsg.Show vbModal
'        Exit Function
'    End If
'    If mvarStatus = TempRecieved Then
'        frmMsg.fwlblMsg.Caption = " —”Ìœ „Êﬁ  ﬁ«»· «’·«Õ ‰Ì”   "
'        frmMsg.fwBtn(0).Visible = False
'        frmMsg.fwBtn(1).ButtonType = flwButtonOk
'        frmMsg.fwBtn(1).Caption = "ﬁ»Ê·"
'        frmMsg.Show vbModal
'        Exit Function
'    End If
    
    Dim Answer As Boolean
    Dim CanAdd As Boolean
    Dim AmountVar As Double

    If lblNum.Caption = "-" Then
        lblNum.Caption = "-1"
    End If
    
    If txtScale.Text = "-" Then
        txtScale.Text = "-1"
    End If
    
    
    Select Case mvarUnitGood
        'Weight Good
        Case 1
            If clsStation.DirectBascule And clsStation.BasculeOn Then
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
            
        Case Else     'No Weight Good
        
            If lblNum.Caption <> "" Then
                AmountVar = Round(Val(lblNum.Caption), 0)
            Else
                AmountVar = 1
            End If
            If clsStation.NumberOfUnitBuy = True Then
                If MyFormAddEditMode = AddMode Then
                    AmountVar = AmountVar * mvarNumberOfUnit
                End If
            Else
                mvarNumberOfUnit = 1
            End If
    End Select
    
    Dim Row_Find As Integer
            
    If clsStation.RepetitiveGood = False And FindRecord_FlexGrid(mvarGoodCode) = True Then  'Exist Good In Fich
    
        If Left(lblNum.Caption, 1) = "-" And mvarUnitGood = 1 And clsStation.DeletedGood = True Then   'Weight Good & Delete
            FlxDetail.TextMatrix(FlxDetail.Row, 1) = 0
        Else
            If clsStation.NumberOfUnitBuy = True Then
                mvarNumberOfUnit = Val(FlxDetail.TextMatrix(FlxDetail.Row, 15))
            Else
                mvarNumberOfUnit = 1
            End If
            If mvarNumberOfUnit = 1 Or clsStation.NumberOfUnitBuy = False Then
                FlxDetail.TextMatrix(FlxDetail.Row, 1) = AmountVar + Val(FlxDetail.TextMatrix(FlxDetail.Row, 1))
            Else
                FlxDetail.TextMatrix(FlxDetail.Row, 1) = (AmountVar * Val(FlxDetail.TextMatrix(FlxDetail.Row, 15))) + Val(FlxDetail.TextMatrix(FlxDetail.Row, 1))
            End If
        End If
        
        If FlxDetail.TextMatrix(FlxDetail.Row, 1) < 0 Then
             frmMsg.fwlblMsg.Caption = " . „ﬁœ«— ﬂ«·« ‰„Ì  Ê«‰œ „‰›Ì »«‘œ"
             frmMsg.fwBtn(0).Visible = False
             frmMsg.fwBtn(1).ButtonType = flwButtonOk
             frmMsg.fwBtn(1).Caption = "ﬁ»Ê·"
             frmMsg.Show vbModal
             FlxDetail.TextMatrix(FlxDetail.Row, 1) = 0
        End If
        
        If FlxDetail.TextMatrix(FlxDetail.Row, 1) = 0 And (fwBtnDailyHavale.Visible = False And fwBtnDailyHavale.Enabled = False) Then '
            FlxDetail.RemoveItem (FlxDetail.Row)
            If FlxDetail.Rows < MaxPurchaseRows Then
                AddEmptyRow     'add row Instead of Remove
            End If
            MaxRowFlexGrid = MaxRowFlexGrid - 1
            
            frmMsg.fwlblMsg.Caption = " .ò«·«Ì „Ê—œ ‰Ÿ— «“ ·Ì”  Õ–› ‘œ "
            frmMsg.fwBtn(0).Visible = False
            frmMsg.fwBtn(1).ButtonType = flwButtonOk
            frmMsg.fwBtn(1).Caption = "ﬁ»Ê·"
            frmMsg.Show vbModal
        Else
            FlxDetail.TextMatrix(FlxDetail.Row, 4) = CLng(Val(FlxDetail.TextMatrix(FlxDetail.Row, 1)) * Val(FlxDetail.TextMatrix(FlxDetail.Row, 3)))
        End If
        
        
           
    Else                         'Not Exist in Fich
                      
        If AmountVar <= 0 Then
        
            AmountVar = 0
            frmMsg.fwlblMsg.Caption = " . „ﬁœ«— ﬂ«·« ‰„Ì  Ê«‰œ ’›— Ì« „‰›Ì »«‘œ"
            frmMsg.fwBtn(0).Visible = False
            frmMsg.fwBtn(1).ButtonType = flwButtonOk
            frmMsg.fwBtn(1).Caption = "ﬁ»Ê·"
            frmMsg.Show vbModal
            
            lblNum.Caption = ""
            txtScale.Text = ""
            Exit Function
        
        End If
         
        FlxDetail.Row = MaxRowFlexGrid
        
        FlxDetail.TextMatrix(FlxDetail.Row, 0) = MaxRowFlexGrid
        FlxDetail.TextMatrix(FlxDetail.Row, 1) = AmountVar
        
        FlxDetail.TextMatrix(FlxDetail.Row, 5) = mvarGoodCode
        FlxDetail.TextMatrix(FlxDetail.Row, 8) = mvarServePlace
        FlxDetail.TextMatrix(FlxDetail.Row, 10) = 0  ' mvardiscount
        FlxDetail.TextMatrix(FlxDetail.Row, 11) = 1
        FlxDetail.TextMatrix(FlxDetail.Row, 13) = mvarInventoryNo
        FlxDetail.TextMatrix(FlxDetail.Row, 14) = IIf(DestInventoryNo = 0, "", DestInventoryNo)
        FlxDetail.TextMatrix(FlxDetail.Row, 15) = mvarNumberOfUnit
        

        FlxDetail.ShowCell FlxDetail.Row, 0
        
        FlxDetail.TextMatrix(FlxDetail.Row, 2) = mvarGoodName
        FlxDetail.TextMatrix(FlxDetail.Row, 6) = mvarGoodWeight
        FlxDetail.TextMatrix(FlxDetail.Row, 3) = mvarBuyPrice
        FlxDetail.TextMatrix(FlxDetail.Row, 7) = mvarUnitGood
        
        FlxDetail.TextMatrix(FlxDetail.Row, 16) = mvarMojodi
        FlxDetail.TextMatrix(FlxDetail.Row, 17) = mvarSellPrice
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
        
        If FlxDetail.TextMatrix(FlxDetail.Row, 3) = "" Then
           FlxDetail.TextMatrix(FlxDetail.Row, 1) = ""
        End If
        
        FlxDetail.TextMatrix(FlxDetail.Row, 19) = mvarDuty
        FlxDetail.TextMatrix(FlxDetail.Row, 20) = mvarTax
        On Error GoTo ErrHandler
        
        FlxDetail.TextMatrix(FlxDetail.Row, 4) = CLng(Val(FlxDetail.TextMatrix(FlxDetail.Row, 1)) * Val(FlxDetail.TextMatrix(FlxDetail.Row, 3)))
        On Error GoTo 0
        
        If FlxDetail.Row = (FlxDetail.Rows - 1) Then
           AddEmptyRow
           'FlxDetail.Row = FlxDetail.Row - 1
        End If
        
        FlxDetail.Row = FlxDetail.Row + 1       'Next Row
        MaxRowFlexGrid = FlxDetail.Row            'Last Row

    End If
    
    FlxDetail.Row = MaxRowFlexGrid     'Last Row
    
    lblNum.Caption = ""
    RefreshLables  'Set Lables
    
    FlxDetail.TopRow = FlxDetail.Rows - (MaxPurchaseRows - 1)
    
    txtScale.Text = ""
    
    Exit Function
    
ErrHandler:
    Select Case err.Number
        Case 6
            MsgBox "„ﬁœ«— ò«·«Ì Ê«—œ ‘œÂ »Ì‘ — «“ „ﬁœ«—Ì”  òÂ »—‰«„Â „Ì  Ê«‰œ ﬁ»Ê· ò‰œ " & vbCrLf & "·ÿ›« Ìò ⁄œœ òÊçò — Ê«—œ ‰„«ÌÌœ"
            Cancel
    End Select
    
End Function

Public Sub ChangeLanguage()
    
    Dim Rst As New ADODB.Recordset
    
    With FlxDetail
        Select Case clsStation.Language
            Case EnumLanguage.Farsi
            
                mdifrm.Caption = "                                                " & clsArya.Company
                BtnMenu(0).Caption = "Ã” ÃÊÌ ﬂ«·«"
    ''''            If mvarStatus = Purchase Then
    ''''               FWLabel1.Caption = "›«ﬂ Ê— Œ—Ìœ"
    ''''            Else
    ''''               FWLabel1.Caption = " ÷«Ì⁄«  "
    ''''            End If
                lblDate.Caption = " «—ÌŒ ›«ﬂ Ê—"
                fwBtnCustFind.Caption = " «„Ì‰ ò‰‰œÂ"
                BtnKalaDelete.Caption = "Õ–› ò«·«"
                SumPriceLabel.Caption = "„»·€"
                CmdDiscountTotal.Caption = " Œ›Ì›"
                cmbCarryFeeTotal.Caption = "ò—«ÌÂ Õ„·"
                cmbPackingTotal.Caption = "»” Â »‰œÌ"
              '  cmbServiceTotal.Caption = "”—ÊÌ”"
                fwCash.Caption = "«Ì” ê«Â ‘„«—Â " & clsArya.StationNo
                
                ReDim Parameter(1) As Parameter
                Parameter(0) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
                Parameter(1) = GenerateInputParameter("@PartitionID", adInteger, 4, clsStation.PartitionID)
                
                Set Rst = RunParametricStoredProcedure2Rec("RetrivePartitionDescription", Parameter)
                
                .Font.Name = "Nazanin"
                .Font.Size = 14
                .Font.Bold = True
                .ForeColor = &H80&
        
            
                .RightToLeft = True
                .TextMatrix(0, 0) = "—œÌ›"
                .TextMatrix(0, 1) = "„ﬁœ«—"
                .TextMatrix(0, 2) = "‰«„ ò«·«"
                .TextMatrix(0, 3) = "›Ì"
                .TextMatrix(0, 4) = "Ã„⁄"
                .TextMatrix(0, 5) = "òœ ò«·«"
                .TextMatrix(0, 7) = "Ê«Õœ ﬂ«·«"
                .TextMatrix(0, 8) = "‰Ê⁄ ”—Ê"
                .TextMatrix(0, 10) = " Œ›Ì›"
                .TextMatrix(0, 11) = "‰—Œ"
                .TextMatrix(0, 12) = "«‰ﬁ÷«¡"
                .TextMatrix(0, 13) = "„»œ«"
                .TextMatrix(0, 14) = "„ﬁ’œ"
                .TextMatrix(0, 15) = " ⁄œ«œÊ«Õœ"
                .TextMatrix(0, 16) = "„ÊÃÊœÌ"
                .TextMatrix(0, 17) = "›Ì ›—Ê‘"
                .TextMatrix(0, 18) = "Œ—Ìœﬁ»·Ì"
                .TextMatrix(0, 19) = "⁄Ê«—÷"
                .TextMatrix(0, 20) = "„«·Ì« "
            
            Case EnumLanguage.English
            
                mdifrm.Caption = "                                                " & clsArya.LatinCompany
                BtnMenu(0).Caption = "Search"
    ''''            If mvarStatus = Purchase Then
    ''''               FWLabel1.Caption = "purchase"
    ''''            Else
    ''''               FWLabel1.Caption = "Losses"
    ''''            End If
                lblDate.Caption = "purchase Date"
                fwBtnCustFind.Caption = "Supplier"
                BtnKalaDelete.Caption = "Delete Item"
                SumPriceLabel.Caption = "Price"
                CmdDiscountTotal.Caption = "Discount"
                cmbCarryFeeTotal.Caption = "Shipping"
                cmbPackingTotal.Caption = "Packing"
              '  cmbServiceTotal.Caption = "Service"
                
                fwCash.Caption = "Station #" & clsArya.StationNo
                
                .RightToLeft = False
                .TextMatrix(0, 0) = "Row"
                .TextMatrix(0, 1) = "Amount"
                .TextMatrix(0, 2) = "Good Name"
                .TextMatrix(0, 3) = "Unit Price"
                .TextMatrix(0, 4) = "Price"
                .TextMatrix(0, 5) = "GoodCode"
                .TextMatrix(0, 7) = "Unit"
                .TextMatrix(0, 8) = "ServePlace"
                .TextMatrix(0, 10) = "Discount"
                .TextMatrix(0, 11) = "Rate"
                .TextMatrix(0, 12) = "Expire"
                .TextMatrix(0, 13) = "From"
                .TextMatrix(0, 14) = "To Store"
                .TextMatrix(0, 15) = "NumberOfUnit"
                .TextMatrix(0, 16) = "Stock"
                .TextMatrix(0, 17) = "SalePrice"
                .TextMatrix(0, 18) = "PreviousBuy"
                .TextMatrix(0, 19) = "Duty"
                .TextMatrix(0, 20) = "Tax"
                
        End Select
        
        .ColFormat(3) = "###,###"
        .ColFormat(4) = "###,###"
        .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter
        .ColAlignment(-1) = flexAlignCenterCenter
        .ColAlignment(2) = flexAlignRightCenter
    End With
    
    Dim strTemp As String
    
    
    ReDim Parameter(1) As Parameter
    Parameter(0) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
    Parameter(1) = GenerateInputParameter("@Type", adInteger, 4, 0) 'All Inventory
    Set rctmp = RunParametricStoredProcedure2Rec("GetInventory", Parameter)
    
    FlxDetail.ColComboList(13) = FlxDetail.BuildComboList(rctmp, "Description", "InventoryNo")
    
    ReDim Parameter(1) As Parameter
    Parameter(0) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
    Parameter(1) = GenerateInputParameter("@Type", adInteger, 4, 1) 'All Inventory In Branch And Central Branch
    Set rctmp = RunParametricStoredProcedure2Rec("GetInventory", Parameter)
    FlxDetail.ColComboList(14) = FlxDetail.BuildComboList(rctmp, "Description", "InventoryNo")
    
    rctmp.Close
    
    ReDim Parameter(0) As Parameter
    Parameter(0) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
    Set rctmp = RunParametricStoredProcedure2Rec("GetUnitGood", Parameter)
    
    FlxDetail.ColComboList(7) = FlxDetail.BuildComboList(rctmp, "Description", "Code")
        
    If Rst.State = 1 Then Rst.Close
    
    ReDim Parameter(0) As Parameter
    Parameter(0) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
    Set Rst = RunParametricStoredProcedure2Rec("Get_Serveplace", Parameter)
    
    strTemp = FlxDetail.BuildComboList(Rst, "Description", "intServePlace")
    FlxDetail.ColComboList(8) = strTemp
    If Rst.State <> 0 Then Rst.Close

    UpdatelblServePlace
   ' ValueBtnMenu
    SetFirstToolBar
    
    Set Rst = Nothing
    
End Sub

Sub AddEmptyRow()

    With FlxDetail
        .Rows = .Rows + 1
    End With
    
End Sub

Private Function FindRecord_FlexGrid(TempGoodCode As Double) As Boolean

    FindRecord_FlexGrid = False
    Dim jj As Integer
    If TempGoodCode = Val(FlxDetail.TextMatrix(FlxDetail.Row, 5)) And (mvarServePlace = Val(FlxDetail.TextMatrix(FlxDetail.Row, 8)) Or Val(lblNum.Caption) < 0) Then
        FindRecord_FlexGrid = True
        Exit Function
    End If
    For jj = 1 To FlxDetail.Rows - 1
        If TempGoodCode = Val(FlxDetail.TextMatrix(jj, 5)) And (mvarServePlace = Val(FlxDetail.TextMatrix(jj, 8)) And (Val(FlxDetail.TextMatrix(jj, 3)) = mvarBuyPrice)) Then
     '   If (TempGoodCode = Val(FlxDetail.TextMatrix(jj, 5))) And (mvarServePlace = Val(FlxDetail.TextMatrix(jj, 8))) And (Val(FlxDetail.TextMatrix(jj, 3)) = mvarSellPrice And FlxDetail.TextMatrix(jj, 10) = "") Then
            FindRecord_FlexGrid = True
            FlxDetail.Row = jj
            Exit Function
        End If
    Next
    
    
    
End Function

Sub ClearDataFlexGrid()

    With FlxDetail
        .Rows = 1
        .Rows = MaxPurchaseRows
        .Row = 1
        MaxRowFlexGrid = 1
                
    End With

    
End Sub

Sub GetDataDetail()
On Error Resume Next
    Dim Rst As New ADODB.Recordset
    
    cmbDestination.ListIndex = -1
    ClearDataFlexGrid
    txtDescription.Text = ""
    ReDim Parameter(4) As Parameter
    
    Parameter(0) = GenerateInputParameter("@No", adBigInt, 8, Val(txtNo.Text))
    Parameter(1) = GenerateInputParameter("@Status", adInteger, 4, TempStatus)
    Parameter(2) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
    Parameter(3) = GenerateInputParameter("@AccountYear", adSmallInt, 2, AccountYear)
    Parameter(4) = GenerateInputParameter("@Branch", adInteger, 4, cmbBranch.ItemData(cmbBranch.ListIndex))
    
    Set Rst = RunParametricStoredProcedure2Rec("Get_FacMD_Good", Parameter)
    
    Dim ii As Integer
    
    If Not (Rst.BOF Or Rst.EOF) Then
        If Not IsNull(Rst!DestinationId) Then
            For ii = 0 To cmbDestination.ListCount - 1
                If cmbDestination.ItemData(ii) = Rst!DestinationId Then
                    cmbDestination.ListIndex = ii
                    ii = 0
                    Exit For
                End If
            Next ii
        Else
            cmbDestination.ListIndex = -1
        End If
    End If
    If Not (Rst.BOF Or Rst.EOF) Then
        If Not IsNull(Rst!intInventoryNo) Then
            For ii = 0 To cmbInventory.ListCount - 1
                If cmbInventory.ItemData(ii) = Rst!intInventoryNo Then
                    cmbInventory.ListIndex = ii
                    ii = 0
                    Exit For
                End If
            Next ii
        Else
            cmbInventory.ListIndex = -1
        End If
        If Not IsNull(Rst!DestInventoryNo) Then
            For ii = 0 To cmbDestInventory.ListCount - 1
                If cmbDestInventory.ItemData(ii) = Rst!DestInventoryNo Then
                    cmbDestInventory.ListIndex = ii
                    ii = 0
                    Exit For
                End If
            Next ii
            cmbDestInventory.Enabled = True
        Else
            cmbDestInventory.ListIndex = -1
            cmbDestInventory.Enabled = False
        End If
        txtDescription.Text = IIf(IsNull(Rst!NvcDescription), "", Rst!NvcDescription)
        TmpGoodDiscount = 0
        boolPayment = Rst!FacPayment
        ii = 0
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
            .TextMatrix(ii, 10) = Rst!Discount
            .TextMatrix(ii, 11) = Rst!Rate
            .TextMatrix(ii, 12) = IIf(IsNull(Rst!ExpireDate), "", Rst!ExpireDate)
            .TextMatrix(ii, 13) = Rst!intInventoryNo
            .TextMatrix(ii, 14) = IIf(IsNull(Rst!DestInventoryNo), "", Rst!DestInventoryNo)
            .TextMatrix(ii, 15) = Rst!NumberOfUnit
            TmpGoodDiscount = TmpGoodDiscount + (Rst!Discount * Rst!amount * Rst!FeeUnit / 100)
            
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
            .TextMatrix(ii, 19) = Rst!DutyBuy
            .TextMatrix(ii, 20) = Rst!TaxBuy
            
            Rst.MoveNext
            MaxRowFlexGrid = ii + 1
            If ii >= FlxDetail.Rows - 2 And Rst.EOF = False Then
                AddEmptyRow
            End If

        Loop
        End With
        
        FlxDetail.Row = MaxRowFlexGrid - 1
        mvarServePlace = FlxDetail.TextMatrix(MaxRowFlexGrid - 1, 8)
    End If
    If Rst.State <> 0 Then Rst.Close
    
        
    ReDim Parameter(3) As Parameter
    
    Parameter(0) = GenerateInputParameter("@No", adBigInt, 8, Val(txtNo.Text))
    Parameter(1) = GenerateInputParameter("@Status", adInteger, 4, TempStatus)
    Parameter(2) = GenerateInputParameter("@AccountYear", adSmallInt, 2, AccountYear)
    Parameter(3) = GenerateInputParameter("@Branch", adInteger, 4, cmbBranch.ItemData(cmbBranch.ListIndex))
    
    Set Rst = RunParametricStoredProcedure2Rec("Get_FacM_Per", Parameter)
    
    sbrFactorProp.Panels(1).Text = ""
    sbrFactorProp.Panels(2).Text = ""
    sbrFactorProp.Panels(3).Text = ""
    sbrFactorProp.Panels(4).Text = ""
    
    
    If Not (Rst.EOF = True And Rst.BOF = True) Then
    
        chKTax.Value = Rst!Rasmi
        chKTax_Click
        dblFichUser = Rst.Fields("User").Value
        intSerialNo = Rst.Fields("intSerialNo").Value
        mvarStationNo = Rst.Fields("StationId").Value
        BitTempReceived = Rst.Fields("BitTempReceived").Value
        Select Case clsStation.Language
            Case EnumLanguage.Farsi
            
                sbrFactorProp.Panels(1).Text = "ò«—»—" & " : " & Rst.Fields("nvcFirstName").Value & " " & Rst.Fields("nvcSurName").Value
                sbrFactorProp.Panels(2).Text = "”«⁄  : " & Rst.Fields("Time").Value
                sbrFactorProp.Panels(3).Text = "‘Ì›  : " & Rst.Fields("ShiftDescription").Value
                sbrFactorProp.Panels(3).Tag = Rst.Fields("ShiftNo").Value
                sbrFactorProp.Panels(4).Text = "«Ì” ê«Â : " & Rst.Fields("StationId").Value
            
            Case EnumLanguage.English
            
                sbrFactorProp.Panels(1).Text = "User : " & Rst.Fields("nvcFirstName").Value & " " & Rst.Fields("nvcSurName").Value
                sbrFactorProp.Panels(2).Text = "Time : " & Rst.Fields("Time").Value
                sbrFactorProp.Panels(3).Text = "Shift : " & Rst.Fields("ShiftNo").Value
                sbrFactorProp.Panels(3).Tag = Rst.Fields("ShiftNo").Value
                sbrFactorProp.Panels(4).Text = "StationId :" & Rst.Fields("StationId").Value
        End Select
        Refrence_Acc = IIf(IsNull(Rst!Refrence_Acc), 0, Rst!Refrence_Acc)
        
        FWChkAcc.Value = Rst!TransferAccounting
        If FWChkAcc.Value Then LblAccNo.Caption = Refrence_Acc
                        
    End If
    Set Rst = Nothing
End Sub

Sub DefaultValueLables()
    
    cmbDestination.ListIndex = -1
    Refrence_Acc = 0
    LblAccNo.Caption = ""
    FWChkAcc.Value = 0
    lblTaxTotal.Caption = 0
    LblDutyTotal.Caption = 0
    txtSumCountNo.Caption = 0
    txtSumCountWeight.Caption = 0
    txtSumWeightTotal.Caption = 0
    txtSumFeeTotal.Text = 0
    txtDiscount.Text = 0
    lblDiscountTotal.Caption = 0
    txtDiscountPercent.Text = 0
    txtCarryFee.Text = 0
    txtCarryFeePercent.Text = 0
    lblCarryFeeTotal.Caption = 0
    lblPackingTotal.Caption = 0
    txtPacking.Text = 0
    txtPackingPercent.Text = 0
     
    
    lblSumPrice.Caption = 0
    lblSumPrice.Tag = 0
    
    txtCustomer.Tag = -1
    UpdatetxtCustomer
    
    For i = 1 To sbrFactorProp.Panels.Count
        sbrFactorProp.Panels(i).Text = ""
    Next i
    TempStatus = mvarStatus
    cmbDestination.ListIndex = -1
    Tafsili_3 = 0
    cmbDestInventory.Enabled = True
        chKTax.Value = False
End Sub

Public Sub DefaultStatusbar()

End Sub
Public Sub SetFirstToolBar()

    AllButton vbOff, True

    mdifrm.Toolbar1.Buttons(1).Enabled = True   'Home
    mdifrm.Toolbar1.Buttons(2).Enabled = True   'PageUp
    mdifrm.Toolbar1.Buttons(3).Enabled = True   'PageDown
    mdifrm.Toolbar1.Buttons(4).Enabled = True   'End
    
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
    mdifrm.Toolbar1.Buttons(10).Enabled = True   'Delete
    mdifrm.Toolbar1.Buttons(18).Enabled = True   'Reffer
    If mvarStatus = TempRecieved Then mdifrm.Toolbar1.Buttons(8).Enabled = True
ElseIf MyFormAddEditMode = AddMode Then    'Add Mode
    
    mdifrm.Toolbar1.Buttons(6).Enabled = False  'Add
    mdifrm.Toolbar1.Buttons(7).Enabled = False  'Edit
    mdifrm.Toolbar1.Buttons(8).Enabled = True   'Enter
    mdifrm.Toolbar1.Buttons(9).Enabled = True   'Esc
    mdifrm.Toolbar1.Buttons(10).Enabled = False   'Delete
    mdifrm.Toolbar1.Buttons(18).Enabled = False   'Reffer
    chKTax.Enabled = True

ElseIf MyFormAddEditMode = EditMode Or MyFormAddEditMode = ManipulateMode Then     'Edit
  
    mdifrm.Toolbar1.Buttons(6).Enabled = True  'Add
    mdifrm.Toolbar1.Buttons(7).Enabled = False  'Edit
    mdifrm.Toolbar1.Buttons(8).Enabled = True   'Enter
    mdifrm.Toolbar1.Buttons(9).Enabled = True   'Esc
    mdifrm.Toolbar1.Buttons(10).Enabled = False   'Delete
    mdifrm.Toolbar1.Buttons(18).Enabled = True   'Reffer
    chKTax.Enabled = True

End If

HeaderLabel Val(MyFormAddEditMode), fwlblMode
  
End Sub


Private Sub UpdatetxtCustomer()

    If txtCustomer.Tag <> "" Then
        Dim Rst As New ADODB.Recordset
        Dim mvarMemberShipId, mvarTel, mvarAddress, mvarDescription As String
        
        fwScrollTextCust.Caption = ""
        fwStatusBarCust.Caption = ""
        
        ReDim Parameter(0) As Parameter
        Parameter(0) = GenerateInputParameter("@Code", adInteger, 4, Val(txtCustomer.Tag))
        Set Rst = RunParametricStoredProcedure2Rec("Get_vw_Suppliers", Parameter)
        
        If Rst.EOF = False And Rst.BOF = False Then
            
            Tafsili = Val(IIf(IsNull(Rst!Tafsili), 0, Rst!Tafsili))
            If clsArya.ExternalAccounting = True And Val(txtCustomer.Tag) > 0 And Tafsili = 0 Then
                ShowMessage "«» œ« »—«Ì «Ì‰  «„Ì‰ ﬂ‰‰œÂ œ— ”Ì” „ Õ”«»œ«—Ì ,  ›÷Ì·Ì «ÌÃ«œ ﬂ‰Ìœ", True, False, " «∆Ìœ", ""
                txtCustomer.Tag = 0
                Exit Sub
            End If
            mvarTel = ""
            If Rst.Fields("tel1") <> "" Then
                    mvarTel = " ...  ·›‰ : " + Rst.Fields("tel1")
            End If
            If Rst.Fields("tel2") <> "" Then
                    mvarTel = mvarTel + " ; " + Rst.Fields("tel2")
            End If
            If Rst.Fields("FullAddress") <> "" Then
                    mvarAddress = " ... ¬œ—” : " & Rst.Fields("FullAddress")
            End If
            
            txtCustomer.Text = Rst.Fields("FullName")
            mvarMemberShipId = "«‘ —«ﬂ : " & Rst.Fields("MemberShipId")
            mvarDescription = Rst.Fields("Description")
            If Rst.Fields("Code") <> -1 Then
                fwScrollTextCust.Caption = mvarDescription
                If mvarDescription <> "" Then
                    'fwScrollTextCust.Visible = True
                End If
                fwStatusBarCust.Caption = mvarMemberShipId & mvarTel & mvarAddress
            Else
                fwScrollTextCust.Caption = ""
                fwStatusBarCust.Caption = ""
            End If
                        
        End If
        Set Rst = Nothing
    End If
End Sub

Private Function InsertAutoHavale() As Long
    Dim Result As Long
    ReDim Parameter(9) As Parameter
    Parameter(0) = GenerateInputParameter("@Branch", adInteger, 4, CurrentBranch)
    Parameter(1) = GenerateInputParameter("@InventoryNo", adInteger, 4, cmbInventory.ItemData(cmbInventory.ListIndex))
    Parameter(2) = GenerateInputParameter("@AccountYear", adSmallInt, 2, AccountYear)
    Parameter(3) = GenerateInputParameter("@Status", adInteger, 4, 2)
    Parameter(4) = GenerateInputParameter("@FromDate", adVarWChar, 8, FromDate)
    Parameter(5) = GenerateInputParameter("@ToDate", adVarWChar, 8, ToDate)
    Parameter(6) = GenerateInputParameter("@Date", adVarWChar, 8, txtDate)
    Parameter(7) = GenerateInputParameter("@User", adInteger, 4, mvarCurUserNo)
    Parameter(8) = GenerateInputParameter("@NvcDescription", adVarWChar, 150, Right(txtDescription.Text, 150))
    Parameter(9) = GenerateOutputParameter("@Result", adInteger, 4)
    Result = RunParametricStoredProcedure("Insert_AutoHavale", Parameter)
    If Result > 0 Then
        Select Case mvarStatus
            Case 6
                ShowMessage "«‰ ﬁ«· ÕÊ«·Â Â«Ì —Ê“«‰Â ›—Ê‘ «‰Ã«„ ‘œ ", True, False, " «∆Ìœ", ""
            Case 7
                ShowMessage "«‰ ﬁ«· —”Ìœ Â«Ì —Ê“«‰Â »—ê‘  «“ ›—Ê‘ «‰Ã«„ ‘œ.", True, False, " «∆Ìœ", ""
        End Select
    Else
        ShowMessage "œ—«‰ ﬁ«· ÕÊ«·Â Ì« —”Ìœ „‘ﬂ· ÊÃÊœ œ«—œ.", True, False, " «∆Ìœ", ""
    End If
    InsertAutoHavale = Result
End Function

