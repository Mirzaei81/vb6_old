VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{7AEDC602-D94C-11D1-BB7A-00E0290EA3C9}#1.0#0"; "ResizeKit.ocx"
Object = "{75D4F767-8785-11D3-93AD-0000832EF44D}#2.14#0"; "FAST2001.ocx"
Begin VB.Form frmSupplierComp 
   Caption         =   "                                                                                                                   ÊÌ“Ì Ê—Â«"
   ClientHeight    =   9180
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11130
   Icon            =   "frmSupplierComp.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9180
   ScaleWidth      =   11130
   Begin VB.Frame frameAccounting 
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
      Height          =   3315
      Left            =   3720
      RightToLeft     =   -1  'True
      TabIndex        =   31
      Top             =   120
      Width           =   3135
      Begin VB.TextBox txtPrimaryBedehi 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   480
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   36
         ToolTipText     =   "„»·€ »œÂÌ »œÊ‰ ⁄·«„  Ê „»·€ »” «‰ﬂ«—Ì »« ⁄·«„  „‰›Ì Ê«—œ ‘Êœ"
         Top             =   1560
         Width           =   1455
      End
      Begin VB.TextBox txtPrimaryTalab 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   35
         ToolTipText     =   "„»·€ »œÂÌ »œÊ‰ ⁄·«„  Ê „»·€ »” «‰ﬂ«—Ì »« ⁄·«„  „‰›Ì Ê«—œ ‘Êœ"
         Top             =   2160
         Width           =   1455
      End
      Begin VB.TextBox txtSanadNo 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   34
         Top             =   2760
         Width           =   1455
      End
      Begin VB.TextBox txtAtf 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
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
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   33
         Top             =   480
         Width           =   2175
      End
      Begin VB.TextBox txtTafsiliCode 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   32
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "„«‰œÂ «Ê·ÌÂ - »œÂÌ"
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
         Height          =   525
         Left            =   1440
         TabIndex        =   40
         Top             =   1680
         Width           =   1635
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "„«‰œÂ «Ê·ÌÂ-ÿ·»"
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
         Height          =   525
         Left            =   1440
         TabIndex        =   39
         Top             =   2280
         Width           =   1635
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "‘„«—Â ”‰œ"
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
         Height          =   525
         Left            =   1800
         RightToLeft     =   -1  'True
         TabIndex        =   38
         Top             =   2760
         Width           =   1155
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblTafsili 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
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
         Height          =   525
         Left            =   1920
         RightToLeft     =   -1  'True
         TabIndex        =   37
         Top             =   1200
         Width           =   915
         WordWrap        =   -1  'True
      End
   End
   Begin VB.Frame Frame8 
      Caption         =   "‘⁄»Â"
      BeginProperty Font 
         Name            =   "B Homa"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   3720
      RightToLeft     =   -1  'True
      TabIndex        =   29
      Top             =   3480
      Width           =   3135
      Begin VB.ComboBox cmbBranch 
         Enabled         =   0   'False
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
         Style           =   2  'Dropdown List
         TabIndex        =   30
         Top             =   360
         Width           =   2085
      End
   End
   Begin VB.Frame frameCustomerInfo 
      BeginProperty Font 
         Name            =   "B Homa"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3975
      Left            =   6980
      TabIndex        =   18
      Top             =   495
      Width           =   4080
      Begin VB.TextBox txtFamily 
         Alignment       =   1  'Right Justify
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
         Height          =   495
         Left            =   105
         RightToLeft     =   -1  'True
         TabIndex        =   23
         Top             =   2760
         Width           =   2790
      End
      Begin VB.ComboBox cmbPrefix 
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
         Height          =   480
         ItemData        =   "frmSupplierComp.frx":A4C2
         Left            =   1200
         List            =   "frmSupplierComp.frx":A4C4
         RightToLeft     =   -1  'True
         TabIndex        =   22
         Top             =   1560
         Width           =   1695
      End
      Begin VB.TextBox txtName 
         Alignment       =   1  'Right Justify
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
         Height          =   495
         Left            =   720
         RightToLeft     =   -1  'True
         TabIndex        =   21
         Top             =   2145
         Width           =   2175
      End
      Begin VB.ComboBox cmbGender 
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
         Height          =   480
         ItemData        =   "frmSupplierComp.frx":A4C6
         Left            =   720
         List            =   "frmSupplierComp.frx":A4C8
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Top             =   960
         Width           =   2175
      End
      Begin VB.TextBox txtMembershipId 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFC0C0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Nazanin"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   720
         RightToLeft     =   -1  'True
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   360
         Width           =   2175
      End
      Begin VB.Label lblFamily 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "‰«„ Œ«‰Ê«œêÌ"
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   495
         Left            =   2820
         RightToLeft     =   -1  'True
         TabIndex        =   28
         Top             =   2760
         Width           =   1155
      End
      Begin VB.Label lblName 
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
         ForeColor       =   &H80000002&
         Height          =   495
         Left            =   2820
         RightToLeft     =   -1  'True
         TabIndex        =   27
         Top             =   2190
         Width           =   1155
      End
      Begin VB.Label lblSex 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
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
         ForeColor       =   &H80000002&
         Height          =   495
         Left            =   2820
         RightToLeft     =   -1  'True
         TabIndex        =   26
         Top             =   930
         Width           =   1155
      End
      Begin VB.Label lblCode 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "«‘ —«ò"
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   495
         Left            =   2820
         RightToLeft     =   -1  'True
         TabIndex        =   25
         Top             =   360
         Width           =   1155
      End
      Begin VB.Label lblPrefix 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "⁄‰Ê«‰"
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   495
         Left            =   2820
         RightToLeft     =   -1  'True
         TabIndex        =   24
         Top             =   1620
         Width           =   1155
      End
   End
   Begin VB.Frame frameContactInfo 
      BeginProperty Font 
         Name            =   "B Homa"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3735
      Left            =   240
      TabIndex        =   4
      Top             =   720
      Width           =   3400
      Begin VB.TextBox txtDiscount 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   1830
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   3135
         Visible         =   0   'False
         Width           =   585
      End
      Begin VB.TextBox txtUnit 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   240
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   2580
         Width           =   2175
      End
      Begin VB.TextBox txtFlour 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   240
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   2025
         Width           =   2175
      End
      Begin VB.TextBox txtMobile 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   240
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   930
         Width           =   2175
      End
      Begin VB.TextBox txtTel 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   240
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   375
         Width           =   2175
      End
      Begin VB.TextBox txtintTel 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   240
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   1470
         Width           =   2175
      End
      Begin VB.Label lblDiscount2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "œ—’œ"
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
         Left            =   1080
         RightToLeft     =   -1  'True
         TabIndex        =   17
         Top             =   3180
         Visible         =   0   'False
         Width           =   705
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   " Œ›Ì›"
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   465
         Left            =   2430
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Top             =   3150
         Visible         =   0   'False
         Width           =   705
      End
      Begin VB.Label lblUnit 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Ê«Õœ"
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   465
         Left            =   2430
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   2595
         Width           =   705
      End
      Begin VB.Label lblPart 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "ÿ»ﬁÂ"
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   465
         Left            =   2430
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   2040
         Width           =   705
      End
      Begin VB.Label lblMobile 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "„Ê»«Ì·"
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   465
         Left            =   2430
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   930
         Width           =   705
      End
      Begin VB.Label lblTel 
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
         ForeColor       =   &H80000002&
         Height          =   465
         Left            =   2430
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   375
         Width           =   705
      End
      Begin VB.Label lblintTel 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "œ«Œ·Ì"
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   465
         Left            =   2430
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   1485
         Width           =   705
      End
   End
   Begin VSFlex7LCtl.VSFlexGrid vsMembers 
      Height          =   4545
      Left            =   240
      TabIndex        =   0
      Top             =   4560
      Width           =   10845
      _cx             =   19129
      _cy             =   8017
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
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   16761024
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   -2147483643
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
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   11
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   500
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmSupplierComp.frx":A4CA
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
      ExplorerBar     =   0
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
   Begin FLWCtrls.FWLabel3D fwlblMode 
      Height          =   480
      Left            =   9705
      Top             =   15
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   847
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
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
   Begin VB.TextBox txtCode 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFC0C0&
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
      Left            =   240
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   240
      Visible         =   0   'False
      Width           =   2055
   End
   Begin RESIZEKITLibCtl.ResizeKit ResizeKit1 
      Height          =   480
      Left            =   4320
      OleObjectBlob   =   "frmSupplierComp.frx":A5BA
      TabIndex        =   3
      Top             =   120
      Width           =   480
   End
   Begin VB.Label lblNo 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "—œÌ›"
      BeginProperty Font 
         Name            =   "B Homa"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   315
      Left            =   2385
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   255
      Visible         =   0   'False
      Width           =   705
   End
End
Attribute VB_Name = "frmSupplierComp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cn As New ADODB.Connection
Dim MyFormAddEditMode As EnumAddEditMode
Dim i As Integer
Dim Parameter() As Parameter
Dim OldTafsili As Long
Dim frmact As Form

Public Sub Add()
    MyFormAddEditMode = AddMode
    
    DefaultSettings
    SetFirstToolBar
    FillvsMembers
End Sub

Public Sub Cancel()
    Add
    
End Sub

Public Sub DefaultSettings()

    txtPrimaryBedehi = ""
    txtPrimaryTalab = ""
    txtFamily.Text = ""
    txtFlour.Text = ""
    txtintTel.Text = ""
   ' txtMembershipId = ""
    txtMobile.Text = ""
    TxtName.Text = ""
    txtTel.Text = ""
    txtUnit.Text = ""
    txtDiscount.Text = ""
    
    cmbGender.ListIndex = 0
    cmbPrefix.ListIndex = 0
    txtTafsiliCode.Text = ""
    OldTafsili = 0
    OldAtf = 0
End Sub
Public Sub Delete()
    On Error GoTo ErrHandler
    ReDim Parameter(1) As Parameter
    Parameter(0) = GenerateInputParameter("@Code", adInteger, 4, txtMembershipId.Tag)
    Parameter(1) = GenerateInputParameter("@Branch", adInteger, 4, CurrentBranch)
    RunParametricStoredProcedure "Delete_Supplier", Parameter
    FillvsMembers
    Add

ErrHandler:
    Select Case err.Number
        Case -2147217873
            
            frmMsg.fwlblMsg.Caption = "›«ò Ê—Â«ÌÌ œ— —«»ÿÂ »« «Ì‰ ÊÌ“Ì Ê— ÊÃÊœ œ«—œ" & vbCrLf + " ‘„« ‰„Ì  Ê«‰Ìœ «Ì‰ ÊÌ“Ì Ê— —« Õ–› ò‰Ìœ "
            frmMsg.fwBtn(0).Visible = False
            frmMsg.fwBtn(1).ButtonType = flwButtonOk
            frmMsg.fwBtn(1).Caption = "ﬁ»Ê·"
            frmMsg.Show vbModal
    
    End Select
    
End Sub
Public Sub Edit()
    MyFormAddEditMode = EditMode
    SetFirstToolBar

End Sub
Public Sub ExitForm()

    Unload Me
    
End Sub

Public Sub GetDataDetail()
    On Error GoTo ErrHandler
    DefaultSettings
    With vsMembers
    
        If .Rows > 1 Then
        
            If .Row > 0 Then
                ReDim Parameter(1) As Parameter
                Dim Rst As ADODB.Recordset
                
                Parameter(0) = GenerateInputParameter("@Code", adInteger, 4, Val(.TextMatrix(.Row, 1)))
                Parameter(1) = GenerateInputParameter("@Branch", adInteger, 4, CurrentBranch)
                Set Rst = RunParametricStoredProcedure2Rec("Get_Supplier_info", Parameter)
                
                txtMembershipId.Tag = Rst!Code '.TextMatrix(.Row, 1)
                TxtName.Text = Rst!Name ' .TextMatrix(.Row, 2)
                txtFamily.Text = Rst!Family ' .TextMatrix(.Row, 3)
                txtTel.Text = Rst!Tel1 ' .TextMatrix(.Row, 5)
                txtMobile.Text = Rst!Mobile ' .TextMatrix(.Row, 7)
                txtintTel.Text = Rst!internalNo '.TextMatrix(.Row, 8)
                txtFlour.Text = Rst!Flour ' .TextMatrix(.Row, 9)
                txtUnit.Text = Rst!Unit ' .TextMatrix(.Row, 10)
                txtDiscount.Text = Rst!Discount
                txtTafsiliCode.Text = IIf(IsNull(Rst!Tafsili), "", Rst!Tafsili)
                OldTafsili = Val(txtTafsiliCode.Text)
                txtSanadNo.Text = IIf(IsNull(Rst!SanadNo), "", Rst!SanadNo)
                
                If IsNull(Rst!TotalRemainingAmount) = False Then
                    If Val(Rst!TotalRemainingAmount) > 0 Then
                        txtPrimaryBedehi.Text = Val(Rst!TotalRemainingAmount)
                    Else
                        txtPrimaryTalab.Text = -1 * Val(Rst!TotalRemainingAmount)
                    End If
                End If
                For i = 0 To cmbGender.ListCount - 1
                    If cmbGender.ItemData(i) = Rst!Sex Then  ' 'Val(.TextMatrix(.Row, 6))
                        cmbGender.ListIndex = i
                        Exit For
                    End If
                Next i
                
                For i = 0 To cmbPrefix.ListCount - 1
                    If cmbPrefix.ItemData(i) = Rst!Prefix Then ' Val(.TextMatrix(.Row, 4))
                        cmbPrefix.ListIndex = i
                        Exit For
                    End If
                Next i
            
            End If
            
        End If
    End With
Exit Sub

ErrHandler:
    ShowDisMessage err.Description, 2000
End Sub

Sub SetFirstToolBar()
    
    Dim Obj As Object
    
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
 
    If MyFormAddEditMode = ViewMode Then  ' View Mode
 
        On Error Resume Next
        For Each Obj In Me
            If TypeOf Obj Is OptionButton Or TypeOf Obj Is FWCheck Then
                Obj.Enabled = False
            ElseIf TypeOf Obj Is TextBox Or TypeOf Obj Is ComboBox Then
                Obj.Locked = True
            End If
        Next Obj
        On Error GoTo 0
        
        mdifrm.Toolbar1.Buttons(6).Enabled = True  'Add
        mdifrm.Toolbar1.Buttons(7).Enabled = True  'Edit
        mdifrm.Toolbar1.Buttons(8).Enabled = False  'Enter
        mdifrm.Toolbar1.Buttons(9).Enabled = False   'Esc
        mdifrm.Toolbar1.Buttons(10).Enabled = True   'Delete
        frameCustomerInfo.Enabled = False
        frameContactInfo.Enabled = False
        txtTafsiliCode.Enabled = False
'        fwlblMode.Caption = "„—Ê—"
        
    ElseIf MyFormAddEditMode = AddMode Then    'Add Mode
        
        On Error Resume Next
        For Each Obj In Me
            If TypeOf Obj Is OptionButton Or TypeOf Obj Is FWCheck Then
                Obj.Enabled = True
            Else
                Obj.Locked = False
            End If
        Next Obj
        On Error GoTo 0
        
        mdifrm.Toolbar1.Buttons(6).Enabled = False  'Add
        mdifrm.Toolbar1.Buttons(7).Enabled = False  'Edit
        mdifrm.Toolbar1.Buttons(8).Enabled = True   'Enter
        mdifrm.Toolbar1.Buttons(9).Enabled = True   'Esc
        mdifrm.Toolbar1.Buttons(10).Enabled = False   'Delete
        frameCustomerInfo.Enabled = True
        frameContactInfo.Enabled = True
'        fwlblMode.Caption = "ÃœÌœ"
    
    ElseIf MyFormAddEditMode = EditMode Then     'Edit
    
        On Error Resume Next
        For Each Obj In Me
            If TypeOf Obj Is OptionButton Or TypeOf Obj Is FWCheck Then
                Obj.Enabled = True
            Else
                Obj.Locked = False
            End If
        Next Obj
        On Error GoTo 0
        
        mdifrm.Toolbar1.Buttons(6).Enabled = True  'Add
        mdifrm.Toolbar1.Buttons(7).Enabled = False  'Edit
        mdifrm.Toolbar1.Buttons(8).Enabled = True   'Enter
        mdifrm.Toolbar1.Buttons(9).Enabled = True   'Esc
        mdifrm.Toolbar1.Buttons(10).Enabled = False   'Delete
        frameCustomerInfo.Enabled = True
        frameContactInfo.Enabled = True
'        fwlblMode.Caption = "«’·«Õ"
    
    End If
    HeaderLabel Val(MyFormAddEditMode), fwlblMode
End Sub

Public Sub Update()
    
    If Trim(txtMembershipId.Text) = "" Or Trim(cmbGender.Text) = "" Or Trim(txtFamily.Text) = "" Then
        frmMsg.fwlblMsg.Caption = "·ÿ›« «ÿ·«⁄«  —« »Â ÿÊ— ò«„· Ê«—œ ‰„«ÌÌœ"
        frmMsg.fwBtn(0).Visible = False
        frmMsg.fwBtn(1).ButtonType = flwButtonOk
        frmMsg.fwBtn(1).Caption = "ﬁ»Ê·"
        frmMsg.Show vbModal
        Exit Sub
    End If
    
    If Val(txtDiscount.Text) < 0 Or Val(txtDiscount.Text) > 100 Then
        frmMsg.fwlblMsg.Caption = "„ﬁœ«—  Œ›Ì› ‰„Ì  Ê«‰œ ò„ — «“ ’›— Ì« »Ì‘ — «“ ’œ œ—’œ »«‘œ "
        frmMsg.fwBtn(0).Visible = False
        frmMsg.fwBtn(1).ButtonType = flwButtonOk
        frmMsg.fwBtn(1).Caption = "ﬁ»Ê·"
        frmMsg.Show vbModal
        Exit Sub
    End If
    Select Case MyFormAddEditMode
    
        Case AddMode
            
            ReDim Parameter(30) As Parameter
            Parameter(0) = GenerateInputParameter("@MembershipId", adInteger, 4, 0)
            Parameter(1) = GenerateInputParameter("@MasterCode", adInteger, 4, txtCode.Text)
            Parameter(2) = GenerateInputParameter("@Owner", adInteger, 4, 1)
            Parameter(3) = GenerateInputParameter("@Name", adVarWChar, 50, TxtName.Text)
            Parameter(4) = GenerateInputParameter("@Family", adVarWChar, 50, txtFamily.Text)
            Parameter(5) = GenerateInputParameter("@Sex", adInteger, 4, cmbGender.ItemData(cmbGender.ListIndex))
            Parameter(6) = GenerateInputParameter("@WorkName", adVarWChar, 50, "")
            Parameter(7) = GenerateInputParameter("@InternalNo", adVarWChar, 50, txtintTel.Text)
            Parameter(8) = GenerateInputParameter("@Unit", adVarWChar, 50, txtUnit.Text)
            Parameter(9) = GenerateInputParameter("@State", adInteger, 4, 0)
            Parameter(10) = GenerateInputParameter("@City", adInteger, 4, 0)
            Parameter(11) = GenerateInputParameter("@ActKind", adInteger, 4, 0)
            Parameter(12) = GenerateInputParameter("@ActDeAct", adInteger, 4, 1)
            Parameter(13) = GenerateInputParameter("@Prefix", adInteger, 4, cmbPrefix.ItemData(cmbPrefix.ListIndex))
            Parameter(14) = GenerateInputParameter("@Address", adVarWChar, 255, "")
            Parameter(15) = GenerateInputParameter("@PostalCode", adVarWChar, 50, "")
            Parameter(16) = GenerateInputParameter("@Tel1", adVarWChar, 50, Trim(txtTel.Text))
            Parameter(17) = GenerateInputParameter("@Tel2", adVarWChar, 50, "")
            Parameter(18) = GenerateInputParameter("@Tel3", adVarWChar, 50, "")
            Parameter(19) = GenerateInputParameter("@Tel4", adVarWChar, 50, "")
            Parameter(20) = GenerateInputParameter("@Mobile", adVarWChar, 50, Trim(txtMobile.Text))
            Parameter(21) = GenerateInputParameter("@Fax", adVarWChar, 50, "")
            Parameter(22) = GenerateInputParameter("@Email", adVarWChar, 50, "")
            Parameter(23) = GenerateInputParameter("@Flour", adVarWChar, 50, txtFlour.Text)
            Parameter(24) = GenerateInputParameter("@Discount", adDouble, 8, Val(txtDiscount.Text))
            Parameter(25) = GenerateInputParameter("@Description", adVarWChar, 255, "")
            Parameter(26) = GenerateInputParameter("@User", adInteger, 4, mvarCurUserNo)
            Parameter(27) = GenerateInputParameter("@TotalRemainingAmount", adDouble, 8, IIf(Val(txtPrimaryBedehi) > 0, Val(txtPrimaryBedehi), -1 * Val(txtPrimaryTalab)))
            Parameter(28) = GenerateInputParameter("@EconomicCode", adVarWChar, 20, Trim(frmSupplier.txtEconomicCode.Text))
            Parameter(29) = GenerateInputParameter("@NationalCode", adVarWChar, 20, Trim(frmSupplier.txtNationalCode.Text))
            Parameter(30) = GenerateOutputParameter("@Code", adBigInt, 8)
            
            Dim LastCode As Long
            LastCode = RunParametricStoredProcedure("Insert_Supplier", Parameter)
            If LastCode <> -1 Then
                frmMsg.fwlblMsg.Caption = "À»  ÊÌ“Ì Ê— ÃœÌœ »« „Ê›ﬁÌ  «‰Ã«„ ‘œ"
                frmMsg.fwBtn(0).Visible = False
                frmMsg.fwBtn(1).ButtonType = flwButtonOk
                frmMsg.fwBtn(1).Caption = "ﬁ»Ê·"
                frmMsg.Show vbModal
                
                If clsArya.ExternalAccounting = True Or HasMiniAcc = True Then
                    Insert_Tafsili LastCode, True
                End If
            Else
                frmMsg.fwlblMsg.Caption = "À»  «‰Ã«„ ‰‘œ" + vbCrLf + "‘„«—Â «‘ —«ò —« »——”Ì ‰„«ÌÌœ."
                frmMsg.fwBtn(0).Visible = False
                frmMsg.fwBtn(1).ButtonType = flwButtonOk
                frmMsg.fwBtn(1).Caption = "ﬁ»Ê·"
                frmMsg.Show vbModal
'                txtMembershipId.SetFocus
                Exit Sub
            End If
            
            
        Case EditMode
            
            ReDim Parameter(31) As Parameter
            Parameter(0) = GenerateInputParameter("@MembershipId", adInteger, 4, 0)
            Parameter(1) = GenerateInputParameter("@MasterCode", adInteger, 4, txtCode.Text)
            Parameter(2) = GenerateInputParameter("@Owner", adInteger, 4, 1)
            Parameter(3) = GenerateInputParameter("@Name", adVarWChar, 50, TxtName.Text)
            Parameter(4) = GenerateInputParameter("@Family", adVarWChar, 50, txtFamily.Text)
            Parameter(5) = GenerateInputParameter("@Sex", adInteger, 4, cmbGender.ItemData(cmbGender.ListIndex))
            Parameter(6) = GenerateInputParameter("@WorkName", adVarWChar, 50, "")
            Parameter(7) = GenerateInputParameter("@InternalNo", adVarWChar, 50, txtintTel.Text)
            Parameter(8) = GenerateInputParameter("@Unit", adVarWChar, 50, txtUnit.Text)
            Parameter(9) = GenerateInputParameter("@State", adInteger, 4, 0)
            Parameter(10) = GenerateInputParameter("@City", adInteger, 4, 0)
            Parameter(11) = GenerateInputParameter("@ActKind", adInteger, 4, 0)
            Parameter(12) = GenerateInputParameter("@ActDeAct", adInteger, 4, 1)
            Parameter(13) = GenerateInputParameter("@Prefix", adInteger, 4, cmbPrefix.ItemData(cmbPrefix.ListIndex))
            Parameter(14) = GenerateInputParameter("@Address", adVarWChar, 255, "")
            Parameter(15) = GenerateInputParameter("@PostalCode", adVarWChar, 50, "")
            Parameter(16) = GenerateInputParameter("@Tel1", adVarWChar, 50, Trim(txtTel.Text))
            Parameter(17) = GenerateInputParameter("@Tel2", adVarWChar, 50, "")
            Parameter(18) = GenerateInputParameter("@Tel3", adVarWChar, 50, "")
            Parameter(19) = GenerateInputParameter("@Mobile", adVarWChar, 50, Trim(txtMobile.Text))
            Parameter(20) = GenerateInputParameter("@Tel4", adVarWChar, 50, "")
            Parameter(21) = GenerateInputParameter("@Fax", adVarWChar, 50, "")
            Parameter(22) = GenerateInputParameter("@Email", adVarWChar, 50, "")
            Parameter(23) = GenerateInputParameter("@Flour", adVarWChar, 50, txtFlour.Text)
            Parameter(24) = GenerateInputParameter("@Discount", adDouble, 8, Val(txtDiscount.Text))
            Parameter(25) = GenerateInputParameter("@Description", adVarWChar, 255, "")
            Parameter(26) = GenerateInputParameter("@User", adInteger, 4, mvarCurUserNo)
            Parameter(27) = GenerateInputParameter("@Code", adBigInt, 8, Val(txtMembershipId.Tag))
            Parameter(28) = GenerateInputParameter("@TotalRemainingAmount", adDouble, 8, IIf(Val(txtPrimaryBedehi) > 0, Val(txtPrimaryBedehi), -1 * Val(txtPrimaryTalab)))
            Parameter(29) = GenerateInputParameter("@EconomicCode", adVarWChar, 20, Trim(frmSupplier.txtEconomicCode.Text))
            Parameter(30) = GenerateInputParameter("@NationalCode", adVarWChar, 20, Trim(frmSupplier.txtNationalCode.Text))
            Parameter(31) = GenerateOutputParameter("@Updated", adBigInt, 8)
            Dim Updated As Long
            Updated = RunParametricStoredProcedure("Update_Supplier", Parameter)
            If Updated <> False Then
                frmMsg.fwlblMsg.Caption = " €ÌÌ—«  »« „Ê›ﬁÌ  «‰Ã«„ ‘œ"
                frmMsg.fwBtn(0).Visible = False
                frmMsg.fwBtn(1).ButtonType = flwButtonOk
                frmMsg.fwBtn(1).Caption = "ﬁ»Ê·"
                frmMsg.Show vbModal
                
                If clsArya.ExternalAccounting = True Or HasMiniAcc = True Then
                    Insert_Tafsili Updated, True
                End If
            Else
                frmMsg.fwlblMsg.Caption = " €ÌÌ—«  «‰Ã«„ ‰‘œ" + vbCrLf + "‘„«—Â «‘ —«ò —« »——”Ì ‰„«ÌÌœ."
                frmMsg.fwBtn(0).Visible = False
                frmMsg.fwBtn(1).ButtonType = flwButtonOk
                frmMsg.fwBtn(1).Caption = "ﬁ»Ê·"
                frmMsg.Show vbModal
                txtMembershipId.SetFocus
                Exit Sub
            End If
    
    End Select
    
    MyFormAddEditMode = AddMode
    SetFirstToolBar
    Add
    
    Exit Sub
RollBack:
    

End Sub
Private Sub FillvsMembers()

    Dim Rst As New ADODB.Recordset
    ReDim Parameter(0) As Parameter
    
    With vsMembers
        .Rows = 1
        Parameter(0) = GenerateInputParameter("@MasterCode", adInteger, 4, Val(txtCode.Text))
        
        Set Rst = RunParametricStoredProcedure2Rec("Get_MemberSupplier", Parameter)
        
        'txtCode.Text = frmact.mvarcode
        If Not (Rst.EOF = True And Rst.BOF = True) Then
            txtMembershipId.Text = Rst!MembershipId
        End If
        i = 0
        While Rst.EOF <> True
            i = i + 1
            .Rows = .Rows + 1
            .TextMatrix(i, 0) = i
            .TextMatrix(i, 1) = Rst.Fields("Code").Value
            .TextMatrix(i, 2) = Rst.Fields("Name").Value
            .TextMatrix(i, 3) = Rst.Fields("Family").Value
            .TextMatrix(i, 4) = Rst.Fields("Prefix").Value
            .TextMatrix(i, 5) = Rst.Fields("Tel1").Value
            .TextMatrix(i, 6) = Rst.Fields("Sex").Value
            .TextMatrix(i, 7) = Rst.Fields("Mobile").Value
            .TextMatrix(i, 8) = Rst.Fields("InternalNo").Value
            .TextMatrix(i, 9) = Rst.Fields("Flour").Value
            .TextMatrix(i, 10) = Rst.Fields("Unit").Value
            Rst.MoveNext
        Wend
    End With
    
    Set Rst = Nothing
End Sub

Private Sub Form_Activate()
    VarActForm = Me.Name
    SetFirstToolBar
    
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

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0

End Sub

Private Sub Form_Load()

    If ClsFormAccess.frmSupplierComp = False Then
        Unload Me
    End If
    
    VarActForm = Me.Name
    
    CenterTop Me
    
    Dim varForm As Form
    For Each varForm In Forms
        If LCase(varForm.Name) = "frmsupplier" Then
            Set frmact = varForm
            Exit For
        End If
    Next
    frmact.Hide
    txtCode.Text = mvarcode 'frmact.mvarcode
    txtMembershipId.Text = frmact.txtMembershipId.Text
    
    vsMembers.Cell(flexcpAlignment, 0, 0, 0, vsMembers.Cols - 1) = flexAlignCenterCenter
    vsMembers.Cell(flexcpAlignment, 0, 0, vsMembers.Rows - 1, 0) = flexAlignCenterCenter
    
    cmbGender.Clear
    Select Case clsStation.Language
    
        Case EnumLanguage.Farsi
        
            cmbGender.AddItem "¬ﬁ«"
            cmbGender.ItemData(cmbGender.ListCount - 1) = 1
            cmbGender.AddItem "Œ«‰„"
            cmbGender.ItemData(cmbGender.ListCount - 1) = 0
            vsMembers.ColComboList(6) = "#1;¬ﬁ«|#0;Œ«‰„"
            
        Case EnumLanguage.English
        
            cmbGender.AddItem "Male"
            cmbGender.ItemData(cmbGender.ListCount - 1) = 1
            cmbGender.AddItem "Female"
            cmbGender.ItemData(cmbGender.ListCount - 1) = 0
            vsMembers.ColComboList(6) = "#1;Male|#2;Female"
            
    End Select
    
    Dim Rst As New ADODB.Recordset
    Set Rst = RunStoredProcedure2RecordSet("Get_All_tPrefix")
    Dim tmpStr As String
    tmpStr = vsMembers.BuildComboList(Rst, "Description", "Code")
    vsMembers.ColComboList(4) = tmpStr
    cmbPrefix.Clear
    Rst.MoveFirst
    While Rst.EOF <> True
        cmbPrefix.AddItem Rst!Description
        cmbPrefix.ItemData(cmbPrefix.ListCount - 1) = Rst!Code
        Rst.MoveNext
    Wend
    If cmbPrefix.ListCount > 0 Then cmbPrefix.ListIndex = 0
    Set Rst = Nothing
    
    With vsMembers
        .TextMatrix(0, 1) = "òœ"
        .TextMatrix(0, 2) = "‰«„"
        .TextMatrix(0, 3) = "‰«„ Œ«‰Ê«œêÌ"
        .TextMatrix(0, 4) = "⁄‰Ê«‰"
        .TextMatrix(0, 5) = " ·›‰"
        .TextMatrix(0, 6) = "Ã‰”Ì "
        .TextMatrix(0, 7) = "„Ê»«Ì·"
        .TextMatrix(0, 8) = "œ«Œ·Ì"
        .TextMatrix(0, 9) = "ÿ»ﬁÂ"
        .TextMatrix(0, 10) = "Ê«Õœ"

         For i = 0 To .Cols - 1
            .ColWidth(i) = Val(GetSetting(strMainKey, "frmSupplierComp_vsMembers", "Col" & i))
            If .ColWidth(i) = 0 Then
                .ColWidth(i) = 1000       '
            End If
         Next i
    
    End With
    
    FillBranch
    
    FillAtf
    
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

    Add
    
    If ClsFormAccess.ChangeTotalRemainingAmount = True Then
        txtPrimaryBedehi.Enabled = True
        txtPrimaryTalab.Enabled = True
    Else
        txtPrimaryBedehi.Enabled = False
        txtPrimaryTalab.Enabled = False
    End If

End Sub
Private Sub FillBranch()
    Dim rctmp As New ADODB.Recordset
    Dim i As Long
    cmbBranch.Clear
    Set rctmp = RunStoredProcedure2RecordSet("Get_All_Branches")
    Do While rctmp.EOF = False
        cmbBranch.AddItem rctmp!nvcBranchName
        cmbBranch.ItemData(cmbBranch.NewIndex) = rctmp!Branch
        rctmp.MoveNext
    Loop
    rctmp.Close
    For i = 0 To cmbBranch.ListCount - 1
        If CurrentBranch = cmbBranch.ItemData(i) Then
            cmbBranch.ListIndex = i
            Exit For
        End If
    Next

End Sub

Private Sub Form_Unload(Cancel As Integer)
    If PosConnection.State = adStateOpen Then PosConnection.Close: Set PosConnection = Nothing
    Set cn = Nothing


    SaveSetting strMainKey, Me.Name, "Left", Me.left
    SaveSetting strMainKey, Me.Name, "Top", Me.top

    frmact.Show
   
End Sub

Private Sub ResizeKit1_ExitResize(ByVal XScale As Double, ByVal YScale As Double)


    If formloadFlag = True Then
        SaveSetting strMainKey, Me.Name, "Height", Me.Height
        SaveSetting strMainKey, Me.Name, "Width", Me.Width
    End If

End Sub

Private Sub vsMembers_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    If Col = -1 Then Exit Sub
    For i = 0 To vsMembers.Cols - 1
        SaveSetting strMainKey, "frmSupplierComp_vsMembers", "Col" & i, vsMembers.ColWidth(i)
    Next

End Sub

Private Sub vsMembers_Click()

    MyFormAddEditMode = ViewMode
    GetDataDetail
    SetFirstToolBar
    HeaderLabel Val(MyFormAddEditMode), Me.fwlblMode

End Sub

Private Sub Insert_Tafsili(CustCode As Long, ShowMessageflag As Boolean)
On Error GoTo ErrHandler
    Dim rs As New ADODB.Recordset
    Dim TafsiliName As String
    TafsiliName = Trim(TxtName.Text) & " " & Trim(txtFamily.Text)
    If txtTafsiliCode.Text = "" Then
        txtTafsiliCode.Text = Accounting.Insert_TafsiliDll(ShowMessageflag, cmbBranch.ItemData(cmbBranch.ListIndex), TafsiliName, EnumAtf.Companies)
    Else
        Accounting.Update_TafsiliDll ShowMessageflag, cmbBranch.ItemData(cmbBranch.ListIndex), Val(txtTafsiliCode.Text), TafsiliName, EnumAtf.Companies
    End If
   
    If (Val(txtPrimaryBedehi) <> 0 Or Val(txtPrimaryTalab) <> 0) And Val(txtSanadNo) = 0 Then
        Accounting.Insert_PrimarySand_Cust CustCode, Val(txtTafsiliCode.Text), Val(txtPrimaryBedehi), Val(txtPrimaryTalab), 1, 1
            
    End If
    If Val(txtTafsiliCode.Text) > 0 Then
        ReDim Parameter(1) As Parameter
        Parameter(0) = GenerateInputParameter("@TafsiliId", adInteger, 4, Val(txtTafsiliCode.Text))
        Parameter(1) = GenerateInputParameter("@Code", adInteger, 4, CustCode)
        RunParametricStoredProcedure "Update_tCust_tafsili", Parameter
    End If
    If ShowMessageflag = False Then Exit Sub
Exit Sub
ErrHandler:
    MsgBox err.Description
    modgl.LogSave "frmSupplier", err, "Insert_Tafsili"
    Resume Next
End Sub


Private Sub FillAtf()
    txtAtf.Text = "«‘Œ«’ Ê ‘—ò Â«"
End Sub
