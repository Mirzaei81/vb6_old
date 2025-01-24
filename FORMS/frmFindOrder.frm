VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{7AEDC602-D94C-11D1-BB7A-00E0290EA3C9}#1.0#0"; "ResizeKit.ocx"
Object = "{75D4F767-8785-11D3-93AD-0000832EF44D}#2.14#0"; "FAST2001.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{687FF23F-0E9B-449D-A782-2C5E7116EDB0}#1.0#0"; "Fardate.ocx"
Begin VB.Form frmFindOrder 
   BackColor       =   &H00E0E0E0&
   Caption         =   "        Ã” ÃÊÌ ”›«—‘"
   ClientHeight    =   8865
   ClientLeft      =   2775
   ClientTop       =   3870
   ClientWidth     =   12435
   Icon            =   "frmFindOrder.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   8865
   ScaleWidth      =   12435
   Begin VB.Timer Timer1 
      Left            =   0
      Top             =   240
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
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
      Height          =   1095
      Left            =   2760
      RightToLeft     =   -1  'True
      TabIndex        =   23
      Top             =   7680
      Width           =   4215
      Begin VB.CheckBox chk_Reprint 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Caption         =   "ç«Å „Ãœœ"
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
         Height          =   255
         Left            =   2040
         RightToLeft     =   -1  'True
         TabIndex        =   26
         Top             =   720
         Width           =   1815
      End
      Begin VB.CheckBox chk_Print 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Caption         =   "ç«Å"
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
         Height          =   255
         Left            =   2040
         RightToLeft     =   -1  'True
         TabIndex        =   25
         Top             =   360
         Width           =   1815
      End
      Begin VB.CommandButton cmdPrint 
         BackColor       =   &H00FF8080&
         Caption         =   "ç«Å"
         Default         =   -1  'True
         BeginProperty Font 
            Name            =   "B Homa"
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
         TabIndex        =   24
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.CommandButton CancelButton 
      BackColor       =   &H000000C0&
      Cancel          =   -1  'True
      Caption         =   "«‰’—«›"
      BeginProperty Font 
         Name            =   "B Homa"
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
      TabIndex        =   10
      Top             =   8040
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      BackColor       =   &H0000C000&
      Caption         =   "ﬁ»Ê·"
      BeginProperty Font 
         Name            =   "B Homa"
         Size            =   11.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1440
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   8040
      Width           =   1215
   End
   Begin VB.CommandButton RepBotton 
      BackColor       =   &H000000C0&
      Caption         =   "ê“«—‘"
      BeginProperty Font 
         Name            =   "B Homa"
         Size            =   12
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      MaskColor       =   &H00C0C0FF&
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   1320
      Width           =   1935
   End
   Begin VB.PictureBox Picture1 
      Height          =   1335
      Left            =   9120
      RightToLeft     =   -1  'True
      ScaleHeight     =   1275
      ScaleWidth      =   3195
      TabIndex        =   5
      Top             =   600
      Width           =   3255
      Begin VB.TextBox txtNo 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0000FFFF&
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
         Left            =   840
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   120
         Width           =   915
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0000C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "‘„«—Â ”›«—‘"
         BeginProperty Font 
            Name            =   "Nazanin"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1800
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   120
         Width           =   1215
      End
   End
   Begin VB.PictureBox Picture2 
      Height          =   1335
      Left            =   5760
      RightToLeft     =   -1  'True
      ScaleHeight     =   1275
      ScaleWidth      =   3195
      TabIndex        =   1
      Top             =   600
      Width           =   3255
      Begin VB.OptionButton optbalance 
         Alignment       =   1  'Right Justify
         Caption         =   "ò·ÌÂ ”›«—‘« "
         BeginProperty Font 
            Name            =   "Nazanin"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   2
         Left            =   240
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   840
         Width           =   2775
      End
      Begin VB.OptionButton optbalance 
         Alignment       =   1  'Right Justify
         Caption         =   "”›«—‘«  «‰Ã«„ ‘œÂ"
         BeginProperty Font 
            Name            =   "Nazanin"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   240
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   480
         Width           =   2775
      End
      Begin VB.OptionButton optbalance 
         Alignment       =   1  'Right Justify
         Caption         =   "”›«—‘«  «‰Ã«„ ‰‘œÂ"
         BeginProperty Font 
            Name            =   "Nazanin"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   240
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   120
         Width           =   2775
      End
   End
   Begin VB.TextBox TxtMembershipId 
      Alignment       =   2  'Center
      BackColor       =   &H0000FFFF&
      BeginProperty Font 
         Name            =   "B Nazanin"
         Size            =   15.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3360
      TabIndex        =   0
      Top             =   240
      Width           =   975
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   840
      Top             =   240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      PrintFileLinesPerPage=   60
   End
   Begin FLWCtrls.FWProgressBar FWProgressBar1 
      Height          =   375
      Left            =   480
      Top             =   6840
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   661
      Max             =   1000
      BorderStyle     =   10
   End
   Begin VSFlex7LCtl.VSFlexGrid vsFactors 
      Height          =   4725
      Left            =   0
      TabIndex        =   11
      Top             =   2040
      Width           =   12315
      _cx             =   21722
      _cy             =   8334
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Nazanin"
         Size            =   11.25
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
      SelectionMode   =   3
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   500
      RowHeightMax    =   500
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmFindOrder.frx":A4C2
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
      AllowUserFreezing=   1
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
   Begin RESIZEKITLibCtl.ResizeKit ResizeKit1 
      Height          =   480
      Left            =   9360
      OleObjectBlob   =   "frmFindOrder.frx":A690
      TabIndex        =   12
      Top             =   -480
      Width           =   480
   End
   Begin FLWCtrls.FWCoolButton fwBtnCustFind 
      Height          =   930
      Left            =   3360
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   960
      Width           =   2370
      _ExtentX        =   4180
      _ExtentY        =   1640
      BackColor       =   16777088
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Nazanin"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmFindOrder.frx":A716
      PictureAlign    =   4
      Caption         =   "„‘ —Ì"
      MaskColor       =   -2147483633
   End
   Begin FarDate1.FarDate txtDate1 
      Height          =   345
      Left            =   120
      TabIndex        =   14
      Top             =   600
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   609
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Object.ToolTipText     =   "FarDate"
   End
   Begin FarDate1.FarDate txtDate2 
      Height          =   345
      Left            =   120
      TabIndex        =   15
      Top             =   960
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   609
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Object.ToolTipText     =   "FarDate"
   End
   Begin VSFlex7LCtl.VSFlexGrid vsFactorDetail 
      Height          =   1935
      Left            =   7080
      TabIndex        =   22
      Top             =   6840
      Width           =   5280
      _cx             =   9313
      _cy             =   3413
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
      BackColorFixed  =   16777088
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
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   5
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmFindOrder.frx":AA30
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
   Begin FLWCtrls.FWNumericTextBox txtInterval 
      Height          =   480
      Left            =   720
      TabIndex        =   28
      Top             =   120
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   847
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
   Begin VB.Label lblSecondTitle 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "À«‰ÌÂ"
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
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   30
      Top             =   165
      Width           =   495
   End
   Begin VB.Label lblIntervalTitle 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "“„«‰ »—Ê“ —”«‰Ì :"
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
      TabIndex        =   29
      Top             =   165
      Width           =   1695
   End
   Begin VB.Shape ShapeBalance 
      BorderColor     =   &H000000FF&
      FillColor       =   &H000000FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   120
      Top             =   7560
      Width           =   615
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "„—ÃÊ⁄Ì"
      BeginProperty Font 
         Name            =   "Nazanin"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   840
      TabIndex        =   27
      Top             =   7560
      Width           =   615
   End
   Begin VB.Label LblFindFactor 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
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
      Height          =   495
      Left            =   9000
      RightToLeft     =   -1  'True
      TabIndex        =   21
      Top             =   120
      Width           =   3255
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "œ— Õ«·  «‰ Œ«» ›ﬁÿ ”Â —ﬁ„ ¬Œ— ' ”Ì” „ ﬂ·ÌÂ «ﬁ·«„ ›«ﬂ Ê—Â« Ì «Ì‰ ﬂ«—»— —« œ— ’Ê—  ÊÃÊœ œ” —”Ì  ‰„«Ì‘ „Ì œÂœ"
      BeginProperty Font 
         Name            =   "Nazanin"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   735
      Left            =   360
      RightToLeft     =   -1  'True
      TabIndex        =   20
      Top             =   7200
      Width           =   6495
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   " «  «—ÌŒ  ÕÊÌ· :"
      BeginProperty Font 
         Name            =   "B Homa"
         Size            =   11.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   2040
      RightToLeft     =   -1  'True
      TabIndex        =   19
      Top             =   960
      Width           =   1185
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "«“  «—ÌŒ  ÕÊÌ· :"
      BeginProperty Font 
         Name            =   "B Homa"
         Size            =   11.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   2040
      RightToLeft     =   -1  'True
      TabIndex        =   18
      Top             =   600
      Width           =   1305
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Ê÷⁄Ì  ”›«—‘"
      BeginProperty Font 
         Name            =   "Nazanin"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   375
      Left            =   5640
      RightToLeft     =   -1  'True
      TabIndex        =   17
      Top             =   120
      Width           =   3135
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "«‘ —«ò :"
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
      Left            =   4560
      RightToLeft     =   -1  'True
      TabIndex        =   16
      Top             =   360
      Width           =   855
   End
End
Attribute VB_Name = "frmFindOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim clsDate As New clsDate
Dim i As Long
Dim FactorType As EnumFactorType
Dim Parameter() As Parameter
Dim SearchOrderType As Integer
Dim ClsPrint As New Printing
Dim IsPrinting As Boolean

Private Sub CancelButton_Click()
    mvarcode = 0
    Unload Me
End Sub

Private Sub cmdprint_Click()
    On Error GoTo ErrHandler
    With vsFactors
        If vsFactors.Row < 1 Then Exit Sub
        mvarStatus = Order
        If chk_Print.Value = 0 Then
            IsPrinting = ClsPrint.Printing(vsFactors.ValueMatrix(.Row, 2), clsArya.StationNo, EnumAddEditMode.AddMode, EnumActionLog.Printing)
        Else
            IsPrinting = ClsPrint.Printing(vsFactors.ValueMatrix(.Row, 2), clsArya.StationNo, EnumAddEditMode.ViewMode, EnumActionLog.Reprint)
        End If
    End With
    If IsPrinting = True Then
        ShowDisMessage "ç«Å «‰Ã«„ ‘œ", 1500
        FillvsFactorDetail
    End If
Exit Sub
ErrHandler:
    ShowDisMessage err.Description, 1500
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call PresetScreenSaver
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call PresetScreenSaver
End Sub

Private Sub Form_Load()

    CenterCenterinSecondScreen Me
    
    mvarcode = 0
    
    FWProgressBar1.Visible = False
     
    With vsFactors
        .Cols = 21
        .TextMatrix(0, 0) = "—œÌ›"
        .TextMatrix(0, 1) = "”—Ì«·"
        .TextMatrix(0, 2) = "‘„«—Â ”›«—‘"
        .TextMatrix(0, 3) = " «—ÌŒ ”›«—‘"
        .TextMatrix(0, 4) = " «—ÌŒ  ÕÊÌ·"
        .TextMatrix(0, 5) = "“„«‰  ÕÊÌ·"
        .TextMatrix(0, 6) = "—Ê“  ÕÊÌ·"
        .TextMatrix(0, 7) = " «—ÌŒ  ÕÊÌ· ‘œÂ"
        .TextMatrix(0, 8) = "“„«‰  ÕÊÌ· ‘œÂ"
        .TextMatrix(0, 9) = "„»·€"
        .TextMatrix(0, 10) = "‰«„ ﬂ«—»—"
        .TextMatrix(0, 11) = " ÕÊÌ·Ì"
        .TextMatrix(0, 12) = "„‘ —Ì"
        .TextMatrix(0, 13) = "ﬂœ „‘ —Ì"
        .TextMatrix(0, 14) = "„Õ· ﬂ«—"
        .TextMatrix(0, 15) = "«‘ —«ﬂ"
        .TextMatrix(0, 16) = "“„«‰ ”›«—‘"
        .TextMatrix(0, 17) = "“„«‰ ”Å—Ì"
        .TextMatrix(0, 18) = "¬œ—” „Êﬁ "
        .TextMatrix(0, 19) = " ⁄œ«œ ‰›—« "
        .TextMatrix(0, 20) = "„—ÃÊ⁄Ì"
        .ColDataType(11) = flexDTBoolean
        .ColDataType(20) = flexDTBoolean
        For i = 0 To .Cols - 1
            .ColWidth(i) = Val(GetSetting(strMainKey, Me.Name, "Col" & i))
            If .ColWidth(i) = 0 Then
                .ColWidth(i) = .Width / 10     'Row
            End If
        Next i
    End With
    With vsFactorDetail
         .Cols = 5
         .TextMatrix(0, 0) = "—œÌ›"
         .TextMatrix(0, 1) = " ⁄œ«œ"
         .TextMatrix(0, 2) = "‰«„ ﬂ«·«"
         .TextMatrix(0, 3) = "›Ì"
         .TextMatrix(0, 4) = "Ã„⁄"
          For i = 0 To .Cols - 1
              .ColWidth(i) = Val(GetSetting(strMainKey, Me.Name & "vsFactorDetail", "Col" & i))
             If .ColWidth(i) = 0 Then
                 .ColWidth(i) = 1000       'Row
             End If
          Next i
         .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter
         .ColAlignment(2) = flexAlignRightCenter
    End With
    
    optAllDigits_Click
    
    If clsStation.SearchOrderType = 0 Then
        optbalance(0).Value = True
       '' optbalance_Click 0
     ElseIf clsStation.SearchOrderType = 1 Then
       optbalance(1).Value = True
       ''optbalance_Click 1
     Else
        optbalance(2).Value = True
       '' optbalance_Click 2
     End If
    
   ' txtDate1.Text = AccountYear & "/01/01"
    txtDate1.Text = "13" & Right(clsDate.shamsi(Date), 8)
    txtDate2.Text = "13" & Right(clsDate.shamsi(Date), 8)
    
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

    If Val(GetSetting(strMainKey, Me.Name, "TimerInterval")) > 0 Then
        Timer1.Interval = Val(GetSetting(strMainKey, Me.Name, "TimerInterval"))
'    Else
'         Timer1.Interval = 5000
    End If
    txtInterval.Value = CStr(Timer1.Interval / 1000)
    If txtInterval.Value = 0 Then Timer1.Enabled = False Else txtInterval.Enabled = True
    
    formloadFlag = True

End Sub

Private Sub Form_Unload(Cancel As Integer)
    If PosConnection.State = adStateOpen Then PosConnection.Close: Set PosConnection = Nothing

    SaveSetting strMainKey, Me.Name, "TimerInterval", CStr(Val(txtInterval.Value) * 1000)

    SaveSetting strMainKey, Me.Name, "Left", Me.Left
    SaveSetting strMainKey, Me.Name, "Top", Me.Top


End Sub

Private Sub fwBtnCustFind_Click()
    FindCust
End Sub
Private Sub Timer1_Timer()
    Timer1.Interval = txtInterval.Value * 1000
    FillvsFactors
End Sub

Public Sub FindCust()
    frmFindCust.Show vbModal
    
    If mvarcode <> 0 Then
        fwBtnCustFind.Tag = mvarcode
        fwBtnCustFind.Caption = mvarName
        mvarcode = 0
    Else
        fwBtnCustFind.Tag = 0
        fwBtnCustFind.Caption = "„‘ —Ì"
    End If
   
End Sub

Private Sub OKButton_Click()
    If vsFactors.Row > 0 Then
        mvarcode = vsFactors.TextMatrix(vsFactors.Row, 2)
    Else
        mvarcode = 0
    End If
    Unload Me

End Sub
Sub ClearDataFlexGrid()

    With vsFactors
        .Rows = 1
    End With
    
End Sub

Private Sub optAllDigits_Click()
    FillvsFactors
    vsFactors.Row = 0
    If vsFactors.Rows > 1 Then
       vsFactors.ShowCell 1, 0
       vsFactors.Sort = flexSortGenericDescending
    End If
End Sub

Private Sub optbalance_Click(index As Integer)
    If optbalance(0).Value = True Then
        SearchOrderType = 0
    ElseIf optbalance(1).Value = True Then
        SearchOrderType = 1
    Else
        SearchOrderType = 2
    End If
    ClearDataFlexGrid
    vsFactors.Row = 0
'    If Val(txtMembershipId.Text) = 0 Then
        FillvsFactors
'    Else
'        Define_OrderFactors
'    End If
End Sub

Private Sub RepBotton_Click()
    On Error GoTo Err_Handler
''If ClsFormAccess.DailyReport = True Then        Dim ArrayUbound  As Integer
    
    ReDim Parameter(4) As Parameter
    Parameter(0) = GenerateInputParameter("@FromDate", adVarWChar, 8, Right(txtDate1.Text, 8))
    Parameter(1) = GenerateInputParameter("@ToDate", adVarWChar, 8, Right(txtDate2.Text, 8))
    Parameter(2) = GenerateInputParameter("@AccountYear", adSmallInt, 2, AccountYear)
    Parameter(3) = GenerateInputParameter("@Balance", adSmallInt, 2, SearchOrderType)
    Parameter(4) = GenerateInputParameter("@CustCode", adBigInt, 8, Val(fwBtnCustFind.Tag))
    
    
'    Dim CustomerCode As Long
'    If vsFactors.Row >= 1 Then
'        ShowMessage "¬Ì« „ÌŒÊ«ÂÌœ ê“«—‘ ›ﬁÿ »—«Ì „‘ —Ì «‰ Œ«» ‘œÂ ‰„«Ì‘ œ«œÂ ‘Êœø", True, True, "»·Ì", "ŒÌ—"
'        If mvarMsgIdx = vbNo Then
'            CustomerCode = -1
'        Else
'            CustomerCode = IIf(Val(vsFactors.TextMatrix(vsFactors.Row, 12)) = 0, -1, CLng(Val(vsFactors.TextMatrix(vsFactors.Row, 12))))
'        End If
'    Else
'        CustomerCode = -1
'    End If
'
'    Parameter(5) = GenerateInputParameter("@Customer", adInteger, 4, CustomerCode)
'
'    Dim OrderType As Integer
'    OrderType = 1
'    'ShowInputForm True, True, False, "’⁄ÊœÌ", "‰“Ê·Ì", "", "„— » ”«“Ì ‘„«—Â ”›«—‘ »Â ’Ê— ", True, True, False
'
''    If mvarInput = "" Then Exit Sub
''
''    If mvarInput = "0" Then
''         OrderType = 1
''    ElseIf mvarInput = "1" Then
''         OrderType = -1
''    Else
''        OrderType = 1
''    End If
'
'    Parameter(6) = GenerateInputParameter("@OrderType", adTinyInt, 1, OrderType)
    
    CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\RepOrderByDate.rpt"
     
    Dim fileSystem As New FileSystemObject
    Dim IsFileExist As Boolean
     
    IsFileExist = fileSystem.FileExists(CrystalReport1.ReportFileName)
     
    If IsFileExist = False Then
        frmDisMsg.lblMessage = "›«Ì· " & CrystalReport1.ReportFileName & "ÅÌœ« ‰‘œ  "
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
     
Exit Sub
        
Err_Handler:
    ShowErrorMessage
    LogSaveNew "frmFindOrder => ", err.Description, err.Number, err.Source, "RepButton_Click"
End Sub


Private Sub ResizeKit1_ExitResize(ByVal XScale As Double, ByVal YScale As Double)
    If formloadFlag = True Then
        SaveSetting strMainKey, Me.Name, "Height", Me.Height
        SaveSetting strMainKey, Me.Name, "Width", Me.Width
    End If

End Sub

Private Sub txtDate1_Change()
    If Len(txtDate1.Text) = 10 Then
        FillvsFactors
    End If
End Sub

Private Sub txtDate2_Change()
    If Len(txtDate2.Text) = 10 Then
        FillvsFactors
    End If
End Sub

Private Sub txtInterval_Changed()
    If txtInterval.Value = 0 Then Timer1.Enabled = False Else txtInterval.Enabled = True
End Sub

Private Sub txtMembershipId_Change()
    Dim Rst As New ADODB.Recordset
''    fwBtnCustFind.Tag = TxtMembershipId.Text
    vsFactors.Rows = 1
    ReDim Parameter(1) As Parameter
    Parameter(0) = GenerateInputParameter("@Membershipid", adBigInt, 8, Val(TxtMembershipId.Text))
    Parameter(1) = GenerateInputParameter("@Branch", adInteger, 4, CurrentBranch)
    Set Rst = RunParametricStoredProcedure2Rec("Get_vw_Customers_ByMembership", Parameter)
    
    If Rst.EOF = False And Rst.BOF = False Then
        If fwBtnCustFind.Tag <> Rst!Code Then
            fwBtnCustFind.Tag = Rst!Code
             fwBtnCustFind.Caption = Rst!FullName
        End If
    Else
        fwBtnCustFind.Tag = 0
        fwBtnCustFind.Caption = "„‘ —Ì"
    End If
'    If Val(fwBtnCustFind.Tag) > 0 Or Val(txtMembershipId.Text) > 0 Then
'        Define_OrderFactors
'    Else
        FillvsFactors
'    End If
End Sub

Private Sub txtNo_Change()
    i = -1
    If vsFactors.Rows >= 1000 Then
        If Len(txtNo.Text) = 3 Then
           i = vsFactors.FindRow(txtNo.Text, 1, 2, True, True)
        End If
    Else
        i = vsFactors.FindRow(txtNo.Text, 1, 2, True, True)
    End If
    If i > 0 Then
        vsFactors.Row = i
        vsFactors.ShowCell i, 0
        LblFindFactor.Caption = ""
    Else
        vsFactors.Row = 0
        vsFactors.ShowCell 0, 0
        If Val(txtNo.Text) > 0 Then
           LblFindFactor.Caption = " ”›«—‘ " & Val(txtNo.Text) & "  œ— ”Ì” „  ⁄—Ì› ‰‘œÂ"
         Else
            LblFindFactor.Caption = "‘„«—Â ”›«—‘ —« Ê«—œ ﬂ‰Ìœ  "
         End If
    
    End If
    
End Sub

Private Sub txtNo_GotFocus()

    vsFactors.Row = 0
    vsFactors.Select vsFactors.Row, 2
'    vsFactors.Sort = flexSortGenericAscending
    vsFactors.Sort = flexSortGenericDescending
    LblFindFactor.Caption = "‘„«—Â ”›«—‘ —« Ê«—œ ﬂ‰Ìœ  "
    
End Sub

Private Sub FillvsFactors()

    On Error GoTo ErrHandler
    FWProgressBar1.Visible = True
    FWProgressBar1.Value = 0
    
    Dim Rst As New ADODB.Recordset
    
    ReDim Parameter(4) As Parameter
    Parameter(0) = GenerateInputParameter("@FromDate", adVarWChar, 8, Right(txtDate1.Text, 8))
    Parameter(1) = GenerateInputParameter("@ToDate", adVarWChar, 8, Right(txtDate2.Text, 8))
    Parameter(2) = GenerateInputParameter("@AccountYear", adSmallInt, 2, AccountYear)
    Parameter(3) = GenerateInputParameter("@Balance", adSmallInt, 2, SearchOrderType)
    Parameter(4) = GenerateInputParameter("@CustCode", adInteger, 4, Val(fwBtnCustFind.Tag))
    Set Rst = RunParametricStoredProcedure2Rec("Get_All_OrderFactors", Parameter)

    With vsFactors
        .Rows = 1
    
        i = 0
        While Rst.EOF <> True
            .Rows = .Rows + 1
            i = .Rows - 1
            .TextMatrix(i, 0) = i
            .TextMatrix(i, 1) = Rst!intSerialNo
            .TextMatrix(i, 2) = Rst!No
            .TextMatrix(i, 3) = Rst!RegDate
            .TextMatrix(i, 4) = IIf(IsNull(Rst!OldDeliveredDate), "", Rst!OldDeliveredDate)
            .TextMatrix(i, 5) = IIf(IsNull(Rst!OldDeliveredTime), "", Rst!OldDeliveredTime)
            .TextMatrix(i, 6) = IIf(IsNull(Rst!oldDeliveredDayName), "", Rst!oldDeliveredDayName)
            .TextMatrix(i, 7) = IIf(IsNull(Rst!DeliveredDate), "", Rst!DeliveredDate)
            .TextMatrix(i, 8) = IIf(IsNull(Rst!DeliveredTime), "", Rst!DeliveredTime)
            .TextMatrix(i, 9) = Rst!sumPrice
            .TextMatrix(i, 10) = Rst!nvcFirstName & " " & Rst!nvcSurname
            .TextMatrix(i, 11) = Rst!Delivered
            .TextMatrix(i, 12) = IIf(IsNull(Rst!CustomerName), "", Rst!CustomerName)
            .TextMatrix(i, 13) = IIf(IsNull(Rst!Customer), -1, Rst!Customer)
            .TextMatrix(i, 14) = Rst!WorkName
            .TextMatrix(i, 15) = IIf(Rst!Customer = -1, " ", Rst!MembershipId)
            .TextMatrix(i, 16) = IIf(IsNull(Rst!time), "", Rst!time)
            .TextMatrix(i, 17) = IIf(IsNull(Rst!RemainMinute), "", Rst!RemainMinute)
            .TextMatrix(i, 18) = IIf(IsNull(Rst!TempAddress), "", Rst!TempAddress)
            .TextMatrix(i, 19) = IIf(IsNull(Rst!GuestNo), "", Rst!GuestNo)
            If Rst!Recursive = 1 Then
                .Cell(flexcpBackColor, i, 0, i, .Cols - 1) = vbRed
                .TextMatrix(i, 20) = "1"
            End If
            Rst.MoveNext
             vsFactors.ColHidden(9) = True
            FWProgressBar1.Value = FWProgressBar1.Value + 1
            If FWProgressBar1.Value = 1000 Then
               FWProgressBar1.Value = 0
            End If
        Wend

    .ColHidden(13) = True
    .Cell(flexcpAlignment, 0, 0, 0, 13) = flexAlignCenterCenter
    .ColAlignment(-1) = flexAlignCenterCenter
    .ColAlignment(12) = flexAlignRightCenter
'    .AutoSizeMode = flexAutoSizeColWidth
'    .AutoSize 0, .Cols - 1
    FWProgressBar1.Value = 0
    FWProgressBar1.Visible = False
    End With
    
    Set Rst = Nothing

Exit Sub
ErrHandler:
    ShowDisMessage err.Description, 1500

End Sub


Private Sub vsFactorDetail_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    For i = 0 To vsFactorDetail.Cols - 1
        SaveSetting strMainKey, Me.Name & "vsFactorDetail", "Col" & i, vsFactorDetail.ColWidth(i)
    Next

End Sub

Public Sub FillvsFactorDetail() ' fills the detail of the current factor
    On Error GoTo Err_Handler
    
    Dim i As Integer
    Dim intselFactor As Double
    
    Dim Rst As New ADODB.Recordset
    
    If Rst.State = 1 Then Rst.Close
    
    With vsFactorDetail
        .Rows = 1
        If vsFactors.Row < 1 Then Exit Sub
        intselFactor = Val(vsFactors.TextMatrix(vsFactors.Row, 1))
        
        ReDim Parameter(1) As Parameter
        Parameter(0) = GenerateInputParameter("@intSelFactor", adInteger, 4, intselFactor)
        Parameter(1) = GenerateInputParameter("@Branch", adInteger, 4, CurrentBranch)
        Set Rst = RunParametricStoredProcedure2Rec("GetvwFactorDetailsInfo", Parameter)
        
        .Rows = 1
        
        If Not (Rst.EOF = True And Rst.BOF = True) Then
        
            i = 1
            While Rst.EOF = False ' fill the grid
                .Rows = .Rows + 1
                .TextMatrix(i, 0) = Rst.Fields("intRow").Value
                .TextMatrix(i, 1) = Rst.Fields("Amount").Value
                .TextMatrix(i, 2) = Rst.Fields("Name").Value
                .TextMatrix(i, 3) = Rst.Fields("FeeUnit").Value
                .TextMatrix(i, 4) = Rst.Fields("FeeUnit").Value * Rst.Fields("Amount").Value
                Rst.MoveNext
                i = i + 1
            Wend
            
        End If
       ''.AutoSizeMode = flexAutoSizeColWidth  ' set the collumns' width
       '' .AutoSize 0, .Cols - 1
    End With
    
    If Rst.State <> 0 Then Rst.Close
     
    If vsFactors.TextMatrix(vsFactors.Row, 20) = "" Then cmdPrint.Enabled = True Else cmdPrint.Enabled = False
    
    Dim CountPrinting, CountRePrint, CountInvoicePrint As Integer
    
    ReDim Parameter(2) As Parameter
    Parameter(0) = GenerateInputParameter("@No", adBigInt, 8, Val(vsFactors.ValueMatrix(vsFactors.Row, 2)))
    Parameter(1) = GenerateInputParameter("@Status", adInteger, 4, mvarStatus)
    Parameter(2) = GenerateInputParameter("@AccountYear", adSmallInt, 2, AccountYear)
    Set Rst = RunParametricStoredProcedure2Rec("Get_CountPrint_tAction", Parameter)
    
    CountPrinting = Rst!CountPrinting
    CountRePrint = Rst!CountRePrint
    CountInvoicePrint = Rst!CountInvoicePrint
    
    If Rst.State <> 0 Then Rst.Close
    Set Rst = Nothing
    If CountPrinting > 0 Then chk_Print.Value = 1 Else chk_Print.Value = 0
    If CountRePrint > 0 Then chk_Reprint.Value = 1 Else chk_Reprint.Value = 0
    
    Exit Sub
    
Err_Handler:
    ShowErrorMessage
    LogSaveNew "frmFindOrder => ", err.Description, err.Number, err.Source, "FillVsFactorDetail"
End Sub

Private Sub vsFactors_AfterSort(ByVal Col As Long, Order As Integer)
    For i = 1 To vsFactors.Rows - 1
        vsFactors.TextMatrix(i, 0) = i
    Next
    
End Sub

Private Sub vsFactors_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    
    For i = 0 To vsFactors.Cols - 1
        SaveSetting strMainKey, Me.Name, "Col" & Col, vsFactors.ColWidth(Col)
    Next

End Sub

Private Sub vsFactors_DblClick()
    If vsFactors.Row > 0 Then
        OKButton_Click
    End If
End Sub

Private Sub Define_OrderFactors()

    On Error GoTo ErrHandler
    Dim Rst As New ADODB.Recordset
    If fwBtnCustFind.Tag > 0 Then
        ReDim Parameter(2) As Parameter
        Parameter(0) = GenerateInputParameter("@CustCode", adBigInt, 8, Val(fwBtnCustFind.Tag))
        Parameter(1) = GenerateInputParameter("@AccountYear", adSmallInt, 2, AccountYear)
        Parameter(2) = GenerateInputParameter("@Balance", adSmallInt, 2, SearchOrderType)
        Set Rst = RunParametricStoredProcedure2Rec("Get_Define_OrderFactors_byCustCode", Parameter)
'    Else
'        ReDim Parameter(2) As Parameter
'        Parameter(0) = GenerateInputParameter("@No", adBigInt, 8, Val(txtNo.Text))
'        Parameter(1) = GenerateInputParameter("@AccountYear", adSmallInt, 2, AccountYear)
'        Parameter(2) = GenerateInputParameter("@Balance", adSmallInt, 2, SearchOrderType)
'        Set Rst = RunParametricStoredProcedure2Rec("Get_Define_OrderFactors", Parameter)
    End If
    With vsFactors
        .Rows = 1
       
        i = 0
        While Rst.EOF <> True
            .Rows = .Rows + 1
             i = .Rows - 1
            .TextMatrix(i, 0) = i
            .TextMatrix(i, 1) = Rst!intSerialNo
            .TextMatrix(i, 2) = Rst!No
            .TextMatrix(i, 3) = Rst!RegDate
            .TextMatrix(i, 4) = IIf(IsNull(Rst!OldDeliveredDate), "", Rst!OldDeliveredDate)
            .TextMatrix(i, 5) = IIf(IsNull(Rst!OldDeliveredTime), "", Rst!OldDeliveredTime)
            .TextMatrix(i, 6) = IIf(IsNull(Rst!oldDeliveredDayName), "", Rst!oldDeliveredDayName)
            .TextMatrix(i, 7) = IIf(IsNull(Rst!DeliveredDate), "", Rst!DeliveredDate)
            .TextMatrix(i, 8) = IIf(IsNull(Rst!DeliveredTime), "", Rst!DeliveredTime)
            .TextMatrix(i, 9) = Rst!sumPrice
            .TextMatrix(i, 10) = Rst!nvcFirstName & " " & Rst!nvcSurname
            .TextMatrix(i, 11) = Rst!Delivered
            .TextMatrix(i, 12) = IIf(Rst!Customer = -1, " ", Rst!CustomerName)
            .TextMatrix(i, 13) = IIf(IsNull(Rst!Customer), -1, Rst!Customer)
            .TextMatrix(i, 14) = Rst!WorkName
            .TextMatrix(i, 15) = IIf(Rst!Customer = -1, " ", Rst!MembershipId)
            .TextMatrix(i, 16) = IIf(IsNull(Rst!time), "", Rst!time)
            .TextMatrix(i, 17) = IIf(IsNull(Rst!RemainMinute), "", Rst!RemainMinute)
            .TextMatrix(i, 18) = IIf(IsNull(Rst!TempAddress), "", Rst!TempAddress)
            .TextMatrix(i, 19) = IIf(IsNull(Rst!GuestNo), "", Rst!GuestNo)
            Rst.MoveNext

        Wend
'        .AutoSizeMode = flexAutoSizeColWidth
       .ColAlignment(11) = flexAlignRightCenter

    End With
    
    Set Rst = Nothing
Exit Sub
ErrHandler:
    ShowDisMessage err.Description, 1500
End Sub


Private Sub vsFactors_RowColChange()
    
    FillvsFactorDetail

End Sub
