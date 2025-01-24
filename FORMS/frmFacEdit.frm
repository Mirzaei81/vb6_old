VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{7AEDC602-D94C-11D1-BB7A-00E0290EA3C9}#1.0#0"; "ResizeKit.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmFacEdit 
   ClientHeight    =   9030
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14970
   Icon            =   "frmFacEdit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9030
   ScaleWidth      =   14970
   Begin VB.ComboBox cmbBranch 
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
      Left            =   1440
      RightToLeft     =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   25
      Top             =   360
      Width           =   2475
   End
   Begin VB.CheckBox ChkLessEdited 
      Alignment       =   1  'Right Justify
      Caption         =   "ç«Å «’·«ÕÌ Â«Ì ﬂ„ ‘œÂ"
      BeginProperty Font 
         Name            =   "B Homa"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   435
      Left            =   12480
      TabIndex        =   24
      Top             =   5520
      Value           =   1  'Checked
      Width           =   2385
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   120
      TabIndex        =   13
      Top             =   5160
      Width           =   12255
      Begin VB.Label LblOldSum 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "B Mitra"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   345
         Left            =   4560
         TabIndex        =   29
         Top             =   480
         Width           =   1545
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "„»·€ ›«ò Ê—Â«ﬁ»· «“ «’·«Õ "
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
         Height          =   345
         Left            =   6000
         TabIndex        =   28
         Top             =   480
         Width           =   2385
      End
      Begin VB.Label lblCount 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "B Mitra"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   8550
         TabIndex        =   21
         Top             =   120
         Width           =   1425
      End
      Begin VB.Label lblSum 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "B Mitra"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   345
         Left            =   4560
         TabIndex        =   20
         Top             =   120
         Width           =   1545
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "„»·€ ›«ò Ê—Â«Ì «’·«Õ ‘œÂ"
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
         Height          =   345
         Left            =   6120
         TabIndex        =   19
         Top             =   120
         Width           =   2205
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   " ⁄œ«œ ›«ò Ê—Â«Ì «’·«Õ ‘œÂ"
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   9960
         TabIndex        =   18
         Top             =   120
         Width           =   2205
      End
      Begin VB.Label lblDailySum 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "B Mitra"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   120
         TabIndex        =   17
         Top             =   480
         Width           =   1485
      End
      Begin VB.Label lblDailyCount 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "B Mitra"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   120
         TabIndex        =   16
         Top             =   120
         Width           =   1365
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "„»·€ ›«ò Ê—Â«Ì «’·«Õ ‘œÂ «„—Ê“"
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1605
         TabIndex        =   15
         Top             =   480
         Width           =   2745
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   " ⁄œ«œ ›«ò Ê—Â«Ì «’·«Õ ‘œÂ «„—Ê“"
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
         Left            =   1605
         TabIndex        =   14
         Top             =   120
         Width           =   2745
      End
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   240
      Top             =   600
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
   Begin VB.CommandButton cmdView 
      BackColor       =   &H008080FF&
      Caption         =   "‰„«Ì‘"
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
      Left            =   120
      TabIndex        =   11
      Top             =   240
      Width           =   1215
   End
   Begin VB.CheckBox ChkDailyView 
      Alignment       =   1  'Right Justify
      Caption         =   "›«ò Ê—Â«Ì «’·«ÕÌ «„—Ê“"
      BeginProperty Font 
         Name            =   "B Homa"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   435
      Left            =   12480
      TabIndex        =   8
      Top             =   5160
      Value           =   1  'Checked
      Width           =   2385
   End
   Begin VSFlex7LCtl.VSFlexGrid vsEditedFactorDetails 
      Height          =   4125
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   3975
      _cx             =   7011
      _cy             =   7276
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
      BackColorFixed  =   -2147483633
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
      AllowUserResizing=   0
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
      FormatString    =   $"frmFacEdit.frx":A4C2
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
   Begin VSFlex7LCtl.VSFlexGrid vsEditedFactors 
      Height          =   4125
      Left            =   4200
      TabIndex        =   2
      Top             =   960
      Width           =   10695
      _cx             =   18865
      _cy             =   7276
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
      SelectionMode   =   3
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
      FormatString    =   $"frmFacEdit.frx":A5A1
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
   Begin VSFlex7LCtl.VSFlexGrid vsPreviousFactorDetails 
      Height          =   2655
      Left            =   120
      TabIndex        =   1
      Top             =   6240
      Width           =   3975
      _cx             =   7011
      _cy             =   4683
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
      BackColorFixed  =   -2147483633
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
      AllowUserResizing=   0
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
      FormatString    =   $"frmFacEdit.frx":A681
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
   Begin VSFlex7LCtl.VSFlexGrid vsPreviousFactors 
      Height          =   2655
      Left            =   4080
      TabIndex        =   0
      Top             =   6240
      Width           =   10815
      _cx             =   19076
      _cy             =   4683
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
      BackColorFixed  =   16777152
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
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmFacEdit.frx":A760
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
   Begin MSMask.MaskEdBox txtDateFrom 
      Height          =   465
      Left            =   9240
      TabIndex        =   22
      Top             =   480
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   820
      _Version        =   393216
      MaxLength       =   8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "##/##/##"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox txtDateTo 
      Height          =   465
      Left            =   5640
      TabIndex        =   23
      Top             =   480
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   820
      _Version        =   393216
      MaxLength       =   8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "##/##/##"
      PromptChar      =   " "
   End
   Begin RESIZEKITLibCtl.ResizeKit ResizeKit1 
      Height          =   480
      Left            =   0
      OleObjectBlob   =   "frmFacEdit.frx":A840
      TabIndex        =   27
      Top             =   0
      Width           =   480
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   " : ‘⁄»Â"
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
      Left            =   3960
      TabIndex        =   26
      Top             =   480
      Width           =   855
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "›«ò Ê—Â«Ì «’·«ÕÌ"
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
      Height          =   615
      Left            =   6480
      TabIndex        =   12
      Top             =   -120
      Width           =   2655
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   " «  «—ÌŒ"
      BeginProperty Font 
         Name            =   "B Homa"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   7320
      TabIndex        =   10
      Top             =   480
      Width           =   945
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "«“  «—ÌŒ"
      BeginProperty Font 
         Name            =   "B Homa"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   11040
      TabIndex        =   9
      Top             =   480
      Width           =   825
   End
   Begin VB.Label lblPreviousFactorDetails 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   840
      TabIndex        =   7
      Top             =   6360
      Width           =   3255
   End
   Begin VB.Label lblPreviousFactors 
      Alignment       =   1  'Right Justify
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
      Height          =   345
      Left            =   12360
      TabIndex        =   6
      Top             =   5880
      Width           =   2535
   End
   Begin VB.Label lblEditedFactors 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "›«ò Ê— Â«Ì «’·«Õ ‘œÂ"
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
      Height          =   345
      Left            =   12720
      TabIndex        =   5
      Top             =   600
      Width           =   2055
   End
   Begin VB.Label lblEditedFactorDetails 
      Alignment       =   1  'Right Justify
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
      Height          =   345
      Left            =   600
      TabIndex        =   4
      Top             =   1080
      Width           =   3495
   End
End
Attribute VB_Name = "frmFacEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim clsDate As New clsDate
Dim i As Integer
Dim Parameter() As Parameter

Private Sub FillvsEditedFactors()

    Dim Rst As New ADODB.Recordset
    
    lblDailySum.Caption = 0
    lblDailyCount.Caption = 0
    lblSum.Caption = 0
    LblCount.Caption = 0
    
    With vsEditedFactors
        
        .Rows = 1
        vsEditedFactorDetails.Rows = 1
        ReDim Parameter(3) As Parameter
        Parameter(0) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
        Parameter(1) = GenerateInputParameter("@DateAfter", adVarWChar, 20, CStr(IIf(Trim(txtDateFrom.ClipText) = "", "", Trim(txtDateFrom.Text))))
        Parameter(2) = GenerateInputParameter("@DateBefore", adVarWChar, 20, CStr(IIf(Trim(txtDateTo.ClipText) = "", "", Trim(txtDateTo.Text))))
        Parameter(3) = GenerateInputParameter("@Branch", adInteger, 4, cmbBranch.ItemData(cmbBranch.ListIndex))
        Set Rst = RunParametricStoredProcedure2Rec("Get_EditedFactors", Parameter)
        
        If Rst.EOF = True And Rst.BOF = True Then
            .Rows = 1
        Else
            Dim VarToday As String
            VarToday = mvarDate
            While Rst.EOF = False
                If ChkDailyView.Value = False Or Rst.Fields("Date").Value = VarToday Then
                    .Rows = .Rows + 1
                    i = .Rows - 1
                    .TextMatrix(i, 0) = i
                    .TextMatrix(i, 1) = Rst.Fields("intSerialNo").Value
                    .TextMatrix(i, 2) = Rst.Fields("No").Value
                    .TextMatrix(i, 3) = Rst.Fields("Status").Value
                    .TextMatrix(i, 4) = Rst.Fields("ShiftDescription").Value
                    .TextMatrix(i, 5) = Rst.Fields("Recursive").Value
                    .TextMatrix(i, 6) = Rst.Fields("nvcFirstName").Value & " " & Rst.Fields("nvcSurname").Value
                    .TextMatrix(i, 7) = Rst.Fields("SumPrice").Value
                    .TextMatrix(i, 8) = Rst.Fields("Time").Value
                    .TextMatrix(i, 9) = Rst.Fields("Date").Value
                    .TextMatrix(i, 10) = Rst.Fields("RegDate").Value
                    .TextMatrix(i, 11) = Rst.Fields("DiscountTotal").Value
                    .TextMatrix(i, 12) = IIf(IsNull(Rst.Fields("TableName").Value), "", Rst.Fields("TableName").Value)
                    .TextMatrix(i, 13) = Rst.Fields("StationId").Value
                    .TextMatrix(i, 14) = Rst.Fields("CarryFeeTotal").Value
                    .TextMatrix(i, 15) = Rst.Fields("ServiceTotal").Value
                    .TextMatrix(i, 16) = Rst.Fields("PackingTotal").Value
'                    .Row = i
                    
                    If Rst.Fields("Date").Value = VarToday Then
                        lblDailySum.Caption = lblDailySum.Caption + Rst.Fields("SumPrice").Value
                        lblDailyCount.Caption = lblDailyCount.Caption + 1
                    End If
                End If
                Rst.MoveNext
            Wend
            If .Rows > 1 Then

                lblSum.Caption = .Aggregate(flexSTSum, .FixedRows, 7, .Rows - 1, 7)
                LblCount.Caption = .Aggregate(flexSTCount, .FixedRows, 7, .Rows - 1, 7)
                .Row = 0
                .Row = 1
            End If
        End If
        .AutoSizeMode = flexAutoSizeColWidth
        .AutoSize 0, .Cols - 1
    End With
    ReDim Parameter(3) As Parameter
    Parameter(0) = GenerateInputParameter("@DateAfter", adVarWChar, 8, CStr(IIf(Trim(txtDateFrom.ClipText) = "", "", Trim(txtDateFrom.Text))))
    Parameter(1) = GenerateInputParameter("@DateBefore", adVarWChar, 8, CStr(IIf(Trim(txtDateTo.ClipText) = "", "", Trim(txtDateTo.Text))))
    Parameter(2) = GenerateInputParameter("@AccountYear", adSmallInt, 2, AccountYear)
    Parameter(3) = GenerateInputParameter("@Branch", adInteger, 4, cmbBranch.ItemData(cmbBranch.ListIndex))
    Set Rst = RunParametricStoredProcedure2Rec("Get_EditedPrice", Parameter)
    
    lblSum.RightToLeft = True
    LblOldSum.RightToLeft = True
    If Rst.EOF = True And Rst.BOF = True Then
        lblSum.Caption = 0
        LblOldSum.Caption = 0
    Else
        lblSum.Caption = Format(Rst!NewSumprice, "#,## ") & "  —Ì«·"
        LblOldSum.Caption = Format(Rst!OldSumPrice, "#,## ") & "  —Ì«·"
    End If
    Set Rst = Nothing

End Sub
Public Sub ExitForm()

    Unload Me
    
End Sub

Private Sub ChkDailyView_Click()
    FillvsEditedFactors
    FillvsPreviousFactorDetails
End Sub

Private Sub cmbBranch_Change()
    ChkDailyView_Click
    FillvsEditedFactors
    FillvsPreviousFactors

End Sub

Private Sub cmbBranch_Click()
    cmbBranch_Change
End Sub

Private Sub cmdView_Click()
    ChkDailyView.Value = False
    FillvsEditedFactors
    FillvsPreviousFactors
End Sub

Private Sub Form_Activate()

     VarActForm = Me.Name
End Sub
Private Sub FillBranch()
    Dim rctmp As New ADODB.Recordset
    cmbBranch.Clear
'    cmbBranch.AddItem "Â„Â ‘⁄»« "
'    cmbBranch.ItemData(cmbBranch.NewIndex) = 0
    Set rctmp = RunStoredProcedure2RecordSet("Get_All_Branches")
    Do While rctmp.EOF = False
        cmbBranch.AddItem rctmp!nvcBranchName
        cmbBranch.ItemData(cmbBranch.NewIndex) = rctmp!Branch
        rctmp.MoveNext
    Loop
    rctmp.Close
    If cmbBranch.ListCount > 0 Then cmbBranch.ListIndex = 0
    
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

    If ClsFormAccess.frmFacEdit = False Then
        Unload Me
        Exit Sub
    End If

    Dim Rst As New ADODB.Recordset
    Dim tmpString As String
    
    CenterTop Me
    
    VarActForm = Me.Name
    Dim s As String
    With vsEditedFactors
        .Cols = 17
        .ColHidden(1) = True
        .ColHidden(15) = False
        s = ""
        ReDim Parameter(1) As Parameter
        Parameter(0) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
        Parameter(1) = GenerateInputParameter("@ReturnType", adInteger, 4, 0)
        Set Rst = RunParametricStoredProcedure2Rec("Get_All_tStatusType", Parameter)
        s = .BuildComboList(Rst, "NvcDescription", "intStatusNo")
        .ColComboList(3) = s
        '.ColComboList(3) = "#1;›«ò Ê— Œ—Ìœ|#2;›«ò Ê— ›—Ê‘"
        .ColDataType(5) = flexDTBoolean  'edited 11/26
        
        .TextMatrix(0, 1) = ""
        .TextMatrix(0, 2) = "”—Ì«·"
        .TextMatrix(0, 3) = "‰Ê⁄ ›«ò Ê—"
        .TextMatrix(0, 4) = "‘Ì› "
        .TextMatrix(0, 5) = "„—ÃÊ⁄Ì"
        .TextMatrix(0, 6) = "‰«„ ò«—»—"
        .TextMatrix(0, 7) = "Ã„⁄"
        .TextMatrix(0, 8) = "”«⁄ "
        .TextMatrix(0, 9) = " «—ÌŒ ’œÊ—"
        .TextMatrix(0, 10) = " «—ÌŒ «’·«Õ"
        .TextMatrix(0, 11) = " Œ›Ì›"
        .TextMatrix(0, 12) = "„Ì“"
        .TextMatrix(0, 13) = "«Ì” ê«Â"
        .TextMatrix(0, 14) = "Â“Ì‰Â Õ„·"
        .TextMatrix(0, 15) = "”—ÊÌ”"
        .TextMatrix(0, 6) = "»” Â »‰œÌ"
        .FixedAlignment(-1) = flexAlignCenterCenter
        .ColAlignment(-1) = flexAlignCenterCenter
        .ColAlignment(6) = flexAlignRightCenter
    
    End With
    
    With vsEditedFactorDetails
        .Cols = 6
        
        If Rst.State = 1 Then Rst.Close
        
        ReDim Parameter(0) As Parameter
        Parameter(0) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
        Set Rst = RunParametricStoredProcedure2Rec("Get_Serveplace", Parameter)
        
        strTemp = .BuildComboList(Rst, "Description", "intServePlace")
        .ColComboList(5) = strTemp
        If Rst.State <> 0 Then Rst.Close
        
        .TextMatrix(0, 1) = "„ﬁœ«—"
        .TextMatrix(0, 2) = "‰«„ ò«·«"
        .TextMatrix(0, 3) = "›Ì"
        .TextMatrix(0, 4) = "ﬁÌ„ "
        .TextMatrix(0, 5) = "„Õ· ”—Ê"
        .ColAlignment(-1) = flexAlignCenterCenter
        .ColAlignment(2) = flexAlignRightCenter
        
  
    End With
    
    With vsPreviousFactors
        .Cols = 18
        .ColHidden(1) = True
        .ColComboList(3) = s
       ' .ColComboList(3) = "#1;›«ò Ê— Œ—Ìœ|#2;›«ò Ê— ›—Ê‘"
    
        .TextMatrix(0, 1) = ""
        .TextMatrix(0, 2) = "”—Ì«·"
        .TextMatrix(0, 3) = "‰Ê⁄ ›«ò Ê—"
        .TextMatrix(0, 4) = "‘Ì› "
        .TextMatrix(0, 5) = "„—ÃÊ⁄Ì"
        .TextMatrix(0, 6) = "‰«„ ò«—»—"
        .TextMatrix(0, 7) = "Ã„⁄"
        .TextMatrix(0, 8) = "”«⁄ "
        .TextMatrix(0, 9) = " «—ÌŒ ’œÊ—"
        .TextMatrix(0, 10) = " «—ÌŒ «’·«Õ"
        .TextMatrix(0, 11) = " Œ›Ì›"
        .TextMatrix(0, 12) = "„Ì“"
        .TextMatrix(0, 13) = "«Ì” ê«Â"
        .TextMatrix(0, 14) = "Â“Ì‰Â Õ„·"
        .TextMatrix(0, 16) = "”—ÊÌ”"
        .TextMatrix(0, 17) = "»” Â »‰œÌ"
        .ColAlignment(-1) = flexAlignCenterCenter
        .FixedAlignment(-1) = flexAlignCenterCenter
        .ColAlignment(6) = flexAlignRightCenter
    
    End With

    With vsPreviousFactorDetails
        .Cols = 6
        
        If Rst.State = 1 Then Rst.Close
        
        ReDim Parameter(0) As Parameter
        Parameter(0) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
        Set Rst = RunParametricStoredProcedure2Rec("Get_Serveplace", Parameter)
        
        strTemp = .BuildComboList(Rst, "Description", "intServePlace")
        .ColComboList(5) = strTemp
        If Rst.State <> 0 Then Rst.Close
        
        .TextMatrix(0, 1) = "„ﬁœ«—"
        .TextMatrix(0, 2) = "‰«„ ò«·«"
        .TextMatrix(0, 3) = "›Ì"
        .TextMatrix(0, 4) = "ﬁÌ„ "
        .TextMatrix(0, 5) = " ”—Ê"
        .ColAlignment(-1) = flexAlignCenterCenter
        .ColAlignment(2) = flexAlignRightCenter
   
    End With
    
    Set Rst = Nothing
    
    txtDateFrom.Text = mvarDate
    txtDateTo.Text = Mid(clsDate.shamsi(Date), 3)
    FillBranch

    ChkDailyView_Click
    mdifrm.Toolbar1.Buttons(15).Enabled = True  'printing
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

Private Sub Form_Unload(Cancel As Integer)
    VarActForm = ""
    If PosConnection.State = adStateOpen Then PosConnection.Close: Set PosConnection = Nothing

    SaveSetting strMainKey, Me.Name, "Left", Me.Left
    SaveSetting strMainKey, Me.Name, "Top", Me.Top


End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call PresetScreenSaver
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call PresetScreenSaver
End Sub



Private Sub ResizeKit1_ExitResize(ByVal XScale As Double, ByVal YScale As Double)

    If formloadFlag = True Then
        SaveSetting strMainKey, Me.Name, "Height", Me.Height
        SaveSetting strMainKey, Me.Name, "Width", Me.Width
    End If

End Sub


Private Sub vsEditedFactors_Click()
'    FillvsEditedFactorDetails
'    FillvsPreviousFactors
End Sub

Private Sub vsEditedFactors_RowColChange()
    FillvsEditedFactorDetails
    FillvsPreviousFactors

End Sub

Private Sub FillvsPreviousFactors()

    If vsEditedFactors.Row < 1 Or vsEditedFactors.Rows < 2 Then
        lblPreviousFactors.Caption = "›«ò Ê—Â«Ì ”«»ﬁ "
        vsPreviousFactors.Rows = 1
        vsPreviousFactorDetails.Rows = 1
        Exit Sub
    End If
    
    Dim Rst As New ADODB.Recordset
    ReDim Parameter(2) As Parameter
    
    With vsPreviousFactors
        
        .Rows = 1
        vsPreviousFactorDetails.Rows = 1
        vsPreviousFactors.Rows = 1
        Parameter(0) = GenerateInputParameter("@intSerialNo", adBigInt, 8, Val(vsEditedFactors.TextMatrix(vsEditedFactors.Row, 1)))
        Parameter(1) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
        Parameter(2) = GenerateInputParameter("@Branch", adInteger, 4, cmbBranch.ItemData(cmbBranch.ListIndex))
        Set Rst = RunParametricStoredProcedure2Rec("Get_PreviousFactors", Parameter)
        
        If Rst.EOF = True And Rst.BOF = True Then
        
        Else
            lblPreviousFactors.Caption = "›«ò Ê—Â«Ì ”«»ﬁ " & vsEditedFactors.TextMatrix(vsEditedFactors.Row, 2)
            i = 1
            While Rst.EOF = False
                .Rows = .Rows + 1
                i = .Rows - 1
                .TextMatrix(i, 0) = i
                .TextMatrix(i, 1) = Rst.Fields("intSerialNo").Value
                .TextMatrix(i, 2) = Rst.Fields("No").Value
                .TextMatrix(i, 3) = Rst.Fields("Status").Value
                .TextMatrix(i, 4) = Rst.Fields("ShiftDescription").Value
                .TextMatrix(i, 5) = Rst.Fields("Recursive").Value
                .TextMatrix(i, 6) = Rst.Fields("nvcFirstName").Value & " " & Rst.Fields("nvcSurname").Value
                .TextMatrix(i, 7) = Rst.Fields("SumPrice").Value
                .TextMatrix(i, 8) = Rst.Fields("Time").Value
                .TextMatrix(i, 9) = Rst.Fields("Date").Value
                .TextMatrix(i, 10) = Rst.Fields("RegDate").Value
                .TextMatrix(i, 11) = Rst.Fields("DiscountTotal").Value
                .TextMatrix(i, 12) = IIf(IsNull(Rst.Fields("TableName").Value), "", Rst.Fields("TableName").Value)
                .TextMatrix(i, 13) = Rst.Fields("StationId").Value
                .TextMatrix(i, 14) = Rst.Fields("CarryFeeTotal").Value
                .TextMatrix(i, 16) = Rst.Fields("ServiceTotal").Value
                .TextMatrix(i, 17) = Rst.Fields("PackingTotal").Value
                .TextMatrix(i, 15) = Rst.Fields("Code").Value
                Rst.MoveNext
            Wend
            .Row = 0
            .Row = 1
        End If
        .AutoSizeMode = flexAutoSizeColWidth
        .AutoSize 0, .Cols - 1
        If .Rows > 1 Then
            If .Aggregate(flexSTMax, .FixedRows, 7, .Rows - 1, 7) > vsEditedFactors.TextMatrix(vsEditedFactors.Row, 7) Then
                vsEditedFactors.Cell(flexcpBackColor, vsEditedFactors.Row, vsEditedFactors.FixedCols, vsEditedFactors.Row, vsEditedFactors.Cols - 1) = 8421631
            End If
        End If
        
    End With
    Set Rst = Nothing

End Sub


Private Sub FillvsEditedFactorDetails()

    If vsEditedFactors.Row < 1 Or vsEditedFactors.Rows < 2 Then
       vsPreviousFactorDetails.Rows = 1
        Exit Sub
    End If
    
    Dim Rst As New ADODB.Recordset
    ReDim Parameter(2) As Parameter
 
    With vsEditedFactorDetails
        
        .Rows = 1
     
        
        Parameter(0) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
        Parameter(1) = GenerateInputParameter("@intSerialNo", adBigInt, 8, Val(vsEditedFactors.TextMatrix(vsEditedFactors.Row, 1)))
        Parameter(2) = GenerateInputParameter("@Branch", adInteger, 4, CurrentBranch)
        
        Set Rst = RunParametricStoredProcedure2Rec("Get_Factor_Detail", Parameter)
        
        .Rows = 1
        
        If Not (Rst.EOF = True And Rst.BOF = True) Then
            lblEditedFactorDetails.Caption = "—Ì“ «ﬁ·«„ ›«ò Ê— " & vsEditedFactors.TextMatrix(vsEditedFactors.Row, 2)
            i = 1
            While Rst.EOF = False ' fill the grid
                .Rows = .Rows + 1
                .TextMatrix(i, 0) = Rst.Fields("intRow").Value
                .TextMatrix(i, 1) = Rst.Fields("Amount").Value
                .TextMatrix(i, 2) = Rst.Fields("Name").Value
                .TextMatrix(i, 3) = Rst.Fields("FeeUnit").Value
                .TextMatrix(i, 4) = Rst.Fields("FeeUnit").Value * Rst.Fields("Amount").Value
                .TextMatrix(i, 5) = Rst.Fields("ServePlace").Value
                
                Rst.MoveNext
                i = i + 1
            Wend
            
        End If
        .AutoSizeMode = flexAutoSizeColWidth
        .AutoSize 0, .Cols - 1
        
    End With
    Set Rst = Nothing
End Sub


Private Sub FillvsPreviousFactorDetails()

    If vsPreviousFactors.Row < 1 Or vsPreviousFactors.Rows < 2 Then
        Exit Sub
    End If
    
    Dim Rst As New ADODB.Recordset
    ReDim Parameter(2) As Parameter
    
    With vsPreviousFactorDetails
        
        .Rows = 1
        Parameter(0) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
        Parameter(1) = GenerateInputParameter("@Code", adInteger, 4, Val(vsPreviousFactors.TextMatrix(vsPreviousFactors.Row, 15)))
        Parameter(2) = GenerateInputParameter("@Branch", adInteger, 4, cmbBranch.ItemData(cmbBranch.ListIndex))
        Set Rst = RunParametricStoredProcedure2Rec("Get_Previous_Factor_Detail", Parameter)
                
        If Not (Rst.EOF = True And Rst.BOF = True) Then
            lblPreviousFactorDetails.Caption = "—Ì“ «ﬁ·«„ ›«ò Ê—Â«Ì ”«»ﬁ " & vsPreviousFactors.TextMatrix(vsPreviousFactors.Row, 2)
            i = 1
            While Rst.EOF = False ' fill the grid
                .Rows = .Rows + 1
                .TextMatrix(i, 0) = Rst.Fields("intRow").Value
                .TextMatrix(i, 1) = Rst.Fields("Amount").Value
                .TextMatrix(i, 2) = Rst.Fields("Name").Value
                .TextMatrix(i, 3) = Rst.Fields("FeeUnit").Value
                .TextMatrix(i, 4) = Rst.Fields("FeeUnit").Value * Rst.Fields("Amount").Value
                .TextMatrix(i, 5) = Rst.Fields("ServePlace").Value
                
                Rst.MoveNext
                i = i + 1
            Wend
        Else
            .Rows = 1
        End If
        .AutoSizeMode = flexAutoSizeColWidth
        .AutoSize 0, .Cols - 1
        
    End With
    Set Rst = Nothing
End Sub

Private Sub vsPreviousFactors_SelChange()

    FillvsPreviousFactorDetails
    
End Sub

Public Sub Printing()
    On Error GoTo ErrorHandler
''''    frmInput.fwlblInput.Caption = "òœ«„ ”«Ì“ ò«€– „Ê—œ ‰Ÿ— ‘„«”  "
''''    frmInput.OptionLevel(0).Caption = "A4"
''''    frmInput.OptionLevel(1).Caption = " ›—Ê‘ê«ÂÌ"
''''    frmInput.OptionLevel(0).Value = True
''''    frmInput.btnCancel.Visible = True
''''    frmInput.Picture1.Visible = True
''''    frmInput.txtInput.Visible = False
''''
''''    frmInput.Show vbModal
''''    If mvarInput = "" Then
''''        Exit Sub
''''    End If
    
    ReDim Parameter(4) As Parameter
    Parameter(0) = GenerateInputParameter("@SystemDate", adVarWChar, 50, Mid(clsDate.shamsi(Date), 3, 8))
    Parameter(1) = GenerateInputParameter("@SystemDay", adVarWChar, 50, clsDate.Find_DayOfWeek(Weekday(Date, vbSaturday)))
    Parameter(2) = GenerateInputParameter("@SystemTime", adVarWChar, 50, Mid(Str(time), 1, 5))
    Parameter(3) = GenerateInputParameter("@DateAfter", adVarWChar, 20, CStr(IIf(Trim(txtDateFrom.ClipText) = "", "", Trim(txtDateFrom.Text))))
    Parameter(4) = GenerateInputParameter("@DateBefore", adVarWChar, 20, CStr(IIf(Trim(txtDateTo.ClipText) = "", "", Trim(txtDateTo.Text))))

    CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\RepEditedFactor_A4.rpt"

    Dim fileSystem As New FileSystemObject
    Dim IsFileExist As Boolean
    IsFileExist = fileSystem.FileExists(CrystalReport1.ReportFileName)
    If IsFileExist = False Then
        frmDisMsg.lblMessage = " ›«Ì· " & vbLf & Mid(CrystalReport1.ReportFileName, Len(CrystalReport1.ReportFileName) - 36) & " ÅÌœ« ‰‘œ "
        frmDisMsg.Timer1.Interval = 3000
        frmDisMsg.Timer1.Enabled = True
        frmDisMsg.Show vbModal
        Exit Sub
    End If
    
    CrystalReport1.ReportTitle = "ê“«—‘ «“ ›Ì‘ Â«Ì «’·«ÕÌ"
    CrystalReport1.Destination = crptToWindow 'crptToPrinter '
    Dim intIndex As Integer

    For intIndex = 0 To UBound(Parameter) - LBound(Parameter)
        CrystalReport1.ParameterFields(intIndex) = CStr(Parameter(intIndex).Name) & ";" & CStr(Parameter(intIndex).Value) & ";" & "True"
    Next intIndex

    CrystalReport1.RetrieveDataFiles
    ODBCSetting clsArya.ServerName, clsArya.DbName
    CrystalReport1.Connect = CrystallConnection
    CrystalReport1.Action = 1
    If Screen.Width > 12000 Then
        CrystalReport1.PageZoom (100)
    Else
        CrystalReport1.PageZoom (75)
    End If
    
    
    ReDim Parameter(3) As Parameter
    Parameter(0) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
    Parameter(1) = GenerateInputParameter("@DateAfter", adVarWChar, 20, CStr(IIf(Trim(txtDateFrom.ClipText) = "", "", Trim(txtDateFrom.Text))))
    Parameter(2) = GenerateInputParameter("@DateBefore", adVarWChar, 20, CStr(IIf(Trim(txtDateTo.ClipText) = "", "", Trim(txtDateTo.Text))))
    Parameter(3) = GenerateInputParameter("@Flag", adInteger, 4, Val(ChkLessEdited.Value))
    
    If clsStation.Language = Farsi Then
        CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\RepEditedFich_A4.rpt"
    Else
        CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\English\RepEditedFich_A4_En.rpt"
    End If
    IsFileExist = fileSystem.FileExists(CrystalReport1.ReportFileName)
    If IsFileExist = False Then
        frmDisMsg.lblMessage = " ›«Ì· " & vbLf & Mid(CrystalReport1.ReportFileName, Len(CrystalReport1.ReportFileName) - 33) & " ÅÌœ« ‰‘œ "
        frmDisMsg.Timer1.Interval = 3000
        frmDisMsg.Timer1.Enabled = True
        frmDisMsg.Show vbModal
        Exit Sub
    End If
    
    CrystalReport1.ReportTitle = "ê“«—‘ «“ ›Ì‘ Â«Ì «’·«ÕÌ"
    CrystalReport1.Destination = crptToWindow 'crptToPrinter '
   ' Dim intIndex As Integer
   
    For intIndex = 0 To UBound(Parameter) - LBound(Parameter)
        CrystalReport1.ParameterFields(intIndex) = CStr(Parameter(intIndex).Name) & ";" & CStr(Parameter(intIndex).Value) & ";" & "True"
    Next intIndex
  
    CrystalReport1.RetrieveDataFiles
    ODBCSetting clsArya.ServerName, clsArya.DbName
    CrystalReport1.Connect = CrystallConnection
    CrystalReport1.Action = 1
    If Screen.Width > 12000 Then
        CrystalReport1.PageZoom (100)
    Else
        CrystalReport1.PageZoom (75)
    End If

Exit Sub
ErrorHandler:
    MsgBox err.Description
    Resume Next
End Sub


