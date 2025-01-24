VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{7AEDC602-D94C-11D1-BB7A-00E0290EA3C9}#1.0#0"; "ResizeKit.ocx"
Object = "{75D4F767-8785-11D3-93AD-0000832EF44D}#2.14#0"; "FAST2001.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmPayk 
   ClientHeight    =   9840
   ClientLeft      =   4440
   ClientTop       =   5145
   ClientWidth     =   15105
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   12
      Charset         =   178
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPayk.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   9840
   ScaleWidth      =   15105
   Begin VB.TextBox txtBarcode 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   29
      Top             =   600
      Width           =   2415
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7935
      Left            =   120
      TabIndex        =   26
      Top             =   1080
      Width           =   12615
      _ExtentX        =   22251
      _ExtentY        =   13996
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   706
      ForeColor       =   128
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "B Nazanin"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "›«ò Ê—Â«Ì «—”«· ‰‘œÂ"
      TabPicture(0)   =   "frmPayk.frx":A4C2
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "vsNotDeliveredFactors"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "›«ò Ê—Â«Ì «—”«· ‘œÂ"
      TabPicture(1)   =   "frmPayk.frx":A4DE
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "vsDeliveredFactors"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin VSFlex7LCtl.VSFlexGrid vsDeliveredFactors 
         Height          =   7215
         Left            =   150
         TabIndex        =   27
         Top             =   600
         Width           =   12435
         _cx             =   21934
         _cy             =   12726
         _ConvInfo       =   1
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483624
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483645
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483643
         BackColorAlternate=   12648447
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   3
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   1
         Cols            =   12
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   400
         RowHeightMax    =   1200
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmPayk.frx":A4FA
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
      Begin VSFlex7LCtl.VSFlexGrid vsNotDeliveredFactors 
         Height          =   7335
         Left            =   -74880
         TabIndex        =   28
         Top             =   480
         Width           =   12315
         _cx             =   21722
         _cy             =   12938
         _ConvInfo       =   1
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483624
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483645
         ForeColorFixed  =   -2147483630
         BackColorSel    =   16711680
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483643
         BackColorAlternate=   12648447
         GridColor       =   -2147483633
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
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   1
         Cols            =   12
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   400
         RowHeightMax    =   1200
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmPayk.frx":A5DA
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
   End
   Begin VB.ComboBox CmbPayk 
      BackColor       =   &H80000003&
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
      Left            =   11280
      RightToLeft     =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   23
      TabStop         =   0   'False
      ToolTipText     =   "›Ì· — ﬂ—œ‰ ›«ﬂ Ê—Â«Ì «—”«·Ì  Ê”ÿ ÅÌﬂ Â«"
      Top             =   525
      Width           =   2610
   End
   Begin RESIZEKITLibCtl.ResizeKit ResizeKit1 
      Height          =   480
      Left            =   5400
      OleObjectBlob   =   "frmPayk.frx":A6BA
      TabIndex        =   21
      Top             =   120
      Width           =   480
   End
   Begin VB.CheckBox ChkFichUpdate 
      Alignment       =   1  'Right Justify
      Caption         =   "»Â —Ê“ —”«‰Ì ›Ì‘ Â« œ— ‘»ﬂÂ"
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
      Height          =   465
      Left            =   12240
      RightToLeft     =   -1  'True
      TabIndex        =   19
      Top             =   9360
      Value           =   1  'Checked
      Width           =   2655
   End
   Begin VB.CommandButton cmdPaySome 
      BackColor       =   &H000000C0&
      Caption         =   " ”ÊÌÂ Õ”«»"
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
      RightToLeft     =   -1  'True
      TabIndex        =   18
      Top             =   9120
      Width           =   1575
   End
   Begin VB.Timer timRefreshForm 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   7920
      Top             =   9120
   End
   Begin VB.CheckBox ChkDaily 
      Alignment       =   1  'Right Justify
      Caption         =   "«—”«·Ì Â«Ì «„—Ê“"
      BeginProperty Font 
         Name            =   "B Homa"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   405
      Left            =   12840
      RightToLeft     =   -1  'True
      TabIndex        =   12
      Top             =   9000
      Value           =   1  'Checked
      Width           =   2025
   End
   Begin VB.CommandButton cmdReturnFromPaykAccount 
      BackColor       =   &H0080C0FF&
      Caption         =   "»—ê‘  «“ Õ”«» ÅÌò"
      BeginProperty Font 
         Name            =   "B Homa"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   2040
      MaskColor       =   &H000000C0&
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   9120
      Width           =   1245
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   8400
      Top             =   9120
   End
   Begin VSFlex7LCtl.VSFlexGrid vsAvailablePayks 
      Height          =   6975
      Left            =   12840
      TabIndex        =   0
      Top             =   1080
      Width           =   2175
      _cx             =   3836
      _cy             =   12303
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483624
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483645
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483643
      BackColorAlternate=   12648447
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
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   400
      RowHeightMax    =   1200
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmPayk.frx":A740
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
   Begin FLWCtrls.FWScrollText fwScrollTextCust 
      Height          =   555
      Left            =   11280
      TabIndex        =   22
      Top             =   0
      Width           =   3795
      _ExtentX        =   6694
      _ExtentY        =   979
      Caption         =   " »—«Ì ‰„«Ì‘ «—”«·Ì Â««“ ﬂ·Ìœ Ã” ÃÊ  «” ›«œÂ ò‰Ìœ  (F2)"
      BorderStyle     =   6
      FontName        =   "B Homa"
      FontBold        =   0   'False
      FontSize        =   11.25
      Interval        =   10
   End
   Begin Total.UcFont UcFont1 
      Height          =   615
      Left            =   12810
      TabIndex        =   31
      Top             =   8160
      Width           =   2190
      _extentx        =   3863
      _extenty        =   1085
   End
   Begin VB.Label lblBarCode 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      ForeColor       =   &H00FF0000&
      Height          =   405
      Left            =   2520
      RightToLeft     =   -1  'True
      TabIndex        =   30
      ToolTipText     =   "‰„«Ì‘ »«—ﬂœ"
      Top             =   600
      Width           =   2175
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      Caption         =   "«Ì‰ —‰ "
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
      Left            =   3015
      RightToLeft     =   -1  'True
      TabIndex        =   25
      Top             =   135
      Width           =   750
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00004080&
      Height          =   375
      Left            =   3870
      RightToLeft     =   -1  'True
      TabIndex        =   24
      Top             =   135
      Width           =   495
   End
   Begin VB.Label LblAccountYear 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Nazanin"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   495
      Left            =   9240
      TabIndex        =   20
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "»Ì—Ê‰"
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
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   17
      Top             =   120
      Width           =   615
   End
   Begin VB.Label lblPayk 
      Alignment       =   1  'Right Justify
      Caption         =   "œ·ÌÊ—Ì"
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
      Left            =   1440
      RightToLeft     =   -1  'True
      TabIndex        =   16
      Top             =   120
      Width           =   750
   End
   Begin VB.Label lblColorOut 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FF80FF&
      Height          =   375
      Left            =   840
      RightToLeft     =   -1  'True
      TabIndex        =   15
      Top             =   120
      Width           =   495
   End
   Begin VB.Label lblColorPayk 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0000C0C0&
      Height          =   375
      Left            =   2280
      RightToLeft     =   -1  'True
      TabIndex        =   14
      Top             =   120
      Width           =   495
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   " «Œ ’«’ «—”«·Ì »Â ÅÌò"
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
      Left            =   6360
      RightToLeft     =   -1  'True
      TabIndex        =   13
      Top             =   -120
      Width           =   3015
   End
   Begin VB.Label SelectedFactorsNo 
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
      Height          =   285
      Left            =   3510
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   9120
      Width           =   1725
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      Caption         =   "„»·€ ›«ò Ê—Â«Ì «‰ Œ«» ‘œÂ"
      BeginProperty Font 
         Name            =   "B Homa"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   5520
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   9480
      Width           =   2385
   End
   Begin VB.Label SelectedFactorsSum 
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
      Height          =   315
      Left            =   3510
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   9480
      Width           =   1725
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   " ⁄œ«œ ›«ò Ê—Â«Ì «‰ Œ«» ‘œÂ"
      BeginProperty Font 
         Name            =   "B Homa"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   5520
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   9120
      Width           =   2385
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Caption         =   " ⁄œ«œ ›«ò Ê—Â«Ì «—”«· ‰‘œÂ"
      BeginProperty Font 
         Name            =   "B Homa"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   9840
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   9120
      Width           =   2385
   End
   Begin VB.Label lblNotDeliveredFactorsPrice 
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
      Height          =   315
      Left            =   8040
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   9480
      Width           =   1665
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "„»·€ ›«ò Ê—Â«Ì «—”«· ‰‘œÂ"
      BeginProperty Font 
         Name            =   "B Homa"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   9720
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   9480
      Width           =   2505
   End
   Begin VB.Label lblNotDeliveredFactorsNo 
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
      Height          =   285
      Left            =   8040
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   9120
      Width           =   1785
   End
   Begin VB.Label lblMessage 
      Alignment       =   2  'Center
      BackColor       =   &H80000001&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "B Homa"
         Size            =   15.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   390
      Left            =   4800
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   630
      Width           =   6375
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "·Ì”  ÅÌòÂ«"
      BeginProperty Font 
         Name            =   "Titr"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   375
      Left            =   13920
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   600
      Width           =   885
   End
   Begin VB.Menu PaykContextMenu 
      Caption         =   "PaykContextMenu "
      Visible         =   0   'False
      Begin VB.Menu mnuReturnFromPaykAccount 
         Caption         =   "»—ê‘  «“ Õ”«» ÅÌò"
      End
   End
End
Attribute VB_Name = "frmPayk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim clsDate As New clsDate
Private intPpno As Integer
Private mvarbarcode As Boolean
Dim Incharge As EnumIncharge
Dim i As Integer
Dim Parameter() As Parameter
Dim MyFormAddEditMode As EnumAddEditMode
Dim BarcodePaykFlag As Boolean
Dim BarcodeFichFlag As Boolean
Dim intServePlace As EnumServePlace
Dim ClsPrint As New Printing
Public Sub Find()
        frmFindSendDeliveries.Show vbModal
        MyFormAddEditMode = ViewMode   'view
        SetFirstToolBar
End Sub

Private Sub CalculateSelected()
    
    Dim tempPrice As Double
    
    With vsNotDeliveredFactors
        SelectedFactorsNo.Caption = 0
        SelectedFactorsSum.Caption = 0
        If .Rows < 2 Then
            Exit Sub
        End If
        For i = 1 To .Rows - 1
            If Val(.TextMatrix(i, 1)) <> 0 Then
                tempPrice = tempPrice + Val(.TextMatrix(i, 7))
                SelectedFactorsNo.Caption = SelectedFactorsNo.Caption + 1
            End If
        Next i
        SelectedFactorsSum.Caption = tempPrice
    End With
End Sub
Public Sub ExitForm()

    Unload Me

End Sub
Public Sub UpdateDbByFactor()

    Dim S2 As String
    Dim S3 As String
    Dim S4 As String
    Dim S5 As String
    Dim i As Integer

    If vsAvailablePayks.Rows = 1 Then Exit Sub

    timRefreshForm.Enabled = False      ' For Barcode No Change Form
    If Val(vsAvailablePayks.TextMatrix(vsAvailablePayks.Row, 1)) = -1 Then

         S4 = ""
         S5 = ""
        For i = 1 To vsNotDeliveredFactors.Rows - 1
             If Val(vsNotDeliveredFactors.TextMatrix(i, 1)) = -1 Then
                 S4 = S4 & vsNotDeliveredFactors.TextMatrix(i, 0) & ","
                 S5 = S5 & vsNotDeliveredFactors.TextMatrix(i, 2) & ","
             End If
         Next i
         If S4 = "" Then
             FillvsDeliveredFactors
             txtBarcode.SetFocus
             Exit Sub
         End If
         S4 = Left(S4, Len(S4) - 1)
         S5 = Left(S5, Len(S5) - 1)

     End If

    With vsNotDeliveredFactors
        If .Rows > 1 Then
            If Val(.TextMatrix(.Row, 1)) = -1 Then
            
                For i = 1 To vsAvailablePayks.Rows - 1
                    If Val(vsAvailablePayks.TextMatrix(i, 1)) = -1 Then
                        S2 = vsAvailablePayks.TextMatrix(i, 0)
                        S3 = vsAvailablePayks.TextMatrix(i, 2)
                    End If
                Next i
                If S2 <> "" Then
                    Dim ret As Integer
                    ReDim Parameter(5) As Parameter
                    Parameter(0) = GenerateInputParameter("@strSelectedFactors", adVarWChar, 4000, .TextMatrix(.Row, 0))
                    Parameter(1) = GenerateInputParameter("@InCharge", adInteger, 4, S2)
                    Parameter(2) = GenerateInputParameter("@Job", adInteger, 4, EnumIncharge.Payk)
                    Parameter(3) = GenerateInputParameter("@UID", adInteger, 4, mvarCurUserNo)
                    Parameter(4) = GenerateInputParameter("@ActionCode", adInteger, 4, EnumActionLog.GiveFactorToPayk)
                    Parameter(5) = GenerateOutputParameter("@Update", adInteger, 4)
                   
                    ret = RunParametricStoredProcedure("Update_tFacM_InCharge", Parameter)
                    If ret = -1 Then
                        frmMsg.fwlblMsg.Caption = "Ã„⁄ „»·€ «Œ ’«’ œ«œÂ »Â ÅÌﬂ »Ì‘ «“ ”ﬁ›  ⁄ÌÌ‰ ‘œÂ «”  "
                        frmMsg.fwBtn(0).ButtonType = flwButtonOk
                        frmMsg.fwBtn(0).Caption = "ﬁ»Ê·"
                        frmMsg.fwBtn(1).Visible = False
                        frmMsg.Show vbModal
                       
                    Else
                    ' Move to Update_tFacM_InCharge
''''                      '  If mdifrm.ClsActionLog.LogGiveFactorToPayk Then
''''                             ReDim Parameter(2) As Parameter
''''                             Parameter(0) = GenerateInputParameter("@strSelectedFactors", adVarWChar, 4000, S4)
''''                             Parameter(1) = GenerateInputParameter("@UID", adInteger, 4, mvarCurUserNo)
''''                             Parameter(2) = GenerateInputParameter("@ActionCode", adInteger, 4, EnumActionLog.GiveFactorToPayk)
''''                             RunParametricStoredProcedure "InsertHistory_Batch", Parameter
''''
''''                       '  End If
                        
                        Timer1.Enabled = False
                        lblMessage = "›«ò Ê— ‘„«—Â " & .TextMatrix(.Row, 2) & " »Â " & S3 & " «Œ ’«’ œ«œÂ ‘œ"
                            
                      '  Timer1.Interval = 10000
                      '  Timer1.Enabled = True
                     
                        If clsStation.PrintAfterPayk Then
                            PrintAfterPaykAssign S5
                        End If
                        If ChkFichUpdate.Value = 1 Then
                            timRefreshForm.Enabled = True
                        End If
                    End If
                    If vsAvailablePayks.Rows > 1 Then
                        For i = 1 To vsAvailablePayks.Rows - 1
                             vsAvailablePayks.TextMatrix(i, 1) = ""
                        Next i
                    End If
                    
                    FillvsNotDeliveredFactors
                    FillvsDeliveredFactors
                End If
            End If
         End If
    End With
    
    txtBarcode.SetFocus
    
End Sub
Private Sub PrintAfterPaykAssign(strSelectedNo As String)
Dim ii As Long
    While Len(strSelectedNo) <> 0
        ii = InStr(1, strSelectedNo, ",", vbBinaryCompare)
        If ii > 1 Then
            ClsPrint.Printing CLng(Left(strSelectedNo, ii - 1)), clsArya.StationNo, InvoiceFactor, EnumActionLog.InvoicePrint
            strSelectedNo = Mid(strSelectedNo, ii + 1)
'            Debug.Print strRet
        Else
            ClsPrint.Printing Val(strSelectedNo), clsArya.StationNo, InvoiceFactor, EnumActionLog.InvoicePrint
            strSelectedNo = ""
        End If
    Wend

End Sub


Public Sub UpdateDbByPayk()
    Dim i As Integer
    Dim S2 As String
    Dim S3 As String
    
   timRefreshForm.Enabled = False      ' For Barcode No Change Form
    With vsAvailablePayks
        
            If Val(.TextMatrix(.Row, 1)) = -1 Then
            
                S2 = ""
                S3 = ""
                For i = 1 To vsNotDeliveredFactors.Rows - 1
                    If Val(vsNotDeliveredFactors.TextMatrix(i, 1)) = -1 Then
                        S2 = S2 & vsNotDeliveredFactors.TextMatrix(i, 0) & ","
                        S3 = S3 & vsNotDeliveredFactors.TextMatrix(i, 2) & ","
                    End If
                Next i
                If S2 = "" Then
                    FillvsDeliveredFactors
                    txtBarcode.SetFocus
                    Exit Sub
                End If
                S2 = Left(S2, Len(S2) - 1)
                S3 = Left(S3, Len(S3) - 1)
                
                ReDim Parameter(5) As Parameter
                Parameter(0) = GenerateInputParameter("@strSelectedFactors", adVarWChar, 4000, S2)
                Parameter(1) = GenerateInputParameter("@InCharge", adInteger, 4, Val(.TextMatrix(.Row, 0)))
                Parameter(2) = GenerateInputParameter("@Job", adInteger, 4, EnumIncharge.Payk)
                Parameter(3) = GenerateInputParameter("@UID", adInteger, 4, mvarCurUserNo)
                Parameter(4) = GenerateInputParameter("@ActionCode", adInteger, 4, EnumActionLog.GiveFactorToPayk)
                Parameter(5) = GenerateOutputParameter("@Update", adInteger, 4)
                ret = RunParametricStoredProcedure("Update_tFacM_InCharge", Parameter)
                    
                If ret = -1 Then
                        frmMsg.fwlblMsg.Caption = "Ã„⁄ „»·€ «Œ ’«’ œ«œÂ »Â ÅÌﬂ »Ì‘ «“ ”ﬁ›  ⁄ÌÌ‰ ‘œÂ «”  "
                        frmMsg.fwBtn(0).ButtonType = flwButtonOk
                        frmMsg.fwBtn(0).Caption = "ﬁ»Ê·"
                        frmMsg.fwBtn(1).Visible = False
                        frmMsg.Show vbModal
                       
                Else
                    ' Move to Update_tFacM_InCharge
''''                   ' If mdifrm.ClsActionLog.LogGiveFactorToPayk Then
''''                        ReDim Parameter(2) As Parameter
''''                        Parameter(0) = GenerateInputParameter("@strSelectedFactors", adVarWChar, 4000, S2)
''''                        Parameter(1) = GenerateInputParameter("@UID", adInteger, 4, mvarCurUserNo)
''''                        Parameter(2) = GenerateInputParameter("@ActionCode", adInteger, 4, EnumActionLog.GiveFactorToPayk)
''''                        RunParametricStoredProcedure "InsertHistory_Batch", Parameter
''''
''''                  '  End If
                    
                    Timer1.Enabled = False
                    If InStr(1, S2, ",") > 0 Then
                        lblMessage = "›«ò Ê—Â«Ì ‘„«—Â " & S3 & " »Â " & .TextMatrix(.Row, 2) & " «Œ ’«’ œ«œÂ ‘œ"
                    Else
                        lblMessage = "›«ò Ê— ‘„«—Â " & S3 & " »Â " & .TextMatrix(.Row, 2) & " «Œ ’«’ œ«œÂ ‘œ"
                    End If
                   
                    If clsStation.PrintAfterPayk Then
                        PrintAfterPaykAssign S3
                    End If
                    
                    If ChkFichUpdate.Value = 1 Then
                        timRefreshForm.Enabled = True
                    End If
                  ' Timer1.Interval = 10000
                   ' Timer1.Enabled = True
                  
                End If
                If .Rows > 1 Then
                    For i = 1 To .Rows - 1
                         .TextMatrix(i, 1) = 0
                    Next i
                End If
                
                FillvsNotDeliveredFactors
                FillvsDeliveredFactors
            Else
                FillvsDeliveredFactors

            End If
    End With
    txtBarcode.SetFocus
End Sub


Public Sub barcode()
    
    Dim i As Integer
    Dim intRepeatedFactor
    Dim Rst As New ADODB.Recordset
    Dim TempInt As Integer
    
    If Len(lblBarCode) = 12 Then
       lblBarCode = "0" + lblBarCode
'it's correct and don't need change it
'    ElseIf Len(lblBarCode ) = 13 Then
'        If Left(lblBarCode , 1) <> "0" Or (Mid(lblBarCode , 2, 1) = "3" Or Mid(lblBarCode , 2, 1) = "9") Then
'            lblBarCode  = "0" + Left(lblBarCode , 12)
'        End If
    End If
    
    timRefreshForm.Enabled = False      ' For Barcode No Change Form
        
    Select Case Left(lblBarCode, 3)
    
            
        Case EnumIncharge.Payk 'Payk BarCode
                                
            If Rst.State <> 0 Then Rst.Close
            
            ReDim Parameter(0) As Parameter
            
            Parameter(0) = GenerateInputParameter("@pPNo", adInteger, 4, Mid(lblBarCode, 4, 10))
            Set Rst = RunParametricStoredProcedure2Rec("Get_Per_By_pPNo", Parameter)
            
            If Not (Rst.EOF = True And Rst.BOF) Then
                intPpno = Val(Mid(lblBarCode, 4, 10))
            
                  
                
                For i = 1 To vsAvailablePayks.Rows - 1
                    
                    If vsAvailablePayks.TextMatrix(i, 0) = Val(Right(lblBarCode, 10)) Then
                        vsAvailablePayks.TextMatrix(i, 1) = -1 ' True
                        vsAvailablePayks.Row = i
                        BarcodePaykFlag = True
                        
                        Timer1.Enabled = False
                        lblMessage = "ÅÌò" & " " & Val(Right(lblBarCode, 10)) & " " & vsAvailablePayks.TextMatrix(i, 2)
                        
                       ' If clsStation.SoundAlarm = True Then
                            Beep 1000, 500
                            txtBarcode = ""
                       ' End If
                      '  Timer1.Interval = 10000
                      '  Timer1.Enabled = True
                        
                        Exit For
                    End If
                Next i
                
                UpdateDbByPayk
                
            Else
            
                Timer1.Enabled = False
                lblMessage = " ÅÌò" & Val(Right(lblBarCode, 10)) & " œ— ”Ì” „ ÊÃÊœ ‰œ«—œ "
                       
               ' Timer1.Interval = 10000
               ' Timer1.Enabled = True
                 
                
            End If
    
            
        Case Else
        
            TempInt = InStr(1, ConvertToBin(EnumServePlace.Delivery + 10, 5), "1")
            
            If Mid(Left(lblBarCode, 3), TempInt, 1) <> "1" Then Exit Sub
            
                
            If Rst.State <> 0 Then Rst.Close
            ReDim Parameter(2) As Parameter
            Parameter(0) = GenerateInputParameter("@No", adBigInt, 8, Val(Right(lblBarCode, 10)))
            Parameter(1) = GenerateInputParameter("@AccountYear", adSmallInt, 2, AccountYear)
            Parameter(2) = GenerateInputParameter("@Branch", adInteger, 4, CurrentBranch)
            Set Rst = RunParametricStoredProcedure2Rec("Get_DeliveryFactor_By_No", Parameter)
                    
                     
            If Not (Rst.EOF = True And Rst.BOF) Then
                If Rst!Balance = True Or Rst!FacPayment = True Then
                    Timer1.Enabled = False
                    lblMessage = "›Ì‘ " & Val(Right(lblBarCode, 10)) & " ﬁ»·«  ”ÊÌÂ ‘œÂ  "
                    
                   ' Timer1.Interval = 10000
                   ' Timer1.Enabled = True
                ElseIf Not IsNull(Rst!Incharge) Then
                    Timer1.Enabled = False
                    lblMessage = "›Ì‘ " & Val(Right(lblBarCode, 10)) & " ﬁ»·« »Â ÅÌò «Œ ’«’ Ì«› Â «”  "
                    
                  '  Timer1.Interval = 10000
                  '  Timer1.Enabled = True
                ElseIf IsNull(Rst!Incharge) Then
                    Timer1.Enabled = False
                    
                    lblMessage = "›Ì‘ «—”«·Ì ‘„«—Â " & Val(Right(lblBarCode, 10))
                    
                  '  Timer1.Interval = 10000
                  '  Timer1.Enabled = True
                      
                    
                    For i = 1 To vsNotDeliveredFactors.Rows - 1
                    
                        If Val(vsNotDeliveredFactors.TextMatrix(i, 2)) = Val(Mid(lblBarCode, 4, 10)) Then
                            vsNotDeliveredFactors.TextMatrix(i, 1) = -1 ' True
                            Beep 1000, 500
                            txtBarcode = ""
                            BarcodeFichFlag = True
                            vsNotDeliveredFactors.Row = i
                            vsNotDeliveredFactors.ShowCell i, 1
                            Exit For
                        End If
                    Next i
                    
                    UpdateDbByFactor
                End If
            Else
                lblMessage = "›Ì‘ «—”«·Ì ‘„«—Â " & Val(Right(lblBarCode, 10)) & "ÊÃÊœ ‰œ«—œ"
                
              '  Timer1.Interval = 10000
              '  Timer1.Enabled = True
            End If
            
           
    End Select
    
End Sub

Public Sub FillvsAvailablePayks()

    Dim Rst As New ADODB.Recordset
    
    If Rst.State = 1 Then Rst.Close
    
    ReDim Parameter(1) As Parameter
    Parameter(0) = GenerateInputParameter("@Job", adInteger, 4, EnumIncharge.Payk)
    Parameter(1) = GenerateInputParameter("@Branch", adInteger, 4, CurrentBranch)
    Set Rst = RunParametricStoredProcedure2Rec("Get_vw_Per_BY_Job", Parameter)
    
    With vsAvailablePayks
        .Rows = 1
        
        If Not (Rst.EOF = True And Rst.BOF = True) Then
        
            i = 1
            While Rst.EOF = False
                .Rows = .Rows + 1
                .TextMatrix(i, 0) = Rst.Fields("pPno").Value
                .TextMatrix(i, 2) = Rst.Fields("nvcFirstName").Value & " " & Rst.Fields("nvcSurname").Value
'                .TextMatrix(i, 3) = Rst.Fields("pFami").Value
                Rst.MoveNext
                i = i + 1
            Wend
            
        End If
        .AutoSizeMode = flexAutoSizeColWidth
        .AutoSize 0, .Cols - 1
        
    End With
    
    If Rst.State <> 0 Then Rst.Close
    Set Rst = Nothing
    
End Sub
Public Sub FillvsNotDeliveredFactors()

On Error GoTo ErrHandler

    Dim s As String
    Dim Rst As New ADODB.Recordset
    Dim intDistance As Integer
    Dim intWarn As Integer
    If Rst.State = 1 Then Rst.Close
    ReDim Parameter(1) As Parameter
    Parameter(0) = GenerateInputParameter("@AccountYear", adSmallInt, 2, AccountYear)
    Parameter(1) = GenerateInputParameter("@Branch", adInteger, 4, CurrentBranch)
    Set Rst = RunParametricStoredProcedure2Rec("GetCustomersInfo", Parameter)
    
    With vsNotDeliveredFactors
        .Rows = 1
        
        If Not (Rst.EOF = True And Rst.BOF = True) Then
        
            Dim VarToday As String
            VarToday = mvarDate
            i = 1
            While Rst.EOF = False
                If ChkDaily.Value = False Or Rst.Fields("Date").Value = VarToday Then
                    .Rows = .Rows + 1
                    .TextMatrix(i, 0) = Rst.Fields("intSerialNo").Value
                    .TextMatrix(i, 2) = Rst.Fields("No").Value ' Val(Right(Rst.Fields("No").Value, 3))
                    .TextMatrix(i, 3) = Rst.Fields("TempNo").Value ' Val(Right(Rst.Fields("No").Value, 3))
                   .TextMatrix(i, 4) = Rst.Fields("NvcDescription").Value
                   
                    'Dim intCode As Long
                    'intCode = Rst.Fields("Code").Value
                    If Val(Rst!Code) = -1 Then
                        .TextMatrix(i, 5) = ""
                    Else
                        .TextMatrix(i, 5) = Val(Rst!Code)
                    End If
                    .TextMatrix(i, 6) = Rst.Fields("Full Name").Value
                    .TextMatrix(i, 7) = Rst.Fields("SumPrice").Value
                    .TextMatrix(i, 8) = Rst.Fields("Time").Value
                    .TextMatrix(i, 9) = Rst.Fields("Date").Value
                    .TextMatrix(i, 10) = Rst.Fields("ServePlaceName").Value
                    .TextMatrix(i, 11) = Rst.Fields("shiftDescription").Value
                    .TextMatrix(i, 12) = Rst.Fields("Address").Value
                    .TextMatrix(i, 13) = IIf(IsNull(Rst.Fields("TempAddress").Value), "", Rst.Fields("TempAddress").Value)
                    .TextMatrix(i, 14) = IIf(IsNull(Rst!Tafsili), 0, Rst!Tafsili)
                    intServePlace = Rst.Fields("ServePlace").Value
                    intDistance = Rst.Fields("distance").Value
                    If intServePlace = Delivery Then
                        If Rst.Fields("StationId").Value = -1 Then
                            .Cell(flexcpBackColor, .Rows - 1, .FixedCols, .Rows - 1, .Cols - 1) = &H4080&
                        Else
                            .Cell(flexcpBackColor, .Rows - 1, .FixedCols, .Rows - 1, .Cols - 1) = &HC0C0&
                        End If
                    ElseIf intServePlace = Out Then
                        .Cell(flexcpBackColor, .Rows - 1, .FixedCols, .Rows - 1, .Cols - 1) = &HFF80FF
                    End If
                    
                    intWarn = Rst.Fields("intWarn").Value
                    If intWarn = -1 Then
                        .Cell(flexcpBackColor, .Rows - 1, 3, .Rows - 1, 3) = &HFF&
                    End If
                    i = i + 1
                End If
                
                Rst.MoveNext
            Wend
            If .Rows > 1 Then
                lblNotDeliveredFactorsPrice = .Aggregate(flexSTSum, .FixedRows, 7, .Rows - 1, 7)
                lblNotDeliveredFactorsNo = .Aggregate(flexSTCount, .FixedRows, 7, .Rows - 1, 7)
            End If
        End If
        .AutoSizeMode = flexAutoSizeColWidth
        .AutoSize 0, .Cols - 1
        CalculateSelected
    End With
    
    If Rst.State <> 0 Then Rst.Close
    Set Rst = Nothing

Exit Sub
ErrHandler:

     ShowDisMessage "StationId ____" & err.Description, 1000
End Sub
Public Sub FillvsDeliveredFactors()
    
    Dim Rst As New ADODB.Recordset
    Dim intSelectedPayk As Integer
    Dim intDistance As Integer
    Dim intWarn As Integer
    For i = 1 To vsAvailablePayks.Rows - 1
        If Val(vsAvailablePayks.TextMatrix(i, 1)) = -1 Then
            intSelectedPayk = Val(vsAvailablePayks.TextMatrix(i, 0))
        End If
    Next i
    If Rst.State = 1 Then Rst.Close
    
    If intSelectedPayk = 0 Then
        ReDim Parameter(1) As Parameter
        Parameter(0) = GenerateInputParameter("@Job", adInteger, 4, EnumIncharge.Payk)
        Parameter(1) = GenerateInputParameter("@AccountYear", adSmallInt, 2, AccountYear)
        Set Rst = RunParametricStoredProcedure2Rec("Get_vw_NotPaidFactors_By_Job", Parameter)
    Else
        ReDim Parameter(2) As Parameter
        Parameter(0) = GenerateInputParameter("@Job", adInteger, 4, EnumIncharge.Payk)
        Parameter(1) = GenerateInputParameter("@InCharge", adInteger, 4, intSelectedPayk)
        Parameter(2) = GenerateInputParameter("@AccountYear", adSmallInt, 2, AccountYear)
        Set Rst = RunParametricStoredProcedure2Rec("Get_vw_NotPaidFactors_By_Job_InCharge", Parameter)
    End If
    
    With vsDeliveredFactors
        .Rows = 1
        
        If Not (Rst.EOF = True And Rst.BOF = True) Then
        
            i = 1
            Dim VarToday As String
            VarToday = mvarDate
            While Rst.EOF = False
                If ChkDaily.Value = False Or Rst.Fields("Date").Value = VarToday Then
                    If Rst.Fields("Incharge").Value = CmbPayk.ItemData(CmbPayk.ListIndex) Or CmbPayk.ItemData(CmbPayk.ListIndex) = 0 Then
                        .Rows = .Rows + 1
                        .TextMatrix(i, 0) = Rst.Fields("intSerialNo").Value
                        
                        .TextMatrix(i, 2) = Rst.Fields("nvcFirstName").Value & " " & Rst.Fields("nvcSurName").Value
                        .TextMatrix(i, 3) = ""
                        .TextMatrix(i, 4) = Rst.Fields("No").Value ' Val(Right(Rst.Fields("No").Value, 3))
                        .TextMatrix(i, 5) = Rst.Fields("TempNo").Value
                        
                        If Val(Rst!Code) = -1 Then
                            .TextMatrix(i, 6) = ""
                        Else
                            .TextMatrix(i, 6) = Val(Rst!Code)
                        End If
                        
                        .TextMatrix(i, 7) = Rst.Fields("Full Name").Value
                        .TextMatrix(i, 8) = Rst.Fields("SumPrice").Value
                        .TextMatrix(i, 9) = Rst.Fields("Time").Value
                        .TextMatrix(i, 10) = Rst.Fields("Date").Value
                        .TextMatrix(i, 11) = Rst.Fields("ServePlaceName").Value
                        .TextMatrix(i, 12) = Rst.Fields("shiftDescription").Value
                        .TextMatrix(i, 13) = Rst.Fields("Address").Value
                        .TextMatrix(i, 14) = IIf(IsNull(Rst.Fields("TempAddress").Value), "", Rst.Fields("TempAddress").Value)
                        .TextMatrix(i, 15) = IIf(IsNull(Rst!Tafsili), 0, Rst!Tafsili)
                        intServePlace = Rst.Fields("ServePlace").Value
                        intDistance = Rst.Fields("distance").Value
                        If intServePlace = Delivery Then
                            If intDistance = 1 Then
                                .Cell(flexcpBackColor, .Rows - 1, .FixedCols, .Rows - 1, .Cols - 1) = &HFF00&
                            Else
                                .Cell(flexcpBackColor, .Rows - 1, .FixedCols, .Rows - 1, .Cols - 1) = &HC0C0&
                            End If
                        ElseIf intServePlace = Out Then
                            .Cell(flexcpBackColor, .Rows - 1, .FixedCols, .Rows - 1, .Cols - 1) = &HFF80FF
                        ElseIf intServePlace = Internet Then
                            .Cell(flexcpBackColor, .Rows - 1, .FixedCols, .Rows - 1, .Cols - 1) = &H4080&
                        End If
                        intWarn = Rst.Fields("intWarn").Value
                        If intWarn = -1 Then
                            .Cell(flexcpBackColor, .Rows - 1, 3, .Rows - 1, 3) = &HFF&
                        End If
                        i = i + 1
                    End If
                End If
                
                Rst.MoveNext
            Wend
            
        End If
        .AutoSizeMode = flexAutoSizeColWidth
        .AutoSize 0, .Cols - 1
    End With
    
    If Rst.State <> 0 Then Rst.Close
    Set Rst = Nothing
End Sub

Private Sub ChkDaily_Click()
    FillvsDeliveredFactors
    FillvsNotDeliveredFactors
End Sub

Private Sub ChkFichUpdate_Click()
    If ChkFichUpdate.Value = 1 Then
        timRefreshForm.Enabled = True
    Else
        timRefreshForm.Enabled = False
    End If
End Sub



Private Sub CmbPayk_Click()
    If CmbPayk.ListIndex <> -1 Then FillvsDeliveredFactors
End Sub

Private Sub cmdPaySome_Click()

    Dim i As Integer
    Dim s As String
    Dim strPayk As String
    
    s = ""
    With vsNotDeliveredFactors
    
        For i = 1 To .Rows - 1
            If Val(.TextMatrix(i, 1)) = -1 Then
                s = s & .TextMatrix(i, 0) & ","
            End If
        Next i
    End With
    With vsDeliveredFactors
    
        For i = 1 To .Rows - 1
            If Val(.TextMatrix(i, 1)) = -1 Then
                s = s & .TextMatrix(i, 0) & ","
            End If
        Next i
    End With
    
    If s = "" Then txtBarcode.SetFocus: Exit Sub
        
        frmMsg.fwlblMsg.Caption = "¬Ì« „«Ì·Ìœ ›«ò Ê—Â«Ì «‰ Œ«» ‘œÂ —«  ”ÊÌÂ ‰„«ÌÌœ ø"
        
        frmMsg.fwBtn(0).ButtonType = flwButtonOk
        frmMsg.fwBtn(1).ButtonType = flwButtonCancel
        
        frmMsg.fwBtn(0).Caption = "»·Ì"
        frmMsg.fwBtn(1).Caption = "ŒÌ—"
        
        frmMsg.Show vbModal
        
        If modgl.mvarMsgIdx = vbNo Then
            Exit Sub
        End If
                   
        s = Left(s, Len(s) - 1)
        ReDim Parameter(1) As Parameter
        
        Parameter(0) = GenerateInputParameter("@strSelectedFactors", adVarWChar, 4000, s)
        Parameter(1) = GenerateInputParameter("@UID", adInteger, 4, mvarCurUserNo)
        RunParametricStoredProcedure "PayFactors_Payk", Parameter
        
        If clsArya.ExternalAccounting = True Then
            ReceivedSanad
        End If
        
        If mdifrm.ClsActionLog.LogPayCustomerFactor Then
            ReDim Parameter(2) As Parameter
            Parameter(0) = GenerateInputParameter("@strSelectedFactors", adVarWChar, 4000, s)
            Parameter(1) = GenerateInputParameter("@UID", adInteger, 4, mvarCurUserNo)
            Parameter(2) = GenerateInputParameter("@ActionCode", adInteger, 4, EnumActionLog.PayPaykFactor)
            RunParametricStoredProcedure "InsertHistory_Batch", Parameter
            
        End If
        
        If InStr(1, s, ",") > 0 Then
             lblMessage = "›«ò Ê—Â«Ì ‘„«—Â" & s & " Å—œ«Œ  ‘œ "
        Else
             lblMessage = "›«ò Ê— ‘„«—Â" & s & " Å—œ«Œ  ‘œ "
        End If
        
        FillvsNotDeliveredFactors
        FillvsDeliveredFactors

       ' Timer1.Interval = 3000
       ' Timer1.Enabled = True
        txtBarcode.SetFocus
End Sub
Private Sub ReceivedSanad()
Dim SanadNo As Long
Dim Sanadstring As String
Sanadstring = ""
With vsNotDeliveredFactors
    For i = 1 To .Rows - 1
        If Val(.TextMatrix(i, 14)) > 0 And Val(.TextMatrix(i, 1)) = -1 Then
            SanadNo = Accounting.Insert_ReceivedSanadDll(0, EnumRecieveType.CustomerRecieve, CStr(Val(.TextMatrix(i, 14))), "»«»  œ—Ì«›  «“  " & .TextMatrix(i, 6), CLng(Val(.TextMatrix(i, 7))), mvarTafsili)
            If SanadNo > 0 Then
                Sanadstring = Sanadstring & SanadNo & ","
            End If
        End If
    Next i
End With
With vsDeliveredFactors
    For i = 1 To .Rows - 1
        If Val(.TextMatrix(i, 15)) > 0 And Val(.TextMatrix(i, 1)) = -1 Then
            SanadNo = Accounting.Insert_ReceivedSanadDll(0, EnumRecieveType.CustomerRecieve, CStr(Val(.TextMatrix(i, 15))), "»«»  œ—Ì«›  «“  " & .TextMatrix(i, 7), CLng(Val(.TextMatrix(i, 8))), mvarTafsili)
            If SanadNo > 0 Then
                Sanadstring = Sanadstring & SanadNo & ","
            End If
        End If
    Next i
End With
If Sanadstring <> "" Then ShowDisMessage " Ê·Ìœ ”‰œ Õ”«»œ«—Ì »« ‘„«—Â Â«Ì  " & Mid(Sanadstring, 1, Len(Sanadstring) - 1) & " »—«Ì À»  „»«·€ œ—Ì«› Ì «‰Ã«„ ‘œ", 2000
End Sub

Private Sub cmdReturnFromPaykAccount_Click()
    Dim strTemp As String
    If ClsFormAccess.ReturnFromAccountDeliver = False Then
        frmMsg.fwlblMsg.Caption = "‘„« »Â «Ì‰ «„ﬂ«‰ œ” —”Ì ‰œ«—Ìœ"
        frmMsg.fwBtn(0).ButtonType = flwButtonOk
        frmMsg.fwBtn(0).Caption = "ﬁ»Ê·"
        frmMsg.fwBtn(1).Visible = False
        frmMsg.Show vbModal
        txtBarcode.SetFocus
        Exit Sub
    End If
    
    With vsDeliveredFactors
        If .Rows > 1 And .Row > 0 Then
            For i = 1 To .Rows - 1
                If Val(.TextMatrix(i, 1)) = 0 Then
                Else
                    strTemp = strTemp & .TextMatrix(i, 0) & ","
                End If
            Next i
            If strTemp <> "" Then
                strTemp = Left(strTemp, Len(strTemp) - 1)
                
                ReDim Parameter(0) As Parameter
                Parameter(0) = GenerateInputParameter("@strSelectedFactors", adVarWChar, 4000, strTemp)
                RunParametricStoredProcedure "Update_tFacM_InCharge_Null", Parameter
            
                If mdifrm.ClsActionLog.LogRefferFromPaykAccount Then
                    ReDim Parameter(2) As Parameter
                    Parameter(0) = GenerateInputParameter("@strSelectedFactors", adVarWChar, 4000, strTemp)
                    Parameter(1) = GenerateInputParameter("@UID", adInteger, 4, mvarCurUserNo)
                    Parameter(2) = GenerateInputParameter("@ActionCode", adInteger, 4, EnumActionLog.RefferFromPaykAccount)
                    RunParametricStoredProcedure "InsertHistory_Batch", Parameter
                    
                End If
                
            End If
        End If
        FillvsNotDeliveredFactors
        FillvsDeliveredFactors
    End With
    txtBarcode.SetFocus
End Sub

Private Sub Form_Activate()
    Dim i As Integer
    MyFormAddEditMode = ViewMode
    SetFirstToolBar
    VarActForm = Me.Name
    LblAccountYear.Caption = "”«· „«·Ì :" & CInt(AccountYear)
    Me.barcode
    ChkFichUpdate.Value = IIf(clsStation.RefreshFichNo, 1, 0)
    ChkFichUpdate_Click
    lblBarCode = ""
    mvarbarcode = False
    txtBarcode.SetFocus

    If GetSetting(strMainKey, Me.Name, "Flexgrid_Name") <> "" Then
        vsAvailablePayks.Font.Name = GetSetting(strMainKey, Me.Name, "Flexgrid_Name")
    End If
    If GetSetting(strMainKey, Me.Name, "Flexgrid_Size") <> "" Then
        vsAvailablePayks.Font.Size = GetSetting(strMainKey, Me.Name, "Flexgrid_Size")
    End If
    If GetSetting(strMainKey, Me.Name, "Flexgrid_Bold") <> "" Then
        vsAvailablePayks.Font.Bold = GetSetting(strMainKey, Me.Name, "Flexgrid_Bold")
    End If
    
    UcFont1.FontName = vsAvailablePayks.Font.Name
    UcFont1.FontSize = vsAvailablePayks.Font.Size
    UcFont1.FontBold = vsAvailablePayks.Font.Bold
    UcFont1.VarActForm = Me.Name

    vsAvailablePayks.Font.Name = vsAvailablePayks.Font.Name
    vsAvailablePayks.Font.Size = vsAvailablePayks.Font.Size
    vsAvailablePayks.Font.Bold = vsAvailablePayks.Font.Bold
    
    vsDeliveredFactors.Font.Name = vsAvailablePayks.Font.Name
    vsDeliveredFactors.Font.Size = vsAvailablePayks.Font.Size
    vsDeliveredFactors.Font.Bold = vsAvailablePayks.Font.Bold
    
    vsNotDeliveredFactors.Font.Name = vsAvailablePayks.Font.Name
    vsNotDeliveredFactors.Font.Size = vsAvailablePayks.Font.Size
    vsNotDeliveredFactors.Font.Bold = vsAvailablePayks.Font.Bold

    vsAvailablePayks.RowHeightMax = vsAvailablePayks.Height * (vsAvailablePayks.Font.Size) / 100 '8.2
    vsDeliveredFactors.RowHeightMax = vsDeliveredFactors.Height * (vsAvailablePayks.Font.Size) / 100 '8.2
    vsNotDeliveredFactors.RowHeightMax = vsNotDeliveredFactors.Height * (vsAvailablePayks.Font.Size) / 100 '8.2

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call PresetScreenSaver
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call PresetScreenSaver
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    

    If mvarbarcode = True Then
    
    
        Select Case KeyCode
    
            Case 111, 191: '/ Barcode
            
                Me.barcode
                If clsStation.BarcodeAutoEscape = True And BarcodeFichFlag = True And BarcodePaykFlag = True Then
                    Me.ExitForm
                Else
                    lblBarCode = ""
                    mvarbarcode = False
                End If
            Case 48, 96:                '0
                lblBarCode = lblBarCode + "0"
            Case 49, 97:                ' 1
                lblBarCode = lblBarCode + "1"
            Case 50, 98:                '2
                lblBarCode = lblBarCode + "2"
            Case 51, 99:                '3
                lblBarCode = lblBarCode + "3"
            Case 52, 100:               '4
                lblBarCode = lblBarCode + "4"
            Case 53, 101:   '5
                lblBarCode = lblBarCode + "5"
            Case 54, 102:   '6
                lblBarCode = lblBarCode + "6"
            Case 55, 103:       '7
                lblBarCode = lblBarCode + "7"
            Case 56, 104:       '8
                lblBarCode = lblBarCode + "8"
            Case 57, 105:       '9
                lblBarCode = lblBarCode + "9"
    
        End Select
        
    Else
    
        Select Case KeyCode
        
            Case 111, 191: '/ Barcode
            
                mvarbarcode = True
                lblBarCode = ""
                
            
        End Select
        
    End If
    Select Case Shift
          Case 0
              Select Case KeyCode
                  Case 27  ' Esc
                      Me.ExitForm
                  Case 113  ' Esc
                       frmFindSendDeliveries.Show
                  
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
Private Sub FillsPaykCombo()
    Dim rctmp As New ADODB.Recordset
    CmbPayk.Clear
    ReDim Parameter(1) As Parameter
    Parameter(0) = GenerateInputParameter("@Job", adInteger, 4, EnumIncharge.Payk)
    Parameter(1) = GenerateInputParameter("@Branch", adInteger, 4, CurrentBranch)
    Set rctmp = RunParametricStoredProcedure2Rec("Get_vw_Per_BY_Job", Parameter)
    CmbPayk.AddItem "Â„Â ÅÌﬂ Â«"
    CmbPayk.ItemData(0) = 0
    If Not (rctmp.EOF = True And rctmp.BOF = True) Then
        Do While rctmp.EOF <> True
        
            CmbPayk.AddItem CStr(rctmp.Fields("nvcFirstName")) & " " & CStr(rctmp.Fields("nvcSurName"))
            CmbPayk.ItemData(CmbPayk.ListCount - 1) = Val(rctmp.Fields("pPNo"))
            rctmp.MoveNext
        Loop
    End If
    CmbPayk.ListIndex = 0
    rctmp.Close
    Set rctmp = Nothing
End Sub

Private Sub Form_Load()

    If ClsFormAccess.frmPayk = False Then
        Unload Me
        Exit Sub
    End If

    CenterTop Me
    
    VarActForm = Me.Name
        
    Incharge = Payk
    
    With vsAvailablePayks
    
        .Rows = 1
        .Cols = 3
        .ColAlignment(-1) = flexAlignRightCenter
        .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter
        .ColWidth(1) = 500
       ' .ColDataType(1) = flexDTBoolean
        .ColHidden(0) = True
        .TextMatrix(0, 0) = "òœ ÅÌò"
        .TextMatrix(0, 1) = ""
        .TextMatrix(0, 2) = "‰«„ Ê ‰«„ Œ«‰Ê«œêÌ"
        
        .AutoSearch = flexSearchFromCursor
        
    End With
    
    With vsNotDeliveredFactors
        .Rows = 1
        .Cols = 15
        .ColWidth(1) = 500
        '.ColDataType(1) = flexDTBoolean
        .ColHidden(0) = True
        .TextMatrix(0, 0) = "òœ ›Ì‘"
        .TextMatrix(0, 1) = "«‰ Œ«»"
        .TextMatrix(0, 2) = "”—Ì«·"
        .TextMatrix(0, 3) = "‘„«—Â"
        .TextMatrix(0, 4) = "«Œÿ«—"
        .TextMatrix(0, 5) = "«‘ —«ò"
        .TextMatrix(0, 6) = "„‘ —Ì"
        .TextMatrix(0, 7) = "„»·€"
        .TextMatrix(0, 8) = "”«⁄ "
        .TextMatrix(0, 9) = " «—ÌŒ"
        .TextMatrix(0, 10) = "‰Ê⁄"
        .TextMatrix(0, 11) = "‘Ì› "
        .TextMatrix(0, 12) = "¬œ—”"
        .TextMatrix(0, 13) = "¬œ—” „Êﬁ "
        .TextMatrix(0, 14) = " ›÷Ì·Ì"
        .AutoSearch = flexSearchFromCursor
        .ColAlignment(-1) = flexAlignCenterCenter
        .ColAlignment(12) = flexAlignRightCenter
        .ColAlignment(13) = flexAlignRightCenter
        .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter
    
        .ColFormat(7) = "###,###"
    End With
    
    With vsDeliveredFactors
        .Rows = 1
        .Cols = 16
        .ColWidth(1) = 500
      '  .ColDataType(1) = flexDTBoolean
        .ColHidden(0) = True
        .TextMatrix(0, 0) = "òœ ›Ì‘"
        .TextMatrix(0, 1) = "«‰ Œ«»"
        .TextMatrix(0, 2) = "ÅÌﬂ"
        .TextMatrix(0, 3) = "«Œÿ«—"
        .TextMatrix(0, 4) = "”—Ì«·"
        .TextMatrix(0, 5) = "‘„«—Â"
        .TextMatrix(0, 6) = "«‘ —«ò"
        .TextMatrix(0, 7) = "„‘ —Ì"
        .TextMatrix(0, 8) = "„»·€"
        .TextMatrix(0, 9) = "”«⁄  «—”«·"
        .TextMatrix(0, 10) = " «—ÌŒ"
        .TextMatrix(0, 11) = "‰Ê⁄"
        .TextMatrix(0, 12) = "‘Ì› "
        .TextMatrix(0, 13) = "¬œ—”"
        .TextMatrix(0, 14) = "¬œ—” „Êﬁ "
        .TextMatrix(0, 15) = " ›÷Ì·Ì"
        .AutoSearch = flexSearchFromCursor
        .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter
        .ColAlignment(-1) = flexAlignCenterCenter
        .ColAlignment(13) = flexAlignRightCenter
        .ColAlignment(14) = flexAlignRightCenter
    
        .ColFormat(8) = "###,###"
    End With
    
    
    FillvsAvailablePayks
    FillvsNotDeliveredFactors
    FillsPaykCombo
    
    FillvsDeliveredFactors
     
    BarcodeFichFlag = False
    BarcodePaykFlag = False
    
    For Each Obj In Forms
        If TypeOf Obj Is Form And LCase(Obj.Name) = "frminvoice" Then
            lblBarCode = frmInvoice.lblBarCode
        End If
    Next Obj
    
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


''''    timRefreshForm.Enabled = True
    Me.SSTab1.Tab = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    VarActForm = ""
    If PosConnection.State = adStateOpen Then PosConnection.Close: Set PosConnection = Nothing
    Dim Obj As Object

    For Each Obj In Forms
        If LCase(Obj.Name) = "frminvoice" Then
            If ClsFormAccess.frmInvoice = True Then
                frmPayk.Hide
                frmInvoice.Show
                frmInvoice.SetFirstToolBar
            End If
        End If
    Next Obj
    SaveSetting strMainKey, Me.Name, "Left", Me.Left
    SaveSetting strMainKey, Me.Name, "Top", Me.Top


End Sub



Private Sub mnuReturnFromPaykAccount_Click()

    Dim strTemp As String
    
    With vsDeliveredFactors
        If .Rows > 1 And .Row > 0 Then
            For i = 1 To .Rows - 1
                If Val(.TextMatrix(i, 1)) = 0 Then
                Else
                    strTemp = strTemp & .TextMatrix(i, 0) & ","
                End If
            Next i
            If strTemp <> "" Then
                strTemp = Left(strTemp, Len(strTemp) - 1)
                
                ReDim Parameter(0) As Parameter
                Parameter(0) = GenerateInputParameter("@strSelectedFactors", adVarWChar, 4000, strTemp)
                RunParametricStoredProcedure "Update_tFacM_InCharge_Null", Parameter
            
            End If
        End If
        FillvsNotDeliveredFactors
        FillvsDeliveredFactors
    End With
    
End Sub

Private Sub ResizeKit1_ExitResize(ByVal XScale As Double, ByVal YScale As Double)

    If formloadFlag = True Then
        SaveSetting strMainKey, Me.Name, "Height", Me.Height
        SaveSetting strMainKey, Me.Name, "Width", Me.Width
    End If

End Sub

Private Sub Timer1_Timer()

    lblMessage = ""
    Timer1.Enabled = False
End Sub
Private Sub timRefreshForm_Timer()
    FillvsNotDeliveredFactors
    FillvsDeliveredFactors
End Sub

Private Sub UcFont1_FontProperty(m_FontName As Variant, m_FontSize As Variant, m_FontBold As Variant)
'    On Error Resume Next
'    For Each Obj In Me
'        Obj.Font.Name = m_FontName
'        Obj.Font.Size = m_FontSize
'        Obj.Font.Bold = m_FontBold
'        '  Obj.FontName = "times new roman"
'        '  Obj.Alignment = vbLeftJustify
'    Next Obj
    vsAvailablePayks.Font.Name = m_FontName
    vsAvailablePayks.Font.Size = m_FontSize
    vsAvailablePayks.Font.Bold = m_FontBold
    vsAvailablePayks.Refresh
    
    vsDeliveredFactors.Font.Name = m_FontName
    vsDeliveredFactors.Font.Size = m_FontSize
    vsDeliveredFactors.Font.Bold = m_FontBold
    vsDeliveredFactors.Refresh
    
    vsNotDeliveredFactors.Font.Name = m_FontName
    vsNotDeliveredFactors.Font.Size = m_FontSize
    vsNotDeliveredFactors.Font.Bold = m_FontBold
    vsNotDeliveredFactors.Refresh

End Sub

Private Sub vsAvailablePayks_KeyDown(KeyCode As Integer, Shift As Integer)

'    If KeyCode <> 32 Then Exit Sub
'
'    With vsAvailablePayks
'        If .Row > 0 And .Rows > 1 Then
'            For i = 1 To .Rows - 1
'                If i <> .Row Then
'                    .TextMatrix(i, 1) = 0
'                End If
'            Next i
'            .Select .Row, 1
'            .EditCell
'            UpdateDbByPayk
'        End If
'    End With
End Sub


Private Sub vsAvailablePayks_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    With vsAvailablePayks
    
        If .Col = 1 And .Row > 0 And .Rows > 1 And Button = 1 Then
            For i = 1 To .Rows - 1
                If i <> .Row Then
                    .TextMatrix(i, 1) = 0
                End If
            Next i
            .Select .Row, .Col
            .EditCell
            UpdateDbByPayk
        End If
    End With
     
End Sub


Private Sub vsDeliveredFactors_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode <> 32 Then Exit Sub
'    With vsDeliveredFactors
'        If .Row > 0 And .Rows > 1 Then
'            .Select .Row, 1
'            .EditCell
'        End If
'    End With
End Sub

Private Sub vsDeliveredFactors_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
        With vsDeliveredFactors
            If .Col = 1 And .Row > 0 And .Rows > 1 And Button = 1 Then
                .Select .Row, .Col
                .EditCell
                
            End If
        End With
End Sub

Private Sub vsDeliveredFactors_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 And vsDeliveredFactors.MouseRow = vsDeliveredFactors.Row Then
        Me.PopupMenu PaykContextMenu
        
    End If

End Sub

Private Sub vsNotDeliveredFactors_EnterCell()
    If MyFormAddEditMode = ViewMode Then Exit Sub
    With vsNotDeliveredFactors
        If .Row > 0 And (.Col = 3) Then
            
               .Select .Row, .Col
               .EditCell
        End If
    End With
End Sub

Private Sub vsNotDeliveredFactors_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode <> 32 Then Exit Sub
'    With vsNotDeliveredFactors
'        If .Col = 1 And .Row > 0 And .Rows > 1 Then
'            .Select .Row, .Col
'            .EditCell
'            UpdateDbByFactor
'        End If
'    End With
'    CalculateSelected
End Sub

Private Sub vsNotDeliveredFactors_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
        
    With vsNotDeliveredFactors
        If .Col = 1 And .Row > 0 And .Rows > 1 And Button = 1 Then
            .Select .Row, .Col
            .EditCell
            UpdateDbByFactor
        End If
    End With
    CalculateSelected
End Sub
Public Sub SetFirstToolBar()

    AllButton vbOff, True

   mdifrm.Toolbar1.Buttons(13).Enabled = True
    mdifrm.Toolbar1.Buttons(23).Enabled = True
    mdifrm.Toolbar1.Buttons(24).Enabled = True
    mdifrm.Toolbar1.Buttons(25).Enabled = True
    mdifrm.Toolbar1.Buttons(26).Enabled = True
    mdifrm.Toolbar1.Buttons(27).Enabled = True
 
If MyFormAddEditMode = ViewMode Or MyFormAddEditMode = RefferedMode Then   ' View Mode
 
    mdifrm.Toolbar1.Buttons(6).Enabled = False  'Add
    mdifrm.Toolbar1.Buttons(7).Enabled = True  'Edit
    mdifrm.Toolbar1.Buttons(8).Enabled = False  'Enter
    mdifrm.Toolbar1.Buttons(9).Enabled = True   'Esc
    mdifrm.Toolbar1.Buttons(18).Enabled = False   'Reffer

ElseIf MyFormAddEditMode = EditMode Or MyFormAddEditMode = ManipulateMode Then     'Edit
  
    mdifrm.Toolbar1.Buttons(6).Enabled = False  'Add
    mdifrm.Toolbar1.Buttons(7).Enabled = False  'Edit
    mdifrm.Toolbar1.Buttons(8).Enabled = True   'Enter
    mdifrm.Toolbar1.Buttons(9).Enabled = True   'Esc
    mdifrm.Toolbar1.Buttons(18).Enabled = False   'Reffer
    timRefreshForm.Enabled = False
    
End If

  
End Sub

Public Sub Cancel()

    MyFormAddEditMode = ViewMode
    SetFirstToolBar
    If ChkFichUpdate.Value = 1 Then
        timRefreshForm.Enabled = True
    End If
    
End Sub
Public Sub Edit()
    MyFormAddEditMode = EditMode
    SetFirstToolBar
End Sub
Public Sub Update()
            
    Dim Update As Integer
    ReDim Parameter(1) As Parameter
    With vsNotDeliveredFactors
     .Select .Row, 0
    Parameter(0) = GenerateInputParameter("@intSerialNo", adBigInt, 8, Val(.TextMatrix(.Row, 0)))
    Parameter(1) = GenerateInputParameter("@NvcDescription", adVarWChar, 50, Trim(.TextMatrix(.Row, 3)))
    End With
    Update = RunParametricStoredProcedure("Update_tFacM_On_NvcDescription", Parameter)
    MyFormAddEditMode = ViewMode
    SetFirstToolBar
    FillvsNotDeliveredFactors
    If ChkFichUpdate.Value = 1 Then
        timRefreshForm.Enabled = True
    End If
End Sub

