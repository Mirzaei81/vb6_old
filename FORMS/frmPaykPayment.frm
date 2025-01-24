VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{7AEDC602-D94C-11D1-BB7A-00E0290EA3C9}#1.0#0"; "ResizeKit.ocx"
Begin VB.Form frmPaykPayment 
   Caption         =   "                            "
   ClientHeight    =   9240
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   11670
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPaykPayment.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9240
   ScaleWidth      =   11670
   Begin VB.CommandButton CmdCancel 
      BackColor       =   &H00404080&
      Cancel          =   -1  'True
      Caption         =   "Œ—ÊÃ"
      BeginProperty Font 
         Name            =   "B Homa"
         Size            =   11.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   19
      Top             =   8400
      Width           =   1335
   End
   Begin VB.Timer Timer1 
      Left            =   240
      Top             =   720
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000016&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6045
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   600
      Width           =   8895
      Begin VB.CommandButton cmdCreditMove 
         BackColor       =   &H008080FF&
         Caption         =   "«‰ ﬁ«· »Â Õ”«» „‘ —Ì«‰ «⁄ »«—Ì"
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
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   5400
         Width           =   2295
      End
      Begin VB.CheckBox ChkDailyView 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000016&
         Caption         =   "›«ò Ê—Â«Ì «—”«·Ì «„—Ê“"
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   435
         Left            =   6720
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   4440
         Value           =   1  'Checked
         Width           =   2025
      End
      Begin VB.CheckBox chkNoPaykDelivery 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000016&
         Caption         =   "›«ò Ê—Â«Ì «—”«·Ì »œÊ‰ ÅÌò"
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   6480
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   3960
         Width           =   2265
      End
      Begin VB.CommandButton cmdPaySome 
         BackColor       =   &H000000C0&
         Caption         =   " ”ÊÌÂ Õ”«»"
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
         Left            =   6000
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   5520
         Width           =   2775
      End
      Begin VB.CommandButton cmdPayAll 
         BackColor       =   &H00000080&
         Caption         =   " ”ÊÌÂ Õ”«» ò·Ì"
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
         Left            =   3360
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   5520
         Width           =   2295
      End
      Begin VSFlex7LCtl.VSFlexGrid vsDeliveredFactors 
         Height          =   3315
         Left            =   0
         TabIndex        =   1
         Top             =   600
         Width           =   8835
         _cx             =   15584
         _cy             =   5847
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
         FocusRect       =   2
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   3
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   11
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmPaykPayment.frx":A4C2
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
      Begin VB.Label lblSelectedCarryfee 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1200
         RightToLeft     =   -1  'True
         TabIndex        =   27
         Top             =   5040
         Visible         =   0   'False
         Width           =   1995
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "„»·€ ò· ﬂ—«ÌÂ Õ„·"
         BeginProperty Font 
            Name            =   "Nazanin"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   345
         Left            =   1560
         RightToLeft     =   -1  'True
         TabIndex        =   25
         Top             =   3960
         Width           =   1755
      End
      Begin VB.Label lblCarryfee 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   240
         RightToLeft     =   -1  'True
         TabIndex        =   24
         Top             =   3960
         Width           =   1275
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "ﬂ—«ÌÂ Õ„· ›«ﬂ Ê— Â«Ì «‰ Œ«» ‘œÂ"
         BeginProperty Font 
            Name            =   "Nazanin"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   345
         Left            =   3240
         RightToLeft     =   -1  'True
         TabIndex        =   23
         Top             =   5040
         Visible         =   0   'False
         Width           =   3195
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "„»·€ ò· »œÂÌ"
         BeginProperty Font 
            Name            =   "Nazanin"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   345
         Left            =   4800
         RightToLeft     =   -1  'True
         TabIndex        =   17
         Top             =   3960
         Width           =   1635
      End
      Begin VB.Label Lable1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   " ⁄œ«œ ò· ›«ò Ê—Â«"
         BeginProperty Font 
            Name            =   "Nazanin"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   345
         Index           =   1
         Left            =   4800
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Top             =   4320
         Width           =   1635
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "„»·€ ›«ò Ê—Â«Ì «‰ Œ«» ‘œÂ"
         BeginProperty Font 
            Name            =   "Nazanin"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   345
         Left            =   3840
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   4680
         Visible         =   0   'False
         Width           =   2655
      End
      Begin VB.Label lblSelected 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1200
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   4680
         Visible         =   0   'False
         Width           =   1995
      End
      Begin VB.Label lblShouldBePaid 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   3360
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   3960
         Width           =   1575
      End
      Begin VB.Label lblNoOfFactors 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   3360
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   4320
         Width           =   1395
      End
      Begin VB.Label Label1 
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
         Height          =   525
         Index           =   2
         Left            =   3000
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   120
         Width           =   5895
      End
   End
   Begin VSFlex7LCtl.VSFlexGrid vsFactorDetail 
      Height          =   2265
      Left            =   3360
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   6645
      Width           =   8265
      _cx             =   14579
      _cy             =   3995
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
      FocusRect       =   2
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   3
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
      FormatString    =   ""
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
   Begin VSFlex7LCtl.VSFlexGrid vsOwedPayks 
      Height          =   4335
      Left            =   9000
      TabIndex        =   0
      Top             =   1200
      Width           =   2535
      _cx             =   4471
      _cy             =   7646
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
      FocusRect       =   2
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
      FormatString    =   ""
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
   Begin RESIZEKITLibCtl.ResizeKit ResizeKit1 
      Height          =   480
      Left            =   0
      OleObjectBlob   =   "frmPaykPayment.frx":A5B5
      TabIndex        =   22
      Top             =   0
      Width           =   480
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Nazanin"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   10200
      RightToLeft     =   -1  'True
      TabIndex        =   26
      Top             =   8280
      Visible         =   0   'False
      Width           =   1995
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
      Left            =   480
      TabIndex        =   21
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   " œ—Ì«›  „»·€ «—”«·Ì «“ ÅÌò"
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
      Height          =   495
      Left            =   4440
      RightToLeft     =   -1  'True
      TabIndex        =   20
      Top             =   0
      Width           =   3375
   End
   Begin VB.Label lblMessage 
      Alignment       =   1  'Right Justify
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
      Left            =   210
      TabIndex        =   12
      Top             =   6540
      Width           =   3045
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "«ﬁ·«„  ›«ò Ê—"
      BeginProperty Font 
         Name            =   "B Homa"
         Size            =   12
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   525
      Left            =   9360
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   6120
      Width           =   2175
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "·Ì”  ÅÌòÂ«Ì »œÂò«—"
      BeginProperty Font 
         Name            =   "B Homa"
         Size            =   12
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   525
      Left            =   9240
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   720
      Width           =   2055
   End
   Begin VB.Menu PaykContextMenu 
      Caption         =   "PaykContextMenu"
      Visible         =   0   'False
      Begin VB.Menu mnuReturnFromPaykAccount 
         Caption         =   "»—ê‘  «“ Õ”«» ÅÌò"
      End
   End
End
Attribute VB_Name = "frmPaykPayment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim clsDate As New clsDate
Dim Incharge As EnumIncharge
Dim i As Integer
Dim Parameter() As Parameter

Public Sub ExitForm()

    Unload Me
End Sub
Private Sub CalculateSelected()
    
    Dim tempPrice As Double
    Dim tempCarryFee
    With vsDeliveredFactors
        lblSelected.Caption = ""
        If .Rows < 2 Then
            Exit Sub
        End If
        For i = 1 To .Rows - 1
            If Val(.TextMatrix(i, 1)) <> 0 Then
                tempPrice = tempPrice + Val(.TextMatrix(i, 8))
                tempCarryFee = tempCarryFee + Val(.TextMatrix(i, 9))
            End If
        Next i
        If tempPrice <> 0 Then
            lblSelected.Visible = True
            lblSelectedCarryfee.Visible = True
            Label7.Visible = True
            Label5.Visible = True
            lblSelected.Caption = tempPrice
            lblSelectedCarryfee.Caption = tempCarryFee
        Else
            lblSelected.Visible = False
            Label7.Visible = False
            lblSelectedCarryfee.Visible = False
            Label5.Visible = False
        End If
    End With
End Sub

Public Sub FillvsOwedPayks()

    Dim Rst As New ADODB.Recordset

    If Rst.State = 1 Then Rst.Close
    
    'find all payks who owe The store
    ReDim Parameter(1) As Parameter
    Parameter(0) = GenerateInputParameter("@Job", adInteger, 4, EnumIncharge.Payk)
    Parameter(1) = GenerateInputParameter("@Branch", adInteger, 4, CurrentBranch)
    Set Rst = RunParametricStoredProcedure2Rec("Get_Owed_Payks_By_Job", Parameter)
    With vsOwedPayks
        .Rows = 1
        
        If Not (Rst.EOF = True And Rst.BOF = True) Then
        
            i = 1
            While Rst.EOF = False ' fill the Grid
                .Rows = .Rows + 1
                .TextMatrix(i, 0) = Rst.Fields("InCharge").Value
                .TextMatrix(i, 2) = Rst.Fields("nvcFirstName").Value & " " & Rst.Fields("nvcSurName").Value
                Rst.MoveNext
                i = i + 1
            Wend
            
        End If
        .AutoSizeMode = flexAutoSizeColWidth ' set the grid collumns width
        .AutoSize 0, .Cols - 1
        
    End With
    
    If Rst.State <> 0 Then Rst.Close
    
    Set Rst = Nothing

    
End Sub


Public Sub FillvsDeliveredFactors()

    Dim i As Integer
    Dim intSelectedPayk As Integer
    Dim Rst As New ADODB.Recordset
    
    intSelectedPayk = -1
    
    With vsDeliveredFactors 'find all the factors which this payk have to pay them
        
        .Rows = 1
        lblNoOfFactors.Caption = 0
        lblShouldBePaid.Caption = 0
        LblCarryFee.Caption = 0
        vsFactorDetail.Rows = 1
        
        
        If vsOwedPayks.Rows > 1 Then
            For i = 1 To vsOwedPayks.Rows - 1
                If Val(vsOwedPayks.TextMatrix(i, 1)) = -1 Then
                    intSelectedPayk = vsOwedPayks.TextMatrix(i, 0)
                    Exit For
                End If
            Next i
        End If
            
        
        If Rst.State = 1 Then Rst.Close
        ReDim Parameter(2) As Parameter
        If intSelectedPayk <> -1 Then
            Parameter(0) = GenerateInputParameter("@InCharge", adInteger, 4, intSelectedPayk)
        ElseIf intSelectedPayk = -1 And chkNoPaykDelivery.Value = 1 Then
            Parameter(0) = GenerateInputParameter("@InCharge", adInteger, 4, 0)
        Else
            Exit Sub
        End If
        Parameter(1) = GenerateInputParameter("@AccountYear", adSmallInt, 2, AccountYear)
        Parameter(2) = GenerateInputParameter("@Branch", adInteger, 4, CurrentBranch)
        Set Rst = RunParametricStoredProcedure2Rec("Get_DeliveryFactor", Parameter)
                    
        
        If Not (Rst.EOF = True And Rst.BOF = True) Then
        
            i = 1
            On Error Resume Next
            
            Dim VarToday As String
            VarToday = mvarDate
            While Rst.EOF = False 'fill the grid
                If ChkDailyView.Value = False Or Rst.Fields("Date").Value = VarToday Then
                    .Rows = .Rows + 1
                    
                    .TextMatrix(i, 0) = Rst.Fields("intSerialNo").Value
                    .TextMatrix(i, 2) = Rst.Fields("nvcFirstName").Value & " " & Rst.Fields("nvcSurname").Value
                    .TextMatrix(i, 3) = Rst.Fields("No").Value
                    .TextMatrix(i, 4) = Rst.Fields("TempNo").Value
                    .TextMatrix(i, 5) = Rst.Fields("Code").Value
                    .TextMatrix(i, 6) = Rst.Fields("Full Name").Value ' Rst.Fields("Name").Value & " " & Rst.Fields("Family")
                    .TextMatrix(i, 7) = IIf(Rst.Fields("blnCreditCust").Value <> 0, -1, 0) ' Rst.Fields("Name").Value & " " & Rst.Fields("Family")
                    .TextMatrix(i, 8) = Rst.Fields("SumPrice").Value
                    .TextMatrix(i, 9) = Rst.Fields("CarryFeeTotal").Value
                    .TextMatrix(i, 10) = Rst.Fields("Time").Value
                    .TextMatrix(i, 11) = Rst.Fields("Date").Value
                    .TextMatrix(i, 12) = Rst.Fields("SentTime").Value
                    .TextMatrix(i, 13) = Rst.Fields("SentMinute").Value
                    .TextMatrix(i, 14) = Rst.Fields("shiftDescription").Value
                    .TextMatrix(i, 15) = Rst.Fields("Address").Value
                    .TextMatrix(i, 16) = IIf(IsNull(Rst!Tafsili), 0, Rst!Tafsili)
                    
                    lblNoOfFactors.Caption = Val(lblNoOfFactors.Caption) + 1
                    lblShouldBePaid.Caption = Val(lblShouldBePaid.Caption) + Rst.Fields("SumPrice").Value
                    LblCarryFee.Caption = Val(LblCarryFee.Caption) + Rst.Fields("CarryFeeTotal")
                    i = i + 1
                End If
                Rst.MoveNext
            Wend
            On Error GoTo 0
            
        End If
        .AutoSizeMode = flexAutoSizeColWidth ' set the columns' width
        .AutoSize 0, .Cols - 1
    End With
    
    If Rst.State <> 0 Then Rst.Close
    Set Rst = Nothing
    FillvsFactorDetail ' fills the detail of the current factor
    
End Sub

Public Sub FillvsFactorDetail() ' fills the detail of the current factor
    Dim i As Integer
    Dim intselFactor As Double
    
    Dim Rst As New ADODB.Recordset
    
    If Rst.State = 1 Then Rst.Close
    
    With vsFactorDetail
        If vsDeliveredFactors.Rows <= 1 Then Exit Sub ' if there is no factor in the grid
        ' if at least there is one , choose the current one
        intselFactor = vsDeliveredFactors.TextMatrix(vsDeliveredFactors.Row, 0)
        
        ReDim Parameter(1) As Parameter
        
        Parameter(0) = GenerateInputParameter("@intSerialNo", adBigInt, 8, intselFactor)
        Parameter(1) = GenerateInputParameter("@Branch", adInteger, 4, CurrentBranch)
        Set Rst = RunParametricStoredProcedure2Rec("Get_vwFactorDetails_By_intSerialNo", Parameter)
        
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
        .AutoSizeMode = flexAutoSizeColWidth  ' set the collumns' width
        .AutoSize 0, .Cols - 1
    End With
    
    If Rst.State <> 0 Then Rst.Close
    Set Rst = Nothing
    

End Sub

Private Sub ChkDailyView_Click()
    
    FillvsDeliveredFactors
    
End Sub

Private Sub chkNoPaykDelivery_Click()
    If chkNoPaykDelivery.Value = 1 Then
        With vsOwedPayks
            For i = 1 To .Rows - 1
                .TextMatrix(i, 1) = ""
            Next i
        End With
        
        Label1(2).Caption = "·Ì”  ›«ò Ê—Â«Ì «—”«·Ì »œÊ‰ ÅÌò"
        Dim s As String
             
        FillvsDeliveredFactors
        
    Else
        lblNoOfFactors = 0
        lblShouldBePaid = 0
        LblCarryFee = 0
        Label1(2).Caption = ""
        vsDeliveredFactors.Rows = 1
        vsFactorDetail.Rows = 1
        
    End If
End Sub

Private Sub cmdCreditMove_Click()
    Dim i As Integer
    Dim s As String
    Dim strPayk As String
    
    If vsOwedPayks.Rows < 2 And chkNoPaykDelivery.Value = 0 Then
        Exit Sub
    End If

    With vsOwedPayks
        For i = 1 To .Rows - 1
            If Val(.TextMatrix(i, 1)) = -1 Then
                strPayk = .TextMatrix(i, 2)
            End If
        Next i
    End With
    
    If strPayk = "" And chkNoPaykDelivery.Value = False Then
        Exit Sub
    End If
    
    s = ""
    With vsDeliveredFactors
    
        For i = 1 To .Rows - 1
            If Val(.TextMatrix(i, 1)) = -1 And Val(.TextMatrix(i, 7)) = -1 Then
                s = s & .TextMatrix(i, 0) & ","
            End If
        Next i
        If s = "" Then ShowDisMessage " ›Ì‘ «‰ Œ«» ‰‘œÂ Ì« „‘ —Ì «⁄ »«—Ì ‰Ì”  ", 2000: Exit Sub
        
        frmMsg.fwlblMsg.Caption = "¬Ì« „«Ì·Ìœ ›«ò Ê—Â«Ì «‰ Œ«» ‘œÂ —« »Â Õ”«» „‘ —Ì «⁄ »«—Ì „‰ ﬁ· ‰„«∆Ìœ ø"
        
        frmMsg.fwBtn(0).ButtonType = flwButtonOk
        frmMsg.fwBtn(1).ButtonType = flwButtonCancel
        
        frmMsg.fwBtn(0).Caption = "»·Ì"
        frmMsg.fwBtn(1).Caption = "ŒÌ—"
        
        frmMsg.Show vbModal
        
        If modgl.mvarMsgIdx = vbNo Then
            Exit Sub
        End If
                   
        s = Left(s, Len(s) - 1)
        ReDim Parameter(0) As Parameter
        
        Parameter(0) = GenerateInputParameter("@strSelectedFactors", adVarWChar, 4000, s)
        RunParametricStoredProcedure "PayFactors_PayktoCustCredit", Parameter
        
        If mdifrm.ClsActionLog.LogMovePaykToCustomCredit Then
            ReDim Parameter(2) As Parameter
            Parameter(0) = GenerateInputParameter("@strSelectedFactors", adVarWChar, 4000, s)
            Parameter(1) = GenerateInputParameter("@UID", adInteger, 4, mvarCurUserNo)
            Parameter(2) = GenerateInputParameter("@ActionCode", adInteger, 4, EnumActionLog.MovePaykToCustomCredit)
            RunParametricStoredProcedure "InsertHistory_Batch", Parameter
            
        End If
        
        If strPayk <> "" Then
        
            If InStr(1, s, ",") > 0 Then
                 lblMessage = "›«ò Ê—Â«Ì ‘„«—Â" & s & "»Â Õ”«» „‘ —Ì «⁄ »«—Ì „‰ ﬁ· ‘œ‰œ "
            Else
                 lblMessage = "›«ò Ê— ‘„«—Â" & s & "»Â Õ”«» „‘ —Ì «⁄ »«—Ì „‰ ﬁ· ‘œ"
            End If
            
            FillvsOwedPayks
            If vsOwedPayks.Rows > 1 Then
                For i = 1 To vsOwedPayks.Rows - 1
                    If vsOwedPayks.TextMatrix(i, 2) = strPayk Then
                        vsOwedPayks.TextMatrix(i, 1) = -1
                    End If
                Next i
            End If
            
            FillvsDeliveredFactors
            
        Else
            If InStr(1, s, ",") > 0 Then
                 lblMessage = "›«ò Ê—Â«Ì ‘„«—Â" & s & " Å—œ«Œ  ‘œ "
            Else
                 lblMessage = "›«ò Ê— ‘„«—Â" & s & " Å—œ«Œ  ‘œ "
            End If
            
            FillvsDeliveredFactors

        End If
        Timer1.Interval = 3000
        Timer1.Enabled = True
            
    End With

End Sub

Private Sub cmdPayAll_Click()

    Dim i As Integer
    Dim s As String
    Dim strPayk As String
    
        
        If vsOwedPayks.Rows < 2 And chkNoPaykDelivery.Value = 0 Then
            Exit Sub
        End If
    
        With vsOwedPayks
            For i = 1 To .Rows - 1
                If Val(.TextMatrix(i, 1)) = -1 Then
                    strPayk = .TextMatrix(i, 2)
                End If
            Next i
        End With
        
        If strPayk = "" And chkNoPaykDelivery.Value = 0 Then
            Exit Sub
        End If
        
        s = ""
        With vsDeliveredFactors
        
            If .Rows < 2 Then Exit Sub
            
            For i = 1 To .Rows - 1
                    s = s & .TextMatrix(i, 0) & ","
            Next i
            If s = "" Then Exit Sub
            
            If chkNoPaykDelivery.Value = 1 Then
                frmMsg.fwlblMsg.Caption = "¬Ì« „«Ì·Ìœ  „«„ ›«ò Ê—Â«Ì »œÊ‰ ÅÌò —«  ”ÊÌÂ ‰„«ÌÌœ ø "
            ElseIf strPayk <> "" Then
                frmMsg.fwlblMsg.Caption = "¬Ì« „«Ì·Ìœ  „«„ ›«ò Ê—Â«Ì «Ì‰ ÅÌò —«  ”ÊÌÂ ‰„«ÌÌœ ø "
            End If
            frmMsg.fwBtn(0).ButtonType = flwButtonOk
            frmMsg.fwBtn(1).ButtonType = flwButtonCancel
            
            frmMsg.fwBtn(0).Caption = "»·Ì"
            frmMsg.fwBtn(1).Caption = "ŒÌ—"
            
            frmMsg.Show vbModal
            
            If modgl.mvarMsgIdx = vbNo Then
                Exit Sub
            End If
            
            s = Left(s, Len(s) - 1)
            If Len(s) > 4000 Then
                cmdPayAll.Enabled = False
                frmMsg.fwlblMsg.Caption = " ⁄œ«œ ›Ì‘ Â« »Ì‘ «“ Õœ „Ã«“ „Ì »«‘œ " & vbLf & "«“ ﬂ·Ìœ  ”ÊÌÂ Õ”«» «” ›«œÂ ﬂ‰Ìœ"
                frmMsg.fwBtn(0).ButtonType = flwButtonOk
                frmMsg.fwBtn(0).Caption = "ﬁ»Ê·"
                frmMsg.fwBtn(1).Visible = False
                frmMsg.Show vbModal
                Exit Sub
            End If
            
            ReDim Parameter(1) As Parameter
            Parameter(0) = GenerateInputParameter("@strSelectedFactors", adVarWChar, 4000, s)
            Parameter(1) = GenerateInputParameter("@UID", adInteger, 4, mvarCurUserNo)
            RunParametricStoredProcedure "PayFactors_Payk", Parameter
            
            If clsArya.ExternalAccounting = True Then
                ReceivedSanad 1
            End If
                
            If chkNoPaykDelivery.Value = 0 Then
                If mdifrm.ClsActionLog.LogPayPaykFactor Then
                    ReDim Parameter(2) As Parameter
                    Parameter(0) = GenerateInputParameter("@strSelectedFactors", adVarWChar, 4000, s)
                    Parameter(1) = GenerateInputParameter("@UID", adInteger, 4, mvarCurUserNo)
                    Parameter(2) = GenerateInputParameter("@ActionCode", adInteger, 4, EnumActionLog.PayPaykFactor)
                    RunParametricStoredProcedure "InsertHistory_Batch", Parameter
                    
                End If
            Else
                If mdifrm.ClsActionLog.LogPayCustomerFactor Then
                    ReDim Parameter(2) As Parameter
                    Parameter(0) = GenerateInputParameter("@strSelectedFactors", adVarWChar, 4000, s)
                    Parameter(1) = GenerateInputParameter("@UID", adInteger, 4, mvarCurUserNo)
                    Parameter(2) = GenerateInputParameter("@ActionCode", adInteger, 4, EnumActionLog.PayCustomerFactor)
                    RunParametricStoredProcedure "InsertHistory_Batch", Parameter
                    
                End If
            End If
                
            If strPayk <> "" Then
                If InStr(1, s, ",") > 0 Then
                     lblMessage = "›«ò Ê—Â«Ì ‘„«—Â" & s & "  Ê”ÿ " & strPayk & " Å—œ«Œ  ‘œ "
                Else
                     lblMessage = "›«ò Ê— ‘„«—Â" & s & "  Ê”ÿ " & strPayk & " Å—œ«Œ  ‘œ "
                End If
                
                FillvsOwedPayks
                If vsOwedPayks.Rows > 1 Then
                    For i = 1 To vsOwedPayks.Rows - 1
                        If vsOwedPayks.TextMatrix(i, 2) = strPayk Then
                            vsOwedPayks.TextMatrix(i, 1) = -1
                        End If
                    Next i
                End If
                
                FillvsDeliveredFactors
                
            Else
            
                If InStr(1, s, ",") > 0 Then
                     lblMessage = "›«ò Ê—Â«Ì ‘„«—Â" & s & " Å—œ«Œ  ‘œ "
                Else
                     lblMessage = "›«ò Ê— ‘„«—Â" & s & " Å—œ«Œ  ‘œ "
                End If
                
                FillvsDeliveredFactors
                
            End If
            Timer1.Interval = 3000
            Timer1.Enabled = True
                  
                
        End With

End Sub

Private Sub cmdCancel_Click()
  Unload Me
End Sub

Private Sub Form_Activate()

    Dim i As Integer
    AllButton vbOff, True
    mdifrm.Toolbar1.Buttons(23).Enabled = True
    mdifrm.Toolbar1.Buttons(24).Enabled = True
    mdifrm.Toolbar1.Buttons(25).Enabled = True
    mdifrm.Toolbar1.Buttons(26).Enabled = True
    mdifrm.Toolbar1.Buttons(27).Enabled = True
    
    VarActForm = Me.Name
    LblAccountYear.Caption = "”«· „«·Ì :" & CInt(AccountYear)

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

Private Sub Form_Unload(Cancel As Integer)
    VarActForm = ""
    If PosConnection.State = adStateOpen Then PosConnection.Close: Set PosConnection = Nothing
    
    SaveSetting strMainKey, Me.Name, "Left", Me.Left
    SaveSetting strMainKey, Me.Name, "Top", Me.Top


End Sub



Private Sub mnuReturnFromPaykAccount_Click()

    ReDim Parameter(0) As Parameter
    Parameter(0) = GenerateInputParameter("@strSelectedFactors", adVarWChar, 4000, vsDeliveredFactors.TextMatrix(vsDeliveredFactors.Row, 0))
    RunParametricStoredProcedure "Update_tFacM_InCharge_Null", Parameter
    
    If mdifrm.ClsActionLog.LogRefferFromPaykAccount Then
    
        ReDim Parameter(2) As Parameter
        Parameter(0) = GenerateInputParameter("@strSelectedFactors", adVarWChar, 4000, vsDeliveredFactors.TextMatrix(vsDeliveredFactors.Row, 0))
        Parameter(1) = GenerateInputParameter("@UID", adInteger, 4, mvarCurUserNo)
        Parameter(2) = GenerateInputParameter("@ActionCode", adInteger, 4, EnumActionLog.RefferFromPaykAccount)
        RunParametricStoredProcedure "InsertHistory_Batch", Parameter
    
    End If
    
    FillvsDeliveredFactors

End Sub

Private Sub ResizeKit1_ExitResize(ByVal XScale As Double, ByVal YScale As Double)

    If formloadFlag = True Then
        SaveSetting strMainKey, Me.Name, "Height", Me.Height
        SaveSetting strMainKey, Me.Name, "Width", Me.Width
    End If

End Sub

Private Sub Timer1_Timer()
    lblMessage.Caption = ""
    Timer1.Enabled = False
End Sub

Private Sub cmdPaySome_Click()

    Dim i As Integer
    Dim s As String
    Dim strPayk As String
    
    If vsOwedPayks.Rows < 2 And chkNoPaykDelivery.Value = 0 Then
        Exit Sub
    End If

    With vsOwedPayks
        For i = 1 To .Rows - 1
            If Val(.TextMatrix(i, 1)) = -1 Then
                strPayk = .TextMatrix(i, 2)
            End If
        Next i
    End With
    
    If strPayk = "" And chkNoPaykDelivery.Value = False Then
        Exit Sub
    End If
    
    s = ""
    With vsDeliveredFactors
    
        For i = 1 To .Rows - 1
            If Val(.TextMatrix(i, 1)) = -1 Then
                s = s & .TextMatrix(i, 0) & ","
            End If
        Next i
        If s = "" Then Exit Sub
        
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
            ReceivedSanad 0
        End If
        If chkNoPaykDelivery.Value = 0 Then
            If mdifrm.ClsActionLog.LogPayPaykFactor Then
                ReDim Parameter(2) As Parameter
                Parameter(0) = GenerateInputParameter("@strSelectedFactors", adVarWChar, 4000, s)
                Parameter(1) = GenerateInputParameter("@UID", adInteger, 4, mvarCurUserNo)
                Parameter(2) = GenerateInputParameter("@ActionCode", adInteger, 4, EnumActionLog.PayPaykFactor)
                RunParametricStoredProcedure "InsertHistory_Batch", Parameter
                
            End If
        Else
            If mdifrm.ClsActionLog.LogPayCustomerFactor Then
                ReDim Parameter(2) As Parameter
                Parameter(0) = GenerateInputParameter("@strSelectedFactors", adVarWChar, 4000, s)
                Parameter(1) = GenerateInputParameter("@UID", adInteger, 4, mvarCurUserNo)
                Parameter(2) = GenerateInputParameter("@ActionCode", adInteger, 4, EnumActionLog.PayCustomerFactor)
                RunParametricStoredProcedure "InsertHistory_Batch", Parameter
                
            End If
        End If
        
        If strPayk <> "" Then
        
            If InStr(1, s, ",") > 0 Then
                 lblMessage = "›«ò Ê—Â«Ì ‘„«—Â" & s & "  Ê”ÿ " & strPayk & " Å—œ«Œ  ‘œ "
            Else
                 lblMessage = "›«ò Ê— ‘„«—Â" & s & "  Ê”ÿ " & strPayk & " Å—œ«Œ  ‘œ "
            End If
            
            FillvsOwedPayks
            If vsOwedPayks.Rows > 1 Then
                For i = 1 To vsOwedPayks.Rows - 1
                    If vsOwedPayks.TextMatrix(i, 2) = strPayk Then
                        vsOwedPayks.TextMatrix(i, 1) = -1
                    End If
                Next i
            End If
            
            FillvsDeliveredFactors
            
        Else
            If InStr(1, s, ",") > 0 Then
                 lblMessage = "›«ò Ê—Â«Ì ‘„«—Â" & s & " Å—œ«Œ  ‘œ "
            Else
                 lblMessage = "›«ò Ê— ‘„«—Â" & s & " Å—œ«Œ  ‘œ "
            End If
            
            FillvsDeliveredFactors

        End If
        Timer1.Interval = 3000
        Timer1.Enabled = True
            
    End With
End Sub
Private Sub ReceivedSanad(mvarType As Integer)
Dim SanadNo As Long
Dim Sanadstring As String
Sanadstring = ""
With vsDeliveredFactors
    For i = 1 To .Rows - 1
        If Val(.TextMatrix(i, 16)) > 0 And ((Val(.TextMatrix(i, 1)) = -1 And Val(.TextMatrix(i, 7)) = -1) Or mvarType = 1) Then
            SanadNo = Accounting.Insert_ReceivedSanadDll(0, EnumRecieveType.CustomerRecieve, CStr(Val(.TextMatrix(i, 16))), "»«»  œ—Ì«›  «“  " & .TextMatrix(i, 6), CLng(Val(.TextMatrix(i, 8))), mvarTafsili)
            If SanadNo > 0 Then
                Sanadstring = Sanadstring & SanadNo & ","
            End If
        End If
    Next i
    If Sanadstring <> "" Then ShowDisMessage " Ê·Ìœ ”‰œ Õ”«»œ«—Ì »« ‘„«—Â Â«Ì  " & Mid(Sanadstring, 1, Len(Sanadstring) - 1) & " »—«Ì À»  „»«·€ œ—Ì«› Ì «‰Ã«„ ‘œ", 2000
End With
End Sub

Private Sub Form_Load()
    
    If ClsFormAccess.frmPaykPayment = False Then
        Unload Me
        Exit Sub
    End If
    
'    If intVersion = Min Then
'        ShowDisMessage " ”ÊÌÂ Õ”«» »« ÅÌﬂ œ— ‰”ŒÂ Â«Ì »«·« — «„ﬂ«‰ Å–Ì— «” ", 1500
'        Unload Me
'        Exit Sub
'    End If
    
    CenterTop Me
    Incharge = Payk
    
    VarActForm = Me.Name
    
    With vsOwedPayks
        .Rows = 1
        .Cols = 3
        .ColAlignment(-1) = flexAlignRightCenter
        .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter
        .ColWidth(1) = 500
        .ColDataType(1) = flexDTBoolean ' the data in this column is boolean
        .ColHidden(0) = True
        'set the headers of the columns
        .TextMatrix(0, 0) = "òœ ÅÌò"
        .TextMatrix(0, 1) = ""
        .TextMatrix(0, 2) = "‰«„ Ê ‰«„ Œ«‰Ê«œêÌ"
    
        .AutoSearch = flexSearchFromCursor
    
    End With
    
    With vsDeliveredFactors
        .Rows = 1
        .Cols = 17
        .ColAlignment(-1) = flexAlignCenterCenter
        .ColAlignment(6) = flexAlignRightCenter
        .ColAlignment(15) = flexAlignRightCenter
        .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter
'        .ColWidth(1) = 500
        .ColDataType(1) = flexDTBoolean
        .ColDataType(7) = flexDTBoolean
        .ColHidden(0) = True
        
        'set the headers of the columns
        .TextMatrix(0, 0) = "òœ ›Ì‘"
        .TextMatrix(0, 1) = "«‰ Œ«»"
        .TextMatrix(0, 2) = "ÅÌò"
        .TextMatrix(0, 3) = "”—Ì«·"
        .TextMatrix(0, 4) = "‘„«—Â"
        .TextMatrix(0, 5) = "òœ"
        .TextMatrix(0, 6) = "„‘ —Ì"
        .TextMatrix(0, 7) = "«⁄ »«—Ì"
        .TextMatrix(0, 8) = "„»·€"
        .TextMatrix(0, 9) = "ﬂ—«ÌÂ Õ„·"
        .TextMatrix(0, 10) = "”«⁄ "
        .TextMatrix(0, 11) = " «—ÌŒ"
        .TextMatrix(0, 12) = "“„«‰ «—”«·"
        .TextMatrix(0, 13) = "„œ  «—”«·"
        .TextMatrix(0, 14) = "‘Ì› "
        .TextMatrix(0, 15) = "¬œ—”"
        .TextMatrix(0, 16) = " ›÷Ì·Ì"
        .AutoSearch = flexSearchFromCursor
    
        .ColFormat(8) = "###,###"
    End With
    
    With vsFactorDetail
        .Rows = 1
        .Cols = 5
        .ColAlignment(-1) = flexAlignCenterCenter
        .ColAlignment(2) = flexAlignRightCenter
        .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter
        
        'set the headers of the columns
        .TextMatrix(0, 0) = "—œÌ›"
        .TextMatrix(0, 1) = " ⁄œ«œ"
        .TextMatrix(0, 2) = "‰«„ ò«·«"
        .TextMatrix(0, 3) = "›Ì"
        .TextMatrix(0, 4) = "ﬁÌ„ "
    
        .AutoSearch = flexSearchFromCursor
    
    End With
        
    FillvsOwedPayks
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


Private Sub vsDeliveredFactors_KeyDown(KeyCode As Integer, Shift As Integer)
    
    FillvsFactorDetail
    
    If KeyCode <> 32 Then Exit Sub ' if the key is not space bar
    Dim i As Integer
    Dim S2 As String
    
    With vsDeliveredFactors
        If .Row > 0 And .Rows > 1 Then
        
'            For i = 1 To .Rows - 1
'                If i <> .Row Then
'                    .TextMatrix(i, 1) = False
'                End If
'            Next i
            .Select .Row, 1
            .EditCell
            
        End If
    End With

    CalculateSelected
    
End Sub

Private Sub vsDeliveredFactors_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    FillvsFactorDetail
            
    Dim i As Integer
    Dim S2 As String
    
    With vsDeliveredFactors
        If .Col = 1 And .Row > 0 And .Rows > 1 And Button = 1 Then
'            For i = 1 To .Rows - 1
'                If i <> .Row Then
'                    .TextMatrix(i, 1) = False
'                End If
'            Next i
            .Select .Row, .Col
            .EditCell
            
            
        End If
    End With
    
    CalculateSelected

End Sub

Private Sub vsDeliveredFactors_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
        If Button = 2 And vsDeliveredFactors.MouseRow = vsDeliveredFactors.Row Then
            Me.PopupMenu PaykContextMenu
        
        End If


End Sub

Private Sub vsDeliveredFactors_SelChange()
    FillvsFactorDetail
End Sub



Private Sub vsOwedPayks_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> 32 Then Exit Sub ' if the key is not space bar
    Dim i As Integer
    Dim S2 As String
    
    With vsOwedPayks
        If .Row > 0 And .Rows > 1 Then
            For i = 1 To .Rows - 1
                If i <> .Row Then
                    .TextMatrix(i, 1) = ""
                End If
            Next i
            .Select .Row, 1
            .EditCell
            
            If Val(.TextMatrix(.Row, 1)) = -1 Then
                chkNoPaykDelivery.Value = 0
                FillvsDeliveredFactors
                Label1(2).Caption = "·Ì”  ›«ò Ê—Â«Ì «—”«· ‘œÂ  Ê”ÿ " & .TextMatrix(.Row, 2)
            Else
                vsDeliveredFactors.Rows = 1
                vsFactorDetail.Rows = 1
                Label1(2).Caption = ""
                lblNoOfFactors.Caption = 0
                lblShouldBePaid.Caption = 0
            End If
        End If
    End With

End Sub

Private Sub vsOwedPayks_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Dim i As Integer
    
    With vsOwedPayks
        If .Col = 1 And .Row > 0 And .Rows > 1 And Button = 1 Then
            For i = 1 To .Rows - 1
                If i <> .Row Then
                    .TextMatrix(i, 1) = ""
                End If
            Next i
            .Select .Row, .Col
            .EditCell
            
            If Val(.TextMatrix(.Row, 1)) = -1 Then
                chkNoPaykDelivery.Value = 0
                FillvsDeliveredFactors
                Label1(2).Caption = "·Ì”  ›«ò Ê—Â«Ì «—”«· ‘œÂ  Ê”ÿ " & .TextMatrix(.Row, 2)
            Else
                vsDeliveredFactors.Rows = 1
                vsFactorDetail.Rows = 1
                Label1(2).Caption = ""
                lblNoOfFactors.Caption = 0
                lblShouldBePaid.Caption = 0
            End If
        End If
    End With

End Sub


