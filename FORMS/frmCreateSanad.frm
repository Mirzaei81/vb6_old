VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{7AEDC602-D94C-11D1-BB7A-00E0290EA3C9}#1.0#0"; "ResizeKit.ocx"
Object = "{75D4F767-8785-11D3-93AD-0000832EF44D}#2.14#0"; "FAST2001.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmCreateSanad 
   Caption         =   "               "
   ClientHeight    =   10155
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   14115
   Icon            =   "frmCreateSanad.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   10155
   ScaleWidth      =   14115
   Tag             =   "frmCreateSanad"
   Begin VB.Frame Frame4 
      Height          =   1335
      Left            =   3840
      RightToLeft     =   -1  'True
      TabIndex        =   37
      Top             =   6960
      Width           =   4215
      Begin VB.CheckBox chkMoveSandoogh 
         Alignment       =   1  'Right Justify
         Caption         =   "«‰ ﬁ«· ÊÃÊÂ ò«—»—«‰ »Â «Ì‰ ’‰œÊﬁ"
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   11.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   960
         RightToLeft     =   -1  'True
         TabIndex        =   39
         Top             =   240
         Width           =   3015
      End
      Begin VB.ComboBox cmbSandoogh 
         BeginProperty Font 
            Name            =   "B Traffic"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         ItemData        =   "frmCreateSanad.frx":A4C2
         Left            =   1200
         List            =   "frmCreateSanad.frx":A4C4
         RightToLeft     =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   38
         Top             =   720
         Width           =   2775
      End
   End
   Begin VB.CommandButton cmdViewRepUser 
      BackColor       =   &H00008000&
      Caption         =   "‰„«Ì‘ ê“«—‘ ò«—»—«‰"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   840
      RightToLeft     =   -1  'True
      TabIndex        =   36
      Top             =   7560
      Width           =   1935
   End
   Begin VB.ComboBox cmbUsers 
      BeginProperty Font 
         Name            =   "B Traffic"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      ItemData        =   "frmCreateSanad.frx":A4C6
      Left            =   8040
      List            =   "frmCreateSanad.frx":A4C8
      RightToLeft     =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   35
      Top             =   960
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Frame Frame3 
      Height          =   735
      Left            =   8160
      RightToLeft     =   -1  'True
      TabIndex        =   31
      Top             =   7560
      Width           =   5895
      Begin VB.Label LblTafsiliName 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   615
         Left            =   0
         RightToLeft     =   -1  'True
         TabIndex        =   33
         Top             =   120
         Width           =   3855
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   " ⁄‰Ê«‰  ›÷Ì·Ì: "
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   11.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4440
         RightToLeft     =   -1  'True
         TabIndex        =   32
         Top             =   120
         Width           =   1335
      End
   End
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   8160
      RightToLeft     =   -1  'True
      TabIndex        =   29
      Top             =   6960
      Width           =   5895
      Begin VB.Label LblKolMoeinName 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   495
         Left            =   0
         RightToLeft     =   -1  'True
         TabIndex        =   34
         Top             =   120
         Width           =   3855
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   " ⁄‰Ê«‰ ﬂ· Ê „⁄Ì‰: "
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   11.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4200
         RightToLeft     =   -1  'True
         TabIndex        =   30
         Top             =   120
         Width           =   1695
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1815
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   17
      Top             =   8280
      Width           =   13935
      Begin VB.CheckBox chkOldFormat 
         Alignment       =   1  'Right Justify
         Caption         =   "›Ê—„  ﬁœÌ„"
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   11.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   12000
         RightToLeft     =   -1  'True
         TabIndex        =   40
         Top             =   1200
         Width           =   1695
      End
      Begin VB.CommandButton cmbEscape 
         BackColor       =   &H000000C0&
         Cancel          =   -1  'True
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
         Height          =   735
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   26
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox txtNoSanad 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Left            =   9720
         RightToLeft     =   -1  'True
         TabIndex        =   21
         Top             =   1080
         Width           =   975
      End
      Begin VB.CommandButton btnCreate 
         BackColor       =   &H00008000&
         Caption         =   "«ÌÃ«œ ”‰œ Õ”«»œ«—Ì"
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
         Height          =   735
         Left            =   3600
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Top             =   960
         Width           =   1935
      End
      Begin VB.TextBox txtTitle 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   585
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   19
         Top             =   240
         Width           =   5415
      End
      Begin VB.CommandButton cmdShowDocument 
         Caption         =   "‰„«Ì‘ ”‰œ"
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
         Height          =   735
         Left            =   1680
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   960
         Width           =   1575
      End
      Begin MSMask.MaskEdBox txtDate3 
         Height          =   585
         Left            =   6720
         TabIndex        =   22
         Top             =   1080
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   1032
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
      Begin VB.Label LblTafsiliNotice 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "œ— ”Ì” „  ›÷Ì·Ì Œ«·Ì ÊÃÊœ œ«—œ . »—«Ì  Ê·Ìœ ”‰œ «» œ« »—«Ì „‘ —Ì«‰ «⁄ »«—Ì Ê Ì« Å—”‰·  ,   ›÷Ì·Ì «ÌÃ«œ ﬂ‰Ìœ."
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   615
         Left            =   6600
         RightToLeft     =   -1  'True
         TabIndex        =   28
         Top             =   480
         Visible         =   0   'False
         Width           =   7215
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblBalanceNotice 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "”‰œ »«·«‰” ‰Ì”  . œﬁ  ‘Êœ œ— ”Ì” „ ›Ì‘ «—”«· ‰‘œÂ Ì« ›Ì‘  ”ÊÌÂ ‰‘œÂ ÊÃÊœ œ«—œ "
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   495
         Left            =   7440
         RightToLeft     =   -1  'True
         TabIndex        =   27
         Top             =   120
         Visible         =   0   'False
         Width           =   6255
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "‘„«—Â ”‰œ : "
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   11.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   10680
         RightToLeft     =   -1  'True
         TabIndex        =   25
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "  «—ÌŒ ”‰œ: "
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   11.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   8400
         RightToLeft     =   -1  'True
         TabIndex        =   24
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "⁄‰Ê«‰ ”‰œ"
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   11.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5640
         RightToLeft     =   -1  'True
         TabIndex        =   23
         Top             =   360
         Width           =   855
      End
   End
   Begin FLWCtrls.FWLed FWLed1 
      Height          =   615
      Left            =   120
      TabIndex        =   15
      Top             =   120
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   1085
      ColorOff        =   0
   End
   Begin VB.ComboBox cboBranch 
      BeginProperty Font 
         Name            =   "B Homa"
         Size            =   11.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   10920
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   360
      Width           =   2115
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   5160
      Top             =   4320
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
   Begin VB.CommandButton cmdSanadView 
      BackColor       =   &H008080FF&
      Caption         =   " Ê·Ìœ ”‰œ"
      BeginProperty Font 
         Name            =   "B Homa"
         Size            =   11.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   960
      Width           =   1575
   End
   Begin VB.ComboBox cboActionType 
      BeginProperty Font 
         Name            =   "B Homa"
         Size            =   11.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   10920
      Style           =   2  'Dropdown List
      TabIndex        =   8
      Top             =   960
      Width           =   2115
   End
   Begin VSFlex7LCtl.VSFlexGrid VsSanadView 
      Height          =   5385
      Left            =   120
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1560
      Width           =   13935
      _cx             =   24580
      _cy             =   9499
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   -2147483635
      ForeColor       =   -2147483640
      BackColorFixed  =   16761024
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483645
      BackColorAlternate=   -2147483637
      GridColor       =   -2147483632
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
      GridLinesFixed  =   12
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   12
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   500
      RowHeightMax    =   500
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmCreateSanad.frx":A4CA
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
      WordWrap        =   -1  'True
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
   Begin MSMask.MaskEdBox txtDate1 
      Height          =   540
      Left            =   4440
      TabIndex        =   0
      Top             =   960
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   953
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
   Begin MSMask.MaskEdBox txtDate2 
      Height          =   540
      Left            =   1920
      TabIndex        =   1
      Top             =   960
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   953
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
      Left            =   4080
      OleObjectBlob   =   "frmCreateSanad.frx":A5C6
      TabIndex        =   16
      Top             =   120
      Width           =   480
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "”«· „«·Ì"
      BeginProperty Font 
         Name            =   "B Homa"
         Size            =   11.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1800
      RightToLeft     =   -1  'True
      TabIndex        =   14
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "‘⁄»Â"
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
      Left            =   12840
      RightToLeft     =   -1  'True
      TabIndex        =   13
      Top             =   360
      Width           =   1065
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "  Ê·Ìœ ”‰œ Õ”«»œ«—Ì"
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
      Left            =   5880
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   0
      Width           =   2655
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "‰Ê⁄ ⁄„·Ì« "
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
      Left            =   12960
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   960
      Width           =   1065
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Ã„⁄"
      BeginProperty Font 
         Name            =   "B Homa"
         Size            =   11.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3840
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   6960
      Width           =   615
   End
   Begin VB.Label lblSumBede 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   1920
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   6960
      Width           =   1815
   End
   Begin VB.Label lblSumBes 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   240
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   6960
      Width           =   1695
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   " «  «—ÌŒ :"
      BeginProperty Font 
         Name            =   "B Homa"
         Size            =   11.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3240
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   960
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "”‰œ  Ê·Ìœ ‘œÂ «“  «—ÌŒ :"
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
      Left            =   5880
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   960
      Width           =   2175
   End
End
Attribute VB_Name = "frmCreateSanad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim clsDate As New clsDate
Dim i As Integer
Dim Parameter() As Parameter
Dim rs As New ADODB.Recordset
Dim Rst As New ADODB.Recordset
Dim Rctemp As New ADODB.Recordset
Dim Totalprice As Long
Dim documentType As EnumAccDocumentType
Dim KolSandoogh, MoeinSandoogh As Long
Dim KolHazineMali, MoeinHazineMali As Long

Private Sub btnCreate_Click()
    If clsArya.ExternalAccounting = True Then
        If Trim(txtDate3.ClipText) = "" Then
             frmDisMsg.lblMessage = "  «—ÌŒ „⁄ »— Ê«—œ ﬂ‰Ìœ "
             frmDisMsg.Timer1.Enabled = True
             frmDisMsg.Show vbModal
             Exit Sub
        ElseIf Val(txtNoSanad) = 0 Then
             frmDisMsg.lblMessage = " ‘„«—Â ÃœÌœ ”‰œ Õ”«»œ«—Ì „‘Œ’ ‰Ì”   "
             frmDisMsg.Timer1.Enabled = True
             frmDisMsg.Show vbModal
             Exit Sub
        End If
        'ValidateDocument
        If ValidateKol And ValidateMoein And ValidateTafsili Then
'            Select Case LCase(clsArya.AccountSystemName)
'                Case Is = "samar"
'                    'For Kind of the document
'                    documentType = SetKind()
'                    If documentType <> NoDefinition Then
                        Insert_Sanad
'                    End If
'                Case Is = "hamyar"
'                    Insert_Sanad
'            End Select
        End If
     Else
        frmDisMsg.lblMessage = " ‘„« »Â «ÌÃ«œ ”‰œ Õ”«»œ«—Ì œ” —”Ì ‰œ«—Ìœ "
        frmDisMsg.Timer1.Enabled = True
        frmDisMsg.Show vbModal
     End If
End Sub

Private Sub cboActionType_Click()
    TitleView cboActionType.ListIndex
    FillvsSanadView
End Sub
Private Sub TitleView(index As Integer)
    Select Case index
       Case 0
            txtTitle.Text = " Ê·Ìœ « Ê„« Ìﬂ ”‰œ -ﬂ·Ì-«“ " & txtDate1.Text & "  « " & txtDate2.Text & " ‘⁄»Â : " & cboBranch.Text
       Case 1
            txtTitle.Text = " Ê·Ìœ « Ê„« Ìﬂ ”‰œ - ›—Ê‘ -«“ " & txtDate1.Text & "  « " & txtDate2.Text
       Case 2
            txtTitle.Text = " Ê·Ìœ « Ê„« Ìﬂ ”‰œ - Œ—Ìœ -«“ " & txtDate1.Text & "  « " & txtDate2.Text
       Case 3
            txtTitle.Text = " Ê·Ìœ « Ê„« Ìﬂ ”‰œ - Å—œ«Œ  -«“ " & txtDate1.Text & "  « " & txtDate2.Text
       Case 4
            txtTitle.Text = " Ê·Ìœ « Ê„« Ìﬂ ”‰œ - œ—Ì«›  -«“ " & txtDate1.Text & "  « " & txtDate2.Text
    End Select

End Sub


Private Sub cboBranch_Click()
    cboActionType_Click
End Sub

Private Sub chkMoveSandoogh_Click()
Dim i As Long
    If chkMoveSandoogh = 0 Then
        FillvsSanadView
    ElseIf chkMoveSandoogh = 1 Then
        If cmbSandoogh.ListIndex = -1 Then
            cmbSandoogh.SetFocus: Sendkey "{F4}", False
        Else
            FillMoveSandoogh
        End If
    End If
End Sub
Private Sub DeleteMoveSandoogh()
With VsSanadView
    For i = 1 To .Rows - 1
        If InStr(1, .TextMatrix(i, 0), "*") > 0 Then
            .RemoveItem i
        End If
    Next
End With
End Sub
Private Sub FillMoveSandoogh()
    Dim TotalNewSandoogh As Currency
    FillvsSanadView
    With VsSanadView
        For i = 1 To .Rows - 1
            If .ValueMatrix(i, 1) = KolSandoogh And .ValueMatrix(i, 2) = MoeinSandoogh Then
                .Rows = .Rows + 1
                .TextMatrix(.Rows - 1, 0) = .Rows - 1 & "*"
                .TextMatrix(.Rows - 1, 1) = KolSandoogh
                .TextMatrix(.Rows - 1, 2) = MoeinSandoogh
                .TextMatrix(.Rows - 1, 3) = .ValueMatrix(i, 3)
                .TextMatrix(.Rows - 1, 4) = " «‰ ﬁ«· »Â ’‰œÊﬁ " & cmbSandoogh.Text
                .TextMatrix(.Rows - 1, 5) = 0
                .TextMatrix(.Rows - 1, 6) = .ValueMatrix(i, 5)
                TotalNewSandoogh = TotalNewSandoogh + .ValueMatrix(i, 5)
            
            End If
        Next
        If TotalNewSandoogh > 0 Then
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, 0) = .Rows - 1 & "*"
            .TextMatrix(.Rows - 1, 1) = KolSandoogh
            .TextMatrix(.Rows - 1, 2) = MoeinSandoogh
            .TextMatrix(.Rows - 1, 3) = cmbSandoogh.ItemData(cmbSandoogh.ListIndex)
            .TextMatrix(.Rows - 1, 4) = " «‰ ﬁ«· «“ ’‰œÊﬁ ò«—»—«‰"
            .TextMatrix(.Rows - 1, 5) = TotalNewSandoogh
            .TextMatrix(.Rows - 1, 6) = 0
        End If
    End With
    
    DoCalculate

End Sub

Public Function FillAddDecrease() As Boolean
    FillAddDecrease = True
    With VsSanadView
        For i = 1 To .Rows - 1
            If .ValueMatrix(i, 1) = KolHazineMali And .ValueMatrix(i, 2) = MoeinHazineMali And .ValueMatrix(i, 3) = Val(frmReceivedSummary.txtTafsili) Then
                FillAddDecrease = False
                Exit For
            End If
        Next
        If FillAddDecrease = False Then Exit Function
        .Rows = .Rows + 1
        .TextMatrix(.Rows - 1, 0) = .Rows - 1 & "**"
        .TextMatrix(.Rows - 1, 1) = KolHazineMali
        .TextMatrix(.Rows - 1, 2) = MoeinHazineMali
        .TextMatrix(.Rows - 1, 3) = Val(frmReceivedSummary.txtTafsili)
        If Val(frmReceivedSummary.txtDecPrice) > 0 Then
            .TextMatrix(.Rows - 1, 4) = " »«»   ò”—Ì ’‰œÊﬁ  " & frmReceivedSummary.cmbPerson.Text
            .TextMatrix(.Rows - 1, 5) = frmReceivedSummary.txtDecPrice
            .TextMatrix(.Rows - 1, 6) = 0
        Else
            .TextMatrix(.Rows - 1, 4) = " »«»  «÷«›«  ’‰œÊﬁ  " & frmReceivedSummary.cmbPerson.Text
            .TextMatrix(.Rows - 1, 5) = 0
            .TextMatrix(.Rows - 1, 6) = frmReceivedSummary.txtAddPrice
        End If
        
        .Rows = .Rows + 1
        .TextMatrix(.Rows - 1, 0) = .Rows - 1 & "**"
        .TextMatrix(.Rows - 1, 1) = KolSandoogh
        .TextMatrix(.Rows - 1, 2) = MoeinSandoogh
        .TextMatrix(.Rows - 1, 3) = Val(frmReceivedSummary.txtTafsili)
        If Val(frmReceivedSummary.txtDecPrice) > 0 Then
            .TextMatrix(.Rows - 1, 4) = " »«»  Ã»—«‰ ò”—Ì ’‰œÊﬁ  " & frmReceivedSummary.cmbPerson.Text
            .TextMatrix(.Rows - 1, 5) = 0
            .TextMatrix(.Rows - 1, 6) = frmReceivedSummary.txtDecPrice
        Else
            .TextMatrix(.Rows - 1, 4) = " »«»  Ã»—«‰ «÷«›«  ’‰œÊﬁ  " & frmReceivedSummary.cmbPerson.Text
            .TextMatrix(.Rows - 1, 5) = frmReceivedSummary.txtAddPrice
            .TextMatrix(.Rows - 1, 6) = 0
        End If
    
    End With
    
    DoCalculate

End Function



Private Sub cmbEscape_Click()
     Unload Me
End Sub

Private Sub cmbSandoogh_Click()
    chkMoveSandoogh_Click
End Sub

Private Sub cmbUsers_Change()
    FillvsSanadView
End Sub

Private Sub cmdSanadView_Click()

    If Trim(txtDate1.ClipText) = "" Or Trim(txtDate2.ClipText) = "" Then
        frmDisMsg.lblMessage = "  «—ÌŒ „⁄ »— Ê«—œ ﬂ‰Ìœ "
        frmDisMsg.Timer1.Enabled = True
        frmDisMsg.Show vbModal
        ClearDataFlexGrid
    Else
       FillvsSanadView
    End If
'    txtTitle.Text = " Ê·Ìœ « Ê„« Ìﬂ ”‰œ «“ ”Ì” „ ”„—  --" & ClsDate.shamsi(Date)
End Sub

Public Sub Printing()
frmInput.fwlblInput.Caption = "òœ«„ ”«Ì“ ò«€– „Ê—œ ‰Ÿ— ‘„«”  "
frmInput.OptionLevel(0).Caption = "A4"
frmInput.OptionLevel(1).Caption = " ›—Ê‘ê«ÂÌ"
frmInput.OptionLevel(0).Value = True
frmInput.btnCancel.Visible = True
frmInput.Picture1.Visible = True
frmInput.txtInput.Visible = False
                
frmInput.Show vbModal
If mvarInput = "" Then
    Exit Sub
End If

With VsSanadView
    
    RunNonParametricStoredProcedure "Delete_tblPrint_Sanad"
    
    ReDim Parameter(6) As Parameter
    For i = 1 To .Rows - 1
            Parameter(0) = GenerateInputParameter("@Row", adInteger, 4, .TextMatrix(i, 0))
            Parameter(1) = GenerateInputParameter("@Kol", adInteger, 4, .TextMatrix(i, 1))
            Parameter(2) = GenerateInputParameter("@Moein", adInteger, 4, .TextMatrix(i, 2))
            Parameter(3) = GenerateInputParameter("@Tafsili", adInteger, 4, IIf(.TextMatrix(i, 3) = "", 0, .TextMatrix(i, 3)))
            Parameter(4) = GenerateInputParameter("@Description", adWChar, 50, Left(.TextMatrix(i, 4), 50))
            Parameter(5) = GenerateInputParameter("@Bedehkar", adBigInt, 8, .TextMatrix(i, 5))
            Parameter(6) = GenerateInputParameter("@Bestankar", adBigInt, 8, .TextMatrix(i, 6))
            
            RunParametricStoredProcedure "Insert_tblPrint_Sanad", Parameter
    Next i
    ReDim Parameter(2) As Parameter
    Parameter(0) = GenerateInputParameter("@SystemDate", adVarWChar, 50, Mid(clsDate.shamsi(Date), 3, 8))
    Parameter(1) = GenerateInputParameter("@SystemDay", adVarWChar, 50, clsDate.Find_DayOfWeek(Weekday(Date, vbSaturday)))
    Parameter(2) = GenerateInputParameter("@SystemTime", adVarWChar, 50, Mid(Str(time), 1, 5))
    
    If mvarInput = "0" Then
        CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\RepSanad_A4.rpt"
    Else
        CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\RepSanad.rpt"
    End If
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
    CrystalReport1.ReportTitle = txtTitle.Text
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

End With
End Sub

Private Sub cmdShowDocument_Click()
    If clsArya.ExternalAccounting And ClsFormAccess.AccfrmAsnad Then
        modgl.ShowAccountingForm "frmAsnad", "’œÊ— «”‰«œ"
    Else
        ShowDisMessage "‘„« »Â «Ì‰ ›„ œ” —”Ì ‰œ«—Ìœ", 1200
    End If
End Sub

Private Sub cmdViewRepUser_Click()
    frmReceivedSummary.AccessUser = True
    frmReceivedSummary.Show vbModal
End Sub

Private Sub Form_Load()
    
    formloadFlag = False
    
    If clsArya.ExternalAccounting = False Then
        ShowDisMessage "«„ﬂ«‰  Ê·Ìœ ”‰œ Õ”«»œ«—Ì ›ﬁÿ œ— ‰”ŒÂ »« Õ”«»œ«—Ì ÊÃÊœ œ«—œ", 1500
        Unload Me
        Exit Sub
    End If
    
    On Error GoTo ErrHandler
    CenterTop Me
    FWLed1.BackColor = Me.BackColor
    FWLed1.ColorOff = Me.BackColor
    Dim L_Rst  As New ADODB.Recordset
    If clsArya.ExternalAccounting = True Then
       cmdShowDocument.Enabled = True
       btnCreate.Enabled = True
       txtNoSanad.Text = Accounting.MaxSanadNoDll()
        
        ReDim Parameter(0) As Parameter
        Parameter(0) = GenerateInputParameter("@Code", adInteger, 4, EnumAccountCodes.HazineMali)
        Set L_Rst = RunParametricStoredProcedure2Rec("Get_tblAcc_Sale_ByCode", Parameter)
        If L_Rst.EOF <> True And L_Rst.BOF <> True Then
            KolHazineMali = L_Rst.Fields("Kol").Value
            MoeinHazineMali = L_Rst.Fields("Moein").Value
        End If

        ReDim Parameter(0) As Parameter
        Parameter(0) = GenerateInputParameter("@Code", adInteger, 4, EnumAccountCodes.CashRemains)
        Set L_Rst = RunParametricStoredProcedure2Rec("Get_tblAcc_Sale_ByCode", Parameter)
        If L_Rst.EOF <> True And L_Rst.BOF <> True Then
            KolSandoogh = L_Rst.Fields("Kol").Value
            MoeinSandoogh = L_Rst.Fields("Moein").Value
        End If
        
        Set L_Rst = Accounting.FillTafsiliSandooghDll
        If Not (L_Rst.BOF = True And L_Rst.EOF = True) Then
            While L_Rst.EOF = False
                cmbSandoogh.AddItem CStr(L_Rst.Fields("TafsiliName"))
                cmbSandoogh.ItemData(cmbSandoogh.ListCount - 1) = Val(L_Rst.Fields("TafsiliId"))
                L_Rst.MoveNext
            Wend
            cmbSandoogh.ListIndex = -1
        End If
    End If
    Set L_Rst = Nothing
    txtDate1.Text = Mid(clsDate.shamsi(Date), 3)
    txtDate2.Text = Mid(clsDate.shamsi(Date), 3)
    txtDate3.Text = Mid(clsDate.shamsi(Date), 3)

    cboBranch.Clear
'    cboBranch.AddItem "Â„Â ‘⁄»« "
'    cboBranch.ItemData(cboBranch.NewIndex) = 0
    Set rs = RunStoredProcedure2RecordSet("Get_All_Branches")
    Do While rs.EOF = False
        cboBranch.AddItem rs!nvcBranchName
        cboBranch.ItemData(cboBranch.NewIndex) = rs!Branch
        rs.MoveNext
    Loop
    rs.Close
    If cboBranch.ListCount > 0 Then cboBranch.ListIndex = 0

    txtNoSanad.Enabled = False
    With cboActionType
        .Clear
        If clsArya.ExternalAccounting = True Then
            .AddItem "Â— ⁄„·Ì« Ì"
            .ItemData(.NewIndex) = 0
        Else
        
''''        ReDim Parameter(0) As Parameter
''''        Parameter(0) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
''''        Set Rst = RunParametricStoredProcedure2Rec("Get_Action", Parameter)
''''        While Rst.EOF <> True
            
            .AddItem "›—Ê‘"
            .ItemData(.ListCount - 1) = EnumAccountingType.Sale
            .AddItem "Œ—Ìœ"
            .ItemData(.ListCount - 1) = EnumAccountingType.Buy
            .AddItem "Å—œ«Œ "
            .ItemData(.ListCount - 1) = EnumAccountingType.Payment
            .AddItem "œ—Ì«› "
            .ItemData(.ListCount - 1) = EnumAccountingType.Recieved
        End If
        .ListIndex = 0

    End With

    FillUsers

'    frmInput.Show vbModal


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

    FWLed1.Value = CInt(AccountYear)
    FWLed1.BackColor = Me.BackColor
    FWLed1.ColorOff = Me.BackColor
    
    SetFirstToolBar
    Call FlexGridActive
    FillvsSanadView

Exit Sub

ErrHandler:
    ShowDisMessage err.Description, 1000
    

End Sub

Public Sub ExitForm()

    Unload Me
    
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call PresetScreenSaver
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call PresetScreenSaver
End Sub

Public Sub FillvsSanadView()
If formloadFlag = False Then Exit Sub

VsSanadView.Rows = 1
    i = 0

    Select Case cboActionType.ListIndex
      Case 0  'All Operation
        If clsArya.ExternalAccounting Then
            If chkOldFormat.Value = 1 Then
                Call SaleSummary   ''' SaleReturnSummary included
            Else
                Call SaleSummaryCustom   ''' SaleReturnSummary included
            End If
        Else
          Call SaleSummary
          Call SupplierSummaryNoCash
          Call BuyReturn
          Call UserExpensivePayment
          Call UserPersonPayment
          Call UserSupplierPayment
          Call UserCustomerPayment
          Call UserPersonRecieve
          Call UserCustomerRecieve
          Call UserCustomerCheckRecieve
          Call UserSupplierRecieve
        End If
      Case 1  ' Sale
         Call SaleSummary
      Case 2  'Buy
          Call SupplierSummaryNoCash
          Call BuyReturn
      Case 3  'Payment
          Call UserExpensivePayment
          Call UserPersonPayment
          Call UserSupplierPayment
          Call UserCustomerPayment
      Case 4  'Recieved
          Call UserPersonRecieve
          Call UserCustomerRecieve
          Call UserCustomerCheckRecieve
          Call UserSupplierRecieve
    
    End Select

    
    VsSanadView.AutoSizeMode = flexAutoSizeColWidth
    lblSumBede.Caption = 0
    lblSumBes.Caption = 0
        
    DoCalculate
    
End Sub
Private Sub DoCalculate()

    Dim NotExistTafsili As Boolean
    Dim TotalBed, TotalBes As Long
    With VsSanadView
       For i = 1 To .Rows - 1
          TotalBed = TotalBed + Val(.TextMatrix(i, 5))
          TotalBes = TotalBes + Val(.TextMatrix(i, 6))
       
          If Trim(.TextMatrix(i, 3)) = "" Then
             NotExistTafsili = True
          End If
       Next
            
    End With
    lblSumBede = TotalBed
    lblSumBes = TotalBes
    If TotalBed <> TotalBes Then
       lblBalanceNotice.Visible = True
    Else
       lblBalanceNotice.Visible = False
    End If
    If NotExistTafsili = False Then
       LblTafsiliNotice.Visible = False
    Else
       LblTafsiliNotice.Visible = True
    End If
    If clsArya.ExternalAccounting = True Then
        If VsSanadView.Rows > 2 And NotExistTafsili = False Then  'And TotalBed = TotalBes
           btnCreate.Enabled = True
           txtNoSanad.Enabled = True
           txtDate3.Enabled = True
        Else
           btnCreate.Enabled = False
           txtNoSanad.Enabled = False
           txtDate3.Enabled = False
        End If
    End If

End Sub
Private Sub SaleReturnSummary()
    Totalprice = 0
    
    ReDim Parameter(3) As Parameter
    Parameter(0) = GenerateInputParameter("@Branch", adInteger, 4, cboBranch.ItemData(cboBranch.ListIndex))
    Parameter(1) = GenerateInputParameter("@DateBefore", adVarWChar, 8, CStr(IIf(Trim(txtDate1.ClipText) = "", "", Trim(txtDate1.Text))))
    Parameter(2) = GenerateInputParameter("@DateAfter", adVarWChar, 8, CStr(IIf(Trim(txtDate2.ClipText) = "", "", Trim(txtDate2.Text))))
    Parameter(3) = GenerateInputParameter("@Uid", adInteger, 4, cmbUsers.ItemData(cmbUsers.ListIndex))
    Set Rst = RunParametricStoredProcedure2Rec("Get_SaleReturnSummary", Parameter)

    If Rst.EOF <> True And Rst.BOF <> True Then
        
        ReDim Parameter2(0) As Parameter
        Parameter2(0) = GenerateInputParameter("@Code", adInteger, 4, EnumAccountCodes.SaleReturn)
        Set Rctemp = RunParametricStoredProcedure2Rec("Get_tblAcc_Sale_ByCode", Parameter2)
    
        If Rctemp.EOF <> True And Rctemp.BOF <> True Then
           With VsSanadView
                While Rst.EOF <> True
                    .Rows = .Rows + 1
                    i = i + 1
                    .TextMatrix(.Rows - 1, 0) = i
                    .TextMatrix(.Rows - 1, 1) = Rctemp.Fields("Kol").Value
                    .TextMatrix(.Rows - 1, 2) = Rctemp.Fields("Moein").Value
                    .TextMatrix(.Rows - 1, 3) = IIf(IsNull(Rctemp.Fields("Tafsili").Value), "", Rctemp.Fields("Tafsili").Value)
                    .TextMatrix(.Rows - 1, 4) = Rctemp.Fields("Description").Value & "-- ‰ﬁœÌ " & Rst!Date
                    .TextMatrix(.Rows - 1, 5) = Rst!SumPriceTotal  'TotalPrice
                    .TextMatrix(.Rows - 1, 6) = 0
                    Rst.MoveNext
                Wend
          End With
        End If
    End If
End Sub

Private Sub Form_Activate()
    VarActForm = Me.Name
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
    
    If rs.State = 1 Then rs.Close
    If Rst.State = 1 Then Rst.Close
    If Rctemp.State = 1 Then Rctemp.Close
    Set rs = Nothing
    Set Rst = Nothing
    Set Rctemp = Nothing
    If PosConnection.State = adStateOpen Then PosConnection.Close: Set PosConnection = Nothing
    VarActForm = ""
    SaveSetting strMainKey, Me.Name, "Left", Me.Left
    SaveSetting strMainKey, Me.Name, "Top", Me.Top

End Sub

Private Sub SetFirstToolBar()

    AllButton vbOff, True
    
    mdifrm.Toolbar1.Buttons(15).Enabled = True  'printing
    mdifrm.Toolbar1.Buttons(23).Enabled = True
    mdifrm.Toolbar1.Buttons(24).Enabled = True
    mdifrm.Toolbar1.Buttons(25).Enabled = True
    mdifrm.Toolbar1.Buttons(26).Enabled = True
    mdifrm.Toolbar1.Buttons(27).Enabled = True

End Sub






Private Sub lblSumBede_Change()
    If Val(lblSumBede) > 0 Then lblSumBede = Format(lblSumBede, "###,###")
    If Val(lblSumBede) = 0 Then lblSumBede = 0
End Sub

Private Sub lblSumBes_Change()
    If Val(lblSumBes) > 0 Then lblSumBes = Format(lblSumBes, "###,###")
    If Val(lblSumBes) = 0 Then lblSumBes = 0
End Sub

Private Sub ResizeKit1_ExitResize(ByVal XScale As Double, ByVal YScale As Double)

    If formloadFlag = True Then
        SaveSetting strMainKey, Me.Name, "Height", Me.Height
        SaveSetting strMainKey, Me.Name, "Width", Me.Width
    End If

End Sub

Private Sub txtDate1_Change()
    TitleView cboActionType.ListIndex
End Sub

Private Sub txtDate2_Change()
    TitleView cboActionType.ListIndex
    If Len(txtDate2.ClipText) = 6 Then txtDate3.Text = txtDate2.Text
End Sub

Private Sub vsSanadView_AfterSort(ByVal Col As Long, Order As Integer)
    For i = 1 To VsSanadView.Rows - 1
        VsSanadView.TextMatrix(i, 0) = i
    Next
End Sub
Sub ClearDataFlexGrid()

    With VsSanadView
        .Rows = 1
        .Rows = 8
        .Row = 1
    End With
    lblSumBede.Caption = ""
    lblSumBes.Caption = ""
End Sub

Private Sub FlexGridActive()

    With VsSanadView
        .Rows = 8
        .Cols = 7
        
        Select Case clsStation.Language
            Case EnumLanguage.Farsi
                
                .RightToLeft = True
                .TextMatrix(0, 0) = "—œÌ›"
                .TextMatrix(0, 1) = "ﬂ·"
                .TextMatrix(0, 2) = "„⁄Ì‰"
                .TextMatrix(0, 3) = " ›÷Ì·Ì"
                .TextMatrix(0, 4) = "‘—Õ ”‰œ"
                .TextMatrix(0, 5) = "»œÂﬂ«—"
                .TextMatrix(0, 6) = "»” «‰ﬂ«—"
            
            Case EnumLanguage.English
            
                .RightToLeft = False
                .TextMatrix(0, 0) = "Row"
                .TextMatrix(0, 1) = "Kol"
                .TextMatrix(0, 2) = "Moein"
                .TextMatrix(0, 3) = "Tafsili"
                .TextMatrix(0, 4) = "Description"
                .TextMatrix(0, 5) = "Bede"
                .TextMatrix(0, 6) = "Bes"
                
        End Select
        .ColFormat(5) = "###,###"
        .ColFormat(6) = "###,###"

        .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter
        .ColAlignment(-1) = flexAlignCenterCenter
        .ColAlignment(4) = flexAlignRightBottom
        
        
            .ColWidth(0) = .Width / 20
            .ColWidth(1) = .Width / 15
            .ColWidth(2) = .Width / 15
            .ColWidth(3) = .Width / 13
            .ColWidth(4) = .Width / 2
            .ColWidth(5) = .Width / 9.1
            .ColWidth(6) = .Width / 9.1
        

        .RowHeightMax = .Height / 10
        .RowHeightMin = .Height / 10.5
        .ScrollBars = flexScrollBarVertical
        .Row = 1
    End With


End Sub

Private Sub Insert_Sanad()
 
 Dim st As String, st1 As String, st2 As String, st3 As String
 Dim kk, jj, mm As Integer
 Dim Result As Long
 On Error GoTo ErrHandler
    
        With VsSanadView
            kk = (.Rows - 1) \ 20
            jj = (.Rows - 1) Mod 20
            If kk > 0 Then
                For mm = 1 To kk
                    st = ""
                    For i = (mm - 1) * 20 + 1 To mm * 20
                        If Len(.TextMatrix(i, 1)) > 0 And Len(.TextMatrix(i, 2)) > 0 And Len(.TextMatrix(i, 3)) > 0 Then st = GenerateDetailsStringAccount(st, AccountYear, 1, txtNoSanad.Text, CStr(i), .TextMatrix(i, 1), .TextMatrix(i, 2), .TextMatrix(i, 3), .TextMatrix(i, 4), CLng(Val(.TextMatrix(i, 5))), CLng(Val(.TextMatrix(i, 6))), IIf(Val(.TextMatrix(i, 5)) > 0, 0, 1), DateToNumber(clsDate.shamsi(Date)), mvarCurUserNo, "", "")
                    Next i
                    Select Case mm
                        Case 1
                            st1 = st
                        Case 2
                            st2 = st
                        Case 3
                            st3 = st
                    End Select
                Next mm
                st = ""
                For i = (kk * 20) + 1 To .Rows - 1
                    If Len(.TextMatrix(i, 1)) > 0 And Len(.TextMatrix(i, 2)) > 0 And Len(.TextMatrix(i, 3)) > 0 Then st = GenerateDetailsStringAccount(st, AccountYear, 1, txtNoSanad.Text, CStr(i), .TextMatrix(i, 1), .TextMatrix(i, 2), .TextMatrix(i, 3), .TextMatrix(i, 4), CLng(Val(.TextMatrix(i, 5))), CLng(Val(.TextMatrix(i, 6))), IIf(Val(.TextMatrix(i, 5)) > 0, 0, 1), DateToNumber(clsDate.shamsi(Date)), mvarCurUserNo, "", "")
                Next i
                If st2 = "" Then
                    st2 = st
                ElseIf st3 = "" Then
                    st3 = st
                Else
                    MsgBox " ⁄œ«œ —œÌ› Â« «“ Õœ „Ã«“ »Ì‘ — «”  . „ÕœÊœÂ  «—ÌŒ —« »—«Ì  Ê·Ìœ ”‰œ ﬂ„ — ﬂ‰Ìœ"
                    Exit Sub
                End If
            Else
                st = ""
                For i = (kk * 20) + 1 To .Rows - 1
                    If Len(.TextMatrix(i, 1)) > 0 And Len(.TextMatrix(i, 2)) > 0 And Len(.TextMatrix(i, 3)) > 0 Then st = GenerateDetailsStringAccount(st, AccountYear, CStr(CurrentBranch), txtNoSanad.Text, CStr(i), .TextMatrix(i, 1), .TextMatrix(i, 2), .TextMatrix(i, 3), .TextMatrix(i, 4), CLng(Val(.TextMatrix(i, 5))), CLng(Val(.TextMatrix(i, 6))), IIf(Val(.TextMatrix(i, 5)) > 0, 0, 1), DateToNumber(clsDate.shamsi(Date)), mvarCurUserNo, "", "")
                Next i
                st1 = st
            End If
        End With
    Dim Status  As Integer
    Status = mvarStatus
    Result = Accounting.Insert_Sanad_Dll(CLng(txtNoSanad.Text), txtDate3.Text, txtTitle.Text _
            , st1, st2, st3, 0, Status, 0)
            
    If Result > 0 Then
            
            UpdateTransferAccounting Result
            frmDisMsg.lblMessage = " ”‰œ Õ”«»œ«—Ì »Â ‘„«—Â  " & Result & "  »« „Ê›ﬁÌ  «ÌÃ«œ ê—œÌœ"
            frmDisMsg.Timer1.Enabled = True
            frmDisMsg.Show vbModal
            txtNoSanad.Enabled = False
            ClearDataFlexGrid
            txtNoSanad.Text = Accounting.MaxSanadNoDll()
    
    Else
        GoTo ErrHandler
    End If

    txtNoSanad.Enabled = False

Exit Sub
ErrHandler:
        frmDisMsg.lblMessage = "œ— À»  ”‰œ Õ”«»œ«—Ì „‘ﬂ· ÊÃÊœ œ«—œ "
        frmDisMsg.Timer1.Enabled = True
        frmDisMsg.Show vbModal
        txtNoSanad.Enabled = True
        txtNoSanad.SetFocus

End Sub
  
Private Sub FillUsers()
    Dim Rst As New ADODB.Recordset
    cmbUsers.Visible = True
    
    ReDim Parameter(0) As Parameter
    Parameter(0) = GenerateInputParameter("@Language", adInteger, 4, 0)
    Set Rst = RunParametricStoredProcedure2Rec("Get_All_tUser", Parameter)
    
    cmbUsers.Clear
    cmbUsers.AddItem "Â„Â ﬂ«—»—«‰ "
    cmbUsers.ItemData(0) = 0
    If Not (Rst.BOF = True And Rst.EOF = True) Then
        While Rst.EOF <> True
            If Rst!UserID <> 0 Then
                cmbUsers.AddItem Trim(Rst.Fields("FullUserName"))
                cmbUsers.ItemData(cmbUsers.NewIndex) = Rst.Fields("UserId")
            End If
            Rst.MoveNext
        Wend
    End If
'    If cmbUsers.ListCount > 0 Then
'        For i = 0 To cmbUsers.ListCount - 1
'            If cmbUsers.ItemData(i) = mvarCurUserNo Then
'                 cmbUsers.ListIndex = i
'                 Exit For
'            End If
'        Next
'    End If
    cmbUsers.ListIndex = 0
    If Rst.State = 1 Then Rst.Close

    Set Rst = Nothing

End Sub

Private Sub SaleSummary()
' Sale Summary
    
    Dim sumPrice As Long
    Dim Discount As Long
    Dim sumPacking As Long
    Dim sumCarryFee As Long
    Dim sumService As Long

    ReDim Parameter(3) As Parameter
    Parameter(0) = GenerateInputParameter("@Branch", adInteger, 4, cboBranch.ItemData(cboBranch.ListIndex))
    Parameter(1) = GenerateInputParameter("@DateBefore", adVarWChar, 8, CStr(IIf(Trim(txtDate1.ClipText) = "", "", Trim(txtDate1.Text))))
    Parameter(2) = GenerateInputParameter("@DateAfter", adVarWChar, 8, CStr(IIf(Trim(txtDate2.ClipText) = "", "", Trim(txtDate2.Text))))
    Parameter(3) = GenerateInputParameter("@Uid", adInteger, 4, cmbUsers.ItemData(cmbUsers.ListIndex))
    Set Rst = RunParametricStoredProcedure2Rec("Get_SaleSummary", Parameter)

'''''''''''   ›—Ê‘


    If Rst.EOF <> True And Rst.BOF <> True Then

        ReDim Parameter2(0) As Parameter
        Parameter2(0) = GenerateInputParameter("@Code", adInteger, 4, EnumAccountCodes.Foroosh)
        Set Rctemp = RunParametricStoredProcedure2Rec("Get_tblAcc_Sale_ByCode", Parameter2)
        With VsSanadView
        While Rst.EOF <> True
                If Rctemp.EOF <> True And Rctemp.BOF <> True Then
                    .Rows = .Rows + 1
                    i = i + 1
                    .TextMatrix(.Rows - 1, 0) = i
                    .TextMatrix(.Rows - 1, 1) = Rctemp.Fields("Kol").Value
                    .TextMatrix(.Rows - 1, 2) = Rctemp.Fields("Moein").Value
                    .TextMatrix(.Rows - 1, 3) = Rst.Fields("Tafsili").Value  'inventoryName
                    .TextMatrix(.Rows - 1, 4) = Rctemp.Fields("Description").Value & "-" & Rst.Fields("Date").Value & " -- " & Rst.Fields("inventoryName").Value
                    .TextMatrix(.Rows - 1, 5) = 0
                    .TextMatrix(.Rows - 1, 6) = Rst.Fields("SumPriceTotal").Value
                End If
                Rst.MoveNext
        Wend
        End With
    End If
    ReDim Parameter(3) As Parameter
    Parameter(0) = GenerateInputParameter("@Branch", adInteger, 4, cboBranch.ItemData(cboBranch.ListIndex))
    Parameter(1) = GenerateInputParameter("@DateBefore", adVarWChar, 8, CStr(IIf(Trim(txtDate1.ClipText) = "", "", Trim(txtDate1.Text))))
    Parameter(2) = GenerateInputParameter("@DateAfter", adVarWChar, 8, CStr(IIf(Trim(txtDate2.ClipText) = "", "", Trim(txtDate2.Text))))
    Parameter(3) = GenerateInputParameter("@Uid", adInteger, 4, cmbUsers.ItemData(cmbUsers.ListIndex))
    Set Rst = RunParametricStoredProcedure2Rec("Get_SaleSummary_Added", Parameter)

    

    If Rst.EOF <> True And Rst.BOF <> True Then
        ReDim Parameter2(0) As Parameter
        Parameter2(0) = GenerateInputParameter("@Code", adInteger, 4, EnumAccountCodes.TakhfifateForoosh)
        Set Rctemp = RunParametricStoredProcedure2Rec("Get_tblAcc_Sale_ByCode", Parameter2)
        
        While Rst.EOF <> True
    
            If Rst!SumDiscount <> 0 Then

        '''''''''''     Œ›Ì›«  ›—Ê‘
                With VsSanadView
                If Rctemp.EOF <> True And Rctemp.BOF <> True Then
                    If Rst.Fields("SumDiscount").Value <> 0 Then
                        .Rows = .Rows + 1
                        i = i + 1
                        .TextMatrix(.Rows - 1, 0) = i
                        .TextMatrix(.Rows - 1, 1) = Rctemp.Fields("Kol").Value
                        .TextMatrix(.Rows - 1, 2) = Rctemp.Fields("Moein").Value
                        .TextMatrix(.Rows - 1, 3) = IIf(IsNull(Rctemp.Fields("Tafsili").Value), 0, Rctemp.Fields("Tafsili").Value)
                        .TextMatrix(.Rows - 1, 4) = Rctemp.Fields("Description").Value & " -- " & Rst.Fields("Date").Value
    '                        If Rst.Fields("SumDiscount").Value > 0 Then   '' Negative Discount
                            .TextMatrix(.Rows - 1, 5) = Rst.Fields("SumDiscount").Value
                            .TextMatrix(.Rows - 1, 6) = 0
    '                        Else
    '                            .TextMatrix(.Rows - 1, 5) = 0
    '                            .TextMatrix(.Rows - 1, 6) = Rst.Fields("SumDiscount").Value * (-1)
    '                        End If
                    End If
               End If
             End With
        End If
        Rst.MoveNext
     Wend
    End If
    ReDim Parameter(3) As Parameter
    Parameter(0) = GenerateInputParameter("@Branch", adInteger, 4, cboBranch.ItemData(cboBranch.ListIndex))
    Parameter(1) = GenerateInputParameter("@DateBefore", adVarWChar, 8, CStr(IIf(Trim(txtDate1.ClipText) = "", "", Trim(txtDate1.Text))))
    Parameter(2) = GenerateInputParameter("@DateAfter", adVarWChar, 8, CStr(IIf(Trim(txtDate2.ClipText) = "", "", Trim(txtDate2.Text))))
    Parameter(3) = GenerateInputParameter("@Uid", adInteger, 4, cmbUsers.ItemData(cmbUsers.ListIndex))
    Set Rst = RunParametricStoredProcedure2Rec("Get_SaleSummary_Added", Parameter)

    If Rst.EOF <> True And Rst.BOF <> True Then
        
        ReDim Parameter2(0) As Parameter
        Parameter2(0) = GenerateInputParameter("@Code", adInteger, 4, EnumAccountCodes.packing)
        Set Rctemp = RunParametricStoredProcedure2Rec("Get_tblAcc_Sale_ByCode", Parameter2)
        
        While Rst.EOF <> True

            If Rst!sumPacking <> 0 Then

            ''''''''''' »” Â »‰œÌ
                With VsSanadView  ' '
                    If Rctemp.EOF <> True And Rctemp.BOF <> True Then
                        If Rst.Fields("SumPacking").Value <> 0 Then
                            .Rows = .Rows + 1
                            i = i + 1
                            .TextMatrix(.Rows - 1, 0) = i
                            .TextMatrix(.Rows - 1, 1) = Rctemp.Fields("Kol").Value
                            .TextMatrix(.Rows - 1, 2) = Rctemp.Fields("Moein").Value
                            .TextMatrix(.Rows - 1, 3) = IIf(IsNull(Rctemp.Fields("Tafsili").Value), 0, Rctemp.Fields("Tafsili").Value)
                            .TextMatrix(.Rows - 1, 4) = Rctemp.Fields("Description").Value & " -- " & Rst.Fields("Date").Value
                            .TextMatrix(.Rows - 1, 5) = 0
                            .TextMatrix(.Rows - 1, 6) = Rst.Fields("SumPacking").Value
                        End If
                    End If
                End With
            End If
            Rst.MoveNext
        Wend
     End If
        '''''''''' ﬂ—«ÌÂ Õ„·
    ReDim Parameter(3) As Parameter
    Parameter(0) = GenerateInputParameter("@Branch", adInteger, 4, cboBranch.ItemData(cboBranch.ListIndex))
    Parameter(1) = GenerateInputParameter("@DateBefore", adVarWChar, 8, CStr(IIf(Trim(txtDate1.ClipText) = "", "", Trim(txtDate1.Text))))
    Parameter(2) = GenerateInputParameter("@DateAfter", adVarWChar, 8, CStr(IIf(Trim(txtDate2.ClipText) = "", "", Trim(txtDate2.Text))))
    Parameter(3) = GenerateInputParameter("@Uid", adInteger, 4, cmbUsers.ItemData(cmbUsers.ListIndex))
    Set Rst = RunParametricStoredProcedure2Rec("Get_SaleSummary_Added", Parameter)

    If Rst.EOF <> True And Rst.BOF <> True Then
        ReDim Parameter2(0) As Parameter
        Parameter2(0) = GenerateInputParameter("@Code", adInteger, 4, EnumAccountCodes.carryfee)
        Set Rctemp = RunParametricStoredProcedure2Rec("Get_tblAcc_Sale_ByCode", Parameter2)
        
        While Rst.EOF <> True
            If Rst!sumCarryFee <> 0 Then
                With VsSanadView
                    If Rctemp.EOF <> True And Rctemp.BOF <> True Then
                        If Rst.Fields("SumCarryFee").Value <> 0 Then
                            .Rows = .Rows + 1
                            i = i + 1
                            .TextMatrix(.Rows - 1, 0) = i
                            .TextMatrix(.Rows - 1, 1) = Rctemp.Fields("Kol").Value
                            .TextMatrix(.Rows - 1, 2) = Rctemp.Fields("Moein").Value
                            .TextMatrix(.Rows - 1, 3) = IIf(IsNull(Rctemp.Fields("Tafsili").Value), 0, Rctemp.Fields("Tafsili").Value)
                            .TextMatrix(.Rows - 1, 4) = Rctemp.Fields("Description").Value & " -- " & Rst.Fields("Date").Value
                            .TextMatrix(.Rows - 1, 5) = 0
                            .TextMatrix(.Rows - 1, 6) = Rst.Fields("SumCarryFee").Value
                        End If
                     End If
               End With
            End If
            Rst.MoveNext
        Wend
    End If
    ReDim Parameter(3) As Parameter
    Parameter(0) = GenerateInputParameter("@Branch", adInteger, 4, cboBranch.ItemData(cboBranch.ListIndex))
    Parameter(1) = GenerateInputParameter("@DateBefore", adVarWChar, 8, CStr(IIf(Trim(txtDate1.ClipText) = "", "", Trim(txtDate1.Text))))
    Parameter(2) = GenerateInputParameter("@DateAfter", adVarWChar, 8, CStr(IIf(Trim(txtDate2.ClipText) = "", "", Trim(txtDate2.Text))))
    Parameter(3) = GenerateInputParameter("@Uid", adInteger, 4, cmbUsers.ItemData(cmbUsers.ListIndex))
    Set Rst = RunParametricStoredProcedure2Rec("Get_SaleSummary_Added", Parameter)

    If Rst.EOF <> True And Rst.BOF <> True Then
        ReDim Parameter2(0) As Parameter
        Parameter2(0) = GenerateInputParameter("@Code", adInteger, 4, EnumAccountCodes.service)
        Set Rctemp = RunParametricStoredProcedure2Rec("Get_tblAcc_Sale_ByCode", Parameter2)

        While Rst.EOF <> True
            If Rst!sumService <> 0 Then
        ''''''''''' ”—ÊÌ”
                With VsSanadView
                    If Rctemp.EOF <> True And Rctemp.BOF <> True Then
                        If Rst.Fields("SumService").Value <> 0 Then
                            .Rows = .Rows + 1
                            i = i + 1
                            .TextMatrix(.Rows - 1, 0) = i
                            .TextMatrix(.Rows - 1, 1) = Rctemp.Fields("Kol").Value
                            .TextMatrix(.Rows - 1, 2) = Rctemp.Fields("Moein").Value
                            .TextMatrix(.Rows - 1, 3) = IIf(IsNull(Rctemp.Fields("Tafsili").Value), 0, Rctemp.Fields("Tafsili").Value)
                            .TextMatrix(.Rows - 1, 4) = Rctemp.Fields("Description").Value & " -- " & Rst.Fields("Date").Value
                            .TextMatrix(.Rows - 1, 5) = 0
                            .TextMatrix(.Rows - 1, 6) = Rst.Fields("SumService").Value
                        End If
                    End If
                End With
            End If
            Rst.MoveNext
        Wend
    End If
'
    ' Duty
    ReDim Parameter(3) As Parameter
    Parameter(0) = GenerateInputParameter("@Branch", adInteger, 4, cboBranch.ItemData(cboBranch.ListIndex))
    Parameter(1) = GenerateInputParameter("@DateBefore", adVarWChar, 8, CStr(IIf(Trim(txtDate1.ClipText) = "", "", Trim(txtDate1.Text))))
    Parameter(2) = GenerateInputParameter("@DateAfter", adVarWChar, 8, CStr(IIf(Trim(txtDate2.ClipText) = "", "", Trim(txtDate2.Text))))
    Parameter(3) = GenerateInputParameter("@Uid", adInteger, 4, cmbUsers.ItemData(cmbUsers.ListIndex))
    Set Rst = RunParametricStoredProcedure2Rec("Get_SaleSummary_Added", Parameter)
            
    If Rst.EOF <> True And Rst.BOF <> True Then
        ReDim Parameter2(0) As Parameter
        Parameter2(0) = GenerateInputParameter("@Code", adInteger, 4, EnumAccountCodes.TaxSale)
        Set Rctemp = RunParametricStoredProcedure2Rec("Get_tblAcc_Sale_ByCode", Parameter2)

        While Rst.EOF <> True
        
        ''''''''''' ⁄Ê«—÷
            If Rst.Fields("DutyTotal").Value <> 0 Then
                With VsSanadView
                    If Rctemp.EOF <> True And Rctemp.BOF <> True Then
                        If Rst.Fields("DutyTotal").Value <> 0 Then
                            .Rows = .Rows + 1
                            i = i + 1
                            .TextMatrix(.Rows - 1, 0) = i
                            .TextMatrix(.Rows - 1, 1) = Rctemp.Fields("Kol").Value
                            .TextMatrix(.Rows - 1, 2) = Rctemp.Fields("Moein").Value
                            .TextMatrix(.Rows - 1, 3) = IIf(IsNull(Rctemp.Fields("Tafsili").Value), 0, Rctemp.Fields("Tafsili").Value)
                            .TextMatrix(.Rows - 1, 4) = Rctemp.Fields("Description").Value & " -- " & Rst.Fields("Date").Value
                            .TextMatrix(.Rows - 1, 5) = 0
                            .TextMatrix(.Rows - 1, 6) = Rst.Fields("DutyTotal").Value
                        End If
                    End If
                End With
            End If
            Rst.MoveNext
        Wend
    End If
    ' Tax
    ReDim Parameter(3) As Parameter
    Parameter(0) = GenerateInputParameter("@Branch", adInteger, 4, cboBranch.ItemData(cboBranch.ListIndex))
    Parameter(1) = GenerateInputParameter("@DateBefore", adVarWChar, 8, CStr(IIf(Trim(txtDate1.ClipText) = "", "", Trim(txtDate1.Text))))
    Parameter(2) = GenerateInputParameter("@DateAfter", adVarWChar, 8, CStr(IIf(Trim(txtDate2.ClipText) = "", "", Trim(txtDate2.Text))))
    Parameter(3) = GenerateInputParameter("@Uid", adInteger, 4, cmbUsers.ItemData(cmbUsers.ListIndex))
    Set Rst = RunParametricStoredProcedure2Rec("Get_SaleSummary_Added", Parameter)
            
    If Rst.EOF <> True And Rst.BOF <> True Then
        ReDim Parameter2(0) As Parameter
        Parameter2(0) = GenerateInputParameter("@Code", adInteger, 4, EnumAccountCodes.DutySale)
        Set Rctemp = RunParametricStoredProcedure2Rec("Get_tblAcc_Sale_ByCode", Parameter2)

        While Rst.EOF <> True
            If Rst.Fields("TaxTotal").Value <> 0 Then
        ''''''''''' ”—ÊÌ”
                With VsSanadView
                    If Rctemp.EOF <> True And Rctemp.BOF <> True Then
                        If Rst.Fields("TaxTotal").Value <> 0 Then
                            .Rows = .Rows + 1
                            i = i + 1
                            .TextMatrix(.Rows - 1, 0) = i
                            .TextMatrix(.Rows - 1, 1) = Rctemp.Fields("Kol").Value
                            .TextMatrix(.Rows - 1, 2) = Rctemp.Fields("Moein").Value
                            .TextMatrix(.Rows - 1, 3) = IIf(IsNull(Rctemp.Fields("Tafsili").Value), 0, Rctemp.Fields("Tafsili").Value)
                            .TextMatrix(.Rows - 1, 4) = Rctemp.Fields("Description").Value & " -- " & Rst.Fields("Date").Value
                            .TextMatrix(.Rows - 1, 5) = 0
                            .TextMatrix(.Rows - 1, 6) = Rst.Fields("TaxTotal").Value
                        End If
                    End If
                End With
            End If
            Rst.MoveNext
        Wend
    End If

    Rst.Close
    Set Rst = Nothing
    Call SaleReturnSummary
    Call Daryaft
    Call UserCustomerRecieve
    Call UserCustomerPayment
    Call UserSupplierRecieve
    Call UserSupplierPayment
    Call UserExpensivePayment
    
    
End Sub
Private Sub SaleSummaryCustom()
' Sale Summary
    
    ReDim Parameter(3) As Parameter
    Parameter(0) = GenerateInputParameter("@Branch", adInteger, 4, cboBranch.ItemData(cboBranch.ListIndex))
    Parameter(1) = GenerateInputParameter("@DateBefore", adVarWChar, 8, CStr(IIf(Trim(txtDate1.ClipText) = "", "", Trim(txtDate1.Text))))
    Parameter(2) = GenerateInputParameter("@DateAfter", adVarWChar, 8, CStr(IIf(Trim(txtDate2.ClipText) = "", "", Trim(txtDate2.Text))))
    Parameter(3) = GenerateInputParameter("@Uid", adInteger, 4, cmbUsers.ItemData(cmbUsers.ListIndex))
    Set Rst = RunParametricStoredProcedure2Rec("Get_SaleSummaryCustom", Parameter)

'''''''''''   ›—Ê‘


    If Rst.EOF <> True And Rst.BOF <> True Then
        With VsSanadView
        While Rst.EOF <> True
            If Rst.Fields("SumBedehKar").Value <> 0 Or Rst.Fields("SumBestankar").Value Then
                .Rows = .Rows + 1
                i = i + 1
                .TextMatrix(.Rows - 1, 0) = i
                .TextMatrix(.Rows - 1, 1) = Rst.Fields("Kol").Value
                .TextMatrix(.Rows - 1, 2) = Rst.Fields("Moein").Value
                .TextMatrix(.Rows - 1, 3) = IIf(IsNull(Rst.Fields("Tafsili").Value), 0, Rst.Fields("Tafsili").Value)    'inventoryName
                .TextMatrix(.Rows - 1, 4) = Rst.Fields("Name").Value
                .TextMatrix(.Rows - 1, 5) = Rst.Fields("SumBedehKar").Value
                .TextMatrix(.Rows - 1, 6) = Rst.Fields("SumBestankar").Value
            End If
            Rst.MoveNext
        Wend
        End With
    End If
    
End Sub

Private Sub Daryaft()
    
    ReDim Parameter(4) As Parameter

    Parameter(0) = GenerateInputParameter("@Branch", adInteger, 4, cboBranch.ItemData(cboBranch.ListIndex))
    Parameter(1) = GenerateInputParameter("@DateBefore", adVarWChar, 8, CStr(IIf(Trim(txtDate1.ClipText) = "", "", Trim(txtDate1.Text))))
    Parameter(2) = GenerateInputParameter("@DateAfter", adVarWChar, 8, CStr(IIf(Trim(txtDate2.ClipText) = "", "", Trim(txtDate2.Text))))
    Parameter(3) = GenerateInputParameter("@Code", adInteger, 4, EnumAccountType.Cash)
    Parameter(4) = GenerateInputParameter("@Uid", adInteger, 4, cmbUsers.ItemData(cmbUsers.ListIndex))
    Set Rst = RunParametricStoredProcedure2Rec("Get_AccountDocument", Parameter)

    If Rst.EOF <> True And Rst.BOF <> True Then
        While Rst.EOF <> True
            If Rst.Fields("sp").Value > 0 Then
                ReDim Parameter2(0) As Parameter
                Parameter2(0) = GenerateInputParameter("@Code", adInteger, 4, EnumAccountCodes.CashRemains)
                Set Rctemp = RunParametricStoredProcedure2Rec("Get_tblAcc_Sale_ByCode", Parameter2)
                With VsSanadView
                    If Rctemp.EOF <> True And Rctemp.BOF <> True Then
                        .Rows = .Rows + 1
                        i = i + 1
                        .TextMatrix(.Rows - 1, 0) = i
                        .TextMatrix(.Rows - 1, 1) = Rctemp.Fields("Kol").Value
                        .TextMatrix(.Rows - 1, 2) = Rctemp.Fields("Moein").Value
                        .TextMatrix(.Rows - 1, 3) = IIf(IsNull(Rst.Fields("PersonTafsili").Value), "", Rst.Fields("PersonTafsili").Value)
                        .TextMatrix(.Rows - 1, 4) = Rctemp.Fields("Description").Value & " -- " & Rst.Fields("UserFullName").Value & " -- " & Rst.Fields("Date").Value
                        .TextMatrix(.Rows - 1, 5) = Rst.Fields("sp").Value
                        .TextMatrix(.Rows - 1, 6) = "0"
                    End If
                End With
            End If
            Rst.MoveNext
        Wend
    End If
    Rst.Close

    Parameter(3) = GenerateInputParameter("@Code", adInteger, 4, EnumAccountType.Card)  ' Card
    Set Rst = RunParametricStoredProcedure2Rec("Get_AccountDocument", Parameter)

    If Rst.EOF <> True And Rst.BOF <> True Then
        While Rst.EOF <> True
            ReDim Parameter2(0) As Parameter
            Parameter2(0) = GenerateInputParameter("@Code", adInteger, 4, EnumAccountCodes.Banks)
            Set Rctemp = RunParametricStoredProcedure2Rec("Get_tblAcc_Sale_ByCode", Parameter2)
            With VsSanadView
                If Rctemp.EOF <> True And Rctemp.BOF <> True Then
                    If Val(Rst.Fields("TafsiliId").Value) = 0 Then ShowMessage " Õ”«» »«‰òÌ »—«Ì ÅÊ“ »«‰òÌ œ——œÌ›   " & i + 1 & "   ⁄—Ì› ‰‘œÂ. „»·€ ’›— „‰ŸÊ— „Ì ê—œœ  « œ— ›—„ ÅÊ“ »«‰òÌ œ—”  ê—œœ ", True, False, "ﬁ»Ê·", ""
                    .Rows = .Rows + 1
                    i = i + 1
                    .TextMatrix(.Rows - 1, 0) = i
                    .TextMatrix(.Rows - 1, 1) = Rctemp.Fields("Kol").Value
                    .TextMatrix(.Rows - 1, 2) = Rctemp.Fields("Moein").Value
                    .TextMatrix(.Rows - 1, 3) = IIf(IsNull(Rst.Fields("TafsiliId").Value), "", Rst.Fields("TafsiliId").Value)
                    .TextMatrix(.Rows - 1, 4) = Rst.Fields("nvcDescription").Value & " Ê«—Ì“ «“ ÿ—Ìﬁ ò«—   " & " -- " & Rst.Fields("Date").Value
                    If Val(Rst.Fields("TafsiliId").Value) <> 0 Then .TextMatrix(.Rows - 1, 5) = Rst.Fields("sp").Value Else .TextMatrix(.Rows - 1, 5) = 0
                    .TextMatrix(.Rows - 1, 6) = "0"
                End If
            End With
            Rst.MoveNext
        Wend
    End If
    Rst.Close

End Sub

Private Sub SaleGarsoons()
        ReDim Parameter(1) As Parameter
        Parameter(0) = GenerateInputParameter("@DateBefore", adVarWChar, 8, CStr(IIf(Trim(txtDate1.ClipText) = "", "", Trim(txtDate1.Text))))
        Parameter(1) = GenerateInputParameter("@DateAfter", adVarWChar, 8, CStr(IIf(Trim(txtDate2.ClipText) = "", "", Trim(txtDate2.Text))))
        Set Rst = RunParametricStoredProcedure2Rec("Get_GarsonBillPaymentSummary", Parameter)
        
        If Rst.EOF <> True And Rst.BOF <> True Then
        
            While Rst.EOF <> True
               
                ReDim Parameter2(0) As Parameter
                Parameter2(0) = GenerateInputParameter("@Code", adInteger, 4, 8)
                Set Rctemp = RunParametricStoredProcedure2Rec("Get_tblAcc_Sale_ByCode", Parameter2)
                 With VsSanadView
                    
                        If Rctemp.EOF <> True And Rctemp.BOF <> True Then
                            .Rows = .Rows + 1
                            i = i + 1
                            .TextMatrix(.Rows - 1, 0) = i
                            .TextMatrix(.Rows - 1, 1) = Rctemp.Fields("Kol").Value
                            .TextMatrix(.Rows - 1, 2) = Rctemp.Fields("Moein").Value
                            .TextMatrix(.Rows - 1, 3) = IIf(IsNull(Rst.Fields("Tafsili").Value), "", Rst.Fields("Tafsili").Value)
                            .TextMatrix(.Rows - 1, 4) = Rctemp.Fields("Description").Value & " -- " & Rst.Fields("FullName").Value & " -- " & Rst.Fields("Date").Value
                            .TextMatrix(.Rows - 1, 5) = Rst.Fields("SumPrice").Value
                            .TextMatrix(.Rows - 1, 6) = 0
                        End If
                End With
            Rst.MoveNext
            Wend
        
        End If

End Sub
Private Sub SaleCouriers()
     'Courier BillPayment Summary
    
        ReDim Parameter(1) As Parameter
        Parameter(0) = GenerateInputParameter("@DateBefore", adVarWChar, 8, CStr(IIf(Trim(txtDate1.ClipText) = "", "", Trim(txtDate1.Text))))
        Parameter(1) = GenerateInputParameter("@DateAfter", adVarWChar, 8, CStr(IIf(Trim(txtDate2.ClipText) = "", "", Trim(txtDate2.Text))))
        Set Rst = RunParametricStoredProcedure2Rec("Get_CarrierBillPaymentSummary", Parameter)
        
        If Rst.EOF <> True And Rst.BOF <> True Then
        
            While Rst.EOF <> True
               
                ReDim Parameter2(0) As Parameter
                Parameter2(0) = GenerateInputParameter("@Code", adInteger, 4, 9)
                Set Rctemp = RunParametricStoredProcedure2Rec("Get_tblAcc_Sale_ByCode", Parameter2)
                 With VsSanadView
                    
                        If Rctemp.EOF <> True And Rctemp.BOF <> True Then
                            .Rows = .Rows + 1
                            i = i + 1
                            .TextMatrix(.Rows - 1, 0) = i
                            .TextMatrix(.Rows - 1, 1) = Rctemp.Fields("Kol").Value
                            .TextMatrix(.Rows - 1, 2) = Rctemp.Fields("Moein").Value
                            .TextMatrix(.Rows - 1, 3) = IIf(IsNull(Rst.Fields("Tafsili").Value), "", Rst.Fields("Tafsili").Value)
                            .TextMatrix(.Rows - 1, 4) = Rctemp.Fields("Description").Value & " -- " & Rst.Fields("CarrierFullName").Value & " -- " & Rst.Fields("Date").Value
                            .TextMatrix(.Rows - 1, 5) = Rst.Fields("SumPrice").Value
                            .TextMatrix(.Rows - 1, 6) = 0
                        End If
                End With
            Rst.MoveNext
            Wend
        
        End If
    
End Sub
Private Sub UserExpensivePayment()
'User Payment Summary  ( Å—œ«Œ  Â“Ì‰Â Â« )
  Dim Uid_Temp As Integer
  Dim Uid_Desc, Uid_Tafsili As String
  Totalprice = 0
  
  Uid_Temp = 0
    
With VsSanadView
    ReDim Parameter(3) As Parameter
    Parameter(0) = GenerateInputParameter("@PaymentType", adInteger, 4, EnumPaymentType.Expensive)
    Parameter(1) = GenerateInputParameter("@DateBefore", adVarWChar, 8, CStr(IIf(Trim(txtDate1.ClipText) = "", "", Trim(txtDate1.Text))))
    Parameter(2) = GenerateInputParameter("@DateAfter", adVarWChar, 8, CStr(IIf(Trim(txtDate2.ClipText) = "", "", Trim(txtDate2.Text))))
    Parameter(3) = GenerateInputParameter("@Uid", adInteger, 4, cmbUsers.ItemData(cmbUsers.ListIndex))
    Set Rst = RunParametricStoredProcedure2Rec("Get_UserPayment", Parameter)
    
    If Rst.EOF <> True And Rst.BOF <> True Then
        Uid_Temp = Rst!Uid
        While Rst.EOF <> True
   
            Uid_Desc = "Å—œ«Œ  »«»  Â“Ì‰Â Â«" & " -- " & Rst.Fields("User_Name").Value & " -- " & Rst.Fields("Date").Value
            Uid_Tafsili = IIf(IsNull(Rst.Fields("Tafsili").Value), "", Rst.Fields("Tafsili").Value)
            
            ReDim Parameter2(0) As Parameter
            Parameter2(0) = GenerateInputParameter("@Code", adInteger, 4, EnumAccountCodes.Hazineh)
            Set Rctemp = RunParametricStoredProcedure2Rec("Get_tblAcc_Sale_ByCode", Parameter2)
            If Rctemp.EOF <> True And Rctemp.BOF <> True Then
                .Rows = .Rows + 1
                i = i + 1
                .TextMatrix(.Rows - 1, 0) = i
                .TextMatrix(.Rows - 1, 1) = Rctemp.Fields("Kol").Value
                .TextMatrix(.Rows - 1, 2) = Rctemp.Fields("Moein").Value
                .TextMatrix(.Rows - 1, 3) = IIf(IsNull(Rctemp.Fields("Tafsili").Value), "", Rctemp.Fields("Tafsili").Value)
                .TextMatrix(.Rows - 1, 4) = " Â“Ì‰Â Â«Ì ⁄„Ê„Ì " & " -- " & Rst.Fields("Person_Name").Value & " -- " & Rst.Fields("Date").Value
                .TextMatrix(.Rows - 1, 5) = Rst.Fields("TotalBestankar").Value
                .TextMatrix(.Rows - 1, 6) = 0
                Totalprice = Totalprice + Rst.Fields("TotalBestankar").Value
            End If
            
            Rst.MoveNext
            If Rst.EOF <> True And Rst.BOF <> True Then
                 If Uid_Temp <> Rst!Uid Then
                     ReDim Parameter2(0) As Parameter
                     Parameter2(0) = GenerateInputParameter("@Code", adInteger, 4, 7)
                     Set Rctemp = RunParametricStoredProcedure2Rec("Get_tblAcc_Sale_ByCode", Parameter2)
                     
                     If Rctemp.EOF <> True And Rctemp.BOF <> True Then
                         .Rows = .Rows + 1
                         i = i + 1
                         .TextMatrix(.Rows - 1, 0) = i
                         .TextMatrix(.Rows - 1, 1) = Rctemp.Fields("Kol").Value
                         .TextMatrix(.Rows - 1, 2) = Rctemp.Fields("Moein").Value
                         .TextMatrix(.Rows - 1, 3) = Uid_Tafsili
                         .TextMatrix(.Rows - 1, 4) = Uid_Desc
                         .TextMatrix(.Rows - 1, 5) = 0
                         .TextMatrix(.Rows - 1, 6) = Totalprice
                         Totalprice = 0
                         Uid_Temp = Rst!Uid
                     End If
                  End If
            Else
                 ReDim Parameter2(0) As Parameter
                 Parameter2(0) = GenerateInputParameter("@Code", adInteger, 4, 7)
                 Set Rctemp = RunParametricStoredProcedure2Rec("Get_tblAcc_Sale_ByCode", Parameter2)
                 If Rctemp.EOF <> True And Rctemp.BOF <> True Then
                     .Rows = .Rows + 1
                     i = i + 1
                     .TextMatrix(.Rows - 1, 0) = i
                     .TextMatrix(.Rows - 1, 1) = Rctemp.Fields("Kol").Value
                     .TextMatrix(.Rows - 1, 2) = Rctemp.Fields("Moein").Value
                     .TextMatrix(.Rows - 1, 3) = Uid_Tafsili
                     .TextMatrix(.Rows - 1, 4) = Uid_Desc
                     .TextMatrix(.Rows - 1, 5) = 0
                     .TextMatrix(.Rows - 1, 6) = Totalprice
                     Totalprice = 0
                 End If
            End If
       Wend
    End If
End With

End Sub
Private Sub UserPersonPayment()
'User Payment Summary  ( Å—œ«Œ  »Â Å—”‰· )
  Totalprice = 0
  Dim Uid_Temp As Integer
  Dim Uid_Desc, Uid_Tafsili As String
  Totalprice = 0
  Uid_Temp = 0
  With VsSanadView

    ReDim Parameter(3) As Parameter
    Parameter(0) = GenerateInputParameter("@PaymentType", adInteger, 4, EnumPaymentType.TempPersonPayment)
    Parameter(1) = GenerateInputParameter("@DateBefore", adVarWChar, 8, CStr(IIf(Trim(txtDate1.ClipText) = "", "", Trim(txtDate1.Text))))
    Parameter(2) = GenerateInputParameter("@DateAfter", adVarWChar, 8, CStr(IIf(Trim(txtDate2.ClipText) = "", "", Trim(txtDate2.Text))))
    Parameter(3) = GenerateInputParameter("@Uid", adInteger, 4, cmbUsers.ItemData(cmbUsers.ListIndex))
    Set Rst = RunParametricStoredProcedure2Rec("Get_UserPayment", Parameter)
    
    If Rst.EOF <> True And Rst.BOF <> True Then
    
        Uid_Temp = Rst!Uid
        While Rst.EOF <> True
           
            Uid_Desc = "⁄·Ì «·Õ”«» - Å—œ«Œ  »Â Å—”‰· " & " -- " & Rst.Fields("User_Name").Value & " -- " & Rst.Fields("Date").Value
            Uid_Tafsili = IIf(IsNull(Rst.Fields("Tafsili").Value), "", Rst.Fields("Tafsili").Value)
            
            ReDim Parameter2(0) As Parameter
            Parameter2(0) = GenerateInputParameter("@Code", adInteger, 4, 10)
            Set Rctemp = RunParametricStoredProcedure2Rec("Get_tblAcc_Sale_ByCode", Parameter2)
            If Rctemp.EOF <> True And Rctemp.BOF <> True Then
                .Rows = .Rows + 1
                i = i + 1
                .TextMatrix(.Rows - 1, 0) = i
                .TextMatrix(.Rows - 1, 1) = Rctemp.Fields("Kol").Value
                .TextMatrix(.Rows - 1, 2) = Rctemp.Fields("Moein").Value
                .TextMatrix(.Rows - 1, 3) = IIf(IsNull(Rst.Fields("Person_Tafsili").Value), "", Rst.Fields("Person_Tafsili").Value)
                .TextMatrix(.Rows - 1, 4) = "⁄·Ì «·Õ”«» -œ—Ì«›  «“ ’‰œÊﬁ" & " -- " & Rst.Fields("Person_Name").Value & " -- " & Rst.Fields("Date").Value
                .TextMatrix(.Rows - 1, 5) = Rst.Fields("TotalBestankar").Value
                .TextMatrix(.Rows - 1, 6) = 0
                Totalprice = Totalprice + Rst.Fields("TotalBestankar").Value
                        
            End If
            Rst.MoveNext
            If Rst.EOF <> True And Rst.BOF <> True Then
                 If Uid_Temp <> Rst!Uid Then
                     ReDim Parameter2(0) As Parameter
                     Parameter2(0) = GenerateInputParameter("@Code", adInteger, 4, 7)
                     Set Rctemp = RunParametricStoredProcedure2Rec("Get_tblAcc_Sale_ByCode", Parameter2)
                     
                     If Rctemp.EOF <> True And Rctemp.BOF <> True Then
                         .Rows = .Rows + 1
                         i = i + 1
                         .TextMatrix(.Rows - 1, 0) = i
                         .TextMatrix(.Rows - 1, 1) = Rctemp.Fields("Kol").Value
                         .TextMatrix(.Rows - 1, 2) = Rctemp.Fields("Moein").Value
                         .TextMatrix(.Rows - 1, 3) = Uid_Tafsili
                         .TextMatrix(.Rows - 1, 4) = Uid_Desc
                         .TextMatrix(.Rows - 1, 5) = 0
                         .TextMatrix(.Rows - 1, 6) = Totalprice
                         Totalprice = 0
                         Uid_Temp = Rst!Uid
                     End If
                  End If
            Else
                 ReDim Parameter2(0) As Parameter
                 Parameter2(0) = GenerateInputParameter("@Code", adInteger, 4, 7)
                 Set Rctemp = RunParametricStoredProcedure2Rec("Get_tblAcc_Sale_ByCode", Parameter2)
                 If Rctemp.EOF <> True And Rctemp.BOF <> True Then
                     .Rows = .Rows + 1
                     i = i + 1
                     .TextMatrix(.Rows - 1, 0) = i
                     .TextMatrix(.Rows - 1, 1) = Rctemp.Fields("Kol").Value
                     .TextMatrix(.Rows - 1, 2) = Rctemp.Fields("Moein").Value
                     .TextMatrix(.Rows - 1, 3) = Uid_Tafsili
                     .TextMatrix(.Rows - 1, 4) = Uid_Desc
                     .TextMatrix(.Rows - 1, 5) = 0
                     .TextMatrix(.Rows - 1, 6) = Totalprice
                     Totalprice = 0
                 End If
            End If
       Wend
   
    End If
  End With
  With VsSanadView

    ReDim Parameter(3) As Parameter
    Parameter(0) = GenerateInputParameter("@PaymentType", adInteger, 4, EnumPaymentType.VamPersonPayment)
    Parameter(1) = GenerateInputParameter("@DateBefore", adVarWChar, 8, CStr(IIf(Trim(txtDate1.ClipText) = "", "", Trim(txtDate1.Text))))
    Parameter(2) = GenerateInputParameter("@DateAfter", adVarWChar, 8, CStr(IIf(Trim(txtDate2.ClipText) = "", "", Trim(txtDate2.Text))))
    Parameter(3) = GenerateInputParameter("@Uid", adInteger, 4, cmbUsers.ItemData(cmbUsers.ListIndex))
    Set Rst = RunParametricStoredProcedure2Rec("Get_UserPayment", Parameter)
    
    If Rst.EOF <> True And Rst.BOF <> True Then
    
        Uid_Temp = Rst!Uid
        While Rst.EOF <> True
           
            Uid_Desc = "Ê«„ -  Å—œ«Œ  »Â Å—”‰· " & " -- " & Rst.Fields("User_Name").Value & " -- " & Rst.Fields("Date").Value
            Uid_Tafsili = IIf(IsNull(Rst.Fields("Tafsili").Value), "", Rst.Fields("Tafsili").Value)
            
            ReDim Parameter2(0) As Parameter
            Parameter2(0) = GenerateInputParameter("@Code", adInteger, 4, 11)
            Set Rctemp = RunParametricStoredProcedure2Rec("Get_tblAcc_Sale_ByCode", Parameter2)
            If Rctemp.EOF <> True And Rctemp.BOF <> True Then
                .Rows = .Rows + 1
                i = i + 1
                .TextMatrix(.Rows - 1, 0) = i
                .TextMatrix(.Rows - 1, 1) = Rctemp.Fields("Kol").Value
                .TextMatrix(.Rows - 1, 2) = Rctemp.Fields("Moein").Value
                .TextMatrix(.Rows - 1, 3) = IIf(IsNull(Rst.Fields("Person_Tafsili").Value), "", Rst.Fields("Person_Tafsili").Value)
                .TextMatrix(.Rows - 1, 4) = "Ê«„ - œ—Ì«›  «“ ’‰œÊﬁ" & " -- " & Rst.Fields("Person_Name").Value & " -- " & Rst.Fields("Date").Value
                .TextMatrix(.Rows - 1, 5) = Rst.Fields("TotalBestankar").Value
                .TextMatrix(.Rows - 1, 6) = 0
                Totalprice = Totalprice + Rst.Fields("TotalBestankar").Value
                        
            End If
            Rst.MoveNext
            If Rst.EOF <> True And Rst.BOF <> True Then
                 If Uid_Temp <> Rst!Uid Then
                     ReDim Parameter2(0) As Parameter
                     Parameter2(0) = GenerateInputParameter("@Code", adInteger, 4, 7)
                     Set Rctemp = RunParametricStoredProcedure2Rec("Get_tblAcc_Sale_ByCode", Parameter2)
                     
                     If Rctemp.EOF <> True And Rctemp.BOF <> True Then
                         .Rows = .Rows + 1
                         i = i + 1
                         .TextMatrix(.Rows - 1, 0) = i
                         .TextMatrix(.Rows - 1, 1) = Rctemp.Fields("Kol").Value
                         .TextMatrix(.Rows - 1, 2) = Rctemp.Fields("Moein").Value
                         .TextMatrix(.Rows - 1, 3) = Uid_Tafsili
                         .TextMatrix(.Rows - 1, 4) = Uid_Desc
                         .TextMatrix(.Rows - 1, 5) = 0
                         .TextMatrix(.Rows - 1, 6) = Totalprice
                         Totalprice = 0
                         Uid_Temp = Rst!Uid
                     End If
                  End If
            Else
                 ReDim Parameter2(0) As Parameter
                 Parameter2(0) = GenerateInputParameter("@Code", adInteger, 4, 7)
                 Set Rctemp = RunParametricStoredProcedure2Rec("Get_tblAcc_Sale_ByCode", Parameter2)
                 If Rctemp.EOF <> True And Rctemp.BOF <> True Then
                     .Rows = .Rows + 1
                     i = i + 1
                     .TextMatrix(.Rows - 1, 0) = i
                     .TextMatrix(.Rows - 1, 1) = Rctemp.Fields("Kol").Value
                     .TextMatrix(.Rows - 1, 2) = Rctemp.Fields("Moein").Value
                     .TextMatrix(.Rows - 1, 3) = Uid_Tafsili
                     .TextMatrix(.Rows - 1, 4) = Uid_Desc
                     .TextMatrix(.Rows - 1, 5) = 0
                     .TextMatrix(.Rows - 1, 6) = Totalprice
                     Totalprice = 0
                 End If
            End If
       Wend
   
    End If
  End With
  With VsSanadView

    ReDim Parameter(3) As Parameter
    Parameter(0) = GenerateInputParameter("@PaymentType", adInteger, 4, EnumPaymentType.MosaedePersonPayment)
    Parameter(1) = GenerateInputParameter("@DateBefore", adVarWChar, 8, CStr(IIf(Trim(txtDate1.ClipText) = "", "", Trim(txtDate1.Text))))
    Parameter(2) = GenerateInputParameter("@DateAfter", adVarWChar, 8, CStr(IIf(Trim(txtDate2.ClipText) = "", "", Trim(txtDate2.Text))))
    Parameter(3) = GenerateInputParameter("@Uid", adInteger, 4, cmbUsers.ItemData(cmbUsers.ListIndex))
    Set Rst = RunParametricStoredProcedure2Rec("Get_UserPayment", Parameter)
    
    If Rst.EOF <> True And Rst.BOF <> True Then
    
        Uid_Temp = Rst!Uid
        While Rst.EOF <> True
           
            Uid_Desc = "„”«⁄œÂ - Å—œ«Œ  »Â Å—”‰· " & " -- " & Rst.Fields("User_Name").Value & " -- " & Rst.Fields("Date").Value
            Uid_Tafsili = IIf(IsNull(Rst.Fields("Tafsili").Value), "", Rst.Fields("Tafsili").Value)
            
            ReDim Parameter2(0) As Parameter
            Parameter2(0) = GenerateInputParameter("@Code", adInteger, 4, 12)
            Set Rctemp = RunParametricStoredProcedure2Rec("Get_tblAcc_Sale_ByCode", Parameter2)
            If Rctemp.EOF <> True And Rctemp.BOF <> True Then
                .Rows = .Rows + 1
                i = i + 1
                .TextMatrix(.Rows - 1, 0) = i
                .TextMatrix(.Rows - 1, 1) = Rctemp.Fields("Kol").Value
                .TextMatrix(.Rows - 1, 2) = Rctemp.Fields("Moein").Value
                .TextMatrix(.Rows - 1, 3) = IIf(IsNull(Rst.Fields("Person_Tafsili").Value), "", Rst.Fields("Person_Tafsili").Value)
                .TextMatrix(.Rows - 1, 4) = "„”«⁄œÂ -œ—Ì«›  «“ ’‰œÊﬁ" & " -- " & Rst.Fields("Person_Name").Value & " -- " & Rst.Fields("Date").Value
                .TextMatrix(.Rows - 1, 5) = Rst.Fields("TotalBestankar").Value
                .TextMatrix(.Rows - 1, 6) = 0
                Totalprice = Totalprice + Rst.Fields("TotalBestankar").Value
                        
            End If
            Rst.MoveNext
            If Rst.EOF <> True And Rst.BOF <> True Then
                 If Uid_Temp <> Rst!Uid Then
                     ReDim Parameter2(0) As Parameter
                     Parameter2(0) = GenerateInputParameter("@Code", adInteger, 4, 7)
                     Set Rctemp = RunParametricStoredProcedure2Rec("Get_tblAcc_Sale_ByCode", Parameter2)
                     
                     If Rctemp.EOF <> True And Rctemp.BOF <> True Then
                         .Rows = .Rows + 1
                         i = i + 1
                         .TextMatrix(.Rows - 1, 0) = i
                         .TextMatrix(.Rows - 1, 1) = Rctemp.Fields("Kol").Value
                         .TextMatrix(.Rows - 1, 2) = Rctemp.Fields("Moein").Value
                         .TextMatrix(.Rows - 1, 3) = Uid_Tafsili
                         .TextMatrix(.Rows - 1, 4) = Uid_Desc
                         .TextMatrix(.Rows - 1, 5) = 0
                         .TextMatrix(.Rows - 1, 6) = Totalprice
                         Totalprice = 0
                         Uid_Temp = Rst!Uid
                     End If
                  End If
            Else
                 ReDim Parameter2(0) As Parameter
                 Parameter2(0) = GenerateInputParameter("@Code", adInteger, 4, 7)
                 Set Rctemp = RunParametricStoredProcedure2Rec("Get_tblAcc_Sale_ByCode", Parameter2)
                 If Rctemp.EOF <> True And Rctemp.BOF <> True Then
                     .Rows = .Rows + 1
                     i = i + 1
                     .TextMatrix(.Rows - 1, 0) = i
                     .TextMatrix(.Rows - 1, 1) = Rctemp.Fields("Kol").Value
                     .TextMatrix(.Rows - 1, 2) = Rctemp.Fields("Moein").Value
                     .TextMatrix(.Rows - 1, 3) = Uid_Tafsili
                     .TextMatrix(.Rows - 1, 4) = Uid_Desc
                     .TextMatrix(.Rows - 1, 5) = 0
                     .TextMatrix(.Rows - 1, 6) = Totalprice
                     Totalprice = 0
                 End If
            End If
       Wend
   
    End If
  End With
  With VsSanadView

    ReDim Parameter(3) As Parameter
    Parameter(0) = GenerateInputParameter("@PaymentType", adInteger, 4, EnumPaymentType.SalaryPersonPayment)
    Parameter(1) = GenerateInputParameter("@DateBefore", adVarWChar, 8, CStr(IIf(Trim(txtDate1.ClipText) = "", "", Trim(txtDate1.Text))))
    Parameter(2) = GenerateInputParameter("@DateAfter", adVarWChar, 8, CStr(IIf(Trim(txtDate2.ClipText) = "", "", Trim(txtDate2.Text))))
    Parameter(3) = GenerateInputParameter("@Uid", adInteger, 4, cmbUsers.ItemData(cmbUsers.ListIndex))
    Set Rst = RunParametricStoredProcedure2Rec("Get_UserPayment", Parameter)
    
    If Rst.EOF <> True And Rst.BOF <> True Then
    
        Uid_Temp = Rst!Uid
        While Rst.EOF <> True
           
            Uid_Desc = "ÕﬁÊﬁ - Å—œ«Œ  »Â Å—”‰· " & " -- " & Rst.Fields("User_Name").Value & " -- " & Rst.Fields("Date").Value
            Uid_Tafsili = IIf(IsNull(Rst.Fields("Tafsili").Value), "", Rst.Fields("Tafsili").Value)
            
            ReDim Parameter2(0) As Parameter
            Parameter2(0) = GenerateInputParameter("@Code", adInteger, 4, 13)
            Set Rctemp = RunParametricStoredProcedure2Rec("Get_tblAcc_Sale_ByCode", Parameter2)
            If Rctemp.EOF <> True And Rctemp.BOF <> True Then
                .Rows = .Rows + 1
                i = i + 1
                .TextMatrix(.Rows - 1, 0) = i
                .TextMatrix(.Rows - 1, 1) = Rctemp.Fields("Kol").Value
                .TextMatrix(.Rows - 1, 2) = Rctemp.Fields("Moein").Value
                .TextMatrix(.Rows - 1, 3) = IIf(IsNull(Rst.Fields("Person_Tafsili").Value), "", Rst.Fields("Person_Tafsili").Value)
                .TextMatrix(.Rows - 1, 4) = "ÕﬁÊﬁ -œ—Ì«›  «“ ’‰œÊﬁ" & " -- " & Rst.Fields("Person_Name").Value & " -- " & Rst.Fields("Date").Value
                .TextMatrix(.Rows - 1, 5) = Rst.Fields("TotalBestankar").Value
                .TextMatrix(.Rows - 1, 6) = 0
                Totalprice = Totalprice + Rst.Fields("TotalBestankar").Value
                        
            End If
            Rst.MoveNext
            If Rst.EOF <> True And Rst.BOF <> True Then
                 If Uid_Temp <> Rst!Uid Then
                     ReDim Parameter2(0) As Parameter
                     Parameter2(0) = GenerateInputParameter("@Code", adInteger, 4, 7)
                     Set Rctemp = RunParametricStoredProcedure2Rec("Get_tblAcc_Sale_ByCode", Parameter2)
                     
                     If Rctemp.EOF <> True And Rctemp.BOF <> True Then
                         .Rows = .Rows + 1
                         i = i + 1
                         .TextMatrix(.Rows - 1, 0) = i
                         .TextMatrix(.Rows - 1, 1) = Rctemp.Fields("Kol").Value
                         .TextMatrix(.Rows - 1, 2) = Rctemp.Fields("Moein").Value
                         .TextMatrix(.Rows - 1, 3) = Uid_Tafsili
                         .TextMatrix(.Rows - 1, 4) = Uid_Desc
                         .TextMatrix(.Rows - 1, 5) = 0
                         .TextMatrix(.Rows - 1, 6) = Totalprice
                         Totalprice = 0
                         Uid_Temp = Rst!Uid
                     End If
                  End If
            Else
                 ReDim Parameter2(0) As Parameter
                 Parameter2(0) = GenerateInputParameter("@Code", adInteger, 4, 7)
                 Set Rctemp = RunParametricStoredProcedure2Rec("Get_tblAcc_Sale_ByCode", Parameter2)
                 If Rctemp.EOF <> True And Rctemp.BOF <> True Then
                     .Rows = .Rows + 1
                     i = i + 1
                     .TextMatrix(.Rows - 1, 0) = i
                     .TextMatrix(.Rows - 1, 1) = Rctemp.Fields("Kol").Value
                     .TextMatrix(.Rows - 1, 2) = Rctemp.Fields("Moein").Value
                     .TextMatrix(.Rows - 1, 3) = Uid_Tafsili
                     .TextMatrix(.Rows - 1, 4) = Uid_Desc
                     .TextMatrix(.Rows - 1, 5) = 0
                     .TextMatrix(.Rows - 1, 6) = Totalprice
                     Totalprice = 0
                 End If
            End If
       Wend
   
    End If
  End With

End Sub
Private Sub UserSupplierPayment()
'User Payment Summary  ( Å—œ«Œ  »Â  «„Ì‰ ﬂ‰‰œê«‰ )
  Totalprice = 0
  Dim Uid_Temp As Integer
  Dim Uid_Desc, Uid_Tafsili As String
  Totalprice = 0
  Uid_Temp = 0
  With VsSanadView

    ReDim Parameter(3) As Parameter
    Parameter(0) = GenerateInputParameter("@PaymentType", adInteger, 4, EnumPaymentType.SupplierPayment)
    Parameter(1) = GenerateInputParameter("@DateBefore", adVarWChar, 8, CStr(IIf(Trim(txtDate1.ClipText) = "", "", Trim(txtDate1.Text))))
    Parameter(2) = GenerateInputParameter("@DateAfter", adVarWChar, 8, CStr(IIf(Trim(txtDate2.ClipText) = "", "", Trim(txtDate2.Text))))
    Parameter(3) = GenerateInputParameter("@Uid", adInteger, 4, cmbUsers.ItemData(cmbUsers.ListIndex))
    Set Rst = RunParametricStoredProcedure2Rec("Get_UserPayment", Parameter)
    
    If Rst.EOF <> True And Rst.BOF <> True Then
    
        Uid_Temp = Rst!Uid
        While Rst.EOF <> True
           
            Uid_Desc = "   Å—œ«Œ   »Â  «„Ì‰ ò‰‰œÂ " & " -- " & Rst.Fields("User_Name").Value & " -- " & Rst.Fields("Date").Value
            Uid_Tafsili = IIf(IsNull(Rst.Fields("Tafsili").Value), "", Rst.Fields("Tafsili").Value)
            
            ReDim Parameter2(0) As Parameter
            Parameter2(0) = GenerateInputParameter("@Code", adInteger, 4, EnumAccountCodes.Bestankaran)
            Set Rctemp = RunParametricStoredProcedure2Rec("Get_tblAcc_Sale_ByCode", Parameter2)
            If Rctemp.EOF <> True And Rctemp.BOF <> True Then
                .Rows = .Rows + 1
                i = i + 1
                .TextMatrix(.Rows - 1, 0) = i
                .TextMatrix(.Rows - 1, 1) = Rctemp.Fields("Kol").Value
                .TextMatrix(.Rows - 1, 2) = Rctemp.Fields("Moein").Value
                .TextMatrix(.Rows - 1, 3) = IIf(IsNull(Rst.Fields("Person_Tafsili").Value), "", Rst.Fields("Person_Tafsili").Value)
                .TextMatrix(.Rows - 1, 4) = "œ—Ì«›  «“ ’‰œÊﬁ(Œ—Ìœ)" & " -- " & Rst.Fields("Person_Name").Value & " -- " & Rst.Fields("Date").Value
                .TextMatrix(.Rows - 1, 5) = Rst.Fields("TotalBestankar").Value
                .TextMatrix(.Rows - 1, 6) = 0
                Totalprice = Totalprice + Rst.Fields("TotalBestankar").Value
                        
            End If
            Rst.MoveNext
            If Rst.EOF <> True And Rst.BOF <> True Then
                 If Uid_Temp <> Rst!Uid Then
                     ReDim Parameter2(0) As Parameter
                     Parameter2(0) = GenerateInputParameter("@Code", adInteger, 4, 7)
                     Set Rctemp = RunParametricStoredProcedure2Rec("Get_tblAcc_Sale_ByCode", Parameter2)
                     
                     If Rctemp.EOF <> True And Rctemp.BOF <> True Then
                         .Rows = .Rows + 1
                         i = i + 1
                         .TextMatrix(.Rows - 1, 0) = i
                         .TextMatrix(.Rows - 1, 1) = Rctemp.Fields("Kol").Value
                         .TextMatrix(.Rows - 1, 2) = Rctemp.Fields("Moein").Value
                         .TextMatrix(.Rows - 1, 3) = Uid_Tafsili
                         .TextMatrix(.Rows - 1, 4) = Uid_Desc
                         .TextMatrix(.Rows - 1, 5) = 0
                         .TextMatrix(.Rows - 1, 6) = Totalprice
                         Totalprice = 0
                         Uid_Temp = Rst!Uid
                     End If
                  End If
            Else
                 ReDim Parameter2(0) As Parameter
                 Parameter2(0) = GenerateInputParameter("@Code", adInteger, 4, 7)
                 Set Rctemp = RunParametricStoredProcedure2Rec("Get_tblAcc_Sale_ByCode", Parameter2)
                 If Rctemp.EOF <> True And Rctemp.BOF <> True Then
                     .Rows = .Rows + 1
                     i = i + 1
                     .TextMatrix(.Rows - 1, 0) = i
                     .TextMatrix(.Rows - 1, 1) = Rctemp.Fields("Kol").Value
                     .TextMatrix(.Rows - 1, 2) = Rctemp.Fields("Moein").Value
                     .TextMatrix(.Rows - 1, 3) = Uid_Tafsili
                     .TextMatrix(.Rows - 1, 4) = Uid_Desc
                     .TextMatrix(.Rows - 1, 5) = 0
                     .TextMatrix(.Rows - 1, 6) = Totalprice
                     Totalprice = 0
                 End If
            End If
       Wend
   
    End If
 End With

End Sub
Private Sub SupplierSummaryNoCash()
 
 'Supplier - €Ì— ‰ﬁœÌ - Œ—Ìœ «“  «„Ì‰ ﬂ‰‰œê«‰
  
Dim TotalDutyBuy As Long
Dim TotaltaxBuy As Long
Dim TotalDutyBuyReturn As Long
Dim TotaltaxBuyReturn As Long
TotalDutyBuy = 0
TotaltaxBuy = 0
TotalDutyBuyReturn = 0
TotaltaxBuyReturn = 0
Totalprice = 0
    
    ReDim Parameter(2) As Parameter
    Parameter(0) = GenerateInputParameter("@Branch", adInteger, 4, cboBranch.ItemData(cboBranch.ListIndex))
    Parameter(1) = GenerateInputParameter("@DateBefore", adVarWChar, 8, CStr(IIf(Trim(txtDate1.ClipText) = "", "", Trim(txtDate1.Text))))
    Parameter(2) = GenerateInputParameter("@DateAfter", adVarWChar, 8, CStr(IIf(Trim(txtDate2.ClipText) = "", "", Trim(txtDate2.Text))))
    Set Rst = RunParametricStoredProcedure2Rec("Get_SupplierBuySummary", Parameter)

    If Rst.EOF <> True And Rst.BOF <> True Then
    
         With VsSanadView
          While Rst.EOF <> True
                  ReDim Parameter2(0) As Parameter
                  Parameter2(0) = GenerateInputParameter("@Code", adInteger, 4, 15)
                  Set Rctemp = RunParametricStoredProcedure2Rec("Get_tblAcc_Sale_ByCode", Parameter2)
              
                  If Rctemp.EOF <> True And Rctemp.BOF <> True Then
                     .Rows = .Rows + 1
                     i = i + 1
                     .TextMatrix(.Rows - 1, 0) = i
                     .TextMatrix(.Rows - 1, 1) = Rctemp.Fields("Kol").Value
                     .TextMatrix(.Rows - 1, 2) = Rctemp.Fields("Moein").Value
                     .TextMatrix(.Rows - 1, 3) = IIf(IsNull(Rst.Fields("SupplierTafsili").Value), "", Rst.Fields("SupplierTafsili").Value)
                     .TextMatrix(.Rows - 1, 4) = Rst.Fields("SupplierName").Value & " -- " & Rst.Fields("NvcDescription").Value
                     .TextMatrix(.Rows - 1, 5) = 0
                     .TextMatrix(.Rows - 1, 6) = Rst.Fields("SumPrice").Value
                     
                     Totalprice = Totalprice + Rst.Fields("SumPrice").Value
                     
                      ReDim Parameter2(0) As Parameter
                      Parameter2(0) = GenerateInputParameter("@Code", adInteger, 4, 16)
                      Set Rctemp = RunParametricStoredProcedure2Rec("Get_tblAcc_Sale_ByCode", Parameter2)
                  
                      If Rctemp.EOF <> True And Rctemp.BOF <> True Then
                          .Rows = .Rows + 1
                          i = i + 1
                          .TextMatrix(.Rows - 1, 0) = i
                          .TextMatrix(.Rows - 1, 1) = Rctemp.Fields("Kol").Value
                          .TextMatrix(.Rows - 1, 2) = Rctemp.Fields("Moein").Value
                          If Rctemp.Fields("Tafsili").Value = "0" Then
                            .TextMatrix(.Rows - 1, 3) = IIf(IsNull(Rst.Fields("SupplierTafsili").Value), "", Rst.Fields("SupplierTafsili").Value)
                          Else
                            .TextMatrix(.Rows - 1, 3) = Rctemp.Fields("Tafsili").Value
                          End If
                          .TextMatrix(.Rows - 1, 4) = "„ÊÃÊœÌ „Ê«œ Ê ﬂ«·« -- " & Rst.Fields("SupplierName").Value '& " -- " & Rst.Fields("NvcDescription").Value
                          .TextMatrix(.Rows - 1, 5) = Rst.Fields("SumPrice").Value - Val(Rst.Fields("DutyTotal").Value) - Val(Rst.Fields("TaxTotal").Value)
                          .TextMatrix(.Rows - 1, 6) = 0
                      
                            If Rst!Status = 1 Then
                                  TotalDutyBuy = TotalDutyBuy + Val(Rst.Fields("DutyTotal").Value)
                                  TotaltaxBuy = TotaltaxBuy + Val(Rst.Fields("TaxTotal").Value)
                            Else
                                  TotalDutyBuyReturn = TotalDutyBuyReturn + Val(Rst.Fields("DutyTotal").Value)
                                  TotaltaxBuyReturn = TotaltaxBuyReturn + Val(Rst.Fields("TaxTotal").Value)
                            End If
                      
                      End If
                  End If
                  Rst.MoveNext
             Wend
           
                ReDim Parameter2(0) As Parameter
                Parameter2(0) = GenerateInputParameter("@Code", adInteger, 4, 25)
                Set Rctemp = RunParametricStoredProcedure2Rec("Get_tblAcc_Sale_ByCode", Parameter2)
    
                If Rctemp.EOF <> True And Rctemp.BOF <> True Then
                     If TotalDutyBuy <> 0 Then
                        .Rows = .Rows + 1
                        i = i + 1
                        .TextMatrix(.Rows - 1, 0) = i
                        .TextMatrix(.Rows - 1, 1) = Rctemp.Fields("Kol").Value
                        .TextMatrix(.Rows - 1, 2) = Rctemp.Fields("Moein").Value
                        .TextMatrix(.Rows - 1, 3) = IIf(IsNull(Rctemp.Fields("Tafsili").Value), "", Rctemp.Fields("Tafsili").Value)
                        .TextMatrix(.Rows - 1, 4) = Rctemp.Fields("Description").Value
                        .TextMatrix(.Rows - 1, 5) = TotalDutyBuy
                        .TextMatrix(.Rows - 1, 6) = 0
                   End If
                     If TotalDutyBuyReturn <> 0 Then
                        .Rows = .Rows + 1
                        i = i + 1
                        .TextMatrix(.Rows - 1, 0) = i
                        .TextMatrix(.Rows - 1, 1) = Rctemp.Fields("Kol").Value
                        .TextMatrix(.Rows - 1, 2) = Rctemp.Fields("Moein").Value
                        .TextMatrix(.Rows - 1, 3) = IIf(IsNull(Rctemp.Fields("Tafsili").Value), "", Rctemp.Fields("Tafsili").Value)
                        .TextMatrix(.Rows - 1, 4) = "⁄Ê«—÷ »—ê‘  «“ Œ—Ìœ"
                        .TextMatrix(.Rows - 1, 5) = 0
                        .TextMatrix(.Rows - 1, 6) = TotalDutyBuyReturn
                   End If
                End If
                ReDim Parameter2(0) As Parameter
                Parameter2(0) = GenerateInputParameter("@Code", adInteger, 4, 27)
                Set Rctemp = RunParametricStoredProcedure2Rec("Get_tblAcc_Sale_ByCode", Parameter2)
    
                If Rctemp.EOF <> True And Rctemp.BOF <> True Then
                    If TotaltaxBuy <> 0 Then
                        .Rows = .Rows + 1
                        i = i + 1
                        .TextMatrix(.Rows - 1, 0) = i
                        .TextMatrix(.Rows - 1, 1) = Rctemp.Fields("Kol").Value
                        .TextMatrix(.Rows - 1, 2) = Rctemp.Fields("Moein").Value
                        .TextMatrix(.Rows - 1, 3) = IIf(IsNull(Rctemp.Fields("Tafsili").Value), "", Rctemp.Fields("Tafsili").Value)
                        .TextMatrix(.Rows - 1, 4) = Rctemp("Description").Value
                        .TextMatrix(.Rows - 1, 5) = TotaltaxBuy
                        .TextMatrix(.Rows - 1, 6) = 0
                  End If
                    If TotaltaxBuyReturn <> 0 Then
                        .Rows = .Rows + 1
                        i = i + 1
                        .TextMatrix(.Rows - 1, 0) = i
                        .TextMatrix(.Rows - 1, 1) = Rctemp.Fields("Kol").Value
                        .TextMatrix(.Rows - 1, 2) = Rctemp.Fields("Moein").Value
                        .TextMatrix(.Rows - 1, 3) = IIf(IsNull(Rctemp.Fields("Tafsili").Value), "", Rctemp.Fields("Tafsili").Value)
                        .TextMatrix(.Rows - 1, 4) = "„«·Ì«  »—ê‘  «“ Œ—Ìœ"
                        .TextMatrix(.Rows - 1, 5) = 0
                        .TextMatrix(.Rows - 1, 6) = TotaltaxBuyReturn
                  End If
                End If
           
           End With
       
                                          ' Total Buy
    
    End If

End Sub
Private Sub BuyReturn()
 'Supplier - €Ì— ‰ﬁœÌ - Œ—Ìœ «“  «„Ì‰ ﬂ‰‰œê«‰
  Totalprice = 0
    
    ReDim Parameter(2) As Parameter
    Parameter(0) = GenerateInputParameter("@Branch", adInteger, 4, cboBranch.ItemData(cboBranch.ListIndex))
    Parameter(1) = GenerateInputParameter("@DateBefore", adVarWChar, 8, CStr(IIf(Trim(txtDate1.ClipText) = "", "", Trim(txtDate1.Text))))
    Parameter(2) = GenerateInputParameter("@DateAfter", adVarWChar, 8, CStr(IIf(Trim(txtDate2.ClipText) = "", "", Trim(txtDate2.Text))))
    Set Rst = RunParametricStoredProcedure2Rec("Get_BuyReturnSummary", Parameter)

    If Rst.EOF <> True And Rst.BOF <> True Then
    With VsSanadView
    
        ReDim Parameter2(0) As Parameter
        Parameter2(0) = GenerateInputParameter("@Code", adInteger, 4, 15)
        Set Rctemp = RunParametricStoredProcedure2Rec("Get_tblAcc_Sale_ByCode", Parameter2)
    
        If Rctemp.EOF <> True And Rctemp.BOF <> True Then
            While Rst.EOF <> True
                   .Rows = .Rows + 1
                   i = i + 1
                   .TextMatrix(.Rows - 1, 0) = i
                   .TextMatrix(.Rows - 1, 1) = Rctemp.Fields("Kol").Value
                   
                   .TextMatrix(.Rows - 1, 2) = Rctemp.Fields("Moein").Value
                   .TextMatrix(.Rows - 1, 3) = IIf(IsNull(Rst.Fields("SupplierTafsili").Value), "", Rst.Fields("SupplierTafsili").Value)
                   
                   .TextMatrix(.Rows - 1, 4) = Rst.Fields("SupplierName").Value & " -- " & Rst.Fields("Date").Value & "-- »«»  »—ê‘  «“ Œ—Ìœ" & Rst.Fields("No").Value
                   .TextMatrix(.Rows - 1, 5) = Rst.Fields("SumPrice").Value
                   .TextMatrix(.Rows - 1, 6) = 0
                   
                   Totalprice = Totalprice + Rst.Fields("SumPrice").Value
                   
        
                                          ' Total Buy
                ReDim Parameter2(0) As Parameter
                Parameter2(0) = GenerateInputParameter("@Code", adInteger, 4, 18)
                Set Rctemp = RunParametricStoredProcedure2Rec("Get_tblAcc_Sale_ByCode", Parameter2)
            
                If Rctemp.EOF <> True And Rctemp.BOF <> True Then
                           .Rows = .Rows + 1
                        i = i + 1
                        .TextMatrix(.Rows - 1, 0) = i
                        .TextMatrix(.Rows - 1, 1) = Rctemp.Fields("Kol").Value
                        .TextMatrix(.Rows - 1, 2) = Rctemp.Fields("Moein").Value
                          If Rctemp.Fields("Tafsili").Value = "0" Then
                            .TextMatrix(.Rows - 1, 3) = IIf(IsNull(Rst.Fields("SupplierTafsili").Value), "", Rst.Fields("SupplierTafsili").Value)
                          Else
                            .TextMatrix(.Rows - 1, 3) = Rctemp.Fields("Tafsili").Value
                        End If
                        .TextMatrix(.Rows - 1, 4) = Rctemp.Fields("Description").Value & "-- „ÊÃÊœÌ „Ê«œ Ê ﬂ«·«  "
                        .TextMatrix(.Rows - 1, 5) = 0
                        .TextMatrix(.Rows - 1, 6) = Rst.Fields("SumPrice").Value  'Totalprice
                End If
            
                Rst.MoveNext
                Wend
        End If
    End With
    End If
   

End Sub


Private Sub UserPersonRecieve()
'User Payment Summary  ( œ—Ì«›  «“ Å—”‰· )
  Totalprice = 0
  Dim Uid_Temp As Integer
  Dim Uid_Desc, Uid_Tafsili As String
  Totalprice = 0
  Uid_Temp = 0
  With VsSanadView

    ReDim Parameter(3) As Parameter
    Parameter(0) = GenerateInputParameter("@RecieveType", adInteger, 4, EnumRecieveType.TempPersonRecieve)
    Parameter(1) = GenerateInputParameter("@DateBefore", adVarWChar, 8, CStr(IIf(Trim(txtDate1.ClipText) = "", "", Trim(txtDate1.Text))))
    Parameter(2) = GenerateInputParameter("@DateAfter", adVarWChar, 8, CStr(IIf(Trim(txtDate2.ClipText) = "", "", Trim(txtDate2.Text))))
    Parameter(3) = GenerateInputParameter("@Uid", adInteger, 4, cmbUsers.ItemData(cmbUsers.ListIndex))
    Set Rst = RunParametricStoredProcedure2Rec("Get_UserRecieve", Parameter)
    
    If Rst.EOF <> True And Rst.BOF <> True Then
    
        Uid_Temp = Rst!Uid
        While Rst.EOF <> True
           
            Uid_Desc = "œ—Ì«›  «“ Å—”‰· - ⁄·Ì «·Õ”«»" & " -- " & Rst.Fields("User_Name").Value & " -- " & Rst.Fields("Date").Value
            Uid_Tafsili = IIf(IsNull(Rst.Fields("Tafsili").Value), "", Rst.Fields("Tafsili").Value)
            
            ReDim Parameter2(0) As Parameter
            Parameter2(0) = GenerateInputParameter("@Code", adInteger, 4, 10)
            Set Rctemp = RunParametricStoredProcedure2Rec("Get_tblAcc_Sale_ByCode", Parameter2)
             With VsSanadView
                If Rctemp.EOF <> True And Rctemp.BOF <> True Then
                    .Rows = .Rows + 1
                    i = i + 1
                    .TextMatrix(.Rows - 1, 0) = i
                    .TextMatrix(.Rows - 1, 1) = Rctemp.Fields("Kol").Value
                    .TextMatrix(.Rows - 1, 2) = Rctemp.Fields("Moein").Value
                    .TextMatrix(.Rows - 1, 3) = IIf(IsNull(Rst.Fields("Person_Tafsili").Value), "", Rst.Fields("Person_Tafsili").Value)
                    .TextMatrix(.Rows - 1, 4) = " Å—œ«Œ  »Â ’‰œÊﬁ- ⁄·Ì «·Õ”«»" & " -- " & Rst.Fields("Person_Name").Value & " -- " & Rst.Fields("Date").Value
                    .TextMatrix(.Rows - 1, 5) = 0
                    .TextMatrix(.Rows - 1, 6) = Rst.Fields("TotalBestankar").Value
                    Totalprice = Totalprice + Rst.Fields("TotalBestankar").Value
                End If
            End With
            Rst.MoveNext
            If Rst.EOF <> True And Rst.BOF <> True Then
                 If Uid_Temp <> Rst!Uid Then
                     ReDim Parameter2(0) As Parameter
                     Parameter2(0) = GenerateInputParameter("@Code", adInteger, 4, 7)
                     Set Rctemp = RunParametricStoredProcedure2Rec("Get_tblAcc_Sale_ByCode", Parameter2)
                     
                     If Rctemp.EOF <> True And Rctemp.BOF <> True Then
                         .Rows = .Rows + 1
                         i = i + 1
                         .TextMatrix(.Rows - 1, 0) = i
                         .TextMatrix(.Rows - 1, 1) = Rctemp.Fields("Kol").Value
                         .TextMatrix(.Rows - 1, 2) = Rctemp.Fields("Moein").Value
                         .TextMatrix(.Rows - 1, 3) = Uid_Tafsili
                         .TextMatrix(.Rows - 1, 4) = Uid_Desc
                         .TextMatrix(.Rows - 1, 5) = Totalprice
                         .TextMatrix(.Rows - 1, 6) = 0
                         Totalprice = 0
                         Uid_Temp = Rst!Uid
                     End If
                  End If
            Else
                 ReDim Parameter2(0) As Parameter
                 Parameter2(0) = GenerateInputParameter("@Code", adInteger, 4, 7)
                 Set Rctemp = RunParametricStoredProcedure2Rec("Get_tblAcc_Sale_ByCode", Parameter2)
                 If Rctemp.EOF <> True And Rctemp.BOF <> True Then
                     .Rows = .Rows + 1
                     i = i + 1
                     .TextMatrix(.Rows - 1, 0) = i
                     .TextMatrix(.Rows - 1, 1) = Rctemp.Fields("Kol").Value
                     .TextMatrix(.Rows - 1, 2) = Rctemp.Fields("Moein").Value
                     .TextMatrix(.Rows - 1, 3) = Uid_Tafsili
                     .TextMatrix(.Rows - 1, 4) = Uid_Desc
                     .TextMatrix(.Rows - 1, 5) = Totalprice
                     .TextMatrix(.Rows - 1, 6) = 0
                     Totalprice = 0
                 End If
            End If
       Wend
   
    End If
    End With
    
  With VsSanadView

    ReDim Parameter(3) As Parameter
    Parameter(0) = GenerateInputParameter("@RecieveType", adInteger, 4, EnumRecieveType.VamPersonRecieve)
    Parameter(1) = GenerateInputParameter("@DateBefore", adVarWChar, 8, CStr(IIf(Trim(txtDate1.ClipText) = "", "", Trim(txtDate1.Text))))
    Parameter(2) = GenerateInputParameter("@DateAfter", adVarWChar, 8, CStr(IIf(Trim(txtDate2.ClipText) = "", "", Trim(txtDate2.Text))))
    Parameter(3) = GenerateInputParameter("@Uid", adInteger, 4, cmbUsers.ItemData(cmbUsers.ListIndex))
    Set Rst = RunParametricStoredProcedure2Rec("Get_UserRecieve", Parameter)
    
    If Rst.EOF <> True And Rst.BOF <> True Then
    
        Uid_Temp = Rst!Uid
        While Rst.EOF <> True
           
            Uid_Desc = "œ—Ì«›  «“ Å—”‰· - Ê«„" & " -- " & Rst.Fields("User_Name").Value & " -- " & Rst.Fields("Date").Value
            Uid_Tafsili = IIf(IsNull(Rst.Fields("Tafsili").Value), "", Rst.Fields("Tafsili").Value)
            
            ReDim Parameter2(0) As Parameter
            Parameter2(0) = GenerateInputParameter("@Code", adInteger, 4, 11)
            Set Rctemp = RunParametricStoredProcedure2Rec("Get_tblAcc_Sale_ByCode", Parameter2)
             With VsSanadView
                If Rctemp.EOF <> True And Rctemp.BOF <> True Then
                    .Rows = .Rows + 1
                    i = i + 1
                    .TextMatrix(.Rows - 1, 0) = i
                    .TextMatrix(.Rows - 1, 1) = Rctemp.Fields("Kol").Value
                    .TextMatrix(.Rows - 1, 2) = Rctemp.Fields("Moein").Value
                    .TextMatrix(.Rows - 1, 3) = IIf(IsNull(Rst.Fields("Person_Tafsili").Value), "", Rst.Fields("Person_Tafsili").Value)
                    .TextMatrix(.Rows - 1, 4) = " Å—œ«Œ  »Â ’‰œÊﬁ- Ê«„" & " -- " & Rst.Fields("Person_Name").Value & " -- " & Rst.Fields("Date").Value
                    .TextMatrix(.Rows - 1, 5) = 0
                    .TextMatrix(.Rows - 1, 6) = Rst.Fields("TotalBestankar").Value
                    Totalprice = Totalprice + Rst.Fields("TotalBestankar").Value
                End If
            End With
            Rst.MoveNext
            If Rst.EOF <> True And Rst.BOF <> True Then
                 If Uid_Temp <> Rst!Uid Then
                     ReDim Parameter2(0) As Parameter
                     Parameter2(0) = GenerateInputParameter("@Code", adInteger, 4, 7)
                     Set Rctemp = RunParametricStoredProcedure2Rec("Get_tblAcc_Sale_ByCode", Parameter2)
                     
                     If Rctemp.EOF <> True And Rctemp.BOF <> True Then
                         .Rows = .Rows + 1
                         i = i + 1
                         .TextMatrix(.Rows - 1, 0) = i
                         .TextMatrix(.Rows - 1, 1) = Rctemp.Fields("Kol").Value
                         .TextMatrix(.Rows - 1, 2) = Rctemp.Fields("Moein").Value
                         .TextMatrix(.Rows - 1, 3) = Uid_Tafsili
                         .TextMatrix(.Rows - 1, 4) = Uid_Desc
                         .TextMatrix(.Rows - 1, 5) = Totalprice
                         .TextMatrix(.Rows - 1, 6) = 0
                         Totalprice = 0
                         Uid_Temp = Rst!Uid
                     End If
                  End If
            Else
                 ReDim Parameter2(0) As Parameter
                 Parameter2(0) = GenerateInputParameter("@Code", adInteger, 4, 7)
                 Set Rctemp = RunParametricStoredProcedure2Rec("Get_tblAcc_Sale_ByCode", Parameter2)
                 If Rctemp.EOF <> True And Rctemp.BOF <> True Then
                     .Rows = .Rows + 1
                     i = i + 1
                     .TextMatrix(.Rows - 1, 0) = i
                     .TextMatrix(.Rows - 1, 1) = Rctemp.Fields("Kol").Value
                     .TextMatrix(.Rows - 1, 2) = Rctemp.Fields("Moein").Value
                     .TextMatrix(.Rows - 1, 3) = Uid_Tafsili
                     .TextMatrix(.Rows - 1, 4) = Uid_Desc
                     .TextMatrix(.Rows - 1, 5) = Totalprice
                     .TextMatrix(.Rows - 1, 6) = 0
                     Totalprice = 0
                 End If
            End If
       Wend
   
    End If
End With
  With VsSanadView

    ReDim Parameter(3) As Parameter
    Parameter(0) = GenerateInputParameter("@RecieveType", adInteger, 4, EnumRecieveType.MosaedePersonRecieve)
    Parameter(1) = GenerateInputParameter("@DateBefore", adVarWChar, 8, CStr(IIf(Trim(txtDate1.ClipText) = "", "", Trim(txtDate1.Text))))
    Parameter(2) = GenerateInputParameter("@DateAfter", adVarWChar, 8, CStr(IIf(Trim(txtDate2.ClipText) = "", "", Trim(txtDate2.Text))))
    Parameter(3) = GenerateInputParameter("@Uid", adInteger, 4, cmbUsers.ItemData(cmbUsers.ListIndex))
    Set Rst = RunParametricStoredProcedure2Rec("Get_UserRecieve", Parameter)
    
    If Rst.EOF <> True And Rst.BOF <> True Then
    
        Uid_Temp = Rst!Uid
        While Rst.EOF <> True
           
            Uid_Desc = "œ—Ì«›  «“ Å—”‰· - „”«⁄œÂ" & " -- " & Rst.Fields("User_Name").Value & " -- " & Rst.Fields("Date").Value
            Uid_Tafsili = IIf(IsNull(Rst.Fields("Tafsili").Value), "", Rst.Fields("Tafsili").Value)
            
            ReDim Parameter2(0) As Parameter
            Parameter2(0) = GenerateInputParameter("@Code", adInteger, 4, 12)
            Set Rctemp = RunParametricStoredProcedure2Rec("Get_tblAcc_Sale_ByCode", Parameter2)
             With VsSanadView
                If Rctemp.EOF <> True And Rctemp.BOF <> True Then
                    .Rows = .Rows + 1
                    i = i + 1
                    .TextMatrix(.Rows - 1, 0) = i
                    .TextMatrix(.Rows - 1, 1) = Rctemp.Fields("Kol").Value
                    .TextMatrix(.Rows - 1, 2) = Rctemp.Fields("Moein").Value
                    .TextMatrix(.Rows - 1, 3) = IIf(IsNull(Rst.Fields("Person_Tafsili").Value), "", Rst.Fields("Person_Tafsili").Value)
                    .TextMatrix(.Rows - 1, 4) = " Å—œ«Œ  »Â ’‰œÊﬁ- „”«⁄œÂ" & " -- " & Rst.Fields("Person_Name").Value & " -- " & Rst.Fields("Date").Value
                    .TextMatrix(.Rows - 1, 5) = 0
                    .TextMatrix(.Rows - 1, 6) = Rst.Fields("TotalBestankar").Value
                    Totalprice = Totalprice + Rst.Fields("TotalBestankar").Value
                End If
            End With
            Rst.MoveNext
            If Rst.EOF <> True And Rst.BOF <> True Then
                 If Uid_Temp <> Rst!Uid Then
                     ReDim Parameter2(0) As Parameter
                     Parameter2(0) = GenerateInputParameter("@Code", adInteger, 4, 7)
                     Set Rctemp = RunParametricStoredProcedure2Rec("Get_tblAcc_Sale_ByCode", Parameter2)
                     
                     If Rctemp.EOF <> True And Rctemp.BOF <> True Then
                         .Rows = .Rows + 1
                         i = i + 1
                         .TextMatrix(.Rows - 1, 0) = i
                         .TextMatrix(.Rows - 1, 1) = Rctemp.Fields("Kol").Value
                         .TextMatrix(.Rows - 1, 2) = Rctemp.Fields("Moein").Value
                         .TextMatrix(.Rows - 1, 3) = Uid_Tafsili
                         .TextMatrix(.Rows - 1, 4) = Uid_Desc
                         .TextMatrix(.Rows - 1, 5) = Totalprice
                         .TextMatrix(.Rows - 1, 6) = 0
                         Totalprice = 0
                         Uid_Temp = Rst!Uid
                     End If
                  End If
            Else
                 ReDim Parameter2(0) As Parameter
                 Parameter2(0) = GenerateInputParameter("@Code", adInteger, 4, 7)
                 Set Rctemp = RunParametricStoredProcedure2Rec("Get_tblAcc_Sale_ByCode", Parameter2)
                 If Rctemp.EOF <> True And Rctemp.BOF <> True Then
                     .Rows = .Rows + 1
                     i = i + 1
                     .TextMatrix(.Rows - 1, 0) = i
                     .TextMatrix(.Rows - 1, 1) = Rctemp.Fields("Kol").Value
                     .TextMatrix(.Rows - 1, 2) = Rctemp.Fields("Moein").Value
                     .TextMatrix(.Rows - 1, 3) = Uid_Tafsili
                     .TextMatrix(.Rows - 1, 4) = Uid_Desc
                     .TextMatrix(.Rows - 1, 5) = Totalprice
                     .TextMatrix(.Rows - 1, 6) = 0
                     Totalprice = 0
                 End If
            End If
       Wend
   
    End If
End With
    
End Sub
Private Sub UserCustomerRecieve()
'User Payment Summary  ( œ—Ì«›  «“ „‘ —Ì«‰ )
  Totalprice = 0
  Dim Uid_Temp As Integer
  Dim Uid_Desc, Uid_Tafsili As String
  Totalprice = 0
  Uid_Temp = 0
  With VsSanadView

    ReDim Parameter(3) As Parameter
    Parameter(0) = GenerateInputParameter("@RecieveType", adInteger, 4, EnumRecieveType.CustomerRecieve)
    Parameter(1) = GenerateInputParameter("@DateBefore", adVarWChar, 8, CStr(IIf(Trim(txtDate1.ClipText) = "", "", Trim(txtDate1.Text))))
    Parameter(2) = GenerateInputParameter("@DateAfter", adVarWChar, 8, CStr(IIf(Trim(txtDate2.ClipText) = "", "", Trim(txtDate2.Text))))
    Parameter(3) = GenerateInputParameter("@Uid", adInteger, 4, cmbUsers.ItemData(cmbUsers.ListIndex))
    Set Rst = RunParametricStoredProcedure2Rec("Get_UserRecieve", Parameter)
    
    If Rst.EOF <> True And Rst.BOF <> True Then
    
        Uid_Temp = Rst!Uid
        While Rst.EOF <> True
           
            Uid_Desc = "œ—Ì«›  «“ „‘ —Ì«‰" & " -- " & Rst.Fields("User_Name").Value & " -- " & Rst.Fields("Date").Value
            Uid_Tafsili = IIf(IsNull(Rst.Fields("Tafsili").Value), "", Rst.Fields("Tafsili").Value)
            
            ReDim Parameter2(0) As Parameter
            Parameter2(0) = GenerateInputParameter("@Code", adInteger, 4, EnumAccountCodes.Bedehkaran)
            Set Rctemp = RunParametricStoredProcedure2Rec("Get_tblAcc_Sale_ByCode", Parameter2)
             With VsSanadView
                If Rctemp.EOF <> True And Rctemp.BOF <> True Then
                    .Rows = .Rows + 1
                    i = i + 1
                    .TextMatrix(.Rows - 1, 0) = i
                    .TextMatrix(.Rows - 1, 1) = Rctemp.Fields("Kol").Value
                    .TextMatrix(.Rows - 1, 2) = Rctemp.Fields("Moein").Value
                    .TextMatrix(.Rows - 1, 3) = IIf(IsNull(Rst.Fields("Person_Tafsili").Value), "", Rst.Fields("Person_Tafsili").Value)
                    .TextMatrix(.Rows - 1, 4) = " Å—œ«Œ   »Â ’‰œÊﬁ" & " -- " & Rst.Fields("Person_Name").Value & " -- " & Rst.Fields("Date").Value
                    .TextMatrix(.Rows - 1, 5) = 0
                    .TextMatrix(.Rows - 1, 6) = Rst.Fields("TotalBestankar").Value
                    Totalprice = Totalprice + Rst.Fields("TotalBestankar").Value
                            
                End If
            End With
            Rst.MoveNext
            If Rst.EOF <> True And Rst.BOF <> True Then
                 If Uid_Temp <> Rst!Uid Then
                     ReDim Parameter2(0) As Parameter
                     Parameter2(0) = GenerateInputParameter("@Code", adInteger, 4, 7)
                     Set Rctemp = RunParametricStoredProcedure2Rec("Get_tblAcc_Sale_ByCode", Parameter2)
                     
                     If Rctemp.EOF <> True And Rctemp.BOF <> True Then
                         .Rows = .Rows + 1
                         i = i + 1
                         .TextMatrix(.Rows - 1, 0) = i
                         .TextMatrix(.Rows - 1, 1) = Rctemp.Fields("Kol").Value
                         .TextMatrix(.Rows - 1, 2) = Rctemp.Fields("Moein").Value
                         .TextMatrix(.Rows - 1, 3) = Uid_Tafsili
                         .TextMatrix(.Rows - 1, 4) = Uid_Desc
                         .TextMatrix(.Rows - 1, 5) = Totalprice
                         .TextMatrix(.Rows - 1, 6) = 0
                         Totalprice = 0
                         Uid_Temp = Rst!Uid
                     End If
                  End If
            Else
                 ReDim Parameter2(0) As Parameter
                 Parameter2(0) = GenerateInputParameter("@Code", adInteger, 4, 7)
                 Set Rctemp = RunParametricStoredProcedure2Rec("Get_tblAcc_Sale_ByCode", Parameter2)
                 If Rctemp.EOF <> True And Rctemp.BOF <> True Then
                     .Rows = .Rows + 1
                     i = i + 1
                     .TextMatrix(.Rows - 1, 0) = i
                     .TextMatrix(.Rows - 1, 1) = Rctemp.Fields("Kol").Value
                     .TextMatrix(.Rows - 1, 2) = Rctemp.Fields("Moein").Value
                     .TextMatrix(.Rows - 1, 3) = Uid_Tafsili
                     .TextMatrix(.Rows - 1, 4) = Uid_Desc
                     .TextMatrix(.Rows - 1, 5) = Totalprice
                     .TextMatrix(.Rows - 1, 6) = 0
                     Totalprice = 0
                 End If
            End If
       Wend
   
    End If


End With
End Sub
Private Sub UserCustomerCheckRecieve()
'User Payment Summary  ( œ—Ì«›  çﬂ «“ „‘ —Ì«‰ )
  Totalprice = 0
  Dim Uid_Temp As Integer
  Dim Uid_Desc, Uid_Tafsili As String
  Totalprice = 0
  Uid_Temp = 0
  With VsSanadView

    
    ReDim Parameter(1)
    Parameter(0) = GenerateInputParameter("@DateBefore", adVarWChar, 8, CStr(IIf(Trim(txtDate1.ClipText) = "", "", Trim(txtDate1.Text))))
    Parameter(1) = GenerateInputParameter("@DateAfter", adVarWChar, 8, CStr(IIf(Trim(txtDate2.ClipText) = "", "", Trim(txtDate2.Text))))
    
    Set Rst = RunParametricStoredProcedure2Rec("Get_ChequeRecieved", Parameter)
    
    If Rst.EOF <> True And Rst.BOF <> True Then
    
        While Rst.EOF <> True
           
            Uid_Desc = Rst.Fields("Description").Value & "‘„«—Â - " & Rst.Fields("intChequeAcc").Value
            Uid_Tafsili = IIf(IsNull(Rst.Fields("Tafsili").Value), "", Rst.Fields("Tafsili").Value)
            
            ReDim Parameter2(0) As Parameter
            Parameter2(0) = GenerateInputParameter("@Code", adInteger, 4, 6)
            Set Rctemp = RunParametricStoredProcedure2Rec("Get_tblAcc_Sale_ByCode", Parameter2)
            If Rctemp.EOF <> True And Rctemp.BOF <> True Then
                .Rows = .Rows + 1
                i = i + 1
                .TextMatrix(.Rows - 1, 0) = i
                .TextMatrix(.Rows - 1, 1) = Rctemp.Fields("Kol").Value
                .TextMatrix(.Rows - 1, 2) = Rctemp.Fields("Moein").Value
                .TextMatrix(.Rows - 1, 3) = Uid_Tafsili
                .TextMatrix(.Rows - 1, 4) = Uid_Desc
                .TextMatrix(.Rows - 1, 5) = 0
                .TextMatrix(.Rows - 1, 6) = Rst.Fields("intChequeAmount").Value
                Totalprice = Rst.Fields("intChequeAmount").Value
                        
            End If
            ReDim Parameter2(0) As Parameter
            Parameter2(0) = GenerateInputParameter("@Code", adInteger, 4, 19)
            Set Rctemp = RunParametricStoredProcedure2Rec("Get_tblAcc_Sale_ByCode", Parameter2)
            
            If Rctemp.EOF <> True And Rctemp.BOF <> True Then
                .Rows = .Rows + 1
                i = i + 1
                .TextMatrix(.Rows - 1, 0) = i
                .TextMatrix(.Rows - 1, 1) = Rctemp.Fields("Kol").Value
                .TextMatrix(.Rows - 1, 2) = Rctemp.Fields("Moein").Value
                .TextMatrix(.Rows - 1, 3) = Uid_Tafsili
                .TextMatrix(.Rows - 1, 4) = Uid_Desc
                .TextMatrix(.Rows - 1, 5) = Totalprice
                .TextMatrix(.Rows - 1, 6) = 0
                Totalprice = 0
            End If
            Rst.MoveNext
       Wend
   
    End If
End With

End Sub
Private Sub UserSupplierRecieve()
'User Payment Summary  ( œ—Ì«›  «“  «„Ì‰ ò‰‰œê«‰ )
  Totalprice = 0
  Dim Uid_Temp As Integer
  Dim Uid_Desc, Uid_Tafsili As String
  Totalprice = 0
  Uid_Temp = 0
  With VsSanadView

    ReDim Parameter(3) As Parameter
    Parameter(0) = GenerateInputParameter("@RecieveType", adInteger, 4, EnumRecieveType.SupplierRecieve)
    Parameter(1) = GenerateInputParameter("@DateBefore", adVarWChar, 8, CStr(IIf(Trim(txtDate1.ClipText) = "", "", Trim(txtDate1.Text))))
    Parameter(2) = GenerateInputParameter("@DateAfter", adVarWChar, 8, CStr(IIf(Trim(txtDate2.ClipText) = "", "", Trim(txtDate2.Text))))
    Parameter(3) = GenerateInputParameter("@Uid", adInteger, 4, cmbUsers.ItemData(cmbUsers.ListIndex))
    Set Rst = RunParametricStoredProcedure2Rec("Get_UserRecieve", Parameter)
    
    If Rst.EOF <> True And Rst.BOF <> True Then
    
        Uid_Temp = Rst!Uid
        While Rst.EOF <> True
           
            Uid_Desc = "œ—Ì«›  «“  «„Ì‰ ò‰‰œê«‰" & " -- " & Rst.Fields("User_Name").Value & " -- " & Rst.Fields("Date").Value
            Uid_Tafsili = IIf(IsNull(Rst.Fields("Tafsili").Value), "", Rst.Fields("Tafsili").Value)
            
            ReDim Parameter2(0) As Parameter
            Parameter2(0) = GenerateInputParameter("@Code", adInteger, 4, EnumAccountCodes.Bestankaran)
            Set Rctemp = RunParametricStoredProcedure2Rec("Get_tblAcc_Sale_ByCode", Parameter2)
             With VsSanadView
                If Rctemp.EOF <> True And Rctemp.BOF <> True Then
                    .Rows = .Rows + 1
                    i = i + 1
                    .TextMatrix(.Rows - 1, 0) = i
                    .TextMatrix(.Rows - 1, 1) = Rctemp.Fields("Kol").Value
                    .TextMatrix(.Rows - 1, 2) = Rctemp.Fields("Moein").Value
                    .TextMatrix(.Rows - 1, 3) = IIf(IsNull(Rst.Fields("Person_Tafsili").Value), "", Rst.Fields("Person_Tafsili").Value)
                    .TextMatrix(.Rows - 1, 4) = " Å—œ«Œ   »Â ’‰œÊﬁ" & " -- " & Rst.Fields("Person_Name").Value & " -- " & Rst.Fields("Date").Value
                    .TextMatrix(.Rows - 1, 5) = 0
                    .TextMatrix(.Rows - 1, 6) = Rst.Fields("TotalBestankar").Value
                    Totalprice = Totalprice + Rst.Fields("TotalBestankar").Value
                            
                End If
            End With
            Rst.MoveNext
            If Rst.EOF <> True And Rst.BOF <> True Then
                 If Uid_Temp <> Rst!Uid Then
                     ReDim Parameter2(0) As Parameter
                     Parameter2(0) = GenerateInputParameter("@Code", adInteger, 4, 7)
                     Set Rctemp = RunParametricStoredProcedure2Rec("Get_tblAcc_Sale_ByCode", Parameter2)
                     
                     If Rctemp.EOF <> True And Rctemp.BOF <> True Then
                         .Rows = .Rows + 1
                         i = i + 1
                         .TextMatrix(.Rows - 1, 0) = i
                         .TextMatrix(.Rows - 1, 1) = Rctemp.Fields("Kol").Value
                         .TextMatrix(.Rows - 1, 2) = Rctemp.Fields("Moein").Value
                         .TextMatrix(.Rows - 1, 3) = Uid_Tafsili
                         .TextMatrix(.Rows - 1, 4) = Uid_Desc
                         .TextMatrix(.Rows - 1, 5) = Totalprice
                         .TextMatrix(.Rows - 1, 6) = 0
                         Totalprice = 0
                         Uid_Temp = Rst!Uid
                     End If
                  End If
            Else
                 ReDim Parameter2(0) As Parameter
                 Parameter2(0) = GenerateInputParameter("@Code", adInteger, 4, 7)
                 Set Rctemp = RunParametricStoredProcedure2Rec("Get_tblAcc_Sale_ByCode", Parameter2)
                 If Rctemp.EOF <> True And Rctemp.BOF <> True Then
                     .Rows = .Rows + 1
                     i = i + 1
                     .TextMatrix(.Rows - 1, 0) = i
                     .TextMatrix(.Rows - 1, 1) = Rctemp.Fields("Kol").Value
                     .TextMatrix(.Rows - 1, 2) = Rctemp.Fields("Moein").Value
                     .TextMatrix(.Rows - 1, 3) = Uid_Tafsili
                     .TextMatrix(.Rows - 1, 4) = Uid_Desc
                     .TextMatrix(.Rows - 1, 5) = Totalprice
                     .TextMatrix(.Rows - 1, 6) = 0
                     Totalprice = 0
                 End If
            End If
       Wend
   
    End If
End With
End Sub

Private Sub VsSanadView_Click()
    With VsSanadView
        If .Col > 0 And .Col < 5 Then
           .Select .Row, .Col
           .EditCell
        End If
    End With
End Sub

Private Sub VsSanadView_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
 '   With VsSanadView
 '       If .Col > 0 And .Row > 0 Then
 '           .Select .Row, .Col
 '           .EditCell
 '       End If
 '   End With

End Sub
Private Sub UserCustomerPayment()
'User Payment Summary  ( Å—œ«Œ  »Â „‘ —Ì«‰ )
  Totalprice = 0
  Dim Uid_Temp As Integer
  Dim Uid_Desc, Uid_Tafsili As String
  Totalprice = 0
  Uid_Temp = 0
  With VsSanadView

    ReDim Parameter(3) As Parameter
    Parameter(0) = GenerateInputParameter("@PaymentType", adInteger, 4, EnumPaymentType.CustomerPayment)
    Parameter(1) = GenerateInputParameter("@DateBefore", adVarWChar, 8, CStr(IIf(Trim(txtDate1.ClipText) = "", "", Trim(txtDate1.Text))))
    Parameter(2) = GenerateInputParameter("@DateAfter", adVarWChar, 8, CStr(IIf(Trim(txtDate2.ClipText) = "", "", Trim(txtDate2.Text))))
    Parameter(3) = GenerateInputParameter("@Uid", adInteger, 4, cmbUsers.ItemData(cmbUsers.ListIndex))
    Set Rst = RunParametricStoredProcedure2Rec("Get_UserPayment", Parameter)
    
    If Rst.EOF <> True And Rst.BOF <> True Then
    
        Uid_Temp = Rst!Uid
        While Rst.EOF <> True
           
            Uid_Desc = "  Å—œ«Œ  »Â „‘ —Ì«‰  " & " -- " & Rst.Fields("User_Name").Value & " -- " & Rst.Fields("Date").Value
            Uid_Tafsili = IIf(IsNull(Rst.Fields("Tafsili").Value), "", Rst.Fields("Tafsili").Value)
            
            ReDim Parameter2(0) As Parameter
            Parameter2(0) = GenerateInputParameter("@Code", adInteger, 4, EnumAccountCodes.Bedehkaran)
            Set Rctemp = RunParametricStoredProcedure2Rec("Get_tblAcc_Sale_ByCode", Parameter2)
            If Rctemp.EOF <> True And Rctemp.BOF <> True Then
                .Rows = .Rows + 1
                i = i + 1
                .TextMatrix(.Rows - 1, 0) = i
                .TextMatrix(.Rows - 1, 1) = Rctemp.Fields("Kol").Value
                .TextMatrix(.Rows - 1, 2) = Rctemp.Fields("Moein").Value
                .TextMatrix(.Rows - 1, 3) = IIf(IsNull(Rst.Fields("Person_Tafsili").Value), "", Rst.Fields("Person_Tafsili").Value)
                .TextMatrix(.Rows - 1, 4) = "œ—Ì«›  «“ ’‰œÊﬁ" & " -- " & Rst.Fields("Person_Name").Value & " -- " & Rst.Fields("Date").Value
                .TextMatrix(.Rows - 1, 5) = Rst.Fields("TotalBestankar").Value
                .TextMatrix(.Rows - 1, 6) = 0
                Totalprice = Totalprice + Rst.Fields("TotalBestankar").Value
                        
            End If
            Rst.MoveNext
            If Rst.EOF <> True And Rst.BOF <> True Then
                 If Uid_Temp <> Rst!Uid Then
                     ReDim Parameter2(0) As Parameter
                     Parameter2(0) = GenerateInputParameter("@Code", adInteger, 4, 7)
                     Set Rctemp = RunParametricStoredProcedure2Rec("Get_tblAcc_Sale_ByCode", Parameter2)
                     
                     If Rctemp.EOF <> True And Rctemp.BOF <> True Then
                         .Rows = .Rows + 1
                         i = i + 1
                         .TextMatrix(.Rows - 1, 0) = i
                         .TextMatrix(.Rows - 1, 1) = Rctemp.Fields("Kol").Value
                         .TextMatrix(.Rows - 1, 2) = Rctemp.Fields("Moein").Value
                         .TextMatrix(.Rows - 1, 3) = Uid_Tafsili
                         .TextMatrix(.Rows - 1, 4) = Uid_Desc
                         .TextMatrix(.Rows - 1, 5) = 0
                         .TextMatrix(.Rows - 1, 6) = Totalprice
                         Totalprice = 0
                         Uid_Temp = Rst!Uid
                     End If
                  End If
            Else
                 ReDim Parameter2(0) As Parameter
                 Parameter2(0) = GenerateInputParameter("@Code", adInteger, 4, 7)
                 Set Rctemp = RunParametricStoredProcedure2Rec("Get_tblAcc_Sale_ByCode", Parameter2)
                 If Rctemp.EOF <> True And Rctemp.BOF <> True Then
                     .Rows = .Rows + 1
                     i = i + 1
                     .TextMatrix(.Rows - 1, 0) = i
                     .TextMatrix(.Rows - 1, 1) = Rctemp.Fields("Kol").Value
                     .TextMatrix(.Rows - 1, 2) = Rctemp.Fields("Moein").Value
                     .TextMatrix(.Rows - 1, 3) = Uid_Tafsili
                     .TextMatrix(.Rows - 1, 4) = Uid_Desc
                     .TextMatrix(.Rows - 1, 5) = 0
                     .TextMatrix(.Rows - 1, 6) = Totalprice
                     Totalprice = 0
                 End If
            End If
       Wend
   
    End If
 End With

End Sub
''''
Private Function ValidateKol() As Boolean
    On Error GoTo ErrHandler
        Dim boolResult As Boolean
        boolResult = True
        ReDim Parameter(0) As Parameter
        With VsSanadView
            For i = 1 To .Rows - 1
                If .TextMatrix(i, 1) <> "" Then
                    Parameter(0) = GenerateInputParameter("@KolId", adInteger, 4, .TextMatrix(i, 1))
                    Set rs = RunParametricStoredProcedure2Rec("Get_tblAcc_Kol_ByID", Parameter)
                    If rs.EOF = True And rs.BOF = True Then
                        frmDisMsg.lblMessage.Caption = " ﬂ· ‰«„⁄ »— «”  "
                        frmDisMsg.Timer1.Interval = 1000
                        frmDisMsg.Timer1.Enabled = True
                        frmDisMsg.Show vbModal
                        
                        .Select i, 1
                        .EditCell
                        boolResult = False
                        Exit For
                        
                    End If
                Else
                    frmDisMsg.lblMessage.Caption = " ﬂ· ‰«„⁄ »— «”  "
                    frmDisMsg.Timer1.Interval = 1000
                    frmDisMsg.Timer1.Enabled = True
                    frmDisMsg.Show vbModal
                    .Select i, 1
                    .EditCell
                    boolResult = False
                    Exit For
                End If
            Next i
        End With
        ValidateKol = boolResult
    Exit Function
ErrHandler:
    MsgBox err.Description
    modgl.LogSave "frmCreateSanad", err, "ValidateKol"

End Function


Private Function ValidateMoein() As Boolean
     On Error GoTo ErrHandler
        Dim Result As Boolean
        Result = True
        ReDim Parameter(1) As Parameter
        With VsSanadView
            For i = 1 To .Rows - 1
                If .TextMatrix(i, 2) <> "" Then
                    Parameter(0) = GenerateInputParameter("@KolId", adInteger, 4, .TextMatrix(i, 1))
                    Parameter(1) = GenerateInputParameter("@MoeinId", adInteger, 4, .TextMatrix(i, 2))
                    Set rs = RunParametricStoredProcedure2Rec("Get_tblAcc_Moein_ByID", Parameter)
                    If rs.EOF = True And rs.BOF = True Then
                        frmDisMsg.lblMessage.Caption = " „⁄Ì‰ ‰«„⁄ »— «”  "
                        frmDisMsg.Timer1.Interval = 1000
                        frmDisMsg.Timer1.Enabled = True
                        frmDisMsg.Show vbModal
                        
                        .Select i, 2
                        .EditCell
                        Result = False
                        Exit For
                        
                    End If
                Else
                
                    frmDisMsg.lblMessage.Caption = " „⁄Ì‰ ‰«„⁄ »— «”  "
                    frmDisMsg.Timer1.Interval = 1000
                    frmDisMsg.Timer1.Enabled = True
                    frmDisMsg.Show vbModal
                    .Select i, 2
                    .EditCell
                    Result = False
                    Exit For
                End If
            Next i
        End With
        ValidateMoein = Result
    Exit Function
ErrHandler:
    MsgBox err.Description
    modgl.LogSave "frmCreateSanad", err, "ValidateMoein"
End Function

Private Function ValidateTafsili() As Boolean
     On Error GoTo ErrHandler
        Dim Result As Boolean
        Result = True
        ReDim Parameter(1) As Parameter
        With VsSanadView
            For i = 1 To .Rows - 1
                If .TextMatrix(i, 3) <> "" Then
                    Parameter(0) = GenerateInputParameter("@Branch", adInteger, 4, cboBranch.ItemData(cboBranch.ListIndex))
                    Parameter(1) = GenerateInputParameter("@TafsiliId", adInteger, 4, .TextMatrix(i, 3))
                    
                    Set rs = RunParametricStoredProcedure2Rec("Get_tblAcc_Tafsilis_ByPK_Branch_TafsiliId", Parameter)
                    If rs.EOF = True And rs.BOF = True Then
                        frmDisMsg.lblMessage.Caption = "  ›÷Ì·Ì ‰«„⁄ »— «”  "
                        frmDisMsg.Timer1.Interval = 1000
                        frmDisMsg.Timer1.Enabled = True
                        frmDisMsg.Show vbModal
                        
                        .Select i, 3
                        .EditCell
                        Result = False
                        Exit For
                        
                    End If
                Else
                    frmDisMsg.lblMessage.Caption = "  ›÷Ì·Ì ‰«„⁄ »— «”  "
                    frmDisMsg.Timer1.Interval = 1000
                    frmDisMsg.Timer1.Enabled = True
                    frmDisMsg.Show vbModal
                    .Select i, 3
                    .EditCell
                    Result = False
                    Exit For
                End If
                
            Next i
        End With
        ValidateTafsili = Result
    Exit Function
ErrHandler:
    MsgBox err.Description
    modgl.LogSave "frmCreateSanad", err, "ValidateTafsili"

End Function


Private Function SetKind() As EnumAccDocumentType
    On Error GoTo ErrHandler
        frmInput.OptionLevel(0).Caption = "ÊÌ—«Ì‘"
        frmInput.OptionLevel(1).Caption = "„Êﬁ "
        frmInput.OptionLevel(2).Caption = "œ«∆„Ì"
        frmInput.txtInput.Visible = False
        For i = 0 To frmInput.OptionLevel.Count - 1
            frmInput.OptionLevel(i).Visible = True
        Next i
        frmInput.fwlblInput.Caption = "‰Ê⁄ ”‰œ —« «‰ Œ«» ﬂ‰Ìœ"
        frmInput.fwlblInput.Visible = True
        frmInput.btnCancel.Visible = True
        frmInput.Picture1.Visible = True
        frmInput.txtInput.Visible = False
        frmInput.OptionLevel(2).Value = True
        frmInput.Show vbModal
'        For i = 0 To frmInput.OptionLevel.Count - 1
'            frmInput.OptionLevel(i).Visible = False
'        Next i
'        frmInput.fwlblInput.Caption = ""
'        frmInput.fwlblInput.Visible = False
        If mvarInput <> "" Then
            Select Case mvarInput
                Case "0"
                    SetKind = Editable
                Case "1"
                    SetKind = Temporary
                Case "2"
                    SetKind = Permanently
                End Select
        Else
            SetKind = NoDefinition
        End If
        
    Exit Function
ErrHandler:
    MsgBox err.Description
    modgl.LogSave "frmCreateSanad", err, "SetKind"
End Function
Private Sub ShowKolMoeinTafsili()
    If clsArya.ExternalAccounting = False Then Exit Sub
    On Error GoTo ErrHandler
    'If LCase(clsArya.AccountSystemName) <> "samar" Then Exit Sub
    LblKolMoeinName = ""
    LblTafsiliName = ""
    
    With VsSanadView
        ReDim Parameter(1) As Parameter
        Parameter(0) = GenerateInputParameter("@KolId", adInteger, 4, Val(.TextMatrix(.Row, 1)))
        Parameter(1) = GenerateInputParameter("@MoeinId", adInteger, 4, Val(.TextMatrix(.Row, 2)))
        Set Rst = RunParametricStoredProcedure2Rec("Get_KolNameMoeinName", Parameter)
        LblKolMoeinName.Caption = Rst!des
        Rst.Close
        Parameter(0) = GenerateInputParameter("@Branch", adInteger, 4, cboBranch.ItemData(cboBranch.ListIndex))
        Parameter(1) = GenerateInputParameter("@TafsiliId", adInteger, 4, Val(.TextMatrix(.Row, 3)))
        Set Rst = RunParametricStoredProcedure2Rec("Get_TafsiliName", Parameter)
        LblTafsiliName.Caption = Rst!des
        Rst.Close
        Set Rst = Nothing
    End With
Exit Sub
ErrHandler:
    MsgBox err.Description
    modgl.LogSave "frmCreateSanad", err, "ShowKolMoeinTafsili"

End Sub


Private Sub VsSanadView_LeaveCell()
    'ShowKolMoeinTafsili
End Sub

Private Sub VsSanadView_RowColChange()
    ShowKolMoeinTafsili
End Sub

Private Sub UpdateTransferAccounting(SanadNo As Long)
    On Error GoTo ErrHandler
    ReDim Parameter(4) As Parameter
     'All Operation Sale * ReturnSale
    Parameter(0) = GenerateInputParameter("@Branch", adInteger, 4, cboBranch.ItemData(cboBranch.ListIndex))
    Parameter(1) = GenerateInputParameter("@DateBefore", adVarWChar, 8, CStr(IIf(Trim(txtDate1.ClipText) = "", "", Trim(txtDate1.Text))))
    Parameter(2) = GenerateInputParameter("@DateAfter", adVarWChar, 8, CStr(IIf(Trim(txtDate2.ClipText) = "", "", Trim(txtDate2.Text))))
    Parameter(3) = GenerateInputParameter("@SanadNo", adInteger, 4, SanadNo)
    Parameter(4) = GenerateInputParameter("@Uid", adInteger, 4, cmbUsers.ItemData(cmbUsers.ListIndex))
    RunParametricStoredProcedure "Update_transferAccounting", Parameter
    
    Exit Sub
ErrHandler:
    MsgBox err.Description
    modgl.LogSave "frmCreateSanad", err, "UpdateTransferAccounting"
End Sub
