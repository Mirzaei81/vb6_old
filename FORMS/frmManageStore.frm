VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{7AEDC602-D94C-11D1-BB7A-00E0290EA3C9}#1.0#0"; "ResizeKit.ocx"
Object = "{75D4F767-8785-11D3-93AD-0000832EF44D}#2.14#0"; "FAST2001.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmManageStore 
   ClientHeight    =   9360
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15105
   Icon            =   "frmManageStore.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9360
   ScaleWidth      =   15105
   Begin VB.Frame Frame4 
      Caption         =   "ﬂ”— Ê «÷«›Â «‰»«—"
      BeginProperty Font 
         Name            =   "B Homa"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   23
      Top             =   2760
      Width           =   14955
      Begin VB.CommandButton StoreDataUpdate 
         Caption         =   "»Â —Ê“ —”«‰Ì „ÊÃÊœÌ ﬂ«·«Â«"
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   2880
         RightToLeft     =   -1  'True
         TabIndex        =   32
         Top             =   240
         Width           =   2295
      End
      Begin VB.CheckBox CheckNotZeroMojodi 
         Alignment       =   1  'Right Justify
         Caption         =   "ﬂ«·«Â«Ì  »« „ÊÃÊœÌ €Ì— ’›—"
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   345
         Left            =   120
         TabIndex        =   30
         Top             =   720
         Width           =   2535
      End
      Begin VB.TextBox txtBarcode 
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
         Height          =   435
         Left            =   120
         MultiLine       =   -1  'True
         RightToLeft     =   -1  'True
         TabIndex        =   29
         Top             =   240
         Width           =   1905
      End
      Begin VB.CommandButton Cmd_Kasr_Ezafe 
         Caption         =   "ﬂ”— Ê «÷«›Â"
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
         Height          =   615
         Left            =   13560
         TabIndex        =   24
         Top             =   360
         Width           =   1215
      End
      Begin MSMask.MaskEdBox txtDateTo 
         Height          =   465
         Left            =   5280
         TabIndex        =   33
         Top             =   480
         Width           =   1335
         _ExtentX        =   2355
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
      Begin VB.Label Label4 
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
         Left            =   5640
         RightToLeft     =   -1  'True
         TabIndex        =   34
         Top             =   120
         Width           =   945
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "»«—òœ"
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
         Left            =   2160
         RightToLeft     =   -1  'True
         TabIndex        =   31
         Top             =   240
         Width           =   465
      End
      Begin VB.Label LblResidBeAnbar 
         Alignment       =   1  'Right Justify
         Caption         =   "Ã„⁄ —Ì«·Ì ”‰œÂ«Ì —”Ìœ »Â «‰»«—:"
         BeginProperty Font 
            Name            =   "B Traffic"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   11880
         RightToLeft     =   -1  'True
         TabIndex        =   28
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label LblEnteghaliAzAnbar 
         Alignment       =   1  'Right Justify
         Caption         =   "Ã„⁄ —Ì«·Ì ÕÊ«·Â Â«Ì «‰ ﬁ«·Ì «“  «‰»«—:"
         BeginProperty Font 
            Name            =   "B Traffic"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   8520
         RightToLeft     =   -1  'True
         TabIndex        =   27
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label SumOfResidBeAnbar 
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
         Height          =   495
         Left            =   10080
         RightToLeft     =   -1  'True
         TabIndex        =   26
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label SumOfEnteghaliAzAnbar 
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
         Height          =   495
         Left            =   6660
         RightToLeft     =   -1  'True
         TabIndex        =   25
         Top             =   360
         Width           =   1815
      End
   End
   Begin VB.CheckBox chkActiveGood 
      Caption         =   "›ﬁÿ ﬂ«·«Â«Ì œ— ê—œ‘"
      BeginProperty Font 
         Name            =   "B Homa"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   1095
      Left            =   5280
      TabIndex        =   15
      Top             =   0
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Height          =   2655
      Left            =   120
      TabIndex        =   8
      Top             =   240
      Width           =   5535
      Begin VB.CommandButton cmbDataUpdateDifference 
         Caption         =   "»—Ê“ —”«‰Ì „€«Ì—  ﬂ«·« Â«"
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   22
         Top             =   1320
         Width           =   2295
      End
      Begin VB.CommandButton cmbTransToOtherYear 
         Caption         =   "«‰ ﬁ«· »Â ”«· œÌê—"
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3240
         RightToLeft     =   -1  'True
         TabIndex        =   21
         Top             =   1320
         Width           =   1815
      End
      Begin VB.PictureBox Picture1 
         Height          =   1215
         Left            =   120
         ScaleHeight     =   1155
         ScaleWidth      =   2235
         TabIndex        =   17
         Top             =   120
         Width           =   2295
         Begin VB.OptionButton opnCounting1 
            Alignment       =   1  'Right Justify
            Caption         =   "»— «”«” ‘„«—‘ «Ê·"
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
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   20
            Top             =   0
            Width           =   1935
         End
         Begin VB.OptionButton opnCounting2 
            Alignment       =   1  'Right Justify
            Caption         =   "»— «”«” ‘„«—‘ œÊ„"
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
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   19
            Top             =   360
            Width           =   1935
         End
         Begin VB.OptionButton opnCounting3 
            Alignment       =   1  'Right Justify
            Caption         =   "»— «”«” ‘„«—‘ ”Ê„"
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
            Left            =   120
            RightToLeft     =   -1  'True
            TabIndex        =   18
            Top             =   720
            Width           =   1935
         End
      End
      Begin VB.ComboBox cmbOtherSalMali 
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
         Left            =   2760
         RightToLeft     =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   720
         Width           =   1335
      End
      Begin VB.ComboBox cmbSalMali 
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
         Left            =   2760
         RightToLeft     =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   240
         Width           =   1335
      End
      Begin FLWCtrls.FWProgressBar FWProgressBar1 
         Height          =   375
         Left            =   120
         Top             =   2040
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   661
         BorderStyle     =   10
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "”«· „«·Ì œÌê—"
         BeginProperty Font 
            Name            =   "B Traffic"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4200
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "”«· „«·Ì Ã«—Ì"
         BeginProperty Font 
            Name            =   "B Traffic"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4200
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame3 
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
      ForeColor       =   &H8000000D&
      Height          =   1080
      Left            =   5760
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   480
      Width           =   3735
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
         Left            =   240
         RightToLeft     =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   360
         Width           =   2475
      End
   End
   Begin VB.Frame Frame28 
      Caption         =   "«‰»«—Â«"
      BeginProperty Font 
         Name            =   "B Homa"
         Size            =   11.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   1080
      Left            =   5760
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   1560
      Width           =   3735
      Begin VB.ComboBox cmbInventory 
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
         Left            =   240
         RightToLeft     =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   360
         Width           =   2475
      End
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   8880
      Top             =   1320
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
   Begin VB.ListBox lstGoodLevel1 
      BeginProperty Font 
         Name            =   "Nazanin"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2010
      Left            =   12360
      RightToLeft     =   -1  'True
      Style           =   1  'Checkbox
      TabIndex        =   1
      Top             =   840
      Width           =   2655
   End
   Begin VB.ListBox lstGoodLevel2 
      BeginProperty Font 
         Name            =   "Nazanin"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2010
      Left            =   9600
      RightToLeft     =   -1  'True
      Style           =   1  'Checkbox
      TabIndex        =   0
      Top             =   840
      Width           =   2625
   End
   Begin VSFlex7LCtl.VSFlexGrid vsGood 
      Height          =   5460
      Left            =   0
      TabIndex        =   2
      Top             =   3900
      Width           =   14985
      _cx             =   26432
      _cy             =   9631
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
      BackColor       =   -2147483624
      ForeColor       =   -2147483640
      BackColorFixed  =   16777152
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483634
      BackColorAlternate=   12648447
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   3
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   23
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmManageStore.frx":A4C2
      ScrollTrack     =   -1  'True
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
      OwnerDraw       =   5
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
   Begin FLWCtrls.FWLabel3D fwlblMode 
      Height          =   495
      Left            =   13440
      Top             =   0
      Width           =   1545
      _ExtentX        =   2725
      _ExtentY        =   873
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   178
         Weight          =   400
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
   Begin RESIZEKITLibCtl.ResizeKit ResizeKit1 
      Height          =   480
      Left            =   0
      OleObjectBlob   =   "frmManageStore.frx":A6A5
      TabIndex        =   16
      Top             =   0
      Width           =   480
   End
   Begin FLWCtrls.FWToolTip toolTipForManageStore 
      Left            =   480
      Top             =   0
      _ExtentX        =   926
      _ExtentY        =   926
      ForeColor       =   -2147483630
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "«‰»«— ê—œ«‰Ì"
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
      Left            =   5640
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   0
      Width           =   3855
   End
   Begin VB.Label lblGoodLevel1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H000040C0&
      BackStyle       =   0  'Transparent
      Caption         =   "ê—ÊÂ «’·Ì"
      BeginProperty Font 
         Name            =   "Titr"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   12720
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   480
      Width           =   2175
   End
   Begin VB.Label lblGoodLevel2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H000080FF&
      BackStyle       =   0  'Transparent
      Caption         =   "ê—ÊÂ ›—⁄Ì ò«·«Â«"
      BeginProperty Font 
         Name            =   "Titr"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   10080
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   480
      Width           =   2145
   End
End
Attribute VB_Name = "frmManageStore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private MyFormAddEditMode As EnumAddEditMode
Dim tmpTextMatrix As String
Dim Parameter() As Parameter
Dim clsDate As New clsDate
Dim i As Integer
Dim rs As New ADODB.Recordset
Dim SortItem As Integer
'„ €Ì— “Ì— »—«Ì  ‘ŒÌ’ ›‘«— œ«œÂ ‘œ‰ œﬂ„Â »Â —Ê“ —”«‰Ì „€«Ì—  ﬂ«·«Â« «” 
'Êﬁ Ì ›‘«— œ«œÂ ‘Êœ° „ €Ì— “Ì— „ﬁœ«— œ—”  —« „ÌêÌ—œ
Dim DifferencesUpdated As Boolean
    
Public Sub Find()
    
    frmFindGoods.Show vbModal
    
    i = vsGood.FindRow(mvarcode, 1, 1, True, True)
    If i > 0 Then
        vsGood.Row = i
        vsGood.ShowCell i, 0
    End If

End Sub

Public Sub ExitForm()

    Unload Me
    
End Sub

Public Sub SetFirstToolBar()

    Dim i As Integer
    
    AllButton vbOff, True
    mdifrm.Toolbar1.Buttons(13).Enabled = True   'Find
    
    mdifrm.Toolbar1.Buttons(15).Enabled = True  'printing
    mdifrm.Toolbar1.Buttons(23).Enabled = True
    mdifrm.Toolbar1.Buttons(24).Enabled = True
    mdifrm.Toolbar1.Buttons(25).Enabled = True
    mdifrm.Toolbar1.Buttons(26).Enabled = True
    mdifrm.Toolbar1.Buttons(27).Enabled = True
    
    Select Case MyFormAddEditMode
    
        Case EnumAddEditMode.ViewMode
        
            For i = 7 To 9
                mdifrm.Toolbar1.Buttons(i).Enabled = True
            Next i
            vsGood.Editable = flexEDNone
            
        Case EnumAddEditMode.AddMode
        
            mdifrm.Toolbar1.Buttons(8).Enabled = True 'enter key
            mdifrm.Toolbar1.Buttons(9).Enabled = True 'cancel key
 '           vsGood.Editable = flexEDKbdMouse
            
        Case EnumAddEditMode.EditMode
                    
            mdifrm.Toolbar1.Buttons(8).Enabled = True 'enter key
            mdifrm.Toolbar1.Buttons(9).Enabled = True 'cancel key
'            vsGood.Editable = flexEDKbdMouse
    End Select
    
    HeaderLabel Val(MyFormAddEditMode), fwlblMode
    
End Sub

Public Sub DefaultSetting()
    lstGoodLevel1.Clear
    lstGoodLevel2.Clear
    vsGood.Rows = 1
    If cmbInventory.ListIndex <> -1 And cmbBranch.ListIndex <> -1 Then
        FillLstGoodLevel1
    End If
End Sub

Public Sub FillLstGoodLevel1() ' it fills the lstGoodLevel1 using table tgoodlevel1
    Dim Rst As New ADODB.Recordset
    
    lstGoodLevel1.Clear
    lstGoodLevel2.Clear
    vsGood.Rows = 1
    
    MyFormAddEditMode = EnumAddEditMode.ViewMode
    SetFirstToolBar
    
    ReDim Parameter(1) As Parameter
    Parameter(0) = GenerateInputParameter("@Branch", adInteger, 4, cmbBranch.ItemData(cmbBranch.ListIndex))
    Parameter(1) = GenerateInputParameter("@InventoryNo", adInteger, 4, cmbInventory.ItemData(cmbInventory.ListIndex))
    Set Rst = RunParametricStoredProcedure2Rec("Get_Segment_Level1", Parameter)
        
    If (Rst.EOF = True And Rst.BOF = True) Then
        Exit Sub
    End If
    
    While Rst.EOF = False
        lstGoodLevel1.AddItem Rst.Fields("Description")
        lstGoodLevel1.ItemData(lstGoodLevel1.ListCount - 1) = Rst.Fields("Code")
        Rst.MoveNext
    Wend
    
    
    lstGoodLevel1.ListIndex = 0
    FillLstGoodLevel2
    Set Rst = Nothing
End Sub

Public Sub FillLstGoodLevel2() ' it fills the lstGoodLevel2 using table tgoodlevel2

    Dim Rst As New ADODB.Recordset
    Dim i As Integer
    Dim intSelectedItem As Integer
        
    MyFormAddEditMode = EnumAddEditMode.ViewMode
    SetFirstToolBar
    
    lstGoodLevel2.Clear
    vsGood.Rows = 1
    
    
    If lstGoodLevel1.ListIndex = -1 Then
        Set Rst = Nothing
        Exit Sub
    Else
        ReDim Parameter(0) As Parameter
        Parameter(0) = GenerateInputParameter("@Code", adInteger, 4, lstGoodLevel1.ItemData(lstGoodLevel1.ListIndex))
        
        If Rst.State <> 0 Then Rst.Close
        Set Rst = RunParametricStoredProcedure2Rec("FillLstGoodLevel2", Parameter)
        If (Rst.EOF = True And Rst.BOF = True) Then
            Set Rst = Nothing
            Exit Sub
        End If
       ' rst.moveFirst
        While Rst.EOF = False
            Select Case clsStation.Language
                Case 0
                    lstGoodLevel2.AddItem Rst.Fields("Description")
                Case 1
                    lstGoodLevel2.AddItem Rst.Fields("LatinDescription")
            End Select
            
            lstGoodLevel2.ItemData(lstGoodLevel2.ListCount - 1) = Rst.Fields("Code")
            Rst.MoveNext
        Wend
        
        Set Rst = Nothing
        lstGoodLevel2.ListIndex = 0
        FillvsGood
        
    End If
    
End Sub

Public Sub FillvsGood() 'it fills the grid using vw_Good
    
    MyFormAddEditMode = EnumAddEditMode.ViewMode 'VIEW Mode
    SetFirstToolBar
    
    vsGood.Rows = 1
    If lstGoodLevel1.ListCount < 1 Then Exit Sub
    If lstGoodLevel2.ListCount < 1 Then Exit Sub
    If cmbSalMali.ListIndex = -1 Then Exit Sub
    
    
    Dim i As Integer
    Dim j As Integer
    Dim intSelectedLevel1 As Integer
    Dim intSelectedLevel2 As Integer
    Dim strSelectedLevels As String
    Dim Rst As New ADODB.Recordset
    Dim Rst2 As New ADODB.Recordset
    
    intSelectedLevel1 = -1
    intSelectedLevel2 = -1
    
    For i = 0 To lstGoodLevel1.ListCount - 1
        If lstGoodLevel1.Selected(i) = True Then
            intSelectedLevel1 = i
        End If
    Next i
    
    strSelectedLevels = ""
    For i = 0 To lstGoodLevel2.ListCount - 1
        If lstGoodLevel2.Selected(i) = True Then
            intSelectedLevel2 = i
            strSelectedLevels = strSelectedLevels + "," + CStr(lstGoodLevel2.ItemData(i))
        End If
    Next i
    
    If Rst.State <> 0 Then Rst.Close
    Dim level1 As Integer
    If intSelectedLevel1 <> -1 And intSelectedLevel2 <> -1 Then
        level1 = lstGoodLevel1.ItemData(intSelectedLevel1)
        strSelectedLevels = Right(strSelectedLevels, Len(strSelectedLevels) - 1)
    ElseIf intSelectedLevel1 <> -1 And intSelectedLevel2 = -1 Then  'Or intSelectedLevel2 = -1
       level1 = lstGoodLevel1.ItemData(intSelectedLevel1)
       strSelectedLevels = ""
    Else
        strSelectedLevels = ""
        level1 = -1
    End If
    ReDim Parameter(10) As Parameter
    Parameter(0) = GenerateInputParameter("@Level1", adInteger, 4, level1)
    Parameter(1) = GenerateInputParameter("@strSelectedLevels", adVarWChar, 4000, strSelectedLevels)
    Parameter(2) = GenerateInputParameter("@Type", adInteger, 4, -1)
    Parameter(3) = GenerateInputParameter("@InventoryNo", adInteger, 4, cmbInventory.ItemData(cmbInventory.ListIndex))
    Parameter(4) = GenerateInputParameter("@Branch", adInteger, 4, cmbBranch.ItemData(cmbBranch.ListIndex))
    Parameter(5) = GenerateInputParameter("@AccountYear", adSmallInt, 2, Val(cmbSalMali.Text))
    Parameter(6) = GenerateInputParameter("@CheckNotZeroMojodi", adInteger, 4, CheckNotZeroMojodi.Value)
    Parameter(7) = GenerateInputParameter("@CheckFirstMojodi", adInteger, 4, 0)
    Parameter(8) = GenerateInputParameter("@CheckOrder", adInteger, 4, 0)
    Parameter(9) = GenerateInputParameter("@Flag", adInteger, 4, 1)
    Parameter(10) = GenerateInputParameter("@SortItem", adInteger, 4, 1)
    
    Set Rst = RunParametricStoredProcedure2Rec("Get_tblTotal_tGood_By_Prams", Parameter)
       
    If (Rst.EOF = True And Rst.BOF = True) Then Exit Sub
    
    With vsGood
        
        i = 1
        
        While Rst.EOF = False
            'If CheckFirstMojodi.Value = 0 Or (CheckFirstMojodi.Value = 1 And Rst.Fields("FirstMojodi").Value > 0) Then
                 .Rows = .Rows + 1
                 .TextMatrix(i, 0) = i
                 .TextMatrix(i, 1) = Rst.Fields("GoodCode").Value
                 .TextMatrix(i, 2) = Rst.Fields("Barcode").Value
                 .TextMatrix(i, 3) = Left(Rst.Fields("Name").Value, 25)
                 .TextMatrix(i, 4) = Rst.Fields("UnitDescription").Value
                ' .TextMatrix(i, 4) = Rst.Fields("TypeDescription").Value
                 If Rst.Fields("Mojodi").Value >= 0 Then
                     If Rst.Fields("Mojodi").Value <> Int(Rst.Fields("Mojodi").Value) Then
                         .TextMatrix(i, 5) = Format(Rst.Fields("Mojodi").Value, "##.000")
                         .TextMatrix(i, 5) = Val(.TextMatrix(i, 5)) ' Delete Last Zeros
                     Else
                          .TextMatrix(i, 5) = Rst.Fields("Mojodi").Value
                     End If
                 Else
                     If Rst.Fields("Mojodi").Value <> Int(Rst.Fields("Mojodi").Value) Then
                         .TextMatrix(i, 5) = -Format(Rst.Fields("Mojodi").Value, "##.000")
                         .TextMatrix(i, 5) = Val(.TextMatrix(i, 5)) & "-" ' Delete Last Zeros
                     Else
                          .TextMatrix(i, 5) = -Rst.Fields("Mojodi").Value & "-"
                     End If
                 End If
                .TextMatrix(i, 7) = IIf(IsNull(Rst!AverageBuyPrice), Rst.Fields("BuyPrice").Value, Rst!AverageBuyPrice) ' Rst.Fields("BuyPrice").Value
                .TextMatrix(i, 8) = IIf(IsNull(Rst!Counting1), "", Rst!Counting1)
                .TextMatrix(i, 9) = IIf(IsNull(Rst!Counting2), "", Rst!Counting2)
                .TextMatrix(i, 10) = IIf(IsNull(Rst!Counting3), "", Rst!Counting3)
                .TextMatrix(i, 11) = IIf(IsNull(Rst!CountDifference), "", Rst!CountDifference)
                If Val(.TextMatrix(i, 11)) <> Int(Val(.TextMatrix(i, 11))) Then
                   .TextMatrix(i, 11) = Format(Val(.TextMatrix(i, 11)), "##.000")
                   .TextMatrix(i, 11) = Val(.TextMatrix(i, 11)) ' Delete Last Zeros
                End If
                .TextMatrix(i, 12) = IIf(IsNull(Rst!CountDifference), "", Rst!CountDifference * Rst!LastSellPrice)
                If Val(.TextMatrix(i, 12)) <> Int(Val(.TextMatrix(i, 12))) Then
                  .TextMatrix(i, 12) = Format(Val(.TextMatrix(i, 12)), "##.000")
                  .TextMatrix(i, 12) = Val(.TextMatrix(i, 12)) ' Delete Last Zeros
                End If
                .TextMatrix(i, 13) = IIf(IsNull(Rst!CountDifference), "", (Rst!CountDifference) * (Rst!AverageBuyPrice))
                If Val(.TextMatrix(i, 13)) <> Int(Val(.TextMatrix(i, 13))) Then
                  .TextMatrix(i, 13) = Format(Val(.TextMatrix(i, 13)), "##.000")
                  .TextMatrix(i, 13) = Val(.TextMatrix(i, 13)) ' Delete Last Zeros
                End If
                .TextMatrix(i, 14) = IIf(IsNull(Rst!bitActiveDifference), "0", Rst!bitActiveDifference)
                 
                 i = i + 1
            'End If
            Rst.MoveNext
            
        Wend
        Set Rst = Nothing
        
        .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter
        If .Rows > 1 Then
            .Cell(flexcpAlignment, 1, 2, .Rows - 1, 2) = flexAlignRightCenter
        End If
'        .AutoSizeMode = flexAutoSizeColWidth
'        .AutoSize 0, .Cols - 1
        
    End With
        
End Sub

Public Sub BeforeUpdate()

End Sub

Public Sub Edit()
    
    With vsGood
        
 '       .Editable = flexEDKbdMouse

        MyFormAddEditMode = EnumAddEditMode.EditMode
        SetFirstToolBar
    End With
End Sub

Public Sub Update()
    
    Dim i As Integer
    Dim j As Integer
    Dim LongTemp As Integer
    Dim lngSelectedSubGroup  As Long
    
    Dim Rst As New ADODB.Recordset
    
    
    lngSelectedSubGroup = -1
    
    If MyFormAddEditMode = EnumAddEditMode.ViewMode Then Exit Sub
    
    vsGood_ValidateEdit vsGood.Row, vsGood.Col, False
    
    With vsGood
        If .Rows < 2 Then
            MyFormAddEditMode = EnumAddEditMode.ViewMode
            SetFirstToolBar
            Exit Sub
        End If
        
        For i = 1 To .Rows - 1
            .Row = i
            If InStr(.TextMatrix(i, 0), "*") > 0 Then 'new or edited records
            
''''                If ((Trim(.TextMatrix(i, 2)) = "" And Trim(.TextMatrix(i, 3)) = "") Or Trim(.TextMatrix(i, 5)) = "") Or .Cell(flexcpText, i, 8) = "" Or .Cell(flexcpText, i, 7) = "" Then
''''
''''                    Select Case clsStation.Language
''''
''''                        Case 0
''''
''''                            frmMsg.fwlblMsg.Caption = "‘„« „Ì »«Ì”  «ÿ·«⁄«  —« »ÿÊ— ò«„· Ê«—œ ‰„«ÌÌœ"
''''                            frmMsg.Fwbtn(0).Caption = "ﬁ»Ê·"
''''                        Case 1
''''
''''                            frmMsg.fwlblMsg.Caption = "You Have to complete the information"
''''                            frmMsg.Fwbtn(0).Caption = "Ok"
''''                            frmMsg.fwlblMsg.Alignment = vbLeftJustify
''''
''''                    End Select
''''
''''                    frmMsg.Fwbtn(0).ButtonType = flwButtonOk
''''                    frmMsg.Fwbtn(1).Visible = False
''''                    frmMsg.Show vbModal
''''
''''                    Exit Sub
''''
''''                End If
                
''''                If Val(.TextMatrix(i, 11)) < 0 Then     '
''''                        Select Case clsStation.Language
''''
''''                            Case 0
''''
''''                                frmMsg.fwlblMsg.Caption = "‘„« „Ì »«Ì”  ‰ﬁÿÂ ”›«—‘  —« Ê«—œ ‰„«ÌÌœ"
''''                                frmMsg.Fwbtn(0).Caption = "ﬁ»Ê·"
''''                            Case 1
''''
''''                                frmMsg.fwlblMsg.Caption = "You Have to complete the information"
''''                                frmMsg.Fwbtn(0).Caption = "Ok"
''''                                frmMsg.fwlblMsg.Alignment = vbLeftJustify
''''
''''                        End Select
''''
''''                        frmMsg.Fwbtn(0).ButtonType = flwButtonOk
''''                        frmMsg.Fwbtn(1).Visible = False
''''                        frmMsg.Show vbModal
''''
''''                        Exit Sub
''''
''''                End If
                
            End If
        Next i
        
        For j = 0 To lstGoodLevel2.ListCount - 1
            If lstGoodLevel2.Selected(j) = True Then
                lngSelectedSubGroup = j
                Exit For
            End If
        Next j

        RunNonParametricStoredProcedure "Update_bitActiveDifference"

        Select Case MyFormAddEditMode
        
            Case EnumAddEditMode.EditMode
                
                For i = 1 To .Rows - 1
                    
                    If InStr(.TextMatrix(i, 0), "*") > 0 Then 'Edited records
                            
                        ReDim Parameter(7) As Parameter
                        Parameter(0) = GenerateInputParameter("@Code", adInteger, 4, Val(Trim(.TextMatrix(i, 1))))
                        Parameter(1) = GenerateInputParameter("@IntInventoryNo", adInteger, 4, cmbInventory.ItemData(cmbInventory.ListIndex))
                        Parameter(2) = GenerateInputParameter("@Counting1", adDouble, 8, IIf(.TextMatrix(i, 8) = "", 0, Val(.TextMatrix(i, 8))))
                        Parameter(3) = GenerateInputParameter("@Counting2", adDouble, 8, IIf(.TextMatrix(i, 9) = "", 0, Val(.TextMatrix(i, 9))))
                        Parameter(4) = GenerateInputParameter("@Counting3", adDouble, 8, IIf(.TextMatrix(i, 10) = "", 0, Val(.TextMatrix(i, 10))))
                        Parameter(5) = GenerateInputParameter("@Branch", adInteger, 4, cmbBranch.ItemData(cmbBranch.ListIndex))
                        Parameter(6) = GenerateInputParameter("@AccountYear", adSmallInt, 2, Val(cmbSalMali.Text))
                        Parameter(7) = GenerateInputParameter("@bitActiveDifference", adBoolean, 1, IIf(.ValueMatrix(i, 14) <> 0, 1, 0))
                        
                        RunParametricStoredProcedure "Update_tblTotal_tInventory_Good_By_Counting", Parameter
                            
                    End If
                                        
                Next i
                
            
            End Select
            
        FillvsGood
        
    End With
    
    Set Rst = Nothing
End Sub


Public Sub Cancel()

    MyFormAddEditMode = EnumAddEditMode.ViewMode
    SetFirstToolBar
    
    FillvsGood
    
End Sub

Private Sub CheckNotZeroMojodi_Click()
    FillvsGood
End Sub

Private Sub ChkNotZeroDifference_Click()
    FillvsGood
End Sub



Private Sub cmbBranch_Click()
    FillInventory
End Sub


Private Sub cmbDataUpdateDifference_Click()
    If cmbInventory.ListIndex = -1 Then Exit Sub

   If opnCounting1.Value = 0 And opnCounting2.Value = 0 And opnCounting3.Value = 0 Then
        frmDisMsg.lblMessage = " ‘„«—‘ œ·ŒÊ«Â —« «‰ Œ«» ﬂ‰Ìœ "
        frmDisMsg.Timer1.Enabled = True
        frmDisMsg.Show vbModal
        Exit Sub
   End If
   
   If lstGoodLevel1.ListCount < 1 Then Exit Sub
    If lstGoodLevel2.ListCount < 1 Then Exit Sub
    If cmbSalMali.ListIndex = -1 Then Exit Sub
    
    
    Dim i As Integer
    Dim j As Integer
    Dim intSelectedLevel1 As Integer
    Dim intSelectedLevel2 As Integer
    Dim strSelectedLevels As String
    Dim Rst As New ADODB.Recordset
    Dim Rst2 As New ADODB.Recordset
    
    intSelectedLevel1 = -1
    intSelectedLevel2 = -1
    
    For i = 0 To lstGoodLevel1.ListCount - 1
        If lstGoodLevel1.Selected(i) = True Then
            intSelectedLevel1 = i
        End If
    Next i
    
    strSelectedLevels = ""
    For i = 0 To lstGoodLevel2.ListCount - 1
        If lstGoodLevel2.Selected(i) = True Then
            intSelectedLevel2 = i
            strSelectedLevels = strSelectedLevels + "," + CStr(lstGoodLevel2.ItemData(i))
        End If
    Next i
    
    If Rst.State <> 0 Then Rst.Close
    Dim level1 As Integer
    If intSelectedLevel1 <> -1 And intSelectedLevel2 <> -1 Then
        level1 = lstGoodLevel1.ItemData(intSelectedLevel1)
        strSelectedLevels = Right(strSelectedLevels, Len(strSelectedLevels) - 1)
    ElseIf intSelectedLevel1 <> -1 And intSelectedLevel2 = -1 Then  'Or intSelectedLevel2 = -1
       level1 = lstGoodLevel1.ItemData(intSelectedLevel1)
       strSelectedLevels = ""
    Else
        strSelectedLevels = ""
        level1 = -1
    End If
    Dim CountingNo As Integer
    If opnCounting1.Value = True Then CountingNo = 1
    If opnCounting2.Value = True Then CountingNo = 2
    If opnCounting3.Value = True Then CountingNo = 3
    
    ReDim Parameter(7) As Parameter
    Parameter(0) = GenerateInputParameter("@Level1", adInteger, 4, level1)
    Parameter(1) = GenerateInputParameter("@strSelectedLevels", adVarWChar, 4000, strSelectedLevels)
    Parameter(2) = GenerateInputParameter("@InventoryNo", adInteger, 4, cmbInventory.ItemData(cmbInventory.ListIndex))
    Parameter(3) = GenerateInputParameter("@Branch", adInteger, 4, cmbBranch.ItemData(cmbBranch.ListIndex))
    Parameter(4) = GenerateInputParameter("@AccountYear", adSmallInt, 2, Val(cmbSalMali.Text))
    Parameter(5) = GenerateInputParameter("@CheckNotZeroMojodi", adBoolean, 1, CheckNotZeroMojodi.Value)
    Parameter(6) = GenerateInputParameter("@CheckFirstMojodi", adBoolean, 1, 0)
    Parameter(7) = GenerateInputParameter("@CountingNo", adInteger, 4, CountingNo)
     
    RunParametricStoredProcedure "Update_tblTotal_tGood_By_Prams", Parameter
       
       
    FWProgressBar1.Value = 0
    
    
    
    FWProgressBar1.Value = FWProgressBar1.Value + 1
    If FWProgressBar1.Value = 100 Then
       FWProgressBar1.Value = 0
    End If
        
    DefaultSetting
    FWProgressBar1.Value = 0
    FillvsGood
    
    cmbDataUpdateDifference.Enabled = True
    'Indicates that Differences has updated and ready to add Differences to the Inventory
    DifferencesUpdated = True
        
    '„ €Ì—Â«ÌÌ »—«Ì „Õ«”»Â Ã„⁄ —Ì«·Ì „»·€ „€«Ì—  Œ—Ìœ Ê ›—Ê‘
    Dim SumOfResid As Currency
    Dim SumOfEnteghali As Currency
    
    SumOfResid = 0
    SumOfEnteghali = 0
    '=========================

    '„Õ«”»Â Ã„⁄ —Ì«·Ì „€«Ì— Â«
     'Dim i As Long
     i = 0
     With vsGood
     For i = 1 To vsGood.Rows - 1
                     '„Õ«”»Â Ã„⁄ —Ì«·Ì —”Ìœ »Â «‰»«— Ê Ã„⁄ —Ì«·Ì «‰ ﬁ«·Ì «“ «‰»«—
         If Val(.TextMatrix(i, 11)) < 0 Then
             'Ã„⁄ —Ì«·Ì «‰ ﬁ«·Ì «“ «‰»«—
             SumOfEnteghali = SumOfEnteghali + Abs(Val(.TextMatrix(i, 13))) '* Val(.TextMatrix(i, 7)
         ElseIf Val(.TextMatrix(i, 11)) > 0 Then
             'Ã„⁄ —Ì«·Ì —”Ìœ »Â «‰»«—
             SumOfResid = SumOfResid + Abs(Val(.TextMatrix(i, 13)))   ' Val(.TextMatrix(i, 6)
         End If
     Next i
     End With
     
     Me.SumOfResidBeAnbar.Caption = Format(SumOfResid, "#,##")
     Me.SumOfEnteghaliAzAnbar.Caption = Format(SumOfEnteghali, "#,##")
     
     Me.Cmd_Kasr_Ezafe.Enabled = True
     
     cmbDataUpdateDifference.Enabled = True
     frmDisMsg.lblMessage = " »Â —Ê“ —”«‰Ì «‰Ã«„ ‘œ "
     frmDisMsg.Timer1.Enabled = True
     frmDisMsg.Show vbModal

End Sub

Private Sub cmbInventory_Click()
    If cmbBranch.ListIndex = -1 Then Exit Sub
    FillLstGoodLevel1
    txtBarcode.SetFocus
End Sub

Private Sub cmbSalMali_Click()
    If AccountYear = cmbSalMali.Text Then
        txtDateTo.Text = Mid(clsDate.shamsi(Date), 3)
    Else
        txtDateTo.Text = Right(cmbSalMali.Text, 2) & "/12" & "/29"
    End If
    FillvsGood
End Sub

Private Sub cmbTransToOtherYear_Click()
     If cmbInventory.ListIndex = -1 Then Exit Sub

   If opnCounting1.Value = True Then
        s = "„ÊÃÊœÌ ‘„«—‘ «Ê·"
   ElseIf opnCounting2.Value = True Then
        s = "„ÊÃÊœÌ ‘„«—‘ œÊ„"
   ElseIf opnCounting3.Value = True Then
        s = "„ÊÃÊœÌ ‘„«—‘ ”Ê„"
   Else
        s = "„ÊÃÊœÌ ‰Â«∆Ì"
   End If
''''   If opnCounting1.Value = 0 And opnCounting2.Value = 0 And opnCounting3.Value = 0 Then
''''        frmDisMsg.lblMessage = " ‘„«—‘ œ·ŒÊ«Â —« «‰ Œ«» ﬂ‰Ìœ "
''''        frmDisMsg.Timer1.Enabled = True
''''        frmDisMsg.Show vbModal
''''        Exit Sub
''''   End If
   
        frmMsg.fwlblMsg.Caption = "¬Ì« „«Ì·Ìœ " & s & " ò«·«Â« —« »Â „ÊÃÊœÌ «Ê·ÌÂ ”«· œÌê— «‰ ﬁ«· œÂÌœ" & vbLf & "œﬁ  ﬂ‰Ìœ Â„Â «ÿ·«⁄«  ”«· ÃœÌœ œ— «‰»«— Å«ﬂ „Ì ‘Êœ "
        frmMsg.fwBtn(0).ButtonType = flwButtonOk
        frmMsg.fwBtn(0).Caption = "»·Ì"
        frmMsg.fwBtn(1).Visible = flwButtonCancel
        frmMsg.fwBtn(1).Caption = "ŒÌ—"
        frmMsg.fwBtn(1).Default = True
        frmMsg.Show vbModal
        
    If mvarMsgIdx <> 1 Then
        Exit Sub
    End If
    If lstGoodLevel1.ListCount < 1 Then Exit Sub
    If lstGoodLevel2.ListCount < 1 Then Exit Sub
    If cmbSalMali.ListIndex = -1 Then Exit Sub
    If cmbOtherSalMali.ListIndex = -1 Then
        frmDisMsg.lblMessage = " ”«· „ﬁ’œ —« «‰ Œ«» ﬂ‰Ìœ "
        frmDisMsg.Timer1.Enabled = True
        frmDisMsg.Show vbModal
        Exit Sub
    End If
    If Val(cmbSalMali.Text) = Val(cmbOtherSalMali.Text) Then
        frmDisMsg.lblMessage = " ”«·Â« »«Ìœ „ ›«Ê  »«‘œ "
        frmDisMsg.Timer1.Enabled = True
        frmDisMsg.Show vbModal
        Exit Sub
    End If
    
    Dim i As Integer
    Dim j As Integer
    Dim intSelectedLevel1 As Integer
    Dim intSelectedLevel2 As Integer
    Dim strSelectedLevels As String
    Dim Rst As New ADODB.Recordset
    Dim Rst2 As New ADODB.Recordset
    
    intSelectedLevel1 = -1
    intSelectedLevel2 = -1
    
    For i = 0 To lstGoodLevel1.ListCount - 1
        If lstGoodLevel1.Selected(i) = True Then
            intSelectedLevel1 = i
        End If
    Next i
    
    strSelectedLevels = ""
    For i = 0 To lstGoodLevel2.ListCount - 1
        If lstGoodLevel2.Selected(i) = True Then
            intSelectedLevel2 = i
            strSelectedLevels = strSelectedLevels + "," + CStr(lstGoodLevel2.ItemData(i))
        End If
    Next i
    
    If Rst.State <> 0 Then Rst.Close
    Dim level1 As Integer
    If intSelectedLevel1 <> -1 And intSelectedLevel2 <> -1 Then
        level1 = lstGoodLevel1.ItemData(intSelectedLevel1)
        strSelectedLevels = Right(strSelectedLevels, Len(strSelectedLevels) - 1)
    ElseIf intSelectedLevel1 <> -1 And intSelectedLevel2 = -1 Then  'Or intSelectedLevel2 = -1
       level1 = lstGoodLevel1.ItemData(intSelectedLevel1)
       strSelectedLevels = ""
    Else
        strSelectedLevels = ""
        level1 = -1
    End If
    Dim CountingNo As Integer
    If opnCounting1.Value = True Then
        CountingNo = 1
    ElseIf opnCounting2.Value = True Then
        CountingNo = 2
    ElseIf opnCounting3.Value = True Then
        CountingNo = 3
    Else
        CountingNo = 0
    End If
    ReDim Parameter(8) As Parameter
    Parameter(0) = GenerateInputParameter("@Level1", adInteger, 4, level1)
    Parameter(1) = GenerateInputParameter("@strSelectedLevels", adVarWChar, 4000, strSelectedLevels)
    Parameter(2) = GenerateInputParameter("@InventoryNo", adInteger, 4, cmbInventory.ItemData(cmbInventory.ListIndex))
    Parameter(3) = GenerateInputParameter("@Branch", adInteger, 4, cmbBranch.ItemData(cmbBranch.ListIndex))
    Parameter(4) = GenerateInputParameter("@AccountYear", adSmallInt, 2, Val(cmbSalMali.Text))
    Parameter(5) = GenerateInputParameter("@CheckNotZeroMojodi", adBoolean, 1, CheckNotZeroMojodi.Value)
    Parameter(6) = GenerateInputParameter("@CheckFirstMojodi", adBoolean, 1, 0)
    Parameter(7) = GenerateInputParameter("@CountingNo", adInteger, 4, CountingNo)
    Parameter(8) = GenerateInputParameter("@ToOtherAccountYear", adSmallInt, 2, Val(cmbOtherSalMali.Text))
     
    RunParametricStoredProcedure "Transport_tblTotal_tGood_By_Prams", Parameter
       
       
    FWProgressBar1.Value = 0
    
    FWProgressBar1.Value = FWProgressBar1.Value + 1
    If FWProgressBar1.Value = 100 Then
       FWProgressBar1.Value = 0
    End If
        
    DefaultSetting
    FWProgressBar1.Value = 0
    cmbDataUpdateDifference.Enabled = True
    frmDisMsg.lblMessage = " «‰ ﬁ«· »Â ”«· „«·Ì ÃœÌœ «‰Ã«„ ‘œ "
    frmDisMsg.Timer1.Enabled = True
    frmDisMsg.Show vbModal
    FillvsGood

End Sub

Private Sub Cmd_Kasr_Ezafe_Click()
        
    If intVersion <> Diamond Then ShowDisMessage " ›ﬁÿ Ê—é‰ Â«Ì «·„«” „Ì  Ê«‰‰œ «“ «Ì‰ ﬁ«»·Ì  «” ›«œÂ ﬂ‰‰œ ", 2000: Exit Sub
    
    If DifferencesUpdated = False Then Exit Sub
    
    frmMsg.fwBtn(0).Visible = True
    frmMsg.fwBtn(1).Visible = True
    frmMsg.fwBtn(0).Caption = " «ÌÌœ"
    Dim strMsg As String
    strMsg = "! ÊÃÂ: «Ì‰ ⁄„·Ì«  €Ì— ﬁ«»· »—ê‘  «” " & vbCrLf
    strMsg = strMsg & "»—«Ì ﬂ”—Ì ﬂ«·« Ìﬂ ÕÊ«·Â «“ «‰»«— Ê »—«Ì «÷«›«  ﬂ«·«Â«Ì «÷«›Â° —”Ìœ »Â «‰»«— «ÌÃ«œ „Ìê—œœ. ¬Ì« „«Ì·Ìœ «Ì‰ ﬂ«— «œ«„Â ÅÌœ« ﬂ‰œø" & vbCrLf & vbCrLf
    frmMsg.fwlblMsg.Caption = strMsg

    frmMsg.Show vbModal
    
    If mvarMsgIdx = vbNo Then
        Exit Sub
    End If
    
    Unload frmMsg
    
    frmDisMsg.lblMessage = "·ÿ›« „‰ Ÿ— »„«‰Ìœ..."
    frmDisMsg.Timer1.Enabled = True
    frmDisMsg.Show vbModal
    
    Dim i As Long
    Dim j As Long
    
    i = 0
    j = 0
    Dim TempDetailsString As String
    With vsGood
        TempDetailsString = ""
        For i = 1 To vsGood.Rows - 1
            If Val(.TextMatrix(i, 11)) > 0 Then
                TempDetailsString = GenerateDetailsString3(TempDetailsString, Abs(Val(.TextMatrix(i, 11))), .TextMatrix(i, 1), .TextMatrix(i, 7), 0, 1, "", "", cmbInventory.ItemData(cmbInventory.ListIndex), "", 1)
                If Len(TempDetailsString) >= 3500 Then
                    '«ÌÃ«œ Ê  Œ’Ì’ Å«—«„ —Â« »—«Ì ›—«ŒÊ«‰Ì —Ê«· œ—Ã —”Ìœ «‰ ﬁ«· »Â «‰»«— Ì« ÕÊ«·Â «‰ ﬁ«· «“ «‰»«—
                    InsertHavaleResid 7, TempDetailsString, "ﬂ”— Ê «÷«›Â - —”Ìœ «‰ ﬁ«·Ì »Â «‰»«—"
                    TempDetailsString = ""
                End If
            End If
        Next i
        If Len(TempDetailsString) > 0 Then InsertHavaleResid 7, TempDetailsString, "ﬂ”— Ê «÷«›Â - —”Ìœ «‰ ﬁ«·Ì »Â «‰»«—"
        TempDetailsString = ""
        For i = 1 To vsGood.Rows - 1
            If Val(.TextMatrix(i, 11)) < 0 Then
                TempDetailsString = GenerateDetailsString3(TempDetailsString, Abs(Val(.TextMatrix(i, 11))), .TextMatrix(i, 1), .TextMatrix(i, 7), 0, 1, "", "", cmbInventory.ItemData(cmbInventory.ListIndex), "", 1)
                If Len(TempDetailsString) >= 3500 Then
                    '«ÌÃ«œ Ê  Œ’Ì’ Å«—«„ —Â« »—«Ì ›—«ŒÊ«‰Ì —Ê«· œ—Ã —”Ìœ «‰ ﬁ«· »Â «‰»«— Ì« ÕÊ«·Â «‰ ﬁ«· «“ «‰»«—
                    InsertHavaleResid 6, TempDetailsString, "ﬂ”— Ê «÷«›Â - ÕÊ«·Â «‰ ﬁ«·Ì «“ «‰»«—"
                    TempDetailsString = ""
                End If
            End If
        Next i
        If Len(TempDetailsString) > 0 Then InsertHavaleResid 6, TempDetailsString, "ﬂ”— Ê «÷«›Â - ÕÊ«·Â «‰ ﬁ«·Ì «“ «‰»«—"
    
    End With
    
    
    Unload frmDisMsg
    
    frmMsg.fwBtn(0).Visible = True
    frmMsg.fwBtn(0).Caption = " «ÌÌœ"
    frmMsg.fwBtn(1).Visible = False
    frmMsg.fwlblMsg.Caption = "Â„Â „€«Ì—  Â«Ì  ‘ŒÌ’ œ«œÂ ‘œÂ œ— «‰»«— " & cmbInventory.List(cmbInventory.ListIndex) & " ﬂ”— Ê Ì« «÷«›Â ‘œ"
    frmMsg.Show vbModal
    
    '€Ì— ›⁄«· ”«“Ì œﬂ„Â ﬂ”— Ê «÷«›Â »—«Ì Ã·ÊêÌ—Ì «“ œ—Ã œÊ»«—Â ﬂ”— Ê «÷«›« 
    Me.Cmd_Kasr_Ezafe.Enabled = False
    
    FWProgressBar1.Value = 0
    ReDim Parameter(11) As Parameter

    Parameter(0) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
    Parameter(1) = GenerateInputParameter("@SystemDate", adVarWChar, 50, Mid(clsDate.shamsi(Date), 3, 8))
    Parameter(2) = GenerateInputParameter("@SystemDay", adVarWChar, 50, clsDate.Find_DayOfWeek(Weekday(Date, vbSaturday)))
    Parameter(3) = GenerateInputParameter("@SystemTime", adVarWChar, 50, Mid(Str(time), 1, 5))
    Parameter(4) = GenerateInputParameter("@DateBefore", adVarWChar, 50, cmbSalMali.Text & "/01/01")
    Parameter(5) = GenerateInputParameter("@DateAfter", adVarWChar, 50, Mid(clsDate.shamsi(Date), 3, 8))
    Parameter(6) = GenerateInputParameter("@Type", adInteger, 4, EnumGoodType.forBuySale)
    Parameter(7) = GenerateInputParameter("@InVentoryNo1", adInteger, 4, cmbInventory.ItemData(cmbInventory.ListIndex))
    Parameter(8) = GenerateInputParameter("@InVentoryNo2", adInteger, 4, cmbInventory.ItemData(cmbInventory.ListIndex))
    Parameter(9) = GenerateInputParameter("@UsePercentFlag", adInteger, 4, 0)
    Parameter(10) = GenerateInputParameter("@Branch", adInteger, 4, cmbBranch.ItemData(cmbBranch.ListIndex))
    Parameter(11) = GenerateInputParameter("@AccountYear", adSmallInt, 2, Val(cmbSalMali.Text))
    
    RunParametricStoredProcedure "Update_tblTotal_tInventory_tGood_For_Mojodi", Parameter
    ''for
    Parameter(6) = GenerateInputParameter("@Type", adInteger, 4, EnumGoodType.forBuy)
    RunParametricStoredProcedure "Update_tblTotal_tInventory_tGood_For_Mojodi", Parameter
    
    FWProgressBar1.Value = 100

'        Set Rst = RunParametricStoredProcedure2Rec("GetInventoryAtomicReport_Mojodi", Parameter )

    cmbDataUpdateDifference_Click
    FWProgressBar1.Value = 0
    
    Exit Sub
    
Err_Handler:
    Unload frmDisMsg
    MsgBox err.Description & vbCrLf & "ﬂ”— Ê «÷«›Â", vbCritical + vbOKOnly, "Œÿ«"
    Me.Cmd_Kasr_Ezafe.Enabled = False

End Sub
Private Function InsertHavaleResid(Status As Integer, TempDetailsString As String, NvcDescription As String) As Boolean
        
    Dim Update As Long
    Update = -1
    
    InsertHavaleResid = False
    ReDim Parameter(28) As Parameter

    Parameter(0) = GenerateInputParameter("@Status", adInteger, 4, Status)
    Parameter(1) = GenerateInputParameter("@Owner", adInteger, 4, -1)
    Parameter(2) = GenerateInputParameter("@Customer", adInteger, 4, 0)
    Parameter(3) = GenerateInputParameter("@DiscountTotal", adInteger, 4, 0)
    Parameter(4) = GenerateInputParameter("@CarryFeeTotal", adInteger, 4, 0)
    Parameter(5) = GenerateInputParameter("@Recursive", adInteger, 4, 0)
    Parameter(6) = GenerateInputParameter("@InCharge", adInteger, 4, 0)
    Parameter(7) = GenerateInputParameter("@FacPayment", adBoolean, 1, 0)
    Parameter(8) = GenerateInputParameter("@OrderType", adInteger, 4, 1)
    Parameter(9) = GenerateInputParameter("@StationId", adInteger, 4, clsArya.StationNo)
    Parameter(10) = GenerateInputParameter("@ServiceTotal", adDouble, 8, 0)
    Parameter(11) = GenerateInputParameter("@PackingTotal", adInteger, 4, 0)
    Parameter(12) = GenerateInputParameter("@TableNo", adInteger, 4, 0)
    Parameter(13) = GenerateInputParameter("@User", adInteger, 4, mvarCurUserNo)
    Parameter(14) = GenerateInputParameter("@Date", adVarWChar, 8, txtDateTo)
    Parameter(15) = GenerateInputParameter("@DetailsString", adVarWChar, 4000, TempDetailsString)
    Parameter(16) = GenerateInputParameter("@ds", adVarWChar, 4000, "")
    Parameter(17) = GenerateInputParameter("@Balance", adBoolean, 1, 0)
    Parameter(18) = GenerateInputParameter("@AccountYear", adSmallInt, 2, CInt(cmbSalMali.List(cmbSalMali.ListIndex)))
    Parameter(19) = GenerateInputParameter("@NvcDescription", adVarWChar, 150, Trim(NvcDescription))
    Parameter(20) = GenerateInputParameter("@HavaleNo", adInteger, 4, 0)
    Parameter(21) = GenerateInputParameter("@TempAddress", adVarWChar, 255, "")
    Parameter(22) = GenerateInputParameter("@GuestNo", adSmallInt, 2, Null)
    Parameter(23) = GenerateInputParameter("@DetailsString2", adVarWChar, 4000, "")
    Parameter(24) = GenerateInputParameter("@DetailsString3", adVarWChar, 4000, "")
    Parameter(25) = GenerateInputParameter("@DetailsString4", adVarWChar, 4000, "")
    Parameter(26) = GenerateInputParameter("@AddedTotal", adInteger, 4, 0)
    Parameter(27) = GenerateInputParameter("@Rasmi", adBoolean, 1, 0)
    Parameter(28) = GenerateOutputParameter("@LastFacMNo", adInteger, 4)
                           
    Update = RunParametricStoredProcedure("InsertFactorMasterDetails", Parameter)
    If Update = -1 Then GoTo Err_Handler
    InsertHavaleResid = True

Exit Function

Err_Handler:
    Unload frmDisMsg
    MsgBox err.Description & vbCrLf & "ﬂ”— Ê «÷«›Â", vbCritical + vbOKOnly, "Œÿ«"

End Function

Private Sub Form_Activate()
    'LblAccountYear.Caption = "”«· „«·Ì :" & CInt(AccountYear)
    
 '   frmFindGoods.Hide
    VarActForm = Me.Name
    
    ChangeLanguage
      
    txtBarcode.Text = ""
    Frame3.BackColor = Me.BackColor
    Frame28.BackColor = Me.BackColor
    
    SortItem = 1    'Code Sort
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

Private Sub Form_Load()

    If ClsFormAccess.frmManageStore = False Then
        Unload Me
        Exit Sub
    End If
    If intVersion = Min Or intVersion = Normal Then
        ShowDisMessage "«„ﬂ«‰ «‰»«—ê—œ«‰Ì œ— ‰”ŒÂ Â«Ì ÅÌ‘—› Â Ê »«·« — ÊÃÊœ œ«—œ", 1500
        Unload Me
        Exit Sub
    End If
    
    CenterTop Me
    VarActForm = Me.Name
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

    DifferencesUpdated = False
    SetToolTip
End Sub


Private Sub Form_Unload(Cancel As Integer)
    
    VarActForm = ""
    If PosConnection.State = adStateOpen Then PosConnection.Close: Set PosConnection = Nothing
    Dim i As Integer
    
    AllButton vbOff, True
    
    Unload frmFindGoods
    
    SaveSetting strMainKey, Me.Name, "Left", Me.Left
    SaveSetting strMainKey, Me.Name, "Top", Me.Top

    
End Sub



Private Sub lstGoodLevel1_Click()

    FillLstGoodLevel2
End Sub

Private Sub lstGoodLevel1_ItemCheck(Item As Integer)
    
    Dim i As Integer
    
    If lstGoodLevel1.Selected(Item) = True Then
        For i = 0 To lstGoodLevel1.ListCount - 1
            If i <> Item And lstGoodLevel1.Selected(i) = True Then
                lstGoodLevel1.Selected(i) = False
            
            End If
        Next i
    End If
    
''''    FillvsGood
''''
''''    MyFormAddEditMode = EnumAddEditMode.ViewMode
''''    SetFirstToolbar
    
End Sub

Private Sub lstGoodLevel1_Scroll()
 '   FillLstGoodLevel2
End Sub

Private Sub lstGoodLevel2_ItemCheck(Item As Integer)
    
    FillvsGood

End Sub

Public Sub ChangeLanguage()

Dim Obj As Object

    Select Case clsStation.Language    ' LCase(mdifrm.Toolbar1.Buttons(25).Key)
        
        Case English
            
            
            Me.Caption = "Counting & Move Store"
            mdifrm.Caption = clsArya.LatinCompany
            Me.RightToLeft = False
            
            For Each Obj In Me
                On Error Resume Next
                    Obj.RightToLeft = False
                On Error GoTo 0
            Next Obj
            lblGoodLevel1.Caption = "Goods Main Groups"
            lblGoodLevel2.Caption = "Goods SubGroups"
        
        Case Farsi
            
            
            Me.Caption = ""
            mdifrm.Caption = clsArya.Company
            Me.RightToLeft = True
            
            For Each Obj In Me
                On Error Resume Next
                    Obj.RightToLeft = True
                On Error GoTo 0
            Next Obj
            
            lblGoodLevel1.Caption = " ê—ÊÂ «’·Ì ò«·«Â«-»Œ‘ Â«"
            lblGoodLevel2.Caption = "ê—ÊÂ ›—⁄Ì ò«·«Â«"
            
    End Select
    
'    lstGoodLevel1.Left = Me.Width - (lstGoodLevel1.Left + lstGoodLevel1.Width)
'    lstGoodLevel2.Left = Me.Width - (lstGoodLevel2.Left + lstGoodLevel2.Width)
    
'    lblGoodLevel1.Left = Me.Width - (lblGoodLevel1.Left + lblGoodLevel1.Width)
'    lblGoodLevel2.Left = Me.Width - (lblGoodLevel2.Left + lblGoodLevel2.Width)
        
    
    With vsGood
    
        .Cols = 16
        
        Select Case clsStation.Language
            Case Farsi
                .TextMatrix(0, 0) = "—œÌ›"
                .TextMatrix(0, 1) = "òœ"
                .TextMatrix(0, 2) = "»«—òœ"
                .TextMatrix(0, 3) = "‰«„ ò«·«"
                .TextMatrix(0, 4) = "Ê«Õœ "
                .TextMatrix(0, 5) = "„ÊÃÊœÌ"
                .TextMatrix(0, 6) = "›Ì ›—Ê‘"
                .TextMatrix(0, 7) = "›Ì „Ì«‰êÌ‰"
                .TextMatrix(0, 8) = "‘„«—‘ «Ê· "
                .TextMatrix(0, 9) = "‘„«—‘ œÊ„"
                .TextMatrix(0, 10) = "‘„«—‘ ”Ê„"
                .TextMatrix(0, 11) = " ⁄œ«œ „€«Ì— "
                .TextMatrix(0, 12) = "„»·€ „€«Ì—  - ›—Ê‘"
                .TextMatrix(0, 13) = "„»·€ „€«Ì—  - Œ—Ìœ"
                .TextMatrix(0, 14) = "„€«Ì—  ›⁄«·"
                .TextMatrix(0, 15) = "     "
            
            Case English
                .TextMatrix(0, 0) = "Row"
                .TextMatrix(0, 1) = "Code"
                .TextMatrix(0, 2) = "Barcode"
                .TextMatrix(0, 3) = "Name"
                .TextMatrix(0, 4) = " Unit"
                .TextMatrix(0, 5) = "Mojodi"
                .TextMatrix(0, 6) = "Fee 1"
                .TextMatrix(0, 7) = "Buyprice"
                .TextMatrix(0, 8) = "Counting1"
                .TextMatrix(0, 9) = " Counting2"
                .TextMatrix(0, 10) = "Counting3"
                .TextMatrix(0, 11) = "CountDifference"
                .TextMatrix(0, 12) = "SumDifferenceBySellPrice"
                .TextMatrix(0, 13) = "SumDifferenceByBuyPrice"
                .TextMatrix(0, 14) = "Active"
                .TextMatrix(0, 15) = "      "
            
       End Select
       
   '     .ColSort(5) = flexSortNumericAscending + flexSortNumericDescending
        .ColAlignment(-1) = flexAlignCenterCenter
        .ColAlignment(3) = flexAlignRightCenter
        .FocusRect = flexFocusHeavy
        .ColFormat(12) = "###,###"
        .ColFormat(13) = "###,###"
        .ColDataType(14) = flexDTBoolean
      '  .ColHidden(1) = True
        '.ColHidden(4) = True
        '.ColHidden(10) = True
''        .AutoSizeMode = flexAutoSizeColWidth
''        .AutoSize 0, .Cols - 1

        .AutoSearch = flexSearchFromCursor
    
        .RowHeight(0) = .RowHeight(0) * 1.5
        For i = 0 To .Cols - 1
            .ColWidth(i) = Val(GetSetting(strMainKey, Me.Name & "_vsGood", "Col" & i))
            If .ColWidth(i) = 0 Then
                .ColWidth(i) = .Width / 10     'Row
            End If
        Next i
    
    End With
    
    FillBranch
    FillInventory
    FillSalMali
    FillOtherSalMali
    DefaultSetting
            
    SetFirstToolBar

End Sub
Private Sub FillBranch()
    Dim rctmp As New ADODB.Recordset
    cmbBranch.Clear
    Set rctmp = RunStoredProcedure2RecordSet("Get_All_Branches")
    Do While rctmp.EOF = False
        cmbBranch.AddItem rctmp!nvcBranchName
        cmbBranch.ItemData(cmbBranch.NewIndex) = rctmp!Branch
        rctmp.MoveNext
    Loop
    rctmp.Close
    'If cmbBranch.ListCount > 0 Then cmbBranch.ListIndex = 0
    

End Sub
Private Sub FillInventory()
    If cmbBranch.ListIndex = -1 Then Exit Sub
    Dim rctmp As New ADODB.Recordset
    
    cmbInventory.Clear
    ReDim Parameter(1) As Parameter
    Parameter(0) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
    Parameter(1) = GenerateInputParameter("@Branch", adInteger, 4, cmbBranch.ItemData(cmbBranch.ListIndex))
    Set rctmp = RunParametricStoredProcedure2Rec("GetInventory_Branch", Parameter)
    
    If Not (rctmp.EOF = True And rctmp.BOF = True) Then
        Do While rctmp.EOF <> True
            cmbInventory.AddItem rctmp.Fields("Description")
            cmbInventory.ItemData(cmbInventory.ListCount - 1) = Val(rctmp.Fields("InventoryNo"))
            rctmp.MoveNext
        Loop
    End If
    rctmp.Close
  '  cmbInventory.ListIndex = 0

End Sub
Private Sub FillSalMali()
    cmbSalMali.Clear
    Set rs = RunStoredProcedure2RecordSet("Get_All_tAccountYears")
    Do While rs.EOF = False
        cmbSalMali.AddItem rs!AccountYear
        rs.MoveNext
    Loop
    Dim i As Integer
    For i = 0 To cmbSalMali.ListCount - 1
        cmbSalMali.ListIndex = i
        If AccountYear = cmbSalMali.Text Then
            Exit For
        End If
    Next
    'If cmbSalMali.ListCount > 0 Then cmbSalMali.ListIndex = 0
    rs.Close
End Sub
Private Sub FillOtherSalMali()
    cmbOtherSalMali.Clear
    Set rs = RunStoredProcedure2RecordSet("Get_All_tAccountYears")
    Do While rs.EOF = False
        cmbOtherSalMali.AddItem rs!AccountYear
        rs.MoveNext
    Loop
    'If cmbSalMali.ListCount > 0 Then cmbSalMali.ListIndex = 0
    rs.Close
End Sub



Private Sub ResizeKit1_ExitResize(ByVal XScale As Double, ByVal YScale As Double)


    If formloadFlag = True Then
        SaveSetting strMainKey, Me.Name, "Height", Me.Height
        SaveSetting strMainKey, Me.Name, "Width", Me.Width
    End If

End Sub



Private Sub StoreDataUpdate_Click()
    If cmbInventory.ListIndex = -1 Then Exit Sub
    If Len(Trim(txtDateTo.ClipText)) < 6 Then
        frmDisMsg.lblMessage = "  «—ÌŒ „⁄ »— Ê«—œ ﬂ‰Ìœ "
        frmDisMsg.Timer1.Enabled = True
        frmDisMsg.Show vbModal
        Exit Sub
    End If
      '  StoreDataUpdate.Enabled = False
        FWProgressBar1.Value = 0
        ReDim Parameter(11) As Parameter
    
        Parameter(0) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
        Parameter(1) = GenerateInputParameter("@SystemDate", adVarWChar, 50, Mid(clsDate.shamsi(Date), 3, 8))
        Parameter(2) = GenerateInputParameter("@SystemDay", adVarWChar, 50, clsDate.Find_DayOfWeek(Weekday(Date, vbSaturday)))
        Parameter(3) = GenerateInputParameter("@SystemTime", adVarWChar, 50, Mid(Str(time), 1, 5))
        Parameter(4) = GenerateInputParameter("@DateBefore", adVarWChar, 50, Right(cmbSalMali.Text, 2) & "/01" & "/01")
        Parameter(5) = GenerateInputParameter("@DateAfter", adVarWChar, 50, CStr(IIf(Trim(txtDateTo.ClipText) = "", "", Trim(txtDateTo.Text))))
        Parameter(6) = GenerateInputParameter("@Type", adInteger, 4, EnumGoodType.forBuy)
        Parameter(7) = GenerateInputParameter("@InVentoryNo1", adInteger, 4, cmbInventory.ItemData(cmbInventory.ListIndex))
        Parameter(8) = GenerateInputParameter("@InVentoryNo2", adInteger, 4, cmbInventory.ItemData(cmbInventory.ListIndex))
        Parameter(9) = GenerateInputParameter("@Branch", adInteger, 4, cmbBranch.ItemData(cmbBranch.ListIndex))
        Parameter(10) = GenerateInputParameter("@UsePercentFlag", adInteger, 4, 1)
        Parameter(11) = GenerateInputParameter("@AccountYear", adSmallInt, 2, Val(cmbSalMali.Text))
        RunParametricStoredProcedure "Update_tblTotal_tInventory_tGood_For_Mojodi", Parameter
        
        
        Parameter(6) = GenerateInputParameter("@Type", adInteger, 4, EnumGoodType.forBuySale)
        Parameter(10) = GenerateInputParameter("@UsePercentFlag", adInteger, 4, 0)
        RunParametricStoredProcedure "Update_tblTotal_tInventory_tGood_For_Mojodi", Parameter
        
        DoEvents
        
        FWProgressBar1.Value = 100
        DefaultSetting
        FWProgressBar1.Value = 0
        StoreDataUpdate.Enabled = True
        frmDisMsg.lblMessage = " »Â —Ê“ —”«‰Ì «‰Ã«„ ‘œ "
        frmDisMsg.Timer1.Enabled = True
        frmDisMsg.Show vbModal

End Sub

Private Sub txtBarcode_Change()
    If Len(txtBarcode.Text) > 2 Then
    If Asc(Mid(txtBarcode.Text, Len(txtBarcode.Text) - 1, 1)) = 13 Then
        txtBarcode.Text = Left(txtBarcode.Text, Len(txtBarcode.Text) - 2)
    End If
    End If
    i = vsGood.FindRow(Trim(txtBarcode.Text), 1, 2, True, True)
    If i > 0 Then
        vsGood.Row = i
        vsGood.ShowCell i, 5
    Else
        vsGood.Row = 0
        vsGood.ShowCell 0, 0
    End If

End Sub

Private Sub txtBarcode_GotFocus()
    txtBarcode.Text = ""

End Sub
Private Sub txtBarcode_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case Shift
        Case 0
            Select Case KeyCode
                Case 13
                    vsGood.SetFocus
                   ' KeyCode = 0
                 '   txtBarcode.Text = ""
                    If i > 0 Then
                        vsGood.Row = i
                        vsGood.ShowCell i, 8
                        vsGood.Row = i
                        vsGood.Col = 8
               '         vsGood.Selec vsGood.Row, vsGood.Col
                        vsGood.EditCell
                        
                    End If
            End Select
    
    End Select

End Sub

Private Sub vsGood_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    With vsGood
        If (.TextMatrix(Row, 0) = "*" Or MyFormAddEditMode = EnumAddEditMode.EditMode) And Col > 1 And tmpTextMatrix <> .TextMatrix(Row, Col) Then
        
            If MyFormAddEditMode = EnumAddEditMode.EditMode And InStr(.TextMatrix(Row, 0), "*") = 0 Then
                .TextMatrix(Row, 0) = Trim(.TextMatrix(Row, 0)) & "*"
            End If
            
        Else

        End If
        

    End With


End Sub

Private Sub vsGood_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    For i = 0 To vsGood.Cols - 1
        SaveSetting strMainKey, Me.Name & "_vsGood", "Col" & Col, vsGood.ColWidth(Col)
    Next

End Sub

Private Sub vsGood_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    tmpTextMatrix = vsGood.TextMatrix(Row, Col)
End Sub

Private Sub vsGood_BeforeSort(ByVal Col As Long, Order As Integer)
If Col = 5 Then
    With vsGood
        For i = 1 To .Rows - 1
            If InStr(1, .TextMatrix(i, Col), "-", vbTextCompare) Then
                .TextMatrix(i, Col) = -1 * (.TextMatrix(i, Col))
                .TextMatrix(i, Col) = -1 * (.TextMatrix(i, Col))
            End If
        Next i
    End With
End If
End Sub
Private Sub vsGood_AfterSort(ByVal Col As Long, Order As Integer)
SortItem = Col
If Col = 5 Then
    With vsGood
        For i = 1 To .Rows - 1
            If InStr(1, .TextMatrix(i, Col), "-", vbTextCompare) Then
                .TextMatrix(i, Col) = -1 * (.TextMatrix(i, Col))
                .TextMatrix(i, Col) = (.TextMatrix(i, Col)) & "-"
            End If
        Next i
    End With
End If
End Sub

Private Sub vsGood_Click()
    With vsGood
        If (.TextMatrix(.Row, 0) = "*" Or MyFormAddEditMode = EnumAddEditMode.EditMode) And ((.Col < 11 And .Col > 7) Or .Col = 14) Then
               .Select .Row, .Col
               .EditCell
        End If
    
    End With

End Sub

Private Sub vsGood_KeyDown(KeyCode As Integer, Shift As Integer)
    With vsGood
        If (.TextMatrix(.Row, 0) = "*" Or MyFormAddEditMode = EnumAddEditMode.EditMode) And (.Col < 11 And .Col > 7) Then
            
               .Select .Row, .Col
               .EditCell
            
        End If
    End With
    
End Sub


Private Sub vsGood_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    With vsGood
        If KeyAscii = 39 Then KeyAscii = 0
        
        If (Col < 7 Or Col > 11) Or (IsNumeric(Chr(KeyAscii)) = False And KeyAscii = 8) Then
            
            KeyAscii = 0
            
        ElseIf IsNumeric(Chr(KeyAscii)) = False And KeyAscii <> 46 Then
            
            KeyAscii = 0
            
        ElseIf (Col < 7 Or Col > 11) Or KeyAscii = 8 Then
            
            KeyAscii = 0
            
        ElseIf MyFormAddEditMode = EditMode Then
            
            If Row > 0 And InStr(1, .TextMatrix(Row, 0), "*") = 0 Then
                .TextMatrix(Row, 0) = .TextMatrix(Row, 0) & "*"
                .TextMatrix(Row, 14) = "1"
            End If
            
        End If
        
    End With
    
End Sub


Private Sub vsGood_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsGood
        .Row = Row
        .Col = Col
    End With
End Sub

Public Sub Printing()
    
    frmInput.fwlblInput.Caption = "‰Ê⁄ ê“«—‘ "
    frmInput.OptionLevel(0).Caption = "ê“«—‘ «—“‘ „ÊÃÊœÌ"
    frmInput.OptionLevel(1).Caption = " ê“«—‘ «—“‘ „€«Ì— "
    frmInput.OptionLevel(0).Value = True
    frmInput.btnCancel.Visible = True
    frmInput.Picture1.Visible = True
    frmInput.txtInput.Visible = False
                    
    frmInput.Show vbModal
    If mvarInput = "" Then
        Exit Sub
    End If
    If lstGoodLevel1.ListCount < 1 Then Exit Sub
    If lstGoodLevel2.ListCount < 1 Then Exit Sub
    If cmbSalMali.ListIndex = -1 Then Exit Sub
    
    
    Dim i As Integer
    Dim j As Integer
    Dim intSelectedLevel1 As Integer
    Dim intSelectedLevel2 As Integer
    Dim strSelectedLevels As String
    Dim Rst As New ADODB.Recordset
    Dim Rst2 As New ADODB.Recordset
    
    intSelectedLevel1 = -1
    intSelectedLevel2 = -1
    
    For i = 0 To lstGoodLevel1.ListCount - 1
        If lstGoodLevel1.Selected(i) = True Then
            intSelectedLevel1 = i
        End If
    Next i
    
    strSelectedLevels = ""
    For i = 0 To lstGoodLevel2.ListCount - 1
        If lstGoodLevel2.Selected(i) = True Then
            intSelectedLevel2 = i
            strSelectedLevels = strSelectedLevels + "," + CStr(lstGoodLevel2.ItemData(i))
        End If
    Next i
    
    If Rst.State <> 0 Then Rst.Close
    Dim level1 As Integer
    If intSelectedLevel1 <> -1 And intSelectedLevel2 <> -1 Then
        level1 = lstGoodLevel1.ItemData(intSelectedLevel1)
        strSelectedLevels = Right(strSelectedLevels, Len(strSelectedLevels) - 1)
    ElseIf intSelectedLevel1 <> -1 And intSelectedLevel2 = -1 Then  'Or intSelectedLevel2 = -1
       level1 = lstGoodLevel1.ItemData(intSelectedLevel1)
       strSelectedLevels = ""
    Else
        strSelectedLevels = ""
        level1 = -1
    End If
    ReDim Parameter(10) As Parameter
    Parameter(0) = GenerateInputParameter("@Level1", adInteger, 4, level1)
    Parameter(1) = GenerateInputParameter("@strSelectedLevels", adVarWChar, 4000, strSelectedLevels)
    Parameter(2) = GenerateInputParameter("@Type", adInteger, 4, -1)
    Parameter(3) = GenerateInputParameter("@InventoryNo", adInteger, 4, cmbInventory.ItemData(cmbInventory.ListIndex))
    Parameter(4) = GenerateInputParameter("@Branch", adInteger, 4, cmbBranch.ItemData(cmbBranch.ListIndex))
    Parameter(5) = GenerateInputParameter("@AccountYear", adSmallInt, 2, Val(cmbSalMali.Text))
    Parameter(6) = GenerateInputParameter("@CheckNotZeroMojodi", adInteger, 4, CheckNotZeroMojodi.Value)
    Parameter(7) = GenerateInputParameter("@CheckFirstMojodi", adInteger, 4, 0)
    Parameter(8) = GenerateInputParameter("@CheckOrder", adInteger, 4, 0)
    Parameter(9) = GenerateInputParameter("@Flag", adInteger, 4, 1)
    Parameter(10) = GenerateInputParameter("@SortItem", adInteger, 4, SortItem)
    
    Set Rst = RunParametricStoredProcedure2Rec("Get_tblTotal_tGood_By_Prams", Parameter)

    If mvarInput = "0" Then
        CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\RepPriceStore_A4.rpt"
        CrystalReport1.ReportTitle = "  ê“«—‘ «—“‘ „ÊÃÊœÌ -" & cmbInventory.Text
    Else
        CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\RepPriceDifference_A4.rpt"
        CrystalReport1.ReportTitle = "  ê“«—‘ «—“‘ „€«Ì—  -" & cmbInventory.Text
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

End Sub

Public Sub SetToolTip()
    
    With toolTipForManageStore
        .BackColor = vbYellow
        .DelayTime(flwToolTipDelayDefault) = 100
        .DelayTime(flwToolTipDelayShow) = 4000
        .DelayTime(flwToolTipDelayReshow) = 1000
        .Text(Picture1) = "„ÊÃÊœÌ »—«”«” «Ì‰ ‘„«—‘ »Â „ÊÃÊœÌ «Ê·ÌÂ ”«· »⁄œ „‰ ﬁ· „Ì ‘Êœ"
        .Text(cmbTransToOtherYear) = "„ÊÃÊœÌ ”«· „«·Ì Ã«—Ì —« »—«Ì Ìﬂ «‰»«— »Â ”«· „«·Ì œÌê— «‰ ﬁ«· „Ì œÂœ"
        .Text(cmbDataUpdateDifference) = "„€«Ì—  »Ì‰ „ÊÃÊœÌ Ê „ﬁœ«— ‘„«—‘ —« ‰‘«‰ „Ì œÂœ Ê «Ì‰ „ﬁœ«— „»‰«Ì „Õ«”»Â ”‰œ ò”— Ê «÷«›Â «‰»«— „Ì »«‘œ . ›ﬁÿ ò«·«Â«ÌÌòÂ ›Ì·œ ›⁄«· ¬‰Â« «‰ Œ«» ‘œÂ »«‘œ œ— „€«Ì—  êÌ—Ì ‘—ò  „Ì ò‰‰œ"
        
    End With

End Sub

