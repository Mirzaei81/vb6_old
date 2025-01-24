VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{7AEDC602-D94C-11D1-BB7A-00E0290EA3C9}#1.0#0"; "ResizeKit.ocx"
Object = "{75D4F767-8785-11D3-93AD-0000832EF44D}#2.14#0"; "FAST2001.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmCycleStock 
   ClientHeight    =   10470
   ClientLeft      =   5235
   ClientTop       =   645
   ClientWidth     =   15015
   Icon            =   "frmCycleStock.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10470
   ScaleWidth      =   15015
   Begin VB.ComboBox cmbSalMali 
      BeginProperty Font 
         Name            =   "Nazanin"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   540
      Left            =   1320
      RightToLeft     =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   21
      Top             =   120
      Width           =   1695
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   12360
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
   Begin FLWCtrls.FWLabel3D fwlblMode 
      Height          =   495
      Left            =   13320
      Top             =   80
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
      Left            =   240
      OleObjectBlob   =   "frmCycleStock.frx":A4C2
      TabIndex        =   1
      Top             =   120
      Width           =   480
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   9735
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   14775
      _ExtentX        =   26061
      _ExtentY        =   17171
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   794
      ForeColor       =   192
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Nazanin"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   " «‰»«— ê—œ«‰Ì"
      TabPicture(0)   =   "frmCycleStock.frx":A548
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "vsGood"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      TabCaption(1)   =   " ⁄—Ì› œÊ—Â «‰»«— ê—œ«‰Ì"
      TabPicture(1)   =   "frmCycleStock.frx":A564
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lblEndDate"
      Tab(1).Control(1)=   "lblDoreDes"
      Tab(1).Control(2)=   "lblStartDate"
      Tab(1).Control(3)=   "Label3"
      Tab(1).Control(4)=   "Label4"
      Tab(1).Control(5)=   "Label7"
      Tab(1).Control(6)=   "mskEndTime"
      Tab(1).Control(7)=   "mskStartTime"
      Tab(1).Control(8)=   "mskEndDate"
      Tab(1).Control(9)=   "mskStartDate"
      Tab(1).Control(10)=   "vsCycleStock"
      Tab(1).Control(11)=   "txtDoreDes"
      Tab(1).Control(12)=   "CmbBranch2"
      Tab(1).Control(13)=   "CmdOpenCycleNo"
      Tab(1).ControlCount=   14
      Begin VB.CommandButton CmdOpenCycleNo 
         Caption         =   "»«“ ò—œ‰ œÊ—Â «‰ Œ«» ‘œÂ"
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   -72960
         TabIndex        =   32
         Top             =   2280
         Width           =   1935
      End
      Begin VB.ComboBox CmbBranch2 
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
         Left            =   -65640
         RightToLeft     =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   30
         Top             =   1080
         Width           =   2475
      End
      Begin VB.Frame Frame1 
         Height          =   2175
         Left            =   240
         TabIndex        =   11
         Top             =   600
         Width           =   14295
         Begin VB.ComboBox CmbNextCycleStock 
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
            Left            =   3240
            RightToLeft     =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   38
            Top             =   1440
            Width           =   1335
         End
         Begin VB.Frame Frame2 
            Height          =   1935
            Left            =   240
            TabIndex        =   33
            Top             =   120
            Width           =   2895
            Begin VB.PictureBox Picture1 
               Height          =   975
               Left            =   240
               ScaleHeight     =   915
               ScaleWidth      =   2355
               TabIndex        =   35
               Top             =   160
               Width           =   2415
               Begin VB.OptionButton optMojodi 
                  Alignment       =   1  'Right Justify
                  Caption         =   "»— «”«” „ÊÃÊœÌ"
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
                  Left            =   120
                  RightToLeft     =   -1  'True
                  TabIndex        =   37
                  Top             =   120
                  Value           =   -1  'True
                  Width           =   2055
               End
               Begin VB.OptionButton optMojodi 
                  Alignment       =   1  'Right Justify
                  Caption         =   "»— «”«” „ÊÃÊœÌ Ê«ﬁ⁄Ì"
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
                  Left            =   120
                  RightToLeft     =   -1  'True
                  TabIndex        =   36
                  Top             =   480
                  Width           =   2055
               End
            End
            Begin VB.CommandButton cmbTransToOtherYear 
               Caption         =   "«‰ ﬁ«· »Â œÊ—Â »⁄œÌ"
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
               Left            =   240
               TabIndex        =   34
               Top             =   1200
               Width           =   2415
            End
         End
         Begin VB.CommandButton cmbDataUpdateDifference 
            Caption         =   "»—Ê“ —”«‰Ì „€«Ì—  ò«·«Â«"
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
            Left            =   6000
            TabIndex        =   31
            Top             =   1440
            Width           =   2295
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
            Left            =   8520
            MultiLine       =   -1  'True
            RightToLeft     =   -1  'True
            TabIndex        =   27
            Top             =   1560
            Width           =   2145
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
            Height          =   960
            Left            =   11400
            RightToLeft     =   -1  'True
            TabIndex        =   25
            Top             =   120
            Width           =   2775
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
               Left            =   120
               RightToLeft     =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   26
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
            Height          =   960
            Left            =   11400
            RightToLeft     =   -1  'True
            TabIndex        =   23
            Top             =   1080
            Width           =   2775
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
               TabIndex        =   24
               Top             =   360
               Width           =   2475
            End
         End
         Begin VB.ComboBox cmbCycleStock 
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
            Left            =   8520
            RightToLeft     =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   12
            Top             =   480
            Width           =   1335
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            Caption         =   "œÊ—Â »⁄œÌ"
            BeginProperty Font 
               Name            =   "B Traffic"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   375
            Left            =   4560
            RightToLeft     =   -1  'True
            TabIndex        =   39
            Top             =   1440
            Width           =   975
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
            Left            =   10680
            RightToLeft     =   -1  'True
            TabIndex        =   28
            Top             =   1560
            Width           =   585
         End
         Begin VB.Label LblCurrentCycleOpen 
            Alignment       =   1  'Right Justify
            Caption         =   "œÊ—Â Ã«—Ì »«“ ‰Ì” "
            BeginProperty Font 
               Name            =   "B Traffic"
               Size            =   11.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   495
            Left            =   5640
            RightToLeft     =   -1  'True
            TabIndex        =   20
            Top             =   240
            Visible         =   0   'False
            Width           =   2535
         End
         Begin VB.Label LblNextCycle 
            Alignment       =   1  'Right Justify
            Caption         =   "œÊ—Â »⁄œÌ »«“ ÊÃÊœ ‰œ«—œ"
            BeginProperty Font 
               Name            =   "B Traffic"
               Size            =   11.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   375
            Left            =   5280
            RightToLeft     =   -1  'True
            TabIndex        =   19
            Top             =   960
            Visible         =   0   'False
            Width           =   2895
         End
         Begin VB.Label LblCurrentCycle 
            Alignment       =   1  'Right Justify
            Caption         =   "«Ì‰ œÊ—Â ›⁄«· «”  Ê·Ì œÊ—Â ›⁄«· Å«∆Ì‰  — «“ ¬‰ ÊÃÊœ œ«—œ"
            BeginProperty Font 
               Name            =   "B Traffic"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   495
            Left            =   3240
            RightToLeft     =   -1  'True
            TabIndex        =   18
            Top             =   600
            Visible         =   0   'False
            Width           =   5055
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "œÊ—Â Ã«—Ì"
            BeginProperty Font 
               Name            =   "B Traffic"
               Size            =   11.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   10080
            RightToLeft     =   -1  'True
            TabIndex        =   13
            Top             =   480
            Width           =   975
         End
      End
      Begin VB.TextBox txtDoreDes 
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
         Height          =   510
         Left            =   -66480
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   1800
         Width           =   3345
      End
      Begin VSFlex7LCtl.VSFlexGrid vsGood 
         Height          =   6660
         Left            =   240
         TabIndex        =   3
         Top             =   2880
         Width           =   14265
         _cx             =   25162
         _cy             =   11747
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
         BackColorFixed  =   16761024
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
         AllowUserResizing=   0
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
         FormatString    =   $"frmCycleStock.frx":A580
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
      Begin VSFlex7LCtl.VSFlexGrid vsCycleStock 
         Height          =   4515
         Left            =   -74400
         TabIndex        =   5
         Top             =   3840
         Width           =   12555
         _cx             =   22146
         _cy             =   7964
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
         BackColor       =   16777215
         ForeColor       =   -2147483640
         BackColorFixed  =   16761024
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483636
         BackColorAlternate=   16777215
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
         RowHeightMin    =   500
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmCycleStock.frx":A763
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
         ExplorerBar     =   3
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
      Begin MSMask.MaskEdBox mskStartDate 
         Height          =   495
         Left            =   -64800
         TabIndex        =   6
         Top             =   2400
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   873
         _Version        =   393216
         MaxLength       =   8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Nazanin"
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
      Begin MSMask.MaskEdBox mskEndDate 
         Height          =   495
         Left            =   -64800
         TabIndex        =   7
         Top             =   3120
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   873
         _Version        =   393216
         MaxLength       =   8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Nazanin"
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
      Begin MSMask.MaskEdBox mskStartTime 
         Height          =   495
         Left            =   -67680
         TabIndex        =   16
         Top             =   2400
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   873
         _Version        =   393216
         MaxLength       =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Nazanin"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##:##"
         PromptChar      =   " "
      End
      Begin MSMask.MaskEdBox mskEndTime 
         Height          =   495
         Left            =   -67680
         TabIndex        =   17
         Top             =   3000
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   873
         _Version        =   393216
         MaxLength       =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Nazanin"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##:##"
         PromptChar      =   " "
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "* ‘⁄»Â"
         BeginProperty Font 
            Name            =   "Nazanin"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   405
         Left            =   -63120
         RightToLeft     =   -1  'True
         TabIndex        =   29
         Top             =   1080
         Width           =   1125
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "*”«⁄  Å«Ì«‰"
         BeginProperty Font 
            Name            =   "Nazanin"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   405
         Left            =   -66960
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   3120
         Width           =   1395
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "*”«⁄  ‘—Ê⁄"
         BeginProperty Font 
            Name            =   "Nazanin"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   405
         Left            =   -66960
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   2520
         Width           =   1395
      End
      Begin VB.Label lblStartDate 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "*  «—ÌŒ ‘—Ê⁄"
         BeginProperty Font 
            Name            =   "Nazanin"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   405
         Left            =   -63360
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   2520
         Width           =   1395
      End
      Begin VB.Label lblDoreDes 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "* ‘—Õ œÊ—Â"
         BeginProperty Font 
            Name            =   "Nazanin"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   405
         Left            =   -63120
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   1800
         Width           =   1125
      End
      Begin VB.Label lblEndDate 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "*  «—ÌŒ Å«Ì«‰"
         BeginProperty Font 
            Name            =   "Nazanin"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   435
         Left            =   -63240
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   3120
         Width           =   1275
      End
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "”«· „«·Ì"
      BeginProperty Font 
         Name            =   "B Traffic"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3120
      RightToLeft     =   -1  'True
      TabIndex        =   22
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "«‰»«— ê—œ«‰Ì œÊ—Â «Ì"
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
      TabIndex        =   0
      Top             =   0
      Width           =   3855
   End
End
Attribute VB_Name = "frmCycleStock"
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
Dim intCycleStockNo As Integer

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
   If SSTab1.Tab = 0 Then
        FillCycleStock
        vsGood.Rows = 1
        If cmbInventory.ListIndex <> -1 And cmbBranch.ListIndex <> -1 Then
            FillvsGood
        End If
    ElseIf SSTab1.Tab = 1 Then
         txtDoreDes = ""
         mskStartDate.Text = Mid(clsDate.shamsi(Date), 3)   ' "  /  /  "
         mskEndDate.Text = Mid(clsDate.shamsi(Date), 3) '"  /  /  "
         mskStartTime.Text = "00:00"
         mskEndTime.Text = "23:59"
         MyFormAddEditMode = EnumAddEditMode.AddMode
         FillvsCycleStock
    End If
End Sub


Public Sub FillvsGood() 'it fills the grid using vw_Good
    
    vsGood.Rows = 1
    If cmbInventory.ListIndex = -1 Then Exit Sub
    MyFormAddEditMode = EnumAddEditMode.ViewMode 'VIEW Mode
    SetFirstToolBar
    
    If cmbCycleStock.ListIndex = -1 Then Exit Sub
    
    
    Dim i As Integer
    Dim j As Integer
    Dim Rst As New ADODB.Recordset
    Dim Rst2 As New ADODB.Recordset
    
    
    ReDim Parameter(2) As Parameter
    Parameter(0) = GenerateInputParameter("@Branch", adInteger, 4, cmbBranch.ItemData(cmbBranch.ListIndex))
    Parameter(1) = GenerateInputParameter("@IntCycleStockNo", adSmallInt, 2, cmbCycleStock.ItemData(cmbCycleStock.ListIndex))
    Parameter(2) = GenerateInputParameter("@InventoryNo", adInteger, 4, cmbInventory.ItemData(cmbInventory.ListIndex))
    
    Set Rst = RunParametricStoredProcedure2Rec("Get_tblTotal_Inventory_Good_CycleStock", Parameter)
       
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
                 .TextMatrix(i, 5) = Rst.Fields("FirstMojodi").Value
                 .TextMatrix(i, 6) = Rst.Fields("BuyAmount").Value
                 .TextMatrix(i, 7) = Rst.Fields("SaleAmount").Value
                 .TextMatrix(i, 8) = Rst.Fields("LossAmount").Value
                 .TextMatrix(i, 9) = IIf(IsNull(Rst!BuyReturnAmount), "", Rst!BuyReturnAmount)
                 .TextMatrix(i, 10) = IIf(IsNull(Rst!SaleReturnAmount), "", Rst!SaleReturnAmount)
                 .TextMatrix(i, 11) = IIf(IsNull(Rst!FromStoreAmount), "", Rst!FromStoreAmount)
                 .TextMatrix(i, 12) = IIf(IsNull(Rst!toStoreAmount), "", Rst!toStoreAmount)
                 If Rst.Fields("Mojodi").Value >= 0 Then
                     If Rst.Fields("Mojodi").Value <> Int(Rst.Fields("Mojodi").Value) Then
                         .TextMatrix(i, 13) = Format(Rst.Fields("Mojodi").Value, "##.000")
                         .TextMatrix(i, 13) = Val(.TextMatrix(i, 13)) ' Delete Last Zeros
                     Else
                          .TextMatrix(i, 13) = Rst.Fields("Mojodi").Value
                     End If
                 Else
                     If Rst.Fields("Mojodi").Value <> Int(Rst.Fields("Mojodi").Value) Then
                         .TextMatrix(i, 13) = -Format(Rst.Fields("Mojodi").Value, "##.000")
                         .TextMatrix(i, 13) = Val(.TextMatrix(i, 13)) & "-" ' Delete Last Zeros
                     Else
                          .TextMatrix(i, 13) = -Rst.Fields("Mojodi").Value & "-"
                     End If
                 End If
                 .TextMatrix(i, 14) = IIf(IsNull(Rst!RealMojodi), "", Rst!RealMojodi)
                 
                 .TextMatrix(i, 15) = IIf(IsNull(Rst!CountDifference), "", Rst!CountDifference)
                 If Val(.TextMatrix(i, 15)) <> Int(Val(.TextMatrix(i, 15))) Then
                    .TextMatrix(i, 15) = Format(Val(.TextMatrix(i, 15)), "##.000")
                    .TextMatrix(i, 15) = Val(.TextMatrix(i, 15)) ' Delete Last Zeros
                 End If
                 
                 i = i + 1
            'End If
            Rst.MoveNext
           
        Wend
        Set Rst = Nothing
        
        .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter
        If .Rows > 1 Then
            .Cell(flexcpAlignment, 1, 2, .Rows - 1, 2) = flexAlignRightCenter
        End If
        .AutoSizeMode = flexAutoSizeColWidth
        .AutoSize 0, .Cols - 1
        
    End With
        
End Sub

Public Sub BeforeUpdate()
If SSTab1.Tab = 0 Then
    With vsGood
        MyFormAddEditMode = EnumAddEditMode.EditMode
        SetFirstToolBar
    End With
End If
End Sub

Public Sub Edit()
   If SSTab1.Tab = 0 Then
        
        With vsGood
            
     '       .Editable = flexEDKbdMouse
    
            MyFormAddEditMode = EnumAddEditMode.EditMode
            SetFirstToolBar
        End With
   ElseIf SSTab1.Tab = 1 Then
        MyFormAddEditMode = EnumAddEditMode.EditMode
        SetFirstToolBar
   End If
  End Sub

Public Sub Update()
 
 If SSTab1.Tab = 0 Then
    Dim i As Integer
    Dim j As Integer
    Dim LongTemp As Integer
    Dim lngSelectedSubGroup  As Long
    
    Dim Rst As New ADODB.Recordset
    
    If cmbSalMali.ListIndex = -1 Then Exit Sub
    If cmbBranch.ListIndex = -1 Then Exit Sub
    
    lngSelectedSubGroup = -1
    
'''    If MyFormAddEditMode = EnumAddEditMode.ViewMode Then Exit Sub
    
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
            Exit For
                
            End If
        Next i
        

        Select Case MyFormAddEditMode
        
                
            Case EnumAddEditMode.EditMode
                
                ReDim Parameter(3) As Parameter
                Parameter(0) = GenerateInputParameter("@Branch", adInteger, 4, cmbBranch.ItemData(cmbBranch.ListIndex))
                Parameter(1) = GenerateInputParameter("@intCycleStockNo", adInteger, 4, cmbCycleStock.ItemData(cmbCycleStock.ListIndex))
                Parameter(2) = GenerateInputParameter("@AccountYear", adSmallInt, 2, cmbSalMali.ItemData(cmbSalMali.ListIndex))
                Parameter(3) = GenerateOutputParameter("@Result", adInteger, 4)
                Dim Result As Long
                Result = RunParametricStoredProcedure("CheckMinCycleStockNo", Parameter)
                If Result = 0 Then
                    frmMsg.fwlblMsg.Caption = "«Ì‰ œÊ—Â ›⁄«· ‰Ì”  À»  «‰Ã«„ ‰„Ì ‘Êœ ."
                    frmMsg.fwBtn(0).Visible = False
                    frmMsg.fwBtn(1).ButtonType = flwButtonOk
                    frmMsg.fwBtn(1).Caption = "ﬁ»Ê·"
                    frmMsg.Show vbModal
                    FillvsGood
                    Exit Sub
                ElseIf Result = 1 Then
                    frmMsg.fwlblMsg.Caption = " «Ì‰ œÊ—Â ›⁄«· «”  Ê·Ì œÊ—Â ›⁄«· Å«∆Ì‰  — «“ ¬‰ ÊÃÊœ œ«—œ . À»  «‰Ã«„ ‰„Ì ‘Êœ "
                    frmMsg.fwBtn(0).Visible = False
                    frmMsg.fwBtn(1).ButtonType = flwButtonOk
                    frmMsg.fwBtn(1).Caption = "ﬁ»Ê·"
                    frmMsg.Show vbModal
                    FillvsGood
                    Exit Sub
                ElseIf Result = 2 Then
                    ' It's Ok
                End If
                
                Dim Qty As Boolean
                Qty = False
                For i = 1 To .Rows - 1
                    If InStr(.TextMatrix(i, 0), "*") > 0 Then 'Edited records
                        Qty = True
                        ReDim Parameter(6) As Parameter
                        Parameter(0) = GenerateInputParameter("@Code", adInteger, 4, Val(Trim(.TextMatrix(i, 1))))
                        Parameter(1) = GenerateInputParameter("@IntInventoryNo", adInteger, 4, cmbInventory.ItemData(cmbInventory.ListIndex))
                        Parameter(2) = GenerateInputParameter("@RealMojodi", adDouble, 8, IIf(.TextMatrix(i, 14) = "", 0, Val(.TextMatrix(i, 14))))
                        Parameter(3) = GenerateInputParameter("@Branch", adInteger, 4, cmbBranch.ItemData(cmbBranch.ListIndex))
                        Parameter(4) = GenerateInputParameter("@intCycleStockNo", adInteger, 4, cmbCycleStock.ItemData(cmbCycleStock.ListIndex))
                        Parameter(5) = GenerateInputParameter("@FirstMojodi", adDouble, 8, IIf(.TextMatrix(i, 5) = "", 0, Val(.TextMatrix(i, 5))))
                        Parameter(6) = GenerateInputParameter("@AccountYear", adSmallInt, 2, cmbSalMali.ItemData(cmbSalMali.ListIndex))
                        RunParametricStoredProcedure "Update_tblTotal_Inventory_Good_CycleStock", Parameter
                            
                    End If
                                        
                Next i
                If Qty = True Then
                    ShowDisMessage "À»   €ÌÌ—«  »« „Ê›ﬁÌ  «‰Ã«„ ‘œ", 2000
                End If
            
            End Select
            
        FillvsGood
        
    End With
    
    Set Rst = Nothing
ElseIf SSTab1.Tab = 1 Then
    If MyFormAddEditMode = ViewMode Then Exit Sub
        Dim strBinBuyState As String
        Dim intBuyState As Integer
        If txtDoreDes.Text = "" Or mskStartDate.Text = "" Or mskEndDate.Text = "" Or mskStartTime.Text = "" Or mskEndTime.Text = "" Then
            frmMsg.fwlblMsg.Caption = "«ÿ·«⁄«  ·«“„ —« Å— ﬂ‰Ìœ"
            frmMsg.fwBtn(0).Visible = False
            frmMsg.fwBtn(1).ButtonType = flwButtonOk
            frmMsg.fwBtn(1).Caption = "ﬁ»Ê·"
            frmMsg.Show vbModal
            Exit Sub
        End If
        If mskStartDate.Text > mskEndDate.Text Then
            frmMsg.fwlblMsg.Caption = " «—ÌŒ Å«Ì«‰ »«Ìœ »“—ê — «“  «—ÌŒ ‘—Ê⁄ »«‘œ"
            frmMsg.fwBtn(0).Visible = False
            frmMsg.fwBtn(1).ButtonType = flwButtonOk
            frmMsg.fwBtn(1).Caption = "ﬁ»Ê·"
            frmMsg.Show vbModal
            Exit Sub
        ElseIf (mskStartDate.Text = mskEndDate.Text) And (mskStartTime.Text >= mskEndTime.Text) Then
            frmMsg.fwlblMsg.Caption = "”«⁄  Å«Ì«‰ »«Ìœ »“—ê — «“ ”«⁄  ‘—Ê⁄ »«‘œ"
            frmMsg.fwBtn(0).Visible = False
            frmMsg.fwBtn(1).ButtonType = flwButtonOk
            frmMsg.fwBtn(1).Caption = "ﬁ»Ê·"
            frmMsg.Show vbModal
            Exit Sub
        End If
        
        Select Case MyFormAddEditMode
            Case AddMode
                ReDim Parameter(7) As Parameter
                Parameter(0) = GenerateInputParameter("@nvcDescription", adVarWChar, 50, txtDoreDes.Text)
                Parameter(1) = GenerateInputParameter("@StartDateCycle", adVarWChar, 10, mskStartDate.Text)
                Parameter(2) = GenerateInputParameter("@EndDateCycle", adVarWChar, 10, mskEndDate.Text)
                Parameter(3) = GenerateInputParameter("@StartTimeCycle", adVarWChar, 10, mskStartTime.Text)
                Parameter(4) = GenerateInputParameter("@EndTimeCycle", adVarWChar, 10, mskEndTime.Text)
                Parameter(5) = GenerateInputParameter("@AccountYear", adSmallInt, 2, cmbSalMali.ItemData(cmbSalMali.ListIndex))
                Parameter(6) = GenerateInputParameter("@Branch", adInteger, 4, cmbBranch2.ItemData(cmbBranch2.ListIndex))
                Parameter(7) = GenerateOutputParameter("@intCycleStockNo", adInteger, 4)
                
                Dim LastCode As Long
                LastCode = RunParametricStoredProcedure("Insert_tblTotal_CycleStock", Parameter)
                If LastCode > 0 Then
                    frmMsg.fwlblMsg.Caption = "À»  œÊ—Â ÃœÌœ »« „Ê›ﬁÌ  «‰Ã«„ ‘œ"
                    frmMsg.fwBtn(0).Visible = False
                    frmMsg.fwBtn(1).ButtonType = flwButtonOk
                    frmMsg.fwBtn(1).Caption = "ﬁ»Ê·"
                    frmMsg.Show vbModal
                    DefaultSetting
                ElseIf LastCode = -2 Then
                    frmMsg.fwlblMsg.Caption = " «Ì‰ œÊ—Â »« œÌê— œÊ—Â Â«  œ«Œ· œ«—œ Ê  €ÌÌ—«  «‰Ã«„ ‰„Ì ‘Êœ"
                    frmMsg.fwBtn(0).Visible = False
                    frmMsg.fwBtn(1).ButtonType = flwButtonOk
                    frmMsg.fwBtn(1).Caption = "ﬁ»Ê·"
                    frmMsg.Show vbModal
                    Exit Sub
                Else
                    frmMsg.fwlblMsg.Caption = "À»  «‰Ã«„ ‰‘œ"
                    frmMsg.fwBtn(0).Visible = False
                    frmMsg.fwBtn(1).ButtonType = flwButtonOk
                    frmMsg.fwBtn(1).Caption = "ﬁ»Ê·"
                    frmMsg.Show vbModal
                    txtDoreDes.SetFocus
                    Exit Sub
                End If
                
                
            Case EditMode
            
                ReDim Parameter(8) As Parameter
                Parameter(0) = GenerateInputParameter("@nvcDescription", adVarWChar, 50, txtDoreDes.Text)
                Parameter(1) = GenerateInputParameter("@StartDateCycle", adVarWChar, 10, mskStartDate.Text)
                Parameter(2) = GenerateInputParameter("@EndDateCycle", adVarWChar, 10, mskEndDate.Text)
                Parameter(3) = GenerateInputParameter("@StartTimeCycle", adVarWChar, 10, mskStartTime.Text)
                Parameter(4) = GenerateInputParameter("@EndTimeCycle", adVarWChar, 10, mskEndTime.Text)
                Parameter(5) = GenerateInputParameter("@AccountYear", adSmallInt, 2, cmbSalMali.ItemData(cmbSalMali.ListIndex))
                Parameter(6) = GenerateInputParameter("@intCycleStockNo", adInteger, 4, intCycleStockNo)
                Parameter(7) = GenerateInputParameter("@Branch", adInteger, 4, cmbBranch2.ItemData(cmbBranch2.ListIndex))
                Parameter(8) = GenerateOutputParameter("@Updated", adInteger, 4)
                
                Dim Updated As Long
                Updated = RunParametricStoredProcedure("Update_tblTotal_CycleStock", Parameter)
                If Updated = 1 Then
                    frmMsg.fwlblMsg.Caption = " €ÌÌ—«  »« „Ê›ﬁÌ  «‰Ã«„ ‘œ"
                    frmMsg.fwBtn(0).Visible = False
                    frmMsg.fwBtn(1).ButtonType = flwButtonOk
                    frmMsg.fwBtn(1).Caption = "ﬁ»Ê·"
                    frmMsg.Show vbModal
                    DefaultSetting
                ElseIf Updated = -2 Then
                    frmMsg.fwlblMsg.Caption = " «Ì‰ œÊ—Â »« œÌê— œÊ—Â Â«  œ«Œ· œ«—œ Ê  €ÌÌ—«  «‰Ã«„ ‰„Ì ‘Êœ"
                    frmMsg.fwBtn(0).Visible = False
                    frmMsg.fwBtn(1).ButtonType = flwButtonOk
                    frmMsg.fwBtn(1).Caption = "ﬁ»Ê·"
                    frmMsg.Show vbModal
                    Exit Sub
                Else
                    frmMsg.fwlblMsg.Caption = " €ÌÌ—«  «‰Ã«„ ‰‘œ"
                    frmMsg.fwBtn(0).Visible = False
                    frmMsg.fwBtn(1).ButtonType = flwButtonOk
                    frmMsg.fwBtn(1).Caption = "ﬁ»Ê·"
                    frmMsg.Show vbModal
                    Exit Sub
                End If
    
            End Select
        
        MyFormAddEditMode = AddMode
        SetFirstToolBar
        
End If
End Sub


Public Sub Cancel()
If SSTab1.Tab = 0 Then
    MyFormAddEditMode = EnumAddEditMode.ViewMode
    SetFirstToolBar
    FillCycleStock
    FillvsGood
ElseIf SSTab1.Tab = 1 Then
    MyFormAddEditMode = EnumAddEditMode.AddMode
    SetFirstToolBar
    FillCycleStock
    DefaultSetting
    LblCurrentCycle.Visible = False
    LblNextCycle.Visible = False

End If
End Sub
Private Sub cmbBranch_Click()
    FillInventory
End Sub

Private Sub CmbBranch2_Click()
    txtDoreDes = ""
    FillvsCycleStock
End Sub

Private Sub cmbDataUpdateDifference_Click()
    If cmbBranch.ListIndex = -1 Then Exit Sub
    If cmbInventory.ListIndex = -1 Then Exit Sub
    If cmbCycleStock.ListIndex = -1 Then Exit Sub

    
    Dim i As Integer
    Dim j As Integer
    Dim Rst As New ADODB.Recordset
    Dim Rst2 As New ADODB.Recordset
    
    
    
    ReDim Parameter(6) As Parameter
    Parameter(0) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
    Parameter(1) = GenerateInputParameter("@SystemDate", adVarWChar, 50, Mid(clsDate.shamsi(Date), 3, 8))
    Parameter(2) = GenerateInputParameter("@SystemDay", adVarWChar, 50, clsDate.Find_DayOfWeek(Weekday(Date, vbSaturday)))
    Parameter(3) = GenerateInputParameter("@SystemTime", adVarWChar, 50, Mid(Str(time), 1, 5))
    Parameter(4) = GenerateInputParameter("@IntCycleStockNo", adInteger, 4, cmbCycleStock.ItemData(cmbCycleStock.ListIndex))
    Parameter(5) = GenerateInputParameter("@InventoryNo1", adInteger, 4, cmbInventory.ItemData(cmbInventory.ListIndex))
    Parameter(6) = GenerateInputParameter("@Branch", adInteger, 4, cmbBranch.ItemData(cmbBranch.ListIndex))
    
     
    RunParametricStoredProcedure "Update_tblTotal_Inventory_Good_CycleStock_formojodi", Parameter
       
       
            
'        DefaultSetting
        cmbDataUpdateDifference.Enabled = True
        frmDisMsg.lblMessage = " »Â —Ê“ —”«‰Ì «‰Ã«„ ‘œ "
        frmDisMsg.Timer1.Enabled = True
        frmDisMsg.Show vbModal
        FillvsGood

End Sub

Private Sub cmbInventory_Click()
    If cmbBranch.ListIndex = -1 Then Exit Sub
    FillvsGood
    txtBarcode.SetFocus
End Sub
Private Sub cmbCycleStock_Click()
If (cmbInventory.ListIndex <> -1 And cmbBranch.ListIndex <> -1) Then
    FillvsGood
    CmbNextCycleStock.ListIndex = 0
    CmbNextCycleStock.Enabled = False
    ReDim Parameter(3) As Parameter
    Parameter(0) = GenerateInputParameter("@Branch", adInteger, 4, cmbBranch.ItemData(cmbBranch.ListIndex))
    Parameter(1) = GenerateInputParameter("@intCycleStockNo", adInteger, 4, cmbCycleStock.ItemData(cmbCycleStock.ListIndex))
    Parameter(2) = GenerateInputParameter("@AccountYear", adSmallInt, 2, cmbSalMali.ItemData(cmbSalMali.ListIndex))
    Parameter(3) = GenerateOutputParameter("@Result", adInteger, 4)
    Dim Result As Long
    Result = RunParametricStoredProcedure("CheckCurrentCycleStockNo", Parameter)
    If Result = 0 Then
       ' frmMsg.fwlblMsg.Caption = "«Ì‰ œÊ—Â ›⁄«· ‰Ì”  ."
        cmbTransToOtherYear.Enabled = False
        cmbDataUpdateDifference.Enabled = False
        LblCurrentCycleOpen.Visible = True
        LblCurrentCycle.Visible = False
        LblNextCycle.Visible = False
    ElseIf Result = 1 Then
'        frmMsg.fwlblMsg.Caption = " «Ì‰ œÊ—Â ›⁄«· «”  Ê·Ì œÊ—Â ›⁄«· Å«∆Ì‰  — «“ ¬‰ ÊÃÊœ œ«—œ . "
        cmbTransToOtherYear.Enabled = False
        cmbDataUpdateDifference.Enabled = False
        LblCurrentCycleOpen.Visible = False
        LblCurrentCycle.Visible = True
        LblNextCycle.Visible = True
    ElseIf Result = 2 Then
        cmbTransToOtherYear.Enabled = False
        cmbDataUpdateDifference.Enabled = True
        LblCurrentCycleOpen.Visible = False
        LblCurrentCycle.Visible = False
        LblNextCycle.Visible = True
    Else
        cmbTransToOtherYear.Enabled = True
        cmbDataUpdateDifference.Enabled = True
        LblCurrentCycleOpen.Visible = False
        LblCurrentCycle.Visible = False
        LblNextCycle.Visible = False
        ReDim Parameter(1) As Parameter
        Parameter(0) = GenerateInputParameter("@Branch", adInteger, 4, cmbBranch.ItemData(cmbBranch.ListIndex))
        Parameter(1) = GenerateInputParameter("@CycleStockNo", adInteger, 4, cmbCycleStock.ItemData(cmbCycleStock.ListIndex))
        Set rctmp = RunParametricStoredProcedure2Rec("Get_tblTotal_Inventory_Good_CycleStock_Name", Parameter)
        For i = 0 To CmbNextCycleStock.ListCount - 1
            CmbNextCycleStock.ListIndex = i
            If rctmp!NextUnlockCycleStock = CmbNextCycleStock.ItemData(CmbNextCycleStock.ListIndex) Then
               Exit For
            End If
        Next i
    End If
End If
    
    
End Sub
Private Sub cmbSalMali_Click()
DefaultSetting
If cmbSalMali.ListIndex <> -1 Then FillvsCycleStock

End Sub

Private Sub cmbTransToOtherYear_Click()
    If cmbBranch.ItemData(cmbBranch.ListIndex) = -1 Then Exit Sub
    Call FrmMsgTransport.SendVariables(cmbBranch.ItemData(cmbBranch.ListIndex), cmbCycleStock.ItemData(cmbCycleStock.ListIndex))
 
    FrmMsgTransport.Show vbModal

    If mvarIndexNo = 1 Then
        ReDim Parameter(3)
        Parameter(0) = GenerateInputParameter("@Branch", adInteger, 4, cmbBranch.ItemData(cmbBranch.ListIndex))
        Parameter(1) = GenerateInputParameter("@InventoryNo1", adInteger, 4, cmbInventory.ItemData(cmbInventory.ListIndex))
        Parameter(2) = GenerateInputParameter("@CycleStockNo", adInteger, 4, cmbCycleStock.ItemData(cmbCycleStock.ListIndex))
        Parameter(3) = GenerateInputParameter("@MojodiFlag", adBoolean, 4, IIf(optMojodi(0).Value = True, 1, 0))
        Set rctmp = RunParametricStoredProcedure2Rec("Transport_tblTotal_Inventory_Good_CycleStock", Parameter)
    Else
        Exit Sub
    End If
    frmDisMsg.lblMessage = " «‰ ﬁ«· »Â œÊ—Â ÃœÌœ «‰Ã«„ ‘œ "
    frmDisMsg.Timer1.Enabled = True
    frmDisMsg.Show vbModal
    'FillCycleStock
    
End Sub

Private Sub CmdOpenCycleNo_Click()
        
    If cmbBranch2.ListIndex = -1 Then Exit Sub
    frmMsg.fwlblMsg.Caption = "¬Ì« «“ »«“ ò—œ‰ œÊ—Â «‰ Œ«» ‘œÂ „ÿ„∆‰ Â” Ìœ ø"

    frmMsg.fwBtn(0).ButtonType = flwButtonOk
    frmMsg.fwBtn(1).ButtonType = flwButtonCancel

    frmMsg.fwBtn(0).Caption = "»·Ì"
    frmMsg.fwBtn(1).Caption = "ŒÌ—"

    frmMsg.Show vbModal
    If modgl.mvarMsgIdx = vbNo Then
        Exit Sub
    End If
    
    ReDim Parameter(2) As Parameter
    
    Parameter(0) = GenerateInputParameter("@Branch", adInteger, 4, cmbBranch2.ItemData(cmbBranch2.ListIndex))
    Parameter(1) = GenerateInputParameter("@intCycleStockNo", adInteger, 4, intCycleStockNo)
    Parameter(2) = GenerateInputParameter("@AccountYear", adSmallInt, 2, cmbSalMali.ItemData(cmbSalMali.ListIndex))
    
    RunParametricStoredProcedure "Update_tblTotal_Inventory_Good_CycleStock_OpenLock", Parameter
    
    frmDisMsg.lblMessage = "œÊ—Â „Ê—œ ‰Ÿ— »«“ ‘œ "
    frmDisMsg.Timer1.Enabled = True
    frmDisMsg.Show vbModal
    
    Cancel

End Sub

Private Sub Form_Activate()
    'LblAccountYear.Caption = "”«· „«·Ì :" & CInt(AccountYear)
    
    frmFindGoods.Hide
    VarActForm = Me.Name
    MyFormAddEditMode = ViewMode
    
    ChangeLanguage
      
    txtBarcode.Text = ""
    Frame3.BackColor = Me.BackColor
    Frame28.BackColor = Me.BackColor
    
    SortItem = 1    'Code Sort
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

    If ClsFormAccess.frmCycleStock = False Then
        Unload Me
        Exit Sub
    End If
    If intVersion = Min Or intVersion = Normal Then
        ShowDisMessage "«‰»«—ê—œ«‰Ì œÊ—Â «Ì œ— ‰”ŒÂ Â«Ì ÅÌ‘—› Â Ê »«·« — ÊÃÊœ œ«—œ", 1500
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

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call PresetScreenSaver
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call PresetScreenSaver
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
        
        Case Farsi
            
            
            Me.Caption = ""
            mdifrm.Caption = clsArya.Company
            Me.RightToLeft = True
            
            For Each Obj In Me
                On Error Resume Next
                    Obj.RightToLeft = True
                On Error GoTo 0
            Next Obj
            
            
    End Select
    
    
    With vsGood
    
        .Cols = 17
        
        Select Case clsStation.Language
            Case Farsi
                .TextMatrix(0, 0) = "—œÌ›"
                .TextMatrix(0, 1) = "òœ"
                .TextMatrix(0, 2) = "»«—òœ"
                .TextMatrix(0, 3) = "‰«„ ò«·«"
                .TextMatrix(0, 4) = "Ê«Õœ "
                .TextMatrix(0, 5) = "„ «Ê·ÌÂ"
                .TextMatrix(0, 6) = "Œ—Ìœ "
                .TextMatrix(0, 7) = "›—Ê‘ "
                .TextMatrix(0, 8) = "÷«Ì⁄« "
                .TextMatrix(0, 9) = "» «“ Œ—Ìœ"
                .TextMatrix(0, 10) = "» «“ ›—Ê‘"
                .TextMatrix(0, 11) = "ÕÊ«·Â «“ «‰»«—"
                .TextMatrix(0, 12) = "—”Ìœ »Â «‰»«—"
                .TextMatrix(0, 13) = "„ÊÃÊœÌ"
                .TextMatrix(0, 14) = " ⁄œ«œ Ê«ﬁ⁄Ì "
                .TextMatrix(0, 15) = " ⁄œ«œ „€«Ì— "
                .TextMatrix(0, 16) = "     "
            
            Case English
                .TextMatrix(0, 0) = "Row"
                .TextMatrix(0, 1) = "Code"
                .TextMatrix(0, 2) = "Barcode"
                .TextMatrix(0, 3) = "Name"
                .TextMatrix(0, 4) = " Unit"
                .TextMatrix(0, 5) = "FirstStock"
                .TextMatrix(0, 6) = "Purchase"
                .TextMatrix(0, 7) = "Sale"
                .TextMatrix(0, 8) = "Losses"
                .TextMatrix(0, 9) = "PurchaseReturn"
                .TextMatrix(0, 10) = "SaleReturn"
                .TextMatrix(0, 11) = "FRomStore"
                .TextMatrix(0, 12) = "toStore"
                .TextMatrix(0, 13) = "Mojodi"
                .TextMatrix(0, 14) = "Counting"
                .TextMatrix(0, 15) = "CountDifference"
                .TextMatrix(0, 16) = "      "
            
       End Select
       
   '     .ColSort(5) = flexSortNumericAscending + flexSortNumericDescending
        .ColAlignment(-1) = flexAlignCenterCenter
        .ColAlignment(3) = flexAlignRightCenter
        .FocusRect = flexFocusHeavy
       ' .ColHidden(1) = True
        .AutoSizeMode = flexAutoSizeColWidth
        .AutoSize 0, .Cols - 1
        .AutoSearch = flexSearchFromCursor
    End With
    
    FillBranch
    FillInventory
    FillSalMali
    FillCycleStock
    FillvsCycleStock
    DefaultSetting
            
    SetFirstToolBar

End Sub
Private Sub FillBranch()
    Dim rctmp As New ADODB.Recordset
    cmbBranch.Clear
    cmbBranch2.Clear
    Set rctmp = RunStoredProcedure2RecordSet("Get_All_Branches")
    Do While rctmp.EOF = False
        cmbBranch.AddItem rctmp!nvcBranchName
        cmbBranch.ItemData(cmbBranch.NewIndex) = rctmp!Branch
        cmbBranch2.AddItem rctmp!nvcBranchName
        cmbBranch2.ItemData(cmbBranch2.NewIndex) = rctmp!Branch
        rctmp.MoveNext
    Loop
    rctmp.Close
    If cmbBranch.ListCount > 0 Then cmbBranch.ListIndex = 0
    If cmbBranch2.ListCount > 0 Then cmbBranch2.ListIndex = 0
    

End Sub
Private Sub FillInventory()
 '   If cmbBranch.ListIndex = -1 Then Exit Sub
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
    For i = 0 To cmbInventory.ListCount - 1
        If cmbInventory.ItemData(i) = clsStation.CycleStockNoDefault Then
            cmbInventory.ListIndex = i
            i = 0
            Exit For
        End If
    Next i
  '  cmbInventory.ListIndex = 0
    FillCycleStock
    FillvsGood
End Sub
Private Sub FillCycleStock()
    If cmbSalMali.ListIndex = -1 Then Exit Sub
    If cmbBranch.ListIndex = -1 Then Exit Sub
    cmbCycleStock.Clear
    ReDim Parameter(1) As Parameter
    Parameter(0) = GenerateInputParameter("@AccountYear", adSmallInt, 2, cmbSalMali.ItemData(cmbSalMali.ListIndex))
    Parameter(1) = GenerateInputParameter("@Branch", adInteger, 4, cmbBranch.ItemData(cmbBranch.ListIndex))
    Set rctmp = RunParametricStoredProcedure2Rec("Get_All_tblTotal_CycleStock", Parameter)
    
    cmbCycleStock.Clear
    CmbNextCycleStock.Clear
    CmbNextCycleStock.AddItem ""
    CmbNextCycleStock.ItemData(0) = 0
    If Not (rctmp.EOF = True And rctmp.BOF = True) Then
        Do While rctmp.EOF <> True
            cmbCycleStock.AddItem rctmp.Fields("NvcDescription")
            cmbCycleStock.ItemData(cmbCycleStock.ListCount - 1) = Val(rctmp.Fields("intCycleStockNo"))
            CmbNextCycleStock.AddItem rctmp.Fields("NvcDescription")
            CmbNextCycleStock.ItemData(CmbNextCycleStock.ListCount - 1) = Val(rctmp.Fields("intCycleStockNo"))
            rctmp.MoveNext
        Loop
    End If
    rctmp.Close
    If cmbCycleStock.ListCount = 0 Then Exit Sub
    cmbCycleStock.ListIndex = 0
    
    ReDim Parameter(1) As Parameter
    Parameter(0) = GenerateInputParameter("@Branch", adInteger, 4, cmbBranch.ItemData(cmbBranch.ListIndex))
    Parameter(1) = GenerateInputParameter("@CycleStockNo", adInteger, 4, cmbCycleStock.ItemData(cmbCycleStock.ListIndex))
    Set rctmp = RunParametricStoredProcedure2Rec("Get_tblTotal_Inventory_Good_CycleStock_Name", Parameter)
    For i = 0 To cmbCycleStock.ListCount - 1
       cmbCycleStock.ListIndex = i
       If rctmp!FirstUnlockCycleStockName = cmbCycleStock.Text Then
           Exit For
       End If
    Next
End Sub
Private Sub FillSalMali()

    cmbSalMali.Clear

    Set rs = RunStoredProcedure2RecordSet("Get_All_tAccountYears")
    Do While rs.EOF = False
        cmbSalMali.AddItem rs!AccountYear
        cmbSalMali.ItemData(cmbSalMali.ListCount - 1) = Val(rs!AccountYear)
        rs.MoveNext
    Loop
    Dim i As Integer
    For i = 0 To cmbSalMali.ListCount - 1
        cmbSalMali.ListIndex = i
        If AccountYear = cmbSalMali.Text Then
            Exit For
        End If
    Next
  
    rs.Close
End Sub


Private Sub ResizeKit1_ExitResize(ByVal XScale As Double, ByVal YScale As Double)


    If formloadFlag = True Then
        SaveSetting strMainKey, Me.Name, "Height", Me.Height
        SaveSetting strMainKey, Me.Name, "Width", Me.Width
    End If

End Sub


Private Sub SSTab1_Click(PreviousTab As Integer)
    If SSTab1.Tab = 0 Then
        MyFormAddEditMode = ViewMode
    ElseIf SSTab1.Tab = 1 Then
        MyFormAddEditMode = AddMode
    End If
    SetFirstToolBar
    Cancel
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
                        vsGood.ShowCell i, 5
                        vsGood.Row = i
                        vsGood.Col = 5
               '         vsGood.Selec vsGood.Row, vsGood.Col
                        vsGood.EditCell
                        
                    End If
            End Select
    
    End Select

End Sub



Private Sub vsCycleStock_Click()
    
    intCycleStockNo = vsCycleStock.TextMatrix(vsCycleStock.Row, 1)
    GetDataDetail
    MyFormAddEditMode = ViewMode
    SetFirstToolBar
    HeaderLabel Val(MyFormAddEditMode), Me.fwlblMode

End Sub

Private Sub vsGood_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    With vsGood
        If (.TextMatrix(Row, 0) = "*" Or MyFormAddEditMode = EnumAddEditMode.EditMode) And Col > 1 And tmpTextMatrix <> .TextMatrix(Row, Col) Then
        
            If MyFormAddEditMode = EnumAddEditMode.EditMode And InStr(.TextMatrix(Row, 0), "*") = 0 Then
                .TextMatrix(Row, 0) = Trim(.TextMatrix(Row, 0)) & "*"
            End If
            
        Else

        End If
        .AutoSizeMode = flexAutoSizeColWidth
        .AutoSize 0, .Cols - 1
        

    End With


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
        If (.TextMatrix(.Row, 0) = "*" Or MyFormAddEditMode = EnumAddEditMode.EditMode) And (.Col = 5 Or .Col = 14) Then
               .Select .Row, .Col
               .EditCell
        End If
    
    End With

End Sub

Private Sub vsGood_KeyDown(KeyCode As Integer, Shift As Integer)
    With vsGood
        If (.TextMatrix(.Row, 0) = "*" Or MyFormAddEditMode = EnumAddEditMode.EditMode) And (.Col = 5 Or .Col = 14) Then
            
               .Select .Row, .Col
               .EditCell
            
        End If
    End With
    
End Sub


Private Sub vsGood_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    With vsGood
        If KeyAscii = 39 Then KeyAscii = 0
        
        If (Col <> 5 And Col <> 14) Or (IsNumeric(Chr(KeyAscii)) = False And KeyAscii = 8) Then
            
            KeyAscii = 0
            
        ElseIf IsNumeric(Chr(KeyAscii)) = False And KeyAscii <> 46 Then
            
            KeyAscii = 0
            
        ElseIf (Col <> 5 And Col <> 14) Or KeyAscii = 8 Then
            
            KeyAscii = 0
            
        ElseIf MyFormAddEditMode = EditMode Then
            
            If Row > 0 And InStr(1, .TextMatrix(Row, 0), "*") = 0 Then
                .TextMatrix(Row, 0) = .TextMatrix(Row, 0) & "*"
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
    frmInput.OptionLevel(2).Visible = True
    frmInput.fwlblInput.Caption = "‰Ê⁄ ê“«—‘ "
    frmInput.OptionLevel(0).Caption = "ê“«—‘ «—“‘ „ÊÃÊœÌ"
    frmInput.OptionLevel(1).Caption = " ê“«—‘ «—“‘ „€«Ì— "
    frmInput.OptionLevel(2).Caption = "ê“«—‘ ⁄„·ò—œ —Ê“«‰Â"
    frmInput.OptionLevel(0).Value = True
    frmInput.btnCancel.Visible = True
    frmInput.Picture1.Visible = True
    frmInput.txtInput.Visible = False
                    
    frmInput.Show vbModal
    If mvarInput = "" Then
        Exit Sub
    End If
'''    If lstGoodLevel1.ListCount < 1 Then Exit Sub
'''    If lstGoodLevel2.ListCount < 1 Then Exit Sub
    If cmbCycleStock.ListIndex = -1 Then Exit Sub
    If cmbSalMali.ListIndex = -1 Then Exit Sub
    
    
'''    Dim i As Integer
'''    Dim j As Integer
'''    Dim intSelectedLevel1 As Integer
'''    Dim intSelectedLevel2 As Integer
'''    Dim strSelectedLevels As String
    Dim Rst As New ADODB.Recordset
'''    Dim Rst2 As New ADODB.Recordset
    
'''    intSelectedLevel1 = -1
'''    intSelectedLevel2 = -1
    
'''''    For i = 0 To lstGoodLevel1.ListCount - 1
'''''        If lstGoodLevel1.Selected(i) = True Then
'''''            intSelectedLevel1 = i
'''''        End If
'''''    Next i
'''''
'''''    strSelectedLevels = ""
'''''    For i = 0 To lstGoodLevel2.ListCount - 1
'''''        If lstGoodLevel2.Selected(i) = True Then
'''''            intSelectedLevel2 = i
'''''            strSelectedLevels = strSelectedLevels + "," + CStr(lstGoodLevel2.ItemData(i))
'''''        End If
'''''    Next i
'''''
'''''    If Rst.State <> 0 Then Rst.Close
'''''    Dim level1 As Integer
'''''    If intSelectedLevel1 <> -1 And intSelectedLevel2 <> -1 Then
'''''        level1 = lstGoodLevel1.ItemData(intSelectedLevel1)
'''''        strSelectedLevels = Right(strSelectedLevels, Len(strSelectedLevels) - 1)
'''''    ElseIf intSelectedLevel1 <> -1 And intSelectedLevel2 = -1 Then  'Or intSelectedLevel2 = -1
'''''       level1 = lstGoodLevel1.ItemData(intSelectedLevel1)
'''''       strSelectedLevels = ""
'''''    Else
'''''        strSelectedLevels = ""
'''''        level1 = -1
'''''    End If
   
    ReDim Parameter(2) As Parameter
    Parameter(0) = GenerateInputParameter("@Branch", adInteger, 4, cmbBranch.ItemData(cmbBranch.ListIndex))
    Parameter(1) = GenerateInputParameter("@IntCycleStockNo", adSmallInt, 2, cmbCycleStock.ItemData(cmbCycleStock.ListIndex))
    Parameter(2) = GenerateInputParameter("@InventoryNo", adInteger, 4, cmbInventory.ItemData(cmbInventory.ListIndex))
    
    Set Rst = RunParametricStoredProcedure2Rec("Get_tblTotal_Inventory_Good_CycleStock", Parameter)

    If mvarInput = "0" Then
        CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\RepCycleStockPrice_A4.rpt"
        CrystalReport1.ReportTitle = "  ê“«—‘ «—“‘ „ÊÃÊœÌ -" & cmbInventory.Text
    ElseIf mvarInput = "1" Then
        CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\RepCycleStocPriceDifference_A4.rpt"
        CrystalReport1.ReportTitle = "  ê“«—‘ «—“‘ „€«Ì—  -" & cmbInventory.Text
    ElseIf mvarInput = "2" Then
        CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\RepDailyCycleStoc_A4.rpt"
        CrystalReport1.ReportTitle = "  ê“«—‘ ⁄„·ò—œ —Ê“«‰Â -" & cmbInventory.Text
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

Private Sub FillvsCycleStock()
    
    If cmbSalMali.ListIndex = -1 Then Exit Sub
    If cmbBranch.ListIndex = -1 Then Exit Sub
    Dim Rst As New ADODB.Recordset
    ReDim Parameter(1) As Parameter
    
    
    Parameter(0) = GenerateInputParameter("@AccountYear", adSmallInt, 2, cmbSalMali.ItemData(cmbSalMali.ListIndex))
    Parameter(1) = GenerateInputParameter("@Branch", adInteger, 4, cmbBranch2.ItemData(cmbBranch2.ListIndex))
    Set Rst = RunParametricStoredProcedure2Rec("Get_All_tblTotal_CycleStock", Parameter)
    
    With vsCycleStock
        .Rows = 1
        i = 0
        While Rst.EOF <> True
            .Rows = .Rows + 1
            i = .Rows - 1
            .TextMatrix(i, 0) = i
            .TextMatrix(i, 1) = Rst!intCycleStockNo
            .TextMatrix(i, 2) = Rst!NvcDescription
            .TextMatrix(i, 3) = Rst!StartDateCycle
            .TextMatrix(i, 4) = Rst!StartTimeCycle
            .TextMatrix(i, 5) = Rst!EndDateCycle
            .TextMatrix(i, 6) = Rst!EndTimeCycle
            .TextMatrix(i, 7) = Rst!AccountYear
            .TextMatrix(i, 8) = Not (Rst!IsLock)
            Rst.MoveNext
        Wend
    End With
    Set Rst = Nothing
End Sub
Sub GetDataDetail()
    
    If cmbBranch.ListIndex = -1 Then Exit Sub
    DefaultSetting
    
    Dim TempStr As String
    
    ReDim Parameter(1) As Parameter
    Parameter(0) = GenerateInputParameter("@intCycleStockNo", adInteger, 4, intCycleStockNo)
    Parameter(1) = GenerateInputParameter("@Branch", adInteger, 4, cmbBranch2.ItemData(cmbBranch2.ListIndex))
    Set rctmp = RunParametricStoredProcedure2Rec("Get_tblTotal_CycleStock_intCycleStockNo", Parameter)
    Dim ii As Integer
    If Not (rctmp.BOF Or rctmp.EOF) Then
        
            txtDoreDes.Text = rctmp!NvcDescription
            mskStartDate.Text = rctmp!StartDateCycle
            mskEndDate.Text = rctmp!EndDateCycle
            mskStartTime.Text = rctmp!StartTimeCycle
            mskEndTime.Text = rctmp!EndTimeCycle
           For i = 0 To cmbSalMali.ListCount - 1
                If rctmp.Fields("AccountYear").Value = cmbSalMali.ItemData(i) Then
                    cmbSalMali.ListIndex = i
                    Exit For
                End If
            Next i
''            cmbSalMali.ItemData(= rctmp!AccountYear
          
    End If
    rctmp.Close
    
    
End Sub

