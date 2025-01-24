VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{7AEDC602-D94C-11D1-BB7A-00E0290EA3C9}#1.0#0"; "ResizeKit.ocx"
Object = "{75D4F767-8785-11D3-93AD-0000832EF44D}#2.14#0"; "FAST2001.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmGood 
   ClientHeight    =   10335
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14985
   Icon            =   "frmGood.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   10335
   ScaleWidth      =   14985
   Begin TabDlg.SSTab SSTab1 
      Height          =   9615
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   14895
      _ExtentX        =   26273
      _ExtentY        =   16960
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   706
      WordWrap        =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Nazanin"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "ê—ÊÂ »‰œÌ ò«·« Â«"
      TabPicture(0)   =   "frmGood.frx":A4C2
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(1)=   "Frame3"
      Tab(0).Control(2)=   "Frame2"
      Tab(0).ControlCount=   3
      TabCaption(1)   =   " ⁄—Ì› ò«·«"
      TabPicture(1)   =   "frmGood.frx":A4DE
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "LblNumberOfRecords"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label2"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label1"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "lblGoodLevel2"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "lblGoodLevel1"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Image1"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "vsGood"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "fwBtnCustFind"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "txtBarcode"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "lstGoodLevel2"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "lstGoodLevel1"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "UcFont1"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "cmdUpdateBuyPrice"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).ControlCount=   13
      Begin VB.CommandButton cmdUpdateBuyPrice 
         Caption         =   "»Â —Ê“ —”«‰Ì ›Ì Œ—Ìœ »« ¬Œ—Ì‰ ﬁÌ„  Œ—Ìœ"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   2280
         RightToLeft     =   -1  'True
         TabIndex        =   35
         Top             =   1800
         Width           =   2415
      End
      Begin VB.Frame Frame2 
         Caption         =   " €ÌÌ— òœ ¬Ì „ «‰ Œ«»Ì - ê—ÊÂ ›—⁄Ì"
         BeginProperty Font 
            Name            =   "Nazanin"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2775
         Left            =   -74640
         RightToLeft     =   -1  'True
         TabIndex        =   25
         Top             =   2760
         Width           =   5775
         Begin VB.TextBox txtLevel1Primary 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   450
            Left            =   720
            RightToLeft     =   -1  'True
            TabIndex        =   34
            Top             =   600
            Width           =   615
         End
         Begin VB.TextBox TxtLevel2 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   450
            Left            =   3480
            RightToLeft     =   -1  'True
            TabIndex        =   29
            Top             =   600
            Width           =   975
         End
         Begin VB.TextBox TxtLevel2New 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   450
            Left            =   1440
            RightToLeft     =   -1  'True
            TabIndex        =   28
            ToolTipText     =   "òœ ê—ÊÂ ›—⁄Ì »«Ìœ »Ì‰ 01 Ê 99 Ê  «»⁄Ì «“ òœ ê—ÊÂ «’·Ì »«‘œ"
            Top             =   600
            Width           =   615
         End
         Begin VB.CommandButton CmdChangeLevel2 
            BackColor       =   &H0000C0C0&
            Caption         =   " €ÌÌ— "
            BeginProperty Font 
               Name            =   "Nazanin"
               Size            =   11.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   240
            RightToLeft     =   -1  'True
            TabIndex        =   27
            Top             =   2040
            Width           =   1215
         End
         Begin VB.ComboBox CmbLevel1New 
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
            Left            =   240
            RightToLeft     =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   26
            Top             =   1200
            Width           =   2000
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            Caption         =   "òœ ê—ÊÂ ›—⁄Ì"
            BeginProperty Font 
               Name            =   "Nazanin"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   4560
            RightToLeft     =   -1  'True
            TabIndex        =   33
            Top             =   600
            Width           =   1095
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            Caption         =   " »œÌ· ‘Êœ »Â"
            BeginProperty Font 
               Name            =   "Nazanin"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000040C0&
            Height          =   390
            Left            =   2160
            RightToLeft     =   -1  'True
            TabIndex        =   32
            Top             =   600
            Width           =   1215
         End
         Begin VB.Label Label8 
            Alignment       =   1  'Right Justify
            Caption         =   " ***òœ ê—ÊÂ ›—⁄Ì œ— œÊ —ﬁ„"
            BeginProperty Font 
               Name            =   "Nazanin"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000040&
            Height          =   495
            Left            =   2280
            RightToLeft     =   -1  'True
            TabIndex        =   31
            Top             =   2040
            Width           =   2895
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            Caption         =   "ê—ÊÂ «’·Ì"
            BeginProperty Font 
               Name            =   "Nazanin"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   2280
            RightToLeft     =   -1  'True
            TabIndex        =   30
            Top             =   1200
            Width           =   1095
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   " €ÌÌ— òœ ¬Ì „ «‰ Œ«»Ì - ê—ÊÂ «’·Ì"
         BeginProperty Font 
            Name            =   "Nazanin"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2175
         Left            =   -74640
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   480
         Width           =   5775
         Begin VB.CommandButton CmdChangeLevel1 
            BackColor       =   &H0000C0C0&
            Caption         =   " €ÌÌ— "
            BeginProperty Font 
               Name            =   "Nazanin"
               Size            =   11.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   240
            RightToLeft     =   -1  'True
            TabIndex        =   21
            Top             =   1320
            Width           =   1215
         End
         Begin VB.TextBox TxtLevel1New 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   450
            Left            =   1200
            RightToLeft     =   -1  'True
            TabIndex        =   20
            ToolTipText     =   "òœ ê—ÊÂ «’·Ì »«Ìœ »Ì‰ 11 Ê 99  »«‘œ"
            Top             =   720
            Width           =   855
         End
         Begin VB.TextBox TxtLevel1 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   450
            Left            =   3600
            RightToLeft     =   -1  'True
            TabIndex        =   19
            Top             =   720
            Width           =   735
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            Caption         =   " *** òœ ê—ÊÂ «’·Ì œ— œÊ —ﬁ„ "
            BeginProperty Font 
               Name            =   "Nazanin"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000040&
            Height          =   495
            Left            =   2520
            RightToLeft     =   -1  'True
            TabIndex        =   24
            Top             =   1560
            Width           =   2895
         End
         Begin VB.Label Label7 
            Alignment       =   1  'Right Justify
            Caption         =   " »œÌ· ‘Êœ »Â"
            BeginProperty Font 
               Name            =   "Nazanin"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000040C0&
            Height          =   390
            Left            =   2160
            RightToLeft     =   -1  'True
            TabIndex        =   23
            Top             =   720
            Width           =   1215
         End
         Begin VB.Label Label6 
            Alignment       =   1  'Right Justify
            Caption         =   "òœ ê—ÊÂ «’·Ì"
            BeginProperty Font 
               Name            =   "Nazanin"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   4560
            RightToLeft     =   -1  'True
            TabIndex        =   22
            Top             =   720
            Width           =   1095
         End
      End
      Begin Total.UcFont UcFont1 
         Height          =   615
         Left            =   360
         TabIndex        =   17
         Top             =   600
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   1085
      End
      Begin VB.Frame Frame1 
         Height          =   8655
         Left            =   -68520
         TabIndex        =   12
         Top             =   480
         Width           =   7515
         Begin VB.CommandButton cmdAddMainGroup 
            BackColor       =   &H00FFC0C0&
            Caption         =   "ê—ÊÂ «’·Ì ÃœÌœ"
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
            Left            =   3960
            Style           =   1  'Graphical
            TabIndex        =   14
            Top             =   8040
            Width           =   2295
         End
         Begin VB.CommandButton cmdAddSubGroup 
            BackColor       =   &H00FFC0C0&
            Caption         =   "“Ì— ê—ÊÂ ÃœÌœ"
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
            Left            =   1110
            Style           =   1  'Graphical
            TabIndex        =   13
            Top             =   8040
            Width           =   2295
         End
         Begin MSComctlLib.TreeView tvLevels 
            Height          =   7215
            Left            =   1080
            TabIndex        =   15
            Top             =   720
            Width           =   5415
            _ExtentX        =   9551
            _ExtentY        =   12726
            _Version        =   393217
            LabelEdit       =   1
            LineStyle       =   1
            Style           =   7
            Appearance      =   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Nazanin"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Label Label3 
            Alignment       =   2  'Center
            Caption         =   "ê—ÊÂÂ«Ì ò«·«"
            BeginProperty Font 
               Name            =   "Titr"
               Size            =   9.75
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   525
            Left            =   4680
            TabIndex        =   16
            Top             =   240
            Width           =   1845
         End
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
         Height          =   2400
         Left            =   11880
         Style           =   1  'Checkbox
         TabIndex        =   10
         Top             =   720
         Width           =   2775
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
         Height          =   2400
         Left            =   9120
         Style           =   1  'Checkbox
         TabIndex        =   9
         Top             =   720
         Width           =   2745
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
         Left            =   5760
         MultiLine       =   -1  'True
         TabIndex        =   2
         Top             =   2370
         Width           =   2385
      End
      Begin FLWCtrls.FWCoolButton fwBtnCustFind 
         Height          =   930
         Left            =   5760
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   840
         Width           =   2370
         _ExtentX        =   4180
         _ExtentY        =   1640
         BackColor       =   16777088
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Nazanin"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "frmGood.frx":A4FA
         PictureAlign    =   4
         Caption         =   " «„Ì‰ ﬂ‰‰œÂ"
         MaskColor       =   -2147483633
      End
      Begin VSFlex7LCtl.VSFlexGrid vsGood 
         Height          =   5940
         Left            =   120
         TabIndex        =   11
         Top             =   3150
         Width           =   14625
         _cx             =   25797
         _cy             =   10477
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
         BackColorFixed  =   -2147483643
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
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   16
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
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
         ExplorerBar     =   1
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   -1  'True
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   5
         Editable        =   2
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
         BackColorFrozen =   -2147483645
         ForeColorFrozen =   0
         WallPaperAlignment=   9
      End
      Begin VB.Image Image1 
         Height          =   1215
         Left            =   600
         Stretch         =   -1  'True
         Top             =   1560
         Width           =   1335
      End
      Begin VB.Label lblGoodLevel1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H000040C0&
         BackStyle       =   0  'Transparent
         Caption         =   "ê—ÊÂ «’·Ì ò«·«Â«"
         BeginProperty Font 
            Name            =   "Titr"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   435
         Left            =   11865
         TabIndex        =   8
         Top             =   345
         Width           =   2655
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
         ForeColor       =   &H8000000D&
         Height          =   435
         Left            =   9000
         TabIndex        =   7
         Top             =   360
         Width           =   2745
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   " «„Ì‰ ﬂ‰‰œÂ"
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
         Left            =   6600
         TabIndex        =   6
         Top             =   480
         Width           =   2175
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "»«—òœ"
         BeginProperty Font 
            Name            =   "Nazanin"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   7920
         TabIndex        =   5
         Top             =   2400
         Width           =   825
      End
      Begin VB.Label LblNumberOfRecords 
         Alignment       =   1  'Right Justify
         Caption         =   " ⁄œ«œ —ﬂÊ—œÂ« :"
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
         Left            =   8220
         TabIndex        =   4
         Top             =   9120
         Width           =   3015
      End
   End
   Begin FLWCtrls.FWLabel3D fwlblMode 
      Height          =   495
      Left            =   13320
      Top             =   0
      Width           =   1545
      _ExtentX        =   2725
      _ExtentY        =   873
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Nazanin"
         Size            =   15.75
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
   Begin FLWCtrls.FWLabel FWLabel1 
      Height          =   495
      Left            =   0
      Top             =   0
      Width           =   14985
      _ExtentX        =   26432
      _ExtentY        =   873
      Enabled         =   -1  'True
      Caption         =   " ⁄—Ì› Ê ê—ÊÂ »‰œÌ ò«·« Â«"
      FirstColor      =   9412754
      SecondColor     =   14215660
      Angle           =   0
      ForeColor       =   7362318
      BackColor       =   12640511
      FontName        =   "Nazanin"
      FontBold        =   -1  'True
      FontSize        =   15.75
      Alignment       =   2
      Picture         =   "frmGood.frx":A814
   End
   Begin RESIZEKITLibCtl.ResizeKit ResizeKit1 
      Height          =   480
      Left            =   0
      OleObjectBlob   =   "frmGood.frx":A830
      TabIndex        =   0
      Top             =   480
      Width           =   480
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmGood"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private MyFormAddEditMode As EnumAddEditMode
Dim tmpTextMatrix As String
Dim Parameter() As Parameter
Dim AsciiNamePrn As String
Dim i As Long
Dim Rst As New ADODB.Recordset
Dim filetemp As New FileSystemObject
Dim strStream
Dim intRepeat As Long

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
    
    mdifrm.Toolbar1.Buttons(20).Enabled = True
    mdifrm.Toolbar1.Buttons(23).Enabled = True
    mdifrm.Toolbar1.Buttons(24).Enabled = True
    mdifrm.Toolbar1.Buttons(25).Enabled = True
    mdifrm.Toolbar1.Buttons(26).Enabled = True
    mdifrm.Toolbar1.Buttons(27).Enabled = True
    
    Select Case MyFormAddEditMode
    
        Case EnumAddEditMode.ViewMode
        
            vsGood.Editable = flexEDNone
            mdifrm.Toolbar1.Buttons(6).Enabled = True  'Add
            mdifrm.Toolbar1.Buttons(7).Enabled = True  'Edit
            mdifrm.Toolbar1.Buttons(8).Enabled = False  'Enter
            mdifrm.Toolbar1.Buttons(9).Enabled = False   'Esc
            mdifrm.Toolbar1.Buttons(10).Enabled = True 'Delete
        
            tvLevels.LabelEdit = tvwManual
            
        Case EnumAddEditMode.AddMode
        
            vsGood.Editable = flexEDKbdMouse
            vsGood.ColHidden(14) = True
            vsGood.ColHidden(15) = True
            mdifrm.Toolbar1.Buttons(6).Enabled = False  'Add
            mdifrm.Toolbar1.Buttons(7).Enabled = False  'Edit
            mdifrm.Toolbar1.Buttons(8).Enabled = True   'Enter
            mdifrm.Toolbar1.Buttons(9).Enabled = True   'Esc
            mdifrm.Toolbar1.Buttons(10).Enabled = False 'Delete
            
            tvLevels.LabelEdit = tvwManual
        Case EnumAddEditMode.EditMode
                    
            vsGood.Editable = flexEDKbdMouse
            vsGood.ColHidden(14) = False
            vsGood.ColHidden(15) = False
    
    '        mdifrm.Toolbar1.Buttons(6).Enabled = True  'Add
            mdifrm.Toolbar1.Buttons(7).Enabled = False  'Edit
            mdifrm.Toolbar1.Buttons(8).Enabled = True   'Enter
            mdifrm.Toolbar1.Buttons(9).Enabled = True   'Esc
            mdifrm.Toolbar1.Buttons(10).Enabled = False 'Delete
            
            tvLevels.LabelEdit = tvwAutomatic
    End Select
    
    HeaderLabel Val(MyFormAddEditMode), fwlblMode
    
End Sub

Private Sub CmbLevel1New_Click()
    If CmbLevel1New.ListIndex = -1 Then Exit Sub
    txtLevel1Primary = CmbLevel1New.ItemData(CmbLevel1New.ListIndex)

End Sub

Private Sub cmdAddMainGroup_Click()

    Add
    mvarInput = ""
    i = 3
'    While (mvarInput = "" And i > 0)
        frmInput.fwlblInput.Caption = " ‰«„ ÃœÌœ »—«Ì ê—ÊÂ «’·Ì ò«·« Ê«—œ ‰„«ÌÌœ "
        frmInput.Picture1.Visible = False
        frmInput.txtInput.Text = ""
        frmInput.MyForm = Me.Name
        frmInput.Show vbModal
'        i = i - 1
        mvarInput = Trim(mvarInput)
'    Wend
    
    If mvarInput = "" Then
        Cancel
        Exit Sub
    End If
    
    On Error GoTo RollBack
    ReDim Parameter(0) As Parameter
    Parameter(0) = GenerateInputParameter("@Description", adVarWChar, 50, mvarInput)
    RunParametricStoredProcedure "InsertGoodLevel1", Parameter
    
    On Error GoTo 0
    
    MyFormAddEditMode = ViewMode
    SetFirstToolBar
    DefaultSetting
    
    Exit Sub
    
RollBack:
    
    frmMsg.fwlblMsg.Caption = "‘„« ‰„Ì  Ê«‰Ìœ «Ì‰ ê—ÊÂ —« «÷«›Â ò‰Ìœ "
    frmMsg.fwBtn(0).ButtonType = flwButtonOk
    frmMsg.fwBtn(0).Caption = "ﬁ»Ê·"
    frmMsg.fwBtn(1).Visible = False
    frmMsg.Show vbModal
    
    MyFormAddEditMode = ViewMode
    SetFirstToolBar
    DefaultSetting

End Sub

Private Sub cmdAddSubGroup_Click()
    Add
    mvarInput = ""
    i = 3
'    While (mvarInput = "" And i > 0)
        frmInput.fwlblInput.Caption = " ‰«„ ÃœÌœ »—«Ì “Ì— ê—ÊÂ ò«·« Ê«—œ ‰„«ÌÌœ "
        frmInput.Picture1.Visible = False
        frmInput.txtInput.Text = ""
        frmInput.MyForm = Me.Name
        frmInput.Show vbModal
'        i = i - 1
        mvarInput = Trim(mvarInput)
'    Wend
    
    If mvarInput = "" Then
        Cancel
        Exit Sub
    End If
    
    On Error GoTo RollBack
    
    ReDim Parameter(1) As Parameter
    Parameter(0) = GenerateInputParameter("@Description", adVarWChar, 50, mvarInput)
    Parameter(1) = GenerateInputParameter("@IntLevel1", adInteger, 4, Left(tvLevels.SelectedItem.Tag, 2))
    RunParametricStoredProcedure "InsertGoodLevel2", Parameter
    
    On Error GoTo 0
    
    MyFormAddEditMode = ViewMode
    SetFirstToolBar
    DefaultSetting
    
    Exit Sub
    
RollBack:
    
    frmMsg.fwlblMsg.Caption = "‘„« ‰„Ì  Ê«‰Ìœ «Ì‰ ê—ÊÂ —« «÷«›Â ò‰Ìœ "
    frmMsg.fwBtn(0).ButtonType = flwButtonOk
    frmMsg.fwBtn(0).Caption = "ﬁ»Ê·"
    frmMsg.fwBtn(1).Visible = False
    frmMsg.Show vbModal
    
    MyFormAddEditMode = ViewMode
    SetFirstToolBar
    DefaultSetting

End Sub

Public Sub DefaultSetting()

'    If SSTab1.Tab = 0 Then
        
        Dim Rst As New ADODB.Recordset
        Dim varNode As node
        ReDim Parameter(0) As Parameter
        
        Parameter(0) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
        
        tvLevels.Nodes.Clear
        If Rst.State <> 0 Then Rst.Close
        
       Set Rst = RunParametricStoredProcedure2Rec("GetGoodLevel1_Description", Parameter)
    
        If Not (Rst.EOF = True And Rst.BOF = True) Then
            i = 1
            While Rst.EOF <> True
                Set varNode = tvLevels.Nodes.Add
                varNode.Text = Rst.Fields("Description").Value
                varNode.Tag = Rst.Fields("code").Value
                Rst.MoveNext
            Wend
        End If
        
        i = tvLevels.Nodes.Count
        
        While i >= 1
        
            ReDim Parameter(1) As Parameter
            Parameter(0) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
            Parameter(1) = GenerateInputParameter("@CurrentItem", adInteger, 4, CInt(tvLevels.Nodes.Item(i).Tag))
            If Rst.State <> 0 Then Rst.Close
            Set Rst = RunParametricStoredProcedure2Rec("GetGoodLevel2_Description", Parameter)
            
            If Not (Rst.EOF = True And Rst.BOF = True) Then
                While Rst.EOF <> True
                    Set varNode = tvLevels.Nodes.Add(i, tvwChild)
                    varNode.Text = Rst.Fields("Description").Value
                    varNode.Tag = Rst.Fields("code").Value
                    Rst.MoveNext
                Wend
                
            End If
            tvLevels.Nodes.Item(i).Expanded = True
            i = i - 1
        Wend
        If tvLevels.Nodes.Count > 0 Then
            tvLevels.Nodes.Item(1).Selected = True
        End If
        CmbLevel1New.Clear
        ReDim Parameter(0) As Parameter
        Parameter(0) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
        Set Rst = RunParametricStoredProcedure2Rec("Get_tGoodLevel1", Parameter)
            
        If Not (Rst.EOF = True And Rst.BOF = True) Then
            While Rst.EOF = False
                CmbLevel1New.AddItem Rst.Fields("Description")
                CmbLevel1New.ItemData(CmbLevel1New.ListCount - 1) = Rst.Fields("Code")
                Rst.MoveNext
            Wend
        End If
    
'    ElseIf SSTab1.Tab = 1 Then
    
        lstGoodLevel1.Clear
        lstGoodLevel2.Clear
        vsGood.Rows = 1
        
        FillLstGoodLevel1
'    End If
End Sub

Public Sub FillLstGoodLevel1() ' it fills the lstGoodLevel1 using table tgoodlevel1
    Dim Rst As New ADODB.Recordset
    
    lstGoodLevel1.Clear
    lstGoodLevel2.Clear
    vsGood.Rows = 1
    
    MyFormAddEditMode = EnumAddEditMode.ViewMode
    SetFirstToolBar
    
    ReDim Parameter(0) As Parameter
    Parameter(0) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
    Set Rst = RunParametricStoredProcedure2Rec("Get_tGoodLevel1", Parameter)
        
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
        ReDim Parameter(1) As Parameter
        Parameter(0) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
        Parameter(1) = GenerateInputParameter("@CurrentItem", adInteger, 4, lstGoodLevel1.ItemData(lstGoodLevel1.ListIndex))
        Set rctmp = RunParametricStoredProcedure2Rec("GetGoodLevel2_Description", Parameter)
        
        vsGood.ColComboList(15) = vsGood.BuildComboList(rctmp, "Description", "Code")
        
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
    
    Dim i As Integer
    Dim j As Integer
    Dim intSelectedLevel1 As Integer
    Dim intSelectedLevel2 As Integer
    Dim strSelectedLevels As String

    
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
    
    If intSelectedLevel1 <> -1 And intSelectedLevel2 <> -1 Then
        
        strSelectedLevels = Right(strSelectedLevels, Len(strSelectedLevels) - 1)
        ReDim Parameter(3) As Parameter
        Parameter(0) = GenerateInputParameter("@Level1", adInteger, 4, lstGoodLevel1.ItemData(intSelectedLevel1))
        Parameter(1) = GenerateInputParameter("@strSelectedLevels", adVarWChar, 4000, strSelectedLevels)
        Parameter(2) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
        Parameter(3) = GenerateInputParameter("@ProductCompany", adInteger, 4, Val(fwBtnCustFind.Tag))
        Set Rst = RunParametricStoredProcedure2Rec("Get_Good_In_Levels", Parameter)
    
    ElseIf intSelectedLevel1 <> -1 And intSelectedLevel2 = -1 Then  'Or intSelectedLevel2 = -1
        ReDim Parameter(2) As Parameter
        Parameter(0) = GenerateInputParameter("@GoodLevel1Code", adInteger, 4, lstGoodLevel1.ItemData(intSelectedLevel1))
        Parameter(1) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
        Parameter(2) = GenerateInputParameter("@ProductCompany", adInteger, 4, Val(fwBtnCustFind.Tag))
        Set Rst = RunParametricStoredProcedure2Rec("GetVw_GoodInfo", Parameter)
    
    Else
        ReDim Parameter(1) As Parameter
        Parameter(0) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
        Parameter(1) = GenerateInputParameter("@ProductCompany", adInteger, 4, Val(fwBtnCustFind.Tag))
        Set Rst = RunParametricStoredProcedure2Rec("GetVwGoodInfo", Parameter)
    
    End If
    
    If (Rst.EOF = True And Rst.BOF = True) Then Exit Sub
    
    With vsGood
        
        
        i = 1
        
        MousePointer = 11
        While Rst.EOF = False
            .Rows = .Rows + 1
            .TextMatrix(i, 0) = i
            .TextMatrix(i, 1) = Rst.Fields("Code").Value
            .TextMatrix(i, 2) = Rst.Fields("GoodName").Value
            .TextMatrix(i, 3) = Rst.Fields("NamePrn").Value
            .TextMatrix(i, 4) = Rst.Fields("BarCode").Value
            .TextMatrix(i, 5) = Rst.Fields("SellPrice").Value
             .TextMatrix(i, 6) = Rst.Fields("SellPrice2").Value
             .TextMatrix(i, 7) = Rst.Fields("SellPrice3").Value
             .TextMatrix(i, 8) = Rst.Fields("SellPrice4").Value
             .TextMatrix(i, 9) = Rst.Fields("SellPrice5").Value
             .TextMatrix(i, 10) = Rst.Fields("SellPrice6").Value
            .TextMatrix(i, 11) = Rst.Fields("BuyPrice").Value
          '  .TextMatrix(i, 12) = Rst.Fields("unitdescription").Value
            .TextMatrix(i, 13) = Rst.Fields("GoodType").Value

            .Cell(flexcpText, i, 12) = CStr(Rst.Fields("unit").Value)
            .Cell(flexcpText, i, 14) = CStr(Rst.Fields("Level1").Value)
            .Cell(flexcpText, i, 15) = CStr(Rst.Fields("Level2").Value)
            .Cell(flexcpText, i, 16) = CStr(Rst.Fields("ProductCompany").Value)  'Rst.Fields("CompDes").Value
'             If Not (Val(strCategory) >= 0 And Val(strCategory) <= 7) Then
                .TextMatrix(i, 17) = Rst.Fields("Weight").Value
                .TextMatrix(i, 18) = Rst.Fields("NumberOfUnit").Value
'             End If
            .TextMatrix(i, 19) = IIf(Rst.Fields("MainType").Value = True, -1, 0)
            .TextMatrix(i, 20) = IIf(IsNull(Rst.Fields("CategoryShow").Value), "", Rst.Fields("CategoryShow").Value)
           .TextMatrix(i, 21) = IIf(IsNull(Rst.Fields("PicturePath").Value), "", Rst.Fields("PicturePath").Value)
           .TextMatrix(i, 22) = IIf(IsNull(Rst.Fields("nvcDescription").Value), "", Rst.Fields("nvcDescription").Value)
           .TextMatrix(i, 23) = IIf(IsNull(Rst.Fields("GoodNamePrn2").Value), "", Rst.Fields("GoodNamePrn2").Value)
           .TextMatrix(i, 24) = IIf(IsNull(Rst.Fields("GoodNamePrn3").Value), "", Rst.Fields("GoodNamePrn3").Value)
            .TextMatrix(i, 25) = Rst.Fields("Code").Value
            i = i + 1
            Rst.MoveNext
            
        Wend
        MousePointer = 0

        Set Rst = Nothing
        LblNumberOfRecords.Caption = " ⁄œ«œ —ﬂÊ—œÂ« :  " & i - 1
'        .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter
'        .Cell(flexcpAlignment, 1, 4, .Rows - 1, .Cols - 1) = flexAlignCenterCenter
'        .AutoSizeMode = flexAutoSizeColWidth
'        .AutoSize 1, 4
'        .AutoSize 13, 13
'        If .ColWidth(2) > 3500 Then .ColWidth(2) = 3500
'        If .ColWidth(3) > 3000 Then .ColWidth(3) = 3000
    End With
        
End Sub

Public Sub BeforeUpdate()

End Sub

Public Sub Edit()
    MyFormAddEditMode = EnumAddEditMode.EditMode
    SetFirstToolBar
    With vsGood
        
        .Editable = flexEDKbdMouse
'        .EditCell

    End With
End Sub

Public Sub Update()
    intRepeat = 0
    On Error GoTo ErrHandler
    Dim i As Integer
    Dim j As Integer
    Dim LongTemp As Integer
    Dim lngSelectedSubGroup  As Long
    
    Dim Rst As New ADODB.Recordset
    Dim NewCode As Double

If SSTab1.Tab = 0 Then
    Select Case MyFormAddEditMode
        Case AddMode 'add
        
        Case EditMode 'edit
        
            tvLevels.SetFocus
            ReDim Parameter(2) As Parameter
            For i = 1 To tvLevels.Nodes.Count
                Parameter(0) = GenerateInputParameter("@code", adInteger, 4, CInt(tvLevels.Nodes.Item(i).Tag))
                Parameter(1) = GenerateInputParameter("@Description", adVarWChar, 50, tvLevels.Nodes.Item(i).Text)
                Parameter(2) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
                
                If InStr(1, tvLevels.Nodes.Item(i).FullPath, "\") = 0 Then 'root
                    RunParametricStoredProcedure "UpdateGoodLevel1", Parameter
                Else
                    RunParametricStoredProcedure "UpdateGoodLevel2", Parameter
                End If
            Next i
            
            frmMsg.fwlblMsg.Caption = "À»   €ÌÌ—«  »« „Ê›ﬁÌ  »Â Å«Ì«‰ —”Ìœ"
            frmMsg.fwBtn(0).Visible = False
            frmMsg.fwBtn(1).ButtonType = flwButtonOk
            frmMsg.fwBtn(1).Caption = "ﬁ»Ê·"
            frmMsg.Show vbModal

    End Select


    DefaultSetting
    MyFormAddEditMode = ViewMode
    SetFirstToolBar
    

ElseIf SSTab1.Tab = 1 Then
    
    Dim Result As Integer
    
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
           ' .Row = i
            If InStr(.TextMatrix(i, 0), "*") > 0 Then 'new or edited records
            
                If Not ((Trim(.TextMatrix(i, 2)) = "" And Trim(.TextMatrix(i, 3)) = "" And Trim(.TextMatrix(i, 5)) = "")) Then  'And .Cell(flexcpText, i, 12) = "" And .Cell(flexcpText, i, 13) = ""
                    If ((Trim(.TextMatrix(i, 2)) = "" And Trim(.TextMatrix(i, 3)) = "") Or Trim(.TextMatrix(i, 5)) = "") Or .Cell(flexcpText, i, 12) = "" Or .Cell(flexcpText, i, 13) = "" Then
                        
                        Select Case clsStation.Language
                        
                            Case 0
                            
                                frmMsg.fwlblMsg.Caption = "‘„« „Ì »«Ì”  «ÿ·«⁄«  —« »ÿÊ— ò«„· Ê«—œ ‰„«ÌÌœ"
                                frmMsg.fwBtn(0).Caption = "ﬁ»Ê·"
                            Case 1
                            
                                frmMsg.fwlblMsg.Caption = "You Have to complete the information"
                                frmMsg.fwBtn(0).Caption = "Ok"
                                frmMsg.fwlblMsg.Alignment = vbLeftJustify
                        
                        End Select
                        
                        frmMsg.fwBtn(0).ButtonType = flwButtonOk
                        frmMsg.fwBtn(1).Visible = False
                        frmMsg.Show vbModal
                        
                        Exit Sub
                        
                    End If
                End If
                
                If (Val(.TextMatrix(i, 13)) = 2 Or Val(.TextMatrix(i, 13)) = 3) And Val(.TextMatrix(i, 5)) < 0 Then    '  By NHashemi For Save Service With Zero Fee
                        Select Case clsStation.Language
                        
                            Case 0
                            
                                frmMsg.fwlblMsg.Caption = "‘„« „Ì »«Ì”  ﬁÌ„  ›—Ê‘ —« Ê«—œ ‰„«ÌÌœ"
                                frmMsg.fwBtn(0).Caption = "ﬁ»Ê·"
                            Case 1
                            
                                frmMsg.fwlblMsg.Caption = "You Have to complete the information"
                                frmMsg.fwBtn(0).Caption = "Ok"
                                frmMsg.fwlblMsg.Alignment = vbLeftJustify
                        
                        End Select
                        
                        frmMsg.fwBtn(0).ButtonType = flwButtonOk
                        frmMsg.fwBtn(1).Visible = False
                        frmMsg.Show vbModal
                        .Row = i
                        .Col = 5
                        .Select .Row, .Col
                        .EditCell
                        
                        Exit Sub

                End If
                
''''                If (Val(.TextMatrix(i, 10)) = 1 Or Val(.TextMatrix(i, 10)) = 3) And Val(.TextMatrix(i, 8)) = 0 Then
''''                        Select Case clsStation.Language
''''
''''                            Case 0
''''
''''                                frmMsg.fwlblMsg.Caption = "‘„« „Ì »«Ì”  ﬁÌ„  Œ—Ìœ Ê«—œ ‰„«ÌÌœ"
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
''''                End If
                If Len(.TextMatrix(i, 4)) > 0 Then
                    If InStr(1, .TextMatrix(i, 4), "/", 1) Then
                        LongTemp = InStr(2, .TextMatrix(i, 4), "/", 1)
                        If LongTemp > 2 Then
                           .TextMatrix(i, 4) = Mid(.TextMatrix(i, 4), 2, LongTemp - 2)
                        End If
                    End If
                    If GetGoodBarcode(.TextMatrix(i, 4), Val(Trim(.TextMatrix(i, 1)))) = True Then
                        frmMsg.fwlblMsg.Caption = " . «Ì‰ »«—ﬂœ ﬁ»·« œ— ”Ì” „  ⁄—Ì› ‘œÂ «”  "
                        frmMsg.fwBtn(0).Caption = "ﬁ»Ê·"
                        frmMsg.fwBtn(1).Visible = False
                        frmMsg.fwBtn(0).ButtonType = flwButtonOk
                        frmMsg.Show vbModal
                    '    .TextMatrix(.Row, 4) = ""
                        .Row = i
                        .Col = 4
                        .Select .Row, .Col
                        .EditCell
                        Exit Sub
                    End If
            
                End If
                If Len(.TextMatrix(i, 1)) <> 8 And clsArya.HardLockSerialNo <> "93070903507" Then
                    ShowMessage "òœ ò«·« »«Ìœ 8 —ﬁ„ »«‘œ", True, False, "ﬁ»Ê·", ""
                    .Row = i
                    .Col = 1
                    .Select .Row, .Col
                    .EditCell
                    Exit Sub
                End If
            End If
        Next i
        
        For j = 0 To lstGoodLevel2.ListCount - 1
            If lstGoodLevel2.Selected(j) = True Then
                lngSelectedSubGroup = j
                Exit For
            End If
        Next j

        Select Case MyFormAddEditMode
        
            Case EnumAddEditMode.AddMode
            
'                If clsArya.LimitedVersion = True Then
'                    Dim strtemporary As String
'                    Dim cnn As New ADODB.Connection
'                    Dim rctmp As New Recordset
'                    Dim CountRecord As Long
'                    cnn.Open strConnectionString
'                    strtemporary = "Select Count(*) as CountRecord from tGood"
'                    rctmp.Open strtemporary, cnn, adOpenDynamic, adLockOptimistic, adCmdText
'                    If Not (rctmp.EOF = True And rctmp.BOF = True) Then
'                       CountRecord = Val(rctmp!CountRecord)
'                       If CountRecord >= GoodsCountingRecord Then
'                          MsgBox " ‰”ŒÂ ¬“„«Ì‘Ì - ‘„« »Ì‘ «“ «Ì‰ „Ã«“ »Â «” ›«œÂ «“ «Ì‰ ”Ì” „ ‰Ì” Ìœ -2 " & vbCrLf & " »« ‘—ﬂ  «› ÃÌ ¬—Ì«  „«” »êÌ—Ìœ" & vbLf & " ·›‰  „«” : 88554488-88554477-88554466-88554455"
'                          Exit Sub
'                       End If
'                    End If
'                    rctmp.Close
'                    cnn.Close
'                    Set cnn = Nothing
'                End If
                
                For i = 1 To .Rows - 1
                    
                    If .TextMatrix(i, 0) = "*" Then 'new records
                        
                        If Not (Trim(.TextMatrix(i, 2)) = "" And Trim(.TextMatrix(i, 3)) = "" And Trim(.TextMatrix(i, 5)) = "") Then  'And Trim(.TextMatrix(i, 11)) = ""
                            
                            If Rst.State <> 0 Then Rst.Close
                                
                            ReDim Parameter(0) As Parameter

                            Parameter(0) = GenerateInputParameter("@Level2Code", adInteger, 4, lstGoodLevel2.ItemData(lngSelectedSubGroup))
                            Set Rst = RunParametricStoredProcedure2Rec("ChecktGoodData", Parameter)
        
                            If Not (Rst.EOF And Rst.BOF) Then
                                If InStr(1, .TextMatrix(i, 4), "/", 1) Then
                                   LongTemp = InStr(2, .TextMatrix(i, 4), "/", 1)
                                    If LongTemp > 2 Then
                                       .TextMatrix(i, 4) = Mid(.TextMatrix(i, 4), 2, LongTemp - 2)
''''                                       If Left(.TextMatrix(i, 4), 2) <> "62" Then
''''                                          .TextMatrix(i, 4) = "62" & .TextMatrix(i, 4)
''''                                       End If
                                    End If
                                Else
''''                                    If Len(Trim(.TextMatrix(i, 4))) > 0 Then
''''                                        frmDisMsg.lblMessage.Caption = " ›Ê—„  »«—ﬂœ  Ê«—œ ‘œÂ ’ÕÌÕ ‰Ì”  "
''''                                        frmDisMsg.Timer1.Interval = 1000
''''                                        frmDisMsg.Timer1.Enabled = True
''''                                        frmDisMsg.Show vbModal
''''                                    End If
                                End If
                                Set strStream = New ADODB.Stream
                                strStream.Type = adTypeBinary
                                strStream.Open
                                If Len(.TextMatrix(i, 21)) = 0 Then

                                Else
                                    If filetemp.FileExists(App.Path & .TextMatrix(i, 21)) = False Then
                                        ShowDisMessage "›«Ì· „Õ ÊÌ ¬ÌﬂÊ‰ œ— „”Ì— „‘Œ’ ‘œÂ ÊÃÊœ ‰œ«—œ", 1000
                                    Else
                                        strStream.LoadFromFile App.Path & .TextMatrix(i, 21)
                                    End If
                                End If
                                NewCode = Val(Trim(.TextMatrix(i, 1)))
                                ReDim Parameter(27) As Parameter
                                
                                Parameter(0) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
                                Parameter(1) = GenerateInputParameter("@Code", adInteger, 4, NewCode)
                                Parameter(2) = GenerateInputParameter("@GoodName", adVarWChar, 50, .TextMatrix(i, 2))
                                Parameter(3) = GenerateInputParameter("@GoodNamePrn", adVarWChar, 50, .TextMatrix(i, 3))
                                Parameter(4) = GenerateInputParameter("@SellPrice", adDouble, 8, Val(.TextMatrix(i, 5)))
                                Parameter(5) = GenerateInputParameter("@BuyPrice", adDouble, 8, Val(.TextMatrix(i, 11)))
                                Parameter(6) = GenerateInputParameter("@Barcode", adVarWChar, 50, .TextMatrix(i, 4))
                                Parameter(7) = GenerateInputParameter("@Level1", adInteger, 4, Val(Left(Format(CStr(lstGoodLevel2.ItemData(lngSelectedSubGroup)), "0000"), 2)))
                                Parameter(8) = GenerateInputParameter("@Level2", adInteger, 4, Val(lstGoodLevel2.ItemData(lngSelectedSubGroup)))
                                Parameter(9) = GenerateInputParameter("@Model", adInteger, 4, 1)
                                Parameter(10) = GenerateInputParameter("@Supplier", adInteger, 4, Val(fwBtnCustFind.Tag))
                                Parameter(11) = GenerateInputParameter("@Unit", adInteger, 4, Val(.Cell(flexcpText, i, 12)))
                                Parameter(12) = GenerateInputParameter("@GoodType", adInteger, 4, Val(.Cell(flexcpText, i, 13)))
                                Parameter(13) = GenerateInputParameter("@Weight", adDouble, 8, IIf(.TextMatrix(i, 17) = "", 1, .TextMatrix(i, 17)))
                                Parameter(14) = GenerateInputParameter("@NumberOfUnit", adInteger, 4, IIf(.TextMatrix(i, 18) = "", 1, .TextMatrix(i, 18)))
                                Parameter(15) = GenerateInputParameter("@SellPrice2", adDouble, 8, Val(.TextMatrix(i, 6)))
                                Parameter(16) = GenerateInputParameter("@SellPrice3", adDouble, 8, Val(.TextMatrix(i, 7)))
                                Parameter(17) = GenerateInputParameter("@MainType", adBoolean, 1, IIf(Val(.TextMatrix(i, 19)) = -1, 1, 0))
                                Parameter(18) = GenerateInputParameter("@SellPrice4", adDouble, 8, Val(.TextMatrix(i, 8)))
                                Parameter(19) = GenerateInputParameter("@SellPrice5", adDouble, 8, Val(.TextMatrix(i, 9)))
                                Parameter(20) = GenerateInputParameter("@SellPrice6", adDouble, 8, Val(.TextMatrix(i, 10)))
                                Parameter(21) = GenerateInputParameter("@CategoryShow", adInteger, 4, IIf(Val(.Cell(flexcpText, i, 20)) = 0, Null, Val(.Cell(flexcpText, i, 20))))
                                Parameter(22) = GenerateInputParameter("@PicturePath", adVarWChar, 100, .TextMatrix(i, 21))
                                Parameter(23) = GenerateInputParameter("@nvcDescription", adVarWChar, 100, .TextMatrix(i, 22))
                                Parameter(24) = GenerateInputParameter("@Picture", adLongVarBinary, strStream.Size + 1, strStream.Read)
                                Parameter(25) = GenerateInputParameter("@GoodNamePrn2", adVarWChar, 100, .TextMatrix(i, 23))
                                Parameter(26) = GenerateInputParameter("@GoodNamePrn3", adVarWChar, 100, .TextMatrix(i, 24))
                                Parameter(27) = GenerateOutputParameter("@Result", adInteger, 4)
                                                             
                                Result = RunParametricStoredProcedure("InserttGood", Parameter)
                                
                                If Result > 0 Then
                                    ShowDisMessage "À»   - " & .TextMatrix(i, 2) & "  -«‰Ã«„ ‘œ", 1000
                                Else
                                    ShowDisMessage "œ— À»   - " & .TextMatrix(i, 2) & "-  À»  «‰Ã«„ ‰‘œ - „‘ò· ÊÃÊœ œ«—œ", 1000
                                End If
                            
                            
                            End If
                            
                        End If
                        
                    End If
                
                Next i
                
                
            Case EnumAddEditMode.EditMode
                For i = 1 To .Rows - 1
                    
                    If InStr(.TextMatrix(i, 0), "*") > 0 Then 'Edited records
                            
                            If InStr(1, .TextMatrix(i, 4), "/", 1) Then
                               LongTemp = InStr(2, .TextMatrix(i, 4), "/", 1)
                               If LongTemp > 2 Then
                                  .TextMatrix(i, 4) = Mid(.TextMatrix(i, 4), 2, LongTemp - 2)
''''                                  If Left(.TextMatrix(i, 4), 2) <> "62" Then
''''                                     .TextMatrix(i, 4) = "62" & .TextMatrix(i, 11)
''''                                  End If
                               End If
                            End If
                            Set strStream = New ADODB.Stream
                            strStream.Type = adTypeBinary
                            strStream.Open
                            If Len(.TextMatrix(i, 21)) = 0 Then

                            Else
                                If filetemp.FileExists(App.Path & .TextMatrix(i, 21)) = False Then
                                    ShowDisMessage "›«Ì· „Õ ÊÌ ¬ÌﬂÊ‰ œ— „”Ì— „‘Œ’ ‘œÂ ÊÃÊœ ‰œ«—œ", 1000
                                Else
                                    strStream.LoadFromFile App.Path & .TextMatrix(i, 21)
                                End If
                            End If
                            If Val(Trim(.TextMatrix(i, 1))) = Val(Trim(.TextMatrix(i, 25))) Then
                                NewCode = 0
                            Else
                                NewCode = Val(Trim(.TextMatrix(i, 1)))
                            End If
                            
                            ReDim Parameter(27) As Parameter
                            
                            Parameter(0) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
                            Parameter(1) = GenerateInputParameter("@Goodname", adVarWChar, 50, .TextMatrix(i, 2))
                            Parameter(2) = GenerateInputParameter("@GoodNamePrn", adVarWChar, 50, .TextMatrix(i, 3))
                            Parameter(3) = GenerateInputParameter("@SellPrice", adDouble, 8, Val(.TextMatrix(i, 5)))
                            Parameter(4) = GenerateInputParameter("@BuyPrice", adDouble, 8, Val(.TextMatrix(i, 11)))
                            Parameter(5) = GenerateInputParameter("@Unit", adInteger, 4, Val(.Cell(flexcpText, i, 12)))
                            Parameter(6) = GenerateInputParameter("@GoodType", adInteger, 4, Val(.Cell(flexcpText, i, 13)))
                            Parameter(7) = GenerateInputParameter("@Barcode", adVarWChar, 50, .TextMatrix(i, 4))
                            Parameter(8) = GenerateInputParameter("@Code", adInteger, 4, Val(Trim(.TextMatrix(i, 25))))
                            Parameter(9) = GenerateInputParameter("@Weight", adDouble, 8, IIf(.TextMatrix(i, 17) = "", 1, .TextMatrix(i, 17)))
                            Parameter(10) = GenerateInputParameter("@NumberOfUnit", adInteger, 4, IIf(.TextMatrix(i, 18) = "", 1, .TextMatrix(i, 18)))
                            Parameter(11) = GenerateInputParameter("@SellPrice2", adDouble, 8, Val(.TextMatrix(i, 6)))
                            Parameter(12) = GenerateInputParameter("@SellPrice3", adDouble, 8, Val(.TextMatrix(i, 7)))
                            Parameter(13) = GenerateInputParameter("@MainType", adBoolean, 1, IIf(Val(.TextMatrix(i, 19)) = -1, 1, 0))
                            Parameter(14) = GenerateInputParameter("@Supplier", adInteger, 4, IIf(Val(.Cell(flexcpText, i, 16)) = 0, -1, Val(.Cell(flexcpText, i, 16))))
                            Parameter(15) = GenerateInputParameter("@Level1", adInteger, 4, Val(.Cell(flexcpText, i, 14)))
                            Parameter(16) = GenerateInputParameter("@Level2", adInteger, 4, Val(.Cell(flexcpText, i, 15)))
                            Parameter(17) = GenerateInputParameter("@SellPrice4", adDouble, 8, Val(.TextMatrix(i, 8)))
                            Parameter(18) = GenerateInputParameter("@SellPrice5", adDouble, 8, Val(.TextMatrix(i, 9)))
                            Parameter(19) = GenerateInputParameter("@SellPrice6", adDouble, 8, Val(.TextMatrix(i, 10)))
                            Parameter(20) = GenerateInputParameter("@CategoryShow", adInteger, 4, IIf(Val(.Cell(flexcpText, i, 20)) = 0, Null, Val(.Cell(flexcpText, i, 20))))
                            Parameter(21) = GenerateInputParameter("@PicturePath", adVarWChar, 100, .TextMatrix(i, 21))
                            Parameter(22) = GenerateInputParameter("@nvcDescription", adVarWChar, 100, .TextMatrix(i, 22))
                            Parameter(23) = GenerateInputParameter("@Picture", adLongVarBinary, strStream.Size + 1, strStream.Read)
                            Parameter(24) = GenerateInputParameter("@GoodNamePrn2", adVarWChar, 100, .TextMatrix(i, 23))
                            Parameter(25) = GenerateInputParameter("@GoodNamePrn3", adVarWChar, 100, .TextMatrix(i, 24))
                            Parameter(26) = GenerateInputParameter("@RealNewCode", adInteger, 4, NewCode)
                            Parameter(27) = GenerateOutputParameter("@Result", adInteger, 4)
                            Result = RunParametricStoredProcedure("UpdatetGood", Parameter)
                            
                            If Result > 0 Then
                                ShowDisMessage " €ÌÌ—«   - " & .TextMatrix(i, 2) & "  -«‰Ã«„ ‘œ", 1000
                            Else
                                ShowDisMessage "  €ÌÌ—«   - " & .TextMatrix(i, 2) & "-  À»  «‰Ã«„ ‰‘œ - „‘ò· ÊÃÊœ œ«—œ", 2000
                            End If
                            
                    End If
                                        
                Next i
            
            End Select
            
        FillvsGood
        
    End With
    
    Set Rst = Nothing
End If

    If clsArya.LimitedVersion = True And HardLockFlagTrial = False And (RemaindateFlag = True Or maxRecordCountFlag = True) Then
        TrialCountFlag = TrialCountFlag + 1
        If TrialCountFlag Mod 2 = 0 Then
            ShowMessage " ‘„« œ— Õ«· «” ›«œÂ «“ ‰”ŒÂ ¬“„«Ì‘Ì ”Ì” „ Â«Ì ›—Ê‘ê«ÂÌ «› ÃÌ ¬—Ì« „Ì »«‘Ìœ" & " »—«Ì  ÂÌÂ ‰”ŒÂ «’·Ì ‰—„ «›“«—»« ‘—ﬂ  «› ÃÌ ¬—Ì« Ì« ‰„«Ì‰œê«‰ ›—Ê‘  „«” »êÌ—Ìœ ", True, False, " «∆Ìœ", ""
            Sleep 1000 * TrialCountFlag / 2
        End If
    End If
    Set strStream = Nothing

Exit Sub

ErrHandler:
    
    MsgBox err.Description
'    frmMsg.fwlblMsg.Caption = " €ÌÌ—«  „Ê—œ ‰Ÿ— «⁄„«· ‰‘œ" + vbCrLf + "·ÿ›« «ÿ·«⁄«  ò«„· Ê œ—”  Ê«—œ ‰„«ÌÌœ"
'    frmMsg.fwBtn(0).Visible = False
'    frmMsg.fwBtn(1).ButtonType = flwButtonOk
'    frmMsg.fwBtn(1).Caption = "ﬁ»Ê·"
'    frmMsg.Show vbModal
End Sub

Public Sub BeforeAdd()

Dim lngSelectedSubGroup  As Long
lngSelectedSubGroup = -1
If SSTab1.Tab = 1 Then
    Dim i As Integer
    Dim j As Integer
    
    i = 0
    For j = 0 To lstGoodLevel1.ListCount - 1
        If lstGoodLevel1.Selected(j) = True Then
            i = i + 1
        End If
    Next j
    
    If i <> 1 Then
        Select Case clsStation.Language
            Case 0
                frmMsg.fwlblMsg.Caption = "‘„« „Ì »«Ì”  1 Ê ›ﬁÿ 1 ê—ÊÂ «’·Ì —« «‰ Œ«» ‰„«ÌÌœ"
                frmMsg.fwBtn(0).Caption = "ﬁ»Ê·"
            Case 1
                frmMsg.fwlblMsg.Caption = "You Have to choose 1 and only 1 main Group"
                frmMsg.fwBtn(0).Caption = "Ok"
                frmMsg.fwlblMsg.Alignment = vbLeftJustify
        
        End Select
        frmMsg.fwBtn(0).ButtonType = flwButtonOk
        frmMsg.fwBtn(1).Visible = False
        frmMsg.Show vbModal
        Exit Sub
        
    End If
    
    i = 0
    For j = 0 To lstGoodLevel2.ListCount - 1
        If lstGoodLevel2.Selected(j) = True Then
            lngSelectedSubGroup = j
            i = i + 1
        End If
    Next j
    
    If i <> 1 Then
    
        Select Case clsStation.Language
            Case 0
                frmMsg.fwlblMsg.Caption = "‘„« „Ì »«Ì”  1 Ê ›ﬁÿ 1 ê—ÊÂ ›—⁄Ì —« «‰ Œ«» ‰„«ÌÌœ"
                frmMsg.fwBtn(0).Caption = "ﬁ»Ê·"
            Case 1
                frmMsg.fwlblMsg.Caption = "You Have to choose 1 and only 1 SubGroup"
                frmMsg.fwBtn(0).Caption = "Ok"
                frmMsg.fwlblMsg.Alignment = vbLeftJustify
        
        End Select
        frmMsg.fwBtn(0).ButtonType = flwButtonOk
        frmMsg.fwBtn(1).Visible = False
        frmMsg.Show vbModal
        Exit Sub
        
    End If
    Dim NewCode As Double
    ReDim Parameter(0) As Parameter

    Parameter(0) = GenerateInputParameter("@Level2Code", adInteger, 4, lstGoodLevel2.ItemData(lngSelectedSubGroup))
    Set Rst = RunParametricStoredProcedure2Rec("ChecktGoodData", Parameter)

    If Not (Rst.EOF And Rst.BOF) Then
        NewCode = Val(lstGoodLevel2.ItemData(lngSelectedSubGroup) & Format(Rst.Fields("code").Value, "0000"))
        NewCode = NewCode + intRepeat
    End If
    With vsGood
        .Rows = .Rows + 1
        .Row = .Rows - 1
        .TextMatrix(.Row, 0) = "*"
        .TextMatrix(.Row, 1) = NewCode
'        .TextMatrix(.Row, 6) = "⁄œœ"
'        .Cell(flexcpText, .Row, 6) = "0"
       ' .Cell(flexcpText, .Row, 5) = "0"
        .TextMatrix(.Row, 12) = "0"
'        .TextMatrix(.Row, 13) = "2"
        .TextMatrix(.Row, 16) = "-1"
        .ShowCell .Row, 1
        .Select .Row, 1
    End With
    intRepeat = intRepeat + 1
    MyFormAddEditMode = EnumAddEditMode.AddMode
    SetFirstToolBar
End If

    
End Sub

Public Sub Add()
    
    MyFormAddEditMode = AddMode
    SetFirstToolBar
End Sub


Public Sub Cancel()
    intRepeat = 0
    MyFormAddEditMode = EnumAddEditMode.ViewMode
    SetFirstToolBar
    If SSTab1.Tab = 1 Then
        FillvsGood
    End If
End Sub

Public Sub Delete()
    
If SSTab1.Tab = 0 Then
    ReDim Parameter(0) As Parameter
    frmMsg.fwlblMsg.Caption = "‘„« œ— Õ«· Õ–› ê—ÊÂ " & tvLevels.SelectedItem.Text & " Â” Ìœ " & vbCrLf & "¬Ì« „ÿ„∆‰Ìœø"
    frmMsg.fwBtn(0).ButtonType = flwButtonOk
    frmMsg.fwBtn(0).Caption = "»·Ì"
    frmMsg.fwBtn(1).ButtonType = flwButtonNo
    frmMsg.fwBtn(1).Caption = "ŒÌ—"
    frmMsg.Show vbModal
    varAnswer = modgl.mvarMsgIdx
    
    On Error GoTo DBErrHandler
    Select Case varAnswer
        Case 1 'yes
            If InStr(1, tvLevels.SelectedItem.FullPath, "\") = 0 Then 'root
                Parameter(0) = GenerateInputParameter("@CurrentItem", adInteger, 4, CInt(tvLevels.SelectedItem.Tag))
                RunParametricStoredProcedure2Rec "RemoveGoodLevel1", Parameter
            Else
                Parameter(0) = GenerateInputParameter("@CurrentItem", adInteger, 4, CInt(tvLevels.SelectedItem.Tag))
                RunParametricStoredProcedure2Rec "RemoveGoodLevel2", Parameter
            End If
        Case 2 'no
    End Select
    On Error GoTo 0
    
    MyFormAddEditMode = ViewMode
    DefaultSetting
    SetFirstToolBar
    

ElseIf SSTab1.Tab = 1 Then
    
    Select Case clsStation.Language
        Case 0
            frmMsg.fwlblMsg.Caption = "¬Ì« „ÿ„∆‰Ìœ „Ì ŒÊ«ÂÌœ " & "'" & vsGood.TextMatrix(vsGood.Row, 2) & "'" & " —« Õ–› ﬂ‰Ìœø"
            frmMsg.fwBtn(0).Caption = "»·Ì"
            frmMsg.fwBtn(1).Caption = "ŒÌ—"
        Case 1
            frmMsg.fwlblMsg.Caption = "You are going to delete '" & vsGood.TextMatrix(vsGood.Row, 2) & "'" + vbNewLine + "Are you sure ?"
            frmMsg.fwBtn(0).Caption = "Yes"
            frmMsg.fwBtn(1).Caption = "No"
            frmMsg.fwlblMsg.Alignment = vbLeftJustify
    End Select
    
    frmMsg.Show vbModal
    
    If modgl.mvarMsgIdx = vbNo Then
        Exit Sub
    End If
    
    ReDim Parameter(1) As Parameter
    Parameter(0) = GenerateInputParameter("@GoodCode", adInteger, 4, vsGood.TextMatrix(vsGood.Row, 1))
    Parameter(1) = GenerateOutputParameter("@Result", adInteger, 4)
    Dim Result As Integer
    Result = RunParametricStoredProcedure("Delete_Good", Parameter)
    
    If Result = 0 Then
    
        Select Case clsStation.Language

            Case 0
                frmMsg.fwlblMsg.Caption = "›«ò Ê—Â«ÌÌ „— »ÿ »« «Ì‰ ò«·« ÊÃÊœ œ«—‰œ ‘„« ‰„Ì  Ê«‰Ìœ «Ì‰ ò«·« —« Õ–› ò‰Ìœ"
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
                frmMsg.fwlblMsg.Caption = "‘„« Ìò ò«·« —« Õ–› ò—œÌœ"
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
    
    FillvsGood
End If
Exit Sub
    
DBErrHandler:
    Select Case err.Number
        Case -2147217873
            frmMsg.fwlblMsg.Caption = "œ— «Ì‰ “Ì— ê—ÊÂ ò«·« ÊÃÊœ œ«—œ «» œ« ò«·«Â« —« Õ–› ò‰Ìœ "
            frmMsg.fwBtn(0).ButtonType = flwButtonOk
            frmMsg.fwBtn(0).Caption = "ﬁ»Ê·"
            frmMsg.fwBtn(1).Visible = False
            frmMsg.Show vbModal
    End Select
    
    MyFormAddEditMode = ViewMode
    DefaultSetting
    SetFirstToolBar
    
End Sub

Private Sub CmdChangeLevel1_Click()
    
    If Len(TxtLevel1New) < 2 Then
        ShowDisMessage "ê—ÊÂ «’·Ì »«Ìœ œÊ —ﬁ„Ì »«‘œ", 1500
        Exit Sub
    End If
    On Error GoTo DBErrHandler
    If TxtLevel1New.Text <> "" Then
        ReDim Parameter(3) As Parameter
        Parameter(0) = GenerateInputParameter("@OldLevel1", adInteger, 4, Val(TxtLevel1.Text))
        Parameter(1) = GenerateInputParameter("@NewLevel1", adInteger, 4, Val(TxtLevel1New.Text))
        Parameter(2) = GenerateInputParameter("@Replace", adBoolean, 1, 0)
        Parameter(3) = GenerateOutputParameter("@Update", adInteger, 4)
        Dim Result As Integer
        Result = RunParametricStoredProcedure("ChangeLevel1", Parameter)
                  
        If Result = 0 Then
            GoTo DBErrHandler
        Else
            frmDisMsg.lblMessage.Caption = "  €ÌÌ— ê—ÊÂÂ« «‰Ã«„ ‘œ "
            frmDisMsg.Timer1.Interval = 1000
            frmDisMsg.Timer1.Enabled = True
            frmDisMsg.Show vbModal
            DefaultSetting
        End If
    End If
Exit Sub
    
DBErrHandler:
    'PosConnection.RollbackTrans
        ReDim Parameter(3) As Parameter
        Parameter(0) = GenerateInputParameter("@OldLevel1", adInteger, 4, Val(TxtLevel1.Text))
        Parameter(1) = GenerateInputParameter("@NewLevel1", adInteger, 4, Val(TxtLevel1New.Text))
        Parameter(2) = GenerateInputParameter("@Replace", adBoolean, 1, 1)
        Parameter(3) = GenerateOutputParameter("@Update", adInteger, 4)
        Result = RunParametricStoredProcedure("ChangeLevel1", Parameter)
        If Result = 0 Then
            frmMsg.fwlblMsg.Caption = "œ—  €ÌÌ— ê—ÊÂÂ« „‘ò· ÊÃÊœ œ«—œ "
            frmMsg.fwBtn(0).ButtonType = flwButtonOk
            frmMsg.fwBtn(0).Caption = "ﬁ»Ê·"
            frmMsg.fwBtn(1).Visible = False
            frmMsg.Show vbModal
        Else
            frmDisMsg.lblMessage.Caption = "  €ÌÌ— ê—ÊÂÂ« Â„—«Â »« Ã«Ìê“Ì‰Ì «‰Ã«„ ‘œ "
            frmDisMsg.Timer1.Interval = 1000
            frmDisMsg.Timer1.Enabled = True
            frmDisMsg.Show vbModal
            DefaultSetting
        End If

End Sub

Private Sub CmdChangeLevel2_Click()
    On Error GoTo DBErrHandler
    
    If Len(TxtLevel2New) < 2 Then
        ShowDisMessage "ê—ÊÂ ›—⁄Ì »«Ìœ œÊ —ﬁ„Ì »«‘œ", 1500
        Exit Sub
    End If
    If Len(txtLevel1Primary) < 2 Then
        ShowDisMessage "ê—ÊÂ «’·Ì «‰ Œ«» ‰‘œÂ", 1500
        Exit Sub
    End If
    
    If TxtLevel2New.Text <> "" Then
        ReDim Parameter(4) As Parameter
        Parameter(0) = GenerateInputParameter("@OldLevel2", adInteger, 4, Val(TxtLevel2.Text))
        Parameter(1) = GenerateInputParameter("@NewLevel2", adInteger, 4, Val(txtLevel1Primary & TxtLevel2New.Text))
        Parameter(2) = GenerateInputParameter("@Level1", adInteger, 4, CmbLevel1New.ItemData(CmbLevel1New.ListIndex))
        Parameter(3) = GenerateInputParameter("@Replace", adBoolean, 1, 0)
        Parameter(4) = GenerateOutputParameter("@Update", adInteger, 4)
        Dim Result As Integer
        Result = RunParametricStoredProcedure("ChangeLevel2", Parameter)
        If Result = 0 Then
            GoTo DBErrHandler
        Else
            frmDisMsg.lblMessage.Caption = "  €ÌÌ— “Ì— ê—ÊÂÂ« Â„—«Â »«  €ÌÌ— òœ ò«·«Â««‰Ã«„ ‘œ "
            frmDisMsg.Timer1.Interval = 1000
            frmDisMsg.Timer1.Enabled = True
            frmDisMsg.Show vbModal
            DefaultSetting
        End If
    End If
Exit Sub
    
DBErrHandler:
   ' PosConnection.RollbackTrans
'        ReDim Parameter(4) As Parameter
'        Parameter(0) = GenerateInputParameter("@OldLevel2", adInteger, 4, Val(TxtLevel2.Text))
'        Parameter(1) = GenerateInputParameter("@NewLevel2", adInteger, 4, Val(txtLevel1Primary & TxtLevel2New.Text))
'        Parameter(2) = GenerateInputParameter("@Level1", adInteger, 4, CmbLevel1New.ItemData(CmbLevel1New.ListIndex))
'        Parameter(3) = GenerateInputParameter("@Replace", adBoolean, 1, 1)
'        Parameter(4) = GenerateOutputParameter("@Update", adInteger, 4)
'        Result = RunParametricStoredProcedure("ChangeLevel2", Parameter)
'        If Result = 0 Then
            frmMsg.fwlblMsg.Caption = "œ—  €ÌÌ— “Ì—ê—ÊÂÂ« „‘ò· ÊÃÊœ œ«—œ "
            frmMsg.fwBtn(0).ButtonType = flwButtonOk
            frmMsg.fwBtn(0).Caption = "ﬁ»Ê·"
            frmMsg.fwBtn(1).Visible = False
            frmMsg.Show vbModal
'        Else
'            frmDisMsg.lblMessage.Caption = "  €ÌÌ— “Ì— ê—ÊÂÂ« Â„—«Â »« Ã«Ìê“Ì‰Ì «‰Ã«„ ‘œ "
'            frmDisMsg.Timer1.Interval = 1000
'            frmDisMsg.Timer1.Enabled = True
'            frmDisMsg.Show vbModal
'            DefaultSetting
'        End If

End Sub

Private Sub cmdUpdateBuyPrice_Click()
    
    RunNonParametricStoredProcedure "Update_BuyPrice_by_LastPrice"
    ShowDisMessage "ﬁÌ„  Œ—Ìœ ò«·«Â« »« ¬Œ—Ì‰ ﬁÌ„  Œ—Ìœ ¬‰ ò«·« »Â —Ê“ ‘œ", 1500
    Cancel
    
End Sub

Private Sub Form_Activate()
    VarActForm = Me.Name
    SetFirstToolBar
    If GetSetting(strMainKey, Me.Name, "Flexgrid_Name") <> "" Then
        vsGood.Font.Name = GetSetting(strMainKey, Me.Name, "Flexgrid_Name")
    End If
    If GetSetting(strMainKey, Me.Name, "Flexgrid_Size") <> "" Then
        vsGood.Font.Size = GetSetting(strMainKey, Me.Name, "Flexgrid_Size")
    End If
    If GetSetting(strMainKey, Me.Name, "Flexgrid_Bold") <> "" Then
        vsGood.Font.Bold = GetSetting(strMainKey, Me.Name, "Flexgrid_Bold")
    End If
    
    
    UcFont1.FontName = vsGood.Font.Name
    UcFont1.FontSize = vsGood.Font.Size
    UcFont1.FontBold = vsGood.Font.Bold
    UcFont1.VarActForm = Me.Name
    
    On Error Resume Next
    For Each Obj In Me
        If Obj.Name <> "UcFont1" And Obj.Name <> "Image1" And Obj.Name <> "ResizeKit1" And Obj.Name <> "CommonDialog1" Then
            Obj.Font.Name = vsGood.Font.Name
            Obj.Font.Size = vsGood.Font.Size
            Obj.Font.Bold = vsGood.Font.Bold
        End If
    Next Obj
    
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case Shift
          Case 0
              Select Case KeyCode
''                  Case 13  ' Enter
''                    Sendkey "{Left}", False
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

    If ClsFormAccess.frmGood = False Then
        Unload Me
        Exit Sub
    End If
    CenterTop Me
    VarActForm = Me.Name
     
    fwBtnCustFind.Tag = -1
    
    UpdatelblSupplier

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
    vsGood.Cols = 27
    If SSTab1.Tab = 0 Then
        MyFormAddEditMode = ViewMode
        DefaultSetting
        SetFirstToolBar
    ElseIf SSTab1.Tab = 1 Then
        ChangeLanguage
        DefaultSetting
    End If
    
End Sub


Private Sub Form_Unload(Cancel As Integer)
    
    If NewGoodFlag = True Then
        NewGoodFlag = False
    End If
    VarActForm = ""
    
    If PosConnection.State = adStateOpen Then PosConnection.Close: Set PosConnection = Nothing
    Unload frmFindGoods
    SaveSetting strMainKey, Me.Name, "Left", Me.Left
    SaveSetting strMainKey, Me.Name, "Top", Me.Top

End Sub

Private Sub fwBtnCustFind_Click()
    Me.FindCust
    FillvsGood
End Sub
Public Sub FindCust()
    
    frmFindSupplier.Show vbModal
    
    If mvarcode <> 0 Then
        fwBtnCustFind.Tag = mvarcode
        mvarcode = 0
    Else
        fwBtnCustFind.Tag = -1
    End If
    UpdatelblSupplier
   
End Sub
Private Sub UpdatelblSupplier()

    If fwBtnCustFind.Tag <> "" Then
        Dim Rst As New ADODB.Recordset
        Dim mvarMemberShipId, mvarTel, mvarAddress, mvarDescription As String
        
    
        ReDim Parameter(0) As Parameter
        Parameter(0) = GenerateInputParameter("@Code", adInteger, 4, Val(fwBtnCustFind.Tag))
        Set Rst = RunParametricStoredProcedure2Rec("Get_vw_Suppliers", Parameter)
        
        If Rst.EOF = False And Rst.BOF = False Then
            
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
            fwBtnCustFind.Caption = Rst.Fields("FullName")
            mvarMemberShipId = "«‘ —«ﬂ : " & Rst.Fields("MemberShipId")
            mvarDescription = Rst.Fields("Description")
           
            
        End If
        
        Set Rst = Nothing
    End If
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
''''
''''    FillvsGood
''''
''''    MyFormAddEditMode = EnumAddEditMode.ViewMode
''''    SetFirstToolbar
    
End Sub

Private Sub lstGoodLevel1_Scroll()
  '  FillLstGoodLevel2
End Sub

Private Sub lstGoodLevel2_ItemCheck(Item As Integer)
    
    FillvsGood

End Sub

Public Sub ChangeLanguage()

Dim Obj As Object
If SSTab1.Tab = 0 Then
    DefaultSetting
ElseIf SSTab1.Tab = 1 Then

    Select Case clsStation.Language    ' LCase(mdifrm.Toolbar1.Buttons(25).Key)
        
        Case English
            
'            mdifrm.Toolbar1.Buttons(25).Key = "›«—”Ì"
'            ClsStation.Language = 1 'Switch to English
            
            Me.Caption = "Define Goods"
            mdifrm.Caption = clsArya.LatinCompany
            Me.RightToLeft = False
            
            For Each Obj In Me
                On Error Resume Next
                    Obj.RightToLeft = False
                  '  Obj.FontName = "times new roman"
'                    Obj.Alignment = vbLeftJustify
                On Error GoTo 0
            Next Obj
            lblGoodLevel1.Caption = "Goods Main Groups"
            lblGoodLevel2.Caption = "Goods SubGroups"
'            lblPrinters.Caption = "Print Formats"
        
        Case Farsi
            
'            mdifrm.Toolbar1.Buttons(25).Key = "English"
'            ClsStation.Language = 0 'Switch to Farsi
            
            Me.Caption = ""
            mdifrm.Caption = clsArya.Company
            Me.RightToLeft = True
            
            For Each Obj In Me
                On Error Resume Next
                    Obj.RightToLeft = True
             '       Obj.FontName = "traffic"
'                    Obj.Alignment = vbRightJustify
                On Error GoTo 0
            Next Obj
            
            lblGoodLevel1.Caption = " ê—ÊÂ «’·Ì ò«·«Â« - »Œ‘ Â«"
            lblGoodLevel2.Caption = "ê—ÊÂ ›—⁄Ì ò«·«Â«"
'            lblPrinters.Caption = "›—„ Â«Ì ç«Å"
            
    End Select
    
    'change the position of the object on the form
    If clsStation.Language = English Then
''''       lstGoodLevel1.Left = Me.Width - (lstGoodLevel1.Left + lstGoodLevel1.Width)
''''       lstGoodLevel2.Left = Me.Width - (lstGoodLevel2.Left + lstGoodLevel2.Width)
''''
''''       lblGoodLevel1.Left = Me.Width - (lblGoodLevel1.Left + lblGoodLevel1.Width)
''''       lblGoodLevel2.Left = Me.Width - (lblGoodLevel2.Left + lblGoodLevel2.Width)
    End If
    With vsGood
    
        .Cols = 27
        Select Case clsStation.Language
            Case Farsi
                
                .TextMatrix(0, 0) = "—œÌ›"
                .TextMatrix(0, 1) = "òœ"
                .TextMatrix(0, 2) = "‰«„ ò«·«"
                .TextMatrix(0, 3) = "‰«„ œ— ç«Å"
                .TextMatrix(0, 4) = "»«—òœ"
                .TextMatrix(0, 5) = "  ›Ì ›—Ê‘  "
                .TextMatrix(0, 6) = "›Ì ›—Ê‘ 2"
                .TextMatrix(0, 7) = "›Ì ›—Ê‘ 3"
                .TextMatrix(0, 8) = "›Ì ›—Ê‘ 4"
                .TextMatrix(0, 9) = "›Ì ›—Ê‘ 5"
                .TextMatrix(0, 10) = "›Ì ›—Ê‘ 6"
                .TextMatrix(0, 11) = "ﬁÌ„  Œ—Ìœ"
                .TextMatrix(0, 12) = "Ê«Õœ ò«·«"
                .TextMatrix(0, 13) = "‰Ê⁄ ò«·«"
                .TextMatrix(0, 14) = "ê—ÊÂ «’·Ì"
                .TextMatrix(0, 15) = "ê—ÊÂ ›—⁄Ì"
                .TextMatrix(0, 16) = "›—Ê‘‰œÂ"
                .TextMatrix(0, 17) = "Ê“‰ Ê«Õœ"
                .TextMatrix(0, 18) = " ⁄œ«œœ—Ê«Õœ"
                .TextMatrix(0, 19) = "ê—ÊÂ «’·Ì"
                .TextMatrix(0, 20) = "„‰Ê ‰„«Ì‘"
                .TextMatrix(0, 21) = "„”Ì— ¬ÌòÊ‰"
                .TextMatrix(0, 22) = " Ê÷ÌÕ« "
                .TextMatrix(0, 23) = "‰«„ œ— ç«Å 2"
                .TextMatrix(0, 24) = "‰«„ œ— ç«Å 3"
                .TextMatrix(0, 25) = "òœ ﬁ»·Ì"
                .TextMatrix(0, 26) = "    "
            
            Case English
                .TextMatrix(0, 0) = "Row"
                .TextMatrix(0, 1) = "Code"
                .TextMatrix(0, 2) = "Name"
                .TextMatrix(0, 3) = "Name in Print"
                .TextMatrix(0, 4) = "BarCode"
                .TextMatrix(0, 5) = "Price"
                .TextMatrix(0, 6) = "Price_2"
                .TextMatrix(0, 7) = "Price_3"
                .TextMatrix(0, 8) = "Price_4"
                .TextMatrix(0, 9) = "Price_5"
                .TextMatrix(0, 10) = "Price_6"
                .TextMatrix(0, 11) = "Purchase Price"
                .TextMatrix(0, 12) = "Good Unit"
                .TextMatrix(0, 13) = "Good Type"
                .TextMatrix(0, 14) = "Level1"
                .TextMatrix(0, 15) = "Level2"
                .TextMatrix(0, 16) = "Supplier"
                .TextMatrix(0, 17) = "Weight"
                .TextMatrix(0, 18) = "NoOfUnit"
                .TextMatrix(0, 19) = "Main Group"
                .TextMatrix(0, 20) = "GoodMenu"
                .TextMatrix(0, 21) = "PicturePath"
                .TextMatrix(0, 22) = "Description"
                .TextMatrix(0, 23) = "Name in Print2"
                .TextMatrix(0, 24) = "Name in Print3"
                .TextMatrix(0, 25) = " Code"
                .TextMatrix(0, 26) = "     "
        
        End Select
        .ColFormat(5) = "###,###"
        .ColFormat(6) = "###,###"
        .ColFormat(7) = "###,###"
        .ColFormat(8) = "###,###"
        .ColFormat(9) = "###,###"
        .ColFormat(10) = "###,###"
        .ColFormat(11) = "###,###"
         For i = 0 To .Cols - 1
            .ColWidth(i) = Val(GetSetting(strMainKey, "frmGood_vsGoods", "Col" & i))
            If .ColWidth(i) = 0 Then
                .ColWidth(i) = .Width / 6     'Row
            End If
         Next i
        .ColAlignment(-1) = flexAlignCenterCenter
        Select Case clsStation.Language
            Case 0
                .ColAlignment(2) = flexAlignRightCenter
                .ColAlignment(3) = flexAlignRightCenter
                .ColAlignment(22) = flexAlignRightCenter
                .ColAlignment(24) = flexAlignRightCenter
                .ColAlignment(22) = flexAlignRightCenter
            Case 1
                .ColAlignment(2) = flexAlignLeftCenter
                .ColAlignment(3) = flexAlignLeftCenter
                .ColAlignment(22) = flexAlignLeftCenter
                .ColAlignment(23) = flexAlignLeftCenter
                .ColAlignment(24) = flexAlignLeftCenter
        End Select
        .ColAlignment(21) = flexAlignLeftCenter
        .ColDataType(19) = flexDTBoolean
''''        .ColHidden(3) = True
        If (Val(strCategory) >= 0 And Val(strCategory) <= 7) Then
          ' .ColHidden(17) = True
           .ColHidden(18) = True
        Else
'''           .ColHidden(19) = True
        End If
        If clsArya.MultiPrice = False Then
           .ColHidden(6) = True
           .ColHidden(7) = True
           .ColHidden(8) = True
           .ColHidden(9) = True
           .ColHidden(10) = True
        Else
            If clsStation.MaxPrices = 5 Then
                .ColHidden(10) = True
            ElseIf clsStation.MaxPrices = 4 Then
                .ColHidden(9) = True
                .ColHidden(10) = True
            ElseIf clsStation.MaxPrices = 3 Then
                .ColHidden(8) = True
                .ColHidden(9) = True
                .ColHidden(10) = True
            ElseIf clsStation.MaxPrices = 2 Then
                .ColHidden(7) = True
                .ColHidden(8) = True
                .ColHidden(9) = True
                .ColHidden(10) = True
            ElseIf clsStation.MaxPrices = 1 Then
                .ColHidden(6) = True
                .ColHidden(7) = True
                .ColHidden(8) = True
                .ColHidden(9) = True
                .ColHidden(10) = True
            End If
        End If
        .FixedAlignment(-1) = flexAlignCenterCenter
        .ColSort(5) = flexSortNumericAscending + flexSortNumericDescending
        .FocusRect = flexFocusHeavy
'        .ColHidden(1) = True
'        .AutoSizeMode = flexAutoSizeColWidth
'        .AutoSize 0, .Cols - 2
'        If .ColWidth(2) > 3500 Then .ColWidth(2) = 3500
'        If .ColWidth(3) > 3000 Then .ColWidth(3) = 3000
        .AutoSearch = flexSearchFromCursor
        .ScrollBars = flexScrollBarBoth
        .AllowUserResizing = flexResizeColumns
        
        
        ReDim Parameter(0) As Parameter
        Parameter(0) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
        Set rctmp = RunParametricStoredProcedure2Rec("GetGoodLevel1_Description", Parameter)
        
        .ColComboList(14) = .BuildComboList(rctmp, "Description", "Code")
        ReDim Parameter(1) As Parameter
        Parameter(0) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
        Parameter(1) = GenerateInputParameter("@CurrentItem", adInteger, 4, 0)
        Set rctmp = RunParametricStoredProcedure2Rec("GetGoodLevel2_Description", Parameter)
        
        .ColComboList(15) = .BuildComboList(rctmp, "Description", "Code")
        
        ReDim Parameter(0) As Parameter
        Parameter(0) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
        Set rctmp = RunParametricStoredProcedure2Rec("GetUnitGood", Parameter)
       
        s = .BuildComboList(rctmp, "Description", "Code")
        .ColComboList(12) = s
        
        ReDim Parameter(0) As Parameter
        Parameter(0) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
        Set rctmp = RunParametricStoredProcedure2Rec("GetGoodType", Parameter)
      
        s = .BuildComboList(rctmp, "Description", "Code")
        .ColComboList(13) = s
        
        ReDim Parameter(0) As Parameter
        Parameter(0) = GenerateInputParameter("@Branch", adInteger, 4, Null)
        Set rctmp = RunParametricStoredProcedure2Rec("Get_All_Supplier", Parameter)
        .ColComboList(16) = "#-1" & "; „ ›—ﬁÂ|"
        .ColComboList(16) = .ColComboList(16) & .BuildComboList(rctmp, "Name", "Code")
        
        Set rctmp = RunStoredProcedure2RecordSet("Get_tblTotal_GoodShow")
        .ColComboList(20) = .BuildComboList(rctmp, "Category", "AutoId")
    
    End With
    rctmp.Close
    
    FillLstGoodLevel1
    
    Set rctmp = Nothing
  
    SetFirstToolBar
End If
End Sub

Private Sub ResizeKit1_ExitResize(ByVal XScale As Double, ByVal YScale As Double)

    If formloadFlag = True Then
        SaveSetting strMainKey, Me.Name, "Height", Me.Height
        SaveSetting strMainKey, Me.Name, "Width", Me.Width
    End If

End Sub
Private Sub SSTab1_Click(PreviousTab As Integer)
    If SSTab1.Tab = 0 Then
        If ClsFormAccess.frmCodingGood = False Then
            Unload Me
            Exit Sub
        End If
        MyFormAddEditMode = ViewMode
    ElseIf SSTab1.Tab = 1 Then
        If ClsFormAccess.frmGood = False Then
            Unload Me
            Exit Sub
        End If
        MyFormAddEditMode = AddMode
    End If
    ChangeLanguage
    SetFirstToolBar
    Cancel

End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call PresetScreenSaver
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call PresetScreenSaver
End Sub

Private Sub tvLevels_NodeClick(ByVal node As MSComctlLib.node)
    TxtLevel1.Text = ""
    TxtLevel2.Text = ""
    TxtLevel1New.Text = ""
    TxtLevel2New.Text = ""
    txtLevel1Primary.Text = ""
    CmbLevel1New.ListIndex = -1
    
    Dim i As Integer
    If tvLevels.Nodes.Count = 0 Then Exit Sub
    i = InStr(1, tvLevels.SelectedItem.Text, " - ")
    If InStr(1, tvLevels.SelectedItem.FullPath, "\") = 0 Then 'root
         
'        If i > 0 Then  '
'            tvLevels.SelectedItem.Text = Trim(Mid(tvLevels.SelectedItem.Text, i + 2))
'        End If
        TxtLevel1.Text = tvLevels.SelectedItem.Tag
    Else
        
'        If i > 0 Then  '
'            tvLevels.SelectedItem.Text = Trim(Mid(tvLevels.SelectedItem.Text, i + 2))
'        End If
        
        TxtLevel2.Text = tvLevels.SelectedItem.Tag
        For i = 0 To CmbLevel1New.ListCount - 1
            CmbLevel1New.ListIndex = i
            If tvLevels.SelectedItem.Parent.Tag = CmbLevel1New.ItemData(CmbLevel1New.ListIndex) Then
                Exit For
            End If
        Next
    End If

End Sub

Private Sub txtBarcode_Change()
    If Right(txtBarcode.Text, 1) = "/" Then
        txtBarcode.Text = Left(txtBarcode.Text, Len(txtBarcode.Text) - 1)
    ElseIf Left(txtBarcode.Text, 1) = "/" Then
        txtBarcode.Text = Right(txtBarcode.Text, Len(txtBarcode.Text) - 1)
    End If
    If Len(txtBarcode.Text) > 2 Then
        If Asc(Mid(txtBarcode.Text, Len(txtBarcode.Text) - 1, 1)) = 13 Then
            txtBarcode.Text = Left(txtBarcode.Text, Len(txtBarcode.Text) - 2)
        End If
    End If
    
    i = vsGood.FindRow(Trim(txtBarcode.Text), 1, 4, True, True)
    If i > 0 Then
        vsGood.Row = i
        vsGood.ShowCell i, 4
    Else
        vsGood.Row = 0
        vsGood.ShowCell 0, 0
    End If
    
End Sub

Private Sub txtBarcode_GotFocus()
    txtBarcode.Text = ""
 '   vsGood.Select vsGood.Row, 20
 '   vsGood.Sort = flexSortGenericAscending

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
                        vsGood.ShowCell i, 4
                        vsGood.Row = i
                        vsGood.Col = 4
               '         vsGood.Selec vsGood.Row, vsGood.Col
                        vsGood.EditCell
                        
                    End If
            End Select
    
    End Select

End Sub


Private Sub TxtLevel1New_Change()
    If SSTab1.Tab = 1 Then Exit Sub
    If Len(TxtLevel1New.Text) = 3 Then
        TxtLevel1New.Text = Val(Left(TxtLevel1New.Text, 2))
        ShowMessage "òœ ê—ÊÂ «’·Ì »«Ìœ »Ì‰ 11 Ê 99  »«‘œ", True, False, "ﬁ»Ê·", ""
    ElseIf Val(Len(TxtLevel1New.Text)) = 2 Then
        If Val(TxtLevel1New.Text) <= 10 Then
        ShowMessage "òœ ê—ÊÂ «’·Ì »«Ìœ »Ì‰ 11 Ê 99  »«‘œ", True, False, "ﬁ»Ê·", ""
        TxtLevel1New.Text = ""
        End If
    End If
End Sub

Private Sub TxtLevel1New_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 8 And Shift = 0 Then
    If Len(TxtLevel1New.Text) > 1 Then TxtLevel1New.Text = Left(TxtLevel1New.Text, Len(TxtLevel1New.Text) - 1)
End If
End Sub

Private Sub TxtLevel2New_Change()
    
    If SSTab1.Tab = 1 Then Exit Sub
    If Len(TxtLevel2New.Text) >= 3 Then
        TxtLevel2New.Text = Val(Left(TxtLevel2New.Text, 2))
        ShowMessage "òœ ê—ÊÂ ›—⁄Ì »«Ìœ »Ì‰ 01 Ê 99 Ê  «»⁄Ì «“ òœ ê—ÊÂ «’·Ì »«‘œ", True, False, "ﬁ»Ê·", ""
'    ElseIf Val(Len(TxtLevel1New.Text)) = 2 Then
'        If Val(TxtLevel2New.Text) >= 10 Then
'            frmMsg.fwlblMsg.Caption = "òœ ê—ÊÂ ›—⁄Ì »«Ìœ »Ì‰ 01 Ê 99 Ê  «»⁄Ì «“ òœ ê—ÊÂ «’·Ì »«‘œ"
'            frmMsg.fwBtn(0).Visible = False
'            frmMsg.fwBtn(1).ButtonType = flwButtonOk
'            frmMsg.fwBtn(1).Caption = "ﬁ»Ê·"
'            frmMsg.Show vbModal
'            TxtLevel2New.Text = ""
'        End If
    End If
End Sub

Private Sub TxtLevel2New_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error Resume Next
    If KeyCode = 8 And Shift = 0 Then
        If Len(TxtLevel2New.Text) > 1 Then TxtLevel2New.Text = Left(TxtLevel2New.Text, Len(TxtLevel2New.Text) - 1)
    End If
End Sub

Private Sub UcFont1_FontProperty(m_FontName As Variant, m_FontSize As Variant, m_FontBold As Variant)
    On Error Resume Next
'    vsGood.Font.Name = m_FontName
'    vsGood.Font.Size = m_FontSize
'    vsGood.Font.Bold = m_FontBold
    For Each Obj In Me
        Obj.Font.Name = m_FontName
        Obj.Font.Size = m_FontSize
        Obj.Font.Bold = m_FontBold
        '  Obj.FontName = "times new roman"
        '  Obj.Alignment = vbLeftJustify
    Next Obj
    vsGood.Refresh
    
End Sub

Private Sub vsGood_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
''''    If flgRow = True And (OldCol = 12 Or OldCol = 13 Or OldCol = 14 Or OldCol = 15 Or OldCol = 16) Then
''''        flgRow = False
''''        Exit Sub
''''   End If
    With vsGood
''        If (.TextMatrix(OldRow, 0) = "*" Or MyFormAddEditMode = EnumAddEditMode.EditMode) And OldCol > 1 And tmpTextMatrix <> .TextMatrix(OldRow, OldCol) Then
        If (.TextMatrix(OldRow, 0) = "*" Or MyFormAddEditMode = EnumAddEditMode.EditMode) And tmpTextMatrix <> .TextMatrix(OldRow, OldCol) Then
        
            If OldCol = 2 Or OldCol = 3 Then

                If Trim(.TextMatrix(OldRow, 2)) = "" And OldCol <> 2 Then
                    .TextMatrix(OldRow, 2) = Trim(.TextMatrix(OldRow, OldCol))
                End If

                If Trim(.TextMatrix(OldRow, 3)) = "" And Col <> 3 Then
                    .TextMatrix(OldRow, 3) = Trim(.TextMatrix(OldRow, OldCol))
                End If

                If Trim(.TextMatrix(OldRow, OldCol)) = "" Then
                
                    If .TextMatrix(OldRow, 2) <> "" Then
                        .TextMatrix(OldRow, OldCol) = Trim(.TextMatrix(OldRow, 2))
                    ElseIf .TextMatrix(OldRow, 3) <> "" Then
                        .TextMatrix(OldRow, OldCol) = Trim(.TextMatrix(OldRow, 3))
                    ElseIf .TextMatrix(OldRow, 4) <> "" Then
                        .TextMatrix(OldRow, OldCol) = Trim(.TextMatrix(OldRow, 4))
                    End If
                    
                End If

            End If
            Dim LongTemp As Integer
            
            If OldCol = 4 Then
                If InStr(1, .TextMatrix(OldRow, 4), "/", 1) Then
                    LongTemp = InStr(2, .TextMatrix(OldRow, 4), "/", 1)
                    If LongTemp > 2 Then
                       .TextMatrix(OldRow, 4) = Mid(.TextMatrix(OldRow, 4), 2, LongTemp - 2)
                    End If
                End If
                If GetGoodBarcode(.TextMatrix(OldRow, 4), Val(Trim(.TextMatrix(OldRow, 1)))) = True Then
                    frmMsg.fwlblMsg.Caption = " . «Ì‰ »«—ﬂœ ﬁ»·« œ— ”Ì” „  ⁄—Ì› ‘œÂ «”  "
                    frmMsg.fwBtn(0).Caption = "ﬁ»Ê·"
                    frmMsg.fwBtn(1).Visible = False
                    frmMsg.fwBtn(0).ButtonType = flwButtonOk
                    frmMsg.Show vbModal
                '    .TextMatrix(.Row, 4) = ""
                    .Row = OldRow
                    .Col = OldCol
                    .Select .Row, .Col
                    .EditCell
                End If
            End If
            
            If NewCol = 12 Or NewCol = 13 Or NewCol = 14 Or NewCol = 15 Or NewCol = 16 Then
                If MyFormAddEditMode = EnumAddEditMode.EditMode And InStr(.TextMatrix(OldRow, 0), "*") = 0 Then
                    .TextMatrix(OldRow, 0) = Trim(.TextMatrix(OldRow, 0)) & "*"
                End If
                
                If OldCol = 14 Or OldCol = 15 Then
                    ReDim Parameter(1) As Parameter
                    Parameter(0) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
                    Parameter(1) = GenerateInputParameter("@CurrentItem", adInteger, 4, .TextMatrix(NewRow, 14))
                    Set rctmp = RunParametricStoredProcedure2Rec("GetGoodLevel2_Description", Parameter)
                    
                    .ColComboList(15) = .BuildComboList(rctmp, "Description", "Code")
                    rctmp.Close
                
                End If
            End If
''            If MyFormAddEditMode = EnumAddEditMode.EditMode And InStr(.TextMatrix(OldRow, 0), "*") = 0 Then
''                .TextMatrix(OldRow, 0) = Trim(.TextMatrix(OldRow, 0)) & "*"
''            End If
            
        Else

        End If
        

    End With
End Sub

Private Sub vsGood_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    With vsGood
        If (.TextMatrix(Row, 0) = "*" Or MyFormAddEditMode = EnumAddEditMode.EditMode) And Col > 1 And tmpTextMatrix <> .TextMatrix(Row, Col) Then

            If Col = 2 Or Col = 3 Then

                If Trim(.TextMatrix(Row, 2)) = "" And Col <> 2 Then
                    .TextMatrix(Row, 2) = Trim(.TextMatrix(Row, Col))
                End If

                If Trim(.TextMatrix(Row, 3)) = "" And Col <> 3 Then
                    .TextMatrix(Row, 3) = Trim(.TextMatrix(Row, Col))
                End If

                If Trim(.TextMatrix(Row, Col)) = "" Then

                    If .TextMatrix(Row, 2) <> "" Then
                        .TextMatrix(Row, Col) = Trim(.TextMatrix(Row, 2))
                    ElseIf .TextMatrix(Row, 3) <> "" Then
                        .TextMatrix(Row, Col) = Trim(.TextMatrix(Row, 3))
                    ElseIf .TextMatrix(Row, 4) <> "" Then
                        .TextMatrix(Row, Col) = Trim(.TextMatrix(Row, 4))
                    End If

                End If

            End If
            Dim LongTemp As Integer

            If Col = 4 Then
                If InStr(1, .TextMatrix(.Row, 4), "/", 1) Then
                    LongTemp = InStr(2, .TextMatrix(.Row, 4), "/", 1)
                    If LongTemp > 2 Then
                       .TextMatrix(.Row, 4) = Mid(.TextMatrix(.Row, 4), 2, LongTemp - 2)
                    End If
                End If
                If GetGoodBarcode(.TextMatrix(.Row, 4), Val(Trim(.TextMatrix(.Row, 1)))) = True Then
                    frmMsg.fwlblMsg.Caption = " . «Ì‰ »«—ﬂœ ﬁ»·« œ— ”Ì” „  ⁄—Ì› ‘œÂ «”  "
                    frmMsg.fwBtn(0).Caption = "ﬁ»Ê·"
                    frmMsg.fwBtn(1).Visible = False
                    frmMsg.fwBtn(0).ButtonType = flwButtonOk
                    frmMsg.Show vbModal
                '    .TextMatrix(.Row, 4) = ""
                    .Row = Row
                    .Col = Col
                    .Select .Row, .Col
                    .EditCell
                End If
            End If
            If Col = 14 Or Col = 15 Then
                ReDim Parameter(1) As Parameter
                Parameter(0) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
                Parameter(1) = GenerateInputParameter("@CurrentItem", adInteger, 4, .TextMatrix(.Row, 14))
                Set rctmp = RunParametricStoredProcedure2Rec("GetGoodLevel2_Description", Parameter)

                .ColComboList(15) = .BuildComboList(rctmp, "Description", "Code")
                rctmp.Close
            End If
            If MyFormAddEditMode = EnumAddEditMode.EditMode And InStr(.TextMatrix(Row, 0), "*") = 0 Then
                .TextMatrix(Row, 0) = Trim(.TextMatrix(Row, 0)) & "*"
            End If

        Else

        End If


    End With


End Sub

Private Sub vsGood_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    With vsGood
        If Col >= 0 Then
            For i = 0 To .Cols - 1
                SaveSetting strMainKey, "frmGood_vsGoods", "Col" & i, .ColWidth(i)
            Next
        End If
    End With

End Sub

Private Sub vsGood_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = 25 Then Cancel = True
'    If MyFormAddEditMode = EditMode And Col = 1 Then Cancel = True
    tmpTextMatrix = vsGood.TextMatrix(Row, Col)
End Sub

Private Sub vsGood_KeyDown(KeyCode As Integer, Shift As Integer)
    With vsGood
        
        If (.TextMatrix(.Row, 0) = "*" Or MyFormAddEditMode = EnumAddEditMode.EditMode) And .Col > 0 Then
            If .Col = 16 And KeyCode = 13 And Shift = 0 Then
                .Col = 23
            ElseIf .Col = 23 And KeyCode = 13 And Shift = 0 Then
                .Col = 24
            ElseIf .Col = 26 And KeyCode = 13 And Shift = 0 Then
                If ((Trim(.TextMatrix(.Row, 2)) = "" And Trim(.TextMatrix(.Row, 3)) = "") Or Trim(.TextMatrix(.Row, 5)) = "") Or .Cell(flexcpText, .Row, 12) = "" Or .Cell(flexcpText, .Row, 13) = "" Then
                Else
                    BeforeAdd
                    Exit Sub
                End If
            ElseIf KeyCode = 38 Or KeyCode = 39 Or KeyCode = 40 Then
            ElseIf KeyCode = 37 And (.Col = 18 Or .Col = 19) Then
            ElseIf KeyCode = 32 And .Col = 19 Then
            ElseIf KeyCode = 13 And Not (.Col = 12 Or .Col = 13 Or .Col = 16 Or .Col = 20) Then
                Sendkey "{Left}", True
            ElseIf KeyCode = 115 And (.Col = 12 Or .Col = 13 Or .Col = 16 Or .Col = 20) Then
                Sendkey "{Left}", True
            Else
                .Select .Row, .Col
                .EditCell
            End If
        End If
    End With
    
End Sub


Private Sub vsGood_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    With vsGood
        If KeyAscii = 39 Then KeyAscii = 0
        If KeyAscii = 13 And (Col = 12 Or Col = 13 Or Col = 20) Then
            Sendkey "{F4}", True
        End If
        If Col = 5 And IsNumeric(Chr(KeyAscii)) = False And (KeyAscii <> 8 And KeyAscii <> 13) Then
            
            KeyAscii = 0
            
        ElseIf MyFormAddEditMode = EditMode Then
            
            If Row > 0 And InStr(1, .TextMatrix(Row, 0), "*") = 0 Then
                .TextMatrix(Row, 0) = .TextMatrix(Row, 0) & "*"
            End If
            
        End If
        
        If .Col = 21 Then
'''''''''''''''''''''
            CommonDialog1.InitDir = App.Path & "\IMAGE\FOOD_PIC"
            CommonDialog1.Filter = "Pictures (*.bmp;*.ico;*.gif;*.jpg;*.jpeg)|*.bmp;*.ico;*.gif;*.jpg;*.jpeg"
            CommonDialog1.CancelError = True
            On Error GoTo ErrorHandler
            CommonDialog1.ShowOpen
           
            Dim fso As New FileSystemObject
            If fso.FileExists(CommonDialog1.InitDir & "\" & fso.GetBaseName(CommonDialog1.Filename) & "." & fso.GetExtensionName(CommonDialog1.Filename)) = True And LCase(CommonDialog1.Filename) <> LCase(CommonDialog1.InitDir & "\" & fso.GetBaseName(CommonDialog1.Filename) & "." & fso.GetExtensionName(CommonDialog1.Filename)) Then
                Dim f As file
                
                Set f = fso.GetFile(CommonDialog1.InitDir & "\" & fso.GetBaseName(CommonDialog1.Filename) & "." & fso.GetExtensionName(CommonDialog1.Filename))
                If Mid(ConvertToBin(f.Attributes, 8), 8, 1) = "1" Then
                    'If f.Attributes = ReadOnly Then
                    frmMsg.fwBtn(1).Caption = "ŒÌ—"
                    frmMsg.fwBtn(0).Caption = "»·Â"
                    frmMsg.fwBtn(1).Visible = True
                    frmMsg.fwBtn(0).Visible = True
                    frmMsg.fwlblMsg.Caption = "„ÊÃÊœ „Ì »«‘œ" & CommonDialog1.InitDir & "«Ì‰ ›«Ì· œ— " & vbCrLf & "¬Ì« „«Ì·Ìœ «“ ¬‰ «” ›«œÂ ‰„«ÌÌœø"
                    frmMsg.Show vbModal
                    f.Attributes = Normal
                    If mvarMsgIdx = vbYes Then
                
                    Else
                        fso.CopyFile CommonDialog1.Filename, CommonDialog1.InitDir & "\" & fso.GetBaseName(CommonDialog1.Filename) & "." & fso.GetExtensionName(CommonDialog1.Filename), True
                    End If
                End If
            ElseIf LCase(CommonDialog1.Filename) <> LCase(CommonDialog1.InitDir & "\" & fso.GetBaseName(CommonDialog1.Filename) & "." & fso.GetExtensionName(CommonDialog1.Filename)) Then
                fso.CopyFile CommonDialog1.Filename, CommonDialog1.InitDir & "\" & fso.GetBaseName(CommonDialog1.Filename) & "." & fso.GetExtensionName(CommonDialog1.Filename), True
            End If
            
            .TextMatrix(.Row, .Col) = "\IMAGE\FOOD_PIC\" & fso.GetBaseName(CommonDialog1.Filename) & "." & fso.GetExtensionName(CommonDialog1.Filename)

''''''''''''''''''''
        End If
    End With
Exit Sub
ErrorHandler:
         vsGood.TextMatrix(vsGood.Row, vsGood.Col) = ""

End Sub


Private Sub vsGood_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    With vsGood
        If MyFormAddEditMode = EnumAddEditMode.ViewMode Then Exit Sub
'        If (.TextMatrix(.Row, 0) = "*" Or MyFormAddEditMode = EnumAddEditMode.EditMode) And .Col > 1 Then
'            .Select .Row, .Col
'            If .Col = 19 Then Exit Sub
'            .EditCell
'        End If
    
'        If .Row > 0 And InStr(1, .TextMatrix(.Row, 0), "*") = 0 Then
'            .TextMatrix(.Row, 0) = .TextMatrix(.Row, 0) & "*"
'        End If
        If .Col = 21 Then
'''''''''''''''''''''
            CommonDialog1.InitDir = App.Path & "\IMAGE\FOOD_PIC"
            CommonDialog1.Filter = "Pictures (*.bmp;*.ico;*.gif;*.jpg;*.jpeg)|*.bmp;*.ico;*.gif;*.jpg;*.jpeg"
            CommonDialog1.CancelError = True
            On Error GoTo ErrorHandler
            CommonDialog1.ShowOpen
           
            Dim fso As New FileSystemObject
            If fso.FileExists(CommonDialog1.InitDir & "\" & fso.GetBaseName(CommonDialog1.Filename) & "." & fso.GetExtensionName(CommonDialog1.Filename)) = True And LCase(CommonDialog1.Filename) <> LCase(CommonDialog1.InitDir & "\" & fso.GetBaseName(CommonDialog1.Filename) & "." & fso.GetExtensionName(CommonDialog1.Filename)) Then
                Dim f As file
                
                Set f = fso.GetFile(CommonDialog1.InitDir & "\" & fso.GetBaseName(CommonDialog1.Filename) & "." & fso.GetExtensionName(CommonDialog1.Filename))
                If Mid(ConvertToBin(f.Attributes, 8), 8, 1) = "1" Then
                    'If f.Attributes = ReadOnly Then
                    frmMsg.fwBtn(1).Caption = "ŒÌ—"
                    frmMsg.fwBtn(0).Caption = "»·Â"
                    frmMsg.fwBtn(1).Visible = True
                    frmMsg.fwBtn(0).Visible = True
                    frmMsg.fwlblMsg.Caption = "„ÊÃÊœ „Ì »«‘œ" & CommonDialog1.InitDir & "«Ì‰ ›«Ì· œ— " & vbCrLf & "¬Ì« „«Ì·Ìœ «“ ¬‰ «” ›«œÂ ‰„«ÌÌœø"
                    frmMsg.Show vbModal
                    f.Attributes = Normal
                    If mvarMsgIdx = vbYes Then
                
                    Else
                        fso.CopyFile CommonDialog1.Filename, CommonDialog1.InitDir & "\" & fso.GetBaseName(CommonDialog1.Filename) & "." & fso.GetExtensionName(CommonDialog1.Filename), True
                    End If
                End If
            ElseIf LCase(CommonDialog1.Filename) <> LCase(CommonDialog1.InitDir & "\" & fso.GetBaseName(CommonDialog1.Filename) & "." & fso.GetExtensionName(CommonDialog1.Filename)) Then
                fso.CopyFile CommonDialog1.Filename, CommonDialog1.InitDir & "\" & fso.GetBaseName(CommonDialog1.Filename) & "." & fso.GetExtensionName(CommonDialog1.Filename), True
            End If
            
            .TextMatrix(.Row, .Col) = "\IMAGE\FOOD_PIC\" & fso.GetBaseName(CommonDialog1.Filename) & "." & fso.GetExtensionName(CommonDialog1.Filename)

            If MyFormAddEditMode = EditMode Then
                
                If InStr(1, .TextMatrix(.Row, 0), "*") = 0 Then
                    .TextMatrix(.Row, 0) = .TextMatrix(.Row, 0) & "*"
                End If
                
            End If
''''''''''''''''''''
        End If
    
    End With

Exit Sub
ErrorHandler:
         vsGood.TextMatrix(vsGood.Row, vsGood.Col) = ""
    
End Sub


Private Sub vsGood_RowColChange()

    Image1.Picture = LoadPicture("")
    With vsGood
        If .Row = 0 Then Exit Sub
        If .TextMatrix(.Row, 21) = "" Then Exit Sub
'        If IsNull(rctmp.Fields("Picture").Value) Then
            If filetemp.FileExists(App.Path & .TextMatrix(.Row, 21)) Then
                Image1.Picture = LoadPicture(App.Path & .TextMatrix(.Row, 21))
            End If
'        Else
'            Set strStream = New ADODB.Stream
'            strStream.Type = adTypeBinary
'            strStream.Open
'            strStream.Write rctmp.Fields("Picture").Value
'            strStream.SaveToFile "C:\Temp.bmp", adSaveCreateOverWrite
'            Image1.Picture = LoadPicture("C:\Temp.bmp")
'            Kill ("C:\Temp.bmp")
''            LoadPictureFromDB = True
'            Set strStream = Nothing
'        End If
    End With

End Sub

Private Sub vsGood_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)

If Col = 14 Then Exit Sub
With vsGood
    .Row = Row
    .Col = Col
End With

End Sub

Private Function GetGoodBarcode(Code As String, GoodCode As Double)
    Dim ReturnValue As Boolean
    ReturnValue = False
    ReDim Parameter(1) As Parameter
    Parameter(0) = GenerateInputParameter("@Barcode", adVarWChar, 50, Code)
    Parameter(1) = GenerateInputParameter("@GoodCode", adInteger, 4, GoodCode)
    Set rctmp = RunParametricStoredProcedure2Rec("Get_Good_Barcode_Check", Parameter)
    
    If Not (rctmp.BOF Or rctmp.EOF) Then
        ReturnValue = True
    End If
    GetGoodBarcode = ReturnValue
    
End Function

