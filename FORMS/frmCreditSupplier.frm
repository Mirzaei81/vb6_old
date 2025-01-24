VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{7AEDC602-D94C-11D1-BB7A-00E0290EA3C9}#1.0#0"; "ResizeKit.ocx"
Object = "{75D4F767-8785-11D3-93AD-0000832EF44D}#2.14#0"; "FAST2001.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmCreditSupplier 
   ClientHeight    =   10935
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11955
   Icon            =   "frmCreditSupplier.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   10935
   ScaleWidth      =   11955
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
      Left            =   120
      RightToLeft     =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   38
      Top             =   0
      Width           =   1695
   End
   Begin VB.Frame Frame4 
      Height          =   4095
      Left            =   9240
      TabIndex        =   32
      Top             =   5160
      Width           =   2655
      Begin FLWCtrls.FWScrollText fwScrollTextCust 
         Height          =   555
         Left            =   120
         TabIndex        =   33
         Top             =   240
         Width           =   2355
         _ExtentX        =   4154
         _ExtentY        =   979
         Caption         =   ""
         BorderStyle     =   9
         FontName        =   "Nazanin"
         FontBold        =   -1  'True
         FontSize        =   12
      End
      Begin VB.Label lblPreRemaining 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
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
         ForeColor       =   &H80000008&
         Height          =   450
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   42
         Top             =   2880
         Width           =   2325
      End
      Begin VB.Label fwStatusBarCust 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Nazanin"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   2010
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   41
         Top             =   840
         Width           =   2325
      End
      Begin VB.Label lblRemaining 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000013&
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
         ForeColor       =   &H80000008&
         Height          =   450
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   34
         Top             =   3360
         Width           =   2325
      End
   End
   Begin VB.Frame Frame3 
      Height          =   4455
      Left            =   9240
      TabIndex        =   25
      Top             =   600
      Width           =   2655
      Begin VB.CommandButton cmdTurnOver 
         Caption         =   "ê—œ‘ Õ”«» «Ì‰  «„Ì‰ ò‰‰œÂ"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         TabIndex        =   44
         Top             =   3720
         Width           =   2415
      End
      Begin VB.ComboBox cmbBranch 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   43
         Text            =   "Combo1"
         Top             =   240
         Width           =   2355
      End
      Begin VB.CommandButton cmdUpdate 
         BackColor       =   &H008080FF&
         Caption         =   "»Â —Ê“ —”«‰Ì «ÿ·«⁄«   «„Ì‰ ﬂ‰‰œÂ"
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
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   26
         Top             =   2880
         Width           =   2415
      End
      Begin FLWCtrls.FWCoolButton fwBtnCustFind 
         Height          =   810
         Left            =   120
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   720
         Width           =   2370
         _ExtentX        =   4180
         _ExtentY        =   1429
         BackColor       =   16777088
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "frmCreditSupplier.frx":A4C2
         PictureAlign    =   4
         Caption         =   " «„Ì‰ ﬂ‰‰œÂ"
         MaskColor       =   -2147483633
      End
      Begin MSMask.MaskEdBox txtDatefrom 
         Height          =   585
         Left            =   120
         TabIndex        =   28
         Top             =   1560
         Width           =   1395
         _ExtentX        =   2461
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
      Begin MSMask.MaskEdBox txtDateto 
         Height          =   585
         Left            =   120
         TabIndex        =   29
         Top             =   2160
         Width           =   1395
         _ExtentX        =   2461
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
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   " «  «—ÌŒ :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   1440
         RightToLeft     =   -1  'True
         TabIndex        =   31
         Top             =   2160
         Width           =   1095
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "«“  «—ÌŒ :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   1440
         RightToLeft     =   -1  'True
         TabIndex        =   30
         Top             =   1680
         Width           =   1095
      End
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   360
      Top             =   6000
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
   Begin VB.Frame Frame2 
      Height          =   2775
      Left            =   120
      TabIndex        =   11
      Top             =   600
      Width           =   9015
      Begin VSFlex7LCtl.VSFlexGrid vsPayment 
         Height          =   2415
         Left            =   1800
         TabIndex        =   12
         Top             =   240
         Width           =   7095
         _cx             =   12515
         _cy             =   4260
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
         FormatString    =   $"frmCreditSupplier.frx":A7DC
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
      Begin VB.Label lblRecieved 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "B Mitra"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   885
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   1800
         Width           =   1605
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "„»·€ ò· Å—œ«Œ Ì"
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
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   1320
         Width           =   1605
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "·Ì”  Å—œ«Œ Ì Â«"
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   1005
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   120
         Width           =   1575
      End
   End
   Begin VB.CommandButton CmdCancel 
      BackColor       =   &H00404080&
      Cancel          =   -1  'True
      Caption         =   "Œ—ÊÃ"
      BeginProperty Font 
         Name            =   "B Homa"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   0
      TabIndex        =   10
      Top             =   9000
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Height          =   5085
      Left            =   -240
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   3480
      Width           =   9375
      Begin VB.TextBox txtSelected 
         BeginProperty Font 
            Name            =   "B Mitra"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   600
         TabIndex        =   24
         Top             =   3840
         Width           =   1695
      End
      Begin VB.CommandButton cmdPaySome 
         BackColor       =   &H000000C0&
         Caption         =   "Å—œ«Œ "
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
         Left            =   360
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   4440
         Width           =   1455
      End
      Begin VSFlex7LCtl.VSFlexGrid vsOwedFactors 
         Height          =   2925
         Left            =   240
         TabIndex        =   2
         Top             =   810
         Width           =   9075
         _cx             =   16007
         _cy             =   5159
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
         BackColorFixed  =   12648384
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
         FormatString    =   $"frmCreditSupplier.frx":A856
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
      Begin VB.Label lblBesNo 
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
         Height          =   435
         Left            =   7080
         RightToLeft     =   -1  'True
         TabIndex        =   23
         Top             =   4200
         Width           =   675
      End
      Begin VB.Label LblBestankar 
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
         Height          =   525
         Left            =   4320
         RightToLeft     =   -1  'True
         TabIndex        =   22
         Top             =   4080
         Width           =   1485
      End
      Begin VB.Label Lable1 
         Alignment       =   1  'Right Justify
         Caption         =   " ⁄œ«œ ‰ﬁœÌ "
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
         Height          =   525
         Index           =   2
         Left            =   8040
         RightToLeft     =   -1  'True
         TabIndex        =   21
         Top             =   4200
         Width           =   1170
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         Caption         =   "„»·€ ‰ﬁœÌ "
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
         Height          =   525
         Left            =   5880
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Top             =   4200
         Width           =   1125
      End
      Begin VB.Label lblBedeNo 
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
         Left            =   7080
         RightToLeft     =   -1  'True
         TabIndex        =   19
         Top             =   4680
         Width           =   885
      End
      Begin VB.Label lblBedehkar 
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
         Height          =   405
         Left            =   4320
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   4560
         Width           =   1365
      End
      Begin VB.Label Lable1 
         Alignment       =   1  'Right Justify
         Caption         =   " ⁄œ«œ €Ì— ‰ﬁœÌ "
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
         Height          =   315
         Index           =   0
         Left            =   7800
         RightToLeft     =   -1  'True
         TabIndex        =   17
         Top             =   4680
         Width           =   1410
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "„»·€ €Ì— ‰ﬁœÌ"
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
         Height          =   285
         Left            =   5760
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Top             =   4680
         Width           =   1245
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "„»·€ ›«ò Ê—Â«Ì «‰ Œ«» ‘œÂ Ì« „»·€ Å—œ«Œ Ì »Â  «„Ì‰ ﬂ‰‰œÂ"
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   1125
         Left            =   2325
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   3840
         Visible         =   0   'False
         Width           =   1890
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "„»·€ ò· "
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
         Height          =   525
         Left            =   6120
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   3720
         Width           =   885
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   525
         Index           =   2
         Left            =   1080
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   240
         Width           =   8175
      End
      Begin VB.Label Lable1 
         Alignment       =   1  'Right Justify
         Caption         =   " ⁄œ«œ ò· ›«ò Ê—Â«"
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
         Height          =   525
         Index           =   1
         Left            =   7680
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   3720
         Width           =   1530
      End
      Begin VB.Label lblNoOfFactors 
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
         Height          =   435
         Left            =   7080
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   3720
         Width           =   645
      End
      Begin VB.Label lblSumPrice 
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
         Height          =   525
         Left            =   4320
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   3720
         Width           =   1125
      End
      Begin VB.Label lblMessage 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   1920
         TabIndex        =   3
         Top             =   4440
         Width           =   3135
      End
   End
   Begin VB.Timer Timer1 
      Left            =   120
      Top             =   630
   End
   Begin VSFlex7LCtl.VSFlexGrid vsFactorDetail 
      Height          =   2295
      Left            =   1080
      TabIndex        =   36
      Top             =   8640
      Width           =   8055
      _cx             =   14208
      _cy             =   4048
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
      BackColorFixed  =   12648384
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
      Cols            =   6
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
   Begin RESIZEKITLibCtl.ResizeKit ResizeKit1 
      Height          =   480
      Left            =   0
      OleObjectBlob   =   "frmCreditSupplier.frx":A935
      TabIndex        =   40
      Top             =   0
      Width           =   480
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
      Left            =   1920
      RightToLeft     =   -1  'True
      TabIndex        =   39
      Top             =   0
      Width           =   1095
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "«ﬁ·«„ ›«ò Ê—"
      BeginProperty Font 
         Name            =   "Nazanin"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   9240
      RightToLeft     =   -1  'True
      TabIndex        =   37
      Top             =   9360
      Width           =   1215
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ê—œ‘ Õ”«»  «„Ì‰ ﬂ‰‰œê«‰"
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
      Left            =   4080
      RightToLeft     =   -1  'True
      TabIndex        =   35
      Top             =   0
      Width           =   4095
   End
End
Attribute VB_Name = "frmCreditSupplier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim clsDate As New clsDate
Dim Incharge As EnumIncharge
Dim i As Integer
Dim Parameter() As Parameter
Dim Rst As New ADODB.Recordset

Public Sub ExitForm()
    Unload Me

End Sub
Private Sub CalculateSelected()
    
    Dim tempPrice As Double
    
    With vsOwedFactors
        txtSelected.Text = ""
        If .Rows < 2 Then
            Exit Sub
        End If
        For i = 1 To .Rows - 1
            If Val(.TextMatrix(i, 1)) <> 0 Then
                tempPrice = tempPrice + Val(.TextMatrix(i, 6))
            End If
        Next i
        txtSelected.Text = tempPrice
    End With
End Sub
Private Sub FillBranch()
    Dim L_Rst As New ADODB.Recordset
    cmbBranch.Clear
    cmbBranch.AddItem "Â„Â ‘⁄»Â Â«"
    cmbBranch.ItemData(cmbBranch.NewIndex) = 0
    Set L_Rst = RunStoredProcedure2RecordSet("Get_All_Branches")
    
    Do While L_Rst.EOF = False
        cmbBranch.AddItem L_Rst!nvcBranchName
        cmbBranch.ItemData(cmbBranch.NewIndex) = L_Rst!Branch
        L_Rst.MoveNext
    Loop
    
    L_Rst.Close: Set L_Rst = Nothing
    If cmbBranch.ListCount > 0 Then cmbBranch.ListIndex = 0
End Sub

Public Sub FillvsOwedFactors()

    If fwBtnCustFind.Tag = "" Then Exit Sub
    With vsOwedFactors 'find all the factors which this payk have to pay them
        
        .Rows = 1
        lblNoOfFactors.Caption = 0
        lblSumPrice.Caption = 0
        lblBedeNo.Caption = 0
        lblBedehkar.Caption = 0
        lblBesNo.Caption = 0
        LblBestankar.Caption = 0
        txtSelected.Text = 0
        If Rst.State = 1 Then Rst.Close
        
        ReDim Parameter(4)
        Parameter(0) = GenerateInputParameter("@Owner", adBigInt, 8, fwBtnCustFind.Tag)
        Parameter(1) = GenerateInputParameter("@DateBefore", adVarWChar, 20, CStr(IIf(Trim(txtDateFrom.ClipText) = "", "", Trim(txtDateFrom.Text))))
        Parameter(2) = GenerateInputParameter("@DateAfter", adVarWChar, 20, CStr(IIf(Trim(txtDateTo.ClipText) = "", "", Trim(txtDateTo.Text))))
        Parameter(3) = GenerateInputParameter("@AccountYear", adSmallInt, 2, Val(cmbSalMali.Text))
        Parameter(4) = GenerateInputParameter("@Branch", adInteger, 4, cmbBranch.ItemData(cmbBranch.ListIndex))
        
        Set Rst = RunParametricStoredProcedure2Rec("Get_SupplierFactor", Parameter)
        
        If Not (Rst.EOF = True And Rst.BOF = True) Then
        
            i = 1
            On Error Resume Next
            
            While Rst.EOF = False 'fill the grid
                .Rows = .Rows + 1
                .TextMatrix(i, 0) = Rst.Fields("intSerialNo").Value
                .TextMatrix(i, 2) = Rst.Fields("nvcFirstName").Value & " " & Rst.Fields("nvcSurname").Value
                .TextMatrix(i, 3) = Val(Right(Rst.Fields("No").Value, 3))
                .TextMatrix(i, 4) = Rst.Fields("Code").Value
                .TextMatrix(i, 5) = Rst.Fields("Full Name").Value ' Rst.Fields("Name").Value & " " & Rst.Fields("Family")
                If Rst.Fields("Status").Value = 1 Then
                    .TextMatrix(i, 6) = Rst.Fields("SumPrice").Value
                ElseIf Rst.Fields("Status").Value = 4 Then
                    .TextMatrix(i, 6) = Rst.Fields("SumPrice").Value & "-"
                End If
                .TextMatrix(i, 7) = Rst.Fields("Time").Value
                .TextMatrix(i, 8) = Rst.Fields("Date").Value
                If Rst.Fields("Balance").Value = False Then 'And Rst.Fields("Facpayment").Value = True Then
                   .TextMatrix(i, 9) = 1
                    lblBedeNo.Caption = Val(lblBedeNo.Caption) + 1
                    If Rst.Fields("Status").Value = 1 Then
                        lblBedehkar.Caption = Val(lblBedehkar.Caption) + Rst.Fields("SumPrice").Value
                    ElseIf Rst.Fields("Status").Value = 4 Then
                        lblBedehkar.Caption = Val(lblBedehkar.Caption) - Rst.Fields("SumPrice").Value
                    End If
                Else
                   .TextMatrix(i, 9) = 0
                    lblBesNo.Caption = Val(lblBesNo.Caption) + 1
                    If Rst.Fields("Status").Value = 1 Then
                        LblBestankar.Caption = Val(LblBestankar.Caption) + Rst.Fields("SumPrice").Value
                    ElseIf Rst.Fields("Status").Value = 4 Then
                        LblBestankar.Caption = Val(LblBestankar.Caption) + Rst.Fields("SumPrice").Value
                    End If
                End If
                .TextMatrix(i, 10) = Rst.Fields("Branch").Value
                lblNoOfFactors.Caption = Val(lblNoOfFactors.Caption) + 1
                If Rst.Fields("Status").Value = 1 Then
                    lblSumPrice.Caption = Val(lblSumPrice.Caption) + Rst.Fields("SumPrice").Value
                ElseIf Rst.Fields("Status").Value = 4 Then
                    lblSumPrice.Caption = Val(lblSumPrice.Caption) - Rst.Fields("SumPrice").Value
                End If
                
                Rst.MoveNext
                i = i + 1
            Wend
            
        End If
        .AutoSizeMode = flexAutoSizeColWidth ' set the columns' width
        .AutoSize 0, .Cols - 1
        
    End With
        
End Sub

Private Sub chkRemain_Click()
    cmdUpdate_Click
End Sub

Private Sub cmbBranch_Click()
    
    cmdUpdate_Click

End Sub

Private Sub cmdCancel_Click()
     
    If Rst.State = 1 Then Rst.Close
    Set Rst = Nothing
    Unload Me
End Sub

Private Sub cmdTurnOver_Click()
    If ClsFormAccess.AccfrmKartHesabReport = True Then
        If Val(Tafsili) > 0 Then
            Accounting.KartHesabShowDll "KolBestankaran", CStr(Tafsili), fwBtnCustFind.Caption, txtDateFrom.Text, txtDateTo.Text
        Else
            ShowDisMessage "«Ì‰ „‘ —Ì œ— ”Ì” „ Õ”«»œ«—Ì œ«—«Ì ﬂœ  ›÷Ì·Ì ‰Ì” ", 2000
        End If
    Else
        ShowDisMessage "‘„« »Â «Ì‰ «„ﬂ«‰ œ” —”Ì ‰œ«—Ìœ", 1500
    End If
    
End Sub

Private Sub cmdUpdate_Click()
    
    If Val(fwBtnCustFind.Tag) > 0 Then
        Label1(2).Caption = "·Ì”  ›«ò Ê—Â«Ì ’«œ—‘œÂ »Â ‰«„ " & fwBtnCustFind.Caption
        Label4.Caption = "·Ì”  Å—œ«Œ Ì Â« »Â " & fwBtnCustFind.Caption
    Else
        Label1(2).Caption = "·Ì”  ›«ò Ê—Â«Ì ’«œ—‘œÂ  "
        Label4.Caption = "·Ì”  Å—œ«Œ Ì Â«  "
    End If
    FillvsOwedFactors
    FillvsOwedPaid
    RemainingCalculate

End Sub
Private Sub RemainingCalculate()
    Dim PreRemain As Long
    PreRemain = 0
   
    ReDim Parameter(7) As Parameter
    Parameter(0) = GenerateInputParameter("@SystemDate", adVarWChar, 50, Mid(clsDate.shamsi(Date), 3, 8))
    Parameter(1) = GenerateInputParameter("@SystemDay", adVarWChar, 50, clsDate.Find_DayOfWeek(Weekday(Date, vbSaturday)))
    Parameter(2) = GenerateInputParameter("@SystemTime", adVarWChar, 50, Mid(Str(time), 1, 5))
    Parameter(3) = GenerateInputParameter("@DateBefore", adVarWChar, 50, CStr(IIf(Trim(txtDateFrom.ClipText) = "", "", Trim(txtDateFrom.Text))))
    Parameter(4) = GenerateInputParameter("@DateAfter", adVarWChar, 50, CStr(IIf(Trim(txtDateTo.ClipText) = "", "", Trim(txtDateTo.Text))))
    
    Parameter(5) = GenerateInputParameter("@FromSupplierCode", adInteger, 4, Val(fwBtnCustFind.Tag))
    Parameter(6) = GenerateInputParameter("@ToSupplierCode", adInteger, 4, Val(fwBtnCustFind.Tag))
    Parameter(7) = GenerateInputParameter("@AccountYear", adSmallInt, 2, Val(cmbSalMali.Text))

    Set Rst = RunParametricStoredProcedure2Rec("Get_SupplierBillPayment_Remain", Parameter)
    If Not (Rst.EOF = True And Rst.BOF = True) Then
           '
            PreRemain = Rst.Fields("BeforeRemain").Value
        
    End If
    If PreRemain < 0 Then
         lblPreRemaining.ForeColor = vbRed
         lblPreRemaining.Caption = " »œÂÌ  ﬁ»·Ì:     " & Abs(PreRemain)
    ElseIf PreRemain = 0 Then
        lblPreRemaining.Caption = " ÿ·» ﬁ»·Ì ‰œ«—œ  "
    Else
        fwScrollTextCust.Visible = False
        lblPreRemaining.ForeColor = 0
        lblPreRemaining.BackColor = Me.BackColor
        lblPreRemaining.Caption = " ÿ·»  ﬁ»·Ì :     " & Abs(PreRemain)
    End If
    
    If Val(lblRecieved.Caption) - Val(lblBedehkar.Caption) + PreRemain > 0 Then  '> 0
        fwScrollTextCust.Visible = True
        fwScrollTextCust.ForeColor = vbRed
        fwScrollTextCust.Caption = fwScrollTextCust.Caption & "  -  »œÂÌ œ«—œ : œﬁ  ‘Êœ"
        lblRemaining.ForeColor = vbRed
        lblRemaining.Caption = " »œÂÌ :     " & Abs(Val(lblRecieved.Caption) - Val(lblBedehkar.Caption) + PreRemain) ' + PreRemain
    Else
        fwScrollTextCust.Visible = False
        lblRemaining.ForeColor = 0
        lblRemaining.BackColor = Me.BackColor
        lblRemaining.Caption = " ÿ·»ﬂ«— :     " & Abs(Val(lblBedehkar.Caption) - Val(lblRecieved.Caption) + PreRemain) ' + PreRemain
    End If

End Sub

Private Sub FillvsOwedPaid()

  If fwBtnCustFind.Tag = "" Then Exit Sub
  With vsPayment
        .Rows = 1
        lblRecieved.Caption = 0
        
        ReDim Parameter(5)
        Parameter(0) = GenerateInputParameter("@PaymentType", adInteger, 4, EnumPaymentType.SupplierPayment)
        Parameter(1) = GenerateInputParameter("@Code", adBigInt, 8, fwBtnCustFind.Tag)
        Parameter(2) = GenerateInputParameter("@DateBefore", adVarWChar, 20, CStr(IIf(Trim(txtDateFrom.ClipText) = "", "", Trim(txtDateFrom.Text))))
        Parameter(3) = GenerateInputParameter("@DateAfter", adVarWChar, 20, CStr(IIf(Trim(txtDateTo.ClipText) = "", "", Trim(txtDateTo.Text))))
        Parameter(4) = GenerateInputParameter("@AccountYear", adSmallInt, 2, Val(cmbSalMali.Text))
        Parameter(5) = GenerateInputParameter("@Branch", adInteger, 4, cmbBranch.ItemData(cmbBranch.ListIndex))
        
        Set Rst = RunParametricStoredProcedure2Rec("Get_Paid", Parameter)
        
        If Not (Rst.EOF = True And Rst.BOF = True) Then
        
            i = 1
            
            While Rst.EOF = False 'fill the grid
                .Rows = .Rows + 1
                .TextMatrix(i, 0) = i
                .TextMatrix(i, 1) = Rst.Fields("No").Value
                .TextMatrix(i, 2) = Rst.Fields("Date").Value
                .TextMatrix(i, 3) = Rst.Fields("Regtime").Value
                .TextMatrix(i, 4) = Rst.Fields("User_Name").Value
                .TextMatrix(i, 5) = Rst.Fields("Description").Value
                .TextMatrix(i, 6) = Rst.Fields("Bestankar").Value
                .TextMatrix(i, 7) = Rst.Fields("Branch").Value
                
                lblRecieved.Caption = Val(lblRecieved.Caption) + Rst.Fields("Bestankar").Value
                
                Rst.MoveNext
                i = i + 1
            Wend
            
        End If
        ReDim Parameter(5)
        Parameter(0) = GenerateInputParameter("@RecieveType", adInteger, 4, EnumRecieveType.SupplierRecieve)
        Parameter(1) = GenerateInputParameter("@Code", adBigInt, 8, fwBtnCustFind.Tag)
        Parameter(2) = GenerateInputParameter("@DateBefore", adVarWChar, 20, CStr(IIf(Trim(txtDateFrom.ClipText) = "", "", Trim(txtDateFrom.Text))))
        Parameter(3) = GenerateInputParameter("@DateAfter", adVarWChar, 20, CStr(IIf(Trim(txtDateTo.ClipText) = "", "", Trim(txtDateTo.Text))))
        Parameter(4) = GenerateInputParameter("@AccountYear", adSmallInt, 2, Val(cmbSalMali.Text))
        Parameter(5) = GenerateInputParameter("@Branch", adInteger, 4, cmbBranch.ItemData(cmbBranch.ListIndex))
        
        Set Rst = RunParametricStoredProcedure2Rec("Get_Recieved", Parameter)
        
        If Not (Rst.EOF = True And Rst.BOF = True) Then
        
            While Rst.EOF = False 'fill the grid
                .Rows = .Rows + 1
                .TextMatrix(i, 0) = i
                .TextMatrix(i, 1) = Rst.Fields("No").Value
                .TextMatrix(i, 2) = Rst.Fields("Date").Value
                .TextMatrix(i, 3) = Rst.Fields("Regtime").Value
                .TextMatrix(i, 4) = Rst.Fields("User_Name").Value
                .TextMatrix(i, 5) = Rst.Fields("Description").Value
                .TextMatrix(i, 6) = "(" & Rst.Fields("Bestankar").Value & ")"
                .TextMatrix(i, 7) = Rst.Fields("Branch").Value
                
                lblRecieved.Caption = Val(lblRecieved.Caption) + (-1 * Rst.Fields("Bestankar").Value)
                
                Rst.MoveNext
                i = i + 1
            Wend
            
        End If
        .AutoSizeMode = flexAutoSizeColWidth ' set the columns' width
        .AutoSize 0, .Cols - 1
    
    End With

End Sub
Private Sub Form_Activate()

    Dim i As Integer
    AllButton vbOff, True
    mdifrm.Toolbar1.Buttons(15).Enabled = True  'printing
    mdifrm.Toolbar1.Buttons(23).Enabled = True
    mdifrm.Toolbar1.Buttons(24).Enabled = True
    mdifrm.Toolbar1.Buttons(25).Enabled = True
    mdifrm.Toolbar1.Buttons(26).Enabled = True
    mdifrm.Toolbar1.Buttons(27).Enabled = True
    
''''    FWLed1.Value = CInt(AccountYear)
''''    FWLed1.BackColor = Me.BackColor
''''    FWLed1.ColorOff = Me.BackColor
    FillSalMali
    
    VarActForm = Me.Name
    txtDateFrom.Text = Mid(clsDate.shamsi(Date), 3, 2) & "/01" & "/01"
    txtDateTo.Text = Mid(clsDate.shamsi(Date), 3)
    txtSelected.Visible = False
    Label7.Visible = False
    
End Sub
Private Sub cmbSalMali_Change()
    cmbSalMali_Click
End Sub

Private Sub cmbSalMali_Click()
    txtDateFrom.Text = Right(cmbSalMali.Text, 2) & "/01" & "/01"
    If AccountYear = cmbSalMali.Text Then
        txtDateTo.Text = Mid(clsDate.shamsi(Date), 3)
    Else
        If clsArya.MiladiDate = 0 Then
            txtDateTo.Text = Right(cmbSalMali.Text, 2) & "/12" & "/29"
        Else
            txtDateTo.Text = Right(cmbSalMali.Text, 2) & "/12" & "/31"
        End If
    End If
    FillvsOwedFactors
    FillvsOwedPaid
    RemainingCalculate
End Sub

Private Sub Form_Unload(Cancel As Integer)

    If PosConnection.State = adStateOpen Then PosConnection.Close: Set PosConnection = Nothing
    
    SaveSetting strMainKey, Me.Name, "Left", Me.Left
    SaveSetting strMainKey, Me.Name, "Top", Me.Top

     VarActForm = ""
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call PresetScreenSaver
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call PresetScreenSaver
End Sub

Private Sub fwBtnCustFind_Click()
    Me.FindCust
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
    txtSelected.Visible = True
    Label7.Visible = True
    cmdUpdate_Click
   
End Sub
Private Sub UpdatelblSupplier()

    If fwBtnCustFind.Tag <> "" Then
        Dim Rst As New ADODB.Recordset
        Dim mvarMemberShipId, mvarTel, mvarAddress, mvarDescription As String
        
        fwScrollTextCust.Caption = ""
        fwStatusBarCust.Caption = ""
        lblRemaining.Caption = ""
        lblRemaining.BackColor = Me.BackColor
        lblPreRemaining.Caption = ""
        lblPreRemaining.BackColor = Me.BackColor
        
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
            Tafsili = IIf(IsNull(Rst!Tafsili), 0, Rst!Tafsili)
            If Tafsili > 0 Then cmdTurnOver.Enabled = True Else cmdTurnOver.Enabled = False
            If Rst.Fields("Code") <> -1 Then
                fwScrollTextCust.Caption = mvarDescription
                fwStatusBarCust.Caption = mvarMemberShipId & mvarTel & mvarAddress
                fwStatusBarCust.Visible = True
                lblRemaining.Visible = True
                lblPreRemaining.Visible = True
            Else
                fwScrollTextCust.Caption = ""
                fwStatusBarCust.Caption = ""
            End If
            
            
        End If
        
        Set Rst = Nothing
    End If
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
Private Sub cmdPaySome_Click()

    Dim i As Integer
    Dim s As String
    Dim strPayk As String
    
        s = ""
        With vsOwedFactors
            For i = 1 To .Rows - 1
                If Val(.TextMatrix(i, 1)) = -1 Then
                    s = s & .TextMatrix(i, 0) & ","
                End If
            Next i
            If Val(txtSelected.Text) = 0 Then Exit Sub
            If s = "" Then
                frmMsg.fwlblMsg.Caption = "¬Ì« „«Ì·Ìœ ÊÃÂ Ê«—œ ‘œÂ —« »Â Õ”«»  «„Ì‰ ﬂ‰‰œÂ „‰ŸÊ— ‰„«ÌÌœ ø"
            Else
                frmMsg.fwlblMsg.Caption = "¬Ì« „«Ì·Ìœ ÊÃÂ ›«ò Ê—Â«Ì «‰ Œ«» ‘œÂ —« Å—œ«Œ  ‰„«ÌÌœ ø"
            End If
            frmMsg.fwBtn(0).ButtonType = flwButtonOk
            frmMsg.fwBtn(1).ButtonType = flwButtonCancel
            
            frmMsg.fwBtn(0).Caption = "»·Ì"
            frmMsg.fwBtn(1).Caption = "ŒÌ—"
            
            frmMsg.Show vbModal
            
            If modgl.mvarMsgIdx = vbNo Then
                Exit Sub
            End If

            
            If s = "" Then s = ","
            s = Left(s, Len(s) - 1)
            ReDim Parameter(5) As Parameter
            Parameter(0) = GenerateInputParameter("@strSelectedFactors", adVarWChar, 4000, s)
            Parameter(1) = GenerateInputParameter("@UID", adInteger, 4, mvarCurUserNo)
            Parameter(2) = GenerateInputParameter("@Owner", adBigInt, 8, fwBtnCustFind.Tag)
            Parameter(3) = GenerateInputParameter("@SumPrice", adBigInt, 8, Val(txtSelected.Text))
            Parameter(4) = GenerateInputParameter("@AccountYear", adSmallInt, 2, AccountYear)
            Parameter(5) = GenerateInputParameter("@Branch", adInteger, 4, CurrentBranch)
            RunParametricStoredProcedure "PayFactors_SupplierCredit_Account", Parameter
                
            If InStr(1, s, ",") > 0 Then
                 lblMessage = "›«ò Ê—Â«Ì ‘„«—Â" & s & " Å—œ«Œ  ‘œ "
            Else
                 lblMessage = "›«ò Ê— ‘„«—Â" & s & " Å—œ«Œ  ‘œ "
            End If
            
            cmdUpdate_Click
            
            Timer1.Interval = 3000
            Timer1.Enabled = True
                  
                
        End With
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
    
    If ClsFormAccess.frmCreditSupplier = False Then
        Unload Me
        Exit Sub
    End If
    
    If intVersion = Min Or intVersion = Normal Then
        ShowDisMessage "ê—œ‘ Õ”«»  «„Ì‰ ﬂ‰‰œê«‰ œ— ‰”ŒÂ ÅÌ‘—› Â Ê »«·« — ÊÃÊœ œ«—œ", 1500
        Unload Me
        Exit Sub
    End If
    CenterTop Me
    
    VarActForm = Me.Name
    FillBranch
    With vsOwedFactors
        .Rows = 1
        .Cols = 11
        .ColAlignment(-1) = flexAlignRightCenter
        .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter
'        .ColWidth(1) = 500
        .ColDataType(1) = flexDTBoolean
        .ColDataType(9) = flexDTBoolean
        .ColHidden(0) = True
        
        'set the headers of the columns
        .TextMatrix(0, 0) = "òœ ›Ì‘"
        .TextMatrix(0, 1) = "«‰ Œ«»"
        .TextMatrix(0, 2) = "ÅÌò"
        .TextMatrix(0, 3) = "”—Ì«·"
        .TextMatrix(0, 4) = "òœ"
        .TextMatrix(0, 5) = " «„Ì‰ ﬂ‰‰œÂ"
        .TextMatrix(0, 6) = "„»·€"
        .TextMatrix(0, 7) = "”«⁄ "
        .TextMatrix(0, 8) = " «—ÌŒ"
        .TextMatrix(0, 9) = "€Ì—‰ﬁœÌ"
        .TextMatrix(0, 10) = "‘⁄»Â"
        .ColFormat(6) = "###,###"
        .AutoSearch = flexSearchFromCursor
        .ColHidden(2) = True
        Set Rst = RunStoredProcedure2RecordSet("Get_All_Branches")
        .ColComboList(10) = .BuildComboList(Rst, "nvcBranchName", "Branch")
    
    End With
    
    With vsPayment
        .Rows = 1
        .Cols = 8
        .ColAlignment(-1) = flexAlignRightCenter
        .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter
        
        'set the headers of the columns
        .TextMatrix(0, 0) = "—œÌ›"
        .TextMatrix(0, 1) = "”—Ì«·"
        .TextMatrix(0, 2) = " «—ÌŒ"
        .TextMatrix(0, 3) = "”«⁄ "
        .TextMatrix(0, 4) = "Å—œ«Œ  ﬂ‰‰œÂ"
        .TextMatrix(0, 5) = "‘—Õ"
        .TextMatrix(0, 6) = "„»·€ Å—œ«Œ Ì"
        .TextMatrix(0, 7) = "‘⁄»Â"
    
        .AutoSearch = flexSearchFromCursor
    
        Set Rst = RunStoredProcedure2RecordSet("Get_All_Branches")
        .ColComboList(7) = .BuildComboList(Rst, "nvcBranchName", "Branch")
    End With
        
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

Private Sub vsOwedFactors_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode <> 32 Then Exit Sub ' if the key is not space bar
    Dim i As Integer
    Dim S2 As String
    
    With vsOwedFactors
        If .Row > 0 And .Rows > 1 And .TextMatrix(.Row, 9) <> 0 Then
        
            .Select .Row, 1
            .EditCell
            
        End If
    End With

    CalculateSelected
End Sub
Public Sub FillvsFactorDetail() ' fills the detail of the current factor
    Dim i As Integer
    Dim intselFactor As Double
    
    Dim Rst As New ADODB.Recordset
    
    If Rst.State = 1 Then Rst.Close
    
    With vsFactorDetail
        If vsOwedFactors.Rows <= 1 Then Exit Sub ' if there is no factor in the grid
        ' if at least there is one , choose the current one
        intselFactor = vsOwedFactors.TextMatrix(vsOwedFactors.Row, 0)
        
        ReDim Parameter(1) As Parameter
        Parameter(0) = GenerateInputParameter("@intSelFactor", adInteger, 4, intselFactor)
        Parameter(1) = GenerateInputParameter("@Branch", adInteger, 4, vsOwedFactors.ValueMatrix(vsOwedFactors.Row, 10))
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
        .AutoSizeMode = flexAutoSizeColWidth  ' set the collumns' width
        .AutoSize 0, .Cols - 1
    End With
    
    If Rst.State <> 0 Then Rst.Close
    Set Rst = Nothing
    

End Sub

Private Sub vsOwedFactors_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    FillvsFactorDetail
    
    Dim i As Integer
    Dim S2 As String
    
    With vsOwedFactors
        If .Col = 1 And .Row > 0 And .Rows > 1 And Button = 1 And .TextMatrix(.Row, 9) <> 0 Then
            
            .Select .Row, .Col
            .EditCell
            
        End If
    End With

    CalculateSelected
End Sub

Private Sub vsOwedFactors_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
        If Button = 2 And vsOwedFactors.MouseRow = vsOwedFactors.Row Then
            Me.PopupMenu PaykContextMenu
        
        End If


End Sub

Public Sub Printing()
With vsOwedFactors
    
    RunNonParametricStoredProcedure "Delete_tblPrint_CreditCustomer"
    
    ReDim Parameter(6) As Parameter
    For i = 1 To .Rows - 1
            Parameter(0) = GenerateInputParameter("@Date", adVarChar, 50, .TextMatrix(i, 8))
            Parameter(1) = GenerateInputParameter("@Time", adVarChar, 50, .TextMatrix(i, 7))
            Parameter(2) = GenerateInputParameter("@SerialNo", adDouble, 8, Val(.TextMatrix(i, 3)))
            Parameter(3) = GenerateInputParameter("@Bedehkar", adBigInt, 8, 0)
            Parameter(4) = GenerateInputParameter("@Bestankar", adBigInt, 8, Val(.TextMatrix(i, 6)))
            Parameter(5) = GenerateInputParameter("@CustCode", adInteger, 4, fwBtnCustFind.Tag)
            Parameter(6) = GenerateInputParameter("@ResiveType", adTinyInt, 1, 2)
            RunParametricStoredProcedure "Insert_tblPrint_CreditCustomer", Parameter
    Next i
End With
With vsPayment
    ReDim Parameter(4) As Parameter
    For i = 1 To .Rows - 1
            Parameter(0) = GenerateInputParameter("@Date", adVarChar, 50, .TextMatrix(i, 2))
            Parameter(1) = GenerateInputParameter("@Time", adVarChar, 50, .TextMatrix(i, 3))
            Parameter(2) = GenerateInputParameter("@SerialNo", adDouble, 8, Val(.TextMatrix(i, 1)))
            Parameter(3) = GenerateInputParameter("@Bedehkar", adBigInt, 8, Val(.TextMatrix(i, 6)))
            Parameter(4) = GenerateInputParameter("@Bestankar", adBigInt, 8, 0)
            
            RunParametricStoredProcedure "Insert_tblPrint_CreditCustomer", Parameter
    Next i
    
End With
    ReDim Parameter(5) As Parameter
    Parameter(0) = GenerateInputParameter("@SystemDate", adVarWChar, 50, Mid(clsDate.shamsi(Date), 3, 8))
    Parameter(1) = GenerateInputParameter("@SystemDay", adVarWChar, 50, clsDate.Find_DayOfWeek(Weekday(Date, vbSaturday)))
    Parameter(2) = GenerateInputParameter("@SystemTime", adVarWChar, 50, Mid(Str(time), 1, 5))
    Parameter(3) = GenerateInputParameter("@DateBefore", adVarWChar, 20, CStr(IIf(Trim(txtDateFrom.ClipText) = "", "", Trim(txtDateFrom.Text))))
    Parameter(4) = GenerateInputParameter("@DateAfter", adVarWChar, 20, CStr(IIf(Trim(txtDateTo.ClipText) = "", "", Trim(txtDateTo.Text))))
    Parameter(5) = GenerateInputParameter("@CustCode", adInteger, 4, fwBtnCustFind.Tag)
    
  '  CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\RepCreditCust.rpt"
    CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\RepCreditCust_A4.rpt"
    
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
    CrystalReport1.ReportTitle = fwBtnCustFind.Caption
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



