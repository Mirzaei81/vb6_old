VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{7AEDC602-D94C-11D1-BB7A-00E0290EA3C9}#1.0#0"; "ResizeKit.ocx"
Object = "{75D4F767-8785-11D3-93AD-0000832EF44D}#2.14#0"; "FAST2001.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmCreditCustomerAccount 
   ClientHeight    =   10935
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12855
   Icon            =   "frmCreditCustomerAccount.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   10935
   ScaleWidth      =   12855
   Begin VB.CommandButton cmdPayCheque 
      BackColor       =   &H000000C0&
      Caption         =   "œ—Ì«›  çﬂ"
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
      Left            =   240
      RightToLeft     =   -1  'True
      TabIndex        =   48
      Top             =   9240
      Width           =   1335
   End
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
      TabIndex        =   37
      Top             =   120
      Width           =   1695
   End
   Begin VB.Frame Frame4 
      Height          =   3975
      Left            =   10200
      TabIndex        =   30
      Top             =   5160
      Width           =   2535
      Begin FLWCtrls.FWScrollText fwScrollTextCust 
         Height          =   555
         Left            =   120
         TabIndex        =   31
         Top             =   240
         Width           =   2235
         _ExtentX        =   3942
         _ExtentY        =   979
         Caption         =   ""
         BorderStyle     =   9
         FontName        =   "B Homa"
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
         TabIndex        =   50
         Top             =   2880
         Width           =   2205
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
         TabIndex        =   49
         Top             =   840
         Width           =   2205
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
         TabIndex        =   32
         Top             =   3360
         Width           =   2205
      End
   End
   Begin VB.Frame Frame3 
      Height          =   4935
      Left            =   10200
      TabIndex        =   23
      Top             =   120
      Width           =   2535
      Begin VB.CommandButton cmdTurnOver 
         Caption         =   "ê—œ‘ Õ”«» «Ì‰ „‘ —Ì"
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
         TabIndex        =   54
         Top             =   4200
         Width           =   2295
      End
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
         TabIndex        =   51
         Text            =   "Combo1"
         Top             =   240
         Width           =   2355
      End
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   41
         Top             =   1680
         Width           =   1335
      End
      Begin VB.CommandButton cmdUpdate 
         BackColor       =   &H008080FF&
         Caption         =   "»Â —Ê“ —”«‰Ì «ÿ·«⁄«  „‘ —Ì"
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
         RightToLeft     =   -1  'True
         TabIndex        =   24
         Top             =   3600
         Width           =   2295
      End
      Begin FLWCtrls.FWCoolButton fwBtnCustFind 
         Height          =   810
         Left            =   120
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   840
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
         Picture         =   "frmCreditCustomerAccount.frx":A4C2
         PictureAlign    =   4
         Caption         =   "„‘ —Ì"
         MaskColor       =   -2147483633
      End
      Begin MSMask.MaskEdBox txtDatefrom 
         Height          =   585
         Left            =   120
         TabIndex        =   26
         Top             =   2280
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
         TabIndex        =   27
         Top             =   2880
         Width           =   1455
         _ExtentX        =   2566
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
      Begin VB.Label Label13 
         Alignment       =   2  'Center
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
         Left            =   1560
         RightToLeft     =   -1  'True
         TabIndex        =   42
         Top             =   1800
         Width           =   855
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   " «  «—ÌŒ :"
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
         Left            =   1560
         RightToLeft     =   -1  'True
         TabIndex        =   29
         Top             =   3000
         Width           =   855
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "«“  «—ÌŒ :"
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
         Left            =   1560
         RightToLeft     =   -1  'True
         TabIndex        =   28
         Top             =   2400
         Width           =   855
      End
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   0
      Top             =   6840
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
   Begin VB.Frame Frame_NoAcc 
      Height          =   3015
      Left            =   120
      TabIndex        =   11
      Top             =   600
      Width           =   9975
      Begin VB.OptionButton OptionPaid 
         Alignment       =   1  'Right Justify
         Caption         =   "œ—Ì«›  ‰ﬁœ"
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   12
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   375
         Index           =   0
         Left            =   8280
         RightToLeft     =   -1  'True
         TabIndex        =   53
         Top             =   120
         Value           =   -1  'True
         Width           =   1455
      End
      Begin VB.OptionButton OptionPaid 
         Alignment       =   1  'Right Justify
         Caption         =   "œ—Ì«›  çﬂ"
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   12
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   375
         Index           =   1
         Left            =   5760
         RightToLeft     =   -1  'True
         TabIndex        =   52
         Top             =   120
         Width           =   1455
      End
      Begin VSFlex7LCtl.VSFlexGrid vsRecieved 
         Height          =   2295
         Left            =   1800
         TabIndex        =   12
         Top             =   600
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
         FormatString    =   $"frmCreditCustomerAccount.frx":A7DC
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
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "·Ì”  œ—Ì«› Ì Â«"
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
         TabIndex        =   36
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label lblRecieved 
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
         Height          =   645
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   1680
         Width           =   1485
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "„»·€ ò· œ—Ì«› Ì"
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
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   1200
         Width           =   1365
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
      Left            =   240
      TabIndex        =   10
      Top             =   10200
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Height          =   5565
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   3600
      Width           =   9975
      Begin VB.CommandButton cmdRecursive 
         BackColor       =   &H000000FF&
         Caption         =   "»—ê‘  ›Ì‘ Â«"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1005
         Left            =   2280
         RightToLeft     =   -1  'True
         TabIndex        =   47
         Top             =   4440
         Width           =   1095
      End
      Begin VB.TextBox txtSelected 
         Enabled         =   0   'False
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
         Left            =   120
         TabIndex        =   46
         Top             =   3840
         Width           =   1455
      End
      Begin VB.CommandButton cmdPaySome 
         BackColor       =   &H000000C0&
         Caption         =   "œ—Ì«›  ÊÃÂ ›Ì‘ Â«"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1005
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   45
         Top             =   4440
         Width           =   1095
      End
      Begin VB.CheckBox ChkSale 
         Alignment       =   1  'Right Justify
         Caption         =   "Â„Â ›—Ê‘ Â«"
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   12
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   345
         Index           =   0
         Left            =   2760
         RightToLeft     =   -1  'True
         TabIndex        =   40
         Top             =   240
         Value           =   1  'Checked
         Width           =   1785
      End
      Begin VB.CheckBox ChkSale 
         Alignment       =   1  'Right Justify
         Caption         =   "›ﬁÿ ›—Ê‘ Â«Ì »œÂò«—"
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   12
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   345
         Index           =   1
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   39
         Top             =   240
         Width           =   2625
      End
      Begin VB.CommandButton cmdBalance 
         BackColor       =   &H000000FF&
         Caption         =   " ”ÊÌÂ  ›Ì‘ Â« œ— Â„«‰ —Ê“"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1005
         Left            =   1200
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   4440
         Width           =   1095
      End
      Begin VSFlex7LCtl.VSFlexGrid vsOwedFactors 
         Height          =   2925
         Left            =   180
         TabIndex        =   2
         Top             =   720
         Width           =   9675
         _cx             =   17066
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
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmCreditCustomerAccount.frx":A856
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
      Begin VB.Label LblSelected 
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
         Height          =   495
         Left            =   1560
         TabIndex        =   43
         Top             =   3960
         Width           =   1935
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
         Left            =   7800
         RightToLeft     =   -1  'True
         TabIndex        =   22
         Top             =   4200
         Width           =   525
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
         Height          =   405
         Left            =   5520
         RightToLeft     =   -1  'True
         TabIndex        =   21
         Top             =   4200
         Width           =   1365
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
         Left            =   8640
         RightToLeft     =   -1  'True
         TabIndex        =   20
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
         Left            =   6600
         RightToLeft     =   -1  'True
         TabIndex        =   19
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
         Height          =   405
         Left            =   7800
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   4680
         Width           =   645
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
         Height          =   525
         Left            =   5160
         RightToLeft     =   -1  'True
         TabIndex        =   17
         Top             =   4560
         Width           =   1395
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
         Height          =   405
         Index           =   0
         Left            =   8520
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Top             =   4680
         Width           =   1290
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
         Height          =   405
         Left            =   6480
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   4680
         Width           =   1245
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "„»·€ ›«ò Ê—Â«Ì «‰ Œ«» ‘œÂ Ì« „»·€ œ—Ì«› Ì «“ „‘ —Ì"
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
         Left            =   3480
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   3840
         Visible         =   0   'False
         Width           =   1650
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
         Left            =   6960
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   3720
         Width           =   765
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   525
         Index           =   2
         Left            =   4680
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   240
         Width           =   5175
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
         Left            =   8280
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
         Left            =   7800
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   3720
         Width           =   525
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
         Height          =   405
         Left            =   5400
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   3720
         Width           =   1365
      End
      Begin VB.Label lblMessage 
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
         Height          =   525
         Left            =   1920
         TabIndex        =   3
         Top             =   4560
         Width           =   3135
      End
   End
   Begin VB.Timer Timer1 
      Left            =   120
      Top             =   630
   End
   Begin VSFlex7LCtl.VSFlexGrid vsFactorDetail 
      Height          =   1695
      Left            =   2880
      TabIndex        =   34
      Top             =   9240
      Width           =   7200
      _cx             =   12700
      _cy             =   2990
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
      OleObjectBlob   =   "frmCreditCustomerAccount.frx":A935
      TabIndex        =   44
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
      TabIndex        =   38
      Top             =   240
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
      Left            =   10200
      RightToLeft     =   -1  'True
      TabIndex        =   35
      Top             =   9360
      Width           =   1215
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ê—œ‘ Õ”«» „‘ —Ì«‰"
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
      Left            =   4680
      RightToLeft     =   -1  'True
      TabIndex        =   33
      Top             =   0
      Width           =   3375
   End
End
Attribute VB_Name = "frmCreditCustomerAccount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim clsDate As New clsDate
Dim Incharge As EnumIncharge
Dim i As Integer
Dim Parameter() As Parameter
Dim Rst As New ADODB.Recordset
Public CustomerPaymentType As EnumCustomerPaymentType
Dim ResiveCash, PaymentCash As Long
Dim ResiveCheque As Long
Dim PreRemain, Currentsale, Curentremain As Currency

'================================
Public Sub ExitForm()
    Unload Me
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

Private Sub CalculateSelected()
    Dim tempPrice As Double
    
    With vsOwedFactors
        txtSelected.Text = ""
        lblSelected.Caption = ""
        If .Rows < 2 Then
            Exit Sub
        End If
        For i = 1 To .Rows - 1
            If Val(.TextMatrix(i, 1)) <> 0 Then
                tempPrice = tempPrice + Val(.TextMatrix(i, 6))
            End If
        Next i
        lblSelected.Caption = tempPrice
        txtSelected.Text = tempPrice
    End With
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
        Parameter(1) = GenerateInputParameter("@Branch", adInteger, 4, vsOwedFactors.ValueMatrix(vsOwedFactors.Row, 11))
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

Private Sub ChkSale_Click(index As Integer)
    If index = 0 Then
    
        If ChkSale(0).Value = 1 Then
            ChkSale(1).Value = 0
        Else
            ChkSale(1).Value = 1
        End If
    Else
        If ChkSale(1).Value = 1 Then
            ChkSale(0).Value = 0
        Else
            ChkSale(0).Value = 1
        End If
    End If

  FillvsOwedFactors

End Sub

Private Sub cmbBranch_Click()
    cmdUpdate_Click
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
    FillvsOwedRecieved
    RemainingCalculate
End Sub
Public Sub FillvsOwedFactors()
    On Error GoTo Err_Handler
    
    If fwBtnCustFind.Tag = "" Or fwBtnCustFind.Tag = "-1" Then Exit Sub
    With vsOwedFactors 'find all the factors which this payk have to pay them
        .Rows = 1
        lblNoOfFactors.Caption = 0
        lblSumPrice.Caption = 0
        lblBedeNo.Caption = 0
        lblBedehkar.Caption = 0
        lblBesNo.Caption = 0
        LblBestankar.Caption = 0
        lblSelected.Caption = 0
        txtSelected.Text = 0
        If Rst.State = 1 Then Rst.Close
       Dim i As Integer
            ReDim Parameter(4)
            Parameter(0) = GenerateInputParameter("@Customer", adBigInt, 8, fwBtnCustFind.Tag)
            Parameter(1) = GenerateInputParameter("@DateBefore", adVarWChar, 20, CStr(IIf(Trim(txtDateFrom.ClipText) = "", "", Trim(txtDateFrom.Text))))
            Parameter(2) = GenerateInputParameter("@DateAfter", adVarWChar, 20, CStr(IIf(Trim(txtDateTo.ClipText) = "", "", Trim(txtDateTo.Text))))
            Parameter(3) = GenerateInputParameter("@AccountYear", adSmallInt, 2, Val(cmbSalMali.Text))
            Parameter(4) = GenerateInputParameter("@Branch", adInteger, 4, cmbBranch.ItemData(cmbBranch.ListIndex))
            'Parameter(4) = GenerateInputParameter("@Branch", adInteger, 4, CurrentBranch)
            Set Rst = RunParametricStoredProcedure2Rec("Get_CustomerFactor", Parameter)
            
            If Not (Rst.EOF = True And Rst.BOF = True) Then
            
                i = 1
                'On Error Resume Next
                
                While Rst.EOF = False 'fill the grid
                    If (ChkSale(0).Value = 1) Or (Rst!Balance = False And Rst!FacPayment = True) Then      'Right(ClsDate.shamsi(Date), 10)
                        .Rows = .Rows + 1
                        .TextMatrix(i, 0) = Rst!intSerialNo
                        .TextMatrix(i, 2) = Rst!nvcFirstName & " " & Rst!nvcSurname
                        .TextMatrix(i, 3) = Val(Rst!No)
                        .TextMatrix(i, 4) = Rst!Code
                        .TextMatrix(i, 5) = Rst![Full Name] ' Rst!Name & " " & Rst!Family
                        
                        If Rst!Status = 2 Then
                            .TextMatrix(i, 6) = Rst!sumPrice
                        ElseIf Rst!Status = 5 Then
                            .TextMatrix(i, 6) = Rst!sumPrice & "-"
                        End If
                        
                        .TextMatrix(i, 7) = Rst!time
                        .TextMatrix(i, 8) = Rst!Date
'                        If Rst.Fields("Balance").Value = False And Rst.Fields("Facpayment").Value = True Then
                        
'                        If Rst!CreditBalance = False Then
                           .TextMatrix(i, 9) = 1
                            lblBedeNo.Caption = Val(lblBedeNo.Caption) + 1
                            If Rst!Status = 2 Then
                                lblBedehkar.Caption = Val(lblBedehkar.Caption) + Rst!sumPrice
                            ElseIf Rst!Status = 5 Then
                                lblBedehkar.Caption = Val(lblBedehkar.Caption) - Rst!sumPrice
                            End If
'                        Else
'                           .TextMatrix(i, 9) = 0
'                            lblBesNo.Caption = Val(lblBesNo.Caption) + 1
'                            If Rst!Status = 2 Then
'                                LblBestankar.Caption = Val(LblBestankar.Caption) + Rst!sumPrice
'                            ElseIf Rst!Status = 5 Then
'                                LblBestankar.Caption = Val(LblBestankar.Caption) - Rst!sumPrice
'                            End If
'                        End If
                        
                        If Rst!Balance = True Then
                        'If Rst!CreditBalance = True Then
                           .TextMatrix(i, 10) = 1
                        End If
                        
                        lblNoOfFactors.Caption = Val(lblNoOfFactors.Caption) + 1
                        
                        If Rst!Status = 2 Then
                            lblSumPrice.Caption = Val(lblSumPrice.Caption) + Rst!sumPrice
                        ElseIf Rst!Status = 5 Then
                            lblSumPrice.Caption = Val(lblSumPrice.Caption) - Rst!sumPrice
                        End If
                        .TextMatrix(i, 11) = Rst!Branch
                        i = i + 1
                    End If
                    Rst.MoveNext
                Wend
                
            End If
            .AutoSizeMode = flexAutoSizeColWidth ' set the columns' width
            .AutoSize 0, .Cols - 1
  End With

Exit Sub
Err_Handler:
    LogSaveNew "frmCreditCustomerAccount => ", err.Description, err.Number, err.Source, "FillVsOwedFactors"
    ShowErrorMessage
End Sub

Private Sub chkRemain_Click()
    cmdUpdate_Click
End Sub

Private Sub cmdCancel_Click()
    If Rst.State <> 0 Then Rst.Close
    Set Rst = Nothing
    Unload Me
End Sub

Private Sub cmdPayCheque_Click()
    CustomerPaymentType = Cheque
    frmCustomerPayment.Show vbModal
    cmdUpdate_Click
End Sub


Private Sub cmdPaySome_Click()
    Dim i As Integer
    Dim S1, S2 As String
    Dim strPayk As String
    
        S1 = ""
        S2 = ""
        With vsOwedFactors
            For i = 1 To .Rows - 1
                If Val(.TextMatrix(i, 1)) = -1 Then
                    S1 = S1 & .TextMatrix(i, 3) & ","
                    S2 = S2 & .TextMatrix(i, 0) & ","
                End If
            Next i
            
            If Val(txtSelected.Text) = 0 Then Exit Sub
            
            If S1 = "" Then
                ShowMessage "¬Ì« „«Ì·Ìœ ÊÃÂ Ê«—œ ‘œÂ —« »Â Õ”«» „‘ —Ì „‰ŸÊ— ‰„«ÌÌœ ø", True, True, "»·Ì", "ŒÌ—"
            Else
                ShowMessage "¬Ì« „«Ì·Ìœ ›«ò Ê—Â«Ì «‰ Œ«» ‘œÂ —«  ”ÊÌÂ ‰„«ÌÌœ ø", True, True, "»·Ì", "ŒÌ—"
            End If
            
            If modgl.mvarMsgIdx = vbNo Then
                Exit Sub
            End If
            
            If S1 = "" Then S1 = ","
            If S2 = "" Then S2 = ","
            S1 = Left(S1, Len(S1) - 1)
            S2 = Left(S2, Len(S2) - 1)
            
            ReDim Parameter(6) As Parameter
            Parameter(0) = GenerateInputParameter("@strSelectedFactors", adVarWChar, 4000, S1)
            Parameter(1) = GenerateInputParameter("@strSelectedIntSerialNos", adVarWChar, 4000, S2)
            Parameter(2) = GenerateInputParameter("@UID", adInteger, 4, mvarCurUserNo)
            Parameter(3) = GenerateInputParameter("@Customer", adBigInt, 8, fwBtnCustFind.Tag)
            Parameter(4) = GenerateInputParameter("@SumPrice", adBigInt, 8, Val(txtSelected.Text))
            Parameter(5) = GenerateInputParameter("@AccountYear", adSmallInt, 2, Val(cmbSalMali.Text))
            Parameter(6) = GenerateInputParameter("@intSerialNo", adBigInt, 4, 0)
            RunParametricStoredProcedure "PayFactors_CustCredit_Account2", Parameter
                
            If InStr(1, S1, ",") > 0 Then
                 ShowMessage "›«ò Ê—Â«Ì ‘„«—Â" & S1 & " Å—œ«Œ  ‘œ ", True, False, " «ÌÌœ", ""
            Else
                 ShowMessage "›«ò Ê— ‘„«—Â" & S1 & " Å—œ«Œ  ‘œ ", True, False, " «ÌÌœ", ""
            End If
            
            cmdUpdate_Click
            
            Timer1.Interval = 3000
            Timer1.Enabled = True
        End With
End Sub

Private Sub cmdRecursive_Click()
 Dim i As Integer
    Dim s As String
    Dim strPayk As String
    
           With vsOwedFactors
            For i = 1 To .Rows - 1
                If Val(.TextMatrix(i, 1)) = -1 Then
                    s = s & .TextMatrix(i, 0) & ","
                End If
            Next i
            If Val(lblSelected.Caption) = 0 Then Exit Sub
            
           
                frmMsg.fwlblMsg.Caption = "¬Ì« „«Ì·Ìœ ›«ò Ê—Â«Ì «‰ Œ«» ‘œÂ —« «“ Õ«·   ”ÊÌÂ Ê «—”«· ‘œÂ Œ«—Ã ‰„«ÌÌœ ø"
          
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
            ReDim Parameter(2) As Parameter
            Parameter(0) = GenerateInputParameter("@strSelectedFactors", adVarWChar, 4000, s)
            Parameter(1) = GenerateInputParameter("@UID", adInteger, 4, mvarCurUserNo)
            Parameter(2) = GenerateInputParameter("@AccountYear", adSmallInt, 2, Val(cmbSalMali.Text))
            RunParametricStoredProcedure "RecursiveFactors_CustCredit", Parameter
                
            
            cmdUpdate_Click
            
            Timer1.Interval = 3000
            Timer1.Enabled = True
                  
                
        End With
End Sub

Private Sub cmdTurnOver_Click()
    If Val(fwBtnCustFind.Tag) <= 0 Then Exit Sub
    If ClsFormAccess.AccfrmKartHesabReport = True Then
        If Val(Tafsili) > 0 Then
            Accounting.KartHesabShowDll "KolBedehkaran", CStr(Tafsili), fwBtnCustFind.Caption, txtDateFrom.Text, txtDateTo.Text
        Else
            ShowDisMessage "«Ì‰ „‘ —Ì œ— ”Ì” „ Õ”«»œ«—Ì œ«—«Ì ﬂœ  ›÷Ì·Ì ‰Ì” ", 2000
        End If
    Else
        ShowDisMessage "‘„« »Â «Ì‰ «„ﬂ«‰ œ” —”Ì ‰œ«—Ìœ", 1500
    End If
End Sub

Private Sub cmdUpdate_Click()
    
    If Len(txtDateFrom.ClipText) < 6 Or Len(txtDateTo.ClipText) < 6 Then ShowDisMessage "›Ê—„   «—ÌŒ ’ÕÌÕ ‰Ì” ", 2000: Exit Sub
    If Val(fwBtnCustFind.Tag) > 0 Then
        Label1(2).Caption = "·Ì”  ›«ò Ê—Â«Ì ’«œ—‘œÂ »Â ‰«„ " & fwBtnCustFind.Caption
        Label4.Caption = "·Ì”  œ—Ì«› Ì Â« «“ " & fwBtnCustFind.Caption
    Else
        Label1(2).Caption = "·Ì”  ›«ò Ê—Â«Ì ’«œ—‘œÂ  "
        Label4.Caption = "·Ì”  œ—Ì«› Ì Â«  "
    End If
    FillvsOwedFactors
    FillvsOwedRecieved
    RemainingCalculate

End Sub


Private Sub FillvsOwedRecieved()

  If fwBtnCustFind.Tag = "" Or fwBtnCustFind.Tag = "-1" Then Exit Sub
   
  If OptionPaid(0).Value = True Then
  
        With vsRecieved
        .Rows = 1
        lblRecieved.Caption = 0
        
        i = 1
        
        ReDim Parameter(5)
        Parameter(0) = GenerateInputParameter("@RecieveType", adInteger, 4, EnumRecieveType.CustomerRecieve)
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
                .TextMatrix(i, 6) = Rst.Fields("Bestankar").Value
                .TextMatrix(i, 7) = Rst.Fields("Branch").Value
                
                lblRecieved.Caption = Val(lblRecieved.Caption) + Rst.Fields("Bestankar").Value
                
                Rst.MoveNext
                i = i + 1
            Wend
            
        End If
        
        ReDim Parameter(4)
        Parameter(0) = GenerateInputParameter("@Code", adBigInt, 8, fwBtnCustFind.Tag)
        Parameter(1) = GenerateInputParameter("@DateBefore", adVarWChar, 8, CStr(IIf(Trim(txtDateFrom.ClipText) = "", "", Trim(txtDateFrom.Text))))
        Parameter(2) = GenerateInputParameter("@DateAfter", adVarWChar, 8, CStr(IIf(Trim(txtDateTo.ClipText) = "", "", Trim(txtDateTo.Text))))
        Parameter(3) = GenerateInputParameter("@AccountYear", adSmallInt, 2, Val(cmbSalMali.Text))
        Parameter(4) = GenerateInputParameter("@Branch", adInteger, 4, cmbBranch.ItemData(cmbBranch.ListIndex))
        
        Set Rst = RunParametricStoredProcedure2Rec("Get_Recieved_tFaccash", Parameter)
        
        If Not (Rst.EOF = True And Rst.BOF = True) Then
        
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
                
                lblRecieved.Caption = Val(lblRecieved.Caption) + (Rst.Fields("Bestankar").Value)
                
                Rst.MoveNext
                i = i + 1
            Wend
        End If
        
        ReDim Parameter(5)
        Parameter(0) = GenerateInputParameter("@PaymentType", adInteger, 4, EnumPaymentType.CustomerPayment)
        Parameter(1) = GenerateInputParameter("@Code", adBigInt, 8, fwBtnCustFind.Tag)
        Parameter(2) = GenerateInputParameter("@DateBefore", adVarWChar, 20, CStr(IIf(Trim(txtDateFrom.ClipText) = "", "", Trim(txtDateFrom.Text))))
        Parameter(3) = GenerateInputParameter("@DateAfter", adVarWChar, 20, CStr(IIf(Trim(txtDateTo.ClipText) = "", "", Trim(txtDateTo.Text))))
        Parameter(4) = GenerateInputParameter("@AccountYear", adSmallInt, 2, Val(cmbSalMali.Text))
        Parameter(5) = GenerateInputParameter("@Branch", adInteger, 4, cmbBranch.ItemData(cmbBranch.ListIndex))
        
        Set Rst = RunParametricStoredProcedure2Rec("Get_Paid", Parameter)
        
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
    ElseIf OptionPaid(1).Value = True Then
       
        With vsRecieved
        .Rows = 1
        lblRecieved.Caption = 0
        
        ReDim Parameter(4)
        Parameter(0) = GenerateInputParameter("@Code", adBigInt, 8, fwBtnCustFind.Tag)
        Parameter(1) = GenerateInputParameter("@DateBefore", adVarWChar, 20, CStr(IIf(Trim(txtDateFrom.ClipText) = "", "", Trim(txtDateFrom.Text))))
        Parameter(2) = GenerateInputParameter("@DateAfter", adVarWChar, 20, CStr(IIf(Trim(txtDateTo.ClipText) = "", "", Trim(txtDateTo.Text))))
        Parameter(3) = GenerateInputParameter("@AccountYear", adSmallInt, 2, Val(cmbSalMali.Text))
        Parameter(4) = GenerateInputParameter("@Branch", adInteger, 4, cmbBranch.ItemData(cmbBranch.ListIndex))
        
        Set Rst = RunParametricStoredProcedure2Rec("Get_Recieved_Cheque", Parameter)
        
        If Not (Rst.EOF = True And Rst.BOF = True) Then
        
            i = 1
            
            While Rst.EOF = False 'fill the grid
                .Rows = .Rows + 1
                .TextMatrix(i, 0) = i
                .TextMatrix(i, 1) = Rst.Fields("No").Value
                .TextMatrix(i, 2) = Rst.Fields("intChequeSerial").Value
                .TextMatrix(i, 3) = Rst.Fields("intChequeAcc").Value
                .TextMatrix(i, 4) = Rst.Fields("ChequeDate").Value
                .TextMatrix(i, 5) = Rst.Fields("nvcBankName").Value
                .TextMatrix(i, 6) = Rst.Fields("nvcBranch").Value
                .TextMatrix(i, 7) = Rst.Fields("RegDate").Value
                .TextMatrix(i, 8) = Rst.Fields("Regtime").Value
                .TextMatrix(i, 9) = Rst.Fields("User_Name").Value
                .TextMatrix(i, 10) = Rst.Fields("Description").Value
                .TextMatrix(i, 11) = Rst.Fields("intChequeAmount").Value
                .TextMatrix(i, 12) = Rst.Fields("Branch").Value
                
                lblRecieved.Caption = Val(lblRecieved.Caption) + Rst.Fields("intChequeAmount").Value
                
                Rst.MoveNext
                i = i + 1
            Wend
            
        End If
        .AutoSizeMode = flexAutoSizeColWidth ' set the columns' width
        .AutoSize 0, .Cols - 1
            
        End With
        
    End If
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
    Label7.Visible = False
    txtSelected.Visible = False
    cmdPayCheque.Visible = False
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    If PosConnection.State = adStateOpen Then PosConnection.Close: Set PosConnection = Nothing
    Unload frmFindCust

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
    If clsArya.Customers = True Then
            frmFindCust.Show vbModal
            
            If mvarcode <> 0 Then
                fwBtnCustFind.Tag = mvarcode
                mvarcode = 0
            Else
                fwBtnCustFind.Tag = -1
            End If
            UpdatelblCustomer
           If fwBtnCustFind.Tag <> -1 And Text1.Text <> "" And Text1.Text <> "-1" Then
                txtSelected.Visible = True
                cmdPayCheque.Visible = True
                Label7.Visible = True
                cmdUpdate_Click
           End If
     Else
                    
        frmDisMsg.lblMessage.Caption = " ‘„« »Â «Ì‰ «„ﬂ«‰ œ” —”Ì ‰œ«—Ìœ "
        frmDisMsg.Timer1.Interval = 1000
        frmDisMsg.Timer1.Enabled = True
        frmDisMsg.Show vbModal
       
     End If
   
End Sub
Private Sub UpdatelblCustomer()

    If Val(fwBtnCustFind.Tag) > 0 Then
        Dim Rst As New ADODB.Recordset
        Dim mvarMemberShipId, mvarTel, mvarAddress, mvarDescription As String
        
        fwScrollTextCust.Caption = ""
        fwStatusBarCust.Caption = ""
        lblRemaining.Caption = ""
        lblPreRemaining.Caption = ""
        lblRemaining.BackColor = Me.BackColor
        lblPreRemaining.BackColor = Me.BackColor
        
        ReDim Parameter(1) As Parameter
        Parameter(0) = GenerateInputParameter("@Code", adInteger, 4, Val(fwBtnCustFind.Tag))
        Parameter(1) = GenerateInputParameter("@Branch", adInteger, 4, CurrentBranch)
        Set Rst = RunParametricStoredProcedure2Rec("Get_vw_Customers", Parameter)
        
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
            Text1.Text = Rst.Fields("Membershipid")
            mvarCustCredit = Rst.Fields("Credit")
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
            
            blnCreditCust = IIf(Rst!Credit > 0, True, False)
            
        End If
        
        Set Rst = Nothing
    End If
End Sub

Private Sub OptionPaid_Click(index As Integer)
    SetvsRecieved
    FillvsOwedRecieved
End Sub

Private Sub ResizeKit1_ExitResize(ByVal XScale As Double, ByVal YScale As Double)
    If formloadFlag = True Then
        SaveSetting strMainKey, Me.Name, "Height", Me.Height
        SaveSetting strMainKey, Me.Name, "Width", Me.Width
    End If

End Sub

Private Sub Text1_Change()
   
''    fwBtnCustFind.Tag = Text1.Text
    vsOwedFactors.Rows = 1
    vsFactorDetail.Rows = 1
    vsRecieved.Rows = 1
    ReDim Parameter(1) As Parameter
    Parameter(0) = GenerateInputParameter("@Membershipid", adBigInt, 8, Val(Text1.Text))
    'Parameter(1) = GenerateInputParameter("@Branch", adInteger, 4, CurrentBranch)
    Parameter(1) = GenerateInputParameter("@Branch", adInteger, 4, cmbBranch.ItemData(cmbBranch.ListIndex))
    Set rctmp = RunParametricStoredProcedure2Rec("Get_vw_Customers_ByMembership", Parameter)
    
    If rctmp.EOF = False And rctmp.BOF = False Then
        If fwBtnCustFind.Tag <> rctmp!Code Then
            fwBtnCustFind.Tag = rctmp!Code
             fwBtnCustFind.Caption = ""
             UpdatelblCustomer
        End If
    Else
        fwBtnCustFind.Tag = ""
        fwBtnCustFind.Caption = "„‘ —Ì"
        fwStatusBarCust.Caption = ""
        UpdatelblCustomer
    End If
  
   If Val(fwBtnCustFind.Tag) > 0 And Text1.Text <> "" And Text1.Text <> "-1" Then
       txtSelected.Visible = True
       cmdPayCheque.Visible = True
   ElseIf Val(fwBtnCustFind.Tag) = 0 And Text1.Text <> "" And Text1.Text = "-1" Then
       txtSelected.Visible = False
       cmdPayCheque.Visible = False
   Else
       txtSelected.Visible = False
       cmdPayCheque.Visible = False
   End If
    cmdUpdate_Click
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 And Shift = 0 Then
        cmdUpdate_Click
    End If
End Sub

Private Sub Timer1_Timer()
    lblMessage.Caption = ""
    Timer1.Enabled = False
End Sub
Private Sub cmdBalance_Click()

    Dim i As Integer
    Dim s As String
    Dim strIntSerialNos As String
    Dim strPayk As String

    s = ""
    strIntSerialNos = ""
    
    With vsOwedFactors
        For i = 1 To .Rows - 1
            If Val(.TextMatrix(i, 1)) = -1 Then
                s = s & .TextMatrix(i, 3) & ","
                strIntSerialNos = strIntSerialNos & .TextMatrix(i, 0) & ","
            End If
        Next i
        If Val(lblSelected.Caption) = 0 Then Exit Sub
        If s = "" Then
            ShowMessage "¬Ì« „«Ì·Ìœ ÊÃÂ Ê«—œ ‘œÂ —« »Â Õ”«» „‘ —Ì „‰ŸÊ— ‰„«ÌÌœ ø", True, True, "»·Ì", "ŒÌ—"
        Else
            ShowMessage "¬Ì« „«Ì·Ìœ ›«ò Ê—Â«Ì «‰ Œ«» ‘œÂ —«  ”ÊÌÂ ‰„«ÌÌœ ø", True, True, "»·Ì", "ŒÌ—"
        End If
        
        If modgl.mvarMsgIdx = vbNo Then
            Exit Sub
        End If
        
        If s = "" Then s = ","
        s = Left(s, Len(s) - 1)
        ReDim Parameter(2) As Parameter
        Parameter(0) = GenerateInputParameter("@strSelectedIntSerialNos", adVarWChar, 4000, strIntSerialNos)
        Parameter(1) = GenerateInputParameter("@UID", adInteger, 4, mvarCurUserNo)
        Parameter(2) = GenerateInputParameter("@AccountYear", adSmallInt, 2, Val(cmbSalMali.Text))
        RunParametricStoredProcedure "PayFactors_CustCredit_Balance", Parameter
            
        If InStr(1, s, ",") > 0 Then
             ShowMessage "›«ò Ê—Â«Ì ‘„«—Â" & s & " Å—œ«Œ  ‘œ ", True, False, " «ÌÌœ", ""
        Else
             ShowMessage "›«ò Ê— ‘„«—Â" & s & " Å—œ«Œ  ‘œ ", True, False, " «ÌÌœ", ""
        End If
        
        cmdUpdate_Click
        
        Timer1.Interval = 3000
        Timer1.Enabled = True
    End With
End Sub

Private Sub Form_Load()
    
    If ClsFormAccess.frmCreditCustomer = False Then
        Unload Me
        Exit Sub
    End If
    
    CenterTop Me
    
    Incharge = Payk
    
''''    Dim obj As Object
''''    For Each obj In Forms
''''        If TypeOf obj Is Form Then
''''            If obj.Name <> "mdifrm" And obj.Name <> Me.Name And (obj.Name <> "frminvoice" Or obj.Name <> "frminvoice_shop") Then
''''                Unload obj
''''            End If
''''        End If
''''
''''    Next obj
    
    txtDateFrom.Text = Mid(clsDate.shamsi(Date), 3, 2) & "/01" & "/01"
    txtDateTo.Text = Mid(clsDate.shamsi(Date), 3)
    FillBranch
    
    SetvsOwedFactor
    SetvsRecieved
'    With vsRecieved
'        .Rows = 1
'        .Cols = 8
'        .ColAlignment(-1) = flexAlignCenterCenter
'        .ColAlignment(4) = flexAlignRightCenter
'        .ColAlignment(5) = flexAlignRightCenter
'        .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter
'
'        'set the headers of the columns
'
'        .TextMatrix(0, 0) = "—œÌ›"
'        .TextMatrix(0, 1) = "”—Ì«·"
'        .TextMatrix(0, 2) = " «—ÌŒ"
'        .TextMatrix(0, 3) = "”«⁄ "
'        .TextMatrix(0, 4) = "œ—Ì«›  ﬂ‰‰œÂ"
'        .TextMatrix(0, 5) = "‘—Õ"
'        .TextMatrix(0, 6) = "„»·€ œ—Ì«› Ì"
'        .TextMatrix(0, 7) = "„»·€ œ—Ì«› Ì"
'        .AutoSearch = flexSearchFromCursor
'
'    End With
        
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

    If strCategory = "04" And strDelegate = "00" And clsArya.CustomerId = 102 Then cmdBalance.Enabled = False
    If clsArya.ExternalAccounting = True Then cmdTurnOver.Visible = True Else cmdTurnOver.Visible = False
End Sub



Private Sub vsOwedFactors_Click()
    
    FillvsFactorDetail

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

Private Sub vsOwedFactors_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
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
    
    If PreRemain <> 0 Then
        ReDim Parameter(6) As Parameter
        Parameter(0) = GenerateInputParameter("@Date", adVarChar, 50, "")
        Parameter(1) = GenerateInputParameter("@Time", adVarChar, 50, "„«‰œÂ «“ ﬁ»·")
        Parameter(2) = GenerateInputParameter("@SerialNo", adDouble, 8, 0)
        If PreRemain > 0 Then
            Parameter(3) = GenerateInputParameter("@Bedehkar", adBigInt, 8, PreRemain)
            Parameter(4) = GenerateInputParameter("@Bestankar", adBigInt, 8, 0)
        Else  ' Status = 5
            Parameter(3) = GenerateInputParameter("@Bedehkar", adBigInt, 8, 0)
            Parameter(4) = GenerateInputParameter("@Bestankar", adBigInt, 8, PreRemain)
        End If
        Parameter(5) = GenerateInputParameter("@CustCode", adInteger, 4, fwBtnCustFind.Tag)
        Parameter(6) = GenerateInputParameter("@ResiveType", adTinyInt, 1, 0)
        RunParametricStoredProcedure "Insert_tblPrint_CreditCustomer", Parameter
    End If
        
    ReDim Parameter(4)
    Parameter(0) = GenerateInputParameter("@Customer", adBigInt, 8, fwBtnCustFind.Tag)
    Parameter(1) = GenerateInputParameter("@DateBefore", adVarWChar, 20, CStr(IIf(Trim(txtDateFrom.ClipText) = "", "", Trim(txtDateFrom.Text))))
    Parameter(2) = GenerateInputParameter("@DateAfter", adVarWChar, 20, CStr(IIf(Trim(txtDateTo.ClipText) = "", "", Trim(txtDateTo.Text))))
    Parameter(3) = GenerateInputParameter("@AccountYear", adSmallInt, 2, Val(cmbSalMali.Text))
    'Parameter(4) = GenerateInputParameter("@Branch", adInteger, 4, CurrentBranch)
    Parameter(4) = GenerateInputParameter("@Branch", adInteger, 4, cmbBranch.ItemData(cmbBranch.ListIndex))
    Set Rst = RunParametricStoredProcedure2Rec("Get_CustomerFactor", Parameter)
    
    If Not (Rst.EOF = True And Rst.BOF = True) Then
                
        ReDim Parameter(6) As Parameter
        While Rst.EOF = False
'            If Rst.Fields("Balance").Value = False And Rst.Fields("Facpayment").Value = True Then      'Right(ClsDate.shamsi(Date), 10)
            Parameter(0) = GenerateInputParameter("@Date", adVarChar, 50, Rst.Fields("Date").Value)
            Parameter(1) = GenerateInputParameter("@Time", adVarChar, 50, Rst.Fields("Time").Value)
            Parameter(2) = GenerateInputParameter("@SerialNo", adDouble, 8, Rst.Fields("No").Value)
            If Rst.Fields("Status").Value = 2 Then
                Parameter(3) = GenerateInputParameter("@Bedehkar", adBigInt, 8, Rst.Fields("SumPrice").Value)
                Parameter(4) = GenerateInputParameter("@Bestankar", adBigInt, 8, 0)
            Else  ' Status = 5
                Parameter(3) = GenerateInputParameter("@Bedehkar", adBigInt, 8, 0)
                Parameter(4) = GenerateInputParameter("@Bestankar", adBigInt, 8, Rst.Fields("SumPrice").Value)
            End If
            Parameter(5) = GenerateInputParameter("@CustCode", adInteger, 4, fwBtnCustFind.Tag)
            Parameter(6) = GenerateInputParameter("@ResiveType", adTinyInt, 1, 0)
            RunParametricStoredProcedure "Insert_tblPrint_CreditCustomer", Parameter
'            End If
          Rst.MoveNext
        Wend
        
    End If
End With
     ReDim Parameter(5)
     Parameter(0) = GenerateInputParameter("@RecieveType", adInteger, 4, EnumRecieveType.CustomerRecieve)
     Parameter(1) = GenerateInputParameter("@Code", adBigInt, 8, fwBtnCustFind.Tag)
     Parameter(2) = GenerateInputParameter("@DateBefore", adVarWChar, 20, CStr(IIf(Trim(txtDateFrom.ClipText) = "", "", Trim(txtDateFrom.Text))))
     Parameter(3) = GenerateInputParameter("@DateAfter", adVarWChar, 20, CStr(IIf(Trim(txtDateTo.ClipText) = "", "", Trim(txtDateTo.Text))))
     Parameter(4) = GenerateInputParameter("@AccountYear", adSmallInt, 2, Val(cmbSalMali.Text))
     Parameter(5) = GenerateInputParameter("@Branch", adInteger, 4, cmbBranch.ItemData(cmbBranch.ListIndex))
     
     Set Rst = RunParametricStoredProcedure2Rec("Get_Recieved", Parameter)
     
    If Not (Rst.EOF = True And Rst.BOF = True) Then
        ReDim Parameter(6) As Parameter
        While Rst.EOF = False
           Parameter(0) = GenerateInputParameter("@Date", adVarChar, 50, Rst.Fields("Date").Value)
           Parameter(1) = GenerateInputParameter("@Time", adVarChar, 50, Rst.Fields("regTime").Value)
           Parameter(2) = GenerateInputParameter("@SerialNo", adDouble, 8, Rst.Fields("No").Value)
           Parameter(3) = GenerateInputParameter("@Bedehkar", adBigInt, 8, 0)
           Parameter(4) = GenerateInputParameter("@Bestankar", adBigInt, 8, Rst.Fields("Bestankar").Value)
           Parameter(5) = GenerateInputParameter("@CustCode", adInteger, 4, fwBtnCustFind.Tag)
           Parameter(6) = GenerateInputParameter("@ResiveType", adTinyInt, 1, 1)
           RunParametricStoredProcedure "Insert_tblPrint_CreditCustomer", Parameter
           Rst.MoveNext
        Wend
    End If
            
    ReDim Parameter(4)
    Parameter(0) = GenerateInputParameter("@Code", adBigInt, 8, fwBtnCustFind.Tag)
    Parameter(1) = GenerateInputParameter("@DateBefore", adVarWChar, 8, CStr(IIf(Trim(txtDateFrom.ClipText) = "", "", Trim(txtDateFrom.Text))))
    Parameter(2) = GenerateInputParameter("@DateAfter", adVarWChar, 8, CStr(IIf(Trim(txtDateTo.ClipText) = "", "", Trim(txtDateTo.Text))))
    Parameter(3) = GenerateInputParameter("@AccountYear", adSmallInt, 2, Val(cmbSalMali.Text))
    Parameter(4) = GenerateInputParameter("@Branch", adInteger, 4, cmbBranch.ItemData(cmbBranch.ListIndex))
    
    Set Rst = RunParametricStoredProcedure2Rec("Get_Recieved_tFaccash", Parameter)
    
    If Not (Rst.EOF = True And Rst.BOF = True) Then
        ReDim Parameter(6) As Parameter
        While Rst.EOF = False 'fill the grid
           Parameter(0) = GenerateInputParameter("@Date", adVarChar, 50, Rst.Fields("Date").Value)
           Parameter(1) = GenerateInputParameter("@Time", adVarChar, 50, Rst.Fields("regTime").Value)
           Parameter(2) = GenerateInputParameter("@SerialNo", adDouble, 8, Rst.Fields("No").Value)
           Parameter(3) = GenerateInputParameter("@Bedehkar", adBigInt, 8, 0)
           Parameter(4) = GenerateInputParameter("@Bestankar", adBigInt, 8, Rst.Fields("Bestankar").Value)
           Parameter(5) = GenerateInputParameter("@CustCode", adInteger, 4, fwBtnCustFind.Tag)
           Parameter(6) = GenerateInputParameter("@ResiveType", adTinyInt, 1, 1)
           RunParametricStoredProcedure "Insert_tblPrint_CreditCustomer", Parameter
           Rst.MoveNext
        Wend
    End If
        
    ReDim Parameter(4)
    Parameter(0) = GenerateInputParameter("@Code", adBigInt, 8, fwBtnCustFind.Tag)
    Parameter(1) = GenerateInputParameter("@DateBefore", adVarWChar, 20, CStr(IIf(Trim(txtDateFrom.ClipText) = "", "", Trim(txtDateFrom.Text))))
    Parameter(2) = GenerateInputParameter("@DateAfter", adVarWChar, 20, CStr(IIf(Trim(txtDateTo.ClipText) = "", "", Trim(txtDateTo.Text))))
    Parameter(3) = GenerateInputParameter("@AccountYear", adSmallInt, 2, Val(cmbSalMali.Text))
    Parameter(4) = GenerateInputParameter("@Branch", adInteger, 4, cmbBranch.ItemData(cmbBranch.ListIndex))
    
    Set Rst = RunParametricStoredProcedure2Rec("Get_Recieved_Cheque", Parameter)
    
    If Not (Rst.EOF = True And Rst.BOF = True) Then
        ReDim Parameter(6) As Parameter
        While Rst.EOF = False
           Parameter(0) = GenerateInputParameter("@Date", adVarChar, 50, Rst.Fields("RegDate").Value)
           Parameter(1) = GenerateInputParameter("@Time", adVarChar, 50, Rst.Fields("RegTime").Value)
           Parameter(2) = GenerateInputParameter("@SerialNo", adDouble, 8, Rst.Fields("No").Value)
           Parameter(3) = GenerateInputParameter("@Bedehkar", adBigInt, 8, 0)
           Parameter(4) = GenerateInputParameter("@Bestankar", adBigInt, 8, Rst.Fields("intChequeAmount").Value)
           Parameter(5) = GenerateInputParameter("@CustCode", adInteger, 4, fwBtnCustFind.Tag)
           Parameter(6) = GenerateInputParameter("@ResiveType", adTinyInt, 1, 2)
           RunParametricStoredProcedure "Insert_tblPrint_CreditCustomer", Parameter
           Rst.MoveNext
        Wend
         
    End If
    ReDim Parameter(5)
    Parameter(0) = GenerateInputParameter("@PaymentType", adInteger, 4, EnumPaymentType.CustomerPayment)
    Parameter(1) = GenerateInputParameter("@Code", adBigInt, 8, fwBtnCustFind.Tag)
    Parameter(2) = GenerateInputParameter("@DateBefore", adVarWChar, 20, CStr(IIf(Trim(txtDateFrom.ClipText) = "", "", Trim(txtDateFrom.Text))))
    Parameter(3) = GenerateInputParameter("@DateAfter", adVarWChar, 20, CStr(IIf(Trim(txtDateTo.ClipText) = "", "", Trim(txtDateTo.Text))))
    Parameter(4) = GenerateInputParameter("@AccountYear", adSmallInt, 2, Val(cmbSalMali.Text))
    Parameter(5) = GenerateInputParameter("@Branch", adInteger, 4, cmbBranch.ItemData(cmbBranch.ListIndex))
    
    Set Rst = RunParametricStoredProcedure2Rec("Get_Paid", Parameter)
    
    If Not (Rst.EOF = True And Rst.BOF = True) Then
        ReDim Parameter(6) As Parameter
        While Rst.EOF = False 'fill the grid
           Parameter(0) = GenerateInputParameter("@Date", adVarChar, 50, Rst.Fields("RegDate").Value)
           Parameter(1) = GenerateInputParameter("@Time", adVarChar, 50, Rst.Fields("RegTime").Value)
           Parameter(2) = GenerateInputParameter("@SerialNo", adDouble, 8, Rst.Fields("No").Value)
           Parameter(3) = GenerateInputParameter("@Bedehkar", adBigInt, 8, Rst.Fields("Bestankar").Value)
           Parameter(4) = GenerateInputParameter("@Bestankar", adBigInt, 8, 0)
           Parameter(5) = GenerateInputParameter("@CustCode", adInteger, 4, fwBtnCustFind.Tag)
           Parameter(6) = GenerateInputParameter("@ResiveType", adTinyInt, 1, 1)
           RunParametricStoredProcedure "Insert_tblPrint_CreditCustomer", Parameter
           Rst.MoveNext
        Wend
    End If
    
    
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
Private Sub SetvsOwedFactor()
With vsOwedFactors
      .Rows = 1
      .Cols = 12
   '   .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter
'        .ColWidth(1) = 500
      .ColDataType(1) = flexDTBoolean
      .ColDataType(9) = flexDTBoolean
      .ColDataType(10) = flexDTBoolean
      .ColHidden(0) = True
      
      'set the headers of the columns
      .TextMatrix(0, 0) = "òœ ›Ì‘"
      .TextMatrix(0, 1) = "«‰ Œ«»"
      .TextMatrix(0, 2) = "ÅÌò"
      .TextMatrix(0, 3) = "”—Ì«·"
      .TextMatrix(0, 4) = "òœ"
      .TextMatrix(0, 5) = "„‘ —Ì"
      .TextMatrix(0, 6) = "„»·€"
      .TextMatrix(0, 7) = "”«⁄ "
      .TextMatrix(0, 8) = " «—ÌŒ"
      .TextMatrix(0, 9) = "€Ì—‰ﬁœÌ"
      .TextMatrix(0, 10) = " ”ÊÌÂ"
      .TextMatrix(0, 11) = "‘⁄»Â"
      .ColFormat(6) = "###,###"
      
      .ColAlignment(-1) = flexAlignCenterCenter
      .ColAlignment(2) = flexAlignRightCenter
      .ColAlignment(5) = flexAlignRightCenter
     
      .AutoSearch = flexSearchFromCursor
  
      .ColHidden(9) = True
        Set Rst = RunStoredProcedure2RecordSet("Get_All_Branches")
        .ColComboList(11) = .BuildComboList(Rst, "nvcBranchName", "Branch")
  End With
End Sub

Private Sub SetvsRecieved()

If OptionPaid(0).Value = True Then
        With vsRecieved
               .Rows = 1
               .Cols = 8
               
               'set the headers of the columns
           
               .TextMatrix(0, 0) = "—œÌ›"
               .TextMatrix(0, 1) = "”—Ì«·"
               .TextMatrix(0, 2) = " «—ÌŒ"
               .TextMatrix(0, 3) = "”«⁄ "
               .TextMatrix(0, 4) = "ﬂ«—»—"
               .TextMatrix(0, 5) = "‘—Õ"
               .TextMatrix(0, 6) = "„»·€ œ—Ì«› Ì"
               .TextMatrix(0, 7) = "‘⁄»Â"
               
               .AutoSearch = flexSearchFromCursor
               .ColAlignment(-1) = flexAlignRightCenter
             '  .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter
               .ColAlignment(4) = flexAlignRightCenter
               .ColAlignment(5) = flexAlignRightCenter
                 Set Rst = RunStoredProcedure2RecordSet("Get_All_Branches")
                .ColComboList(7) = .BuildComboList(Rst, "nvcBranchName", "Branch")
         
           End With
ElseIf OptionPaid(1).Value = True Then
        With vsRecieved
               .Rows = 1
               .Cols = 13
               .ColAlignment(-1) = flexAlignRightCenter
               .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter
               
               'set the headers of the columns
           
               .TextMatrix(0, 0) = "—œÌ›"
               .TextMatrix(0, 1) = "”—Ì«·"
               .TextMatrix(0, 2) = "”—Ì«· çﬂ"
               .TextMatrix(0, 3) = "‘„«—Â Õ”«»"
               .TextMatrix(0, 4) = " «—ÌŒ ”— —”Ìœ "
               .TextMatrix(0, 5) = "»«‰ﬂ"
               .TextMatrix(0, 6) = "‘⁄»Â"
               .TextMatrix(0, 7) = " «—ÌŒ"
               .TextMatrix(0, 8) = "”«⁄ "
               .TextMatrix(0, 9) = "ﬂ«—»—"
               .TextMatrix(0, 10) = "‘—Õ"
               .TextMatrix(0, 11) = "„»·€ œ—Ì«› Ì"
               .TextMatrix(0, 12) = "‘⁄»Â"
               .AutoSearch = flexSearchFromCursor
           
                 Set Rst = RunStoredProcedure2RecordSet("Get_All_Branches")
                .ColComboList(12) = .BuildComboList(Rst, "nvcBranchName", "Branch")
           End With
End If
End Sub
Private Sub RemainingCalculate()
        If Val(fwBtnCustFind.Tag) <= 0 Then Exit Sub
        ResiveCash = 0
        PaymentCash = 0
        ResiveCheque = 0
        
        ReDim Parameter(5)
        Parameter(0) = GenerateInputParameter("@RecieveType", adInteger, 4, EnumRecieveType.CustomerRecieve)
        Parameter(1) = GenerateInputParameter("@Code", adBigInt, 8, fwBtnCustFind.Tag)
        Parameter(2) = GenerateInputParameter("@DateBefore", adVarWChar, 20, CStr(IIf(Trim(txtDateFrom.ClipText) = "", "", Trim(txtDateFrom.Text))))
        Parameter(3) = GenerateInputParameter("@DateAfter", adVarWChar, 20, CStr(IIf(Trim(txtDateTo.ClipText) = "", "", Trim(txtDateTo.Text))))
        Parameter(4) = GenerateInputParameter("@AccountYear", adSmallInt, 2, Val(cmbSalMali.Text))
        Parameter(5) = GenerateInputParameter("@Branch", adInteger, 4, cmbBranch.ItemData(cmbBranch.ListIndex))
        
        Set Rst = RunParametricStoredProcedure2Rec("Get_Recieved", Parameter)
         If Not (Rst.EOF = True And Rst.BOF = True) Then
                 While Rst.EOF = False
                       ResiveCash = ResiveCash + Rst.Fields("Bestankar").Value
                        
                        Rst.MoveNext
                 Wend
          End If
          
         ReDim Parameter(5)
        Parameter(0) = GenerateInputParameter("@PaymentType", adInteger, 4, EnumPaymentType.CustomerPayment)
        Parameter(1) = GenerateInputParameter("@Code", adBigInt, 8, fwBtnCustFind.Tag)
        Parameter(2) = GenerateInputParameter("@DateBefore", adVarWChar, 20, CStr(IIf(Trim(txtDateFrom.ClipText) = "", "", Trim(txtDateFrom.Text))))
        Parameter(3) = GenerateInputParameter("@DateAfter", adVarWChar, 20, CStr(IIf(Trim(txtDateTo.ClipText) = "", "", Trim(txtDateTo.Text))))
        Parameter(4) = GenerateInputParameter("@AccountYear", adSmallInt, 2, Val(cmbSalMali.Text))
        Parameter(5) = GenerateInputParameter("@Branch", adInteger, 4, cmbBranch.ItemData(cmbBranch.ListIndex))
        
        Set Rst = RunParametricStoredProcedure2Rec("Get_Paid", Parameter)
        
        If Not (Rst.EOF = True And Rst.BOF = True) Then
        
            While Rst.EOF = False 'fill the grid
                 While Rst.EOF = False
                       PaymentCash = PaymentCash + Rst.Fields("Bestankar").Value
                        
                        Rst.MoveNext
                 Wend
            Wend
        End If
        ReDim Parameter(4)
        Parameter(0) = GenerateInputParameter("@Code", adBigInt, 8, fwBtnCustFind.Tag)
        Parameter(1) = GenerateInputParameter("@DateBefore", adVarWChar, 20, CStr(IIf(Trim(txtDateFrom.ClipText) = "", "", Trim(txtDateFrom.Text))))
        Parameter(2) = GenerateInputParameter("@DateAfter", adVarWChar, 20, CStr(IIf(Trim(txtDateTo.ClipText) = "", "", Trim(txtDateTo.Text))))
        Parameter(3) = GenerateInputParameter("@AccountYear", adSmallInt, 2, Val(cmbSalMali.Text))
        Parameter(4) = GenerateInputParameter("@Branch", adInteger, 4, cmbBranch.ItemData(cmbBranch.ListIndex))
        
        Set Rst = RunParametricStoredProcedure2Rec("Get_Recieved_Cheque", Parameter)
        
        If Not (Rst.EOF = True And Rst.BOF = True) Then
                While Rst.EOF = False
                ResiveCheque = ResiveCheque + Rst.Fields("intChequeAmount").Value
                
                Rst.MoveNext
               
            Wend
            
        End If
    
   
    ReDim Parameter(7) As Parameter
    'ArrayUbound = 12
    Parameter(0) = GenerateInputParameter("@SystemDate", adVarWChar, 50, Mid(clsDate.shamsi(Date), 3, 8))
    Parameter(1) = GenerateInputParameter("@SystemDay", adVarWChar, 50, clsDate.Find_DayOfWeek(Weekday(Date, vbSaturday)))
    Parameter(2) = GenerateInputParameter("@SystemTime", adVarWChar, 50, Mid(Str(time), 1, 5))
    Parameter(3) = GenerateInputParameter("@DateBefore", adVarWChar, 50, CStr(IIf(Trim(txtDateFrom.ClipText) = "", "", Trim(txtDateFrom.Text))))
    Parameter(4) = GenerateInputParameter("@DateAfter", adVarWChar, 50, CStr(IIf(Trim(txtDateTo.ClipText) = "", "", Trim(txtDateTo.Text))))
    
    Parameter(5) = GenerateInputParameter("@FromCustCode", adInteger, 4, fwBtnCustFind.Tag)
    Parameter(6) = GenerateInputParameter("@ToCustCode", adInteger, 4, fwBtnCustFind.Tag)
    Parameter(7) = GenerateInputParameter("@AccountYear", adSmallInt, 2, Val(cmbSalMali.Text))
    
    Set Rst = RunParametricStoredProcedure2Rec("Get_CustomerBillPayment_Remain", Parameter)
    PreRemain = 0
    Currentsale = 0
    Curentremain = 0
    If Not (Rst.EOF = True And Rst.BOF = True) Then
           '
         '  While Rst.EOF = False
            PreRemain = Rst.Fields("BeforeRemain").Value
            Currentsale = Rst.Fields("Currentsale").Value
            Curentremain = Rst.Fields("Curentremain").Value
         '   Rst.MoveNext
         '   Wend
        
    End If
    If PreRemain > 0 Then
         lblPreRemaining.ForeColor = vbRed
         lblPreRemaining.Caption = " »œÂÌ  ﬁ»·Ì:     " & Abs(PreRemain)
    ElseIf PreRemain = 0 Then
        lblPreRemaining.Caption = " »œÂÌ ﬁ»·Ì ‰œ«—œ  "
    Else
        fwScrollTextCust.Visible = False
        lblPreRemaining.ForeColor = 0
        lblPreRemaining.BackColor = Me.BackColor
        lblPreRemaining.Caption = " ÿ·»  ﬁ»·Ì :     " & PreRemain
    End If
            
   If (Curentremain + PreRemain) > 0 Then
        fwScrollTextCust.Visible = True
        fwScrollTextCust.ForeColor = vbRed
        fwScrollTextCust.Caption = fwScrollTextCust.Caption & "  -  »œÂÌ œ«—œ : œﬁ  ‘Êœ"
        lblRemaining.ForeColor = vbRed
        lblRemaining.Caption = " »œÂÌ ﬂ·:     " & Abs(Curentremain + PreRemain)
    Else
        fwScrollTextCust.Visible = False
        lblRemaining.ForeColor = 0
        lblRemaining.BackColor = Me.BackColor
        lblRemaining.Caption = " ÿ·» ﬂ· :     " & Abs(Curentremain + PreRemain)
    End If
'   If ResiveCash - PaymentCash + ResiveCheque - Val(lblBedehkar.Caption) + PreRemain < 0 Then
'        fwScrollTextCust.Visible = True
'        fwScrollTextCust.ForeColor = vbRed
'        fwScrollTextCust.Caption = fwScrollTextCust.Caption & "  -  »œÂÌ œ«—œ : œﬁ  ‘Êœ"
'        lblRemaining.ForeColor = vbRed
'        lblRemaining.Caption = " »œÂÌ :     " & Abs(ResiveCash - PaymentCash + ResiveCheque - Val(lblBedehkar.Caption) + PreRemain)
'    Else
'        fwScrollTextCust.Visible = False
'        lblRemaining.ForeColor = 0
'        lblRemaining.BackColor = Me.BackColor
'        lblRemaining.Caption = " ÿ·»ﬂ«— :     " & Abs(ResiveCash - PaymentCash + ResiveCheque - Val(lblBedehkar.Caption) + PreRemain)
'    End If
         
            
End Sub


