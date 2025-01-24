VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{7AEDC602-D94C-11D1-BB7A-00E0290EA3C9}#1.0#0"; "ResizeKit.ocx"
Object = "{75D4F767-8785-11D3-93AD-0000832EF44D}#2.14#0"; "FAST2001.ocx"
Begin VB.Form frmPrinter 
   ClientHeight    =   9675
   ClientLeft      =   300
   ClientTop       =   420
   ClientWidth     =   14475
   Icon            =   "frmPrinter.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   9675
   ScaleWidth      =   14475
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0FFFF&
      Caption         =   " „Ê«—œ ç«Å  ⁄ÌÌ‰ ‘œÂ"
      BeginProperty Font 
         Name            =   "Titr"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5295
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   27
      Top             =   4290
      Width           =   14445
      Begin FLWCtrls.FWButton FWDeleteButton 
         Height          =   480
         Left            =   360
         TabIndex        =   28
         Top             =   4680
         Width           =   1875
         _ExtentX        =   3307
         _ExtentY        =   847
         ButtonType      =   6
         Caption         =   "Õ–› „Ê—œ ç«Å"
         BackColor       =   255
         FontName        =   "B Homa"
         FontBold        =   -1  'True
         FontSize        =   9.75
      End
      Begin VSFlex7LCtl.VSFlexGrid FlxDetail 
         Height          =   4035
         Left            =   -120
         TabIndex        =   29
         Top             =   480
         Width           =   14415
         _cx             =   25426
         _cy             =   7117
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
         BackColor       =   8438015
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483636
         BackColorAlternate=   8438015
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
         SelectionMode   =   3
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   22
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   500
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmPrinter.frx":A4C2
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
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   $"frmPrinter.frx":A71D
         BeginProperty Font 
            Name            =   "Nazanin"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   975
         Left            =   3360
         RightToLeft     =   -1  'True
         TabIndex        =   31
         Top             =   4485
         Width           =   10815
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
      Caption         =   " ⁄—Ì› ç«Åê—Â«"
      BeginProperty Font 
         Name            =   "Titr"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4245
      Left            =   6720
      RightToLeft     =   -1  'True
      TabIndex        =   24
      Top             =   30
      Width           =   7665
      Begin VSFlex7LCtl.VSFlexGrid VSFlexGrid1 
         Height          =   2955
         Left            =   0
         TabIndex        =   25
         Top             =   450
         Width           =   7545
         _cx             =   13309
         _cy             =   5212
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
         BackColor       =   8454143
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483636
         BackColorAlternate=   8454143
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
         Rows            =   1
         Cols            =   4
         FixedRows       =   0
         FixedCols       =   1
         RowHeightMin    =   500
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmPrinter.frx":A7D8
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
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
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
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
      End
      Begin FLWCtrls.FWButton fwBtn 
         Height          =   495
         Left            =   120
         TabIndex        =   26
         Top             =   3600
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   873
         Caption         =   "À» "
         FontName        =   "B Homa"
         FontBold        =   -1  'True
         FontSize        =   9.75
         Alignment       =   1
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H0080C0FF&
      Caption         =   " ⁄ÌÌ‰ „Ê—œ ç«Å  "
      BeginProperty Font 
         Name            =   "Titr"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4275
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   0
      Width           =   6705
      Begin VB.ComboBox cmbPartition 
         BackColor       =   &H0080C0FF&
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
         Left            =   3600
         RightToLeft     =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   32
         Top             =   3720
         Width           =   1875
      End
      Begin VB.TextBox TxtRepeatNo 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0080C0FF&
         Height          =   435
         Left            =   2160
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Text            =   "1"
         Top             =   3120
         Width           =   495
      End
      Begin VB.CheckBox ArmCheck 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0080C0FF&
         Caption         =   "¬—„"
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
         Left            =   2235
         RightToLeft     =   -1  'True
         TabIndex        =   19
         Top             =   660
         Width           =   1080
      End
      Begin VB.CheckBox BarcodeCheck 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0080C0FF&
         Caption         =   "»«—ﬂœ"
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
         Left            =   2235
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   1035
         Width           =   1080
      End
      Begin VB.CheckBox CutterCheck 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0080C0FF&
         Caption         =   "»—‘"
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
         Left            =   2235
         RightToLeft     =   -1  'True
         TabIndex        =   17
         Top             =   1410
         Width           =   1080
      End
      Begin VB.TextBox TxtLineFeed 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0080C0FF&
         Height          =   435
         Left            =   2160
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Top             =   2640
         Width           =   495
      End
      Begin VB.ComboBox cmbState 
         BackColor       =   &H0080C0FF&
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
         Left            =   3600
         RightToLeft     =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   1560
         Width           =   1875
      End
      Begin VB.ComboBox CmbStationNo 
         BackColor       =   &H0080C0FF&
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
         Left            =   3600
         RightToLeft     =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   1020
         Width           =   1875
      End
      Begin VB.ComboBox CmbPrinterNo 
         BackColor       =   &H0080C0FF&
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
         Left            =   3600
         RightToLeft     =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   2100
         Width           =   1875
      End
      Begin VB.ComboBox CmbPrintFormat 
         BackColor       =   &H0080C0FF&
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
         Left            =   3600
         RightToLeft     =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   2640
         Width           =   1875
      End
      Begin VB.CheckBox SerialNoCheck 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0080C0FF&
         Caption         =   "”—Ì«·"
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
         Left            =   2235
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   1785
         Width           =   1080
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H0080C0FF&
         Caption         =   "ç«Å œ— Õ«· "
         BeginProperty Font 
            Name            =   "Titr"
            Size            =   8.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2235
         Left            =   240
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   240
         Width           =   1875
         Begin VB.CheckBox EditCheck 
            Alignment       =   1  'Right Justify
            BackColor       =   &H0080C0FF&
            Caption         =   "ÊÌ—«Ì‘"
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
            Left            =   630
            RightToLeft     =   -1  'True
            TabIndex        =   10
            Top             =   1050
            Value           =   1  'Checked
            Width           =   1080
         End
         Begin VB.CheckBox AddCheck 
            Alignment       =   1  'Right Justify
            BackColor       =   &H0080C0FF&
            Caption         =   "ÃœÌœ"
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
            Left            =   630
            RightToLeft     =   -1  'True
            TabIndex        =   9
            Top             =   675
            Value           =   1  'Checked
            Width           =   1080
         End
         Begin VB.CheckBox ViewCheck 
            Alignment       =   1  'Right Justify
            BackColor       =   &H0080C0FF&
            Caption         =   "„—Ê—"
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
            Left            =   600
            RightToLeft     =   -1  'True
            TabIndex        =   8
            Top             =   300
            Value           =   1  'Checked
            Width           =   1080
         End
         Begin VB.CheckBox ManipulateCheck 
            Alignment       =   1  'Right Justify
            BackColor       =   &H0080C0FF&
            Caption         =   " €ÌÌ—«   ⁄œ«œ ﬂ«·«"
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
            Left            =   30
            RightToLeft     =   -1  'True
            TabIndex        =   7
            Top             =   1425
            Width           =   1680
         End
         Begin VB.CheckBox RefferCheck 
            Alignment       =   1  'Right Justify
            BackColor       =   &H0080C0FF&
            Caption         =   "„—ÃÊ⁄Ì"
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
            Left            =   630
            RightToLeft     =   -1  'True
            TabIndex        =   6
            Top             =   1800
            Width           =   1080
         End
      End
      Begin VB.ComboBox CmbPrintType 
         BackColor       =   &H0080C0FF&
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
         Left            =   3600
         RightToLeft     =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   3180
         Width           =   1875
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H0080C0FF&
         Height          =   1695
         Left            =   240
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   2520
         Width           =   1815
         Begin VB.CheckBox LableCheck 
            Alignment       =   1  'Right Justify
            BackColor       =   &H0080C0FF&
            Caption         =   "Å—Ì‰  ·Ì»· Ê Å—›—«é"
            BeginProperty Font 
               Name            =   "Nazanin"
               Size            =   12
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   240
            RightToLeft     =   -1  'True
            TabIndex        =   33
            Top             =   720
            Width           =   1440
         End
         Begin VB.CheckBox InvoiceCheck 
            Alignment       =   1  'Right Justify
            BackColor       =   &H0080C0FF&
            Caption         =   "›«ﬂ Ê— ›—Ê‘"
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
            Left            =   240
            RightToLeft     =   -1  'True
            TabIndex        =   3
            Top             =   240
            Width           =   1440
         End
      End
      Begin VB.ComboBox CmbStatus 
         BackColor       =   &H0080C0FF&
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
         Left            =   3600
         RightToLeft     =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   480
         Width           =   1875
      End
      Begin FLWCtrls.FWLabel3D FWLblPrintState 
         Height          =   375
         Left            =   5310
         Top             =   1644
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Nazanin"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Deep            =   100
         ForeColor1      =   16761024
         ForeColor2      =   0
         BackColor       =   8438015
         Caption         =   "  :Õ«·  ’œÊ— "
         Alignment       =   1
      End
      Begin FLWCtrls.FWLabel3D FWLabel3D1 
         Height          =   375
         Left            =   5310
         Top             =   2166
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Nazanin"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Deep            =   100
         ForeColor1      =   16761024
         ForeColor2      =   0
         BackColor       =   8438015
         Caption         =   ":«“ Å—Ì‰ — "
         Alignment       =   1
      End
      Begin FLWCtrls.FWLabel3D FWLabel3D3 
         Height          =   375
         Left            =   5310
         Top             =   2688
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Nazanin"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Deep            =   100
         ForeColor1      =   16761024
         ForeColor2      =   0
         BackColor       =   8438015
         Caption         =   "  :Õ«·  ç«Å "
         Alignment       =   1
      End
      Begin FLWCtrls.FWButton FWBtnSave 
         Height          =   495
         Left            =   2160
         TabIndex        =   21
         Top             =   3600
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   873
         Caption         =   "À» "
         FontName        =   "B Homa"
         FontBold        =   -1  'True
         FontSize        =   9.75
         Alignment       =   1
      End
      Begin FLWCtrls.FWLabel3D FWLabel3D5 
         Height          =   375
         Left            =   5520
         Top             =   1122
         Width           =   885
         _ExtentX        =   1561
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Nazanin"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Deep            =   100
         ForeColor1      =   16761024
         ForeColor2      =   0
         BackColor       =   8438015
         Caption         =   "  :«Ì” ê«Â "
         Alignment       =   1
      End
      Begin FLWCtrls.FWLabel3D FWLabel3D2 
         Height          =   375
         Left            =   5310
         Top             =   3210
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Nazanin"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Deep            =   100
         ForeColor1      =   16761024
         ForeColor2      =   0
         BackColor       =   8438015
         Caption         =   "  :‰Ê⁄ ç«Å "
         Alignment       =   1
      End
      Begin FLWCtrls.FWLabel3D FWLabel3D4 
         Height          =   375
         Left            =   5520
         Top             =   600
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Nazanin"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Deep            =   100
         ForeColor1      =   16761024
         ForeColor2      =   0
         BackColor       =   8438015
         Caption         =   "  :Ê÷⁄Ì "
         Alignment       =   1
      End
      Begin FLWCtrls.FWLabel3D FWLabel3D6 
         Height          =   375
         Left            =   5430
         Top             =   3735
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Nazanin"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Deep            =   100
         ForeColor1      =   16761024
         ForeColor2      =   0
         BackColor       =   8438015
         Caption         =   "  :»Œ‘"
         Alignment       =   1
      End
      Begin VB.Label lblRepeatNo 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0080C0FF&
         Caption         =   "  ﬂ—«—      "
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
         Left            =   2640
         RightToLeft     =   -1  'True
         TabIndex        =   23
         Top             =   3120
         Width           =   615
      End
      Begin VB.Label LblLineFeed 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0080C0FF&
         Caption         =   " Œÿ Œ«·Ì    "
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
         Left            =   2640
         RightToLeft     =   -1  'True
         TabIndex        =   22
         Top             =   2640
         Width           =   975
      End
   End
   Begin RESIZEKITLibCtl.ResizeKit ResizeKit1 
      Height          =   480
      Left            =   0
      OleObjectBlob   =   "frmPrinter.frx":A83E
      TabIndex        =   30
      Top             =   0
      Width           =   480
   End
End
Attribute VB_Name = "frmPrinter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private ClsGl As New ClsGl
Private cmd As New ADODB.command
Private Rc As New ADODB.Recordset
Private clsDate As New clsDate
Private rctmp As New ADODB.Recordset
Private MaxRowFlexGrid As Integer
Dim Parameter() As Parameter
Dim MyFormAddEditMode As EnumAddEditMode
Dim i As Integer
Dim intPrintingNo As Integer

Sub AddEmptyRow()

    FlxDetail.Rows = FlxDetail.Rows + 1

End Sub

Sub ClearDataFlexGrid()
   
    FlxDetail.Rows = 1
    FlxDetail.Rows = 8
    MaxRowFlexGrid = 1
    FlxDetail.TopRow = 1
    
End Sub

Sub GetDataDetail()

    ClearDataFlexGrid
    Dim Parameters(1) As Parameter
    
    Parameters(0) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
    Parameters(1) = GenerateInputParameter("@PrintFormat", adInteger, 4, 0)
    
    Set rctmp = RunParametricStoredProcedure2Rec("PrintersInfo", Parameters)
    
    Dim ii As Integer
    If Not (rctmp.BOF Or rctmp.EOF) Then
           With FlxDetail
           Do While Not (rctmp.EOF)
                  
                ii = ii + 1
                FlxDetail.TextMatrix(ii, 0) = rctmp!Number
                FlxDetail.TextMatrix(ii, 1) = Val(rctmp.Fields("StationId").Value)
                FlxDetail.TextMatrix(ii, 2) = rctmp!ServePlaceDescription
               
                FlxDetail.TextMatrix(ii, 3) = Val(rctmp!PrinterNo)
                FlxDetail.TextMatrix(ii, 4) = rctmp!PrintFormatName
                If rctmp!Arm.Value Then
                    .TextMatrix(ii, 5) = -1
                End If
                If rctmp!barcode.Value Then
                    .TextMatrix(ii, 6) = -1
                End If
                If rctmp!Cutter.Value Then
                    .TextMatrix(ii, 7) = -1
                End If
                FlxDetail.TextMatrix(ii, 8) = rctmp!LineFeed.Value
                FlxDetail.TextMatrix(ii, 9) = rctmp!RepeatNo.Value
                If rctmp!SerialNo.Value Then
                    .TextMatrix(ii, 10) = -1
                End If
                If ((rctmp!PermittedModes.Value And ViewMode) = ViewMode) Then
                    .TextMatrix(ii, 11) = -1
                End If
                If ((rctmp!PermittedModes.Value And AddMode) = AddMode) Then
                    .TextMatrix(ii, 12) = -1
                End If
                
                If ((rctmp!PermittedModes.Value And EditMode) = EditMode) Then
                    .TextMatrix(ii, 13) = -1
                End If
                
                If ((rctmp!PermittedModes.Value And ManipulateMode) = ManipulateMode) Then
                    .TextMatrix(ii, 14) = -1
                End If
                
                If ((rctmp!PermittedModes.Value And RefferedMode) = RefferedMode) Then
                    .TextMatrix(ii, 15) = -1
                End If
                If ((rctmp!PermittedModes.Value And InvoiceFactor) = InvoiceFactor) Then
                    .TextMatrix(ii, 16) = -1
                End If
                If ((rctmp!PermittedModes.Value And Perfrage) = Perfrage) Then
                    .TextMatrix(ii, 18) = -1
                End If
                FlxDetail.Row = ii
                FlxDetail.Col = 19
                If rctmp!DirectRpt.Value = 0 Then
                   FlxDetail.TextMatrix(ii, 19) = "ê—«›ÌﬂÌ"
                ElseIf rctmp!DirectRpt.Value = 1 Then
                   FlxDetail.TextMatrix(ii, 19) = "„⁄„Ê·Ì"
                ElseIf rctmp!DirectRpt.Value = 2 Then
                   FlxDetail.TextMatrix(ii, 19) = "”—Ì⁄"
                End If
                
                .Cell(flexcpText, ii, 20) = CStr(rctmp.Fields("Status").Value)
                .Cell(flexcpText, ii, 21) = CStr(rctmp.Fields("PartitionId").Value)
                rctmp.MoveNext
               
               If ii >= FlxDetail.Rows - 1 Then
                  AddEmptyRow
               Else
                  FlxDetail.Row = FlxDetail.Row + 1
               End If
''''               If ii > 50 Then  ' Error
''''                 MsgBox " Error In Get Row From DataBase" & ii
''''                  rctmp.Close
''''                  MaxRowFlexGrid = 1
''''
''''                  Exit Sub
''''                End If
            Loop
            End With
            MaxRowFlexGrid = FlxDetail.Row
    End If
    rctmp.Close
    FlxDetail.Cell(flexcpAlignment, 0, 0, FlxDetail.Rows - 1, FlxDetail.Cols - 1) = flexAlignCenterCenter

End Sub

Private Sub FlxDetail_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    For i = 0 To FlxDetail.Cols - 1
        SaveSetting strMainKey, Me.Name & "_FlxDetail", "Col" & i, FlxDetail.ColWidth(i)
    Next
End Sub

Private Sub FlxDetail_Click()
 If FlxDetail.TextMatrix(FlxDetail.Row, 0) = "" Then Exit Sub
 intPrintingNo = FlxDetail.TextMatrix(FlxDetail.Row, 0)
 mdifrm.Toolbar1.Buttons(7).Enabled = True  'Edit
 GetDataDetailForEdit
'' MyFormAddEditMode = ViewMode
End Sub

Private Sub FlxDetail_KeyDown(KeyCode As Integer, Shift As Integer)
    
    If KeyCode <> 46 Or FlxDetail.TextMatrix(FlxDetail.Row, 0) = "" Or FlxDetail.Col <> 9 Then Exit Sub
    
    If rctmp.State <> 0 Then rctmp.Close
    
    ReDim Parameter(0) As Parameter
    Parameter(0) = GenerateInputParameter("@Number", adInteger, 4, FlxDetail.TextMatrix(FlxDetail.Row, 0))
    RunParametricStoredProcedure "Update_tPrinting_By_Number", Parameter
    GetDataDetail
    
End Sub


Private Sub Form_Activate()

    VarActForm = Me.Name

    GetDataDetail
    
    mdifrm.Toolbar1.Buttons(20).Enabled = False
    mdifrm.Toolbar1.Buttons(21).Enabled = False
    mdifrm.Toolbar1.Buttons(23).Enabled = True
    mdifrm.Toolbar1.Buttons(24).Enabled = True
    mdifrm.Toolbar1.Buttons(25).Enabled = True
    mdifrm.Toolbar1.Buttons(26).Enabled = True
    mdifrm.Toolbar1.Buttons(27).Enabled = True
    
    If CmbPrinterNo.ListCount = 0 Or CmbPrintFormat.ListCount = 0 Or CmbState.ListCount = 0 Or CmbStationNo.ListCount = 0 Then
        FWBtnSave.Enabled = False
    Else
        FWBtnSave.Enabled = True
    End If
    
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

    If ClsFormAccess.frmPrinter = False Then
        Unload Me
        Exit Sub
    End If

    MyFormAddEditMode = AddMode
    CenterTop Me
    
    VSFlexGrid1.Cols = 5
    
    FillVSFlexGrid1
    VSFlexGrid1.ColHidden(3) = True
    
    ReDim Parameter(0) As Parameter
    Parameter(0) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
    Set rctmp = RunParametricStoredProcedure2Rec("Get_StatusType", Parameter)
    
    If Not (rctmp.EOF = True And rctmp.BOF = True) Then
        While Not rctmp.EOF
            CmbStatus.AddItem rctmp!NvcDescription
            CmbStatus.ItemData(CmbStatus.NewIndex) = rctmp!intStatusNo
            rctmp.MoveNext
        Wend
        Me.CmbStatus.ListIndex = 0
    End If
    rctmp.Close
    
    
    ReDim Parameter(0) As Parameter
    Parameter(0) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
    Set rctmp = RunParametricStoredProcedure2Rec("Get_Serveplace_Composite", Parameter)
    
    If Not (rctmp.EOF = True And rctmp.BOF = True) Then
        While Not rctmp.EOF
            CmbState.AddItem rctmp!Description
            CmbState.ItemData(CmbState.NewIndex) = rctmp!intServePlace
            rctmp.MoveNext
        Wend
        Me.CmbState.ListIndex = 0
    End If
    rctmp.Close
    
    If rctmp.State <> 0 Then rctmp.Close
    Set rctmp = RunStoredProcedure2RecordSet("Get_All_Active_PrintFormats")
    If Not (rctmp.EOF = True And rctmp.BOF = True) Then
        rctmp.MoveFirst
        While Not rctmp.EOF
            CmbPrintFormat.AddItem rctmp!PrintFormatName
            CmbPrintFormat.ItemData(CmbPrintFormat.ListCount - 1) = rctmp!PrintFormat
            rctmp.MoveNext
        Wend
        CmbPrintFormat.ListIndex = 0
        
    End If
    rctmp.Close
    
    If rctmp.State <> 0 Then rctmp.Close
    
    ReDim Parameter(2) As Parameter
    Parameter(0) = GenerateInputParameter("@MaxStationNo", adInteger, 4, clsArya.MaxStationNo)
    Parameter(1) = GenerateInputParameter("@MaxPocketPcNo", adInteger, 4, clsArya.MaxPocketPcNo)
    Parameter(2) = GenerateInputParameter("@MaxKitchenNo", adInteger, 4, clsArya.MaxKitchenNo)
    Set rctmp = RunParametricStoredProcedure2Rec("Get_Stations", Parameter)
    
    If Not (rctmp.EOF = True And rctmp.BOF = True) Then
        While Not rctmp.EOF
            CmbStationNo.AddItem rctmp!Description
            CmbStationNo.ItemData(CmbStationNo.ListCount - 1) = rctmp!StationId
            rctmp.MoveNext
        Wend
        CmbStationNo.ListIndex = 0
    End If
'    ReDim Parameter(2) As Parameter
'    Parameter(0) = GenerateInputParameter("@MaxStationNo", adInteger, 4, clsArya.MaxStationNo)
'    Parameter(1) = GenerateInputParameter("@MaxPocketPcNo", adInteger, 4, clsArya.MaxPocketPcNo)
'    Parameter(2) = GenerateInputParameter("@MaxKitchenNo", adInteger, 4, clsArya.MaxKitchenNo)
'    Set rctmp = RunParametricStoredProcedure2Rec("Get_Stations", Parameter)
'    FlxDetail.ColComboList(1) = FlxDetail.BuildComboList(rctmp, "Description", "StationId")
    
    cmbPartition.Clear
    ReDim Parameter(0) As Parameter
    Parameter(0) = GenerateInputParameter("@Branch", adInteger, 4, CurrentBranch)
    Set rctmp = RunParametricStoredProcedure2Rec("Get_tPartitions", Parameter)
    If Not (rctmp.EOF = True And rctmp.BOF = True) Then
        While Not rctmp.EOF
            cmbPartition.AddItem rctmp!PartitionDescription
            cmbPartition.ItemData(cmbPartition.ListCount - 1) = rctmp!PartitionID
            rctmp.MoveNext
        Wend
        cmbPartition.ListIndex = 0
    End If
    
    ReDim Parameter(0) As Parameter
    Parameter(0) = GenerateInputParameter("@Branch", adInteger, 4, CurrentBranch)
    Set rctmp = RunParametricStoredProcedure2Rec("Get_tPartitions", Parameter)
    FlxDetail.ColComboList(21) = FlxDetail.BuildComboList(rctmp, "PartitionDescription", "PartitionID")

    If clsArya.FastPrint = True Then
        CmbPrintType.AddItem " ê—«›ÌﬂÌ "
        CmbPrintType.ItemData(CmbPrintType.ListCount - 1) = 0
        CmbPrintType.AddItem " „⁄„Ê·Ì "
        CmbPrintType.ItemData(CmbPrintType.ListCount - 1) = 1
        CmbPrintType.AddItem " ”—Ì⁄ "
        CmbPrintType.ItemData(CmbPrintType.ListCount - 1) = 2
        CmbPrintType.ListIndex = 0
    Else
        CmbPrintType.AddItem " ê—«›ÌﬂÌ "
        CmbPrintType.ItemData(CmbPrintType.ListCount - 1) = 0
        CmbPrintType.ListIndex = 0
    End If
    If CmbPrinterNo.ListCount = 0 Or CmbPrintFormat.ListCount = 0 Or CmbState.ListCount = 0 Or CmbStationNo.ListCount = 0 Then
        FWBtnSave.Enabled = False
    End If

    ReDim Parameter(0) As Parameter
    Parameter(0) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
    Set rctmp = RunParametricStoredProcedure2Rec("Get_StatusType", Parameter)
   
    FlxDetail.ColComboList(20) = FlxDetail.BuildComboList(rctmp, "NvcDescription", "intStatusNo")
    
    With FlxDetail
        .ColAlignment(0) = flexAlignRightCenter
        .ColAlignment(1) = flexAlignLeftCenter
        .ColAlignment(2) = flexAlignCenterCenter
        .ColAlignment(4) = flexAlignCenterCenter
        .ColHidden(17) = True
        '''.ColHidden(18) = True  ''perfrage
    
         For i = 0 To .Cols - 1
            .ColWidth(i) = Val(GetSetting(strMainKey, Me.Name & "_flxDetail", "Col" & i))
            If .ColWidth(i) = 0 Then
                .ColWidth(i) = 1000       '
            End If
         Next i
        
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

'    If intVersion = Min Then
'        CmbStatus.ListIndex = 1
'        CmbStatus.Enabled = False
'        CmbPrinterNo.Enabled = False
'        CmbStationNo.Enabled = False
'    End If


End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call PresetScreenSaver
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call PresetScreenSaver
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Set rctmp = Nothing
    Set Rc = Nothing
    Set cmd = Nothing
    If PosConnection.State = adStateOpen Then PosConnection.Close: Set PosConnection = Nothing
    
    mdifrm.Toolbar1.Buttons(20).Enabled = False
    mdifrm.Toolbar1.Buttons(21).Enabled = False
 '   mdifrm.Toolbar1.Buttons(27).Enabled = False
    
    VarActForm = ""
    
    mdifrm.Toolbar1.Buttons(20).Enabled = False
    mdifrm.Toolbar1.Buttons(21).Enabled = False
    mdifrm.Toolbar1.Buttons(23).Enabled = True
    mdifrm.Toolbar1.Buttons(24).Enabled = True
    mdifrm.Toolbar1.Buttons(25).Enabled = True
    mdifrm.Toolbar1.Buttons(26).Enabled = True
    mdifrm.Toolbar1.Buttons(27).Enabled = True

    SaveSetting strMainKey, Me.Name, "Left", Me.Left
    SaveSetting strMainKey, Me.Name, "Top", Me.Top


    
End Sub

Private Sub fwBtn_Click()       ' Save Data
    
    Dim i As Integer
    Dim j As Integer
    
    With VSFlexGrid1
        
        For i = 0 To .Rows - 2
            If Trim(.TextMatrix(i, 1)) <> "" Then
                For j = i + 1 To .Rows - 1
                    If LCase(Trim(.TextMatrix(i, 1)) & Trim(.TextMatrix(i, 2))) = LCase(Trim(.TextMatrix(j, 1)) & Trim(.TextMatrix(j, 2))) Then
                        MsgBox "«‰ Œ«» Å—Ì‰ —  ò—«—Ì œ— ·Ì”  Å—Ì‰ —Â« „Ã«“ ‰„Ì »«‘œ"
                        Exit Sub
                    End If
                Next j
            End If
        Next i
        
    End With
    
    On Error GoTo RollBack
    PosConnection.BeginTrans

    RunNonParametricStoredProcedure "DroptPrintersCK"
    For i = 0 To VSFlexGrid1.Rows - 1
    
        If VSFlexGrid1.TextMatrix(i, 1) <> "" Then
        
            If rctmp.State <> 0 Then rctmp.Close
            
            rctmp.CursorType = adOpenDynamic
            rctmp.LockType = adLockOptimistic
            
            ReDim Parameter(0) As Parameter
            Parameter(0) = GenerateInputParameter("@PrinterNo", adInteger, 4, Val(VSFlexGrid1.TextMatrix(i, 3)))
            Set rctmp = RunMPSP2Rec("Get_Printers_By_PrinterNo", Parameter)
            
            If (rctmp.EOF = True And rctmp.BOF = True) Then  'not exists
                
                rctmp.AddNew
                
            End If
            
            rctmp.Fields("PrinterNo").Value = i + 1
            
            If InStr(1, VSFlexGrid1.TextMatrix(i, 1), "\\", vbTextCompare) Then 'Network Printer
                Dim HostTemp As Integer
                
                HostTemp = InStr(3, VSFlexGrid1.TextMatrix(i, 1), "\", vbTextCompare)
                rctmp.Fields("PrinterName").Value = Mid(VSFlexGrid1.TextMatrix(i, 1), HostTemp + 1)
                rctmp.Fields("HostName").Value = "\\" & Mid(VSFlexGrid1.TextMatrix(i, 1), 3, HostTemp - 3) & "\"
            Else
                If VSFlexGrid1.TextMatrix(i, 2) = "" Then
                    rctmp.Fields("PrinterName").Value = VSFlexGrid1.TextMatrix(i, 1)
                     ' PrinterName Not Empty
                    rctmp.Fields("HostName").Value = "\\" & MachineName & "\"   'Only For Local Printers
                End If
            End If
            
'            rctmp.Fields("Port").Value = VSFlexGrid1.TextMatrix(i, 4) ' Local Port
            rctmp.Fields("Branch").Value = CurrentBranch ' Local Branch

            Dim X As Printer
            For Each X In Printers
                If X.DeviceName Like VSFlexGrid1.TextMatrix(i, 1) Then
                    Set Printer = X
                    rctmp.Fields("Port").Value = Mid(Printer.Port, 1, Len(Printer.Port) - 1) ' Delete : From End Of String
                    Exit For
                End If
            Next
            rctmp.Update
        Else
        
            If rctmp.State <> 0 Then rctmp.Close
            
            ReDim Parameter(0) As Parameter
            Parameter(0) = GenerateInputParameter("@PrinterNo", adInteger, 4, Val(VSFlexGrid1.TextMatrix(i, 3)))
            Set rctmp = RunMPSP2Rec("Get_Printers_By_PrinterNo", Parameter)
            
            If Not (rctmp.EOF = True And rctmp.BOF = True) Then ' exists
                
                rctmp.Delete
                
            End If
            
        End If
        
    Next i
    RunNonParametricStoredProcedure "AddtPrintersCK"
        
    PosConnection.CommitTrans
    On Error GoTo 0
    
    FillVSFlexGrid1
    GetDataDetail

    Exit Sub

RollBack:
    PosConnection.RollbackTrans
    Select Case err.Number
        Case -2147217873
            MsgBox "«‰ Œ«» Å—Ì‰ —  ò—«—Ì œ— ·Ì”  Å—Ì‰ —Â« „Ã«“ ‰„Ì »«‘œ"
    End Select
    
End Sub

Public Sub ExitForm()
    Unload Me
End Sub

Private Sub fwBtn_KeyUp(KeyCode As Integer, Shift As Integer)
    KeyActi vbtxtbox, KeyCode, Shift, frmPrinter
End Sub

Private Sub FWDeleteButton_Click()

    Dim s As String
    Dim i As Integer
    
    If FlxDetail.SelectedRows > 0 Then
    
        For i = 0 To FlxDetail.SelectedRows - 1
            If FlxDetail.TextMatrix(FlxDetail.SelectedRow(i), 0) <> "" Then
                s = s & FlxDetail.TextMatrix(FlxDetail.SelectedRow(i), 0) & ","
            End If
        Next i
        
        If s <> "" Then
            s = Left(s, Len(s) - 1)
        Else
            Exit Sub
        End If
        
        frmMsg.fwlblMsg.Caption = "¬Ì« „Ì ŒÊ«ÂÌœ «Ì‰ „Ê«—œ ç«Å —« Õ–› ò‰Ìœ ø"
        
        frmMsg.fwBtn(0).ButtonType = flwButtonOk
        frmMsg.fwBtn(1).ButtonType = flwButtonCancel
        
        frmMsg.fwBtn(0).Caption = "»·Ì"
        frmMsg.fwBtn(1).Caption = "ŒÌ—"
        
        frmMsg.Show vbModal
        
        If modgl.mvarMsgIdx = vbNo Then
            Exit Sub
        End If
        
        ReDim Parameter(0) As Parameter
        
        Parameter(0) = GenerateInputParameter("@Numbers", adVarWChar, 200, s)
        
        RunParametricStoredProcedure "DeletePrinting", Parameter
        GetDataDetail
        
    End If
End Sub

Private Sub FWBtnSave_Click()

    Dim PermittedModes As String
    
    If ViewCheck.Value = 1 Then
        PermittedModes = "1" & PermittedModes
    Else
        PermittedModes = "0" & PermittedModes
    End If
    If AddCheck.Value = 1 Then
        PermittedModes = "1" & PermittedModes
    Else
        PermittedModes = "0" & PermittedModes
    End If
    If EditCheck.Value = 1 Then
        PermittedModes = "1" & PermittedModes
    Else
        PermittedModes = "0" & PermittedModes
    End If
    If ManipulateCheck.Value = 1 Then
        PermittedModes = "1" & PermittedModes
    Else
        PermittedModes = "0" & PermittedModes
    End If
    If RefferCheck.Value = 1 Then
        PermittedModes = "1" & PermittedModes
    Else
        PermittedModes = "0" & PermittedModes
    End If
    If InvoiceCheck.Value = 1 Then
        PermittedModes = "1" & PermittedModes
    Else
        PermittedModes = "0" & PermittedModes
    End If
    If LableCheck.Value = 1 Then
        PermittedModes = "1" & PermittedModes
    Else
        PermittedModes = "0" & PermittedModes
    End If
    
  Select Case MyFormAddEditMode
        Case EnumAddEditMode.AddMode
                ReDim Parameter(15) As Parameter
                
                Parameter(0) = GenerateInputParameter("@StationId", adInteger, 4, Me.CmbStationNo.ItemData(CmbStationNo.ListIndex))
                Parameter(1) = GenerateInputParameter("@ServePlace", adInteger, 4, Me.CmbState.ItemData(Me.CmbState.ListIndex))
                Parameter(2) = GenerateInputParameter("@PrinterNo", adInteger, 4, Me.CmbPrinterNo.ItemData(CmbPrinterNo.ListIndex))
                Parameter(3) = GenerateInputParameter("@PrintFormat", adInteger, 4, Me.CmbPrintFormat.ItemData(CmbPrintFormat.ListIndex))
                Parameter(4) = GenerateInputParameter("@Arm", adBoolean, 1, Val(Me.ArmCheck.Value))
                Parameter(5) = GenerateInputParameter("@Barcode", adBoolean, 1, Val(Me.BarcodeCheck.Value))
                Parameter(6) = GenerateInputParameter("@SerialNo", adBoolean, 1, Val(Me.SerialNoCheck.Value))
                Parameter(7) = GenerateInputParameter("@Cutter", adBoolean, 1, Val(Me.CutterCheck.Value))
                Parameter(8) = GenerateInputParameter("@LineFeed", adInteger, 4, Val(Me.TxtLineFeed.Text))
                Parameter(9) = GenerateInputParameter("@RepeatNo", adInteger, 4, Val(Me.TxtRepeatNo.Text))
                Parameter(10) = GenerateInputParameter("@User", adInteger, 4, mvarCurUserNo)
                Parameter(11) = GenerateInputParameter("@PermittedModes", adInteger, 4, ConvertBinToInt(PermittedModes))
                Parameter(12) = GenerateInputParameter("@DirectRpt", adInteger, 4, Me.CmbPrintType.ItemData(CmbPrintType.ListIndex))
                Parameter(13) = GenerateInputParameter("@Status", adInteger, 4, Me.CmbStatus.ItemData(CmbStatus.ListIndex))
                Parameter(14) = GenerateInputParameter("@partitionId", adInteger, 4, Me.cmbPartition.ItemData(cmbPartition.ListIndex))
                Parameter(15) = GenerateOutputParameter("@Number", adInteger, 4)
                
                On Error Resume Next
                
                If RunParametricStoredProcedure("InserttPrinting", Parameter) = -1 Then
                    MsgBox "À»  «‰Ã«„ ‰‘œ.„Ê—œ ç«Å  ò—«—Ì „Ì »«‘œ", , "Œÿ«"
                    Exit Sub
            
                End If
                On Error GoTo 0
                
                ClearDataFlexGrid
                GetDataDetail
         Case EnumAddEditMode.EditMode
                
                ReDim Parameter(16) As Parameter
                
                Parameter(0) = GenerateInputParameter("@StationId", adInteger, 4, Me.CmbStationNo.ItemData(CmbStationNo.ListIndex))
                Parameter(1) = GenerateInputParameter("@ServePlace", adInteger, 4, Me.CmbState.ItemData(Me.CmbState.ListIndex))
                Parameter(2) = GenerateInputParameter("@PrinterNo", adInteger, 4, Me.CmbPrinterNo.ItemData(CmbPrinterNo.ListIndex))
                Parameter(3) = GenerateInputParameter("@PrintFormat", adInteger, 4, Me.CmbPrintFormat.ItemData(CmbPrintFormat.ListIndex))
                Parameter(4) = GenerateInputParameter("@Arm", adBoolean, 1, Val(Me.ArmCheck.Value))
                Parameter(5) = GenerateInputParameter("@Barcode", adBoolean, 1, Val(Me.BarcodeCheck.Value))
                Parameter(6) = GenerateInputParameter("@SerialNo", adBoolean, 1, Val(Me.SerialNoCheck.Value))
                Parameter(7) = GenerateInputParameter("@Cutter", adBoolean, 1, Val(Me.CutterCheck.Value))
                Parameter(8) = GenerateInputParameter("@LineFeed", adInteger, 4, Val(Me.TxtLineFeed.Text))
                Parameter(9) = GenerateInputParameter("@RepeatNo", adInteger, 4, Val(Me.TxtRepeatNo.Text))
                Parameter(10) = GenerateInputParameter("@User", adInteger, 4, mvarCurUserNo)
                Parameter(11) = GenerateInputParameter("@PermittedModes", adInteger, 4, ConvertBinToInt(PermittedModes))
                Parameter(12) = GenerateInputParameter("@DirectRpt", adInteger, 4, Me.CmbPrintType.ItemData(CmbPrintType.ListIndex))
                Parameter(13) = GenerateInputParameter("@Status", adInteger, 4, Me.CmbStatus.ItemData(CmbStatus.ListIndex))
                Parameter(14) = GenerateInputParameter("@partitionId", adInteger, 4, Me.cmbPartition.ItemData(cmbPartition.ListIndex))
                Parameter(15) = GenerateInputParameter("@Number", adInteger, 4, intPrintingNo)
                Parameter(16) = GenerateOutputParameter("@Update", adInteger, 4)
                
                On Error Resume Next
                
                If RunParametricStoredProcedure("UpdatetPrinting", Parameter) = -1 Then
                    MsgBox "À»  «‰Ã«„ ‰‘œ.„Ê—œ ç«Å  ò—«—Ì „Ì »«‘œ", , "Œÿ«"
                    Exit Sub
            
                End If
                On Error GoTo 0
                
                ClearDataFlexGrid
                GetDataDetail
                MyFormAddEditMode = EnumAddEditMode.AddMode
                mdifrm.Toolbar1.Buttons(9).Enabled = False  'Cancel

         
         End Select

End Sub

Private Sub InvoiceCheck_Click()
    If InvoiceCheck.Value = 1 Then
       ViewCheck.Value = 0
       AddCheck.Value = 0
       ManipulateCheck = 0
       EditCheck = 0
       RefferCheck = 0
       LableCheck = 0
    End If

End Sub

Private Sub LableCheck_Click()
    If LableCheck.Value = 1 Then
       ViewCheck.Value = 0
       AddCheck.Value = 0
       ManipulateCheck = 0
       EditCheck = 0
       RefferCheck = 0
       InvoiceCheck = 0
    End If

End Sub

Private Sub ResizeKit1_ExitResize(ByVal XScale As Double, ByVal YScale As Double)

    If formloadFlag = True Then
        SaveSetting strMainKey, Me.Name, "Height", Me.Height
        SaveSetting strMainKey, Me.Name, "Width", Me.Width
    End If

End Sub

Private Sub VSFlexGrid1_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    VSFlexGrid1.TextMatrix(Row, 2) = ""
    VSFlexGrid1_ValidateEdit VSFlexGrid1.Row, VSFlexGrid1.Col, False
End Sub

Private Sub VSFlexGrid1_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim s As String
    If Col = 4 Then
        
        s = ""
        If rctmp.State <> 0 Then
            rctmp.Close
        End If
        Set rctmp = RunStoredProcedure2RecordSet("Get_All_Ports")
        s = VSFlexGrid1.BuildComboList(rctmp, "PortName", "PortCode")
        VSFlexGrid1.ColComboList(4) = s
    
        Exit Sub
    End If
    If Col <> 1 Then Cancel = True
    
    Dim i As Integer
    Dim DontAdd As Boolean
    Dim X As Printer
    s = ""
    
    With VSFlexGrid1
        For Each X In Printers
            For i = 0 To .Rows - 1
                If i <> .Row Then
                    If LCase(X.DeviceName) = LCase(.TextMatrix(i, 1)) Or LCase(X.DeviceName) = LCase(.TextMatrix(i, 2)) & LCase(.TextMatrix(i, 1)) Then
                        DontAdd = True
                    End If
                End If
            Next i
            If DontAdd = False Then
                s = s + "|" + X.DeviceName
            Else
                DontAdd = False
            End If
        Next X
    End With
    
    VSFlexGrid1.ComboList = s
   
End Sub

Private Sub VSFlexGrid1_Click()
     With VSFlexGrid1
        If .Col = 4 Then
           .Select .Row, .Col
           .EditCell
        End If
     End With
End Sub

Private Sub VSFlexGrid1_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    
    With VSFlexGrid1
        .Row = Row
        .Col = Col
    End With
    
End Sub

Private Sub FillVSFlexGrid1()

    VSFlexGrid1.Rows = 0
    CmbPrinterNo.Clear
    
    With VSFlexGrid1
        
        Set rctmp = RunStoredProcedure2RecordSet("Get_All_tPrinters")
        
        If Not (rctmp.EOF = True And rctmp.BOF = True) Then
        
            rctmp.MoveFirst
            
            Dim temp As String

            i = 1
            While Not rctmp.EOF
            
                If clsArya.MaxprinterNo >= rctmp!PrinterNo Then
                    .Rows = .Rows + 1
                    .TextMatrix(.Rows - 1, 0) = "Å—Ì‰ — # " & i
                    CmbPrinterNo.AddItem "Å—Ì‰ — #" & i
                    CmbPrinterNo.ItemData(CmbPrinterNo.ListCount - 1) = rctmp!PrinterNo
                    temp = temp & "#" & rctmp!PrinterNo & ";" & "Å—Ì‰ — " & i & "|"
                    FlxDetail.ColComboList(3) = temp
                    .TextMatrix(i - 1, 1) = CStr(rctmp!PrinterName)
                    .TextMatrix(i - 1, 2) = CStr(rctmp!HostName)
                    .TextMatrix(i - 1, 3) = CStr(rctmp!PrinterNo)
                    .TextMatrix(i - 1, 4) = CStr(rctmp!Port)
                End If
                i = i + 1
                rctmp.MoveNext
            Wend
            
            If temp <> "" Then
                temp = Mid(temp, 1, Len(temp) - 1)
            End If
            FlxDetail.ColComboList(3) = temp
            
            For i = i To clsArya.MaxprinterNo
                .Rows = .Rows + 1
                .TextMatrix(.Rows - 1, 0) = "Å—Ì‰ — # " & i
            Next
            
            CmbPrinterNo.ListIndex = 0
        Else
           
            For i = 1 To clsArya.MaxprinterNo
                .Rows = .Rows + 1
                .TextMatrix(.Rows - 1, 0) = "Å—Ì‰ — # " & i
            Next
            
        End If
        
        If CmbPrinterNo.ListCount = 0 Or CmbPrintFormat.ListCount = 0 Or CmbState.ListCount = 0 Or CmbStationNo.ListCount = 0 Then
            FWBtnSave.Enabled = False
        Else
            FWBtnSave.Enabled = True
        End If

    End With
    
End Sub
Private Sub GetDataDetailForEdit()
    
    ''DefaultSettings
    
    Dim TempStr As String
    
    ReDim Parameter(0) As Parameter
    Parameter(0) = GenerateInputParameter("@intPrintingNo", adInteger, 4, Val(intPrintingNo))
    Set rctmp = RunParametricStoredProcedure2Rec("Get_tPrinting_intPrintingNo", Parameter)
    Dim ii As Integer
    If Not (rctmp.BOF Or rctmp.EOF) Then
       
        TxtLineFeed.Text = rctmp!LineFeed
        TxtRepeatNo.Text = rctmp!RepeatNo
        ArmCheck.Value = IIf(rctmp!Arm = True, 1, 0)
        BarcodeCheck.Value = IIf(rctmp!barcode = True, 1, 0)
        CutterCheck.Value = IIf(rctmp!Cutter = True, 1, 0)
        SerialNoCheck.Value = IIf(rctmp!SerialNo = True, 1, 0)
        ViewCheck.Value = IIf(((rctmp!PermittedModes) And (1)) = 1, 1, 0)
        AddCheck.Value = IIf(((rctmp!PermittedModes) And (2)) = 2, 1, 0)
        EditCheck.Value = IIf(((rctmp!PermittedModes) And (4)) = 4, 1, 0)
        ManipulateCheck.Value = IIf(((rctmp!PermittedModes) And (8)) = 8, 1, 0)
        RefferCheck.Value = IIf(((rctmp!PermittedModes) And (16)) = 16, 1, 0)
        InvoiceCheck.Value = IIf(((rctmp!PermittedModes) And (32)) = 32, 1, 0)
        LableCheck.Value = IIf(((rctmp!PermittedModes) And (64)) = 64, 1, 0)
        For i = 0 To CmbStatus.ListCount - 1
            If CmbStatus.ItemData(i) = rctmp!Status Then
                CmbStatus.ListIndex = i
                Exit For
            End If
        Next i
        
        For i = 0 To CmbStationNo.ListCount - 1
            If CmbStationNo.ItemData(i) = rctmp!StationId Then
                CmbStationNo.ListIndex = i
                Exit For
            End If
        Next i
        
        For i = 0 To CmbState.ListCount - 1
            If CmbState.ItemData(i) = rctmp!ServePlace Then
                CmbState.ListIndex = i
                Exit For
            End If
        Next i
        
        For i = 0 To CmbPrinterNo.ListCount - 1
            If CmbPrinterNo.ItemData(i) = rctmp!PrinterNo Then
                CmbPrinterNo.ListIndex = i
                Exit For
            End If
        Next i
        
        For i = 0 To CmbPrintFormat.ListCount - 1
            If CmbPrintFormat.ItemData(i) = rctmp!PrintFormat Then
                CmbPrintFormat.ListIndex = i
                Exit For
            End If
        Next i
        
        For i = 0 To CmbPrintType.ListCount - 1
            If CmbPrintType.ItemData(i) = rctmp!DirectRpt Then
                CmbPrintType.ListIndex = i
                Exit For
            End If
        Next i
        
        For i = 0 To cmbPartition.ListCount - 1
            If cmbPartition.ItemData(i) = rctmp!PartitionID Then
                cmbPartition.ListIndex = i
                Exit For
            End If
        Next i
        
        
   
    End If
    rctmp.Close
    
    
End Sub

Public Sub Edit()

MyFormAddEditMode = EnumAddEditMode.EditMode
mdifrm.Toolbar1.Buttons(7).Enabled = False  'Edit
mdifrm.Toolbar1.Buttons(9).Enabled = True  'Cancel

End Sub
Public Sub Cancel()
MyFormAddEditMode = EnumAddEditMode.AddMode
mdifrm.Toolbar1.Buttons(9).Enabled = False  'Cancel
'ClearDataFlexGrid
'GetDataDetail
End Sub
