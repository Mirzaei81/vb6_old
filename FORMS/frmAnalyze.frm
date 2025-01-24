VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{7AEDC602-D94C-11D1-BB7A-00E0290EA3C9}#1.0#0"; "ResizeKit.ocx"
Object = "{75D4F767-8785-11D3-93AD-0000832EF44D}#2.14#0"; "FAST2001.ocx"
Begin VB.Form frmAnalyze 
   ClientHeight    =   9990
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   13275
   ClipControls    =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   9990
   ScaleWidth      =   13275
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtDate 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   435
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   2040
      Width           =   1485
   End
   Begin VB.CommandButton CancelButton 
      BackColor       =   &H000000C0&
      Caption         =   "Œ—ÊÃ"
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
      TabIndex        =   7
      Top             =   9240
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "Nazanin"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   480
      Width           =   13095
      Begin VB.ComboBox cmbDestInventory 
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "Nazanin"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   4440
         RightToLeft     =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   960
         Width           =   2595
      End
      Begin VB.ComboBox CmbGood 
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
         Left            =   9600
         RightToLeft     =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   960
         Width           =   2445
      End
      Begin VB.ComboBox cmbGoodType 
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
         Left            =   9600
         RightToLeft     =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   360
         Width           =   2445
      End
      Begin VB.TextBox txtFeeUnit 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "B Nazanin"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   570
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   19
         Top             =   840
         Width           =   1575
      End
      Begin VB.ComboBox cmbInventory 
         BackColor       =   &H00FFC0C0&
         BeginProperty Font 
            Name            =   "Nazanin"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   4440
         RightToLeft     =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   360
         Width           =   2595
      End
      Begin VB.TextBox TxtAmount 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "B Nazanin"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   570
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   240
         Width           =   1575
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackColor       =   &H00EACCEC&
         BackStyle       =   0  'Transparent
         Caption         =   "«‰»«— „ﬁ’œ"
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
         Height          =   375
         Left            =   7800
         RightToLeft     =   -1  'True
         TabIndex        =   27
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "ò«·«"
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
         Height          =   465
         Left            =   11880
         RightToLeft     =   -1  'True
         TabIndex        =   24
         Top             =   960
         Width           =   975
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "‰Ê⁄ ò«·«"
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
         Height          =   465
         Left            =   11880
         RightToLeft     =   -1  'True
         TabIndex        =   23
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00EACCEC&
         BackStyle       =   0  'Transparent
         Caption         =   "›Ì ÕÊ«·Â"
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000001&
         Height          =   375
         Left            =   2160
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Top             =   960
         Width           =   855
      End
      Begin VB.Label lblInventory 
         Alignment       =   2  'Center
         BackColor       =   &H00EACCEC&
         BackStyle       =   0  'Transparent
         Caption         =   "«‰»«— „»œ«"
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
         Height          =   375
         Left            =   7800
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackColor       =   &H00EACCEC&
         BackStyle       =   0  'Transparent
         Caption         =   "„ﬁœ«—ÕÊ«·Â"
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000001&
         Height          =   375
         Left            =   2040
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.CommandButton CmdDone 
      BackColor       =   &H00008000&
      Caption         =   " «ÌÌœ"
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
      Left            =   11160
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   9240
      Width           =   1785
   End
   Begin VB.CommandButton cmddeSelectGoods 
      Caption         =   ">"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   885
      Left            =   6720
      TabIndex        =   2
      Top             =   5160
      Width           =   885
   End
   Begin VB.CommandButton cmdSelectGoods 
      Caption         =   "<"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   885
      Left            =   6720
      TabIndex        =   1
      Top             =   4080
      Width           =   885
   End
   Begin VB.CommandButton CmdDelete 
      BackColor       =   &H000000C0&
      Cancel          =   -1  'True
      Caption         =   "Õ–›  „Ê«—œ¬‰«·Ì“ ‘œÂ"
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
      Left            =   7680
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   9240
      Width           =   2175
   End
   Begin VSFlex7LCtl.VSFlexGrid vsDefinedGoods 
      Height          =   5685
      Left            =   120
      TabIndex        =   8
      Top             =   2520
      Width           =   6555
      _cx             =   11562
      _cy             =   10028
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
      ForeColor       =   8388608
      BackColorFixed  =   12632064
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
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   3
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   3
      FixedRows       =   1
      FixedCols       =   0
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
   Begin FLWCtrls.FWLabel3D fwlblMode 
      Height          =   525
      Left            =   11640
      Top             =   0
      Visible         =   0   'False
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   926
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
   Begin VSFlex7LCtl.VSFlexGrid vsNotDefinedGoods 
      Height          =   5685
      Left            =   7770
      TabIndex        =   9
      Top             =   2520
      Width           =   5445
      _cx             =   9604
      _cy             =   10028
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
      ForeColor       =   128
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
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   3
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   2
      FixedRows       =   1
      FixedCols       =   0
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
      OleObjectBlob   =   "frmAnalyze.frx":0000
      TabIndex        =   10
      Top             =   0
      Width           =   480
   End
   Begin VB.Label lblDate 
      BackColor       =   &H00EACCEC&
      BackStyle       =   0  'Transparent
      Caption         =   " «—ÌŒ ¬„«œÂ ”«“Ì"
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
      Left            =   2040
      RightToLeft     =   -1  'True
      TabIndex        =   29
      Top             =   2040
      Width           =   1485
   End
   Begin VB.Label lblBuyPriceFirst 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "»Â«¡ »— «”«” ﬁÌ„  Œ—Ìœ ‰«Œ«·’"
      BeginProperty Font 
         Name            =   "B Homa"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   465
      Left            =   2280
      RightToLeft     =   -1  'True
      TabIndex        =   25
      Top             =   9480
      Width           =   4695
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "¬‰«·Ì“ Ê ¬„«œÂ ”«“Ì „Ê«œ «Ê·ÌÂ"
      BeginProperty Font 
         Name            =   "Nazanin"
         Size            =   15.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   495
      Left            =   4920
      RightToLeft     =   -1  'True
      TabIndex        =   16
      Top             =   0
      Width           =   3735
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H008080FF&
      BackStyle       =   0  'Transparent
      Caption         =   "ò«·«Â«Ì ¬„«œÂ ‘œÂ"
      BeginProperty Font 
         Name            =   "B Homa"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008080&
      Height          =   465
      Left            =   4200
      RightToLeft     =   -1  'True
      TabIndex        =   15
      Top             =   2040
      Width           =   2445
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "ò«·«Â«Ì  ⁄—Ì› ‰‘œÂ œ— „‰Ê  Ê·ÌœÌ"
      BeginProperty Font 
         Name            =   "B Homa"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   465
      Left            =   9480
      RightToLeft     =   -1  'True
      TabIndex        =   14
      Top             =   2085
      Width           =   3615
   End
   Begin VB.Label LblPercent 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "Ã„⁄ œ—’œ"
      BeginProperty Font 
         Name            =   "B Homa"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   465
      Left            =   2280
      RightToLeft     =   -1  'True
      TabIndex        =   13
      Top             =   8205
      Width           =   4695
   End
   Begin VB.Label LblsellPrice 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "»Â«¡ »— «”«” ﬁÌ„  ›—Ê‘"
      BeginProperty Font 
         Name            =   "B Homa"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   465
      Left            =   2280
      RightToLeft     =   -1  'True
      TabIndex        =   12
      Top             =   9000
      Width           =   4695
   End
   Begin VB.Label LblBuyPrice 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "»Â«¡ »— «”«” ﬁÌ„  Œ—Ìœ"
      BeginProperty Font 
         Name            =   "B Homa"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   465
      Left            =   2400
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   8580
      Width           =   4575
   End
End
Attribute VB_Name = "frmAnalyze"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MyFormAddEditMode As EnumAddEditMode
Dim i As Long
Dim Parameter() As Parameter
Dim formloadFlag As Boolean
Dim Rst As New Recordset
Const IndexColRow = 0
Const IndexColID = 1
Const IndexColName = 2
Const IndexColBuyPrice = 3
Const IndexColSellPrice = 4
Const IndexColWeight = 5
Const IndexColPercent = 6
Const IndexColAmount = 7
Const CountColDefine = 8
Const CountColUnDefine = 6
Dim sumPrice As Double

Private Sub CancelButton_Click()
    Unload Me
End Sub


Private Sub CmbGood_Click()
    FillGrid

End Sub
Private Sub InsertGoodAnalyze()
    
    If CmbGood.ListIndex = -1 Then Exit Sub
    Dim i As Integer
    Dim j As Integer
    Dim strSelectedSeller As String
    Dim s As String
    Dim U As String
    s = "": U = ""
    With vsDefinedGoods
        For i = 1 To .Rows - 1
            s = s & .TextMatrix(i, IndexColID) & ","
            U = U & CDbl(.TextMatrix(i, IndexColPercent)) & ","
        Next i
        
        If s <> "" Then
            s = left(s, Len(s) - 1)
            U = left(U, Len(U) - 1)
        End If
        ReDim Parameter(2) As Parameter
        Parameter(0) = GenerateInputParameter("@GoodCode", adInteger, 4, CmbGood.ItemData(CmbGood.ListIndex))
        Parameter(1) = GenerateInputParameter("@nvcGoodFirstCode", adVarWChar, 4000, s)
        Parameter(2) = GenerateInputParameter("@nvcfltUsedValue", adVarWChar, 4000, U)
        
        RunParametricStoredProcedure "Insert_tblTotal_Good_Analyze", Parameter
    
    End With
End Sub

Private Sub cmdDelete_Click()
    
    If CmbGood.ListIndex = -1 Then Exit Sub
    
    ShowMessage "¬Ì« „Ì ŒÊ«ÂÌœ „Ê«—œ  ⁄—Ì› ‘œÂ «“ ·Ì”  Õ–› ‘Ê‰œø", True, True, "»·Ì", "ŒÌ—"
    If mvarMsgIdx = vbYes Then
    
        ReDim Parameter(0) As Parameter
        Parameter(0) = GenerateInputParameter("@GoodCode", adInteger, 4, CmbGood.ItemData(CmbGood.ListIndex))
        RunParametricStoredProcedure "Delete_tblTotal_Good_Analyze", Parameter
        FillGrid
    
    End If
End Sub

Private Sub CmdDone_Click()

    If CmbGood.ListIndex = -1 Then ShowDisMessage "ò«·« «‰ Œ«» ‰‘œÂ", 1500: Exit Sub
    If cmbInventory.ListIndex = -1 Then ShowDisMessage "«‰»«— „»œ« —« «‰ Œ«» ò‰Ìœ", 1500: Exit Sub
    If cmbDestInventory.ListIndex = -1 Then ShowDisMessage "«‰»«— „ﬁ’œ —« «‰ Œ«» ò‰Ìœ", 1500: Exit Sub
    If cmbInventory.ListIndex = cmbDestInventory.ListIndex Then ShowDisMessage "«‰»«— „»œ« Ê „ﬁ’œ »«Ìœ „ ›«Ê  »«‘‰œ", 1500: Exit Sub
    If Len(txtDate.Text) <> 8 Then ShowDisMessage " «—ÌŒ —« Ê«—œ ò‰Ìœ", 1500: Exit Sub
    With vsDefinedGoods
        If .Rows = 1 Then Exit Sub
        DetailsString = ""
        For i = 1 To .Rows - 1
            If Val(Trim(.TextMatrix(i, IndexColAmount))) = 0 Then
                ShowMessage "„ﬁœ«— ’›— ﬁ«»· ﬁ»Ê· ‰Ì”  - " & .TextMatrix(i, IndexColName), True, False, "ﬁ»Ê·", ""
                Exit Sub
'            Else
'                DetailsString = GenerateDetailsString3(DetailsString, .TextMatrix(i, IndexColAmount), .TextMatrix(i, IndexColID), Val(.TextMatrix(i, IndexColBuyPrice)), 0, 0, " ", "", cmbInventory.ItemData(cmbInventory.ListIndex), cmbDestInventory.ItemData(cmbDestInventory.ListIndex), 1, "")
            End If
        Next i
        InsertGoodAnalyze
        If LCase(VarActForm) = "frmpurchase" Then
            '''  À»  ÕÊ«·Â
            ShowDisMessage "„—Õ·Â «Ê· - À»  ÕÊ«·Â ò«·«Ì ‰«Œ«·’ «“ «‰»«—  " & cmbInventory.Text, 1500
            frmPurchase.txtDate = txtDate
            For i = 0 To frmPurchase.CmbStatus.ListCount - 1
                If frmPurchase.CmbStatus.ItemData(i) = EnumFactorType.fromStore Then
                    frmPurchase.CmbStatus.ListIndex = i: Exit For
                End If
            Next
            For i = 0 To frmPurchase.cmbInventory.ListCount - 1
                If frmPurchase.cmbInventory.ItemData(i) = cmbInventory.ItemData(cmbInventory.ListIndex) Then
                    frmPurchase.cmbInventory.ListIndex = i: Exit For
                End If
            Next
''            For i = 0 To frmPurchase.cmbDestInventory.ListCount - 1
''                If frmPurchase.cmbDestInventory.ItemData(i) = cmbDestInventory.ItemData(cmbDestInventory.ListIndex) Then
''                    frmPurchase.cmbDestInventory.ListIndex = i: Exit For
''                End If
''            Next
            mvarcode = CmbGood.ItemData(CmbGood.ListIndex)
            If frmPurchase.GetGoodCode(mvarcode) = True Then
                frmPurchase.lblNum = TxtAmount
                frmPurchase.ChangeGoodquantity
                If frmPurchase.Update <> -1 Then
                    ShowDisMessage "À»  «” «‰œ«—œ ÕÊ«·Â ò«·«Ì ‰«Œ«·’  «‰Ã«„ ê—›  ", 1000
                Else
                    ShowMessage "œ— À»  «” «‰œ«—œ ÕÊ«·Â ò«·«Ì ‰«Œ«·’ „‘ò· ÊÃÊœ œ«—œ", True, False, "ﬁ»Ê·", ""
                    Exit Sub
                End If
            Else
                ShowMessage "œ— «‰ Œ«» ò«·«Ì ‰«Œ«·’ »—«Ì «” «‰œ«—œ ÕÊ«·Â „‘ò· ÊÃÊœ œ«—œ", True, False, "ﬁ»Ê·", ""
                Exit Sub
            End If
            
'' Change Havale to StandardHavale
''Do not neccessary Resid for standard
'''            '''  À»  —”Ìœ
            ShowDisMessage "„—Õ·Â œÊ„ -À»   —”Ìœ «” «‰œ«—œ ò«·«Â«Ì ¬„«œÂ ‘œÂ »Â «‰»«—  " & cmbDestInventory.Text, 1500
            For i = 0 To frmPurchase.CmbStatus.ListCount - 1
                If frmPurchase.CmbStatus.ItemData(i) = EnumFactorType.toStore Then
                    frmPurchase.CmbStatus.ListIndex = i: Exit For
                End If
            Next
            For i = 0 To frmPurchase.cmbInventory.ListCount - 1
                If frmPurchase.cmbInventory.ItemData(i) = cmbDestInventory.ItemData(cmbDestInventory.ListIndex) Then
                    frmPurchase.cmbInventory.ListIndex = i: Exit For
                End If
            Next
''            For i = 0 To frmPurchase.cmbDestInventory.ListCount - 1
''                If frmPurchase.cmbDestInventory.ItemData(i) = cmbDestInventory.ItemData(cmbDestInventory.ListIndex) Then
''                    frmPurchase.cmbDestInventory.ListIndex = i: Exit For
''                End If
''            Next
            For i = 1 To .Rows - 1
                mvarcode = .TextMatrix(i, IndexColID)
                If frmPurchase.GetGoodCode(mvarcode) = True Then
                    frmPurchase.lblNum = .TextMatrix(i, IndexColAmount)
                    frmPurchase.ChangeGoodquantity
                Else
                    ShowMessage "œ— «‰ Œ«»  ò«·«Â«Ì ¬„«œÂ ‘œÂ »—«Ì ÕÊ«·Â „‘ò· ÊÃÊœ œ«—œ", True, False, "ﬁ»Ê·", ""
                    Exit Sub
                End If
            Next
            If frmPurchase.Update <> -1 Then
                ShowDisMessage "À»  ÕÊ«·Â Ê —”Ìœ  ò«·«Â«Ì ¬„«œÂ ‘œÂ «‰Ã«„ ê—›  ", 1000
            Else
                ShowMessage "œ— À»  ÕÊ«·Â Ê —”Ìœ  ò«·«Â«Ì ¬„«œÂ ‘œÂ „‘ò· ÊÃÊœ œ«—œ", True, False, "ﬁ»Ê·", ""
                Exit Sub
            End If
        End If
    
    End With
    If PosConnection.State = adStateOpen Then PosConnection.Close: Set PosConnection = Nothing
    Unload Me
End Sub
Private Sub cmbGoodType_Click()
    FillGoods
End Sub
Private Sub cmddeSelectGoods_Click()
    If vsDefinedGoods.SelectedRows = 0 Then Exit Sub
    For i = 0 To vsDefinedGoods.SelectedRows - 1
        vsNotDefinedGoods.Rows = vsNotDefinedGoods.Rows + 1
        vsNotDefinedGoods.TextMatrix(vsNotDefinedGoods.Rows - 1, IndexColID) = vsDefinedGoods.TextMatrix(vsDefinedGoods.SelectedRow(i), IndexColID)
        vsNotDefinedGoods.TextMatrix(vsNotDefinedGoods.Rows - 1, IndexColName) = vsDefinedGoods.TextMatrix(vsDefinedGoods.SelectedRow(i), IndexColName)
        vsNotDefinedGoods.TextMatrix(vsNotDefinedGoods.Rows - 1, IndexColBuyPrice) = vsDefinedGoods.TextMatrix(vsDefinedGoods.SelectedRow(i), IndexColBuyPrice)
        vsNotDefinedGoods.TextMatrix(vsNotDefinedGoods.Rows - 1, IndexColSellPrice) = vsDefinedGoods.TextMatrix(vsDefinedGoods.SelectedRow(i), IndexColSellPrice)
    Next i
    For i = 0 To vsDefinedGoods.SelectedRows - 1
        vsDefinedGoods.RemoveItem vsDefinedGoods.SelectedRow(0)
    Next i
End Sub
Private Sub cmdSelectGoods_Click()
    If vsNotDefinedGoods.SelectedRows = 0 Then Exit Sub
    For i = 0 To vsNotDefinedGoods.SelectedRows - 1
        vsDefinedGoods.Rows = vsDefinedGoods.Rows + 1
        vsDefinedGoods.TextMatrix(vsDefinedGoods.Rows - 1, IndexColRow) = vsDefinedGoods.Rows - 1
        vsDefinedGoods.TextMatrix(vsDefinedGoods.Rows - 1, IndexColID) = vsNotDefinedGoods.TextMatrix(vsNotDefinedGoods.SelectedRow(i), IndexColID)
        vsDefinedGoods.TextMatrix(vsDefinedGoods.Rows - 1, IndexColName) = vsNotDefinedGoods.TextMatrix(vsNotDefinedGoods.SelectedRow(i), IndexColName)
        vsDefinedGoods.TextMatrix(vsDefinedGoods.Rows - 1, IndexColBuyPrice) = vsNotDefinedGoods.TextMatrix(vsNotDefinedGoods.SelectedRow(i), IndexColBuyPrice)
        vsDefinedGoods.TextMatrix(vsDefinedGoods.Rows - 1, IndexColSellPrice) = vsNotDefinedGoods.TextMatrix(vsNotDefinedGoods.SelectedRow(i), IndexColSellPrice)
        vsDefinedGoods.TextMatrix(vsDefinedGoods.Rows - 1, IndexColWeight) = vsNotDefinedGoods.TextMatrix(vsNotDefinedGoods.SelectedRow(i), IndexColWeight)
        vsDefinedGoods.TextMatrix(vsDefinedGoods.Rows - 1, IndexColPercent) = 0
        vsDefinedGoods.TextMatrix(vsDefinedGoods.Rows - 1, IndexColAmount) = 0
    Next i
    For i = 0 To vsNotDefinedGoods.SelectedRows - 1
        vsNotDefinedGoods.RemoveItem vsNotDefinedGoods.SelectedRow(0)
    Next i
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call PresetScreenSaver
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call PresetScreenSaver
End Sub
Private Sub Form_Load()
    
     mvarAnalyzeForm = True  '' For Purchase Form to not show message
    CenterCenterinSecondScreen Me
    formloadFlag = False
    Me.left = Val(GetSetting(strMainKey, Me.Name, "Left"))
    If Val(GetSetting(strMainKey, Me.Name, "Height")) > 5000 Then
        Me.Height = Val(GetSetting(strMainKey, Me.Name, "Height"))
    End If
    If Val(GetSetting(strMainKey, Me.Name, "Width")) > 5000 Then
        Me.Width = Val(GetSetting(strMainKey, Me.Name, "Width"))
    End If
    Me.top = Val(GetSetting(strMainKey, Me.Name, "Top"))
    If Me.left < 0 Then Me.left = 0
    If Me.top < 0 Then Me.top = 0
    If Me.top > Me.ScaleHeight Then Me.top = 0

    MyFormAddEditMode = ViewMode
    formloadFlag = True
    
    With vsDefinedGoods
        .Rows = 1
        .Cols = CountColDefine
        .TextMatrix(0, IndexColRow) = "—œÌ›"
        .TextMatrix(0, IndexColID) = "òœ ò«·«"
        .TextMatrix(0, IndexColName) = "‰«„ ò«·«"
        .TextMatrix(0, IndexColBuyPrice) = "ﬁÌ„  Œ—Ìœ"
        .TextMatrix(0, IndexColSellPrice) = "ﬁÌ„  ›—Ê‘"
        .TextMatrix(0, IndexColWeight) = "Ê“‰ Ê«Õœ"
        .TextMatrix(0, IndexColPercent) = "œ—’œ Œ«·’"
        .TextMatrix(0, IndexColAmount) = "„ﬁœ«—"
        .ColHidden(IndexColID) = True
        .ColDataType(IndexColPercent) = flexDTDecimal
        .ColDataType(IndexColAmount) = flexDTDecimal
        .FixedAlignment(-1) = flexAlignCenterCenter
        .ColAlignment(-1) = flexAlignCenterCenter
        .ColAlignment(IndexColName) = flexAlignRightCenter
        For i = 0 To .Cols - 1
            .ColWidth(i) = Val(GetSetting(strMainKey, Me.Name & "vsDefinedGoods", "Col" & i))
            If .ColWidth(i) = 0 Then
                .ColWidth(i) = .Width / 10     'Row
            End If
        Next i
    End With
    With vsNotDefinedGoods
        .Rows = 1
        .Cols = CountColUnDefine
        .TextMatrix(0, IndexColRow) = "—œÌ›"
        .TextMatrix(0, IndexColID) = "òœ ò«·«"
        .TextMatrix(0, IndexColName) = "‰«„ ò«·«"
        .TextMatrix(0, IndexColBuyPrice) = "ﬁÌ„  Œ—Ìœ"
        .TextMatrix(0, IndexColSellPrice) = "ﬁÌ„  ›—Ê‘"
        .TextMatrix(0, IndexColWeight) = "Ê“‰ Ê«Õœ"
        .ColHidden(IndexColID) = True
        .FixedAlignment(-1) = flexAlignCenterCenter
        .ColAlignment(-1) = flexAlignCenterCenter
        .ColAlignment(IndexColName) = flexAlignRightCenter
        For i = 0 To .Cols - 1
            .ColWidth(i) = Val(GetSetting(strMainKey, Me.Name & "vsNotDefinedGoods", "Col" & i))
            If .ColWidth(i) = 0 Then
                .ColWidth(i) = .Width / 10     'Row
            End If
        Next i
    End With

    FillComb
    txtDate = Mid(clsDate.shamsi(Date), 3, 8)
End Sub
Private Sub Form_Unload(Cancel As Integer)
    If PosConnection.State = adStateOpen Then PosConnection.Close: Set PosConnection = Nothing
    SaveSetting strMainKey, Me.Name, "Left", Me.left
    SaveSetting strMainKey, Me.Name, "Top", Me.top
     mvarAnalyzeForm = False  '' For Purchase Form to not show message

End Sub
Private Sub FillGrid()
    DoEvents
    If CmbGood.ListIndex = -1 Then Exit Sub
    
    ReDim Parameter(4) As Parameter
    Parameter(0) = GenerateInputParameter("@Code", adInteger, 4, CmbGood.ItemData(CmbGood.ListIndex))
    Parameter(1) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
    Parameter(2) = GenerateInputParameter("@StationId", adInteger, 4, clsArya.StationNo)
    Parameter(3) = GenerateInputParameter("@Flag", adBoolean, 1, 0)
    Parameter(4) = GenerateInputParameter("@AccountYear", adSmallInt, 2, AccountYear)
    Set Rst = RunParametricStoredProcedure2Rec("Get_Good_Code", Parameter)
    
    If Not (Rst.BOF Or Rst.EOF) Then
        txtFeeUnit = Rst!BuyPrice
    End If
    vsNotDefinedGoods.Rows = 1
    vsDefinedGoods.Rows = 1
    If Rst.State <> 0 Then Rst.Close
    ReDim Parameter(0) As Parameter
    Parameter(0) = GenerateInputParameter("@GoodCode", adInteger, 4, CmbGood.ItemData(CmbGood.ListIndex))
    Set Rst = RunParametricStoredProcedure2Rec("Get_undefined_Good", Parameter)
    While Rst.EOF <> True
        With vsNotDefinedGoods
            .Rows = .Rows + 1
            i = .Rows - 1
            .TextMatrix(i, IndexColRow) = i
            .TextMatrix(i, IndexColID) = Rst.Fields("code").Value
            .TextMatrix(i, IndexColName) = Rst.Fields("Name").Value
            .TextMatrix(i, IndexColBuyPrice) = Rst.Fields("BuyPrice").Value
            .TextMatrix(i, IndexColSellPrice) = Rst.Fields("SellPrice").Value
            .TextMatrix(i, IndexColWeight) = Rst.Fields("Weight").Value
        End With
        Rst.MoveNext
    Wend
    
    If Rst.State <> 0 Then Rst.Close
    ReDim Parameter(0) As Parameter
    Parameter(0) = GenerateInputParameter("@GoodCode", adInteger, 4, CmbGood.ItemData(CmbGood.ListIndex))
    Set Rst = RunParametricStoredProcedure2Rec("Get_defined_Good", Parameter)
    While Rst.EOF <> True
        With vsDefinedGoods
            .Rows = .Rows + 1
            i = .Rows - 1
            .TextMatrix(i, IndexColRow) = i
            .TextMatrix(i, IndexColID) = Rst.Fields("code").Value
            .TextMatrix(i, IndexColName) = Rst.Fields("Name").Value
            .TextMatrix(i, IndexColBuyPrice) = Rst.Fields("BuyPrice").Value
            .TextMatrix(i, IndexColSellPrice) = Rst.Fields("SellPrice").Value
            .TextMatrix(i, IndexColPercent) = Rst.Fields("fltUsedValue").Value
            .TextMatrix(i, IndexColWeight) = Rst.Fields("Weight").Value
        End With
        Rst.MoveNext
    Wend
    vsDefinedGoods_AfterEdit 0, 0
    If Rst.State = adStateOpen Then Rst.Close: Set Rst = Nothing
    If PosConnection.State = adStateOpen Then PosConnection.Close: Set PosConnection = Nothing
End Sub
Public Sub ExitForm()
    Unload Me
End Sub
Private Sub ResizeKit1_ExitResize(ByVal XScale As Double, ByVal YScale As Double)
    If formloadFlag = True Then
        SaveSetting strMainKey, Me.Name, "Height", Me.Height
        SaveSetting strMainKey, Me.Name, "Width", Me.Width
    End If
End Sub

Private Sub TxtAmount_Change()
    
    sumPrice = Val(TxtAmount) * Val(txtFeeUnit)
    CalculateSumPrice

End Sub

Private Sub txtDate_KeyPress(KeyAscii As Integer)
If Len(txtDate.Text) >= 8 And (KeyAscii >= 48 And KeyAscii <= 57) Then
    KeyAscii = 0
    Exit Sub
End If
If txtDate.SelStart = 0 Then Exit Sub
If KeyAscii = 8 Then
    If Len(txtDate.Text) = txtDate.SelStart Then
        Exit Sub
    End If
    If Mid(txtDate.Text, txtDate.SelStart, 1) = "/" Then
        KeyAscii = 0
        Exit Sub
    End If
    Exit Sub
End If
If KeyAscii < 48 Or KeyAscii > 57 Then
    KeyAscii = 0
    Exit Sub
End If
If Len(txtDate.Text) <> txtDate.SelStart Then
    Exit Sub
End If
Select Case Len(txtDate.Text)
Case 2
    txtDate.Text = txtDate.Text & "/"
    txtDate.SelStart = Len(txtDate.Text) + 1
Case 5
    txtDate.Text = txtDate.Text & "/"
    txtDate.SelStart = Len(txtDate.Text) + 1
End Select

End Sub

Private Sub txtDate_KeyDown(KeyCode As Integer, Shift As Integer)
    If txtDate.Locked = True Then Exit Sub
    If KeyCode = vbKeyDelete Then
        If Mid(txtDate.Text, txtDate.SelStart + 1, 1) = "/" Then
            KeyCode = 0
        End If
    End If

'    If KeyCode = vbKeyDelete Then
'        If Mid(txtDate.Text, txtDate.SelStart + 1, 1) = "/" Then
'            KeyCode = 0
'        End If
'    End If
End Sub

Private Sub txtFeeUnit_Change()
    
    sumPrice = Val(TxtAmount) * Val(txtFeeUnit)
    CalculateSumPrice
End Sub

Private Sub CalculateSumPrice()
    
    If sumPrice = 0 Then lblBuyPriceFirst = "»Â«¡ »— «”«” ﬁÌ„  Œ—Ìœ ‰«Œ«·’:  " Else lblBuyPriceFirst = "»Â«¡ »— «”«” ﬁÌ„  Œ—Ìœ ‰«Œ«·’:  " & Format(sumPrice, "#,## ") & clsArya.UnitPrice
    DoCalculate
End Sub
Private Sub DoCalculate()
    With vsDefinedGoods
        For i = 1 To .Rows - 1
            If Val(.TextMatrix(i, IndexColWeight)) <> 0 Then .TextMatrix(i, IndexColAmount) = CDbl(.TextMatrix(i, IndexColPercent)) * Val(TxtAmount) / (100 * Val(.TextMatrix(i, IndexColWeight)))
           .TextMatrix(i, IndexColAmount) = CLng(.TextMatrix(i, IndexColAmount))
        Next
    End With
    CalculateLables
End Sub
Private Sub CalculateLables()
    Dim BuyPrice, SellPrice As Long
    BuyPrice = 0: SellPrice = 0
    With vsDefinedGoods
        For i = 1 To .Rows - 1
            BuyPrice = BuyPrice + CLng(Val(.TextMatrix(i, IndexColAmount)) * Val(.TextMatrix(i, IndexColBuyPrice)))
            SellPrice = SellPrice + CLng(Val(.TextMatrix(i, IndexColAmount)) * Val(.TextMatrix(i, IndexColSellPrice)))
        Next
    End With
    If BuyPrice = 0 Then LblBuyPrice = "»Â«¡ »— «”«” ﬁÌ„  Œ—Ìœ :  " Else LblBuyPrice = "»Â«¡ »— «”«” ﬁÌ„  Œ—Ìœ :  " & Format(BuyPrice, "#,## ") & clsArya.UnitPrice
    If SellPrice = 0 Then LblsellPrice = "»Â«¡ »— «”«” ﬁÌ„  ›—Ê‘ :  " Else LblsellPrice = "»Â«¡ »— «”«” ﬁÌ„  ›—Ê‘ :  " & Format(SellPrice, "#,##  ") & clsArya.UnitPrice
End Sub


Private Sub vsDefinedGoods_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    
    Dim Percent As Double
    Percent = 0
    With vsDefinedGoods
        For i = 1 To .Rows - 1
            If .TextMatrix(i, IndexColPercent) = "" Then .TextMatrix(i, IndexColPercent) = 0
            Percent = Percent + CDbl(.TextMatrix(i, IndexColPercent))
        Next
    End With
    LblPercent.Caption = "Ã„⁄ œ—’œ :  " & Percent
    If Col = IndexColPercent Then DoCalculate Else CalculateLables
End Sub

Private Sub vsDefinedGoods_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    For i = 0 To vsDefinedGoods.Cols - 1
        SaveSetting strMainKey, Me.Name & "vsDefinedGoods", "Col" & i, vsDefinedGoods.ColWidth(i)
    Next
End Sub

Private Sub vsDefinedGoods_Click()
    With vsDefinedGoods
        If .Row >= 1 And (.Col = IndexColPercent Or .Col = IndexColAmount) Then
            .Select .Row, .Col
            .EditCell
        End If
    End With
End Sub
Private Sub FillComb()
    
    ReDim Parameter(0) As Parameter
    Parameter(0) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
    Set Rst = RunParametricStoredProcedure2Rec("GetGoodType", Parameter)
    If Not (Rst.EOF = True And Rst.BOF = True) Then
        While Not Rst.EOF
            cmbGoodType.AddItem Rst.Fields("Description").Value
            cmbGoodType.ItemData(cmbGoodType.ListCount - 1) = Rst.Fields("Code").Value
            Rst.MoveNext
        Wend
    End If
    
    If Rst.State = 1 Then Rst.Close
    cmbInventory.Clear
    
    ReDim Parameter(1) As Parameter
    Parameter(0) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
    Parameter(1) = GenerateInputParameter("@Type", adInteger, 4, 1) 'All Inventory
    Set Rst = RunParametricStoredProcedure2Rec("GetInventory", Parameter)
    
    If Not (Rst.EOF = True And Rst.BOF = True) Then
        Do While Rst.EOF <> True
            cmbInventory.AddItem Rst.Fields("Description")
            cmbInventory.ItemData(cmbInventory.ListCount - 1) = Val(Rst.Fields("InventoryNo"))
            Rst.MoveNext
        Loop
    End If
    cmbInventory.ListIndex = 0
    
    If Rst.State = 1 Then Rst.Close
    cmbDestInventory.Clear
    
    ReDim Parameter(1) As Parameter
    Parameter(0) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
    Parameter(1) = GenerateInputParameter("@Type", adInteger, 4, 1)  ' All Inventory
    Set Rst = RunParametricStoredProcedure2Rec("GetInventory", Parameter)
    
    If Not (Rst.EOF = True And Rst.BOF = True) Then
        Do While Rst.EOF <> True
            cmbDestInventory.AddItem Rst.Fields("Description")
            cmbDestInventory.ItemData(cmbDestInventory.ListCount - 1) = Val(Rst.Fields("InventoryNo"))
            Rst.MoveNext
        Loop
         
    End If
    cmbDestInventory.ListIndex = 0
    Rst.Close
    
    If Rst.State = adStateOpen Then Rst.Close: Set Rst = Nothing
    If PosConnection.State = adStateOpen Then PosConnection.Close: Set PosConnection = Nothing

End Sub
Private Sub FillGoods()
    CmbGood.Clear
    If cmbGoodType.ListIndex = -1 Then Exit Sub
    ReDim Parameter(0) As Parameter
    Parameter(0) = GenerateInputParameter("@GoodType", adInteger, 4, cmbGoodType.ItemData(cmbGoodType.ListIndex))
    Set Rst = RunParametricStoredProcedure2Rec("Get_Good_By_GoodType", Parameter)
    If Not (Rst.EOF = True And Rst.BOF = True) Then
        While Not Rst.EOF
            CmbGood.AddItem Rst.Fields("Name").Value
            CmbGood.ItemData(CmbGood.ListCount - 1) = Rst.Fields("Code").Value
            Rst.MoveNext
        Wend
    End If
    
    If Rst.State = adStateOpen Then Rst.Close: Set Rst = Nothing
    If PosConnection.State = adStateOpen Then PosConnection.Close: Set PosConnection = Nothing

End Sub

Private Sub vsDefinedGoods_DblClick()
   ' cmddeSelectGoods_Click
End Sub

Private Sub vsNotDefinedGoods_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    For i = 0 To vsNotDefinedGoods.Cols - 1
        SaveSetting strMainKey, Me.Name & "vsNotDefinedGoods", "Col" & i, vsNotDefinedGoods.ColWidth(i)
    Next
End Sub


Private Sub vsNotDefinedGoods_DblClick()
    cmdSelectGoods_Click
End Sub

