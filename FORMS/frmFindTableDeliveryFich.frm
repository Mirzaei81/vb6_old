VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{7AEDC602-D94C-11D1-BB7A-00E0290EA3C9}#1.0#0"; "ResizeKit.ocx"
Object = "{75D4F767-8785-11D3-93AD-0000832EF44D}#2.14#0"; "FAST2001.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{687FF23F-0E9B-449D-A782-2C5E7116EDB0}#1.0#0"; "Fardate.ocx"
Begin VB.Form frmFindTableDeliveryFich 
   BackColor       =   &H00E0E0E0&
   ClientHeight    =   9960
   ClientLeft      =   2775
   ClientTop       =   3870
   ClientWidth     =   12600
   Icon            =   "frmFindTableDeliveryFich.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   9960
   ScaleWidth      =   12600
   Begin VB.Frame Frame1 
      Height          =   2415
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   20
      Top             =   6720
      Width           =   12375
      Begin VSFlex7LCtl.VSFlexGrid vsFactorDetail 
         Height          =   1935
         Left            =   8040
         TabIndex        =   21
         Top             =   240
         Width           =   4200
         _cx             =   7408
         _cy             =   3413
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
         ForeColor       =   -2147483640
         BackColorFixed  =   16777088
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
         FormatString    =   $"frmFindTableDeliveryFich.frx":A4C2
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
      Begin VSFlex7LCtl.VSFlexGrid VSReceived 
         Height          =   1935
         Left            =   3960
         TabIndex        =   29
         Top             =   240
         Width           =   3960
         _cx             =   6985
         _cy             =   3413
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
         ForeColor       =   -2147483640
         BackColorFixed  =   8454143
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
         FormatString    =   $"frmFindTableDeliveryFich.frx":A555
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
      Begin VSFlex7LCtl.VSFlexGrid VSHistory 
         Height          =   1935
         Left            =   120
         TabIndex        =   30
         Top             =   240
         Width           =   3720
         _cx             =   6562
         _cy             =   3413
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
         ForeColor       =   -2147483640
         BackColorFixed  =   8438015
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
         FormatString    =   $"frmFindTableDeliveryFich.frx":A5EB
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
   End
   Begin VB.CommandButton OKButton 
      BackColor       =   &H0000C000&
      Caption         =   "ﬁ»Ê·"
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
      TabIndex        =   5
      Top             =   9240
      Width           =   1215
   End
   Begin VB.CommandButton CancelButton 
      BackColor       =   &H000000C0&
      Caption         =   "«‰’—«›"
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
      TabIndex        =   4
      Top             =   9240
      Width           =   1215
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6495
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   12405
      _ExtentX        =   21881
      _ExtentY        =   11456
      _Version        =   393216
      Tabs            =   5
      TabsPerRow      =   5
      TabHeight       =   706
      WordWrap        =   0   'False
      ForeColor       =   128
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Titr"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "›Ì‘"
      TabPicture(0)   =   "frmFindTableDeliveryFich.frx":A681
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label4"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label7"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label11"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label12"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label13"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label14"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "ShapeBalance"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "ShapeRecursive"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label15"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "txtDate2"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "txtDate1"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "vsFactors_Fich"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "FWProgressBar1"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "TxtNo_Fich"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "TxtNo_Temp"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Picture1"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).ControlCount=   18
      TabCaption(1)   =   "„Ì“"
      TabPicture(1)   =   "frmFindTableDeliveryFich.frx":A69D
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label1"
      Tab(1).Control(1)=   "Label6"
      Tab(1).Control(2)=   "vsFactors_Table"
      Tab(1).Control(3)=   "txtNo_Table"
      Tab(1).ControlCount=   4
      TabCaption(2)   =   "«—”«·Ì"
      TabPicture(2)   =   "frmFindTableDeliveryFich.frx":A6B9
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label8"
      Tab(2).Control(1)=   "Label5"
      Tab(2).Control(2)=   "vsFactors_Delivery"
      Tab(2).Control(3)=   "TxtNo_Delivery"
      Tab(2).Control(4)=   "ChkDaily"
      Tab(2).ControlCount=   5
      TabCaption(3)   =   "„Ì“ ê—ÊÂÌ"
      TabPicture(3)   =   "frmFindTableDeliveryFich.frx":A6D5
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Timer1"
      Tab(3).Control(1)=   "timRefreshForm"
      Tab(3).Control(2)=   "txtNo_Table1"
      Tab(3).Control(3)=   "vsMultiFactors_Table"
      Tab(3).Control(4)=   "FWBtnPrint"
      Tab(3).Control(5)=   "CrystalReport1"
      Tab(3).Control(6)=   "FWButton1"
      Tab(3).Control(7)=   "lblMessage"
      Tab(3).Control(8)=   "Label10"
      Tab(3).Control(9)=   "Label9"
      Tab(3).ControlCount=   10
      TabCaption(4)   =   "„Ì“ Â«Ì  ”ÊÌÂ ‘œÂ"
      TabPicture(4)   =   "frmFindTableDeliveryFich.frx":A6F1
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "txtNo_Table_Tasvie"
      Tab(4).Control(1)=   "vsFactors_Table_Tasvie"
      Tab(4).Control(2)=   "Label17"
      Tab(4).Control(3)=   "Label16"
      Tab(4).ControlCount=   4
      Begin VB.TextBox txtNo_Table_Tasvie 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00808000&
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
         Left            =   -66240
         RightToLeft     =   -1  'True
         TabIndex        =   42
         Top             =   720
         Width           =   1155
      End
      Begin VB.PictureBox Picture1 
         Height          =   615
         Left            =   240
         ScaleHeight     =   555
         ScaleWidth      =   4875
         TabIndex        =   39
         Top             =   840
         Width           =   4935
         Begin VB.OptionButton optShowFich 
            Alignment       =   1  'Right Justify
            Caption         =   "‰„«Ì‘ Â„Â œ— „ÕœÊœÂ  «—ÌŒ"
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
            Index           =   1
            Left            =   0
            RightToLeft     =   -1  'True
            TabIndex        =   41
            Top             =   0
            Value           =   -1  'True
            Width           =   2955
         End
         Begin VB.OptionButton optShowFich 
            Alignment       =   1  'Right Justify
            Caption         =   "›ﬁÿ ›«ﬂ Ê—«‰ Œ«»Ì"
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
            Index           =   0
            Left            =   3000
            RightToLeft     =   -1  'True
            TabIndex        =   40
            Top             =   0
            Width           =   1875
         End
      End
      Begin VB.TextBox TxtNo_Temp 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0080C0FF&
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
         Left            =   9840
         RightToLeft     =   -1  'True
         TabIndex        =   35
         Top             =   600
         Width           =   1155
      End
      Begin VB.Timer Timer1 
         Interval        =   1000
         Left            =   -74280
         Top             =   600
      End
      Begin VB.Timer timRefreshForm 
         Enabled         =   0   'False
         Interval        =   30000
         Left            =   -74880
         Top             =   600
      End
      Begin VB.TextBox txtNo_Table1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00C0C000&
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
         Left            =   -67080
         RightToLeft     =   -1  'True
         TabIndex        =   24
         Top             =   840
         Width           =   1155
      End
      Begin VB.TextBox txtNo_Table 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00808000&
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
         Left            =   -66120
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Top             =   840
         Width           =   1155
      End
      Begin VB.CheckBox ChkDaily 
         Alignment       =   1  'Right Justify
         Caption         =   "«—”«·Ì Â«Ì «„—Ê“"
         BeginProperty Font 
            Name            =   "Nazanin"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   -74040
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   840
         Value           =   1  'Checked
         Width           =   2145
      End
      Begin VB.TextBox TxtNo_Delivery 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFC0C0&
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
         Left            =   -65760
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   840
         Width           =   1155
      End
      Begin VB.TextBox TxtNo_Fich 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0000C000&
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
         Left            =   9840
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   1080
         Width           =   1155
      End
      Begin VSFlex7LCtl.VSFlexGrid vsFactors_Delivery 
         Height          =   4965
         Left            =   -74400
         TabIndex        =   3
         Top             =   1365
         Width           =   11115
         _cx             =   19606
         _cy             =   8758
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
         SelectionMode   =   3
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   8
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   500
         RowHeightMax    =   500
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmFindTableDeliveryFich.frx":A70D
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
         ExplorerBar     =   5
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
      Begin FLWCtrls.FWProgressBar FWProgressBar1 
         Height          =   375
         Left            =   3240
         Top             =   5685
         Width           =   5895
         _ExtentX        =   10398
         _ExtentY        =   661
         Max             =   1000
         BorderStyle     =   10
      End
      Begin VSFlex7LCtl.VSFlexGrid vsFactors_Fich 
         Height          =   4005
         Left            =   240
         TabIndex        =   7
         Top             =   1560
         Width           =   11955
         _cx             =   21087
         _cy             =   7064
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
         ForeColor       =   -2147483640
         BackColorFixed  =   16761024
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483633
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
         SelectionMode   =   3
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   8
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   500
         RowHeightMax    =   500
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmFindTableDeliveryFich.frx":A7D1
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
         ExplorerBar     =   5
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
      Begin VSFlex7LCtl.VSFlexGrid vsFactors_Table 
         Height          =   4845
         Left            =   -74400
         TabIndex        =   11
         Top             =   1440
         Width           =   11355
         _cx             =   20029
         _cy             =   8546
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
         SelectionMode   =   3
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   7
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   500
         RowHeightMax    =   500
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmFindTableDeliveryFich.frx":A89E
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
         ExplorerBar     =   5
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
      Begin VSFlex7LCtl.VSFlexGrid vsMultiFactors_Table 
         Height          =   4725
         Left            =   -74640
         TabIndex        =   22
         Top             =   1440
         Width           =   10395
         _cx             =   18336
         _cy             =   8334
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
         SelectionMode   =   3
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   9
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   500
         RowHeightMax    =   500
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmFindTableDeliveryFich.frx":A962
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
         ExplorerBar     =   5
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
      Begin FLWCtrls.FWButton FWBtnPrint 
         Height          =   615
         Left            =   -64200
         TabIndex        =   26
         Top             =   5400
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   1085
         ButtonType      =   5
         Caption         =   "ç«Å"
         FontName        =   "Nazanin"
         FontBold        =   -1  'True
         FontSize        =   14.25
         Alignment       =   1
      End
      Begin Crystal.CrystalReport CrystalReport1 
         Left            =   -75000
         Top             =   0
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
      Begin FLWCtrls.FWButton FWButton1 
         Height          =   615
         Left            =   -64200
         TabIndex        =   27
         Top             =   4560
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   1085
         Caption         =   " ”ÊÌÂ"
         FontName        =   "Nazanin"
         FontBold        =   -1  'True
         FontSize        =   14.25
         Alignment       =   1
      End
      Begin FarDate1.FarDate txtDate1 
         Height          =   345
         Left            =   7560
         TabIndex        =   33
         Top             =   960
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   609
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Object.ToolTipText     =   "FarDate"
      End
      Begin FarDate1.FarDate txtDate2 
         Height          =   345
         Left            =   5280
         TabIndex        =   34
         Top             =   960
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   609
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Object.ToolTipText     =   "FarDate"
      End
      Begin VSFlex7LCtl.VSFlexGrid vsFactors_Table_Tasvie 
         Height          =   4845
         Left            =   -74520
         TabIndex        =   43
         Top             =   1320
         Width           =   11355
         _cx             =   20029
         _cy             =   8546
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
         SelectionMode   =   3
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   7
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   500
         RowHeightMax    =   500
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmFindTableDeliveryFich.frx":AA5C
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
         ExplorerBar     =   5
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
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         BackColor       =   &H0000C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "‰«„ „Ì“"
         BeginProperty Font 
            Name            =   "Nazanin"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   375
         Left            =   -65040
         RightToLeft     =   -1  'True
         TabIndex        =   45
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Ã” ÃÊÌ „Ì“Â«Ì  ”ÊÌÂ ‘œÂ"
         BeginProperty Font 
            Name            =   "Nazanin"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   495
         Left            =   -70680
         RightToLeft     =   -1  'True
         TabIndex        =   44
         Top             =   600
         Width           =   3255
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "„—ÃÊ⁄Ì"
         BeginProperty Font 
            Name            =   "Nazanin"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1080
         TabIndex        =   38
         Top             =   5640
         Width           =   615
      End
      Begin VB.Shape ShapeRecursive 
         BorderColor     =   &H00FFFF00&
         FillColor       =   &H00FFFF00&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   120
         Top             =   5640
         Width           =   615
      End
      Begin VB.Shape ShapeBalance 
         BorderColor     =   &H000000FF&
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   255
         Left            =   120
         Top             =   6000
         Width           =   615
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   " ”ÊÌÂ ‰‘œÂ"
         BeginProperty Font 
            Name            =   "Nazanin"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   840
         TabIndex        =   37
         Top             =   6000
         Width           =   975
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0000C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "‘„«—Â —Ê“«‰Â "
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
         Left            =   11160
         RightToLeft     =   -1  'True
         TabIndex        =   36
         Top             =   600
         Width           =   1095
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0000C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   " «"
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
         Left            =   7200
         RightToLeft     =   -1  'True
         TabIndex        =   32
         Top             =   960
         Width           =   255
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0000C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "«“"
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
         Left            =   9480
         RightToLeft     =   -1  'True
         TabIndex        =   31
         Top             =   960
         Width           =   255
      End
      Begin VB.Label lblMessage 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000001&
         BackStyle       =   0  'Transparent
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
         Height          =   705
         Left            =   -74760
         RightToLeft     =   -1  'True
         TabIndex        =   28
         Top             =   960
         Width           =   6495
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackColor       =   &H0000C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "‰«„ „Ì“"
         BeginProperty Font 
            Name            =   "Nazanin"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   375
         Left            =   -65880
         RightToLeft     =   -1  'True
         TabIndex        =   25
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Ã” ÃÊÌ „Ì“Â«Ì  ”ÊÌÂ ‰‘œÂ"
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
         Left            =   -71040
         RightToLeft     =   -1  'True
         TabIndex        =   23
         Top             =   600
         Width           =   3495
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Ã” ÃÊÌ „Ì“Â«Ì  ”ÊÌÂ ‰‘œÂ"
         BeginProperty Font 
            Name            =   "Nazanin"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   495
         Left            =   -70560
         RightToLeft     =   -1  'True
         TabIndex        =   19
         Top             =   720
         Width           =   3255
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Ã” ÃÊÌ ›«ò Ê—"
         BeginProperty Font 
            Name            =   "Nazanin"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   495
         Left            =   5760
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   480
         Width           =   3015
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H0000C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "‰«„ „Ì“"
         BeginProperty Font 
            Name            =   "Nazanin"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   375
         Left            =   -64920
         RightToLeft     =   -1  'True
         TabIndex        =   17
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H0000C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "«‘ —«ò"
         BeginProperty Font 
            Name            =   "Nazanin"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   375
         Left            =   -64320
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   840
         Width           =   975
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   " Ã” ÃÊÌ «—”«· ‰‘œÂ Â«"
         BeginProperty Font 
            Name            =   "Nazanin"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   495
         Left            =   -70800
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   600
         Width           =   3375
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "‰„«Ì‘ ›«ﬂ Ê—Â«"
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
         Left            =   1680
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "œ— Õ«·  «‰ Œ«» ‰„«Ì‘ Â„Â ' ”Ì” „ ﬂ·ÌÂ «ﬁ·«„ ›«ﬂ Ê—Â« Ì «Ì‰ ﬂ«—»— —« œ— ’Ê—  ÊÃÊœ œ” —”Ì  ‰„«Ì‘ „Ì œÂœ"
         BeginProperty Font 
            Name            =   "Nazanin"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   615
         Left            =   1320
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   6120
         Width           =   9135
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0000C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "‘„«—Â ”—Ì«·"
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
         Left            =   11040
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   1080
         Width           =   1215
      End
   End
   Begin RESIZEKITLibCtl.ResizeKit ResizeKit1 
      Height          =   480
      Left            =   0
      OleObjectBlob   =   "frmFindTableDeliveryFich.frx":AB20
      TabIndex        =   1
      Top             =   0
      Width           =   480
   End
   Begin VB.Label LblFindFactor 
      Alignment       =   2  'Center
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
      ForeColor       =   &H000000C0&
      Height          =   615
      Left            =   720
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   720
      Width           =   5415
   End
End
Attribute VB_Name = "frmFindTableDeliveryFich"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim clsDate As New clsDate
Dim i As Long
Dim FactorType As EnumFactorType
Dim Parameter() As Parameter
Dim strTableNoDetailString As String
Private Const indexColTableNo As Integer = 8

Private Sub CancelButton_Click()
    mvarcode = 0
    Unload Me
End Sub

Private Sub ChkDaily_Click()
    FillvsFactors_Delivery
End Sub

Private Sub Form_Activate()
    
    formloadFlag = True
    If clsStation.SearchType = 0 Then
        If clsStation.TemporaryNo = True Then
            TxtNo_Temp.SetFocus
        Else
            TxtNo_Fich.SetFocus
        End If
    ElseIf clsStation.SearchType = 1 Then
        txtNo_Table.SetFocus
    ElseIf clsStation.SearchType = 2 Then
        TxtNo_Delivery.SetFocus
    ElseIf clsStation.SearchType = 3 Then
        txtNo_Table1.SetFocus
    End If

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   KeyDown KeyCode, Shift
End Sub

Private Sub Form_Load()

    
    CenterCenterinSecondScreen Me
    
    FactorType = EnumFactorType.Invoice
    mvarcode = 0
   
    FWProgressBar1.Visible = False
    ChangeLanguage
'    If clsStation.SearchFichDefault = True Then
'        optShowFich(0).Value = True
'    Else
        optShowFich(0).Value = False
'        If clsStation.SearchType = 0 Then
'            optShowFich_Click 1
'        End If
'    End If
    frmInvoice.FindFlag = False
    formloadFlag = False
    
    txtDate1.Text = "13" & mvarDate
    txtDate2.Text = txtDate1.Text
    
    FillvsFactors_Delivery
    FillvsFactors_Table
    FillvsFactors_Table_Tasvie
    FillvsMultiFactors_Table
    
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


    If clsStation.SearchType = 0 Then
        SSTab1.Tab = 0
    ElseIf clsStation.SearchType = 1 Then
        SSTab1.Tab = 1
    ElseIf clsStation.SearchType = 2 Then
        SSTab1.Tab = 2
    ElseIf clsStation.SearchType = 3 Then
        SSTab1.Tab = 3
    End If
'    If clsStation.TemporaryNo = True Then
'        TxtNo_Temp.Enabled = True
'        TxtNo_Fich.Enabled = False
'    Else
'        TxtNo_Temp.Enabled = False
'        TxtNo_Fich.Enabled = True
'    End If
    optShowFich_Click 1
    
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call PresetScreenSaver
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call PresetScreenSaver
End Sub
Public Sub ChangeLanguage()
''''If clsStation.Language = English Then
''''    CancelButton.Caption = "Cancel"
''''    ChkDaily.Caption = "Today delivered"
''''    FWBtnPrint.Caption = "print"
''''    FWButton1.Caption = "payoff"
''''    Label1.Caption = "Table's name"
''''    Label10.Caption = "Table's name"
''''    Label3.Caption = "Number of digits of invoice"
''''    Label4.Caption = "Invoice number"
''''    Label5.Caption = "Customer ID"
''''    Label6.Caption = "Searching through not paied tables"
''''    Label7.Caption = "Finding invoice"
''''    Label8.Caption = "Finding not sent"
''''    Label9.Caption = "Searching through not paied tables"
''''    OKButton.Caption = "Accept"
''''    optShowFich(0).Caption = "Regular"
''''    optShowFich(1).Caption = "Last three digits"
''''    SSTab1.TabCaption(0) = "Invoice"
''''    SSTab1.TabCaption(1) = "Table"
''''    SSTab1.TabCaption(2) = "Transmited"
''''    SSTab1.TabCaption(3) = "Group table"
''''    With vsFactors_Delivery
''''            .TextMatrix(0, 0) = "Invoice ID"
''''            .TextMatrix(0, 1) = ""
''''            .TextMatrix(0, 2) = "Serial"
''''            .TextMatrix(0, 3) = "Customer ID"
''''            .TextMatrix(0, 4) = "Customer"
''''            .TextMatrix(0, 5) = "Cost"
''''            .TextMatrix(0, 6) = "Time"
''''            .TextMatrix(0, 7) = "Date"
''''            .TextMatrix(0, 8) = "Address"
''''   End With
''''   With vsFactors_Table
''''            .RightToLeft = False
''''            .TextMatrix(0, 0) = "Row"
''''            .TextMatrix(0, 2) = "Table"
''''            .TextMatrix(0, 3) = "Garsoon"
''''            .TextMatrix(0, 4) = "Serial"
''''            .TextMatrix(0, 5) = "Cost"
''''            .TextMatrix(0, 6) = "Time"
''''   End With
''''   With vsFactorDetail
''''            .TextMatrix(0, 0) = "Row"
''''            .TextMatrix(0, 1) = "Quantity"
''''            .TextMatrix(0, 2) = "Goods Name"
''''            .TextMatrix(0, 3) = "Fee"
''''            .TextMatrix(0, 4) = "Sum"
''''   End With
''''   With vsMultiFactors_Table
''''            .TextMatrix(0, 0) = "Row"
''''            .TextMatrix(0, 2) = "Select"
''''            .TextMatrix(0, 3) = "Table"
''''            .TextMatrix(0, 4) = "Garsoon"
''''            .TextMatrix(0, 5) = "Serial"
''''            .TextMatrix(0, 6) = "Cost"
''''            .TextMatrix(0, 7) = "Time"
''''   End With
''''   With vsFactors_Fich
''''            .TextMatrix(0, 0) = "Row"
''''            .TextMatrix(0, 1) = "SerialNo"
''''            .TextMatrix(0, 2) = "invoice ID"
''''            .TextMatrix(0, 3) = "Customer"
''''            .TextMatrix(0, 4) = "Date"
''''            .TextMatrix(0, 5) = "Time"
''''            .TextMatrix(0, 6) = "Price"
''''            .TextMatrix(0, 7) = "No3"
''''            .TextMatrix(0, 8) = "User"
''''            .TextMatrix(0, 9) = "Balance"
''''            .TextMatrix(0, 10) = "Void"
''''            .TextMatrix(0, 11) = "Service"
''''            .TextMatrix(0, 12) = "Shipment"
''''            .TextMatrix(0, 13) = "Discount"
''''            .TextMatrix(0, 14) = "Description"
''''   End With
''''Else
    CancelButton.Caption = "«‰’—«›"
    ChkDaily.Caption = "«—”«·Ì «„—Ê“"
    FWBtnPrint.Caption = "ç«Å"
    FWButton1.Caption = " ”ÊÌÂ"
    Label1.Caption = "‰«„ „Ì“"
    Label10.Caption = "‰«„ „Ì“"
    Label3.Caption = " ⁄œ«œ «—ﬁ«„ ›«ﬂ Ê—"
    Label4.Caption = "”—Ì«· ›Ì‘"
    Label5.Caption = "«‘ —«ﬂ"
    Label6.Caption = "Ã” ÃÊÌ „Ì“Â«Ì  ”ÊÌÂ ‰‘œÂ"
    Label7.Caption = "Ã” ÃÊÌ ›Ì‘"
    Label8.Caption = " Ã” ÃÊÌ «—”«· ‰‘œÂ Â«"
    Label9.Caption = "Ã” ÃÊÌ „Ì“Â«Ì  ”ÊÌÂ ‰‘œÂ"
    OKButton.Caption = "ﬁ»Ê·"
    optShowFich(0).Caption = "›ﬁÿ ›«ﬂ Ê— «‰ Œ«»Ì"
    optShowFich(1).Caption = "‰„«Ì‘ Â„Â œ— „ÕœÊœÂ  «—ÌŒ"
    SSTab1.TabCaption(0) = "›Ì‘"
    SSTab1.TabCaption(1) = "„Ì“"
    SSTab1.TabCaption(2) = "«—”«·Ì Â«"
    SSTab1.TabCaption(3) = "„Ì“ ê—ÊÂÌ"
    With vsFactors_Fich
        .Cols = 16
        .TextMatrix(0, 0) = "—œÌ›"
        .TextMatrix(0, 1) = "ﬂœ"
        .TextMatrix(0, 2) = "‘„«—Â"
        .TextMatrix(0, 3) = "„‘ —Ì"
        .TextMatrix(0, 4) = " «—ÌŒ"
        .TextMatrix(0, 5) = "“„«‰"
        .TextMatrix(0, 6) = "„»·€"
        .TextMatrix(0, 7) = "”—Ì«·"
        .TextMatrix(0, 8) = "ò«—»—"
        .TextMatrix(0, 9) = " ”ÊÌÂ"
        .TextMatrix(0, 10) = "„—ÃÊ⁄Ì"
        .TextMatrix(0, 11) = "”—ÊÌ”"
        .TextMatrix(0, 12) = "ò—«ÌÂ Õ„·"
        .TextMatrix(0, 13) = " Œ›Ì›"
        .TextMatrix(0, 14) = "‘Ì› "
        .TextMatrix(0, 15) = " Ê÷ÌÕ« "
         For i = 0 To .Cols - 1
             .ColWidth(i) = Val(GetSetting(strMainKey, Me.Name & "vsFactors_Fich", "Col" & i))
            If .ColWidth(i) = 0 Then
                .ColWidth(i) = 1000       'Row
            End If
         Next i
        .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter
        .ColAlignment(-1) = flexAlignCenterCenter
        .ColAlignment(3) = flexAlignRightCenter
        .ColAlignment(8) = flexAlignRightCenter
        .ColAlignment(15) = flexAlignRightCenter
        .ColDataType(9) = flexDTBoolean
        .ColDataType(10) = flexDTBoolean
'        .ColHidden(7) = True
   End With
   With vsMultiFactors_Table
        .Cols = 12
        .TextMatrix(0, 0) = "—œÌ›"
        .TextMatrix(0, 2) = "«‰ Œ«»"
        .TextMatrix(0, 3) = "„Ì“"
        .TextMatrix(0, 4) = "ê«—”Ê‰"
        .TextMatrix(0, 5) = "‘„«—Â"
        .TextMatrix(0, 6) = "„»·€"
        .TextMatrix(0, 7) = "”«⁄ "
        .TextMatrix(0, 8) = "ﬂœ „Ì“"
        .TextMatrix(0, 9) = "”—Ì«·"
        .TextMatrix(0, 10) = "‘Ì› "
        .TextMatrix(0, 11) = "‰›—« "
        .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter
        .ColAlignment(-1) = flexAlignCenterCenter
         For i = 0 To .Cols - 1
             .ColWidth(i) = Val(GetSetting(strMainKey, Me.Name & "vsMultiFactors_Table", "Col" & i))
            If .ColWidth(i) = 0 Then
                .ColWidth(i) = 1000       'Row
            End If
         Next i
   End With
   With vsFactorDetail
        .Cols = 5
        .TextMatrix(0, 0) = "—œÌ›"
        .TextMatrix(0, 1) = " ⁄œ«œ"
        .TextMatrix(0, 2) = "‰«„ ﬂ«·«"
        .TextMatrix(0, 3) = "›Ì"
        .TextMatrix(0, 4) = "Ã„⁄"
         For i = 0 To .Cols - 1
             .ColWidth(i) = Val(GetSetting(strMainKey, Me.Name & "vsFactorDetail", "Col" & i))
            If .ColWidth(i) = 0 Then
                .ColWidth(i) = 1000       'Row
            End If
         Next i
        .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter
        .ColAlignment(2) = flexAlignRightCenter
   End With
   With vsFactors_Table
        .Cols = 10
        .RightToLeft = True
        .TextMatrix(0, 0) = "—œÌ›"
        .TextMatrix(0, 2) = "„Ì“"
        .TextMatrix(0, 3) = "ê«—”Ê‰"
        .TextMatrix(0, 4) = "‘„«—Â"
        .TextMatrix(0, 5) = "„»·€"
        .TextMatrix(0, 6) = "”«⁄ "
        .TextMatrix(0, 7) = "”—Ì«·"
        .TextMatrix(0, 8) = "‘Ì› "
        .TextMatrix(0, 9) = "‰›—« "
        .ColAlignment(-1) = flexAlignCenterCenter
         For i = 0 To .Cols - 1
             .ColWidth(i) = Val(GetSetting(strMainKey, Me.Name & "vsFactors_Table", "Col" & i))
            If .ColWidth(i) = 0 Then
                .ColWidth(i) = 1000       'Row
            End If
         Next i
    End With
   With vsFactors_Table_Tasvie
        .Cols = 10
        .RightToLeft = True
        .TextMatrix(0, 0) = "—œÌ›"
        .TextMatrix(0, 2) = "„Ì“"
        .TextMatrix(0, 3) = "ê«—”Ê‰"
        .TextMatrix(0, 4) = "‘„«—Â"
        .TextMatrix(0, 5) = "„»·€"
        .TextMatrix(0, 6) = "”«⁄ "
        .TextMatrix(0, 7) = "”—Ì«·"
        .TextMatrix(0, 8) = "‘Ì› "
        .TextMatrix(0, 9) = "‰›—« "
        .ColAlignment(-1) = flexAlignCenterCenter
         For i = 0 To .Cols - 1
             .ColWidth(i) = Val(GetSetting(strMainKey, Me.Name & "vsFactors_Table_Tasvie", "Col" & i))
            If .ColWidth(i) = 0 Then
                .ColWidth(i) = 1000       'Row
            End If
         Next i
    End With
    With vsFactors_Delivery
        .Rows = 1
        .Cols = 11
        .TextMatrix(0, 0) = "—œÌ›"
        .TextMatrix(0, 1) = "òœ ›Ì‘"
        .TextMatrix(0, 2) = "‘„«—Â"
        .TextMatrix(0, 3) = "«‘ —«ò"
        .TextMatrix(0, 4) = "„‘ —Ì"
        .TextMatrix(0, 5) = "„»·€"
        .TextMatrix(0, 6) = "”«⁄ "
        .TextMatrix(0, 7) = " «—ÌŒ"
        .TextMatrix(0, 8) = "‘Ì› "
        .TextMatrix(0, 9) = "”—Ì«·"
        .TextMatrix(0, 10) = "¬œ—”"
         For i = 0 To .Cols - 1
             .ColWidth(i) = Val(GetSetting(strMainKey, Me.Name & "vsFactors_Delivery", "Col" & i))
            If .ColWidth(i) = 0 Then
                .ColWidth(i) = 1000       'Row
            End If
         Next i
        .ColAlignment(-1) = flexAlignCenterCenter
        .ColAlignment(4) = flexAlignRightCenter
        .ColAlignment(10) = flexAlignRightCenter
        .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter
        .ColHidden(1) = True
        .AutoSearch = flexSearchFromCursor
        End With

''''End If
    With vsReceived
        .Cols = 5
        .TextMatrix(0, 0) = "—œÌ›"
        .TextMatrix(0, 1) = "œ—Ì«› Ì"
        .TextMatrix(0, 2) = " «—ÌŒ"
        .TextMatrix(0, 3) = "“„«‰"
        .TextMatrix(0, 4) = "‰Ê⁄"
         For i = 0 To .Cols - 1
             .ColWidth(i) = Val(GetSetting(strMainKey, Me.Name & "VSReceived", "Col" & i))
            If .ColWidth(i) = 0 Then
                .ColWidth(i) = 1000       'Row
            End If
         Next i
    End With
    With VSHistory
        .Cols = 4
        .TextMatrix(0, 0) = "—œÌ›"
        .TextMatrix(0, 1) = "—ÊÌœ«œ À» Ì"
        .TextMatrix(0, 2) = " «—ÌŒ"
        .TextMatrix(0, 3) = "“„«‰"
         For i = 0 To .Cols - 1
             .ColWidth(i) = Val(GetSetting(strMainKey, Me.Name & "VSHistory", "Col" & i))
            If .ColWidth(i) = 0 Then
                .ColWidth(i) = 1000       'Row
            End If
         Next i
    End With
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If PosConnection.State = adStateOpen Then PosConnection.Close: Set PosConnection = Nothing
    SaveSetting strMainKey, Me.Name, "Left", Me.Left
    SaveSetting strMainKey, Me.Name, "Top", Me.Top
End Sub

Private Sub FWBtnPrint_Click()
    On Error GoTo Err_Handler
    
    Dim i As Integer
    Dim s As String
    
    s = ""
     With vsMultiFactors_Table
         For i = 1 To .Rows - 1
             If Val(.TextMatrix(i, 2)) = -1 Then
                  s = s & .TextMatrix(i, 1) & ","
             End If
         Next i
     End With
    If s = "" Then
        frmDisMsg.lblMessage = "ÂÌç „Ê—œÌ «‰ Œ«» ‰‘œÂ «” "
        frmDisMsg.Timer1.Enabled = True
        frmDisMsg.Show vbModal
        Exit Sub
    End If
    
    frmInput.fwlblInput.Caption = "”«Ì“ ﬂ«€–"
    frmInput.OptionLevel(0).Caption = "ﬂ«€– A5"
    frmInput.OptionLevel(1).Caption = "ﬂ«€– Recipt"
    frmInput.txtInput.Visible = False
    frmInput.Picture1.Visible = True
    frmInput.Show vbModal
    ReDim Parameter(0) As Parameter
    Parameter(0) = GenerateInputParameter("@strSelectedFactors", adVarWChar, 4000, s)
   ''  RunParametricStoredProcedure "Get_Selected_Table", Parameter
    
    
    If mvarInput = "" Then
        Exit Sub
    ElseIf mvarInput = 0 Then
       CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\A5\Invoicefich_MultiTable_A5.rpt"
    ElseIf mvarInput = 1 Then
       CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\Invoicefich_MultiTable.rpt"
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
   'CrystalReport1.ReportTitle = " ê—œ‘ ò«·« œ— «‰»«— "
    CrystalReport1.Destination = crptToWindow 'crptToPrinter '
    Dim intIndex As Integer
   
    For intIndex = 0 To UBound(Parameter) - LBound(Parameter)
        CrystalReport1.ParameterFields(intIndex) = CStr(Parameter(intIndex).Name) & ";" & CStr(Parameter(intIndex).Value) & ";" & "True"
    Next intIndex
  
    CrystalReport1.RetrieveDataFiles
    ODBCSetting clsArya.ServerName, clsArya.DbName
    CrystalReport1.Connect = CrystallConnection
        CrystalReport1.PageZoom (100)
    CrystalReport1.Action = 1
'    If Screen.Width > 12000 Then
'    Else
'        CrystalReport1.PageZoom (75)
'    End If
    
    ReDim Parameter(2) As Parameter
    Parameter(0) = GenerateInputParameter("@strSelectedFactors", adVarWChar, 4000, s)
    Parameter(1) = GenerateInputParameter("@UID", adInteger, 4, mvarCurUserNo)
    Parameter(2) = GenerateInputParameter("@ActionCode", adInteger, 4, EnumActionLog.InvoicePrint)
    RunParametricStoredProcedure "InsertHistory_Batch", Parameter

    Exit Sub
    
Err_Handler:
    ShowErrorMessage
    LogSaveNew "frmFindTableDeliveryFich => ", err.Description, err.Number, err.Source, "FWBtnPrint_Click"
End Sub

Private Sub FWButton1_Click()
    On Error GoTo Err_Handler
    
    Dim i As Integer
    Dim s, S2 As String
    Dim strPayk As String
    Dim InvoiceNoString As String
    InvoiceNoString = ""
    s = ""
    S2 = ""
    
     With vsMultiFactors_Table
         For i = 1 To .Rows - 1
             If Val(.TextMatrix(i, 2)) = -1 Then
                  s = s & .TextMatrix(i, 1) & ","
                  S2 = S2 & .TextMatrix(i, indexColTableNo) & ","
                  InvoiceNoString = InvoiceNoString & .TextMatrix(i, 9) & ","
             End If
         Next i
''     End With
''    If i = "" Then
''        frmDisMsg.lblMessage = "ÂÌç „Ê—œÌ «‰ Œ«» ‰‘œÂ «” "
''        frmDisMsg.Timer1.Enabled = True
''        frmDisMsg.Show vbModal
''        Exit Sub
''    End If
    
    
    
''    Dim i As Integer
''    Dim s As String
''    Dim strPayk As String
    
''    s = ""
''    With vsMultiFactors_Table
    
''        For i = 1 To .Rows - 1
''            If Val(.TextMatrix(i, 1)) = -1 Then
''                s = s & .TextMatrix(i, 0) & ","
''            End If
''        Next i
   
        If s = "" Then
        frmDisMsg.lblMessage = "ÂÌç „Ê—œÌ «‰ Œ«» ‰‘œÂ «” "
        frmDisMsg.Timer1.Enabled = True
        frmDisMsg.Show vbModal
        Exit Sub
        End If
        
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
        S2 = Left(S2, Len(S2) - 1)
        InvoiceNoString = Left(InvoiceNoString, Len(InvoiceNoString) - 1)
        ReDim Parameter(2) As Parameter
        
            ReDim Parameter(2) As Parameter
            Parameter(0) = GenerateInputParameter("@strSelectedFactors", adVarWChar, 4000, s)
            Parameter(1) = GenerateInputParameter("@strSelectedTables", adVarWChar, 4000, S2)
            Parameter(2) = GenerateInputParameter("@UID", adInteger, 4, mvarCurUserNo)
            RunParametricStoredProcedure "PayFactors_Table", Parameter
        
        If mdifrm.ClsActionLog.LogPayCustomerFactor Then
            ReDim Parameter(2) As Parameter
            Parameter(0) = GenerateInputParameter("@strSelectedFactors", adVarWChar, 4000, s)
            Parameter(1) = GenerateInputParameter("@UID", adInteger, 4, mvarCurUserNo)
            Parameter(2) = GenerateInputParameter("@ActionCode", adInteger, 4, EnumActionLog.PayCustomerFactor)
            RunParametricStoredProcedure "InsertHistory_Batch", Parameter
            
        End If
        
        If InStr(1, s, ",") > 0 Then
             lblMessage = "›«ò Ê—Â«Ì ‘„«—Â" & InvoiceNoString & " Å—œ«Œ  ‘œ "
        Else
             lblMessage = "›«ò Ê— ‘„«—Â" & InvoiceNoString & " Å—œ«Œ  ‘œ "
        End If
        
        FillvsMultiFactors_Table

        Timer1.Interval = 3000
        Timer1.Enabled = True
            
    End With
    
    Exit Sub
    
Err_Handler:
    ShowErrorMessage
    LogSaveNew "frmFindTableDeliveryFich => ", err.Description, err.Number, err.Source, "FWButton1_Click"
End Sub

Private Sub OKButton_Click()
    On Error GoTo Err_Handler
    
    If SSTab1.Tab = 0 Then
        If vsFactors_Fich.Row > 0 Then
            mvarcode = vsFactors_Fich.TextMatrix(vsFactors_Fich.Row, 7)
            frmInvoice.FindFlag = True
        Else
            mvarcode = 0
        End If
    ElseIf SSTab1.Tab = 1 Then
        If vsFactors_Table.Row > 0 Then
            mvarcode = vsFactors_Table.TextMatrix(vsFactors_Table.Row, 7)
            frmInvoice.FindFlag = True
        Else
            mvarcode = 0
        End If
    ElseIf SSTab1.Tab = 2 Then
        If vsFactors_Delivery.Row > 0 Then
            mvarcode = vsFactors_Delivery.TextMatrix(vsFactors_Delivery.Row, 9)
            frmInvoice.FindFlag = True
        Else
            mvarcode = 0
        End If
    ElseIf SSTab1.Tab = 4 Then
        If vsFactors_Table_Tasvie.Row > 0 Then
            mvarcode = vsFactors_Table_Tasvie.TextMatrix(vsFactors_Table_Tasvie.Row, 7)
            frmInvoice.FindFlag = True
        Else
            mvarcode = 0
        End If
    End If
    Unload Me
    
    Exit Sub
    
Err_Handler:
    ShowErrorMessage
    LogSaveNew "frmFindTableDeliveryFich => ", err.Description, err.Number, err.Source, "OKButton_Click"
End Sub

Sub ClearDataFlexGrid()

'''    With vsFactors_Table
'''        .Rows = 1
'''
'''    End With
    With vsFactors_Fich
        .Rows = 1
               
    End With
''    With vsFactors_Delivery
''        .Rows = 1
''
''    End With
End Sub

Private Sub ResizeKit1_ExitResize(ByVal XScale As Double, ByVal YScale As Double)

    If formloadFlag = True Then
        SaveSetting strMainKey, Me.Name, "Height", Me.Height
        SaveSetting strMainKey, Me.Name, "Width", Me.Width
    End If

End Sub
Private Sub optShowFich_Click(index As Integer)
    If optShowFich(1).Value = 0 Then
        ClearDataFlexGrid
        vsFactors_Fich.Row = 0
        txtDate1.Enabled = False
        txtDate2.Enabled = False
        If clsStation.TemporaryNo = True Then TxtNo_Temp.SetFocus Else TxtNo_Fich.SetFocus
        
    '    vsFactors_Fich.ShowCell 1, 0
    '    vsFactors_Fich.Sort = flexSortGenericDescending
    Else
        txtDate1.Enabled = True
        txtDate2.Enabled = True
        FillvsFactors_Fich
        vsFactors_Fich.Row = 0
'        If vsFactors_Fich.Rows > 1 Then
'           vsFactors_Fich.ShowCell 1, 0
'           vsFactors_Fich.Sort = flexSortGenericDescending
'        End If
   End If
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
    On Error GoTo ErrHandler
    FillvsFactorDetail
    FillvsReceived
    FillvsHistory
    If formloadFlag = False Then
        formloadFlag = True
        Exit Sub
    End If
    If SSTab1.Tab = 0 Then
        FillvsFactors_Fich
        If clsStation.TemporaryNo = True Then
            TxtNo_Temp.SetFocus
        Else
            TxtNo_Fich.SetFocus
        End If
    ElseIf SSTab1.Tab = 1 Then
        FillvsFactors_Table
        txtNo_Table.SetFocus
    ElseIf SSTab1.Tab = 2 Then
        TxtNo_Delivery.SetFocus
    ElseIf SSTab1.Tab = 3 Then
        FillvsMultiFactors_Table
        txtNo_Table1.SetFocus
    ElseIf SSTab1.Tab = 4 Then
        FillvsFactors_Table_Tasvie
        txtNo_Table_Tasvie.SetFocus
    End If
Exit Sub
    
ErrHandler:
    ShowErrorMessage
    LogSaveNew "frmFindTableDeliveryFich => ", err.Description, err.Number, err.Source, "SStab1_Click"
    
End Sub


Private Sub SSTab1_KeyDown(KeyCode As Integer, Shift As Integer)
    KeyDown KeyCode, Shift
End Sub

Private Sub txtDate1_Change()
    If formloadFlag = False Then Exit Sub
    If Len(txtDate1.Text) = 10 Then
        FillvsFactors_Fich
        vsFactors_Fich.Row = 0
'        If vsFactors_Fich.Rows > 1 Then
'            vsFactors_Fich.ShowCell 1, 0
'            vsFactors_Fich.Sort = flexSortGenericDescending
'        End If
    End If
End Sub

Private Sub txtDate2_Change()
    If formloadFlag = False Then Exit Sub
    If Len(txtDate2.Text) = 10 Then
        FillvsFactors_Fich
        vsFactors_Fich.Row = 0
'        If vsFactors_Fich.Rows > 1 Then
'            vsFactors_Fich.ShowCell 1, 0
'            vsFactors_Fich.Sort = flexSortGenericDescending
'        End If
    End If
End Sub

Private Sub txtNo_Table_Change()
    i = -1
    If Len(txtNo_Table.Text) <> 0 Then
       i = vsFactors_Table.FindRow(txtNo_Table.Text, 1, 2, True, True)
    End If
    If i > 0 Then
        vsFactors_Table.Row = i
        vsFactors_Table.ShowCell i, 0
        LblFindFactor.Caption = ""
        'vsFactors_Table.SetFocus
    Else
        vsFactors_Table.Row = 0
        vsFactors_Table.ShowCell 0, 0
        If Val(txtNo_Table.Text) > 0 Then
           LblFindFactor.Caption = "›—Ê‘ œ— „Ì“ " & Val(txtNo_Table.Text) & " «‰Ã«„ ‰‘œÂ"
         Else
            LblFindFactor.Caption = "‘„«—Â „Ì“ —« Ê«—œ ﬂ‰Ìœ  "
         End If
    
    End If
    FillvsFactorDetail
    FillvsReceived
    FillvsHistory
End Sub

Private Sub txtNo_Table_Tasvie_Change()
    i = -1
    If Len(txtNo_Table_Tasvie.Text) <> 0 Then
       i = vsFactors_Table_Tasvie.FindRow(txtNo_Table_Tasvie.Text, 1, 2, True, True)
    End If
    If i > 0 Then
        vsFactors_Table_Tasvie.Row = i
        vsFactors_Table_Tasvie.ShowCell i, 0
        LblFindFactor.Caption = ""
        'vsFactors_Table.SetFocus
    Else
        vsFactors_Table_Tasvie.Row = 0
        vsFactors_Table_Tasvie.ShowCell 0, 0
        If Val(txtNo_Table_Tasvie.Text) > 0 Then
           LblFindFactor.Caption = "›—Ê‘ œ— „Ì“ " & Val(txtNo_Table_Tasvie.Text) & " «‰Ã«„ ‰‘œÂ"
         Else
            LblFindFactor.Caption = "‘„«—Â „Ì“ —« Ê«—œ ﬂ‰Ìœ  "
         End If
    
    End If
    FillvsFactorDetail
    FillvsReceived
    FillvsHistory
End Sub

Private Sub txtNo_Delivery_Change()
    i = -1
    If Len(TxtNo_Delivery.Text) <> 0 Then
       i = vsFactors_Delivery.FindRow(TxtNo_Delivery.Text, 1, 3, True, True)
    End If
    If i > 0 Then
        vsFactors_Delivery.Row = i
        vsFactors_Delivery.ShowCell i, 0
        LblFindFactor.Caption = ""
      '  vsFactors_Delivery.SetFocus
        vsFactors_Delivery.TopRow = i
    Else
        vsFactors_Delivery.Row = 0
        vsFactors_Delivery.ShowCell 0, 0
        If Val(TxtNo_Delivery.Text) > 0 Then
           LblFindFactor.Caption = "»—«Ì «‘ —«ò  " & Val(TxtNo_Delivery.Text) & " ›«ò Ê— «—”«· ‰‘œÂ ‰œ«—Ì„"
         Else
            LblFindFactor.Caption = "‘„«—Â «‘ —«ò —« Ê«—œ ﬂ‰Ìœ  "
         End If
    
    End If
    FillvsFactorDetail
    FillvsReceived
    FillvsHistory
End Sub

Private Sub txtNo_Fich_Change()
    i = -1
    If optShowFich(0).Value = True Then
  '      i = vsFactors_Fich.FindRow(txtNo_Fich.Text, 1, 2, True, True)
        If Val(TxtNo_Fich.Text) > 0 Then
           Define_Factor
        Else
            vsFactors_Fich.Rows = 1
         
        End If
    Else
'        If vsFactors_Fich.Rows >= 1000 Then
'            If Len(TxtNo_Fich.Text) = 3 Then
'               i = vsFactors_Fich.FindRow(TxtNo_Fich.Text, 1, 7, True, True)
'            End If
'        Else
            i = vsFactors_Fich.FindRow(TxtNo_Fich.Text, 1, 7, True, True)
'        End If
    End If
    If i > 0 Then
        vsFactors_Fich.Row = i
        vsFactors_Fich.ShowCell i, 0
        LblFindFactor.Caption = ""
    Else
        vsFactors_Fich.Row = 0
        vsFactors_Fich.ShowCell 0, 0
        If Val(TxtNo_Fich.Text) > 0 Then
           LblFindFactor.Caption = " ›«ﬂ Ê— " & Val(TxtNo_Fich.Text) & "  œ— ”Ì” „  ⁄—Ì› ‰‘œÂ"
         Else
            LblFindFactor.Caption = "‘„«—Â ›«ﬂ Ê— —« Ê«—œ ﬂ‰Ìœ  "
         End If
    
    End If
    FillvsFactorDetail
    FillvsReceived
    FillvsHistory
End Sub

Private Sub txtNo_Table_GotFocus()

    vsFactors_Table.Row = 0
    vsFactors_Table.Select vsFactors_Table.Row, 2
    vsFactors_Table.Sort = flexSortNumericAscending
  '  vsFactors_Table.Sort = flexSortGenericDescending
    LblFindFactor.Caption = "‘„«—Â „Ì“ —« Ê«—œ ﬂ‰Ìœ  "
    For i = 1 To vsFactors_Table.Rows - 1
        vsFactors_Table.TextMatrix(i, 0) = i
    Next
    
End Sub

Private Sub txtNo_Table_Tasvie_GotFocus()

    vsFactors_Table_Tasvie.Row = 0
    vsFactors_Table_Tasvie.Select vsFactors_Table_Tasvie.Row, 2
    vsFactors_Table_Tasvie.Sort = flexSortNumericAscending
  '  vsFactors_Table.Sort = flexSortGenericDescending
    LblFindFactor.Caption = "‘„«—Â „Ì“ —« Ê«—œ ﬂ‰Ìœ  "
    For i = 1 To vsFactors_Table_Tasvie.Rows - 1
        vsFactors_Table_Tasvie.TextMatrix(i, 0) = i
    Next
    
End Sub
Private Sub txtNo_Delivery_GotFocus()

    vsFactors_Delivery.Row = 0
    vsFactors_Delivery.Select vsFactors_Delivery.Row, 3
    vsFactors_Delivery.Sort = flexSortGenericAscending
  '  vsFactors_Delivery.Sort = flexSortGenericDescending
    LblFindFactor.Caption = "‘„«—Â «‘ —«ò —« Ê«—œ ﬂ‰Ìœ  "
    For i = 1 To vsFactors_Delivery.Rows - 1
        vsFactors_Delivery.TextMatrix(i, 0) = i
    Next
    
End Sub
Private Sub txtNo_Fich_GotFocus()

    vsFactors_Fich.Row = 0
    vsFactors_Fich.Select vsFactors_Fich.Row, 7
'    vsFactors_Fich.Sort = flexSortGenericAscending
    vsFactors_Fich.Sort = flexSortGenericDescending
    LblFindFactor.Caption = "‘„«—Â ›«ﬂ Ê— —« Ê«—œ ﬂ‰Ìœ  "
    
End Sub

Private Sub FillvsFactors_Table()
    On Error GoTo Err_Handler
    
    Dim Rst As New ADODB.Recordset
    Dim CountTableFich As Integer
    CountTableFich = 0
    ReDim Parameter(2) As Parameter
    Parameter(0) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
    Parameter(1) = GenerateInputParameter("@nvcDate", adWChar, 8, mvarDate)
    Parameter(2) = GenerateInputParameter("@Branch", adInteger, 4, CurrentBranch)
    Set Rst = RunParametricStoredProcedure2Rec("Get_Factors_Tables", Parameter)

    With vsFactors_Table
        .Rows = 1
        i = 0
        While Rst.EOF <> True
             If Rst!FacPayment = False Then
                If Rst!PartitionID = clsStation.PartitionID Or clsStation.OtherPartition = False Then
                    .Rows = .Rows + 1
                    CountTableFich = CountTableFich + 1
                    i = .Rows - 1
                    .TextMatrix(i, 0) = i
                    .TextMatrix(i, 1) = Rst!intSerialNo
                    .TextMatrix(i, 2) = Rst!TableName
                    .TextMatrix(i, 3) = IIf(IsNull(Rst!FullName), "", Rst!FullName)
                    .TextMatrix(i, 4) = Rst!tempNo
                    .TextMatrix(i, 5) = Rst!sumPrice
                    .TextMatrix(i, 6) = Rst!time
                    .TextMatrix(i, 7) = Rst!No
                    .TextMatrix(i, 8) = Rst!ShiftDescription
                    .TextMatrix(i, 9) = IIf(IsNull(Rst!GuestNo), "", Rst!GuestNo)
                End If
            End If
            Rst.MoveNext
         
        Wend
    If CountTableFich > 0 Then
      .Row = 1
    End If
    End With

    vsFactors_Table.Cell(flexcpAlignment, 0, 0, 0, 6) = flexAlignCenterCenter
    vsFactors_Table.ColAlignment(-1) = flexAlignCenterCenter
    
    Set Rst = Nothing

    Exit Sub
    
Err_Handler:
    ShowErrorMessage
    LogSaveNew "frmFindTableDeliveryFich => ", err.Description, err.Number, err.Source, "FillVsFactors_Table"
End Sub
Private Sub FillvsFactors_Table_Tasvie()
    On Error GoTo Err_Handler
    
    Dim Rst As New ADODB.Recordset
    Dim CountTableFich As Integer
    CountTableFich = 0
    ReDim Parameter(2) As Parameter
    Parameter(0) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
    Parameter(1) = GenerateInputParameter("@nvcDate", adWChar, 8, mvarDate)
    Parameter(2) = GenerateInputParameter("@Branch", adInteger, 4, CurrentBranch)
    
    Set Rst = RunParametricStoredProcedure2Rec("Get_Factors_Tables", Parameter)

    With vsFactors_Table_Tasvie
        .Rows = 1
        i = 0
        While Rst.EOF <> True
             If Rst!FacPayment = True Then
                If Rst!PartitionID = clsStation.PartitionID Or clsStation.OtherPartition = False Then
                    .Rows = .Rows + 1
                    CountTableFich = CountTableFich + 1
                    i = .Rows - 1
                    .TextMatrix(i, 0) = i
                    .TextMatrix(i, 1) = Rst!intSerialNo
                    .TextMatrix(i, 2) = Rst!TableName
                    .TextMatrix(i, 3) = IIf(IsNull(Rst!FullName), "", Rst!FullName)
                    .TextMatrix(i, 4) = Rst!tempNo
                    .TextMatrix(i, 5) = Rst!sumPrice
                    .TextMatrix(i, 6) = Rst!time
                    .TextMatrix(i, 7) = Rst!No
                    .TextMatrix(i, 8) = Rst!ShiftDescription
                    .TextMatrix(i, 9) = IIf(IsNull(Rst!GuestNo), "", Rst!GuestNo)
                End If
            End If
            Rst.MoveNext
         
        Wend
    If CountTableFich > 0 Then
      .Row = 1
    End If
    End With

    vsFactors_Table_Tasvie.Cell(flexcpAlignment, 0, 0, 0, 6) = flexAlignCenterCenter
    vsFactors_Table_Tasvie.ColAlignment(-1) = flexAlignCenterCenter
    
    Set Rst = Nothing

    Exit Sub
    
Err_Handler:
    ShowErrorMessage
    LogSaveNew "frmFindTableDeliveryFich => ", err.Description, err.Number, err.Source, "FillVsFactors_Table_Tasvie"
End Sub

Private Sub FillvsFactors_Delivery()
    On Error GoTo Err_Handler
    
    Dim Rst As New ADODB.Recordset
    Dim CountDeliveryFich As Integer
    CountDeliveryFich = 0
    If Rst.State = 1 Then Rst.Close
    
    ReDim Parameter(1) As Parameter
    Parameter(0) = GenerateInputParameter("@AccountYear", adSmallInt, 2, AccountYear)
    Parameter(1) = GenerateInputParameter("@Branch", adInteger, 4, CurrentBranch)
    Set Rst = RunParametricStoredProcedure2Rec("GetCustomersInfo", Parameter)
    
    Dim VarToday As String
    VarToday = mvarDate
    With vsFactors_Delivery
        .Rows = 1
        
        If Not (Rst.EOF = True And Rst.BOF = True) Then
        
            'Rst.MoveFirst
            i = 1
            While Rst.EOF = False
                If ChkDaily.Value = False Or Rst.Fields("Date").Value = VarToday Then
                    .Rows = .Rows + 1
                     CountDeliveryFich = CountDeliveryFich + 1
                    .TextMatrix(i, 0) = i
                    .TextMatrix(i, 1) = Rst.Fields("intSerialNo").Value
                    .TextMatrix(i, 2) = Rst.Fields("TempNo").Value
                    .TextMatrix(i, 3) = Rst.Fields("Code").Value
                    .TextMatrix(i, 4) = Rst.Fields("Full Name").Value
                    .TextMatrix(i, 5) = Rst.Fields("SumPrice").Value
                    .TextMatrix(i, 6) = Rst.Fields("Time").Value
                    .TextMatrix(i, 7) = Rst.Fields("Date").Value
                    .TextMatrix(i, 8) = Rst.Fields("ShiftDescription").Value
                    .TextMatrix(i, 9) = Rst.Fields("No").Value
                    .TextMatrix(i, 10) = Rst.Fields("Address").Value
                    i = i + 1
                End If
                
                Rst.MoveNext
            Wend
        End If
        '.AutoSizeMode = flexAutoSizeColWidth
        '.AutoSize 0, .Cols - 1
        If CountDeliveryFich > 0 Then
        .Row = 1
        End If
    End With
    
    If Rst.State <> 0 Then Rst.Close
    Set Rst = Nothing
   
    Exit Sub
    
Err_Handler:
    ShowErrorMessage
    LogSaveNew "frmFindTableDeliveryFich => ", err.Description, err.Number, err.Source, "FillVsFactors_Delivery"
End Sub
Private Sub FillvsFactors_Fich()
    On Error GoTo Err_Handler
    
    If Len(txtDate1.Text) <> 10 Or Len(txtDate1.Text) <> 10 Then
        ShowDisMessage "  «—ÌŒ „⁄ »— Ê«—œ ﬂ‰Ìœ ", 1000
        Exit Sub
    End If
    FWProgressBar1.Visible = True
    FWProgressBar1.Value = 0
    
    Dim Rst As New ADODB.Recordset
    
    ReDim Parameter(5) As Parameter
    Parameter(0) = GenerateInputParameter("@Status", adInteger, 4, mvarStatus)
    Parameter(1) = GenerateInputParameter("@User", adInteger, 4, mvarCurUserNo)
    Parameter(2) = GenerateInputParameter("@AccountYear", adSmallInt, 2, AccountYear)
    Parameter(3) = GenerateInputParameter("@Branch", adInteger, 4, CurrentBranch)
    Parameter(4) = GenerateInputParameter("@DateAfter", adVarWChar, 8, Right(txtDate1.Text, 8))
    Parameter(5) = GenerateInputParameter("@DateBefore", adVarWChar, 8, Right(txtDate2.Text, 8))
    
    Set Rst = RunParametricStoredProcedure2Rec("Get_All_Factors", Parameter)

    With vsFactors_Fich
        .Rows = 1
        i = 0
        While Rst.EOF <> True
            .Rows = .Rows + 1
            i = .Rows - 1
            .TextMatrix(i, 0) = i
            .TextMatrix(i, 1) = Rst!intSerialNo
            .TextMatrix(i, 2) = Rst!tempNo
            .TextMatrix(i, 3) = Rst![CustomerName]
            .TextMatrix(i, 4) = Rst![Date]
            .TextMatrix(i, 5) = Rst![time]
            .TextMatrix(i, 6) = Rst!sumPrice
            .TextMatrix(i, 7) = Rst!No
            .TextMatrix(i, 8) = Rst!nvcFirstName & " " & Rst!nvcSurname
            .TextMatrix(i, 9) = Rst!Balance
            .TextMatrix(i, 10) = Rst!Recursive
            .TextMatrix(i, 11) = Rst!ServiceTotal
            .TextMatrix(i, 12) = Rst!CarryFeeTotal
            .TextMatrix(i, 13) = Rst!DiscountTotal
            .TextMatrix(i, 14) = Rst!ShiftDescription
            .TextMatrix(i, 15) = IIf(IsNull(Rst!NvcDescription), "", Rst!NvcDescription)
            If Rst!Balance = False And Rst!Recursive = 0 Then .Cell(flexcpBackColor, i, 0, i, .Cols - 1) = vbRed
            If Rst!Recursive = 1 Then .Cell(flexcpBackColor, i, 0, i, .Cols - 1) = vbCyan
            Rst.MoveNext
         
            FWProgressBar1.Value = FWProgressBar1.Value + 1
            If FWProgressBar1.Value = 1000 Then
               FWProgressBar1.Value = 0
            End If
        Wend
    End With

    FWProgressBar1.Value = 0
    FWProgressBar1.Visible = False
    
    Set Rst = Nothing

    Exit Sub
    
Err_Handler:
    ShowErrorMessage
    LogSaveNew "frmFindTableDeliveryFich => ", err.Description, err.Number, err.Source, "FillVsFactors_Fich"

End Sub

Private Sub txtNo_Table1_Change()
  i = -1
    If Len(txtNo_Table1.Text) <> 0 Then
       i = vsMultiFactors_Table.FindRow(txtNo_Table1.Text, 1, 3, True, True)
    End If
    If i > 0 Then
        vsMultiFactors_Table.Row = i
        vsMultiFactors_Table.ShowCell i, 0
        LblFindFactor.Caption = ""
        'vsFactors_Table.SetFocus
    Else
        vsMultiFactors_Table.Row = 0
        vsMultiFactors_Table.ShowCell 0, 0
        If Val(txtNo_Table1.Text) > 0 Then
           LblFindFactor.Caption = "›—Ê‘ œ— „Ì“ " & Val(txtNo_Table1.Text) & " «‰Ã«„ ‰‘œÂ"
         Else
            LblFindFactor.Caption = "‘„«—Â „Ì“ —« Ê«—œ ﬂ‰Ìœ  "
         End If
    
    End If
    FillvsFactorDetail
    FillvsReceived
    FillvsHistory
End Sub

Private Sub txtNo_Table1_GotFocus()
  vsMultiFactors_Table.Row = 0
    vsMultiFactors_Table.Select vsMultiFactors_Table.Row, 3
    vsMultiFactors_Table.Sort = flexSortGenericAscending
  '  vsFactors_Table.Sort = flexSortGenericDescending
    LblFindFactor.Caption = "‘„«—Â „Ì“ —« Ê«—œ ﬂ‰Ìœ  "
    For i = 1 To vsMultiFactors_Table.Rows - 1
        vsMultiFactors_Table.TextMatrix(i, 0) = i
    Next
End Sub


Private Sub TxtNo_Temp_Change()
    i = -1
    If optShowFich(0).Value = True Then
  '      i = vsFactors_Fich.FindRow(txtNo_Fich.Text, 1, 2, True, True)
        If Val(TxtNo_Temp.Text) > 0 Then
           Define_Factor_Temp
        Else
            vsFactors_Fich.Rows = 1
         
        End If
    Else
'        If vsFactors_Fich.Rows >= 1000 Then
'            If Len(TxtNo_Fich.Text) = 3 Then
'               i = vsFactors_Fich.FindRow(TxtNo_Fich.Text, 1, 2, True, True)
'            End If
'        Else
            i = vsFactors_Fich.FindRow(TxtNo_Temp.Text, 1, 2, True, True)
'        End If
    End If
    If i > 0 Then
        vsFactors_Fich.Row = i
        vsFactors_Fich.ShowCell i, 0
        LblFindFactor.Caption = ""
    Else
        vsFactors_Fich.Row = 0
        vsFactors_Fich.ShowCell 0, 0
        If Val(TxtNo_Temp.Text) > 0 Then
           LblFindFactor.Caption = " ›«ﬂ Ê— " & Val(TxtNo_Temp.Text) & "  œ— ”Ì” „  ⁄—Ì› ‰‘œÂ"
         Else
            LblFindFactor.Caption = "‘„«—Â ›«ﬂ Ê— —« Ê«—œ ﬂ‰Ìœ  "
         End If
    
    End If
    FillvsFactorDetail
    FillvsReceived
    FillvsHistory

End Sub

Private Sub TxtNo_Temp_GotFocus()
    vsFactors_Fich.Row = 0
    vsFactors_Fich.Select vsFactors_Fich.Row, 7
'    vsFactors_Fich.Sort = flexSortGenericAscending
    vsFactors_Fich.Sort = flexSortGenericDescending
    LblFindFactor.Caption = "‘„«—Â ›«ﬂ Ê— —Ê“«‰Â —« Ê«—œ ﬂ‰Ìœ  "

End Sub

Private Sub vsFactorDetail_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    
    For i = 0 To vsFactorDetail.Cols - 1
        SaveSetting strMainKey, Me.Name & "vsFactorDetail", "Col" & i, vsFactorDetail.ColWidth(i)
    Next

End Sub
Private Sub VSReceived_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    
    For i = 0 To vsReceived.Cols - 1
        SaveSetting strMainKey, Me.Name & "VSReceived", "Col" & i, vsReceived.ColWidth(i)
    Next

End Sub
Private Sub VSHistory_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    
    For i = 0 To VSHistory.Cols - 1
        SaveSetting strMainKey, Me.Name & "VSHistory", "Col" & i, VSHistory.ColWidth(i)
    Next

End Sub

Private Sub vsFactors_Delivery_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    
    For i = 0 To vsFactors_Delivery.Cols - 1
        SaveSetting strMainKey, Me.Name & "vsFactors_Delivery", "Col" & i, vsFactors_Delivery.ColWidth(i)
    Next

End Sub

Private Sub vsFactors_Delivery_RowColChange()
    FillvsFactorDetail
    FillvsReceived
    FillvsHistory
End Sub

Private Sub vsFactors_Fich_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    
    For i = 0 To vsFactors_Fich.Cols - 1
        SaveSetting strMainKey, Me.Name & "vsFactors_Fich", "Col" & i, vsFactors_Fich.ColWidth(i)
    Next

End Sub

Private Sub vsFactors_Fich_RowColChange()
    FillvsFactorDetail
    FillvsReceived
    FillvsHistory
End Sub
    
Private Sub vsFactors_Table_AfterSort(ByVal Col As Long, Order As Integer)
    For i = 1 To vsFactors_Table.Rows - 1
        vsFactors_Table.TextMatrix(i, 0) = i
    Next
End Sub
Private Sub vsFactors_Table_Tasvie_AfterSort(ByVal Col As Long, Order As Integer)
    For i = 1 To vsFactors_Table_Tasvie.Rows - 1
        vsFactors_Table_Tasvie.TextMatrix(i, 0) = i
    Next
End Sub

Private Sub vsFactors_Table_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    
    For i = 0 To vsFactors_Table.Cols - 1
        SaveSetting strMainKey, Me.Name & "vsFactors_Table", "Col" & i, vsFactors_Table.ColWidth(i)
    Next

End Sub
Private Sub vsFactors_Table_Tasvie_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    
    For i = 0 To vsFactors_Table_Tasvie.Cols - 1
        SaveSetting strMainKey, Me.Name & "vsFactors_Table_Tasvie", "Col" & i, vsFactors_Table_Tasvie.ColWidth(i)
    Next

End Sub

Private Sub vsFactors_Table_DblClick()
    On Error GoTo Err_Handler
    
    If vsFactors_Table.Row > 0 Then
        OKButton_Click
    End If
    
    Exit Sub
    
Err_Handler:
    ShowErrorMessage
    LogSaveNew "frmFindTableDeliveryFich => ", err.Description, err.Number, err.Source, "vsFactors_Table_DblClick"
End Sub
Private Sub vsFactors_Table_Tasvie_DblClick()
    On Error GoTo Err_Handler
    
    If vsFactors_Table_Tasvie.Row > 0 Then
        OKButton_Click
    End If
    
    Exit Sub
    
Err_Handler:
    ShowErrorMessage
    LogSaveNew "frmFindTableDeliveryFich => ", err.Description, err.Number, err.Source, "vsFactors_Table_Tasvie_DblClick"
End Sub

Private Sub vsFactors_Delivery_AfterSort(ByVal Col As Long, Order As Integer)
    For i = 1 To vsFactors_Delivery.Rows - 1
        vsFactors_Delivery.TextMatrix(i, 0) = i
    Next
End Sub

Private Sub vsFactors_Delivery_DblClick()
    If vsFactors_Delivery.Row > 0 Then
        OKButton_Click
    End If
End Sub

Private Sub vsFactors_Fich_AfterSort(ByVal Col As Long, Order As Integer)
    For i = 1 To vsFactors_Fich.Rows - 1
        vsFactors_Fich.TextMatrix(i, 0) = i
    Next
    
End Sub

Private Sub vsFactors_Fich_DblClick()
    If vsFactors_Fich.Row > 0 Then
        OKButton_Click
    End If
End Sub

Private Sub Define_Factor()
    On Error GoTo Err_Handler
    
    If SSTab1.Tab = 0 Then
        Dim Rst As New ADODB.Recordset
        
        ReDim Parameter(4) As Parameter
        Parameter(0) = GenerateInputParameter("@Status", adInteger, 4, mvarStatus)
        Parameter(1) = GenerateInputParameter("@User", adInteger, 4, mvarCurUserNo)
        Parameter(2) = GenerateInputParameter("@No", adBigInt, 8, Val(TxtNo_Fich.Text))
        Parameter(3) = GenerateInputParameter("@AccountYear", adSmallInt, 2, AccountYear)
        Parameter(4) = GenerateInputParameter("@Branch", adInteger, 4, CurrentBranch)
        Set Rst = RunParametricStoredProcedure2Rec("Get_Define_Factors", Parameter)
    
        With vsFactors_Fich
            .Rows = 1
            i = 0
            While Rst.EOF <> True
                .Rows = .Rows + 1
                i = .Rows - 1
                .TextMatrix(i, 0) = i
                .TextMatrix(i, 1) = Rst!intSerialNo
                .TextMatrix(i, 2) = Rst!tempNo
                .TextMatrix(i, 3) = Rst![CustomerName]
                .TextMatrix(i, 4) = Rst![Date]
                .TextMatrix(i, 5) = Rst![time]
                .TextMatrix(i, 6) = Rst!sumPrice
                .TextMatrix(i, 7) = Rst!No
                .TextMatrix(i, 8) = Rst!nvcFirstName & " " & Rst!nvcSurname
                .TextMatrix(i, 9) = Rst!Balance
                .TextMatrix(i, 10) = Rst!Recursive
                .TextMatrix(i, 11) = Rst!ServiceTotal
                .TextMatrix(i, 12) = Rst!CarryFeeTotal
                .TextMatrix(i, 13) = Rst!DiscountTotal
                .TextMatrix(i, 14) = Rst!ShiftDescription
                .TextMatrix(i, 15) = IIf(IsNull(Rst!NvcDescription), "", Rst!NvcDescription)
                If Rst!Balance = False And Rst!Recursive = 0 Then .Cell(flexcpBackColor, i, 0, i, .Cols - 1) = vbRed
                If Rst!Recursive = 1 Then .Cell(flexcpBackColor, i, 0, i, .Cols - 1) = vbCyan
                Rst.MoveNext
             
            Wend
        End With
        
        Set Rst = Nothing
    ElseIf SSTab1.Tab = 1 Then
    
    
    ElseIf SSTab1.Tab = 2 Then
    
    
    End If
Exit Sub

Err_Handler:
    ShowErrorMessage
    LogSaveNew "frmFindTableDeliveryFich => ", err.Description, err.Number, err.Source, "Define_Factor"
End Sub
Private Sub Define_Factor_Temp()
    On Error GoTo Err_Handler
    
    If SSTab1.Tab = 0 Then
        Dim Rst As New ADODB.Recordset
        
        ReDim Parameter(4) As Parameter
        Parameter(0) = GenerateInputParameter("@Status", adInteger, 4, mvarStatus)
        Parameter(1) = GenerateInputParameter("@User", adInteger, 4, mvarCurUserNo)
        Parameter(2) = GenerateInputParameter("@TempNo", adBigInt, 8, Val(TxtNo_Temp.Text))
        Parameter(3) = GenerateInputParameter("@AccountYear", adSmallInt, 2, AccountYear)
        Parameter(4) = GenerateInputParameter("@Branch", adInteger, 4, CurrentBranch)
        
        Set Rst = RunParametricStoredProcedure2Rec("Get_Define_Factors_Temp", Parameter)
    
        With vsFactors_Fich
            .Rows = 1
            i = 0
            While Rst.EOF <> True
                .Rows = .Rows + 1
                i = .Rows - 1
                .TextMatrix(i, 0) = i
                .TextMatrix(i, 1) = Rst!intSerialNo
                .TextMatrix(i, 2) = Rst!tempNo
                .TextMatrix(i, 3) = Rst![CustomerName]
                .TextMatrix(i, 4) = Rst![Date]
                .TextMatrix(i, 5) = Rst![time]
                .TextMatrix(i, 6) = Rst!sumPrice
                .TextMatrix(i, 7) = Rst!No
                .TextMatrix(i, 8) = Rst!nvcFirstName & " " & Rst!nvcSurname
                .TextMatrix(i, 9) = Rst!Balance
                .TextMatrix(i, 10) = Rst!Recursive
                .TextMatrix(i, 11) = Rst!ServiceTotal
                .TextMatrix(i, 12) = Rst!CarryFeeTotal
                .TextMatrix(i, 13) = Rst!DiscountTotal
                .TextMatrix(i, 14) = Rst!ShiftDescription
                .TextMatrix(i, 15) = IIf(IsNull(Rst!NvcDescription), "", Rst!NvcDescription)
                If Rst!Balance = False And Rst!Recursive = 0 Then .Cell(flexcpBackColor, i, 0, i, .Cols - 1) = vbRed
                If Rst!Recursive = 1 Then .Cell(flexcpBackColor, i, 0, i, .Cols - 1) = vbCyan
                Rst.MoveNext
             
            Wend
        End With
        
        Set Rst = Nothing
    ElseIf SSTab1.Tab = 1 Then
    
    
    ElseIf SSTab1.Tab = 2 Then
    
    
    End If
Exit Sub

Err_Handler:
    ShowErrorMessage
    LogSaveNew "frmFindTableDeliveryFich => ", err.Description, err.Number, err.Source, "Define_Factor"
End Sub

Public Sub FillvsFactorDetail() ' fills the detail of the current factor
    On Error GoTo Err_Handler
    
    Dim i As Integer
    Dim intselFactor As Double
    
    Dim Rst As New ADODB.Recordset
    
    If Rst.State = 1 Then Rst.Close
    
    With vsFactorDetail
       
        ' if at least there is one , choose the current one
        If SSTab1.Tab = 0 Then
           If vsFactors_Fich.Rows <= 1 Or vsFactors_Fich.Row = 0 Then
                vsFactorDetail.Rows = 1
                Exit Sub
           End If
         '  On Error Resume Next
           intselFactor = Val(vsFactors_Fich.TextMatrix(vsFactors_Fich.Row, 1))
        ElseIf SSTab1.Tab = 1 Then
           If vsFactors_Table.Rows <= 1 Or vsFactors_Table.Row = 0 Then ''Or clsStation.SearchFichDefault = False Then
                vsFactorDetail.Rows = 1
                Exit Sub
           End If
        '   On Error Resume Next
           intselFactor = Val(vsFactors_Table.TextMatrix(vsFactors_Table.Row, 1))
        ElseIf SSTab1.Tab = 2 Then
           If vsFactors_Delivery.Rows <= 1 Or vsFactors_Delivery.Row = 0 Then
                vsFactorDetail.Rows = 1
                Exit Sub
           End If
        '   On Error Resume Next
           intselFactor = Val(vsFactors_Delivery.TextMatrix(vsFactors_Delivery.Row, 1))
         ElseIf SSTab1.Tab = 3 Then
           If vsMultiFactors_Table.Rows <= 1 Or vsMultiFactors_Table.Row = 0 Then
                vsFactorDetail.Rows = 1
                Exit Sub
           End If
        '   On Error Resume Next
           intselFactor = Val(vsMultiFactors_Table.TextMatrix(vsMultiFactors_Table.Row, 1))
        ElseIf SSTab1.Tab = 4 Then
           If vsFactors_Table_Tasvie.Rows <= 1 Or vsFactors_Table_Tasvie.Row = 0 Then ''Or clsStation.SearchFichDefault = False Then
                vsFactorDetail.Rows = 1
                Exit Sub
           End If
        '   On Error Resume Next
           intselFactor = Val(vsFactors_Table_Tasvie.TextMatrix(vsFactors_Table_Tasvie.Row, 1))
        End If
        
        
        ReDim Parameter(1) As Parameter
        Parameter(0) = GenerateInputParameter("@intSelFactor", adInteger, 4, intselFactor)
        Parameter(1) = GenerateInputParameter("@Branch", adInteger, 4, CurrentBranch)
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
       ''.AutoSizeMode = flexAutoSizeColWidth  ' set the collumns' width
       '' .AutoSize 0, .Cols - 1
    End With
    
    If Rst.State <> 0 Then Rst.Close
    Set Rst = Nothing
    
    Exit Sub
    
Err_Handler:
    ShowErrorMessage
    LogSaveNew "frmFindTableDeliveryFich => ", err.Description, err.Number, err.Source, "FillVsFactorDetail"
End Sub
Public Sub FillvsReceived() ' fills the detail of the current factor
    On Error GoTo Err_Handler
    
    Dim i As Integer
    Dim intselFactor As Long
    
    Dim Rst As New ADODB.Recordset
    
    If Rst.State = 1 Then Rst.Close
    
    With vsReceived
       
        ' if at least there is one , choose the current one
        If SSTab1.Tab = 0 Then
           If vsFactors_Fich.Rows <= 1 Or vsFactors_Fich.Row = 0 Then
                .Rows = 1
                Exit Sub
           End If
         '  On Error Resume Next
           intselFactor = Val(vsFactors_Fich.TextMatrix(vsFactors_Fich.Row, 1))
        ElseIf SSTab1.Tab = 1 Then
           If vsFactors_Table.Rows <= 1 Or vsFactors_Table.Row = 0 Then ''Or clsStation.SearchFichDefault = False Then
                .Rows = 1
                Exit Sub
           End If
        '   On Error Resume Next
           intselFactor = Val(vsFactors_Table.TextMatrix(vsFactors_Table.Row, 1))
        ElseIf SSTab1.Tab = 2 Then
           If vsFactors_Delivery.Rows <= 1 Or vsFactors_Delivery.Row = 0 Then
                .Rows = 1
                Exit Sub
           End If
        '   On Error Resume Next
           intselFactor = Val(vsFactors_Delivery.TextMatrix(vsFactors_Delivery.Row, 1))
         ElseIf SSTab1.Tab = 3 Then
           If vsMultiFactors_Table.Rows <= 1 Or vsMultiFactors_Table.Row = 0 Then
                .Rows = 1
                Exit Sub
           End If
        '   On Error Resume Next
           intselFactor = Val(vsMultiFactors_Table.TextMatrix(vsMultiFactors_Table.Row, 1))
        ElseIf SSTab1.Tab = 3 Then
           If vsFactors_Table_Tasvie.Rows <= 1 Or vsFactors_Table_Tasvie.Row = 0 Then ''Or clsStation.SearchFichDefault = False Then
                .Rows = 1
                Exit Sub
           End If
        '   On Error Resume Next
           intselFactor = Val(vsFactors_Table_Tasvie.TextMatrix(vsFactors_Table_Tasvie.Row, 1))
        End If
        
        ReDim Parameter(1) As Parameter
        
        Parameter(0) = GenerateInputParameter("@intSerialNo", adBigInt, 8, intselFactor)
        Parameter(1) = GenerateInputParameter("@Branch", adInteger, 4, CurrentBranch)
        
        Set Rst = RunParametricStoredProcedure2Rec("Get_PayFactors", Parameter)
   
        .Rows = 1
        
        If Not (Rst.EOF = True And Rst.BOF = True) Then
        
            i = 1
            While Rst.EOF = False ' fill the grid
                .Rows = .Rows + 1
                .TextMatrix(i, 0) = i
                .TextMatrix(i, 1) = Rst.Fields("intAmount").Value
                .TextMatrix(i, 2) = Rst.Fields("Date").Value
                .TextMatrix(i, 3) = Rst.Fields("RegTime").Value
                .TextMatrix(i, 4) = Rst.Fields("Type").Value
                Rst.MoveNext
                i = i + 1
            Wend
            
        End If
    End With
    
    If Rst.State <> 0 Then Rst.Close
    Set Rst = Nothing
    
    Exit Sub
    
Err_Handler:
    ShowErrorMessage
    LogSaveNew "frmFindTableDeliveryFich => ", err.Description, err.Number, err.Source, "FillVsReceived"
End Sub
Public Sub FillvsHistory() ' fills the detail of the current factor
    On Error GoTo Err_Handler
    
    Dim i As Integer
    Dim intselFactor As Long
    
    Dim Rst As New ADODB.Recordset
    
    If Rst.State = 1 Then Rst.Close
    
    With VSHistory
       
        ' if at least there is one , choose the current one
        If SSTab1.Tab = 0 Then
           If vsFactors_Fich.Rows <= 1 Or vsFactors_Fich.Row = 0 Then
                .Rows = 1
                Exit Sub
           End If
         '  On Error Resume Next
           intselFactor = Val(vsFactors_Fich.TextMatrix(vsFactors_Fich.Row, 1))
        ElseIf SSTab1.Tab = 1 Then
           If vsFactors_Table.Rows <= 1 Or vsFactors_Table.Row = 0 Then ''Or clsStation.SearchFichDefault = False Then
                .Rows = 1
                Exit Sub
           End If
        '   On Error Resume Next
           intselFactor = Val(vsFactors_Table.TextMatrix(vsFactors_Table.Row, 1))
        ElseIf SSTab1.Tab = 2 Then
           If vsFactors_Delivery.Rows <= 1 Or vsFactors_Delivery.Row = 0 Then
                .Rows = 1
                Exit Sub
           End If
        '   On Error Resume Next
           intselFactor = Val(vsFactors_Delivery.TextMatrix(vsFactors_Delivery.Row, 1))
         ElseIf SSTab1.Tab = 3 Then
           If vsMultiFactors_Table.Rows <= 1 Or vsMultiFactors_Table.Row = 0 Then
                .Rows = 1
                Exit Sub
           End If
        '   On Error Resume Next
           intselFactor = Val(vsMultiFactors_Table.TextMatrix(vsMultiFactors_Table.Row, 1))
        ElseIf SSTab1.Tab = 4 Then
           If vsFactors_Table_Tasvie.Rows <= 1 Or vsFactors_Table_Tasvie.Row = 0 Then ''Or clsStation.SearchFichDefault = False Then
                .Rows = 1
                Exit Sub
           End If
        '   On Error Resume Next
           intselFactor = Val(vsFactors_Table_Tasvie.TextMatrix(vsFactors_Table_Tasvie.Row, 1))
        End If
        
        ReDim Parameter(0) As Parameter
        
        Parameter(0) = GenerateInputParameter("@intSerialNo", adBigInt, 8, intselFactor)
        
        Set Rst = RunParametricStoredProcedure2Rec("Get_History_By_intSerialNo", Parameter)
   
        .Rows = 1
        
        If Not (Rst.EOF = True And Rst.BOF = True) Then
        
            i = 1
            While Rst.EOF = False ' fill the grid
                .Rows = .Rows + 1
                .TextMatrix(i, 0) = i
                .TextMatrix(i, 1) = Rst.Fields("ActionDescription").Value
                .TextMatrix(i, 2) = Rst.Fields("RegDate").Value
                .TextMatrix(i, 3) = Rst.Fields("RegTime").Value
                Rst.MoveNext
                i = i + 1
            Wend
            
        End If
    End With
    
    If Rst.State <> 0 Then Rst.Close
    Set Rst = Nothing
    
    Exit Sub
    
Err_Handler:
    ShowErrorMessage
    LogSaveNew "frmFindTableDeliveryFich => ", err.Description, err.Number, err.Source, "FillVsHistory"
End Sub

Private Sub vsFactors_Table_RowColChange()
    FillvsFactorDetail
    FillvsReceived
    FillvsHistory
End Sub
Private Sub vsFactors_Table_Tasvie_RowColChange()
    FillvsFactorDetail
    FillvsReceived
    FillvsHistory
End Sub

Private Sub KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then
        Unload Me
    ElseIf KeyCode = 13 Then
        OKButton_Click
    ElseIf KeyCode = 38 Or KeyCode = 39 Or KeyCode = 40 Or KeyCode = 41 Then
        If SSTab1.Tab = 0 Then
            vsFactors_Fich.SetFocus
        ElseIf SSTab1.Tab = 1 Then
            vsFactors_Table.SetFocus
        ElseIf SSTab1.Tab = 2 Then
            vsFactors_Delivery.SetFocus
        End If
    End If

End Sub

Private Sub FillvsMultiFactors_Table()
    On Error GoTo Err_Handler
    strTableNoDetailString = ""
    
    Dim Rst As New ADODB.Recordset
    Dim CountTableFich As Integer
    CountTableFich = 0
    ReDim Parameter(2) As Parameter
    Parameter(0) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
    Parameter(1) = GenerateInputParameter("@nvcDate", adWChar, 8, mvarDate)
    Parameter(2) = GenerateInputParameter("@Branch", adInteger, 4, CurrentBranch)
    
    Set Rst = RunParametricStoredProcedure2Rec("Get_Factors_Tables", Parameter)

    With vsMultiFactors_Table
        .Rows = 1
        i = 0
       .ColDataType(1) = flexDTBoolean
        While Rst.EOF <> True
             If Rst!FacPayment = False Then
                If Rst!PartitionID = clsStation.PartitionID Or clsStation.OtherPartition = False Then
                    .Rows = .Rows + 1
                    CountTableFich = CountTableFich + 1
                    i = .Rows - 1
                    .TextMatrix(i, 0) = i
                    .TextMatrix(i, 1) = Rst!intSerialNo
                    ''.TextMatrix(i, 1) = 1
                    .TextMatrix(i, 3) = Rst!TableName
                    .TextMatrix(i, 4) = IIf(IsNull(Rst!FullName), "", Rst!FullName)
                    .TextMatrix(i, 5) = Rst!tempNo
                    .TextMatrix(i, 6) = Rst!sumPrice
                    .TextMatrix(i, 7) = Rst!time
                    .TextMatrix(i, indexColTableNo) = Rst!TableNo
                    .TextMatrix(i, 9) = Rst!No
                    .TextMatrix(i, 10) = Rst!ShiftDescription
                    .TextMatrix(i, 11) = IIf(IsNull(Rst!GuestNo), "", Rst!GuestNo)
                End If
            End If
            Rst.MoveNext
        Wend
        
    If CountTableFich > 0 Then
      .Row = 1
    End If
    End With

    vsMultiFactors_Table.Cell(flexcpAlignment, 0, 0, 0, 6) = flexAlignCenterCenter
    vsMultiFactors_Table.ColAlignment(-1) = flexAlignCenterCenter
    
    Set Rst = Nothing

    Exit Sub
    
Err_Handler:
    ShowErrorMessage
    LogSaveNew "frmFindTableDeliveryFich => ", err.Description, err.Number, err.Source, "FillVsMultiFactors_Table"
End Sub

Private Sub vsMultiFactors_Table_AfterSort(ByVal Col As Long, Order As Integer)
 For i = 1 To vsMultiFactors_Table.Rows - 1
        vsMultiFactors_Table.TextMatrix(i, 0) = i
    Next
End Sub

Private Sub vsMultiFactors_Table_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    
    For i = 0 To vsMultiFactors_Table.Cols - 1
        SaveSetting strMainKey, Me.Name & "vsMultiFactors_Table", "Col" & i, vsMultiFactors_Table.ColWidth(i)
    Next

End Sub

Private Sub vsMultiFactors_Table_Click()
    With vsMultiFactors_Table
        If .Row > 0 And .Col = 2 Then
         .Select .Row, .Col
         .EditCell
        End If
    End With
End Sub

Private Sub vsMultiFactors_Table_RowColChange()
    FillvsFactorDetail
    FillvsReceived
    FillvsHistory
End Sub
