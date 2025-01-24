VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{7AEDC602-D94C-11D1-BB7A-00E0290EA3C9}#1.0#0"; "ResizeKit.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmCreditCustomer 
   ClientHeight    =   8655
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12540
   BeginProperty Font 
      Name            =   "B Homa"
      Size            =   11.25
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCreditCustomer.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8655
   ScaleWidth      =   12540
   Begin VB.Frame Frame2 
      Height          =   2295
      Left            =   9840
      TabIndex        =   24
      Top             =   6240
      Width           =   2655
      Begin VB.CommandButton cmdUpdate 
         BackColor       =   &H008080FF&
         Caption         =   "»Â —Ê“ —”«‰Ì «ÿ·«⁄«  „‘ —Ì«‰"
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
         Left            =   360
         RightToLeft     =   -1  'True
         TabIndex        =   29
         Top             =   1440
         Width           =   1815
      End
      Begin MSMask.MaskEdBox txtDatefrom 
         Height          =   465
         Left            =   120
         TabIndex        =   25
         Top             =   360
         Width           =   1395
         _ExtentX        =   2461
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
      Begin MSMask.MaskEdBox txtDateto 
         Height          =   465
         Left            =   120
         TabIndex        =   26
         Top             =   840
         Width           =   1455
         _ExtentX        =   2566
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
      Begin VB.Label Label6 
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
         Height          =   495
         Left            =   1560
         RightToLeft     =   -1  'True
         TabIndex        =   28
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label5 
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
         Height          =   375
         Left            =   1560
         RightToLeft     =   -1  'True
         TabIndex        =   27
         Top             =   840
         Width           =   855
      End
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
      Left            =   9960
      RightToLeft     =   -1  'True
      TabIndex        =   22
      Text            =   "Combo1"
      Top             =   600
      Width           =   2355
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
      TabIndex        =   20
      Top             =   0
      Width           =   1695
   End
   Begin VB.CommandButton CmdCancel 
      BackColor       =   &H00404080&
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
      Left            =   360
      TabIndex        =   17
      Top             =   7560
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5565
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   540
      Width           =   9735
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
         Height          =   765
         Left            =   3720
         RightToLeft     =   -1  'True
         TabIndex        =   23
         Top             =   4680
         Width           =   1095
      End
      Begin VB.CommandButton cmdPayAll 
         BackColor       =   &H00000080&
         Caption         =   " ”ÊÌÂ Õ”«» ò·Ì œ— Â„«‰ —Ê“"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   4680
         Width           =   1815
      End
      Begin VB.CommandButton cmdPaySome 
         BackColor       =   &H000000C0&
         Caption         =   " ”ÊÌÂ Õ”«» œ— Â„«‰ —Ê“"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   2040
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   4680
         Width           =   1575
      End
      Begin VSFlex7LCtl.VSFlexGrid vsOwedFactors 
         Height          =   2985
         Left            =   180
         TabIndex        =   4
         Top             =   810
         Width           =   9435
         _cx             =   16642
         _cy             =   5265
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
      Begin VB.CheckBox chkNoPaykDelivery 
         Alignment       =   1  'Right Justify
         Caption         =   "›«ò Ê—Â«Ì «—”«·Ì »œÊ‰ ÅÌò"
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
         Left            =   6840
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   4920
         Visible         =   0   'False
         Width           =   2625
      End
      Begin VB.Label lblSelected 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   1680
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Top             =   3840
         Visible         =   0   'False
         Width           =   1965
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "„»·€ ›«ò Ê—Â«Ì «‰ Œ«» ‘œÂ"
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
         Height          =   525
         Left            =   3645
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   3900
         Visible         =   0   'False
         Width           =   2265
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "„»·€ ò· »œÂÌ"
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
         Left            =   8160
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   3900
         Width           =   1245
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
         Left            =   3480
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   240
         Width           =   5895
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
         Left            =   7920
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   4365
         Width           =   1530
      End
      Begin VB.Label lblNoOfFactors 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   6240
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   4320
         Width           =   1605
      End
      Begin VB.Label lblShouldBePaid 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   6000
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   3900
         Width           =   2205
      End
      Begin VB.Label lblMessage 
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
         Height          =   525
         Left            =   60
         TabIndex        =   5
         Top             =   4200
         Width           =   5775
      End
   End
   Begin VB.Timer Timer1 
      Left            =   120
      Top             =   630
   End
   Begin VSFlex7LCtl.VSFlexGrid vsFactorDetail 
      Height          =   2415
      Left            =   1920
      TabIndex        =   11
      Top             =   6120
      Width           =   7935
      _cx             =   13996
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
   Begin VSFlex7LCtl.VSFlexGrid vsOwedCustomers 
      Height          =   4575
      Left            =   9840
      TabIndex        =   12
      Top             =   1560
      Width           =   2655
      _cx             =   4683
      _cy             =   8070
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
      Cols            =   10
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   350
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmCreditCustomer.frx":A4C2
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
      OleObjectBlob   =   "frmCreditCustomer.frx":A5A1
      TabIndex        =   19
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
      TabIndex        =   21
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   " ”ÊÌÂ Õ”«» »« „‘ —Ì«‰ «⁄ »«—Ì Ê »œÂﬂ«—"
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
      Left            =   3840
      RightToLeft     =   -1  'True
      TabIndex        =   18
      Top             =   0
      Width           =   4935
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "·Ì”  „‘ —Ì«‰ »œÂò«—"
      Height          =   525
      Left            =   9840
      RightToLeft     =   -1  'True
      TabIndex        =   14
      Top             =   1080
      Width           =   2535
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "«ﬁ·«„ ›«ò Ê—"
      Height          =   525
      Left            =   600
      RightToLeft     =   -1  'True
      TabIndex        =   13
      Top             =   6240
      Width           =   1215
   End
End
Attribute VB_Name = "frmCreditCustomer"
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
    
    With vsOwedFactors
        lblSelected.Caption = ""
        If .Rows < 2 Then
            Exit Sub
        End If
        For i = 1 To .Rows - 1
            If Val(.TextMatrix(i, 1)) <> 0 Then
                tempPrice = tempPrice + Val(.TextMatrix(i, 6))
            End If
        Next i
        If tempPrice <> 0 Then
            lblSelected.Visible = True
            Label7.Visible = True
            lblSelected.Caption = tempPrice
        Else
            lblSelected.Visible = False
            Label7.Visible = False
        End If
    End With
End Sub

Public Sub FillvsOwedCustomers()

    If Len(txtDateFrom.ClipText) < 6 Or Len(txtDateTo.ClipText) < 6 Then ShowDisMessage "›Ê—„   «—ÌŒ ’ÕÌÕ ‰Ì” ", 2000: Exit Sub
    Dim Rst As New ADODB.Recordset

    If Rst.State = 1 Then Rst.Close
    
    'find all payks who owe The store
    ReDim Parameter(3) As Parameter
    Parameter(0) = GenerateInputParameter("@AccountYear", adSmallInt, 2, Val(cmbSalMali.Text))
    Parameter(1) = GenerateInputParameter("@Branch", adInteger, 4, CurrentBranch)
    Parameter(2) = GenerateInputParameter("@DateBefore", adVarWChar, 10, CStr(IIf(Trim(txtDateFrom.ClipText) = "", "", Trim(txtDateFrom.Text))))
    Parameter(3) = GenerateInputParameter("@DateAfter", adVarWChar, 10, CStr(IIf(Trim(txtDateTo.ClipText) = "", "", Trim(txtDateTo.Text))))
    Set Rst = RunParametricStoredProcedure2Rec("Get_OwedCreditCustomer", Parameter)
    With vsOwedCustomers
        .Rows = 1
        
        If Not (Rst.EOF = True And Rst.BOF = True) Then
        
            'Rst.moveFirst
            i = 1
            While Rst.EOF = False ' fill the Grid
                .Rows = .Rows + 1
                .TextMatrix(i, 0) = Rst.Fields("code").Value
                .TextMatrix(i, 2) = Rst.Fields("Full Name").Value
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


Public Sub FillvsOwedFactors()

    If Len(txtDateFrom.ClipText) < 6 Or Len(txtDateTo.ClipText) < 6 Then ShowDisMessage "›Ê—„   «—ÌŒ ’ÕÌÕ ‰Ì” ", 2000: Exit Sub
    Dim i As Integer
    Dim intSelectedPayk As Long
    Dim Rst As New ADODB.Recordset
    
    intSelectedPayk = -1
    
    With vsOwedFactors 'find all the factors which this payk have to pay them
        
        .Rows = 1
        lblNoOfFactors.Caption = 0
        lblShouldBePaid.Caption = 0
        
        vsFactorDetail.Rows = 1
        
        If vsOwedCustomers.Rows > 1 Then
            For i = 1 To vsOwedCustomers.Rows - 1
                If Val(vsOwedCustomers.TextMatrix(i, 1)) = -1 Then
                    intSelectedPayk = vsOwedCustomers.TextMatrix(i, 0)
                    Exit For
                End If
            Next i
        End If
        
        
        If Rst.State = 1 Then Rst.Close
        If intSelectedPayk = -1 Then Exit Sub
        
        ReDim Parameter(4) As Parameter
        Parameter(0) = GenerateInputParameter("@Customer", adBigInt, 8, intSelectedPayk)
        Parameter(1) = GenerateInputParameter("@AccountYear", adSmallInt, 2, Val(cmbSalMali.Text))
        Parameter(2) = GenerateInputParameter("@Branch", adInteger, 4, cmbBranch.ItemData(cmbBranch.ListIndex))
        Parameter(3) = GenerateInputParameter("@DateBefore", adVarWChar, 10, CStr(IIf(Trim(txtDateFrom.ClipText) = "", "", Trim(txtDateFrom.Text))))
        Parameter(4) = GenerateInputParameter("@DateAfter", adVarWChar, 10, CStr(IIf(Trim(txtDateTo.ClipText) = "", "", Trim(txtDateTo.Text))))
        Set Rst = RunParametricStoredProcedure2Rec("Get_CreditFactor", Parameter)
        
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
                .TextMatrix(i, 6) = Rst.Fields("SumPrice").Value
                .TextMatrix(i, 7) = Rst.Fields("Time").Value
                .TextMatrix(i, 8) = Rst.Fields("Date").Value
                .TextMatrix(i, 9) = Rst.Fields("Address").Value
                .TextMatrix(i, 10) = Rst.Fields("Branch").Value
                
                lblNoOfFactors.Caption = Val(lblNoOfFactors.Caption) + 1
                lblShouldBePaid.Caption = Val(lblShouldBePaid.Caption) + Rst.Fields("SumPrice").Value
                
                Rst.MoveNext
                i = i + 1
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



Private Sub chkNoPaykDelivery_Click()
    If chkNoPaykDelivery.Value = 1 Then
        With vsOwedCustomers
            For i = 1 To .Rows - 1
                .TextMatrix(i, 1) = ""
            Next i
        End With
        
        Label1(2).Caption = "·Ì”  ›«ò Ê—Â«Ì «—”«·Ì »œÊ‰ ÅÌò"
        
        FillvsOwedFactors
        
    Else
        lblNoOfFactors = 0
        lblShouldBePaid = 0
        Label1(2).Caption = ""
        vsOwedFactors.Rows = 1
        vsFactorDetail.Rows = 1
        
    End If
End Sub

Private Sub cmdCancel_Click()
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

Private Sub cmdPayAll_Click()

    Dim i As Integer
    Dim s As String
    Dim strPayk As String
    
        
        If vsOwedCustomers.Rows < 2 And chkNoPaykDelivery.Value = 0 Then
            Exit Sub
        End If
    
        With vsOwedCustomers
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
        With vsOwedFactors
        
            If .Rows < 2 Then Exit Sub
            
            For i = 1 To .Rows - 1
                    s = s & .TextMatrix(i, 0) & ","
            Next i
            If s = "" Then Exit Sub
            
            If chkNoPaykDelivery.Value = 1 Then
                frmMsg.fwlblMsg.Caption = "¬Ì« „«Ì·Ìœ  „«„ ›«ò Ê—Â«Ì »œÊ‰ ÅÌò —«  ”ÊÌÂ ‰„«ÌÌœ ø "
            ElseIf strPayk <> "" Then
                frmMsg.fwlblMsg.Caption = "¬Ì« „«Ì·Ìœ  „«„ ›«ò Ê—Â«Ì «Ì‰ „‘ —Ì —«  ”ÊÌÂ ‰„«ÌÌœ ø "
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
            
            ReDim Parameter(2) As Parameter
            Parameter(0) = GenerateInputParameter("@strSelectedFactors", adVarWChar, 4000, s)
            Parameter(1) = GenerateInputParameter("@UID", adInteger, 4, mvarCurUserNo)
            Parameter(2) = GenerateInputParameter("@AccountYear", adSmallInt, 2, Val(cmbSalMali.Text))
            RunParametricStoredProcedure "PayFactors_CustCredit_Balance", Parameter
                
                If strPayk <> "" Then
                    If InStr(1, s, ",") > 0 Then
                         lblMessage = "›«ò Ê—Â«Ì ‘„«—Â" & s & "  Ê”ÿ " & strPayk & " Å—œ«Œ  ‘œ "
                    Else
                         lblMessage = "›«ò Ê— ‘„«—Â" & s & "  Ê”ÿ " & strPayk & " Å—œ«Œ  ‘œ "
                    End If
                    
                    FillvsOwedCustomers
                    If vsOwedCustomers.Rows > 1 Then
                        For i = 1 To vsOwedCustomers.Rows - 1
                            If vsOwedCustomers.TextMatrix(i, 2) = strPayk Then
                                vsOwedCustomers.TextMatrix(i, 1) = -1
                            End If
                        Next i
                    End If
                    
                    FillvsOwedFactors
                    
                Else
                
                    If InStr(1, s, ",") > 0 Then
                         lblMessage = "›«ò Ê—Â«Ì ‘„«—Â" & s & " Å—œ«Œ  ‘œ "
                    Else
                         lblMessage = "›«ò Ê— ‘„«—Â" & s & " Å—œ«Œ  ‘œ "
                    End If
                            
                    FillvsOwedFactors
                    
                End If
                Timer1.Interval = 3000
                Timer1.Enabled = True
                  
                
        End With

End Sub

Private Sub Command1_Click()
   
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

Private Sub cmdUpdate_Click()
        vsOwedFactors.Rows = 1
        vsOwedCustomers.Rows = 1
        FillvsOwedCustomers

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
    
    VarActForm = Me.Name
    
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call PresetScreenSaver
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call PresetScreenSaver
End Sub

Private Sub Form_Unload(Cancel As Integer)

    If PosConnection.State = adStateOpen Then PosConnection.Close: Set PosConnection = Nothing
    
    SaveSetting strMainKey, Me.Name, "Left", Me.Left
    SaveSetting strMainKey, Me.Name, "Top", Me.Top

     VarActForm = ""
End Sub

Private Sub mnuReturnFromPaykAccount_Click()

    ReDim Parameter(0) As Parameter
    Parameter(0) = GenerateInputParameter("@strSelectedFactors", adVarWChar, 4000, vsOwedFactors.TextMatrix(vsOwedFactors.Row, 0))
    RunParametricStoredProcedure "Update_tFacM_InCharge_Null", Parameter
        
    FillvsOwedFactors

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
    
    
        If vsOwedCustomers.Rows < 2 Then
            Exit Sub
        End If
    
        With vsOwedCustomers
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
        With vsOwedFactors
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
            ReDim Parameter(2) As Parameter
            Parameter(0) = GenerateInputParameter("@strSelectedFactors", adVarWChar, 4000, s)
            Parameter(1) = GenerateInputParameter("@UID", adInteger, 4, mvarCurUserNo)
            Parameter(2) = GenerateInputParameter("@AccountYear", adSmallInt, 2, Val(cmbSalMali.Text))
            RunParametricStoredProcedure "PayFactors_CustCredit_Balance", Parameter
                
                If strPayk <> "" Then
                    If InStr(1, s, ",") > 0 Then
                         lblMessage = "›«ò Ê—Â«Ì ‘„«—Â" & s & "  Ê”ÿ " & strPayk & " Å—œ«Œ  ‘œ "
                    Else
                         lblMessage = "›«ò Ê— ‘„«—Â" & s & "  Ê”ÿ " & strPayk & " Å—œ«Œ  ‘œ "
                    End If
                    
                    FillvsOwedCustomers
                    If vsOwedCustomers.Rows > 1 Then
                        For i = 1 To vsOwedCustomers.Rows - 1
                            If vsOwedCustomers.TextMatrix(i, 2) = strPayk Then
                                vsOwedCustomers.TextMatrix(i, 1) = -1
                            End If
                        Next i
                    End If
                    
                    FillvsOwedFactors
                    
                Else
                    If InStr(1, s, ",") > 0 Then
                         lblMessage = "›«ò Ê—Â«Ì ‘„«—Â" & s & " Å—œ«Œ  ‘œ "
                    Else
                         lblMessage = "›«ò Ê— ‘„«—Â" & s & " Å—œ«Œ  ‘œ "
                    End If
                    
                    FillvsOwedFactors

                End If
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
    
    VarActForm = Me.Name
    
    With vsOwedCustomers
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
        .TextMatrix(0, 2) = "       ‰«„    "
    
        .AutoSearch = flexSearchFromCursor
    
    End With
    
    FillBranch
    
    With vsOwedFactors
        .Rows = 1
        .Cols = 11
        .ColAlignment(-1) = flexAlignRightCenter
        .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter
'        .ColWidth(1) = 500
        .ColDataType(1) = flexDTBoolean
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
        .TextMatrix(0, 9) = "¬œ—”"
        .TextMatrix(0, 10) = "‘⁄»Â"
        .ColFormat(6) = "###,###"
        
        .AutoSearch = flexSearchFromCursor
    
        Set Rst = RunStoredProcedure2RecordSet("Get_All_Branches")
        .ColComboList(10) = .BuildComboList(Rst, "nvcBranchName", "Branch")
    End With
    
    With vsFactorDetail
        .Rows = 1
        .Cols = 5
        .ColAlignment(-1) = flexAlignRightCenter
        .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter
        
        'set the headers of the columns
        .TextMatrix(0, 0) = "—œÌ›"
        .TextMatrix(0, 1) = " ⁄œ«œ"
        .TextMatrix(0, 2) = "‰«„ ò«·«"
        .TextMatrix(0, 3) = "›Ì"
        .TextMatrix(0, 4) = "ﬁÌ„ "
    
        .AutoSearch = flexSearchFromCursor
    
    End With
    FillSalMali
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

    txtDateFrom.Text = Mid(clsDate.shamsi(Date), 3, 2) & "/01" & "/01"
    txtDateTo.Text = Mid(clsDate.shamsi(Date), 3)
    
    cmdUpdate_Click
End Sub



Private Sub vsOwedFactors_KeyDown(KeyCode As Integer, Shift As Integer)
    
    FillvsFactorDetail
    
    If KeyCode <> 32 Then Exit Sub ' if the key is not space bar
    Dim i As Integer
    Dim S2 As String
    
    With vsOwedFactors
        If .Row > 0 And .Rows > 1 Then
        
            .Select .Row, 1
            .EditCell
            
        End If
    End With

    CalculateSelected
End Sub

Private Sub vsOwedFactors_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    FillvsFactorDetail
            
    Dim i As Integer
    Dim S2 As String
    
    With vsOwedFactors
        If .Col = 1 And .Row > 0 And .Rows > 1 And Button = 1 Then
            
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

Private Sub vsOwedFactors_SelChange()
  FillvsFactorDetail
End Sub

Private Sub vsOwedCustomers_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> 32 Then Exit Sub ' if the key is not space bar
    Dim i As Integer
    Dim S2 As String
    
    With vsOwedCustomers
        If .Col = 1 And .Row > 0 And .Rows > 1 Then
            For i = 1 To .Rows - 1
                If i <> .Row Then
                    .TextMatrix(i, 1) = ""
                End If
            Next i
            .Select .Row, .Col
            .EditCell
            
            If Val(.TextMatrix(.Row, 1)) = -1 Then
                chkNoPaykDelivery.Value = 0
                FillvsOwedFactors
                Label1(2).Caption = "·Ì”  ›«ò Ê—Â«Ì Å—œ«Œ  ‰‘œÂ  Ê”ÿ " & .TextMatrix(.Row, 2)
            Else
                vsOwedFactors.Rows = 1
                vsFactorDetail.Rows = 1
                Label1(2).Caption = ""
                lblNoOfFactors.Caption = 0
                lblShouldBePaid.Caption = 0
            End If
        End If
    End With

End Sub

Private Sub vsOwedCustomers_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Dim i As Integer
    
    With vsOwedCustomers
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
                FillvsOwedFactors
                Label1(2).Caption = "·Ì”  ›«ò Ê—Â«Ì «—”«· ‘œÂ »Â " & .TextMatrix(.Row, 2)
            Else
                vsOwedFactors.Rows = 1
                vsFactorDetail.Rows = 1
                Label1(2).Caption = ""
                lblNoOfFactors.Caption = 0
                lblShouldBePaid.Caption = 0
            End If
        End If
    End With

End Sub



