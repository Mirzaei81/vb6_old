VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{7AEDC602-D94C-11D1-BB7A-00E0290EA3C9}#1.0#0"; "ResizeKit.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmGarson 
   ClientHeight    =   9255
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12525
   Icon            =   "frmGarson.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9255
   ScaleWidth      =   12525
   Begin VB.CommandButton CmdCancel 
      BackColor       =   &H00404080&
      Cancel          =   -1  'True
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
      Left            =   240
      TabIndex        =   21
      Top             =   8520
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00808000&
      Caption         =   "Œ«·Ì ﬂ—œ‰ „Ì“Â«Ì Å—"
      BeginProperty Font 
         Name            =   "B Homa"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   2640
      TabIndex        =   19
      Top             =   7680
      Width           =   2295
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   480
      Top             =   7320
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
   Begin VB.Frame Frame1 
      Height          =   5625
      Left            =   150
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   600
      Width           =   9735
      Begin VB.CommandButton cmdCreditMove 
         BackColor       =   &H008080FF&
         Caption         =   "«‰ ﬁ«· »Â Õ”«» „‘ —Ì«‰ «⁄ »«—Ì"
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   11.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   0
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Top             =   4800
         Width           =   2295
      End
      Begin VB.CheckBox ChkDaily 
         Alignment       =   1  'Right Justify
         Caption         =   "„Ì“Â«Ì «„—Ê“"
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   11.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   525
         Left            =   7320
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   4440
         Width           =   2265
      End
      Begin VB.CheckBox chkPrint 
         Alignment       =   1  'Right Justify
         Caption         =   "ç«Å"
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   11.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   8400
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Top             =   5040
         Width           =   1215
      End
      Begin VB.CheckBox chkNoPaykDelivery 
         Alignment       =   1  'Right Justify
         Caption         =   "„Ì“Â«Ì »œÊ‰ ê«—”Ê‰"
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   11.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   525
         Left            =   6840
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   3840
         Width           =   2745
      End
      Begin VB.CommandButton cmdPayAll 
         BackColor       =   &H00000080&
         Caption         =   " ”ÊÌÂ Õ”«» ò·Ì"
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   11.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   2400
         MaskColor       =   &H000000C0&
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   5040
         Width           =   2115
      End
      Begin VB.CommandButton cmdPaySome 
         BackColor       =   &H000000C0&
         Caption         =   " ”ÊÌÂ Õ”«»"
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   11.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   4560
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   5040
         Width           =   1875
      End
      Begin VSFlex7LCtl.VSFlexGrid vsDeliveredFactors 
         Height          =   3105
         Left            =   360
         TabIndex        =   1
         Top             =   600
         Width           =   9315
         _cx             =   16431
         _cy             =   5477
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
         FocusRect       =   2
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
         FormatString    =   $"frmGarson.frx":A4C2
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
         Height          =   435
         Left            =   2160
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   4560
         Width           =   1965
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "„»·€ ›«ò Ê—Â«Ì «‰ Œ«» ‘œÂ"
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   11.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   435
         Left            =   3750
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   4680
         Visible         =   0   'False
         Width           =   2625
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "„»·€ ò· »œÂÌ"
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   11.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   435
         Left            =   4785
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   3780
         Width           =   1605
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
         Height          =   525
         Index           =   2
         Left            =   3480
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   120
         Width           =   5895
      End
      Begin VB.Label Lable1 
         Alignment       =   1  'Right Justify
         Caption         =   " ⁄œ«œ ò· ›«ò Ê—Â«"
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   11.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   435
         Index           =   1
         Left            =   4530
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   4260
         Width           =   1845
      End
      Begin VB.Label lblNoOfFactors 
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
         Height          =   435
         Left            =   2880
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   4260
         Width           =   1605
      End
      Begin VB.Label lblShouldBePaid 
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
         Height          =   435
         Left            =   3180
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   3780
         Width           =   1605
      End
   End
   Begin VB.Timer Timer1 
      Left            =   120
      Top             =   630
   End
   Begin VSFlex7LCtl.VSFlexGrid vsFactorDetail 
      Height          =   2895
      Left            =   5400
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   6240
      Width           =   6705
      _cx             =   11827
      _cy             =   5106
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
      FocusRect       =   2
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
   Begin VSFlex7LCtl.VSFlexGrid vsOwedPayks 
      Height          =   3615
      Left            =   9840
      TabIndex        =   0
      Top             =   1200
      Width           =   2655
      _cx             =   4683
      _cy             =   6376
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
      FocusRect       =   2
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
      OleObjectBlob   =   "frmGarson.frx":A5A1
      TabIndex        =   24
      Top             =   0
      Width           =   480
   End
   Begin VB.Label lblBarCode 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   435
      Left            =   1080
      RightToLeft     =   -1  'True
      TabIndex        =   25
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label LblAccountYear 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Nazanin"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   495
      Left            =   10320
      TabIndex        =   23
      Top             =   0
      Width           =   1815
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "œ—Ì«›  „»·€ ’Ê—  Õ”«» «“ ê«—”Ê‰"
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
      Left            =   4560
      RightToLeft     =   -1  'True
      TabIndex        =   22
      Top             =   0
      Width           =   4095
   End
   Begin VB.Label lblMessage 
      Alignment       =   1  'Right Justify
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
      Height          =   825
      Left            =   60
      TabIndex        =   17
      Top             =   6330
      Width           =   5325
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "·Ì”  ê«—”Ê‰Â«Ì »œÂò«—"
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
      Left            =   9960
      RightToLeft     =   -1  'True
      TabIndex        =   13
      Top             =   720
      Width           =   2295
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "«ﬁ·«„  ›«ò Ê—"
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
      Left            =   10200
      RightToLeft     =   -1  'True
      TabIndex        =   12
      Top             =   5760
      Width           =   1815
   End
End
Attribute VB_Name = "frmGarson"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim clsDate As New clsDate
Dim Incharge As EnumIncharge
Dim i As Integer
Dim Parameter() As Parameter
Private Const indexColTableNo As Integer = 13


Public Sub barcode()
    If Len(lblBarCode.Caption) = 12 Then
       lblBarCode.Caption = "0" + lblBarCode.Caption
    ElseIf Len(lblBarCode.Caption) = 13 Then
        If Left(lblBarCode.Caption, 1) <> "0" Or (Mid(lblBarCode.Caption, 2, 1) = "3" Or Mid(lblBarCode.Caption, 2, 1) = "9") Then
            lblBarCode.Caption = "0" + Left(lblBarCode.Caption, 12)
        End If
    End If
    Dim Rst As New ADODB.Recordset
    Select Case Left(lblBarCode.Caption, 3)
                
        Case EnumIncharge.Garson, 26
                                
            If Rst.State <> 0 Then Rst.Close
            
            ReDim Parameter(0) As Parameter
            
            Parameter(0) = GenerateInputParameter("@pPNo", adInteger, 4, Mid(lblBarCode.Caption, 4, 10))
            Set Rst = RunParametricStoredProcedure2Rec("Get_Per_By_pPNo", Parameter)
            
            If Not (Rst.EOF = True And Rst.BOF) Then
                intPpno = Val(Mid(lblBarCode.Caption, 4, 10))
            
                  
                
                For i = 1 To vsOwedPayks.Rows - 1
                    
                    If vsOwedPayks.TextMatrix(i, 0) = Val(Mid(lblBarCode.Caption, 4, 10)) Then
                        vsOwedPayks.TextMatrix(i, 1) = -1 ' True
                        vsOwedPayks.Row = i
                        
                        Timer1.Enabled = False
                        lblMessage = "ê«—”Ê‰" & " " & Val(Mid(lblBarCode.Caption, 4, 10)) & " " & vsOwedPayks.TextMatrix(i, 2)
                        
                        Timer1.Interval = 10000
                        Timer1.Enabled = True
                        
                        Exit For
                    End If
                Next i
            End If
    End Select
    Set Rst = Nothing
    
End Sub

Public Sub ExitForm()

    Unload Me
    
End Sub

Public Sub FillvsOwedPayks()

    Dim Rst As New ADODB.Recordset
    ReDim Parameter(1) As Parameter

    If Rst.State = 1 Then Rst.Close
        
    Parameter(0) = GenerateInputParameter("@Job", adInteger, 4, Incharge)
    Parameter(1) = GenerateInputParameter("@Branch", adInteger, 4, CurrentBranch)
    Set Rst = RunParametricStoredProcedure2Rec("GetNotPaidGarsons", Parameter)
        
   
    With vsOwedPayks
        .Rows = 1
        
        If Not (Rst.EOF = True And Rst.BOF = True) Then
        
            i = 1
            While Rst.EOF = False ' fill the Grid
                .Rows = .Rows + 1
                .TextMatrix(i, 0) = Rst.Fields("InCharge").Value
                .TextMatrix(i, 2) = Rst.Fields("nvcFirstName").Value & " " & Rst.Fields("nvcSurName").Value
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



Public Sub FillvsDeliveredFactors()

    Dim i As Integer
    Dim intSelectedPayk As Integer
    Dim Rst As New ADODB.Recordset
    
    intSelectedPayk = -1
    
    With vsDeliveredFactors 'find all the factors which this payk have to pay them
        
        .Rows = 1
        lblNoOfFactors.Caption = 0
        lblShouldBePaid.Caption = 0
        
        vsFactorDetail.Rows = 1
        
        
        If vsOwedPayks.Rows > 1 Then
            For i = 1 To vsOwedPayks.Rows - 1
                If Val(vsOwedPayks.TextMatrix(i, 1)) = -1 Then
                    intSelectedPayk = vsOwedPayks.TextMatrix(i, 0)
                    Exit For
                End If
            Next i
        End If
        
        
        If Rst.State = 1 Then Rst.Close
        ReDim Parameter(2) As Parameter
        If intSelectedPayk <> -1 Then
            Parameter(0) = GenerateInputParameter("@InCharge", adInteger, 4, intSelectedPayk)
        ElseIf intSelectedPayk = -1 And chkNoPaykDelivery.Value = 1 Then
            Parameter(0) = GenerateInputParameter("@InCharge", adInteger, 4, 0)
        Else
            Exit Sub
        End If
        Parameter(1) = GenerateInputParameter("@AccountYear", adSmallInt, 2, AccountYear)
        Parameter(2) = GenerateInputParameter("@Branch", adInteger, 4, CurrentBranch)
        Set Rst = RunParametricStoredProcedure2Rec("Get_TableFactor", Parameter)
        
        If Not (Rst.EOF = True And Rst.BOF = True) Then
        
            i = 1
            On Error Resume Next
            
            Dim VarToday As String
            VarToday = mvarDate
            While Rst.EOF = False 'fill the grid
                If ChkDaily.Value = False Or Rst.Fields("Date").Value = VarToday Then
                    .Rows = .Rows + 1
                    .TextMatrix(i, 0) = Rst.Fields("intSerialNo").Value
                    .TextMatrix(i, 2) = Rst.Fields("nvcFirstName").Value & " " & Rst.Fields("nvcSurname").Value
                    .TextMatrix(i, 3) = Val(Right(Rst.Fields("No").Value, 3))
                    .TextMatrix(i, 4) = Val(Right(Rst.Fields("TempNo").Value, 3))
                    If Rst.Fields("Code").Value = -1 Then
                       .TextMatrix(i, 5) = ""
                       .TextMatrix(i, 6) = ""
                   Else
                      .TextMatrix(i, 5) = Rst.Fields("Code").Value
                      .TextMatrix(i, 6) = Rst.Fields("Full Name").Value ' Rst.Fields("Name").Value & " " & Rst.Fields("Family")
                    End If
                    If Rst.Fields("Credit").Value > 0 Then
                       .TextMatrix(i, 7) = 1
                    End If
                    .TextMatrix(i, 8) = Rst.Fields("TableName").Value
                    .TextMatrix(i, 9) = Rst.Fields("SumPrice").Value
                    .TextMatrix(i, 10) = Rst.Fields("Time").Value
                    .TextMatrix(i, 11) = Rst.Fields("Date").Value
                    .TextMatrix(i, 12) = Rst.Fields("ShiftDescription").Value
                    .TextMatrix(i, indexColTableNo) = Rst!TableNo
                    
                    lblNoOfFactors.Caption = Val(lblNoOfFactors.Caption) + 1
                    lblShouldBePaid.Caption = Val(lblShouldBePaid.Caption) + Rst.Fields("SumPrice").Value
                    i = i + 1
                End If
                Rst.MoveNext
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

Public Sub FillvsFactorDetail() ' fills the detail of the current factor
    Dim i As Integer
    Dim intselFactor As Double
    
    Dim Rst As New ADODB.Recordset
    
    If Rst.State = 1 Then Rst.Close
    
    With vsFactorDetail
        If vsDeliveredFactors.Rows <= 1 Then Exit Sub ' if there is no factor in the grid
        ' if at least there is one , choose the current one
        intselFactor = vsDeliveredFactors.TextMatrix(vsDeliveredFactors.Row, 0)
        
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
        .AutoSizeMode = flexAutoSizeColWidth  ' set the collumns' width
        .AutoSize 0, .Cols - 1
    End With
    
    If Rst.State <> 0 Then Rst.Close
    Set Rst = Nothing
    

End Sub

Private Sub ChkDaily_Click()
    FillvsDeliveredFactors
End Sub

Private Sub chkNoPaykDelivery_Click()
    If chkNoPaykDelivery.Value = 1 Then
        With vsOwedPayks
            For i = 1 To .Rows - 1
                .TextMatrix(i, 1) = ""
            Next i
        End With
        
        Label1(2).Caption = "·Ì”  „Ì“Â«Ì »œÊ‰ ê«—”Ê‰"
        Dim s As String
                
        FillvsDeliveredFactors
        
    Else
        lblNoOfFactors = 0
        lblShouldBePaid = 0
        Label1(2).Caption = ""
        vsDeliveredFactors.Rows = 1
        vsFactorDetail.Rows = 1
        
    End If
End Sub
Private Sub CalculateSelected()
    
    Dim tempPrice As Double
    
    With vsDeliveredFactors
        lblSelected.Caption = ""
        If .Rows < 2 Then
            Exit Sub
        End If
        For i = 1 To .Rows - 1
            If Val(.TextMatrix(i, 1)) <> 0 Then
                tempPrice = tempPrice + Val(.TextMatrix(i, 9))
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


Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdCreditMove_Click()
    Dim i As Integer
    Dim s, S2 As String
    Dim strPayk As String
    
    If vsOwedPayks.Rows < 2 And chkNoPaykDelivery.Value = 0 Then
        Exit Sub
    End If

    With vsOwedPayks
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
    S2 = ""
    With vsDeliveredFactors
    
        For i = 1 To .Rows - 1
            If .TextMatrix(i, 1) <> "" And .TextMatrix(i, 7) <> "" Then '
                s = s & .TextMatrix(i, 0) & ","
                S2 = S2 & .TextMatrix(i, indexColTableNo) & ","
            End If
        Next i
        If s = "" Then ShowDisMessage " ›Ì‘ «‰ Œ«» ‰‘œÂ Ì« „‘ —Ì «⁄ »«—Ì ‰Ì”  ", 2000: Exit Sub
        
        frmMsg.fwlblMsg.Caption = "¬Ì« „«Ì·Ìœ ›«ò Ê—Â«Ì «‰ Œ«» ‘œÂ —« »Â Õ”«» „‘ —Ì «⁄ »«—Ì „‰ ﬁ· ‰„«∆Ìœ ø"
        
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
        ReDim Parameter(2) As Parameter
        
        Parameter(0) = GenerateInputParameter("@strSelectedFactors", adVarWChar, 4000, s)
        Parameter(1) = GenerateInputParameter("@strSelectedTables", adVarWChar, 4000, S2)
        Parameter(2) = GenerateInputParameter("@Branch", adInteger, 4, CurrentBranch)
        RunParametricStoredProcedure "PayFactors_TabletoCustCredit", Parameter
        
        If mdifrm.ClsActionLog.LogMoveTableToCustomCredit Then
            ReDim Parameter(2) As Parameter
            Parameter(0) = GenerateInputParameter("@strSelectedFactors", adVarWChar, 4000, s)
            Parameter(1) = GenerateInputParameter("@UID", adInteger, 4, mvarCurUserNo)
            Parameter(2) = GenerateInputParameter("@ActionCode", adInteger, 4, EnumActionLog.MoveTableToCustomCredit)
            RunParametricStoredProcedure "InsertHistory_Batch", Parameter
            
        End If
        
        If strPayk <> "" Then
        
            If InStr(1, s, ",") > 0 Then
                 lblMessage = "›«ò Ê—Â«Ì ‘„«—Â" & s & "»Â Õ”«» „‘ —Ì «⁄ »«—Ì „‰ ﬁ· ‘œ‰œ "
            Else
                 lblMessage = "›«ò Ê— ‘„«—Â" & s & "»Â Õ”«» „‘ —Ì «⁄ »«—Ì „‰ ﬁ· ‘œ"
            End If
            
            FillvsOwedPayks
            If vsOwedPayks.Rows > 1 Then
                For i = 1 To vsOwedPayks.Rows - 1
                    If vsOwedPayks.TextMatrix(i, 2) = strPayk Then
                        vsOwedPayks.TextMatrix(i, 1) = -1
                    End If
                Next i
            End If
            
            FillvsDeliveredFactors
            
        Else
            If InStr(1, s, ",") > 0 Then
                 lblMessage = "›«ò Ê—Â«Ì ‘„«—Â" & s & " Å—œ«Œ  ‘œ "
            Else
                 lblMessage = "›«ò Ê— ‘„«—Â" & s & " Å—œ«Œ  ‘œ "
            End If
            
            FillvsDeliveredFactors

        End If
        Timer1.Interval = 3000
        Timer1.Enabled = True
            
    End With
End Sub

Private Sub cmdPayAll_Click()

    Dim i As Integer
    Dim s, S2 As String
    Dim strPayk As String
    
        
        If vsOwedPayks.Rows < 2 And chkNoPaykDelivery.Value = 0 Then
            Exit Sub
        End If
    
        With vsOwedPayks
            For i = 1 To .Rows - 1
                If Val(.TextMatrix(i, 1)) = -1 Then
                    strPayk = .TextMatrix(i, 2)
                End If
            Next i
        End With
        
        
        s = ""
        S2 = ""
        With vsDeliveredFactors
        
            If .Rows < 2 Then Exit Sub
            
            For i = 1 To .Rows - 1
                    s = s & .TextMatrix(i, 0) & ","
                    S2 = S2 & .TextMatrix(i, indexColTableNo) & ","
            Next i
            If s = "" Then Exit Sub
            
            If chkNoPaykDelivery.Value = 1 Then
                frmMsg.fwlblMsg.Caption = "¬Ì« „«Ì·Ìœ  „«„ ›«ò Ê—Â«Ì »œÊ‰ ê«—”Ê‰ —«  ”ÊÌÂ ‰„«ÌÌœ ø "
            ElseIf strPayk <> "" Then
                frmMsg.fwlblMsg.Caption = "¬Ì« „«Ì·Ìœ  „«„ ›«ò Ê—Â«Ì «Ì‰ ê«—”Ê‰ —«  ”ÊÌÂ ‰„«ÌÌœ ø "
            Else
                Exit Sub
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
            If Len(s) > 4000 Then
                cmdPayAll.Enabled = False
                frmMsg.fwlblMsg.Caption = " ⁄œ«œ ›Ì‘ Â« »Ì‘ «“ Õœ „Ã«“ „Ì »«‘œ " & vbLf & "«“ ﬂ·Ìœ  ”ÊÌÂ Õ”«» «” ›«œÂ ﬂ‰Ìœ"
                frmMsg.fwBtn(0).ButtonType = flwButtonOk
                frmMsg.fwBtn(0).Caption = "ﬁ»Ê·"
                frmMsg.fwBtn(1).Visible = False
                frmMsg.Show vbModal
                Exit Sub
            End If
            S2 = Left(S2, Len(S2) - 1)
            If Len(S2) > 4000 Then
                cmdPayAll.Enabled = False
                frmMsg.fwlblMsg.Caption = " ⁄œ«œ ›Ì‘ Â« »Ì‘ «“ Õœ „Ã«“ „Ì »«‘œ " & vbLf & "«“ ﬂ·Ìœ  ”ÊÌÂ Õ”«» «” ›«œÂ ﬂ‰Ìœ"
                frmMsg.fwBtn(0).ButtonType = flwButtonOk
                frmMsg.fwBtn(0).Caption = "ﬁ»Ê·"
                frmMsg.fwBtn(1).Visible = False
                frmMsg.Show vbModal
                Exit Sub
            End If
              
            ReDim Parameter(2) As Parameter
            Parameter(0) = GenerateInputParameter("@strSelectedFactors", adVarWChar, 4000, s)
            Parameter(1) = GenerateInputParameter("@strSelectedTables", adVarWChar, 4000, S2)
            Parameter(2) = GenerateInputParameter("@UID", adInteger, 4, mvarCurUserNo)
            RunParametricStoredProcedure "PayFactors_Table", Parameter
            
            If chkNoPaykDelivery.Value = 0 Then
                If mdifrm.ClsActionLog.LogPayGarsonFactor Then
                    ReDim Parameter(2) As Parameter
                    Parameter(0) = GenerateInputParameter("@strSelectedFactors", adVarWChar, 4000, s)
                    Parameter(1) = GenerateInputParameter("@UID", adInteger, 4, mvarCurUserNo)
                    Parameter(2) = GenerateInputParameter("@ActionCode", adInteger, 4, EnumActionLog.PayGarsonFactor)
                    RunParametricStoredProcedure "InsertHistory_Batch", Parameter
                    
                End If
            Else
                If mdifrm.ClsActionLog.LogPayCustomerFactor Then
                    ReDim Parameter(2) As Parameter
                    Parameter(0) = GenerateInputParameter("@strSelectedFactors", adVarWChar, 4000, s)
                    Parameter(1) = GenerateInputParameter("@UID", adInteger, 4, mvarCurUserNo)
                    Parameter(2) = GenerateInputParameter("@ActionCode", adInteger, 4, EnumActionLog.PayCustomerFactor)
                    RunParametricStoredProcedure "InsertHistory_Batch", Parameter
                    
                End If
            End If
            If strPayk <> "" Then
                If InStr(1, s, ",") > 0 Then
                     lblMessage = "›«ò Ê—Â«Ì ‘„«—Â" & s & "  Ê”ÿ " & strPayk & " Å—œ«Œ  ‘œ "
                Else
                     lblMessage = "›«ò Ê— ‘„«—Â" & s & "  Ê”ÿ " & strPayk & " Å—œ«Œ  ‘œ "
                End If
                
                FillvsOwedPayks
                If vsOwedPayks.Rows > 1 Then
                    For i = 1 To vsOwedPayks.Rows - 1
                        If vsOwedPayks.TextMatrix(i, 2) = strPayk Then
                            vsOwedPayks.TextMatrix(i, 1) = -1
                        End If
                    Next i
                End If
                
                FillvsDeliveredFactors
                
            Else
            
                If InStr(1, s, ",") > 0 Then
                     lblMessage = "›«ò Ê—Â«Ì ‘„«—Â" & s & " Å—œ«Œ  ‘œ "
                Else
                     lblMessage = "›«ò Ê— ‘„«—Â" & s & " Å—œ«Œ  ‘œ "
                End If
                
                FillvsDeliveredFactors
                
            End If
            Timer1.Interval = 3000
            Timer1.Enabled = True
                  
                
        End With

End Sub

Private Sub Command1_Click()
    
    RunNonParametricStoredProcedure "Update_Table_Empty"
    frmMsg.fwlblMsg.Caption = " ﬂ·ÌÂ „Ì“Â« Œ«·Ì ‘œ‰œ  "
    frmMsg.fwBtn(0).Visible = False
    frmMsg.fwBtn(1).ButtonType = flwButtonOk
    frmMsg.fwBtn(1).Caption = "ﬁ»Ê·"
    frmMsg.Show vbModal

End Sub

Private Sub Form_Activate()

    Dim i As Integer
    AllButton vbOff, True
    mdifrm.Toolbar1.Buttons(23).Enabled = True
    mdifrm.Toolbar1.Buttons(24).Enabled = True
    mdifrm.Toolbar1.Buttons(25).Enabled = True
    mdifrm.Toolbar1.Buttons(26).Enabled = True
    mdifrm.Toolbar1.Buttons(27).Enabled = True
    
    VarActForm = Me.Name
    LblAccountYear.Caption = "”«· „«·Ì :" & CInt(AccountYear)

End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call PresetScreenSaver
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call PresetScreenSaver
End Sub

Private Sub Form_Unload(Cancel As Integer)
    VarActForm = ""
    If PosConnection.State = adStateOpen Then PosConnection.Close: Set PosConnection = Nothing

    SaveSetting strMainKey, Me.Name, "Left", Me.Left
    SaveSetting strMainKey, Me.Name, "Top", Me.Top
End Sub

Private Sub mnuReturnFromPaykAccount_Click()
    ReDim Parameter(0) As Parameter
    Parameter(0) = GenerateInputParameter("@strSelectedFactors", adVarWChar, 4000, vsDeliveredFactors.TextMatrix(vsDeliveredFactors.Row, 0))
    RunParametricStoredProcedure "Update_tFacM_InCharge_Null", Parameter
    
    FillvsDeliveredFactors
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
    Dim s, S2 As String
    Dim strPayk As String
    ReDim Parameter(0) As Parameter
    
    
        If vsOwedPayks.Rows < 2 And chkNoPaykDelivery.Value = 0 Then
            Exit Sub
        End If
    
        With vsOwedPayks
            For i = 1 To .Rows - 1
                If Val(.TextMatrix(i, 1)) = -1 Then
                    strPayk = .TextMatrix(i, 2)
                End If
            Next i
        End With
        
        If strPayk = "" And chkNoPaykDelivery = False Then
            Exit Sub
        End If
        
        s = ""
        S2 = ""
        With vsDeliveredFactors
            For i = 1 To .Rows - 1
                If Val(.TextMatrix(i, 1)) = -1 Then
                    s = s & .TextMatrix(i, 0) & ","
                    S2 = S2 & .TextMatrix(i, indexColTableNo) & ","
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
            S2 = Left(S2, Len(S2) - 1)
            
            
            ReDim Parameter(2) As Parameter
            Parameter(0) = GenerateInputParameter("@strSelectedFactors", adVarWChar, 4000, s)
            Parameter(1) = GenerateInputParameter("@strSelectedTables", adVarWChar, 4000, S2)
            Parameter(2) = GenerateInputParameter("@UID", adInteger, 4, mvarCurUserNo)
            RunParametricStoredProcedure "PayFactors_Table", Parameter
                
            If chkNoPaykDelivery.Value = 0 Then
                If mdifrm.ClsActionLog.LogPayGarsonFactor Then
                    ReDim Parameter(2) As Parameter
                    Parameter(0) = GenerateInputParameter("@strSelectedFactors", adVarWChar, 4000, s)
                    Parameter(1) = GenerateInputParameter("@UID", adInteger, 4, mvarCurUserNo)
                    Parameter(2) = GenerateInputParameter("@ActionCode", adInteger, 4, EnumActionLog.PayGarsonFactor)
                    RunParametricStoredProcedure "InsertHistory_Batch", Parameter
                    
                End If
            Else
                If mdifrm.ClsActionLog.LogPayCustomerFactor Then
                    ReDim Parameter(2) As Parameter
                    Parameter(0) = GenerateInputParameter("@strSelectedFactors", adVarWChar, 4000, s)
                    Parameter(1) = GenerateInputParameter("@UID", adInteger, 4, mvarCurUserNo)
                    Parameter(2) = GenerateInputParameter("@ActionCode", adInteger, 4, EnumActionLog.PayCustomerFactor)
                    RunParametricStoredProcedure "InsertHistory_Batch", Parameter
                    
                End If
            End If
            
            If chkPrint.Value = vbChecked Then
                
                CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\RepGarsonPaidBills.rpt"
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
                CrystalReport1.ReportTitle = clsArya.StationName
                CrystalReport1.Destination = crptToWindow 'crptToPrinter '
                
                Parameter(0) = GenerateInputParameter("@AttachedString", adVarWChar, 1000, s)
                
                CrystalReport1.ParameterFields(0) = CStr(Parameter(0).Name) & ";" & CStr(Parameter(0).Value) & ";" & "True"
  
                CrystalReport1.RetrieveDataFiles
                CrystalReport1.WindowShowPrintBtn = True
                CrystalReport1.WindowShowPrintSetupBtn = True
                CrystalReport1.WindowState = crptMaximized
                ODBCSetting clsArya.ServerName, clsArya.DbName
                CrystalReport1.Connect = CrystallConnection
                CrystalReport1.Action = 1
                If Screen.Width > 12000 Then
                    CrystalReport1.PageZoom (100)
                Else
                    CrystalReport1.PageZoom (75)
                End If

                chkPrint.Value = vbUnchecked
            End If
            
            If strPayk <> "" Then
                If InStr(1, s, ",") > 0 Then
                     lblMessage = "›«ò Ê—Â«Ì ‘„«—Â" & s & "  Ê”ÿ " & strPayk & " Å—œ«Œ  ‘œ "
                Else
                     lblMessage = "›«ò Ê— ‘„«—Â" & s & "  Ê”ÿ " & strPayk & " Å—œ«Œ  ‘œ "
                End If
                
                FillvsOwedPayks
                If vsOwedPayks.Rows > 1 Then
                    For i = 1 To vsOwedPayks.Rows - 1
                        If vsOwedPayks.TextMatrix(i, 2) = strPayk Then
                            vsOwedPayks.TextMatrix(i, 1) = -1
                        End If
                    Next i
                End If
                
                FillvsDeliveredFactors
                
            Else
                If InStr(1, s, ",") > 0 Then
                     lblMessage = "›«ò Ê—Â«Ì ‘„«—Â" & s & " Å—œ«Œ  ‘œ "
                Else
                     lblMessage = "›«ò Ê— ‘„«—Â" & s & " Å—œ«Œ  ‘œ "
                End If
                
                FillvsDeliveredFactors '(S)

            End If
            Timer1.Interval = 3000
            Timer1.Enabled = True
        End With
End Sub

Private Sub Form_Load()
    
    If ClsFormAccess.frmGarson = False Then
        Unload Me
        Exit Sub
    End If
    
    If intVersion = Min Or intVersion = Normal Then
        ShowDisMessage "›—„ ê«—”Ê‰ œ— ‰”ŒÂ Â«Ì ÅÌ‘—› Â Ê »«·« — ÊÃÊœ œ«—œ", 1500
        Unload Me
        Exit Sub
    End If
    
    CenterTop Me
    
    Incharge = Garson
    
''''    Dim obj As Object
''''    For Each obj In Forms
''''        If TypeOf obj Is Form Then
''''            If LCase(obj.Name) <> "mdifrm" And obj.Name <> Me.Name And (LCase(obj.Name) <> "frminvoice" Or LCase(obj.Name) <> "frminvoice_shop") Then
''''                Unload obj
''''            End If
''''        End If
''''
''''    Next obj
    
    VarActForm = Me.Name
    
    With vsOwedPayks
        .Rows = 1
        .Cols = 3
        .ColAlignment(-1) = flexAlignRightCenter
        .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter
        .ColWidth(1) = 500
        .ColDataType(1) = flexDTBoolean ' the data in this column is boolean
        .ColHidden(0) = True
        'set the headers of the columns
        .TextMatrix(0, 0) = "òœ ê«—”Ê‰"
        .TextMatrix(0, 1) = ""
        .TextMatrix(0, 2) = "‰«„ Ê ‰«„ Œ«‰Ê«œêÌ"
    
        .AutoSearch = flexSearchFromCursor
    
    End With
    
    With vsDeliveredFactors
        .Rows = 1
        .Cols = 14
        .ColAlignment(-1) = flexAlignCenterCenter
        .ColAlignment(6) = flexAlignRightCenter
        .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter
'        .ColWidth(1) = 500
        .ColDataType(1) = flexDTBoolean
        .ColDataType(7) = flexDTBoolean
        .ColHidden(0) = True
        .ColHidden(indexColTableNo) = True
        .ColDataType(indexColTableNo) = flexDTLong
        'set the headers of the columns
        .TextMatrix(0, 0) = "òœ ›Ì‘"
        .TextMatrix(0, 1) = " ”ÊÌÂ"
        .TextMatrix(0, 2) = "ê«—”Ê‰"
        .TextMatrix(0, 3) = "”—Ì«·"
        .TextMatrix(0, 4) = "‘„«—Â"
        .TextMatrix(0, 5) = "òœ"
        .TextMatrix(0, 6) = "„‘ —Ì"
        .TextMatrix(0, 7) = "«⁄ »«—Ì"
        .TextMatrix(0, 8) = "„Ì“"
        .TextMatrix(0, 9) = "„»·€"
        .TextMatrix(0, 10) = "”«⁄ "
        .TextMatrix(0, 11) = " «—ÌŒ"
        .TextMatrix(0, 12) = "‘Ì› "
        .TextMatrix(0, indexColTableNo) = "ﬂœ „Ì“"
        .AutoSearch = flexSearchFromCursor
    
        .ColFormat(9) = "###,###"
    End With
    
    With vsFactorDetail
        .Rows = 1
        .Cols = 5
        .ColAlignment(-1) = flexAlignCenterCenter
        .ColAlignment(2) = flexAlignRightCenter
        .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter
        
        'set the headers of the columns
        .TextMatrix(0, 0) = "—œÌ›"
        .TextMatrix(0, 1) = " ⁄œ«œ"
        .TextMatrix(0, 2) = "‰«„ ò«·«"
        .TextMatrix(0, 3) = "›Ì"
        .TextMatrix(0, 4) = "ﬁÌ„ "
    
        .AutoSearch = flexSearchFromCursor
        .ColFormat(3) = "###,###"
        .ColFormat(4) = "###,###"
    End With
        
    FillvsOwedPayks
    
    For Each Obj In Forms
        If TypeOf Obj Is Form And LCase(Obj.Name) = "frminvoice" Then
            lblBarCode.Caption = frmInvoice.lblBarCode.Caption
''''        ElseIf TypeOf obj Is Form And LCase(obj.Name) = "frminvoice_shop" Then
''''            lblBarCode.Caption = frmInvoice_Shop.lblBarCode.Caption
        End If
    Next Obj
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


    Me.barcode
    
End Sub


Private Sub vsDeliveredFactors_KeyDown(KeyCode As Integer, Shift As Integer)
    FillvsFactorDetail
    
    If KeyCode <> 32 Then Exit Sub ' if the key is not space bar
    Dim i As Integer
    Dim S2 As String
    
    With vsDeliveredFactors
        If .Row > 0 And .Rows > 1 Then
'            For i = 1 To .Rows - 1
'                If i <> .Row Then
'                    .TextMatrix(i, 1) = False
'                End If
'            Next i
            .Select .Row, 1
            .EditCell
        End If
    End With

    CalculateSelected
End Sub

Private Sub vsDeliveredFactors_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    FillvsFactorDetail
            
    Dim i As Integer
    Dim S2 As String
    
    With vsDeliveredFactors
        If .Col = 1 And .Row > 0 And .Rows > 1 And Button = 1 Then
'            For i = 1 To .Rows - 1
'                If i <> .Row Then
'                    .TextMatrix(i, 1) = False
'                End If
'            Next i
            .Select .Row, .Col
            .EditCell
            
            
        End If
    End With
    
    CalculateSelected

End Sub

Private Sub vsDeliveredFactors_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 And vsDeliveredFactors.MouseRow = vsDeliveredFactors.Row Then
        Me.PopupMenu PaykContextMenu
    End If
End Sub

Private Sub vsDeliveredFactors_SelChange()
    FillvsFactorDetail
End Sub

Private Sub vsOwedPayks_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> 32 Then Exit Sub ' if the key is not space bar
    Dim i As Integer
    Dim S2 As String
    
    With vsOwedPayks
        If .Row > 0 And .Rows > 1 Then
            For i = 1 To .Rows - 1
                If i <> .Row Then
                    .TextMatrix(i, 1) = ""
                End If
            Next i
            .Select .Row, 1
            .EditCell
            
            If Val(.TextMatrix(.Row, 1)) = -1 Then

                FillvsDeliveredFactors
                Label1(2).Caption = "·Ì”  ›«ò Ê—Â«Ì  ”ÊÌÂ Õ”«» ‰‘œÂ " & .TextMatrix(.Row, 2)
            Else
                vsDeliveredFactors.Rows = 1
                vsFactorDetail.Rows = 1
                Label1(2).Caption = ""
                lblNoOfFactors.Caption = 0
                lblShouldBePaid.Caption = 0
            End If
        End If
    End With

End Sub

Private Sub vsOwedPayks_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Dim i As Integer
    
    With vsOwedPayks
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
                FillvsDeliveredFactors
                Label1(2).Caption = "·Ì”  ›«ò Ê—Â«Ì  ”ÊÌÂ Õ”«» ‰‘œÂ " & .TextMatrix(.Row, 2)

            Else
                vsDeliveredFactors.Rows = 1
                vsFactorDetail.Rows = 1
                Label1(2).Caption = ""
                lblNoOfFactors.Caption = 0
                lblShouldBePaid.Caption = 0
            End If
        End If
    End With
End Sub
