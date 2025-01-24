VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{7AEDC602-D94C-11D1-BB7A-00E0290EA3C9}#1.0#0"; "ResizeKit.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmFacRecursive 
   ClientHeight    =   8835
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12675
   Icon            =   "frmFacRecursive.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8835
   ScaleWidth      =   12675
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
      Left            =   1440
      RightToLeft     =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   21
      Top             =   600
      Width           =   2475
   End
   Begin VB.Frame Frame1 
      Height          =   2295
      Left            =   8040
      TabIndex        =   9
      Top             =   6240
      Width           =   4455
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "„»·€ ›«ò Ê—Â«Ì „—ÃÊ⁄ ‘œÂ"
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
         Left            =   1860
         TabIndex        =   17
         Top             =   585
         Width           =   2445
      End
      Begin VB.Label lblCount 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "B Mitra"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   240
         TabIndex        =   16
         Top             =   120
         Width           =   1365
      End
      Begin VB.Label lblSum 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "B Mitra"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   240
         TabIndex        =   15
         Top             =   585
         Width           =   1425
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   " ⁄œ«œ ›«ò Ê—Â«Ì „—ÃÊ⁄ ‘œÂ"
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
         Left            =   1740
         TabIndex        =   14
         Top             =   120
         Width           =   2565
      End
      Begin VB.Label lblDailySum 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "B Mitra"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   120
         TabIndex        =   13
         Top             =   1740
         Width           =   1185
      End
      Begin VB.Label lblDailyCount 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "B Mitra"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   240
         TabIndex        =   12
         Top             =   1260
         Width           =   1065
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         Caption         =   "„»·€ ›«ò Ê—Â«Ì „—ÃÊ⁄ ‘œÂ «„—Ê“"
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
         Left            =   1290
         TabIndex        =   11
         Top             =   1740
         Width           =   3015
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         Caption         =   " ⁄œ«œ ›«ò Ê—Â«Ì „—ÃÊ⁄ ‘œÂ «„—Ê“"
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
         Left            =   1305
         TabIndex        =   10
         Top             =   1260
         Width           =   3015
      End
   End
   Begin VB.CommandButton cmdView 
      BackColor       =   &H008080FF&
      Caption         =   "‰„«Ì‘"
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
      TabIndex        =   5
      Top             =   480
      Width           =   1215
   End
   Begin VB.CheckBox ChkDailyView 
      Alignment       =   1  'Right Justify
      Caption         =   "›«ò Ê—Â«Ì „—ÃÊ⁄Ì «„—Ê“"
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
      Height          =   435
      Left            =   9240
      TabIndex        =   4
      Top             =   5760
      Value           =   1  'Checked
      Width           =   3225
   End
   Begin VSFlex7LCtl.VSFlexGrid vsFactorItems 
      Height          =   2265
      Left            =   120
      TabIndex        =   0
      Top             =   6240
      Width           =   7755
      _cx             =   13679
      _cy             =   3995
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "B Nazanin"
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
      BackColorFixed  =   -2147483645
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483643
      BackColorAlternate=   12648447
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
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   10
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
   Begin VSFlex7LCtl.VSFlexGrid vsRefferedFactors 
      Height          =   4395
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   12435
      _cx             =   21934
      _cy             =   7752
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "B Nazanin"
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
      BackColorBkg    =   -2147483643
      BackColorAlternate=   12648447
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
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   10
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
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   120
      Top             =   600
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
   Begin MSMask.MaskEdBox txtDateFrom 
      Height          =   465
      Left            =   7560
      TabIndex        =   18
      Top             =   600
      Width           =   1755
      _ExtentX        =   3096
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
   Begin MSMask.MaskEdBox txtDateTo 
      Height          =   465
      Left            =   4800
      TabIndex        =   19
      Top             =   600
      Width           =   1755
      _ExtentX        =   3096
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
   Begin RESIZEKITLibCtl.ResizeKit ResizeKit1 
      Height          =   480
      Left            =   0
      OleObjectBlob   =   "frmFacRecursive.frx":A4C2
      TabIndex        =   20
      Top             =   0
      Width           =   480
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   " : ‘⁄»Â"
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
      Left            =   3840
      TabIndex        =   22
      Top             =   600
      Width           =   855
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "›«ò Ê—Â«Ì „—ÃÊ⁄Ì"
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
      Height          =   495
      Left            =   4680
      TabIndex        =   8
      Top             =   0
      Width           =   2655
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "«“  «—ÌŒ"
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
      Left            =   9240
      TabIndex        =   7
      Top             =   600
      Width           =   945
   End
   Begin VB.Label Label3 
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
      Left            =   6480
      TabIndex        =   6
      Top             =   600
      Width           =   825
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "›«ò Ê—Â«Ì „—ÃÊ⁄ ‘œÂ"
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
      Height          =   465
      Left            =   10320
      TabIndex        =   3
      Top             =   720
      Width           =   2055
   End
   Begin VB.Label lblFactorDetail 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "—Ì“ «ﬁ·«„ ›«ò Ê—"
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
      Height          =   465
      Left            =   5160
      TabIndex        =   2
      Top             =   5760
      Width           =   2415
   End
End
Attribute VB_Name = "frmFacRecursive"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim clsDate As New clsDate
Dim i As Integer
Dim Parameter() As Parameter

Public Sub ExitForm()

    Unload Me

End Sub
Private Sub FillBranch()
    Dim rctmp As New ADODB.Recordset
    cmbBranch.Clear
    cmbBranch.AddItem "Â„Â ‘⁄»« "
    cmbBranch.ItemData(cmbBranch.NewIndex) = 0
    Set rctmp = RunStoredProcedure2RecordSet("Get_All_Branches")
    Do While rctmp.EOF = False
        cmbBranch.AddItem rctmp!nvcBranchName
        cmbBranch.ItemData(cmbBranch.NewIndex) = rctmp!Branch
        rctmp.MoveNext
    Loop
    rctmp.Close
    If cmbBranch.ListCount > 0 Then cmbBranch.ListIndex = 0
    
End Sub

Public Sub FillvsRefferedFactors()
    
    Dim Rst As New ADODB.Recordset

    If Rst.State = 1 Then Rst.Close
    
    lblDailyCount.Caption = 0
    lblDailySum.Caption = 0
    LblCount.Caption = 0
    lblSum.Caption = 0

    With vsRefferedFactors
    
        ReDim Parameter(3) As Parameter
        
        Parameter(0) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
        Parameter(1) = GenerateInputParameter("@DateAfter", adVarWChar, 20, CStr(IIf(Trim(txtDateFrom.ClipText) = "", "", Trim(txtDateFrom.Text))))
        Parameter(2) = GenerateInputParameter("@DateBefore", adVarWChar, 20, CStr(IIf(Trim(txtDateTo.ClipText) = "", "", Trim(txtDateTo.Text))))
        Parameter(3) = GenerateInputParameter("@Branch", adInteger, 4, cmbBranch.ItemData(cmbBranch.ListIndex))
        
        Set Rst = RunParametricStoredProcedure2Rec("Get_RefferedFactors", Parameter)
        .Rows = 1
        
        If Not (Rst.EOF = True And Rst.BOF = True) Then
        
            i = 1
            Dim VarToday As String
            VarToday = mvarDate
            While Rst.EOF = False
                If ChkDailyView.Value = False Or Rst.Fields("Date").Value = VarToday Then
                    .Rows = .Rows + 1
                    .TextMatrix(i, 0) = i
                    .TextMatrix(i, 1) = Rst.Fields("intSerialNo").Value
                    .TextMatrix(i, 2) = Rst.Fields("NO").Value
                    .TextMatrix(i, 3) = Rst.Fields("nvcFirstName").Value & " " & Rst.Fields("nvcSurName").Value
                    .TextMatrix(i, 4) = Rst.Fields("Status").Value ' Val(Right(Rst.Fields("No").Value, 3))
                    .TextMatrix(i, 5) = Rst.Fields("SumPrice").Value
                    .TextMatrix(i, 6) = Rst.Fields("ShiftDescription").Value
                    .TextMatrix(i, 7) = Rst.Fields("Date").Value
                    .TextMatrix(i, 8) = Rst.Fields("Time").Value
                    .TextMatrix(i, 9) = Rst.Fields("DiscountTotal").Value
                    .TextMatrix(i, 10) = Rst.Fields("CarryFeeTotal").Value
                    .TextMatrix(i, 11) = Rst.Fields("ServiceTotal").Value
                    .TextMatrix(i, 12) = Rst.Fields("PackingTotal").Value
                    .TextMatrix(i, 13) = Rst.Fields("Branch").Value
                    
                    If Rst.Fields("Date").Value = VarToday Then
                        lblDailySum.Caption = lblDailySum.Caption + Rst.Fields("SumPrice").Value
                        lblDailyCount.Caption = lblDailyCount.Caption + 1
                    End If
                    i = i + 1
                End If
                Rst.MoveNext
            Wend
            If .Rows > 1 Then

                LblCount.Caption = .Aggregate(flexSTCount, .FixedRows, 5, .Rows - 1, 5)
                lblSum.Caption = .Aggregate(flexSTSum, .FixedRows, 5, .Rows - 1, 5)
            End If
        End If
        .AutoSizeMode = flexAutoSizeColWidth
        .AutoSize 0, .Cols - 1
    End With
    
    If Rst.State <> 0 Then Rst.Close
    Set Rst = Nothing
End Sub
Public Sub FillvsFactorItems()
    
    Dim i As Integer
    Dim intselFactor As Long
    
    Dim Rst As New ADODB.Recordset
    
    If Rst.State = 1 Then Rst.Close
    lblFactorDetail.Caption = ""
    
    With vsFactorItems
        .Rows = 1
        
        If vsRefferedFactors.Rows <= 1 Then Exit Sub
        intselFactor = vsRefferedFactors.TextMatrix(vsRefferedFactors.Row, 1)
        
        ReDim Parameter(2) As Parameter
        
        Parameter(0) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
        Parameter(1) = GenerateInputParameter("@intSerialNo", adBigInt, 8, intselFactor)
        Parameter(2) = GenerateInputParameter("@Branch", adInteger, 4, vsRefferedFactors.ValueMatrix(vsRefferedFactors.Row, 13))
        
        Set Rst = RunParametricStoredProcedure2Rec("Get_Factor_Detail", Parameter)
        
        If Not (Rst.EOF = True And Rst.BOF = True) Then
            lblFactorDetail.Caption = "—Ì“ «ﬁ·«„ ›«ò Ê— " & vsRefferedFactors.TextMatrix(vsRefferedFactors.Row, 2)
'            Rst.moveFirst
            i = 1
            While Rst.EOF = False ' fill the grid
                .Rows = .Rows + 1
                .TextMatrix(i, 0) = Rst.Fields("intRow").Value
                .TextMatrix(i, 1) = Rst.Fields("Amount").Value
                .TextMatrix(i, 2) = Rst.Fields("Name").Value
                .TextMatrix(i, 3) = Rst.Fields("FeeUnit").Value
                .TextMatrix(i, 4) = Rst.Fields("FeeUnit").Value * Rst.Fields("Amount").Value
                .TextMatrix(i, 5) = Rst.Fields("ServePlace").Value
                
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

Private Sub ChkDailyView_Click()
    FillvsRefferedFactors
    FillvsFactorItems
End Sub

Private Sub cmbBranch_Change()
    
    ChkDailyView_Click
    FillvsRefferedFactors

End Sub

Private Sub cmbBranch_Click()
    cmbBranch_Change
End Sub

Private Sub cmdView_Click()
    ChkDailyView.Value = False
    FillvsRefferedFactors
    FillvsFactorItems
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

    If ClsFormAccess.frmFacRecursive = False Then
        Unload Me
        Exit Sub
    End If

    Dim Rst As New ADODB.Recordset
    Dim s As String
    
    CenterTop Me
    
    VarActForm = Me.Name
   
    FillBranch
    
    With vsRefferedFactors
        .Rows = 1
        .Cols = 14
'        .ColAlignment(-1) = flexAlignRightCenter
'        .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter
        .ColHidden(1) = True
        
        s = ""
        ReDim Parameter(1) As Parameter
        Parameter(0) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
        Parameter(1) = GenerateInputParameter("@ReturnType", adInteger, 4, 0)
        Set Rst = RunParametricStoredProcedure2Rec("Get_All_tStatusType", Parameter)
        s = .BuildComboList(Rst, "NvcDescription", "intStatusNo")
        .ColComboList(3) = s
       ' .ColComboList(4) = "#1;›«ò Ê— Œ—Ìœ|#2;›«ò Ê— ›—Ê‘"
        .TextMatrix(0, 0) = ""
        .TextMatrix(0, 1) = "”—Ì«· ›Ì‘"
        .TextMatrix(0, 2) = "òœ ›Ì‘"
        .TextMatrix(0, 3) = "‰«„ ò«—»—"
        .TextMatrix(0, 4) = "‰Ê⁄ ›«ò Ê—"
        .TextMatrix(0, 5) = "Ã„⁄"
        .TextMatrix(0, 6) = "‘Ì› "
        .TextMatrix(0, 7) = " «—ÌŒ"
        .TextMatrix(0, 8) = "”«⁄ "
        .TextMatrix(0, 9) = " Œ›Ì›"
        .TextMatrix(0, 10) = "Â“Ì‰Â Õ„·"
        .TextMatrix(0, 11) = "”—ÊÌ”"
        .TextMatrix(0, 12) = "Â“Ì‰Â »” Â »‰œÌ"
        .TextMatrix(0, 13) = "‘⁄»Â"
        .AutoSearch = flexSearchFromCursor
        .ColAlignment(-1) = flexAlignCenterCenter
        .ColAlignment(3) = flexAlignRightCenter
    
        Set Rst = RunStoredProcedure2RecordSet("Get_All_Branches")
        .ColComboList(13) = .BuildComboList(Rst, "nvcBranchName", "Branch")
    
    End With
    
    
    With vsFactorItems
        .Rows = 1
        .Cols = 6
'        .ColAlignment(-1) = flexAlignRightCenter
'        .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter
        .TextMatrix(0, 0) = "—œÌ›"
        .TextMatrix(0, 1) = "„ﬁœ«—"
        .TextMatrix(0, 2) = "‰«„ ò«·«"
        .TextMatrix(0, 3) = "›Ì"
        .TextMatrix(0, 4) = "ﬁÌ„ "
        .TextMatrix(0, 5) = "„Õ· ”—Ê"
        .ColAlignment(-1) = flexAlignCenterCenter
        .ColAlignment(2) = flexAlignRightCenter
        
        Dim strTemp As String
        
        If Rst.State = 1 Then Rst.Close
        
        ReDim Parameter(0) As Parameter
        Parameter(0) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
        Set Rst = RunParametricStoredProcedure2Rec("Get_Serveplace", Parameter)
        
        strTemp = .BuildComboList(Rst, "Description", "intServePlace")
        .ColComboList(5) = strTemp
        If Rst.State <> 0 Then Rst.Close
        
        
        Set Rst = Nothing
        
        .AutoSearch = flexSearchFromCursor
    
    End With
    
    txtDateFrom.Text = mvarDate
    txtDateTo.Text = Mid(clsDate.shamsi(Date), 3)
    FillvsRefferedFactors
    FillvsFactorItems
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

    SaveSetting strMainKey, Me.Name, "Left", Me.Left
    SaveSetting strMainKey, Me.Name, "Top", Me.Top



End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call PresetScreenSaver
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call PresetScreenSaver
End Sub

Private Sub ResizeKit1_ExitResize(ByVal XScale As Double, ByVal YScale As Double)

    If formloadFlag = True Then
        SaveSetting strMainKey, Me.Name, "Height", Me.Height
        SaveSetting strMainKey, Me.Name, "Width", Me.Width
    End If

End Sub

Private Sub vsRefferedFactors_SelChange()
    FillvsFactorItems
End Sub

Public Sub Printing()
    
    ReDim Parameter(4) As Parameter
    Parameter(0) = GenerateInputParameter("@SystemDate", adVarWChar, 50, Mid(clsDate.shamsi(Date), 3, 8))
    Parameter(1) = GenerateInputParameter("@SystemDay", adVarWChar, 50, clsDate.Find_DayOfWeek(Weekday(Date, vbSaturday)))
    Parameter(2) = GenerateInputParameter("@SystemTime", adVarWChar, 50, Mid(Str(time), 1, 5))
    Parameter(3) = GenerateInputParameter("@DateAfter", adVarWChar, 20, CStr(IIf(Trim(txtDateFrom.ClipText) = "", "", Trim(txtDateFrom.Text))))
    Parameter(4) = GenerateInputParameter("@DateBefore", adVarWChar, 20, CStr(IIf(Trim(txtDateTo.ClipText) = "", "", Trim(txtDateTo.Text))))
    
    CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\RepRefferedFactor_A4.rpt"
    
    Dim fileSystem As New FileSystemObject
    Dim IsFileExist As Boolean
    IsFileExist = fileSystem.FileExists(CrystalReport1.ReportFileName)
    If IsFileExist = False Then
        frmDisMsg.lblMessage = " ›«Ì· " & vbLf & Mid(CrystalReport1.ReportFileName, Len(CrystalReport1.ReportFileName) - 36) & " ÅÌœ« ‰‘œ "
        frmDisMsg.Timer1.Interval = 3000
        frmDisMsg.Timer1.Enabled = True
        frmDisMsg.Show vbModal
        Exit Sub
    End If
    
    CrystalReport1.ReportTitle = "ê“«—‘ «“ ›Ì‘ Â«Ì „—ÃÊ⁄Ì"
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

