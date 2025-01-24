VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{7AEDC602-D94C-11D1-BB7A-00E0290EA3C9}#1.0#0"; "ResizeKit.ocx"
Object = "{75D4F767-8785-11D3-93AD-0000832EF44D}#2.14#0"; "FAST2001.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmMojodiControl 
   ClientHeight    =   9360
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15105
   Icon            =   "frmMojodiControl.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9360
   ScaleWidth      =   15105
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
      Left            =   12360
      MultiLine       =   -1  'True
      TabIndex        =   26
      Top             =   2640
      Width           =   2145
   End
   Begin VB.Frame Frame2 
      Height          =   840
      Left            =   5925
      TabIndex        =   3
      Top             =   2160
      Width           =   3255
      Begin VB.CheckBox CheckOrder 
         Alignment       =   1  'Right Justify
         Caption         =   "›ﬁÿ ﬂ«·«Â«Ì »Â ‰ﬁÿÂ ”›«—‘ —”ÌœÂ"
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
         Height          =   615
         Left            =   120
         TabIndex        =   4
         Top             =   120
         Width           =   3015
      End
   End
   Begin VB.CheckBox CheckFirstMojodi 
      Alignment       =   1  'Right Justify
      Caption         =   "›ﬁÿ ﬂ«·«Â«Ì »« „ÊÃÊœÌ «Ê·ÌÂ"
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
      Height          =   345
      Left            =   9480
      TabIndex        =   11
      Top             =   2640
      Width           =   2535
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
      ForeColor       =   &H8000000C&
      Height          =   960
      Left            =   5925
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   360
      Width           =   3255
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
         Left            =   240
         RightToLeft     =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   9
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
      ForeColor       =   &H8000000C&
      Height          =   960
      Left            =   5925
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   1200
      Width           =   3255
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
         TabIndex        =   7
         Top             =   360
         Width           =   2475
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
      Height          =   2010
      Left            =   12360
      RightToLeft     =   -1  'True
      Style           =   1  'Checkbox
      TabIndex        =   1
      Top             =   600
      Width           =   2655
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
      Height          =   2010
      Left            =   9360
      RightToLeft     =   -1  'True
      Style           =   1  'Checkbox
      TabIndex        =   0
      Top             =   600
      Width           =   2745
   End
   Begin VSFlex7LCtl.VSFlexGrid vsGood 
      Height          =   5940
      Left            =   0
      TabIndex        =   2
      Top             =   3240
      Width           =   14985
      _cx             =   26432
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
      BackColorFixed  =   12648384
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
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   3
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
      FormatString    =   $"frmMojodiControl.frx":A4C2
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
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
   Begin FLWCtrls.FWLabel3D fwlblMode 
      Height          =   495
      Left            =   0
      Top             =   0
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
      Left            =   0
      OleObjectBlob   =   "frmMojodiControl.frx":A6A6
      TabIndex        =   13
      Top             =   0
      Width           =   480
   End
   Begin VB.CheckBox chkActiveGood 
      Alignment       =   1  'Right Justify
      Caption         =   "›ﬁÿ ﬂ«·«Â«Ì œ— ê—œ‘"
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
      Height          =   615
      Left            =   6240
      RightToLeft     =   -1  'True
      TabIndex        =   15
      Top             =   2160
      Visible         =   0   'False
      Width           =   2055
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   720
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
   Begin VB.Frame frmMenuGoodFirst 
      Height          =   2535
      Left            =   120
      TabIndex        =   12
      Top             =   480
      Width           =   5535
      Begin VB.ComboBox cmbSalMali 
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
         TabIndex        =   20
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton cmdInventoryGood_Delete 
         Caption         =   " Õ–› ò«·«Â« «“ «‰»«—"
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   19
         Top             =   210
         Width           =   1335
      End
      Begin VB.CommandButton cmdInventoryGood_Add 
         Caption         =   " «÷«›Â ò—œ‰ ‰«„ ò«·«Â« »Â «‰»«—"
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   1560
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   210
         Width           =   1575
      End
      Begin VB.CommandButton StoreDataUpdate 
         Caption         =   "»Â —Ê“ —”«‰Ì „ÊÃÊœÌ ﬂ«·«Â«Ì  Œ—Ìœ‰Ì - ›—ÊŒ ‰Ì"
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   360
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   1000
         Width           =   2295
      End
      Begin FLWCtrls.FWProgressBar FWProgressBar1 
         Height          =   375
         Left            =   120
         Top             =   2100
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   661
         BorderStyle     =   10
      End
      Begin MSMask.MaskEdBox txtDateTo 
         Height          =   465
         Left            =   3240
         TabIndex        =   21
         Top             =   1560
         Width           =   1335
         _ExtentX        =   2355
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
      Begin MSMask.MaskEdBox txtDateFrom 
         Height          =   465
         Left            =   3240
         TabIndex        =   22
         Top             =   907
         Width           =   1335
         _ExtentX        =   2355
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
      Begin VB.Label Label4 
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
         Left            =   4320
         RightToLeft     =   -1  'True
         TabIndex        =   25
         Top             =   1560
         Width           =   1065
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
         Left            =   4320
         RightToLeft     =   -1  'True
         TabIndex        =   24
         Top             =   855
         Width           =   1065
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "”«· „«·Ì"
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
         Left            =   4320
         RightToLeft     =   -1  'True
         TabIndex        =   23
         Top             =   240
         Width           =   1065
      End
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
      Height          =   435
      Left            =   10080
      RightToLeft     =   -1  'True
      TabIndex        =   17
      Top             =   240
      Width           =   2025
   End
   Begin VB.Label lblGoodLevel1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H000040C0&
      BackStyle       =   0  'Transparent
      Caption         =   "»Œ‘ Â«"
      BeginProperty Font 
         Name            =   "Titr"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   12480
      RightToLeft     =   -1  'True
      TabIndex        =   16
      Top             =   240
      Width           =   2415
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
      Height          =   315
      Left            =   14430
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   2640
      Width           =   585
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ﬂ‰ —· „ÊÃÊœÌ ﬂ«·« Â«Ì Œ—Ìœ‰Ì Ê ›—ÊŒ ‰Ì"
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
      Left            =   5400
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   -120
      Width           =   4455
   End
End
Attribute VB_Name = "frmMojodiControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private MyFormAddEditMode As EnumAddEditMode
Dim tmpTextMatrix As String
Dim Parameter() As Parameter
Dim clsDate As New clsDate
Dim i As Integer
    
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

    lstGoodLevel1.Clear
    lstGoodLevel2.Clear
    vsGood.Rows = 1
    If cmbInventory.ListIndex <> -1 And cmbBranch.ListIndex <> -1 Then
        FillLstGoodLevel1
    End If
End Sub

Public Sub FillLstGoodLevel1() ' it fills the lstGoodLevel1 using table tgoodlevel1
    Dim Rst As New ADODB.Recordset
    
    lstGoodLevel1.Clear
    lstGoodLevel2.Clear
    vsGood.Rows = 1
    
    MyFormAddEditMode = EnumAddEditMode.ViewMode
    SetFirstToolBar
    
    ReDim Parameter(1) As Parameter
    Parameter(0) = GenerateInputParameter("@Branch", adInteger, 4, cmbBranch.ItemData(cmbBranch.ListIndex))
    Parameter(1) = GenerateInputParameter("@InventoryNo", adInteger, 4, cmbInventory.ItemData(cmbInventory.ListIndex))
    Set Rst = RunParametricStoredProcedure2Rec("Get_Segment_Level1", Parameter)
        
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
    Dim Rst As New ADODB.Recordset
    Dim Rst2 As New ADODB.Recordset
    
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
    Dim level1 As Integer
    If intSelectedLevel1 <> -1 And intSelectedLevel2 <> -1 Then
        level1 = lstGoodLevel1.ItemData(intSelectedLevel1)
        strSelectedLevels = Right(strSelectedLevels, Len(strSelectedLevels) - 1)
    ElseIf intSelectedLevel1 <> -1 And intSelectedLevel2 = -1 Then  'Or intSelectedLevel2 = -1
       level1 = lstGoodLevel1.ItemData(intSelectedLevel1)
       strSelectedLevels = ""
    Else
        strSelectedLevels = ""
        level1 = -1
    End If
    ReDim Parameter(10) As Parameter
    Parameter(0) = GenerateInputParameter("@Level1", adInteger, 4, level1)
    Parameter(1) = GenerateInputParameter("@strSelectedLevels", adVarWChar, 4000, strSelectedLevels)
    Parameter(2) = GenerateInputParameter("@Type", adInteger, 4, EnumGoodType.forBuySale)
    Parameter(3) = GenerateInputParameter("@InventoryNo", adInteger, 4, cmbInventory.ItemData(cmbInventory.ListIndex))
    Parameter(4) = GenerateInputParameter("@Branch", adInteger, 4, cmbBranch.ItemData(cmbBranch.ListIndex))
    Parameter(5) = GenerateInputParameter("@AccountYear", adSmallInt, 2, Val(cmbSalMali.Text))
    Parameter(6) = GenerateInputParameter("@CheckNotZeroMojodi", adInteger, 4, 0)
    Parameter(7) = GenerateInputParameter("@CheckFirstMojodi", adInteger, 4, CheckFirstMojodi.Value)
    Parameter(8) = GenerateInputParameter("@CheckOrder", adInteger, 4, CheckOrder.Value)
    Parameter(9) = GenerateInputParameter("@Flag", adInteger, 4, 0)
    Parameter(10) = GenerateInputParameter("@SortItem", adInteger, 4, 1)
    
    Set Rst = RunParametricStoredProcedure2Rec("Get_tblTotal_tGood_By_Prams", Parameter)
      
    
    If (Rst.EOF = True And Rst.BOF = True) Then Exit Sub
    
    With vsGood
        
        i = 1
        
        While Rst.EOF = False
            'If CheckFirstMojodi.Value = 0 Or (CheckFirstMojodi.Value = 1 And Rst.Fields("FirstMojodi").Value > 0) Then
                 .Rows = .Rows + 1
                 .TextMatrix(i, 0) = i
                 .TextMatrix(i, 1) = Rst.Fields("GoodCode").Value
                 .TextMatrix(i, 2) = Left(Rst.Fields("Name").Value, 40)
                 .TextMatrix(i, 3) = Rst.Fields("Barcode").Value
                 .TextMatrix(i, 4) = Rst.Fields("CompDes").Value
                 If Rst.Fields("FirstMojodi").Value >= 0 Then
                    .TextMatrix(i, 5) = Rst.Fields("FirstMojodi").Value
                 Else
                    .TextMatrix(i, 5) = -Rst.Fields("FirstMojodi").Value & "-"
                 End If
                 .TextMatrix(i, 6) = Rst.Fields("BuyAmount").Value
                 .TextMatrix(i, 7) = Val(Format(Rst.Fields("SaleAmount").Value, "##.000"))
                 .TextMatrix(i, 8) = Val(Format(Rst.Fields("LossAmount").Value, "##.000"))
                 If Rst.Fields("Mojodi").Value >= 0 Then
                     If Rst.Fields("Mojodi").Value <> Int(Rst.Fields("Mojodi").Value) Then
                         .TextMatrix(i, 9) = Format(Rst.Fields("Mojodi").Value, "##.000")
                         .TextMatrix(i, 9) = Val(.TextMatrix(i, 9)) ' Delete Last Zeros
                     Else
                          .TextMatrix(i, 9) = Rst.Fields("Mojodi").Value
                     End If
                 Else
                     If Rst.Fields("Mojodi").Value <> Int(Rst.Fields("Mojodi").Value) Then
                         .TextMatrix(i, 9) = -Format(Rst.Fields("Mojodi").Value, "##.000")
                         .TextMatrix(i, 9) = Val(.TextMatrix(i, 9)) & "-" ' Delete Last Zeros
                     Else
                          .TextMatrix(i, 9) = -Rst.Fields("Mojodi").Value & "-"
                     End If
                 End If
                 .TextMatrix(i, 10) = Val(Format(Rst.Fields("BuyReturnAmount").Value, "##.000"))
                 .TextMatrix(i, 11) = Val(Format(Rst.Fields("SaleReturnAmount").Value, "##.000"))
                 .TextMatrix(i, 12) = Val(Format(Rst.Fields("FromStoreAmount").Value, "##.000"))
                 .TextMatrix(i, 13) = Val(Format(Rst.Fields("toStoreAmount").Value, "##.000"))
                 .TextMatrix(i, 14) = Rst.Fields("SellPrice").Value
                 .TextMatrix(i, 15) = Rst.Fields("SellPrice2").Value
                 .TextMatrix(i, 16) = Rst.Fields("SellPrice3").Value
                 .TextMatrix(i, 17) = Rst.Fields("SellPrice4").Value
                 .TextMatrix(i, 18) = Rst.Fields("SellPrice5").Value
                 .TextMatrix(i, 19) = Rst.Fields("SellPrice6").Value
                 .TextMatrix(i, 20) = Rst.Fields("BuyPrice").Value
                 .TextMatrix(i, 21) = Rst.Fields("UnitDescription").Value
                 .TextMatrix(i, 22) = IIf(Rst.Fields("MojodiControl").Value = True, -1, 0)
                 .TextMatrix(i, 23) = Rst.Fields("OrderPoint").Value
                 .TextMatrix(i, 24) = Rst.Fields("MinValue").Value
                 .TextMatrix(i, 25) = Rst.Fields("MaxValue").Value
                 
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
   '     .AutoSize 0, .Cols - 1
        .AutoSize 1, 14
        
    End With
        
End Sub

Public Sub BeforeUpdate()

End Sub

Public Sub Edit()
    If strCategory = "24" And strDelegate = "00" And (clsArya.CustomerId = 5 Or clsArya.CustomerId = 6) Then Exit Sub
    
    With vsGood
        
 '       .Editable = flexEDKbdMouse

        MyFormAddEditMode = EnumAddEditMode.EditMode
        SetFirstToolBar
    End With
End Sub

Public Sub Update()
    
    Dim i As Integer
    Dim j As Integer
    Dim LongTemp As Integer
    Dim lngSelectedSubGroup  As Long
    
    Dim Rst As New ADODB.Recordset
    
    
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
            .Row = i
            If InStr(.TextMatrix(i, 0), "*") > 0 Then 'new or edited records
            
''''                If ((Trim(.TextMatrix(i, 2)) = "" And Trim(.TextMatrix(i, 3)) = "") Or Trim(.TextMatrix(i, 5)) = "") Or .Cell(flexcpText, i, 8) = "" Or .Cell(flexcpText, i, 7) = "" Then
''''
''''                    Select Case clsStation.Language
''''
''''                        Case 0
''''
''''                            frmMsg.fwlblMsg.Caption = "‘„« „Ì »«Ì”  «ÿ·«⁄«  —« »ÿÊ— ò«„· Ê«—œ ‰„«ÌÌœ"
''''                            frmMsg.Fwbtn(0).Caption = "ﬁ»Ê·"
''''                        Case 1
''''
''''                            frmMsg.fwlblMsg.Caption = "You Have to complete the information"
''''                            frmMsg.Fwbtn(0).Caption = "Ok"
''''                            frmMsg.fwlblMsg.Alignment = vbLeftJustify
''''
''''                    End Select
''''
''''                    frmMsg.Fwbtn(0).ButtonType = flwButtonOk
''''                    frmMsg.Fwbtn(1).Visible = False
''''                    frmMsg.Show vbModal
''''
''''                    Exit Sub
''''
''''                End If
                
                If Val(.TextMatrix(i, 11)) < 0 Then     '
                        Select Case clsStation.Language
                        
                            Case 0
                            
                                frmMsg.fwlblMsg.Caption = "‘„« „Ì »«Ì”  ‰ﬁÿÂ ”›«—‘  —« Ê«—œ ‰„«ÌÌœ"
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
        Next i
        
        For j = 0 To lstGoodLevel2.ListCount - 1
            If lstGoodLevel2.Selected(j) = True Then
                lngSelectedSubGroup = j
                Exit For
            End If
        Next j

        Select Case MyFormAddEditMode
        
                
            Case EnumAddEditMode.EditMode
                
                For i = 1 To .Rows - 1
                    
                    If InStr(.TextMatrix(i, 0), "*") > 0 Then 'Edited records
                            
                        ReDim Parameter(8) As Parameter

                        Parameter(0) = GenerateInputParameter("@FirstMojodi", adDouble, 8, Val(.TextMatrix(i, 5)))
                        Parameter(1) = GenerateInputParameter("@MojodiControl", adBoolean, 1, IIf(Val(.TextMatrix(i, 22)) = -1, 1, 0))
                        Parameter(2) = GenerateInputParameter("@OrderPoint", adDouble, 8, Val(.TextMatrix(i, 23)))
                        Parameter(3) = GenerateInputParameter("@MinValue", adDouble, 8, Val(.TextMatrix(i, 24)))
                        Parameter(4) = GenerateInputParameter("@MaxValue", adDouble, 8, Val(.TextMatrix(i, 25)))
                        Parameter(5) = GenerateInputParameter("@Code", adInteger, 4, Val(Trim(.TextMatrix(i, 1))))
                        Parameter(6) = GenerateInputParameter("@IntInventoryNo", adInteger, 4, cmbInventory.ItemData(cmbInventory.ListIndex))
                        Parameter(7) = GenerateInputParameter("@Branch", adInteger, 4, cmbBranch.ItemData(cmbBranch.ListIndex))
                        Parameter(8) = GenerateInputParameter("@AccountYear", adSmallInt, 2, Val(cmbSalMali.Text))
                        
                        RunParametricStoredProcedure "Update_Good_Store", Parameter
                            
                    End If
                                        
                Next i
                
            
            End Select
            
        FillvsGood
        
    End With
    
    Set Rst = Nothing
End Sub


Public Sub Cancel()

    MyFormAddEditMode = EnumAddEditMode.ViewMode
    SetFirstToolBar
    
    FillvsGood
    
End Sub



Private Sub CheckFirstMojodi_Click()
    FillvsGood
End Sub

Private Sub CheckOrder_Click()
    FillvsGood
End Sub

Private Sub cmbBranch_Click()
    FillInventory
End Sub


Private Sub cmbInventory_Click()
    If cmbBranch.ListIndex = -1 Then Exit Sub
    FillLstGoodLevel1
    txtBarcode.SetFocus
End Sub

Private Sub cmdInventoryGood_Add_Click()
        
    Dim intSelectedLevel1 As Integer
    
    intSelectedLevel1 = -1
    
    For i = 0 To lstGoodLevel1.ListCount - 1
        If lstGoodLevel1.Selected(i) = True Then
            intSelectedLevel1 = i
            Exit For
        End If
    Next i
    If intSelectedLevel1 = -1 Then
        frmMsg.fwlblMsg.Caption = "‘„« »«Ìœ Õœ«ﬁ· Ìò ê—ÊÂ «‰ Œ«» ò‰Ìœ "
        frmMsg.fwBtn(0).ButtonType = flwButtonOk
        frmMsg.fwBtn(1).Visible = False
        frmMsg.fwBtn(0).Caption = "ﬁ»Ê·"
        frmMsg.Show vbModal
        Exit Sub
     End If
        frmMsg.fwlblMsg.Caption = "¬Ì« »—«Ì «÷«›Â ò—œ‰ ‰«„ ò·ÌÂ ò«·«Â« »Â «‰»«— «ÿ„Ì‰«‰ œ«—Ìœ"
        frmMsg.fwBtn(0).ButtonType = flwButtonOk
        frmMsg.fwBtn(1).ButtonType = flwButtonCancel
        frmMsg.fwBtn(0).Caption = "»·Ì"
        frmMsg.fwBtn(1).Caption = "ŒÌ—"
        frmMsg.Show vbModal
        If mvarMsgIdx = vbYes Then
            ReDim Parameter(3) As Parameter
    
            Parameter(0) = GenerateInputParameter("@IntInventoryNo", adInteger, 4, cmbInventory.ItemData(cmbInventory.ListIndex))
            Parameter(1) = GenerateInputParameter("@Level1", adInteger, 4, lstGoodLevel1.ItemData(intSelectedLevel1))
            Parameter(2) = GenerateInputParameter("@Branch", adInteger, 4, cmbBranch.ItemData(cmbBranch.ListIndex))
            Parameter(3) = GenerateInputParameter("@AccountYear", adSmallInt, 2, Val(cmbSalMali.Text))

            RunParametricStoredProcedure "Insert_tinventory_Good_All", Parameter
            DefaultSetting
            frmDisMsg.lblMessage = "«›“«Ì‘ ‰«„ ò·ÌÂ ò«·«Â« »Â «‰»«— «‰Ã«„ ‘œ "
            frmDisMsg.Timer1.Enabled = True
            frmDisMsg.Show vbModal
        End If

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

Private Sub cmdInventoryGood_Delete_Click()
    If cmbBranch.ListIndex = -1 Then Exit Sub
    If cmbInventory.ListIndex = -1 Then Exit Sub
    If cmbSalMali.ListIndex = -1 Then Exit Sub
    
    Dim intSelectedLevel1 As Integer
    
    frmMsg.fwlblMsg.Caption = "¬Ì« »—«Ì Õ–› ‰«„  ò«·«Â«Ì  »œÊ‰ ê—œ‘ Ê »œÊ‰ „ÊÃÊœÌ «Ê·ÌÂ «“ «‰»«— «ÿ„Ì‰«‰ œ«—Ìœ"
    frmMsg.fwBtn(0).ButtonType = flwButtonOk
    frmMsg.fwBtn(1).ButtonType = flwButtonCancel
    frmMsg.fwBtn(0).Caption = "»·Ì"
    frmMsg.fwBtn(1).Caption = "ŒÌ—"
    frmMsg.Show vbModal
    If mvarMsgIdx = vbYes Then
        ReDim Parameter(2) As Parameter

        Parameter(0) = GenerateInputParameter("@IntInventoryNo", adInteger, 4, cmbInventory.ItemData(cmbInventory.ListIndex))
        Parameter(1) = GenerateInputParameter("@Branch", adInteger, 4, cmbBranch.ItemData(cmbBranch.ListIndex))
        Parameter(2) = GenerateInputParameter("@AccountYear", adSmallInt, 2, Val(cmbSalMali.Text))
        
        RunParametricStoredProcedure "Delete_tinventory_Good_All", Parameter
        DefaultSetting
        frmDisMsg.lblMessage = "Õ–› ‰«„  ò«·«Â«Ì  »œÊ‰ ê—œ‘ Ê »œÊ‰ „ÊÃÊœÌ «Ê·ÌÂ «“ «‰»«— «‰Ã«„ ‘œ "
        frmDisMsg.Timer1.Enabled = True
        frmDisMsg.Show vbModal
    End If

End Sub

Private Sub Form_Activate()
    VarActForm = Me.Name
    
    ChangeLanguage
      
    txtDateFrom.Text = Mid(AccountYear, 3, 2) & "/01" & "/01"
    txtDateTo.Text = Mid(clsDate.shamsi(Date), 3)
    txtBarcode.Text = ""
    Frame3.BackColor = Me.BackColor
    Frame28.BackColor = Me.BackColor
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
    FillvsGood
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case Shift
          Case 0
              Select Case KeyCode
                  Case 13  ' Enter
                    SendKeys "{Left}", True
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

    If ClsFormAccess.frmMojodiControl1 = False Then
        Unload Me
        Exit Sub
    End If
        
    CenterTop Me
    VarActForm = Me.Name
    formloadFlag = False
    Me.Left = Val(GetSetting(strMainKey, Me.Name, "Left"))
    If Val(GetSetting(strMainKey, Me.Name, "Height")) > 0 Then
        Me.Height = Val(GetSetting(strMainKey, Me.Name, "Height"))
    End If
    If Val(GetSetting(strMainKey, Me.Name, "Width")) > 0 Then
        Me.Width = Val(GetSetting(strMainKey, Me.Name, "Width"))
    End If
    Me.Top = Val(GetSetting(strMainKey, Me.Name, "Top"))
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
    
''''    FillvsGood
''''
''''    MyFormAddEditMode = EnumAddEditMode.ViewMode
''''    SetFirstToolbar
    
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call PresetScreenSaver
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call PresetScreenSaver
End Sub
Private Sub lstGoodLevel1_Scroll()
 '   FillLstGoodLevel2
End Sub

Private Sub lstGoodLevel2_ItemCheck(Item As Integer)
    
    FillvsGood

End Sub

Public Sub ChangeLanguage()

Dim Obj As Object

    Select Case clsStation.Language    ' LCase(mdifrm.Toolbar1.Buttons(25).Key)
        
        Case English
            
            
            Me.Caption = "Mojodi Control"
            mdifrm.Caption = clsArya.LatinCompany
            Me.RightToLeft = False
            
            For Each Obj In Me
                On Error Resume Next
                    Obj.RightToLeft = False
                On Error GoTo 0
            Next Obj
            lblGoodLevel1.Caption = "Goods Main Groups"
            lblGoodLevel2.Caption = "Goods SubGroups"
        
        Case Farsi
            
            
            Me.Caption = ""
            mdifrm.Caption = clsArya.Company
            Me.RightToLeft = True
            
            For Each Obj In Me
                On Error Resume Next
                    Obj.RightToLeft = True
                On Error GoTo 0
            Next Obj
            
            lblGoodLevel1.Caption = " ê—ÊÂ «’·Ì ò«·«Â«-»Œ‘ Â«"
            lblGoodLevel2.Caption = "ê—ÊÂ ›—⁄Ì ò«·«Â«"
            
    End Select
    
'    lstGoodLevel1.Left = Me.Width - (lstGoodLevel1.Left + lstGoodLevel1.Width)
'    lstGoodLevel2.Left = Me.Width - (lstGoodLevel2.Left + lstGoodLevel2.Width)
    
'    lblGoodLevel1.Left = Me.Width - (lblGoodLevel1.Left + lblGoodLevel1.Width)
'    lblGoodLevel2.Left = Me.Width - (lblGoodLevel2.Left + lblGoodLevel2.Width)
        
    
    With vsGood
    
        .Cols = 27
        
        Select Case clsStation.Language
            Case Farsi
                .TextMatrix(0, 0) = "—œÌ›"
                .TextMatrix(0, 1) = "òœ"
                .TextMatrix(0, 2) = "‰«„ ò«·«"
                .TextMatrix(0, 3) = "»«—òœ"
                .TextMatrix(0, 4) = "›—Ê‘‰œÂ "
                .TextMatrix(0, 5) = "„ «Ê·ÌÂ"
                .TextMatrix(0, 6) = "Œ—Ìœ "
                .TextMatrix(0, 7) = "›—Ê‘ "
                .TextMatrix(0, 8) = "÷«Ì⁄« "
                .TextMatrix(0, 9) = "„ÊÃÊœÌ"
                .TextMatrix(0, 10) = "» «“ Œ—Ìœ"
                .TextMatrix(0, 11) = "» «“ ›—Ê‘"
                .TextMatrix(0, 12) = "ÕÊ«·Â «“ «‰»«—"
                .TextMatrix(0, 13) = "—”Ìœ »Â «‰»«—"
                .TextMatrix(0, 14) = "›Ì ›—Ê‘"
                .TextMatrix(0, 15) = "›Ì ›—Ê‘ 2"
                .TextMatrix(0, 16) = "›Ì ›—Ê‘ 3"
                .TextMatrix(0, 17) = "›Ì ›—Ê‘ 4"
                .TextMatrix(0, 18) = "›Ì ›—Ê‘ 5"
                .TextMatrix(0, 19) = "›Ì ›—Ê‘ 6"
                .TextMatrix(0, 20) = "›Ì Œ—Ìœ"
                .TextMatrix(0, 21) = "Ê«Õœ "
                .TextMatrix(0, 22) = " ﬂ‰ —· "
                .TextMatrix(0, 23) = "‰ﬁÿÂ ”›«—‘ "
                .TextMatrix(0, 24) = "Õœ«ﬁ·"
                .TextMatrix(0, 25) = "Õœ«ﬂÀ—"
                .TextMatrix(0, 26) = "    "
            
            Case English
                .TextMatrix(0, 0) = "Row"
                .TextMatrix(0, 1) = "Code"
                .TextMatrix(0, 2) = "Name"
                .TextMatrix(0, 3) = "Barcode"
                .TextMatrix(0, 4) = "Seller"
                .TextMatrix(0, 5) = "FirstStock"
                .TextMatrix(0, 6) = "Purchase"
                .TextMatrix(0, 7) = "Sale"
                .TextMatrix(0, 8) = "Losses"
                .TextMatrix(0, 9) = "Stock"
                .TextMatrix(0, 10) = "PurchaseReturn"
                .TextMatrix(0, 11) = "SaleReturn"
                .TextMatrix(0, 12) = "FRomStore"
                .TextMatrix(0, 13) = "toStore"
                .TextMatrix(0, 14) = "Fee 1"
                .TextMatrix(0, 15) = "Fee 2"
                .TextMatrix(0, 16) = "Fee 3"
                .TextMatrix(0, 17) = "Fee 4"
                .TextMatrix(0, 18) = "Fee 5"
                .TextMatrix(0, 19) = "Fee 6"
                .TextMatrix(0, 20) = "Buyprice"
                .TextMatrix(0, 21) = " Unit"
                .TextMatrix(0, 22) = "Control"
                .TextMatrix(0, 23) = "Ordre"
                .TextMatrix(0, 24) = "Minimum"
                .TextMatrix(0, 25) = "Maximum"
                .TextMatrix(0, 26) = "   "
       End Select
        If clsArya.MultiPrice = False Then
           .ColHidden(15) = True
           .ColHidden(16) = True
           .ColHidden(17) = True
           .ColHidden(18) = True
           .ColHidden(19) = True
        Else
            If clsStation.MaxPrices = 5 Then
                .ColHidden(19) = True
            ElseIf clsStation.MaxPrices = 4 Then
                .ColHidden(18) = True
                .ColHidden(19) = True
            ElseIf clsStation.MaxPrices = 3 Then
                .ColHidden(17) = True
                .ColHidden(18) = True
                .ColHidden(19) = True
            ElseIf clsStation.MaxPrices = 2 Then
                .ColHidden(16) = True
                .ColHidden(17) = True
                .ColHidden(18) = True
                .ColHidden(19) = True
            ElseIf clsStation.MaxPrices = 1 Then
                .ColHidden(15) = True
                .ColHidden(16) = True
                .ColHidden(17) = True
                .ColHidden(18) = True
                .ColHidden(19) = True
            End If
        End If
       .ColDataType(0) = flexDTDouble
       .ColDataType(10) = flexDTString
       .ColDataType(22) = flexDTBoolean
    '   .ColDataType(9) = flexDTDecimal
       ' .ColSort(9) = flexSortNumericAscending + flexSortNumericDescending
        .ColSort(9) = flexSortCustom
        .ColAlignment(-1) = flexAlignCenterCenter
      '  .ColAlignment(25) = flexAlignRightCenter
        .FocusRect = flexFocusHeavy
       ' .ColHidden(1) = True
        .AutoSizeMode = flexAutoSizeColWidth
        .AutoSize 0, .Cols - 2
       ' .AutoSize 2, 14
        .AutoSearch = flexSearchFromCursor
    End With
    
    FillBranch
    FillInventory
    FillSalMali
    DefaultSetting
            
    SetFirstToolBar

End Sub
Private Sub FillBranch()
    Dim rctmp As New ADODB.Recordset
    cmbBranch.Clear
    Set rctmp = RunStoredProcedure2RecordSet("Get_All_Branches")
    Do While rctmp.EOF = False
        cmbBranch.AddItem rctmp!nvcBranchName
        cmbBranch.ItemData(cmbBranch.NewIndex) = rctmp!Branch
        rctmp.MoveNext
    Loop
    rctmp.Close
    'If cmbBranch.ListCount > 0 Then cmbBranch.ListIndex = 0
    

End Sub
Private Sub FillInventory()
    cmbInventory.Clear
    If cmbBranch.ListIndex = -1 Then Exit Sub
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
  '  cmbInventory.ListIndex = 0

End Sub

Private Sub ResizeKit1_ExitResize(ByVal XScale As Double, ByVal YScale As Double)

    If formloadFlag = True Then
        SaveSetting strMainKey, Me.Name, "Height", Me.Height
        SaveSetting strMainKey, Me.Name, "Width", Me.Width
    End If

End Sub

Private Sub StoreDataUpdate_Click()
    If cmbInventory.ListIndex = -1 Then Exit Sub

    If Trim(txtDateFrom.ClipText) = "" Or Trim(txtDateTo.ClipText) = "" Then
        frmDisMsg.lblMessage = "  «—ÌŒ „⁄ »— Ê«—œ ﬂ‰Ìœ "
        frmDisMsg.Timer1.Enabled = True
        frmDisMsg.Show vbModal
        Exit Sub
    End If
       ' StoreDataUpdate.Enabled = False
        FWProgressBar1.Value = 0
        ReDim Parameter(11) As Parameter
    
        Parameter(0) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
        Parameter(1) = GenerateInputParameter("@SystemDate", adVarWChar, 50, Mid(clsDate.shamsi(Date), 3, 8))
        Parameter(2) = GenerateInputParameter("@SystemDay", adVarWChar, 50, clsDate.Find_DayOfWeek(Weekday(Date, vbSaturday)))
        Parameter(3) = GenerateInputParameter("@SystemTime", adVarWChar, 50, Mid(Str(Time), 1, 5))
        Parameter(4) = GenerateInputParameter("@DateBefore", adVarWChar, 50, CStr(IIf(Trim(txtDateFrom.ClipText) = "", "", Trim(txtDateFrom.Text))))
        Parameter(5) = GenerateInputParameter("@DateAfter", adVarWChar, 50, CStr(IIf(Trim(txtDateTo.ClipText) = "", "", Trim(txtDateTo.Text))))
        Parameter(6) = GenerateInputParameter("@Type", adInteger, 4, EnumGoodType.forBuySale)
        Parameter(7) = GenerateInputParameter("@InVentoryNo1", adInteger, 4, cmbInventory.ItemData(cmbInventory.ListIndex))
        Parameter(8) = GenerateInputParameter("@InVentoryNo2", adInteger, 4, cmbInventory.ItemData(cmbInventory.ListIndex))
        Parameter(9) = GenerateInputParameter("@Branch", adInteger, 4, cmbBranch.ItemData(cmbBranch.ListIndex))
        Parameter(10) = GenerateInputParameter("@UsePercentFlag", adInteger, 4, 0)
        Parameter(11) = GenerateInputParameter("@AccountYear", adSmallInt, 2, Val(cmbSalMali.Text))
        
        RunParametricStoredProcedure "Update_tblTotal_tInventory_tGood_For_Mojodi", Parameter
        FWProgressBar1.Value = 100

'        Set Rst = RunParametricStoredProcedure2Rec("GetInventoryAtomicReport_Mojodi", Parameter)
    
        DefaultSetting
        FWProgressBar1.Value = 0
        StoreDataUpdate.Enabled = True
        frmDisMsg.lblMessage = " »Â —Ê“ —”«‰Ì «‰Ã«„ ‘œ "
        frmDisMsg.Timer1.Enabled = True
        frmDisMsg.Show vbModal

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
    i = vsGood.FindRow(Trim(txtBarcode.Text), 1, 3, True, True)
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

Private Sub vsGood_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    With vsGood
        If (.TextMatrix(Row, 0) = "*" Or MyFormAddEditMode = EnumAddEditMode.EditMode) And Col > 1 And tmpTextMatrix <> .TextMatrix(Row, Col) Then
        
            If MyFormAddEditMode = EnumAddEditMode.EditMode And InStr(.TextMatrix(Row, 0), "*") = 0 Then
                .TextMatrix(Row, 0) = Trim(.TextMatrix(Row, 0)) & "*"
            End If
            
        Else

        End If
      '  .AutoSizeMode = flexAutoSizeColWidth
      '  .AutoSize Col, Col
        

    End With


End Sub

Private Sub vsGood_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    tmpTextMatrix = vsGood.TextMatrix(Row, Col)
End Sub

Private Sub vsGood_BeforeSort(ByVal Col As Long, Order As Integer)
If Col = 5 Or Col = 9 Then
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
If Col = 5 Or Col = 9 Then
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
 If strCategory = "24" And strDelegate = "00" And (clsArya.CustomerId = 5 Or clsArya.CustomerId = 6) Then Exit Sub
    With vsGood
        If (.TextMatrix(.Row, 0) = "*" Or MyFormAddEditMode = EnumAddEditMode.EditMode) And (.Col = 5 Or .Col > 21) Then
            If .Col = 5 Or .Col > 21 Then
               .Select .Row, .Col
               .EditCell
            End If
        End If
    
    End With

End Sub

Private Sub vsGood_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 37 Or KeyCode = 38 Or KeyCode = 39 Or KeyCode = 40 Then Exit Sub
    If strCategory = "24" And strDelegate = "00" And (clsArya.CustomerId = 5 Or clsArya.CustomerId = 6) Then Exit Sub
    With vsGood
        If (.TextMatrix(.Row, 0) = "*" Or MyFormAddEditMode = EnumAddEditMode.EditMode) And (.Col = 5 Or .Col > 21) Then
            If .Col = 5 Or .Col > 21 Then
               .Select .Row, .Col
               .EditCell
            End If
        End If
    End With
    
End Sub


Private Sub vsGood_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    With vsGood
        If KeyAscii = 39 Then KeyAscii = 0
        
        If (Col <> 5 And Col < 22) Or (IsNumeric(Chr(KeyAscii)) = False And KeyAscii = 8) Then
            
            KeyAscii = 0
            
        ElseIf IsNumeric(Chr(KeyAscii)) = False And KeyAscii <> 46 Then
            
            KeyAscii = 0
            
        ElseIf (Col <> 5 And Col < 22) Or KeyAscii = 8 Then
            
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
    On Error GoTo Err_Handler
    
    If cmbInventory.ListIndex = -1 Then
        ShowMessage "«‰»«— «‰ Œ«» ‰‘œÂ «” ", True, False, " «ÌÌœ", ""
        Exit Sub
    ElseIf cmbBranch.ListIndex = -1 Then
        ShowMessage "‘⁄»Â «‰ Œ«» ‰‘œÂ «” ", True, False, " «ÌÌœ", ""
        Exit Sub
    End If
    
    frmInput.fwlblInput.Caption = "‰Ê⁄ ê“«—‘ "
    frmInput.OptionLevel(0).Caption = "ê“«—‘ „ÊÃÊœÌ"
    frmInput.OptionLevel(1).Caption = "ê“«—‘ ”›«—‘ ﬂ«·«"
    frmInput.OptionLevel(0).Value = True
    frmInput.btnCancel.Visible = True
    frmInput.Picture1.Visible = True
    frmInput.txtInput.Visible = False
                    
    frmInput.Show vbModal
    If mvarInput = "" Then
        Exit Sub
    End If
    If lstGoodLevel1.ListCount < 1 Then Exit Sub
    If lstGoodLevel2.ListCount < 1 Then Exit Sub
    If cmbSalMali.ListIndex = -1 Then Exit Sub
  Dim i As Long
    Dim j As Long
    Dim intSelectedLevel1 As Integer
    Dim intSelectedLevel2 As Integer
    Dim strSelectedLevels As String
    Dim Rst As New ADODB.Recordset
    'Dim Rst2 As New ADODB.Recordset
    
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
    
    Dim level1 As Integer
    If intSelectedLevel1 <> -1 And intSelectedLevel2 <> -1 Then
        level1 = lstGoodLevel1.ItemData(intSelectedLevel1)
        strSelectedLevels = Right(strSelectedLevels, Len(strSelectedLevels) - 1)
    ElseIf intSelectedLevel1 <> -1 And intSelectedLevel2 = -1 Then  'Or intSelectedLevel2 = -1
       level1 = lstGoodLevel1.ItemData(intSelectedLevel1)
       strSelectedLevels = ""
    Else
        strSelectedLevels = ""
        level1 = -1
    End If
    
    With vsGood
        RunNonParametricStoredProcedure "Delete_tblPrint_Order"

        If mvarInput = "0" Then
            ReDim Parameter(8) As Parameter
            Parameter(0) = GenerateInputParameter("@SystemDate", adVarWChar, 50, Mid(clsDate.shamsi(Date), 3, 8))
            Parameter(1) = GenerateInputParameter("@SystemDay", adVarWChar, 50, clsDate.Find_DayOfWeek(Weekday(Date, vbSaturday)))
            Parameter(2) = GenerateInputParameter("@SystemTime", adVarWChar, 50, Mid(Str(Time), 1, 5))
            Parameter(3) = GenerateInputParameter("@Level1", adInteger, 4, level1)
            Parameter(4) = GenerateInputParameter("@strSelectedLevels", adVarWChar, 4000, strSelectedLevels)
            Parameter(5) = GenerateInputParameter("@GoodType", adInteger, 4, EnumGoodType.forBuySale)
            Parameter(6) = GenerateInputParameter("@InventoryNo", adInteger, 4, cmbInventory.ItemData(cmbInventory.ListIndex))
            Parameter(7) = GenerateInputParameter("@Branch", adInteger, 4, cmbBranch.ItemData(cmbBranch.ListIndex))
            Parameter(8) = GenerateInputParameter("@AccountYear", adSmallInt, 2, Val(cmbSalMali.Text))
            
            CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\RepMojodiControl.rpt"
            CrystalReport1.ReportTitle = "  ê“«—‘ „ÊÃÊœÌ -" & cmbInventory.Text
            
        ElseIf mvarInput = "1" Then
            ReDim Parameter(12) As Parameter
            For i = 1 To .Rows - 1
                    Parameter(0) = GenerateInputParameter("@Row", adInteger, 4, i)
                    Parameter(1) = GenerateInputParameter("@GoodName", adVarChar, 50, .TextMatrix(i, 2))
                    Parameter(2) = GenerateInputParameter("@UnitName", adVarChar, 50, .TextMatrix(i, 21))
                    Parameter(3) = GenerateInputParameter("@Mojodi", adDouble, 8, .TextMatrix(i, 9))
                    Parameter(4) = GenerateInputParameter("@OrderPoint", adDouble, 8, Val(.TextMatrix(i, 23)))
                    Parameter(5) = GenerateInputParameter("@Minimum", adDouble, 8, Val(.TextMatrix(i, 24)))
                    Parameter(6) = GenerateInputParameter("@Maximum", adDouble, 8, Val(.TextMatrix(i, 25)))
                    Parameter(7) = GenerateInputParameter("@BuyPrice", adInteger, 4, Val(.TextMatrix(i, 20)))
                    Parameter(8) = GenerateInputParameter("@Sellprice", adInteger, 4, Val(.TextMatrix(i, 14)))
                    Parameter(9) = GenerateInputParameter("@Sellprice2", adInteger, 4, Val(.TextMatrix(i, 15)))
                    Parameter(10) = GenerateInputParameter("@Sellprice3", adInteger, 4, Val(.TextMatrix(i, 16)))
                    Parameter(11) = GenerateInputParameter("@Barcode", adVarChar, 20, .TextMatrix(i, 3))
                    Parameter(12) = GenerateInputParameter("@FirstMojodi", adDouble, 8, .TextMatrix(i, 5))
                    
                    RunParametricStoredProcedure "Insert_tblPrint_Order", Parameter
            Next i
            
            ReDim Parameter(2) As Parameter
            Parameter(0) = GenerateInputParameter("@SystemDate", adVarWChar, 50, Mid(clsDate.shamsi(Date), 3, 8))
            Parameter(1) = GenerateInputParameter("@SystemDay", adVarWChar, 50, clsDate.Find_DayOfWeek(Weekday(Date, vbSaturday)))
            Parameter(2) = GenerateInputParameter("@SystemTime", adVarWChar, 50, Mid(Str(Time), 1, 5))
            CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\RepOrder_A4.rpt"
            CrystalReport1.ReportTitle = "”›«—‘ »—«Ì ﬂ«·«Â«Ì Œ—Ìœ‰Ì Ê ›—ÊŒ ‰Ì"
        Else
            Exit Sub
        End If

      '  CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\RepOrder.rpt"
'        CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\RepOrder_A4.rpt"'
'        CrystalReport1.ReportTitle = "”›«—‘ »—«Ì ﬂ«·«Â«Ì Œ—Ìœ‰Ì Ê ›—ÊŒ ‰Ì"
        
        Dim intIndex As Integer
        For intIndex = 0 To 100
            CrystalReport1.ParameterFields(intIndex) = ""
        Next intIndex
        
        For intIndex = 0 To UBound(Parameter) - LBound(Parameter)
            CrystalReport1.ParameterFields(intIndex) = CStr(Parameter(intIndex).Name) & ";" & CStr(Parameter(intIndex).Value) & ";" & "True"
        Next intIndex
      
        CrystalReport1.Destination = crptToWindow 'crptToPrinter '
        CrystalReport1.RetrieveDataFiles
        CrystalReport1.Connect = CrystallConnection
        CrystalReport1.Action = 1
        
        If Screen.Width > 12000 Then
            CrystalReport1.PageZoom (100)
        Else
            CrystalReport1.PageZoom (75)
        End If
    End With

Exit Sub
Err_Handler:
    LogSaveNew "frmMojodiControl => ", err.Description, err.Number, err.Source, "Printing"
    ShowErrorMessage
End Sub

