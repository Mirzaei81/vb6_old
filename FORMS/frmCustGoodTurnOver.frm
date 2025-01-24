VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{7AEDC602-D94C-11D1-BB7A-00E0290EA3C9}#1.0#0"; "ResizeKit.ocx"
Object = "{75D4F767-8785-11D3-93AD-0000832EF44D}#2.14#0"; "FAST2001.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmCustGoodTurnOver 
   ClientHeight    =   9450
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10095
   Icon            =   "frmCustGoodTurnOver.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9450
   ScaleWidth      =   10095
   Begin VB.Frame Frame4 
      Caption         =   "„‘ —ﬂ"
      BeginProperty Font 
         Name            =   "B Homa"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   6000
      RightToLeft     =   -1  'True
      TabIndex        =   16
      Top             =   600
      Width           =   3975
      Begin VB.TextBox Text1 
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
         Height          =   495
         Left            =   120
         MultiLine       =   -1  'True
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   1080
         Width           =   2655
      End
      Begin FLWCtrls.FWCoolButton fwBtnCustFind 
         Height          =   570
         Left            =   120
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   360
         Width           =   3570
         _ExtentX        =   6297
         _ExtentY        =   1005
         BackColor       =   16777088
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Nazanin"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Picture         =   "frmCustGoodTurnOver.frx":A4C2
         PictureAlign    =   4
         Caption         =   "„‘ —Ì"
         MaskColor       =   -2147483633
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         Caption         =   "«‘ —«ò "
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
         Left            =   2880
         RightToLeft     =   -1  'True
         TabIndex        =   19
         Top             =   1080
         Width           =   855
      End
   End
   Begin RESIZEKITLibCtl.ResizeKit ResizeKit1 
      Height          =   480
      Left            =   240
      OleObjectBlob   =   "frmCustGoodTurnOver.frx":A7DC
      TabIndex        =   15
      Top             =   0
      Width           =   480
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
      ForeColor       =   &H8000000D&
      Height          =   960
      Left            =   3120
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   840
      Width           =   2775
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
         Style           =   2  'Dropdown List
         TabIndex        =   11
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
      ForeColor       =   &H8000000D&
      Height          =   960
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   840
      Width           =   3015
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
         Left            =   120
         RightToLeft     =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   360
         Width           =   2475
      End
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   9600
      Top             =   120
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
      Caption         =   "ò«·«          "
      BeginProperty Font 
         Name            =   "B Homa"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   6000
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   2280
      Width           =   3975
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
         Left            =   120
         MultiLine       =   -1  'True
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Top             =   1080
         Width           =   2745
      End
      Begin FLWCtrls.FWCoolButton fwBtnGoodFind 
         Height          =   570
         Left            =   120
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   360
         Width           =   3570
         _ExtentX        =   6297
         _ExtentY        =   1005
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
         Picture         =   "frmCustGoodTurnOver.frx":A862
         PictureAlign    =   4
         Caption         =   "ò«·«"
         MaskColor       =   -2147483633
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
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
         Height          =   435
         Left            =   2640
         RightToLeft     =   -1  'True
         TabIndex        =   21
         Top             =   1080
         Width           =   825
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2055
      Left            =   120
      TabIndex        =   1
      Top             =   1920
      Width           =   5775
      Begin VB.CommandButton StoreDataUpdate 
         Caption         =   "„Õ«”»Â ê—œ‘ ﬂ«·« "
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   240
         TabIndex        =   22
         Top             =   840
         Width           =   1935
      End
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
         Left            =   120
         RightToLeft     =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   240
         Width           =   1455
      End
      Begin MSMask.MaskEdBox txtDateTo 
         Height          =   465
         Left            =   2640
         TabIndex        =   2
         Top             =   840
         Width           =   2115
         _ExtentX        =   3731
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
         Left            =   2640
         TabIndex        =   3
         Top             =   240
         Width           =   2115
         _ExtentX        =   3731
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
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "”«· „«·Ì"
         BeginProperty Font 
            Name            =   "B Traffic"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1800
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   360
         Width           =   735
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
         Left            =   4440
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   840
         Width           =   945
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
         Left            =   4560
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   240
         Width           =   825
      End
   End
   Begin VSFlex7LCtl.VSFlexGrid vsGood 
      Height          =   5265
      Left            =   120
      TabIndex        =   0
      Top             =   4080
      Width           =   9900
      _cx             =   17462
      _cy             =   9287
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
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   11
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmCustGoodTurnOver.frx":AB7C
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
      Height          =   495
      Left            =   8400
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
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ê—œ‘ ﬂ«·«Ì „‘ —Ì«‰"
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
      Left            =   2520
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   0
      Width           =   5175
   End
End
Attribute VB_Name = "frmCustGoodTurnOver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private MyFormAddEditMode As EnumAddEditMode
Dim tmpTextMatrix As String
Dim Parameter() As Parameter
Dim clsDate As New clsDate

Public Sub ExitForm()

    Unload Me
        
End Sub

Public Sub SetFirstToolBar()

    Dim i As Integer
    
    AllButton vbOff, True
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
        
            mdifrm.Toolbar1.Buttons(8).Enabled = False 'enter key
            mdifrm.Toolbar1.Buttons(9).Enabled = False 'cancel key
 '           vsGood.Editable = flexEDKbdMouse
            
        Case EnumAddEditMode.EditMode
                    
            mdifrm.Toolbar1.Buttons(8).Enabled = False 'enter key
            mdifrm.Toolbar1.Buttons(9).Enabled = False 'cancel key
'            vsGood.Editable = flexEDKbdMouse
    End Select
    
    HeaderLabel Val(MyFormAddEditMode), fwlblMode
    
End Sub

Public Sub DefaultSetting()
    vsGood.Rows = 1
End Sub

Public Sub FillvsGood() 'it fills the grid using vw_Good
    vsGood.Rows = 1
    'If fwBtnGoodFind.Tag = "" Then Exit Sub
    If Val(fwBtnCustFind.Tag) = 0 Then Exit Sub
    MyFormAddEditMode = EnumAddEditMode.ViewMode 'VIEW Mode
    SetFirstToolBar
    Dim InventoryNo As Integer
    If cmbInventory.ListIndex = -1 Then
        InventoryNo = 0
    Else
        InventoryNo = cmbInventory.ItemData(cmbInventory.ListIndex)
    End If
    ReDim Parameter(10) As Parameter

    Parameter(0) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
    Parameter(1) = GenerateInputParameter("@SystemDate", adVarWChar, 50, Mid(clsDate.shamsi(Date), 3, 8))
    Parameter(2) = GenerateInputParameter("@SystemDay", adVarWChar, 50, clsDate.Find_DayOfWeek(Weekday(Date, vbSaturday)))
    Parameter(3) = GenerateInputParameter("@SystemTime", adVarWChar, 50, Mid(Str(time), 1, 5))
    Parameter(4) = GenerateInputParameter("@DateBefore", adVarWChar, 50, CStr(IIf(Trim(txtDateFrom.ClipText) = "", "", Trim(txtDateFrom.Text))))
    Parameter(5) = GenerateInputParameter("@DateAfter", adVarWChar, 50, CStr(IIf(Trim(txtDateTo.ClipText) = "", "", Trim(txtDateTo.Text))))
    Parameter(6) = GenerateInputParameter("@GoodCode", adInteger, 4, Val(fwBtnGoodFind.Tag))
    Parameter(7) = GenerateInputParameter("@InVentoryNo", adInteger, 4, InventoryNo)
    Parameter(8) = GenerateInputParameter("@Branch", adInteger, 4, cmbBranch.ItemData(cmbBranch.ListIndex))
    Parameter(9) = GenerateInputParameter("@AccountYear", adSmallInt, 2, Val(cmbSalMali.Text))
    Parameter(10) = GenerateInputParameter("@Customer", adInteger, 4, fwBtnCustFind.Tag)
    
    Set Rst = RunParametricStoredProcedure2Rec("Get_CustomerGood", Parameter)
    
    frmDisMsg.lblMessage = "„Õ«”»Â «‰Ã«„ ‘œ "
    frmDisMsg.Timer1.Enabled = True
    frmDisMsg.Show vbModal
    Dim Mojodi As Double
    If Not (Rst.EOF = True And Rst.BOF = True) Then
                 With vsGood
           i = 1
            While Rst.EOF = False
                .Rows = .Rows + 1
                .TextMatrix(i, 0) = i
                .TextMatrix(i, 1) = Rst.Fields("No").Value
                .TextMatrix(i, 2) = Rst.Fields("Date").Value
                .TextMatrix(i, 3) = Left(Rst.Fields("InventoryName").Value, 25)
                .TextMatrix(i, 4) = Rst.Fields("GoodName").Value
                .TextMatrix(i, 5) = Rst.Fields("Amount").Value
                .TextMatrix(i, 6) = Rst.Fields("FeeUnit").Value
                .TextMatrix(i, 7) = Rst.Fields("Amount").Value * Rst.Fields("FeeUnit").Value
                 Rst.MoveNext
                 i = i + 1
            Wend
            Set Rst = Nothing
            
                
''''            .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter
''''            .Cell(flexcpAlignment, 1, 2, .Rows - 1, 2) = flexAlignRightCenter
''''            ''.AutoSizeMode = flexAutoSizeColWidth
''''            .AutoSize 0, .Cols - 1
            
        End With
    
    End If
End Sub


Public Sub Cancel()

    MyFormAddEditMode = EnumAddEditMode.ViewMode
    SetFirstToolBar
    fwBtnGoodFind.Tag = ""
    fwBtnGoodFind.Caption = ""
    vsGood.Rows = 1
    
End Sub

Private Sub cmbBranch_Click()
    FillInventory
End Sub

Private Sub cmbInventory_Click()
    If cmbBranch.ListIndex = -1 Then Exit Sub
    fwBtnGoodFind.Tag = ""
    fwBtnGoodFind.Caption = ""
End Sub

Private Sub cmbSalMali_Change()
    If cmbSalMali.Text <> "" Then
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
    End If
End Sub

Private Sub cmbSalMali_Click()
    cmbSalMali_Change
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
    rs.Close
End Sub

Private Sub Form_Activate()
    VarActForm = Me.Name
    SetFirstToolBar
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

    If ClsFormAccess.frmCustGoodTurnOver = False Then
        Unload Me
        Exit Sub
    End If
    CenterTop Me
    VarActForm = Me.Name
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

    txtDateFrom.Text = Mid(AccountYear, 3, 2) & "/01" & "/01"
    txtDateTo.Text = Mid(clsDate.shamsi(Date), 3)
    
    Frame1.BackColor = Me.BackColor
    FillBranch
    FillInventory
    FillSalMali
    
    ChangeLanguage
    DefaultSetting
    
End Sub


Private Sub Form_Unload(Cancel As Integer)
     Unload frmFindCust
    VarActForm = ""

    Dim i As Integer
    
    AllButton vbOff, True
    Unload frmFindGoods

    SaveSetting strMainKey, Me.Name, "Left", Me.Left
    SaveSetting strMainKey, Me.Name, "Top", Me.Top


    If PosConnection.State = adStateOpen Then PosConnection.Close: Set PosConnection = Nothing
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call PresetScreenSaver
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call PresetScreenSaver
End Sub

Public Sub ChangeLanguage()

'Dim obj As Object
'
'    Select Case clsStation.Language    ' LCase(mdifrm.Toolbar1.Buttons(25).Key)
'
'        Case English
'
'
'            mdifrm.Caption = clsArya.LatinCompany
'            Me.RightToLeft = False
'
'            For Each obj In Me
'                On Error Resume Next
'                    obj.RightToLeft = False
'                On Error GoTo 0
'            Next obj
'
'        Case Farsi
'
'
'            mdifrm.Caption = clsArya.Company
'            Me.RightToLeft = True
'
'            For Each obj In Me
'                On Error Resume Next
'                    obj.RightToLeft = True
'                On Error GoTo 0
'            Next obj
'
'
'    End Select
    
       
    
    With vsGood
    
        .Cols = 8
        
        Select Case clsStation.Language
            Case Farsi
                .TextMatrix(0, 0) = "—œÌ›"
                .TextMatrix(0, 1) = "”‰œ"
                .TextMatrix(0, 2) = " «—ÌŒ"
                .TextMatrix(0, 3) = " «‰»«— "
                .TextMatrix(0, 4) = "‰«„ ò«·«"
                .TextMatrix(0, 5) = " ⁄œ«œ"
                .TextMatrix(0, 6) = " ›Ì ›—Ê‘ "
                .TextMatrix(i, 7) = "Ã„⁄ „»·€"
                
            Case English
                .TextMatrix(0, 0) = "Row"
                .TextMatrix(0, 1) = "SanadNo"
                .TextMatrix(0, 2) = "Date"
                .TextMatrix(0, 3) = "Store"
                .TextMatrix(0, 4) = "GoodName"
                .TextMatrix(0, 5) = "Amount"
                .TextMatrix(0, 6) = " SellPrice "
                .TextMatrix(i, 7) = ""
                
       End Select
   '     .ColSort(5) = flexSortNumericAscending + flexSortNumericDescending
        .FixedAlignment(-1) = flexAlignCenterCenter
        .ColAlignment(-1) = flexAlignCenterCenter
        .ColAlignment(4) = flexAlignRightCenter
        .FocusRect = flexFocusHeavy
'        .AutoSizeMode = flexAutoSizeColWidth
'        .AutoSize 0, .Cols - 1
'        .AutoSearch = flexSearchFromCursor
    End With
    
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
Private Sub FillInventory()
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
    If cmbInventory.ListCount > 0 Then cmbInventory.ListIndex = 0

End Sub



Private Sub fwBtnCustFind_Click()
Me.FindCust
End Sub
Private Sub fwBtnGoodFind_Click()
    frmFindGoods.Show vbModal
    fwBtnGoodFind.Caption = mvarName
    fwBtnGoodFind.Tag = mvarcode
'    If mvarcode = 0 Then
'       ' StoreDataUpdate.Enabled = False
'    Else
'        StoreDataUpdate.Enabled = True
'    End If
    txtBarcode.Text = mvarBarcodeName
End Sub


Private Sub ResizeKit1_ExitResize(ByVal XScale As Double, ByVal YScale As Double)

    If formloadFlag = True Then
        SaveSetting strMainKey, Me.Name, "Height", Me.Height
        SaveSetting strMainKey, Me.Name, "Width", Me.Width
    End If

End Sub

Private Sub StoreDataUpdate_Click()
    
    If Trim(txtDateFrom.ClipText) = "" Or Trim(txtDateTo.ClipText) = "" Then
        frmDisMsg.lblMessage = "  «—ÌŒ „⁄ »— Ê«—œ ﬂ‰Ìœ "
        frmDisMsg.Timer1.Enabled = True
        frmDisMsg.Show vbModal
        Exit Sub
    End If
    FillvsGood
End Sub


Public Sub Printing()
   ReDim Parameter(10) As Parameter

    Parameter(0) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
    Parameter(1) = GenerateInputParameter("@SystemDate", adVarWChar, 50, Mid(clsDate.shamsi(Date), 3, 8))
    Parameter(2) = GenerateInputParameter("@SystemDay", adVarWChar, 50, clsDate.Find_DayOfWeek(Weekday(Date, vbSaturday)))
    Parameter(3) = GenerateInputParameter("@SystemTime", adVarWChar, 50, Mid(Str(time), 1, 5))
    Parameter(4) = GenerateInputParameter("@DateBefore", adVarWChar, 50, CStr(IIf(Trim(txtDateFrom.ClipText) = "", "", Trim(txtDateFrom.Text))))
    Parameter(5) = GenerateInputParameter("@DateAfter", adVarWChar, 50, CStr(IIf(Trim(txtDateTo.ClipText) = "", "", Trim(txtDateTo.Text))))
    Parameter(6) = GenerateInputParameter("@GoodCode", adInteger, 4, Val(fwBtnGoodFind.Tag))
    Parameter(7) = GenerateInputParameter("@InVentoryNo", adInteger, 4, InventoryNo)
    Parameter(8) = GenerateInputParameter("@Branch", adInteger, 4, cmbBranch.ItemData(cmbBranch.ListIndex))
    Parameter(9) = GenerateInputParameter("@AccountYear", adSmallInt, 2, Val(cmbSalMali.Text))
    Parameter(10) = GenerateInputParameter("@Customer", adInteger, 4, fwBtnCustFind.Tag)
    
  '  CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\RepOrder.rpt"
    CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\RepCustOmerGoodTurnOver_A4.rpt"
    
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
    CrystalReport1.ReportTitle = " ê—œ‘ ò«·« „‘ —Ì«‰ "
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

Private Sub Text1_Change()
  If Val(Text1.Text) = -1 Then vsGood.Rows = 1: Exit Sub
  ReDim Parameter(1) As Parameter
    Parameter(0) = GenerateInputParameter("@Membershipid", adBigInt, 8, Val(Text1.Text))
    Parameter(1) = GenerateInputParameter("@Branch", adInteger, 4, CurrentBranch)
    Set rctmp = RunParametricStoredProcedure2Rec("Get_vw_Customers_ByMembership", Parameter)
    
    If rctmp.EOF = False And rctmp.BOF = False Then
        If fwBtnCustFind.Tag <> rctmp!Code Then
            fwBtnCustFind.Tag = rctmp!Code
             fwBtnCustFind.Caption = ""
             UpdatelblCustomer
        End If
    Else
        fwBtnCustFind.Tag = 0
        fwBtnCustFind.Caption = "„‘ —Ì"
        UpdatelblCustomer
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
                    CheckBarcode
            End Select
    End Select

End Sub

Private Sub CheckBarcode()
    ReDim Parameter(3) As Parameter
    Parameter(0) = GenerateInputParameter("@Barcode", adVarWChar, 50, txtBarcode.Text)
    Parameter(1) = GenerateInputParameter("@StationId", adInteger, 4, clsArya.StationNo)
    Parameter(2) = GenerateInputParameter("@Flag", adBoolean, 1, 0)
    Parameter(3) = GenerateInputParameter("@AccountYear", adSmallInt, 2, Val(cmbSalMali.Text))
    Set rctmp = RunParametricStoredProcedure2Rec("Get_Good_Barcode", Parameter)
    
    If Not (rctmp.BOF Or rctmp.EOF) Then
        fwBtnGoodFind.Caption = rctmp.Fields("Name")
        fwBtnGoodFind.Tag = rctmp.Fields("Code")
    Else
        fwBtnGoodFind.Tag = 0
        frmDisMsg.lblMessage.Caption = " . «Ì‰ »«—ﬂœ œ— ”Ì” „  ⁄—Ì› ‰‘œÂ «”  "
        frmDisMsg.Timer1.Interval = 1000
        frmDisMsg.Timer1.Enabled = True
        frmDisMsg.Show vbModal
    End If
    If fwBtnGoodFind.Tag = 0 Then
       ' StoreDataUpdate.Enabled = False
    Else
        StoreDataUpdate.Enabled = True
        StoreDataUpdate.SetFocus
    End If

End Sub

Private Sub vsGood_Click()
With vsGood
    If .Col = 1 Then
        .Sort = flexSortNumericAscending
        .ColSort(1) = flexSortGenericAscending + flexSortGenericDescending
    End If
End With
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
      Else
                    
        frmDisMsg.lblMessage.Caption = " ‘„« »Â «Ì‰ «„ﬂ«‰ œ” —”Ì ‰œ«—Ìœ "
        frmDisMsg.Timer1.Interval = 1000
        frmDisMsg.Timer1.Enabled = True
        frmDisMsg.Show vbModal
       
     End If
   
End Sub
Private Sub UpdatelblCustomer()

    If fwBtnCustFind.Tag <> "" Then
        Dim Rst As New ADODB.Recordset
        Dim mvarMemberShipId, mvarTel, mvarAddress, mvarDescription As String
        
        
        ReDim Parameter(1) As Parameter
        Parameter(0) = GenerateInputParameter("@Code", adInteger, 4, Val(fwBtnCustFind.Tag))
        Parameter(1) = GenerateInputParameter("@Branch", adInteger, 4, CurrentBranch)
        Set Rst = RunParametricStoredProcedure2Rec("Get_vw_Customers", Parameter)
        
        If Rst.EOF = False And Rst.BOF = False Then
            
            fwBtnCustFind.Caption = Rst.Fields("FullName")
            Text1.Text = Rst.Fields("Membershipid")
            mvarCustCredit = Rst.Fields("Credit")
            mvarMemberShipId = "«‘ —«ﬂ : " & Rst.Fields("MemberShipId")
            mvarDescription = Rst.Fields("Description")
            blnCreditCust = IIf(Rst!Credit > 0, True, False)
        End If
        
        Set Rst = Nothing
    End If
    FillvsGood
End Sub

