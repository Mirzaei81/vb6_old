VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{7AEDC602-D94C-11D1-BB7A-00E0290EA3C9}#1.0#0"; "ResizeKit.ocx"
Object = "{75D4F767-8785-11D3-93AD-0000832EF44D}#2.14#0"; "FAST2001.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmMojodiControl_4 
   ClientHeight    =   9210
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15105
   Icon            =   "frmMojodiControl_4.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9210
   ScaleWidth      =   15105
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
      Height          =   2400
      Left            =   9120
      RightToLeft     =   -1  'True
      Style           =   1  'Checkbox
      TabIndex        =   13
      Top             =   840
      Width           =   2865
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
      Height          =   2400
      Left            =   12240
      RightToLeft     =   -1  'True
      Style           =   1  'Checkbox
      TabIndex        =   12
      Top             =   840
      Width           =   2655
   End
   Begin VB.Frame Frame3 
      Caption         =   "����"
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
      Left            =   6120
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   600
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
      Caption         =   "�������"
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
      Left            =   6120
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   1500
      Width           =   2775
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
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   6120
      TabIndex        =   6
      Top             =   2520
      Width           =   2775
      Begin VB.CheckBox CheckOrder 
         Alignment       =   1  'Right Justify
         Caption         =   "��� ������� ���� ������"
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
         Height          =   495
         Left            =   240
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   120
         Width           =   2415
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2535
      Left            =   240
      TabIndex        =   1
      Top             =   720
      Width           =   5655
      Begin VB.CommandButton StoreDataUpdate 
         Caption         =   "�� ��� ����� ������ ������� �������"
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   360
         RightToLeft     =   -1  'True
         TabIndex        =   21
         Top             =   1080
         Width           =   2295
      End
      Begin VB.CommandButton cmdInventoryGood_Add 
         Caption         =   " ����� ���� ��� ������ �� �����"
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   1560
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton cmdInventoryGood_Delete 
         Caption         =   " ��� ������ �� �����"
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   19
         Top             =   240
         Width           =   1335
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
         Left            =   3360
         RightToLeft     =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   240
         Width           =   1275
      End
      Begin FLWCtrls.FWProgressBar FWProgressBar1 
         Height          =   375
         Left            =   120
         Top             =   2040
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   661
         BorderStyle     =   10
      End
      Begin MSMask.MaskEdBox txtDateTo 
         Height          =   465
         Left            =   3360
         TabIndex        =   2
         Top             =   1440
         Width           =   1275
         _ExtentX        =   2249
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
         Left            =   3360
         TabIndex        =   3
         Top             =   855
         Width           =   1275
         _ExtentX        =   2249
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
         Caption         =   "��� ����"
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   4560
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Top             =   240
         Width           =   945
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "�� �����"
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
         TabIndex        =   5
         Top             =   1440
         Width           =   945
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "�� �����"
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
         Top             =   855
         Width           =   945
      End
   End
   Begin VSFlex7LCtl.VSFlexGrid vsGood 
      Height          =   5700
      Left            =   180
      TabIndex        =   0
      Top             =   3420
      Width           =   14865
      _cx             =   26220
      _cy             =   10054
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
      BackColor       =   -2147483633
      ForeColor       =   -2147483640
      BackColorFixed  =   16777152
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483634
      BackColorAlternate=   -2147483633
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   3
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
      FormatString    =   $"frmMojodiControl_4.frx":A4C2
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
      Caption         =   "����"
      Alignment       =   2
   End
   Begin FLWCtrls.FWLabel FWLabel1 
      Height          =   495
      Left            =   0
      Top             =   0
      Width           =   15105
      _ExtentX        =   26644
      _ExtentY        =   873
      Enabled         =   -1  'True
      Caption         =   "����� ���� ��� �������"
      FirstColor      =   9412754
      SecondColor     =   14215660
      Angle           =   0
      ForeColor       =   7362318
      BackColor       =   12640511
      FontName        =   "B Homa"
      FontSize        =   15.75
      Alignment       =   2
      Picture         =   "frmMojodiControl_4.frx":A5B6
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   7440
      Top             =   960
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
   Begin RESIZEKITLibCtl.ResizeKit ResizeKit1 
      Height          =   480
      Left            =   0
      OleObjectBlob   =   "frmMojodiControl_4.frx":A5D2
      TabIndex        =   18
      Top             =   0
      Width           =   480
   End
   Begin VB.Label lblGoodLevel2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H000080FF&
      BackStyle       =   0  'Transparent
      Caption         =   "���� ���� ������"
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
      Left            =   9120
      RightToLeft     =   -1  'True
      TabIndex        =   15
      Top             =   500
      Width           =   2745
   End
   Begin VB.Label lblGoodLevel1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H000040C0&
      BackStyle       =   0  'Transparent
      Caption         =   "��� ��"
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
      Left            =   12120
      RightToLeft     =   -1  'True
      TabIndex        =   14
      Top             =   500
      Width           =   2655
   End
End
Attribute VB_Name = "frmMojodiControl_4"
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
    
    mdifrm.Toolbar1.Buttons(15).Enabled = True
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
    
    ReDim Parameter(0) As Parameter
    Parameter(0) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
    Set Rst = RunParametricStoredProcedure2Rec("Get_tGoodLevel1", Parameter)
        
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
    If cmbInventory.ListCount < 1 Then Exit Sub
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
    Parameter(2) = GenerateInputParameter("@Type", adInteger, 4, EnumGoodType.forSale)
    Parameter(3) = GenerateInputParameter("@InventoryNo", adInteger, 4, cmbInventory.ItemData(cmbInventory.ListIndex))
    Parameter(4) = GenerateInputParameter("@Branch", adInteger, 4, cmbBranch.ItemData(cmbBranch.ListIndex))
    Parameter(5) = GenerateInputParameter("@AccountYear", adSmallInt, 2, Val(cmbSalMali.Text))
    Parameter(6) = GenerateInputParameter("@CheckNotZeroMojodi", adInteger, 4, 0)
    Parameter(7) = GenerateInputParameter("@CheckFirstMojodi", adInteger, 4, 0)
    Parameter(8) = GenerateInputParameter("@CheckOrder", adInteger, 4, CheckOrder.Value)
    Parameter(9) = GenerateInputParameter("@Flag", adInteger, 4, 0)
    Parameter(10) = GenerateInputParameter("@SortItem", adInteger, 4, 1)
    
    Set Rst = RunParametricStoredProcedure2Rec("Get_tblTotal_tGood_By_Prams", Parameter)
      
    
    If (Rst.EOF = True And Rst.BOF = True) Then Exit Sub
    
    With vsGood
        
        i = 1
        
        While Rst.EOF = False
            .Rows = .Rows + 1
            .TextMatrix(i, 0) = i
            .TextMatrix(i, 1) = Rst.Fields("GoodCode").Value
            .TextMatrix(i, 2) = left(Rst.Fields("Name").Value, 40)
            .TextMatrix(i, 3) = Rst.Fields("UnitDescription").Value
            .TextMatrix(i, 4) = Rst.Fields("TypeDescription").Value
            If Rst.Fields("FirstMojodi").Value >= 0 Then
               .TextMatrix(i, 5) = Format(Rst.Fields("FirstMojodi").Value, "##.000")
            Else
               .TextMatrix(i, 5) = -Format(Rst.Fields("FirstMojodi").Value, "##.000") & "-"
            End If
            .TextMatrix(i, 6) = Format(Rst.Fields("BuyAmount").Value, "##.000")
            .TextMatrix(i, 7) = Format(Rst.Fields("SaleAmount").Value, "##.000")
            .TextMatrix(i, 8) = Format(Rst.Fields("LossAmount").Value, "##.000")
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
            .TextMatrix(i, 10) = IIf(Rst.Fields("MojodiControl").Value = True, -1, 0)
            .TextMatrix(i, 11) = Rst.Fields("OrderPoint").Value
            .TextMatrix(i, 12) = Rst.Fields("MinValue").Value
            .TextMatrix(i, 13) = Rst.Fields("MaxValue").Value
            .TextMatrix(i, 14) = Rst.Fields("BuyReturnAmount").Value
            .TextMatrix(i, 15) = Format(Rst.Fields("SaleReturnAmount").Value, "##.000")
            
          
            i = i + 1
            Rst.MoveNext
            
        Wend
        Set Rst = Nothing
        
        .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter
        .Cell(flexcpAlignment, 1, 2, .Rows - 1, 2) = flexAlignRightCenter
        .AutoSizeMode = flexAutoSizeColWidth
        .AutoSize 0, .Cols - 1
        
    End With
        
End Sub

Public Sub BeforeUpdate()

End Sub

Public Sub Edit()
    MyFormAddEditMode = EnumAddEditMode.EditMode
    SetFirstToolBar
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
                
                If Val(.TextMatrix(i, 5)) < 0 Then     '
                        Select Case clsStation.Language
                        
                            Case 0
                            
                                frmMsg.fwlblMsg.Caption = "��� �� �����  ������ ����� ��  ���� ���� ������"
                                frmMsg.fwBtn(0).Caption = "����"
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
                        Parameter(1) = GenerateInputParameter("@MojodiControl", adBoolean, 1, IIf(Val(.TextMatrix(i, 10)) = -1, 1, 0))
                        Parameter(2) = GenerateInputParameter("@OrderPoint", adDouble, 8, Val(.TextMatrix(i, 11)))
                        Parameter(3) = GenerateInputParameter("@MinValue", adDouble, 8, Val(.TextMatrix(i, 12)))
                        Parameter(4) = GenerateInputParameter("@MaxValue", adDouble, 8, Val(.TextMatrix(i, 13)))
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
Private Sub cmbBranch_Click()
    FillInventory
End Sub

Private Sub cmbInventory_Click()
    If cmbBranch.ListIndex = -1 Then Exit Sub
    FillLstGoodLevel1
'    FillvsGood
End Sub

Private Sub cmbSalMali_Change()
    If cmbSalMali.Text <> "" Then
        txtDateFrom.Text = Mid(clsDate.shamsi(Date), 3) 'Right(cmbSalMali.Text, 2) & "/01" & "/01"
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
     End If
End Sub

Private Sub cmbSalMali_Click()
    cmbSalMali_Change
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
        frmMsg.fwlblMsg.Caption = "��� ���� ����� � ���� ������ ���� "
        frmMsg.fwBtn(0).ButtonType = flwButtonOk
        frmMsg.fwBtn(1).Visible = False
        frmMsg.fwBtn(0).Caption = "����"
        frmMsg.Show vbModal
        Exit Sub
     End If
        frmMsg.fwlblMsg.Caption = "��� ���� ����� ���� ��� ���� ������ �� ����� ������� �����"
        frmMsg.fwBtn(0).ButtonType = flwButtonOk
        frmMsg.fwBtn(1).ButtonType = flwButtonCancel
        frmMsg.fwBtn(0).Caption = "���"
        frmMsg.fwBtn(1).Caption = "���"
        frmMsg.Show vbModal
        If mvarMsgIdx = vbYes Then
            ReDim Parameter(3) As Parameter
    
            Parameter(0) = GenerateInputParameter("@IntInventoryNo", adInteger, 4, cmbInventory.ItemData(cmbInventory.ListIndex))
            Parameter(1) = GenerateInputParameter("@Level1", adInteger, 4, lstGoodLevel1.ItemData(intSelectedLevel1))
            Parameter(2) = GenerateInputParameter("@Branch", adInteger, 4, cmbBranch.ItemData(cmbBranch.ListIndex))
            Parameter(3) = GenerateInputParameter("@AccountYear", adSmallInt, 2, Val(cmbSalMali.Text))
        
            RunParametricStoredProcedure "Insert_tinventory_Good_All", Parameter
            DefaultSetting
            frmDisMsg.lblMessage = "������ ��� ���� ������ �� ����� ����� �� "
            frmDisMsg.Timer1.Enabled = True
            frmDisMsg.Show vbModal
        End If

End Sub

Private Sub cmdInventoryGood_Delete_Click()
    If cmbBranch.ListIndex = -1 Then Exit Sub
    If cmbInventory.ListIndex = -1 Then Exit Sub
    If cmbSalMali.ListIndex = -1 Then Exit Sub
    Dim intSelectedLevel1 As Integer
    
    frmMsg.fwlblMsg.Caption = "��� ���� ��� ��� ���� ������ �� ����� ������� �����"
    frmMsg.fwBtn(0).ButtonType = flwButtonOk
    frmMsg.fwBtn(1).ButtonType = flwButtonCancel
    frmMsg.fwBtn(0).Caption = "���"
    frmMsg.fwBtn(1).Caption = "���"
    frmMsg.Show vbModal
    If mvarMsgIdx = vbYes Then
        ReDim Parameter(2) As Parameter

        Parameter(0) = GenerateInputParameter("@IntInventoryNo", adInteger, 4, cmbInventory.ItemData(cmbInventory.ListIndex))
        Parameter(1) = GenerateInputParameter("@Branch", adInteger, 4, cmbBranch.ItemData(cmbBranch.ListIndex))
        Parameter(2) = GenerateInputParameter("@AccountYear", adSmallInt, 2, Val(cmbSalMali.Text))
        
        RunParametricStoredProcedure "Delete_tinventory_Good_All", Parameter
        DefaultSetting
        frmDisMsg.lblMessage = "��� ��� ���� ������ �� ����� ����� �� "
        frmDisMsg.Timer1.Enabled = True
        frmDisMsg.Show vbModal
    End If

End Sub

Private Sub Form_Activate()
    VarActForm = Me.Name
    Frame3.BackColor = Me.BackColor
    Frame28.BackColor = Me.BackColor
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call PresetScreenSaver
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call PresetScreenSaver
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

    If ClsFormAccess.frmMojodiControl_4 = False Then
        Unload Me
        Exit Sub
    End If
        
    If intVersion = Normal Or intVersion = Min Then
        ShowDisMessage "������ ����� ������� ������� �� ���� ��� ������� � ������ ���� ����", 1500
        Unload Me
        Exit Sub
    End If
    CenterTop Me
    VarActForm = Me.Name
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

    formloadFlag = True

    SetFirstToolBar
    ChangeLanguage
    
    DefaultSetting
    
    txtDateFrom.Text = Mid(clsDate.shamsi(Date), 3, 2) & "/01" & "/01"
    txtDateTo.Text = Mid(clsDate.shamsi(Date), 3)
    
End Sub


Private Sub Form_Unload(Cancel As Integer)
    
    VarActForm = ""

    Dim i As Integer
    
    AllButton vbOff, True
    
    If PosConnection.State = adStateOpen Then PosConnection.Close: Set PosConnection = Nothing
    Unload frmFindGoods

    SaveSetting strMainKey, Me.Name, "Left", Me.left
    SaveSetting strMainKey, Me.Name, "Top", Me.top



End Sub

Private Sub CheckOrder_Click()
    FillvsGood
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
    
    FillvsGood
    
    MyFormAddEditMode = EnumAddEditMode.ViewMode
    SetFirstToolBar
    
End Sub

Private Sub lstGoodLevel1_Scroll()
    FillLstGoodLevel2
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
            
            lblGoodLevel1.Caption = " ���� ���� ������ - ��� ��"
            lblGoodLevel2.Caption = "���� ���� ������"
            
    End Select
    
'    lstGoodLevel1.Left = Me.Width - (lstGoodLevel1.Left + lstGoodLevel1.Width)
'    lstGoodLevel2.Left = Me.Width - (lstGoodLevel2.Left + lstGoodLevel2.Width)
    
'    lblGoodLevel1.Left = Me.Width - (lblGoodLevel1.Left + lblGoodLevel1.Width)
'    lblGoodLevel2.Left = Me.Width - (lblGoodLevel2.Left + lblGoodLevel2.Width)
        
    
    With vsGood
    
        .Cols = 17
        
        Select Case clsStation.Language
            Case Farsi
                .TextMatrix(0, 0) = "����"
                .TextMatrix(0, 1) = "��"
                .TextMatrix(0, 2) = "��� ����"
                .TextMatrix(0, 3) = "���� "
                .TextMatrix(0, 4) = "��� ����"
                .TextMatrix(0, 5) = "������ �����"
                .TextMatrix(0, 6) = "���� "
                .TextMatrix(0, 7) = "���� "
                .TextMatrix(0, 8) = "������"
                .TextMatrix(0, 9) = "��������� ����"
                .TextMatrix(0, 10) = " ����� "
                .TextMatrix(0, 11) = "���� ����� "
                .TextMatrix(0, 12) = "����� �����"
                .TextMatrix(0, 13) = "�����������"
                .TextMatrix(0, 14) = "�ѐ�� �� ����"
                .TextMatrix(0, 15) = "�ѐ�� �� ����"
                .TextMatrix(0, 16) = "    "
            
            Case English
                .TextMatrix(0, 0) = "Row"
                .TextMatrix(0, 1) = "Code"
                .TextMatrix(0, 2) = "Name"
                .TextMatrix(0, 3) = " Unit"
                .TextMatrix(0, 4) = " Type"
                .TextMatrix(0, 5) = "FirstStock"
                .TextMatrix(0, 6) = "Purchase"
                .TextMatrix(0, 7) = "Sale"
                .TextMatrix(0, 8) = "Losses"
                .TextMatrix(0, 9) = "Remaining"
                .TextMatrix(0, 10) = "Control"
                .TextMatrix(0, 11) = "Ordre"
                .TextMatrix(0, 12) = "Minimum"
                .TextMatrix(0, 13) = "Maximum"
                .TextMatrix(0, 14) = "PurchaseReturn"
                .TextMatrix(0, 15) = "SaleReturn"
                .TextMatrix(0, 16) = "    "
       End Select
        
   '     .ColSort(5) = flexSortNumericAscending + flexSortNumericDescending
        .ColAlignment(-1) = flexAlignCenterCenter
        .FocusRect = flexFocusHeavy
       ' .ColHidden(1) = True
        .ColHidden(6) = True
       ' .ColHidden(8) = True
       ' .ColHidden(10) = True
        .ColHidden(11) = True
        .ColHidden(12) = True
        .ColHidden(13) = True
        .ColHidden(14) = True
        
        .AutoSizeMode = flexAutoSizeColWidth
        .AutoSize 0, .Cols - 1
        .AutoSearch = flexSearchFromCursor
    End With
    FillBranch
    FillInventory
    FillSalMali
    DefaultSetting
            
    SetFirstToolBar


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
    If cmbBranch.ListCount > 0 Then cmbBranch.ListIndex = 0
    

End Sub
Private Sub FillInventory()
    cmbInventory.Clear
    If cmbBranch.ListIndex = -1 Then Exit Sub
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
    If cmbInventory.ListCount > 0 Then cmbInventory.ListIndex = 0

End Sub

Private Sub ResizeKit1_ExitResize(ByVal XScale As Double, ByVal YScale As Double)


    If formloadFlag = True Then
        SaveSetting strMainKey, Me.Name, "Height", Me.Height
        SaveSetting strMainKey, Me.Name, "Width", Me.Width
    End If

End Sub

Private Sub StoreDataUpdate_Click()
   If cmbInventory.ListIndex = -1 Then Exit Sub
 
    If Len(txtDateFrom.ClipText) <> 6 Or Len(txtDateTo.ClipText) <> 6 Then
        frmDisMsg.lblMessage = " ����� ����� ���� ���� "
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
        Parameter(3) = GenerateInputParameter("@SystemTime", adVarWChar, 50, Mid(Str(time), 1, 5))
        Parameter(4) = GenerateInputParameter("@DateBefore", adVarWChar, 50, CStr(IIf(Trim(txtDateFrom.ClipText) = "", "", Trim(txtDateFrom.Text))))
        Parameter(5) = GenerateInputParameter("@DateAfter", adVarWChar, 50, CStr(IIf(Trim(txtDateTo.ClipText) = "", "", Trim(txtDateTo.Text))))
        Parameter(6) = GenerateInputParameter("@Type", adInteger, 4, EnumGoodType.forSale)
        Parameter(7) = GenerateInputParameter("@InVentoryNo1", adInteger, 4, cmbInventory.ItemData(cmbInventory.ListIndex))
        Parameter(8) = GenerateInputParameter("@InVentoryNo2", adInteger, 4, cmbInventory.ItemData(cmbInventory.ListIndex))
        Parameter(9) = GenerateInputParameter("@Branch", adInteger, 4, cmbBranch.ItemData(cmbBranch.ListIndex))
        Parameter(10) = GenerateInputParameter("@UsePercentFlag", adInteger, 4, 0)
        Parameter(11) = GenerateInputParameter("@AccountYear", adSmallInt, 2, Val(cmbSalMali.Text))
        
        DoEvents
        RunParametricStoredProcedure "Update_tblTotal_tInventory_tGood_For_Mojodi", Parameter
        FWProgressBar1.Value = 100

''        Set Rst = RunParametricStoredProcedure2Rec("GetInventoryAtomicReport_Mojodi", Parameter, Cnn)
''
''   If (Rst.EOF = True And Rst.BOF = True) Then Exit Sub
''
''
''        While Rst.EOF = False
''
''            ReDim Parameter(10) As Parameter
''
''            Parameter(0) = GenerateInputParameter("@GoodCode", adInteger, 4, Rst.Fields("GoodCode").Value)
''            Parameter(1) = GenerateInputParameter("@BuyAmount", adDouble, 8, Rst.Fields("BuyAmount").Value)
''            Parameter(2) = GenerateInputParameter("@SaleAmount", adDouble, 8, Rst.Fields("SaleAmount").Value)
''            Parameter(3) = GenerateInputParameter("@LossAmount", adDouble, 8, Rst.Fields("LossAmount").Value)
''            Parameter(4) = GenerateInputParameter("@BuyReturnAmount", adDouble, 8, Rst.Fields("BuyReturnAmount").Value)
''            Parameter(5) = GenerateInputParameter("@SaleReturnAmount", adDouble, 8, Rst.Fields("SaleReturnAmount").Value)
''            Parameter(6) = GenerateInputParameter("@FromStoreAmount", adDouble, 8, Rst.Fields("FromStoreAmount").Value)
''            Parameter(7) = GenerateInputParameter("@toStoreAmount", adDouble, 8, Rst.Fields("toStoreAmount").Value)
''            Parameter(8) = GenerateInputParameter("@Mojodi", adDouble, 8, Rst.Fields("Mojodi").Value)
''            Parameter(9) = GenerateInputParameter("@InVentoryNo", adInteger, 4, cmbInventory.ItemData(cmbInventory.ListIndex))
''            Parameter(10) = GenerateInputParameter("@Branch", adInteger, 4, cmbBranch.ItemData(cmbBranch.ListIndex))
''
''            RunParametricStoredProcedure "Update_Calculated_Store", Parameter, Cnn
''
''            Rst.MoveNext
''            FWProgressBar1.Value = FWProgressBar1.Value + 1
''            If FWProgressBar1.Value = 100 Then
''               FWProgressBar1.Value = 0
''            End If
''
''        Wend
''        Set Rst = Nothing
        DefaultSetting
        FWProgressBar1.Value = 0
        StoreDataUpdate.Enabled = True
        frmDisMsg.lblMessage = " �� ��� ����� ����� �� "
        frmDisMsg.Timer1.Enabled = True
        frmDisMsg.Show vbModal

End Sub

Private Sub vsGood_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    With vsGood
        If (.TextMatrix(Row, 0) = "*" Or MyFormAddEditMode = EnumAddEditMode.EditMode) And Col > 1 And tmpTextMatrix <> .TextMatrix(Row, Col) Then
        
            If MyFormAddEditMode = EnumAddEditMode.EditMode And InStr(.TextMatrix(Row, 0), "*") = 0 Then
                .TextMatrix(Row, 0) = Trim(.TextMatrix(Row, 0)) & "*"
            End If
            
        Else

        End If
        .AutoSizeMode = flexAutoSizeColWidth
        .AutoSize 0, .Cols - 1
        

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
    With vsGood
        If (.TextMatrix(.Row, 0) = "*" Or MyFormAddEditMode = EnumAddEditMode.EditMode) And (.Col = 5 Or .Col > 9) Then
            If .Col = 5 And ClsFormAccess.FirstMojodiControl = True Then
               .Select .Row, .Col
               .EditCell
            ElseIf .Col = 5 Then
                ShowDisMessage "��� ����� ������ �� ��� ���� �� ������", 2000
            End If
            If .Col > 9 Then
               .Select .Row, .Col
               .EditCell
            End If
        End If
    
    End With

End Sub
Private Sub vsGood_DblClick()
    With vsGood
        If (.TextMatrix(.Row, IdxColRow) = "*" Or MyFormAddEditMode = EnumAddEditMode.EditMode) And (.Col = IdxColFirstStock Or .Col > IdxColGoodUnit) Then
            If .Col = 5 And ClsFormAccess.FirstMojodiControl = True Then
               .Select .Row, .Col
               .EditCell
            ElseIf .Col = 5 Then
                ShowDisMessage "��� ����� ������ �� ��� ���� �� ������", 2000
            End If
            If .Col > 9 Then
               .Select .Row, .Col
               .EditCell
            End If
        Else
            If .Col = 2 Then
                Load frmGoodTurnOver
                frmGoodTurnOver.cmbBranch.ListIndex = cmbBranch.ListIndex
                frmGoodTurnOver.cmbInventory.ListIndex = cmbInventory.ListIndex
                frmGoodTurnOver.cmbSalMali.ListIndex = cmbSalMali.ListIndex
                
                frmGoodTurnOver.fwBtnGoodFind.Caption = .TextMatrix(.Row, 2)
                frmGoodTurnOver.fwBtnGoodFind.Tag = .TextMatrix(.Row, 1)
                frmGoodTurnOver.txtDateFrom.Text = txtDateFrom.Text
                frmGoodTurnOver.txtDateTo.Text = txtDateTo.Text
                frmGoodTurnOver.StoreDataUpdate.Enabled = True
'                frmGoodTurnOver.txtBarcode.Text = .TextMatrix(.Row, 3)
                frmGoodTurnOver.StoreDataUpdate_Click
                frmGoodTurnOver.Show
            End If
        End If
    End With

End Sub

Private Sub vsGood_KeyDown(KeyCode As Integer, Shift As Integer)
    With vsGood
        If (.TextMatrix(.Row, 0) = "*" Or MyFormAddEditMode = EnumAddEditMode.EditMode) And (.Col = 5 Or .Col > 9) Then
            If .Col = 5 And ClsFormAccess.FirstMojodiControl = True Then
               .Select .Row, .Col
               .EditCell
            ElseIf .Col = 5 Then
                ShowDisMessage "��� ����� ������ �� ��� ���� �� ������", 2000
            End If
            If .Col > 9 Then
               .Select .Row, .Col
               .EditCell
            End If
        End If
    End With
    
End Sub


Private Sub vsGood_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    With vsGood
        If KeyAscii = 39 Then KeyAscii = 0
        
        If (Col <> 5 And Col < 10) Or (IsNumeric(Chr(KeyAscii)) = False And KeyAscii = 8) Then
            
            KeyAscii = 0
            
        ElseIf IsNumeric(Chr(KeyAscii)) = False Then
            
            KeyAscii = 0
            
        ElseIf (Col <> 5 And Col < 10) Or KeyAscii = 8 Then
            
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
With vsGood
    
    RunNonParametricStoredProcedure "Delete_tblPrint_Order"
    
    ReDim Parameter(12) As Parameter
    For i = 1 To .Rows - 1
        
        Parameter(0) = GenerateInputParameter("@Row", adInteger, 4, i)
        Parameter(1) = GenerateInputParameter("@GoodName", adVarChar, 50, .TextMatrix(i, 2))
        Parameter(2) = GenerateInputParameter("@UnitName", adVarChar, 50, .TextMatrix(i, 3))
        Parameter(3) = GenerateInputParameter("@Mojodi", adDouble, 8, .TextMatrix(i, 9))
        Parameter(4) = GenerateInputParameter("@OrderPoint", adDouble, 8, Val(.TextMatrix(i, 11)))
        Parameter(5) = GenerateInputParameter("@Minimum", adDouble, 8, Val(.TextMatrix(i, 12)))
        Parameter(6) = GenerateInputParameter("@Maximum", adDouble, 8, Val(.TextMatrix(i, 13)))
        Parameter(7) = GenerateInputParameter("@BuyPrice", adInteger, 4, 0)
        Parameter(8) = GenerateInputParameter("@Sellprice", adInteger, 4, 0)
        Parameter(9) = GenerateInputParameter("@Sellprice2", adInteger, 4, 0)
        Parameter(10) = GenerateInputParameter("@Sellprice3", adInteger, 4, 0)
        Parameter(11) = GenerateInputParameter("@Barcode", adVarChar, 20, "")
        Parameter(12) = GenerateInputParameter("@FirstMojodi", adDouble, 8, .TextMatrix(i, 5))
        
        RunParametricStoredProcedure "Insert_tblPrint_Order", Parameter
    Next i
    ReDim Parameter(2) As Parameter
    Parameter(0) = GenerateInputParameter("@SystemDate", adVarWChar, 50, Mid(clsDate.shamsi(Date), 3, 8))
    Parameter(1) = GenerateInputParameter("@SystemDay", adVarWChar, 50, clsDate.Find_DayOfWeek(Weekday(Date, vbSaturday)))
    Parameter(2) = GenerateInputParameter("@SystemTime", adVarWChar, 50, Mid(Str(time), 1, 5))
    
  '  CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\RepOrder.rpt"
    CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\RepOrder_A4.rpt"
    Dim fileSystem As New FileSystemObject
    Dim IsFileExist As Boolean
    IsFileExist = fileSystem.FileExists(CrystalReport1.ReportFileName)
    If IsFileExist = False Then
            frmDisMsg.lblMessage = " ����  " & CrystalReport1.ReportFileName & "���� ��� "
            frmDisMsg.Timer1.Interval = 3000
            frmDisMsg.Timer1.Enabled = True
            frmDisMsg.Show vbModal
            Exit Sub
    End If
    
    CrystalReport1.ReportTitle = "����� ���� ������� ������� "
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

End With
End Sub

