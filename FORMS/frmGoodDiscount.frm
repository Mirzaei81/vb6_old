VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{7AEDC602-D94C-11D1-BB7A-00E0290EA3C9}#1.0#0"; "ResizeKit.ocx"
Object = "{75D4F767-8785-11D3-93AD-0000832EF44D}#2.14#0"; "FAST2001.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmGoodDiscount 
   ClientHeight    =   9255
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   15105
   Icon            =   "frmGoodDiscount.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9255
   ScaleWidth      =   15105
   Begin VB.Frame Frame1 
      Height          =   2775
      Left            =   240
      TabIndex        =   10
      Top             =   600
      Width           =   6495
      Begin VB.TextBox txtDiscount 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Nazanin"
            Size            =   14.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4200
         TabIndex        =   16
         Text            =   "0"
         Top             =   240
         Width           =   1095
      End
      Begin VB.CheckBox ChkDutySale 
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
         Left            =   4320
         TabIndex        =   21
         Top             =   2160
         Width           =   855
      End
      Begin VB.CheckBox ChkDutyBuy 
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
         Left            =   4320
         TabIndex        =   20
         Top             =   1680
         Width           =   855
      End
      Begin VB.CheckBox ChkTaxSale 
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
         Left            =   4320
         TabIndex        =   19
         Top             =   1200
         Width           =   855
      End
      Begin VB.CheckBox ChkTaxBuy 
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
         Left            =   4320
         TabIndex        =   18
         Top             =   720
         Width           =   855
      End
      Begin VB.CommandButton NewDiscountUpdate 
         BackColor       =   &H00008000&
         Caption         =   " ⁄ÌÌ‰ œ—’œ  Œ›Ì› —ÊÌ «Ì‰  ê—ÊÂ ﬂ«·«Â« "
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
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   240
         Width           =   3915
      End
      Begin VB.CommandButton NewDutyBuyUpdate 
         BackColor       =   &H00008000&
         Caption         =   " ⁄ÌÌ‰ ⁄Ê«—÷ Œ—Ìœ —ÊÌ «Ì‰  ê—ÊÂ ﬂ«·«Â«"
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
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   1680
         Width           =   3915
      End
      Begin VB.CommandButton NewDutySaleUpdate 
         BackColor       =   &H00008000&
         Caption         =   " ⁄ÌÌ‰ ⁄Ê«—÷ ›—Ê‘ —ÊÌ «Ì‰  ê—ÊÂ ﬂ«·«Â« "
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
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   2160
         Width           =   3915
      End
      Begin VB.CommandButton NewTaxSaleUpdate 
         BackColor       =   &H00008000&
         Caption         =   " ⁄ÌÌ‰ „«·Ì«  ›—Ê‘ —ÊÌ «Ì‰  ê—ÊÂ ﬂ«·«Â« "
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
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   1200
         Width           =   3915
      End
      Begin VB.CommandButton NewTaxBuyUpdate 
         BackColor       =   &H00008000&
         Caption         =   " ⁄ÌÌ‰ „«·Ì«  Œ—Ìœ —ÊÌ «Ì‰  ê—ÊÂ ﬂ«·«Â«"
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
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   720
         Width           =   3915
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "œ—’œ  Œ›Ì›"
         BeginProperty Font 
            Name            =   "Nazanin"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   5400
         RightToLeft     =   -1  'True
         TabIndex        =   17
         Top             =   240
         Width           =   585
      End
   End
   Begin RESIZEKITLibCtl.ResizeKit ResizeKit1 
      Height          =   480
      Left            =   6000
      OleObjectBlob   =   "frmGoodDiscount.frx":A4C2
      TabIndex        =   9
      Top             =   600
      Width           =   480
   End
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
      Left            =   6840
      MultiLine       =   -1  'True
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   2160
      Width           =   2025
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   6360
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
      Left            =   12330
      RightToLeft     =   -1  'True
      Style           =   1  'Checkbox
      TabIndex        =   1
      Top             =   960
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
      Height          =   2400
      Left            =   9360
      RightToLeft     =   -1  'True
      Style           =   1  'Checkbox
      TabIndex        =   0
      Top             =   960
      Width           =   2745
   End
   Begin VSFlex7LCtl.VSFlexGrid vsGood 
      Height          =   5700
      Left            =   180
      TabIndex        =   2
      Top             =   3480
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
      BackColor       =   -2147483624
      ForeColor       =   -2147483640
      BackColorFixed  =   16777152
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
      Left            =   13440
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
   Begin FLWCtrls.FWLabel FWLabel1 
      Height          =   495
      Left            =   0
      Top             =   0
      Width           =   14985
      _ExtentX        =   26432
      _ExtentY        =   873
      Enabled         =   -1  'True
      Caption         =   " Œ›Ì› —ÊÌ ﬂ«·«Â«"
      FirstColor      =   9412754
      SecondColor     =   14215660
      Angle           =   0
      ForeColor       =   7362318
      BackColor       =   12640511
      FontName        =   "B Homa"
      FontSize        =   15.75
      Alignment       =   2
      Picture         =   "frmGoodDiscount.frx":A548
   End
   Begin FLWCtrls.FWCoolButton fwBtnCustFind 
      Height          =   930
      Left            =   6840
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   960
      Width           =   2370
      _ExtentX        =   4180
      _ExtentY        =   1640
      BackColor       =   16777088
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "B Homa"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "frmGoodDiscount.frx":A564
      PictureAlign    =   4
      Caption         =   " «„Ì‰ ﬂ‰‰œÂ"
      MaskColor       =   -2147483633
   End
   Begin FLWCtrls.FWProgressBar FWProgressBar1 
      Height          =   495
      Left            =   6840
      Top             =   2760
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   873
      BorderStyle     =   10
   End
   Begin VB.Label Label2 
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
      Left            =   8760
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   2160
      Width           =   585
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   " «„Ì‰ ﬂ‰‰œÂ"
      BeginProperty Font 
         Name            =   "B Homa"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7080
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   600
      Width           =   2175
   End
   Begin VB.Label lblGoodLevel1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H000040C0&
      BackStyle       =   0  'Transparent
      Caption         =   "ê—ÊÂ «’·Ì ò«·«Â«"
      BeginProperty Font 
         Name            =   "Titr"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   435
      Left            =   12360
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   480
      Width           =   2655
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
      ForeColor       =   &H8000000D&
      Height          =   435
      Left            =   9360
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   480
      Width           =   2745
   End
End
Attribute VB_Name = "frmGoodDiscount"
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
    
    FillLstGoodLevel1
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
    
    MyFormAddEditMode = EnumAddEditMode.ViewMode 'VIEW Mode
    SetFirstToolBar
    
    vsGood.Rows = 1
    If lstGoodLevel1.ListCount < 1 Then Exit Sub
    If lstGoodLevel2.ListCount < 1 Then Exit Sub
    
    Dim i As Long
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
    
    If intSelectedLevel1 <> -1 And intSelectedLevel2 <> -1 Then
        
            strSelectedLevels = Right(strSelectedLevels, Len(strSelectedLevels) - 1)
            ReDim Parameter(3) As Parameter
            Parameter(0) = GenerateInputParameter("@Level1", adInteger, 4, lstGoodLevel1.ItemData(intSelectedLevel1))
            Parameter(1) = GenerateInputParameter("@strSelectedLevels", adVarWChar, 4000, strSelectedLevels)
            Parameter(2) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
            Parameter(3) = GenerateInputParameter("@ProductCompany", adInteger, 4, Val(fwBtnCustFind.Tag))
            Set Rst = RunParametricStoredProcedure2Rec("Get_Good_In_Levels", Parameter)
        
    ElseIf intSelectedLevel1 <> -1 And intSelectedLevel2 = -1 Then  'Or intSelectedLevel2 = -1
            
            ReDim Parameter(2) As Parameter
            Parameter(0) = GenerateInputParameter("@GoodLevel1Code", adInteger, 4, lstGoodLevel1.ItemData(intSelectedLevel1))
            Parameter(1) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
            Parameter(2) = GenerateInputParameter("@ProductCompany", adInteger, 4, Val(fwBtnCustFind.Tag))
            Set Rst = RunParametricStoredProcedure2Rec("GetVw_GoodInfo", Parameter)
    Else
            ReDim Parameter(1) As Parameter
            Parameter(0) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
            Parameter(1) = GenerateInputParameter("@ProductCompany", adInteger, 4, Val(fwBtnCustFind.Tag))
            Set Rst = RunParametricStoredProcedure2Rec("GetVwGoodInfo", Parameter)
       
    End If
    If (Rst.EOF = True And Rst.BOF = True) Then Exit Sub
    
    With vsGood
        
        i = 1
        
        While Rst.EOF = False
            .Rows = .Rows + 1
            .TextMatrix(i, 0) = i
            .TextMatrix(i, 1) = Rst.Fields("Code").Value
            .TextMatrix(i, 2) = Rst.Fields("Name").Value
            .TextMatrix(i, 3) = Rst.Fields("UnitDescription").Value
            .TextMatrix(i, 4) = Rst.Fields("TypeDescription").Value
            .TextMatrix(i, 5) = Rst.Fields("Barcode").Value
            .TextMatrix(i, 6) = Rst.Fields("Discount").Value
            .TextMatrix(i, 7) = IIf(Rst.Fields("DutyBuy").Value = True, -1, 0)
            .TextMatrix(i, 8) = IIf(Rst.Fields("DutySale").Value = True, -1, 0)
            .TextMatrix(i, 9) = IIf(Rst.Fields("TaxBuy").Value = True, -1, 0)
            .TextMatrix(i, 10) = IIf(Rst.Fields("TaxSale").Value = True, -1, 0)
            
            i = i + 1
            Rst.MoveNext
            
        Wend
        Set Rst = Nothing
        
        .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter
       ' .AutoSizeMode = flexAutoSizeColWidth
       ' .AutoSize 0, .Cols - 1
        
    End With
        
End Sub

Public Sub BeforeUpdate()

End Sub

Public Sub Edit()
    With vsGood
        
 '       .Editable = flexEDKbdMouse

        MyFormAddEditMode = EnumAddEditMode.EditMode
        SetFirstToolBar
    End With
End Sub

Public Sub Update()
    
    Dim i As Long
    Dim j As Long
    Dim LongTemp As Integer
    Dim lngSelectedGroup  As Long
    Dim lngSelectedSubGroup  As Long
    
    Dim Rst As New ADODB.Recordset
    
    
    lngSelectedGroup = -1
    lngSelectedSubGroup = -1
    
    If MyFormAddEditMode = EnumAddEditMode.ViewMode Then Exit Sub
    
    vsGood_ValidateEdit vsGood.Row, vsGood.Col, False
    
    With vsGood
        If .Rows < 2 Then
            MyFormAddEditMode = EnumAddEditMode.ViewMode
            SetFirstToolBar
            Exit Sub
        End If
        
'        For i = 1 To .Rows - 1
'            .Row = i
'            If InStr(.TextMatrix(i, 0), "*") > 0 Then 'new or edited records
'
'                If Val(.TextMatrix(i, 6)) > 100 Then     '
'                        Select Case clsStation.Language
'
'                            Case 0
'
'                                frmMsg.fwlblMsg.Caption = " Œ›Ì› »“—ê — «“ 100 œ—’œ ﬁ«»· ﬁ»Ê· ‰Ì” "
'                                frmMsg.fwBtn(0).Caption = "ﬁ»Ê·"
'                            Case 1
'
'                                frmMsg.fwlblMsg.Caption = "Error In Discount Greater Than 100"
'                                frmMsg.fwBtn(0).Caption = "Ok"
'                                frmMsg.fwlblMsg.Alignment = vbLeftJustify
'
'                        End Select
'
'                        frmMsg.fwBtn(0).ButtonType = flwButtonOk
'                        frmMsg.fwBtn(1).Visible = False
'                        frmMsg.Show vbModal
'
'                        Exit Sub
'
'                End If
'
'            End If
'        Next i
        
        For j = 0 To lstGoodLevel1.ListCount - 1
            If lstGoodLevel1.Selected(j) = True Then
                lngSelectedGroup = j
                Exit For
            End If
        Next j
        For j = 0 To lstGoodLevel2.ListCount - 1
            If lstGoodLevel2.Selected(j) = True Then
                lngSelectedSubGroup = j
                Exit For
            End If
        Next j

        Select Case MyFormAddEditMode
                
            Case EnumAddEditMode.EditMode
                If ValidateUIData() = True Then
                    If UpdateChanges() = True Then
                        FillvsGood
                        ShowDisMessage " €ÌÌ—«  »« „Ê›ﬁÌ  À»  ‘œ", 2000
                    End If
                End If
            
        End Select
        
    End With
    Set Rst = Nothing
End Sub


Public Sub Cancel()

    MyFormAddEditMode = EnumAddEditMode.ViewMode
    SetFirstToolBar
    
    FillvsGood
    
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

    If ClsFormAccess.frmGoodDiscount = False Then
        Unload Me
        Exit Sub
    End If
    
    If intVersion = Min Then
        ShowDisMessage "«” ›«œÂ «“  Œ›Ì›«  ê—ÊÂÌ Ê„«·Ì«  Ê ⁄Ê«—÷ œ— ‰”ŒÂ „ Ê”ÿ Ê »«·« — «„ﬂ«‰ Å–Ì— «” ", 1500
        Unload Me
        Exit Sub
    End If
    
    CenterTop Me
    VarActForm = Me.Name
    
    
    
    ChangeLanguage
    fwBtnCustFind.Tag = -1
    UpdatelblSupplier
    DefaultSetting
    
    fwBtnCustFind.Visible = False
    
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

    Dim i As Integer
    
    AllButton vbOff, True
    
    SaveSetting strMainKey, Me.Name, "Left", Me.Left
    SaveSetting strMainKey, Me.Name, "Top", Me.Top


    If PosConnection.State = adStateOpen Then PosConnection.Close: Set PosConnection = Nothing
End Sub

Private Sub fwBtnCustFind_Click()
    Me.FindCust
    FillvsGood
End Sub
Public Sub FindCust()
    
    frmFindSupplier.Show vbModal
    
    If mvarcode <> 0 Then
        fwBtnCustFind.Tag = mvarcode
        mvarcode = 0
    Else
        fwBtnCustFind.Tag = -1
    End If
    UpdatelblSupplier
   
End Sub
Private Sub UpdatelblSupplier()

    If fwBtnCustFind.Tag <> "" Then
        Dim Rst As New ADODB.Recordset
        Dim mvarMemberShipId, mvarTel, mvarAddress, mvarDescription As String
        
    
        ReDim Parameter(0) As Parameter
        Parameter(0) = GenerateInputParameter("@Code", adInteger, 4, Val(fwBtnCustFind.Tag))
        Set Rst = RunParametricStoredProcedure2Rec("Get_vw_Suppliers", Parameter)
        
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
            mvarMemberShipId = "«‘ —«ﬂ : " & Rst.Fields("MemberShipId")
            mvarDescription = Rst.Fields("Description")
           
            
        End If
        
        Set Rst = Nothing
    End If
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
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call PresetScreenSaver
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call PresetScreenSaver
End Sub

Public Sub ChangeLanguage()

Dim Obj As Object

    Select Case clsStation.Language    ' LCase(mdifrm.Toolbar1.Buttons(25).Key)
        
        Case English
            
            
            Me.Caption = "Good Discount"
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
            
            
            Me.Caption = " ⁄—Ì›  Œ›Ì› ò«·«"
            mdifrm.Caption = clsArya.Company
            Me.RightToLeft = True
            
            For Each Obj In Me
                On Error Resume Next
                    Obj.RightToLeft = True
                On Error GoTo 0
            Next Obj
            
            lblGoodLevel1.Caption = " (ê—ÊÂ «’·Ì ò«·«Â«(»Œ‘ Â«"
            lblGoodLevel2.Caption = "ê—ÊÂ ›—⁄Ì ò«·«Â«"
            
    End Select
    
'    lstGoodLevel1.Left = Me.Width - (lstGoodLevel1.Left + lstGoodLevel1.Width)
'    lstGoodLevel2.Left = Me.Width - (lstGoodLevel2.Left + lstGoodLevel2.Width)
    
'    lblGoodLevel1.Left = Me.Width - (lblGoodLevel1.Left + lblGoodLevel1.Width)
'    lblGoodLevel2.Left = Me.Width - (lblGoodLevel2.Left + lblGoodLevel2.Width)
        
    
    With vsGood
    
        .Cols = 11
        
        Select Case clsStation.Language
            Case Farsi
                .TextMatrix(0, 0) = "—œÌ›"
                .TextMatrix(0, 1) = "òœ"
                .TextMatrix(0, 2) = "‰«„ ò«·«"
                .TextMatrix(0, 3) = "Ê«Õœ "
                .TextMatrix(0, 4) = "‰Ê⁄ ﬂ«·«"
                .TextMatrix(0, 5) = "»«—ﬂœ"
                .TextMatrix(0, 6) = "œ—’œ  Œ›Ì›"
                .TextMatrix(0, 7) = "⁄Ê«—÷ Œ—Ìœ"
                .TextMatrix(0, 8) = "⁄Ê«—÷ ›—Ê‘"
                .TextMatrix(0, 9) = "„«·Ì«  Œ—Ìœ"
                .TextMatrix(0, 10) = "„«·Ì«  ›—Ê‘"
            
            Case English
                .TextMatrix(0, 0) = "Row"
                .TextMatrix(0, 1) = "Code"
                .TextMatrix(0, 2) = "Name"
                .TextMatrix(0, 3) = " Unit"
                .TextMatrix(0, 4) = " Type"
                .TextMatrix(0, 5) = " Barcode"
                .TextMatrix(0, 6) = "Discount"
                .TextMatrix(0, 7) = "Duty_Buy"
                .TextMatrix(0, 8) = "Duty_Sale"
                .TextMatrix(0, 9) = "Tax_Buy"
                .TextMatrix(0, 10) = "Tax_Sale"
            
       End Select
         .ColDataType(7) = flexDTBoolean
         .ColDataType(8) = flexDTBoolean
         .ColDataType(9) = flexDTBoolean
         .ColDataType(10) = flexDTBoolean
   '     .ColSort(5) = flexSortNumericAscending + flexSortNumericDescending
        .ColAlignment(-1) = flexAlignCenterCenter
        .ColAlignment(2) = flexAlignRightCenter
        .FocusRect = flexFocusHeavy
      '  .ColHidden(1) = True
'        .AutoSizeMode = flexAutoSizeColWidth
'        .AutoSize 0, .Cols - 1
        .AutoSearch = flexSearchFromCursor
         For i = 0 To .Cols - 1
             .ColWidth(i) = Val(GetSetting(strMainKey, "frmGoodDiscount_vsGood", "Col" & i))
            If .ColWidth(i) = 0 Then
                .ColWidth(i) = .Width / 10       'Row
            End If
         Next i
         
    End With

    FillLstGoodLevel1
            
    SetFirstToolBar

End Sub

Private Sub NewDiscountUpdate_Click()
    
    On Error GoTo Err_Handler
    
    Dim Discount As Double
    Dim strDiscount As String
    strDiscount = Trim(txtDiscount.Text)
    
    If IsNumeric(strDiscount) = False Then
        ShowMessage "„ﬁœ«— Ê«—œ ‘œÂ »—«Ì  Œ›Ì› ﬂ«·« „⁄ »— ‰Ì” . ·ÿ›« Ìﬂ „ﬁœ«— „⁄ »— —« Ê«—œ ﬂ‰Ìœ", True, False, " «ÌÌœ", ""
        Exit Sub
    End If
    
    Discount = Val(strDiscount)
    
    If Discount > 100 Or Discount < 0 Then
        ShowMessage " «—ﬁ«„ „⁄ »— Ê«—œ ﬂ‰Ìœ ", True, False, "ﬁ»Ê·", ""
        Exit Sub
    End If


    Me.MousePointer = vbHourglass
    Dim GoodCodesString As String
    Dim GoodCodesLength As Long
    Dim HadGreaterLength As Boolean
    HadGreaterLength = False
    GoodCodesString = ""

    With vsGood
        For i = 1 To .Rows - 1
            GoodCodesString = GoodCodesString & .TextMatrix(i, 1) & STRING_DELIMITER
            GoodCodesLength = Len(GoodCodesString)
            If GoodCodesLength > 3800 Then
                HadGreaterLength = True
                
'                '==========================
'                Debug.Print GoodCodesString
'                '==========================
                
                If Len(GoodCodesString) > 1 Then
                    GoodCodesString = Left(GoodCodesString, Len(GoodCodesString) - 1)
                End If
                
                ReDim Parameter(1) As Parameter
                Parameter(0) = GenerateInputParameter("@Discount", adDouble, 8, Discount)
                Parameter(1) = GenerateInputParameter("@GoodCodesString", adVarWChar, 4000, GoodCodesString)
                RunParametricStoredProcedure "Update_Good_Discount", Parameter
                FWProgressBar1.Value = 0
                FWProgressBar1.Max = GoodCodesLength
                Dim j As Long
                For j = 0 To GoodCodesLength - 1
                    FWProgressBar1.Value = FWProgressBar1.Value + 1
                Next j
                GoodCodesString = ""
            End If
        Next i
    End With
    
    If GoodCodesString <> "" Then
        If Len(GoodCodesString) > 1 Then
            GoodCodesString = Left(GoodCodesString, Len(GoodCodesString) - 1)
        End If
        
        ReDim Parameter(1) As Parameter
        Parameter(0) = GenerateInputParameter("@Discount", adDouble, 8, Discount)
        Parameter(1) = GenerateInputParameter("@GoodCodesString", adVarWChar, 4000, GoodCodesString)
        RunParametricStoredProcedure "Update_Good_Discount", Parameter
    End If
    FillvsGood
 '   DefaultSetting
    Me.MousePointer = vbDefault
    FWProgressBar1.Value = 0
    ShowDisMessage " €ÌÌ— œ—’œ  Œ›Ì› «‰Ã«„ ‘œ ", 2000
    
Exit Sub
Err_Handler:
    MsgBox err.Description
    Me.MousePointer = vbDefault
   
End Sub

Private Sub NewDutyBuyUpdate_Click()
    On Error GoTo Err_Handler
    
'    Dim buyDuty As Double
'    Dim strBuyDuty As String
'    strBuyDuty = Trim(txtDutyBuy.Text)
'
'    If IsNumeric(strBuyDuty) = False Then
'        ShowMessage "„ﬁœ«— Ê«—œ ‘œÂ »—«Ì ⁄Ê«—÷ Œ—Ìœ ﬂ«·« „⁄ »— ‰Ì” . ·ÿ›« Ìﬂ „ﬁœ«— „⁄ »— —« Ê«—œ ﬂ‰Ìœ", True, False, " «ÌÌœ", ""
'        Exit Sub
'    End If
'
'    buyDuty = Val(strBuyDuty)
'
'    If buyDuty > 100 Then
'        ShowMessage " «—ﬁ«„ „⁄ »— Ê«—œ ﬂ‰Ìœ ", True, False, "ﬁ»Ê·", ""
'        Exit Sub
'    End If


    Me.MousePointer = vbHourglass
    Dim GoodCodesString As String
    Dim GoodCodesLength As Long
    Dim HadGreaterLength As Boolean
    HadGreaterLength = False
    GoodCodesString = ""

    With vsGood
        For i = 1 To .Rows - 1
            GoodCodesString = GoodCodesString & .TextMatrix(i, 1) & STRING_DELIMITER
            GoodCodesLength = Len(GoodCodesString)
            If GoodCodesLength > 3800 Then
                HadGreaterLength = True
                
'                '==========================
'                Debug.Print GoodCodesString
'                '==========================
                
                If Len(GoodCodesString) > 1 Then
                    GoodCodesString = Left(GoodCodesString, Len(GoodCodesString) - 1)
                End If
                
                ReDim Parameter(1) As Parameter
                Parameter(0) = GenerateInputParameter("@DutyBuy", adBoolean, 1, ChkDutyBuy.Value)
                Parameter(1) = GenerateInputParameter("@GoodCodesString", adVarWChar, 4000, GoodCodesString)
                RunParametricStoredProcedure "Update_Good_DutyBuy", Parameter
                FWProgressBar1.Value = 0
                FWProgressBar1.Max = GoodCodesLength
                Dim j As Long
                For j = 0 To GoodCodesLength - 1
                    FWProgressBar1.Value = FWProgressBar1.Value + 1
                Next j
                GoodCodesString = ""
            End If
        Next i
    End With
    
    If GoodCodesString <> "" Then
        If Len(GoodCodesString) > 1 Then
            GoodCodesString = Left(GoodCodesString, Len(GoodCodesString) - 1)
        End If
        
        ReDim Parameter(1) As Parameter
        Parameter(0) = GenerateInputParameter("@DutyBuy", adBoolean, 1, ChkDutyBuy.Value)
        Parameter(1) = GenerateInputParameter("@GoodCodesString", adVarWChar, 4000, GoodCodesString)
        RunParametricStoredProcedure "Update_Good_DutyBuy", Parameter
    End If
    
    FillvsGood
'    DefaultSetting
    Me.MousePointer = vbDefault
    FWProgressBar1.Value = 0
    ShowDisMessage " €ÌÌ— œ—’œ ⁄Ê«—÷ Œ—Ìœ «‰Ã«„ ‘œ ", 2000

Exit Sub
Err_Handler:
    MsgBox err.Description
    Me.MousePointer = vbDefault

End Sub

Private Sub NewDutySaleUpdate_Click()
    On Error GoTo Err_Handler
    
'    Dim saleDuty As Double
'    Dim strSaleDuty As String
'    strSaleDuty = Trim(txtDutySale.Text)
'
'    If IsNumeric(strSaleDuty) = False Then
'        ShowMessage "„ﬁœ«— Ê«—œ ‘œÂ »—«Ì ⁄Ê«—÷ ›—Ê‘ ﬂ«·« „⁄ »— ‰Ì” . ·ÿ›« Ìﬂ „ﬁœ«— „⁄ »— —« Ê«—œ ﬂ‰Ìœ", True, False, " «ÌÌœ", ""
'        Exit Sub
'    End If
'
'    saleDuty = Val(strSaleDuty)
'
'    If saleDuty > 100 Then
'        ShowMessage " «—ﬁ«„ „⁄ »— Ê«—œ ﬂ‰Ìœ ", True, False, "ﬁ»Ê·", ""
'        Exit Sub
'    End If

    Me.MousePointer = vbHourglass
    Dim GoodCodesString As String
    Dim GoodCodesLength As Long
    Dim HadGreaterLength As Boolean
    HadGreaterLength = False
    GoodCodesString = ""

    With vsGood
        For i = 1 To .Rows - 1
            GoodCodesString = GoodCodesString & .TextMatrix(i, 1) & STRING_DELIMITER
            GoodCodesLength = Len(GoodCodesString)
            If GoodCodesLength > 3800 Then
                HadGreaterLength = True
                
'                '==========================
'                Debug.Print GoodCodesString
'                '==========================
                
                If Len(GoodCodesString) > 1 Then
                    GoodCodesString = Left(GoodCodesString, Len(GoodCodesString) - 1)
                End If
                
                ReDim Parameter(1) As Parameter
                Parameter(0) = GenerateInputParameter("@DutySale", adBoolean, 1, ChkDutySale.Value)
                Parameter(1) = GenerateInputParameter("@GoodCodesString", adVarWChar, 4000, GoodCodesString)
                RunParametricStoredProcedure "Update_Good_DutySale", Parameter
                FWProgressBar1.Value = 0
                FWProgressBar1.Max = GoodCodesLength
                Dim j As Long
                For j = 0 To GoodCodesLength - 1
                    FWProgressBar1.Value = FWProgressBar1.Value + 1
                Next j
                GoodCodesString = ""
            End If
        Next i
    End With
    
    If GoodCodesString <> "" Then
        If Len(GoodCodesString) > 1 Then
            GoodCodesString = Left(GoodCodesString, Len(GoodCodesString) - 1)
        End If
        
        ReDim Parameter(1) As Parameter
        Parameter(0) = GenerateInputParameter("@DutySale", adBoolean, 1, ChkDutySale.Value)
        Parameter(1) = GenerateInputParameter("@GoodCodesString", adVarWChar, 4000, GoodCodesString)
        RunParametricStoredProcedure "Update_Good_DutySale", Parameter
    End If
    FillvsGood
  '  DefaultSetting
    Me.MousePointer = vbDefault
    FWProgressBar1.Value = 0
    ShowDisMessage " €ÌÌ— œ—’œ ⁄Ê«—÷ ›—Ê‘ «‰Ã«„ ‘œ ", 2000

Exit Sub
Err_Handler:
    MsgBox err.Description
    Me.MousePointer = vbDefault

End Sub

Private Sub NewTaxBuyUpdate_Click()
    On Error GoTo Err_Handler
    
'    Dim buyTax As Double
'    Dim strBuyTax As String
'    strBuyTax = Trim(txtTaxBuy.Text)
    
'    If IsNumeric(strBuyTax) = False Then
'        ShowMessage "„ﬁœ«— Ê«—œ ‘œÂ »—«Ì „«·Ì«  Œ—Ìœ ﬂ«·« „⁄ »— ‰Ì” . ·ÿ›« Ìﬂ „ﬁœ«— „⁄ »— —« Ê«—œ ﬂ‰Ìœ", True, False, " «ÌÌœ", ""
'        Exit Sub
'    End If
    
'    buyTax = Val(strBuyTax)
'
'    If buyTax > 100 Then
'        ShowMessage " «—ﬁ«„ „⁄ »— Ê«—œ ﬂ‰Ìœ ", True, False, "ﬁ»Ê·", ""
'        Exit Sub
'    End If
    
    Me.MousePointer = vbHourglass
    Dim GoodCodesString As String
    Dim GoodCodesLength As Long
    Dim HadGreaterLength As Boolean
    HadGreaterLength = False
    GoodCodesString = ""

    With vsGood
        For i = 1 To .Rows - 1
            GoodCodesString = GoodCodesString & .TextMatrix(i, 1) & STRING_DELIMITER
            GoodCodesLength = Len(GoodCodesString)
            If GoodCodesLength > 3800 Then
                HadGreaterLength = True
                
'                '==========================
'                Debug.Print GoodCodesString
'                '==========================
                
                If Len(GoodCodesString) > 1 Then
                    GoodCodesString = Left(GoodCodesString, Len(GoodCodesString) - 1)
                End If
                
                ReDim Parameter(1) As Parameter
                Parameter(0) = GenerateInputParameter("@TaxBuy", adBoolean, 1, ChkTaxBuy.Value)
                Parameter(1) = GenerateInputParameter("@GoodCodesString", adVarWChar, 4000, GoodCodesString)
                RunParametricStoredProcedure "Update_Good_TaxBuy", Parameter
                FWProgressBar1.Value = 0
                FWProgressBar1.Max = GoodCodesLength
                Dim j As Long
                For j = 0 To GoodCodesLength - 1
                    FWProgressBar1.Value = FWProgressBar1.Value + 1
                Next j
                GoodCodesString = ""
            End If
        Next i
    End With
    
    If GoodCodesString <> "" Then
        If Len(GoodCodesString) > 1 Then
            GoodCodesString = Left(GoodCodesString, Len(GoodCodesString) - 1)
        End If
        
        ReDim Parameter(1) As Parameter
        Parameter(0) = GenerateInputParameter("@TaxBuy", adBoolean, 1, ChkTaxBuy.Value)
        Parameter(1) = GenerateInputParameter("@GoodCodesString", adVarWChar, 4000, GoodCodesString)
        RunParametricStoredProcedure "Update_Good_TaxBuy", Parameter
    End If
    FillvsGood
 '   DefaultSetting
    Me.MousePointer = vbDefault
    FWProgressBar1.Value = 0
    ShowDisMessage " €ÌÌ— œ—’œ „«·Ì«  Œ—Ìœ «‰Ã«„ ‘œ ", 2000

Exit Sub
Err_Handler:
    MsgBox err.Description
    Me.MousePointer = vbDefault

End Sub

Private Sub NewTaxSaleUpdate_Click()
    On Error GoTo Err_Handler
    
'    Dim saleTax As Double
'    Dim strSaleTax As String
'    strSaleTax = Trim(txtTaxSale.Text)
'
'    If IsNumeric(strSaleTax) = False Then
'        ShowMessage "„ﬁœ«— Ê«—œ ‘œÂ »—«Ì „«·Ì«  ›—Ê‘ ﬂ«·« „⁄ »— ‰Ì” . ·ÿ›« Ìﬂ „ﬁœ«— „⁄ »— —« Ê«—œ ﬂ‰Ìœ", True, False, " «ÌÌœ", ""
'        Exit Sub
'    End If
'
'    saleTax = Val(strSaleTax)
'
'    If saleTax > 100 Then
'        ShowMessage " «—ﬁ«„ „⁄ »— Ê«—œ ﬂ‰Ìœ ", True, False, "ﬁ»Ê·", ""
'        Exit Sub
'    End If
    
    Me.MousePointer = vbHourglass
    Dim GoodCodesString As String
    Dim GoodCodesLength As Long
    Dim HadGreaterLength As Boolean
    HadGreaterLength = False
    GoodCodesString = ""

    With vsGood
        For i = 1 To .Rows - 1
            GoodCodesString = GoodCodesString & .TextMatrix(i, 1) & STRING_DELIMITER
            GoodCodesLength = Len(GoodCodesString)
            If GoodCodesLength > 3800 Then
                HadGreaterLength = True
                
'                '==========================
'                Debug.Print GoodCodesString
'                '==========================
                
                If Len(GoodCodesString) > 1 Then
                    GoodCodesString = Left(GoodCodesString, Len(GoodCodesString) - 1)
                End If
                
                ReDim Parameter(1) As Parameter
                Parameter(0) = GenerateInputParameter("@TaxSale", adBoolean, 1, ChkTaxSale.Value)
                Parameter(1) = GenerateInputParameter("@GoodCodesString", adVarWChar, 4000, GoodCodesString)
                RunParametricStoredProcedure "Update_Good_TaxSale", Parameter
                FWProgressBar1.Value = 0
                FWProgressBar1.Max = GoodCodesLength
                Dim j As Long
                For j = 0 To GoodCodesLength - 1
                    FWProgressBar1.Value = FWProgressBar1.Value + 1
                Next j
                GoodCodesString = ""
            End If
        Next i
    End With
    
    If GoodCodesString <> "" Then
        If Len(GoodCodesString) > 1 Then
            GoodCodesString = Left(GoodCodesString, Len(GoodCodesString) - 1)
        End If
        
        ReDim Parameter(1) As Parameter
        Parameter(0) = GenerateInputParameter("@TaxSale", adBoolean, 1, ChkTaxSale.Value)
        Parameter(1) = GenerateInputParameter("@GoodCodesString", adVarWChar, 4000, GoodCodesString)
        RunParametricStoredProcedure "Update_Good_TaxSale", Parameter
    End If
    FillvsGood
'    DefaultSetting
    Me.MousePointer = vbDefault
    FWProgressBar1.Value = 0
    ShowDisMessage " €ÌÌ— œ—’œ „«·Ì«  ›—Ê‘ «‰Ã«„ ‘œ ", 2000

Exit Sub
Err_Handler:
    MsgBox err.Description

    Me.MousePointer = vbDefault

End Sub

Private Sub ResizeKit1_ExitResize(ByVal XScale As Double, ByVal YScale As Double)


    If formloadFlag = True Then
        SaveSetting strMainKey, Me.Name, "Height", Me.Height
        SaveSetting strMainKey, Me.Name, "Width", Me.Width
    End If

End Sub

Private Sub txtBarcode_Change()
    If Len(txtBarcode.Text) > 2 Then
    If Asc(Mid(txtBarcode.Text, Len(txtBarcode.Text) - 1, 1)) = 13 Then
        txtBarcode.Text = Left(txtBarcode.Text, Len(txtBarcode.Text) - 2)
    End If
    End If
    i = vsGood.FindRow(Trim(txtBarcode.Text), 1, 6, True, True)
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

Private Sub txtDiscount_Change()
  If Len(txtDiscount.Text) > 2 Then
     txtDiscount.Text = Val(Left(txtDiscount.Text, 2))
  Else
     txtDiscount.Text = Val(txtDiscount.Text)
  End If
  
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

Private Sub vsGood_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    
    For i = 0 To vsGood.Cols - 1
        SaveSetting strMainKey, "frmGoodDiscount_vsGood", "Col" & i, vsGood.ColWidth(i)
    Next

End Sub

Private Sub vsGood_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    tmpTextMatrix = vsGood.TextMatrix(Row, Col)
End Sub


Private Sub vsGood_Click()
'    With vsGood
'        If (.TextMatrix(.Row, 0) = "*" Or MyFormAddEditMode = EnumAddEditMode.EditMode) Then
'            If .Col = 6 Or .Col = 7 Or .Col = 8 Or .Col = 9 Or .Col = 10 Then
'               .Select .Row, .Col
'               .EditCell
'            End If
'        End If
'
'    End With

End Sub

Private Sub vsGood_KeyDown(KeyCode As Integer, Shift As Integer)
    With vsGood
        If (.TextMatrix(.Row, 0) = "*" Or MyFormAddEditMode = EnumAddEditMode.EditMode) Then
            If .Col = 6 Or .Col = 7 Or .Col = 8 Or .Col = 9 Or .Col = 10 Then
               .Select .Row, .Col
               .EditCell
            End If
        End If
    End With
    
End Sub


Private Sub vsGood_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    With vsGood
        If KeyAscii = 39 Then KeyAscii = 0
        
        If Col = 6 And (IsNumeric(Chr(KeyAscii)) = False And KeyAscii <> 8 And KeyAscii <> 46) Then
            
            KeyAscii = 0
            
        ElseIf MyFormAddEditMode = EditMode Then
            
            If Row > 0 And InStr(1, .TextMatrix(Row, 0), "*") = 0 Then
                .TextMatrix(Row, 0) = .TextMatrix(Row, 0) & "*"
            End If
            
        End If
        
    End With
    
End Sub


Private Sub vsGood_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    With vsGood
        If (.TextMatrix(.Row, 0) = "*" Or MyFormAddEditMode = EnumAddEditMode.EditMode) Then
            If .Col = 6 Or .Col = 7 Or .Col = 8 Or .Col = 9 Or .Col = 10 Then
               .Select .Row, .Col
               .EditCell
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

Private Function UpdateChanges() As Boolean
    On Error GoTo Err_Handler
    Dim blnResult As Boolean
    blnResult = False
    
    Me.MousePointer = vbHourglass
    Dim GoodCodesString As String
    Dim GoodCodesLength As Long
    Dim HadGreaterLength As Boolean
    
    Dim Discount As Double
    Dim buyDuty As Double
    Dim saleDuty As Double
    Dim buyTax As Double
    Dim saleTax As Double
    Dim changeCount As Long
    
    With vsGood
        For i = 1 To .Rows - 1
            If InStr(.TextMatrix(i, 0), "*") > 0 Then changeCount = changeCount + 1
            
        Next i
    End With
    HadGreaterLength = False
    GoodCodesString = ""
    
    FWProgressBar1.Value = 0
    FWProgressBar1.Max = changeCount
    
    With vsGood
        For i = 1 To .Rows - 1
            If InStr(.TextMatrix(i, 0), "*") > 0 Then 'edited records
                GoodCodesString = GoodCodesString & .TextMatrix(i, 1) '& STRING_DELIMITER
                
                Discount = Val(.TextMatrix(i, 6))
                buyDuty = IIf(Val(.TextMatrix(i, 7)) = -1, 1, 0)
                saleDuty = IIf(Val(.TextMatrix(i, 8)) = -1, 1, 0)
                buyTax = IIf(Val(.TextMatrix(i, 9)) = -1, 1, 0)
                saleTax = IIf(Val(.TextMatrix(i, 10)) = -1, 1, 0)
                '==============================
                
                ReDim Parameter(1) As Parameter
                Parameter(0) = GenerateInputParameter("@Discount", adDouble, 8, Discount)
                Parameter(1) = GenerateInputParameter("@GoodCodesString", adVarWChar, 4000, GoodCodesString)
                DoEvents
                RunParametricStoredProcedure "Update_Good_Discount", Parameter
                '==============================
                
                ReDim Parameter(1) As Parameter
                Parameter(0) = GenerateInputParameter("@DutyBuy", adBoolean, 1, buyDuty)
                Parameter(1) = GenerateInputParameter("@GoodCodesString", adVarWChar, 4000, GoodCodesString)
                DoEvents
                RunParametricStoredProcedure "Update_Good_DutyBuy", Parameter
                '==============================
                
                ReDim Parameter(1) As Parameter
                Parameter(0) = GenerateInputParameter("@DutySale", adBoolean, 1, saleDuty)
                Parameter(1) = GenerateInputParameter("@GoodCodesString", adVarWChar, 4000, GoodCodesString)
                DoEvents
                RunParametricStoredProcedure "Update_Good_DutySale", Parameter
                '==============================
                
                ReDim Parameter(1) As Parameter
                Parameter(0) = GenerateInputParameter("@TaxBuy", adBoolean, 1, buyTax)
                Parameter(1) = GenerateInputParameter("@GoodCodesString", adVarWChar, 4000, GoodCodesString)
                DoEvents
                RunParametricStoredProcedure "Update_Good_TaxBuy", Parameter
                '==============================

                ReDim Parameter(1) As Parameter
                Parameter(0) = GenerateInputParameter("@TaxSale", adBoolean, 1, saleTax)
                Parameter(1) = GenerateInputParameter("@GoodCodesString", adVarWChar, 4000, GoodCodesString)
                DoEvents
                RunParametricStoredProcedure "Update_Good_TaxSale", Parameter
                
                FWProgressBar1.Value = FWProgressBar1.Value + 1

                GoodCodesString = ""
            End If
        Next i
    End With
    
    blnResult = True
    UpdateChanges = blnResult
    Me.MousePointer = vbDefault
Exit Function
Err_Handler:
    LogSaveNew "frmGoodDiscount => ", err.Description, err.Number, err.Source, "UpdateChanges"
    ShowErrorMessage
    UpdateChanges = False
    Me.MousePointer = vbDefault
End Function

Private Function ValidateUIData() As Boolean
    Dim blnResult As Boolean
    blnResult = True
    
    With vsGood
        For i = 1 To .Rows - 1
            .Row = i
            If InStr(.TextMatrix(i, 0), "*") > 0 Then 'new or edited records
                If Val(.TextMatrix(i, 6)) > 100 Then     '
                        Select Case clsStation.Language
                            Case 0
                                frmMsg.fwlblMsg.Caption = " Œ›Ì› »“—ê — «“ 100 œ—’œ ﬁ«»· ﬁ»Ê· ‰Ì” "
                                frmMsg.fwBtn(0).Caption = "ﬁ»Ê·"

                            Case 1
                                frmMsg.fwlblMsg.Caption = "Error In Discount Greater Than 100"
                                frmMsg.fwBtn(0).Caption = "Ok"
                                frmMsg.fwlblMsg.Alignment = vbLeftJustify
                        End Select

                        frmMsg.fwBtn(0).ButtonType = flwButtonOk
                        frmMsg.fwBtn(1).Visible = False
                        frmMsg.Show vbModal
                        
                        blnResult = False
                        Exit For
                End If
            End If
        Next i
    End With
    
    ValidateUIData = blnResult
End Function


