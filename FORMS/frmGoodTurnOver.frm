VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{7AEDC602-D94C-11D1-BB7A-00E0290EA3C9}#1.0#0"; "ResizeKit.ocx"
Object = "{75D4F767-8785-11D3-93AD-0000832EF44D}#2.14#0"; "FAST2001.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmGoodTurnOver 
   ClientHeight    =   9450
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13695
   Icon            =   "frmGoodTurnOver.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9450
   ScaleWidth      =   13695
   Begin VB.Frame Frame4 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   885
      Left            =   60
      TabIndex        =   19
      Top             =   495
      Width           =   5700
      Begin VB.CommandButton cmdSetFirstPrice 
         Caption         =   "»Â —Ê“ —”«‰Ì ›Ì «Ê·ÌÂ »« ﬁÌ„  Œ—Ìœ ﬂ«·«"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   855
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Top             =   165
         Width           =   4005
      End
   End
   Begin RESIZEKITLibCtl.ResizeKit ResizeKit1 
      Height          =   480
      Left            =   240
      OleObjectBlob   =   "frmGoodTurnOver.frx":A4C2
      TabIndex        =   17
      Top             =   0
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
      Left            =   6240
      MultiLine       =   -1  'True
      RightToLeft     =   -1  'True
      TabIndex        =   13
      Top             =   2640
      Width           =   2385
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
      Left            =   10800
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
      Left            =   10800
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   1800
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
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   9480
      Top             =   480
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
      Height          =   1575
      Left            =   6240
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   960
      Width           =   3495
      Begin FLWCtrls.FWCoolButton fwBtnGoodFind 
         Height          =   930
         Left            =   120
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   360
         Width           =   3090
         _ExtentX        =   5450
         _ExtentY        =   1640
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
         Picture         =   "frmGoodTurnOver.frx":A548
         PictureAlign    =   4
         Caption         =   "ò«·«"
         MaskColor       =   -2147483633
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1815
      Left            =   0
      TabIndex        =   1
      Top             =   1350
      Width           =   5760
      Begin VB.CommandButton StoreDataUpdate 
         BackColor       =   &H000000C0&
         Caption         =   "„Õ«”»Â ê—œ‘ ﬂ«·« œ— «‰»«—"
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   840
         Left            =   435
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   840
         Width           =   2295
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
         Left            =   495
         RightToLeft     =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   225
         Width           =   1455
      End
      Begin MSMask.MaskEdBox txtDateTo 
         Height          =   465
         Left            =   3225
         TabIndex        =   2
         Top             =   1110
         Width           =   1440
         _ExtentX        =   2540
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
         Left            =   3255
         TabIndex        =   3
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
         Left            =   2190
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Top             =   255
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
         Top             =   1110
         Width           =   975
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
         Left            =   4500
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   390
         Width           =   915
      End
   End
   Begin VSFlex7LCtl.VSFlexGrid vsGood 
      Height          =   5985
      Left            =   0
      TabIndex        =   0
      Top             =   3240
      Width           =   13665
      _cx             =   24104
      _cy             =   10557
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
      FormatString    =   $"frmGoodTurnOver.frx":A862
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
      AllowUserFreezing=   3
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
   End
   Begin FLWCtrls.FWLabel3D fwlblMode 
      Height          =   495
      Left            =   12000
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
      Left            =   8760
      RightToLeft     =   -1  'True
      TabIndex        =   14
      Top             =   2640
      Width           =   825
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ò«—œò” «‰»«— - ê—œ‘ ò«·«Â« œ— «‰»«—"
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
      Left            =   4320
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   0
      Width           =   5175
   End
End
Attribute VB_Name = "frmGoodTurnOver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private MyFormAddEditMode As EnumAddEditMode
Dim tmpTextMatrix As String
Dim Parameter() As Parameter
Dim clsDate As New clsDate
'==================================================
Const GRID_COLUMNS_COUNT As Integer = 18

Const IndexColRow As Integer = 0
Const IndexColCode As Integer = 1
Const IndexColName As Integer = 2
Const IndexColSanad As Integer = 3
Const IndexColDate  As Integer = 4
Const IndexColFromStore  As Integer = 5
Const IndexColDescription As Integer = 6
Const IndexColToStore  As Integer = 7
Const IndexColInput  As Integer = 8
Const IndexColOutput  As Integer = 9
Const IndexColMojodi  As Integer = 10
Const IndexColFee  As Integer = 11
Const IndexColInputTotal  As Integer = 12
Const IndexColOutputTotal As Integer = 13
Const IndexColMojodiTotal As Integer = 14
Const IndexColAverageFee  As Integer = 15
Const IndexColBlank  As Integer = 16
'==================================================

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
    FillvsGood ("")
End Sub

Public Sub FillvsGood(DateSearch As String) 'it fills the grid using vw_Good
    On Error GoTo Err_Handler
    
    vsGood.Rows = 1
    
    If fwBtnGoodFind.Tag = "" Then Exit Sub
    MyFormAddEditMode = EnumAddEditMode.ViewMode 'VIEW Mode
    SetFirstToolBar
    
    Dim InventoryNo As Integer
    If cmbInventory.ListIndex = -1 Then
        InventoryNo = 0
    Else
        InventoryNo = cmbInventory.ItemData(cmbInventory.ListIndex)
    End If
    
    ReDim Parameter(9) As Parameter
    Parameter(0) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
    Parameter(1) = GenerateInputParameter("@SystemDate", adVarWChar, 50, Mid(clsDate.shamsi(Date), 3, 8))
    Parameter(2) = GenerateInputParameter("@SystemDay", adVarWChar, 50, clsDate.Find_DayOfWeek(Weekday(Date, vbSaturday)))
    Parameter(3) = GenerateInputParameter("@SystemTime", adVarWChar, 50, Mid(Str(time), 1, 5))
    Parameter(4) = GenerateInputParameter("@DateBefore", adVarWChar, 50, CStr(IIf(Trim(txtDateFrom.ClipText) = "", "", Trim(txtDateFrom.Text))))
    Parameter(5) = GenerateInputParameter("@DateAfter", adVarWChar, 50, CStr(IIf(Trim(txtDateTo.ClipText) = "", "", Trim(txtDateTo.Text))))
    Parameter(6) = GenerateInputParameter("@GoodCode", adInteger, 4, Val(fwBtnGoodFind.Tag))
    Parameter(7) = GenerateInputParameter("@InventoryNo", adInteger, 4, InventoryNo)
    Parameter(8) = GenerateInputParameter("@Branch", adInteger, 4, cmbBranch.ItemData(cmbBranch.ListIndex))
    Parameter(9) = GenerateInputParameter("@AccountYear", adSmallInt, 2, Val(cmbSalMali.Text))
'    Parameter(10) = GenerateInputParameter("@DateSearch", adVarWChar, 20, DateSearch)
    
    Set Rst = RunParametricStoredProcedure2Rec("GetInventoryGood_Mojodi_New", Parameter)
    
    ShowDisMessage "„Õ«”»Â «‰Ã«„ ‘œ ", 1400
    
    Dim Mojodi As Double
    Mojodi = 0#
    
    If Not (Rst.EOF = True And Rst.BOF = True) Then
        With vsGood
            Dim amount As Double
            Dim Status As Long
            Dim Fee As Double
            Dim Import As Double
            Dim Export As Double
            Import = 0#
            Export = 0#
            i = 1
            Dim Store As String
            Store = ""

            
            If InStr(1, Store, Rst!fromStore, 1) = 0 Then
                Dim FirstMojodi As Double
                Dim FirstPrice As Double
                
                FirstMojodi = IIf(IsNull(Rst!FirstMojodi), 0, Rst!FirstMojodi)
                FirstPrice = IIf(IsNull(Rst!FirstPrice), 0, Rst!FirstPrice)
                
                .Rows = .Rows + 1
                .TextMatrix(i, IndexColRow) = i
                .TextMatrix(i, IndexColCode) = Rst!GoodCode
                .TextMatrix(i, IndexColName) = Left(Rst!Name, 25)
                .TextMatrix(i, IndexColSanad) = ""
                .TextMatrix(i, IndexColDate) = ""
                .TextMatrix(i, IndexColFromStore) = Rst!fromStore
                .TextMatrix(i, IndexColDescription) = "„ÊÃÊœÌ «Ê·ÌÂ"
                .TextMatrix(i, IndexColToStore) = ""
                .TextMatrix(i, IndexColInput) = FirstMojodi
                .TextMatrix(i, IndexColOutput) = 0
                .TextMatrix(i, IndexColMojodi) = FirstMojodi
                .TextMatrix(i, IndexColFee) = Price
                .TextMatrix(i, IndexColInputTotal) = FirstMojodi * FirstPrice
                .TextMatrix(i, IndexColOutputTotal) = 0
                .TextMatrix(i, IndexColMojodiTotal) = FirstMojodi * FirstPrice
                .TextMatrix(i, IndexColAverageFee) = 0
                
                Mojodi = Mojodi + FirstMojodi
                Import = Import + FirstMojodi
                i = i + 1
                Store = Store & " " & Rst!fromStore
            End If
                
            While Rst.EOF = False
                amount = IIf(IsNull(Rst!amount), 0, Rst!amount)
'                Price = IIf(IsNull(Rst!FeeUnit), 0, Rst!FeeUnit)
                Status = Rst!Status
                Fee = Rst!FeeUnit
                
                .Rows = .Rows + 1
                .TextMatrix(i, IndexColRow) = i
                .TextMatrix(i, IndexColCode) = Rst!GoodCode
                .TextMatrix(i, IndexColName) = Left(Rst!Name, 25)
                .TextMatrix(i, IndexColSanad) = Rst!No
                .TextMatrix(i, IndexColDate) = Rst!Date
                .TextMatrix(i, IndexColFromStore) = Rst!fromStore
                .TextMatrix(i, IndexColDescription) = Rst!NvcDescription
                .TextMatrix(i, IndexColToStore) = Rst!DestDescription
                
                If amount > 0 Then
                    .TextMatrix(i, IndexColInput) = amount
                    .TextMatrix(i, IndexColOutput) = 0
                    .TextMatrix(i, IndexColInputTotal) = amount * Fee
                    .TextMatrix(i, IndexColOutputTotal) = 0
                    Import = Import + amount
                Else
                    .TextMatrix(i, IndexColInput) = 0
                    .TextMatrix(i, IndexColOutput) = -amount
                    .TextMatrix(i, IndexColOutputTotal) = -amount * Fee
                    .TextMatrix(i, IndexColInputTotal) = 0
                    Export = Export + (-amount)
                End If
                
                Mojodi = Mojodi + amount
                If Mojodi <> Int(Mojodi) Then
                    Mojodi = Format(Mojodi, "##.000")
                End If
                .TextMatrix(i, IndexColMojodi) = Mojodi
                .TextMatrix(i, IndexColFee) = Fee
'                .TextMatrix(i, IndexColInputTotal) = Amount * Fee
             '   .TextMatrix(i, IndexColMojodiTotal) = (Mojodi * Fee) '+ Val(.TextMatrix(i - 1, IndexColMojodiTotal))
                 If amount > 0 Then
                    .TextMatrix(i, IndexColMojodiTotal) = Val(.TextMatrix(i, IndexColInputTotal)) + Val(.TextMatrix(i - 1, IndexColMojodiTotal))
                 Else
                    .TextMatrix(i, IndexColMojodiTotal) = (-1 * Val(.TextMatrix(i, IndexColOutputTotal))) + Val(.TextMatrix(i - 1, IndexColMojodiTotal))
                End If
                If Mojodi <> 0 Then
                    .TextMatrix(i, IndexColAverageFee) = Format(Rst!BuyPrice, "##")
                Else
                    .TextMatrix(i, IndexColMojodiTotal) = 0
                    .TextMatrix(i, IndexColAverageFee) = 0
                End If
                
                i = i + 1
                Rst.MoveNext
            Wend

            Rst.Close: Set Rst = Nothing
            
            .Rows = .Rows + 1
            .TextMatrix(i, IndexColRow) = i
            .TextMatrix(i, IndexColCode) = ""
            .TextMatrix(i, IndexColName) = ""
            .TextMatrix(i, IndexColSanad) = ""
            .TextMatrix(i, IndexColDate) = ""
            .TextMatrix(i, IndexColFromStore) = ""
            .TextMatrix(i, IndexColDescription) = "Ã„⁄"
            .TextMatrix(i, IndexColToStore) = ""
            .TextMatrix(i, IndexColInput) = Import
            .TextMatrix(i, IndexColOutput) = Export
            .TextMatrix(i, IndexColMojodi) = ""
            .TextMatrix(i, IndexColFee) = ""
            .TextMatrix(i, IndexColInputTotal) = ""
            .TextMatrix(i, IndexColOutputTotal) = ""
            .TextMatrix(i, IndexColMojodiTotal) = ""
            .TextMatrix(i, IndexColAverageFee) = ""
            .TextMatrix(i, IndexColBlank) = ""
                
            .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter
            .Cell(flexcpAlignment, 1, IndexColName, .Rows - 1, IndexColName) = flexAlignRightCenter
            .AutoSizeMode = flexAutoSizeColWidth
            .AutoSize 0, .Cols - 1
            
            ' for disable sorting
            .ExplorerBar = flexExNone  ' for disable move and sort header column
               '  sort must be disable because we calculate sum of column (TotalMojodi)

        End With
    End If

Exit Sub
Err_Handler:
    LogSaveNew "frmGoodTurnOver =>", err.Description, err.Number, err.Source, "FillVsGood"
    ShowErrorMessage
    Resume Next
End Sub

Public Sub Cancel()
    MyFormAddEditMode = EnumAddEditMode.ViewMode
    SetFirstToolBar
    fwBtnGoodFind.Tag = ""
    fwBtnGoodFind.Caption = ""
    FillvsGood ("")
End Sub

Private Sub cmbBranch_Click()
    FillInventory
End Sub

Private Sub cmbInventory_Click()
    If cmbBranch.ListIndex = -1 Then Exit Sub
    fwBtnGoodFind.Tag = ""
    fwBtnGoodFind.Caption = ""
    FillvsGood ("")
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
        FillvsGood ("")
    End If
End Sub

Private Sub cmbSalMali_Click()
    cmbSalMali_Change
End Sub

Private Sub FillSalMali()
    Dim L_Rst As New ADODB.Recordset
    
    cmbSalMali.Clear
    Set L_Rst = RunStoredProcedure2RecordSet("Get_All_tAccountYears")
    Do While L_Rst.EOF = False
        cmbSalMali.AddItem L_Rst!AccountYear
        cmbSalMali.ItemData(cmbSalMali.NewIndex) = L_Rst!AccountYear
        L_Rst.MoveNext
    Loop
    Dim i As Integer
    For i = 0 To cmbSalMali.ListCount - 1
        If AccountYear = cmbSalMali.ItemData(i) Then
            cmbSalMali.ListIndex = i
            Exit For
        End If
    Next
    L_Rst.Close: Set L_Rst = Nothing
End Sub

Private Sub cmdSetFirstPrice_Click()
    If cmbSalMali.Text = "" Or cmbInventory.ListIndex = -1 Then Exit Sub
    On Error GoTo ErrHandler
    
    Me.MousePointer = vbHourglass
    
    ReDim Parameter(2) As Parameter
    Parameter(0) = GenerateInputParameter("@AccountYear", adInteger, 4, Val(cmbSalMali.Text))
    Parameter(1) = GenerateInputParameter("@Flag", adInteger, 4, 1)
    Parameter(2) = GenerateInputParameter("@InventoryNO", adInteger, 4, cmbInventory.ItemData(cmbInventory.ListIndex))
    RunParametricStoredProcedure "Update_FirstPriceByBuyPrice", Parameter
     
    frmDisMsg.lblMessage.Caption = " ›Ì «Ê·ÌÂ Â„Â ﬂ«·« »Â —Ê“ —”«‰Ì ‘œ‰œ"
    frmDisMsg.Timer1.Enabled = True
    frmDisMsg.Show
    Me.MousePointer = vbDefault
    
        
Exit Sub
ErrHandler:
    LogSaveNew "frmGoodTurnOver=>", err.Description, err.Number, err.Source, "CmdSetFirstPrice_Click"
    ShowErrorMessage
    Me.MousePointer = vbDefault

End Sub

Private Sub Form_Activate()
''''    FWLed1.Value = CInt(AccountYear)
''''    FWLed1.BackColor = Me.BackColor
''''    FWLed1.ColorOff = Me.BackColor
    
    VarActForm = Me.Name
    
    Frame1.BackColor = Me.BackColor
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
    If ClsFormAccess.frmGoodTurnOver = False Then
        Unload Me
        Exit Sub
    End If
    
    If intVersion = Normal Or intVersion = Min Then
        ShowDisMessage "‰„«Ì‘ ﬂ«—œﬂ” ﬂ«·«Â« œ— ‰”ŒÂ Â«Ì ÅÌ‘—› Â Ê »«·« — ÊÃÊœ œ«—œ", 1500
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

    ChangeLanguage
    
    txtDateFrom.Text = Mid(AccountYear, 3, 2) & "/01" & "/01"
    txtDateTo.Text = Mid(clsDate.shamsi(Date), 3)
    
End Sub


Private Sub Form_Unload(Cancel As Integer)
    
    VarActForm = ""

    Dim i As Integer
    
    AllButton vbOff, True
    Unload frmFindGoods

    SaveSetting strMainKey, Me.Name, "Left", Me.Left
    SaveSetting strMainKey, Me.Name, "Top", Me.Top
    If PosConnection.State = adStateOpen Then PosConnection.Close: Set PosConnection = Nothing
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call PresetScreenSaver
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call PresetScreenSaver
End Sub

Public Sub ChangeLanguage()
    Dim Obj As Object

    Select Case clsStation.Language    ' LCase(mdifrm.Toolbar1.Buttons(25).Key)
        Case English
            mdifrm.Caption = clsArya.LatinCompany
            Me.RightToLeft = False
            
            For Each Obj In Me
                On Error Resume Next
                    Obj.RightToLeft = False
                On Error GoTo 0
            Next Obj
        
        Case Farsi
            mdifrm.Caption = clsArya.Company
            Me.RightToLeft = True
            
            For Each Obj In Me
                On Error Resume Next
                    Obj.RightToLeft = True
                On Error GoTo 0
            Next Obj
    End Select
    
    With vsGood
        .Cols = GRID_COLUMNS_COUNT
        
        Select Case clsStation.Language
            Case Farsi
                .TextMatrix(0, IndexColRow) = "—œÌ›"
                .TextMatrix(0, IndexColCode) = "òœ"
                .TextMatrix(0, IndexColName) = "‰«„ ò«·«"
                .TextMatrix(0, IndexColSanad) = "”‰œ"
                .TextMatrix(0, IndexColDate) = " «—ÌŒ"
                .TextMatrix(0, IndexColFromStore) = " «‰»«— „»œ√ "
                .TextMatrix(0, IndexColDescription) = " ‘—Õ "
                .TextMatrix(0, IndexColToStore) = " «‰»«— „ﬁ’œ "
                .TextMatrix(0, IndexColInput) = "Ê«—œÂ "
                .TextMatrix(0, IndexColOutput) = " ’«œ—Â "
                .TextMatrix(0, IndexColMojodi) = " „ÊÃÊœÌ "
                .TextMatrix(0, IndexColFee) = " ›Ì Œ—Ìœ "
                .TextMatrix(0, IndexColInputTotal) = " «—“‘ Ê«—œÂ "
                .TextMatrix(0, IndexColOutputTotal) = " «—“‘ ’«œ—Â "
                .TextMatrix(0, IndexColMojodiTotal) = " «—“‘ „ÊÃÊœÌ "
                .TextMatrix(0, IndexColAverageFee) = " ›Ì „Ì«‰êÌ‰ Œ—Ìœ "
                .TextMatrix(i, IndexColBlank) = ""
                
            Case English
                .TextMatrix(0, IndexColRow) = "Row"
                .TextMatrix(0, IndexColCode) = "Code"
                .TextMatrix(0, IndexColName) = "Name"
                .TextMatrix(0, IndexColSanad) = "Sanad"
                .TextMatrix(0, IndexColDate) = "Date"
                .TextMatrix(0, IndexColFromStore) = " From Store "
                .TextMatrix(0, IndexColDescription) = " Description "
                .TextMatrix(0, IndexColToStore) = " To Store "
                .TextMatrix(0, IndexColInput) = "Input "
                .TextMatrix(0, IndexColOutput) = " Output "
                .TextMatrix(0, IndexColMojodi) = " Stock "
                .TextMatrix(0, IndexColFee) = " Fee "
                .TextMatrix(0, IndexColInputTotal) = " Input Total "
                .TextMatrix(0, IndexColOutputTotal) = " Output Total "
                .TextMatrix(0, IndexColMojodiTotal) = " Stock Total "
                .TextMatrix(0, IndexColAverageFee) = " Average Fee "
                .TextMatrix(i, IndexColBlank) = ""
                
                .ColDataType(IndexColSanad) = flexDTLong8
                .ColDataType(IndexColDate) = flexDTStringW
                .ColDataType(IndexColDescription) = flexDTStringW
                .ColDataType(IndexColInput) = flexDTDouble
                .ColDataType(IndexColOutput) = flexDTDouble
                .ColDataType(IndexColMojodi) = flexDTDouble
                .ColDataType(IndexColFee) = flexDTDouble
                .ColDataType(IndexColInputTotal) = flexDTDouble
                .ColDataType(IndexColOutputTotal) = flexDTDouble
                .ColDataType(IndexColMojodiTotal) = flexDTDouble
                .ColDataType(IndexColAverageFee) = flexDTDouble
       End Select
       .ColSort(-1) = flexSortNone
       .Sort = flexSortNone
       
        .ColAlignment(-1) = flexAlignCenterCenter
        .FocusRect = flexFocusHeavy
        '.ColHidden(1) = True
'        .ColHidden(2) = True
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

Private Sub FillBranch()
    Dim L_Rst As New ADODB.Recordset
    cmbBranch.Clear
'    cmbBranch.AddItem "Â„Â ‘⁄»Â Â«"
'    cmbBranch.ItemData(cmbBranch.NewIndex) = 0
    Set L_Rst = RunStoredProcedure2RecordSet("Get_All_Branches")
    
    Do While L_Rst.EOF = False
        cmbBranch.AddItem L_Rst!nvcBranchName
        cmbBranch.ItemData(cmbBranch.NewIndex) = L_Rst!Branch
        L_Rst.MoveNext
    Loop
    
    L_Rst.Close: Set L_Rst = Nothing
    
    If cmbBranch.ListCount > 0 Then cmbBranch.ListIndex = 0
End Sub

Private Sub FillInventory()
    Dim L_Rst As New ADODB.Recordset
    
    cmbInventory.Clear
    ReDim Parameter(1) As Parameter
    Parameter(0) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
    Parameter(1) = GenerateInputParameter("@Branch", adInteger, 4, cmbBranch.ItemData(cmbBranch.ListIndex))
    Set L_Rst = RunParametricStoredProcedure2Rec("GetInventory_Branch", Parameter)
    
    If Not (L_Rst.EOF = True And L_Rst.BOF = True) Then
        Do While L_Rst.EOF <> True
            cmbInventory.AddItem L_Rst.Fields("Description")
            cmbInventory.ItemData(cmbInventory.NewIndex) = L_Rst!InventoryNo
            L_Rst.MoveNext
        Loop
    End If
    
    L_Rst.Close: Set L_Rst = Nothing
    If cmbInventory.ListCount > 0 Then cmbInventory.ListIndex = 0
End Sub

Private Sub fwBtnGoodFind_Click()
    frmFindGoods.Show vbModal
    fwBtnGoodFind.Caption = mvarName
    fwBtnGoodFind.Tag = mvarcode
    If mvarcode = 0 Then
        StoreDataUpdate.Enabled = False
    Else
        StoreDataUpdate.Enabled = True
        Call StoreDataUpdate_Click
    End If
    txtBarcode.Text = mvarBarcodeName
End Sub

Private Sub ResizeKit1_ExitResize(ByVal XScale As Double, ByVal YScale As Double)
    If formloadFlag = True Then
        SaveSetting strMainKey, Me.Name, "Height", Me.Height
        SaveSetting strMainKey, Me.Name, "Width", Me.Width
    End If
End Sub

Public Sub StoreDataUpdate_Click()
    If Trim(txtDateFrom.ClipText) = "" Or Trim(txtDateTo.ClipText) = "" Then
        ShowDisMessage "·ÿ›«  «—ÌŒ „⁄ »—Ì —« Ê«—œ ﬂ‰Ìœ ", 2500
        Exit Sub
    End If
    FillvsGood ("")
End Sub

Public Sub Printing()
    On Error GoTo Err_Handler
    If cmbSalMali.ListIndex = -1 Then
        ShowDisMessage "·ÿ›« ”«· „«·Ì —« «‰ Œ«» ﬂ‰Ìœ", 2000
        Exit Sub
    ElseIf cmbBranch.ListIndex = -1 Then
        ShowDisMessage "·ÿ›« ‘⁄»Â —« «‰ Œ«» ﬂ‰Ìœ", 2000
        Exit Sub
    ElseIf cmbInventory.ListIndex = -1 Then
        ShowDisMessage "·ÿ›« «‰»«— —« «‰ Œ«» ﬂ‰Ìœ", 2000
        Exit Sub
    ElseIf Val(fwBtnGoodFind.Tag) = 0 Then
        ShowDisMessage "·ÿ›« Ìﬂ ﬂ«·« «‰ Œ«» ﬂ‰Ìœ", 2000
        Exit Sub
    End If
'    With vsGood
'        RunNonParametricStoredProcedure "Delete_tblPrint_TurnOver"
        
'        ReDim Parameter(5) As Parameter
'        For i = 1 To .Rows - 2
'                Parameter(0) = GenerateInputParameter("@SanadNo", adInteger, 4, Val(.TextMatrix(i, IndexColSanad)))
'                Parameter(1) = GenerateInputParameter("@Date", adVarWChar, 20, .TextMatrix(i, IndexColDate))
'                Parameter(2) = GenerateInputParameter("@Description", adVarWChar, 50, .TextMatrix(i, IndexColDescription))
'                Parameter(3) = GenerateInputParameter("@Input", adDouble, 8, Val(.TextMatrix(i, IndexColInput)))
'                Parameter(4) = GenerateInputParameter("@Output", adDouble, 8, Val(.TextMatrix(i, IndexColOutput)))
'                Parameter(5) = GenerateInputParameter("@Mojodi", adDouble, 8, Val(.TextMatrix(i, IndexColMojodi)))
'
'                RunParametricStoredProcedure "Insert_tblPrint_TurnOver", Parameter
'        Next i
        
        ReDim Parameter(8) As Parameter
        Parameter(0) = GenerateInputParameter("@SystemDate", adVarWChar, 50, Mid(clsDate.shamsi(Date), 3, 8))
        Parameter(1) = GenerateInputParameter("@SystemDay", adVarWChar, 50, clsDate.Find_DayOfWeek(Weekday(Date, vbSaturday)))
        Parameter(2) = GenerateInputParameter("@SystemTime", adVarWChar, 50, Mid(Str(time), 1, 5))
        Parameter(3) = GenerateInputParameter("@GoodCode1", adInteger, 4, fwBtnGoodFind.Tag)
        Parameter(4) = GenerateInputParameter("@Date1", adVarWChar, 50, txtDateFrom.Text)
        Parameter(5) = GenerateInputParameter("@Date2", adVarWChar, 50, txtDateTo.Text)
        Parameter(6) = GenerateInputParameter("@InventoryNo1", adInteger, 4, cmbInventory.ItemData(cmbInventory.ListIndex))
        Parameter(7) = GenerateInputParameter("@Branch1", adInteger, 4, cmbBranch.ItemData(cmbBranch.ListIndex))
        Parameter(8) = GenerateInputParameter("@AccountYear1", adInteger, 2, cmbSalMali.ItemData(cmbSalMali.ListIndex))
        
      '  CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\RepOrder.rpt"
        CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\RepTurnOver_A4.rpt"
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
        
        CrystalReport1.ReportTitle = " ê—œ‘ ò«·« œ— «‰»«— "
        CrystalReport1.Destination = crptToWindow 'crptToPrinter '
        
        Dim intIndex As Integer
'        For intIndex = 0 To 100
'            CrystalReport1.ParameterFields(i) = ""
'        Next intIndex
        
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
'    End With

Exit Sub
Err_Handler:
    LogSaveNew "frmGoodTurnOver => ", err.Description, err.Number, err.Source, "Printing"
    ShowErrorMessage
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
        ShowDisMessage " . «Ì‰ »«—ﬂœ œ— ”Ì” „  ⁄—Ì› ‰‘œÂ «”  ", 2000
    End If
    If fwBtnGoodFind.Tag = 0 Then
        StoreDataUpdate.Enabled = False
    Else
        StoreDataUpdate.Enabled = True
        StoreDataUpdate.SetFocus
    End If
End Sub

Private Sub vsGood_Click()
    With vsGood
'        If .Col = IndexColCode Then
'            .Sort = flexSortNumericAscending
'            .ColSort(IndexColCode) = flexSortGenericAscending + flexSortGenericDescending
'        End If
    End With
End Sub

Private Sub vsGood_DblClick()
    With vsGood
        If .Row > 0 Then
            If Trim(.TextMatrix(.Row, IndexColDescription)) = "›—Ê‘" Then
                FillvsGood (.TextMatrix(.Row, IndexColDate))
            End If
        End If
    End With
End Sub

