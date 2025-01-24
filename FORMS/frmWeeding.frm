VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{7AEDC602-D94C-11D1-BB7A-00E0290EA3C9}#1.0#0"; "ResizeKit.ocx"
Object = "{75D4F767-8785-11D3-93AD-0000832EF44D}#2.14#0"; "FAST2001.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmWeeding 
   ClientHeight    =   8940
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14040
   BeginProperty Font 
      Name            =   "B Homa"
      Size            =   9.75
      Charset         =   178
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmWeeding.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8940
   ScaleWidth      =   14040
   Begin VB.TextBox txtmembershipid 
      Alignment       =   1  'Right Justify
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
      Left            =   7680
      RightToLeft     =   -1  'True
      TabIndex        =   22
      Top             =   480
      Width           =   2265
   End
   Begin VB.ComboBox CmbPartition 
      Height          =   465
      ItemData        =   "frmWeeding.frx":A4C2
      Left            =   2880
      List            =   "frmWeeding.frx":A4C4
      TabIndex        =   20
      Top             =   3000
      Width           =   2655
   End
   Begin VB.ComboBox cmbDay 
      Height          =   465
      ItemData        =   "frmWeeding.frx":A4C6
      Left            =   4200
      List            =   "frmWeeding.frx":A4C8
      TabIndex        =   4
      Top             =   1080
      Width           =   1335
   End
   Begin VB.TextBox txtGetPrice 
      Alignment       =   1  'Right Justify
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
      Left            =   3600
      TabIndex        =   8
      Top             =   2400
      Width           =   1900
   End
   Begin VB.TextBox txtName 
      Alignment       =   1  'Right Justify
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
      Left            =   600
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   480
      Width           =   1900
   End
   Begin VB.TextBox txtTotalPrice 
      Alignment       =   1  'Right Justify
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
      Left            =   8040
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   2400
      Width           =   1900
   End
   Begin VB.TextBox txtMoaref 
      Alignment       =   1  'Right Justify
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
      Left            =   7200
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   3000
      Width           =   2745
   End
   Begin VB.TextBox txtFamily 
      Alignment       =   1  'Right Justify
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
      Left            =   3720
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   480
      Width           =   1900
   End
   Begin VSFlex7LCtl.VSFlexGrid vsWeeding 
      Height          =   5235
      Left            =   0
      TabIndex        =   0
      Top             =   3600
      Width           =   13995
      _cx             =   24686
      _cy             =   9234
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
      BackColor       =   16777215
      ForeColor       =   -2147483640
      BackColorFixed  =   16761024
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   16777215
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
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   12
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   500
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmWeeding.frx":A4CA
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
      ExplorerBar     =   3
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
      Left            =   12600
      Top             =   0
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   926
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   178
         Weight          =   700
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
   Begin MSMask.MaskEdBox mskDateWeeding 
      Height          =   585
      Left            =   8520
      TabIndex        =   3
      Top             =   1080
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   1032
      _Version        =   393216
      MaxLength       =   8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Nazanin"
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
   Begin MSMask.MaskEdBox mskStartTime 
      Height          =   585
      Left            =   8520
      TabIndex        =   5
      Top             =   1680
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   1032
      _Version        =   393216
      MaxLength       =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Nazanin"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "##:##"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox mskEndTime 
      Height          =   585
      Left            =   4200
      TabIndex        =   6
      Top             =   1680
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   1032
      _Version        =   393216
      MaxLength       =   5
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Nazanin"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "##:##"
      PromptChar      =   " "
   End
   Begin RESIZEKITLibCtl.ResizeKit ResizeKit1 
      Height          =   480
      Left            =   0
      OleObjectBlob   =   "frmWeeding.frx":A618
      TabIndex        =   24
      Top             =   0
      Width           =   480
   End
   Begin VB.Label lblMembershipWeeding 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "òœ «‘ —«ò"
      ForeColor       =   &H80000002&
      Height          =   405
      Left            =   10080
      RightToLeft     =   -1  'True
      TabIndex        =   23
      Top             =   600
      Width           =   1275
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "»Œ‘"
      ForeColor       =   &H80000002&
      Height          =   405
      Left            =   5640
      RightToLeft     =   -1  'True
      TabIndex        =   21
      Top             =   3000
      Width           =   1395
   End
   Begin VB.Label lblFamily 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "* ‰«„ Œ«‰Ê«œêÌ"
      ForeColor       =   &H80000002&
      Height          =   405
      Left            =   5640
      RightToLeft     =   -1  'True
      TabIndex        =   19
      Top             =   480
      Width           =   1365
   End
   Begin VB.Label lblName 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "‰«„"
      ForeColor       =   &H80000002&
      Height          =   405
      Left            =   2520
      RightToLeft     =   -1  'True
      TabIndex        =   18
      Top             =   480
      Width           =   645
   End
   Begin VB.Label lblTitel 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "«ÿ·«⁄«   „Ã«·”"
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
      Left            =   5760
      RightToLeft     =   -1  'True
      TabIndex        =   17
      Top             =   0
      Width           =   2775
   End
   Begin VB.Label lblStartTime 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "”«⁄  Ê—Êœ"
      ForeColor       =   &H80000002&
      Height          =   405
      Left            =   9960
      RightToLeft     =   -1  'True
      TabIndex        =   16
      Top             =   1800
      Width           =   1395
   End
   Begin VB.Label lblEndTime 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "”«⁄  Œ—ÊÃ"
      ForeColor       =   &H80000002&
      Height          =   405
      Left            =   5640
      RightToLeft     =   -1  'True
      TabIndex        =   15
      Top             =   1800
      Width           =   1395
   End
   Begin VB.Label lblMoaref 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "„⁄—›"
      ForeColor       =   &H80000002&
      Height          =   405
      Left            =   9960
      RightToLeft     =   -1  'True
      TabIndex        =   14
      Top             =   3000
      Width           =   1395
   End
   Begin VB.Label lblDateWeeding 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   " «—ÌŒ "
      ForeColor       =   &H80000002&
      Height          =   405
      Left            =   9960
      RightToLeft     =   -1  'True
      TabIndex        =   13
      Top             =   1200
      Width           =   1395
   End
   Begin VB.Label lblTotalPrice 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "„»·€ ﬂ·"
      ForeColor       =   &H80000002&
      Height          =   405
      Left            =   9960
      RightToLeft     =   -1  'True
      TabIndex        =   12
      Top             =   2400
      Width           =   1395
   End
   Begin VB.Label lblGetPrice 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "»Ì⁄«‰Â"
      ForeColor       =   &H80000002&
      Height          =   405
      Left            =   5640
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   2400
      Width           =   1395
   End
   Begin VB.Label lblFax 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "—Ê“"
      ForeColor       =   &H80000002&
      Height          =   405
      Left            =   5640
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   1200
      Width           =   1395
   End
End
Attribute VB_Name = "frmWeeding"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private clsDate As New clsDate
Private Rc As New ADODB.Recordset
Private rctmp As New ADODB.Recordset
Public mvarcode As String
Private MyFormAddEditMode As EnumAddEditMode
Dim Parameter() As Parameter
Dim i As Integer
Dim OldTafsili As Long
Dim intWeedingNo As Integer

Public Sub Delete()
    'Case
        
        
            ReDim Parameter(1) As Parameter
            Parameter(0) = GenerateInputParameter("@intWeedingNo", adInteger, 4, intWeedingNo)
            Parameter(1) = GenerateOutputParameter("@Deleted", adInteger, 4)
            
            Dim Deleted As Long
            Deleted = RunParametricStoredProcedure("Delete_tblTotal_Weeding_ByPk_intWeedingNo", Parameter)
            If Deleted <> False Then
                frmMsg.fwlblMsg.Caption = "Õ–› »« „Ê›ﬁÌ  «‰Ã«„ ‘œ"
                frmMsg.fwBtn(0).Visible = False
                frmMsg.fwBtn(1).ButtonType = flwButtonOk
                frmMsg.fwBtn(1).Caption = "ﬁ»Ê·"
                frmMsg.Show vbModal
                
            Else
                frmMsg.fwlblMsg.Caption = "Õ–› «‰Ã«„ ‰‘œ"
                frmMsg.fwBtn(0).Visible = False
                frmMsg.fwBtn(1).ButtonType = flwButtonOk
                frmMsg.fwBtn(1).Caption = "ﬁ»Ê·"
                frmMsg.Show vbModal
                TxtName.SetFocus
                Exit Sub
            End If

        'End Select
        MyFormAddEditMode = AddMode
    SetFirstToolBar
    Add
    
    Exit Sub
RollBack:
End Sub

Private Sub FillvsWeeding()

    Dim Rst As New ADODB.Recordset
    ReDim Parameter(0) As Parameter
    
    
    Parameter(0) = GenerateInputParameter("@intWeedingNo", adInteger, 4, -1)
    Set Rst = RunParametricStoredProcedure2Rec("Get_tblTotal_Weeding_ByPK_intWeedingNo", Parameter)
    
    With vsWeeding
        .Rows = 1
        i = 0
        While Rst.EOF <> True
            .Rows = .Rows + 1
            i = .Rows - 1
            .TextMatrix(i, 0) = i
            .TextMatrix(i, 1) = IIf(IsNull(Rst!intPartitionId), "", Rst!intPartitionId)
            .TextMatrix(i, 2) = Rst!nvcName & " " & Rst!nvcFamily
            .TextMatrix(i, 3) = IIf(IsNull(Rst!Membercode), "", Rst!Membercode)
            .TextMatrix(i, 4) = Rst!nvcDateWeeding
            .TextMatrix(i, 5) = Rst!intDay
            .TextMatrix(i, 6) = Rst!nvcMoaref
            .TextMatrix(i, 7) = Rst!nvcUseStartTime
            .TextMatrix(i, 8) = Rst!nvcUseEndTime
            .TextMatrix(i, 9) = Rst!intTotalPrice
            .TextMatrix(i, 10) = Rst!intGetPrice
            .TextMatrix(i, 11) = Rst!intWeedingNo
            Rst.MoveNext
        Wend
    End With
    Set Rst = Nothing
End Sub

Private Sub Form_Activate()

    VarActForm = Me.Name
    SetFirstToolBar

End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call PresetScreenSaver
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
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

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()

    CenterTop Me
    
''    If ClsFormAccess.frmSupplier = False Then
''        Unload Me
''        Exit Sub
''    End If
    
    VarActForm = Me.Name
     
     cmbDay.Clear
     cmbDay.AddItem "‘‰»Â"
     cmbDay.ItemData(cmbDay.ListCount - 1) = 1
     vsWeeding.ColComboList(5) = "#1" & "; ‘‰»Â|"
     cmbDay.AddItem "Ìﬂ‘‰»Â"
     cmbDay.ItemData(cmbDay.ListCount - 1) = 2
     vsWeeding.ColComboList(5) = vsWeeding.ColComboList(5) & "#2" & "; Ìﬂ‘‰»Â|"
     cmbDay.AddItem "œÊ‘‰»Â"
     cmbDay.ItemData(cmbDay.ListCount - 1) = 3
     vsWeeding.ColComboList(5) = vsWeeding.ColComboList(5) & "#3" & "; œÊ‘‰»Â|"
     cmbDay.AddItem "”Â ‘‰»Â"
     cmbDay.ItemData(cmbDay.ListCount - 1) = 4
     vsWeeding.ColComboList(5) = vsWeeding.ColComboList(5) & "#4" & "; ”Â ‘‰»Â|"
     cmbDay.AddItem "çÂ«—‘‰»Â"
     cmbDay.ItemData(cmbDay.ListCount - 1) = 5
     vsWeeding.ColComboList(5) = vsWeeding.ColComboList(5) & "#5" & "; çÂ«—‘‰»Â|"
     cmbDay.AddItem "Å‰Ã‘‰»Â"
     cmbDay.ItemData(cmbDay.ListCount - 1) = 6
     vsWeeding.ColComboList(5) = vsWeeding.ColComboList(5) & "#6" & "; Å‰Ã‘‰»Â|"
     cmbDay.AddItem "Ã„⁄Â"
     cmbDay.ItemData(cmbDay.ListCount - 1) = 7
     vsWeeding.ColComboList(5) = vsWeeding.ColComboList(5) & "#7" & "; Ã„⁄Â|"
     cmbDay.ListIndex = 0
     
'     vsWeeding.ColHidden(1) = True
     
    cmbPartition.Clear
    ReDim Parameter(0) As Parameter
    Parameter(0) = GenerateInputParameter("@Branch", adInteger, 4, CurrentBranch)
    Set rctmp = RunParametricStoredProcedure2Rec("Get_tPartitions", Parameter)
    vsWeeding.ColComboList(1) = vsWeeding.BuildComboList(rctmp, "PartitionDescription", "PartitionID")
    Set rctmp = RunParametricStoredProcedure2Rec("Get_tPartitions", Parameter)
    
    Do While rctmp.EOF = False
        cmbPartition.AddItem rctmp!PartitionDescription
        cmbPartition.ItemData(cmbPartition.NewIndex) = rctmp!PartitionID
        rctmp.MoveNext
    Loop
    rctmp.Close
    If cmbPartition.ListCount > 0 Then cmbPartition.ListIndex = 0
      
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

     
     Add

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Set Rc = Nothing
    Set rctmp = Nothing
    
    If PosConnection.State = adStateOpen Then PosConnection.Close: Set PosConnection = Nothing
    
    Set clsDate = Nothing
    Set mdifrm.FileCls = Nothing
        
    VarActForm = ""
    Unload frmWeeding
    
    mdifrm.Toolbar1.Buttons(20).Enabled = False
    mdifrm.Toolbar1.Buttons(21).Enabled = False
    mdifrm.Toolbar1.Buttons(23).Enabled = True
    mdifrm.Toolbar1.Buttons(24).Enabled = True
    mdifrm.Toolbar1.Buttons(25).Enabled = True
    mdifrm.Toolbar1.Buttons(26).Enabled = True
    mdifrm.Toolbar1.Buttons(27).Enabled = True

    SaveSetting strMainKey, Me.Name, "Left", Me.Left
    SaveSetting strMainKey, Me.Name, "Top", Me.Top

End Sub


Public Sub Cancel()
    Select Case MyFormAddEditMode
        Case AddMode 'new
            DefaultSettings
            MyFormAddEditMode = AddMode
            SetFirstToolBar
        Case EditMode 'edit
            GetDataDetail
            MyFormAddEditMode = ViewMode
            SetFirstToolBar
    End Select
End Sub

Public Sub DefaultSettings()

    On Error Resume Next
    
    
    cmbDay.ListIndex = 0
    
    On Error GoTo 0
    
    txtFamily.Text = ""
    txtGetPrice.Text = ""
    txtMoaref.Text = ""
    TxtName.Text = ""
    txtTotalPrice.Text = ""
    mskDateWeeding.Text = "  /  /  "
    mskStartTime.Text = "  :  "
    mskEndTime.Text = "  :  "
      
End Sub

Public Sub Add()

    If MyFormAddEditMode = EditMode Then
        DefaultSettings
    End If
    MyFormAddEditMode = AddMode
    DefaultSettings
    SetFirstToolBar
    FillvsWeeding
End Sub

Public Sub ExitSub()
If PosConnection.State = adStateOpen Then PosConnection.Close: Set PosConnection = Nothing
    Unload Me
End Sub

Public Sub Update()
    If MyFormAddEditMode = ViewMode Then Exit Sub
    Dim strBinBuyState As String
    Dim intBuyState As Integer
    Select Case MyFormAddEditMode
        Case AddMode
            If txtFamily.Text = "" Then
                frmMsg.fwlblMsg.Caption = "‰«„ Œ«‰Ê«œêÌ —« Å— ﬂ‰Ìœ"
                frmMsg.fwBtn(0).Visible = False
                frmMsg.fwBtn(1).ButtonType = flwButtonOk
                frmMsg.fwBtn(1).Caption = "ﬁ»Ê·"
                frmMsg.Show vbModal
                Exit Sub
            End If
            ReDim Parameter(11) As Parameter
            Parameter(0) = GenerateInputParameter("@nvcName", adVarWChar, 50, TxtName.Text)
            Parameter(1) = GenerateInputParameter("@nvcFamily", adVarWChar, 50, txtFamily.Text)
            Parameter(2) = GenerateInputParameter("@nvcDateWeeding", adVarWChar, 10, mskDateWeeding.Text)
            Parameter(3) = GenerateInputParameter("@intDay", adInteger, 4, cmbDay.ItemData(cmbDay.ListIndex))
            Parameter(4) = GenerateInputParameter("@nvcMoaref", adVarWChar, 50, txtMoaref)
            Parameter(5) = GenerateInputParameter("@nvcUseStartTime", adVarWChar, 5, mskStartTime.Text)
            Parameter(6) = GenerateInputParameter("@nvcUseEndTime", adVarWChar, 5, mskEndTime.Text)
            Parameter(7) = GenerateInputParameter("@intTotalPrice", adInteger, 4, Val(txtTotalPrice.Text))
            Parameter(8) = GenerateInputParameter("@intGetPrice", adInteger, 4, Val(txtGetPrice.Text))
            Parameter(9) = GenerateInputParameter("@intPartitionId", adInteger, 4, cmbPartition.ItemData(cmbPartition.ListIndex))
            Parameter(10) = GenerateInputParameter("@Membercode", adBigInt, 8, Val(txtMembershipId.Text))
            Parameter(11) = GenerateOutputParameter("@intWeedingNo", adInteger, 4)
            
            Dim LastCode As Long
            LastCode = RunParametricStoredProcedure("Insert_tblTotal_Weeding", Parameter)
            If LastCode <> -1 Then
                frmMsg.fwlblMsg.Caption = "À»  «ÿ·«⁄«  ÃœÌœ »« „Ê›ﬁÌ  «‰Ã«„ ‘œ"
                frmMsg.fwBtn(0).Visible = False
                frmMsg.fwBtn(1).ButtonType = flwButtonOk
                frmMsg.fwBtn(1).Caption = "ﬁ»Ê·"
                frmMsg.Show vbModal
            Else
                frmMsg.fwlblMsg.Caption = "À»  «‰Ã«„ ‰‘œ"
                frmMsg.fwBtn(0).Visible = False
                frmMsg.fwBtn(1).ButtonType = flwButtonOk
                frmMsg.fwBtn(1).Caption = "ﬁ»Ê·"
                frmMsg.Show vbModal
                txtFamily.SetFocus
                Exit Sub
            End If
            
            
        Case EditMode
        
        
            ReDim Parameter(12) As Parameter
            Parameter(0) = GenerateInputParameter("@nvcName", adVarWChar, 50, TxtName.Text)
            Parameter(1) = GenerateInputParameter("@nvcFamily", adVarWChar, 50, txtFamily.Text)
            Parameter(2) = GenerateInputParameter("@nvcDateWeeding", adVarWChar, 10, mskDateWeeding.Text)
            Parameter(3) = GenerateInputParameter("@intDay", adInteger, 4, cmbDay.ItemData(cmbDay.ListIndex))
            Parameter(4) = GenerateInputParameter("@nvcMoaref", adVarWChar, 50, txtMoaref)
            Parameter(5) = GenerateInputParameter("@nvcUseStartTime", adVarWChar, 5, mskStartTime.Text)
            Parameter(6) = GenerateInputParameter("@nvcUseEndTime", adVarWChar, 5, mskEndTime.Text)
            Parameter(7) = GenerateInputParameter("@intTotalPrice", adInteger, 4, Val(txtTotalPrice.Text))
            Parameter(8) = GenerateInputParameter("@intGetPrice", adInteger, 4, Val(txtGetPrice.Text))
            Parameter(9) = GenerateInputParameter("@intWeedingNo", adInteger, 4, intWeedingNo)
            Parameter(10) = GenerateInputParameter("@intPartitionId", adInteger, 4, cmbPartition.ItemData(cmbPartition.ListIndex))
            Parameter(11) = GenerateInputParameter("@Membercode", adBigInt, 8, Val(txtMembershipId.Text))
            Parameter(12) = GenerateOutputParameter("@Updated", adInteger, 4)
            
            Dim Updated As Long
            Updated = RunParametricStoredProcedure("Update_tblTotal_Weeding_ByPk_intWeedingNo", Parameter)
            If Updated <> False Then
                frmMsg.fwlblMsg.Caption = " €ÌÌ—«  »« „Ê›ﬁÌ  «‰Ã«„ ‘œ"
                frmMsg.fwBtn(0).Visible = False
                frmMsg.fwBtn(1).ButtonType = flwButtonOk
                frmMsg.fwBtn(1).Caption = "ﬁ»Ê·"
                frmMsg.Show vbModal
                
            Else
                frmMsg.fwlblMsg.Caption = " €ÌÌ—«  «‰Ã«„ ‰‘œ"
                frmMsg.fwBtn(0).Visible = False
                frmMsg.fwBtn(1).ButtonType = flwButtonOk
                frmMsg.fwBtn(1).Caption = "ﬁ»Ê·"
                frmMsg.Show vbModal
                TxtName.SetFocus
                Exit Sub
            End If

        End Select
    
    MyFormAddEditMode = AddMode
    SetFirstToolBar
    Add
    
    Exit Sub
RollBack:
        
End Sub


Public Sub Edit()
 
    MyFormAddEditMode = EditMode
    SetFirstToolBar
    
End Sub

Public Sub ExitForm()
    Unload Me
End Sub

Sub SetFirstToolBar()
    
    Dim Obj As Object
    
    AllButton vbOff, True
    
    mdifrm.Toolbar1.Buttons(1).Enabled = True   'Home
    mdifrm.Toolbar1.Buttons(2).Enabled = True   'PageUp
    mdifrm.Toolbar1.Buttons(3).Enabled = True   'PageDown
    mdifrm.Toolbar1.Buttons(4).Enabled = True   'End
    
    mdifrm.Toolbar1.Buttons(15).Enabled = True   'Print
    
    mdifrm.Toolbar1.Buttons(23).Enabled = True
    mdifrm.Toolbar1.Buttons(24).Enabled = True
    mdifrm.Toolbar1.Buttons(25).Enabled = True
    mdifrm.Toolbar1.Buttons(26).Enabled = True
    mdifrm.Toolbar1.Buttons(27).Enabled = True
    
 
    If MyFormAddEditMode = ViewMode Then  ' View Mode
 
        On Error Resume Next
        For Each Obj In Me
           Obj.Locked = True
        Next Obj
        On Error GoTo 0
        mdifrm.Toolbar1.Buttons(10).Enabled = True
        mdifrm.Toolbar1.Buttons(6).Enabled = True  'Add
        mdifrm.Toolbar1.Buttons(7).Enabled = True  'Edit
        mdifrm.Toolbar1.Buttons(8).Enabled = False  'Enter
        mdifrm.Toolbar1.Buttons(9).Enabled = False   'Esc
        
    ElseIf MyFormAddEditMode = AddMode Then    'Add Mode
        
        On Error Resume Next
        For Each Obj In Me
                Obj.Locked = False
        Next Obj
        On Error GoTo 0
        
        mdifrm.Toolbar1.Buttons(6).Enabled = False  'Add
        mdifrm.Toolbar1.Buttons(7).Enabled = False  'Edit
        mdifrm.Toolbar1.Buttons(8).Enabled = True   'Enter
        mdifrm.Toolbar1.Buttons(9).Enabled = True   'Esc
    
    ElseIf MyFormAddEditMode = EditMode Then     'Edit
    
        On Error Resume Next
        For Each Obj In Me
                Obj.Locked = False
        Next Obj
        On Error GoTo 0
        
        mdifrm.Toolbar1.Buttons(6).Enabled = True  'Add
        mdifrm.Toolbar1.Buttons(7).Enabled = False  'Edit
        mdifrm.Toolbar1.Buttons(8).Enabled = True   'Enter
        mdifrm.Toolbar1.Buttons(9).Enabled = True   'Esc
        
    End If
    
    HeaderLabel Val(MyFormAddEditMode), fwlblMode
    
End Sub
Sub GetDataDetail()
    
    DefaultSettings
    
    Dim TempStr As String
    
    ReDim Parameter(0) As Parameter
    Parameter(0) = GenerateInputParameter("@intWeedingNo", adInteger, 4, Val(intWeedingNo))
    Set rctmp = RunParametricStoredProcedure2Rec("Get_tblTotal_Weeding_ByPK_intWeedingNo", Parameter)
    Dim ii As Integer
    If Not (rctmp.BOF Or rctmp.EOF) Then
        
            TxtName.Text = rctmp!nvcName
            txtFamily.Text = rctmp!nvcFamily
            mskDateWeeding.Text = rctmp!nvcDateWeeding
            txtMoaref.Text = rctmp!nvcMoaref
            mskStartTime.Text = rctmp!nvcUseStartTime
            mskEndTime.Text = rctmp!nvcUseEndTime
            txtTotalPrice.Text = rctmp!intTotalPrice
            txtGetPrice.Text = rctmp!intGetPrice
        
        For i = 0 To cmbDay.ListCount - 1
            If cmbDay.ItemData(i) = rctmp!intDay Then
                cmbDay.ListIndex = i
                Exit For
            End If
        Next i
               
    End If
    rctmp.Close
    
    
End Sub




Private Sub ResizeKit1_ExitResize(ByVal XScale As Double, ByVal YScale As Double)
    If formloadFlag = True Then
        SaveSetting strMainKey, Me.Name, "Height", Me.Height
        SaveSetting strMainKey, Me.Name, "Width", Me.Width
    End If

End Sub

Private Sub txtFamily_GotFocus()
    Dim Rst As New ADODB.Recordset
    If Rst.State = 1 Then Rst.Close
    ReDim Parameter(0) As Parameter
    Parameter(0) = GenerateInputParameter("@MembershipId", adBigInt, 8, Val(txtMembershipId.Text))
    Set Rst = RunParametricStoredProcedure2Rec("Get_Customer_weeding", Parameter)
         
     If Not (Rst.EOF = True And Rst.BOF = True) Then
        
        txtFamily.Text = Rst!FullName
        
    End If
               
End Sub


Private Sub vsWeeding_AfterSort(ByVal Col As Long, Order As Integer)
    With vsWeeding
        If Col = 3 And .Rows > 1 Then
            For i = 1 To .Rows - 2
                If (Val(.TextMatrix(i, 3)) - Val(.TextMatrix(i + 1, 3)) > 1 And Order = 2) Or (Val(.TextMatrix(i + 1, 3)) - Val(.TextMatrix(i, 3)) > 1 And Order = 1) Then
                    .Cell(flexcpBackColor, i, .FixedCols, i, .Cols - 1) = 8421631
                Else
                    .Cell(flexcpBackColor, i, .FixedCols, i, .Cols - 1) = &H80000005
                End If
            Next i
        End If
    End With
End Sub

Private Sub vsWeeding_Click()
    
    intWeedingNo = vsWeeding.TextMatrix(vsWeeding.Row, 11)
    MyFormAddEditMode = ViewMode
    GetDataDetail
    SetFirstToolBar
    HeaderLabel Val(MyFormAddEditMode), Me.fwlblMode

End Sub




