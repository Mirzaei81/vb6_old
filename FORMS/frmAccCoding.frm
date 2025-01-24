VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{7AEDC602-D94C-11D1-BB7A-00E0290EA3C9}#1.0#0"; "ResizeKit.ocx"
Object = "{75D4F767-8785-11D3-93AD-0000832EF44D}#2.14#0"; "FAST2001.ocx"
Begin VB.Form frmAccCoding 
   ClientHeight    =   9420
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9540
   Icon            =   "frmAccCoding.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9420
   ScaleWidth      =   9540
   Begin VB.CommandButton cmdAccounting 
      Caption         =   "ÍÓÇÈÏÇÑí"
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
      TabIndex        =   12
      Top             =   1680
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Frame frameDescription 
      Height          =   1695
      Left            =   1440
      TabIndex        =   3
      Top             =   600
      Width           =   7545
      Begin VB.TextBox txtDesc 
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
         Height          =   450
         Left            =   4170
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   360
         Width           =   2415
      End
      Begin VB.TextBox txtTafsili 
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         Left            =   720
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   960
         Width           =   2415
      End
      Begin VB.ComboBox cmbKol 
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
         ItemData        =   "frmAccCoding.frx":A4C2
         Left            =   720
         List            =   "frmAccCoding.frx":A4C4
         RightToLeft     =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   360
         Width           =   2415
      End
      Begin VB.ComboBox cmbMoein 
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
         ItemData        =   "frmAccCoding.frx":A4C6
         Left            =   4170
         List            =   "frmAccCoding.frx":A4C8
         RightToLeft     =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1005
         Width           =   2415
      End
      Begin VB.Label lblTafsili 
         Alignment       =   1  'Right Justify
         Caption         =   "ÊÝÖíáí"
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   375
         Left            =   3000
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   960
         Width           =   735
      End
      Begin VB.Label lblKol 
         Alignment       =   1  'Right Justify
         Caption         =   "ßá"
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   375
         Left            =   3000
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   360
         Width           =   735
      End
      Begin VB.Label lblMoein 
         Alignment       =   1  'Right Justify
         Caption         =   "ãÚíä"
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   375
         Left            =   6405
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   960
         Width           =   735
      End
      Begin VB.Label lblDesc 
         Alignment       =   1  'Right Justify
         Caption         =   "ÔÑÍ"
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   375
         Left            =   6405
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Top             =   435
         Width           =   735
      End
   End
   Begin FLWCtrls.FWLabel3D fwlblMode 
      Height          =   525
      Left            =   8040
      Top             =   0
      Width           =   1545
      _ExtentX        =   2725
      _ExtentY        =   926
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Nazanin"
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
      Caption         =   "ãÑæÑ"
      Alignment       =   2
   End
   Begin VSFlex7LCtl.VSFlexGrid vsAccCode 
      Height          =   6915
      Left            =   1440
      TabIndex        =   0
      Top             =   2400
      Width           =   7545
      _cx             =   13309
      _cy             =   12197
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Nazanin"
         Size            =   11.25
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
      AllowUserResizing=   4
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   6
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmAccCoding.frx":A4CA
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   0
      AutoSearch      =   0
      AutoSearchDelay =   2
      MultiTotals     =   0   'False
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
   Begin RESIZEKITLibCtl.ResizeKit ResizeKit1 
      Height          =   480
      Left            =   0
      OleObjectBlob   =   "frmAccCoding.frx":A57F
      TabIndex        =   2
      Top             =   0
      Width           =   480
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ßÏ åÇí ÍÓÇÈÏÇÑí"
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
      Left            =   3480
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   0
      Width           =   2175
   End
End
Attribute VB_Name = "frmAccCoding"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i As Integer
Dim cnn As New ADODB.Connection
Dim Parameter() As Parameter
Dim MyFormAddEditMode As EnumAddEditMode
Dim Rst As New ADODB.Recordset
Dim cn As New ADODB.Connection


Public Sub ExitForm()
    Unload Me
End Sub

Private Sub CancelButton_Click()
  Unload Me
End Sub

Private Sub cmbKol_Click()
    If cmbKol.ListIndex > 0 Then
        FillMoein
    End If
End Sub


Private Sub cmdAccounting_Click()
'    If clsArya.ExternalAccounting Then
'         modgl.ShowAccountingForm "frmKol", "ÍÓÇÈåÇí ßá"
'     '    modgl.ShowAccountingForm "frmKolList", "áíÓÊ ÍÓÇÈåÇí ˜á"
'     End If
End Sub

Private Sub Form_Activate()
    SetFirstToolBar
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
    If KeyCode = 13 Then
        With vsAccCode
           .Rows = .Rows + 1
           .Row = .Rows - 1
        End With
    End If

End Sub

Private Sub Form_Load()
    
    On Error GoTo ErrHandler
    If ClsFormAccess.frmAccCoding = False Then
        Unload Me
        Exit Sub
    End If

    CenterTop Me
    
    VarActForm = Me.Name
    
    Dim s As String
    
    With vsAccCode
        .Rows = 1
        .Cols = 7
        .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter
        .ColHidden(4) = True
        .ColHidden(5) = True
        .TextMatrix(0, 3) = " ßÏ ãÚíä"
        .TextMatrix(0, 6) = "ÔÑÍ ãÚíä"
        .ColHidden(3) = False
        .ColHidden(6) = False
         For i = 0 To .Cols - 1
            .ColWidth(i) = Val(GetSetting(strMainKey, "vsAccCode", "Col" & i))
            If .ColWidth(i) = 0 Then
                .ColWidth(i) = 1000       '
            End If
         Next i
        .ColAlignment(1) = flexAlignRightCenter
        .ColAlignment(2) = flexAlignRightCenter
        .ColAlignment(3) = flexAlignCenterCenter
        .ColAlignment(4) = flexAlignRightCenter
        .ColAlignment(5) = flexAlignRightCenter
        .ColAlignment(6) = flexAlignRightCenter
    End With
        
    GetDataDetail
    MyFormAddEditMode = ViewMode
    SetFirstToolBar
'
    FillKol
    cmbKol.ListIndex = -1
'    FillMoein
    
    
    If clsArya.ExternalAccounting = True And ClsFormAccess.AccountingAccess = True Then
        cmdAccounting.Enabled = True
    Else
        cmdAccounting.Enabled = False
    End If
    formloadFlag = False
    Me.Left = Val(GetSetting(strMainKey, "frmAccCoding", "Left"))
    If Val(GetSetting(strMainKey, "frmAccCoding", "Height")) > 0 Then
        Me.Height = Val(GetSetting(strMainKey, "frmAccCoding", "Height"))
    End If
    If Val(GetSetting(strMainKey, "frmAccCoding", "Width")) > 0 Then
        Me.Width = Val(GetSetting(strMainKey, "frmAccCoding", "Width"))
    End If
    Me.Top = Val(GetSetting(strMainKey, "frmAccCoding", "Top"))
    formloadFlag = True

Exit Sub
ErrHandler:
    ShowDisMessage err.Description, 1000
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    VarActForm = ""
    If cnn.State = adStateOpen Then cnn.Close: Set cnn = Nothing
    
    SaveSetting strMainKey, "frmAccCoding", "Left", Me.Left
    SaveSetting strMainKey, "frmAccCoding", "Top", Me.Top
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call PresetScreenSaver
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call PresetScreenSaver
End Sub

Public Sub Add()
    If MyFormAddEditMode <> EditMode Then
        MyFormAddEditMode = AddMode
        SetFirstToolBar
'        With vsAccCode
'            If Trim(.TextMatrix(.Row, 5)) <> "" Then
'                .Rows = .Rows + 1
'                .Row = .Rows - 1
'                .TextMatrix(.Row, 0) = Val(.TextMatrix(.Row - 1, 0)) + 1
'                .Cell(flexcpAlignment, 1, 5, .Row, 5) = flexAlignRightCenter
'            End If
'        End With
    txtDesc.Text = ""
    txtTafsili.Text = ""
    cmbKol.ListIndex = -1
    cmbMoein.ListIndex = -1
    End If
End Sub

Public Sub Edit()
    MyFormAddEditMode = EditMode
    SetFirstToolBar
End Sub

Public Sub Cancel()
    MyFormAddEditMode = ViewMode
    vsAccCode.Rows = 1
    GetDataDetail
    SetFirstToolBar
End Sub

Private Sub SetFirstToolBar()
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
    mdifrm.Toolbar1.Buttons(6).Enabled = True  'Add
    
    If MyFormAddEditMode = ViewMode Then  ' View Mode
    
'        mdifrm.Toolbar1.Buttons(6).Enabled = False  'Add
        mdifrm.Toolbar1.Buttons(6).Enabled = True
        mdifrm.Toolbar1.Buttons(7).Enabled = True  'Edit
        mdifrm.Toolbar1.Buttons(8).Enabled = False  'Enter
        mdifrm.Toolbar1.Buttons(9).Enabled = True   'Esc
        
        frameDescription.Enabled = False
          
    ElseIf MyFormAddEditMode = AddMode Then    'Add Mode
    
    '    mdifrm.Toolbar1.Buttons(6).Enabled = False  'Add
        mdifrm.Toolbar1.Buttons(7).Enabled = False  'Edit
        mdifrm.Toolbar1.Buttons(8).Enabled = True   'Enter
        mdifrm.Toolbar1.Buttons(9).Enabled = True   'Esc
        
        frameDescription.Enabled = True
        
    ElseIf MyFormAddEditMode = EditMode Then     'Edit
        
'        mdifrm.Toolbar1.Buttons(6).Enabled = True  'Add
        mdifrm.Toolbar1.Buttons(6).Enabled = False
        mdifrm.Toolbar1.Buttons(7).Enabled = False  'Edit
        mdifrm.Toolbar1.Buttons(8).Enabled = True   'Enter
        mdifrm.Toolbar1.Buttons(9).Enabled = True   'Esc
        
        frameDescription.Enabled = True
    End If
    
    HeaderLabel Val(MyFormAddEditMode), fwlblMode
End Sub

Private Sub ResizeKit1_ExitResize(ByVal XScale As Double, ByVal YScale As Double)
    If formloadFlag = True Then
        SaveSetting strMainKey, "frmAccCoding", "Height", Me.Height
        SaveSetting strMainKey, "frmAccCoding", "Width", Me.Width
    End If
End Sub

Private Sub vsAccCode_AfterSort(ByVal Col As Long, Order As Integer)
    For i = 1 To vsAccCode.Rows - 1
        vsAccCode.TextMatrix(i, 0) = i
    Next
End Sub

Public Sub Update()
    
    Dim Result As Integer
    Dim Obj As Object
    Dim TotalPayment As Long

With vsAccCode
'        For i = 1 To .Rows - 1
'            If Trim(.TextMatrix(i, 1)) = "" Or Trim(.TextMatrix(i, 2)) = "" Or Trim(.TextMatrix(i, 3)) = "" Or Trim(.TextMatrix(i, 4)) = "" Or Trim(.TextMatrix(i, 5)) = "" Then
'                frmMsg.fwlblMsg.Caption = "áØÝÇ ÇØáÇÚÇÊ ÖÑæÑí ÑÇ æÇÑÏ äãÇííÏ"
'                frmMsg.fwBtn(0).ButtonType = flwButtonOk
'                frmMsg.fwBtn(0).Caption = "ÞÈæá"
'                frmMsg.fwBtn(1).Visible = False
'                frmMsg.Show vbModal
'                Exit Sub
'             End If
'        Next i
        If txtDesc.Text = "" Or txtTafsili.Text = "" Or cmbKol.ListIndex = -1 Or cmbMoein.ListIndex = -1 Then
            frmMsg.fwlblMsg.Caption = "áØÝÇ ÇØáÇÚÇÊ ÖÑæÑí ÑÇ æÇÑÏ äãÇííÏ"
            frmMsg.fwBtn(0).ButtonType = flwButtonOk
            frmMsg.fwBtn(0).Caption = "ÞÈæá"
            frmMsg.fwBtn(1).Visible = False
            frmMsg.Show vbModal
            Exit Sub
        End If
   
    Select Case MyFormAddEditMode
        Case AddMode
            ReDim Parameter(6) As Parameter
'            Parameter(0) = GenerateInputParameter("@Code", adInteger, 4, Val(.TextMatrix(.Row, 0)) + 1)
            Parameter(0) = GenerateInputParameter("@Code", adInteger, 4, vsAccCode.Rows)
'            Parameter(1) = GenerateInputParameter("@Description", adVarChar, 50, .TextMatrix(.Row, 1))
'            Parameter(2) = GenerateInputParameter("@Description", adVarChar, 50, .TextMatrix(.Row, 1))
'            Parameter(3) = GenerateInputParameter("@Moein", adInteger, 4, .TextMatrix(.Row, 3))
'            Parameter(4) = GenerateInputParameter("@Tafsili", adInteger, 4, .TextMatrix(.Row, 4))
            Parameter(1) = GenerateInputParameter("@Description", adVarChar, 50, txtDesc.Text)
            Parameter(2) = GenerateInputParameter("@Kol", adInteger, 1, cmbKol.ItemData(cmbKol.ListIndex))
            Parameter(3) = GenerateInputParameter("@Moein", adInteger, 4, cmbMoein.ItemData(cmbMoein.ListIndex))
            Parameter(4) = GenerateInputParameter("@Tafsili", adInteger, 4, txtTafsili.Text)
            Parameter(5) = GenerateInputParameter("@Active", adBoolean, 1, 1)
            Parameter(6) = GenerateInputParameter("@MoeinDesc", adVarChar, 50, cmbMoein.Text)
            
            Result = RunParametricStoredProcedure("Insert_tblAcc_Sale", Parameter)
            
            If Result = 1 Then
                frmMsg.fwlblMsg.Caption = "ËÈÊ ÇØáÇÚÇÊ ÌÏíÏ ÈÇ ãæÝÞíÊ ÇíÇä íÇÝÊ"
                frmMsg.fwBtn(0).ButtonType = flwButtonOk
                frmMsg.fwBtn(0).Caption = "ÞÈæá"
                frmMsg.fwBtn(1).Visible = False
                frmMsg.Show vbModal
                vsAccCode.Rows = 1
'                MyFormAddEditMode = AddMode
                MyFormAddEditMode = ViewMode
                SetFirstToolBar
                GetDataDetail
            End If
        Case EditMode
            
            ReDim Parameter(5) As Parameter
'            For i = 1 To .Rows - 1
'                If InStr(.TextMatrix(i, 0), "*") > 0 Then 'new or edited records
                    ReDim Parameter(6) As Parameter
                    Parameter(0) = GenerateInputParameter("@Code", adInteger, 4, Val(.TextMatrix(vsAccCode.Row, 0)))
'                    Parameter(1) = GenerateInputParameter("@Description", adVarChar, 50, .TextMatrix(i, 1))
'                    Parameter(2) = GenerateInputParameter("@Kol", adInteger, 4, .TextMatrix(i, 2))
'                    Parameter(3) = GenerateInputParameter("@Moein", adInteger, 4, .TextMatrix(i, 3))
'                    Parameter(4) = GenerateInputParameter("@Tafsili", adInteger, 4, .TextMatrix(i, 4))
                    Parameter(1) = GenerateInputParameter("@Description", adVarChar, 50, txtDesc.Text)
                    Parameter(2) = GenerateInputParameter("@Kol", adInteger, 4, cmbKol.ItemData(cmbKol.ListIndex))
                    Parameter(3) = GenerateInputParameter("@Moein", adInteger, 4, cmbMoein.ItemData(cmbMoein.ListIndex))
                    Parameter(4) = GenerateInputParameter("@Tafsili", adInteger, 4, txtTafsili.Text)
                    Parameter(5) = GenerateInputParameter("@Active", adBoolean, 1, 1)
                    Parameter(6) = GenerateInputParameter("@MoeinDesc", adVarChar, 50, cmbMoein.Text)
                    
                    Result = RunParametricStoredProcedure("Update_tblAcc_Sale", Parameter)
'                End If
'            Next i
'
            
            frmMsg.fwlblMsg.Caption = "ÊÛííÑ ÇØáÇÚÇÊ  ÈÇ ãæÝÞíÊ ÇíÇä íÇÝÊ"
            frmMsg.fwBtn(0).ButtonType = flwButtonOk
            frmMsg.fwBtn(0).Caption = "ÞÈæá"
            frmMsg.fwBtn(1).Visible = False
            frmMsg.Show vbModal
            vsAccCode.Rows = 1
'            MyFormAddEditMode = AddMode
            MyFormAddEditMode = ViewMode
            SetFirstToolBar
            
            GetDataDetail
    End Select
End With
Exit Sub
ErrHandler:
    Select Case err.Number
        Case -2147217873
                ShowMessage "ËÈÊ ÇäÌÇã äÔÏ" + vbCrLf + "ÇØáÇÚÇÊ Ê˜ÑÇÑí ãí ÈÇÔÏ", True, False, "ÊÇííÏ", ""
        Case Else
    End Select
End Sub

Private Sub vsAccCode_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
''''    With vsAccCode
''''        If .Row > 0 Then
''''            Select Case .Col
''''
''''
''''                Case 3
''''
''''
''''
''''                Case 4
''''
''''
''''            End Select
''''        End If
''''    End With
''''    Set Rst = Nothing
End Sub

Private Sub vsAccCode_Click()
  If MyFormAddEditMode = EditMode Then
     With vsAccCode
        If .Row > 1 And (.Col = 5 Or .Col = 4) Then
            .Select .Row, .Col
            .EditCell
        End If
     End With
  End If
  
  If MyFormAddEditMode = ViewMode Then
    If vsAccCode.Row > 1 Then
        With vsAccCode
           txtDesc.Text = .TextMatrix(.Row, 1)
           txtTafsili.Text = .TextMatrix(.Row, 4)
           FillKol
           FillMoein
        End With
    End If
  End If
End Sub

Private Sub vsAccCode_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then
     With vsAccCode
        .Rows = .Rows + 1
        .Row = .Rows - 1
     End With
  End If
End Sub

Private Sub vsAccCode_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    For i = 0 To vsAccCode.Cols - 1
        SaveSetting strMainKey, "vsAccCode", "Col" & i, vsAccCode.ColWidth(i)
    Next
End Sub

Private Sub GetDataDetail()
    On Error Resume Next
    If cn.State = adStateClosed And clsArya.ExternalAccounting Then cn.Open AccstrConnectionString
    Dim s As String
    On Error GoTo ErrHandler
    
    Set Rst = RunStoredProcedure2RecordSet("Get_tblAcc_Sale")
    
    If Rst.EOF = True And Rst.BOF = True Then Exit Sub
    vsAccCode.Rows = 1
    Dim ii As Integer
    ii = 0
    If Rst.EOF = False Then
        Do While Not (Rst.EOF)
            vsAccCode.Rows = vsAccCode.Rows + 1
            ii = ii + 1
            vsAccCode.TextMatrix(ii, 0) = Rst!Code '
            vsAccCode.TextMatrix(ii, 1) = Rst!Description '
            vsAccCode.TextMatrix(ii, 2) = Rst!Kol 'nvcFirstName & " " & Rst!nvcSurName
            vsAccCode.TextMatrix(ii, 3) = Rst!Moein
            vsAccCode.TextMatrix(ii, 4) = Rst!Tafsili
            vsAccCode.TextMatrix(ii, 5) = Rst!Active
            vsAccCode.TextMatrix(ii, 6) = IIf(IsNull(Rst!MoeinDesc), "", Rst!MoeinDesc)
            
            Rst.MoveNext
        Loop
        vsAccCode.Cell(flexcpAlignment, 1, 1, vsAccCode.Rows - 1, vsAccCode.Cols - 1) = flexAlignRightCenter
'        vsAccCode.AutoSizeMode = flexAutoSizeColWidth
'        vsAccCode.AutoSize 1, 5
    End If
    
    If clsArya.ExternalAccounting Then
        With vsAccCode
            s = ""
            Set Rst = RunStoredProcedure2RecordSet("Get_All_tblAcc_Kols", cn)
            If LCase(clsArya.AccountSystemName) = "samar" Then
                s = .BuildComboList(Rst, Trim("KolName"), "KolId")
            Else
                s = .BuildComboList(Rst, Trim("Descs"), "Kol")
            End If
            .ColComboList(2) = s
    
        End With
    End If
'
''    With vsAccCode
''        .Col = 2
''        i = 2
''        Do While i < .Rows
''
''            ReDim Parameter(0) As Parameter
''            Parameter(0) = GenerateInputParameter("@KolID", adInteger, 4, .ValueMatrix(i, .Col))
''            Set Rst = RunParametricStoredProcedure2Rec("Get_tblAcc_Moeins_ByFK_KolID", Parameter, cn)
''            s = .BuildComboList(Rst, "MoeinName", "MoeinId")
''            .ColComboList(3) = s
''            i = i + 1
''        Loop
''    End With
    
'    vsAccCode.Cell(flexcpAlignment, 0, 0, vsAccCode.Rows - 1, vsAccCode.Cols - 1) = flexAlignCenterCenter
'    vsAccCode.Cell(flexcpAlignment, 1, 5, vsAccCode.Rows - 1, 5) = flexAlignRightCenter
    Exit Sub
ErrHandler:
    LogSave "frmAccCoding", err, "GetDataDetail"
    ShowErrorMessage
   ' Resume Next
End Sub

Private Sub vsAccCode_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
    With vsAccCode
        If MyFormAddEditMode = EditMode Then
            
            If Row > 0 And InStr(1, .TextMatrix(Row, 0), "*") = 0 Then
                .TextMatrix(Row, 0) = .TextMatrix(Row, 0) & "*"
            End If
            
        End If
    End With
End Sub

Private Sub FillKol()
    If clsArya.ExternalAccounting Then
        If cn.State = adStateClosed Then cn.Open AccstrConnectionString
        On Error GoTo ErrHandler
        
        Set Rst = RunStoredProcedure2RecordSet("Get_All_tblAcc_Kols", cn)
        cmbKol.Clear
        If Rst.EOF <> True And Rst.BOF <> True Then
            Do While Rst.EOF = False
                If LCase(clsArya.AccountSystemName) = "samar" Then
                    cmbKol.AddItem Trim(Rst.Fields("KolName"))
                    cmbKol.ItemData(cmbKol.NewIndex) = Rst!KolId
                Else
                    cmbKol.AddItem Trim(Rst.Fields("Descs"))
                    cmbKol.ItemData(cmbKol.NewIndex) = Rst!Kol
                End If
                Rst.MoveNext
            Loop
            If MyFormAddEditMode = ViewMode Then
                For i = 0 To cmbKol.ListCount - 1
                    If vsAccCode.TextMatrix(vsAccCode.Row, 2) = cmbKol.ItemData(i) Then
                      cmbKol.ListIndex = i
                      Exit For
                    End If
                Next
    '            cmbKol.ListIndex = 0
    '            vsAccCode.TextMatrix(vsAccCode.Row, 2) = cmbKol.ItemData(0)
            End If
        End If
        If Rst.State = adStateOpen Then Rst.Close
        If cn.State = adStateOpen Then cn.Close
      End If
    Exit Sub
ErrHandler:
    modgl.LogSave "frmAccCoding => ", err, "FillKol"
    ShowErrorMessage
End Sub

Private Sub FillMoein()
    If cmbKol.ListIndex = -1 Or clsArya.ExternalAccounting <> True Then cmbMoein.Clear: Exit Sub
    If cn.State = adStateClosed Then cn.Open AccstrConnectionString
    On Error GoTo ErrHandler
    
        ReDim Parameter(0) As Parameter
'        Parameter(0) = GenerateInputParameter("@KolID", adInteger, 4, vsAccCode.TextMatrix(vsAccCode.Row, 2))
        Parameter(0) = GenerateInputParameter("@KolID", adInteger, 4, cmbKol.ItemData(cmbKol.ListIndex))
        Set Rst = RunParametricStoredProcedure2Rec("Get_tblAcc_Moeins_ByFK_KolID", Parameter, cn)
        cmbMoein.Clear
        
        If Rst.EOF <> True And Rst.BOF <> True Then
        Do While Rst.EOF = False
            If LCase(clsArya.AccountSystemName) = "samar" Then
                cmbMoein.AddItem Trim(Rst.Fields("MoeinName"))
                cmbMoein.ItemData(cmbMoein.NewIndex) = Rst.Fields("MoeinId")
            Else
                cmbMoein.AddItem Trim(Rst.Fields("Descs"))
                cmbMoein.ItemData(cmbMoein.NewIndex) = Rst.Fields("M1")
            End If
            Rst.MoveNext
        Loop
        
        If MyFormAddEditMode = ViewMode Then
            For i = 0 To cmbMoein.ListCount - 1
                If vsAccCode.TextMatrix(vsAccCode.Row, 3) = cmbMoein.ItemData(i) Then
                  cmbMoein.ListIndex = i
                  Exit For
                End If
            Next
        End If
        End If
    Exit Sub
ErrHandler:
    LogSave "frmAccCoding", err, "FillMoein"
    ShowErrorMessage
End Sub

Private Sub vsAccCode_SelChange()
    If vsAccCode.Row >= 1 Then
        txtDesc.Text = vsAccCode.TextMatrix(vsAccCode.Row, 1)
        txtTafsili.Text = vsAccCode.TextMatrix(vsAccCode.Row, 4)
        FillKol
        FillMoein
    End If
End Sub


