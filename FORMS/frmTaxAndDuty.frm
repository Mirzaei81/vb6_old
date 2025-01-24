VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{7AEDC602-D94C-11D1-BB7A-00E0290EA3C9}#1.0#0"; "ResizeKit.ocx"
Object = "{75D4F767-8785-11D3-93AD-0000832EF44D}#2.14#0"; "FAST2001.ocx"
Begin VB.Form frmTaxAndDuty 
   ClientHeight    =   4755
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8130
   Icon            =   "frmTaxAndDuty.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4755
   ScaleWidth      =   8130
   Begin FLWCtrls.FWLabel3D fwlblMode 
      Height          =   495
      Left            =   6720
      Top             =   0
      Width           =   1380
      _ExtentX        =   2434
      _ExtentY        =   873
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   14.25
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
      OleObjectBlob   =   "frmTaxAndDuty.frx":A4C2
      TabIndex        =   0
      Top             =   0
      Width           =   480
   End
   Begin FLWCtrls.FWLabel fwlblPartition 
      Height          =   495
      Left            =   0
      Top             =   0
      Width           =   8115
      _ExtentX        =   14314
      _ExtentY        =   873
      Enabled         =   -1  'True
      Caption         =   "›—„  ⁄—Ì› ⁄Ê«—÷ Ê „«·Ì« "
      FirstColor      =   9412754
      SecondColor     =   14215660
      Angle           =   0
      ForeColor       =   7362318
      BackColor       =   12640511
      FontName        =   "B Homa"
      FontSize        =   15.75
      Alignment       =   2
      Picture         =   "frmTaxAndDuty.frx":A548
   End
   Begin VSFlex7LCtl.VSFlexGrid vsTaxAndDuty 
      Height          =   4155
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   7995
      _cx             =   14102
      _cy             =   7329
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
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   12648384
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
      AllowBigSelection=   0   'False
      AllowUserResizing=   2
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   6
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
End
Attribute VB_Name = "frmTaxAndDuty"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private MyFormAddEditMode As EnumAddEditMode
 
Dim Parameter() As Parameter

'===================================
Const GRID_COLUMNS_COUNT As Long = 5
Const IdxColRow As Integer = 0
Const IdxColName As Integer = 1
Const IdxColAmountPercent As Integer = 2
Const IdxColActive As Integer = 3
Const IdxColIntServiceNo As Integer = 4
'===================================
Public Sub SetFirstToolBar()
    
    AllButton vbOff, True
    
    mdifrm.Toolbar1.Buttons(1).Enabled = False   'Home
    mdifrm.Toolbar1.Buttons(2).Enabled = False   'PageUp
    mdifrm.Toolbar1.Buttons(3).Enabled = False   'PageDown
    mdifrm.Toolbar1.Buttons(4).Enabled = False   'End
        
    mdifrm.Toolbar1.Buttons(23).Enabled = True
    mdifrm.Toolbar1.Buttons(24).Enabled = True
    mdifrm.Toolbar1.Buttons(25).Enabled = True
    mdifrm.Toolbar1.Buttons(26).Enabled = True
    mdifrm.Toolbar1.Buttons(27).Enabled = True
    
    If MyFormAddEditMode = ViewMode Then  ' View Mode
    
        mdifrm.Toolbar1.Buttons(10).Enabled = True 'Delete
        mdifrm.Toolbar1.Buttons(6).Enabled = True  'Add
        mdifrm.Toolbar1.Buttons(7).Enabled = True  'Edit
        mdifrm.Toolbar1.Buttons(8).Enabled = False  'Enter
        mdifrm.Toolbar1.Buttons(9).Enabled = False   'Esc
        mdifrm.Toolbar1.Buttons(10).Enabled = True 'Delete
        
    ElseIf MyFormAddEditMode = AddMode Then    'Add Mode
    
        mdifrm.Toolbar1.Buttons(6).Enabled = False  'Add
        mdifrm.Toolbar1.Buttons(7).Enabled = False  'Edit
        mdifrm.Toolbar1.Buttons(8).Enabled = True   'Enter
        mdifrm.Toolbar1.Buttons(9).Enabled = True   'Esc
        mdifrm.Toolbar1.Buttons(10).Enabled = False 'Delete

    ElseIf MyFormAddEditMode = EditMode Then     'Edit
        
        mdifrm.Toolbar1.Buttons(6).Enabled = True  'Add
        mdifrm.Toolbar1.Buttons(7).Enabled = False  'Edit
        mdifrm.Toolbar1.Buttons(8).Enabled = True   'Enter
        mdifrm.Toolbar1.Buttons(9).Enabled = True   'Esc
        mdifrm.Toolbar1.Buttons(10).Enabled = False 'Delete
        
    End If
    
    HeaderLabel Val(MyFormAddEditMode), fwlblMode
    
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

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call PresetScreenSaver
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call PresetScreenSaver
End Sub

Private Sub Form_Load()
    If ClsFormAccess.frmTaxAndDuty = False Then
        Unload Me
        Exit Sub
    End If
    
    If intVersion = Min Then
        ShowDisMessage "«” ›«œÂ «“ „«·Ì«  Ê ⁄Ê«—÷ œ— ‰”ŒÂ „ Ê”ÿ Ê »«·« — «„ﬂ«‰ Å–Ì— «” ", 1500
        Unload Me
        Exit Sub
    End If
'    CenterTop Me

    VarActForm = Me.Name
    
    With vsTaxAndDuty
        .Cols = GRID_COLUMNS_COUNT
        .TextMatrix(0, IdxColRow) = "—œÌ›"
        .TextMatrix(0, IdxColName) = "‰«„"
        .TextMatrix(0, IdxColAmountPercent) = "„ﬁœ«—- œ—’œ"
        .TextMatrix(0, IdxColActive) = "›⁄«·"
        
        .ColAlignment(-1) = flexAlignCenterCenter
        .ColWidth(IdxColRow) = 500
        .ColWidth(IdxColName) = 2500
        .ColWidth(IdxColAmountPercent) = 2500
        .ColWidth(IdxColActive) = 2000
        
        .ColDataType(IdxColActive) = flexDTBoolean
        
        .ColAlignment(IdxColName) = flexAlignCenterCenter
        .ColAlignment(IdxColAmountPercent) = flexAlignCenterCenter
        .ColAlignment(IdxColActive) = flexAlignCenterCenter
        
        .ColHidden(IdxColIntServiceNo) = True
        
    End With
    MyFormAddEditMode = ViewMode
    DefaultSetting
    SetFirstToolBar
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
    SaveSetting strMainKey, Me.Name, "Left", Me.Left
    SaveSetting strMainKey, Me.Name, "Top", Me.Top
    If PosConnection.State = adStateOpen Then PosConnection.Close: Set PosConnection = Nothing
    
End Sub

Public Sub ExitForm()
    Unload Me
End Sub

Public Sub Edit()
    MyFormAddEditMode = EditMode
    SetFirstToolBar
End Sub


''''Public Sub Delete()
''''
''''    If vsInventory.Rows < 2 Then Exit Sub
''''
''''    On Error GoTo ErrHandler
''''    ReDim Parameter(0) As Parameter
''''    Parameter(0) = GenerateInputParameter("@intId", adInteger, 4, Val(vsInventory.TextMatrix(vsInventory.Row, 1)))
''''    RunParametricStoredProcedure "Delete_Inventory", Parameter
''''
''''    frmMsg.fwlblMsg.Caption = "«‰»«— „Ê—œ ‰Ÿ— »« „Ê›ﬁÌ  Õ–› ‘œ"
''''    frmMsg.fwBtn(0).Visible = False
''''    frmMsg.fwBtn(1).ButtonType = flwButtonOk
''''    frmMsg.fwBtn(1).Caption = "ﬁ»Ê·"
''''    frmMsg.Show vbModal
''''
''''    DefaultSetting
''''Exit Sub
''''
''''ErrHandler:
''''If Err.Number = -2147217873 Then
''''
''''    frmMsg.fwlblMsg.Caption = "„ «”›«‰Â ‘„« ﬁ«œ— »Â Õ–› ‰„Ì »«‘Ìœ"
''''    frmMsg.fwBtn(0).Visible = False
''''    frmMsg.fwBtn(1).ButtonType = flwButtonOk
''''    frmMsg.fwBtn(1).Caption = "ﬁ»Ê·"
''''    frmMsg.Show vbModal
''''Else
''''    MsgBox Err.Description
''''End If
''''
''''End Sub

Public Sub DefaultSetting()
    
'    If cmbBranch.ListIndex = -1 Then Exit Sub
    Dim Rst As New ADODB.Recordset
    
 
    Set Rst = RunStoredProcedure2RecordSet("GetTaxAndDuty")
    
    With vsTaxAndDuty
        .Rows = 1
        If Not (Rst.EOF = True And Rst.BOF = True) Then
            While Rst.EOF <> True
                .Rows = .Rows + 1

                .TextMatrix(.Rows - 1, IdxColRow) = .Rows - 1
                .TextMatrix(.Rows - 1, IdxColName) = Rst!Fa_Name
                .TextMatrix(.Rows - 1, IdxColAmountPercent) = Rst!AmountPercent
                .TextMatrix(.Rows - 1, IdxColActive) = Rst!Active
                .TextMatrix(.Rows - 1, IdxColIntServiceNo) = Rst!intServiceNo
                Rst.MoveNext
            Wend
        End If
    
    End With
    
    
    If Rst.State = 1 Then Rst.Close
     
''''    Dim obj As Object
''''    For Each obj In Me
''''        If TypeOf obj Is TextBox Then
''''            obj.Text = ""
''''            obj.Tag = 0
''''        ElseIf TypeOf obj Is ComboBox Then
''''            obj.ListIndex = 0
''''        ElseIf TypeOf obj Is CheckBox Then
''''            obj.Value = False
''''        End If
''''    Next obj
    
    Set Rst = Nothing
    
End Sub

Public Sub Add()
    MyFormAddEditMode = AddMode
    DefaultSetting
    SetFirstToolBar
End Sub

Public Sub Cancel()
    MyFormAddEditMode = ViewMode
    SetFirstToolBar
    DefaultSetting
    
End Sub

Public Sub ChangeLanguage()
    Select Case clsStation.Language
        Case Farsi
        
        Case English
    End Select
End Sub

Public Sub Update()
    ReDim Parameter(4) As Parameter
    Dim Result As Integer
    Dim Obj As Object
    Dim i As Integer
    
'''    If Trim(txtInventory.Text) = "" Or InStr(txtInventory.Text, "'") <> 0 Then
'''
'''            frmMsg.fwlblMsg.Caption = "·ÿ›« Ìò ‰«„ „⁄ »— »—«Ì «‰»«— Ê«—œ ‰„«ÌÌœ"
'''            frmMsg.fwBtn(0).ButtonType = flwButtonOk
'''            frmMsg.fwBtn(0).Caption = "ﬁ»Ê·"
'''            frmMsg.Show vbModal
'''
'''            txtInventory.SetFocus
'''
'''            Exit Sub
'''
'''    End If
    
''''    For i = 1 To vsInventory.Rows - 1
''''
''''        If vsInventory.TextMatrix(i, 4) = 1 Then
''''            CentralBranch = True
''''            CentralBranchCode = vsInventory.TextMatrix(i, 2)
''''        End If
''''    Next i
    
'    Select Case MyFormAddEditMode
    
'        Case AddMode
'
'''            If cboInventoryType.ItemData(cboInventoryType.ListIndex) = 1 And CentralBranch = True Then
'''                frmMsg.fwlblMsg.Caption = "‘„« ‰„Ì  Ê«‰Ìœ »Ì‘ «“ Ìò ‘⁄»Â „—ò“Ì œ«‘ Â »«‘Ìœ"
'''                frmMsg.Fwbtn(0).ButtonType = flwButtonOk
'''                frmMsg.Fwbtn(0).Caption = "ﬁ»Ê·"
'''                frmMsg.Show vbModal
'''                Exit Sub
'''            End If
            
'            Parameter(0) = GenerateInputParameter("@Fa_name", adVarWChar, 50, Trim(txtInventory.Text))
'            Parameter(1) = GenerateInputParameter("@AmountPercent", adInteger, 4, Val(chkActive.Value))
'            Parameter(2) = GenerateInputParameter("@Branch", adInteger, 4, cmbBranch.ItemData(cmbBranch.ListIndex))
'            Parameter(2) = GenerateOutputParameter("@Active", adBoolean, 1)
'            Result = RunParametricStoredProcedure("Insert_tinventory", Parameter )
'
'            If Parameter(3).Value <> -1 Then
'                txtInventory.Tag = Parameter(3).Value
'                frmMsg.fwlblMsg.Caption = "À»  «ÿ·«⁄«  ÃœÌœ »« „Ê›ﬁÌ  Å«Ì«‰ Ì«› "
'                frmMsg.fwBtn(0).ButtonType = flwButtonOk
'                frmMsg.fwBtn(0).Caption = "ﬁ»Ê·"
'                frmMsg.fwBtn(1).Visible = False
'                frmMsg.Show vbModal
'
'                Add
'            Else
'                frmMsg.fwlblMsg.Caption = "«ÿ·«⁄«  ÃœÌœ À»  ‰‘œ. ·ÿ›« œÊ»«—Â ”⁄Ì ‰„«ÌÌœ"
'                frmMsg.fwBtn(0).ButtonType = flwButtonOk
'                frmMsg.fwBtn(0).Caption = "ﬁ»Ê·"
'                frmMsg.fwBtn(1).Visible = False
'                frmMsg.Show vbModal
'            End If
            
'           Case EditMode
        
''''            If CentralBranch = True And CentralBranchCode <> txtInventory.Tag Then
''''                frmMsg.fwlblMsg.Caption = "‘„« ‰„Ì  Ê«‰Ìœ »Ì‘ «“ Ìò ‘⁄»Â „—ò“Ì œ«‘ Â »«‘Ìœ"
''''                frmMsg.Fwbtn(0).ButtonType = flwButtonOk
''''                frmMsg.Fwbtn(0).Caption = "ﬁ»Ê·"
''''                frmMsg.Show vbModal
''''                Exit Sub
''''            End If
        
'           ReDim Parameter(3) As Parameter
            With vsTaxAndDuty
                For i = 1 To .Rows - 1
                   If InStr(.TextMatrix(i, IdxColRow), "*") > 0 Then 'Edited records
                       Parameter(0) = GenerateInputParameter("@Fa_Name", adVarWChar, 50, .TextMatrix(i, IdxColName))
                       Parameter(1) = GenerateInputParameter("@AmountPercent", adDouble, 8, Val(.TextMatrix(i, IdxColAmountPercent)))
                       Parameter(2) = GenerateInputParameter("@Active", adBoolean, 1, .TextMatrix(i, IdxColActive))  ''Val(.TextMatrix(i, 3)))
                       Parameter(3) = GenerateInputParameter("@intServiceNo", adInteger, 4, Val(.TextMatrix(i, IdxColIntServiceNo)))
                       Parameter(4) = GenerateOutputParameter("@Result", adInteger, 4)
                
                       Result = RunParametricStoredProcedure("Update_tAdditionalService", Parameter)
                
                       If Result <> -1 Then
                           ShowMessage " €ÌÌ— «ÿ·«⁄«  " & .TextMatrix(i, 1) & " »« „Ê›ﬁÌ  Å«Ì«‰ Ì«› ", True, False, " «ÌÌœ", ""
                       
                       Else
                           ShowMessage "„ «”›«‰Â «ÿ·«⁄«  " & .TextMatrix(i, 1) & "  €ÌÌ— ‰Ì«› . ·ÿ›« œÊ»«—Â ”⁄Ì ‰„«ÌÌœ", True, False, " «ÌÌœ", ""
                       End If
                   End If
                Next i
                
                Cancel
            End With
               
'            End Select
            
'        FillvsGood
'    End Select
End Sub

Private Sub ResizeKit1_ExitResize(ByVal XScale As Double, ByVal YScale As Double)
    If formloadFlag = True Then
        SaveSetting strMainKey, Me.Name, "Height", Me.Height
        SaveSetting strMainKey, Me.Name, "Width", Me.Width
    End If
End Sub

''''Private Sub vsTaxAndDuty_Click()
'''' If vsTaxAndDuty.Row = 1 Then Exit Sub
''''    With vsTaxAndDuty
''''        If .Row = 0 Then Exit Sub
''''        MyFormAddEditMode = ViewMode
''''        SetFirstToolbar
''''    End With
'''''''
'''''''    MyFormAddEditMode = EnumAddEditMode.EditMode
'''''''    SetFirstToolbar
'''''''    With vsGood
'''''''        .Editable = flexEDKbdMouse
'''''''    End With
''''End Sub


Private Sub vsTaxAndDuty_KeyDown(KeyCode As Integer, Shift As Integer)
    With vsTaxAndDuty
        If (.TextMatrix(.Row, IdxColRow) = "*" Or MyFormAddEditMode = EnumAddEditMode.EditMode) And .Col > IdxColName Then
            If KeyCode = 38 Or KeyCode = 39 Or KeyCode = 40 Then
            ElseIf KeyCode = 37 Then
            ElseIf KeyCode = 32 And (.Col = IdxColAmountPercent Or Col = IdxColActive) Then
                .Select .Row, .Col
                .EditCell
                If .Row > 0 And InStr(1, .TextMatrix(.Row, IdxColRow), "*") = 0 Then
                    .TextMatrix(.Row, IdxColRow) = .TextMatrix(.Row, IdxColRow) & "*"
                End If
            Else
                .Select .Row, .Col
                .EditCell
            End If
        End If
    End With
End Sub

Private Sub vsTaxAndDuty_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    With vsTaxAndDuty
        If KeyAscii = 39 Then KeyAscii = 0
            
        If MyFormAddEditMode = EditMode Then
            If Row > 0 And InStr(1, .TextMatrix(Row, IdxColRow), "*") = 0 Then
                .TextMatrix(Row, IdxColRow) = .TextMatrix(Row, IdxColRow) & "*"
            End If
        End If
    End With
End Sub

Private Sub vsTaxAndDuty_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    With vsTaxAndDuty
        If (MyFormAddEditMode = EnumAddEditMode.EditMode) And .Col > IdxColName Then
            .Select .Row, .Col
            If .Col <> IdxColAmountPercent And .Col <> IdxColActive Then Exit Sub
            If .Row > 0 And InStr(1, .TextMatrix(.Row, IdxColRow), "*") = 0 Then
                .TextMatrix(.Row, IdxColRow) = .TextMatrix(.Row, IdxColRow) & "*"
            End If
            .EditCell
        End If
    End With
End Sub

Private Sub vsTaxAndDuty_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> IdxColAmountPercent And Col <> IdxColActive Then Exit Sub
    With vsTaxAndDuty
        .Row = Row
        .Col = Col
    End With
End Sub

