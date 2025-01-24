VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{7AEDC602-D94C-11D1-BB7A-00E0290EA3C9}#1.0#0"; "ResizeKit.ocx"
Object = "{75D4F767-8785-11D3-93AD-0000832EF44D}#2.14#0"; "FAST2001.ocx"
Begin VB.Form frmTable 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   ClientHeight    =   8925
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10365
   Icon            =   "frmTable.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   8925
   ScaleWidth      =   10365
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   8040
      Width           =   10215
      Begin VB.CommandButton cmdCopyTable 
         Caption         =   "òÅÌ „Ì“ Â« «“ „Ì“"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6720
         TabIndex        =   6
         Top             =   240
         Width           =   1815
      End
      Begin FLWCtrls.FWNumericTextBox FWNumericTextBox1 
         Height          =   495
         Left            =   3480
         TabIndex        =   8
         Top             =   240
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   873
         Max             =   120
         Min             =   1
         Value           =   100
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "—œÌ› „‘Œ’ ‘œÂ  «"
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
         Left            =   4440
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   240
         Width           =   2175
      End
   End
   Begin VB.ComboBox cmbBranch 
      Enabled         =   0   'False
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
      TabIndex        =   3
      Top             =   360
      Width           =   2955
   End
   Begin FLWCtrls.FWLabel3D fwlblMode 
      Height          =   525
      Left            =   8880
      Top             =   0
      Width           =   1380
      _ExtentX        =   2434
      _ExtentY        =   926
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "B Yekan"
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
   Begin VSFlex7LCtl.VSFlexGrid vsTable 
      Height          =   7020
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   10185
      _cx             =   17965
      _cy             =   12382
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
      BackColorFixed  =   16761024
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
      AllowUserResizing=   0
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
   Begin RESIZEKITLibCtl.ResizeKit ResizeKit1 
      Height          =   480
      Left            =   0
      OleObjectBlob   =   "frmTable.frx":A4C2
      TabIndex        =   2
      Top             =   0
      Width           =   480
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "‰«„ ‘⁄»Â"
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
      Left            =   3120
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   360
      Width           =   975
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   " ⁄—Ì› „Ì“"
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
      Left            =   4800
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   240
      Width           =   3015
   End
End
Attribute VB_Name = "frmTable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private MyFormAddEditMode As EnumAddEditMode
Dim Parameter() As Parameter

Public Sub Add()
    With vsTable
        .Rows = .Rows + 1
        .Row = .Rows - 1
        .TextMatrix(.Row, 0) = "*"
'        .Cell(flexcpText, .Row, 6) = "0"
'        .Cell(flexcpText, .Row, 5) = "0"
        .ShowCell .Row, 0
    End With
    
    MyFormAddEditMode = EnumAddEditMode.AddMode
    
    MyFormAddEditMode = AddMode
    SetFirstToolBar
    
End Sub

Public Sub Cancel()

    FillsvsTable
'    MyFormAddEditMode = AddMode
'    SetFirstToolBar
    
End Sub
Public Sub ChangeLanguage()
    
    
End Sub
Private Sub FillsGarsonCombo()
    If rctmp.State = 1 Then rctmp.Close
    cmbGarson.Clear
    ReDim Parameter(0) As Parameter
    Parameter(0) = GenerateInputParameter("@Branch", adInteger, 4, CurrentBranch)
    Set rctmp = RunParametricStoredProcedure2Rec("Get_vw_Garson", Parameter)
    cmbGarson.AddItem ""
    cmbGarson.ItemData(0) = 0
    If Not (rctmp.EOF = True And rctmp.BOF = True) Then
        Do While rctmp.EOF <> True
        
            cmbGarson.AddItem CStr(rctmp.Fields("nvcFirstName")) & " " & CStr(rctmp.Fields("nvcSurName"))
            cmbGarson.ItemData(cmbGarson.ListCount - 1) = Val(rctmp.Fields("pPNo"))
            rctmp.MoveNext
            
        Loop
         
    End If
    cmbGarson.ListIndex = 0
    rctmp.Close

End Sub

Private Sub FillBranch()
    Dim rctmp As New ADODB.Recordset
    Dim i As Long
    cmbBranch.Clear
    Set rctmp = RunStoredProcedure2RecordSet("Get_All_Branches")
    Do While rctmp.EOF = False
        cmbBranch.AddItem rctmp!nvcBranchName
        cmbBranch.ItemData(cmbBranch.NewIndex) = rctmp!Branch
        rctmp.MoveNext
    Loop
    rctmp.Close
    For i = 0 To cmbBranch.ListCount - 1
        If CurrentBranch = cmbBranch.ItemData(i) Then
            cmbBranch.ListIndex = i
            Exit For
        End If
    Next

End Sub

Public Sub Edit()
    MyFormAddEditMode = EditMode
    SetFirstToolBar
End Sub

Public Sub ExitForm()

    Unload Me
    
End Sub

Public Sub SetFirstToolBar()

    AllButton vbOff, True
    
    mdifrm.Toolbar1.Buttons(1).Enabled = True   'Home
    mdifrm.Toolbar1.Buttons(2).Enabled = True   'PageUp
    mdifrm.Toolbar1.Buttons(3).Enabled = True   'PageDown
    mdifrm.Toolbar1.Buttons(4).Enabled = True   'End
    
    mdifrm.Toolbar1.Buttons(10).Enabled = True 'Delete
    
    mdifrm.Toolbar1.Buttons(23).Enabled = True
    mdifrm.Toolbar1.Buttons(24).Enabled = True
    mdifrm.Toolbar1.Buttons(25).Enabled = True
    mdifrm.Toolbar1.Buttons(26).Enabled = True
    mdifrm.Toolbar1.Buttons(27).Enabled = True
    
    If MyFormAddEditMode = ViewMode Then  ' View Mode
    
        mdifrm.Toolbar1.Buttons(6).Enabled = True  'Add
        mdifrm.Toolbar1.Buttons(7).Enabled = True  'Edit
        mdifrm.Toolbar1.Buttons(8).Enabled = False  'Enter
        mdifrm.Toolbar1.Buttons(9).Enabled = False   'Esc
        
    ElseIf MyFormAddEditMode = AddMode Then    'Add Mode
    
        mdifrm.Toolbar1.Buttons(6).Enabled = False  'Add
        mdifrm.Toolbar1.Buttons(7).Enabled = False  'Edit
        mdifrm.Toolbar1.Buttons(8).Enabled = True   'Enter
        mdifrm.Toolbar1.Buttons(9).Enabled = True   'Esc

    ElseIf MyFormAddEditMode = EditMode Then     'Edit
        
        mdifrm.Toolbar1.Buttons(6).Enabled = True  'Add
        mdifrm.Toolbar1.Buttons(7).Enabled = False  'Edit
        mdifrm.Toolbar1.Buttons(8).Enabled = True   'Enter
        mdifrm.Toolbar1.Buttons(9).Enabled = True   'Esc
        
    End If
    
    HeaderLabel Val(MyFormAddEditMode), fwlblMode

End Sub
Public Sub Update()
    
    If MyFormAddEditMode = EnumAddEditMode.ViewMode Then Exit Sub
    
    vstable_ValidateEdit vsTable.Row, vsTable.Col, False
    
    With vsTable
        If .Rows < 2 Then
            MyFormAddEditMode = EnumAddEditMode.ViewMode
            SetFirstToolBar
            Exit Sub
        End If
        
        For i = 1 To .Rows - 1
            .Row = i
            If InStr(.TextMatrix(i, 0), "*") > 0 Then 'new or edited records
            
               ' If Not ((Trim(.TextMatrix(i, 2)) = "" And Trim(.TextMatrix(i, 5)) = "" And .Cell(flexcpText, i, 4) = "")) Then
                    If ((Trim(.TextMatrix(i, 2)) = "" And Trim(.TextMatrix(i, 5)) = "") Or .Cell(flexcpText, i, 4) = "") Then
                        
                        Select Case clsStation.Language
                        
                            Case 0
                            
                                frmMsg.fwlblMsg.Caption = "‘„« „Ì »«Ì”  «ÿ·«⁄«  —« »ÿÊ— ò«„· Ê«—œ ‰„«ÌÌœ"
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
             '   End If
             End If
        Next i
        

        Select Case MyFormAddEditMode
        
            Case EnumAddEditMode.AddMode
            
                For i = 1 To .Rows - 1
                    
                    If .TextMatrix(i, 0) = "*" Then 'new records
                        
                        ReDim Parameter(8) As Parameter
                        Parameter(0) = GenerateInputParameter("@Name", adVarChar, 50, .TextMatrix(i, 2))
                        Parameter(1) = GenerateInputParameter("@NumberOfChair", adInteger, 4, .TextMatrix(i, 5))
                        Parameter(2) = GenerateInputParameter("@Person", adInteger, 4, IIf(.TextMatrix(i, 3) = "", 0, .TextMatrix(i, 3)))
                        Parameter(3) = GenerateInputParameter("@PartitionID", adInteger, 4, IIf(.TextMatrix(i, 4) = "", 1, .TextMatrix(i, 4)))
                        Parameter(4) = GenerateInputParameter("@Empty", adBoolean, 1, IIf(.TextMatrix(i, 6) = "-1", 1, 0))
                        Parameter(5) = GenerateInputParameter("@Reserve", adBoolean, 1, IIf(.TextMatrix(i, 7) = "-1", 1, 0))
                        Parameter(6) = GenerateInputParameter("@nvcmaxuseTime", adVarChar, 10, .TextMatrix(i, 8))
                        Parameter(7) = GenerateInputParameter("@Branch", adInteger, 4, CurrentBranch)
                        Parameter(8) = GenerateOutputParameter("@No", adInteger, 4)
                        
                        On Error GoTo ErrHandler
                        Result = RunParametricStoredProcedure("InsertTable", Parameter)
                    End If
                
                Next i
                If Result > 0 Then
                    frmMsg.fwlblMsg.Caption = "À»  «ÿ·«⁄«  ÃœÌœ »« „Ê›ﬁÌ  Å«Ì«‰ Ì«› "
                    frmMsg.fwBtn(0).ButtonType = flwButtonOk
                    frmMsg.fwBtn(0).Caption = "ﬁ»Ê·"
                    frmMsg.fwBtn(1).Visible = False
                    frmMsg.Show vbModal
                End If
               
                
            Case EnumAddEditMode.EditMode
                For i = 1 To .Rows - 1
                    
                    If InStr(.TextMatrix(i, 0), "*") > 0 Then 'Edited records
                            
                        ReDim Parameter(9) As Parameter
                        Parameter(0) = GenerateInputParameter("@Name", adVarChar, 50, .TextMatrix(i, 2))
                        Parameter(1) = GenerateInputParameter("@NumberOfChair", adInteger, 4, .TextMatrix(i, 5))
                        Parameter(2) = GenerateInputParameter("@Person", adInteger, 4, IIf(.TextMatrix(i, 3) = "", 0, .TextMatrix(i, 3)))
                        Parameter(3) = GenerateInputParameter("@PartitionID", adInteger, 4, .TextMatrix(i, 4))
                        Parameter(4) = GenerateInputParameter("@Empty", adBoolean, 1, IIf(.TextMatrix(i, 6) = -1, 1, 0))
                        Parameter(5) = GenerateInputParameter("@Reserve", adBoolean, 1, IIf(.TextMatrix(i, 7) = -1, 1, 0))
                        Parameter(6) = GenerateInputParameter("@No", adInteger, 4, .TextMatrix(i, 1))
                        Parameter(7) = GenerateInputParameter("@nvcmaxuseTime", adVarChar, 10, .TextMatrix(i, 8))
                        Parameter(8) = GenerateInputParameter("@Branch", adInteger, 4, CurrentBranch)
                        Parameter(9) = GenerateOutputParameter("@Result", adInteger, 4)
                        
                        On Error GoTo ErrHandler
                        Result = RunParametricStoredProcedure("UpdateTable", Parameter)
                            
                    End If
                                        
                Next i
                If Result > 0 Then
                    frmMsg.fwlblMsg.Caption = " €ÌÌ— «ÿ·«⁄«  ÃœÌœ »« „Ê›ﬁÌ  Å«Ì«‰ Ì«› "
                    frmMsg.fwBtn(0).ButtonType = flwButtonOk
                    frmMsg.fwBtn(0).Caption = "ﬁ»Ê·"
                    frmMsg.fwBtn(1).Visible = False
                    frmMsg.Show vbModal
                End If
            
            End Select
            
        FillsvsTable
        
    End With
Exit Sub
ErrHandler:
    Select Case err.Number
        Case -2147217873
            frmMsg.fwlblMsg.Caption = "À»  «‰Ã«„ ‰‘œ" + vbCrLf + "«ÿ·«⁄«   ò—«—Ì „Ì »«‘œ"
            frmMsg.fwBtn(0).ButtonType = flwButtonOk
            frmMsg.fwBtn(0).Caption = "ﬁ»Ê·"
            frmMsg.fwBtn(1).Visible = False
            frmMsg.Show vbModal
        Case Else
            MsgBox err.Description
    End Select
    
End Sub

Private Sub cmbPerson_Change()

End Sub



Private Sub cmdCopyTable_Click()
 With vsTable
    If .Row < 1 Then ShowDisMessage "¬Œ—Ì‰ —œÌ› „Ì“ „Ê—œ ‰Ÿ— —« «‰ Œ«» ò‰Ìœ", 1500: Exit Sub
    ShowMessage "¬Ì« »—«Ì òÅÌ „Ì“Â«  « ‘„«—Â „Ê—œ ‰Ÿ— „ÿ„∆‰ Â” Ìœø ", True, True, "»·Ì", "ŒÌ—"
    Dim Result As Long
    If modgl.mvarMsgIdx = vbYes Then
        
        ReDim Parameter(9) As Parameter
        Parameter(0) = GenerateInputParameter("@FromTableNo", adInteger, 4, Val(.TextMatrix(.Row, 2)))
        Parameter(1) = GenerateInputParameter("@NumberOfChair", adInteger, 4, 4)
        Parameter(2) = GenerateInputParameter("@Person", adInteger, 4, IIf(.TextMatrix(.Row, 3) = "", 0, .TextMatrix(.Row, 3)))
        Parameter(3) = GenerateInputParameter("@PartitionID", adInteger, 4, IIf(.TextMatrix(.Row, 4) = "", 1, .TextMatrix(.Row, 4)))
        Parameter(4) = GenerateInputParameter("@Empty", adBoolean, 1, 1)
        Parameter(5) = GenerateInputParameter("@Reserve", adBoolean, 1, 0)
        Parameter(6) = GenerateInputParameter("@nvcmaxuseTime", adVarChar, 10, .TextMatrix(.Row, 8))
        Parameter(7) = GenerateInputParameter("@Branch", adInteger, 4, CurrentBranch)
        Parameter(8) = GenerateInputParameter("@ToTableNo", adInteger, 4, FWNumericTextBox1.Value)
        Parameter(9) = GenerateOutputParameter("@intStatus", adInteger, 4)
        
        Result = RunParametricStoredProcedure("Copy_tTables", Parameter)
        If Result > 0 Then
            ShowDisMessage "«÷«›Â ò—œ‰  „Ì“Â« «‰Ã«„ ‘œ", 1500
            FillsvsTable
        Else
            ShowDisMessage " œ— À»  «÷«›Â ò—œ‰ „Ì“Â« „‘ò· ÊÃÊœ œ«—œ ", 1500
        End If
        
    End If
End With
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

    If ClsFormAccess.frmTable = False Then
        Unload Me
        Exit Sub
    End If
    
    If intVersion = Min Then
        ShowDisMessage "«” ›«œÂ «“ „Ì“ Ê ê«—”Ê‰ œ— ‰”ŒÂ Â«Ì „⁄„Ê·Ì Ê »«·« — «„ﬂ«‰ Å–Ì— «” ", 1500
        Unload Me
        Exit Sub
    End If
    
    CenterTop Me
    
    VarActForm = Me.Name
    
    MyFormAddEditMode = AddMode
    SetFirstToolBar
    With vsTable
    
        .Cols = 9
        .TextMatrix(0, 0) = "—œÌ›"
        .TextMatrix(0, 1) = "òœ"
        .TextMatrix(0, 2) = "„Ì“"
        .TextMatrix(0, 3) = "„”∆Ê·"
        .TextMatrix(0, 4) = "»Œ‘"
        .TextMatrix(0, 5) = "  ’‰œ·Ì  "
        .TextMatrix(0, 6) = " Œ«·Ì "
        .TextMatrix(0, 7) = "—“—Ê"
        .TextMatrix(0, 8) = "Õœ«òÀ— “„«‰ «” ›«œÂ"
            
        
        .ColDataType(6) = flexDTBoolean
        .ColDataType(7) = flexDTBoolean
        .ColSort(2) = flexSortNumericAscending + flexSortNumericDescending
        .FocusRect = flexFocusHeavy
        .ColHidden(1) = True
        .AutoSizeMode = flexAutoSizeColWidth
        .AutoSize 0, .Cols - 1
        .AutoSearch = flexSearchFromCursor
        
        .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter
     '   .Cell(flexcpAlignment, 0, 5, 0, 5) = flexAlignRightCenter
        
        Dim Rst As New ADODB.Recordset
        
        ReDim Parameter(1) As Parameter
        Parameter(0) = GenerateInputParameter("@Job", adInteger, 4, EnumIncharge.Garson)
        Parameter(1) = GenerateInputParameter("@Branch", adInteger, 4, CurrentBranch)
        s = ""
        Set Rst = RunParametricStoredProcedure2Rec("Get_vw_Per_BY_Job", Parameter)

        s = .BuildComboList(Rst, "nvcSurName", "pPno")   '"nvcFirstName" & " " &
            .ColComboList(3) = s
    
        FillBranch
        
        ReDim Parameter(1) As Parameter
        Parameter(0) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
        Parameter(1) = GenerateInputParameter("@Branch", adInteger, 4, CurrentBranch)
        Set Rst = RunParametricStoredProcedure2Rec("Get_All_tPartitions", Parameter)
        s = ""
        s = .BuildComboList(Rst, "PartitionDescription", "PartitionID")
        .ColComboList(4) = s
    
        
    End With
    FillsvsTable
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
Public Sub FillsvsTable() 'it fills the grid using
    
    MyFormAddEditMode = EnumAddEditMode.ViewMode 'VIEW Mode
    SetFirstToolBar
    
    vsTable.Rows = 1
    
    Dim Rst As New ADODB.Recordset
    
    
    With vsTable
    
    ReDim Parameter(0) As Parameter
    Parameter(0) = GenerateInputParameter("@Branch", adInteger, 4, CurrentBranch)
    Set Rst = RunParametricStoredProcedure2Rec("Get_Tables", Parameter)
    
    If (Rst.EOF = True And Rst.BOF = True) Then Exit Sub
    
        i = 1
        
        While Rst.EOF = False
            .Rows = .Rows + 1
            .TextMatrix(i, 0) = i
            .TextMatrix(i, 1) = Rst!No
            .TextMatrix(i, 2) = Rst!Name
            .TextMatrix(i, 3) = IIf(IsNull(Rst!Person), "", Rst!Person)
            .TextMatrix(i, 4) = Rst!PartitionID
            .TextMatrix(i, 5) = Rst!NumberOfChair
            .TextMatrix(i, 6) = IIf(Rst!Empty = True, -1, 0)
            .TextMatrix(i, 7) = IIf(Rst!Reserve = True, -1, 0)
            .TextMatrix(i, 8) = IIf(IsNull(Rst!nvcMaxUseTime), "", Rst!nvcMaxUseTime)
            
            
           
          '  .Cell(flexcpText, i, 9) = Rst.Fields("unit").Value
            i = i + 1
            Rst.MoveNext
            
        Wend
        Set Rst = Nothing
        
        .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter
        .Cell(flexcpAlignment, 1, 4, .Rows - 1, .Cols - 1) = flexAlignCenterCenter
        .AutoSizeMode = flexAutoSizeColWidth
        .AutoSize 0, .Cols - 1
        
    End With
        
End Sub

Private Sub ResizeKit1_ExitResize(ByVal XScale As Double, ByVal YScale As Double)
    If formloadFlag = True Then
        SaveSetting strMainKey, Me.Name, "Height", Me.Height
        SaveSetting strMainKey, Me.Name, "Width", Me.Width
    End If

End Sub

Private Sub vsTable_Click()
    
    With vsTable
        If .Col = 6 Or .Col = 7 Then Exit Sub
        If (MyFormAddEditMode = EnumAddEditMode.EditMode) Then
            .Select .Row, .Col
            .EditCell
        End If
    
    End With
End Sub
Private Sub vsTable_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    With vsTable
        
        If MyFormAddEditMode = EditMode Then
            
            If Row > 0 And InStr(1, .TextMatrix(Row, 0), "*") = 0 Then
                .TextMatrix(Row, 0) = .TextMatrix(Row, 0) & "*"
            End If
            
        End If
        
    End With
    
End Sub


Private Sub vsTable_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    With vsTable
        If .Row > 0 And InStr(1, .TextMatrix(.Row, 0), "*") = 0 And MyFormAddEditMode = EditMode Then
            .TextMatrix(.Row, 0) = .TextMatrix(.Row, 0) & "*"
        End If
        If (.TextMatrix(.Row, 0) = "*" Or MyFormAddEditMode = EnumAddEditMode.EditMode) And .Row > 0 And .Col > 1 Then
            .Select .Row, .Col
            .EditCell
        End If
    
    End With
    
End Sub


Private Sub vstable_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
''''    With vsTable
''''        .Row = Row
''''        .Col = Col
''''        If MyFormAddEditMode = EditMode Then
''''
''''            If Row > 0 And InStr(1, .TextMatrix(Row, 0), "*") = 0 Then
''''                .TextMatrix(Row, 0) = .TextMatrix(Row, 0) & "*"
''''            End If
''''
''''        End If
''''    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    VarActForm = ""
    If PosConnection.State = adStateOpen Then PosConnection.Close: Set PosConnection = Nothing

    SaveSetting strMainKey, Me.Name, "Left", Me.Left
    SaveSetting strMainKey, Me.Name, "Top", Me.Top

End Sub


Private Sub txtNumberOfChairs_KeyPress(KeyAscii As Integer)
    If IsNumeric(Chr(KeyAscii)) <> True And KeyAscii <> 8 Then
        KeyAscii = 0
    End If
End Sub

