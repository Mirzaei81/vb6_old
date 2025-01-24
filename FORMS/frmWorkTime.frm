VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{7AEDC602-D94C-11D1-BB7A-00E0290EA3C9}#1.0#0"; "ResizeKit.ocx"
Object = "{75D4F767-8785-11D3-93AD-0000832EF44D}#2.14#0"; "FAST2001.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmWorkTime 
   ClientHeight    =   5385
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8415
   Icon            =   "frmWorkTime.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   5385
   ScaleWidth      =   8415
   Begin RESIZEKITLibCtl.ResizeKit ResizeKit1 
      Height          =   480
      Left            =   480
      OleObjectBlob   =   "frmWorkTime.frx":A4C2
      TabIndex        =   9
      Top             =   240
      Width           =   480
   End
   Begin VB.TextBox txtDescription 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Nazanin"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   4680
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   1680
      Width           =   2265
   End
   Begin VB.Frame Frame1 
      Height          =   1755
      Left            =   360
      TabIndex        =   1
      Top             =   630
      Width           =   3675
      Begin MSComCtl2.DTPicker DTPStart 
         Height          =   345
         Left            =   420
         TabIndex        =   2
         Top             =   390
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   609
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   81592321
         CurrentDate     =   38330
      End
      Begin MSComCtl2.DTPicker DTPEnd 
         Height          =   345
         Left            =   420
         TabIndex        =   3
         Top             =   1050
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   609
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   81592321
         CurrentDate     =   38330
      End
      Begin VB.Label lblStartTime 
         Alignment       =   1  'Right Justify
         Caption         =   "”«⁄  ‘—Ê⁄ ﬂ«—"
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   2040
         TabIndex        =   5
         Top             =   360
         Width           =   1485
      End
      Begin VB.Label lblStopTime 
         Alignment       =   1  'Right Justify
         Caption         =   "”«⁄  « „«„ ﬂ«—"
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1920
         TabIndex        =   4
         Top             =   1050
         Width           =   1605
      End
   End
   Begin VSFlex7LCtl.VSFlexGrid vsShift 
      Height          =   2895
      Left            =   360
      TabIndex        =   0
      Top             =   2460
      Width           =   7965
      _cx             =   14049
      _cy             =   5106
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
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmWorkTime.frx":A548
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
      ExplorerBar     =   0
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
      Height          =   570
      Left            =   6720
      Top             =   0
      Width           =   1680
      _ExtentX        =   2963
      _ExtentY        =   1005
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
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   " ‰ŸÌ„ ”«⁄  ò«—Ì"
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
      Left            =   2880
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   0
      Width           =   3015
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "‰«„ ‘Ì› "
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
      Left            =   6960
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   1800
      Width           =   1245
   End
End
Attribute VB_Name = "frmWorkTime"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private MyFormAddEditMode As EnumAddEditMode
Dim Parameter() As Parameter

Public Sub Add()
    cmbBranch.Enabled = True
    MyFormAddEditMode = AddMode 'Add
    SetFirstToolBar
    
End Sub
Public Sub AfterCancel()
    FillVsShift
End Sub

Public Sub Cancel()
    MyFormAddEditMode = ViewMode 'View
    SetFirstToolBar
End Sub

Public Sub Delete()

    If vsShift.Rows < 2 Then Exit Sub

    If MyFormAddEditMode <> 0 Then
        Cancel
    End If
    If txtDescription.Tag = "" Then Exit Sub
    On Error GoTo ErrHandler
    ReDim Parameter(0) As Parameter
    Parameter(0) = GenerateInputParameter("@Code", adInteger, 4, txtDescription.Tag)
    RunParametricStoredProcedure "Delete_tShift_By_Code", Parameter
    
    frmMsg.fwlblMsg.Caption = "‘Ì›  ﬂ«—Ì »« „Ê›ﬁÌ  Õ–› ‘œ"
    frmMsg.fwBtn(0).Visible = False
    frmMsg.fwBtn(1).ButtonType = flwButtonOk
    frmMsg.fwBtn(1).Caption = "ﬁ»Ê·"
    frmMsg.Show vbModal
    
    FillVsShift
Exit Sub
    
ErrHandler:
If err.Number = -2147217873 Then

    frmMsg.fwlblMsg.Caption = "„ «”›«‰Â ‘„« ﬁ«œ— »Â Õ–› «Ì‰ ‘Ì›  ‰„Ì »«‘Ìœ"
    frmMsg.fwBtn(0).Visible = False
    frmMsg.fwBtn(1).ButtonType = flwButtonOk
    frmMsg.fwBtn(1).Caption = "ﬁ»Ê·"
    frmMsg.Show vbModal
End If
    
End Sub

Public Sub ExitForm()
    Unload Me
End Sub

Public Sub Edit()
'    cmbBranch.Enabled = False
    If vsShift.Rows > 1 Then
        MyFormAddEditMode = EditMode 'Edit
        SetFirstToolBar
    End If
End Sub
Private Sub cmbBranch_Click()
    FillVsShift
End Sub
'Private Sub FillBranch()
'
'    cmbBranch.Clear
'    Set rctmp = RunStoredProcedure2RecordSet("Get_All_Branches")
'    Do While rctmp.EOF = False
'        cmbBranch.AddItem rctmp!nvcBranchName
'        cmbBranch.ItemData(cmbBranch.NewIndex) = rctmp!Branch
'        rctmp.MoveNext
'    Loop
'    rctmp.Close
'    For i = 0 To cmbBranch.ListCount - 1
'        cmbBranch.ListIndex = i
'        If CurrentBranch = cmbBranch.ItemData(cmbBranch.ListIndex) Then
'            Exit For
'        End If
'    Next
'
'End Sub

Public Sub FillVsShift()
    Dim Rst As New ADODB.Recordset
    
    With vsShift
        .Rows = 1
'        ReDim Parameter(0) As Parameter
'        Parameter(0) = GenerateInputParameter("@Branch", adInteger, 4, cmbBranch.ItemData(cmbBranch.ListIndex))
'        Set Rst = RunParametricStoredProcedure2Rec("Get_All_tShift", Parameter)
        Set Rst = RunStoredProcedure2RecordSet("Get_All_tShift")
        If Not (Rst.EOF = True And Rst.BOF = True) Then
            While Rst.EOF <> True
                .Rows = .Rows + 1
                .TextMatrix(.Rows - 1, 0) = .Rows - 1
                .TextMatrix(.Rows - 1, 1) = Rst.Fields("Code").Value
                .TextMatrix(.Rows - 1, 2) = Rst.Fields("Description").Value
                .TextMatrix(.Rows - 1, 3) = FormatDateTime(Rst.Fields("StartTime").Value, vbShortTime)
                .TextMatrix(.Rows - 1, 4) = FormatDateTime(Rst.Fields("EndTime").Value, vbShortTime)
                Rst.MoveNext
            Wend
            .ColAlignment(0) = flexAlignCenterCenter
            .ColAlignment(1) = flexAlignCenterCenter
            .ColAlignment(2) = flexAlignRightCenter
            .ColAlignment(3) = flexAlignCenterCenter
            .ColAlignment(4) = flexAlignCenterCenter
            
            .ColHidden(1) = True
            
            .AutoSizeMode = flexAutoSizeColWidth
            .AutoSize 0, .Cols - 1
            .Row = 0
            .Row = 1
        End If
        
    End With
End Sub

Public Sub SetFirstToolBar()
    
    AllButton vbOff, True
   
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

Public Sub Update()
    
    Select Case MyFormAddEditMode
        Case AddMode 'add
            
            If ValidRange = True Then
                ReDim Parameter(4) As Parameter
                Parameter(0) = GenerateInputParameter("@Description", adVarWChar, 50, txtDescription.Text)
                Parameter(1) = GenerateInputParameter("@StartTime", adDBTime, 8, FormatDateTime(DTPStart.Value, vbShortTime))
                Parameter(2) = GenerateInputParameter("@EndTime", adDBTime, 8, FormatDateTime(DTPEnd.Value, vbShortTime))
                Parameter(3) = GenerateInputParameter("@Branch", adInteger, 4, CurrentBranch)
                Parameter(4) = GenerateOutputParameter("@Result", adInteger, 4)
                
                If RunParametricStoredProcedure("InsertShift", Parameter) <> -1 Then
                
                    frmMsg.fwlblMsg.Caption = "‘Ì›  ﬂ«—Ì ÃœÌœ «ÌÃ«œ ‘œ"
                    frmMsg.fwBtn(0).Visible = False
                    frmMsg.fwBtn(1).ButtonType = flwButtonOk
                    frmMsg.fwBtn(1).Caption = "ﬁ»Ê·"
                    frmMsg.Show vbModal
                
                    FillVsShift
                    MyFormAddEditMode = ViewMode
                    SetFirstToolBar
                Else
                
                    frmMsg.fwlblMsg.Caption = "‘Ì›  ﬂ«—Ì «ÌÃ«œ ‰‘œ" & vbCrLf & "·ÿ›« œÊ»«—Â ”⁄Ì ò‰Ìœ"
                    frmMsg.fwBtn(0).Visible = False
                    frmMsg.fwBtn(1).ButtonType = flwButtonOk
                    frmMsg.fwBtn(1).Caption = "ﬁ»Ê·"
                    frmMsg.Show vbModal
                    
                End If
                    
            Else
                
                frmMsg.fwlblMsg.Caption = "‘„« „Ì »«Ì”  „ÕœÊœÂ “„«‰Ì „⁄ »—Ì Ê«—œ ‰„«ÌÌœ"
                frmMsg.fwBtn(0).Visible = False
                frmMsg.fwBtn(1).ButtonType = flwButtonOk
                frmMsg.fwBtn(1).Caption = "ﬁ»Ê·"
                frmMsg.Show vbModal
                
            End If
            
        Case EditMode 'Edit
        
            If ValidRange = True Then
                ReDim Parameter(6) As Parameter
                Parameter(0) = GenerateInputParameter("@Code", adInteger, 4, txtDescription.Tag)
                Parameter(1) = GenerateInputParameter("@Description", adVarWChar, 50, txtDescription.Text)
                Parameter(2) = GenerateInputParameter("@StartTime", adDBTime, 8, FormatDateTime(DTPStart.Value, vbShortTime))
                Parameter(3) = GenerateInputParameter("@EndTime", adDBTime, 8, FormatDateTime(DTPEnd.Value, vbShortTime))
                Parameter(4) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
                Parameter(5) = GenerateInputParameter("@Branch", adInteger, 4, CurrentBranch)
                Parameter(6) = GenerateOutputParameter("@Result", adInteger, 4)
                
                If RunParametricStoredProcedure("Update_tShift", Parameter) <> -1 Then
                
                    frmMsg.fwlblMsg.Caption = "‘Ì›  ﬂ«—Ì  €ÌÌ— ﬂ—œ"
                    frmMsg.fwBtn(0).Visible = False
                    frmMsg.fwBtn(1).ButtonType = flwButtonOk
                    frmMsg.fwBtn(1).Caption = "ﬁ»Ê·"
                    frmMsg.Show vbModal
                
                    MyFormAddEditMode = ViewMode
                    SetFirstToolBar
                    FillVsShift
                Else
                
                    frmMsg.fwlblMsg.Caption = "‘Ì›  ﬂ«—Ì  €ÌÌ— ‰ﬂ—œ" & vbCrLf & "·ÿ›« œÊ»«—Â ”⁄Ì ò‰Ìœ"
                    frmMsg.fwBtn(0).Visible = False
                    frmMsg.fwBtn(1).ButtonType = flwButtonOk
                    frmMsg.fwBtn(1).Caption = "ﬁ»Ê·"
                    frmMsg.Show vbModal
                    
                End If
                
            Else
                
                frmMsg.fwlblMsg.Caption = "‘„« „Ì »«Ì”  „ÕœÊœÂ “„«‰Ì „⁄ »—Ì Ê—œ ‰„«ÌÌœ"
                frmMsg.fwBtn(0).Visible = False
                frmMsg.fwBtn(1).ButtonType = flwButtonOk
                frmMsg.fwBtn(1).Caption = "ﬁ»Ê·"
                frmMsg.Show vbModal
                
            End If
            
            
    End Select
    FillVsShift

End Sub

Public Function ValidRange() As Boolean
    
    Dim RstTemp As New ADODB.Recordset
    Select Case MyFormAddEditMode
        Case AddMode 'add
            
'            ReDim Parameter(0) As Parameter
'            Parameter(0) = GenerateInputParameter("@Branch", adInteger, 4, cmbBranch.ItemData(cmbBranch.ListIndex))
'            Set RstTemp = RunParametricStoredProcedure2Rec("Get_All_tShift", Parameter)
            Set RstTemp = RunStoredProcedure2RecordSet("Get_All_tShift")
        Case EditMode 'Edit
            
            ReDim Parameter(0) As Parameter
            Parameter(0) = GenerateInputParameter("@Code", adInteger, 4, txtDescription.Tag)
'            Parameter(1) = GenerateInputParameter("@Branch", adInteger, 4, cmbBranch.ItemData(cmbBranch.ListIndex))
            Set RstTemp = RunParametricStoredProcedure2Rec("Get_tShift_By_Code_Not_In", Parameter)
            
        Case Else
        
            Set RstTemp = Nothing
            Exit Function
            
    End Select
        
    If Not (RstTemp.EOF = True And RstTemp.BOF = True) Then
        While RstTemp.EOF <> True
            If DTPStart.Value <= DTPEnd.Value Then
            
                If RstTemp.Fields("StartTime").Value < RstTemp.Fields("EndTime").Value Then
                    
                    If (DTPStart.Value < RstTemp.Fields("StartTime").Value And DTPEnd.Value <= RstTemp.Fields("StartTime").Value) Or (DTPStart.Value >= RstTemp.Fields("EndTime").Value And DTPEnd.Value > RstTemp.Fields("EndTime").Value) Then
                        RstTemp.MoveNext
                    Else
                        ValidRange = False
                        Set RstTemp = Nothing
                        Exit Function
                    End If
                    
                Else
                    
                    If (DTPStart.Value < RstTemp.Fields("StartTime").Value And DTPEnd.Value <= RstTemp.Fields("StartTime").Value And DTPStart.Value >= RstTemp.Fields("EndTime").Value) Or (DTPStart.Value >= RstTemp.Fields("EndTime").Value And DTPEnd.Value > RstTemp.Fields("EndTime").Value) Then
                        RstTemp.MoveNext
                    Else
                        ValidRange = False
                        Set RstTemp = Nothing
                        Exit Function
                    End If
                    
                End If
                
            Else
            
                If RstTemp.Fields("StartTime").Value < RstTemp.Fields("EndTime").Value Then
                    
                    If (DTPStart.Value < RstTemp.Fields("StartTime").Value And DTPEnd.Value <= RstTemp.Fields("StartTime").Value) Or (DTPStart.Value >= RstTemp.Fields("EndTime").Value And DTPEnd.Value <= RstTemp.Fields("startTime").Value) Then
                        RstTemp.MoveNext
                    Else
                        ValidRange = False
                        Set RstTemp = Nothing
                        Exit Function
                    End If
                    
                Else
                    
                    If (DTPStart.Value < RstTemp.Fields("StartTime").Value And DTPEnd.Value <= RstTemp.Fields("StartTime").Value And DTPStart.Value >= RstTemp.Fields("EndTime").Value) Or (DTPStart.Value >= RstTemp.Fields("EndTime").Value And DTPEnd.Value > RstTemp.Fields("EndTime").Value And DTPEnd.Value >= RstTemp.Fields("StartTime").Value) Then
                        RstTemp.MoveNext
                    Else
                        ValidRange = False
                        Set RstTemp = Nothing
                        Exit Function
                    End If
                    
                End If
            
            End If
        Wend
        ValidRange = True
    Else
        ValidRange = True
    End If
    
    Set RstTemp = Nothing
End Function

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

Private Sub Form_Load()

    If ClsFormAccess.frmWorkTime = False Then
        Unload Me
        Exit Sub
    End If
        
    CenterCenter Me
    
    mdifrm.Toolbar1.Buttons(8).Enabled = True
    
    MyFormAddEditMode = ViewMode 'view mode
    
    VarActForm = Me.Name
    
    DTPStart.Format = dtpTime
    DTPStart.CustomFormat = "HH:mm"

    DTPEnd.Format = dtpTime
    DTPEnd.CustomFormat = "HH:mm"

    With vsShift
        .Cols = 5
        .Rows = 1
        .TextMatrix(0, 1) = "òœ"
        .TextMatrix(0, 2) = "‰«„ ‘Ì› "
        .TextMatrix(0, 3) = "”«⁄  ‘—Ê⁄ ﬂ«—"
        .TextMatrix(0, 4) = "”«⁄  « „«„ ﬂ«—"
        
        .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter
        .AutoSizeMode = flexAutoSizeColWidth
        .AutoSize 0, .Cols - 1
   End With
'    FillBranch
    FillVsShift
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

    mdifrm.Toolbar1.Buttons(20).Enabled = False
    mdifrm.Toolbar1.Buttons(21).Enabled = False
 '   mdifrm.Toolbar1.Buttons(27).Enabled = False
    mdifrm.Toolbar1.Buttons(8).Enabled = False
    VarActForm = ""
    'mdifrm.Arrange 0
    mdifrm.Toolbar1.Buttons(20).Enabled = False
    mdifrm.Toolbar1.Buttons(21).Enabled = False
    mdifrm.Toolbar1.Buttons(23).Enabled = True
    mdifrm.Toolbar1.Buttons(24).Enabled = True
    mdifrm.Toolbar1.Buttons(25).Enabled = True
    mdifrm.Toolbar1.Buttons(26).Enabled = True
    mdifrm.Toolbar1.Buttons(27).Enabled = True
    
    If PosConnection.State = adStateOpen Then PosConnection.Close: Set PosConnection = Nothing


    SaveSetting strMainKey, Me.Name, "Left", Me.Left
    SaveSetting strMainKey, Me.Name, "Top", Me.Top


    
End Sub


Private Sub ResizeKit1_ExitResize(ByVal XScale As Double, ByVal YScale As Double)


    If formloadFlag = True Then
        SaveSetting strMainKey, Me.Name, "Height", Me.Height
        SaveSetting strMainKey, Me.Name, "Width", Me.Width
    End If

End Sub

'Private Sub txtStartTime_Validate(Cancel As Boolean)
'    If InStr(1, txtStartTime.Text, "_") <> 0 Then
'        Dim tmpstr As String
'        tmpstr = Replace(txtStartTime.Text, ":", "")
'        tmpstr = Replace(tmpstr, "_", "")
'        Do While Len(tmpstr) < 4
'            tmpstr = "0" & tmpstr
'        Loop
'        txtStartTime.Text = Format(tmpstr, "00:00")
'    End If
'End Sub
'
'
'Private Sub txtStopTime_Validate(Cancel As Boolean)
'    If InStr(1, txtStopTime.Text, "_") <> 0 Then
'        Dim tmpstr As String
'        tmpstr = Replace(txtStopTime.Text, ":", "")
'        tmpstr = Replace(tmpstr, "_", "")
'        Do While Len(tmpstr) < 4
'            tmpstr = "0" & tmpstr
'        Loop
'        txtStopTime.Text = Format(tmpstr, "00:00")
'    End If
'
'
'End Sub

Private Sub vsShift_RowColChange()
    With vsShift
        If .Row > 0 Then
            txtDescription.Tag = .TextMatrix(.Row, 1)
            txtDescription.Text = .TextMatrix(.Row, 2)
            DTPStart.Value = .TextMatrix(.Row, 3)
            DTPEnd.Value = .TextMatrix(.Row, 4)
        ElseIf .Row = 0 Then
            .Row = 1
        End If
    End With
End Sub
