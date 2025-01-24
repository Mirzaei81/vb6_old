VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{7AEDC602-D94C-11D1-BB7A-00E0290EA3C9}#1.0#0"; "ResizeKit.ocx"
Object = "{75D4F767-8785-11D3-93AD-0000832EF44D}#2.14#0"; "FAST2001.ocx"
Begin VB.Form frmNotice 
   ClientHeight    =   6375
   ClientLeft      =   5055
   ClientTop       =   450
   ClientWidth     =   11715
   Icon            =   "frmNotice.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   6375
   ScaleMode       =   0  'User
   ScaleWidth      =   11715
   Begin VB.CommandButton CmdNoticForSms 
      BackColor       =   &H00008000&
      Caption         =   "«—”«· ÃÂ  SMS"
      BeginProperty Font 
         Name            =   "B Homa"
         Size            =   12
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   1800
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   5520
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.CommandButton cmdDone 
      BackColor       =   &H00008000&
      Caption         =   " «ÌÌœ"
      BeginProperty Font 
         Name            =   "B Homa"
         Size            =   12
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   5520
      Width           =   1455
   End
   Begin VSFlex7LCtl.VSFlexGrid vsNotice 
      Height          =   4695
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   8895
      _cx             =   15690
      _cy             =   8281
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
      BackColorFixed  =   -2147483643
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
      AllowUserResizing=   0
      SelectionMode   =   0
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
   Begin VB.ListBox lstPrintFormat 
      BeginProperty Font 
         Name            =   "Nazanin"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3960
      Left            =   9000
      RightToLeft     =   -1  'True
      Style           =   1  'Checkbox
      TabIndex        =   1
      Top             =   1200
      Width           =   2655
   End
   Begin FLWCtrls.FWLabel3D fwlblMode 
      Height          =   495
      Left            =   10200
      Top             =   0
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   873
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
   Begin RESIZEKITLibCtl.ResizeKit ResizeKit1 
      Height          =   480
      Left            =   0
      OleObjectBlob   =   "frmNotice.frx":A4C2
      TabIndex        =   5
      Top             =   0
      Width           =   480
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ç«Å ‘⁄«—Â«  œ—«‰ Â«Ì ›Ì‘"
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
      TabIndex        =   4
      Top             =   0
      Width           =   3015
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      BackStyle       =   0  'Transparent
      Caption         =   "‰Ê⁄ ç«Å"
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
      Left            =   10440
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   600
      Width           =   1215
   End
End
Attribute VB_Name = "frmNotice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i As Integer
Private MyFormAddEditMode As EnumAddEditMode
Dim Parameter() As Parameter

Public Sub Add()
    
    MyFormAddEditMode = AddMode
    SetFirstToolBar
    With vsNotice
        .Rows = .Rows + 1
        .TextMatrix(.Rows - 1, 0) = "*"
        
    End With
    
End Sub
Public Sub Cancel()

    FillvsNotice
    MyFormAddEditMode = ViewMode
    SetFirstToolBar
    
End Sub
Public Sub Edit()

    MyFormAddEditMode = EditMode
    SetFirstToolBar

End Sub

Public Sub Update()
    
    Dim Rst As New ADODB.Recordset
    Dim intNoticeNo As Integer
    vsNotice_ValidateEdit vsNotice.Row, vsNotice.Col, False
'    vsNotice.SetFocus
    Select Case MyFormAddEditMode
        Case AddMode 'add
        
            With vsNotice
            
                For i = 1 To .Rows - 1
                    If InStr(1, .TextMatrix(i, 0), "*") > 0 Then
                        If Trim(.TextMatrix(i, 3)) <> "" Then
                            Set Rst = RunStoredProcedure2RecordSet("CheckNoticeNo")
                            intNoticeNo = Rst.Fields("MaxNoticeNo").Value + 1
                            Set Rst = Nothing
                            
                            ReDim Parameter(2) As Parameter
                            
                            Parameter(0) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
                            Parameter(1) = GenerateInputParameter("@NoticeNo", adInteger, 4, intNoticeNo)
                            Parameter(2) = GenerateInputParameter("@NoticeDescription", adVarWChar, 255, .TextMatrix(i, 3))
                            
                            RunParametricStoredProcedure "InserttNoticeDescription", Parameter
                        End If
                    End If
                Next i
                frmMsg.fwlblMsg.Caption = "«ÿ·«⁄«  ÃœÌœ »« „Ê›ﬁÌ  À»  ‘œ"
                frmMsg.fwBtn(0).Visible = False
                frmMsg.fwBtn(1).ButtonType = flwButtonOk
                frmMsg.fwBtn(1).Caption = "ﬁ»Ê·"
                frmMsg.Show vbModal
            End With
        Case EditMode 'edit
        
            With vsNotice
                For i = 1 To .Rows - 1
                    If InStr(1, .TextMatrix(i, 0), "*") > 0 Then
                        If Trim(.TextMatrix(i, 3)) <> "" Then
                            
                            ReDim Parameter(2) As Parameter
                            Parameter(0) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
                            Parameter(1) = GenerateInputParameter("@NoticeNo", adInteger, 4, .TextMatrix(i, 1))
                            Parameter(2) = GenerateInputParameter("@NoticeDescription", adVarWChar, 255, .TextMatrix(i, 3))
                            
                            RunParametricStoredProcedure "UpdatetNoticeDescription", Parameter
                                
                        End If
                    End If
                Next i
            End With
            frmMsg.fwlblMsg.Caption = "À»   €ÌÌ—«  »« „Ê›ﬁÌ  »Â Å«Ì«‰ —”Ìœ"
            frmMsg.fwBtn(0).Visible = False
            frmMsg.fwBtn(1).ButtonType = flwButtonOk
            frmMsg.fwBtn(1).Caption = "ﬁ»Ê·"
            frmMsg.Show vbModal
            
    End Select
    
    
    DefaultSetting
    MyFormAddEditMode = ViewMode
    SetFirstToolBar
    
    Exit Sub
RollBack:
    
    frmMsg.fwlblMsg.Caption = " €ÌÌ—«  „Ê—œ ‰Ÿ— «⁄„«· ‰‘œ" + vbCrLf + "·ÿ›« «ÿ·«⁄«  ò«„· Ê œ—”  Ê«—œ ‰„«ÌÌœ"
    frmMsg.fwBtn(0).Visible = False
    frmMsg.fwBtn(1).ButtonType = flwButtonOk
    frmMsg.fwBtn(1).Caption = "ﬁ»Ê·"
    frmMsg.Show vbModal
    
End Sub

Public Sub SetFirstToolBar()
    
    AllButton vbOff, True
    
    mdifrm.Toolbar1.Buttons(15).Enabled = True   'Print
    
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

Public Sub DefaultSetting()
    lstPrintFormat.Clear
    
    FilllstPrintFormat
    FillvsNotice
End Sub

Public Sub ExitForm()
    
    Unload Me
    
End Sub

Public Sub FilllstPrintFormat()
    
    Dim Rst As New ADODB.Recordset
    ReDim Parameter(0) As Parameter
    
    Parameter(0) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
      
    lstPrintFormat.Clear
    
    Set Rst = RunParametricStoredProcedure2Rec("FillPrintFormatList", Parameter)
    
    If Not (Rst.EOF = True And Rst.BOF = True) Then
        
        While Rst.EOF <> True
            lstPrintFormat.AddItem Rst.Fields("PrintFormatName").Value
            lstPrintFormat.ItemData(lstPrintFormat.ListCount - 1) = Rst.Fields("PrintFormat").Value
            Rst.MoveNext
        Wend
    
    End If
    
    
End Sub

Public Sub FillvsNotice()
    
    Dim Rst As New ADODB.Recordset
    ReDim Parameter(0) As Parameter
    
    Parameter(0) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
    Set Rst = RunParametricStoredProcedure2Rec("FillNoticeList", Parameter)
    
    With vsNotice
        .Rows = 1
        If Not (Rst.BOF = True And Rst.EOF = True) Then

            While Rst.EOF <> True
                .Rows = .Rows + 1
                .Row = .Rows - 1
                .TextMatrix(.Row, 0) = .Row
                .TextMatrix(.Row, 1) = Rst.Fields("NoticeNo").Value
                .TextMatrix(.Row, 3) = IIf(IsNull(Rst.Fields("NoticeDescription").Value), "", Rst.Fields("NoticeDescription").Value)
                Rst.MoveNext
            Wend
            .AutoSizeMode = flexAutoSizeColWidth
            .AutoSize 0, .Cols - 1
        End If
        If .Rows > 1 Then
            .Cell(flexcpText, 1, 2, .Rows - 1, 2) = 0
        End If
    End With


End Sub

Private Sub CmdDone_Click()

    Dim intSelNoticeType As Integer
    Dim dt As New clsDate
    ReDim Parameter(1) As Parameter
    
    i = -1
    For i = 0 To lstPrintFormat.ListCount - 1
    
        If lstPrintFormat.Selected(i) = True Then
            intSelNoticeType = i
            Exit For
        End If
    
    Next i
    
    If i <> -1 Then
        On Error GoTo RollBack
        Dim blnDone As Boolean


        With vsNotice
            For i = 1 To .Rows - 1
                If Val(.TextMatrix(i, 2)) = -1 Then
                
                            Parameter(0) = GenerateInputParameter("@PrintFormat", adInteger, 4, lstPrintFormat.ItemData(intSelNoticeType))
                            Parameter(1) = GenerateInputParameter("@NoticeNo", adInteger, 4, .TextMatrix(i, 1))

                            RunParametricStoredProcedure "UpdatetPrintFormat", Parameter
                    
                    blnDone = True
                    Exit For
                End If
            Next i
            If blnDone <> True Then
            
                Parameter(0) = GenerateInputParameter("@PrintFormat", adInteger, 4, lstPrintFormat.ItemData(intSelNoticeType))
                Parameter(1) = GenerateInputParameter("@NoticeNo", adInteger, 4, 0)

                RunParametricStoredProcedure "UpdatetPrintFormat", Parameter
                
            End If
        End With
        On Error GoTo 0
        
        frmMsg.fwlblMsg.Caption = "À»   €ÌÌ—«  »« „Ê›ﬁÌ  »Â Å«Ì«‰ —”Ìœ"
        frmMsg.fwBtn(0).Visible = False
        frmMsg.fwBtn(1).ButtonType = flwButtonOk
        frmMsg.fwBtn(1).Caption = "ﬁ»Ê·"
        frmMsg.Show vbModal
        
    End If
    Exit Sub
RollBack:

    err.Clear
    frmMsg.fwlblMsg.Caption = " €ÌÌ—«  „Ê—œ ‰Ÿ— «⁄„«· ‰‘œ" + vbCrLf + "·ÿ›« «ÿ·«⁄«  ò«„· Ê œ—”  Ê«—œ ‰„«ÌÌœ"
    frmMsg.fwBtn(0).Visible = False
    frmMsg.fwBtn(1).ButtonType = flwButtonOk
    frmMsg.fwBtn(1).Caption = "ﬁ»Ê·"
    frmMsg.Show vbModal
    
End Sub

Private Sub CmdNoticForSms_Click()
    If vsNotice.Row > 0 Then
        If frmSms.txtSMSMessage.Text <> "" Then
           frmSms.txtSMSMessage.Text = frmSms.txtSMSMessage.Text & vsNotice.TextMatrix(vsNotice.Row, 3)
        Else
            frmSms.txtSMSMessage.Text = vsNotice.TextMatrix(vsNotice.Row, 3)
        End If
    End If
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

Private Sub Form_Load()
    
    If ClsFormAccess.frmNotice = False Then
        Unload Me
        Exit Sub
    End If
    
    CenterTop Me
    
  ' Me.Height = 10020
  '  Me.Width = 14835
    
    VarActForm = Me.Name
    
    With vsNotice
        .Cols = 4
        .Rows = 1
        .Row = 0
        .TextMatrix(.Row, 0) = "—œÌ›"
        .TextMatrix(.Row, 1) = "òœ"
        .TextMatrix(.Row, 2) = "«‰ Œ«»"
        .TextMatrix(.Row, 3) = "„ ‰ ‘⁄«—"
        .ColHidden(1) = True
        .ColDataType(2) = flexDTBoolean
        .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter
        .ColAlignment(3) = flexAlignRightCenter
        .ColAlignment(0) = flexAlignCenterCenter
        .AutoSizeMode = flexAutoSizeColWidth
        .AutoSize 0, .Cols - 1
    End With

    DefaultSetting
    MyFormAddEditMode = ViewMode
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
    
        If PosConnection.State = adStateOpen Then PosConnection.Close: Set PosConnection = Nothing
        VarActForm = ""
        If CmdNoticForSms.Visible = True Then
            ExitForm
        Else
    
            SaveSetting strMainKey, Me.Name, "Left", Me.Left
            SaveSetting strMainKey, Me.Name, "Top", Me.Top
        End If
End Sub

Private Sub lstPrintFormat_ItemCheck(Item As Integer)
    
    Dim Rst As New ADODB.Recordset
    
    Dim SelectPrintFormat As Integer
    If lstPrintFormat.Selected(Item) = True Then
        SelectPrintFormat = lstPrintFormat.ItemData(Item)
        For i = 0 To lstPrintFormat.ListCount - 1
            If i <> Item And lstPrintFormat.Selected(i) = True Then
                lstPrintFormat.Selected(i) = False
            
            End If
        Next i
    
        FillvsNotice
        
        With vsNotice
            If .Rows > 1 Then
                .Cell(flexcpText, 1, 2, .Rows - 1, 2) = 0
                ReDim Parameter(0) As Parameter
                Parameter(0) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
                Set Rst = RunParametricStoredProcedure2Rec("Get_tPrintFormat_By_PrintFormat", Parameter)
                
                If Not (Rst.EOF = True And Rst.BOF = True) Then
                    While Rst.EOF <> True
                       If Rst.Fields("PrintFormat").Value = SelectPrintFormat Then
                            For i = 0 To .Rows - 1
                                If Rst.Fields("NoticeNo").Value = .TextMatrix(i, 1) Then
                                     .TextMatrix(i, 2) = -1
                                     Exit Sub
                                End If
                            
                            Next i
                        End If
                        Rst.MoveNext
                    Wend
                End If
            End If
        End With
    
    Else
        FillvsNotice
    End If

End Sub



Private Sub ResizeKit1_ExitResize(ByVal XScale As Double, ByVal YScale As Double)

    If formloadFlag = True Then
        SaveSetting strMainKey, Me.Name, "Height", Me.Height
        SaveSetting strMainKey, Me.Name, "Width", Me.Width
    End If

End Sub

Private Sub vsNotice_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If MyFormAddEditMode = EditMode Then
        With vsNotice
            If Row > 0 And InStr(1, .TextMatrix(Row, 0), "*") = 0 Then
                .TextMatrix(Row, 0) = .TextMatrix(Row, 0) & "*"
            End If
        End With
    End If
End Sub

Private Sub vsNotice_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    With vsNotice
        Select Case MyFormAddEditMode
            Case ViewMode
                If .Col = 2 And .Row > 0 Then
                    .Select .Row, .Col
                    .EditCell
                    If .TextMatrix(.Row, .Col) = -1 Then
                        For i = 1 To .Rows - 1
                            If i <> .Row Then
                                .TextMatrix(i, .Col) = 0
                            End If
                        Next i
                    End If
                End If
            Case AddMode
                If Mid(.TextMatrix(.Row, 0), 1, 1) = "*" And .Col > 0 And .Row > 0 Then
                    .Select .Row, .Col
                    .EditCell
                End If
            Case EditMode
                If .Col > 0 And .Row > 0 Then
                    .Select .Row, .Col
                    .EditCell
                End If
        End Select
    End With

End Sub

Private Sub vsNotice_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsNotice
        .Row = Row
        .Col = Col
    End With
End Sub
