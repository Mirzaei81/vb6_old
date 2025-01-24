VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{7AEDC602-D94C-11D1-BB7A-00E0290EA3C9}#1.0#0"; "ResizeKit.ocx"
Object = "{75D4F767-8785-11D3-93AD-0000832EF44D}#2.14#0"; "FAST2001.ocx"
Begin VB.Form frmGoodDifferences 
   Caption         =   "       ¬Å‘‰ Â«Ì  ò«·«"
   ClientHeight    =   9195
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8475
   Icon            =   "frmGoodDifferences.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   9195
   ScaleWidth      =   8475
   Begin VB.CommandButton cmdDeleteAll 
      BackColor       =   &H00008000&
      Caption         =   "Å«ò ò—œ‰ Â„Â"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Tag             =   "1"
      Top             =   7200
      Width           =   1095
   End
   Begin VB.CommandButton cmdSelectAll 
      BackColor       =   &H00008000&
      Caption         =   "«‰ Œ«» Â„Â"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Tag             =   "1"
      Top             =   5640
      Width           =   1095
   End
   Begin VSFlex7LCtl.VSFlexGrid vsDifferences 
      Height          =   2925
      Left            =   1290
      TabIndex        =   2
      Top             =   1050
      Width           =   7125
      _cx             =   12568
      _cy             =   5159
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
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   3
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   7
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmGoodDifferences.frx":A4C2
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
   Begin VB.CommandButton cmdDone 
      BackColor       =   &H00008000&
      Caption         =   "«Œ ’«’ ¬Å‘‰ »Â ﬂ«·«Â«"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1485
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   1
      ToolTipText     =   "»« «‰ Œ«»  €ÌÌ—«  Ê «‰ Œ«» ﬂ«·«Â« «— »«ÿ ¬‰Â« »« Â„ «‰Ã«„ „Ì êÌ—œ"
      Top             =   3360
      Width           =   1035
   End
   Begin VSFlex7LCtl.VSFlexGrid vsGood 
      Height          =   4425
      Left            =   1320
      TabIndex        =   0
      Top             =   4680
      Width           =   7050
      _cx             =   12435
      _cy             =   7805
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
      BackColorFixed  =   12640511
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
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   7
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   500
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmGoodDifferences.frx":A568
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
   Begin FLWCtrls.FWLabel3D fwlblMode 
      Height          =   495
      Left            =   7005
      Top             =   0
      Width           =   1455
      _ExtentX        =   2566
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
      OleObjectBlob   =   "frmGoodDifferences.frx":A607
      TabIndex        =   5
      Top             =   0
      Width           =   480
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "»« «‰ Œ«» ê—ÊÂÌ ¬Å‘‰  Â« Â„Â ﬂ«·«Â«∆Ì ﬂÂ œ«—«Ì «Ì‰ ¬Å‘‰  Â« Â” ‰œ ‰„«Ì‘ œ«œÂ „Ì ‘Ê‰œ"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   615
      Left            =   2880
      TabIndex        =   7
      Top             =   4080
      Width           =   4695
      WordWrap        =   -1  'True
   End
   Begin VB.Label LblNote 
      Alignment       =   1  'Right Justify
      Caption         =   $"frmGoodDifferences.frx":A68D
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   975
      Left            =   1320
      TabIndex        =   6
      Top             =   0
      Width           =   5655
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ò«·«Â«"
      BeginProperty Font 
         Name            =   "B Homa"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   555
      Left            =   7680
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   4200
      Width           =   645
   End
   Begin VB.Label lable1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "¬Å‘‰ "
      BeginProperty Font 
         Name            =   "B Homa"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   555
      Left            =   7200
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   510
      Width           =   1125
   End
End
Attribute VB_Name = "frmGoodDifferences"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i As Integer
Private MyFormAddEditMode As EnumAddEditMode
Dim Parameter() As Parameter

Public Sub Add()
    
    Cancel
    MyFormAddEditMode = AddMode
    SetFirstToolBar
    
    With vsDifferences
        .Rows = .Rows + 1
        .TextMatrix(.Rows - 1, 0) = "*"
        .Row = .Rows - 1
        .ShowCell .Row, 0
        .Select .Row, 0
    End With
    
End Sub
Public Sub Cancel()

    DefaultSetting
    MyFormAddEditMode = ViewMode
    SetFirstToolBar
    
End Sub
Public Sub Edit()

    MyFormAddEditMode = EditMode
    SetFirstToolBar

End Sub
Public Sub Delete()

    Dim tmpStr As String
    With vsDifferences
        If .SelectedRows > 0 Then
            tmpStr = Trim(.TextMatrix(.Row, 3))
            If tmpStr <> "" Then
                If Trim(.TextMatrix(.Row, 5)) <> "" Then
                    tmpStr = "„Ê—œ " & tmpStr & " Ê " & Trim(.TextMatrix(.Row, 5)) & " —« "
                Else
                    tmpStr = "„Ê—œ " & tmpStr & " —« "
                End If
            Else
                tmpStr = "„Ê—œ " & Trim(.TextMatrix(.Row, 5)) & " —« "
            End If
            
            frmMsg.fwlblMsg.Caption = "¬Ì« „ÿ„∆‰Ìœ „Ì ŒÊ«ÂÌœ " & tmpStr & " Õ–› ﬂ‰Ìœø"
            frmMsg.Show vbModal
            If modgl.mvarMsgIdx = vbYes Then
                
                ReDim Parameter(0) As Parameter
                Parameter(0) = GenerateInputParameter("@code", adInteger, 4, .TextMatrix(.Row, 1))
                RunParametricStoredProcedure "DeleteDifferences", Parameter
                
                frmMsg.fwlblMsg.Caption = " .‘„« " & tmpStr & " Õ–› ﬂ—œÂ «Ìœ"
                frmMsg.fwBtn(1).Visible = False
                frmMsg.fwBtn(0).ButtonType = flwButtonOk
                frmMsg.fwBtn(0).Caption = "ﬁ»Ê·"
                frmMsg.Show vbModal
            End If
        End If
        DefaultSetting
    End With
    
    
    MyFormAddEditMode = ViewMode
    SetFirstToolBar

End Sub

Public Sub Update()
    
    On Error GoTo ErrHandler
    
    Dim RecordIsAdded As Boolean
    Dim Rst As New ADODB.Recordset
    
    Select Case MyFormAddEditMode
        Case AddMode
             
            With vsDifferences
                vsDifferences_ValidateEdit .Row, .Col, False
                For i = 1 To .Rows - 1
                    If InStr(1, .TextMatrix(i, 0), "*") > 0 And (Trim(.TextMatrix(i, 3)) <> "" Or Trim(.TextMatrix(i, 5)) <> "") Then
                        ReDim Parameter(4) As Parameter
                        Parameter(0) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
                        Parameter(1) = GenerateInputParameter("@Defference", adVarWChar, 200, .TextMatrix(i, 3))
                        Parameter(2) = GenerateInputParameter("@NegativeDefference", adVarWChar, 200, .TextMatrix(i, 5))
                        Parameter(3) = GenerateInputParameter("@CostDifference", adInteger, 4, Val(.TextMatrix(i, 6)))
                        Parameter(4) = GenerateOutputParameter("@LastCode", adInteger, 4)
                        Dim Result As Long
                        Result = RunParametricStoredProcedure("Insert_Differences", Parameter)
                        RecordIsAdded = True
                    End If
                Next i
            End With
            If RecordIsAdded Then
                frmMsg.fwlblMsg.Caption = "«ÿ·«⁄«  ÃœÌœ »« „Ê›ﬁÌ  À»  ‘œ"
                frmMsg.fwBtn(0).Visible = False
                frmMsg.fwBtn(1).ButtonType = flwButtonOk
                frmMsg.fwBtn(1).Caption = "ﬁ»Ê·"
                frmMsg.Show vbModal
            
            End If
        Case EditMode
        
            With vsDifferences
                vsDifferences_ValidateEdit .Row, .Col, False
                For i = 1 To .Rows - 1
                    If InStr(1, .TextMatrix(i, 0), "*") > 0 And (Trim(.TextMatrix(i, 3)) <> "" Or Trim(.TextMatrix(i, 5)) <> "") Then
                        ReDim Parameter(3) As Parameter
                        Parameter(0) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
                        Parameter(1) = GenerateInputParameter("@Code", adInteger, 4, Val(.TextMatrix(i, 1)))
                        Parameter(2) = GenerateInputParameter("@Difference", adVarWChar, 200, .TextMatrix(i, 3))
                        Parameter(3) = GenerateInputParameter("@CostDifference", adInteger, 4, .TextMatrix(i, 6))
                        RunParametricStoredProcedure "Edit_Differences", Parameter
                        RecordIsAdded = True
                        
                        Parameter(0) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
                        Parameter(1) = GenerateInputParameter("@Code", adInteger, 4, -Val(.TextMatrix(i, 1)))
                        Parameter(2) = GenerateInputParameter("@Difference", adVarWChar, 200, .TextMatrix(i, 5))
                        Parameter(3) = GenerateInputParameter("@CostDifference", adInteger, 4, 0)
                        RunParametricStoredProcedure "Edit_Differences", Parameter
                        RecordIsAdded = True
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
ErrHandler:
    
    frmMsg.fwlblMsg.Caption = " €ÌÌ—«  „Ê—œ ‰Ÿ— «⁄„«· ‰‘œ"
    frmMsg.fwBtn(0).Visible = False
    frmMsg.fwBtn(1).ButtonType = flwButtonOk
    frmMsg.fwBtn(1).Caption = "ﬁ»Ê·"
    frmMsg.Show vbModal
    
End Sub

Private Sub cmdDeleteAll_Click()
    If vsGood.Rows > 1 Then
        vsGood.Cell(flexcpText, vsGood.FixedRows, 0, vsGood.Rows - 1, 0) = 0
    End If

End Sub

Private Sub cmdSelectAll_Click()
    If vsGood.Rows > 1 Then
'        If cmdSelectAll.Tag = 1 Then
'            cmdSelectAll.Tag = 0
            vsGood.Cell(flexcpText, vsGood.FixedRows, 0, vsGood.Rows - 1, 0) = -1
'            cmdSelectAll.Caption = "Å«ò ò—œ‰ Â„Â"
'        Else
'            cmdSelectAll.Tag = 1
'            vsGood.Cell(flexcpText, vsGood.FixedRows, 0, vsGood.Rows - 1, 0) = 0
'            cmdSelectAll.Caption = "«‰ Œ«» Â„Â"
'        End If
    End If
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call PresetScreenSaver
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call PresetScreenSaver
End Sub

Public Sub ChangeLanguage()
    
   FillvsDifferences
   FillvsGood
   
End Sub
Public Sub SetFirstToolBar()
    
    AllButton vbOff, True
    
    mdifrm.Toolbar1.Buttons(15).Enabled = True   'Print
    
    mdifrm.Toolbar1.Buttons(23).Enabled = True
    mdifrm.Toolbar1.Buttons(24).Enabled = True
    mdifrm.Toolbar1.Buttons(25).Enabled = True
    mdifrm.Toolbar1.Buttons(26).Enabled = True
    mdifrm.Toolbar1.Buttons(27).Enabled = True
 
    If MyFormAddEditMode = ViewMode Then
 
        mdifrm.Toolbar1.Buttons(6).Enabled = True  'Add
        mdifrm.Toolbar1.Buttons(7).Enabled = True  'Edit
        mdifrm.Toolbar1.Buttons(8).Enabled = False  'Enter
        mdifrm.Toolbar1.Buttons(9).Enabled = False   'Esc
        mdifrm.Toolbar1.Buttons(10).Enabled = True   'Esc
        
    ElseIf MyFormAddEditMode = AddMode Then    'Add Mode
                
        mdifrm.Toolbar1.Buttons(6).Enabled = False  'Add
        mdifrm.Toolbar1.Buttons(7).Enabled = False  'Edit
        mdifrm.Toolbar1.Buttons(8).Enabled = True   'Enter
        mdifrm.Toolbar1.Buttons(9).Enabled = True   'Esc
        mdifrm.Toolbar1.Buttons(10).Enabled = False   'Esc
            
    ElseIf MyFormAddEditMode = EditMode Then     'Edit
        
        mdifrm.Toolbar1.Buttons(6).Enabled = True  'Add
        mdifrm.Toolbar1.Buttons(7).Enabled = False  'Edit
        mdifrm.Toolbar1.Buttons(8).Enabled = True   'Enter
        mdifrm.Toolbar1.Buttons(9).Enabled = True   'Esc
        mdifrm.Toolbar1.Buttons(10).Enabled = False   'Esc
        
    End If

    HeaderLabel Val(MyFormAddEditMode), fwlblMode
End Sub

Public Sub DefaultSetting()
    vsDifferences.Rows = 1
    
    FillvsDifferences
    FillvsGood
End Sub

Public Sub ExitForm()

    Unload Me
    
End Sub

Public Sub FillvsDifferences()
    
    Dim Rst As New ADODB.Recordset
    ReDim Parameter(0) As Parameter
    
    Parameter(0) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
    Set Rst = RunParametricStoredProcedure2Rec("GetDifferences", Parameter)
    
    With vsDifferences
        .Rows = 1
        While Rst.EOF <> True
        
            .Rows = .Rows + 1
            i = .Rows - 1
            .TextMatrix(i, 1) = Rst!absCode
            .TextMatrix(i, 3) = IIf(IsNull(Rst!Difference), "", Rst!Difference)
            .TextMatrix(i, 5) = IIf(IsNull(Rst!NegativeDifference), "", Rst!NegativeDifference)
           .TextMatrix(i, 6) = IIf(IsNull(Rst!CostDifference), "0", Rst!CostDifference)
            
            Rst.MoveNext
        Wend
    End With
    Set Rst = Nothing
End Sub
Private Sub UpdateGoodDiffernces(Row As Long)
    Dim Rst As New ADODB.Recordset
    ReDim Parameter(1) As Parameter
    
    Parameter(0) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
    
    With vsDifferences
        
        If Val(.TextMatrix(Row, 2)) <> 0 Then
            Parameter(1) = GenerateInputParameter("@DifferenceCode", adInteger, 4, Val(.TextMatrix(Row, 1)))
        ElseIf Val(.TextMatrix(Row, 4)) <> 0 Then
            Parameter(1) = GenerateInputParameter("@DifferenceCode", adInteger, 4, -Val(.TextMatrix(Row, 1)))
'            Else
'                Parameter(1) = GenerateInputParameter("@DifferenceCode", adInteger, 4, 0)
        End If
'        Else
'            Parameter(1) = GenerateInputParameter("@DifferenceCode", adInteger, 4, 1)
    
        Set Rst = RunParametricStoredProcedure2Rec("Get_Good_Level_Difference", Parameter)
        If Not (Rst.EOF = True And Rst.BOF = True) Then
            While Rst.EOF = False
                For i = 1 To vsGood.Rows - 1
                    If Rst.Fields("Code").Value = vsGood.TextMatrix(i, 3) Then
                        vsGood.TextMatrix(i, 0) = -1
                        Exit For
                    End If
                Next i
                Rst.MoveNext
            Wend
        End If
    End With

End Sub
Public Sub FillvsGood()
    
    Dim Rst As New ADODB.Recordset
    
    ReDim Parameter(1) As Parameter
    Parameter(0) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
    Parameter(1) = GenerateInputParameter("@DifferenceCode", adInteger, 4, -1)
    Set Rst = RunParametricStoredProcedure2Rec("Get_Good_Level_Difference", Parameter)
    
    With vsGood
        .Rows = 1
        If Not (Rst.BOF = True And Rst.EOF = True) Then
            While Rst.EOF <> True
                .Rows = .Rows + 1
                i = .Rows - 1
                .TextMatrix(i, 0) = Rst!Selected
                .TextMatrix(i, 1) = Rst!level1
                .TextMatrix(i, 2) = Rst!Level2
                .TextMatrix(i, 3) = Rst!Code
                .TextMatrix(i, 4) = Rst!DesLevel1
                .TextMatrix(i, 5) = Rst!DesLevel2
                .TextMatrix(i, 6) = Rst!Name
 
                
                Rst.MoveNext
            Wend
            .AutoSearch = flexSearchFromCursor
            .AutoSizeMode = flexAutoSizeColWidth
            .AutoSize 0, .Cols - 1
        End If
    End With
    
    Set Rst = Nothing

End Sub


Private Sub CmdDone_Click()

    Dim intSelDifference As Integer
    Dim dt As New clsDate
    Dim SelectedGoods, strDiffernce  As String
    SelectedGoods = ""
    
    With vsGood
        For i = 1 To .Rows - 1
            If Val(.TextMatrix(i, 0)) <> 0 Then
                SelectedGoods = SelectedGoods & "," & .TextMatrix(i, 3)
            End If
        Next i
    End With

    With vsDifferences
        For i = 1 To .Rows - 1
            intSelDifference = 0
            strDiffernce = ""
            If Val(.TextMatrix(i, 2)) <> 0 Or Val(.TextMatrix(i, 4)) <> 0 Then
                If Val(.TextMatrix(i, 2)) <> 0 Then
                    intSelDifference = Val(.TextMatrix(i, 1))
                    strDiffernce = .TextMatrix(i, 3)
                ElseIf Val(.TextMatrix(i, 4)) <> 0 Then
                    intSelDifference = -Val(.TextMatrix(i, 1))
                    strDiffernce = .TextMatrix(i, 5)
                End If
            End If
            If intSelDifference <> 0 Then
                ReDim Parameter(1) As Parameter
                Parameter(0) = GenerateInputParameter("@GoodCode", adBSTR, 4000, SelectedGoods)
                Parameter(1) = GenerateInputParameter("@DifferenceCode", adInteger, 4, intSelDifference)
                RunParametricStoredProcedure "Insert_Goods_Difference", Parameter
                
                ShowDisMessage "À»   €ÌÌ—«  - " & strDiffernce & "  -«‰Ã«„ ‘œ", 1000
            End If
        
        Next i
    End With
    
    DefaultSetting
    MyFormAddEditMode = ViewMode
    SetFirstToolBar
    
    Exit Sub
RollBack:

    err.Clear
    frmMsg.fwlblMsg.Caption = " €ÌÌ—«  „Ê—œ ‰Ÿ— «⁄„«· ‰‘œ" + vbCrLf + "·ÿ›« «ÿ·«⁄«  ò«„· Ê œ—”  Ê«—œ ‰„«ÌÌœ"
    frmMsg.fwBtn(0).Visible = False
    frmMsg.fwBtn(1).ButtonType = flwButtonOk
    frmMsg.fwBtn(1).Caption = "ﬁ»Ê·"
    frmMsg.Show vbModal

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
    
    If ClsFormAccess.frmNotice = False Then
        Unload Me
        Exit Sub
    End If
    
    CenterTop Me
    
    VarActForm = Me.Name
    
    With vsGood
        .Cols = 7
        .Rows = 1
        .Row = 0
        .TextMatrix(.Row, 0) = "«‰ Œ«»"
        .TextMatrix(.Row, 4) = "ê—ÊÂ"
        .TextMatrix(.Row, 5) = "“Ì— ê—ÊÂ"
        .TextMatrix(.Row, 6) = "‰«„ ò«·«"
        .AutoSearch = flexSearchFromCursor
        .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter
        .ColAlignment(-1) = flexAlignRightCenter
        .AutoSizeMode = flexAutoSizeColWidth
        .AutoSize 0, .Cols - 1
    End With
    
    With vsDifferences
        .Rows = 1
        .TextMatrix(.Row, 2) = "«‰ Œ«»"
        .TextMatrix(.Row, 4) = "«‰ Œ«»"
        .TextMatrix(.Row, 3) = " €ÌÌ— „À» "
        .TextMatrix(.Row, 5) = " €ÌÌ— „‰›Ì"
        .TextMatrix(.Row, 6) = "Â“Ì‰Â  €Ì—« "
        .AutoSearch = flexSearchFromCursor
        .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter
        .ColAlignment(-1) = flexAlignRightCenter
        .ColAlignment(6) = flexAlignCenterCenter
        .AutoSizeMode = flexAutoSizeColWidth
        '.AutoSize 0, .Cols - 1
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

    SaveSetting strMainKey, Me.Name, "Left", Me.Left
    SaveSetting strMainKey, Me.Name, "Top", Me.Top

     VarActForm = ""
     
End Sub

Private Sub ResizeKit1_ExitResize(ByVal XScale As Double, ByVal YScale As Double)
    If formloadFlag = True Then
        SaveSetting strMainKey, Me.Name, "Height", Me.Height
        SaveSetting strMainKey, Me.Name, "Width", Me.Width
    End If

End Sub

Private Sub vsDifferences_KeyPress(KeyAscii As Integer)
    If MyFormAddEditMode = EditMode Then
        With vsDifferences
            If .Row > 0 Then
                If .Col = 3 Or .Col = 5 Then
                    
                    .Select .Row, .Col
                    .EditCell
                End If
            End If
        End With
    End If

End Sub

Private Sub vsDifferences_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If MyFormAddEditMode = EditMode Then
        With vsDifferences
            If Row > 0 And InStr(1, .TextMatrix(Row, 0), "*") = 0 Then
                .TextMatrix(Row, 0) = .TextMatrix(Row, 0) & "*"
            End If
        End With
    End If
End Sub


'Private Sub vsDifferences_MouseDown(Button As Integer, Shift As Integer, X As Single, y As Single)
'    With vsDifferences
'        Select Case MyFormAddEditMode
'            Case ViewMode
'                If .Col = 2 And .Row > 0 Then
'                    .Select .Row, .Col
'                    .EditCell
'                    If .TextMatrix(.Row, .Col) = -1 Then
'                        For i = 1 To .Rows - 1
'                            If i <> .Row Then
'                                .TextMatrix(i, .Col) = 0
'                            End If
'                        Next i
'                    End If
'                End If
'            Case AddMode
'                If Mid(.TextMatrix(.Row, 0), 1, 1) = "*" And .Col > 0 And .Row > 0 Then
'                    .Select .Row, .Col
'                    .EditCell
'                End If
'            Case EditMode
'                If .Col > 0 And .Row > 0 Then
'                    .Select .Row, .Col
'                    .EditCell
'                End If
'        End Select
'    End With
'
'End Sub
'
Private Sub vsDifferences_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsDifferences
        .Row = Row
        .Col = Col
    End With
End Sub

Private Sub vsDifferences_Click()
 Dim i As Long
    With vsDifferences
        If (.Col = 2 Or .Col = 4) And .Row > 0 And .TextMatrix(.Row, 1) <> "" Then
            .Select .Row, .Col
            .EditCell
            For i = 1 To vsGood.Rows - 1
                vsGood.TextMatrix(i, 0) = ""
            Next
            For i = 1 To .Rows - 1
                If Val(.TextMatrix(i, 2)) <> 0 Or Val(.TextMatrix(i, 4)) <> 0 Then
                    UpdateGoodDiffernces i
    '                For i = 1 To .Rows - 1
    '                    If Val(.TextMatrix(i, 2)) <> 0 And (i <> .Row Or 2 <> .Col) Then
    '                        .TextMatrix(i, 2) = 0
    '                    ElseIf Val(.TextMatrix(i, 4)) <> 0 And (i <> .Row Or 4 <> .Col) Then
    '                        .TextMatrix(i, 4) = 0
    '
    '                    End If
    '                Next i
                End If
            Next
        ElseIf .Col <> 2 And .Col <> 4 And .Row > 0 And .TextMatrix(.Row, 1) = "" And MyFormAddEditMode = AddMode Then
            .Select .Row, .Col
            .EditCell
        
        ElseIf .Col <> 2 And .Col <> 4 And .Row > 0 And .TextMatrix(.Row, 1) <> "" And MyFormAddEditMode = EditMode Then
            .Select .Row, .Col
            .EditCell
        End If
    End With
    
End Sub

Private Sub vsGood_Click()
    With vsGood
        If .Row > 0 And .Col = 0 Then
            .Select .Row, .Col
            .EditCell
        End If
    End With
End Sub

Private Sub vsGood_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 32 And Shift = 0 Then
        With vsGood
            If .Rows > 1 And .Row > 0 Then
                If Val(.TextMatrix(.Row, 0)) = 0 Then
                    .TextMatrix(.Row, 0) = -1
                Else
                    .TextMatrix(.Row, 0) = 0
                End If
            End If
        End With
    End If
End Sub
