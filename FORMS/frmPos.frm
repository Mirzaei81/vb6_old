VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{7AEDC602-D94C-11D1-BB7A-00E0290EA3C9}#1.0#0"; "ResizeKit.ocx"
Object = "{75D4F767-8785-11D3-93AD-0000832EF44D}#2.14#0"; "FAST2001.ocx"
Begin VB.Form frmPos 
   ClientHeight    =   6570
   ClientLeft      =   120
   ClientTop       =   510
   ClientWidth     =   9300
   Icon            =   "frmPos.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   6570
   ScaleWidth      =   9300
   Begin VB.TextBox txtPosSerialNo 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "B Homa"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   4680
      RightToLeft     =   -1  'True
      TabIndex        =   16
      Top             =   2160
      Width           =   2895
   End
   Begin VB.ComboBox cmbBank 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   240
      RightToLeft     =   -1  'True
      TabIndex        =   15
      Top             =   720
      Width           =   2895
   End
   Begin VB.ComboBox cboStations 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   4680
      RightToLeft     =   -1  'True
      TabIndex        =   13
      Top             =   720
      Width           =   2895
   End
   Begin VB.ComboBox cmbCommunication 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   4680
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   1680
      Width           =   2895
   End
   Begin VB.TextBox txtPosAddress 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   240
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   1680
      Width           =   2895
   End
   Begin VB.ComboBox CmbPosModel 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   4680
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   1200
      Width           =   2895
   End
   Begin VB.ComboBox cmbAccountNo 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   240
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   1200
      Width           =   2895
   End
   Begin VB.TextBox TxtAccountNo 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "B Homa"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   240
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   1200
      Width           =   2895
   End
   Begin VSFlex7LCtl.VSFlexGrid vsfgBank 
      Height          =   3555
      Left            =   120
      TabIndex        =   0
      Top             =   2880
      Width           =   9075
      _cx             =   16007
      _cy             =   6271
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
      BackColorFixed  =   -2147483633
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
      Rows            =   2
      Cols            =   2
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
   Begin FLWCtrls.FWLabel3D fwlblMode 
      Height          =   495
      Left            =   7800
      Top             =   120
      Width           =   1380
      _ExtentX        =   2434
      _ExtentY        =   873
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Titr"
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
      OleObjectBlob   =   "frmPos.frx":A4C2
      TabIndex        =   5
      Top             =   0
      Width           =   480
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "‘„«—Â ”—Ì«· ÅÊ“"
      BeginProperty Font 
         Name            =   "B Homa"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   7680
      RightToLeft     =   -1  'True
      TabIndex        =   17
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "‰«„ »«‰ò"
      BeginProperty Font 
         Name            =   "B Homa"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   3000
      TabIndex        =   14
      Top             =   720
      Width           =   1365
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "‰«„ «Ì” ê«Â"
      BeginProperty Font 
         Name            =   "B Homa"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   7560
      TabIndex        =   12
      Top             =   720
      Width           =   1365
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "¬œ—” ‘»òÂ"
      BeginProperty Font 
         Name            =   "B Homa"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   3240
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "‰ÕÊÂ « ’«·"
      BeginProperty Font 
         Name            =   "B Homa"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   7200
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   1680
      Width           =   1695
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "‘„«—Â Õ”«»"
      BeginProperty Font 
         Name            =   "B Homa"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   3240
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "‰Ê⁄ ÅÊ“ »«‰òÌ"
      BeginProperty Font 
         Name            =   "B Homa"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   7320
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   " ⁄—Ì› ÅÊ“ »«‰òÌ"
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
      Left            =   3360
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   0
      Width           =   2175
   End
End
Attribute VB_Name = "frmPos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private MyFormAddEditMode As EnumAddEditMode
Dim Parameter() As Parameter
Dim i As Long

Public Sub SetFirstToolBar()
    Dim i As Integer

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

    If ClsFormAccess.frmPos = False Then
        Unload Me
        Exit Sub
    End If
    
    CenterCenter Me
        
    Dim L_Rst As New ADODB.Recordset
    With vsfgBank
        .Cols = 10
        .TextMatrix(0, 0) = "—œÌ›"
        .TextMatrix(0, 1) = "‘„«—Â ÅÊ“"
        .TextMatrix(0, 2) = "«Ì” ê«Â ÅÊ“"
        .TextMatrix(0, 3) = "‰«„ »«‰ò"
        .TextMatrix(0, 4) = "‘„«—Â Õ”«»"
        .TextMatrix(0, 5) = "‘„«—Â Õ”«»"
        .TextMatrix(0, 6) = "‰Ê⁄ ÅÊ“ "
        .TextMatrix(0, 7) = "‰ÕÊÂ « ’«· "
        .TextMatrix(0, 8) = "¬œ—” ‘»òÂ "
        .TextMatrix(0, 9) = "”—Ì«· ÅÊ“ "
        If clsArya.ExternalAccounting = True Then
            .ColHidden(4) = True
            cmbAccountNo.Visible = True
            TxtAccountNo.Visible = False
            cmbAccountNo.Clear
            Set L_Rst = Accounting.FillTafsiliAccountDll
            If Not (L_Rst.BOF = True And L_Rst.EOF = True) Then
                While L_Rst.EOF = False
                    cmbAccountNo.AddItem CStr(L_Rst.Fields("TafsiliName"))
                    cmbAccountNo.ItemData(cmbAccountNo.ListCount - 1) = Val(L_Rst.Fields("TafsiliId"))
                    L_Rst.MoveNext
                Wend
                cmbAccountNo.ListIndex = -1
            End If
            Set L_Rst = Accounting.FillTafsiliAccountDll
            .ColComboList(5) = .BuildComboList(L_Rst, "TafsiliName", "TafsiliId")
        
        Else
            .ColHidden(5) = True
            cmbAccountNo.Visible = False
            TxtAccountNo.Visible = True
        End If
        Dim Rst As New ADODB.Recordset
        
        Set Rst = RunStoredProcedure2RecordSet("Get_All_PosType")
        .ColComboList(6) = .BuildComboList(Rst, "PosName", "PosTypeId")
        
        Set Rst = RunStoredProcedure2RecordSet("Get_Pc_Stations")
        .ColComboList(2) = .BuildComboList(Rst, "Description", "StationId")
        
        Set Rst = RunStoredProcedure2RecordSet("Get_All_tBanks")
        .ColComboList(3) = .BuildComboList(Rst, "nvcBankName", "tintBank")
        
        
        .FixedAlignment(-1) = flexAlignCenterCenter
        .ColAlignment(-1) = flexAlignCenterCenter

'        .ColComboList(1) = "#1;ÅÊ“ »«‰òÌ ¬”«‰ Å—œ«Œ |#2;ÅÊ“ »«‰òÌ Å«”«—ê«œ|#3;ÅÊ“ »«‰òÌ „· |#4;ÅÊ“ »«‰òÌ ’«œ—« |#5;ÅÊ“ »«‰òÌ «ﬁ ’«œ ‰ÊÌ‰|#6;ÅÊ“ »«‰òÌ ”«„«‰"
        
'        CmbPosModel.Clear
'        CmbPosModel.AddItem "ÅÊ“ »«‰òÌ ¬”«‰ Å—œ«Œ "
'        CmbPosModel.ItemData(0) = 1
'        CmbPosModel.AddItem "ÅÊ“ »«‰òÌ Å«”«—ê«œ"
'        CmbPosModel.ItemData(1) = 2
'        CmbPosModel.AddItem "ÅÊ“ »«‰òÌ „· "
'        CmbPosModel.ItemData(2) = 3
'        CmbPosModel.AddItem "ÅÊ“ »«‰òÌ ’«œ—« "
'        CmbPosModel.ItemData(3) = 4
'        CmbPosModel.AddItem "ÅÊ“ »«‰òÌ «ﬁ ’«œ ‰ÊÌ‰"
'        CmbPosModel.ItemData(4) = 5
'        CmbPosModel.AddItem "ÅÊ“ »«‰òÌ ”«„«‰"
'        CmbPosModel.ItemData(5) = 6
        
        Set Rst = RunStoredProcedure2RecordSet("Get_All_PosType")
        CmbPosModel.Clear
        If Not (Rst.EOF = True And Rst.BOF = True) Then
            While Not Rst.EOF
                i = i + 1
                CmbPosModel.AddItem Rst.Fields("PosName").Value
                CmbPosModel.ItemData(CmbPosModel.ListCount - 1) = Rst.Fields("PosTypeId").Value
                Rst.MoveNext
            Wend
        End If
        If Rst.State <> 0 Then Rst.Close
        
        Set Rst = RunStoredProcedure2RecordSet("Get_All_PosPort")
        .ColComboList(7) = .BuildComboList(Rst, CStr("PortName"), "PortId")
        Rst.Close
        Set Rst = RunStoredProcedure2RecordSet("Get_All_PosPort")
        cmbCommunication.Clear
        If Not (Rst.EOF = True And Rst.BOF = True) Then
            While Not Rst.EOF
                i = i + 1
                cmbCommunication.AddItem Rst.Fields("PortName").Value
                cmbCommunication.ItemData(cmbCommunication.ListCount - 1) = Rst.Fields("PortId").Value
                Rst.MoveNext
            Wend
        End If
        Rst.Close
        
        For i = 0 To .Cols - 1
           .ColWidth(i) = Val(GetSetting(strMainKey, Me.Name, "Col" & i))
           If .ColWidth(i) = 0 Then
               .ColWidth(i) = 1000       '
           End If
        Next i
    
    
    End With

    MyFormAddEditMode = ViewMode
    DefaultSetting
    SetFirstToolBar
    
    Set Rst = RunStoredProcedure2RecordSet("Get_Pc_Stations")

    i = 0
    cboStations.Clear
    If Not (Rst.EOF = True And Rst.BOF = True) Then
        While Not Rst.EOF
            i = i + 1
            cboStations.AddItem Rst.Fields("Description").Value
            cboStations.ItemData(cboStations.ListCount - 1) = Rst.Fields("StationID").Value
            Rst.MoveNext
        Wend
    End If
    If Rst.State <> 0 Then Rst.Close
    
    Set Rst = RunStoredProcedure2RecordSet("Get_All_tBanks")

    i = 0
    cmbBank.Clear
    If Not (Rst.EOF = True And Rst.BOF = True) Then
        While Not Rst.EOF
            i = i + 1
            cmbBank.AddItem Rst.Fields("nvcBankName").Value
            cmbBank.ItemData(cmbBank.ListCount - 1) = Rst.Fields("tintBank").Value
            Rst.MoveNext
        Wend
    End If
    If Rst.State <> 0 Then Rst.Close
    
    Set Rst = Nothing
    
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
    
    Set L_Rst = Nothing
        
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveSetting strMainKey, Me.Name, "Left", Me.Left
    SaveSetting strMainKey, Me.Name, "Top", Me.Top
    If PosConnection.State = adStateOpen Then PosConnection.Close: Set PosConnection = Nothing
    VarActForm = ""
    
End Sub
Public Sub ExitForm()

    Unload Me
    
End Sub
Public Sub Edit()
    If vsfgBank.Rows > 1 Then
        MyFormAddEditMode = EditMode 'Edit
        SetFirstToolBar
    End If
End Sub

Public Sub Delete()

    If vsfgBank.Rows < 2 Then Exit Sub
    If vsfgBank.Row < 1 Then Exit Sub

    On Error GoTo ErrHandler
    ReDim Parameter(0) As Parameter
    Parameter(0) = GenerateInputParameter("@PosId", adInteger, 4, Val(vsfgBank.TextMatrix(vsfgBank.Row, 1)))
    RunParametricStoredProcedure "Delete_tblPub_Pos", Parameter
    
    frmMsg.fwlblMsg.Caption = "»« „Ê›ﬁÌ  Õ–› ‘œ"
    frmMsg.fwBtn(0).Visible = False
    frmMsg.fwBtn(1).ButtonType = flwButtonOk
    frmMsg.fwBtn(1).Caption = "ﬁ»Ê·"
    frmMsg.Show vbModal
    
    DefaultSetting
Exit Sub
    
ErrHandler:
If err.Number = -2147217873 Then

    frmMsg.fwlblMsg.Caption = "„ «”›«‰Â ‘„« ﬁ«œ— »Â Õ–› ‰„Ì »«‘Ìœ"
    frmMsg.fwBtn(0).Visible = False
    frmMsg.fwBtn(1).ButtonType = flwButtonOk
    frmMsg.fwBtn(1).Caption = "ﬁ»Ê·"
    frmMsg.Show vbModal
End If
    
End Sub

Public Sub DefaultSetting()

    Dim Rst As New ADODB.Recordset
    
    Set Rst = RunStoredProcedure2RecordSet("Get_All_tblPub_Pos")
    
    With vsfgBank
        .Rows = 1
        If Not (Rst.BOF = True And Rst.EOF = True) Then
            While Rst.EOF <> True
                .Rows = .Rows + 1

                .TextMatrix(.Rows - 1, 0) = .Rows - 1
                .TextMatrix(.Rows - 1, 1) = Rst!PosId
                .TextMatrix(.Rows - 1, 2) = Rst!StationId
                .TextMatrix(.Rows - 1, 3) = IIf(IsNull(Rst!intBank), "", Rst!intBank)
                .TextMatrix(.Rows - 1, 4) = Rst!nvcAccountNo
                .TextMatrix(.Rows - 1, 5) = IIf(IsNull(Rst!AccountId), "", Rst!AccountId)
                .TextMatrix(.Rows - 1, 6) = Rst!POSType
                .TextMatrix(.Rows - 1, 7) = IIf(IsNull(Rst!ComunicationType), "", Rst!ComunicationType)
                .TextMatrix(.Rows - 1, 8) = IIf(IsNull(Rst!PosAddress), "", Rst!PosAddress)
                .TextMatrix(.Rows - 1, 9) = IIf(IsNull(Rst!nvcPosSerialNo), "", Rst!nvcPosSerialNo)
                Rst.MoveNext
            Wend
        End If
    
    End With
    
    If Rst.State = 1 Then Rst.Close
     
'    Dim Obj As Object
'    For Each Obj In Me
'        If TypeOf Obj Is TextBox Then
'            Obj.Text = ""
'            Obj.Tag = 0
'        ElseIf TypeOf Obj Is ComboBox Then
'            Obj.ListIndex = 0
'        ElseIf TypeOf Obj Is OptionButton Then
'            Obj.Value = False
'        ElseIf TypeOf Obj Is CheckBox Then
'            Obj.Value = vbUnchecked
'        End If
'    Next Obj
    
    Set Rst = Nothing
    
End Sub
Public Sub Add()
    
    Cancel
    MyFormAddEditMode = AddMode
    
    SetFirstToolBar
    
End Sub

Public Sub Cancel()

    MyFormAddEditMode = ViewMode
    SetFirstToolBar
    DefaultSetting
    CmbPosModel.ListIndex = -1
    cmbAccountNo.ListIndex = -1
    TxtAccountNo = ""
    cboStations.ListIndex = -1
    cmbBank.ListIndex = -1
    cmbCommunication.ListIndex = -1
    txtPosAddress = ""
    txtPosSerialNo = ""

End Sub
Public Sub ChangeLanguage()

    Select Case clsStation.Language
    
        Case Farsi
        
        Case English
        
    End Select
    
End Sub

Public Sub Update()
    
    On Error GoTo ErrHandler
    Dim i As Integer
    Dim Result As Long
    Dim Obj As Object

    If clsArya.ExternalAccounting = True And cmbAccountNo.ListIndex = -1 Then ShowDisMessage "Õ”«» »«‰òÌ «‰ Œ«» ‰‘œÂ", 1500: Exit Sub
    If Trim$(CmbPosModel.Text) = "" Or cboStations.ListIndex = -1 Or cmbBank.ListIndex = -1 Or (Trim$(TxtAccountNo.Text) = "" And clsArya.ExternalAccounting = False) Then
            frmMsg.fwlblMsg.Caption = "·ÿ›« «ÿ·«⁄«  —« ò«„· Ê«—œ ‰„«ÌÌœ"
            frmMsg.fwBtn(0).ButtonType = flwButtonOk
            frmMsg.fwBtn(0).Caption = "ﬁ»Ê·"
            frmMsg.Show vbModal
            
            Exit Sub

    End If
    
'''    For i = 1 To vsfgBank.Rows - 1
'''        If vsfgBank.Row <> i Then
'''            If (Trim$(vsfgBank.TextMatrix(i, 2)) = cboStations.ItemData(cboStations.ListIndex) And clsArya.ExternalAccounting = False) Or (Trim$(vsfgBank.TextMatrix(i, 4)) = Trim$(cmbAccountNo.Text) And clsArya.ExternalAccounting = True) Then
'''                frmMsg.fwlblMsg.Caption = "ﬁ»·« À»  ‘œÂ «” "
'''                frmMsg.fwBtn(0).ButtonType = flwButtonOk
'''                frmMsg.fwBtn(0).Caption = "ﬁ»Ê·"
'''                frmMsg.fwBtn(0).Visible = False
'''                frmMsg.Show vbModal
'''                Exit Sub
'''            End If
'''        End If
'''    Next i
    
    Select Case MyFormAddEditMode
    
        Case AddMode
            ReDim Parameter(8) As Parameter
            Parameter(0) = GenerateInputParameter("@PosType", adInteger, 4, CmbPosModel.ItemData(CmbPosModel.ListIndex))
            Parameter(1) = GenerateInputParameter("@StationId", adInteger, 4, cboStations.ItemData(cboStations.ListIndex))
            Parameter(2) = GenerateInputParameter("@BankNo", adTinyInt, 16, cmbBank.ItemData(cmbBank.ListIndex))
            If clsArya.ExternalAccounting = False Then
                Parameter(3) = GenerateInputParameter("@nvcAccountNo", adVarWChar, 50, Trim(TxtAccountNo.Text))
                Parameter(4) = GenerateInputParameter("@AccountId", adInteger, 4, Null)
            Else
                Parameter(3) = GenerateInputParameter("@nvcAccountNo", adVarWChar, 50, Trim(cmbAccountNo.Text))
                Parameter(4) = GenerateInputParameter("@AccountId", adInteger, 4, cmbAccountNo.ItemData(cmbAccountNo.ListIndex))
            End If
            Parameter(5) = GenerateInputParameter("@CommunicationType", adInteger, 4, cmbCommunication.ItemData(cmbCommunication.ListIndex))
            Parameter(6) = GenerateInputParameter("@PosAddress", adVarWChar, 20, Trim(txtPosAddress.Text))
            Parameter(7) = GenerateInputParameter("@nvcPosSerialNo", adVarWChar, 20, Trim(txtPosSerialNo.Text))
            Parameter(8) = GenerateOutputParameter("@intStatus", adInteger, 4)
            
            Result = RunParametricStoredProcedure("Insert_tblPub_Pos", Parameter)
            
            If Result <> -1 Then
                frmMsg.fwlblMsg.Caption = "À»  «ÿ·«⁄«  ÃœÌœ »« „Ê›ﬁÌ  Å«Ì«‰ Ì«› "
                frmMsg.fwBtn(0).ButtonType = flwButtonOk
                frmMsg.fwBtn(0).Caption = "ﬁ»Ê·"
                frmMsg.fwBtn(1).Visible = False
                frmMsg.Show vbModal

                MyFormAddEditMode = ViewMode
                DefaultSetting
                SetFirstToolBar
            Else
                frmMsg.fwlblMsg.Caption = "«ÿ·«⁄«  ÃœÌœ À»  ‰‘œ. ·ÿ›« œÊ»«—Â ”⁄Ì ‰„«ÌÌœ"
                frmMsg.fwBtn(0).ButtonType = flwButtonOk
                frmMsg.fwBtn(0).Caption = "ﬁ»Ê·"
                frmMsg.fwBtn(1).Visible = False
                frmMsg.Show vbModal
            End If
            
        Case EditMode
        
            ReDim Parameter(9) As Parameter
            
            Parameter(0) = GenerateInputParameter("@PosId", adInteger, 4, Val(vsfgBank.TextMatrix(vsfgBank.Row, 1)))
            Parameter(1) = GenerateInputParameter("@StationId", adInteger, 4, cboStations.ItemData(cboStations.ListIndex))
            Parameter(2) = GenerateInputParameter("@BankNo", adTinyInt, 16, cmbBank.ItemData(cmbBank.ListIndex))
            If clsArya.ExternalAccounting = False Then
                Parameter(3) = GenerateInputParameter("@nvcAccountNo", adVarWChar, 50, Trim(TxtAccountNo.Text))
                Parameter(4) = GenerateInputParameter("@AccountId", adInteger, 4, Null)
            Else
                Parameter(3) = GenerateInputParameter("@nvcAccountNo", adVarWChar, 50, Trim(cmbAccountNo.Text))
                Parameter(4) = GenerateInputParameter("@AccountId", adInteger, 4, cmbAccountNo.ItemData(cmbAccountNo.ListIndex))
            End If
            Parameter(5) = GenerateInputParameter("@PosType", adInteger, 4, CmbPosModel.ItemData(CmbPosModel.ListIndex))
            Parameter(6) = GenerateInputParameter("@CommunicationType", adInteger, 4, cmbCommunication.ItemData(cmbCommunication.ListIndex))
            Parameter(7) = GenerateInputParameter("@PosAddress", adVarWChar, 20, Trim(txtPosAddress.Text))
            Parameter(8) = GenerateInputParameter("@nvcPosSerialNo", adVarWChar, 20, Trim(txtPosSerialNo.Text))
            Parameter(9) = GenerateOutputParameter("@intStatus", adInteger, 4)
            
            Result = RunParametricStoredProcedure("Update_tblPub_Pos", Parameter)

            If Result <> -1 Then
            
                frmMsg.fwlblMsg.Caption = " €ÌÌ— «ÿ·«⁄«  »« „Ê›ﬁÌ  Å«Ì«‰ Ì«› "
                frmMsg.fwBtn(0).ButtonType = flwButtonOk
                frmMsg.fwBtn(0).Caption = "ﬁ»Ê·"
                frmMsg.fwBtn(1).Visible = False
                frmMsg.Show vbModal
                
                MyFormAddEditMode = ViewMode
                DefaultSetting
                SetFirstToolBar
            Else
            
                frmMsg.fwlblMsg.Caption = "„ «”›«‰Â «ÿ·«⁄«   €ÌÌ— ‰Ì«› . ·ÿ›« œÊ»«—Â ”⁄Ì ‰„«ÌÌœ"
                frmMsg.fwBtn(0).ButtonType = flwButtonOk
                frmMsg.fwBtn(0).Caption = "ﬁ»Ê·"
                frmMsg.fwBtn(1).Visible = False
                frmMsg.Show vbModal
                
            End If
            
    End Select
Exit Sub
ErrHandler:
    ShowDisMessage err.Description, 2000
    
End Sub


Private Sub ResizeKit1_ExitResize(ByVal XScale As Double, ByVal YScale As Double)
    If formloadFlag = True Then
        SaveSetting strMainKey, Me.Name, "Height", Me.Height
        SaveSetting strMainKey, Me.Name, "Width", Me.Width
    End If

End Sub

Private Sub vsfgBank_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    
    With vsfgBank
    For i = 0 To vsfgBank.Cols - 1
        SaveSetting strMainKey, Me.Name, "Col" & i, .ColWidth(i)
    Next
    End With

End Sub

Private Sub vsfgBank_Click()
    If MyFormAddEditMode <> EditMode Then Exit Sub
    With vsfgBank
        Dim i, ii As Long
        If .Row = 0 Then Exit Sub
        For i = 0 To cboStations.ListCount - 1
            If cboStations.ItemData(i) = .TextMatrix(.Row, 2) Then
               Me.cboStations.ListIndex = i
               Exit For
            End If
        Next i
        For i = 0 To CmbPosModel.ListCount - 1
            If CmbPosModel.ItemData(i) = .TextMatrix(.Row, 6) Then
               Me.CmbPosModel.ListIndex = i
               Exit For
            End If
        Next i
        For i = 0 To cmbBank.ListCount - 1
            If cmbBank.ItemData(i) = Val(.TextMatrix(.Row, 3)) Then
               Me.cmbBank.ListIndex = i
               Exit For
            End If
        Next i
        TxtAccountNo.Text = .TextMatrix(.Row, 4)
        cmbAccountNo.ListIndex = -1
        If clsArya.ExternalAccounting = True And .ValueMatrix(.Row, 5) > 0 Then
            For ii = 0 To cmbAccountNo.ListCount - 1
                If cmbAccountNo.ItemData(ii) = .ValueMatrix(.Row, 5) Then
                    cmbAccountNo.ListIndex = ii
                    Exit For
                End If
            Next ii
        End If
        For i = 0 To cmbCommunication.ListCount - 1
            If cmbCommunication.ItemData(i) = .ValueMatrix(.Row, 7) Then
               Me.cmbCommunication.ListIndex = i
               Exit For
            End If
        Next i
        txtPosAddress.Text = .TextMatrix(.Row, 8)
        txtPosSerialNo.Text = .TextMatrix(.Row, 9)
''        MyFormAddEditMode = ViewMode
''        SetFirstToolBar
    End With
    
End Sub
