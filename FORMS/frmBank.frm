VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{75D4F767-8785-11D3-93AD-0000832EF44D}#2.14#0"; "FAST2001.ocx"
Begin VB.Form frmBank 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6420
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4725
   Icon            =   "frmBank.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   6420
   ScaleWidth      =   4725
   Begin VB.TextBox Text1 
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
      Height          =   450
      Left            =   240
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   720
      Width           =   3855
   End
   Begin VSFlex7LCtl.VSFlexGrid vsfgBank 
      Height          =   4755
      Left            =   120
      TabIndex        =   1
      Top             =   1320
      Width           =   4515
      _cx             =   7964
      _cy             =   8387
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
      AllowUserResizing=   0
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
      Left            =   3360
      Top             =   0
      Width           =   1380
      _ExtentX        =   2434
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
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   " ⁄—Ì› »«‰ò"
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
      Left            =   1080
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   0
      Width           =   2175
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "‰«„"
      BeginProperty Font 
         Name            =   "B Homa"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   4200
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   720
      Width           =   375
   End
End
Attribute VB_Name = "frmBank"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private MyFormAddEditMode As EnumAddEditMode
Dim Parameter() As Parameter

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

    If ClsFormAccess.frmBank = False Then
        Unload Me
        Exit Sub
    End If
    
    CenterCenter Me
    
    VarActForm = Me.Name
    
    With vsfgBank
        .Cols = 3
        .TextMatrix(0, 1) = "‰«„ »«‰ò"

        .ColHidden(2) = True

        .FixedAlignment(-1) = flexAlignCenterCenter
        .ColAlignment(-1) = flexAlignRightCenter

        .ColWidth(0) = 510
        .ColWidth(1) = 3600
    End With

    MyFormAddEditMode = ViewMode
    DefaultSetting
    SetFirstToolBar
End Sub

Private Sub Form_Unload(Cancel As Integer)
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

    If MyFormAddEditMode <> 0 Then
        Cancel
    End If
    On Error GoTo ErrHandler
    ReDim Parameter(0) As Parameter
    Parameter(0) = GenerateInputParameter("@intId", adInteger, 4, Text1.Tag)
    RunParametricStoredProcedure "Delete_tBank_By_tintBank", Parameter
    
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
    
    Set Rst = RunStoredProcedure2RecordSet("Get_All_tBanks")
    
    With vsfgBank
        .Rows = 1
        If Not (Rst.BOF = True And Rst.EOF = True) Then
            While Rst.EOF <> True
                .Rows = .Rows + 1

                .TextMatrix(.Rows - 1, 0) = .Rows - 1
                .TextMatrix(.Rows - 1, 1) = Rst!nvcBankName
                .TextMatrix(.Rows - 1, 2) = Rst!tintBank
                Rst.MoveNext
            Wend
        End If
    
    End With
    
    If Rst.State = 1 Then Rst.Close
     
    Dim Obj As Object
    For Each Obj In Me
        If TypeOf Obj Is TextBox Then
            Obj.Text = ""
            Obj.Tag = 0
        ElseIf TypeOf Obj Is ComboBox Then
            Obj.ListIndex = 0
        ElseIf TypeOf Obj Is OptionButton Then
            Obj.Value = False
        ElseIf TypeOf Obj Is CheckBox Then
            Obj.Value = vbUnchecked
        End If
    Next Obj
    
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
    
End Sub
Public Sub ChangeLanguage()

    Select Case clsStation.Language
    
        Case Farsi
        
        Case English
        
    End Select
    
End Sub

Public Sub Update()
    Dim i As Integer
    ReDim Parameter(1) As Parameter
    Dim Result As Integer
    Dim Obj As Object

    If Trim$(Text1.Text) = "" Then
            frmMsg.fwlblMsg.Caption = "·ÿ›« «ÿ·«⁄«  —« ò«„· Ê«—œ ‰„«ÌÌœ"
            frmMsg.fwBtn(0).ButtonType = flwButtonOk
            frmMsg.fwBtn(0).Caption = "ﬁ»Ê·"
            frmMsg.Show vbModal
            
            Text1.SetFocus
            
            Exit Sub

    End If

    For i = 1 To vsfgBank.Rows - 1
        If vsfgBank.Row <> i Then
            If Trim$(vsfgBank.TextMatrix(i, 1)) = Trim$(Text1.Text) Then
                frmMsg.fwlblMsg.Caption = "ﬁ»·« À»  ‘œÂ «” "
                frmMsg.fwBtn(0).ButtonType = flwButtonOk
                frmMsg.fwBtn(0).Caption = "ﬁ»Ê·"
                frmMsg.Show vbModal
                Exit Sub
            End If
        End If
    Next i
    
    Select Case MyFormAddEditMode
    
        Case AddMode
            Parameter(0) = GenerateInputParameter("@nvcBankName", adVarWChar, 25, Trim(Text1.Text))
            Parameter(1) = GenerateOutputParameter("@intStatus", adInteger, 4)
            
            Result = RunParametricStoredProcedure("Insert_tBank", Parameter)
            
            If Parameter(1).Value <> -1 Then
                Text1.Tag = Parameter(1).Value
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
        
            ReDim Parameter(2) As Parameter
            
            Parameter(0) = GenerateInputParameter("@tintBank", adInteger, 4, Text1.Tag)
            Parameter(1) = GenerateInputParameter("@nvcBankName", adVarWChar, 25, Trim(Text1.Text))
            Parameter(2) = GenerateOutputParameter("@intStatus", adInteger, 4)
            
            Result = RunParametricStoredProcedure("Update_tBank", Parameter)

            If Parameter(2).Value <> -1 Then
            
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

End Sub

Private Sub vsfgBank_Click()
    
    With vsfgBank
        If .Row = 0 Then Exit Sub
        Text1.Tag = .TextMatrix(.Row, 2)
        Text1.Text = .TextMatrix(.Row, 1)

        MyFormAddEditMode = ViewMode
        SetFirstToolBar
    End With
    
End Sub
