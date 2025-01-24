VERSION 5.00
Object = "{75D4F767-8785-11D3-93AD-0000832EF44D}#2.14#0"; "FAST2001.ocx"
Begin VB.Form frmPartition 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5355
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5400
   Icon            =   "frmPartition.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5355
   ScaleWidth      =   5400
   Begin VB.Frame Frame1 
      Height          =   4575
      Left            =   0
      TabIndex        =   0
      Top             =   600
      Width           =   5295
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
         Left            =   600
         RightToLeft     =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   240
         Width           =   2565
      End
      Begin VB.Frame Frame3 
         Caption         =   "ÅÌ‘ ›—÷ œ—’œ „Ì“Â«Ì —“—Ê"
         BeginProperty Font 
            Name            =   "B Yekan"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   1200
         Left            =   600
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Top             =   3360
         Width           =   4335
         Begin FLWCtrls.FWNumericTextBox FWNumericReserveService 
            Height          =   495
            Left            =   1200
            TabIndex        =   8
            Top             =   360
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   873
            Value           =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "ÅÌ‘ ›—÷ œ—’œ ”—ÊÌ”"
         BeginProperty Font 
            Name            =   "B Yekan"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   1200
         Left            =   600
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   2040
         Width           =   4335
         Begin FLWCtrls.FWNumericTextBox FWNumericService 
            Height          =   495
            Left            =   1200
            TabIndex        =   6
            Top             =   360
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   873
            Value           =   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin VB.ComboBox cmbExistPartitons 
         BeginProperty Font 
            Name            =   "Arial"
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
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1560
         Width           =   2565
      End
      Begin VB.TextBox txtPartition 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   600
         TabIndex        =   2
         Top             =   960
         Width           =   2565
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
         Left            =   3840
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   240
         Width           =   975
      End
      Begin VB.Label lblExistPartitions 
         Alignment       =   1  'Right Justify
         Caption         =   "»Œ‘ Â«Ì „ÊÃÊœ"
         BeginProperty Font 
            Name            =   "B Yekan"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3120
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   1560
         Width           =   1815
      End
      Begin VB.Label lblPartitionName 
         Alignment       =   1  'Right Justify
         Caption         =   "‰«„ »Œ‘"
         BeginProperty Font 
            Name            =   "B Yekan"
            Size            =   11.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3960
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   960
         Width           =   855
      End
   End
   Begin FLWCtrls.FWLabel3D fwlblMode 
      Height          =   525
      Left            =   4080
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
   Begin FLWCtrls.FWLabel fwlblPartition 
      Height          =   495
      Left            =   0
      Top             =   0
      Width           =   5325
      _ExtentX        =   9393
      _ExtentY        =   873
      Enabled         =   -1  'True
      Caption         =   " ⁄—Ì› »Œ‘"
      FirstColor      =   9412754
      SecondColor     =   14215660
      Angle           =   0
      ForeColor       =   7362318
      BackColor       =   12640511
      FontName        =   "B Yekan"
      FontBold        =   -1  'True
      FontItalic      =   -1  'True
      FontSize        =   15.75
      Alignment       =   2
      Picture         =   "frmPartition.frx":A4C2
   End
End
Attribute VB_Name = "frmPartition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private MyFormAddEditMode As EnumAddEditMode
Dim Parameter() As Parameter
Public Sub ChangeLanguage()
    
    DefaultSetting
    
End Sub


Public Sub SetFirstToolBar()
    
    AllButton vbOff, True
    
    mdifrm.Toolbar1.Buttons(1).Enabled = True   'Home
    mdifrm.Toolbar1.Buttons(2).Enabled = True   'PageUp
    mdifrm.Toolbar1.Buttons(3).Enabled = True   'PageDown
    mdifrm.Toolbar1.Buttons(4).Enabled = True   'End
        
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
        Frame1.Enabled = False
        
    ElseIf MyFormAddEditMode = AddMode Then    'Add Mode
    
        mdifrm.Toolbar1.Buttons(6).Enabled = False  'Add
        mdifrm.Toolbar1.Buttons(7).Enabled = False  'Edit
        mdifrm.Toolbar1.Buttons(8).Enabled = True   'Enter
        mdifrm.Toolbar1.Buttons(9).Enabled = True   'Esc
        Frame1.Enabled = True

    ElseIf MyFormAddEditMode = EditMode Then     'Edit
        
        mdifrm.Toolbar1.Buttons(6).Enabled = True  'Add
        mdifrm.Toolbar1.Buttons(7).Enabled = False  'Edit
        mdifrm.Toolbar1.Buttons(8).Enabled = True   'Enter
        mdifrm.Toolbar1.Buttons(9).Enabled = True   'Esc
        Frame1.Enabled = True
'        Frame1.Visible = False
        
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
   
    CenterCenter Me
    If Not (Val(strCategory) >= 0 And Val(strCategory) <= 7) Then
       Frame2.Visible = False
       Frame3.Visible = False
    End If
    
    VarActForm = Me.Name
    
    MyFormAddEditMode = AddMode
    FillBranch
    DefaultSetting
    SetFirstToolBar

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

Private Sub Form_Unload(Cancel As Integer)
    VarActForm = ""
    If PosConnection.State = adStateOpen Then PosConnection.Close: Set PosConnection = Nothing
End Sub
Public Sub ExitForm()

    Unload Me
    
End Sub
Public Sub Edit()
    MyFormAddEditMode = EditMode
    SetFirstToolBar
End Sub


Public Sub FirstKey()
    ReDim Parameter(2) As Parameter
    Dim Rst As New ADODB.Recordset
    
    Parameter(0) = GenerateInputParameter("@CurrentPartitionID", adInteger, 4, Val(txtPartition.Tag))
    Parameter(1) = GenerateInputParameter("@Direction", adInteger, 4, EnumDirection.FirstRecord)
    Parameter(2) = GenerateInputParameter("@Branch", adInteger, 4, CurrentBranch)
    
    Set Rst = RunParametricStoredProcedure2Rec("NavigateInPartitions", Parameter)
    If Rst.State = 1 Then
        If Not (Rst.EOF = True And Rst.BOF = True) Then
            GetRecrdsetDetail Rst
        End If
    End If
    Set Rst = Nothing
    
    MyFormAddEditMode = ViewMode
    SetFirstToolBar
End Sub

Public Sub PreviousKey()
    ReDim Parameter(2) As Parameter
    Dim Rst As New ADODB.Recordset
    
    Parameter(0) = GenerateInputParameter("@CurrentPartitionID", adInteger, 4, Val(txtPartition.Tag))
    Parameter(1) = GenerateInputParameter("@Direction", adInteger, 4, EnumDirection.PreviousRecord)
    Parameter(2) = GenerateInputParameter("@Branch", adInteger, 4, CurrentBranch)
    
    Set Rst = RunParametricStoredProcedure2Rec("NavigateInPartitions", Parameter)
    If Rst.State = 1 Then
        If Not (Rst.EOF = True And Rst.BOF = True) Then
            GetRecrdsetDetail Rst
        End If
    End If
    Set Rst = Nothing
    
    MyFormAddEditMode = ViewMode
    SetFirstToolBar
End Sub

Public Sub NextKey()
    ReDim Parameter(2) As Parameter
    Dim Rst As New ADODB.Recordset
    
    Parameter(0) = GenerateInputParameter("@CurrentPartitionID", adInteger, 4, Val(txtPartition.Tag))
    Parameter(1) = GenerateInputParameter("@Direction", adInteger, 4, EnumDirection.NextRecord)
    Parameter(2) = GenerateInputParameter("@Branch", adInteger, 4, CurrentBranch)
    
    Set Rst = RunParametricStoredProcedure2Rec("NavigateInPartitions", Parameter)
    If Rst.State = 1 Then
        If Not (Rst.EOF = True And Rst.BOF = True) Then
            GetRecrdsetDetail Rst
        End If
    End If
    Set Rst = Nothing
    
    MyFormAddEditMode = ViewMode
    SetFirstToolBar
End Sub

Public Sub LastKey()
    ReDim Parameter(2) As Parameter
    Dim Rst As New ADODB.Recordset
    
    Parameter(0) = GenerateInputParameter("@CurrentPartitionID", adInteger, 4, Val(txtPartition.Tag))
    Parameter(1) = GenerateInputParameter("@Direction", adInteger, 4, EnumDirection.LastRecord)
    Parameter(2) = GenerateInputParameter("@Branch", adInteger, 4, CurrentBranch)
    
    Set Rst = RunParametricStoredProcedure2Rec("NavigateInPartitions", Parameter)
    If Rst.State = 1 Then
        If Not (Rst.EOF = True And Rst.BOF = True) Then
            GetRecrdsetDetail Rst
        End If
    End If
    Set Rst = Nothing
    
    MyFormAddEditMode = ViewMode
    SetFirstToolBar
End Sub
Private Sub GetRecrdsetDetail(tempRst As ADODB.Recordset)

    DefaultSetting
    
    If tempRst.EOF = True And tempRst.BOF = True Then Exit Sub
    
    txtPartition.Text = tempRst.Fields("PartitionDescription").Value
    txtPartition.Tag = tempRst.Fields("PartitionID").Value
    FWNumericService.Value = tempRst.Fields("DefaultServicePercent").Value
    FWNumericReserveService.Value = tempRst.Fields("ReserveServiceRate").Value
    
    If IsNull(tempRst.Fields("PartitionID").Value) = False Then
        
        For i = 1 To cmbExistPartitons.ListCount - 1
            If tempRst.Fields("PartitionID").Value = cmbExistPartitons.ItemData(i) Then
                cmbExistPartitons.ListIndex = i
                Exit For
            End If
        Next i
        
    Else
        cmbExistPartitons.ListIndex = 0
    
    
    End If
        
End Sub

Public Sub DefaultSetting()

    Dim Rst As New ADODB.Recordset
    
    ReDim Parameter(1) As Parameter
    Parameter(0) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
    Parameter(1) = GenerateInputParameter("@Branch", adInteger, 4, CurrentBranch)
    Set Rst = RunParametricStoredProcedure2Rec("Get_All_tPartitions", Parameter)
    
    cmbExistPartitons.Clear
    cmbExistPartitons.AddItem ""
    cmbExistPartitons.ItemData(cmbExistPartitons.ListCount - 1) = 0
    
    If Not (Rst.EOF = True And Rst.BOF = True) Then
        i = 1
        While Rst.EOF <> True
            cmbExistPartitons.AddItem Rst.Fields("PartitionDescription").Value
            cmbExistPartitons.ItemData(cmbExistPartitons.ListCount - 1) = Rst.Fields("PartitionID").Value
            Rst.MoveNext
        Wend
    End If
    
    If Rst.State <> 0 Then Rst.Close
    FWNumericService.Value = 0
    FWNumericReserveService.Value = 0
    Dim Obj As Object
    For Each Obj In Me
        If TypeOf Obj Is TextBox Then
            Obj.Text = ""
            Obj.Tag = 0
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

    MyFormAddEditMode = AddMode
    SetFirstToolBar
    DefaultSetting
    
End Sub
Public Sub Update()
    
    Dim Result As Integer
    Dim Obj As Object
    
    If Trim(txtPartition.Text) = "" Then
            
            frmMsg.fwlblMsg.Caption = "·ÿ›« ‰«„ »Œ‘ —« Ê«—œ ‰„«ÌÌœ"
            frmMsg.fwBtn(0).ButtonType = flwButtonOk
            frmMsg.fwBtn(0).Caption = "ﬁ»Ê·"
            frmMsg.Show vbModal
            Exit Sub

    End If
    
    
    
    Select Case MyFormAddEditMode
        Case AddMode
            ReDim Parameter(5) As Parameter
            Parameter(0) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
            Parameter(1) = GenerateInputParameter("@ServicePercentDefault", adInteger, 4, FWNumericService.Value)
            Parameter(2) = GenerateInputParameter("@PartitionName", adVarWChar, 50, txtPartition.Text)
            Parameter(3) = GenerateInputParameter("@ReserveServiceRate", adInteger, 4, FWNumericReserveService.Value)
            Parameter(4) = GenerateInputParameter("@Branch", adInteger, 4, CurrentBranch)
            Parameter(5) = GenerateOutputParameter("@PartitionID", adInteger, 4)
            
            On Error GoTo ErrHandler
            
            Result = RunParametricStoredProcedure("InsertPartition", Parameter)
            On Error GoTo 0
            
            txtPartition.Tag = Parameter(1).Value
            frmMsg.fwlblMsg.Caption = "À»  «ÿ·«⁄«  ÃœÌœ »« „Ê›ﬁÌ  Å«Ì«‰ Ì«› "
            frmMsg.fwBtn(0).ButtonType = flwButtonOk
            frmMsg.fwBtn(0).Caption = "ﬁ»Ê·"
            frmMsg.fwBtn(1).Visible = False
            frmMsg.Show vbModal
            
            Add
            
        Case EditMode
        
            ReDim Parameter(5) As Parameter
            
            Parameter(0) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
            Parameter(1) = GenerateInputParameter("@ServicePercentDefault", adInteger, 4, FWNumericService.Value)
            Parameter(2) = GenerateInputParameter("@PartitionName", adVarWChar, 50, txtPartition.Text)
            Parameter(3) = GenerateInputParameter("@PartitionID", adInteger, 4, cmbExistPartitons.ItemData(cmbExistPartitons.ListIndex))
            Parameter(4) = GenerateInputParameter("@ReserveServiceRate", adInteger, 4, FWNumericReserveService.Value)
            Parameter(5) = GenerateInputParameter("@Branch", adInteger, 4, CurrentBranch)
            
            On Error GoTo ErrHandler
            Result = RunParametricStoredProcedure("UpdatePartition", Parameter)
            On Error GoTo 0
            
            frmMsg.fwlblMsg.Caption = " €ÌÌ— «ÿ·«⁄«  »« „Ê›ﬁÌ  Å«Ì«‰ Ì«› "
            frmMsg.fwBtn(0).ButtonType = flwButtonOk
            frmMsg.fwBtn(0).Caption = "ﬁ»Ê·"
            frmMsg.fwBtn(1).Visible = False
            frmMsg.Show vbModal
            
            Add
            
    End Select

ErrHandler:
    Select Case err.Number
        Case -2147217873
            frmMsg.fwlblMsg.Caption = "À»  «‰Ã«„ ‰‘œ" + vbCrLf + "«ÿ·«⁄«   ò—«—Ì „Ì »«‘œ"
            frmMsg.fwBtn(0).ButtonType = flwButtonOk
            frmMsg.fwBtn(0).Caption = "ﬁ»Ê·"
            frmMsg.fwBtn(1).Visible = False
            frmMsg.Show vbModal
        Case Else
        
    End Select
    
End Sub
'*********************************************************************************
Public Sub Delete()
    
    Select Case clsStation.Language
        Case 0
            frmMsg.fwlblMsg.Caption = "¬Ì« „ÿ„∆‰Ìœ „Ì ŒÊ«ÂÌœ " & "'" & cmbExistPartitons.List(cmbExistPartitons.ListIndex) & "'" & " —« Õ–› ﬂ‰Ìœø"
            frmMsg.fwBtn(0).Caption = "»·Ì"
            frmMsg.fwBtn(1).Caption = "ŒÌ—"
        Case 1
            frmMsg.fwlblMsg.Caption = "You are going to delete  "  '& vsGood.TextMatrix(vsGood.Row, 2) & "'" + vbNewLine + "Are you sure ?"
            frmMsg.fwBtn(0).Caption = "Yes"
            frmMsg.fwBtn(1).Caption = "No"
            frmMsg.fwlblMsg.Alignment = vbLeftJustify
    End Select
        
    frmMsg.Show vbModal
        
    If modgl.mvarMsgIdx = vbNo Then
        Exit Sub
    End If
    
    ReDim Parameter(1) As Parameter
    Parameter(0) = GenerateInputParameter("@PartitionID", adInteger, 4, cmbExistPartitons.ItemData(cmbExistPartitons.ListIndex))
    Parameter(1) = GenerateInputParameter("@Branch", adInteger, 4, CurrentBranch)
    RunParametricStoredProcedure "DeletePartition", Parameter
    
    Select Case clsStation.Language
        Case 0
            frmMsg.fwlblMsg.Caption = "Õ–› Ê«Õœ „Ê—œ ‰Ÿ— »« „Ê›ﬁÌ  «‰Ã«„ ‘œ"
            frmMsg.fwBtn(0).Caption = "ﬁ»Ê·"
            frmMsg.fwBtn(1).Visible = False
        Case 1
            frmMsg.fwlblMsg.Caption = "The partition has deleted successfuly"
            frmMsg.fwBtn(0).Caption = "Yes"
            frmMsg.fwBtn(1).Visible = False
            frmMsg.fwlblMsg.Alignment = vbLeftJustify
    End Select
        
    frmMsg.Show vbModal
End Sub

