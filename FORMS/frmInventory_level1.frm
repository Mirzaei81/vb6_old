VERSION 5.00
Object = "{7AEDC602-D94C-11D1-BB7A-00E0290EA3C9}#1.0#0"; "ResizeKit.ocx"
Object = "{75D4F767-8785-11D3-93AD-0000832EF44D}#2.14#0"; "FAST2001.ocx"
Begin VB.Form frmInventory_level1 
   ClientHeight    =   6645
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7335
   Icon            =   "frmInventory_level1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   6645
   ScaleWidth      =   7335
   Begin VB.Frame StoreDataUpdate 
      Height          =   4095
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   2400
      Width           =   2295
      Begin VB.CommandButton cmdInventoryGood_Add 
         BackColor       =   &H000000C0&
         Caption         =   "«÷«›Â ò—œ‰ ò«·«Â«Ì   ê—ÊÂÂ« »Â «‰»«—"
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1215
         Left            =   360
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   1560
         Width           =   1575
      End
      Begin VB.CommandButton cmdInventoryGood_Delete 
         BackColor       =   &H000000C0&
         Caption         =   " Õ–› ﬂ·ÌÂ ò«·«Â« «“ «‰»«—"
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   360
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   3000
         Width           =   1575
      End
      Begin VB.ComboBox cmbSalMali 
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
         Left            =   360
         RightToLeft     =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "”«· „«·Ì"
         BeginProperty Font 
            Name            =   "B Traffic"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1320
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.ComboBox cmbBranch 
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
      Left            =   2640
      RightToLeft     =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   840
      Width           =   3195
   End
   Begin VB.TextBox txtGroupName 
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
      Height          =   435
      Left            =   2640
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   2040
      Visible         =   0   'False
      Width           =   3195
   End
   Begin VB.ListBox lstGoodLevel1 
      BeginProperty Font 
         Name            =   "Nazanin"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3570
      Left            =   2640
      RightToLeft     =   -1  'True
      Style           =   1  'Checkbox
      TabIndex        =   1
      Top             =   2520
      Width           =   3195
   End
   Begin VB.ComboBox cmbInventory 
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
      Left            =   2640
      RightToLeft     =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   1440
      Width           =   3195
   End
   Begin FLWCtrls.FWLabel3D fwlblMode 
      Height          =   525
      Left            =   5640
      Top             =   0
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   926
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
   Begin RESIZEKITLibCtl.ResizeKit ResizeKit1 
      Height          =   480
      Left            =   0
      OleObjectBlob   =   "frmInventory_level1.frx":A4C2
      TabIndex        =   11
      Top             =   0
      Width           =   480
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
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
      Left            =   6000
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   840
      Width           =   975
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   " ⁄—Ì› ê—ÊÂÂ« œ— «‰»«—Â«"
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
      Left            =   2040
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   0
      Width           =   2895
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "‰«„ ê—ÊÂ ò«·«Â«"
      BeginProperty Font 
         Name            =   "B Homa"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   5760
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   2520
      Width           =   1305
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "‰«„ «‰»«—"
      BeginProperty Font 
         Name            =   "B Homa"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   5880
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   1440
      Width           =   1065
   End
End
Attribute VB_Name = "frmInventory_level1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MyFormAddEditMode As EnumAddEditMode
Dim Parameter() As Parameter


Private Sub cmbBranch_Click()
    DefaultSetting
End Sub

Private Sub cmbInventory_Click()
    FillLevel1
    MyFormAddEditMode = EditMode
    SetFirstToolBar
End Sub
Private Sub FillLevel1()
   If cmbInventory.ListIndex = -1 Then Exit Sub
    For i = 0 To lstGoodLevel1.ListCount - 1
         lstGoodLevel1.Selected(i) = False
    Next i
   
    ReDim Parameter(1) As Parameter
    Parameter(0) = GenerateInputParameter("@Branch", adInteger, 4, cmbBranch.ItemData(cmbBranch.ListIndex))
    Parameter(1) = GenerateInputParameter("@InventoryNo", adInteger, 4, cmbInventory.ItemData(cmbInventory.ListIndex))
    Set Rst = RunParametricStoredProcedure2Rec("Get_tInventory_Level1", Parameter)
        
    
    If Not (Rst.EOF = True And Rst.BOF = True) Then
        While Rst.EOF <> True
            
            For i = 0 To lstGoodLevel1.ListCount - 1
                If Rst.Fields("Level1").Value = lstGoodLevel1.ItemData(i) Then
                      lstGoodLevel1.Selected(i) = True
                      Exit For
                      lstGoodLevel1.Selected(i) = False
             
                End If
            Next i
            Rst.MoveNext
        Wend
    End If

End Sub

Public Sub Update()
    
    Dim i As Integer
    
'    If lstGoodLevel1.SelCount = 0 Then Exit Sub
    Select Case MyFormAddEditMode
      Case EditMode
        Dim SelectedGroups As String
        For i = 0 To lstGoodLevel1.ListCount - 1
            If lstGoodLevel1.Selected(i) = True Then
                SelectedGroups = SelectedGroups & lstGoodLevel1.ItemData(i) & ","
            End If
        Next i
        If Len(SelectedGroups) > 0 Then
            SelectedGroups = Left(SelectedGroups, Len(SelectedGroups) - 1)
        Else
            SelectedGroups = ""
        End If
        If cmbInventory.ListIndex > -1 Then
            ReDim Parameter(2) As Parameter
            Parameter(0) = GenerateInputParameter("@Branch", adInteger, 4, cmbBranch.ItemData(cmbBranch.ListIndex))
            Parameter(1) = GenerateInputParameter("@InventoryNo", adInteger, 4, cmbInventory.ItemData(cmbInventory.ListIndex))
            Parameter(2) = GenerateInputParameter("@Level1Code", adVarWChar, 400, SelectedGroups)
            
            RunParametricStoredProcedure "Update_tInventory_Level1", Parameter
        
        End If
        frmDisMsg.lblMessage = "  €ÌÌ—«  «‰Ã«„ ‘œ "
        frmDisMsg.Timer1.Enabled = True
        frmDisMsg.Show vbModal
 End Select
End Sub
Public Function AddGoodstoInventory()
    
    Dim StrSelectedLevel1 As String
    StrSelectedLevel1 = ""
    For i = 0 To lstGoodLevel1.ListCount - 1
        If lstGoodLevel1.Selected(i) = True Then
            StrSelectedLevel1 = StrSelectedLevel1 & lstGoodLevel1.ItemData(i) & ","
            Exit For
        End If
    Next i
    If StrSelectedLevel1 = "" Then
        frmMsg.fwlblMsg.Caption = "‘„« »«Ìœ Õœ«ﬁ· Ìò ê—ÊÂ «‰ Œ«» ò‰Ìœ "
        frmMsg.fwBtn(0).ButtonType = flwButtonOk
        frmMsg.fwBtn(1).Visible = False
        frmMsg.fwBtn(0).Caption = "ﬁ»Ê·"
        frmMsg.Show vbModal
        Exit Function
    End If
        frmMsg.fwlblMsg.Caption = "¬Ì« »—«Ì «÷«›Â ò—œ‰ ‰«„ ò«·«Â« Ì «Ì‰ ê—ÊÂÂ« »Â «‰»«— «ÿ„Ì‰«‰ œ«—Ìœ"
        frmMsg.fwBtn(0).ButtonType = flwButtonOk
        frmMsg.fwBtn(1).ButtonType = flwButtonCancel
        frmMsg.fwBtn(0).Caption = "»·Ì"
        frmMsg.fwBtn(1).Caption = "ŒÌ—"
        frmMsg.Show vbModal
        If mvarMsgIdx = vbYes Then
            ReDim Parameter(3) As Parameter
    
            For i = 0 To lstGoodLevel1.ListCount - 1
                If lstGoodLevel1.Selected(i) = True Then
                
                    Parameter(0) = GenerateInputParameter("@IntInventoryNo", adInteger, 4, cmbInventory.ItemData(cmbInventory.ListIndex))
                    Parameter(1) = GenerateInputParameter("@Level1", adInteger, 4, lstGoodLevel1.ItemData(i))
                    Parameter(2) = GenerateInputParameter("@Branch", adInteger, 4, cmbBranch.ItemData(cmbBranch.ListIndex))
                    Parameter(3) = GenerateInputParameter("@AccountYear", adSmallInt, 2, Val(cmbSalMali.Text))
        
                    RunParametricStoredProcedure "Insert_tinventory_Good_All", Parameter
                End If
            Next i
            DefaultSetting
            frmDisMsg.lblMessage = "«›“«Ì‘ ‰«„ ò·ÌÂ ò«·«Â« »Â «‰»«— «‰Ã«„ ‘œ "
            frmDisMsg.Timer1.Enabled = True
            frmDisMsg.Show vbModal
        End If

End Function

Private Sub cmdInventoryGood_Add_Click()
    AddGoodstoInventory

End Sub

Private Sub cmdInventoryGood_Delete_Click()
    If cmbBranch.ListIndex = -1 Then Exit Sub
    If cmbInventory.ListIndex = -1 Then Exit Sub
    
    frmMsg.fwlblMsg.Caption = "¬Ì« »—«Ì Õ–› ‰«„  ò«·«Â«Ì  »œÊ‰ ê—œ‘ Ê »œÊ‰ „ÊÃÊœÌ «Ê·ÌÂ «“ «‰»«— «ÿ„Ì‰«‰ œ«—Ìœ"
    frmMsg.fwBtn(0).ButtonType = flwButtonOk
    frmMsg.fwBtn(1).ButtonType = flwButtonCancel
    frmMsg.fwBtn(0).Caption = "»·Ì"
    frmMsg.fwBtn(1).Caption = "ŒÌ—"
    frmMsg.Show vbModal
    If mvarMsgIdx = vbYes Then
        ReDim Parameter(2) As Parameter

        Parameter(0) = GenerateInputParameter("@IntInventoryNo", adInteger, 4, cmbInventory.ItemData(cmbInventory.ListIndex))
        Parameter(1) = GenerateInputParameter("@Branch", adInteger, 4, cmbBranch.ItemData(cmbBranch.ListIndex))
        Parameter(2) = GenerateInputParameter("@AccountYear", adSmallInt, 2, Val(cmbSalMali.Text))
        
        RunParametricStoredProcedure "Delete_tinventory_Good_All", Parameter
        DefaultSetting
        frmDisMsg.lblMessage = "Õ–› ‰«„  ò«·«Â«Ì  »œÊ‰ ê—œ‘ Ê »œÊ‰ „ÊÃÊœÌ «Ê·ÌÂ «“ «‰»«— «‰Ã«„ ‘œ "
        frmDisMsg.Timer1.Enabled = True
        frmDisMsg.Show vbModal
    End If

End Sub
Private Sub FillSalMali()
    cmbSalMali.Clear
    Set rs = RunStoredProcedure2RecordSet("Get_All_tAccountYears")
    Do While rs.EOF = False
        cmbSalMali.AddItem rs!AccountYear
        rs.MoveNext
    Loop
    Dim i As Integer
    For i = 0 To cmbSalMali.ListCount - 1
        cmbSalMali.ListIndex = i
        If AccountYear = cmbSalMali.Text Then
            Exit For
        End If
    Next
    'If cmbSalMali.ListCount > 0 Then cmbSalMali.ListIndex = 0
    rs.Close
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

Private Sub Form_Load()

    CenterCenter Me
    
    VarActForm = Me.Name
    
    FillBranch
    FillSalMali
    MyFormAddEditMode = ViewMode
    SetFirstToolBar
    DefaultSetting
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
Private Sub FillBranch()
    Dim rctmp As New ADODB.Recordset
    cmbBranch.Clear
    Set rctmp = RunStoredProcedure2RecordSet("Get_All_Branches")
    Do While rctmp.EOF = False
        cmbBranch.AddItem rctmp!nvcBranchName
        cmbBranch.ItemData(cmbBranch.NewIndex) = rctmp!Branch
        rctmp.MoveNext
    Loop
    rctmp.Close

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

Private Sub Form_Unload(Cancel As Integer)
    VarActForm = ""
    SaveSetting strMainKey, Me.Name, "Left", Me.Left
    SaveSetting strMainKey, Me.Name, "Top", Me.Top

    If PosConnection.State = adStateOpen Then PosConnection.Close: Set PosConnection = Nothing
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
Public Sub Delete()

End Sub

Public Sub Edit()
    MyFormAddEditMode = EditMode
    SetFirstToolBar
End Sub

Public Sub ChangeLanguage()
    
    DefaultSetting
    
End Sub

Private Sub DefaultSetting()

    If cmbBranch.ListIndex = -1 Then Exit Sub
    
    cmbInventory.Clear
    lstGoodLevel1.Clear
    txtGroupName.Text = ""
    txtGroupName.Locked = True
    
    Dim rctmp As New ADODB.Recordset
    
    If rctmp.State <> 0 Then rctmp.Close
''''    Set rctmp = RunStoredProcedure2RecordSet("Get_tInventoryType")
    Dim Rst As New ADODB.Recordset
    
    ReDim Parameter(1) As Parameter
    Parameter(0) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
    Parameter(1) = GenerateInputParameter("@Branch", adInteger, 4, cmbBranch.ItemData(cmbBranch.ListIndex))
    Set rctmp = RunParametricStoredProcedure2Rec("GetInventory_Branch", Parameter)
    
    If Not (rctmp.EOF = True And rctmp.BOF = True) Then
        Do While rctmp.EOF <> True
            cmbInventory.AddItem rctmp.Fields("Description")
            cmbInventory.ItemData(cmbInventory.ListCount - 1) = Val(rctmp.Fields("InventoryNo"))
            rctmp.MoveNext
        Loop
    End If
    rctmp.Close
    
    ReDim Parameter(0) As Parameter
    Parameter(0) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
    Set rctmp = RunParametricStoredProcedure2Rec("Get_tGoodLevel1", Parameter)
        
    If (rctmp.EOF = True And rctmp.BOF = True) Then
        Exit Sub
    End If
    
    While rctmp.EOF = False
        lstGoodLevel1.AddItem rctmp.Fields("Description")
        lstGoodLevel1.ItemData(lstGoodLevel1.ListCount - 1) = rctmp.Fields("Code")
        rctmp.MoveNext
    Wend
    
    
   
End Sub

Public Sub ExitForm()
    Unload Me

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
    
    If MyFormAddEditMode = ViewMode Then  ' View Mode
    
        mdifrm.Toolbar1.Buttons(6).Enabled = True  'Add
        mdifrm.Toolbar1.Buttons(7).Enabled = True  'Edit
        mdifrm.Toolbar1.Buttons(8).Enabled = False  'Enter
        mdifrm.Toolbar1.Buttons(9).Enabled = False   'Esc
        txtGroupName.Locked = True
        
    ElseIf MyFormAddEditMode = AddMode Then    'Add Mode
    
        mdifrm.Toolbar1.Buttons(6).Enabled = False  'Add
        mdifrm.Toolbar1.Buttons(7).Enabled = False  'Edit
        mdifrm.Toolbar1.Buttons(8).Enabled = True   'Enter
        mdifrm.Toolbar1.Buttons(9).Enabled = True   'Esc
        txtGroupName.Locked = False
        
    ElseIf MyFormAddEditMode = EditMode Then     'Edit
        
        mdifrm.Toolbar1.Buttons(6).Enabled = False  'Add
        mdifrm.Toolbar1.Buttons(7).Enabled = False  'Edit
        mdifrm.Toolbar1.Buttons(8).Enabled = True   'Enter
        mdifrm.Toolbar1.Buttons(9).Enabled = True   'Esc
        txtGroupName.Locked = False
    
    End If
    
    HeaderLabel Val(MyFormAddEditMode), fwlblMode


End Sub


Private Sub ResizeKit1_ExitResize(ByVal XScale As Double, ByVal YScale As Double)


    If formloadFlag = True Then
        SaveSetting strMainKey, Me.Name, "Height", Me.Height
        SaveSetting strMainKey, Me.Name, "Width", Me.Width
    End If

End Sub
