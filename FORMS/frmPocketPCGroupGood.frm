VERSION 5.00
Object = "{75D4F767-8785-11D3-93AD-0000832EF44D}#2.14#0"; "FAST2001.ocx"
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Begin VB.Form frmPocketPCGroupGood 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7335
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9285
   Icon            =   "frmPocketPCGroupGood.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   7335
   ScaleWidth      =   9285
   Begin VB.CommandButton CmdDone 
      BackColor       =   &H00008000&
      Caption         =   " «ÌÌœ"
      BeginProperty Font 
         Name            =   "B Homa"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   6600
      Width           =   1305
   End
   Begin VSFlex7LCtl.VSFlexGrid vsDefinedGoods 
      Height          =   4485
      Left            =   120
      TabIndex        =   6
      Top             =   1980
      Width           =   4605
      _cx             =   8123
      _cy             =   7911
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
      ForeColor       =   8388608
      BackColorFixed  =   16761024
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
      AllowUserResizing=   0
      SelectionMode   =   3
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   3
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmPocketPCGroupGood.frx":A4C2
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
   Begin VB.CommandButton cmddeSelectGoods 
      Caption         =   ">"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   765
      Left            =   4710
      TabIndex        =   5
      Top             =   3915
      Width           =   765
   End
   Begin VB.CommandButton cmdSelectGoods 
      Caption         =   "<"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   765
      Left            =   4710
      TabIndex        =   4
      Top             =   3120
      Width           =   765
   End
   Begin VB.ComboBox cboGroup 
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
      Left            =   5640
      RightToLeft     =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   1080
      Width           =   3525
   End
   Begin FLWCtrls.FWLabel3D fwlblMode 
      Height          =   525
      Left            =   7680
      Top             =   0
      Visible         =   0   'False
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
   Begin VSFlex7LCtl.VSFlexGrid vsNotDefinedGoods 
      Height          =   4485
      Left            =   5490
      TabIndex        =   7
      Top             =   1980
      Width           =   3645
      _cx             =   6429
      _cy             =   7911
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
      ForeColor       =   128
      BackColorFixed  =   16761024
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
      AllowUserResizing=   0
      SelectionMode   =   3
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   2
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmPocketPCGroupGood.frx":A538
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
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   " ⁄—Ì› ò«·«Â« œ— ê—ÊÂÂ«Ì  PocketPC"
      BeginProperty Font 
         Name            =   "Nazanin"
         Size            =   15.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   495
      Left            =   1560
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   0
      Width           =   5295
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H008080FF&
      BackStyle       =   0  'Transparent
      Caption         =   "ò«·«Â«Ì  ⁄—Ì› ‘œÂ œ— ê—ÊÂ"
      BeginProperty Font 
         Name            =   "B Homa"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   465
      Left            =   1620
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   1500
      Width           =   3045
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00E0E0E0&
      BackStyle       =   0  'Transparent
      Caption         =   "ò«·«Â«Ì  ⁄—Ì› ‰‘œÂ œ— ê—ÊÂ"
      BeginProperty Font 
         Name            =   "B Homa"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   465
      Left            =   6720
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   1500
      Width           =   2415
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "ê—ÊÂ ò«·«"
      BeginProperty Font 
         Name            =   "B Homa"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   465
      Left            =   6720
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   600
      Width           =   2415
   End
End
Attribute VB_Name = "frmPocketPCGroupGood"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MyFormAddEditMode As EnumAddEditMode
Dim i As Integer
Dim Parameter() As Parameter

Private Sub CmdDone_Click()

    Dim SelectedGoods As String
    Dim NameDisplays As String
    
    ReDim Parameter(2) As Parameter
    
    With vsDefinedGoods
        For i = 1 To .Rows - 1
            SelectedGoods = SelectedGoods & .TextMatrix(i, 0) & ","
            If Trim(.TextMatrix(i, 2)) <> "" Then
                NameDisplays = NameDisplays & .TextMatrix(i, 2) & ","
            Else
                MsgBox "‰«„ ‰„«Ì‘Ì ‰»«Ìœ Œ«·Ì »«‘œ"
                Exit Sub
            End If
        Next i
        
    End With
    
    If SelectedGoods <> "" Then
        SelectedGoods = Left(SelectedGoods, Len(SelectedGoods) - 1)
        NameDisplays = Left(NameDisplays, Len(NameDisplays) - 1)
    End If
    
    If cboGroup.ListIndex > -1 Then
        Parameter(0) = GenerateInputParameter("@PocketPCGroupCode", adInteger, 4, cboGroup.ItemData(cboGroup.ListIndex))
        Parameter(1) = GenerateInputParameter("@StrGoodCode", adVarWChar, 4000, SelectedGoods)
        Parameter(2) = GenerateInputParameter("@NameDisplay", adVarWChar, 4000, NameDisplays)
        RunParametricStoredProcedure "Insert_PocketPC_Good", Parameter
        cboGroup.ListIndex = cboGroup.ListIndex
        ShowDisMessage "«Œ ’«’ ò«·« »Â „‰Ê «‰Ã«„ ‘œ", 1500
    End If
End Sub



Private Sub cboGroup_Click()
    
    Dim Rst As New ADODB.Recordset
    Dim intTemp As Integer
    
    vsDefinedGoods.Rows = 1
    vsNotDefinedGoods.Rows = 1
    
    If cboGroup.ListIndex > -1 Then
        intTemp = cboGroup.ItemData(cboGroup.ListIndex)
    End If
    
    If Rst.State <> 0 Then Rst.Close
    
    ReDim Parameter(0) As Parameter
    Parameter(0) = GenerateInputParameter("@PocketPCGroupCode", adInteger, 4, intTemp)
    Set Rst = RunParametricStoredProcedure2Rec("Get_undefined_PocketPC_Good", Parameter)
    
    While Rst.EOF <> True
        
        With vsNotDefinedGoods
            .Rows = .Rows + 1
            i = .Rows - 1
            .TextMatrix(i, 0) = Rst.Fields("code").Value
            .TextMatrix(i, 1) = Rst.Fields("Name").Value
        
        End With
        Rst.MoveNext
    
    Wend
    
    If Rst.State <> 0 Then Rst.Close
    
    ReDim Parameter(0) As Parameter
    Parameter(0) = GenerateInputParameter("@PocketPCGroupCode", adInteger, 4, intTemp)
    Set Rst = RunParametricStoredProcedure2Rec("Get_defined_PocketPC_Good", Parameter)
    
    While Rst.EOF <> True
        
        With vsDefinedGoods
            .Rows = .Rows + 1
            i = .Rows - 1
            .TextMatrix(i, 0) = Rst.Fields("code").Value
            .TextMatrix(i, 1) = Rst.Fields("Name").Value
            If IsNull(Rst.Fields("NameDisplay").Value) = False Then
                .TextMatrix(i, 2) = Rst.Fields("NameDisplay").Value
            End If
        End With
        Rst.MoveNext
    
    Wend
    Set Rst = Nothing
End Sub

Private Sub cmddeSelectGoods_Click()

    If vsDefinedGoods.SelectedRows = 0 Then Exit Sub
    
    For i = 0 To vsDefinedGoods.SelectedRows - 1
        vsNotDefinedGoods.Rows = vsNotDefinedGoods.Rows + 1
        vsNotDefinedGoods.TextMatrix(vsNotDefinedGoods.Rows - 1, 0) = vsDefinedGoods.TextMatrix(vsDefinedGoods.SelectedRow(i), 0)
        vsNotDefinedGoods.TextMatrix(vsNotDefinedGoods.Rows - 1, 1) = vsDefinedGoods.TextMatrix(vsDefinedGoods.SelectedRow(i), 1)
    Next i
    For i = 0 To vsDefinedGoods.SelectedRows - 1
        vsDefinedGoods.RemoveItem vsDefinedGoods.SelectedRow(0)
    Next i
    
End Sub

Private Sub cmdSelectGoods_Click()
    
    If vsNotDefinedGoods.SelectedRows = 0 Then Exit Sub
    
    If vsNotDefinedGoods.SelectedRows + vsDefinedGoods.Rows - 1 > 12 Then
        ShowDisMessage "œﬁ  ﬂ‰Ìœ œ— Å«ﬂ  ÅÌ ”Ì  ‰ŸÌ„«  —« ÿÊ—Ì «‰Ã«„ œÂÌœ ﬂÂ «Ì‰  ⁄œ«œ ﬂ«·« —« ‰„«Ì‘ œÂœ", 1000
        'Exit Sub
    End If
    
    For i = 0 To vsNotDefinedGoods.SelectedRows - 1
        vsDefinedGoods.Rows = vsDefinedGoods.Rows + 1
        vsDefinedGoods.TextMatrix(vsDefinedGoods.Rows - 1, 0) = vsNotDefinedGoods.TextMatrix(vsNotDefinedGoods.SelectedRow(i), 0)
        vsDefinedGoods.TextMatrix(vsDefinedGoods.Rows - 1, 1) = vsNotDefinedGoods.TextMatrix(vsNotDefinedGoods.SelectedRow(i), 1)
        vsDefinedGoods.TextMatrix(vsDefinedGoods.Rows - 1, 2) = vsNotDefinedGoods.TextMatrix(vsNotDefinedGoods.SelectedRow(i), 1)
    Next i
    For i = 0 To vsNotDefinedGoods.SelectedRows - 1
        vsNotDefinedGoods.RemoveItem vsNotDefinedGoods.SelectedRow(0)
    Next i
    
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

    If intVersion <> gold And intVersion <> Diamond Then
        ShowDisMessage "«„ﬂ«‰  ⁄—Ì› ê—ÊÂÂ«Ì Å«ﬂ  ÅÌ ”Ì ›ﬁÿ œ— ‰”ŒÂ ÊÌéÂ ÊÃÊœ œ«—œ", 1500
        Unload Me
        Exit Sub
    End If
    CenterCenter Me
    
    VarActForm = Me.Name
    
    MyFormAddEditMode = ViewMode
    DefaultSetting
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

Private Sub Form_Unload(Cancel As Integer)
    If PosConnection.State = adStateOpen Then PosConnection.Close: Set PosConnection = Nothing
    VarActForm = ""
    
End Sub

Public Sub Add()

''    MyFormAddEditMode = AddMode
''    SetFirstToolbar
''
End Sub
Public Sub Cancel()

'    DefaultSetting
'    MyFormAddEditMode = ViewMode
'    SetFirstToolbar
'
End Sub
Public Sub Delete()

End Sub

Public Sub Edit()
'    MyFormAddEditMode = EditMode
'    SetFirstToolbar
End Sub

Public Sub Update()
''
''    Dim intResult As Integer
''
''    Select Case MyFormAddEditMode
''
''        Case AddMode
''
''            If txtGroupName.Text = "" Or InStr(txtGroupName.Text, "'") <> 0 Then
''                Exit Sub
''            End If
''
''            Dim Parameter(1) As Parameter
''
''            Parameter(0) = GenerateInputParameter("@Description", adVarWChar, 50, txtGroupName.Text)
''            Parameter(1) = GenerateOutputParameter("@Result", adInteger, 4)
''            intResult = RunParametricStoredProcedure("Insert_PocketPCGroup", Parameter, Cnn)
''            If intResult <> -1 Then
''
''            Else
''
''            End If
''
''        Case EditMode
''
''            If txtGroupName.Text = "" Or InStr(txtGroupName.Text, "'") <> 0 Then
''                Exit Sub
''            End If
''            Dim Parameter2(3) As Parameter
''            Parameter2(0) = GenerateInputParameter("@PocketPCGroup", adInteger, 4, lstGroups.ItemData(lstGroups.ListIndex))
''            Parameter2(1) = GenerateInputParameter("@Description", adVarWChar, 50, txtGroupName.Text)
''            Parameter2(2) = GenerateInputParameter("@intLanguage", adInteger, 4, ClsStation.Language)
''            Parameter2(3) = GenerateOutputParameter("@Result", adInteger, 4)
''            intResult = RunParametricStoredProcedure("Update_PocketPCGroup", Parameter2, Cnn)
''            If intResult <> -1 Then
''
''            Else
''
''            End If
''    End Select
''
''    DefaultSetting
''    MyFormAddEditMode = ViewMode
''    SetFirstToolbar
''    HeaderLabel CInt(MyFormAddEditMode), Me.fwlblMode
End Sub

Private Sub DefaultSetting()

    cboGroup.Clear
    vsDefinedGoods.Rows = 1
    vsNotDefinedGoods.Rows = 1
    
    Dim Rst As New ADODB.Recordset
   
    ReDim Parameter(0) As Parameter
    Parameter(0) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
    
    If Rst.State <> 0 Then Rst.Close
    Set Rst = RunParametricStoredProcedure2Rec("Get_PocketPCGroups", Parameter)
    
    While Rst.EOF <> True
        If Rst.Fields("PocketPCGroupCode").Value > 8 Then
            cboGroup.AddItem Rst.Fields("Description").Value
            cboGroup.ItemData(cboGroup.ListCount - 1) = Rst.Fields("PocketPCGroupCode").Value
        End If
        Rst.MoveNext
    Wend
    
    Dim intTemp As Integer
    If cboGroup.ListIndex > -1 Then
        intTemp = cboGroup.ItemData(cboGroup.ListIndex)
    End If
    
    If Rst.State <> 0 Then Rst.Close
    
    ReDim Parameter(0) As Parameter
    Parameter(0) = GenerateInputParameter("@PocketPCGroupCode", adInteger, 4, intTemp)
    Set Rst = RunParametricStoredProcedure2Rec("Get_undefined_PocketPC_Good", Parameter)
    
    While Rst.EOF <> True
        
        With vsNotDefinedGoods
            .Rows = .Rows + 1
            i = .Rows - 1
            .TextMatrix(i, 0) = Rst.Fields("code").Value
            .TextMatrix(i, 1) = Rst.Fields("Name").Value
        
        End With
        Rst.MoveNext
    
    Wend
    
    If Rst.State <> 0 Then Rst.Close
    
    ReDim Parameter(0) As Parameter
    Parameter(0) = GenerateInputParameter("@PocketPCGroupCode", adInteger, 4, intTemp)
    Set Rst = RunParametricStoredProcedure2Rec("Get_defined_PocketPC_Good", Parameter)
    
    While Rst.EOF <> True
        
        With vsDefinedGoods
            .Rows = .Rows + 1
            i = .Rows - 1
            .TextMatrix(i, 0) = Rst.Fields("code").Value
            .TextMatrix(i, 1) = Rst.Fields("Name").Value
            If IsNull(Rst.Fields("NameDisplay").Value) = False Then
                .TextMatrix(i, 2) = Rst.Fields("NameDisplay").Value
            End If
        End With
        Rst.MoveNext
    
    Wend
    
    If cboGroup.ListCount > 0 Then
        cboGroup.ListIndex = 0
    End If
    Set Rst = Nothing
    
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
    
        mdifrm.Toolbar1.Buttons(6).Enabled = False  'Add
        mdifrm.Toolbar1.Buttons(7).Enabled = False  'Edit
        mdifrm.Toolbar1.Buttons(8).Enabled = False  'Enter
        mdifrm.Toolbar1.Buttons(9).Enabled = False   'Esc
        
    ElseIf MyFormAddEditMode = AddMode Then    'Add Mode
    
        mdifrm.Toolbar1.Buttons(6).Enabled = False  'Add
        mdifrm.Toolbar1.Buttons(7).Enabled = False  'Edit
        mdifrm.Toolbar1.Buttons(8).Enabled = True   'Enter
        mdifrm.Toolbar1.Buttons(9).Enabled = True   'Esc
        
    ElseIf MyFormAddEditMode = EditMode Then     'Edit
        
        mdifrm.Toolbar1.Buttons(6).Enabled = False  'Add
        mdifrm.Toolbar1.Buttons(7).Enabled = False  'Edit
        mdifrm.Toolbar1.Buttons(8).Enabled = True   'Enter
        mdifrm.Toolbar1.Buttons(9).Enabled = True   'Esc
    
    End If
    
    HeaderLabel Val(MyFormAddEditMode), fwlblMode


End Sub


Private Sub vsDefinedGoods_Click()
    
    With vsDefinedGoods
        If .Row >= 1 And .Col = 2 Then
            .Select .Row, .Col
            .EditCell
        End If
    
    End With
End Sub

