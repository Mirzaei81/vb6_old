VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{7AEDC602-D94C-11D1-BB7A-00E0290EA3C9}#1.0#0"; "ResizeKit.ocx"
Object = "{75D4F767-8785-11D3-93AD-0000832EF44D}#2.14#0"; "FAST2001.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmPayment 
   ClientHeight    =   7830
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12900
   Icon            =   "frmPayment.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7830
   ScaleWidth      =   12900
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
      Left            =   10200
      RightToLeft     =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   17
      Top             =   720
      Width           =   1965
   End
   Begin VB.ComboBox cmbPerson 
      BackColor       =   &H00FFC0C0&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Nazanin"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   5520
      RightToLeft     =   -1  'True
      TabIndex        =   13
      TabStop         =   0   'False
      Text            =   "cmbPerson"
      Top             =   1560
      Width           =   3135
   End
   Begin FLWCtrls.FWLabel3D fwlblMode 
      Height          =   525
      Left            =   11280
      Top             =   0
      Width           =   1545
      _ExtentX        =   2725
      _ExtentY        =   926
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
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
   Begin FLWCtrls.FWLed FWLed1 
      Height          =   735
      Left            =   10680
      TabIndex        =   3
      Top             =   1320
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   1296
      ColorOff        =   8438015
      BackColor       =   8438015
   End
   Begin VSFlex7LCtl.VSFlexGrid vsPayment 
      Height          =   5475
      Left            =   120
      TabIndex        =   0
      Top             =   2160
      Width           =   12705
      _cx             =   22410
      _cy             =   9657
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
      BackColorFixed  =   -2147483645
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
      AllowUserResizing=   4
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   9
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmPayment.frx":A4C2
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
   Begin FLWCtrls.FWLed FWLed2 
      Height          =   615
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   1085
      ColorOff        =   0
   End
   Begin MSMask.MaskEdBox txtDate1 
      Height          =   585
      Left            =   120
      TabIndex        =   8
      Top             =   840
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   1032
      _Version        =   393216
      MaxLength       =   8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "##/##/##"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox txtDate2 
      Height          =   585
      Left            =   120
      TabIndex        =   9
      Top             =   1440
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   1032
      _Version        =   393216
      MaxLength       =   8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "##/##/##"
      PromptChar      =   " "
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   3600
      Top             =   1440
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowGroupTree=   -1  'True
      WindowAllowDrillDown=   -1  'True
      WindowShowCloseBtn=   -1  'True
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin RESIZEKITLibCtl.ResizeKit ResizeKit1 
      Height          =   480
      Left            =   3720
      OleObjectBlob   =   "frmPayment.frx":A5F2
      TabIndex        =   10
      Top             =   120
      Width           =   480
   End
   Begin MSMask.MaskEdBox txtDate 
      Height          =   585
      Left            =   7080
      TabIndex        =   11
      Top             =   840
      Width           =   1515
      _ExtentX        =   2672
      _ExtentY        =   1032
      _Version        =   393216
      MaxLength       =   8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "##/##/##"
      PromptChar      =   " "
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "‘⁄»Â"
      BeginProperty Font 
         Name            =   "B Homa"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   12120
      RightToLeft     =   -1  'True
      TabIndex        =   18
      Top             =   720
      Width           =   615
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Å—œ«Œ  ﬂ‰‰œÂ: "
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
      Left            =   8880
      RightToLeft     =   -1  'True
      TabIndex        =   16
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Label LblUserName 
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
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   3000
      TabIndex        =   15
      Top             =   960
      Width           =   3015
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ﬂ«—»— :  "
      BeginProperty Font 
         Name            =   "B Homa"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   6120
      RightToLeft     =   -1  'True
      TabIndex        =   14
      Top             =   960
      Width           =   705
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "  «—ÌŒ ”‰œ: "
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
      Left            =   9120
      RightToLeft     =   -1  'True
      TabIndex        =   12
      Top             =   960
      Width           =   975
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   " «  «—ÌŒ :"
      BeginProperty Font 
         Name            =   "B Homa"
         Size            =   11.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   1560
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   1440
      Width           =   945
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      Caption         =   "”«· „«·Ì"
      BeginProperty Font 
         Name            =   "B Homa"
         Size            =   11.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1560
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   0
      Width           =   1215
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   " Å—œ«Œ  «“ ’‰œÊﬁ"
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
      Left            =   5280
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   0
      Width           =   2655
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "«“  «—ÌŒ :"
      BeginProperty Font 
         Name            =   "B Homa"
         Size            =   11.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   1560
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   840
      Width           =   945
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "‘„«—Â"
      BeginProperty Font 
         Name            =   "B Homa"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   12120
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   1440
      Width           =   615
   End
End
Attribute VB_Name = "frmPayment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim clsDate As New clsDate
Dim i, ii As Integer
Dim Parameter() As Parameter
Dim MyFormAddEditMode As EnumAddEditMode
Dim Rst As New ADODB.Recordset


Public Sub ExitForm()

    Unload Me
    
End Sub

Private Sub Form_Activate()
    
    FWLed2.Value = CInt(AccountYear)
    FWLed1.BackColor = Me.BackColor
    FWLed1.ColorOff = Me.BackColor
    FWLed2.BackColor = Me.BackColor
    FWLed2.ColorOff = Me.BackColor
    
    SetFirstToolBar
    VarActForm = Me.Name
    
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call PresetScreenSaver
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call PresetScreenSaver
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case Shift
          Case 0
              Select Case KeyCode
                  Case 27  ' Esc
                      Me.ExitForm
                  Case Else
                    vsPayment_KeyDown KeyCode, Shift
              End Select
          Case 2
               Select Case KeyCode
                  Case 123  'Exit
                     If clsStation.KeyboardType = EnumKeyBoardType.Rb2 Then
                        Me.ExitForm
                     End If
                  Case Else
                    vsPayment_KeyDown KeyCode, Shift
              End Select

    End Select
End Sub

Private Sub Form_Load()

    If ClsFormAccess.frmPayment = False Then
        Unload Me
        Exit Sub
    End If

    If intVersion = Min Or intVersion = Normal Then
        ShowDisMessage "Å—œ«Œ  «“ ’‰œÊﬁ œ— ‰”ŒÂ Â«Ì ÅÌ‘—› Â Ê »«·« — ÊÃÊœ œ«—œ", 1500
        Unload Me
        Exit Sub
    End If
    
    CenterTop Me
    
    VarActForm = Me.Name
    
    txtDate1.Text = Right(clsDate.shamsi(Date), 8)
    txtDate2.Text = Right(clsDate.shamsi(Date), 8)
    Dim s As String
    
    With vsPayment
        .Rows = 1
       ' .ColAlignment(-1) = flexAlignRightCenter
        .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter
     '   .Cell(flexcpAlignment, 0, 5, 0, 5) = flexAlignRightCenter
'        .ColHidden(1) = True
        .ColHidden(2) = True
        .ColHidden(8) = True
        .ColHidden(3) = True
        .ColHidden(1) = True    'date
        
        .ColFormat(7) = "###,###"
        s = ""
        Set Rst = RunStoredProcedure2RecordSet("Get_User")
      '  s = .BuildComboList(Rst, "PersonName", "Uid")
      '  .ColComboList(3) = s
         
        cmbPerson.Clear
        If Not (Rst.EOF = True And Rst.BOF = True) Then
            Do While Rst.EOF <> True

                cmbPerson.AddItem CStr(Rst.Fields("PersonName"))
                cmbPerson.ItemData(cmbPerson.ListCount - 1) = Val(Rst.Fields("Uid"))
                Rst.MoveNext

            Loop
        End If
        
        s = ""
        If clsArya.ExternalAccounting = False Then
            Set Rst = RunStoredProcedure2RecordSet("Get_PaymentType")
        Else
            Set Rst = RunStoredProcedure2RecordSet("Get_PaymentType_Acc")
        End If
        s = .BuildComboList(Rst, "Description", "Code")
        .ColComboList(4) = s
        
        s = ""
        Set Rst = RunStoredProcedure2RecordSet("Get_ExpensiveType")
        s = .BuildComboList(Rst, "Description", "Code")
        .ColComboList(6) = s
         
         For i = 0 To .Cols - 1
            .ColWidth(i) = Val(GetSetting(strMainKey, "vsPayment", "Col" & i))
            If .ColWidth(i) = 0 Then
                .ColWidth(i) = 1000       '
            End If
         Next i
        
    End With
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

    FillBranch
    Add
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
    If Rst.State = adStateOpen Then Rst.Close: Set Rst = Nothing
    If PosConnection.State = adStateOpen Then PosConnection.Close: Set PosConnection = Nothing
    Unload frmFindCust
    

    SaveSetting strMainKey, Me.Name, "Left", Me.Left
    SaveSetting strMainKey, Me.Name, "Top", Me.Top


    
End Sub
Public Sub Add()
    On Error GoTo ErrHandler
    If MyFormAddEditMode <> EditMode Then
        MyFormAddEditMode = AddMode
        SetFirstToolBar
        With vsPayment
            If Trim(.TextMatrix(.Row, 5)) <> "" And Trim(.TextMatrix(.Row, 7)) <> "" Then
                .Rows = .Rows + 1
                .Row = .Rows - 1
                .TextMatrix(.Row, 0) = .Row
                .Cell(flexcpAlignment, 1, 5, .Row, 5) = flexAlignRightCenter
            End If
        End With
'        vsPayment.ColHidden(1) = True
        vsPayment.ColHidden(2) = True
        Number
    Else     'Edit Mode
        With vsPayment
            If Trim(.TextMatrix(.Row, 5)) <> "" And Trim(.TextMatrix(.Row, 7)) <> "" Then
                .Rows = .Rows + 1
                .Row = .Rows - 1
                .TextMatrix(.Row, 0) = .Row
            End If
        End With
    End If
    ReDim Parameter(0) As Parameter
    
    Parameter(0) = GenerateInputParameter("@UserId", adInteger, 4, mvarCurUserNo)
    Set Rst = RunParametricStoredProcedure2Rec("Get_UserName", Parameter)
    
    LblUserName.Caption = IIf(IsNull(Rst!AddUserName), "", Rst!AddUserName)
    txtDate.Text = Mid(clsDate.shamsi(Date), 3)
    
    For ii = 0 To cmbPerson.ListCount - 1
        If cmbPerson.ItemData(ii) = mvarCurUserNo Then
            cmbPerson.ListIndex = ii
            ii = 0
            Exit For
        End If
    Next ii
    Exit Sub
ErrHandler:
    LogSave Me.Name, err, "Add"
    MsgBox err.Description
End Sub
Public Sub Edit()
    MyFormAddEditMode = EditMode
    SetFirstToolBar
End Sub
Public Sub Cancel()
    MyFormAddEditMode = ViewMode
    vsPayment.Rows = 1
    Add
    SetFirstToolBar
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
    mdifrm.Toolbar1.Buttons(6).Enabled = True  'Add
    
    If MyFormAddEditMode = ViewMode Then  ' View Mode
    
        mdifrm.Toolbar1.Buttons(6).Enabled = False  'Add
        mdifrm.Toolbar1.Buttons(7).Enabled = True  'Edit
        mdifrm.Toolbar1.Buttons(8).Enabled = False  'Enter
        mdifrm.Toolbar1.Buttons(9).Enabled = True   'Esc
          
       vsPayment.ColHidden(2) = False
    ElseIf MyFormAddEditMode = AddMode Then    'Add Mode
    
    '    mdifrm.Toolbar1.Buttons(6).Enabled = False  'Add
        mdifrm.Toolbar1.Buttons(7).Enabled = False  'Edit
        mdifrm.Toolbar1.Buttons(8).Enabled = True   'Enter
        mdifrm.Toolbar1.Buttons(9).Enabled = True   'Esc

    ElseIf MyFormAddEditMode = EditMode Then     'Edit
        
        mdifrm.Toolbar1.Buttons(6).Enabled = True  'Add
        mdifrm.Toolbar1.Buttons(7).Enabled = False  'Edit
        mdifrm.Toolbar1.Buttons(8).Enabled = True   'Enter
        mdifrm.Toolbar1.Buttons(9).Enabled = True   'Esc
 '      vsPayment.ColHidden(1) = True
       vsPayment.ColHidden(2) = True
        
    End If
    
    HeaderLabel Val(MyFormAddEditMode), fwlblMode

End Sub



Private Sub ResizeKit1_ExitResize(ByVal XScale As Double, ByVal YScale As Double)

    If formloadFlag = True Then
        SaveSetting strMainKey, Me.Name, "Height", Me.Height
        SaveSetting strMainKey, Me.Name, "Width", Me.Width
    End If

End Sub

Private Sub vsPayment_AfterSort(ByVal Col As Long, Order As Integer)
    For i = 1 To vsPayment.Rows - 1
        vsPayment.TextMatrix(i, 0) = i
    Next
End Sub
Public Sub Number()
    
    ReDim Parameter(2) As Parameter
    Parameter(0) = GenerateInputParameter("@AccountYear", adSmallInt, 2, AccountYear)
    Parameter(1) = GenerateInputParameter("@Branch", adInteger, 4, CurrentBranch)
    Parameter(2) = GenerateOutputParameter("@No", adBigInt, 8)
    
    FWLed1.Tag = RunParametricStoredProcedure("Get_New_tblAcc_Cash", Parameter)
    FWLed1.Value = FWLed1.Tag Mod 1000
    
End Sub

Public Sub Update()
    
Dim Result As Long
Dim Obj As Object
Dim TotalPayment As Long
vsPayment_ValidateEdit vsPayment.Row, vsPayment.Col, False
With vsPayment
        For i = 1 To .Rows - 1
                If Not Trim(txtDate.ClipText) <> "" And Trim(.TextMatrix(i, 4)) <> "" And Trim(.TextMatrix(i, 5)) <> "" And Trim(.TextMatrix(i, 7)) <> "" Then
                frmMsg.fwlblMsg.Caption = "·ÿ›« «ÿ·«⁄«  ÷—Ê—Ì —« Ê«—œ ‰„«ÌÌœ"
                frmMsg.fwBtn(0).ButtonType = flwButtonOk
                frmMsg.fwBtn(0).Caption = "ﬁ»Ê·"
                frmMsg.fwBtn(1).Visible = False
                frmMsg.Show vbModal
                Exit Sub
            End If
        Next i
    
   
    Select Case MyFormAddEditMode
        Case AddMode
        
            '' get new max number
            Number
            
            ReDim Parameter(10) As Parameter
            For i = 1 To .Rows - 1
                If Trim(.TextMatrix(i, 4)) <> "" And Trim(.TextMatrix(i, 5)) <> "" And Trim(.TextMatrix(i, 7)) <> "" Then 'Trim(.TextMatrix(i, 3)) <> "" And
                      
                    Parameter(0) = GenerateInputParameter("@No", adBigInt, 8, FWLed1.Tag)
                    Parameter(1) = GenerateInputParameter("@List", adTinyInt, 1, .TextMatrix(i, 0))
                    Parameter(2) = GenerateInputParameter("@Date", adVarChar, 50, Trim(txtDate.Text))
                    Parameter(3) = GenerateInputParameter("@Uid", adInteger, 4, cmbPerson.ItemData(cmbPerson.ListIndex))
                    Parameter(4) = GenerateInputParameter("@Description", adVarChar, 300, .TextMatrix(i, 5))
                    Parameter(5) = GenerateInputParameter("@Bestankar", adBigInt, 8, .TextMatrix(i, 7))
                    Parameter(6) = GenerateInputParameter("@PaymentType", adInteger, 4, .TextMatrix(i, 4))
                    Parameter(7) = GenerateInputParameter("@Uid_Bede", adInteger, 4, IIf(.TextMatrix(i, 8) = "", 0, .TextMatrix(i, 8)))
                    Parameter(8) = GenerateInputParameter("@AddUser", adInteger, 4, mvarCurUserNo)
                    Parameter(9) = GenerateInputParameter("@AccountYear", adSmallInt, 2, AccountYear)
                    Parameter(10) = GenerateInputParameter("@Branch", adInteger, 4, CurrentBranch)
                    
                    Result = RunParametricStoredProcedure("Insert_tblAcc_Cash", Parameter)
                End If
            Next i
            
            
            frmMsg.fwlblMsg.Caption = "À»  «ÿ·«⁄«  ÃœÌœ »« „Ê›ﬁÌ  Å«Ì«‰ Ì«› "
            frmMsg.fwBtn(0).ButtonType = flwButtonOk
            frmMsg.fwBtn(0).Caption = "ﬁ»Ê·"
            frmMsg.fwBtn(1).Visible = False
            frmMsg.Show vbModal
            vsPayment.Rows = 1
            MyFormAddEditMode = AddMode
            Add
            
        Case EditMode
            ReDim Parameter(2) As Parameter
            Parameter(0) = GenerateInputParameter("@No", adBigInt, 8, FWLed1.Tag)
            Parameter(1) = GenerateInputParameter("@AccountYear", adSmallInt, 2, AccountYear)
            Parameter(2) = GenerateInputParameter("@Branch", adInteger, 4, CurrentBranch)
                
            Result = RunParametricStoredProcedure("Update_tblAcc_Cash", Parameter)
            
            ReDim Parameter(10) As Parameter
            For i = 1 To vsPayment.Rows - 1
                If Trim(.TextMatrix(i, 4)) <> "" And Trim(.TextMatrix(i, 5)) <> "" And Trim(.TextMatrix(i, 7)) <> "" Then 'Trim(.TextMatrix(i, 3)) <> "" And
                    Parameter(0) = GenerateInputParameter("@No", adBigInt, 8, FWLed1.Tag)
                    Parameter(1) = GenerateInputParameter("@List", adTinyInt, 1, .TextMatrix(i, 0))
                    Parameter(2) = GenerateInputParameter("@Date", adVarChar, 50, Trim(txtDate.Text))
                    Parameter(3) = GenerateInputParameter("@Uid", adInteger, 4, cmbPerson.ItemData(cmbPerson.ListIndex))
                    Parameter(4) = GenerateInputParameter("@Description", adVarChar, 300, .TextMatrix(i, 5))
                    Parameter(5) = GenerateInputParameter("@Bestankar", adBigInt, 8, .TextMatrix(i, 7))
                    Parameter(6) = GenerateInputParameter("@PaymentType", adInteger, 4, .TextMatrix(i, 4))
                    Parameter(7) = GenerateInputParameter("@Uid_Bede", adInteger, 4, IIf(.TextMatrix(i, 8) = "", 0, .TextMatrix(i, 8)))
                    Parameter(8) = GenerateInputParameter("@AddUser", adInteger, 4, mvarCurUserNo)
                    Parameter(9) = GenerateInputParameter("@AccountYear", adSmallInt, 2, AccountYear)
                    Parameter(10) = GenerateInputParameter("@Branch", adInteger, 4, CurrentBranch)
                    Result = RunParametricStoredProcedure("Insert_tblAcc_Cash", Parameter)
                End If
            Next i
            
            
            frmMsg.fwlblMsg.Caption = " €ÌÌ— «ÿ·«⁄«   »« „Ê›ﬁÌ  Å«Ì«‰ Ì«› "
            frmMsg.fwBtn(0).ButtonType = flwButtonOk
            frmMsg.fwBtn(0).Caption = "ﬁ»Ê·"
            frmMsg.fwBtn(1).Visible = False
            frmMsg.Show vbModal
            vsPayment.Rows = 1
            MyFormAddEditMode = AddMode
            Add
    End Select
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
        
    End Select
End Sub

Private Sub vsPayment_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    For i = 0 To vsPayment.Cols - 1
        SaveSetting strMainKey, "vsPayment", "Col" & i, vsPayment.ColWidth(i)
    Next
End Sub

Private Sub vsPayment_Click()
    With vsPayment
        If (MyFormAddEditMode = EnumAddEditMode.EditMode Or MyFormAddEditMode = EnumAddEditMode.AddMode) Then
             If .Col = 6 And .TextMatrix(.Row, 4) = "" Then Exit Sub
             If .Col <> 6 Then
                .Select .Row, .Col
                .EditCell
             ElseIf .Col = 6 And .TextMatrix(.Row, 4) = 0 Then
                .Select .Row, .Col
                .EditCell
             ElseIf .Col = 6 And (.TextMatrix(.Row, 4) = 1 Or .TextMatrix(.Row, 4) = 2 Or .TextMatrix(.Row, 4) = 3 Or .TextMatrix(.Row, 4) = 4) Then
                frmFindPerson.Show vbModal
            
                If mvarcode <> 0 Then
                  .TextMatrix(.Row, 6) = mvarName     ' lblCustomer.Tag = mvarcode
                  .TextMatrix(.Row, 8) = mvarcode
                  mvarcode = 0
                End If
                
             ElseIf .Col = 6 And .TextMatrix(.Row, 4) = 5 Then      'Suppliers
                frmFindSupplier.Show vbModal
            
                If mvarcode <> 0 Then
                  .TextMatrix(.Row, 6) = mvarName     ' lblCustomer.Tag = mvarcode
                  .TextMatrix(.Row, 8) = mvarcode
                  mvarcode = 0
                End If
             ElseIf .Col = 6 And .TextMatrix(.Row, 4) = 6 Then   'Customers
                frmFindCust.Show vbModal
            
                If mvarcode <> 0 Then
                  .TextMatrix(.Row, 6) = mvarName     ' lblCustomer.Tag = mvarcode
                  .TextMatrix(.Row, 8) = mvarcode
                  mvarcode = 0
                End If
                
             End If
        End If
    
    End With
End Sub

Private Sub vsPayment_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then
     With vsPayment
         If .TextMatrix(.Row, 4) <> "" And .TextMatrix(.Row, 5) <> "" And .TextMatrix(.Row, 6) <> "" And .TextMatrix(.Row, 7) <> "" Then '.TextMatrix(.Row, 3) <> "" And
           .Rows = .Rows + 1
           .Row = .Rows - 1
           .TextMatrix(.Row, 0) = .Row
         ElseIf .TextMatrix(.Row, 4) <> "" And .TextMatrix(.Row, 5) <> "" And .TextMatrix(.Row, 7) <> "" And .TextMatrix(.Row, 4) = "0" Then '.TextMatrix(.Row, 3) <> "" And
           .Rows = .Rows + 1
           .Row = .Rows - 1
           .TextMatrix(.Row, 0) = .Row
        End If
     End With
  End If
End Sub
Public Sub FirstKey()
    On Error GoTo ErrHandler
    ReDim Parameter(3) As Parameter
    
    Parameter(0) = GenerateInputParameter("@No", adBigInt, 8, FWLed1.Tag)
    Parameter(1) = GenerateInputParameter("@Direction", adInteger, 4, 0)
    Parameter(2) = GenerateInputParameter("@AccountYear", adSmallInt, 2, AccountYear)
    Parameter(3) = GenerateInputParameter("@Branch", adInteger, 4, CurrentBranch)
    Set Rst = RunParametricStoredProcedure2Rec("Get_tblAcc_Cash", Parameter)
    If Not (Rst.EOF = True And Rst.BOF = True) Then
        GetRecrdsetDetail
    End If
    Set Rst = Nothing
    
    MyFormAddEditMode = ViewMode
    SetFirstToolBar
    Exit Sub
ErrHandler:
    LogSave Me.Name, err, "FirstKey"
    MsgBox err.Description
End Sub

Public Sub PreviousKey()
    On Error GoTo ErrHandler
    If Val(FWLed1.Tag) <= 0 Then Exit Sub
    ReDim Parameter(3) As Parameter
    
    Parameter(0) = GenerateInputParameter("@No", adBigInt, 8, FWLed1.Tag)
    Parameter(1) = GenerateInputParameter("@Direction", adInteger, 4, 1)
    Parameter(2) = GenerateInputParameter("@AccountYear", adSmallInt, 2, AccountYear)
    Parameter(3) = GenerateInputParameter("@Branch", adInteger, 4, CurrentBranch)
    Set Rst = RunParametricStoredProcedure2Rec("Get_tblAcc_Cash", Parameter)
    If Not (Rst.EOF = True And Rst.BOF = True) Then
        GetRecrdsetDetail
    End If
    Set Rst = Nothing
    
    MyFormAddEditMode = ViewMode
    SetFirstToolBar
    Exit Sub
ErrHandler:
    LogSave Me.Name, err, "PreviousKey"
    MsgBox err.Description
End Sub

Public Sub NextKey()
    ReDim Parameter(3) As Parameter
    
    Parameter(0) = GenerateInputParameter("@No", adBigInt, 8, FWLed1.Tag)
    Parameter(1) = GenerateInputParameter("@Direction", adInteger, 4, 2)
    Parameter(2) = GenerateInputParameter("@AccountYear", adSmallInt, 2, AccountYear)
    Parameter(3) = GenerateInputParameter("@Branch", adInteger, 4, CurrentBranch)
    Set Rst = RunParametricStoredProcedure2Rec("Get_tblAcc_Cash", Parameter)
    If Not (Rst.EOF = True And Rst.BOF = True) Then
        GetRecrdsetDetail
    End If
    Set Rst = Nothing
    
    MyFormAddEditMode = ViewMode
    SetFirstToolBar
End Sub

Public Sub LastKey()
    On Error GoTo ErrHandler
    ReDim Parameter(3) As Parameter
    
    Parameter(0) = GenerateInputParameter("@No", adBigInt, 8, FWLed1.Tag)
    Parameter(1) = GenerateInputParameter("@Direction", adInteger, 4, 3)
    Parameter(2) = GenerateInputParameter("@AccountYear", adSmallInt, 2, AccountYear)
    Parameter(3) = GenerateInputParameter("@Branch", adInteger, 4, CurrentBranch)
    Set Rst = RunParametricStoredProcedure2Rec("Get_tblAcc_Cash", Parameter)
    If Not (Rst.EOF = True And Rst.BOF = True) Then
        GetRecrdsetDetail
    End If
    Set Rst = Nothing
    
    MyFormAddEditMode = ViewMode
    SetFirstToolBar
    Exit Sub
ErrHandler:
    LogSave Me.Name, err, "LastKey"
    MsgBox err.Description
End Sub

Private Sub GetRecrdsetDetail()
    Dim UserID As Integer
    If Rst.EOF = True And Rst.BOF = True Then Exit Sub
    vsPayment.Rows = 1
    Dim ii As Integer
    ii = 0
    If Rst.State = adStateOpen Then
       If Rst.EOF = False Then
           UserID = Rst!AddUser
           Do While Not (Rst.EOF)
               vsPayment.Rows = vsPayment.Rows + 1
               ii = ii + 1
               vsPayment.TextMatrix(ii, 0) = Rst!List '
             '  vsPayment.TextMatrix(ii, 1) = Rst!Date '
               txtDate.Text = Rst!Date '
               vsPayment.TextMatrix(ii, 2) = Rst!RegTime '
               vsPayment.TextMatrix(ii, 3) = Rst!Uid 'nvcFirstName & " " & Rst!nvcSurName
               vsPayment.TextMatrix(ii, 4) = Rst!PaymentType
               vsPayment.TextMatrix(ii, 5) = Rst!Description
               vsPayment.TextMatrix(ii, 6) = IIf(IsNull(Rst!Person_Name), "", Rst!Person_Name)
               vsPayment.TextMatrix(ii, 7) = Rst!Bestankar
               vsPayment.TextMatrix(ii, 8) = Rst!Uid_Bede
               
               FWLed1.Tag = Rst!No
               FWLed1.Value = Rst!No Mod 1000
    '           txtDate1.Text = Rst!Date
               Rst.MoveNext
       
           Loop
           For ii = 0 To cmbPerson.ListCount - 1
               If cmbPerson.ItemData(ii) = vsPayment.TextMatrix(1, 3) Then
                   cmbPerson.ListIndex = ii
                   ii = 0
                   Exit For
               End If
           Next ii
           ReDim Parameter(0) As Parameter
           
           Parameter(0) = GenerateInputParameter("@UserId", adInteger, 4, UserID)
           Set Rst = RunParametricStoredProcedure2Rec("Get_UserName", Parameter)
           
           LblUserName.Caption = IIf(IsNull(Rst!AddUserName), "", Rst!AddUserName)
       
       End If
       vsPayment.Cell(flexcpAlignment, 0, 0, vsPayment.Rows - 1, vsPayment.Cols - 1) = flexAlignCenterCenter
       vsPayment.Cell(flexcpAlignment, 1, 5, vsPayment.Rows - 1, 5) = flexAlignRightCenter
    End If
End Sub

Private Sub vsPayment_LeaveCell()
    With vsPayment
        If .Col = 6 And Val(.TextMatrix(.Row, 4)) = 0 Then
            .TextMatrix(.Row, 8) = .TextMatrix(.Row, 6)
        End If
    End With
End Sub

Public Sub Printing()
    On Error GoTo ErrHandler
    
    ReDim Parameter(5) As Parameter
    Parameter(0) = GenerateInputParameter("@SystemDate", adVarWChar, 50, Mid(clsDate.shamsi(Date), 3, 8))
    Parameter(1) = GenerateInputParameter("@SystemDay", adVarWChar, 50, clsDate.Find_DayOfWeek(Weekday(Date, vbSaturday)))
    Parameter(2) = GenerateInputParameter("@SystemTime", adVarWChar, 50, Mid(Str(time), 1, 5))
    Parameter(3) = GenerateInputParameter("@DateBefore", adVarChar, 50, txtDate1.Text)
    Parameter(4) = GenerateInputParameter("@DateAfter", adVarChar, 50, txtDate2.Text)
    Parameter(5) = GenerateInputParameter("@Branch", adInteger, 4, CurrentBranch)
    
    CrystalReport1.ReportFileName = App.Path & "\Reports" & RepVer & "\RepPayment_A4.rpt"
    Dim fileSystem As New FileSystemObject
    Dim IsFileExist As Boolean
    IsFileExist = fileSystem.FileExists(CrystalReport1.ReportFileName)
    If IsFileExist = False Then
            frmDisMsg.lblMessage = " ›«Ì·  " & CrystalReport1.ReportFileName & "ÅÌœ« ‰‘œ "
            frmDisMsg.Timer1.Interval = 3000
            frmDisMsg.Timer1.Enabled = True
            frmDisMsg.Show vbModal
            Exit Sub
    End If
    CrystalReport1.ReportTitle = "ê“«—‘  Å—œ«Œ Â«Ì «‰Ã«„ ‘œÂ "
    CrystalReport1.Destination = crptToWindow 'crptToPrinter '
    Dim intIndex As Integer
   
    For intIndex = 0 To UBound(Parameter) - LBound(Parameter)
        CrystalReport1.ParameterFields(intIndex) = CStr(Parameter(intIndex).Name) & ";" & CStr(Parameter(intIndex).Value) & ";" & "True"
    Next intIndex
  
    CrystalReport1.RetrieveDataFiles
    ODBCSetting clsArya.ServerName, clsArya.DbName
    CrystalReport1.Connect = CrystallConnection
    CrystalReport1.Action = 1
    If Screen.Width > 12000 Then
        CrystalReport1.PageZoom (100)
    Else
        CrystalReport1.PageZoom (75)
    End If
Exit Sub

ErrHandler:
    MsgBox err.Description
    LogSave Me.Name, err, "Printing"
    Resume Next
End Sub

Private Sub vsPayment_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsPayment
        .Row = Row
        .Col = Col
    End With

End Sub
