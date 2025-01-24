VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{7AEDC602-D94C-11D1-BB7A-00E0290EA3C9}#1.0#0"; "ResizeKit.ocx"
Object = "{75D4F767-8785-11D3-93AD-0000832EF44D}#2.14#0"; "FAST2001.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmReportGenerator 
   Caption         =   "                                                                                                      Ê·Ìœ ê“«—‘«  "
   ClientHeight    =   9045
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   14460
   Icon            =   "frmReportGenerator.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9045
   ScaleWidth      =   14460
   Begin VB.CommandButton cmd_Esc 
      BackColor       =   &H0000C0C0&
      Caption         =   "Œ—ÊÃ"
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
      TabIndex        =   13
      Top             =   8040
      Width           =   1455
   End
   Begin VB.TextBox txtReportName 
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
      Height          =   525
      Left            =   2880
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   2760
      Width           =   3705
   End
   Begin VB.TextBox txtItemReports 
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
      Height          =   525
      Left            =   2880
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   1920
      Width           =   3705
   End
   Begin VB.TextBox txtGroupReports 
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
      Height          =   525
      Left            =   2880
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   1080
      Width           =   3705
   End
   Begin VB.CommandButton cmd_Ok 
      BackColor       =   &H0000C000&
      Caption         =   "ﬁ»Ê·"
      BeginProperty Font 
         Name            =   "Nazanin"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   11880
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   8160
      Width           =   1365
   End
   Begin VB.ListBox lstItemReports 
      BeginProperty Font 
         Name            =   "Nazanin"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2400
      ItemData        =   "frmReportGenerator.frx":A4C2
      Left            =   8760
      List            =   "frmReportGenerator.frx":A4C4
      RightToLeft     =   -1  'True
      Style           =   1  'Checkbox
      TabIndex        =   1
      Top             =   960
      Width           =   2625
   End
   Begin VB.ListBox lstGroupReports 
      BeginProperty Font 
         Name            =   "Nazanin"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2400
      ItemData        =   "frmReportGenerator.frx":A4C6
      Left            =   11640
      List            =   "frmReportGenerator.frx":A4C8
      RightToLeft     =   -1  'True
      Style           =   1  'Checkbox
      TabIndex        =   0
      Top             =   960
      Width           =   2655
   End
   Begin RESIZEKITLibCtl.ResizeKit ResizeKit1 
      Height          =   480
      Left            =   0
      OleObjectBlob   =   "frmReportGenerator.frx":A4CA
      TabIndex        =   4
      Top             =   0
      Width           =   480
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   600
      Top             =   120
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
   Begin FLWCtrls.FWLabel3D fwlblMode 
      Height          =   495
      Left            =   12600
      Top             =   50
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
   Begin VSFlex7LCtl.VSFlexGrid vsItemReportsDetails 
      Height          =   4035
      Left            =   120
      TabIndex        =   12
      Top             =   3840
      Width           =   14115
      _cx             =   24897
      _cy             =   7117
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9
         Charset         =   178
         Weight          =   400
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
      Rows            =   1
      Cols            =   15
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmReportGenerator.frx":A550
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
      Editable        =   2
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
   Begin VB.Label LblIo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Nazanin"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   495
      Index           =   18
      Left            =   1080
      RightToLeft     =   -1  'True
      TabIndex        =   14
      Top             =   120
      Width           =   5055
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "‰«„ ›«Ì· ê“«—‘"
      BeginProperty Font 
         Name            =   "Nazanin"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   6960
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   2880
      Width           =   1455
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "‰«„ ê“«—‘"
      BeginProperty Font 
         Name            =   "Nazanin"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   6960
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   1920
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "‰«„ ê—ÊÂ ê“«—‘"
      BeginProperty Font 
         Name            =   "Nazanin"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   6840
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   1200
      Width           =   1575
   End
   Begin VB.Label lblGoodLevel2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H000080FF&
      BackStyle       =   0  'Transparent
      Caption         =   "‰«„ ê“«—‘« "
      BeginProperty Font 
         Name            =   "Titr"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   435
      Left            =   9120
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   480
      Width           =   2025
   End
   Begin VB.Label lblGoodLevel1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H000040C0&
      BackStyle       =   0  'Transparent
      Caption         =   "ê—ÊÂ ê“«—‘« "
      BeginProperty Font 
         Name            =   "Titr"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   435
      Left            =   11640
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   600
      Width           =   2415
   End
End
Attribute VB_Name = "frmReportGenerator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private MyFormAddEditMode As EnumAddEditMode
 
Dim Rst As New ADODB.Recordset
Dim Parameter() As Parameter
Dim ReportFileName As String
Dim CurrentCol As Integer
Dim GroupReportId As Integer
Dim ReportId As Integer
Dim flgerr As Integer

Public Sub ExitForm()
    Unload Me
End Sub

Private Sub cmd_Esc_Click()
    ExitForm
End Sub

Private Sub Form_Activate()
    VarActForm = Me.Name
     With vsItemReportsDetails
        s = ""
        If Rst.State <> 0 Then
            Rst.Close
        End If
        Set Rst = RunStoredProcedure2RecordSet("Get_All_ParameterType")
        s = .BuildComboList(Rst, "ParameterTypeName", "ParameterTypeId")
        .ColComboList(4) = s
        
        s = ""
        Set Rst = RunStoredProcedure2RecordSet("Get_All_ObjectType")
        s = .BuildComboList(Rst, "ObjectName", "ObjectType")
        .ColComboList(6) = s
    End With
    Set Rst = Nothing
    CurrentCol = -1
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
                  Case 13  ' Enter
                    SendKeys "{Left}", True
                  Case 27  ' Esc
                    Me.ExitForm
              End Select
          Case 2
               Select Case KeyCode
                  Case 123  'Exit
              End Select

    End Select

End Sub

Private Sub Form_Load()
'    If clsFormAccess.frmReports = False Then
'        Unload Me
'        Exit Sub
'    End If
     
     
        
    CenterTop Me
    VarActForm = Me.Name
    
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
    MyFormAddEditMode = ViewMode
    HeaderLabel Val(MyFormAddEditMode), fwlblMode
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    VarActForm = ""
    If PosConnection.State = adStateOpen Then PosConnection.Close: Set PosConnection = Nothing

    Dim i As Integer
    
    AllButton vbOff, True
    
    SaveSetting strMainKey, Me.Name, "Left", Me.Left
    SaveSetting strMainKey, Me.Name, "Top", Me.Top
    
    frmGroupReport.Show
End Sub

Public Sub DefaultSetting()
    txtGroupReports.Text = ""
    txtItemReports.Text = ""
    txtReportName.Text = ""
    txtGroupReports.Enabled = False
    txtItemReports.Enabled = False
    txtReportName.Enabled = False
    lstGroupReports.Clear
    lstItemReports.Clear
    FilllstGroupReports
    MyFormAddEditMode = ViewMode
    HeaderLabel Val(MyFormAddEditMode), fwlblMode
    mdifrm.Toolbar1.Buttons(6).Enabled = True
    mdifrm.Toolbar1.Buttons(7).Enabled = False
    mdifrm.Toolbar1.Buttons(8).Enabled = True
    mdifrm.Toolbar1.Buttons(9).Enabled = True
    mdifrm.Toolbar1.Buttons(10).Enabled = False
    vsItemReportsDetails.Rows = 1
    vsItemReportsDetails.Clear 1
    lstItemReports.Enabled = False
    vsItemReportsDetails.Editable = flexEDNone
    
    ReportId = -1
    GroupReportId = -1
End Sub

Public Sub FilllstGroupReports() ' it fills the lstGroupReports using table tgoodlevel1
    
    lstGroupReports.Clear
    lstItemReports.Clear
    
    Set Rst = RunStoredProcedure2RecordSet("Get_All_GroupReports")
        
    If (Rst.EOF = True And Rst.BOF = True) Then
        Exit Sub
    End If
    
    While Rst.EOF = False
        lstGroupReports.AddItem Rst.Fields("GroupReportName")
        lstGroupReports.ItemData(lstGroupReports.ListCount - 1) = Rst.Fields("intGroupreportId")
        Rst.MoveNext
    Wend
    lstGroupReports.ListIndex = 0
    FilllstItemReports
    Set Rst = Nothing
End Sub

Public Sub FilllstItemReports() ' it fills the lstItemReports using table tgoodlevel2

    lstItemReports.Clear
    ReportFileName = ""
    LblIo(18).Caption = ""
    
    If lstGroupReports.ListIndex = -1 Then
        Set Rst = Nothing
        Exit Sub
    Else
        ReDim Parameter(1) As Parameter
        Parameter(0) = GenerateInputParameter("@intGroupreportId", adInteger, 4, GroupReportId)
        Parameter(1) = GenerateInputParameter("@AccessLevel", adInteger, 4, mVarAccessLevel)
        
        If Rst.State <> 0 Then Rst.Close
        Set Rst = RunParametricStoredProcedure2Rec("Get_ItemReports_ByGroupId", Parameter)
        If (Rst.EOF = True And Rst.BOF = True) Then
            Set Rst = Nothing
            Exit Sub
        End If
       ' rst.moveFirst
        While Rst.EOF = False
            lstItemReports.AddItem Rst.Fields("ReportName")
            lstItemReports.ItemData(lstItemReports.ListCount - 1) = Rst.Fields("intReportId")
            Rst.MoveNext
        Wend
        
        Set Rst = Nothing
  '      lstItemReports.ListIndex = 0
        
    End If
    
End Sub

Private Sub lstGroupReports_Click()
 For i = 0 To lstGroupReports.ListCount - 1
    If i <> (lstGroupReports.ListIndex) Then
        lstGroupReports.Selected(i) = False
    End If
 Next i
    If lstGroupReports.Selected(lstGroupReports.ListIndex) = True Then
        txtGroupReports.Text = lstGroupReports.Text
        txtGroupReports.Enabled = False
''        txtItemReports.Enabled = True
''        txtReportName.Enabled = True
        lstItemReports.Enabled = True
        mdifrm.Toolbar1.Buttons(7).Enabled = True
        mdifrm.Toolbar1.Buttons(10).Enabled = True
        GroupReportId = lstGroupReports.ItemData(lstGroupReports.ListIndex)
    Else
        txtGroupReports.Text = ""
        txtGroupReports.Enabled = True
'''        txtItemReports.Enabled = False
'''        txtReportName.Enabled = False
        lstItemReports.Enabled = False
        mdifrm.Toolbar1.Buttons(7).Enabled = False
        mdifrm.Toolbar1.Buttons(10).Enabled = False
        GroupReportId = -1
    End If
    FilllstItemReports
    vsItemReportsDetails.Rows = 1
    vsItemReportsDetails.Clear 1
    txtItemReports.Text = ""
    txtReportName.Text = ""
End Sub

Private Sub lstItemReports_Click()
    For i = 0 To lstItemReports.ListCount - 1
    If i <> (lstItemReports.ListIndex) Then
        lstItemReports.Selected(i) = False
    End If
   Next i
   
    If lstItemReports.Selected(lstItemReports.ListIndex) = True Then
        txtItemReports.Text = lstItemReports.Text
        txtItemReports.Enabled = False
        txtReportName.Enabled = False
        ReportId = lstItemReports.ItemData(lstItemReports.ListIndex)
        GetDataDetaile
    Else
        txtItemReports.Text = ""
        txtReportName.Text = ""
        txtItemReports.Enabled = True
        txtReportName.Enabled = True
        vsItemReportsDetails.Rows = 1
        vsItemReportsDetails.Clear 1
        ReportId = -1
    End If
End Sub

Private Sub ClearParameters()
    On Error GoTo ErrHandler

    For i = 0 To 17
        LblIo(i).Visible = False
        LblIo(i).Caption = ""
        LblIo(i).Tag = ""
    Next i
Exit Sub
ErrHandler:
    MsgBox err.Description
    Resume Next
End Sub

Private Sub FillParameters()

End Sub


Private Sub ResizeKit1_ExitResize(ByVal XScale As Double, ByVal YScale As Double)
    If formloadFlag = True Then
        SaveSetting strMainKey, Me.Name, "Height", Me.Height
        SaveSetting strMainKey, Me.Name, "Width", Me.Width
    End If
End Sub

Public Sub ReportShow()
    On Error GoTo ErrorHandler
    '-----------------------
    CrystalReport1.ReportTitle = LblIo(18).Caption  ' ReportHeader
    CrystalReport1.Destination = crptToWindow 'crptToPrinter '
    Dim intIndex As Integer
   
    For intIndex = 0 To UBound(Parameter) - LBound(Parameter)
        CrystalReport1.ParameterFields(intIndex) = CStr(Parameter(intIndex).Name) & ";" & CStr(Parameter(intIndex).Value) & ";" & "True"
    Next intIndex
   
    CrystalReport1.RetrieveDataFiles
    ODBCSetting clsArya.ServerName, clsArya.DbName
    CrystalReport1.Connect = CrystallConnection
    CrystalReport1.Action = 1
    If PaperType = 1 Then
       CrystalReport1.PageZoom (100)
    Else
       CrystalReport1.PageZoom (100)
       
    End If
    Exit Sub
    
ErrorHandler:
    MsgBox err.Description & "  File Name:  " & CrystalReport1.ReportFileName
    Resume Next
End Sub

Public Sub Update()
    
    Dim DetailsString As String
    
    Dim Result As Integer
    Dim Obj As Object
    Dim CentralBranch As Boolean
    Dim CentralBranchCode As Integer
    
    If Trim(txtGroupReports.Text) = "" And txtGroupReports.Enabled = True Then
        ShowMessage "‰«„ ê—ÊÂ ê“«—‘«  Ê«—œ ‰‘œÂ «” ", True, False, " «ÌÌœ", ""
        txtGroupReports.SetFocus
        Exit Sub
    ElseIf Trim(txtItemReports.Text) = "" And (Trim(txtReportName.Text) <> "" Or ReportId <> -1) Then
        ShowMessage "‰«„ ê“«—‘ Ê«—œ ‰‘œÂ «” ", True, False, " «ÌÌœ", ""
        txtItemReports.SetFocus
        Exit Sub
    ElseIf Trim(txtReportName.Text) = "" And (Trim(txtItemReports.Text) <> "" Or ReportId <> -1) Then
        ShowMessage "‰«„ ›«Ì· ê“«—‘ Ê«—œ ‰‘œÂ «” ", True, False, " «ÌÌœ", ""
        txtItemReports.SetFocus
        Exit Sub
    End If
    
'''
'''    Select Case MyFormAddEditMode
'''
'''        Case AddMode
            
            
                ReDim Parameter(2) As Parameter
                Parameter(0) = GenerateInputParameter("@GroupReportId", adInteger, 4, GroupReportId)
                Parameter(1) = GenerateInputParameter("@GroupReportName", adVarWChar, 50, Trim(txtGroupReports.Text))
                Parameter(2) = GenerateOutputParameter("@intGroupReportId", adInteger, 4)
                
                Result = RunParametricStoredProcedure("Insert_tbltotal_GroupReports", Parameter)
                
                If Parameter(2).Value <> -1 Then
                    ShowMessage "À»  ê—ÊÂ ê“«—‘ »« „Ê›ﬁÌ  Å«Ì«‰ Ì«› ", True, False, " «ÌÌœ", ""
                Else
                    ShowMessage "ê—ÊÂ ê“«—‘ ÃœÌœ À»  ‰‘œ. ·ÿ›« œÊ»«—Â ”⁄Ì ‰„«ÌÌœ", True, False, " «ÌÌœ", ""
                End If
           If Trim(txtItemReports.Text) <> "" Then
                ReDim Parameter(4) As Parameter
                Parameter(0) = GenerateInputParameter("@ReportId", adInteger, 4, ReportId)
                Parameter(1) = GenerateInputParameter("@ReportName", adVarWChar, 100, Trim(txtItemReports.Text))
                Parameter(2) = GenerateInputParameter("@LatinReportName", adVarWChar, 50, Trim(txtReportName.Text))
                Parameter(3) = GenerateInputParameter("@intGroupReportId", adInteger, 4, lstGroupReports.ItemData(lstGroupReports.ListIndex))
                Parameter(4) = GenerateOutputParameter("@intReportId", adInteger, 4)
                
                Result = RunParametricStoredProcedure("Insert_tbltotal_ItemReports", Parameter)
                
                If Parameter(4).Value <> -1 Then
                    ShowMessage "À»  ‰«„ ê“«—‘ »« „Ê›ﬁÌ  Å«Ì«‰ Ì«› ", True, False, " «ÌÌœ", ""
                Else
                    ShowMessage "‰«„ ê“«—‘ ÃœÌœ À»  ‰‘œ. ·ÿ›« œÊ»«—Â ”⁄Ì ‰„«ÌÌœ", True, False, " «ÌÌœ", ""
                End If
           End If
            
    If ReportId <> -1 Then
        st = ""
        Dim j As Integer
        With vsItemReportsDetails
            For i = 1 To .Rows - 1
                If Val(.TextMatrix(i, 0)) <> 0 Then
                    If Not (Val(.TextMatrix(i, 0)) <> 0 And Trim(.TextMatrix(i, 1)) <> "" _
                    And Trim(.TextMatrix(i, 3)) <> "" And Trim(.TextMatrix(i, 4)) <> "" And Trim(.TextMatrix(i, 5)) <> "" _
                    And Trim(.TextMatrix(i, 6)) <> "" And Trim(.TextMatrix(i, 7)) <> "") Then 'And Trim(.TextMatrix(i, 8)) <> "") Then
                        ShowMessage "«ÿ·«⁄«  —« ﬂ«„· Ê«—œ ﬂ‰Ìœ", True, False, " «ÌÌœ", ""
                        Exit Sub
                     Else
                        j = i
                     End If
                     
                 End If
             Next i
        End With
        
        DetailsString = ""
        
        With vsItemReportsDetails
            For i = 1 To j
                DetailsString = GenerateDetailsStringReportGenarator(DetailsString, Trim(.TextMatrix(i, 0)), Trim(.TextMatrix(i, 1)), Trim(.TextMatrix(i, 2)), Trim(.TextMatrix(i, 3)), Trim(.TextMatrix(i, 4)), Trim(.TextMatrix(i, 5)), Trim(.TextMatrix(i, 6)), Trim(.TextMatrix(i, 7)), Trim(.TextMatrix(i, 8)), Trim(.TextMatrix(i, 9)), Trim(.TextMatrix(i, 10)), Trim(.TextMatrix(i, 11)), Trim(.TextMatrix(i, 12)), IIf(Trim(.TextMatrix(i, 13)) = "-1", "1", "0"))
            Next i
        End With
        
        ReDim Parameter(1) As Parameter
        Parameter(0) = GenerateInputParameter("@ReportId", adInteger, 4, ReportId)
        Parameter(1) = GenerateInputParameter("@ds", adVarWChar, 4000, DetailsString)
        Result = RunParametricStoredProcedure("INSERT_tblTotal_ItemReports_Details", Parameter)
    End If
        DefaultSetting
End Sub

Public Sub SetFirstToolBar()
    HeaderLabel Val(MyFormAddEditMode), fwlblMode
End Sub

Public Sub Add()
    
    
    DefaultSetting
    txtGroupReports.Enabled = True
    mdifrm.Toolbar1.Buttons(6).Enabled = False
    MyFormAddEditMode = AddMode
    SetFirstToolBar
End Sub

Public Sub Cancel()

    MyFormAddEditMode = ViewMode
    DefaultSetting
    
End Sub

Private Sub GetDataDetaile()
 
         ReDim Parameter(1) As Parameter
        Parameter(0) = GenerateInputParameter("@intReportId", adInteger, 4, ReportId)
        Parameter(1) = GenerateInputParameter("@Status", adBoolean, 1, 1)
        
        If Rst.State <> 0 Then Rst.Close
        Set Rst = RunParametricStoredProcedure2Rec("Get_ItemReports_Details_ByReportId_Rg", Parameter)
        If (Rst.EOF = True And Rst.BOF = True) Then
            Set Rst = Nothing
            Exit Sub
        End If

       txtReportName.Text = Rst.Fields("LatinReportName")
        
        Set Rst = Nothing
 
        ReDim Parameter(1) As Parameter
        Parameter(0) = GenerateInputParameter("@intReportId", adInteger, 4, lstItemReports.ItemData(lstItemReports.ListIndex))
        Parameter(1) = GenerateInputParameter("@Status", adBoolean, 1, 0)
        
        If Rst.State <> 0 Then Rst.Close
        Set Rst = RunParametricStoredProcedure2Rec("Get_ItemReports_Details_ByReportId_Rg", Parameter)
        If (Rst.EOF = True And Rst.BOF = True) Then
            Set Rst = Nothing
            Exit Sub
        End If

    With vsItemReportsDetails
        .Rows = 1
        If Not (Rst.EOF = True And Rst.BOF = True) Then
            While Rst.EOF <> True
                .Rows = .Rows + 1

                .TextMatrix(.Rows - 1, 0) = .Rows - 1
                .TextMatrix(.Rows - 1, 1) = IIf(IsNull(Rst!FromText), "", Rst!FromText)
                .TextMatrix(.Rows - 1, 2) = IIf(IsNull(Rst!ToText), "", Rst!ToText)
                .TextMatrix(.Rows - 1, 3) = IIf(IsNull(Rst!ParameterName), "", Rst!ParameterName)
                .TextMatrix(.Rows - 1, 4) = IIf(IsNull(Rst!ParameterType), "", Rst!ParameterType)
                .TextMatrix(.Rows - 1, 5) = IIf(IsNull(Rst!parameterLengh), "", Rst!parameterLengh)
                .TextMatrix(.Rows - 1, 6) = IIf(IsNull(Rst!ObjectType), "", Rst!ObjectType)
                .TextMatrix(.Rows - 1, 7) = IIf(IsNull(Rst!Quantity), "", Rst!Quantity)
                .TextMatrix(.Rows - 1, 8) = IIf(IsNull(Rst!MinValue), "", Rst!MinValue)
                .TextMatrix(.Rows - 1, 9) = IIf(IsNull(Rst!MaxValue), "", Rst!MaxValue)
                .TextMatrix(.Rows - 1, 10) = IIf(IsNull(Rst!ComboQuery), "", Rst!ComboQuery)
                .TextMatrix(.Rows - 1, 11) = IIf(IsNull(Rst!ComboFieldCode), "", Rst!ComboFieldCode)
                .TextMatrix(.Rows - 1, 12) = IIf(IsNull(Rst!ComboFieldDescr), "", Rst!ComboFieldDescr)
                .TextMatrix(.Rows - 1, 13) = IIf(IsNull(Rst!RightToLeft), 0, Rst!RightToLeft)
                Rst.MoveNext
            Wend
        End If

    End With

    If Rst.State = 1 Then Rst.Close

End Sub


Private Sub vsItemReportsDetails_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    With vsItemReportsDetails
        If .Col < 13 Then
           .Col = .Col + 1
           CurrentCol = -1
        ElseIf .Row = .Rows - 1 Then
        .TextMatrix(.Row, 0) = .Row
        .Rows = .Rows + 1
        .Row = .Row + 1
        .Col = 1
        End If
    End With
End Sub
Private Sub vsItemReportsDetails_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If CurrentCol <> OldCol Then
        With vsItemReportsDetails
            If (OldCol = 1 Or OldCol = 4 Or OldCol = 5 Or OldCol = 6 Or OldCol = 7 Or OldCol = 8) And .TextMatrix(.Row, OldCol) = "" Then
                CurrentCol = NewCol
                .Col = OldCol
            End If
            
        End With
   
    End If
End Sub
Public Sub Edit()
    mdifrm.Toolbar1.Buttons(7).Enabled = False
    MyFormAddEditMode = EditMode
    HeaderLabel Val(MyFormAddEditMode), fwlblMode
    txtGroupReports.Enabled = True
'    If ReportId <> -1 Then
        txtItemReports.Enabled = True
        txtReportName.Enabled = True
''    End If
    lstItemReports.Enabled = True
    vsItemReportsDetails.Editable = flexEDKbdMouse
    If ReportId <> -1 Then
        vsItemReportsDetails.Rows = vsItemReportsDetails.Rows + 1
    End If
    
End Sub
Public Sub Delete()
If GroupReportId <> -1 And ReportId = -1 Then
        frmMsg.fwlblMsg.Caption = "¬Ì« „«Ì·Ìœ ê—ÊÂ ê“«—‘«  «‰ Œ«» ‘œÂ Õ–› ‘Êœ"
        frmMsg.fwBtn(0).ButtonType = flwButtonOk
        frmMsg.fwBtn(0).Caption = "»·Â"
        frmMsg.fwBtn(1).ButtonType = flwButtonCancel
        frmMsg.fwBtn(1).Caption = "ŒÌ—"
        frmMsg.Show vbModal
        If mvarMsgIdx = vbNo Then
            Exit Sub
        Else
                On Error GoTo ErrHandler
                flgerr = 1
                ReDim Parameter(0) As Parameter
                Parameter(0) = GenerateInputParameter("@intGroupreportId", adInteger, 4, GroupReportId)
                Result = RunParametricStoredProcedure("Delete_tbltotal_GroupReports", Parameter)
                
                frmMsg.fwlblMsg.Caption = "ê—ÊÂ ê“«—‘«  »« „Ê›ﬁÌ  Õ–› ‘œ"
                frmMsg.fwBtn(0).Visible = False
                frmMsg.fwBtn(1).ButtonType = flwButtonCancel
                frmMsg.fwBtn(1).Caption = "ﬁ»Ê·"
                frmMsg.Show vbModal
        End If
ElseIf GroupReportId <> -1 And ReportId <> -1 Then
    If vsItemReportsDetails.Rows = 1 Then
        frmMsg.fwlblMsg.Caption = "¬Ì« „«Ì·Ìœ ê“«—‘ «‰ Œ«» ‘œÂ Õ–› ‘Êœ"
        frmMsg.fwBtn(0).ButtonType = flwButtonOk
        frmMsg.fwBtn(0).Caption = "»·Â"
        frmMsg.fwBtn(1).ButtonType = flwButtonCancel
        frmMsg.fwBtn(1).Caption = "ŒÌ—"
        frmMsg.Show vbModal
        If mvarMsgIdx = vbNo Then
            Exit Sub
        Else
            On Error GoTo ErrHandler
        flgerr = 2
        ReDim Parameter(0) As Parameter
        Parameter(0) = GenerateInputParameter("@intReportId", adInteger, 4, ReportId)
        Result = RunParametricStoredProcedure("Delete_tbltotal_ItemReports", Parameter)
        
        frmMsg.fwlblMsg.Caption = "ê“«—‘ „Ê—œ ‰Ÿ— »« „Ê›ﬁÌ  Õ–› ‘œ"
        frmMsg.fwBtn(0).Visible = False
        frmMsg.fwBtn(1).ButtonType = flwButtonCancel
        frmMsg.fwBtn(1).Caption = "ﬁ»Ê·"
        frmMsg.Show vbModal
        End If
    Else
        frmMsg.fwlblMsg.Caption = "¬Ì« „«Ì·Ìœ —ﬂÊ—œ «‰ Œ«»Ì «“ ê“«—‘ «‰ Œ«» ‘œÂ Õ–› ‘Êœ"
        frmMsg.fwBtn(0).ButtonType = flwButtonOk
        frmMsg.fwBtn(0).Caption = "»·Â"
        frmMsg.fwBtn(1).ButtonType = flwButtonCancel
        frmMsg.fwBtn(1).Caption = "ŒÌ—"
        frmMsg.Show vbModal
        If mvarMsgIdx = vbNo Then
            Exit Sub
        Else
            On Error GoTo ErrHandler
        flgerr = 2
        ReDim Parameter(1) As Parameter
        Parameter(0) = GenerateInputParameter("@intReportId", adInteger, 4, ReportId)
        Parameter(1) = GenerateInputParameter("@Row", adInteger, 4, Val(vsItemReportsDetails.TextMatrix(vsItemReportsDetails.Row, 0)))
        Result = RunParametricStoredProcedure("Delete_tblTotal_ItemReports_Details", Parameter)
        
        frmMsg.fwlblMsg.Caption = "—ﬂÊ—œ «‰ Œ«»Ì »« „Ê›ﬁÌ  Õ–› ‘œ"
        frmMsg.fwBtn(0).Visible = False
        frmMsg.fwBtn(1).ButtonType = flwButtonCancel
        frmMsg.fwBtn(1).Caption = "ﬁ»Ê·"
        frmMsg.Show vbModal
        End If
    End If
End If
DefaultSetting
ErrHandler:
If err.Number = -2147217873 Then
 If flgerr = 1 Then
    frmMsg.fwlblMsg.Caption = "«ÿ·«⁄«  „— »ÿ  »« ê—ÊÂ ê“«—‘«  ÊÃÊœœ«—œ" & vbLf & "ﬁ«œ— »Â Õ–› ‰„Ì »«‘Ìœ"
 ElseIf flgerr = 2 Then
    frmMsg.fwlblMsg.Caption = "«ÿ·«⁄«  „— »ÿ  »« ê“«—‘ ÊÃÊœœ«—œ" & vbLf & "ﬁ«œ— »Â Õ–› ‰„Ì »«‘Ìœ"
 End If
    frmMsg.fwBtn(0).Visible = False
    frmMsg.fwBtn(1).ButtonType = flwButtonCancel
    frmMsg.fwBtn(1).Caption = "ﬁ»Ê·"
    frmMsg.Show vbModal
End If
End Sub
