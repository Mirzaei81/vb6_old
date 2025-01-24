VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{7AEDC602-D94C-11D1-BB7A-00E0290EA3C9}#1.0#0"; "ResizeKit.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmHistory 
   ClientHeight    =   8985
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11565
   Icon            =   "frmHistory.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8985
   ScaleWidth      =   11565
   Begin VB.ComboBox cmbBranch2 
      BackColor       =   &H00009000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   2760
      RightToLeft     =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   25
      Top             =   2280
      Width           =   2115
   End
   Begin VB.ComboBox cmbBranch1 
      BackColor       =   &H00009000&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   7920
      RightToLeft     =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   22
      Top             =   2280
      Width           =   2115
   End
   Begin VB.CommandButton cmdSerach 
      BackColor       =   &H00008000&
      Caption         =   "‘—Ê⁄ Ã” ÃÊ"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "B Homa"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   120
      TabIndex        =   8
      Top             =   2400
      Width           =   2115
   End
   Begin VB.ComboBox cboActionType 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   7920
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   2760
      Width           =   2115
   End
   Begin VB.ComboBox cboFactorType 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   2760
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   2760
      Width           =   2115
   End
   Begin MSComCtl2.UpDown UpDownTo 
      Height          =   465
      Left            =   4680
      TabIndex        =   18
      Top             =   750
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   820
      _Version        =   393216
      Value           =   999999999
      BuddyControl    =   "txtTo"
      BuddyDispid     =   196614
      OrigLeft        =   4560
      OrigTop         =   570
      OrigRight       =   4815
      OrigBottom      =   1125
      Max             =   999999999
      SyncBuddy       =   -1  'True
      BuddyProperty   =   65547
      Enabled         =   -1  'True
   End
   Begin MSComCtl2.DTPicker DTPickerTo 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "HH:mm"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1065
         SubFormatType   =   4
      EndProperty
      Height          =   465
      Left            =   2760
      TabIndex        =   5
      Top             =   1680
      Width           =   2115
      _ExtentX        =   3731
      _ExtentY        =   820
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "HH:mm"
      Format          =   81592322
      CurrentDate     =   0.999988425925926
   End
   Begin MSComCtl2.DTPicker DTPickerFrom 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "HH:mm"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1065
         SubFormatType   =   4
      EndProperty
      Height          =   465
      Left            =   7920
      TabIndex        =   4
      Top             =   1680
      Width           =   2115
      _ExtentX        =   3731
      _ExtentY        =   820
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "hh:mm AMPM"
      Format          =   81592322
      CurrentDate     =   38486
   End
   Begin VB.TextBox txtTo 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   2760
      MaxLength       =   9
      TabIndex        =   1
      Top             =   720
      Width           =   2115
   End
   Begin VB.TextBox txtFrom 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   7920
      MaxLength       =   9
      TabIndex        =   0
      Top             =   750
      Width           =   2115
   End
   Begin VSFlex7LCtl.VSFlexGrid vsFactorHistory 
      Height          =   5595
      Left            =   0
      TabIndex        =   9
      Top             =   3240
      Width           =   11385
      _cx             =   20082
      _cy             =   9869
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
      BackColorFixed  =   16777152
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
      AllowUserResizing=   1
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
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmHistory.frx":A4C2
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
   Begin MSMask.MaskEdBox txtDateFrom 
      Height          =   465
      Left            =   7920
      TabIndex        =   2
      Top             =   1215
      Width           =   2115
      _ExtentX        =   3731
      _ExtentY        =   820
      _Version        =   393216
      MaxLength       =   8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "##/##/##"
      PromptChar      =   " "
   End
   Begin MSMask.MaskEdBox txtDateTo 
      Height          =   465
      Left            =   2760
      TabIndex        =   3
      Top             =   1215
      Width           =   2115
      _ExtentX        =   3731
      _ExtentY        =   820
      _Version        =   393216
      MaxLength       =   8
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "##/##/##"
      PromptChar      =   " "
   End
   Begin MSComCtl2.UpDown UpDownFrom 
      Height          =   465
      Left            =   9720
      TabIndex        =   19
      Top             =   750
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   820
      _Version        =   393216
      BuddyControl    =   "txtFrom"
      BuddyDispid     =   196615
      OrigLeft        =   9720
      OrigTop         =   660
      OrigRight       =   9975
      OrigBottom      =   1215
      Max             =   999999999
      SyncBuddy       =   -1  'True
      BuddyProperty   =   65547
      Enabled         =   -1  'True
   End
   Begin RESIZEKITLibCtl.ResizeKit ResizeKit1 
      Height          =   480
      Left            =   0
      OleObjectBlob   =   "frmHistory.frx":A5F0
      TabIndex        =   21
      Top             =   0
      Width           =   480
   End
   Begin VB.Label Label11 
      Alignment       =   1  'Right Justify
      Caption         =   " « ‘⁄»Â"
      BeginProperty Font 
         Name            =   "B Homa"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   4920
      RightToLeft     =   -1  'True
      TabIndex        =   24
      Top             =   2280
      Width           =   1305
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "«“ ‘⁄»Â"
      BeginProperty Font 
         Name            =   "B Homa"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   9960
      RightToLeft     =   -1  'True
      TabIndex        =   23
      Top             =   2280
      Width           =   1305
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "”Ê«»ﬁ ›«ò Ê—"
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
      Height          =   615
      Left            =   4560
      RightToLeft     =   -1  'True
      TabIndex        =   20
      Top             =   0
      Width           =   2655
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      Caption         =   "«“ ›«ò Ê—"
      BeginProperty Font 
         Name            =   "B Homa"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   10020
      RightToLeft     =   -1  'True
      TabIndex        =   17
      Top             =   750
      Width           =   1305
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      Caption         =   "«“  «—ÌŒ"
      BeginProperty Font 
         Name            =   "B Homa"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   10020
      RightToLeft     =   -1  'True
      TabIndex        =   16
      Top             =   1215
      Width           =   1305
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Caption         =   "«“ ”«⁄ "
      BeginProperty Font 
         Name            =   "B Homa"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   10020
      RightToLeft     =   -1  'True
      TabIndex        =   15
      Top             =   1680
      Width           =   1305
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "‰Ê⁄ ⁄„·Ì« "
      BeginProperty Font 
         Name            =   "B Homa"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   9960
      RightToLeft     =   -1  'True
      TabIndex        =   14
      Top             =   2880
      Width           =   1305
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "‰Ê⁄ ›«ò Ê—"
      BeginProperty Font 
         Name            =   "B Homa"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   4920
      RightToLeft     =   -1  'True
      TabIndex        =   13
      Top             =   2880
      Width           =   1305
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   " «  «—ÌŒ"
      BeginProperty Font 
         Name            =   "B Homa"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   4980
      RightToLeft     =   -1  'True
      TabIndex        =   12
      Top             =   1215
      Width           =   1305
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   " « ›«ò Ê—"
      BeginProperty Font 
         Name            =   "B Homa"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   4980
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   750
      Width           =   1305
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   " « ”«⁄ "
      BeginProperty Font 
         Name            =   "B Homa"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   4980
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   1680
      Width           =   1305
   End
End
Attribute VB_Name = "frmHistory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim clsDate As New clsDate
Dim i As Integer
Dim Parameter() As Parameter

Public Sub ExitForm()

    Unload Me
    
End Sub

Public Sub FillvsFactorHistory()

    Dim Rst As New ADODB.Recordset

    If Rst.State = 1 Then Rst.Close
    ReDim Parameter(9) As Parameter
    
    Parameter(0) = GenerateInputParameter("@FromNo", adBigInt, 8, Val(txtFrom.Text))
    Parameter(1) = GenerateInputParameter("@ToNo", adBigInt, 8, Val(txtTo.Text))
    Parameter(2) = GenerateInputParameter("@FromDate", adVarWChar, 50, CStr(IIf(Trim(txtDateFrom.ClipText) = "", "", Trim(txtDateFrom.Text))))
    Parameter(3) = GenerateInputParameter("@ToDate", adVarWChar, 50, IIf(Trim(txtDateTo.ClipText) = "", "", Trim(txtDateTo.Text)))
    Parameter(4) = GenerateInputParameter("@FromTime", adVarWChar, 50, CStr(FormatDateTime(DTPickerFrom.Value, vbShortTime)))
    Parameter(5) = GenerateInputParameter("@ToTime", adVarWChar, 50, CStr(FormatDateTime(DTPickerTo.Value, vbShortTime)))
    Parameter(6) = GenerateInputParameter("@ActionCode", adInteger, 4, cboActionType.ItemData(cboActionType.ListIndex))
    Parameter(7) = GenerateInputParameter("@Status", adInteger, 4, cboFactorType.ItemData(cboFactorType.ListIndex))
    Parameter(8) = GenerateInputParameter("@Branch1", adInteger, 4, cmbBranch1.ItemData(cmbBranch1.ListIndex))
    Parameter(9) = GenerateInputParameter("@Branch2", adInteger, 4, cmbBranch2.ItemData(cmbBranch2.ListIndex))
    
    Set Rst = RunParametricStoredProcedure2Rec("Get_History_By_Parameters", Parameter)
    With vsFactorHistory
        .Rows = 1
        i = 0
        While Rst.EOF <> True
            .Rows = .Rows + 1
            i = i + 1
            .TextMatrix(.Rows - 1, 0) = i
            .TextMatrix(.Rows - 1, 1) = Rst.Fields("Code").Value
            .TextMatrix(.Rows - 1, 2) = Rst.Fields("intSerialNo").Value
            .TextMatrix(.Rows - 1, 3) = Rst.Fields("No").Value
            .TextMatrix(.Rows - 1, 4) = Rst.Fields("Status").Value
            .TextMatrix(.Rows - 1, 5) = Rst.Fields("ActionCode").Value
            .TextMatrix(.Rows - 1, 6) = Rst.Fields("nvcFirstName").Value & " " & Rst.Fields("nvcSurName").Value
            .TextMatrix(.Rows - 1, 7) = Rst.Fields("RegDate").Value
            .TextMatrix(.Rows - 1, 8) = Rst.Fields("RegTime").Value
            .TextMatrix(.Rows - 1, 9) = Rst.Fields("Branch").Value
            
            Rst.MoveNext
        Wend
        .AutoSizeMode = flexAutoSizeColWidth
        .AutoSize 0, .Cols - 1
    End With
    
    
    Set Rst = Nothing
    
End Sub



Private Sub cmdSerach_Click()

    FillvsFactorHistory
End Sub

Private Sub Form_Activate()
    SetFirstToolBar
    VarActForm = Me.Name
    
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

    If ClsFormAccess.frmHistory = False Then
        Unload Me
        Exit Sub
    End If
    
    If intVersion = Min Then
        ShowDisMessage "‰„«Ì‘ ”Ê«»ﬁ ›«ﬂ Ê—Â« œ— ‰”ŒÂ Â«Ì »«·« — ÊÃÊœ œ«—œ", 1500
        Unload Me
        Exit Sub
    End If
    Dim Rst As New ADODB.Recordset
    
    CenterTop Me
    
    VarActForm = Me.Name
    
    txtDateFrom.Text = Mid(clsDate.shamsi(Date), 3)
    txtDateTo.Text = Mid(clsDate.shamsi(Date), 3)

   
    With vsFactorHistory
        .Rows = 1
'        .ColAlignment(-1) = flexAlignRightCenter
'        .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter
        .ColHidden(1) = True
        
        s = ""
        ReDim Parameter(1) As Parameter
        Parameter(0) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
        Parameter(1) = GenerateInputParameter("@ReturnType", adInteger, 4, 0)
        Set Rst = RunParametricStoredProcedure2Rec("Get_All_tStatusType", Parameter)
        s = .BuildComboList(Rst, "NvcDescription", "intStatusNo")
        .ColComboList(4) = s
        
''''        .ColComboList(4) = "#" & EnumFactorType.InvoiceReturn & "; »—ê‘  «“ ›—Ê‘|#" & EnumFactorType.PurchaseReturn & "; »—ê‘  «“ Œ—Ìœ|#" & EnumFactorType.Losses & ";÷«Ì⁄« |#" & EnumFactorType.invoice & ";›«ò Ê— ›—Ê‘|#" & EnumFactorType.Purchase & ";›«ò Ê— Œ—Ìœ"
        
        ReDim Parameter(0) As Parameter
        Parameter(0) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
        Set Rst = RunParametricStoredProcedure2Rec("Get_Action", Parameter)
        .ColComboList(5) = .BuildComboList(Rst, "ActionDescription", "ActionCode")
        .AutoSearch = flexSearchFromCursor
    
        Set Rst = RunStoredProcedure2RecordSet("Get_All_Branches")
        .ColComboList(9) = .BuildComboList(Rst, "nvcBranchName", "Branch")
    
    End With
    With cboActionType
        .Clear
        .AddItem "Â— ⁄„·Ì« Ì"
        .ItemData(0) = 0
        
        ReDim Parameter(0) As Parameter
        Parameter(0) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
        Set Rst = RunParametricStoredProcedure2Rec("Get_Action", Parameter)
        While Rst.EOF <> True
            
            .AddItem Rst!ActionDescription
            .ItemData(.ListCount - 1) = Rst!ActionCode
            Rst.MoveNext
        
        Wend
        If Rst.State <> 0 Then Rst.Close
        
        .ListIndex = 0

    End With
    
    With cboFactorType
    
        .Clear
        .AddItem "Â— ›«ò Ê—Ì"
        .ItemData(0) = 0
        ReDim Parameter(1) As Parameter
        Parameter(0) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
        Parameter(1) = GenerateInputParameter("@ReturnType", adInteger, 4, 0)
        Set Rst = RunParametricStoredProcedure2Rec("Get_All_tStatusType", Parameter)
        While Rst.EOF <> True
            
            .AddItem Rst!NvcDescription
            .ItemData(.ListCount - 1) = Rst!intStatusNo
            Rst.MoveNext
        
        Wend
        If Rst.State <> 0 Then Rst.Close
    
        .ListIndex = 0
    End With
    
    cmbBranch1.Clear
    cmbBranch2.Clear
    Set Rst = RunStoredProcedure2RecordSet("Get_All_Branches")
    Do While Rst.EOF = False
        cmbBranch1.AddItem Rst!nvcBranchName
        cmbBranch1.ItemData(cmbBranch1.NewIndex) = Rst!Branch
        cmbBranch2.AddItem Rst!nvcBranchName
        cmbBranch2.ItemData(cmbBranch2.NewIndex) = Rst!Branch
        Rst.MoveNext
    Loop
    Rst.Close
    For i = 0 To cmbBranch1.ListCount - 1
        If CurrentBranch = cmbBranch1.ItemData(i) Then
            cmbBranch1.ListIndex = i
            Exit For
        End If
    Next
    For i = 0 To cmbBranch2.ListCount - 1
        If CurrentBranch = cmbBranch2.ItemData(i) Then
            cmbBranch2.ListIndex = i
            Exit For
        End If
    Next
    
    
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


    Set Rst = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    VarActForm = ""
    If PosConnection.State = adStateOpen Then PosConnection.Close: Set PosConnection = Nothing
    SaveSetting strMainKey, Me.Name, "Left", Me.Left
    SaveSetting strMainKey, Me.Name, "Top", Me.Top


End Sub

Private Sub SetFirstToolBar()

    AllButton vbOff, True
    
    mdifrm.Toolbar1.Buttons(23).Enabled = True
    mdifrm.Toolbar1.Buttons(24).Enabled = True
    mdifrm.Toolbar1.Buttons(25).Enabled = True
    mdifrm.Toolbar1.Buttons(26).Enabled = True
    mdifrm.Toolbar1.Buttons(27).Enabled = True

End Sub

Private Sub ResizeKit1_ExitResize(ByVal XScale As Double, ByVal YScale As Double)

    If formloadFlag = True Then
        SaveSetting strMainKey, Me.Name, "Height", Me.Height
        SaveSetting strMainKey, Me.Name, "Width", Me.Width
    End If

End Sub

Private Sub vsFactorHistory_AfterSort(ByVal Col As Long, Order As Integer)
    For i = 1 To vsFactorHistory.Rows - 1
        vsFactorHistory.TextMatrix(i, 0) = i
    Next
End Sub

