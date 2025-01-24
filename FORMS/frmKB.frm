VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{7AEDC602-D94C-11D1-BB7A-00E0290EA3C9}#1.0#0"; "ResizeKit.ocx"
Object = "{75D4F767-8785-11D3-93AD-0000832EF44D}#2.14#0"; "FAST2001.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmKB 
   ClientHeight    =   9165
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12840
   Icon            =   "frmKB.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9165
   ScaleWidth      =   12840
   Begin VB.CommandButton SetBtnAscDefault 
      Caption         =   "„— » ”«“Ì »«‰ò «ÿ·«⁄« Ì »— «”«” Õ—Ê› «·›»«¡"
      BeginProperty Font 
         Name            =   "B Homa"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6480
      TabIndex        =   9
      Top             =   0
      Visible         =   0   'False
      Width           =   3615
   End
   Begin VB.ComboBox cboStations 
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
      Left            =   6480
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   600
      Width           =   2115
   End
   Begin VSFlex7LCtl.VSFlexGrid vsNoKeyBoardDefined 
      Height          =   7605
      Left            =   0
      TabIndex        =   4
      Top             =   1440
      Width           =   6405
      _cx             =   11298
      _cy             =   13414
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
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmKB.frx":A4C2
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
   Begin VSFlex7LCtl.VSFlexGrid vsKeyBoardDefined 
      Height          =   7635
      Left            =   6480
      TabIndex        =   3
      Top             =   1440
      Width           =   6255
      _cx             =   11033
      _cy             =   13467
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
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmKB.frx":A5A1
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
   Begin VB.TextBox txtKey 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   10680
      MaxLength       =   1
      TabIndex        =   1
      Top             =   480
      Width           =   585
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Height          =   840
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   6435
      _ExtentX        =   11351
      _ExtentY        =   1482
      ButtonWidth     =   2963
      ButtonHeight    =   1429
      Wrappable       =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "«Œ ’«’ ﬂ·Ìœ »Â ﬂ«·«Â«"
            Object.ToolTipText     =   "«Œ ’«’ ﬂ·Ìœ »Â ﬂ«·«Â«"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Ã«»Ã«ÌÌ ﬂ·ÌœÂ«"
            Object.ToolTipText     =   "Ã«»Ã«ÌÌ ﬂ·ÌœÂ«"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Õ–› ﬂ«·« «“ ﬂ·ÌœÂ«"
            Object.ToolTipText     =   "Õ–› ﬂ«·« «“ ﬂ·ÌœÂ«"
            ImageIndex      =   5
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   240
      Top             =   1800
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmKB.frx":A680
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmKB.frx":AF5C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmKB.frx":B838
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmKB.frx":C114
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmKB.frx":C9F0
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmKB.frx":D2CC
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmKB.frx":DBA8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin FLWCtrls.FWLabel3D FWLabel3D1 
      Height          =   375
      Left            =   6480
      Top             =   1080
      Width           =   6300
      _ExtentX        =   11113
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Titr"
         Size            =   9
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor1      =   -2147483647
      BackColor       =   12648447
      Caption         =   "        : ﬂ«·«Â«Ì „⁄—›Ì ‘œÂ —ÊÌ ﬂ·ÌœÂ« "
      Alignment       =   1
   End
   Begin FLWCtrls.FWLabel3D FWLabel3D2 
      Height          =   375
      Left            =   0
      Top             =   1080
      Width           =   6420
      _ExtentX        =   11324
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Titr"
         Size            =   9
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor1      =   -2147483647
      BackColor       =   12648384
      Caption         =   "        : ﬂ«·«Â«Ì „⁄—›Ì ‰‘œÂ —ÊÌ ﬂ·ÌœÂ«"
      Alignment       =   1
   End
   Begin RESIZEKITLibCtl.ResizeKit ResizeKit1 
      Height          =   480
      Left            =   9240
      OleObjectBlob   =   "frmKB.frx":E264
      TabIndex        =   8
      Top             =   120
      Width           =   480
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   " ⁄—Ì› ﬂ·ÌœÂ«Ì ﬂ«·«"
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
      Left            =   9960
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   0
      Width           =   2655
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
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
      Left            =   8640
      TabIndex        =   6
      Top             =   600
      Width           =   1005
   End
   Begin VB.Label lblPHazi 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ﬂ·Ìœ ﬂ«·«"
      BeginProperty Font 
         Name            =   "B Homa"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   375
      Left            =   11400
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   480
      Width           =   765
   End
End
Attribute VB_Name = "frmKB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i As Integer
Dim MyFormAddEditMode As KBEditMode
''dim ClsArya.StationNo As Integer
Dim ClsCnvKeyBoard As New ClsCnvKeyBoard
Dim TempKey, TempShift As Integer
Dim TempKeyCode, TempShiftKey As Integer
Dim tmpStationNo As Integer
Dim Parameter() As Parameter

Enum KBEditMode

    ViewKey = 1
    KeyToGood = 2
    ChangeKey = 4
    DeleteKey = 8
    
End Enum
Public Sub ChangeLanguage()

    FillvsKeyBoardDefined
    FillvsNoKeyBoardDefined

End Sub

Public Sub Cancel()

    For i = 1 To Toolbar1.Buttons.Count
        Toolbar1.Buttons(i).Enabled = True
    Next i
    
    txtKey.Locked = False
    TempKey = 0
    TempShift = 0
    MyFormAddEditMode = KBEditMode.ViewKey
    SetFirstToolBar
    
End Sub

Public Sub ExitForm()
    Unload Me
End Sub

Public Sub FillvsKeyBoardDefined()
    Dim Rst As New ADODB.Recordset
   
    With vsKeyBoardDefined
        .Rows = 1
        
        ReDim Parameter(2) As Parameter
        Parameter(0) = GenerateInputParameter("@StationId", adInteger, 4, tmpStationNo)
        Parameter(1) = GenerateInputParameter("@KeyCode", adInteger, 4, TempKeyCode)
        Parameter(2) = GenerateInputParameter("@ShiftKey", adInteger, 4, TempShiftKey)
        Set Rst = RunParametricStoredProcedure2Rec("Get_KeyBoard_Defined_Good", Parameter)
        
        If Rst.EOF = True And Rst.BOF = True Then
            Set Rst = Nothing
            Exit Sub
        End If
        While Rst.EOF <> True
            .Rows = .Rows + 1
            i = .Rows - 1
            .TextMatrix(i, 0) = i
            .TextMatrix(i, 1) = Rst.Fields("code").Value
            .TextMatrix(i, 2) = Rst.Fields("Level1").Value
            .TextMatrix(i, 3) = Rst.Fields("Level2").Value
            .TextMatrix(i, 4) = 0
            Select Case clsStation.Language
              Case Farsi
                .TextMatrix(i, 5) = Rst.Fields("Deslevel1").Value
                .TextMatrix(i, 6) = Rst.Fields("Deslevel2").Value
                .TextMatrix(i, 7) = Rst.Fields("Name").Value
              Case English
                .TextMatrix(i, 5) = Rst.Fields("LatinDeslevel1").Value
                .TextMatrix(i, 6) = Rst.Fields("LatinDeslevel2").Value
                .TextMatrix(i, 7) = Rst.Fields("LatinName").Value
            End Select
            Rst.MoveNext
        Wend
        
        .Row = 0
        .AutoSizeMode = flexAutoSizeColWidth
        .AutoSize 0, .Cols - 1
        
    End With
    Set Rst = Nothing
End Sub


Public Sub FillvsNoKeyBoardDefined()
    Dim Rst As New ADODB.Recordset
   
    With vsNoKeyBoardDefined
        .Rows = 1
        
        ReDim Parameter(0) As Parameter
        Parameter(0) = GenerateInputParameter("@StationId", adInteger, 4, tmpStationNo)
        Set Rst = RunParametricStoredProcedure2Rec("Get_No_KeyBoard_Defined_Good", Parameter)
        
        If Rst.EOF = True And Rst.BOF = True Then
            Set Rst = Nothing
            Exit Sub
        End If
        While Rst.EOF <> True
            .Rows = .Rows + 1
            i = .Rows - 1
            .TextMatrix(i, 0) = i
            .TextMatrix(i, 1) = Rst.Fields("code").Value
            .TextMatrix(i, 2) = Rst.Fields("Level1").Value
            .TextMatrix(i, 3) = Rst.Fields("Level2").Value
            .TextMatrix(i, 4) = 0
            Select Case clsStation.Language
               Case Farsi
                .TextMatrix(i, 5) = Rst.Fields("Deslevel1").Value
                .TextMatrix(i, 6) = Rst.Fields("Deslevel2").Value
                .TextMatrix(i, 7) = Rst.Fields("Name").Value
               Case English
                .TextMatrix(i, 5) = Rst.Fields("LatinDeslevel1").Value
                .TextMatrix(i, 6) = Rst.Fields("LatinDeslevel2").Value
                .TextMatrix(i, 7) = Rst.Fields("LatinName").Value
            End Select
            Rst.MoveNext
        Wend
        
        .Row = 0
        .AutoSizeMode = flexAutoSizeColWidth
        .AutoSize 0, .Cols - 1
        
    End With
    Set Rst = Nothing
End Sub

Public Sub SetFirstToolBar()
    
    FillvsKeyBoardDefined
    FillvsNoKeyBoardDefined
    
    AllButton vbOff, True
    
    mdifrm.Toolbar1.Buttons(23).Enabled = True
    mdifrm.Toolbar1.Buttons(24).Enabled = True
    mdifrm.Toolbar1.Buttons(25).Enabled = True
    mdifrm.Toolbar1.Buttons(26).Enabled = True
    mdifrm.Toolbar1.Buttons(27).Enabled = True
    
    mdifrm.Toolbar1.Buttons(9).Enabled = True
    
    For i = 1 To Toolbar1.Buttons.Count
         Toolbar1.Buttons.Item(i).Enabled = True
    Next i
    
    txtKey.Locked = False
    TempKey = 0
    TempShift = 0
    
    Select Case MyFormAddEditMode
        Case KBEditMode.KeyToGood
            Toolbar1.Buttons.Item(1).Enabled = False
        Case KBEditMode.ChangeKey
            Toolbar1.Buttons.Item(2).Enabled = False
        Case KBEditMode.DeleteKey
            Toolbar1.Buttons.Item(3).Enabled = False
            
    End Select

End Sub

Private Sub cboStations_Click()
    tmpStationNo = cboStations.ItemData(cboStations.ListIndex)
    MyFormAddEditMode = ViewKey
    SetFirstToolBar

End Sub

Private Sub Form_Activate()

    AllButton vbOff, True
    
    mdifrm.Toolbar1.Buttons(23).Enabled = True
    mdifrm.Toolbar1.Buttons(24).Enabled = True
    mdifrm.Toolbar1.Buttons(25).Enabled = True
    mdifrm.Toolbar1.Buttons(26).Enabled = True
    mdifrm.Toolbar1.Buttons(27).Enabled = True
    mdifrm.Toolbar1.Buttons(9).Enabled = True
    
    VarActForm = Me.Name
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case Shift
          Case 0
              Select Case KeyCode
                  Case 27  ' Esc
            
                    Me.ExitForm
                  Case Else
                    If Me.ActiveControl.Name <> txtKey.Name And txtKey.Locked = False Then
                        txtKey.SetFocus
                        txtKey_KeyDown KeyCode, Shift
                    End If
              End Select
          Case 2
               Select Case KeyCode
                  Case 123  'Exit
                     If clsStation.KeyboardType = EnumKeyBoardType.Rb2 Then
                        Me.ExitForm
                     End If
                  Case Else
                    If Me.ActiveControl.Name <> txtKey.Name And txtKey.Locked = False Then
                        txtKey.SetFocus
                        txtKey_KeyDown KeyCode, Shift
                    End If
              End Select

    End Select

End Sub

Private Sub Form_Load()

    
    If ClsFormAccess.frmKB = False Then
        Unload Me
        Exit Sub
    End If
    
    TempKey = 0
    TempShift = 0
    TempKeyCode = 0
    TempShiftKey = 0
    
    VarActForm = Me.Name
    CenterTop Me
    
    Dim Rst As New ADODB.Recordset
    Set Rst = RunStoredProcedure2RecordSet("Get_Pc_Stations")
    Dim i As Integer
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
    If i > clsArya.MaxStationNo And DebugMode = False And HardLockFlagTrial = False Then
       MsgBox "Œÿ« œ—  ⁄œ«œ «Ì” ê«ÂÂ«Ì Pc"
       End
    End If
    
    If cboStations.ListCount > 0 Then
        For i = 0 To cboStations.ListCount - 1
            If clsArya.StationNo = cboStations.ItemData(i) Then
                cboStations.ListIndex = i
                Exit For
            End If
        Next
    Else
        Unload Me
        Exit Sub
    End If
    Set Rst = Nothing
    
    With vsKeyBoardDefined
        
        .Rows = 1
        .Cols = 8
        .TextMatrix(0, 0) = "—œÌ›"
        .TextMatrix(0, 1) = "òœ ò«·«"
        .TextMatrix(0, 2) = "òœ ”ÿÕ «Ê· ò«·«"
        .TextMatrix(0, 3) = "òœ ”ÿÕ œÊ„ ò«·«"
        .TextMatrix(0, 4) = "«‰ Œ«»"
        .TextMatrix(0, 5) = "ê—ÊÂ «’·Ì"
        .TextMatrix(0, 6) = "“Ì— ê—ÊÂ"
        .TextMatrix(0, 7) = "‰«„ ò«·«"
        
        .ColDataType(4) = flexDTBoolean
      '  .ColHidden(1) = True
        .ColHidden(2) = True
        .ColHidden(3) = True
'        .ColHidden(4) = True
        .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter
        .ColAlignment(5) = flexAlignRightCenter
        .ColAlignment(6) = flexAlignRightCenter
        .ColAlignment(7) = flexAlignRightCenter
        .ColAlignment(0) = flexAlignCenterCenter
       
    End With
    
    
    With vsNoKeyBoardDefined
        
        .Rows = 1
        .Cols = 8
        .TextMatrix(0, 0) = "—œÌ›"
        .TextMatrix(0, 1) = "òœ ò«·«"
        .TextMatrix(0, 2) = "òœ ”ÿÕ «Ê· ò«·«"
        .TextMatrix(0, 3) = "òœ ”ÿÕ œÊ„ ò«·«"
        .TextMatrix(0, 4) = "«‰ Œ«»"
        .TextMatrix(0, 5) = "ê—ÊÂ «’·Ì"
        .TextMatrix(0, 6) = "“Ì— ê—ÊÂ"
        .TextMatrix(0, 7) = "‰«„ ò«·«"
        
        .ColDataType(4) = flexDTBoolean
       ' .ColHidden(1) = True
        .ColHidden(2) = True
        .ColHidden(3) = True
        .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter
        .ColAlignment(5) = flexAlignRightCenter
        .ColAlignment(6) = flexAlignRightCenter
        .ColAlignment(7) = flexAlignRightCenter
        .ColAlignment(0) = flexAlignCenterCenter
       
        .AutoSearch = flexSearchFromCursor
    End With

    MyFormAddEditMode = ViewKey
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
    VarActForm = ""
    If PosConnection.State = adStateOpen Then PosConnection.Close: Set PosConnection = Nothing
    Set ClsCnvKeyBoard = Nothing
    SaveSetting strMainKey, Me.Name, "Left", Me.Left
    SaveSetting strMainKey, Me.Name, "Top", Me.Top


    
End Sub


Private Sub ResizeKit1_ExitResize(ByVal XScale As Double, ByVal YScale As Double)


    If formloadFlag = True Then
        SaveSetting strMainKey, Me.Name, "Height", Me.Height
        SaveSetting strMainKey, Me.Name, "Width", Me.Width
    End If

End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call PresetScreenSaver
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call PresetScreenSaver
End Sub

Private Sub SetBtnAscDefault_Click()
    
    RunNonParametricStoredProcedure "Update_Good_btnAscDefault"
    frmMsg.fwlblMsg.Caption = " „— » ”«“Ì »«‰ò «ÿ·«⁄« Ì »— „»‰«Ì Õ—Ê› «·›»« «‰Ã«„ ‘œ "
    frmMsg.fwBtn(0).Visible = False
    frmMsg.fwBtn(1).ButtonType = flwButtonOk
    frmMsg.fwBtn(1).Caption = "ﬁ»Ê·"
    frmMsg.Show vbModal

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)

    MyFormAddEditMode = 2 ^ Button.index
            
    For i = 1 To Toolbar1.Buttons.Count
    
        If i <> Button.index Then
            Toolbar1.Buttons.Item(i).Enabled = True
        Else
            Toolbar1.Buttons.Item(i).Enabled = False
        End If
    
    Next i
    
    Select Case MyFormAddEditMode
        
        Case KBEditMode.KeyToGood
        
            frmMsg.fwlblMsg.Caption = "„—Õ·Â «Ê· : ·ÿ›« ﬂ·Ìœ „Ê—œ ‰Ÿ— —« «‰ Œ«» ‰„«ÌÌœ"
            frmMsg.fwBtn(0).Visible = False
            frmMsg.fwBtn(1).ButtonType = flwButtonOk
            frmMsg.fwBtn(1).Caption = "ﬁ»Ê·"
            frmMsg.Show vbModal
            
            txtKey.SetFocus
            
        Case KBEditMode.ChangeKey
        
            frmMsg.fwlblMsg.Caption = "·ÿ›« ﬂ·Ìœ —« »—«Ì  €ÌÌ— «‰ Œ«» ‰„«ÌÌœ"
            frmMsg.fwBtn(0).Visible = False
            frmMsg.fwBtn(1).ButtonType = flwButtonOk
            frmMsg.fwBtn(1).Caption = "ﬁ»Ê·"
            frmMsg.Show vbModal
            
            txtKey.SetFocus
        
        Case KBEditMode.DeleteKey
        
            frmMsg.fwlblMsg.Caption = " ·ÿ›« ﬂ·Ìœ Õ–›Ì —«  ⁄ÌÌ‰  ‰„«ÌÌœ"
            frmMsg.fwBtn(0).Visible = False
            frmMsg.fwBtn(1).ButtonType = flwButtonOk
            frmMsg.fwBtn(1).Caption = "ﬁ»Ê·"
            frmMsg.Show vbModal
            
            txtKey.SetFocus
    
    End Select
    
End Sub


Private Sub txtKey_Change()
    On Error Resume Next

    FillvsKeyBoardDefined
    
    If TempKeyCode = 0 Then Exit Sub
    Select Case MyFormAddEditMode
    
        Case KBEditMode.ViewKey
        
            MyFormAddEditMode = ViewKey
            SetFirstToolBar
            TempKeyCode = 0
            TempShiftKey = 0
            
        Case KBEditMode.KeyToGood
        
'            If txtKey.Text = "" Then Exit Sub
            
            frmMsg.fwlblMsg.Caption = "ﬂ«·«Ì „Ê—œ ‰Ÿ— —« «“ ÃœÊ· ﬂ«·«Â«Ì „⁄—›Ì ‰‘œÂ —ÊÌ ﬂ·Ìœ ﬂ«·« «‰ Œ«» ﬂ‰Ìœ  "
            frmMsg.fwBtn(0).Visible = False
            frmMsg.fwBtn(1).ButtonType = flwButtonOk
            frmMsg.fwBtn(1).Caption = "ﬁ»Ê·"
            frmMsg.Show vbModal
            
            vsNoKeyBoardDefined.SetFocus
            
        Case KBEditMode.ChangeKey
        
'            If txtKey.Text = "" Then Exit Sub
            
            If TempKey = 0 Then 'first key
            
                TempKey = TempKeyCode
                TempShift = TempShiftKey
                
                frmMsg.fwlblMsg.Caption = "ﬂ·Ìœ «‰ Œ«»Ì »Â ﬂœ«„ ﬂ·Ìœ «‰ ﬁ«· Ì«»œ"
                frmMsg.fwBtn(0).Visible = False
                frmMsg.fwBtn(1).ButtonType = flwButtonOk
                frmMsg.fwBtn(1).Caption = "ﬁ»Ê·"
                frmMsg.Show vbModal
                TempKeyCode = 0
                TempShiftKey = 0
            
            Else 'second key
            
                ReDim Parameter(4) As Parameter
                Parameter(0) = GenerateInputParameter("@StationId", adInteger, 4, tmpStationNo)
                Parameter(1) = GenerateInputParameter("@KeyCode1", adInteger, 4, TempKeyCode)
                Parameter(2) = GenerateInputParameter("@ShiftKey1", adInteger, 4, TempShiftKey)
                Parameter(3) = GenerateInputParameter("@KeyCode2", adInteger, 4, TempKey)
                Parameter(4) = GenerateInputParameter("@ShiftKey2", adInteger, 4, TempShift)
                RunParametricStoredProcedure "Exchange_key_In_tGood_KB", Parameter
                
                frmMsg.fwlblMsg.Caption = "Ã«»Ã«ÌÌ ò·ÌœÂ« »« „Ê›ﬁÌ  «‰Ã«„ ‘œ"
                frmMsg.fwBtn(0).Visible = False
                frmMsg.fwBtn(1).ButtonType = flwButtonOk
                frmMsg.fwBtn(1).Caption = "ﬁ»Ê·"
                frmMsg.Show vbModal
                
                TempKey = 0
                TempShift = 0
                
                MyFormAddEditMode = ViewKey
                SetFirstToolBar
            
            End If
        
        
    
        Case KBEditMode.DeleteKey
                    
            If vsKeyBoardDefined.Rows < 2 Then
                frmMsg.fwlblMsg.Caption = "»Â «Ì‰ ò·Ìœ ò«·«ÌÌ «Œ ’«’ œ«œÂ ‰‘œÂ «” "
                frmMsg.fwBtn(0).Visible = False
                frmMsg.fwBtn(1).ButtonType = flwButtonOk
                frmMsg.fwBtn(1).Caption = "ﬁ»Ê·"
                frmMsg.Show vbModal
                Exit Sub
            End If
            
            TempKey = TempKeyCode
            TempShift = TempShiftKey
            
            frmMsg.fwBtn(1).Caption = "ŒÌ—"
            frmMsg.fwBtn(0).Caption = "»·Â"
            frmMsg.fwlblMsg.Caption = "¬Ì« „Ì ŒÊ«ÂÌœ ﬂ· ﬂ«·«Â«Ì ﬂ·Ìœ —« Å«ﬂ ﬂ‰Ìœø "
            frmMsg.Show vbModal
            
            If modgl.mvarMsgIdx = vbYes Then
                
                ReDim Parameter(2) As Parameter
                Parameter(0) = GenerateInputParameter("@KeyCode", adInteger, 4, TempKey)
                Parameter(1) = GenerateInputParameter("@ShiftKey", adInteger, 4, TempShift)
                Parameter(2) = GenerateInputParameter("@StationId", adInteger, 4, tmpStationNo)
                RunParametricStoredProcedure "Delete_tGood_KB_By_KeyCode", Parameter
                    
''''                frmMsg.fwlblMsg.Caption = " . ﬂ·ÌÂ ﬂ«·«Â« «“ ﬂ·Ìœ ›Êﬁ Å«ﬂ ‘œ "
''''                frmMsg.Fwbtn(0).Visible = False
''''                frmMsg.Fwbtn(1).ButtonType = flwButtonOk
''''                frmMsg.Fwbtn(1).Caption = "ﬁ»Ê·"
''''                frmMsg.Show vbModal
                    frmDisMsg.Timer1.Interval = 1000
                    frmDisMsg.lblMessage = "  ﬂ·ÌÂ ﬂ«·«Â« «“ ﬂ·Ìœ ›Êﬁ Å«ﬂ ‘œ"
                    frmDisMsg.Timer1.Enabled = True
                    frmDisMsg.Show vbModal
                
                MyFormAddEditMode = ViewKey
                SetFirstToolBar
                
            Else
            
                frmMsg.fwlblMsg.Caption = "ﬂ«·«Ì „Ê—œ ‰Ÿ— —« «“ ÃœÊ· ﬂ«·«Â«Ì „⁄—›Ì ‘œÂ —ÊÌ ﬂ·Ìœ ﬂ«·« «‰ Œ«» ﬂ‰Ìœ  "
                frmMsg.fwBtn(0).Visible = False
                frmMsg.fwBtn(1).ButtonType = flwButtonOk
                frmMsg.fwBtn(1).Caption = "ﬁ»Ê·"
                frmMsg.Show vbModal
                txtKey.Locked = True
                
                
            End If
            
    End Select


End Sub

Private Sub txtKey_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode < 47 Or KeyCode = 144 Then Exit Sub
    If txtKey.Locked <> True Then
        
''''        If KeyCode = 8 Then
''''            Exit Sub
''''        End If
        TempKeyCode = 0
        TempShiftKey = 0
        txtKey.SelStart = 0
        txtKey.SelLength = Len(txtKey.Text)
            
''''            If (KeyAscii >= 127) Or Val(GetKbLayout) = Val(LANG_Pr_IR) Then
''''
''''                KeyAscii = ClsCnvKeyBoard.CnvKeyBoard(KeyAscii)
''''
''''            End If
            
            If IsUserDefinedKey(KeyCode, Shift) = False Then
            
                KeyCode = 0
                TempKeyCode = KeyCode
                TempShiftKey = Shift
                frmDisMsg.lblMessage = "«Ì‰ ò·Ìœ Ã“ ò·ÌœÂ«Ì —“—Ê ‘œÂ „Ì »«‘œ"
                frmDisMsg.Timer1.Enabled = True
                frmDisMsg.Show vbModal
                txtKey.SelText = ""
                Exit Sub
                
            Else
                TempKeyCode = KeyCode
                TempShiftKey = Shift
                FillvsKeyBoardDefined
                
                txtKey.Text = Chr(KeyCode)
                
                If (KeyCode = 78 And Shift = 2) Then   ' Ctrl + N
                   txtKey_Change
                End If
                If (KeyCode = 121 Or KeyCode = 122 Or KeyCode = 123) And Shift = 1 Then    ' Shift  + f10 ~ Shift f12
                   txtKey_Change
                End If
              '  txtKey.SelStart = 0
              '  txtKey.SelLength = Len(txtKey.Text)
            End If
    End If

End Sub

Private Sub vsKeyBoardDefined_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    
    With vsKeyBoardDefined
    
        Select Case MyFormAddEditMode
        
            Case KBEditMode.DeleteKey
                
                Dim Parameter(1) As Parameter
                Parameter(0) = GenerateInputParameter("@GoodCode", adInteger, 4, .TextMatrix(.Row, 1))
                Parameter(1) = GenerateInputParameter("@StationId", adInteger, 4, tmpStationNo)
                RunParametricStoredProcedure "Delete_tGood_KB_By_GoodCode", Parameter
                
                frmMsg.fwlblMsg.Caption = "ò«·«Ì „Ê—œ ‰Ÿ— «“ ò·Ìœ Õ–› ‘œ"
                frmMsg.fwBtn(0).Visible = False
                frmMsg.fwBtn(1).ButtonType = flwButtonOk
                frmMsg.fwBtn(1).Caption = "ﬁ»Ê·"
                frmMsg.Show vbModal
                
                MyFormAddEditMode = ViewKey
                SetFirstToolBar
                
        End Select
    End With

End Sub

Private Sub vsKeyBoardDefined_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    With vsKeyBoardDefined
    
        If Button <> 1 Or .Rows < 2 Or .Row < 1 Or .Col <> 4 Or TempKey = 0 Then Exit Sub
        
        Select Case MyFormAddEditMode
        
            Case KBEditMode.DeleteKey
            
                    .Select .Row, .Col
                    .EditCell
                
                If Val(.TextMatrix(.Row, .Col)) = -1 Then
                    For i = 1 To .Rows - 1
                        If i <> .Row Then
                            .TextMatrix(i, 4) = 0
                        End If
                    Next i
                End If
                
        End Select
        
    End With

End Sub

Private Sub vsNoKeyBoardDefined_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    With vsNoKeyBoardDefined
    
        Select Case MyFormAddEditMode
            Case KBEditMode.KeyToGood
            
'                If Val(.TextMatrix(Row, Col)) = -1 And txtKey.Text <> "" Then
                If Val(.TextMatrix(Row, Col)) = -1 And TempKeyCode <> 0 Then
                    ReDim Parameter(3) As Parameter
                    
                    Parameter(0) = GenerateInputParameter("@GoodCode", adInteger, 4, .TextMatrix(Row, 1))
                    Parameter(1) = GenerateInputParameter("@KeyCode", adInteger, 4, TempKeyCode)
                    Parameter(2) = GenerateInputParameter("@ShiftKey", adInteger, 4, TempShiftKey)
                    Parameter(3) = GenerateInputParameter("@StationId", adInteger, 4, tmpStationNo)
                    
                    RunParametricStoredProcedure "Insert_tGood_KB", Parameter
                    
''''                    frmMsg.fwlblMsg.Caption = "ò«·« »Â ò·Ìœ «Œ ’«’ œ«œÂ ‘œ"
''''                    frmMsg.Fwbtn(0).Visible = False
''''                    frmMsg.Fwbtn(1).ButtonType = flwButtonOk
''''                    frmMsg.Fwbtn(1).Caption = "ﬁ»Ê·"
''''                    frmMsg.Show vbModal
                    frmDisMsg.Timer1.Interval = 1000
                    frmDisMsg.lblMessage = " ò«·« »Â ò·Ìœ «Œ ’«’ œ«œÂ ‘œ"
                    frmDisMsg.Timer1.Enabled = True
                    frmDisMsg.Show vbModal
                    
                    MyFormAddEditMode = ViewKey
                    SetFirstToolBar
                    TempKeyCode = 0
                    TempShiftKey = 0
                End If
        End Select
        
    End With

End Sub
'
'Private Sub vsNoKeyBoardDefined_KeyDown(KeyCode As Integer, Shift As Integer)
'
'    With vsNoKeyBoardDefined
'
'        If KeyCode <> 32 Or .Rows < 2 Or .Row < 1 Or .Col <> 4 Or txtKey.Text = "" Then Exit Sub
'
'        Select Case MyFormAddEditMode
'            Case KBEditMode.KeyToGood
'                .Select .Row, .Col
'                .EditCell
'
'                If Val(.TextMatrix(.Row, .Col)) = -1 Then
'                    For i = 1 To .Rows - 1
'                        If i <> .Row Then
'                            .TextMatrix(i, 4) = 0
'                        End If
'                    Next i
'                End If
'
'        End Select
'
'
'    End With
'
'End Sub

Private Sub vsNoKeyBoardDefined_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    With vsNoKeyBoardDefined
    
'        If Button <> 1 Or .Rows < 2 Or .Row < 1 Or .Col <> 4 Or txtKey.Text = "" Then Exit Sub
        If Button <> 1 Or .Rows < 2 Or .Row < 1 Or .Col <> 4 Or TempKeyCode = 0 Then Exit Sub
        
        Select Case MyFormAddEditMode
            Case KBEditMode.KeyToGood
                .Select .Row, .Col
                .EditCell
                
                If Val(.TextMatrix(.Row, .Col)) = -1 Then
                    For i = 1 To .Rows - 1
                        If i <> .Row Then
                            .TextMatrix(i, 4) = 0
                        End If
                    Next i
                End If
                
        End Select
        
        
    End With
End Sub
