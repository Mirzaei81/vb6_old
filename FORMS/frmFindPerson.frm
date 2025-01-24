VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{7AEDC602-D94C-11D1-BB7A-00E0290EA3C9}#1.0#0"; "ResizeKit.ocx"
Begin VB.Form frmFindPerson 
   BackColor       =   &H80000016&
   Caption         =   "                                               Ã” ÃÊÌ Å—”‰·"
   ClientHeight    =   9015
   ClientLeft      =   2775
   ClientTop       =   3870
   ClientWidth     =   8190
   Icon            =   "frmFindPerson.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   9015
   ScaleWidth      =   8190
   Begin VB.ComboBox cmbBranch 
      BackColor       =   &H8000000F&
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
      Left            =   4680
      RightToLeft     =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   480
      Width           =   2475
   End
   Begin VB.CommandButton OKButton 
      BackColor       =   &H0000C000&
      Caption         =   "ﬁ»Ê·"
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
      Height          =   615
      Left            =   1680
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   8280
      Width           =   975
   End
   Begin VB.CommandButton CancelButton 
      BackColor       =   &H000000C0&
      Cancel          =   -1  'True
      Caption         =   "«‰’—«›"
      BeginProperty Font 
         Name            =   "B Homa"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   480
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   8280
      Width           =   975
   End
   Begin VB.TextBox txtName 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000018&
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
      Left            =   4440
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   1080
      Width           =   2715
   End
   Begin VSFlex7LCtl.VSFlexGrid vsPerson 
      Height          =   6525
      Left            =   480
      TabIndex        =   1
      Top             =   1680
      Width           =   7275
      _cx             =   12832
      _cy             =   11509
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
      Rows            =   2
      Cols            =   5
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   500
      RowHeightMax    =   500
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmFindPerson.frx":A4C2
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
      ExplorerBar     =   5
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
   Begin RESIZEKITLibCtl.ResizeKit ResizeKit1 
      Height          =   480
      Left            =   0
      OleObjectBlob   =   "frmFindPerson.frx":A553
      TabIndex        =   4
      Top             =   0
      Width           =   480
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "‘⁄»Â"
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
      Height          =   435
      Left            =   7080
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   480
      Width           =   585
   End
   Begin VB.Label LblFindPerson 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "B Homa"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   495
      Left            =   720
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   0
      Width           =   6975
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
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
      ForeColor       =   &H8000000D&
      Height          =   435
      Left            =   6960
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   1080
      Width           =   585
   End
End
Attribute VB_Name = "frmFindPerson"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim i As Integer

Private Sub CancelButton_Click()
    mvarcode = 0
    Unload Me
End Sub

Private Sub cmbBranch_Click()
    FillvsPerson
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call PresetScreenSaver
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call PresetScreenSaver
End Sub

Private Sub Form_Load()

    CenterCenterinSecondScreen Me
    
    mvarcode = 0
    

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
    FillvsPerson
    vsPerson.Row = 0


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
    For i = 0 To cmbBranch.ListCount - 1
        If CurrentBranch = cmbBranch.ItemData(i) Then
            cmbBranch.ListIndex = i
'        If CurrentBranch = cmbBranch.ItemData(cmbBranch.ListIndex) Then
            Exit For
        End If
    Next

End Sub

Private Sub Form_Unload(Cancel As Integer)
    If PosConnection.State = adStateOpen Then PosConnection.Close: Set PosConnection = Nothing

    SaveSetting strMainKey, Me.Name, "Left", Me.Left
    SaveSetting strMainKey, Me.Name, "Top", Me.Top


End Sub

Private Sub OKButton_Click()
    If vsPerson.Row > 0 Then
        mvarcode = vsPerson.TextMatrix(vsPerson.Row, 1)
        mvarName = vsPerson.TextMatrix(vsPerson.Row, 3)
        mvarBranch = cmbBranch.ItemData(cmbBranch.ListIndex)
        
    Else
        mvarcode = 0
        mvarName = ""
    End If
    Unload Me

End Sub

Private Sub ResizeKit1_ExitResize(ByVal XScale As Double, ByVal YScale As Double)
    If formloadFlag = True Then
        SaveSetting strMainKey, Me.Name, "Height", Me.Height
        SaveSetting strMainKey, Me.Name, "Width", Me.Width
    End If

End Sub

Private Sub txtName_Change()

    i = vsPerson.FindRow(TxtName.Text, 1, 3, False, False)
    If i > 0 Then
        vsPerson.Row = i
        vsPerson.ShowCell i, 0
        LblFindPerson.Caption = ""
    Else
        vsPerson.Row = 0
        vsPerson.ShowCell 0, 0
        If TxtName.Text <> "" Then
           LblFindPerson.Caption = " ‰«„ ( " & TxtName.Text & " ) œ— ”Ì” „  ⁄—Ì› ‰‘œÂ"
         Else
            LblFindPerson.Caption = "‰«„ Å—”‰· —« Ê«—œ ﬂ‰Ìœ  "
         End If

    End If

End Sub

Private Sub txtName_GotFocus()

    vsPerson.Row = 0
    vsPerson.Select vsPerson.Row, 3
    vsPerson.Sort = flexSortGenericAscending
    LblFindPerson.Caption = "‰«„ Å—”‰· —« Ê«—œ ﬂ‰Ìœ  "

End Sub

Private Sub FillvsPerson()
   ReDim Parameter(1) As Parameter
   Dim Rst As New ADODB.Recordset
    
    Parameter(0) = GenerateInputParameter("@AccessLevelCurrentUser", adInteger, 4, mVarAccessLevel)
    Parameter(1) = GenerateInputParameter("@Branch", adInteger, 4, cmbBranch.ItemData(cmbBranch.ListIndex))
    Set Rst = RunParametricStoredProcedure2Rec("Get_tblTotal_tPer_ByAccessLevel", Parameter)
    
    With vsPerson
        .Rows = 1
        i = 0
        While Rst.EOF <> True
            .Rows = .Rows + 1
            i = .Rows - 1
            .TextMatrix(i, 0) = i
            .TextMatrix(i, 1) = Rst!ppno
            .TextMatrix(i, 2) = Rst!PersonnelNumber
            .TextMatrix(i, 3) = Rst!PersonName
            .TextMatrix(i, 4) = Rst!ActDeAct
             .ColDataType(4) = flexDTBoolean
            Rst.MoveNext
        Wend
    End With
    Set Rst = Nothing
    vsPerson.Cell(flexcpAlignment, 0, 0, 0, 3) = flexAlignCenterCenter
    vsPerson.ColAlignment(-1) = flexAlignCenterCenter


End Sub


Private Sub vsPerson_AfterSort(ByVal Col As Long, Order As Integer)
    For i = 1 To vsPerson.Rows - 1
        vsPerson.TextMatrix(i, 0) = i
    Next
    
End Sub

Private Sub vsPerson_DblClick()
    If vsPerson.Row > 0 Then
        OKButton_Click
    End If
End Sub
