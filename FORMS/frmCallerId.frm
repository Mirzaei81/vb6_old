VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{7AEDC602-D94C-11D1-BB7A-00E0290EA3C9}#1.0#0"; "ResizeKit.ocx"
Object = "{687FF23F-0E9B-449D-A782-2C5E7116EDB0}#1.0#0"; "Fardate.ocx"
Begin VB.Form frmCallerId 
   BackColor       =   &H00FFC0C0&
   Caption         =   "                                    ·Ì”   „«” Â«                        "
   ClientHeight    =   6720
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7245
   Icon            =   "FrmCallerId.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   6720
   ScaleWidth      =   7245
   Begin VB.CommandButton CmdSave 
      BackColor       =   &H00008000&
      Caption         =   "À»  œ” Ì"
      Enabled         =   0   'False
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
      Left            =   120
      TabIndex        =   6
      Top             =   6000
      Width           =   1575
   End
   Begin VB.PictureBox Frame3 
      BackColor       =   &H00FFC0C0&
      Height          =   615
      Left            =   1800
      RightToLeft     =   -1  'True
      ScaleHeight     =   555
      ScaleWidth      =   5235
      TabIndex        =   0
      Top             =   0
      Width           =   5295
      Begin VB.OptionButton Option1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFC0C0&
         Caption         =   " „«” Â«Ì »œÊ‰ ÃÊ«»"
         BeginProperty Font 
            Name            =   "Nazanin"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   2
         Top             =   120
         Width           =   2655
      End
      Begin VB.OptionButton Option1 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFC0C0&
         Caption         =   "ò·ÌÂ  „«” Â«Ì —Ê“"
         BeginProperty Font 
            Name            =   "Nazanin"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   2880
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   120
         Value           =   -1  'True
         Width           =   2295
      End
   End
   Begin RESIZEKITLibCtl.ResizeKit ResizeKit1 
      Height          =   480
      Left            =   0
      OleObjectBlob   =   "FrmCallerId.frx":A4C2
      TabIndex        =   3
      Top             =   0
      Width           =   480
   End
   Begin VSFlex7LCtl.VSFlexGrid vsCallerId 
      Height          =   5145
      Left            =   0
      TabIndex        =   4
      Top             =   720
      Width           =   7155
      _cx             =   12621
      _cy             =   9075
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Nazanin"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MousePointer    =   0
      BackColor       =   12648447
      ForeColor       =   -2147483640
      BackColorFixed  =   8454143
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   12648447
      BackColorAlternate=   12648447
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   3
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   3
      SelectionMode   =   3
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   3
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   500
      RowHeightMax    =   500
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"FrmCallerId.frx":A548
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
   Begin FarDate1.FarDate FarDate1 
      Height          =   500
      Left            =   0
      TabIndex        =   5
      Top             =   100
      Visible         =   0   'False
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   873
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Object.ToolTipText     =   "FarDate"
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "»« ò·Ìò —ÊÌ  —œÌ› ” Ê‰ ‰«„ „‘ —Ì  „Ì  Ê«‰Ìœ »Â ’Ê—  œ” Ì ‰«„ „‘ —Ì —« Ê«—œ ò‰Ìœ."
      BeginProperty Font 
         Name            =   "Nazanin"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004080&
      Height          =   735
      Left            =   1800
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   6000
      Width           =   5295
   End
End
Attribute VB_Name = "FrmCallerId"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim i As Integer
Dim Parameter() As Parameter

Private Sub CmdSave_Click()
    vsCallerId_ValidateEdit vsCallerId.Row, vsCallerId.Col, True
    Update
    ShowDisMessage " €ÌÌ— «ÿ·«⁄«  «‰Ã«„ ê—› ", 1000
    CmdSave.Enabled = False

End Sub

Private Sub FarDate1_Change()
    DefaultSetting
End Sub
Public Sub ShowActiveRow()
    With vsCallerId
        .Cell(flexcpBackColor, .Rows - 1, .FixedCols, .Rows - 1, .Cols - 1) = 8421631
    End With
End Sub
Private Sub Form_Activate()
    
    mdifrm.Toolbar3.Visible = False
'    SetFirstToolBar
 If LastRecordshow = True Then ShowActiveRow
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

    With vsCallerId
        .Cols = 8
        .TextMatrix(0, 0) = "—œÌ›"
        .TextMatrix(0, 1) = "“„«‰"
        .TextMatrix(0, 2) = "Œÿ"
        .TextMatrix(0, 3) = " ·›‰ "
        .TextMatrix(0, 4) = "òœ "
        .TextMatrix(0, 5) = "«‘ —«ò "
        .TextMatrix(0, 6) = "„‘ —Ì "
        .TextMatrix(0, 7) = "AutoId "

        .ColHidden(1) = False
        .ColHidden(4) = True
        .ColHidden(7) = True

        .FixedAlignment(-1) = flexAlignCenterCenter
        .ColAlignment(0) = flexAlignCenterCenter
        .ColAlignment(1) = flexAlignCenterCenter
        .ColAlignment(2) = flexAlignCenterCenter
        .ColAlignment(3) = flexAlignCenterCenter
        .ColAlignment(5) = flexAlignCenterCenter
        .ColAlignment(6) = flexAlignRightCenter
         For i = 0 To .Cols - 1
            .ColWidth(i) = Val(GetSetting(strMainKey, "frmCallerId_vsCallerId", "Col" & i))
            If .ColWidth(i) = 0 Then
                .ColWidth(i) = .Width / 6     'Row
            End If
         Next i

    End With

'    MyFormAddEditMode = ViewMode
    DefaultSetting
'    SetFirstToolBar

    formloadFlag = False
    Me.Left = Val(GetSetting(strMainKey, Me.Name, "Left"))
    If Val(GetSetting(strMainKey, Me.Name, "Height")) > 0 Then
        Me.Height = Val(GetSetting(strMainKey, Me.Name, "Height"))
    End If
    If Val(GetSetting(strMainKey, Me.Name, "Width")) > 0 Then
        Me.Width = Val(GetSetting(strMainKey, Me.Name, "Width"))
    End If
    Me.Top = Val(GetSetting(strMainKey, Me.Name, "Top"))
    formloadFlag = True
    FarDate1.Visible = True
    FarDate1.Text = "13" + mvarDate
       
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If PosConnection.State = adStateOpen Then PosConnection.Close: Set PosConnection = Nothing
    
    SaveSetting strMainKey, Me.Name, "Left", Me.Left
    SaveSetting strMainKey, Me.Name, "Top", Me.Top
    LastRecordshow = False
End Sub
Public Sub ExitForm()

    Unload Me
    
End Sub
Public Sub Edit()
    If vsCallerId.Rows > 1 Then
    End If
End Sub

Public Sub Delete()

'    If vsCallerId.Rows < 2 Then Exit Sub
'
'    If MyFormAddEditMode <> 0 Then
'        Cancel
'    End If
'    On Error GoTo ErrHandler
'    ReDim Parameter(0) As Parameter
'    Parameter(0) = GenerateInputParameter("@intId", adInteger, 4, Text1.Tag)
'    RunParametricStoredProcedure "Delete_tBank_By_tintBank", Parameter
'
'    frmMsg.fwlblMsg.Caption = "»« „Ê›ﬁÌ  Õ–› ‘œ"
'    frmMsg.fwBtn(0).Visible = False
'    frmMsg.fwBtn(1).ButtonType = flwButtonOk
'    frmMsg.fwBtn(1).Caption = "ﬁ»Ê·"
'    frmMsg.Show vbModal
'
'    DefaultSetting
'Exit Sub
'
'ErrHandler:
'If err.Number = -2147217873 Then
'
'    frmMsg.fwlblMsg.Caption = "„ «”›«‰Â ‘„« ﬁ«œ— »Â Õ–› ‰„Ì »«‘Ìœ"
'    frmMsg.fwBtn(0).Visible = False
'    frmMsg.fwBtn(1).ButtonType = flwButtonOk
'    frmMsg.fwBtn(1).Caption = "ﬁ»Ê·"
'    frmMsg.Show vbModal
'End If
    
End Sub

Public Sub DefaultSetting()

    Dim Rst As New ADODB.Recordset
    
    
    ReDim Parameter(1) As Parameter
    Parameter(0) = GenerateInputParameter("@Flag", adInteger, 4, IIf(Option1(0).Value = True, 0, 1))
    Parameter(1) = GenerateInputParameter("@nvcDate", adWChar, 8, Mid(FarDate1.Text, 3, 8))
    Set Rst = RunParametricStoredProcedure2Rec("Get_tblTotal_CallerId", Parameter)
    
    With vsCallerId
        .Rows = 1
        If Not (Rst.BOF = True And Rst.EOF = True) Then
            While Rst.EOF <> True
                .Rows = .Rows + 1
                .TextMatrix(.Rows - 1, 0) = Rst!intRow
                .TextMatrix(.Rows - 1, 1) = Rst!nvcTime
                .TextMatrix(.Rows - 1, 2) = Rst!LineNumber
                .TextMatrix(.Rows - 1, 3) = Trim(Rst!nvcCallerId)
                .TextMatrix(.Rows - 1, 4) = IIf(IsNull(Rst!intCustomer), "", Rst!intCustomer)
                .TextMatrix(.Rows - 1, 5) = IIf(IsNull(Rst!MembershipId), "", Rst!MembershipId)
                .TextMatrix(.Rows - 1, 6) = IIf(IsNull(Rst!nvcName), "", Trim(Rst!nvcName))
                .TextMatrix(.Rows - 1, 7) = Rst!AutoId
                Rst.MoveNext
            Wend
            .ShowCell .Rows - 1, 1
        End If
    
    End With
    
    If Rst.State = 1 Then Rst.Close
     
    Set Rst = Nothing
    
End Sub
Public Sub Add()
    
    
End Sub

Public Sub Cancel()
    
End Sub

Public Sub Update()

    ReDim Parameter(3) As Parameter
    
    Parameter(0) = GenerateInputParameter("@AutoId", adInteger, 4, Val(vsCallerId.TextMatrix(vsCallerId.Row, 7)))
    Parameter(1) = GenerateInputParameter("@intCustomer", adInteger, 4, IIf(Val(frmInvoice.lblCustomer.Tag) = -1, Null, Val(frmInvoice.lblCustomer.Tag)))
    Parameter(2) = GenerateInputParameter("@MembershipId", adInteger, 4, IIf(mvarMemberShipId = 0, Null, mvarMemberShipId))
    Parameter(3) = GenerateInputParameter("@nvcname", adWChar, 50, IIf(Trim(frmInvoice.lblCustomer.Caption) = "€Ì— „‘ —ò", Trim(vsCallerId.TextMatrix(vsCallerId.Row, 6)), frmInvoice.lblCustomer.Caption))
    
    RunParametricStoredProcedure "Update_tblTotal_CallerId", Parameter
    
    DefaultSetting

End Sub

Private Sub Option1_Click(Index As Integer)
    DefaultSetting
End Sub

Private Sub ResizeKit1_ExitResize(ByVal XScale As Double, ByVal YScale As Double)
    
    If formloadFlag = True Then
        SaveSetting strMainKey, Me.Name, "Height", Me.Height
        SaveSetting strMainKey, Me.Name, "Width", Me.Width
    End If

End Sub

Private Sub vsCallerId_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    
    For i = 0 To vsCallerId.Cols - 1
        SaveSetting strMainKey, "frmCallerId_vsCallerId", "Col" & i, vsCallerId.ColWidth(i)
    Next

End Sub

Private Sub vsCallerId_Click()
    
    With vsCallerId
        If .Row = 0 Then Exit Sub
        If Val(.TextMatrix(.Row, 4)) = 0 And .Col <> 6 Then
'            Call_RealNumber = .TextMatrix(.Row, 3)
'            frmInvoice.FindCust
            mvarcode = 0
            Call frmInvoice.FWModem_MouseDown(Val(.TextMatrix(.Row, 2)) - 1, 1, 0, 10, 10)
            If frmInvoice.lblCustomer.Tag > 0 Then
                Update
                frmInvoice.ChkCallerId.Value = False
                Unload Me
            End If
        ElseIf Val(.TextMatrix(.Row, 4)) = 0 And .Col = 6 Then
            .Select .Row, .Col
            .EditCell
            CmdSave.Enabled = True
        Else
            ShowDisMessage "«Ì‰  ·›‰ ﬁ»·« ÃÊ«» œ«œÂ ‘œÂ", 1000
        End If
    
    
    End With
    
End Sub

Private Sub vsCallerId_DblClick()
    vsCallerId_Click
End Sub

Private Sub vsCallerId_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsCallerId
        .Row = Row
        .Col = Col
    End With
    
End Sub
