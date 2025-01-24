VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{75D4F767-8785-11D3-93AD-0000832EF44D}#2.14#0"; "FAST2001.ocx"
Begin VB.Form frmPrizeType 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6360
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5055
   BeginProperty Font 
      Name            =   "Nazanin"
      Size            =   12
      Charset         =   178
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPrizeType.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6360
   ScaleWidth      =   5055
   Begin VB.TextBox txtPrizeType 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Nazanin"
         Size            =   12
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   570
      Left            =   240
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   840
      Width           =   2865
   End
   Begin VSFlex7LCtl.VSFlexGrid vsPrizeType 
      Height          =   4635
      Left            =   120
      TabIndex        =   0
      Top             =   1560
      Width           =   4635
      _cx             =   8176
      _cy             =   8176
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
      BackColor       =   16777215
      ForeColor       =   -2147483640
      BackColorFixed  =   16761024
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   16777215
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
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   3
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   500
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmPrizeType.frx":A4C2
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
      ExplorerBar     =   3
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
      Height          =   525
      Left            =   3600
      Top             =   0
      Width           =   1425
      _ExtentX        =   2514
      _ExtentY        =   926
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   178
         Weight          =   700
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
   Begin VB.Label lblPrizeType 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "* ‰Ê⁄ Ã«Ì“Â"
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
      Height          =   405
      Left            =   3240
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   960
      Width           =   1365
   End
   Begin VB.Label lblTitel 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   " ⁄—Ì› ‰Ê⁄ Ã«Ì“Â"
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
      Left            =   1200
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   0
      Width           =   2415
   End
End
Attribute VB_Name = "frmPrizeType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private clsDate As New clsDate
Private Rc As New ADODB.Recordset
Private rctmp As New ADODB.Recordset
Public mvarcode As String
Private MyFormAddEditMode As EnumAddEditMode
Dim Parameter() As Parameter
Dim i As Integer
Dim OldTafsili As Long
Dim intPrizeTypeNo As Integer

Public Sub Delete()
    'Case
        
        
            ReDim Parameter(1) As Parameter
            Parameter(0) = GenerateInputParameter("@intPrizeTypeNo", adInteger, 4, intPrizeTypeNo)
            Parameter(1) = GenerateOutputParameter("@Deleted", adInteger, 4)
            
            Dim Deleted As Long
            Deleted = RunParametricStoredProcedure("Delete_tblTotal_PrizeType_ByPk_intPrizeTypeNo", Parameter)
            If Deleted <> False Then
                frmMsg.fwlblMsg.Caption = "Õ–› »« „Ê›ﬁÌ  «‰Ã«„ ‘œ"
                frmMsg.fwBtn(0).Visible = False
                frmMsg.fwBtn(1).ButtonType = flwButtonOk
                frmMsg.fwBtn(1).Caption = "ﬁ»Ê·"
                frmMsg.Show vbModal
                
            Else
                frmMsg.fwlblMsg.Caption = "Õ–› «‰Ã«„ ‰‘œ"
                frmMsg.fwBtn(0).Visible = False
                frmMsg.fwBtn(1).ButtonType = flwButtonOk
                frmMsg.fwBtn(1).Caption = "ﬁ»Ê·"
                frmMsg.Show vbModal
                txtPrizeType.SetFocus
                Exit Sub
            End If

        'End Select
        MyFormAddEditMode = AddMode
    SetFirstToolBar
    Add
    
    Exit Sub
RollBack:
End Sub

Private Sub FillvsPrizeType()

    Dim Rst As New ADODB.Recordset
    ReDim Parameter(0) As Parameter
    
    
    Parameter(0) = GenerateInputParameter("@intPrizeTypeNo", adInteger, 4, -1)
    Set Rst = RunParametricStoredProcedure2Rec("Get_tblTotal_PrizeType_ByPK_intPrizeTypeNo", Parameter)
    
    With vsPrizeType
        .Rows = 1
        i = 0
        While Rst.EOF <> True
            .Rows = .Rows + 1
            i = .Rows - 1
            .TextMatrix(i, 0) = i
            .TextMatrix(i, 1) = Rst!intPrizeTypeNo
            .TextMatrix(i, 2) = Rst!nvcPrizeTypeName
            Rst.MoveNext
        Wend
    End With
    Set Rst = Nothing
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

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()

    CenterTop Me
    
''    If ClsFormAccess.frmSupplier = False Then
''        Unload Me
''        Exit Sub
''    End If
    
    VarActForm = Me.Name
    
     vsPrizeType.ColHidden(1) = True
     Add

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Set Rc = Nothing
    Set rctmp = Nothing
    
    If PosConnection.State = adStateOpen Then PosConnection.Close: Set PosConnection = Nothing
    
    Set clsDate = Nothing
    Set mdifrm.FileCls = Nothing
        
    VarActForm = ""
    Dim obj As Object
    Dim Exit_Form As Boolean
   
''''
    
    If Exit_Form = False Then
        mdifrm.Toolbar1.Buttons(20).Enabled = False
        mdifrm.Toolbar1.Buttons(21).Enabled = False
        mdifrm.Toolbar1.Buttons(23).Enabled = True
        mdifrm.Toolbar1.Buttons(24).Enabled = True
        mdifrm.Toolbar1.Buttons(25).Enabled = True
        mdifrm.Toolbar1.Buttons(26).Enabled = True
        mdifrm.Toolbar1.Buttons(27).Enabled = True
    End If

End Sub


Public Sub Cancel()
    Select Case MyFormAddEditMode
        Case AddMode 'new
            DefaultSettings
            MyFormAddEditMode = AddMode
            SetFirstToolBar
        Case EditMode 'edit
            GetDataDetail
            MyFormAddEditMode = ViewMode
            SetFirstToolBar
    End Select
End Sub

Public Sub DefaultSettings()

    On Error Resume Next
    
    On Error GoTo 0
    txtPrizeType.Text = ""
End Sub

Public Sub Add()

    If MyFormAddEditMode = EditMode Then
        DefaultSettings
    End If
    MyFormAddEditMode = AddMode
    DefaultSettings
    SetFirstToolBar
    FillvsPrizeType
End Sub

Public Sub ExitSub()
If PosConnection.State = adStateOpen Then PosConnection.Close: Set PosConnection = Nothing
    Unload Me
End Sub

Public Sub Update()
    If MyFormAddEditMode = ViewMode Then Exit Sub
    Dim strBinBuyState As String
    Dim intBuyState As Integer
    If txtPrizeType.Text = "" Then
                frmMsg.fwlblMsg.Caption = "‰Ê⁄ Ã«Ì“Â —« Å— ﬂ‰Ìœ"
                frmMsg.fwBtn(0).Visible = False
                frmMsg.fwBtn(1).ButtonType = flwButtonOk
                frmMsg.fwBtn(1).Caption = "ﬁ»Ê·"
                frmMsg.Show vbModal
                Exit Sub
    End If
    Select Case MyFormAddEditMode
        Case AddMode
            ReDim Parameter(1) As Parameter
            Parameter(0) = GenerateInputParameter("@nvcPrizeTypeName", adVarWChar, 50, txtPrizeType.Text)
            Parameter(1) = GenerateOutputParameter("@intPrizeTypeNo", adInteger, 4)
            
            Dim LastCode As Long
            LastCode = RunParametricStoredProcedure("Insert_tblTotal_PrizeType", Parameter)
            If LastCode <> -1 Then
                frmMsg.fwlblMsg.Caption = "À»  ‰Ê⁄ Ã«Ì“Â ÃœÌœ »« „Ê›ﬁÌ  «‰Ã«„ ‘œ"
                frmMsg.fwBtn(0).Visible = False
                frmMsg.fwBtn(1).ButtonType = flwButtonOk
                frmMsg.fwBtn(1).Caption = "ﬁ»Ê·"
                frmMsg.Show vbModal
            Else
                frmMsg.fwlblMsg.Caption = "À»  «‰Ã«„ ‰‘œ"
                frmMsg.fwBtn(0).Visible = False
                frmMsg.fwBtn(1).ButtonType = flwButtonOk
                frmMsg.fwBtn(1).Caption = "ﬁ»Ê·"
                frmMsg.Show vbModal
                txtPrizeType.SetFocus
                Exit Sub
            End If
            
            
        Case EditMode
        
        
            ReDim Parameter(2) As Parameter
            Parameter(0) = GenerateInputParameter("@nvcPrizeTypeName", adVarWChar, 50, txtPrizeType.Text)
            Parameter(1) = GenerateInputParameter("@intPrizeTypeNo", adInteger, 4, intPrizeTypeNo)
            Parameter(2) = GenerateOutputParameter("@Updated", adInteger, 4)
            
            Dim Updated As Long
            Updated = RunParametricStoredProcedure("Update_tblTotal_PrizeType_ByPk_intPrizeTypeNo", Parameter)
            If Updated <> False Then
                frmMsg.fwlblMsg.Caption = " €ÌÌ—«  »« „Ê›ﬁÌ  «‰Ã«„ ‘œ"
                frmMsg.fwBtn(0).Visible = False
                frmMsg.fwBtn(1).ButtonType = flwButtonOk
                frmMsg.fwBtn(1).Caption = "ﬁ»Ê·"
                frmMsg.Show vbModal
                
            Else
                frmMsg.fwlblMsg.Caption = " €ÌÌ—«  «‰Ã«„ ‰‘œ"
                frmMsg.fwBtn(0).Visible = False
                frmMsg.fwBtn(1).ButtonType = flwButtonOk
                frmMsg.fwBtn(1).Caption = "ﬁ»Ê·"
                frmMsg.Show vbModal
                txtPrizeType.SetFocus
                Exit Sub
            End If

        End Select
    
    MyFormAddEditMode = AddMode
    SetFirstToolBar
    Add
    
    Exit Sub
RollBack:
        
End Sub


Public Sub Edit()
 
    MyFormAddEditMode = EditMode
    SetFirstToolBar
    
End Sub

Public Sub ExitForm()
    Unload Me
End Sub

Sub SetFirstToolBar()
    
    Dim obj As Object
    
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
 
        On Error Resume Next
        For Each obj In Me
           obj.Locked = True
        Next obj
        On Error GoTo 0
        mdifrm.Toolbar1.Buttons(10).Enabled = True
        mdifrm.Toolbar1.Buttons(6).Enabled = True  'Add
        mdifrm.Toolbar1.Buttons(7).Enabled = True  'Edit
        mdifrm.Toolbar1.Buttons(8).Enabled = False  'Enter
        mdifrm.Toolbar1.Buttons(9).Enabled = False   'Esc
        
    ElseIf MyFormAddEditMode = AddMode Then    'Add Mode
        
        On Error Resume Next
        For Each obj In Me
                obj.Locked = False
        Next obj
        On Error GoTo 0
        
        mdifrm.Toolbar1.Buttons(6).Enabled = False  'Add
        mdifrm.Toolbar1.Buttons(7).Enabled = False  'Edit
        mdifrm.Toolbar1.Buttons(8).Enabled = True   'Enter
        mdifrm.Toolbar1.Buttons(9).Enabled = True   'Esc
    
    ElseIf MyFormAddEditMode = EditMode Then     'Edit
    
        On Error Resume Next
        For Each obj In Me
                obj.Locked = False
        Next obj
        On Error GoTo 0
        
        mdifrm.Toolbar1.Buttons(6).Enabled = True  'Add
        mdifrm.Toolbar1.Buttons(7).Enabled = False  'Edit
        mdifrm.Toolbar1.Buttons(8).Enabled = True   'Enter
        mdifrm.Toolbar1.Buttons(9).Enabled = True   'Esc
        
    End If
    
    HeaderLabel Val(MyFormAddEditMode), fwlblMode
    
End Sub
Sub GetDataDetail()
    
    DefaultSettings
    
    Dim TempStr As String
    
    ReDim Parameter(0) As Parameter
    Parameter(0) = GenerateInputParameter("@intPrizeTypeNo", adInteger, 4, intPrizeTypeNo)
    Set rctmp = RunParametricStoredProcedure2Rec("Get_tblTotal_PrizeType_ByPK_intPrizeTypeNo", Parameter)
    Dim ii As Integer
    If Not (rctmp.BOF Or rctmp.EOF) Then
        
        txtPrizeType.Text = rctmp!nvcPrizeTypeName
               
    End If
    rctmp.Close
    
    
End Sub





Private Sub vsPrizeType_AfterSort(ByVal Col As Long, Order As Integer)
    With vsPrizeType
        If Col = 3 And .Rows > 1 Then
            For i = 1 To .Rows - 2
                If (Val(.TextMatrix(i, 3)) - Val(.TextMatrix(i + 1, 3)) > 1 And Order = 2) Or (Val(.TextMatrix(i + 1, 3)) - Val(.TextMatrix(i, 3)) > 1 And Order = 1) Then
                    .Cell(flexcpBackColor, i, .FixedCols, i, .Cols - 1) = 8421631
                Else
                    .Cell(flexcpBackColor, i, .FixedCols, i, .Cols - 1) = &H80000005
                End If
            Next i
        End If
    End With
End Sub

Private Sub vsPrizeType_Click()
    
    intPrizeTypeNo = vsPrizeType.TextMatrix(vsPrizeType.Row, 1)
    MyFormAddEditMode = ViewMode
    GetDataDetail
    SetFirstToolBar
    HeaderLabel Val(MyFormAddEditMode), Me.fwlblMode

End Sub




