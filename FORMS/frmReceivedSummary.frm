VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{7AEDC602-D94C-11D1-BB7A-00E0290EA3C9}#1.0#0"; "ResizeKit.ocx"
Object = "{75D4F767-8785-11D3-93AD-0000832EF44D}#2.14#0"; "FAST2001.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmReceivedSummary 
   Caption         =   "À»  „ÊÃÊœÌ ò«—»—«‰"
   ClientHeight    =   8100
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6360
   LinkTopic       =   "Form1"
   ScaleHeight     =   8100
   ScaleWidth      =   6360
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtTafsili 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   25
      Top             =   1200
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Frame frame_Change 
      Height          =   855
      Left            =   120
      TabIndex        =   19
      Top             =   7200
      Width           =   6135
      Begin VB.CommandButton cmdCreateAddDecrease 
         BackColor       =   &H00008000&
         Caption         =   " Ê·Ìœ ”‰œò”— Ê «÷«›Â ’‰œÊﬁ"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   4440
         RightToLeft     =   -1  'True
         TabIndex        =   24
         Top             =   120
         Width           =   1575
      End
      Begin VB.TextBox txtDecPrice 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   495
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   21
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox txtAddPrice 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2280
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "ò«Â‘: "
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   11.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   495
         Left            =   1560
         RightToLeft     =   -1  'True
         TabIndex        =   23
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "«›“«Ì‘: "
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   11.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3600
         RightToLeft     =   -1  'True
         TabIndex        =   22
         Top             =   240
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   120
      TabIndex        =   15
      Top             =   6240
      Width           =   6135
      Begin FLWCtrls.FWButton FWBtnOK 
         Height          =   615
         Left            =   4440
         TabIndex        =   16
         Top             =   240
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   1085
         Caption         =   "À» (F12)"
         FontName        =   "B Homa"
         FontSize        =   11.25
         Alignment       =   1
      End
      Begin FLWCtrls.FWButton FWBtnCancel 
         Height          =   615
         Left            =   120
         TabIndex        =   17
         ToolTipText     =   "«‰’—«› Ê Œ—ÊÃ"
         Top             =   240
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   1085
         ButtonType      =   1
         Caption         =   "«‰’—«›"
         BackColor       =   12632256
         FontName        =   "B Homa"
         FontSize        =   11.25
         Alignment       =   1
         Object.ToolTipText     =   "«‰’—«› Ê Œ—ÊÃ"
      End
      Begin FLWCtrls.FWButton FWBtnDelete 
         Height          =   615
         Left            =   2040
         TabIndex        =   18
         ToolTipText     =   "«‰’—«› Ê Œ—ÊÃ"
         Top             =   240
         Width           =   1875
         _ExtentX        =   3307
         _ExtentY        =   1085
         ButtonType      =   6
         Caption         =   "Õ–› ”ÿ—(F8)"
         BackColor       =   12632256
         FontName        =   "B Homa"
         FontSize        =   11.25
         Alignment       =   1
         Object.ToolTipText     =   "«‰’—«› Ê Œ—ÊÃ"
      End
   End
   Begin VB.ComboBox cmbPerson 
      BackColor       =   &H00FFFFFF&
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
      ItemData        =   "frmReceivedSummary.frx":0000
      Left            =   2280
      List            =   "frmReceivedSummary.frx":0002
      RightToLeft     =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   13
      Top             =   1440
      Width           =   2775
   End
   Begin VB.TextBox txtSanadNo 
      Alignment       =   1  'Right Justify
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   120
      Width           =   975
   End
   Begin VB.ComboBox cmbShift 
      BackColor       =   &H00FFFFFF&
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
      ItemData        =   "frmReceivedSummary.frx":0004
      Left            =   2280
      List            =   "frmReceivedSummary.frx":0006
      RightToLeft     =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   840
      Width           =   2775
   End
   Begin VB.TextBox txtDescription 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      MultiLine       =   -1  'True
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   4800
      Width           =   6135
   End
   Begin RESIZEKITLibCtl.ResizeKit ResizeKit1 
      Height          =   480
      Left            =   240
      OleObjectBlob   =   "frmReceivedSummary.frx":0008
      TabIndex        =   0
      Top             =   0
      Width           =   480
   End
   Begin MSMask.MaskEdBox txtDate 
      Height          =   585
      Left            =   3600
      TabIndex        =   5
      Top             =   120
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
   Begin VSFlex7LCtl.VSFlexGrid vsReceived 
      Height          =   2715
      Left            =   120
      TabIndex        =   14
      Top             =   1920
      Width           =   6195
      _cx             =   10927
      _cy             =   4789
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
      Rows            =   2
      Cols            =   2
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
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
   Begin VB.Label Label21 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "‘„«—Â : "
      BeginProperty Font 
         Name            =   "B Homa"
         Size            =   11.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1080
      RightToLeft     =   -1  'True
      TabIndex        =   12
      Top             =   180
      Width           =   735
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "‘Ì›  "
      BeginProperty Font 
         Name            =   "B Homa"
         Size            =   11.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5520
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   840
      Width           =   735
   End
   Begin VB.Label lblDay 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "B Homa"
         Size            =   11.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   420
      Left            =   2880
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   360
      Width           =   60
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "  «—ÌŒ ⁄„·ò—œ:"
      BeginProperty Font 
         Name            =   "B Homa"
         Size            =   11.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4920
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label LblTotal 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "B Nazanin"
         Size            =   14.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   240
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   5760
      Width           =   2145
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Ã„⁄ : "
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
      Left            =   2520
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   5760
      Width           =   585
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ﬂ«—»— :  "
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
      Left            =   5520
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   1320
      Width           =   585
   End
   Begin VB.Label LblUserName 
      Alignment       =   1  'Right Justify
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
      Height          =   465
      Left            =   9240
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   360
      Width           =   2505
   End
End
Attribute VB_Name = "frmReceivedSummary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Rst As New ADODB.Recordset
Dim p() As Parameter
Dim Flag As Boolean
Public intSerialNo As Long
Dim Updated As Long
Dim clsDate As New clsDate
Public AccessUser As Boolean

Private Sub cmbPerson_Click()
    If formloadFlag = False Then Exit Sub
    FillGrid
    Set Rst = RunStoredProcedure2RecordSet("Get_User")
    
    If Not (Rst.EOF = True And Rst.BOF = True) Then
        Do While Rst.EOF <> True
            If cmbPerson.ItemData(cmbPerson.ListIndex) = Rst!Uid Then
                txtTafsili = Rst!Tafsili
                Exit Do
            End If
        Loop
    End If
    Rst.Close
End Sub

Private Sub cmbShift_Click()
    If formloadFlag = False Then Exit Sub
    FillGrid
End Sub

Private Sub cmbShift_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        VSReceived.SetFocus: Sendkey "{F4}", True: VSReceived.Select VSReceived.Row, VSReceived.Col
    End If
End Sub

Private Sub cmdCreateAddDecrease_Click()
Dim varForm As Form
Dim frmact As Form

For Each varForm In Forms
    If varForm.Name = "frmCreateSanad" Then
        Set frmact = varForm
        Exit For
    End If
Next
If frmact Is Nothing Then Exit Sub
    If frmact.FillAddDecrease = False Then
        ShowDisMessage "”‰œ ò”— Ê «÷«›Â »—«Ì «Ì‰ ò«—»— ﬁ»·« Ê«—œ ‘œÂ", 1300
    End If
End Sub

Private Sub Form_Activate()

    txtDate.SetFocus
    If AccessUser = False Then cmbPerson.Enabled = False Else cmbPerson.Enabled = True

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case Shift
          Case 0
              Select Case KeyCode
                  Case 27  ' Esc
                     FWBtnCancel_Click
                  Case vbKeyF12  ' Esc
                        FWBtnOK_Click
              End Select
          Case 2
               Select Case KeyCode
                  Case 123  'Exit
                    FWBtnCancel_Click
              End Select
    
    End Select
End Sub

Private Sub Form_Load()
    
Dim ii As Long
    formloadFlag = False
    
    SetGrid

    AddRowInGrid
    
    Set Rst = RunStoredProcedure2RecordSet("Get_User")
    
    cmbPerson.Clear
    If Not (Rst.EOF = True And Rst.BOF = True) Then
        Do While Rst.EOF <> True
            cmbPerson.AddItem CStr(Rst.Fields("PersonName"))
            cmbPerson.ItemData(cmbPerson.ListCount - 1) = Val(Rst.Fields("Uid"))
            Rst.MoveNext
    
        Loop
        Rst.Close
        For ii = 0 To cmbPerson.ListCount - 1
            If cmbPerson.ItemData(ii) = mvarCurUserNo Then
                cmbPerson.ListIndex = ii
                Exit For
            End If
        Next
    End If

    cmbShift.Clear
    Set Rst = RunStoredProcedure2RecordSet("Get_All_tShift")
    If Not (Rst.EOF = True And Rst.BOF = True) Then
        While Not Rst.EOF
            cmbShift.AddItem Rst!Description
            cmbShift.ItemData(cmbShift.NewIndex) = Rst!Code
            Rst.MoveNext
        Wend
    End If
    If Rst.State = adStateOpen Then If Rst.State = adStateOpen Then Rst.Close
    
    LoadForm Me.Name
    formloadFlag = True
'    CenterCenterOffset Me
'    CenterTop Me

    txtDate = mvarDate
    cmdCreateAddDecrease.Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    If PosConnection.State = adStateOpen Then PosConnection.Close: Set PosConnection = Nothing
    SaveSetting strMainKey, Me.Name, "Left", Me.Left
    SaveSetting strMainKey, Me.Name, "Top", Me.Top

End Sub

Private Sub FWBtnCancel_Click()
    If Flag = True Then
        frmMsg.fwlblMsg.Caption = "¬Ì« »—«Ì Œ—ÊÃ «“ ›—„ «ÿ„Ì‰«‰ œ«—Ìœ"
        frmMsg.fwBtn(0).ButtonType = flwButtonOk
        frmMsg.fwBtn(0).Caption = "»·Â"
        frmMsg.fwBtn(1).ButtonType = flwButtonCancel
        frmMsg.fwBtn(1).Caption = "ŒÌ—"
        frmMsg.Show vbModal
        If mvarMsgIdx <> vbYes Then Exit Sub
    End If
    mvarIndexNo = 0
    Unload Me
End Sub

Private Sub FWBtnOK_Click()
    Update
End Sub

Private Sub Update()
On Error GoTo ErrHandler
If cmbPerson.ListIndex = -1 Then ShowDisMessage "ò«—»— «‰ Œ«» ‰‘œÂ", 1200: Exit Sub
If cmbShift.ListIndex = -1 Then ShowDisMessage "‘Ì›  ò«—Ì «‰ Œ«» ‰‘œÂ", 1200: Exit Sub
Dim i As Long
Dim Result As Integer
    
With VSReceived
    For i = 1 To .Rows - 1
        If Len(txtDate.ClipText) < 6 Or (Trim(.TextMatrix(i, 1)) <> "" And Trim(.TextMatrix(i, 2)) = "") Or (Trim(.TextMatrix(i, 1)) = 5 And Trim(.TextMatrix(i, 3)) = "") Then
            frmMsg.fwlblMsg.Caption = "·ÿ›« «ÿ·«⁄«  ÷—Ê—Ì —« Ê«—œ ‰„«ÌÌœ"
            frmMsg.fwBtn(0).ButtonType = flwButtonOk
            frmMsg.fwBtn(0).Caption = "ﬁ»Ê·"
            frmMsg.fwBtn(1).Visible = False
            frmMsg.Show vbModal
            Exit Sub
        End If
    Next i
    
    ReDim Parameter(0) As Parameter
    Parameter(0) = GenerateInputParameter("@SanadNo", adInteger, 4, Val(txtSanadNo))

    RunParametricStoredProcedure "Delete_tblAcc_ReceivedSummary", Parameter
    
    ReDim Parameter(9) As Parameter
    For i = 1 To .Rows - 1
        If Trim(.TextMatrix(i, 1)) <> "" Then
            Parameter(0) = GenerateInputParameter("@nvcDate", adWChar, 8, Trim(txtDate.Text))
            Parameter(1) = GenerateInputParameter("@ShiftNo", adInteger, 4, cmbShift.ItemData(cmbShift.ListIndex)) ' .TextMatrix(i, 3)
            Parameter(2) = GenerateInputParameter("@Uid", adInteger, 4, cmbPerson.ItemData(cmbPerson.ListIndex)) ' .TextMatrix(i, 3)
            Parameter(3) = GenerateInputParameter("@intRow", adInteger, 4, i)
            Parameter(4) = GenerateInputParameter("@ReceivedType", adTinyInt, 1, .TextMatrix(i, 1))
            Parameter(5) = GenerateInputParameter("@Price", adBigInt, 8, .TextMatrix(i, 2))
            Parameter(6) = GenerateInputParameter("@PosId", adInteger, 4, Val(.TextMatrix(i, 3)))
            Parameter(7) = GenerateInputParameter("@nvcDescription", adVarChar, 255, Right(txtDescription, 255))
            Parameter(8) = GenerateInputParameter("@SanadNo", adInteger, 4, Val(txtSanadNo))
            Parameter(9) = GenerateOutputParameter("@Result", adInteger, 4)
            
            Result = RunParametricStoredProcedure("Insert_tblAcc_ReceivedSummary", Parameter)
        End If
    Next i
            
            
    frmMsg.fwlblMsg.Caption = "À»  «ÿ·«⁄«  ÃœÌœ »« „Ê›ﬁÌ  Å«Ì«‰ Ì«› "
    frmMsg.fwBtn(0).ButtonType = flwButtonOk
    frmMsg.fwBtn(0).Caption = "ﬁ»Ê·"
    frmMsg.fwBtn(1).Visible = False
    frmMsg.Show vbModal
    VSReceived.Rows = 1
    FillGrid
            
 End With


Exit Sub
ErrHandler:
    ShowDisMessage "Œÿ«. »⁄÷Ì «“ «ÿ·«⁄«   ò—«—Ì „Ì »«‘œ", 1500
    FillGrid
End Sub
Private Sub MaxSanadNo()
    Dim Rst As New ADODB.Recordset
    Set Rst = RunStoredProcedure2RecordSet("Get_Max_tblAcc_ReceivedSummary")
    
    If Not (Rst.BOF = True And Rst.EOF = True) Then
        txtSanadNo = Rst!SanadNo
    End If
    
    If Rst.State = 1 Then Rst.Close
    Set Rst = Nothing

End Sub
Private Function DoCalculate() As String
    Dim i As Integer
    Dim s As Double
    s = 0#
    For i = 1 To VSReceived.Rows - 1
        If Val(VSReceived.TextMatrix(i, 1)) > 0 Then
            s = s + Val(VSReceived.TextMatrix(i, 2))
        End If
    Next i
    DoCalculate = CStr(s)
    LblTotal = DoCalculate
End Function

Private Sub FWBtnDelete_Click()
    With VSReceived
        If .Row >= 1 And .Rows > 1 Then
            .RemoveItem .Row
            DoCalculate
            vsReceived_AfterSort 1, 0
        End If
    End With
End Sub

Private Sub LblTotal_Change()
    If Val(LblTotal) > 0 Then LblTotal = Format(LblTotal, "###,###") & clsArya.UnitPrice Else LblTotal = 0
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

Private Sub txtDate_Change()
    lblDay = ""
    If Len(txtDate.ClipText) = 6 Then lblDay = clsDate.Find_DayOfWeekShamsi("13" & txtDate.Text): FillGrid
    
End Sub

Private Sub txtDate_GotFocus()
    txtDate.SelStart = 3
    txtDate.SelLength = 5
End Sub

Private Sub txtDate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then cmbShift.SetFocus

End Sub

Private Sub txtDate_LostFocus()
    If CheckDate6Digit(txtDate.Text) = False Then
        txtDate.SetFocus
    End If
End Sub
Private Sub cmbShift_GotFocus()

    Sendkey "{F4}", True
End Sub

Private Sub vsReceived_AfterSort(ByVal Col As Long, Order As Integer)
    Dim i As Integer

    For i = 1 To VSReceived.Rows - 1
        VSReceived.TextMatrix(i, 0) = CStr(i)
    Next i
End Sub

Private Sub VSReceived_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    Dim i As Long
    For i = 0 To VSReceived.Cols - 1
        SaveSetting strMainKey, Me.Name, "Col" & Col, VSReceived.ColWidth(Col)
    Next

End Sub

Private Sub vsReceived_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Select Case CInt(Val(VSReceived.TextMatrix(Row, 1)))
        Case 1
            If Col = 3 Then Cancel = True
    End Select
End Sub

Private Sub SetGrid()
    With VSReceived
        .Rows = 2
        .Cols = 4
        .ColHidden(-1) = False
        .TextMatrix(0, 0) = "—œÌ›"
        .TextMatrix(0, 1) = "‰Ê⁄ œ—Ì«› "
        .TextMatrix(0, 2) = "„»·€"
        .TextMatrix(0, 3) = "ÅÊ“ »«‰òÌ"
        
        .FixedAlignment(-1) = flexAlignCenterCenter
        .ColAlignment(-1) = flexAlignCenterCenter

        .ColFormat(2) = "###,###"
        Set Rst = RunStoredProcedure2RecordSet("Get_All_tRecvType")
        .ColComboList(1) = .BuildComboList(Rst, "nvcDescription", "tintType")
        Rst.Close

        Set Rst = RunStoredProcedure2RecordSet("Get_All_tblPub_Pos")
        .ColComboList(3) = .BuildComboList(Rst, "nvcBankName", "PosId")
        Rst.Close
        Dim i As Long
        For i = 0 To .Cols - 1
            .ColWidth(i) = Val(GetSetting(strMainKey, Me.Name, "Col" & i))
            If .ColWidth(i) = 0 Then
                .ColWidth(i) = .Width / 3
            End If
        Next i
    End With
End Sub
Private Sub FillGrid()
cmdCreateAddDecrease.Enabled = False
With VSReceived
    If cmbShift.ListIndex = -1 Then Exit Sub
'    txtDate = " /  /  "
    .Rows = 1
    ReDim Parameter(2) As Parameter
    Parameter(0) = GenerateInputParameter("@nvcDate", adWChar, 8, Trim(txtDate.Text))
    Parameter(1) = GenerateInputParameter("@ShiftNo", adInteger, 4, cmbShift.ItemData(cmbShift.ListIndex)) ' .TextMatrix(i, 3)
    Parameter(2) = GenerateInputParameter("@Uid", adInteger, 4, cmbPerson.ItemData(cmbPerson.ListIndex)) ' .TextMatrix(i, 3)
    
    Set Rst = RunParametricStoredProcedure2Rec("Get_tblAcc_ReceivedSummary", Parameter)
    If Not (Rst.BOF = True And Rst.EOF = True) Then
        txtSanadNo = Rst!SanadNo
        While Rst.EOF <> True
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, 0) = .Rows - 1
            .TextMatrix(.Rows - 1, 1) = Rst!ReceivedType
            .TextMatrix(.Rows - 1, 2) = Rst!Price
            .TextMatrix(.Rows - 1, 3) = IIf(IsNull(Rst!PosId), "", Rst!PosId)
            Rst.MoveNext
        Wend
        DoCalculate
        If frame_Change.Visible = True Then ChangeCalculate
    Else
        .Rows = 2
        MaxSanadNo
'        cmbShift_KeyDown vbKeyReturn, 0
    End If
    
    If Rst.State = 1 Then Rst.Close
    Set Rst = Nothing

End With
End Sub
Private Sub ChangeCalculate()

Dim varForm As Form
Dim frmact As Form

For Each varForm In Forms
    If varForm.Name = "frmCreateSanad" Then
        Set frmact = varForm
        Exit For
    End If
Next
If frmact Is Nothing Then Exit Sub

Dim Rst As New ADODB.Recordset
Dim i As Long

With VSReceived
    
    ReDim Parameter(4) As Parameter

    Parameter(0) = GenerateInputParameter("@Branch", adInteger, 4, frmact.cboBranch.ItemData(frmact.cboBranch.ListIndex))
    Parameter(1) = GenerateInputParameter("@DateBefore", adVarWChar, 8, frmact.txtDate1.Text)
    Parameter(2) = GenerateInputParameter("@DateAfter", adVarWChar, 8, frmact.txtDate2.Text)
    Parameter(3) = GenerateInputParameter("@Code", adInteger, 4, EnumAccountType.Cash)
    Parameter(4) = GenerateInputParameter("@Uid", adInteger, 4, cmbPerson.ItemData(cmbPerson.ListIndex))
    Set Rst = RunParametricStoredProcedure2Rec("Get_AccountDocument", Parameter)

    If Rst.EOF <> True And Rst.BOF <> True Then
        For i = 1 To .Rows - 1
            If .ValueMatrix(i, 1) = 1 Then
                If .ValueMatrix(i, 2) - Rst.Fields("sp").Value > 0 Then
                    txtAddPrice = .ValueMatrix(i, 2) - Rst.Fields("sp").Value
                    txtDecPrice = ""
                    cmdCreateAddDecrease.Enabled = True
                ElseIf .ValueMatrix(i, 2) - Rst.Fields("sp").Value < 0 Then
                    txtDecPrice = Rst.Fields("sp").Value - .ValueMatrix(i, 2)
                    txtAddPrice = ""
                    cmdCreateAddDecrease.Enabled = True
                Else
                    txtAddPrice = ""
                    txtDecPrice = ""
                End If
                Exit For
            End If
        Next
    End If
    Rst.Close

End With
Set Rst = Nothing
End Sub
Private Sub AddRowInGrid()
    Dim flgAddRow As Boolean
    Dim s As String
    Dim c As Integer
    DoCalculate
    flgAddRow = False
    If VSReceived.Rows > VSReceived.FixedRows Then
        s = ""
        For c = 1 To VSReceived.Cols - 1
            s = s + VSReceived.TextMatrix(VSReceived.Rows - 1, c)
        Next c
        If Len(s) > 0 Then flgAddRow = True
    Else
        flgAddRow = True
    End If
    If flgAddRow = True Then
        VSReceived.Rows = VSReceived.Rows + 1
        VSReceived.Row = VSReceived.Rows - 1
        VSReceived.Col = 1
    End If
End Sub

Private Sub vsReceived_Click()
With VSReceived
    .Select .Row, .Col: .EditCell
    If .Col = 1 Then Sendkey "{F4}", False
End With

End Sub

Private Sub vsReceived_EnterCell()
With VSReceived
    .Select .Row, .Col: .EditCell
    If .Col = 1 Then Sendkey "{F4}", False
End With
End Sub

Private Sub vsReceived_KeyDown(KeyCode As Integer, Shift As Integer)
With VSReceived
    If KeyCode = vbKeyReturn Then
        Select Case CInt(Val(VSReceived.TextMatrix(.Row, 1)))
            Case 1
                If .Col = 1 Then
                    .Col = 2
                ElseIf .Col = 2 Then
                    AddRowInGrid
                End If
            Case 5
                If .Col = 1 Then
                    .Col = 2
                ElseIf .Col = 2 Then
                    .Col = 3
                    Sendkey "{F4}", False
                ElseIf .Col = 3 Then
                    AddRowInGrid
                End If
        End Select
        DoCalculate
    ElseIf KeyCode = vbKeyF12 Then
        FWBtnOK_Click
    End If
End With
End Sub

Private Sub vsReceived_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    Select Case Col
        Case 2
            If (KeyAscii < vbKey0 Or KeyAscii > vbKey9) And KeyAscii <> vbKeyBack And KeyAscii <> vbKeyReturn And KeyAscii <> 45 Then KeyAscii = 0
    End Select
End Sub

