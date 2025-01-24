VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{7AEDC602-D94C-11D1-BB7A-00E0290EA3C9}#1.0#0"; "ResizeKit.ocx"
Begin VB.Form frmSmsNotSend 
   Caption         =   "                                                        ·Ì”  «” «„ «” Â«Ì À»  ‘œÂ"
   ClientHeight    =   6210
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10335
   Icon            =   "frmSmsNotSend.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   6210
   ScaleWidth      =   10335
   Begin VB.CheckBox ChkSendSms 
      Alignment       =   1  'Right Justify
      Caption         =   "              ›ﬁÿ «—”«· ‰‘œÂ Â« "
      BeginProperty Font 
         Name            =   "Nazanin"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7200
      TabIndex        =   3
      Top             =   5640
      Width           =   2775
   End
   Begin VB.CommandButton CmdView 
      Caption         =   "‰„«Ì‘"
      BeginProperty Font 
         Name            =   "Nazanin"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2400
      TabIndex        =   2
      Top             =   5520
      Width           =   1575
   End
   Begin VB.CommandButton CmdSend 
      Caption         =   " «ÌÌœ ÃÂ  «—”«· „Ãœœ"
      BeginProperty Font 
         Name            =   "Nazanin"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   1
      Top             =   5520
      Width           =   2055
   End
   Begin RESIZEKITLibCtl.ResizeKit ResizeKit1 
      Height          =   480
      Left            =   0
      OleObjectBlob   =   "frmSmsNotSend.frx":A4C2
      TabIndex        =   0
      Top             =   0
      Width           =   480
   End
   Begin VSFlex7LCtl.VSFlexGrid vsSmsList 
      Height          =   5055
      Left            =   240
      TabIndex        =   4
      Top             =   360
      Width           =   9855
      _cx             =   17383
      _cy             =   8916
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
      AllowUserResizing=   3
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   5
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmSmsNotSend.frx":A548
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
End
Attribute VB_Name = "frmSmsNotSend"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Dim i As Integer
'Private MyFormAddEditMode As EnumAddEditMode
'Dim Parameter() As Parameter
'
'Public Sub FillvsSmsList()
'
'    Dim Flag As Integer
'
'    Flag = 0
'    If ChkSendSms.Value = 1 Then
'        Flag = 1
'    End If
'
'    Dim Rst As New ADODB.Recordset
'    ReDim Parameter(0) As Parameter
'
'    Parameter(0) = GenerateInputParameter("@Flag", adInteger, 4, Flag)
'    Set Rst = RunParametricStoredProcedure2Rec("Get_tblPub_Sms", Parameter)
'
'    With vsSmsList
'        .Rows = 1
'        If Not (Rst.BOF = True And Rst.EOF = True) Then
'
'            While Rst.EOF <> True
'                .Rows = .Rows + 1
'                .Row = .Rows - 1
'                .TextMatrix(.Row, 0) = .Row
'                .TextMatrix(.Row, 1) = Rst.Fields("Tel").Value
'                .TextMatrix(.Row, 2) = Rst.Fields("Description").Value 'IIf(IsNull(Rst.Fields("Description").Value), "", Rst.Fields("Description").Value)
'                .TextMatrix(.Row, 3) = Rst.Fields("SmsStatus").Value
'                .TextMatrix(.Row, 4) = Rst.Fields("id").Value
'                Rst.MoveNext
'            Wend
'            .AutoSizeMode = flexAutoSizeColWidth
'            .AutoSize 0, .Cols - 1
'        End If
'        If .Rows > 1 Then
'            .Cell(flexcpText, 1, 2, .Rows - 1, 2) = 0
'        End If
'    End With
'
'
'End Sub
'
'
'Private Sub ChkSendSms_Click()
'    FillvsSmsList
'End Sub
'
'Private Sub CmdSend_Click()
'    If vsSmsList.Row > 0 Then
'        If frmSms.txtSMSMessage.Text <> "" Then
'           frmSms.txtSMSMessage.Text = frmSms.txtSMSMessage.Text & vsSmsList.TextMatrix(vsSmsList.Row, 3)
'        Else
'            frmSms.txtSMSMessage.Text = vsSmsList.TextMatrix(vsSmsList.Row, 3)
'        End If
'        frmSms.LblEdit.Caption = vsSmsList.TextMatrix(vsSmsList.Row, 4)
'    End If
'End Sub
'
'Private Sub cmdView_Click()
'    FillvsSmsList
'End Sub
'Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    Call PresetScreenSaver
'End Sub
'
'Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'    Call PresetScreenSaver
'End Sub
'
'Private Sub Form_Load()
'
''    With vsSmsList
''        .Cols = 4
''        .Rows = 1
''        .Row = 0
''        .TextMatrix(.Row, 0) = "—œÌ›"
''        .TextMatrix(.Row, 1) = "‘„«—Â"
''        .TextMatrix(.Row, 2) = "„ ‰"
''        .TextMatrix(.Row, 3) = "Ê÷⁄Ì "
''        .ColHidden(1) = True
''        .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter
''        .ColAlignment(3) = flexAlignRightCenter
''        .ColAlignment(0) = flexAlignCenterCenter
''        .AutoSizeMode = flexAutoSizeColWidth
''        .AutoSize 0, .Cols - 1
''    End With
'
'    FillvsSmsList
'
'End Sub
'
'Private Sub Form_Unload(Cancel As Integer)
'    If PosConnection.State = adStateOpen Then PosConnection.Close: Set PosConnection = Nothing
''''
'End Sub
