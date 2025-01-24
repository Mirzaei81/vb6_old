VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{7AEDC602-D94C-11D1-BB7A-00E0290EA3C9}#1.0#0"; "ResizeKit.ocx"
Begin VB.Form frmMojodiReduce 
   BackColor       =   &H80000016&
   ClientHeight    =   8340
   ClientLeft      =   2775
   ClientTop       =   3870
   ClientWidth     =   7905
   Icon            =   "frmMojodiReduce.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   8340
   ScaleWidth      =   7905
   Begin VSFlex7LCtl.VSFlexGrid vsGood 
      Height          =   5445
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   7635
      _cx             =   13467
      _cy             =   9604
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
      AllowUserResizing=   0
      SelectionMode   =   3
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   6
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   500
      RowHeightMax    =   500
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmMojodiReduce.frx":A4C2
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
   Begin VB.CommandButton CancelButton 
      BackColor       =   &H000000C0&
      Cancel          =   -1  'True
      Caption         =   "ÎíÑ"
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
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   7200
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      BackColor       =   &H0000C000&
      Caption         =   "Èáí"
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
      Left            =   1560
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   7200
      Width           =   1215
   End
   Begin RESIZEKITLibCtl.ResizeKit ResizeKit1 
      Height          =   480
      Left            =   0
      OleObjectBlob   =   "frmMojodiReduce.frx":A578
      TabIndex        =   7
      Top             =   0
      Width           =   480
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ÏÑ ÕæÑÊ ßÓÑí ÔãÇ ãÌÇÒ Èå ÝÑæÔ äíÓÊíÏ "
      BeginProperty Font 
         Name            =   "Nazanin"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   435
      Left            =   1920
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   7920
      Visible         =   0   'False
      Width           =   3825
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ÂíÇ Èå ËÈÊ ÝíÔ ÇÏÇãå ãí ÏåíÏ"
      BeginProperty Font 
         Name            =   "Nazanin"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   435
      Left            =   4080
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   6840
      Width           =   3225
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ßäÊÑá ãæÌæÏí æ ÈÇÞíãÇäÏå ßÇáÇ"
      BeginProperty Font 
         Name            =   "B Homa"
         Size            =   12
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2520
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   0
      Width           =   2895
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0E0FF&
      Height          =   495
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   0
      Width           =   7935
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "ÇÞáÇã ÏÇÑÇí ßÓÑ ãæÌæÏí"
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
      Height          =   435
      Left            =   4800
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   600
      Width           =   2865
   End
End
Attribute VB_Name = "frmMojodiReduce"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Result As Boolean
Dim i As Integer

Private Sub Form_Activate()
    If OKButton.Enabled = True Then OKButton.SetFocus
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call PresetScreenSaver
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call PresetScreenSaver
End Sub

Private Sub Form_Load()

    If ClsFormAccess.MojodiControl = False Then
        OKButton.Enabled = False
        Label3.Visible = True
    End If
    CenterCenterinSecondScreen Me
    
    Result = False
    
    FillvsGood

    formloadFlag = False
    Me.left = Val(GetSetting(strMainKey, Me.Name, "Left"))
    If Val(GetSetting(strMainKey, Me.Name, "Height")) > 5000 Then
        Me.Height = Val(GetSetting(strMainKey, Me.Name, "Height"))
    End If
    If Val(GetSetting(strMainKey, Me.Name, "Width")) > 5000 Then
        Me.Width = Val(GetSetting(strMainKey, Me.Name, "Width"))
    End If
    Me.top = Val(GetSetting(strMainKey, Me.Name, "Top"))
    If Me.left < 0 Then Me.left = 0
    If Me.top < 0 Then Me.top = 0
    If Me.top > Me.ScaleHeight Then Me.top = 0

    formloadFlag = True



End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveSetting strMainKey, Me.Name, "Left", Me.left
    SaveSetting strMainKey, Me.Name, "Top", Me.top


    If PosConnection.State = adStateOpen Then PosConnection.Close: Set PosConnection = Nothing
End Sub

Private Sub OKButton_Click()
    Result = True
    Unload Me

End Sub
Private Sub CancelButton_Click()
    Result = False
    Unload Me
End Sub

Private Sub FillvsGood()
Dim Rst As New ADODB.Recordset

If mvarAddeditMode = AddMode Then
    ReDim Parameter(3) As Parameter
    Parameter(0) = GenerateInputParameter("@DetailsString", adVarWChar, 4000, DetailsString1)
     Parameter(1) = GenerateInputParameter("@DetailsString2", adVarWChar, 4000, DetailsString2)
     Parameter(2) = GenerateInputParameter("@DetailsString3", adVarWChar, 4000, DetailsString3)
     Parameter(3) = GenerateInputParameter("@DetailsString4", adVarWChar, 4000, DetailsString4)
    Set Rst = RunParametricStoredProcedure2Rec("CheckPreSave", Parameter)
    If Not (Rst.BOF Or Rst.EOF) Then
        With vsGood
            .Rows = 1
            i = 0
            While Rst.EOF <> True
                .Rows = .Rows + 1
                i = i + 1
                .TextMatrix(i, 0) = i
                .TextMatrix(i, 1) = Rst!GoodCode
                .TextMatrix(i, 2) = Rst!GoodName
                .TextMatrix(i, 3) = Rst!Decrease
                .TextMatrix(i, 4) = Rst!Description
                .TextMatrix(i, 5) = Rst!MojodiControl
                Rst.MoveNext
            Wend
        End With
    End If
    Set Rst = Nothing
    vsGood.Cell(flexcpAlignment, 0, 0, 0, 3) = flexAlignCenterCenter
    vsGood.ColAlignment(-1) = flexAlignCenterCenter

ElseIf mvarAddeditMode = EditMode Or mvarAddeditMode = ManipulateMode Then
    ReDim Parameter(5) As Parameter
    Parameter(0) = GenerateInputParameter("@DetailsString", adVarWChar, 4000, DetailsString1)
    Parameter(1) = GenerateInputParameter("@Status", adInteger, 4, mvarStatus)
    Parameter(2) = GenerateInputParameter("@No", adBigInt, 8, mvarNo)
     Parameter(3) = GenerateInputParameter("@DetailsString2", adVarWChar, 4000, DetailsString2)
     Parameter(4) = GenerateInputParameter("@DetailsString3", adVarWChar, 4000, DetailsString3)
     Parameter(5) = GenerateInputParameter("@DetailsString4", adVarWChar, 4000, DetailsString4)
    Set Rst = RunParametricStoredProcedure2Rec("CheckPreSave_Edit", Parameter)
    If Not (Rst.BOF Or Rst.EOF) Then
        With vsGood
            .Rows = 1
            i = 0
            While Rst.EOF <> True
                .Rows = .Rows + 1
                i = i + 1
                .TextMatrix(i, 0) = i
                .TextMatrix(i, 1) = Rst!FirstCode
                .TextMatrix(i, 2) = Rst!GoodName
                .TextMatrix(i, 3) = Val(Rst!Remain) * -1
                .TextMatrix(i, 4) = Rst!Description
                .TextMatrix(i, 5) = Rst!MojodiControl
                Rst.MoveNext
            Wend
        End With
    End If
    Set Rst = Nothing
    vsGood.Cell(flexcpAlignment, 0, 0, 0, 3) = flexAlignCenterCenter
    vsGood.ColAlignment(-1) = flexAlignCenterCenter
End If

End Sub


Private Sub ResizeKit1_ExitResize(ByVal XScale As Double, ByVal YScale As Double)


    If formloadFlag = True Then
        SaveSetting strMainKey, Me.Name, "Height", Me.Height
        SaveSetting strMainKey, Me.Name, "Width", Me.Width
    End If

End Sub

Private Sub vsGood_AfterSort(ByVal Col As Long, Order As Integer)
    For i = 1 To vsGood.Rows - 1
        vsGood.TextMatrix(i, 0) = i
    Next
    
End Sub

Private Sub vsGood_DblClick()
    If ClsFormAccess.MojodiControl = False Then
            OKButton.Enabled = False
            Label3.Visible = True
            Exit Sub
    End If
    If vsGood.Row > 0 Then
        OKButton_Click
    End If
End Sub
