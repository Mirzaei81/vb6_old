VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{7AEDC602-D94C-11D1-BB7A-00E0290EA3C9}#1.0#0"; "ResizeKit.ocx"
Begin VB.Form frmSelectInventory 
   BackColor       =   &H80000016&
   ClientHeight    =   4470
   ClientLeft      =   2775
   ClientTop       =   3870
   ClientWidth     =   6240
   Icon            =   "frmSelectInventory.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   4470
   ScaleWidth      =   6240
   Begin VSFlex7LCtl.VSFlexGrid vsInventory 
      Height          =   2715
      Left            =   240
      TabIndex        =   0
      Top             =   720
      Width           =   5835
      _cx             =   10292
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
      AllowUserResizing=   3
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
      FormatString    =   $"frmSelectInventory.frx":A4C2
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
      Caption         =   "ÇäÕÑÇÝ"
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
      Left            =   240
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3720
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      BackColor       =   &H0000C000&
      Caption         =   "ÞÈæá"
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
      Height          =   495
      Left            =   1680
      RightToLeft     =   -1  'True
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3720
      Width           =   1215
   End
   Begin RESIZEKITLibCtl.ResizeKit ResizeKit1 
      Height          =   480
      Left            =   0
      OleObjectBlob   =   "frmSelectInventory.frx":A55E
      TabIndex        =   5
      Top             =   0
      Width           =   480
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ÇäÊÎÇÈ ÇäÈÇÑ ÈÑÇí ßÇáÇ"
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
      Left            =   1560
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   0
      Width           =   2895
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0E0FF&
      Caption         =   "      "
      Height          =   495
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   0
      Width           =   7935
   End
End
Attribute VB_Name = "frmSelectInventory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Result As Boolean
Dim i As Integer

Private Sub Form_Activate()
'''    If strCategory = "24" And strDelegate = "00" And (clsArya.CustomerId = 4 Or clsArya.CustomerId = 5 Or clsArya.CustomerId = 6 Or clsArya.CustomerId = 11) Then
'''
'''        CancelButton.Caption = "ÞÈæá"
'''        CancelButton.SetFocus
'''        OKButton.Visible = False
'''    Else
'''        OKButton.SetFocus
'''    End If
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call PresetScreenSaver
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call PresetScreenSaver
End Sub
Private Sub Form_Load()

    CenterCenterinSecondScreen Me
    
    Result = False
    
    FillvsInventory

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
    SaveSetting strMainKey, Me.Name, "Left", Me.Left
    SaveSetting strMainKey, Me.Name, "Top", Me.Top


    If PosConnection.State = adStateOpen Then PosConnection.Close: Set PosConnection = Nothing
End Sub

Private Sub OKButton_Click()
   If vsInventory.Row > 0 Then
        mvarInventoryNo = vsInventory.TextMatrix(vsInventory.Row, 3)
        mvarMojodi = vsInventory.TextMatrix(vsInventory.Row, 4)
   End If
    Unload Me

End Sub
Private Sub CancelButton_Click()
    mvarInventoryNo = 0
    Unload Me
End Sub

Private Sub FillvsInventory()
Dim Rst As New ADODB.Recordset

    ReDim Parameter(4) As Parameter
    Parameter(0) = GenerateInputParameter("@Code", adInteger, 4, mvarGoodCode)
    Parameter(1) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
    Parameter(2) = GenerateInputParameter("@StationId", adInteger, 4, clsArya.StationNo)
    Parameter(3) = GenerateInputParameter("@Flag", adBoolean, 1, 1)
    Parameter(4) = GenerateInputParameter("@AccountYear", adSmallInt, 2, AccountYear)
    Set Rst = RunParametricStoredProcedure2Rec("Get_Good_Code", Parameter)
    
    If Not (Rst.BOF Or Rst.EOF) Then
        With vsInventory
            .Rows = 1
            i = 0
            While Rst.EOF <> True
                .Rows = .Rows + 1
                i = i + 1
                .TextMatrix(i, 0) = i
                .TextMatrix(i, 1) = Rst!Code
                .TextMatrix(i, 2) = Rst!InventoryName
                .TextMatrix(i, 3) = Rst!InventoryNo
                .TextMatrix(i, 4) = Rst!Mojodi
                Rst.MoveNext
            Wend
        End With
    End If
    Set Rst = Nothing
'    vsInventory.Cell(flexcpAlignment, 0, 0, 0, 3) = flexAlignCenterCenter
    vsInventory.ColAlignment(-1) = flexAlignCenterCenter
    vsInventory.Row = 1
    
End Sub
Private Sub ResizeKit1_ExitResize(ByVal XScale As Double, ByVal YScale As Double)


    If formloadFlag = True Then
        SaveSetting strMainKey, Me.Name, "Height", Me.Height
        SaveSetting strMainKey, Me.Name, "Width", Me.Width
    End If

End Sub
Private Sub vsInventory_AfterSort(ByVal Col As Long, Order As Integer)
    For i = 1 To vsInventory.Rows - 1
        vsInventory.TextMatrix(i, 0) = i
    Next
    
End Sub
Private Sub vsInventory_DblClick()
   If vsInventory.Row > 0 Then
        OKButton_Click
    End If
End Sub
