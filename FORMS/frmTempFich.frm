VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{7AEDC602-D94C-11D1-BB7A-00E0290EA3C9}#1.0#0"; "ResizeKit.ocx"
Begin VB.Form frmTempFich 
   BackColor       =   &H00FFC0C0&
   ClientHeight    =   8520
   ClientLeft      =   2775
   ClientTop       =   3870
   ClientWidth     =   11895
   Icon            =   "frmTempFich.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   8520
   ScaleWidth      =   11895
   Begin VB.CommandButton cmdDelete 
      BackColor       =   &H000000FF&
      Caption         =   "Õ–› ›Ì‘"
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
      TabIndex        =   1
      Top             =   4440
      Width           =   1365
   End
   Begin VB.CommandButton CancelButton 
      BackColor       =   &H000080FF&
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
      Left            =   120
      TabIndex        =   3
      Top             =   7800
      Width           =   1365
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
      TabIndex        =   2
      Top             =   7800
      Width           =   1485
   End
   Begin VSFlex7LCtl.VSFlexGrid vsFactorItems 
      Height          =   2565
      Left            =   1680
      TabIndex        =   4
      Top             =   5160
      Width           =   9765
      _cx             =   17224
      _cy             =   4524
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
      BackColorFixed  =   16761024
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
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
      AllowUserResizing=   0
      SelectionMode   =   0
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
      FormatString    =   ""
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
   Begin VSFlex7LCtl.VSFlexGrid vsTempFactors 
      Height          =   3765
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   11295
      _cx             =   19923
      _cy             =   6641
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
      BackColorFixed  =   16761024
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   12648447
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   0   'False
      AllowUserResizing=   3
      SelectionMode   =   3
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
      FormatString    =   $"frmTempFich.frx":A4C2
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
   Begin RESIZEKITLibCtl.ResizeKit ResizeKit1 
      Height          =   480
      Left            =   0
      OleObjectBlob   =   "frmTempFich.frx":A5A2
      TabIndex        =   7
      Top             =   0
      Width           =   480
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "›Ì‘ „Êﬁ "
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
      Left            =   4320
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   0
      Width           =   3015
   End
   Begin VB.Label lblEditedFactorDetails 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "«ﬁ·«„ ›«ò Ê—"
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
      Height          =   345
      Left            =   9360
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   4560
      Width           =   2175
   End
End
Attribute VB_Name = "frmTempFich"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Public mvarcode As String

Dim clsDate As New clsDate
Dim i As Integer
Dim Parameter() As Parameter

Public Sub FillvsTempFactors()
   
    Dim Rst As New ADODB.Recordset

    If Rst.State = 1 Then Rst.Close
    
    With vsTempFactors
    
        Dim ParametersArrays(2) As Parameter
        
        ParametersArrays(0) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
        ParametersArrays(1) = GenerateInputParameter("@Status", adInteger, 4, mvarStatus)
        ParametersArrays(2) = GenerateInputParameter("@Branch", adInteger, 4, CurrentBranch)
        Set Rst = RunParametricStoredProcedure2Rec("Get_TempFactors", ParametersArrays)
        .Rows = 1
        vsFactorItems.Rows = 1
        If Not (Rst.EOF = True And Rst.BOF = True) Then
        
            i = 1
            While Rst.EOF = False
                .Rows = .Rows + 1
                .TextMatrix(i, 0) = i
                .TextMatrix(i, 1) = Rst.Fields("intSerialNo").Value
                .TextMatrix(i, 2) = Rst.Fields("NO").Value
                .TextMatrix(i, 3) = IIf(Rst.Fields("Membershipid").Value = "-1", " ", Rst.Fields("FullName").Value)
                .TextMatrix(i, 4) = IIf(IsNull(Rst.Fields("NvcDescription").Value), "", Rst.Fields("NvcDescription").Value)
                .TextMatrix(i, 5) = IIf(IsNull(Rst.Fields("Address").Value), " ", Rst.Fields("Address").Value) '
                .TextMatrix(i, 6) = Rst.Fields("SumPrice").Value
                .TextMatrix(i, 7) = IIf(Rst.Fields("Membershipid").Value = "-1", " ", Rst.Fields("Membershipid").Value)
                .TextMatrix(i, 8) = Rst.Fields("Date").Value
                .TextMatrix(i, 9) = Rst.Fields("Time").Value
                .TextMatrix(i, 10) = Rst.Fields("nvcFirstName").Value & " " & Rst.Fields("nvcSurName").Value
                .TextMatrix(i, 11) = Rst.Fields("ServePlace").Value  'OrderType
                .TextMatrix(i, 12) = Rst.Fields("DiscountTotal").Value
                .TextMatrix(i, 13) = Rst.Fields("CarryFeeTotal").Value
                .TextMatrix(i, 14) = Rst.Fields("ServiceTotal").Value
                .TextMatrix(i, 15) = Rst.Fields("PackingTotal").Value
                .TextMatrix(i, 16) = IIf(IsNull(Rst.Fields("TempAddress").Value), " ", Rst.Fields("TempAddress").Value) '
                
                i = i + 1
                Rst.MoveNext
            Wend
        End If
      '  .AutoSizeMode = flexAutoSizeColWidth
      '  .AutoSize 0, .Cols - 1
        If .Rows > 1 Then .Row = 1
    End With
    
    If Rst.State <> 0 Then Rst.Close
    Set Rst = Nothing
    
End Sub
Public Sub FillvsFactorItems()
    
    Dim i As Integer
    Dim intselFactor As Long
    
    Dim Rst As New ADODB.Recordset
    
    If Rst.State = 1 Then Rst.Close
'    lblFactorDetail.Caption = ""
    
    With vsFactorItems
        .Rows = 1
        If vsTempFactors.Rows <= 1 Then Exit Sub
        intselFactor = vsTempFactors.TextMatrix(vsTempFactors.Row, 1)
        
        Dim ParametersArrays(2) As Parameter
        
        ParametersArrays(0) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
        ParametersArrays(1) = GenerateInputParameter("@intSerialNo", adBigInt, 8, intselFactor)
        ParametersArrays(2) = GenerateInputParameter("@Branch", adInteger, 4, CurrentBranch)
        Set Rst = RunParametricStoredProcedure2Rec("Get_Factor_Detail_Temp", ParametersArrays)
        
               
        If Not (Rst.EOF = True And Rst.BOF = True) Then
'            lblFactorDetail.Caption = "—Ì“ «ﬁ·«„ ›«ò Ê— " & vsTempFactors.TextMatrix(vsTempFactors.Row, 2)
'            Rst.moveFirst
            i = 1
            While Rst.EOF = False ' fill the grid
                .Rows = .Rows + 1
                .TextMatrix(i, 0) = Rst.Fields("intRow").Value
                .TextMatrix(i, 1) = Rst.Fields("Amount").Value
                .TextMatrix(i, 2) = Rst.Fields("Name").Value
                .TextMatrix(i, 3) = Rst.Fields("FeeUnit").Value
                .TextMatrix(i, 4) = Rst.Fields("FeeUnit").Value * Rst.Fields("Amount").Value
'                .TextMatrix(i, 5) = Rst.Fields("ServePlace").Value
                .TextMatrix(i, 5) = IIf(IsNull(Rst!DifferencesDescription), "", Rst!DifferencesDescription)
                
                Rst.MoveNext
                i = i + 1
            Wend
            
        End If
        .AutoSizeMode = flexAutoSizeColWidth  ' set the collumns' width
        .AutoSize 0, .Cols - 1
    End With
    
    If Rst.State <> 0 Then Rst.Close
    Set Rst = Nothing
    
End Sub

Private Sub CancelButton_Click()
    Me.mvarcode = ""
    Unload Me
End Sub


Private Sub cmdDelete_Click()
    With vsTempFactors
        If .SelectedRows < 1 Then Exit Sub
        
        Dim strSelectedFactors As String
               
        For i = 0 To .SelectedRows - 1
            strSelectedFactors = strSelectedFactors & "," & .TextMatrix(.SelectedRow(i), 1)
        Next i
        If strSelectedFactors <> "" Then
            ReDim Parameter(1) As Parameter
            Parameter(0) = GenerateInputParameter("@strSelectedFactors", adVarWChar, 4000, strSelectedFactors)
            Parameter(1) = GenerateInputParameter("@Branch", adInteger, 4, CurrentBranch)
            RunParametricStoredProcedure "DeleteTempFactors", Parameter
        End If
    End With
    FillvsTempFactors
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If (KeyCode = 46 Or KeyCode = 189) And Shift = 0 Then
        cmdDelete_Click
    End If
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call PresetScreenSaver
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call PresetScreenSaver
End Sub

Private Sub Form_Load()

    Dim Rst As New ADODB.Recordset
    
    CenterCenterinSecondScreen Me
    
    Dim strTemp As String
        
    With vsTempFactors
        .Rows = 1
        .Cols = 17
        
        .TextMatrix(0, 0) = ""
        .TextMatrix(0, 1) = "”—Ì«· ›Ì‘"
        .TextMatrix(0, 2) = "òœ ›Ì‘"
        .TextMatrix(0, 3) = "‰«„ „‘ —Ì"
        .TextMatrix(0, 4) = " Ê÷ÌÕ« "
        .TextMatrix(0, 5) = "¬œ—”"
        .TextMatrix(0, 6) = "Ã„⁄"
        .TextMatrix(0, 7) = "«‘ —«ﬂ"
        .TextMatrix(0, 8) = " «—ÌŒ"
        .TextMatrix(0, 9) = "”«⁄ "
        .TextMatrix(0, 10) = "‰«„ ò«—»—"
        .TextMatrix(0, 11) = "‰Ê⁄ ”—Ê"
        .TextMatrix(0, 12) = " Œ›Ì›"
        .TextMatrix(0, 13) = "Â“Ì‰Â Õ„·"
        .TextMatrix(0, 14) = "”—ÊÌ”"
        .TextMatrix(0, 15) = " »” Â »‰œÌ"
        .TextMatrix(0, 16) = " ¬œ—” „Êﬁ  "
        .AutoSearch = flexSearchFromCursor
                
        If Rst.State = 1 Then Rst.Close
        
        ReDim Parameter(0) As Parameter
        Parameter(0) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
        Set Rst = RunParametricStoredProcedure2Rec("Get_Serveplace", Parameter)
        
        strTemp = .BuildComboList(Rst, "Description", "intServePlace")
        .ColComboList(11) = strTemp
        If Rst.State <> 0 Then Rst.Close
        
    End With
    With vsFactorItems
        .Rows = 1
        .Cols = 6
        .ColAlignment(-1) = flexAlignRightCenter
        .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter
        .TextMatrix(0, 0) = "—œÌ›"
        .TextMatrix(0, 1) = "„ﬁœ«—"
        .TextMatrix(0, 2) = " ‰«„ "
        .TextMatrix(0, 3) = "›Ì"
        .TextMatrix(0, 4) = "ﬁÌ„ "
        .TextMatrix(0, 5) = " €ÌÌ—« "
        
        If Rst.State = 1 Then Rst.Close
        
        Dim Parameter1(0) As Parameter
        Parameter1(0) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
        Set Rst = RunParametricStoredProcedure2Rec("Get_Serveplace", Parameter1)
        
        strTemp = .BuildComboList(Rst, "Description", "intServePlace")
        .ColComboList(5) = strTemp
        If Rst.State <> 0 Then Rst.Close
        
        Set Rst = Nothing
        
        .AutoSearch = flexSearchFromCursor
    
    End With
    
    FlexGridActive
    FillvsTempFactors
'    FillvsFactorItems
    If vsTempFactors.Rows > 1 Then vsTempFactors.Row = 1
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


    Set clsDate = Nothing
    If PosConnection.State = adStateOpen Then PosConnection.Close: Set PosConnection = Nothing
End Sub

Private Sub OKButton_Click()
    With vsTempFactors
        If .SelectedRows < 1 Then
            frmMsg.fwlblMsg.Caption = "‘„« »«Ìœ Ìò ›«ò Ê— «‰ Œ«» ‰„«ÌÌœ"
            frmMsg.fwBtn(0).ButtonType = flwButtonOk
            frmMsg.fwBtn(0).Caption = "ﬁ»Ê·"
            frmMsg.Show vbModal
                    
        Else
            Me.mvarcode = .TextMatrix(.SelectedRow(0), 2)
            Unload Me
        End If
    End With
End Sub

Private Sub ResizeKit1_ExitResize(ByVal XScale As Double, ByVal YScale As Double)
    If formloadFlag = True Then
        SaveSetting strMainKey, Me.Name, "Height", Me.Height
        SaveSetting strMainKey, Me.Name, "Width", Me.Width
    End If

End Sub
Private Sub FlexGridActive()

    With vsTempFactors
        .ForeColor = &H40&
        
        .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter
        .ColAlignment(-1) = flexAlignCenterCenter
        .ColAlignment(5) = flexAlignRightCenter
        
         For i = 1 To .Cols - 1
             .ColWidth(i) = Val(GetSetting(strMainKey, "vsTempFactors", "Col" & i))
         Next i
         
        If .ColWidth(1) = 0 Then
            .ColWidth(1) = .Width / 20       '
        End If
        If .ColWidth(2) = 0 Then
            .ColWidth(2) = .Width / 20        '
        End If
        If .ColWidth(3) = 0 Then
            .ColWidth(3) = .Width / 8      '
        End If
        If .ColWidth(4) = 0 Then
            .ColWidth(4) = .Width / 6
        End If
        If .ColWidth(5) = 0 Then
            .ColWidth(5) = .Width / 6
        End If
        If .ColWidth(8) = 0 Then
            .ColWidth(8) = .Width / 8        '
        End If
        If .ColWidth(10) = 0 Then
            .ColWidth(10) = .Width / 7       '
        End If
        If .ColWidth(12) <= 20 Then
            .ColWidth(12) = .Width / 12       '
        End If
        If .ColWidth(13) = 0 Then
            .ColWidth(13) = .Width / 12       '
        End If
        If .ColWidth(14) <= 20 Then
            .ColWidth(14) = .Width / 5       '
        End If
        If .ColWidth(15) = 0 Then
            .ColWidth(15) = .Width / 6       '
        End If
        If .ColWidth(16) = 0 Then
            .ColWidth(16) = .Width / 6       '
        End If
''''        .AutoSizeMode = flexAutoSizeColWidth
''''        .AutoSize 8, .Cols - 1
   
    End With


End Sub

Private Sub vsTempFactors_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    For i = 0 To vsTempFactors.Cols - 1
        SaveSetting strMainKey, "vsTempFactors", "Col" & i, vsTempFactors.ColWidth(i)
    Next
End Sub


Private Sub vsTempFactors_DblClick()
    OKButton_Click
End Sub

Private Sub vsTempFactors_SelChange()
    FillvsFactorItems
End Sub
