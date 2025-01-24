VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{7AEDC602-D94C-11D1-BB7A-00E0290EA3C9}#1.0#0"; "ResizeKit.ocx"
Begin VB.Form frmGoodPrintFormats 
   ClientHeight    =   9465
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10665
   Icon            =   "frmGoodPrintFormats.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   9465
   ScaleWidth      =   10665
   Begin VB.Frame Frame1 
      Height          =   4275
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   5160
      Width           =   10305
      Begin VB.CommandButton cmdSelectAllAdded 
         BackColor       =   &H000000C0&
         Caption         =   "«‰ Œ«» Â„Â"
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
         Left            =   690
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Tag             =   "1"
         Top             =   2640
         Width           =   2415
      End
      Begin VB.CommandButton cmdDelete 
         BackColor       =   &H000000C0&
         Caption         =   "Õ–› ò«·« «“ Õ«·  ç«Å"
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
         Left            =   720
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   3480
         Width           =   2415
      End
      Begin VSFlex7LCtl.VSFlexGrid vsAddedGoods 
         Height          =   3660
         Left            =   3810
         TabIndex        =   7
         Top             =   540
         Width           =   6255
         _cx             =   11033
         _cy             =   6456
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
         ForeColor       =   8388608
         BackColorFixed  =   16777152
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
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmGoodPrintFormats.frx":A4C2
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
      Begin VB.Label lblAddedGoods 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFC0C0&
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
         ForeColor       =   &H00FF0000&
         Height          =   435
         Left            =   5130
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   120
         Width           =   4935
      End
   End
   Begin VB.Frame Frame2 
      Height          =   4665
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   480
      Width           =   10305
      Begin VB.CommandButton cmdSelectAllDeleted 
         BackColor       =   &H00008000&
         Caption         =   "«‰ Œ«» Â„Â"
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
         Left            =   2940
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Tag             =   "1"
         Top             =   4080
         Width           =   1695
      End
      Begin VB.CommandButton cmdAdd 
         BackColor       =   &H00008000&
         Caption         =   "«÷«›Â ò—œ‰ ò«·« »Â Õ«·  ç«Å "
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
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   4080
         Width           =   2655
      End
      Begin VB.ListBox lstPrintFormats 
         BeginProperty Font 
            Name            =   "Nazanin"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2790
         Left            =   6240
         RightToLeft     =   -1  'True
         Style           =   1  'Checkbox
         TabIndex        =   2
         Top             =   600
         Width           =   3975
      End
      Begin VSFlex7LCtl.VSFlexGrid vsDeletedGoods 
         Height          =   3300
         Left            =   120
         TabIndex        =   1
         Top             =   600
         Width           =   6075
         _cx             =   10716
         _cy             =   5821
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
         ForeColor       =   128
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
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmGoodPrintFormats.frx":A598
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
      Begin VB.Label lblPrintFormats 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FF0000&
         BackStyle       =   0  'Transparent
         Caption         =   "›Ê—„ Â«Ì ç«Å"
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
         Left            =   6240
         RightToLeft     =   -1  'True
         TabIndex        =   4
         Top             =   120
         Width           =   3975
      End
      Begin VB.Label lblDeletedGoods 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFC0C0&
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
         ForeColor       =   &H00000080&
         Height          =   435
         Left            =   1080
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   120
         Width           =   5055
      End
   End
   Begin RESIZEKITLibCtl.ResizeKit ResizeKit1 
      Height          =   480
      Left            =   0
      OleObjectBlob   =   "frmGoodPrintFormats.frx":A66E
      TabIndex        =   13
      Top             =   0
      Width           =   480
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "ò«·« œ— ç«Å"
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
      TabIndex        =   12
      Top             =   0
      Width           =   1455
   End
End
Attribute VB_Name = "frmGoodPrintFormats"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim i As Integer
Dim Parameter() As Parameter

Public Sub ExitForm()
    Unload Me
End Sub
Public Sub ChangeLanguage()
    Select Case clsStation.Language
        Case EnumLanguage.Farsi
        
            With vsAddedGoods
                .TextMatrix(0, 3) = "ê—ÊÂ"
                .TextMatrix(0, 5) = "“Ì— ê—ÊÂ"
                .TextMatrix(0, 6) = "‰«„ ò«·«"
                .AutoSizeMode = flexAutoSizeColWidth
                .AutoSize 0, .Cols - 1
            End With
            
            With vsDeletedGoods
            
                .TextMatrix(0, 3) = "ê—ÊÂ"
                .TextMatrix(0, 5) = "“Ì— ê—ÊÂ"
                .TextMatrix(0, 6) = "‰«„ ò«·«"
                .AutoSizeMode = flexAutoSizeColWidth
                .AutoSize 0, .Cols - 1
            End With
            
        Case EnumLanguage.English
        
            With vsAddedGoods
                .TextMatrix(0, 3) = "Group"
                .TextMatrix(0, 5) = "Sub Group"
                .TextMatrix(0, 6) = "Good Name"
                .AutoSizeMode = flexAutoSizeColWidth
                .AutoSize 0, .Cols - 1
            
            End With
            
            With vsDeletedGoods
                .TextMatrix(0, 3) = "Group"
                .TextMatrix(0, 5) = "Sub Group"
                .TextMatrix(0, 6) = "Good Name"
                .AutoSizeMode = flexAutoSizeColWidth
                .AutoSize 0, .Cols - 1
            End With
    
    End Select
    FilllstPrintFormats
    FillvsAddedGoods 0
    FillvsDeletedGoods 0
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call PresetScreenSaver
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call PresetScreenSaver
End Sub

Private Sub FillvsAddedGoods(Item As Integer)
    
    With vsAddedGoods
    
        .Rows = 1
        If Item = -1 Then
            Exit Sub
        End If
        
        PrintFormatCode = lstPrintFormats.ItemData(Item)
        lblAddedGoods.Caption = "ò«·«Â«Ì „ÊÃÊœ »—«Ì ç«Å œ—  " & lstPrintFormats.List(Item)
        
        Dim Rst As New ADODB.Recordset
        ReDim Parameter(1) As Parameter
        Parameter(0) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
        Parameter(1) = GenerateInputParameter("@intPrinterFormat", adInteger, 4, PrintFormatCode)
        Set Rst = RunParametricStoredProcedure2Rec("Get_AddedGoods", Parameter)
    
        i = 1
        While Rst.EOF <> True
            With vsAddedGoods
                .Rows = .Rows + 1
                
                .TextMatrix(i, 1) = Rst!Code
                .TextMatrix(i, 2) = Rst!level1
                .TextMatrix(i, 3) = Rst!DesLevel1
                .TextMatrix(i, 4) = Rst!Level2
                .TextMatrix(i, 5) = Rst!DesLevel2
                .TextMatrix(i, 6) = Rst!Name
                
                i = i + 1
                Rst.MoveNext
            End With
        Wend
        
        If .Rows > 1 Then
            .Cell(flexcpAlignment, 1, 1, .Rows - 1, .Cols - 1) = flexAlignRightCenter
            cmdSelectAllAdded.Enabled = True
        Else
            cmdSelectAllAdded.Enabled = False

        End If
        
        .AutoSizeMode = flexAutoSizeColWidth
        .AutoSize 0, .Cols - 1
    
    End With
    
    Set Rst = Nothing
End Sub

Private Sub FillvsDeletedGoods(Item As Integer)
    
    With vsDeletedGoods
    
        .Rows = 1
        
        If Item = -1 Then
            Exit Sub
        End If
        
        PrintFormatCode = lstPrintFormats.ItemData(Item)
        lblDeletedGoods.Caption = "ò«·«Â«Ì Õ–› ‘œÂ «“  " & lstPrintFormats.List(Item)
        
        Dim Rst As New ADODB.Recordset
        ReDim Parameter(1) As Parameter
        Parameter(0) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
        Parameter(1) = GenerateInputParameter("@intPrinterFormat", adInteger, 4, PrintFormatCode)
        Set Rst = RunParametricStoredProcedure2Rec("Get_DeletedGoods", Parameter)
        
        i = 1
        While Rst.EOF <> True
            With vsDeletedGoods
            
                .Rows = .Rows + 1
                
                .TextMatrix(i, 1) = Rst!Code
                .TextMatrix(i, 2) = Rst!level1
                .TextMatrix(i, 3) = Rst!DesLevel1
                .TextMatrix(i, 4) = Rst!Level2
                .TextMatrix(i, 5) = Rst!DesLevel2
                .TextMatrix(i, 6) = Rst!Name
                
                i = i + 1
                Rst.MoveNext
            End With
        Wend
        If .Rows > 1 Then
            .Cell(flexcpAlignment, 1, 1, .Rows - 1, .Cols - 1) = flexAlignRightCenter
            cmdSelectAllDeleted.Enabled = True
        Else
            cmdSelectAllDeleted.Enabled = False
        End If
        .AutoSizeMode = flexAutoSizeColWidth
        .AutoSize 0, .Cols - 1
    
    End With
    
    Set Rst = Nothing
End Sub
Private Sub FilllstPrintFormats()

    Dim Rst As New ADODB.Recordset
    ReDim Parameter(0) As Parameter
    
    Parameter(0) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
    
    Set Rst = RunParametricStoredProcedure2Rec("GetPrintFormat", Parameter)
    
    lstPrintFormats.Clear
    
    While Rst.EOF <> True
    
        lstPrintFormats.AddItem Rst!PrintFormatName
        lstPrintFormats.ItemData(lstPrintFormats.ListCount - 1) = Rst!PrintFormat
        Rst.MoveNext
    Wend
    
    Set Rst = Nothing
End Sub



Private Sub cmdAdd_Click()

    
    Dim intPrintFormats As Integer
            
    Dim strMessage As String
    Dim SelectedGoods As String
    
    If lstPrintFormats.SelCount = 0 Then
        strMessage = strMessage & "·ÿ›« «» œ« Ìò ›Ê—„  ç«Å «‰ Œ«» ‰„«ÌÌœ" & vbCrLf
    End If
    
    If vsDeletedGoods.Rows < 2 Then
        strMessage = strMessage & "‘„« „Ì »«Ì”  Õœ«ﬁ· Ìò ò«·« «‰ Œ«» ‰„«ÌÌœ" & vbCrLf
    Else
        For i = 1 To vsDeletedGoods.Rows - 1
            If Val(vsDeletedGoods.TextMatrix(i, 0)) <> 0 Then
                SelectedGoods = SelectedGoods & vsDeletedGoods.TextMatrix(i, 1) & ","
            End If
        
        Next i
        If SelectedGoods = "" Then
            strMessage = strMessage & "‘„« „Ì »«Ì”  Õœ«ﬁ· Ìò ò«·« «‰ Œ«» ‰„«ÌÌœ" & vbCrLf
        Else
            SelectedGoods = Left(SelectedGoods, Len(SelectedGoods) - 1)
        End If
    End If
   
    If strMessage <> "" Then
        frmMsg.fwlblMsg.Caption = strMessage
        frmMsg.fwBtn(0).Visible = False
        frmMsg.fwBtn(1).ButtonType = flwButtonOk
        frmMsg.fwBtn(1).Caption = "ﬁ»Ê·"
        frmMsg.Show vbModal
        Exit Sub
    End If
    
    With lstPrintFormats
        For i = 0 To .ListCount - 1
            If .Selected(i) = True Then
                intPrintFormats = .ItemData(i)
                Exit For
            End If
        Next i
    End With
    
    Dim Item As Integer
    
    Item = i
    
    strMessage = "ò«·«Â«Ì „Ê—œ ‰Ÿ— »Â  " & lstPrintFormats.List(Item) & " «÷«›Â ‘œ"
   
    ReDim Parameter(1) As Parameter
    
    For i = 1 To vsDeletedGoods.Rows - 1
        If Val(vsDeletedGoods.TextMatrix(i, 0)) <> 0 Then
        
            Parameter(0) = GenerateInputParameter("@intPrinterFormat", adInteger, 4, intPrintFormats)
            Parameter(1) = GenerateInputParameter("@GoodCode", adInteger, 4, vsDeletedGoods.TextMatrix(i, 1))
            RunParametricStoredProcedure "Delete_DeletedGoodsFromPrintFormat", Parameter
            
        End If
    
    Next i
    
    FillvsAddedGoods Item
    FillvsDeletedGoods Item
    
    frmMsg.fwlblMsg.Caption = strMessage
    frmMsg.fwBtn(0).Visible = False
    frmMsg.fwBtn(1).ButtonType = flwButtonOk
    frmMsg.fwBtn(1).Caption = "ﬁ»Ê·"
    frmMsg.Show vbModal


End Sub

Private Sub cmdDelete_Click()
    
    Dim intPrintFormats As Integer
            
    Dim strMessage As String
    Dim SelectedGoods As String
    
    If lstPrintFormats.SelCount = 0 Then
        strMessage = strMessage & "·ÿ›« «» œ« Ìò ›Ê—„  ç«Å «‰ Œ«» ‰„«ÌÌœ" & vbCrLf
    End If
    
    If vsAddedGoods.Rows < 2 Then
        strMessage = strMessage & "‘„« „Ì »«Ì”  Õœ«ﬁ· Ìò ò«·« «‰ Œ«» ‰„«ÌÌœ" & vbCrLf
    Else
        For i = 1 To vsAddedGoods.Rows - 1
            If Val(vsAddedGoods.TextMatrix(i, 0)) <> 0 Then
                SelectedGoods = SelectedGoods & vsAddedGoods.TextMatrix(i, 1) & ","
            End If
        
        Next i
        If SelectedGoods = "" Then
            strMessage = strMessage & "‘„« „Ì »«Ì”  Õœ«ﬁ· Ìò ò«·« «‰ Œ«» ‰„«ÌÌœ" & vbCrLf
        Else
            SelectedGoods = Left(SelectedGoods, Len(SelectedGoods) - 1)
        End If
    End If
   
    If strMessage <> "" Then
        frmMsg.fwlblMsg.Caption = strMessage
        frmMsg.fwBtn(0).Visible = False
        frmMsg.fwBtn(1).ButtonType = flwButtonOk
        frmMsg.fwBtn(1).Caption = "ﬁ»Ê·"
        frmMsg.Show vbModal
        Exit Sub
    End If
    
    With lstPrintFormats
        For i = 0 To .ListCount - 1
            If .Selected(i) = True Then
                intPrintFormats = .ItemData(i)
                Exit For
            End If
        Next i
    End With
    
    Dim Item As Integer
    
    Item = i
    
    strMessage = "ò«·«Â«Ì „Ê—œ ‰Ÿ— «“  " & lstPrintFormats.List(Item) & " Õ–› ‘œ"
   
    ReDim Parameter(1) As Parameter
    
    For i = 1 To vsAddedGoods.Rows - 1
        If Val(vsAddedGoods.TextMatrix(i, 0)) <> 0 Then
        
            Parameter(0) = GenerateInputParameter("@intPrinterFormat", adInteger, 4, intPrintFormats)
            Parameter(1) = GenerateInputParameter("@GoodCode", adInteger, 4, vsAddedGoods.TextMatrix(i, 1))
            RunParametricStoredProcedure "Insert_DeletedGoodsFromPrintFormat", Parameter
            
        End If
    
    Next i
    
    FillvsAddedGoods Item
    FillvsDeletedGoods Item
    
    frmMsg.fwlblMsg.Caption = strMessage
    frmMsg.fwBtn(0).Visible = False
    frmMsg.fwBtn(1).ButtonType = flwButtonOk
    frmMsg.fwBtn(1).Caption = "ﬁ»Ê·"
    frmMsg.Show vbModal
    
    
End Sub

Private Sub cmdSelectAllAdded_Click()
    If vsAddedGoods.Rows > 1 Then
        If cmdSelectAllAdded.Tag = 1 Then
            cmdSelectAllAdded.Tag = 0
            vsAddedGoods.Cell(flexcpText, vsAddedGoods.FixedRows, 0, vsAddedGoods.Rows - 1, 0) = -1
            cmdSelectAllAdded.Caption = "Å«ò ò—œ‰ Â„Â"
        Else
            cmdSelectAllAdded.Tag = 1
            vsAddedGoods.Cell(flexcpText, vsAddedGoods.FixedRows, 0, vsAddedGoods.Rows - 1, 0) = 0
            cmdSelectAllAdded.Caption = "«‰ Œ«» Â„Â"
        End If
    End If

End Sub

Private Sub cmdSelectAllDeleted_Click()
    If vsDeletedGoods.Rows > 1 Then
        If cmdSelectAllDeleted.Tag = 1 Then
            cmdSelectAllDeleted.Tag = 0
            vsDeletedGoods.Cell(flexcpText, vsDeletedGoods.FixedRows, 0, vsDeletedGoods.Rows - 1, 0) = -1
            cmdSelectAllDeleted.Caption = "Å«ò ò—œ‰ Â„Â"
        Else
            cmdSelectAllDeleted.Tag = 1
            vsDeletedGoods.Cell(flexcpText, vsDeletedGoods.FixedRows, 0, vsDeletedGoods.Rows - 1, 0) = 0
            cmdSelectAllDeleted.Caption = "«‰ Œ«» Â„Â"
        End If
    End If
End Sub

Private Sub Form_Activate()

    mdifrm.Toolbar1.Buttons(23).Enabled = True
    mdifrm.Toolbar1.Buttons(24).Enabled = True
    mdifrm.Toolbar1.Buttons(25).Enabled = True
    mdifrm.Toolbar1.Buttons(26).Enabled = True
    mdifrm.Toolbar1.Buttons(27).Enabled = True

    VarActForm = Me.Name
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
        
    If intVersion = Min Then
        ShowDisMessage "ﬁ«»·Ì  ›Ì· — ﬂ—œ‰ ﬂ«·«Â« œ— ç«Å œ— ‰”ŒÂ Â«Ì »«·« — ÊÃÊœ œ«—œ", 1500
        Unload Me
        Exit Sub
    End If
    
    CenterTop Me
    
    With vsAddedGoods
        .Cols = 7
        .Rows = 1
        .ColHidden(1) = True
        .ColHidden(2) = True
        .ColHidden(4) = True
        .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter
        .AutoSizeMode = flexAutoSizeColWidth
        .AutoSize 0, .Cols - 1
    End With
    
    With vsDeletedGoods
        .Cols = 7
        .Rows = 1
        .ColHidden(1) = True
        .ColHidden(2) = True
        .ColHidden(4) = True
        .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter
        .AutoSizeMode = flexAutoSizeColWidth
        .AutoSize 0, .Cols - 1
    End With
    
    ChangeLanguage
    
    
    FilllstPrintFormats
    
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
    
     VarActForm = ""
     
End Sub

Private Sub lstPrintFormats_ItemCheck(Item As Integer)

    With lstPrintFormats
    
        If lstPrintFormats.Selected(Item) = True Then
        
            For i = 0 To lstPrintFormats.ListCount - 1
                If i <> Item Then
                    lstPrintFormats.Selected(i) = False
                End If
            Next i
            
            FillvsDeletedGoods Item
            FillvsAddedGoods Item
        Else
            vsDeletedGoods.Rows = 1
            vsAddedGoods.Rows = 1
            
        End If
        
    End With
    
    
End Sub




Private Sub ResizeKit1_ExitResize(ByVal XScale As Double, ByVal YScale As Double)


    If formloadFlag = True Then
        SaveSetting strMainKey, Me.Name, "Height", Me.Height
        SaveSetting strMainKey, Me.Name, "Width", Me.Width
    End If

End Sub

Private Sub vsAddedGoods_KeyPress(KeyAscii As Integer)
    If KeyAscii <> 32 Then Exit Sub
    With vsAddedGoods
        If .Col = 0 And .Row > 0 Then
            .Select .Row, .Col
            .EditCell
        End If
    End With

End Sub

Private Sub vsAddedGoods_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    With vsAddedGoods
        If Button = 1 And .Col = 0 And .Row > 0 Then
            .Select .Row, .Col
            .EditCell
        End If
    End With

End Sub

Private Sub vsDeletedGoods_KeyPress(KeyAscii As Integer)
    If KeyAscii <> 32 Then Exit Sub
    With vsDeletedGoods
        If .Col = 0 And .Row > 0 Then
            .Select .Row, .Col
            .EditCell
        End If
    End With
    
End Sub

Private Sub vsDeletedGoods_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    With vsDeletedGoods
        If Button = 1 And .Col = 0 And .Row > 0 Then
            .Select .Row, .Col
            .EditCell
        End If
    End With
    
End Sub
