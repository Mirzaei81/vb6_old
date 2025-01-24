VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{7AEDC602-D94C-11D1-BB7A-00E0290EA3C9}#1.0#0"; "ResizeKit.ocx"
Object = "{75D4F767-8785-11D3-93AD-0000832EF44D}#2.14#0"; "FAST2001.ocx"
Begin VB.Form frmFindFactor 
   BackColor       =   &H00E0E0E0&
   Caption         =   "        Ã” ÃÊÌ ›«ò Ê—"
   ClientHeight    =   8865
   ClientLeft      =   2775
   ClientTop       =   3870
   ClientWidth     =   9540
   Icon            =   "frmFindFactor.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   8865
   ScaleWidth      =   9540
   Begin VB.TextBox txtDescription 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   4920
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   600
      Width           =   3195
   End
   Begin VB.PictureBox Picture2 
      Height          =   615
      Left            =   240
      ScaleHeight     =   555
      ScaleWidth      =   3315
      TabIndex        =   16
      Top             =   600
      Width           =   3375
      Begin VB.OptionButton optStatus 
         Alignment       =   1  'Right Justify
         Caption         =   "—”Ìœ „Êﬁ "
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
         Left            =   0
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   120
         Width           =   1875
      End
      Begin VB.OptionButton optStatus 
         Alignment       =   1  'Right Justify
         Caption         =   "Œ—Ìœ"
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
         Left            =   2160
         RightToLeft     =   -1  'True
         TabIndex        =   17
         Top             =   120
         Value           =   -1  'True
         Width           =   915
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   975
      Left            =   4200
      ScaleHeight     =   915
      ScaleWidth      =   5115
      TabIndex        =   7
      Top             =   1320
      Width           =   5175
      Begin VB.TextBox txtDate2 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   150
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Top             =   180
         Width           =   2325
      End
      Begin VB.TextBox txtDate1 
         Alignment       =   1  'Right Justify
         Height          =   345
         Left            =   2760
         RightToLeft     =   -1  'True
         TabIndex        =   19
         Top             =   120
         Width           =   1935
      End
      Begin VB.OptionButton optShowFich 
         Alignment       =   1  'Right Justify
         Caption         =   "‰„«Ì‘ Â„Â œ— „ÕœÊœÂ  «—ÌŒ"
         BeginProperty Font 
            Name            =   "Nazanin"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   1
         Left            =   0
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Top             =   480
         Value           =   -1  'True
         Width           =   2955
      End
      Begin VB.OptionButton optShowFich 
         Alignment       =   1  'Right Justify
         Caption         =   "›ﬁÿ ”‰œ «‰ Œ«»Ì"
         BeginProperty Font 
            Name            =   "Nazanin"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   0
         Left            =   3240
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   480
         Width           =   1875
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0000C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   "«“"
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
         Left            =   4680
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   120
         Width           =   255
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         BackColor       =   &H0000C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   " «"
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
         Left            =   2400
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   120
         Width           =   255
      End
   End
   Begin FLWCtrls.FWProgressBar FWProgressBar1 
      Height          =   375
      Left            =   1920
      Top             =   8040
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   661
      Max             =   1000
      BorderStyle     =   10
   End
   Begin VB.CommandButton OKButton 
      BackColor       =   &H0000C000&
      Caption         =   "ﬁ»Ê·"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "B Homa"
         Size            =   11.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2400
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton CancelButton 
      BackColor       =   &H000000C0&
      Cancel          =   -1  'True
      Caption         =   "«‰’—«›"
      BeginProperty Font 
         Name            =   "B Homa"
         Size            =   11.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   240
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   1440
      Width           =   1215
   End
   Begin VB.TextBox txtNo 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   6960
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   1155
   End
   Begin VSFlex7LCtl.VSFlexGrid vsFactors_Fich 
      Height          =   5565
      Left            =   60
      TabIndex        =   1
      Top             =   2400
      Width           =   9435
      _cx             =   16642
      _cy             =   9816
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
      BackColorBkg    =   -2147483633
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
      Cols            =   8
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   500
      RowHeightMax    =   500
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmFindFactor.frx":A4C2
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
      OleObjectBlob   =   "frmFindFactor.frx":A589
      TabIndex        =   6
      Top             =   0
      Width           =   480
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0000C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "‘—Õ ”‰œ"
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
      Left            =   8160
      RightToLeft     =   -1  'True
      TabIndex        =   12
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "œ— Õ«·  «‰ Œ«» ‰„«Ì‘ Â„Â ' ”Ì” „ ﬂ·ÌÂ «ﬁ·«„ ›«ﬂ Ê—Â« Ì «Ì‰ ﬂ«—»— —« œ— ’Ê—  ÊÃÊœ œ” —”Ì  ‰„«Ì‘ „Ì œÂœ"
      BeginProperty Font 
         Name            =   "Nazanin"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   615
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   15
      Top             =   8400
      Width           =   9135
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0000C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "‰„«Ì‘ ”‰œ Â«"
      BeginProperty Font 
         Name            =   "Nazanin"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   375
      Left            =   7920
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   960
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0000C0C0&
      BackStyle       =   0  'Transparent
      Caption         =   "‘„«—Â ”‰œ"
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
      Left            =   8160
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label LblFindFactor 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
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
      Height          =   615
      Left            =   240
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   0
      Width           =   3375
   End
End
Attribute VB_Name = "frmFindFactor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim i As Long
Dim FactorType As EnumFactorType
Dim Parameter() As Parameter
Dim clsDate As New clsDate

Private Sub CancelButton_Click()
    mvarcode = 0
    Unload Me
End Sub

Private Sub Form_Activate()
    With vsFactors_Fich
        .Cols = 8
        .TextMatrix(0, 7) = " Ê÷ÌÕ« "
        .FixedAlignment(-1) = flexAlignCenterCenter
        .ColAlignment(-1) = flexAlignCenterCenter
        .ColAlignment(7) = flexAlignRightCenter
        .ColHidden(7) = False
    End With
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call PresetScreenSaver
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call PresetScreenSaver
End Sub

Private Sub Form_Load()

    CenterCenterinSecondScreen Me
    
    mvarcode = 0
    
    FWProgressBar1.Visible = False
'    If clsStation.SearchFichDefault = True Then
'        optShowFich(0).Value = True
'    Else
        optShowFich(0).Value = False
'    End If

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

    txtDate1.Text = "13" & Right(clsDate.shamsi(Date), 8)
   ' txtDate1.Text = AccountYear & "/01/01"
    txtDate2.Text = "13" & Right(clsDate.shamsi(Date), 8)
    optShowFich_Click 1

    If mvarStatus <> Purchase Then Picture2.Visible = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If PosConnection.State = adStateOpen Then PosConnection.Close: Set PosConnection = Nothing

    SaveSetting strMainKey, Me.Name, "Left", Me.left
    SaveSetting strMainKey, Me.Name, "Top", Me.top


End Sub


Private Sub OKButton_Click()
    If vsFactors_Fich.Row > 0 Then
'        If TempStatus = TempRecieved And Val(vsFactors_Fich.TextMatrix(vsFactors_Fich.Row, 8)) <> 0 Then ShowDisMessage "", 1500
        mvarcode = vsFactors_Fich.TextMatrix(vsFactors_Fich.Row, 2)
    Else
        mvarcode = 0
    End If
    Unload Me

End Sub
Sub ClearDataFlexGrid()

    With vsFactors_Fich
        .Rows = 1
        .Cols = 8
               
    End With

    
End Sub

Private Sub optShowFich_Click(index As Integer)
    If optShowFich(1).Value = 0 Then
        ClearDataFlexGrid
        vsFactors_Fich.Row = 0
        txtDate1.Enabled = False
        txtDate2.Enabled = False
    '    vsFactors_Fich.ShowCell 1, 0
    '    vsFactors_Fich.Sort = flexSortGenericDescending
   Else
        txtDate1.Enabled = True
        txtDate2.Enabled = True
        FillvsFactors_Fich
        vsFactors_Fich.Row = 0
        If vsFactors_Fich.Rows > 1 Then
           vsFactors_Fich.ShowCell 1, 0
           vsFactors_Fich.Sort = flexSortGenericDescending
        End If
   End If
End Sub


Private Sub optStatus_Click(index As Integer)
    vsFactors_Fich.Rows = 1
    If optStatus(0).Value = True Then
        TempStatus = Purchase
    Else
        TempStatus = TempRecieved
    End If
    If formloadFlag = False Then Exit Sub
    FillvsFactors_Fich

End Sub

Private Sub ResizeKit1_ExitResize(ByVal XScale As Double, ByVal YScale As Double)

    If formloadFlag = True Then
        SaveSetting strMainKey, Me.Name, "Height", Me.Height
        SaveSetting strMainKey, Me.Name, "Width", Me.Width
    End If

End Sub

Private Sub txtDescription_Change()
    FWProgressBar1.Visible = True
    FWProgressBar1.Value = 0
    
    Dim Rst As New ADODB.Recordset
    
    ReDim Parameter(3) As Parameter
    Parameter(0) = GenerateInputParameter("@Status", adInteger, 4, TempStatus)
    Parameter(1) = GenerateInputParameter("@AccountYear", adSmallInt, 2, AccountYear)
    Parameter(2) = GenerateInputParameter("@Branch", adInteger, 4, frmPurchase.cmbBranch.ItemData(frmPurchase.cmbBranch.ListIndex))
    Parameter(3) = GenerateInputParameter("@nvcDescription", adVarWChar, 255, Trim(txtDescription.Text))
    
    Set Rst = RunParametricStoredProcedure2Rec("Get_tFacM_Description", Parameter)

    With vsFactors_Fich
        .Rows = 1
        .Cols = 8
        i = 0
        While Rst.EOF <> True
            .Rows = .Rows + 1
            i = .Rows - 1
            .TextMatrix(i, 0) = i
            .TextMatrix(i, 1) = Rst!intSerialNo
            .TextMatrix(i, 2) = Rst!No
            .TextMatrix(i, 3) = Rst![Date]
            .TextMatrix(i, 4) = Rst![time]
            .TextMatrix(i, 5) = Rst!sumPrice
            .TextMatrix(i, 6) = Rst!nvcFirstName & " " & Rst!nvcSurname
            .TextMatrix(i, 7) = Rst!NvcDescription
            Rst.MoveNext
         
        FWProgressBar1.Value = FWProgressBar1.Value + 1
        If FWProgressBar1.Value = 1000 Then
           FWProgressBar1.Value = 0
        End If
        Wend
    End With

    FWProgressBar1.Value = 0
    FWProgressBar1.Visible = False
    
    Set Rst = Nothing

End Sub

Private Sub txtDescription_GotFocus()
    vsFactors_Fich.Rows = 1
End Sub

Private Sub txtNo_Change()
    i = -1
    If optShowFich(0).Value = True Then
  '      i = vsFactors_Fich.FindRow(txtNo.Text, 1, 2, True, True)
       If Val(txtNo.Text) > 0 Then
          Define_Factor
       Else
        vsFactors_Fich.Rows = 1
        vsFactors_Fich.Cols = 8
        
       End If
    Else
        If Len(txtNo.Text) = 3 Then
           i = vsFactors_Fich.FindRow(txtNo.Text, 1, 7, True, True)
        End If
    End If
    If i > 0 Then
        vsFactors_Fich.Row = i
        vsFactors_Fich.ShowCell i, 0
        LblFindFactor.Caption = ""
    Else
        vsFactors_Fich.Row = 0
        vsFactors_Fich.ShowCell 0, 0
        If Val(txtNo.Text) > 0 Then
           LblFindFactor.Caption = " ”‰œ " & Val(txtNo.Text) & "  œ— ”Ì” „  ⁄—Ì› ‰‘œÂ"
         Else
            LblFindFactor.Caption = "‘„«—Â ”‰œ —« Ê«—œ ﬂ‰Ìœ  "
         End If
    
    End If
    
End Sub

Private Sub txtNo_GotFocus()

    vsFactors_Fich.Row = 0
    vsFactors_Fich.Select vsFactors_Fich.Row, 2
'    vsFactors_Fich.Sort = flexSortGenericAscending
    vsFactors_Fich.Sort = flexSortGenericDescending
    LblFindFactor.Caption = "‘„«—Â ”‰œ —« Ê«—œ ﬂ‰Ìœ  "
    
End Sub
Private Sub txtDate1_Change()
    If formloadFlag = False Then Exit Sub
    If Len(txtDate1.Text) = 10 Then
        FillvsFactors_Fich
        vsFactors_Fich.Row = 0
'        If vsFactors_Fich.Rows > 1 Then
'            vsFactors_Fich.ShowCell 1, 0
'            vsFactors_Fich.Sort = flexSortGenericDescending
'        End If
    End If
End Sub

Private Sub txtDate2_Change()
    If formloadFlag = False Then Exit Sub
    If Len(txtDate2.Text) = 10 Then
        FillvsFactors_Fich
        vsFactors_Fich.Row = 0
'        If vsFactors_Fich.Rows > 1 Then
'            vsFactors_Fich.ShowCell 1, 0
'            vsFactors_Fich.Sort = flexSortGenericDescending
'        End If
    End If
End Sub

Private Sub FillvsFactors_Fich()

    If Len(txtDate1.Text) <> 10 Or Len(txtDate1.Text) <> 10 Then
        ShowDisMessage "  «—ÌŒ „⁄ »— Ê«—œ ﬂ‰Ìœ ", 1000
        Exit Sub
    End If
    FWProgressBar1.Visible = True
    FWProgressBar1.Value = 0
    
    Dim Rst As New ADODB.Recordset
    
    ReDim Parameter(5) As Parameter
    Parameter(0) = GenerateInputParameter("@Status", adInteger, 4, TempStatus)
    Parameter(1) = GenerateInputParameter("@User", adInteger, 4, mvarCurUserNo)
    Parameter(2) = GenerateInputParameter("@AccountYear", adSmallInt, 2, AccountYear)
    Parameter(3) = GenerateInputParameter("@Branch", adInteger, 4, frmPurchase.cmbBranch.ItemData(frmPurchase.cmbBranch.ListIndex))
    Parameter(4) = GenerateInputParameter("@DateAfter", adVarWChar, 8, Right(txtDate1.Text, 8))
    Parameter(5) = GenerateInputParameter("@DateBefore", adVarWChar, 8, Right(txtDate2.Text, 8))
    
    Set Rst = RunParametricStoredProcedure2Rec("Get_All_Factors", Parameter)

    With vsFactors_Fich
        .Rows = 1
        .Cols = 9
        .TextMatrix(0, 8) = "—”Ìœ œ«∆„"
        .ColDataType(8) = flexDTBoolean
        i = 0
        While Rst.EOF <> True
            .Rows = .Rows + 1
            i = .Rows - 1
            .TextMatrix(i, 0) = i
            .TextMatrix(i, 1) = Rst!intSerialNo
            .TextMatrix(i, 2) = Rst!No
            .TextMatrix(i, 3) = Rst![Date]
            .TextMatrix(i, 4) = Rst![time]
            .TextMatrix(i, 5) = Rst!sumPrice
            .TextMatrix(i, 6) = Rst!nvcFirstName & " " & Rst!nvcSurname
            .TextMatrix(i, 7) = Rst!NvcDescription
            .TextMatrix(i, 8) = Rst!BitTempReceived
            If Rst!BitTempReceived = True Then .Cell(flexcpBackColor, i, 0, i, .Cols - 1) = vbRed
            Rst.MoveNext
         
        FWProgressBar1.Value = FWProgressBar1.Value + 1
        If FWProgressBar1.Value = 1000 Then
           FWProgressBar1.Value = 0
        End If
        Wend
    End With

    FWProgressBar1.Value = 0
    FWProgressBar1.Visible = False
    
    Set Rst = Nothing


End Sub

Private Sub vsFactors_Fich_AfterSort(ByVal Col As Long, Order As Integer)
    For i = 1 To vsFactors_Fich.Rows - 1
        vsFactors_Fich.TextMatrix(i, 0) = i
    Next
    
End Sub

Private Sub vsFactors_Fich_DblClick()
    If vsFactors_Fich.Row > 0 Then
        OKButton_Click
    End If
End Sub

Private Sub Define_Factor()

    
    Dim Rst As New ADODB.Recordset
    
    ReDim Parameter(4) As Parameter
    Parameter(0) = GenerateInputParameter("@Status", adInteger, 4, mvarStatus)
    Parameter(1) = GenerateInputParameter("@User", adInteger, 4, mvarCurUserNo)
    Parameter(2) = GenerateInputParameter("@No", adBigInt, 8, txtNo.Text)
    Parameter(3) = GenerateInputParameter("@AccountYear", adSmallInt, 2, AccountYear)
    Parameter(4) = GenerateInputParameter("@Branch", adInteger, 4, frmPurchase.cmbBranch.ItemData(frmPurchase.cmbBranch.ListIndex))
    
    Set Rst = RunParametricStoredProcedure2Rec("Get_Define_Factors", Parameter)

    With vsFactors_Fich
        .Rows = 1
        .Cols = 9
        .TextMatrix(0, 8) = "—”Ìœ œ«∆„"
        .ColDataType(8) = flexDTBoolean
        i = 0
        While Rst.EOF <> True
            .Rows = .Rows + 1
            i = .Rows - 1
            .TextMatrix(i, 0) = i
            .TextMatrix(i, 1) = Rst!intSerialNo
            .TextMatrix(i, 2) = Rst!No
            .TextMatrix(i, 3) = Rst![Date]
            .TextMatrix(i, 4) = Rst![time]
            .TextMatrix(i, 5) = Rst!sumPrice
            .TextMatrix(i, 6) = Rst!nvcFirstName & " " & Rst!nvcSurname
            .TextMatrix(i, 7) = Rst!NvcDescription
            .TextMatrix(i, 8) = Rst!BitTempReceived
            If Rst!BitTempReceived = True Then .Cell(flexcpBackColor, i, 0, i, .Cols - 1) = vbRed
            Rst.MoveNext
         
        Wend
    End With
    
    Set Rst = Nothing


End Sub

