VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Begin VB.Form frmFindGoods_Kb 
   BackColor       =   &H80000016&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "                                           Ã” ÃÊÌ ﬂ«·«Â«Ì ﬂÌ»Ê—œ «Ì‰ «Ì” ê«Â"
   ClientHeight    =   6285
   ClientLeft      =   6045
   ClientTop       =   2430
   ClientWidth     =   8070
   Icon            =   "frmFindGoods_Kb.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   6285
   ScaleWidth      =   8070
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   1935
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   120
      Width           =   5295
      Begin VB.CommandButton BtnKeypad 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000D&
         Caption         =   "Å«ﬂ ﬂ‰"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1635
         Index           =   10
         Left            =   4320
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   120
         Width           =   840
      End
      Begin VB.CommandButton BtnKeypad 
         BackColor       =   &H8000000D&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "Titr"
            Size            =   15.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Index           =   0
         Left            =   120
         TabIndex        =   14
         Tag             =   "0"
         Top             =   120
         Width           =   795
      End
      Begin VB.CommandButton BtnKeypad 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000D&
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Index           =   3
         Left            =   2640
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Tag             =   "3"
         Top             =   120
         Width           =   795
      End
      Begin VB.CommandButton BtnKeypad 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000D&
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Index           =   2
         Left            =   1800
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Tag             =   "2"
         Top             =   120
         Width           =   795
      End
      Begin VB.CommandButton BtnKeypad 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000D&
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Index           =   1
         Left            =   960
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Tag             =   "1"
         Top             =   120
         Width           =   795
      End
      Begin VB.CommandButton BtnKeypad 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000D&
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Index           =   6
         Left            =   960
         RightToLeft     =   -1  'True
         TabIndex        =   10
         Tag             =   "6"
         Top             =   960
         Width           =   795
      End
      Begin VB.CommandButton BtnKeypad 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000D&
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Index           =   5
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Tag             =   "5"
         Top             =   960
         Width           =   795
      End
      Begin VB.CommandButton BtnKeypad 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000D&
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Index           =   4
         Left            =   3480
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Tag             =   "4"
         Top             =   120
         Width           =   795
      End
      Begin VB.CommandButton BtnKeypad 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000D&
         Caption         =   "9"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Index           =   9
         Left            =   3480
         RightToLeft     =   -1  'True
         TabIndex        =   7
         Tag             =   "9"
         Top             =   960
         Width           =   795
      End
      Begin VB.CommandButton BtnKeypad 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000D&
         Caption         =   "8"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Index           =   8
         Left            =   2640
         RightToLeft     =   -1  'True
         TabIndex        =   6
         Tag             =   "8"
         Top             =   960
         Width           =   795
      End
      Begin VB.CommandButton BtnKeypad 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000D&
         Caption         =   "7"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   15.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   795
         Index           =   7
         Left            =   1800
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Tag             =   "7"
         Top             =   960
         Width           =   795
      End
   End
   Begin VB.TextBox txtRow 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   5520
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   480
      Width           =   2505
   End
   Begin VSFlex7LCtl.VSFlexGrid vsGoods 
      Height          =   3945
      Left            =   120
      TabIndex        =   0
      Top             =   2160
      Width           =   7875
      _cx             =   13891
      _cy             =   6959
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   178
         Weight          =   400
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
      AllowUserResizing=   0
      SelectionMode   =   3
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   7
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   500
      RowHeightMax    =   600
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmFindGoods_Kb.frx":A4C2
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
   Begin VB.Label LblFindGoods 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   735
      Left            =   5640
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   1200
      Width           =   2295
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "—œÌ› ò«·«"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   6840
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   0
      Width           =   1185
   End
End
Attribute VB_Name = "frmFindGoods_Kb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim NumberCode As Integer
Dim i, TmpLineCode As Long
Dim TmpGoodName1, TmpGoodName2 As String
Dim j As Integer
Dim Parameter() As Parameter
Dim NotSupportedGoodType As EnumGoodType
Dim SearchType As Integer
Dim MvarUserDefine As Boolean
Dim mvarKeyCode, MvarShiftKey, KeyAscii As Integer
Public Function SendVariables(ByRef UserDefine, ByRef KeyCode, ByRef Shift, ByRef AsciiCode)
    MvarUserDefine = UserDefine
    mvarKeyCode = KeyCode
    MvarShiftKey = Shift
    KeyAscii = AsciiCode
End Function

Private Sub CancelButton_Click()
    If LCase(VarActForm) = "frminvoice" Then
        frmInvoice.lblNum.Caption = ""
    End If
    mvarcode = 0
    Unload Me
End Sub

Private Sub BtnKeypad_Click(index As Integer)
    If BtnKeypad(index).Tag = "" Then
        If LCase(VarActForm) = "frminvoice" Then
            If Len(Trim(frmInvoice.lblNum.Caption)) >= 1 Then
                frmInvoice.lblNum.Caption = Left(frmInvoice.lblNum.Caption, Len(Trim(frmInvoice.lblNum.Caption)) - 1)
            End If
        End If
    Else
        If LCase(VarActForm) = "frminvoice" Then
            frmInvoice.lblNum.Caption = frmInvoice.lblNum.Caption + BtnKeypad(index).Tag
        End If
    End If

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call PresetScreenSaver
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call PresetScreenSaver
End Sub

Private Sub Form_Activate()
    TxtRow.SetFocus

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then
        Unload Me
    ElseIf KeyCode = 13 Then
        OKButton_Click
    ElseIf KeyCode = 8 Then
        TxtRow.SetFocus
    End If
    NumberCode = KeyCode
End Sub

Private Sub Form_Load()

   ' CenterCenterinSecondScreen Me
    If Screen.Width > 12000 Then
        frmFindGoods_Kb.Left = 7400
        frmFindGoods_Kb.Top = 200
    Else
        frmFindGoods_Kb.Left = 4000
        frmFindGoods_Kb.Top = 200
    End If
    
    If LCase(VarActForm) = "frminvoice" Then
        NotSupportedGoodType = EnumGoodType.forBuy
        If clsStation.MultiPrice = False Then
           vsGoods.ColHidden(4) = True
           vsGoods.ColHidden(5) = True
        End If
        vsGoods.ColHidden(6) = True
    Else
        NotSupportedGoodType = EnumGoodType.forSale
        vsGoods.ColHidden(4) = True
        vsGoods.ColHidden(5) = True
    End If
    If clsStation.AlphabetGoodSearch = True Then
            Label2.Caption = "—œÌ› ﬂ«·«"
    Else
            Label2.Caption = "‰«„ ﬂ«·«"
    End If
    mvarcode = 0
     
    FillvsGoods
    
    vsGoods.Row = 1
    formloadFlag = False
    Me.Left = Val(GetSetting(strMainKey, Me.Name, "Left"))
''''    If Val(GetSetting(strMainKey,Me.Name, "Height")) > 5000Then
''''        Me.Height = Val(GetSetting(strMainKey, Me.Name, "Height"))
''''    End If
''''    If Val(GetSetting(strMainKey, Me.Name, "Width")) > 5000 Then
''''        Me.Width = Val(GetSetting(strMainKey, Me.Name, "Width"))
''''    End If
    Me.Top = Val(GetSetting(strMainKey, Me.Name, "Top"))
    If Me.Left < 0 Then Me.Left = 0
    If Me.Top < 0 Then Me.Top = 0
    If Me.Top > Me.ScaleHeight Then Me.Top = 0

    formloadFlag = True


    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If PosConnection.State = adStateOpen Then PosConnection.Close: Set PosConnection = Nothing
    
    SaveSetting strMainKey, Me.Name, "Left", Me.Left
    SaveSetting strMainKey, Me.Name, "Top", Me.Top

End Sub

Private Sub OKButton_Click()
    If vsGoods.Row > 0 Then
        mvarcode = vsGoods.TextMatrix(vsGoods.Row, 1)
        If LCase(VarActForm) = "frminvoice" Then
            If frmInvoice.GetGoodCode(mvarcode) = True Then
                frmInvoice.ChangeGoodquantity
                frmInvoice.lblNum.Caption = ""
            End If
        ElseIf LCase(VarActForm) = "frmpurchase" Then
            If frmPurchase.GetGoodCode(mvarcode) = True Then
                frmPurchase.ChangeGoodquantity
                frmPurchase.lblNum.Caption = ""
            End If
        End If
    Else
        mvarcode = 0
    End If
    Unload Me

End Sub

Private Sub Label5_Click()

End Sub

Private Sub LblFindGoods_Click()

    If formloadFlag = True Then
        SaveSetting strMainKey, Me.Name, "Height", Me.Height
        SaveSetting strMainKey, Me.Name, "Width", Me.Width
    End If

End Sub


Private Sub txtRow_Change()
'   If (NumberCode > 48 And NumberCode < 58) Or (NumberCode > 96 And NumberCode < 105) Then
      If clsStation.AlphabetGoodSearch = True Then
            i = vsGoods.FindRow(TxtRow.Text, 1, 0, False, False)
            If i > 0 Then
                vsGoods.Row = i
                vsGoods.ShowCell i, 0
            Else
                vsGoods.Row = 0
                vsGoods.ShowCell 0, 0
                If TxtRow.Text <> "" Then
                   LblFindGoods.Caption = " —œÌ› ﬂ«·« »« - " & TxtRow.Text & " - ÊÃÊœ ‰œ«—œ"
                 Else
                    LblFindGoods.Caption = " ‘„«—Â —œÌ› ﬂ«·« —« Ê«—œ ﬂ‰Ìœ "
                 End If
            End If
       Else
              
            i = vsGoods.FindRow(TxtRow.Text, 1, 2, False, False)
            If i > 0 Then
                vsGoods.Row = i
                vsGoods.ShowCell i, 0
            Else
                vsGoods.Row = 0
                vsGoods.ShowCell 0, 0
                If TxtRow.Text <> "" Then
                   LblFindGoods.Caption = " ‰«„ ﬂ«·« »« - " & TxtRow.Text & " - ÊÃÊœ ‰œ«—œ"
                 Else
                    LblFindGoods.Caption = " ‰«„ ﬂ«·« —« Ê«—œ ﬂ‰Ìœ "
                 End If
            End If
       End If
'   End If
   
End Sub

Private Sub txtRow_GotFocus()
    vsGoods.Select vsGoods.Row, 2
    vsGoods.Sort = flexSortGenericAscending
End Sub

Private Sub FillvsGoods()

    Dim Rst2 As New ADODB.Recordset
    If MvarUserDefine Then
         ReDim Parameter(4) As Parameter
         Parameter(0) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
         Parameter(1) = GenerateInputParameter("@StationId", adInteger, 4, clsArya.StationNo)
         Parameter(2) = GenerateInputParameter("@KeyCode", adInteger, 4, mvarKeyCode)
         Parameter(3) = GenerateInputParameter("@ShiftKey", adInteger, 4, MvarShiftKey)
         Parameter(4) = GenerateInputParameter("@NotSupportedGoodType", adInteger, 4, NotSupportedGoodType)
         Set Rst2 = RunParametricStoredProcedure2Rec("Get_Good_KB", Parameter)
    Else
         ReDim Parameter(2) As Parameter
         Parameter(0) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
         Parameter(1) = GenerateInputParameter("@BtnAscDefault", adInteger, 4, KeyAscii)
         Parameter(2) = GenerateInputParameter("@NotSupportedGoodType", adInteger, 4, NotSupportedGoodType)
         Set Rst2 = RunParametricStoredProcedure2Rec("Get_Good_DefaultKB", Parameter)
    End If
    
    
    With vsGoods
        .Rows = 1
        i = 0
        While Rst2.EOF <> True
            .Rows = .Rows + 1
            i = .Rows - 1
            .TextMatrix(i, 0) = i
            .TextMatrix(i, 1) = Rst2!Code
            .TextMatrix(i, 2) = Rst2![Name]
            .TextMatrix(i, 3) = Rst2!SellPrice
            .TextMatrix(i, 4) = Rst2!Sellprice2
            .TextMatrix(i, 5) = Rst2!Sellprice3
            .TextMatrix(i, 6) = Rst2!BuyPrice
            Rst2.MoveNext
        Wend
    End With
   
    Set Rst2 = Nothing
    If Screen.Width > 12000 Then
        If i > 25 Then i = 25
    Else
        If i > 18 Then i = 18
    End If
    frmFindGoods_Kb.Height = 3400 + i * 500
    vsGoods.Height = 1200 + i * 500
End Sub

Private Sub txtRow_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then
        Unload Me
    End If
    If vsGoods.Row = 0 Then
        If KeyCode = 40 Then
            SendKeys "{Tab}", True
            vsGoods.Row = 1
            vsGoods.ShowCell 1, 0
        End If
    Else
        If KeyCode = 40 Or KeyCode = 38 Then
            vsGoods.SetFocus
        End If
    End If
End Sub

Private Sub vsGoods_AfterSort(ByVal Col As Long, Order As Integer)
    For i = 1 To vsGoods.Rows - 1
        vsGoods.TextMatrix(i, 0) = i
    Next
    
End Sub

Private Sub vsGoods_DblClick()
    If vsGoods.Row > 0 Then
        OKButton_Click
    End If
End Sub


