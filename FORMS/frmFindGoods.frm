VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{7AEDC602-D94C-11D1-BB7A-00E0290EA3C9}#1.0#0"; "ResizeKit.ocx"
Object = "{75D4F767-8785-11D3-93AD-0000832EF44D}#2.14#0"; "FAST2001.ocx"
Begin VB.Form frmFindGoods 
   Caption         =   "         Ã” ÃÊÌ ò«·«"
   ClientHeight    =   7635
   ClientLeft      =   6060
   ClientTop       =   2445
   ClientWidth     =   9795
   BeginProperty Font 
      Name            =   "Nazanin"
      Size            =   8.25
      Charset         =   178
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmFindGoods.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   7635
   ScaleWidth      =   9795
   Begin VB.CommandButton SetBtnAscDefault 
      BackColor       =   &H000080FF&
      Caption         =   " «’·«Õ Ì Ê ﬂ ›«—”Ì œ— »«‰ﬂ «ÿ·«⁄« Ì"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2280
      TabIndex        =   27
      Top             =   120
      Width           =   2535
   End
   Begin VB.Frame Frame1 
      Height          =   6015
      Left            =   7920
      RightToLeft     =   -1  'True
      TabIndex        =   15
      Top             =   1440
      Width           =   1815
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
         TabIndex        =   26
         Tag             =   "0"
         Top             =   3600
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
         Left            =   960
         RightToLeft     =   -1  'True
         TabIndex        =   25
         Tag             =   "3"
         Top             =   1080
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
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   24
         Tag             =   "2"
         Top             =   240
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
         TabIndex        =   23
         Tag             =   "1"
         Top             =   240
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
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   22
         Tag             =   "6"
         Top             =   1920
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
         Left            =   960
         RightToLeft     =   -1  'True
         TabIndex        =   21
         Tag             =   "5"
         Top             =   1920
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
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Tag             =   "4"
         Top             =   1080
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
         Left            =   960
         RightToLeft     =   -1  'True
         TabIndex        =   19
         Tag             =   "9"
         Top             =   3600
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
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Tag             =   "8"
         Top             =   2760
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
         Left            =   960
         RightToLeft     =   -1  'True
         TabIndex        =   17
         Tag             =   "7"
         Top             =   2760
         Width           =   795
      End
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
         Height          =   915
         Index           =   10
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Top             =   4560
         Width           =   1635
      End
   End
   Begin VB.PictureBox Frame3 
      Height          =   615
      Left            =   5640
      RightToLeft     =   -1  'True
      ScaleHeight     =   555
      ScaleWidth      =   3915
      TabIndex        =   12
      Top             =   840
      Width           =   3975
      Begin VB.OptionButton Option1 
         Alignment       =   1  'Right Justify
         Caption         =   "Ã” ÃÊÌ ”—Ì⁄"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   2160
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Top             =   120
         Width           =   1575
      End
      Begin VB.OptionButton Option1 
         Alignment       =   1  'Right Justify
         Caption         =   "Ã” ÃÊÌ „⁄„Ê·Ì"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   120
         Width           =   1935
      End
   End
   Begin RESIZEKITLibCtl.ResizeKit ResizeKit1 
      Height          =   480
      Left            =   120
      OleObjectBlob   =   "frmFindGoods.frx":A4C2
      TabIndex        =   11
      Top             =   0
      Width           =   480
   End
   Begin VB.CommandButton OKButton 
      BackColor       =   &H0000FF00&
      Caption         =   "ﬁ»Ê·"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6480
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   6720
      Width           =   1215
   End
   Begin VB.CommandButton CancelButton 
      BackColor       =   &H000080FF&
      Cancel          =   -1  'True
      Caption         =   "«‰’—«›"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   6720
      Width           =   1215
   End
   Begin VB.TextBox txtName1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   7920
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   360
      Width           =   1065
   End
   Begin VB.TextBox txtBarcode 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   2280
      MultiLine       =   -1  'True
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   960
      Width           =   2505
   End
   Begin VSFlex7LCtl.VSFlexGrid vsGoods 
      Height          =   4905
      Left            =   120
      TabIndex        =   3
      Top             =   1560
      Width           =   7755
      _cx             =   13679
      _cy             =   8652
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
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
      Cols            =   8
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   500
      RowHeightMax    =   500
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmFindGoods.frx":A548
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
   Begin VB.TextBox txtName3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   5760
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   360
      Width           =   1065
   End
   Begin VB.TextBox txtName2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   6840
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   360
      Width           =   1065
   End
   Begin FLWCtrls.FWProgressBar FWProgressBar1 
      Height          =   375
      Left            =   1440
      Top             =   7080
      Width           =   4935
      _ExtentX        =   8705
      _ExtentY        =   661
      BackColor       =   -2147483626
      BorderStyle     =   10
   End
   Begin VB.Label LblCount 
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
      Height          =   495
      Left            =   2400
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   6480
      Width           =   3135
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
      Height          =   1455
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   0
      Width           =   2055
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "»«—òœ"
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
      Left            =   4920
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   960
      Width           =   465
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "‰«„ ò«·«"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   435
      Left            =   9000
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   360
      Width           =   585
   End
End
Attribute VB_Name = "frmFindGoods"
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

Private Sub BtnKeypad_Click(index As Integer)
    If BtnKeypad(index).Tag = "" Then
        If LCase(VarActForm) = "frminvoice" Then
            If Len(Trim(frmInvoice.lblNum.Caption)) >= 1 Then
                frmInvoice.lblNum.Caption = left(frmInvoice.lblNum.Caption, Len(Trim(frmInvoice.lblNum.Caption)) - 1)
            End If
        End If
    Else
        If LCase(VarActForm) = "frminvoice" Then
            frmInvoice.lblNum.Caption = frmInvoice.lblNum.Caption + BtnKeypad(index).Tag
        End If
    End If

End Sub

Private Sub CancelButton_Click()
    If LCase(VarActForm) = "frminvoice" Then
        frmInvoice.lblNum.Caption = ""
    End If
    mvarcode = 0
    mvarName = ""
    'Me.Hide
    Unload Me
End Sub

Private Sub Form_Activate()
    Dim hMenu As Long
    On Error Resume Next
    
    hMenu = GetSystemMenu(Me.hwnd, False)

    DeleteMenu hMenu, 6, MF_BYPOSITION
    
    CancelButton.Visible = False
    OKButton.Visible = False
    
    Option1(0).Value = clsStation.GoodSearchDefault
    Option1(1).Value = Not (clsStation.GoodSearchDefault)
    txtName1.SetFocus
    
    CancelButton.Visible = True
    OKButton.Visible = True
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    NumberCode = KeyCode
''''    If KeyCode > 48 And KeyCode < 58 Then
''''        If LCase(VarActForm) = "frminvoice" Then
''''            frmInvoice.lblNum.Caption = frmInvoice.lblNum.Caption & Chr(KeyCode)
''''        ElseIf LCase(VarActForm) = "frminvoice" Then
''''            frminvoice.lblNum.Caption = frminvoice.lblNum.Caption & Chr(KeyCode)
''''        End If
''''    End If
End Sub

Private Sub Form_Load()
    vsGoods.Rows = 8
    RightTop Me
    If Screen.Width > 12000 Then
''''        frmFindGoods.Left = 7400
''''        frmFindGoods.Top = 1000
    Else
        frmFindGoods.left = 4000
        frmFindGoods.top = 1000
    End If
    
    If LCase(VarActForm) = "frminvoice" Then
        NotSupportedGoodType = EnumGoodType.forBuy
        If clsStation.MultiPrice = False Then
           vsGoods.ColHidden(5) = True
           vsGoods.ColHidden(6) = True
           vsGoods.ColHidden(7) = False
        End If
    ElseIf LCase(VarActForm) = "frmpurchase" Or LCase(VarActForm) = "frmgoodturnover" Or LCase(VarActForm) = "frmusepercent" Then
        NotSupportedGoodType = EnumGoodType.forSale
        vsGoods.ColHidden(5) = True
        vsGoods.ColHidden(6) = True
        vsGoods.TextMatrix(0, 4) = "  ›Ì Œ—Ìœ  "
        vsGoods.TextMatrix(0, 7) = "  ›Ì ›—Ê‘ "
    Else
        NotSupportedGoodType = EnumGoodType.All
        vsGoods.ColHidden(5) = False
        vsGoods.ColHidden(6) = False
        vsGoods.TextMatrix(0, 4) = "  ›Ì Œ—Ìœ  "
        vsGoods.TextMatrix(0, 7) = "  ›Ì ›—Ê‘ "
    End If
    
    mvarcode = 0
    mvarName = ""
    

    vsGoods.Row = 0
    vsGoods.ColWidth(0) = 700
    vsGoods.ColWidth(2) = 2700
    vsGoods.ColWidth(3) = 1850
    
    FlexGridActive
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
    If PosConnection.State = adStateOpen Then PosConnection.Close: Set PosConnection = Nothing

    SaveSetting strMainKey, Me.Name, "Left", Me.left
    SaveSetting strMainKey, Me.Name, "Top", Me.top


End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call PresetScreenSaver
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call PresetScreenSaver
End Sub

Private Sub OKButton_Click()
    If vsGoods.Row > 0 Then
        mvarcode = vsGoods.TextMatrix(vsGoods.Row, 1)
        mvarName = vsGoods.TextMatrix(vsGoods.Row, 2)
        mvarBarcodeName = vsGoods.TextMatrix(vsGoods.Row, 3)
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
    txtName3.Text = ""
    txtName2.Text = ""
    txtName3.Visible = False
    txtName2.Visible = False
    txtName1.Text = ""
    txtName1.SetFocus
    If LCase(VarActForm) <> "frminvoice" And LCase(VarActForm) <> "frmpurchase" Then
       'Me.Hide
       Unload Me
    End If
End Sub

Private Sub ResizeKit1_ExitResize(ByVal XScale As Double, ByVal YScale As Double)

    If formloadFlag = True Then
        SaveSetting strMainKey, Me.Name, "Height", Me.Height
        SaveSetting strMainKey, Me.Name, "Width", Me.Width
    End If

End Sub

Private Sub SetBtnAscDefault_Click()
On Error GoTo ErrHandler
    RunNonParametricStoredProcedure "Update_Good_btnAscDefault"
    ShowMessage "  «’·«Õ ﬂ«—«ﬂ —Â«Ì ›«—”Ì «‰Ã«„ ‘œ ", True, False, " «ÌÌœ", ""
Exit Sub
ErrHandler:
    ShowDisMessage err.Description, 1500
End Sub

Private Sub txtBarcode_Change()

    i = vsGoods.FindRow(txtBarcode.Text, 1, 3, True, True)
    If i > 0 Then
        vsGoods.Row = i
        vsGoods.ShowCell i, 0
        LblFindGoods.Caption = ""
    Else
        vsGoods.Row = 0
        vsGoods.ShowCell 0, 0
        If txtBarcode.Text <> "" Then
           LblFindGoods.Caption = " »«—ﬂœ ﬂ«·« »« - " & txtBarcode.Text & " - œ— ”Ì” „  ⁄—Ì› ‰‘œÂ"
         Else
            LblFindGoods.Caption = "  »«—ﬂœ ﬂ«·« —« Ê«—œ ﬂ‰Ìœ "
         End If
    End If

End Sub

Private Sub txtBarcode_GotFocus()

    vsGoods.Select vsGoods.Row, 3
    vsGoods.Sort = flexSortGenericAscending
    LblFindGoods.Caption = " »«—ﬂœ —« —ÊÌ ﬂ«·« »ﬂ‘Ìœ "

End Sub

Private Sub txtBarcode_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Select Case Shift
        Case 0
            Select Case KeyCode
                Case 40
                    vsGoods.SetFocus
                    KeyCode = 0
                    SendKeys "{DOWN}", True
                    txtBarcode.SetFocus
                Case 38
                    
                    vsGoods.SetFocus
                    KeyCode = 0
                    SendKeys "{UP}", True
                    txtBarcode.SetFocus
            End Select
    
    End Select

End Sub

Private Sub txtName1_Change()
''''   If NumberCode > 48 And NumberCode < 58 Then
''''        If Len(txtName1.Text) = 0 Then Exit Sub
''''        txtName1.Text = Right(txtName1.Text, Len(txtName1.Text) - 1)
''''        Exit Sub
''''   End If
   If Option1(0).Value = False Then     'Normal Search
        i = vsGoods.FindRow(txtName1.Text, 1, 2, False, False)
        If i > 0 Then
            vsGoods.Row = i
            vsGoods.ShowCell i, 0
            LblFindGoods.Caption = ""
            txtName2.Text = ""
            If txtName1.Text <> "" And clsStation.ThreeSegmentSearch = True Then
                txtName2.Visible = True
            End If
        Else
            vsGoods.Row = 0
            vsGoods.ShowCell 0, 0
            If txtName1.Text <> "" Then
               LblFindGoods.Caption = " ‰«„ ﬂ«·« »« - " & txtName1.Text & " - œ— ”Ì” „  ⁄—Ì› ‰‘œÂ"
             Else
                LblFindGoods.Caption = " Õ—Ê› «Ê· ‰«„ ﬂ«·« —« Ê«—œ ﬂ‰Ìœ "
             End If
        End If
   Else                                 'Fast Serach
       If Len(txtName1.Text) > 0 Then
            SearchType = 1
            Define_Good
            If vsGoods.Rows > 1 Then
                vsGoods.Row = 1
                vsGoods.ShowCell 1, 0
                LblFindGoods.Caption = ""
                txtName2.Text = ""
                If txtName1.Text <> "" And clsStation.ThreeSegmentSearch = True Then
                    txtName2.Visible = True
                    txtName2.SetFocus
                End If
            Else
                vsGoods.Row = 0
                vsGoods.ShowCell 0, 0
                If Len(txtName1.Text) > 0 Then
                    LblFindGoods.Caption = " ‰«„ ﬂ«·« »« - " & txtName1.Text & " - œ— ”Ì” „  ⁄—Ì› ‰‘œÂ"
                 Else
                    LblFindGoods.Caption = " Õ—Ê› «Ê· ‰«„ ﬂ«·« —« Ê«—œ ﬂ‰Ìœ "
                 End If
            
            End If
       Else
          vsGoods.Rows = 1
        
       End If
  
   End If

End Sub

Private Sub txtName1_GotFocus()
    txtName2.Visible = False
    txtName3.Visible = False
    txtName2.Text = ""
    txtName3.Text = ""
    vsGoods.Select vsGoods.Row, 2
    vsGoods.Sort = flexSortGenericAscending
    If Option1(0).Value = False Then txtName1_Change
End Sub

Private Sub txtName2_Change()
   Dim Txt2 As String
    If Option1(0).Value = False Then
        Txt2 = TmpGoodName1 & txtName2.Text
        i = vsGoods.FindRow(Txt2, 1, 2, False, False)
        If i > 0 Then
            LblFindGoods.Caption = ""
            txtName3.Text = ""
            vsGoods.Row = i
            vsGoods.ShowCell i, 0
            If txtName2.Text <> "" And clsStation.ThreeSegmentSearch = True Then
                txtName3.Visible = True
            End If
        Else
            vsGoods.Row = 0
            vsGoods.ShowCell 0, 0
            LblFindGoods.Caption = " ‰«„ ﬂ«·« »« - " & Txt2 & " - œ— ”Ì” „  ⁄—Ì› ‰‘œÂ"
        End If
   Else
        SearchType = 2
        Define_Good
        If vsGoods.Rows > 1 Then
            vsGoods.Row = 1
            vsGoods.ShowCell 1, 0
            LblFindGoods.Caption = ""
            txtName3.Text = ""
            If txtName2.Text <> "" And clsStation.ThreeSegmentSearch = True Then
                txtName3.Visible = True
                txtName3.SetFocus
            End If
        Else
            vsGoods.Row = 0
            vsGoods.ShowCell 0, 0
            If Len(txtName2.Text) > 0 Then
                LblFindGoods.Caption = " ‰«„ ﬂ«·« »« - " & txtName2.Text & " - œ— ”Ì” „  ⁄—Ì› ‰‘œÂ"
             Else
                LblFindGoods.Caption = " Õ—Ê› «Ê· ‰«„ ﬂ«·« —« Ê«—œ ﬂ‰Ìœ "
             End If
        
        End If
  
   End If

End Sub

Private Sub txtName2_GotFocus()
    txtName3.Visible = False
    txtName3.Text = ""

    vsGoods.Select vsGoods.Row, 2
    vsGoods.Sort = flexSortGenericAscending
    
    If vsGoods.Row > 0 And Option1(0).Value = False Then
        j = InStr(1, vsGoods.TextMatrix(i, 2), " ", vbTextCompare)
        TmpGoodName1 = left(vsGoods.TextMatrix(i, 2), j)
        txtName2_Change
    End If
End Sub


Private Sub txtName3_Change()

   Dim Txt3 As String
    If Option1(0).Value = False Then
        Txt3 = TmpGoodName2 & txtName3.Text
        i = vsGoods.FindRow(Txt3, 1, 2, False, False)
        If i > 0 Then
            LblFindGoods.Caption = ""
            vsGoods.Row = i
            vsGoods.ShowCell i, 0
        Else
            vsGoods.Row = 0
            vsGoods.ShowCell 0, 0
            LblFindGoods.Caption = " ‰«„ ﬂ«·« »« - " & Txt3 & " - œ— ”Ì” „  ⁄—Ì› ‰‘œÂ"
        End If
   Else
        SearchType = 3
        Define_Good
        If vsGoods.Rows > 1 Then
            vsGoods.Row = 1
            vsGoods.ShowCell 1, 0
            LblFindGoods.Caption = ""
        Else
            vsGoods.Row = 0
            vsGoods.ShowCell 0, 0
            If Len(txtName3.Text) > 0 Then
                LblFindGoods.Caption = " ‰«„ ﬂ«·« »« - " & txtName2.Text & " - œ— ”Ì” „  ⁄—Ì› ‰‘œÂ"
             Else
                LblFindGoods.Caption = " Õ—Ê› «Ê· ‰«„ ﬂ«·« —« Ê«—œ ﬂ‰Ìœ "
             End If
        
        End If
  
   End If

End Sub

Private Sub txtName3_GotFocus()
    If vsGoods.Row > 0 And Option1(0).Value = False Then
        vsGoods.Select vsGoods.Row, 2
        vsGoods.Sort = flexSortGenericAscending
        j = InStr(j + 1, vsGoods.TextMatrix(i, 2), " ", vbTextCompare)
        TmpGoodName2 = left(vsGoods.TextMatrix(i, 2), j)
        txtName3_Change
    End If
End Sub

Private Sub Option1_Click(index As Integer)
    If Option1(0).Value = False Then
        txtName2.Visible = False
        txtName3.Visible = False
        txtName2.Text = ""
        txtName3.Text = ""
        FillvsGoods
    Else
        vsGoods.Rows = 1
        txtName2.Visible = False
        txtName3.Visible = False
        txtName2.Text = ""
        txtName3.Text = ""
    End If
    On Error Resume Next
    txtName1.SetFocus
End Sub

Private Sub FillvsGoods()
    Dim Rst As New ADODB.Recordset

    If LCase(VarActForm) = "frminvoice" Then
        ReDim Parameter(4) As Parameter
        Parameter(0) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
        Parameter(1) = GenerateInputParameter("@NotSupportedGoodType", adInteger, 4, NotSupportedGoodType)
        Parameter(2) = GenerateInputParameter("@StationId", adInteger, 4, clsArya.StationNo)
        Parameter(3) = GenerateInputParameter("@Flag", adBoolean, 1, 0)
        Parameter(4) = GenerateInputParameter("@AccountYear", adSmallInt, 2, AccountYear)
        Set Rst = RunParametricStoredProcedure2Rec("Get_GoodInfo_By_GoodType", Parameter)
    Else  'If LCase(VarActForm) = "frmpurchase" Or LCase(VarActForm) = "frmgoodturnover" Or LCase(VarActForm) = "frmgood" Or LCase(VarActForm) = "frmusepercent" Then
        ReDim Parameter(4) As Parameter
        Parameter(0) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
        Parameter(1) = GenerateInputParameter("@NotSupportedGoodType", adInteger, 4, NotSupportedGoodType)
        Parameter(2) = GenerateInputParameter("@StationId", adInteger, 4, clsArya.StationNo)
        Parameter(3) = GenerateInputParameter("@Flag", adBoolean, 1, 1)
        Parameter(4) = GenerateInputParameter("@AccountYear", adSmallInt, 2, AccountYear)
        Set Rst = RunParametricStoredProcedure2Rec("Get_GoodInfo_By_GoodType", Parameter)
    End If
    
    With vsGoods
        .Rows = 1
        i = 0
        FWProgressBar1.Value = 0
        MousePointer = 11
        While Rst.EOF <> True
            .Rows = .Rows + 1
            i = .Rows - 1
            .TextMatrix(i, 0) = i
            .TextMatrix(i, 1) = Rst!Code
            .TextMatrix(i, 2) = Rst![GoodName]
            .TextMatrix(i, 3) = IIf(IsNull(Rst!barcode), "", Rst!barcode)
            If NotSupportedGoodType = EnumGoodType.forBuy Then
               .TextMatrix(i, 4) = Rst!SellPrice
            Else
               .TextMatrix(i, 4) = Rst!BuyPrice
            End If
            .TextMatrix(i, 5) = Rst!Sellprice2
            .TextMatrix(i, 6) = Rst!Sellprice3
            .TextMatrix(i, 7) = Rst!SellPrice
            
            If i Mod 1000 = 0 Then DoEvents
            LblCount.Caption = " ⁄œ«œ —ﬂÊ—œÂ«   -  " & i
            If i Mod 100 = 0 Then
                FWProgressBar1.Value = FWProgressBar1 + 1
                If FWProgressBar1.Value = 101 Then
                    FWProgressBar1.Value = 1
                End If
            End If
            Rst.MoveNext
        Wend
        MousePointer = 0
    End With
   
    Set Rst = Nothing


End Sub


Private Sub txtName1_KeyDown(KeyCode As Integer, Shift As Integer)
    If (txtName1.Text = "" Or Len(txtName1.Text) = 1) And KeyCode = 8 Then
        txtName2.Visible = False
    End If
    If KeyCode = 40 Or KeyCode = 38 Then
        vsGoods.SetFocus
    End If
End Sub
Private Sub txtName2_KeyDown(KeyCode As Integer, Shift As Integer)
   If txtName2.Text = "" And KeyCode = 8 Then
       txtName1.SetFocus
   End If
   If Len(txtName2.Text) = 1 And KeyCode = 8 Then
       txtName3.Visible = False
   End If
    If KeyCode = 40 Or KeyCode = 38 Then
        vsGoods.SetFocus
    End If
End Sub

Private Sub txtName3_KeyDown(KeyCode As Integer, Shift As Integer)
   If txtName3.Text = "" And KeyCode = 8 Then
       txtName2.SetFocus
   End If
    If KeyCode = 40 Or KeyCode = 38 Then
        vsGoods.SetFocus
    End If

End Sub

Private Sub vsGoods_AfterSort(ByVal Col As Long, Order As Integer)
    For i = 1 To vsGoods.Rows - 1
        vsGoods.TextMatrix(i, 0) = i
    Next
    
End Sub
Private Sub vsGoods_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    For i = 0 To vsGoods.Cols - 1
        SaveSetting strMainKey, "FindGoods", "Col" & i, vsGoods.ColWidth(i)
    Next

End Sub


Private Sub vsGoods_DblClick()
    If vsGoods.Row > 0 Then
        OKButton_Click
    End If
End Sub

Private Sub Define_Good()
    
    Dim Rst As New ADODB.Recordset
    
    Select Case SearchType
        Case 1
            ReDim Parameter(1) As Parameter
            Parameter(0) = GenerateInputParameter("@Name1", adVarWChar, 20, Trim(left(txtName1.Text, 20)))
            Parameter(1) = GenerateInputParameter("@NotSupportedGoodType", adInteger, 4, NotSupportedGoodType)
            Set Rst = RunParametricStoredProcedure2Rec("Get_Good_Name", Parameter)
        Case 2
            ReDim Parameter(2) As Parameter
            Parameter(0) = GenerateInputParameter("@Name1", adVarWChar, 20, Trim(left(txtName1.Text, 20)))
            Parameter(1) = GenerateInputParameter("@Name2", adVarWChar, 20, left(" " & txtName2.Text, 20))
            Parameter(2) = GenerateInputParameter("@NotSupportedGoodType", adInteger, 4, NotSupportedGoodType)
            Set Rst = RunParametricStoredProcedure2Rec("Get_Good_Name2", Parameter)
        Case 3
            ReDim Parameter(3) As Parameter
            Parameter(0) = GenerateInputParameter("@Name1", adVarWChar, 20, Trim(left(txtName1.Text, 20)))
            Parameter(1) = GenerateInputParameter("@Name2", adVarWChar, 20, left(" " & txtName2.Text, 20))
            Parameter(2) = GenerateInputParameter("@Name2", adVarWChar, 20, left(" " & txtName3.Text, 20))
            Parameter(3) = GenerateInputParameter("@NotSupportedGoodType", adInteger, 4, NotSupportedGoodType)
            Set Rst = RunParametricStoredProcedure2Rec("Get_Good_Name3", Parameter)
    End Select
    Dim TmpTel As String
    Dim jj As Integer
    
    With vsGoods
        .Rows = 1
        i = 0
        While Rst.EOF <> True
            .Rows = .Rows + 1
            i = .Rows - 1
            .TextMatrix(i, 0) = i
            .TextMatrix(i, 1) = Rst!Code
            .TextMatrix(i, 2) = Rst![Name]
            .TextMatrix(i, 3) = Rst!barcode
            If NotSupportedGoodType = EnumGoodType.forBuy Then
               .TextMatrix(i, 4) = Rst!SellPrice
            Else
               .TextMatrix(i, 4) = Rst!BuyPrice
            End If
            .TextMatrix(i, 5) = Rst!Sellprice2
            .TextMatrix(i, 6) = Rst!Sellprice3
            .TextMatrix(i, 7) = Rst!SellPrice
            Rst.MoveNext
        Wend
    End With

''''        If i > 0 Then
''''            LblCount.Caption = " ⁄œ«œ —ﬂÊ—œÂ«   -  " & i
''''        Else
''''            LblCount.Caption = ""
''''        End If

    Set Rst = Nothing

End Sub
Private Sub FlexGridActive()

    With vsGoods
             
         For i = 0 To .Cols - 1
             .ColWidth(i) = Val(GetSetting(strMainKey, "FindGoods", "Col" & i))
         Next i
'''    vsGoods.ColWidth(0) = 700
'''    vsGoods.ColWidth(2) = 2700
'''    vsGoods.ColWidth(3) = 1850
        
        If .ColWidth(0) = 0 Then
            .ColWidth(0) = 700       '
        End If
        If .ColWidth(2) = 0 Then
            .ColWidth(2) = 2700        '
        End If
        If .ColWidth(3) = 0 Then
            .ColWidth(3) = 1850      '
        End If
        If .ColWidth(4) = 0 Then
            .ColWidth(4) = .Width / 6
        End If
        If .ColWidth(5) = 0 Then
            .ColWidth(5) = .Width / 8        '
        End If
        If .ColWidth(6) = 0 Then
            .ColWidth(6) = .Width / 7       '
        End If
        If .ColWidth(7) <= 20 Then
            .ColWidth(7) = .Width / 12       '
        End If
       
       
   End With
End Sub

