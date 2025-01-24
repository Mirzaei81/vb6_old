VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{7AEDC602-D94C-11D1-BB7A-00E0290EA3C9}#1.0#0"; "ResizeKit.ocx"
Begin VB.Form frmFindSupplier 
   BackColor       =   &H80000016&
   Caption         =   $"frmFindSupplier.frx":0000
   ClientHeight    =   8625
   ClientLeft      =   2775
   ClientTop       =   3765
   ClientWidth     =   10680
   Icon            =   "frmFindSupplier.frx":00BF
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   8625
   ScaleWidth      =   10680
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   0
      RightToLeft     =   -1  'True
      TabIndex        =   22
      Top             =   600
      Width           =   3015
      Begin VB.CommandButton cmdTurnOver 
         Caption         =   "ê—œ‘ Õ”«» «Ì‰  «„Ì‰ ò‰‰œÂ"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   23
         Top             =   960
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.Label LblTotalCreditDebit 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "B Nazanin"
            Size            =   14.25
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   450
         Left            =   120
         RightToLeft     =   -1  'True
         TabIndex        =   25
         Top             =   480
         Width           =   2775
      End
      Begin VB.Label LblTotalCreditDebitLabel 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "»œÂÌ- ÿ·» ﬂ·:"
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   9
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   1440
         RightToLeft     =   -1  'True
         TabIndex        =   24
         Top             =   120
         Width           =   1335
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   615
      Left            =   3120
      RightToLeft     =   -1  'True
      ScaleHeight     =   555
      ScaleWidth      =   3435
      TabIndex        =   19
      Top             =   1560
      Width           =   3495
      Begin VB.OptionButton Option1 
         Alignment       =   1  'Right Justify
         Caption         =   "Ã” ÃÊÌ ”—Ì⁄"
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   1680
         RightToLeft     =   -1  'True
         TabIndex        =   21
         Top             =   120
         Value           =   -1  'True
         Width           =   1575
      End
      Begin VB.OptionButton Option1 
         Alignment       =   1  'Right Justify
         Caption         =   "Ã” ÃÊÌ „⁄„Ê·Ì"
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   -120
         RightToLeft     =   -1  'True
         TabIndex        =   20
         Top             =   120
         Width           =   1815
      End
   End
   Begin VB.Frame Frame5 
      Height          =   855
      Left            =   3120
      RightToLeft     =   -1  'True
      TabIndex        =   13
      Top             =   720
      Width           =   3495
      Begin VB.TextBox TxtTimer 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   1080
         RightToLeft     =   -1  'True
         TabIndex        =   14
         Text            =   "500"
         Top             =   120
         Width           =   615
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "„Ì·Ì À«‰ÌÂ"
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
         TabIndex        =   16
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "“„«‰  «ŒÌ— »Ì‰ ﬂ·ÌœÂ«Ì Ê—ÊœÌ"
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   9.75
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   1800
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   120
         Width           =   1215
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   0
      Top             =   600
   End
   Begin VB.CommandButton CmdNewSupplier 
      BackColor       =   &H0000C0C0&
      Caption         =   " «„Ì‰ ò‰‰œÂ ÃœÌœ"
      BeginProperty Font 
         Name            =   "B Homa"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   600
      RightToLeft     =   -1  'True
      TabIndex        =   12
      Top             =   0
      Width           =   2055
   End
   Begin VB.CommandButton OKButton 
      BackColor       =   &H00008000&
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
      Left            =   1560
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   7920
      Width           =   1215
   End
   Begin VB.CommandButton CancelButton 
      BackColor       =   &H000000C0&
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
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   7920
      Width           =   1215
   End
   Begin VB.TextBox txtName 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "Nazanin"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   6720
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   600
      Width           =   2595
   End
   Begin VB.TextBox txtMembershipId 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "Nazanin"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   6720
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   165
      Width           =   2595
   End
   Begin VB.TextBox txtAddress 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "Nazanin"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   705
      Left            =   6720
      MultiLine       =   -1  'True
      RightToLeft     =   -1  'True
      TabIndex        =   1
      Top             =   1470
      Width           =   2595
   End
   Begin VB.TextBox txtPhone 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000F&
      BeginProperty Font 
         Name            =   "Nazanin"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   6720
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   1035
      Width           =   2595
   End
   Begin VSFlex7LCtl.VSFlexGrid vsCustomer 
      Height          =   5565
      Left            =   0
      TabIndex        =   4
      Top             =   2280
      Width           =   10635
      _cx             =   18759
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
      SelectionMode   =   3
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   7
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   500
      RowHeightMax    =   500
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmFindSupplier.frx":A581
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
      OleObjectBlob   =   "frmFindSupplier.frx":A6BD
      TabIndex        =   18
      Top             =   0
      Width           =   480
   End
   Begin VB.Label LblCount 
      Alignment       =   2  'Center
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
      Height          =   495
      Left            =   3840
      RightToLeft     =   -1  'True
      TabIndex        =   17
      Top             =   8040
      Width           =   3855
   End
   Begin VB.Label LblFindCust 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Nazanin"
         Size            =   11.25
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   615
      Left            =   3600
      RightToLeft     =   -1  'True
      TabIndex        =   11
      Top             =   120
      Width           =   3015
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "‰«„"
      BeginProperty Font 
         Name            =   "B Homa"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   9330
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   600
      Width           =   1185
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "òœ «‘ —«ò"
      BeginProperty Font 
         Name            =   "B Homa"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   9360
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   165
      Width           =   1185
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "¬œ—”"
      BeginProperty Font 
         Name            =   "B Homa"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   9360
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   1560
      Width           =   1185
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   " ·›‰"
      BeginProperty Font 
         Name            =   "B Homa"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   9360
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   1035
      Width           =   1185
   End
End
Attribute VB_Name = "frmFindSupplier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim clsDate As New clsDate
Dim SearchType As Integer
Dim i, j As Long
Dim Rst As New ADODB.Recordset
Dim TotalCreditDebit As Currency

Private Sub CancelButton_Click()
    mvarcode = 0
    Unload Me
End Sub

Private Sub CmdNewSupplier_Click()
    If ClsFormAccess.frmSupplier = True Then
        Unload Me
        frmSupplier.Show
    End If
    
End Sub

Private Sub cmdTurnOver_Click()
    If vsCustomer.Row < 1 Then Exit Sub
    
    If ClsFormAccess.AccfrmKartHesabReport = True Then
        If vsCustomer.ValueMatrix(vsCustomer.Row, 8) > 0 Then
            Accounting.KartHesabShowDll "KolBestankaran", CStr(vsCustomer.TextMatrix(vsCustomer.Row, 8)), CStr(vsCustomer.TextMatrix(vsCustomer.Row, 3)), Mid(AccountYear & "/01/01", 3), Mid(clsDate.shamsi(Date), 3)
        Else
            ShowDisMessage "«Ì‰ „‘ —Ì œ— ”Ì” „ Õ”«»œ«—Ì œ«—«Ì ﬂœ  ›÷Ì·Ì ‰Ì” ", 2000
        End If
    Else
        ShowDisMessage "‘„« »Â «Ì‰ «„ﬂ«‰ œ” —”Ì ‰œ«—Ìœ", 1500
    End If
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call PresetScreenSaver
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call PresetScreenSaver
End Sub

Private Sub Form_Activate()

    Select Case clsStation.DefaultCustSearch
        Case EnumDefaultCustSearch.MembershipId
            txtMembershipId.SetFocus
        Case EnumDefaultCustSearch.address
            TxtAddress.SetFocus
        Case EnumDefaultCustSearch.Name
            TxtName.SetFocus
        Case EnumDefaultCustSearch.Phone
            txtPhone.SetFocus
    End Select
    
End Sub

Private Sub Form_Load()

    formloadFlag = False
    CenterCenterinSecondScreen Me
    
    mvarcode = 0
    With vsCustomer
        .Cols = 8
        .TextMatrix(0, 7) = " ›÷Ì·Ì"
        For i = 0 To .Cols - 1
            .ColWidth(i) = Val(GetSetting(strMainKey, Me.Name, "Col" & i))
            If .ColWidth(i) = 0 Then
                .ColWidth(i) = .Width / 5     'Row
            End If
        Next i
    End With
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

    FillvsCustomer

    If clsArya.ExternalAccounting = True Then cmdTurnOver.Visible = True Else cmdTurnOver.Visible = False
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveSetting strMainKey, Me.Name, "Left", Me.Left
    SaveSetting strMainKey, Me.Name, "Top", Me.Top


    If PosConnection.State = adStateOpen Then PosConnection.Close: Set PosConnection = Nothing
End Sub

Private Sub OKButton_Click()
    If vsCustomer.Row > 0 Then
        mvarcode = vsCustomer.TextMatrix(vsCustomer.Row, 1)
        mvarName = vsCustomer.TextMatrix(vsCustomer.Row, 3)
    Else
        mvarcode = 0
        mvarName = ""
        mvarPublicOrderType = inPerson
    End If
    Unload Me

End Sub

Private Sub Option1_Click(index As Integer)
    If Option1(0).Value = False Then
        FillvsCustomer
    Else
        vsCustomer.Rows = 1
    End If
    txtMembershipId.SetFocus
End Sub

Private Sub ResizeKit1_ExitResize(ByVal XScale As Double, ByVal YScale As Double)

    If formloadFlag = True Then
        SaveSetting strMainKey, Me.Name, "Height", Me.Height
        SaveSetting strMainKey, Me.Name, "Width", Me.Width
    End If

End Sub

Private Sub txtAddress_Change()

    If Option1(0).Value = False Then
        i = vsCustomer.FindRow(TxtAddress.Text, 1, 5, False, False)
        If i > 0 Then
            vsCustomer.Row = i
            vsCustomer.ShowCell i, 0
            LblFindCust.Caption = ""
        Else
            vsCustomer.Row = 0
            vsCustomer.ShowCell 0, 0
            If TxtAddress.Text <> "" Then
               LblFindCust.Caption = " ¬œ—” ( " & TxtAddress.Text & " )œ— ”Ì” „  ⁄—Ì› ‰‘œÂ"
             Else
                LblFindCust.Caption = "¬œ—” „‘ —ﬂ —« Ê«—œ ﬂ‰Ìœ  "
             End If
    
        End If
   Else
       If Len(TxtAddress.Text) > 0 Then
            SearchType = 4
            Timer1.Interval = Val(TxtTimer.Text)
            Timer1.Enabled = True
       Else
          vsCustomer.Rows = 1
      
       End If
  
   End If

End Sub
Private Sub Timer1_Timer()
    Timer1.Enabled = False
    Define_Customer
    If vsCustomer.Rows > 1 Then
        vsCustomer.Row = 1
        vsCustomer.ShowCell 1, 0
        LblFindCust.Caption = ""
    Else
        vsCustomer.Row = 0
        vsCustomer.ShowCell 0, 0
        Select Case SearchType
            Case 1:
                 If Val(txtMembershipId.Text) > 0 Then
                   LblFindCust.Caption = " «‘ —«ﬂ ( " & txtMembershipId.Text & " ) œ— ”Ì” „  ⁄—Ì› ‰‘œÂ"
                 Else
                    LblFindCust.Caption = "ﬂœ «‘ —«ﬂ —« Ê«—œ ﬂ‰Ìœ  "
                 End If
            Case 2:
                 If Len(TxtName.Text) > 0 Then
                   LblFindCust.Caption = " ‰«„ ( " & TxtName.Text & " ) œ— ”Ì” „  ⁄—Ì› ‰‘œÂ"
                 Else
                    LblFindCust.Caption = "‰«„ „‘ —ﬂ —« Ê«—œ ﬂ‰Ìœ  "
                 End If
            Case 3:
                 If Len(txtPhone.Text) > 0 Then
                    LblFindCust.Caption = "  ·›‰ ( " & txtPhone.Text & " ) œ— ”Ì” „  ⁄—Ì› ‰‘œÂ"
                 Else
                    LblFindCust.Caption = " ·›‰ „‘ —ﬂ —« Ê«—œ ﬂ‰Ìœ  "
                 End If
            Case 4:
                 If Len(TxtAddress.Text) > 0 Then
                   LblFindCust.Caption = " ¬œ—” ( " & TxtAddress.Text & " ) œ— ”Ì” „  ⁄—Ì› ‰‘œÂ"
                 Else
                    LblFindCust.Caption = "¬œ—” «‘ —«ﬂ —« Ê«—œ ﬂ‰Ìœ  "
                 End If
        End Select
    End If
            
End Sub

Private Sub txtAddress_GotFocus()
    txtPhone.Text = ""
    TxtName.Text = ""
    txtMembershipId.Text = ""
    vsCustomer.Row = 0
    LblCount = ""

    vsCustomer.Select vsCustomer.Row, 5
    vsCustomer.Sort = flexSortGenericAscending

End Sub

Private Sub txtMembershipId_Change()
    If Option1(0).Value = False Then
        i = vsCustomer.FindRow(txtMembershipId.Text, 1, 2, True, True)
        If i > 0 Then
            vsCustomer.Row = i
            vsCustomer.ShowCell i, 0
            LblFindCust.Caption = ""
        Else
            vsCustomer.Row = 0
            vsCustomer.ShowCell 0, 0
            If Val(txtMembershipId.Text) > 0 Then
               LblFindCust.Caption = " «‘ —«ﬂ ( " & txtMembershipId.Text & " ) œ— ”Ì” „  ⁄—Ì› ‰‘œÂ"
             Else
                LblFindCust.Caption = "ﬂœ «‘ —«ﬂ —« Ê«—œ ﬂ‰Ìœ  "
             End If
        
        End If
   Else
       If Val(txtMembershipId.Text) > 0 Then
            SearchType = 1
            Timer1.Interval = Val(TxtTimer.Text)
            Timer1.Enabled = True
       Else
          vsCustomer.Rows = 1
       
       End If
  
   End If
    
    
End Sub

Private Sub txtMembershipId_GotFocus()
    txtPhone.Text = ""
    TxtAddress.Text = ""
    TxtName.Text = ""
    vsCustomer.Row = 0
    LblCount = ""

    vsCustomer.Select vsCustomer.Row, 2
    vsCustomer.Sort = flexSortGenericAscending
    
End Sub



Private Sub txtName_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 40 Then
    If vsCustomer.Row >= 1 Then
         vsCustomer.SetFocus
         If vsCustomer.Rows > 2 Then
            vsCustomer.Row = 2
         End If
    End If
End If
End Sub

Private Sub txtName_Change()

    If Option1(0).Value = False Then
        i = vsCustomer.FindRow(TxtName.Text, 1, 3, False, False)
        If i > 0 Then
            vsCustomer.Row = i
            vsCustomer.ShowCell i, 0
            LblFindCust.Caption = ""
        Else
            vsCustomer.Row = 0
            vsCustomer.ShowCell 0, 0
            If TxtName.Text <> "" Then
               LblFindCust.Caption = " ‰«„ ( " & TxtName.Text & " ) œ— ”Ì” „  ⁄—Ì› ‰‘œÂ"
             Else
                LblFindCust.Caption = "‰«„ „‘ —ﬂ —« Ê«—œ ﬂ‰Ìœ  "
             End If
    
        End If
    Else
       If Len(TxtName.Text) > 0 Then
            SearchType = 2
            Timer1.Interval = Val(TxtTimer.Text)
            Timer1.Enabled = True
       Else
          vsCustomer.Rows = 1

       End If
  
   End If

End Sub

Private Sub txtName_GotFocus()
    txtPhone.Text = ""
    TxtAddress.Text = ""
    txtMembershipId.Text = ""
    vsCustomer.Row = 0
    LblCount = ""

    vsCustomer.Select vsCustomer.Row, 3
    vsCustomer.Sort = flexSortGenericAscending
End Sub

Private Sub FillvsCustomer()

    Dim Rst As New ADODB.Recordset
    ReDim Parameter(0) As Parameter
    Parameter(0) = GenerateInputParameter("@Branch", adInteger, 4, CurrentBranch)
    Set Rst = RunParametricStoredProcedure2Rec("Get_All_Supplier", Parameter)
    
    With vsCustomer
        .Rows = 1
        i = 0
        While Rst.EOF <> True
            .Rows = .Rows + 1
            i = .Rows - 1
            .TextMatrix(i, 0) = i
            .TextMatrix(i, 1) = Rst!Code
            .TextMatrix(i, 2) = Rst!MembershipId
            .TextMatrix(i, 3) = Rst![Name]
            .TextMatrix(i, 4) = Rst!Tel1
            .TextMatrix(i, 5) = Rst!address
            .TextMatrix(i, 6) = Rst!Discount
            .TextMatrix(i, 7) = IIf(IsNull(Rst!Tafsili), " ", Rst!Tafsili)
            Rst.MoveNext
        Wend
    End With
    Set Rst = Nothing


End Sub

Private Sub txtPhone_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 40 Then
    If vsCustomer.Row >= 1 Then
         vsCustomer.SetFocus
         If vsCustomer.Rows > 2 Then
            vsCustomer.Row = 2
         End If
    End If
End If

End Sub
Private Sub txtPhone_Change()

   If Option1(0).Value = False Then
        i = vsCustomer.FindRow(txtPhone.Text, 1, 4, False, False)
        If i > 0 Then
            vsCustomer.Row = i
            vsCustomer.ShowCell i, 0
            LblFindCust.Caption = ""
        Else
            vsCustomer.Row = 0
            vsCustomer.ShowCell 0, 0
            If txtPhone.Text <> "" Then
               LblFindCust.Caption = "  ·›‰ ( " & txtPhone.Text & " ) œ— ”Ì” „  ⁄—Ì› ‰‘œÂ"
             Else
                LblFindCust.Caption = " ·›‰ „‘ —ﬂ —« Ê«—œ ﬂ‰Ìœ  "
             End If
    
        End If
    Else
       If Len(txtPhone.Text) > 0 Then
            SearchType = 3
            Timer1.Interval = Val(TxtTimer.Text)
            Timer1.Enabled = True
       Else
          vsCustomer.Rows = 1
      
       End If
  
   End If

End Sub

Private Sub txtPhone_GotFocus()
    TxtName.Text = ""
    TxtAddress.Text = ""
    txtMembershipId.Text = ""
    vsCustomer.Row = 0
    LblCount = ""

    vsCustomer.Select vsCustomer.Row, 4
    vsCustomer.Sort = flexSortGenericAscending

End Sub


Private Sub vsCustomer_AfterSort(ByVal Col As Long, Order As Integer)
    For i = 1 To vsCustomer.Rows - 1
        vsCustomer.TextMatrix(i, 0) = i
    Next
    
End Sub

Private Sub vsCustomer_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    For i = 0 To vsCustomer.Cols - 1
        SaveSetting strMainKey, Me.Name, "Col" & Col, vsCustomer.ColWidth(Col)
    Next
End Sub

Private Sub vsCustomer_DblClick()
    If vsCustomer.Row > 0 Then
        OKButton_Click
    End If
End Sub

Private Sub Define_Customer()
    
    ReDim Parameter(1) As Parameter
    If VarActForm = "frmSupplier" Then
        Parameter(0) = GenerateInputParameter("@ActDeact", adInteger, 4, 2) ' All Customers
    Else
        Parameter(0) = GenerateInputParameter("@ActDeact", adInteger, 4, 0) ' Only Active
    End If
    Select Case SearchType
        Case 1
            Parameter(1) = GenerateInputParameter("@Code", adBigInt, 8, Val(txtMembershipId.Text))
            Set Rst = RunParametricStoredProcedure2Rec("Get_Supplier_Code", Parameter)
        Case 2
            Parameter(1) = GenerateInputParameter("@Name", adVarWChar, 50, Left(TxtName.Text, 50))
            Set Rst = RunParametricStoredProcedure2Rec("Get_Supplier_Name", Parameter)
        Case 3
            Parameter(1) = GenerateInputParameter("@Tel", adVarWChar, 10, Left(txtPhone.Text, 10))
            Set Rst = RunParametricStoredProcedure2Rec("Get_Supplier_Tel", Parameter)
        Case 4
            Parameter(1) = GenerateInputParameter("@Addresse", adVarWChar, 100, Left(TxtAddress.Text, 100))
            Set Rst = RunParametricStoredProcedure2Rec("Get_Supplier_Address", Parameter)
    End Select
    Dim TmpTel As String
    Dim jj As Integer
    
    With vsCustomer
        .Rows = 1
        i = 0
        j = 0
        Do While Rst.EOF <> True
            TmpTel = Rst!Tel1
            j = j + 1
            For jj = 1 To 5
                If TmpTel <> "" Or (jj = 1) Then
                    i = i + 1
                    .Rows = .Rows + 1
                    .TextMatrix(i, 0) = j
                    .TextMatrix(i, 1) = Rst!Code
                    .TextMatrix(i, 2) = Rst!MembershipId
                    .TextMatrix(i, 3) = Rst![Name]
                    .TextMatrix(i, 4) = Left(Trim(TmpTel), 10)
                    .TextMatrix(i, 5) = IIf(IsNull(Rst!address), " ", Rst!address)
                    .TextMatrix(i, 7) = IIf(IsNull(Rst!Tafsili), " ", Rst!Tafsili)
''''                    .TextMatrix(i, 6) = Rst!discount
''''                    .TextMatrix(i, 7) = Rst!Credit
''''                    .TextMatrix(i, 8) = Rst!carryfee
''''                    .TextMatrix(i, 9) = Rst!PaykFee
''''                    .TextMatrix(i, 10) = Rst!Distance
                    TmpTel = ""
                End If
                If jj = 1 And Trim(Rst!Tel2) <> "" Then
                    TmpTel = Rst!Tel2
                ElseIf jj = 2 And Trim(Rst!Tel3) <> "" Then
                    TmpTel = Rst!Tel3
                ElseIf jj = 3 And Trim(Rst!Tel4) <> "" Then
                    TmpTel = Rst!Tel4
                ElseIf jj = 4 And Trim(Rst!Mobile) <> "" Then
                    TmpTel = Rst!Mobile
                End If
            Next jj
             
            
            Rst.MoveNext
        Loop
        vsCustomer.MergeCompare = flexMCTrimNoCase
        vsCustomer.MergeCells = flexMergeRestrictRows
        vsCustomer.MergeRow(vsCustomer.Rows - 1) = True
        vsCustomer.MergeCol(0) = True
        vsCustomer.MergeCol(1) = True
        vsCustomer.MergeCol(2) = True
        vsCustomer.MergeCol(3) = True
    ''''    vsCustomer.MergeCol(4) = True
        vsCustomer.MergeCol(5) = True
        If i > 0 Then
            LblCount.Caption = " ⁄œ«œ —ﬂÊ—œÂ«   -  " & i
        Else
            LblCount.Caption = ""
        End If
        vsCustomer.AutoSizeMode = flexAutoSizeColWidth
        vsCustomer.AutoSize 0, .Cols - 1
    If vsCustomer.ColWidth(3) < 3000 Then
        vsCustomer.ColWidth(3) = 3000
    End If
    If vsCustomer.ColWidth(4) < 1500 Then
        vsCustomer.ColWidth(4) = 1500
    End If
    If vsCustomer.ColWidth(5) < 4000 Then
        vsCustomer.ColWidth(5) = 4000
    End If
    
    End With
    Set Rst = Nothing
    
End Sub

Private Sub GetCreditDebit()
    On Error GoTo Err_Handler
    Dim TotalBedehkar, TotalBestankar As Double
    Dim L_Rst As New ADODB.Recordset
    
    Me.LblTotalCreditDebit.Caption = ""
    If vsCustomer.Row < 1 Then Exit Sub
    If clsArya.ExternalAccounting = True And Val(vsCustomer.TextMatrix(vsCustomer.Row, 7)) > 0 Then
        Set L_Rst = Accounting.GetCreditDebitDll(Val(vsCustomer.TextMatrix(vsCustomer.Row, 7)), 1)
        If L_Rst.BOF = True And L_Rst.EOF = True Then
            Set L_Rst = Nothing
            Exit Sub
        Else
            TotalCreditDebit = 0: TotalBedehkar = 0: TotalBestankar = 0
            While L_Rst.EOF = False
                TotalBedehkar = TotalBedehkar + L_Rst.Fields("Bedehkar").Value
                TotalBestankar = TotalBestankar + L_Rst.Fields("Bestankar").Value
                L_Rst.MoveNext
            Wend
        
        End If
        TotalCreditDebit = TotalBedehkar - TotalBestankar
        L_Rst.Close
        Set L_Rst = Nothing
        If TotalCreditDebit > 0 Then
            LblTotalCreditDebitLabel = "»œÂÌ ﬂ·: "
            LblTotalCreditDebit = Format(TotalCreditDebit, "#,##") & clsArya.UnitPrice
            LblTotalCreditDebitLabel.ForeColor = vbRed
            LblTotalCreditDebit.ForeColor = vbRed
        ElseIf TotalCreditDebit = 0 Then
            LblTotalCreditDebitLabel.Caption = "»œÂÌ- ÿ·» ﬂ·: "
            LblTotalCreditDebit.Caption = Format(TotalCreditDebit, "#,##") & clsArya.UnitPrice
            LblTotalCreditDebitLabel.ForeColor = vbGreen
            LblTotalCreditDebit.ForeColor = vbGreen
        Else
            TotalCreditDebit = Abs(TotalCreditDebit)
            LblTotalCreditDebitLabel.Caption = "ÿ·» ﬂ·: "
            LblTotalCreditDebit.Caption = Format((TotalCreditDebit), "#,##") & clsArya.UnitPrice
            LblTotalCreditDebitLabel.ForeColor = vbGreen
            LblTotalCreditDebit.ForeColor = vbGreen
            
        End If
        
    End If
    
Exit Sub

Err_Handler:
    LogSaveNew "frmFindCust => ", err.Description, err.Number, err.Source, "GetCreditDebit"
    ShowErrorMessage
    err.Clear
    Resume Next
    Set L_Rst = Nothing
End Sub

Private Sub vsCustomer_RowColChange()
    GetCreditDebit
End Sub
