VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{7AEDC602-D94C-11D1-BB7A-00E0290EA3C9}#1.0#0"; "ResizeKit.ocx"
Begin VB.Form frmFindPhoneBook 
   Caption         =   "                                                              Ã” ÃÊÌ œ› —çÂ  ·›‰"
   ClientHeight    =   7725
   ClientLeft      =   2775
   ClientTop       =   3870
   ClientWidth     =   11070
   BeginProperty Font 
      Name            =   "Nazanin"
      Size            =   12
      Charset         =   178
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmFindPhoneBook.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   7725
   ScaleWidth      =   11070
   Begin VB.PictureBox Picture1 
      Height          =   615
      Left            =   3000
      RightToLeft     =   -1  'True
      ScaleHeight     =   555
      ScaleWidth      =   3795
      TabIndex        =   20
      Top             =   360
      Width           =   3855
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
         Left            =   0
         RightToLeft     =   -1  'True
         TabIndex        =   22
         Top             =   0
         Width           =   1815
      End
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
         Left            =   1800
         RightToLeft     =   -1  'True
         TabIndex        =   21
         Top             =   0
         Width           =   1815
      End
   End
   Begin VB.Frame Frame7 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   6960
      RightToLeft     =   -1  'True
      TabIndex        =   14
      Top             =   240
      Width           =   3975
      Begin VB.TextBox txtFirstName 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         Height          =   435
         Left            =   0
         RightToLeft     =   -1  'True
         TabIndex        =   18
         Top             =   1080
         Width           =   2595
      End
      Begin VB.TextBox txtLastName 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         Height          =   435
         Left            =   0
         RightToLeft     =   -1  'True
         TabIndex        =   0
         Top             =   195
         Width           =   2595
      End
      Begin VB.TextBox txtTel 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000018&
         Height          =   435
         Left            =   0
         RightToLeft     =   -1  'True
         TabIndex        =   1
         Top             =   630
         Width           =   2595
      End
      Begin VB.Label lblFirstName 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H8000000D&
         Height          =   435
         Left            =   2640
         RightToLeft     =   -1  'True
         TabIndex        =   17
         Top             =   1080
         Width           =   1065
      End
      Begin VB.Label lblLastName 
         Alignment       =   1  'Right Justify
         Caption         =   "‰«„ Œ«‰Ê«œêÌ"
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
         Left            =   2640
         RightToLeft     =   -1  'True
         TabIndex        =   16
         Top             =   240
         Width           =   1065
      End
      Begin VB.Label lblTel 
         Alignment       =   1  'Right Justify
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
         ForeColor       =   &H8000000D&
         Height          =   435
         Left            =   2640
         RightToLeft     =   -1  'True
         TabIndex        =   15
         Top             =   600
         Width           =   1065
      End
   End
   Begin VB.Frame Frame5 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3000
      RightToLeft     =   -1  'True
      TabIndex        =   10
      Top             =   960
      Width           =   3855
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
         Left            =   1440
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Text            =   "500"
         Top             =   120
         Width           =   615
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
         Left            =   2280
         RightToLeft     =   -1  'True
         TabIndex        =   13
         Top             =   120
         Width           =   1215
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
         Left            =   360
         RightToLeft     =   -1  'True
         TabIndex        =   12
         Top             =   120
         Width           =   855
      End
   End
   Begin VB.Frame Frame4 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6720
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   6840
      Width           =   3495
      Begin VB.TextBox txtMaxRecord 
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
         Height          =   495
         Left            =   0
         RightToLeft     =   -1  'True
         TabIndex        =   9
         Text            =   "300"
         Top             =   120
         Width           =   615
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         Caption         =   "Õœ«ﬂÀ—  ⁄œ«œ Œ—ÊÃÌÂ« »—«Ì ‰„«Ì‘"
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
         Left            =   600
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   120
         Width           =   2775
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   600
      Top             =   0
   End
   Begin VSFlex7LCtl.VSFlexGrid vsPhoneBook 
      Height          =   4845
      Left            =   120
      TabIndex        =   2
      Top             =   1920
      Width           =   10875
      _cx             =   19182
      _cy             =   8546
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
      AllowUserResizing=   0
      SelectionMode   =   3
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   11
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   500
      RowHeightMax    =   500
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmFindPhoneBook.frx":A4C2
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
      Height          =   495
      Left            =   120
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   6960
      Width           =   1215
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
      Height          =   495
      Left            =   1560
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   6960
      Width           =   1215
   End
   Begin RESIZEKITLibCtl.ResizeKit ResizeKit1 
      Height          =   480
      Left            =   0
      OleObjectBlob   =   "frmFindPhoneBook.frx":A610
      TabIndex        =   19
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
      Left            =   3480
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   6960
      Width           =   2775
   End
   Begin VB.Label LblFindPhoneBook 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Nazanin"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   1455
      Left            =   240
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   360
      Width           =   2655
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmFindPhoneBook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim i, j As Long
Dim SearchType As Integer
Dim Rst As New ADODB.Recordset
Dim tmpflag As Boolean

Private Sub CancelButton_Click()
    mvarcode = 0
    Unload Me
     frmPhoneBook.SetFocus
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call PresetScreenSaver
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call PresetScreenSaver
End Sub

Private Sub Form_Activate()
     txtMaxRecord.Text = clsStation.MaxRecordCount
    TxtTimer.Text = clsStation.SrarchInputDelayKeyboard
    Option1(0).Value = clsStation.CustomerSearchDefault
    Option1(1).Value = Not (clsStation.CustomerSearchDefault)
    Select Case 0 'clsStation.DefaultCustSearch
        Case EnumDefaultPhoneBookSearch.LastName
            txtLastName.SetFocus
            LblFindPhoneBook.Caption = "‰«„ Œ«‰Ê«œêÌ  —« Ê«—œ ﬂ‰Ìœ  "
        Case EnumDefaultPhoneBookSearch.Tel
            txtTel.SetFocus
            LblFindPhoneBook.Caption = " ·›‰  —« Ê«—œ ﬂ‰Ìœ  "
        Case EnumDefaultPhoneBookSearch.FirstName
            txtFirstName.SetFocus
            LblFindPhoneBook.Caption = "‰«„  —« Ê«—œ ﬂ‰Ìœ  "
    End Select
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 115 Then
    If ClsFormAccess.frmPhoneBook = True Then
        Unload Me
        frmPhoneBook.Show
    End If
End If

End Sub

Private Sub Form_Load()
    CenterCenterinSecondScreen Me
    
    mvarcode = 0
    vsPhoneBook.ColHidden(1) = True

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

Private Sub Option1_Click(index As Integer)
    If Option1(0).Value = False Then
        Frame5.Visible = False
        FillvsPhoneBook
    Else
        Frame5.Visible = True
        vsPhoneBook.Rows = 1
        labelClear

    End If
    txtLastName.SetFocus
End Sub

Private Sub Form_Unload(Cancel As Integer)


    SaveSetting strMainKey, Me.Name, "Left", Me.Left
    SaveSetting strMainKey, Me.Name, "Top", Me.Top


    If PosConnection.State = adStateOpen Then PosConnection.Close: Set PosConnection = Nothing
End Sub

Private Sub OKButton_Click()
    If vsPhoneBook.Row > 0 Then
        mvarcode = vsPhoneBook.TextMatrix(vsPhoneBook.Row, 1)
    Else
        mvarcode = 0
    End If
    Unload Me

End Sub

Private Sub ResizeKit1_ExitResize(ByVal XScale As Double, ByVal YScale As Double)


    If formloadFlag = True Then
        SaveSetting strMainKey, Me.Name, "Height", Me.Height
        SaveSetting strMainKey, Me.Name, "Width", Me.Width
    End If

End Sub

Private Sub Timer1_Timer()
    Timer1.Enabled = False
    Define_Customer
    If vsPhoneBook.Rows > 1 Then
        vsPhoneBook.Row = 1
        vsPhoneBook.ShowCell 1, 0
        LblFindPhoneBook.Caption = ""
    Else
        vsPhoneBook.Row = 0
        vsPhoneBook.ShowCell 0, 0
        labelClear
        Select Case SearchType
            Case 1:
                 If Val(txtLastName.Text) > 0 Then
                   LblFindPhoneBook.Caption = "  ( " & txtLastName.Text & " ) œ— ”Ì” „  ⁄—Ì› ‰‘œÂ"
                 Else
                    LblFindPhoneBook.Caption = "‰«„ Œ«‰Ê«œêÌ  —« Ê«—œ ﬂ‰Ìœ  "
                 End If
            
            Case 2:
                 If Len(txtTel.Text) > 0 Then
                    LblFindPhoneBook.Caption = "  ·›‰ ( " & txtTel.Text & " ) œ— ”Ì” „  ⁄—Ì› ‰‘œÂ"
                 Else
                    LblFindPhoneBook.Caption = " ·›‰  —« Ê«—œ ﬂ‰Ìœ  "
                 End If
            Case 3:
                 If Len(txtFirstName.Text) > 0 Then
                   LblFindPhoneBook.Caption = " ‰«„ ( " & txtFirstName.Text & " ) œ— ”Ì” „  ⁄—Ì› ‰‘œÂ"
                 Else
                    LblFindPhoneBook.Caption = "‰«„  —« Ê«—œ ﬂ‰Ìœ  "
                 End If
            
        End Select
    End If
            
End Sub

Private Sub txtMaxRecord_Change()
    clsStation.MaxRecordCount = Val(txtMaxRecord.Text)
    SetStationSettingFile
End Sub

Private Sub txtLastName_Change()
    If Option1(0).Value = False Then
        i = vsPhoneBook.FindRow(txtLastName.Text, 1, 3, False, False)
        If i > 0 Then
            vsPhoneBook.Row = i
            vsPhoneBook.ShowCell i, 0
            LblFindPhoneBook.Caption = ""
        Else
            vsPhoneBook.Row = 0
            vsPhoneBook.ShowCell 0, 0
            labelClear
            If Val(txtLastName.Text) > 0 Then
               LblFindPhoneBook.Caption = "  ( " & txtLastName.Text & " ) œ— ”Ì” „  ⁄—Ì› ‰‘œÂ"
             Else
                LblFindPhoneBook.Caption = "‰«„ Œ«‰Ê«œêÌ  —« Ê«—œ ﬂ‰Ìœ  "
             End If
        
        End If
   Else
       If Len(txtLastName.Text) > 0 Then
            SearchType = 1
            Timer1.Interval = Val(TxtTimer.Text)
            Timer1.Enabled = True
       Else
          vsPhoneBook.Rows = 1
          labelClear
        
       End If
  
   End If
    
End Sub

Private Sub txtLastName_GotFocus()

    txtTel.Text = ""
    txtFirstName.Text = ""
    vsPhoneBook.Row = 0
    labelClear
    LblCount = ""
    vsPhoneBook.Select vsPhoneBook.Row, 3
    vsPhoneBook.Sort = flexSortGenericAscending
    LblFindPhoneBook.Caption = "‰«„ Œ«‰Ê«œêÌ  —« Ê«—œ ﬂ‰Ìœ  "
    
End Sub
Private Sub txtFirstName_Change()

    If Option1(0).Value = False Then
        i = vsPhoneBook.FindRow(txtFirstName.Text, 1, 2, False, False)
        If i > 0 Then
            vsPhoneBook.Row = i
            vsPhoneBook.ShowCell i, 0
            LblFindPhoneBook.Caption = ""
        Else
            vsPhoneBook.Row = 0
            vsPhoneBook.ShowCell 0, 0
            labelClear
            If txtFirstName.Text <> "" Then
               LblFindPhoneBook.Caption = " ‰«„ ( " & txtFirstName.Text & " ) œ— ”Ì” „  ⁄—Ì› ‰‘œÂ"
             Else
                LblFindPhoneBook.Caption = "‰«„  —« Ê«—œ ﬂ‰Ìœ  "
             End If
    
        End If
    Else
       If Len(txtFirstName.Text) > 0 Then
            SearchType = 3
            Timer1.Interval = Val(TxtTimer.Text)
            Timer1.Enabled = True
       Else
          vsPhoneBook.Rows = 1
          labelClear

       End If
  
   End If

End Sub

Private Sub txtFirstName_GotFocus()
    txtTel.Text = ""
    txtLastName.Text = ""
    vsPhoneBook.Row = 0
    labelClear
    LblCount = ""
    vsPhoneBook.Select vsPhoneBook.Row, 2
    vsPhoneBook.Sort = flexSortGenericAscending
    LblFindPhoneBook.Caption = "‰«„  —« Ê«—œ ﬂ‰Ìœ  "

End Sub


Private Sub FillvsPhoneBook()
    Dim i As Integer
    ReDim Parameter(0) As Parameter
    Parameter(0) = GenerateInputParameter("@intPhoneBookNo", adInteger, 4, -1) ' All Customers
    Set Rst = RunParametricStoredProcedure2Rec("Get_tblTotal_tPhoneBook_Info", Parameter)
    i = 0
    With vsPhoneBook
        
        .Rows = 1
         While Rst.EOF <> True
            .Rows = .Rows + 1
            i = i + 1
            .TextMatrix(i, 0) = i
            .TextMatrix(i, 1) = Rst!intPhoneBookNo
            .TextMatrix(i, 2) = IIf(IsNull(Rst!nvcFirstName), "", Rst!nvcFirstName)
            .TextMatrix(i, 3) = Rst!nvcLastName
            .TextMatrix(i, 4) = Trim(Rst!nvcTelCollection)
            .TextMatrix(i, 5) = IIf(IsNull(Rst!nvcMobile), "", Rst!nvcMobile)
            .TextMatrix(i, 6) = IIf(IsNull(Rst!nvcTelCompany), "", Rst!nvcTelCompany)
            .TextMatrix(i, 7) = IIf(IsNull(Rst!nvcFax), "", Rst!nvcFax)
            .TextMatrix(i, 8) = IIf(IsNull(Rst!nvcEmail), "", Rst!nvcEmail)
            .TextMatrix(i, 9) = IIf(IsNull(Rst!nvcAddress), "", Rst!nvcAddress)
            .TextMatrix(i, 10) = Rst!nvcDate
            
            Rst.MoveNext
        Wend
    End With
    LblCount.Caption = " ⁄œ«œ —ﬂÊ—œÂ«   -  " & i
    Set Rst = Nothing
'vsFactors.set_MergeRow(vsFactors.Rows -1 , true)
'vsFactors.set_MergeCol(2,true);
    vsPhoneBook.MergeCompare = flexMCTrimNoCase
    vsPhoneBook.MergeCells = flexMergeRestrictRows
    vsPhoneBook.MergeRow(vsPhoneBook.Rows - 1) = True
    vsPhoneBook.MergeCol(0) = True
    vsPhoneBook.MergeCol(1) = True
    vsPhoneBook.MergeCol(2) = True
    vsPhoneBook.MergeCol(3) = True
    vsPhoneBook.MergeCol(4) = True
    vsPhoneBook.MergeCol(5) = True
    vsPhoneBook.MergeCol(6) = True
    vsPhoneBook.MergeCol(7) = True
    vsPhoneBook.ColWidth(3) = vsPhoneBook.ColWidth(3) * 1.1
    vsPhoneBook.AutoSizeMode = flexAutoSizeColWidth
    vsPhoneBook.AutoSize 0, vsPhoneBook.Cols - 1
    If vsPhoneBook.ColWidth(3) < 3000 Then
        vsPhoneBook.ColWidth(3) = 3000
    End If
    If vsPhoneBook.ColWidth(4) < 1500 Then
        vsPhoneBook.ColWidth(4) = 1500
    End If
    If vsPhoneBook.ColWidth(5) < 4000 Then
        vsPhoneBook.ColWidth(5) = 4000
    End If


End Sub


Private Sub txtFirstName_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 40 Then
    If vsPhoneBook.Row >= 1 Then
         vsPhoneBook.SetFocus
         If vsPhoneBook.Rows > 2 Then
            vsPhoneBook.Row = 2
         End If
    End If
End If
End Sub

Private Sub txtTel_Change()

    If Option1(0).Value = False Then
        i = vsPhoneBook.FindRow(txtTel.Text, 1, 4, False, False)
        If i > 0 Then
            vsPhoneBook.Row = i
            vsPhoneBook.ShowCell i, 0
            LblFindPhoneBook.Caption = ""
        Else
            vsPhoneBook.Row = 0
            vsPhoneBook.ShowCell 0, 0
            labelClear
            If txtTel.Text <> "" Then
               LblFindPhoneBook.Caption = "  ·›‰ ( " & txtTel.Text & " ) œ— ”Ì” „  ⁄—Ì› ‰‘œÂ"
             Else
                LblFindPhoneBook.Caption = " ·›‰ „‘ —ﬂ —« Ê«—œ ﬂ‰Ìœ  "
             End If
    
        End If
    Else
       If Len(txtTel.Text) > 0 Then
            SearchType = 2
            Timer1.Interval = Val(TxtTimer.Text)
            Timer1.Enabled = True
       Else
          vsPhoneBook.Rows = 1
          labelClear
      
       End If
  
   End If

End Sub

Private Sub txtTel_GotFocus()
    If tmpflag = True Then
        tmpflag = False
        Exit Sub
    End If
    txtFirstName.Text = ""
    txtLastName.Text = ""
    vsPhoneBook.Row = 0
    labelClear
    LblCount = ""
    vsPhoneBook.Select vsPhoneBook.Row, 4
    vsPhoneBook.Sort = flexSortGenericAscending
    LblFindPhoneBook.Caption = " ·›‰ „‘ —ﬂ —« Ê«—œ ﬂ‰Ìœ  "

End Sub

Private Sub txtTel_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 40 Then
    If vsPhoneBook.Row >= 1 Then
         vsPhoneBook.SetFocus
         If vsPhoneBook.Rows > 2 Then
            vsPhoneBook.Row = 2
         End If
    End If
End If

End Sub

Private Sub TxtTimer_Change()
    clsStation.SrarchInputDelayKeyboard = Val(TxtTimer.Text)
    SetStationSettingFile
End Sub

Private Sub vsPhoneBook_AfterSort(ByVal Col As Long, Order As Integer)
    j = 1
    For i = 1 To vsPhoneBook.Rows - 1
        If i = vsPhoneBook.Rows - 1 Then
            vsPhoneBook.TextMatrix(i, 0) = j
            Exit For
        End If
        vsPhoneBook.TextMatrix(i, 0) = j
        If vsPhoneBook.TextMatrix(i, 3) <> vsPhoneBook.TextMatrix(i + 1, 3) Then
            j = j + 1
        End If
    Next
    vsPhoneBook.MergeCells = flexMergeRestrictRows
    vsPhoneBook.MergeRow(vsPhoneBook.Rows - 1) = True
    vsPhoneBook.MergeCol(0) = True
    vsPhoneBook.MergeCol(0) = True
    
End Sub
Private Sub vsPhoneBook_DblClick()
    If vsPhoneBook.Row > 0 Then
        OKButton_Click
    End If
End Sub

Private Sub Define_Customer()
    
    labelClear
    
    ReDim Parameter(0) As Parameter
    Select Case SearchType
        Case 1
            Parameter(0) = GenerateInputParameter("@nvcLastName", adVarWChar, 50, Left(txtLastName.Text, 50))
            Set Rst = RunParametricStoredProcedure2Rec("Get_tblTotal_tPhoneBook_By_nvcLastName", Parameter)
       Case 2
            Parameter(0) = GenerateInputParameter("@nvcTel", adVarWChar, 30, Left(txtTel.Text, 30))
            Set Rst = RunParametricStoredProcedure2Rec("Get_tblTotal_tPhoneBook_By_nvcTel", Parameter)
         
        Case 3
            Parameter(0) = GenerateInputParameter("@nvcFirstName", adVarWChar, 50, Left(txtFirstName.Text, 50))
            Set Rst = RunParametricStoredProcedure2Rec("Get_tblTotal_tPhoneBook_By_nvcFirstName", Parameter)
        
    End Select
    i = 0
    With vsPhoneBook
        
        .Rows = 1
         Do While Rst.EOF <> True
            .Rows = .Rows + 1
            i = i + 1
            .TextMatrix(i, 0) = i
            .TextMatrix(i, 1) = Rst!intPhoneBookNo
            .TextMatrix(i, 2) = IIf(IsNull(Rst!nvcFirstName), "", Rst!nvcFirstName)
            .TextMatrix(i, 3) = Rst!nvcLastName
            .TextMatrix(i, 4) = Trim(Rst!nvcTelCollection)
            .TextMatrix(i, 5) = IIf(IsNull(Rst!nvcMobile), "", Rst!nvcMobile)
            .TextMatrix(i, 6) = IIf(IsNull(Rst!nvcTelCompany), "", Rst!nvcTelCompany)
            .TextMatrix(i, 7) = IIf(IsNull(Rst!nvcFax), "", Rst!nvcFax)
            .TextMatrix(i, 8) = IIf(IsNull(Rst!nvcEmail), "", Rst!nvcEmail)
            .TextMatrix(i, 9) = IIf(IsNull(Rst!nvcAddress), "", Rst!nvcAddress)
            .TextMatrix(i, 10) = Rst!nvcDate
            If i > Val(txtMaxRecord.Text) Then Exit Do
            Rst.MoveNext
        Loop
    End With
            
        vsPhoneBook.MergeCompare = flexMCTrimNoCase
        'vsPhoneBook.MergeCells = flexMergeRestrictRows
        vsPhoneBook.MergeRow(vsPhoneBook.Rows - 1) = True
        vsPhoneBook.MergeCol(0) = True
        vsPhoneBook.MergeCol(1) = True
        vsPhoneBook.MergeCol(2) = True
        vsPhoneBook.MergeCol(3) = True
    ''''    vsPhoneBook.MergeCol(4) = True
        vsPhoneBook.MergeCol(5) = True
        If i > 0 Then
            LblCount.Caption = " ⁄œ«œ —ﬂÊ—œÂ«   -  " & i
        Else
            LblCount.Caption = ""
        End If
        vsPhoneBook.AutoSizeMode = flexAutoSizeColWidth
        vsPhoneBook.AutoSize 0, vsPhoneBook.Cols - 1
    If vsPhoneBook.ColWidth(3) < 3000 Then
        vsPhoneBook.ColWidth(3) = 3000
    End If
    If vsPhoneBook.ColWidth(4) < 1500 Then
        vsPhoneBook.ColWidth(4) = 1500
    End If
    If vsPhoneBook.ColWidth(5) < 4000 Then
        vsPhoneBook.ColWidth(5) = 4000
    End If
    
    Set Rst = Nothing
    
End Sub


Private Sub labelClear()

End Sub

