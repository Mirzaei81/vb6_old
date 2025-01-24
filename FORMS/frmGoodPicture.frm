VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{7AEDC602-D94C-11D1-BB7A-00E0290EA3C9}#1.0#0"; "ResizeKit.ocx"
Object = "{75D4F767-8785-11D3-93AD-0000832EF44D}#2.14#0"; "FAST2001.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmGoodPicture 
   ClientHeight    =   9105
   ClientLeft      =   5235
   ClientTop       =   645
   ClientWidth     =   13725
   Icon            =   "frmGoodPicture.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9105
   ScaleWidth      =   13725
   Begin VB.TextBox txtBarcode 
      Alignment       =   1  'Right Justify
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
      Left            =   4680
      MultiLine       =   -1  'True
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   1920
      Width           =   2145
   End
   Begin VB.ListBox lstGoodLevel2 
      BeginProperty Font 
         Name            =   "Nazanin"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2010
      Left            =   7680
      RightToLeft     =   -1  'True
      Style           =   1  'Checkbox
      TabIndex        =   5
      Top             =   1080
      Width           =   2745
   End
   Begin VB.ListBox lstGoodLevel1 
      BeginProperty Font 
         Name            =   "Nazanin"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2010
      Left            =   10800
      RightToLeft     =   -1  'True
      Style           =   1  'Checkbox
      TabIndex        =   4
      Top             =   1080
      Width           =   2775
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   9360
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowGroupTree=   -1  'True
      WindowAllowDrillDown=   -1  'True
      WindowShowCloseBtn=   -1  'True
      WindowShowSearchBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin FLWCtrls.FWLabel3D fwlblMode 
      Height          =   495
      Left            =   12000
      Top             =   120
      Width           =   1545
      _ExtentX        =   2725
      _ExtentY        =   873
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor1      =   32896
      ForeColor2      =   128
      BackColor       =   9412754
      Caption         =   "„—Ê—"
      Alignment       =   2
   End
   Begin VSFlex7LCtl.VSFlexGrid vsGood 
      Height          =   5715
      Left            =   6960
      TabIndex        =   1
      Top             =   3360
      Width           =   6675
      _cx             =   11774
      _cy             =   10081
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
      BackColor       =   16777215
      ForeColor       =   -2147483640
      BackColorFixed  =   16761024
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483636
      BackColorAlternate=   16777215
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
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   5
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   500
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmGoodPicture.frx":A4C2
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
      ExplorerBar     =   3
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
      OleObjectBlob   =   "frmGoodPicture.frx":A53B
      TabIndex        =   2
      Top             =   0
      Width           =   480
   End
   Begin FLWCtrls.FWButton FWBtnpicture 
      Height          =   585
      Left            =   465
      TabIndex        =   3
      Top             =   2250
      Width           =   2115
      _ExtentX        =   3731
      _ExtentY        =   1032
      ButtonType      =   8
      Caption         =   " «‰ Œ«»  ’ÊÌ— "
      BackColor       =   49152
      ForeColor       =   16384
      FontName        =   "B Homa"
      FontBold        =   -1  'True
      FontSize        =   9.75
      Alignment       =   1
   End
   Begin MSComDlg.CommonDialog Cdlg 
      Left            =   720
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Image Image1 
      Height          =   5655
      Left            =   105
      Stretch         =   -1  'True
      Top             =   3360
      Width           =   6735
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "»«—òœ"
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
      Left            =   5880
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   1200
      Width           =   825
   End
   Begin VB.Label lblGoodLevel2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H000080FF&
      BackStyle       =   0  'Transparent
      Caption         =   "ê—ÊÂ ›—⁄Ì ò«·«Â«"
      BeginProperty Font 
         Name            =   "Titr"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   435
      Left            =   7440
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   600
      Width           =   2745
   End
   Begin VB.Label lblGoodLevel1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H000040C0&
      BackStyle       =   0  'Transparent
      Caption         =   "ê—ÊÂ «’·Ì ò«·«Â«"
      BeginProperty Font 
         Name            =   "Titr"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   435
      Left            =   10920
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   720
      Width           =   2655
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "«Œ ’«’  ’ÊÌ— »Â ﬂ«·«"
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
      Left            =   5160
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   0
      Width           =   3855
   End
End
Attribute VB_Name = "frmGoodPicture"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private MyFormAddEditMode As EnumAddEditMode
Dim tmpTextMatrix As String
Dim Parameter() As Parameter
Dim clsDate As New clsDate
Dim i As Integer
Dim rs As New ADODB.Recordset
Dim Rst As New ADODB.Recordset
Dim SortItem As Integer
Dim intCode As Long
Dim CustCode As Integer
Dim strFileName As String
Dim strStream
Public Sub ExitForm()

    Unload Me
    
End Sub

Public Sub SetFirstToolBar()

    Dim i As Integer
    
    AllButton vbOff, True
    mdifrm.Toolbar1.Buttons(13).Enabled = False   'Find
    
    mdifrm.Toolbar1.Buttons(15).Enabled = False  'printing
    mdifrm.Toolbar1.Buttons(23).Enabled = True
    mdifrm.Toolbar1.Buttons(24).Enabled = True
    mdifrm.Toolbar1.Buttons(25).Enabled = True
    mdifrm.Toolbar1.Buttons(26).Enabled = True
    mdifrm.Toolbar1.Buttons(27).Enabled = True
    
    Select Case MyFormAddEditMode
    
        Case EnumAddEditMode.ViewMode
        
            For i = 6 To 9
                mdifrm.Toolbar1.Buttons(i).Enabled = True
            Next i
''            vsGood.Editable = flexEDNone
           mdifrm.Toolbar1.Buttons(10).Enabled = True
        Case EnumAddEditMode.AddMode
        
            mdifrm.Toolbar1.Buttons(8).Enabled = True 'enter key
            mdifrm.Toolbar1.Buttons(9).Enabled = True 'cancel key

            
        Case EnumAddEditMode.EditMode
                    
            mdifrm.Toolbar1.Buttons(8).Enabled = True 'enter key
            mdifrm.Toolbar1.Buttons(9).Enabled = True 'cancel key

    End Select
    
    HeaderLabel Val(MyFormAddEditMode), fwlblMode
    
End Sub

''''Public Sub DefaultSetting()
'''' vsGood.ColHidden(4) = True
'''' txtMembershipId = frmCust.mvarMemberShipId
'''' txtName = frmCust.mvarCustName
'''' txtFamily = frmCust.mvarCustFamily
'''' MyFormAddEditMode = EnumAddEditMode.AddMode
'''' CustCode = Val(frmCust.mvarcode2)
'''' txtPicNo = ""
'''' strFileName = ""
''''
'''' Image1.Picture = LoadPicture("")
''''FWBtnpicture.Enabled = True
''''
''''
''''End Sub


Public Sub Edit()
       MyFormAddEditMode = EnumAddEditMode.EditMode
       SetFirstToolBar
       FWBtnpicture.Enabled = True
    
  End Sub

Public Sub Update()
    If MyFormAddEditMode = ViewMode Then Exit Sub

    If intCode = 0 Then   'Or strFileName = ""
        frmMsg.fwlblMsg.Caption = "«ÿ·«⁄«  ·«“„ —« Å— ﬂ‰Ìœ"
        frmMsg.fwBtn(0).Visible = False
        frmMsg.fwBtn(1).ButtonType = flwButtonOk
        frmMsg.fwBtn(1).Caption = "ﬁ»Ê·"
        frmMsg.Show vbModal
        Exit Sub
    End If
     
    Set strStream = New ADODB.Stream
    strStream.Type = adTypeBinary
    strStream.Open
    If strFileName <> "" Then
        strStream.LoadFromFile strFileName
    End If
'    rs.Fields("***YourImageField***").Value = strStream.Read
    
'    Image1.Picture = LoadPicture(strFileName)
'    SavePictureToDB = True
                
     ReDim Parameter(3) As Parameter
     Parameter(0) = GenerateInputParameter("@PicturePath", adVarWChar, 300, strFileName)
     Parameter(1) = GenerateInputParameter("@GoodCode", adInteger, 4, intCode)
     Parameter(2) = GenerateInputParameter("@Picture", adLongVarBinary, strStream.Size + 1, strStream.Read)
     Parameter(3) = GenerateOutputParameter("@Updated", adInteger, 4)

     Dim Updated As Long
     Updated = RunParametricStoredProcedure("Update_tblTotal_GoodPic", Parameter)
     If Updated = 1 Then
         frmMsg.fwlblMsg.Caption = " €ÌÌ—«  »« „Ê›ﬁÌ  «‰Ã«„ ‘œ"
         frmMsg.fwBtn(0).Visible = False
         frmMsg.fwBtn(1).ButtonType = flwButtonOk
         frmMsg.fwBtn(1).Caption = "ﬁ»Ê·"
         frmMsg.Show vbModal
    Else
         frmMsg.fwlblMsg.Caption = " €ÌÌ—«  «‰Ã«„ ‰‘œ"
         frmMsg.fwBtn(0).Visible = False
         frmMsg.fwBtn(1).ButtonType = flwButtonOk
         frmMsg.fwBtn(1).Caption = "ﬁ»Ê·"
         frmMsg.Show vbModal
         Exit Sub
     End If
     MyFormAddEditMode = AddMode
     DefaultSetting
     SetFirstToolBar
     FillvsGood

End Sub


Public Sub Cancel()
    MyFormAddEditMode = EnumAddEditMode.AddMode
    SetFirstToolBar
    DefaultSetting
  
End Sub
Private Sub CmbGoodlevel1_Click()
FillvsGood
End Sub

Private Sub Form_Activate()
    
   
    VarActForm = Me.Name
    MyFormAddEditMode = AddMode
    SortItem = 1    'Code Sort
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
        ShowDisMessage "«” ›«œÂ «“ ⁄ò” ò«·«Â« œ— ‰”ŒÂ Â«Ì „ Ê”ÿ Ê »«·« — «„ﬂ«‰ Å–Ì— «” ", 1500
        Unload Me
        Exit Sub
    End If
    CenterTop Me
    VarActForm = Me.Name
    MyFormAddEditMode = AddMode
    DefaultSetting
    SetFirstToolBar
    ''FillvsGood
 
      
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
    VarActForm = ""
    If PosConnection.State = adStateOpen Then PosConnection.Close: Set PosConnection = Nothing
   
    SaveSetting strMainKey, Me.Name, "Left", Me.Left
    SaveSetting strMainKey, Me.Name, "Top", Me.Top
    
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call PresetScreenSaver
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call PresetScreenSaver
End Sub


Private Sub FWBtnpicture_Click()
                 
 On Error GoTo NoFileOpened
 
    With Cdlg
         .CancelError = True
         .Filter = "Pictures (*.bmp;*.ico;*.gif;*.jpg;*.jpeg;*.png)|*.bmp;*.ico;*.gif;*.jpg;*.jpeg;*.png"
         .DialogTitle = "Picture Search"
         .InitDir = App.Path & "\Image"
         On Error GoTo NoFileOpened
         .ShowOpen
         strFileName = .Filename
    End With
    
'    Image1.Picture = LoadPicture(strFileName)

     Dim Token As Long
     Dim c
        
     c = Me.BackColor
        
     If c < 0 Then c = GetSysColor(c - &H80000000)
        
     Token = InitGDIPlus
        
    ' Picture1(0).Picture = LoadPictureGDIPlus(App.Path & "\1.png", , , vbWhite)
    ' Picture1(1).Picture = LoadPictureGDIPlus(App.Path & "\1.png", , , vbCyan)
    ' Picture1(2).Picture = LoadPictureGDIPlus(App.Path & "\1.png", , , vbGreen)
     Image1.Picture = LoadPictureGDIPlus(strFileName, , , c)
        
     FreeGDIPlus Token

Exit Sub
   
NoFileOpened:
    strFileName = ""
    ShowDisMessage err.Description, 1000
End Sub


Private Sub lstGoodLevel1_Click()

    FillLstGoodLevel2
End Sub


Private Sub lstGoodLevel2_ItemCheck(Item As Integer)
    
    FillvsGood
End Sub




Private Sub ResizeKit1_ExitResize(ByVal XScale As Double, ByVal YScale As Double)


    If formloadFlag = True Then
        SaveSetting strMainKey, Me.Name, "Height", Me.Height
        SaveSetting strMainKey, Me.Name, "Width", Me.Width
    End If

End Sub
Private Sub FillvsGood()
    MyFormAddEditMode = EnumAddEditMode.ViewMode 'VIEW Mode
    SetFirstToolBar
    
    vsGood.Rows = 1
    If lstGoodLevel1.ListCount < 1 Then Exit Sub
    If lstGoodLevel2.ListCount < 1 Then Exit Sub
    
    Dim i As Integer
    Dim j As Integer
    Dim intSelectedLevel1 As Integer
    Dim intSelectedLevel2 As Integer
    Dim strSelectedLevels As String

    
    intSelectedLevel1 = -1
    intSelectedLevel2 = -1
    
    For i = 0 To lstGoodLevel1.ListCount - 1
        If lstGoodLevel1.Selected(i) = True Then
            intSelectedLevel1 = i
        End If
    Next i
    
    strSelectedLevels = ""
    For i = 0 To lstGoodLevel2.ListCount - 1
        If lstGoodLevel2.Selected(i) = True Then
            intSelectedLevel2 = i
            strSelectedLevels = strSelectedLevels + "," + CStr(lstGoodLevel2.ItemData(i))
        End If
    Next i
    
    If Rst.State <> 0 Then Rst.Close
    
    If intSelectedLevel1 <> -1 And intSelectedLevel2 <> -1 Then
        
        strSelectedLevels = Right(strSelectedLevels, Len(strSelectedLevels) - 1)
        ReDim Parameter(2) As Parameter
        Parameter(0) = GenerateInputParameter("@Level1", adInteger, 4, lstGoodLevel1.ItemData(intSelectedLevel1))
        Parameter(1) = GenerateInputParameter("@strSelectedLevels", adVarWChar, 4000, strSelectedLevels)
        Parameter(2) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
        Set Rst = RunParametricStoredProcedure2Rec("Get_Good_In_Levels", Parameter)
    
    ElseIf intSelectedLevel1 <> -1 And intSelectedLevel2 = -1 Then  'Or intSelectedLevel2 = -1
        ReDim Parameter(1) As Parameter
        Parameter(0) = GenerateInputParameter("@GoodLevel1Code", adInteger, 4, lstGoodLevel1.ItemData(intSelectedLevel1))
        Parameter(1) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
        Set Rst = RunParametricStoredProcedure2Rec("GetVw_GoodInfo", Parameter)
    
    Else
        ReDim Parameter(0) As Parameter
        Parameter(0) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
        Set Rst = RunParametricStoredProcedure2Rec("GetVwGoodInfo", Parameter)
    
    End If
    
    If (Rst.EOF = True And Rst.BOF = True) Then Exit Sub
    
    With vsGood
        
        
        i = 1
        
        MousePointer = 11
        While Rst.EOF = False
            .Rows = .Rows + 1
            .TextMatrix(i, 0) = i
            .TextMatrix(i, 1) = Rst.Fields("Code").Value
            .TextMatrix(i, 2) = Rst.Fields("GoodName").Value
            .TextMatrix(i, 3) = Rst.Fields("BarCode").Value
            .TextMatrix(i, 4) = Rst.Fields("HavePic").Value
            i = i + 1
            Rst.MoveNext
            
        Wend
        MousePointer = 0
        Set Rst = Nothing
        Select Case clsStation.Language
            Case 0
                .ColAlignment(-1) = flexAlignRightCenter
            Case 1
                .ColAlignment(-1) = flexAlignLeftCenter
        End Select

        .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter
        .Cell(flexcpAlignment, 1, 4, .Rows - 1, .Cols - 1) = flexAlignCenterCenter
        .AutoSizeMode = flexAutoSizeColWidth
'        .AutoSize 0, .Cols - 1
        .AutoSize 1, 4

''        If .ColWidth(2) > 3500 Then .ColWidth(2) = 3500
''        If .ColWidth(3) > 3000 Then .ColWidth(3) = 3000
    End With
     vsGood_RowColChange
End Sub
    

Sub GetDataDetail()
    
    On Error GoTo err
    
    Image1.Picture = LoadPicture("")
    FWBtnpicture.Tag = 0
    Dim TempStr As String
    Dim Token As Long
    Dim c
    
    ReDim Parameter(0) As Parameter
    Parameter(0) = GenerateInputParameter("@intCode", adInteger, 4, intCode)
    Set rctmp = RunParametricStoredProcedure2Rec("Get_tGood_Picture", Parameter)
    Dim ii As Integer
    If Not (rctmp.BOF Or rctmp.EOF) Then
        
           FWBtnpicture.Tag = rctmp!GoodCode
           strFileName = rctmp!PicturePath
          
          '' On Error Resume Next
    
'            If IsNull(rctmp.Fields("Picture").Value) Then
                    
                 c = Me.BackColor
                    
                 If c < 0 Then c = GetSysColor(c - &H80000000)
                    
                 Token = InitGDIPlus
                    
                ' Picture1(0).Picture = LoadPictureGDIPlus(App.Path & "\1.png", , , vbWhite)
                ' Picture1(1).Picture = LoadPictureGDIPlus(App.Path & "\1.png", , , vbCyan)
                ' Picture1(2).Picture = LoadPictureGDIPlus(App.Path & "\1.png", , , vbGreen)
                 Image1.Picture = LoadPictureGDIPlus(strFileName, , , c)
                    
                 FreeGDIPlus Token
'            Else
'                Set strStream = New ADODB.Stream
'                strStream.Type = adTypeBinary
'                strStream.Open
'                strStream.Write rctmp.Fields("Picture").Value
'                strStream.SaveToFile "C:\Temp.bmp", adSaveCreateOverWrite
'                Image1.Picture = LoadPicture("C:\Temp.bmp")
'                Kill ("C:\Temp.bmp")
'    '            LoadPictureFromDB = True
'                Set strStream = Nothing
'            End If
    Else
           Image1.Picture = LoadPicture("")
          
    End If
    
err:
 If err.Number = 53 Then
 
        Image1.Picture = LoadPicture("")
        frmMsg.fwlblMsg.Caption = "⁄ﬂ” „Ê—œ ‰Ÿ— Å«ﬂ ‘œÂ «” "
        frmMsg.fwBtn(0).Caption = "ﬁ»Ê·"

        frmMsg.fwBtn(1).Visible = False
        frmMsg.fwBtn(0).ButtonType = flwButtonOk
        frmMsg.Show vbModal
 End If
    rctmp.Close
    
    
End Sub


Private Sub txtBarcode_Change()
    If Right(txtBarcode.Text, 1) = "/" Then
        txtBarcode.Text = Left(txtBarcode.Text, Len(txtBarcode.Text) - 1)
    ElseIf Left(txtBarcode.Text, 1) = "/" Then
        txtBarcode.Text = Right(txtBarcode.Text, Len(txtBarcode.Text) - 1)
    End If
    If Len(txtBarcode.Text) > 2 Then
        If Asc(Mid(txtBarcode.Text, Len(txtBarcode.Text) - 1, 1)) = 13 Then
            txtBarcode.Text = Left(txtBarcode.Text, Len(txtBarcode.Text) - 2)
        End If
    End If
    
    i = vsGood.FindRow(Trim(txtBarcode.Text), 1, 3, True, True)
    If i > 0 Then
        vsGood.Row = i
        vsGood.ShowCell i, 3
    Else
        vsGood.Row = 0
        vsGood.ShowCell 0, 0
    End If
    
End Sub


Private Sub vsGood_Click()
''    If vsGood.Row = 0 Then Exit Sub
''    intCode = Val(vsGood.TextMatrix(vsGood.Row, 1))
''    GetDataDetail
''    MyFormAddEditMode = ViewMode
''    SetFirstToolbar
''    FWBtnpicture.Enabled = False
''
''    HeaderLabel Val(MyFormAddEditMode), Me.fwlblMode
End Sub
Public Sub Delete()
    
            frmMsg.fwlblMsg.Caption = "¬Ì« „ÿ„∆‰Ìœ „Ì ŒÊ«ÂÌœ ⁄ﬂ” „—œ ‰Ÿ— —« Õ–› ﬂ‰Ìœø"
            frmMsg.fwBtn(0).Caption = "»·Ì"
            frmMsg.fwBtn(1).Caption = "ŒÌ—"
       
    
    frmMsg.Show vbModal
    
    If modgl.mvarMsgIdx = vbNo Then
        Exit Sub
    End If
    ReDim Parameter(1) As Parameter
    Parameter(0) = GenerateInputParameter("@GoodCode", adBigInt, 8, intCode)
    Parameter(1) = GenerateOutputParameter("@Result", adInteger, 4)
    Dim Result As Integer
    Result = RunParametricStoredProcedure("Delete_tblTotal_GoodPic", Parameter)
    
    If Result = 0 Then
    
       
                frmMsg.fwlblMsg.Caption = "„‘ò·Ì œ—Õ–› «Ì‰ ⁄ﬂ” ÊÃÊœ œ«—œ ‘„« ‰„Ì  Ê«‰Ìœ «Ì‰ ⁄ﬂ” —« Õ–› ò‰Ìœ"
                frmMsg.fwBtn(0).Caption = "ﬁ»Ê·"

        frmMsg.fwBtn(1).Visible = False
        frmMsg.fwBtn(0).ButtonType = flwButtonOk
        frmMsg.Show vbModal
    
    Else
    
         frmMsg.fwlblMsg.Caption = "‘„« Ìò ⁄ﬂ” —« Õ–› ò—œÌœ"
         frmMsg.fwBtn(0).Caption = "ﬁ»Ê·"
         
        
        frmMsg.fwBtn(1).Visible = False
        frmMsg.fwBtn(0).ButtonType = flwButtonOk
        frmMsg.Show vbModal
        
    End If
    
    MyFormAddEditMode = AddMode
    DefaultSetting
    SetFirstToolBar
    FillvsGood
   
End Sub
Public Sub DefaultSetting()

Dim Obj As Object




            
            For Each Obj In Me
                On Error Resume Next
                    Obj.RightToLeft = True
                On Error GoTo 0
            Next Obj
            
    
    With vsGood
    
        .Cols = 5
                .TextMatrix(0, 0) = "—œÌ›"
                .TextMatrix(0, 1) = "òœ"
                .TextMatrix(0, 2) = "‰«„ ò«·«"
                .TextMatrix(0, 3) = "»«—òœ"
                .TextMatrix(0, 4) = " ’ÊÌ—"
               
            
        
        
        .ColDataType(4) = flexDTBoolean
''        .ColSort(5) = flexSortNumericAscending + flexSortNumericDescending
        .FocusRect = flexFocusHeavy
       ' .ColHidden(1) = True
        .AutoSizeMode = flexAutoSizeColWidth
        .AutoSize 0, .Cols - 2
'''        If .ColWidth(2) > 3500 Then .ColWidth(2) = 3500
        .AutoSearch = flexSearchFromCursor
        .ScrollBars = flexScrollBarBoth
        .AllowUserResizing = flexResizeColumns
               
    End With
 
    
    FillLstGoodLevel1
    
    Set rctmp = Nothing
  
    SetFirstToolBar

End Sub
Public Sub FillLstGoodLevel1() ' it fills the lstGoodLevel1 using table tgoodlevel1
    Dim Rst As New ADODB.Recordset
    
    lstGoodLevel1.Clear
    lstGoodLevel2.Clear
    vsGood.Rows = 1
    
    MyFormAddEditMode = EnumAddEditMode.ViewMode
    SetFirstToolBar
    
    ReDim Parameter(0) As Parameter
    Parameter(0) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
    Set Rst = RunParametricStoredProcedure2Rec("Get_tGoodLevel1", Parameter)
        
    If (Rst.EOF = True And Rst.BOF = True) Then
        Exit Sub
    End If
    
    While Rst.EOF = False
        lstGoodLevel1.AddItem Rst.Fields("Description")
        lstGoodLevel1.ItemData(lstGoodLevel1.ListCount - 1) = Rst.Fields("Code")
        Rst.MoveNext
    Wend
    
    
    lstGoodLevel1.ListIndex = 0
    FillLstGoodLevel2
    Set Rst = Nothing
End Sub

Public Sub FillLstGoodLevel2() ' it fills the lstGoodLevel2 using table tgoodlevel2

    Dim Rst As New ADODB.Recordset
    Dim i As Integer
    Dim intSelectedItem As Integer
        
    MyFormAddEditMode = EnumAddEditMode.ViewMode
    SetFirstToolBar
    
    lstGoodLevel2.Clear
    vsGood.Rows = 1
    
    
    If lstGoodLevel1.ListIndex = -1 Then
        Set Rst = Nothing
        Exit Sub
    Else
        ReDim Parameter(0) As Parameter
        Parameter(0) = GenerateInputParameter("@Code", adInteger, 4, lstGoodLevel1.ItemData(lstGoodLevel1.ListIndex))
        
        If Rst.State <> 0 Then Rst.Close
        Set Rst = RunParametricStoredProcedure2Rec("FillLstGoodLevel2", Parameter)
        If (Rst.EOF = True And Rst.BOF = True) Then
            Set Rst = Nothing
            Exit Sub
        End If
       ' rst.moveFirst
        While Rst.EOF = False
            Select Case clsStation.Language
                Case 0
                    lstGoodLevel2.AddItem Rst.Fields("Description")
                Case 1
                    lstGoodLevel2.AddItem Rst.Fields("LatinDescription")
            End Select
            
            lstGoodLevel2.ItemData(lstGoodLevel2.ListCount - 1) = Rst.Fields("Code")
            Rst.MoveNext
        Wend
        ReDim Parameter(1) As Parameter
        Parameter(0) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
        Parameter(1) = GenerateInputParameter("@CurrentItem", adInteger, 4, lstGoodLevel1.ItemData(lstGoodLevel1.ListIndex))
        Set rctmp = RunParametricStoredProcedure2Rec("GetGoodLevel2_Description", Parameter)
        
        ''vsGood.ColComboList(15) = vsGood.BuildComboList(rctmp, "Description", "Code")
        
        Set Rst = Nothing
        lstGoodLevel2.ListIndex = 0
        FillvsGood
        
    End If
    
End Sub


Private Sub vsGood_RowColChange()
    intCode = Val(vsGood.TextMatrix(vsGood.Row, 1))
    GetDataDetail
    MyFormAddEditMode = ViewMode
    SetFirstToolBar
    FWBtnpicture.Enabled = False
   
    HeaderLabel Val(MyFormAddEditMode), Me.fwlblMode
End Sub
