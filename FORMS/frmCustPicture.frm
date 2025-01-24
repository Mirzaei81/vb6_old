VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{7AEDC602-D94C-11D1-BB7A-00E0290EA3C9}#1.0#0"; "ResizeKit.ocx"
Object = "{75D4F767-8785-11D3-93AD-0000832EF44D}#2.14#0"; "FAST2001.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmCustPicture 
   ClientHeight    =   7980
   ClientLeft      =   5235
   ClientTop       =   645
   ClientWidth     =   11220
   Icon            =   "frmCustPicture.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7980
   ScaleWidth      =   11220
   Begin VB.TextBox txtPicturePath 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   7800
      MultiLine       =   -1  'True
      TabIndex        =   11
      Text            =   "frmCustPicture.frx":A4C2
      Top             =   1800
      Width           =   1665
   End
   Begin VB.PictureBox PictureBox 
      Height          =   5055
      Left            =   240
      ScaleHeight     =   4995
      ScaleWidth      =   4635
      TabIndex        =   10
      Top             =   2760
      Width           =   4695
   End
   Begin VB.TextBox txtPicNo 
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
      Height          =   465
      Left            =   3240
      RightToLeft     =   -1  'True
      TabIndex        =   8
      Top             =   1800
      Width           =   2175
   End
   Begin VB.TextBox txtName 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFC0C0&
      Enabled         =   0   'False
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
      Left            =   3240
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   960
      Width           =   2895
   End
   Begin VB.TextBox txtMembershipId 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFC0C0&
      Enabled         =   0   'False
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
      Left            =   7800
      RightToLeft     =   -1  'True
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   960
      Width           =   2175
   End
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   8280
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
      Left            =   9120
      Top             =   75
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
   Begin VSFlex7LCtl.VSFlexGrid vsCustPic 
      Height          =   5115
      Left            =   5040
      TabIndex        =   1
      Top             =   2760
      Width           =   6075
      _cx             =   10716
      _cy             =   9022
      _ConvInfo       =   1
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Nazanin"
         Size            =   14.25
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
      SelectionMode   =   0
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
      FormatString    =   $"frmCustPicture.frx":A4CB
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
      OleObjectBlob   =   "frmCustPicture.frx":A54C
      TabIndex        =   2
      Top             =   0
      Width           =   480
   End
   Begin FLWCtrls.FWButton FWBtnpicture 
      Height          =   705
      Left            =   360
      TabIndex        =   7
      Top             =   1800
      Width           =   2115
      _ExtentX        =   3731
      _ExtentY        =   1244
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
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "„”Ì— ›Ê·œ— ⁄ò” „‘ —Ì«‰ œ— ÅÊ‘Â Ã«—Ì"
      BeginProperty Font 
         Name            =   "B Homa"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   1155
      Left            =   9720
      RightToLeft     =   -1  'True
      TabIndex        =   12
      Top             =   1560
      Width           =   1305
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "*‘„«—Â ⁄ﬂ”"
      BeginProperty Font 
         Name            =   "B Homa"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   495
      Left            =   6000
      RightToLeft     =   -1  'True
      TabIndex        =   9
      Top             =   1800
      Width           =   1515
   End
   Begin VB.Label lblName 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "‰«„"
      BeginProperty Font 
         Name            =   "B Homa"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   495
      Left            =   6480
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   960
      Width           =   675
   End
   Begin VB.Label lblCode 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "«‘ —«ò"
      BeginProperty Font 
         Name            =   "B Homa"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   495
      Left            =   10200
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   960
      Width           =   795
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   " ’ÊÌ— „‘ —ﬂÌ‰"
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
      Left            =   3360
      RightToLeft     =   -1  'True
      TabIndex        =   0
      Top             =   0
      Width           =   3855
   End
End
Attribute VB_Name = "frmCustPicture"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private MyFormAddEditMode As EnumAddEditMode
Dim tmpTextMatrix As String
 
Dim Parameter() As Parameter
 
Dim i As Integer
Dim rs As New ADODB.Recordset
Dim SortItem As Integer
Dim intCode As Integer
Dim CustCode As Long
Dim strFileName As String
Dim frmact As Form

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

Public Sub DefaultSetting()
 vsCustPic.ColHidden(4) = True
 txtMembershipId = frmact.txtMembershipId
 txtName = frmact.txtName + " " + frmact.txtFamily
 MyFormAddEditMode = EnumAddEditMode.AddMode
 CustCode = Val(frmact.mvarcode2)
 txtPicNo = ""
 strFileName = ""
 txtPicNo.Tag = 0
 PictureBox.Picture = LoadPicture("")
FWBtnpicture.Enabled = True
txtPicNo.Locked = False

End Sub


Public Sub Edit()
       MyFormAddEditMode = EnumAddEditMode.EditMode
       SetFirstToolBar
       FWBtnpicture.Enabled = True
       txtPicNo.Locked = False
  End Sub

Public Sub Update()
 
    If MyFormAddEditMode = ViewMode Then Exit Sub

        If txtPicNo = "" Or strFileName = "" Then
            frmMsg.fwlblMsg.Caption = "«ÿ·«⁄«  ·«“„ —« Å— ﬂ‰Ìœ"
            frmMsg.fwBtn(0).Visible = False
            frmMsg.fwBtn(1).ButtonType = flwButtonOk
            frmMsg.fwBtn(1).Caption = "ﬁ»Ê·"
            frmMsg.Show vbModal
            Exit Sub
        End If

          ReDim Parameter(1) As Parameter
                Parameter(0) = GenerateInputParameter("@PictureNo", adBigInt, 8, Val(txtPicNo.Text))
                Parameter(1) = GenerateInputParameter("@intserialno", adBigInt, 8, txtPicNo.Tag)
                Set rctmp = RunParametricStoredProcedure2Rec("Get_Check_tCust_pictureNo", Parameter)
                 If rctmp!CountPictureNo = 2 Then
                        frmMsg.fwlblMsg.Caption = " «Ì‰ ‘„«—Â  ﬁ»·« œ— ”Ì” „  ⁄—Ì› ‘œÂ «” . ‘„«—Â  œÌê—Ì  ⁄—Ì› ò‰Ìœ "
                        frmMsg.fwBtn(0).Visible = False
                        frmMsg.fwBtn(1).ButtonType = flwButtonOk
                        frmMsg.fwBtn(1).Caption = "ﬁ»Ê·"
                        frmMsg.Show vbModal
                        txtPicNo.SetFocus
                        Exit Sub
                    End If
            ReDim Parameter(1) As Parameter
                  Parameter(0) = GenerateInputParameter("@PicturePath", adVarChar, 300, strFileName)
                  Parameter(1) = GenerateInputParameter("@intserialno", adBigInt, 8, txtPicNo.Tag)
                Set rctmp = RunParametricStoredProcedure2Rec("Get_Check_tCust_picturePath", Parameter)
                 If rctmp!CountPictureNo = 2 Then
                        frmMsg.fwlblMsg.Caption = " «Ì‰ ⁄ﬂ”  ﬁ»·« œ— ”Ì” „ « Œ«» ‘œÂ «” . ⁄ﬂ”  œÌê—Ì «‰ Œ«» ò‰Ìœ "
                        frmMsg.fwBtn(0).Visible = False
                        frmMsg.fwBtn(1).ButtonType = flwButtonOk
                        frmMsg.fwBtn(1).Caption = "ﬁ»Ê·"
                        frmMsg.Show vbModal
                        txtPicNo.SetFocus
                        Exit Sub
                    End If


        Select Case MyFormAddEditMode
            Case AddMode
                  
                ReDim Parameter(3) As Parameter
                
                Parameter(0) = GenerateInputParameter("@Code", adInteger, 4, CustCode)
                Parameter(1) = GenerateInputParameter("@PictureNo", adBigInt, 8, Val(txtPicNo.Text))
                Parameter(2) = GenerateInputParameter("@PicturePath", adVarWChar, 300, strFileName)
                Parameter(3) = GenerateOutputParameter("@Result", adInteger, 4)
               
               
             Dim Resault As Integer
                Resault = RunParametricStoredProcedure("InsertCustomerPicture", Parameter)
                If Resault > 0 Then
                    frmMsg.fwlblMsg.Caption = "À»  «ÿ·«⁄«  ⁄ﬂ” ÃœÌœ »« „Ê›ﬁÌ  «‰Ã«„ ‘œ"
                    frmMsg.fwBtn(0).Visible = False
                    frmMsg.fwBtn(1).ButtonType = flwButtonOk
                    frmMsg.fwBtn(1).Caption = "ﬁ»Ê·"
                    frmMsg.Show vbModal
             
                Else
                    frmMsg.fwlblMsg.Caption = "À»  «‰Ã«„ ‰‘œ"
                    frmMsg.fwBtn(0).Visible = False
                    frmMsg.fwBtn(1).ButtonType = flwButtonOk
                    frmMsg.fwBtn(1).Caption = "ﬁ»Ê·"
                    frmMsg.Show vbModal
                    Exit Sub
                End If
               
            Case EditMode

                ReDim Parameter(3) As Parameter
                Parameter(0) = GenerateInputParameter("@PictureNo", adBigInt, 8, Val(txtPicNo.Text))
                Parameter(1) = GenerateInputParameter("@PicturePath", adVarWChar, 300, strFileName)
                Parameter(2) = GenerateInputParameter("@intserial", adInteger, 4, txtPicNo.Tag)
                Parameter(3) = GenerateOutputParameter("@Updated", adInteger, 4)

                Dim Updated As Long
                Updated = RunParametricStoredProcedure("Update_TCust_Picture", Parameter)
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

            End Select

 
                
                MyFormAddEditMode = AddMode
                DefaultSetting
                SetFirstToolBar
                FillvsCustPic

End Sub


Public Sub Cancel()
    MyFormAddEditMode = EnumAddEditMode.AddMode
    SetFirstToolBar
    DefaultSetting
  
End Sub
Private Sub CmbGoodlevel1_Click()
FillvsCustPic
End Sub

Private Sub Form_Activate()
    
    VarActForm = Me.Name
    MyFormAddEditMode = AddMode
    SortItem = 1    'Code Sort

End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call PresetScreenSaver
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call PresetScreenSaver
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
          
    CenterTop Me
    VarActForm = Me.Name
    MyFormAddEditMode = AddMode
    
    Dim varForm As Form
    For Each varForm In Forms
        If LCase(varForm.Name) = "frmcust" Then
            Set frmact = varForm
            Exit For
        End If
    Next
    frmact.Hide
    
    DefaultSetting
    SetFirstToolBar
    FillvsCustPic
 
      
    formloadFlag = False
    Me.left = Val(GetSetting(strMainKey, Me.Name, "Left"))
    If Val(GetSetting(strMainKey, Me.Name, "Height")) > 0 Then
        Me.Height = Val(GetSetting(strMainKey, Me.Name, "Height"))
    End If
    If Val(GetSetting(strMainKey, Me.Name, "Width")) > 0 Then
        Me.Width = Val(GetSetting(strMainKey, Me.Name, "Width"))
    End If
    Me.top = Val(GetSetting(strMainKey, Me.Name, "Top"))
    formloadFlag = True
    
End Sub
Private Sub Form_Unload(Cancel As Integer)
    
    If PosConnection.State = adStateOpen Then PosConnection.Close: Set PosConnection = Nothing
    frmact.Show
    SaveSetting strMainKey, Me.Name, "Left", Me.left
    SaveSetting strMainKey, Me.Name, "Top", Me.top

    
End Sub


Private Sub FWBtnpicture_Click()
                 
    
    With Cdlg
         .CancelError = True
         .Filter = "Pictures (*.bmp;*.ico;*.gif;*.jpg;*.jpeg)|*.bmp;*.ico;*.gif;*.jpg;*.jpeg"
         .DialogTitle = "Picture Search"
         .InitDir = App.Path & "\Image"
         On Error GoTo NoFileOpened
         .ShowOpen
         strFileName = .FileName
NoFileOpened:
    End With
    
    
    PictureBox.Picture = LoadPicture(Trim(strFileName))
   
End Sub

Private Sub ResizeKit1_ExitResize(ByVal XScale As Double, ByVal YScale As Double)


    If formloadFlag = True Then
        SaveSetting strMainKey, Me.Name, "Height", Me.Height
        SaveSetting strMainKey, Me.Name, "Width", Me.Width
    End If

End Sub
Private Sub FillvsCustPic()
    
    Dim Rst As New ADODB.Recordset
    ReDim Parameter(0) As Parameter
    
    
   
    Parameter(0) = GenerateInputParameter("@code", adInteger, 4, CustCode)
    Set Rst = RunParametricStoredProcedure2Rec("Get_All_TCust_Picture", Parameter)
    
    With vsCustPic
        .Rows = 1
        i = 0
        While Rst.EOF <> True
            .Rows = .Rows + 1
            i = .Rows - 1
            .TextMatrix(i, 0) = i
            .TextMatrix(i, 1) = Rst!PictureNo
            .TextMatrix(i, 2) = Rst!AddDate
            .TextMatrix(i, 3) = Rst!AddTime
            .TextMatrix(i, 4) = Rst!Intserial
           
       
            Rst.MoveNext
        Wend
    End With
    Set Rst = Nothing
    
    
End Sub
Sub GetDataDetail()
    
    txtDiscount = ""
    txtFromSumPrice = ""
    txtToSumPrice = ""
    txtPicNo.Tag = 0
    Dim TempStr As String
    
    ReDim Parameter(0) As Parameter
    Parameter(0) = GenerateInputParameter("@intCode", adInteger, 4, intCode)
    Set rctmp = RunParametricStoredProcedure2Rec("Get_TCust_Picture", Parameter)
    Dim ii As Integer
    If Not (rctmp.BOF Or rctmp.EOF) Then
        
           txtPicNo.Text = rctmp!PictureNo
           txtPicNo.Tag = rctmp!Intserial
           strFileName = rctmp!PicturePath
          
          '' On Error Resume Next
          On Error GoTo ErrHandler
          PictureBox.Picture = LoadPicture(rctmp!PicturePath)
          
    End If
    
ErrHandler:
 If err.Number = 53 Then
 
        PictureBox.Picture = LoadPicture("")
        frmMsg.fwlblMsg.Caption = "⁄ﬂ” „Ê—œ ‰Ÿ— Å«ﬂ ‘œÂ «” "
        frmMsg.fwBtn(0).Caption = "ﬁ»Ê·"

        frmMsg.fwBtn(1).Visible = False
        frmMsg.fwBtn(0).ButtonType = flwButtonOk
        frmMsg.Show vbModal
 End If
    If rctmp.State = adStateOpen Then rctmp.Close
    
    
End Sub

Private Sub vsCustPic_Click()
    If vsCustPic.Row = 0 Then Exit Sub
    intCode = vsCustPic.TextMatrix(vsCustPic.Row, 1)
    GetDataDetail
    MyFormAddEditMode = ViewMode
    SetFirstToolBar
    FWBtnpicture.Enabled = False
    txtPicNo.Locked = True
    HeaderLabel Val(MyFormAddEditMode), Me.fwlblMode
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
    Parameter(0) = GenerateInputParameter("@intserial", adBigInt, 8, txtPicNo.Tag)
    Parameter(1) = GenerateOutputParameter("@Result", adInteger, 4)
    Dim Result As Integer
    Result = RunParametricStoredProcedure("Delete_TCust_Picture", Parameter)
    
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
    FillvsCustPic
   
End Sub
