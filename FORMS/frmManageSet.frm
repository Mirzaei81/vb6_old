VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{7AEDC602-D94C-11D1-BB7A-00E0290EA3C9}#1.0#0"; "ResizeKit.ocx"
Object = "{75D4F767-8785-11D3-93AD-0000832EF44D}#2.14#0"; "FAST2001.ocx"
Begin VB.Form frmManageSet 
   ClientHeight    =   7035
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   8445
   Icon            =   "frmManageSet.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7035
   ScaleWidth      =   8445
   Begin RESIZEKITLibCtl.ResizeKit ResizeKit1 
      Height          =   480
      Left            =   0
      OleObjectBlob   =   "frmManageSet.frx":A4C2
      TabIndex        =   3
      Top             =   0
      Width           =   480
   End
   Begin FLWCtrls.FWLabel3D fwlblMode 
      Height          =   495
      Left            =   7080
      Top             =   0
      Width           =   1380
      _ExtentX        =   2434
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
   Begin VB.ComboBox comboTables 
      BeginProperty Font 
         Name            =   "Nazanin"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   480
      ItemData        =   "frmManageSet.frx":A548
      Left            =   5880
      List            =   "frmManageSet.frx":A54A
      RightToLeft     =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1200
      Width           =   2265
   End
   Begin VSFlex7LCtl.VSFlexGrid vsFlex 
      Height          =   4395
      Left            =   2280
      TabIndex        =   0
      Top             =   2040
      Width           =   5955
      _cx             =   10504
      _cy             =   7752
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
      AllowUserResizing=   3
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   5
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
      AutoResize      =   0   'False
      AutoSizeMode    =   0
      AutoSearch      =   1
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
   Begin FLWCtrls.FWLabel fwlblPartition 
      Height          =   495
      Left            =   -2280
      Top             =   0
      Width           =   12810
      _ExtentX        =   22595
      _ExtentY        =   873
      Enabled         =   -1  'True
      Caption         =   " ⁄«—Ì› Å«ÌÂ"
      FirstColor      =   9412754
      SecondColor     =   14215660
      Angle           =   0
      ForeColor       =   7362318
      BackColor       =   12640511
      FontName        =   "B Homa"
      FontSize        =   14.25
      Alignment       =   2
      Picture         =   "frmManageSet.frx":A54C
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "‰«„ ÃœÊ·"
      BeginProperty Font 
         Name            =   "Nazanin"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   4680
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   1200
      Width           =   975
   End
End
Attribute VB_Name = "frmManageSet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private MyFormAddEditMode As EnumAddEditMode
Dim Parameter() As Parameter
Dim formloadFlag As Boolean
Private Sub comboTables_Click()
    
    LoadDataStation

End Sub
Private Sub Form_Activate()
    VarActForm = Me.Name
    LoadDataStation
    
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call PresetScreenSaver
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call PresetScreenSaver
End Sub
Public Sub LoadDataStation()

    Dim Rst As New ADODB.Recordset
    SetFirstToolBar
    With vsFlex
    .Rows = 1
    Select Case comboTables.ListIndex
    
        Case 0
            
                .Rows = 1
                .Cols = 3
                .TextMatrix(0, 0) = "—œÌ›"
                .TextMatrix(0, 1) = "⁄‰Ê«‰"
                .ColAlignment(-1) = flexAlignCenterCenter
                .ColWidth(0) = 1000
                .ColHidden(1) = False
                .ColHidden(2) = True
             
                
              
               
                Set Rst = RunStoredProcedure2RecordSet("Get_All_tPrefix")
                
                
               
                If Not (Rst.EOF = True And Rst.BOF = True) Then
                    While Rst.EOF <> True
                        .Rows = .Rows + 1
        
                        .TextMatrix(.Rows - 1, 0) = .Rows - 1
                        .TextMatrix(.Rows - 1, 1) = Rst!Description
                        .TextMatrix(.Rows - 1, 2) = Rst!Code
                        Rst.MoveNext
                    Wend
                End If
        Case 1
'                .Rows = 1
'                .Cols = 1
                .Rows = 1
                .Cols = 3
                .TextMatrix(0, 0) = "—œÌ›"
                .TextMatrix(0, 1) = "‰Ê⁄ ›⁄«·Ì "
                              
                .ColAlignment(-1) = flexAlignCenterCenter
                .ColHidden(1) = False
                .ColHidden(2) = True
                
                         

               
                Set Rst = RunStoredProcedure2RecordSet("Get_All_tWorkType")
                If Not (Rst.EOF = True And Rst.BOF = True) Then
                    While Rst.EOF <> True
                        .Rows = .Rows + 1
        
                        .TextMatrix(.Rows - 1, 0) = .Rows - 1
                        .TextMatrix(.Rows - 1, 1) = Rst!Description
                        .TextMatrix(.Rows - 1, 2) = Rst!Code
                        Rst.MoveNext
                    Wend
                End If
        
        
        Case 2
                 .Rows = 1
                .Cols = 3
                .TextMatrix(0, 0) = "—œÌ›"
                .TextMatrix(0, 1) = "ﬂœ «” «‰"
                .TextMatrix(0, 2) = "‰«„ «” «‰"
                .ColHidden(1) = False
                .ColHidden(2) = False
               
               
              
                          
                               
                .ColAlignment(-1) = flexAlignCenterCenter
                                     
               
                Set Rst = RunStoredProcedure2RecordSet("Get_tState")
                If Not (Rst.EOF = True And Rst.BOF = True) Then
                    While Rst.EOF <> True
                        .Rows = .Rows + 1
        
                        .TextMatrix(.Rows - 1, 0) = .Rows - 1
                        .TextMatrix(.Rows - 1, 1) = Rst!Code
                        .TextMatrix(.Rows - 1, 2) = Rst!Description
                       
                                       
                        Rst.MoveNext
                    Wend
                      End If
        
        Case 3
                 .Rows = 1
                .Cols = 4
                .TextMatrix(0, 0) = "—œÌ›"
                .TextMatrix(0, 1) = "ﬂœ ‘Â—"
                .TextMatrix(0, 2) = "‰«„ ‘Â—"
                .TextMatrix(0, 3) = "ﬂœ «” «‰"
                .ColHidden(1) = True
                .ColHidden(2) = False
                .ColHidden(3) = False
                          
                               
                .ColAlignment(-1) = flexAlignCenterCenter
                                     
               
                Set Rst = RunStoredProcedure2RecordSet("Get_tCity")
                If Not (Rst.EOF = True And Rst.BOF = True) Then
                    While Rst.EOF <> True
                        .Rows = .Rows + 1
        
                        .TextMatrix(.Rows - 1, 0) = .Rows - 1
                        .TextMatrix(.Rows - 1, 1) = Rst!Code
                        .TextMatrix(.Rows - 1, 2) = Rst!Description
                        .TextMatrix(.Rows - 1, 3) = IIf(IsNull(Rst!State), "", Rst!State)
                                       
                        Rst.MoveNext
                    Wend
                End If
        
        Case 4
                 .Rows = 1
                .Cols = 4
                .TextMatrix(0, 0) = "—œÌ›"
                .TextMatrix(0, 1) = "Ê«Õœ ﬂ«·«"
                .TextMatrix(0, 2) = "Ê«Õœ »’Ê—  ·« Ì‰"
                .ColHidden(1) = False
                .ColHidden(2) = False
                .ColHidden(3) = True
              
                               
                .ColAlignment(-1) = flexAlignCenterCenter
               
                
                         
               
                Set Rst = RunStoredProcedure2RecordSet("Get_tUnitGood")
                If Not (Rst.EOF = True And Rst.BOF = True) Then
                    While Rst.EOF <> True
                        .Rows = .Rows + 1
        
                        .TextMatrix(.Rows - 1, 0) = .Rows - 1
                        .TextMatrix(.Rows - 1, 1) = Rst!Description
                        .TextMatrix(.Rows - 1, 2) = IIf(IsNull(Rst!LatinDescription), "", Rst!LatinDescription)
                        .TextMatrix(.Rows - 1, 3) = Rst!Code
                        
                       
                        Rst.MoveNext
                    Wend
                End If
        


                 
          Case 5
                 .Rows = 1
                .Cols = 3
                .TextMatrix(0, 0) = "—œÌ›"
                .TextMatrix(0, 1) = "ﬂœ"
                .TextMatrix(0, 2) = "Â“Ì‰Â"
                .ColHidden(1) = True
                .ColHidden(2) = False
                            
                          
                               
                .ColAlignment(-1) = flexAlignCenterCenter
                                     
               
                Set Rst = RunStoredProcedure2RecordSet("Get_tblAcc_ExpensiveType")
                If Not (Rst.EOF = True And Rst.BOF = True) Then
                    While Rst.EOF <> True
                        .Rows = .Rows + 1
        
                        .TextMatrix(.Rows - 1, 0) = .Rows - 1
                        .TextMatrix(.Rows - 1, 1) = Rst!Code
                        .TextMatrix(.Rows - 1, 2) = Rst!Description
                                       
                        Rst.MoveNext
                    Wend
                End If
    
        Case 6
            
                .Rows = 1
                .Cols = 3
                .TextMatrix(0, 0) = "—œÌ›"
                .TextMatrix(0, 1) = "„ﬁ’œ À«‰ÊÌÂ ÕÊ«·Â"
                .ColAlignment(-1) = flexAlignCenterCenter
                .ColWidth(0) = 1000
                .ColHidden(1) = False
                .ColHidden(2) = True
             
                
              
               
                Set Rst = RunStoredProcedure2RecordSet("Get_All_tblPub_Destination")
                
                
               
                If Not (Rst.EOF = True And Rst.BOF = True) Then
                    While Rst.EOF <> True
                        .Rows = .Rows + 1
        
                        .TextMatrix(.Rows - 1, 0) = .Rows - 1
                        .TextMatrix(.Rows - 1, 1) = Trim(Rst!nvcDestination)
                        .TextMatrix(.Rows - 1, 2) = Rst!DestinationId
                        Rst.MoveNext
                    Wend
                End If
     
    End Select
    End With
    
    If Rst.State = 1 Then Rst.Close
 

End Sub
Public Sub Delete()

    On Error GoTo ErrHandler
    If vsFlex.Rows < 2 Then Exit Sub

    With vsFlex
    Select Case comboTables.ListIndex
    
        Case 0
            ReDim Parameter(0) As Parameter
            Parameter(0) = GenerateInputParameter("@intId", adInteger, 4, Val(vsFlex.TextMatrix(vsFlex.Row, 2)))
            RunParametricStoredProcedure "Delete_tPrefix", Parameter
    
        Case 1
            ReDim Parameter(0) As Parameter
            Parameter(0) = GenerateInputParameter("@intId", adInteger, 4, Val(vsFlex.TextMatrix(vsFlex.Row, 2)))
            RunParametricStoredProcedure "Delete_tWorkType", Parameter
        
        Case 2
            ReDim Parameter(0) As Parameter
            Parameter(0) = GenerateInputParameter("@intId", adInteger, 4, Val(vsFlex.TextMatrix(vsFlex.Row, 1)))
            RunParametricStoredProcedure "Delete_tState", Parameter
        
        Case 3
            ReDim Parameter(0) As Parameter
            Parameter(0) = GenerateInputParameter("@intId", adInteger, 4, Val(vsFlex.TextMatrix(vsFlex.Row, 1)))
            RunParametricStoredProcedure "Delete_tCity", Parameter
            
        Case 4
            ReDim Parameter(0) As Parameter
            Parameter(0) = GenerateInputParameter("@intId", adInteger, 4, Val(vsFlex.TextMatrix(vsFlex.Row, 3)))
            RunParametricStoredProcedure "Delete_tUnitGood", Parameter
            
        Case 5
            ReDim Parameter(0) As Parameter
            Parameter(0) = GenerateInputParameter("@intId", adInteger, 4, Val(vsFlex.TextMatrix(vsFlex.Row, 1)))
            RunParametricStoredProcedure "Delete_tblAcc_ExpensiveType", Parameter
        Case 6
            ReDim Parameter(0) As Parameter
            Parameter(0) = GenerateInputParameter("@DestinationId", adInteger, 4, Val(vsFlex.TextMatrix(vsFlex.Row, 2)))
            RunParametricStoredProcedure "Delete_tblPub_Destination", Parameter
    End Select
    End With
    
    
    frmMsg.fwlblMsg.Caption = "»« „Ê›ﬁÌ  Õ–› ‘œ"
    frmMsg.fwBtn(0).Visible = False
    frmMsg.fwBtn(1).ButtonType = flwButtonOk
    frmMsg.fwBtn(1).Caption = "ﬁ»Ê·"
    frmMsg.Show vbModal
    
    LoadDataStation
Exit Sub
    
ErrHandler:
If err.Number = -2147217873 Then

    frmMsg.fwlblMsg.Caption = "„ «”›«‰Â ‘„« ﬁ«œ— »Â Õ–› ‰„Ì »«‘Ìœ"
    frmMsg.fwBtn(0).Visible = False
    frmMsg.fwBtn(1).ButtonType = flwButtonOk
    frmMsg.fwBtn(1).Caption = "ﬁ»Ê·"
    frmMsg.Show vbModal
Else
    MsgBox err.Description
End If
    
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
   
  If ClsFormAccess.frmManageSet = False Then
        Unload Me
        Exit Sub
    End If
    
   CenterCenterinSecondScreen Me
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

    
    
    MyFormAddEditMode = ViewMode
    
    VarActForm = Me.Name
    
    comboTables.Clear
    comboTables.AddItem "⁄‰Ê«‰"
    comboTables.ItemData(comboTables.NewIndex) = 0
    comboTables.AddItem "‰Ê⁄ ›⁄«·Ì "
    comboTables.ItemData(comboTables.NewIndex) = 1
    comboTables.AddItem "«” «‰"
    comboTables.ItemData(comboTables.NewIndex) = 2
    comboTables.AddItem "‘Â—"
    comboTables.ItemData(comboTables.NewIndex) = 3
    comboTables.AddItem "Ê«Õœ ﬂ«·«"
    comboTables.ItemData(comboTables.NewIndex) = 4
    comboTables.AddItem "Â“Ì‰Â Â«"
    comboTables.ItemData(comboTables.NewIndex) = 5
    comboTables.AddItem " „ﬁ’œ »⁄œÌ ÕÊ«·Â"
    comboTables.ItemData(comboTables.NewIndex) = 6
     
    If comboTables.ListCount > 0 Then comboTables.ListIndex = 0

    
    CenterCenter Me
    
    
    With vsFlex
    Select Case comboTables.ListIndex
    
        Case 0
    
        Case 1
        Case 2
    End Select
    End With
  formloadFlag = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
 VarActForm = ""
    If PosConnection.State = adStateOpen Then PosConnection.Close: Set PosConnection = Nothing
    Dim i As Integer
    
    AllButton vbOff, True
    
    Unload frmFindGoods

    SaveSetting strMainKey, Me.Name, "Left", Me.Left
    SaveSetting strMainKey, Me.Name, "Top", Me.Top
   
End Sub
Public Sub ExitForm()

    Unload Me
    
End Sub
Public Sub Edit()
   Select Case comboTables.ListIndex
    Case 0, 1, 4, 5, 6
    MyFormAddEditMode = EditMode
     mdifrm.Toolbar1.Buttons(7).Enabled = False
  
    SetFirstToolBar
    vsFlex.Editable = flexEDKbdMouse
    Case 2, 3
     MyFormAddEditMode = EditMode
      mdifrm.Toolbar1.Buttons(7).Enabled = False
   
     SetFirstToolBar
     vsFlex.Editable = flexEDKbdMouse
     vsFlex.ColHidden(1) = True
     
      End Select
End Sub

Public Sub Add()
    
    Select Case comboTables.ListIndex
    Case 0, 1, 4, 5, 6
        MyFormAddEditMode = AddMode
        mdifrm.Toolbar1.Buttons(6).Enabled = False
        mdifrm.Toolbar1.Buttons(7).Enabled = False
        SetFirstToolBar
        LoadDataStation
        With vsFlex
            .Rows = .Rows + 1
            .Row = .Rows - 1
            .TextMatrix(.Row, 0) = "*"
            .ShowCell .Row, 1
            .Select .Row, 1
        End With
        vsFlex.Editable = flexEDKbdMouse
    Case 2, 3
        MyFormAddEditMode = AddMode
        mdifrm.Toolbar1.Buttons(6).Enabled = False
        mdifrm.Toolbar1.Buttons(7).Enabled = False
        SetFirstToolBar
        LoadDataStation
        With vsFlex
           .ColHidden(1) = True
           .Rows = .Rows + 1
           .Row = .Rows - 1
           .TextMatrix(.Row, 0) = "*"
           .ShowCell .Row, 1
           .Select .Row, 1
        End With
        vsFlex.Editable = flexEDKbdMouse
    End Select
End Sub

Public Sub Cancel()

    MyFormAddEditMode = ViewMode
    mdifrm.Toolbar1.Buttons(6).Enabled = True
    mdifrm.Toolbar1.Buttons(7).Enabled = True
    SetFirstToolBar
    LoadDataStation
    vsFlex.Editable = flexEDNone
      
End Sub
Public Sub ChangeLanguage()

    Select Case clsStation.Language
    
        Case Farsi
        
        Case English
        
    End Select
    
End Sub
Public Sub Update()
    
    ReDim Parameter(3) As Parameter
    
    With vsFlex
    vsFlex_ValidateEdit vsFlex.Row, vsFlex.Col, False
    Select Case MyFormAddEditMode
        Case AddMode
            
'''            If cboBranchType.ItemData(cboBranchType.ListIndex) = 1 And CentralBranch = True Then
'''                frmMsg.fwlblMsg.Caption = "‘„« ‰„Ì  Ê«‰Ìœ »Ì‘ «“ Ìò ‘⁄»Â „—ò“Ì œ«‘ Â »«‘Ìœ"
'''                frmMsg.fwBtn(0).ButtonType = flwButtonOk
'''                frmMsg.fwBtn(0).Caption = "ﬁ»Ê·"
'''                frmMsg.Show vbModal
'''                Exit Sub
'''            End If
             Result = -1
             Select Case comboTables.ListIndex
        
                Case 0
                             
                             For i = 1 To .Rows - 1
                        .Row = i
                        If InStr(.TextMatrix(i, 0), "*") > 0 Then 'new or edited records

                            ReDim Parameter(1) As Parameter
                            Parameter(0) = GenerateInputParameter("@Description", adVarChar, 50, .TextMatrix(i, 1))
                            Parameter(1) = GenerateOutputParameter("@Check", adInteger, 4)
                            Result = RunParametricStoredProcedure("Insert_tPrefix", Parameter)
                          End If
                        Next
                        If Result <> -1 Then
                            frmMsg.fwlblMsg.Caption = "À»  «ÿ·«⁄«  ÃœÌœ »« „Ê›ﬁÌ  Å«Ì«‰ Ì«› "
                            frmMsg.fwBtn(0).ButtonType = flwButtonOk
                            frmMsg.fwBtn(0).Caption = "ﬁ»Ê·"
                            frmMsg.fwBtn(1).Visible = False
                            frmMsg.Show vbModal
                            
                            Add
                        Else
                            frmMsg.fwlblMsg.Caption = "«ÿ·«⁄«  ÃœÌœ À»  ‰‘œ. ·ÿ›« œÊ»«—Â ”⁄Ì ‰„«ÌÌœ"
                            frmMsg.fwBtn(0).ButtonType = flwButtonOk
                            frmMsg.fwBtn(0).Caption = "ﬁ»Ê·"
                            frmMsg.fwBtn(1).Visible = False
                            frmMsg.Show vbModal
                        End If
                    Case 1
                           
                             For i = 1 To .Rows - 1
                        .Row = i
                        If InStr(.TextMatrix(i, 0), "*") > 0 Then 'new or edited records

                           ReDim Parameter(1) As Parameter
                            Parameter(0) = GenerateInputParameter("@Description", adVarChar, 50, .TextMatrix(i, 1))
                            Parameter(1) = GenerateOutputParameter("@Check", adInteger, 4)
                            Result = RunParametricStoredProcedure("Insert_tWorkType", Parameter)
                            
                          End If
                        Next
                        If Result <> -1 Then
                            frmMsg.fwlblMsg.Caption = "À»  «ÿ·«⁄«  ÃœÌœ »« „Ê›ﬁÌ  Å«Ì«‰ Ì«› "
                            frmMsg.fwBtn(0).ButtonType = flwButtonOk
                            frmMsg.fwBtn(0).Caption = "ﬁ»Ê·"
                            frmMsg.fwBtn(1).Visible = False
                            frmMsg.Show vbModal
                            
                           
                        Else
                            frmMsg.fwlblMsg.Caption = "«ÿ·«⁄«  ÃœÌœ À»  ‰‘œ. ·ÿ›« œÊ»«—Â ”⁄Ì ‰„«ÌÌœ"
                            frmMsg.fwBtn(0).ButtonType = flwButtonOk
                            frmMsg.fwBtn(0).Caption = "ﬁ»Ê·"
                            frmMsg.fwBtn(1).Visible = False
                            frmMsg.Show vbModal
                        End If
                    
              Case 2
                           
                             For i = 1 To .Rows - 1
                        .Row = i
                        If InStr(.TextMatrix(i, 0), "*") > 0 Then 'new or edited records

                           ReDim Parameter(1) As Parameter
                             Parameter(0) = GenerateInputParameter("@Description", adWChar, 50, .TextMatrix(i, 2))
                             Parameter(1) = GenerateOutputParameter("@Check", adInteger, 4)
                            Result = RunParametricStoredProcedure("insert_tState", Parameter)
                            
                          End If
                        Next
                        If Result <> -1 Then
                            frmMsg.fwlblMsg.Caption = "À»  «ÿ·«⁄«  ÃœÌœ »« „Ê›ﬁÌ  Å«Ì«‰ Ì«› "
                            frmMsg.fwBtn(0).ButtonType = flwButtonOk
                            frmMsg.fwBtn(0).Caption = "ﬁ»Ê·"
                            frmMsg.fwBtn(1).Visible = False
                            frmMsg.Show vbModal
                            LoadDataStation
                           
                        Else
                            frmMsg.fwlblMsg.Caption = "«ÿ·«⁄«  ÃœÌœ À»  ‰‘œ. ·ÿ›« œÊ»«—Â ”⁄Ì ‰„«ÌÌœ"
                            frmMsg.fwBtn(0).ButtonType = flwButtonOk
                            frmMsg.fwBtn(0).Caption = "ﬁ»Ê·"
                            frmMsg.fwBtn(1).Visible = False
                            frmMsg.Show vbModal
                        End If
'''''                    Case 2
'''''
'''''                             For i = 1 To .Rows - 1
'''''                        .Row = i
'''''                        If InStr(.TextMatrix(i, 0), "*") > 0 Then 'new or edited records
'''''
'''''                           ReDim Parameter(2) As Parameter
'''''                             Parameter(0) = GenerateInputParameter("@Code", adInteger, 4, .TextMatrix(i, 1))
'''''                             Parameter(1) = GenerateInputParameter("@Description", adWChar, 50, .TextMatrix(i, 2))
'''''                             Parameter(2) = GenerateOutputParameter("@Check", adInteger, 4)
'''''                             Result = RunParametricStoredProcedure("Insert_tState", Parameter)
'''''
'''''                          End If
'''''                        Next
'''''                        If Result <> -1 Then
'''''                            frmMsg.fwlblMsg.Caption = "À»  «ÿ·«⁄«  ÃœÌœ »« „Ê›ﬁÌ  Å«Ì«‰ Ì«› "
'''''                            frmMsg.fwBtn(0).ButtonType = flwButtonOk
'''''                            frmMsg.fwBtn(0).Caption = "ﬁ»Ê·"
'''''                            frmMsg.fwBtn(1).Visible = False
'''''                            frmMsg.Show vbModal
'''''
'''''
'''''                        Else
'''''                            frmMsg.fwlblMsg.Caption = "«ÿ·«⁄«  ÃœÌœ À»  ‰‘œ. ·ÿ›« œÊ»«—Â ”⁄Ì ‰„«ÌÌœ"
'''''                            frmMsg.fwBtn(0).ButtonType = flwButtonOk
'''''                            frmMsg.fwBtn(0).Caption = "ﬁ»Ê·"
'''''                            frmMsg.fwBtn(1).Visible = False
'''''                            frmMsg.Show vbModal
'''''                        End If
                    Case 3
                           
                             For i = 1 To .Rows - 1
                        .Row = i
                        If InStr(.TextMatrix(i, 0), "*") > 0 Then 'new or edited records

                           ReDim Parameter(2) As Parameter
                             Parameter(0) = GenerateInputParameter("@Description", adWChar, 50, .TextMatrix(i, 2))
                             Parameter(1) = GenerateInputParameter("@State", adInteger, 4, IIf(.TextMatrix(i, 3) = "", Null, Val(.TextMatrix(i, 3))))
                             Parameter(2) = GenerateOutputParameter("@Check", adInteger, 4)
                             Result = RunParametricStoredProcedure("Insert_tCity", Parameter)
                            
                          End If
                        Next
                        If Result <> -1 Then
                            frmMsg.fwlblMsg.Caption = "À»  «ÿ·«⁄«  ÃœÌœ »« „Ê›ﬁÌ  Å«Ì«‰ Ì«› "
                            frmMsg.fwBtn(0).ButtonType = flwButtonOk
                            frmMsg.fwBtn(0).Caption = "ﬁ»Ê·"
                            frmMsg.fwBtn(1).Visible = False
                            frmMsg.Show vbModal
                            
                           
                        Else
                            frmMsg.fwlblMsg.Caption = "«ÿ·«⁄«  ÃœÌœ À»  ‰‘œ. ·ÿ›« œÊ»«—Â ”⁄Ì ‰„«ÌÌœ"
                            frmMsg.fwBtn(0).ButtonType = flwButtonOk
                            frmMsg.fwBtn(0).Caption = "ﬁ»Ê·"
                            frmMsg.fwBtn(1).Visible = False
                            frmMsg.Show vbModal
                        End If
                      Case 4
                           
                             For i = 1 To .Rows - 1
                        .Row = i
                        If InStr(.TextMatrix(i, 0), "*") > 0 Then 'new or edited records

                           ReDim Parameter(2) As Parameter
                             Parameter(0) = GenerateInputParameter("@Description", adWChar, 50, .TextMatrix(i, 1))
                             Parameter(1) = GenerateInputParameter("@LatinDescription", adWChar, 50, .TextMatrix(i, 2))
                             Parameter(2) = GenerateOutputParameter("@Check", adInteger, 4)
                             Result = RunParametricStoredProcedure("Insert_tUnitGood", Parameter)
                            
                          End If
                        Next
                        If Result <> -1 Then
                            frmMsg.fwlblMsg.Caption = "À»  «ÿ·«⁄«  ÃœÌœ »« „Ê›ﬁÌ  Å«Ì«‰ Ì«› "
                            frmMsg.fwBtn(0).ButtonType = flwButtonOk
                            frmMsg.fwBtn(0).Caption = "ﬁ»Ê·"
                            frmMsg.fwBtn(1).Visible = False
                            frmMsg.Show vbModal
                            
                           
                        Else
                            frmMsg.fwlblMsg.Caption = "«ÿ·«⁄«  ÃœÌœ À»  ‰‘œ. ·ÿ›« œÊ»«—Â ”⁄Ì ‰„«ÌÌœ"
                            frmMsg.fwBtn(0).ButtonType = flwButtonOk
                            frmMsg.fwBtn(0).Caption = "ﬁ»Ê·"
                            frmMsg.fwBtn(1).Visible = False
                            frmMsg.Show vbModal
                        End If
                     Case 5
                           
                             For i = 1 To .Rows - 1
                        .Row = i
                        If InStr(.TextMatrix(i, 0), "*") > 0 Then 'new or edited records

                           ReDim Parameter(1) As Parameter
                             Parameter(0) = GenerateInputParameter("@Description", adWChar, 50, .TextMatrix(i, 2))
                             Parameter(1) = GenerateOutputParameter("@Check", adInteger, 4)
                             Result = RunParametricStoredProcedure("Insert_tblAcc_ExpensiveType", Parameter)
                            
                          End If
                        Next
                        If Result <> -1 Then
                            frmMsg.fwlblMsg.Caption = "À»  «ÿ·«⁄«  ÃœÌœ »« „Ê›ﬁÌ  Å«Ì«‰ Ì«› "
                            frmMsg.fwBtn(0).ButtonType = flwButtonOk
                            frmMsg.fwBtn(0).Caption = "ﬁ»Ê·"
                            frmMsg.fwBtn(1).Visible = False
                            frmMsg.Show vbModal
                            
                           
                        Else
                            frmMsg.fwlblMsg.Caption = "«ÿ·«⁄«  ÃœÌœ À»  ‰‘œ. ·ÿ›« œÊ»«—Â ”⁄Ì ‰„«ÌÌœ"
                            frmMsg.fwBtn(0).ButtonType = flwButtonOk
                            frmMsg.fwBtn(0).Caption = "ﬁ»Ê·"
                            frmMsg.fwBtn(1).Visible = False
                            frmMsg.Show vbModal
                        End If
                
                Case 6
                             
                    For i = 1 To .Rows - 1
                        .Row = i
                        If InStr(.TextMatrix(i, 0), "*") > 0 Then 'new or edited records

                            ReDim Parameter(1) As Parameter
                            Parameter(0) = GenerateInputParameter("@nvcDestination", adVarChar, 50, Trim(.TextMatrix(i, 1)))
                            Parameter(1) = GenerateOutputParameter("@intStatue", adInteger, 4)
                            Result = RunParametricStoredProcedure("Insert_tblPub_Destination", Parameter)
                          End If
                        Next
                        If Result <> -1 Then
                            frmMsg.fwlblMsg.Caption = "À»  «ÿ·«⁄«  ÃœÌœ »« „Ê›ﬁÌ  Å«Ì«‰ Ì«› "
                            frmMsg.fwBtn(0).ButtonType = flwButtonOk
                            frmMsg.fwBtn(0).Caption = "ﬁ»Ê·"
                            frmMsg.fwBtn(1).Visible = False
                            frmMsg.Show vbModal
                            
                            Add
                        Else
                            frmMsg.fwlblMsg.Caption = "«ÿ·«⁄«  ÃœÌœ À»  ‰‘œ. ·ÿ›« œÊ»«—Â ”⁄Ì ‰„«ÌÌœ"
                            frmMsg.fwBtn(0).ButtonType = flwButtonOk
                            frmMsg.fwBtn(0).Caption = "ﬁ»Ê·"
                            frmMsg.fwBtn(1).Visible = False
                            frmMsg.Show vbModal
                        End If
                 
               
                 End Select
        Case EditMode
            Select Case comboTables.ListIndex
        
                Case 0
                    For i = 1 To .Rows - 1
                        .Row = i
                        If InStr(.TextMatrix(i, 0), "*") > 0 Then 'new or edited records

                            ReDim Parameter(2) As Parameter
                            Parameter(0) = GenerateInputParameter("@Code", adInteger, 4, .TextMatrix(i, 2))
                            Parameter(1) = GenerateInputParameter("@Description", adWChar, 50, .TextMatrix(i, 1))
                            Parameter(2) = GenerateOutputParameter("@Check", adInteger, 4)
                            Result = RunParametricStoredProcedure("Update_tPrefix", Parameter)
                            
                        End If
                    Next
                Case 1
                 For i = 1 To .Rows - 1
                        .Row = i
                        If InStr(.TextMatrix(i, 0), "*") > 0 Then 'new or edited records

                            ReDim Parameter(2) As Parameter
                            Parameter(0) = GenerateInputParameter("@Code", adInteger, 4, .TextMatrix(i, 2))
                            Parameter(1) = GenerateInputParameter("@Description", adWChar, 50, .TextMatrix(i, 1))
                            Parameter(2) = GenerateOutputParameter("@Check", adInteger, 4)
                            Result = RunParametricStoredProcedure("Update_tWorkType", Parameter)
                            
                        End If
                    Next
                Case 2
                 For i = 1 To .Rows - 1
                        .Row = i
                        If InStr(.TextMatrix(i, 0), "*") > 0 Then 'new or edited records

                            ReDim Parameter(2) As Parameter
                            Parameter(0) = GenerateInputParameter("@Description", adWChar, 50, .TextMatrix(i, 2))
                            Parameter(1) = GenerateInputParameter("@Code", adInteger, 4, .TextMatrix(i, 1))
                            Parameter(2) = GenerateOutputParameter("@Check", adInteger, 4)
                            Result = RunParametricStoredProcedure("Update_tState", Parameter)
                            
                        End If
                    Next
                Case 3
                 For i = 1 To .Rows - 1
                        .Row = i
                        If InStr(.TextMatrix(i, 0), "*") > 0 Then 'new or edited records
                            If .TextMatrix(i, 3) = "" Then
            
                                frmMsg.fwlblMsg.Caption = "ﬂœ «” «‰ —« Ê«—œ ﬂ‰Ìœ"
                                frmMsg.fwBtn(0).ButtonType = flwButtonOk
                                frmMsg.fwBtn(0).Caption = "ﬁ»Ê·"
                                frmMsg.fwBtn(1).Visible = False
                                frmMsg.Show vbModal
                                Exit Sub
                            End If
                            ReDim Parameter(3) As Parameter
                            Parameter(0) = GenerateInputParameter("@Code", adInteger, 4, .TextMatrix(i, 1))
                            Parameter(1) = GenerateInputParameter("@Description", adWChar, 50, .TextMatrix(i, 2))
                            Parameter(2) = GenerateInputParameter("@State", adInteger, 4, .TextMatrix(i, 3))
                            Parameter(3) = GenerateOutputParameter("@Check", adInteger, 4)
                            Result = RunParametricStoredProcedure("Update_tCity", Parameter)
                            
                        End If
                    Next
               Case 4
                 For i = 1 To .Rows - 1
                        .Row = i
                        If InStr(.TextMatrix(i, 0), "*") > 0 Then 'new or edited records

                            ReDim Parameter(3) As Parameter
                            Parameter(0) = GenerateInputParameter("Code", adInteger, 4, .TextMatrix(i, 3))
                            Parameter(1) = GenerateInputParameter("@Description", adWChar, 50, .TextMatrix(i, 1))
                            Parameter(2) = GenerateInputParameter("@LatinDescription", adWChar, 50, .TextMatrix(i, 2))
                            Parameter(3) = GenerateOutputParameter("@Check", adInteger, 4)
                            Result = RunParametricStoredProcedure("Update_tUnitGood", Parameter)
                            
                        End If
                    Next
            Case 5
                 For i = 1 To .Rows - 1
                        .Row = i
                        If InStr(.TextMatrix(i, 0), "*") > 0 Then 'new or edited records

                            ReDim Parameter(2) As Parameter
                            Parameter(0) = GenerateInputParameter("Code", adInteger, 4, .TextMatrix(i, 1))
                            Parameter(1) = GenerateInputParameter("@Description", adWChar, 50, .TextMatrix(i, 2))
                            Parameter(2) = GenerateOutputParameter("@Check", adInteger, 4)
                            Result = RunParametricStoredProcedure("Update_tblAcc_ExpensiveType", Parameter)
                            
                        End If
                    Next
            Case 6
                For i = 1 To .Rows - 1
                    .Row = i
                    If InStr(.TextMatrix(i, 0), "*") > 0 Then 'new or edited records

                        ReDim Parameter(2) As Parameter
                        Parameter(0) = GenerateInputParameter("@DestinationId", adInteger, 4, .TextMatrix(i, 2))
                        Parameter(1) = GenerateInputParameter("@nvcDestination", adWChar, 50, Trim(.TextMatrix(i, 1)))
                        Parameter(2) = GenerateOutputParameter("@intStatus", adInteger, 4)
                        Result = RunParametricStoredProcedure("Update_tblPub_Destination", Parameter)
                        
                    End If
                Next
        
        End Select
            If Result <> -1 Then
            
                frmMsg.fwlblMsg.Caption = " €ÌÌ— «ÿ·«⁄«  »« „Ê›ﬁÌ  Å«Ì«‰ Ì«› "
                frmMsg.fwBtn(0).ButtonType = flwButtonOk
                frmMsg.fwBtn(0).Caption = "ﬁ»Ê·"
                frmMsg.fwBtn(1).Visible = False
                frmMsg.Show vbModal
                LoadDataStation
            Else

                frmMsg.fwlblMsg.Caption = "„ «”›«‰Â «ÿ·«⁄«   €ÌÌ— ‰Ì«› . ·ÿ›« œÊ»«—Â ”⁄Ì ‰„«ÌÌœ"
                frmMsg.fwBtn(0).ButtonType = flwButtonOk
                frmMsg.fwBtn(0).Caption = "ﬁ»Ê·"
                frmMsg.fwBtn(1).Visible = False
                frmMsg.Show vbModal
                
            End If
     
      End Select

    End With
    vsFlex.Editable = flexEDNone
''''    mdifrm.Toolbar1.Buttons(6).Enabled = True
''''    mdifrm.Toolbar1.Buttons(7).Enabled = True
    MyFormAddEditMode = ViewMode
End Sub
''    HeaderLabel Val(MyFormAddEditMode), fwlblMode
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
        
            For i = 6 To 10
                mdifrm.Toolbar1.Buttons(i).Enabled = True
            Next i
''            vsGood.Editable = flexEDNone
          
    mdifrm.Toolbar1.Buttons(7).Enabled = True
        Case EnumAddEditMode.AddMode
        
            mdifrm.Toolbar1.Buttons(8).Enabled = True 'enter key
            mdifrm.Toolbar1.Buttons(9).Enabled = True 'cancel key

            
        Case EnumAddEditMode.EditMode
                    
            mdifrm.Toolbar1.Buttons(8).Enabled = True 'enter key
            mdifrm.Toolbar1.Buttons(9).Enabled = True 'cancel key

    End Select
    
    HeaderLabel Val(MyFormAddEditMode), fwlblMode
    
End Sub
Private Sub vsFlex_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    With vsFlex
        If MyFormAddEditMode = EnumAddEditMode.EditMode And InStr(.TextMatrix(Row, 0), "*") = 0 Then
            .TextMatrix(Row, 0) = Trim(.TextMatrix(Row, 0)) & "*"
        End If
    End With
End Sub

Private Sub vsFlex_Click()
    
If MyFormAddEditMode = EditMode Then
    With vsFlex
        If .Row = 0 Then Exit Sub
        
       .EditCell
       
    End With
End If
End Sub
Private Sub ResizeKit1_ExitResize(ByVal XScale As Double, ByVal YScale As Double)

    If formloadFlag = True Then
        SaveSetting strMainKey, Me.Name, "Height", Me.Height
        SaveSetting strMainKey, Me.Name, "Width", Me.Width
    End If

End Sub

Private Sub vsFlex_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsFlex
        .Row = Row
        .Col = Col
    End With

End Sub
