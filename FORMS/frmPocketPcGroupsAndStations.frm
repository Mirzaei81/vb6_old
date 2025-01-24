VERSION 5.00
Object = "{75D4F767-8785-11D3-93AD-0000832EF44D}#2.14#0"; "FAST2001.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmPocketPcGroupsAndStations 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6810
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7695
   Icon            =   "frmPocketPcGroupsAndStations.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   6810
   ScaleWidth      =   7695
   Begin VB.CommandButton Command1 
      BackColor       =   &H00008000&
      Caption         =   "À»   ’ÊÌ— ê—ÊÂ ò«·«"
      BeginProperty Font 
         Name            =   "B Homa"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   630
      Left            =   405
      RightToLeft     =   -1  'True
      TabIndex        =   7
      Top             =   6000
      Width           =   1785
   End
   Begin VB.CommandButton CmdDone 
      BackColor       =   &H00008000&
      Caption         =   " «ÌÌœ"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "B Homa"
         Size            =   9.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   4710
      RightToLeft     =   -1  'True
      TabIndex        =   5
      Top             =   6045
      Width           =   1410
   End
   Begin VB.TextBox txtGroupName 
      Alignment       =   1  'Right Justify
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
      Left            =   4335
      RightToLeft     =   -1  'True
      TabIndex        =   2
      Top             =   1920
      Width           =   3195
   End
   Begin VB.ListBox lstGroups 
      BeginProperty Font 
         Name            =   "Nazanin"
         Size            =   12
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3180
      Left            =   4335
      RightToLeft     =   -1  'True
      Style           =   1  'Checkbox
      TabIndex        =   1
      Top             =   2535
      Width           =   3195
   End
   Begin VB.ComboBox cboStation 
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
      Left            =   4305
      RightToLeft     =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   960
      Width           =   3195
   End
   Begin FLWCtrls.FWLabel3D fwlblMode 
      Height          =   525
      Left            =   6330
      Top             =   -15
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   926
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
   Begin FLWCtrls.FWButton FWBtnpicture 
      Height          =   690
      Left            =   210
      TabIndex        =   8
      Top             =   810
      Width           =   2595
      _ExtentX        =   4577
      _ExtentY        =   1217
      ButtonType      =   8
      Caption         =   " «‰ Œ«»  ’ÊÌ— ê—ÊÂ ò«·« "
      BackColor       =   49152
      ForeColor       =   16384
      FontName        =   "B Homa"
      FontBold        =   -1  'True
      FontSize        =   9.75
      Alignment       =   1
   End
   Begin MSComDlg.CommonDialog Cdlg 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Image Image1 
      Height          =   3330
      Left            =   285
      Stretch         =   -1  'True
      Top             =   2310
      Width           =   3720
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   " ⁄—Ì› ê—ÊÂÂ« œ—PocketPC"
      BeginProperty Font 
         Name            =   "Nazanin"
         Size            =   15.75
         Charset         =   178
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   495
      Left            =   900
      RightToLeft     =   -1  'True
      TabIndex        =   6
      Top             =   0
      Width           =   4845
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "‰«„ ê—ÊÂ"
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
      Left            =   6465
      RightToLeft     =   -1  'True
      TabIndex        =   4
      Top             =   1500
      Width           =   1065
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "‰«„ «Ì” ê«Â"
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
      Left            =   6510
      RightToLeft     =   -1  'True
      TabIndex        =   3
      Top             =   555
      Width           =   1065
   End
End
Attribute VB_Name = "frmPocketPcGroupsAndStations"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MyFormAddEditMode As EnumAddEditMode
Dim Parameter() As Parameter
Dim strFileName As String

Private Sub cboStation_Click()

    Dim Rst As New ADODB.Recordset
    ReDim Parameter(1) As Parameter
    Dim TempStationId As Integer
    
    lstGroups.Clear
    
    If cboStation.ListIndex > -1 Then
        TempStationId = cboStation.ItemData(cboStation.ListIndex)
    Else
        TempStationId = 0
    End If
    
    Parameter(0) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
    Parameter(1) = GenerateInputParameter("@StationId", adInteger, 4, TempStationId)
    
    If Rst.State <> 0 Then Rst.Close
    Set Rst = RunParametricStoredProcedure2Rec("Get_tPocketPCGroup", Parameter)
    
    While Rst.EOF <> True
        If Rst.Fields("PocketPCGroupCode").Value > 8 Then
        Select Case clsStation.Language
            Case Farsi
                lstGroups.AddItem Rst.Fields("Description").Value
            Case English
                lstGroups.AddItem Rst.Fields("LatinDescription").Value
        End Select

        lstGroups.ItemData(lstGroups.ListCount - 1) = Rst.Fields("PocketPCGroupCode").Value
        If IsNull(Rst.Fields("StationId").Value) <> True Then
                    lstGroups.Selected(lstGroups.ListCount - 1) = True

        End If
        End If
        Rst.MoveNext
    Wend
    
    Set Rst = Nothing
End Sub

Private Sub CmdDone_Click()
    
    Dim i As Integer
    
    If lstGroups.SelCount = 0 Then Exit Sub
    
    Dim SelectedGroups As String
    ReDim Parameter(1) As Parameter
    For i = 0 To lstGroups.ListCount - 1
        If lstGroups.Selected(i) = True Then
            SelectedGroups = SelectedGroups & lstGroups.ItemData(i) & ","
        End If
    Next i
    SelectedGroups = Left(SelectedGroups, Len(SelectedGroups) - 1)
    If cboStation.ListIndex > -1 Then
        Parameter(0) = GenerateInputParameter("@StationId", adInteger, 4, cboStation.ItemData(cboStation.ListIndex))
        Parameter(1) = GenerateInputParameter("@PocketPCGroupCode", adVarWChar, 4000, SelectedGroups)
        
        RunParametricStoredProcedure "Update_tPocketPC_StationGroups", Parameter
        ShowDisMessage "«Œ ’«’ ê—ÊÂ »Â «Ì” ê«Â «‰Ã«„ ‘œ", 1500
    End If
End Sub

Private Sub Command1_Click()
     ReDim Parameter(2) As Parameter
     Parameter(0) = GenerateInputParameter("@PicturePath", adVarWChar, 300, strFileName)
     Parameter(1) = GenerateInputParameter("@PocketPcGroupCode", adInteger, 4, lstGroups.ItemData(lstGroups.ListIndex))
     Parameter(2) = GenerateOutputParameter("@Updated", adInteger, 4)

     Dim Updated As Long
     Updated = RunParametricStoredProcedure("Update_tPocketPCGroup_Pic", Parameter)
     If Updated = 1 Then
         frmMsg.fwlblMsg.Caption = " €ÌÌ—«  »« „Ê›ﬁÌ  «‰Ã«„ ‘œ"
         frmMsg.fwBtn(0).Visible = False
         frmMsg.fwBtn(1).ButtonType = flwButtonOk
         frmMsg.fwBtn(1).Caption = "ﬁ»Ê·"
         frmMsg.Show vbModal
         lstGroups_Click
    Else
         frmMsg.fwlblMsg.Caption = " €ÌÌ—«  «‰Ã«„ ‰‘œ"
         frmMsg.fwBtn(0).Visible = False
         frmMsg.fwBtn(1).ButtonType = flwButtonOk
         frmMsg.fwBtn(1).Caption = "ﬁ»Ê·"
         frmMsg.Show vbModal
         Exit Sub
     End If

End Sub

Private Sub Form_Activate()
    VarActForm = Me.Name
    SetFirstToolBar
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call PresetScreenSaver
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call PresetScreenSaver
End Sub

Private Sub Form_Load()

    If intVersion <> gold And intVersion <> Diamond Then
        ShowDisMessage "«„ﬂ«‰  ⁄—Ì› ﬂ«·«Â« œ— Å«ﬂ  ÅÌ ”Ì Ê  »·  ›ﬁÿ œ— ‰”ŒÂ ÊÌéÂ ÊÃÊœ œ«—œ", 1500
        Unload Me
        Exit Sub
    End If
    CenterCenter Me
    
    VarActForm = Me.Name
    
    MyFormAddEditMode = ViewMode
    DefaultSetting
    SetFirstToolBar
    
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

Private Sub Form_Unload(Cancel As Integer)
    If PosConnection.State = adStateOpen Then PosConnection.Close: Set PosConnection = Nothing
    VarActForm = ""
End Sub

Public Sub Add()

    MyFormAddEditMode = AddMode
    SetFirstToolBar
    
End Sub
Public Sub Cancel()

    DefaultSetting
    MyFormAddEditMode = ViewMode
    SetFirstToolBar
    
End Sub
Public Sub Delete()

End Sub

Public Sub Edit()
    MyFormAddEditMode = EditMode
    SetFirstToolBar
End Sub

Public Sub Update()
    
    Dim intResult As Integer
    
    Select Case MyFormAddEditMode
    
        Case AddMode
        
            If txtGroupName.Text = "" Or InStr(txtGroupName.Text, "'") <> 0 Then
                Exit Sub
            End If
            
            ReDim Parameter(1) As Parameter
            
            Parameter(0) = GenerateInputParameter("@Description", adVarWChar, 50, txtGroupName.Text)
            Parameter(1) = GenerateOutputParameter("@Result", adInteger, 4)
            intResult = RunParametricStoredProcedure("Insert_PocketPCGroup", Parameter)
            If intResult <> -1 Then
            
            Else
                ShowDisMessage "«÷«›Â ‘œ‰ ê—ÊÂ «‰Ã«„ ‘œ", 1500
            End If
            
        Case EditMode
        
            If txtGroupName.Text = "" Or InStr(txtGroupName.Text, "'") <> 0 Then
                Exit Sub
            End If
            Dim Parameter2(3) As Parameter
            Parameter2(0) = GenerateInputParameter("@PocketPCGroup", adInteger, 4, lstGroups.ItemData(lstGroups.ListIndex))
            Parameter2(1) = GenerateInputParameter("@Description", adVarWChar, 50, txtGroupName.Text)
            Parameter2(2) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
            Parameter2(3) = GenerateOutputParameter("@Result", adInteger, 4)
            intResult = RunParametricStoredProcedure("Update_PocketPCGroup", Parameter2)
            If intResult <> -1 Then
            
            Else
            
            End If
    End Select
    
    DefaultSetting
    MyFormAddEditMode = ViewMode
    SetFirstToolBar
    HeaderLabel CInt(MyFormAddEditMode), Me.fwlblMode
End Sub
Public Sub ChangeLanguage()
    
    DefaultSetting
    
End Sub

Private Sub DefaultSetting()

    cboStation.Clear
    lstGroups.Clear
    txtGroupName.Text = ""
    txtGroupName.Locked = True
    
    Dim Rst As New ADODB.Recordset
    
    If Rst.State <> 0 Then Rst.Close
    Rst.Open "Select * from  dbo.tStations Where (StationType & 8 = 8 or StationType & 16 = 16 ) And Branch =  " & CurrentBranch, PosConnection, adOpenDynamic, adLockOptimistic, adCmdText
    ' ﬂ‰ —·  «ÌÅ «Ì” ê«Â Â«—œ ﬂœ ‘œ
'    Set Rst = RunStoredProcedure2RecordSet("Get_Pocket_Stations")
    
    i = 0
    If Not (Rst.EOF = True And Rst.BOF = True) Then
        While Not Rst.EOF
            i = i + 1
            cboStation.AddItem Rst.Fields("Description").Value
            cboStation.ItemData(cboStation.ListCount - 1) = Rst.Fields("StationID").Value
            Rst.MoveNext
        Wend
    End If
    If Rst.State <> 0 Then Rst.Close
    If i = 0 Then
       MsgBox " Ì«  »·   ÊÃÊœ ‰œ«—œ  Pocket_Pc «Ì” ê«Â"
       Exit Sub
    ElseIf i > clsArya.MaxPocketPcNo + clsArya.MaxTabletNo Then
       MsgBox "Œÿ« œ—  ⁄œ«œ «Ì” ê«ÂÂ«Ì Pocket_Pc"
       End
    End If
    
    
    ReDim Parameter(1) As Parameter
    Dim TempStationId As Integer
    
    If cboStation.ListIndex > -1 Then
        TempStationId = cboStation.ItemData(cboStation.ListIndex)
    Else
        TempStationId = 0
    End If
    Parameter(0) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
    Parameter(1) = GenerateInputParameter("@StationId", adInteger, 4, TempStationId)
    
    If Rst.State <> 0 Then Rst.Close
    Set Rst = RunParametricStoredProcedure2Rec("Get_tPocketPCGroup", Parameter)
    
    If Not (Rst.EOF = True And Rst.BOF = True) Then
        While Rst.EOF <> True
            Select Case clsStation.Language
                Case Farsi
                    lstGroups.AddItem Rst.Fields("Description").Value
                Case English
                    lstGroups.AddItem Rst.Fields("LatinDescription").Value
            End Select
            lstGroups.ItemData(lstGroups.ListCount - 1) = Rst.Fields("PocketPCGroupCode").Value
            If IsNull(Rst.Fields("StationId").Value) <> True Then
                        lstGroups.Selected(lstGroups.ListCount - 1) = True
    
            End If
            Rst.MoveNext
        Wend
        cboStation.ListIndex = 0
    End If
    Set Rst = Nothing
    
    
End Sub

Public Sub ExitForm()
    Unload Me

End Sub

Private Sub SetFirstToolBar()

    AllButton vbOff, True
    mdifrm.Toolbar1.Buttons(1).Enabled = True   'Home
    mdifrm.Toolbar1.Buttons(2).Enabled = True   'PageUp
    mdifrm.Toolbar1.Buttons(3).Enabled = True   'PageDown
    mdifrm.Toolbar1.Buttons(4).Enabled = True   'End
    
    mdifrm.Toolbar1.Buttons(15).Enabled = True   'Print
    
    mdifrm.Toolbar1.Buttons(23).Enabled = True
    mdifrm.Toolbar1.Buttons(24).Enabled = True
    mdifrm.Toolbar1.Buttons(25).Enabled = True
    mdifrm.Toolbar1.Buttons(26).Enabled = True
    mdifrm.Toolbar1.Buttons(27).Enabled = True
    
    If MyFormAddEditMode = ViewMode Then  ' View Mode
    
        mdifrm.Toolbar1.Buttons(6).Enabled = True  'Add
        mdifrm.Toolbar1.Buttons(7).Enabled = True  'Edit
        mdifrm.Toolbar1.Buttons(8).Enabled = False  'Enter
        mdifrm.Toolbar1.Buttons(9).Enabled = False   'Esc
        txtGroupName.Locked = True
        
    ElseIf MyFormAddEditMode = AddMode Then    'Add Mode
    
        mdifrm.Toolbar1.Buttons(6).Enabled = False  'Add
        mdifrm.Toolbar1.Buttons(7).Enabled = False  'Edit
        mdifrm.Toolbar1.Buttons(8).Enabled = True   'Enter
        mdifrm.Toolbar1.Buttons(9).Enabled = True   'Esc
        txtGroupName.Locked = False
        
    ElseIf MyFormAddEditMode = EditMode Then     'Edit
        
        mdifrm.Toolbar1.Buttons(6).Enabled = False  'Add
        mdifrm.Toolbar1.Buttons(7).Enabled = False  'Edit
        mdifrm.Toolbar1.Buttons(8).Enabled = True   'Enter
        mdifrm.Toolbar1.Buttons(9).Enabled = True   'Esc
        txtGroupName.Locked = False
    
    End If
    
    HeaderLabel Val(MyFormAddEditMode), fwlblMode


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

Private Sub lstGroups_Click()
    On Error GoTo err
    Dim intCode As Long
    If MyFormAddEditMode = EditMode Then
    
       txtGroupName.Text = lstGroups.List(lstGroups.ListIndex)
       
    End If
    FWBtnpicture.Tag = 0
    Dim TempStr As String
    Dim Token As Long
    Dim c
    intCode = Val(lstGroups.ItemData(lstGroups.ListIndex))
    
    ReDim Parameter(0) As Parameter
    Parameter(0) = GenerateInputParameter("@PocketPcGroupCode", adInteger, 4, intCode)
    Set rctmp = RunParametricStoredProcedure2Rec("Get_tPocketPCGroup_ByID", Parameter)
    Dim ii As Integer
    If Not (rctmp.BOF Or rctmp.EOF) Then
           Image1.Picture = LoadPicture("")
           FWBtnpicture.Tag = rctmp!PocketPcGroupCode
           strFileName = IIf(IsNull(rctmp!PicturePath), "", rctmp!PicturePath)
          
          '' On Error Resume Next
    
'            If IsNull(rctmp.Fields("Picture").Value) Then
'               Image1.Picture = LoadPicture(rctmp!PicturePath)
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
             If strFileName <> "" Then
                 c = Me.BackColor
                    
                 If c < 0 Then c = GetSysColor(c - &H80000000)
                    
                 Token = InitGDIPlus
                    
                ' Picture1(0).Picture = LoadPictureGDIPlus(App.Path & "\1.png", , , vbWhite)
                ' Picture1(1).Picture = LoadPictureGDIPlus(App.Path & "\1.png", , , vbCyan)
                ' Picture1(2).Picture = LoadPictureGDIPlus(App.Path & "\1.png", , , vbGreen)
                 Image1.Picture = LoadPictureGDIPlus(strFileName, , , c)
                    
                 FreeGDIPlus Token
            End If
    Else
           Image1.Picture = LoadPicture("")
          
    End If
Exit Sub
err:
 If err.Number = 53 Then
 
        Image1.Picture = LoadPicture("")
        frmMsg.fwlblMsg.Caption = "⁄ﬂ” „Ê—œ ‰Ÿ— Å«ﬂ ‘œÂ «” "
        frmMsg.fwBtn(0).Caption = "ﬁ»Ê·"

        frmMsg.fwBtn(1).Visible = False
        frmMsg.fwBtn(0).ButtonType = flwButtonOk
        frmMsg.Show vbModal
 Else
    ShowDisMessage err.Description, 1000
 End If
    rctmp.Close

End Sub

Private Sub lstGroups_ItemCheck(Item As Integer)
'''
    If cboStation.ListIndex > -1 Then
        CmdDone.Enabled = True
    End If
End Sub
