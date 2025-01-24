VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{7AEDC602-D94C-11D1-BB7A-00E0290EA3C9}#1.0#0"; "ResizeKit.ocx"
Object = "{75D4F767-8785-11D3-93AD-0000832EF44D}#2.14#0"; "FAST2001.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmStation_Inventory 
   Caption         =   "                                                                                                   ÇÎÊÕÇÕ ÇäÈÇÑ Èå ÇíÓÊÇå     "
   ClientHeight    =   9735
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11280
   Icon            =   "frmStation_Inventory.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   9735
   ScaleWidth      =   11280
   Begin Crystal.CrystalReport CrystalReport1 
      Left            =   3240
      Top             =   0
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
   Begin VB.Frame Frame2 
      Height          =   3135
      Left            =   2250
      TabIndex        =   1
      Top             =   360
      Width           =   9015
      Begin VB.Frame Frame1 
         Caption         =   "ÓÇá ãÇáí"
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   11.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   960
         Left            =   6120
         RightToLeft     =   -1  'True
         TabIndex        =   11
         Top             =   2040
         Width           =   2775
         Begin VB.ComboBox cmbSalMali 
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
            Left            =   890
            RightToLeft     =   -1  'True
            TabIndex        =   12
            Top             =   360
            Width           =   1815
         End
      End
      Begin VB.CommandButton cmdStationDelete 
         BackColor       =   &H0080C0FF&
         Caption         =   "ÍÐÝ ÇíÓÊÇå ÇÒ ÇäÈÇÑ"
         BeginProperty Font 
            Name            =   "Nazanin"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   480
         TabIndex        =   9
         Top             =   1680
         Width           =   1695
      End
      Begin VB.ListBox lsttStations 
         BeginProperty Font 
            Name            =   "Nazanin"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2400
         Left            =   2400
         RightToLeft     =   -1  'True
         Style           =   1  'Checkbox
         TabIndex        =   7
         Top             =   480
         Width           =   3195
      End
      Begin VB.Frame Frame28 
         Caption         =   "ÇäÈÇÑåÇ"
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   11.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   960
         Left            =   6120
         RightToLeft     =   -1  'True
         TabIndex        =   5
         Top             =   1080
         Width           =   2775
         Begin VB.ComboBox cmbInventory 
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
            Left            =   240
            RightToLeft     =   -1  'True
            TabIndex        =   6
            Top             =   360
            Width           =   2475
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "ÔÚÈå"
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   11.25
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   960
         Left            =   6120
         RightToLeft     =   -1  'True
         TabIndex        =   3
         Top             =   120
         Width           =   2775
         Begin VB.ComboBox cmbBranch 
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
            Left            =   240
            RightToLeft     =   -1  'True
            TabIndex        =   4
            Top             =   360
            Width           =   2475
         End
      End
      Begin VB.CommandButton CmdStationAdd 
         BackColor       =   &H00FF8080&
         Caption         =   "ÇÖÇÝå ˜ÑÏä ÇíÓÊÇå Èå ÇäÈÇÑ"
         BeginProperty Font 
            Name            =   "Nazanin"
            Size            =   12
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   480
         TabIndex        =   2
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "äÇã ÇíÓÊÇåÇ"
         BeginProperty Font 
            Name            =   "B Homa"
            Size            =   9.75
            Charset         =   178
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   435
         Left            =   4440
         RightToLeft     =   -1  'True
         TabIndex        =   8
         Top             =   120
         Width           =   1305
      End
   End
   Begin VSFlex7LCtl.VSFlexGrid vsGood 
      Height          =   5940
      Left            =   120
      TabIndex        =   0
      Top             =   3600
      Width           =   11145
      _cx             =   19659
      _cy             =   10477
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
      BackColor       =   -2147483624
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483643
      ForeColorFixed  =   -2147483630
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483634
      BackColorAlternate=   12648447
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483642
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   50
      Cols            =   23
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frmStation_Inventory.frx":A4C2
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
      OwnerDraw       =   5
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
   Begin FLWCtrls.FWLabel3D fwlblMode 
      Height          =   495
      Left            =   0
      Top             =   0
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
      Caption         =   "ãÑæÑ"
      Alignment       =   2
   End
   Begin RESIZEKITLibCtl.ResizeKit ResizeKit1 
      Height          =   480
      Left            =   0
      OleObjectBlob   =   "frmStation_Inventory.frx":A5B6
      TabIndex        =   10
      Top             =   840
      Width           =   480
   End
End
Attribute VB_Name = "frmStation_Inventory"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private MyFormAddEditMode As EnumAddEditMode
Dim tmpTextMatrix As String
Dim Parameter() As Parameter
Dim clsDate As New clsDate
Dim i As Integer
    
Public Sub Find()
    frmFindGoods.Show vbModal
    i = vsGood.FindRow(mvarcode, 1, 1, True, True)
    
    If i > 0 Then
        vsGood.Row = i
        vsGood.ShowCell i, 0
    End If
End Sub

Public Sub ExitForm()
    Unload Me
End Sub

Private Sub FillStation()
    Dim L_Rst As New ADODB.Recordset
    
'    MsgBox "cmbSalMali.ListIndex = " & cmbSalMali.ListIndex & vbCrLf & _
'            "cmbBranch.ListIndex = " & cmbBranch.ListIndex & vbCrLf & _
'            "cmbInventory.ListIndex = " & cmbInventory.ListIndex
    
    If cmbInventory.ListIndex < 0 Or cmbBranch.ListIndex < 0 Or cmbSalMali.ListIndex < 0 Then Exit Sub
   
    lsttStations.Clear
    ReDim Parameter(0) As Parameter
    Parameter(0) = GenerateInputParameter("@Branch", adInteger, 4, cmbBranch.ItemData(cmbBranch.ListIndex))
    Set L_Rst = RunParametricStoredProcedure2Rec("Get_tStations", Parameter)
        
    If (L_Rst.EOF = True And L_Rst.BOF = True) Then
        Exit Sub
    Else
        While L_Rst.EOF = False
            lsttStations.AddItem L_Rst!Description
            lsttStations.ItemData(lsttStations.NewIndex) = L_Rst!StationId
            L_Rst.MoveNext
        Wend
    End If
    
    If L_Rst.State = adStateOpen Then L_Rst.Close
    
    For i = 0 To lsttStations.ListCount - 1
         lsttStations.Selected(i) = False
    Next i
    
    Dim a As Integer
    a = cmbInventory.ItemData(cmbInventory.ListIndex)
    
    ReDim Parameter(2) As Parameter
    Parameter(0) = GenerateInputParameter("@Branch", adInteger, 4, cmbBranch.ItemData(cmbBranch.ListIndex))
    Parameter(1) = GenerateInputParameter("@InventoryNo", adInteger, 4, cmbInventory.ItemData(cmbInventory.ListIndex))
    Parameter(2) = GenerateInputParameter("@AccountYear", adSmallInt, 2, Val(cmbSalMali.ItemData(cmbSalMali.ListIndex)))
    Set L_Rst = RunParametricStoredProcedure2Rec("Get_tStation_Inventory", Parameter)
    
    If Not (L_Rst.EOF = True And L_Rst.BOF = True) Then
        While L_Rst.EOF = False
            For i = 0 To lsttStations.ListCount - 1
                If L_Rst!StationId = lsttStations.ItemData(i) Then
                      lsttStations.Selected(i) = True
                      Exit For
                      lsttStations.Selected(i) = False
                End If
            Next i
            L_Rst.MoveNext
        Wend
    End If
    
    If L_Rst.State = adStateOpen Then L_Rst.Close
    Set L_Rst = Nothing
End Sub

Private Sub FillSalMali()
    Dim L_Rst As New ADODB.Recordset
    
    cmbSalMali.Clear
    
    Set L_Rst = RunStoredProcedure2RecordSet("Get_All_tAccountYears")
    Do While L_Rst.EOF = False
        cmbSalMali.AddItem L_Rst!AccountYear
        cmbSalMali.ItemData(cmbSalMali.NewIndex) = L_Rst!AccountYear
        L_Rst.MoveNext
    Loop
    
'    Dim i As Integer
'    For i = 0 To cmbSalMali.ListCount - 1
'        If AccountYear = cmbSalMali.ItemData(i) Then
'            cmbSalMali.ListIndex = i
'            Exit For
'        End If
'    Next
    
'    If cmbSalMali.ListCount > 0 Then cmbSalMali.ListIndex = cmbSalMali.ListCount - 1
'    MsgBox "cmbSalMali.ListIndex = " & cmbSalMali.ListIndex
    If L_Rst.State = adStateOpen Then L_Rst.Close
    Set L_Rst = Nothing
End Sub

Public Sub SetFirstToolBar()

    Dim i As Integer
    
    AllButton vbOff, True
    mdifrm.Toolbar1.Buttons(6).Enabled = False   'Add
    mdifrm.Toolbar1.Buttons(7).Enabled = True   'Edit
    mdifrm.Toolbar1.Buttons(8).Enabled = False   'Enter
    mdifrm.Toolbar1.Buttons(9).Enabled = False   'Cancel
    mdifrm.Toolbar1.Buttons(10).Enabled = False   'Delete
    
    mdifrm.Toolbar1.Buttons(13).Enabled = True   'Find
    
    mdifrm.Toolbar1.Buttons(15).Enabled = True  'printing
    mdifrm.Toolbar1.Buttons(23).Enabled = True  'Language
    mdifrm.Toolbar1.Buttons(24).Enabled = True  'PhoneBook
    mdifrm.Toolbar1.Buttons(25).Enabled = False 'Keyboard
    mdifrm.Toolbar1.Buttons(26).Enabled = True  'Calculator
    mdifrm.Toolbar1.Buttons(27).Enabled = True  '
    
    Select Case MyFormAddEditMode
        Case EnumAddEditMode.ViewMode
            mdifrm.Toolbar1.Buttons(6).Enabled = False 'Add
            mdifrm.Toolbar1.Buttons(7).Enabled = True  'Edit
            mdifrm.Toolbar1.Buttons(8).Enabled = False  'Enter
            mdifrm.Toolbar1.Buttons(9).Enabled = False  'Cancel
            mdifrm.Toolbar1.Buttons(10).Enabled = False  'Delete
            vsGood.Editable = flexEDNone
            
        Case EnumAddEditMode.AddMode
            mdifrm.Toolbar1.Buttons(6).Enabled = False 'Add
            mdifrm.Toolbar1.Buttons(7).Enabled = False  'Edit
            mdifrm.Toolbar1.Buttons(8).Enabled = True  'Enter
            mdifrm.Toolbar1.Buttons(9).Enabled = True  'Cancel
            mdifrm.Toolbar1.Buttons(10).Enabled = False  'Delete

 '           vsGood.Editable = flexEDKbdMouse
            
        Case EnumAddEditMode.EditMode
            mdifrm.Toolbar1.Buttons(6).Enabled = False 'Add
            mdifrm.Toolbar1.Buttons(7).Enabled = False  'Edit
            mdifrm.Toolbar1.Buttons(8).Enabled = True  'Enter
            mdifrm.Toolbar1.Buttons(9).Enabled = True  'Cancel
            mdifrm.Toolbar1.Buttons(10).Enabled = False  'Delete

'            vsGood.Editable = flexEDKbdMouse
    End Select
    
    HeaderLabel Val(MyFormAddEditMode), fwlblMode
End Sub

Public Sub DefaultSetting()
    vsGood.Rows = 1
    
    If cmbInventory.ListIndex > -1 And cmbBranch.ListIndex > -1 Then
        FillStation
    End If
       '' ReDim Parameter(0) As Parameter
        ''Parameter(0) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
    CmdStationAdd.Enabled = False
    cmdStationDelete.Enabled = False
End Sub

Public Sub FillvsGood() 'it fills the grid using vw_Good
    If ValidateUIData() = False Then Exit Sub
    
'    If cmbInventory.ListIndex < 0 Or cmbBranch.ListIndex < 0 Or cmbSalMali.ListIndex < 0 Then Exit Sub

    MyFormAddEditMode = EnumAddEditMode.ViewMode 'VIEW Mode
    SetFirstToolBar
    
    vsGood.Rows = 1
    
    Dim i As Long
    Dim j As Integer
    Dim intSelectedLevel1 As Integer
    Dim intSelectedLevel2 As Integer
    Dim strSelectedLevels As String
    Dim Rst As New ADODB.Recordset
    Dim Rst2 As New ADODB.Recordset
    
    intSelectedLevel1 = -1
    intSelectedLevel2 = -1
    
    Dim SelectedGroups As String
    ReDim Parameter(1) As Parameter
    For i = 0 To lsttStations.ListCount - 1
        If lsttStations.Selected(i) = True Then
            SelectedGroups = SelectedGroups & lsttStations.ItemData(i) & ","
        End If
    Next i
    
    If Len(SelectedGroups) > 0 Then
        SelectedGroups = Left(SelectedGroups, Len(SelectedGroups) - 1)
    Else
        SelectedGroups = ""
    End If
    
    ReDim Parameter(3) As Parameter
    Parameter(0) = GenerateInputParameter("@InventoryNo", adInteger, 4, cmbInventory.ItemData(cmbInventory.ListIndex))
    Parameter(1) = GenerateInputParameter("@Branch", adInteger, 4, cmbBranch.ItemData(cmbBranch.ListIndex))
    Parameter(2) = GenerateInputParameter("@StationID", adVarWChar, 400, SelectedGroups)
    Parameter(3) = GenerateInputParameter("@AccountYear", adSmallInt, 2, Val(cmbSalMali.Text))
    Set Rst = RunParametricStoredProcedure2Rec("GetStation_Inventory_Goods", Parameter)
    
    If (Rst.EOF = True And Rst.BOF = True) Then Exit Sub
    
    With vsGood
''''        Dim jj As Integer
        i = 1
        While Rst.EOF = False
            .Rows = .Rows + 1
            .TextMatrix(i, 0) = i
''''            .Cell(flexcpText, i, 1) = CStr(Rst.Fields("Branch").Value)
''''            .Cell(flexcpText, i, 2) = CStr(Rst.Fields("InventoryNo").Value)
            .Cell(flexcpText, i, 3) = CStr(Rst.Fields("StationId").Value)
            .TextMatrix(i, 4) = Rst.Fields("GoodCode").Value
            .TextMatrix(i, 5) = Left(Rst.Fields("Name").Value, 40)
            .TextMatrix(i, 6) = IIf(Rst.Fields("Active").Value = True, -1, 0)
            
            i = i + 1
''''            jj = i Mod 10
''''            If jj = 0 Then
''''                FWOdometter1.Value = FWOdometter1.Value + 1
''''            End If
''''            If FWOdometter1.Value = 100 Then
''''                FWOdometter1.Value = 0
''''            End If
            Rst.MoveNext
            
        Wend
        
        If Rst.State = adStateOpen Then Rst.Close
        Set Rst = Nothing
        
        .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter
        If .Rows > 1 Then
            .Cell(flexcpAlignment, 1, 2, .Rows - 1, 2) = flexAlignRightCenter
        End If
    '    .AutoSizeMode = flexAutoSizeColWidth
    '    .AutoSize 0, .Cols - 1
        
    End With
End Sub

Public Sub BeforeUpdate()

End Sub

Public Sub Edit()
    
    With vsGood
 '       .Editable = flexEDKbdMouse
        MyFormAddEditMode = EnumAddEditMode.EditMode
        SetFirstToolBar
    End With
    
    CmdStationAdd.Enabled = True
    cmdStationDelete.Enabled = True
End Sub

Public Sub Update()
    On Error GoTo Err_Handler
    
    Dim i As Integer
    Dim j As Integer
    Dim LongTemp As Integer
    Dim lngSelectedSubGroup  As Long
    
    Dim Rst As New ADODB.Recordset
        
    lngSelectedSubGroup = -1
    
    If MyFormAddEditMode = EnumAddEditMode.ViewMode Then Exit Sub
    
    vsGood_ValidateEdit vsGood.Row, vsGood.Col, False
    
    With vsGood
        If .Rows < 2 Then
            MyFormAddEditMode = EnumAddEditMode.ViewMode
            SetFirstToolBar
            Exit Sub
        End If
        
        For i = 1 To .Rows - 1
            .Row = i
            If InStr(.TextMatrix(i, 0), "*") > 0 Then 'new or edited records
            
''''                If ((Trim(.TextMatrix(i, 2)) = "" And Trim(.TextMatrix(i, 3)) = "") Or Trim(.TextMatrix(i, 5)) = "") Or .Cell(flexcpText, i, 8) = "" Or .Cell(flexcpText, i, 7) = "" Then
''''
''''                    Select Case clsStation.Language
''''
''''                        Case 0
''''
''''                            frmMsg.fwlblMsg.Caption = "ÔãÇ ãí ÈÇíÓÊ ÇØáÇÚÇÊ ÑÇ ÈØæÑ ˜Çãá æÇÑÏ äãÇííÏ"
''''                            frmMsg.Fwbtn(0).Caption = "ÞÈæá"
''''                        Case 1
''''
''''                            frmMsg.fwlblMsg.Caption = "You Have to complete the information"
''''                            frmMsg.Fwbtn(0).Caption = "Ok"
''''                            frmMsg.fwlblMsg.Alignment = vbLeftJustify
''''
''''                    End Select
''''
''''                    frmMsg.Fwbtn(0).ButtonType = flwButtonOk
''''                    frmMsg.Fwbtn(1).Visible = False
''''                    frmMsg.Show vbModal
''''
''''                    Exit Sub
''''
''''                End If
                
                
            End If
        Next i
        

        Select Case MyFormAddEditMode
        
                
            Case EnumAddEditMode.EditMode
                
                For i = 1 To .Rows - 1
                    
                    If InStr(.TextMatrix(i, 0), "*") > 0 Then 'Edited records
                            
                        ReDim Parameter(5) As Parameter

                        Parameter(0) = GenerateInputParameter("@Branch", adInteger, 4, cmbBranch.ItemData(cmbBranch.ListIndex))
                        Parameter(1) = GenerateInputParameter("@InventoryNo", adInteger, 4, cmbInventory.ItemData(cmbInventory.ListIndex))
                        Parameter(2) = GenerateInputParameter("@StationId", adInteger, 4, Val(.TextMatrix(i, 3)))
                        Parameter(3) = GenerateInputParameter("@GoodCode", adInteger, 4, Val(.TextMatrix(i, 4)))
                        Parameter(4) = GenerateInputParameter("@Active", adBoolean, 1, IIf(Val(.TextMatrix(i, 6)) = -1, 1, 0))
                        Parameter(5) = GenerateInputParameter("@AccountYear", adSmallInt, 2, Val(cmbSalMali.Text))
                        
                        RunParametricStoredProcedure "Update_tStation_Inventory_Good", Parameter
                            
                    End If
                                        
                Next i
                
            
            End Select
            
        FillvsGood
        
    End With
    
    CmdStationAdd.Enabled = False
    cmdStationDelete.Enabled = False
    
    If Rst.State = adStateOpen Then Rst.Close
    Set Rst = Nothing
    
    Exit Sub
Err_Handler:
    LogSaveNew "frmStation_Inventory => ", err.Description, err.Number, err.Source, "Update"
    ShowErrorMessage
    err.Clear
End Sub

Public Sub Cancel()
    MyFormAddEditMode = EnumAddEditMode.ViewMode
    SetFirstToolBar
    FillvsGood
End Sub

Private Sub CheckFirstMojodi_Click()
    FillvsGood
End Sub

Private Sub CheckOrder_Click()
    FillvsGood
End Sub

Private Sub cmbBranch_Click()
    FillInventory
End Sub

Private Sub cmbInventory_Click()
    If cmbBranch.ListIndex < 0 Then Exit Sub
    
    FillStation
    FillvsGood
End Sub

'Private Sub cmbSalMali_Change()
'    cmbInventory_Click
'End Sub

Private Sub cmbSalMali_Click()
'    cmbSalMali_Change
    cmbInventory_Click
End Sub

Private Sub CmdStationAdd_Click()
    Dim i As Integer
    
'    If lsttStations.SelCount = 0 Then Exit Sub
    Select Case MyFormAddEditMode
        Case EditMode
            Dim SelectedGroups As String
            For i = 0 To lsttStations.ListCount - 1
                If lsttStations.Selected(i) = True Then
                    SelectedGroups = SelectedGroups & lsttStations.ItemData(i) & ","
                End If
            Next i
        
            If Len(SelectedGroups) > 0 Then
                SelectedGroups = Left(SelectedGroups, Len(SelectedGroups) - 1)
            Else
                SelectedGroups = ""
            End If
        
            If cmbInventory.ListIndex > -1 Then
                ReDim Parameter(3) As Parameter
                Parameter(0) = GenerateInputParameter("@Branch", adInteger, 4, cmbBranch.ItemData(cmbBranch.ListIndex))
                Parameter(1) = GenerateInputParameter("@InventoryNo", adInteger, 4, cmbInventory.ItemData(cmbInventory.ListIndex))
                Parameter(2) = GenerateInputParameter("@AccountYear", adSmallInt, 2, Val(cmbSalMali.Text))
                Parameter(3) = GenerateInputParameter("@StationID", adVarWChar, 400, SelectedGroups)
                
                RunParametricStoredProcedure "Update_tStation_Inventory", Parameter
                
                ShowDisMessage " ÇÖÇÝå ˜ÑÏä ÇíÓÊÇååÇ Èå ÇäÈÇÑ ÇäÌÇã ÔÏ", 2000
                GotoViewMode
                CmdStationAdd.Enabled = False
                cmdStationDelete.Enabled = False
                FillStation
                FillvsGood
            End If
           'End of case EditMode
    End Select
End Sub

Private Sub GotoViewMode()
    MyFormAddEditMode = ViewMode
    SetFirstToolBar
End Sub

Private Sub cmdStationDelete_Click()
    Dim i As Integer
    
    Select Case MyFormAddEditMode
      Case EditMode
        Dim SelectedGroups As String
        For i = 0 To lsttStations.ListCount - 1
            If lsttStations.Selected(i) = True Then
                SelectedGroups = SelectedGroups & lsttStations.ItemData(i) & ","
            End If
        Next i
        
        If Len(SelectedGroups) > 0 Then
            SelectedGroups = Left(SelectedGroups, Len(SelectedGroups) - 1)
        Else
            SelectedGroups = ""
        End If
        
        If cmbInventory.ListIndex > -1 And cmbBranch.ListIndex > -1 Then
            ReDim Parameter(3) As Parameter
            Parameter(0) = GenerateInputParameter("@Branch", adInteger, 4, cmbBranch.ItemData(cmbBranch.ListIndex))
            Parameter(1) = GenerateInputParameter("@InventoryNo", adInteger, 4, cmbInventory.ItemData(cmbInventory.ListIndex))
            Parameter(2) = GenerateInputParameter("@StationID", adVarWChar, 400, SelectedGroups)
            Parameter(3) = GenerateInputParameter("@AccountYear", adSmallInt, 2, Val(cmbSalMali.Text))
            
            RunParametricStoredProcedure "Delete_tStation_Inventory", Parameter
            
            ShowDisMessage "ÍÐÝ ÇíÓÊÇååÇ ÇÒÇäÈÇÑ ÇäÌÇã ÔÏ ", 2000
            GotoViewMode
            CmdStationAdd.Enabled = False
            cmdStationDelete.Enabled = False
            FillStation
            FillvsGood
        End If
        'End of case EditMode
     End Select
End Sub

Private Sub Form_Activate()
    VarActForm = Me.Name
    
    Frame3.BackColor = Me.BackColor
    Frame28.BackColor = Me.BackColor
'    FWOdometter1.Value = 0
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call PresetScreenSaver
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
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
    If ClsFormAccess.frmStation_Inventory = False Then
        Unload Me
        Exit Sub
    End If
    
    CenterTop Me
    VarActForm = Me.Name
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
    
    ChangeLanguage
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    VarActForm = ""
    If PosConnection.State = adStateOpen Then PosConnection.Close: Set PosConnection = Nothing

    Dim i As Integer
    
    Unload frmFindGoods

    SaveSetting strMainKey, Me.Name, "Left", Me.Left
    SaveSetting strMainKey, Me.Name, "Top", Me.Top
End Sub

Public Sub ChangeLanguage()
    On Error Resume Next
    
    Call FormActivateOperations
    
    Dim Obj As Object

    Select Case clsStation.Language    ' LCase(mdifrm.Toolbar1.Buttons(25).Key)
        Case English
            Me.Caption = "Assign Station to Inventory"
            mdifrm.Caption = clsArya.LatinCompany
            Me.RightToLeft = False
            
            For Each Obj In Me
                On Error Resume Next
                    Obj.RightToLeft = False
                'On Error GoTo 0
            Next Obj
            
        Case Farsi
            Me.Caption = "                                                                  ÇÎÊÕÇÕ ÇäÈÇÑ Èå ÇíÓÊÇå     "
            mdifrm.Caption = clsArya.Company
            Me.RightToLeft = True
            
            For Each Obj In Me
                On Error Resume Next
                    Obj.RightToLeft = True
                On Error GoTo 0
            Next Obj
            
    End Select

Exit Sub
Err_Handler:
    LogSaveNew "frmStation_Inventory => ", err.Description, err.Number, err.Source, "ChangeLanguage"
    err.Clear
End Sub

Private Sub FormActivateOperations()
    On Error Resume Next
    
    FillSalMali
    FillBranch
    FillInventory
    
    With vsGood
        .Rows = 1
        .Cols = 7
        
        Select Case clsStation.Language
            Case Farsi
                .TextMatrix(0, 0) = "ÑÏíÝ"
                .TextMatrix(0, 1) = "ÔÚÈå"
                .TextMatrix(0, 2) = "ÇäÈÇÑ"
                .TextMatrix(0, 3) = "ÇíÓÊÇå "
                .TextMatrix(0, 4) = "˜Ï ßÇáÇ"
                .TextMatrix(0, 5) = "äÇã ßÇáÇ"
                .TextMatrix(0, 6) = " ßäÊÑá "
            
            Case English
                .TextMatrix(0, 0) = "Row"
                .TextMatrix(0, 1) = "Branch"
                .TextMatrix(0, 2) = "Inventory"
                .TextMatrix(0, 3) = " Station"
                .TextMatrix(0, 4) = " GoodCode"
                .TextMatrix(0, 5) = "GoodName"
                .TextMatrix(0, 6) = "Control"
       End Select
       
        .ColWidth(3) = 2000
        .ColWidth(5) = 4500
        .ColDataType(6) = flexDTBoolean
   '     .ColSort(5) = flexSortNumericAscending + flexSortNumericDescending
        .ColAlignment(-1) = flexAlignCenterCenter
        .ColAlignment(5) = flexAlignRightCenter
        .FocusRect = flexFocusHeavy
        .ColHidden(1) = True
        .ColHidden(2) = True
      '  .ColHidden(4) = True
        .AutoSizeMode = flexAutoSizeColWidth
     '   .AutoSize 5, 5
        .AutoSearch = flexSearchFromCursor
        
        Set rctmp = RunStoredProcedure2RecordSet("Get_All_Branches")
        .ColComboList(1) = .BuildComboList(rctmp, "nvcBranchName", "Branch")
        rctmp.Close
        
        Set rctmp = RunStoredProcedure2RecordSet("Get_All_tinventory")
        .ColComboList(2) = .BuildComboList(rctmp, "Description", "InventoryNo")
        rctmp.Close
        
        ReDim Parameter(0) As Parameter
        Parameter(0) = GenerateInputParameter("@Branch", adInteger, 4, cmbBranch.ItemData(cmbBranch.ListIndex))
        Set rctmp = RunParametricStoredProcedure2Rec("Get_tStations", Parameter)
        .ColComboList(3) = .BuildComboList(rctmp, "Description", "StationId")
        rctmp.Close
    
    End With
    
    DefaultSetting
    SetFirstToolBar
End Sub

Private Sub FillBranch()
    On Error Resume Next
    
    Dim L_Rst As New ADODB.Recordset
    
    cmbBranch.Clear
    
    Set L_Rst = RunStoredProcedure2RecordSet("Get_All_Branches")
    
    Do While L_Rst.EOF = False
        cmbBranch.AddItem L_Rst!nvcBranchName
        cmbBranch.ItemData(cmbBranch.NewIndex) = L_Rst!Branch
        L_Rst.MoveNext
    Loop
    
'    Dim i As Integer
'    For i = 0 To cmbBranch.ListCount - 1
'        If CurrentBranch = cmbBranch.ItemData(i) Then
'            cmbBranch.ListIndex = i
'            Exit For
'        End If
'    Next i
    
    If L_Rst.State = adStateOpen Then L_Rst.Close
    Set L_Rst = Nothing
'    MsgBox "cmbBranch.ListIndex = " & cmbBranch.ListIndex
'    cmbBranch.ListIndex = -1
'    If cmbBranch.ListCount > 0 Then cmbBranch.ListIndex = 0
End Sub

Private Sub FillInventory()
    Dim L_Rst As New ADODB.Recordset
    
    cmbInventory.Clear
    lsttStations.Clear

    If cmbBranch.ListIndex < 0 Then Exit Sub
    
    cmbInventory.Clear
    
    ReDim Parameter(1) As Parameter
    Parameter(0) = GenerateInputParameter("@intLanguage", adInteger, 4, clsStation.Language)
    Parameter(1) = GenerateInputParameter("@Branch", adInteger, 4, cmbBranch.ItemData(cmbBranch.ListIndex))
    Set L_Rst = RunParametricStoredProcedure2Rec("GetInventory_Branch", Parameter)
    
    If Not (L_Rst.EOF = True And L_Rst.BOF = True) Then
        Do While L_Rst.EOF = False
            cmbInventory.AddItem L_Rst!Description
            cmbInventory.ItemData(cmbInventory.NewIndex) = Val(L_Rst!InventoryNo)
            L_Rst.MoveNext
        Loop
    End If
    
    If L_Rst.State = adStateOpen Then L_Rst.Close
    Set L_Rst = Nothing
    
'    cmbInventory.ListIndex = -1
'    If cmbInventory.ListCount > 0 Then cmbInventory.ListIndex = 0
'    MsgBox "cmbInventory.ListIndex = " & cmbInventory.ListIndex
    FillStation
End Sub

Private Sub lsttStations_Click()
    FillvsGood
End Sub

Private Sub ResizeKit1_ExitResize(ByVal XScale As Double, ByVal YScale As Double)
    If formloadFlag = True Then
        SaveSetting strMainKey, Me.Name, "Height", Me.Height
        SaveSetting strMainKey, Me.Name, "Width", Me.Width
    End If

End Sub

Private Sub vsGood_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    With vsGood
        If (.TextMatrix(Row, 0) = "*" Or MyFormAddEditMode = EnumAddEditMode.EditMode) And Col > 1 And tmpTextMatrix <> .TextMatrix(Row, Col) Then
        
            If MyFormAddEditMode = EnumAddEditMode.EditMode And InStr(.TextMatrix(Row, 0), "*") = 0 Then
                .TextMatrix(Row, 0) = Trim(.TextMatrix(Row, 0)) & "*"
            End If
            
        Else

        End If
     '   .AutoSizeMode = flexAutoSizeColWidth
     '   .AutoSize 0, .Cols - 1

    End With
End Sub

Private Sub vsGood_KeyDown(KeyCode As Integer, Shift As Integer)
    With vsGood
        If .Col = 6 Then
           .Select .Row, .Col
           .EditCell
        End If
    End With
    
End Sub

Private Sub vsGood_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    With vsGood
        If .Col = 6 Then
           .Select .Row, .Col
           .EditCell
        End If
    End With

End Sub

Private Sub vsGood_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsGood
        .Row = Row
        .Col = Col
    End With
End Sub

Public Sub Printing()
End Sub

Private Function ValidateUIData() As Boolean
    Dim blnResult As Boolean
    blnResult = False
    
    If cmbSalMali.ListIndex < 0 Then
        ShowMessage "áØÝÇð ÓÇá ãÇáí ÑÇ ÇäÊÎÇÈ ßäíÏ", True, False, "ÊÇííÏ", ""
    ElseIf cmbBranch.ListIndex < 0 Then
        ShowMessage "áØÝÇð ÔÚÈå ÑÇ ÇäÊÎÇÈ ßäíÏ ", True, False, "ÊÇííÏ", ""
    ElseIf cmbInventory.ListIndex < 0 Then
        ShowMessage "áØÝÇð ÇäÈÇÑ ãæÑÏ äÙÑ ÑÇ ÇäÊÎÇÈ ßäíÏ", True, False, "ÊÇííÏ", ""
    Else
        blnResult = True
    End If
    
    ValidateUIData = blnResult
End Function
