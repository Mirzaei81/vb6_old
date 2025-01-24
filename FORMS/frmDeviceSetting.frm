VERSION 5.00
Object = "{C0A63B80-4B21-11D3-BD95-D426EF2C7949}#1.0#0"; "Vsflex7L.ocx"
Object = "{7AEDC602-D94C-11D1-BB7A-00E0290EA3C9}#1.0#0"; "ResizeKit.ocx"
Object = "{75D4F767-8785-11D3-93AD-0000832EF44D}#2.14#0"; "FAST2001.ocx"
Begin VB.Form frmDeviceSetting 
   ClientHeight    =   7005
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   12915
   Icon            =   "frmDeviceSetting.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   RightToLeft     =   -1  'True
   ScaleHeight     =   7005
   ScaleWidth      =   12915
   Begin FLWCtrls.FWLabel3D fwlblMode 
      Height          =   495
      Left            =   11520
      Top             =   0
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   873
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "B Homa"
         Size            =   14.25
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
   Begin VSFlex7LCtl.VSFlexGrid VSDeviceSetting 
      Height          =   5865
      Left            =   0
      TabIndex        =   0
      Top             =   960
      Width           =   12735
      _cx             =   22463
      _cy             =   10345
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
      Rows            =   50
      Cols            =   9
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmDeviceSetting.frx":A4C2
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
      ExplorerBar     =   0
      PicturesOver    =   0   'False
      FillStyle       =   0
      RightToLeft     =   -1  'True
      PictureType     =   0
      TabBehavior     =   0
      OwnerDraw       =   0
      Editable        =   2
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
   Begin FLWCtrls.FWLabel FWLabel1 
      Height          =   495
      Left            =   0
      Top             =   0
      Width           =   12885
      _ExtentX        =   22728
      _ExtentY        =   873
      Enabled         =   -1  'True
      Caption         =   " ‰ŸÌ„ Ê”«Ì· Ã«‰»Ì"
      FirstColor      =   9412754
      SecondColor     =   14215660
      Angle           =   0
      ForeColor       =   7362318
      BackColor       =   12640511
      FontName        =   "B Homa"
      FontItalic      =   -1  'True
      FontSize        =   15.75
      Alignment       =   2
      Picture         =   "frmDeviceSetting.frx":A5D5
   End
   Begin RESIZEKITLibCtl.ResizeKit ResizeKit1 
      Height          =   480
      Left            =   120
      OleObjectBlob   =   "frmDeviceSetting.frx":A5F1
      TabIndex        =   1
      Top             =   360
      Width           =   480
   End
End
Attribute VB_Name = "frmDeviceSetting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MyFormAddEditMode As EnumAddEditMode
Dim i As Integer
Dim Parameter() As Parameter


Public Sub Add()

    With VSDeviceSetting
        .Rows = .Rows + 1
        .TextMatrix(.Rows - 1, 0) = "*"
    End With
    MyFormAddEditMode = AddMode
    SetFirstToolBar
    
End Sub
Public Sub Cancel()

    MyFormAddEditMode = EnumAddEditMode.ViewMode
    SetFirstToolBar
    
    FillVSDeviceSetting
    
End Sub
Public Sub Delete()

    With VSDeviceSetting
        If .SelectedRows < 1 Then
            frmMsg.fwlblMsg.Caption = "·ÿ›« «» œ« Ìò Ì« ç‰œ Ê”Ì·Â —« «‰ Œ«» ‰„«ÌÌœ"
            frmMsg.fwBtn(0).Caption = "ﬁ»Ê·"
            frmMsg.Show vbModal
            Exit Sub
        End If
        
        Dim s As String
        Dim S2 As String
        
        s = ""
        For i = 0 To .SelectedRows - 1
            s = s & .TextMatrix(.SelectedRow(i), 0) & " ,"
            S2 = S2 & .TextMatrix(.SelectedRow(i), 8) & ","
        Next i
        s = Left(s, Len(s) - 1)
        S2 = Left(S2, Len(S2) - 1)
        
        frmMsg.fwlblMsg.Caption = "¬Ì« „ÿ„∆‰Ìœ „Ì ŒÊ«ÂÌœ Ê”Ì·Â (Ê”«Ì·) Ã«‰»Ì " & s & " —« Õ–› ﬂ‰Ìœø"
        frmMsg.fwBtn(0).Caption = "»·Ì"
        frmMsg.fwBtn(1).Caption = "ŒÌ—"
        frmMsg.Show vbModal
        
        If modgl.mvarMsgIdx = vbNo Then
           Exit Sub
        End If
        
        ReDim Parameter(1) As Parameter
        
        Parameter(0) = GenerateInputParameter("@String", adVarWChar, 4000, S2)
        Parameter(1) = GenerateOutputParameter("@Result", adInteger, 4)
        If RunParametricStoredProcedure("Delete_DeviceSetting", Parameter) = -1 Then
            frmMsg.fwlblMsg.Caption = "Õ–› «‰Ã«„ ‰‘œ"
            frmMsg.fwBtn(0).Caption = "ﬁ»Ê·"
            frmMsg.Show vbModal
        Else
            frmMsg.fwlblMsg.Caption = "Õ–› »« „Ê›ﬁÌ  «‰Ã«„ ‘œ"
            frmMsg.fwBtn(0).Caption = "ﬁ»Ê·"
            frmMsg.Show vbModal
        End If
        
    End With
    Cancel
End Sub

Public Sub Edit()
        MyFormAddEditMode = EnumAddEditMode.EditMode
        SetFirstToolBar
End Sub

Public Sub ExitForm()

    Unload Me
    
End Sub

Public Sub FillVSDeviceSetting()

    MyFormAddEditMode = EnumAddEditMode.ViewMode
    SetFirstToolBar
    
    Dim Rst As New ADODB.Recordset
    
    With VSDeviceSetting
        .Rows = 1
        Set Rst = RunStoredProcedure2RecordSet("Get_DeviceSettingAll")
        
        While Rst.EOF <> True
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, 0) = .Rows - 1
            .TextMatrix(.Rows - 1, 1) = Rst.Fields("StationId").Value
            .TextMatrix(.Rows - 1, 2) = Rst.Fields("DeviceTypeCode").Value
            .TextMatrix(.Rows - 1, 3) = Rst.Fields("DeviceCode").Value
            .TextMatrix(.Rows - 1, 4) = Rst.Fields("PortCode").Value
            .TextMatrix(.Rows - 1, 5) = Rst.Fields("BaudRate").Value
            .TextMatrix(.Rows - 1, 6) = Rst.Fields("BufferSize").Value
            .TextMatrix(.Rows - 1, 7) = Rst.Fields("RThreshold").Value
            .TextMatrix(.Rows - 1, 8) = Rst.Fields("Code").Value
            Rst.MoveNext
        Wend
        
    End With
    Set Rst = Nothing
    
End Sub

Public Sub SetFirstToolBar()

    Dim i As Integer
    
    AllButton vbOff, True
    
    mdifrm.Toolbar1.Buttons(23).Enabled = True
    mdifrm.Toolbar1.Buttons(24).Enabled = True
    mdifrm.Toolbar1.Buttons(25).Enabled = True
    mdifrm.Toolbar1.Buttons(26).Enabled = True
    mdifrm.Toolbar1.Buttons(27).Enabled = True
    
    VSDeviceSetting.Editable = flexEDNone
    
    Select Case MyFormAddEditMode
    
        Case EnumAddEditMode.ViewMode
        
            For i = 6 To 10
                mdifrm.Toolbar1.Buttons(i).Enabled = True
            Next i
            
        Case EnumAddEditMode.AddMode
        
            mdifrm.Toolbar1.Buttons(6).Enabled = True 'add key
            mdifrm.Toolbar1.Buttons(8).Enabled = True 'enter key
            mdifrm.Toolbar1.Buttons(9).Enabled = True 'cancel key
'            VSDeviceSetting.Editable = flexEDKbdMouse
            
        Case EnumAddEditMode.EditMode
                    
            mdifrm.Toolbar1.Buttons(8).Enabled = True 'enter key
            mdifrm.Toolbar1.Buttons(9).Enabled = True 'cancel key
'            VSDeviceSetting.Editable = flexEDKbdMouse
    End Select
    
    HeaderLabel Val(MyFormAddEditMode), fwlblMode
    
End Sub

Public Sub DefaultSetting()

    Dim Rst As New ADODB.Recordset
    Dim s As String
    
    With VSDeviceSetting
    
        s = ""
        If Rst.State <> 0 Then
            Rst.Close
        End If
        Set Rst = RunStoredProcedure2RecordSet("Get_PC_Stations")
        s = .BuildComboList(Rst, "Description", "StationId")
        .ColComboList(1) = s
        
        
        s = ""
        If Rst.State <> 0 Then
            Rst.Close
        End If
        
        Set Rst = RunStoredProcedure2RecordSet("Get_All_DeviceType")
        If clsStation.Language = Farsi Then
            s = .BuildComboList(Rst, "DeviceTypeName", "DeviceTypeCode")
        Else
            s = .BuildComboList(Rst, "DeviceTypeLatinName", "DeviceTypeCode")
        End If
        .ColComboList(2) = s

        s = ""
        If Rst.State <> 0 Then
            Rst.Close
        End If
        Set Rst = RunStoredProcedure2RecordSet("Get_All_Device")
        If clsStation.Language = Farsi Then
            s = .BuildComboList(Rst, "DeviceName", "DeviceCode")
        Else
            s = .BuildComboList(Rst, "DeviceLatinName", "DeviceCode")
        End If
        .ColComboList(3) = s

        s = ""
        If Rst.State <> 0 Then
            Rst.Close
        End If
        Set Rst = RunStoredProcedure2RecordSet("Get_All_Ports")
        s = .BuildComboList(Rst, "PortName", "PortCode")
        .ColComboList(4) = s
        
        s = ""
        If Rst.State <> 0 Then
            Rst.Close
        End If
        Set Rst = RunStoredProcedure2RecordSet("Get_All_BaudRate")
        s = .BuildComboList(Rst, "BaudRate")
        .ColComboList(5) = s
        
    End With
    Set Rst = Nothing
     
    FillVSDeviceSetting
    
End Sub

Public Sub Update()
    
    Dim j As Integer
    With VSDeviceSetting
        If .Rows < 2 Then Exit Sub
        
        Select Case MyFormAddEditMode
            Case AddMode
                For i = .Rows - 1 To 1 Step -1
                    If .TextMatrix(i, 0) = "*" Then
                        j = i
                        Exit For
                    End If
                Next i
                If j <> 0 Then
                    ReDim Parameter(4) As Parameter
                    Parameter(0) = GenerateInputParameter("@DeviceCode", adInteger, 4, .TextMatrix(j, 3))
                    Parameter(1) = GenerateInputParameter("@StationId", adInteger, 4, .TextMatrix(j, 1))
                    Parameter(2) = GenerateInputParameter("@PortCode", adInteger, 4, .TextMatrix(j, 4))
                    Parameter(3) = GenerateInputParameter("@BaudRate", adInteger, 4, .TextMatrix(j, 5))
                    Parameter(4) = GenerateOutputParameter("@Code", adInteger, 4)
                    
                    If RunParametricStoredProcedure("Insert_DeviceSetting", Parameter) = -1 Then
                    
                    End If
                End If
            Case EditMode
            
                For i = .Rows - 1 To 1 Step -1
                    If InStr(1, .TextMatrix(i, 0), "*") <> 0 Then
                        j = i
                        Exit For
                    End If
                Next i
                If j <> 0 Then
                    ReDim Parameter(5) As Parameter
                    Parameter(0) = GenerateInputParameter("@DeviceCode", adInteger, 4, .TextMatrix(j, 3))
                    Parameter(1) = GenerateInputParameter("@StationId", adInteger, 4, .TextMatrix(j, 1))
                    Parameter(2) = GenerateInputParameter("@PortCode", adInteger, 4, .TextMatrix(j, 4))
                    Parameter(3) = GenerateInputParameter("@BaudRate", adInteger, 4, .TextMatrix(j, 5))
                    Parameter(4) = GenerateInputParameter("@Code", adInteger, 4, .TextMatrix(j, 8))
                    Parameter(5) = GenerateOutputParameter("@Result", adInteger, 4)
                    If RunParametricStoredProcedure("Update_DeviceSetting", Parameter) = -1 Then
                        
                    End If
                End If
        End Select
        FillVSDeviceSetting
    End With
End Sub

Private Sub Form_Activate()
    VarActForm = Me.Name
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

Private Sub Form_Load()

    CenterCenter Me
    VarActForm = Me.Name
    
    
    DefaultSetting
    
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

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call PresetScreenSaver
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call PresetScreenSaver
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Integer
    
    AllButton vbOff, True
    
    If PosConnection.State = adStateOpen Then PosConnection.Close: Set PosConnection = Nothing

    SaveSetting strMainKey, Me.Name, "Left", Me.Left
    SaveSetting strMainKey, Me.Name, "Top", Me.Top
    VarActForm = ""

End Sub

Private Sub ResizeKit1_ExitResize(ByVal XScale As Double, ByVal YScale As Double)

    If formloadFlag = True Then
        SaveSetting strMainKey, Me.Name, "Height", Me.Height
        SaveSetting strMainKey, Me.Name, "Width", Me.Width
    End If

End Sub

Private Sub VSDeviceSetting_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim Rst As New ADODB.Recordset
    Dim s As String
    
    With VSDeviceSetting
        If .Row > 0 Then
            Select Case .Col
                Case 1
                    s = ""
                    If Rst.State <> 0 Then
                        Rst.Close
                    End If
                    Set Rst = RunStoredProcedure2RecordSet("Get_PC_Stations")
                    s = .BuildComboList(Rst, "Description", "StationId")
                    .ColComboList(.Col) = s
                
                    
                Case 3
                
                    s = ""
                    If Rst.State <> 0 Then
                        Rst.Close
                    End If
                    Set Rst = RunStoredProcedure2RecordSet("Get_All_Device")
                    If clsStation.Language = Farsi Then
                        s = .BuildComboList(Rst, "DeviceName", "DeviceCode")
                    Else
                        s = .BuildComboList(Rst, "DeviceLatinName", "DeviceCode")
                    End If
                    .ColComboList(.Col) = s
                    
                
                Case 4
                
                    s = ""
                    If Rst.State <> 0 Then
                        Rst.Close
                    End If
                    Set Rst = RunStoredProcedure2RecordSet("Get_All_Ports")
                    s = .BuildComboList(Rst, "PortName", "PortCode")
                    .ColComboList(.Col) = s
                
                Case 5
                
                    s = ""
                    If Rst.State <> 0 Then
                        Rst.Close
                    End If
                    Set Rst = RunStoredProcedure2RecordSet("Get_All_BaudRate")
                    s = .BuildComboList(Rst, "BaudRate")
                    .ColComboList(.Col) = s
                    
                    
            End Select
        End If
    End With
    Set Rst = Nothing

End Sub



Private Sub VSDeviceSetting_ChangeEdit()

    Dim Rst As New ADODB.Recordset
    With VSDeviceSetting
        If .Row > 0 And .ColComboList(.Col) <> "" Then
            VSDeviceSetting_ValidateEdit .Row, .Col, False
            Select Case .Col
                Case 3
                    s = ""
                    If Rst.State <> 0 Then
                        Rst.Close
                    End If
                    
                    ReDim Parameter(0) As Parameter
                    Parameter(0) = GenerateInputParameter("@DeviceCode", adInteger, 4, .TextMatrix(.Row, .Col))
                    Set Rst = RunParametricStoredProcedure2Rec("Get_DeviceType_From_DeviceCode", Parameter)
                    .TextMatrix(.Row, 2) = Rst.Fields("DeviceTypeCode").Value
                    .TextMatrix(.Row, 6) = Rst.Fields("BufferSize").Value
                    .TextMatrix(.Row, 7) = Rst.Fields("RThreshold").Value
                
                Case 2
                    s = ""
                    If Rst.State <> 0 Then
                        Rst.Close
                    End If
                    
                    ReDim Parameter(0) As Parameter
                    Parameter(0) = GenerateInputParameter("@DeviceTypeCode", adInteger, 4, .TextMatrix(.Row, .Col))
                    Set Rst = RunParametricStoredProcedure2Rec("Get_DeviceCode_From_DeviceType", Parameter)
                    .TextMatrix(.Row, 3) = Rst.Fields("DeviceCode").Value
                    .TextMatrix(.Row, 6) = Rst.Fields("BufferSize").Value
                    .TextMatrix(.Row, 7) = Rst.Fields("RThreshold").Value
            End Select
        End If
    End With
    Set Rst = Nothing
End Sub

Private Sub VSDeviceSetting_KeyDown(KeyCode As Integer, Shift As Integer)
    
    Dim j As Integer
    With VSDeviceSetting
        If .Rows > 1 And .Row <> 0 Then
            Select Case MyFormAddEditMode
                Case AddMode
                    If .Row = .Rows - 1 Then
                        .Select .Row, .Col
                        .EditCell
    
                    End If
                
                Case EditMode
                    If InStr(1, .TextMatrix(.Row, 0), "*") = 0 Then
                        For i = 1 To .Rows - 1
                            If InStr(1, .TextMatrix(i, 0), "*") <> 0 Then
                                j = i
                                Exit For
                            End If
                        Next
                        If j = 0 Then
                            .TextMatrix(.Row, 0) = .TextMatrix(.Row, 0) & "*"
                            .Select .Row, .Col
                            .EditCell
                        
                        End If
                    Else
                        .Select .Row, .Col
                        .EditCell
                    End If
                    
            End Select
        End If
    End With

End Sub

Private Sub VSDeviceSetting_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Dim j As Integer
    With VSDeviceSetting
        If .Rows > 1 And .Row <> 0 Then
            Select Case MyFormAddEditMode
                Case AddMode
                    If .Row = .Rows - 1 Then
                        .Select .Row, .Col
                        .EditCell
    
                    End If
                
                Case EditMode
                    If InStr(1, .TextMatrix(.Row, 0), "*") = 0 Then
                        For i = 1 To .Rows - 1
                            If InStr(1, .TextMatrix(i, 0), "*") <> 0 Then
                                j = i
                                Exit For
                            End If
                        Next
                        If j = 0 Then
                            .TextMatrix(.Row, 0) = .TextMatrix(.Row, 0) & "*"
                            .Select .Row, .Col
                            .EditCell
                        
                        End If
                    Else
                        .Select .Row, .Col
                        .EditCell
                    End If
                    
            End Select
        End If
    End With

End Sub

Private Sub VSDeviceSetting_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If MyFormAddEditMode = EditMode Then
        With VSDeviceSetting
            If InStr(1, .TextMatrix(Row, 0), "*") = 0 Then
                .TextMatrix(Row, 0) = .TextMatrix(Row, 0) & "*"
            End If
        End With
    End If
End Sub

Private Sub VSDeviceSetting_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With VSDeviceSetting
        .Row = Row
        .Col = Col
    End With
End Sub
